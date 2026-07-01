// Line-level lube / oil usage (oil logs + stores issues to equipment / work orders).

function hasColumn(db, table, col) {
  try {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name) === col);
  } catch {
    return false;
  }
}

function hasTable(db, table) {
  try {
    const r = db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name=? LIMIT 1`).get(String(table));
    return Boolean(r?.name);
  } catch {
    return false;
  }
}

function lubePartNameCodeSql(alias = "p") {
  const p = alias;
  return `
    LOWER(COALESCE(${p}.part_name, '')) LIKE '%oil%'
    OR LOWER(COALESCE(${p}.part_name, '')) LIKE '%lube%'
    OR LOWER(COALESCE(${p}.part_name, '')) LIKE '%grease%'
    OR LOWER(COALESCE(${p}.part_name, '')) LIKE '%hydraulic%'
    OR LOWER(COALESCE(${p}.part_code, '')) LIKE '%oil%'
    OR LOWER(COALESCE(${p}.part_code, '')) LIKE '%lube%'
  `;
}

/** SQL fragment: part row is oil / lube / grease (alias e.g. p). */
export function lubePartWhereSql(db, alias = "p") {
  const p = alias;
  const nameCode = lubePartNameCodeSql(p);
  if (hasColumn(db, "parts", "consumable_kind")) {
    return `(
      LOWER(COALESCE(${p}.consumable_kind, '')) IN ('oil', 'lube', 'lubricant', 'hydraulic', 'hydraulic_oil', 'coolant', 'grease')
      OR ${nameCode}
    )`;
  }
  return `(${nameCode})`;
}

const ROLE_OIL_TYPES = new Set(["admin", "supervisor", "manager", "stores", "artisan", "operator"]);

function normalizeOilTypeKey(raw) {
  const t = String(raw || "").trim();
  if (!t) return "UNSPECIFIED";
  if (ROLE_OIL_TYPES.has(t.toLowerCase())) return "UNSPECIFIED";
  return t;
}

function num(v, digits = 2) {
  const n = Number(v);
  if (!Number.isFinite(n)) return 0;
  return Number(n.toFixed(digits));
}

function resolveSmrForAssetDate(db, assetId, usageDate, cache) {
  const aid = Number(assetId);
  const d = String(usageDate || "").trim();
  if (!Number.isFinite(aid) || aid <= 0 || !d) return null;
  const key = `${aid}|${d}`;
  if (cache.has(key)) return cache.get(key);

  let smr = null;
  if (hasTable(db, "daily_hours") && hasColumn(db, "daily_hours", "closing_hours")) {
    const exact = db.prepare(`
      SELECT closing_hours
      FROM daily_hours
      WHERE asset_id = ?
        AND work_date = ?
        AND closing_hours IS NOT NULL
      LIMIT 1
    `).get(aid, d);
    if (exact?.closing_hours != null) {
      smr = num(exact.closing_hours, 1);
    } else {
      const near = db.prepare(`
        SELECT closing_hours
        FROM daily_hours
        WHERE asset_id = ?
          AND work_date <= ?
          AND closing_hours IS NOT NULL
        ORDER BY work_date DESC
        LIMIT 1
      `).get(aid, d);
      if (near?.closing_hours != null) smr = num(near.closing_hours, 1);
    }
  }
  if (smr == null && hasTable(db, "asset_hours")) {
    const ah = db.prepare(`SELECT total_hours FROM asset_hours WHERE asset_id = ?`).get(aid);
    if (ah?.total_hours != null) smr = num(ah.total_hours, 1);
  }
  cache.set(key, smr);
  return smr;
}

function enrichLinesWithSmr(db, lines) {
  const cache = new Map();
  return lines.map((line) => ({
    ...line,
    smr: resolveSmrForAssetDate(db, line.asset_id, line.usage_date, cache),
  }));
}

function stockMovementDateExpr(db) {
  const hasCreated = hasColumn(db, "stock_movements", "created_at");
  const hasMvDate = hasColumn(db, "stock_movements", "movement_date");
  if (hasCreated && hasMvDate) return "date(COALESCE(sm.created_at, sm.movement_date))";
  if (hasCreated) return "date(sm.created_at)";
  if (hasMvDate) return "date(sm.movement_date)";
  return "date('1970-01-01')";
}

/**
 * @param {import('better-sqlite3').Database} db
 * @param {{ start: string, end: string, lubeUnitFallback?: number }} opts
 */
export function fetchLubeUsageLines(db, { start, end, lubeUnitFallback = 4 }) {
  const fallback = Number.isFinite(Number(lubeUnitFallback)) && Number(lubeUnitFallback) > 0
    ? Number(lubeUnitFallback)
    : 4;
  const lines = [];
  const hasConsumableKind = hasColumn(db, "parts", "consumable_kind");
  const hasLubeMappings = hasTable(db, "lube_type_mappings");
  const lubeTypeExpr = hasConsumableKind && hasLubeMappings
    ? `COALESCE(NULLIF(TRIM(p.consumable_kind), ''), NULLIF(TRIM(ltm.part_code), ''), 'lube')`
    : hasConsumableKind
      ? `COALESCE(NULLIF(TRIM(p.consumable_kind), ''), 'lube')`
      : hasLubeMappings
        ? `COALESCE(NULLIF(TRIM(ltm.part_code), ''), 'lube')`
        : `'lube'`;
  const ltmJoin = hasLubeMappings
    ? `LEFT JOIN lube_type_mappings ltm ON LOWER(TRIM(ltm.oil_key)) = LOWER(TRIM(COALESCE(ol.oil_type, '')))`
    : "";

  if (hasTable(db, "oil_logs") && hasTable(db, "assets")) {
    const oilRows = db.prepare(`
      SELECT
        ol.id,
        ol.log_date AS usage_date,
        ol.asset_id,
        a.asset_code,
        a.asset_name,
        COALESCE(
          NULLIF(TRIM(p.part_code), ''),
          CASE
            WHEN LOWER(TRIM(COALESCE(ol.oil_type, ''))) IN ('admin','supervisor','manager','stores','artisan','operator') THEN NULL
            ELSE NULLIF(TRIM(ol.oil_type), '')
          END,
          'UNSPECIFIED'
        ) AS part_code,
        COALESCE(NULLIF(TRIM(p.part_name), ''), NULLIF(TRIM(ol.oil_type), ''), '') AS part_name,
        ${lubeTypeExpr} AS lube_type,
        ol.quantity,
        COALESCE(ol.unit_cost, p.unit_cost, ?) AS unit_cost,
        'oil_log' AS source,
        NULL AS work_order_id
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      LEFT JOIN parts p ON UPPER(TRIM(p.part_code)) = UPPER(TRIM(COALESCE(ol.oil_type, '')))
      ${ltmJoin}
      WHERE ol.log_date BETWEEN ? AND ?
      ORDER BY ol.log_date ASC, a.asset_code ASC, ol.id ASC
      LIMIT 10000
    `).all(fallback, start, end);

    for (const r of oilRows) {
      const qty = num(r.quantity, 3);
      const unit = num(r.unit_cost, 4);
      lines.push({
        id: Number(r.id || 0),
        usage_date: String(r.usage_date || ""),
        asset_id: Number(r.asset_id || 0) || null,
        asset_code: String(r.asset_code || ""),
        asset_name: String(r.asset_name || ""),
        part_code: String(r.part_code || "UNSPECIFIED"),
        part_name: String(r.part_name || ""),
        lube_type: normalizeOilTypeKey(r.lube_type) === "UNSPECIFIED" ? "lube" : String(r.lube_type || "lube"),
        quantity: qty,
        unit_cost: unit,
        line_cost: num(qty * unit, 2),
        source: "oil_log",
        work_order_id: null,
      });
    }
  }

  if (hasTable(db, "stock_movements") && hasTable(db, "parts") && hasTable(db, "assets")) {
    const dateExpr = stockMovementDateExpr(db);
    const hasMvUnit = hasColumn(db, "stock_movements", "unit_cost");
    const unitExpr = hasMvUnit
      ? `COALESCE(sm.unit_cost, p.unit_cost, ${fallback})`
      : `COALESCE(p.unit_cost, ${fallback})`;
    const hasWorkOrders = hasTable(db, "work_orders");
    const lubeTypeCol = hasConsumableKind
      ? `COALESCE(NULLIF(TRIM(p.consumable_kind), ''), 'lube')`
      : `'lube'`;

    const woJoin = hasWorkOrders
      ? `LEFT JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
         LEFT JOIN assets aw ON aw.id = w.asset_id`
      : `LEFT JOIN assets aw ON 0`;
    const woIdExpr = hasWorkOrders
      ? `CASE WHEN sm.reference LIKE 'work_order:%' THEN CAST(substr(sm.reference, 12) AS INTEGER) ELSE NULL END`
      : `NULL`;

    const stockRows = db.prepare(`
      SELECT
        sm.id,
        ${dateExpr} AS usage_date,
        COALESCE(aw.id, aa.id) AS asset_id,
        COALESCE(aw.asset_code, aa.asset_code) AS asset_code,
        COALESCE(aw.asset_name, aa.asset_name) AS asset_name,
        p.part_code,
        p.part_name,
        ${lubeTypeCol} AS lube_type,
        ABS(sm.quantity) AS quantity,
        ${unitExpr} AS unit_cost,
        CASE WHEN sm.reference LIKE 'work_order:%' THEN 'work_order' ELSE 'stores_issue' END AS source,
        ${woIdExpr} AS work_order_id
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      ${woJoin}
      LEFT JOIN assets aa ON sm.reference LIKE ('asset:' || aa.id || ':stores')
      WHERE sm.movement_type = 'out'
        AND ${lubePartWhereSql(db, "p")}
        AND ${dateExpr} BETWEEN ? AND ?
        AND sm.reference NOT LIKE 'lube_issue:%'
        AND COALESCE(aw.id, aa.id) IS NOT NULL
      ORDER BY ${dateExpr} ASC, asset_code ASC, sm.id ASC
      LIMIT 10000
    `).all(start, end);

    for (const r of stockRows) {
      const qty = num(r.quantity, 3);
      const unit = num(r.unit_cost, 4);
      lines.push({
        id: Number(r.id || 0),
        usage_date: String(r.usage_date || ""),
        asset_id: Number(r.asset_id || 0) || null,
        asset_code: String(r.asset_code || ""),
        asset_name: String(r.asset_name || ""),
        part_code: String(r.part_code || ""),
        part_name: String(r.part_name || ""),
        lube_type: String(r.lube_type || "lube"),
        quantity: qty,
        unit_cost: unit,
        line_cost: num(qty * unit, 2),
        source: String(r.source || "stores_issue"),
        work_order_id: r.work_order_id != null ? Number(r.work_order_id) : null,
      });
    }
  }

  lines.sort((a, b) => {
    const d = String(a.usage_date).localeCompare(String(b.usage_date));
    if (d !== 0) return d;
    const ac = String(a.asset_code).localeCompare(String(b.asset_code));
    if (ac !== 0) return ac;
    return String(a.part_code).localeCompare(String(b.part_code));
  });

  return enrichLinesWithSmr(db, lines);
}
