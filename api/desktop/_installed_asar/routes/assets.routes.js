// IRONLOG/api/routes/assets.routes.js
import { db } from "../db/client.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

export default async function assetRoutes(app) {
  /* =========================
     PREPARED STATEMENTS
  ========================= */

  const getAssetByCode = db.prepare(`
    SELECT
      id, asset_code, asset_name, category,
      active, is_standby,
      archived, archive_reason, archived_at,
      created_at
    FROM assets
    WHERE asset_code = ?
  `);

  const insertAsset = db.prepare(`
    INSERT INTO assets (
      asset_code, asset_name, category,
      active, is_standby,
      archived, archive_reason, archived_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const updateActive = db.prepare(`UPDATE assets SET active = ? WHERE asset_code = ?`);
  const updateStandby = db.prepare(`UPDATE assets SET is_standby = ? WHERE asset_code = ?`);

  /* =========================
     ROUTES
  ========================= */

  // GET /api/assets?include_archived=1
  app.get("/", async (req) => {
    const includeArchived = String(req.query?.include_archived || "0") === "1";

    const rows = db.prepare(`
      SELECT
        id, asset_code, asset_name, category,
        active, is_standby,
        archived, archive_reason, archived_at,
        created_at
      FROM assets
      WHERE (? = 1 OR archived = 0)
      ORDER BY asset_code ASC
    `).all(includeArchived ? 1 : 0);

    return rows.map((r) => ({
      ...r,
      active: Number(r.active),
      is_standby: Number(r.is_standby),
      archived: Number(r.archived),
    }));
  });

  // POST /api/assets
  app.post("/", async (req, reply) => {
    const body = req.body || {};
    const asset_code = String(body.asset_code || "").trim();
    const asset_name = String(body.asset_name || "").trim();
    const category = String(body.category || "").trim() || null;

    const active = body.active === 0 || body.active === false ? 0 : 1;
    const is_standby = body.is_standby ? 1 : 0;

    // archived fields (optional)
    const archived = body.archived ? 1 : 0;
    const archive_reason = archived ? (String(body.archive_reason || "").trim() || null) : null;
    const archived_at = archived ? (String(body.archived_at || "").trim() || null) : null;

    if (!asset_code || !asset_name) {
      return reply.code(400).send({ error: "asset_code and asset_name are required" });
    }

    try {
      const r = insertAsset.run(
        asset_code,
        asset_name,
        category,
        active,
        is_standby,
        archived,
        archive_reason,
        archived_at
      );
      return reply.code(201).send({ ok: true, id: Number(r.lastInsertRowid) });
    } catch (e) {
      return reply.code(400).send({ error: e.message || String(e) });
    }
  });

  // POST /api/assets/:asset_code/archive  { archived: true/false, reason? }
  app.post("/:asset_code/archive", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    const body = req.body || {};

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const archived = body.archived ? 1 : 0;
    const reason = archived ? (String(body.reason || "").trim() || null) : null;

    if (archived) {
      db.prepare(`
        UPDATE assets
        SET archived = 1, archive_reason = ?, archived_at = datetime('now')
        WHERE asset_code = ?
      `).run(reason, asset_code);
    } else {
      db.prepare(`
        UPDATE assets
        SET archived = 0, archive_reason = NULL, archived_at = NULL
        WHERE asset_code = ?
      `).run(asset_code);
    }

    return { ok: true };
  });

  // PATCH /api/assets/:asset_code  { active?, is_standby? }
  app.patch("/:asset_code", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    const body = req.body || {};

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const tx = db.transaction(() => {
      if (body.active !== undefined) updateActive.run(body.active ? 1 : 0, asset_code);
      if (body.is_standby !== undefined) updateStandby.run(body.is_standby ? 1 : 0, asset_code);
    });

    tx();

    const updated = getAssetByCode.get(asset_code);
    return {
      ok: true,
      asset: {
        ...updated,
        active: Number(updated.active),
        is_standby: Number(updated.is_standby),
        archived: Number(updated.archived),
      },
    };
  });

  // ============================================================
  // ASSET HISTORY (Timeline)
  // GET /api/assets/:asset_code/history?start=YYYY-MM-DD&end=YYYY-MM-DD
  // ============================================================
  app.get("/:asset_code/history", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const startOk = start && isDate(start);
    const endOk = end && isDate(end);

    function dateFilter(col) {
      const clauses = [];
      const params = [];
      if (startOk) { clauses.push(`${col} >= ?`); params.push(start); }
      if (endOk) { clauses.push(`${col} <= ?`); params.push(end); }
      return { sql: clauses.length ? " AND " + clauses.join(" AND ") : "", params };
    }

    function hasTable(tableName) {
      const row = db.prepare(`
        SELECT name
        FROM sqlite_master
        WHERE type = 'table' AND name = ?
      `).get(tableName);
      return Boolean(row);
    }

    function hasColumn(tableName, columnName) {
      if (!hasTable(tableName)) return false;
      const cols = db.prepare(`PRAGMA table_info(${tableName})`).all();
      return cols.some((c) => String(c.name) === columnName);
    }

    function firstExistingTable(candidates) {
      for (const name of candidates) {
        if (hasTable(name)) return name;
      }
      return null;
    }

    function firstExistingColumn(tableName, candidates) {
      for (const name of candidates) {
        if (hasColumn(tableName, name)) return name;
      }
      return null;
    }

    // ---- BREAKDOWNS (include downtime total + downtime log lines)
    const bdF = dateFilter("b.breakdown_date");
    let breakdowns = [];
    let downtimeLogsByBreakdown = new Map();

    try {
      breakdowns = db.prepare(`
        SELECT
          b.id,
          b.breakdown_date AS date,
          b.description AS title,
          b.critical,
          COALESCE(SUM(l.hours_down), 0) AS downtime_hours
        FROM breakdowns b
        LEFT JOIN breakdown_downtime_logs l ON l.breakdown_id = b.id
        WHERE b.asset_id = ? ${bdF.sql}
        GROUP BY b.id
        ORDER BY b.breakdown_date DESC, b.id DESC
        LIMIT 200
      `).all(asset.id, ...bdF.params);

      const ids = breakdowns.map((b) => b.id);
      if (ids.length) {
        const placeholders = ids.map(() => "?").join(",");
        const logs = db.prepare(`
          SELECT breakdown_id, log_date, hours_down, notes, created_at
          FROM breakdown_downtime_logs
          WHERE breakdown_id IN (${placeholders})
          ORDER BY log_date DESC, id DESC
        `).all(...ids);

        downtimeLogsByBreakdown = new Map();
        for (const l of logs) {
          const arr = downtimeLogsByBreakdown.get(l.breakdown_id) || [];
          arr.push({
            log_date: l.log_date,
            hours_down: Number(l.hours_down || 0),
            notes: l.notes || "",
            created_at: l.created_at,
          });
          downtimeLogsByBreakdown.set(l.breakdown_id, arr);
        }
      }
    } catch {
      breakdowns = db.prepare(`
        SELECT
          b.id,
          b.breakdown_date AS date,
          b.description AS title,
          b.critical,
          b.downtime_hours AS downtime_hours
        FROM breakdowns b
        WHERE b.asset_id = ? ${bdF.sql}
        ORDER BY b.breakdown_date DESC, b.id DESC
        LIMIT 200
      `).all(asset.id, ...bdF.params);

      downtimeLogsByBreakdown = new Map();
    }

    const woForBreakdown = db.prepare(`
      SELECT id
      FROM work_orders
      WHERE source = 'breakdown' AND reference_id = ?
      ORDER BY id DESC
      LIMIT 1
    `);

    const breakdownEvents = breakdowns.map((b) => ({
      type: "breakdown",
      date: b.date,
      title: b.title,
      work_order_id: woForBreakdown.get(b.id)?.id ?? null,
      details: {
        breakdown_id: b.id,
        critical: Boolean(b.critical),
        downtime_hours: Number(b.downtime_hours || 0),
        downtime_logs: downtimeLogsByBreakdown.get(b.id) || [],
        photos: [],
        photo: null,
      },
    }));

    if (hasTable("breakdown_photos") && breakdownEvents.length) {
      const ids = breakdownEvents.map((e) => Number(e.details.breakdown_id)).filter((n) => Number.isFinite(n));
      if (ids.length) {
        const placeholders = ids.map(() => "?").join(",");
        const rows = db.prepare(`
          SELECT breakdown_id, file_path, photo_stage, created_at
          FROM breakdown_photos
          WHERE breakdown_id IN (${placeholders})
          ORDER BY id DESC
        `).all(...ids);

        const byBreakdown = new Map();
        for (const r of rows) {
          const arr = byBreakdown.get(r.breakdown_id) || [];
          arr.push({
            file_path: r.file_path,
            photo_stage: r.photo_stage || null,
            created_at: r.created_at || null,
          });
          byBreakdown.set(r.breakdown_id, arr);
        }

        for (const ev of breakdownEvents) {
          const photos = byBreakdown.get(Number(ev.details.breakdown_id)) || [];
          ev.details.photos = photos;
          ev.details.photo = photos[0]?.file_path || null;
        }
      }
    }

    // ---- WORK ORDERS
    const woF = dateFilter("DATE(w.opened_at)");
    const workOrders = db.prepare(`
      SELECT
        w.id,
        DATE(w.opened_at) AS date,
        w.source,
        w.status,
        w.opened_at,
        w.closed_at
      FROM work_orders w
      WHERE w.asset_id = ? ${woF.sql}
      ORDER BY w.id DESC
      LIMIT 200
    `).all(asset.id, ...woF.params).map((w) => ({
      type: "work_order",
      date: w.date,
      title: `WO #${w.id} (${w.source})`,
      work_order_id: w.id,
      details: {
        source: w.source,
        status: w.status,
        opened_at: w.opened_at,
        closed_at: w.closed_at,
      },
    }));

    if (hasTable("work_order_photos") && workOrders.length) {
      const ids = workOrders.map((w) => Number(w.work_order_id)).filter((n) => Number.isFinite(n));
      if (ids.length) {
        const placeholders = ids.map(() => "?").join(",");
        const rows = db.prepare(`
          SELECT work_order_id, file_path, photo_stage, created_at
          FROM work_order_photos
          WHERE work_order_id IN (${placeholders})
          ORDER BY id DESC
        `).all(...ids);

        const byWo = new Map();
        for (const r of rows) {
          const arr = byWo.get(r.work_order_id) || [];
          arr.push({
            file_path: r.file_path,
            photo_stage: r.photo_stage || null,
            created_at: r.created_at || null,
          });
          byWo.set(r.work_order_id, arr);
        }

        for (const ev of workOrders) {
          const photos = byWo.get(Number(ev.work_order_id)) || [];
          ev.details.photos = photos;
          ev.details.photo = photos[0]?.file_path || null;
        }
      }
    }

    // ---- GET CHANGE SLIPS (optional table)
    let getSlipEvents = [];
    if (hasTable("get_change_slips")) {
      const gsF = dateFilter("g.slip_date");
      getSlipEvents = db.prepare(`
        SELECT g.id, g.slip_date AS date, g.location, g.notes
        FROM get_change_slips g
        WHERE g.asset_id = ? ${gsF.sql}
        ORDER BY g.slip_date DESC, g.id DESC
        LIMIT 200
      `).all(asset.id, ...gsF.params).map((g) => {
        const items = hasTable("get_change_items")
          ? db.prepare(`
              SELECT position, part_code, part_name, qty, reason
              FROM get_change_items
              WHERE slip_id = ?
              ORDER BY id ASC
            `).all(g.id)
          : [];

        const wo = hasTable("work_order_links")
          ? db.prepare(`
              SELECT work_order_id
              FROM work_order_links
              WHERE link_type = 'get_slip' AND link_id = ?
              ORDER BY id DESC
              LIMIT 1
            `).get(g.id)
          : null;

        return {
          type: "get_slip",
          date: g.date,
          title: `GET Change Slip #${g.id}`,
          work_order_id: wo?.work_order_id ?? null,
          details: {
            slip_id: g.id,
            location: g.location,
            notes: g.notes,
            items,
          },
        };
      });
    }

    // ---- COMPONENT CHANGE SLIPS (optional table)
    let componentSlipEvents = [];
    if (hasTable("component_change_slips")) {
      const csF = dateFilter("c.slip_date");
      componentSlipEvents = db.prepare(`
        SELECT
          c.id,
          c.slip_date AS date,
          c.component,
          c.serial_out,
          c.serial_in,
          c.hours_at_change,
          c.notes
        FROM component_change_slips c
        WHERE c.asset_id = ? ${csF.sql}
        ORDER BY c.slip_date DESC, c.id DESC
        LIMIT 200
      `).all(asset.id, ...csF.params).map((c) => {
        const wo = hasTable("work_order_links")
          ? db.prepare(`
              SELECT work_order_id
              FROM work_order_links
              WHERE link_type = 'component_slip' AND link_id = ?
              ORDER BY id DESC
              LIMIT 1
            `).get(c.id)
          : null;

        return {
          type: "component_slip",
          date: c.date,
          title: `Component Change: ${c.component} (#${c.id})`,
          work_order_id: wo?.work_order_id ?? null,
          details: {
            slip_id: c.id,
            component: c.component,
            serial_out: c.serial_out,
            serial_in: c.serial_in,
            hours_at_change: c.hours_at_change,
            notes: c.notes,
          },
        };
      });
    }

    // ---- DAMAGE REPORTS (schema-flexible / optional tables)
    let damageEvents = [];
    const damageTable = firstExistingTable([
      "equipment_damage_reports",
      "damage_reports",
      "asset_damage_reports",
    ]);

    if (damageTable) {
      const dDateCol = firstExistingColumn(damageTable, ["report_date", "damage_date", "date", "created_at"]);
      const dTitleCol = firstExistingColumn(damageTable, ["title", "damage_type", "component", "area", "category"]);
      const dDescCol = firstExistingColumn(damageTable, ["description", "notes", "details", "report"]);
      const dPhotoCol = firstExistingColumn(damageTable, ["photo_path", "photo_url", "image_path", "image_url", "image"]);

      const dDateExpr = dDateCol ? `DATE(d.${dDateCol})` : "DATE('now')";
      const dF = dateFilter(dDateExpr);

      const dTitleExpr = dTitleCol ? `COALESCE(d.${dTitleCol}, 'Damage Report')` : "'Damage Report'";
      const dDescExpr = dDescCol ? `d.${dDescCol}` : "NULL";
      const dPhotoExpr = dPhotoCol ? `d.${dPhotoCol}` : "NULL";

      damageEvents = db.prepare(`
        SELECT
          d.id,
          ${dDateExpr} AS date,
          ${dTitleExpr} AS title,
          ${dDescExpr} AS notes,
          ${dPhotoExpr} AS photo
        FROM ${damageTable} d
        WHERE d.asset_id = ? ${dF.sql}
        ORDER BY date DESC, d.id DESC
        LIMIT 200
      `).all(asset.id, ...dF.params).map((d) => ({
        type: "damage_report",
        date: d.date,
        title: `Damage: ${d.title || `Report #${d.id}`}`,
        work_order_id: null,
        details: {
          damage_id: d.id,
          notes: d.notes || null,
          photo: d.photo || null,
        },
      }));
    }

    // ---- TYRE EVENTS (change slips + inspections, schema-flexible)
    let tyreChangeEvents = [];
    let tyreInspectionEvents = [];

    const tyreChangeTable = firstExistingTable(["tyre_change_slips", "tire_change_slips"]);
    if (tyreChangeTable) {
      const tcDateCol = firstExistingColumn(tyreChangeTable, ["slip_date", "change_date", "date", "created_at"]);
      const tcPositionCol = firstExistingColumn(tyreChangeTable, ["position", "wheel_position", "tyre_position"]);
      const tcOutCol = firstExistingColumn(tyreChangeTable, ["serial_out", "tyre_out", "tire_out"]);
      const tcInCol = firstExistingColumn(tyreChangeTable, ["serial_in", "tyre_in", "tire_in"]);
      const tcHoursCol = firstExistingColumn(tyreChangeTable, ["hours_at_change", "machine_hours", "hour_meter"]);
      const tcNotesCol = firstExistingColumn(tyreChangeTable, ["notes", "reason", "description"]);
      const tcPhotoCol = firstExistingColumn(tyreChangeTable, ["photo_path", "photo_url", "image_path", "image_url", "image"]);

      const tcDateExpr = tcDateCol ? `DATE(t.${tcDateCol})` : "DATE('now')";
      const tcF = dateFilter(tcDateExpr);
      const tcPositionExpr = tcPositionCol ? `t.${tcPositionCol}` : "NULL";
      const tcOutExpr = tcOutCol ? `t.${tcOutCol}` : "NULL";
      const tcInExpr = tcInCol ? `t.${tcInCol}` : "NULL";
      const tcHoursExpr = tcHoursCol ? `t.${tcHoursCol}` : "NULL";
      const tcNotesExpr = tcNotesCol ? `t.${tcNotesCol}` : "NULL";
      const tcPhotoExpr = tcPhotoCol ? `t.${tcPhotoCol}` : "NULL";

      tyreChangeEvents = db.prepare(`
        SELECT
          t.id,
          ${tcDateExpr} AS date,
          ${tcPositionExpr} AS position,
          ${tcOutExpr} AS serial_out,
          ${tcInExpr} AS serial_in,
          ${tcHoursExpr} AS hours_at_change,
          ${tcNotesExpr} AS notes,
          ${tcPhotoExpr} AS photo
        FROM ${tyreChangeTable} t
        WHERE t.asset_id = ? ${tcF.sql}
        ORDER BY date DESC, t.id DESC
        LIMIT 200
      `).all(asset.id, ...tcF.params).map((t) => ({
        type: "tyre_change",
        date: t.date,
        title: `Tyre Change #${t.id}${t.position ? ` (${t.position})` : ""}`,
        work_order_id: null,
        details: {
          slip_id: t.id,
          position: t.position || null,
          serial_out: t.serial_out || null,
          serial_in: t.serial_in || null,
          hours_at_change: t.hours_at_change != null ? Number(t.hours_at_change) : null,
          notes: t.notes || null,
          photo: t.photo || null,
        },
      }));
    }

    const tyreInspectTable = firstExistingTable(["tyre_inspections", "tire_inspections"]);
    if (tyreInspectTable) {
      const tiDateCol = firstExistingColumn(tyreInspectTable, ["inspection_date", "check_date", "date", "created_at"]);
      const tiPositionCol = firstExistingColumn(tyreInspectTable, ["position", "wheel_position", "tyre_position"]);
      const tiConditionCol = firstExistingColumn(tyreInspectTable, ["condition", "status", "result"]);
      const tiPressureCol = firstExistingColumn(tyreInspectTable, ["pressure_kpa", "pressure_psi", "pressure"]);
      const tiTreadCol = firstExistingColumn(tyreInspectTable, ["tread_depth", "tread_mm"]);
      const tiNotesCol = firstExistingColumn(tyreInspectTable, ["notes", "comment", "observation"]);
      const tiPhotoCol = firstExistingColumn(tyreInspectTable, ["photo_path", "photo_url", "image_path", "image_url", "image"]);

      const tiDateExpr = tiDateCol ? `DATE(t.${tiDateCol})` : "DATE('now')";
      const tiF = dateFilter(tiDateExpr);
      const tiPositionExpr = tiPositionCol ? `t.${tiPositionCol}` : "NULL";
      const tiConditionExpr = tiConditionCol ? `t.${tiConditionCol}` : "NULL";
      const tiPressureExpr = tiPressureCol ? `t.${tiPressureCol}` : "NULL";
      const tiTreadExpr = tiTreadCol ? `t.${tiTreadCol}` : "NULL";
      const tiNotesExpr = tiNotesCol ? `t.${tiNotesCol}` : "NULL";
      const tiPhotoExpr = tiPhotoCol ? `t.${tiPhotoCol}` : "NULL";

      tyreInspectionEvents = db.prepare(`
        SELECT
          t.id,
          ${tiDateExpr} AS date,
          ${tiPositionExpr} AS position,
          ${tiConditionExpr} AS condition,
          ${tiPressureExpr} AS pressure,
          ${tiTreadExpr} AS tread_depth,
          ${tiNotesExpr} AS notes,
          ${tiPhotoExpr} AS photo
        FROM ${tyreInspectTable} t
        WHERE t.asset_id = ? ${tiF.sql}
        ORDER BY date DESC, t.id DESC
        LIMIT 200
      `).all(asset.id, ...tiF.params).map((t) => ({
        type: "tyre_inspection",
        date: t.date,
        title: `Tyre Inspection #${t.id}${t.position ? ` (${t.position})` : ""}`,
        work_order_id: null,
        details: {
          inspection_id: t.id,
          position: t.position || null,
          condition: t.condition || null,
          pressure: t.pressure != null ? Number(t.pressure) : null,
          tread_depth: t.tread_depth != null ? Number(t.tread_depth) : null,
          notes: t.notes || null,
          photo: t.photo || null,
        },
      }));
    }

    // ---- OIL TOTALS (qty + optional cost)
    const oilF = dateFilter("o.log_date");
    const hasOilTotalCost = hasColumn("oil_logs", "total_cost");
    const hasOilUnitCost = hasColumn("oil_logs", "unit_cost");
    const oilCostExpr = hasOilTotalCost
      ? "COALESCE(SUM(o.total_cost), 0)"
      : hasOilUnitCost
      ? "COALESCE(SUM(o.quantity * o.unit_cost), 0)"
      : "0";

    const oilTotals = hasTable("oil_logs")
      ? db.prepare(`
          SELECT
            COALESCE(SUM(o.quantity), 0) AS oil_qty_total,
            ${oilCostExpr} AS oil_cost_total
          FROM oil_logs o
          WHERE o.asset_id = ? ${oilF.sql}
        `).get(asset.id, ...oilF.params)
      : { oil_qty_total: 0, oil_cost_total: 0 };

    // ---- PARTS USED + COST (via WO issue stock movements)
    const hasSmCreatedAt = hasColumn("stock_movements", "created_at");
    const hasSmMovementDate = hasColumn("stock_movements", "movement_date");
    const smDateCol = hasSmCreatedAt
      ? "DATE(sm.created_at)"
      : hasSmMovementDate
      ? "DATE(sm.movement_date)"
      : "DATE('now')";
    const smF = dateFilter(smDateCol);

    const hasSmUnitCost = hasColumn("stock_movements", "unit_cost");
    const hasSmTotalCost = hasColumn("stock_movements", "total_cost");
    const hasPartsUnitCost = hasColumn("parts", "unit_cost");

    const partCostExpr = hasSmTotalCost
      ? "COALESCE(SUM(ABS(sm.total_cost)), 0)"
      : hasSmUnitCost
      ? "COALESCE(SUM(ABS(sm.quantity) * sm.unit_cost), 0)"
      : hasPartsUnitCost
      ? "COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0)"
      : "0";

    const partTotals = hasTable("stock_movements")
      ? db.prepare(`
          SELECT
            COALESCE(SUM(ABS(sm.quantity)), 0) AS parts_qty_total,
            ${partCostExpr} AS parts_cost_total
          FROM stock_movements sm
          JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
          LEFT JOIN parts p ON p.id = sm.part_id
          WHERE w.asset_id = ?
            AND sm.movement_type = 'out'
            ${smF.sql}
        `).get(asset.id, ...smF.params)
      : { parts_qty_total: 0, parts_cost_total: 0 };

    const history = [
      ...breakdownEvents,
      ...workOrders,
      ...getSlipEvents,
      ...componentSlipEvents,
      ...damageEvents,
      ...tyreChangeEvents,
      ...tyreInspectionEvents,
    ].sort((a, b) =>
      String(b.date).localeCompare(String(a.date))
    );

    const summary = {
      period_start: startOk ? start : null,
      period_end: endOk ? end : null,
      counts: {
        breakdowns: breakdownEvents.length,
        work_orders: workOrders.length,
        get_slips: getSlipEvents.length,
        component_slips: componentSlipEvents.length,
        damage_reports: damageEvents.length,
        tyre_changes: tyreChangeEvents.length,
        tyre_inspections: tyreInspectionEvents.length,
        events_total: history.length,
      },
      totals: {
        parts_qty_total: Number(partTotals?.parts_qty_total || 0),
        parts_cost_total: Number(partTotals?.parts_cost_total || 0),
        oil_qty_total: Number(oilTotals?.oil_qty_total || 0),
        oil_cost_total: Number(oilTotals?.oil_cost_total || 0),
      },
      maintenance_cost_total:
        Number(partTotals?.parts_cost_total || 0) + Number(oilTotals?.oil_cost_total || 0),
    };

    return {
      ok: true,
      asset: {
        asset_code: asset.asset_code,
        asset_name: asset.asset_name,
        category: asset.category,
        archived: Number(asset.archived),
        archive_reason: asset.archive_reason,
        archived_at: asset.archived_at,
      },
      summary,
      history,
    };
  });
}