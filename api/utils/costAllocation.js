/** Schema + helpers for site / cost center on assets, fuel, and lube logs. */

export function ensureCostAllocationSchema(db) {
  if (!db) return;
  const hasTable = (name) => {
    const r = db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name=? LIMIT 1`).get(String(name));
    return Boolean(r?.name);
  };
  const hasColumn = (table, col) => {
    if (!hasTable(table)) return false;
    return db.prepare(`PRAGMA table_info(${table})`).all().some((c) => String(c.name) === col);
  };
  const ensureColumn = (table, colName, colDef) => {
    if (!hasTable(table) || hasColumn(table, colName)) return;
    try {
      db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
    } catch {}
  };
  ensureColumn("assets", "site_code", "site_code TEXT");
  ensureColumn("assets", "cost_center_code", "cost_center_code TEXT");
  ensureColumn("assets", "department_code", "department_code TEXT");
  ensureColumn("fuel_logs", "cost_center_code", "cost_center_code TEXT");
  ensureColumn("oil_logs", "cost_center_code", "cost_center_code TEXT");
}

export function normalizeCostCenterCode(raw) {
  const s = String(raw || "").trim();
  return s ? s.toUpperCase() : null;
}

export function normalizeSiteCode(raw) {
  const s = String(raw || "").trim();
  return s ? s.toLowerCase() : null;
}

/** Cost center for a log: explicit body value, else asset default. */
export function resolveLogCostCenterCode(db, assetId, bodyOverride) {
  const fromBody = normalizeCostCenterCode(bodyOverride);
  if (fromBody) return fromBody;
  if (!assetId) return null;
  try {
    const row = db.prepare(`SELECT cost_center_code FROM assets WHERE id = ? LIMIT 1`).get(Number(assetId));
    return normalizeCostCenterCode(row?.cost_center_code);
  } catch {
    return null;
  }
}
