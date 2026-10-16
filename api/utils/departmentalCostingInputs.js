// Monthly departmental entries that are not represented by operational logs.
// Keep these separate from source transactions so the costing export can show
// precisely what was entered manually and never manufacture a financial value.

export const PLANT_DEPARTMENT = "AML-3001";
export const WORKSHOP_DEPARTMENT = "AML-3002";

export const DEPARTMENTAL_COST_INPUT_TYPES = Object.freeze([
  {
    key: "plant_revenue",
    label: "External plant hire revenue",
    department: PLANT_DEPARTMENT,
  },
  {
    key: "plant_other_operating",
    label: "Plant other operating expense",
    department: PLANT_DEPARTMENT,
  },
  {
    key: "plant_depreciation",
    label: "Plant depreciation",
    department: PLANT_DEPARTMENT,
  },
  {
    key: "workshop_duties",
    label: "Workshop import / export duties",
    department: WORKSHOP_DEPARTMENT,
  },
  {
    key: "workshop_other_operating",
    label: "Workshop other operating expense",
    department: WORKSHOP_DEPARTMENT,
  },
]);

const INPUT_TYPE_BY_KEY = new Map(DEPARTMENTAL_COST_INPUT_TYPES.map((item) => [item.key, item]));

function normalizedSiteCode(value) {
  return String(value || "main").trim().toLowerCase() || "main";
}

export function ensureDepartmentalCostingInputSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS departmental_costing_inputs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      period TEXT NOT NULL,
      site_code TEXT NOT NULL DEFAULT 'main',
      entry_type TEXT NOT NULL,
      asset_code TEXT,
      description TEXT,
      reference TEXT,
      amount REAL NOT NULL DEFAULT 0,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `);
  db.exec(`
    CREATE INDEX IF NOT EXISTS idx_departmental_costing_inputs_period_site
    ON departmental_costing_inputs (period, site_code, entry_type)
  `);
}

export function getDepartmentalCostInputType(value) {
  return INPUT_TYPE_BY_KEY.get(String(value || "").trim().toLowerCase()) || null;
}

export function listDepartmentalCostingInputs(db, { period, siteCode = "main" } = {}) {
  ensureDepartmentalCostingInputSchema(db);
  return db.prepare(`
    SELECT id, period, site_code, entry_type, asset_code, description, reference, amount, created_by, created_at
    FROM departmental_costing_inputs
    WHERE period = ? AND LOWER(COALESCE(NULLIF(TRIM(site_code), ''), 'main')) = ?
    ORDER BY created_at ASC, id ASC
  `).all(String(period || "").trim(), normalizedSiteCode(siteCode)).map((row) => ({
    ...row,
    amount: Number(row.amount || 0),
    input_type: getDepartmentalCostInputType(row.entry_type),
  }));
}

export function createDepartmentalCostingInput(db, {
  period,
  siteCode = "main",
  entryType,
  assetCode = null,
  description = null,
  reference = null,
  amount,
  createdBy = null,
} = {}) {
  ensureDepartmentalCostingInputSchema(db);
  const cleanPeriod = String(period || "").trim();
  const type = getDepartmentalCostInputType(entryType);
  const numericAmount = Number(amount);
  if (!/^\d{4}-\d{2}$/.test(cleanPeriod)) throw new Error("period (YYYY-MM) required");
  if (!type) throw new Error("valid entry_type required");
  if (!Number.isFinite(numericAmount) || numericAmount === 0) throw new Error("amount must be a non-zero number");

  const info = db.prepare(`
    INSERT INTO departmental_costing_inputs
      (period, site_code, entry_type, asset_code, description, reference, amount, created_by)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    cleanPeriod,
    normalizedSiteCode(siteCode),
    type.key,
    String(assetCode || "").trim().toUpperCase() || null,
    String(description || "").trim() || null,
    String(reference || "").trim() || null,
    numericAmount,
    String(createdBy || "").trim() || null,
  );

  return db.prepare(`
    SELECT id, period, site_code, entry_type, asset_code, description, reference, amount, created_by, created_at
    FROM departmental_costing_inputs WHERE id = ?
  `).get(Number(info.lastInsertRowid));
}

export function deleteDepartmentalCostingInput(db, { id, siteCode = "main" } = {}) {
  ensureDepartmentalCostingInputSchema(db);
  const out = db.prepare(`
    DELETE FROM departmental_costing_inputs
    WHERE id = ? AND LOWER(COALESCE(NULLIF(TRIM(site_code), ''), 'main')) = ?
  `).run(Number(id), normalizedSiteCode(siteCode));
  return Number(out.changes || 0);
}
