/** Plant / contractor hire billing helpers. */

export function monthStart(period) {
  return `${String(period).slice(0, 7)}-01`;
}

export function monthEnd(period) {
  const [y, m] = String(period).split("-").map((n) => Number(n));
  const d = new Date(Date.UTC(y, m, 0));
  return d.toISOString().slice(0, 10);
}

export function prevMonth(period) {
  const [y, m] = String(period).split("-").map((n) => Number(n));
  const d = new Date(Date.UTC(y, m - 2, 1));
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
}

export function isHiredCategory(category) {
  const c = String(category || "").toLowerCase();
  return c.includes("contractor hire") || c.includes("contractor") || c.includes("hire");
}

export function ensurePlantHireSchema(db) {
  const ensure = (col, ddl) => {
    const cols = db.prepare(`PRAGMA table_info(assets)`).all();
    if (!cols.some((c) => String(c.name) === col)) {
      db.prepare(`ALTER TABLE assets ADD COLUMN ${ddl}`).run();
    }
  };
  ensure("hire_billing_mode", "hire_billing_mode TEXT");
  ensure("hire_rate_per_hour", "hire_rate_per_hour REAL");
  ensure("hire_fixed_monthly", "hire_fixed_monthly REAL");
}

export function normalizeHireBillingMode(raw) {
  const s = String(raw || "").trim().toLowerCase();
  if (s === "hourly" || s === "fixed_monthly" || s === "fixed") {
    return s === "fixed" ? "fixed_monthly" : s;
  }
  return "";
}

/**
 * Per hired asset cost lines for a calendar month.
 * @returns {Array<{asset_code, asset_name, category, site_code, cost_center_code, billing_mode, hours_run, rate, amount}>}
 */
export function buildPlantHireLines(db, period) {
  ensurePlantHireSchema(db);
  const start = monthStart(period);
  const end = monthEnd(period);

  const assets = db.prepare(`
    SELECT
      id, asset_code, asset_name, category, site_code, cost_center_code,
      hire_billing_mode, hire_rate_per_hour, hire_fixed_monthly, active, archived
    FROM assets
    WHERE COALESCE(archived, 0) = 0
      AND (
        LOWER(COALESCE(category, '')) LIKE '%contractor%'
        OR LOWER(COALESCE(category, '')) LIKE '%hire%'
      )
  `).all();

  const hoursStmt = db.prepare(`
    SELECT COALESCE(SUM(COALESCE(hours_run, 0)), 0) AS h
    FROM daily_hours
    WHERE asset_id = ?
      AND work_date BETWEEN DATE(?) AND DATE(?)
      AND COALESCE(is_used, 1) = 1
  `);

  const lines = [];
  for (const a of assets) {
    const mode = normalizeHireBillingMode(a.hire_billing_mode);
    if (!mode) continue;

    let hours_run = 0;
    let rate = 0;
    let amount = 0;

    if (mode === "hourly") {
      hours_run = Number(hoursStmt.get(a.id, start, end)?.h || 0);
      rate = Number(a.hire_rate_per_hour || 0);
      amount = hours_run * rate;
    } else if (mode === "fixed_monthly") {
      rate = Number(a.hire_fixed_monthly || 0);
      amount = rate;
    }

    if (amount <= 0) continue;

    lines.push({
      asset_code: a.asset_code,
      asset_name: a.asset_name,
      category: a.category || "Plant Hire",
      site_code: a.site_code || null,
      cost_center_code: a.cost_center_code || null,
      billing_mode: mode,
      hours_run: Number(hours_run.toFixed(2)),
      rate: Number(rate.toFixed(4)),
      amount: Number(amount.toFixed(2)),
    });
  }

  return lines.sort((a, b) => String(a.asset_code).localeCompare(String(b.asset_code)));
}

/** Finance actuals rows (category plant_hire) aggregated by site / cost center / equipment type. */
export function buildPlantHireFinanceRows(db, period) {
  const lines = buildPlantHireLines(db, period);
  const out = new Map();
  for (const line of lines) {
    const key = `${line.site_code || ""}|${line.cost_center_code || ""}|${line.category || ""}`;
    const prev = out.get(key) || {
      period,
      site_code: line.site_code || null,
      cost_center_code: line.cost_center_code || null,
      equipment_type: line.category || null,
      category: "plant_hire",
      actual_amount: 0,
    };
    prev.actual_amount += Number(line.amount || 0);
    out.set(key, prev);
  }
  return Array.from(out.values()).map((r) => ({
    ...r,
    actual_amount: Number(r.actual_amount.toFixed(2)),
  }));
}

export function listHireAssetsRegister(db) {
  ensurePlantHireSchema(db);
  return db.prepare(`
    SELECT
      asset_code, asset_name, category, site_code, cost_center_code,
      hire_billing_mode, hire_rate_per_hour, hire_fixed_monthly,
      active, archived
    FROM assets
    WHERE COALESCE(archived, 0) = 0
      AND (
        LOWER(COALESCE(category, '')) LIKE '%contractor%'
        OR LOWER(COALESCE(category, '')) LIKE '%hire%'
      )
    ORDER BY asset_code ASC
  `).all();
}
