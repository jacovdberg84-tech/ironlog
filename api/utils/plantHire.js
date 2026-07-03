/** Plant / contractor hire billing helpers. */

import { getRunFromFuelRows } from "./fuelRunFromLogs.js";
import { sqlFuelMetricModeExpr } from "./fuelMetricMode.js";
import { isHireCategory, sqlIncludeArchivedHireAssets } from "./hiredEquipment.js";

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

function hireAssetsWhereSql() {
  return `
    COALESCE(a.active, 1) = 1
    AND (
      NULLIF(TRIM(COALESCE(a.hire_billing_mode, '')), '') IS NOT NULL
      OR LOWER(COALESCE(a.category, '')) LIKE '%contractor%'
      OR LOWER(COALESCE(a.category, '')) LIKE '%hire%'
      OR UPPER(COALESCE(a.asset_code, '')) LIKE 'BMP%'
      OR UPPER(COALESCE(a.asset_code, '')) LIKE 'PTT%'
    )
    AND ${sqlIncludeArchivedHireAssets("a")}
  `;
}

function getFuelDerivedRun(db, assetId, start, end, metricMode) {
  const logs = db.prepare(`
    SELECT
      log_date,
      COALESCE(LOWER(meter_unit), '') AS meter_unit,
      COALESCE(meter_run_value, 0) AS meter_run_value,
      COALESCE(hours_run, 0) AS hours_run,
      open_meter_value,
      close_meter_value
    FROM fuel_logs
    WHERE asset_id = ?
      AND log_date BETWEEN ? AND ?
    ORDER BY log_date ASC, id ASC
  `).all(assetId, start, end);

  const prev = db.prepare(`
    SELECT
      log_date,
      COALESCE(LOWER(meter_unit), '') AS meter_unit,
      COALESCE(meter_run_value, 0) AS meter_run_value,
      COALESCE(hours_run, 0) AS hours_run,
      open_meter_value,
      close_meter_value
    FROM fuel_logs
    WHERE asset_id = ?
      AND log_date < ?
      AND (COALESCE(meter_run_value, 0) > 0 OR COALESCE(hours_run, 0) > 0)
    ORDER BY log_date DESC, id DESC
    LIMIT 1
  `).get(assetId, start);

  const mode = String(metricMode || "hours").toLowerCase() === "km" ? "km" : "hours";
  const run = getRunFromFuelRows(logs, prev, mode);
  return {
    hours_run: Number(run.hours_run || 0),
    km_run: Number(run.km_run || 0),
    source: logs.length ? "fams_fuel" : null,
  };
}

/**
 * Per hired asset cost lines for a calendar month.
 * Uses daily_hours when present; otherwise FAMS fuel meter deltas.
 */
export function buildPlantHireLines(db, period) {
  ensurePlantHireSchema(db);
  const start = monthStart(period);
  const end = monthEnd(period);
  const metricExpr = sqlFuelMetricModeExpr("a");

  const assets = db.prepare(`
    SELECT
      a.id, a.asset_code, a.asset_name, a.category, a.site_code, a.cost_center_code,
      a.hire_billing_mode, a.hire_rate_per_hour, a.hire_fixed_monthly,
      a.active, a.archived, ${metricExpr} AS metric_mode,
      COALESCE(NULLIF(a.km_per_hour_factor, 0), 10.0) AS km_per_hour_factor
    FROM assets a
    WHERE ${hireAssetsWhereSql()}
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
    let km_run = 0;
    let run_source = "none";
    let rate = 0;
    let amount = 0;

    if (mode === "hourly") {
      hours_run = Number(hoursStmt.get(a.id, start, end)?.h || 0);
      if (hours_run > 0) {
        run_source = "daily_hours";
      } else {
        const fuelRun = getFuelDerivedRun(db, a.id, start, end, a.metric_mode);
        const metric = String(a.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
        if (metric === "km" && fuelRun.km_run > 0) {
          km_run = fuelRun.km_run;
          const factor = Number(a.km_per_hour_factor || 10);
          hours_run = factor > 0 ? km_run / factor : 0;
          run_source = fuelRun.source || "fams_fuel";
        } else if (fuelRun.hours_run > 0) {
          hours_run = fuelRun.hours_run;
          run_source = fuelRun.source || "fams_fuel";
        }
      }
      rate = Number(a.hire_rate_per_hour || 0);
      amount = hours_run * rate;
    } else if (mode === "fixed_monthly") {
      rate = Number(a.hire_fixed_monthly || 0);
      amount = rate;
      run_source = "fixed_monthly";
    }

    if (amount <= 0 && mode !== "fixed_monthly") continue;

    lines.push({
      asset_code: a.asset_code,
      asset_name: a.asset_name,
      category: a.category || "Plant Hire",
      site_code: a.site_code || null,
      cost_center_code: a.cost_center_code || null,
      billing_mode: mode,
      metric_mode: String(a.metric_mode || "hours"),
      hours_run: Number(hours_run.toFixed(2)),
      km_run: Number(km_run.toFixed(2)),
      run_source,
      archived: Number(a.archived || 0),
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
  const metricExpr = sqlFuelMetricModeExpr("a");
  return db.prepare(`
    SELECT
      a.asset_code, a.asset_name, a.category, a.site_code, a.cost_center_code,
      a.hire_billing_mode, a.hire_rate_per_hour, a.hire_fixed_monthly,
      a.active, a.archived, ${metricExpr} AS utilization_mode
    FROM assets a
    WHERE ${hireAssetsWhereSql()}
    ORDER BY a.asset_code ASC
  `).all();
}

// Re-export for callers that used isHireCategory from here.
export { isHireCategory };
