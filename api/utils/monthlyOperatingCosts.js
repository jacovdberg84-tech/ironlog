/** Monthly operating cost actuals — same source queries as Reports → Cost Monthly XLSX. */

import { buildPlantHireFinanceRows } from "./plantHire.js";

export function monthPeriodBounds(monthStr) {
  const [y, m] = String(monthStr).split("-").map((n) => Number(n));
  const start = new Date(Date.UTC(y, m - 1, 1)).toISOString().slice(0, 10);
  const end = new Date(Date.UTC(y, m, 0)).toISOString().slice(0, 10);
  return { start, end };
}

export function readCostDefaults(db) {
  const d = {
    fuel_cost_per_liter_default: 1.5,
    lube_cost_per_qty_default: 4.0,
    labor_cost_per_hour_default: 35.0,
    downtime_cost_per_hour_default: 120.0,
  };
  try {
    const rows = db.prepare(`
      SELECT key, value FROM cost_settings
      WHERE key IN (
        'fuel_cost_per_liter_default',
        'lube_cost_per_qty_default',
        'labor_cost_per_hour_default',
        'downtime_cost_per_hour_default'
      )
    `).all();
    for (const r of rows) {
      const k = String(r.key || "").trim();
      const v = Number(r.value);
      if (k && Number.isFinite(v)) d[k] = v;
    }
  } catch { /* ignore */ }
  return d;
}

function hasTable(db, name) {
  return Boolean(db.prepare(`SELECT 1 FROM sqlite_master WHERE type='table' AND name=?`).get(name));
}

/**
 * Per-asset costs for a calendar period (inclusive YYYY-MM-DD bounds).
 * Mirrors /api/reports/cost-monthly.xlsx buildAssetCosts.
 */
export function buildAssetPeriodCosts(db, start, end, defaults = null) {
  const defs = defaults || readCostDefaults(db);
  const map = new Map();

  const ensure = (r) => {
    const code = String(r.asset_code || "UNLINKED");
    if (!map.has(code)) {
      map.set(code, {
        asset_code: code,
        asset_name: r.asset_name || "Unlinked",
        category: r.category || "Unassigned",
        fuel_cost: 0,
        lube_cost: 0,
        parts_cost: 0,
        labor_hours: 0,
        labor_cost: 0,
        downtime_hours: 0,
        downtime_cost: 0,
        total_cost: 0,
      });
    }
    return map.get(code);
  };

  if (hasTable(db, "fuel_logs") && hasTable(db, "assets")) {
    const fuelRows = db.prepare(`
      SELECT a.asset_code, a.asset_name, a.category,
        COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS fuel_cost
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date BETWEEN ? AND ?
      GROUP BY a.id
    `).all(defs.fuel_cost_per_liter_default, start, end);
    for (const r of fuelRows) ensure(r).fuel_cost += Number(r.fuel_cost || 0);
  }

  if (hasTable(db, "oil_logs") && hasTable(db, "assets")) {
    const lubeRows = db.prepare(`
      SELECT a.asset_code, a.asset_name, a.category,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.id
    `).all(defs.lube_cost_per_qty_default, start, end);
    for (const r of lubeRows) ensure(r).lube_cost += Number(r.lube_cost || 0);
  }

  if (hasTable(db, "stock_movements") && hasTable(db, "parts")) {
    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
    const partsRows = db.prepare(`
      SELECT
        COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
        COALESCE(a.asset_name, 'Unlinked') AS asset_name,
        COALESCE(a.category, 'Unassigned') AS category,
        COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS parts_cost
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      LEFT JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
      LEFT JOIN assets a ON a.id = w.asset_id
      WHERE sm.movement_type = 'out'
        AND ${smDateExpr} BETWEEN ? AND ?
      GROUP BY a.id
    `).all(start, end);
    for (const r of partsRows) ensure(r).parts_cost += Number(r.parts_cost || 0);
  }

  if (hasTable(db, "work_orders") && hasTable(db, "assets")) {
    const laborRows = db.prepare(`
      SELECT a.asset_code, a.asset_name, a.category,
        COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS labor_hours,
        COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE DATE(COALESCE(w.completed_at, w.closed_at)) BETWEEN ? AND ?
        AND w.status IN ('completed', 'approved', 'closed')
      GROUP BY a.id
    `).all(defs.labor_cost_per_hour_default, start, end);
    for (const r of laborRows) {
      const row = ensure(r);
      row.labor_hours += Number(r.labor_hours || 0);
      row.labor_cost += Number(r.labor_cost || 0);
    }
  }

  if (hasTable(db, "breakdown_downtime_logs") && hasTable(db, "breakdowns") && hasTable(db, "assets")) {
    const downtimeRows = db.prepare(`
      SELECT a.asset_code, a.asset_name, a.category,
        COALESCE(SUM(l.hours_down), 0) AS downtime_hours,
        COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS downtime_cost
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date BETWEEN ? AND ?
      GROUP BY a.id
    `).all(defs.downtime_cost_per_hour_default, start, end);
    for (const r of downtimeRows) {
      const row = ensure(r);
      row.downtime_hours += Number(r.downtime_hours || 0);
      row.downtime_cost += Number(r.downtime_cost || 0);
    }
  }

  return Array.from(map.values())
    .map((r) => {
      const total =
        Number(r.fuel_cost || 0) +
        Number(r.lube_cost || 0) +
        Number(r.parts_cost || 0) +
        Number(r.labor_cost || 0) +
        Number(r.downtime_cost || 0);
      return {
        ...r,
        fuel_cost: Number(r.fuel_cost.toFixed(2)),
        lube_cost: Number(r.lube_cost.toFixed(2)),
        parts_cost: Number(r.parts_cost.toFixed(2)),
        labor_hours: Number(r.labor_hours.toFixed(2)),
        labor_cost: Number(r.labor_cost.toFixed(2)),
        downtime_hours: Number(r.downtime_hours.toFixed(2)),
        downtime_cost: Number(r.downtime_cost.toFixed(2)),
        total_cost: Number(total.toFixed(2)),
      };
    })
    .filter((r) => r.total_cost > 0);
}

/** Roll per-asset rows up to finance categories for budget meeting / BVA. */
export function rollupOperatingCategoryRows(assetRows, plantHireAmount = 0) {
  const totals = { parts: 0, labor: 0, fuel: 0, lube: 0, downtime: 0, plant_hire: Number(plantHireAmount || 0) };
  for (const r of assetRows) {
    totals.parts += Number(r.parts_cost || 0);
    totals.labor += Number(r.labor_cost || 0);
    totals.fuel += Number(r.fuel_cost || 0);
    totals.lube += Number(r.lube_cost || 0);
    totals.downtime += Number(r.downtime_cost || 0);
  }
  const rows = Object.entries(totals)
    .filter(([, amount]) => Number(amount) > 0)
    .map(([category, amount]) => ({ category, amount: Number(amount.toFixed(2)) }))
    .sort((a, b) => b.amount - a.amount);
  const total = Number(rows.reduce((s, r) => s + r.amount, 0).toFixed(2));
  return { rows, total };
}

/**
 * Category actuals for a calendar month (YYYY-MM).
 * Uses cost-monthly source data so Word doc matches Cost Monthly XLSX.
 */
export function buildMonthlyOperatingActuals(db, period) {
  const { start, end } = monthPeriodBounds(period);
  const defaults = readCostDefaults(db);
  const assetRows = buildAssetPeriodCosts(db, start, end, defaults);
  const plantHire = buildPlantHireFinanceRows(db, period).reduce(
    (s, r) => s + Number(r.actual_amount || 0),
    0,
  );
  return rollupOperatingCategoryRows(assetRows, plantHire);
}

/** Downtime detail for Word doc — logged breakdown hours × asset downtime rate. */
export function buildBreakdownDowntimeDetail(db, period) {
  const { start, end } = monthPeriodBounds(period);
  const defs = readCostDefaults(db);
  if (!hasTable(db, "breakdown_downtime_logs") || !hasTable(db, "breakdowns") || !hasTable(db, "assets")) {
    return { detail: [], total_hours: 0, total_cost: 0 };
  }

  const rows = db.prepare(`
    SELECT
      a.asset_code,
      a.asset_name,
      a.category,
      COUNT(DISTINCT l.log_date) AS down_days,
      COALESCE(SUM(l.hours_down), 0) AS downtime_hours,
      COALESCE(a.downtime_cost_per_hour, ?) AS downtime_rate,
      COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS downtime_cost
    FROM breakdown_downtime_logs l
    JOIN breakdowns b ON b.id = l.breakdown_id
    JOIN assets a ON a.id = b.asset_id
    WHERE l.log_date BETWEEN ? AND ?
      AND COALESCE(l.hours_down, 0) > 0
    GROUP BY a.id
    HAVING downtime_hours > 0
    ORDER BY downtime_cost DESC
  `).all(defs.downtime_cost_per_hour_default, defs.downtime_cost_per_hour_default, start, end);

  const detail = rows.map((r) => ({
    asset_code: r.asset_code,
    asset_name: r.asset_name,
    category: r.category || "",
    down_days: Number(r.down_days || 0),
    downtime_hours: Number(Number(r.downtime_hours || 0).toFixed(2)),
    downtime_rate: Number(Number(r.downtime_rate || 0).toFixed(2)),
    downtime_cost: Number(Number(r.downtime_cost || 0).toFixed(2)),
  }));

  const total_hours = Number(detail.reduce((s, r) => s + r.downtime_hours, 0).toFixed(2));
  const total_cost = Number(detail.reduce((s, r) => s + r.downtime_cost, 0).toFixed(2));
  return { detail, total_hours, total_cost };
}
