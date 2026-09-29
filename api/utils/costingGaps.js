// IRONLOG/api/utils/costingGaps.js — holes in maintenance costing that someone
// should fill (or mark "not needed"). Found with plain queries so the list is
// right even when Borris is offline; Borris only helps fill them.
//
// Gap types:
//   service_unpriced   upcoming or overdue service with no price from any source
//   part_zero_cost     store part at $0 that sits in a service kit or was issued lately
//   service_no_labour  finished service work order with no labour hours booked
//   labour_rate_missing no default workshop labour rate set

const FINISHED_WO = ["completed", "approved", "closed"];

function hasTable(db, name) {
  return Boolean(db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?`).get(name));
}

function hasColumn(db, table, col) {
  return hasTable(db, table) && db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
}

export function ensureCostingGapSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS costing_gap_dismissals (
      gap_key TEXT PRIMARY KEY,
      reason TEXT,
      dismissed_by TEXT,
      dismissed_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `);
}

export function dismissCostingGap(db, gapKey, { reason = "", user = "" } = {}) {
  ensureCostingGapSchema(db);
  db.prepare(`
    INSERT INTO costing_gap_dismissals (gap_key, reason, dismissed_by, dismissed_at)
    VALUES (?, ?, ?, datetime('now'))
    ON CONFLICT(gap_key) DO UPDATE SET reason = excluded.reason, dismissed_by = excluded.dismissed_by, dismissed_at = excluded.dismissed_at
  `).run(String(gapKey), String(reason || "").slice(0, 300) || null, String(user || "") || null);
}

function unpricedServices(forecasts) {
  return (forecasts || [])
    .filter((r) => r && (r.needs_manual_input || String(r.forecast?.cost_source || "none") === "none"))
    .map((r) => ({
      key: `service_unpriced:${r.plan_id}`,
      type: "service_unpriced",
      severity: String(r.status || "").toUpperCase() === "OVERDUE" ? "high" : "medium",
      asset_code: r.asset_code,
      plan_id: Number(r.plan_id),
      title: `${r.asset_code} — ${r.service_name || "service"} has no price`,
      detail: Number(r.remaining_hours) < 0
        ? `Overdue by ${Math.abs(Number(r.remaining_hours)).toFixed(0)} h. No kit, template or past cost to price it.`
        : `Due in ${Number(r.remaining_hours).toFixed(0)} h. No kit, template or past cost to price it.`,
      borris: true,
    }));
}

function zeroCostParts(db, today) {
  if (!hasTable(db, "parts") || !hasColumn(db, "parts", "unit_cost")) return [];
  const rows = new Map();
  if (hasTable(db, "service_template_items") && hasTable(db, "service_templates")) {
    for (const r of db.prepare(`
      SELECT p.id, p.part_code, p.part_name, GROUP_CONCAT(DISTINCT t.name) AS templates
      FROM service_template_items i
      JOIN service_templates t ON t.id = i.service_template_id AND COALESCE(t.active, 1) = 1
      JOIN parts p ON p.id = i.stock_item_id
      WHERE COALESCE(p.unit_cost, 0) <= 0
      GROUP BY p.id
    `).all()) rows.set(r.id, { ...r, where: `in service kit ${r.templates}` });
  }
  if (hasTable(db, "stock_movements")) {
    const dateCol = hasColumn(db, "stock_movements", "created_at") ? "sm.created_at" : hasColumn(db, "stock_movements", "movement_date") ? "sm.movement_date" : null;
    const costCol = hasColumn(db, "stock_movements", "unit_cost_usd") ? "COALESCE(sm.unit_cost_usd, 0)" : "0";
    if (dateCol) {
      for (const r of db.prepare(`
        SELECT p.id, p.part_code, p.part_name, COUNT(*) AS issues, ABS(SUM(sm.quantity)) AS qty
        FROM stock_movements sm
        JOIN parts p ON p.id = sm.part_id
        WHERE sm.quantity < 0
          AND COALESCE(p.unit_cost, 0) <= 0
          AND ${costCol} <= 0
          AND date(${dateCol}) >= date(?, '-90 days')
        GROUP BY p.id
      `).all(today)) {
        const prev = rows.get(r.id);
        const issued = `issued ${Number(r.qty)} in the last 90 days`;
        rows.set(r.id, { ...r, where: prev ? `${prev.where}; ${issued}` : issued });
      }
    }
  }
  return [...rows.values()].map((r) => ({
    key: `part_zero_cost:${r.part_code}`,
    type: "part_zero_cost",
    severity: "medium",
    part_code: r.part_code,
    title: `${r.part_code} — ${r.part_name || "part"} costs $0`,
    detail: `${r.where.charAt(0).toUpperCase()}${r.where.slice(1)}. Services and jobs using it are under-costed.`,
    borris: true,
  }));
}

function servicesWithoutLabour(db, today) {
  if (!hasTable(db, "work_orders") || !hasColumn(db, "work_orders", "labor_hours")) return [];
  const doneAt = hasColumn(db, "work_orders", "completed_at") ? "COALESCE(w.completed_at, w.closed_at)" : "w.closed_at";
  return db.prepare(`
    SELECT w.id, a.asset_code, ${doneAt} AS done_at
    FROM work_orders w
    JOIN assets a ON a.id = w.asset_id
    WHERE LOWER(COALESCE(w.source, '')) = 'service'
      AND REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN (${FINISHED_WO.map((s) => `'${s}'`).join(", ")})
      AND COALESCE(w.labor_hours, 0) <= 0
      AND date(COALESCE(${doneAt}, w.opened_at)) >= date(?, '-60 days')
    ORDER BY ${doneAt} DESC
  `).all(today).map((r) => ({
    key: `service_no_labour:${r.id}`,
    type: "service_no_labour",
    severity: "low",
    asset_code: r.asset_code,
    work_order_id: Number(r.id),
    title: `WO #${r.id} (${r.asset_code}) has no labour hours`,
    detail: `Service finished ${String(r.done_at || "").slice(0, 10) || "recently"} with no labour booked, so its cost is parts only.`,
    borris: false,
  }));
}

function labourRateMissing(db) {
  if (!hasTable(db, "cost_settings")) return [];
  const row = db.prepare(`SELECT value FROM cost_settings WHERE key = 'labor_cost_per_hour_default' LIMIT 1`).get();
  if (Number(row?.value) > 0) return [];
  return [{
    key: "labour_rate_missing",
    type: "labour_rate_missing",
    severity: "high",
    title: "No default labour rate set",
    detail: "Labour on services and jobs is priced at a built-in $35/h until you set your own rate.",
    borris: false,
  }];
}

const SEVERITY = { high: 0, medium: 1, low: 2 };

/**
 * All open costing gaps, most important first, without the ones someone marked
 * "not needed". forecasts come from buildUpcomingServiceCostForecasts.
 */
export function findCostingGaps(db, { forecasts = [], today = new Date().toISOString().slice(0, 10) } = {}) {
  ensureCostingGapSchema(db);
  const dismissed = new Set(db.prepare(`SELECT gap_key FROM costing_gap_dismissals`).all().map((r) => r.gap_key));
  return [
    ...labourRateMissing(db),
    ...unpricedServices(forecasts),
    ...zeroCostParts(db, today),
    ...servicesWithoutLabour(db, today),
  ]
    .filter((g) => !dismissed.has(g.key))
    .sort((a, b) => (SEVERITY[a.severity] ?? 3) - (SEVERITY[b.severity] ?? 3));
}

export function summarizeCostingGaps(gaps) {
  const by = (t) => gaps.filter((g) => g.type === t).length;
  return {
    total: gaps.length,
    services_unpriced: by("service_unpriced"),
    parts_zero_cost: by("part_zero_cost"),
    services_no_labour: by("service_no_labour"),
    labour_rate_missing: by("labour_rate_missing") > 0,
  };
}

// The maintenance routes know how to read meters and price services; Borris's
// Ask lives elsewhere. The maintenance side registers a provider here so Ask
// can pull the same planning picture without importing route internals.
let planningSnapshotProvider = null;

export function setPlanningSnapshotProvider(fn) {
  planningSnapshotProvider = typeof fn === "function" ? fn : null;
}

/** { upcoming: [...], gaps: {...} } for planning questions, or null when unavailable. */
export function getPlanningSnapshot() {
  try {
    return planningSnapshotProvider ? planningSnapshotProvider() : null;
  } catch {
    return null;
  }
}

/** Question is about planning or costing (gets the planning snapshot and a longer answer). */
export function isPlanningQuestion(question) {
  return /\b(cost|costs|costing|price|prices|pricing|budget|spend|service|services|kit|kits|plan|planning|forecast|labour|labor|quote|interval|due|overdue|gap|gaps)\b/i.test(String(question || ""));
}
