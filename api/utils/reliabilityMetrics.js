/**
 * MTBF / LTTR helpers — aligned with breakdown downtime logs and breakdown work orders.
 *
 * MTBF = operating hours ÷ failure count
 * LTTR = downtime hours ÷ failure count
 *
 * Failures = distinct incidents with machine downtime recorded in the selected window.
 *
 * Daily downtime logs are authoritative because their log date is the actual date
 * lost. Older records without daily logs may use their breakdown-header total,
 * but only where the breakdown itself was reported in the selected window. A
 * linked work order remains audit information only: its elapsed calendar time
 * must never turn a historic, planned, or later-entered record into downtime for
 * the selected period.
 */

function parseTsMs(raw, fallbackMs) {
  if (raw == null || raw === "") return fallbackMs;
  const s = String(raw).trim();
  const ms = Date.parse(s.includes("T") ? s : s.replace(" ", "T"));
  return Number.isFinite(ms) ? ms : fallbackMs;
}

/** Wall-clock hours a work order was open, clipped to [start, end] (inclusive dates). */
export function woWallClockHoursInRange(openedAt, closedAt, start, end) {
  const periodStartMs = parseTsMs(`${start}T00:00:00`, NaN);
  const periodEndMs = parseTsMs(`${end}T23:59:59`, NaN);
  if (!Number.isFinite(periodStartMs) || !Number.isFinite(periodEndMs)) return 0;
  const openedMs = parseTsMs(openedAt, periodStartMs);
  const closedMs = parseTsMs(closedAt, periodEndMs);
  const winStart = Math.max(openedMs, periodStartMs);
  const winEnd = Math.min(closedMs, periodEndMs);
  if (winEnd <= winStart) return 0;
  return (winEnd - winStart) / 3600000;
}

function breakdownHeaderHours(row) {
  return Math.max(0, Number(row.header_downtime_hours ?? row.downtime_total_hours ?? row.downtime_hours ?? 0));
}

/**
 * @param {import('better-sqlite3').Database} db
 * @param {{ assetIds: number[], start: string, end: string, hasTable: (n:string)=>boolean, hasColumn: (t:string,c:string)=>boolean }} ctx
 */
export function buildReliabilityIncidentsForAssets(db, ctx) {
  const { assetIds, start, end, hasTable, hasColumn } = ctx;
  const empty = { incidents: [], byAsset: new Map() };
  if (!assetIds.length || !hasTable("breakdowns")) return empty;

  const marks = assetIds.map(() => "?").join(",");
  const canLogs = hasTable("breakdown_downtime_logs");
  const canWo = hasTable("work_orders");
  const dtCol = hasColumn("breakdowns", "downtime_total_hours")
    ? "downtime_total_hours"
    : (hasColumn("breakdowns", "downtime_hours") ? "downtime_hours" : null);
  const breakdownDateExpr = hasColumn("breakdowns", "breakdown_date")
    ? "b.breakdown_date"
    : "DATE(COALESCE(b.created_at, b.updated_at))";

  const breakdownRows = db.prepare(`
    SELECT
      b.id AS breakdown_id,
      b.asset_id,
      ${breakdownDateExpr} AS breakdown_date,
      b.description,
      b.primary_work_order_id,
      ${dtCol ? `COALESCE(b.${dtCol}, 0)` : "0"} AS header_downtime_hours
    FROM breakdowns b
    WHERE b.asset_id IN (${marks})
  `).all(...assetIds);

  const logDtByBreakdown = new Map();
  if (canLogs) {
    for (const r of db.prepare(`
      SELECT l.breakdown_id, COALESCE(SUM(l.hours_down), 0) AS log_dt
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      WHERE b.asset_id IN (${marks})
        AND DATE(l.log_date) BETWEEN DATE(?) AND DATE(?)
        AND COALESCE(l.hours_down, 0) > 0
      GROUP BY l.breakdown_id
    `).all(...assetIds, start, end)) {
      logDtByBreakdown.set(Number(r.breakdown_id), Number(r.log_dt || 0));
    }
  }

  const woById = new Map();
  const woByBreakdownRef = new Map();
  if (canWo) {
    const woHasCompleted = hasColumn("work_orders", "completed_at");
    const woCols = [
      "w.id",
      "w.asset_id",
      "w.reference_id",
      "w.source",
      "w.status",
      "w.opened_at",
      "w.closed_at",
      woHasCompleted ? "w.completed_at" : "NULL AS completed_at",
      hasColumn("work_orders", "labor_hours") ? "w.labor_hours" : "0 AS labor_hours",
    ].join(", ");
    for (const wo of db.prepare(`
      SELECT ${woCols}
      FROM work_orders w
      WHERE w.asset_id IN (${marks})
        AND LOWER(TRIM(COALESCE(w.source, ''))) = 'breakdown'
    `).all(...assetIds)) {
      woById.set(Number(wo.id), wo);
      const ref = Number(wo.reference_id || 0);
      if (ref > 0) woByBreakdownRef.set(ref, wo);
    }
  }

  const incidents = [];
  const byAsset = new Map();

  for (const row of breakdownRows) {
    const breakdown_id = Number(row.breakdown_id || 0);
    const asset_id = Number(row.asset_id || 0);
    if (!breakdown_id || !asset_id) continue;

    const breakdown_date = String(row.breakdown_date || "").slice(0, 10);
    const log_dt = Number(logDtByBreakdown.get(breakdown_id) || 0);
    const header_dt = breakdownHeaderHours(row);
    const inReportWindow = breakdown_date >= start && breakdown_date <= end;

    const primaryWoId = Number(row.primary_work_order_id || 0);
    const wo = (primaryWoId && woById.get(primaryWoId))
      || woByBreakdownRef.get(breakdown_id)
      || null;

    let downtime_hours = 0;
    let downtime_source = "none";

    if (log_dt > 0) {
      downtime_hours = log_dt;
      downtime_source = "downtime_logs";
    } else if (inReportWindow && header_dt > 0) {
      downtime_hours = header_dt;
      downtime_source = "breakdown_header";
    }

    if (downtime_hours <= 0) continue;

    const incident = {
      breakdown_id,
      asset_id,
      breakdown_date,
      description: String(row.description || ""),
      work_order_id: wo ? Number(wo.id) : (primaryWoId || null),
      work_order_status: wo ? String(wo.status || "") : null,
      work_order_opened_at: wo ? (wo.opened_at || null) : null,
      work_order_closed_at: wo ? (wo.completed_at || wo.closed_at || null) : null,
      work_order_labor_hours: wo ? Number(wo.labor_hours || 0) : null,
      downtime_hours: Number(downtime_hours.toFixed(2)),
      downtime_source,
      log_downtime_in_period: Number(log_dt.toFixed(2)),
      header_downtime_hours: Number(header_dt.toFixed(2)),
    };
    incidents.push(incident);

    const cur = byAsset.get(asset_id) || { failure_count: 0, downtime_hours: 0 };
    cur.failure_count += 1;
    cur.downtime_hours += downtime_hours;
    byAsset.set(asset_id, cur);
  }

  return { incidents, byAsset };
}

export function round2(n) {
  return Number(Number(n).toFixed(2));
}

export function computeMtbfLttr(operating_hours, failure_count, downtime_hours) {
  const fc = Number(failure_count || 0);
  const op = Number(operating_hours || 0);
  const dt = Number(downtime_hours || 0);
  const mtbf_hours = fc > 0 ? op / fc : null;
  const lttr_hours = fc > 0 ? dt / fc : null;
  return {
    mtbf_hours: mtbf_hours == null ? null : round2(mtbf_hours),
    lttr_hours: lttr_hours == null ? null : round2(lttr_hours),
  };
}

/**
 * Reliability rows (one per machine) added up per group: model, area, class or site.
 * Each group gets its machine count, failures, operating and downtime hours, MTBF, LTTR (MTTR),
 * availability (available ÷ scheduled hours) and maintenance cost per operating hour.
 */
export function groupReliabilityRows(rows, keyOf) {
  const groups = new Map();
  for (const r of rows || []) {
    const key = String(keyOf(r) || "").trim() || "Not set";
    const g = groups.get(key) || { key, asset_count: 0, failure_count: 0, operating_hours: 0, downtime_hours: 0, scheduled_hours: 0, available_hours: 0, maintenance_cost: 0, has_cost: false };
    g.asset_count += 1;
    g.failure_count += Number(r.failure_count || 0);
    g.operating_hours += Number(r.operating_hours || 0);
    g.downtime_hours += Number(r.downtime_hours || 0);
    g.scheduled_hours += Number(r.scheduled_hours || 0);
    g.available_hours += Number(r.available_hours || 0);
    if (r.maintenance_cost != null) { g.maintenance_cost += Number(r.maintenance_cost || 0); g.has_cost = true; }
    groups.set(key, g);
  }
  return [...groups.values()].map((g) => ({
    key: g.key,
    asset_count: g.asset_count,
    failure_count: g.failure_count,
    operating_hours: round2(g.operating_hours),
    downtime_hours: round2(g.downtime_hours),
    ...computeMtbfLttr(g.operating_hours, g.failure_count, g.downtime_hours),
    availability_pct: g.scheduled_hours > 0 ? Number(((g.available_hours / g.scheduled_hours) * 100).toFixed(1)) : null,
    maintenance_cost: g.has_cost ? round2(g.maintenance_cost) : null,
    cost_per_hour: g.has_cost && g.operating_hours > 0 ? round2(g.maintenance_cost / g.operating_hours) : null,
  })).sort((a, b) => b.failure_count - a.failure_count || String(a.key).localeCompare(String(b.key)));
}

/**
 * Failures per failed component (failure codes): how often each component fails across the machines
 * in scope, its downtime and MTBF (all the machines' operating hours ÷ that component's failures).
 */
export function componentReliability(incidents, operatingHours, labels = {}) {
  const groups = new Map();
  for (const inc of incidents || []) {
    const key = inc.fc_component || "";
    const g = groups.get(key) || { code: key || null, failure_count: 0, downtime_hours: 0 };
    g.failure_count += 1;
    g.downtime_hours += Number(inc.downtime_hours || 0);
    groups.set(key, g);
  }
  return [...groups.values()].map((g) => ({
    code: g.code,
    label: g.code ? labels[g.code] || g.code : "No failure code yet",
    failure_count: g.failure_count,
    downtime_hours: round2(g.downtime_hours),
    ...computeMtbfLttr(operatingHours, g.failure_count, g.downtime_hours),
  })).sort((a, b) => (a.code ? 0 : 1) - (b.code ? 0 : 1) || b.failure_count - a.failure_count);
}
