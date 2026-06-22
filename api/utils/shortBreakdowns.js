/**
 * Short breakdowns logged via POST /api/breakdowns/short-complete
 * (downtime log notes: "Short breakdown — …").
 */
export function listShortBreakdownsForDate(db, logDate) {
  const hasLogs = Boolean(
    db.prepare(`SELECT 1 AS ok FROM sqlite_master WHERE type='table' AND name='breakdown_downtime_logs' LIMIT 1`).get()
  );
  if (!hasLogs) return [];

  return db.prepare(`
    SELECT
      a.asset_code,
      a.asset_name,
      b.description,
      b.component,
      b.critical,
      l.hours_down,
      l.notes,
      l.log_date,
      b.breakdown_date
    FROM breakdown_downtime_logs l
    JOIN breakdowns b ON b.id = l.breakdown_id
    JOIN assets a ON a.id = b.asset_id
    WHERE l.log_date = ?
      AND COALESCE(l.notes, '') LIKE 'Short breakdown%'
    ORDER BY a.asset_code ASC, l.id ASC
  `).all(logDate).map((r) => {
    const notes = String(r.notes || "");
    const desc = String(r.description || "");
    const shortDesc = notes.replace(/^Short breakdown\s*[—-]\s*/i, "").trim() || desc;
    return {
      asset_code: String(r.asset_code || ""),
      asset_name: String(r.asset_name || ""),
      description: shortDesc,
      component: r.component ? String(r.component) : null,
      critical: Boolean(r.critical),
      hours_down: Number(r.hours_down || 0),
      log_date: String(r.log_date || ""),
      breakdown_date: r.breakdown_date ? String(r.breakdown_date) : null,
    };
  });
}

export function listPlannedMaintenanceForDate(db, logDate) {
  const hasWo = Boolean(
    db.prepare(`SELECT 1 AS ok FROM sqlite_master WHERE type='table' AND name='work_orders' LIMIT 1`).get()
  );
  if (!hasWo) return [];

  const hasPlans = Boolean(
    db.prepare(`SELECT 1 AS ok FROM sqlite_master WHERE type='table' AND name='maintenance_plans' LIMIT 1`).get()
  );
  const planJoin = hasPlans
    ? "LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id"
    : "";
  const serviceSelect = hasPlans ? "mp.service_name" : "NULL AS service_name";

  return db.prepare(`
    SELECT
      w.id AS wo_id,
      a.asset_code,
      a.asset_name,
      ${serviceSelect},
      w.status,
      w.opened_at,
      w.completed_at,
      w.closed_at
    FROM work_orders w
    JOIN assets a ON a.id = w.asset_id
    ${planJoin}
    WHERE LOWER(COALESCE(w.source, '')) = 'service'
      AND (
        DATE(COALESCE(w.opened_at, '')) = ?
        OR DATE(COALESCE(w.completed_at, w.closed_at, '')) = ?
      )
    ORDER BY a.asset_code ASC, w.id ASC
  `).all(logDate, logDate).map((r) => {
    const status = String(r.status || "").replace(/_/g, " ").trim();
    const service = r.service_name ? String(r.service_name).trim() : "Service";
    const woId = Number(r.wo_id || 0);
    const detail = woId ? `${service} (WO #${woId}${status ? `, ${status}` : ""})` : service;
    return {
      asset_code: String(r.asset_code || ""),
      asset_name: String(r.asset_name || ""),
      service_name: service,
      wo_id: woId,
      status,
      description: detail,
    };
  });
}

export function shiftDateYmd(dateStr, days) {
  const d = new Date(`${String(dateStr).trim()}T12:00:00`);
  d.setDate(d.getDate() + Math.round(Number(days) || 0));
  return d.toISOString().slice(0, 10);
}
