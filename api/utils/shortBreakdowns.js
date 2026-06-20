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

export function shiftDateYmd(dateStr, days) {
  const d = new Date(`${String(dateStr).trim()}T12:00:00`);
  d.setDate(d.getDate() + Math.round(Number(days) || 0));
  return d.toISOString().slice(0, 10);
}
