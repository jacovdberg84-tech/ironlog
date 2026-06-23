/**
 * Estimated service date from recent daily usage (matches maintenance history logic).
 */
export function addDaysYmd(ymd, days) {
  const d = new Date(`${String(ymd).trim()}T12:00:00`);
  d.setDate(d.getDate() + Number(days || 0));
  return d.toISOString().slice(0, 10);
}

export function estimateServiceDateFromUsage(db, opts = {}) {
  const assetId = Number(opts.asset_id || 0);
  const remaining = Number(opts.remaining_hours ?? opts.remaining ?? 0);
  const asOfLabel = String(opts.as_of || "").trim() || new Date().toISOString().slice(0, 10);
  const historyDays = Math.max(3, Number(opts.history_days || 14));

  if (!assetId || !/^\d{4}-\d{2}-\d{2}$/.test(asOfLabel)) {
    return { estimated_service_date: null, avg_daily_hours: 0 };
  }

  const startDate = addDaysYmd(asOfLabel, -(historyDays - 1));
  const avgRow = db.prepare(`
    SELECT
      COALESCE(SUM(hours_run), 0) AS total_run,
      COUNT(DISTINCT work_date) AS day_count
    FROM daily_hours
    WHERE asset_id = ?
      AND is_used = 1
      AND hours_run > 0
      AND work_date BETWEEN ? AND ?
  `).get(assetId, startDate, asOfLabel);

  const totalRun = Number(avgRow?.total_run || 0);
  const dayCount = Number(avgRow?.day_count || 0);
  const avgDaily = dayCount > 0 ? totalRun / dayCount : 0;
  const estDays = avgDaily > 0 ? Math.max(0, remaining / avgDaily) : null;
  const estimated_service_date = estDays == null ? null : addDaysYmd(asOfLabel, Math.round(estDays));

  return {
    estimated_service_date,
    avg_daily_hours: Number(avgDaily.toFixed(2)),
  };
}

export function enrichDueRowsWithEstimates(db, rows, opts = {}) {
  const list = Array.isArray(rows) ? rows : [];
  return list.map((r) => {
    const est = estimateServiceDateFromUsage(db, {
      asset_id: r.asset_id,
      remaining_hours: r.remaining_hours,
      as_of: opts.as_of,
      history_days: opts.history_days,
    });
    return { ...r, ...est };
  });
}
