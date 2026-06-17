/** Downtime cost = down-days × daily scheduled hours × asset downtime $/hr. */

export function monthStart(period) {
  return `${String(period).slice(0, 7)}-01`;
}

export function monthEnd(period) {
  const [y, m] = String(period).split("-").map((n) => Number(n));
  return new Date(Date.UTC(y, m, 0)).toISOString().slice(0, 10);
}

export function eachDateInclusiveYmd(startStr, endStr) {
  const out = [];
  const d = new Date(`${startStr}T12:00:00`);
  const end = new Date(`${endStr}T12:00:00`);
  while (d <= end) {
    out.push(d.toISOString().slice(0, 10));
    d.setDate(d.getDate() + 1);
  }
  return out;
}

function normWoStatus(status) {
  return String(status || "").trim().toLowerCase().replace(/\s+/g, "_");
}

function isOpenWoStatus(status) {
  return ["open", "assigned", "in_progress"].includes(normWoStatus(status));
}

/**
 * @param {import('better-sqlite3').Database} db
 * @param {string} period YYYY-MM
 * @param {{ defaultScheduledHours?: number, downtimeRateDefault?: number }} opts
 */
export function buildScheduledDowntimeCost(db, period, opts = {}) {
  const defaultScheduledHours = Number(opts.defaultScheduledHours ?? 10);
  const downtimeRateDefault = Number(opts.downtimeRateDefault ?? 120);
  const start = monthStart(period);
  const end = monthEnd(period);
  const periodDates = eachDateInclusiveYmd(start, end);

  const hasTable = (name) =>
    Boolean(db.prepare(`SELECT 1 FROM sqlite_master WHERE type='table' AND name=?`).get(name));
  if (!hasTable("assets")) {
    return { detail: [], financeRows: [], total_hours: 0, total_cost: 0 };
  }

  const assetsHasDownCost = db
    .prepare(`PRAGMA table_info(assets)`)
    .all()
    .some((c) => String(c.name) === "downtime_cost_per_hour");
  const assetsHasSite = db
    .prepare(`PRAGMA table_info(assets)`)
    .all()
    .some((c) => String(c.name) === "site_code");
  const assetsHasCC = db
    .prepare(`PRAGMA table_info(assets)`)
    .all()
    .some((c) => String(c.name) === "cost_center_code");

  const assets = db.prepare(`
    SELECT
      id,
      asset_code,
      asset_name,
      category,
      ${assetsHasSite ? "site_code" : "NULL AS site_code"},
      ${assetsHasCC ? "cost_center_code" : "NULL AS cost_center_code"},
      ${assetsHasDownCost ? "COALESCE(downtime_cost_per_hour, ?)" : String(downtimeRateDefault)} AS downtime_rate
    FROM assets
    WHERE COALESCE(archived, 0) = 0
      AND active = 1
  `).all(...(assetsHasDownCost ? [downtimeRateDefault] : []));

  const assetById = new Map(assets.map((a) => [Number(a.id), a]));

  /** @type {Set<string>} assetId|YYYY-MM-DD */
  const downDayKeys = new Set();

  if (hasTable("breakdown_downtime_logs") && hasTable("breakdowns")) {
    const logDays = db.prepare(`
      SELECT b.asset_id, l.log_date AS work_date
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      WHERE l.log_date BETWEEN DATE(?) AND DATE(?)
        AND COALESCE(l.hours_down, 0) > 0
    `).all(start, end);
    for (const r of logDays) {
      downDayKeys.add(`${Number(r.asset_id)}|${r.work_date}`);
    }
  }

  if (hasTable("work_orders")) {
    const woCols = db.prepare(`PRAGMA table_info(work_orders)`).all();
    const hasOpened = woCols.some((c) => String(c.name) === "opened_at");
    const hasCreated = woCols.some((c) => String(c.name) === "created_at");
    let openExpr = "datetime('now')";
    if (hasOpened && hasCreated) {
      openExpr = "COALESCE(NULLIF(TRIM(w.opened_at), ''), w.created_at, datetime('now'))";
    } else if (hasOpened) {
      openExpr = "COALESCE(NULLIF(TRIM(w.opened_at), ''), datetime('now'))";
    } else if (hasCreated) {
      openExpr = "COALESCE(w.created_at, datetime('now'))";
    }

    const woRows = db.prepare(`
      SELECT
        w.asset_id,
        w.source,
        w.status,
        DATE(${openExpr}) AS opened_date,
        CASE
          WHEN w.closed_at IS NULL OR TRIM(COALESCE(w.closed_at, '')) = '' THEN DATE(?)
          ELSE DATE(w.closed_at)
        END AS closed_date
      FROM work_orders w
      WHERE DATE(${openExpr}) <= DATE(?)
        AND (
          w.closed_at IS NULL OR TRIM(COALESCE(w.closed_at, '')) = ''
          OR DATE(w.closed_at) >= DATE(?)
        )
    `).all(end, end, start);

    for (const wo of woRows) {
      const assetId = Number(wo.asset_id || 0);
      if (!assetId || !assetById.has(assetId)) continue;

      const includeInterval =
        String(wo.source || "").toLowerCase() === "breakdown"
        || isOpenWoStatus(wo.status);

      if (!includeInterval) continue;

      const opened = String(wo.opened_date || start);
      const closed = String(wo.closed_date || end);
      const from = opened > start ? opened : start;
      const to = closed < end ? closed : end;
      for (const d of eachDateInclusiveYmd(from, to)) {
        downDayKeys.add(`${assetId}|${d}`);
      }
    }
  }

  const scheduledStmt = hasTable("daily_hours")
    ? db.prepare(`
        SELECT scheduled_hours
        FROM daily_hours
        WHERE asset_id = ? AND work_date = ?
        ORDER BY id DESC
        LIMIT 1
      `)
    : null;

  const byAsset = new Map();

  for (const key of downDayKeys) {
    const [idStr, workDate] = key.split("|");
    const assetId = Number(idStr);
    const asset = assetById.get(assetId);
    if (!asset) continue;

    let sched = defaultScheduledHours;
    if (scheduledStmt) {
      const row = scheduledStmt.get(assetId, workDate);
      const v = Number(row?.scheduled_hours);
      if (Number.isFinite(v) && v > 0) sched = v;
    }

    const prev = byAsset.get(assetId) || {
      asset_id: assetId,
      asset_code: asset.asset_code,
      asset_name: asset.asset_name,
      category: asset.category,
      site_code: asset.site_code,
      cost_center_code: asset.cost_center_code,
      downtime_rate: Number(asset.downtime_rate || downtimeRateDefault),
      down_days: 0,
      downtime_hours: 0,
      downtime_cost: 0,
    };
    prev.down_days += 1;
    prev.downtime_hours += sched;
    prev.downtime_cost += sched * prev.downtime_rate;
    byAsset.set(assetId, prev);
  }

  const detail = Array.from(byAsset.values())
    .map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      category: r.category || "",
      down_days: r.down_days,
      downtime_hours: Number(r.downtime_hours.toFixed(2)),
      downtime_rate: Number(r.downtime_rate.toFixed(2)),
      downtime_cost: Number(r.downtime_cost.toFixed(2)),
    }))
    .filter((r) => r.downtime_hours > 0)
    .sort((a, b) => b.downtime_cost - a.downtime_cost);

  const financeMap = new Map();
  for (const r of byAsset.values()) {
    const amount = Number(r.downtime_cost.toFixed(2));
    if (amount <= 0) continue;
    const fKey = `${r.site_code || ""}|${r.cost_center_code || ""}|${r.category || ""}`;
    const prev = financeMap.get(fKey) || {
      site_code: r.site_code || null,
      cost_center_code: r.cost_center_code || null,
      equipment_type: r.category || null,
      amount: 0,
    };
    prev.amount += amount;
    financeMap.set(fKey, prev);
  }

  const financeRows = Array.from(financeMap.values()).map((r) => ({
    ...r,
    amount: Number(r.amount.toFixed(2)),
  }));

  const total_hours = Number(detail.reduce((s, r) => s + r.downtime_hours, 0).toFixed(2));
  const total_cost = Number(detail.reduce((s, r) => s + r.downtime_cost, 0).toFixed(2));

  return { detail, financeRows, total_hours, total_cost, period, start, end };
}
