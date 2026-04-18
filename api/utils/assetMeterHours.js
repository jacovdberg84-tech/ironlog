// Shared hour-meter resolution (fleet daily vs asset_hours), used by API routes.
import { db } from "../db/client.js";

/** @param {number} assetId */
export function getAssetCurrentHoursInfo(assetId) {
  const id = Number(assetId);
  if (!Number.isFinite(id) || id <= 0) {
    return { hours: 0, source: "none" };
  }

  const fromAssetHours = db.prepare(`
    SELECT total_hours
    FROM asset_hours
    WHERE asset_id = ?
  `).get(id);

  const assetHours = fromAssetHours?.total_hours == null ? null : Number(fromAssetHours.total_hours);

  const latestMeter = db.prepare(`
    SELECT closing_hours AS latest_closing, work_date AS latest_work_date
    FROM daily_hours
    WHERE asset_id = ?
      AND closing_hours IS NOT NULL
    ORDER BY work_date DESC, id DESC
    LIMIT 1
  `).get(id);
  const latestClosing = latestMeter?.latest_closing == null ? null : Number(latestMeter.latest_closing);
  const latestWorkDate = latestMeter?.latest_work_date ? String(latestMeter.latest_work_date) : null;

  if (assetHours != null && latestClosing != null) {
    if (Math.abs(assetHours - latestClosing) > 5000) {
      return { hours: latestClosing, source: "daily_closing", latest_work_date: latestWorkDate };
    }
    if (latestClosing >= assetHours) {
      return { hours: latestClosing, source: "daily_closing", latest_work_date: latestWorkDate };
    }
    return { hours: assetHours, source: "asset_hours", latest_work_date: latestWorkDate };
  }

  if (latestClosing != null) {
    return { hours: latestClosing, source: "daily_closing", latest_work_date: latestWorkDate };
  }
  if (assetHours != null) return { hours: assetHours, source: "asset_hours", latest_work_date: null };

  const fromDailyHours = db.prepare(`
    SELECT COALESCE(SUM(hours_run), 0) AS total_hours
    FROM daily_hours
    WHERE asset_id = ?
      AND is_used = 1
      AND hours_run > 0
  `).get(id);

  return {
    hours: Number(fromDailyHours?.total_hours || 0),
    source: "daily_sum",
    latest_work_date: null,
  };
}
