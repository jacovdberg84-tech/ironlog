import { db } from "../db/client.js";
import { buildDueListFromPlans, groupActivePlansByAsset, resolveNextServiceForAssetPlans } from "../utils/serviceSchedule.js";
import { enrichDueRowsWithEstimates } from "../utils/maintenanceEstimates.js";

function getAssetCurrentHours(assetId) {
  const latestMeter = db.prepare(`
    SELECT closing_hours AS latest_closing FROM daily_hours
    WHERE asset_id = ? AND closing_hours IS NOT NULL AND DATE(work_date) IS NOT NULL
    ORDER BY work_date DESC, id DESC LIMIT 1
  `).get(assetId);
  if (latestMeter?.latest_closing != null) return Number(latestMeter.latest_closing);
  const fromDailyHours = db.prepare(`
    SELECT COALESCE(SUM(hours_run), 0) AS total_hours FROM daily_hours
    WHERE asset_id = ? AND is_used = 1 AND hours_run > 0
  `).get(assetId);
  return Number(fromDailyHours?.total_hours || 0);
}

const rows = db.prepare(`
  SELECT mp.id AS plan_id, mp.asset_id, mp.service_name, mp.interval_hours, mp.last_service_hours,
    a.asset_code, a.asset_name, a.category, mp.active
  FROM maintenance_plans mp
  JOIN assets a ON a.id = mp.asset_id
  WHERE mp.active = 1 AND a.active = 1 AND a.is_standby = 0 AND a.archived = 0
  ORDER BY a.asset_code ASC
`).all();

console.log("raw plan rows:", rows.length, rows.map((r) => ({ id: r.plan_id, asset: r.asset_code, active: r.active })));

const byAsset = groupActivePlansByAsset(rows);
console.log("byAsset size:", byAsset.size);
for (const [aid, plans] of byAsset) {
  const cur = getAssetCurrentHours(aid);
  const resolved = resolveNextServiceForAssetPlans(plans, cur, plans[0]?.asset_code);
  console.log("asset", plans[0]?.asset_code, "current", cur, "resolved", resolved);
}

const due = enrichDueRowsWithEstimates(
  db,
  buildDueListFromPlans(rows, getAssetCurrentHours, 50),
  { as_of: new Date().toISOString().slice(0, 10), history_days: 14 }
);
console.log("Total due rows:", due.length);
for (const r of due) {
  console.log(`  ${r.asset_code} rem=${r.remaining_hours} status=${r.status}`);
}
