// IRONLOG/api/routes/dashboard.routes.js
import { db } from "../db/client.js";
import { ensureAuditTable } from "../utils/audit.js";
import { andDailyHoursFleetHoursOnly, andAssetFleetHoursOnly } from "../utils/fleetHoursKpiScope.js";
import { listDailyPrestarts, PRESTART_DEDUCTION_HOURS } from "../utils/prestartDaily.js";
import { getRunFromFuelRows } from "../utils/fuelRunFromLogs.js";
import { sqlFuelMetricModeExpr } from "../utils/fuelMetricMode.js";
import { sqlIncludeArchivedHireAssets } from "../utils/hiredEquipment.js";
import { ensureCostAllocationSchema } from "../utils/costAllocation.js";
import { createManagementSummary, styleManagementDetailSheet } from "../utils/managementWorkbook.js";
import { buildShiftScenario } from "../utils/shiftScenario.js";
import { registerAssetKpiRangeBuilder } from "../utils/assetKpiRangeProvider.js";
import registerOverviewRoutes from "./dashboard/overview.routes.js";
import registerAssetKpiRoutes from "./dashboard/asset-kpi.routes.js";
import registerLubeRoutes from "./dashboard/lube.routes.js";
import registerCostSettingsRoutes from "./dashboard/cost-settings.routes.js";
import registerFuelRoutes from "./dashboard/fuel.routes.js";
import { holdsAnyRole } from "../utils/request.js";

function todayYYYYMMDD() {
  return new Date().toISOString().slice(0, 10);
}

function monthRangeFromYYYYMM(monthStr) {
  const [y, m] = String(monthStr).split("-").map((n) => Number(n));
  const start = new Date(Date.UTC(y, m - 1, 1));
  const end = new Date(Date.UTC(y, m, 0));
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

function monthIdFromDateStr(dateStr) {
  return String(dateStr || "").slice(0, 7);
}

/** First day of month for a YYYY-MM-DD date string. */
function monthStartIso(dateStr) {
  const s = String(dateStr || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return `${s.slice(0, 7)}-01`;
}

function isToyotaHiluxAsset(asset) {
  const code = String(asset?.asset_code || "").toLowerCase();
  const name = String(asset?.asset_name || "").toLowerCase();
  return name.includes("toyota") && name.includes("hilux") || code.includes("hilux");
}

function eachDateInclusiveYMD(startStr, endStr, fn) {
  const start = new Date(`${startStr}T12:00:00`);
  const end = new Date(`${endStr}T12:00:00`);
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    fn(d.toISOString().slice(0, 10));
  }
}

function siteCodeFromReq(req) {
  return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
}

export default async function dashboardRoutes(app) {
  ensureCostAllocationSchema(db);
  ensureAuditTable(db);

  function hasColumn(table, col) {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name) === col);
  }
  function getBreakdownDowntimeColumn() {
    const rows = db.prepare(`PRAGMA table_info(breakdowns)`).all();
    const names = new Set(rows.map((r) => String(r.name)));
    if (names.has("downtime_total_hours")) return "downtime_total_hours";
    if (names.has("downtime_hours")) return "downtime_hours";
    return "downtime_hours";
  }

  function ensureColumn(table, colName, colDef) {
    if (!hasColumn(table, colName)) {
      db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
    }
  }

  function getRole(req) {
    return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
  }

  function requireRoles(req, reply, roles) {
    const role = getRole(req);
    if (!holdsAnyRole(req, roles)) {
      reply.code(403).send({ error: `role '${role || "unknown"}' not allowed` });
      return false;
    }
    return true;
  }

  db.prepare(`
    CREATE TABLE IF NOT EXISTS lube_type_mappings (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      oil_key TEXT NOT NULL UNIQUE,
      part_code TEXT NOT NULL,
      updated_by TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  function slaPriority(status, ageHours) {
    const s = String(status || "").toLowerCase();
    const a = Number(ageHours || 0);
    if (s === "completed" && a > 48) return "P1";
    if (s === "in_progress" && a > 72) return "P1";
    if ((s === "open" || s === "assigned") && a > 72) return "P1";

    if (s === "completed" && a > 24) return "P2";
    if (s === "in_progress" && a > 48) return "P2";
    if ((s === "open" || s === "assigned") && a > 48) return "P2";

    return "P3";
  }

  // OEM baseline fuel benchmark (L/hr) per asset.
  ensureColumn("assets", "baseline_fuel_l_per_hour", "baseline_fuel_l_per_hour REAL DEFAULT 5.0");
  ensureColumn("assets", "baseline_fuel_km_per_l", "baseline_fuel_km_per_l REAL DEFAULT 2.0");
  ensureColumn("assets", "fuel_cost_per_liter", "fuel_cost_per_liter REAL");
  ensureColumn("assets", "downtime_cost_per_hour", "downtime_cost_per_hour REAL");
  ensureColumn("daily_hours", "input_unit", "input_unit TEXT DEFAULT 'hours'");
  ensureColumn("assets", "utilization_mode", "utilization_mode TEXT DEFAULT 'hours'");
  ensureColumn("assets", "km_per_hour_factor", "km_per_hour_factor REAL DEFAULT 10.0");
  ensureColumn("parts", "unit_cost", "unit_cost REAL DEFAULT 0");
  ensureColumn("oil_logs", "unit_cost", "unit_cost REAL");
  ensureColumn("fuel_logs", "unit_cost_per_liter", "unit_cost_per_liter REAL");
  ensureColumn("fuel_logs", "hours_run", "hours_run REAL");
  ensureColumn("fuel_logs", "meter_run_value", "meter_run_value REAL");
  ensureColumn("fuel_logs", "meter_unit", "meter_unit TEXT");
  ensureColumn("fuel_logs", "open_meter_value", "open_meter_value REAL");
  ensureColumn("fuel_logs", "close_meter_value", "close_meter_value REAL");
  ensureColumn("work_orders", "labor_hours", "labor_hours REAL DEFAULT 0");
  ensureColumn("work_orders", "labor_rate_per_hour", "labor_rate_per_hour REAL");

  db.prepare(`
    CREATE TABLE IF NOT EXISTS cost_settings (
      key TEXT PRIMARY KEY,
      value REAL NOT NULL DEFAULT 0,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  const upsertCostSetting = db.prepare(`
    INSERT INTO cost_settings (key, value, updated_at)
    VALUES (?, ?, datetime('now'))
    ON CONFLICT(key) DO NOTHING
  `);
  upsertCostSetting.run("fuel_cost_per_liter_default", 1.5);
  upsertCostSetting.run("lube_cost_per_qty_default", 4.0);
  upsertCostSetting.run("labor_cost_per_hour_default", 35.0);
  upsertCostSetting.run("downtime_cost_per_hour_default", 120.0);

  // -----------------------------
  // Prepared statements (reuse)
  // -----------------------------

  const dailyHoursHasSite = hasColumn("daily_hours", "site_code");
  const breakdownsHasSite = hasColumn("breakdowns", "site_code");
  const downtimeLogsHasSite = hasColumn("breakdown_downtime_logs", "site_code");
  const dhSiteSql = dailyHoursHasSite
    ? `AND LOWER(TRIM(COALESCE(NULLIF(dh.site_code, ''), 'main'))) = ?`
    : "";
  const bdLogSiteSql = downtimeLogsHasSite
    ? `AND LOWER(TRIM(COALESCE(NULLIF(l.site_code, ''), NULLIF(b.site_code, ''), 'main'))) = ?`
    : breakdownsHasSite
      ? `AND LOWER(TRIM(COALESCE(NULLIF(b.site_code, ''), 'main'))) = ?`
      : "";
  const bdOnlySiteSql = breakdownsHasSite
    ? `AND LOWER(TRIM(COALESCE(NULLIF(b.site_code, ''), 'main'))) = ?`
    : "";

  // Per-asset production rows for the day (daily standby + master standby + km assets excluded from hour-based KPI pool)
  const getDayAssetHoursNoSite = db.prepare(`
    SELECT
      dh.asset_id,
      a.asset_code,
      a.asset_name,
      a.category,
      COALESCE(dh.is_used, 1) AS is_used,
      COALESCE(NULLIF(TRIM(dh.input_unit), ''), '') AS input_unit,
      CASE
        WHEN (
          (INSTR(LOWER(COALESCE(a.asset_name, '')), 'toyota') > 0 AND INSTR(LOWER(COALESCE(a.asset_name, '')), 'hilux') > 0)
          OR INSTR(LOWER(COALESCE(a.asset_code, '')), 'hilux') > 0
        ) THEN 'km'
        ELSE 'hours'
      END AS utilization_mode,
      COALESCE(NULLIF(a.km_per_hour_factor, 0), 10.0) AS km_per_hour_factor,
      COALESCE(dh.scheduled_hours, 0) AS scheduled_hours,
      COALESCE(dh.hours_run, 0) AS run_hours
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date = ?
      ${andDailyHoursFleetHoursOnly("dh", "a")}
  `);
  const getDayAssetHoursWithSite = dailyHoursHasSite
    ? db.prepare(`
    SELECT
      dh.asset_id,
      a.asset_code,
      a.asset_name,
      a.category,
      COALESCE(dh.is_used, 1) AS is_used,
      COALESCE(NULLIF(TRIM(dh.input_unit), ''), '') AS input_unit,
      CASE
        WHEN (
          (INSTR(LOWER(COALESCE(a.asset_name, '')), 'toyota') > 0 AND INSTR(LOWER(COALESCE(a.asset_name, '')), 'hilux') > 0)
          OR INSTR(LOWER(COALESCE(a.asset_code, '')), 'hilux') > 0
        ) THEN 'km'
        ELSE 'hours'
      END AS utilization_mode,
      COALESCE(NULLIF(a.km_per_hour_factor, 0), 10.0) AS km_per_hour_factor,
      COALESCE(dh.scheduled_hours, 0) AS scheduled_hours,
      COALESCE(dh.hours_run, 0) AS run_hours
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date = ?
      ${dhSiteSql}
      ${andDailyHoursFleetHoursOnly("dh", "a")}
  `)
    : null;
  const getActiveFleetAssets = db.prepare(`
    SELECT id AS asset_id, asset_code, asset_name, category, utilization_mode, km_per_hour_factor
    FROM assets
    WHERE active = 1
      AND is_standby = 0
  `);

  /** Assets marked not used for the day in Daily Input (daily standby). */
  const getDayDailyStandbyAssetIdsNoSite = db.prepare(`
    SELECT DISTINCT dh.asset_id
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date = ?
      AND COALESCE(dh.is_used, 1) = 0
      ${andAssetFleetHoursOnly("a")}
  `);
  const getDayDailyStandbyAssetIdsWithSite = dailyHoursHasSite
    ? db.prepare(`
    SELECT DISTINCT dh.asset_id
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date = ?
      ${dhSiteSql}
      AND COALESCE(dh.is_used, 1) = 0
      ${andAssetFleetHoursOnly("a")}
  `)
    : null;

  const getDayAssetDowntimeNoSite = db.prepare(`
    SELECT
      b.asset_id,
      COALESCE(SUM(l.hours_down), 0) AS downtime_hours
    FROM breakdown_downtime_logs l
    JOIN breakdowns b ON b.id = l.breakdown_id
    JOIN assets a ON a.id = b.asset_id
    WHERE l.log_date = ?
      ${andAssetFleetHoursOnly("a")}
    GROUP BY b.asset_id
  `);
  const getDayAssetDowntimeWithSite = bdLogSiteSql
    ? db.prepare(`
    SELECT
      b.asset_id,
      COALESCE(SUM(l.hours_down), 0) AS downtime_hours
    FROM breakdown_downtime_logs l
    JOIN breakdowns b ON b.id = l.breakdown_id
    JOIN assets a ON a.id = b.asset_id
    WHERE l.log_date = ?
      ${andAssetFleetHoursOnly("a")}
      ${bdLogSiteSql}
    GROUP BY b.asset_id
  `)
    : null;

  /** Breakdown WOs open on a KPI day (interval overlap with day) used for downtime imputation when no logs exist. */
  const getOpenBreakdownAssetIdsByDayNoSite = db.prepare(`
    SELECT DISTINCT w.asset_id
    FROM work_orders w
    JOIN assets a ON a.id = w.asset_id
    LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
    WHERE w.source = 'breakdown'
      AND DATE(COALESCE(NULLIF(TRIM(w.opened_at), ''), datetime('now'))) <= ?
      AND (
        w.closed_at IS NULL
        OR TRIM(COALESCE(w.closed_at, '')) = ''
        OR DATE(w.closed_at) >= ?
      )
      ${andAssetFleetHoursOnly("a")}
  `);
  const getOpenBreakdownAssetIdsByDayWithSite = bdOnlySiteSql
    ? db.prepare(`
    SELECT DISTINCT w.asset_id
    FROM work_orders w
    JOIN assets a ON a.id = w.asset_id
    LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
    WHERE w.source = 'breakdown'
      AND DATE(COALESCE(NULLIF(TRIM(w.opened_at), ''), datetime('now'))) <= ?
      AND (
        w.closed_at IS NULL
        OR TRIM(COALESCE(w.closed_at, '')) = ''
        OR DATE(w.closed_at) >= ?
      )
      ${andAssetFleetHoursOnly("a")}
      ${bdOnlySiteSql}
  `)
    : null;

  // A breakdown can be valid before its linked work order is created. Include the
  // incident itself so an open/down machine can never leave availability at 100%.
  const breakdownStatusSql = hasColumn("breakdowns", "status")
    ? `REPLACE(TRIM(LOWER(COALESCE(b.status, 'open'))), ' ', '_') IN ('open', 'in_progress')`
    : "1 = 1";
  const breakdownEndSql = hasColumn("breakdowns", "end_at")
    ? `(b.end_at IS NULL OR TRIM(COALESCE(b.end_at, '')) = '' OR DATE(b.end_at) >= ?)`
    : "1 = 1";
  const getOpenIncidentAssetIdsByDayNoSite = db.prepare(`
    SELECT DISTINCT b.asset_id
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    WHERE DATE(b.breakdown_date) <= ?
      AND ${breakdownEndSql}
      AND ${breakdownStatusSql}
      AND UPPER(TRIM(COALESCE(b.description, ''))) NOT LIKE 'MANAGER INSPECTION ALERT%'
      AND NOT EXISTS (
        SELECT 1 FROM work_orders wx
        WHERE wx.source = 'breakdown' AND COALESCE(wx.reference_id, -1) = b.id
          AND REPLACE(TRIM(LOWER(COALESCE(wx.status, ''))), ' ', '_') IN ('completed', 'approved', 'closed')
      )
      ${andAssetFleetHoursOnly("a")}
  `);
  const getOpenIncidentAssetIdsByDayWithSite = bdOnlySiteSql
    ? db.prepare(`
      SELECT DISTINCT b.asset_id
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      WHERE DATE(b.breakdown_date) <= ?
        AND ${breakdownEndSql}
        AND ${breakdownStatusSql}
        AND UPPER(TRIM(COALESCE(b.description, ''))) NOT LIKE 'MANAGER INSPECTION ALERT%'
        AND NOT EXISTS (
          SELECT 1 FROM work_orders wx
          WHERE wx.source = 'breakdown' AND COALESCE(wx.reference_id, -1) = b.id
            AND REPLACE(TRIM(LOWER(COALESCE(wx.status, ''))), ' ', '_') IN ('completed', 'approved', 'closed')
        )
        ${andAssetFleetHoursOnly("a")}
        ${bdOnlySiteSql}
    `)
    : null;

  const breakdownDowntimeCol = getBreakdownDowntimeColumn();
  const getDayAssetDowntimeFallbackNoSite = db.prepare(`
    SELECT
      b.asset_id,
      COALESCE(SUM(COALESCE(b.${breakdownDowntimeCol}, 0)), 0) AS downtime_hours
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    WHERE b.breakdown_date = ?
      ${andAssetFleetHoursOnly("a")}
    GROUP BY b.asset_id
  `);
  const getDayAssetDowntimeFallbackWithSite = bdOnlySiteSql
    ? db.prepare(`
    SELECT
      b.asset_id,
      COALESCE(SUM(COALESCE(b.${breakdownDowntimeCol}, 0)), 0) AS downtime_hours
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    WHERE b.breakdown_date = ?
      ${andAssetFleetHoursOnly("a")}
      ${bdOnlySiteSql}
    GROUP BY b.asset_id
  `)
    : null;

  const getDowntimeReasonsNoSite = db.prepare(`
    SELECT
      CASE
        WHEN l.notes IS NULL OR TRIM(l.notes) = '' THEN 'Unspecified'
        WHEN INSTR(l.notes, '—') > 0 THEN TRIM(SUBSTR(l.notes, INSTR(l.notes, '—') + 1))
        ELSE 'Unspecified'
      END AS reason,
      COALESCE(SUM(l.hours_down), 0) AS hours_down,
      COUNT(DISTINCT l.breakdown_id) AS incidents
    FROM breakdown_downtime_logs l
    JOIN breakdowns b ON b.id = l.breakdown_id
    JOIN assets a ON a.id = b.asset_id
    WHERE l.log_date = ?
      ${andAssetFleetHoursOnly("a")}
    GROUP BY reason
    ORDER BY hours_down DESC, incidents DESC
    LIMIT 8
  `);
  const getDowntimeReasonsWithSite = bdLogSiteSql
    ? db.prepare(`
    SELECT
      CASE
        WHEN l.notes IS NULL OR TRIM(l.notes) = '' THEN 'Unspecified'
        WHEN INSTR(l.notes, '—') > 0 THEN TRIM(SUBSTR(l.notes, INSTR(l.notes, '—') + 1))
        ELSE 'Unspecified'
      END AS reason,
      COALESCE(SUM(l.hours_down), 0) AS hours_down,
      COUNT(DISTINCT l.breakdown_id) AS incidents
    FROM breakdown_downtime_logs l
    JOIN breakdowns b ON b.id = l.breakdown_id
    JOIN assets a ON a.id = b.asset_id
    WHERE l.log_date = ?
      ${andAssetFleetHoursOnly("a")}
      ${bdLogSiteSql}
    GROUP BY reason
    ORDER BY hours_down DESC, incidents DESC
    LIMIT 8
  `)
    : null;

  const majorDowntimeNoSite = db.prepare(`
      SELECT
        a.asset_code,
        b.description,
        SUM(l.hours_down) AS downtime_hours,
        b.critical
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date = ?
        ${andAssetFleetHoursOnly("a")}
      GROUP BY l.breakdown_id
      ORDER BY downtime_hours DESC
      LIMIT 5
    `);
  const majorDowntimeWithSite = bdLogSiteSql
    ? db.prepare(`
      SELECT
        a.asset_code,
        b.description,
        SUM(l.hours_down) AS downtime_hours,
        b.critical
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date = ?
        ${andAssetFleetHoursOnly("a")}
        ${bdLogSiteSql}
      GROUP BY l.breakdown_id
      ORDER BY downtime_hours DESC
      LIMIT 5
    `)
    : null;

  /**
   * Fleet KPI for one calendar day (same rules as dashboard split KPI).
   * When includePerAsset=true, builds per_asset_kpi for that day (debug / detail).
   */
  function computeFleetKpiForDay(dayStr, scheduledFallback, opts = {}) {
    const includePerAsset = Boolean(opts.includePerAsset);
    const siteCode = String(opts.siteCode || "main").trim().toLowerCase() || "main";
    const activeFleetIds = new Set(getActiveFleetAssets.all().map((r) => Number(r.asset_id || 0)).filter((id) => id > 0));
    const dailyStandbyIds = new Set(
      (dailyHoursHasSite && getDayDailyStandbyAssetIdsWithSite
        ? getDayDailyStandbyAssetIdsWithSite.all(dayStr, siteCode)
        : getDayDailyStandbyAssetIdsNoSite.all(dayStr)
      )
        .map((r) => Number(r.asset_id || 0))
        .filter((id) => activeFleetIds.has(id))
    );

    const assetRows = dailyHoursHasSite && getDayAssetHoursWithSite
      ? getDayAssetHoursWithSite.all(dayStr, siteCode)
      : getDayAssetHoursNoSite.all(dayStr);

    const logDowntimeRows =
      getDayAssetDowntimeWithSite && bdLogSiteSql
        ? getDayAssetDowntimeWithSite.all(dayStr, siteCode)
        : getDayAssetDowntimeNoSite.all(dayStr);
    const fallbackDowntimeRows =
      getDayAssetDowntimeFallbackWithSite && bdOnlySiteSql
        ? getDayAssetDowntimeFallbackWithSite.all(dayStr, siteCode)
        : getDayAssetDowntimeFallbackNoSite.all(dayStr);

    // Merge per asset: prefer explicit downtime logs for that asset/day,
    // otherwise use fallback downtime from breakdown header rows.
    // Previous global-switch behavior dropped fallback rows whenever any log existed.
    const logDowntimeByAsset = new Map(
      (logDowntimeRows || [])
        .map((r) => [Number(r.asset_id || 0), Number(r.downtime_hours || 0)])
        .filter(([assetId]) => activeFleetIds.has(assetId))
    );
    const fallbackDowntimeByAsset = new Map(
      (fallbackDowntimeRows || [])
        .map((r) => [Number(r.asset_id || 0), Number(r.downtime_hours || 0)])
        .filter(([assetId]) => activeFleetIds.has(assetId))
    );
    const downtimeByAsset = new Map();
    const downtimeAssetIds = new Set([
      ...Array.from(logDowntimeByAsset.keys()),
      ...Array.from(fallbackDowntimeByAsset.keys()),
    ]);
    for (const assetId of downtimeAssetIds) {
      const logged = Number(logDowntimeByAsset.get(assetId) || 0);
      const fallback = Number(fallbackDowntimeByAsset.get(assetId) || 0);
      downtimeByAsset.set(assetId, logged > 0 ? logged : fallback);
    }
    const openWorkOrderAssetIds = getOpenBreakdownAssetIdsByDayWithSite && bdOnlySiteSql
        ? getOpenBreakdownAssetIdsByDayWithSite.all(dayStr, dayStr, siteCode)
        : getOpenBreakdownAssetIdsByDayNoSite.all(dayStr, dayStr);
    const incidentParams = hasColumn("breakdowns", "end_at") ? [dayStr, dayStr] : [dayStr];
    if (getOpenIncidentAssetIdsByDayWithSite && bdOnlySiteSql) incidentParams.push(siteCode);
    const openIncidentAssetIds = getOpenIncidentAssetIdsByDayWithSite && bdOnlySiteSql
      ? getOpenIncidentAssetIdsByDayWithSite.all(...incidentParams)
      : getOpenIncidentAssetIdsByDayNoSite.all(...incidentParams);
    const openBreakdownAssets = new Set(
      [...openWorkOrderAssetIds, ...openIncidentAssetIds]
        .map((r) => Number(r.asset_id || 0))
        .filter((assetId) => activeFleetIds.has(assetId))
    );
    const prestartAssetIds = new Set(
      listDailyPrestarts(db, dayStr).rows
        .map((r) => Number(r.asset_id || 0))
        .filter((assetId) => activeFleetIds.has(assetId))
    );
    const eligibleAssetRows = assetRows.filter((r) => {
      const assetId = Number(r.asset_id || 0);
      return activeFleetIds.has(assetId) && !dailyStandbyIds.has(assetId);
    });
    const assetIdsInHours = new Set(eligibleAssetRows.map((r) => Number(r.asset_id || 0)));

    let scheduled_hours = 0;
    let run_hours = 0;
    let downtime_hours = 0;
    let inspection_hours = 0;
    let availability_loss_hours = 0;
    let utilization_base_hours = 0;
    const per_asset_kpi = includePerAsset ? [] : null;
    const contributingAssetIds = new Set();

    eligibleAssetRows.forEach((r) => {
      const assetId = Number(r.asset_id || 0);
      const rowScheduled = Number(r.scheduled_hours);
      const rowIsUsed = Number(r.is_used ?? 1);
      const hasDailyUsageSignal = rowIsUsed === 1;
      const hasReportedRunSignal = Math.max(0, Number(r.run_hours || 0)) > 0;
      const scheduled = Math.max(
        0,
        Number.isFinite(rowScheduled) && rowScheduled > 0
          ? rowScheduled
          : (hasDailyUsageSignal || hasReportedRunSignal ? Number(scheduledFallback || 0) : 0)
      );
      const runRaw = Math.max(0, Number(r.run_hours || 0));
      const mode = isToyotaHiluxAsset(r) ? "km" : "hours";
      const kmPerHour = Math.max(0.1, Number(r.km_per_hour_factor || 10));
      const run = mode === "km" ? (runRaw / kmPerHour) : runRaw;
      const loggedDownRaw = Math.max(0, Number(downtimeByAsset.get(assetId) || 0));
      // If run is reported, do not auto-impute full-day downtime from OPEN breakdown.
      // Imputation is only for days marked used with zero reported run and no logged downtime rows.
      const allowOpenBreakdownImpute = scheduled > 0 && hasDailyUsageSignal && !hasReportedRunSignal;
      const loggedDown = loggedDownRaw > 0
        ? loggedDownRaw
        : (allowOpenBreakdownImpute && openBreakdownAssets.has(assetId) ? scheduled : 0);
      const cappedDown = Math.min(loggedDown, scheduled);
      const inspection = prestartAssetIds.has(assetId) ? PRESTART_DEDUCTION_HOURS : 0;
      const cappedLoss = Math.min(scheduled, cappedDown + inspection);
      const contributes_to_kpi = true;
      const runEff = Math.min(run, scheduled);
      scheduled_hours += scheduled;
      run_hours += runEff;
      downtime_hours += cappedDown;
      inspection_hours += Math.max(0, cappedLoss - cappedDown);
      availability_loss_hours += cappedLoss;
      utilization_base_hours += scheduled;
      const s = Number(scheduled.toFixed(2));
      const runN = Number(runEff.toFixed(2));
      const downN = Number(cappedDown.toFixed(2));
      if (s > 0 || runN > 0 || downN > 0) contributingAssetIds.add(assetId);
      if (includePerAsset) {
        const downtime_source = loggedDownRaw > 0
          ? "logged"
          : (allowOpenBreakdownImpute && openBreakdownAssets.has(assetId) ? "open_breakdown_imputed" : "none");
        per_asset_kpi.push({
          asset_id: assetId,
          asset_code: String(r.asset_code || ""),
          asset_name: String(r.asset_name || ""),
          category: String(r.category || ""),
          utilization_mode: mode,
          contributes_to_kpi,
          km_per_hour_factor: Number(kmPerHour.toFixed(2)),
          scheduled_hours: s,
          meter_run_value: Number(runRaw.toFixed(2)),
          run_hours: runN,
          downtime_hours: downN,
          inspection_hours: Number(Math.max(0, cappedLoss - cappedDown).toFixed(2)),
          available_hours: Number(Math.max(0, scheduled - cappedLoss).toFixed(2)),
          debug: {
            has_daily_row: true,
            is_used: rowIsUsed === 1,
            used_schedule_fallback: !(Number.isFinite(rowScheduled) && rowScheduled > 0) && scheduled > 0,
            downtime_source,
          },
        });
      }
    });

    // Do not force-add all active assets without daily rows.
    // KPI should primarily reflect reported data for the selected period/site.
    const missingFleetIds = [];
    const downtimeOnlyIds = Array.from(downtimeByAsset.keys())
      .map((id) => Number(id || 0))
      .filter((id) => id > 0 && !assetIdsInHours.has(id));
    const openBreakdownOnlyIds = Array.from(openBreakdownAssets).filter(
      (id) => id > 0 && !assetIdsInHours.has(id)
    );
    const includeIds = Array.from(new Set([...downtimeOnlyIds, ...openBreakdownOnlyIds, ...missingFleetIds]));
    if (includeIds.length) {
      const uniq = includeIds.slice(0, 500);
      const placeholders = uniq.map(() => "?").join(",");
      const extraAssets = db.prepare(`
        SELECT id AS asset_id, asset_code, asset_name, category, utilization_mode, km_per_hour_factor
        FROM assets
        WHERE id IN (${placeholders})
      `).all(...uniq);

      for (const a of extraAssets) {
        const assetId = Number(a.asset_id || 0);
        const modeExtra = isToyotaHiluxAsset(a)
          ? "km"
          : String(a.utilization_mode || "").trim().toLowerCase() === "km"
            ? "km"
            : "hours";
        if (modeExtra === "km") continue;

        const scheduled = Math.max(0, Number(scheduledFallback || 0));
        const loggedDownRaw = Math.max(0, Number(downtimeByAsset.get(assetId) || 0));
        const loggedDown = loggedDownRaw > 0
          ? loggedDownRaw
          : (openBreakdownAssets.has(assetId) ? scheduled : 0);
        const cappedDown = Math.min(loggedDown, scheduled);
        const inspection = prestartAssetIds.has(assetId) ? PRESTART_DEDUCTION_HOURS : 0;
        const cappedLoss = Math.min(scheduled, cappedDown + inspection);
        const cat = String(a.category || "");
        const mode = isToyotaHiluxAsset(a) ? "km" : "hours";
        const kmPerHour = Math.max(0.1, Number(a.km_per_hour_factor || 10));
        const runRaw = 0;
        const run = mode === "km" ? (runRaw / kmPerHour) : runRaw;
        const runEff = Math.min(run, scheduled);

        scheduled_hours += scheduled;
        run_hours += runEff;
        downtime_hours += cappedDown;
        inspection_hours += Math.max(0, cappedLoss - cappedDown);
        availability_loss_hours += cappedLoss;
        utilization_base_hours += scheduled;

        const s = Number(scheduled.toFixed(2));
        const runN = Number(runEff.toFixed(2));
        const downN = Number(cappedDown.toFixed(2));
        if (s > 0 || runN > 0 || downN > 0) contributingAssetIds.add(assetId);
        if (includePerAsset) {
          const downtime_source = loggedDownRaw > 0
            ? "logged"
            : (openBreakdownAssets.has(assetId) ? "open_breakdown_imputed" : "none");
          per_asset_kpi.push({
            asset_id: assetId,
            asset_code: String(a.asset_code || ""),
            asset_name: String(a.asset_name || ""),
            category: cat,
            utilization_mode: mode,
            contributes_to_kpi: true,
            km_per_hour_factor: Number(kmPerHour.toFixed(2)),
            scheduled_hours: s,
            meter_run_value: Number(runRaw.toFixed(2)),
            run_hours: runN,
            downtime_hours: downN,
            inspection_hours: Number(Math.max(0, cappedLoss - cappedDown).toFixed(2)),
            available_hours: Number(Math.max(0, scheduled - cappedLoss).toFixed(2)),
            debug: {
              has_daily_row: false,
              is_used: false,
              used_schedule_fallback: scheduled > 0,
              downtime_source,
            },
          });
        }
      }
    }

    return {
      scheduled_hours,
      run_hours,
      downtime_hours,
      inspection_hours,
      availability_loss_hours,
      utilization_base_hours,
      contributingAssetIds,
      per_asset_kpi: includePerAsset ? per_asset_kpi : [],
    };
  }

  function buildAssetKpiRange(start, end, scheduledFallback, siteCode, assetCodeFilter) {
    const allowSet = Array.isArray(assetCodeFilter) && assetCodeFilter.length
      ? new Set(assetCodeFilter.map((c) => String(c || "").trim().toUpperCase()).filter(Boolean))
      : null;
    const assetMap = new Map();
    const daily_series = [];
    let daysInRange = 0;
    eachDateInclusiveYMD(start, end, (dayStr) => {
      daysInRange += 1;
      const k = computeFleetKpiForDay(dayStr, scheduledFallback, { includePerAsset: true, siteCode });
      let dayScheduled = Number(k.scheduled_hours || 0);
      let dayRun = Number(k.run_hours || 0);
      let dayDown = Number(k.downtime_hours || 0);
      // When filtering to selected equipment, recompute the day's fleet totals
      // from just the included assets so the trend/fleet reflect the selection.
      if (allowSet) { dayScheduled = 0; dayRun = 0; dayDown = 0; }
      for (const row of k.per_asset_kpi || []) {
        const id = Number(row.asset_id || 0);
        if (!id) continue;
        const code = String(row.asset_code || "").trim().toUpperCase();
        if (allowSet && !allowSet.has(code)) continue;
        if (allowSet) {
          dayScheduled += Number(row.scheduled_hours || 0);
          dayRun += Number(row.run_hours || 0);
          dayDown += Number(row.downtime_hours || 0);
        }
        if (!assetMap.has(id)) {
          assetMap.set(id, {
            asset_id: id,
            asset_code: String(row.asset_code || ""),
            asset_name: String(row.asset_name || ""),
            category: String(row.category || ""),
            utilization_mode: String(row.utilization_mode || "hours"),
            scheduled_hours: 0,
            run_hours: 0,
            downtime_hours: 0,
            available_hours: 0,
            days_with_data: 0,
            daily_points: [],
            debug_counts: {
              reported_days: 0,
              no_daily_row_days: 0,
              used_flag_days: 0,
              standby_flag_days: 0,
              fallback_schedule_days: 0,
              logged_downtime_days: 0,
              imputed_open_breakdown_days: 0,
            },
          });
        }
        const a = assetMap.get(id);
        const s = Number(row.scheduled_hours || 0);
        const r = Number(row.run_hours || 0);
        const d = Number(row.downtime_hours || 0);
        const v = Number(row.available_hours || 0);
        a.scheduled_hours += s;
        a.run_hours += r;
        a.downtime_hours += d;
        a.available_hours += v;
        a.daily_points.push({
          date: dayStr,
          scheduled_hours: Number(s.toFixed(2)),
          available_hours: Number(v.toFixed(2)),
          run_hours: Number(r.toFixed(2)),
          downtime_hours: Number(d.toFixed(2)),
        });
        const dbg = row.debug || {};
        if (dbg.has_daily_row) a.debug_counts.reported_days += 1;
        else a.debug_counts.no_daily_row_days += 1;
        if (dbg.is_used === true) a.debug_counts.used_flag_days += 1;
        if (dbg.has_daily_row && dbg.is_used === false) a.debug_counts.standby_flag_days += 1;
        if (dbg.used_schedule_fallback) a.debug_counts.fallback_schedule_days += 1;
        if (dbg.downtime_source === "logged") a.debug_counts.logged_downtime_days += 1;
        if (dbg.downtime_source === "open_breakdown_imputed") a.debug_counts.imputed_open_breakdown_days += 1;
        if (s > 0 || r > 0 || d > 0) a.days_with_data += 1;
      }
      const dayAvail = Math.max(0, dayScheduled - dayDown);
      daily_series.push({
        date: dayStr,
        scheduled_hours: Number(dayScheduled.toFixed(2)),
        available_hours: Number(dayAvail.toFixed(2)),
        run_hours: Number(dayRun.toFixed(2)),
        downtime_hours: Number(dayDown.toFixed(2)),
      });
    });

    const pct = (num, den) =>
      den > 0 && Number.isFinite(num) ? Number(((num / den) * 100).toFixed(1)) : null;

    const by_asset = Array.from(assetMap.values()).map((a) => {
      const sched = a.scheduled_hours;
      const avail = a.available_hours;
      const run = a.run_hours;
      return {
        asset_id: a.asset_id,
        asset_code: a.asset_code,
        asset_name: a.asset_name,
        category: a.category,
        utilization_mode: a.utilization_mode,
        days_with_data: a.days_with_data,
        days_in_range: daysInRange,
        scheduled_hours: Number(sched.toFixed(2)),
        run_hours: Number(run.toFixed(2)),
        downtime_hours: Number(a.downtime_hours.toFixed(2)),
        available_hours: Number(avail.toFixed(2)),
        availability_pct: pct(avail, sched),
        utilization_pct: pct(run, sched),
        daily_points: a.daily_points,
        debug: {
          reported_days: Number(a.debug_counts.reported_days || 0),
          no_daily_row_days: Number(a.debug_counts.no_daily_row_days || 0),
          used_flag_days: Number(a.debug_counts.used_flag_days || 0),
          standby_flag_days: Number(a.debug_counts.standby_flag_days || 0),
          fallback_schedule_days: Number(a.debug_counts.fallback_schedule_days || 0),
          logged_downtime_days: Number(a.debug_counts.logged_downtime_days || 0),
          imputed_open_breakdown_days: Number(a.debug_counts.imputed_open_breakdown_days || 0),
        },
      };
    });

    by_asset.sort((x, y) => {
      if (x.utilization_pct == null && y.utilization_pct == null) {
        return String(x.asset_code || "").localeCompare(String(y.asset_code || ""));
      }
      if (x.utilization_pct == null) return 1;
      if (y.utilization_pct == null) return -1;
      return y.utilization_pct - x.utilization_pct;
    });

    const govSelect = [];
    if (hasColumn("assets", "department_code")) govSelect.push("department_code");
    if (hasColumn("assets", "cost_center_code")) govSelect.push("cost_center_code");
    if (hasColumn("assets", "data_owner_username")) govSelect.push("data_owner_username");
    if (govSelect.length) {
      const ids = by_asset.map((a) => Number(a.asset_id || 0)).filter((id) => id > 0);
      if (ids.length) {
        const ph = ids.map(() => "?").join(",");
        const rows = db
          .prepare(
            `SELECT id AS asset_id, ${govSelect.join(", ")} FROM assets WHERE id IN (${ph})`
          )
          .all(...ids);
        const gmap = new Map(rows.map((r) => [Number(r.asset_id), r]));
        for (const a of by_asset) {
          const g = gmap.get(Number(a.asset_id));
          if (!g) continue;
          if (govSelect.includes("department_code")) a.department_code = g.department_code ?? null;
          if (govSelect.includes("cost_center_code")) a.cost_center_code = g.cost_center_code ?? null;
          if (govSelect.includes("data_owner_username")) a.data_owner_username = g.data_owner_username ?? null;
        }
      }
    }

    const catMap = new Map();
    for (const a of by_asset) {
      const catKey = String(a.category || "").trim() || "Uncategorized";
      if (!catMap.has(catKey)) {
        catMap.set(catKey, {
          category: catKey,
          scheduled_hours: 0,
          run_hours: 0,
          downtime_hours: 0,
          available_hours: 0,
          asset_ids: new Set(),
        });
      }
      const c = catMap.get(catKey);
      c.scheduled_hours += a.scheduled_hours;
      c.run_hours += a.run_hours;
      c.downtime_hours += a.downtime_hours;
      c.available_hours += a.available_hours;
      c.asset_ids.add(a.asset_id);
    }

    const by_category = Array.from(catMap.values()).map((c) => {
      const sched = c.scheduled_hours;
      const avail = c.available_hours;
      const run = c.run_hours;
      return {
        category: c.category,
        asset_count: c.asset_ids.size,
        scheduled_hours: Number(sched.toFixed(2)),
        run_hours: Number(run.toFixed(2)),
        downtime_hours: Number(c.downtime_hours.toFixed(2)),
        available_hours: Number(avail.toFixed(2)),
        availability_pct: pct(avail, sched),
        utilization_pct: pct(run, sched),
      };
    });

    by_category.sort((x, y) => {
      if (x.utilization_pct == null && y.utilization_pct == null) {
        return String(x.category || "").localeCompare(String(y.category || ""));
      }
      if (x.utilization_pct == null) return 1;
      if (y.utilization_pct == null) return -1;
      return y.utilization_pct - x.utilization_pct;
    });

    const fleet_sched = by_asset.reduce((s, r) => s + r.scheduled_hours, 0);
    const fleet_avail = by_asset.reduce((s, r) => s + r.available_hours, 0);
    const fleet_run = by_asset.reduce((s, r) => s + r.run_hours, 0);
    const fleet_down = by_asset.reduce((s, r) => s + r.downtime_hours, 0);

    return {
      ok: true,
      range: { start, end },
      scheduled_fallback: scheduledFallback,
      days_in_range: daysInRange,
      definitions: {
        availability_pct: "(sum of available hours) / (sum of scheduled hours) × 100; available = scheduled − downtime (capped per day).",
        utilization_pct: "(sum of effective run hours) / (sum of scheduled hours) × 100; same denominator as the plan (matches main dashboard MTD utilization).",
      },
      fleet: {
        scheduled_hours: Number(fleet_sched.toFixed(2)),
        available_hours: Number(fleet_avail.toFixed(2)),
        run_hours: Number(fleet_run.toFixed(2)),
        downtime_hours: Number(fleet_down.toFixed(2)),
        availability_pct: pct(fleet_avail, fleet_sched),
        utilization_pct: pct(fleet_run, fleet_sched),
      },
      by_category,
      by_asset,
      daily_series,
    };
  }

  // Maintenance reliability uses this exact daily calculation for its selected
  // equipment, so the downtime and operating-hour bases cannot diverge from
  // the Asset KPI report.
  registerAssetKpiRangeBuilder(buildAssetKpiRange);

  function fuelPreviousPeriodRange(start, end) {
    const [sy, sm, sd] = start.split("-").map(Number);
    const [ey, em, ed] = end.split("-").map(Number);
    const startTs = Date.UTC(sy, sm - 1, sd);
    const endTs = Date.UTC(ey, em - 1, ed);
    const days = Math.max(1, Math.round((endTs - startTs) / 86400000) + 1);
    const prevEndTs = startTs - 86400000;
    const prevStartTs = prevEndTs - (days - 1) * 86400000;
    const fmt = (ts) => {
      const d = new Date(ts);
      return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
    };
    return {
      start: fmt(prevStartTs),
      end: fmt(prevEndTs),
      days,
    };
  }

  function buildFuelDailyUsageSeries(start, end, modeFilter, assetFilter) {
    const metricExpr = sqlFuelMetricModeExpr("a");
    const where = [
      "fl.log_date BETWEEN ? AND ?",
      sqlIncludeArchivedHireAssets("a"),
    ];
    const params = [start, end];
    if (assetFilter) {
      where.push("LOWER(TRIM(a.asset_code)) = LOWER(TRIM(?))");
      params.push(assetFilter);
    }
    if (modeFilter === "km") {
      where.push(`(${metricExpr}) = 'km'`);
    } else if (modeFilter === "hours") {
      where.push(`(${metricExpr}) = 'hours'`);
    }
    const rows = db.prepare(`
      SELECT fl.log_date AS log_date, COALESCE(SUM(fl.liters), 0) AS liters
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE ${where.join(" AND ")}
      GROUP BY fl.log_date
      ORDER BY fl.log_date ASC
    `).all(...params);
    const byDate = new Map((rows || []).map((r) => [String(r.log_date || ""), Number(r.liters || 0)]));
    const series = [];
    const [sy, sm, sd] = start.split("-").map(Number);
    const [ey, em, ed] = end.split("-").map(Number);
    let ts = Date.UTC(sy, sm - 1, sd);
    const endTs = Date.UTC(ey, em - 1, ed);
    let dayIndex = 0;
    while (ts <= endTs) {
      const d = new Date(ts);
      const date = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
      const liters = byDate.get(date) || 0;
      series.push({ day_index: dayIndex, date, liters: Number(liters.toFixed(2)) });
      dayIndex += 1;
      ts += 86400000;
    }
    const total_liters = Number(series.reduce((s, r) => s + Number(r.liters || 0), 0).toFixed(2));
    return { start, end, total_liters, series };
  }

  function scenarioHoursFromRequest(value, fallback) {
    const parsed = Number(value ?? fallback);
    return Number.isFinite(parsed) && parsed >= 0.25 && parsed <= 24 ? parsed : fallback;
  }

  function buildShiftScenarioData(start, end, baseHours, scenarioHours, assetCode, siteCode) {
    const metricExpr = sqlFuelMetricModeExpr("a");
    const assets = db.prepare(`
      SELECT
        a.id AS asset_id,
        a.asset_code,
        a.asset_name,
        COALESCE(a.category, '') AS category,
        ${metricExpr} AS metric_mode
      FROM assets a
      WHERE ${sqlIncludeArchivedHireAssets("a")}
      ORDER BY a.asset_code ASC
    `).all().filter((row) => !assetCode || String(row.asset_code || "").trim().toLowerCase() === assetCode.toLowerCase());

    const byAsset = new Map(assets.map((asset) => [Number(asset.asset_id), {
      ...asset,
      active_days: 0,
      run_hours: 0,
      fuel_liters: 0,
      fuel_cost: 0,
      nonfuel_cost: 0,
      lube_cost: 0,
      parts_cost: 0,
      labor_cost: 0,
      mechanic_labor_cost: 0,
      downtime_cost: 0,
    }]));
    if (!byAsset.size) {
      return buildShiftScenario([], { baseHours, scenarioHours });
    }

    const settings = new Map(db.prepare(`
      SELECT key, value
      FROM cost_settings
      WHERE key IN ('fuel_cost_per_liter_default', 'lube_cost_per_qty_default', 'labor_cost_per_hour_default', 'downtime_cost_per_hour_default')
    `).all().map((row) => [String(row.key || ""), Number(row.value || 0)]));
    const fuelDefault = Number.isFinite(settings.get("fuel_cost_per_liter_default")) ? settings.get("fuel_cost_per_liter_default") : 1.5;
    const lubeDefault = Number.isFinite(settings.get("lube_cost_per_qty_default")) ? settings.get("lube_cost_per_qty_default") : 4.0;
    const laborDefault = Number.isFinite(settings.get("labor_cost_per_hour_default")) ? settings.get("labor_cost_per_hour_default") : 35.0;
    const downtimeDefault = Number.isFinite(settings.get("downtime_cost_per_hour_default")) ? settings.get("downtime_cost_per_hour_default") : 120.0;

    const dailyParams = [start, end];
    const dailySiteFilter = dailyHoursHasSite
      ? "AND LOWER(TRIM(COALESCE(NULLIF(dh.site_code, ''), 'main'))) = ?"
      : "";
    if (dailyHoursHasSite) dailyParams.push(siteCode);
    const dailyRows = db.prepare(`
      SELECT
        dh.asset_id,
        COUNT(DISTINCT dh.work_date) AS active_days,
        COALESCE(SUM(dh.hours_run), 0) AS run_hours
      FROM daily_hours dh
      WHERE dh.work_date BETWEEN ? AND ?
        AND COALESCE(dh.is_used, 1) = 1
        AND LOWER(COALESCE(NULLIF(TRIM(dh.input_unit), ''), 'hours')) <> 'km'
        AND COALESCE(dh.hours_run, 0) > 0
        ${dailySiteFilter}
      GROUP BY dh.asset_id
    `).all(...dailyParams);
    for (const row of dailyRows) {
      const target = byAsset.get(Number(row.asset_id));
      if (!target) continue;
      target.active_days = Number(row.active_days || 0);
      target.run_hours = Number(row.run_hours || 0);
    }

    const fuelRows = db.prepare(`
      SELECT
        fl.asset_id,
        COALESCE(SUM(fl.liters), 0) AS fuel_liters,
        COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS fuel_cost,
        COUNT(DISTINCT fl.log_date) AS fuel_days
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date BETWEEN ? AND ?
      GROUP BY fl.asset_id
    `).all(fuelDefault, start, end);
    const fuelDays = new Map();
    for (const row of fuelRows) {
      const target = byAsset.get(Number(row.asset_id));
      if (!target) continue;
      target.fuel_liters = Number(row.fuel_liters || 0);
      target.fuel_cost = Number(row.fuel_cost || 0);
      fuelDays.set(Number(row.asset_id), Number(row.fuel_days || 0));
    }

    const fuelLogsInRange = db.prepare(`
      SELECT
        id,
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
    `);
    const fuelLogBeforeRange = db.prepare(`
      SELECT
        id,
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
    `);
    for (const asset of byAsset.values()) {
      if (String(asset.metric_mode || "hours").toLowerCase() === "km") continue;
      if (asset.run_hours <= 0 && asset.fuel_liters > 0) {
        const run = getRunFromFuelRows(
          fuelLogsInRange.all(asset.asset_id, start, end),
          fuelLogBeforeRange.get(asset.asset_id, start),
          "hours",
        );
        asset.run_hours = Number(run?.hours_run || 0);
      }
      if (asset.active_days <= 0 && asset.run_hours > 0) {
        asset.active_days = Number(fuelDays.get(Number(asset.asset_id)) || 0);
      }
    }

    const addCost = (rows, key) => {
      for (const row of rows) {
        const target = byAsset.get(Number(row.asset_id));
        if (!target) continue;
        target[key] += Number(row[key] || 0);
      }
    };
    addCost(db.prepare(`
      SELECT ol.asset_id, COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS lube_cost
      FROM oil_logs ol
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY ol.asset_id
    `).all(lubeDefault, start, end), "lube_cost");

    const stockColumns = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const stockDateExpr = stockColumns.some((column) => String(column.name) === "created_at")
      ? "DATE(sm.created_at)"
      : "DATE(sm.movement_date)";
    addCost(db.prepare(`
      SELECT w.asset_id, COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS parts_cost
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
      WHERE sm.movement_type = 'out'
        AND ${stockDateExpr} BETWEEN ? AND ?
      GROUP BY w.asset_id
    `).all(start, end), "parts_cost");
    addCost(db.prepare(`
      SELECT w.asset_id, COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      WHERE DATE(COALESCE(w.completed_at, w.closed_at)) BETWEEN ? AND ?
        AND LOWER(REPLACE(TRIM(COALESCE(w.status, '')), ' ', '_')) IN ('completed', 'approved', 'closed')
      GROUP BY w.asset_id
    `).all(laborDefault, start, end), "labor_cost");
    const mechanicsTableExists = Boolean(db.prepare(`
      SELECT 1 AS ok FROM sqlite_master WHERE type = 'table' AND name = 'mechanic_labor_entries' LIMIT 1
    `).get()?.ok);
    if (mechanicsTableExists) {
      const mechanicRows = db.prepare(`
        SELECT
          a.id AS asset_id,
          COALESCE(SUM(
            COALESCE(m.hours, 0) * COALESCE(NULLIF(m.labor_rate_per_hour, 0), ?)
          ), 0) AS mechanic_labor_cost
        FROM mechanic_labor_entries m
        JOIN assets a ON UPPER(TRIM(a.asset_code)) = UPPER(TRIM(m.asset_code))
        WHERE m.work_date BETWEEN ? AND ?
          AND LOWER(TRIM(COALESCE(m.site_code, 'main'))) = ?
        GROUP BY a.id
      `).all(laborDefault, start, end, siteCode);
      for (const row of mechanicRows) {
        const target = byAsset.get(Number(row.asset_id));
        if (!target) continue;
        target.mechanic_labor_cost = Number(row.mechanic_labor_cost || 0);
        // The time sheet is the actual labour record; use it in preference to
        // a work-order estimate for the same asset and period.
        if (target.mechanic_labor_cost > 0) target.labor_cost = target.mechanic_labor_cost;
      }
    }
    addCost(db.prepare(`
      SELECT b.asset_id, COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS downtime_cost
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date BETWEEN ? AND ?
      GROUP BY b.asset_id
    `).all(downtimeDefault, start, end), "downtime_cost");

    for (const asset of byAsset.values()) {
      asset.nonfuel_cost = asset.lube_cost + asset.parts_cost + asset.labor_cost + asset.downtime_cost;
    }
    return buildShiftScenario(Array.from(byAsset.values()), { baseHours, scenarioHours });
  }

  function addShiftScenarioWorkbook(workbook, data, { start, end, assetCode } = {}) {
    workbook.creator = "IRONLOG";
    workbook.created = new Date();
    const baseHours = Number(data.base_hours || 11);
    const scenarioHours = Number(data.scenario_hours || 8);
    const fleet = data.fleet || {};
    const summary = createManagementSummary(workbook, {
      title: "IRONLOG Shift Cost Scenario",
      periodLabel: `${start} to ${end}${assetCode ? ` · ${assetCode}` : " · All equipment"}`,
      cards: [
        { label: `${baseHours} H FUEL`, value: Number(fleet.base_fuel_liters || 0), numFmt: '#,##0.0" L"' },
        { label: `${scenarioHours} H FUEL`, value: Number(fleet.scenario_fuel_liters || 0), numFmt: '#,##0.0" L"' },
        { label: "FUEL SAVING", value: Number(fleet.fuel_liters_saved || 0), tone: "attention", numFmt: '#,##0.0" L"' },
        { label: `${baseHours} H COST / H`, value: Number(fleet.base_cost_per_operating_hour || 0), numFmt: '$#,##0.00' },
        { label: `${scenarioHours} H COST / H`, value: Number(fleet.scenario_cost_per_operating_hour || 0), numFmt: '$#,##0.00' },
        { label: "PERIOD FUEL COST SAVING", value: Number(fleet.fuel_cost_saved || 0), tone: "attention", numFmt: '$#,##0.00' },
        { label: "RECORDED NON-FUEL COST", value: Number(fleet.nonfuel_cost || 0), numFmt: '$#,##0.00' },
        { label: "EQUIPMENT ANALYSED", value: Number(data.rows?.length || 0), numFmt: '#,##0' },
      ],
      scopeLines: [
        `The selected period contains ${Number(fleet.active_shifts || 0).toFixed(0)} active equipment shifts across ${Number(data.rows?.length || 0)} equipment units.`,
        "Fuel scales with each machine's logged L/hr. Lube, parts, labour and downtime costs are held constant per active equipment shift.",
        "This is a planning comparison, not a production forecast. Assets with incomplete fuel/hour data or kilometre-based utilisation are listed separately.",
      ],
    });
    summary.ws.getColumn(1).width = 20;

    const detail = workbook.addWorksheet("Equipment comparison");
    detail.columns = [
      { header: "Asset", key: "asset_code", width: 14 },
      { header: "Equipment", key: "asset_name", width: 26 },
      { header: "Category", key: "category", width: 18 },
      { header: "Active shifts", key: "active_days", width: 13 },
      { header: "Logged run h", key: "logged_run_hours", width: 14 },
      { header: "Actual L/hr", key: "actual_liters_per_hour", width: 13 },
      { header: "Non-fuel cost / shift", key: "nonfuel_cost_per_shift", width: 19 },
      { header: `${baseHours}h fuel L / shift`, key: "base_fuel_liters_per_shift", width: 18 },
      { header: `${scenarioHours}h fuel L / shift`, key: "scenario_fuel_liters_per_shift", width: 19 },
      { header: "Fuel L saved / shift", key: "fuel_liters_saved_per_shift", width: 18 },
      { header: `${baseHours}h total cost / shift`, key: "base_total_cost_per_shift", width: 21 },
      { header: `${scenarioHours}h total cost / shift`, key: "scenario_total_cost_per_shift", width: 22 },
      { header: `${baseHours}h cost / h`, key: "base_cost_per_operating_hour", width: 16 },
      { header: `${scenarioHours}h cost / h`, key: "scenario_cost_per_operating_hour", width: 17 },
      { header: "Cost / h change", key: "cost_per_operating_hour_change", width: 16 },
      { header: "Period fuel saving L", key: "period_fuel_liters_saved", width: 19 },
      { header: "Period cost saving", key: "period_cost_saved", width: 19 },
    ];
    for (const row of data.rows || []) detail.addRow(row);
    detail.addRow({
      asset_code: "TOTAL",
      asset_name: `${Number(data.rows?.length || 0)} equipment units`,
      active_days: Number(fleet.active_shifts || 0),
      logged_run_hours: (data.rows || []).reduce((total, row) => total + Number(row.logged_run_hours || 0), 0),
      actual_liters_per_hour: Number(fleet.base_operating_hours || 0) > 0 ? Number(fleet.base_fuel_liters || 0) / Number(fleet.base_operating_hours || 1) : 0,
      nonfuel_cost_per_shift: Number(fleet.active_shifts || 0) > 0 ? Number(fleet.nonfuel_cost || 0) / Number(fleet.active_shifts || 1) : 0,
      base_fuel_liters_per_shift: Number(fleet.active_shifts || 0) > 0 ? Number(fleet.base_fuel_liters || 0) / Number(fleet.active_shifts || 1) : 0,
      scenario_fuel_liters_per_shift: Number(fleet.active_shifts || 0) > 0 ? Number(fleet.scenario_fuel_liters || 0) / Number(fleet.active_shifts || 1) : 0,
      fuel_liters_saved_per_shift: Number(fleet.active_shifts || 0) > 0 ? Number(fleet.fuel_liters_saved || 0) / Number(fleet.active_shifts || 1) : 0,
      base_total_cost_per_shift: Number(fleet.active_shifts || 0) > 0 ? Number(fleet.base_total_cost || 0) / Number(fleet.active_shifts || 1) : 0,
      scenario_total_cost_per_shift: Number(fleet.active_shifts || 0) > 0 ? Number(fleet.scenario_total_cost || 0) / Number(fleet.active_shifts || 1) : 0,
      base_cost_per_operating_hour: Number(fleet.base_cost_per_operating_hour || 0),
      scenario_cost_per_operating_hour: Number(fleet.scenario_cost_per_operating_hour || 0),
      cost_per_operating_hour_change: Number(fleet.cost_per_operating_hour_change || 0),
      period_fuel_liters_saved: Number(fleet.fuel_liters_saved || 0),
      period_cost_saved: Number(fleet.total_cost_saved || 0),
    });
    const styled = styleManagementDetailSheet(detail, {
      title: "Equipment shift comparison",
      subtitle: `Comparison of ${baseHours}-hour and ${scenarioHours}-hour shifts · ${start} to ${end}`,
      frozenColumns: 2,
      numberFormats: {
        D: '#,##0', E: '#,##0.0', F: '#,##0.000', G: '$#,##0.00', H: '#,##0.0', I: '#,##0.0', J: '#,##0.0',
        K: '$#,##0.00', L: '$#,##0.00', M: '$#,##0.00', N: '$#,##0.00', O: '$#,##0.00', P: '#,##0.0', Q: '$#,##0.00',
      },
    });
    detail.autoFilter = { from: { row: styled.headerRow, column: 1 }, to: { row: styled.lastRow, column: detail.columnCount } };

    const excluded = workbook.addWorksheet("Excluded equipment");
    excluded.columns = [
      { header: "Asset", key: "asset_code", width: 16 },
      { header: "Equipment", key: "asset_name", width: 34 },
      { header: "Reason excluded", key: "reason", width: 38 },
    ];
    if (data.excluded?.length) {
      data.excluded.forEach((row) => excluded.addRow(row));
    } else {
      excluded.addRow({ asset_code: "—", asset_name: "", reason: "All selected assets had a valid hours and fuel basis." });
    }
    const excludedStyled = styleManagementDetailSheet(excluded, {
      title: "Excluded equipment",
      subtitle: "Assets excluded from the shift comparison to avoid assumptions based on incomplete or kilometre data.",
      frozenColumns: 1,
    });
    excluded.autoFilter = { from: { row: excludedStyled.headerRow, column: 1 }, to: { row: excludedStyled.lastRow, column: excluded.columnCount } };
  }

  // Route groups live in routes/dashboard/. They receive the shared helpers above through ctx.
  const ctx = {
    addShiftScenarioWorkbook,
    bdLogSiteSql,
    buildAssetKpiRange,
    buildFuelDailyUsageSeries,
    buildShiftScenarioData,
    computeFleetKpiForDay,
    dailyHoursHasSite,
    eachDateInclusiveYMD,
    fuelPreviousPeriodRange,
    getDowntimeReasonsNoSite,
    getDowntimeReasonsWithSite,
    hasColumn,
    majorDowntimeNoSite,
    majorDowntimeWithSite,
    monthRangeFromYYYYMM,
    monthStartIso,
    requireRoles,
    scenarioHoursFromRequest,
    siteCodeFromReq,
    slaPriority,
    todayYYYYMMDD,
  };
  registerOverviewRoutes(app, ctx);
  registerAssetKpiRoutes(app, ctx);
  registerLubeRoutes(app, ctx);
  registerCostSettingsRoutes(app, ctx);
  registerFuelRoutes(app, ctx);
}
