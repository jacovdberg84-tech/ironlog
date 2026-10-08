// Review-only maintenance schedule builder for June.
// The voice model only selects the requested equipment and horizon. Meter
// readings, service due logic and the downloadable workbook remain entirely
// deterministic and inside Ironlog.
import ExcelJS from "exceljs";
import { getAssetHoursInfoAsOf } from "./assetMeterHours.js";
import {
  classifyServiceDue,
  groupActivePlansByAsset,
  meterUnitForAsset,
  resolveNextServiceForAssetPlans,
} from "./serviceSchedule.js";
import { createManagementSummary, styleManagementDetailSheet } from "./managementWorkbook.js";

const DEFAULT_HORIZON_DAYS = 30;
const MIN_HORIZON_DAYS = 7;
const MAX_HORIZON_DAYS = 365;
const MAX_ASSETS = 30;

function round(value, digits = 1) {
  const number = Number(value);
  if (!Number.isFinite(number)) return 0;
  const factor = 10 ** digits;
  return Math.round(number * factor) / factor;
}

function normaliseCodes(value) {
  const values = Array.isArray(value) ? value : [value];
  return [...new Set(values
    .map((code) => String(code || "").trim().toUpperCase())
    .filter(Boolean))]
    .slice(0, MAX_ASSETS);
}

function boundedHorizon(value) {
  const days = Number(value);
  if (!Number.isFinite(days)) return DEFAULT_HORIZON_DAYS;
  return Math.max(MIN_HORIZON_DAYS, Math.min(MAX_HORIZON_DAYS, Math.round(days)));
}

function addDays(ymd, days) {
  const date = new Date(`${String(ymd).slice(0, 10)}T12:00:00Z`);
  date.setUTCDate(date.getUTCDate() + Math.max(0, Math.round(Number(days) || 0)));
  return date.toISOString().slice(0, 10);
}

function isYmd(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "").trim());
}

function recentUsage(assetId, asOf, dbConn) {
  const row = dbConn.prepare(`
    SELECT
      COALESCE(SUM(day_run), 0) AS run_hours,
      COUNT(*) AS operating_days
    FROM (
      SELECT work_date, SUM(hours_run) AS day_run
      FROM daily_hours
      WHERE asset_id = ?
        AND is_used = 1
        AND hours_run > 0
        AND work_date BETWEEN date(?, '-27 days') AND ?
      GROUP BY work_date
    )
  `).get(assetId, asOf, asOf) || {};
  const runHours = Number(row.run_hours || 0);
  const operatingDays = Number(row.operating_days || 0);
  return {
    run_hours: round(runHours),
    operating_days: operatingDays,
    average_per_operating_day: operatingDays > 0 ? round(runHours / operatingDays, 2) : 0,
  };
}

function scheduleStatus(due, forecastDays, horizonDays) {
  if (due.status === "OVERDUE") return "PLAN NOW";
  if (Number.isFinite(forecastDays) && forecastDays <= horizonDays) return "PLAN IN WINDOW";
  if (due.status === "ALMOST DUE") return "WATCH CLOSELY";
  return "MONITOR";
}

function scheduleRank(row) {
  const ranks = { "PLAN NOW": 1, "PLAN IN WINDOW": 2, "WATCH CLOSELY": 3, MONITOR: 4, "NO ACTIVE PLAN": 5 };
  return ranks[row.schedule_status] || 9;
}

/**
 * Build a factual, review-only service schedule for selected equipment.
 * Forecast dates use the most recent 28-day average per operating day; they
 * are deliberately planning estimates and do not change a maintenance plan.
 */
export function buildJuneMaintenanceSchedule({ assetCodes, asOf, horizonDays, dbConn } = {}) {
  if (!dbConn?.prepare) throw new Error("A database connection is required to build June's maintenance schedule.");
  const requestedCodes = normaliseCodes(assetCodes);
  const selectedAsOf = isYmd(asOf) ? String(asOf) : new Date().toISOString().slice(0, 10);
  const selectedHorizon = boundedHorizon(horizonDays);
  if (!requestedCodes.length) {
    return {
      as_of: selectedAsOf,
      horizon_days: selectedHorizon,
      requested_asset_codes: [],
      missing_asset_codes: [],
      rows: [],
      summary: { selected_assets: 0, overdue: 0, planned_in_window: 0, no_active_plan: 0 },
    };
  }

  const placeholders = requestedCodes.map(() => "?").join(", ");
  const assets = dbConn.prepare(`
    SELECT id, asset_code, asset_name, category
    FROM assets
    WHERE UPPER(TRIM(asset_code)) IN (${placeholders})
  `).all(...requestedCodes);
  const assetsByCode = new Map(assets.map((asset) => [String(asset.asset_code || "").trim().toUpperCase(), asset]));
  const foundIds = assets.map((asset) => Number(asset.id)).filter((id) => id > 0);
  const planRows = foundIds.length
    ? dbConn.prepare(`
      SELECT mp.id, mp.asset_id, mp.service_name, mp.interval_hours, mp.last_service_hours, mp.active,
        a.asset_code, a.asset_name, a.category
      FROM maintenance_plans mp
      JOIN assets a ON a.id = mp.asset_id
      WHERE mp.active = 1 AND mp.asset_id IN (${foundIds.map(() => "?").join(", ")})
      ORDER BY a.asset_code ASC, mp.id ASC
    `).all(...foundIds)
    : [];
  const plansByAsset = groupActivePlansByAsset(planRows);
  const rows = [];

  for (const requestedCode of requestedCodes) {
    const asset = assetsByCode.get(requestedCode);
    if (!asset) continue;
    const meter = getAssetHoursInfoAsOf(asset.id, selectedAsOf, dbConn);
    const currentMeter = round(meter.hours);
    const usage = recentUsage(asset.id, selectedAsOf, dbConn);
    const plans = plansByAsset.get(Number(asset.id)) || [];
    const meterUnit = meterUnitForAsset(asset.asset_code);
    const next = resolveNextServiceForAssetPlans(plans, currentMeter, asset.asset_code);

    if (!next) {
      rows.push({
        asset_code: String(asset.asset_code || requestedCode),
        asset_name: String(asset.asset_name || ""),
        category: String(asset.category || ""),
        current_meter: currentMeter,
        meter_unit: meterUnit,
        meter_source: String(meter.source || "unknown"),
        meter_date: meter.latest_work_date || null,
        service_name: "No active service plan",
        due_meter: null,
        remaining: null,
        average_daily_run: usage.average_per_operating_day,
        forecast_due_date: null,
        forecast_days: null,
        plan_status: "NO PLAN",
        schedule_status: "NO ACTIVE PLAN",
        planning_note: "Set up an active service interval before scheduling this asset.",
      });
      continue;
    }

    const due = classifyServiceDue(next.remaining_hours, asset.asset_code, next.next_service_interval || next.interval_hours);
    const remaining = round(next.remaining_hours);
    const forecastDays = remaining <= 0
      ? 0
      : usage.average_per_operating_day > 0
        ? Math.ceil(remaining / usage.average_per_operating_day)
        : null;
    const forecastDueDate = Number.isFinite(forecastDays) ? addDays(selectedAsOf, forecastDays) : null;
    const status = scheduleStatus(due, forecastDays, selectedHorizon);
    rows.push({
      asset_code: String(asset.asset_code || requestedCode),
      asset_name: String(asset.asset_name || ""),
      category: String(asset.category || ""),
      current_meter: currentMeter,
      meter_unit: meterUnit,
      meter_source: String(meter.source || "unknown"),
      meter_date: meter.latest_work_date || null,
      service_name: String(next.service_name || "Service"),
      due_meter: round(next.next_due_hours),
      remaining,
      average_daily_run: usage.average_per_operating_day,
      forecast_due_date: forecastDueDate,
      forecast_days: forecastDays,
      plan_status: due.status,
      schedule_status: status,
      planning_note: due.status === "OVERDUE"
        ? "Overdue — confirm parts, labour and downtime window before creating work."
        : forecastDueDate
          ? `Forecast based on ${usage.operating_days} recorded operating day(s) in the last 28 days.`
          : "No recent operating pattern is available; confirm a target date manually.",
    });
  }

  rows.sort((a, b) => scheduleRank(a) - scheduleRank(b)
    || String(a.forecast_due_date || "9999-12-31").localeCompare(String(b.forecast_due_date || "9999-12-31"))
    || a.asset_code.localeCompare(b.asset_code));
  const summary = {
    selected_assets: rows.length,
    overdue: rows.filter((row) => row.plan_status === "OVERDUE").length,
    planned_in_window: rows.filter((row) => row.schedule_status === "PLAN IN WINDOW").length,
    no_active_plan: rows.filter((row) => row.schedule_status === "NO ACTIVE PLAN").length,
  };
  return {
    as_of: selectedAsOf,
    horizon_days: selectedHorizon,
    requested_asset_codes: requestedCodes,
    missing_asset_codes: requestedCodes.filter((code) => !assetsByCode.has(code)),
    rows,
    summary,
  };
}

export async function buildJuneMaintenanceScheduleWorkbook(schedule = {}) {
  const rows = Array.isArray(schedule.rows) ? schedule.rows : [];
  const summary = schedule.summary || {};
  const asOf = String(schedule.as_of || "");
  const horizon = Number(schedule.horizon_days || DEFAULT_HORIZON_DAYS);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "IRONLOG — June";
  workbook.created = new Date();

  createManagementSummary(workbook, {
    title: "IRONLOG June — Draft Maintenance Schedule",
    periodLabel: `Planning horizon: ${asOf} to ${addDays(asOf, horizon)} (${horizon} days)`,
    cards: [
      { label: "SELECTED EQUIPMENT", value: Number(summary.selected_assets || 0), numFmt: "#,##0" },
      { label: "OVERDUE", value: Number(summary.overdue || 0), numFmt: "#,##0", tone: "warning" },
      { label: "PLAN IN WINDOW", value: Number(summary.planned_in_window || 0), numFmt: "#,##0", tone: "attention" },
      { label: "NO ACTIVE PLAN", value: Number(summary.no_active_plan || 0), numFmt: "#,##0", tone: "warning" },
    ],
    scopeLines: [
      "Review-only planning draft prepared by June. It does not create work orders, service records or purchase requisitions.",
      "Forecast dates use the selected asset's latest reliable meter and its recent 28-day operating pattern. Confirm labour, parts and downtime before execution.",
      schedule.missing_asset_codes?.length ? `Asset codes not found: ${schedule.missing_asset_codes.join(", ")}.` : "All requested asset codes were found in Ironlog.",
    ],
  });

  const detail = workbook.addWorksheet("Maintenance Schedule");
  detail.columns = [
    { header: "Fleet number", key: "asset_code", width: 15 },
    { header: "Equipment", key: "asset_name", width: 30 },
    { header: "Category", key: "category", width: 18 },
    { header: "Current meter", key: "current_meter", width: 15 },
    { header: "Unit", key: "meter_unit", width: 10 },
    { header: "Meter source", key: "meter_source", width: 16 },
    { header: "Next service", key: "service_name", width: 24 },
    { header: "Due meter", key: "due_meter", width: 14 },
    { header: "Remaining", key: "remaining", width: 14 },
    { header: "Avg / operating day", key: "average_daily_run", width: 20 },
    { header: "Forecast due", key: "forecast_due_date", width: 15 },
    { header: "Service status", key: "plan_status", width: 16 },
    { header: "Schedule action", key: "schedule_status", width: 18 },
    { header: "Planning note", key: "planning_note", width: 58 },
  ];
  detail.addRows(rows.map((row) => ({
    ...row,
    current_meter: Number(row.current_meter || 0),
    due_meter: row.due_meter == null ? null : Number(row.due_meter),
    remaining: row.remaining == null ? null : Number(row.remaining),
    average_daily_run: Number(row.average_daily_run || 0),
  })));
  styleManagementDetailSheet(detail, {
    title: "Draft maintenance schedule",
    subtitle: `As at ${asOf} · ${horizon}-day planning horizon · review before creating any operational record`,
    frozenColumns: 2,
    numberFormats: {
      current_meter: "#,##0.0",
      due_meter: "#,##0.0",
      remaining: "#,##0.0",
      average_daily_run: "#,##0.00",
    },
  });
  detail.getColumn("planning_note").alignment = { vertical: "top", wrapText: true };
  detail.getColumn("schedule_status").eachCell({ includeEmpty: false }, (cell) => {
    const value = String(cell.value || "");
    if (value === "PLAN NOW" || value === "NO ACTIVE PLAN") {
      cell.font = { ...(cell.font || {}), bold: true, color: { argb: "FF9C0006" } };
    } else if (value === "PLAN IN WINDOW" || value === "WATCH CLOSELY") {
      cell.font = { ...(cell.font || {}), bold: true, color: { argb: "FF9C6500" } };
    }
  });

  return workbook.xlsx.writeBuffer();
}

