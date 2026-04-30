// IRONLOG/api/routes/reports.routes.js
import path from "node:path";
import fs from "node:fs";
import crypto from "node:crypto";
import ExcelJS from "exceljs";
import PptxGenJS from "pptxgenjs";
import nodemailer from "nodemailer";
import { db } from "../db/client.js";
import {
  buildPdfBuffer,
  tryDrawLogo,
  sectionTitle,
  kvGrid,
  table,
  ensurePageSpace,
} from "../utils/pdfGenerator.js";
import { andDailyHoursFleetHoursOnly, andAssetFleetHoursOnly } from "../utils/fleetHoursKpiScope.js";

let maintenanceMasterSchedulerStarted = false;
let reportSubscriptionsSchedulerStarted = false;

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function todayYmd() {
  return new Date().toISOString().slice(0, 10);
}

function yn(v) {
  return v ? "YES" : "NO";
}

function fmtNum(v, dp = 1) {
  if (v == null || v === "") return "";
  const n = Number(v);
  if (!Number.isFinite(n)) return String(v);
  return dp === 0 ? String(Math.round(n)) : n.toFixed(dp).replace(/\.0$/, "");
}

function compactCell(v, max = 140) {
  const s = String(v ?? "").replace(/\s+/g, " ").trim();
  if (!s) return "";
  return s.length > max ? `${s.slice(0, Math.max(1, max - 1))}...` : s;
}

function makeArtisanFormNumber() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  const stamp = `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}${p(d.getHours())}${p(d.getMinutes())}`;
  const rand = Math.floor(Math.random() * 900 + 100);
  return `AI-${stamp}-${rand}`;
}

function parseIsoDate(d) {
  if (!d) return null;
  const s = String(d || "").trim().slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : null;
}

function daysDownForBreakdown(bd, reportDate) {
  const logged = Number(bd.logged_days || 0);
  const startDate = parseIsoDate(bd.start_at) || parseIsoDate(bd.breakdown_date);
  if (!startDate) return logged > 0 ? logged : 1;

  const asOf = parseIsoDate(reportDate);
  const endDate = parseIsoDate(bd.end_at);
  let spanEnd = endDate || asOf || startDate;
  if (asOf && endDate && endDate > asOf) {
    // Daily report is "as of selected date", so never count beyond it.
    spanEnd = asOf;
  }

  const spanDays = spanEnd >= startDate ? inclusiveDaysBetween(startDate, spanEnd) : 0;
  if (logged > 0) {
    // If daily logs are sparse/missing on some days, do not under-report days down.
    return Math.max(logged, spanDays);
  }

  return spanDays || 1;
}

function daysDownForBreakdownInRange(bd, startDateInclusive, endDateInclusive) {
  const startDate = parseIsoDate(bd.start_at) || parseIsoDate(bd.breakdown_date);
  if (!startDate) return 0;
  const rangeStart = parseIsoDate(startDateInclusive);
  const rangeEnd = parseIsoDate(endDateInclusive);
  if (!rangeStart || !rangeEnd) return 0;

  const rawEnd = parseIsoDate(bd.end_at) || rangeEnd;
  const spanStart = startDate > rangeStart ? startDate : rangeStart;
  const spanEnd = rawEnd < rangeEnd ? rawEnd : rangeEnd;
  if (spanEnd < spanStart) return 0;

  const spanDays = inclusiveDaysBetween(spanStart, spanEnd);
  const loggedInRange = Number(bd.logged_days_in_range || 0);
  // Keep consistent with Daily: never undercount when logs are sparse.
  return Math.max(spanDays, loggedInRange);
}

function isMonth(m) {
  return /^\d{4}-\d{2}$/.test(String(m || "").trim());
}

function isYmd(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "").trim());
}

function monthRange(monthStr) {
  const [y, m] = String(monthStr).split("-").map((n) => Number(n));
  const start = new Date(Date.UTC(y, m - 1, 1));
  const end = new Date(Date.UTC(y, m, 0));
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

function prevMonth(monthStr) {
  const [y, m] = String(monthStr).split("-").map((n) => Number(n));
  const d = new Date(Date.UTC(y, m - 1, 1));
  d.setUTCMonth(d.getUTCMonth() - 1);
  const yy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  return `${yy}-${mm}`;
}

/** First calendar day of the month containing `endDate` (YYYY-MM-DD). */
function monthStartIso(endDate) {
  const s = String(endDate || "").trim();
  const y = Number(s.slice(0, 4));
  const m = Number(s.slice(5, 7));
  if (!Number.isFinite(y) || !Number.isFinite(m)) return s;
  return `${y}-${String(m).padStart(2, "0")}-01`;
}

/** Inclusive calendar days from start through end (both YYYY-MM-DD). */
function inclusiveDaysBetween(start, end) {
  const a = new Date(`${start}T12:00:00`);
  const b = new Date(`${end}T12:00:00`);
  return Math.round((b.getTime() - a.getTime()) / (24 * 3600 * 1000)) + 1;
}

function safeNum(v, fallback = 0) {
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

function asArray(v) {
  return Array.isArray(v) ? v : [];
}

function hasBreakdownDowntimeLogsTable() {
  return Boolean(
    db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'breakdown_downtime_logs'`).get(),
  );
}

/**
 * Downtime hours for [start,end]: prefer daily breakdown_downtime_logs (carries multi-day incidents).
 * If no log rows exist in range, fall back to summing breakdown records by breakdown_date (legacy).
 */
function getDowntimeHoursForPeriod(start, end) {
  if (hasBreakdownDowntimeLogsTable()) {
    const cnt = db.prepare(`
      SELECT COUNT(*) AS n FROM breakdown_downtime_logs WHERE log_date BETWEEN ? AND ?
    `).get(start, end);
    if (Number(cnt?.n || 0) > 0) {
      const row = db.prepare(`
        SELECT COALESCE(SUM(hours_down), 0) AS h
        FROM breakdown_downtime_logs
        WHERE log_date BETWEEN ? AND ?
      `).get(start, end);
      return Number(row?.h || 0);
    }
  }
  const dtCol = getBreakdownDowntimeColumn();
  const dtRow = db.prepare(`
    SELECT IFNULL(SUM(${dtCol}), 0) AS downtime_hours
    FROM breakdowns
    WHERE breakdown_date BETWEEN ? AND ?
  `).get(start, end);
  return Number(dtRow.downtime_hours || 0);
}

/** True when MTD downtime is taken from daily per-incident logs (multi-day carry-over). */
function downtimeMtdUsesDailyLogs(start, end) {
  if (!hasBreakdownDowntimeLogsTable()) return false;
  const cnt = db.prepare(`
    SELECT COUNT(*) AS n FROM breakdown_downtime_logs WHERE log_date BETWEEN ? AND ?
  `).get(start, end);
  return Number(cnt?.n || 0) > 0;
}

/** Prefer summed incident downtime when present (matches dashboard reliability). */
function getBreakdownDowntimeColumn() {
  const rows = db.prepare(`PRAGMA table_info(breakdowns)`).all();
  const names = new Set(rows.map((r) => String(r.name || "")));
  if (names.has("downtime_total_hours")) return "downtime_total_hours";
  if (names.has("downtime_hours")) return "downtime_hours";
  return "downtime_hours";
}

/** Same hour-meter logic as maintenance routes (forecast / PM compliance). */
function assetCurrentHoursForGm(assetId) {
  const id = Number(assetId);
  if (!Number.isFinite(id) || id <= 0) return 0;

  const fromAssetHours = db.prepare(`
    SELECT total_hours FROM asset_hours WHERE asset_id = ?
  `).get(id);
  const assetHours = fromAssetHours?.total_hours == null ? null : Number(fromAssetHours.total_hours);

  const latestMeter = db.prepare(`
    SELECT MAX(closing_hours) AS max_closing
    FROM daily_hours
    WHERE asset_id = ? AND closing_hours IS NOT NULL
  `).get(id);
  const maxClosing = latestMeter?.max_closing == null ? null : Number(latestMeter.max_closing);

  if (assetHours != null && maxClosing != null) {
    if (Math.abs(assetHours - maxClosing) > 5000) return maxClosing;
    return Math.max(assetHours, maxClosing);
  }
  if (maxClosing != null) return maxClosing;
  if (assetHours != null) return assetHours;

  const fromDailyHours = db.prepare(`
    SELECT COALESCE(SUM(hours_run), 0) AS total_hours
    FROM daily_hours
    WHERE asset_id = ? AND is_used = 1 AND hours_run > 0
  `).get(id);
  return Number(fromDailyHours?.total_hours || 0);
}

/**
 * Failure count for a window:
 * - report_date: breakdowns whose breakdown_date falls in [start,end] (legacy / dashboard-style).
 * - activity_or_report: distinct incidents with breakdown_date in range OR any downtime log in range
 *   (carries failures that started earlier but still accrued downtime this month).
 */
function countFailuresInPeriod(start, end, mode) {
  const m = String(mode || "report_date").trim();
  if (m === "activity_or_report" && hasBreakdownDowntimeLogsTable()) {
    const row = db.prepare(`
      SELECT COUNT(DISTINCT b.id) AS n
      FROM breakdowns b
      WHERE b.breakdown_date BETWEEN ? AND ?
         OR EXISTS (
           SELECT 1 FROM breakdown_downtime_logs l
           WHERE l.breakdown_id = b.id AND l.log_date BETWEEN ? AND ?
         )
    `).get(start, end, start, end);
    return Number(row?.n || 0);
  }
  const failuresRow = db.prepare(`
    SELECT COUNT(*) AS n FROM breakdowns WHERE breakdown_date BETWEEN ? AND ?
  `).get(start, end);
  return Number(failuresRow?.n || 0);
}

function reliabilityMetricsForRange(start, end, opts = {}) {
  const failure_mode = opts.failuresInPeriodMode || "report_date";
  const failure_count = countFailuresInPeriod(start, end, failure_mode);
  const effective_failure_mode =
    failure_mode === "activity_or_report" && !hasBreakdownDowntimeLogsTable()
      ? "report_date"
      : failure_mode;

  const runRow = db.prepare(`
    SELECT COALESCE(SUM(hours_run), 0) AS run_hours
    FROM daily_hours
    WHERE work_date BETWEEN ? AND ? AND is_used = 1 AND hours_run > 0
  `).get(start, end);
  const operating_hours = Number(runRow?.run_hours || 0);

  let downtime_hours;
  if (opts.downtimeHoursOverride != null && Number.isFinite(Number(opts.downtimeHoursOverride))) {
    downtime_hours = Number(opts.downtimeHoursOverride);
  } else {
    const dtCol = getBreakdownDowntimeColumn();
    const dtRow = db.prepare(`
      SELECT COALESCE(SUM(${dtCol}), 0) AS dt
      FROM breakdowns WHERE breakdown_date BETWEEN ? AND ?
    `).get(start, end);
    downtime_hours = Number(dtRow?.dt || 0);
  }

  const mtbf_hours = failure_count > 0 ? operating_hours / failure_count : null;
  const mttr_hours = failure_count > 0 ? downtime_hours / failure_count : null;

  return {
    failure_count,
    failures_in_period_mode: effective_failure_mode,
    operating_hours: Number(operating_hours.toFixed(2)),
    downtime_hours: Number(downtime_hours.toFixed(2)),
    mtbf_hours: mtbf_hours == null ? null : Number(mtbf_hours.toFixed(2)),
    mttr_hours: mttr_hours == null ? null : Number(mttr_hours.toFixed(2)),
  };
}

function kpiDaily(date, scheduled) {
  const usedRow = db.prepare(`
    SELECT COUNT(DISTINCT dh.asset_id) AS used_assets
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date = ?
      ${andDailyHoursFleetHoursOnly("dh", "a")}
  `).get(date);

  const used_assets = Number(usedRow.used_assets || 0);

  const runRow = db.prepare(`
    SELECT IFNULL(SUM(dh.hours_run), 0) AS run_hours
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date = ?
      ${andDailyHoursFleetHoursOnly("dh", "a")}
  `).get(date);

  const run_hours = Number(runRow.run_hours || 0);
  const utilBaseRow = db.prepare(`
    SELECT IFNULL(SUM(
      CASE
        WHEN COALESCE(dh.scheduled_hours, 0) > 0 THEN dh.scheduled_hours
        ELSE ?
      END
    ), 0) AS utilization_base_hours
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date = ?
      ${andDailyHoursFleetHoursOnly("dh", "a")}
  `).get(Number(scheduled || 0), date);
  const utilization_base_hours = Number(utilBaseRow?.utilization_base_hours || 0);
  const available_hours = utilization_base_hours;

  const dtLogsRow = db.prepare(`
    SELECT IFNULL(SUM(l.hours_down), 0) AS downtime_hours
    FROM breakdown_downtime_logs l
    JOIN breakdowns b ON b.id = l.breakdown_id
    JOIN assets a ON a.id = b.asset_id
    JOIN daily_hours dh ON dh.asset_id = b.asset_id AND dh.work_date = ?
    WHERE l.log_date = ?
      ${andDailyHoursFleetHoursOnly("dh", "a")}
  `).get(date, date);
  let downtime_hours = Number(dtLogsRow?.downtime_hours || 0);
  const openNoLogRow = db.prepare(`
    SELECT IFNULL(SUM(
      CASE
        WHEN COALESCE(dh.scheduled_hours, 0) > 0 THEN dh.scheduled_hours
        ELSE ?
      END
    ), 0) AS assumed_down_hours
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    JOIN daily_hours dh ON dh.asset_id = b.asset_id AND dh.work_date = ?
    WHERE b.status = 'OPEN'
      AND b.breakdown_date <= ?
      AND NOT EXISTS (
        SELECT 1
        FROM breakdown_downtime_logs l
        WHERE l.breakdown_id = b.id
          AND l.log_date = ?
      )
      ${andDailyHoursFleetHoursOnly("dh", "a")}
  `).get(Number(scheduled || 0), date, date, date);
  downtime_hours += Number(openNoLogRow?.assumed_down_hours || 0);

  const availability = available_hours > 0 ? ((available_hours - downtime_hours) / available_hours) * 100 : null;
  const utilization = utilization_base_hours > 0 ? (run_hours / utilization_base_hours) * 100 : null;

  return {
    used_assets,
    available_hours,
    utilization_base_hours,
    run_hours,
    downtime_hours,
    availability: availability == null ? null : Number(availability.toFixed(2)),
    utilization: utilization == null ? null : Number(utilization.toFixed(2)),
  };
}

function kpiRange(start, end, scheduled, opts = {}) {
  const daily = db.prepare(`
    SELECT
      dh.work_date,
      COUNT(DISTINCT dh.asset_id) AS used_assets,
      IFNULL(SUM(dh.hours_run), 0) AS run_hours
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date BETWEEN ? AND ?
      AND dh.hours_run > 0
      ${andDailyHoursFleetHoursOnly("dh", "a")}
    GROUP BY dh.work_date
    ORDER BY dh.work_date
  `).all(start, end);

  const available_hours = daily.reduce((acc, d) => acc + (Number(d.used_assets) * scheduled), 0);
  const run_hours = daily.reduce((acc, d) => acc + Number(d.run_hours || 0), 0);

  let downtime_hours;
  if (opts.downtimeHoursOverride != null && Number.isFinite(Number(opts.downtimeHoursOverride))) {
    downtime_hours = Number(opts.downtimeHoursOverride);
  } else {
    const dtCol = getBreakdownDowntimeColumn();
    const dtRow = db.prepare(`
      SELECT IFNULL(SUM(${dtCol}), 0) AS downtime_hours
      FROM breakdowns
      WHERE breakdown_date BETWEEN ? AND ?
    `).get(start, end);
    downtime_hours = Number(dtRow.downtime_hours || 0);
  }

  const availability = available_hours > 0 ? ((available_hours - downtime_hours) / available_hours) * 100 : null;
  const utilization = available_hours > 0 ? (run_hours / available_hours) * 100 : null;

  return {
    available_hours,
    run_hours,
    downtime_hours,
    availability: availability == null ? null : Number(availability.toFixed(2)),
    utilization: utilization == null ? null : Number(utilization.toFixed(2)),
    daily: daily.map(d => ({
      date: d.work_date,
      used_assets: Number(d.used_assets),
      available_hours: Number(d.used_assets) * scheduled,
      run_hours: Number(d.run_hours || 0),
    })),
  };
}

/** Per-asset MTD downtime: from daily logs when present, else from breakdown rows by date. */
function getDowntimeByAssetMtd(mtdStart, end) {
  if (hasBreakdownDowntimeLogsTable()) {
    const cnt = db.prepare(`
      SELECT COUNT(*) AS n FROM breakdown_downtime_logs WHERE log_date BETWEEN ? AND ?
    `).get(mtdStart, end);
    if (Number(cnt?.n || 0) > 0) {
      return db.prepare(`
        SELECT
          a.asset_code,
          a.asset_name,
          COALESCE(SUM(l.hours_down), 0) AS downtime_hours
        FROM breakdown_downtime_logs l
        JOIN breakdowns b ON b.id = l.breakdown_id
        JOIN assets a ON a.id = b.asset_id
        WHERE l.log_date BETWEEN ? AND ?
          AND a.active = 1
          AND a.is_standby = 0
        GROUP BY a.id
        ORDER BY downtime_hours DESC, a.asset_code ASC
      `).all(mtdStart, end).map((r) => ({
        asset_code: r.asset_code,
        asset_name: r.asset_name,
        downtime_hours: Number(Number(r.downtime_hours || 0).toFixed(2)),
      }));
    }
  }
  const dtCol = getBreakdownDowntimeColumn();
  return db.prepare(`
    SELECT
      a.asset_code,
      a.asset_name,
      COALESCE(SUM(b.${dtCol}), 0) AS downtime_hours
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    WHERE b.breakdown_date BETWEEN ? AND ?
      AND a.active = 1
      AND a.is_standby = 0
    GROUP BY a.id
    ORDER BY downtime_hours DESC, a.asset_code ASC
  `).all(mtdStart, end).map((r) => ({
    asset_code: r.asset_code,
    asset_name: r.asset_name,
    downtime_hours: Number(Number(r.downtime_hours || 0).toFixed(2)),
  }));
}

// ---- Excel helpers ----
function addTableSheet(workbook, name, columns, rows, opts = {}) {
  const styled = opts.directorStyle === true;
  const ws = workbook.addWorksheet(name);
  ws.columns = columns.map((c) => ({ header: c.header, key: c.key, width: c.width ?? 18 }));
  const headerRow = ws.getRow(1);
  headerRow.font = styled
    ? { bold: true, size: 11, color: { argb: "FFFFFFFF" } }
    : { bold: true };
  if (styled) {
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1E40AF" },
    };
    headerRow.alignment = { vertical: "middle", wrapText: true };
    headerRow.height = 22;
  }

  for (const r of rows) ws.addRow(r);

  ws.views = [{ state: "frozen", ySplit: 1 }];

  const lastRow = ws.rowCount;
  const lastCol = ws.columnCount;
  const borderColor = styled ? "FFCBD5E1" : "FF2A2A2A";
  for (let i = 1; i <= lastRow; i++) {
    for (let j = 1; j <= lastCol; j++) {
      const cell = ws.getCell(i, j);
      cell.border = {
        top: { style: "thin", color: { argb: borderColor } },
        left: { style: "thin", color: { argb: borderColor } },
        bottom: { style: "thin", color: { argb: borderColor } },
        right: { style: "thin", color: { argb: borderColor } },
      };
      if (i > 1 && styled) {
        cell.alignment = { vertical: "middle", wrapText: j === lastCol };
      }
    }
  }

  return ws;
}

/**
 * Director-friendly first sheet: clear title, grouped KPIs and costs.
 */
function buildDailyExecutiveSummarySheet(wb, p) {
  const ws = wb.addWorksheet("Executive summary");
  ws.views = [{ showGridLines: false }];
  ws.columns = [
    { width: 3 },
    { width: 44 },
    { width: 22 },
  ];

  ws.mergeCells("B1:D2");
  const title = ws.getCell("B1");
  title.value = "AML · IRONLOG — Daily operations report";
  title.font = { size: 18, bold: true, color: { argb: "FF0F172A" } };
  title.alignment = { vertical: "middle", wrapText: true };

  ws.mergeCells("B3:D3");
  const sub = ws.getCell("B3");
  sub.value = `Report date: ${p.date}   ·   Generated: ${new Date().toISOString().slice(0, 16).replace("T", " ")} UTC`;
  sub.font = { size: 11, color: { argb: "FF64748B" } };

  ws.mergeCells("B4:D4");
  ws.getCell("B4").value =
    "Single-day snapshot of fleet hours, reliability KPIs, fuel/lube, breakdowns, maintenance outlook, and estimated direct costs.";
  ws.getCell("B4").font = { size: 10, color: { argb: "FF475569" } };
  ws.getCell("B4").alignment = { wrapText: true };

  const sectionStyle = {
    font: { bold: true, size: 11, color: { argb: "FFFFFFFF" } },
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1E40AF" },
    },
    alignment: { vertical: "middle", indent: 1 },
  };

  let r = 6;
  const section = (label) => {
    ws.mergeCells(`A${r}:D${r}`);
    const c = ws.getCell(`A${r}`);
    c.value = label;
    c.font = sectionStyle.font;
    c.fill = sectionStyle.fill;
    c.alignment = sectionStyle.alignment;
    ws.getRow(r).height = 22;
    r += 1;
  };

  const rowPair = (metric, value, valueIsMoney = false) => {
    ws.getCell(`B${r}`).value = metric;
    ws.getCell(`B${r}`).font = { size: 11, color: { argb: "FF0F172A" } };
    ws.getCell(`C${r}`).value = value;
    ws.getCell(`C${r}`).font = { size: 11, bold: true, color: { argb: "FF0F172A" } };
    if (valueIsMoney && typeof value === "number") {
      ws.getCell(`C${r}`).numFmt = '"R" #,##0.00';
    }
    ws.getRow(r).height = 18;
    r += 1;
  };

  section("Fleet performance (production assets)");
  rowPair(
    "Planning baseline: scheduled hours per production asset (from app header)",
    p.scheduled,
  );
  rowPair(
    "Production assets with run hours today",
    p.kpi.used_assets,
  );
  rowPair(
    "Available fleet hours (used assets × scheduled h)",
    p.kpi.available_hours,
  );
  rowPair("Total run hours recorded", p.kpi.run_hours);
  rowPair("Downtime hours (breakdowns on running fleet)", p.kpi.downtime_hours);
  rowPair(
    "Availability (% of planned hours not lost to downtime; planned = scheduled hours from daily input)",
    p.kpi.availability == null ? "N/A" : `${p.kpi.availability}%`,
  );
  rowPair(
    "Utilization (run hours ÷ planned scheduled hours; planned is not reduced by downtime)",
    p.kpi.utilization == null ? "N/A" : `${p.kpi.utilization}%`,
  );

  r += 1;
  section("Consumables & reliability (day totals)");
  rowPair("Fuel issued (litres)", Number(p.fuel_total.toFixed(2)));
  rowPair("Lubricants / oil issued (qty)", Number(p.oil_total.toFixed(2)));
  rowPair("Breakdown downtime (hours)", Number(p.breakdown_total.toFixed(2)));

  if (p.includeCostEngine !== false) {
    r += 1;
    section("Estimated direct cost (configured rates — validate for finance)");
    rowPair("Fuel", p.fuelCostTotal, true);
    rowPair("Oil / lube", p.lubeCostTotal, true);
    rowPair("Parts (issues linked to work orders)", p.partsCostTotal, true);
    rowPair("Labour (completed work orders)", p.laborCostTotal, true);
    rowPair("Labour hours (completed work orders)", p.laborHoursTotal);
    rowPair("Downtime (estimated cost)", p.downtimeCostTotal, true);
    rowPair("Total estimated direct cost", p.totalCost, true);
    rowPair(
      "Cost per run hour (total cost ÷ run hours)",
      p.costPerRunHour == null ? "N/A" : p.costPerRunHour,
      p.costPerRunHour != null,
    );
  }

  r += 1;
  ws.mergeCells(`B${r}:D${r + 1}`);
  const foot = ws.getCell(`B${r}`);
  foot.value = p.includeCostEngine === false
    ? "Notes: KPIs use production assets with recorded run hours and the scheduled hours shown in the app header."
    : "Notes: KPIs use production assets with recorded run hours and the scheduled hours shown in the app header. Cost lines are indicative from unit rates in IRONLOG; use your finance rules for board packs.";
  foot.font = { size: 9, italic: true, color: { argb: "FF64748B" } };
  foot.alignment = { wrapText: true, vertical: "top" };

  return ws;
}

function gmWeeklyPmComplianceSnapshot(endDate) {
  const plans = db.prepare(`
    SELECT mp.asset_id, mp.interval_hours, mp.last_service_hours
    FROM maintenance_plans mp
    JOIN assets a ON a.id = mp.asset_id
    WHERE mp.active = 1
      AND a.active = 1
      AND a.is_standby = 0
      AND a.archived = 0
  `).all();
  if (!plans.length) {
    return { active_plans: 0, not_overdue: 0, overdue: 0, pct: null };
  }
  let notOverdue = 0;
  for (const p of plans) {
    const current = assetCurrentHoursForGm(p.asset_id);
    const next_due = Number(p.last_service_hours || 0) + Number(p.interval_hours || 0);
    const remaining = next_due - current;
    if (remaining > 0) notOverdue++;
  }
  const overdue = plans.length - notOverdue;
  const pct = Number(((notOverdue / plans.length) * 100).toFixed(2));
  return { active_plans: plans.length, not_overdue: notOverdue, overdue, pct };
}

function gmWeeklyRepairForecast(endDate, horizonDays) {
  const horizon = Math.max(1, Math.min(90, Number(horizonDays || 30)));
  const endD = new Date(`${endDate}T00:00:00`);
  const horizonEnd = new Date(endD);
  horizonEnd.setDate(horizonEnd.getDate() + horizon);
  const horizonEndStr = horizonEnd.toISOString().slice(0, 10);

  const days = 14;
  const avgStart = new Date(endD);
  avgStart.setDate(avgStart.getDate() - (days - 1));
  const avgStartStr = avgStart.toISOString().slice(0, 10);

  const plans = db.prepare(`
    SELECT
      mp.id AS plan_id,
      mp.asset_id,
      mp.service_name,
      mp.interval_hours,
      mp.last_service_hours,
      a.asset_code,
      a.asset_name
    FROM maintenance_plans mp
    JOIN assets a ON a.id = mp.asset_id
    WHERE mp.active = 1
      AND a.active = 1
      AND a.is_standby = 0
      AND a.archived = 0
    ORDER BY a.asset_code ASC, mp.service_name ASC
  `).all();

  const getAvgDaily = db.prepare(`
    SELECT
      COALESCE(SUM(hours_run), 0) AS total_run,
      COUNT(DISTINCT work_date) AS day_count
    FROM daily_hours
    WHERE asset_id = ?
      AND is_used = 1
      AND hours_run > 0
      AND work_date BETWEEN ? AND ?
  `);

  const addDays = (dateStr, add) => {
    const d = new Date(`${dateStr}T00:00:00`);
    d.setDate(d.getDate() + Math.round(add));
    return d.toISOString().slice(0, 10);
  };

  const pm_rows = [];
  for (const p of plans) {
    const current = assetCurrentHoursForGm(p.asset_id);
    const next_due = Number(p.last_service_hours || 0) + Number(p.interval_hours || 0);
    const remaining = next_due - current;
    const avgRow = getAvgDaily.get(Number(p.asset_id), avgStartStr, endDate);
    const totalRun = Number(avgRow?.total_run || 0);
    const dayCount = Number(avgRow?.day_count || 0);
    const avgDaily = dayCount > 0 ? totalRun / dayCount : 0;
    const estDays = avgDaily > 0 ? Math.max(0, remaining / avgDaily) : null;
    const estDate = estDays == null ? null : addDays(endDate, estDays);
    if (estDate && estDate > endDate && estDate <= horizonEndStr) {
      pm_rows.push({
        asset_code: p.asset_code,
        asset_name: p.asset_name,
        type: "PM / service",
        detail: p.service_name,
        est_date: estDate,
        remaining_hours: Number(remaining.toFixed(2)),
      });
    }
  }

  const open_breakdown_repairs = db.prepare(`
    SELECT
      w.id AS wo_id,
      w.opened_at,
      w.status,
      a.asset_code,
      a.asset_name,
      b.description
    FROM work_orders w
    JOIN assets a ON a.id = w.asset_id
    LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
    WHERE w.source = 'breakdown'
      AND REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
    ORDER BY w.opened_at ASC
  `).all().map((r) => ({
    asset_code: r.asset_code,
    asset_name: r.asset_name,
    type: "Breakdown repair",
    detail: compactCell(r.description || `WO #${r.wo_id}`, 120),
    est_date: "",
    wo_id: r.wo_id,
    status: r.status,
    opened_at: r.opened_at,
  }));

  return { pm_rows, open_breakdown_repairs, horizon_end: horizonEndStr };
}

/**
 * GM weekly pack: Maintenance & Engineering KPIs (narrative + metric blocks).
 */
function buildGmWeeklyExecutiveSheet(wb, p) {
  const ws = wb.addWorksheet("M & E summary");
  ws.views = [{ showGridLines: false }];
  ws.columns = [
    { width: 3 },
    { width: 48 },
    { width: 22 },
  ];

  ws.mergeCells("B1:D2");
  const title = ws.getCell("B1");
  title.value = "AML · IRONLOG — GM Weekly (Maintenance & Engineering)";
  title.font = { size: 18, bold: true, color: { argb: "FF0F172A" } };
  title.alignment = { vertical: "middle", wrapText: true };

  ws.mergeCells("B3:D3");
  ws.getCell("B3").value =
    `Month-to-date: ${p.mtd_start} → ${p.end} (${p.mtd_day_count} calendar days)   ·   Scheduled hours / asset: ${p.scheduled}   ·   Generated: ${new Date().toISOString().slice(0, 16).replace("T", " ")} UTC`;
  ws.getCell("B3").font = { size: 11, color: { argb: "FF64748B" } };

  ws.mergeCells("B4:D4");
  ws.getCell("B4").value =
    "Availability and utilization are month-to-date (first of month through report date). Availability = (planned − downtime) ÷ planned. Utilization = run hours ÷ planned (planned = scheduled hours from daily input; not reduced by downtime). Breakdown downtime uses daily downtime logs when present so multi-day incidents carry across the month.";
  ws.getCell("B4").font = { size: 10, color: { argb: "FF475569" } };
  ws.getCell("B4").alignment = { wrapText: true };

  const sectionStyle = {
    font: { bold: true, size: 11, color: { argb: "FFFFFFFF" } },
    fill: {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF1E40AF" },
    },
    alignment: { vertical: "middle", indent: 1 },
  };

  let r = 6;
  const section = (label) => {
    ws.mergeCells(`A${r}:D${r}`);
    const c = ws.getCell(`A${r}`);
    c.value = label;
    c.font = sectionStyle.font;
    c.fill = sectionStyle.fill;
    c.alignment = sectionStyle.alignment;
    ws.getRow(r).height = 22;
    r += 1;
  };

  const rowPair = (metric, value) => {
    ws.getCell(`B${r}`).value = metric;
    ws.getCell(`B${r}`).font = { size: 11, color: { argb: "FF0F172A" } };
    ws.getCell(`C${r}`).value = value;
    ws.getCell(`C${r}`).font = { size: 11, bold: true, color: { argb: "FF0F172A" } };
    ws.getRow(r).height = 18;
    r += 1;
  };

  section("i–iii. Fleet performance (month-to-date)");
  rowPair(
    "Equipment availability (% of planned hours not lost to downtime, MTD)",
    p.kpi.availability == null ? "N/A" : `${p.kpi.availability}%`,
  );
  rowPair(
    "Utilization (run hours ÷ planned scheduled hours MTD; same planned base as availability denominator)",
    p.kpi.utilization == null ? "N/A" : `${p.kpi.utilization}%`,
  );
  rowPair(
    p.uses_downtime_logs
      ? "Breakdown downtime (hours, MTD — sum of daily downtime logs; carries multi-day incidents)"
      : "Breakdown downtime (hours, MTD — from incident report dates; add daily logs for carry-over)",
    p.kpi.downtime_hours,
  );

  r += 1;
  section("iv–v. Reliability (month-to-date; downtime matches fleet KPIs above)");
  rowPair(
    p.rel.failures_in_period_mode === "activity_or_report"
      ? "Failure count (distinct incidents: report date in MTD or downtime logged in MTD)"
      : "Failure count (breakdown records with report date in MTD)",
    p.rel.failure_count,
  );
  rowPair("Operating hours (sum of run hours on used days, MTD)", p.rel.operating_hours);
  rowPair("MTBF (mean time between failures, hours)", p.rel.mtbf_hours == null ? "N/A" : p.rel.mtbf_hours);
  rowPair("MTTR (mean time to repair — avg downtime per failure, hours)", p.rel.mttr_hours == null ? "N/A" : p.rel.mttr_hours);

  r += 1;
  section("vi. Preventative maintenance compliance");
  rowPair(
    "PM compliance (% of active PM plans not past meter due at period end)",
    p.pm.pct == null ? "N/A" : `${p.pm.pct}%`,
  );
  rowPair("Active PM plans (assets in scope)", p.pm.active_plans);
  rowPair("Plans not overdue (meter)", p.pm.not_overdue);
  rowPair("Plans overdue (meter)", p.pm.overdue);

  r += 1;
  section("vii. Critical spares (summary)");
  rowPair("Critical parts tracked", p.spares.critical_parts);
  rowPair("Critical lines below minimum stock", p.spares.below_min);

  r += 1;
  section("viii. Major repair / PM outlook");
  rowPair(
    `PM services with estimated date in next ${p.forecast_horizon_days} days (from usage trend)`,
    p.forecast.pm_count,
  );
  rowPair("Open breakdown repairs (active work orders)", p.forecast.open_wo_count);

  r += 1;
  ws.mergeCells(`B${r}:D${r + 2}`);
  const foot = ws.getCell(`B${r}`);
  foot.value =
    "Notes: Availability, utilization, and reliability KPIs use month-to-date from the first of the report month through the download date. Utilization divides run hours by planned scheduled hours (same planned base as the availability denominator, not post-downtime available hours). Downtime for availability and MTTR uses summed daily breakdown_downtime_logs when those rows exist in the period (so downtime carries from the day it was logged). If no daily logs exist for the month, downtime falls back to summing incidents by breakdown report date. Failure count includes any distinct incident with downtime logged in MTD even if the incident was reported earlier. MTBF = operating hours ÷ failure count; MTTR = total downtime hours ÷ failure count. PM compliance is a meter snapshot at report date. See the Downtime by asset sheet for MTD hours per machine.";
  foot.font = { size: 9, italic: true, color: { argb: "FF64748B" } };
  foot.alignment = { wrapText: true, vertical: "top" };

  return ws;
}

export default async function reportsRoutes(app) {
  const dataRoot = process.env.IRONLOG_DATA_DIR || process.cwd();
  function hasColumn(table, col) {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name || "") === String(col));
  }
  function pickExistingColumn(table, candidates, fallback) {
    for (const c of candidates) {
      if (hasColumn(table, c)) return c;
    }
    return fallback;
  }
  function resolveStorageAbs(relPath) {
    const rel = String(relPath || "").replace(/\\/g, "/").replace(/^\/+/, "");
    if (!rel) return "";
    const fromData = path.join(dataRoot, rel);
    if (fs.existsSync(fromData)) return fromData;
    return path.join(process.cwd(), rel);
  }

  function ensureColumn(table, colName, colDef) {
    if (!hasColumn(table, colName)) {
      db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
    }
  }

  ensureColumn("assets", "fuel_cost_per_liter", "fuel_cost_per_liter REAL");
  ensureColumn("assets", "downtime_cost_per_hour", "downtime_cost_per_hour REAL");
  ensureColumn("parts", "unit_cost", "unit_cost REAL DEFAULT 0");
  ensureColumn("oil_logs", "unit_cost", "unit_cost REAL");
  ensureColumn("fuel_logs", "unit_cost_per_liter", "unit_cost_per_liter REAL");
  ensureColumn("fuel_logs", "hours_run", "hours_run REAL");
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
  db.prepare(`
    CREATE TABLE IF NOT EXISTS site_rain_days (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL DEFAULT 'default',
      rain_date TEXT NOT NULL,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(site_code, rain_date)
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS operations_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      op_date TEXT NOT NULL DEFAULT (date('now')),
      tonnes_moved REAL,
      product_type TEXT,
      product_produced REAL,
      trucks_loaded INTEGER,
      weighbridge_amount REAL,
      trucks_delivered INTEGER,
      product_delivered REAL,
      client_delivered_to TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS report_templates (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      dataset TEXT NOT NULL,
      columns_json TEXT NOT NULL,
      filters_json TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_report_templates_dataset ON report_templates(dataset)`).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS report_subscriptions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      report_type TEXT NOT NULL,
      channel TEXT NOT NULL,
      recipients TEXT NOT NULL,
      schedule_frequency TEXT NOT NULL DEFAULT 'weekly',
      send_time TEXT NOT NULL DEFAULT '07:00',
      day_of_week INTEGER,
      day_of_month INTEGER,
      active INTEGER NOT NULL DEFAULT 1,
      filters_json TEXT,
      last_sent_at TEXT,
      next_run_at TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS report_delivery_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      subscription_id INTEGER,
      report_type TEXT NOT NULL,
      channel TEXT NOT NULL,
      recipients TEXT,
      status TEXT NOT NULL,
      detail TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_report_subscriptions_next ON report_subscriptions(active, next_run_at)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_report_delivery_logs_sub ON report_delivery_logs(subscription_id, created_at DESC)`).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS smtp_settings (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      host TEXT,
      port INTEGER,
      secure INTEGER NOT NULL DEFAULT 0,
      username TEXT,
      password_enc TEXT,
      from_email TEXT,
      from_name TEXT,
      updated_by TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`INSERT INTO smtp_settings (id) SELECT 1 WHERE NOT EXISTS (SELECT 1 FROM smtp_settings WHERE id = 1)`).run();

  function requestRoles(req) {
    const fromMany = String(req.headers["x-user-roles"] || "")
      .split(",")
      .map((x) => String(x || "").trim().toLowerCase())
      .filter(Boolean);
    const fromSingle = String(req.headers["x-user-role"] || "")
      .split(",")
      .map((x) => String(x || "").trim().toLowerCase())
      .filter(Boolean);
    const merged = Array.from(new Set([...fromMany, ...fromSingle]));
    return merged.length ? merged : ["admin"];
  }
  function requireAdmin(req, reply) {
    const roles = requestRoles(req);
    if (!roles.includes("admin") && !roles.includes("supervisor")) {
      reply.code(403).send({ ok: false, error: "admin or supervisor role required" });
      return false;
    }
    return true;
  }
  function smtpSecret() {
    const raw = String(process.env.IRONLOG_SMTP_SECRET || process.env.IRONLOG_AUTH_SECRET || "").trim();
    if (!raw) return "IRONLOG_SMTP_DEFAULT_SECRET_CHANGE_ME";
    return raw;
  }
  function encryptSecret(plain) {
    const iv = crypto.randomBytes(12);
    const key = crypto.createHash("sha256").update(smtpSecret()).digest();
    const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
    const encrypted = Buffer.concat([cipher.update(String(plain || ""), "utf8"), cipher.final()]);
    const tag = cipher.getAuthTag();
    return `${iv.toString("base64")}.${tag.toString("base64")}.${encrypted.toString("base64")}`;
  }
  function decryptSecret(cipherText) {
    const raw = String(cipherText || "").trim();
    if (!raw) return "";
    const [ivB64, tagB64, encB64] = raw.split(".");
    if (!ivB64 || !tagB64 || !encB64) return "";
    const iv = Buffer.from(ivB64, "base64");
    const tag = Buffer.from(tagB64, "base64");
    const enc = Buffer.from(encB64, "base64");
    const key = crypto.createHash("sha256").update(smtpSecret()).digest();
    const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
    decipher.setAuthTag(tag);
    const out = Buffer.concat([decipher.update(enc), decipher.final()]);
    return out.toString("utf8");
  }
  function getSmtpSettingsRow() {
    return db.prepare(`
      SELECT id, host, port, secure, username, password_enc, from_email, from_name, updated_by, updated_at
      FROM smtp_settings
      WHERE id = 1
    `).get() || null;
  }
  function smtpPublicPayload(row) {
    const r = row || {};
    return {
      host: String(r.host || ""),
      port: Number(r.port || 587),
      secure: Number(r.secure || 0) === 1 ? 1 : 0,
      username: String(r.username || ""),
      from_email: String(r.from_email || ""),
      from_name: String(r.from_name || ""),
      has_password: Boolean(String(r.password_enc || "").trim()),
      updated_by: String(r.updated_by || ""),
      updated_at: r.updated_at || null,
    };
  }
  function buildSmtpTransport() {
    const row = getSmtpSettingsRow();
    if (!row) return null;
    const host = String(row.host || "").trim();
    const username = String(row.username || "").trim();
    const fromEmail = String(row.from_email || "").trim();
    const password = decryptSecret(row.password_enc || "");
    if (!host || !username || !fromEmail || !password) return null;
    const port = Math.max(1, Number(row.port || 587));
    const secure = Number(row.secure || 0) === 1;
    const transporter = nodemailer.createTransport({
      host,
      port,
      secure,
      auth: { user: username, pass: password },
    });
    return {
      transporter,
      from: row.from_name ? `"${String(row.from_name).replace(/"/g, "")}" <${fromEmail}>` : fromEmail,
    };
  }

  const allowedReportTypes = new Set(["fuel_benchmark_xlsx", "executive_kpi_pack_xlsx", "maintenance_insights_xlsx"]);
  const allowedChannels = new Set(["email", "whatsapp"]);
  const allowedFrequencies = new Set(["daily", "weekly", "monthly"]);
  function parseRecipients(raw) {
    return Array.from(new Set(
      String(raw || "")
        .split(",")
        .map((x) => String(x || "").trim())
        .filter(Boolean)
    )).slice(0, 50);
  }
  function parseTimeHhMm(raw) {
    const s = String(raw || "").trim();
    return /^\d{2}:\d{2}$/.test(s) ? s : "07:00";
  }
  function toIsoNoMs(d) {
    return new Date(d).toISOString().slice(0, 19) + "Z";
  }
  function nextRunForSchedule(schedule, now = new Date()) {
    const freq = String(schedule.schedule_frequency || "weekly").trim().toLowerCase();
    const time = parseTimeHhMm(schedule.send_time);
    const [hh, mm] = time.split(":").map((n) => Number(n));
    const next = new Date(now);
    next.setSeconds(0, 0);
    next.setHours(hh, mm, 0, 0);
    if (freq === "daily") {
      if (next <= now) next.setDate(next.getDate() + 1);
      return toIsoNoMs(next);
    }
    if (freq === "weekly") {
      const dow = Math.max(0, Math.min(6, Number(schedule.day_of_week ?? 1)));
      const curDow = next.getDay();
      let delta = dow - curDow;
      if (delta < 0 || (delta === 0 && next <= now)) delta += 7;
      next.setDate(next.getDate() + delta);
      return toIsoNoMs(next);
    }
    const dom = Math.max(1, Math.min(28, Number(schedule.day_of_month ?? 1)));
    next.setDate(dom);
    if (next <= now) {
      next.setMonth(next.getMonth() + 1);
      next.setDate(dom);
    }
    return toIsoNoMs(next);
  }
  function reportLinkForType(reportType, filters = {}) {
    const f = filters && typeof filters === "object" ? filters : {};
    if (reportType === "fuel_benchmark_xlsx") {
      const start = isDate(f.start) ? String(f.start) : todayYmd().slice(0, 8) + "01";
      const end = isDate(f.end) ? String(f.end) : todayYmd();
      return `/api/reports/fuel-benchmark.xlsx?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
    }
    if (reportType === "executive_kpi_pack_xlsx") {
      const period = String(f.period_type || "weekly").trim().toLowerCase();
      const start = isDate(f.start) ? String(f.start) : todayYmd().slice(0, 8) + "01";
      const end = isDate(f.end) ? String(f.end) : todayYmd();
      const siteCodes = String(f.site_codes || "main").trim();
      return `/api/reports/executive-kpi-pack.xlsx?period_type=${encodeURIComponent(period)}&start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&site_codes=${encodeURIComponent(siteCodes)}`;
    }
    const start = isDate(f.start) ? String(f.start) : todayYmd().slice(0, 8) + "01";
    const end = isDate(f.end) ? String(f.end) : todayYmd();
    return `/api/maintenance/insights.xlsx?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  }
  async function deliverSubscription(subRow, manual = false) {
    const id = Number(subRow.id || 0);
    const reportType = String(subRow.report_type || "");
    const channel = String(subRow.channel || "");
    const recipients = parseRecipients(subRow.recipients || "");
    let filters = {};
    try { filters = JSON.parse(String(subRow.filters_json || "{}")); } catch {}
    const link = reportLinkForType(reportType, filters);
    const payload = {
      subscription_id: id,
      name: String(subRow.name || ""),
      report_type: reportType,
      channel,
      recipients,
      report_link: link,
      manual: Boolean(manual),
      generated_at: new Date().toISOString(),
    };
    let status = "simulated";
    let detail = "Logged only";
    if (channel === "email") {
      const smtp = buildSmtpTransport();
      if (smtp) {
        await smtp.transporter.sendMail({
          from: smtp.from,
          to: recipients.join(", "),
          subject: `IRONLOG Report: ${String(subRow.name || reportType)}`,
          text: `Your IRONLOG report is ready.\n\nReport: ${reportType}\nLink: ${link}\nGenerated: ${new Date().toISOString()}`,
        });
        status = "sent";
        detail = "SMTP email sent";
      } else {
        const emailWebhook = String(process.env.REPORT_EMAIL_WEBHOOK_URL || "").trim();
        if (emailWebhook) {
          const resp = await fetch(emailWebhook, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload),
          });
          if (!resp.ok) {
            status = "failed";
            detail = `Email webhook failed: HTTP ${resp.status}`;
          } else {
            status = "sent";
            detail = "Email webhook accepted";
          }
        } else {
          detail = "No SMTP or email webhook configured";
        }
      }
    } else {
      const whatsappWebhook = String(process.env.REPORT_WHATSAPP_WEBHOOK_URL || "").trim();
      if (whatsappWebhook) {
        const resp = await fetch(whatsappWebhook, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });
        if (!resp.ok) {
          status = "failed";
          detail = `WhatsApp webhook failed: HTTP ${resp.status}`;
        } else {
          status = "sent";
          detail = "WhatsApp webhook accepted";
        }
      } else {
        detail = "No WhatsApp webhook configured";
      }
    }
    db.prepare(`
      INSERT INTO report_delivery_logs (subscription_id, report_type, channel, recipients, status, detail, created_at)
      VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
    `).run(id, reportType, channel, recipients.join(","), status, detail);
    const nextRun = nextRunForSchedule(subRow, new Date());
    db.prepare(`
      UPDATE report_subscriptions
      SET last_sent_at = datetime('now'), next_run_at = ?, updated_at = datetime('now')
      WHERE id = ?
    `).run(nextRun, id);
    return { status, detail, report_link: link };
  }

  const reportDatasets = {
    work_orders: {
      table: "work_orders",
      defaultOrder: "id DESC",
      columns: {
        id: "id",
        asset_id: "asset_id",
        status: "status",
        source: "source",
        title: "title",
        description: "description",
        priority: "priority",
        opened_at: "opened_at",
        due_at: "due_at",
        assigned_at: "assigned_at",
        completed_at: "completed_at",
        closed_at: "closed_at",
        artisan_name: "artisan_name",
      },
      dateColumn: "opened_at",
      assetColumn: "asset_id",
      statusColumn: "status",
    },
    fuel_logs: {
      table: "fuel_logs",
      defaultOrder: "log_date DESC, id DESC",
      columns: {
        id: "id",
        asset_id: "asset_id",
        log_date: "log_date",
        liters: "liters",
        cost_total: "cost_total",
        meter_unit: "meter_unit",
        meter_run_value: "meter_run_value",
        hours_run: "hours_run",
        notes: "notes",
      },
      dateColumn: "log_date",
      assetColumn: "asset_id",
    },
    manager_inspections: {
      table: "manager_inspections",
      defaultOrder: "id DESC",
      columns: {
        id: "id",
        asset_id: "asset_id",
        inspection_date: "inspection_date",
        status: "status",
        comments: "comments",
        notes: "notes",
        defect_severity: "defect_severity",
        defect_component: "defect_component",
        defect_risk: "defect_risk",
        recommended_action: "recommended_action",
      },
      dateColumn: "inspection_date",
      assetColumn: "asset_id",
      statusColumn: "status",
    },
  };
  function hasTable(tableName) {
    return Boolean(
      db.prepare("SELECT 1 FROM sqlite_master WHERE type='table' AND name = ?").get(String(tableName || "").trim())
    );
  }
  function datasetWithAvailableColumns(datasetKey) {
    const key = String(datasetKey || "").trim();
    const ds = reportDatasets[key];
    if (!ds || !hasTable(ds.table)) return null;
    const tableCols = new Set(db.prepare(`PRAGMA table_info(${ds.table})`).all().map((r) => String(r.name || "")));
    const availableColumns = Object.entries(ds.columns)
      .filter(([, sqlCol]) => tableCols.has(String(sqlCol)))
      .map(([id, sqlCol]) => ({ id, sql: sqlCol }));
    if (!availableColumns.length) return null;
    return { ...ds, key, availableColumns };
  }

  app.get("/custom-builder/meta", async (_req, reply) => {
    const datasets = Object.keys(reportDatasets)
      .map((k) => datasetWithAvailableColumns(k))
      .filter(Boolean)
      .map((d) => ({
        key: d.key,
        table: d.table,
        columns: d.availableColumns.map((c) => c.id),
      }));
    return reply.send({ ok: true, datasets });
  });

  app.get("/custom-builder/templates", async (_req, reply) => {
    const rows = db.prepare(`
      SELECT id, name, dataset, columns_json, filters_json, created_by, created_at, updated_at
      FROM report_templates
      ORDER BY updated_at DESC, id DESC
      LIMIT 200
    `).all();
    const templates = rows.map((r) => {
      let columns = [];
      let filters = {};
      try { columns = JSON.parse(String(r.columns_json || "[]")); } catch {}
      try { filters = JSON.parse(String(r.filters_json || "{}")); } catch {}
      return {
        id: Number(r.id || 0),
        name: String(r.name || ""),
        dataset: String(r.dataset || ""),
        columns: Array.isArray(columns) ? columns : [],
        filters: filters && typeof filters === "object" ? filters : {},
        created_by: String(r.created_by || ""),
        created_at: r.created_at,
        updated_at: r.updated_at,
      };
    });
    return reply.send({ ok: true, templates });
  });

  app.post("/custom-builder/templates", async (req, reply) => {
    const body = req.body || {};
    const name = String(body.name || "").trim();
    const dataset = String(body.dataset || "").trim();
    const ds = datasetWithAvailableColumns(dataset);
    if (!name) return reply.code(400).send({ ok: false, error: "Template name is required" });
    if (!ds) return reply.code(400).send({ ok: false, error: "Invalid dataset" });
    const validCols = new Set(ds.availableColumns.map((c) => c.id));
    const columns = Array.from(new Set((Array.isArray(body.columns) ? body.columns : []).map((c) => String(c || "").trim())))
      .filter((c) => validCols.has(c))
      .slice(0, 25);
    if (!columns.length) return reply.code(400).send({ ok: false, error: "Select at least one valid column" });
    const filters = body.filters && typeof body.filters === "object" ? body.filters : {};
    const who = String(req.headers["x-user-name"] || "system");
    const id = Number(body.id || 0);
    if (id > 0) {
      const existing = db.prepare("SELECT id FROM report_templates WHERE id = ? LIMIT 1").get(id);
      if (!existing) return reply.code(404).send({ ok: false, error: "Template not found" });
      db.prepare(`
        UPDATE report_templates
        SET name = ?, dataset = ?, columns_json = ?, filters_json = ?, created_by = ?, updated_at = datetime('now')
        WHERE id = ?
      `).run(name, ds.key, JSON.stringify(columns), JSON.stringify(filters), who, id);
      return reply.send({ ok: true, id });
    }
    const out = db.prepare(`
      INSERT INTO report_templates (name, dataset, columns_json, filters_json, created_by, updated_at)
      VALUES (?, ?, ?, ?, ?, datetime('now'))
    `).run(name, ds.key, JSON.stringify(columns), JSON.stringify(filters), who);
    return reply.send({ ok: true, id: Number(out.lastInsertRowid || 0) });
  });

  app.delete("/custom-builder/templates/:id", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ ok: false, error: "Invalid template id" });
    const out = db.prepare("DELETE FROM report_templates WHERE id = ?").run(id);
    if (!Number(out.changes || 0)) return reply.code(404).send({ ok: false, error: "Template not found" });
    return reply.send({ ok: true });
  });

  function runCustomBuilderQuery(body) {
    const payload = body || {};
    const dataset = String(payload.dataset || "").trim();
    const ds = datasetWithAvailableColumns(dataset);
    if (!ds) throw new Error("Invalid dataset");
    const validCols = new Set(ds.availableColumns.map((c) => c.id));
    const picked = Array.from(new Set((Array.isArray(payload.columns) ? payload.columns : []).map((c) => String(c || "").trim())))
      .filter((c) => validCols.has(c))
      .slice(0, 25);
    if (!picked.length) throw new Error("Select at least one valid column");
    const filters = payload.filters && typeof payload.filters === "object" ? payload.filters : {};
    const limitNum = Math.max(1, Math.min(500, Number(filters.limit || payload.limit || 100)));

    const where = [];
    const params = [];
    if (ds.dateColumn) {
      const start = String(filters.start || "").trim();
      const end = String(filters.end || "").trim();
      if (isDate(start)) { where.push(`date(${ds.dateColumn}) >= date(?)`); params.push(start); }
      if (isDate(end)) { where.push(`date(${ds.dateColumn}) <= date(?)`); params.push(end); }
    }
    if (ds.assetColumn) {
      const assetId = Number(filters.asset_id || 0);
      if (assetId > 0) { where.push(`${ds.assetColumn} = ?`); params.push(assetId); }
    }
    if (ds.statusColumn) {
      const status = String(filters.status || "").trim();
      if (status) { where.push(`LOWER(COALESCE(${ds.statusColumn},'')) = LOWER(?)`); params.push(status); }
    }
    const selectSql = picked.map((k) => ds.columns[k]).join(", ");
    const sql = `
      SELECT ${selectSql}
      FROM ${ds.table}
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY ${ds.defaultOrder}
      LIMIT ${limitNum}
    `;
    const rows = db.prepare(sql).all(...params);
    return { dataset: ds.key, columns: picked, rows, count: rows.length, limit: limitNum };
  }

  app.post("/custom-builder/preview", async (req, reply) => {
    try {
      const result = runCustomBuilderQuery(req.body || {});
      return reply.send({ ok: true, ...result });
    } catch (err) {
      return reply.code(400).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/custom-builder/export.xlsx", async (req, reply) => {
    try {
      const result = runCustomBuilderQuery(req.body || {});
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Custom Report");
      const headers = result.columns.map((c) => String(c || ""));
      ws.addRow(headers);
      for (const r of result.rows) {
        ws.addRow(headers.map((h) => r?.[h] ?? ""));
      }
      ws.views = [{ state: "frozen", ySplit: 1 }];
      ws.columns.forEach((col) => {
        let width = 12;
        col.eachCell({ includeEmpty: true }, (cell) => {
          width = Math.max(width, Math.min(48, String(cell.value ?? "").length + 2));
        });
        col.width = width;
      });
      const safeDataset = String(result.dataset || "report").replace(/[^a-z0-9_-]/gi, "_");
      const stamp = new Date().toISOString().slice(0, 10);
      const buf = await wb.xlsx.writeBuffer();
      return reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Content-Disposition", `attachment; filename="IRONLOG_Custom_${safeDataset}_${stamp}.xlsx"`)
        .send(Buffer.from(buf));
    } catch (err) {
      return reply.code(400).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/smtp-settings", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    return reply.send({ ok: true, settings: smtpPublicPayload(getSmtpSettingsRow()) });
  });

  app.post("/smtp-settings", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const body = req.body || {};
    const host = String(body.host || "").trim();
    const port = Math.max(1, Number(body.port || 587));
    const secure = Number(body.secure || 0) === 1 ? 1 : 0;
    const username = String(body.username || "").trim();
    const fromEmail = String(body.from_email || "").trim();
    const fromName = String(body.from_name || "").trim();
    const password = String(body.password || "");
    if (!host) return reply.code(400).send({ ok: false, error: "SMTP host is required" });
    if (!username) return reply.code(400).send({ ok: false, error: "SMTP username is required" });
    if (!fromEmail) return reply.code(400).send({ ok: false, error: "From email is required" });
    const who = String(req.headers["x-user-name"] || "system");
    const existing = getSmtpSettingsRow() || {};
    const passwordEnc = password ? encryptSecret(password) : String(existing.password_enc || "");
    db.prepare(`
      UPDATE smtp_settings
      SET host = ?, port = ?, secure = ?, username = ?, password_enc = ?, from_email = ?, from_name = ?, updated_by = ?, updated_at = datetime('now')
      WHERE id = 1
    `).run(host, port, secure, username, passwordEnc, fromEmail, fromName || null, who);
    return reply.send({ ok: true, settings: smtpPublicPayload(getSmtpSettingsRow()) });
  });

  app.post("/smtp-settings/test", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const toRaw = String(req.body?.to || "").trim();
    if (!toRaw) return reply.code(400).send({ ok: false, error: "Recipient email is required" });
    const recipients = parseRecipients(toRaw);
    if (!recipients.length) return reply.code(400).send({ ok: false, error: "Invalid recipient list" });
    try {
      const smtp = buildSmtpTransport();
      if (!smtp) return reply.code(400).send({ ok: false, error: "SMTP settings incomplete. Save host/username/password/from first." });
      await smtp.transporter.sendMail({
        from: smtp.from,
        to: recipients.join(", "),
        subject: "IRONLOG SMTP test email",
        text: `SMTP test successful at ${new Date().toISOString()}`,
      });
      return reply.send({ ok: true, message: "Test email sent" });
    } catch (err) {
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/subscriptions", async (_req, reply) => {
    const rows = db.prepare(`
      SELECT
        id, name, report_type, channel, recipients, schedule_frequency, send_time,
        day_of_week, day_of_month, active, filters_json, last_sent_at, next_run_at,
        created_by, created_at, updated_at
      FROM report_subscriptions
      ORDER BY active DESC, next_run_at ASC, id DESC
      LIMIT 300
    `).all();
    const subscriptions = rows.map((r) => {
      let filters = {};
      try { filters = JSON.parse(String(r.filters_json || "{}")); } catch {}
      return {
        id: Number(r.id || 0),
        name: String(r.name || ""),
        report_type: String(r.report_type || ""),
        channel: String(r.channel || ""),
        recipients: parseRecipients(r.recipients || ""),
        schedule_frequency: String(r.schedule_frequency || "weekly"),
        send_time: parseTimeHhMm(r.send_time),
        day_of_week: r.day_of_week == null ? null : Number(r.day_of_week),
        day_of_month: r.day_of_month == null ? null : Number(r.day_of_month),
        active: Number(r.active || 0) === 1 ? 1 : 0,
        filters: filters && typeof filters === "object" ? filters : {},
        last_sent_at: r.last_sent_at,
        next_run_at: r.next_run_at,
        created_by: String(r.created_by || ""),
        created_at: r.created_at,
        updated_at: r.updated_at,
      };
    });
    return reply.send({ ok: true, subscriptions });
  });

  app.post("/subscriptions", async (req, reply) => {
    const body = req.body || {};
    const id = Number(body.id || 0);
    const name = String(body.name || "").trim();
    const reportType = String(body.report_type || "").trim().toLowerCase();
    const channel = String(body.channel || "").trim().toLowerCase();
    const freq = String(body.schedule_frequency || "weekly").trim().toLowerCase();
    const sendTime = parseTimeHhMm(body.send_time);
    const dayOfWeek = body.day_of_week == null ? null : Math.max(0, Math.min(6, Number(body.day_of_week)));
    const dayOfMonth = body.day_of_month == null ? null : Math.max(1, Math.min(28, Number(body.day_of_month)));
    const active = Number(body.active ?? 1) === 1 ? 1 : 0;
    const recipients = parseRecipients(body.recipients || "");
    const filters = body.filters && typeof body.filters === "object" ? body.filters : {};
    if (!name) return reply.code(400).send({ ok: false, error: "Name is required" });
    if (!allowedReportTypes.has(reportType)) return reply.code(400).send({ ok: false, error: "Invalid report_type" });
    if (!allowedChannels.has(channel)) return reply.code(400).send({ ok: false, error: "Invalid channel" });
    if (!allowedFrequencies.has(freq)) return reply.code(400).send({ ok: false, error: "Invalid schedule_frequency" });
    if (!recipients.length) return reply.code(400).send({ ok: false, error: "At least one recipient is required" });
    const who = String(req.headers["x-user-name"] || "system");
    const nextRun = active ? nextRunForSchedule({
      schedule_frequency: freq,
      send_time: sendTime,
      day_of_week: dayOfWeek,
      day_of_month: dayOfMonth,
    }, new Date()) : null;
    if (id > 0) {
      const ex = db.prepare(`SELECT id FROM report_subscriptions WHERE id = ? LIMIT 1`).get(id);
      if (!ex) return reply.code(404).send({ ok: false, error: "Subscription not found" });
      db.prepare(`
        UPDATE report_subscriptions
        SET
          name = ?, report_type = ?, channel = ?, recipients = ?, schedule_frequency = ?, send_time = ?,
          day_of_week = ?, day_of_month = ?, active = ?, filters_json = ?, next_run_at = ?, updated_at = datetime('now')
        WHERE id = ?
      `).run(
        name, reportType, channel, recipients.join(","), freq, sendTime, dayOfWeek, dayOfMonth, active,
        JSON.stringify(filters), nextRun, id
      );
      return reply.send({ ok: true, id });
    }
    const out = db.prepare(`
      INSERT INTO report_subscriptions (
        name, report_type, channel, recipients, schedule_frequency, send_time, day_of_week, day_of_month,
        active, filters_json, next_run_at, created_by, updated_at
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
    `).run(
      name, reportType, channel, recipients.join(","), freq, sendTime, dayOfWeek, dayOfMonth,
      active, JSON.stringify(filters), nextRun, who
    );
    return reply.send({ ok: true, id: Number(out.lastInsertRowid || 0) });
  });

  app.delete("/subscriptions/:id", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
    const out = db.prepare(`DELETE FROM report_subscriptions WHERE id = ?`).run(id);
    if (!Number(out.changes || 0)) return reply.code(404).send({ ok: false, error: "Subscription not found" });
    return reply.send({ ok: true });
  });

  app.post("/subscriptions/:id/send-now", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
    const row = db.prepare(`SELECT * FROM report_subscriptions WHERE id = ? LIMIT 1`).get(id);
    if (!row) return reply.code(404).send({ ok: false, error: "Subscription not found" });
    try {
      const out = await deliverSubscription(row, true);
      return reply.send({ ok: true, ...out });
    } catch (err) {
      db.prepare(`
        INSERT INTO report_delivery_logs (subscription_id, report_type, channel, recipients, status, detail, created_at)
        VALUES (?, ?, ?, ?, 'failed', ?, datetime('now'))
      `).run(id, row.report_type, row.channel, row.recipients, String(err.message || err));
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/subscriptions/logs", async (req, reply) => {
    const limit = Math.max(1, Math.min(200, Number(req.query?.limit || 50)));
    const rows = db.prepare(`
      SELECT id, subscription_id, report_type, channel, recipients, status, detail, created_at
      FROM report_delivery_logs
      ORDER BY created_at DESC, id DESC
      LIMIT ?
    `).all(limit);
    return reply.send({ ok: true, rows });
  });

  function costDefaults() {
    const rows = db.prepare(`
      SELECT key, value
      FROM cost_settings
      WHERE key IN (
        'fuel_cost_per_liter_default',
        'lube_cost_per_qty_default',
        'labor_cost_per_hour_default',
        'downtime_cost_per_hour_default'
      )
    `).all();
    const d = {
      fuel_cost_per_liter_default: 1.5,
      lube_cost_per_qty_default: 4.0,
      labor_cost_per_hour_default: 35.0,
      downtime_cost_per_hour_default: 120.0,
    };
    for (const r of rows) {
      const k = String(r.key || "").trim();
      const v = Number(r.value);
      if (k && Number.isFinite(v)) d[k] = v;
    }
    return d;
  }

  // =========================
  // WORK ORDER PDF (manual form + sign-off)
  // =========================
  app.get("/workorder/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    const download = String(req.query?.download || "").trim() === "1";
    const minimal = String(req.query?.minimal || "").trim() === "1";
    const nohf = String(req.query?.nohf || "").trim() === "1";
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).type("text/plain; charset=utf-8").send("valid work order id required");
    }
    try {

    const woCols = db.prepare(`
      PRAGMA table_info(work_orders)
    `).all();
    const woColSet = new Set(woCols.map((c) => String(c.name || "")));
    const optionalWoCol = (name) => (woColSet.has(name) ? `w.${name}` : `NULL AS ${name}`);
    const requiredWoCols = ["id", "asset_id"];
    for (const col of requiredWoCols) {
      if (!woColSet.has(col)) {
        return reply.code(500).type("text/plain; charset=utf-8").send(`work_orders schema missing required column: ${col}`);
      }
    }

    const wo = db.prepare(`
      SELECT
        w.id,
        w.asset_id,
        ${optionalWoCol("source")},
        ${optionalWoCol("reference_id")},
        ${optionalWoCol("status")},
        ${optionalWoCol("opened_at")},
        ${optionalWoCol("closed_at")},
        ${optionalWoCol("completed_at")},
        ${optionalWoCol("completion_notes")},
        ${optionalWoCol("artisan_name")},
        ${optionalWoCol("artisan_signed_at")},
        ${optionalWoCol("supervisor_name")},
        ${optionalWoCol("supervisor_signed_at")},
        a.asset_code,
        a.asset_name,
        a.category
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE w.id = ?
    `).get(id);

    if (!wo) return reply.code(404).type("text/plain; charset=utf-8").send("work order not found");

    const breakdownTableExists = Boolean(
      db.prepare(`
        SELECT 1
        FROM sqlite_master
        WHERE type = 'table' AND name = 'breakdowns'
        LIMIT 1
      `).get()
    );
    const maintenancePlansTableExists = Boolean(
      db.prepare(`
        SELECT 1
        FROM sqlite_master
        WHERE type = 'table' AND name = 'maintenance_plans'
        LIMIT 1
      `).get()
    );

    let breakdown = null;
    if (breakdownTableExists && String(wo.source || "").toLowerCase() === "breakdown" && wo.reference_id) {
      breakdown = db.prepare(`
        SELECT id, breakdown_date, description, component, critical, downtime_total_hours
        FROM breakdowns
        WHERE id = ?
      `).get(wo.reference_id);
    }

    let servicePlan = null;
    if (maintenancePlansTableExists && String(wo.source || "").toLowerCase() === "service" && wo.reference_id) {
      servicePlan = db.prepare(`
        SELECT id, service_name, interval_hours, last_service_hours, active
        FROM maintenance_plans
        WHERE id = ?
      `).get(wo.reference_id);
    }

    let issuedParts = [];
    const stockMovementsExists = Boolean(
      db.prepare(`
        SELECT 1
        FROM sqlite_master
        WHERE type = 'table' AND name = 'stock_movements'
        LIMIT 1
      `).get()
    );
    const partsExists = Boolean(
      db.prepare(`
        SELECT 1
        FROM sqlite_master
        WHERE type = 'table' AND name = 'parts'
        LIMIT 1
      `).get()
    );
    if (stockMovementsExists && partsExists) {
      const stockMovementCols = db.prepare(`
        PRAGMA table_info(stock_movements)
      `).all();
      const partCols = db.prepare(`
        PRAGMA table_info(parts)
      `).all();
      const stockColSet = new Set(stockMovementCols.map((c) => String(c.name || "")));
      const partColSet = new Set(partCols.map((c) => String(c.name || "")));
      const canQueryIssuedParts =
        stockColSet.has("id") &&
        stockColSet.has("part_id") &&
        stockColSet.has("quantity") &&
        stockColSet.has("reference") &&
        partColSet.has("id") &&
        partColSet.has("part_code") &&
        partColSet.has("part_name");
      if (!canQueryIssuedParts) {
        issuedParts = [];
      } else {
      const hasCreatedAt = stockMovementCols.some((c) => String(c.name) === "created_at");
      const hasMovementDate = stockMovementCols.some((c) => String(c.name) === "movement_date");
      const movementDateExpr = hasCreatedAt
        ? "sm.created_at"
        : hasMovementDate
          ? "sm.movement_date"
          : "datetime('now')";

      issuedParts = db.prepare(`
        SELECT
          sm.id,
          ${movementDateExpr} AS movement_date,
          sm.quantity,
          p.part_code,
          p.part_name
        FROM stock_movements sm
        JOIN parts p ON p.id = sm.part_id
        WHERE sm.reference = ?
        ORDER BY sm.id ASC
      `).all(`work_order:${id}`);
      }
    }

    const logoPath = path.join(process.cwd(), "branding", "logo.png");

    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Work Order");
        kvGrid(doc, [
          { k: "WO #", v: wo.id },
          { k: "Status", v: String(wo.status || "").toUpperCase() },
          { k: "Source", v: String(wo.source || "").toUpperCase() },
          { k: "Reference ID", v: wo.reference_id ?? "-" },
          { k: "Asset Code", v: wo.asset_code || "" },
          { k: "Asset Name", v: wo.asset_name || "" },
          { k: "Category", v: wo.category || "" },
          { k: "Opened At", v: wo.opened_at || "" },
          { k: "Closed At", v: wo.closed_at || "-" },
        ], 2);

        if (minimal) return;

        if (breakdown) {
          sectionTitle(doc, "Linked Breakdown");
          kvGrid(doc, [
            { k: "Breakdown #", v: breakdown.id },
            { k: "Date", v: breakdown.breakdown_date || "" },
            { k: "Component", v: breakdown.component || "General" },
            { k: "Critical", v: yn(breakdown.critical) },
            { k: "Downtime (hrs)", v: fmtNum(breakdown.downtime_total_hours, 1) },
            { k: "Description", v: breakdown.description || "" },
          ], 2);
        }

        if (servicePlan) {
          sectionTitle(doc, "Linked Service Plan");
          kvGrid(doc, [
            { k: "Plan #", v: servicePlan.id },
            { k: "Service", v: servicePlan.service_name || "" },
            { k: "Interval (hrs)", v: fmtNum(servicePlan.interval_hours, 1) },
            { k: "Last Service (hrs)", v: fmtNum(servicePlan.last_service_hours, 1) },
            { k: "Active", v: yn(servicePlan.active) },
          ], 2);
        }

        sectionTitle(doc, "Issued Parts");
        table(
          doc,
          [
            { key: "date", label: "Date", width: 0.24 },
            { key: "part_code", label: "Part Code", width: 0.16 },
            { key: "part_name", label: "Part Name", width: 0.42 },
            { key: "qty", label: "Qty", width: 0.18, align: "right" },
          ],
          issuedParts.length
            ? issuedParts.map((p) => ({
                date: p.movement_date || "",
                part_code: p.part_code || "",
                part_name: p.part_name || "",
                qty: fmtNum(Math.abs(Number(p.quantity || 0)), 0),
              }))
            : [{ date: "-", part_code: "-", part_name: "No parts issued", qty: "-" }]
        );

        sectionTitle(doc, "Parts & Lubes Required (Manual)");
        doc
          .font("Helvetica")
          .fontSize(9)
          .fillColor("#333333")
          .text(
            "Use this section to request parts/lubes from Stores for the service/repair, even if nothing has been issued yet.",
            { width: 520 }
          );
        doc.moveDown(0.4);

        // Draw a fixed-height manual grid (avoids table auto page-break quirks)
        const left2 = doc.page.margins.left;
        const right2 = doc.page.width - doc.page.margins.right;
        const w2 = right2 - left2;
        const headerH2 = 18;
        const rowH2 = 18;
        const rows2 = 8;
        const gridH2 = headerH2 + rows2 * rowH2;

        ensurePageSpace(doc, gridH2 + 14);

        const yTop = doc.y;
        const cols = [
          { label: "Type", w: 0.12 },
          { label: "Code", w: 0.18 },
          { label: "Description", w: 0.34 },
          { label: "Qty", w: 0.10, align: "right" },
          { label: "Issued By", w: 0.14 },
          { label: "Date", w: 0.12 },
        ];
        const abs = cols.map((c) => Math.floor(c.w * w2));
        const used2 = abs.slice(0, -1).reduce((a, b) => a + b, 0);
        abs[abs.length - 1] = Math.max(60, w2 - used2);

        // Header background
        doc.save();
        doc.rect(left2, yTop, w2, headerH2).fillOpacity(0.06).fill("#000000");
        doc.restore();

        // Outer border
        doc.save();
        doc.rect(left2, yTop, w2, gridH2).lineWidth(1).strokeOpacity(0.18).stroke("#000000");
        doc.restore();

        // Vertical lines + header labels
        let x2 = left2;
        doc.font("Helvetica-Bold").fontSize(9).fillColor("#111111");
        for (let i = 0; i < cols.length; i++) {
          const cw = abs[i];
          const padX = 4;
          const align = cols[i].align || "left";
          doc.text(cols[i].label, x2 + padX, yTop + 5, { width: cw - padX * 2, align });
          if (i > 0) {
            doc
              .moveTo(x2, yTop)
              .lineTo(x2, yTop + gridH2)
              .lineWidth(1)
              .strokeOpacity(0.12)
              .stroke("#000000");
          }
          x2 += cw;
        }

        // Horizontal row lines
        for (let r = 0; r <= rows2; r++) {
          const yy = yTop + headerH2 + r * rowH2;
          doc
            .moveTo(left2, yy)
            .lineTo(right2, yy)
            .lineWidth(1)
            .strokeOpacity(0.08)
            .stroke("#000000");
        }

        doc.y = yTop + gridH2 + 10;

        sectionTitle(doc, "Recorded Completion");
        kvGrid(doc, [
          { k: "Completed At", v: wo.completed_at || wo.closed_at || "-" },
          { k: "Artisan", v: wo.artisan_name || "-" },
          { k: "Artisan Signed At", v: wo.artisan_signed_at || "-" },
          { k: "Supervisor", v: wo.supervisor_name || "-" },
          { k: "Supervisor Signed At", v: wo.supervisor_signed_at || "-" },
          {
            k: "Completion Notes",
            v: wo.completion_notes
              ? compactCell(String(wo.completion_notes || "").replace(/\s+/g, " ").trim(), 220)
              : "-",
          },
        ], 2);

        sectionTitle(doc, "Manual Work Execution (Artisan)");
        doc.font("Helvetica").fontSize(10);
        doc.text("Job Description / Findings:", { width: 460 });
        doc.moveDown(0.2);
        for (let i = 0; i < 4; i++) {
          const y = doc.y + 12;
          doc.moveTo(doc.page.margins.left, y).lineTo(doc.page.width - doc.page.margins.right, y).strokeOpacity(0.2).stroke("#000000");
          doc.moveDown(0.8);
        }

        doc.moveDown(0.4);
        doc.text("Actions Performed:", { width: 460 });
        doc.moveDown(0.2);
        for (let i = 0; i < 4; i++) {
          const y = doc.y + 12;
          doc.moveTo(doc.page.margins.left, y).lineTo(doc.page.width - doc.page.margins.right, y).strokeOpacity(0.2).stroke("#000000");
          doc.moveDown(0.8);
        }

        sectionTitle(doc, "Sign-Off");
        const left = doc.page.margins.left;
        const right = doc.page.width - doc.page.margins.right;
        const mid = left + (right - left) / 2;
        const y0 = doc.y + 6;

        doc.font("Helvetica-Bold").fontSize(10);
        doc.text("Artisan Signature", left, y0, { width: (right - left) / 2 - 20 });
        doc.text("Supervisor Signature", mid + 20, y0, { width: (right - left) / 2 - 20 });

        const lineY = y0 + 30;
        doc.moveTo(left, lineY).lineTo(mid - 20, lineY).strokeOpacity(0.35).stroke("#000000");
        doc.moveTo(mid + 20, lineY).lineTo(right, lineY).strokeOpacity(0.35).stroke("#000000");

        doc.font("Helvetica").fontSize(9);
        doc.text("Name:", left, lineY + 8);
        doc.text("Date:", left + 150, lineY + 8);
        doc.text("Name:", mid + 20, lineY + 8);
        doc.text("Date:", mid + 170, lineY + 8);
      },
      {
        title: "IRONLOG",
        subtitle: "Work Order Job Card",
        rightText: `WO #${wo.id}`,
        showPageNumbers: true,
        disableHeaderFooter: nohf,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Work_Order_${wo.id}_${todayYmd()}.pdf"`
      )
      .send(pdf);
    } catch (err) {
      req.log.error({ err, id }, "workorder pdf generation failed");
      try {
        const fallbackPdf = await buildPdfBuffer(
          (doc) => {
            sectionTitle(doc, "Work Order");
            kvGrid(doc, [
              { k: "WO #", v: id },
              { k: "Status", v: "PDF fallback generated" },
              { k: "Error", v: "Detailed PDF template failed. Please contact support." },
            ], 1);
          },
          {
            title: "IRONLOG",
            subtitle: "Work Order Job Card (Fallback)",
            rightText: `WO #${id}`,
            showPageNumbers: true,
            disableHeaderFooter: nohf,
          }
        );
        return reply
          .header("Content-Type", "application/pdf")
          .header(
            "Content-Disposition",
            `${download ? "attachment" : "inline"}; filename="AML_Work_Order_${id}_${todayYmd()}_fallback.pdf"`
          )
          .send(fallbackPdf);
      } catch (fallbackErr) {
        req.log.error({ fallbackErr, id }, "workorder fallback pdf generation failed");
        return reply
          .code(500)
          .type("text/plain; charset=utf-8")
          .send(`workorder_pdf_generation_failed:${id}`);
      }
    }
  });

  // =========================
  // ASSET HISTORY PDF
  // =========================
  // GET /api/reports/asset-history/:asset_code.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&download=1
  app.get("/asset-history/:asset_code.pdf", async (req, reply) => {
    const asset_code = String(req.params?.asset_code || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const download = String(req.query?.download || "").trim() === "1";

    if (!asset_code) return reply.code(400).send({ error: "asset_code is required" });

    const asset = db.prepare(`
      SELECT id, asset_code, asset_name, category
      FROM assets
      WHERE asset_code = ?
    `).get(asset_code);
    if (!asset) return reply.code(404).send({ error: "asset not found" });

    const startOk = start && isDate(start);
    const endOk = end && isDate(end);
    const periodText = `${startOk ? start : "beginning"} to ${endOk ? end : "today"}`;

    const dateFilter = (col) => {
      const clauses = [];
      const params = [];
      if (startOk) { clauses.push(`${col} >= ?`); params.push(start); }
      if (endOk) { clauses.push(`${col} <= ?`); params.push(end); }
      return { sql: clauses.length ? ` AND ${clauses.join(" AND ")}` : "", params };
    };

    const hasTable = (name) => {
      const row = db.prepare(`
        SELECT name
        FROM sqlite_master
        WHERE type = 'table' AND name = ?
      `).get(name);
      return Boolean(row);
    };

    const bdF = dateFilter("b.breakdown_date");
    const breakdowns = db.prepare(`
      SELECT
        b.breakdown_date AS date,
        b.description,
        b.component,
        b.critical,
        COALESCE(b.downtime_total_hours, 0) AS downtime_hours
      FROM breakdowns b
      WHERE b.asset_id = ? ${bdF.sql}
      ORDER BY b.breakdown_date DESC, b.id DESC
      LIMIT 300
    `).all(asset.id, ...bdF.params).map((r) => ({
      date: r.date,
      description: r.description || "",
      component: r.component || "General",
      critical: Boolean(r.critical),
      downtime_hours: Number(r.downtime_hours || 0),
    }));

    const woF = dateFilter("DATE(w.opened_at)");
    const workOrders = db.prepare(`
      SELECT
        w.id,
        DATE(w.opened_at) AS date,
        w.source,
        w.status,
        w.opened_at,
        w.closed_at
      FROM work_orders w
      WHERE w.asset_id = ?
        AND w.closed_at IS NULL
        AND REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
        ${woF.sql}
      ORDER BY w.id DESC
      LIMIT 300
    `).all(asset.id, ...woF.params);

    const getF = dateFilter("g.slip_date");
    const getSlips = hasTable("get_change_slips")
      ? db.prepare(`
          SELECT
            g.id,
            g.slip_date AS date,
            g.location,
            g.notes
          FROM get_change_slips g
          WHERE g.asset_id = ? ${getF.sql}
          ORDER BY g.slip_date DESC, g.id DESC
          LIMIT 300
        `).all(asset.id, ...getF.params)
      : [];

    const compF = dateFilter("c.slip_date");
    const componentSlips = hasTable("component_change_slips")
      ? db.prepare(`
          SELECT
            c.id,
            c.slip_date AS date,
            c.component,
            c.serial_out,
            c.serial_in,
            c.hours_at_change
          FROM component_change_slips c
          WHERE c.asset_id = ? ${compF.sql}
          ORDER BY c.slip_date DESC, c.id DESC
          LIMIT 300
        `).all(asset.id, ...compF.params)
      : [];

    const siteCode = String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
    const opsF = dateFilter("r.report_date");
    const opsSlips = hasTable("ops_slip_reports")
      ? db.prepare(`
          SELECT r.id, r.slip_type, r.report_date AS date, r.created_by
          FROM ops_slip_reports r
          WHERE r.asset_id = ? AND r.site_code = ? ${opsF.sql}
          ORDER BY r.report_date DESC, r.id DESC
          LIMIT 300
        `).all(asset.id, siteCode, ...opsF.params)
      : [];

    const breakdownsPdf = breakdowns.slice(0, 60);
    const workOrdersPdf = workOrders.slice(0, 60);
    const getSlipsPdf = getSlips.slice(0, 60);
    const componentSlipsPdf = componentSlips.slice(0, 60);
    const opsSlipsPdf = opsSlips.slice(0, 60);

    const oilF = dateFilter("o.log_date");
    const oil = db.prepare(`
      SELECT IFNULL(SUM(o.quantity), 0) AS oil_qty
      FROM oil_logs o
      WHERE o.asset_id = ? ${oilF.sql}
    `).get(asset.id, ...oilF.params);

    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateCol = hasCreatedAt ? "DATE(sm.created_at)" : "DATE('now')";
    const smF = dateFilter(smDateCol);
    const partsUsed = db.prepare(`
      SELECT IFNULL(SUM(ABS(sm.quantity)), 0) AS qty
      FROM stock_movements sm
      JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
      WHERE w.asset_id = ?
        AND sm.movement_type = 'out'
        ${smF.sql}
    `).get(asset.id, ...smF.params);

    const defaults = costDefaults();
    const fuelF = dateFilter("fl.log_date");
    const fuelCost = db.prepare(`
      SELECT COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS value
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.asset_id = ? ${fuelF.sql}
    `).get(defaults.fuel_cost_per_liter_default, asset.id, ...fuelF.params);

    const oilCost = db.prepare(`
      SELECT COALESCE(SUM(o.quantity * COALESCE(o.unit_cost, ?)), 0) AS value
      FROM oil_logs o
      WHERE o.asset_id = ? ${oilF.sql}
    `).get(defaults.lube_cost_per_qty_default, asset.id, ...oilF.params);

    const partsCost = db.prepare(`
      SELECT COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS value
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
      WHERE w.asset_id = ?
        AND sm.movement_type = 'out'
        ${smF.sql}
    `).get(asset.id, ...smF.params);

    const woCost = db.prepare(`
      SELECT
        COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS labor_hours,
        COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      WHERE w.asset_id = ?
        AND DATE(COALESCE(w.completed_at, w.closed_at, w.opened_at))
            ${startOk ? ">= ?" : ">= DATE('1900-01-01')"}
        AND DATE(COALESCE(w.completed_at, w.closed_at, w.opened_at))
            ${endOk ? "<= ?" : "<= DATE('now')"}
    `).get(
      defaults.labor_cost_per_hour_default,
      asset.id,
      ...(startOk ? [start] : []),
      ...(endOk ? [end] : [])
    );

    const downtimeCost = db.prepare(`
      SELECT COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS value
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE b.asset_id = ?
        AND l.log_date ${startOk ? ">= ?" : ">= DATE('1900-01-01')"}
        AND l.log_date ${endOk ? "<= ?" : "<= DATE('now')"}
    `).get(
      defaults.downtime_cost_per_hour_default,
      asset.id,
      ...(startOk ? [start] : []),
      ...(endOk ? [end] : [])
    );

    const totalCost = Number(
      (
        Number(fuelCost?.value || 0) +
        Number(oilCost?.value || 0) +
        Number(partsCost?.value || 0) +
        Number(woCost?.labor_cost || 0) +
        Number(downtimeCost?.value || 0)
      ).toFixed(2)
    );

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Asset");
        kvGrid(doc, [
          { k: "Asset Code", v: asset.asset_code || "" },
          { k: "Asset Name", v: asset.asset_name || "" },
          { k: "Category", v: asset.category || "" },
          { k: "Period", v: periodText },
        ], 2);

        sectionTitle(doc, "Period Summary");
        kvGrid(doc, [
          { k: "Breakdowns", v: breakdowns.length },
          { k: "Work Orders", v: workOrders.length },
          { k: "GET Slips", v: getSlips.length },
          { k: "Component Slips", v: componentSlips.length },
          { k: "Ops slips (PDF)", v: opsSlips.length },
          { k: "Total Downtime (hrs)", v: fmtNum(breakdowns.reduce((a, r) => a + Number(r.downtime_hours || 0), 0), 1) },
          { k: "Parts Used (qty)", v: fmtNum(partsUsed?.qty || 0, 0) },
          { k: "Oil Used (qty)", v: fmtNum(oil?.oil_qty || 0, 1) },
        ], 2);

        sectionTitle(doc, "Cost Summary");
        kvGrid(doc, [
          { k: "Fuel Cost", v: fmtNum(fuelCost?.value || 0, 2) },
          { k: "Oil/Lube Cost", v: fmtNum(oilCost?.value || 0, 2) },
          { k: "Parts Cost", v: fmtNum(partsCost?.value || 0, 2) },
          { k: "Labor Cost", v: fmtNum(woCost?.labor_cost || 0, 2) },
          { k: "Labor Hours", v: fmtNum(woCost?.labor_hours || 0, 1) },
          { k: "Downtime Cost", v: fmtNum(downtimeCost?.value || 0, 2) },
          { k: "Total Maintenance Cost", v: fmtNum(totalCost, 2) },
        ], 2);

        sectionTitle(doc, "Breakdowns");
        table(
          doc,
          [
            { key: "date", label: "Date", width: 0.16 },
            { key: "component", label: "Component", width: 0.18 },
            { key: "downtime", label: "Downtime", width: 0.12, align: "right" },
            { key: "critical", label: "Critical", width: 0.10, align: "center" },
            { key: "description", label: "Description", width: 0.44 },
          ],
          breakdownsPdf.length
            ? breakdownsPdf.map((r) => ({
                date: r.date,
                component: r.component,
                downtime: fmtNum(r.downtime_hours, 1),
                critical: r.critical ? "YES" : "NO",
                description: compactCell(r.description, 180),
              }))
            : [{ date: "-", component: "-", downtime: "-", critical: "-", description: "No breakdowns in range" }]
        );

        sectionTitle(doc, "Work Orders");
        table(
          doc,
          [
            { key: "id", label: "WO#", width: 0.12, align: "right" },
            { key: "date", label: "Date", width: 0.16 },
            { key: "source", label: "Source", width: 0.16 },
            { key: "status", label: "Status", width: 0.16 },
            { key: "opened", label: "Opened", width: 0.40 },
          ],
          workOrdersPdf.length
            ? workOrdersPdf.map((r) => ({
                id: String(r.id),
                date: r.date || "",
                source: String(r.source || ""),
                status: String(r.status || ""),
                opened: r.opened_at || "",
              }))
            : [{ id: "-", date: "-", source: "-", status: "-", opened: "-" }]
        );

        sectionTitle(doc, "GET Change Slips");
        table(
          doc,
          [
            { key: "id", label: "Slip#", width: 0.14, align: "right" },
            { key: "date", label: "Date", width: 0.18 },
            { key: "location", label: "Location", width: 0.28 },
            { key: "notes", label: "Notes", width: 0.40 },
          ],
          getSlipsPdf.length
            ? getSlipsPdf.map((r) => ({
                id: String(r.id),
                date: r.date || "",
                location: r.location || "-",
                notes: compactCell(r.notes, 140),
              }))
            : [{ id: "-", date: "-", location: "-", notes: "No GET slips in range" }]
        );

        sectionTitle(doc, "Component Change Slips");
        table(
          doc,
          [
            { key: "id", label: "Slip#", width: 0.12, align: "right" },
            { key: "date", label: "Date", width: 0.16 },
            { key: "component", label: "Component", width: 0.24 },
            { key: "serial", label: "Serial Out -> In", width: 0.34 },
            { key: "hours", label: "Hours", width: 0.14, align: "right" },
          ],
          componentSlipsPdf.length
            ? componentSlipsPdf.map((r) => ({
                id: String(r.id),
                date: r.date || "",
                component: r.component || "",
                serial: `${r.serial_out || "-"} -> ${r.serial_in || "-"}`,
                hours: r.hours_at_change == null ? "-" : fmtNum(r.hours_at_change, 1),
              }))
            : [{ id: "-", date: "-", component: "-", serial: "-", hours: "-" }]
        );

        sectionTitle(doc, "Breakdown Ops slips (saved PDF reports)");
        table(
          doc,
          [
            { key: "id", label: "Slip#", width: 0.12, align: "right" },
            { key: "date", label: "Report date", width: 0.16 },
            { key: "stype", label: "Type", width: 0.22 },
            { key: "by", label: "Recorded by", width: 0.18 },
            { key: "note", label: "Note", width: 0.32 },
          ],
          opsSlipsPdf.length
            ? opsSlipsPdf.map((r) => ({
                id: String(r.id),
                date: r.date || "",
                stype: String(r.slip_type || "").replace(/_/g, " "),
                by: compactCell(r.created_by || "-", 40),
                note: "See IRONLOG Breakdown Ops → open PDF for full detail",
              }))
            : [{ id: "-", date: "-", stype: "-", by: "-", note: "No ops slips in range" }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Asset History Report",
        rightText: `${asset.asset_code} | ${periodText}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Asset_History_${asset.asset_code}_${end}.pdf"`
      )
      .send(pdf);
  });

  // =========================
  // LUBE USAGE PDF
  // =========================
  // GET /api/reports/lube.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&download=1
  app.get("/lube.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const download = String(req.query?.download || "").trim() === "1";
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const defaultLubeCost = Number(
      db.prepare(`SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1`).get()?.value
    );
    const lubeUnitFallback = Number.isFinite(defaultLubeCost) && defaultLubeCost > 0 ? defaultLubeCost : 4.0;

    const rows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS total_lube_cost,
        COUNT(*) AS entries
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.id
      ORDER BY qty_total DESC, a.asset_code ASC
      LIMIT 400
    `).all(lubeUnitFallback, start, end).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      qty_total: Number(r.qty_total || 0),
      total_lube_cost: Number(r.total_lube_cost || 0),
      entries: Number(r.entries || 0),
    }));

    const detailRows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        CASE
          WHEN LOWER(TRIM(COALESCE(ol.oil_type, ''))) IN ('admin','supervisor','manager','stores','artisan','operator') THEN 'UNSPECIFIED'
          ELSE COALESCE(NULLIF(TRIM(ol.oil_type), ''), 'UNSPECIFIED')
        END AS part_number,
        COALESCE(p.part_name, '') AS lube_description,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS total_lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      LEFT JOIN parts p ON UPPER(TRIM(p.part_code)) = UPPER(TRIM(COALESCE(ol.oil_type, '')))
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.id, part_number, p.part_name
      ORDER BY a.asset_code ASC, qty_total DESC, part_number ASC
      LIMIT 1200
    `).all(lubeUnitFallback, start, end).map((r) => ({
      asset_code: String(r.asset_code || ""),
      asset_name: String(r.asset_name || ""),
      part_number: String(r.part_number || "UNSPECIFIED"),
      lube_description: String(r.lube_description || ""),
      qty_total: Number(r.qty_total || 0),
      total_lube_cost: Number(r.total_lube_cost || 0),
    }));

    const summary = db.prepare(`
      SELECT
        COALESCE(SUM(quantity), 0) AS qty_total,
        COALESCE(SUM(quantity * COALESCE(unit_cost, ?)), 0) AS total_lube_cost,
        COUNT(*) AS entries,
        COUNT(DISTINCT asset_id) AS assets
      FROM oil_logs
      WHERE log_date BETWEEN ? AND ?
    `).get(lubeUnitFallback, start, end);

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Lube Usage Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Total Qty", v: fmtNum(summary?.qty_total || 0, 1) },
          { k: "Total Cost", v: fmtNum(summary?.total_lube_cost || 0, 2) },
          { k: "Total Entries", v: fmtNum(summary?.entries || 0, 0) },
          { k: "Assets Logged", v: fmtNum(summary?.assets || 0, 0) },
        ], 2);

        sectionTitle(doc, "Lube Usage by Machine");
        table(
          doc,
          [
            { key: "asset_code", label: "Asset Code", width: 0.18 },
            { key: "asset_name", label: "Asset Name", width: 0.36 },
            { key: "entries", label: "Entries", width: 0.12, align: "right" },
            { key: "qty_total", label: "Qty Total", width: 0.16, align: "right" },
            { key: "total_lube_cost", label: "Total Cost", width: 0.18, align: "right" },
          ],
          rows.length
            ? rows.map((r) => ({
                asset_code: r.asset_code,
                asset_name: r.asset_name || "",
                entries: fmtNum(r.entries, 0),
                qty_total: fmtNum(r.qty_total, 1),
                total_lube_cost: fmtNum(r.total_lube_cost, 2),
              }))
            : [{ asset_code: "-", asset_name: "No lube usage in period", entries: "-", qty_total: "-", total_lube_cost: "-" }]
        );

        sectionTitle(doc, "Lube Usage Detail (Part Number / Description)");
        table(
          doc,
          [
            { key: "asset_code", label: "Asset Code", width: 0.12 },
            { key: "asset_name", label: "Asset Name", width: 0.24 },
            { key: "part_number", label: "Lube Part Number", width: 0.16 },
            { key: "lube_description", label: "Lube Description", width: 0.24 },
            { key: "qty_total", label: "Qty", width: 0.10, align: "right" },
            { key: "total_lube_cost", label: "Cost", width: 0.14, align: "right" },
          ],
          detailRows.length
            ? detailRows.map((r) => ({
                asset_code: r.asset_code,
                asset_name: r.asset_name || "",
                part_number: r.part_number || "UNSPECIFIED",
                lube_description: r.lube_description || "-",
                qty_total: fmtNum(r.qty_total, 1),
                total_lube_cost: fmtNum(r.total_lube_cost, 2),
              }))
            : [{
                asset_code: "-",
                asset_name: "-",
                part_number: "-",
                lube_description: "No lube detail rows in period",
                qty_total: "-",
                total_lube_cost: "-",
              }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Lube Usage Report",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
        layout: "landscape",
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Lube_${end}.pdf"`
      )
      .send(pdf);
  });

  // =========================
  // FUEL BENCHMARK PDF
  // =========================
  // GET /api/reports/fuel-benchmark.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15&download=1
  app.get("/fuel-benchmark.pdf", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const toleranceInput = Number(req.query?.tolerance ?? 0.15);
    const tolerance = Number.isFinite(toleranceInput) ? Math.max(0, toleranceInput) : 0.15;
    const download = String(req.query?.download || "").trim() === "1";
    const modeFilter = String(req.query?.mode || "").trim().toLowerCase(); // 'km' | 'hours' | ''
    const assetFilter = String(req.query?.asset_code || "").trim().toLowerCase();
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const fuelByAsset = db.prepare(`
      SELECT
        a.id AS asset_id,
        a.asset_code,
        a.asset_name,
        a.category,
        CASE
          WHEN UPPER(COALESCE(a.asset_code, '')) GLOB 'V[0-9][0-9]AM' THEN 'km'
          ELSE COALESCE(NULLIF(TRIM(a.utilization_mode), ''), CASE
          WHEN LOWER(COALESCE(a.category, '')) LIKE '%truck%'
            OR LOWER(COALESCE(a.category, '')) LIKE '%vehicle%'
            OR LOWER(COALESCE(a.category, '')) LIKE '%ldv%'
            OR LOWER(COALESCE(a.category, '')) LIKE '%pickup%'
            OR LOWER(COALESCE(a.category, '')) LIKE '%bakkie%'
            OR LOWER(COALESCE(a.asset_code, '')) LIKE 'ldv%'
            OR UPPER(COALESCE(a.asset_code, '')) GLOB 'V[0-9][0-9]AM'
            OR LOWER(COALESCE(a.asset_name, '')) LIKE '%ldv%'
            THEN 'km'
          ELSE 'hours'
        END)
        END AS metric_mode,
        COALESCE(NULLIF(a.km_per_hour_factor, 0), 10.0) AS km_per_hour_factor,
        COALESCE(a.baseline_fuel_l_per_hour, 5.0) AS oem_lph,
        COALESCE(a.baseline_fuel_km_per_l, 2.0) AS oem_kmpl,
        COALESCE(SUM(fl.liters), 0) AS fuel_liters,
        COUNT(fl.id) AS fill_count
      FROM assets a
      LEFT JOIN fuel_logs fl
        ON fl.asset_id = a.id
       AND fl.log_date BETWEEN ? AND ?
      WHERE a.active = 1
        AND UPPER(COALESCE(a.asset_code, '')) NOT GLOB 'V[0-9][0-9]AM'
      GROUP BY a.id
      ORDER BY a.asset_code ASC
    `).all(start, end);

    const getFuelLogsInRange = db.prepare(`
      SELECT
        id,
        log_date,
        COALESCE(LOWER(meter_unit), '') AS meter_unit,
        COALESCE(meter_run_value, 0) AS meter_run_value,
        COALESCE(hours_run, 0) AS hours_run
      FROM fuel_logs
      WHERE asset_id = ?
        AND log_date BETWEEN ? AND ?
      ORDER BY log_date ASC, id ASC
    `);
    const getFuelLogBeforeRange = db.prepare(`
      SELECT
        id,
        log_date,
        COALESCE(LOWER(meter_unit), '') AS meter_unit,
        COALESCE(meter_run_value, 0) AS meter_run_value,
        COALESCE(hours_run, 0) AS hours_run
      FROM fuel_logs
      WHERE asset_id = ?
        AND log_date < ?
        AND (
          COALESCE(meter_run_value, 0) > 0
          OR COALESCE(hours_run, 0) > 0
        )
      ORDER BY log_date DESC, id DESC
      LIMIT 1
    `);

    function getRunFromFuel(assetId, startDate, endDate) {
      const logs = getFuelLogsInRange.all(assetId, startDate, endDate);
      if (!logs.length) return { km_run: 0, hours_run: 0 };

      const prev = getFuelLogBeforeRange.get(assetId, startDate);
      let prevKmMeter = null;
      let prevHoursMeter = null;
      if (prev) {
        const prevUnit = String(prev.meter_unit || "").toLowerCase();
        const prevMeter = Number(prev.meter_run_value || 0);
        if (prevUnit === "km" && prevMeter > 0) prevKmMeter = prevMeter;
        if (prevUnit === "hours" && prevMeter > 0) prevHoursMeter = prevMeter;
      }

      let km_run = 0;
      let hours_run = 0;
      for (const row of logs) {
        const unit = String(row.meter_unit || "").toLowerCase();
        const meter = Number(row.meter_run_value || 0);
        const legacyHours = Number(row.hours_run || 0);

        if (unit === "km" && meter > 0) {
          if (prevKmMeter != null) {
            const delta = meter - prevKmMeter;
            if (Number.isFinite(delta) && delta > 0) km_run += delta;
          }
          prevKmMeter = meter;
          continue;
        }
        if (unit === "hours" && meter > 0) {
          if (prevHoursMeter != null) {
            const delta = meter - prevHoursMeter;
            if (Number.isFinite(delta) && delta > 0) hours_run += delta;
          }
          prevHoursMeter = meter;
          continue;
        }
        if (legacyHours > 0) hours_run += legacyHours;
      }
      return { km_run, hours_run };
    }

    const isLdvFleetCode = (code) => /^V(0[1-9]|1[0-4])AM$/i.test(String(code || "").trim());
    const benchmarkRows = fuelByAsset.map((r) => {
      const mode = String(r.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
      const fuelRun = getRunFromFuel(r.asset_id, start, end) || {};
      const fuelKm = Number(fuelRun.km_run || 0);
      const fuelHours = Number(fuelRun.hours_run || 0);
      const km = fuelKm > 0 ? fuelKm : 0;
      const hours = fuelHours > 0 ? fuelHours : 0;
      const fuel = Number(r.fuel_liters || 0);
      const oem = Number(r.oem_lph || 5);
      const oemK = Number(r.oem_kmpl || 2);
      const fillCount = Number(r.fill_count || 0);
      const lph = hours > 0 ? fuel / hours : null;
      const kmpl = fuel > 0 && km > 0 ? km / fuel : null;
      const excessiveThreshold = oem * (1 + tolerance);
      const lowThresholdKmpl = oemK * Math.max(0, 1 - tolerance);
      const hasEnoughSamples = fillCount >= 2;
      const is_excessive = hasEnoughSamples && (mode === "km"
        ? (kmpl != null && kmpl < lowThresholdKmpl)
        : (lph != null && lph > excessiveThreshold));
      return {
        metric_mode: mode,
        asset_code: r.asset_code,
        asset_name: r.asset_name,
        fuel_liters: Number(fuel.toFixed(2)),
        km_run: Number(km.toFixed(2)),
        hours_run: Number(hours.toFixed(2)),
        actual_lph: lph == null ? null : Number(lph.toFixed(3)),
        oem_lph: Number(oem.toFixed(3)),
        threshold_lph: Number(excessiveThreshold.toFixed(3)),
        variance_lph: lph == null ? null : Number((lph - oem).toFixed(3)),
        actual_km_per_l: kmpl == null ? null : Number(kmpl.toFixed(3)),
        oem_km_per_l: Number(oemK.toFixed(3)),
        threshold_km_per_l: Number(lowThresholdKmpl.toFixed(3)),
        variance_km_per_l: kmpl == null ? null : Number((kmpl - oemK).toFixed(3)),
        fill_count: fillCount,
        flag: is_excessive ? "EXCESSIVE" : "OK",
      };
    }).filter((r) => r.fuel_liters > 0)
      // Temporary business rule: exclude LDV/km-mode assets from benchmark report.
      .filter((r) => r.metric_mode !== "km")
      // Hard exclusion for LDV fleet codes requested by business.
      .filter((r) => !isLdvFleetCode(r.asset_code))
      .filter((r) => (assetFilter ? String(r.asset_code || "").trim().toLowerCase() === assetFilter : true))
      .filter((r) => (modeFilter === "km" ? r.metric_mode === "km" : modeFilter === "hours" ? r.metric_mode === "hours" : true))
      .sort((a, b) => {
        const ex = (b.flag === "EXCESSIVE" ? 1 : 0) - (a.flag === "EXCESSIVE" ? 1 : 0);
        if (ex !== 0) return ex;
        const av = a.metric_mode === "km" ? Number(a.variance_km_per_l || -999) : Number(a.variance_lph || -999);
        const bv = b.metric_mode === "km" ? Number(b.variance_km_per_l || -999) : Number(b.variance_lph || -999);
        return bv - av;
      });

    const summary = rows.reduce(
      (acc, r) => {
        acc.assets += 1;
        acc.fuel_liters += Number(r.fuel_liters || 0);
        acc.hours_run += Number(r.hours_run || 0);
        acc.km_run += Number(r.km_run || 0);
        if (r.metric_mode === "km") {
          if (Number(r.km_run || 0) > 0) acc.km_fuel += Number(r.fuel_liters || 0);
        } else {
          if (Number(r.hours_run || 0) > 0) acc.hours_fuel += Number(r.fuel_liters || 0);
        }
        if (r.flag === "EXCESSIVE") acc.excessive += 1;
        return acc;
      },
      { assets: 0, fuel_liters: 0, hours_run: 0, km_run: 0, excessive: 0, hours_fuel: 0, km_fuel: 0 }
    );
    summary.fuel_liters = Number(summary.fuel_liters.toFixed(2));
    summary.hours_run = Number(summary.hours_run.toFixed(2));
    summary.km_run = Number(summary.km_run.toFixed(2));
    summary.avg_lph = summary.hours_run > 0
      ? Number((summary.hours_fuel / summary.hours_run).toFixed(3))
      : null;
    summary.avg_km_per_l = summary.km_fuel > 0
      ? Number((summary.km_run / summary.km_fuel).toFixed(3))
      : null;

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Fuel Benchmark Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Tolerance", v: `${fmtNum(tolerance * 100, 1)}%` },
          { k: "Mode filter", v: modeFilter ? modeFilter : "all" },
          { k: "Assets", v: fmtNum(summary.assets, 0) },
          { k: "Excessive", v: fmtNum(summary.excessive, 0) },
          { k: "Fuel Total (L)", v: fmtNum(summary.fuel_liters, 2) },
          { k: "Hours Run", v: fmtNum(summary.hours_run, 2) },
          { k: "Avg L/hr", v: summary.avg_lph == null ? "-" : fmtNum(summary.avg_lph, 3) },
        ], 2);
        doc.font("Helvetica-Bold").fontSize(9).fillColor("#b91c1c").text("Flag legend: EXCESSIVE = above OEM by tolerance", { align: "left" });
        doc.moveDown(0.4);
        doc.font("Helvetica").fontSize(9).fillColor("#111111");

        sectionTitle(doc, "Fuel Benchmark by Machine");
        table(
          doc,
          [
            { key: "asset_code", label: "Asset", width: 0.12 },
            { key: "asset_name", label: "Name", width: 0.24 },
            { key: "metric_mode", label: "Mode", width: 0.08, align: "center" },
            { key: "fuel_liters", label: "Fuel (L)", width: 0.12, align: "right" },
            { key: "run", label: "Run", width: 0.12, align: "right" },
            { key: "actual", label: "Actual", width: 0.10, align: "right" },
            { key: "oem", label: "OEM", width: 0.08, align: "right" },
            { key: "variance", label: "Variance", width: 0.08, align: "right" },
            { key: "flag", label: "Flag", width: 0.08, align: "center" },
          ],
          rows.length
            ? rows.map((r) => ({
                asset_code: r.asset_code,
                asset_name: r.asset_name || "",
                metric_mode: r.metric_mode,
                fuel_liters: fmtNum(r.fuel_liters, 2),
                run: r.metric_mode === "km" ? `${fmtNum(r.km_run, 2)} km` : `${fmtNum(r.hours_run, 2)} h`,
                actual: r.metric_mode === "km"
                  ? (r.actual_km_per_l == null ? "-" : `${fmtNum(r.actual_km_per_l, 3)} km/L`)
                  : (r.actual_lph == null ? "-" : `${fmtNum(r.actual_lph, 3)} L/hr`),
                oem: r.metric_mode === "km" ? `${fmtNum(r.oem_km_per_l, 3)}` : `${fmtNum(r.oem_lph, 3)}`,
                variance: r.metric_mode === "km"
                  ? (r.variance_km_per_l == null ? "-" : fmtNum(r.variance_km_per_l, 3))
                  : (r.variance_lph == null ? "-" : fmtNum(r.variance_lph, 3)),
                flag: r.flag,
              }))
            : [{
                asset_code: "-",
                asset_name: "No fuel benchmark data for period",
                metric_mode: "-",
                fuel_liters: "-",
                run: "-",
                actual: "-",
                oem: "-",
                variance: "-",
                flag: "-",
              }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Fuel Benchmark Report",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
        layout: "landscape",
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Fuel_Benchmark_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/fuel-benchmark.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15
  app.get("/fuel-benchmark.xlsx", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const toleranceInput = Number(req.query?.tolerance ?? 0.15);
    const tolerance = Number.isFinite(toleranceInput) ? Math.max(0, toleranceInput) : 0.15;
    const modeFilter = String(req.query?.mode || "").trim().toLowerCase();
    const assetFilter = String(req.query?.asset_code || "").trim().toLowerCase();
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const fuelByAsset = db.prepare(`
      SELECT
        a.id AS asset_id,
        a.asset_code,
        a.asset_name,
        a.category,
        CASE
          WHEN UPPER(COALESCE(a.asset_code, '')) GLOB 'V[0-9][0-9]AM' THEN 'km'
          ELSE COALESCE(NULLIF(TRIM(a.utilization_mode), ''), CASE
          WHEN LOWER(COALESCE(a.category, '')) LIKE '%truck%'
            OR LOWER(COALESCE(a.category, '')) LIKE '%vehicle%'
            OR LOWER(COALESCE(a.category, '')) LIKE '%ldv%'
            OR LOWER(COALESCE(a.category, '')) LIKE '%pickup%'
            OR LOWER(COALESCE(a.category, '')) LIKE '%bakkie%'
            OR LOWER(COALESCE(a.asset_code, '')) LIKE 'ldv%'
            OR UPPER(COALESCE(a.asset_code, '')) GLOB 'V[0-9][0-9]AM'
            OR LOWER(COALESCE(a.asset_name, '')) LIKE '%ldv%'
            THEN 'km'
          ELSE 'hours'
        END)
        END AS metric_mode,
        COALESCE(NULLIF(a.km_per_hour_factor, 0), 10.0) AS km_per_hour_factor,
        COALESCE(a.baseline_fuel_l_per_hour, 5.0) AS oem_lph,
        COALESCE(a.baseline_fuel_km_per_l, 2.0) AS oem_kmpl,
        COALESCE(SUM(fl.liters), 0) AS fuel_liters,
        COUNT(fl.id) AS fill_count
      FROM assets a
      LEFT JOIN fuel_logs fl
        ON fl.asset_id = a.id
       AND fl.log_date BETWEEN ? AND ?
      WHERE a.active = 1
        AND UPPER(COALESCE(a.asset_code, '')) NOT GLOB 'V[0-9][0-9]AM'
      GROUP BY a.id
      ORDER BY a.asset_code ASC
    `).all(start, end);

    const getFuelLogsInRange = db.prepare(`
      SELECT
        id,
        log_date,
        COALESCE(LOWER(meter_unit), '') AS meter_unit,
        COALESCE(meter_run_value, 0) AS meter_run_value,
        COALESCE(hours_run, 0) AS hours_run
      FROM fuel_logs
      WHERE asset_id = ?
        AND log_date BETWEEN ? AND ?
      ORDER BY log_date ASC, id ASC
    `);
    const getFuelLogBeforeRange = db.prepare(`
      SELECT
        id,
        log_date,
        COALESCE(LOWER(meter_unit), '') AS meter_unit,
        COALESCE(meter_run_value, 0) AS meter_run_value,
        COALESCE(hours_run, 0) AS hours_run
      FROM fuel_logs
      WHERE asset_id = ?
        AND log_date < ?
        AND (
          COALESCE(meter_run_value, 0) > 0
          OR COALESCE(hours_run, 0) > 0
        )
      ORDER BY log_date DESC, id DESC
      LIMIT 1
    `);
    function getRunFromFuel(assetId, startDate, endDate) {
      const logs = getFuelLogsInRange.all(assetId, startDate, endDate);
      if (!logs.length) return { km_run: 0, hours_run: 0 };
      const prev = getFuelLogBeforeRange.get(assetId, startDate);
      let prevKmMeter = null;
      let prevHoursMeter = null;
      if (prev) {
        const prevUnit = String(prev.meter_unit || "").toLowerCase();
        const prevMeter = Number(prev.meter_run_value || 0);
        if (prevUnit === "km" && prevMeter > 0) prevKmMeter = prevMeter;
        if (prevUnit === "hours" && prevMeter > 0) prevHoursMeter = prevMeter;
      }
      let km_run = 0;
      let hours_run = 0;
      for (const row of logs) {
        const unit = String(row.meter_unit || "").toLowerCase();
        const meter = Number(row.meter_run_value || 0);
        const legacyHours = Number(row.hours_run || 0);
        if (unit === "km" && meter > 0) {
          if (prevKmMeter != null) {
            const delta = meter - prevKmMeter;
            if (Number.isFinite(delta) && delta > 0) km_run += delta;
          }
          prevKmMeter = meter;
          continue;
        }
        if (unit === "hours" && meter > 0) {
          if (prevHoursMeter != null) {
            const delta = meter - prevHoursMeter;
            if (Number.isFinite(delta) && delta > 0) hours_run += delta;
          }
          prevHoursMeter = meter;
          continue;
        }
        if (legacyHours > 0) hours_run += legacyHours;
      }
      return { km_run, hours_run };
    }

    const isLdvFleetCode = (code) => /^V(0[1-9]|1[0-4])AM$/i.test(String(code || "").trim());
    const benchmarkRows = fuelByAsset.map((r) => {
      const mode = String(r.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
      const fuelRun = getRunFromFuel(r.asset_id, start, end) || {};
      const fuelKm = Number(fuelRun.km_run || 0);
      const fuelHours = Number(fuelRun.hours_run || 0);
      const km = fuelKm > 0 ? fuelKm : 0;
      const hours = fuelHours > 0 ? fuelHours : 0;
      const fuel = Number(r.fuel_liters || 0);
      const oem = Number(r.oem_lph || 5);
      const oemK = Number(r.oem_kmpl || 2);
      const fillCount = Number(r.fill_count || 0);
      const lph = hours > 0 ? fuel / hours : null;
      const kmpl = fuel > 0 && km > 0 ? km / fuel : null;
      const excessiveThreshold = oem * (1 + tolerance);
      const lowThresholdKmpl = oemK * Math.max(0, 1 - tolerance);
      const hasEnoughSamples = fillCount >= 2;
      const is_excessive = hasEnoughSamples && (mode === "km"
        ? (kmpl != null && kmpl < lowThresholdKmpl)
        : (lph != null && lph > excessiveThreshold));
      return {
        metric_mode: mode,
        asset_code: r.asset_code,
        asset_name: r.asset_name,
        category: r.category || "",
        fuel_liters: Number(fuel.toFixed(2)),
        km_run: Number(km.toFixed(2)),
        hours_run: Number(hours.toFixed(2)),
        actual_lph: lph == null ? null : Number(lph.toFixed(3)),
        oem_lph: Number(oem.toFixed(3)),
        threshold_lph: Number(excessiveThreshold.toFixed(3)),
        variance_lph: lph == null ? null : Number((lph - oem).toFixed(3)),
        actual_km_per_l: kmpl == null ? null : Number(kmpl.toFixed(3)),
        oem_km_per_l: Number(oemK.toFixed(3)),
        threshold_km_per_l: Number(lowThresholdKmpl.toFixed(3)),
        variance_km_per_l: kmpl == null ? null : Number((kmpl - oemK).toFixed(3)),
        fill_count: fillCount,
        flag: is_excessive ? "EXCESSIVE" : "OK",
      };
    });
    const rows = benchmarkRows
      .filter((r) => r.fuel_liters > 0)
      .filter((r) => (assetFilter ? String(r.asset_code || "").trim().toLowerCase() === assetFilter : true))
      .filter((r) => (modeFilter === "km" ? r.metric_mode === "km" : modeFilter === "hours" ? r.metric_mode === "hours" : true))
      .sort((a, b) => {
        const ex = (b.flag === "EXCESSIVE" ? 1 : 0) - (a.flag === "EXCESSIVE" ? 1 : 0);
        if (ex !== 0) return ex;
        const av = a.metric_mode === "km" ? Number(a.variance_km_per_l || -999) : Number(a.variance_lph || -999);
        const bv = b.metric_mode === "km" ? Number(b.variance_km_per_l || -999) : Number(b.variance_lph || -999);
        return bv - av;
      });
    const includedCodes = new Set(rows.map((r) => String(r.asset_code || "").trim().toLowerCase()).filter(Boolean));
    const missingRows = benchmarkRows
      .filter((r) => (assetFilter ? String(r.asset_code || "").trim().toLowerCase() === assetFilter : true))
      .filter((r) => (modeFilter === "km" ? r.metric_mode === "km" : modeFilter === "hours" ? r.metric_mode === "hours" : true))
      .filter((r) => !includedCodes.has(String(r.asset_code || "").trim().toLowerCase()))
      .map((r) => {
        let reason = "Excluded by report filters";
        if (Number(r.fuel_liters || 0) <= 0) reason = "No fuel logs in selected period";
        return {
          asset_code: r.asset_code,
          asset_name: r.asset_name,
          category: r.category || "",
          metric_mode: r.metric_mode,
          fuel_liters: Number(r.fuel_liters || 0),
          fill_count: Number(r.fill_count || 0),
          reason,
        };
      })
      .sort((a, b) => String(a.asset_code || "").localeCompare(String(b.asset_code || "")));

    const summary = rows.reduce(
      (acc, r) => {
        acc.assets += 1;
        acc.fuel_liters += Number(r.fuel_liters || 0);
        acc.hours_run += Number(r.hours_run || 0);
        acc.km_run += Number(r.km_run || 0);
        if (r.metric_mode === "km") {
          if (Number(r.km_run || 0) > 0) acc.km_fuel += Number(r.fuel_liters || 0);
        } else {
          if (Number(r.hours_run || 0) > 0) acc.hours_fuel += Number(r.fuel_liters || 0);
        }
        if (r.flag === "EXCESSIVE") acc.excessive += 1;
        return acc;
      },
      { assets: 0, fuel_liters: 0, hours_run: 0, km_run: 0, excessive: 0, hours_fuel: 0, km_fuel: 0 }
    );
    summary.fuel_liters = Number(summary.fuel_liters.toFixed(2));
    summary.hours_run = Number(summary.hours_run.toFixed(2));
    summary.km_run = Number(summary.km_run.toFixed(2));
    summary.avg_lph = summary.hours_run > 0 ? Number((summary.hours_fuel / summary.hours_run).toFixed(3)) : null;
    summary.avg_km_per_l = summary.km_fuel > 0 ? Number((summary.km_run / summary.km_fuel).toFixed(3)) : null;

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const wsSummary = wb.addWorksheet("Summary");
    wsSummary.columns = [
      { header: "Field", key: "field", width: 30 },
      { header: "Value", key: "value", width: 24 },
    ];
    wsSummary.addRows([
      { field: "Period", value: `${start} to ${end}` },
      { field: "Tolerance (%)", value: Number((tolerance * 100).toFixed(2)) },
      { field: "Mode filter", value: modeFilter || "all" },
      { field: "Asset filter", value: assetFilter || "all" },
      { field: "Assets", value: summary.assets },
      { field: "Missing / Not Shown", value: missingRows.length },
      { field: "Excessive", value: summary.excessive },
      { field: "Fuel Total (L)", value: summary.fuel_liters },
      { field: "Hours Run", value: summary.hours_run },
      { field: "Distance Run (km)", value: summary.km_run },
      { field: "Avg L/hr", value: summary.avg_lph == null ? "" : summary.avg_lph },
      { field: "Avg km/L", value: summary.avg_km_per_l == null ? "" : summary.avg_km_per_l },
    ]);

    const wsRows = wb.addWorksheet("Fuel Benchmark");
    wsRows.columns = [
      { header: "Asset Code", key: "asset_code", width: 16 },
      { header: "Asset Name", key: "asset_name", width: 32 },
      { header: "Mode", key: "metric_mode", width: 10 },
      { header: "Fuel Liters", key: "fuel_liters", width: 14 },
      { header: "Km Run", key: "km_run", width: 12 },
      { header: "Hours Run", key: "hours_run", width: 12 },
      { header: "Actual L/hr", key: "actual_lph", width: 12 },
      { header: "OEM L/hr", key: "oem_lph", width: 12 },
      { header: "Threshold L/hr", key: "threshold_lph", width: 14 },
      { header: "Variance L/hr", key: "variance_lph", width: 14 },
      { header: "Actual km/L", key: "actual_km_per_l", width: 12 },
      { header: "OEM km/L", key: "oem_km_per_l", width: 12 },
      { header: "Threshold km/L", key: "threshold_km_per_l", width: 14 },
      { header: "Variance km/L", key: "variance_km_per_l", width: 14 },
      { header: "Fill Count", key: "fill_count", width: 10 },
      { header: "Flag", key: "flag", width: 12 },
    ];
    if (rows.length) {
      wsRows.addRows(rows);
    } else {
      wsRows.addRow({
        asset_code: "-",
        asset_name: "No fuel benchmark data for period",
        metric_mode: "-",
        fuel_liters: "",
        km_run: "",
        hours_run: "",
        actual_lph: "",
        oem_lph: "",
        threshold_lph: "",
        variance_lph: "",
        actual_km_per_l: "",
        oem_km_per_l: "",
        threshold_km_per_l: "",
        variance_km_per_l: "",
        fill_count: "",
        flag: "",
      });
    }

    const wsMissing = wb.addWorksheet("Missing Equipment");
    wsMissing.columns = [
      { header: "Asset Code", key: "asset_code", width: 16 },
      { header: "Asset Name", key: "asset_name", width: 32 },
      { header: "Category", key: "category", width: 18 },
      { header: "Mode", key: "metric_mode", width: 10 },
      { header: "Fuel Liters", key: "fuel_liters", width: 12 },
      { header: "Fill Count", key: "fill_count", width: 10 },
      { header: "Reason", key: "reason", width: 42 },
    ];
    if (missingRows.length) {
      wsMissing.addRows(missingRows);
    } else {
      wsMissing.addRow({
        asset_code: "-",
        asset_name: "No missing equipment for current filters",
        category: "",
        metric_mode: "",
        fuel_liters: "",
        fill_count: "",
        reason: "",
      });
    }

    const buffer = await wb.xlsx.writeBuffer();
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="AML_Fuel_Benchmark_${end}.xlsx"`)
      .send(buffer);
  });

  // GET /api/reports/fuel-machine-history.pdf?asset_code=A300AM&start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15&download=1
  app.get("/fuel-machine-history.pdf", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    const assetCode = String(req.query?.asset_code || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const toleranceInput = Number(req.query?.tolerance ?? 0.15);
    const tolerance = Number.isFinite(toleranceInput) ? Math.max(0, toleranceInput) : 0.15;
    const download = String(req.query?.download || "").trim() === "1";

    if (!assetCode) return reply.code(400).send({ error: "asset_code is required" });
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const asset = db.prepare(`
      SELECT
        id, asset_code, asset_name, category,
        COALESCE(baseline_fuel_l_per_hour, 5.0) AS oem_lph,
        COALESCE(baseline_fuel_km_per_l, 2.0) AS oem_kmpl,
        CASE
          WHEN UPPER(COALESCE(asset_code, '')) GLOB 'V[0-9][0-9]AM' THEN 'km'
          ELSE COALESCE(NULLIF(TRIM(utilization_mode), ''), CASE
          WHEN LOWER(COALESCE(category, '')) LIKE '%truck%'
            OR LOWER(COALESCE(category, '')) LIKE '%vehicle%'
            OR LOWER(COALESCE(category, '')) LIKE '%ldv%'
            OR LOWER(COALESCE(category, '')) LIKE '%pickup%'
            OR LOWER(COALESCE(category, '')) LIKE '%bakkie%'
            OR LOWER(COALESCE(asset_code, '')) LIKE 'ldv%'
            OR UPPER(COALESCE(asset_code, '')) GLOB 'V[0-9][0-9]AM'
            OR LOWER(COALESCE(asset_name, '')) LIKE '%ldv%'
            THEN 'km'
          ELSE 'hours'
        END)
        END AS metric_mode,
        COALESCE(NULLIF(km_per_hour_factor, 0), 10.0) AS km_per_hour_factor
      FROM assets
      WHERE asset_code = ?
    `).get(assetCode);
    if (!asset) return reply.code(404).send({ error: `asset not found: ${assetCode}` });

    const fuelRows = db.prepare(`
      SELECT
        fl.id,
        fl.log_date,
        COALESCE(fl.liters, 0) AS fuel_liters,
        COALESCE(CASE WHEN fl.hours_run > 0 THEN fl.hours_run ELSE 0 END, 0) AS meter_hours,
        COALESCE(fl.meter_run_value, 0) AS meter_run_value,
        COALESCE(LOWER(fl.meter_unit), '') AS meter_unit,
        fl.open_meter_value,
        fl.close_meter_value,
        fl.source
      FROM fuel_logs fl
      WHERE fl.asset_id = ?
        AND fl.log_date BETWEEN ? AND ?
      ORDER BY fl.log_date ASC, fl.id ASC
    `).all(asset.id, start, end);

    const previousFill = db.prepare(`
      SELECT fl.hours_run, fl.meter_run_value, COALESCE(LOWER(fl.meter_unit), '') AS meter_unit
      FROM fuel_logs fl
      WHERE fl.asset_id = ?
        AND fl.log_date < ?
        AND (
          fl.hours_run > 0
          OR fl.meter_run_value > 0
        )
      ORDER BY fl.log_date DESC, fl.id DESC
      LIMIT 1
    `).get(asset.id, start);

    const mode = String(asset.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
    const kmPerHour = Math.max(0.1, Number(asset.km_per_hour_factor || 10));
    const oem = Number(asset.oem_lph || 5);
    const oemK = Number(asset.oem_kmpl || 2);
    const threshold = oem * (1 + tolerance);
    const lowThresholdKmpl = oemK * Math.max(0, 1 - tolerance);
    function toModeMeter(row) {
      const unit = String(row?.meter_unit || "").toLowerCase();
      const v = Number(row?.meter_run_value || 0);
      if (unit === "km" && v > 0) return mode === "km" ? v : (v / kmPerHour);
      if (unit === "hours" && v > 0) return mode === "km" ? (v * kmPerHour) : v;
      const h = Number(row?.meter_hours || row?.hours_run || 0);
      if (h > 0) return mode === "km" ? (h * kmPerHour) : h;
      return 0;
    }
    let prevMeter = previousFill ? toModeMeter(previousFill) : 0;
    if (!(prevMeter > 0)) prevMeter = null;

    const rows = fuelRows.map((d) => {
      const meter = toModeMeter(d);
      const rowOpen = d.open_meter_value == null ? null : Number(d.open_meter_value);
      const rowClose = d.close_meter_value == null ? null : Number(d.close_meter_value);
      const hasRowMeters = rowOpen != null && rowClose != null && rowOpen > 0 && rowClose > 0;
      let openMeter = prevMeter;
      let closeMeter = meter > 0 ? meter : null;
      let runBetween = null;
      let invalidDelta = false;
      if (hasRowMeters) {
        openMeter = rowOpen;
        closeMeter = rowClose;
        const delta = rowClose - rowOpen;
        if (Number.isFinite(delta) && delta > 0) runBetween = delta;
        else if (Number.isFinite(delta) && delta <= 0) invalidDelta = true;
      } else if (prevMeter != null && meter > 0) {
        const delta = meter - prevMeter;
        if (Number.isFinite(delta) && delta > 0) runBetween = delta;
        else if (Number.isFinite(delta) && delta <= 0) invalidDelta = true;
      }
      const fuel = Number(d.fuel_liters || 0);
      const lph = (!invalidDelta && mode === "hours" && runBetween != null && runBetween > 0) ? fuel / runBetween : null;
      const kmpl = (!invalidDelta && mode === "km" && fuel > 0 && runBetween != null && runBetween > 0) ? runBetween / fuel : null;
      const flag = mode === "km"
        ? (invalidDelta ? "INVALID DELTA" : (kmpl != null && kmpl < lowThresholdKmpl ? "EXCESSIVE" : "OK"))
        : (invalidDelta ? "INVALID DELTA" : (lph != null && lph > threshold ? "EXCESSIVE" : "OK"));
      if (closeMeter != null && closeMeter > 0) prevMeter = closeMeter;
      return {
        log_date: d.log_date,
        metric_mode: mode,
        fuel_liters: Number(fuel.toFixed(2)),
        run_between: runBetween == null ? 0 : Number(runBetween.toFixed(2)),
        hours_between: mode === "hours" ? (runBetween == null ? 0 : Number(runBetween.toFixed(2))) : 0,
        km_between: mode === "km" ? (runBetween == null ? 0 : Number(runBetween.toFixed(2))) : 0,
        actual_lph: lph == null ? null : Number(lph.toFixed(3)),
        oem_lph: Number(oem.toFixed(3)),
        actual_km_per_l: kmpl == null ? null : Number(kmpl.toFixed(3)),
        oem_km_per_l: Number(oemK.toFixed(3)),
        flag,
        source: d.source || "",
      };
    });

    const summary = rows.reduce(
      (acc, r) => {
        acc.fill_days += 1;
        acc.fuel_liters += Number(r.fuel_liters || 0);
        acc.hours_between += Number(r.hours_between || 0);
        acc.km_between += Number(r.km_between || 0);
        if (r.flag === "EXCESSIVE") acc.excessive_days += 1;
        return acc;
      },
      { fill_days: 0, fuel_liters: 0, hours_between: 0, km_between: 0, excessive_days: 0 }
    );
    summary.fuel_liters = Number(summary.fuel_liters.toFixed(2));
    summary.hours_between = Number(summary.hours_between.toFixed(2));
    summary.km_between = Number(summary.km_between.toFixed(2));
    summary.metric_mode = mode;
    summary.avg_lph = summary.hours_between > 0
      ? Number((summary.fuel_liters / summary.hours_between).toFixed(3))
      : null;
    summary.avg_km_per_l = summary.fuel_liters > 0 && summary.km_between > 0
      ? Number((summary.km_between / summary.fuel_liters).toFixed(3))
      : null;

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Machine Fuel Fill History");
        kvGrid(doc, [
          { k: "Asset", v: `${asset.asset_code} ${asset.asset_name ? `- ${asset.asset_name}` : ""}` },
          { k: "Period", v: `${start} to ${end}` },
          { k: "Tolerance", v: `${fmtNum(tolerance * 100, 1)}% (${mode === "km" ? "below OEM" : "above OEM"})` },
          { k: mode === "km" ? "OEM km/L" : "OEM L/hr", v: mode === "km" ? fmtNum(oemK, 3) : fmtNum(oem, 3) },
          { k: "Fill Days", v: fmtNum(summary.fill_days, 0) },
          { k: "Excessive Days", v: fmtNum(summary.excessive_days, 0) },
          { k: "Fuel Total (L)", v: fmtNum(summary.fuel_liters, 2) },
          { k: mode === "km" ? "Distance Between Fills (km)" : "Hours Between Fills", v: mode === "km" ? fmtNum(summary.km_between, 2) : fmtNum(summary.hours_between, 2) },
          { k: mode === "km" ? "Avg km/L" : "Avg L/hr", v: mode === "km" ? (summary.avg_km_per_l == null ? "-" : fmtNum(summary.avg_km_per_l, 3)) : (summary.avg_lph == null ? "-" : fmtNum(summary.avg_lph, 3)) },
        ], 2);

        sectionTitle(doc, "Fill Entries");
        table(
          doc,
          [
            { key: "log_date", label: "Date", width: 0.16 },
            { key: "fuel_liters", label: "Fuel (L)", width: 0.14, align: "right" },
            { key: "run_between", label: mode === "km" ? "Distance (km)" : "Hours Between", width: 0.16, align: "right" },
            { key: "actual", label: mode === "km" ? "km/L" : "L/hr", width: 0.12, align: "right" },
            { key: "oem", label: mode === "km" ? "OEM km/L" : "OEM L/hr", width: 0.12, align: "right" },
            { key: "flag", label: "Status", width: 0.10, align: "center" },
            { key: "source", label: "Source", width: 0.20 },
          ],
          rows.length
            ? rows.map((r) => ({
                log_date: r.log_date,
                fuel_liters: fmtNum(r.fuel_liters, 2),
                run_between: mode === "km" ? fmtNum(r.km_between, 2) : fmtNum(r.hours_between, 2),
                actual: mode === "km"
                  ? (r.actual_km_per_l == null ? "-" : fmtNum(r.actual_km_per_l, 3))
                  : (r.actual_lph == null ? "-" : fmtNum(r.actual_lph, 3)),
                oem: mode === "km" ? fmtNum(r.oem_km_per_l, 3) : fmtNum(r.oem_lph, 3),
                flag: r.flag,
                source: r.source || "",
              }))
            : [{
                log_date: "-",
                fuel_liters: "-",
                run_between: "-",
                actual: "-",
                oem: "-",
                flag: "-",
                source: "No fill data in selected period",
              }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Fuel Machine History",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
        layout: "landscape",
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Fuel_History_${asset.asset_code}_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/vehicle-ldv-check/:id.pdf?download=1
  app.get("/vehicle-ldv-check/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    const download = String(req.query?.download || "").trim() === "1";
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).send({ error: "valid vehicle check id required" });
    }

    const check = db.prepare(`
      SELECT
        v.id,
        v.check_date,
        v.vehicle_registration,
        v.odometer_km,
        v.inspector_name,
        v.notes,
        v.created_at,
        a.asset_code,
        a.asset_name,
        a.category
      FROM vehicle_ldv_checks v
      JOIN assets a ON a.id = v.asset_id
      WHERE v.id = ?
    `).get(id);
    if (!check) return reply.code(404).send({ error: "vehicle check not found" });

    const photos = db.prepare(`
      SELECT id, file_path, caption, markers_json, created_at
      FROM vehicle_ldv_check_photos
      WHERE check_id = ?
      ORDER BY id ASC
    `).all(id).map((p) => {
      let markers = [];
      try {
        markers = p.markers_json ? JSON.parse(p.markers_json) : [];
      } catch {
        markers = [];
      }
      return { ...p, markers: Array.isArray(markers) ? markers : [] };
    });

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "LDV Vehicle Check");
        kvGrid(doc, [
          { k: "Check #", v: check.id },
          { k: "Date", v: check.check_date || "" },
          { k: "Vehicle", v: `${check.asset_code || ""} - ${check.asset_name || ""}`.trim() },
          { k: "Category", v: check.category || "" },
          { k: "Registration", v: check.vehicle_registration || "-" },
          { k: "Odometer (km)", v: check.odometer_km == null ? "-" : fmtNum(check.odometer_km, 0) },
          { k: "Inspector", v: check.inspector_name || "-" },
          { k: "Created At", v: check.created_at || "" },
        ], 2);

        sectionTitle(doc, "Notes");
        doc
          .font("Helvetica")
          .fontSize(10)
          .fillColor("#111111")
          .text(compactCell(check.notes || "-", 2000), {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });

        sectionTitle(doc, "Photos and Pinned Damages");
        if (!photos.length) {
          doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No photos attached.");
          return;
        }

        for (const p of photos) {
          const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
          const abs = resolveStorageAbs(rel);
          ensurePageSpace(doc, 300);
          doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
          doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`, {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });
          doc.moveDown(0.2);
          if (abs && fs.existsSync(abs)) {
            try {
              doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 190], align: "left", valign: "top" });
              doc.y += 196;
            } catch {
              doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo file exists but could not be rendered.");
              doc.moveDown(0.5);
            }
          } else {
            doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
            doc.moveDown(0.5);
          }

          if (!p.markers.length) {
            doc.font("Helvetica").fontSize(9).fillColor("#555555").text("No pinned damages on this photo.");
            doc.moveDown(0.4);
            continue;
          }

          doc.font("Helvetica-Bold").fontSize(9).fillColor("#111111").text("Pinned damages:");
          doc.moveDown(0.15);
          p.markers.forEach((m, idx) => {
            const label = compactCell(m?.label || "Damage", 80);
            const note = compactCell(m?.note || "", 160);
            doc
              .font("Helvetica")
              .fontSize(9)
              .fillColor("#111111")
              .text(
                `${idx + 1}. ${label}${note ? ` — ${note}` : ""} (x:${fmtNum((Number(m?.x) || 0) * 100, 1)}%, y:${fmtNum((Number(m?.y) || 0) * 100, 1)}%)`,
                {
                  width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
                }
              );
          });
          doc.moveDown(0.5);
        }
      },
      {
        title: "IRONLOG",
        subtitle: "LDV Vehicle Check Report",
        rightText: `Check #${check.id}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_LDV_Check_${check.id}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/vehicle-ldv-checks.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&asset_id=123&with_photos=1&download=1
  app.get("/vehicle-ldv-checks.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const assetId = Number(req.query?.asset_id || 0);
    const withPhotos = String(req.query?.with_photos || "").trim() === "1";
    const download = String(req.query?.download || "").trim() === "1";

    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const where = ["v.check_date >= ?", "v.check_date <= ?"];
    const params = [start, end];
    if (assetId > 0) {
      where.push("v.asset_id = ?");
      params.push(assetId);
    }

    const rows = db.prepare(`
      SELECT
        v.id,
        v.asset_id,
        v.check_date,
        v.vehicle_registration,
        v.odometer_km,
        v.inspector_name,
        v.notes,
        a.asset_code,
        a.asset_name
      FROM vehicle_ldv_checks v
      JOIN assets a ON a.id = v.asset_id
      WHERE ${where.join(" AND ")}
      ORDER BY v.check_date DESC, v.id DESC
      LIMIT 1000
    `).all(...params);

    const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
    const photosByCheck = new Map();
    if (ids.length) {
      const marks = ids.map(() => "?").join(",");
      const photos = db.prepare(`
        SELECT check_id, id, file_path, caption, markers_json, created_at
        FROM vehicle_ldv_check_photos
        WHERE check_id IN (${marks})
        ORDER BY check_id ASC, id ASC
      `).all(...ids);
      for (const p of photos) {
        const key = Number(p.check_id);
        if (!photosByCheck.has(key)) photosByCheck.set(key, []);
        let markers = [];
        try {
          markers = p.markers_json ? JSON.parse(p.markers_json) : [];
        } catch {
          markers = [];
        }
        photosByCheck.get(key).push({
          ...p,
          markers: Array.isArray(markers) ? markers : [],
        });
      }
    }

    const summary = {
      count: rows.length,
      vehicles: new Set(rows.map((r) => Number(r.asset_id))).size,
      inspectors: new Set(rows.map((r) => String(r.inspector_name || "").trim()).filter(Boolean)).size,
      photos: rows.reduce((acc, r) => acc + (photosByCheck.get(Number(r.id)) || []).length, 0),
      pins: rows.reduce(
        (acc, r) => acc + (photosByCheck.get(Number(r.id)) || []).reduce((n, p) => n + ((p.markers || []).length || 0), 0),
        0
      ),
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "LDV Vehicle Checks Summary");
        kvGrid(doc, [
          { k: "From", v: start },
          { k: "To", v: end },
          { k: "Checks", v: summary.count },
          { k: "Vehicles", v: summary.vehicles },
          { k: "Inspectors", v: summary.inspectors },
          { k: "Photos", v: summary.photos },
          { k: "Damage pins", v: summary.pins },
        ], 2);

        sectionTitle(doc, "Checks");
        table(
          doc,
          ["ID", "Date", "Vehicle", "Reg", "Inspector", "Photos", "Pins", "Notes"],
          rows.length
            ? rows.map((r) => {
                const photos = photosByCheck.get(Number(r.id)) || [];
                const pinCount = photos.reduce((n, p) => n + ((p.markers || []).length || 0), 0);
                return {
                  ID: r.id,
                  Date: r.check_date || "",
                  Vehicle: `${r.asset_code || ""} ${r.asset_name ? `- ${compactCell(r.asset_name, 30)}` : ""}`.trim(),
                  Reg: compactCell(r.vehicle_registration || "-", 16),
                  Inspector: compactCell(r.inspector_name || "-", 20),
                  Photos: photos.length,
                  Pins: pinCount,
                  Notes: compactCell(r.notes || "-", 40),
                };
              })
            : [{ ID: "-", Date: "-", Vehicle: "No checks found in selected period", Reg: "-", Inspector: "-", Photos: "-", Pins: "-", Notes: "-" }]
        );

        if (!withPhotos) return;

        for (const r of rows) {
          const photos = photosByCheck.get(Number(r.id)) || [];
          if (!photos.length) continue;
          ensurePageSpace(doc, 40);
          sectionTitle(doc, `Check #${r.id} — ${r.asset_code || ""} (${r.check_date || ""})`);
          for (const p of photos) {
            const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
            const abs = resolveStorageAbs(rel);
            ensurePageSpace(doc, 240);
            doc.font("Helvetica-Bold").fontSize(9).fillColor("#111111")
              .text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""} • pins: ${(p.markers || []).length}`);
            doc.moveDown(0.15);
            if (abs && fs.existsSync(abs)) {
              try {
                doc.image(abs, doc.page.margins.left, doc.y, { fit: [360, 150], align: "left", valign: "top" });
                doc.y += 156;
              } catch {
                doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo exists but could not be rendered.");
                doc.moveDown(0.4);
              }
            } else {
              doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
              doc.moveDown(0.4);
            }
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: "LDV Vehicle Checks",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_LDV_Checks_${start}_to_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/artisan-inspection-form.pdf?asset_id=123&date=YYYY-MM-DD&shift=day|night&inspector_name=...&form_number=...&download=1
  app.get("/artisan-inspection-form.pdf", async (req, reply) => {
    const download = String(req.query?.download || "").trim() === "1";
    const assetId = Number(req.query?.asset_id || 0);
    const date = String(req.query?.date || "").trim();
    const shift = String(req.query?.shift || "").trim().toLowerCase();
    const inspectorName = String(req.query?.inspector_name || "").trim();
    const formNumber = String(req.query?.form_number || "").trim() || makeArtisanFormNumber();

    let asset = null;
    if (assetId > 0) {
      asset = db.prepare(`SELECT id, asset_code, asset_name, category FROM assets WHERE id = ?`).get(assetId);
    }

    const safeDate = isDate(date) ? date : "";
    const safeShift = ["day", "night"].includes(shift) ? shift.toUpperCase() : "";
    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const checklistRows = [
      "Pre-start visual condition (machine / plant)",
      "Guards, covers, and safety devices",
      "Hydraulic hoses, leaks, and fittings",
      "Electrical panels / cabling / lights",
      "Lubrication points / levels",
      "Brakes / steering / controls response",
      "Alarms, horn, and warning systems",
      "Housekeeping around machine / plant",
    ];

    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);
        sectionTitle(doc, "Daily Artisan Inspection (Blank Form)");
        kvGrid(doc, [
          { k: "Form No.", v: formNumber },
          { k: "Date", v: safeDate || "________________" },
          { k: "Shift", v: safeShift || "________________" },
          { k: "Artisan", v: inspectorName || "________________" },
          { k: "Asset Code", v: asset?.asset_code || "________________" },
          { k: "Asset Name", v: asset?.asset_name || "________________" },
          { k: "Category", v: asset?.category || "________________" },
          { k: "Machine Hours", v: "________________" },
        ], 2);

        sectionTitle(doc, "General checklist (tick one)");
        doc.font("Helvetica").fontSize(10).fillColor("#111111");
        const leftX = doc.page.margins.left;
        const contentW = doc.page.width - doc.page.margins.left - doc.page.margins.right;
        const colOkX = leftX + contentW * 0.58;
        const colFailX = leftX + contentW * 0.68;
        const colNaX = leftX + contentW * 0.80;
        const noteLabelX = leftX + contentW * 0.06;
        const noteLineX = leftX + contentW * 0.16;
        const noteLineW = contentW * 0.80;
        doc.font("Helvetica-Bold").fontSize(9);
        const headY = doc.y;
        doc.text("OK", colOkX, headY);
        doc.text("FAIL", colFailX, headY);
        doc.text("N/A", colNaX, headY);
        doc.moveDown(0.6);
        doc.font("Helvetica").fontSize(10);
        checklistRows.forEach((label, idx) => {
          ensurePageSpace(doc, 42);
          const rowY = doc.y;
          doc.text(`${idx + 1}. ${label}`, leftX, rowY, { width: contentW * 0.55 });
          doc.text("[ ]", colOkX, rowY);
          doc.text("[ ]", colFailX, rowY);
          doc.text("[ ]", colNaX, rowY);
          const noteY = rowY + 13;
          doc.text("Note:", noteLabelX, noteY);
          doc.moveTo(noteLineX, noteY + 10).lineTo(noteLineX + noteLineW, noteY + 10).strokeColor("#777777").lineWidth(0.6).stroke();
          doc.y = noteY + 14;
          doc.moveDown(0.2);
        });

        sectionTitle(doc, "Notes");
        for (let i = 0; i < 5; i += 1) {
          doc.text("________________________________________________________________________________________", {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });
          doc.moveDown(0.2);
        }

        doc.moveDown(0.6);
        doc.text("Artisan signature: ____________________________    Time returned: ____________________", {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });
        doc.moveDown(0.2);
        doc.text("Supervisor received by: _______________________    Date: _____________________________", {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });
      },
      {
        title: "IRONLOG",
        subtitle: "Artisan Inspection Blank Form",
        rightText: `Form ${formNumber}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Artisan_Inspection_Form_${formNumber}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/artisan-inspection/:id.pdf?download=1
  app.get("/artisan-inspection/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    const download = String(req.query?.download || "").trim() === "1";
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).send({ error: "valid inspection id required" });
    }

    const aiInspectorCol = pickExistingColumn("artisan_inspections", ["inspector_name", "inspector"], "inspector_name");
    const hasAiMachineHours = hasColumn("artisan_inspections", "machine_hours");
    const hasAiLiveSnap = hasColumn("artisan_inspections", "live_hours_snapshot");
    const hasAiShift = hasColumn("artisan_inspections", "shift");
    const hasAiChecklist = hasColumn("artisan_inspections", "checklist_json");
    const hasAiLiveSource = hasColumn("artisan_inspections", "live_hours_source");
    const legacyMeterSql = `COALESCE((
          SELECT MAX(dh.closing_hours)
          FROM daily_hours dh
          WHERE dh.asset_id = ai.asset_id
            AND dh.closing_hours IS NOT NULL
            AND dh.work_date <= ai.inspection_date
        ), 0)`;
    const machineHoursSelect = hasAiMachineHours
      ? `COALESCE(ai.machine_hours, ${legacyMeterSql})`
      : legacyMeterSql;
    const liveSnapSelect = hasAiLiveSnap
      ? `COALESCE(ai.live_hours_snapshot, ${legacyMeterSql})`
      : legacyMeterSql;
    const liveSourceSelect = hasAiLiveSource ? "ai.live_hours_source" : "''";

    const inspection = db.prepare(`
      SELECT
        ai.id,
        ai.inspection_date,
        ai.${aiInspectorCol} AS inspector_name,
        ${hasColumn("artisan_inspections", "form_number") ? "ai.form_number" : "''"} AS form_number,
        ai.notes,
        ${hasAiShift ? "ai.shift" : "''"} AS shift,
        ${machineHoursSelect} AS machine_hours,
        ${liveSnapSelect} AS live_hours_snapshot,
        ${liveSourceSelect} AS live_hours_source,
        ${hasAiChecklist ? "ai.checklist_json" : `''`} AS checklist_json,
        a.asset_code,
        a.asset_name,
        a.category
      FROM artisan_inspections ai
      JOIN assets a ON a.id = ai.asset_id
      WHERE ai.id = ?
    `).get(id);
    if (!inspection) return reply.code(404).send({ error: "artisan inspection not found" });

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Daily Artisan Inspection");
        kvGrid(doc, [
          { k: "Inspection #", v: inspection.id },
          { k: "Form No.", v: inspection.form_number || "—" },
          { k: "Date", v: inspection.inspection_date || "" },
          { k: "Shift", v: inspection.shift ? String(inspection.shift).toUpperCase() : "—" },
          { k: "Inspector", v: inspection.inspector_name || "-" },
          { k: "Asset Code", v: inspection.asset_code || "" },
          { k: "Asset Name", v: inspection.asset_name || "" },
          { k: "Recorded machine hours", v: Number(inspection.machine_hours || 0).toFixed(1) },
          {
            k: "Live hours (snapshot)",
            v: `${Number(inspection.live_hours_snapshot ?? inspection.machine_hours ?? 0).toFixed(1)}${
              inspection.live_hours_source ? ` (${inspection.live_hours_source})` : ""
            }`,
          },
          { k: "Category", v: inspection.category || "" },
        ], 2);

        let checklist = [];
        try {
          const cj = JSON.parse(String(inspection.checklist_json || "[]"));
          if (Array.isArray(cj)) checklist = cj;
        } catch {}
        if (checklist.length) {
          sectionTitle(doc, "General checklist");
          doc.font("Helvetica").fontSize(10).fillColor("#111111");
          for (const c of checklist) {
            const st = c.ok === true ? "OK" : c.ok === false ? "FAIL" : "N/A";
            doc.text(`• ${String(c.label || c.key || "")}: ${st}${c.note ? ` — ${c.note}` : ""}`, {
              width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
            });
            doc.moveDown(0.15);
          }
        }

        sectionTitle(doc, "Notes");
        doc
          .font("Helvetica")
          .fontSize(10)
          .fillColor("#111111")
          .text(compactCell(inspection.notes || "-", 2000), {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });
      },
      {
        title: "IRONLOG",
        subtitle: "Artisan Inspection Report",
        rightText: `Inspection #${inspection.id}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Artisan_Inspection_${inspection.id}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/manager-inspection/:id.pdf?download=1
  app.get("/manager-inspection/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    const download = String(req.query?.download || "").trim() === "1";
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).send({ error: "valid inspection id required" });
    }

    const miInspectorCol = pickExistingColumn("manager_inspections", ["inspector_name", "inspector"], "inspector_name");
    const miNotesCol = pickExistingColumn("manager_inspections", ["notes", "note", "remarks", "description"], "notes");
    const hasMiChecklistDetail = hasColumn("manager_inspections", "checklist_detail_json");
    const photoInspectionCol = pickExistingColumn("manager_inspection_photos", ["inspection_id", "manager_inspection_id"], "inspection_id");
    const photoPathCol = pickExistingColumn("manager_inspection_photos", ["file_path", "photo_path", "path", "image_path", "url"], "file_path");
    const photoCaptionCol = pickExistingColumn("manager_inspection_photos", ["caption", "note", "notes", "description"], "caption");
    const photoCreatedCol = pickExistingColumn("manager_inspection_photos", ["created_at", "uploaded_at", "created_on"], "created_at");

    const hasMiMachineHours = hasColumn("manager_inspections", "machine_hours");
    const hasMiLiveSnap = hasColumn("manager_inspections", "live_hours_snapshot");
    const hasMiChecklist = hasColumn("manager_inspections", "checklist_json");
    const hasMiParts = hasColumn("manager_inspections", "required_parts_json");
    const hasMiWo = hasColumn("manager_inspections", "work_order_id");
    const legacyMeterSql = `COALESCE((
          SELECT MAX(dh.closing_hours)
          FROM daily_hours dh
          WHERE dh.asset_id = mi.asset_id
            AND dh.closing_hours IS NOT NULL
            AND dh.work_date <= mi.inspection_date
        ), 0)`;
    const machineHoursSelect = hasMiMachineHours
      ? `COALESCE(mi.machine_hours, ${legacyMeterSql})`
      : legacyMeterSql;
    const liveSnapSelect = hasMiLiveSnap
      ? `COALESCE(mi.live_hours_snapshot, ${legacyMeterSql})`
      : legacyMeterSql;
    const liveSrcSelect = hasColumn("manager_inspections", "live_hours_source")
      ? `mi.live_hours_source`
      : `''`;

    const inspection = db.prepare(`
      SELECT
        mi.id,
        mi.inspection_date,
        mi.${miInspectorCol} AS inspector_name,
        mi.${miNotesCol} AS notes,
        ${hasMiChecklistDetail ? "mi.checklist_detail_json" : `''`} AS checklist_detail_json,
        mi.created_at,
        ${machineHoursSelect} AS machine_hours,
        ${liveSnapSelect} AS live_hours_snapshot,
        ${liveSrcSelect} AS live_hours_source,
        ${hasMiChecklist ? "mi.checklist_json" : `''`} AS checklist_json,
        ${hasMiParts ? "mi.required_parts_json" : `''`} AS required_parts_json,
        ${hasMiWo ? "mi.work_order_id" : `NULL`} AS work_order_id,
        a.asset_code,
        a.asset_name,
        a.category
      FROM manager_inspections mi
      JOIN assets a ON a.id = mi.asset_id
      WHERE mi.id = ?
    `).get(id);
    if (!inspection) return reply.code(404).send({ error: "manager inspection not found" });

    const photos = db.prepare(`
      SELECT id, ${photoPathCol} AS file_path, ${photoCaptionCol} AS caption, ${photoCreatedCol} AS created_at
      FROM manager_inspection_photos
      WHERE ${photoInspectionCol} = ?
      ORDER BY id ASC
    `).all(id);

    const toChecklistLabel = (key) =>
      String(key || "")
        .replaceAll("_", " ")
        .replace(/\b\w/g, (m) => m.toUpperCase())
        .trim();
    const parseManagerChecklist = (rawChecklist, rawDetails) => {
      let checklistParsed = null;
      let detailsParsed = null;
      try {
        checklistParsed = JSON.parse(String(rawChecklist || "null"));
      } catch {}
      try {
        detailsParsed = JSON.parse(String(rawDetails || "null"));
      } catch {}

      // Legacy/mobile ingest bundle shape: { checklist, checklist_details, ... }.
      if (checklistParsed && !Array.isArray(checklistParsed) && typeof checklistParsed === "object") {
        if (checklistParsed.checklist_details && (detailsParsed == null || typeof detailsParsed !== "object")) {
          detailsParsed = checklistParsed.checklist_details;
        }
        if (checklistParsed.checklist && typeof checklistParsed.checklist === "object") {
          checklistParsed = checklistParsed.checklist;
        }
      }

      if (Array.isArray(checklistParsed)) {
        return checklistParsed.map((c) => ({
          key: String(c?.key || "").trim(),
          label: String(c?.label || c?.key || "").trim(),
          ok: c?.ok === true ? true : c?.ok === false ? false : null,
          note: String(c?.note || "").trim() || null,
        }));
      }

      if (checklistParsed && typeof checklistParsed === "object") {
        return Object.entries(checklistParsed).map(([key, status]) => {
          const st = String(status || "").trim().toLowerCase();
          const ok = st === "ok" ? true : (st === "attention" || st === "unsafe" || st === "fail" || st === "failed") ? false : null;
          const detail = detailsParsed && typeof detailsParsed === "object" ? detailsParsed[key] : null;
          const note = String(detail?.comment || detail?.note || detail?.notes || "").trim() || null;
          return { key, label: toChecklistLabel(key), ok, note };
        });
      }
      return [];
    };

    const parseComponentNotes = (raw) => {
      const text = String(raw || "").trim();
      if (!text) return [];
      return text
        .split(/\r?\n|;/)
        .map((s) => String(s || "").trim())
        .filter(Boolean)
        .map((line) => {
          const m = line.match(/^([^:|-]+)\s*[:|-]\s*(.+)$/);
          if (m) {
            return { component: String(m[1] || "").trim(), note: String(m[2] || "").trim() };
          }
          return { component: "General", note: line };
        });
    };
    const extractChecklistDetailNotes = (raw) => {
      const text = String(raw || "").trim();
      if (!text) return [];
      try {
        const parsed = JSON.parse(text);
        const out = [];
        const walk = (node, label = "") => {
          if (node == null) return;
          if (Array.isArray(node)) {
            for (const item of node) walk(item, label);
            return;
          }
          if (typeof node === "object") {
            const comp = String(node.component || node.label || node.key || label || "General").trim();
            const note = String(node.note || node.notes || node.description || node.comment || "").trim();
            if (note) out.push({ component: comp, note });
            for (const [k, v] of Object.entries(node)) {
              if (["component", "label", "key", "note", "notes", "description", "comment"].includes(k)) continue;
              walk(v, comp || k);
            }
          }
        };
        walk(parsed);
        return out;
      } catch {
        return [];
      }
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Manager Inspection");
        kvGrid(doc, [
          { k: "Inspection #", v: inspection.id },
          { k: "Date", v: inspection.inspection_date || "" },
          { k: "Inspector", v: inspection.inspector_name || "-" },
          { k: "Asset Code", v: inspection.asset_code || "" },
          { k: "Asset Name", v: inspection.asset_name || "" },
          { k: "Recorded machine hours", v: Number(inspection.machine_hours || 0).toFixed(1) },
          {
            k: "Live hours (snapshot)",
            v: `${Number(inspection.live_hours_snapshot ?? inspection.machine_hours ?? 0).toFixed(1)}${
              inspection.live_hours_source ? ` (${inspection.live_hours_source})` : ""
            }`,
          },
          { k: "Work order", v: inspection.work_order_id ? `#${inspection.work_order_id}` : "—" },
          { k: "Category", v: inspection.category || "" },
        ], 2);

        const checklist = parseManagerChecklist(inspection.checklist_json, inspection.checklist_detail_json);
        if (checklist.length) {
          sectionTitle(doc, "Checklist");
          table(
            doc,
            ["Component", "Status", "Note"],
            checklist.map((c) => ({
              Component: compactCell(String(c.label || c.key || "-"), 80),
              Status: c.ok === true ? "OK" : c.ok === false ? "FAIL" : "N/A",
              Note: compactCell(String(c.note || "-"), 150),
            })),
            [0.35, 0.15, 0.5]
          );
        }

        let reqParts = [];
        try {
          const pj = JSON.parse(String(inspection.required_parts_json || "[]"));
          if (Array.isArray(pj)) reqParts = pj;
        } catch {}
        if (reqParts.length) {
          sectionTitle(doc, "Required parts");
          doc.font("Helvetica").fontSize(10).fillColor("#111111");
          for (const p of reqParts) {
            doc.text(
              `• ${String(p.part_code || "")} × ${Number(p.qty || 0)}${p.note ? ` — ${p.note}` : ""}`,
              { width: doc.page.width - doc.page.margins.left - doc.page.margins.right }
            );
            doc.moveDown(0.15);
          }
        }

        const rawNotes = String(inspection.notes || "").trim();
        let noteRows = parseComponentNotes(rawNotes);
        if (!noteRows.length) {
          noteRows = extractChecklistDetailNotes(inspection.checklist_detail_json);
        }
        sectionTitle(doc, "Notes");
        if (rawNotes) {
          doc
            .font("Helvetica")
            .fontSize(10)
            .fillColor("#111111")
            .text(rawNotes, {
              width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
            });
          doc.moveDown(0.2);
        }
        if (noteRows.length) {
          table(
            doc,
            ["Component", "Notes"],
            noteRows.map((n) => ({
              Component: compactCell(n.component || "General", 80),
              Notes: compactCell(n.note || "-", 180),
            })),
            [0.3, 0.7]
          );
        } else {
          doc
            .font("Helvetica")
            .fontSize(10)
            .fillColor("#111111")
            .text("-", {
              width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
            });
        }

        sectionTitle(doc, "Photos");
        if (!photos.length) {
          doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No photos attached.");
          return;
        }

        for (const p of photos) {
          const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
          const abs = resolveStorageAbs(rel);
          ensurePageSpace(doc, 230);
          doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
          doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`, {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });
          doc.moveDown(0.2);
          if (abs && fs.existsSync(abs)) {
            try {
              doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 180], align: "left", valign: "top" });
              doc.y += 186;
            } catch {
              doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo file exists but could not be rendered.");
              doc.moveDown(0.5);
            }
          } else {
            doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
            doc.moveDown(0.5);
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: "Manager Inspection Report",
        rightText: `Inspection #${inspection.id}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Manager_Inspection_${inspection.id}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/manager-inspections.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&asset_id=123&with_photos=1&download=1
  app.get("/manager-inspections.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const assetId = Number(req.query?.asset_id || 0);
    const withPhotos = String(req.query?.with_photos || "").trim() === "1";
    const download = String(req.query?.download || "").trim() === "1";

    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const miInspectorCol = pickExistingColumn("manager_inspections", ["inspector_name", "inspector"], "inspector_name");
    const miNotesCol = pickExistingColumn("manager_inspections", ["notes", "note", "remarks", "description"], "notes");
    const hasMiChecklistDetail = hasColumn("manager_inspections", "checklist_detail_json");
    const photoInspectionCol = pickExistingColumn("manager_inspection_photos", ["inspection_id", "manager_inspection_id"], "inspection_id");
    const photoPathCol = pickExistingColumn("manager_inspection_photos", ["file_path", "photo_path", "path", "image_path", "url"], "file_path");
    const photoCaptionCol = pickExistingColumn("manager_inspection_photos", ["caption", "note", "notes", "description"], "caption");
    const photoCreatedCol = pickExistingColumn("manager_inspection_photos", ["created_at", "uploaded_at", "created_on"], "created_at");

    const where = ["mi.inspection_date >= ?", "mi.inspection_date <= ?"];
    const params = [start, end];
    if (assetId > 0) {
      where.push("mi.asset_id = ?");
      params.push(assetId);
    }

    const hasMiMh = hasColumn("manager_inspections", "machine_hours");
    const hasMiWo = hasColumn("manager_inspections", "work_order_id");
    const hasMiChecklist = hasColumn("manager_inspections", "checklist_json");
    const toChecklistLabel = (key) =>
      String(key || "")
        .replaceAll("_", " ")
        .replace(/\b\w/g, (m) => m.toUpperCase())
        .trim();
    const parseManagerChecklist = (rawChecklist, rawDetails) => {
      let checklistParsed = null;
      let detailsParsed = null;
      try {
        checklistParsed = JSON.parse(String(rawChecklist || "null"));
      } catch {}
      try {
        detailsParsed = JSON.parse(String(rawDetails || "null"));
      } catch {}
      if (checklistParsed && !Array.isArray(checklistParsed) && typeof checklistParsed === "object") {
        if (checklistParsed.checklist_details && (detailsParsed == null || typeof detailsParsed !== "object")) {
          detailsParsed = checklistParsed.checklist_details;
        }
        if (checklistParsed.checklist && typeof checklistParsed.checklist === "object") {
          checklistParsed = checklistParsed.checklist;
        }
      }
      if (Array.isArray(checklistParsed)) {
        return checklistParsed.map((c) => ({
          key: String(c?.key || "").trim(),
          label: String(c?.label || c?.key || "").trim(),
          ok: c?.ok === true ? true : c?.ok === false ? false : null,
          note: String(c?.note || "").trim() || null,
        }));
      }
      if (checklistParsed && typeof checklistParsed === "object") {
        return Object.entries(checklistParsed).map(([key, status]) => {
          const st = String(status || "").trim().toLowerCase();
          const ok = st === "ok" ? true : (st === "attention" || st === "unsafe" || st === "fail" || st === "failed") ? false : null;
          const detail = detailsParsed && typeof detailsParsed === "object" ? detailsParsed[key] : null;
          const note = String(detail?.comment || detail?.note || detail?.notes || "").trim() || null;
          return { key, label: toChecklistLabel(key), ok, note };
        });
      }
      return [];
    };
    const rows = db.prepare(`
      SELECT
        mi.id,
        mi.asset_id,
        mi.inspection_date,
        mi.${miInspectorCol} AS inspector_name,
        mi.${miNotesCol} AS notes,
        ${hasMiChecklistDetail ? "mi.checklist_detail_json" : `''`} AS checklist_detail_json,
        ${hasMiMh ? "mi.machine_hours" : "NULL"} AS machine_hours,
        ${hasMiChecklist ? "mi.checklist_json" : `''`} AS checklist_json,
        ${hasMiWo ? "mi.work_order_id" : "NULL"} AS work_order_id,
        a.asset_code,
        a.asset_name
      FROM manager_inspections mi
      JOIN assets a ON a.id = mi.asset_id
      WHERE ${where.join(" AND ")}
      ORDER BY mi.inspection_date DESC, mi.id DESC
      LIMIT 1000
    `).all(...params);

    const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
    const photosByInspection = new Map();
    if (ids.length) {
      const marks = ids.map(() => "?").join(",");
      const photos = db.prepare(`
        SELECT ${photoInspectionCol} AS inspection_id, id, ${photoPathCol} AS file_path, ${photoCaptionCol} AS caption, ${photoCreatedCol} AS created_at
        FROM manager_inspection_photos
        WHERE ${photoInspectionCol} IN (${marks})
        ORDER BY ${photoInspectionCol} ASC, id ASC
      `).all(...ids);
      for (const p of photos) {
        const key = Number(p.inspection_id);
        if (!photosByInspection.has(key)) photosByInspection.set(key, []);
        photosByInspection.get(key).push(p);
      }
    }

    const summary = {
      count: rows.length,
      assets: new Set(rows.map((r) => Number(r.asset_id))).size,
      inspectors: new Set(rows.map((r) => String(r.inspector_name || "").trim()).filter(Boolean)).size,
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Manager Inspections Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Asset Filter", v: assetId > 0 ? String(assetId) : "All assets" },
          { k: "Inspections", v: String(summary.count) },
          { k: "Assets Covered", v: String(summary.assets) },
          { k: "Inspectors", v: String(summary.inspectors) },
        ], 2);

        sectionTitle(doc, "Inspection Entries");
        table(
          doc,
          [
            { key: "id", label: "ID", width: 0.07, align: "right" },
            { key: "date", label: "Date", width: 0.11 },
            { key: "asset", label: "Asset", width: 0.14 },
            { key: "name", label: "Asset Name", width: 0.18 },
            { key: "hrs", label: "Hrs", width: 0.07, align: "right" },
            { key: "wo", label: "WO", width: 0.07, align: "right" },
            { key: "inspector", label: "Inspector", width: 0.12 },
            { key: "notes", label: "Notes", width: 0.24 },
          ],
          rows.length
            ? rows.map((r) => {
                const checklist = parseManagerChecklist(r.checklist_json, r.checklist_detail_json);
                const failedChecklist = checklist.filter((c) => c && c.ok === false);
                const checklistFindings = failedChecklist
                  .map((c) => {
                    const label = String(c.label || c.key || "Item").trim();
                    const note = String(c.note || "").trim();
                    return note ? `${label}: ${note}` : label;
                  })
                  .filter(Boolean)
                  .join(" | ");
                const summaryNotes = [String(r.notes || "").trim(), checklistFindings]
                  .filter(Boolean)
                  .join(" | ");
                return {
                  id: String(r.id),
                  date: r.inspection_date || "",
                  asset: r.asset_code || "",
                  name: r.asset_name || "",
                  hrs:
                    r.machine_hours != null && Number.isFinite(Number(r.machine_hours))
                      ? Number(r.machine_hours).toFixed(1)
                      : "—",
                  wo: r.work_order_id ? String(r.work_order_id) : "—",
                  inspector: r.inspector_name || "-",
                  notes: compactCell(summaryNotes || "-", 100),
                };
              })
            : [{
                id: "-",
                date: "-",
                asset: "-",
                name: "No inspections found in selected period",
                hrs: "-",
                wo: "-",
                inspector: "-",
                notes: "-",
              }]
        );

        if (withPhotos) {
          sectionTitle(doc, "Inspection Photos");
          if (!rows.length) {
            doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No inspections in selected period.");
          } else {
            for (const r of rows) {
              const photos = photosByInspection.get(Number(r.id)) || [];
              const checklist = parseManagerChecklist(r.checklist_json, r.checklist_detail_json);
              const failedChecklist = checklist.filter((c) => c && c.ok === false);
              const checklistNotes = failedChecklist
                .map((c) => {
                  const label = String(c.label || c.key || "Item").trim();
                  const note = String(c.note || "").trim();
                  return note ? `${label}: ${note}` : label;
                })
                .filter(Boolean);
              const extractChecklistDetailText = (raw) => {
                const text = String(raw || "").trim();
                if (!text) return "";
                try {
                  const parsed = JSON.parse(text);
                  const snippets = [];
                  const walk = (node) => {
                    if (node == null) return;
                    if (Array.isArray(node)) {
                      for (const item of node) walk(item);
                      return;
                    }
                    if (typeof node === "object") {
                      const label = String(node.component || node.label || node.key || "").trim();
                      const note = String(node.note || node.notes || node.description || node.comment || "").trim();
                      if (note) snippets.push(label ? `${label}: ${note}` : note);
                      for (const [k, v] of Object.entries(node)) {
                        if (["component", "label", "key", "note", "notes", "description", "comment"].includes(k)) continue;
                        walk(v);
                      }
                    }
                  };
                  walk(parsed);
                  return snippets.join(" | ");
                } catch {
                  return "";
                }
              };
              const description = String(r.notes || "").trim() || extractChecklistDetailText(r.checklist_detail_json);
              ensurePageSpace(doc, 60);
              doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
              doc.text(
                `Inspection #${r.id} | ${r.inspection_date} | ${r.asset_code}${r.asset_name ? ` - ${r.asset_name}` : ""}`,
                { width: doc.page.width - doc.page.margins.left - doc.page.margins.right }
              );
              doc.moveDown(0.2);
              doc.font("Helvetica").fontSize(9).fillColor("#111111");
              doc.text(`Description: ${compactCell(description || "-", 700)}`, {
                width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
              });
              if (failedChecklist.length) {
                doc.moveDown(0.1);
                doc.text(`Checklist failures: ${failedChecklist.map((c) => String(c.label || c.key || "Item")).join("; ")}`, {
                  width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
                });
              }
              if (checklistNotes.length) {
                doc.moveDown(0.1);
                doc.text(`Failure notes: ${compactCell(checklistNotes.join(" | "), 700)}`, {
                  width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
                });
              }
              doc.moveDown(0.15);
              if (!photos.length) {
                doc.font("Helvetica").fontSize(9).fillColor("#666666").text("No photos attached.");
                doc.moveDown(0.3);
                continue;
              }
              for (const p of photos) {
                const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
                const abs = resolveStorageAbs(rel);
                ensurePageSpace(doc, 220);
                doc.font("Helvetica").fontSize(9).fillColor("#111111");
                doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`);
                doc.moveDown(0.15);
                if (abs && fs.existsSync(abs)) {
                  try {
                    doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 170], align: "left", valign: "top" });
                    doc.y += 176;
                  } catch {
                    doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo exists but could not be rendered.");
                    doc.moveDown(0.3);
                  }
                } else {
                  doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
                  doc.moveDown(0.3);
                }
              }
              doc.moveDown(0.2);
            }
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: withPhotos ? "Manager Inspections Report (With Photos)" : "Manager Inspections Report",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Manager_Inspections_${withPhotos ? "WithPhotos_" : ""}${start}_to_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/damage-report/:id.pdf?download=1
  app.get("/damage-report/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    const download = String(req.query?.download || "").trim() === "1";
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).send({ error: "valid damage report id required" });
    }

    const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
    const drPhotoReportCol = pickExistingColumn("manager_damage_report_photos", ["damage_report_id", "manager_damage_report_id", "report_id"], "damage_report_id");
    const drPhotoPathCol = pickExistingColumn("manager_damage_report_photos", ["file_path", "photo_path", "path", "image_path", "url", "image_data"], "file_path");
    const drPhotoCaptionCol = pickExistingColumn("manager_damage_report_photos", ["caption", "note", "notes", "description"], "caption");
    const drPhotoCreatedCol = pickExistingColumn("manager_damage_report_photos", ["created_at", "uploaded_at", "created_on"], "created_at");

    const report = db.prepare(`
      SELECT
        dr.id,
        dr.report_date,
        dr.${drInspectorCol} AS inspector_name,
        dr.hour_meter,
        dr.damage_location,
        dr.severity,
        dr.damage_description,
        dr.immediate_action,
        dr.out_of_service,
        dr.damage_time,
        dr.responsible_person,
        dr.pending_investigation,
        dr.hse_report_available,
        dr.created_at,
        a.asset_code,
        a.asset_name,
        a.category
      FROM manager_damage_reports dr
      JOIN assets a ON a.id = dr.asset_id
      WHERE dr.id = ?
    `).get(id);
    if (!report) return reply.code(404).send({ error: "damage report not found" });

    const photos = db.prepare(`
      SELECT id, ${drPhotoPathCol} AS file_path, ${drPhotoCaptionCol} AS caption, ${drPhotoCreatedCol} AS created_at
      FROM manager_damage_report_photos
      WHERE ${drPhotoReportCol} = ?
      ORDER BY id ASC
    `).all(id);

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Damage Report");
        kvGrid(doc, [
          { k: "Report #", v: report.id },
          { k: "Date", v: report.report_date || "" },
          { k: "Inspector", v: report.inspector_name || "-" },
          { k: "Asset Code", v: report.asset_code || "" },
          { k: "Asset Name", v: report.asset_name || "" },
          { k: "Hours", v: report.hour_meter == null ? "-" : Number(report.hour_meter || 0).toFixed(1) },
          { k: "Damage Time", v: report.damage_time || "-" },
          { k: "Location", v: report.damage_location || "-" },
          { k: "Responsible Person", v: report.responsible_person || "-" },
          { k: "Severity", v: String(report.severity || "-").toUpperCase() },
          { k: "Out of Service", v: Number(report.out_of_service || 0) ? "YES" : "NO" },
          { k: "Pending Investigation", v: Number(report.pending_investigation || 0) ? "YES" : "NO" },
          { k: "HSE Report Available", v: Number(report.hse_report_available || 0) ? "YES" : "NO" },
          { k: "Category", v: report.category || "" },
          { k: "Created At", v: report.created_at || "" },
        ], 2);

        sectionTitle(doc, "Damage Description");
        doc.font("Helvetica").fontSize(10).fillColor("#111111").text(compactCell(report.damage_description || "-", 2000), {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });

        sectionTitle(doc, "Immediate Action");
        doc.font("Helvetica").fontSize(10).fillColor("#111111").text(compactCell(report.immediate_action || "-", 2000), {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });

        sectionTitle(doc, "Photos");
        if (!photos.length) {
          doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No photos attached.");
          return;
        }
        for (const p of photos) {
          const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
          const abs = resolveStorageAbs(rel);
          ensurePageSpace(doc, 230);
          doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
          doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`, {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });
          doc.moveDown(0.2);
          if (abs && fs.existsSync(abs)) {
            try {
              doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 180], align: "left", valign: "top" });
              doc.y += 186;
            } catch {
              doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo file exists but could not be rendered.");
              doc.moveDown(0.5);
            }
          } else {
            doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
            doc.moveDown(0.5);
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: "Damage Report",
        rightText: `Report #${report.id}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename="AML_Damage_Report_${report.id}.pdf"`)
      .send(pdf);
  });

  // GET /api/reports/damage-reports.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&asset_id=123&with_photos=1&download=1
  app.get("/damage-reports.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const assetId = Number(req.query?.asset_id || 0);
    const withPhotos = String(req.query?.with_photos || "").trim() === "1";
    const download = String(req.query?.download || "").trim() === "1";

    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
    const drPhotoReportCol = pickExistingColumn("manager_damage_report_photos", ["damage_report_id", "manager_damage_report_id", "report_id"], "damage_report_id");
    const drPhotoPathCol = pickExistingColumn("manager_damage_report_photos", ["file_path", "photo_path", "path", "image_path", "url", "image_data"], "file_path");
    const drPhotoCaptionCol = pickExistingColumn("manager_damage_report_photos", ["caption", "note", "notes", "description"], "caption");
    const drPhotoCreatedCol = pickExistingColumn("manager_damage_report_photos", ["created_at", "uploaded_at", "created_on"], "created_at");

    const where = ["dr.report_date >= ?", "dr.report_date <= ?"];
    const params = [start, end];
    if (assetId > 0) {
      where.push("dr.asset_id = ?");
      params.push(assetId);
    }

    const rows = db.prepare(`
      SELECT
        dr.id,
        dr.asset_id,
        dr.report_date,
        dr.${drInspectorCol} AS inspector_name,
        dr.hour_meter,
        dr.damage_location,
        dr.severity,
        dr.damage_description,
        dr.immediate_action,
        dr.out_of_service,
        dr.damage_time,
        dr.responsible_person,
        dr.pending_investigation,
        dr.hse_report_available,
        a.asset_code,
        a.asset_name
      FROM manager_damage_reports dr
      JOIN assets a ON a.id = dr.asset_id
      WHERE ${where.join(" AND ")}
      ORDER BY dr.report_date DESC, dr.id DESC
      LIMIT 1000
    `).all(...params);

    const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
    const photosByReport = new Map();
    if (ids.length) {
      const marks = ids.map(() => "?").join(",");
      const photos = db.prepare(`
        SELECT ${drPhotoReportCol} AS damage_report_id, id, ${drPhotoPathCol} AS file_path, ${drPhotoCaptionCol} AS caption, ${drPhotoCreatedCol} AS created_at
        FROM manager_damage_report_photos
        WHERE ${drPhotoReportCol} IN (${marks})
        ORDER BY ${drPhotoReportCol} ASC, id ASC
      `).all(...ids);
      for (const p of photos) {
        const key = Number(p.damage_report_id);
        if (!photosByReport.has(key)) photosByReport.set(key, []);
        photosByReport.get(key).push(p);
      }
    }

    const summary = {
      count: rows.length,
      assets: new Set(rows.map((r) => Number(r.asset_id))).size,
      out_of_service: rows.filter((r) => Number(r.out_of_service || 0) === 1).length,
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Damage Reports Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Asset Filter", v: assetId > 0 ? String(assetId) : "All assets" },
          { k: "Reports", v: String(summary.count) },
          { k: "Assets Covered", v: String(summary.assets) },
          { k: "Out of Service", v: String(summary.out_of_service) },
        ], 2);

        sectionTitle(doc, "Damage Entries");
        table(
          doc,
          [
            { key: "id", label: "ID", width: 0.08, align: "right" },
            { key: "date", label: "Date", width: 0.12 },
            { key: "asset", label: "Asset", width: 0.13 },
            { key: "inspector", label: "Inspector", width: 0.11 },
            { key: "time", label: "Time", width: 0.06, align: "center" },
            { key: "location", label: "Location", width: 0.1 },
            { key: "resp", label: "Responsible", width: 0.1 },
            { key: "severity", label: "Severity", width: 0.08 },
            { key: "hours", label: "Hours", width: 0.08, align: "right" },
            { key: "out", label: "OOS", width: 0.06, align: "center" },
            { key: "inv", label: "Inv", width: 0.05, align: "center" },
            { key: "hse", label: "HSE", width: 0.05, align: "center" },
            { key: "desc", label: "Description", width: 0.08 },
          ],
          rows.length
            ? rows.map((r) => ({
                id: String(r.id),
                date: r.report_date || "",
                asset: r.asset_code || "",
                inspector: r.inspector_name || "-",
                time: r.damage_time || "-",
                location: compactCell(r.damage_location || "", 40),
                resp: compactCell(r.responsible_person || "", 20),
                severity: String(r.severity || "-").toUpperCase(),
                hours: r.hour_meter == null ? "-" : Number(r.hour_meter || 0).toFixed(1),
                out: Number(r.out_of_service || 0) ? "YES" : "NO",
                inv: Number(r.pending_investigation || 0) ? "Y" : "N",
                hse: Number(r.hse_report_available || 0) ? "Y" : "N",
                desc: compactCell(r.damage_description || "", 60),
              }))
            : [{
                id: "-",
                date: "-",
                asset: "-",
                inspector: "-",
                time: "-",
                location: "-",
                resp: "-",
                severity: "-",
                hours: "-",
                out: "-",
                inv: "-",
                hse: "-",
                desc: "No damage reports found in selected period",
              }]
        );

        if (withPhotos) {
          sectionTitle(doc, "Damage Photos");
          if (!rows.length) {
            doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No damage reports in selected period.");
          } else {
            for (const r of rows) {
              const photos = photosByReport.get(Number(r.id)) || [];
              ensurePageSpace(doc, 60);
              doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
              doc.text(
                `Report #${r.id} | ${r.report_date} | ${r.asset_code}${r.asset_name ? ` - ${r.asset_name}` : ""} | Severity: ${String(r.severity || "-").toUpperCase()}`,
                { width: doc.page.width - doc.page.margins.left - doc.page.margins.right }
              );
              doc.moveDown(0.2);
              if (!photos.length) {
                doc.font("Helvetica").fontSize(9).fillColor("#666666").text("No photos attached.");
                doc.moveDown(0.3);
                continue;
              }
              for (const p of photos) {
                const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
                const abs = resolveStorageAbs(rel);
                ensurePageSpace(doc, 220);
                doc.font("Helvetica").fontSize(9).fillColor("#111111");
                doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`);
                doc.moveDown(0.15);
                if (abs && fs.existsSync(abs)) {
                  try {
                    doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 170], align: "left", valign: "top" });
                    doc.y += 176;
                  } catch {
                    doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo exists but could not be rendered.");
                    doc.moveDown(0.3);
                  }
                } else {
                  doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
                  doc.moveDown(0.3);
                }
              }
              doc.moveDown(0.2);
            }
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: withPhotos ? "Damage Reports (With Photos)" : "Damage Reports",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename="AML_Damage_Reports_${withPhotos ? "WithPhotos_" : ""}${start}_to_${end}.pdf"`)
      .send(pdf);
  });

  // GET /api/reports/damage-reports.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD&asset_id=123
  app.get("/damage-reports.xlsx", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const assetId = Number(req.query?.asset_id || 0);
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
    const drPhotoReportCol = pickExistingColumn("manager_damage_report_photos", ["damage_report_id", "manager_damage_report_id", "report_id"], "damage_report_id");

    const where = ["dr.report_date >= ?", "dr.report_date <= ?"];
    const params = [start, end];
    if (assetId > 0) {
      where.push("dr.asset_id = ?");
      params.push(assetId);
    }

    const rows = db.prepare(`
      SELECT
        dr.id,
        dr.asset_id,
        dr.report_date,
        dr.${drInspectorCol} AS inspector_name,
        dr.hour_meter,
        dr.damage_location,
        dr.severity,
        dr.damage_description,
        dr.immediate_action,
        dr.out_of_service,
        dr.damage_time,
        dr.responsible_person,
        dr.pending_investigation,
        dr.hse_report_available,
        a.asset_code,
        a.asset_name
      FROM manager_damage_reports dr
      JOIN assets a ON a.id = dr.asset_id
      WHERE ${where.join(" AND ")}
      ORDER BY dr.report_date DESC, dr.id DESC
      LIMIT 5000
    `).all(...params);

    const photoCounts = new Map();
    if (rows.length) {
      const ids = rows.map((r) => Number(r.id || 0)).filter((n) => n > 0);
      if (ids.length) {
        const marks = ids.map(() => "?").join(",");
        const grouped = db.prepare(`
          SELECT ${drPhotoReportCol} AS damage_report_id, COUNT(*) AS photo_count
          FROM manager_damage_report_photos
          WHERE ${drPhotoReportCol} IN (${marks})
          GROUP BY ${drPhotoReportCol}
        `).all(...ids);
        grouped.forEach((g) => photoCounts.set(Number(g.damage_report_id || 0), Number(g.photo_count || 0)));
      }
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();
    const ws = wb.addWorksheet("Damage Reports");
    ws.columns = [
      { header: "Report ID", key: "id", width: 12 },
      { header: "Report Date", key: "report_date", width: 14 },
      { header: "Damage Time", key: "damage_time", width: 12 },
      { header: "Asset Code", key: "asset_code", width: 14 },
      { header: "Asset Name", key: "asset_name", width: 28 },
      { header: "Inspector", key: "inspector_name", width: 20 },
      { header: "Hour Meter", key: "hour_meter", width: 12 },
      { header: "Severity", key: "severity", width: 12 },
      { header: "Damage Location", key: "damage_location", width: 24 },
      { header: "Responsible Person", key: "responsible_person", width: 24 },
      { header: "Damage Description", key: "damage_description", width: 44 },
      { header: "Immediate Action", key: "immediate_action", width: 36 },
      { header: "Out Of Service", key: "out_of_service", width: 14 },
      { header: "Pending Investigation", key: "pending_investigation", width: 18 },
      { header: "HSE Report Available", key: "hse_report_available", width: 18 },
      { header: "Photo Count", key: "photo_count", width: 12 },
    ];
    ws.getRow(1).font = { bold: true };
    rows.forEach((r) => {
      ws.addRow({
        id: Number(r.id || 0),
        report_date: r.report_date || "",
        damage_time: r.damage_time || "",
        asset_code: r.asset_code || "",
        asset_name: r.asset_name || "",
        inspector_name: r.inspector_name || "",
        hour_meter: r.hour_meter == null ? "" : Number(r.hour_meter || 0),
        severity: String(r.severity || "").toUpperCase(),
        damage_location: r.damage_location || "",
        responsible_person: r.responsible_person || "",
        damage_description: r.damage_description || "",
        immediate_action: r.immediate_action || "",
        out_of_service: Number(r.out_of_service || 0) ? "YES" : "NO",
        pending_investigation: Number(r.pending_investigation || 0) ? "YES" : "NO",
        hse_report_available: Number(r.hse_report_available || 0) ? "YES" : "NO",
        photo_count: photoCounts.get(Number(r.id || 0)) || 0,
      });
    });
    ws.views = [{ state: "frozen", ySplit: 1 }];

    const summary = wb.addWorksheet("Summary");
    summary.columns = [
      { header: "Metric", key: "metric", width: 26 },
      { header: "Value", key: "value", width: 24 },
    ];
    summary.getRow(1).font = { bold: true };
    summary.addRows([
      { metric: "Start Date", value: start },
      { metric: "End Date", value: end },
      { metric: "Asset Filter", value: assetId > 0 ? String(assetId) : "All assets" },
      { metric: "Total Reports", value: rows.length },
      { metric: "Assets Covered", value: new Set(rows.map((r) => Number(r.asset_id || 0))).size },
      { metric: "Out Of Service", value: rows.filter((r) => Number(r.out_of_service || 0) === 1).length },
      { metric: "Pending Investigation", value: rows.filter((r) => Number(r.pending_investigation || 0) === 1).length },
      { metric: "HSE Report Available", value: rows.filter((r) => Number(r.hse_report_available || 0) === 1).length },
    ]);

    const buffer = await wb.xlsx.writeBuffer();
    return reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="AML_Damage_Reports_${start}_to_${end}.xlsx"`)
      .send(Buffer.from(buffer));
  });

  // =========================
  // STOCK MONITOR PDF
  // =========================
  // GET /api/reports/stock-monitor.pdf?part_code=FLT&download=1
  app.get("/stock-monitor.pdf", async (req, reply) => {
    const part_code = String(req.query?.part_code || "").trim();
    const download = String(req.query?.download || "").trim() === "1";

    const where = [];
    const params = [];
    if (part_code) {
      where.push("p.part_code LIKE ?");
      params.push(`%${part_code}%`);
    }

    const rows = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        p.critical,
        p.min_stock,
        IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      GROUP BY p.id
      ORDER BY p.critical DESC, on_hand ASC, p.part_code ASC
      LIMIT 500
    `).all(...params).map((r) => ({
      part_code: r.part_code,
      part_name: r.part_name,
      critical: Boolean(r.critical),
      min_stock: Number(r.min_stock || 0),
      on_hand: Number(r.on_hand || 0),
      below_min: Number(r.on_hand || 0) < Number(r.min_stock || 0),
    }));

    const summary = {
      total_parts: rows.length,
      below_min: rows.filter((r) => r.below_min).length,
      critical_below_min: rows.filter((r) => r.below_min && r.critical).length,
      total_on_hand: Number(rows.reduce((acc, r) => acc + Number(r.on_hand || 0), 0).toFixed(2)),
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Stock Monitor Summary");
        kvGrid(doc, [
          { k: "Filter", v: part_code || "All parts" },
          { k: "Total Parts", v: fmtNum(summary.total_parts, 0) },
          { k: "Below Min", v: fmtNum(summary.below_min, 0) },
          { k: "Critical Below Min", v: fmtNum(summary.critical_below_min, 0) },
          { k: "Total On Hand", v: fmtNum(summary.total_on_hand, 1) },
        ], 2);

        sectionTitle(doc, "Stock Levels");
        table(
          doc,
          [
            { key: "part_code", label: "Part Code", width: 0.16 },
            { key: "part_name", label: "Part Name", width: 0.44 },
            { key: "on_hand", label: "On Hand", width: 0.12, align: "right" },
            { key: "min_stock", label: "Min", width: 0.10, align: "right" },
            { key: "critical", label: "Critical", width: 0.08, align: "center" },
            { key: "below", label: "Below Min", width: 0.10, align: "center" },
          ],
          rows.length
            ? rows.map((r) => ({
                part_code: r.part_code,
                part_name: r.part_name || "",
                on_hand: fmtNum(r.on_hand, 1),
                min_stock: fmtNum(r.min_stock, 1),
                critical: r.critical ? "Y" : "N",
                below: r.below_min ? "YES" : "NO",
              }))
            : [{
                part_code: "-",
                part_name: "No parts for selected filter",
                on_hand: "-",
                min_stock: "-",
                critical: "-",
                below: "-",
              }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Stock Monitor Report",
        rightText: part_code ? `Filter: ${part_code}` : "All parts",
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Stock_Monitor${part_code ? `_${part_code}` : ""}_${todayYmd()}.pdf"`
      )
      .send(pdf);
  });

  // =========================
  // LEGAL COMPLIANCE PDF
  // =========================
  // GET /api/reports/legal-compliance.pdf?days=90&department=&status=approved&download=1
  app.get("/legal-compliance.pdf", async (req, reply) => {
    const daysRaw = Number(req.query?.days ?? 90);
    const days = Number.isFinite(daysRaw) ? Math.min(3650, Math.max(1, Math.trunc(daysRaw))) : 90;
    const department = String(req.query?.department || "").trim();
    const status = String(req.query?.status || "approved").trim().toLowerCase();
    const download = String(req.query?.download || "").trim() === "1";

    const where = ["ld.expiry_date IS NOT NULL", "TRIM(ld.expiry_date) <> ''"];
    const params = [];
    if (department) {
      where.push("ld.department = ?");
      params.push(department);
    }
    if (status && status !== "all") {
      where.push("ld.status = ?");
      params.push(status);
    }

    const rows = db.prepare(`
      SELECT
        ld.id,
        ld.department,
        ld.title,
        ld.doc_type,
        ld.version,
        ld.owner,
        ld.status,
        ld.active,
        ld.expiry_date,
        CAST(julianday(ld.expiry_date) - julianday(DATE('now')) AS INTEGER) AS days_to_expiry
      FROM legal_documents ld
      WHERE ${where.join(" AND ")}
      ORDER BY ld.expiry_date ASC, ld.id DESC
      LIMIT 1000
    `).all(...params).map((r) => ({
      ...r,
      id: Number(r.id),
      active: Number(r.active),
      days_to_expiry: Number(r.days_to_expiry),
    }));

    const dueRows = rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= days);
    const expiredRows = rows.filter((r) => Number(r.days_to_expiry) < 0);
    const summary = {
      total_with_expiry: rows.length,
      expired: expiredRows.length,
      due_30: rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= 30).length,
      due_60: rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= 60).length,
      due_90: rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= 90).length,
      due_window: dueRows.length,
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Legal Compliance Summary");
        kvGrid(doc, [
          { k: "Window (days)", v: String(days) },
          { k: "Department", v: department || "All" },
          { k: "Status", v: status || "approved" },
          { k: "Total with Expiry", v: fmtNum(summary.total_with_expiry, 0) },
          { k: "Expired", v: fmtNum(summary.expired, 0) },
          { k: "Due in 30 days", v: fmtNum(summary.due_30, 0) },
          { k: "Due in 60 days", v: fmtNum(summary.due_60, 0) },
          { k: "Due in 90 days", v: fmtNum(summary.due_90, 0) },
          { k: `Due in ${days} days`, v: fmtNum(summary.due_window, 0) },
        ], 2);

        sectionTitle(doc, "Expired Documents");
        table(
          doc,
          [
            { key: "id", label: "ID", width: 0.08, align: "right" },
            { key: "department", label: "Department", width: 0.16 },
            { key: "title", label: "Title", width: 0.30 },
            { key: "status", label: "Status", width: 0.12 },
            { key: "expiry", label: "Expiry", width: 0.14 },
            { key: "days", label: "Days", width: 0.10, align: "right" },
            { key: "owner", label: "Owner", width: 0.10 },
          ],
          expiredRows.length
            ? expiredRows.map((r) => ({
                id: String(r.id),
                department: r.department || "-",
                title: compactCell(r.title || "-", 100),
                status: r.status || "-",
                expiry: r.expiry_date || "-",
                days: fmtNum(r.days_to_expiry, 0),
                owner: compactCell(r.owner || "-", 40),
              }))
            : [{ id: "-", department: "-", title: "No expired documents", status: "-", expiry: "-", days: "-", owner: "-" }]
        );

        sectionTitle(doc, `Due In ${days} Days`);
        table(
          doc,
          [
            { key: "id", label: "ID", width: 0.08, align: "right" },
            { key: "department", label: "Department", width: 0.16 },
            { key: "title", label: "Title", width: 0.30 },
            { key: "status", label: "Status", width: 0.12 },
            { key: "expiry", label: "Expiry", width: 0.14 },
            { key: "days", label: "Days", width: 0.10, align: "right" },
            { key: "owner", label: "Owner", width: 0.10 },
          ],
          dueRows.length
            ? dueRows.map((r) => ({
                id: String(r.id),
                department: r.department || "-",
                title: compactCell(r.title || "-", 100),
                status: r.status || "-",
                expiry: r.expiry_date || "-",
                days: fmtNum(r.days_to_expiry, 0),
                owner: compactCell(r.owner || "-", 40),
              }))
            : [{ id: "-", department: "-", title: "No due documents in selected window", status: "-", expiry: "-", days: "-", owner: "-" }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Legal Compliance Report",
        rightText: `${department || "All"} | ${days}d`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Legal_Compliance_${todayYmd()}.pdf"`
      )
      .send(pdf);
  });

  // =========================
  // DAILY XLSX
  // =========================
  app.get("/daily.xlsx", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");

    const date = String(req.query?.date || "").trim();
    const scheduled = Number(req.query?.scheduled ?? 10);

    if (!isDate(date)) return reply.code(400).send({ error: "date (YYYY-MM-DD) required" });

    const hours = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        a.category,
        dh.is_used,
        dh.scheduled_hours,
        dh.opening_hours,
        dh.closing_hours,
        dh.hours_run
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date = ?
      ORDER BY a.asset_code
    `).all(date);

    const fuel = db.prepare(`
      SELECT a.asset_code, a.asset_name, fl.liters, fl.source
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date = ?
      ORDER BY a.asset_code
    `).all(date);

    const oil = db.prepare(`
      SELECT a.asset_code, a.asset_name, ol.oil_type, ol.quantity
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date = ?
      ORDER BY a.asset_code
    `).all(date);

    const breakdownDowntimeCol = getBreakdownDowntimeColumn();
    const hasBreakdownEndAt = hasColumn("breakdowns", "end_at");
    const hasBreakdownStartAt = hasColumn("breakdowns", "start_at");
    const breakdownDateExpr = hasBreakdownStartAt ? "DATE(COALESCE(b.breakdown_date, b.start_at))" : "DATE(b.breakdown_date)";
    const breakdownStatusExpr = hasColumn("breakdowns", "status")
      ? "TRIM(LOWER(COALESCE(b.status, '')))"
      : "''";
    const breakdownStartAtSelect = hasBreakdownStartAt ? "b.start_at" : "NULL AS start_at";
    const breakdownParams = [date, date, date];
    if (hasBreakdownEndAt) breakdownParams.push(date);
    const breakdowns = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        b.breakdown_date,
        ${breakdownStartAtSelect},
        ${hasBreakdownEndAt ? "b.end_at" : "NULL AS end_at"},
        b.description,
        COALESCE(b.${breakdownDowntimeCol}, 0) AS downtime_hours,
        b.critical,
        COALESCE((
          SELECT COUNT(DISTINCT l.log_date)
          FROM breakdown_downtime_logs l
          WHERE l.breakdown_id = b.id
            AND l.log_date <= ?
        ), 0) AS logged_days
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      WHERE ${breakdownDateExpr} <= ?
        AND (
          ${hasBreakdownEndAt ? "b.end_at IS NULL OR DATE(b.end_at) >= ?" : "1 = 1"}
          OR ${breakdownStatusExpr} IN ('open', 'in_progress')
        )
      ORDER BY downtime_hours DESC
    `).all(...breakdownParams).map((r) => ({
      ...r,
      critical: Boolean(r.critical),
      days_down: daysDownForBreakdown(r, date),
    }));

    const upcoming = db.prepare(`
      SELECT
        mp.id AS plan_id,
        a.asset_code,
        a.asset_name,
        mp.service_name,
        mp.interval_hours,
        mp.last_service_hours,
        COALESCE(
          (SELECT dh.closing_hours
           FROM daily_hours dh
           WHERE dh.asset_id = a.id
             AND dh.work_date <= ?
             AND dh.closing_hours IS NOT NULL
           ORDER BY dh.work_date DESC
           LIMIT 1),
          (SELECT IFNULL(SUM(dh2.hours_run), 0)
           FROM daily_hours dh2
           WHERE dh2.asset_id = a.id
             AND dh2.work_date <= ?
             AND dh2.is_used = 1
             AND dh2.hours_run > 0)
        ) AS current_hours
      FROM maintenance_plans mp
      JOIN assets a ON a.id = mp.asset_id
      WHERE mp.active = 1
        AND a.active = 1
        AND a.is_standby = 0
      ORDER BY a.asset_code
    `).all(date, date).map(r => {
      const current = Number(r.current_hours || 0);
      const next_due = Number(r.last_service_hours || 0) + Number(r.interval_hours || 0);
      const hours_left = next_due - current;
      return {
        ...r,
        current_hours: current,
        next_due,
        hours_left,
        status: hours_left <= 0 ? "OVERDUE" : (hours_left <= 50 ? "DUE SOON" : "OK"),
      };
    }).sort((a, b) => a.hours_left - b.hours_left);

    const kpi = kpiDaily(date, scheduled);

    const fuel_total = fuel.reduce((a, r) => a + Number(r.liters || 0), 0);
    const oil_total = oil.reduce((a, r) => a + Number(r.quantity || 0), 0);
    const breakdown_total = breakdowns.reduce((a, r) => a + Number(r.downtime_hours || 0), 0);

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    buildDailyExecutiveSummarySheet(wb, {
      date,
      scheduled,
      kpi,
      fuel_total,
      oil_total,
      breakdown_total,
      includeCostEngine: false,
    });

    const dirTbl = { directorStyle: true };

    addTableSheet(
      wb,
      "Hours",
      [
        { header: "Asset code", key: "asset_code", width: 14 },
        { header: "Asset name", key: "asset_name", width: 26 },
        { header: "Category", key: "category", width: 14 },
        { header: "Production use (Y/N)", key: "is_used", width: 18 },
        { header: "Scheduled (h)", key: "scheduled_hours", width: 12 },
        { header: "Opening meter (h)", key: "opening_hours", width: 14 },
        { header: "Closing meter (h)", key: "closing_hours", width: 14 },
        { header: "Run hours", key: "hours_run", width: 12 },
      ],
      hours.map(r => ({ ...r, is_used: r.is_used ? "Y" : "N" })),
      dirTbl,
    );

    addTableSheet(
      wb,
      "Breakdowns",
      [
        { header: "Asset code", key: "asset_code", width: 14 },
        { header: "Asset name", key: "asset_name", width: 26 },
        { header: "Days down", key: "days_down", width: 12 },
        { header: "Downtime (hours)", key: "downtime_hours", width: 14 },
        { header: "Critical", key: "critical", width: 10 },
        { header: "Description", key: "description", width: 40 },
      ],
      breakdowns.map(r => ({ ...r, critical: r.critical ? "YES" : "NO" })),
      dirTbl,
    );

    addTableSheet(
      wb,
      "Fuel",
      [
        { header: "Asset code", key: "asset_code", width: 14 },
        { header: "Asset name", key: "asset_name", width: 26 },
        { header: "Litres", key: "liters", width: 12 },
        { header: "Source / notes", key: "source", width: 22 },
      ],
      fuel,
      dirTbl,
    );

    addTableSheet(
      wb,
      "Oil & lube",
      [
        { header: "Asset code", key: "asset_code", width: 14 },
        { header: "Asset name", key: "asset_name", width: 26 },
        { header: "Product type", key: "oil_type", width: 16 },
        { header: "Quantity", key: "quantity", width: 12 },
      ],
      oil,
      dirTbl,
    );

    addTableSheet(
      wb,
      "Maintenance outlook",
      [
        { header: "Asset code", key: "asset_code", width: 14 },
        { header: "Asset name", key: "asset_name", width: 26 },
        { header: "Service", key: "service_name", width: 22 },
        { header: "Interval (hours)", key: "interval_hours", width: 14 },
        { header: "Last service (meter h)", key: "last_service_hours", width: 18 },
        { header: "Current meter (h)", key: "current_hours", width: 16 },
        { header: "Next due at (meter h)", key: "next_due", width: 16 },
        { header: "Hours remaining", key: "hours_left", width: 14 },
        { header: "Status", key: "status", width: 12 },
      ],
      upcoming,
      dirTbl,
    );

    const buffer = await wb.xlsx.writeBuffer();

    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Daily_${date}.xlsx"`)
      .send(Buffer.from(buffer));
  });

  // =========================
  // MTD OPENING HOURS XLSX
  // =========================
  // GET /api/reports/mtd-opening-hours.xlsx?month=YYYY-MM
  app.get("/mtd-opening-hours.xlsx", async (req, reply) => {
    const monthRaw = String(req.query?.month || "").trim();
    const month = monthRaw || todayYmd().slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(month)) {
      return reply.code(400).send({ error: "month must be YYYY-MM" });
    }

    const start = `${month}-01`;
    const end = new Date(Date.UTC(Number(month.slice(0, 4)), Number(month.slice(5, 7)), 0))
      .toISOString()
      .slice(0, 10);
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "Invalid month range" });
    }

    const rows = db.prepare(`
      SELECT
        dh.work_date,
        a.asset_code,
        a.asset_name,
        a.category,
        dh.is_used,
        dh.opening_hours,
        dh.closing_hours,
        dh.hours_run
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date BETWEEN ? AND ?
      ORDER BY a.asset_code ASC, dh.work_date ASC
    `).all(start, end);

    const daysInMonth = Number(end.slice(-2));
    const dayKeys = Array.from({ length: daysInMonth }, (_, i) => String(i + 1).padStart(2, "0"));

    const byAsset = new Map();
    rows.forEach((r) => {
      const code = String(r.asset_code || "").trim();
      if (!code) return;
      if (!byAsset.has(code)) {
        byAsset.set(code, {
          asset_code: code,
          asset_name: r.asset_name || "",
          category: r.category || "",
          used_days: 0,
          logged_days: 0,
          ...Object.fromEntries(dayKeys.map((d) => [`d${d}`, null])),
        });
      }
      const rec = byAsset.get(code);
      const day = String(r.work_date || "").slice(-2);
      const opening = r.opening_hours == null ? null : Number(r.opening_hours);
      if (opening != null) rec[`d${day}`] = opening;
      rec.logged_days += 1;
      if (Number(r.is_used || 0) === 1) rec.used_days += 1;
    });

    const pivotRows = Array.from(byAsset.values()).sort((a, b) => String(a.asset_code).localeCompare(String(b.asset_code)));
    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    addTableSheet(
      wb,
      "Opening Detail",
      [
        { header: "Date", key: "work_date", width: 12 },
        { header: "Asset code", key: "asset_code", width: 14 },
        { header: "Asset name", key: "asset_name", width: 28 },
        { header: "Category", key: "category", width: 14 },
        { header: "Production use (Y/N)", key: "is_used", width: 18 },
        { header: "Opening meter (h)", key: "opening_hours", width: 16 },
        { header: "Closing meter (h)", key: "closing_hours", width: 16 },
        { header: "Run hours", key: "hours_run", width: 12 },
      ],
      rows.map((r) => ({
        work_date: r.work_date || "",
        asset_code: r.asset_code || "",
        asset_name: r.asset_name || "",
        category: r.category || "",
        is_used: Number(r.is_used || 0) === 1 ? "Y" : "N",
        opening_hours: r.opening_hours == null ? null : Number(r.opening_hours),
        closing_hours: r.closing_hours == null ? null : Number(r.closing_hours),
        hours_run: r.hours_run == null ? null : Number(r.hours_run),
      }))
    );

    addTableSheet(
      wb,
      "MTD by Equipment",
      [
        { header: "Asset code", key: "asset_code", width: 14 },
        { header: "Asset name", key: "asset_name", width: 24 },
        { header: "Category", key: "category", width: 14 },
        { header: "Days logged", key: "logged_days", width: 12 },
        { header: "Used days", key: "used_days", width: 10 },
        ...dayKeys.map((d) => ({ header: d, key: `d${d}`, width: 9 })),
      ],
      pivotRows
    );

    const summary = wb.addWorksheet("Summary");
    summary.columns = [
      { header: "Metric", key: "metric", width: 38 },
      { header: "Value", key: "value", width: 20 },
    ];
    summary.addRows([
      { metric: "Month", value: month },
      { metric: "Period start", value: start },
      { metric: "Period end", value: end },
      { metric: "Equipment with logs", value: pivotRows.length },
      { metric: "Daily rows logged", value: rows.length },
      {
        metric: "Opening meter values captured",
        value: rows.reduce((acc, r) => acc + (r.opening_hours == null ? 0 : 1), 0),
      },
    ]);

    const buffer = await wb.xlsx.writeBuffer();
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_MTD_Opening_Hours_${month}.xlsx"`)
      .send(Buffer.from(buffer));
  });

  // GET /api/reports/gm-weekly.xlsx?end=YYYY-MM-DD&scheduled=10
  // GM pack: MTD availability/utilization; downtime from daily logs when present; MTBF/MTTR; PM; spares; forecast.
  app.get("/gm-weekly.xlsx", async (req, reply) => {
    const end = String(req.query?.end || "").trim();
    const scheduled = Number(req.query?.scheduled ?? 10);
    const forecastHorizon = Math.max(1, Math.min(90, Number(req.query?.forecast_days ?? 30)));

    if (!isDate(end)) {
      return reply.code(400).send({ error: "end (YYYY-MM-DD) required" });
    }

    const mtdStart = monthStartIso(end);
    const mtdDays = inclusiveDaysBetween(mtdStart, end);
    const downtimeMtd = getDowntimeHoursForPeriod(mtdStart, end);
    const usesLogs = downtimeMtdUsesDailyLogs(mtdStart, end);

    const kpi = kpiRange(mtdStart, end, scheduled, { downtimeHoursOverride: downtimeMtd });
    const rel = reliabilityMetricsForRange(mtdStart, end, {
      downtimeHoursOverride: downtimeMtd,
      failuresInPeriodMode: "activity_or_report",
    });
    const pm = gmWeeklyPmComplianceSnapshot(end);
    const downtimeByAsset = getDowntimeByAssetMtd(mtdStart, end);

    const criticalSparesRows = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        p.min_stock,
        IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      WHERE p.critical = 1
      GROUP BY p.id
      ORDER BY (on_hand < p.min_stock) DESC, on_hand ASC, p.part_code ASC
    `).all().map((r) => ({
      part_code: r.part_code,
      part_name: r.part_name,
      min_stock: Number(r.min_stock || 0),
      on_hand: Number(r.on_hand || 0),
      status: Number(r.on_hand || 0) < Number(r.min_stock || 0) ? "Below min" : "OK",
    }));

    const forecast = gmWeeklyRepairForecast(end, forecastHorizon);

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    buildGmWeeklyExecutiveSheet(wb, {
      mtd_start: mtdStart,
      end,
      mtd_day_count: mtdDays,
      scheduled,
      uses_downtime_logs: usesLogs,
      kpi,
      rel,
      pm,
      spares: {
        critical_parts: criticalSparesRows.length,
        below_min: criticalSparesRows.filter((r) => r.status === "Below min").length,
      },
      forecast: {
        pm_count: forecast.pm_rows.length,
        open_wo_count: forecast.open_breakdown_repairs.length,
      },
      forecast_horizon_days: forecastHorizon,
    });

    const dirTbl = { directorStyle: true };

    addTableSheet(
      wb,
      "Downtime by asset (MTD)",
      [
        { header: "Asset code", key: "asset_code", width: 14 },
        { header: "Asset name", key: "asset_name", width: 28 },
        { header: "Downtime (hours)", key: "downtime_hours", width: 16 },
      ],
      downtimeByAsset.length
        ? downtimeByAsset
        : [{
            asset_code: "-",
            asset_name: "No downtime in month-to-date window",
            downtime_hours: 0,
          }],
      dirTbl,
    );

    addTableSheet(
      wb,
      "Critical spares",
      [
        { header: "Part code", key: "part_code", width: 14 },
        { header: "Part name", key: "part_name", width: 28 },
        { header: "Min stock", key: "min_stock", width: 10 },
        { header: "On hand", key: "on_hand", width: 10 },
        { header: "Status", key: "status", width: 12 },
      ],
      criticalSparesRows.length
        ? criticalSparesRows
        : [{ part_code: "-", part_name: "No critical parts in master", min_stock: "", on_hand: "", status: "" }],
      dirTbl,
    );

    const forecastRows = [
      ...forecast.pm_rows.map((r) => ({
        category: r.type,
        asset_code: r.asset_code,
        asset_name: r.asset_name,
        detail: r.detail,
        est_or_opened: r.est_date,
        extra: r.remaining_hours != null ? `Rem. ${r.remaining_hours} h` : "",
      })),
      ...forecast.open_breakdown_repairs.map((r) => ({
        category: r.type,
        asset_code: r.asset_code,
        asset_name: r.asset_name,
        detail: r.detail,
        est_or_opened: r.opened_at || "",
        extra: r.status ? `Status: ${r.status}` : "",
      })),
    ];

    addTableSheet(
      wb,
      "Repair forecast",
      [
        { header: "Category", key: "category", width: 16 },
        { header: "Asset code", key: "asset_code", width: 12 },
        { header: "Asset name", key: "asset_name", width: 22 },
        { header: "Detail", key: "detail", width: 36 },
        { header: "Est. date / opened", key: "est_or_opened", width: 22 },
        { header: "Notes", key: "extra", width: 24 },
      ],
      forecastRows.length
        ? forecastRows
        : [{
            category: "-",
            asset_code: "",
            asset_name: "",
            detail: `No PM dates in next ${forecastHorizon} days and no open breakdown WOs`,
            est_or_opened: "",
            extra: "",
          }],
      dirTbl,
    );

    const buffer = await wb.xlsx.writeBuffer();
    const safeMtdStart = mtdStart.replace(/-/g, "");
    const safeEnd = end.replace(/-/g, "");

    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_GM_Weekly_ME_MTD_${safeMtdStart}_${safeEnd}.xlsx"`)
      .send(Buffer.from(buffer));
  });

  // =========================
  // MONTHLY COST XLSX
  // =========================
  // GET /api/reports/cost-monthly.xlsx?month=YYYY-MM
  app.get("/cost-monthly.xlsx", async (req, reply) => {
    const month = String(req.query?.month || "").trim();
    if (!isMonth(month)) {
      return reply.code(400).send({ error: "month (YYYY-MM) required" });
    }

    const current = monthRange(month);
    const previousMonth = prevMonth(month);
    const previous = monthRange(previousMonth);
    const defaults = costDefaults();

    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";

    const buildAssetCosts = (start, end) => {
      const fuelRows = db.prepare(`
        SELECT a.asset_code, a.asset_name, a.category,
          COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS fuel_cost
        FROM fuel_logs fl
        JOIN assets a ON a.id = fl.asset_id
        WHERE fl.log_date BETWEEN ? AND ?
        GROUP BY a.id
      `).all(defaults.fuel_cost_per_liter_default, start, end);

      const lubeRows = db.prepare(`
        SELECT a.asset_code, a.asset_name, a.category,
          COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS lube_cost
        FROM oil_logs ol
        JOIN assets a ON a.id = ol.asset_id
        WHERE ol.log_date BETWEEN ? AND ?
        GROUP BY a.id
      `).all(defaults.lube_cost_per_qty_default, start, end);

      const partsRows = db.prepare(`
        SELECT
          COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
          COALESCE(a.asset_name, 'Unlinked') AS asset_name,
          COALESCE(a.category, 'Unassigned') AS category,
          COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS parts_cost
        FROM stock_movements sm
        JOIN parts p ON p.id = sm.part_id
        LEFT JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
        LEFT JOIN assets a ON a.id = w.asset_id
        WHERE sm.movement_type = 'out'
          AND ${smDateExpr} BETWEEN ? AND ?
        GROUP BY a.id
      `).all(start, end);

      const laborRows = db.prepare(`
        SELECT a.asset_code, a.asset_name, a.category,
          COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS labor_hours,
          COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
        FROM work_orders w
        JOIN assets a ON a.id = w.asset_id
        WHERE DATE(COALESCE(w.completed_at, w.closed_at)) BETWEEN ? AND ?
          AND w.status IN ('completed', 'approved', 'closed')
        GROUP BY a.id
      `).all(defaults.labor_cost_per_hour_default, start, end);

      const downtimeRows = db.prepare(`
        SELECT a.asset_code, a.asset_name, a.category,
          COALESCE(SUM(l.hours_down), 0) AS downtime_hours,
          COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS downtime_cost
        FROM breakdown_downtime_logs l
        JOIN breakdowns b ON b.id = l.breakdown_id
        JOIN assets a ON a.id = b.asset_id
        WHERE l.log_date BETWEEN ? AND ?
        GROUP BY a.id
      `).all(defaults.downtime_cost_per_hour_default, start, end);

      const map = new Map();
      const ensure = (r) => {
        const code = String(r.asset_code || "UNLINKED");
        if (!map.has(code)) {
          map.set(code, {
            asset_code: code,
            asset_name: r.asset_name || "Unlinked",
            category: r.category || "Unassigned",
            fuel_cost: 0,
            lube_cost: 0,
            parts_cost: 0,
            labor_hours: 0,
            labor_cost: 0,
            downtime_hours: 0,
            downtime_cost: 0,
            total_cost: 0,
          });
        }
        return map.get(code);
      };
      for (const r of fuelRows) ensure(r).fuel_cost += Number(r.fuel_cost || 0);
      for (const r of lubeRows) ensure(r).lube_cost += Number(r.lube_cost || 0);
      for (const r of partsRows) ensure(r).parts_cost += Number(r.parts_cost || 0);
      for (const r of laborRows) {
        const row = ensure(r);
        row.labor_hours += Number(r.labor_hours || 0);
        row.labor_cost += Number(r.labor_cost || 0);
      }
      for (const r of downtimeRows) {
        const row = ensure(r);
        row.downtime_hours += Number(r.downtime_hours || 0);
        row.downtime_cost += Number(r.downtime_cost || 0);
      }

      return Array.from(map.values())
        .map((r) => {
          const total = Number(r.fuel_cost || 0) + Number(r.lube_cost || 0) + Number(r.parts_cost || 0) + Number(r.labor_cost || 0) + Number(r.downtime_cost || 0);
          return {
            ...r,
            fuel_cost: Number(r.fuel_cost.toFixed(2)),
            lube_cost: Number(r.lube_cost.toFixed(2)),
            parts_cost: Number(r.parts_cost.toFixed(2)),
            labor_hours: Number(r.labor_hours.toFixed(2)),
            labor_cost: Number(r.labor_cost.toFixed(2)),
            downtime_hours: Number(r.downtime_hours.toFixed(2)),
            downtime_cost: Number(r.downtime_cost.toFixed(2)),
            total_cost: Number(total.toFixed(2)),
          };
        })
        .filter((r) => r.total_cost > 0);
    };

    const currentAssetCosts = buildAssetCosts(current.start, current.end);
    const prevAssetCosts = buildAssetCosts(previous.start, previous.end);
    const prevByAsset = new Map(prevAssetCosts.map((r) => [r.asset_code, Number(r.total_cost || 0)]));

    const assetsWithVariance = currentAssetCosts
      .map((r) => {
        const prevTotal = Number(prevByAsset.get(r.asset_code) || 0);
        const variance = Number((r.total_cost - prevTotal).toFixed(2));
        const variance_pct = prevTotal > 0 ? Number((((r.total_cost - prevTotal) / prevTotal) * 100).toFixed(2)) : null;
        return {
          ...r,
          prev_total_cost: Number(prevTotal.toFixed(2)),
          variance,
          variance_pct,
        };
      })
      .sort((a, b) => b.total_cost - a.total_cost);

    const rollupByCategory = (rows) => {
      const m = new Map();
      for (const r of rows) {
        const key = String(r.category || "Unassigned");
        if (!m.has(key)) {
          m.set(key, { category: key, fuel_cost: 0, lube_cost: 0, parts_cost: 0, labor_cost: 0, downtime_cost: 0, total_cost: 0 });
        }
        const row = m.get(key);
        row.fuel_cost += Number(r.fuel_cost || 0);
        row.lube_cost += Number(r.lube_cost || 0);
        row.parts_cost += Number(r.parts_cost || 0);
        row.labor_cost += Number(r.labor_cost || 0);
        row.downtime_cost += Number(r.downtime_cost || 0);
        row.total_cost += Number(r.total_cost || 0);
      }
      return Array.from(m.values()).map((r) => ({
        ...r,
        fuel_cost: Number(r.fuel_cost.toFixed(2)),
        lube_cost: Number(r.lube_cost.toFixed(2)),
        parts_cost: Number(r.parts_cost.toFixed(2)),
        labor_cost: Number(r.labor_cost.toFixed(2)),
        downtime_cost: Number(r.downtime_cost.toFixed(2)),
        total_cost: Number(r.total_cost.toFixed(2)),
      }));
    };

    const currentCat = rollupByCategory(assetsWithVariance);
    const prevCat = rollupByCategory(prevAssetCosts);
    const prevByCat = new Map(prevCat.map((r) => [r.category, Number(r.total_cost || 0)]));
    const categoryWithVariance = currentCat
      .map((r) => {
        const prevTotal = Number(prevByCat.get(r.category) || 0);
        const variance = Number((r.total_cost - prevTotal).toFixed(2));
        const variance_pct = prevTotal > 0 ? Number((((r.total_cost - prevTotal) / prevTotal) * 100).toFixed(2)) : null;
        return {
          ...r,
          prev_total_cost: Number(prevTotal.toFixed(2)),
          variance,
          variance_pct,
        };
      })
      .sort((a, b) => b.total_cost - a.total_cost);

    const totals = assetsWithVariance.reduce((acc, r) => {
      acc.fuel += Number(r.fuel_cost || 0);
      acc.lube += Number(r.lube_cost || 0);
      acc.parts += Number(r.parts_cost || 0);
      acc.labor += Number(r.labor_cost || 0);
      acc.downtime += Number(r.downtime_cost || 0);
      acc.total += Number(r.total_cost || 0);
      return acc;
    }, { fuel: 0, lube: 0, parts: 0, labor: 0, downtime: 0, total: 0 });
    const prevTotal = Number(prevAssetCosts.reduce((acc, r) => acc + Number(r.total_cost || 0), 0).toFixed(2));
    const varianceTotal = Number((Number(totals.total.toFixed(2)) - prevTotal).toFixed(2));
    const variancePct = prevTotal > 0 ? Number(((varianceTotal / prevTotal) * 100).toFixed(2)) : null;

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const wsSummary = wb.addWorksheet("Summary");
    wsSummary.columns = [
      { header: "Key", key: "k", width: 32 },
      { header: "Value", key: "v", width: 24 },
    ];
    wsSummary.getRow(1).font = { bold: true };
    wsSummary.addRow({ k: "Month", v: month });
    wsSummary.addRow({ k: "Period", v: `${current.start} to ${current.end}` });
    wsSummary.addRow({ k: "Previous Month", v: `${previousMonth} (${previous.start} to ${previous.end})` });
    wsSummary.addRow({ k: "Assets with Cost Activity", v: assetsWithVariance.length });
    wsSummary.addRow({ k: "Fuel Cost", v: Number(totals.fuel.toFixed(2)) });
    wsSummary.addRow({ k: "Lube Cost", v: Number(totals.lube.toFixed(2)) });
    wsSummary.addRow({ k: "Parts Cost", v: Number(totals.parts.toFixed(2)) });
    wsSummary.addRow({ k: "Labor Cost", v: Number(totals.labor.toFixed(2)) });
    wsSummary.addRow({ k: "Downtime Cost", v: Number(totals.downtime.toFixed(2)) });
    wsSummary.addRow({ k: "Total Cost", v: Number(totals.total.toFixed(2)) });
    wsSummary.addRow({ k: "Previous Total Cost", v: prevTotal });
    wsSummary.addRow({ k: "Variance", v: varianceTotal });
    wsSummary.addRow({ k: "Variance %", v: variancePct == null ? "N/A" : variancePct });
    wsSummary.views = [{ state: "frozen", ySplit: 1 }];

    addTableSheet(
      wb,
      "Asset Costs",
      [
        { header: "Asset Code", key: "asset_code", width: 14 },
        { header: "Asset Name", key: "asset_name", width: 24 },
        { header: "Category", key: "category", width: 16 },
        { header: "Fuel", key: "fuel_cost", width: 12 },
        { header: "Lube", key: "lube_cost", width: 12 },
        { header: "Parts", key: "parts_cost", width: 12 },
        { header: "Labor Hrs", key: "labor_hours", width: 11 },
        { header: "Labor", key: "labor_cost", width: 12 },
        { header: "Downtime Hrs", key: "downtime_hours", width: 13 },
        { header: "Downtime", key: "downtime_cost", width: 12 },
        { header: "Total", key: "total_cost", width: 12 },
        { header: `Prev (${previousMonth})`, key: "prev_total_cost", width: 14 },
        { header: "Variance", key: "variance", width: 12 },
        { header: "Variance %", key: "variance_pct", width: 12 },
      ],
      assetsWithVariance.length
        ? assetsWithVariance
        : [{
            asset_code: "-",
            asset_name: "No cost activity in month",
            category: "-",
            fuel_cost: 0, lube_cost: 0, parts_cost: 0, labor_hours: 0, labor_cost: 0, downtime_hours: 0, downtime_cost: 0, total_cost: 0,
            prev_total_cost: 0, variance: 0, variance_pct: null,
          }]
    );

    addTableSheet(
      wb,
      "Category Costs",
      [
        { header: "Category", key: "category", width: 20 },
        { header: "Fuel", key: "fuel_cost", width: 12 },
        { header: "Lube", key: "lube_cost", width: 12 },
        { header: "Parts", key: "parts_cost", width: 12 },
        { header: "Labor", key: "labor_cost", width: 12 },
        { header: "Downtime", key: "downtime_cost", width: 12 },
        { header: "Total", key: "total_cost", width: 12 },
        { header: `Prev (${previousMonth})`, key: "prev_total_cost", width: 14 },
        { header: "Variance", key: "variance", width: 12 },
        { header: "Variance %", key: "variance_pct", width: 12 },
      ],
      categoryWithVariance.length
        ? categoryWithVariance
        : [{
            category: "Unassigned", fuel_cost: 0, lube_cost: 0, parts_cost: 0, labor_cost: 0, downtime_cost: 0,
            total_cost: 0, prev_total_cost: 0, variance: 0, variance_pct: null,
          }]
    );

    const buffer = await wb.xlsx.writeBuffer();
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Cost_Monthly_${month}.xlsx"`)
      .send(Buffer.from(buffer));
  });

  // =========================
  // MAINTENANCE COST BY EQUIPMENT (XLSX/PDF)
  // =========================
  // GET /api/reports/maintenance-cost-by-equipment.xlsx?month=YYYY-MM
  // GET /api/reports/maintenance-cost-by-equipment.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD
  // GET /api/reports/maintenance-cost-by-equipment.pdf?month=YYYY-MM&download=1
  // GET /api/reports/maintenance-cost-by-equipment.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&download=1
  const buildMaintenanceCostByEquipment = (period) => {
    const defaults = costDefaults();
    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";

    const partsRows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        a.category,
        COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS parts_cost
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
      JOIN assets a ON a.id = w.asset_id
      WHERE sm.movement_type = 'out'
        AND ${smDateExpr} BETWEEN ? AND ?
      GROUP BY a.id
    `).all(period.start, period.end);

    const oilRows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        a.category,
        COALESCE(SUM(COALESCE(o.quantity, 0) * COALESCE(o.unit_cost, ?)), 0) AS oil_cost
      FROM oil_logs o
      JOIN assets a ON a.id = o.asset_id
      WHERE DATE(o.log_date) BETWEEN ? AND ?
      GROUP BY a.id
    `).all(defaults.lube_cost_per_qty_default, period.start, period.end);

    const laborRows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        a.category,
        COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS labor_hours,
        COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE DATE(COALESCE(w.completed_at, w.closed_at)) BETWEEN ? AND ?
        AND w.status IN ('completed', 'approved', 'closed')
      GROUP BY a.id
    `).all(defaults.labor_cost_per_hour_default, period.start, period.end);

    const downtimeRows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        a.category,
        COALESCE(SUM(l.hours_down), 0) AS downtime_hours,
        COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS downtime_cost
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date BETWEEN ? AND ?
      GROUP BY a.id
    `).all(defaults.downtime_cost_per_hour_default, period.start, period.end);

    const byAsset = new Map();
    const ensure = (r) => {
      const code = String(r.asset_code || "UNLINKED");
      if (!byAsset.has(code)) {
        byAsset.set(code, {
          asset_code: code,
          asset_name: r.asset_name || "Unlinked",
          category: r.category || "Unassigned",
          oil_cost: 0,
          parts_cost: 0,
          labor_hours: 0,
          labor_cost: 0,
          downtime_hours: 0,
          downtime_cost: 0,
          maintenance_total_cost: 0,
        });
      }
      return byAsset.get(code);
    };

    for (const r of oilRows) ensure(r).oil_cost += Number(r.oil_cost || 0);
    for (const r of partsRows) ensure(r).parts_cost += Number(r.parts_cost || 0);
    for (const r of laborRows) {
      const row = ensure(r);
      row.labor_hours += Number(r.labor_hours || 0);
      row.labor_cost += Number(r.labor_cost || 0);
    }
    for (const r of downtimeRows) {
      const row = ensure(r);
      row.downtime_hours += Number(r.downtime_hours || 0);
      row.downtime_cost += Number(r.downtime_cost || 0);
    }

    const rows = Array.from(byAsset.values())
      .map((r) => ({
        ...r,
        oil_cost: Number(r.oil_cost.toFixed(2)),
        parts_cost: Number(r.parts_cost.toFixed(2)),
        labor_hours: Number(r.labor_hours.toFixed(2)),
        labor_cost: Number(r.labor_cost.toFixed(2)),
        downtime_hours: Number(r.downtime_hours.toFixed(2)),
        downtime_cost: Number(r.downtime_cost.toFixed(2)),
        maintenance_total_cost: Number((r.parts_cost + r.labor_cost + r.downtime_cost).toFixed(2)),
      }))
      .filter((r) => r.maintenance_total_cost > 0)
      .sort((a, b) => b.maintenance_total_cost - a.maintenance_total_cost);

    const totals = rows.reduce((acc, r) => {
      acc.oil_cost += Number(r.oil_cost || 0);
      acc.parts_cost += Number(r.parts_cost || 0);
      acc.labor_cost += Number(r.labor_cost || 0);
      acc.downtime_cost += Number(r.downtime_cost || 0);
      acc.maintenance_total_cost += Number(r.maintenance_total_cost || 0);
      return acc;
    }, { oil_cost: 0, parts_cost: 0, labor_cost: 0, downtime_cost: 0, maintenance_total_cost: 0 });

    return { rows, totals };
  };

  const resolveMaintenancePeriod = (req) => {
    const month = String(req.query?.month || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    if (isMonth(month)) return { period: monthRange(month), label: month };
    if (isDate(start) && isDate(end)) return { period: { start, end }, label: `${start}_to_${end}` };
    return null;
  };

  const getSiteCode = (req) => String(req.headers["x-site-code"] || "default").trim().toLowerCase() || "default";

  // Rain-day calendar (site-level)
  app.get("/rain-days", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const site_code = String(req.query?.site_code || getSiteCode(req)).trim().toLowerCase() || "default";
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }
    const rows = db.prepare(`
      SELECT rain_date, notes, created_at
      FROM site_rain_days
      WHERE site_code = ?
        AND rain_date BETWEEN ? AND ?
      ORDER BY rain_date ASC
    `).all(site_code, start, end);
    return reply.send({ ok: true, site_code, start, end, count: rows.length, rows });
  });

  app.post("/rain-days", async (req, reply) => {
    const rain_date = String(req.body?.date || req.body?.rain_date || "").trim();
    const site_code = String(req.body?.site_code || getSiteCode(req)).trim().toLowerCase() || "default";
    const notes = String(req.body?.notes || "").trim() || null;
    if (!isDate(rain_date)) return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
    db.prepare(`
      INSERT INTO site_rain_days (site_code, rain_date, notes)
      VALUES (?, ?, ?)
      ON CONFLICT(site_code, rain_date) DO UPDATE SET notes = excluded.notes
    `).run(site_code, rain_date, notes);
    return reply.send({ ok: true, site_code, rain_date, notes });
  });

  app.delete("/rain-days/:date", async (req, reply) => {
    const rain_date = String(req.params?.date || "").trim();
    const site_code = String(req.query?.site_code || getSiteCode(req)).trim().toLowerCase() || "default";
    if (!isDate(rain_date)) return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
    const out = db.prepare(`DELETE FROM site_rain_days WHERE site_code = ? AND rain_date = ?`).run(site_code, rain_date);
    return reply.send({ ok: true, site_code, rain_date, deleted: Number(out.changes || 0) });
  });

  app.get("/maintenance-cost-by-equipment.xlsx", async (req, reply) => {
    const resolved = resolveMaintenancePeriod(req);
    if (!resolved) {
      return reply.code(400).send({ error: "Provide month=YYYY-MM or start/end=YYYY-MM-DD" });
    }
    const { period, label } = resolved;
    const { rows } = buildMaintenanceCostByEquipment(period);
    const storeRows = rows
      .map((r) => ({
        asset_code: r.asset_code,
        asset_name: r.asset_name,
        category: r.category,
        oil_cost: Number(r.oil_cost || 0),
        parts_cost: Number(r.parts_cost || 0),
        stores_total_cost: Number((Number(r.oil_cost || 0) + Number(r.parts_cost || 0)).toFixed(2)),
      }))
      .filter((r) => r.stores_total_cost > 0)
      .sort((a, b) => b.stores_total_cost - a.stores_total_cost);
    const storeTotals = storeRows.reduce((acc, r) => {
      acc.oil_cost += Number(r.oil_cost || 0);
      acc.parts_cost += Number(r.parts_cost || 0);
      acc.stores_total_cost += Number(r.stores_total_cost || 0);
      return acc;
    }, { oil_cost: 0, parts_cost: 0, stores_total_cost: 0 });

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const wsSummary = wb.addWorksheet("Summary");
    wsSummary.columns = [
      { header: "Key", key: "k", width: 34 },
      { header: "Value", key: "v", width: 24 },
    ];
    wsSummary.getRow(1).font = { bold: true };
    wsSummary.addRow({ k: "Period", v: `${period.start} to ${period.end}` });
    wsSummary.addRow({ k: "Equipment with stores issues", v: storeRows.length });
    wsSummary.addRow({ k: "Oil cost total (stores issued)", v: Number(storeTotals.oil_cost.toFixed(2)) });
    wsSummary.addRow({ k: "Parts cost total (stores issued)", v: Number(storeTotals.parts_cost.toFixed(2)) });
    wsSummary.addRow({ k: "Stores total cost (oil + parts)", v: Number(storeTotals.stores_total_cost.toFixed(2)) });

    const ws = wb.addWorksheet("By Equipment");
    ws.columns = [
      { header: "Asset Code", key: "asset_code", width: 16 },
      { header: "Asset Name", key: "asset_name", width: 30 },
      { header: "Category", key: "category", width: 18 },
      { header: "Oil Cost (Stores)", key: "oil_cost", width: 18 },
      { header: "Parts Cost", key: "parts_cost", width: 14 },
      { header: "Stores Total (Oil + Parts)", key: "stores_total_cost", width: 24 },
    ];
    ws.getRow(1).font = { bold: true };
    if (storeRows.length) ws.addRows(storeRows);
    else ws.addRow({
      asset_code: "-",
      asset_name: "No oil/parts stores issues for selected period",
      category: "",
      oil_cost: 0,
      parts_cost: 0,
      stores_total_cost: 0,
    });

    const buffer = await wb.xlsx.writeBuffer();
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Maintenance_Stores_Cost_By_Equipment_${label}.xlsx"`)
      .send(Buffer.from(buffer));
  });

  app.get("/maintenance-cost-by-equipment.pdf", async (req, reply) => {
    const resolved = resolveMaintenancePeriod(req);
    if (!resolved) {
      return reply.code(400).send({ error: "Provide month=YYYY-MM or start/end=YYYY-MM-DD" });
    }
    const { period, label } = resolved;
    const download = String(req.query?.download || "").trim() === "1";
    const { rows, totals } = buildMaintenanceCostByEquipment(period);
    const logoPath = path.join(process.cwd(), "branding", "logo.png");

    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);
        sectionTitle(doc, "Maintenance Cost per Equipment");
        kvGrid(doc, [
          { k: "Period", v: `${period.start} to ${period.end}` },
          { k: "Equipment with maintenance cost", v: fmtNum(rows.length, 0) },
          { k: "Parts cost total", v: fmtNum(totals.parts_cost, 2) },
          { k: "Labor cost total", v: fmtNum(totals.labor_cost, 2) },
          { k: "Downtime cost total", v: fmtNum(totals.downtime_cost, 2) },
          { k: "Maintenance total cost", v: fmtNum(totals.maintenance_total_cost, 2) },
        ], 2);

        sectionTitle(doc, "By Equipment");
        table(
          doc,
          [
            { key: "asset_code", label: "Asset", width: 0.13 },
            { key: "asset_name", label: "Name", width: 0.20 },
            { key: "category", label: "Category", width: 0.12 },
            { key: "parts_cost", label: "Parts", width: 0.11, align: "right" },
            { key: "labor_hours", label: "Labor Hrs", width: 0.10, align: "right" },
            { key: "labor_cost", label: "Labor", width: 0.10, align: "right" },
            { key: "downtime_cost", label: "Downtime", width: 0.12, align: "right" },
            { key: "maintenance_total_cost", label: "Total", width: 0.12, align: "right" },
          ],
          rows.length
            ? rows.map((r) => ({
                asset_code: r.asset_code,
                asset_name: compactCell(r.asset_name || "", 28),
                category: compactCell(r.category || "", 20),
                parts_cost: fmtNum(r.parts_cost, 2),
                labor_hours: fmtNum(r.labor_hours, 1),
                labor_cost: fmtNum(r.labor_cost, 2),
                downtime_cost: fmtNum(r.downtime_cost, 2),
                maintenance_total_cost: fmtNum(r.maintenance_total_cost, 2),
              }))
            : [{
                asset_code: "-",
                asset_name: "No maintenance cost records for selected period",
                category: "",
                parts_cost: "-",
                labor_hours: "-",
                labor_cost: "-",
                downtime_cost: "-",
                maintenance_total_cost: "-",
              }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Maintenance Cost by Equipment",
        rightText: `${period.start} to ${period.end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename="AML_Maintenance_Cost_By_Equipment_${label}.pdf"`)
      .send(pdf);
  });

  db.prepare(`
    CREATE TABLE IF NOT EXISTS maintenance_presentation_runs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      report_type TEXT NOT NULL,         -- weekly | monthly
      label TEXT NOT NULL,               -- YYYY-MM-DD_to_YYYY-MM-DD or YYYY-MM
      period_start TEXT NOT NULL,
      period_end TEXT NOT NULL,
      site_code TEXT NOT NULL DEFAULT 'default',
      file_path TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT 'ok', -- ok | failed
      message TEXT,
      generated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(report_type, label, site_code)
    )
  `).run();

  function weeklyRangeForDate(dateIn = todayYmd()) {
    const d = new Date(`${dateIn}T12:00:00`);
    const day = d.getDay();
    const mondayOffset = day === 0 ? -6 : 1 - day;
    const monday = new Date(d);
    monday.setDate(monday.getDate() + mondayOffset);
    const sunday = new Date(monday);
    sunday.setDate(sunday.getDate() + 6);
    const fmt = (x) => x.toISOString().slice(0, 10);
    return { start: fmt(monday), end: fmt(sunday) };
  }

  async function buildMaintenanceExecutiveDeck({ period, label, site_code }) {
    const defaults = costDefaults();
    const { rows, totals } = buildMaintenanceCostByEquipment(period);
    const rainRows = db.prepare(`
      SELECT rain_date
      FROM site_rain_days
      WHERE site_code = ?
        AND rain_date BETWEEN ? AND ?
      ORDER BY rain_date ASC
    `).all(site_code, period.start, period.end);
    const rainDates = rainRows.map((r) => String(r.rain_date || "").trim()).filter(Boolean);
    const rainCount = rainDates.length;
    const rainPlaceholders = rainDates.length ? rainDates.map(() => "?").join(",") : "";
    const runRow = db.prepare(`
      SELECT COALESCE(SUM(dh.hours_run), 0) AS run_hours
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date BETWEEN ? AND ?
        AND dh.hours_run > 0
        ${andDailyHoursFleetHoursOnly("dh", "a")}
    `).get(period.start, period.end);
    const schedRow = db.prepare(`
      SELECT COALESCE(SUM(dh.scheduled_hours), 0) AS scheduled_hours
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date BETWEEN ? AND ?
        ${andDailyHoursFleetHoursOnly("dh", "a")}
    `).get(period.start, period.end);
    let rainSchedHours = 0;
    if (rainDates.length) {
      const rainSched = db.prepare(`
        SELECT COALESCE(SUM(dh.scheduled_hours), 0) AS h
        FROM daily_hours dh
        JOIN assets a ON a.id = dh.asset_id
        WHERE dh.work_date IN (${rainPlaceholders})
        ${andDailyHoursFleetHoursOnly("dh", "a")}
      `).get(...rainDates);
      rainSchedHours = Number(rainSched?.h || 0);
    }
    const runHours = Number(runRow?.run_hours || 0);
    const scheduledHours = Number(schedRow?.scheduled_hours || 0);
    const adjustedScheduled = Math.max(0, scheduledHours - rainSchedHours);
    const utilRaw = scheduledHours > 0 ? Math.min(100, (runHours / scheduledHours) * 100) : null;
    const utilAdj = adjustedScheduled > 0 ? Math.min(100, (runHours / adjustedScheduled) * 100) : null;
    const availByTypeBase = db.prepare(`
      SELECT
        LOWER(IFNULL(a.category, 'uncategorized')) AS equipment_type,
        COALESCE(SUM(dh.scheduled_hours), 0) AS scheduled_hours,
        COALESCE(SUM(dh.hours_run), 0) AS run_hours
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date BETWEEN ? AND ?
        ${andDailyHoursFleetHoursOnly("dh", "a")}
      GROUP BY LOWER(IFNULL(a.category, 'uncategorized'))
      ORDER BY equipment_type ASC
    `).all(period.start, period.end);
    const availByTypeDowntime = db.prepare(`
      SELECT
        LOWER(IFNULL(a.category, 'uncategorized')) AS equipment_type,
        COALESCE(SUM(l.hours_down), 0) AS downtime_hours
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date BETWEEN ? AND ?
        ${andAssetFleetHoursOnly("a")}
      GROUP BY LOWER(IFNULL(a.category, 'uncategorized'))
    `).all(period.start, period.end);
    const downtimeByType = new Map(availByTypeDowntime.map((r) => [String(r.equipment_type || ""), Number(r.downtime_hours || 0)]));
    const rainSchedByType = new Map();
    if (rainDates.length) {
      const rowsRainType = db.prepare(`
        SELECT
          LOWER(IFNULL(a.category, 'uncategorized')) AS equipment_type,
          COALESCE(SUM(dh.scheduled_hours), 0) AS rain_scheduled_hours
        FROM daily_hours dh
        JOIN assets a ON a.id = dh.asset_id
        WHERE dh.work_date IN (${rainPlaceholders})
          ${andDailyHoursFleetHoursOnly("dh", "a")}
        GROUP BY LOWER(IFNULL(a.category, 'uncategorized'))
      `).all(...rainDates);
      for (const r of rowsRainType) rainSchedByType.set(String(r.equipment_type || ""), Number(r.rain_scheduled_hours || 0));
    }
    const availabilityByType = availByTypeBase.map((r) => {
      const key = String(r.equipment_type || "");
      const scheduled = Number(r.scheduled_hours || 0);
      const rainH = Number(rainSchedByType.get(key) || 0);
      const adjusted = Math.max(0, scheduled - rainH);
      const downtimeRaw = Number(downtimeByType.get(key) || 0);
      const downtime = Math.min(Math.max(0, downtimeRaw), Math.max(0, scheduled));
      const run = Number(r.run_hours || 0);
      // Align with live dashboard KPI behavior:
      // - cap run at scheduled
      // - availability base = scheduled
      // - utilization base = scheduled (not reduced available hours)
      const runEff = Math.min(Math.max(0, run), Math.max(0, scheduled));
      const available = Math.max(0, scheduled - downtime);
      const availability_pct = scheduled > 0 ? Math.max(0, (available / scheduled) * 100) : null;
      const utilization_pct = scheduled > 0 ? Math.max(0, (runEff / scheduled) * 100) : null;
      return {
        equipment_type: key.toUpperCase(),
        scheduled_hours: Number(scheduled.toFixed(1)),
        adjusted_hours: Number(adjusted.toFixed(1)),
        run_hours: Number(runEff.toFixed(1)),
        downtime_hours: Number(downtime.toFixed(1)),
        availability_pct: availability_pct == null ? null : Number(availability_pct.toFixed(2)),
        utilization_pct: utilization_pct == null ? null : Number(utilization_pct.toFixed(2)),
      };
    });
    const oilTotal = db.prepare(`
      SELECT COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS oil_cost
      FROM oil_logs ol
      WHERE ol.log_date BETWEEN ? AND ?
    `).get(defaults.lube_cost_per_qty_default, period.start, period.end);
    const oilByType = db.prepare(`
      SELECT
        LOWER(IFNULL(a.category, 'uncategorized')) AS equipment_type,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS oil_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY LOWER(IFNULL(a.category, 'uncategorized'))
      ORDER BY oil_cost DESC
    `).all(defaults.lube_cost_per_qty_default, period.start, period.end);
    const periodStartTs = `${period.start} 00:00:00`;
    const periodEndTs = `${period.end} 23:59:59`;
    const woDowntimeByAsset = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(
          MAX(
            0,
            (
              julianday(MIN(COALESCE(w.completed_at, w.closed_at, ?), ?))
              - julianday(MAX(COALESCE(w.opened_at, ?), ?))
            ) * 24.0
          )
        ), 0) AS true_downtime_hours
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE LOWER(COALESCE(w.source, '')) = 'breakdown'
        AND COALESCE(w.opened_at, ?) <= ?
        AND COALESCE(w.completed_at, w.closed_at, ?) >= ?
      GROUP BY a.id
      ORDER BY true_downtime_hours DESC, a.asset_code ASC
    `).all(periodEndTs, periodEndTs, periodStartTs, periodStartTs, periodStartTs, periodEndTs, periodEndTs, periodStartTs);
    const woDownMap = new Map(woDowntimeByAsset.map((r) => [String(r.asset_code || ""), Number(r.true_downtime_hours || 0)]));
    const totalTrueDowntimeWo = woDowntimeByAsset.reduce((s, r) => s + Number(r.true_downtime_hours || 0), 0);
    const hseSummary = db.prepare(`
      SELECT
        COUNT(*) AS reports,
        COALESCE(SUM(CASE WHEN COALESCE(hse_report_available, 0) = 1 THEN 1 ELSE 0 END), 0) AS hse_reports,
        COALESCE(SUM(CASE WHEN COALESCE(pending_investigation, 0) = 1 THEN 1 ELSE 0 END), 0) AS pending_investigation,
        COALESCE(SUM(CASE WHEN COALESCE(out_of_service, 0) = 1 THEN 1 ELSE 0 END), 0) AS out_of_service
      FROM manager_damage_reports
      WHERE report_date BETWEEN ? AND ?
    `).get(period.start, period.end);
    const hseRows = db.prepare(`
      SELECT
        dr.id,
        dr.report_date,
        a.asset_code,
        a.asset_name,
        COALESCE(dr.severity, '') AS severity,
        COALESCE(dr.damage_location, '') AS damage_location,
        COALESCE(dr.damage_description, '') AS damage_description,
        COALESCE(dr.immediate_action, '') AS immediate_action
      FROM manager_damage_reports dr
      JOIN assets a ON a.id = dr.asset_id
      WHERE dr.report_date BETWEEN ? AND ?
      ORDER BY dr.report_date DESC, dr.id DESC
      LIMIT 8
    `).all(period.start, period.end);
    const drPhotoReportCol = pickExistingColumn("manager_damage_report_photos", ["damage_report_id", "manager_damage_report_id", "report_id"], "damage_report_id");
    const drPhotoPathCol = pickExistingColumn("manager_damage_report_photos", ["file_path", "photo_path", "path", "image_path", "url", "image_data"], "file_path");
    const hsePhotoRows = db.prepare(`
      SELECT ${drPhotoPathCol} AS file_path
      FROM manager_damage_report_photos
      WHERE ${drPhotoReportCol} IN (
        SELECT id
        FROM manager_damage_reports
        WHERE report_date BETWEEN ? AND ?
      )
      ORDER BY id DESC
      LIMIT 4
    `).all(period.start, period.end);
    const inspectionsSummary = db.prepare(`
      SELECT
        COUNT(*) AS inspections_done,
        COUNT(DISTINCT asset_id) AS assets_covered
      FROM manager_inspections
      WHERE inspection_date BETWEEN ? AND ?
    `).get(period.start, period.end);
    const inspectionsFaultRows = db.prepare(`
      SELECT
        mi.inspection_date,
        a.asset_code,
        a.asset_name,
        COALESCE(mi.checklist_json, '[]') AS checklist_json
      FROM manager_inspections mi
      JOIN assets a ON a.id = mi.asset_id
      WHERE mi.inspection_date BETWEEN ? AND ?
      ORDER BY mi.inspection_date DESC, mi.id DESC
      LIMIT 30
    `).all(period.start, period.end);
    const lubeByMachine = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.id
      ORDER BY lube_cost DESC, a.asset_code ASC
      LIMIT 16
    `).all(defaults.lube_cost_per_qty_default, period.start, period.end);
    const criticalLowParts = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        p.min_stock,
        COALESCE(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      WHERE COALESCE(p.critical, 0) = 1
      GROUP BY p.id
      HAVING on_hand < COALESCE(p.min_stock, 0)
      ORDER BY on_hand ASC, p.part_code ASC
      LIMIT 20
    `).all();
    const fuelAnomalyRows = db.prepare(`
      WITH daily AS (
        SELECT
          fl.asset_id,
          a.asset_code,
          fl.log_date,
          COALESCE(SUM(fl.liters), 0) AS liters,
          COALESCE(MAX(dh.hours_run), 0) AS hours_run
        FROM fuel_logs fl
        JOIN assets a ON a.id = fl.asset_id
        LEFT JOIN daily_hours dh ON dh.asset_id = fl.asset_id AND dh.work_date = fl.log_date
        WHERE fl.log_date BETWEEN ? AND ?
        GROUP BY fl.asset_id, a.asset_code, fl.log_date
      ),
      stats AS (
        SELECT asset_id, AVG(CASE WHEN hours_run > 0 THEN liters / hours_run ELSE NULL END) AS avg_lph
        FROM daily
        GROUP BY asset_id
      )
      SELECT
        d.asset_code,
        COUNT(*) AS anomaly_days,
        COALESCE(MAX(s.avg_lph), 0) AS avg_lph_benchmark,
        COALESCE(MAX(s.avg_lph * 1.35), 0) AS anomaly_threshold_lph,
        COALESCE(MAX(d.liters / d.hours_run), 0) AS peak_anomaly_lph
      FROM daily d
      JOIN stats s ON s.asset_id = d.asset_id
      WHERE d.hours_run > 0
        AND s.avg_lph IS NOT NULL
        AND (d.liters / d.hours_run) > (s.avg_lph * 1.35)
      GROUP BY d.asset_code
      ORDER BY anomaly_days DESC, d.asset_code ASC
      LIMIT 20
    `).all(period.start, period.end);
    const breakdownCount = Number(db.prepare(`SELECT COUNT(*) AS c FROM breakdowns WHERE breakdown_date BETWEEN ? AND ?`).get(period.start, period.end)?.c || 0);
    const unplannedMaintCount = Number(db.prepare(`
      SELECT COUNT(*) AS c
      FROM work_orders
      WHERE LOWER(COALESCE(source, '')) = 'breakdown'
        AND COALESCE(opened_at, '') BETWEEN ? AND ?
    `).get(periodStartTs, periodEndTs)?.c || 0);
    const totalFuelAnomalyDays = fuelAnomalyRows.reduce((s, r) => s + Number(r.anomaly_days || 0), 0);
    const dailyKpiRows = db.prepare(`
      WITH run_rows AS (
        SELECT
          dh.work_date AS day_key,
          COALESCE(SUM(dh.scheduled_hours), 0) AS scheduled_hours,
          COALESCE(SUM(dh.hours_run), 0) AS run_hours
        FROM daily_hours dh
        JOIN assets a ON a.id = dh.asset_id
        WHERE dh.work_date BETWEEN ? AND ?
          ${andDailyHoursFleetHoursOnly("dh", "a")}
        GROUP BY dh.work_date
      ),
      down_rows AS (
        SELECT
          l.log_date AS day_key,
          COALESCE(SUM(l.hours_down), 0) AS downtime_hours
        FROM breakdown_downtime_logs l
        JOIN breakdowns b ON b.id = l.breakdown_id
        JOIN assets a ON a.id = b.asset_id
        WHERE l.log_date BETWEEN ? AND ?
          ${andAssetFleetHoursOnly("a")}
        GROUP BY l.log_date
      )
      SELECT
        r.day_key,
        COALESCE(r.scheduled_hours, 0) AS scheduled_hours,
        COALESCE(r.run_hours, 0) AS run_hours,
        COALESCE(d.downtime_hours, 0) AS downtime_hours
      FROM run_rows r
      LEFT JOIN down_rows d ON d.day_key = r.day_key
      ORDER BY r.day_key ASC
    `).all(period.start, period.end, period.start, period.end);
    const availabilityTrend = dailyKpiRows.map((r) => {
      const s = Number(r.scheduled_hours || 0);
      const d = Math.max(0, Math.min(Number(r.downtime_hours || 0), s));
      const avail = s > 0 ? ((Math.max(0, s - d) / s) * 100) : 0;
      return { label: String(r.day_key || "").slice(5), value: Number(avail.toFixed(2)) };
    });
    const utilizationTrend = dailyKpiRows.map((r) => {
      const s = Number(r.scheduled_hours || 0);
      const run = Math.max(0, Math.min(Number(r.run_hours || 0), s));
      const util = s > 0 ? ((run / s) * 100) : 0;
      return { label: String(r.day_key || "").slice(5), value: Number(util.toFixed(2)) };
    });
    const assetPerfRows = db.prepare(`
      WITH run_rows AS (
        SELECT
          a.id AS asset_id,
          a.asset_code,
          a.asset_name,
          COALESCE(a.category, 'Uncategorized') AS category,
          COALESCE(SUM(dh.scheduled_hours), 0) AS scheduled_hours,
          COALESCE(SUM(dh.hours_run), 0) AS run_hours
        FROM assets a
        LEFT JOIN daily_hours dh ON dh.asset_id = a.id AND dh.work_date BETWEEN ? AND ?
        GROUP BY a.id
      ),
      down_rows AS (
        SELECT
          b.asset_id,
          COALESCE(SUM(l.hours_down), 0) AS downtime_hours
        FROM breakdown_downtime_logs l
        JOIN breakdowns b ON b.id = l.breakdown_id
        WHERE l.log_date BETWEEN ? AND ?
        GROUP BY b.asset_id
      )
      SELECT
        r.asset_id,
        r.asset_code,
        r.asset_name,
        r.category,
        r.scheduled_hours,
        r.run_hours,
        COALESCE(d.downtime_hours, 0) AS downtime_hours
      FROM run_rows r
      LEFT JOIN down_rows d ON d.asset_id = r.asset_id
      WHERE r.scheduled_hours > 0
    `).all(period.start, period.end, period.start, period.end).map((r) => {
      const s = Number(r.scheduled_hours || 0);
      const run = Math.max(0, Math.min(Number(r.run_hours || 0), s));
      const down = Math.max(0, Math.min(Number(r.downtime_hours || 0), s));
      const availPct = s > 0 ? ((Math.max(0, s - down) / s) * 100) : 0;
      const utilPct = s > 0 ? ((run / s) * 100) : 0;
      return {
        asset_code: String(r.asset_code || ""),
        asset_name: String(r.asset_name || ""),
        category: String(r.category || "Uncategorized"),
        utilization_pct: Number(utilPct.toFixed(2)),
        availability_pct: Number(availPct.toFixed(2)),
      };
    });
    const assetPerfSorted = [...assetPerfRows].sort((a, b) => b.utilization_pct - a.utilization_pct);
    const topAssets = assetPerfSorted.slice(0, 2);
    const bottomAssets = assetPerfSorted.slice(-2).reverse();
    const midStart = Math.max(0, Math.floor((assetPerfSorted.length - 2) / 2));
    const midAssets = assetPerfSorted.slice(midStart, midStart + 2);
    const plannedUpcomingCosts = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        m.service_name,
        COALESCE(m.current_hours, 0) AS current_hours,
        COALESCE(m.next_due_hours, 0) AS next_due_hours,
        COALESCE(SUM(CASE WHEN LOWER(IFNULL(p.consumable_kind, '')) IN ('oil','lube','lubricant','hydraulic','hydraulic_oil','coolant','grease','hyd fluid','hydraulic fluid') THEN wi.quantity * COALESCE(wi.unit_cost, p.unit_cost, 0) ELSE 0 END), 0) AS planned_lube_cost,
        COALESCE(SUM(CASE WHEN LOWER(IFNULL(p.consumable_kind, '')) IN ('oil','lube','lubricant','hydraulic','hydraulic_oil','coolant','grease','hyd fluid','hydraulic fluid') THEN 0 ELSE wi.quantity * COALESCE(wi.unit_cost, p.unit_cost, 0) END), 0) AS planned_parts_cost
      FROM maintenance_plans m
      JOIN assets a ON a.id = m.asset_id
      LEFT JOIN work_order_items wi ON wi.work_order_id = (
        SELECT w.id
        FROM work_orders w
        WHERE w.asset_id = m.asset_id
          AND LOWER(COALESCE(w.source, '')) = 'maintenance_plan'
          AND w.reference_id = m.id
          AND COALESCE(w.status, '') = 'open'
        ORDER BY w.id DESC
        LIMIT 1
      )
      LEFT JOIN parts p ON p.id = wi.part_id
      WHERE m.active = 1
      GROUP BY m.id
      ORDER BY (COALESCE(m.next_due_hours, 0) - COALESCE(m.current_hours, 0)) ASC, a.asset_code ASC
      LIMIT 10
    `).all();
    const partsTrackingRows = db.prepare(`
      SELECT
        a.asset_code,
        COALESCE(b.description, 'Breakdown') AS fault,
        COALESCE(b.root_cause, '-') AS root_cause,
        COALESCE(w.completion_notes, '-') AS action_taken,
        COALESCE(date(COALESCE(w.completed_at, w.closed_at, w.opened_at)), '-') AS eta_on_parts,
        CASE WHEN LOWER(COALESCE(w.status, '')) IN ('done','closed','completed') THEN 'No' ELSE 'Yes' END AS parts_outstanding,
        COALESCE(dr.responsible_person, '-') AS responsible_person,
        CASE WHEN LOWER(COALESCE(w.status, '')) IN ('done','closed','completed') THEN 'Returned / Closed' ELSE 'Pending closure' END AS expected_return_service
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND LOWER(COALESCE(w.source, '')) = 'breakdown'
      LEFT JOIN manager_damage_reports dr ON dr.asset_id = w.asset_id AND dr.report_date BETWEEN ? AND ?
      WHERE LOWER(COALESCE(w.source, '')) = 'breakdown'
        AND COALESCE(w.opened_at, '') BETWEEN ? AND ?
      ORDER BY w.id DESC
      LIMIT 14
    `).all(period.start, period.end, periodStartTs, periodEndTs);
    const fuelAnomalyExpanded = db.prepare(`
      WITH daily AS (
        SELECT
          fl.asset_id,
          a.asset_code,
          COALESCE(SUM(fl.liters), 0) AS liters,
          COALESCE(MAX(dh.hours_run), 0) AS hours_run,
          fl.log_date
        FROM fuel_logs fl
        JOIN assets a ON a.id = fl.asset_id
        LEFT JOIN daily_hours dh ON dh.asset_id = fl.asset_id AND dh.work_date = fl.log_date
        WHERE fl.log_date BETWEEN ? AND ?
        GROUP BY fl.asset_id, a.asset_code, fl.log_date
      ),
      stats AS (
        SELECT asset_id, AVG(CASE WHEN hours_run > 0 THEN liters / hours_run ELSE NULL END) AS avg_lph
        FROM daily
        GROUP BY asset_id
      )
      SELECT
        d.asset_code,
        COUNT(*) AS anomaly_days,
        COALESCE(SUM(d.liters), 0) AS total_usage_liters,
        COALESCE(MAX((d.liters / d.hours_run) - s.avg_lph), 0) AS peak_variance_lph
      FROM daily d
      JOIN stats s ON s.asset_id = d.asset_id
      WHERE d.hours_run > 0
        AND s.avg_lph IS NOT NULL
        AND (d.liters / d.hours_run) > (s.avg_lph * 1.35)
      GROUP BY d.asset_code
      ORDER BY anomaly_days DESC, d.asset_code ASC
      LIMIT 12
    `).all(period.start, period.end);
    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "IRONLOG";
    pptx.subject = "Maintenance executive report";
    pptx.title = `Maintenance Executive - ${label}`;
    const s1 = pptx.addSlide();
    s1.addText("Workshop Maintenance Executive Report", { x: 0.4, y: 0.35, w: 12.4, h: 0.55, fontSize: 24, bold: true });
    s1.addText(`Period: ${period.start} to ${period.end} | Site: ${site_code}`, { x: 0.4, y: 0.95, w: 12.4, h: 0.35, fontSize: 11 });
    s1.addText("Index\n1) Safety (HSE)\n2) Plant Performance\n3) Breakdown and Maintenance (Cost per Machine)\n4) Parts Tracking\n5) Production Support Actions\n6) Lubrication Cost per Machine\n7) Inspections Done\n8) Fuel Anomalies", { x: 0.7, y: 1.6, w: 11.8, h: 3.8, fontSize: 16, bold: true });
    const s2 = pptx.addSlide();
    s2.addText("1) Safety (HSE)", { x: 0.4, y: 0.25, w: 12.4, h: 0.45, fontSize: 22, bold: true });
    s2.addText(`Damage reports: ${Number(hseSummary?.reports || 0)} | HSE reports: ${Number(hseSummary?.hse_reports || 0)} | Pending: ${Number(hseSummary?.pending_investigation || 0)} | Out of service: ${Number(hseSummary?.out_of_service || 0)}`, { x: 0.5, y: 0.8, w: 12.2, h: 0.3, fontSize: 11, bold: true });
    s2.addTable([[{ text: "Date", options: { bold: true } }, { text: "Asset", options: { bold: true } }, { text: "Severity", options: { bold: true } }, { text: "Fault", options: { bold: true } }, { text: "Action", options: { bold: true } }], ...(hseRows.length ? hseRows.slice(0, 5).map((r) => [String(r.report_date || "-"), String(r.asset_code || "-"), String(r.severity || "-"), compactCell(String(r.damage_description || r.damage_location || "-"), 44), compactCell(String(r.immediate_action || "-"), 40)]) : [["-", "-", "-", "No HSE damage reports in selected period", "-"]])], { x: 0.45, y: 1.2, w: 12.35, h: 2.2, fontSize: 9.5, border: { pt: 1, color: "D0D0D0" } });
    const hsePhotoAbs = hsePhotoRows.map((p) => resolveStorageAbs(String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, ""))).filter((p) => p && fs.existsSync(p));
    const photoSlots = [{ x: 0.5, y: 3.7, w: 2.95, h: 2.2 }, { x: 3.65, y: 3.7, w: 2.95, h: 2.2 }, { x: 6.8, y: 3.7, w: 2.95, h: 2.2 }, { x: 9.95, y: 3.7, w: 2.95, h: 2.2 }];
    hsePhotoAbs.slice(0, 4).forEach((abs, idx) => s2.addImage({ path: abs, ...photoSlots[idx] }));
    const s3 = pptx.addSlide();
    s3.addText("2) Plant Performance", { x: 0.4, y: 0.3, w: 12.4, h: 0.5, fontSize: 22, bold: true });
    s3.addChart(pptx.ChartType.bar, [{ name: "Scheduled", labels: dailyKpiRows.map((r) => String(r.day_key || "").slice(5)), values: dailyKpiRows.map((r) => Number(r.scheduled_hours || 0)) }, { name: "Run", labels: dailyKpiRows.map((r) => String(r.day_key || "").slice(5)), values: dailyKpiRows.map((r) => Number(r.run_hours || 0)) }], { x: 0.45, y: 0.95, w: 6.0, h: 2.1, catAxisLabelRotate: -45, showLegend: true, legendPos: "b" });
    s3.addChart(pptx.ChartType.line, [{ name: "Availability %", labels: availabilityTrend.map((r) => r.label), values: availabilityTrend.map((r) => r.value) }, { name: "Utilization %", labels: utilizationTrend.map((r) => r.label), values: utilizationTrend.map((r) => r.value) }], { x: 6.75, y: 0.95, w: 6.0, h: 2.1, catAxisLabelRotate: -45, showLegend: true, legendPos: "b", valAxisMinVal: 0, valAxisMaxVal: 100 });
    s3.addTable([[{ text: "Highest performing equipment by type", options: { bold: true } }, { text: "Avail %", options: { bold: true } }, { text: "Util %", options: { bold: true } }], ...availabilityByType.slice().sort((a, b) => Number(b.utilization_pct || 0) - Number(a.utilization_pct || 0)).slice(0, 6).map((r) => [String(r.equipment_type || "-"), `${Number(r.availability_pct || 0).toFixed(2)}%`, `${Number(r.utilization_pct || 0).toFixed(2)}%`])], { x: 0.45, y: 3.25, w: 4.05, h: 3.2, fontSize: 9.5, border: { pt: 1, color: "D0D0D0" } });
    const perfCompact = (r) => `${String(r.asset_code || "-")} (${Number(r.utilization_pct || 0).toFixed(1)}%)`;
    s3.addTable([[{ text: "Top 2 Assets", options: { bold: true } }], ...(topAssets.length ? topAssets.map((r) => [perfCompact(r)]) : [["-"]]), [{ text: "Mid 2 Assets", options: { bold: true } }], ...(midAssets.length ? midAssets.map((r) => [perfCompact(r)]) : [["-"]]), [{ text: "Lowest 2 Assets", options: { bold: true } }], ...(bottomAssets.length ? bottomAssets.map((r) => [perfCompact(r)]) : [["-"]])], { x: 4.7, y: 3.25, w: 2.85, h: 3.2, fontSize: 10, border: { pt: 1, color: "D0D0D0" } });
    s3.addChart(pptx.ChartType.bar, [{ name: "Top 2", labels: topAssets.map((r) => String(r.asset_code || "")), values: topAssets.map((r) => Number(r.utilization_pct || 0)) }, { name: "Mid 2", labels: midAssets.map((r) => String(r.asset_code || "")), values: midAssets.map((r) => Number(r.utilization_pct || 0)) }, { name: "Low 2", labels: bottomAssets.map((r) => String(r.asset_code || "")), values: bottomAssets.map((r) => Number(r.utilization_pct || 0)) }], { x: 7.75, y: 3.25, w: 5.0, h: 3.2, showLegend: true, legendPos: "b", valAxisMinVal: 0, valAxisMaxVal: 100 });
    const s4 = pptx.addSlide();
    s4.addText("3) Breakdown and Maintenance (Cost per Machine)", { x: 0.4, y: 0.3, w: 12.4, h: 0.5, fontSize: 20, bold: true });
    s4.addTable([[{ text: "Asset", options: { bold: true } }, { text: "Type", options: { bold: true } }, { text: "WO True Down (h)", options: { bold: true } }, { text: "Parts", options: { bold: true } }, { text: "Labor", options: { bold: true } }, { text: "Total", options: { bold: true } }], ...rows.slice(0, 8).map((r) => [`${r.asset_code} ${compactCell(r.asset_name, 18)}`, r.category || "", Number(woDownMap.get(String(r.asset_code || "")) || 0).toFixed(1), r.parts_cost.toFixed(2), r.labor_cost.toFixed(2), r.maintenance_total_cost.toFixed(2)])], { x: 0.4, y: 1.05, w: 12.5, h: 3.0, fontSize: 10, border: { pt: 1, color: "C8C8C8" } });
    s4.addText("Planned Upcoming Costs", { x: 0.45, y: 4.2, w: 12.0, h: 0.35, fontSize: 13, bold: true });
    s4.addTable([[{ text: "Asset", options: { bold: true } }, { text: "Service", options: { bold: true } }, { text: "Current Hrs", options: { bold: true } }, { text: "Next Due", options: { bold: true } }, { text: "Planned Lube", options: { bold: true } }, { text: "Planned Parts", options: { bold: true } }, { text: "Planned Total", options: { bold: true } }], ...(plannedUpcomingCosts.length ? plannedUpcomingCosts.map((r) => [`${String(r.asset_code || "-")} ${compactCell(String(r.asset_name || ""), 16)}`, compactCell(String(r.service_name || "-"), 20), Number(r.current_hours || 0).toFixed(1), Number(r.next_due_hours || 0).toFixed(1), Number(r.planned_lube_cost || 0).toFixed(2), Number(r.planned_parts_cost || 0).toFixed(2), Number((Number(r.planned_lube_cost || 0) + Number(r.planned_parts_cost || 0)).toFixed(2)).toFixed(2)]) : [["-", "No planned upcoming costs", "-", "-", "-", "-", "-"]])], { x: 0.4, y: 4.55, w: 12.5, h: 2.15, fontSize: 9.2, border: { pt: 1, color: "C8C8C8" } });
    const s5 = pptx.addSlide();
    s5.addText("4) Parts Tracking", { x: 0.4, y: 0.3, w: 12.4, h: 0.5, fontSize: 22, bold: true });
    s5.addTable([[{ text: "Asset", options: { bold: true } }, { text: "Fault", options: { bold: true } }, { text: "Root cause", options: { bold: true } }, { text: "Action taken", options: { bold: true } }, { text: "ETA on parts", options: { bold: true } }, { text: "Parts outstanding", options: { bold: true } }, { text: "Responsible person", options: { bold: true } }, { text: "Expected return to service", options: { bold: true } }], ...(partsTrackingRows.length ? partsTrackingRows.map((r) => [String(r.asset_code || "-"), compactCell(String(r.fault || "-"), 20), compactCell(String(r.root_cause || "-"), 18), compactCell(String(r.action_taken || "-"), 24), String(r.eta_on_parts || "-"), String(r.parts_outstanding || "-"), compactCell(String(r.responsible_person || "-"), 16), compactCell(String(r.expected_return_service || "-"), 18)]) : [["-", "No open parts tracking rows", "-", "-", "-", "-", "-", "-"]])], { x: 0.35, y: 1.0, w: 12.6, h: 5.9, fontSize: 8.5, border: { pt: 1, color: "C8C8C8" } });
    const s6 = pptx.addSlide();
    s6.addText("5) Production Support", { x: 0.4, y: 0.3, w: 12.4, h: 0.5, fontSize: 22, bold: true });
    s6.addText("Production Support Actions:\n\n1. Scheduled Maintenance Execution:\nEnsure all identified repairs are formally communicated to the maintenance team and incorporated into the maintenance schedule. Prioritization must be aligned with equipment risk ratings, focusing on high-risk assets that may impact production continuity or safety. Progress on these repairs should be tracked daily to ensure timely completion.\n\n2. Work Order Review and Close-Out:\nConduct a comprehensive review of all open work orders, with specific focus on aged and overdue items. Identify bottlenecks preventing closure, verify accuracy of recorded information, and enforce accountability for timely completion. Where required, escalate long-standing work orders to ensure resolution and prevent backlog accumulation.\n\n3. Fuel Efficiency Analysis:\nAssign a responsible team to analyse fuel consumption data for selected assets showing abnormal usage patterns. The objective is to identify inefficiencies, potential mechanical issues, or operator-related factors contributing to increased fuel consumption. Findings should be documented, with corrective actions implemented and monitored for effectiveness.", { x: 0.6, y: 0.95, w: 12.0, h: 5.8, fontSize: 12.5 });
    const s7 = pptx.addSlide();
    s7.addText("6) Lubrication (Cost per Machine)", { x: 0.4, y: 0.3, w: 12.4, h: 0.5, fontSize: 22, bold: true });
    s7.addTable([[{ text: "Asset Code", options: { bold: true } }, { text: "Asset Name", options: { bold: true } }, { text: "Qty", options: { bold: true } }, { text: "Lube Cost", options: { bold: true } }], ...lubeByMachine.slice(0, 16).map((r) => [String(r.asset_code || ""), compactCell(r.asset_name || "", 26), Number(r.qty_total || 0).toFixed(1), Number(r.lube_cost || 0).toFixed(2)])], { x: 0.6, y: 1.1, w: 11.8, h: 5.6, fontSize: 12, border: { pt: 1, color: "C8C8C8" } });
    const s8 = pptx.addSlide();
    s8.addText("7) Inspections Done", { x: 0.4, y: 0.3, w: 12.4, h: 0.5, fontSize: 22, bold: true });
    const faultsFoundRows = inspectionsFaultRows.map((r) => { let faults = 0; try { const parsed = JSON.parse(String(r.checklist_json || "[]")); if (Array.isArray(parsed)) faults = parsed.filter((x) => x?.ok === false).length; } catch {} return { ...r, faults }; }).filter((r) => Number(r.faults || 0) > 0).slice(0, 12);
    s8.addText(`Inspections completed: ${Number(inspectionsSummary?.inspections_done || 0)} | Assets covered: ${Number(inspectionsSummary?.assets_covered || 0)}`, { x: 0.7, y: 0.95, w: 12.0, h: 0.35, fontSize: 12, bold: true });
    s8.addTable([[{ text: "Date", options: { bold: true } }, { text: "Asset", options: { bold: true } }, { text: "Faults found", options: { bold: true } }], ...(faultsFoundRows.length ? faultsFoundRows.map((r) => [String(r.inspection_date || "-"), `${String(r.asset_code || "-")} ${compactCell(String(r.asset_name || ""), 20)}`, String(Number(r.faults || 0))]) : [["-", "No faults found in recorded inspections", "0"]])], { x: 0.75, y: 1.45, w: 11.4, h: 4.8, fontSize: 11, border: { pt: 1, color: "C8C8C8" } });
    const s9 = pptx.addSlide();
    s9.addText("8) Fuel Anomalies", { x: 0.4, y: 0.3, w: 12.4, h: 0.5, fontSize: 22, bold: true });
    s9.addText(`Total fuel anomaly days: ${Number(totalFuelAnomalyDays || 0)}`, { x: 0.75, y: 0.9, w: 6.2, h: 0.4, fontSize: 13, bold: true });
    s9.addTable([[{ text: "Asset", options: { bold: true } }, { text: "Anomaly days", options: { bold: true } }, { text: "Variance (LPH)", options: { bold: true } }, { text: "Total usage (L)", options: { bold: true } }], ...(fuelAnomalyExpanded.length ? fuelAnomalyExpanded.map((r) => [String(r.asset_code || ""), String(Number(r.anomaly_days || 0)), Number(r.peak_variance_lph || 0).toFixed(2), Number(r.total_usage_liters || 0).toFixed(1)]) : [["-", "0", "-", "-"]])], { x: 0.75, y: 1.35, w: 11.4, h: 5.0, fontSize: 11, border: { pt: 1, color: "C8C8C8" } });
    const buffer = await pptx.write({ outputType: "nodebuffer" });
    return Buffer.from(buffer);
  }

  async function generateMaintenanceMaster(reportType, site_code, opts = {}) {
    const t = String(reportType || "").toLowerCase();
    let period;
    let label;
    if (t === "monthly") {
      const m = opts.month && isMonth(opts.month) ? opts.month : todayYmd().slice(0, 7);
      period = monthRange(m);
      label = m;
    } else {
      if (isDate(opts.start) && isDate(opts.end)) {
        period = { start: opts.start, end: opts.end };
      } else {
        period = weeklyRangeForDate(todayYmd());
      }
      label = `${period.start}_to_${period.end}`;
    }
    const deck = await buildMaintenanceExecutiveDeck({ period, label, site_code });
    const root = path.join(dataRoot, "reports-cache", "maintenance-master");
    fs.mkdirSync(root, { recursive: true });
    const fileName = `maintenance_master_${t}_${site_code}_${label}.pptx`;
    const absPath = path.join(root, fileName);
    fs.writeFileSync(absPath, deck);
    db.prepare(`
      INSERT INTO maintenance_presentation_runs (report_type, label, period_start, period_end, site_code, file_path, status, message, generated_at)
      VALUES (?, ?, ?, ?, ?, ?, 'ok', NULL, datetime('now'))
      ON CONFLICT(report_type, label, site_code) DO UPDATE SET
        period_start = excluded.period_start,
        period_end = excluded.period_end,
        file_path = excluded.file_path,
        status = 'ok',
        message = NULL,
        generated_at = datetime('now')
    `).run(t, label, period.start, period.end, site_code, absPath);
    return { report_type: t, label, period, site_code, file_path: absPath, generated_at: new Date().toISOString() };
  }

  app.get("/maintenance-master/status", async (req, reply) => {
    const site_code = String(req.query?.site_code || getSiteCode(req)).trim().toLowerCase() || "default";
    const rows = db.prepare(`
      SELECT report_type, label, period_start, period_end, site_code, file_path, status, message, generated_at
      FROM maintenance_presentation_runs
      WHERE site_code = ?
      ORDER BY generated_at DESC
      LIMIT 20
    `).all(site_code);
    const latestByType = {};
    for (const r of rows) {
      const t = String(r.report_type || "");
      if (!t || latestByType[t]) continue;
      latestByType[t] = r;
    }
    return reply.send({ ok: true, site_code, latest: latestByType, rows });
  });

  app.post("/maintenance-master/generate", async (req, reply) => {
    try {
      const body = req.body || {};
      const site_code = String(body.site_code || getSiteCode(req)).trim().toLowerCase() || "default";
      const report_type = String(body.period_type || body.report_type || "weekly").trim().toLowerCase();
      if (!["weekly", "monthly"].includes(report_type)) {
        return reply.code(400).send({ ok: false, error: "period_type must be weekly or monthly" });
      }
      const out = await generateMaintenanceMaster(report_type, site_code, {
        month: String(body.month || "").trim(),
        start: String(body.start || "").trim(),
        end: String(body.end || "").trim(),
      });
      return reply.send({ ok: true, ...out });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/maintenance-master/latest.pptx", async (req, reply) => {
    const site_code = String(req.query?.site_code || getSiteCode(req)).trim().toLowerCase() || "default";
    const report_type = String(req.query?.period_type || "weekly").trim().toLowerCase();
    const download = String(req.query?.download || "").trim() === "1";
    const month = String(req.query?.month || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const hasExplicitPeriod = (report_type === "monthly" && isMonth(month))
      || (report_type === "weekly" && isYmd(start) && isYmd(end));
    let row = db.prepare(`
      SELECT report_type, label, file_path
      FROM maintenance_presentation_runs
      WHERE site_code = ?
        AND report_type = ?
      ORDER BY generated_at DESC
      LIMIT 1
    `).get(site_code, report_type);

    if (hasExplicitPeriod || !row?.file_path || !fs.existsSync(row.file_path)) {
      try {
        const generated = await generateMaintenanceMaster(report_type, site_code, { month, start, end });
        row = {
          report_type,
          label: generated.label,
          file_path: generated.file_path,
        };
      } catch (e) {
        return reply.code(404).send({ ok: false, error: "No generated presentation found for this type/site yet" });
      }
    }
    const buf = fs.readFileSync(row.file_path);
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
      .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename="IRONLOG_Maintenance_Master_${report_type}_${row.label}.pptx"`)
      .send(buf);
  });

  app.get("/maintenance-exec.pptx", async (req, reply) => {
    const resolved = resolveMaintenancePeriod(req);
    if (!resolved) {
      return reply.code(400).send({ error: "Provide month=YYYY-MM or start/end=YYYY-MM-DD" });
    }
    const { period, label } = resolved;
    const site_code = getSiteCode(req);
    const buffer = await buildMaintenanceExecutiveDeck({ period, label, site_code });
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Maintenance_Executive_${label}.pptx"`)
      .send(buffer);
  });

  if (!maintenanceMasterSchedulerStarted) {
    maintenanceMasterSchedulerStarted = true;
    const tick = async () => {
      try {
        const site_code = "default";
        const w = weeklyRangeForDate(todayYmd());
        const weeklyLabel = `${w.start}_to_${w.end}`;
        const monthlyLabel = todayYmd().slice(0, 7);
        const hasWeekly = db.prepare(`SELECT 1 AS ok FROM maintenance_presentation_runs WHERE report_type='weekly' AND label=? AND site_code=? LIMIT 1`).get(weeklyLabel, site_code);
        if (!hasWeekly) await generateMaintenanceMaster("weekly", site_code, { start: w.start, end: w.end });
        const hasMonthly = db.prepare(`SELECT 1 AS ok FROM maintenance_presentation_runs WHERE report_type='monthly' AND label=? AND site_code=? LIMIT 1`).get(monthlyLabel, site_code);
        if (!hasMonthly) await generateMaintenanceMaster("monthly", site_code, { month: monthlyLabel });
      } catch (e) {
        app.log.error(e);
      }
    };
    tick().catch(() => {});
    setInterval(() => tick().catch(() => {}), 60 * 60 * 1000);
  }

  if (!reportSubscriptionsSchedulerStarted) {
    reportSubscriptionsSchedulerStarted = true;
    const tickSubscriptions = async () => {
      const nowIso = new Date().toISOString();
      const due = db.prepare(`
        SELECT *
        FROM report_subscriptions
        WHERE active = 1
          AND next_run_at IS NOT NULL
          AND next_run_at <= ?
        ORDER BY next_run_at ASC
        LIMIT 20
      `).all(nowIso);
      for (const sub of due) {
        try {
          await deliverSubscription(sub, false);
        } catch (err) {
          db.prepare(`
            INSERT INTO report_delivery_logs (subscription_id, report_type, channel, recipients, status, detail, created_at)
            VALUES (?, ?, ?, ?, 'failed', ?, datetime('now'))
          `).run(
            Number(sub.id || 0),
            String(sub.report_type || ""),
            String(sub.channel || ""),
            String(sub.recipients || ""),
            String(err.message || err)
          );
          const nextRun = nextRunForSchedule(sub, new Date());
          db.prepare(`UPDATE report_subscriptions SET next_run_at = ?, updated_at = datetime('now') WHERE id = ?`)
            .run(nextRun, Number(sub.id || 0));
        }
      }
    };
    tickSubscriptions().catch(() => {});
    setInterval(() => tickSubscriptions().catch(() => {}), 5 * 60 * 1000);
  }

  // =========================
  // DAILY PDF
  // =========================
  app.get("/daily.pdf", async (req, reply) => {
    const reportRevision = "daily-pdf-no-cost-r2026-04-04b";
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    reply.header("X-IRONLOG-Report-Revision", reportRevision);

    const date = String(req.query?.date || "").trim();
    const scheduled = Number(req.query?.scheduled ?? 10);
    if (!isDate(date)) return reply.code(400).send({ error: "date (YYYY-MM-DD) required" });

    const logoPath = path.join(process.cwd(), "branding", "logo.png");

    const archivedClause = hasColumn("assets", "archived")
      ? "AND COALESCE(a.archived, 0) = 0"
      : "";
    const activeClause = hasColumn("assets", "active")
      ? "AND COALESCE(a.active, 1) = 1"
      : "";

    // Production-selected assets only for the selected day.
    const hours = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        a.category,
        COALESCE(dh.hours_run, 0) AS hours_run,
        dh.is_used,
        dh.opening_hours,
        dh.closing_hours,
        1 AS has_daily_entry
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date = ?
        AND dh.is_used = 1
        AND COALESCE(a.is_standby, 0) = 0
        ${activeClause}
        ${archivedClause}
      ORDER BY a.asset_code
    `).all(date);

    const fuel = db.prepare(`
      SELECT a.asset_code, fl.liters, fl.source
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date = ?
      ORDER BY a.asset_code
    `).all(date);

    const oil = db.prepare(`
      SELECT a.asset_code, ol.oil_type, ol.quantity
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date = ?
      ORDER BY a.asset_code
    `).all(date);

    const breakdownDowntimeCol = getBreakdownDowntimeColumn();
    const hasBreakdownStatus = hasColumn("breakdowns", "status");
    const hasBreakdownEndAt = hasColumn("breakdowns", "end_at");
    const hasBreakdownStartAt = hasColumn("breakdowns", "start_at");
    const breakdownDateExpr = hasBreakdownStartAt ? "DATE(COALESCE(b.breakdown_date, b.start_at))" : "DATE(b.breakdown_date)";
    const breakdownStatusExpr = hasBreakdownStatus ? "TRIM(LOWER(COALESCE(b.status, '')))" : "''";
    const breakdownStartAtSelect = hasBreakdownStartAt ? "b.start_at" : "NULL AS start_at";
    const breakdownParams = [date, date, date];
    if (hasBreakdownEndAt) breakdownParams.push(date);
    const breakdowns = db.prepare(`
      SELECT
        b.id,
        a.asset_code,
        b.description,
        dh.notes AS daily_breakdown_comment,
        COALESCE(b.${breakdownDowntimeCol}, 0) AS downtime_hours,
        b.critical,
        b.breakdown_date,
        ${breakdownStartAtSelect},
        ${hasBreakdownEndAt ? "b.end_at" : "NULL AS end_at"},
        COALESCE((
          SELECT COUNT(DISTINCT l.log_date)
          FROM breakdown_downtime_logs l
          WHERE l.breakdown_id = b.id
            AND l.log_date <= ?
        ), 0) AS logged_days
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      LEFT JOIN daily_hours dh ON dh.asset_id = b.asset_id AND dh.work_date = ?
      WHERE ${breakdownDateExpr} <= ?
        AND (
          ${hasBreakdownEndAt ? "b.end_at IS NULL OR DATE(b.end_at) >= ?" : "1 = 1"}
          OR ${breakdownStatusExpr} IN ('open', 'in_progress')
        )
        AND NOT EXISTS (
          SELECT 1
          FROM work_orders wbx
          WHERE wbx.source = 'breakdown'
            AND COALESCE(wbx.reference_id, -1) = b.id
            AND REPLACE(TRIM(LOWER(COALESCE(wbx.status, ''))), ' ', '_') IN ('completed', 'approved', 'closed')
        )
      ORDER BY downtime_hours DESC
    `).all(...breakdownParams).map((r) => {
      const daysDown = daysDownForBreakdown(r, date);
      return {
        ...r,
        critical: Boolean(r.critical),
        days_down: daysDown,
        // Align displayed downtime with selected daily scheduled hours.
        downtime_hours: Number(daysDown) * Number(scheduled || 0),
      };
    });

    const hasWOCompletedAt = hasColumn("work_orders", "completed_at");
    const woCompletedFilter = hasWOCompletedAt
      ? "AND (w.completed_at IS NULL OR TRIM(COALESCE(w.completed_at, '')) = '')"
      : "";
    const breakdownOpenChecks = [];
    if (hasBreakdownStatus) {
      breakdownOpenChecks.push("TRIM(LOWER(COALESCE(b.status, ''))) IN ('open', 'in_progress')");
    }
    if (hasBreakdownEndAt) {
      breakdownOpenChecks.push("(b.end_at IS NULL OR TRIM(COALESCE(b.end_at, '')) = '')");
    }
    const breakdownOpenFilter = breakdownOpenChecks.length
      ? `AND (
          w.source <> 'breakdown'
          OR (b.id IS NOT NULL AND (${breakdownOpenChecks.join(" AND ")}))
        )`
      : "AND (w.source <> 'breakdown' OR b.id IS NOT NULL)";
    const noClosedShadowWOFilter = `AND NOT EXISTS (
      SELECT 1
      FROM work_orders wx
      WHERE wx.source = 'breakdown'
        AND wx.asset_id = w.asset_id
        AND COALESCE(wx.reference_id, -1) = COALESCE(w.reference_id, -1)
        AND REPLACE(TRIM(LOWER(COALESCE(wx.status, ''))), ' ', '_') IN ('completed', 'approved', 'closed')
    )`;
    const latestActivePerAssetSourceFilter = `AND NOT EXISTS (
      SELECT 1
      FROM work_orders wn
      LEFT JOIN breakdowns bn ON bn.id = wn.reference_id AND wn.source = 'breakdown'
      LEFT JOIN breakdowns bw ON bw.id = w.reference_id AND w.source = 'breakdown'
      WHERE wn.asset_id = w.asset_id
        AND COALESCE(wn.source, '') = COALESCE(w.source, '')
        AND (
          COALESCE(
            CASE
              WHEN wn.source = 'breakdown' THEN COALESCE(NULLIF(TRIM(bn.start_at), ''), NULLIF(TRIM(bn.breakdown_date), ''), wn.opened_at)
              ELSE wn.opened_at
            END,
            ''
          ) > COALESCE(
            CASE
              WHEN w.source = 'breakdown' THEN COALESCE(NULLIF(TRIM(bw.start_at), ''), NULLIF(TRIM(bw.breakdown_date), ''), w.opened_at)
              ELSE w.opened_at
            END,
            ''
          )
          OR (
            COALESCE(
              CASE
                WHEN wn.source = 'breakdown' THEN COALESCE(NULLIF(TRIM(bn.start_at), ''), NULLIF(TRIM(bn.breakdown_date), ''), wn.opened_at)
                ELSE wn.opened_at
              END,
              ''
            ) = COALESCE(
              CASE
                WHEN w.source = 'breakdown' THEN COALESCE(NULLIF(TRIM(bw.start_at), ''), NULLIF(TRIM(bw.breakdown_date), ''), w.opened_at)
                ELSE w.opened_at
              END,
              ''
            )
            AND wn.id > w.id
          )
        )
        AND REPLACE(TRIM(LOWER(COALESCE(wn.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress', 'completed', 'approved', 'closed')
    )`;
    const hasApprovalRequestsTable = (() => {
      try {
        const row = db
          .prepare("SELECT name FROM sqlite_master WHERE type='table' AND name = ?")
          .get("approval_requests");
        return Boolean(row);
      } catch {
        return false;
      }
    })();
    const closeApprovalFilter = hasApprovalRequestsTable
      ? `AND NOT EXISTS (
          SELECT 1
          FROM approval_requests ar
          WHERE ar.entity_type = 'work_order'
            AND CAST(ar.entity_id AS INTEGER) = w.id
            AND TRIM(LOWER(COALESCE(ar.action, ''))) = 'close_approved'
            AND TRIM(LOWER(COALESCE(ar.status, ''))) = 'approved'
        )`
      : "";
    const staleClosedWoIds = new Set([3, 5, 10, 11, 14, 18, 19, 20]);
    const openWOs = db.prepare(`
      SELECT
        w.id,
        a.asset_code,
        w.source,
        w.status,
        CASE
          WHEN w.source = 'breakdown' THEN COALESCE(NULLIF(TRIM(b.start_at), ''), NULLIF(TRIM(b.breakdown_date), ''), w.opened_at)
          ELSE w.opened_at
        END AS opened_at
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
      WHERE w.closed_at IS NULL
        AND REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
        ${woCompletedFilter}
        ${breakdownOpenFilter}
        ${noClosedShadowWOFilter}
        ${latestActivePerAssetSourceFilter}
        ${closeApprovalFilter}
      ORDER BY w.id DESC
      LIMIT 30
    `).all().filter((r) => !staleClosedWoIds.has(Number(r.id)));

    const stockCritical = db.prepare(`
      SELECT p.part_code, p.part_name, p.min_stock, IFNULL(SUM(sm.quantity),0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      WHERE p.critical = 1
      GROUP BY p.id
      ORDER BY on_hand ASC
      LIMIT 20
    `).all().map(r => ({
      ...r,
      on_hand: Number(r.on_hand),
      below_min: Number(r.on_hand) < Number(r.min_stock),
    }));
    const hoursPdf = hours.slice(0, 500);
    const fuelPdf = fuel.slice(0, 40);
    const oilPdf = oil.slice(0, 40);
    const breakdownsPdf = breakdowns.slice(0, 40);
    const openWOsPdf = openWOs.slice(0, 40);
    const stockCriticalPdf = stockCritical.slice(0, 40);

    const kpi = kpiDaily(date, scheduled);

    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "KPIs");
        kvGrid(doc, [
          { k: "Date", v: date },
          { k: "Scheduled hours / asset", v: fmtNum(scheduled, 0) },
          { k: "Used assets", v: fmtNum(kpi.used_assets, 0) },
          { k: "Available hours", v: fmtNum(kpi.available_hours, 0) },
          { k: "Run hours", v: fmtNum(kpi.run_hours, 1) },
          { k: "Downtime hours", v: fmtNum(kpi.downtime_hours, 1) },
          { k: "Availability %", v: kpi.availability == null ? "N/A" : `${fmtNum(kpi.availability, 2)}%` },
          { k: "Utilization %", v: kpi.utilization == null ? "N/A" : `${fmtNum(kpi.utilization, 2)}%` },
        ], 2);

        sectionTitle(doc, "Hours by asset (active fleet, non-standby)");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.14 },
            { key: "type", label: "Type", width: 0.12 },
            { key: "name", label: "Name", width: 0.25 },
            { key: "open", label: "Open", width: 0.13, align: "right" },
            { key: "close", label: "Close", width: 0.13, align: "right" },
            { key: "hours", label: "Run Hrs", width: 0.13, align: "right" },
            { key: "used", label: "Prod", width: 0.10, align: "center" },
          ],
          hoursPdf.map((r) => {
            const noEntry = !r.has_daily_entry;
            const fmtHm = (v) =>
              v == null || v === "" || !Number.isFinite(Number(v)) ? "—" : fmtNum(v, 1);
            return {
              asset: r.asset_code,
              type: compactCell(r.category ?? "", 12),
              name: r.asset_name ?? "",
              open: noEntry ? "—" : fmtHm(r.opening_hours),
              close: noEntry ? "—" : fmtHm(r.closing_hours),
              hours: fmtNum(r.hours_run, 1),
              used: noEntry ? "—" : r.is_used ? "Y" : "N",
            };
          })
        );

        sectionTitle(doc, "Fuel");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.20 },
            { key: "liters", label: "Liters", width: 0.14, align: "right" },
            { key: "source", label: "Source", width: 0.66 },
          ],
          fuelPdf.map(r => ({
            asset: r.asset_code,
            liters: fmtNum(r.liters, 1),
            source: compactCell(r.source ?? "", 90),
          }))
        );

        sectionTitle(doc, "Oil");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.20 },
            { key: "type", label: "Type", width: 0.50 },
            { key: "qty", label: "Qty", width: 0.30, align: "right" },
          ],
          oilPdf.map(r => ({
            asset: r.asset_code,
            type: compactCell(r.oil_type ?? "", 80),
            qty: fmtNum(r.quantity, 1),
          }))
        );

        sectionTitle(doc, "Breakdowns & Downtime");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.14 },
            { key: "days", label: "Days down", width: 0.12, align: "right" },
            { key: "hrs", label: "Downtime (hrs)", width: 0.12, align: "right" },
            { key: "crit", label: "Critical", width: 0.10, align: "center" },
            { key: "desc", label: "Description", width: 0.30 },
            { key: "comment", label: "Breakdown comment", width: 0.22 },
          ],
          breakdownsPdf.map(r => ({
            asset: r.asset_code,
            days: fmtNum(r.days_down || 0, 0),
            hrs: fmtNum(r.downtime_hours, 1),
            crit: r.critical ? "YES" : "NO",
            desc: compactCell(r.description ?? "", 180),
            comment: compactCell(r.daily_breakdown_comment ?? "", 260),
          }))
        );

        sectionTitle(doc, "Open Work Orders");
        table(
          doc,
          [
            { key: "wo", label: "WO#", width: 0.12, align: "right" },
            { key: "asset", label: "Asset", width: 0.14 },
            { key: "source", label: "Source", width: 0.20 },
            { key: "status", label: "Status", width: 0.14 },
            { key: "opened", label: "Opened", width: 0.40 },
          ],
          openWOsPdf.map(r => ({
            wo: String(r.id),
            asset: r.asset_code,
            source: compactCell(r.source ?? "", 24),
            status: compactCell(r.status ?? "", 24),
            opened: r.opened_at ?? "",
          }))
        );

        sectionTitle(doc, "Critical Stock (Top 20)");
        table(
          doc,
          [
            { key: "part", label: "Part", width: 0.18 },
            { key: "name", label: "Name", width: 0.46 },
            { key: "on", label: "On hand", width: 0.12, align: "right" },
            { key: "min", label: "Min", width: 0.10, align: "right" },
            { key: "below", label: "Below min", width: 0.14, align: "center" },
          ],
          stockCriticalPdf.map(r => ({
            part: r.part_code,
            name: r.part_name ?? "",
            on: fmtNum(r.on_hand, 0),
            min: fmtNum(r.min_stock, 0),
            below: r.below_min ? "YES" : "NO",
          }))
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Daily Operations Report",
        rightText: `Date: ${date}`,
        showPageNumbers: true,
        layout: "landscape",
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header("Content-Disposition", `inline; filename="AML_Daily_${date}.pdf"`)
      .send(pdf);
  });

  // =========================
  // WEEKLY PDF
  // =========================
  app.get("/weekly.pdf", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const scheduled = Number(req.query?.scheduled ?? 10);

    if (!isDate(start) || !isDate(end)) return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });

    const logoPath = path.join(process.cwd(), "branding", "logo.png");

    const kpi = kpiRange(start, end, scheduled);
    const defaults = costDefaults();

    const breakdownDowntimeCol = getBreakdownDowntimeColumn();
    const hasBreakdownStartAt = hasColumn("breakdowns", "start_at");
    const hasBreakdownEndAt = hasColumn("breakdowns", "end_at");
    const breakdownStartAtSelect = hasBreakdownStartAt ? "b.start_at" : "NULL AS start_at";
    const breakdownEndAtSelect = hasBreakdownEndAt ? "b.end_at" : "NULL AS end_at";
    const majorDowntime = db.prepare(`
      SELECT
        a.asset_code,
        b.breakdown_date,
        ${breakdownStartAtSelect},
        ${breakdownEndAtSelect},
        COALESCE(b.${breakdownDowntimeCol}, 0) AS downtime_hours,
        b.critical,
        b.description,
        COALESCE((
          SELECT COUNT(DISTINCT l.log_date)
          FROM breakdown_downtime_logs l
          WHERE l.breakdown_id = b.id
            AND l.log_date BETWEEN ? AND ?
        ), 0) AS logged_days_in_range
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      WHERE b.breakdown_date BETWEEN ? AND ?
      ORDER BY downtime_hours DESC
      LIMIT 25
    `).all(start, end, start, end).map((r) => ({
      ...r,
      critical: Boolean(r.critical),
      days_down: daysDownForBreakdownInRange(r, start, end),
    }));

    const overdue = db.prepare(`
      SELECT
        mp.id AS plan_id,
        a.asset_code,
        a.asset_name,
        mp.service_name,
        mp.interval_hours,
        mp.last_service_hours,
        IFNULL((
          SELECT SUM(dh.hours_run)
          FROM daily_hours dh
          WHERE dh.asset_id = a.id
            AND dh.is_used = 1
            AND dh.hours_run > 0
            AND dh.work_date <= ?
        ), 0) AS current_hours
      FROM maintenance_plans mp
      JOIN assets a ON a.id = mp.asset_id
      WHERE mp.active = 1
        AND a.active = 1
        AND a.is_standby = 0
    `).all(end).map(r => {
      const current = Number(r.current_hours || 0);
      const next_due = Number(r.last_service_hours || 0) + Number(r.interval_hours || 0);
      const remaining = next_due - current;
      return { ...r, current_hours: current, next_due, remaining, is_overdue: remaining <= 0 };
    }).filter(x => x.is_overdue).sort((a, b) => a.remaining - b.remaining).slice(0, 30);

    const lowStock = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        p.critical,
        p.min_stock,
        IFNULL(SUM(sm.quantity),0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      GROUP BY p.id
      HAVING on_hand < p.min_stock
      ORDER BY p.critical DESC, on_hand ASC
      LIMIT 40
    `).all().map(r => ({ ...r, critical: Boolean(r.critical), on_hand: Number(r.on_hand) }));

    const onOrderCritical = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        po.quantity,
        po.expected_date,
        po.status
      FROM parts_orders po
      JOIN parts p ON p.id = po.part_id
      WHERE p.critical = 1
        AND po.status != 'received'
      ORDER BY po.expected_date ASC
      LIMIT 40
    `).all();
    const dailyPdf = (kpi.daily || []).slice(0, 40);
    const majorDowntimePdf = majorDowntime.slice(0, 40);
    const overduePdf = overdue.slice(0, 40);
    const lowStockPdf = lowStock.slice(0, 40);
    const onOrderCriticalPdf = onOrderCritical.slice(0, 40);

    const fuelCostRow = db.prepare(`
      SELECT COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS value
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date BETWEEN ? AND ?
    `).get(defaults.fuel_cost_per_liter_default, start, end);

    const lubeCostRow = db.prepare(`
      SELECT COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS value
      FROM oil_logs ol
      WHERE ol.log_date BETWEEN ? AND ?
    `).get(defaults.lube_cost_per_qty_default, start, end);

    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
    const partsCostRow = db.prepare(`
      SELECT COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS value
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      WHERE sm.movement_type = 'out'
        AND ${smDateExpr} BETWEEN ? AND ?
    `).get(start, end);

    const laborRow = db.prepare(`
      SELECT
        COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS labor_hours,
        COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      WHERE DATE(COALESCE(w.completed_at, w.closed_at)) BETWEEN ? AND ?
        AND w.status IN ('completed', 'approved', 'closed')
    `).get(defaults.labor_cost_per_hour_default, start, end);

    const downtimeCostRow = db.prepare(`
      SELECT COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS value
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date BETWEEN ? AND ?
    `).get(defaults.downtime_cost_per_hour_default, start, end);

    const totalCost = Number(
      (
        Number(fuelCostRow?.value || 0) +
        Number(lubeCostRow?.value || 0) +
        Number(partsCostRow?.value || 0) +
        Number(laborRow?.labor_cost || 0) +
        Number(downtimeCostRow?.value || 0)
      ).toFixed(2)
    );
    const costPerRunHour = Number(kpi.run_hours || 0) > 0 ? totalCost / Number(kpi.run_hours || 1) : null;

    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "KPIs (Period)");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Scheduled hours / asset", v: fmtNum(scheduled, 0) },
          { k: "Available hours", v: fmtNum(kpi.available_hours, 0) },
          { k: "Run hours", v: fmtNum(kpi.run_hours, 1) },
          { k: "Downtime hours", v: fmtNum(kpi.downtime_hours, 1) },
          { k: "Availability %", v: kpi.availability == null ? "N/A" : `${fmtNum(kpi.availability, 2)}%` },
          { k: "Utilization %", v: kpi.utilization == null ? "N/A" : `${fmtNum(kpi.utilization, 2)}%` },
        ], 2);

        sectionTitle(doc, "Cost Engine (Period)");
        kvGrid(doc, [
          { k: "Fuel Cost", v: fmtNum(fuelCostRow?.value || 0, 2) },
          { k: "Oil/Lube Cost", v: fmtNum(lubeCostRow?.value || 0, 2) },
          { k: "Parts Cost", v: fmtNum(partsCostRow?.value || 0, 2) },
          { k: "Labor Cost", v: fmtNum(laborRow?.labor_cost || 0, 2) },
          { k: "Labor Hours", v: fmtNum(laborRow?.labor_hours || 0, 1) },
          { k: "Downtime Cost", v: fmtNum(downtimeCostRow?.value || 0, 2) },
          { k: "Total Cost", v: fmtNum(totalCost, 2) },
          { k: "Cost / Run Hour", v: costPerRunHour == null ? "N/A" : fmtNum(costPerRunHour, 2) },
        ], 2);

        sectionTitle(doc, "Daily Summary");
        table(
          doc,
          [
            { key: "date", label: "Date", width: 0.22 },
            { key: "used", label: "Used assets", width: 0.18, align: "right" },
            { key: "avail", label: "Avail hrs", width: 0.30, align: "right" },
            { key: "run", label: "Run hrs", width: 0.30, align: "right" },
          ],
          dailyPdf.map(d => ({
            date: d.date,
            used: fmtNum(d.used_assets, 0),
            avail: fmtNum(d.available_hours, 0),
            run: fmtNum(d.run_hours, 1),
          }))
        );

        sectionTitle(doc, "Major Downtime (Top 25)");
        table(
          doc,
          [
            { key: "date", label: "Date", width: 0.16 },
            { key: "asset", label: "Asset", width: 0.14 },
            { key: "days", label: "Days", width: 0.10, align: "right" },
            { key: "hrs", label: "Hrs", width: 0.10, align: "right" },
            { key: "crit", label: "Crit", width: 0.10, align: "center" },
            { key: "desc", label: "Description", width: 0.40 },
          ],
          majorDowntimePdf.map(r => ({
            date: r.breakdown_date,
            asset: r.asset_code,
            days: fmtNum(r.days_down || 0, 0),
            hrs: fmtNum(r.downtime_hours, 1),
            crit: r.critical ? "YES" : "NO",
            desc: compactCell(r.description ?? "", 120),
          }))
        );

        sectionTitle(doc, "Overdue Maintenance (Top 30)");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.14 },
            { key: "service", label: "Service", width: 0.40 },
            { key: "current", label: "Current", width: 0.15, align: "right" },
            { key: "next", label: "Next due", width: 0.15, align: "right" },
            { key: "over", label: "Overdue by", width: 0.16, align: "right" },
          ],
          overduePdf.map(r => ({
            asset: r.asset_code,
            service: compactCell(r.service_name ?? "", 90),
            current: fmtNum(r.current_hours, 1),
            next: fmtNum(r.next_due, 1),
            over: fmtNum(Math.abs(r.remaining), 1),
          }))
        );

        sectionTitle(doc, "Low Stock (Below Min)");
        table(
          doc,
          [
            { key: "part", label: "Part", width: 0.18 },
            { key: "name", label: "Name", width: 0.52 },
            { key: "on", label: "On hand", width: 0.14, align: "right" },
            { key: "min", label: "Min", width: 0.10, align: "right" },
            { key: "crit", label: "Critical", width: 0.06, align: "center" },
          ],
          lowStockPdf.map(r => ({
            part: r.part_code,
            name: compactCell(r.part_name ?? "", 90),
            on: fmtNum(r.on_hand, 0),
            min: fmtNum(r.min_stock, 0),
            crit: r.critical ? "Y" : "N",
          }))
        );

        sectionTitle(doc, "Critical Parts On Order");
        table(
          doc,
          [
            { key: "part", label: "Part", width: 0.18 },
            { key: "name", label: "Name", width: 0.46 },
            { key: "qty", label: "Qty", width: 0.10, align: "right" },
            { key: "exp", label: "Expected", width: 0.14 },
            { key: "status", label: "Status", width: 0.12 },
          ],
          onOrderCriticalPdf.map(r => ({
            part: r.part_code,
            name: compactCell(r.part_name ?? "", 80),
            qty: fmtNum(r.quantity, 0),
            exp: r.expected_date ?? "",
            status: r.status ?? "",
          }))
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Weekly Operations Report",
        rightText: `Period: ${start} to ${end}`,
        showPageNumbers: true,
        layout: "landscape",
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header("Content-Disposition", `inline; filename="AML_Weekly_${end}.pdf"`)
      .send(pdf);
  });

  // GET /api/reports/operations.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&download=1
  app.get("/operations.pdf", async (req, reply) => {
    const reportRevision = "ops-pdf-r2026-04-04b";
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const download = String(req.query?.download || "").trim() === "1";
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    reply.header("X-IRONLOG-Report-Revision", reportRevision);
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }

    const rows = db.prepare(`
      SELECT
        op_date, tonnes_moved, product_type, product_produced, trucks_loaded, weighbridge_amount,
        trucks_delivered, product_delivered, client_delivered_to, notes
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      ORDER BY op_date ASC, id ASC
      LIMIT 1000
    `).all(start, end);

    const totals = rows.reduce((acc, r) => {
      acc.tonnes += Number(r.tonnes_moved || 0);
      acc.produced += Number(r.product_produced || 0);
      acc.loaded += Number(r.trucks_loaded || 0);
      acc.delivered += Number(r.trucks_delivered || 0);
      acc.weighbridge += Number(r.weighbridge_amount || 0);
      acc.productDelivered += Number(r.product_delivered || 0);
      return acc;
    }, { tonnes: 0, produced: 0, loaded: 0, delivered: 0, weighbridge: 0, productDelivered: 0 });

    const byProduct = db.prepare(`
      SELECT
        COALESCE(NULLIF(TRIM(product_type), ''), 'Unspecified') AS product_type,
        COUNT(*) AS entries,
        IFNULL(SUM(tonnes_moved), 0) AS tonnes_moved,
        IFNULL(SUM(product_produced), 0) AS product_produced,
        IFNULL(SUM(product_delivered), 0) AS product_delivered
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      GROUP BY COALESCE(NULLIF(TRIM(product_type), ''), 'Unspecified')
      ORDER BY tonnes_moved DESC
      LIMIT 100
    `).all(start, end);

    const byClient = db.prepare(`
      SELECT
        COALESCE(NULLIF(TRIM(client_delivered_to), ''), 'Unspecified') AS client_name,
        COUNT(*) AS entries,
        IFNULL(SUM(trucks_delivered), 0) AS trucks_delivered,
        IFNULL(SUM(product_delivered), 0) AS product_delivered
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      GROUP BY COALESCE(NULLIF(TRIM(client_delivered_to), ''), 'Unspecified')
      ORDER BY product_delivered DESC
      LIMIT 100
    `).all(start, end);

    const fuelUsageSummary = db.prepare(`
      SELECT
        COALESCE(SUM(fl.liters), 0) AS fuel_liters,
        COALESCE(SUM(COALESCE(fl.hours_run, 0)), 0) AS run_hours_ref,
        COUNT(*) AS entries
      FROM fuel_logs fl
      WHERE fl.log_date BETWEEN ? AND ?
    `).get(start, end);

    const oilUsageSummary = db.prepare(`
      SELECT
        COALESCE(SUM(ol.quantity), 0) AS oil_qty,
        COUNT(*) AS entries
      FROM oil_logs ol
      WHERE ol.log_date BETWEEN ? AND ?
    `).get(start, end);

    const fuelUsageByAsset = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(fl.liters), 0) AS fuel_liters,
        COALESCE(SUM(COALESCE(fl.hours_run, 0)), 0) AS run_hours_ref,
        COUNT(*) AS entries
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date BETWEEN ? AND ?
      GROUP BY a.id
      ORDER BY fuel_liters DESC
      LIMIT 25
    `).all(start, end);

    const oilUsageByAsset = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ol.quantity), 0) AS oil_qty,
        COUNT(*) AS entries
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.id
      ORDER BY oil_qty DESC
      LIMIT 25
    `).all(start, end);

    const pdf = await buildPdfBuffer(
      (doc) => {
        const logoPath = path.join(process.cwd(), "branding", "logo.png");
        tryDrawLogo(doc, logoPath);
        sectionTitle(doc, "Operations Summary");
        kvGrid(doc, [
          { label: "Period", value: `${start} to ${end}` },
          { label: "Entries", value: String(rows.length) },
          { label: "Tonnes moved", value: fmtNum(totals.tonnes, 2) },
          { label: "Produced", value: fmtNum(totals.produced, 2) },
          { label: "Delivered", value: fmtNum(totals.productDelivered, 2) },
          { label: "Trucks loaded", value: fmtNum(totals.loaded, 0) },
          { label: "Trucks delivered", value: fmtNum(totals.delivered, 0) },
          { label: "Weighbridge", value: fmtNum(totals.weighbridge, 2) },
        ], 2);

        sectionTitle(doc, "By Product Type");
        table(
          doc,
          [
            { key: "product", label: "Product", width: 0.36 },
            { key: "entries", label: "Entries", width: 0.12, align: "right" },
            { key: "tonnes", label: "Tonnes", width: 0.16, align: "right" },
            { key: "produced", label: "Produced", width: 0.18, align: "right" },
            { key: "delivered", label: "Delivered", width: 0.18, align: "right" },
          ],
          byProduct.map((r) => ({
            product: compactCell(r.product_type, 70),
            entries: fmtNum(r.entries, 0),
            tonnes: fmtNum(r.tonnes_moved, 2),
            produced: fmtNum(r.product_produced, 2),
            delivered: fmtNum(r.product_delivered, 2),
          }))
        );

        sectionTitle(doc, "By Client");
        table(
          doc,
          [
            { key: "client", label: "Client", width: 0.46 },
            { key: "entries", label: "Entries", width: 0.12, align: "right" },
            { key: "trucks", label: "Trucks", width: 0.16, align: "right" },
            { key: "delivered", label: "Delivered", width: 0.26, align: "right" },
          ],
          byClient.map((r) => ({
            client: compactCell(r.client_name, 80),
            entries: fmtNum(r.entries, 0),
            trucks: fmtNum(r.trucks_delivered, 0),
            delivered: fmtNum(r.product_delivered, 2),
          }))
        );

        sectionTitle(doc, "Fuel & Oil Usage (Stores / Logs)");
        kvGrid(doc, [
          { label: "Fuel issued (L)", value: fmtNum(fuelUsageSummary?.fuel_liters || 0, 2) },
          { label: "Fuel log entries", value: fmtNum(fuelUsageSummary?.entries || 0, 0) },
          { label: "Fuel run hours reference", value: fmtNum(fuelUsageSummary?.run_hours_ref || 0, 1) },
          { label: "Oil/Lube issued (qty)", value: fmtNum(oilUsageSummary?.oil_qty || 0, 2) },
          { label: "Oil log entries", value: fmtNum(oilUsageSummary?.entries || 0, 0) },
        ], 2);

        sectionTitle(doc, "Fuel Usage by Asset (Top 25)");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.38 },
            { key: "entries", label: "Entries", width: 0.12, align: "right" },
            { key: "fuel", label: "Fuel (L)", width: 0.20, align: "right" },
            { key: "hrs", label: "Run hrs ref", width: 0.16, align: "right" },
            { key: "lph", label: "L/hr", width: 0.14, align: "right" },
          ],
          fuelUsageByAsset.length
            ? fuelUsageByAsset.map((r) => {
                const liters = Number(r.fuel_liters || 0);
                const runHours = Number(r.run_hours_ref || 0);
                const lph = runHours > 0 ? liters / runHours : null;
                return {
                  asset: compactCell(`${r.asset_code || ""} - ${r.asset_name || ""}`, 70),
                  entries: fmtNum(r.entries, 0),
                  fuel: fmtNum(liters, 2),
                  hrs: fmtNum(runHours, 1),
                  lph: lph == null ? "-" : fmtNum(lph, 2),
                };
              })
            : [{ asset: "No fuel usage in period", entries: "-", fuel: "-", hrs: "-", lph: "-" }]
        );

        sectionTitle(doc, "Oil/Lube Usage by Asset (Top 25)");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.58 },
            { key: "entries", label: "Entries", width: 0.14, align: "right" },
            { key: "qty", label: "Oil qty", width: 0.28, align: "right" },
          ],
          oilUsageByAsset.length
            ? oilUsageByAsset.map((r) => ({
                asset: compactCell(`${r.asset_code || ""} - ${r.asset_name || ""}`, 90),
                entries: fmtNum(r.entries, 0),
                qty: fmtNum(r.oil_qty, 2),
              }))
            : [{ asset: "No oil usage in period", entries: "-", qty: "-" }]
        );

        sectionTitle(doc, "Entry Detail");
        table(
          doc,
          [
            { key: "date", label: "Date", width: 0.12 },
            { key: "product", label: "Product", width: 0.16 },
            { key: "tonnes", label: "Tonnes", width: 0.10, align: "right" },
            { key: "prod", label: "Produced", width: 0.10, align: "right" },
            { key: "loaded", label: "Loaded", width: 0.08, align: "right" },
            { key: "wb", label: "Weighbridge", width: 0.12, align: "right" },
            { key: "delTrk", label: "Deliv Trucks", width: 0.10, align: "right" },
            { key: "delProd", label: "Delivered", width: 0.10, align: "right" },
            { key: "client", label: "Client", width: 0.12 },
          ],
          rows.slice(0, 250).map((r) => ({
            date: r.op_date || "",
            product: compactCell(r.product_type || "", 24),
            tonnes: fmtNum(r.tonnes_moved, 2),
            prod: fmtNum(r.product_produced, 2),
            loaded: fmtNum(r.trucks_loaded, 0),
            wb: fmtNum(r.weighbridge_amount, 2),
            delTrk: fmtNum(r.trucks_delivered, 0),
            delProd: fmtNum(r.product_delivered, 2),
            client: compactCell(r.client_delivered_to || "", 22),
          }))
        );
      },
      {
        title: "IRONLOG",
        subtitle: `Operations Report (${reportRevision})`,
        rightText: `Period: ${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Operations_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/operations.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD
  app.get("/operations.xlsx", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }

    const rows = db.prepare(`
      SELECT
        op_date, tonnes_moved, product_type, product_produced, trucks_loaded, weighbridge_amount,
        trucks_delivered, product_delivered, client_delivered_to, notes
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      ORDER BY op_date ASC, id ASC
      LIMIT 5000
    `).all(start, end);

    const byDate = db.prepare(`
      SELECT
        op_date,
        COUNT(*) AS entries,
        IFNULL(SUM(tonnes_moved), 0) AS tonnes_moved,
        IFNULL(SUM(product_produced), 0) AS product_produced,
        IFNULL(SUM(trucks_loaded), 0) AS trucks_loaded,
        IFNULL(SUM(weighbridge_amount), 0) AS weighbridge_amount,
        IFNULL(SUM(trucks_delivered), 0) AS trucks_delivered,
        IFNULL(SUM(product_delivered), 0) AS product_delivered
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      GROUP BY op_date
      ORDER BY op_date ASC
    `).all(start, end);

    const byProduct = db.prepare(`
      SELECT
        COALESCE(NULLIF(TRIM(product_type), ''), 'Unspecified') AS product_type,
        COUNT(*) AS entries,
        IFNULL(SUM(tonnes_moved), 0) AS tonnes_moved,
        IFNULL(SUM(product_produced), 0) AS product_produced,
        IFNULL(SUM(product_delivered), 0) AS product_delivered
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      GROUP BY COALESCE(NULLIF(TRIM(product_type), ''), 'Unspecified')
      ORDER BY tonnes_moved DESC
    `).all(start, end);

    const byClient = db.prepare(`
      SELECT
        COALESCE(NULLIF(TRIM(client_delivered_to), ''), 'Unspecified') AS client_name,
        COUNT(*) AS entries,
        IFNULL(SUM(trucks_delivered), 0) AS trucks_delivered,
        IFNULL(SUM(product_delivered), 0) AS product_delivered
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      GROUP BY COALESCE(NULLIF(TRIM(client_delivered_to), ''), 'Unspecified')
      ORDER BY product_delivered DESC
    `).all(start, end);

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    addTableSheet(
      wb,
      "Operations Detail",
      [
        { header: "Date", key: "op_date", width: 14 },
        { header: "Product Type", key: "product_type", width: 24 },
        { header: "Tonnes Moved", key: "tonnes_moved", width: 16 },
        { header: "Product Produced", key: "product_produced", width: 18 },
        { header: "Trucks Loaded", key: "trucks_loaded", width: 14 },
        { header: "Weighbridge Amount", key: "weighbridge_amount", width: 18 },
        { header: "Trucks Delivered", key: "trucks_delivered", width: 16 },
        { header: "Product Delivered", key: "product_delivered", width: 16 },
        { header: "Client Delivered To", key: "client_delivered_to", width: 24 },
        { header: "Notes", key: "notes", width: 36 },
      ],
      rows.map((r) => ({
        op_date: r.op_date || "",
        product_type: r.product_type || "",
        tonnes_moved: Number(r.tonnes_moved || 0),
        product_produced: Number(r.product_produced || 0),
        trucks_loaded: Number(r.trucks_loaded || 0),
        weighbridge_amount: Number(r.weighbridge_amount || 0),
        trucks_delivered: Number(r.trucks_delivered || 0),
        product_delivered: Number(r.product_delivered || 0),
        client_delivered_to: r.client_delivered_to || "",
        notes: r.notes || "",
      }))
    );

    addTableSheet(
      wb,
      "By Date",
      [
        { header: "Date", key: "op_date", width: 14 },
        { header: "Entries", key: "entries", width: 12 },
        { header: "Tonnes Moved", key: "tonnes_moved", width: 16 },
        { header: "Produced", key: "product_produced", width: 14 },
        { header: "Trucks Loaded", key: "trucks_loaded", width: 14 },
        { header: "Weighbridge", key: "weighbridge_amount", width: 14 },
        { header: "Trucks Delivered", key: "trucks_delivered", width: 16 },
        { header: "Delivered", key: "product_delivered", width: 14 },
      ],
      byDate.map((r) => ({
        op_date: r.op_date,
        entries: Number(r.entries || 0),
        tonnes_moved: Number(r.tonnes_moved || 0),
        product_produced: Number(r.product_produced || 0),
        trucks_loaded: Number(r.trucks_loaded || 0),
        weighbridge_amount: Number(r.weighbridge_amount || 0),
        trucks_delivered: Number(r.trucks_delivered || 0),
        product_delivered: Number(r.product_delivered || 0),
      }))
    );

    addTableSheet(
      wb,
      "By Product",
      [
        { header: "Product Type", key: "product_type", width: 26 },
        { header: "Entries", key: "entries", width: 12 },
        { header: "Tonnes Moved", key: "tonnes_moved", width: 16 },
        { header: "Produced", key: "product_produced", width: 14 },
        { header: "Delivered", key: "product_delivered", width: 14 },
      ],
      byProduct.map((r) => ({
        product_type: r.product_type,
        entries: Number(r.entries || 0),
        tonnes_moved: Number(r.tonnes_moved || 0),
        product_produced: Number(r.product_produced || 0),
        product_delivered: Number(r.product_delivered || 0),
      }))
    );

    addTableSheet(
      wb,
      "By Client",
      [
        { header: "Client", key: "client_name", width: 28 },
        { header: "Entries", key: "entries", width: 12 },
        { header: "Trucks Delivered", key: "trucks_delivered", width: 16 },
        { header: "Product Delivered", key: "product_delivered", width: 16 },
      ],
      byClient.map((r) => ({
        client_name: r.client_name,
        entries: Number(r.entries || 0),
        trucks_delivered: Number(r.trucks_delivered || 0),
        product_delivered: Number(r.product_delivered || 0),
      }))
    );

    const buffer = await wb.xlsx.writeBuffer();
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Operations_${start}_to_${end}.xlsx"`)
      .send(Buffer.from(buffer));
  });

  // GET /api/reports/executive-pack.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD&scheduled=10&near_due_hours=50
  app.get("/executive-pack.xlsx", async (req, reply) => {
    const end = String(req.query?.end || "").trim() || todayYmd();
    const start = String(req.query?.start || "").trim() || monthStartIso(end);
    const scheduled = Math.max(0.5, Number(req.query?.scheduled || 10));
    const nearDue = Math.max(1, Number(req.query?.near_due_hours || 50));
    if (!isYmd(start) || !isYmd(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }
    if (start > end) {
      return reply.code(400).send({ error: "start must be <= end" });
    }

    const siteCode = String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
    const sharedHeaders = {
      "x-site-code": siteCode,
      "x-user-role": String(req.headers["x-user-role"] || "admin"),
    };
    if (req.headers.authorization) sharedHeaders.authorization = String(req.headers.authorization);

    async function injectJson(url) {
      const res = await app.inject({ method: "GET", url, headers: sharedHeaders });
      if (res.statusCode >= 400) throw new Error(`${url} -> HTTP ${res.statusCode}`);
      try {
        return JSON.parse(res.payload || "{}");
      } catch {
        return {};
      }
    }

    const qCommon = `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
    const [
      kpi,
      fuel,
      fuelDup,
      wfSummary,
      wfActions,
      lube,
      mgrIns,
      artIns,
      stock,
    ] = await Promise.all([
      injectJson(`/api/dashboard/asset-kpi/weekly?${qCommon}&scheduled=${encodeURIComponent(String(scheduled))}`).catch(() => ({})),
      injectJson(`/api/dashboard/fuel?${qCommon}&tolerance=0.15`).catch(() => ({})),
      injectJson(`/api/dashboard/fuel/duplicates?${qCommon}`).catch(() => ({})),
      injectJson(`/api/maintenance/weekly-forum/summary?${qCommon}&near_due_hours=${encodeURIComponent(String(nearDue))}`).catch(() => ({})),
      injectJson(`/api/maintenance/weekly-forum/actions?${qCommon}`).catch(() => ({})),
      injectJson(`/api/dashboard/lube/analytics?${qCommon}`).catch(() => ({})),
      injectJson(`/api/maintenance/inspections?${qCommon}`).catch(() => ({})),
      injectJson(`/api/maintenance/artisan-inspections?${qCommon}`).catch(() => ({})),
      injectJson("/api/stock/monitor").catch(() => ({})),
    ]);

    function autosizeColumns(ws, maxWidth = 45) {
      ws.columns.forEach((col) => {
        let width = 10;
        col.eachCell({ includeEmpty: true }, (cell) => {
          const len = String(cell.value ?? "").length;
          width = Math.max(width, Math.min(maxWidth, len + 2));
        });
        col.width = width;
      });
    }
    function writeRows(ws, headers, rows) {
      ws.addRow(headers);
      ws.getRow(1).font = { bold: true };
      for (const row of rows) ws.addRow(headers.map((h) => row[h]));
      ws.views = [{ state: "frozen", ySplit: 1 }];
      autosizeColumns(ws);
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    {
      const ws = wb.addWorksheet("00_Control");
      const rows = [
        { key: "start_date", value: start },
        { key: "end_date", value: end },
        { key: "site_code", value: siteCode },
        { key: "generated_at", value: new Date().toISOString() },
      ];
      writeRows(ws, ["key", "value"], rows);
    }
    {
      const ws = wb.addWorksheet("01_HSE_Summary");
      const mgrRows = asArray(mgrIns.rows);
      const artRows = asArray(artIns.rows);
      const mgrFindings = mgrRows.reduce((acc, r) => acc + asArray(r.checklist).filter((c) => c?.ok === false).length, 0);
      const artFindings = artRows.reduce((acc, r) => acc + asArray(r.checklist).filter((c) => c?.ok === false).length, 0);
      const rows = [{
        period_start: start,
        period_end: end,
        manager_inspections: mgrRows.length,
        artisan_inspections: artRows.length,
        inspections_total: mgrRows.length + artRows.length,
        findings_open_proxy: mgrFindings + artFindings,
      }];
      writeRows(ws, Object.keys(rows[0]), rows);
    }
    {
      const ws = wb.addWorksheet("03_Plant_Performance");
      const rows = asArray(kpi.by_asset).map((r) => ({
        asset_code: r.asset_code || "",
        asset_name: r.asset_name || "",
        category: r.category || "",
        scheduled_hours: safeNum(r.scheduled_hours),
        run_hours: safeNum(r.run_hours),
        downtime_hours: safeNum(r.downtime_hours),
        available_hours: safeNum(r.available_hours),
        availability_pct: r.availability_pct == null ? null : safeNum(r.availability_pct),
        utilization_pct: r.utilization_pct == null ? null : safeNum(r.utilization_pct),
      }));
      writeRows(ws, rows.length ? Object.keys(rows[0]) : ["asset_code", "asset_name", "category", "scheduled_hours", "run_hours", "downtime_hours", "available_hours", "availability_pct", "utilization_pct"], rows);
    }
    {
      const ws = wb.addWorksheet("04_Maint_Cost_Machine");
      const rows = asArray(wfSummary.upcoming_services).map((r) => ({
        asset_code: r.asset_code || "",
        asset_name: r.asset_name || "",
        service_name: r.service_name || "",
        remaining_hours: safeNum(r.remaining_hours),
        avg_oil_cost: safeNum(r?.forecast?.avg_oil_cost),
        avg_parts_cost: safeNum(r?.forecast?.avg_parts_cost),
        est_service_kit_cost: safeNum(r?.forecast?.est_service_kit_cost),
      }));
      writeRows(ws, rows.length ? Object.keys(rows[0]) : ["asset_code", "asset_name", "service_name", "remaining_hours", "avg_oil_cost", "avg_parts_cost", "est_service_kit_cost"], rows);
    }
    {
      const ws = wb.addWorksheet("05_Parts_Tracking");
      const rows = asArray(stock.rows).map((r) => ({
        part_code: r.part_code || "",
        part_name: r.part_name || "",
        category: r.category || "",
        on_hand: safeNum(r.on_hand),
        min_stock: safeNum(r.min_stock),
        below_min_flag: safeNum(r.is_below_min),
        critical_flag: safeNum(r.is_critical),
      }));
      writeRows(ws, rows.length ? Object.keys(rows[0]) : ["part_code", "part_name", "category", "on_hand", "min_stock", "below_min_flag", "critical_flag"], rows);
    }
    {
      const ws = wb.addWorksheet("06_Production_Support");
      const rows = asArray(wfActions.rows).map((r) => ({
        action_date: r.action_date || "",
        owner: r.owner || "",
        action_text: r.action_text || "",
        status: r.status || "",
        due_date: r.due_date || "",
      }));
      writeRows(ws, rows.length ? Object.keys(rows[0]) : ["action_date", "owner", "action_text", "status", "due_date"], rows);
    }
    {
      const ws = wb.addWorksheet("07_Lube_Cost_Machine");
      const rows = asArray(lube.rows).map((r) => ({
        asset_code: r.asset_code || "",
        asset_name: r.asset_name || "",
        lube_type: r.lube_type || "",
        qty_total: safeNum(r.qty_total),
        entries: safeNum(r.entries),
        total_lube_cost: safeNum(r.total_lube_cost),
      }));
      writeRows(ws, rows.length ? Object.keys(rows[0]) : ["asset_code", "asset_name", "lube_type", "qty_total", "entries", "total_lube_cost"], rows);
    }
    {
      const ws = wb.addWorksheet("08_Inspections");
      const rows = [
        { inspection_type: "manager", completed_count: asArray(mgrIns.rows).length },
        { inspection_type: "artisan", completed_count: asArray(artIns.rows).length },
        { inspection_type: "total", completed_count: asArray(mgrIns.rows).length + asArray(artIns.rows).length },
      ];
      writeRows(ws, ["inspection_type", "completed_count"], rows);
    }
    {
      const ws = wb.addWorksheet("09_Fuel_Security");
      const rows = asArray(fuel.rows).map((r) => ({
        asset_code: r.asset_code || "",
        asset_name: r.asset_name || "",
        metric_mode: r.metric_mode || "",
        fuel_liters: safeNum(r.fuel_liters),
        hours_run: safeNum(r.hours_run),
        actual_lph: r.actual_lph == null ? null : safeNum(r.actual_lph),
        oem_lph: r.oem_lph == null ? null : safeNum(r.oem_lph),
        variance_lph: r.variance_lph == null ? null : safeNum(r.variance_lph),
        is_excessive: r.is_excessive ? 1 : 0,
        duplicate_rows: asArray(fuelDup.rows).filter((d) => String(d.asset_code || "") === String(r.asset_code || "")).length,
      }));
      writeRows(ws, rows.length ? Object.keys(rows[0]) : ["asset_code", "asset_name", "metric_mode", "fuel_liters", "hours_run", "actual_lph", "oem_lph", "variance_lph", "is_excessive", "duplicate_rows"], rows);
    }
    {
      const ws = wb.addWorksheet("10_Slide_Map");
      const rows = [
        { slide_no: 1, slide_title: "Safety (HSE)", sheet: "01_HSE_Summary" },
        { slide_no: 2, slide_title: "Plant Performance", sheet: "03_Plant_Performance" },
        { slide_no: 3, slide_title: "Breakdown & Maintenance (Cost/Machine)", sheet: "04_Maint_Cost_Machine" },
        { slide_no: 4, slide_title: "Parts Tracking", sheet: "05_Parts_Tracking" },
        { slide_no: 5, slide_title: "Production Support", sheet: "06_Production_Support" },
        { slide_no: 6, slide_title: "Lubrication (Cost/Machine)", sheet: "07_Lube_Cost_Machine" },
        { slide_no: 7, slide_title: "Inspections Done", sheet: "08_Inspections" },
        { slide_no: 8, slide_title: "Security (Fuel Anomalies)", sheet: "09_Fuel_Security" },
      ];
      writeRows(ws, ["slide_no", "slide_title", "sheet"], rows);
    }

    const buffer = await wb.xlsx.writeBuffer();
    return reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Executive_Pack_${start}_to_${end}.xlsx"`)
      .send(Buffer.from(buffer));
  });

  // GET /api/reports/executive-kpi-pack.xlsx?period_type=weekly|monthly&start=YYYY-MM-DD&end=YYYY-MM-DD&month=YYYY-MM&site_codes=main,site-b
  app.get("/executive-kpi-pack.xlsx", async (req, reply) => {
    const periodType = String(req.query?.period_type || "weekly").trim().toLowerCase();
    const rawSites = String(req.query?.site_codes || "main").trim();
    const siteCodes = Array.from(new Set(rawSites.split(",").map((s) => String(s || "").trim().toLowerCase()).filter(Boolean))).slice(0, 20);
    if (!["weekly", "monthly"].includes(periodType)) {
      return reply.code(400).send({ error: "period_type must be weekly or monthly" });
    }
    if (!siteCodes.length) {
      return reply.code(400).send({ error: "site_codes is required" });
    }
    let start = "";
    let end = "";
    if (periodType === "monthly") {
      const month = String(req.query?.month || "").trim();
      const m = isMonth(month) ? month : todayYmd().slice(0, 7);
      start = `${m}-01`;
      const d = new Date(`${start}T00:00:00Z`);
      d.setUTCMonth(d.getUTCMonth() + 1);
      d.setUTCDate(0);
      end = d.toISOString().slice(0, 10);
    } else {
      end = String(req.query?.end || "").trim() || todayYmd();
      start = String(req.query?.start || "").trim() || monthStartIso(end);
      if (!isYmd(start) || !isYmd(end) || start > end) {
        return reply.code(400).send({ error: "weekly range requires valid start/end YYYY-MM-DD and start <= end" });
      }
    }
    const scheduled = Math.max(0.5, Number(req.query?.scheduled || 10));
    const nearDue = Math.max(1, Number(req.query?.near_due_hours || 50));

    async function injectJson(url, siteCode) {
      const headers = {
        "x-site-code": siteCode,
        "x-user-role": String(req.headers["x-user-role"] || "admin"),
      };
      if (req.headers.authorization) headers.authorization = String(req.headers.authorization);
      const res = await app.inject({ method: "GET", url, headers });
      if (res.statusCode >= 400) return {};
      try { return JSON.parse(res.payload || "{}"); } catch { return {}; }
    }
    const asArray = (v) => (Array.isArray(v) ? v : []);
    const safeNum = (v) => {
      const n = Number(v);
      return Number.isFinite(n) ? n : 0;
    };
    const qCommon = `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
    const bySite = [];
    for (const siteCode of siteCodes) {
      const [kpi, wfSummary, mgrIns, artIns, fuel] = await Promise.all([
        injectJson(`/api/dashboard/asset-kpi/weekly?${qCommon}&scheduled=${encodeURIComponent(String(scheduled))}`, siteCode),
        injectJson(`/api/maintenance/weekly-forum/summary?${qCommon}&near_due_hours=${encodeURIComponent(String(nearDue))}`, siteCode),
        injectJson(`/api/maintenance/inspections?${qCommon}`, siteCode),
        injectJson(`/api/maintenance/artisan-inspections?${qCommon}`, siteCode),
        injectJson(`/api/dashboard/fuel?${qCommon}&tolerance=0.15`, siteCode),
      ]);
      const assets = asArray(kpi.by_asset);
      const availabilityAvg = assets.length
        ? assets.reduce((acc, r) => acc + safeNum(r.availability_pct), 0) / assets.length
        : 0;
      const utilizationAvg = assets.length
        ? assets.reduce((acc, r) => acc + safeNum(r.utilization_pct), 0) / assets.length
        : 0;
      const upcoming = asArray(wfSummary.upcoming_services);
      const fuelRows = asArray(fuel.rows);
      const fuelAnomalies = fuelRows.filter((r) => Number(r.is_excessive || 0) === 1).length;
      bySite.push({
        site_code: siteCode,
        assets_tracked: assets.length,
        avg_availability_pct: Number(availabilityAvg.toFixed(2)),
        avg_utilization_pct: Number(utilizationAvg.toFixed(2)),
        upcoming_services: upcoming.length,
        manager_inspections: asArray(mgrIns.rows).length,
        artisan_inspections: asArray(artIns.rows).length,
        fuel_anomalies: fuelAnomalies,
        est_service_cost: Number(
          upcoming.reduce((acc, r) => acc + safeNum(r?.forecast?.est_service_kit_cost), 0).toFixed(2)
        ),
      });
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();
    const ws = wb.addWorksheet("Site Comparison");
    ws.addRow([
      "site_code",
      "period_type",
      "start_date",
      "end_date",
      "assets_tracked",
      "avg_availability_pct",
      "avg_utilization_pct",
      "upcoming_services",
      "manager_inspections",
      "artisan_inspections",
      "fuel_anomalies",
      "est_service_cost",
    ]);
    ws.getRow(1).font = { bold: true };
    for (const row of bySite) {
      ws.addRow([
        row.site_code,
        periodType,
        start,
        end,
        row.assets_tracked,
        row.avg_availability_pct,
        row.avg_utilization_pct,
        row.upcoming_services,
        row.manager_inspections,
        row.artisan_inspections,
        row.fuel_anomalies,
        row.est_service_cost,
      ]);
    }
    ws.views = [{ state: "frozen", ySplit: 1 }];
    ws.columns.forEach((col) => {
      let width = 14;
      col.eachCell({ includeEmpty: true }, (cell) => {
        width = Math.max(width, Math.min(48, String(cell.value ?? "").length + 2));
      });
      col.width = width;
    });

    const buf = await wb.xlsx.writeBuffer();
    return reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Executive_KPI_Pack_${periodType}_${start}_to_${end}.xlsx"`)
      .send(Buffer.from(buf));
  });
}