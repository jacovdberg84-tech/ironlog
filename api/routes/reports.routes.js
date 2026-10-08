// IRONLOG/api/routes/reports.routes.js
import path from "node:path";
import fs from "node:fs";
import { fileURLToPath } from "node:url";
import multipart from "@fastify/multipart";
import PptxGenJS from "pptxgenjs";
import { db } from "../db/client.js";
import { ensureSmtpTables, buildSmtpTransport, formatSmtpError, sendIronlogMail } from "../utils/mail.js";
import { ensurePageSpace } from "../utils/pdfGenerator.js";
import { andDailyHoursFleetHoursOnly, andAssetFleetHoursOnly, andAssetExcludeLdv } from "../utils/fleetHoursKpiScope.js";
import { getRunFromFuelRows } from "../utils/fuelRunFromLogs.js";
import { inferHireContractorLabel, sqlIncludeArchivedHireAssets } from "../utils/hiredEquipment.js";
import { sqlFuelMetricModeExpr } from "../utils/fuelMetricMode.js";
import { getMachinePrestartTemplate } from "../utils/machinePrestartTemplates.js";
import { resolvePdfImage } from "../utils/imagePdf.js";
import { resolveStorageAbs as resolveStorageAbsUtil, getDataRoot, normalizeStorageRel } from "../utils/storagePaths.js";
import { getPdfBrandingLogoDir } from "../utils/reportSettings.js";
import { prestartDeductionForProductionFleet, PRESTART_DEDUCTION_HOURS } from "../utils/prestartDaily.js";
import { shiftDateYmd } from "../utils/shortBreakdowns.js";
import { isDate } from "../utils/request.js";
import registerSettingsRoutes from "./reports/settings.routes.js";
import registerWorkorderAssetRoutes from "./reports/workorder-asset.routes.js";
import registerFuelLubeRoutes from "./reports/fuel-lube.routes.js";
import registerInspectionsRoutes from "./reports/inspections.routes.js";
import registerStoresRoutes from "./reports/stores.routes.js";
import registerOperationsExportsRoutes from "./reports/operations-exports.routes.js";
import registerPresentationsRoutes from "./reports/presentations.routes.js";
import registerPeriodReportsRoutes from "./reports/period-reports.routes.js";
import { serviceCostSourceLabel } from "../utils/serviceCostSource.js";
import { DONE_WORK_ORDER_STATUSES, isWorkOrderDone, breakdownNextStep, downtimeCell, draftWeeklyFindings, isProductionAsset, issuedToEquipmentSql, workOrderBlockerCell, workOrderOwnerCell, workOrderProgressCell } from "../utils/weeklyForumFindings.js";

const __dirnameReports = path.dirname(fileURLToPath(import.meta.url));
const doneStatusSql = DONE_WORK_ORDER_STATUSES.map((s) => `'${s}'`).join(", ");
const AML_WEEKLY_TEMPLATE_PATH = path.join(
  __dirnameReports,
  "..",
  "templates",
  "AML Weekly Check Sheet V27.xlsx"
);

let maintenanceMasterSchedulerStarted = false;
let reportSubscriptionsSchedulerStarted = false;

function todayYmd() {
  return new Date().toISOString().slice(0, 10);
}

/** Build the authenticated context for a report's internal API request. */
export function maintenanceDeckInsightsHeaders(siteCode, requestHeaders = {}) {
  const headers = {
    "x-user-name": "system",
    "x-user-role": "admin",
    "x-user-roles": "admin",
    "x-site-code": siteCode || "main",
  };
  const authorization = String(requestHeaders?.authorization || requestHeaders?.Authorization || "").trim();
  if (authorization) headers.authorization = authorization;
  return headers;
}

// The Daily PDF is issued the following morning, while all operational data
// belongs to the completed shift on the preceding calendar day.
export function dailyPdfOperationsDate(reportIssueDate) {
  return shiftDateYmd(reportIssueDate, -1);
}

function yn(v) {
  return v ? "YES" : "NO";
}

function machinePrestartProfileFromCheckMode(checkMode) {
  const m = String(checkMode || "").toLowerCase();
  const pfx = "machine_prestart_";
  if (!m.startsWith(pfx)) return null;
  const id = m.slice(pfx.length).trim();
  return id || null;
}

function buildVehicleLdvChecklistPdfRows(checkMode, checklistJsonRaw) {
  let raw = {};
  try {
    const parsed = checklistJsonRaw ? JSON.parse(String(checklistJsonRaw)) : {};
    raw = parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  } catch {
    raw = {};
  }
  const mpId = machinePrestartProfileFromCheckMode(checkMode);
  if (mpId) {
    const tmpl = getMachinePrestartTemplate(mpId);
    if (!tmpl?.sections) return [];
    const rows = [];
    for (const sec of tmpl.sections) {
      for (const it of sec.items || []) {
        const k = String(it.key || "").trim();
        const secTitle = String(sec.title || "").trim();
        const lbl = String(it.label || k);
        rows.push({
          item: secTitle ? `${secTitle}: ${lbl}` : lbl,
          status: raw[k] === true ? "PASS" : "FAIL",
        });
      }
    }
    return rows;
  }
  const labelMap = {
    brakes_ok: "Brakes",
    lights_ok: "Lights",
    tyres_ok: "Tyres",
    oil_coolant_ok: "Oil/Coolant",
    leaks_damage_ok: "Leaks/Damage",
    safety_items_ok: "Safety Items",
  };
  return Object.entries(labelMap).map(([k, label]) => ({
    item: label,
    status: raw[k] === true ? "PASS" : "FAIL",
  }));
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

function dailyPdfLongDate(ymd) {
  const match = String(ymd || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return String(ymd || "");
  const value = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  if (!Number.isFinite(value.getTime())) return String(ymd || "");
  return new Intl.DateTimeFormat("en-GB", {
    timeZone: "UTC",
    day: "numeric",
    month: "long",
    year: "numeric",
  }).format(value);
}

function dailyPdfManagementSection(doc, label) {
  ensurePageSpace(doc, 38);
  const left = doc.page.margins.left;
  const right = doc.page.width - doc.page.margins.right;
  const y = doc.y;
  doc.font("Helvetica-Bold").fontSize(12).fillColor("#143c50").text(String(label || ""), left, y, {
    width: right - left,
    lineBreak: false,
  });
  doc
    .moveTo(left, y + 20)
    .lineTo(right, y + 20)
    .lineWidth(0.8)
    .strokeColor("#0f7182")
    .stroke();
  doc.y = y + 28;
}

function drawDailyPdfMetricCards(doc, cards) {
  const left = doc.page.margins.left;
  const right = doc.page.width - doc.page.margins.right;
  const gap = 11;
  const cardHeight = 62;
  const cardWidth = (right - left - gap * (cards.length - 1)) / cards.length;
  ensurePageSpace(doc, cardHeight + 14);
  const y = doc.y;

  for (let index = 0; index < cards.length; index += 1) {
    const card = cards[index];
    const x = left + index * (cardWidth + gap);
    doc.save();
    doc.roundedRect(x, y, cardWidth, cardHeight, 5).fill("#e8f2f5");
    doc.restore();
    doc
      .font("Helvetica-Bold")
      .fontSize(8)
      .fillColor("#526a7c")
      .text(String(card.label || "").toUpperCase(), x + 10, y + 10, {
        width: cardWidth - 20,
        lineBreak: false,
      });
    doc
      .font("Helvetica-Bold")
      .fontSize(20)
      .fillColor(card.alert ? "#b42318" : "#143c50")
      .text(String(card.value || "-"), x + 10, y + 25, {
        width: cardWidth - 20,
        lineBreak: false,
      });
    if (card.detail) {
      doc
        .font("Helvetica")
        .fontSize(7.5)
        .fillColor("#526a7c")
        .text(String(card.detail), x + 10, y + 49, {
          width: cardWidth - 20,
          lineBreak: false,
        });
    }
  }
  doc.y = y + cardHeight + 12;
}

function drawDailyPdfOperatingBasis(doc, values) {
  const left = doc.page.margins.left;
  const right = doc.page.width - doc.page.margins.right;
  const gap = 11;
  const rowHeight = 39;
  const colWidth = (right - left - gap) / 2;
  ensurePageSpace(doc, rowHeight * 2 + 16);
  const startY = doc.y;

  values.slice(0, 4).forEach((value, index) => {
    const row = Math.floor(index / 2);
    const col = index % 2;
    const x = left + col * (colWidth + gap);
    const y = startY + row * rowHeight;
    doc
      .font("Helvetica")
      .fontSize(8.5)
      .fillColor("#526a7c")
      .text(String(value.label || ""), x, y, { width: colWidth * 0.52, lineBreak: false });
    doc
      .font("Helvetica-Bold")
      .fontSize(9.5)
      .fillColor("#143c50")
      .text(String(value.value || "-"), x + colWidth * 0.48, y, {
        width: colWidth * 0.52,
        align: "right",
        lineBreak: false,
      });
    doc
      .moveTo(x, y + 24)
      .lineTo(x + colWidth, y + 24)
      .lineWidth(0.45)
      .strokeColor("#cbdde3")
      .stroke();
  });
  doc.y = startY + rowHeight * 2 + 4;
}

function drawDailyPdfExceptions(doc, rows, emptyMessage = "Nothing requiring management action was recorded for the operating day.") {
  const left = doc.page.margins.left;
  const right = doc.page.width - doc.page.margins.right;
  const width = right - left;
  if (!rows.length) {
    ensurePageSpace(doc, 42);
    const y = doc.y;
    doc.save();
    doc.roundedRect(left, y, width, 36, 4).fill("#f2f6f7");
    doc.restore();
    doc.font("Helvetica").fontSize(9).fillColor("#526a7c").text(
      String(emptyMessage),
      left + 12,
      y + 12,
      { width: width - 24, lineBreak: false },
    );
    doc.y = y + 46;
    return;
  }

  for (const row of rows) {
    const detail = String(row.detail || "").trim();
    const detailHeight = detail
      ? doc.font("Helvetica").fontSize(8.5).heightOfString(detail, { width: width - 176 })
      : 0;
    const height = Math.max(42, Math.ceil(24 + detailHeight));
    ensurePageSpace(doc, height + 8);
    const y = doc.y;
    const alert = row.alert === true;
    doc.save();
    doc.rect(left, y, 4, height).fill(alert ? "#b42318" : "#0f7182");
    doc.restore();
    doc
      .font("Helvetica-Bold")
      .fontSize(9.5)
      .fillColor("#143c50")
      .text(String(row.title || ""), left + 13, y + 7, { width: width - 176, lineBreak: false });
    if (detail) {
      doc.font("Helvetica").fontSize(8.5).fillColor("#526a7c").text(detail, left + 13, y + 20, {
        width: width - 176,
      });
    }
    doc
      .font("Helvetica-Bold")
      .fontSize(9)
      .fillColor(alert ? "#b42318" : "#143c50")
      .text(String(row.value || ""), left, y + 10, { width, align: "right", lineBreak: false });
    doc.y = y + height + 6;
  }
}

/**
 * Daily Log breakdown entries also record planned services. Detect an actual
 * service interval from the entered component/work text so the Daily PDF can
 * distinguish a maintenance stop from an unplanned breakdown.
 */
export function serviceLabelFromDailyDowntime(...values) {
  const text = values
    .map((value) => String(value ?? ""))
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
  if (!text) return "";

  const hourMatch = text.match(/\b(\d{2,5})\s*(?:h|hr|hrs|hour|hours)\s*(?:service|svc)\b/i);
  if (hourMatch) return `${Number(hourMatch[1])} hour service`;

  const kmMatch = text.match(/\b(\d{3,6})\s*(?:km|kilomet(?:er|re)s?)\s*(?:service|svc)\b/i);
  if (kmMatch) return `${Number(kmMatch[1])} km service`;

  return "";
}

export function dailyPdfDowntimeHours({ hasDailyEntry, isUsed, recordedHours, totalHours }) {
  const recorded = Math.max(0, Number(recordedHours || 0));
  const total = Math.max(0, Number(totalHours || 0));
  // A saved Production row is authoritative: only a logged repair loss can
  // reduce its availability, never the open-incident full-shift fallback.
  return Boolean(hasDailyEntry) && Number(isUsed || 0) === 1 ? recorded : total;
}

export function dailyPdfRepairFallbackHours({ dayCap, loggedHours, allocatedHours, repairHours }) {
  const cap = Math.max(0, Number(dayCap || 0));
  const logged = Math.max(0, Number(loggedHours || 0));
  const allocated = Math.max(0, Number(allocatedHours || 0));
  const repair = Math.max(0, Number(repairHours || 0));
  // Technician repair time is a fallback for a day with no Daily Log
  // downtime. Once a clerk has entered actual downtime, do not add a second
  // estimate from a work order for the same asset/day.
  if (logged > 0) return 0;
  return Math.min(repair, Math.max(0, cap - allocated));
}

function workOrderTerminalStatus(status) {
  const s = String(status || "").toLowerCase();
  return ["completed", "approved", "closed"].includes(s);
}

function jobFindingsTextForPdf(wo, breakdown) {
  const jobDesc = String(wo?.job_description || "").trim();
  const bdDesc = String(breakdown?.description || "").trim();
  const parts = [];
  if (jobDesc) parts.push(jobDesc);
  else if (!workOrderTerminalStatus(wo?.status)) {
    const legacy = String(wo?.completion_notes || "").trim();
    if (legacy) parts.push(legacy);
  }
  if (bdDesc && (!jobDesc || !jobDesc.includes(bdDesc))) parts.push(bdDesc);
  const progress = String(wo?.repair_progress || "").trim();
  if (progress) {
    const when = String(wo?.repair_progress_at || "").trim();
    parts.push(when ? `Progress (${when}): ${progress}` : `Progress: ${progress}`);
  }
  return parts.join("\n\n");
}

function completionNotesForPdf(wo, findingsText) {
  const notes = String(wo?.completion_notes || "").trim();
  if (!notes) return "-";
  const findings = String(findingsText || "").trim();
  if (findings && notes === findings) return "-";
  return compactCell(notes.replace(/\s+/g, " "), 220);
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
  const calculated = Number(bd.calendar_days_down || 0);
  // Prefer selected calendar date (breakdown_date) over start_at — start_at was historically
  // set to create time (datetime('now')), which is not the Date down the operator chose.
  const startDate = parseIsoDate(bd.breakdown_date) || parseIsoDate(bd.start_at);
  if (!startDate) return Math.max(logged, calculated, 1);

  const asOf = parseIsoDate(reportDate);
  const endDate = parseIsoDate(bd.end_at);
  let spanEnd = endDate || asOf || startDate;
  if (asOf && endDate && endDate > asOf) {
    // Daily report is "as of selected date", so never count beyond it.
    spanEnd = asOf;
  }

  const spanDays = spanEnd >= startDate ? inclusiveDaysBetween(startDate, spanEnd) : 0;
  if (logged > 0) {
    // A Daily Log downtime entry is the operational source of truth. Do not
    // turn a one-hour incident into several full days merely because its WO
    // remains open for follow-up work.
    return logged;
  }

  return Math.max(calculated, spanDays, 1);
}

function daysDownForBreakdownInRange(bd, startDateInclusive, endDateInclusive) {
  const startDate = parseIsoDate(bd.breakdown_date) || parseIsoDate(bd.start_at);
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

/** Calendar months overlapping an inclusive date range (YYYY-MM-DD). */
function monthsInRange(start, end) {
  const months = [];
  const [sy, sm] = String(start).split("-").map((n) => Number(n));
  const [ey, em] = String(end).split("-").map((n) => Number(n));
  if (!Number.isFinite(sy) || !Number.isFinite(sm) || !Number.isFinite(ey) || !Number.isFinite(em)) return months;
  let y = sy;
  let m = sm;
  while (y < ey || (y === ey && m <= em)) {
    const label = `${y}-${String(m).padStart(2, "0")}`;
    const { start: monthStart, end: monthEnd } = monthRange(label);
    const rangeStart = monthStart < start ? start : monthStart;
    const rangeEnd = monthEnd > end ? end : monthEnd;
    if (rangeStart <= rangeEnd) {
      months.push({ label, start: rangeStart, end: rangeEnd });
    }
    m += 1;
    if (m > 12) {
      m = 1;
      y += 1;
    }
  }
  return months;
}

/** Calendar quarters for a year, clipped to an inclusive YTD end date. */
function calendarQuartersThrough(year, ytdEnd) {
  const y = Number(year);
  const endCap = String(ytdEnd || `${y}-12-31`).slice(0, 10);
  if (!Number.isFinite(y) || !/^\d{4}-\d{2}-\d{2}$/.test(endCap)) return [];
  const defs = [
    { key: "Q1", label: `Q1 ${y}`, start: `${y}-01-01`, end: `${y}-03-31` },
    { key: "Q2", label: `Q2 ${y}`, start: `${y}-04-01`, end: `${y}-06-30` },
    { key: "Q3", label: `Q3 ${y}`, start: `${y}-07-01`, end: `${y}-09-30` },
    { key: "Q4", label: `Q4 ${y}`, start: `${y}-10-01`, end: `${y}-12-31` },
  ];
  return defs
    .map((q) => {
      const end = q.end > endCap ? endCap : q.end;
      return { ...q, end, active: q.start <= endCap && q.start <= end };
    })
    .filter((q) => q.active);
}

const FUEL_BENCHMARK_BY_ASSET_COLUMNS = [
  { header: "Asset Code", key: "asset_code", width: 16 },
  { header: "Asset Name", key: "asset_name", width: 32 },
  { header: "Category", key: "category", width: 18 },
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
  { header: "Fuel Matched to Readings (L)", key: "matched_liters", width: 16 },
  { header: "Matched %", key: "coverage_pct", width: 11 },
  { header: "Worked Out From", key: "run_source_text", width: 22 },
  { header: "Rejected Readings", key: "suspect_readings", width: 12 },
  { header: "Notes", key: "note", width: 48 },
];

const RUN_SOURCE_TEXT = { fill_meter: "Meter readings on fills", daily_hours: "Daily hours", none: "No usable readings" };

/** Benchmark rows as Excel rows: a missing OEM benchmark reads "Not set". */
function fuelBenchmarkSheetRow(r) {
  const out = { ...r, run_source_text: RUN_SOURCE_TEXT[r.run_source] || "" };
  if (r.metric_mode === "km") {
    if (r.oem_km_per_l == null && "oem_set" in r) out.oem_km_per_l = "Not set";
  } else if (r.oem_lph == null && "oem_set" in r) {
    out.oem_lph = "Not set";
  }
  return out;
}

function addFuelBenchmarkByAssetWorksheet(wb, sheetName, rows, emptyMessage = "No fuel benchmark data for period") {
  const ws = wb.addWorksheet(sheetName);
  ws.columns = FUEL_BENCHMARK_BY_ASSET_COLUMNS;
  if (rows.length) {
    ws.addRows(rows.map(fuelBenchmarkSheetRow));
  } else {
    ws.addRow({
      asset_code: "-",
      asset_name: emptyMessage,
      category: "",
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
  return ws;
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

function kpiDaily(date, scheduled, dailyHoursDate = date, opts = {}) {
  const usedRow = db.prepare(`
    SELECT COUNT(DISTINCT dh.asset_id) AS used_assets
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date = ?
      ${andDailyHoursFleetHoursOnly("dh", "a")}
  `).get(dailyHoursDate);

  const used_assets = Number(usedRow.used_assets || 0);

  const runRow = db.prepare(`
    SELECT IFNULL(SUM(dh.hours_run), 0) AS run_hours
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date = ?
      ${andDailyHoursFleetHoursOnly("dh", "a")}
  `).get(dailyHoursDate);

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
  `).get(Number(scheduled || 0), dailyHoursDate);
  const utilization_base_hours = Number(utilBaseRow?.utilization_base_hours || 0);
  let available_hours = utilization_base_hours;

  const breakdownCols = db.prepare("PRAGMA table_info(breakdowns)").all();
  const hasBreakdownStatus = breakdownCols.some((r) => String(r.name || "") === "status");
  const hasBreakdownEndAt = breakdownCols.some((r) => String(r.name || "") === "end_at");
  const statusExpr = hasBreakdownStatus ? "TRIM(LOWER(COALESCE(b.status, '')))" : "''";
  const openStatePredicate = hasBreakdownStatus
    ? `${statusExpr} IN ('open', 'in_progress')`
    : "1 = 1";
  const endAtPredicate = hasBreakdownEndAt
    ? "(b.end_at IS NULL OR DATE(b.end_at) >= ?)"
    : "1 = 1";
  const activeBreakdownPredicate = hasBreakdownEndAt
    ? `(${endAtPredicate} OR ${openStatePredicate})`
    : `${openStatePredicate}`;
  // Logged downtime per machine, never more than the shift hours it did not
  // run (a machine cannot be down while it runs; see utils/downtimeCap.js).
  const dtLogsRow = db.prepare(`
    SELECT IFNULL(SUM(
      CASE WHEN COALESCE(x.run, 0) > 0 THEN MIN(x.down, MAX(0, x.sched - x.run)) ELSE x.down END
    ), 0) AS downtime_hours
    FROM (
      SELECT
        b.asset_id,
        SUM(l.hours_down) AS down,
        (SELECT MAX(COALESCE(dh.hours_run, 0)) FROM daily_hours dh WHERE dh.asset_id = b.asset_id AND dh.work_date = ?) AS run,
        (SELECT MAX(CASE WHEN COALESCE(dh.scheduled_hours, 0) > 0 THEN dh.scheduled_hours ELSE ? END)
           FROM daily_hours dh WHERE dh.asset_id = b.asset_id AND dh.work_date = ?) AS sched
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date = ?
        AND DATE(COALESCE(b.breakdown_date, l.log_date)) <= ?
        AND UPPER(TRIM(COALESCE(b.description, ''))) NOT LIKE 'MANAGER INSPECTION ALERT%'
        ${andAssetFleetHoursOnly("a")}
      GROUP BY b.asset_id
    ) x
  `).get(dailyHoursDate, Number(scheduled || 0), dailyHoursDate, date, date);
  let downtime_hours = Number(dtLogsRow?.downtime_hours || 0);
  const openNoLogParams = [Number(scheduled || 0), date, dailyHoursDate];
  if (hasBreakdownEndAt) openNoLogParams.push(date);
  openNoLogParams.push(date, date);
  const openNoLogRow = db.prepare(`
    SELECT
      -- A saved Daily Log production entry is the operator's statement that
      -- the asset was available. Do not turn it into a full-shift loss just
      -- because an earlier repair WO remains open or the meter was entered
      -- late. Explicit downtime logs remain the source of actual loss.
      IFNULL(SUM(CASE WHEN x.has_daily_row = 0 AND x.timed_first_day = 0 THEN x.scheduled ELSE 0 END), 0) AS assumed_down_hours,
      IFNULL(SUM(CASE WHEN x.has_daily_row = 0 THEN x.scheduled ELSE 0 END), 0) AS missing_planned_hours
    FROM (
      SELECT
        b.asset_id,
        MAX(CASE WHEN COALESCE(dh.scheduled_hours, 0) > 0 THEN dh.scheduled_hours ELSE ? END) AS scheduled,
        MAX(CASE WHEN dh.asset_id IS NULL THEN 0 ELSE 1 END) AS has_daily_row,
        MAX(COALESCE(dh.hours_run, 0)) AS run_hours,
        MAX(CASE
          WHEN DATE(COALESCE(b.start_at, '')) = ?
            AND TIME(COALESCE(b.start_at, '00:00:00')) > '00:00:00'
          THEN 1 ELSE 0
        END) AS timed_first_day
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      LEFT JOIN daily_hours dh ON dh.asset_id = b.asset_id AND dh.work_date = ? AND dh.is_used = 1
      WHERE ${activeBreakdownPredicate}
        AND b.breakdown_date <= ?
        AND UPPER(TRIM(COALESCE(b.description, ''))) NOT LIKE 'MANAGER INSPECTION ALERT%'
        AND NOT EXISTS (
          SELECT 1 FROM breakdown_downtime_logs l
          WHERE l.breakdown_id = b.id
            AND l.log_date <= ?
        )
        AND NOT EXISTS (
          SELECT 1 FROM work_orders wbx
          WHERE wbx.source = 'breakdown'
            AND COALESCE(wbx.reference_id, -1) = b.id
            AND REPLACE(TRIM(LOWER(COALESCE(wbx.status, ''))), ' ', '_') IN ('completed', 'approved', 'closed')
        )
        ${andAssetFleetHoursOnly("a")}
      GROUP BY b.asset_id
    ) x
  `).get(...openNoLogParams);
  downtime_hours += Number(openNoLogRow?.assumed_down_hours || 0);
  // A completed repair can be recorded after the operating day. The Daily PDF
  // passes the repair time that belongs to this day when no explicit downtime
  // log was captured, so it contributes to availability without rewriting the
  // original incident record.
  downtime_hours += Math.max(0, safeNum(opts.additionalDowntimeHours, 0));
  // A down asset still belongs in planned hours even when no Daily Log row was entered.
  available_hours += Number(openNoLogRow?.missing_planned_hours || 0);

  const prestart = prestartDeductionForProductionFleet(db, date);
  const prestart_hours = Number(prestart.hours || 0);
  const effective_loss_hours = downtime_hours + prestart_hours;
  const availability = available_hours > 0
    ? (Math.max(0, available_hours - effective_loss_hours) / available_hours) * 100
    : null;
  const utilization = utilization_base_hours > 0 ? (run_hours / utilization_base_hours) * 100 : null;

  return {
    used_assets,
    available_hours,
    utilization_base_hours,
    run_hours,
    downtime_hours,
    prestart_hours,
    prestart_count: Number(prestart.count || 0),
    prestart_deduction_hours_per_check: PRESTART_DEDUCTION_HOURS,
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
  const dataRoot = getDataRoot();
  await app.register(multipart, { limits: { fileSize: 2 * 1024 * 1024 } });
  const pdfBrandingLogoDir = getPdfBrandingLogoDir(dataRoot);
  fs.mkdirSync(pdfBrandingLogoDir, { recursive: true });
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
    return resolveStorageAbsUtil(relPath, dataRoot);
  }

  async function resolveCheckPhotoForPdf(photoRow) {
    const rel = normalizeStorageRel(photoRow?.file_path);
    const abs = resolveStorageAbs(rel);
    const resolved = await resolvePdfImage(abs);
    let markers = [];
    try {
      markers = photoRow?.markers_json ? JSON.parse(photoRow.markers_json) : [];
    } catch {
      markers = [];
    }
    if (Array.isArray(photoRow?.markers)) markers = photoRow.markers;
    return {
      ...photoRow,
      rel,
      pdfPath: resolved.path,
      tempPath: resolved.temp ? resolved.path : null,
      markers: Array.isArray(markers) ? markers : [],
    };
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
  ensureColumn("work_orders", "repair_progress", "repair_progress TEXT");
  ensureColumn("work_orders", "repair_progress_at", "repair_progress_at TEXT");

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
  ensureSmtpTables();

  const allowedReportTypes = new Set([
    "fuel_benchmark_xlsx",
    "executive_kpi_pack_xlsx",
    "maintenance_insights_xlsx",
  ]);
  const allowedChannels = new Set(["email", "whatsapp"]);
  const allowedFrequencies = new Set(["daily", "weekly", "monthly"]);

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
  const MAX_SUBSCRIPTION_ATTACHMENT_BYTES = 22 * 1024 * 1024;

  function subscriptionAttachFormat(filters = {}) {
    const raw = String(filters?.attach_format || "pdf").trim().toLowerCase();
    if (raw === "link" || raw === "none") return "link";
    if (raw === "xlsx" || raw === "excel") return "xlsx";
    if (raw === "both" || raw === "pdf_xlsx") return "both";
    return "pdf";
  }

  function internalApiBase() {
    return `http://127.0.0.1:${process.env.PORT_EFFECTIVE || process.env.PORT || 3001}`;
  }

  function publicReportLink(path) {
    const base = String(process.env.IRONLOG_PUBLIC_BASE_URL || "").trim().replace(/\/+$/, "");
    return base ? `${base}${path}` : path;
  }

  function resolveSubscriptionDateRange(filters = {}, reportType = "") {
    const f = filters && typeof filters === "object" ? filters : {};
    const type = String(reportType || "").trim().toLowerCase();
    const mode = String(f.period_mode || "").trim().toLowerCase();
    if ((type === "maintenance_insights_xlsx" || type === "fuel_benchmark_xlsx") && mode !== "fixed") {
      const end = todayYmd();
      const d = new Date(`${end}T12:00:00`);
      d.setDate(d.getDate() - 29);
      return { start: d.toISOString().slice(0, 10), end };
    }
    const start = isDate(f.start) ? String(f.start).trim() : `${todayYmd().slice(0, 8)}01`;
    const end = isDate(f.end) ? String(f.end).trim() : todayYmd();
    return { start, end };
  }

  function reportPathsForType(reportType, filters = {}) {
    const f = filters && typeof filters === "object" ? filters : {};
    const { start, end } = resolveSubscriptionDateRange(f, reportType);
    const tolerance = Number.isFinite(Number(f.tolerance)) ? Number(f.tolerance) : 0.15;
    const near = Number.isFinite(Number(f.near_due_hours)) ? Number(f.near_due_hours) : 50;
    const horizon = Number.isFinite(Number(f.predictive_horizon_hours)) ? Number(f.predictive_horizon_hours) : 100;
    const period = String(f.period_type || "weekly").trim().toLowerCase();
    const siteCodes = String(f.site_codes || "main").trim();
    const out = { pdf: null, xlsx: null, names: { pdf: null, xlsx: null }, start, end };

    if (reportType === "fuel_benchmark_xlsx") {
      const q = `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&tolerance=${tolerance}`;
      out.pdf = `/api/reports/fuel-benchmark.pdf?${q}&download=1`;
      out.xlsx = `/api/reports/fuel-benchmark.xlsx?${q}`;
      out.names.pdf = `IRONLOG_Fuel_Benchmark_${end}.pdf`;
      out.names.xlsx = `IRONLOG_Fuel_Benchmark_${end}.xlsx`;
    } else if (reportType === "executive_kpi_pack_xlsx") {
      const q = `period_type=${encodeURIComponent(period)}&start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&site_codes=${encodeURIComponent(siteCodes)}`;
      out.xlsx = `/api/reports/executive-kpi-pack.xlsx?${q}`;
      out.names.xlsx = `IRONLOG_Executive_KPI_Pack_${end}.xlsx`;
    } else {
      const checklist = Number.isFinite(Number(f.checklist_fail_threshold))
        ? Number(f.checklist_fail_threshold)
        : 2;
      const fuelVar = Number.isFinite(Number(f.fuel_variance_threshold))
        ? Number(f.fuel_variance_threshold)
        : 15;
      const q = `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&near_due_hours=${near}&predictive_horizon_hours=${horizon}&checklist_fail_threshold=${checklist}&fuel_variance_threshold=${fuelVar}`;
      out.pdf = `/api/maintenance/insights.pdf?${q}&download=1`;
      out.xlsx = `/api/maintenance/insights.xlsx?${q}`;
      out.names.pdf = `IRONLOG_Maintenance_Insights_${end}.pdf`;
      out.names.xlsx = `IRONLOG_Maintenance_Insights_${end}.xlsx`;
    }
    return out;
  }

  async function fetchInternalReport(path) {
    const res = await app.inject({
      method: "GET",
      url: path,
      headers: {
        "x-user-name": "system",
        "x-user-role": "admin",
        "x-user-roles": "admin",
        "x-site-code": String(process.env.IRONLOG_DEFAULT_SITE_CODE || "main"),
      },
    });
    if (res.statusCode >= 400) {
      const body = String(res.payload || "");
      let msg = body.slice(0, 300) || path;
      try {
        const j = JSON.parse(body);
        msg = j.error || j.message || msg;
      } catch {}
      throw new Error(`Report generation failed (${res.statusCode}): ${msg}`);
    }
    const buf = Buffer.from(res.rawPayload || res.payload || "");
    if (!buf.length) throw new Error(`Report generation returned empty file: ${path}`);
    return buf;
  }

  async function probeMaintenanceInsightsData(filters = {}) {
    const paths = reportPathsForType("maintenance_insights_xlsx", filters);
    const q = String(paths.pdf || "").split("?")[1] || "";
    const res = await app.inject({
      method: "GET",
      url: `/api/maintenance/insights?${q.replace(/&download=1/g, "").replace(/download=1&/g, "")}`,
      headers: {
        "x-user-name": "system",
        "x-user-role": "admin",
        "x-user-roles": "admin",
        "x-site-code": String(process.env.IRONLOG_DEFAULT_SITE_CODE || "main"),
      },
    });
    if (res.statusCode >= 400) {
      let msg = `HTTP ${res.statusCode}`;
      try {
        const j = JSON.parse(String(res.payload || "{}"));
        msg = j.error || j.message || msg;
      } catch {}
      throw new Error(`Maintenance insights data failed: ${msg}`);
    }
    const data = JSON.parse(String(res.payload || "{}"));
    if (!String(data?.range?.start || "").trim()) {
      throw new Error("Maintenance insights returned an empty payload (check API is up to date)");
    }
    return {
      data,
      at_risk: Array.isArray(data?.predictive?.at_risk_plans) ? data.predictive.at_risk_plans.length : 0,
      cost_rows: Array.isArray(data?.maintenance_cost) ? data.maintenance_cost.length : 0,
    };
  }

  async function buildSubscriptionAttachments(reportType, filters = {}) {
    const fmt = subscriptionAttachFormat(filters);
    if (fmt === "link") return [];
    const paths = reportPathsForType(reportType, filters);
    const attachments = [];
    const wantPdf = fmt === "pdf" || fmt === "both";
    const wantXlsx = fmt === "xlsx" || fmt === "both";

    async function add(path, filename, contentType) {
      if (!path || !filename) return;
      if (reportType === "maintenance_insights_xlsx") {
        await probeMaintenanceInsightsData(filters);
      }
      const content = await fetchInternalReport(path);
      if (content.length > MAX_SUBSCRIPTION_ATTACHMENT_BYTES) {
        const mb = Math.round(content.length / 1024 / 1024);
        throw new Error(`${filename} is too large to email (${mb} MB). Use link delivery or a narrower date range.`);
      }
      attachments.push({ filename, content, contentType });
    }

    if (wantPdf) {
      if (paths.pdf) {
        await add(paths.pdf, paths.names.pdf, "application/pdf");
      } else if (fmt === "pdf" && paths.xlsx) {
        await add(
          paths.xlsx,
          paths.names.xlsx,
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        );
      }
    }
    if (wantXlsx) {
      await add(
        paths.xlsx,
        paths.names.xlsx,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      );
    }
    return attachments;
  }

  function reportLinkForType(reportType, filters = {}) {
    const paths = reportPathsForType(reportType, filters);
    return paths.xlsx || paths.pdf || `/api/reports/${encodeURIComponent(reportType)}`;
  }

  function buildSubscriptionEmailText({ subRow, reportType, filters, link, attachments, generatedAt }) {
    const paths = reportPathsForType(reportType, filters);
    const publicLink = publicReportLink(link);
    const lines = [
      "Your IRONLOG report is ready.",
      "",
      `Report: ${String(subRow.name || reportType)}`,
      `Period: ${paths.start} to ${paths.end}`,
    ];
    if (attachments.length) {
      lines.push(`Attachments: ${attachments.map((a) => a.filename).join(", ")}`);
      lines.push("");
      lines.push(`You can also open the report in IRONLOG: ${publicLink}`);
    } else {
      lines.push("");
      lines.push(`Download: ${publicLink}`);
    }
    lines.push(`Generated: ${generatedAt}`);
    return lines.join("\n");
  }

  async function deliverSubscription(subRow, manual = false) {
    const id = Number(subRow.id || 0);
    const reportType = String(subRow.report_type || "");
    const channel = String(subRow.channel || "");
    const recipients = parseRecipients(subRow.recipients || "");
    let filters = {};
    try { filters = JSON.parse(String(subRow.filters_json || "{}")); } catch {}
    const link = reportLinkForType(reportType, filters);
    const generatedAt = new Date().toISOString();
    const payload = {
      subscription_id: id,
      name: String(subRow.name || ""),
      report_type: reportType,
      channel,
      recipients,
      report_link: link,
      attach_format: subscriptionAttachFormat(filters),
      manual: Boolean(manual),
      generated_at: generatedAt,
    };
    let status = "simulated";
    let detail = "Logged only";
    let attachments = [];
    if (channel === "email") {
      const smtp = buildSmtpTransport();
      if (smtp.error) {
        status = "failed";
        detail = smtp.error;
        if (manual) throw new Error(smtp.error);
      } else if (smtp.transporter) {
        try {
          if (subscriptionAttachFormat(filters) !== "link") {
            attachments = await buildSubscriptionAttachments(reportType, filters);
          }
          const mailOut = await sendIronlogMail({
            to: recipients,
            subject: `IRONLOG Report: ${String(subRow.name || reportType)}`,
            text: buildSubscriptionEmailText({ subRow, reportType, filters, link, attachments, generatedAt }),
            attachments: attachments.length ? attachments : undefined,
          });
          if (!mailOut.ok) throw new Error(mailOut.error || "SMTP send failed");
          status = "sent";
          let insightNote = "";
          if (reportType === "maintenance_insights_xlsx" && subscriptionAttachFormat(filters) !== "link") {
            try {
              const probe = await probeMaintenanceInsightsData(filters);
              insightNote = ` (${probe.at_risk} at-risk, ${probe.cost_rows} cost rows, ${probe.data?.range?.start || "?"} to ${probe.data?.range?.end || "?"})`;
            } catch {}
          }
          detail = attachments.length
            ? `SMTP email sent with ${attachments.length} attachment(s)${insightNote}`
            : `SMTP email sent (link only)${insightNote}`;
        } catch (err) {
          status = "failed";
          detail = formatSmtpError(err);
          if (manual) throw new Error(detail);
        }
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
    return { status, detail, report_link: link, attachment_count: attachments.length };
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

  function stockMovementDateExpr() {
    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    return hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
  }

  function queryPeriodFleetCostTotals(start, end) {
    const defaults = costDefaults();
    const smDateExpr = stockMovementDateExpr();

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

    const fuel_cost = Number(fuelCostRow?.value || 0);
    const lube_cost = Number(lubeCostRow?.value || 0);
    const parts_cost = Number(partsCostRow?.value || 0);
    const labor_cost = Number(laborRow?.labor_cost || 0);
    const labor_hours = Number(laborRow?.labor_hours || 0);
    const downtime_cost = Number(downtimeCostRow?.value || 0);
    const total_cost = Number((fuel_cost + lube_cost + parts_cost + labor_cost + downtime_cost).toFixed(2));

    return {
      defaults,
      fuel_cost,
      lube_cost,
      parts_cost,
      labor_cost,
      labor_hours,
      downtime_cost,
      total_cost,
    };
  }

  function buildPeriodAssetCosts(start, end) {
    const defaults = costDefaults();
    const smDateExpr = stockMovementDateExpr();

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

    return Array.from(map.values()).map((r) => {
      const total = Number(r.fuel_cost || 0) + Number(r.lube_cost || 0) + Number(r.parts_cost || 0)
        + Number(r.labor_cost || 0) + Number(r.downtime_cost || 0);
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
    });
  }

  function queryPeriodRunHoursByAsset(start, end) {
    return db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        a.category,
        COALESCE(SUM(dh.hours_run), 0) AS run_hours
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date BETWEEN ? AND ?
        AND dh.hours_run > 0
        ${andDailyHoursFleetHoursOnly("dh", "a")}
      GROUP BY a.id
    `).all(start, end).map((r) => ({
      asset_code: String(r.asset_code || ""),
      asset_name: String(r.asset_name || ""),
      category: String(r.category || "Unassigned"),
      run_hours: Number(Number(r.run_hours || 0).toFixed(1)),
    }));
  }

  function mergeAssetCostsWithRunHours(assetCosts, runHoursRows) {
    const runByCode = new Map(runHoursRows.map((r) => [r.asset_code, r]));
    const merged = new Map();

    for (const cost of assetCosts) {
      const code = String(cost.asset_code || "");
      const run = runByCode.get(code);
      const run_hours = Number(run?.run_hours || 0);
      const total_cost = Number(cost.total_cost || 0);
      merged.set(code, {
        ...cost,
        run_hours,
        cost_per_run_hour: run_hours > 0 ? Number((total_cost / run_hours).toFixed(2)) : null,
      });
      if (run) runByCode.delete(code);
    }

    for (const [code, run] of runByCode) {
      merged.set(code, {
        asset_code: code,
        asset_name: run.asset_name,
        category: run.category,
        fuel_cost: 0,
        lube_cost: 0,
        parts_cost: 0,
        labor_hours: 0,
        labor_cost: 0,
        downtime_hours: 0,
        downtime_cost: 0,
        total_cost: 0,
        run_hours: Number(run.run_hours || 0),
        cost_per_run_hour: null,
      });
    }

    return Array.from(merged.values())
      .filter((r) => r.total_cost > 0 || r.run_hours > 0)
      .sort((a, b) => Number(b.total_cost || 0) - Number(a.total_cost || 0));
  }

  function rollupFleetCostByCategory(assetRows) {
    const map = new Map();
    for (const r of assetRows) {
      const key = String(r.category || "Unassigned");
      if (!map.has(key)) {
        map.set(key, {
          category: key,
          run_hours: 0,
          fuel_cost: 0,
          lube_cost: 0,
          parts_cost: 0,
          labor_cost: 0,
          downtime_cost: 0,
          total_cost: 0,
        });
      }
      const row = map.get(key);
      row.run_hours += Number(r.run_hours || 0);
      row.fuel_cost += Number(r.fuel_cost || 0);
      row.lube_cost += Number(r.lube_cost || 0);
      row.parts_cost += Number(r.parts_cost || 0);
      row.labor_cost += Number(r.labor_cost || 0);
      row.downtime_cost += Number(r.downtime_cost || 0);
      row.total_cost += Number(r.total_cost || 0);
    }
    return Array.from(map.values())
      .map((r) => ({
        ...r,
        run_hours: Number(r.run_hours.toFixed(1)),
        fuel_cost: Number(r.fuel_cost.toFixed(2)),
        lube_cost: Number(r.lube_cost.toFixed(2)),
        parts_cost: Number(r.parts_cost.toFixed(2)),
        labor_cost: Number(r.labor_cost.toFixed(2)),
        downtime_cost: Number(r.downtime_cost.toFixed(2)),
        total_cost: Number(r.total_cost.toFixed(2)),
        cost_per_run_hour: r.run_hours > 0 ? Number((r.total_cost / r.run_hours).toFixed(2)) : null,
      }))
      .sort((a, b) => b.total_cost - a.total_cost);
  }

  function buildPeriodContractorFuelRows(start, end) {
    const defaults = costDefaults();
    const metricExpr = sqlFuelMetricModeExpr("a");
    const hireWhere = `
      COALESCE(a.active, 1) = 1
      AND (
        NULLIF(TRIM(COALESCE(a.hire_billing_mode, '')), '') IS NOT NULL
        OR LOWER(COALESCE(a.category, '')) LIKE '%contractor%'
        OR LOWER(COALESCE(a.category, '')) LIKE '%hire%'
        OR UPPER(COALESCE(a.asset_code, '')) LIKE 'BMP%'
        OR UPPER(COALESCE(a.asset_code, '')) LIKE 'PTT%'
        OR UPPER(COALESCE(a.asset_code, '')) IN ('E017', 'E018', 'E025', 'BR3', 'BW10', 'BW11')
      )
      AND ${sqlIncludeArchivedHireAssets("a")}
    `;

    const fuelRows = db.prepare(`
      SELECT
        a.id AS asset_id,
        a.asset_code,
        a.asset_name,
        a.category,
        a.hire_billing_mode,
        COALESCE(a.archived, 0) AS archived,
        ${metricExpr} AS metric_mode,
        COALESCE(SUM(fl.liters), 0) AS fuel_liters,
        COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS fuel_cost,
        COUNT(fl.id) AS fill_count
      FROM assets a
      INNER JOIN fuel_logs fl ON fl.asset_id = a.id AND fl.log_date BETWEEN ? AND ?
      WHERE ${hireWhere}
      GROUP BY a.id
      HAVING fuel_liters > 0
      ORDER BY a.asset_code ASC
    `).all(defaults.fuel_cost_per_liter_default, start, end);

    const getFuelLogsInRange = db.prepare(`
      SELECT
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
    const getFuelLogBeforeRange = db.prepare(`
      SELECT
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

    return fuelRows.map((r) => {
      const mode = String(r.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
      const logs = getFuelLogsInRange.all(r.asset_id, start, end);
      const prev = getFuelLogBeforeRange.get(r.asset_id, start);
      const fuelRun = getRunFromFuelRows(logs, prev, mode) || {};
      const km_run = Number(fuelRun.km_run || 0);
      const hours_run = Number(fuelRun.hours_run || 0);
      const fuel_cost = Number(Number(r.fuel_cost || 0).toFixed(2));
      const fuel_liters = Number(Number(r.fuel_liters || 0).toFixed(2));
      const run_value = mode === "km" ? km_run : hours_run;
      const cost_per_run = run_value > 0 ? Number((fuel_cost / run_value).toFixed(2)) : null;
      return {
        asset_code: String(r.asset_code || ""),
        asset_name: String(r.asset_name || ""),
        category: String(r.category || ""),
        contractor: inferHireContractorLabel(r.asset_code, r.category) || "Other",
        metric_mode: mode,
        fuel_liters,
        fuel_cost,
        fill_count: Number(r.fill_count || 0),
        km_run: Number(km_run.toFixed(2)),
        hours_run: Number(hours_run.toFixed(2)),
        run_value: Number(run_value.toFixed(2)),
        run_label: mode === "km" ? "km" : "hrs",
        cost_per_run,
        archived: Number(r.archived || 0),
      };
    }).sort((a, b) => b.fuel_cost - a.fuel_cost);
  }

  function rollupContractorFuelBySupplier(rows) {
    const map = new Map();
    for (const r of rows) {
      const key = String(r.contractor || "Other");
      if (!map.has(key)) {
        map.set(key, {
          contractor: key,
          asset_count: 0,
          fuel_liters: 0,
          fuel_cost: 0,
          hours_run: 0,
          km_run: 0,
        });
      }
      const row = map.get(key);
      row.asset_count += 1;
      row.fuel_liters += Number(r.fuel_liters || 0);
      row.fuel_cost += Number(r.fuel_cost || 0);
      row.hours_run += Number(r.hours_run || 0);
      row.km_run += Number(r.km_run || 0);
    }
    return Array.from(map.values())
      .map((r) => ({
        ...r,
        fuel_liters: Number(r.fuel_liters.toFixed(2)),
        fuel_cost: Number(r.fuel_cost.toFixed(2)),
        hours_run: Number(r.hours_run.toFixed(1)),
        km_run: Number(r.km_run.toFixed(1)),
      }))
      .sort((a, b) => b.fuel_cost - a.fuel_cost);
  }

  // =========================
  // STORES PART ORDERS PDF / XLSX (purchases & forecast)
  // =========================
  function parsePartOrdersPeriod(req) {
    const start = String(req.query?.start || req.query?.date_from || "").trim();
    const end = String(req.query?.end || req.query?.date_to || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
      return { error: "start and end are required (YYYY-MM-DD)" };
    }
    if (start > end) return { error: "start must be on or before end" };
    const site_code = String(req.headers["x-site-code"] || req.query?.site_code || "main").trim().toLowerCase() || "main";
    const status = String(req.query?.status || "").trim().toLowerCase();
    return { start, end, site_code, status };
  }

  function partOrderStatusLabel(status) {
    const s = String(status || "").toLowerCase();
    if (s === "on_order") return "On order";
    if (s === "warehouse_ready") return "Warehouse ready";
    if (s === "in_transit") return "In transit";
    if (s === "arrived") return "Arrived";
    if (s === "cancelled") return "Cancelled";
    return s || "—";
  }

  function fetchStoresPartOrdersForReport({ start, end, site_code, status }) {
    const where = ["LOWER(TRIM(COALESCE(o.site_code, 'main'))) = ?"];
    const params = [site_code];
    where.push("o.order_date >= ?");
    params.push(start);
    where.push("o.order_date <= ?");
    params.push(end);
    const allowed = new Set(["on_order", "warehouse_ready", "in_transit", "arrived", "cancelled"]);
    if (status && allowed.has(status)) {
      where.push("LOWER(COALESCE(o.status, 'on_order')) = ?");
      params.push(status);
    } else {
      where.push("LOWER(COALESCE(o.status, 'on_order')) <> 'cancelled'");
    }

    const rows = db.prepare(`
      SELECT
        o.id,
        o.site_code,
        o.part_code,
        o.part_name,
        o.qty,
        o.unit_cost,
        o.currency,
        o.supplier_name,
        o.po_number,
        o.requisition_number,
        o.invoice_number,
        o.current_location,
        o.warehouse_code,
        o.warehouse_date,
        o.warehouse_waiting_days,
        o.supplier_qty_received,
        o.supplier_outstanding_qty,
        o.sales_order,
        o.source_last_imported_at,
        a.asset_code,
        a.asset_name,
        o.order_date,
        o.expected_arrival_date,
        o.arrived_date,
        o.status,
        o.notes,
        o.created_by
      FROM stores_part_orders o
      LEFT JOIN assets a ON a.id = o.asset_id
      WHERE ${where.join(" AND ")}
      ORDER BY o.order_date DESC, o.id DESC
      LIMIT 5000
    `).all(...params).map((r) => {
      const qty = Number(r.qty || 0);
      const unit_cost = Number(r.unit_cost || 0);
      const line_total = Number((qty * unit_cost).toFixed(2));
      return { ...r, qty, unit_cost, line_total };
    });

    const summary = {
      on_order: { count: 0, qty: 0, value: 0 },
      warehouse_ready: { count: 0, qty: 0, value: 0 },
      in_transit: { count: 0, qty: 0, value: 0 },
      arrived: { count: 0, qty: 0, value: 0 },
      cancelled: { count: 0, qty: 0, value: 0 },
      total_forecast: 0,
      total_arrived: 0,
      total_pending: 0,
    };
    for (const row of rows) {
      const st = String(row.status || "on_order").toLowerCase();
      const bucket = summary[st];
      const value = row.line_total;
      if (bucket) {
        bucket.count += 1;
        bucket.qty += row.qty;
        bucket.value += value;
      }
      if (st === "arrived") summary.total_arrived += value;
      if (st === "on_order" || st === "warehouse_ready" || st === "in_transit") summary.total_pending += value;
      if (st !== "cancelled") summary.total_forecast += value;
    }
    for (const key of ["on_order", "warehouse_ready", "in_transit", "arrived", "cancelled"]) {
      summary[key].qty = Number(summary[key].qty.toFixed(2));
      summary[key].value = Number(summary[key].value.toFixed(2));
    }
    summary.total_forecast = Number(summary.total_forecast.toFixed(2));
    summary.total_arrived = Number(summary.total_arrived.toFixed(2));
    summary.total_pending = Number(summary.total_pending.toFixed(2));
    return { rows, summary };
  }

  // =========================
  // AML WEEKLY CHECK SHEET (creator-protected template)
  // =========================
  function buildAmlWeeklyExportRecords(weekEnding, siteCode) {
    const assetHasArchived = hasColumn("assets", "archived");
    const assetHasSiteCode = hasColumn("assets", "site_code");
    const assetWhere = [];
    const assetParams = [weekEnding, weekEnding, weekEnding, weekEnding];
    if (assetHasArchived) assetWhere.push("COALESCE(a.archived, 0) = 0");
    if (assetHasSiteCode && siteCode) {
      assetWhere.push("LOWER(COALESCE(NULLIF(a.site_code, ''), 'main')) = ?");
      assetParams.push(siteCode);
    }

    const assets = db.prepare(`
      SELECT
        a.id,
        a.asset_code,
        a.asset_name,
        COALESCE(a.active, 1) AS active,
        COALESCE((
          SELECT dh.closing_hours
          FROM daily_hours dh
          WHERE dh.asset_id = a.id
            AND dh.work_date <= ?
            AND dh.closing_hours IS NOT NULL
          ORDER BY dh.work_date DESC, dh.id DESC
          LIMIT 1
        ), (
          SELECT dh.opening_hours
          FROM daily_hours dh
          WHERE dh.asset_id = a.id
            AND dh.work_date <= ?
            AND dh.opening_hours IS NOT NULL
          ORDER BY dh.work_date DESC, dh.id DESC
          LIMIT 1
        )) AS meter_hours,
        COALESCE((
          SELECT dh.is_used
          FROM daily_hours dh
          WHERE dh.asset_id = a.id
            AND dh.work_date <= ?
          ORDER BY dh.work_date DESC, dh.id DESC
          LIMIT 1
        ), 0) AS latest_is_used,
        COALESCE((
          SELECT dh.hours_run
          FROM daily_hours dh
          WHERE dh.asset_id = a.id
            AND dh.work_date <= ?
          ORDER BY dh.work_date DESC, dh.id DESC
          LIMIT 1
        ), 0) AS latest_hours_run
      FROM assets a
      ${assetWhere.length ? `WHERE ${assetWhere.join(" AND ")}` : ""}
      ORDER BY a.asset_code COLLATE NOCASE ASC
    `).all(...assetParams);

    const breakdownAtWeekEnd = db.prepare(`
      SELECT
        b.id,
        b.description,
        b.component,
        b.critical,
        b.ets_repair_date,
        w.id AS work_order_id,
        w.status AS work_order_status,
        w.repair_progress
      FROM breakdowns b
      LEFT JOIN work_orders w ON w.id = b.primary_work_order_id
      WHERE b.asset_id = ?
        AND DATE(COALESCE(NULLIF(b.start_at, ''), b.breakdown_date)) <= ?
        AND (
          LOWER(COALESCE(b.status, 'open')) NOT IN ('closed', 'completed')
          OR DATE(COALESCE(b.end_at, w.closed_at, b.breakdown_date)) > ?
        )
      ORDER BY COALESCE(b.critical, 0) DESC,
        DATE(COALESCE(NULLIF(b.start_at, ''), b.breakdown_date)) DESC,
        b.id DESC
      LIMIT 1
    `);
    const offsiteAtWeekEnd = db.prepare(`
      SELECT
        r.id,
        r.sent_date,
        r.expected_return_date,
        r.actual_return_date,
        r.vendor,
        r.notes,
        r.repair_status
      FROM breakdown_offsite_repairs r
      WHERE r.asset_id = ?
        AND DATE(r.sent_date) <= ?
        AND (r.actual_return_date IS NULL OR DATE(r.actual_return_date) > ?)
      ORDER BY DATE(r.sent_date) DESC, r.id DESC
      LIMIT 1
    `);

    return assets.map((asset) => ({
      assetCode: asset.asset_code,
      meterHours: asset.meter_hours,
      active: asset.active,
      // A unit that logged production on the latest day is operational, even if an
      // older breakdown record remains open while its paperwork is being closed out.
      isOperational: Number(asset.latest_is_used) === 1 || Number(asset.latest_hours_run) > 0,
      breakdown: breakdownAtWeekEnd.get(asset.id, weekEnding, weekEnding) || null,
      offsite: offsiteAtWeekEnd.get(asset.id, weekEnding, weekEnding) || null,
    }));
  }

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

  async function buildMaintenanceExecutiveDeck({ period, label, site_code, requestHeaders = {} }) {
    const defaults = costDefaults();
    const laborRate = Number(defaults.labor_cost_per_hour_default || 35);
    const lubeDefault = Number(defaults.lube_cost_per_qty_default || 4);
    const akpChartColors = ["2563EB", "16A34A", "DC2626"];
    const akpBarOpts = { catAxisLabelRotate: -45, showLegend: true, legendPos: "b", chartColors: akpChartColors };
    const akpLineOpts = { catAxisLabelRotate: -45, showLegend: true, legendPos: "b", valAxisMinVal: 0, valAxisMaxVal: 100, chartColors: ["2563EB", "16A34A"] };

    // The live API requires a bearer session.  Internal inject requests do not
    // inherit it automatically, which previously made this deck silently use
    // an empty forecast even though the signed-in Weekly Forum showed costs.
    const insightsHeaders = maintenanceDeckInsightsHeaders(site_code, requestHeaders);
    const insightsInjected = await app.inject({
      method: "GET",
      url: `/api/maintenance/insights?start=${encodeURIComponent(period.start)}&end=${encodeURIComponent(period.end)}&near_due_hours=50&predictive_horizon_hours=100`,
      headers: insightsHeaders,
    });
    let maintenanceInsights = {};
    if (insightsInjected.statusCode < 400) {
      try {
        maintenanceInsights = JSON.parse(String(insightsInjected.payload || "{}"));
      } catch {
        maintenanceInsights = {};
      }
    }
    const insightsCostRows = Array.isArray(maintenanceInsights?.maintenance_cost)
      ? maintenanceInsights.maintenance_cost
      : [];
    const upcomingCostRows = Array.isArray(maintenanceInsights?.parts_planning?.upcoming_cost_forecasts)
      ? maintenanceInsights.parts_planning.upcoming_cost_forecasts
      : [];
    const hasSiteRainDays = hasTable("site_rain_days");
    const hasBreakdownLogs = hasTable("breakdown_downtime_logs");
    const hasManagerDamageReports = hasTable("manager_damage_reports");
    const hasManagerDamageReportPhotos = hasTable("manager_damage_report_photos");
    const hasManagerInspections = hasTable("manager_inspections");
    const rainRows = hasSiteRainDays
      ? db.prepare(`
      SELECT rain_date
      FROM site_rain_days
      WHERE site_code = ?
        AND rain_date BETWEEN ? AND ?
      ORDER BY rain_date ASC
    `).all(site_code, period.start, period.end)
      : [];
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
        ${andAssetExcludeLdv("a")}
      GROUP BY LOWER(IFNULL(a.category, 'uncategorized'))
      ORDER BY equipment_type ASC
    `).all(period.start, period.end);
    const availByTypeDowntime = hasBreakdownLogs
      ? db.prepare(`
      SELECT
        LOWER(IFNULL(a.category, 'uncategorized')) AS equipment_type,
        COALESCE(SUM(l.hours_down), 0) AS downtime_hours
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date BETWEEN ? AND ?
        ${andAssetFleetHoursOnly("a")}
        ${andAssetExcludeLdv("a")}
      GROUP BY LOWER(IFNULL(a.category, 'uncategorized'))
    `).all(period.start, period.end)
      : [];
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
    const hseSummary = hasManagerDamageReports
      ? db.prepare(`
      SELECT
        COUNT(*) AS reports,
        COALESCE(SUM(CASE WHEN COALESCE(hse_report_available, 0) = 1 THEN 1 ELSE 0 END), 0) AS hse_reports,
        COALESCE(SUM(CASE WHEN COALESCE(pending_investigation, 0) = 1 THEN 1 ELSE 0 END), 0) AS pending_investigation,
        COALESCE(SUM(CASE WHEN COALESCE(out_of_service, 0) = 1 THEN 1 ELSE 0 END), 0) AS out_of_service
      FROM manager_damage_reports
      WHERE report_date BETWEEN ? AND ?
    `).get(period.start, period.end)
      : { reports: 0, hse_reports: 0, pending_investigation: 0, out_of_service: 0 };
    const hseRows = hasManagerDamageReports
      ? db.prepare(`
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
    `).all(period.start, period.end)
      : [];
    const drPhotoReportCol = hasManagerDamageReportPhotos
      ? pickExistingColumn("manager_damage_report_photos", ["damage_report_id", "manager_damage_report_id", "report_id"], "damage_report_id")
      : "damage_report_id";
    const drPhotoPathCol = hasManagerDamageReportPhotos
      ? pickExistingColumn("manager_damage_report_photos", ["file_path", "photo_path", "path", "image_path", "url", "image_data"], "file_path")
      : "file_path";
    const hsePhotoRows = hasManagerDamageReports && hasManagerDamageReportPhotos
      ? db.prepare(`
      SELECT ${drPhotoPathCol} AS file_path
      FROM manager_damage_report_photos
      WHERE ${drPhotoReportCol} IN (
        SELECT id
        FROM manager_damage_reports
        WHERE report_date BETWEEN ? AND ?
      )
      ORDER BY id DESC
      LIMIT 4
    `).all(period.start, period.end)
      : [];
    const inspectionsSummary = hasManagerInspections
      ? db.prepare(`
      SELECT
        COUNT(*) AS inspections_done,
        COUNT(DISTINCT asset_id) AS assets_covered
      FROM manager_inspections
      WHERE inspection_date BETWEEN ? AND ?
    `).get(period.start, period.end)
      : { inspections_done: 0, assets_covered: 0 };
    const smColsDeckEarly = hasTable("stock_movements") ? db.prepare(`PRAGMA table_info(stock_movements)`).all() : [];
    const smHasCreatedAtDeckEarly = smColsDeckEarly.some((c) => String(c.name) === "created_at");
    const smDateExprDeck = smHasCreatedAtDeckEarly ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
    const miNotesColDeck = hasManagerInspections
      ? pickExistingColumn("manager_inspections", ["notes", "note", "remarks", "description"], "notes")
      : "notes";
    const miRecActionCol = hasColumn("manager_inspections", "recommended_action") ? "recommended_action" : null;
    const photoInspectionColDeck = hasTable("manager_inspection_photos")
      ? pickExistingColumn("manager_inspection_photos", ["inspection_id", "manager_inspection_id"], "inspection_id")
      : "inspection_id";
    const photoPathColDeck = hasTable("manager_inspection_photos")
      ? pickExistingColumn("manager_inspection_photos", ["file_path", "photo_path", "path", "image_path", "url"], "file_path")
      : "file_path";
    const inspectionsFaultRows = hasManagerInspections
      ? db.prepare(`
      SELECT
        mi.id,
        mi.inspection_date,
        a.asset_code,
        a.asset_name,
        COALESCE(mi.${miNotesColDeck}, '') AS inspection_notes,
        ${miRecActionCol ? `COALESCE(mi.${miRecActionCol}, '')` : "''"} AS recommended_action,
        COALESCE(mi.checklist_json, '[]') AS checklist_json,
        (
          SELECT ${photoPathColDeck}
          FROM manager_inspection_photos
          WHERE ${photoInspectionColDeck} = mi.id
          ORDER BY id ASC
          LIMIT 1
        ) AS photo_path
      FROM manager_inspections mi
      JOIN assets a ON a.id = mi.asset_id
      WHERE mi.inspection_date BETWEEN ? AND ?
      ORDER BY mi.inspection_date DESC, mi.id DESC
      LIMIT 8
    `).all(period.start, period.end)
      : [];
    // Aggregate per asset + lube type across the whole period so a month of logs fits one slide.
    const lubeByMachine = db.prepare(`
      SELECT
        MIN(ol.log_date) AS first_log_date,
        MAX(ol.log_date) AS last_log_date,
        COUNT(*) AS entries,
        a.asset_code,
        a.asset_name,
        COALESCE(NULLIF(TRIM(ol.oil_type), ''), '-') AS lube_type,
        MAX(COALESCE(NULLIF(TRIM(p.part_name), ''), '')) AS issue_part,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      LEFT JOIN parts p ON UPPER(TRIM(p.part_code)) = UPPER(TRIM(COALESCE(ol.oil_type, '')))
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.id, UPPER(TRIM(COALESCE(ol.oil_type, '')))
      ORDER BY lube_cost DESC, qty_total DESC, a.asset_code ASC
      LIMIT 16
    `).all(lubeDefault, period.start, period.end);
    // Totals must cover every log line in the period, not just the rows shown on the slide.
    const lubeTotalsRow = db.prepare(`
      SELECT
        COUNT(*) AS entries,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN ? AND ?
    `).get(lubeDefault, period.start, period.end);
    const lubeGroupCountRow = db.prepare(`
      SELECT COUNT(*) AS groups
      FROM (
        SELECT 1
        FROM oil_logs ol
        JOIN assets a ON a.id = ol.asset_id
        WHERE ol.log_date BETWEEN ? AND ?
        GROUP BY a.id, UPPER(TRIM(COALESCE(ol.oil_type, '')))
      )
    `).get(period.start, period.end);
    const lubePeriodTotals = {
      qty: Number(lubeTotalsRow?.qty_total || 0),
      cost: Number(lubeTotalsRow?.lube_cost || 0),
      entries: Number(lubeTotalsRow?.entries || 0),
      groups: Number(lubeGroupCountRow?.groups || 0),
    };
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
          ${andAssetExcludeLdv("a")}
        GROUP BY dh.work_date
      ),
      down_rows AS (
        ${hasBreakdownLogs
          ? `
        SELECT
          l.log_date AS day_key,
          COALESCE(SUM(l.hours_down), 0) AS downtime_hours
        FROM breakdown_downtime_logs l
        JOIN breakdowns b ON b.id = l.breakdown_id
        JOIN assets a ON a.id = b.asset_id
        WHERE l.log_date BETWEEN ? AND ?
          ${andAssetFleetHoursOnly("a")}
          ${andAssetExcludeLdv("a")}
        GROUP BY l.log_date
        `
          : `
        SELECT
          NULL AS day_key,
          0 AS downtime_hours
        WHERE 1 = 0
        `
        }
      )
      SELECT
        r.day_key,
        COALESCE(r.scheduled_hours, 0) AS scheduled_hours,
        COALESCE(r.run_hours, 0) AS run_hours,
        COALESCE(d.downtime_hours, 0) AS downtime_hours
      FROM run_rows r
      LEFT JOIN down_rows d ON d.day_key = r.day_key
      ORDER BY r.day_key ASC
    `).all(...(hasBreakdownLogs ? [period.start, period.end, period.start, period.end] : [period.start, period.end]));
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
          AND dh.is_used = 1
        WHERE a.active = 1
          AND a.is_standby = 0
          ${andAssetExcludeLdv("a")}
        GROUP BY a.id
      ),
      down_rows AS (
        ${hasBreakdownLogs
          ? `
        SELECT
          b.asset_id,
          COALESCE(SUM(l.hours_down), 0) AS downtime_hours
        FROM breakdown_downtime_logs l
        JOIN breakdowns b ON b.id = l.breakdown_id
        WHERE l.log_date BETWEEN ? AND ?
        GROUP BY b.asset_id
        `
          : `
        SELECT
          NULL AS asset_id,
          0 AS downtime_hours
        WHERE 1 = 0
        `
        }
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
    `).all(...(hasBreakdownLogs ? [period.start, period.end, period.start, period.end] : [period.start, period.end])).map((r) => {
      const s = Number(r.scheduled_hours || 0);
      const run = Math.max(0, Math.min(Number(r.run_hours || 0), s));
      const down = Math.max(0, Math.min(Number(r.downtime_hours || 0), s));
      const availPct = s > 0 ? ((Math.max(0, s - down) / s) * 100) : 0;
      const utilPct = s > 0 ? ((run / s) * 100) : 0;
      return {
        asset_code: String(r.asset_code || ""),
        asset_name: String(r.asset_name || ""),
        category: String(r.category || "Uncategorized"),
        scheduled_hours: s,
        run_hours: run,
        downtime_hours: down,
        availability_pct: Number(availPct.toFixed(2)),
        utilization_pct: Number(utilPct.toFixed(2)),
      };
    });
    const assetPerfSorted = [...assetPerfRows]
      .filter((r) => Number(r.scheduled_hours || 0) > 0)
      .sort((a, b) => b.utilization_pct - a.utilization_pct);
    const topAssets = assetPerfSorted.slice(0, 2);
    const assetsWithData = assetPerfSorted.filter(
      (r) => Number(r.run_hours || 0) > 0 || Number(r.downtime_hours || 0) > 0,
    );
    const bottomAssets = [...assetsWithData].sort((a, b) => a.utilization_pct - b.utilization_pct).slice(0, 2);
    const midStart = Math.max(0, Math.floor((assetPerfSorted.length - 2) / 2));
    const midAssets = assetPerfSorted.slice(midStart, midStart + 2);
    const fleetScheduled = assetPerfSorted.reduce((s, r) => s + Number(r.scheduled_hours || 0), 0);
    const fleetRun = assetPerfSorted.reduce((s, r) => s + Number(r.run_hours || 0), 0);
    const fleetDown = assetPerfSorted.reduce((s, r) => s + Number(r.downtime_hours || 0), 0);
    const fleetAvailPct = fleetScheduled > 0
      ? Number((((Math.max(0, fleetScheduled - fleetDown)) / fleetScheduled) * 100).toFixed(2))
      : null;
    const fleetUtilPct = fleetScheduled > 0
      ? Number(((fleetRun / fleetScheduled) * 100).toFixed(2))
      : null;
    const hasBreakdownRepairLabor = hasTable("breakdown_repair_labor");
    const woCostDetailRows = hasTable("work_orders")
      ? db.prepare(`
      WITH wo_parts AS (
        SELECT
          CAST(SUBSTR(sm.reference, 13) AS INTEGER) AS wo_id,
          COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS parts_cost
        FROM stock_movements sm
        JOIN parts p ON p.id = sm.part_id
        WHERE sm.movement_type = 'out'
          AND sm.reference LIKE 'work_order:%'
          AND ${smDateExprDeck} BETWEEN ? AND ?
        GROUP BY wo_id
      )
      SELECT *
      FROM (
        SELECT
          w.id AS wo_id,
          a.asset_code,
          a.asset_name,
          COALESCE(a.category, '') AS category,
          LOWER(COALESCE(w.source, '')) AS source,
          COALESCE(b.description, ${hasBreakdownRepairLabor ? "brl.notes," : ""} w.completion_notes, w.repair_progress, '-') AS reason,
          COALESCE(wp.parts_cost, 0) AS parts_cost,
          CASE
            WHEN LOWER(COALESCE(w.source, '')) = 'breakdown' THEN
              COALESCE(${hasBreakdownRepairLabor ? "NULLIF(brl.labor_hours, 0)," : ""} w.labor_hours, 0)
              * COALESCE(NULLIF(w.labor_rate_per_hour, 0), ?)
            ELSE COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)
          END AS labor_cost
        FROM work_orders w
        JOIN assets a ON a.id = w.asset_id
        LEFT JOIN breakdowns b ON b.id = w.reference_id AND LOWER(COALESCE(w.source, '')) = 'breakdown'
        ${hasBreakdownRepairLabor
          ? "LEFT JOIN breakdown_repair_labor brl ON brl.breakdown_id = b.id"
          : ""}
        LEFT JOIN wo_parts wp ON wp.wo_id = w.id
        WHERE DATE(COALESCE(w.completed_at, w.closed_at, w.opened_at)) BETWEEN ? AND ?
          AND LOWER(COALESCE(w.source, '')) IN ('breakdown', 'service')
      ) wo_costs
      WHERE wo_costs.parts_cost > 0 OR wo_costs.labor_cost > 0
      ORDER BY (wo_costs.parts_cost + wo_costs.labor_cost) DESC, wo_costs.wo_id DESC
      LIMIT 14
    `).all(period.start, period.end, laborRate, laborRate, period.start, period.end)
      : [];
    const woCostTotals = woCostDetailRows.reduce(
      (acc, r) => {
        acc.parts_cost += Number(r.parts_cost || 0);
        acc.labor_cost += Number(r.labor_cost || 0);
        acc.total_cost += Number(r.parts_cost || 0) + Number(r.labor_cost || 0);
        return acc;
      },
      { parts_cost: 0, labor_cost: 0, total_cost: 0 },
    );
    const insightsLaborTotal = insightsCostRows.reduce((s, r) => s + Number(r.labor_cost || 0), 0);
    const insightsPartsTotal = insightsCostRows.reduce((s, r) => s + Number(r.parts_cost || 0), 0);
    const insightsMaintTotal = insightsCostRows.reduce((s, r) => s + Number(r.total_cost || 0), 0);
    const contractorFuelRows = buildPeriodContractorFuelRows(period.start, period.end);
    const contractorBySupplier = rollupContractorFuelBySupplier(contractorFuelRows);
    const contractorFuelTotal = contractorFuelRows.reduce(
      (acc, r) => {
        acc.fuel_liters += Number(r.fuel_liters || 0);
        acc.fuel_cost += Number(r.fuel_cost || 0);
        acc.hours_run += Number(r.hours_run || 0);
        acc.km_run += Number(r.km_run || 0);
        return acc;
      },
      { fuel_liters: 0, fuel_cost: 0, hours_run: 0, km_run: 0 },
    );

    const siteCandidates = Array.from(new Set([String(site_code || "main").trim().toLowerCase() || "main", "default"]));
    const siteMarks = siteCandidates.map(() => "?").join(", ");
    const storesPartsRows = hasTable("stores_part_orders")
      ? db.prepare(`
      SELECT
        part_code,
        part_name,
        qty,
        unit_cost,
        supplier_name,
        po_number,
        order_date,
        expected_arrival_date,
        status,
        notes
      FROM stores_part_orders
      WHERE LOWER(TRIM(COALESCE(site_code, 'main'))) IN (${siteMarks})
        AND LOWER(COALESCE(status, 'on_order')) IN ('on_order', 'in_transit')
      ORDER BY
        CASE LOWER(COALESCE(status, 'on_order')) WHEN 'in_transit' THEN 0 ELSE 1 END,
        DATE(COALESCE(expected_arrival_date, order_date)) ASC,
        id DESC
      LIMIT 16
    `).all(...siteCandidates)
      : [];
    const storesPartsTotals = storesPartsRows.reduce(
      (acc, r) => {
        const line = Number(r.qty || 0) * Number(r.unit_cost || 0);
        acc.qty += Number(r.qty || 0);
        acc.value += line;
        const st = String(r.status || "").toLowerCase();
        if (st === "in_transit") acc.in_transit += line;
        if (st === "on_order") acc.on_order += line;
        return acc;
      },
      { qty: 0, value: 0, on_order: 0, in_transit: 0 },
    );
    const partsTrackingRows = storesPartsRows;
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
        COALESCE(MAX(s.avg_lph), 0) AS avg_lph_benchmark,
        COALESCE(MAX((d.liters / d.hours_run) - s.avg_lph), 0) AS peak_variance_lph,
        COALESCE(MAX(CASE WHEN s.avg_lph > 0 THEN (((d.liters / d.hours_run) / s.avg_lph) - 1) * 100 ELSE 0 END), 0) AS peak_variance_pct
      FROM daily d
      JOIN stats s ON s.asset_id = d.asset_id
      WHERE d.hours_run > 0
        AND s.avg_lph IS NOT NULL
        AND (d.liters / d.hours_run) > (s.avg_lph * 1.35)
      GROUP BY d.asset_code
      ORDER BY anomaly_days DESC, d.asset_code ASC
      LIMIT 12
    `).all(period.start, period.end);
    const upcomingKitTotal = upcomingCostRows.reduce((s, r) => s + Number(r?.forecast?.est_service_kit_cost || 0), 0);
    const upcomingLaborTotal = upcomingCostRows.reduce((s, r) => s + Number(r?.forecast?.est_labor_cost || 0), 0);
    const upcomingGrandTotal = upcomingCostRows.reduce((s, r) => s + Number(r?.forecast?.est_total_cost || 0), 0);
    const upcomingUnpricedRows = upcomingCostRows.filter((r) => r?.needs_manual_input);
    const isAllInUpcomingEstimate = (r) => String(r?.forecast?.cost_source || "") === "manual_all_in_estimate";
    const upcomingAllInTotal = upcomingCostRows
      .filter(isAllInUpcomingEstimate)
      .reduce((sum, r) => sum + Number(r?.forecast?.est_total_cost || 0), 0);
    const upcomingCostCell = (r, key) => {
      if (r?.needs_manual_input) return "—";
      if (isAllInUpcomingEstimate(r) && key !== "est_total_cost") return "—";
      return Number(r?.forecast?.[key] || 0).toFixed(2);
    };
    const upcomingSourceLabel = serviceCostSourceLabel;
    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "IRONLOG";
    pptx.subject = "Maintenance executive report";
    pptx.title = `Maintenance Executive - ${label}`;
    const deckNavy = "12355B";
    const deckTeal = "0F766E";
    const deckBlue = "2563EB";
    const deckMist = "F4F8FC";
    const deckLine = "C9D8E6";
    const deckText = "1F2937";
    const deckMuted = "64748B";
    const deckHeaderFont = "Aptos Display";
    const deckBodyFont = "Aptos";
    const tableHeader = (...labels) => labels.map((text) => ({
      text,
      options: { bold: true, color: "FFFFFF", fill: { color: deckNavy }, fontFace: deckBodyFont },
    }));
    const tableOptions = { border: { pt: 0.75, color: deckLine }, color: deckText, fontFace: deckBodyFont };
    const tableHeightFor = (rows, min = 0.9, max = 4.8, rowHeight = 0.43) =>
      Math.max(min, Math.min(max, 0.36 + (Math.max(1, rows) * rowHeight)));
    const addExecutiveFrame = (slide) => {
      slide.background = { color: deckMist };
      slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: 13.333, h: 0.08,
        line: { color: deckTeal, transparency: 100 },
        fill: { color: deckTeal },
      });
      slide.addShape(pptx.ShapeType.line, {
        x: 0.4, y: 7.1, w: 12.45, h: 0,
        line: { color: deckLine, pt: 0.6 },
      });
      slide.addText(`IRONLOG | Maintenance Executive Pack | ${period.start} to ${period.end}`, {
        x: 0.4, y: 7.17, w: 10.5, h: 0.18, fontFace: deckBodyFont, fontSize: 7.5, color: deckMuted,
      });
      slide.addText(String(site_code || "main").toUpperCase(), {
        x: 11.15, y: 7.17, w: 1.7, h: 0.18, fontFace: deckBodyFont, fontSize: 7.5, bold: true, color: deckMuted, align: "right",
      });
    };
    const addMetric = (slide, x, labelText, valueText, accent) => {
      slide.addShape(pptx.ShapeType.roundRect, {
        x, y: 4.95, w: 3.82, h: 1.1,
        rectRadius: 0.08,
        line: { color: accent, transparency: 100 },
        fill: { color: "FFFFFF" },
        shadow: { type: "outer", color: "AAB7C4", opacity: 0.14, blur: 1, angle: 45, distance: 1 },
      });
      slide.addShape(pptx.ShapeType.rect, {
        x, y: 4.95, w: 0.08, h: 1.1,
        line: { color: accent, transparency: 100 }, fill: { color: accent },
      });
      slide.addText(labelText, {
        x: x + 0.24, y: 5.15, w: 3.35, h: 0.2, fontFace: deckBodyFont, fontSize: 8.5, bold: true, color: deckMuted,
      });
      slide.addText(valueText, {
        x: x + 0.24, y: 5.46, w: 3.35, h: 0.33, fontFace: deckHeaderFont, fontSize: 20, bold: true, color: accent,
      });
    };
    const s1 = pptx.addSlide();
    s1.background = { color: deckNavy };
    s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.12, line: { color: deckTeal, transparency: 100 }, fill: { color: deckTeal } });
    s1.addText("IRONLOG", { x: 0.55, y: 0.48, w: 2.1, h: 0.26, fontFace: deckBodyFont, fontSize: 12, bold: true, color: "84E1D2", charSpacing: 1.1 });
    s1.addText("Maintenance Executive Pack", { x: 0.55, y: 1.2, w: 8.2, h: 0.75, fontFace: deckHeaderFont, fontSize: 34, bold: true, color: "FFFFFF" });
    s1.addText(`Weekly forum | ${period.start} to ${period.end}`, { x: 0.58, y: 2.08, w: 6.7, h: 0.28, fontFace: deckBodyFont, fontSize: 15, color: "D9E6F2" });
    s1.addText("A concise review of fleet performance, maintenance risk, upcoming work, stores, lubrication and assurance activity.", { x: 0.58, y: 2.75, w: 8.7, h: 0.7, fontFace: deckBodyFont, fontSize: 17, color: "EAF2F8", breakLine: false });
    s1.addShape(pptx.ShapeType.line, { x: 0.58, y: 4.35, w: 12.1, h: 0, line: { color: "5C7592", pt: 0.75 } });
    addMetric(s1, 0.58, "FLEET AVAILABILITY", fleetAvailPct == null ? "No data" : `${fleetAvailPct.toFixed(1)}%`, deckTeal);
    addMetric(s1, 4.76, "BREAKDOWNS LOGGED", String(breakdownCount), "F59E0B");
    addMetric(s1, 8.94, "UPCOMING MAINTENANCE", `$${fmtNum(upcomingGrandTotal, 0)}`, deckBlue);
    s1.addText(`${String(site_code || "main").toUpperCase()} SITE`, { x: 10.6, y: 0.55, w: 2.15, h: 0.22, fontFace: deckBodyFont, fontSize: 10, bold: true, color: "84E1D2", align: "right" });
    const s2 = pptx.addSlide();
    addExecutiveFrame(s2);
    s2.addText("1) Safety and HSE", { x: 0.4, y: 0.25, w: 12.4, h: 0.45, fontFace: deckHeaderFont, fontSize: 22, bold: true, color: deckNavy });
    s2.addText(`Damage reports: ${Number(hseSummary?.reports || 0)} | HSE reports: ${Number(hseSummary?.hse_reports || 0)} | Pending: ${Number(hseSummary?.pending_investigation || 0)} | Out of service: ${Number(hseSummary?.out_of_service || 0)}`, { x: 0.5, y: 0.8, w: 12.2, h: 0.3, fontFace: deckBodyFont, fontSize: 11, bold: true, color: deckText });
    s2.addTable(
      [
        tableHeader("Date", "Asset", "Severity", "Fault", "Action"),
        ...(hseRows.length
          ? hseRows.slice(0, 5).map((r) => [String(r.report_date || "-"), String(r.asset_code || "-"), String(r.severity || "-"), compactCell(String(r.damage_description || r.damage_location || "-"), 44), compactCell(String(r.immediate_action || "-"), 40)])
          : [["-", "-", "-", "No HSE damage reports in selected period", "-"]]),
      ],
      { x: 0.45, y: 1.2, w: 12.35, h: tableHeightFor(1 + (hseRows.length ? Math.min(5, hseRows.length) : 1), 1.0, 2.2), fontSize: 9.5, ...tableOptions }
    );
    const hsePhotoAbs = hsePhotoRows
      .map((p) => resolveStorageAbs(String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "")))
      .filter((p) => p && fs.existsSync(p));
    const photoSlots = [
      { x: 0.5, y: 3.7, w: 2.95, h: 2.2 },
      { x: 3.65, y: 3.7, w: 2.95, h: 2.2 },
      { x: 6.8, y: 3.7, w: 2.95, h: 2.2 },
      { x: 9.95, y: 3.7, w: 2.95, h: 2.2 },
    ];
    hsePhotoAbs.slice(0, 4).forEach((abs, idx) => {
      s2.addImage({ path: abs, ...photoSlots[idx] });
    });
    const s3 = pptx.addSlide();
    addExecutiveFrame(s3);
    s3.addText("2) Plant Performance", { x: 0.4, y: 0.25, w: 12.4, h: 0.45, fontFace: deckHeaderFont, fontSize: 22, bold: true, color: deckNavy });
    s3.addText(
      `Fleet (excl. LDVs): Availability ${fleetAvailPct == null ? "—" : `${fleetAvailPct.toFixed(1)}%`} | Utilization ${fleetUtilPct == null ? "—" : `${fleetUtilPct.toFixed(1)}%`} | Scheduled ${fleetScheduled.toFixed(1)}h | Run ${fleetRun.toFixed(1)}h | Downtime ${fleetDown.toFixed(1)}h`,
      { x: 0.45, y: 0.72, w: 12.3, h: 0.35, fontFace: deckBodyFont, fontSize: 11, bold: true, color: deckText },
    );
    s3.addChart(pptx.ChartType.bar, [
      { name: "Scheduled", labels: dailyKpiRows.map((r) => String(r.day_key || "").slice(5)), values: dailyKpiRows.map((r) => Number(r.scheduled_hours || 0)) },
      { name: "Run", labels: dailyKpiRows.map((r) => String(r.day_key || "").slice(5)), values: dailyKpiRows.map((r) => Number(r.run_hours || 0)) },
      { name: "Downtime", labels: dailyKpiRows.map((r) => String(r.day_key || "").slice(5)), values: dailyKpiRows.map((r) => Number(r.downtime_hours || 0)) },
    ], { x: 0.45, y: 1.1, w: 6.0, h: 2.05, ...akpBarOpts });
    s3.addChart(pptx.ChartType.line, [
      { name: "Availability %", labels: availabilityTrend.map((r) => r.label), values: availabilityTrend.map((r) => r.value) },
      { name: "Utilization %", labels: utilizationTrend.map((r) => r.label), values: utilizationTrend.map((r) => r.value) },
    ], { x: 6.75, y: 1.1, w: 6.0, h: 2.05, ...akpLineOpts });
    s3.addTable(
      [
        tableHeader("Highest performing equipment by type", "Avail %", "Util %"),
        ...availabilityByType
          .filter((r) => !String(r.equipment_type || "").toLowerCase().includes("ldv"))
          .slice()
          .sort((a, b) => Number(b.utilization_pct || 0) - Number(a.utilization_pct || 0))
          .slice(0, 6)
          .map((r) => [String(r.equipment_type || "-"), `${Number(r.availability_pct || 0).toFixed(2)}%`, `${Number(r.utilization_pct || 0).toFixed(2)}%`]),
      ],
      { x: 0.45, y: 3.35, w: 4.05, h: tableHeightFor(1 + Math.min(6, availabilityByType.length), 0.9, 3.1), fontSize: 9.5, ...tableOptions }
    );
    const perfCompact = (r) => `${String(r.asset_code || "-")} (${Number(r.utilization_pct || 0).toFixed(1)}%)`;
    s3.addTable(
      [
        tableHeader("Top 2 Assets"),
        ...(topAssets.length ? topAssets.map((r) => [perfCompact(r)]) : [["-"]]),
        [{ text: "Mid 2 Assets", options: { bold: true, color: deckNavy, fill: { color: "E6F4F1" }, fontFace: deckBodyFont } }],
        ...(midAssets.length ? midAssets.map((r) => [perfCompact(r)]) : [["-"]]),
        [{ text: "Lowest 2 (with data)", options: { bold: true, color: "9F1239", fill: { color: "FDECEC" }, fontFace: deckBodyFont } }],
        ...(bottomAssets.length ? bottomAssets.map((r) => [perfCompact(r)]) : [["-"]]),
      ],
      {
        x: 4.7, y: 3.35, w: 2.85,
        h: tableHeightFor(
          3
            + (topAssets.length || 1)
            + (midAssets.length || 1)
            + (bottomAssets.length || 1),
          2.4,
          3.1,
        ),
        fontSize: 10,
        ...tableOptions,
      }
    );
    s3.addChart(pptx.ChartType.bar, [
      { name: "Utilization %", labels: [...topAssets, ...midAssets, ...bottomAssets].map((r) => String(r.asset_code || "")), values: [...topAssets, ...midAssets, ...bottomAssets].map((r) => Number(r.utilization_pct || 0)) },
    ], { x: 7.75, y: 3.35, w: 5.0, h: 3.1, showLegend: false, chartColors: ["16A34A"], valAxisMinVal: 0, valAxisMaxVal: 100 });
    const s4 = pptx.addSlide();
    addExecutiveFrame(s4);
    s4.addText("3) Breakdown and Maintenance Costs", { x: 0.4, y: 0.25, w: 12.4, h: 0.45, fontFace: deckHeaderFont, fontSize: 20, bold: true, color: deckNavy });
    s4.addText(
      `Labor from Asset Insights: $${fmtNum(insightsLaborTotal, 2)} | Parts: $${fmtNum(insightsPartsTotal, 2)} | Maintenance total: $${fmtNum(insightsMaintTotal, 2)}`,
      { x: 0.45, y: 0.72, w: 12.2, h: 0.3, fontFace: deckBodyFont, fontSize: 10, bold: true, color: deckText },
    );
    s4.addTable(
      [
        tableHeader("WO", "Asset", "Type", "Reason", "Parts", "Labor", "Total"),
        ...(woCostDetailRows.length
          ? woCostDetailRows.map((r) => {
            const parts = Number(r.parts_cost || 0);
            const labor = Number(r.labor_cost || 0);
            return [
              String(r.wo_id || "-"),
              `${r.asset_code} ${compactCell(r.asset_name, 14)}`,
              String(r.source || "-"),
              compactCell(String(r.reason || "-"), 34),
              parts.toFixed(2),
              labor.toFixed(2),
              (parts + labor).toFixed(2),
            ];
          })
          : insightsCostRows.slice(0, 10).map((r) => [
            "-",
            `${r.asset_code} ${compactCell(r.asset_name, 14)}`,
            "summary",
            "Asset Insights maintenance cost",
            Number(r.parts_cost || 0).toFixed(2),
            Number(r.labor_cost || 0).toFixed(2),
            Number(r.total_cost || 0).toFixed(2),
          ])),
      ],
      {
        x: 0.35, y: 1.05, w: 12.6,
        h: tableHeightFor(1 + (woCostDetailRows.length || Math.min(10, insightsCostRows.length)), 1.2, 4.55),
        fontSize: 9,
        ...tableOptions,
      }
    );
    s4.addText(
      `Work-order totals: Parts $${fmtNum(woCostTotals.parts_cost, 2)} | Labor $${fmtNum(woCostTotals.labor_cost, 2)} | Total $${fmtNum(woCostTotals.total_cost, 2)}`,
      { x: 0.45, y: 5.75, w: 12.2, h: 0.35, fontSize: 11, bold: true },
    );
    if (contractorFuelRows.length) {
      const sContractor = pptx.addSlide();
      addExecutiveFrame(sContractor);
      sContractor.addText("3b) Contractor Fuel Costs", { x: 0.4, y: 0.25, w: 12.4, h: 0.45, fontFace: deckHeaderFont, fontSize: 20, bold: true, color: deckNavy });
      sContractor.addText(
        [
          `Contractor assets: ${contractorFuelRows.length}`,
          `Fuel liters: ${fmtNum(contractorFuelTotal.fuel_liters, 1)}`,
          `Fuel cost: $${fmtNum(contractorFuelTotal.fuel_cost, 2)}`,
          `Run hrs: ${fmtNum(contractorFuelTotal.hours_run, 1)}`,
          `Run km: ${fmtNum(contractorFuelTotal.km_run, 1)}`,
          contractorFuelTotal.hours_run > 0 ? `Avg $/hr: $${fmtNum(contractorFuelTotal.fuel_cost / contractorFuelTotal.hours_run, 2)}` : null,
          contractorFuelTotal.km_run > 0 ? `Avg $/km: $${fmtNum(contractorFuelTotal.fuel_cost / contractorFuelTotal.km_run, 2)}` : null,
        ].filter(Boolean).join("  |  "),
        { x: 0.45, y: 0.72, w: 12.2, h: 0.35, fontFace: deckBodyFont, fontSize: 10, bold: true, color: deckText },
      );
      sContractor.addTable(
        [
          tableHeader("Supplier", "Assets", "Liters", "Fuel $", "Run hrs", "Run km", "$/hr", "$/km"),
          ...contractorBySupplier.map((r) => [
            String(r.contractor || "-"),
            fmtNum(r.asset_count, 0),
            fmtNum(r.fuel_liters, 1),
            fmtNum(r.fuel_cost, 2),
            fmtNum(r.hours_run, 1),
            fmtNum(r.km_run, 1),
            r.hours_run > 0 ? fmtNum(r.fuel_cost / r.hours_run, 2) : "—",
            r.km_run > 0 ? fmtNum(r.fuel_cost / r.km_run, 2) : "—",
          ]),
        ],
        { x: 0.35, y: 1.15, w: 12.6, h: tableHeightFor(1 + contractorBySupplier.length, 1.0, 2.0), fontSize: 9.5, ...tableOptions },
      );
      sContractor.addTable(
        [
          tableHeader("Asset", "Supplier", "Mode", "Liters", "Fuel $", "Run", "Unit", "$/unit", "Fills"),
          ...contractorFuelRows.slice(0, 14).map((r) => [
            String(r.asset_code || "-"),
            String(r.contractor || "-"),
            r.metric_mode === "km" ? "km" : "hrs",
            fmtNum(r.fuel_liters, 1),
            fmtNum(r.fuel_cost, 2),
            fmtNum(r.run_value, 1),
            String(r.run_label || "-"),
            r.cost_per_run == null ? "N/A" : fmtNum(r.cost_per_run, 2),
            fmtNum(r.fill_count, 0),
          ]),
        ],
        { x: 0.35, y: 3.35, w: 12.6, h: tableHeightFor(1 + Math.min(14, contractorFuelRows.length), 1.0, 2.55), fontSize: 9, ...tableOptions },
      );
      sContractor.addText(
        `Contractor fuel total: $${fmtNum(contractorFuelTotal.fuel_cost, 2)} (${fmtNum(contractorFuelTotal.fuel_liters, 1)} L from FAMS meter run)`,
        { x: 0.45, y: 6.05, w: 12.2, h: 0.35, fontSize: 11, bold: true },
      );
    }
    const s4b = pptx.addSlide();
    addExecutiveFrame(s4b);
    s4b.addText("4) Planned Upcoming Costs", { x: 0.4, y: 0.25, w: 12.4, h: 0.45, fontFace: deckHeaderFont, fontSize: 20, bold: true, color: deckNavy });
    const upcomingTableH = tableHeightFor(1 + (upcomingCostRows.length ? Math.min(12, upcomingCostRows.length) : 1), 1.2, 4.8);
    s4b.addTable(
      [
        tableHeader("Asset", "Service", "Remaining hours", "Status", "Kit $", "Labor $", "Total $", "Source"),
        ...(upcomingCostRows.length
          ? upcomingCostRows.slice(0, 12).map((r) => [
            `${String(r.asset_code || "-")} ${compactCell(String(r.asset_name || ""), 14)}`,
            compactCell(String(r.service_name || "-"), 18),
            Number(r.remaining_hours || 0).toFixed(1),
            String(r.status || "-"),
            upcomingCostCell(r, "est_service_kit_cost"),
            upcomingCostCell(r, "est_labor_cost"),
            upcomingCostCell(r, "est_total_cost"),
            compactCell(upcomingSourceLabel(r), 14),
          ])
          : [["-", "No upcoming services in forecast window", "-", "-", "-", "-", "-", "-"]]),
      ],
      { x: 0.35, y: 0.95, w: 12.6, h: upcomingTableH, fontSize: 9.2, ...tableOptions }
    );
    s4b.addText(
      `Priced total: Kit $${fmtNum(upcomingKitTotal, 2)} | Labor $${fmtNum(upcomingLaborTotal, 2)}${upcomingAllInTotal ? ` | All-in $${fmtNum(upcomingAllInTotal, 2)}` : ""} | Upcoming maintenance $${fmtNum(upcomingGrandTotal, 2)}${upcomingUnpricedRows.length ? ` | ${upcomingUnpricedRows.length} awaiting pricing` : ""}`,
      { x: 0.45, y: Math.min(5.9, 0.95 + upcomingTableH + 0.2), w: 12.2, h: 0.35, fontFace: deckBodyFont, fontSize: 11, bold: true, color: deckText },
    );
    const s5 = pptx.addSlide();
    addExecutiveFrame(s5);
    s5.addText("5) Parts Tracking", { x: 0.4, y: 0.25, w: 12.4, h: 0.45, fontFace: deckHeaderFont, fontSize: 22, bold: true, color: deckNavy });
    s5.addText(
      `On order: $${fmtNum(storesPartsTotals.on_order, 2)} | En route (in transit): $${fmtNum(storesPartsTotals.in_transit, 2)} | Combined: $${fmtNum(storesPartsTotals.value, 2)}`,
      { x: 0.45, y: 0.72, w: 12.2, h: 0.3, fontFace: deckBodyFont, fontSize: 10, bold: true, color: deckText },
    );
    const partsTableH = tableHeightFor(1 + (partsTrackingRows.length || 1), 1.2, 5.0);
    s5.addTable(
      [
        tableHeader("Status", "Part", "Description", "Qty", "Line $", "Supplier", "PO #", "ETA"),
        ...(partsTrackingRows.length
          ? partsTrackingRows.map((r) => {
            const st = String(r.status || "").toLowerCase() === "in_transit" ? "En route" : "On order";
            const line = Number(r.qty || 0) * Number(r.unit_cost || 0);
            return [
              st,
              String(r.part_code || "-"),
              compactCell(String(r.part_name || r.notes || "-"), 24),
              fmtNum(Number(r.qty || 0), 1),
              fmtNum(line, 2),
              compactCell(String(r.supplier_name || "-"), 14),
              compactCell(String(r.po_number || "-"), 10),
              String(r.expected_arrival_date || r.order_date || "-"),
            ];
          })
          : [["-", "No parts on order or en route", "-", "-", "-", "-", "-", "-"]]),
      ],
      { x: 0.35, y: 1.05, w: 12.6, h: partsTableH, fontSize: 8.8, ...tableOptions }
    );
    const s7 = pptx.addSlide();
    addExecutiveFrame(s7);
    s7.addText("6) Lubrication", { x: 0.4, y: 0.25, w: 12.4, h: 0.45, fontFace: deckHeaderFont, fontSize: 22, bold: true, color: deckNavy });
    s7.addText(
      `Period ${period.start} to ${period.end}`,
      { x: 0.45, y: 0.72, w: 12.3, h: 0.25, fontFace: deckBodyFont, fontSize: 10, color: deckMuted },
    );
    const lubeTableH = tableHeightFor(1 + (lubeByMachine.length || 1), 1.2, 4.75);
    s7.addTable(
      [
        tableHeader("Dates", "Asset", "Lube type", "Issued part", "Logs", "Qty", "Cost"),
        ...(lubeByMachine.length
          ? lubeByMachine.map((r) => {
            const first = String(r.first_log_date || "").trim();
            const last = String(r.last_log_date || "").trim();
            return [
              first && last && first !== last ? `${first} → ${last}` : (last || first || "-"),
              `${String(r.asset_code || "")} ${compactCell(r.asset_name || "", 16)}`,
              compactCell(String(r.lube_type || "-"), 16),
              compactCell(String(r.issue_part || "-") || "-", 18),
              String(Number(r.entries || 0)),
              Number(r.qty_total || 0).toFixed(1),
              Number(r.lube_cost || 0).toFixed(2),
            ];
          })
          : [["-", "No lube issues in period", "-", "-", "-", "-", "-"]]),
      ],
      { x: 0.45, y: 1.05, w: 12.3, h: lubeTableH, fontSize: 9.5, ...tableOptions }
    );
    const lubeShownNote = lubePeriodTotals.groups > lubeByMachine.length
      ? ` | Showing top ${lubeByMachine.length} of ${lubePeriodTotals.groups} asset/lube lines by cost`
      : "";
    s7.addText(
      `Period total usage: ${lubePeriodTotals.qty.toFixed(1)} qty | Lube cost: $${fmtNum(lubePeriodTotals.cost, 2)} | Log lines: ${lubePeriodTotals.entries}${lubeShownNote}`,
      { x: 0.5, y: Math.min(5.95, 1.05 + lubeTableH + 0.2), w: 12.0, h: 0.35, fontFace: deckBodyFont, fontSize: 12, bold: true, color: deckText },
    );
    const s8 = pptx.addSlide();
    addExecutiveFrame(s8);
    s8.addText("7) Manager Inspections", { x: 0.4, y: 0.25, w: 12.4, h: 0.45, fontFace: deckHeaderFont, fontSize: 22, bold: true, color: deckNavy });
    const inspectionDisplayRows = inspectionsFaultRows.map((r) => {
      let faultNotes = [];
      try {
        const parsed = JSON.parse(String(r.checklist_json || "[]"));
        if (Array.isArray(parsed)) {
          faultNotes = parsed
            .filter((x) => x?.ok === false)
            .map((x) => String(x?.note || x?.label || x?.key || "").trim())
            .filter(Boolean);
        }
      } catch { /* ignore */ }
      const desc = [
        String(r.inspection_notes || "").trim(),
        String(r.recommended_action || "").trim(),
        faultNotes.length ? `Faults: ${faultNotes.slice(0, 3).join("; ")}` : "",
      ].filter(Boolean).join(" | ") || "Inspection completed";
      return { ...r, description: desc };
    });
    s8.addText(`Inspections completed: ${Number(inspectionsSummary?.inspections_done || 0)} | Assets covered: ${Number(inspectionsSummary?.assets_covered || 0)}`, { x: 0.7, y: 0.75, w: 12.0, h: 0.35, fontFace: deckBodyFont, fontSize: 12, bold: true, color: deckText });
    let inspY = 1.2;
    for (const r of inspectionDisplayRows.slice(0, 4)) {
      s8.addText(
        `${String(r.inspection_date || "-")} — ${String(r.asset_code || "-")} ${compactCell(String(r.asset_name || ""), 22)}`,
        { x: 0.55, y: inspY, w: 7.8, h: 0.28, fontFace: deckBodyFont, fontSize: 10, bold: true, color: deckNavy },
      );
      s8.addText(compactCell(String(r.description || "-"), 120), { x: 0.55, y: inspY + 0.28, w: 7.8, h: 0.55, fontFace: deckBodyFont, fontSize: 9, color: deckText });
      const photoAbs = r.photo_path
        ? resolveStorageAbs(String(r.photo_path || "").replace(/\\/g, "/").replace(/^\/+/, ""))
        : null;
      if (photoAbs && fs.existsSync(photoAbs)) {
        s8.addImage({ path: photoAbs, x: 8.55, y: inspY, w: 3.6, h: 0.85 });
      }
      inspY += 1.05;
    }
    if (!inspectionDisplayRows.length) {
      s8.addText("No manager inspections recorded for the selected period.", { x: 0.7, y: 1.5, w: 11.5, h: 0.4, fontFace: deckBodyFont, fontSize: 11, color: deckMuted });
    }
    const s9 = pptx.addSlide();
    addExecutiveFrame(s9);
    s9.addText("8) Fuel Anomalies", { x: 0.4, y: 0.25, w: 12.4, h: 0.45, fontFace: deckHeaderFont, fontSize: 22, bold: true, color: deckNavy });
    s9.addText(`Total fuel anomaly days: ${Number(totalFuelAnomalyDays || 0)}`, { x: 0.75, y: 0.75, w: 6.2, h: 0.35, fontFace: deckBodyFont, fontSize: 13, bold: true, color: deckText });
    const anomalyTableH = tableHeightFor(1 + (fuelAnomalyExpanded.length || 1), 1.2, 5.0);
    s9.addTable(
      [
        tableHeader("Asset", "Anomaly days", "Variance (LPH)", "Variance %", "Total usage (L)"),
        ...(fuelAnomalyExpanded.length
          ? fuelAnomalyExpanded.map((r) => [
            String(r.asset_code || ""),
            String(Number(r.anomaly_days || 0)),
            Number(r.peak_variance_lph || 0).toFixed(2),
            `${Number(r.peak_variance_pct || 0).toFixed(1)}%`,
            Number(r.total_usage_liters || 0).toFixed(1),
          ])
          : [["-", "0", "-", "-", "-"]]),
      ],
      { x: 0.75, y: 1.2, w: 11.4, h: anomalyTableH, fontSize: 11, ...tableOptions }
    );
    const buffer = await pptx.write({ outputType: "nodebuffer" });
    return Buffer.from(buffer);
  }

  async function buildWeeklyForumPresentation({ period, label, site_code, requestHeaders = {} }) {
    const msPerDay = 24 * 60 * 60 * 1000;
    const shiftDate = (ymd, days) => {
      const d = new Date(`${ymd}T12:00:00Z`);
      d.setUTCDate(d.getUTCDate() + days);
      return d.toISOString().slice(0, 10);
    };
    const rangeDays = Math.max(1, Math.round((new Date(`${period.end}T12:00:00Z`) - new Date(`${period.start}T12:00:00Z`)) / msPerDay) + 1);
    const previousPeriod = {
      start: shiftDate(period.start, -rangeDays),
      end: shiftDate(period.end, -rangeDays),
    };
    const dateLabel = (start, end) => {
      const options = { day: "2-digit", month: "short", year: "numeric", timeZone: "UTC" };
      const a = new Intl.DateTimeFormat("en-GB", options).format(new Date(`${start}T12:00:00Z`));
      const b = new Intl.DateTimeFormat("en-GB", options).format(new Date(`${end}T12:00:00Z`));
      return start === end ? a : `${a} – ${b}`;
    };
    const compact = (value, max = 64) => {
      const text = String(value || "").replace(/\s+/g, " ").trim();
      return text.length > max ? `${text.slice(0, Math.max(1, max - 1)).trim()}…` : text;
    };
    const num = (value, places = 1) => Number(Number(value || 0).toFixed(places));
    const money = (value) => `$${Number(value || 0).toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
    const pct = (value) => value == null || !Number.isFinite(Number(value)) ? "—" : `${Number(value).toFixed(1)}%`;
    const hasDowntimeLogs = hasTable("breakdown_downtime_logs") && hasTable("breakdowns");
    const maintenanceScope = andDailyHoursFleetHoursOnly("dh", "a");
    const maintenanceFleetScope = andAssetFleetHoursOnly("a");
    const withoutLdvs = andAssetExcludeLdv("a");

    const rangeKpis = (range) => {
      const hours = db.prepare(`
        SELECT
          COALESCE(SUM(dh.scheduled_hours), 0) AS scheduled_hours,
          COALESCE(SUM(dh.hours_run), 0) AS run_hours
        FROM daily_hours dh
        JOIN assets a ON a.id = dh.asset_id
        WHERE dh.work_date BETWEEN ? AND ?
          ${maintenanceScope}
          ${withoutLdvs}
      `).get(range.start, range.end);
      const downtime = hasDowntimeLogs
        ? db.prepare(`
            SELECT
              COALESCE(SUM(l.hours_down), 0) AS downtime_hours,
              COUNT(DISTINCT b.asset_id) AS assets_down
            FROM breakdown_downtime_logs l
            JOIN breakdowns b ON b.id = l.breakdown_id
            JOIN assets a ON a.id = b.asset_id
            WHERE l.log_date BETWEEN ? AND ?
              ${maintenanceFleetScope}
              ${withoutLdvs}
          `).get(range.start, range.end)
        : { downtime_hours: 0, assets_down: 0 };
      const criticalOpen = db.prepare(`
        SELECT COUNT(*) AS count
        FROM work_orders w
        JOIN breakdowns b ON b.id = w.reference_id
        WHERE LOWER(COALESCE(w.source, '')) = 'breakdown'
          AND COALESCE(b.critical, 0) = 1
          AND LOWER(COALESCE(w.status, 'open')) NOT IN (${doneStatusSql})
      `).get();
      const scheduled = Number(hours?.scheduled_hours || 0);
      const run = Math.max(0, Math.min(Number(hours?.run_hours || 0), scheduled));
      const down = Math.max(0, Math.min(Number(downtime?.downtime_hours || 0), scheduled));
      return {
        scheduled,
        run,
        downtime: down,
        assetsDown: Number(downtime?.assets_down || 0),
        criticalOpen: Number(criticalOpen?.count || 0),
        availability: scheduled > 0 ? ((scheduled - down) / scheduled) * 100 : null,
        utilization: scheduled > 0 ? (run / scheduled) * 100 : null,
      };
    };

    const selectedKpis = rangeKpis(period);
    const previousKpis = rangeKpis(previousPeriod);
    const dailyKpiRows = db.prepare(`
      WITH hours AS (
        SELECT dh.work_date AS day_key,
          COALESCE(SUM(dh.scheduled_hours), 0) AS scheduled_hours,
          COALESCE(SUM(dh.hours_run), 0) AS run_hours
        FROM daily_hours dh
        JOIN assets a ON a.id = dh.asset_id
        WHERE dh.work_date BETWEEN ? AND ?
          ${maintenanceScope}
          ${withoutLdvs}
        GROUP BY dh.work_date
      ), downtime AS (
        ${hasDowntimeLogs
          ? `SELECT l.log_date AS day_key, COALESCE(SUM(l.hours_down), 0) AS downtime_hours
             FROM breakdown_downtime_logs l
             JOIN breakdowns b ON b.id = l.breakdown_id
             JOIN assets a ON a.id = b.asset_id
             WHERE l.log_date BETWEEN ? AND ?
               ${maintenanceFleetScope}
               ${withoutLdvs}
             GROUP BY l.log_date`
          : `SELECT NULL AS day_key, 0 AS downtime_hours WHERE 1 = 0`}
      )
      SELECT hours.day_key, hours.scheduled_hours, hours.run_hours,
        COALESCE(downtime.downtime_hours, 0) AS downtime_hours
      FROM hours
      LEFT JOIN downtime ON downtime.day_key = hours.day_key
      ORDER BY hours.day_key ASC
    `).all(...(hasDowntimeLogs
      ? [period.start, period.end, period.start, period.end]
      : [period.start, period.end]));

    const assetPerformance = db.prepare(`
      WITH hours AS (
        SELECT a.id AS asset_id, a.asset_code, a.asset_name,
          COALESCE(a.category, 'Uncategorised') AS category,
          COALESCE(SUM(dh.scheduled_hours), 0) AS scheduled_hours,
          COALESCE(SUM(dh.hours_run), 0) AS run_hours
        FROM assets a
        LEFT JOIN daily_hours dh ON dh.asset_id = a.id
          AND dh.work_date BETWEEN ? AND ?
          AND dh.is_used = 1
        WHERE a.active = 1 AND a.is_standby = 0
          ${withoutLdvs}
        GROUP BY a.id
      ), downtime AS (
        ${hasDowntimeLogs
          ? `SELECT b.asset_id, COALESCE(SUM(l.hours_down), 0) AS downtime_hours
             FROM breakdown_downtime_logs l
             JOIN breakdowns b ON b.id = l.breakdown_id
             WHERE l.log_date BETWEEN ? AND ?
             GROUP BY b.asset_id`
          : `SELECT NULL AS asset_id, 0 AS downtime_hours WHERE 1 = 0`}
      )
      SELECT hours.*, COALESCE(downtime.downtime_hours, 0) AS downtime_hours
      FROM hours LEFT JOIN downtime ON downtime.asset_id = hours.asset_id
      WHERE hours.scheduled_hours > 0
    `).all(...(hasDowntimeLogs
      ? [period.start, period.end, period.start, period.end]
      : [period.start, period.end])).map((row) => {
      const scheduled = Number(row.scheduled_hours || 0);
      const run = Math.max(0, Math.min(Number(row.run_hours || 0), scheduled));
      const downtime = Math.max(0, Math.min(Number(row.downtime_hours || 0), scheduled));
      return {
        ...row,
        scheduled_hours: scheduled,
        run_hours: run,
        downtime_hours: downtime,
        availability: scheduled > 0 ? ((scheduled - downtime) / scheduled) * 100 : 0,
        utilization: scheduled > 0 ? (run / scheduled) * 100 : 0,
      };
    });
    // Equipment KPIs (rankings, low use) cover production equipment only.
    const productionPerformance = assetPerformance.filter((row) => isProductionAsset(row));
    const categoryRows = [...productionPerformance.reduce((map, row) => {
      const key = String(row.category || "Uncategorised");
      const current = map.get(key) || { category: key, scheduled: 0, run: 0, downtime: 0, count: 0 };
      current.scheduled += Number(row.scheduled_hours || 0);
      current.run += Number(row.run_hours || 0);
      current.downtime += Number(row.downtime_hours || 0);
      current.count += 1;
      map.set(key, current);
      return map;
    }, new Map()).values()].map((row) => ({
      ...row,
      availability: row.scheduled > 0 ? ((row.scheduled - Math.min(row.downtime, row.scheduled)) / row.scheduled) * 100 : 0,
      utilization: row.scheduled > 0 ? (Math.min(row.run, row.scheduled) / row.scheduled) * 100 : 0,
    })).sort((a, b) => b.utilization - a.utilization);
    const rankedAssets = [...productionPerformance].sort((a, b) => b.utilization - a.utilization);
    const lowUseAssets = [...productionPerformance]
      .filter((row) => Number(row.scheduled_hours || 0) > 0)
      .sort((a, b) => a.utilization - b.utilization)
      .slice(0, 8);

    const breakdownRows = hasDowntimeLogs
      ? db.prepare(`
          SELECT b.id, a.asset_code, a.asset_name, b.status, b.component, b.description,
            b.parts_status, b.ets_repair_date, b.end_at,
            MIN(l.log_date) AS start_date, MAX(l.log_date) AS end_date,
            COUNT(DISTINCT l.log_date) AS day_count,
            COALESCE(SUM(l.hours_down), 0) AS downtime_hours
          FROM breakdown_downtime_logs l
          JOIN breakdowns b ON b.id = l.breakdown_id
          JOIN assets a ON a.id = b.asset_id
          WHERE l.log_date BETWEEN ? AND ?
          GROUP BY b.id
          ORDER BY downtime_hours DESC, start_date ASC, a.asset_code ASC
          LIMIT 10
        `).all(period.start, period.end)
      : [];
    const workOrderEndExpr = hasColumn("work_orders", "completed_at")
      ? "COALESCE(w.completed_at, w.closed_at, w.opened_at)"
      : "COALESCE(w.closed_at, w.opened_at)";
    const woCol = (col) => (hasColumn("work_orders", col) ? `w.${col}` : `NULL`);
    const hasPartsRequests = hasTable("maintenance_parts_requests");
    const workOrderRows = db.prepare(`
      SELECT w.id, w.source, w.reference_id, w.status, w.opened_at, w.closed_at,
        ${woCol("completed_at")} AS completed_at,
        ${woCol("assigned_artisan_name")} AS assigned_artisan_name,
        ${woCol("due_date")} AS due_date,
        ${woCol("repair_progress")} AS repair_progress,
        a.asset_code, a.asset_name, a.category,
        b.description AS breakdown_description, b.component AS breakdown_component,
        b.parts_status, b.ets_repair_date,
        mp.service_name,
        ${hasPartsRequests ? `(
          SELECT COUNT(*) FROM maintenance_parts_requests pr
          WHERE pr.work_order_id = w.id
            AND LOWER(COALESCE(pr.status, 'requested')) NOT IN ('received', 'issued', 'closed', 'cancelled', 'rejected')
        )` : "0"} AS open_parts_requests,
        ${hasPartsRequests ? `(
          SELECT COALESCE(pr.part_name, pr.part_code) FROM maintenance_parts_requests pr
          WHERE pr.work_order_id = w.id
            AND LOWER(COALESCE(pr.status, 'requested')) NOT IN ('received', 'issued', 'closed', 'cancelled', 'rejected')
          ORDER BY pr.id LIMIT 1
        )` : "NULL"} AS first_waiting_part
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      LEFT JOIN breakdowns b ON LOWER(COALESCE(w.source, '')) = 'breakdown' AND b.id = w.reference_id
      LEFT JOIN maintenance_plans mp ON LOWER(COALESCE(w.source, '')) = 'service' AND mp.id = w.reference_id
      WHERE DATE(w.opened_at) <= ?
        AND (
          LOWER(COALESCE(w.status, 'open')) NOT IN (${doneStatusSql})
          OR DATE(${workOrderEndExpr}) BETWEEN ? AND ?
        )
        AND LOWER(COALESCE(w.source, '')) IN ('breakdown', 'service')
      ORDER BY
        CASE WHEN LOWER(COALESCE(w.status, 'open')) IN (${doneStatusSql}) THEN 1 ELSE 0 END,
        DATE(w.opened_at) DESC, w.id DESC
      LIMIT 60
    `).all(period.end, period.start, period.end);
    // The repair status slide lists production equipment only.
    const productionWorkOrders = workOrderRows.filter((r) => isProductionAsset(r)).slice(0, 12);

    const defaults = costDefaults();
    const laborRate = Number(defaults.labor_cost_per_hour_default || 0);
    const rangeCosts = (range) => {
      const partsIssued = hasTable("stock_movements") && hasTable("parts") && hasColumn("parts", "unit_cost")
        ? Number(db.prepare(`
            SELECT COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS value
            FROM stock_movements sm
            JOIN parts p ON p.id = sm.part_id
            WHERE sm.movement_type = 'out'
              AND ${issuedToEquipmentSql("sm")}
              AND DATE(sm.created_at) BETWEEN ? AND ?
          `).get(range.start, range.end)?.value || 0)
        : 0;
      const internalLabor = hasColumn("work_orders", "labor_hours") && hasColumn("work_orders", "labor_rate_per_hour")
        ? Number(db.prepare(`
            SELECT COALESCE(SUM(${hasColumn("work_orders", "planned_labor_hours") ? "COALESCE(NULLIF(w.labor_hours, 0), w.planned_labor_hours, 0)" : "COALESCE(w.labor_hours, 0)"} * COALESCE(NULLIF(w.labor_rate_per_hour, 0), ?)), 0) AS value
            FROM work_orders w
            WHERE DATE(${workOrderEndExpr}) BETWEEN ? AND ?
          `).get(laborRate, range.start, range.end)?.value || 0)
        : 0;
      return { partsIssued, internalLabor, externalRepairs: 0, total: partsIssued + internalLabor };
    };
    const selectedCosts = rangeCosts(period);
    const previousCosts = rangeCosts(previousPeriod);
    const costDifference = (current, previous) => ({
      value: Number(current || 0) - Number(previous || 0),
      percent: Number(previous || 0) > 0 ? ((Number(current || 0) - Number(previous || 0)) / Number(previous || 0)) * 100 : null,
    });

    let upcomingServices = [];
    try {
      const injected = await app.inject({
        method: "GET",
        url: `/api/maintenance/insights?start=${encodeURIComponent(period.start)}&end=${encodeURIComponent(period.end)}&near_due_hours=50&predictive_horizon_hours=100`,
        headers: maintenanceDeckInsightsHeaders(site_code, requestHeaders),
      });
      if (injected.statusCode < 400) {
        const parsed = JSON.parse(String(injected.payload || "{}"));
        upcomingServices = Array.isArray(parsed?.parts_planning?.upcoming_cost_forecasts)
          ? parsed.parts_planning.upcoming_cost_forecasts
          : [];
      }
    } catch { /* A deck can still be produced without forecast data. */ }
    const reviewNotes = hasTable("weekly_forum_review_notes")
      ? db.prepare(`
          SELECT area, weekly_finding, action_owner, due_date
          FROM weekly_forum_review_notes
          WHERE period_start = ? AND period_end = ?
          ORDER BY CASE area WHEN 'Downtime' THEN 0 WHEN 'Repairs' THEN 1 WHEN 'Costs' THEN 2 ELSE 3 END, id ASC
        `).all(period.start, period.end)
      : [];
    const actionRows = hasTable("weekly_forum_actions")
      ? db.prepare(`
          SELECT department, action_item, owner_name, due_date, status, notes
          FROM weekly_forum_actions
          WHERE action_date <= ?
            AND COALESCE(due_date, action_date) >= ?
            AND LOWER(COALESCE(status, 'open')) NOT IN ('done', 'closed', 'completed')
          ORDER BY COALESCE(due_date, action_date) ASC, id DESC
          LIMIT 5
        `).all(period.end, period.start)
      : [];

    // Parts ordered: order lines placed in the range (cancelled lines left out), in USD.
    const partsOrdered = (range) => {
      if (!hasTable("stores_part_orders")) return 0;
      const rate = (key, fallback) => {
        const v = hasTable("cost_settings") ? Number(db.prepare(`SELECT value FROM cost_settings WHERE key = ?`).get(key)?.value) : NaN;
        return Number.isFinite(v) && v > 0 ? v : fallback;
      };
      return Number(db.prepare(`
        SELECT COALESCE(SUM(COALESCE(o.qty, 0) * COALESCE(o.unit_cost, 0) /
          CASE UPPER(COALESCE(o.currency, 'USD')) WHEN 'ZAR' THEN ? WHEN 'MZN' THEN ? ELSE 1 END), 0) AS value
        FROM stores_part_orders o
        WHERE DATE(o.order_date) BETWEEN ? AND ?
          AND LOWER(COALESCE(o.status, '')) <> 'cancelled'
          AND LOWER(COALESCE(o.site_code, 'main')) = LOWER(?)
      `).get(rate("zar_per_usd", 18.5), rate("mzn_per_usd", 64), range.start, range.end, String(site_code || "main"))?.value || 0);
    };
    // Month costing: the last full month and the current month to the end of the selected week.
    const endMonth = String(period.end).slice(0, 7);
    const monthStart = (ym) => `${ym}-01`;
    const monthEnd = (ym) => {
      const [y, m] = ym.split("-").map(Number);
      return new Date(Date.UTC(y, m, 0)).toISOString().slice(0, 10);
    };
    const prevMonth = (() => {
      const [y, m] = endMonth.split("-").map(Number);
      const d = new Date(Date.UTC(y, m - 2, 1));
      return d.toISOString().slice(0, 7);
    })();
    const monthName = (ym) => new Intl.DateTimeFormat("en-GB", { month: "long", year: "numeric", timeZone: "UTC" }).format(new Date(`${ym}-15T12:00:00Z`));
    const monthCosting = [
      { label: monthName(prevMonth), range: { start: monthStart(prevMonth), end: monthEnd(prevMonth) } },
      { label: `${monthName(endMonth).split(" ")[0]} to date`, range: { start: monthStart(endMonth), end: period.end } },
    ].map((m) => {
      const c = rangeCosts(m.range);
      return { ...m, partsIssued: c.partsIssued, labour: c.internalLabor, total: c.partsIssued + c.internalLabor, partsOrdered: partsOrdered(m.range) };
    });

    // Equipment away from site for repairs: off-site repair records not yet returned
    // (or returned inside the selected week), plus open breakdowns noted as off site.
    const offsiteRows = [];
    if (hasTable("breakdown_offsite_repairs")) {
      offsiteRows.push(...db.prepare(`
        SELECT o.breakdown_id, a.asset_code, a.asset_name, o.repair_status, o.sent_date, o.expected_return_date, o.actual_return_date,
          COALESCE(NULLIF(TRIM(o.current_location), ''), NULLIF(TRIM(o.vendor), '')) AS place,
          o.vendor, COALESCE(NULLIF(TRIM(o.repair_reason), ''), NULLIF(TRIM(o.notes), ''), b.description, b.component) AS reason
        FROM breakdown_offsite_repairs o
        JOIN assets a ON a.id = o.asset_id
        LEFT JOIN breakdowns b ON b.id = o.breakdown_id
        WHERE LOWER(COALESCE(o.site_code, 'main')) = LOWER(?)
          AND DATE(o.sent_date) <= ?
          AND (LOWER(COALESCE(o.repair_status, '')) <> 'returned' OR DATE(COALESCE(o.actual_return_date, o.updated_at)) BETWEEN ? AND ?)
        ORDER BY CASE WHEN LOWER(COALESCE(o.repair_status, '')) = 'returned' THEN 1 ELSE 0 END, o.sent_date ASC
      `).all(String(site_code || "main"), period.end, period.start, period.end));
    }
    if (hasTable("breakdowns")) {
      const noted = db.prepare(`
        SELECT b.id AS breakdown_id, a.asset_code, a.asset_name, b.status AS repair_status, b.breakdown_date AS sent_date,
          b.ets_repair_date AS expected_return_date, NULL AS actual_return_date, NULL AS place, NULL AS vendor,
          TRIM(COALESCE(b.component, '') || ' ' || COALESCE(b.description, '')) AS reason,
          COALESCE(b.parts_status, '') AS parts_status
        FROM breakdowns b
        JOIN assets a ON a.id = b.asset_id
        WHERE UPPER(TRIM(COALESCE(b.status, 'OPEN'))) <> 'CLOSED'
      `).all().filter((r) => /off.?site|south africa|\bRSA\b|external repair|sent (away|out) for repair/i.test(`${r.reason} ${r.parts_status}`));
      for (const r of noted) {
        if (!offsiteRows.some((o) => (o.breakdown_id && o.breakdown_id === r.breakdown_id) || o.asset_code === r.asset_code)) offsiteRows.push({ ...r, repair_status: "off site (breakdown notes)" });
      }
    }

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "IRONLOG";
    pptx.subject = "Weekly maintenance forum";
    pptx.title = `Weekly Maintenance Forum - ${label}`;
    pptx.company = "AML";
    pptx.lang = "en-ZA";
    // Black and CAT yellow. Yellow is for bars, rules and header text on black;
    // body text stays dark for readability on the projector.
    const navy = "1A1A1A";
    const teal = "FFCD11";
    const blue = "3A3A3A";
    const accentText = "3A3A3A";
    const pale = "F5F4EE";
    const stripe = "FAF9F4";
    const line = "D9D6C8";
    const text = "1A1A1A";
    const muted = "5C5C5C";
    const headingFont = "Aptos Display";
    const bodyFont = "Aptos";
    const headerPeriod = `${dateLabel(period.start, period.end).toUpperCase()}  •  CURRENT WEEK`;
    const tableHeader = (...labels) => labels.map((value) => ({
      text: value,
      options: { bold: true, color: teal, fill: { color: navy }, fontFace: bodyFont },
    }));
    const TOTAL_SLIDES = 11;
    const tableOptions = { border: { pt: 0.5, color: line }, color: text, fontFace: bodyFont, margin: 0.06 };
    const addChrome = (slide, title, page) => {
      slide.background = { color: "FFFFFF" };
      slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.74, line: { color: navy, transparency: 100 }, fill: { color: navy } });
      slide.addText("AML  /  WEEKLY MAINTENANCE", { x: 0.34, y: 0.17, w: 4.9, h: 0.22, fontFace: headingFont, fontSize: 12, bold: true, color: "FFFFFF" });
      slide.addText(headerPeriod, { x: 8.1, y: 0.18, w: 4.86, h: 0.18, fontFace: bodyFont, fontSize: 8.8, color: teal, align: "right" });
      slide.addText(title, { x: 0.38, y: 0.96, w: 9.95, h: 0.4, fontFace: headingFont, fontSize: 22, bold: true, color: navy });
      slide.addShape(pptx.ShapeType.line, { x: 0.38, y: 1.6, w: 12.55, h: 0, line: { color: teal, pt: 2 } });
      slide.addShape(pptx.ShapeType.line, { x: 0.38, y: 7.1, w: 12.55, h: 0, line: { color: line, pt: 0.5 } });
      slide.addText(`${page} / ${TOTAL_SLIDES}`, { x: 12.2, y: 7.16, w: 0.7, h: 0.15, fontFace: bodyFont, fontSize: 7.3, bold: true, color: muted, align: "right" });
    };
    const addKpiTile = (slide, x, labelText, valueText, note) => {
      slide.addShape(pptx.ShapeType.rect, { x, y: 1.95, w: 2.95, h: 1.2, line: { color: line, pt: 0.6 }, fill: { color: pale } });
      slide.addShape(pptx.ShapeType.rect, { x, y: 1.95, w: 0.07, h: 1.2, line: { color: teal, transparency: 100 }, fill: { color: teal } });
      slide.addText(labelText.toUpperCase(), { x: x + 0.16, y: 2.14, w: 2.6, h: 0.14, fontFace: bodyFont, fontSize: 7.4, bold: true, color: muted });
      slide.addText(valueText, { x: x + 0.16, y: 2.42, w: 2.6, h: 0.34, fontFace: headingFont, fontSize: 21, bold: true, color: navy });
      slide.addText(note, { x: x + 0.16, y: 2.89, w: 2.6, h: 0.12, fontFace: bodyFont, fontSize: 6.8, color: muted });
    };

    const s1 = pptx.addSlide();
    s1.background = { color: "FFFFFF" };
    s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.74, line: { color: navy, transparency: 100 }, fill: { color: navy } });
    s1.addText("AML  /  WEEKLY MAINTENANCE", { x: 0.34, y: 0.17, w: 5.2, h: 0.22, fontFace: headingFont, fontSize: 12, bold: true, color: "FFFFFF" });
    s1.addText(headerPeriod, { x: 8.1, y: 0.18, w: 4.86, h: 0.18, fontFace: bodyFont, fontSize: 8.8, color: teal, align: "right" });
    s1.addText("Mechanical performance", { x: 0.38, y: 0.96, w: 8.4, h: 0.4, fontFace: headingFont, fontSize: 22, bold: true, color: navy });
    s1.addText("LIVE IRONLOG DATA  •  Weekly maintenance review", { x: 0.4, y: 1.4, w: 6.5, h: 0.16, fontFace: bodyFont, fontSize: 8, color: muted });
    s1.addShape(pptx.ShapeType.line, { x: 0.38, y: 1.68, w: 12.55, h: 0, line: { color: teal, pt: 2 } });
    s1.addText("Weekly maintenance review", { x: 0.5, y: 2.0, w: 6.5, h: 0.45, fontFace: headingFont, fontSize: 28, bold: true, color: navy });
    s1.addText(`Selected week: ${dateLabel(period.start, period.end)}`, { x: 0.52, y: 2.72, w: 5.4, h: 0.22, fontFace: bodyFont, fontSize: 12, color: blue });
    s1.addText(`Previous week: ${dateLabel(previousPeriod.start, previousPeriod.end)}`, { x: 0.52, y: 3.06, w: 5.4, h: 0.22, fontFace: bodyFont, fontSize: 11, color: muted });
    s1.addText("Mechanical downtime  •  Work orders  •  Off-site repairs  •  Cost movement  •  Upcoming maintenance  •  Month costing", { x: 0.52, y: 4.04, w: 11.6, h: 0.22, fontFace: bodyFont, fontSize: 10, bold: true, color: accentText });
    s1.addShape(pptx.ShapeType.rect, { x: 0.5, y: 4.65, w: 12.0, h: 1.1, line: { color: line, pt: 0.65 }, fill: { color: stripe } });
    s1.addShape(pptx.ShapeType.rect, { x: 0.5, y: 4.65, w: 0.09, h: 1.1, line: { color: teal, transparency: 100 }, fill: { color: teal } });
    s1.addText(`Availability ${pct(selectedKpis.availability)}   |   Utilization ${pct(selectedKpis.utilization)}   |   ${num(selectedKpis.downtime)} mechanical downtime hours   |   ${selectedKpis.assetsDown} assets down`, { x: 0.78, y: 5.01, w: 11.4, h: 0.28, fontFace: headingFont, fontSize: 16, bold: true, color: navy, align: "center" });
    s1.addShape(pptx.ShapeType.line, { x: 0.38, y: 7.1, w: 12.55, h: 0, line: { color: line, pt: 0.5 } });
    s1.addText(`1 / ${TOTAL_SLIDES}`, { x: 12.2, y: 7.16, w: 0.7, h: 0.15, fontFace: bodyFont, fontSize: 7.3, bold: true, color: muted, align: "right" });

    const s2 = pptx.addSlide();
    addChrome(s2, "Weekly maintenance at a glance", 2);
    addKpiTile(s2, 0.42, "Fleet availability", pct(selectedKpis.availability), `Previous: ${pct(previousKpis.availability)}`);
    addKpiTile(s2, 3.54, "Mechanical downtime", `${num(selectedKpis.downtime)} h`, `Previous: ${num(previousKpis.downtime)} h`);
    addKpiTile(s2, 6.66, "Assets down", String(selectedKpis.assetsDown), `Previous: ${previousKpis.assetsDown}`);
    addKpiTile(s2, 9.78, "Open critical WOs", String(selectedKpis.criticalOpen), "Current open backlog");
    s2.addText("What changed this week", { x: 0.42, y: 3.55, w: 5.2, h: 0.25, fontFace: headingFont, fontSize: 14, bold: true, color: navy });
    const reviewByArea = new Map(reviewNotes.map((row) => [String(row.area || ""), row]));
    const drafted = draftWeeklyFindings({ selectedKpis, previousKpis, breakdownRows, workOrderRows, selectedCosts, previousCosts, money });
    const reviewRows = ["Downtime", "Repairs", "Costs"].map((area) => {
      const row = reviewByArea.get(area);
      if (row) return [area, compact(row.weekly_finding, 190)];
      return [area, compact(drafted[area].finding, 190)];
    });
    s2.addTable([tableHeader("Area", "Weekly finding"), ...reviewRows], { x: 0.42, y: 3.92, w: 12.0, h: 2.12, colW: [1.6, 10.4], fontSize: 10, rowH: 0.52, ...tableOptions });

    const s3 = pptx.addSlide();
    addChrome(s3, "KPI hours and downtime", 3);
    s3.addText("Scheduled, run and downtime hours", { x: 0.48, y: 1.95, w: 6.2, h: 0.2, fontFace: headingFont, fontSize: 12, bold: true, color: navy });
    s3.addChart(pptx.ChartType.bar, [
      { name: "Scheduled", labels: dailyKpiRows.map((r) => String(r.day_key || "").slice(5)), values: dailyKpiRows.map((r) => num(r.scheduled_hours)) },
      { name: "Run", labels: dailyKpiRows.map((r) => String(r.day_key || "").slice(5)), values: dailyKpiRows.map((r) => num(r.run_hours)) },
      { name: "Downtime", labels: dailyKpiRows.map((r) => String(r.day_key || "").slice(5)), values: dailyKpiRows.map((r) => num(r.downtime_hours)) },
    ], { x: 0.42, y: 2.24, w: 7.15, h: 3.75, catAxisLabelRotate: -45, showLegend: true, legendPos: "b", chartColors: ["FFCD11", "3A3A3A", "C2410C"], showTitle: false });
    s3.addShape(pptx.ShapeType.rect, { x: 8.05, y: 2.2, w: 4.28, h: 3.45, line: { color: line, pt: 0.65 }, fill: { color: pale } });
    s3.addText("Weekly KPI rules", { x: 8.32, y: 2.5, w: 3.5, h: 0.2, fontFace: headingFont, fontSize: 13, bold: true, color: navy });
    s3.addText("Availability = available hours ÷ scheduled hours\n\nUtilization = run hours ÷ scheduled hours\n\nShow downtime inside the selected week only. Parked, standby and LDV units are excluded.", { x: 8.32, y: 3.0, w: 3.45, h: 1.7, fontFace: bodyFont, fontSize: 10, color: text, breakLine: false });
    s3.addText(`Week totals: ${num(selectedKpis.scheduled)} scheduled h  •  ${num(selectedKpis.run)} run h  •  ${num(selectedKpis.downtime)} downtime h`, { x: 0.6, y: 6.35, w: 11.6, h: 0.2, fontFace: bodyFont, fontSize: 10, bold: true, color: accentText, align: "center" });

    const s4 = pptx.addSlide();
    addChrome(s4, "Top and lowest production equipment by category", 4);
    s4.addText("Availability and utilization by category", { x: 0.48, y: 1.95, w: 6.2, h: 0.2, fontFace: headingFont, fontSize: 12, bold: true, color: navy });
    s4.addChart(pptx.ChartType.bar, [
      { name: "Availability %", labels: categoryRows.slice(0, 8).map((r) => compact(r.category, 16)), values: categoryRows.slice(0, 8).map((r) => num(r.availability)) },
      { name: "Utilization %", labels: categoryRows.slice(0, 8).map((r) => compact(r.category, 16)), values: categoryRows.slice(0, 8).map((r) => num(r.utilization)) },
    ], { x: 0.42, y: 2.24, w: 7.25, h: 3.95, catAxisLabelRotate: -40, showLegend: true, legendPos: "b", valAxisMinVal: 0, valAxisMaxVal: 100, chartColors: ["FFCD11", "3A3A3A"] });
    const topRows = rankedAssets.slice(0, 3).map((r) => [`Top • ${r.asset_code}`, compact(r.asset_name, 24), pct(r.utilization)]);
    const bottomRows = [...rankedAssets].filter((r) => r.run_hours > 0 || r.downtime_hours > 0).sort((a, b) => a.utilization - b.utilization).slice(0, 3).map((r) => [`Low • ${r.asset_code}`, compact(r.asset_name, 24), pct(r.utilization)]);
    s4.addText("Top / lowest asset watch", { x: 8.05, y: 1.95, w: 3.9, h: 0.2, fontFace: headingFont, fontSize: 12, bold: true, color: navy });
    s4.addTable([tableHeader("Rank", "Asset", "Util %"), ...topRows, ...bottomRows], { x: 8.02, y: 2.24, w: 4.3, h: 3.95, fontSize: 8.6, rowH: 0.43, ...tableOptions });

    const s5 = pptx.addSlide();
    addChrome(s5, "Production equipment with low use", 5);
    const lowUseRows = lowUseAssets.length
      ? lowUseAssets.map((r) => [r.asset_code, compact(r.asset_name, 34), compact(r.category, 22), `${num(r.run_hours)} / ${num(r.scheduled_hours)} h`, pct(r.utilization)])
      : [["—", "No production equipment with recorded scheduled hours", "—", "—", "—"]];
    s5.addText("Utilization, lowest first", { x: 0.48, y: 1.95, w: 6.2, h: 0.2, fontFace: headingFont, fontSize: 12, bold: true, color: navy });
    s5.addTable([tableHeader("Asset", "Equipment", "Category", "Run / scheduled", "Utilization"), ...lowUseRows], { x: 0.42, y: 2.24, w: 12.0, h: 3.85, colW: [1.6, 4.4, 2.8, 1.8, 1.4], fontSize: 9.4, rowH: 0.42, ...tableOptions });

    const s6 = pptx.addSlide();
    addChrome(s6, "Equipment down during selected week", 6);
    const downRows = breakdownRows.length ? breakdownRows.map((r) => [
      `${r.asset_code} • ${compact(r.asset_name, 22)}`,
      r.start_date === r.end_date ? String(r.start_date || "—") : `${r.start_date || "—"} → ${r.end_date || "—"}`,
      downtimeCell(r),
      compact([r.component, r.description].filter(Boolean).join(" • "), 60) || "—",
      String(r.status || "OPEN").toUpperCase(),
    ]) : [["—", "—", "0 h", "No mechanical downtime recorded", "—"]];
    s6.addTable([tableHeader("Asset", "Date / window", "Hours down", "Fault / component", "Status"), ...downRows], { x: 0.38, y: 2.05, w: 12.45, h: 4.15, colW: [2.9, 2.2, 1.9, 4.35, 1.1], fontSize: 8.8, rowH: 0.42, ...tableOptions });

    const s7 = pptx.addSlide();
    addChrome(s7, "Mechanical work and repair status", 7);
    const workRows = productionWorkOrders.length ? productionWorkOrders.map((r) => {
      const description = String(r.source || "").toLowerCase() === "service" ? r.service_name : r.breakdown_description;
      return [
        `WO-${r.id} • ${r.asset_code}`,
        compact(description || `${r.source || "Work order"} activity`, 56),
        compact(workOrderProgressCell(r), 38),
        compact(workOrderBlockerCell(r), 36),
      ];
    }) : [["—", "No breakdown or service work orders on production equipment", "—", "—"]];
    s7.addTable([tableHeader("Work order / asset", "Mechanical work", "Progress", "Blocker / parts"), ...workRows], { x: 0.38, y: 2.0, w: 12.45, h: 4.6, colW: [2.3, 4.45, 2.9, 2.8], fontSize: 8.8, rowH: 0.38, ...tableOptions });

    const sOff = pptx.addSlide();
    addChrome(sOff, "Equipment off site for repairs", 8);
    const today = period.end;
    const daysBetween = (a, b) => Math.max(0, Math.round((new Date(`${String(b).slice(0, 10)}T12:00:00Z`) - new Date(`${String(a).slice(0, 10)}T12:00:00Z`)) / msPerDay));
    const statusText = (s) => String(s || "").replace(/_/g, " ").replace(/^\w/, (c) => c.toUpperCase());
    const offRows = offsiteRows.length ? offsiteRows.slice(0, 10).map((r) => [
      `${r.asset_code} • ${compact(r.asset_name, 22)}`,
      compact(r.reason || "—", 52),
      compact(r.place || r.vendor || "—", 26),
      r.sent_date ? `${String(r.sent_date).slice(0, 10)} • ${daysBetween(r.sent_date, r.actual_return_date || today)} days` : "—",
      compact(statusText(r.repair_status), 22),
      r.actual_return_date ? `Returned ${String(r.actual_return_date).slice(0, 10)}` : (r.expected_return_date ? String(r.expected_return_date).slice(0, 10) : "Not set"),
    ]) : [["—", "No equipment is off site for repairs", "—", "—", "—", "—"]];
    sOff.addTable([tableHeader("Equipment", "Repair", "Where", "Sent / days away", "Status", "Expected back"), ...offRows], { x: 0.38, y: 2.0, w: 12.45, h: 4.4, colW: [2.6, 3.7, 1.95, 1.75, 1.2, 1.25], fontSize: 9, rowH: 0.42, ...tableOptions });

    const s8 = pptx.addSlide();
    addChrome(s8, "Maintenance cost against previous week", 9);
    const costs = [
      ["Parts issued", selectedCosts.partsIssued, previousCosts.partsIssued, "Parts and lubes issued to machines and work orders"],
      ["Internal labour", selectedCosts.internalLabor, previousCosts.internalLabor, `Labour hours on work orders (planned hours where none booked) × ${money(laborRate)}/h`],
      ["External repairs", selectedCosts.externalRepairs, previousCosts.externalRepairs, "Recorded external repair costs"],
      ["Total maintenance", selectedCosts.total, previousCosts.total, "Parts issued + internal labour"],
    ];
    const costRows = costs.map(([category, current, previous, note]) => {
      const change = costDifference(current, previous);
      return [category, money(previous), money(current), `${change.value >= 0 ? "+" : ""}${money(change.value)}${change.percent == null ? "" : ` (${change.percent >= 0 ? "+" : ""}${change.percent.toFixed(1)}%)`}`, note];
    });
    s8.addText(`Current: ${dateLabel(period.start, period.end)}     Previous: ${dateLabel(previousPeriod.start, previousPeriod.end)}`, { x: 0.48, y: 1.92, w: 11.2, h: 0.2, fontFace: bodyFont, fontSize: 10, bold: true, color: blue });
    s8.addTable([tableHeader("Cost category", "Previous week", "Selected week", "Change", "Reason / reference"), ...costRows], { x: 0.42, y: 2.28, w: 12.0, h: 2.75, fontSize: 9.3, rowH: 0.5, ...tableOptions });

    const s9 = pptx.addSlide();
    addChrome(s9, "Upcoming maintenance", 10);
    const plannedRows = [
      ...breakdownRows.filter((r) => String(r.status || "").toLowerCase() !== "closed").slice(0, 3).map((r, index) => {
        const wo = workOrderRows.find((w) => String(w.source || "").toLowerCase() === "breakdown" && Number(w.reference_id) === Number(r.id));
        return [
          String(index + 1), `${r.asset_code} • return to service`, r.ets_repair_date || "Set return target",
          compact(r.parts_status ? `Parts: ${r.parts_status}` : (r.description || "Repair plan required"), 34),
          wo ? compact(workOrderOwnerCell({ ...wo, ets_repair_date: r.ets_repair_date }), 34) : "Open a work order and assign an artisan",
        ];
      }),
      ...upcomingServices.slice(0, 3).map((r, index) => [
        String(index + 1 + breakdownRows.filter((x) => String(x.status || "").toLowerCase() !== "closed").slice(0, 3).length),
        `${r.asset_code || "—"} • ${compact(r.service_name || "Scheduled service", 28)}`,
        Number.isFinite(Number(r.remaining_hours))
          ? (Number(r.remaining_hours) < 0 ? `${num(Math.abs(Number(r.remaining_hours)), 0).toLocaleString("en-US")} h overdue` : `${num(r.remaining_hours, 0).toLocaleString("en-US")} h remaining`)
          : "Plan by availability",
        compact(Number(r?.forecast?.est_total_cost) > 0
          ? `$${Math.round(Number(r.forecast.est_total_cost)).toLocaleString("en-US")} • ${serviceCostSourceLabel(r)}`
          : serviceCostSourceLabel(r), 34),
        (() => {
          const wo = workOrderRows.find((w) => String(w.source || "").toLowerCase() === "service" && Number(w.reference_id) === Number(r.plan_id) && !isWorkOrderDone(w.status));
          return wo ? compact(workOrderOwnerCell(wo), 34) : "Raise work order; confirm outage window";
        })(),
      ]),
    ].slice(0, 5);
    const nextRows = plannedRows.length ? plannedRows : [["1", "No open repairs or near-due services", "—", "—", "—"]];
    s9.addTable([tableHeader("Priority", "Asset / job", "Planned date", "Dependency", "Owner / decision"), ...nextRows], { x: 0.38, y: 2.0, w: 12.45, h: 2.7, fontSize: 8.7, rowH: 0.42, ...tableOptions });
    s9.addText("Management decisions required", { x: 0.48, y: 5.05, w: 4.4, h: 0.2, fontFace: headingFont, fontSize: 12, bold: true, color: navy });
    const decisionText = actionRows.length
      ? actionRows.map((r) => `${compact(r.department, 14)}: ${compact(r.action_item, 64)} — ${compact(r.owner_name, 24)}${r.due_date ? ` (${r.due_date})` : ""}`).join("\n")
      : "No unresolved actions in the selected weekly range. Add actions in the Action Tracker so they carry into the next forum pack.";
    s9.addText(decisionText, { x: 0.48, y: 5.38, w: 11.7, h: 0.85, fontFace: bodyFont, fontSize: 9.3, color: muted, breakLine: false });

    // Month costing: last full month and the current month so far.
    const sMonth = pptx.addSlide();
    addChrome(sMonth, `${monthCosting[0].label} costing`, 11);
    const monthRows = [
      ["Parts issued", ...monthCosting.map((m) => money(m.partsIssued))],
      ["Labour", ...monthCosting.map((m) => money(m.labour))],
      [{ text: "Total maintenance cost", options: { bold: true, fill: { color: pale } } }, ...monthCosting.map((m) => ({ text: money(m.total), options: { bold: true, fill: { color: pale } } }))],
      ["Parts ordered", ...monthCosting.map((m) => money(m.partsOrdered))],
    ];
    sMonth.addTable([tableHeader("", ...monthCosting.map((m) => m.label)), ...monthRows], { x: 0.42, y: 2.0, w: 6.2, h: 2.6, colW: [2.6, 1.8, 1.8], fontSize: 12, rowH: 0.52, align: "left", ...tableOptions });
    sMonth.addChart(pptx.ChartType.bar, monthCosting.map((m) => ({
      name: m.label,
      labels: ["Parts issued", "Labour", "Parts ordered"],
      values: [Math.round(m.partsIssued), Math.round(m.labour), Math.round(m.partsOrdered)],
    })), { x: 6.95, y: 1.9, w: 6.0, h: 4.4, barGrouping: "clustered", showLegend: true, legendPos: "b", chartColors: ["1A1A1A", "FFCD11"], valAxisLabelFormatCode: "$#,##0", showValue: true, dataLabelFormatCode: "$#,##0", dataLabelFontSize: 8, catAxisLabelFontSize: 10 });
    sMonth.addText("Total maintenance cost = parts issued + labour. Parts ordered is shown on its own: ordered parts are charged when they are issued.", { x: 0.42, y: 4.85, w: 6.2, h: 0.5, fontFace: bodyFont, fontSize: 9.5, color: muted, breakLine: false });

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
    const deck = t === "weekly"
      ? await buildWeeklyForumPresentation({
          period,
          label,
          site_code,
          requestHeaders: opts.requestHeaders || {},
        })
      : await buildMaintenanceExecutiveDeck({
          period,
          label,
          site_code,
          requestHeaders: opts.requestHeaders || {},
        });
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

  // Route groups live in routes/reports/. They receive the shared helpers above through ctx.
  const ctx = {
    AML_WEEKLY_TEMPLATE_PATH,
    addFuelBenchmarkByAssetWorksheet,
    addTableSheet,
    allowedChannels,
    allowedFrequencies,
    allowedReportTypes,
    asArray,
    buildAmlWeeklyExportRecords,
    buildDailyExecutiveSummarySheet,
    buildGmWeeklyExecutiveSheet,
    buildMaintenanceCostByEquipment,
    buildMaintenanceExecutiveDeck,
    buildPeriodAssetCosts,
    buildPeriodContractorFuelRows,
    buildVehicleLdvChecklistPdfRows,
    calendarQuartersThrough,
    compactCell,
    completionNotesForPdf,
    costDefaults,
    dailyPdfDowntimeHours,
    dailyPdfLongDate,
    dailyPdfManagementSection,
    dailyPdfOperationsDate,
    dailyPdfRepairFallbackHours,
    dataRoot,
    datasetWithAvailableColumns,
    daysDownForBreakdown,
    daysDownForBreakdownInRange,
    deliverSubscription,
    downtimeMtdUsesDailyLogs,
    drawDailyPdfExceptions,
    drawDailyPdfMetricCards,
    drawDailyPdfOperatingBasis,
    fetchStoresPartOrdersForReport,
    fmtNum,
    generateMaintenanceMaster,
    getBreakdownDowntimeColumn,
    getDowntimeByAssetMtd,
    getDowntimeHoursForPeriod,
    getSiteCode,
    gmWeeklyPmComplianceSnapshot,
    gmWeeklyRepairForecast,
    hasColumn,
    hasTable,
    inclusiveDaysBetween,
    isMonth,
    isYmd,
    jobFindingsTextForPdf,
    kpiDaily,
    kpiRange,
    machinePrestartProfileFromCheckMode,
    makeArtisanFormNumber,
    mergeAssetCostsWithRunHours,
    monthRange,
    monthStartIso,
    monthsInRange,
    nextRunForSchedule,
    parseIsoDate,
    parsePartOrdersPeriod,
    parseRecipients,
    parseTimeHhMm,
    partOrderStatusLabel,
    pdfBrandingLogoDir,
    pickExistingColumn,
    prevMonth,
    queryPeriodFleetCostTotals,
    queryPeriodRunHoursByAsset,
    reliabilityMetricsForRange,
    reportDatasets,
    requireAdmin,
    resolveCheckPhotoForPdf,
    resolveMaintenancePeriod,
    resolveStorageAbs,
    rollupContractorFuelBySupplier,
    rollupFleetCostByCategory,
    runCustomBuilderQuery,
    safeNum,
    serviceLabelFromDailyDowntime,
    todayYmd,
    yn,
  };
  registerSettingsRoutes(app, ctx);
  registerWorkorderAssetRoutes(app, ctx);
  registerFuelLubeRoutes(app, ctx);
  registerInspectionsRoutes(app, ctx);
  registerStoresRoutes(app, ctx);
  registerOperationsExportsRoutes(app, ctx);
  registerPresentationsRoutes(app, ctx);
  registerPeriodReportsRoutes(app, ctx);
}
