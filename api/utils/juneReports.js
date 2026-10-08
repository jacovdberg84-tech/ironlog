// IRONLOG/api/utils/juneReports.js — reports and costing for June.
//
// June discusses costs and fleet performance with the administrator. These
// helpers reuse the numbers the reports already produce, so June never has
// her own version of the truth:
// - costs: buildAssetPeriodCosts (same source as Reports → Cost Monthly XLSX)
// - KPIs: the dashboard's Asset KPI range engine (availability, utilisation)
// - files: links to the existing report downloads, opened from a button.
import { db as defaultDb } from "../db/client.js";
import { getAssetKpiRangeBuilder } from "./assetKpiRangeProvider.js";
import { buildAssetPeriodCosts, monthPeriodBounds, readCostDefaults, rollupOperatingCategoryRows } from "./monthlyOperatingCosts.js";
import { buildPlantHireFinanceRows } from "./plantHire.js";

const CURRENCY = "USD";
const MAX_RANGE_DAYS = 370;

const isYmd = (v) => /^\d{4}-\d{2}-\d{2}$/.test(String(v || "").trim());
const isMonth = (v) => /^\d{4}-\d{2}$/.test(String(v || "").trim());
const round = (v, d = 2) => Number((Number(v) || 0).toFixed(d));
const ymd = (ms) => new Date(ms).toISOString().slice(0, 10);
const dayMs = 86_400_000;

function hasTable(dbConn, name) {
  return Boolean(dbConn.prepare(`SELECT 1 FROM sqlite_master WHERE type='table' AND name=?`).get(name));
}

/**
 * The period June was asked about, plus the one to compare against.
 * month → that month vs the month before; from/to → that range vs the same
 * number of days just before; nothing → this month so far vs last month.
 */
export function resolvePeriod({ month, from_date, to_date } = {}, today = new Date().toISOString().slice(0, 10)) {
  if (isYmd(from_date) && isYmd(to_date) && to_date >= from_date) {
    const days = Math.round((Date.parse(to_date) - Date.parse(from_date)) / dayMs) + 1;
    if (days > MAX_RANGE_DAYS) return { error: `Keep the range to ${MAX_RANGE_DAYS} days or less.` };
    const prevEnd = ymd(Date.parse(from_date) - dayMs);
    return {
      start: from_date, end: to_date, label: `${from_date} to ${to_date}`, month: null,
      previous: { start: ymd(Date.parse(prevEnd) - (days - 1) * dayMs), end: prevEnd, label: `previous ${days} days` },
    };
  }
  const m = isMonth(month) ? String(month) : today.slice(0, 7);
  const { start, end: monthEnd } = monthPeriodBounds(m);
  const toDate = !isMonth(month) && monthEnd > today ? today : monthEnd;
  const prevMonth = ymd(Date.parse(`${m}-01`) - dayMs).slice(0, 7);
  const prev = monthPeriodBounds(prevMonth);
  const monthToDate = toDate !== monthEnd;
  // Month to date is compared with the same days of last month, not a whole month.
  const sameDay = `${prevMonth}-${toDate.slice(8, 10)}`;
  const prevEnd = monthToDate ? (sameDay < prev.end ? sameDay : prev.end) : prev.end;
  return {
    start, end: toDate, month: m,
    label: monthToDate ? `${m} (month to date, up to ${toDate})` : m,
    month_to_date: monthToDate,
    previous: monthToDate
      ? { start: prev.start, end: prevEnd, label: `${prevMonth} same days (to ${prevEnd})`, month: null }
      : { start: prev.start, end: prev.end, label: prevMonth, month: prevMonth },
  };
}

function plantHireForMonth(dbConn, month) {
  if (!month) return 0;
  try {
    return buildPlantHireFinanceRows(dbConn, month).reduce((s, r) => s + Number(r.actual_amount || 0), 0);
  } catch {
    return 0;
  }
}

function filterAssets(rows, { asset_code, category }) {
  const code = String(asset_code || "").trim().toUpperCase();
  const cat = String(category || "").trim().toLowerCase();
  return rows.filter((r) =>
    (!code || String(r.asset_code || "").toUpperCase() === code)
    && (!cat || String(r.category || "").toLowerCase().includes(cat)));
}

function costRow(r) {
  return {
    asset_code: r.asset_code,
    asset_name: r.asset_name,
    category: r.category,
    total: round(r.total_cost),
    parts: round(r.parts_cost),
    labour: round(r.labor_cost),
    labour_hours: round(r.labor_hours, 1),
    fuel: round(r.fuel_cost),
    lube: round(r.lube_cost),
    downtime: round(r.downtime_cost),
    downtime_hours: round(r.downtime_hours, 1),
  };
}

function budgetByCategory(dbConn, month) {
  if (!month || !hasTable(dbConn, "finance_budgets_monthly")) return null;
  const rows = dbConn.prepare(`
    SELECT LOWER(TRIM(COALESCE(category, ''))) AS category, SUM(budget_amount) AS amount
    FROM finance_budgets_monthly WHERE period = ?
    GROUP BY LOWER(TRIM(COALESCE(category, '')))
  `).all(month);
  if (!rows.length) return null;
  return {
    total: round(rows.reduce((s, r) => s + Number(r.amount || 0), 0)),
    by_category: rows.map((r) => ({ category: r.category || "(no category)", amount: round(r.amount) })),
  };
}

function change(now, before) {
  const diff = round(now - before);
  return { now: round(now), before: round(before), change: diff, change_pct: before > 0 ? round((diff / before) * 100, 1) : null };
}

/** Run hours per asset from the Asset KPI engine, for cost per hour. */
function runHoursByAsset(range, siteCode, codes) {
  const build = getAssetKpiRangeBuilder();
  if (!build || !codes.length) return new Map();
  try {
    const kpi = build(range.start, range.end, 10, siteCode, codes);
    return new Map((kpi?.by_asset || []).map((a) => [String(a.asset_code || "").toUpperCase(), Number(a.run_hours || 0)]));
  } catch {
    return new Map();
  }
}

/** Costs for a period: categories, machines, budget and the previous period. */
export function buildJuneCostReport(args = {}, { dbConn = defaultDb, siteCode = "main", today } = {}) {
  const period = resolvePeriod(args, today);
  if (period.error) return { error: period.error };
  const defaults = readCostDefaults(dbConn);
  const allNow = buildAssetPeriodCosts(dbConn, period.start, period.end, defaults);
  const allBefore = buildAssetPeriodCosts(dbConn, period.previous.start, period.previous.end, defaults);
  const filtered = Boolean(args.asset_code || args.category);
  const now = filterAssets(allNow, args);
  const before = filterAssets(allBefore, args);
  // Plant hire is a monthly finance figure; it only applies to whole-fleet month views.
  const hireNow = filtered ? 0 : plantHireForMonth(dbConn, period.month);
  const hireBefore = filtered ? 0 : plantHireForMonth(dbConn, period.previous.month);
  const totalsNow = rollupOperatingCategoryRows(now, hireNow);
  const totalsBefore = rollupOperatingCategoryRows(before, hireBefore);
  const beforeByCat = new Map(totalsBefore.rows.map((r) => [r.category, r.amount]));
  const categories = Array.from(new Set([...totalsNow.rows.map((r) => r.category), ...beforeByCat.keys()]))
    .map((cat) => ({ category: cat === "labor" ? "labour" : cat, ...change(totalsNow.rows.find((r) => r.category === cat)?.amount || 0, beforeByCat.get(cat) || 0) }))
    .sort((a, b) => b.now - a.now);

  const byEquipmentType = new Map();
  for (const r of now) {
    const key = String(r.category || "Unassigned");
    byEquipmentType.set(key, (byEquipmentType.get(key) || 0) + Number(r.total_cost || 0));
  }
  const top = [...now].sort((a, b) => b.total_cost - a.total_cost).slice(0, args.asset_code ? 1 : 10);
  const runHours = runHoursByAsset(period, siteCode, top.map((r) => r.asset_code));
  const beforeByAsset = new Map(before.map((r) => [r.asset_code, r]));
  const machines = top.map((r) => {
    const run = runHours.get(String(r.asset_code).toUpperCase());
    const prev = beforeByAsset.get(r.asset_code);
    return {
      ...costRow(r),
      previous_total: prev ? round(prev.total_cost) : 0,
      run_hours: run != null ? round(run, 1) : null,
      cost_per_run_hour: run > 0 ? round(r.total_cost / run) : null,
    };
  });
  const risers = now
    .map((r) => ({ asset_code: r.asset_code, asset_name: r.asset_name, now: Number(r.total_cost || 0), before: Number(beforeByAsset.get(r.asset_code)?.total_cost || 0) }))
    .map((r) => ({ ...r, change: round(r.now - r.before), now: round(r.now), before: round(r.before) }))
    .filter((r) => r.change > 0)
    .sort((a, b) => b.change - a.change)
    .slice(0, 5);

  return {
    review_only: true,
    currency: CURRENCY,
    period: { start: period.start, end: period.end, label: period.label, month_to_date: Boolean(period.month_to_date) },
    compared_with: period.previous,
    filter: filtered ? { asset_code: args.asset_code || null, category: args.category || null } : null,
    total: change(totalsNow.total, totalsBefore.total),
    by_cost_type: categories,
    by_equipment_type: Array.from(byEquipmentType, ([category, amount]) => ({ category, amount: round(amount) })).sort((a, b) => b.amount - a.amount).slice(0, 12),
    top_machines: machines,
    biggest_increases: risers,
    budget: filtered ? null : budgetByCategory(dbConn, period.month),
    machines_with_costs: now.length,
    assumptions: {
      note: "Same source as Reports → Cost Monthly. Fuel, lube, labour and downtime use each machine's rate, else these defaults.",
      fuel_per_litre: defaults.fuel_cost_per_liter_default,
      lube_per_unit: defaults.lube_cost_per_qty_default,
      labour_per_hour: defaults.labor_cost_per_hour_default,
      downtime_per_hour: defaults.downtime_cost_per_hour_default,
      month_to_date_note: period.month_to_date ? "This month is not finished; it is compared with the same days of last month. Plant hire is left out of month-to-date views." : null,
      unlinked_note: now.some((r) => r.asset_code === "UNLINKED") ? "UNLINKED is spend that was not booked to a machine; worth chasing so machine costs are complete." : null,
    },
  };
}

function kpiSummary(kpi) {
  if (!kpi?.fleet) return null;
  return {
    availability_pct: kpi.fleet.availability_pct,
    utilisation_pct: kpi.fleet.utilization_pct,
    scheduled_hours: round(kpi.fleet.scheduled_hours, 1),
    run_hours: round(kpi.fleet.run_hours, 1),
    downtime_hours: round(kpi.fleet.downtime_hours, 1),
  };
}

function breakdownsIn(dbConn, start, end, codes) {
  if (!hasTable(dbConn, "breakdowns")) return null;
  const filter = codes?.length ? `AND UPPER(a.asset_code) IN (${codes.map(() => "?").join(",")})` : "";
  const params = [start, end, ...(codes || []).map((c) => String(c).toUpperCase())];
  const rows = dbConn.prepare(`
    SELECT a.asset_code, a.asset_name, COUNT(*) AS count, SUM(COALESCE(b.critical, 0)) AS critical
    FROM breakdowns b JOIN assets a ON a.id = b.asset_id
    WHERE b.breakdown_date BETWEEN ? AND ? ${filter}
    GROUP BY a.id ORDER BY count DESC, a.asset_code
  `).all(...params);
  return {
    total: rows.reduce((s, r) => s + Number(r.count || 0), 0),
    critical: rows.reduce((s, r) => s + Number(r.critical || 0), 0),
    repeat_machines: rows.filter((r) => Number(r.count) > 1).slice(0, 8).map((r) => ({ asset_code: r.asset_code, asset_name: r.asset_name, breakdowns: Number(r.count) })),
  };
}

/** Availability, utilisation, downtime and breakdowns for a period. */
export function buildJuneKpiReport(args = {}, { dbConn = defaultDb, siteCode = "main", today } = {}) {
  const period = resolvePeriod(args, today);
  if (period.error) return { error: period.error };
  const build = getAssetKpiRangeBuilder();
  if (!build) return { error: "The KPI engine is not available on this server." };
  const codes = Array.isArray(args.asset_codes) ? args.asset_codes.map((c) => String(c || "").trim().toUpperCase()).filter(Boolean).slice(0, 40) : [];
  const now = build(period.start, period.end, 10, siteCode, codes.length ? codes : undefined);
  const before = build(period.previous.start, period.previous.end, 10, siteCode, codes.length ? codes : undefined);
  const cat = String(args.category || "").trim().toLowerCase();
  const assets = (now?.by_asset || [])
    .filter((a) => !cat || String(a.category || "").toLowerCase().includes(cat))
    .filter((a) => Number(a.scheduled_hours) > 0);
  const slim = (a) => ({
    asset_code: a.asset_code,
    asset_name: a.asset_name,
    category: a.category,
    availability_pct: a.availability_pct,
    utilisation_pct: a.utilization_pct,
    run_hours: round(a.run_hours, 1),
    downtime_hours: round(a.downtime_hours, 1),
  });
  return {
    review_only: true,
    period: { start: period.start, end: period.end, label: period.label, month_to_date: Boolean(period.month_to_date) },
    compared_with: period.previous,
    definitions: {
      availability: "available hours ÷ scheduled hours (available = scheduled − downtime)",
      utilisation: "run hours ÷ scheduled hours",
    },
    fleet: kpiSummary(now),
    fleet_previous: kpiSummary(before),
    by_equipment_type: (now?.by_category || [])
      .filter((c) => !cat || String(c.category || "").toLowerCase().includes(cat))
      .slice(0, 12)
      .map((c) => ({ category: c.category, machines: c.asset_count, availability_pct: c.availability_pct, utilisation_pct: c.utilization_pct, downtime_hours: round(c.downtime_hours, 1) })),
    lowest_availability: [...assets].sort((a, b) => (a.availability_pct ?? 101) - (b.availability_pct ?? 101)).slice(0, 8).map(slim),
    most_downtime: [...assets].sort((a, b) => b.downtime_hours - a.downtime_hours).slice(0, 8).map(slim),
    breakdowns: breakdownsIn(dbConn, period.start, period.end, codes),
    breakdowns_previous: breakdownsIn(dbConn, period.previous.start, period.previous.end, codes),
  };
}

// The report files June can put on screen. Each builds a link to an existing
// report download; the browser opens it from a button the administrator taps.
const REPORTS = {
  cost_monthly_xlsx: { label: "Fleet cost — monthly (Excel)", needs: "month", url: (a) => `/api/reports/cost-monthly.xlsx?month=${a.month}`, file: (a) => `fleet-cost-${a.month}.xlsx`, download: true },
  maintenance_cost_by_equipment_pdf: { label: "Maintenance cost by equipment (PDF)", needs: "month_or_range", url: (a) => `/api/reports/maintenance-cost-by-equipment.pdf?${a.q}`, file: () => "maintenance-cost-by-equipment.pdf" },
  maintenance_cost_by_equipment_xlsx: { label: "Maintenance cost by equipment (Excel)", needs: "month_or_range", url: (a) => `/api/reports/maintenance-cost-by-equipment.xlsx?${a.q}`, file: () => "maintenance-cost-by-equipment.xlsx", download: true },
  monthly_pdf: { label: "Monthly report (PDF)", needs: "month", url: (a) => `/api/reports/monthly.pdf?month=${a.month}`, file: (a) => `monthly-${a.month}.pdf` },
  weekly_pdf: { label: "Weekly report (PDF)", needs: "range", url: (a) => `/api/reports/weekly.pdf?start=${a.from_date}&end=${a.to_date}`, file: () => "weekly-report.pdf" },
  daily_pdf: { label: "Daily report (PDF)", needs: "date", url: (a) => `/api/reports/daily.pdf?date=${a.date}`, file: (a) => `daily-${a.date}.pdf` },
  executive_kpi_pack_xlsx: { label: "Executive KPI pack (Excel)", needs: "month", url: (a) => `/api/reports/executive-kpi-pack.xlsx?month=${a.month}`, file: (a) => `executive-kpi-pack-${a.month}.xlsx`, download: true },
  gm_budget_meeting_docx: { label: "GM budget meeting pack (Word)", needs: "month", url: (a) => `/api/reports/gm-budget-meeting.docx?month=${a.month}`, file: (a) => `gm-budget-meeting-${a.month}.docx`, download: true },
  asset_history_pdf: { label: "Machine history (PDF)", needs: "asset", url: (a) => `/api/reports/asset-history/${encodeURIComponent(a.asset_code)}.pdf${a.from_date && a.to_date ? `?start=${a.from_date}&end=${a.to_date}` : ""}`, file: (a) => `asset-history-${a.asset_code}.pdf` },
};
export const JUNE_REPORT_KEYS = Object.keys(REPORTS);

/** A button-ready link to one of the existing report files. */
export function prepareJuneReportFile(args = {}, { today = new Date().toISOString().slice(0, 10) } = {}) {
  const spec = REPORTS[String(args.report || "")];
  if (!spec) return { error: `Unknown report. Choose one of: ${JUNE_REPORT_KEYS.join(", ")}.` };
  const lastMonth = ymd(Date.parse(`${today.slice(0, 7)}-01`) - dayMs).slice(0, 7);
  const a = {
    month: isMonth(args.month) ? args.month : lastMonth,
    from_date: isYmd(args.from_date) ? args.from_date : null,
    to_date: isYmd(args.to_date) ? args.to_date : null,
    date: isYmd(args.date) ? args.date : today,
    asset_code: String(args.asset_code || "").trim().toUpperCase(),
  };
  if (spec.needs === "range" && !(a.from_date && a.to_date)) {
    // Default to last full Monday–Sunday week.
    const d = new Date(`${today}T00:00:00Z`);
    const monday = Date.parse(today) - (((d.getUTCDay() + 6) % 7) + 7) * dayMs;
    a.from_date = ymd(monday);
    a.to_date = ymd(monday + 6 * dayMs);
  }
  if (spec.needs === "asset" && !a.asset_code) return { error: "Say which machine, for example A303AM." };
  a.q = a.from_date && a.to_date && spec.needs === "month_or_range" && !isMonth(args.month)
    ? `start=${a.from_date}&end=${a.to_date}`
    : `month=${a.month}`;
  const period = spec.needs === "date" ? a.date
    : spec.needs === "range" || (spec.needs === "month_or_range" && a.q.startsWith("start")) ? `${a.from_date} to ${a.to_date}`
    : spec.needs === "asset" ? (a.from_date && a.to_date ? `${a.from_date} to ${a.to_date}` : "all history")
    : a.month;
  return {
    ready: true,
    report: args.report,
    label: spec.label,
    period,
    asset_code: a.asset_code || null,
    url: spec.url(a),
    filename: spec.file(a),
    download: Boolean(spec.download),
    next_step: "A button to open it is on screen.",
  };
}
