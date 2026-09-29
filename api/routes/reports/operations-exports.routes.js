// IRONLOG/api/routes/reports/operations-exports.routes.js — Daily, GM, cost and maintenance-cost Excel exports, rain days.
// Registered by routes/reports.routes.js; shared helpers arrive through ctx.
import ExcelJS from "exceljs";
import fs from "node:fs";
import path from "node:path";
import { buildAmlWeeklyCheckSheet } from "../../utils/amlWeeklyCheckSheet.js";
import { buildPdfBuffer, kvGrid, sectionTitle, table, tryDrawLogo } from "../../utils/pdfGenerator.js";
import { db } from "../../db/client.js";
import { isDate } from "../../utils/request.js";
import { oilPartSql } from "../../utils/stockCategory.js";

export default function registerOperationsExportsRoutes(app, ctx) {
  const {
    AML_WEEKLY_TEMPLATE_PATH,
    addTableSheet,
    buildAmlWeeklyExportRecords,
    buildDailyExecutiveSummarySheet,
    buildGmWeeklyExecutiveSheet,
    buildMaintenanceCostByEquipment,
    compactCell,
    costDefaults,
    daysDownForBreakdown,
    downtimeMtdUsesDailyLogs,
    fmtNum,
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
    kpiDaily,
    kpiRange,
    monthRange,
    monthStartIso,
    prevMonth,
    reliabilityMetricsForRange,
    requireAdmin,
    resolveMaintenancePeriod,
    todayYmd,
  } = ctx;

  // GET /api/reports/aml-weekly-check-sheet.xlsx?week_ending=YYYY-MM-DD
  app.get("/aml-weekly-check-sheet.xlsx", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    const weekEnding = String(req.query?.week_ending || req.query?.end || "").trim();
    if (!isDate(weekEnding)) {
      return reply.code(400).send({ ok: false, error: "week_ending (YYYY-MM-DD) required" });
    }
    if (!fs.existsSync(AML_WEEKLY_TEMPLATE_PATH)) {
      req.log.error({ template: AML_WEEKLY_TEMPLATE_PATH }, "AML weekly template is missing");
      return reply.code(503).send({ ok: false, error: "AML weekly template is not installed" });
    }

    const siteCode = String(req.headers["x-site-code"] || req.query?.site_code || "main")
      .trim()
      .toLowerCase();
    const templateBuffer = fs.readFileSync(AML_WEEKLY_TEMPLATE_PATH);
    const records = buildAmlWeeklyExportRecords(weekEnding, siteCode);
    const workbook = await buildAmlWeeklyCheckSheet(templateBuffer, { weekEnding, records });
    const unmatched = workbook.unmatchedAssetCodes.slice(0, 20);
    if (unmatched.length) {
      req.log.warn({ unmatched, siteCode, weekEnding }, "AML weekly export skipped assets not present in the template");
    }

    return reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="AML_Weekly_Check_Sheet_${weekEnding}.xlsx"`)
      .header("X-Ironlog-AML-Filled-Rows", String(workbook.filledRows))
      .header("X-Ironlog-AML-Unmatched-Assets", unmatched.join(","))
      .send(workbook.buffer);
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
            AND COALESCE(l.hours_down, 0) > 0
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
  // PLANT LABOR & OIL (annual workbook — Plant No / Mechanics style)
  // GET /api/reports/plant-labor-oil.xlsx?year=2026&site_code=main
  // =========================
  app.get("/plant-labor-oil.xlsx", async (req, reply) => {
    const year = String(req.query?.year || new Date().getFullYear()).trim();
    if (!/^\d{4}$/.test(year)) {
      return reply.code(400).send({ error: "year (YYYY) required" });
    }
    const siteCode = String(req.query?.site_code || req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
    const start = `${year}-01-01`;
    const end = `${year}-12-31`;
    const defaults = costDefaults();
    const laborRate = Number(defaults.labor_cost_per_hour_default || 35);
    const lubeDefault = Number(defaults.lube_cost_per_qty_default || 4);

    const woOilStockSql = `(
      SELECT COALESCE(SUM(ABS(sm.quantity) * COALESCE(NULLIF(sm.unit_cost, 0), p.unit_cost, ${lubeDefault})), 0)
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      WHERE sm.reference = ('work_order:' || w.id)
        AND sm.movement_type = 'out'
        AND ${oilPartSql("p")}
    )`;

    const woLines = db.prepare(`
      SELECT
        w.id AS work_order_id,
        strftime('%Y-%m', COALESCE(w.completed_at, w.closed_at, w.opened_at)) AS year_month,
        CAST(strftime('%m', COALESCE(w.completed_at, w.closed_at, w.opened_at)) AS INTEGER) AS month_num,
        COALESCE(w.site_code, 'main') AS site_code,
        a.asset_code AS plant_no,
        a.asset_name,
        a.category,
        COALESCE(NULLIF(TRIM(w.assigned_artisan_name), ''), NULLIF(TRIM(w.artisan_name), ''), '') AS technician,
        COALESCE(w.labor_hours, 0) AS labor_hours,
        COALESCE(w.labor_rate_per_hour, ?) AS labor_rate_per_hour,
        COALESCE(w.oil_cost, 0) AS manual_oil_cost,
        ${woOilStockSql} AS issued_oil_cost,
        (COALESCE(w.oil_cost, 0) + ${woOilStockSql}) AS total_oil_cost,
        (COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)) AS labor_cost_usd,
        w.source,
        w.status,
        DATE(COALESCE(w.completed_at, w.closed_at, w.opened_at)) AS work_date
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE DATE(COALESCE(w.completed_at, w.closed_at)) BETWEEN ? AND ?
        AND w.status IN ('completed', 'approved', 'closed')
        AND LOWER(TRIM(COALESCE(w.site_code, 'main'))) = ?
      ORDER BY work_date ASC, w.id ASC
    `).all(laborRate, laborRate, start, end, siteCode);

    const monthNames = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

    const plantYearMap = new Map();
    const monthlyMap = new Map();
    for (const r of woLines) {
      const plant = String(r.plant_no || "");
      if (!plantYearMap.has(plant)) {
        plantYearMap.set(plant, {
          plant_no: plant,
          asset_name: r.asset_name,
          category: r.category,
          labor_hours: 0,
          labor_cost_usd: 0,
          oil_cost_usd: 0,
          work_orders: 0,
        });
      }
      const py = plantYearMap.get(plant);
      py.labor_hours += Number(r.labor_hours || 0);
      py.labor_cost_usd += Number(r.labor_cost_usd || 0);
      py.oil_cost_usd += Number(r.total_oil_cost || 0);
      py.work_orders += 1;

      const mKey = `${r.year_month}|${plant}`;
      if (!monthlyMap.has(mKey)) {
        monthlyMap.set(mKey, {
          month: monthNames[Number(r.month_num || 0)] || String(r.year_month || ""),
          year_month: r.year_month,
          site: r.site_code,
          plant_no: plant,
          asset_name: r.asset_name,
          labor_hours: 0,
          labor_cost_usd: 0,
          oil_cost_usd: 0,
        });
      }
      const mo = monthlyMap.get(mKey);
      mo.labor_hours += Number(r.labor_hours || 0);
      mo.labor_cost_usd += Number(r.labor_cost_usd || 0);
      mo.oil_cost_usd += Number(r.total_oil_cost || 0);
    }

    const plantSummary = Array.from(plantYearMap.values())
      .map((r) => ({
        ...r,
        labor_hours: Number(r.labor_hours.toFixed(2)),
        labor_cost_usd: Number(r.labor_cost_usd.toFixed(2)),
        oil_cost_usd: Number(r.oil_cost_usd.toFixed(2)),
      }))
      .sort((a, b) => a.plant_no.localeCompare(b.plant_no));

    const monthlyRows = Array.from(monthlyMap.values())
      .map((r) => ({
        ...r,
        labor_hours: Number(r.labor_hours.toFixed(2)),
        labor_cost_usd: Number(r.labor_cost_usd.toFixed(2)),
        oil_cost_usd: Number(r.oil_cost_usd.toFixed(2)),
      }))
      .sort((a, b) => String(a.year_month).localeCompare(String(b.year_month)) || a.plant_no.localeCompare(b.plant_no));

    const technicians = db.prepare(`
      SELECT DISTINCT COALESCE(NULLIF(TRIM(assigned_artisan_name), ''), NULLIF(TRIM(artisan_name), '')) AS technician_name
      FROM work_orders
      WHERE DATE(COALESCE(completed_at, closed_at)) BETWEEN ? AND ?
        AND TRIM(COALESCE(NULLIF(TRIM(assigned_artisan_name), ''), NULLIF(TRIM(artisan_name), ''), '')) <> ''
      ORDER BY technician_name ASC
    `).all(start, end);

    const maintPlans = hasTable("maintenance_plans")
      ? db.prepare(`
          SELECT
            a.asset_code AS plant_no,
            a.asset_name,
            a.category,
            mp.service_name,
            mp.interval_hours,
            mp.last_service_hours,
            mp.active
          FROM maintenance_plans mp
          JOIN assets a ON a.id = mp.asset_id
          WHERE COALESCE(mp.active, 1) = 1
          ORDER BY a.asset_code ASC, mp.service_name ASC
        `).all()
      : [];

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const wsHead = wb.addWorksheet("Plant Summary");
    wsHead.columns = [
      { header: "Plant No", key: "plant_no", width: 14 },
      { header: "Description", key: "asset_name", width: 28 },
      { header: "Category", key: "category", width: 14 },
      { header: "Total Hours Worked", key: "labor_hours", width: 18 },
      { header: "Labor USD", key: "labor_cost_usd", width: 14 },
      { header: "Oil USD", key: "oil_cost_usd", width: 12 },
      { header: "Work Orders", key: "work_orders", width: 12 },
    ];
    wsHead.getRow(1).font = { bold: true };
    wsHead.addRows(plantSummary.length ? plantSummary : [{
      plant_no: "-", asset_name: "No labor/oil on work orders for year", category: "", labor_hours: 0, labor_cost_usd: 0, oil_cost_usd: 0, work_orders: 0,
    }]);

    const wsMonthly = wb.addWorksheet("Monthly by Plant");
    wsMonthly.columns = [
      { header: "Month", key: "month", width: 14 },
      { header: "Year-Month", key: "year_month", width: 12 },
      { header: "Site", key: "site", width: 10 },
      { header: "Plant No", key: "plant_no", width: 14 },
      { header: "Description", key: "asset_name", width: 24 },
      { header: "Total Hours Worked", key: "labor_hours", width: 18 },
      { header: "Labor USD", key: "labor_cost_usd", width: 14 },
      { header: "Oil USD", key: "oil_cost_usd", width: 12 },
    ];
    wsMonthly.getRow(1).font = { bold: true };
    wsMonthly.addRows(monthlyRows);

    const wsWo = wb.addWorksheet("Work Orders");
    wsWo.columns = [
      { header: "WO #", key: "work_order_id", width: 8 },
      { header: "Date", key: "work_date", width: 12 },
      { header: "Month", key: "year_month", width: 10 },
      { header: "Plant No", key: "plant_no", width: 14 },
      { header: "Technician", key: "technician", width: 18 },
      { header: "Repair Hours", key: "labor_hours", width: 14 },
      { header: "Labor Rate", key: "labor_rate_per_hour", width: 12 },
      { header: "Labor USD", key: "labor_cost_usd", width: 12 },
      { header: "Manual Oil USD", key: "manual_oil_cost", width: 14 },
      { header: "Issued Oil USD", key: "issued_oil_cost", width: 14 },
      { header: "Total Oil USD", key: "total_oil_cost", width: 14 },
      { header: "Source", key: "source", width: 12 },
      { header: "Status", key: "status", width: 12 },
    ];
    wsWo.getRow(1).font = { bold: true };
    wsWo.addRows(woLines.map((r) => ({
      ...r,
      labor_hours: Number(Number(r.labor_hours || 0).toFixed(2)),
      labor_rate_per_hour: Number(Number(r.labor_rate_per_hour || 0).toFixed(2)),
      labor_cost_usd: Number(Number(r.labor_cost_usd || 0).toFixed(2)),
      manual_oil_cost: Number(Number(r.manual_oil_cost || 0).toFixed(2)),
      issued_oil_cost: Number(Number(r.issued_oil_cost || 0).toFixed(2)),
      total_oil_cost: Number(Number(r.total_oil_cost || 0).toFixed(2)),
    })));

    const wsTech = wb.addWorksheet("Technicians");
    wsTech.columns = [
      { header: "Technician Name", key: "technician_name", width: 28 },
    ];
    wsTech.getRow(1).font = { bold: true };
    wsTech.addRows(technicians);

    const wsSched = wb.addWorksheet("Maintenance Schedule");
    wsSched.columns = [
      { header: "Plant No", key: "plant_no", width: 14 },
      { header: "Description", key: "asset_name", width: 24 },
      { header: "Category", key: "category", width: 14 },
      { header: "Service", key: "service_name", width: 22 },
      { header: "Interval Hrs", key: "interval_hours", width: 12 },
      { header: "Last Service Hrs", key: "last_service_hours", width: 16 },
    ];
    wsSched.getRow(1).font = { bold: true };
    wsSched.addRows(maintPlans);

    const buffer = await wb.xlsx.writeBuffer();
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Plant_Labor_Oil_${year}_${siteCode}.xlsx"`)
      .send(Buffer.from(buffer));
  });

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
}
