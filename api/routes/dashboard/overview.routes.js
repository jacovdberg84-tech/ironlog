// IRONLOG/api/routes/dashboard/overview.routes.js — Main dashboard, reliability, cost trend and work order nudges.
// Registered by routes/dashboard.routes.js; shared helpers arrive through ctx.
import { PRESTART_DEDUCTION_HOURS } from "../../utils/prestartDaily.js";
import { db } from "../../db/client.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerOverviewRoutes(app, ctx) {
  const {
    bdLogSiteSql,
    computeFleetKpiForDay,
    dailyHoursHasSite,
    eachDateInclusiveYMD,
    getDowntimeReasonsNoSite,
    getDowntimeReasonsWithSite,
    hasColumn,
    majorDowntimeNoSite,
    majorDowntimeWithSite,
    monthRangeFromYYYYMM,
    monthStartIso,
    requireRoles,
    siteCodeFromReq,
    slaPriority,
    todayYYYYMMDD,
  } = ctx;

  // GET /api/dashboard?date=YYYY-MM-DD&scheduled=10
  app.get("/", async (req, reply) => {
    const date = String(req.query?.date || todayYYYYMMDD()).trim();
    const scheduledRaw = Number(req.query?.scheduled ?? 10);
    const scheduledFallback = Number.isFinite(scheduledRaw) && scheduledRaw > 0
      ? scheduledRaw
      : 10;
    const siteCode = siteCodeFromReq(req);

    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
    }

    // =========================
    // KPI (Split) — gauges: month-to-date through selected date; per-asset table: selected day
    // Planned hours = sum of per-asset scheduled (daily_hours.scheduled_hours, else header fallback).
    // Machine-available hours (MTD) = planned − downtime (capped per asset-day).
    // Availability = machine-available ÷ planned = (planned − downtime) / planned.
    // Utilization = run ÷ planned (same denominator as the plan, not reduced by downtime).
    // Logged downtime uses breakdown_downtime_logs.log_date = each day. OPEN-breakdown imputed
    // downtime (when logs are zero) only applies if breakdown_date is in the same YYYY-MM as that day
    // so prior-month incidents do not inflate the current month.
    // =========================

    const dayK = computeFleetKpiForDay(date, scheduledFallback, { includePerAsset: true, siteCode });
    const per_asset_kpi = dayK.per_asset_kpi;
    const run_hours = dayK.run_hours;

    const mtdStart = monthStartIso(date);
    let mtd_scheduled = 0;
    let mtd_run = 0;
    let mtd_downtime = 0;
    let mtd_inspection = 0;
    let mtd_availability_loss = 0;
    let mtd_day_count = 0;
    const mtdAssetIds = new Set();
    eachDateInclusiveYMD(mtdStart, date, (dayStr) => {
      mtd_day_count += 1;
      const dr = computeFleetKpiForDay(dayStr, scheduledFallback, { includePerAsset: false, siteCode });
      mtd_scheduled += dr.scheduled_hours;
      mtd_run += dr.run_hours;
      mtd_downtime += dr.downtime_hours;
      mtd_inspection += dr.inspection_hours;
      mtd_availability_loss += dr.availability_loss_hours;
      dr.contributingAssetIds.forEach((id) => mtdAssetIds.add(id));
    });

    // Safety fallback: if MTD scheduled remained zero, derive from active fleet (single-site DBs only).
    if (!dailyHoursHasSite && mtd_scheduled <= 0 && scheduledFallback > 0 && mtd_day_count > 0) {
      const activeFleetCountRow = db.prepare(`
        SELECT COUNT(*) AS c
        FROM assets
        WHERE active = 1
          AND is_standby = 0
      `).get();
      const activeFleetCount = Number(activeFleetCountRow?.c || 0);
      if (activeFleetCount > 0) {
        mtd_scheduled = activeFleetCount * scheduledFallback * mtd_day_count;
      }
    }

    const available_hours_mtd = Math.max(0, mtd_scheduled - mtd_availability_loss);
    const availability_mtd =
      mtd_scheduled > 0 ? (available_hours_mtd / mtd_scheduled) * 100 : null;
    const utilization_mtd =
      mtd_scheduled > 0 ? (mtd_run / mtd_scheduled) * 100 : null;
    const used_assets = mtdAssetIds.size;

    const day_scheduled = dayK.scheduled_hours;
    const day_run = dayK.run_hours;
    const day_downtime = dayK.downtime_hours;
    const day_inspection = dayK.inspection_hours;
    const day_availability_loss = dayK.availability_loss_hours;
    const available_hours_day = Math.max(0, day_scheduled - day_availability_loss);
    const availability_day =
      day_scheduled > 0 ? (available_hours_day / day_scheduled) * 100 : null;
    const utilization_day =
      day_scheduled > 0 ? (day_run / day_scheduled) * 100 : null;
    // Gauges: prefer selected day; if that day has no hour-meter planned time, use MTD so the dash is not stuck on N/A.
    const gauge_basis = day_scheduled > 0 ? "selected_day" : "mtd";
    const availability =
      availability_day != null ? availability_day : availability_mtd;
    const utilization =
      utilization_day != null ? utilization_day : utilization_mtd;

    const scheduled_hours = mtd_scheduled;
    const downtime_hours = mtd_downtime;
    // run_hours = selected day only (for cost-per-run-hour); gauges use mtd_* above

    // Alerts summaries
    const lowStockCount = db.prepare(`
      SELECT COUNT(*) AS c
      FROM (
        SELECT p.id, IFNULL(SUM(sm.quantity),0) AS on_hand, p.min_stock
        FROM parts p
        LEFT JOIN stock_movements sm ON sm.part_id = p.id
        GROUP BY p.id
        HAVING on_hand < p.min_stock
      )
    `).get();

    const overdueMaintCount = db.prepare(`
      SELECT COUNT(*) AS c
      FROM (
        SELECT
          mp.id,
          (IFNULL((
            SELECT SUM(dh.hours_run)
            FROM daily_hours dh
            WHERE dh.asset_id = mp.asset_id
              AND dh.is_used = 1
              AND dh.hours_run > 0
              AND dh.work_date <= ?
          ),0) - (mp.last_service_hours + mp.interval_hours)) AS diff
        FROM maintenance_plans mp
        JOIN assets a ON a.id = mp.asset_id
        WHERE mp.active = 1 AND a.active = 1 AND a.is_standby = 0
      )
      WHERE diff >= 0
    `).get(date);

    const hasWOCompletedAt = hasColumn("work_orders", "completed_at");
    const hasBreakdownStatus = hasColumn("breakdowns", "status");
    const woCompletedFilter = hasWOCompletedAt
      ? "AND (w.completed_at IS NULL OR TRIM(COALESCE(w.completed_at, '')) = '')"
      : "";
    const breakdownOpenFilter = hasBreakdownStatus
      ? `AND (
          w.source <> 'breakdown'
          OR TRIM(LOWER(COALESCE(b.status, ''))) IN ('open', 'in_progress')
        )`
      : "";

    const openWOCount = db.prepare(`
      SELECT COUNT(*) AS c
      FROM work_orders w
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
      WHERE REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
        AND (w.closed_at IS NULL OR TRIM(COALESCE(w.closed_at, '')) = '')
        ${woCompletedFilter}
        ${breakdownOpenFilter}
    `).get();

    const majorDowntime = (
      majorDowntimeWithSite && bdLogSiteSql
        ? majorDowntimeWithSite.all(date, siteCode)
        : majorDowntimeNoSite.all(date)
    ).map((r) => ({
      ...r,
      downtime_hours: Number(r.downtime_hours || 0),
      critical: Boolean(r.critical),
    }));

    const downtimeReasons = (
      getDowntimeReasonsWithSite && bdLogSiteSql
        ? getDowntimeReasonsWithSite.all(date, siteCode)
        : getDowntimeReasonsNoSite.all(date)
    ).map((r) => ({
      reason: r.reason,
      hours_down: Number(r.hours_down || 0),
      incidents: Number(r.incidents || 0)
    }));

    const criticalLowStock = db.prepare(`
      SELECT p.part_code, p.part_name, p.min_stock, IFNULL(SUM(sm.quantity),0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      WHERE p.critical = 1
      GROUP BY p.id
      HAVING on_hand < p.min_stock
      ORDER BY on_hand ASC
      LIMIT 8
    `).all().map(r => ({ ...r, on_hand: Number(r.on_hand) }));

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
      WHERE REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
        AND (w.closed_at IS NULL OR TRIM(COALESCE(w.closed_at, '')) = '')
        ${woCompletedFilter}
        ${breakdownOpenFilter}
      ORDER BY w.id DESC
      LIMIT 8
    `).all();

    const openWOSla = db.prepare(`
      SELECT
        w.id,
        a.asset_code,
        w.source,
        w.status,
        CASE
          WHEN w.source = 'breakdown' THEN COALESCE(NULLIF(TRIM(b.start_at), ''), NULLIF(TRIM(b.breakdown_date), ''), w.opened_at)
          ELSE w.opened_at
        END AS opened_at,
        CAST((
          julianday('now') - julianday(
            COALESCE(
              CASE
                WHEN w.source = 'breakdown' THEN COALESCE(NULLIF(TRIM(b.start_at), ''), NULLIF(TRIM(b.breakdown_date), ''), w.opened_at)
                ELSE w.opened_at
              END,
              datetime('now')
            )
          )
        ) * 24 AS INTEGER) AS age_hours
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
      WHERE REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
        AND (w.closed_at IS NULL OR TRIM(COALESCE(w.closed_at, '')) = '')
        ${woCompletedFilter}
        ${breakdownOpenFilter}
      ORDER BY age_hours DESC, w.id DESC
      LIMIT 200
    `).all().map((r) => ({
      id: Number(r.id),
      asset_code: r.asset_code,
      source: r.source,
      status: r.status,
      opened_at: r.opened_at,
      age_hours: Number(r.age_hours || 0),
    }));

    const sla_summary = {
      open_gt_24h: openWOSla.filter((r) => r.age_hours > 24).length,
      in_progress_gt_48h: openWOSla.filter((r) => String(r.status) === "in_progress" && r.age_hours > 48).length,
      completed_gt_12h: 0,
    };
    const sla_breaches = openWOSla
      .filter((r) => {
        const s = String(r.status || "").toLowerCase();
        if (s === "in_progress") return r.age_hours > 48;
        return r.age_hours > 24;
      })
      .map((r) => ({
        ...r,
        priority: slaPriority(r.status, r.age_hours),
      }))
      .slice(0, 8);

    const lubeDaily = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ol.quantity), 0) AS qty,
        COALESCE(SUM(
          ol.quantity * COALESCE(
            ol.unit_cost,
            (SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1),
            4.0
          )
        ), 0) AS lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date = ?
      GROUP BY a.id
      ORDER BY qty DESC, a.asset_code ASC
      LIMIT 8
    `).all(date).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      qty: Number(r.qty || 0),
      lube_cost: Number(r.lube_cost || 0),
    }));

    const lubeDailyByType = db.prepare(`
      SELECT
        a.asset_code,
        CASE
          WHEN LOWER(TRIM(COALESCE(ol.oil_type, ''))) IN ('admin','supervisor','manager','stores','artisan','operator') THEN 'UNSPECIFIED'
          ELSE COALESCE(NULLIF(TRIM(ol.oil_type), ''), 'UNSPECIFIED')
        END AS oil_type,
        COALESCE(SUM(ol.quantity), 0) AS qty,
        COALESCE(SUM(
          ol.quantity * COALESCE(
            ol.unit_cost,
            (SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1),
            4.0
          )
        ), 0) AS lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date = ?
      GROUP BY a.asset_code, oil_type
      ORDER BY a.asset_code ASC, qty DESC, oil_type ASC
      LIMIT 400
    `).all(date);
    const byTypeMap = new Map();
    for (const r of lubeDailyByType) {
      const code = String(r.asset_code || "");
      if (!byTypeMap.has(code)) byTypeMap.set(code, []);
      byTypeMap.get(code).push({
        oil_type: String(r.oil_type || "UNSPECIFIED"),
        qty: Number(r.qty || 0),
        lube_cost: Number(r.lube_cost || 0),
      });
    }
    for (const row of lubeDaily) {
      row.by_oil_type = byTypeMap.get(String(row.asset_code || "")) || [];
    }

    const lubeTotalRow = db.prepare(`
      SELECT
        COALESCE(SUM(quantity), 0) AS qty_total,
        COALESCE(SUM(
          quantity * COALESCE(
            unit_cost,
            (SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1),
            4.0
          )
        ), 0) AS total_lube_cost
      FROM oil_logs
      WHERE log_date = ?
    `).get(date);

    const settingsRows = db.prepare(`
      SELECT key, value
      FROM cost_settings
      WHERE key IN (
        'fuel_cost_per_liter_default',
        'lube_cost_per_qty_default',
        'labor_cost_per_hour_default',
        'downtime_cost_per_hour_default'
      )
    `).all();
    const settings = {
      fuel_cost_per_liter_default: 1.5,
      lube_cost_per_qty_default: 4.0,
      labor_cost_per_hour_default: 35.0,
      downtime_cost_per_hour_default: 120.0,
    };
    for (const r of settingsRows) {
      const k = String(r.key || "").trim();
      const v = Number(r.value);
      if (k && Number.isFinite(v)) settings[k] = v;
    }

    const fuelCostRow = db.prepare(`
      SELECT COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS fuel_cost
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date = ?
    `).get(settings.fuel_cost_per_liter_default, date);

    const lubeCostRow = db.prepare(`
      SELECT COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS lube_cost
      FROM oil_logs ol
      WHERE ol.log_date = ?
    `).get(settings.lube_cost_per_qty_default, date);

    const woCostRow = db.prepare(`
      SELECT
        COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS labor_hours,
        COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      WHERE DATE(COALESCE(w.completed_at, w.closed_at)) = ?
        AND w.status IN ('completed', 'approved', 'closed')
    `).get(settings.labor_cost_per_hour_default, date);

    const downtimeCostRow = db.prepare(`
      SELECT COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS downtime_cost
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date = ?
    `).get(settings.downtime_cost_per_hour_default, date);

    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
    const partsCostRows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS parts_cost
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      LEFT JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
      LEFT JOIN assets a ON a.id = w.asset_id
      WHERE sm.movement_type = 'out'
        AND ${smDateExpr} = ?
      GROUP BY a.id
    `).all(date).map((r) => ({
      asset_code: r.asset_code || "UNLINKED",
      asset_name: r.asset_name || "Unlinked to WO",
      parts_cost: Number(r.parts_cost || 0),
    }));
    const parts_cost = Number(partsCostRows.reduce((acc, r) => acc + Number(r.parts_cost || 0), 0).toFixed(2));

    const assetFuelCost = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS fuel_cost
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date = ?
      GROUP BY a.id
    `).all(settings.fuel_cost_per_liter_default, date);

    const assetLubeCost = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date = ?
      GROUP BY a.id
    `).all(settings.lube_cost_per_qty_default, date);

    const assetLaborCost = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE DATE(COALESCE(w.completed_at, w.closed_at)) = ?
        AND w.status IN ('completed', 'approved', 'closed')
      GROUP BY a.id
    `).all(settings.labor_cost_per_hour_default, date);

    const assetDowntimeCost = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS downtime_cost
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date = ?
      GROUP BY a.id
    `).all(settings.downtime_cost_per_hour_default, date);

    const byAsset = new Map();
    const putCost = (rows, keyName) => {
      for (const r of rows) {
        const code = String(r.asset_code || "UNLINKED");
        if (!byAsset.has(code)) {
          byAsset.set(code, {
            asset_code: code,
            asset_name: r.asset_name || "Unlinked",
            fuel_cost: 0,
            lube_cost: 0,
            parts_cost: 0,
            labor_cost: 0,
            downtime_cost: 0,
            total_cost: 0,
          });
        }
        const row = byAsset.get(code);
        row[keyName] += Number(r[keyName] || 0);
      }
    };
    putCost(assetFuelCost, "fuel_cost");
    putCost(assetLubeCost, "lube_cost");
    putCost(partsCostRows, "parts_cost");
    putCost(assetLaborCost, "labor_cost");
    putCost(assetDowntimeCost, "downtime_cost");
    const top_asset_costs = Array.from(byAsset.values())
      .map((r) => ({
        ...r,
        fuel_cost: Number(r.fuel_cost.toFixed(2)),
        lube_cost: Number(r.lube_cost.toFixed(2)),
        parts_cost: Number(r.parts_cost.toFixed(2)),
        labor_cost: Number(r.labor_cost.toFixed(2)),
        downtime_cost: Number(r.downtime_cost.toFixed(2)),
        total_cost: Number((r.fuel_cost + r.lube_cost + r.parts_cost + r.labor_cost + r.downtime_cost).toFixed(2)),
      }))
      .filter((r) => r.total_cost > 0)
      .sort((a, b) => b.total_cost - a.total_cost)
      .slice(0, 8);

    const fuel_cost = Number(fuelCostRow?.fuel_cost || 0);
    const lube_cost = Number(lubeCostRow?.lube_cost || 0);
    const labor_cost = Number(woCostRow?.labor_cost || 0);
    const labor_hours = Number(woCostRow?.labor_hours || 0);
    const downtime_cost = Number(downtimeCostRow?.downtime_cost || 0);
    const total_cost = Number((fuel_cost + lube_cost + parts_cost + labor_cost + downtime_cost).toFixed(2));
    const cost_per_run_hour = run_hours > 0 ? Number((total_cost / run_hours).toFixed(2)) : null;

    return {
      ok: true,
      date,

      // Keep this for UI display; KPI uses per-row scheduled from daily_hours
      scheduled_hours_per_asset: scheduledFallback,

      kpi: {
        site_code: siteCode,
        gauge_basis,
        used_assets,
        scheduled_hours,
        available_hours: Number(available_hours_mtd.toFixed(2)),
        utilization_base_hours: Number(mtd_scheduled.toFixed(2)),
        run_hours: day_run,
        run_hours_mtd: Number(mtd_run.toFixed(2)),
        downtime_hours,
        inspection_hours: Number(mtd_inspection.toFixed(2)),
        availability: availability == null ? null : Number(availability.toFixed(2)),
        utilization: utilization == null ? null : Number(utilization.toFixed(2)),
        availability_day: availability_day == null ? null : Number(availability_day.toFixed(2)),
        utilization_day: utilization_day == null ? null : Number(utilization_day.toFixed(2)),
        availability_mtd: availability_mtd == null ? null : Number(availability_mtd.toFixed(2)),
        utilization_mtd: utilization_mtd == null ? null : Number(utilization_mtd.toFixed(2)),
        scheduled_hours_day: Number(day_scheduled.toFixed(2)),
        downtime_hours_day: Number(day_downtime.toFixed(2)),
        inspection_hours_day: Number(day_inspection.toFixed(2)),
        inspection_deduction_hours_per_check: PRESTART_DEDUCTION_HOURS,
        basis: "gauges_day_or_mtd_fallback_meta_mtd",
        mtd_start: mtdStart,
        mtd_end: date,
      },

      alerts: {
        low_stock: Number(lowStockCount.c || 0),
        overdue_maintenance: Number(overdueMaintCount.c || 0),
        open_work_orders: Number(openWOCount.c || 0)
      },

      major_downtime: majorDowntime,
      downtime_reasons: downtimeReasons,
      per_asset_kpi,

      critical_low_stock: criticalLowStock,
      open_work_orders: openWOs,
      workorder_sla: {
        summary: sla_summary,
        breaches: sla_breaches,
      },
      lube_usage: {
        date,
        qty_total: Number(lubeTotalRow?.qty_total || 0),
        total_lube_cost: Number(lubeTotalRow?.total_lube_cost || 0),
        rows: lubeDaily,
      },
      cost_engine: {
        date,
        settings,
        fuel_cost: Number(fuel_cost.toFixed(2)),
        lube_cost: Number(lube_cost.toFixed(2)),
        parts_cost,
        labor_cost: Number(labor_cost.toFixed(2)),
        labor_hours: Number(labor_hours.toFixed(2)),
        downtime_cost: Number(downtime_cost.toFixed(2)),
        total_cost,
        cost_per_run_hour,
        top_asset_costs,
      },
    };
  });

  // GET /api/dashboard/cost/trend?months=12&end_month=YYYY-MM
  app.get("/cost/trend", async (req, reply) => {
    const monthsRaw = Number(req.query?.months ?? 12);
    const months = Number.isFinite(monthsRaw) ? Math.max(1, Math.min(36, Math.trunc(monthsRaw))) : 12;
    const endMonthInput = String(req.query?.end_month || "").trim();
    const nowMonth = todayYYYYMMDD().slice(0, 7);
    const endMonth = /^\d{4}-\d{2}$/.test(endMonthInput) ? endMonthInput : nowMonth;
    if (!/^\d{4}-\d{2}$/.test(endMonth)) {
      return reply.code(400).send({ error: "end_month must be YYYY-MM" });
    }

    const settingsRows = db.prepare(`
      SELECT key, value
      FROM cost_settings
      WHERE key IN (
        'fuel_cost_per_liter_default',
        'lube_cost_per_qty_default',
        'labor_cost_per_hour_default',
        'downtime_cost_per_hour_default'
      )
    `).all();
    const settings = {
      fuel_cost_per_liter_default: 1.5,
      lube_cost_per_qty_default: 4.0,
      labor_cost_per_hour_default: 35.0,
      downtime_cost_per_hour_default: 120.0,
    };
    for (const r of settingsRows) {
      const k = String(r.key || "").trim();
      const v = Number(r.value);
      if (k && Number.isFinite(v)) settings[k] = v;
    }

    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";

    const getMonthCost = (monthId) => {
      const r = monthRangeFromYYYYMM(monthId);

      const fuel = db.prepare(`
        SELECT COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS v
        FROM fuel_logs fl
        JOIN assets a ON a.id = fl.asset_id
        WHERE fl.log_date BETWEEN ? AND ?
      `).get(settings.fuel_cost_per_liter_default, r.start, r.end);

      const lube = db.prepare(`
        SELECT COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS v
        FROM oil_logs ol
        WHERE ol.log_date BETWEEN ? AND ?
      `).get(settings.lube_cost_per_qty_default, r.start, r.end);

      const parts = db.prepare(`
        SELECT COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS v
        FROM stock_movements sm
        JOIN parts p ON p.id = sm.part_id
        WHERE sm.movement_type = 'out'
          AND ${smDateExpr} BETWEEN ? AND ?
      `).get(r.start, r.end);

      const labor = db.prepare(`
        SELECT
          COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS hrs,
          COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS v
        FROM work_orders w
        WHERE DATE(COALESCE(w.completed_at, w.closed_at)) BETWEEN ? AND ?
          AND w.status IN ('completed', 'approved', 'closed')
      `).get(settings.labor_cost_per_hour_default, r.start, r.end);

      const down = db.prepare(`
        SELECT COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS v
        FROM breakdown_downtime_logs l
        JOIN breakdowns b ON b.id = l.breakdown_id
        JOIN assets a ON a.id = b.asset_id
        WHERE l.log_date BETWEEN ? AND ?
      `).get(settings.downtime_cost_per_hour_default, r.start, r.end);

      const run = db.prepare(`
        SELECT COALESCE(SUM(hours_run), 0) AS h
        FROM daily_hours
        WHERE work_date BETWEEN ? AND ?
          AND is_used = 1
          AND hours_run > 0
      `).get(r.start, r.end);

      const fuel_cost = Number(fuel?.v || 0);
      const lube_cost = Number(lube?.v || 0);
      const parts_cost = Number(parts?.v || 0);
      const labor_cost = Number(labor?.v || 0);
      const labor_hours = Number(labor?.hrs || 0);
      const downtime_cost = Number(down?.v || 0);
      const run_hours = Number(run?.h || 0);
      const total_cost = Number((fuel_cost + lube_cost + parts_cost + labor_cost + downtime_cost).toFixed(2));
      return {
        month: monthId,
        fuel_cost: Number(fuel_cost.toFixed(2)),
        lube_cost: Number(lube_cost.toFixed(2)),
        parts_cost: Number(parts_cost.toFixed(2)),
        labor_cost: Number(labor_cost.toFixed(2)),
        labor_hours: Number(labor_hours.toFixed(2)),
        downtime_cost: Number(downtime_cost.toFixed(2)),
        run_hours: Number(run_hours.toFixed(2)),
        total_cost,
        cost_per_run_hour: run_hours > 0 ? Number((total_cost / run_hours).toFixed(2)) : null,
      };
    };

    const endDate = new Date(`${endMonth}-01T00:00:00Z`);
    const rows = [];
    for (let i = months - 1; i >= 0; i--) {
      const d = new Date(endDate);
      d.setUTCMonth(d.getUTCMonth() - i);
      const m = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
      rows.push(getMonthCost(m));
    }

    const prev = rows.length >= 2 ? rows[rows.length - 2] : null;
    const cur = rows.length ? rows[rows.length - 1] : null;
    const mom_variance = prev && cur ? Number((cur.total_cost - prev.total_cost).toFixed(2)) : null;
    const mom_variance_pct = prev && cur && Number(prev.total_cost) > 0
      ? Number((((cur.total_cost - prev.total_cost) / prev.total_cost) * 100).toFixed(2))
      : null;

    return {
      ok: true,
      months,
      end_month: endMonth,
      rows,
      latest: cur,
      mom: {
        previous_month: prev?.month || null,
        current_month: cur?.month || null,
        variance: mom_variance,
        variance_pct: mom_variance_pct,
      },
    };
  });

  // GET /api/dashboard/reliability?start=YYYY-MM-DD&end=YYYY-MM-DD
  // MTBF = operating hours / failure count
  // LTTR = downtime hours / failure count
  app.get("/reliability", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }

    const failuresRow = db.prepare(`
      SELECT COUNT(*) AS n
      FROM breakdowns
      WHERE breakdown_date BETWEEN ? AND ?
    `).get(start, end);
    const failure_count = Number(failuresRow?.n || 0);

    const runRow = db.prepare(`
      SELECT COALESCE(SUM(hours_run), 0) AS run_hours
      FROM daily_hours
      WHERE work_date BETWEEN ? AND ?
        AND is_used = 1
        AND hours_run > 0
    `).get(start, end);
    const operating_hours = Number(runRow?.run_hours || 0);

    let downtime_hours = 0;
    if (hasColumn("breakdowns", "downtime_total_hours")) {
      const dtRow = db.prepare(`
        SELECT COALESCE(SUM(downtime_total_hours), 0) AS dt
        FROM breakdowns
        WHERE breakdown_date BETWEEN ? AND ?
      `).get(start, end);
      downtime_hours = Number(dtRow?.dt || 0);
    } else if (hasColumn("breakdowns", "downtime_hours")) {
      const dtRow = db.prepare(`
        SELECT COALESCE(SUM(downtime_hours), 0) AS dt
        FROM breakdowns
        WHERE breakdown_date BETWEEN ? AND ?
      `).get(start, end);
      downtime_hours = Number(dtRow?.dt || 0);
    }

    const mtbf_hours = failure_count > 0 ? operating_hours / failure_count : null;
    const lttr_hours = failure_count > 0 ? downtime_hours / failure_count : null;

    return {
      ok: true,
      start,
      end,
      failure_count,
      operating_hours: Number(operating_hours.toFixed(2)),
      downtime_hours: Number(downtime_hours.toFixed(2)),
      mtbf_hours: mtbf_hours == null ? null : Number(mtbf_hours.toFixed(2)),
      lttr_hours: lttr_hours == null ? null : Number(lttr_hours.toFixed(2)),
    };
  });

  // GET /api/dashboard/reliability/trend?weeks=12&end=YYYY-MM-DD
  app.get("/reliability/trend", async (req, reply) => {
    const weeks = Math.max(1, Math.min(52, Number(req.query?.weeks || 12)));
    const endStr = String(req.query?.end || todayYYYYMMDD()).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(endStr)) {
      return reply.code(400).send({ error: "end must be YYYY-MM-DD" });
    }
    const endDate = new Date(`${endStr}T00:00:00`);
    const points = [];

    for (let i = weeks - 1; i >= 0; i--) {
      const wEnd = new Date(endDate);
      wEnd.setDate(endDate.getDate() - (i * 7));
      const wStart = new Date(wEnd);
      wStart.setDate(wEnd.getDate() - 6);
      const fmt = (d) => d.toISOString().slice(0, 10);
      const start = fmt(wStart);
      const end = fmt(wEnd);

      const failuresRow = db.prepare(`
        SELECT COUNT(*) AS n
        FROM breakdowns
        WHERE breakdown_date BETWEEN ? AND ?
      `).get(start, end);
      const failure_count = Number(failuresRow?.n || 0);

      const runRow = db.prepare(`
        SELECT COALESCE(SUM(hours_run), 0) AS run_hours
        FROM daily_hours
        WHERE work_date BETWEEN ? AND ?
          AND is_used = 1
          AND hours_run > 0
      `).get(start, end);
      const operating_hours = Number(runRow?.run_hours || 0);

      let downtime_hours = 0;
      if (hasColumn("breakdowns", "downtime_total_hours")) {
        const dtRow = db.prepare(`
          SELECT COALESCE(SUM(downtime_total_hours), 0) AS dt
          FROM breakdowns
          WHERE breakdown_date BETWEEN ? AND ?
        `).get(start, end);
        downtime_hours = Number(dtRow?.dt || 0);
      } else if (hasColumn("breakdowns", "downtime_hours")) {
        const dtRow = db.prepare(`
          SELECT COALESCE(SUM(downtime_hours), 0) AS dt
          FROM breakdowns
          WHERE breakdown_date BETWEEN ? AND ?
        `).get(start, end);
        downtime_hours = Number(dtRow?.dt || 0);
      }

      const mtbf_hours = failure_count > 0 ? operating_hours / failure_count : null;
      const lttr_hours = failure_count > 0 ? downtime_hours / failure_count : null;

      points.push({
        start,
        end,
        label: end.slice(5),
        failure_count,
        mtbf_hours: mtbf_hours == null ? null : Number(mtbf_hours.toFixed(2)),
        lttr_hours: lttr_hours == null ? null : Number(lttr_hours.toFixed(2)),
      });
    }

    return { ok: true, weeks, end: endStr, points };
  });

  // POST /api/dashboard/workorders/:id/nudge
  app.post("/workorders/:id/nudge", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "artisan"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "valid work order id required" });

    const wo = db.prepare(`
      SELECT id, status, opened_at
      FROM work_orders
      WHERE id = ?
    `).get(id);
    if (!wo) return reply.code(404).send({ error: "work order not found" });

    const note = String(req.body?.note || "").trim() || null;
    writeAudit(db, req, {
      module: "workorders",
      action: "nudge_supervisor",
      entity_type: "work_order",
      entity_id: id,
      payload: {
        status: wo.status,
        opened_at: wo.opened_at,
        note,
      },
    });

    return { ok: true, id, nudged: true };
  });
}
