// IRONLOG/api/routes/reports/period-reports.routes.js — Daily, weekly, monthly, operations and executive pack reports.
// Registered by routes/reports.routes.js; shared helpers arrive through ctx.
import ExcelJS from "exceljs";
import path from "node:path";
import { PRESTART_DEDUCTION_HOURS, listDailyPrestarts } from "../../utils/prestartDaily.js";
import { buildPdfBuffer, ensurePageSpace, kvGrid, sectionTitle, table, tryDrawLogo } from "../../utils/pdfGenerator.js";
import { buildSpeedingReportPdfContent, getCartrackSpeedAlertKmh, listCartrackEventsFromDb, summarizeSpeedingEvents } from "../../utils/cartrack.js";
import { db } from "../../db/client.js";
import { getPdfReportBranding } from "../../utils/reportSettings.js";
import { isDate } from "../../utils/request.js";
import { listPlannedMaintenanceForDate } from "../../utils/shortBreakdowns.js";

export default function registerPeriodReportsRoutes(app, ctx) {
  const {
    addTableSheet,
    asArray,
    buildPeriodAssetCosts,
    buildPeriodContractorFuelRows,
    compactCell,
    costDefaults,
    dailyPdfDowntimeHours,
    dailyPdfLongDate,
    dailyPdfManagementSection,
    dailyPdfOperationsDate,
    dailyPdfRepairFallbackHours,
    daysDownForBreakdown,
    daysDownForBreakdownInRange,
    drawDailyPdfExceptions,
    drawDailyPdfMetricCards,
    drawDailyPdfOperatingBasis,
    fmtNum,
    getBreakdownDowntimeColumn,
    getSiteCode,
    hasColumn,
    hasTable,
    isMonth,
    isYmd,
    kpiDaily,
    kpiRange,
    mergeAssetCostsWithRunHours,
    monthRange,
    monthStartIso,
    parseIsoDate,
    queryPeriodFleetCostTotals,
    queryPeriodRunHoursByAsset,
    rollupContractorFuelBySupplier,
    rollupFleetCostByCategory,
    safeNum,
    serviceLabelFromDailyDowntime,
    todayYmd,
  } = ctx;

  // =========================
  // DAILY PDF
  // =========================
  app.get("/daily.pdf", async (req, reply) => {
    const reportRevision = "daily-pdf-management-layout-r2026-09-29";
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    reply.header("X-IRONLOG-Report-Revision", reportRevision);

    const date = String(req.query?.date || "").trim();
    const scheduled = Number(req.query?.scheduled ?? 10);
    if (!isDate(date)) return reply.code(400).send({ error: "date (YYYY-MM-DD) required" });
    // The report is issued today for the previous completed operations day.
    const opsDay = dailyPdfOperationsDate(date);

    const logoPath = path.join(process.cwd(), "branding", "logo.png");

    const archivedClause = hasColumn("assets", "archived")
      ? "AND COALESCE(a.archived, 0) = 0"
      : "";
    const activeClause = hasColumn("assets", "active")
      ? "AND COALESCE(a.active, 1) = 1"
      : "";

    // Daily Input is captured on the report issue date for the previous operations day.
    const hours = db.prepare(`
      SELECT
        a.id AS asset_id,
        a.asset_code,
        a.asset_name,
        a.category,
        COALESCE(dh.hours_run, 0) AS hours_run,
        COALESCE(dh.scheduled_hours, 0) AS scheduled_hours,
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

    const breakdownDowntimeCol = getBreakdownDowntimeColumn();
    const hasBreakdownStatus = hasColumn("breakdowns", "status");
    const hasBreakdownEndAt = hasColumn("breakdowns", "end_at");
    const hasBreakdownStartAt = hasColumn("breakdowns", "start_at");
    const breakdownDateExpr = hasBreakdownStartAt ? "DATE(COALESCE(b.breakdown_date, b.start_at))" : "DATE(b.breakdown_date)";
    const breakdownStatusExpr = hasBreakdownStatus ? "TRIM(LOWER(COALESCE(b.status, '')))" : "''";
    const breakdownStartAtSelect = hasBreakdownStartAt ? "b.start_at" : "NULL AS start_at";
    const hasPartsOrderedDate = hasColumn("breakdowns", "parts_ordered_date");
    const hasPartsStatus = hasColumn("breakdowns", "parts_status");
    const hasPartsReceivedDate = hasColumn("breakdowns", "parts_received_date");
    const hasEtsRepairDate = hasColumn("breakdowns", "ets_repair_date");
    const breakdownParams = [date, date, date];
    if (hasBreakdownEndAt) breakdownParams.push(date);
    const breakdownAssetSeen = new Set();
    const breakdowns = db.prepare(`
      SELECT
        b.id,
        a.asset_code,
        a.asset_name,
        b.description,
        ${hasPartsOrderedDate ? "b.parts_ordered_date" : "NULL AS parts_ordered_date"},
        ${hasPartsStatus ? "b.parts_status" : "NULL AS parts_status"},
        ${hasPartsReceivedDate ? "b.parts_received_date" : "NULL AS parts_received_date"},
        ${hasEtsRepairDate ? "b.ets_repair_date" : "NULL AS ets_repair_date"},
        dh.notes AS daily_breakdown_comment,
        COALESCE(b.${breakdownDowntimeCol}, 0) AS downtime_hours,
        b.critical,
        b.breakdown_date,
        ${breakdownStartAtSelect},
        ${hasBreakdownEndAt ? "b.end_at" : "NULL AS end_at"},
        CAST(
          julianday(
            MIN(
              DATE(?),
              DATE(COALESCE(${hasBreakdownEndAt ? "b.end_at" : "NULL"}, ?))
            )
          ) - julianday(${breakdownDateExpr}) + 1
          AS INTEGER
        ) AS calendar_days_down,
        COALESCE((
          SELECT COUNT(DISTINCT l.log_date)
          FROM breakdown_downtime_logs l
          WHERE l.breakdown_id = b.id
            AND l.log_date <= ?
            AND COALESCE(l.hours_down, 0) > 0
        ), 0) AS logged_days
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      LEFT JOIN daily_hours dh ON dh.asset_id = b.asset_id AND dh.work_date = ?
      WHERE ${breakdownDateExpr} <= ?
        AND UPPER(COALESCE(b.description, '')) NOT LIKE '%WORK ORDER COMPLETED%'
        AND UPPER(COALESCE(b.description, '')) NOT LIKE 'MANAGER INSPECTION ALERT%'
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
    `).all(date, date, ...breakdownParams).map((r) => {
      const daysDown = daysDownForBreakdown(r, date);
      return {
        ...r,
        critical: Boolean(r.critical),
        days_down: daysDown,
        downtime_hours: Number(r.downtime_hours || 0),
      };
    }).sort((a, b) => {
      const bDate = String(b.breakdown_date || b.start_at || "");
      const aDate = String(a.breakdown_date || a.start_at || "");
      return bDate.localeCompare(aDate) || Number(b.id || 0) - Number(a.id || 0);
    }).filter((row) => {
      const key = String(row.asset_code || row.asset_name || row.id || "");
      if (breakdownAssetSeen.has(key)) return false;
      breakdownAssetSeen.add(key);
      return true;
    }).sort((a, b) => Number(b.days_down || 0) - Number(a.days_down || 0));

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
        AND REPLACE(TRIM(LOWER(COALESCE(wn.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
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
        a.asset_name,
        w.source,
        w.status,
        w.assigned_artisan_name,
        w.repair_progress,
        ${hasPartsStatus ? "b.parts_status" : "NULL AS parts_status"},
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

    const hoursPdf = hours.slice(0, 500);
    const breakdownsPdf = breakdowns.slice(0, 40);
    const openWOsPdf = openWOs.slice(0, 40);

    // Daily PDF is opened "today" for yesterday's ops — short BDs, fuel, and
    // per-asset availability downtime all use the previous calendar day.
    // Availability must use the same operations day as the downtime, fuel and
    // pre-start sections. Using the report issue date here could falsely show 100%.
    const scheduledFallback = Math.max(0, Number(scheduled || 0));
    const dailyDowntimeLogs = hasTable("breakdown_downtime_logs")
      ? db.prepare(`
          SELECT
            l.id,
            b.asset_id,
            a.asset_code,
            a.asset_name,
            b.description,
            b.component,
            b.critical,
            CASE
              WHEN ${breakdownStatusExpr} IN ('open', 'in_progress')
                AND NOT EXISTS (
                  SELECT 1
                  FROM work_orders wd
                  WHERE wd.source = 'breakdown'
                    AND wd.reference_id = b.id
                    AND REPLACE(TRIM(LOWER(COALESCE(wd.status, ''))), ' ', '_')
                      IN ('completed', 'approved', 'closed')
                )
              THEN 1 ELSE 0
            END AS effective_active,
            l.hours_down,
            l.notes,
            l.log_date
          FROM breakdown_downtime_logs l
          JOIN breakdowns b ON b.id = l.breakdown_id
          JOIN assets a ON a.id = b.asset_id
          WHERE l.log_date = ?
          ORDER BY l.hours_down DESC, a.asset_code ASC, l.id ASC
        `).all(opsDay)
          .filter((r) => Number(r.hours_down || 0) > 0 || Number(r.effective_active || 0) === 1)
          .slice(0, 60)
      : [];

    // A breakdown may begin and be repaired during an operations day, while the
    // technician completes/closes its work order the following morning. When
    // no explicit downtime log exists for that incident/day, carry the recorded
    // repair hours back to the breakdown day. Explicit daily logs remain the
    // source of truth and are never replaced.
    const hasBreakdownRepairLabor = hasTable("breakdown_repair_labor");
    const repairHoursExpr = hasBreakdownRepairLabor
      ? "COALESCE(NULLIF(brl.labor_hours, 0), w.labor_hours, 0)"
      : "COALESCE(w.labor_hours, 0)";
    const completedBreakdownRepairCandidates = db.prepare(`
      SELECT
        b.id AS breakdown_id,
        b.asset_id,
        a.asset_code,
        a.asset_name,
        b.description,
        b.component,
        MAX(${repairHoursExpr}) AS repair_hours
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      JOIN work_orders w
        ON LOWER(TRIM(COALESCE(w.source, ''))) = 'breakdown'
        AND COALESCE(w.reference_id, -1) = b.id
      ${hasBreakdownRepairLabor ? "LEFT JOIN breakdown_repair_labor brl ON brl.breakdown_id = b.id" : ""}
      WHERE ${breakdownDateExpr} = ?
        AND REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_')
          IN ('completed', 'approved', 'closed')
        AND COALESCE(${repairHoursExpr}, 0) > 0
        AND NOT EXISTS (
          SELECT 1
          FROM breakdown_downtime_logs l
          WHERE l.breakdown_id = b.id
            AND l.log_date = ?
            AND COALESCE(l.hours_down, 0) > 0
        )
      GROUP BY b.id, b.asset_id, a.asset_code, a.asset_name, b.description, b.component
      ORDER BY b.id ASC
    `).all(opsDay, opsDay);

    const scheduledByAssetId = new Map(
      hours.map((r) => {
        const rowScheduled = Number(r.scheduled_hours || 0);
        return [
          Number(r.asset_id || 0),
          rowScheduled > 0 ? rowScheduled : scheduledFallback,
        ];
      }),
    );
    const loggedDowntimeByAssetId = new Map();
    for (const row of dailyDowntimeLogs) {
      const assetId = Number(row.asset_id || 0);
      if (!assetId) continue;
      loggedDowntimeByAssetId.set(
        assetId,
        Math.max(0, Number(loggedDowntimeByAssetId.get(assetId) || 0))
          + Math.max(0, Number(row.hours_down || 0)),
      );
    }
    const repairDowntimeAllocatedByAssetId = new Map();
    const completedBreakdownRepairDowntime = completedBreakdownRepairCandidates
      .map((row) => {
        const assetId = Number(row.asset_id || 0);
        const dayCap = Math.max(0, Number(scheduledByAssetId.get(assetId) || scheduledFallback));
        const alreadyLogged = Math.max(0, Number(loggedDowntimeByAssetId.get(assetId) || 0));
        const alreadyAllocated = Math.max(0, Number(repairDowntimeAllocatedByAssetId.get(assetId) || 0));
        const repairHours = Math.max(0, Number(row.repair_hours || 0));
        // A shift is 06:00-17:00 (11 hours) unless its scheduled-hours value
        // says otherwise. Work-order repair time only fills a blank Daily Log;
        // it must never inflate an asset's explicit recorded downtime.
        const hoursDown = dailyPdfRepairFallbackHours({
          dayCap,
          loggedHours: alreadyLogged,
          allocatedHours: alreadyAllocated,
          repairHours,
        });
        if (hoursDown > 0) {
          repairDowntimeAllocatedByAssetId.set(assetId, alreadyAllocated + hoursDown);
        }
        return {
          ...row,
          repair_hours: repairHours,
          hours_down: Number(hoursDown.toFixed(2)),
        };
      })
      .filter((row) => row.hours_down > 0);
    const completedRepairDowntimeHours = completedBreakdownRepairDowntime.reduce(
      (sum, row) => sum + Number(row.hours_down || 0),
      0,
    );
    const kpi = kpiDaily(opsDay, scheduled, date, {
      additionalDowntimeHours: completedRepairDowntimeHours,
    });
    const dailyPlannedMaintenance = listPlannedMaintenanceForDate(db, opsDay).slice(0, 40);
    const offsiteSiteRaw = String(req.query?.site_code || getSiteCode(req) || "main").trim().toLowerCase() || "main";
    const offsiteSiteAliases =
      offsiteSiteRaw === "main" || offsiteSiteRaw === "default"
        ? ["main", "default"]
        : [offsiteSiteRaw];
    const offsiteRepairsPdf = hasTable("breakdown_offsite_repairs")
      ? db.prepare(`
          SELECT
            r.id,
            a.asset_code,
            a.asset_name,
            r.repair_status,
            r.sent_date,
            r.expected_return_date,
            r.actual_return_date,
            r.vendor,
            r.notes
            ,r.repair_reason
            ,r.approval_status
            ,r.quote_number
            ,r.responsible_person
            ,r.estimated_cost
            ,r.actual_cost
          FROM breakdown_offsite_repairs r
          JOIN assets a ON a.id = r.asset_id
          WHERE LOWER(TRIM(COALESCE(r.site_code, 'main'))) IN (${offsiteSiteAliases.map(() => "?").join(", ")})
            AND r.sent_date <= ?
            AND (
              r.actual_return_date IS NULL
              OR TRIM(COALESCE(r.actual_return_date, '')) = ''
              OR r.actual_return_date >= ?
            )
          ORDER BY r.sent_date ASC, r.id ASC
          LIMIT 40
        `).all(...offsiteSiteAliases, date, date)
      : [];

    // Per-asset downtime for availability (same day as short breakdowns / fuel).
    const downtimeByAssetId = new Map();
    // Keep actual recorded loss separate from the full-shift fallback for an
    // open incident. A saved Production row must only use actual loss.
    const recordedDowntimeByAssetId = new Map();
    try {
      const downRows = db.prepare(`
        SELECT b.asset_id, COALESCE(SUM(l.hours_down), 0) AS hours_down
        FROM breakdown_downtime_logs l
        JOIN breakdowns b ON b.id = l.breakdown_id
        WHERE l.log_date = ?
        GROUP BY b.asset_id
      `).all(opsDay);
      for (const r of downRows) {
        const assetId = Number(r.asset_id || 0);
        const hoursDown = Math.max(0, Number(r.hours_down || 0));
        downtimeByAssetId.set(assetId, hoursDown);
        recordedDowntimeByAssetId.set(assetId, hoursDown);
      }
      for (const r of completedBreakdownRepairDowntime) {
        const assetId = Number(r.asset_id || 0);
        if (!assetId) continue;
        const repairHours = Math.max(0, Number(r.hours_down || 0));
        const totalDown = Math.max(0, Number(downtimeByAssetId.get(assetId) || 0)) + repairHours;
        const recordedDown = Math.max(0, Number(recordedDowntimeByAssetId.get(assetId) || 0)) + repairHours;
        downtimeByAssetId.set(assetId, totalDown);
        recordedDowntimeByAssetId.set(assetId, recordedDown);
      }

      // Only infer a full shift for a brand-new open incident with neither a
      // Daily Log production entry nor an explicit downtime log. A production
      // entry wins over an open WO; actual loss must be entered as downtime.
      const activeDownAssets = db.prepare(`
        SELECT DISTINCT b.asset_id
        FROM breakdowns b
        WHERE ${breakdownDateExpr} <= ?
          AND ${breakdownStatusExpr} IN ('open', 'in_progress')
          AND ${hasBreakdownEndAt ? "(b.end_at IS NULL OR DATE(b.end_at) >= ?)" : "1 = 1"}
          AND NOT EXISTS (
            SELECT 1
            FROM work_orders wa
            WHERE wa.source = 'breakdown'
              AND wa.reference_id = b.id
              AND REPLACE(TRIM(LOWER(COALESCE(wa.status, ''))), ' ', '_')
                IN ('completed', 'approved', 'closed')
          )
          AND NOT EXISTS (
            SELECT 1
            FROM breakdown_downtime_logs l
            WHERE l.breakdown_id = b.id
              AND l.log_date <= ?
          )
          AND NOT EXISTS (
            SELECT 1
            FROM daily_hours dh
            WHERE dh.asset_id = b.asset_id
              AND dh.work_date IN (?, ?)
              AND COALESCE(dh.is_used, 0) = 1
          )
      `).all(...(hasBreakdownEndAt
        ? [opsDay, opsDay, opsDay, date, opsDay]
        : [opsDay, opsDay, date, opsDay]));
      for (const r of activeDownAssets) {
        const assetId = Number(r.asset_id || 0);
        if (assetId > 0 && Number(downtimeByAssetId.get(assetId) || 0) <= 0) {
          downtimeByAssetId.set(assetId, scheduledFallback);
        }
      }
    } catch { /* table may be missing */ }

    // Fuel for previous calendar day (ops day), shown per equipment on the hours table.
    const fuelByAssetId = new Map();
    try {
      const fuelDayRows = db.prepare(`
        SELECT a.id AS asset_id, COALESCE(SUM(fl.liters), 0) AS liters
        FROM fuel_logs fl
        JOIN assets a ON a.id = fl.asset_id
        WHERE fl.log_date = ?
        GROUP BY a.id
      `).all(opsDay);
      for (const r of fuelDayRows) {
        fuelByAssetId.set(Number(r.asset_id), Math.max(0, Number(r.liters || 0)));
      }
    } catch { /* ignore */ }

    const prestartAssetIds = new Set(
      listDailyPrestarts(db, opsDay).rows.map((r) => Number(r.asset_id || 0)).filter((id) => id > 0),
    );

    const hoursPdfEnriched = hoursPdf.map((r) => {
      const assetId = Number(r.asset_id || 0);
      const rowScheduled = Number(r.scheduled_hours);
      const sched = Math.max(
        0,
        Number.isFinite(rowScheduled) && rowScheduled > 0 ? rowScheduled : scheduledFallback,
      );
      const run = Math.max(0, Number(r.hours_run || 0));
      const runEff = sched > 0 ? Math.min(run, sched) : run;
      // A Production entry is the operator's statement that the asset was
      // available. Never mix the open-incident full-shift assumption into that
      // row: retain only an entered downtime log or recorded repair hours.
      const downRaw = dailyPdfDowntimeHours({
        hasDailyEntry: r.has_daily_entry,
        isUsed: r.is_used,
        recordedHours: recordedDowntimeByAssetId.get(assetId),
        totalHours: downtimeByAssetId.get(assetId),
      });
      const down = sched > 0 ? Math.min(downRaw, sched) : downRaw;
      const prestartDone = prestartAssetIds.has(assetId);
      const inspection = prestartDone && sched > down
        ? Math.min(PRESTART_DEDUCTION_HOURS, sched - down)
        : 0;
      const available = Math.max(0, sched - down - inspection);
      const availPct = sched > 0 ? (available / sched) * 100 : null;
      const utilPct = sched > 0 ? (runEff / sched) * 100 : null;
      const fuelLiters = fuelByAssetId.get(assetId);
      return {
        ...r,
        scheduled_hours_eff: sched,
        available_hours: available,
        downtime_hours: down,
        prestart_done: prestartDone,
        inspection_hours: inspection,
        availability_pct: availPct,
        utilization_pct: utilPct,
        fuel_liters: fuelLiters == null ? null : fuelLiters,
      };
    });

    const speedAlertKmh = getCartrackSpeedAlertKmh();
    let cartrackSpeeding = null;
    if (hasTable("cartrack_events")) {
      try {
        const events = listCartrackEventsFromDb({
          startDate: opsDay,
          endDate: opsDay,
          minSpeedKmh: speedAlertKmh,
          limit: 500,
        });
        cartrackSpeeding = summarizeSpeedingEvents(opsDay, events, speedAlertKmh);
      } catch {
        cartrackSpeeding = null;
      }
    }

    const offsiteStatusLabel = (status) => {
      const key = String(status || "").trim().toLowerCase();
      const labels = {
        sent_offsite: "Sent offsite",
        in_repair: "In repair",
        waiting_parts: "Waiting parts",
        ready_return: "Ready for return",
        returned: "Returned",
        diagnosis: "Diagnosis",
      };
      return labels[key] || key || "-";
    };
    const daysBetweenYmd = (from, to) => {
      const start = String(from || "").trim();
      const end = String(to || "").trim();
      if (!start || !end) return null;
      const startTime = Date.parse(`${start}T00:00:00Z`);
      const endTime = Date.parse(`${end}T00:00:00Z`);
      if (!Number.isFinite(startTime) || !Number.isFinite(endTime)) return null;
      return Math.round((endTime - startTime) / 86400000);
    };

    const offsiteReportRows = offsiteRepairsPdf.map((row) => {
      const sent = String(row.sent_date || "").trim();
      const expected = String(row.expected_return_date || "").trim();
      const actual = String(row.actual_return_date || "").trim();
      const elapsed = daysBetweenYmd(sent, actual || date);
      const dueDelta = daysBetweenYmd(date, expected);
      const tracking = actual
        ? `Returned ${actual}`
        : !expected
          ? "Return date required"
          : dueDelta < 0
            ? `OVERDUE ${Math.abs(dueDelta)} day(s)`
            : dueDelta === 0
              ? "Due today"
              : `Due in ${dueDelta} day(s)`;
      return {
        asset: row.asset_code,
        equipment: compactCell(row.asset_name ?? "", 48),
        status: offsiteStatusLabel(row.repair_status),
        owner: compactCell([row.responsible_person, row.vendor].filter(Boolean).join(" / "), 44) || "-",
        sent: sent || "-",
        expected: expected || "-",
        elapsed: elapsed == null ? "-" : String(elapsed),
        tracking,
        approval: compactCell(String(row.approval_status || "not required").replaceAll("_", " "), 28),
        notes: compactCell([row.repair_reason, row.notes].filter(Boolean).join(" / "), 120) || "-",
      };
    });

    const breakdownReportRows = breakdownsPdf.map((row) => ({
      asset: row.asset_code,
      equipment: compactCell(row.asset_name ?? "", 40),
      date_down: parseIsoDate(row.breakdown_date) || parseIsoDate(row.start_at) || "-",
      days: fmtNum(row.days_down || 0, 0),
      parts_ordered: row.parts_ordered_date || "-",
      parts_status: compactCell(row.parts_status ?? "", 36) || "-",
      received: row.parts_received_date || "-",
      ets: row.ets_repair_date || "-",
      desc: compactCell(row.description ?? "", 70) || "-",
    }));

    const loggedDowntimeAssets = new Set();
    const dailyDowntimeRows = dailyDowntimeLogs.map((row) => {
      loggedDowntimeAssets.add(String(row.asset_code || ""));
      const serviceLabel = serviceLabelFromDailyDowntime(row.component, row.notes, row.description);
      const isMaintenance = Boolean(serviceLabel);
      return {
        asset: row.asset_code,
        equipment: compactCell(row.asset_name ?? "", 48),
        type: isMaintenance ? "Maintenance" : "Breakdown",
        hrs: fmtNum(Math.max(0, Number(row.hours_down || 0)), 1),
        area: isMaintenance ? serviceLabel : (compactCell(row.component ?? "", 42) || "-"),
        detail: isMaintenance
          ? serviceLabel
          : compactCell(String(row.notes || row.description || "").replace(/^(?:Auto from Daily Input \(DOWN\)|Daily Log breakdown)\s*[-]?\s*/i, ""), 220),
      };
    });
    for (const row of completedBreakdownRepairDowntime) {
      dailyDowntimeRows.push({
        asset: row.asset_code,
        equipment: compactCell(row.asset_name ?? "", 48),
        type: "Breakdown",
        hrs: fmtNum(row.hours_down || 0, 1),
        area: compactCell(row.component ?? "", 42) || "-",
        detail: compactCell(
          `${String(row.description || "Breakdown repair").trim()} / ${fmtNum(row.repair_hours || 0, 1)} completed repair hr`,
          220,
        ),
      });
    }
    for (const row of dailyPlannedMaintenance) {
      if (loggedDowntimeAssets.has(String(row.asset_code || ""))) continue;
      dailyDowntimeRows.push({
        asset: row.asset_code,
        equipment: compactCell(row.asset_name ?? "", 48),
        type: "Maintenance",
        hrs: "-",
        area: compactCell(row.service_name ?? "", 42) || "Service",
        detail: compactCell(row.description ?? "", 220),
      });
    }

    const openWorkOrderRows = openWOsPdf.map((row) => ({
      wo: String(row.id),
      asset: row.asset_code,
      equipment: compactCell(row.asset_name ?? "", 48),
      source: ({
        breakdown: "Breakdown",
        service: "Maintenance",
        manager_inspection: "Inspection",
        inspection: "Inspection",
      })[String(row.source || "").toLowerCase()] || compactCell(row.source ?? "", 20),
      status: compactCell(String(row.status ?? "").replace(/_/g, " "), 16),
      parts: compactCell(row.parts_status ?? "", 26) || "-",
      tech: compactCell(row.assigned_artisan_name ?? "", 28),
      progress: compactCell(row.repair_progress ?? "", 260) || "No progress update",
    }));

    const fleetReadingRows = hoursPdfEnriched.map((row) => {
      const noEntry = !row.has_daily_entry;
      const formatHours = (value) =>
        value == null || value === "" || !Number.isFinite(Number(value)) ? "-" : fmtNum(value, 1);
      const formatPercent = (value) =>
        value == null || !Number.isFinite(Number(value)) ? "-" : `${fmtNum(value, 1)}%`;
      const isDistanceAsset = /(?:\bldv\b|light vehicle|vehicle)/i.test(`${row.category || ""} ${row.asset_name || ""}`);
      return {
        asset: row.asset_code,
        type: compactCell(row.category ?? "", 12),
        equipment: compactCell(row.asset_name ?? "", 42),
        open: noEntry ? "-" : formatHours(row.opening_hours),
        close: noEntry ? "-" : formatHours(row.closing_hours),
        run: `${fmtNum(row.hours_run, 1)}${isDistanceAsset ? " km" : " h"}`,
        prestart: row.prestart_done ? "Done" : "-",
        avail: formatPercent(row.availability_pct),
        util: formatPercent(row.utilization_pct),
        fuel: row.fuel_liters == null ? "-" : `${fmtNum(row.fuel_liters, 1)} L`,
      };
    });

    const downtimeActionRows = dailyDowntimeRows
      .filter((row) => row.type === "Breakdown" && Number(row.hrs || 0) > 0)
      .sort((a, b) => Number(b.hrs || 0) - Number(a.hrs || 0))
      .slice(0, 3)
      .map((row) => ({
        title: `${row.asset} | ${row.equipment}`,
        detail: `${row.area}: ${row.detail || "Breakdown downtime recorded"}`,
        value: `${row.hrs} h down`,
        alert: true,
      }));
    const offsiteActionRows = offsiteReportRows.slice(0, 2).map((row) => ({
      title: `${row.asset} | ${row.equipment}`,
      detail: `Offsite repair - ${row.status}. ${row.notes}`,
      value: row.elapsed === "-" ? row.tracking : `${row.elapsed} day(s) out`,
      alert: /OVERDUE|required/i.test(row.tracking),
    }));
    const managementActionRows = [...downtimeActionRows, ...offsiteActionRows].slice(0, 5);
    let dailyPdfSiteName = "Quionga, Mozambique";
    try {
      dailyPdfSiteName = String(getPdfReportBranding(db).site_name || dailyPdfSiteName).trim();
    } catch {
      // The report remains available when branding settings have not been initialized.
    }

    const dailyPdfPageLabels = ["MANAGEMENT OVERVIEW"];
    const pdf = await buildPdfBuffer(
      (doc) => {
        {
          const dailyPageLabels = dailyPdfPageLabels;
          let activePageLabel = "MANAGEMENT OVERVIEW";
          const managementTableStyle = {
            compact: true,
            headerColor: "#143c50",
            headerLineColor: "#0f7182",
            zebraColor: "#f2f6f7",
            textColor: "#143c50",
          };
          doc.on("pageAdded", () => dailyPageLabels.push(activePageLabel));

          dailyPdfManagementSection(doc, "Daily performance");
          drawDailyPdfMetricCards(doc, [
            {
              label: "Availability",
              value: kpi.availability == null ? "-" : `${fmtNum(kpi.availability, 1)}%`,
              detail: "scheduled time available",
              alert: kpi.availability != null && Number(kpi.availability) < 80,
            },
            {
              label: "Utilization",
              value: kpi.utilization == null ? "-" : `${fmtNum(kpi.utilization, 1)}%`,
              detail: "run hours / schedule",
            },
            {
              label: "Run hours",
              value: `${fmtNum(kpi.run_hours, 1)} h`,
              detail: `${fmtNum(kpi.used_assets, 0)} assets used`,
            },
            {
              label: "Downtime",
              value: `${fmtNum(kpi.downtime_hours, 1)} h`,
              detail: "repair and maintenance loss",
              alert: Number(kpi.downtime_hours || 0) > 0,
            },
          ]);

          dailyPdfManagementSection(doc, "Operating basis");
          drawDailyPdfOperatingBasis(doc, [
            { label: "Operating day", value: dailyPdfLongDate(opsDay) },
            { label: "Shift pattern", value: "06:00 - 17:00" },
            { label: "Scheduled capacity", value: `${fmtNum(scheduled, 1)} h per asset` },
            {
              label: "Pre-start checks",
              value: `${fmtNum(kpi.prestart_count, 0)} completed (${fmtNum(kpi.prestart_hours, 2)} h)`,
            },
          ]);

          dailyPdfManagementSection(doc, "Exceptions and actions");
          drawDailyPdfExceptions(doc, managementActionRows);

          activePageLabel = "FLEET DETAIL";
          doc.addPage();
          dailyPdfManagementSection(doc, "Fleet readings");
          doc.font("Helvetica").fontSize(8.5).fillColor("#526a7c").text(
            `Production hours captured on ${date} apply to ${opsDay}. Fuel, availability and downtime are shown for ${opsDay}.`,
          );
          doc.moveDown(0.45);
          if (fleetReadingRows.length) {
            table(
              doc,
              [
                { key: "asset", label: "Asset", width: 0.09 },
                { key: "type", label: "Type", width: 0.09 },
                { key: "equipment", label: "Equipment", width: 0.18 },
                { key: "open", label: "Open", width: 0.08, align: "right" },
                { key: "close", label: "Close", width: 0.08, align: "right" },
                { key: "run", label: "Move", width: 0.09, align: "right" },
                { key: "prestart", label: "Pre-start", width: 0.09, align: "center" },
                { key: "avail", label: "Avail", width: 0.08, align: "right" },
                { key: "util", label: "Util", width: 0.08, align: "right" },
                { key: "fuel", label: "Fuel", width: 0.14, align: "right" },
              ],
              fleetReadingRows,
              { ...managementTableStyle, fontSize: 7.5, headerFontSize: 7.5, rowPadY: 3, headerPadY: 4 },
            );
          } else {
            drawDailyPdfExceptions(doc, [], "No fleet readings were captured for this operating day.");
          }

          activePageLabel = "MAINTENANCE & INCIDENTS";
          doc.addPage();
          dailyPdfManagementSection(doc, "Maintenance and incidents");

          dailyPdfManagementSection(doc, "Offsite repair tracking");
          if (offsiteReportRows.length) {
            table(
              doc,
              [
                { key: "asset", label: "Plant #", width: 0.07 },
                { key: "equipment", label: "Equipment", width: 0.13 },
                { key: "status", label: "Repair status", width: 0.10 },
                { key: "owner", label: "Owner / repairer", width: 0.12 },
                { key: "sent", label: "Sent", width: 0.07 },
                { key: "expected", label: "Expected", width: 0.08 },
                { key: "elapsed", label: "Days out", width: 0.06, align: "right" },
                { key: "tracking", label: "Return tracking", width: 0.11 },
                { key: "approval", label: "Approval", width: 0.10 },
                { key: "notes", label: "Reason / next action", width: 0.16 },
              ],
              offsiteReportRows,
              managementTableStyle,
            );
          } else {
            drawDailyPdfExceptions(doc, [], "No assets are currently tracked as offsite for repair.");
          }

          dailyPdfManagementSection(doc, "Breakdown incidents");
          if (breakdownReportRows.length) {
            table(
              doc,
              [
                { key: "asset", label: "Plant #", width: 0.09 },
                { key: "equipment", label: "Equipment", width: 0.15 },
                { key: "date_down", label: "Date down", width: 0.11 },
                { key: "days", label: "Total days down", width: 0.10, align: "right" },
                { key: "parts_ordered", label: "Parts ordered", width: 0.12 },
                { key: "parts_status", label: "Parts status", width: 0.12 },
                { key: "received", label: "Received", width: 0.11 },
                { key: "ets", label: "ETS repair", width: 0.10 },
                { key: "desc", label: "Fault / action", width: 0.10 },
              ],
              breakdownReportRows,
              managementTableStyle,
            );
          } else {
            drawDailyPdfExceptions(doc, [], "No active breakdown incidents were recorded.");
          }

          dailyPdfManagementSection(doc, `Daily downtime - ${opsDay}`);
          if (dailyDowntimeRows.length) {
            table(
              doc,
              [
                { key: "asset", label: "Plant #", width: 0.10 },
                { key: "equipment", label: "Equipment", width: 0.20 },
                { key: "type", label: "Downtime type", width: 0.14 },
                { key: "hrs", label: "Hours down", width: 0.11, align: "right" },
                { key: "area", label: "Component / service", width: 0.17 },
                { key: "detail", label: "Reason / work completed", width: 0.28 },
              ],
              dailyDowntimeRows,
              managementTableStyle,
            );
          } else {
            drawDailyPdfExceptions(doc, [], "No repair or maintenance downtime was recorded.");
          }

          dailyPdfManagementSection(doc, "Open work orders");
          if (openWorkOrderRows.length) {
            table(
              doc,
              [
                { key: "wo", label: "WO#", width: 0.07, align: "right" },
                { key: "asset", label: "Plant #", width: 0.09 },
                { key: "equipment", label: "Equipment", width: 0.16 },
                { key: "source", label: "Source", width: 0.10 },
                { key: "status", label: "Status", width: 0.09 },
                { key: "parts", label: "Parts status", width: 0.12 },
                { key: "tech", label: "Technician", width: 0.11 },
                { key: "progress", label: "Progress / next action", width: 0.26 },
              ],
              openWorkOrderRows,
              managementTableStyle,
            );
          } else {
            drawDailyPdfExceptions(doc, [], "No open work orders require action.");
          }

          if (cartrackSpeeding?.total_speeding_events) {
            activePageLabel = "FLEET TRACKING";
            doc.addPage();
            dailyPdfManagementSection(doc, `Fleet tracking - speeding above ${speedAlertKmh} km/h`);
            doc.font("Helvetica").fontSize(9).fillColor("#526a7c").text(
              `${cartrackSpeeding.total_speeding_events} event(s) across ${cartrackSpeeding.vehicles_with_speeding} vehicle(s).`,
            );
            doc.moveDown(0.4);
            const speedPdf = buildSpeedingReportPdfContent(cartrackSpeeding);
            table(
              doc,
              speedPdf.columns,
              speedPdf.rows.slice(0, 25).map((row) => ({
                time: compactCell(row.time ?? "", 16),
                vehicle: compactCell(String(row.vehicle ?? "").replace(/\n/g, " / "), 24),
                speed: row.speed ?? "-",
                limit: row.limit ?? "-",
                type: compactCell(row.type ?? "", 80),
              })),
              managementTableStyle,
            );
          }
        }
        return;

        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "KPIs");
        kvGrid(doc, [
          { k: "Report date", v: date },
          { k: "Operations date", v: opsDay },
          { k: "Scheduled hours / asset", v: fmtNum(scheduled, 0) },
          { k: "Used assets", v: fmtNum(kpi.used_assets, 0) },
          { k: "Available hours", v: fmtNum(kpi.available_hours, 0) },
          { k: "Run hours", v: fmtNum(kpi.run_hours, 1) },
          { k: "Downtime hours", v: fmtNum(kpi.downtime_hours, 1) },
          ...(kpi.prestart_count > 0
            ? [{
                k: `Pre-start checks (${fmtNum(kpi.prestart_deduction_hours_per_check, 2)} hr each)`,
                v: `${fmtNum(kpi.prestart_count, 0)} check(s) · ${fmtNum(kpi.prestart_hours, 2)} hr deducted`,
              }]
            : []),
          { k: "Availability %", v: kpi.availability == null ? "N/A" : `${fmtNum(kpi.availability, 2)}%` },
          { k: "Utilization %", v: kpi.utilization == null ? "N/A" : `${fmtNum(kpi.utilization, 2)}%` },
          ...(cartrackSpeeding
            ? [{
                k: `Speeding above ${speedAlertKmh} km/h`,
                v: cartrackSpeeding.total_speeding_events
                  ? `${cartrackSpeeding.total_speeding_events} event(s) · ${cartrackSpeeding.vehicles_with_speeding} vehicle(s)`
                  : "None",
              }]
            : []),
        ], 2);

        sectionTitle(doc, "Equipment hours and fuel");
        doc.fontSize(9).fillColor("#64748b");
        doc.text(
          `Production hours captured on ${date} apply to ${opsDay}; fuel, availability and downtime are also for ${opsDay}.`,
        );
        doc.moveDown(0.35);
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.09 },
            { key: "type", label: "Type", width: 0.09 },
            { key: "name", label: "Name", width: 0.14 },
            { key: "open", label: "Open", width: 0.07, align: "right" },
            { key: "close", label: "Close", width: 0.07, align: "right" },
            { key: "hours", label: "Run Hrs", width: 0.07, align: "right" },
            { key: "prestart", label: "Pre-start", width: 0.11, align: "center" },
            { key: "avail", label: "Avail %", width: 0.09, align: "right" },
            { key: "util", label: "Util %", width: 0.09, align: "right" },
            { key: "fuel", label: `Fuel L (${opsDay.slice(5)})`, width: 0.18, align: "right" },
          ],
          hoursPdfEnriched.map((r) => {
            const noEntry = !r.has_daily_entry;
            const fmtHm = (v) =>
              v == null || v === "" || !Number.isFinite(Number(v)) ? "—" : fmtNum(v, 1);
            const fmtPct = (v) => (v == null || !Number.isFinite(Number(v)) ? "—" : `${fmtNum(v, 1)}%`);
            const fuelLiters = r.fuel_liters;
            return {
              asset: r.asset_code,
              type: compactCell(r.category ?? "", 12),
              name: r.asset_name ?? "",
              open: noEntry ? "—" : fmtHm(r.opening_hours),
              close: noEntry ? "—" : fmtHm(r.closing_hours),
              hours: fmtNum(r.hours_run, 1),
              prestart: r.prestart_done ? `Done (-${fmtNum(r.inspection_hours, 2)}h)` : "—",
              avail: fmtPct(r.availability_pct),
              util: fmtPct(r.utilization_pct),
              fuel: fuelLiters == null ? "—" : fmtNum(fuelLiters, 1),
            };
          })
        );

        {
          const statusLabel = (s) => {
            const k = String(s || "").trim().toLowerCase();
            if (k === "sent_offsite") return "Sent offsite";
            if (k === "in_repair") return "In repair";
            if (k === "waiting_parts") return "Waiting parts";
            if (k === "ready_return") return "Ready return";
            if (k === "returned") return "Returned";
            if (k === "diagnosis") return "Diagnosis";
            return k || "—";
          };
          const daysBetweenYmd = (from, to) => {
            const a = String(from || "").trim();
            const b = String(to || "").trim();
            if (!a || !b) return null;
            const t0 = Date.parse(`${a}T00:00:00Z`);
            const t1 = Date.parse(`${b}T00:00:00Z`);
            if (!Number.isFinite(t0) || !Number.isFinite(t1)) return null;
            return Math.round((t1 - t0) / 86400000);
          };
          sectionTitle(doc, "Offsite repair tracking");
          if (!offsiteRepairsPdf.length) {
            doc.font("Helvetica").fontSize(10).fillColor("#555555")
              .text("No assets currently tracked as offsite for repairs.", doc.page.margins.left, doc.y, {
                width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
              });
            doc.moveDown(0.6);
          } else {
            table(
              doc,
              [
                { key: "asset", label: "Plant #", width: 0.07 },
                { key: "equipment", label: "Equipment", width: 0.13 },
                { key: "status", label: "Repair status", width: 0.10 },
                { key: "owner", label: "Owner / repairer", width: 0.12 },
                { key: "sent", label: "Sent", width: 0.07 },
                { key: "expected", label: "Expected", width: 0.08 },
                { key: "elapsed", label: "Days out", width: 0.06, align: "right" },
                { key: "tracking", label: "Return tracking", width: 0.11 },
                { key: "approval", label: "Approval", width: 0.10 },
                { key: "notes", label: "Reason / next action", width: 0.16 },
              ],
              offsiteRepairsPdf.map((r) => {
                const sent = String(r.sent_date || "").trim();
                const expected = String(r.expected_return_date || "").trim();
                const actual = String(r.actual_return_date || "").trim();
                const elapsed = daysBetweenYmd(sent, actual || date);
                const dueDelta = daysBetweenYmd(date, expected);
                const tracking = actual
                  ? `Returned ${actual}`
                  : !expected
                    ? "Return date required"
                    : dueDelta < 0
                      ? `OVERDUE ${Math.abs(dueDelta)} day(s)`
                      : dueDelta === 0
                        ? "Due today"
                        : `Due in ${dueDelta} day(s)`;
                return {
                  asset: r.asset_code,
                  equipment: compactCell(r.asset_name ?? "", 48),
                  status: statusLabel(r.repair_status),
                  owner: compactCell([r.responsible_person, r.vendor].filter(Boolean).join(" / "), 44) || "—",
                  sent: sent || "—",
                  expected: expected || "—",
                  elapsed: elapsed == null ? "—" : String(elapsed),
                  tracking,
                  approval: compactCell(String(r.approval_status || "not required").replaceAll("_", " "), 28),
                  notes: compactCell([r.repair_reason, r.notes].filter(Boolean).join(" · "), 120) || "—",
                };
              })
            );
          }
        }

        ensurePageSpace(doc, 245);
        sectionTitle(doc, "Breakdown incidents");
        table(
          doc,
          [
            { key: "asset", label: "Plant #", width: 0.09 },
            { key: "equipment", label: "Equipment", width: 0.15 },
            { key: "date_down", label: "Date down", width: 0.11 },
            { key: "days", label: "Total days down", width: 0.10, align: "right" },
            { key: "parts_ordered", label: "Date parts ordered", width: 0.12 },
            { key: "parts_status", label: "Status of parts", width: 0.12 },
            { key: "received", label: "Received date", width: 0.11 },
            { key: "ets", label: "ETS Repair", width: 0.10 },
            { key: "desc", label: "Fault / action", width: 0.10 },
          ],
          breakdownsPdf.map((r) => {
            // Selected Date down = breakdown_date (operator-chosen). Avoid start_at create-time stamp.
            const dateDown =
              parseIsoDate(r.breakdown_date) || parseIsoDate(r.start_at) || "—";
            return {
              asset: r.asset_code,
              equipment: compactCell(r.asset_name ?? "", 40),
              date_down: dateDown,
              days: fmtNum(r.days_down || 0, 0),
              parts_ordered: r.parts_ordered_date || "—",
              parts_status: compactCell(r.parts_status ?? "", 36) || "—",
              received: r.parts_received_date || "—",
              ets: r.ets_repair_date || "—",
              desc: compactCell(r.description ?? "", 70),
            };
          })
        );

        // A Daily Log incident can be either an unplanned breakdown or a
        // planned service. Numeric service intervals identify the latter even
        // when the entry originated from the quick breakdown workflow.
        const legacyLoggedDowntimeAssets = new Set();
        const legacyDailyDowntimeRows = dailyDowntimeLogs.map((r) => {
          legacyLoggedDowntimeAssets.add(String(r.asset_code || ""));
          const serviceLabel = serviceLabelFromDailyDowntime(r.component, r.notes, r.description);
          const isMaintenance = Boolean(serviceLabel);
          return {
            asset: r.asset_code,
            equipment: compactCell(r.asset_name ?? "", 48),
            type: isMaintenance ? "Maintenance" : "Breakdown",
            hrs: fmtNum(Math.max(0, Number(r.hours_down || 0)), 1),
            area: isMaintenance
              ? serviceLabel
              : (compactCell(r.component ?? "", 42) || "—"),
            // For a service, the interval is the useful operations summary;
            // do not expose the quick-entry's internal "Short breakdown" text.
            detail: isMaintenance
              ? serviceLabel
              : compactCell(String(r.notes || r.description || "").replace(/^(?:Auto from Daily Input \(DOWN\)|Daily Log breakdown)\s*[—-]?\s*/i, ""), 220),
          };
        });
        for (const r of completedBreakdownRepairDowntime) {
          legacyDailyDowntimeRows.push({
            asset: r.asset_code,
            equipment: compactCell(r.asset_name ?? "", 48),
            type: "Breakdown",
            hrs: fmtNum(r.hours_down || 0, 1),
            area: compactCell(r.component ?? "", 42) || "—",
            detail: compactCell(
              `${String(r.description || "Breakdown repair").trim()} · ${fmtNum(r.repair_hours || 0, 1)} completed repair hr`,
              220,
            ),
          });
        }
        for (const r of dailyPlannedMaintenance) {
          if (legacyLoggedDowntimeAssets.has(String(r.asset_code || ""))) continue;
          legacyDailyDowntimeRows.push({
            asset: r.asset_code,
            equipment: compactCell(r.asset_name ?? "", 48),
            type: "Maintenance",
            hrs: "—",
            area: compactCell(r.service_name ?? "", 42) || "Service",
            detail: compactCell(r.description ?? "", 220),
          });
        }

        sectionTitle(doc, `Daily downtime (${opsDay})`);
        table(
          doc,
          [
            { key: "asset", label: "Plant #", width: 0.10 },
            { key: "equipment", label: "Equipment", width: 0.20 },
            { key: "type", label: "Downtime type", width: 0.14 },
            { key: "hrs", label: "Hours down", width: 0.11, align: "right" },
            { key: "area", label: "Component / service", width: 0.17 },
            { key: "detail", label: "Reason / work completed", width: 0.28 },
          ],
          legacyDailyDowntimeRows.length
            ? legacyDailyDowntimeRows
            : [{ asset: "—", equipment: "No downtime recorded", type: "—", hrs: "0.0", area: "—", detail: "—" }],
        );

        sectionTitle(doc, "Open Work Orders");
        table(
          doc,
          [
            { key: "wo", label: "WO#", width: 0.07, align: "right" },
            { key: "asset", label: "Plant #", width: 0.09 },
            { key: "equipment", label: "Equipment", width: 0.16 },
            { key: "source", label: "Source", width: 0.10 },
            { key: "status", label: "Status", width: 0.09 },
            { key: "parts", label: "Parts status", width: 0.12 },
            { key: "tech", label: "Technician", width: 0.11 },
            { key: "progress", label: "Progress / next action", width: 0.26 },
          ],
          openWOsPdf.map(r => ({
            wo: String(r.id),
            asset: r.asset_code,
            equipment: compactCell(r.asset_name ?? "", 48),
            source: ({
              breakdown: "Breakdown",
              service: "Maintenance",
              manager_inspection: "Inspection",
              inspection: "Inspection",
            })[String(r.source || "").toLowerCase()] || compactCell(r.source ?? "", 20),
            status: compactCell(String(r.status ?? "").replace(/_/g, " "), 16),
            parts: compactCell(r.parts_status ?? "", 26) || "—",
            tech: compactCell(r.assigned_artisan_name ?? "", 28),
            progress: compactCell(r.repair_progress ?? "", 260) || "No progress update",
          }))
        );

        if (cartrackSpeeding) {
          sectionTitle(doc, `Fleet Tracking - speeding above ${speedAlertKmh} km/h`);
          if (!cartrackSpeeding.total_speeding_events) {
            doc.fontSize(10).fillColor("#64748b");
            doc.text(`No speeding events above ${speedAlertKmh} km/h recorded for ${opsDay}.`);
            doc.moveDown(0.5);
          } else {
            doc.fontSize(10).fillColor("#334155");
            doc.text(
              `${cartrackSpeeding.total_speeding_events} event(s) above ${speedAlertKmh} km/h across ${cartrackSpeeding.vehicles_with_speeding} vehicle(s).`,
            );
            doc.moveDown(0.4);
            const speedPdf = buildSpeedingReportPdfContent(cartrackSpeeding);
            table(
              doc,
              speedPdf.columns,
              speedPdf.rows.slice(0, 25).map((r) => ({
                time: compactCell(r.time ?? "", 16),
                vehicle: compactCell(String(r.vehicle ?? "").replace(/\n/g, " / "), 24),
                speed: r.speed ?? "—",
                limit: r.limit ?? "—",
                type: compactCell(r.type ?? "", 80),
              })),
            );
          }
        }

      },
      {
        title: "IRONLOG",
        managementTitle: "AML / DAILY OPERATIONS",
        subtitle: `Operating date ${dailyPdfLongDate(opsDay)} | Issued ${dailyPdfLongDate(date)} | ${dailyPdfSiteName}`,
        pageLabel: ({ pageIndex }) => dailyPdfPageLabels[pageIndex] || "OPERATIONS DETAIL",
        sourceText: `Operating day ${opsDay}`,
        headerStyle: "management",
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

  // =========================
  // MONTHLY FLEET COST PDF
  // =========================
  // GET /api/reports/monthly.pdf?month=YYYY-MM&scheduled=10
  // GET /api/reports/monthly.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&scheduled=10
  app.get("/monthly.pdf", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");

    const month = String(req.query?.month || "").trim();
    let start = String(req.query?.start || "").trim();
    let end = String(req.query?.end || "").trim();
    const scheduled = Number(req.query?.scheduled ?? 10);
    const download = String(req.query?.download || "").trim() === "1";

    if (isMonth(month)) {
      const range = monthRange(month);
      start = range.start;
      end = range.end;
    }
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "month (YYYY-MM) or start and end (YYYY-MM-DD) required" });
    }

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const kpi = kpiRange(start, end, scheduled);
    const costs = queryPeriodFleetCostTotals(start, end);
    const costPerRunHour = Number(kpi.run_hours || 0) > 0
      ? Number((costs.total_cost / Number(kpi.run_hours || 1)).toFixed(2))
      : null;

    const assetCosts = buildPeriodAssetCosts(start, end);
    const runHours = queryPeriodRunHoursByAsset(start, end);
    const assetRows = mergeAssetCostsWithRunHours(assetCosts, runHours);
    const categoryRows = rollupFleetCostByCategory(assetRows);
    const contractorFuelRows = buildPeriodContractorFuelRows(start, end);
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

    const assetPdf = assetRows.slice(0, 50);
    const categoryPdf = categoryRows.slice(0, 20);
    const periodLabel = month || `${start} to ${end}`;
    const fileTag = month ? month.replace("-", "") : end.replace(/-/g, "");

    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Fleet KPIs (Period)");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Scheduled hours / asset", v: fmtNum(scheduled, 0) },
          { k: "Available hours", v: fmtNum(kpi.available_hours, 0) },
          { k: "Run hours", v: fmtNum(kpi.run_hours, 1) },
          { k: "Downtime hours", v: fmtNum(kpi.downtime_hours, 1) },
          { k: "Availability %", v: kpi.availability == null ? "N/A" : `${fmtNum(kpi.availability, 2)}%` },
          { k: "Utilization %", v: kpi.utilization == null ? "N/A" : `${fmtNum(kpi.utilization, 2)}%` },
        ], 2);

        sectionTitle(doc, "Fleet Cost Summary (Period)");
        kvGrid(doc, [
          { k: "Fuel Cost", v: fmtNum(costs.fuel_cost, 2) },
          { k: "Oil / Lube Cost", v: fmtNum(costs.lube_cost, 2) },
          { k: "Parts Cost", v: fmtNum(costs.parts_cost, 2) },
          { k: "Labor Cost", v: fmtNum(costs.labor_cost, 2) },
          { k: "Labor Hours", v: fmtNum(costs.labor_hours, 1) },
          { k: "Downtime Cost", v: fmtNum(costs.downtime_cost, 2) },
          { k: "Total Fleet Cost", v: fmtNum(costs.total_cost, 2) },
          { k: "Fleet Cost / Run Hour", v: costPerRunHour == null ? "N/A" : fmtNum(costPerRunHour, 2) },
        ], 2);

        if (contractorFuelRows.length) {
          sectionTitle(doc, "Contractor Fuel Summary (FAMS — hired / archived active)");
          kvGrid(doc, [
            { k: "Contractor assets (fuel)", v: fmtNum(contractorFuelRows.length, 0) },
            { k: "Contractor fuel liters", v: fmtNum(contractorFuelTotal.fuel_liters, 1) },
            { k: "Contractor fuel cost", v: fmtNum(contractorFuelTotal.fuel_cost, 2) },
            { k: "Contractor run (hrs)", v: fmtNum(contractorFuelTotal.hours_run, 1) },
            { k: "Contractor run (km)", v: fmtNum(contractorFuelTotal.km_run, 1) },
            {
              k: "Avg fuel $/hr (hour units)",
              v: contractorFuelTotal.hours_run > 0
                ? fmtNum(contractorFuelTotal.fuel_cost / contractorFuelTotal.hours_run, 2)
                : "N/A",
            },
            {
              k: "Avg fuel $/km (km units)",
              v: contractorFuelTotal.km_run > 0
                ? fmtNum(contractorFuelTotal.fuel_cost / contractorFuelTotal.km_run, 2)
                : "N/A",
            },
          ], 2);

          sectionTitle(doc, "Contractor Fuel by Supplier");
          table(
            doc,
            [
              { key: "supplier", label: "Supplier", width: 0.14 },
              { key: "assets", label: "Assets", width: 0.08, align: "right" },
              { key: "liters", label: "Liters", width: 0.12, align: "right" },
              { key: "cost", label: "Fuel $", width: 0.12, align: "right" },
              { key: "hrs", label: "Run hrs", width: 0.12, align: "right" },
              { key: "km", label: "Run km", width: 0.12, align: "right" },
              { key: "cph", label: "$/hr", width: 0.10, align: "right" },
              { key: "cpk", label: "$/km", width: 0.10, align: "right" },
            ],
            contractorBySupplier.map((r) => ({
              supplier: r.contractor,
              assets: fmtNum(r.asset_count, 0),
              liters: fmtNum(r.fuel_liters, 1),
              cost: fmtNum(r.fuel_cost, 2),
              hrs: fmtNum(r.hours_run, 1),
              km: fmtNum(r.km_run, 1),
              cph: r.hours_run > 0 ? fmtNum(r.fuel_cost / r.hours_run, 2) : "—",
              cpk: r.km_run > 0 ? fmtNum(r.fuel_cost / r.km_run, 2) : "—",
            })),
          );

          sectionTitle(doc, "Contractor Fuel by Asset (FAMS meter run)");
          table(
            doc,
            [
              { key: "asset", label: "Asset", width: 0.10 },
              { key: "supplier", label: "Supplier", width: 0.10 },
              { key: "mode", label: "Mode", width: 0.07 },
              { key: "liters", label: "Liters", width: 0.10, align: "right" },
              { key: "cost", label: "Fuel $", width: 0.10, align: "right" },
              { key: "run", label: "Run", width: 0.10, align: "right" },
              { key: "unit", label: "Unit", width: 0.06 },
              { key: "cpu", label: "$/unit", width: 0.10, align: "right" },
              { key: "fills", label: "Fills", width: 0.07, align: "right" },
            ],
            contractorFuelRows.map((r) => ({
              asset: r.asset_code,
              supplier: r.contractor,
              mode: r.metric_mode === "km" ? "km" : "hrs",
              liters: fmtNum(r.fuel_liters, 1),
              cost: fmtNum(r.fuel_cost, 2),
              run: fmtNum(r.run_value, 1),
              unit: r.run_label,
              cpu: r.cost_per_run == null ? "N/A" : fmtNum(r.cost_per_run, 2),
              fills: fmtNum(r.fill_count, 0),
            })),
          );
        }

        sectionTitle(doc, "Cost by Category ($/run hr includes labor + lube)");
        table(
          doc,
          [
            { key: "cat", label: "Category", width: 0.16 },
            { key: "run", label: "Run hrs", width: 0.10, align: "right" },
            { key: "fuel", label: "Fuel", width: 0.10, align: "right" },
            { key: "lube", label: "Lube", width: 0.10, align: "right" },
            { key: "labor", label: "Labor", width: 0.10, align: "right" },
            { key: "parts", label: "Parts", width: 0.10, align: "right" },
            { key: "total", label: "Total", width: 0.12, align: "right" },
            { key: "cph", label: "$/hr", width: 0.10, align: "right" },
          ],
          categoryPdf.map((r) => ({
            cat: compactCell(r.category ?? "", 24),
            run: fmtNum(r.run_hours, 1),
            fuel: fmtNum(r.fuel_cost, 0),
            lube: fmtNum(r.lube_cost, 0),
            labor: fmtNum(r.labor_cost, 0),
            parts: fmtNum(r.parts_cost, 0),
            total: fmtNum(r.total_cost, 0),
            cph: r.cost_per_run_hour == null ? "N/A" : fmtNum(r.cost_per_run_hour, 2),
          })),
        );

        sectionTitle(doc, "Cost by Asset (Top 50 — labor + lube + fuel + parts + downtime)");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.12 },
            { key: "run", label: "Run hrs", width: 0.09, align: "right" },
            { key: "fuel", label: "Fuel", width: 0.09, align: "right" },
            { key: "lube", label: "Lube", width: 0.09, align: "right" },
            { key: "labor", label: "Labor", width: 0.09, align: "right" },
            { key: "parts", label: "Parts", width: 0.09, align: "right" },
            { key: "down", label: "Down", width: 0.09, align: "right" },
            { key: "total", label: "Total", width: 0.10, align: "right" },
            { key: "cph", label: "$/hr", width: 0.08, align: "right" },
          ],
          assetPdf.map((r) => ({
            asset: r.asset_code,
            run: fmtNum(r.run_hours, 1),
            fuel: fmtNum(r.fuel_cost, 0),
            lube: fmtNum(r.lube_cost, 0),
            labor: fmtNum(r.labor_cost, 0),
            parts: fmtNum(r.parts_cost, 0),
            down: fmtNum(r.downtime_cost, 0),
            total: fmtNum(r.total_cost, 0),
            cph: r.cost_per_run_hour == null ? "N/A" : fmtNum(r.cost_per_run_hour, 2),
          })),
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Monthly Fleet Cost Report",
        rightText: `Period: ${periodLabel}`,
        showPageNumbers: true,
        layout: "landscape",
      },
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Monthly_Fleet_Cost_${fileTag}.pdf"`,
      )
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
