// IRONLOG/api/routes/maintenance/plans.routes.js — Maintenance plans, services due, service history and backfill.
// Registered by routes/maintenance.routes.js; shared helpers arrive through ctx.
import ExcelJS from "exceljs";
import { buildDueListFromPlans, classifyServiceDue, groupActivePlansByAsset, hasRotatingSchedule, meterUnitForAsset, planIntervalHours, resolveLegacyPlanDue, resolveNextServiceForAssetPlans, snapLastServiceHours } from "../../utils/serviceSchedule.js";
import { buildPdfBuffer, sectionTitle, table } from "../../utils/pdfGenerator.js";
import { createManagementSummary, styleManagementDetailSheet } from "../../utils/managementWorkbook.js";
import { db } from "../../db/client.js";
import { enrichDueRowsWithEstimates } from "../../utils/maintenanceEstimates.js";
import { isDate } from "../../utils/request.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerPlansRoutes(app, ctx) {
  const {
    displayPlanServiceName,
    getAssetCurrentHours,
    getAssetCurrentHoursInfo,
    getAssetHoursInfoAsOf,
    hasTable,
    listMaintenancePlans,
  } = ctx;

  app.get("/plans", async (req, reply) => {
    try {
      const plans = listMaintenancePlans();

      return reply.send({
        ok: true,
        plans
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - EXCEL EXPORT
  // GET /api/maintenance/plans.xlsx?near_due_hours=50
  // =====================================================
  app.get("/plans.xlsx", async (req, reply) => {
    try {
      const requestedThreshold = Number(req.query?.near_due_hours || 50);
      const nearDueHours = Number.isFinite(requestedThreshold)
        ? Math.max(1, Math.min(100000, requestedThreshold))
        : 50;
      const plans = listMaintenancePlans(nearDueHours);
      const planningQueue = buildDueListFromPlans(plans, getAssetCurrentHours, nearDueHours)
        .sort((a, b) => {
          const statusRank = { OVERDUE: 1, "ALMOST DUE": 2, OK: 3 };
          const rank = (statusRank[String(a.status || "OK").toUpperCase()] || 4)
            - (statusRank[String(b.status || "OK").toUpperCase()] || 4);
          if (rank !== 0) return rank;
          const remaining = Number(a.remaining_hours || 0) - Number(b.remaining_hours || 0);
          if (remaining !== 0) return remaining;
          return String(a.asset_code || "").localeCompare(String(b.asset_code || ""));
        });
      const overdueCount = planningQueue.filter((row) => String(row.status || "").toUpperCase() === "OVERDUE").length;
      const dueSoonCount = planningQueue.filter((row) => String(row.status || "").toUpperCase() === "ALMOST DUE").length;
      const activeAssetCount = new Set(planningQueue.map((row) => Number(row.asset_id || 0)).filter(Boolean)).size;
      const dateTag = new Date().toISOString().slice(0, 10);

      const wb = new ExcelJS.Workbook();
      wb.creator = "IRONLOG";
      wb.created = new Date();
      createManagementSummary(wb, {
        title: "IRONLOG Maintenance Plans",
        periodLabel: `Live planning view generated ${dateTag}`,
        cards: [
          { label: "ACTIVE ASSETS", value: activeAssetCount, numFmt: "#,##0" },
          { label: "OVERDUE", value: overdueCount, numFmt: "#,##0", tone: "warning" },
          { label: "DUE SOON", value: dueSoonCount, numFmt: "#,##0", tone: "attention" },
          { label: "CONFIGURED SCHEDULES", value: plans.length, numFmt: "#,##0" },
        ],
        scopeLines: [
          "Planning Queue shows one next service for every active asset. Service Schedules keeps every configured service interval.",
          `Near-due threshold: ${nearDueHours} hours. LDV schedules use their 500 km early-warning threshold.`,
        ],
      });

      const queueSheet = wb.addWorksheet("Planning Queue");
      queueSheet.columns = [
        { header: "Asset code", key: "asset_code", width: 14 },
        { header: "Equipment", key: "asset_name", width: 30 },
        { header: "Category", key: "category", width: 20 },
        { header: "Current meter", key: "current_hours", width: 16 },
        { header: "Unit", key: "meter_unit", width: 10 },
        { header: "Last service", key: "last_service_hours", width: 16 },
        { header: "Next service", key: "service_name", width: 22 },
        { header: "Due meter", key: "next_due_hours", width: 16 },
        { header: "Remaining", key: "remaining_hours", width: 14 },
        { header: "Status", key: "status", width: 15 },
      ];
      queueSheet.addRows(planningQueue.map((row) => ({
        asset_code: row.asset_code || "",
        asset_name: row.asset_name || "",
        category: row.category || "",
        current_hours: Number(row.current_hours || 0),
        meter_unit: row.meter_unit || meterUnitForAsset(row.asset_code),
        last_service_hours: Number(row.last_service_hours || 0),
        service_name: displayPlanServiceName(row),
        next_due_hours: Number(row.next_due_hours || 0),
        remaining_hours: Number(row.remaining_hours || 0),
        status: row.status || "OK",
      })));
      styleManagementDetailSheet(queueSheet, {
        title: "Maintenance planning queue",
        subtitle: `Live meter readings · Near-due threshold: ${nearDueHours} hours`,
        frozenColumns: 2,
        numberFormats: {
          current_hours: "#,##0.0",
          last_service_hours: "#,##0.0",
          next_due_hours: "#,##0.0",
          remaining_hours: "#,##0.0",
        },
      });

      const scheduleSheet = wb.addWorksheet("Service Schedules");
      scheduleSheet.columns = [
        { header: "Asset code", key: "asset_code", width: 14 },
        { header: "Equipment", key: "asset_name", width: 30 },
        { header: "Category", key: "category", width: 20 },
        { header: "Service schedule", key: "service_name", width: 24 },
        { header: "Interval", key: "interval_hours", width: 13 },
        { header: "Unit", key: "meter_unit", width: 10 },
        { header: "Active", key: "active", width: 11 },
        { header: "Last service", key: "last_service_hours", width: 16 },
        { header: "Current meter", key: "current_hours", width: 16 },
        { header: "Next due", key: "next_due_hours", width: 16 },
        { header: "Remaining", key: "remaining_hours", width: 14 },
        { header: "Next for asset", key: "is_next_for_asset", width: 15 },
        { header: "Status", key: "status", width: 15 },
      ];
      scheduleSheet.addRows(plans.map((row) => ({
        asset_code: row.asset_code || "",
        asset_name: row.asset_name || "",
        category: row.category || "",
        service_name: displayPlanServiceName(row),
        interval_hours: Number(row.interval_hours || 0),
        meter_unit: row.meter_unit || meterUnitForAsset(row.asset_code),
        active: Number(row.active || 0) ? "Active" : "Inactive",
        last_service_hours: Number(row.last_service_hours_snapped ?? row.last_service_hours ?? 0),
        current_hours: Number(row.current_hours || 0),
        next_due_hours: row.is_next_for_asset ? Number(row.next_due_hours || 0) : null,
        remaining_hours: row.is_next_for_asset ? Number(row.remaining_hours || 0) : null,
        is_next_for_asset: row.is_next_for_asset ? "Yes" : "",
        status: row.is_next_for_asset ? (row.status || "OK") : "",
      })));
      styleManagementDetailSheet(scheduleSheet, {
        title: "Configured maintenance schedules",
        subtitle: "All service plans, including inactive schedules. Only the next due service is marked per active asset.",
        frozenColumns: 2,
        numberFormats: {
          interval_hours: "#,##0.0",
          last_service_hours: "#,##0.0",
          current_hours: "#,##0.0",
          next_due_hours: "#,##0.0",
          remaining_hours: "#,##0.0",
        },
      });

      const buffer = await wb.xlsx.writeBuffer();
      return reply
        .header("Content-Disposition", `attachment; filename=IRONLOG_Maintenance_Plans_${dateTag}.xlsx`)
        .type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .send(buffer);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || "Unable to export maintenance plans" });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - CREATE
  // POST /api/maintenance/plans
  // =====================================================
  app.post("/plans", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const service_name = String(req.body?.service_name || "").trim();
      const interval_hours = Number(req.body?.interval_hours || 0);
      const active = Number(req.body?.active ?? 1) ? 1 : 0;

      if (!asset_id || !service_name || interval_hours <= 0) {
        return reply.code(400).send({
          ok: false,
          error: "asset_id, service_name and interval_hours are required"
        });
      }

      const asset = db.prepare(`
        SELECT id, asset_code
        FROM assets
        WHERE id = ?
      `).get(asset_id);

      if (!asset) {
        return reply.code(404).send({
          ok: false,
          error: "Asset not found"
        });
      }

      const last_service_hours = snapLastServiceHours(
        Number(req.body?.last_service_hours || 0),
        interval_hours,
        asset.asset_code,
      );

      const result = db.prepare(`
        INSERT INTO maintenance_plans (
          asset_id,
          service_name,
          interval_hours,
          last_service_hours,
          active
        )
        VALUES (?, ?, ?, ?, ?)
      `).run(
        asset_id,
        service_name,
        interval_hours,
        last_service_hours,
        active
      );
      writeAudit(db, req, {
        module: "maintenance",
        action: "plan.create",
        entity_type: "maintenance_plan",
        entity_id: String(Number(result.lastInsertRowid || 0)),
        after: { asset_id, service_name, interval_hours, last_service_hours, active },
      });

      return reply.send({
        ok: true,
        id: Number(result.lastInsertRowid)
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - UPDATE
  // PUT /api/maintenance/plans/:id
  // =====================================================
  app.put("/plans/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) {
        return reply.code(400).send({
          ok: false,
          error: "Invalid plan id"
        });
      }

      const existing = db.prepare(`
        SELECT *
        FROM maintenance_plans
        WHERE id = ?
      `).get(id);

      if (!existing) {
        return reply.code(404).send({
          ok: false,
          error: "Maintenance plan not found"
        });
      }

      const asset_id =
        req.body?.asset_id != null ? Number(req.body.asset_id) : Number(existing.asset_id);

      const service_name =
        req.body?.service_name != null
          ? String(req.body.service_name).trim()
          : String(existing.service_name || "").trim();

      const interval_hours =
        req.body?.interval_hours != null
          ? Number(req.body.interval_hours)
          : Number(existing.interval_hours || 0);

      const last_service_hours_raw =
        req.body?.last_service_hours != null
          ? Number(req.body.last_service_hours)
          : Number(existing.last_service_hours || 0);

      const active =
        req.body?.active != null
          ? (Number(req.body.active) ? 1 : 0)
          : Number(existing.active || 0);

      if (!asset_id || !service_name || interval_hours <= 0) {
        return reply.code(400).send({
          ok: false,
          error: "asset_id, service_name and interval_hours are required"
        });
      }

      const asset = db.prepare(`
        SELECT id, asset_code
        FROM assets
        WHERE id = ?
      `).get(asset_id);

      if (!asset) {
        return reply.code(404).send({
          ok: false,
          error: "Asset not found"
        });
      }

      const last_service_hours = snapLastServiceHours(
        last_service_hours_raw,
        interval_hours,
        asset.asset_code,
      );

      db.prepare(`
        UPDATE maintenance_plans
        SET
          asset_id = ?,
          service_name = ?,
          interval_hours = ?,
          last_service_hours = ?,
          active = ?
        WHERE id = ?
      `).run(
        asset_id,
        service_name,
        interval_hours,
        last_service_hours,
        active,
        id
      );
      writeAudit(db, req, {
        module: "maintenance",
        action: "plan.update",
        entity_type: "maintenance_plan",
        entity_id: String(id),
        before: existing,
        after: { asset_id, service_name, interval_hours, last_service_hours, active },
      });

      return reply.send({ ok: true });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - DELETE
  // DELETE /api/maintenance/plans/:id
  // =====================================================
  app.delete("/plans/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) {
        return reply.code(400).send({ ok: false, error: "Invalid plan id" });
      }

      const existing = db.prepare(`
        SELECT *
        FROM maintenance_plans
        WHERE id = ?
      `).get(id);

      if (!existing) {
        return reply.code(404).send({ ok: false, error: "Maintenance plan not found" });
      }

      db.prepare(`
        DELETE FROM maintenance_plans WHERE id = ?
      `).run(id);
      writeAudit(db, req, {
        module: "maintenance",
        action: "plan.delete",
        entity_type: "maintenance_plan",
        entity_id: String(id),
        before: existing,
      });

      return reply.send({ ok: true });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - TOGGLE ACTIVE
  // PATCH /api/maintenance/plans/:id/toggle
  // =====================================================
  app.patch("/plans/:id/toggle", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) {
        return reply.code(400).send({ ok: false, error: "Invalid plan id" });
      }

      const plan = db.prepare(`
        SELECT id, active
        FROM maintenance_plans
        WHERE id = ?
      `).get(id);

      if (!plan) {
        return reply.code(404).send({ ok: false, error: "Maintenance plan not found" });
      }

      const newActive = plan.active ? 0 : 1;
      db.prepare(`
        UPDATE maintenance_plans
        SET active = ?
        WHERE id = ?
      `).run(newActive, id);

      return reply.send({ ok: true, id, active: newActive });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - REBASE LAST SERVICE TO LIVE HOURS
  // POST /api/maintenance/plans/:id/rebase-last-service
  // =====================================================
  app.post("/plans/:id/rebase-last-service", async (req, reply) => {
    try {
      const idFromPath = Number(req.params?.id || 0);
      const idFromBody = Number(req.body?.plan_id || 0);
      const id = idFromPath > 0 ? idFromPath : (idFromBody > 0 ? idFromBody : 0);

      let plan = null;
      if (id > 0) {
        plan = db.prepare(`
          SELECT mp.id, mp.asset_id, mp.service_name, a.asset_code, a.asset_name
          FROM maintenance_plans mp
          JOIN assets a ON a.id = mp.asset_id
          WHERE mp.id = ?
          LIMIT 1
        `).get(id);
      }
      if (!plan) {
        const assetId = Number(req.body?.asset_id || 0);
        const serviceName = String(req.body?.service_name || "").trim();
        if (assetId > 0 && serviceName) {
          plan = db.prepare(`
            SELECT mp.id, mp.asset_id, mp.service_name, a.asset_code, a.asset_name
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.asset_id = ?
              AND UPPER(TRIM(mp.service_name)) = UPPER(TRIM(?))
            ORDER BY mp.active DESC, mp.id DESC
            LIMIT 1
          `).get(assetId, serviceName);
        }
      }

      if (!plan) {
        return reply.code(400).send({ ok: false, error: "Invalid plan id or plan lookup context" });
      }

      const currentInfo = getAssetCurrentHoursInfo(Number(plan.asset_id || 0));
      const liveHours = Number(currentInfo.hours || 0);
      const planMeta = db.prepare(`
        SELECT interval_hours FROM maintenance_plans WHERE id = ?
      `).get(Number(plan.id));
      const safeHours = snapLastServiceHours(
        Number.isFinite(liveHours) ? liveHours : 0,
        Number(planMeta?.interval_hours || 0),
        plan.asset_code,
      );

      db.prepare(`
        UPDATE maintenance_plans
        SET last_service_hours = ?
        WHERE id = ?
      `).run(safeHours, Number(plan.id));

      db.prepare(`
        UPDATE maintenance_plans
        SET last_service_hours = ?
        WHERE asset_id = ?
          AND active = 1
      `).run(safeHours, Number(plan.asset_id || 0));

      return reply.send({
        ok: true,
        id: Number(plan.id),
        asset_id: Number(plan.asset_id || 0),
        asset_code: plan.asset_code,
        asset_name: plan.asset_name,
        service_name: plan.service_name,
        last_service_hours: safeHours,
        source: currentInfo.source
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

    // =====================================================
  // GET LIVE HOURS FOR ONE ASSET
  // GET /api/maintenance/asset/:id/live-hours?as_of=YYYY-MM-DD (optional; meter/usage up to that date)
  // =====================================================
  app.get("/asset/:id/live-hours", async (req, reply) => {
    try {
      const assetId = Number(req.params?.id || 0);
      if (!assetId) {
        return reply.code(400).send({
          ok: false,
          error: "Invalid asset id"
        });
      }

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name
        FROM assets
        WHERE id = ?
      `).get(assetId);

      if (!asset) {
        return reply.code(404).send({
          ok: false,
          error: "Asset not found"
        });
      }

      const asOf = String(req.query?.as_of || "").trim();
      const currentInfo = isDate(asOf)
        ? getAssetHoursInfoAsOf(assetId, asOf)
        : getAssetCurrentHoursInfo(assetId);
      const current_hours = Number(currentInfo.hours || 0);

      const assetPlans = db.prepare(`
        SELECT id, service_name, interval_hours, last_service_hours, active
        FROM maintenance_plans
        WHERE asset_id = ?
          AND active = 1
        ORDER BY interval_hours ASC, id ASC
      `).all(assetId);
      const next_service = assetPlans.length
        ? resolveNextServiceForAssetPlans(assetPlans, current_hours, asset.asset_code)
        : null;

      return reply.send({
        ok: true,
        asset_id: asset.id,
        asset_code: asset.asset_code,
        asset_name: asset.asset_name,
        as_of: isDate(asOf) ? asOf : null,
        current_hours: Number(current_hours.toFixed(1)),
        current_hours_source: currentInfo.source,
        next_service,
        rotating_schedule: hasRotatingSchedule(assetPlans),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // LIST MAINTENANCE DUE
  // GET /api/maintenance/due?date=2026-02-27&near_due_hours=50
  // =====================================================
  app.get("/due", async (req, reply) => {
    try {
      const date = String(req.query?.date || "").trim();
      const nearDueHours = Math.max(1, Number(req.query?.near_due_hours || 50));
      if (date && !isDate(date)) {
        return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
      }

      const rows = date
        ? db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              mp.active,
              a.asset_code,
              a.asset_name,
              a.category,
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
              AND a.archived = 0
            ORDER BY a.asset_code, mp.id
          `).all(date)
        : db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              mp.active,
              a.asset_code,
              a.asset_name,
              a.category,
              IFNULL((
                SELECT SUM(dh.hours_run)
                FROM daily_hours dh
                WHERE dh.asset_id = a.id
                  AND dh.is_used = 1
                  AND dh.hours_run > 0
              ), 0) AS current_hours
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.is_standby = 0
              AND a.archived = 0
            ORDER BY a.asset_code, mp.id
          `).all();

      const getHours = (assetId) => {
        if (date) return Number(rows.find((r) => Number(r.asset_id) === Number(assetId))?.current_hours ?? getAssetCurrentHours(assetId));
        return getAssetCurrentHours(assetId);
      };
      const due = enrichDueRowsWithEstimates(
        db,
        buildDueListFromPlans(rows, getHours, nearDueHours),
        { as_of: date || new Date().toISOString().slice(0, 10), history_days: 14 },
      );

      return reply.send({
        ok: true,
        as_of: date || null,
        near_due_hours: nearDueHours,
        due
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // MAINTENANCE HISTORY (single-line per asset+service)
  // GET /api/maintenance/history?as_of=YYYY-MM-DD&days=14
  // =====================================================
  app.get("/history", async (req, reply) => {
    try {
      const as_of = String(req.query?.as_of || "").trim();
      const days = Math.max(3, Math.min(120, Number(req.query?.days || 14)));
      if (as_of && !isDate(as_of)) return reply.code(400).send({ error: "as_of must be YYYY-MM-DD" });

      const endDate = as_of || new Date().toISOString().slice(0, 10);
      const startD = new Date(endDate + "T00:00:00");
      startD.setDate(startD.getDate() - (days - 1));
      const startDate = startD.toISOString().slice(0, 10);

      const plans = db.prepare(`
        SELECT
          mp.id AS plan_id,
          mp.asset_id,
          mp.service_name,
          mp.interval_hours,
          mp.last_service_hours,
          mp.active,
          a.asset_code,
          a.asset_name,
          a.category
        FROM maintenance_plans mp
        JOIN assets a ON a.id = mp.asset_id
        WHERE mp.active = 1
          AND a.active = 1
          AND a.is_standby = 0
          AND a.archived = 0
        ORDER BY a.asset_code ASC, mp.service_name ASC
      `).all();

      const getLastServiced = db.prepare(`
        SELECT
          DATE(COALESCE(w.closed_at, w.completed_at)) AS last_serviced_date,
          COALESCE(w.closed_at, w.completed_at) AS last_serviced_at
        FROM work_orders w
        WHERE w.source = 'service'
          AND w.reference_id = ?
          AND w.status IN ('completed', 'approved', 'closed')
        ORDER BY COALESCE(w.closed_at, w.completed_at) DESC
        LIMIT 1
      `);
      const getLastBackfillServiced = db.prepare(`
        SELECT
          h.id AS history_id,
          DATE(h.service_date) AS last_serviced_date,
          h.service_date AS last_serviced_at
        FROM maintenance_service_history h
        WHERE h.asset_id = ?
          AND UPPER(TRIM(h.service_name)) = UPPER(TRIM(?))
        ORDER BY h.service_date DESC, h.id DESC
        LIMIT 1
      `);

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
        const d = new Date(dateStr + "T00:00:00");
        d.setDate(d.getDate() + Math.round(add));
        return d.toISOString().slice(0, 10);
      };

      const getLastServicedAnyOnAsset = db.prepare(`
        SELECT
          DATE(COALESCE(w.closed_at, w.completed_at)) AS last_serviced_date,
          COALESCE(w.closed_at, w.completed_at) AS last_serviced_at
        FROM work_orders w
        WHERE w.source = 'service'
          AND w.asset_id = ?
          AND w.status IN ('completed', 'approved', 'closed')
        ORDER BY COALESCE(w.closed_at, w.completed_at) DESC
        LIMIT 1
      `);

      const rows = [];
      const byAsset = groupActivePlansByAsset(plans);
      for (const [, assetPlans] of byAsset) {
        const sample = assetPlans[0];
        const assetId = Number(sample.asset_id || 0);
        const currentInfo = getAssetCurrentHoursInfo(assetId);
        const current = Number(currentInfo.hours || 0);
        const rotating = hasRotatingSchedule(assetPlans);

        if (rotating) {
          const resolved = resolveNextServiceForAssetPlans(
            assetPlans,
            current,
            String(sample.asset_code || ""),
          );
          const lastWo = getLastServicedAnyOnAsset.get(assetId);
          const lastBackfill = db.prepare(`
            SELECT
              h.id AS history_id,
              DATE(h.service_date) AS last_serviced_date,
              h.service_date AS last_serviced_at
            FROM maintenance_service_history h
            WHERE h.asset_id = ?
            ORDER BY h.service_date DESC, h.id DESC
            LIMIT 1
          `).get(assetId);
          const lastWoAt = String(lastWo?.last_serviced_at || "");
          const lastBackfillAt = String(lastBackfill?.last_serviced_at || "");
          const useBackfill = Boolean(lastBackfillAt && (!lastWoAt || lastBackfillAt > lastWoAt));
          const last = useBackfill ? lastBackfill : lastWo;
          const avgRow = getAvgDaily.get(assetId, startDate, endDate);
          const totalRun = Number(avgRow?.total_run || 0);
          const dayCount = Number(avgRow?.day_count || 0);
          const avgDaily = dayCount > 0 ? totalRun / dayCount : 0;
          const remaining = Number(resolved?.remaining_hours || 0);
          const estDays = avgDaily > 0 ? Math.max(0, remaining / avgDaily) : null;
          const estDate = estDays == null ? null : addDays(endDate, estDays);

          const dueMeta = classifyServiceDue(
            remaining,
            String(sample.asset_code || ""),
            Number(resolved?.interval_hours || resolved?.next_service_interval || 0),
            50,
          );

          rows.push({
            plan_id: Number(resolved?.plan_id || 0),
            asset_id: assetId,
            asset_code: sample.asset_code,
            asset_name: sample.asset_name,
            service_name: resolved?.service_name || sample.service_name,
            last_serviced_date: last?.last_serviced_date || null,
            last_service_source: useBackfill ? "backfill" : (last?.last_serviced_at ? "work_order" : null),
            last_service_history_id: useBackfill ? Number(lastBackfill?.history_id || 0) : null,
            current_hours: Number(current.toFixed(2)),
            current_hours_source: currentInfo.source,
            remaining_hours: Number(remaining.toFixed(2)),
            avg_daily_hours: Number(avgDaily.toFixed(2)),
            estimated_service_date: estDate,
            schedule_mode: "rotating",
            next_due_hours: resolved?.next_due_hours ?? null,
            last_service_hours: resolved?.last_service_hours ?? null,
            meter_unit: dueMeta.meter_unit,
            near_due_threshold: dueMeta.near_due_threshold,
            status: dueMeta.status,
            is_almost_due: dueMeta.is_almost_due,
          });
          continue;
        }

        for (const p of assetPlans) {
          const resolved = resolveLegacyPlanDue(p, current, String(p.asset_code || ""));
          const remaining = resolved.remaining_hours;

          const lastWo = getLastServiced.get(Number(p.plan_id || 0));
          const lastBackfill = getLastBackfillServiced.get(assetId, String(p.service_name || ""));
          const lastWoAt = String(lastWo?.last_serviced_at || "");
          const lastBackfillAt = String(lastBackfill?.last_serviced_at || "");
          const useBackfill = Boolean(lastBackfillAt && (!lastWoAt || lastBackfillAt > lastWoAt));
          const last = useBackfill ? lastBackfill : lastWo;
          const avgRow = getAvgDaily.get(assetId, startDate, endDate);
          const totalRun = Number(avgRow?.total_run || 0);
          const dayCount = Number(avgRow?.day_count || 0);
          const avgDaily = dayCount > 0 ? totalRun / dayCount : 0;

          const estDays = avgDaily > 0 ? Math.max(0, remaining / avgDaily) : null;
          const estDate = estDays == null ? null : addDays(endDate, estDays);

          const dueMeta = classifyServiceDue(
            remaining,
            String(p.asset_code || ""),
            Number(resolved.interval_hours || 0),
            50,
          );

          rows.push({
            plan_id: Number(p.plan_id || 0),
            asset_id: assetId,
            asset_code: p.asset_code,
            asset_name: p.asset_name,
            service_name: p.service_name,
            last_serviced_date: last?.last_serviced_date || null,
            last_service_source: useBackfill ? "backfill" : (last?.last_serviced_at ? "work_order" : null),
            last_service_history_id: useBackfill ? Number(lastBackfill?.history_id || 0) : null,
            current_hours: Number(current.toFixed(2)),
            current_hours_source: currentInfo.source,
            remaining_hours: Number(remaining.toFixed(2)),
            avg_daily_hours: Number(avgDaily.toFixed(2)),
            estimated_service_date: estDate,
            schedule_mode: "grid",
            next_due_hours: resolved.next_due_hours,
            last_service_hours: resolved.last_service_hours,
            meter_unit: dueMeta.meter_unit,
            near_due_threshold: dueMeta.near_due_threshold,
            status: dueMeta.status,
            is_almost_due: dueMeta.is_almost_due,
          });
        }
      }

      return reply.send({ ok: true, as_of: endDate, range: { start: startDate, end: endDate }, rows });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // =====================================================
  // BACKFILL (ANCIENT) SERVICE HISTORY
  // POST /api/maintenance/history/backfill
  // Body: { asset_id, service_name, service_date, service_hours?, notes?, update_plan_last_hours?, plan_id? }
  // =====================================================
  app.post("/history/backfill", async (req, reply) => {
    try {
      const body = req.body || {};
      const assetId = Number(body.asset_id || 0);
      const serviceName = String(body.service_name || "").trim();
      const serviceDate = String(body.service_date || "").trim();
      const serviceHoursIn = body.service_hours;
      const notes = String(body.notes || "").trim() || null;
      const updatePlanLastHours = Number(body.update_plan_last_hours || 0) === 1;
      const planIdIn = Number(body.plan_id || 0);

      if (!assetId) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!serviceName) return reply.code(400).send({ ok: false, error: "service_name is required" });
      if (!isDate(serviceDate)) return reply.code(400).send({ ok: false, error: "service_date must be YYYY-MM-DD" });

      const serviceHours = serviceHoursIn == null || String(serviceHoursIn).trim() === ""
        ? null
        : Number(serviceHoursIn);
      if (serviceHours != null && (!Number.isFinite(serviceHours) || serviceHours < 0)) {
        return reply.code(400).send({ ok: false, error: "service_hours must be a valid number >= 0" });
      }

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name
        FROM assets
        WHERE id = ?
        LIMIT 1
      `).get(assetId);
      if (!asset) return reply.code(404).send({ ok: false, error: "asset not found" });

      let planId = planIdIn > 0 ? planIdIn : null;
      let planInterval = 0;
      if (!planId) {
        const matchedPlan = db.prepare(`
          SELECT id, interval_hours
          FROM maintenance_plans
          WHERE asset_id = ?
            AND UPPER(TRIM(service_name)) = UPPER(TRIM(?))
          ORDER BY active DESC, id DESC
          LIMIT 1
        `).get(assetId, serviceName);
        if (matchedPlan?.id) {
          planId = Number(matchedPlan.id);
          planInterval = Number(matchedPlan.interval_hours || 0);
        }
      } else {
        const matchedPlan = db.prepare(`
          SELECT interval_hours FROM maintenance_plans WHERE id = ?
        `).get(planId);
        planInterval = Number(matchedPlan?.interval_hours || 0);
      }
      if (!planInterval) planInterval = planIntervalHours({ service_name: serviceName, interval_hours: 0 });

      const snappedServiceHours = serviceHours != null
        ? snapLastServiceHours(serviceHours, planInterval, asset.asset_code)
        : null;

      const insert = db.prepare(`
        INSERT INTO maintenance_service_history (
          asset_id, plan_id, service_name, service_date, service_hours, notes, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
      `);
      const updatePlan = db.prepare(`
        UPDATE maintenance_plans
        SET last_service_hours = ?
        WHERE id = ?
      `);

      const tx = db.transaction(() => {
        const r = insert.run(
          assetId,
          planId,
          serviceName,
          serviceDate,
          snappedServiceHours,
          notes,
          String(req.headers?.["x-user-name"] || "system")
        );
        if (updatePlanLastHours && planId && snappedServiceHours != null) {
          updatePlan.run(snappedServiceHours, Number(planId));
          db.prepare(`
            UPDATE maintenance_plans SET last_service_hours = ?
            WHERE asset_id = ? AND active = 1
          `).run(snappedServiceHours, assetId);
        }
        return Number(r.lastInsertRowid || 0);
      });

      const id = tx();
      writeAudit(db, req, {
        module: "maintenance",
        action: "service_history.create",
        entity_type: "maintenance_service_history",
        entity_id: String(id),
        after: {
          asset_id: assetId,
          plan_id: planId || null,
          service_name: serviceName,
          service_date: serviceDate,
          service_hours: snappedServiceHours,
        },
      });
      return reply.send({
        ok: true,
        id,
        asset_id: assetId,
        asset_code: asset.asset_code,
        service_name: serviceName,
        service_date: serviceDate,
        service_hours: snappedServiceHours,
        plan_id: planId || null,
        plan_last_hours_updated: Boolean(updatePlanLastHours && planId && snappedServiceHours != null),
      });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // GET /api/maintenance/history/backfill?asset_id=&asset_code=&limit=200&include_work_orders=1
  app.get("/history/backfill", async (req, reply) => {
    try {
      let assetId = Number(req.query?.asset_id || 0);
      const assetCode = String(req.query?.asset_code || "").trim();
      if (!assetId && assetCode) {
        const a = db.prepare(`
          SELECT id FROM assets WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?)) LIMIT 1
        `).get(assetCode);
        assetId = Number(a?.id || 0);
      }
      const limit = Math.max(1, Math.min(500, Number(req.query?.limit || 200)));
      const includeWorkOrders = String(req.query?.include_work_orders || "1").trim() !== "0";
      const where = assetId > 0 ? "WHERE h.asset_id = ?" : "";
      const backfillRows = db.prepare(`
        SELECT
          h.id,
          h.asset_id,
          h.plan_id,
          h.service_name,
          h.service_date,
          h.service_hours,
          h.notes,
          h.created_by,
          h.created_at,
          a.asset_code,
          a.asset_name,
          'backfill' AS record_source
        FROM maintenance_service_history h
        JOIN assets a ON a.id = h.asset_id
        ${where}
        ORDER BY h.service_date DESC, h.id DESC
        LIMIT ?
      `).all(...(assetId > 0 ? [assetId, limit] : [limit]));

      let woRows = [];
      if (includeWorkOrders && hasTable("work_orders")) {
        const woWhere = assetId > 0 ? "AND w.asset_id = ?" : "";
        woRows = db.prepare(`
          SELECT
            w.id,
            w.asset_id,
            w.reference_id AS plan_id,
            COALESCE(mp.service_name, 'Service') AS service_name,
            DATE(COALESCE(w.closed_at, w.completed_at)) AS service_date,
            NULL AS service_hours,
            w.completion_notes AS notes,
            w.artisan_name AS created_by,
            COALESCE(w.closed_at, w.completed_at) AS created_at,
            a.asset_code,
            a.asset_name,
            'work_order' AS record_source
          FROM work_orders w
          JOIN assets a ON a.id = w.asset_id
          LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id
          WHERE LOWER(COALESCE(w.source, '')) = 'service'
            AND LOWER(COALESCE(w.status, '')) IN ('completed', 'approved', 'closed')
            ${woWhere}
          ORDER BY COALESCE(w.closed_at, w.completed_at) DESC, w.id DESC
          LIMIT ?
        `).all(...(assetId > 0 ? [assetId, limit] : [limit]));
      }

      const merged = [...backfillRows, ...woRows]
        .sort((a, b) => {
          const da = String(a.service_date || a.created_at || "");
          const dbd = String(b.service_date || b.created_at || "");
          return dbd.localeCompare(da) || Number(b.id || 0) - Number(a.id || 0);
        })
        .slice(0, limit);

      return reply.send({ ok: true, rows: merged, asset_id: assetId || null });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // PUT /api/maintenance/history/backfill/:id
  app.put("/history/backfill/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
      const body = req.body || {};
      const serviceName = String(body.service_name || "").trim();
      const serviceDate = String(body.service_date || "").trim();
      const notes = String(body.notes || "").trim() || null;
      const serviceHoursIn = body.service_hours;
      const serviceHours = serviceHoursIn == null || String(serviceHoursIn).trim() === ""
        ? null
        : Number(serviceHoursIn);
      if (!serviceName) return reply.code(400).send({ ok: false, error: "service_name is required" });
      if (!isDate(serviceDate)) return reply.code(400).send({ ok: false, error: "service_date must be YYYY-MM-DD" });
      if (serviceHours != null && (!Number.isFinite(serviceHours) || serviceHours < 0)) {
        return reply.code(400).send({ ok: false, error: "service_hours must be a valid number >= 0" });
      }
      const cur = db.prepare(`SELECT * FROM maintenance_service_history WHERE id = ?`).get(id);
      if (!cur) return reply.code(404).send({ ok: false, error: "backfill entry not found" });
      db.prepare(`
        UPDATE maintenance_service_history
        SET service_name = ?, service_date = ?, service_hours = ?, notes = ?
        WHERE id = ?
      `).run(serviceName, serviceDate, serviceHours, notes, id);
      writeAudit(db, req, {
        module: "maintenance",
        action: "service_history.update",
        entity_type: "maintenance_service_history",
        entity_id: String(id),
        before: cur,
        after: { ...cur, service_name: serviceName, service_date: serviceDate, service_hours: serviceHours, notes },
      });
      return reply.send({ ok: true, id });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // DELETE /api/maintenance/history/backfill/:id
  app.delete("/history/backfill/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
      const cur = db.prepare(`SELECT * FROM maintenance_service_history WHERE id = ?`).get(id);
      if (!cur) return reply.code(404).send({ ok: false, error: "backfill entry not found" });
      db.prepare(`DELETE FROM maintenance_service_history WHERE id = ?`).run(id);
      writeAudit(db, req, {
        module: "maintenance",
        action: "service_history.delete",
        entity_type: "maintenance_service_history",
        entity_id: String(id),
        before: cur,
      });
      return reply.send({ ok: true, id });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // =====================================================
  // AUTO-GENERATE SERVICE WORK ORDERS
  // POST /api/maintenance/generate?date=2026-02-27
  // =====================================================
  app.post("/generate", async (req, reply) => {
    try {
      const date = String(req.query?.date || "").trim();
      if (date && !isDate(date)) {
        return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
      }
      const nearDueHours = Math.max(1, Number(req.body?.near_due_hours || req.query?.near_due_hours || 50));
      const planIdsRaw = Array.isArray(req.body?.plan_ids) ? req.body.plan_ids : [];
      const requestedPlanIds = [...new Set(planIdsRaw.map((v) => Number(v)).filter((v) => Number.isInteger(v) && v > 0))];

      const plans = date
        ? db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              a.asset_code,
              a.asset_name,
              a.category,
              IFNULL((
                SELECT SUM(dh.hours_run)
                FROM daily_hours dh
                WHERE dh.asset_id = mp.asset_id
                  AND dh.is_used = 1
                  AND dh.hours_run > 0
                  AND dh.work_date <= ?
              ), 0) AS current_hours
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.is_standby = 0
              AND a.archived = 0
          `).all(date)
        : db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              a.asset_code,
              a.asset_name,
              a.category,
              IFNULL((
                SELECT SUM(dh.hours_run)
                FROM daily_hours dh
                WHERE dh.asset_id = mp.asset_id
                  AND dh.is_used = 1
                  AND dh.hours_run > 0
              ), 0) AS current_hours
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.is_standby = 0
              AND a.archived = 0
          `).all();

      const findOpenServiceWO = db.prepare(`
        SELECT id, status
        FROM work_orders
        WHERE source = 'service'
          AND reference_id = ?
          AND REPLACE(TRIM(LOWER(COALESCE(status, ''))), ' ', '_')
            NOT IN ('closed', 'completed', 'approved', 'cancelled')
        ORDER BY id DESC
        LIMIT 1
      `);

      const insertWO = db.prepare(`
        INSERT INTO work_orders (asset_id, source, reference_id, status)
        VALUES (?, 'service', ?, 'open')
      `);

      const currentByAsset = new Map();
      for (const plan of plans) {
        const assetId = Number(plan.asset_id || 0);
        if (!currentByAsset.has(assetId)) {
          currentByAsset.set(assetId, date ? Number(plan.current_hours || 0) : getAssetCurrentHours(assetId));
        }
      }
      const nextServices = buildDueListFromPlans(
        plans,
        (assetId) => Number(currentByAsset.get(Number(assetId)) || 0),
        nearDueHours,
      );

      const tx = db.transaction(() => {
        const created = [];
        const skipped = [];

        for (const p of nextServices) {
          const current = Number(p.current_hours || 0);
          const next_due = Number(p.next_due_hours || 0);
          const remaining = Number(p.remaining_hours || 0);
          const isOverdue = remaining <= 0;
          const isRequestedPlan = requestedPlanIds.includes(Number(p.plan_id || 0));
          const shouldCreate = requestedPlanIds.length ? isRequestedPlan : isOverdue;
          if (!shouldCreate) continue;
          const existing = findOpenServiceWO.get(p.plan_id);
          if (existing) {
            skipped.push({
              plan_id: Number(p.plan_id),
              asset_id: Number(p.asset_id),
              asset_code: p.asset_code,
              reason: "open_work_order_exists",
              work_order_id: Number(existing.id),
              work_order_status: String(existing.status || "open"),
            });
            continue;
          }

          const wo = insertWO.run(p.asset_id, p.plan_id);
          created.push({
            work_order_id: Number(wo.lastInsertRowid),
            plan_id: p.plan_id,
            asset_id: p.asset_id,
            service_name: p.service_name,
            current_hours: Number(current.toFixed(2)),
            next_due_hours: Number(next_due.toFixed(2))
          });
        }

        return { created, skipped };
      });

      const result = tx();
      const created = result.created;

      return reply.send({
        ok: true,
        near_due_hours: nearDueHours,
        requested_plan_count: requestedPlanIds.length,
        created_count: created.length,
        created,
        skipped: result.skipped,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // UPCOMING SERVICES PDF
  // GET /api/maintenance/due-upcoming.pdf?date=YYYY-MM-DD&near_due_hours=50&download=1
  // =====================================================
  app.get("/due-upcoming.pdf", async (req, reply) => {
    try {
      const date = String(req.query?.date || "").trim();
      if (date && !isDate(date)) {
        return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
      }
      const nearDueHours = Math.max(1, Number(req.query?.near_due_hours || 50));
      const withinHoursRaw = req.query?.within_hours;
      const withinHours = withinHoursRaw != null && String(withinHoursRaw).trim() !== ""
        ? Math.max(0, Number(withinHoursRaw))
        : null;
      const planIdsRaw = req.query?.plan_ids;
      const requestedPlanIds = Array.isArray(planIdsRaw)
        ? [...new Set(planIdsRaw.map((v) => Number(v)).filter((v) => Number.isInteger(v) && v > 0))]
        : [...new Set(String(planIdsRaw || "").split(/[,\s]+/).map((v) => Number(v)).filter((v) => Number.isInteger(v) && v > 0))];
      const asOfLabel = date || new Date().toISOString().slice(0, 10);
      const historyDays = 14;

      const rows = date
        ? db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              a.asset_code,
              a.asset_name,
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
              AND a.archived = 0
            ORDER BY a.asset_code ASC, mp.service_name ASC
          `).all(date)
        : db.prepare(`
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

      const getHours = (assetId) => {
        if (date) {
          return Number(rows.find((r) => Number(r.asset_id) === Number(assetId))?.current_hours ?? getAssetCurrentHours(assetId));
        }
        return getAssetCurrentHours(assetId);
      };

      const dueRows = enrichDueRowsWithEstimates(
        db,
        buildDueListFromPlans(rows, getHours, nearDueHours),
        { as_of: asOfLabel, history_days: historyDays },
      )
        .map((r) => ({
          plan_id: Number(r.plan_id || 0),
          asset_code: r.asset_code,
          asset_name: r.asset_name,
          service_name: r.service_name,
          current_hours: Number(r.current_hours || 0),
          next_due_hours: Number(r.next_due_hours || 0),
          remaining_hours: Number(r.remaining_hours || 0),
          estimated_service_date: r.estimated_service_date,
          status: r.status,
        }))
        .filter((r) => {
          if (requestedPlanIds.length) return requestedPlanIds.includes(Number(r.plan_id || 0));
          if (withinHours != null && Number.isFinite(withinHours)) {
            return Number(r.remaining_hours || 0) <= withinHours;
          }
          return true;
        })
        .sort((a, b) => Number(a.remaining_hours || 0) - Number(b.remaining_hours || 0));

      const pdfScopeLabel = requestedPlanIds.length
        ? `Selected equipment only (${requestedPlanIds.length})`
        : withinHours != null && Number.isFinite(withinHours)
          ? `Remaining <= ${withinHours.toFixed(0)}h`
          : "All active service schedules";

      const pdf = await buildPdfBuffer(
        (doc) => {
          sectionTitle(doc, "Upcoming Services");
          doc
            .font("Helvetica")
            .fontSize(10)
            .text(
              `As of: ${asOfLabel} | ${pdfScopeLabel} | <= ${nearDueHours.toFixed(0)}h flagged ALMOST DUE | Est. date from ${historyDays}-day avg usage`,
            );
          doc.moveDown(0.4);

          table(
            doc,
            [
              { key: "asset_code", label: "Asset", width: 0.1 },
              { key: "asset_name", label: "Name", width: 0.16 },
              { key: "service_name", label: "Service", width: 0.15 },
              { key: "current_hours", label: "Current", width: 0.1, align: "right" },
              { key: "next_due_hours", label: "Next Due", width: 0.1, align: "right" },
              { key: "remaining_hours", label: "Remaining", width: 0.1, align: "right" },
              { key: "estimated_service_date", label: "Est. Date", width: 0.11 },
              { key: "status", label: "Status", width: 0.12 },
            ],
            dueRows.length
              ? dueRows.map((r) => ({
                  ...r,
                  current_hours: Number(r.current_hours || 0).toFixed(1),
                  next_due_hours: Number(r.next_due_hours || 0).toFixed(1),
                  remaining_hours: Number(r.remaining_hours || 0).toFixed(1),
                  estimated_service_date: r.estimated_service_date || "—",
                }))
              : [
                  {
                    asset_code: "-",
                    asset_name: requestedPlanIds.length
                      ? "No upcoming services for selected equipment"
                      : withinHours != null && Number.isFinite(withinHours)
                        ? `No services with remaining <= ${withinHours.toFixed(0)}h`
                        : "No upcoming services",
                    service_name: "-",
                    current_hours: "-",
                    next_due_hours: "-",
                    remaining_hours: "-",
                    estimated_service_date: "-",
                    status: "-",
                  },
                ]
          );
        },
        {
          title: "IRONLOG",
          subtitle: "Maintenance Upcoming Services",
          rightText: `As of: ${asOfLabel}`,
          layout: "landscape",
        }
      );

      const isDownload = String(req.query?.download || "").trim() === "1";
      reply.header("Cache-Control", "no-store, no-cache, must-revalidate");
      reply.header("Pragma", "no-cache");
      reply.header("Content-Type", "application/pdf");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="AML_Upcoming_Services_${asOfLabel}.pdf"`
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });
}
