// IRONLOG/api/routes/maintenance/weekly.routes.js — Borris weekly plan, weekly forum and weekly inspection roster.
// Registered by routes/maintenance.routes.js; shared helpers arrive through ctx.
import { buildDueListFromPlans, meterUnitForAsset } from "../../utils/serviceSchedule.js";
import { buildPdfBuffer, pdfBodyTop, sectionTitle, table } from "../../utils/pdfGenerator.js";
import { buildWeeklyMaintenancePlan } from "../../utils/weeklyMaintenancePlan.js";
import { db } from "../../db/client.js";
import { getPdfReportBranding } from "../../utils/reportSettings.js";
import { isDate } from "../../utils/request.js";

export default function registerWeeklyRoutes(app, ctx) {
  const {
    addDaysYmd,
    addWeeklyInspectionSlot,
    buildUpcomingServiceCostForecasts,
    buildWeeklyForumSummary,
    buildWeeklyInspectionCalendarData,
    clearWeeklyInspectionRoster,
    copyWeeklyInspectionDay,
    drawWeeklyInspectionCalendarPdfGrid,
    ensureWeeklyInspectionSchema,
    getAssetCurrentHours,
    getAssetCurrentHoursInfo,
    isMonth,
    normalizeEquipCategory,
    updateWeeklyInspectionSlotStatus,
    wiEnsureBodySpace,
    wiFormatMinutesPdf,
  } = ctx;

  // =====================================================
  // WEEKLY FORUM SUMMARY (cross-functional alignment)
  // GET /api/maintenance/weekly-forum/summary?start=YYYY-MM-DD&end=YYYY-MM-DD&near_due_hours=50
  // =====================================================
  // Read-only Borris plan using the same rotating service rules as the queue.
  app.get('/weekly-plan', async (req, reply) => {
    try {
      const today = new Intl.DateTimeFormat('en-CA', { timeZone: 'Africa/Johannesburg', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
      const plans = db.prepare(`SELECT mp.*, mp.id AS plan_id, a.asset_code, a.asset_name, a.category
        FROM maintenance_plans mp JOIN assets a ON a.id=mp.asset_id
        WHERE mp.active=1 AND a.active=1 AND a.is_standby=0 AND a.archived=0`).all();
      const due = buildDueListFromPlans(plans, id => getAssetCurrentHours(id), 50);
      const usageQuery = db.prepare(`SELECT SUM(day_run) AS total_run, COUNT(*) AS day_count,
        SUM(CASE WHEN day_run < 0 OR day_run > 24 THEN 1 ELSE 0 END) AS invalid_days
        FROM (SELECT work_date, SUM(hours_run) AS day_run FROM daily_hours
          WHERE asset_id=? AND is_used=1 AND work_date BETWEEN date(?, '-13 days') AND ? GROUP BY work_date)`);
      const forecasts = buildUpcomingServiceCostForecasts(db, due.map(r => ({...r, last_service_hours: r.next_due_hours-r.interval_hours})), {maxRemainingHours: Number.MAX_SAFE_INTEGER});
      const byPlan = new Map(forecasts.map(r => [r.plan_id, r.forecast]));
      const rows = due.map(r => ({...r, meter_unit: meterUnitForAsset(r.asset_code),
        meter_source: getAssetCurrentHoursInfo(r.asset_id).source, usage: usageQuery.get(r.asset_id, today, today), forecast: byPlan.get(r.plan_id)}));
      return reply.send(buildWeeklyMaintenancePlan(rows, today));
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ok:false, error:'Unable to build weekly maintenance plan'});
    }
  });

  app.get("/weekly-forum/summary", async (req, reply) => {
    try {
      const data = await buildWeeklyForumSummary(req.query || {});
      return reply.send(data);
    } catch (err) {
      req.log.error(err);
      return reply.code(Number(err?.statusCode || 500)).send({ ok: false, error: err.message || String(err) });
    }
  });

  // =====================================================
  // WEEKLY FORUM PDF
  // GET /api/maintenance/weekly-forum.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&near_due_hours=50&download=1
  // =====================================================
  app.get("/weekly-forum.pdf", async (req, reply) => {
    try {
      const data = await buildWeeklyForumSummary(req.query || {});
      const start = String(data?.range?.start || "");
      const end = String(data?.range?.end || "");
      const isDownload = String(req.query?.download || "").trim() === "1";

      const pdf = await buildPdfBuffer(
        (doc) => {
          sectionTitle(doc, "Weekly Forum Summary");
          table(
            doc,
            [
              { key: "metric", label: "Metric", width: 0.62 },
              { key: "value", label: "Value", width: 0.38, align: "right" },
            ],
            [
              { metric: "Range", value: `${start} to ${end}` },
              { metric: "Open Work Orders", value: Number(data?.kpis?.open_work_orders || 0) },
              { metric: "Upcoming Services Flagged", value: Number(data?.kpis?.upcoming_services_flagged || 0) },
              {
                metric: "Stores parts (excl. oil/lube SKUs)",
                value: Number(data?.costs?.stores_parts_cost || 0).toFixed(2),
              },
              {
                metric: "Oil cost — lube log entries",
                value: Number(data?.costs?.stores_oil_from_logs || 0).toFixed(2),
              },
              {
                metric: "Oil cost — WO stock (oil/lube lines)",
                value: Number(data?.costs?.stores_oil_from_work_orders || 0).toFixed(2),
              },
              {
                metric: "Stores oil total",
                value: Number(data?.costs?.stores_oil_cost || 0).toFixed(2),
              },
              { metric: "Maintenance Labor Cost", value: Number(data?.costs?.maintenance_labor_cost || 0).toFixed(2) },
              { metric: "Weekly Total Cost", value: Number(data?.costs?.weekly_total_cost || 0).toFixed(2) },
              { metric: "Upcoming Service Forecast Cost", value: Number(data?.costs?.upcoming_service_forecast_cost || 0).toFixed(2) },
            ]
          );

          const actualRows = Array.isArray(data?.period_actuals_by_asset) ? data.period_actuals_by_asset : [];
          sectionTitle(doc, "Period actuals by equipment (historical — selected date range)");
          table(
            doc,
            [
              { key: "machine", label: "Machine", width: 0.34 },
              { key: "parts", label: "Parts $", width: 0.12, align: "right" },
              { key: "lubes", label: "Lubes $", width: 0.12, align: "right" },
              { key: "labor", label: "Labor $", width: 0.12, align: "right" },
              { key: "total", label: "Total $", width: 0.14, align: "right" },
              { key: "wos", label: "Closed WOs", width: 0.16, align: "right" },
            ],
            actualRows.length
              ? actualRows.map((r) => ({
                  machine: `${String(r.asset_code || "-")} - ${String(r.asset_name || "-")}`,
                  parts: Number(r.parts_cost || 0).toFixed(2),
                  lubes: Number(r.lubes_total_cost || 0).toFixed(2),
                  labor: Number(r.labor_cost || 0).toFixed(2),
                  total: Number(r.period_total_cost || 0).toFixed(2),
                  wos: String(r.closed_work_orders ?? 0),
                }))
              : [
                  {
                    machine: "No equipment consumption recorded in range",
                    parts: "-",
                    lubes: "-",
                    labor: "-",
                    total: "-",
                    wos: "-",
                  },
                ]
          );

          sectionTitle(doc, "Upcoming Services Forecast");
          const rows = Array.isArray(data?.upcoming_services) ? data.upcoming_services : [];
          table(
            doc,
            [
              { key: "machine", label: "Machine", width: 0.18 },
              { key: "service", label: "Service", width: 0.15 },
              { key: "current", label: "Current", width: 0.07, align: "right" },
              { key: "next", label: "Next Due", width: 0.07, align: "right" },
              { key: "remain", label: "Remaining", width: 0.07, align: "right" },
              { key: "status", label: "Status", width: 0.09 },
              { key: "oil", label: "Avg Oil Qty", width: 0.07, align: "right" },
              { key: "oil_cost", label: "Avg Oil $", width: 0.08, align: "right" },
              { key: "parts", label: "Avg Parts Qty", width: 0.07, align: "right" },
              { key: "parts_cost", label: "Avg Parts $", width: 0.08, align: "right" },
              { key: "kit", label: "Est Kit Cost", width: 0.07, align: "right" },
            ],
            rows.length
              ? rows.map((r) => ({
                  machine: `${String(r.asset_code || "-")} - ${String(r.asset_name || "-")}`,
                  service: String(r.service_name || "-"),
                  current: Number(r.current_hours || 0).toFixed(1),
                  next: Number(r.next_due_hours || 0).toFixed(1),
                  remain: Number(r.remaining_hours || 0).toFixed(1),
                  status: String(r.status || "-"),
                  oil: Number(r?.forecast?.avg_oil_qty || 0).toFixed(1),
                  oil_cost: Number(r?.forecast?.avg_oil_cost || 0).toFixed(2),
                  parts: Number(r?.forecast?.avg_parts_qty || 0).toFixed(1),
                  parts_cost: Number(r?.forecast?.avg_parts_cost || 0).toFixed(2),
                  kit: Number(r?.forecast?.est_service_kit_cost || 0).toFixed(2),
                }))
              : [{
                  machine: "-",
                  service: "No upcoming services within threshold",
                  current: "-",
                  next: "-",
                  remain: "-",
                  status: "-",
                  oil: "-",
                  oil_cost: "-",
                  parts: "-",
                  parts_cost: "-",
                  kit: "-",
                }]
          );
        },
        {
          title: "IRONLOG",
          subtitle: "Weekly Forum",
          rightText: `${start} to ${end}`,
          showPageNumbers: true,
          layout: "landscape",
        }
      );

      reply.header("Content-Type", "application/pdf");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="AML_Weekly_Forum_${end}.pdf"`
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(Number(err?.statusCode || 500)).send({ ok: false, error: err.message || String(err) });
    }
  });

  // =====================================================
  // WEEKLY INSPECTION CALENDAR
  // =====================================================
  app.get("/weekly-inspections/candidate-assets", async (req, reply) => {
    try {
      ensureWeeklyInspectionSchema();
      const days = Math.max(7, Math.min(120, Number(req.query?.days || 30) || 30));
      const since = addDaysYmd(new Date().toISOString().slice(0, 10), -days);
      const rows = db.prepare(`
        SELECT
          a.id AS asset_id,
          a.asset_code,
          a.asset_name,
          a.category,
          MAX(dh.work_date) AS last_used_date,
          ROUND(SUM(CASE WHEN dh.work_date >= ? THEN COALESCE(dh.hours_run, 0) ELSE 0 END), 1) AS recent_hours
        FROM assets a
        INNER JOIN daily_hours dh ON dh.asset_id = a.id
        WHERE COALESCE(a.active, 1) = 1
          AND COALESCE(a.archived, 0) = 0
          AND COALESCE(a.is_standby, 0) = 0
          AND dh.work_date >= ?
          AND COALESCE(dh.is_used, 1) = 1
          AND COALESCE(dh.hours_run, 0) > 0
        GROUP BY a.id, a.asset_code, a.asset_name, a.category
        ORDER BY a.category ASC, a.asset_code ASC
      `).all(since, since);
      const groups = {};
      for (const r of rows) {
        const cat = normalizeEquipCategory(r.category);
        if (!groups[cat]) groups[cat] = [];
        groups[cat].push(r);
      }
      return reply.send({
        ok: true,
        since,
        days,
        assets: rows,
        groups,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/weekly-inspections/calendar", async (req, reply) => {
    try {
      const month = String(req.query?.month || "").trim();
      const weekStart = String(req.query?.week_start || "").trim();
      if (month && !isMonth(month)) {
        return reply.code(400).send({ ok: false, error: "month must be YYYY-MM" });
      }
      if (weekStart && !isDate(weekStart)) {
        return reply.code(400).send({ ok: false, error: "week_start must be YYYY-MM-DD" });
      }
      return reply.send(buildWeeklyInspectionCalendarData(req.query || {}));
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/weekly-inspections/assets", async (_req, reply) => {
    try {
      const rows = db.prepare(`
        SELECT
          wia.id,
          wia.asset_id,
          wia.notes,
          wia.sort_order,
          wia.active,
          COALESCE(wia.est_minutes, 30) AS est_minutes,
          a.asset_code,
          a.asset_name
        FROM weekly_inspection_assets wia
        JOIN assets a ON a.id = wia.asset_id
        WHERE COALESCE(wia.active, 1) = 1
        ORDER BY COALESCE(wia.sort_order, 0), a.asset_code ASC
      `).all();
      return reply.send({ ok: true, assets: rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/weekly-inspections/assets", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const notes = String(req.body?.notes || "").trim();
      const est_minutes = Math.max(5, Number(req.body?.est_minutes ?? 30) || 30);
      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      const asset = db.prepare(`
        SELECT id, asset_code, asset_name
        FROM assets
        WHERE id = ?
          AND COALESCE(active, 1) = 1
          AND COALESCE(archived, 0) = 0
      `).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });
      const existing = db.prepare(`SELECT id FROM weekly_inspection_assets WHERE asset_id = ?`).get(asset_id);
      if (existing) {
        db.prepare(`
          UPDATE weekly_inspection_assets
          SET active = 1, notes = ?, est_minutes = ?, updated_at = datetime('now')
          WHERE asset_id = ?
        `).run(notes, est_minutes, asset_id);
        return reply.send({ ok: true, id: Number(existing.id), asset_id, reactivated: true });
      }
      const maxSort = Number(db.prepare(`SELECT COALESCE(MAX(sort_order), 0) AS m FROM weekly_inspection_assets`).get()?.m || 0);
      const ins = db.prepare(`
        INSERT INTO weekly_inspection_assets (asset_id, notes, sort_order, est_minutes, active, created_at, updated_at)
        VALUES (?, ?, ?, ?, 1, datetime('now'), datetime('now'))
      `).run(asset_id, notes, maxSort + 1, est_minutes);
      return reply.send({ ok: true, id: Number(ins.lastInsertRowid), asset_id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/weekly-inspections/assets/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
      const row = db.prepare(`
        SELECT id, asset_id, notes, COALESCE(est_minutes, 30) AS est_minutes
        FROM weekly_inspection_assets
        WHERE id = ? AND COALESCE(active, 1) = 1
      `).get(id);
      if (!row) return reply.code(404).send({ ok: false, error: "Schedule row not found" });
      const notes = req.body?.notes != null ? String(req.body.notes || "").trim() : String(row.notes || "").trim();
      const est_minutes = req.body?.est_minutes != null
        ? Math.max(5, Number(req.body.est_minutes) || 30)
        : Number(row.est_minutes || 30);
      db.prepare(`
        UPDATE weekly_inspection_assets
        SET notes = ?, est_minutes = ?, updated_at = datetime('now')
        WHERE id = ?
      `).run(notes, est_minutes, id);
      db.prepare(`
        UPDATE weekly_inspection_slots
        SET est_minutes = ?, updated_at = datetime('now')
        WHERE asset_id = ?
      `).run(est_minutes, Number(row.asset_id));
      return reply.send({ ok: true, id, asset_id: Number(row.asset_id), notes, est_minutes });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/weekly-inspections/slots", async (req, reply) => {
    try {
      const planned_date = String(req.body?.planned_date || "").trim();
      const asset_id = Number(req.body?.asset_id || 0);
      const est_minutes = req.body?.est_minutes != null
        ? Math.max(5, Number(req.body.est_minutes) || 30)
        : undefined;
      const slot = addWeeklyInspectionSlot({ planned_date, asset_id, est_minutes });
      return reply.send({ ok: true, slot });
    } catch (err) {
      req.log.error(err);
      return reply.code(400).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/weekly-inspections/slots/copy-day", async (req, reply) => {
    try {
      const from_date = String(req.body?.from_date || "").trim();
      const to_date = String(req.body?.to_date || "").trim();
      const result = copyWeeklyInspectionDay(from_date, to_date);
      return reply.send({ ok: true, ...result });
    } catch (err) {
      req.log.error(err);
      return reply.code(400).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.delete("/weekly-inspections/slots/:id", async (req, reply) => {
    try {
      ensureWeeklyInspectionSchema();
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid slot id" });
      const row = db.prepare(`SELECT id, planned_date, asset_id FROM weekly_inspection_slots WHERE id = ?`).get(id);
      if (!row) return reply.code(404).send({ ok: false, error: "Inspection slot not found" });
      db.prepare(`DELETE FROM weekly_inspection_slots WHERE id = ?`).run(id);
      return reply.send({ ok: true, id, planned_date: row.planned_date, asset_id: Number(row.asset_id) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.delete("/weekly-inspections/assets/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
      const row = db.prepare(`SELECT id, asset_id FROM weekly_inspection_assets WHERE id = ?`).get(id);
      if (!row) return reply.code(404).send({ ok: false, error: "Schedule row not found" });
      db.prepare(`UPDATE weekly_inspection_assets SET active = 0, updated_at = datetime('now') WHERE id = ?`).run(id);
      return reply.send({ ok: true, id, asset_id: Number(row.asset_id) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.delete("/weekly-inspections/roster", async (req, reply) => {
    try {
      const clearSlots = String(req.query?.clear_slots ?? "1").trim() !== "0";
      const result = clearWeeklyInspectionRoster({ clear_slots: clearSlots });
      return reply.send({ ok: true, ...result, clear_slots: clearSlots });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/weekly-inspections/entries", async (req, reply) => {
    try {
      const slot_id = Number(req.body?.slot_id || 0);
      const asset_id = Number(req.body?.asset_id || 0);
      const planned_date = String(req.body?.planned_date || "").trim();
      const status = String(req.body?.status || "pending").trim().toLowerCase();
      const inspector_name = String(req.body?.inspector_name || "").trim();
      const slot = updateWeeklyInspectionSlotStatus({
        slot_id: slot_id || undefined,
        asset_id: asset_id || undefined,
        planned_date,
        status,
        inspector_name,
      });
      return reply.send({ ok: true, slot });
    } catch (err) {
      req.log.error(err);
      return reply.code(400).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/weekly-inspections.pdf", async (req, reply) => {
    try {
      const data = buildWeeklyInspectionCalendarData(req.query || {});
      const isDownload = String(req.query?.download || "").trim() === "1";
      const compliance = data.compliance || {};
      const rosterAssets = Array.isArray(data.assets) ? data.assets : [];
      const weeklyGaps = Array.isArray(compliance.weekly_gaps) ? compliance.weekly_gaps : [];
      const monthLabel = String(data.month || "");
      const branding = getPdfReportBranding(db);
      const pdf = await buildPdfBuffer(
        (doc) => {
          doc.y = pdfBodyTop(doc, { siteName: branding.site_name });
          doc.font("Helvetica").fontSize(10).fillColor("#334155");
          doc.text(
            `Released: ${Number(compliance.done_count ?? 0)} / ${Number(compliance.total_slots ?? 0)} · Overdue: ${Number(compliance.not_released_count ?? 0)} · Planned: ${wiFormatMinutesPdf(compliance.est_minutes_total || 0)}`,
            doc.page.margins.left,
            doc.y,
            { width: doc.page.width - doc.page.margins.left - doc.page.margins.right },
          );
          doc.moveDown(0.55);
          drawWeeklyInspectionCalendarPdfGrid(doc, data, { siteName: branding.site_name });
          wiEnsureBodySpace(doc, 36, branding.site_name);
          if (rosterAssets.length || weeklyGaps.length) {
            if (rosterAssets.length) {
              sectionTitle(doc, "Workshop roster");
              doc.font("Helvetica").fontSize(8).fillColor("#334155");
              doc.text(
                rosterAssets
                  .map((a) => `${String(a.asset_code || "-")} (${Number(a.est_minutes || 30)} min default)`)
                  .join("  ·  "),
                { lineGap: 2 },
              );
            }
            if (weeklyGaps.length) {
              wiEnsureBodySpace(doc, 48, branding.site_name);
              doc.moveDown(0.35);
              sectionTitle(doc, "Missing weekly workshop visit");
              doc.fontSize(8).fillColor("#b45309");
              const gapLines = weeklyGaps.slice(0, 40).map((g) =>
                `${String(g.asset_code || "-")} — week ${String(g.week_start || "").slice(5)} to ${String(g.week_end || "").slice(5)}`,
              );
              doc.text(gapLines.join("\n"), { lineGap: 2 });
              if (weeklyGaps.length > 40) doc.text(`…and ${weeklyGaps.length - 40} more`, { lineGap: 2 });
            }
          }
          wiEnsureBodySpace(doc, 20, branding.site_name);
          doc.moveDown(0.35);
          doc.fontSize(8).fillColor("#64748b");
          doc.text("Legend: REL = released  |  SKIP = skipped  |  PEN = pending");
        },
        {
          title: "IRONLOG",
          subtitle: "Workshop Inspection Calendar",
          rightText: monthLabel,
          showPageNumbers: true,
          layout: "landscape",
        },
      );
      reply.header("Content-Type", "application/pdf");
      reply.header("Cache-Control", "no-store");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="IRONLOG_Workshop_Inspections_${monthLabel || "calendar"}.pdf"`
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/weekly-forum/review-notes", async (req, reply) => {
    try {
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      if (!isDate(start) || !isDate(end)) {
        return reply.code(400).send({ ok: false, error: "start and end must be YYYY-MM-DD" });
      }
      const rows = db.prepare(`
        SELECT id, period_start, period_end, area, weekly_finding, action_owner, due_date, created_at, updated_at
        FROM weekly_forum_review_notes
        WHERE period_start = ? AND period_end = ?
        ORDER BY CASE area
          WHEN 'Downtime' THEN 0
          WHEN 'Repairs' THEN 1
          WHEN 'Costs' THEN 2
          ELSE 3
        END ASC, id ASC
      `).all(start, end);
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/weekly-forum/review-notes", async (req, reply) => {
    try {
      const start = String(req.body?.start || "").trim();
      const end = String(req.body?.end || "").trim();
      const area = String(req.body?.area || "").trim();
      const weekly_finding = String(req.body?.weekly_finding || "").trim();
      const action_owner = String(req.body?.action_owner || "").trim();
      const due_date = String(req.body?.due_date || "").trim() || null;
      const allowedAreas = ["Downtime", "Repairs", "Costs", "Other"];
      if (!isDate(start) || !isDate(end) || start > end) {
        return reply.code(400).send({ ok: false, error: "start and end must be valid dates, with start before end" });
      }
      if (!allowedAreas.includes(area)) {
        return reply.code(400).send({ ok: false, error: "area must be Downtime, Repairs, Costs, or Other" });
      }
      if (!weekly_finding || !action_owner) {
        return reply.code(400).send({ ok: false, error: "weekly finding and action / owner are required" });
      }
      if (due_date && !isDate(due_date)) {
        return reply.code(400).send({ ok: false, error: "due_date must be YYYY-MM-DD" });
      }
      const existing = db.prepare(`
        SELECT id FROM weekly_forum_review_notes
        WHERE period_start = ? AND period_end = ? AND area = ?
      `).get(start, end, area);
      db.prepare(`
        INSERT INTO weekly_forum_review_notes (
          period_start, period_end, area, weekly_finding, action_owner, due_date, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
        ON CONFLICT(period_start, period_end, area) DO UPDATE SET
          weekly_finding = excluded.weekly_finding,
          action_owner = excluded.action_owner,
          due_date = excluded.due_date,
          updated_at = datetime('now')
      `).run(start, end, area, weekly_finding, action_owner, due_date);
      const row = db.prepare(`
        SELECT id, period_start, period_end, area, weekly_finding, action_owner, due_date, created_at, updated_at
        FROM weekly_forum_review_notes
        WHERE period_start = ? AND period_end = ? AND area = ?
      `).get(start, end, area);
      return reply.send({ ok: true, id: Number(row?.id || existing?.id || 0), row });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.delete("/weekly-forum/review-notes/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
      const result = db.prepare(`DELETE FROM weekly_forum_review_notes WHERE id = ?`).run(id);
      if (!result.changes) return reply.code(404).send({ ok: false, error: "weekly review note not found" });
      return reply.send({ ok: true, id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/weekly-forum/actions", async (req, reply) => {
    try {
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const status = String(req.query?.status || "").trim().toLowerCase();
      if (start && !isDate(start)) return reply.code(400).send({ ok: false, error: "start must be YYYY-MM-DD" });
      if (end && !isDate(end)) return reply.code(400).send({ ok: false, error: "end must be YYYY-MM-DD" });

      const where = [];
      const params = [];
      if (start) {
        where.push("action_date >= ?");
        params.push(start);
      }
      if (end) {
        where.push("action_date <= ?");
        params.push(end);
      }
      if (status) {
        where.push("LOWER(COALESCE(status,'open')) = ?");
        params.push(status);
      }

      const rows = db.prepare(`
        SELECT
          id, action_date, department, action_item, owner_name, due_date,
          status, notes, created_at, updated_at
        FROM weekly_forum_actions
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY
          CASE LOWER(COALESCE(status, 'open'))
            WHEN 'open' THEN 0
            WHEN 'in_progress' THEN 1
            WHEN 'blocked' THEN 2
            ELSE 3
          END ASC,
          COALESCE(due_date, action_date) ASC,
          id DESC
      `).all(...params);
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/weekly-forum/parts", async (req, reply) => {
    try {
      const hasPartsTable = Boolean(
        db.prepare(`
          SELECT 1
          FROM sqlite_master
          WHERE type = 'table' AND name = 'parts'
          LIMIT 1
        `).get()
      );
      const rows = hasPartsTable
        ? db.prepare(`
            SELECT
              p.id,
              p.part_code,
              p.part_name,
              COALESCE(SUM(sm.quantity), 0) AS on_hand,
              COALESCE((
                SELECT COALESCE(sm2.unit_cost_usd, sm2.cost_input, 0)
                FROM stock_movements sm2
                WHERE sm2.part_id = p.id
                  AND COALESCE(sm2.unit_cost_usd, sm2.cost_input, 0) > 0
                ORDER BY sm2.id DESC
                LIMIT 1
              ), 0) AS latest_unit_cost
            FROM parts p
            LEFT JOIN stock_movements sm ON sm.part_id = p.id
            GROUP BY p.id
            ORDER BY p.part_code ASC
            LIMIT 1500
          `).all()
        : [];
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/weekly-forum/forecast-inputs", async (req, reply) => {
    try {
      const rows = db.prepare(`
        SELECT id, plan_id, oil_part_code, oil_qty, parts_part_code, parts_qty, items_json, notes,
          COALESCE(labor_total, 0) AS labor_total,
          COALESCE(all_in_total, 0) AS all_in_total, updated_at
        FROM weekly_forum_service_inputs
        ORDER BY plan_id ASC
      `).all();
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/weekly-forum/forecast-inputs", async (req, reply) => {
    try {
      const plan_id = Number(req.body?.plan_id || 0);
      const items = Array.isArray(req.body?.items) ? req.body.items : [];
      const oil_part_code = String(req.body?.oil_part_code || "").trim() || null;
      const oil_qty = Math.max(0, Number(req.body?.oil_qty || 0));
      const parts_part_code = String(req.body?.parts_part_code || "").trim() || null;
      const parts_qty = Math.max(0, Number(req.body?.parts_qty || 0));
      const notes = String(req.body?.notes || "").trim() || null;
      const labor_total = Math.max(0, Number(req.body?.labor_total || 0));
      const all_in_total = Math.max(0, Number(req.body?.all_in_total || 0));
      const normalizedItems = items
        .map((it) => ({
          type: String(it?.type || "part").toLowerCase() === "oil" ? "oil" : "part",
          part_code: String(it?.part_code || "").trim(),
          qty: Math.max(0, Number(it?.qty || 0)),
        }))
        .filter((it) => it.part_code && it.qty > 0);
      const items_json = JSON.stringify(normalizedItems);
      if (!plan_id) return reply.code(400).send({ ok: false, error: "plan_id is required" });
      if (!normalizedItems.length && labor_total <= 0 && all_in_total <= 0) {
        return reply.code(400).send({ ok: false, error: "Add store items, labor, or an all-in planned cost" });
      }
      db.prepare(`
        INSERT INTO weekly_forum_service_inputs (
          plan_id, oil_part_code, oil_qty, parts_part_code, parts_qty, items_json, notes, labor_total, all_in_total, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
        ON CONFLICT(plan_id) DO UPDATE SET
          oil_part_code = excluded.oil_part_code,
          oil_qty = excluded.oil_qty,
          parts_part_code = excluded.parts_part_code,
          parts_qty = excluded.parts_qty,
          items_json = excluded.items_json,
          notes = excluded.notes,
          labor_total = excluded.labor_total,
          all_in_total = excluded.all_in_total,
          updated_at = datetime('now')
      `).run(plan_id, oil_part_code, oil_qty, parts_part_code, parts_qty, items_json, notes, labor_total, all_in_total);
      return reply.send({ ok: true, plan_id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/weekly-forum/actions", async (req, reply) => {
    try {
      const action_date = String(req.body?.action_date || "").trim() || new Date().toISOString().slice(0, 10);
      const department = String(req.body?.department || "").trim();
      const action_item = String(req.body?.action_item || "").trim();
      const owner_name = String(req.body?.owner_name || "").trim();
      const due_date = String(req.body?.due_date || "").trim() || null;
      const status = String(req.body?.status || "open").trim().toLowerCase() || "open";
      const notes = String(req.body?.notes || "").trim() || null;

      if (!isDate(action_date)) return reply.code(400).send({ ok: false, error: "action_date must be YYYY-MM-DD" });
      if (!department) return reply.code(400).send({ ok: false, error: "department is required" });
      if (!action_item) return reply.code(400).send({ ok: false, error: "action_item is required" });
      if (!owner_name) return reply.code(400).send({ ok: false, error: "owner_name is required" });
      if (due_date && !isDate(due_date)) return reply.code(400).send({ ok: false, error: "due_date must be YYYY-MM-DD" });

      const allowedStatuses = ["open", "in_progress", "blocked", "done"];
      if (!allowedStatuses.includes(status)) {
        return reply.code(400).send({ ok: false, error: "status must be open|in_progress|blocked|done" });
      }

      const ins = db.prepare(`
        INSERT INTO weekly_forum_actions (
          action_date, department, action_item, owner_name, due_date, status, notes, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(action_date, department, action_item, owner_name, due_date, status, notes);
      return reply.send({ ok: true, id: Number(ins.lastInsertRowid) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/weekly-forum/actions/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
      const existing = db.prepare(`SELECT id FROM weekly_forum_actions WHERE id = ?`).get(id);
      if (!existing) return reply.code(404).send({ ok: false, error: "action not found" });

      const department = req.body?.department != null ? String(req.body.department).trim() : undefined;
      const action_item = req.body?.action_item != null ? String(req.body.action_item).trim() : undefined;
      const owner_name = req.body?.owner_name != null ? String(req.body.owner_name).trim() : undefined;
      const due_date = req.body?.due_date != null ? (String(req.body.due_date).trim() || null) : undefined;
      const status = req.body?.status != null ? String(req.body.status).trim().toLowerCase() : undefined;
      const notes = req.body?.notes != null ? (String(req.body.notes).trim() || null) : undefined;

      if (due_date !== undefined && due_date && !isDate(due_date)) {
        return reply.code(400).send({ ok: false, error: "due_date must be YYYY-MM-DD" });
      }
      if (status !== undefined) {
        const allowedStatuses = ["open", "in_progress", "blocked", "done"];
        if (!allowedStatuses.includes(status)) {
          return reply.code(400).send({ ok: false, error: "status must be open|in_progress|blocked|done" });
        }
      }

      db.prepare(`
        UPDATE weekly_forum_actions
        SET
          department = COALESCE(?, department),
          action_item = COALESCE(?, action_item),
          owner_name = COALESCE(?, owner_name),
          due_date = COALESCE(?, due_date),
          status = COALESCE(?, status),
          notes = COALESCE(?, notes),
          updated_at = datetime('now')
        WHERE id = ?
      `).run(department ?? null, action_item ?? null, owner_name ?? null, due_date ?? null, status ?? null, notes ?? null, id);

      return reply.send({ ok: true, id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
}
