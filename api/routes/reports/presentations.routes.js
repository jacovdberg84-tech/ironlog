// IRONLOG/api/routes/reports/presentations.routes.js — Maintenance master, executive and GM presentations and documents.
// Registered by routes/reports.routes.js; shared helpers arrive through ctx.
import PptxGenJS from "pptxgenjs";
import fs from "node:fs";
import { buildBudgetMeetingDocxBuffer } from "../../utils/budgetMeetingDocx.js";
import { buildMechanicLaborDetail, buildMonthlyOperatingActuals } from "../../utils/monthlyOperatingCosts.js";
import { buildPlantHireLines, prevMonth as hirePrevMonth } from "../../utils/plantHire.js";
import { db } from "../../db/client.js";
import { getOperatingBudgetAmount } from "../../utils/monthlyBudget.js";

export default function registerPresentationsRoutes(app, ctx) {
  const {
    buildMaintenanceExecutiveDeck,
    compactCell,
    fmtNum,
    generateMaintenanceMaster,
    getSiteCode,
    hasTable,
    isMonth,
    isYmd,
    monthRange,
    resolveMaintenancePeriod,
  } = ctx;

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
        requestHeaders: req.headers,
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
        const generated = await generateMaintenanceMaster(report_type, site_code, {
          month,
          start,
          end,
          requestHeaders: req.headers,
        });
        row = {
          report_type,
          label: generated.label,
          file_path: generated.file_path,
        };
      } catch (e) {
        return reply.code(500).send({ ok: false, error: e?.message || "Failed to generate maintenance presentation" });
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
    const siteCodeFromQuery = String(req.query?.site_code || "").trim().toLowerCase();
    const site_code = siteCodeFromQuery || getSiteCode(req);
    const buffer = await buildMaintenanceExecutiveDeck({
      period,
      label,
      site_code,
      requestHeaders: req.headers,
    });
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Maintenance_Executive_${label}.pptx"`)
      .send(buffer);
  });

  app.get("/gm-upcoming-costs.pptx", async (req, reply) => {
    const resolved = resolveMaintenancePeriod(req);
    if (!resolved) {
      return reply.code(400).send({ error: "Provide month=YYYY-MM or start/end=YYYY-MM-DD" });
    }
    const { period, label } = resolved;
    const siteCodeFromQuery = String(req.query?.site_code || "").trim().toLowerCase();
    const siteCodeHeader = getSiteCode(req);
    // Window-open downloads may omit headers; include main-site fallback for legacy flows.
    const siteCandidates = Array.from(new Set(
      siteCodeFromQuery
        ? [siteCodeFromQuery]
        : (siteCodeHeader === "default" ? ["main", "default"] : [siteCodeHeader])
    ));
    const site_code = siteCandidates[0] || "main";

    const hasStoresPartOrders = hasTable("stores_part_orders");
    const siteMarks = siteCandidates.map(() => "?").join(", ");
    const orderRows = hasStoresPartOrders
      ? db.prepare(`
          SELECT
            id,
            part_code,
            part_name,
            qty,
            unit_cost,
            supplier_name,
            po_number,
            order_date,
            expected_arrival_date,
            arrived_date,
            status,
            notes
          FROM stores_part_orders
          WHERE LOWER(TRIM(COALESCE(site_code, 'main'))) IN (${siteMarks})
            AND DATE(order_date) BETWEEN DATE(?) AND DATE(?)
          ORDER BY DATE(order_date) DESC, id DESC
          LIMIT 200
        `).all(...siteCandidates, period.start, period.end)
      : [];

    const statusRows = {
      on_order: orderRows.filter((r) => String(r.status || "").toLowerCase() === "on_order"),
      in_transit: orderRows.filter((r) => String(r.status || "").toLowerCase() === "in_transit"),
      arrived: orderRows.filter((r) => String(r.status || "").toLowerCase() === "arrived"),
    };
    const statusTotals = Object.fromEntries(
      Object.entries(statusRows).map(([k, rows]) => [
        k,
        {
          qty: rows.reduce((s, r) => s + Number(r.qty || 0), 0),
          value: rows.reduce((s, r) => s + Number(r.qty || 0) * Number(r.unit_cost || 0), 0),
          lines: rows.length,
        },
      ]),
    );

    const q = new URLSearchParams();
    q.set("start", period.start);
    q.set("end", period.end);
    q.set("near_due_hours", "50");
    q.set("predictive_horizon_hours", "100");
    q.set("checklist_fail_threshold", "2");
    q.set("fuel_variance_threshold", "15");
    const injected = await app.inject({
      method: "GET",
      url: `/api/maintenance/insights?${q.toString()}`,
      headers: {
        "x-user-name": String(req.headers?.["x-user-name"] || "system"),
        "x-user-role": String(req.headers?.["x-user-role"] || "admin"),
        "x-user-roles": String(req.headers?.["x-user-roles"] || "admin"),
        "x-site-code": site_code,
      },
    });
    if (injected.statusCode >= 400) {
      let payload = {};
      try { payload = JSON.parse(String(injected.payload || "{}")); } catch {}
      return reply.code(injected.statusCode).send(payload?.error ? payload : { ok: false, error: "Failed to build maintenance insights data" });
    }
    const insights = JSON.parse(String(injected.payload || "{}"));
    const maintenanceRows = Array.isArray(insights?.parts_planning?.upcoming_cost_forecasts)
      ? insights.parts_planning.upcoming_cost_forecasts
      : [];

    const maintenanceTotal = maintenanceRows.reduce(
      (s, r) => s + Number(r?.forecast?.est_total_cost || 0),
      0,
    );
    const maintenanceKit = maintenanceRows.reduce(
      (s, r) => s + Number(r?.forecast?.est_service_kit_cost || 0),
      0,
    );
    const maintenanceLabor = maintenanceRows.reduce(
      (s, r) => s + Number(r?.forecast?.est_labor_cost || 0),
      0,
    );

    const pptx = new PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "IRONLOG";
    pptx.subject = "GM upcoming costs report";
    pptx.title = `GM Upcoming Costs - ${label}`;

    const intro = pptx.addSlide();
    intro.addText("GM Upcoming Costs Report", { x: 0.4, y: 0.35, w: 12.4, h: 0.55, fontSize: 26, bold: true });
    intro.addText(`Period: ${period.start} to ${period.end} | Site: ${site_code}`, { x: 0.4, y: 0.95, w: 12.4, h: 0.35, fontSize: 11 });
    intro.addText(
      [
        { text: `Parts on order total: $${fmtNum(statusTotals.on_order?.value || 0, 2)}\n`, options: { bold: true } },
        { text: `Parts in transit total: $${fmtNum(statusTotals.in_transit?.value || 0, 2)}\n`, options: { bold: true } },
        { text: `Parts arrived total: $${fmtNum(statusTotals.arrived?.value || 0, 2)}\n`, options: { bold: true } },
        { text: `Upcoming maintenance total: $${fmtNum(maintenanceTotal, 2)}\n`, options: { bold: true } },
      ],
      { x: 0.7, y: 1.75, w: 11.8, h: 2.2, fontSize: 16 }
    );

    const addPartsStatusSlide = (title, rows, totals) => {
      const s = pptx.addSlide();
      s.addText(title, { x: 0.4, y: 0.3, w: 12.4, h: 0.5, fontSize: 22, bold: true });
      s.addTable(
        [
          [
            { text: "Order Date", options: { bold: true } },
            { text: "Part", options: { bold: true } },
            { text: "Description", options: { bold: true } },
            { text: "Qty", options: { bold: true } },
            { text: "Unit $", options: { bold: true } },
            { text: "Line $", options: { bold: true } },
            { text: "Supplier", options: { bold: true } },
            { text: "PO #", options: { bold: true } },
            { text: "ETA/Arrived", options: { bold: true } },
          ],
          ...(rows.length
            ? rows.slice(0, 14).map((r) => [
                String(r.order_date || "-"),
                String(r.part_code || "-"),
                compactCell(String(r.part_name || "-"), 22),
                fmtNum(Number(r.qty || 0), 2),
                fmtNum(Number(r.unit_cost || 0), 2),
                fmtNum(Number(r.qty || 0) * Number(r.unit_cost || 0), 2),
                compactCell(String(r.supplier_name || "-"), 16),
                compactCell(String(r.po_number || "-"), 12),
                String(r.arrived_date || r.expected_arrival_date || "-"),
              ])
            : [["-", "-", "No rows in selected period", "-", "-", "-", "-", "-", "-"]]),
        ],
        { x: 0.35, y: 1.0, w: 12.6, h: 4.7, fontSize: 9.2, border: { pt: 1, color: "D0D0D0" } }
      );
      s.addText(
        `Total lines: ${Number(totals?.lines || 0)}   |   Total qty: ${fmtNum(Number(totals?.qty || 0), 2)}   |   Total cost: $${fmtNum(Number(totals?.value || 0), 2)}`,
        { x: 0.45, y: 5.95, w: 12.2, h: 0.35, fontSize: 12, bold: true }
      );
    };

    addPartsStatusSlide("Parts on Order", statusRows.on_order, statusTotals.on_order);
    addPartsStatusSlide("Parts in Transit", statusRows.in_transit, statusTotals.in_transit);
    addPartsStatusSlide("Parts Arrived", statusRows.arrived, statusTotals.arrived);

    const maint = pptx.addSlide();
    maint.addText("Upcoming Maintenance Cost (from Maintenance Insights)", { x: 0.4, y: 0.3, w: 12.4, h: 0.5, fontSize: 20, bold: true });
    maint.addTable(
      [
        [
          { text: "Asset", options: { bold: true } },
          { text: "Service", options: { bold: true } },
          { text: "Remaining Hrs", options: { bold: true } },
          { text: "Status", options: { bold: true } },
          { text: "Kit $", options: { bold: true } },
          { text: "Labor $", options: { bold: true } },
          { text: "Total $", options: { bold: true } },
          { text: "Source", options: { bold: true } },
        ],
        ...(maintenanceRows.length
          ? maintenanceRows.slice(0, 14).map((r) => [
              `${String(r.asset_code || "-")} ${compactCell(String(r.asset_name || ""), 16)}`,
              compactCell(String(r.service_name || "-"), 20),
              fmtNum(Number(r.remaining_hours || 0), 1),
              String(r.status || "-"),
              fmtNum(Number(r?.forecast?.est_service_kit_cost || 0), 2),
              fmtNum(Number(r?.forecast?.est_labor_cost || 0), 2),
              fmtNum(Number(r?.forecast?.est_total_cost || 0), 2),
              compactCell(String(r?.forecast?.cost_source || "-").replace(/_/g, " "), 16),
            ])
          : [["-", "-", "-", "-", "-", "-", "-", "No upcoming maintenance rows"]]),
      ],
      { x: 0.35, y: 1.0, w: 12.6, h: 4.7, fontSize: 9.2, border: { pt: 1, color: "D0D0D0" } }
    );
    maint.addText(
      `Total rows: ${maintenanceRows.length}   |   Kit total: $${fmtNum(maintenanceKit, 2)}   |   Labor total: $${fmtNum(maintenanceLabor, 2)}   |   Upcoming maintenance total: $${fmtNum(maintenanceTotal, 2)}`,
      { x: 0.45, y: 5.95, w: 12.2, h: 0.35, fontSize: 12, bold: true }
    );

    const buffer = await pptx.write({ outputType: "nodebuffer" });
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
      .header("Content-Disposition", `attachment; filename="IRONLOG_GM_Upcoming_Costs_${label}.pptx"`)
      .send(Buffer.from(buffer));
  });

  // GET /api/reports/gm-budget-meeting.docx?month=YYYY-MM&site_code=
  app.get("/gm-budget-meeting.docx", async (req, reply) => {
    const month = String(req.query?.month || "").trim();
    if (!isMonth(month)) {
      return reply.code(400).send({ error: "month (YYYY-MM) required" });
    }
    const siteCodeFromQuery = String(req.query?.site_code || "").trim().toLowerCase();
    const siteCodeHeader = getSiteCode(req);
    const site_code = siteCodeFromQuery || (siteCodeHeader === "default" ? "main" : siteCodeHeader);
    const prevLabel = hirePrevMonth(month);
    const period = monthRange(month);

    async function injectJson(path) {
      const injected = await app.inject({
        method: "GET",
        url: path,
        headers: {
          "x-user-name": String(req.headers?.["x-user-name"] || "system"),
          "x-user-role": String(req.headers?.["x-user-role"] || "admin"),
          "x-user-roles": String(req.headers?.["x-user-roles"] || "admin"),
          "x-site-code": site_code,
        },
      });
      if (injected.statusCode >= 400) return null;
      try {
        return JSON.parse(String(injected.payload || "{}"));
      } catch {
        return null;
      }
    }

    const [plantBudget] = await Promise.all([
      injectJson(`/api/finance/plant-hire-budget?period=${encodeURIComponent(month)}&site_code=${encodeURIComponent(site_code)}`),
    ]);

    const currentActuals = buildMonthlyOperatingActuals(db, month);
    const prevActuals = buildMonthlyOperatingActuals(db, prevLabel);
    const operatingBudgetRow = getOperatingBudgetAmount(db, month, site_code);

    const hasStoresPartOrders = hasTable("stores_part_orders");
    const orderRows = hasStoresPartOrders
      ? db.prepare(`
          SELECT qty, unit_cost, status
          FROM stores_part_orders
          WHERE LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
            AND DATE(order_date) BETWEEN DATE(?) AND DATE(?)
        `).all(site_code, period.start, period.end)
      : [];
    const sumStatus = (status) =>
      orderRows
        .filter((r) => String(r.status || "").toLowerCase() === status)
        .reduce((s, r) => s + Number(r.qty || 0) * Number(r.unit_cost || 0), 0);

    const q = new URLSearchParams();
    q.set("start", period.start);
    q.set("end", period.end);
    q.set("near_due_hours", "50");
    q.set("predictive_horizon_hours", "100");
    const insightsInjected = await app.inject({
      method: "GET",
      url: `/api/maintenance/insights?${q.toString()}`,
      headers: {
        "x-user-name": String(req.headers?.["x-user-name"] || "system"),
        "x-user-role": String(req.headers?.["x-user-role"] || "admin"),
        "x-user-roles": String(req.headers?.["x-user-roles"] || "admin"),
        "x-site-code": site_code,
      },
    });
    let maintenanceTotal = 0;
    if (insightsInjected.statusCode < 400) {
      try {
        const insights = JSON.parse(String(insightsInjected.payload || "{}"));
        const maintenanceRows = Array.isArray(insights?.parts_planning?.upcoming_cost_forecasts)
          ? insights.parts_planning.upcoming_cost_forecasts
          : [];
        maintenanceTotal = maintenanceRows.reduce(
          (s, r) => s + Number(r?.forecast?.est_total_cost || 0),
          0,
        );
      } catch { /* ignore */ }
    }

    const partsOnOrder = sumStatus("on_order");
    const partsInTransit = sumStatus("in_transit");
    const partsArrived = sumStatus("arrived");
    const upcomingTotal = partsOnOrder + partsInTransit + partsArrived + maintenanceTotal;
    const plantHireLines = buildPlantHireLines(db, month);
    const mechanicLabor = buildMechanicLaborDetail(db, month, site_code);

    const buffer = await buildBudgetMeetingDocxBuffer({
      periodLabel: month,
      prevPeriodLabel: prevLabel,
      siteCode: site_code,
      operatingBudget: Number(operatingBudgetRow.budget_amount || 0),
      currentActuals,
      prevActuals,
      plantHireBudget: Number(plantBudget?.budget_amount || 0),
      plantHireLines,
      mechanicLaborDetail: mechanicLabor.detail,
      mechanicLaborTotalHours: mechanicLabor.total_hours,
      mechanicLaborTotalCost: mechanicLabor.total_cost,
      mechanicLaborDefaultRate: mechanicLabor.default_rate,
      upcoming: {
        parts_on_order: partsOnOrder,
        parts_in_transit: partsInTransit,
        parts_arrived: partsArrived,
        maintenance_total: maintenanceTotal,
        upcoming_total: upcomingTotal,
      },
    });

    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Budget_Meeting_${month}.docx"`)
      .header("Cache-Control", "no-store")
      .send(Buffer.from(buffer));
  });
}
