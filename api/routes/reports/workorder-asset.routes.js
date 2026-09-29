// IRONLOG/api/routes/reports/workorder-asset.routes.js — Work order and asset history PDFs.
// Registered by routes/reports.routes.js; shared helpers arrive through ctx.
import path from "node:path";
import { buildPdfBuffer, ensurePageSpace, kvGrid, sectionTitle, table, tryDrawLogo } from "../../utils/pdfGenerator.js";
import { db } from "../../db/client.js";
import { isDate } from "../../utils/request.js";

export default function registerWorkorderAssetRoutes(app, ctx) {
  const {
    compactCell,
    completionNotesForPdf,
    costDefaults,
    fmtNum,
    jobFindingsTextForPdf,
    todayYmd,
    yn,
  } = ctx;

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
        ${optionalWoCol("job_description")},
        ${optionalWoCol("repair_progress")},
        ${optionalWoCol("repair_progress_at")},
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
    const jobFindingsText = jobFindingsTextForPdf(wo, breakdown);
    const completionNotesPdf = completionNotesForPdf(wo, jobFindingsText);

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
            v: completionNotesPdf,
          },
        ], 2);

        sectionTitle(doc, "Manual Work Execution (Artisan)");
        doc.font("Helvetica").fontSize(10);
        doc.text("Job Description / Findings:", { width: 460 });
        doc.moveDown(0.2);
        if (jobFindingsText) {
          doc.font("Helvetica").fontSize(9).text(jobFindingsText, { width: 460 });
          doc.moveDown(0.4);
        }
        const findingBlankLines = jobFindingsText ? 2 : 4;
        for (let i = 0; i < findingBlankLines; i++) {
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
}
