// IRONLOG/api/routes/reports/stores.routes.js — Stock monitor, stock movement and part order reports.
// Registered by routes/reports.routes.js; shared helpers arrive through ctx.
import ExcelJS from "exceljs";
import path from "node:path";
import { buildPdfBuffer, kvGrid, sectionTitle, table, tryDrawLogo } from "../../utils/pdfGenerator.js";
import { createManagementSummary, styleManagementDetailSheet } from "../../utils/managementWorkbook.js";
import { db } from "../../db/client.js";

export default function registerStoresRoutes(app, ctx) {
  const {
    compactCell,
    fetchStoresPartOrdersForReport,
    fmtNum,
    hasColumn,
    parsePartOrdersPeriod,
    partOrderStatusLabel,
    todayYmd,
  } = ctx;

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
  // STOCK MOVEMENTS PDF (period ledger)
  // =========================
  // GET /api/reports/stock-movements.pdf?date_from=YYYY-MM-DD&date_to=YYYY-MM-DD&part_code=&download=1
  app.get("/stock-movements.pdf", async (req, reply) => {
    function parseDateOnly(s) {
      const t = String(s || "").trim();
      return /^\d{4}-\d{2}-\d{2}$/.test(t) ? t : null;
    }
    function normalizeStockReportPartFilter(raw) {
      let s = String(raw || "").trim().replace(/\s+/g, " ");
      if (!s) return "";
      const dashIdx = s.indexOf(" - ");
      if (dashIdx > 0) s = s.slice(0, dashIdx).trim();
      return s;
    }

    const date_from = parseDateOnly(req.query?.date_from);
    const date_to = parseDateOnly(req.query?.date_to);
    if (!date_from || !date_to) {
      return reply.code(400).send({ error: "date_from and date_to are required (YYYY-MM-DD)" });
    }
    if (date_from > date_to) {
      return reply.code(400).send({ error: "date_from must be on or before date_to" });
    }

    const part_filter = normalizeStockReportPartFilter(req.query?.part_code);
    const download = String(req.query?.download || "").trim() === "1";

    const smDateCol = hasColumn("stock_movements", "created_at") ? "sm.created_at" : "sm.movement_date";
    const smDateTimeExpr = `datetime(${smDateCol})`;

    const startDt = `${date_from} 00:00:00`;
    const endDt = `${date_to} 23:59:59`;

    const where = [`${smDateTimeExpr} >= datetime(?)`, `${smDateTimeExpr} <= datetime(?)`];
    const params = [startDt, endDt];

    if (part_filter) {
      const exactPart = db.prepare(`
        SELECT id FROM parts WHERE UPPER(TRIM(part_code)) = UPPER(TRIM(?))
      `).get(part_filter);
      if (exactPart) {
        where.push("sm.part_id = ?");
        params.push(Number(exactPart.id));
      } else {
        where.push("instr(LOWER(IFNULL(p.part_code,'')), LOWER(?)) > 0");
        params.push(part_filter);
      }
    }

    const whereSql = where.join(" AND ");

    const summaryRow = db.prepare(`
      SELECT
        COUNT(*) AS movement_count,
        IFNULL(SUM(CASE WHEN sm.quantity > 0 THEN sm.quantity ELSE 0 END), 0) AS qty_in,
        IFNULL(SUM(CASE WHEN sm.quantity < 0 THEN ABS(sm.quantity) ELSE 0 END), 0) AS qty_out,
        IFNULL(SUM(sm.quantity), 0) AS net_qty
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      WHERE ${whereSql}
    `).get(...params);

    const maxRows = 5000;
    const hasBin = hasColumn("stock_movements", "bin_id");
    const binJoin = hasBin ? "LEFT JOIN stock_bins b ON b.id = sm.bin_id" : "";
    const binSelect = hasBin ? "b.bin_code" : "NULL";

    const rows = db.prepare(`
      SELECT
        sm.id,
        ${smDateCol} AS movement_at,
        sm.movement_type,
        sm.quantity,
        sm.reference,
        l.location_code,
        ${binSelect} AS bin_code,
        p.part_code,
        p.part_name
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      LEFT JOIN stock_locations l ON l.id = sm.location_id
      ${binJoin}
      WHERE ${whereSql}
      ORDER BY ${smDateTimeExpr} DESC, sm.id DESC
      LIMIT ${maxRows}
    `).all(...params).map((r) => ({
      ...r,
      quantity: Number(r.quantity || 0),
    }));

    const totalMatching = Number(
      db.prepare(`
        SELECT COUNT(*) AS c
        FROM stock_movements sm
        JOIN parts p ON p.id = sm.part_id
        WHERE ${whereSql}
      `).get(...params)?.c || 0,
    );

    const truncated = totalMatching > rows.length;

    const summary = {
      movement_count: Number(summaryRow?.movement_count || 0),
      qty_in: Number(summaryRow?.qty_in || 0),
      qty_out: Number(summaryRow?.qty_out || 0),
      net_qty: Number(summaryRow?.net_qty || 0),
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Stock movements summary");
        kvGrid(
          doc,
          [
            { k: "Period", v: `${date_from} → ${date_to}` },
            { k: "Part filter", v: part_filter || "All parts" },
            { k: "Movements", v: fmtNum(summary.movement_count, 0) },
            { k: "Qty in", v: fmtNum(summary.qty_in, 2) },
            { k: "Qty out", v: fmtNum(summary.qty_out, 2) },
            { k: "Net", v: fmtNum(summary.net_qty, 2) },
            {
              k: "Rows in PDF",
              v: truncated ? `${rows.length} of ${totalMatching} (cap ${maxRows})` : String(rows.length || 0),
            },
          ],
          2,
        );

        sectionTitle(doc, "Movements");
        table(
          doc,
          [
            { key: "movement_at", label: "When", width: 0.14 },
            { key: "part_code", label: "Part", width: 0.11 },
            { key: "part_name", label: "Name", width: 0.22 },
            { key: "qty", label: "Qty", width: 0.07, align: "right" },
            { key: "type", label: "Type", width: 0.07 },
            { key: "loc", label: "Loc", width: 0.08 },
            { key: "bin", label: "Bin", width: 0.07 },
            { key: "reference", label: "Reference", width: 0.24 },
          ],
          rows.length
            ? rows.map((r) => ({
                movement_at: compactCell(String(r.movement_at || "").replace("T", " "), 22),
                part_code: compactCell(r.part_code, 14),
                part_name: compactCell(r.part_name, 34),
                qty: fmtNum(r.quantity, 2),
                type: String(r.movement_type || ""),
                loc: compactCell(r.location_code || "—", 10),
                bin: compactCell(r.bin_code || "", 8),
                reference: compactCell(r.reference || "", 42),
              }))
            : [
                {
                  movement_at: "-",
                  part_code: "-",
                  part_name: "No movements in period",
                  qty: "-",
                  type: "-",
                  loc: "-",
                  bin: "-",
                  reference: "-",
                },
              ],
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Stock movements report",
        rightText: part_filter ? `Filter: ${part_filter}` : "All parts",
        showPageNumbers: true,
      },
    );

    const safePart = part_filter ? String(part_filter).replace(/[^\w.-]+/g, "_").slice(0, 40) : "";

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Stock_Movements_${date_from}_${date_to}${safePart ? `_${safePart}` : ""}_${todayYmd()}.pdf"`,
      )
      .send(pdf);
  });

  // GET /api/reports/part-orders.pdf?start=&end=&status=&download=1
  app.get("/part-orders.pdf", async (req, reply) => {
    const period = parsePartOrdersPeriod(req);
    if (period.error) return reply.code(400).send({ error: period.error });
    const download = String(req.query?.download || "").trim() === "1";
    const { start, end, site_code, status } = period;
    const { rows, summary } = fetchStoresPartOrdersForReport(period);

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Parts purchases forecast");
        kvGrid(
          doc,
          [
            { k: "Site", v: site_code },
            { k: "Period", v: `${start} → ${end}` },
            { k: "Status filter", v: status ? partOrderStatusLabel(status) : "All active" },
            { k: "On order value", v: `${fmtNum(summary.on_order.value, 2)} (${summary.on_order.count} lines)` },
            { k: "Warehouse ready value", v: `${fmtNum(summary.warehouse_ready.value, 2)} (${summary.warehouse_ready.count} lines)` },
            { k: "In transit value", v: `${fmtNum(summary.in_transit.value, 2)} (${summary.in_transit.count} lines)` },
            { k: "Arrived value", v: `${fmtNum(summary.arrived.value, 2)} (${summary.arrived.count} lines)` },
            { k: "Pending forecast value", v: fmtNum(summary.total_pending, 2) },
            { k: "Total period value", v: fmtNum(summary.total_forecast, 2) },
          ],
          2,
        );

        sectionTitle(doc, "Purchase lines");
        table(
          doc,
          [
            { key: "order_date", label: "Ordered", width: 0.09 },
            { key: "part_code", label: "Part", width: 0.1 },
            { key: "part_name", label: "Description", width: 0.18 },
            { key: "qty", label: "Qty", width: 0.06, align: "right" },
            { key: "unit_cost", label: "Unit cost", width: 0.08, align: "right" },
            { key: "line_total", label: "Line cost", width: 0.08, align: "right" },
            { key: "supplier", label: "Supplier", width: 0.11 },
            { key: "po", label: "PO", width: 0.07 },
            { key: "req", label: "Req #", width: 0.08 },
            { key: "eta", label: "ETA", width: 0.08 },
            { key: "status", label: "Status", width: 0.11 },
          ],
          rows.length
            ? rows.map((r) => ({
                order_date: compactCell(r.order_date, 12),
                part_code: compactCell(r.part_code || "—", 14),
                part_name: compactCell(r.part_name, 30),
                qty: fmtNum(r.qty, 2),
                unit_cost: fmtNum(r.unit_cost, 2),
                line_total: fmtNum(r.line_total, 2),
                supplier: compactCell(r.supplier_name || "—", 18),
                po: compactCell(r.po_number || "—", 10),
                req: compactCell(r.requisition_number || "—", 12),
                eta: compactCell(r.expected_arrival_date || "—", 12),
                status: partOrderStatusLabel(r.status),
              }))
            : [
                {
                  order_date: "-",
                  part_code: "-",
                  part_name: "No purchases in this period",
                  qty: "-",
                  unit_cost: "-",
                  line_total: "-",
                  supplier: "-",
                  po: "-",
                  req: "-",
                  eta: "-",
                  status: "-",
                },
              ],
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Parts order & warehouse status",
        rightText: `${site_code} · ${start} → ${end}`,
        showPageNumbers: true,
      },
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="IRONLOG_Parts_Purchases_${start}_${end}_${todayYmd()}.pdf"`,
      )
      .send(pdf);
  });

  // GET /api/reports/part-orders.xlsx?start=&end=&status=
  app.get("/part-orders.xlsx", async (req, reply) => {
    const period = parsePartOrdersPeriod(req);
    if (period.error) return reply.code(400).send({ error: period.error });
    const { start, end, site_code, status } = period;
    const { rows, summary } = fetchStoresPartOrdersForReport(period);

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();
    createManagementSummary(wb, {
      title: "IRONLOG Parts Purchases Report",
      periodLabel: `Reporting period: ${start} to ${end}`,
      cards: [
        { label: "ON ORDER VALUE", value: summary.on_order.value, numFmt: "#,##0.00" },
        { label: "WAREHOUSE READY VALUE", value: summary.warehouse_ready.value, numFmt: "#,##0.00" },
        { label: "IN TRANSIT VALUE", value: summary.in_transit.value, numFmt: "#,##0.00" },
        { label: "ARRIVED VALUE", value: summary.arrived.value, numFmt: "#,##0.00" },
        { label: "TOTAL PERIOD VALUE", value: summary.total_forecast, numFmt: "#,##0.00" },
      ],
      scopeLines: [
        `Site: ${site_code}. Status filter: ${status ? partOrderStatusLabel(status) : "all active"}.`,
        `${summary.on_order.count} on-order, ${summary.warehouse_ready.count} warehouse-ready, ${summary.in_transit.count} in-transit, and ${summary.arrived.count} arrived line items in the selected period. Amounts retain the currency shown on each line.`,
      ],
    });

    const ws = wb.addWorksheet("Purchases");
    ws.columns = [
      { header: "Order date", key: "order_date", width: 12 },
      { header: "Part code", key: "part_code", width: 14 },
      { header: "Description", key: "part_name", width: 28 },
      { header: "Qty", key: "qty", width: 10 },
      { header: "Unit cost", key: "unit_cost", width: 12 },
      { header: "Line total", key: "line_total", width: 12 },
      { header: "Currency", key: "currency", width: 10 },
      { header: "Supplier", key: "supplier_name", width: 20 },
      { header: "PO number", key: "po_number", width: 14 },
      { header: "Requisition #", key: "requisition_number", width: 14 },
      { header: "Invoice #", key: "invoice_number", width: 14 },
      { header: "Location", key: "current_location", width: 18 },
      { header: "Warehouse", key: "warehouse_code", width: 12 },
      { header: "Warehouse date", key: "warehouse_date", width: 14 },
      { header: "Days waiting", key: "warehouse_waiting_days", width: 13 },
      { header: "Qty at warehouse", key: "supplier_qty_received", width: 17 },
      { header: "Outstanding qty", key: "supplier_outstanding_qty", width: 16 },
      { header: "Sales order", key: "sales_order", width: 16 },
      { header: "Asset", key: "asset_code", width: 13 },
      { header: "ETA on site", key: "expected_arrival_date", width: 14 },
      { header: "Arrived date", key: "arrived_date", width: 14 },
      { header: "Status", key: "status_label", width: 14 },
      { header: "Notes", key: "notes", width: 30 },
      { header: "Created by", key: "created_by", width: 16 },
    ];
    for (const r of rows) {
      ws.addRow({
        ...r,
        status_label: partOrderStatusLabel(r.status),
      });
    }
    ["qty", "supplier_qty_received", "supplier_outstanding_qty", "unit_cost", "line_total"].forEach((key) => {
      ws.getColumn(key).numFmt = "#,##0.00";
    });
    styleManagementDetailSheet(ws, {
      title: "Parts order & warehouse status detail",
      subtitle: `Reporting period: ${start} to ${end} · Site: ${site_code}`,
      frozenColumns: 2,
      numberFormats: { qty: "#,##0.00", supplier_qty_received: "#,##0.00", supplier_outstanding_qty: "#,##0.00", unit_cost: "#,##0.00", line_total: "#,##0.00" },
    });

    const buffer = await wb.xlsx.writeBuffer();
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header(
        "Content-Disposition",
        `attachment; filename="IRONLOG_Parts_Purchases_${start}_${end}.xlsx"`,
      )
      .send(Buffer.from(buffer));
  });
}
