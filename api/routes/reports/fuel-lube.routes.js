// IRONLOG/api/routes/reports/fuel-lube.routes.js — Lube and fuel reports (benchmark, reconciliation, machine history).
// Registered by routes/reports.routes.js; shared helpers arrive through ctx.
import ExcelJS from "exceljs";
import path from "node:path";
import { aggregateFuelBenchmarkByCategory } from "../../utils/fuelBenchmarkAggregate.js";
import { buildMechanicLaborDetail } from "../../utils/monthlyOperatingCosts.js";
import { buildPdfBuffer, kvGrid, sectionTitle, table, tryDrawLogo } from "../../utils/pdfGenerator.js";
import { createManagementSummary, styleManagementDetailSheet } from "../../utils/managementWorkbook.js";
import { db } from "../../db/client.js";
import { fetchLubeMonthStockSnapshot } from "../../utils/lubeMonthStock.js";
import { fetchLubeUsageLines } from "../../utils/lubeUsageLines.js";
import { fuelBenchmarkAssetsInRangeSql, sqlFuelMetricModeExpr } from "../../utils/fuelMetricMode.js";
import { getMonthlyBudgetRow, getOperatingBudgetAmount } from "../../utils/monthlyBudget.js";
import { summarizeFuelBenchmarkRows } from "../../utils/fuelRunFromLogs.js";
import { assetConsumption, consumptionNote, ensureFuelBenchmarkSchema, sortBenchmarkRows } from "../../utils/fuelConsumption.js";
import { isDate } from "../../utils/request.js";

export default function registerFuelLubeRoutes(app, ctx) {
  const {
    addFuelBenchmarkByAssetWorksheet,
    calendarQuartersThrough,
    compactCell,
    fmtNum,
    monthsInRange,
    queryPeriodFleetCostTotals,
  } = ctx;
  ensureFuelBenchmarkSchema(db);

  // =========================
  // LUBE USAGE PDF
  // =========================
  // GET /api/reports/lube.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&download=1
  app.get("/lube.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const download = String(req.query?.download || "").trim() === "1";
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const defaultLubeCost = Number(
      db.prepare(`SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1`).get()?.value
    );
    const lubeUnitFallback = Number.isFinite(defaultLubeCost) && defaultLubeCost > 0 ? defaultLubeCost : 4.0;

    const rows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS total_lube_cost,
        COUNT(*) AS entries
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.id
      ORDER BY qty_total DESC, a.asset_code ASC
      LIMIT 400
    `).all(lubeUnitFallback, start, end).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      qty_total: Number(r.qty_total || 0),
      total_lube_cost: Number(r.total_lube_cost || 0),
      entries: Number(r.entries || 0),
    }));

    const detailRows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        CASE
          WHEN LOWER(TRIM(COALESCE(ol.oil_type, ''))) IN ('admin','supervisor','manager','stores','artisan','operator') THEN 'UNSPECIFIED'
          ELSE COALESCE(NULLIF(TRIM(ol.oil_type), ''), 'UNSPECIFIED')
        END AS part_number,
        COALESCE(p.part_name, '') AS lube_description,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS total_lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      LEFT JOIN parts p ON UPPER(TRIM(p.part_code)) = UPPER(TRIM(COALESCE(ol.oil_type, '')))
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.id, part_number, p.part_name
      ORDER BY a.asset_code ASC, qty_total DESC, part_number ASC
      LIMIT 1200
    `).all(lubeUnitFallback, start, end).map((r) => ({
      asset_code: String(r.asset_code || ""),
      asset_name: String(r.asset_name || ""),
      part_number: String(r.part_number || "UNSPECIFIED"),
      lube_description: String(r.lube_description || ""),
      qty_total: Number(r.qty_total || 0),
      total_lube_cost: Number(r.total_lube_cost || 0),
    }));

    const summary = db.prepare(`
      SELECT
        COALESCE(SUM(quantity), 0) AS qty_total,
        COALESCE(SUM(quantity * COALESCE(unit_cost, ?)), 0) AS total_lube_cost,
        COUNT(*) AS entries,
        COUNT(DISTINCT asset_id) AS assets
      FROM oil_logs
      WHERE log_date BETWEEN ? AND ?
    `).get(lubeUnitFallback, start, end);

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Lube Usage Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Total Qty", v: fmtNum(summary?.qty_total || 0, 1) },
          { k: "Total Cost", v: fmtNum(summary?.total_lube_cost || 0, 2) },
          { k: "Total Entries", v: fmtNum(summary?.entries || 0, 0) },
          { k: "Assets Logged", v: fmtNum(summary?.assets || 0, 0) },
        ], 2);

        sectionTitle(doc, "Lube Usage by Machine");
        table(
          doc,
          [
            { key: "asset_code", label: "Asset Code", width: 0.18 },
            { key: "asset_name", label: "Asset Name", width: 0.36 },
            { key: "entries", label: "Entries", width: 0.12, align: "right" },
            { key: "qty_total", label: "Qty Total", width: 0.16, align: "right" },
            { key: "total_lube_cost", label: "Total Cost", width: 0.18, align: "right" },
          ],
          rows.length
            ? rows.map((r) => ({
                asset_code: r.asset_code,
                asset_name: r.asset_name || "",
                entries: fmtNum(r.entries, 0),
                qty_total: fmtNum(r.qty_total, 1),
                total_lube_cost: fmtNum(r.total_lube_cost, 2),
              }))
            : [{ asset_code: "-", asset_name: "No lube usage in period", entries: "-", qty_total: "-", total_lube_cost: "-" }]
        );

        sectionTitle(doc, "Lube Usage Detail (Part Number / Description)");
        table(
          doc,
          [
            { key: "asset_code", label: "Asset Code", width: 0.12 },
            { key: "asset_name", label: "Asset Name", width: 0.24 },
            { key: "part_number", label: "Lube Part Number", width: 0.16 },
            { key: "lube_description", label: "Lube Description", width: 0.24 },
            { key: "qty_total", label: "Qty", width: 0.10, align: "right" },
            { key: "total_lube_cost", label: "Cost", width: 0.14, align: "right" },
          ],
          detailRows.length
            ? detailRows.map((r) => ({
                asset_code: r.asset_code,
                asset_name: r.asset_name || "",
                part_number: r.part_number || "UNSPECIFIED",
                lube_description: r.lube_description || "-",
                qty_total: fmtNum(r.qty_total, 1),
                total_lube_cost: fmtNum(r.total_lube_cost, 2),
              }))
            : [{
                asset_code: "-",
                asset_name: "-",
                part_number: "-",
                lube_description: "No lube detail rows in period",
                qty_total: "-",
                total_lube_cost: "-",
              }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Lube Usage Report",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
        layout: "landscape",
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Lube_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/lube-usage-by-asset.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD&month=YYYY-MM&location_code=LUBE
  // Sheet 1: oil_logs usage by asset + oil type. Sheet 2: month opening/closing store stock for lube SKUs.
  app.get("/lube-usage-by-asset.xlsx", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    let month = String(req.query?.month || "").trim();
    const location_code = String(req.query?.location_code || "").trim().toUpperCase();
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }
    if (!/^\d{4}-\d{2}$/.test(month)) month = start.slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(month)) {
      return reply.code(400).send({ error: "month could not be derived from start date" });
    }

    const defaultLubeCost = Number(
      db.prepare(`SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1`).get()?.value
    );
    const lubeUnitFallback = Number.isFinite(defaultLubeCost) && defaultLubeCost > 0 ? defaultLubeCost : 4.0;

    let location_id = null;
    let resolvedLocationCode = "";
    if (location_code) {
      const loc = db.prepare(`SELECT id, location_code FROM stock_locations WHERE UPPER(TRIM(location_code)) = ?`).get(location_code);
      if (!loc) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
      location_id = Number(loc.id);
      resolvedLocationCode = String(loc.location_code || location_code);
    }

    const usageLines = fetchLubeUsageLines(db, { start, end, lubeUnitFallback });
    const usageSummary = new Map();
    for (const line of usageLines) {
      // This sheet is the daily oil-log usage summary. Stores/work-order
      // movements are listed on the line sheet but must not be mixed into it.
      if (line.source !== "oil_log") continue;
      const key = [line.asset_code, line.part_code, line.part_name, line.lube_type].join("\u001F");
      const current = usageSummary.get(key) || {
        asset_code: line.asset_code,
        asset_name: line.asset_name,
        part_code: line.part_code,
        part_name: line.part_name,
        lube_type: line.lube_type,
        qty_total: 0,
        total_lube_cost: 0,
        entries: 0,
      };
      current.qty_total += Number(line.quantity || 0);
      current.total_lube_cost += Number(line.line_cost || 0);
      current.entries += 1;
      usageSummary.set(key, current);
    }
    const usageRows = [...usageSummary.values()].sort((a, b) =>
      String(a.asset_code).localeCompare(String(b.asset_code))
      || String(a.part_name).localeCompare(String(b.part_name))
      || Number(b.qty_total) - Number(a.qty_total)
    );

    let stockSnap;
    try {
      stockSnap = fetchLubeMonthStockSnapshot(db, { month, location_id });
    } catch (e) {
      return reply.code(400).send({ error: String(e.message || e) });
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();
    const totalQty = usageRows.reduce((sum, row) => sum + Number(row.qty_total || 0), 0);
    const totalCost = usageRows.reduce((sum, row) => sum + Number(row.total_lube_cost || 0), 0);
    const assetCount = new Set(usageRows.map((row) => String(row.asset_code || "")).filter(Boolean)).size;
    createManagementSummary(wb, {
      title: "IRONLOG Lube Usage Report",
      periodLabel: `Reporting period: ${start} to ${end}`,
      cards: [
        { label: "LUBE QUANTITY ISSUED", value: totalQty, numFmt: "#,##0.00" },
        { label: "ESTIMATED LUBE COST", value: totalCost, numFmt: "#,##0.00" },
        { label: "ASSETS SUPPLIED", value: assetCount, numFmt: "#,##0" },
        { label: "LUBE SKUS ISSUED", value: usageRows.length, numFmt: "#,##0" },
      ],
      scopeLines: [
        `Usage period: ${start} to ${end}. Store stock month: ${month}.`,
        `Location: ${resolvedLocationCode || "all locations"}. The detail tabs retain each issue and the store balance evidence.`,
      ],
    });
    const wsUsage = wb.addWorksheet("Usage by asset", { views: [{ state: "frozen", ySplit: 1 }] });
    wsUsage.columns = [
      { header: "Asset code", key: "asset_code", width: 14 },
      { header: "Asset name", key: "asset_name", width: 28 },
      { header: "Lube part no", key: "part_code", width: 16 },
      { header: "Description", key: "part_name", width: 36 },
      { header: "Type", key: "lube_type", width: 14 },
      { header: "Qty", key: "qty_total", width: 12 },
      { header: "Est. cost", key: "total_lube_cost", width: 12 },
      { header: "Log lines", key: "entries", width: 10 },
    ];
    for (const r of usageRows) {
      wsUsage.addRow({
        asset_code: r.asset_code,
        asset_name: r.asset_name,
        part_code: r.part_code,
        part_name: r.part_name || "",
        lube_type: r.lube_type,
        qty_total: Number(r.qty_total || 0),
        total_lube_cost: Number(r.total_lube_cost || 0),
        entries: Number(r.entries || 0),
      });
    }
    styleManagementDetailSheet(wsUsage, {
      title: "Lube usage by asset",
      subtitle: `Reporting period: ${start} to ${end}`,
      frozenColumns: 1,
      numberFormats: { qty_total: "#,##0.00", total_lube_cost: "#,##0.00", entries: "#,##0" },
    });

    const wsLines = wb.addWorksheet("Usage lines", { views: [{ state: "frozen", ySplit: 1 }] });
    wsLines.columns = [
      { header: "Date", key: "usage_date", width: 12 },
      { header: "Lube part no", key: "part_code", width: 16 },
      { header: "Description", key: "part_name", width: 32 },
      { header: "Type", key: "lube_type", width: 14 },
      { header: "Plant no", key: "asset_code", width: 14 },
      { header: "Machine", key: "asset_name", width: 28 },
      { header: "Machine hrs", key: "smr", width: 12 },
      { header: "Qty", key: "quantity", width: 10 },
      { header: "Unit cost", key: "unit_cost", width: 11 },
      { header: "Line cost", key: "line_cost", width: 11 },
      { header: "Source", key: "source", width: 14 },
      { header: "WO #", key: "work_order_id", width: 8 },
    ];
    for (const r of usageLines) {
      wsLines.addRow({
        usage_date: r.usage_date,
        part_code: r.part_code,
        part_name: r.part_name,
        lube_type: r.lube_type,
        asset_code: r.asset_code,
        asset_name: r.asset_name,
        smr: r.smr != null ? Number(r.smr) : "",
        quantity: Number(r.quantity || 0),
        unit_cost: Number(r.unit_cost || 0),
        line_cost: Number(r.line_cost || 0),
        source: r.source,
        work_order_id: r.work_order_id || "",
      });
    }
    styleManagementDetailSheet(wsLines, {
      title: "Lube issue lines",
      subtitle: `Reporting period: ${start} to ${end}`,
      frozenColumns: 2,
      numberFormats: { smr: "#,##0.0", quantity: "#,##0.00", unit_cost: "#,##0.00", line_cost: "#,##0.00" },
    });

    const wsStock = wb.addWorksheet("Store stock");
    wsStock.columns = [
      { header: "Stock code", key: "part_code", width: 16 },
      { header: "Description", key: "part_name", width: 34 },
      { header: "Min stock", key: "min_stock", width: 14 },
      { header: "Opening qty", key: "opening_qty", width: 14 },
      { header: "Closing qty", key: "closing_qty", width: 14 },
      { header: "Net movement", key: "net_month_movement", width: 16 },
    ];
    for (const r of stockSnap.rows) {
      wsStock.addRow(r);
    }
    styleManagementDetailSheet(wsStock, {
      title: "Lube store stock movement",
      subtitle: `Balance month: ${month} · Opening as of ${stockSnap.opening_as_of} · Closing as of ${stockSnap.closing_as_of}`,
      frozenColumns: 1,
      numberFormats: { min_stock: "#,##0.00", opening_qty: "#,##0.00", closing_qty: "#,##0.00", net_month_movement: "#,##0.00" },
    });

    const safeStart = start.replace(/[^\d-]/g, "");
    const safeEnd = end.replace(/[^\d-]/g, "");
    const buffer = await wb.xlsx.writeBuffer();
    return reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Lube_Usage_${safeStart}_to_${safeEnd}.xlsx"`)
      .send(buffer);
  });

  // =========================
  // FUEL BENCHMARK PDF
  // =========================
  // GET /api/reports/fuel-benchmark.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15&download=1
  app.get("/fuel-benchmark.pdf", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const toleranceInput = Number(req.query?.tolerance ?? 0.15);
    const tolerance = Number.isFinite(toleranceInput) ? Math.max(0, toleranceInput) : 0.15;
    const download = String(req.query?.download || "").trim() === "1";
    const modeFilter = String(req.query?.mode || "").trim().toLowerCase(); // 'km' | 'hours' | ''
    const assetFilter = String(req.query?.asset_code || "").trim().toLowerCase();
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const fuelByAsset = db.prepare(fuelBenchmarkAssetsInRangeSql()).all(start, end);
    const benchmarkRows = sortBenchmarkRows(fuelByAsset
      .map((r) => {
        const row = assetConsumption(db, r, start, end, tolerance);
        return { ...row, note: consumptionNote(row) };
      })
      .filter((r) => r.fuel_liters > 0)
      .filter((r) => (assetFilter ? String(r.asset_code || "").trim().toLowerCase() === assetFilter : true))
      .filter((r) => (modeFilter === "km" ? r.metric_mode === "km" : modeFilter === "hours" ? r.metric_mode === "hours" : true)));

    const rows = benchmarkRows;
    const categoryRows = assetFilter ? [] : aggregateFuelBenchmarkByCategory(
      rows.map((r) => ({ ...r, is_excessive: r.flag === "EXCESSIVE" })),
      tolerance
    );
    const displayRows = categoryRows.length ? categoryRows : rows;

    const summaryBase = summarizeFuelBenchmarkRows(rows);
    const summary = { ...summaryBase, categories: categoryRows.length };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Fuel Benchmark Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Tolerance", v: `${fmtNum(tolerance * 100, 1)}%` },
          { k: "Mode filter", v: modeFilter ? modeFilter : "all" },
          { k: "Categories", v: fmtNum(summary.categories, 0) },
          { k: "Assets (detail)", v: fmtNum(summary.assets, 0) },
          { k: "Excessive", v: fmtNum(summary.excessive, 0) },
          { k: "Fuel Total (L)", v: fmtNum(summary.fuel_liters, 2) },
          { k: "Hours Run", v: fmtNum(summary.hours_run, 2) },
          { k: "Avg L/hr", v: summary.avg_lph == null ? "-" : fmtNum(summary.avg_lph, 3) },
        ], 2);
        doc.font("Helvetica-Bold").fontSize(9).fillColor("#b91c1c").text("Flag legend: EXCESSIVE = above OEM by tolerance", { align: "left" });
        doc.font("Helvetica").fontSize(8.5).fillColor("#475569").text("Worked out fill to fill: each fill's litres over the meter movement since the previous good reading. Impossible readings are rejected; machines without an OEM benchmark are not flagged.", { align: "left" });
        doc.moveDown(0.4);
        doc.font("Helvetica").fontSize(9).fillColor("#111111");

        const byCategory = categoryRows.length > 0;
        sectionTitle(doc, byCategory ? "Fuel Used by Equipment Category" : "Fuel Benchmark by Machine");
        table(
          doc,
          byCategory
            ? [
                { key: "category", label: "Category", width: 0.22 },
                { key: "metric_mode", label: "Mode", width: 0.08, align: "center" },
                { key: "asset_count", label: "Units", width: 0.08, align: "center" },
                { key: "fuel_liters", label: "Fuel (L)", width: 0.14, align: "right" },
                { key: "run", label: "Run", width: 0.12, align: "right" },
                { key: "actual", label: "Actual", width: 0.10, align: "right" },
                { key: "oem", label: "OEM", width: 0.08, align: "right" },
                { key: "variance", label: "Variance", width: 0.08, align: "right" },
                { key: "flag", label: "Flag", width: 0.08, align: "center" },
              ]
            : [
                { key: "asset_code", label: "Asset", width: 0.09 },
                { key: "asset_name", label: "Name", width: 0.15 },
                { key: "metric_mode", label: "Mode", width: 0.06, align: "center" },
                { key: "fuel_liters", label: "Fuel (L)", width: 0.09, align: "right" },
                { key: "run", label: "Run", width: 0.09, align: "right" },
                { key: "actual", label: "Actual", width: 0.10, align: "right" },
                { key: "oem", label: "OEM", width: 0.07, align: "right" },
                { key: "variance", label: "Variance", width: 0.07, align: "right" },
                { key: "flag", label: "Flag", width: 0.08, align: "center" },
                { key: "note", label: "Notes", width: 0.20 },
              ],
          displayRows.length
            ? displayRows.map((r) => {
                if (byCategory) {
                  return {
                    category: r.category,
                    metric_mode: r.metric_mode,
                    asset_count: fmtNum(r.asset_count, 0),
                    fuel_liters: fmtNum(r.fuel_liters, 2),
                    run: r.metric_mode === "km" ? `${fmtNum(r.km_run, 2)} km` : `${fmtNum(r.hours_run, 2)} h`,
                    actual: r.metric_mode === "km"
                      ? (r.actual_km_per_l == null ? "-" : `${fmtNum(r.actual_km_per_l, 3)} km/L`)
                      : (r.actual_lph == null ? "-" : `${fmtNum(r.actual_lph, 3)} L/hr`),
                    oem: (r.metric_mode === "km" ? r.oem_km_per_l : r.oem_lph) == null ? "Not set" : fmtNum(r.metric_mode === "km" ? r.oem_km_per_l : r.oem_lph, 3),
                    variance: r.metric_mode === "km"
                      ? (r.variance_km_per_l == null ? "-" : fmtNum(r.variance_km_per_l, 3))
                      : (r.variance_lph == null ? "-" : fmtNum(r.variance_lph, 3)),
                    flag: r.flag,
                  };
                }
                return {
                  asset_code: r.asset_code,
                  asset_name: r.asset_name || "",
                  metric_mode: r.metric_mode,
                  fuel_liters: fmtNum(r.fuel_liters, 2),
                  run: r.metric_mode === "km" ? `${fmtNum(r.km_run, 2)} km` : `${fmtNum(r.hours_run, 2)} h`,
                  actual: r.metric_mode === "km"
                    ? (r.actual_km_per_l == null ? "-" : `${fmtNum(r.actual_km_per_l, 3)} km/L`)
                    : (r.actual_lph == null ? "-" : `${fmtNum(r.actual_lph, 3)} L/hr`),
                  oem: (r.metric_mode === "km" ? r.oem_km_per_l : r.oem_lph) == null ? "Not set" : fmtNum(r.metric_mode === "km" ? r.oem_km_per_l : r.oem_lph, 3),
                  variance: r.metric_mode === "km"
                    ? (r.variance_km_per_l == null ? "-" : fmtNum(r.variance_km_per_l, 3))
                    : (r.variance_lph == null ? "-" : fmtNum(r.variance_lph, 3)),
                  flag: r.flag,
                  note: r.note || "",
                };
              })
            : [{
                category: "-",
                asset_code: "-",
                asset_name: "No fuel benchmark data for period",
                metric_mode: "-",
                asset_count: "-",
                fuel_liters: "-",
                run: "-",
                actual: "-",
                oem: "-",
                variance: "-",
                flag: "-",
              }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Fuel Benchmark Report",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
        layout: "landscape",
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Fuel_Benchmark_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/fuel-benchmark.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15
  // GET /api/reports/fuel-reconciliation.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15&fuel_price=1.5&download=1
  app.get("/fuel-reconciliation.pdf", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const toleranceInput = Number(req.query?.tolerance ?? 0.15);
    const tolerance = Number.isFinite(toleranceInput) ? Math.max(0, toleranceInput) : 0.15;
    const fuelPriceInput = Number(req.query?.fuel_price ?? 0);
    const fuelPrice = Number.isFinite(fuelPriceInput) ? Math.max(0, fuelPriceInput) : 0;
    const download = String(req.query?.download || "").trim() === "1";
    const modeFilter = String(req.query?.mode || "").trim().toLowerCase();
    const assetFilter = String(req.query?.asset_code || "").trim().toLowerCase();
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const fuelByAsset = db.prepare(fuelBenchmarkAssetsInRangeSql()).all(start, end);

    const rows = fuelByAsset
      .map((asset) => {
        const r = assetConsumption(db, asset, start, end, tolerance);
        const mode = r.metric_mode;
        const km = r.km_run;
        const hours = r.hours_run;
        // All fuel dispensed against what the recorded run should have used; fuel
        // that could not be matched to meter readings is named in the reasons.
        const fuel = Number(r.fuel_liters || 0);
        const expected = mode === "km"
          ? (km > 0 && Number(r.oem_km_per_l || 0) > 0 ? km / Number(r.oem_km_per_l) : 0)
          : (hours > 0 && Number(r.oem_lph || 0) > 0 ? hours * Number(r.oem_lph) : 0);
        const variance = fuel - expected;
        const allowed = expected * tolerance;
        const unexplained = Math.max(0, variance - allowed);
        const variancePct = expected > 0 ? (variance / expected) * 100 : null;
        const reasons = [];
        const runBase = mode === "km" ? km : hours;
        if (runBase <= 0 && fuel > 0) reasons.push("fuel with no run basis");
        if (Number(r.fill_count || 0) < 2) reasons.push("low sample count");
        if (variancePct != null && variancePct > 50) reasons.push("variance > 50%");
        const unmatchedL = Number((r.fuel_liters - Number(r.basis_liters || 0)).toFixed(2));
        if (runBase > 0 && unmatchedL > 0) reasons.push(`${unmatchedL} L not matched to meter readings`);
        if (r.suspect_readings) reasons.push(`${r.suspect_readings} meter reading(s) rejected`);
        if (!r.oem_set) reasons.push("benchmark not set");
        return {
          asset_code: String(r.asset_code || ""),
          asset_name: String(r.asset_name || ""),
          metric_mode: mode,
          fuel_liters: Number(fuel.toFixed(2)),
          expected_liters: Number(expected.toFixed(2)),
          variance_liters: Number(variance.toFixed(2)),
          unexplained_liters: Number(unexplained.toFixed(2)),
          variance_pct: variancePct == null ? null : Number(variancePct.toFixed(1)),
          reason_text: reasons.join("; "),
        };
      })
      .filter((r) => r.fuel_liters > 0)
      .filter((r) => (assetFilter ? String(r.asset_code || "").trim().toLowerCase() === assetFilter : true))
      .filter((r) => (modeFilter === "km" ? r.metric_mode === "km" : modeFilter === "hours" ? r.metric_mode === "hours" : true))
      .sort((a, b) => Number(b.unexplained_liters || 0) - Number(a.unexplained_liters || 0));

    const totals = rows.reduce((acc, r) => {
      acc.actual += Number(r.fuel_liters || 0);
      acc.expected += Number(r.expected_liters || 0);
      acc.variance += Number(r.variance_liters || 0);
      acc.unexplained += Number(r.unexplained_liters || 0);
      if (String(r.reason_text || "").trim()) acc.anomalies += 1;
      return acc;
    }, { actual: 0, expected: 0, variance: 0, unexplained: 0, anomalies: 0 });
    totals.value = totals.unexplained * fuelPrice;

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);
        sectionTitle(doc, "Fuel Reconciliation Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Tolerance", v: `${fmtNum(tolerance * 100, 1)}%` },
          { k: "Fuel price / L", v: fmtNum(fuelPrice, 2) },
          { k: "Assets", v: fmtNum(rows.length, 0) },
          { k: "Anomaly assets", v: fmtNum(totals.anomalies, 0) },
          { k: "Actual liters", v: fmtNum(totals.actual, 2) },
          { k: "Expected liters", v: fmtNum(totals.expected, 2) },
          { k: "Variance liters", v: fmtNum(totals.variance, 2) },
          { k: "Estimated missing (unexplained) liters", v: fmtNum(totals.unexplained, 2) },
          { k: "Estimated missing value", v: fmtNum(totals.value, 2) },
        ], 2);

        sectionTitle(doc, "Top Reconciliation Exceptions");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.10 },
            { key: "name", label: "Name", width: 0.20 },
            { key: "mode", label: "Mode", width: 0.07, align: "center" },
            { key: "actual", label: "Actual L", width: 0.10, align: "right" },
            { key: "expected", label: "Expected L", width: 0.10, align: "right" },
            { key: "variance", label: "Variance L", width: 0.10, align: "right" },
            { key: "unexplained", label: "Unexplained L", width: 0.12, align: "right" },
            { key: "pct", label: "Var %", width: 0.07, align: "right" },
            { key: "reasons", label: "Reasons", width: 0.14 },
          ],
          rows.length
            ? rows.slice(0, 60).map((r) => ({
                asset: r.asset_code || "-",
                name: r.asset_name || "",
                mode: r.metric_mode,
                actual: fmtNum(r.fuel_liters, 2),
                expected: fmtNum(r.expected_liters, 2),
                variance: fmtNum(r.variance_liters, 2),
                unexplained: fmtNum(r.unexplained_liters, 2),
                pct: r.variance_pct == null ? "-" : fmtNum(r.variance_pct, 1),
                reasons: compactCell(r.reason_text || "-", 60),
              }))
            : [{
                asset: "-",
                name: "No reconciliation data for period",
                mode: "-",
                actual: "-",
                expected: "-",
                variance: "-",
                unexplained: "-",
                pct: "-",
                reasons: "-",
              }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Fuel Reconciliation Report",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
        layout: "landscape",
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Fuel_Reconciliation_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/fuel-reconciliation.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15&fuel_price=1.5
  app.get("/fuel-reconciliation.xlsx", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const toleranceInput = Number(req.query?.tolerance ?? 0.15);
    const tolerance = Number.isFinite(toleranceInput) ? Math.max(0, toleranceInput) : 0.15;
    const fuelPriceInput = Number(req.query?.fuel_price ?? 0);
    const fuelPrice = Number.isFinite(fuelPriceInput) ? Math.max(0, fuelPriceInput) : 0;
    const modeFilter = String(req.query?.mode || "").trim().toLowerCase();
    const assetFilter = String(req.query?.asset_code || "").trim().toLowerCase();
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const fuelByAsset = db.prepare(fuelBenchmarkAssetsInRangeSql()).all(start, end);

    const rows = fuelByAsset
      .map((asset) => {
        const r = assetConsumption(db, asset, start, end, tolerance);
        const mode = r.metric_mode;
        const km = r.km_run;
        const hours = r.hours_run;
        // All fuel dispensed against what the recorded run should have used; fuel
        // that could not be matched to meter readings is named in the reasons.
        const fuel = Number(r.fuel_liters || 0);
        const expected = mode === "km"
          ? (km > 0 && Number(r.oem_km_per_l || 0) > 0 ? km / Number(r.oem_km_per_l) : 0)
          : (hours > 0 && Number(r.oem_lph || 0) > 0 ? hours * Number(r.oem_lph) : 0);
        const variance = fuel - expected;
        const allowed = expected * tolerance;
        const unexplained = Math.max(0, variance - allowed);
        const variancePct = expected > 0 ? (variance / expected) * 100 : null;
        const reasons = [];
        const runBase = mode === "km" ? km : hours;
        if (runBase <= 0 && fuel > 0) reasons.push("fuel with no run basis");
        if (Number(r.fill_count || 0) < 2) reasons.push("low sample count");
        if (variancePct != null && variancePct > 50) reasons.push("variance > 50%");
        const unmatchedL = Number((r.fuel_liters - Number(r.basis_liters || 0)).toFixed(2));
        if (runBase > 0 && unmatchedL > 0) reasons.push(`${unmatchedL} L not matched to meter readings`);
        if (r.suspect_readings) reasons.push(`${r.suspect_readings} meter reading(s) rejected`);
        if (!r.oem_set) reasons.push("benchmark not set");
        return {
          asset_code: String(r.asset_code || ""),
          asset_name: String(r.asset_name || ""),
          metric_mode: mode,
          fuel_liters: Number(fuel.toFixed(2)),
          expected_liters: Number(expected.toFixed(2)),
          variance_liters: Number(variance.toFixed(2)),
          unexplained_liters: Number(unexplained.toFixed(2)),
          variance_pct: variancePct == null ? null : Number(variancePct.toFixed(1)),
          reason_text: reasons.join("; "),
        };
      })
      .filter((r) => r.fuel_liters > 0)
      .filter((r) => (assetFilter ? String(r.asset_code || "").trim().toLowerCase() === assetFilter : true))
      .filter((r) => (modeFilter === "km" ? r.metric_mode === "km" : modeFilter === "hours" ? r.metric_mode === "hours" : true))
      .sort((a, b) => Number(b.unexplained_liters || 0) - Number(a.unexplained_liters || 0));

    const totals = rows.reduce((acc, r) => {
      acc.actual += Number(r.fuel_liters || 0);
      acc.expected += Number(r.expected_liters || 0);
      acc.variance += Number(r.variance_liters || 0);
      acc.unexplained += Number(r.unexplained_liters || 0);
      if (String(r.reason_text || "").trim()) acc.anomalies += 1;
      return acc;
    }, { actual: 0, expected: 0, variance: 0, unexplained: 0, anomalies: 0 });
    totals.value = totals.unexplained * fuelPrice;

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const wsSummary = wb.addWorksheet("Summary");
    wsSummary.columns = [
      { header: "Field", key: "field", width: 38 },
      { header: "Value", key: "value", width: 28 },
    ];
    wsSummary.getRow(1).font = { bold: true };
    wsSummary.addRows([
      { field: "Period", value: `${start} to ${end}` },
      { field: "Tolerance (%)", value: Number((tolerance * 100).toFixed(2)) },
      { field: "Fuel price per L", value: Number(fuelPrice.toFixed(2)) },
      { field: "Assets", value: rows.length },
      { field: "Anomaly assets", value: totals.anomalies },
      { field: "Actual liters", value: Number(totals.actual.toFixed(2)) },
      { field: "Expected liters", value: Number(totals.expected.toFixed(2)) },
      { field: "Variance liters", value: Number(totals.variance.toFixed(2)) },
      { field: "Estimated missing liters (unexplained)", value: Number(totals.unexplained.toFixed(2)) },
      { field: "Estimated missing value", value: Number(totals.value.toFixed(2)) },
    ]);

    const ws = wb.addWorksheet("Reconciliation");
    ws.columns = [
      { header: "Asset Code", key: "asset_code", width: 14 },
      { header: "Asset Name", key: "asset_name", width: 30 },
      { header: "Mode", key: "metric_mode", width: 10 },
      { header: "Actual Liters", key: "fuel_liters", width: 14 },
      { header: "Expected Liters", key: "expected_liters", width: 15 },
      { header: "Variance Liters", key: "variance_liters", width: 15 },
      { header: "Unexplained Liters", key: "unexplained_liters", width: 18 },
      { header: "Variance %", key: "variance_pct", width: 12 },
      { header: "Reasons", key: "reason_text", width: 46 },
    ];
    ws.getRow(1).font = { bold: true };
    rows.forEach((r) => ws.addRow(r));
    ws.autoFilter = { from: "A1", to: "I1" };
    ws.views = [{ state: "frozen", ySplit: 1 }];

    const buffer = await wb.xlsx.writeBuffer();
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="AML_Fuel_Reconciliation_${end}.xlsx"`)
      .send(Buffer.from(buffer));
  });

  // GET /api/reports/fuel-benchmark.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15
  app.get("/fuel-benchmark.xlsx", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const toleranceInput = Number(req.query?.tolerance ?? 0.15);
    const tolerance = Number.isFinite(toleranceInput) ? Math.max(0, toleranceInput) : 0.15;
    const modeFilter = String(req.query?.mode || "").trim().toLowerCase();
    const assetFilter = String(req.query?.asset_code || "").trim().toLowerCase();
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const fuelByAssetStmt = db.prepare(fuelBenchmarkAssetsInRangeSql());

    function buildRowsForPeriod(periodStart, periodEnd) {
      const fuelByAsset = fuelByAssetStmt.all(periodStart, periodEnd);
      const benchmarkRows = fuelByAsset.map((r) => {
        const row = assetConsumption(db, r, periodStart, periodEnd, tolerance);
        return { ...row, note: consumptionNote(row) };
      });
      const rows = sortBenchmarkRows(benchmarkRows
        .filter((r) => r.fuel_liters > 0)
        .filter((r) => (assetFilter ? String(r.asset_code || "").trim().toLowerCase() === assetFilter : true))
        .filter((r) => (modeFilter === "km" ? r.metric_mode === "km" : modeFilter === "hours" ? r.metric_mode === "hours" : true)));
      return { rows, benchmarkRows };
    }

    const periodMonths = monthsInRange(start, end);
    const { rows, benchmarkRows } = buildRowsForPeriod(start, end);
    const includedCodes = new Set(rows.map((r) => String(r.asset_code || "").trim().toLowerCase()).filter(Boolean));
    const missingRows = benchmarkRows
      .filter((r) => (assetFilter ? String(r.asset_code || "").trim().toLowerCase() === assetFilter : true))
      .filter((r) => (modeFilter === "km" ? r.metric_mode === "km" : modeFilter === "hours" ? r.metric_mode === "hours" : true))
      .filter((r) => !includedCodes.has(String(r.asset_code || "").trim().toLowerCase()))
      .map((r) => {
        let reason = "Excluded by report filters";
        if (Number(r.fuel_liters || 0) <= 0) reason = "No fuel logs in selected period";
        return {
          asset_code: r.asset_code,
          asset_name: r.asset_name,
          category: r.category || "",
          metric_mode: r.metric_mode,
          fuel_liters: Number(r.fuel_liters || 0),
          fill_count: Number(r.fill_count || 0),
          reason,
        };
      })
      .sort((a, b) => String(a.asset_code || "").localeCompare(String(b.asset_code || "")));

    const categoryRows = assetFilter ? [] : aggregateFuelBenchmarkByCategory(
      rows.map((r) => ({ ...r, is_excessive: r.flag === "EXCESSIVE" })),
      tolerance
    );

    const summaryBase = summarizeFuelBenchmarkRows(rows);
    const summary = { ...summaryBase, categories: categoryRows.length };

    const reportYear = Number(String(end).slice(0, 4));
    const ytdStart = `${reportYear}-01-01`;
    const quarters = calendarQuartersThrough(reportYear, end);
    const quarterSnapshots = quarters.map((q) => {
      const built = buildRowsForPeriod(q.start, q.end);
      const qSummary = summarizeFuelBenchmarkRows(built.rows);
      const costs = queryPeriodFleetCostTotals(q.start, q.end);
      return {
        ...q,
        rows: built.rows,
        summary: qSummary,
        costs,
      };
    });
    const ytdBuilt = buildRowsForPeriod(ytdStart, end);
    const ytdSummary = summarizeFuelBenchmarkRows(ytdBuilt.rows);
    const ytdCosts = queryPeriodFleetCostTotals(ytdStart, end);

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const wsSummary = wb.addWorksheet("Summary");
    wsSummary.columns = [
      { header: "Field", key: "field", width: 34 },
      { header: "Value", key: "value", width: 28 },
    ];
    wsSummary.addRows([
      { field: "Selected period", value: `${start} to ${end}` },
      { field: "Monthly tabs", value: periodMonths.map((m) => `${m.label} (${m.start} to ${m.end})`).join(", ") || start },
      { field: "Year context", value: String(reportYear) },
      { field: "YTD", value: `${ytdStart} to ${end}` },
      { field: "Quarters included", value: quarters.map((q) => q.key).join(", ") || "-" },
      { field: "Tolerance (%)", value: Number((tolerance * 100).toFixed(2)) },
      { field: "Mode filter", value: modeFilter || "all" },
      { field: "Asset filter", value: assetFilter || "all" },
      { field: "Equipment categories", value: summary.categories },
      { field: "Assets (Totals sheet)", value: summary.assets },
      { field: "Missing / Not Shown", value: missingRows.length },
      { field: "Excessive", value: summary.excessive },
      { field: "Fuel Total (L)", value: summary.fuel_liters },
      { field: "Hours Run (machines only)", value: summary.hours_run },
      { field: "Distance Run (km)", value: summary.km_run },
      { field: "Avg L/hr (machines only)", value: summary.avg_lph == null ? "" : summary.avg_lph },
      { field: "Avg km/L (LDV only)", value: summary.avg_km_per_l == null ? "" : summary.avg_km_per_l },
      { field: "YTD Fuel Total (L)", value: ytdSummary.fuel_liters },
      { field: "YTD Avg L/hr", value: ytdSummary.avg_lph == null ? "" : ytdSummary.avg_lph },
      { field: "YTD Avg km/L", value: ytdSummary.avg_km_per_l == null ? "" : ytdSummary.avg_km_per_l },
    ]);

    for (const month of periodMonths) {
      const { rows: monthRows } = buildRowsForPeriod(month.start, month.end);
      addFuelBenchmarkByAssetWorksheet(
        wb,
        month.label,
        monthRows,
        `No fuel benchmark data for ${month.label} (${month.start} to ${month.end})`
      );
    }

    for (const q of quarterSnapshots) {
      addFuelBenchmarkByAssetWorksheet(
        wb,
        q.label,
        q.rows,
        `No fuel benchmark data for ${q.label} (${q.start} to ${q.end})`
      );
    }
    addFuelBenchmarkByAssetWorksheet(
      wb,
      `YTD ${reportYear}`,
      ytdBuilt.rows,
      `No fuel benchmark data for YTD ${reportYear} (${ytdStart} to ${end})`
    );

    addFuelBenchmarkByAssetWorksheet(wb, "Totals", rows);

    const wsQuarterCompare = wb.addWorksheet("Quarter Compare");
    wsQuarterCompare.columns = [
      { header: "Period", key: "period", width: 14 },
      { header: "Start", key: "start", width: 12 },
      { header: "End", key: "end", width: 12 },
      { header: "Fuel (L)", key: "fuel_liters", width: 12 },
      { header: "Hours Run", key: "hours_run", width: 12 },
      { header: "Km Run", key: "km_run", width: 12 },
      { header: "Avg L/hr", key: "avg_lph", width: 12 },
      { header: "Avg km/L", key: "avg_km_per_l", width: 12 },
      { header: "Fuel Cost", key: "fuel_cost", width: 12 },
      { header: "Lube Cost", key: "lube_cost", width: 12 },
      { header: "Mechanics / Labor", key: "labor_cost", width: 16 },
      { header: "Parts Cost", key: "parts_cost", width: 12 },
      { header: "Downtime Cost", key: "downtime_cost", width: 14 },
      { header: "Total Cost", key: "total_cost", width: 12 },
      { header: "Fuel Δ vs prior Q", key: "fuel_delta", width: 16 },
      { header: "Cost Δ vs prior Q", key: "cost_delta", width: 16 },
      { header: "L/hr Δ vs prior Q", key: "lph_delta", width: 16 },
    ];
    let prevQ = null;
    for (const q of quarterSnapshots) {
      const fuelDelta = prevQ ? Number((q.summary.fuel_liters - prevQ.summary.fuel_liters).toFixed(2)) : null;
      const costDelta = prevQ ? Number((q.costs.total_cost - prevQ.costs.total_cost).toFixed(2)) : null;
      const lphDelta = prevQ && q.summary.avg_lph != null && prevQ.summary.avg_lph != null
        ? Number((q.summary.avg_lph - prevQ.summary.avg_lph).toFixed(3))
        : null;
      wsQuarterCompare.addRow({
        period: q.key,
        start: q.start,
        end: q.end,
        fuel_liters: q.summary.fuel_liters,
        hours_run: q.summary.hours_run,
        km_run: q.summary.km_run,
        avg_lph: q.summary.avg_lph == null ? "" : q.summary.avg_lph,
        avg_km_per_l: q.summary.avg_km_per_l == null ? "" : q.summary.avg_km_per_l,
        fuel_cost: Number(q.costs.fuel_cost.toFixed(2)),
        lube_cost: Number(q.costs.lube_cost.toFixed(2)),
        labor_cost: Number(q.costs.labor_cost.toFixed(2)),
        parts_cost: Number(q.costs.parts_cost.toFixed(2)),
        downtime_cost: Number(q.costs.downtime_cost.toFixed(2)),
        total_cost: q.costs.total_cost,
        fuel_delta: fuelDelta == null ? "" : fuelDelta,
        cost_delta: costDelta == null ? "" : costDelta,
        lph_delta: lphDelta == null ? "" : lphDelta,
      });
      prevQ = q;
    }
    wsQuarterCompare.addRow({
      period: `YTD ${reportYear}`,
      start: ytdStart,
      end,
      fuel_liters: ytdSummary.fuel_liters,
      hours_run: ytdSummary.hours_run,
      km_run: ytdSummary.km_run,
      avg_lph: ytdSummary.avg_lph == null ? "" : ytdSummary.avg_lph,
      avg_km_per_l: ytdSummary.avg_km_per_l == null ? "" : ytdSummary.avg_km_per_l,
      fuel_cost: Number(ytdCosts.fuel_cost.toFixed(2)),
      lube_cost: Number(ytdCosts.lube_cost.toFixed(2)),
      labor_cost: Number(ytdCosts.labor_cost.toFixed(2)),
      parts_cost: Number(ytdCosts.parts_cost.toFixed(2)),
      downtime_cost: Number(ytdCosts.downtime_cost.toFixed(2)),
      total_cost: ytdCosts.total_cost,
      fuel_delta: "",
      cost_delta: "",
      lph_delta: "",
    });

    // Monthly budget vs actual (yellow budget cells are editable what-if values in Excel)
    const yearMonths = monthsInRange(ytdStart, end);
    const wsBudget = wb.addWorksheet("Budget vs Actual");
    wsBudget.columns = [
      { header: "Month", key: "month", width: 10 },
      { header: "Budget Fuel", key: "budget_fuel", width: 12 },
      { header: "Budget Lube", key: "budget_lube", width: 12 },
      { header: "Budget Mechanics", key: "budget_labor", width: 16 },
      { header: "Budget Operating", key: "budget_operating", width: 16 },
      { header: "Actual Fuel", key: "actual_fuel", width: 12 },
      { header: "Actual Lube", key: "actual_lube", width: 12 },
      { header: "Actual Mechanics", key: "actual_labor", width: 16 },
      { header: "Actual Parts", key: "actual_parts", width: 12 },
      { header: "Actual Total", key: "actual_total", width: 12 },
      { header: "Fuel Variance", key: "var_fuel", width: 12 },
      { header: "Lube Variance", key: "var_lube", width: 12 },
      { header: "Mechanics Variance", key: "var_labor", width: 16 },
      { header: "Operating Variance", key: "var_operating", width: 16 },
    ];
    wsBudget.getCell("A1").note = "Yellow budget columns pre-fill from Finance budgets when available. Edit yellow cells in Excel for what-if; variances recalculate with formulas.";

    const yellowFill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFCE8A8" },
    };
    yearMonths.forEach((m, idx) => {
      const costs = queryPeriodFleetCostTotals(m.start, m.end);
      const mech = buildMechanicLaborDetail(db, m.label);
      const actualLabor = Number(mech.total_cost || 0) > 0
        ? Number(mech.total_cost || 0)
        : Number(costs.labor_cost || 0);
      const budgetFuel = Number(getMonthlyBudgetRow(db, m.label, "fuel")?.budget_amount || 0);
      const budgetLube = Number(getMonthlyBudgetRow(db, m.label, "lube")?.budget_amount || 0);
      const budgetLabor = Number(getMonthlyBudgetRow(db, m.label, "labor")?.budget_amount || 0);
      const budgetOperating = Number(getOperatingBudgetAmount(db, m.label)?.budget_amount || 0);
      const rowNum = idx + 2;
      const row = wsBudget.addRow({
        month: m.label,
        budget_fuel: budgetFuel,
        budget_lube: budgetLube,
        budget_labor: budgetLabor,
        budget_operating: budgetOperating,
        actual_fuel: Number(costs.fuel_cost.toFixed(2)),
        actual_lube: Number(costs.lube_cost.toFixed(2)),
        actual_labor: Number(actualLabor.toFixed(2)),
        actual_parts: Number(costs.parts_cost.toFixed(2)),
        actual_total: Number((costs.fuel_cost + costs.lube_cost + actualLabor + costs.parts_cost + costs.downtime_cost).toFixed(2)),
        var_fuel: null,
        var_lube: null,
        var_labor: null,
        var_operating: null,
      });
      // Excel formulas: Actual - Budget
      row.getCell("var_fuel").value = { formula: `F${rowNum}-B${rowNum}` };
      row.getCell("var_lube").value = { formula: `G${rowNum}-C${rowNum}` };
      row.getCell("var_labor").value = { formula: `H${rowNum}-D${rowNum}` };
      row.getCell("var_operating").value = { formula: `J${rowNum}-E${rowNum}` };
      for (const col of ["B", "C", "D", "E"]) {
        wsBudget.getCell(`${col}${rowNum}`).fill = yellowFill;
      }
    });

    // Quarter rollup of budget vs actual
    const wsBudgetQ = wb.addWorksheet("Budget by Quarter");
    wsBudgetQ.columns = [
      { header: "Period", key: "period", width: 14 },
      { header: "Budget Fuel", key: "budget_fuel", width: 12 },
      { header: "Budget Lube", key: "budget_lube", width: 12 },
      { header: "Budget Mechanics", key: "budget_labor", width: 16 },
      { header: "Budget Operating", key: "budget_operating", width: 16 },
      { header: "Actual Fuel", key: "actual_fuel", width: 12 },
      { header: "Actual Lube", key: "actual_lube", width: 12 },
      { header: "Actual Mechanics", key: "actual_labor", width: 16 },
      { header: "Actual Total", key: "actual_total", width: 12 },
      { header: "Fuel Variance", key: "var_fuel", width: 12 },
      { header: "Lube Variance", key: "var_lube", width: 12 },
      { header: "Mechanics Variance", key: "var_labor", width: 16 },
      { header: "Operating Variance", key: "var_operating", width: 16 },
      { header: "Cost Δ vs prior Q", key: "cost_delta", width: 16 },
    ];
    let prevBudgetQ = null;
    for (const q of quarterSnapshots) {
      const qMonths = monthsInRange(q.start, q.end);
      let budgetFuel = 0;
      let budgetLube = 0;
      let budgetLabor = 0;
      let budgetOperating = 0;
      let actualFuel = 0;
      let actualLube = 0;
      let actualLabor = 0;
      let actualTotal = 0;
      for (const m of qMonths) {
        const costs = queryPeriodFleetCostTotals(m.start, m.end);
        const mech = buildMechanicLaborDetail(db, m.label);
        const labor = Number(mech.total_cost || 0) > 0 ? Number(mech.total_cost || 0) : Number(costs.labor_cost || 0);
        budgetFuel += Number(getMonthlyBudgetRow(db, m.label, "fuel")?.budget_amount || 0);
        budgetLube += Number(getMonthlyBudgetRow(db, m.label, "lube")?.budget_amount || 0);
        budgetLabor += Number(getMonthlyBudgetRow(db, m.label, "labor")?.budget_amount || 0);
        budgetOperating += Number(getOperatingBudgetAmount(db, m.label)?.budget_amount || 0);
        actualFuel += Number(costs.fuel_cost || 0);
        actualLube += Number(costs.lube_cost || 0);
        actualLabor += labor;
        actualTotal += Number(costs.fuel_cost || 0) + Number(costs.lube_cost || 0) + labor + Number(costs.parts_cost || 0) + Number(costs.downtime_cost || 0);
      }
      const row = {
        period: q.key,
        budget_fuel: Number(budgetFuel.toFixed(2)),
        budget_lube: Number(budgetLube.toFixed(2)),
        budget_labor: Number(budgetLabor.toFixed(2)),
        budget_operating: Number(budgetOperating.toFixed(2)),
        actual_fuel: Number(actualFuel.toFixed(2)),
        actual_lube: Number(actualLube.toFixed(2)),
        actual_labor: Number(actualLabor.toFixed(2)),
        actual_total: Number(actualTotal.toFixed(2)),
        var_fuel: Number((actualFuel - budgetFuel).toFixed(2)),
        var_lube: Number((actualLube - budgetLube).toFixed(2)),
        var_labor: Number((actualLabor - budgetLabor).toFixed(2)),
        var_operating: Number((actualTotal - budgetOperating).toFixed(2)),
        cost_delta: prevBudgetQ == null ? "" : Number((actualTotal - prevBudgetQ).toFixed(2)),
      };
      wsBudgetQ.addRow(row);
      prevBudgetQ = actualTotal;
    }

    const wsCategory = wb.addWorksheet("Fuel by Category");
    wsCategory.columns = [
      { header: "Category", key: "category", width: 22 },
      { header: "Mode", key: "metric_mode", width: 10 },
      { header: "Units", key: "asset_count", width: 8 },
      { header: "Fuel Liters", key: "fuel_liters", width: 14 },
      { header: "Km Run", key: "km_run", width: 12 },
      { header: "Hours Run", key: "hours_run", width: 12 },
      { header: "Actual L/hr", key: "actual_lph", width: 12 },
      { header: "OEM L/hr", key: "oem_lph", width: 12 },
      { header: "Variance L/hr", key: "variance_lph", width: 14 },
      { header: "Actual km/L", key: "actual_km_per_l", width: 12 },
      { header: "OEM km/L", key: "oem_km_per_l", width: 12 },
      { header: "Variance km/L", key: "variance_km_per_l", width: 14 },
      { header: "Fill Count", key: "fill_count", width: 10 },
      { header: "Excessive Assets", key: "excessive_asset_count", width: 14 },
      { header: "Flag", key: "flag", width: 12 },
    ];
    if (categoryRows.length) {
      wsCategory.addRows(categoryRows);
    } else {
      wsCategory.addRow({
        category: assetFilter ? "Single asset filter — see Totals sheet" : "No fuel benchmark data for period",
        metric_mode: "",
        asset_count: "",
        fuel_liters: "",
        km_run: "",
        hours_run: "",
        actual_lph: "",
        oem_lph: "",
        variance_lph: "",
        actual_km_per_l: "",
        oem_km_per_l: "",
        variance_km_per_l: "",
        fill_count: "",
        excessive_asset_count: "",
        flag: "",
      });
    }

    const wsMissing = wb.addWorksheet("Missing Equipment");
    wsMissing.columns = [
      { header: "Asset Code", key: "asset_code", width: 16 },
      { header: "Asset Name", key: "asset_name", width: 32 },
      { header: "Category", key: "category", width: 18 },
      { header: "Mode", key: "metric_mode", width: 10 },
      { header: "Fuel Liters", key: "fuel_liters", width: 12 },
      { header: "Fill Count", key: "fill_count", width: 10 },
      { header: "Reason", key: "reason", width: 42 },
    ];
    if (missingRows.length) {
      wsMissing.addRows(missingRows);
    } else {
      wsMissing.addRow({
        asset_code: "-",
        asset_name: "No missing equipment for current filters",
        category: "",
        metric_mode: "",
        fuel_liters: "",
        fill_count: "",
        reason: "",
      });
    }

    const buffer = await wb.xlsx.writeBuffer();
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="AML_Fuel_Benchmark_${end}.xlsx"`)
      .send(buffer);
  });

  // GET /api/reports/fuel-machine-history.pdf?asset_code=A300AM&start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15&download=1
  app.get("/fuel-machine-history.pdf", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    const assetCode = String(req.query?.asset_code || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const toleranceInput = Number(req.query?.tolerance ?? 0.15);
    const tolerance = Number.isFinite(toleranceInput) ? Math.max(0, toleranceInput) : 0.15;
    const download = String(req.query?.download || "").trim() === "1";

    if (!assetCode) return reply.code(400).send({ error: "asset_code is required" });
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const asset = db.prepare(`
      SELECT a.id AS asset_id, a.id, a.asset_code, a.asset_name, a.category, a.archived, a.hire_billing_mode,
        COALESCE(a.baseline_fuel_l_per_hour, 5.0) AS oem_lph,
        COALESCE(a.baseline_fuel_km_per_l, 2.0) AS oem_kmpl,
        a.fuel_benchmark_set_at,
        ${sqlFuelMetricModeExpr("a")} AS metric_mode
      FROM assets a
      WHERE a.asset_code = ?
    `).get(assetCode);
    if (!asset) return reply.code(404).send({ error: `asset not found: ${assetCode}` });

    // Same calculation as the benchmark, fill by fill.
    const calc = assetConsumption(db, asset, start, end, tolerance, { withTrace: true });
    const mode = calc.metric_mode;
    const oem = calc.oem_lph;
    const oemK = calc.oem_km_per_l;
    const rows = calc.trace.map((t) => {
      const run = Number(t.run || 0);
      const lph = mode === "hours" && run > 0 ? t.interval_liters / run : null;
      const kmpl = mode === "km" && run > 0 && t.interval_liters > 0 ? run / t.interval_liters : null;
      const over = mode === "km"
        ? kmpl != null && calc.threshold_km_per_l != null && kmpl < calc.threshold_km_per_l
        : lph != null && calc.threshold_lph != null && lph > calc.threshold_lph;
      return {
        log_date: t.date,
        metric_mode: mode,
        fuel_liters: Number(Number(t.liters || 0).toFixed(2)),
        reading: t.reading,
        run_between: run,
        hours_between: mode === "hours" ? run : 0,
        km_between: mode === "km" ? run : 0,
        interval_liters: t.interval_liters,
        actual_lph: lph == null ? null : Number(lph.toFixed(3)),
        actual_km_per_l: kmpl == null ? null : Number(kmpl.toFixed(3)),
        flag: t.status === "ok" ? (over ? "EXCESSIVE" : "OK") : "CHECK",
        status: t.status === "ok" ? "" : t.status,
        source: t.source || "",
      };
    });
    const summary = {
      fill_days: rows.length,
      excessive_days: rows.filter((r) => r.flag === "EXCESSIVE").length,
      fuel_liters: calc.fuel_liters,
      matched_liters: calc.matched_liters,
      hours_between: calc.hours_run,
      km_between: calc.km_run,
      avg_lph: calc.actual_lph,
      avg_km_per_l: calc.actual_km_per_l,
      run_source: calc.run_source,
      note: consumptionNote(calc),
      metric_mode: mode,
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Machine Fuel Fill History");
        kvGrid(doc, [
          { k: "Asset", v: `${asset.asset_code} ${asset.asset_name ? `- ${asset.asset_name}` : ""}` },
          { k: "Period", v: `${start} to ${end}` },
          { k: "Tolerance", v: `${fmtNum(tolerance * 100, 1)}% (${mode === "km" ? "below OEM" : "above OEM"})` },
          { k: mode === "km" ? "OEM km/L" : "OEM L/hr", v: (mode === "km" ? oemK : oem) == null ? "Not set" : fmtNum(mode === "km" ? oemK : oem, 3) },
          { k: "Fills", v: fmtNum(summary.fill_days, 0) },
          { k: "Excessive intervals", v: fmtNum(summary.excessive_days, 0) },
          { k: "Fuel Total (L)", v: fmtNum(summary.fuel_liters, 2) },
          { k: "Fuel matched to readings (L)", v: fmtNum(summary.matched_liters, 2) },
          { k: "Worked out from", v: summary.run_source === "fill_meter" ? "Meter readings on the fills" : summary.run_source === "daily_hours" ? "Daily hours (fill readings incomplete)" : "No usable readings" },
          ...(summary.note ? [{ k: "Notes", v: summary.note }] : []),
          { k: mode === "km" ? "Distance (km)" : "Hours", v: mode === "km" ? fmtNum(summary.km_between, 2) : fmtNum(summary.hours_between, 2) },
          { k: mode === "km" ? "Avg km/L" : "Avg L/hr", v: mode === "km" ? (summary.avg_km_per_l == null ? "-" : fmtNum(summary.avg_km_per_l, 3)) : (summary.avg_lph == null ? "-" : fmtNum(summary.avg_lph, 3)) },
        ], 2);

        sectionTitle(doc, "Fill Entries");
        table(
          doc,
          [
            { key: "log_date", label: "Date", width: 0.10 },
            { key: "fuel_liters", label: "Fuel (L)", width: 0.08, align: "right" },
            { key: "reading", label: mode === "km" ? "Odometer" : "Hour meter", width: 0.10, align: "right" },
            { key: "run_between", label: mode === "km" ? "km since last" : "Hours since last", width: 0.11, align: "right" },
            { key: "actual", label: mode === "km" ? "km/L" : "L/hr", width: 0.08, align: "right" },
            { key: "flag", label: "Status", width: 0.09, align: "center" },
            { key: "status", label: "What happened", width: 0.34 },
            { key: "source", label: "Source", width: 0.10 },
          ],
          rows.length
            ? rows.map((r) => ({
                log_date: r.log_date,
                fuel_liters: fmtNum(r.fuel_liters, 2),
                reading: r.reading == null ? "-" : fmtNum(r.reading, 1),
                run_between: r.run_between ? fmtNum(r.run_between, 1) : "-",
                actual: mode === "km"
                  ? (r.actual_km_per_l == null ? "-" : fmtNum(r.actual_km_per_l, 2))
                  : (r.actual_lph == null ? "-" : fmtNum(r.actual_lph, 2)),
                flag: r.flag,
                status: r.status,
                source: r.source || "",
              }))
            : [{
                log_date: "-",
                fuel_liters: "-",
                reading: "-",
                run_between: "-",
                actual: "-",
                flag: "-",
                status: "No fill data in selected period",
                source: "",
              }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Fuel Machine History",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
        layout: "landscape",
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Fuel_History_${asset.asset_code}_${end}.pdf"`
      )
      .send(pdf);
  });
}
