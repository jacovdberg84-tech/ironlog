import ExcelJS from "exceljs";
import { createManagementSummary, styleManagementDetailSheet, MANAGEMENT_XLSX } from "./managementWorkbook.js";
import { normalizeStockCategory, stockCategoryLabel } from "./stockCategory.js";

const { COLORS } = MANAGEMENT_XLSX;
const thinBorder = { style: "thin", color: { argb: COLORS.grid } };

function number(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function dateOnly(value) {
  return String(value || "").slice(0, 10);
}

function stockStatus(row) {
  const closing = number(row.closing_qty);
  const min = number(row.min_qty);
  const max = number(row.max_qty);
  if (closing <= 0 && (min > 0 || row.critical)) return "URGENT - STOCK OUT";
  if (min > 0 && closing <= min) return "REORDER";
  if (max > 0 && closing > max) return "EXCESS";
  return "OK";
}

function actionFor(row) {
  const status = stockStatus(row);
  if (status === "URGENT - STOCK OUT") return "Order urgently";
  if (status === "REORDER") return "Raise replenishment";
  if (status === "EXCESS") return "Review holding";
  return "Monitor";
}

function asWorkbookRows(rows) {
  return (Array.isArray(rows) ? rows : []).map((row) => ({
    ...row,
    stock_category: normalizeStockCategory(row.stock_category) || (row.is_lube ? "oil" : "part"),
    category: stockCategoryLabel(normalizeStockCategory(row.stock_category) || (row.is_lube ? "oil" : "part")),
    supplier: String(row.supplier || "").trim(),
    location: String(row.location || "").trim() || "Unspecified",
    criticality: row.critical ? "Critical" : "Standard",
    status: stockStatus(row),
    action: actionFor(row),
    last_movement_date: dateOnly(row.last_movement_at),
    stock_value: number(row.closing_qty) * number(row.unit_cost),
  }));
}

function applyHeading(cell, value) {
  cell.value = value;
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.navy } };
  cell.font = { name: "Arial", size: 10, bold: true, color: { argb: COLORS.white } };
  cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  cell.border = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
}

function writeAttentionTable(ws, startRow, rows) {
  ws.mergeCells(startRow, 1, startRow, 8);
  const heading = ws.getCell(startRow, 1);
  heading.value = "Reorder and critical attention";
  heading.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.blueLight } };
  heading.font = { name: "Arial", size: 11, bold: true, color: { argb: COLORS.navy } };
  heading.border = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
  heading.alignment = { vertical: "middle" };
  ws.getRow(startRow).height = 22;

  const headers = ["Stock code", "Description", "Location", "Criticality", "Closing qty", "Minimum qty", "Stock value (USD)", "Action"];
  headers.forEach((value, index) => applyHeading(ws.getCell(startRow + 1, index + 1), value));
  ws.getRow(startRow + 1).height = 26;

  const displayRows = rows.length ? rows : [{
    part_code: "—",
    part_name: "No critical or below-minimum items for this period.",
    location: "",
    criticality: "",
    closing_qty: "",
    min_qty: "",
    stock_value: "",
    action: "Monitor",
  }];
  displayRows.forEach((row, index) => {
    const values = [
      row.part_code || "",
      row.part_name || "",
      row.location || "",
      row.criticality || "",
      row.closing_qty === "" ? "" : number(row.closing_qty),
      row.min_qty === "" ? "" : number(row.min_qty),
      row.stock_value === "" ? "" : number(row.stock_value),
      row.action || "Monitor",
    ];
    const excelRow = ws.getRow(startRow + 2 + index);
    values.forEach((value, columnIndex) => {
      const cell = excelRow.getCell(columnIndex + 1);
      cell.value = value;
      cell.font = { name: "Arial", size: 10, color: { argb: COLORS.body } };
      cell.alignment = { vertical: "middle", wrapText: columnIndex === 1 };
      cell.border = { bottom: thinBorder };
      if (index % 2) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
    });
    excelRow.height = 20;
  });
  ws.getColumn(5).numFmt = "#,##0.00";
  ws.getColumn(6).numFmt = "#,##0.00";
  ws.getColumn(7).numFmt = '"$"#,##0.00';
}

const CATEGORY_SHEETS = [
  { key: "part", name: "Parts", title: "Parts register", scope: "Workshop spares and service kits" },
  { key: "component", name: "Components", title: "Components register", scope: "Major assemblies: engines, transmissions, pumps, cylinders, hubs, radiators" },
  { key: "get", name: "G.E.T", title: "G.E.T register", scope: "Ground engaging tools: tips, adapters, cutting edges, blades" },
  { key: "tyre", name: "Tyres", title: "Tyres register", scope: "Tyres" },
  { key: "oil", name: "Oils & Lubricants", title: "Oils and lubricants register", scope: "Oils, greases and fluids" },
];

/** Summary table: items, value and reorder count per stock category. Returns the last row used. */
function writeCategoryTable(ws, startRow, items) {
  ws.mergeCells(startRow, 1, startRow, 8);
  const heading = ws.getCell(startRow, 1);
  heading.value = "Stock by category";
  heading.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.blueLight } };
  heading.font = { name: "Arial", size: 11, bold: true, color: { argb: COLORS.navy } };
  heading.border = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
  ws.getRow(startRow).height = 22;
  ["Category", "Items", "Stock value (USD)", "Below minimum", "Issued in period (qty)"].forEach((value, index) => applyHeading(ws.getCell(startRow + 1, index + 1), value));
  ws.getRow(startRow + 1).height = 26;
  const rows = CATEGORY_SHEETS.map((sheet) => {
    const list = items.filter((row) => row.stock_category === sheet.key);
    return [
      sheet.name,
      list.length,
      list.reduce((sum, row) => sum + number(row.stock_value), 0),
      list.filter((row) => row.status === "REORDER" || row.status === "URGENT - STOCK OUT").length,
      list.reduce((sum, row) => sum + number(row.issues_qty), 0),
    ];
  });
  rows.push(["Total", ...[1, 2, 3, 4].map((i) => rows.reduce((sum, r) => sum + number(r[i]), 0))]);
  rows.forEach((values, index) => {
    const excelRow = ws.getRow(startRow + 2 + index);
    const total = index === rows.length - 1;
    values.forEach((value, columnIndex) => {
      const cell = excelRow.getCell(columnIndex + 1);
      cell.value = value;
      cell.font = { name: "Arial", size: 10, bold: total, color: { argb: COLORS.body } };
      cell.border = total ? { top: thinBorder, bottom: thinBorder } : { bottom: thinBorder };
      if (columnIndex === 2) cell.numFmt = '"$"#,##0.00';
      else if (columnIndex === 4) cell.numFmt = "#,##0.00";
      else if (columnIndex > 0) cell.numFmt = "#,##0";
    });
    excelRow.height = 20;
  });
  return startRow + 1 + rows.length;
}

function addInventorySheet(workbook, name, title, subtitle, rows, includeCriticality = true) {
  const ws = workbook.addWorksheet(name);
  ws.columns = [
    { header: "Stock Code", key: "part_code", width: 16 },
    { header: "Description", key: "part_name", width: 34 },
    { header: "Category", key: "category", width: 18 },
    { header: "Supplier", key: "supplier", width: 18 },
    { header: "Location", key: "location", width: 18 },
    ...(includeCriticality ? [{ header: "Criticality", key: "criticality", width: 13 }] : []),
    { header: "Min Qty", key: "min_qty", width: 12 },
    { header: "Max Qty", key: "max_qty", width: 12 },
    { header: "Reorder Point", key: "reorder_point", width: 14 },
    { header: "Opening Qty", key: "opening_qty", width: 13 },
    { header: "Receipts", key: "receipts_qty", width: 12 },
    { header: "Issues", key: "issues_qty", width: 12 },
    { header: "Returns", key: "returns_qty", width: 12 },
    { header: "Transfers In", key: "transfers_in_qty", width: 13 },
    { header: "Transfers Out", key: "transfers_out_qty", width: 14 },
    { header: "Adjustments", key: "adjustments_qty", width: 14 },
    { header: "Closing Qty", key: "closing_qty", width: 13 },
    { header: "Unit Cost (USD)", key: "unit_cost", width: 15 },
    { header: "Stock Value (USD)", key: "stock_value", width: 17 },
    { header: "Last Movement", key: "last_movement_date", width: 15 },
    { header: "Status", key: "status", width: 22 },
  ];
  ws.addRows(rows);
  styleManagementDetailSheet(ws, {
    title,
    subtitle,
    frozenColumns: 2,
    numberFormats: {
      min_qty: "#,##0.00",
      max_qty: "#,##0.00",
      reorder_point: "#,##0.00",
      opening_qty: "#,##0.00",
      receipts_qty: "#,##0.00",
      issues_qty: "#,##0.00",
      returns_qty: "#,##0.00",
      transfers_in_qty: "#,##0.00",
      transfers_out_qty: "#,##0.00",
      adjustments_qty: "#,##0.00",
      closing_qty: "#,##0.00",
      unit_cost: '"$"#,##0.00',
      stock_value: '"$"#,##0.00',
    },
  });
  return ws;
}

function addMovementsSheet(workbook, rows, periodLabel) {
  const ws = workbook.addWorksheet("Stock Movements");
  ws.columns = [
    { header: "Date", key: "movement_date", width: 14 },
    { header: "Transaction Type", key: "transaction_type", width: 18 },
    { header: "Stock Code", key: "part_code", width: 16 },
    { header: "Description", key: "part_name", width: 34 },
    { header: "Qty", key: "quantity", width: 12 },
    { header: "Location", key: "location", width: 18 },
    { header: "Bin", key: "bin_code", width: 14 },
    { header: "Unit Cost (USD)", key: "unit_cost", width: 15 },
    { header: "Total Value (USD)", key: "total_value", width: 17 },
    { header: "Reference", key: "reference", width: 34 },
  ];
  ws.addRows(rows.map((row) => ({
    movement_date: dateOnly(row.movement_at),
    transaction_type: row.transaction_type || row.movement_type || "Movement",
    part_code: row.part_code || "",
    part_name: row.part_name || "",
    quantity: Math.abs(number(row.quantity)),
    location: row.location_code || "Unspecified",
    bin_code: row.bin_code || "",
    unit_cost: number(row.unit_cost),
    total_value: Math.abs(number(row.quantity)) * number(row.unit_cost),
    reference: row.reference || "",
  })));
  styleManagementDetailSheet(ws, {
    title: "Stock movements",
    subtitle: `Transactions recorded in IRONLOG · ${periodLabel}`,
    frozenColumns: 2,
    numberFormats: {
      quantity: "#,##0.00",
      unit_cost: '"$"#,##0.00',
      total_value: '"$"#,##0.00',
    },
  });
  return ws;
}

/** Builds the GM-friendly weekly or monthly stock workbook from live store data. */
export async function buildWorkshopInventoryReportWorkbook(data = {}, options = {}) {
  const reportType = String(options.reportType || "monthly").toLowerCase() === "weekly" ? "Weekly" : "Monthly";
  const startDate = String(options.startDate || "");
  const endDate = String(options.endDate || "");
  const periodLabel = `${reportType} report: ${startDate} to ${endDate}`;
  const items = asWorkbookRows(data.items);
  const movements = Array.isArray(data.movements) ? data.movements : [];
  const attentionRows = items
    .filter((row) => row.critical || row.status !== "OK")
    .sort((left, right) => (
      Number(Boolean(right.critical)) - Number(Boolean(left.critical))
      || number(right.min_qty) - number(left.min_qty)
      || number(left.closing_qty) - number(right.closing_qty)
    ));
  const belowMin = items.filter((row) => row.status === "REORDER" || row.status === "URGENT - STOCK OUT");
  const criticalBelow = attentionRows.filter((row) => row.critical && row.status !== "OK");
  const totalValue = items.reduce((sum, row) => sum + number(row.stock_value), 0);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "IRONLOG";
  workbook.created = new Date();
  const { ws: summary, nextRow } = createManagementSummary(workbook, {
    title: "IRONLOG Workshop Inventory Report",
    periodLabel,
    cards: [
      { label: "Store items", value: items.length, numFmt: "#,##0" },
      { label: "Inventory value (USD)", value: totalValue, numFmt: '"$"#,##0.00' },
      { label: "Below minimum", value: belowMin.length, numFmt: "#,##0", tone: "warning" },
      { label: "Critical below min", value: criticalBelow.length, numFmt: "#,##0", tone: "attention" },
      { label: "Period movements", value: movements.length, numFmt: "#,##0" },
    ],
    scopeLines: [
      "Closing balances are calculated from stock movements recorded through the report end date.",
      "Values use the current IRONLOG catalogue unit cost. Fields not maintained in the stock master are left blank, not estimated.",
    ],
  });
  summary.name = "GM Summary";
  summary.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  summary.getColumn(1).width = 16;
  summary.getColumn(2).width = 34;
  summary.getColumn(3).width = 18;
  summary.getColumn(4).width = 13;
  summary.getColumn(5).width = 14;
  summary.getColumn(6).width = 14;
  summary.getColumn(7).width = 18;
  summary.getColumn(8).width = 22;
  const afterCategories = writeCategoryTable(summary, nextRow + 2, items);
  writeAttentionTable(summary, afterCategories + 2, attentionRows.slice(0, 15));
  const noteRow = afterCategories + 2 + Math.max(4, attentionRows.slice(0, 15).length + 3);
  summary.mergeCells(noteRow, 1, noteRow, 8);
  const note = summary.getCell(noteRow, 1);
  note.value = "Source: IRONLOG Stores. Each item's category (oils, G.E.T, parts, components, tyres) is set in Stock Control; update supplier, location and min/max there too.";
  note.font = { name: "Arial", size: 9, italic: true, color: { argb: COLORS.muted } };

  // One register per stock category, in the order stores reports them.
  for (const sheet of CATEGORY_SHEETS) {
    addInventorySheet(
      workbook,
      sheet.name,
      sheet.title,
      `${periodLabel} · ${sheet.scope}`,
      items.filter((row) => row.stock_category === sheet.key),
      sheet.key !== "oil",
    );
  }
  addMovementsSheet(workbook, movements, periodLabel);

  const critical = workbook.addWorksheet("Critical Spares");
  critical.columns = [
    { header: "Stock Code", key: "part_code", width: 16 },
    { header: "Description", key: "part_name", width: 34 },
    { header: "Category", key: "category", width: 18 },
    { header: "Location", key: "location", width: 18 },
    { header: "Criticality", key: "criticality", width: 13 },
    { header: "Closing Qty", key: "closing_qty", width: 13 },
    { header: "Min Qty", key: "min_qty", width: 12 },
    { header: "Max Qty", key: "max_qty", width: 12 },
    { header: "Stock Value (USD)", key: "stock_value", width: 17 },
    { header: "Last Movement", key: "last_movement_date", width: 15 },
    { header: "Action Required", key: "action", width: 22 },
  ];
  critical.addRows(attentionRows);
  styleManagementDetailSheet(critical, {
    title: "Critical and below-minimum spares",
    subtitle: `${periodLabel} · Critical items and all items needing replenishment`,
    frozenColumns: 2,
    numberFormats: {
      closing_qty: "#,##0.00",
      min_qty: "#,##0.00",
      max_qty: "#,##0.00",
      stock_value: '"$"#,##0.00',
    },
  });

  return workbook.xlsx.writeBuffer();
}

export function resolveWorkshopInventoryReportPeriod(reportType, reportDate) {
  const type = String(reportType || "monthly").trim().toLowerCase() === "weekly" ? "weekly" : "monthly";
  const raw = String(reportDate || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) throw new Error("report_date must be YYYY-MM-DD");
  const anchor = new Date(`${raw}T00:00:00Z`);
  if (Number.isNaN(anchor.valueOf()) || anchor.toISOString().slice(0, 10) !== raw) {
    throw new Error("report_date must be a valid calendar date");
  }
  if (type === "weekly") {
    const start = new Date(anchor);
    start.setUTCDate(start.getUTCDate() - 6);
    return { report_type: type, start_date: start.toISOString().slice(0, 10), end_date: raw };
  }
  const start = new Date(Date.UTC(anchor.getUTCFullYear(), anchor.getUTCMonth(), 1));
  const end = new Date(Date.UTC(anchor.getUTCFullYear(), anchor.getUTCMonth() + 1, 0));
  return {
    report_type: type,
    start_date: start.toISOString().slice(0, 10),
    end_date: end.toISOString().slice(0, 10),
  };
}
