// IRONLOG/api/utils/monthlyBudgetSpendWorkbook.js
import ExcelJS from "exceljs";
import { createManagementSummary, styleManagementDetailSheet } from "./managementWorkbook.js";

const NAVY = "FF17365D";
const BLUE_LIGHT = "FFD9EAF7";
const GREEN = "FF166534";
const GREEN_LIGHT = "FFDCFCE7";
const RED = "FFB91C1C";
const RED_LIGHT = "FFFEE2E2";
const AMBER = "FFB45309";
const AMBER_LIGHT = "FFFFF2CC";
const GRID = "FFD9E2F3";
const MONEY = '"$"#,##0.00;[Red]("$"#,##0.00);-';

function applyFill(cell, argb) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function applyBorder(cell) {
  const border = { style: "thin", color: { argb: GRID } };
  cell.border = { top: border, right: border, bottom: border, left: border };
}

function moneyCard(value, fallback = "Not set") {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

export async function buildMonthlyBudgetSpendWorkbook(report) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "IRONLOG";
  workbook.created = new Date();
  workbook.properties.title = "IRONLOG Monthly Budget vs Spend";
  workbook.properties.subject = `Budget versus spend for ${report.month}`;

  const varianceTone = report.variance == null ? "attention" : report.variance > 0 ? "warning" : "normal";
  const { ws, nextRow } = createManagementSummary(workbook, {
    title: "IRONLOG Monthly Budget vs Spend",
    periodLabel: `${report.start} to ${report.end} | Site: ${String(report.site_code || "main").toUpperCase()} | Currency: USD`,
    cards: [
      { label: "MONTHLY BUDGET", value: moneyCard(report.budget?.is_set ? report.budget.amount : null), numFmt: report.budget?.is_set ? MONEY : undefined, tone: report.budget?.is_set ? "normal" : "attention" },
      { label: "ACTUAL SPEND", value: report.total_spend, numFmt: MONEY, tone: "normal" },
      { label: "VARIANCE TO BUDGET", value: moneyCard(report.variance, "Budget not set"), numFmt: report.variance != null ? MONEY : undefined, tone: varianceTone },
      { label: "BUDGET USED", value: report.budget_used_pct == null ? "Budget not set" : report.budget_used_pct / 100, numFmt: report.budget_used_pct == null ? undefined : "0.0%", tone: report.budget_used_pct != null && report.budget_used_pct > 100 ? "warning" : "normal" },
    ],
    scopeLines: [
      `Budget source: ${report.budget?.is_set ? `Operating budget saved for ${report.month}` : `No operating budget has been saved for ${report.month}`}.`,
      "Actual spend combines saved mechanic labour, lube usage, non-lubricant stores issues, and completed service work-order labour.",
      "Service work-order labour with a matching mechanic job-card number is excluded from the maintenance line to reduce double capture.",
    ],
  });

  const startRow = nextRow + 2;
  ws.mergeCells(startRow, 1, startRow, 6);
  const titleCell = ws.getCell(startRow, 1);
  titleCell.value = "Spend by category";
  titleCell.font = { name: "Arial", size: 12, bold: true, color: { argb: NAVY } };
  titleCell.alignment = { vertical: "middle" };
  ws.getRow(startRow).height = 22;

  const headerRow = startRow + 1;
  const headers = ["Spend category", "Actual spend (USD)", "Share of spend", "Records", "Data source", "Comment"];
  headers.forEach((header, index) => {
    const cell = ws.getCell(headerRow, index + 1);
    cell.value = header;
    applyFill(cell, NAVY);
    applyBorder(cell);
    cell.font = { name: "Arial", size: 10, bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  });
  ws.getRow(headerRow).height = 30;

  const categoryStart = headerRow + 1;
  report.categories.forEach((row, index) => {
    const excelRow = categoryStart + index;
    const values = [
      row.label,
      Number(row.amount || 0),
      report.total_spend > 0 ? Number(row.amount || 0) / report.total_spend : null,
      Number(row.records || 0),
      row.source,
      row.key === "maintenance" ? "Excludes service WOs with matching mechanic job cards" : "",
    ];
    values.forEach((value, column) => {
      const cell = ws.getCell(excelRow, column + 1);
      cell.value = value;
      applyBorder(cell);
      cell.font = { name: "Arial", size: 10, color: { argb: "FF1F2937" } };
      cell.alignment = { vertical: "middle", wrapText: column >= 4 };
    });
    ws.getCell(excelRow, 2).numFmt = MONEY;
    ws.getCell(excelRow, 3).numFmt = "0.0%";
    ws.getCell(excelRow, 4).numFmt = "#,##0";
    ws.getRow(excelRow).height = 28;
  });

  const totalRow = categoryStart + report.categories.length;
  const totalValues = ["Total actual spend", report.total_spend, 1, report.categories.reduce((sum, row) => sum + Number(row.records || 0), 0), "", ""];
  totalValues.forEach((value, column) => {
    const cell = ws.getCell(totalRow, column + 1);
    cell.value = value;
    applyFill(cell, BLUE_LIGHT);
    applyBorder(cell);
    cell.font = { name: "Arial", size: 10, bold: true, color: { argb: NAVY } };
    cell.alignment = { vertical: "middle" };
  });
  ws.getCell(totalRow, 2).numFmt = MONEY;
  ws.getCell(totalRow, 3).numFmt = "0.0%";
  ws.getCell(totalRow, 4).numFmt = "#,##0";

  const varianceRow = totalRow + 2;
  const varianceLabel = report.budget?.is_set ? (report.variance > 0 ? "Over budget" : "Under budget") : "Budget status";
  ws.mergeCells(varianceRow, 1, varianceRow, 2);
  ws.mergeCells(varianceRow, 3, varianceRow, 4);
  ws.mergeCells(varianceRow, 5, varianceRow, 6);
  const label = ws.getCell(varianceRow, 1);
  label.value = varianceLabel;
  const value = ws.getCell(varianceRow, 3);
  value.value = report.budget?.is_set ? Math.abs(Number(report.variance || 0)) : "Set the operating budget under Assets → Plant Hire & Budget";
  const explanation = ws.getCell(varianceRow, 5);
  explanation.value = report.budget?.is_set
    ? `${report.variance > 0 ? "Actual spend exceeds" : "Actual spend remains within"} the saved monthly operating budget.`
    : "No budget comparison can be calculated until a monthly operating budget is saved.";
  const tone = !report.budget?.is_set ? AMBER_LIGHT : report.variance > 0 ? RED_LIGHT : GREEN_LIGHT;
  const toneText = !report.budget?.is_set ? AMBER : report.variance > 0 ? RED : GREEN;
  [label, value, explanation].forEach((cell) => {
    applyFill(cell, tone);
    applyBorder(cell);
    cell.font = { name: "Arial", size: 10, bold: cell !== explanation, color: { argb: toneText } };
    cell.alignment = { vertical: "middle", wrapText: cell === explanation };
  });
  if (report.budget?.is_set) value.numFmt = MONEY;
  ws.getRow(varianceRow).height = 28;

  ws.getColumn(1).width = 28;
  ws.getColumn(2).width = 20;
  ws.getColumn(3).width = 16;
  ws.getColumn(4).width = 12;
  ws.getColumn(5).width = 44;
  ws.getColumn(6).width = 38;
  ws.views = [{ showGridLines: false }];
  ws.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

  const detail = workbook.addWorksheet("Spend Detail", { views: [{ showGridLines: false }] });
  detail.columns = [
    { header: "Date", key: "date", width: 14 },
    { header: "Spend Category", key: "category", width: 27 },
    { header: "Asset", key: "asset_code", width: 14 },
    { header: "Reference", key: "reference", width: 18 },
    { header: "Description", key: "description", width: 48 },
    { header: "Qty / Hrs", key: "quantity", width: 13 },
    { header: "Unit Cost / Rate", key: "unit_rate", width: 17 },
    { header: "Actual Spend (USD)", key: "amount", width: 20 },
  ];
  detail.addRows(report.details.length ? report.details : [{
    date: "", category: "", asset_code: "", reference: "", description: "No spend activity was recorded for this month.", quantity: null, unit_rate: null, amount: 0,
  }]);
  styleManagementDetailSheet(detail, {
    title: "Monthly Spend Detail",
    subtitle: `${report.start} to ${report.end} | Source entries behind the Budget vs Spend summary`,
    frozenColumns: 3,
    numberFormats: { F: "#,##0.000", G: MONEY, H: MONEY },
  });
  detail.autoFilter = { from: { row: 4, column: 1 }, to: { row: detail.rowCount, column: 8 } };
  detail.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

  return workbook;
}
