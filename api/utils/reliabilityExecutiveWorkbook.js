import ExcelJS from "exceljs";
import { createManagementSummary, styleManagementDetailSheet, MANAGEMENT_XLSX } from "./managementWorkbook.js";
import { addNativeBarCharts } from "./nativeExcelCharts.js";

const { COLORS } = MANAGEMENT_XLSX;
const thinBorder = { style: "thin", color: { argb: COLORS.grid } };

function number(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function hoursOrBlank(value) {
  return value == null || !Number.isFinite(Number(value)) ? "" : Number(value);
}

function applyFill(cell, argb) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function writeExecutiveTable(ws, startRow, rows) {
  const headingRow = startRow;
  ws.mergeCells(headingRow, 1, headingRow, 9);
  const heading = ws.getCell(headingRow, 1);
  heading.value = "Equipment requiring attention";
  applyFill(heading, COLORS.blueLight);
  heading.font = { name: "Arial", size: 11, bold: true, color: { argb: COLORS.navy } };
  heading.border = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
  heading.alignment = { vertical: "middle" };
  ws.getRow(headingRow).height = 22;

  const headerRow = headingRow + 1;
  const headers = ["Asset", "Equipment", "Category", "Failures", "Downtime h", "MTBF h", "LTTR h", "Recorded h", "Action"];
  headers.forEach((value, index) => {
    const cell = ws.getCell(headerRow, index + 1);
    cell.value = value;
    applyFill(cell, COLORS.navy);
    cell.font = { name: "Arial", size: 10, bold: true, color: { argb: COLORS.white } };
    cell.border = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  ws.getRow(headerRow).height = 26;

  const displayRows = rows.length ? rows : [{
    asset_code: "—",
    asset_name: "No breakdown downtime recorded for this scope.",
    category: "",
    failure_count: "",
    downtime_hours: "",
    mtbf_hours: "",
    lttr_hours: "",
    recorded_downtime_hours: "",
    action: "Monitor",
  }];
  displayRows.forEach((row, rowIndex) => {
    const values = [
      row.asset_code || "",
      row.asset_name || "",
      row.category || "",
      row.failure_count === "" ? "" : number(row.failure_count),
      row.downtime_hours === "" ? "" : number(row.downtime_hours),
      hoursOrBlank(row.mtbf_hours),
      hoursOrBlank(row.lttr_hours),
      row.recorded_downtime_hours === "" ? "" : number(row.recorded_downtime_hours),
      row.action || "Review",
    ];
    const excelRow = ws.getRow(headerRow + 1 + rowIndex);
    values.forEach((value, columnIndex) => {
      const cell = excelRow.getCell(columnIndex + 1);
      cell.value = value;
      cell.font = { name: "Arial", size: 10, color: { argb: COLORS.body } };
      cell.alignment = { vertical: "middle", wrapText: columnIndex === 1 };
      cell.border = { bottom: thinBorder };
      if (rowIndex % 2) applyFill(cell, "FFF8FAFC");
    });
    excelRow.height = 20;
  });
  const lastRow = headerRow + displayRows.length;
  [5, 6, 7, 8].forEach((column) => {
    ws.getColumn(column).numFmt = "#,##0.0";
  });
  return lastRow;
}

function addChartData(ws, startColumn, header, rows, valueKey) {
  const valueHeader = valueKey === "downtime_hours" ? "Downtime hours" : "MTBF hours";
  ws.getCell(1, startColumn).value = header;
  ws.getCell(1, startColumn + 1).value = valueHeader;
  rows.forEach((row, index) => {
    ws.getCell(index + 2, startColumn).value = row.asset_code || row.asset_name || "Asset";
    ws.getCell(index + 2, startColumn + 1).value = number(row[valueKey]);
  });
  ws.getColumn(startColumn).hidden = true;
  ws.getColumn(startColumn + 1).hidden = true;
  const left = ws.getColumn(startColumn).letter;
  const right = ws.getColumn(startColumn + 1).letter;
  const last = rows.length + 1;
  return {
    categoryFormula: `'${ws.name}'!$${left}$2:$${left}$${last}`,
    valueFormula: `'${ws.name}'!$${right}$2:$${right}$${last}`,
  };
}

function getAttentionRows(byAsset) {
  return byAsset
    .filter((row) => number(row.failure_count) > 0 || number(row.downtime_hours) > 0)
    .sort((left, right) => (
      number(right.downtime_hours) - number(left.downtime_hours)
      || number(right.failure_count) - number(left.failure_count)
      || String(left.asset_code || "").localeCompare(String(right.asset_code || ""))
    ))
    .slice(0, 10)
    .map((row) => ({
      ...row,
      action: number(row.failure_count) > 0 ? "Review reliability" : "Review downtime",
    }));
}

/** Creates a decision-ready MTBF/LTTR workbook from the same report data used on screen. */
export async function buildReliabilityExecutiveWorkbook(data = {}, options = {}) {
  const startDate = String(options.startDate || "");
  const endDate = String(options.endDate || "");
  const category = String(options.category || "").trim();
  const assetRows = Array.isArray(data.by_asset) ? data.by_asset : [];
  const incidents = Array.isArray(data.incidents) ? data.incidents : [];
  const summaryData = data.summary || {};
  const scopeCount = number(data.asset_filter_count);
  const basis = data.downtime_basis === "asset_kpi_daily"
    ? "Asset KPI daily downtime"
    : "Recorded breakdown downtime";

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "IRONLOG";
  workbook.created = new Date();
  workbook.properties.date1904 = false;

  const { ws: summary, nextRow } = createManagementSummary(workbook, {
    title: "IRONLOG MTBF / LTTR Executive Report",
    periodLabel: `Reporting period: ${startDate} to ${endDate}`,
    cards: [
      { label: "Failures", value: number(summaryData.failure_count), numFmt: "#,##0", tone: "attention" },
      { label: "Operating hours", value: number(summaryData.operating_hours), numFmt: "#,##0.0" },
      { label: "Downtime hours", value: number(summaryData.downtime_hours), numFmt: "#,##0.0", tone: "warning" },
      { label: "MTBF hours", value: hoursOrBlank(summaryData.mtbf_hours), numFmt: "#,##0.0" },
      { label: "LTTR hours", value: hoursOrBlank(summaryData.lttr_hours), numFmt: "#,##0.0", tone: "attention" },
    ],
    scopeLines: [
      `Scope: ${scopeCount} asset(s)${category ? ` · Category: ${category}` : ""}.`,
      `Downtime basis: ${basis}. Recorded incident downtime remains available in the audit tab.`,
    ],
  });
  summary.name = "Exec Summary";
  summary.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  summary.pageMargins = { left: 0.25, right: 0.25, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 };
  summary.getColumn(1).width = 15;
  summary.getColumn(2).width = 28;
  summary.getColumn(3).width = 17;
  summary.getColumn(4).width = 11;
  summary.getColumn(5).width = 14;
  summary.getColumn(6).width = 12;
  summary.getColumn(7).width = 12;
  summary.getColumn(8).width = 14;
  summary.getColumn(9).width = 19;

  const chartStartRow = nextRow + 2;
  summary.mergeCells(chartStartRow, 1, chartStartRow, 15);
  const chartNote = summary.getCell(chartStartRow, 1);
  chartNote.value = "Chart values are linked to hidden source cells on this sheet and remain editable in Excel.";
  chartNote.font = { name: "Arial", size: 9, italic: true, color: { argb: COLORS.muted } };
  chartNote.alignment = { vertical: "middle" };
  summary.getRow(chartStartRow).height = 18;

  const chartBottomRow = chartStartRow + 16;
  for (let rowNumber = chartStartRow + 1; rowNumber <= chartBottomRow; rowNumber += 1) {
    summary.getRow(rowNumber).height = 18;
  }

  const attentionRows = getAttentionRows(assetRows);
  const attentionLastRow = writeExecutiveTable(summary, chartBottomRow + 3, attentionRows);
  summary.mergeCells(attentionLastRow + 2, 1, attentionLastRow + 2, 9);
  const note = summary.getCell(attentionLastRow + 2, 1);
  note.value = `Source: IRONLOG reliability data for ${startDate} to ${endDate}. LTTR uses ${basis.toLowerCase()}.`;
  note.font = { name: "Arial", size: 9, italic: true, color: { argb: COLORS.muted } };
  note.alignment = { vertical: "middle" };

  const topDowntime = assetRows
    .filter((row) => number(row.downtime_hours) > 0)
    .sort((left, right) => number(right.downtime_hours) - number(left.downtime_hours))
    .slice(0, 10);
  const lowestMtbf = assetRows
    .filter((row) => number(row.failure_count) > 0 && Number.isFinite(Number(row.mtbf_hours)))
    .sort((left, right) => number(left.mtbf_hours) - number(right.mtbf_hours))
    .slice(0, 10);
  const charts = [];
  if (topDowntime.length) {
    const source = addChartData(summary, 17, "Downtime asset", topDowntime, "downtime_hours");
    charts.push({
      title: "Breakdown downtime by equipment (hours)",
      seriesName: "Downtime hours",
      ...source,
      categories: topDowntime.map((row) => row.asset_code || row.asset_name || "Asset"),
      values: topDowntime.map((row) => number(row.downtime_hours)),
      color: "C55A11",
      position: { from: { col: 0, row: chartStartRow }, to: { col: 7, row: chartBottomRow } },
    });
  }
  if (lowestMtbf.length) {
    const source = addChartData(summary, 20, "MTBF asset", lowestMtbf, "mtbf_hours");
    charts.push({
      title: "Lowest MTBF by equipment (hours)",
      seriesName: "MTBF hours",
      ...source,
      categories: lowestMtbf.map((row) => row.asset_code || row.asset_name || "Asset"),
      values: lowestMtbf.map((row) => number(row.mtbf_hours)),
      color: "2F75B5",
      position: { from: { col: 8, row: chartStartRow }, to: { col: 15, row: chartBottomRow } },
    });
  }
  if (!charts.length) {
    const emptyMessage = summary.getCell(chartStartRow + 4, 1);
    emptyMessage.value = "No breakdown failures or downtime were recorded for the selected scope, so no reliability charts are shown.";
    emptyMessage.font = { name: "Arial", size: 10, color: { argb: COLORS.muted } };
  }

  const assets = workbook.addWorksheet("Asset Detail");
  assets.columns = [
    { header: "Asset", key: "asset_code", width: 14 },
    { header: "Equipment", key: "asset_name", width: 30 },
    { header: "Category", key: "category", width: 18 },
    { header: "Failures", key: "failure_count", width: 12 },
    { header: "Operating h", key: "operating_hours", width: 14 },
    { header: "Downtime h", key: "downtime_hours", width: 14 },
    { header: "Recorded h", key: "recorded_downtime_hours", width: 14 },
    { header: "MTBF h", key: "mtbf_hours", width: 12 },
    { header: "LTTR h", key: "lttr_hours", width: 12 },
  ];
  assets.addRows(assetRows.map((row) => ({
    ...row,
    mtbf_hours: hoursOrBlank(row.mtbf_hours),
    lttr_hours: hoursOrBlank(row.lttr_hours),
  })));
  styleManagementDetailSheet(assets, {
    title: "MTBF / LTTR by equipment",
    subtitle: `Reporting period: ${startDate} to ${endDate} · Downtime basis: ${basis}`,
    frozenColumns: 2,
    numberFormats: {
      operating_hours: "#,##0.0",
      downtime_hours: "#,##0.0",
      recorded_downtime_hours: "#,##0.0",
      mtbf_hours: "#,##0.0",
      lttr_hours: "#,##0.0",
    },
  });

  const incidentSheet = workbook.addWorksheet("Incident Audit");
  incidentSheet.columns = [
    { header: "Asset", key: "asset_code", width: 14 },
    { header: "Breakdown #", key: "breakdown_id", width: 12 },
    { header: "Report date", key: "breakdown_date", width: 14 },
    { header: "WO #", key: "work_order_id", width: 10 },
    { header: "Downtime h (period)", key: "downtime_hours", width: 18 },
    { header: "Source", key: "downtime_source", width: 18 },
    { header: "Log h (period)", key: "log_downtime_in_period", width: 14 },
    { header: "Header h (total)", key: "header_downtime_hours", width: 16 },
    { header: "WO opened", key: "work_order_opened_at", width: 20 },
    { header: "WO closed", key: "work_order_closed_at", width: 20 },
    { header: "Description", key: "description", width: 44 },
  ];
  incidentSheet.addRows(incidents);
  styleManagementDetailSheet(incidentSheet, {
    title: "Breakdown incident audit",
    subtitle: `Direct recorded downtime for incidents in the reporting period · ${startDate} to ${endDate}`,
    frozenColumns: 2,
    numberFormats: {
      downtime_hours: "#,##0.0",
      log_downtime_in_period: "#,##0.0",
      header_downtime_hours: "#,##0.0",
    },
  });

  const raw = await workbook.xlsx.writeBuffer();
  return addNativeBarCharts(raw, { sheetName: summary.name, charts });
}
