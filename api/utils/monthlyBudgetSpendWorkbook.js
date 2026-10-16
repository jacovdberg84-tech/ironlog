import ExcelJS from "exceljs";
import { styleManagementDetailSheet, MANAGEMENT_XLSX } from "./managementWorkbook.js";
import { addNativeBarCharts } from "./nativeExcelCharts.js";

const { COLORS } = MANAGEMENT_XLSX;
const MONEY = '"$"#,##0.00;[Red]("$"#,##0.00);-';
const PERCENT = "0.0%";
const THIN = { style: "thin", color: { argb: COLORS.grid } };

function number(value, digits = 2) {
  const parsed = Number(value || 0);
  return Number.isFinite(parsed) ? Number(parsed.toFixed(digits)) : 0;
}

function formulaCell(cell, formula, result, numFmt = null) {
  cell.value = { formula: String(formula || "").replace(/^=/, ""), result };
  if (numFmt) cell.numFmt = numFmt;
  return cell;
}

function addSummaryCard(ws, row, startColumn, card) {
  const endColumn = startColumn + 1;
  const fill = card.tone === "attention"
    ? COLORS.amberLight
    : card.tone === "warning"
      ? COLORS.warningLight
      : COLORS.navyLight;
  ws.mergeCells(row, startColumn, row, endColumn);
  ws.mergeCells(row + 1, startColumn, row + 2, endColumn);
  const label = ws.getCell(row, startColumn);
  const value = ws.getCell(row + 1, startColumn);
  label.value = String(card.label || "");
  value.value = card.value ?? "";
  [label, value].forEach((cell) => {
    applyFill(cell, fill);
    cell.border = { top: THIN, bottom: THIN, left: THIN, right: THIN };
    cell.alignment = { vertical: "middle" };
  });
  label.font = { name: "Arial", size: 9, bold: true, color: { argb: COLORS.muted } };
  value.font = { name: "Arial", size: 16, bold: true, color: { argb: COLORS.navy } };
  value.alignment = { vertical: "middle", horizontal: "left" };
  if (card.numFmt) value.numFmt = card.numFmt;
}

function applyFill(cell, argb) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function addSection(ws, row, title, columns) {
  ws.mergeCells(row, 1, row, columns);
  const cell = ws.getCell(row, 1);
  cell.value = title;
  applyFill(cell, COLORS.blueLight);
  cell.font = { name: "Arial", size: 11, bold: true, color: { argb: COLORS.navy } };
  cell.alignment = { vertical: "middle" };
  cell.border = { top: THIN, right: THIN, bottom: THIN, left: THIN };
  ws.getRow(row).height = 21;
}

function styleSummaryTable(ws, headerRow, lastRow) {
  ws.getRow(headerRow).eachCell({ includeEmpty: true }, (cell) => {
    applyFill(cell, COLORS.navy);
    cell.font = { name: "Arial", size: 10, bold: true, color: { argb: COLORS.white } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = { top: THIN, right: THIN, bottom: THIN, left: THIN };
  });
  for (let row = headerRow + 1; row <= lastRow; row += 1) {
    ws.getRow(row).eachCell({ includeEmpty: true }, (cell) => {
      cell.font = { name: "Arial", size: 10, color: { argb: COLORS.body }, bold: row === lastRow };
      cell.alignment = { vertical: "middle", wrapText: cell.column === 5 };
      cell.border = { bottom: THIN };
    });
    if (row === lastRow) applyFill(ws.getCell(row, 1), COLORS.blueLight);
    if (row === lastRow) {
      for (let col = 2; col <= 5; col += 1) applyFill(ws.getCell(row, col), COLORS.blueLight);
    }
  }
}

function buildPlantCostingSheet(workbook, report) {
  const plantData = report.departmental?.plant || {};
  const assets = report.departmental?.plant_assets?.length
    ? report.departmental.plant_assets
    : [{ asset_code: "UNALLOCATED", asset_name: "No Plant revenue or cost records", revenue: 0, labour: 0, lube: 0, parts: 0, other_operating: 0, depreciation: 0, operating_cost: 0, total_cost: 0, net_contribution: 0, profit_margin_pct: null }];
  const ws = workbook.addWorksheet("Plant Costing", { views: [{ showGridLines: false }] });
  ws.columns = [
    { header: "Plant / Asset", key: "asset_code", width: 15 },
    { header: "Equipment", key: "asset_name", width: 30 },
    { header: "External Revenue", key: "revenue", width: 16 },
    { header: "Workshop Labour", key: "labour", width: 16 },
    { header: "Workshop Lubes", key: "lube", width: 16 },
    { header: "Spare Parts", key: "parts", width: 15 },
    { header: "Other Operating", key: "other_operating", width: 17 },
    { header: "Depreciation", key: "depreciation", width: 15 },
    { header: "Operating Cost", key: "operating_cost", width: 16 },
    { header: "Total Plant Cost", key: "total_cost", width: 17 },
    { header: "Net Contribution", key: "net_contribution", width: 18 },
    { header: "Profit Margin", key: "profit_margin_pct", width: 15 },
    { header: "25% Target", key: "target", width: 12 },
  ];
  ws.addRows(assets.map((row) => ({ ...row, target: 0.25 })));
  ws.addRow({ asset_code: "TOTAL" });
  styleManagementDetailSheet(ws, {
    title: `Plant Costing (${report.departmental?.plant_department || "AML-3001"})`,
    subtitle: `${report.start} to ${report.end} | Labour and lube are the matching Workshop internal charges. Parts remain direct Plant costs.`,
    frozenColumns: 2,
    numberFormats: { C: MONEY, D: MONEY, E: MONEY, F: MONEY, G: MONEY, H: MONEY, I: MONEY, J: MONEY, K: MONEY, L: PERCENT, M: PERCENT },
  });

  const firstDataRow = 5;
  const totalRow = firstDataRow + assets.length;
  const expected = report.departmental?.plant_assets || [];
  const expectedByCode = new Map(expected.map((row) => [String(row.asset_code || ""), row]));
  const colFormula = (sourceColumn, row) => `SUMIFS('Source Detail'!$${sourceColumn}:$${sourceColumn},'Source Detail'!$D:$D,$A${row})`;
  for (let index = 0; index < assets.length; index += 1) {
    const rowNumber = firstDataRow + index;
    const asset = expectedByCode.get(String(assets[index].asset_code || "")) || assets[index];
    formulaCell(ws.getCell(rowNumber, 3), `=${colFormula("J", rowNumber)}`, number(asset.revenue), MONEY);
    formulaCell(ws.getCell(rowNumber, 4), `=${colFormula("K", rowNumber)}`, number(asset.labour), MONEY);
    formulaCell(ws.getCell(rowNumber, 5), `=${colFormula("L", rowNumber)}`, number(asset.lube), MONEY);
    formulaCell(ws.getCell(rowNumber, 6), `=${colFormula("M", rowNumber)}`, number(asset.parts), MONEY);
    formulaCell(ws.getCell(rowNumber, 7), `=${colFormula("N", rowNumber)}`, number(asset.other_operating), MONEY);
    formulaCell(ws.getCell(rowNumber, 8), `=${colFormula("O", rowNumber)}`, number(asset.depreciation), MONEY);
    formulaCell(ws.getCell(rowNumber, 9), `=SUM(D${rowNumber}:G${rowNumber})`, number(asset.operating_cost), MONEY);
    formulaCell(ws.getCell(rowNumber, 10), `=I${rowNumber}+H${rowNumber}`, number(asset.total_cost), MONEY);
    formulaCell(ws.getCell(rowNumber, 11), `=C${rowNumber}-J${rowNumber}`, number(asset.net_contribution), MONEY);
    formulaCell(ws.getCell(rowNumber, 12), `=IF(C${rowNumber}>0,K${rowNumber}/C${rowNumber},"")`, asset.profit_margin_pct == null ? "" : number(asset.profit_margin_pct) / 100, PERCENT);
    formulaCell(ws.getCell(rowNumber, 13), "=0.25", 0.25, PERCENT);
  }
  for (let column = 3; column <= 11; column += 1) {
    const letter = ws.getColumn(column).letter;
    const totalResult = column === 3 ? number(plantData.revenue)
      : column === 4 ? number(plantData.labour)
        : column === 5 ? number(plantData.lube)
          : column === 6 ? number(plantData.parts)
            : column === 7 ? number(plantData.other_operating)
              : column === 8 ? number(plantData.depreciation)
                : column === 9 ? number(plantData.operating_cost)
                  : column === 10 ? number(plantData.total_cost)
                    : number(plantData.net_contribution);
    formulaCell(ws.getCell(totalRow, column), `=SUM(${letter}${firstDataRow}:${letter}${totalRow - 1})`, totalResult, MONEY);
  }
  formulaCell(ws.getCell(totalRow, 12), `=IF(C${totalRow}>0,K${totalRow}/C${totalRow},"")`, plantData.profit_margin_pct == null ? "" : number(plantData.profit_margin_pct) / 100, PERCENT);
  formulaCell(ws.getCell(totalRow, 13), "=0.25", 0.25, PERCENT);
  ws.addConditionalFormatting({
    ref: `L${firstDataRow}:L${totalRow}`,
    rules: [{ type: "cellIs", operator: "lessThan", formulae: [0.25], style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } }, font: { bold: true, color: { argb: "FFB91C1C" } } } }],
  });
  ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: totalRow, column: 13 } };
  ws.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  return { ws, totalRow, firstDataRow, assets };
}

function buildWorkshopCostingSheet(workbook, report) {
  const workshop = report.departmental?.workshop || {};
  const ws = workbook.addWorksheet("Workshop Costing", { views: [{ showGridLines: false }] });
  ws.columns = [
    { header: "Workshop Cost Category", key: "category", width: 34 },
    { header: "Actual Workshop Cost", key: "cost", width: 20 },
    { header: "Internal Charge Recovered", key: "recovery", width: 24 },
    { header: "Cost-Recovery Balance", key: "balance", width: 22 },
    { header: "Treatment", key: "comment", width: 49 },
  ];
  const rows = [
    { category: "Mechanics labour", sourceColumn: "P", sourceType: "Mechanic", cost: number(workshop.labour_cost), recovery: number((report.departmental?.source_lines || []).filter((r) => Number(r.workshop_labour_cost || 0) !== 0).reduce((sum, r) => sum + Number(r.workshop_recovery || 0), 0)), comment: "Charged to Plant only when an asset is linked." },
    { category: "Lubricant supply", sourceColumn: "Q", sourceType: "Lube", cost: number(workshop.lube_cost), recovery: number((report.departmental?.source_lines || []).filter((r) => Number(r.workshop_lube_cost || 0) !== 0).reduce((sum, r) => sum + Number(r.workshop_recovery || 0), 0)), comment: "Recovered from Plant as an internal lubricant charge." },
    { category: "Import / export duties", sourceColumn: "R", sourceType: "", cost: number(workshop.duties), recovery: 0, comment: "Manual Workshop cost. Not treated as Plant revenue unless separately allocated." },
    { category: "Other Workshop operating cost", sourceColumn: "S", sourceType: "", cost: number(workshop.other_operating), recovery: 0, comment: "Includes Workshop-only parts or manual operating entries." },
  ];
  ws.addRows(rows);
  ws.addRow({ category: "TOTAL" });
  styleManagementDetailSheet(ws, {
    title: `Workshop Costing (${report.departmental?.workshop_department || "AML-3002"})`,
    subtitle: `${report.start} to ${report.end} | Internal charges are cost recovery, not commercial revenue.`,
    numberFormats: { B: MONEY, C: MONEY, D: MONEY },
  });
  const firstDataRow = 5;
  const totalRow = firstDataRow + rows.length;
  for (let index = 0; index < rows.length; index += 1) {
    const rowNumber = firstDataRow + index;
    const row = rows[index];
    formulaCell(ws.getCell(rowNumber, 2), `=SUM('Source Detail'!$${row.sourceColumn}:$${row.sourceColumn})`, row.cost, MONEY);
    const recoveryFormula = row.sourceType
      ? `=SUMIFS('Source Detail'!$T:$T,'Source Detail'!$${row.sourceColumn}:$${row.sourceColumn},">0")`
      : "=0";
    formulaCell(ws.getCell(rowNumber, 3), recoveryFormula, row.recovery, MONEY);
    formulaCell(ws.getCell(rowNumber, 4), `=C${rowNumber}-B${rowNumber}`, number(row.recovery - row.cost), MONEY);
  }
  formulaCell(ws.getCell(totalRow, 2), `=SUM(B${firstDataRow}:B${totalRow - 1})`, number(workshop.total_operating_cost), MONEY);
  formulaCell(ws.getCell(totalRow, 3), `=SUM(C${firstDataRow}:C${totalRow - 1})`, number(workshop.recovery), MONEY);
  formulaCell(ws.getCell(totalRow, 4), `=C${totalRow}-B${totalRow}`, number(workshop.net_cost_recovery_balance), MONEY);
  ws.addConditionalFormatting({
    ref: `D${firstDataRow}:D${totalRow}`,
    rules: [
      { type: "cellIs", operator: "lessThan", formulae: [0], style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } }, font: { bold: true, color: { argb: "FFB91C1C" } } } },
      { type: "cellIs", operator: "greaterThan", formulae: [0], style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } }, font: { bold: true, color: { argb: "FFB45309" } } } },
    ],
  });
  ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: totalRow, column: 5 } };
  ws.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  return { ws, totalRow };
}

function buildSourceDetailSheet(workbook, report) {
  const sourceRows = report.departmental?.source_lines || [];
  const displayRows = sourceRows.length ? sourceRows : [{
    date: "", source_id: "", source_type: "", asset_code: "", asset_name: "", description: "No source records for the selected month.",
  }];
  const ws = workbook.addWorksheet("Source Detail", { views: [{ showGridLines: false }] });
  ws.columns = [
    { header: "Date", key: "date", width: 13 }, { header: "Source ID", key: "source_id", width: 27 }, { header: "Source", key: "source_type", width: 25 },
    { header: "Asset", key: "asset_code", width: 14 }, { header: "Equipment", key: "asset_name", width: 28 }, { header: "Description", key: "description", width: 48 },
    { header: "Qty / Hrs", key: "quantity", width: 12 }, { header: "Unit Cost", key: "unit_rate", width: 14 }, { header: "Source Amount", key: "source_amount", width: 16 },
    { header: "Plant Revenue", key: "plant_revenue", width: 16 }, { header: "Plant Labour Charge", key: "plant_labour_charge", width: 19 }, { header: "Plant Lube Charge", key: "plant_lube_charge", width: 18 },
    { header: "Plant Parts", key: "plant_parts", width: 15 }, { header: "Plant Other Op", key: "plant_other_operating", width: 18 }, { header: "Plant Depreciation", key: "plant_depreciation", width: 18 },
    { header: "Workshop Labour Cost", key: "workshop_labour_cost", width: 21 }, { header: "Workshop Lube Cost", key: "workshop_lube_cost", width: 20 }, { header: "Workshop Duties", key: "workshop_duties", width: 18 },
    { header: "Workshop Other Op", key: "workshop_other_operating", width: 20 }, { header: "Workshop Recovery", key: "workshop_recovery", width: 19 },
  ];
  ws.addRows(displayRows);
  styleManagementDetailSheet(ws, {
    title: "Source Detail and Departmental Allocation",
    subtitle: `${report.start} to ${report.end} | Original operational sources plus the allocation used by this workbook.`,
    frozenColumns: 4,
    numberFormats: { G: "#,##0.000", H: MONEY, I: MONEY, J: MONEY, K: MONEY, L: MONEY, M: MONEY, N: MONEY, O: MONEY, P: MONEY, Q: MONEY, R: MONEY, S: MONEY, T: MONEY },
  });
  ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: ws.rowCount, column: 20 } };
  ws.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  return { ws, sourceRows };
}

function buildReconciliationSheet(ws, report, sourceRows) {
  const rows = report.departmental?.reconciliation || [];
  const sourceRowById = new Map(sourceRows.map((row, index) => [String(row.source_id || ""), index + 5]));
  const displayRows = rows.length ? rows : [{ source_id: "", asset_code: "", description: "No internal labour or lube charges were found for this period.", workshop_recovery: 0, plant_expense: 0, difference: 0, status: "No internal charges" }];
  ws.columns = [
    { header: "Source Reference", key: "source_id", width: 28 }, { header: "Plant Asset", key: "asset_code", width: 15 }, { header: "Description", key: "description", width: 50 },
    { header: "Workshop Recovery", key: "workshop_recovery", width: 20 }, { header: "Plant Expense", key: "plant_expense", width: 18 }, { header: "Difference", key: "difference", width: 16 }, { header: "Status", key: "status", width: 20 },
  ];
  ws.addRows(displayRows);
  ws.addRow({ source_id: "TOTAL" });
  styleManagementDetailSheet(ws, {
    title: "Interdepartmental Reconciliation",
    subtitle: "Workshop recovery must equal the matching Plant expense. Internal charges are eliminated in the consolidated result.",
    frozenColumns: 2,
    numberFormats: { D: MONEY, E: MONEY, F: MONEY },
  });
  const firstDataRow = 5;
  const totalRow = firstDataRow + displayRows.length;
  for (let index = 0; index < displayRows.length; index += 1) {
    const rowNumber = firstDataRow + index;
    const row = displayRows[index];
    const sourceRow = sourceRowById.get(String(row.source_id || ""));
    if (sourceRow) {
      formulaCell(ws.getCell(rowNumber, 4), `='Source Detail'!$T$${sourceRow}`, number(row.workshop_recovery), MONEY);
      formulaCell(ws.getCell(rowNumber, 5), `=SUM('Source Detail'!$K$${sourceRow}:'Source Detail'!$L$${sourceRow})`, number(row.plant_expense), MONEY);
      formulaCell(ws.getCell(rowNumber, 6), `=D${rowNumber}-E${rowNumber}`, number(row.difference), MONEY);
      const status = String(row.asset_code || "").toUpperCase() === "UNLINKED" ? "Missing plant asset" : Math.abs(number(row.difference)) < 0.01 ? "Matched" : "Unmatched";
      formulaCell(ws.getCell(rowNumber, 7), `=IF($B${rowNumber}="UNLINKED","Missing plant asset",IF(COUNTIF($A$${firstDataRow}:$A$${totalRow - 1},$A${rowNumber})>1,"Duplicate source",IF(ABS($F${rowNumber})<0.01,"Matched","Unmatched")))`, status);
    }
  }
  formulaCell(ws.getCell(totalRow, 4), `=SUM(D${firstDataRow}:D${totalRow - 1})`, number(rows.reduce((sum, row) => sum + number(row.workshop_recovery), 0)), MONEY);
  formulaCell(ws.getCell(totalRow, 5), `=SUM(E${firstDataRow}:E${totalRow - 1})`, number(rows.reduce((sum, row) => sum + number(row.plant_expense), 0)), MONEY);
  formulaCell(ws.getCell(totalRow, 6), `=D${totalRow}-E${totalRow}`, number(rows.reduce((sum, row) => sum + number(row.difference), 0)), MONEY);
  ws.getCell(totalRow, 7).value = "Totals";
  ws.addConditionalFormatting({
    ref: `F${firstDataRow}:F${totalRow}`,
    rules: [{ type: "cellIs", operator: "notEqual", formulae: [0], style: { fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } }, font: { bold: true, color: { argb: "FFB91C1C" } } } }],
  });
  ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: totalRow, column: 7 } };
  ws.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  return { ws, totalRow };
}

function buildManagementSummary(ws, report, plantSheet, workshopSheet) {
  const plant = report.departmental?.plant || {};
  const workshop = report.departmental?.workshop || {};
  const margin = plant.profit_margin_pct == null ? null : number(plant.profit_margin_pct) / 100;
  const budgetTone = report.variance == null ? "attention" : report.variance > 0 ? "warning" : "normal";
  ws.properties.defaultRowHeight = 18;
  ws.columns = Array.from({ length: 12 }, (_, index) => ({ width: [18, 14, 3][index % 3] }));
  ws.mergeCells(1, 1, 1, 12);
  ws.mergeCells(2, 1, 2, 12);
  ws.getCell(1, 1).value = "IRONLOG Departmental Costing Report";
  ws.getCell(1, 1).font = { name: "Arial", size: 16, bold: true, color: { argb: COLORS.navy } };
  ws.getCell(1, 1).alignment = { vertical: "middle" };
  ws.getCell(2, 1).value = `${report.start} to ${report.end} | Site: ${String(report.site_code || "main").toUpperCase()} | USD`;
  ws.getCell(2, 1).font = { name: "Arial", size: 10, color: { argb: COLORS.muted } };
  ws.getCell(2, 1).alignment = { vertical: "middle" };
  ws.getRow(1).height = 26;
  ws.getRow(2).height = 18;
  const cards = [
    { label: "PLANT HIRE REVENUE", value: { formula: `'Plant Costing'!$C$${plantSheet.totalRow}`, result: number(plant.revenue) }, numFmt: MONEY, tone: "normal" },
    { label: "PLANT OPERATING COST", value: { formula: `'Plant Costing'!$I$${plantSheet.totalRow}`, result: number(plant.operating_cost) }, numFmt: MONEY, tone: "normal" },
    { label: "PLANT PROFIT MARGIN", value: margin == null ? "No revenue recorded" : { formula: `'Plant Costing'!$L$${plantSheet.totalRow}`, result: margin }, numFmt: margin == null ? undefined : PERCENT, tone: margin != null && margin < 0.25 ? "warning" : "normal" },
    { label: "WORKSHOP RECOVERY BALANCE", value: { formula: `'Workshop Costing'!$D$${workshopSheet.totalRow}`, result: number(workshop.net_cost_recovery_balance) }, numFmt: MONEY, tone: number(workshop.net_cost_recovery_balance) < 0 ? "attention" : "normal" },
    { label: "CONSOLIDATED RESULT", value: { formula: `'Plant Costing'!$K$${plantSheet.totalRow}+'Workshop Costing'!$D$${workshopSheet.totalRow}`, result: number(report.departmental?.consolidated_result) }, numFmt: MONEY, tone: number(report.departmental?.consolidated_result) < 0 ? "attention" : "normal" },
    { label: "MONTHLY OPERATING BUDGET", value: report.budget?.is_set ? number(report.budget.amount) : "Not set", numFmt: report.budget?.is_set ? MONEY : undefined, tone: report.budget?.is_set ? "normal" : "attention" },
    { label: "BUDGET VARIANCE", value: report.variance == null ? "Budget not set" : { formula: `'Plant Costing'!$I$${plantSheet.totalRow}-${number(report.budget.amount)}`, result: number(report.variance) }, numFmt: report.variance == null ? undefined : MONEY, tone: budgetTone },
    { label: "PLANT TARGET MARGIN", value: { formula: "0.25", result: 0.25 }, numFmt: PERCENT, tone: "normal" },
  ];
  const starts = [1, 4, 7, 10];
  cards.forEach((card, index) => addSummaryCard(ws, 4 + Math.floor(index / 4) * 4, starts[index % 4], card));
  const scopeRow = 13;
  ws.mergeCells(scopeRow, 1, scopeRow, 12);
  ws.getCell(scopeRow, 1).value = "Scope";
  applyFill(ws.getCell(scopeRow, 1), COLORS.blueLight);
  ws.getCell(scopeRow, 1).font = { name: "Arial", size: 10, bold: true, color: { argb: COLORS.navy } };
  ws.getCell(scopeRow, 1).border = { top: THIN, bottom: THIN, left: THIN, right: THIN };
  const scopeLines = [
    `Plant Department: ${report.departmental?.plant_department || "AML-3001"}. Workshop Department: ${report.departmental?.workshop_department || "AML-3002"}.`,
    "Labour and lubricant charges are matched between Workshop recovery and Plant expense, then eliminated in the consolidated result.",
    "Depreciation, import/export duties and other operating costs are shown only when a source or a labelled monthly departmental entry exists.",
  ];
  scopeLines.forEach((line, index) => {
    ws.mergeCells(scopeRow + 1 + index, 1, scopeRow + 1 + index, 12);
    const cell = ws.getCell(scopeRow + 1 + index, 1);
    cell.value = line;
    cell.font = { name: "Arial", size: 10, color: { argb: COLORS.body } };
    cell.alignment = { vertical: "middle" };
  });
  ws.getRow(scopeRow).height = 20;
  ws.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  const start = scopeRow + 1 + scopeLines.length + 2;
  addSection(ws, start, "Departmental financial position", 5);
  const headerRow = start + 1;
  ["Measure", "Plant AML-3001", "Workshop AML-3002", "Consolidated", "Treatment"].forEach((header, index) => { ws.getCell(headerRow, index + 1).value = header; });
  const rows = [
    ["External plant hire revenue", `='Plant Costing'!$C$${plantSheet.totalRow}`, number(plant.revenue), "=0", 0, `=B${headerRow + 1}`, number(plant.revenue), "External revenue only; never a Workshop recovery."],
    ["Plant total operating cost", `='Plant Costing'!$I$${plantSheet.totalRow}`, number(plant.operating_cost), "=0", 0, `=B${headerRow + 2}`, number(plant.operating_cost), "Includes Workshop charges, parts and other Plant operating costs."],
    ["Plant depreciation", `='Plant Costing'!$H$${plantSheet.totalRow}`, number(plant.depreciation), "=0", 0, `=B${headerRow + 3}`, number(plant.depreciation), "Kept separate from operating expenditure."],
    ["Workshop operating cost", "=0", 0, `='Workshop Costing'!$B$${workshopSheet.totalRow}`, number(workshop.total_operating_cost), `=C${headerRow + 4}`, number(workshop.total_operating_cost), "Actual Workshop cost; not commercial revenue."],
    ["Internal Workshop charges eliminated", `='Plant Costing'!$D$${plantSheet.totalRow}+'Plant Costing'!$E$${plantSheet.totalRow}`, number(plant.labour + plant.lube), `='Workshop Costing'!$C$${workshopSheet.totalRow}`, number(workshop.recovery), `=C${headerRow + 5}-B${headerRow + 5}`, 0, "Plant expense and Workshop recovery cancel at company level."],
    ["Net result after internal elimination", `='Plant Costing'!$K$${plantSheet.totalRow}`, number(plant.net_contribution), `='Workshop Costing'!$D$${workshopSheet.totalRow}`, number(workshop.net_cost_recovery_balance), `=B${headerRow + 6}+C${headerRow + 6}`, number(report.departmental?.consolidated_result), "Plant contribution plus Workshop cost-recovery balance."],
  ];
  rows.forEach((row, index) => {
    const rowNumber = headerRow + 1 + index;
    ws.getCell(rowNumber, 1).value = row[0];
    formulaCell(ws.getCell(rowNumber, 2), row[1], row[2], MONEY);
    formulaCell(ws.getCell(rowNumber, 3), row[3], row[4], MONEY);
    formulaCell(ws.getCell(rowNumber, 4), row[5], row[6], MONEY);
    ws.getCell(rowNumber, 5).value = row[7];
  });
  const finalRow = headerRow + rows.length;
  styleSummaryTable(ws, headerRow, finalRow);
  ws.getColumn(1).width = 34; ws.getColumn(2).width = 20; ws.getColumn(3).width = 22; ws.getColumn(4).width = 20; ws.getColumn(5).width = 55;
  return { ws, chartStartRow: finalRow + 3 };
}

/** Builds the formula-driven departmental costing workbook. */
export async function buildMonthlyBudgetSpendWorkbook(report) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "IRONLOG";
  workbook.created = new Date();
  workbook.properties.title = "IRONLOG Departmental Costing Report";
  workbook.properties.subject = `Plant and Workshop costing for ${report.month}`;
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.calcProperties.forceFullCalc = true;

  const summaryWs = workbook.addWorksheet("Management Summary", { views: [{ showGridLines: false }] });
  const plantSheet = buildPlantCostingSheet(workbook, report);
  const workshopSheet = buildWorkshopCostingSheet(workbook, report);
  // Excel sheet names are capped at 31 characters; the full label remains in the sheet title.
  const reconciliationWs = workbook.addWorksheet("Dept Reconciliation", { views: [{ showGridLines: false }] });
  const sourceSheet = buildSourceDetailSheet(workbook, report);
  buildReconciliationSheet(reconciliationWs, report, sourceSheet.sourceRows);
  const summary = buildManagementSummary(summaryWs, report, plantSheet, workshopSheet);
  return { workbook, chartStartRow: summary.chartStartRow, plantSheet };
}

/** Exports the executive workbook with a native, editable chart. */
export async function buildMonthlyBudgetSpendBuffer(report) {
  const { workbook, chartStartRow, plantSheet } = await buildMonthlyBudgetSpendWorkbook(report);
  const plantAssets = (report.departmental?.plant_assets || []).filter((row) => String(row.asset_code || "").toUpperCase() !== "UNALLOCATED").slice(0, 10);
  const raw = await workbook.xlsx.writeBuffer();
  if (!plantAssets.length) return Buffer.from(raw);
  return addNativeBarCharts(raw, {
    sheetName: "Management Summary",
    charts: [{
      title: "Plant net contribution by equipment (USD)", seriesName: "Net contribution",
      categoryFormula: `'Plant Costing'!$A$${plantSheet.firstDataRow}:$A$${plantSheet.firstDataRow + plantAssets.length - 1}`,
      valueFormula: `'Plant Costing'!$K$${plantSheet.firstDataRow}:$K$${plantSheet.firstDataRow + plantAssets.length - 1}`,
      categories: plantAssets.map((row) => String(row.asset_code || "Asset")), values: plantAssets.map((row) => number(row.net_contribution)), color: "2F75B5",
      position: { from: { col: 0, row: chartStartRow }, to: { col: 10, row: chartStartRow + 16 } },
    }],
  });
}
