const COLORS = Object.freeze({
  navy: "FF17365D",
  navyLight: "FFEAF2F8",
  blueLight: "FFD9EAF7",
  amberLight: "FFFFF2CC",
  warningLight: "FFFCE4D6",
  body: "FF1F2937",
  muted: "FF64748B",
  grid: "FFD9E2F3",
  white: "FFFFFFFF",
});

const thinBorder = { style: "thin", color: { argb: COLORS.grid } };

function applyFill(cell, argb) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function applyCard(ws, row, startColumn, card) {
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
    cell.border = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
    cell.alignment = { vertical: "middle" };
  });
  label.font = { name: "Arial", size: 9, bold: true, color: { argb: COLORS.muted } };
  value.font = { name: "Arial", size: 16, bold: true, color: { argb: COLORS.navy } };
  value.alignment = { vertical: "middle", horizontal: "left" };
  if (card.numFmt) value.numFmt = card.numFmt;
}

/**
 * Creates the common first-sheet layout used by Ironlog management exports.
 * The reporting data remains in each export's existing detail sheets.
 */
export function createManagementSummary(workbook, {
  title,
  periodLabel,
  cards = [],
  scopeLines = [],
} = {}) {
  const ws = workbook.addWorksheet("Summary", { views: [{ showGridLines: false }] });
  ws.properties.defaultRowHeight = 18;
  ws.columns = Array.from({ length: 12 }, (_, index) => ({ width: [18, 14, 3][index % 3] }));

  ws.mergeCells(1, 1, 1, 12);
  ws.mergeCells(2, 1, 2, 12);
  const titleCell = ws.getCell(1, 1);
  titleCell.value = String(title || "IRONLOG Management Report");
  titleCell.font = { name: "Arial", size: 16, bold: true, color: { argb: COLORS.navy } };
  titleCell.alignment = { vertical: "middle" };
  const periodCell = ws.getCell(2, 1);
  periodCell.value = String(periodLabel || "");
  periodCell.font = { name: "Arial", size: 10, color: { argb: COLORS.muted } };
  periodCell.alignment = { vertical: "middle" };
  ws.getRow(1).height = 26;
  ws.getRow(2).height = 18;

  const starts = [1, 4, 7, 10];
  cards.slice(0, 8).forEach((card, index) => {
    applyCard(ws, 4 + Math.floor(index / 4) * 4, starts[index % 4], card);
  });

  const scopeRow = 4 + Math.max(1, Math.ceil(Math.min(cards.length, 8) / 4)) * 4 + 1;
  ws.mergeCells(scopeRow, 1, scopeRow, 12);
  const scopeTitle = ws.getCell(scopeRow, 1);
  scopeTitle.value = "Scope";
  applyFill(scopeTitle, COLORS.blueLight);
  scopeTitle.font = { name: "Arial", size: 10, bold: true, color: { argb: COLORS.navy } };
  scopeTitle.border = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
  scopeLines.filter(Boolean).forEach((line, index) => {
    ws.mergeCells(scopeRow + 1 + index, 1, scopeRow + 1 + index, 12);
    const cell = ws.getCell(scopeRow + 1 + index, 1);
    cell.value = String(line);
    cell.font = { name: "Arial", size: 10, color: { argb: COLORS.body } };
    cell.alignment = { vertical: "middle" };
  });

  ws.getRow(scopeRow).height = 20;
  return { ws, nextRow: scopeRow + 1 + scopeLines.filter(Boolean).length };
}

/** Adds the shared title, freeze pane, header, and attention formatting to an existing detail sheet. */
export function styleManagementDetailSheet(ws, {
  title,
  subtitle = "",
  frozenColumns = 0,
  percentageColumns = [],
  numberFormats = {},
} = {}) {
  const columns = Math.max(1, ws.columnCount || ws.actualColumnCount || 1);
  ws.spliceRows(1, 0, [String(title || ws.name || "Detail")]);
  ws.spliceRows(2, 0, [String(subtitle || "")]);
  ws.spliceRows(3, 0, []);
  ws.mergeCells(1, 1, 1, columns);
  ws.mergeCells(2, 1, 2, columns);

  const titleCell = ws.getCell(1, 1);
  titleCell.font = { name: "Arial", size: 14, bold: true, color: { argb: COLORS.navy } };
  titleCell.alignment = { vertical: "middle" };
  const subtitleCell = ws.getCell(2, 1);
  subtitleCell.font = { name: "Arial", size: 10, color: { argb: COLORS.muted } };
  subtitleCell.alignment = { vertical: "middle" };
  ws.getRow(1).height = 24;
  ws.getRow(2).height = 18;

  const headerRow = 4;
  const header = ws.getRow(headerRow);
  header.height = 22;
  header.eachCell({ includeEmpty: true }, (cell) => {
    applyFill(cell, COLORS.navy);
    cell.font = { name: "Arial", size: 10, bold: true, color: { argb: COLORS.white } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = {
      top: { style: "thin", color: { argb: COLORS.white } },
      bottom: { style: "thin", color: { argb: COLORS.white } },
      left: { style: "thin", color: { argb: COLORS.white } },
      right: { style: "thin", color: { argb: COLORS.white } },
    };
  });

  Object.entries(numberFormats).forEach(([key, numFmt]) => {
    const column = ws.getColumn(key);
    if (column) column.numFmt = numFmt;
  });
  percentageColumns.forEach((key) => {
    const column = ws.getColumn(key);
    if (column) column.numFmt = '0.0"%"';
  });

  for (let rowNumber = headerRow + 1; rowNumber <= ws.rowCount; rowNumber += 1) {
    const row = ws.getRow(rowNumber);
    const first = String(row.getCell(1).value || "").trim().toUpperCase();
    const isTotal = first === "TOTAL";
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.font = { name: "Arial", size: 10, bold: isTotal, color: { argb: COLORS.body } };
      cell.alignment = { vertical: "middle" };
      cell.border = { bottom: thinBorder };
      if (isTotal) applyFill(cell, COLORS.blueLight);
    });
  }

  const lastRow = ws.rowCount;
  percentageColumns.forEach((key) => {
    const column = ws.getColumn(key);
    if (!column || lastRow <= headerRow) return;
    ws.addConditionalFormatting({
      ref: `${column.letter}${headerRow + 1}:${column.letter}${lastRow}`,
      rules: [{
        type: "cellIs",
        operator: "lessThan",
        formulae: [80],
        style: {
          fill: { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.warningLight } },
          font: { bold: true, color: { argb: "FF9C0006" } },
        },
      }],
    });
  });

  ws.views = [{ state: "frozen", ySplit: headerRow, xSplit: frozenColumns, showGridLines: false }];
  return { headerRow, lastRow };
}

export const MANAGEMENT_XLSX = Object.freeze({ COLORS });
