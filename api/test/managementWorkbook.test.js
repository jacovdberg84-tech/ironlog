import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { createManagementSummary, styleManagementDetailSheet } from "../utils/managementWorkbook.js";

test("management workbook helper creates a summary and formats detail sheets consistently", async () => {
  const wb = new ExcelJS.Workbook();
  const { ws: summary } = createManagementSummary(wb, {
    title: "IRONLOG Lube Usage Report",
    periodLabel: "Reporting period: 01 August 2026 to 31 August 2026",
    cards: [
      { label: "Total quantity", value: 120, numFmt: "#,##0.00" },
      { label: "Estimated cost", value: 640, numFmt: "#,##0.00" },
    ],
    scopeLines: ["Asset usage and store stock movements."],
  });
  const detail = wb.addWorksheet("Usage by asset");
  detail.columns = [
    { header: "Asset", key: "asset", width: 14 },
    { header: "Availability %", key: "availability", width: 14 },
    { header: "Cost", key: "cost", width: 14 },
  ];
  detail.addRows([
    { asset: "A300AM", availability: 79.5, cost: 10.5 },
    { asset: "TOTAL", availability: 84.5, cost: 20.5 },
  ]);
  styleManagementDetailSheet(detail, {
    title: "Lube usage by asset",
    subtitle: "Reporting period: August 2026",
    frozenColumns: 1,
    percentageColumns: ["availability"],
    numberFormats: { cost: "#,##0.00" },
  });

  assert.equal(summary.getCell("A1").value, "IRONLOG Lube Usage Report");
  assert.equal(detail.getCell("A1").value, "Lube usage by asset");
  assert.equal(detail.getCell("A4").value, "Asset");
  assert.equal(detail.getCell("B5").numFmt, '0.0"%"');
  assert.equal(detail.views[0].ySplit, 4);
  assert.equal(detail.views[0].xSplit, 1);
  assert.equal(detail.getCell("A6").font.bold, true);
  assert.ok(detail.conditionalFormattings.length > 0);

  const buffer = await wb.xlsx.writeBuffer();
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(buffer);
  assert.deepEqual(reopened.worksheets.map((ws) => ws.name), ["Summary", "Usage by asset"]);
  assert.equal(reopened.getWorksheet("Usage by asset").getCell("A4").value, "Asset");
});
