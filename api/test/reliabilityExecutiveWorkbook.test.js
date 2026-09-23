import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { buildReliabilityExecutiveWorkbook } from "../utils/reliabilityExecutiveWorkbook.js";

const sample = {
  asset_filter_count: 2,
  scheduled_fallback: 11,
  downtime_basis: "asset_kpi_daily",
  summary: {
    failure_count: 3,
    operating_hours: 154.6,
    downtime_hours: 22.5,
    recorded_downtime_hours: 14,
    mtbf_hours: 51.53,
    lttr_hours: 7.5,
  },
  by_asset: [
    {
      asset_code: "A300AM",
      asset_name: "Bell B30D",
      category: "ADT",
      failure_count: 2,
      operating_hours: 100.4,
      downtime_hours: 18,
      recorded_downtime_hours: 11,
      mtbf_hours: 50.2,
      lttr_hours: 9,
    },
    {
      asset_code: "G01AM",
      asset_name: "CAT 140H",
      category: "Grader",
      failure_count: 1,
      operating_hours: 54.2,
      downtime_hours: 4.5,
      recorded_downtime_hours: 3,
      mtbf_hours: 54.2,
      lttr_hours: 4.5,
    },
  ],
  incidents: [{
    asset_code: "A300AM",
    breakdown_id: 44,
    breakdown_date: "2026-08-19",
    work_order_id: 122,
    downtime_hours: 9,
    downtime_source: "daily logs",
    log_downtime_in_period: 9,
    header_downtime_hours: 9,
    work_order_opened_at: "2026-08-19 06:00",
    work_order_closed_at: "2026-08-19 15:00",
    description: "Hydraulic hose repair",
  }],
};

test("executive reliability workbook has a management summary, detail, audit, and native charts", async () => {
  const buffer = await buildReliabilityExecutiveWorkbook(sample, {
    startDate: "2026-08-01",
    endDate: "2026-08-31",
    category: "Production fleet",
  });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), ["Exec Summary", "Asset Detail", "Incident Audit"]);
  const summary = workbook.getWorksheet("Exec Summary");
  assert.equal(summary.getCell("A1").value, "IRONLOG MTBF / LTTR Executive Report");
  assert.equal(summary.getCell("A5").value, 3);
  assert.ok(summary.getColumn(1).values.includes("Equipment requiring attention"));
  assert.equal(workbook.getWorksheet("Asset Detail").getCell("A4").value, "Asset");
  assert.equal(workbook.getWorksheet("Incident Audit").getCell("A5").value, "A300AM");

  const zip = await JSZip.loadAsync(buffer);
  const chartParts = Object.keys(zip.files).filter((path) => /^xl\/charts\/chart\d+\.xml$/.test(path));
  assert.equal(chartParts.length, 2);
  const firstChart = await zip.file(chartParts[0]).async("string");
  assert.match(firstChart, /Breakdown downtime by equipment/);
  assert.match(firstChart, /Exec Summary/);
  const summaryXml = await zip.file("xl/worksheets/sheet1.xml").async("string");
  const summaryRels = await zip.file("xl/worksheets/_rels/sheet1.xml.rels").async("string");
  assert.match(summaryXml, /<drawing r:id="rId\d+"\/>/);
  assert.match(summaryRels, /relationships\/drawing/);
});
