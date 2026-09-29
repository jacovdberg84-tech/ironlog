import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import {
  buildWorkshopInventoryReportWorkbook,
  resolveWorkshopInventoryReportPeriod,
} from "../utils/workshopInventoryReport.js";

const sampleData = {
  items: [
    {
      part_code: "KIT-500",
      part_name: "500 hour service kit",
      stock_category: "part",
      supplier: "AML Spares",
      location: "Main stores",
      critical: true,
      min_qty: 2,
      max_qty: 6,
      reorder_point: 2,
      opening_qty: 4,
      receipts_qty: 0,
      issues_qty: 4,
      returns_qty: 0,
      transfers_in_qty: 0,
      transfers_out_qty: 0,
      adjustments_qty: 0,
      closing_qty: 0,
      unit_cost: 125.5,
      last_movement_at: "2026-09-27 10:15:00",
      is_lube: false,
    },
    {
      part_code: "BOLT-M16",
      part_name: "M16 flange bolt",
      stock_category: "part",
      supplier: "",
      location: "Main stores",
      critical: false,
      min_qty: 10,
      max_qty: 50,
      reorder_point: 10,
      opening_qty: 30,
      receipts_qty: 10,
      issues_qty: 5,
      returns_qty: 0,
      transfers_in_qty: 0,
      transfers_out_qty: 0,
      adjustments_qty: 0,
      closing_qty: 35,
      unit_cost: 1.25,
      last_movement_at: "2026-09-26 09:00:00",
      is_lube: false,
    },
    {
      part_code: "OIL-10W40",
      part_name: "10W40 engine oil",
      stock_category: "oil",
      supplier: "Oil supplier",
      location: "Lube store",
      critical: false,
      min_qty: 40,
      max_qty: 200,
      reorder_point: 40,
      opening_qty: 120,
      receipts_qty: 40,
      issues_qty: 30,
      returns_qty: 0,
      transfers_in_qty: 0,
      transfers_out_qty: 0,
      adjustments_qty: 0,
      closing_qty: 130,
      unit_cost: 4.63,
      last_movement_at: "2026-09-25 13:45:00",
      is_lube: true,
    },
  ],
  movements: [
    {
      movement_at: "2026-09-27 10:15:00",
      movement_type: "issue_to_work_order",
      quantity: -4,
      reference: "WO-2026-091",
      part_code: "KIT-500",
      part_name: "500 hour service kit",
      location_code: "Main stores",
      bin_code: "A-01",
      unit_cost: 125.5,
    },
  ],
};

test("workshop inventory report resolves weekly and monthly periods predictably", () => {
  assert.deepEqual(resolveWorkshopInventoryReportPeriod("weekly", "2026-09-28"), {
    report_type: "weekly",
    start_date: "2026-09-22",
    end_date: "2026-09-28",
  });
  assert.deepEqual(resolveWorkshopInventoryReportPeriod("monthly", "2026-09-28"), {
    report_type: "monthly",
    start_date: "2026-09-01",
    end_date: "2026-09-30",
  });
  assert.throws(() => resolveWorkshopInventoryReportPeriod("monthly", "2026-02-30"), /valid calendar date/);
});

test("workshop inventory report creates a GM summary and live stock register tabs", async () => {
  const buffer = await buildWorkshopInventoryReportWorkbook(sampleData, {
    reportType: "monthly",
    startDate: "2026-09-01",
    endDate: "2026-09-30",
  });
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), [
    "GM Summary",
    "Parts",
    "Components",
    "G.E.T",
    "Tyres",
    "Oils & Lubricants",
    "Stock Movements",
    "Critical Spares",
  ]);
  assert.equal(workbook.getWorksheet("GM Summary").getCell("A1").value, "IRONLOG Workshop Inventory Report");
  assert.equal(workbook.getWorksheet("Parts").getCell("A4").value, "Stock Code");
  assert.equal(workbook.getWorksheet("Parts").getCell("A5").value, "KIT-500");
  assert.equal(workbook.getWorksheet("Oils & Lubricants").getCell("A5").value, "OIL-10W40");
  assert.equal(workbook.getWorksheet("Stock Movements").getCell("A5").value, "2026-09-27");
  assert.equal(workbook.getWorksheet("Critical Spares").getCell("A5").value, "KIT-500");
  assert.equal(workbook.getWorksheet("Parts").getCell("U5").value, "URGENT - STOCK OUT");
});

test("each stock category gets its own register and the summary totals by category", async () => {
  const data = {
    items: [
      ...sampleData.items,
      { ...sampleData.items[1], part_code: "TIP-J300", part_name: "Tip Cat J300", stock_category: "get", is_lube: false },
      { ...sampleData.items[1], part_code: "TYRE-2325", part_name: "23.5R25 tyre", stock_category: "tyre", is_lube: false },
      { ...sampleData.items[1], part_code: "ENG-3306", part_name: "Cat 3306B engine", stock_category: "component", is_lube: false },
    ],
    movements: [],
  };
  const buffer = await buildWorkshopInventoryReportWorkbook(data, { reportType: "monthly", startDate: "2026-09-01", endDate: "2026-09-30" });
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  assert.equal(workbook.getWorksheet("G.E.T").getCell("A5").value, "TIP-J300");
  assert.equal(workbook.getWorksheet("Tyres").getCell("A5").value, "TYRE-2325");
  assert.equal(workbook.getWorksheet("Components").getCell("A5").value, "ENG-3306");
  assert.equal(workbook.getWorksheet("Components").getCell("C5").value, "Components");
  assert.equal(workbook.getWorksheet("Parts").getCell("A7").value, null, "only the two parts are on the Parts sheet");
  const summary = workbook.getWorksheet("GM Summary");
  const labels = [];
  summary.eachRow((row) => labels.push(row.getCell(1).value));
  assert.ok(labels.includes("Stock by category"));
  for (const name of ["Parts", "Components", "G.E.T", "Tyres", "Oils & Lubricants", "Total"]) assert.ok(labels.includes(name), name);
});
