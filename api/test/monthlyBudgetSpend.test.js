import assert from "node:assert/strict";
import test from "node:test";
import Database from "better-sqlite3";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { buildMonthlyBudgetSpend } from "../utils/monthlyBudgetSpend.js";
import { buildMonthlyBudgetSpendBuffer, buildMonthlyBudgetSpendWorkbook } from "../utils/monthlyBudgetSpendWorkbook.js";
import { createDepartmentalCostingInput } from "../utils/departmentalCostingInputs.js";

function setupDb() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE assets (
      id INTEGER PRIMARY KEY,
      asset_code TEXT,
      asset_name TEXT,
      site_code TEXT,
      cost_center_code TEXT,
      category TEXT,
      active INTEGER,
      archived INTEGER,
      utilization_mode TEXT,
      km_per_hour_factor REAL,
      hire_billing_mode TEXT,
      hire_rate_per_hour REAL,
      hire_fixed_monthly REAL
    );
    CREATE TABLE daily_hours (
      id INTEGER PRIMARY KEY,
      asset_id INTEGER,
      work_date TEXT,
      hours_run REAL,
      is_used INTEGER
    );
    CREATE TABLE cost_settings (key TEXT PRIMARY KEY, value REAL);
    CREATE TABLE finance_budgets_monthly (
      id INTEGER PRIMARY KEY,
      period TEXT,
      site_code TEXT,
      cost_center_code TEXT,
      equipment_type TEXT,
      category TEXT,
      budget_amount REAL,
      currency TEXT,
      notes TEXT,
      created_by TEXT,
      updated_at TEXT
    );
    CREATE TABLE mechanic_labor_entries (
      id INTEGER PRIMARY KEY,
      work_date TEXT,
      technician_name TEXT,
      hours REAL,
      asset_code TEXT,
      reason TEXT,
      labor_rate_per_hour REAL,
      site_code TEXT,
      category TEXT,
      job_card_no TEXT
    );
    CREATE TABLE parts (
      id INTEGER PRIMARY KEY,
      part_code TEXT,
      part_name TEXT,
      unit_cost REAL,
      stock_category TEXT,
      stock_category_source TEXT
    );
    CREATE TABLE oil_logs (
      id INTEGER PRIMARY KEY,
      asset_id INTEGER,
      log_date TEXT,
      oil_type TEXT,
      quantity REAL,
      unit_cost REAL,
      created_at TEXT
    );
    CREATE TABLE work_orders (
      id INTEGER PRIMARY KEY,
      asset_id INTEGER,
      source TEXT,
      status TEXT,
      opened_at TEXT,
      closed_at TEXT,
      completed_at TEXT,
      labor_hours REAL,
      labor_rate_per_hour REAL,
      site_code TEXT,
      job_description TEXT
    );
    CREATE TABLE stock_movements (
      id INTEGER PRIMARY KEY,
      part_id INTEGER,
      quantity REAL,
      movement_type TEXT,
      reference TEXT,
      created_at TEXT,
      unit_cost REAL,
      total_cost REAL
    );
  `);
  db.exec(`
    INSERT INTO assets VALUES (1, 'A300AM', 'Bell B30D', 'main', 'AML-3001', 'Contractor Hire', 1, 0, 'hours', 10, 'hourly', 20, 0);
    INSERT INTO daily_hours VALUES (1, 1, '2026-10-08', 10, 1);
    INSERT INTO cost_settings VALUES ('labor_cost_per_hour_default', 80), ('lube_cost_per_qty_default', 4);
    INSERT INTO finance_budgets_monthly
      (period, site_code, cost_center_code, equipment_type, category, budget_amount, currency)
      VALUES ('2026-10', 'main', '', '', 'operating', 1000, 'USD');
    INSERT INTO mechanic_labor_entries
      VALUES (1, '2026-10-08', 'Arnold', 2, 'A300AM', '500 hour service', 80, 'main', 'Service', '17');
    INSERT INTO parts VALUES
      (1, 'MLF-10W40', 'Fuchs Titan 10W40 Engine Oil', 6, 'oil', 'manual'),
      (2, 'FLT-01', 'Oil filter', 50, 'part', 'manual');
    INSERT INTO oil_logs VALUES (1, 1, '2026-10-08', 'MLF-10W40', 5, 6, '2026-10-08T06:00:00');
    INSERT INTO work_orders VALUES
      (17, 1, 'service', 'completed', '2026-10-08', '2026-10-08', '2026-10-08', 2, 80, 'main', '500 hour service'),
      (18, 1, 'service', 'completed', '2026-10-09', '2026-10-09', '2026-10-09', 2, 80, 'main', '250 hour service');
    INSERT INTO stock_movements VALUES
      (1, 2, -2, 'out', 'work_order:18', '2026-10-09T08:00:00', 50, 100);
  `);
  return db;
}

test("departmental costing reconciles Plant and Workshop without double-charging direct parts", async () => {
  const db = setupDb();
  try {
    createDepartmentalCostingInput(db, { period: "2026-10", siteCode: "main", entryType: "plant_revenue", assetCode: "A300AM", description: "Approved hire uplift", amount: 100, createdBy: "test" });
    createDepartmentalCostingInput(db, { period: "2026-10", siteCode: "main", entryType: "plant_other_operating", assetCode: "A300AM", description: "Plant operating allocation", amount: 10, createdBy: "test" });
    createDepartmentalCostingInput(db, { period: "2026-10", siteCode: "main", entryType: "plant_depreciation", assetCode: "A300AM", description: "Monthly depreciation", amount: 25, createdBy: "test" });
    createDepartmentalCostingInput(db, { period: "2026-10", siteCode: "main", entryType: "workshop_duties", description: "Import duties", amount: 5, createdBy: "test" });
    createDepartmentalCostingInput(db, { period: "2026-10", siteCode: "main", entryType: "workshop_other_operating", description: "Workshop utilities", amount: 15, createdBy: "test" });
    const report = buildMonthlyBudgetSpend(db, { month: "2026-10", siteCode: "main" });
    const categories = Object.fromEntries(report.categories.map((row) => [row.key, row]));

    assert.equal(report.budget.amount, 1000);
    assert.equal(categories.mechanics.amount, 160);
    assert.equal(categories.lubes.amount, 30);
    assert.equal(categories.parts.amount, 100);
    assert.equal(categories.maintenance.amount, 160);
    assert.equal(categories.maintenance.records, 1, "WO #17 is already captured by its mechanics job card");
    assert.equal(categories.other.amount, 10);
    assert.equal(report.total_spend, 460);
    assert.equal(report.variance, -540);
    assert.equal(report.budget_used_pct, 46);
    assert.deepEqual(new Set(report.details.map((row) => row.category)), new Set([
      "Mechanics cost",
      "Lubes used",
      "Parts used",
      "Maintenance service work", "Plant other operating expense", "Plant depreciation",
      "Workshop import / export duties", "Workshop other operating expense", "External plant hire revenue",
    ]));

    assert.equal(report.departmental.plant.revenue, 300, "hourly hire plus manual external revenue");
    assert.equal(report.departmental.plant.operating_cost, 460);
    assert.equal(report.departmental.plant.depreciation, 25);
    assert.equal(report.departmental.plant.total_cost, 485);
    assert.equal(report.departmental.plant.net_contribution, -185);
    assert.equal(report.departmental.workshop.labour_cost, 320);
    assert.equal(report.departmental.workshop.lube_cost, 30);
    assert.equal(report.departmental.workshop.total_operating_cost, 370);
    assert.equal(report.departmental.workshop.recovery, 350);
    assert.equal(report.departmental.workshop.net_cost_recovery_balance, -20);
    assert.equal(report.departmental.consolidated_result, -205);
    assert.equal(report.departmental.plant.parts, 100, "parts remain direct to Plant and are not Workshop recovery");
    assert.ok(report.departmental.reconciliation.every((row) => row.status === "Matched" && row.difference === 0));

    const { workbook } = await buildMonthlyBudgetSpendWorkbook(report);
    const bytes = await workbook.xlsx.writeBuffer();
    const loaded = new ExcelJS.Workbook();
    await loaded.xlsx.load(bytes);
    assert.deepEqual(loaded.worksheets.map((sheet) => sheet.name), [
      "Management Summary", "Plant Costing", "Workshop Costing", "Dept Reconciliation", "Source Detail",
    ]);
    assert.equal(loaded.getWorksheet("Management Summary").getCell("A1").value, "IRONLOG Departmental Costing Report");
    assert.equal(loaded.getWorksheet("Plant Costing").getCell("C5").value.formula.startsWith("SUMIFS("), true);
    assert.equal(loaded.getWorksheet("Workshop Costing").getCell("D9").value.result, -20);

    const exported = await buildMonthlyBudgetSpendBuffer(report);
    const zip = await JSZip.loadAsync(exported);
    const chartParts = Object.keys(zip.files).filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name));
    assert.equal(chartParts.length, 1, "the executive workbook includes an editable plant contribution chart");
    assert.match(await zip.file(chartParts[0]).async("string"), /Plant net contribution by equipment/);
  } finally {
    db.close();
  }
});
