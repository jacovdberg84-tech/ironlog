import assert from "node:assert/strict";
import test from "node:test";
import Database from "better-sqlite3";
import ExcelJS from "exceljs";
import { buildMonthlyBudgetSpend } from "../utils/monthlyBudgetSpend.js";
import { buildMonthlyBudgetSpendWorkbook } from "../utils/monthlyBudgetSpendWorkbook.js";

function setupDb() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE assets (
      id INTEGER PRIMARY KEY,
      asset_code TEXT,
      asset_name TEXT,
      site_code TEXT
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
    INSERT INTO assets VALUES (1, 'A300AM', 'Bell B30D', 'main');
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

test("monthly budget spend keeps the requested cost sources traceable and avoids linked job-card double capture", async () => {
  const db = setupDb();
  try {
    const report = buildMonthlyBudgetSpend(db, { month: "2026-10", siteCode: "main" });
    const categories = Object.fromEntries(report.categories.map((row) => [row.key, row]));

    assert.equal(report.budget.amount, 1000);
    assert.equal(categories.mechanics.amount, 160);
    assert.equal(categories.lubes.amount, 30);
    assert.equal(categories.parts.amount, 100);
    assert.equal(categories.maintenance.amount, 160);
    assert.equal(categories.maintenance.records, 1, "WO #17 is already captured by its mechanics job card");
    assert.equal(report.total_spend, 450);
    assert.equal(report.variance, -550);
    assert.equal(report.budget_used_pct, 45);
    assert.deepEqual(new Set(report.details.map((row) => row.category)), new Set([
      "Mechanics cost",
      "Lubes used",
      "Parts used",
      "Maintenance service work",
    ]));

    const workbook = await buildMonthlyBudgetSpendWorkbook(report);
    const bytes = await workbook.xlsx.writeBuffer();
    const loaded = new ExcelJS.Workbook();
    await loaded.xlsx.load(bytes);
    assert.deepEqual(loaded.worksheets.map((sheet) => sheet.name), ["Summary", "Spend Detail"]);
    assert.ok(loaded.getWorksheet("Summary").getCell("A1").value.includes("Budget vs Spend"));
  } finally {
    db.close();
  }
});
