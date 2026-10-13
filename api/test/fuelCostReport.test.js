import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import vm from "node:vm";
import Database from "better-sqlite3";

// pdfGenerator opens the app database on import; point it at a throwaway file.
const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-fuel-cost-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
const { buildFuelCostPdf, buildFuelCostReport } = await import("../utils/fuelCostReport.js");
const { buildAssetPeriodCosts, readCostDefaults } = await import("../utils/monthlyOperatingCosts.js");

test.after(() => rmSync(tempDir, { recursive: true, force: true }));

function seed() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT, category TEXT, fuel_cost_per_liter REAL, downtime_cost_per_hour REAL);
    CREATE TABLE fuel_logs (id INTEGER PRIMARY KEY, asset_id INTEGER, log_date TEXT, liters REAL, unit_cost_per_liter REAL);
    CREATE TABLE cost_settings (key TEXT PRIMARY KEY, value TEXT);
    INSERT INTO cost_settings VALUES ('fuel_cost_per_liter_default', '1.40');
    INSERT INTO assets VALUES
      (1, 'E500AM', 'CAT 350 Excavator', 'Excavator', NULL, NULL),
      (2, 'T01AM', 'AXOR Tipper', 'Truck', 1.25, NULL),
      (3, 'T02AM', 'AXOR Tipper 2', 'Truck', NULL, NULL),
      (4, 'G01AM', 'Grader', 'Grader', NULL, NULL);
    INSERT INTO fuel_logs (asset_id, log_date, liters, unit_cost_per_liter) VALUES
      (1, '2026-09-01', 400, 1.50),  -- log price wins
      (1, '2026-09-15', 100, NULL),  -- falls back to default 1.40
      (2, '2026-09-10', 200, NULL),  -- falls back to the asset rate 1.25
      (3, '2026-09-30', 50, NULL),
      (3, '2026-10-01', 999, NULL),  -- outside September
      (4, '2026-08-31', 300, NULL);  -- outside September
  `);
  return db;
}

test("fuel cost report matches the fuel figures of the Monthly Cost Excel", () => {
  const db = seed();
  const period = { start: "2026-09-01", end: "2026-09-30" };
  const defaults = readCostDefaults(db);
  const report = buildFuelCostReport(db, period, defaults);
  const costMonthly = buildAssetPeriodCosts(db, period.start, period.end, defaults);

  const fromExcel = Object.fromEntries(costMonthly.filter((r) => r.fuel_cost > 0).map((r) => [r.asset_code, r.fuel_cost]));
  const fromPdf = Object.fromEntries(report.rows.map((r) => [r.asset_code, r.fuel_cost]));
  assert.deepEqual(fromPdf, fromExcel);

  assert.deepEqual(report.rows.map((r) => r.asset_code), ["E500AM", "T01AM", "T02AM"], "largest cost first, out-of-period machines left out");
  assert.equal(report.rows[0].fuel_cost, 740);
  assert.equal(report.rows[0].liters, 500);
  assert.equal(report.rows[0].fills, 2);
  assert.equal(report.rows[0].cost_per_liter, 1.48);
  assert.equal(report.rows[1].fuel_cost, 250);
  assert.equal(report.rows[2].fuel_cost, 70);

  assert.deepEqual(report.categories.map((c) => [c.category, c.assets, c.fuel_cost]), [["Excavator", 1, 740], ["Truck", 2, 320]]);
  assert.equal(report.totals.fuel_cost, 1060);
  assert.equal(report.totals.liters, 750);
  assert.equal(report.totals.assets, 3);
  assert.equal(report.totals.default_cost_per_liter, 1.4);
});

test("fuel cost report handles a period with no fuel", async () => {
  const db = seed();
  const report = buildFuelCostReport(db, { start: "2025-01-01", end: "2025-01-31" }, readCostDefaults(db));
  assert.equal(report.rows.length, 0);
  assert.equal(report.totals.fuel_cost, 0);
  assert.equal(report.totals.cost_per_liter, null);
  const pdf = await buildFuelCostPdf(report, { db });
  assert.equal(pdf.subarray(0, 4).toString(), "%PDF");
});

test("fuel cost PDF renders for a month with fuel", async () => {
  const db = seed();
  const report = buildFuelCostReport(db, { start: "2026-09-01", end: "2026-09-30" }, readCostDefaults(db));
  const pdf = await buildFuelCostPdf(report, { db });
  assert.ok(Buffer.isBuffer(pdf));
  assert.equal(pdf.subarray(0, 4).toString(), "%PDF");
  assert.ok(pdf.length > 1000);
});

test("Reports fuel cost card builds month or date-range links", () => {
  const source = readFileSync(new URL("../../web/app/reports.js", import.meta.url), "utf8");
  const helper = source.slice(source.indexOf("let fuelCostPdfBlobUrl"), source.indexOf("async function loadFuelCostPdfPreview("))
    + source.slice(source.indexOf("function downloadFuelCostPdf("), source.indexOf("async function printFuelCostPdf("));
  const values = {};
  const calls = { alerts: [], downloads: [] };
  const context = vm.createContext({
    API: "",
    qs: (id) => ({ value: values[id] || "" }),
    alert: (msg) => calls.alerts.push(msg),
    downloadAuthedFile: (url, name) => calls.downloads.push([url, name]),
  });
  vm.runInContext(helper, context);

  values.fuelCostMonth = "2026-09";
  values.fuelCostStart = "2026-09-05";
  values.fuelCostEnd = "2026-09-10";
  assert.equal(context.fuelCostPdfQuery(), "month=2026-09", "month wins when both are filled");

  values.fuelCostMonth = "";
  assert.equal(context.fuelCostPdfQuery(), "start=2026-09-05&end=2026-09-10");
  context.downloadFuelCostPdf();
  assert.deepEqual(calls.downloads, [["/api/reports/fuel-cost.pdf?start=2026-09-05&end=2026-09-10&download=1", "fuel-cost.pdf"]]);

  values.fuelCostStart = "2026-09-20";
  assert.equal(context.fuelCostPdfQuery(), null);
  values.fuelCostEnd = "";
  assert.equal(context.fuelCostPdfQuery(), null);
  assert.equal(calls.alerts.length, 2);
});
