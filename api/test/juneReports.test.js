import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-june-reports-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
const { db } = await import("../db/client.js");
const r = await import("../utils/juneReports.js");
const { registerAssetKpiRangeBuilder } = await import("../utils/assetKpiRangeProvider.js");
process.on("exit", () => { try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

db.exec(`
  CREATE TABLE IF NOT EXISTS assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT, category TEXT, downtime_cost_per_hour REAL);
  CREATE TABLE IF NOT EXISTS breakdowns (id INTEGER PRIMARY KEY, asset_id INTEGER, breakdown_date TEXT, critical INTEGER, status TEXT);
  CREATE TABLE IF NOT EXISTS breakdown_downtime_logs (id INTEGER PRIMARY KEY, breakdown_id INTEGER, log_date TEXT, hours_down REAL);
  CREATE TABLE IF NOT EXISTS finance_budgets_monthly (period TEXT, site_code TEXT, cost_center_code TEXT, equipment_type TEXT, category TEXT, budget_amount REAL, currency TEXT);
`);
db.prepare(`INSERT INTO assets (id, asset_code, asset_name, category, downtime_cost_per_hour) VALUES (1, 'A303AM', 'ADT 303', 'ADT', 100), (2, 'E504AM', 'Excavator 504', 'Excavator', 200)`).run();
db.prepare(`INSERT INTO breakdowns (id, asset_id, breakdown_date, critical, status) VALUES (1, 1, '2026-09-03', 1, 'CLOSED'), (2, 1, '2026-09-20', 0, 'CLOSED'), (3, 2, '2026-08-10', 0, 'CLOSED')`).run();
db.prepare(`INSERT INTO breakdown_downtime_logs (breakdown_id, log_date, hours_down) VALUES (1, '2026-09-03', 10), (2, '2026-09-20', 5), (3, '2026-08-10', 4)`).run();
db.prepare(`INSERT INTO finance_budgets_monthly (period, category, budget_amount) VALUES ('2026-09', 'downtime', 1000), ('2026-09', 'parts', 5000)`).run();

test("periods: a month, this month so far, and a custom range", () => {
  const sep = r.resolvePeriod({ month: "2026-09" }, "2026-10-08");
  assert.deepEqual([sep.start, sep.end, sep.previous.start, sep.previous.end], ["2026-09-01", "2026-09-30", "2026-08-01", "2026-08-31"]);
  const mtd = r.resolvePeriod({}, "2026-10-08");
  assert.equal(mtd.month_to_date, true);
  assert.deepEqual([mtd.start, mtd.end, mtd.previous.start, mtd.previous.end], ["2026-10-01", "2026-10-08", "2026-09-01", "2026-09-08"]);
  // 31 March so far compares with all of February, never past its end.
  assert.equal(r.resolvePeriod({}, "2026-03-31").previous.end, "2026-02-28");
  const range = r.resolvePeriod({ from_date: "2026-09-10", to_date: "2026-09-16" });
  assert.deepEqual([range.previous.start, range.previous.end], ["2026-09-03", "2026-09-09"]);
  assert.match(r.resolvePeriod({ from_date: "2024-01-01", to_date: "2026-01-01" }).error, /370 days/);
});

test("cost report uses the Cost Monthly numbers and compares with the previous month", () => {
  const rep = r.buildJuneCostReport({ month: "2026-09" }, { dbConn: db, today: "2026-10-08" });
  assert.equal(rep.currency, "USD");
  // Downtime: 15 h × $100 for A303AM in September; 4 h × $200 for E504AM in August.
  assert.deepEqual(rep.total, { now: 1500, before: 800, change: 700, change_pct: 87.5 });
  assert.deepEqual(rep.by_cost_type.map((c) => [c.category, c.now, c.before]), [["downtime", 1500, 800]]);
  assert.equal(rep.top_machines[0].asset_code, "A303AM");
  assert.equal(rep.top_machines[0].downtime_hours, 15);
  assert.equal(rep.biggest_increases[0].asset_code, "A303AM");
  assert.deepEqual(rep.budget, { total: 6000, by_category: [{ category: "downtime", amount: 1000 }, { category: "parts", amount: 5000 }] });
  const one = r.buildJuneCostReport({ month: "2026-09", asset_code: "e504am" }, { dbConn: db, today: "2026-10-08" });
  assert.deepEqual([one.total.now, one.total.before, one.budget], [0, 800, null]);
});

test("KPI report slims the Asset KPI engine output and counts breakdowns", () => {
  const calls = [];
  registerAssetKpiRangeBuilder((start, end, sched, site, codes) => {
    calls.push([start, end, codes]);
    return {
      fleet: { availability_pct: 82.5, utilization_pct: 61, scheduled_hours: 400, run_hours: 244, downtime_hours: 70 },
      by_category: [{ category: "ADT", asset_count: 2, availability_pct: 82.5, utilization_pct: 61, downtime_hours: 70 }],
      by_asset: [
        { asset_code: "A303AM", asset_name: "ADT 303", category: "ADT", availability_pct: 70, utilization_pct: 50, scheduled_hours: 200, run_hours: 100, downtime_hours: 60, daily_points: [1, 2] },
        { asset_code: "E504AM", asset_name: "Excavator 504", category: "Excavator", availability_pct: 95, utilization_pct: 72, scheduled_hours: 200, run_hours: 144, downtime_hours: 10 },
      ],
    };
  });
  const rep = r.buildJuneKpiReport({ month: "2026-09" }, { dbConn: db, today: "2026-10-08" });
  assert.deepEqual(calls.map((c) => c.slice(0, 2)), [["2026-09-01", "2026-09-30"], ["2026-08-01", "2026-08-31"]]);
  assert.equal(rep.fleet.availability_pct, 82.5);
  assert.equal(rep.lowest_availability[0].asset_code, "A303AM");
  assert.equal(rep.lowest_availability[0].daily_points, undefined, "no bulky daily series for June");
  assert.deepEqual(rep.breakdowns, { total: 2, critical: 1, repeat_machines: [{ asset_code: "A303AM", asset_name: "ADT 303", breakdowns: 2 }] });
  assert.equal(rep.breakdowns_previous.total, 1);
  registerAssetKpiRangeBuilder(null);
  assert.match(r.buildJuneKpiReport({}, { dbConn: db }).error, /KPI engine/);
});

test("report files only link to existing report downloads", () => {
  const cost = r.prepareJuneReportFile({ report: "cost_monthly_xlsx" }, { today: "2026-10-08" });
  assert.deepEqual([cost.url, cost.download, cost.period], ["/api/reports/cost-monthly.xlsx?month=2026-09", true, "2026-09"]);
  const weekly = r.prepareJuneReportFile({ report: "weekly_pdf" }, { today: "2026-10-08" });
  assert.equal(weekly.url, "/api/reports/weekly.pdf?start=2026-09-28&end=2026-10-04");
  const range = r.prepareJuneReportFile({ report: "maintenance_cost_by_equipment_pdf", from_date: "2026-09-01", to_date: "2026-09-15" });
  assert.equal(range.url, "/api/reports/maintenance-cost-by-equipment.pdf?start=2026-09-01&end=2026-09-15");
  const hist = r.prepareJuneReportFile({ report: "asset_history_pdf", asset_code: "a303am" });
  assert.equal(hist.url, "/api/reports/asset-history/A303AM.pdf");
  assert.match(r.prepareJuneReportFile({ report: "asset_history_pdf" }).error, /which machine/);
  assert.match(r.prepareJuneReportFile({ report: "../../etc/passwd" }).error, /Unknown report/);
});
