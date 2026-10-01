import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";
import ExcelJS from "exceljs";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-fuelbench-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { ensurePlantHireSchema } = await import("../utils/plantHire.js");
ensurePlantHireSchema(db); // done by the assets routes at start-up
const { default: dashboardRoutes } = await import("../routes/dashboard.routes.js");
const { default: reportsRoutes } = await import("../routes/reports.routes.js");
// The reports routes start hourly schedulers; they must not keep the test process alive.
const realSetInterval = globalThis.setInterval;
globalThis.setInterval = (...args) => realSetInterval(...args).unref();
const { default: workOrderRoutes } = await import("../routes/workorders.routes.js");
const { default: financeRoutes } = await import("../routes/finance.routes.js");
const app = Fastify({ logger: false });
// Tables and columns the Excel's cost summary reads are created by these at start-up.
await app.register(workOrderRoutes, { prefix: "/api/workorders" });
await app.register(financeRoutes, { prefix: "/api/finance" });
await app.register(dashboardRoutes, { prefix: "/api/dashboard" });
await app.register(reportsRoutes, { prefix: "/api/reports" });
await app.ready();
globalThis.setInterval = realSetInterval;
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

const H = { "x-user-name": "admin", "x-user-role": "admin", "x-user-roles": "admin", "x-site-code": "main" };
const get = async (url) => {
  const res = await app.inject({ method: "GET", url, headers: H });
  return { code: res.statusCode, body: res.headers["content-type"]?.includes("json") ? res.json() : res.rawPayload };
};

// E501: excavator, OEM 35 L/hr, two typos in September (one jump, one backwards).
// E025: excavator whose fills mostly have no meter reading (Q2 showed 978 L/hr).
// D600: dozer with no benchmark set (column default 5 L/hr).
db.exec(`
  INSERT INTO assets (id, asset_code, asset_name, category, active) VALUES
    (1, 'E501AM', 'CAT 350', '50t Excavator', 1),
    (2, 'E025', 'Hitachi 470-Z', 'Excavator', 1),
    (3, 'D600AM', 'Cat D6R Dozer', 'Dozer', 1);
  UPDATE assets SET baseline_fuel_l_per_hour = 35 WHERE id = 1;
  UPDATE assets SET baseline_fuel_l_per_hour = 40 WHERE id = 2;
`);
const fill = db.prepare(`INSERT INTO fuel_logs (asset_id, log_date, liters, source, meter_unit, meter_run_value) VALUES (?, ?, ?, 'test', 'hours', ?)`);
// E501: 10 h/day at 30 L/hr; day 10 typed 15100 instead of 15000 (jump), day 20 typed 1510 (backwards).
let meter = 14900;
fill.run(1, "2026-08-31", 300, meter);
for (let d = 1; d <= 30; d += 1) {
  meter += 10;
  const typed = d === 10 ? meter + 400 : d === 20 ? 1510 : meter;
  fill.run(1, `2026-09-${String(d).padStart(2, "0")}`, 300, typed);
}
// E025: 20 fills of 400 L, readings only every 5th fill, 12 h/day at 33.3 L/hr.
let m2 = 8000;
fill.run(2, "2026-08-31", 400, m2);
for (let d = 1; d <= 20; d += 1) {
  m2 += 12;
  fill.run(2, `2026-09-${String(d).padStart(2, "0")}`, 400, d % 5 === 0 ? m2 : 0);
}
// D600: 8 h/day at 15.7 L/hr.
let m3 = 3000;
fill.run(3, "2026-08-31", 125, m3);
for (let d = 1; d <= 10; d += 1) {
  m3 += 8;
  fill.run(3, `2026-09-${String(d).padStart(2, "0")}`, 125.6, m3);
}

test("benchmark is worked out fill to fill; bad readings no longer distort it", async (t) => {
  t.after(() => app.close());
  const r = await get("/api/dashboard/fuel?start=2026-09-01&end=2026-09-30");
  assert.equal(r.code, 200, String(r.body?.message || r.body?.error || r.body));
  const by = Object.fromEntries(r.body.rows.map((x) => [x.asset_code, x]));

  const e501 = by.E501AM;
  assert.equal(e501.hours_run, 300, "30 days x 10 h; the typos add nothing");
  assert.equal(e501.actual_lph, 30);
  assert.equal(e501.suspect_readings, 2);
  assert.equal(e501.run_source, "fill_meter");
  assert.equal(e501.is_excessive, false);

  const e025 = by.E025;
  assert.equal(e025.hours_run, 240);
  assert.equal(e025.actual_lph, 33.333, "not 978 L/hr: litres without a reading wait for the next one");
  assert.equal(e025.coverage_pct, 100);

  const d6 = by.D600AM;
  assert.equal(d6.actual_lph, 15.7);
  assert.equal(d6.oem_lph, null, "no benchmark set");
  assert.equal(d6.oem_set, false);
  assert.equal(d6.is_excessive, false, "not flagged against a default 5 L/hr");
  assert.match(d6.note, /benchmark not set/);
  assert.equal(r.body.summary.excessive_count, 0);

  // Saving the benchmark (even 5) makes it count.
  const save = await app.inject({ method: "POST", url: "/api/dashboard/fuel/baseline", headers: H, payload: { asset_code: "D600AM", baseline_fuel_l_per_hour: 12 } });
  assert.equal(save.statusCode, 200, save.body);
  const after = (await get("/api/dashboard/fuel?start=2026-09-01&end=2026-09-30")).body.rows.find((x) => x.asset_code === "D600AM");
  assert.equal(after.oem_lph, 12);
  assert.equal(after.is_excessive, true, "15.7 is over 12 + 15%");

  // Fill-by-fill history names the bad readings.
  const daily = (await get("/api/dashboard/fuel/daily?asset_code=E501AM&start=2026-09-01&end=2026-09-30")).body;
  const bad = daily.rows.filter((x) => x.invalid_delta);
  assert.deepEqual(bad.map((x) => x.log_date), ["2026-09-10", "2026-09-20"]);
  assert.match(bad[0].status, /jump of 410 h: reading ignored, litres carried to the 2026-09-11 fill/);
  assert.equal(daily.summary.avg_lph, 30);
  assert.equal(daily.summary.rejected_readings, 2);

  // Excel export and PDFs build with the new columns.
  const x = await get("/api/reports/fuel-benchmark.xlsx?start=2026-09-01&end=2026-09-30");
  assert.equal(x.code, 200);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(x.body);
  const ws = wb.getWorksheet("2026-09");
  const head = ws.getRow(1).values.slice(1);
  assert.ok(head.includes("Matched %") && head.includes("Notes"));
  const rowFor = (code) => { let found = null; ws.eachRow((row) => { if (row.getCell(1).value === code) found = row; }); return found; };
  assert.equal(rowFor("E025").getCell(head.indexOf("Actual L/hr") + 1).value, 33.333);
  for (const url of ["/api/reports/fuel-benchmark.pdf", "/api/reports/fuel-reconciliation.pdf", "/api/reports/fuel-reconciliation.xlsx", "/api/reports/fuel-machine-history.pdf?asset_code=E501AM&x=1"]) {
    const res = await get(`${url}${url.includes("?") ? "&" : "?"}start=2026-09-01&end=2026-09-30`);
    assert.equal(res.code, 200, `${url}: ${String(res.body).slice(0, 300)}`);
  }
});
