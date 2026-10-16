import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

// /api/kpi and the finance SSOT report read downtime from breakdown_downtime_logs and
// hours from daily_hours.work_date (they used columns the database does not have).
const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-kpi-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;

await import("../db/migrate.js");
(await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
const { db } = await import("../db/client.js");
const realSetInterval = globalThis.setInterval;
globalThis.setInterval = (...args) => realSetInterval(...args).unref();
const app = Fastify({ logger: false });
for (const [file, prefix] of [["stock.routes.js", "/api/stock"], ["dashboard.routes.js", "/api/dashboard"], ["breakdowns.routes.js", "/api/breakdowns"], ["kpi.routes.js", "/api/kpi"], ["finance.routes.js", "/api/finance"]]) {
  await app.register((await import(`../routes/${file}`)).default, { prefix });
}
await app.ready();
globalThis.setInterval = realSetInterval;
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

const H = { "x-user-name": "boss", "x-user-role": "admin", "x-user-roles": "admin", "x-site-code": "main" };
const get = (url) => app.inject({ method: "GET", url, headers: H });

const truck = db.prepare(`INSERT INTO assets (asset_code, asset_name, category, active, is_standby, site_code) VALUES ('T1', 'Truck 1', 'Truck', 1, 0, 'main')`).run().lastInsertRowid;
const dozer = db.prepare(`INSERT INTO assets (asset_code, asset_name, category, active, is_standby, site_code) VALUES ('D1', 'Dozer 1', 'Dozer', 1, 0, 'main')`).run().lastInsertRowid;
for (const id of [truck, dozer]) db.prepare(`INSERT INTO daily_hours (asset_id, work_date, hours_run, is_used, scheduled_hours) VALUES (?, '2026-10-14', 8, 1, 10)`).run(id);
const bd = db.prepare(`INSERT INTO breakdowns (asset_id, breakdown_date, status, description) VALUES (?, '2026-10-14', 'OPEN', 'Hydraulic hose')`).run(dozer).lastInsertRowid;
db.prepare(`INSERT INTO breakdown_downtime_logs (breakdown_id, log_date, hours_down) VALUES (?, '2026-10-14', 2)`).run(bd);

test("daily KPIs: availability from logged downtime, biggest losses listed", async () => {
  const res = await get("/api/kpi/daily/2026-10-14?scheduled=10");
  assert.equal(res.statusCode, 200, res.body);
  const k = res.json();
  assert.deepEqual([k.used_assets_count, k.available_hours, k.run_hours, k.downtime_hours, k.availability, k.utilization], [2, 20, 16, 2, 90, 80]);
  assert.deepEqual(k.major_downtime.map((m) => [m.asset_code, m.downtime_hours]), [["D1", 2]]);
});

test("range KPIs over several days", async () => {
  const res = await get("/api/kpi/range?start=2026-10-13&end=2026-10-15&scheduled=10");
  assert.equal(res.statusCode, 200, res.body);
  const k = res.json();
  assert.deepEqual([k.totals.available_hours, k.totals.run_hours, k.totals.downtime_hours, k.totals.availability], [20, 16, 2, 90]);
  assert.equal(k.major_downtime[0].asset_code, "D1");
});

test("finance SSOT report and its exports open for a month", async () => {
  const res = await get("/api/finance/reports/ssot?period=2026-10");
  assert.equal(res.statusCode, 200, res.body);
  assert.equal((await get("/api/finance/reports/ssot/export.csv?period=2026-10")).statusCode, 200);
  assert.equal((await get("/api/finance/reports/ssot/export.xlsx?period=2026-10")).statusCode, 200);
});
