import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-kpi-standby-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { ensurePlantHireSchema } = await import("../utils/plantHire.js");
ensurePlantHireSchema(db); // done by the assets routes at start-up
const { default: dashboardRoutes } = await import("../routes/dashboard.routes.js");
const { default: workOrderRoutes } = await import("../routes/workorders.routes.js");
const app = Fastify({ logger: false });
await app.register(workOrderRoutes, { prefix: "/api/workorders" });
await app.register(dashboardRoutes, { prefix: "/api/dashboard" });
await app.ready();
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

const H = { "x-user-name": "admin", "x-user-role": "admin", "x-user-roles": "admin", "x-site-code": "main" };

// Three drill rigs, each with a breakdown left open since 1 October:
// TD01AM works on the 8th, then is marked standby in Daily Input for five days.
// TD02AM is marked standby on the 9th, but 6 h downtime was logged that day.
// TD03AM has no Daily Input at all, so the open breakdown means it was down.
db.exec(`
  INSERT INTO assets (id, asset_code, asset_name, category, active, is_standby) VALUES
    (70, 'TD01AM', 'Atlas Copco ECM 660', 'Drill Rig', 1, 0),
    (71, 'TD02AM', 'Drill 2', 'Drill Rig', 1, 0),
    (72, 'TD03AM', 'Drill 3', 'Drill Rig', 1, 0);
  INSERT INTO breakdowns (id, asset_id, breakdown_date, status, description, critical) VALUES
    (700, 70, '2026-10-01', 'OPEN', 'Hydraulic hose', 0),
    (710, 71, '2026-10-01', 'OPEN', 'Engine', 0),
    (720, 72, '2026-10-01', 'OPEN', 'Track', 0);
  INSERT INTO breakdown_downtime_logs (breakdown_id, log_date, hours_down) VALUES (710, '2026-10-09', 6);
`);
const day = db.prepare(`INSERT INTO daily_hours (asset_id, work_date, scheduled_hours, hours_run, is_used) VALUES (?, ?, ?, ?, ?)`);
day.run(70, "2026-10-08", 11, 5.2, 1);
for (const d of ["2026-10-09", "2026-10-10", "2026-10-11", "2026-10-12", "2026-10-13"]) day.run(70, d, 0, 0, 0);
day.run(71, "2026-10-09", 0, 0, 0);

async function kpi(code) {
  const res = await app.inject({
    method: "GET",
    url: `/api/dashboard/asset-kpi/weekly?start=2026-10-08&end=2026-10-13&scheduled=10&asset_codes=${code}`,
    headers: H,
  });
  assert.equal(res.statusCode, 200);
  const a = res.json().by_asset[0] || {};
  return [a.scheduled_hours, a.run_hours, a.downtime_hours, a.availability_pct];
}

test("standby days are not counted as downtime, even with a breakdown left open", async () => {
  // Only the worked day counts: 11 h scheduled, 5.2 h run, nothing down.
  assert.deepEqual(await kpi("TD01AM"), [11, 5.2, 0, 100]);
});

test("downtime explicitly logged on a standby day still counts", async () => {
  // The 9th: 6 h logged of the 10 h default day. The other five days have no
  // Daily Input and an open breakdown, so they stay imputed as down (5 × 10 h).
  assert.deepEqual(await kpi("TD02AM"), [60, 0, 56, 6.7]);
});

test("a machine with an open breakdown and no Daily Input is still counted as down", async () => {
  assert.deepEqual(await kpi("TD03AM"), [60, 0, 60, 0]);
});
