import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-fuelcost-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { ensurePlantHireSchema } = await import("../utils/plantHire.js");
ensurePlantHireSchema(db); // done by the assets routes at start-up
const { default: dashboardRoutes } = await import("../routes/dashboard.routes.js");
// The reports routes start hourly schedulers; they must not keep the test process alive.
const realSetInterval = globalThis.setInterval;
globalThis.setInterval = (...args) => realSetInterval(...args).unref();
const { default: reportsRoutes } = await import("../routes/reports.routes.js");
const app = Fastify({ logger: false });
// Fuel price columns (fuel_logs.unit_cost_per_liter, assets.fuel_cost_per_liter) are added by the dashboard routes.
await app.register(dashboardRoutes, { prefix: "/api/dashboard" });
await app.register(reportsRoutes, { prefix: "/api/reports" });
await app.ready();
globalThis.setInterval = realSetInterval;
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

const H = { "x-user-name": "admin", "x-user-role": "admin", "x-user-roles": "admin", "x-site-code": "main" };
const get = (url) => app.inject({ method: "GET", url, headers: H });

test("fuel cost PDF opens inline for a month and downloads for a date range", async () => {
  const assetId = db.prepare(`INSERT INTO assets (asset_code, asset_name, category) VALUES ('FC01', 'Fuel test truck', 'Truck')`).run().lastInsertRowid;
  db.prepare(`INSERT INTO fuel_logs (asset_id, log_date, liters) VALUES (?, '2026-09-12', 120)`).run(assetId);

  const month = await get("/api/reports/fuel-cost.pdf?month=2026-09");
  assert.equal(month.statusCode, 200);
  assert.equal(month.headers["content-type"], "application/pdf");
  assert.match(month.headers["content-disposition"], /^inline; filename="AML_Fuel_Cost_2026-09\.pdf"/);
  assert.equal(month.rawPayload.subarray(0, 4).toString(), "%PDF");

  const range = await get("/api/reports/fuel-cost.pdf?start=2026-09-01&end=2026-09-15&download=1");
  assert.equal(range.statusCode, 200);
  assert.match(range.headers["content-disposition"], /^attachment; filename="AML_Fuel_Cost_2026-09-01_to_2026-09-15\.pdf"/);
});

test("fuel cost PDF asks for a valid period", async () => {
  assert.equal((await get("/api/reports/fuel-cost.pdf")).statusCode, 400);
  assert.equal((await get("/api/reports/fuel-cost.pdf?start=2026-09-20&end=2026-09-01")).statusCode, 400);
});
