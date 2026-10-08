import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-june-schedule-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
const Fastify = (await import("fastify")).default;
const juneRoutes = (await import("../routes/june.routes.js")).default;
const { db } = await import("../db/client.js");

const admin = { "x-user-role": "admin", "x-user-roles": "admin", "x-user-name": "jaco", "x-site-code": "main" };

test("June schedule draft gives only its owner an authenticated Excel download", async (t) => {
  t.after(async () => {
    db.close();
    rmSync(tempDir, { recursive: true, force: true });
  });
  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT, category TEXT);
    CREATE TABLE asset_hours (asset_id INTEGER, total_hours REAL);
    CREATE TABLE daily_hours (id INTEGER PRIMARY KEY, asset_id INTEGER, work_date TEXT, closing_hours REAL, hours_run REAL, is_used INTEGER);
    CREATE TABLE maintenance_plans (id INTEGER PRIMARY KEY, asset_id INTEGER, service_name TEXT, interval_hours REAL, last_service_hours REAL, active INTEGER);
  `);
  db.prepare("INSERT INTO assets (id, asset_code, asset_name, category) VALUES (1, 'G01AM', 'CAT 140H', 'Grader')").run();
  db.prepare("INSERT INTO daily_hours (asset_id, work_date, closing_hours, hours_run, is_used) VALUES (1, '2026-10-07', 21478.9, 9, 1)").run();
  db.prepare("INSERT INTO maintenance_plans (id, asset_id, service_name, interval_hours, last_service_hours, active) VALUES (1, 1, '1000 hour service', 1000, 21000, 1)").run();

  const app = Fastify();
  await app.register(juneRoutes, { prefix: "/api/june" });
  t.after(() => app.close());

  const draft = await app.inject({
    method: "POST",
    url: "/api/june/gateway/execute",
    headers: admin,
    payload: { name: "june_draft_maintenance_schedule", arguments: { asset_codes: ["G01AM"], horizon_days: 30, as_of: "2026-10-08" } },
  });
  assert.equal(draft.statusCode, 200);
  const result = draft.json().result;
  assert.equal(result.state, "review_only");
  assert.equal(result.type, "maintenance_schedule");
  assert.match(result.download.report_id, /^[0-9a-f-]{36}$/i);

  const report = await app.inject({
    method: "GET",
    url: `/api/june/maintenance-schedule/${result.download.report_id}.xlsx`,
    headers: admin,
  });
  assert.equal(report.statusCode, 200);
  assert.match(String(report.headers["content-disposition"]), /IRONLOG_June_Maintenance_Schedule_2026-10-08\.xlsx/);
  assert.match(String(report.headers["content-type"]), /application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet/);
  assert.ok(report.rawPayload.length > 1000);

  const anotherAdmin = await app.inject({
    method: "GET",
    url: `/api/june/maintenance-schedule/${result.download.report_id}.xlsx`,
    headers: { ...admin, "x-user-name": "another-admin" },
  });
  assert.equal(anotherAdmin.statusCode, 404);
});

