import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-wo-move-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { default: workOrderRoutes } = await import("../routes/workorders.routes.js");

const app = Fastify({ logger: false });
await app.register(workOrderRoutes, { prefix: "/api/workorders" });
await app.ready();

process.on("exit", () => {
  try { db.close(); } catch { /* already closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

db.exec(`
  INSERT INTO assets (id, asset_code, asset_name, category, active) VALUES
    (1, 'E500AM', 'CAT 350', 'Excavator', 1),
    (2, 'E504AM', 'CAT 350', 'Excavator', 1);
  INSERT INTO breakdowns (id, asset_id, breakdown_date, status, description)
    VALUES (90, 1, '2026-10-07', 'OPEN', 'Hydraulic hose failure');
  INSERT INTO work_orders (id, asset_id, source, reference_id, status, site_code)
    VALUES (349, 1, 'breakdown', 90, 'in_progress', 'main');
  UPDATE breakdowns SET primary_work_order_id = 349 WHERE id = 90;
  INSERT INTO breakdown_offsite_repairs (asset_id, breakdown_id, repair_status, sent_date)
    VALUES (1, 90, 'sent_offsite', '2026-10-07');
`);

const as = (role) => ({
  "x-user-name": role === "admin" ? "bj.vandenberg" : "operator",
  "x-user-role": role,
  "x-user-roles": role,
  "x-site-code": "main",
});

test("admin correction moves a breakdown, linked work orders and off-site record together", async (t) => {
  t.after(() => app.close());

  const denied = await app.inject({
    method: "POST",
    url: "/api/workorders/349/move-asset",
    headers: as("supervisor"),
    payload: { asset_code: "E504AM" },
  });
  assert.equal(denied.statusCode, 403);

  const moved = await app.inject({
    method: "POST",
    url: "/api/workorders/349/move-asset",
    headers: as("admin"),
    payload: { asset_code: "E504AM" },
  });
  assert.equal(moved.statusCode, 200, moved.body);
  assert.equal(moved.json().to_asset.asset_code, "E504AM");

  assert.equal(db.prepare("SELECT asset_id FROM breakdowns WHERE id = 90").get().asset_id, 2);
  assert.equal(db.prepare("SELECT asset_id FROM work_orders WHERE id = 349").get().asset_id, 2);
  assert.equal(db.prepare("SELECT asset_id FROM breakdown_offsite_repairs WHERE breakdown_id = 90").get().asset_id, 2);
});
