import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-svcgen-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { default: workOrderRoutes } = await import("../routes/workorders.routes.js");
const { default: maintenanceRoutes } = await import("../routes/maintenance.routes.js");
const app = Fastify({ logger: false });
await app.register(workOrderRoutes, { prefix: "/api/workorders" });
await app.register(maintenanceRoutes, { prefix: "/api/maintenance" });
await app.ready();
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

const H = { "x-user-name": "admin", "x-user-role": "admin", "x-user-roles": "admin", "x-site-code": "main" };
const generate = async (planIds) => {
  const res = await app.inject({ method: "POST", url: "/api/maintenance/generate", headers: H, payload: { plan_ids: planIds } });
  assert.equal(res.statusCode, 200, res.body);
  return res.json();
};

db.exec(`
  INSERT INTO assets (id, asset_code, asset_name, category, active) VALUES
    (1, 'A303AM', 'Bell B30E', '30t ADT', 1), (2, 'A304AM', 'Bell B30E', '30t ADT', 1), (3, 'A305AM', 'Bell B30E', '30t ADT', 1);
  INSERT INTO maintenance_plans (id, asset_id, service_name, interval_hours, last_service_hours, active) VALUES
    (10, 1, '250', 250, 4750, 1), (20, 2, '250', 250, 100, 1), (30, 3, '250', 250, 100, 0);
  INSERT INTO daily_hours (asset_id, work_date, hours_run, is_used) VALUES (1, '2026-09-30', 4900, 1), (2, '2026-09-30', 300, 1), (3, '2026-09-30', 300, 1);
`);

test("a ticked service gets a work order, or a reason when it cannot", async (t) => {
  t.after(() => app.close());

  // A303AM on standby: it is listed in the plans table but was skipped silently before.
  db.exec(`UPDATE assets SET is_standby = 1 WHERE id = 1`);
  const one = await generate([10]);
  assert.equal(one.created_count, 1, JSON.stringify(one));
  assert.equal(one.created[0].asset_code, "A303AM");
  assert.equal(one.created[0].note, "machine is on standby");
  const woId = one.created[0].work_order_id;
  assert.deepEqual(db.prepare("SELECT source, reference_id, status FROM work_orders WHERE id = ?").get(woId), { source: "service", reference_id: 10, status: "open" });

  // Ticked again: says which work order is already open.
  const again = await generate([10]);
  assert.equal(again.created_count, 0);
  assert.equal(again.skipped[0].reason, "open_work_order_exists");
  assert.equal(again.skipped[0].work_order_id, woId);

  // Plan switched off: reason plus the machine's last service work order.
  db.exec(`INSERT INTO work_orders (asset_id, source, reference_id, status, closed_at) VALUES (3, 'service', 30, 'closed', '2026-09-20 10:00:00')`);
  const off = await generate([30]);
  assert.equal(off.created_count, 0);
  assert.equal(off.skipped[0].reason_text, "this service plan is switched off");
  assert.equal(off.skipped[0].last_service_work_order.status, "closed");

  // A normal due machine still works, and the bulk run (nothing ticked) is unchanged.
  const normal = await generate([20]);
  assert.equal(normal.created_count, 1);
});
