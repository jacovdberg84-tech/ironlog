import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import { rememberPlanHours, reopenWorkOrder } from "../utils/workOrderReopen.js";

function seed() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER, status TEXT,
      assigned_artisan_name TEXT, closed_at TEXT, completed_at TEXT);
    CREATE TABLE maintenance_plans (id INTEGER PRIMARY KEY, asset_id INTEGER, service_name TEXT, interval_hours REAL,
      last_service_hours REAL, active INTEGER DEFAULT 1);
    CREATE TABLE breakdowns (id INTEGER PRIMARY KEY, asset_id INTEGER, status TEXT, end_at TEXT);
    INSERT INTO maintenance_plans VALUES (1, 30, '250', 250, 4750, 1), (2, 30, '500', 500, 4500, 1), (3, 31, '250', 250, 100, 1);
  `);
  return db;
}

const plan = (db, id) => db.prepare("SELECT last_service_hours AS h FROM maintenance_plans WHERE id = ?").get(id).h;

test("a service work order closed by mistake reopens with the plans put back exactly", () => {
  const db = seed();
  db.exec(`INSERT INTO work_orders (id, asset_id, source, reference_id, status) VALUES (1, 30, 'service', 1, 'in_progress')`);
  // What close does: remember, then set every active plan on the machine to the meter.
  rememberPlanHours(db, 1, 30);
  db.exec(`UPDATE maintenance_plans SET last_service_hours = 5010 WHERE asset_id = 30; UPDATE work_orders SET status = 'closed', closed_at = datetime('now') WHERE id = 1;`);
  const r = reopenWorkOrder(db, 1);
  assert.equal(r.status, "open", "nobody was assigned");
  assert.equal(plan(db, 1), 4750);
  assert.equal(plan(db, 2), 4500);
  assert.equal(plan(db, 3), 100, "other machines untouched");
  assert.ok(r.plans_restored.every((p) => p.exact));
  const wo = db.prepare("SELECT status, closed_at, completed_at FROM work_orders WHERE id = 1").get();
  assert.deepEqual(wo, { status: "open", closed_at: null, completed_at: null });
  assert.throws(() => reopenWorkOrder(db, 1), /Only a completed, approved or closed/);
});

test("closed before plan hours were kept: plans still on the close hours go back one interval", () => {
  const db = seed();
  db.exec(`
    UPDATE maintenance_plans SET last_service_hours = 5000 WHERE id IN (1, 2);
    INSERT INTO work_orders (id, asset_id, source, reference_id, status, assigned_artisan_name) VALUES (2, 30, 'service', 1, 'closed', 'jose');
  `);
  const r = reopenWorkOrder(db, 2, { rolledHours: 5000 });
  assert.equal(r.status, "assigned", "keeps the technician, ready to reassign or start");
  assert.equal(plan(db, 1), 4750, "250 h service due again at 5000");
  assert.equal(plan(db, 2), 4500);
  assert.ok(r.plans_restored.every((p) => !p.exact));
});

test("a breakdown repair reopened reopens its breakdown", () => {
  const db = seed();
  db.exec(`
    INSERT INTO breakdowns VALUES (9, 30, 'CLOSED', '2026-09-30 10:00:00');
    INSERT INTO work_orders (id, asset_id, source, reference_id, status) VALUES (3, 30, 'breakdown', 9, 'completed');
  `);
  const r = reopenWorkOrder(db, 3);
  assert.equal(r.breakdown_reopened, true);
  assert.deepEqual(db.prepare("SELECT status, end_at FROM breakdowns WHERE id = 9").get(), { status: "OPEN", end_at: null });
});
