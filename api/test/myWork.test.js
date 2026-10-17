import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import { buildMyWork, queryLowStock } from "../utils/myWork.js";

function seed() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT);
    CREATE TABLE tasks (id INTEGER PRIMARY KEY, title TEXT, status TEXT, priority TEXT, project TEXT,
      assigned_to TEXT, due_date TEXT, site_code TEXT);
    CREATE TABLE breakdowns (id INTEGER PRIMARY KEY, asset_id INTEGER, breakdown_date TEXT, status TEXT,
      description TEXT, critical INTEGER, parts_status TEXT, site_code TEXT,
      component TEXT, ets_repair_date TEXT, primary_work_order_id INTEGER);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER, status TEXT, opened_at TEXT,
      closed_at TEXT, completed_at TEXT, assigned_artisan_name TEXT, priority TEXT, site_code TEXT);
    CREATE TABLE parts (id INTEGER PRIMARY KEY, part_code TEXT, part_name TEXT, min_stock REAL, critical INTEGER);
    CREATE TABLE stock_movements (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL);
    CREATE TABLE parts_orders (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL, expected_date TEXT, status TEXT);
  `);
  db.exec(`
    INSERT INTO assets VALUES (1, 'EX01', 'Excavator'), (2, 'DT02', 'Dump truck');
    INSERT INTO tasks VALUES
      (1, 'Late task', 'open', 'low', NULL, 'Jaco', '2026-09-01', 'main'),
      (2, 'Today task', 'in_progress', 'high', NULL, 'jaco', '2026-09-29', 'main'),
      (3, 'No due date', 'open', 'high', NULL, 'jaco', NULL, 'main'),
      (4, 'Finished', 'done', 'high', NULL, 'jaco', '2026-09-01', 'main'),
      (5, 'Someone else', 'open', 'high', NULL, 'pieter', '2026-09-01', 'main'),
      (6, 'Other site', 'open', 'high', NULL, 'jaco', '2026-09-01', 'north');
    INSERT INTO breakdowns (id, asset_id, breakdown_date, status, description, critical, parts_status, site_code) VALUES
      (1, 1, '2026-09-20', 'open', 'Hydraulic leak', 0, NULL, 'main'),
      (2, 2, '2026-09-25', 'OPEN', 'Engine', 1, 'ordered', NULL),
      (3, 2, '2026-09-10', 'closed', 'Old', 1, NULL, 'main'),
      (4, 1, '2026-09-10', 'open', 'North', 0, NULL, 'north');
    INSERT INTO work_orders (id, asset_id, source, status, opened_at, closed_at, completed_at, assigned_artisan_name, priority, site_code) VALUES
      (1, 1, 'breakdown', 'open', '2026-09-20', NULL, NULL, NULL, NULL, 'main'),
      (2, 2, 'service', 'In Progress', '2026-09-21', NULL, NULL, 'Sipho', NULL, 'main'),
      (3, 2, 'service', 'closed', '2026-09-01', '2026-09-02', NULL, NULL, NULL, 'main'),
      (4, 2, 'service', 'open', '2026-09-01', NULL, '2026-09-03', NULL, NULL, 'main');
    INSERT INTO parts VALUES (1, 'FLT', 'Filter', 5, 0), (2, 'BELT', 'Belt', 2, 1), (3, 'OK', 'Plenty', 1, 0);
    INSERT INTO stock_movements VALUES (1, 1, 3), (2, 3, 10);
    INSERT INTO parts_orders VALUES (1, 2, 4, '2026-10-01', 'ordered'), (2, 1, 2, NULL, 'received');
  `);
  return db;
}

const ctx = { user: "Jaco", site: "main", today: "2026-09-29" };

test("tasks are the caller's open tasks on their site, due soonest first", () => {
  const w = buildMyWork(seed(), { ...ctx, roles: ["operator"] });
  assert.deepEqual(Object.keys(w.sections), ["tasks"]);
  assert.deepEqual(w.sections.tasks.items.map((t) => t.id), [1, 2, 3]);
  assert.equal(w.sections.tasks.count, 3);
  assert.equal(w.sections.tasks.overdue, 1);
  assert.equal(w.sections.tasks.due_today, 1);
});

test("workshop roles see open breakdowns and work orders for their site", () => {
  const w = buildMyWork(seed(), { ...ctx, roles: ["plant_manager"] });
  assert.equal(w.sections.low_stock, undefined);
  assert.deepEqual(w.sections.open_breakdowns.items.map((b) => b.id), [2, 1]);
  assert.equal(w.sections.open_breakdowns.items[0].critical, true);
  assert.deepEqual(w.sections.open_work_orders.items.map((x) => x.id), [1, 2]);
  assert.equal(w.sections.open_work_orders.unassigned, 1);
});

test("stores roles see low stock and outstanding part orders", () => {
  const w = buildMyWork(seed(), { ...ctx, roles: ["storeman"] });
  assert.equal(w.sections.open_breakdowns, undefined);
  assert.deepEqual(w.sections.low_stock.items.map((p) => p.part_code), ["BELT", "FLT"]);
  assert.deepEqual(w.sections.parts_on_order.items.map((p) => p.part_code), ["BELT"]);
  // Breakdown 2 is waiting on ordered parts with no part listed yet.
  assert.deepEqual(w.sections.waiting_parts.items.map((r) => `${r.kind}:${r.asset_code}`), ["breakdown:DT02"]);
  assert.equal(w.sections.waiting_parts.machines_down, 1);
});

test("workshop admin covers both workshop and stores sections", () => {
  const w = buildMyWork(seed(), { ...ctx, roles: ["workshop_admin"] });
  // Plant admins also approve completed work orders and see who is on shift (home screen per role).
  assert.deepEqual(Object.keys(w.sections).sort(), ["approvals", "low_stock", "open_breakdowns", "open_shifts", "open_work_orders", "parts_on_order", "tasks", "team", "waiting_parts"]);
});

test("previews are capped at five while count stays complete", () => {
  const db = seed();
  const insert = db.prepare("INSERT INTO tasks (title, status, priority, assigned_to, site_code) VALUES (?, 'open', 'low', 'jaco', 'main')");
  for (let i = 0; i < 6; i++) insert.run(`Extra ${i}`);
  const w = buildMyWork(db, { ...ctx, roles: [] });
  assert.equal(w.sections.tasks.count, 9);
  assert.equal(w.sections.tasks.items.length, 5);
});

test("low stock keeps the alerts ordering and shape", () => {
  const rows = queryLowStock(seed(), 100);
  assert.deepEqual(rows.map((r) => [r.part_code, r.on_hand, r.critical]), [["BELT", 0, true], ["FLT", 3, false]]);
});

// Ease of use, part 2: a home screen per role.
function seedRoles() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE users (id INTEGER PRIMARY KEY, username TEXT, full_name TEXT, role TEXT);
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT, category TEXT, site_code TEXT);
    CREATE TABLE tasks (id INTEGER PRIMARY KEY, title TEXT, status TEXT, priority TEXT, project TEXT, assigned_to TEXT, due_date TEXT, site_code TEXT);
    CREATE TABLE breakdowns (id INTEGER PRIMARY KEY, asset_id INTEGER, breakdown_date TEXT, status TEXT, description TEXT, critical INTEGER,
      parts_status TEXT, site_code TEXT, component TEXT, ets_repair_date TEXT, primary_work_order_id INTEGER);
    CREATE TABLE breakdown_downtime_logs (id INTEGER PRIMARY KEY, breakdown_id INTEGER, log_date TEXT, hours_down REAL);
    CREATE TABLE daily_hours (id INTEGER PRIMARY KEY, asset_id INTEGER, work_date TEXT, hours_run REAL);
    CREATE TABLE maintenance_plans (id INTEGER PRIMARY KEY, service_name TEXT, interval_hours REAL);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER, status TEXT, opened_at TEXT,
      closed_at TEXT, completed_at TEXT, assigned_artisan_name TEXT, assigned_at TEXT, started_at TEXT, due_date TEXT, priority TEXT,
      job_description TEXT, repair_progress TEXT, completion_notes TEXT, labor_hours REAL, labor_rate_per_hour REAL, site_code TEXT);
    CREATE TABLE tech_shifts (id INTEGER PRIMARY KEY, username TEXT, site_code TEXT, started_at TEXT, status TEXT);
    CREATE TABLE approval_requests (id INTEGER PRIMARY KEY, module TEXT, action TEXT, status TEXT, requested_by TEXT, created_at TEXT);
    CREATE TABLE parts (id INTEGER PRIMARY KEY, part_code TEXT, part_name TEXT, min_stock REAL, critical INTEGER, unit_cost REAL);
    CREATE TABLE stock_movements (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL, movement_type TEXT, reference TEXT, created_at TEXT);
    CREATE TABLE parts_orders (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL, expected_date TEXT, status TEXT);

    INSERT INTO users VALUES (1, 'sipho', 'Sipho Ndlovu', 'artisan'), (2, 'ana', 'Ana Silva', 'artisan');
    INSERT INTO assets VALUES (1, 'EX01', 'Excavator', 'Excavator', 'main'), (2, 'DT02', 'Dump truck', 'ADT', 'main'), (3, 'N1', 'North', 'ADT', 'north');
    INSERT INTO breakdowns (id, asset_id, breakdown_date, status, description, critical, site_code) VALUES (1, 2, '2026-10-15', 'OPEN', 'Engine', 1, 'main');
    INSERT INTO breakdown_downtime_logs VALUES (1, 1, '2026-10-15', 10), (2, 1, '2026-10-16', 10);
    INSERT INTO daily_hours (asset_id, work_date, hours_run) VALUES (1, '2026-10-10', 8), (1, '2026-10-11', 8), (2, '2026-10-10', 9), (2, '2026-10-11', 9), (3, '2026-10-11', 9);
    INSERT INTO work_orders (id, asset_id, source, reference_id, status, opened_at, completed_at, assigned_artisan_name, labor_hours, labor_rate_per_hour, site_code) VALUES
      (1, 2, 'breakdown', 1, 'in_progress', '2026-10-15', NULL, 'sipho', NULL, NULL, 'main'),
      (2, 1, 'service', NULL, 'assigned', '2026-10-16', NULL, 'ana', NULL, NULL, 'main'),
      (3, 1, 'service', NULL, 'completed', '2026-10-02', '2026-10-05', 'ana', 4, 50, 'main'),
      (4, 3, 'service', NULL, 'completed', '2026-10-02', '2026-10-05', 'x', 100, 50, 'north');
    INSERT INTO tech_shifts (username, site_code, started_at, status) VALUES ('sipho', 'main', '2026-10-17T05:00:00Z', 'open'), ('ana', 'main', '2026-10-17T05:10:00Z', 'open'), ('zed', 'north', '2026-10-17T05:00:00Z', 'open');
    INSERT INTO approval_requests (module, action, status, requested_by, created_at) VALUES ('stock', 'adjust_movement', 'pending', 'st', '2026-10-16'), ('stock', 'adjust_movement', 'approved', 'st', '2026-10-01');
    INSERT INTO parts VALUES (1, 'FLT', 'Filter', 0, 0, 20);
    INSERT INTO stock_movements (part_id, quantity, movement_type, reference, created_at) VALUES (1, -3, 'out', 'work_order:3', '2026-10-05 08:00:00'), (1, -9, 'out', 'work_order:4', '2026-10-05 08:00:00');
  `);
  return db;
}
const roleCtx = { user: "boss", site: "main", today: "2026-10-17" };

test("a manager sees availability, machines down, this month's work order cost and what waits for approval", () => {
  const w = buildMyWork(seedRoles(), { ...roleCtx, roles: ["supervisor"] });
  const f = w.sections.figures;
  assert.deepEqual(f.period, { start: "2026-10-10", end: "2026-10-16" });
  assert.equal(f.availability_pct, 50, "4 machine-days × 10 h = 40 h available, 20 h down; the north machine is left out");
  assert.equal(f.machines_down, 1);
  assert.equal(f.work_order_cost_usd, 260, "4 h × $50 labour + 3 × $20 parts on this site's work order");
  assert.deepEqual([w.sections.approvals.count, w.sections.approvals.requests, w.sections.approvals.work_orders], [2, 1, 1]);
  assert.equal(buildMyWork(seedRoles(), { ...roleCtx, roles: ["executive"] }).sections.approvals, undefined, "executives see the figures but approve nothing here");
});

test("a foreman sees who is on shift and who is free; a technician sees their own jobs", () => {
  const foreman = buildMyWork(seedRoles(), { ...roleCtx, roles: ["plant_manager"] });
  assert.deepEqual(foreman.sections.team.items.map((p) => [p.name, p.busy_on]), [["Ana Silva", null], ["Sipho Ndlovu", "WO 1 — DT02"]], "free first; the north technician is not here");
  assert.equal(foreman.sections.team.free, 1);
  assert.equal(foreman.sections.approvals.requests, 0, "foremen approve completed work orders, not stock requests");
  const tech = buildMyWork(seedRoles(), { ...roleCtx, user: "sipho", roles: ["artisan"] });
  assert.deepEqual(Object.keys(tech.sections).sort(), ["my_jobs", "tasks"]);
  assert.deepEqual(tech.sections.my_jobs.items.map((j) => [j.id, j.asset_code, j.critical]), [[1, "DT02", true]]);
  assert.equal(tech.sections.my_jobs.urgent, 1);
});
