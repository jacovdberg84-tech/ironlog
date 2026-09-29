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
      description TEXT, critical INTEGER, parts_status TEXT, site_code TEXT);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, status TEXT, opened_at TEXT,
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
    INSERT INTO breakdowns VALUES
      (1, 1, '2026-09-20', 'open', 'Hydraulic leak', 0, NULL, 'main'),
      (2, 2, '2026-09-25', 'OPEN', 'Engine', 1, 'ordered', NULL),
      (3, 2, '2026-09-10', 'closed', 'Old', 1, NULL, 'main'),
      (4, 1, '2026-09-10', 'open', 'North', 0, NULL, 'north');
    INSERT INTO work_orders VALUES
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
});

test("workshop admin covers both workshop and stores sections", () => {
  const w = buildMyWork(seed(), { ...ctx, roles: ["workshop_admin"] });
  assert.deepEqual(Object.keys(w.sections).sort(), ["low_stock", "open_breakdowns", "open_work_orders", "parts_on_order", "tasks"]);
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
