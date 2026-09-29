import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import { summarizeWaiting, workshopWaitingOnParts } from "../utils/partsWaiting.js";

function seed() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT);
    CREATE TABLE breakdowns (id INTEGER PRIMARY KEY, asset_id INTEGER, status TEXT, component TEXT, description TEXT,
      parts_status TEXT, critical INTEGER, breakdown_date TEXT, ets_repair_date TEXT, primary_work_order_id INTEGER, site_code TEXT);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER, status TEXT);
    CREATE TABLE parts (id INTEGER PRIMARY KEY, part_code TEXT, part_name TEXT);
    CREATE TABLE stock_movements (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL);
    CREATE TABLE maintenance_parts_requests (id INTEGER PRIMARY KEY, site_code TEXT, asset_id INTEGER, asset_code TEXT,
      part_code TEXT, part_name TEXT, qty REAL, urgency TEXT, notes TEXT, work_order_id INTEGER, status TEXT,
      requested_by TEXT, status_notes TEXT, created_at TEXT);
    INSERT INTO assets VALUES (1, 'EX01', 'Excavator'), (2, 'DT02', 'Dump truck'), (3, 'LD03', 'Loader');
    INSERT INTO parts VALUES (1, 'SEAL-1', 'Seal kit'), (2, 'PUMP-9', 'Hydraulic pump');
    INSERT INTO stock_movements VALUES (1, 1, 5), (2, 1, -1), (3, 2, 0);
    INSERT INTO breakdowns VALUES
      (1, 1, 'OPEN', 'Hydraulics', 'Boom leak', 'Ordered', 1, '2026-09-20', NULL, 10, 'main'),
      (2, 2, 'OPEN', 'Engine', 'Overheating', 'Waiting OEM', 0, '2026-09-22', '2026-10-05', 11, 'main'),
      (3, 3, 'CLOSED', 'Tyres', 'Flat', 'Ordered', 0, '2026-09-01', NULL, NULL, 'main');
    INSERT INTO work_orders VALUES (10, 1, 'breakdown', 1, 'in_progress'), (11, 2, 'breakdown', 2, 'open'),
      (12, 3, 'service', NULL, 'open'), (13, 3, 'service', NULL, 'closed');
    INSERT INTO maintenance_parts_requests VALUES
      (1, 'main', 1, 'EX01', 'SEAL-1', 'Seal kit', 2, 'normal', NULL, 10, 'requested', 'sipho', NULL, '2026-09-21'),
      (2, 'main', 3, 'LD03', 'PUMP-9', 'Hydraulic pump', 1, 'urgent', NULL, 12, 'ordered', 'johan', 'PO 55', '2026-09-10'),
      (3, 'main', 3, 'LD03', NULL, 'Filter', 1, 'normal', NULL, 13, 'requested', 'johan', NULL, '2026-09-01'),
      (4, 'main', 3, 'LD03', NULL, 'Belt', 1, 'normal', NULL, 12, 'received', 'johan', NULL, '2026-09-02'),
      (5, 'north', 1, 'EX01', NULL, 'Other site', 1, 'normal', NULL, NULL, 'requested', 'x', NULL, '2026-09-02');
  `);
  return db;
}

test("lists open requests and uncovered breakdowns, machines down first", () => {
  const rows = workshopWaitingOnParts(seed(), { site: "main" });
  assert.deepEqual(rows.map((r) => `${r.kind}:${r.id}`), ["request:1", "breakdown:2", "request:2"]);
  const seal = rows[0];
  assert.equal(seal.machine_down, true);
  assert.equal(seal.on_hand, 4);
  assert.equal(seal.in_stock, true, "4 on hand covers qty 2");
  assert.equal(rows[2].in_stock, false);
  assert.equal(rows[2].notes, "PO 55");
});

test("a breakdown whose work order already has a request is not listed twice", () => {
  const rows = workshopWaitingOnParts(seed(), { site: "main" });
  assert.equal(rows.some((r) => r.kind === "breakdown" && r.id === 1), false);
});

test("requests on finished work orders, received requests and other sites are left out", () => {
  const ids = workshopWaitingOnParts(seed(), { site: "main" }).filter((r) => r.kind === "request").map((r) => r.id);
  assert.deepEqual(ids.sort(), [1, 2]);
});

test("summary counts machines down, issuable and to-order requests", () => {
  const rows = workshopWaitingOnParts(seed(), { site: "main" });
  assert.deepEqual(summarizeWaiting(rows), { total: 3, machines_down: 2, in_stock: 1, not_ordered: 0 });
});

test("works before the parts request table exists", () => {
  const db = new Database(":memory:");
  db.exec(`CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT);`);
  assert.deepEqual(workshopWaitingOnParts(db), []);
});
