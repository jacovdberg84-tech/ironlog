import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import { dismissCostingGap, findCostingGaps, isPlanningQuestion, summarizeCostingGaps } from "../utils/costingGaps.js";

function seed() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT);
    CREATE TABLE parts (id INTEGER PRIMARY KEY, part_code TEXT, part_name TEXT, unit_cost REAL);
    CREATE TABLE stock_movements (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL, reference TEXT, created_at TEXT, unit_cost_usd REAL);
    CREATE TABLE service_templates (id INTEGER PRIMARY KEY, name TEXT, active INTEGER);
    CREATE TABLE service_template_items (id INTEGER PRIMARY KEY, service_template_id INTEGER, stock_item_id INTEGER);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, status TEXT, labor_hours REAL, opened_at TEXT, completed_at TEXT, closed_at TEXT);
    CREATE TABLE cost_settings (key TEXT PRIMARY KEY, value TEXT);
    INSERT INTO assets VALUES (1, 'E500AM', 'CAT 350'), (2, 'T01AM', 'AXOR');
    INSERT INTO parts VALUES (1, 'FLT-1', 'Oil filter', 0), (2, 'OIL-15', 'Engine oil 15W40', 3.2), (3, 'BELT', 'Fan belt', 0), (4, 'OLD', 'Old part', 0);
    INSERT INTO service_templates VALUES (1, 'CAT 350 500h', 1), (2, 'Retired', 0);
    INSERT INTO service_template_items VALUES (1, 1, 1), (2, 1, 2), (3, 2, 4);
    INSERT INTO stock_movements VALUES
      (1, 3, -2, 'work_order:9', '2026-09-20 10:00', NULL),
      (2, 4, -1, 'work_order:9', '2026-05-01 10:00', NULL);
    INSERT INTO work_orders VALUES
      (9, 1, 'service', 'closed', 0, '2026-09-19', '2026-09-20', '2026-09-21'),
      (10, 1, 'service', 'closed', 6, '2026-09-19', '2026-09-20', '2026-09-21'),
      (11, 2, 'service', 'in_progress', 0, '2026-09-25', NULL, NULL),
      (12, 2, 'breakdown', 'closed', 0, '2026-09-25', NULL, '2026-09-26');
  `);
  return db;
}

const forecasts = [
  { plan_id: 5, asset_code: "T01AM", service_name: "10000 km service", status: "OVERDUE", remaining_hours: -120, needs_manual_input: true, forecast: { cost_source: "none" } },
  { plan_id: 6, asset_code: "E500AM", service_name: "500 h", status: "ALMOST DUE", remaining_hours: 40, needs_manual_input: false, forecast: { cost_source: "service_template" } },
];

test("finds each kind of gap, most important first", () => {
  const gaps = findCostingGaps(seed(), { forecasts, today: "2026-09-29" });
  assert.deepEqual(gaps.map((g) => g.key), [
    "labour_rate_missing",
    "service_unpriced:5",
    "part_zero_cost:FLT-1",
    "part_zero_cost:BELT",
    "service_no_labour:9",
  ]);
  const svc = gaps.find((g) => g.key === "service_unpriced:5");
  assert.equal(svc.severity, "high");
  assert.match(svc.detail, /Overdue by 120 h/);
  assert.match(gaps.find((g) => g.key === "part_zero_cost:FLT-1").detail, /CAT 350 500h/);
  assert.match(gaps.find((g) => g.key === "part_zero_cost:BELT").detail, /Issued 2 in the last 90 days/);
});

test("a set labour rate and dismissed gaps drop off the list", () => {
  const db = seed();
  db.prepare("INSERT INTO cost_settings VALUES ('labor_cost_per_hour_default', '42')").run();
  dismissCostingGap(db, "part_zero_cost:BELT", { reason: "consumable", user: "jaco" });
  const gaps = findCostingGaps(db, { forecasts, today: "2026-09-29" });
  assert.equal(gaps.some((g) => g.key === "labour_rate_missing"), false);
  assert.equal(gaps.some((g) => g.key === "part_zero_cost:BELT"), false);
  assert.deepEqual(summarizeCostingGaps(gaps), { total: 3, services_unpriced: 1, parts_zero_cost: 1, services_no_labour: 1, labour_rate_missing: false });
});

test("planning questions are recognised", () => {
  assert.equal(isPlanningQuestion("What will services cost next month?"), true);
  assert.equal(isPlanningQuestion("Which costing gaps should I fill first"), true);
  assert.equal(isPlanningQuestion("good morning"), false);
});
