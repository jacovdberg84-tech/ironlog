import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import {
  buildPartEvidence,
  buildServiceEvidence,
  modelTokens,
  proposalForPart,
  proposalFromAi,
  proposalFromHistory,
  serviceProposalMessages,
} from "../utils/costingAssist.js";

function seed() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT, category TEXT);
    CREATE TABLE maintenance_plans (id INTEGER PRIMARY KEY, asset_id INTEGER, service_name TEXT, interval_hours REAL);
    CREATE TABLE parts (id INTEGER PRIMARY KEY, part_code TEXT, part_name TEXT, unit_cost REAL, stock_category TEXT);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER, status TEXT);
    CREATE TABLE stock_movements (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL, reference TEXT, created_at TEXT, unit_cost_usd REAL);
    CREATE TABLE cost_settings (key TEXT PRIMARY KEY, value TEXT);
    CREATE TABLE stores_part_orders (id INTEGER PRIMARY KEY, part_id INTEGER, part_code TEXT, unit_cost REAL, order_date TEXT);
    INSERT INTO cost_settings VALUES ('labor_cost_per_hour_default', '40');
    INSERT INTO assets VALUES (1, 'W200AM', 'BELL B25D 23000L WATER TRUCK', 'Water Truck'), (2, 'W201AM', 'BELL B25D 23000L WATER TRUCK', 'Water Truck');
    INSERT INTO maintenance_plans VALUES (10, 1, '500 h service', 500), (20, 2, '500 h service', 500);
    INSERT INTO parts VALUES
      (1, 'B25D-OF', 'Oil filter B25D', 18.5, 'part'),
      (2, 'OIL15', 'Engine oil 15W40', 3, 'oil'),
      (3, 'B25D-FF', 'Fuel filter B25D', 0, 'part'),
      (4, 'BOLT', 'Bolt', 0.2, 'part');
    INSERT INTO work_orders VALUES (100, 2, 'service', 20, 'closed'), (101, 2, 'service', 20, 'closed');
    INSERT INTO stock_movements VALUES
      (1, 1, -1, 'work_order:100', '2026-06-01', NULL), (2, 2, -28, 'work_order:100', '2026-06-01', NULL),
      (3, 1, -1, 'work_order:101', '2026-08-01', NULL), (4, 2, -30, 'work_order:101', '2026-08-01', NULL),
      (5, 4, -2, 'work_order:101', '2026-08-01', NULL),
      (6, 3, 10, NULL, '2026-07-01', 22.4);
    INSERT INTO stores_part_orders VALUES (1, 3, 'B25D-FF', 21.9, '2026-09-10');
  `);
  return db;
}

test("model tokens skip generic words", () => {
  assert.deepEqual(modelTokens("BELL B25D 23000L WATER TRUCK"), ["BELL", "B25D", "23000L"]);
});

test("service evidence uses sister machines when this one has no history", () => {
  const ev = buildServiceEvidence(seed(), 10);
  assert.equal(ev.own_history.services, 0);
  assert.deepEqual(ev.peers.assets, ["W201AM"]);
  assert.equal(ev.peers.history.services, 2);
  assert.ok(ev.store_candidates.some((c) => c.part_code === "B25D-FF"));
  assert.equal(ev.labour.standard_hours, 4);
  assert.equal(ev.labour.rate, 40);
  const msg = JSON.parse(serviceProposalMessages(ev)[1].content);
  assert.equal(msg.service.asset_code, "W200AM");
  assert.equal(msg.store_candidates[0].unit_cost, undefined, "prices are not sent to the model");
});

test("history proposal keeps parts used on at least half the services", () => {
  const db = seed();
  const p = proposalFromHistory(db, buildServiceEvidence(db, 10));
  assert.equal(p.mode, "history");
  // The bolt was on 1 of 2 services: half counts, as in the template builder.
  assert.deepEqual(p.lines.map((l) => [l.part_code, l.qty, l.type]), [["B25D-OF", 1, "part"], ["OIL15", 30, "oil"], ["BOLT", 2, "part"]]);
  assert.deepEqual(p.labour, { hours: 4, rate: 40, total: 160 });
  assert.equal(p.totals.total, 18.5 + 90 + 0.4 + 160);
});

test("AI proposals are checked and priced from the store", () => {
  const db = seed();
  const ev = buildServiceEvidence(db, 10);
  const reply = 'Sure! {"summary":"Standard B25D 500 h kit","items":[{"part_code":"b25d-of","qty":1,"why":"sister machine"},{"part_code":"B25D-FF","qty":2},{"part_code":"MADE-UP","qty":1},{"part_code":"OIL15","qty":-3}],"labour_hours":99,"questions":["Axle oil?"]}';
  const p = proposalFromAi(db, ev, reply);
  assert.equal(p.mode, "borris");
  assert.deepEqual(p.lines.map((l) => [l.part_code, l.qty, l.unit_cost]), [["B25D-OF", 1, 18.5], ["B25D-FF", 2, 0]]);
  assert.deepEqual(p.rejected_codes, ["MADE-UP"]);
  assert.deepEqual(p.unpriced_codes, ["B25D-FF"]);
  assert.equal(p.labour.hours, 4, "implausible labour falls back to the standard");
  assert.deepEqual(p.questions, ["Axle oil?"]);
  assert.equal(proposalFromAi(db, ev, "no json here"), null);
});

test("part price is proposed only from a real purchase", () => {
  const db = seed();
  const p = proposalForPart(buildPartEvidence(db, "B25D-FF"));
  assert.equal(p.unit_cost, 21.9);
  assert.equal(p.purchases.length, 2);
  db.exec("DELETE FROM stores_part_orders; DELETE FROM stock_movements WHERE unit_cost_usd > 0;");
  const none = proposalForPart(buildPartEvidence(db, "B25D-FF"));
  assert.equal(none.unit_cost, null);
  assert.match(none.summary, /No purchase price/);
});
