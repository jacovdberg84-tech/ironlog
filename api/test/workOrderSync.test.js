import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import {
  closeBreakdownIfWorkFinished,
  closeWorkOrdersForClosedBreakdown,
  repairBreakdownWorkOrderLinks,
} from "../utils/workOrderSync.js";

function seed() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE breakdowns (id INTEGER PRIMARY KEY, asset_id INTEGER, status TEXT, end_at TEXT);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER,
      status TEXT, completed_at TEXT, closed_at TEXT, completion_notes TEXT);
  `);
  return db;
}

const bd = (db, id) => db.prepare("SELECT status, end_at FROM breakdowns WHERE id = ?").get(id);
const wo = (db, id) => db.prepare("SELECT status, closed_at, completion_notes FROM work_orders WHERE id = ?").get(id);

test("a finished breakdown work order closes the breakdown with the job's end time", () => {
  const db = seed();
  db.exec(`
    INSERT INTO breakdowns VALUES (1, 1, 'OPEN', NULL);
    INSERT INTO work_orders VALUES (1, 1, 'breakdown', 1, 'approved', '2026-09-20 10:00:00', NULL, NULL);
  `);
  assert.equal(closeBreakdownIfWorkFinished(db, 1), true);
  assert.deepEqual(bd(db, 1), { status: "CLOSED", end_at: "2026-09-20 10:00:00" });
  assert.equal(closeBreakdownIfWorkFinished(db, 1), false, "already closed");
});

test("a breakdown stays open while another of its work orders is unfinished", () => {
  const db = seed();
  db.exec(`
    INSERT INTO breakdowns VALUES (1, 1, 'open', NULL);
    INSERT INTO work_orders VALUES
      (1, 1, 'breakdown', 1, 'closed', NULL, '2026-09-20', NULL),
      (2, 1, 'breakdown', 1, 'In Progress', NULL, NULL, NULL),
      (3, 1, 'service', 1, 'closed', NULL, '2026-09-20', NULL);
  `);
  assert.equal(closeBreakdownIfWorkFinished(db, 1), false);
  assert.equal(bd(db, 1).status, "open");
});

test("a breakdown without work orders is left alone", () => {
  const db = seed();
  db.exec(`INSERT INTO breakdowns VALUES (1, 1, 'OPEN', NULL);`);
  assert.equal(closeBreakdownIfWorkFinished(db, 1), false);
});

test("closing a breakdown closes its open work orders only", () => {
  const db = seed();
  db.exec(`
    INSERT INTO breakdowns VALUES (1, 1, 'CLOSED', '2026-09-22 08:00:00');
    INSERT INTO work_orders VALUES
      (1, 1, 'breakdown', 1, 'open', NULL, NULL, NULL),
      (2, 1, 'breakdown', 1, 'completed', '2026-09-21', NULL, NULL),
      (3, 1, 'service', 1, 'open', NULL, NULL, NULL);
  `);
  assert.deepEqual(closeWorkOrdersForClosedBreakdown(db, 1), [1]);
  assert.deepEqual(wo(db, 1), { status: "closed", closed_at: "2026-09-22 08:00:00", completion_notes: "Closed when the breakdown was closed" });
  assert.equal(wo(db, 2).status, "completed");
  assert.equal(wo(db, 3).status, "open");
});

test("repair fixes drift in both directions and is idempotent", () => {
  const db = seed();
  db.exec(`
    INSERT INTO breakdowns VALUES (1, 1, 'OPEN', NULL), (2, 2, 'CLOSED', '2026-09-01'), (3, 3, 'OPEN', NULL);
    INSERT INTO work_orders VALUES
      (1, 1, 'breakdown', 1, 'closed', NULL, '2026-09-05', NULL),
      (2, 2, 'breakdown', 2, 'assigned', NULL, NULL, 'Keep this note'),
      (3, 3, 'breakdown', 3, 'open', NULL, NULL, NULL);
  `);
  assert.deepEqual(repairBreakdownWorkOrderLinks(db), { breakdowns_closed: [1], work_orders_closed: [2] });
  assert.equal(bd(db, 1).status, "CLOSED");
  assert.deepEqual(wo(db, 2), { status: "closed", closed_at: "2026-09-01", completion_notes: "Keep this note" });
  assert.equal(bd(db, 3).status, "OPEN", "genuinely open breakdown untouched");
  assert.deepEqual(repairBreakdownWorkOrderLinks(db), { breakdowns_closed: [], work_orders_closed: [] });
});
