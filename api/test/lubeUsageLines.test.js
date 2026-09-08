import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import { fetchLubeUsageLines } from "../utils/lubeUsageLines.js";

function createLubeDb() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT);
    CREATE TABLE parts (
      id INTEGER PRIMARY KEY,
      part_code TEXT,
      part_name TEXT,
      unit_cost REAL,
      consumable_kind TEXT
    );
    CREATE TABLE oil_logs (
      id INTEGER PRIMARY KEY,
      asset_id INTEGER,
      log_date TEXT,
      oil_type TEXT,
      quantity REAL,
      created_at TEXT,
      unit_cost REAL,
      part_id INTEGER
    );
    CREATE TABLE stock_movements (
      id INTEGER PRIMARY KEY,
      part_id INTEGER,
      quantity REAL,
      movement_type TEXT,
      reference TEXT,
      created_at TEXT
    );
  `);
  db.prepare("INSERT INTO assets (id, asset_code, asset_name) VALUES (1, 'G01AM', 'Cat 140H')").run();
  db.prepare(`
    INSERT INTO parts (id, part_code, part_name, unit_cost, consumable_kind)
    VALUES (?, ?, ?, ?, ?)
  `).run(10, "MLFPT2301", "Fuchs Titan UTTO TO-4 SAE 50 210lt", 5.63, "lube");
  return db;
}

test("lube usage export resolves historical generic oil logs to their issued stock part", () => {
  const db = createLubeDb();
  try {
    const createdAt = "2026-09-08 07:48:31";
    db.prepare(`
      INSERT INTO oil_logs (id, asset_id, log_date, oil_type, quantity, created_at, unit_cost)
      VALUES (1, 1, '2026-08-21', 'Top up', 15, ?, 5.63)
    `).run(createdAt);
    db.prepare(`
      INSERT INTO stock_movements (id, part_id, quantity, movement_type, reference, created_at)
      VALUES (1, 10, -15, 'out', 'lube_issue:asset:1', ?)
    `).run(createdAt);

    const [line] = fetchLubeUsageLines(db, { start: "2026-08-01", end: "2026-08-31" });
    assert.equal(line.part_code, "MLFPT2301");
    assert.equal(line.part_name, "Fuchs Titan UTTO TO-4 SAE 50 210lt");
    assert.equal(line.lube_type, "lube");
  } finally {
    db.close();
  }
});

test("lube usage export prefers the directly recorded stock part for new issues", () => {
  const db = createLubeDb();
  try {
    db.prepare(`
      INSERT INTO oil_logs (id, asset_id, log_date, oil_type, quantity, created_at, unit_cost, part_id)
      VALUES (1, 1, '2026-09-08', 'Service', 10, '2026-09-08 08:00:00', 5.63, 10)
    `).run();

    const [line] = fetchLubeUsageLines(db, { start: "2026-09-01", end: "2026-09-30" });
    assert.equal(line.part_code, "MLFPT2301");
    assert.equal(line.part_name, "Fuchs Titan UTTO TO-4 SAE 50 210lt");
  } finally {
    db.close();
  }
});
