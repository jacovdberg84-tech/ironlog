import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import { buildReliabilityIncidentsForAssets } from "../utils/reliabilityMetrics.js";

function hasTable(db, name) {
  return Boolean(db.prepare("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?").get(name));
}

function hasColumn(db, table, column) {
  return db.prepare(`PRAGMA table_info(${table})`).all().some((row) => row.name === column);
}

function buildFixture() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE breakdowns (
      id INTEGER PRIMARY KEY,
      asset_id INTEGER NOT NULL,
      breakdown_date TEXT,
      description TEXT,
      primary_work_order_id INTEGER,
      downtime_total_hours REAL DEFAULT 0
    );
    CREATE TABLE breakdown_downtime_logs (
      id INTEGER PRIMARY KEY,
      breakdown_id INTEGER NOT NULL,
      log_date TEXT NOT NULL,
      hours_down REAL NOT NULL
    );
    CREATE TABLE work_orders (
      id INTEGER PRIMARY KEY,
      asset_id INTEGER NOT NULL,
      reference_id INTEGER,
      source TEXT,
      status TEXT,
      opened_at TEXT,
      closed_at TEXT,
      completed_at TEXT,
      labor_hours REAL
    );
  `);
  return db;
}

test("reliability counts only downtime recorded in the selected period", () => {
  const db = buildFixture();
  db.prepare("INSERT INTO work_orders VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)").run(
    11, 1, 101, "breakdown", "closed", "2026-07-17 06:00:00", "2026-08-11 11:30:00", null, 0
  );
  // This late-entered record points to an old work order. It must not add August
  // downtime simply because the linked work order was open in August.
  db.prepare("INSERT INTO breakdowns VALUES (?, ?, ?, ?, ?, ?)").run(
    101, 1, "2026-09-09", "Planned maintenance", 11, 1
  );
  // A July incident with a real August log is valid August downtime.
  db.prepare("INSERT INTO breakdowns VALUES (?, ?, ?, ?, ?, ?)").run(
    102, 1, "2026-07-30", "Hydraulic repair", null, 0
  );
  db.prepare("INSERT INTO breakdown_downtime_logs VALUES (?, ?, ?, ?)").run(1, 102, "2026-08-04", 2.5);
  // This legacy incident has no daily log, so its header is allowed because it
  // was reported inside the selected period.
  db.prepare("INSERT INTO breakdowns VALUES (?, ?, ?, ?, ?, ?)").run(
    103, 1, "2026-08-12", "Tyre repair", null, 1.25
  );
  // A current-period incident with no recorded downtime is not a failure yet.
  db.prepare("INSERT INTO breakdowns VALUES (?, ?, ?, ?, ?, ?)").run(
    104, 1, "2026-08-20", "Awaiting inspection", null, 0
  );

  const result = buildReliabilityIncidentsForAssets(db, {
    assetIds: [1],
    start: "2026-08-01",
    end: "2026-08-31",
    hasTable: (name) => hasTable(db, name),
    hasColumn: (table, column) => hasColumn(db, table, column),
  });

  assert.deepEqual(
    result.incidents.map((row) => [row.breakdown_id, row.downtime_hours, row.downtime_source]),
    [[102, 2.5, "downtime_logs"], [103, 1.25, "breakdown_header"]]
  );
  assert.deepEqual(result.byAsset.get(1), { failure_count: 2, downtime_hours: 3.75 });
  db.close();
});
