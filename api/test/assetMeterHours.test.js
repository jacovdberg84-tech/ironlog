import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import { getAssetCurrentHoursInfo, getAssetHoursInfoAsOf } from "../utils/assetMeterHours.js";

function seedMeterDb() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE asset_hours (asset_id INTEGER PRIMARY KEY, total_hours REAL);
    CREATE TABLE daily_hours (
      id INTEGER PRIMARY KEY,
      asset_id INTEGER,
      work_date TEXT,
      closing_hours REAL,
      hours_run REAL,
      is_used INTEGER
    );
    INSERT INTO asset_hours (asset_id, total_hours) VALUES (1, 2887);
    INSERT INTO daily_hours (asset_id, work_date, closing_hours, hours_run, is_used) VALUES
      (1, '2026-09-01', 2875, 8, 1),
      (1, '2026-09-02', 2887, 12, 1),
      -- Imported production history can have a very large accumulated run total.
      (1, '2026-09-03', NULL, 35498, 1);
  `);
  return db;
}

test("current meter uses the latest valid closing rather than summed production history", () => {
  const db = seedMeterDb();
  assert.deepEqual(getAssetCurrentHoursInfo(1, db), {
    hours: 2887,
    source: "daily_closing",
    latest_work_date: "2026-09-02",
  });
  db.close();
});

test("as-of meter does not turn imported daily hours into a current SMR", () => {
  const db = seedMeterDb();
  assert.deepEqual(getAssetHoursInfoAsOf(1, "2026-09-03", db), {
    hours: 2887,
    source: "daily_closing",
    latest_work_date: "2026-09-02",
  });
  db.close();
});
