import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import ExcelJS from "exceljs";
import { buildJuneMaintenanceSchedule, buildJuneMaintenanceScheduleWorkbook } from "../utils/juneMaintenanceSchedule.js";

function makeDb() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT, category TEXT);
    CREATE TABLE asset_hours (asset_id INTEGER, total_hours REAL);
    CREATE TABLE daily_hours (id INTEGER PRIMARY KEY, asset_id INTEGER, work_date TEXT, closing_hours REAL, hours_run REAL, is_used INTEGER);
    CREATE TABLE maintenance_plans (id INTEGER PRIMARY KEY, asset_id INTEGER, service_name TEXT, interval_hours REAL, last_service_hours REAL, active INTEGER);
  `);
  return db;
}

test("June maintenance schedule uses the latest closing meter and creates a review workbook", async () => {
  const db = makeDb();
  db.prepare("INSERT INTO assets (id, asset_code, asset_name, category) VALUES (1, 'GS04AM', 'John Deere 30KVA', 'Generator')").run();
  // Asset-hours is stale. The daily closing meter is the reliable source June
  // must use, preventing the old cumulative-hours error from resurfacing.
  db.prepare("INSERT INTO asset_hours (asset_id, total_hours) VALUES (1, 35498.7)").run();
  db.prepare("INSERT INTO daily_hours (asset_id, work_date, closing_hours, hours_run, is_used) VALUES (1, '2026-10-07', 2887, 9, 1)").run();
  db.prepare("INSERT INTO daily_hours (asset_id, work_date, closing_hours, hours_run, is_used) VALUES (1, '2026-10-06', 2878, 8, 1)").run();
  db.prepare("INSERT INTO maintenance_plans (id, asset_id, service_name, interval_hours, last_service_hours, active) VALUES (1, 1, '500 hour service', 500, 2500, 1)").run();
  db.prepare("INSERT INTO maintenance_plans (id, asset_id, service_name, interval_hours, last_service_hours, active) VALUES (2, 1, '1000 hour service', 1000, 2500, 1)").run();

  const schedule = buildJuneMaintenanceSchedule({
    assetCodes: ["gs04am", "missing01"],
    asOf: "2026-10-08",
    horizonDays: 30,
    dbConn: db,
  });

  assert.deepEqual(schedule.missing_asset_codes, ["MISSING01"]);
  assert.equal(schedule.rows.length, 1);
  assert.equal(schedule.rows[0].asset_code, "GS04AM");
  assert.equal(schedule.rows[0].current_meter, 2887);
  assert.equal(schedule.rows[0].meter_source, "daily_closing");
  assert.equal(schedule.rows[0].due_meter, 3000);
  assert.equal(schedule.rows[0].service_name, "1000 hour service");
  assert.equal(schedule.rows[0].schedule_status, "PLAN IN WINDOW");

  const buffer = await buildJuneMaintenanceScheduleWorkbook(schedule);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), ["Summary", "Maintenance Schedule"]);
  assert.equal(workbook.getWorksheet("Maintenance Schedule").getCell("A5").value, "GS04AM");
  db.close();
});

