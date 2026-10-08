import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Database from "better-sqlite3";
// The utility has a default production DB import even though these tests pass
// an explicit in-memory connection. Keep that default import in memory too.
const defaultDbDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-june-calendar-unit-"));
process.env.DB_PATH = path.join(defaultDbDir, "ironlog.db");
const {
  cancelInternalCalendarEvent,
  createInternalCalendarEvent,
  getInternalCalendarEvent,
  listInternalCalendarEvents,
  updateInternalCalendarEvent,
} = await import("../utils/juneInternalCalendar.js");
const { db: defaultDb } = await import("../db/client.js");

test.after(() => {
  defaultDb.close();
  rmSync(defaultDbDir, { recursive: true, force: true });
});

test("June internal calendar keeps an owner's schedule separate and preserves cancelled entries", () => {
  const db = new Database(":memory:");
  const jaco = { siteCode: "main", user: "jaco" };
  const anotherAdmin = { siteCode: "main", user: "another-admin" };

  const created = createInternalCalendarEvent(jaco, {
    title: "GS04AM service planning",
    event_date: "2026-10-12",
    start_time: "08:00",
    end_time: "09:00",
    category: "maintenance",
    asset_code: "gs04am",
    work_order_id: 349,
  }, { dbConn: db });
  assert.equal(created.title, "GS04AM service planning");
  assert.equal(created.asset_code, "GS04AM");
  assert.equal(created.status, "scheduled");
  assert.equal(listInternalCalendarEvents(anotherAdmin, { fromDate: "2026-10-01", dbConn: db }).length, 0);

  const moved = updateInternalCalendarEvent(jaco, created.id, {
    event_date: "2026-10-13",
    start_time: "10:00",
    end_time: "11:30",
  }, { dbConn: db });
  assert.equal(moved.event_date, "2026-10-13");
  assert.equal(moved.start_time, "10:00");
  assert.equal(moved.end_time, "11:30");

  const cancelled = cancelInternalCalendarEvent(jaco, created.id, { dbConn: db });
  assert.equal(cancelled.status, "cancelled");
  assert.equal(getInternalCalendarEvent(jaco, created.id, { dbConn: db }), null);
  assert.equal(getInternalCalendarEvent(jaco, created.id, { dbConn: db, includeCancelled: true }).status, "cancelled");
  assert.equal(listInternalCalendarEvents(jaco, { fromDate: "2026-10-01", dbConn: db }).length, 0);
  db.close();
});

test("June internal calendar rejects nonsensical dates and times", () => {
  const db = new Database(":memory:");
  const context = { siteCode: "main", user: "jaco" };
  assert.throws(() => createInternalCalendarEvent(context, {
    title: "Impossible meeting",
    event_date: "2026-02-30",
    start_time: "10:00",
    end_time: "09:00",
  }, { dbConn: db }), /valid calendar date/i);
  assert.throws(() => createInternalCalendarEvent(context, {
    title: "Backwards meeting",
    event_date: "2026-10-12",
    start_time: "10:00",
    end_time: "09:00",
  }, { dbConn: db }), /later than start/i);
  db.close();
});
