import test from "node:test";
import assert from "node:assert/strict";

const { __test } = await import("../utils/juneIcsCalendar.js");

const ymd = (d) => d.toISOString().slice(0, 10);
const icsStamp = (d) => d.toISOString().replace(/[-:]/g, "").replace(/\.\d{3}/, "");

test("calendar view gets past events in its range; June's briefing only upcoming ones", () => {
  const now = new Date();
  const twoDaysAgo = new Date(now.getTime() - 2 * 86_400_000);
  twoDaysAgo.setUTCHours(8, 0, 0, 0);
  const inThreeDays = new Date(now.getTime() + 3 * 86_400_000);
  inThreeDays.setUTCHours(9, 0, 0, 0);
  const lastWeek = new Date(now.getTime() - 7 * 86_400_000);
  lastWeek.setUTCHours(6, 30, 0, 0);
  const ics = [
    "BEGIN:VCALENDAR",
    "BEGIN:VEVENT", "UID:past", "SUMMARY:Monday planning", `DTSTART:${icsStamp(twoDaysAgo)}`, `DTEND:${icsStamp(new Date(twoDaysAgo.getTime() + 3_600_000))}`, "END:VEVENT",
    "BEGIN:VEVENT", "UID:next", "SUMMARY:Supplier visit", `DTSTART:${icsStamp(inThreeDays)}`, `DTEND:${icsStamp(new Date(inThreeDays.getTime() + 1_800_000))}`, "LOCATION:Workshop", "END:VEVENT",
    "BEGIN:VEVENT", "UID:daily", "SUMMARY:Toolbox talk", `DTSTART:${icsStamp(lastWeek)}`, `DTEND:${icsStamp(new Date(lastWeek.getTime() + 900_000))}`, "RRULE:FREQ=DAILY", "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");

  const upcoming = __test.upcomingEvents(ics).map((e) => e.summary);
  assert.ok(upcoming.includes("Supplier visit"));
  assert.ok(!upcoming.includes("Monday planning"), "briefing skips past events");

  const from = ymd(new Date(now.getTime() - 3 * 86_400_000));
  const to = ymd(new Date(now.getTime() + 3 * 86_400_000));
  const ranged = __test.upcomingEvents(ics, { from, to });
  const names = ranged.map((e) => e.summary);
  assert.ok(names.includes("Monday planning"), "week view shows earlier days");
  assert.ok(names.includes("Supplier visit"));
  assert.equal(names.filter((n) => n === "Toolbox talk").length, 7, "daily recurrence on each of the 7 days");
  const visit = ranged.find((e) => e.summary === "Supplier visit");
  assert.ok(visit.start_at && visit.end_at, "start and end times for the grid");
});
