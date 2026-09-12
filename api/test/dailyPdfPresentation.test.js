import test from "node:test";
import assert from "node:assert/strict";
import {
  dailyPdfDowntimeHours,
  dailyPdfRepairFallbackHours,
  dailyPdfOperationsDate,
  serviceLabelFromDailyDowntime,
} from "../routes/reports.routes.js";

test("Daily PDF speeding uses the same completed operations day", () => {
  assert.equal(dailyPdfOperationsDate("2026-09-12"), "2026-09-11");
  assert.equal(dailyPdfOperationsDate("2027-01-01"), "2026-12-31");
});

test("Daily PDF keeps actual downtime when production was recorded", () => {
  // A300AM: a 10.6-hour Production row must not inherit an 11-hour fallback
  // merely because its repair work order is still open.
  assert.equal(dailyPdfDowntimeHours({
    hasDailyEntry: true,
    isUsed: 1,
    recordedHours: 1,
    totalHours: 11,
  }), 1);
});

test("Daily PDF retains the full-shift fallback only without production", () => {
  assert.equal(dailyPdfDowntimeHours({
    hasDailyEntry: false,
    isUsed: 0,
    recordedHours: 0,
    totalHours: 11,
  }), 11);
});

test("Daily PDF does not stack work-order repair estimates onto logged downtime", () => {
  // A300AM: the clerk entered 1 hour. A later work-order close must not add
  // its repair estimate and turn the row into an 11-hour loss.
  assert.equal(dailyPdfRepairFallbackHours({
    dayCap: 11,
    loggedHours: 1,
    allocatedHours: 0,
    repairHours: 10,
  }), 0);
});

test("numeric service intervals are presented as maintenance", () => {
  assert.equal(
    serviceLabelFromDailyDowntime(
      "500 hour service",
      "Short breakdown — BREAKDOWN — Service — 500 hour service completed",
    ),
    "500 hour service",
  );
});
