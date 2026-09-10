import test from "node:test";
import assert from "node:assert/strict";
import {
  dailyPdfDowntimeHours,
  serviceLabelFromDailyDowntime,
} from "../routes/reports.routes.js";

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

test("numeric service intervals are presented as maintenance", () => {
  assert.equal(
    serviceLabelFromDailyDowntime(
      "500 hour service",
      "Short breakdown — BREAKDOWN — Service — 500 hour service completed",
    ),
    "500 hour service",
  );
});
