import test from "node:test";
import assert from "node:assert/strict";
import { benchmarkIsSet, fillToFill } from "../utils/fuelConsumption.js";

const fill = (date, liters, reading, unit = "hours") => ({ id: date + reading, log_date: date, liters, meter_unit: unit, meter_run_value: reading });
const prev = fill("2026-08-31", 300, 10000);

test("clean fills: litres over hours since the previous fill", () => {
  const r = fillToFill([fill("2026-09-01", 200, 10010), fill("2026-09-02", 200, 10020)], prev, "hours");
  assert.equal(r.run, 20);
  assert.equal(r.matched_liters, 400);
  assert.equal(r.coverage, 1);
  assert.equal(r.suspect.length, 0);
});

test("a typo that jumps the meter is ignored; its litres go with the next good reading", () => {
  // 10010 → 10520 (typed 105 instead of 100) → 10030: without the check this added 510 h.
  const r = fillToFill([fill("2026-09-01", 200, 10010), fill("2026-09-02", 200, 10520), fill("2026-09-03", 200, 10030)], prev, "hours");
  assert.equal(r.run, 30, "10 + 20 hours, the typo is skipped");
  assert.equal(r.matched_liters, 600);
  assert.equal(r.suspect.length, 1);
  assert.equal(r.suspect[0].reading, 10520);
  assert.equal(r.suspect[0].action, "ignored");
  assert.equal(r.trace[1].status, "jump of 510 h: reading ignored, litres carried to the 2026-09-03 fill");
  assert.match(r.trace[2].status, /^ok$/);
  assert.equal(r.trace[2].interval_liters, 400);
});

test("a typo that goes backwards is ignored too", () => {
  const r = fillToFill([fill("2026-09-01", 200, 10010), fill("2026-09-02", 200, 1002), fill("2026-09-03", 200, 10030)], prev, "hours");
  assert.equal(r.run, 30);
  assert.equal(r.matched_liters, 600);
  assert.equal(r.suspect[0].reason, "meter went back");
});

test("missing readings: litres wait for the next reading (no more '978 L/hr' from fuel with no hours)", () => {
  const r = fillToFill([fill("2026-09-01", 400, 0), fill("2026-09-02", 400, null), fill("2026-09-03", 400, 10030)], prev, "hours");
  assert.equal(r.run, 30);
  assert.equal(r.matched_liters, 1200);
  assert.equal((r.matched_liters / r.run).toFixed(1), "40.0");
});

test("meter replaced: counting restarts from the new meter", () => {
  const r = fillToFill([fill("2026-09-01", 200, 10010), fill("2026-09-05", 200, 5), fill("2026-09-06", 200, 15), fill("2026-09-07", 200, 25)], prev, "hours");
  assert.equal(r.run, 30, "10 h on the old meter, 20 h on the new one");
  assert.equal(r.suspect[0].action, "counted from here");
  assert.equal(r.unmatched_liters, 200, "the fill at the meter change cannot be matched");
});

test("no earlier reading: the first fill cannot be matched; fuel after the last reading is not counted", () => {
  const r = fillToFill([fill("2026-09-01", 200, 10010), fill("2026-09-02", 200, 10020), fill("2026-09-03", 150, 0)], null, "hours");
  assert.equal(r.run, 10);
  assert.equal(r.matched_liters, 200);
  assert.equal(r.unmatched_liters, 350);
  assert.match(r.trace[2].status, /no later reading in period/);
});

test("km machines: odometer between fills; readings in the other unit are not mixed in", () => {
  const p = fill("2026-08-31", 60, 50000, "km");
  const r = fillToFill([fill("2026-09-01", 60, 50600, "km"), fill("2026-09-02", 60, 3020, "hours"), fill("2026-09-03", 60, 51200, "km")], p, "km");
  assert.equal(r.run, 1200);
  assert.equal(r.matched_liters, 180);
  assert.equal(r.other_unit_fills, 1);
});

test("benchmark counts as set unless it is still the column default and was never saved", () => {
  assert.equal(benchmarkIsSet(5, 5, null), false);
  assert.equal(benchmarkIsSet(5, 5, "2026-10-01"), true);
  assert.equal(benchmarkIsSet(35, 5, null), true);
  assert.equal(benchmarkIsSet(null, 5, null), false);
});
