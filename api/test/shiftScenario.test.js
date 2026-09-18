import test from "node:test";
import assert from "node:assert/strict";
import { buildShiftScenario } from "../utils/shiftScenario.js";

test("shift scenario scales fuel with hours and holds other recorded cost per active shift", () => {
  const result = buildShiftScenario([{
    asset_id: 1,
    asset_code: "A300AM",
    asset_name: "Bell B30D",
    active_days: 2,
    run_hours: 20,
    fuel_liters: 220,
    fuel_cost: 440,
    nonfuel_cost: 220,
  }], { baseHours: 11, scenarioHours: 8 });

  assert.equal(result.rows.length, 1);
  const row = result.rows[0];
  assert.equal(row.actual_liters_per_hour, 11);
  assert.equal(row.base_fuel_liters_per_shift, 121);
  assert.equal(row.scenario_fuel_liters_per_shift, 88);
  assert.equal(row.fuel_liters_saved_per_shift, 33);
  assert.equal(row.base_total_cost_per_shift, 352);
  assert.equal(row.scenario_total_cost_per_shift, 286);
  assert.equal(row.base_cost_per_operating_hour, 32);
  assert.equal(row.scenario_cost_per_operating_hour, 35.75);
  assert.equal(row.cost_per_operating_hour_change, 3.75);
  assert.equal(result.fleet.fuel_liters_saved, 66);
  assert.equal(result.fleet.total_cost_saved, 132);
});

test("shift scenario excludes distance and missing-data assets instead of inventing a result", () => {
  const result = buildShiftScenario([
    { asset_code: "V01AM", metric_mode: "km", active_days: 2, run_hours: 20, fuel_liters: 30 },
    { asset_code: "A301AM", active_days: 2, run_hours: 0, fuel_liters: 30 },
    { asset_code: "A302AM", active_days: 2, run_hours: 10, fuel_liters: 0 },
  ]);

  assert.deepEqual(result.rows, []);
  assert.deepEqual(result.excluded.map((row) => [row.asset_code, row.reason]), [
    ["V01AM", "Distance-based asset (km)"],
    ["A301AM", "No recorded operating hours"],
    ["A302AM", "No fuel logged in selected period"],
  ]);
});
