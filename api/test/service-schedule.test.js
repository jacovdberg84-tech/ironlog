import test from "node:test";
import assert from "node:assert/strict";
import { resolveNextServiceForAssetPlans } from "../utils/serviceSchedule.js";

const plantPlan = {
  id: 7,
  asset_id: 11,
  asset_code: "A300AM",
  service_name: "500 hour service",
  interval_hours: 500,
  active: 1,
};

test("automatically flags a 500 hour service at the next half-thousand milestone", () => {
  const next = resolveNextServiceForAssetPlans([plantPlan], 19498, "A300AM");
  assert.equal(next.next_due_hours, 19500);
  assert.equal(next.next_service_interval, 500);
  assert.equal(next.service_name, "500 hour service");
  assert.equal(next.remaining_hours, 2);
});

test("automatically flags a 1000 hour service at the next thousand milestone", () => {
  const next = resolveNextServiceForAssetPlans([plantPlan], 19985, "A300AM");
  assert.equal(next.next_due_hours, 20000);
  assert.equal(next.next_service_interval, 1000);
  assert.equal(next.service_name, "1000 hour service");
  assert.equal(next.remaining_hours, 15);
});

test("LDVs always calculate on a 10000 km grid", () => {
  const next = resolveNextServiceForAssetPlans([{
    ...plantPlan,
    asset_code: "V01AM",
    service_name: "Incorrect legacy interval",
    interval_hours: 500,
  }], 19985, "V01AM");
  assert.equal(next.next_due_hours, 20000);
  assert.equal(next.next_service_interval, 10000);
  assert.equal(next.service_name, "10000 km service");
  assert.equal(next.remaining_hours, 15);
});
