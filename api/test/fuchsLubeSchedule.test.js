import test from "node:test";
import assert from "node:assert/strict";
import {
  buildFuchsServiceOilItems,
  getFuchsLubricationProfile,
  getFuchsServiceRequirements,
} from "../utils/fuchsLubeSchedule.js";

test("matches a named fleet model to its Fuchs lubricants and full-system capacities", () => {
  const profile = getFuchsLubricationProfile({
    equipmentName: "Cat 140K Motor Grader",
    category: "Grader",
  });

  assert.equal(profile?.matched_model, "CAT 140K");
  const engine = profile.components.find((component) => component.component === "ENGINE");
  assert.equal(engine?.product, "Fuchs Titan Cargo MC 10W40");
  assert.deepEqual(engine?.capacities_l, [{ label: "CAPACITY", litres: 18 }]);
  assert.equal(engine?.oil_drain_interval, 500);
});

test("creates service oil lines only for Fuchs items due at the selected interval", () => {
  const profile = getFuchsLubricationProfile({ equipmentName: "CAT 140K" });
  const due = getFuchsServiceRequirements(profile, 500);
  const oils = buildFuchsServiceOilItems(profile, 500);

  assert.deepEqual(due.map((item) => item.component), ["ENGINE"]);
  assert.deepEqual(oils, [{
    name: "ENGINE: Fuchs Titan Cargo MC 10W40",
    qty: 18,
    unit: "L",
    part_hint: "Fuchs Titan Cargo MC 10W40",
    component: "ENGINE",
    source: "Fuchs equipment lube schedule",
  }]);
});

test("keeps Fuchs configuration variants visible instead of guessing a requisition quantity", () => {
  const profile = getFuchsLubricationProfile({ equipmentName: "Cat D6R Dozer" });
  const due = getFuchsServiceRequirements(profile, 1000);
  const powerTrain = due.find((item) => item.component === "POWER TRAIN");
  const oils = buildFuchsServiceOilItems(profile, 1000);

  assert.equal(powerTrain?.system_capacity_l, null);
  assert.deepEqual(powerTrain?.capacity_options_l, [
    { label: "BLT/TBC", litres: 148 },
    { label: "S6X", litres: 146 },
  ]);
  assert.equal(oils.some((item) => item.component === "POWER TRAIN"), false);
});

test("matches the shared Fuchs truck profile for Mercedes and Axor equipment", () => {
  const profile = getFuchsLubricationProfile({
    equipmentName: "Mercedes Benz Axor 3335K",
    category: "Truck",
  });

  assert.equal(profile?.matched_model, "ALL MERCEDES BENZ TRUCKS");
  assert.equal(profile?.components.find((component) => component.component === "ENGINE")?.capacities_l[0]?.litres, 39);
});
