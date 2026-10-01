import test from "node:test";
import assert from "node:assert/strict";
import { capDowntime, downtimeCapForRun } from "../utils/downtimeCap.js";

test("a machine cannot be down for more than the shift hours it did not run", () => {
  assert.equal(capDowntime(9, { scheduled: 11, run: 9.5 }), 1.5, "A301AM: 9 h repair, ran 9.5 h of 11");
  assert.equal(capDowntime(9, { scheduled: 11, run: 12 }), 0, "ran the whole shift");
  assert.equal(capDowntime(9, { scheduled: 11, run: 0 }), 9, "did not run: logged downtime stands");
  assert.equal(capDowntime(9, { scheduled: 11, run: null }), 9, "no Daily Input: logged downtime stands");
  assert.equal(capDowntime(1, { scheduled: 11, run: 4 }), 1, "below the cap is unchanged");
  assert.equal(downtimeCapForRun({ scheduled: 11, run: 0 }), Infinity);
});
