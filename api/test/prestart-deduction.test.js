import test from "node:test";
import assert from "node:assert/strict";
import { PRESTART_DEDUCTION_HOURS } from "../utils/prestartDaily.js";

test("completed pre-starts deduct 15 minutes from availability by default", () => {
  assert.equal(PRESTART_DEDUCTION_HOURS, 0.25);
});
