import test from "node:test";
import assert from "node:assert/strict";
import { famsSelectedDateRange } from "../utils/famsFuelRange.js";

test("FAMS catch-up range accepts a completed August period", () => {
  assert.deepEqual(
    famsSelectedDateRange({
      startDate: "2026-08-01",
      endDate: "2026-08-31",
      now: new Date("2026-09-14T10:00:00+02:00"),
    }),
    {
      startYmd: "2026-08-01",
      endYmd: "2026-08-31",
      startSlash: "2026/08/01",
      endSlash: "2026/08/31",
    },
  );
});

test("FAMS catch-up range rejects unsafe calendar ranges", () => {
  const now = new Date("2026-09-14T10:00:00+02:00");
  assert.throws(
    () => famsSelectedDateRange({ startDate: "2026-08-31", endDate: "2026-08-01", now }),
    /cannot be after/i,
  );
  assert.throws(
    () => famsSelectedDateRange({ startDate: "2026-02-30", endDate: "2026-03-01", now }),
    /valid calendar date/i,
  );
  assert.throws(
    () => famsSelectedDateRange({ startDate: "2026-09-01", endDate: "2026-09-15", now }),
    /later than today/i,
  );
});
