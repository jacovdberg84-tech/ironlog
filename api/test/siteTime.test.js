import test from "node:test";
import assert from "node:assert/strict";
import { setSiteTimeZoneProvider, siteClock, siteDay, siteDayStartIso, siteStamp, siteTimeZone, siteWallTimeMs } from "../utils/siteTime.js";
import { localDay, localDayStartIso } from "../utils/techActivity.js";
import { settingsFiles } from "../loadEnv.js";

test("site days and clock times follow the company's time zone", () => {
  setSiteTimeZoneProvider(() => "Africa/Maputo");
  assert.equal(siteTimeZone(), "Africa/Maputo");
  assert.equal(localDay("2026-10-14T22:30:00Z"), "2026-10-15", "00:30 next day in Maputo");
  assert.equal(localDayStartIso("2026-10-15"), "2026-10-14T22:00:00.000Z");
  assert.equal(siteClock("2026-10-15T04:05:00Z"), "06:05");
  assert.equal(siteStamp(Date.parse("2026-10-15T04:05:00Z")), "2026-10-15 06:05");
  assert.equal(siteWallTimeMs("2026-10-15", "06:05:00"), Date.parse("2026-10-15T04:05:00Z"));

  setSiteTimeZoneProvider(() => "Africa/Lusaka");
  assert.equal(siteDay("2026-10-14T22:30:00Z"), "2026-10-15");
  setSiteTimeZoneProvider(() => "America/Santiago");
  assert.equal(siteDay("2026-10-15T02:00:00Z"), "2026-10-14", "a site behind UTC");
  setSiteTimeZoneProvider(() => "Not/AZone");
  assert.equal(siteTimeZone(), "Africa/Maputo", "a bad zone falls back instead of breaking");
});

test("daylight saving is handled for sites that have it", () => {
  setSiteTimeZoneProvider(() => "Europe/Lisbon");
  // Lisbon is UTC+1 in summer, UTC+0 in winter.
  assert.equal(siteDayStartIso("2026-07-01"), "2026-06-30T23:00:00.000Z");
  assert.equal(siteDayStartIso("2026-12-01"), "2026-12-01T00:00:00.000Z");
  assert.equal(siteWallTimeMs("2026-03-29", "12:00:00"), Date.parse("2026-03-29T11:00:00Z"), "the day the clocks go forward");
  setSiteTimeZoneProvider(() => process.env.IRONLOG_TIMEZONE || "Africa/Maputo");
});

test("a customer instance reads only its own settings file", () => {
  assert.deepEqual(settingsFiles({ IRONLOG_ENV_FILE: "/data/ironlog-instances/pilot/instance.env" }, "/x"), ["/data/ironlog-instances/pilot/instance.env"]);
  const normal = settingsFiles({}, "/data/ironlog/api");
  assert.equal(normal.length, 2);
  assert.match(normal[0], /api\/\.env$/);
});
