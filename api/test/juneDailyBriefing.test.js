import test from "node:test";
import assert from "node:assert/strict";
import { briefingSite, buildDailyBriefing, getSiteWeather, siteToday } from "../utils/juneDailyBriefing.js";

const site = briefingSite({});

test("briefing site defaults to Palma and can be overridden on the server", () => {
  assert.equal(site.name, "Palma, Cabo Delgado");
  assert.equal(site.timezone, "Africa/Maputo");
  assert.ok(site.nearby.includes("Mocímboa da Praia"));
  const other = briefingSite({ JUNE_BRIEFING_LOCATION: "Pemba", JUNE_BRIEFING_LAT: "-12.97", JUNE_BRIEFING_LON: "40.52", JUNE_BRIEFING_NEARBY: "Pemba, Metuge" });
  assert.deepEqual([other.name, other.latitude, other.longitude, other.nearby], ["Pemba", -12.97, 40.52, ["Pemba", "Metuge"]]);
});

test("today is the site's date, not UTC", () => {
  // 23:30 UTC on 8 Oct is 01:30 on 9 Oct in Maputo.
  assert.equal(siteToday("Africa/Maputo", new Date("2026-10-08T23:30:00Z")), "2026-10-09");
});

test("weather comes back in plain terms from Open-Meteo", async () => {
  let url = "";
  const fetchImpl = async (u) => {
    url = u;
    return {
      ok: true,
      json: async () => ({
        current: { temperature_2m: 27.4, relative_humidity_2m: 78, weather_code: 2, wind_speed_10m: 14 },
        daily: { weather_code: [80], temperature_2m_max: [31], temperature_2m_min: [23], precipitation_probability_max: [60], precipitation_sum: [4.2], wind_speed_10m_max: [22], wind_gusts_10m_max: [38] },
      }),
    };
  };
  const w = await getSiteWeather(site, { fetchImpl });
  assert.match(url, /latitude=-10\.7768/);
  assert.match(url, /timezone=Africa%2FMaputo/);
  assert.equal(w.available, true);
  assert.equal(w.now.conditions, "partly cloudy");
  assert.deepEqual([w.today.conditions, w.today.high_c, w.today.rain_chance_pct, w.today.gusts_max_kmh], ["light showers", 31, 60, 38]);
});

test("weather failures never break the briefing; June is told to web search instead", async () => {
  const down = await getSiteWeather(site, { fetchImpl: async () => { throw new Error("ECONNREFUSED"); } });
  assert.equal(down.available, false);
  const slow = await getSiteWeather(site, {
    timeoutMs: 20,
    fetchImpl: (_u, { signal }) => new Promise((_, reject) => signal.addEventListener("abort", () => reject(Object.assign(new Error("aborted"), { name: "AbortError" })))),
  });
  assert.equal(slow.reason, "Weather service timed out.");

  const b = await buildDailyBriefing({
    site,
    now: new Date("2026-10-08T05:00:00Z"),
    fetchImpl: async () => ({ ok: false, status: 503 }),
    ironlogBrief: (day) => ({
      as_of: day,
      metrics: { open_breakdowns: 2, critical_breakdowns: 1 },
      open_breakdowns: [{ asset_code: "A303AM", critical: true }, { asset_code: "E504AM", critical: false }],
      pm_due: [{ asset_code: "G01AM", status: "overdue" }, { asset_code: "F500AM", status: "due_soon" }],
    }),
    calendarToday: (day) => ({ date: day, ironlog_schedule: [{ title: "Production meeting" }] }),
  });
  assert.equal(b.date, "2026-10-08");
  assert.equal(b.breakdowns.metrics.critical_breakdowns, 1);
  assert.deepEqual(b.breakdowns.pm_overdue.map((r) => r.asset_code), ["G01AM"]);
  assert.equal(b.calendar_today.ironlog_schedule[0].title, "Production meeting");
  assert.equal(b.weather.available, false);
  assert.match(b.weather.next_step, /web search for today's weather in Palma/);
  assert.equal(b.security_check.method, "web_search");
  assert.match(b.security_check.search_for, /Palma.*last 7 days before 2026-10-08/);
});

test("one failing source does not sink the rest", async () => {
  const b = await buildDailyBriefing({
    site,
    fetchImpl: async () => ({ ok: false, status: 500 }),
    ironlogBrief: () => { throw new Error("db locked"); },
    calendarToday: async () => ({ ironlog_schedule: [] }),
  });
  assert.deepEqual(b.breakdowns, { error: "db locked" });
  assert.deepEqual(b.calendar_today, { ironlog_schedule: [] });
});
