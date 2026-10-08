// IRONLOG/api/utils/juneDailyBriefing.js — facts for June's daily briefing.
//
// The "Daily briefing" button starts June with an instruction to open with
// a short rundown: open breakdowns, today's calendar, the weather at site and
// any recent security incidents nearby. This file supplies the site, today's
// local date and the weather (Open-Meteo, no key needed). Breakdowns and the
// calendar come from June's existing helpers; security news comes from June's
// hosted web search, because there is no reliable feed for it.

const DEFAULT_SITE = {
  name: "Palma, Cabo Delgado",
  country: "Mozambique",
  latitude: -10.7768,
  longitude: 40.4764,
  timezone: "Africa/Maputo",
  // Towns June checks for security incidents around the site.
  nearby: ["Palma", "Afungi", "Mocímboa da Praia", "Nangade", "Muidumbe", "Mueda", "Macomia", "Quionga"],
};

const WEATHER_TIMEOUT_MS = 6_000;

/** The briefing site; override with JUNE_BRIEFING_* settings on the server. */
export function briefingSite(env = process.env) {
  const num = (v, fallback) => (Number.isFinite(Number(v)) && String(v).trim() !== "" ? Number(v) : fallback);
  const nearby = String(env.JUNE_BRIEFING_NEARBY || "").split(",").map((s) => s.trim()).filter(Boolean);
  return {
    name: String(env.JUNE_BRIEFING_LOCATION || DEFAULT_SITE.name).trim(),
    country: String(env.JUNE_BRIEFING_COUNTRY || DEFAULT_SITE.country).trim(),
    latitude: num(env.JUNE_BRIEFING_LAT, DEFAULT_SITE.latitude),
    longitude: num(env.JUNE_BRIEFING_LON, DEFAULT_SITE.longitude),
    timezone: String(env.JUNE_BRIEFING_TZ || DEFAULT_SITE.timezone).trim(),
    nearby: nearby.length ? nearby : DEFAULT_SITE.nearby,
  };
}

/** Today's date (YYYY-MM-DD) at site, not in UTC. */
export function siteToday(timezone, now = new Date()) {
  try {
    return now.toLocaleDateString("en-CA", { timeZone: timezone });
  } catch {
    return now.toISOString().slice(0, 10);
  }
}

// WMO weather interpretation codes used by Open-Meteo.
const WMO = {
  0: "clear sky", 1: "mainly clear", 2: "partly cloudy", 3: "overcast",
  45: "fog", 48: "freezing fog",
  51: "light drizzle", 53: "drizzle", 55: "heavy drizzle",
  61: "light rain", 63: "rain", 65: "heavy rain",
  80: "light showers", 81: "showers", 82: "violent showers",
  95: "thunderstorms", 96: "thunderstorms with hail", 99: "severe thunderstorms with hail",
};
const describe = (code) => WMO[Number(code)] || "mixed conditions";

/** Today's weather at site. Never throws: on failure June falls back to web search. */
export async function getSiteWeather(site, { fetchImpl = globalThis.fetch, timeoutMs = WEATHER_TIMEOUT_MS } = {}) {
  const params = new URLSearchParams({
    latitude: String(site.latitude),
    longitude: String(site.longitude),
    current: "temperature_2m,relative_humidity_2m,weather_code,wind_speed_10m",
    daily: "weather_code,temperature_2m_max,temperature_2m_min,precipitation_probability_max,precipitation_sum,wind_speed_10m_max,wind_gusts_10m_max",
    timezone: site.timezone,
    forecast_days: "1",
    wind_speed_unit: "kmh",
  });
  if (typeof fetchImpl !== "function") return { available: false, reason: "This server cannot make web requests." };
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetchImpl(`https://api.open-meteo.com/v1/forecast?${params}`, { signal: controller.signal });
    if (!res.ok) return { available: false, reason: `Weather service returned HTTP ${res.status}.` };
    const data = await res.json();
    const d = data?.daily || {};
    const c = data?.current || {};
    const first = (arr) => (Array.isArray(arr) ? arr[0] : null);
    return {
      available: true,
      source: "Open-Meteo",
      location: site.name,
      now: {
        temperature_c: c.temperature_2m ?? null,
        humidity_pct: c.relative_humidity_2m ?? null,
        wind_kmh: c.wind_speed_10m ?? null,
        conditions: c.weather_code != null ? describe(c.weather_code) : null,
      },
      today: {
        conditions: describe(first(d.weather_code)),
        high_c: first(d.temperature_2m_max),
        low_c: first(d.temperature_2m_min),
        rain_chance_pct: first(d.precipitation_probability_max),
        rain_mm: first(d.precipitation_sum),
        wind_max_kmh: first(d.wind_speed_10m_max),
        gusts_max_kmh: first(d.wind_gusts_10m_max),
      },
    };
  } catch (error) {
    const reason = error?.name === "AbortError" ? "Weather service timed out." : "Weather service could not be reached from the server.";
    return { available: false, reason };
  } finally {
    clearTimeout(timer);
  }
}

/** What June should search for, so the security part stays factual and local. */
export function securityCheck(site, today) {
  return {
    method: "web_search",
    search_for: `Recent insurgent attacks or security incidents near ${site.nearby.join(", ")} (${site.name}, ${site.country}) in the last 7 days before ${today}.`,
    rules: [
      "Only report incidents dated within the last 7 days, near the listed places.",
      "Give the place, the date and the source for each one; keep it to the facts reported.",
      "If nothing recent is found, say so plainly. Never guess or fill in details.",
      "Keep this part serious: no jokes and no swearing.",
    ],
  };
}

/**
 * The briefing tool result. Breakdowns and calendar are passed in from the
 * routes (they own those helpers); weather is fetched here.
 */
export async function buildDailyBriefing({ site = briefingSite(), now = new Date(), ironlogBrief, calendarToday, fetchImpl } = {}) {
  const today = siteToday(site.timezone, now);
  const [brief, calendar, weather] = await Promise.all([
    Promise.resolve().then(() => ironlogBrief(today)).catch((e) => ({ error: String(e?.message || e) })),
    Promise.resolve().then(() => calendarToday(today)).catch((e) => ({ error: String(e?.message || e) })),
    getSiteWeather(site, fetchImpl ? { fetchImpl } : {}),
  ]);
  const breakdowns = Array.isArray(brief?.open_breakdowns) ? brief.open_breakdowns : [];
  return {
    review_only: true,
    date: today,
    timezone: site.timezone,
    site: { name: site.name, country: site.country },
    breakdowns: brief?.error ? { error: brief.error } : {
      metrics: brief.metrics,
      open: breakdowns,
      pm_overdue: (brief.pm_due || []).filter((r) => r.status === "overdue").slice(0, 8),
    },
    calendar_today: calendar,
    weather: weather.available ? weather : {
      ...weather,
      next_step: `Use web search for today's weather in ${site.name}, ${site.country}.`,
    },
    security_check: securityCheck(site, today),
    order: [
      "Breakdowns: count, critical ones first, what is waiting on parts, anything down a long time.",
      "Today's calendar: meetings and reminders in time order.",
      `Weather in ${site.name} today.`,
      "Security: recent incidents near site, or that nothing recent was found.",
      "Finish by asking what to tackle first.",
    ],
  };
}

export const BRIEFING_LIVE_INSTRUCTIONS = [
  "Daily briefing: as soon as this conversation starts, before Jaco says anything, greet him and give his daily briefing.",
  "Delegate to the backend for it: use the daily briefing tool for breakdowns, today's calendar and weather, and web search for recent security incidents near site.",
  "Follow the order in the tool result. Keep the whole briefing to about a minute of speech: highlights, not every row.",
  "Keep the security part serious and factual, with dates and sources. Then ask what he wants to tackle first.",
].join(" ");
