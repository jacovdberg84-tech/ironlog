// IRONLOG/api/utils/siteTime.js — the site's local time, in the company's time zone.
//
// Each IronLog company has its own time zone (Admin → Company profile). The
// server registers a provider at start-up (server.js) that reads it; until then,
// and in tests, the zone is IRONLOG_TIMEZONE or Africa/Maputo, where IronLog
// started. Works for zones with daylight saving too.

const FALLBACK_TZ = "Africa/Maputo";
let provider = () => process.env.IRONLOG_TIMEZONE || FALLBACK_TZ;
let cached = { at: 0, tz: null };

/** server.js: () => company profile time zone. */
export function setSiteTimeZoneProvider(fn) {
  provider = fn;
  cached = { at: 0, tz: null };
}

function valid(tz) {
  try {
    new Intl.DateTimeFormat("en-CA", { timeZone: tz });
    return true;
  } catch {
    return false;
  }
}

/** The site's IANA time zone, re-read at most once a minute. */
export function siteTimeZone() {
  if (cached.tz && Date.now() - cached.at < 60000) return cached.tz;
  let tz = FALLBACK_TZ;
  try {
    const got = String(provider() || "").trim();
    if (got && valid(got)) tz = got;
  } catch {}
  cached = { at: Date.now(), tz };
  return tz;
}

function parts(date, tz) {
  const p = new Intl.DateTimeFormat("en-US", {
    timeZone: tz, hourCycle: "h23", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit",
  }).formatToParts(date);
  const get = (t) => Number(p.find((x) => x.type === t)?.value);
  return { y: get("year"), m: get("month"), d: get("day"), h: get("hour"), mi: get("minute"), s: get("second") };
}

/** How far the zone is ahead of UTC at that moment, in milliseconds. */
export function zoneOffsetMs(tz, date = new Date()) {
  const x = parts(date, tz);
  const asUtc = Date.UTC(x.y, x.m - 1, x.d, x.h, x.mi, x.s);
  return asUtc - Math.floor(date.getTime() / 1000) * 1000;
}

/** A wall-clock time at site ("2026-10-15", "06:30:00") as milliseconds. */
export function siteWallTimeMs(ymd, hms = "00:00:00", tz = siteTimeZone()) {
  const [y, m, d] = String(ymd).split("-").map(Number);
  const [h = 0, mi = 0, s = 0] = String(hms).split(":").map(Number);
  const guess = Date.UTC(y, m - 1, d, h, mi, s);
  const first = guess - zoneOffsetMs(tz, new Date(guess));
  const second = guess - zoneOffsetMs(tz, new Date(first));
  return second;
}

/** The site's calendar day (YYYY-MM-DD) for an ISO time. */
export function siteDay(iso = new Date().toISOString(), tz = siteTimeZone()) {
  return new Intl.DateTimeFormat("en-CA", { timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date(iso));
}

/** The start of a site day as an ISO time. */
export function siteDayStartIso(day, tz = siteTimeZone()) {
  return new Date(siteWallTimeMs(day, "00:00:00", tz)).toISOString();
}

/** Site clock time (HH:MM) for an ISO time. */
export function siteClock(iso, tz = siteTimeZone()) {
  return new Intl.DateTimeFormat("en-GB", { timeZone: tz, hour: "2-digit", minute: "2-digit", hour12: false }).format(new Date(iso));
}

/** "YYYY-MM-DD HH:MM" at site for a millisecond time. */
export function siteStamp(ms, tz = siteTimeZone()) {
  if (!Number.isFinite(ms)) return "—";
  const x = parts(new Date(ms), tz);
  const pad = (n) => String(n).padStart(2, "0");
  return `${x.y}-${pad(x.m)}-${pad(x.d)} ${pad(x.h)}:${pad(x.mi)}`;
}
