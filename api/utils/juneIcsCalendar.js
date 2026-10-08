// IRONLOG/api/utils/juneIcsCalendar.js
// Private ICS calendar source for June. The signed URL is encrypted at rest,
// fetched only by the API, and never returned to the browser or GPT-Live.
import dns from "node:dns/promises";
import net from "node:net";
import { db } from "../db/client.js";
import { decryptJuneConnectorSecret, encryptJuneConnectorSecret } from "./juneOutlook.js";

const FETCH_TIMEOUT_MS = 20_000;
const MAX_ICS_BYTES = 5 * 1024 * 1024;
const MAX_REDIRECTS = 3;
const WINDOW_DAYS = 21;

function text(value, max = 500) {
  return String(value || "").trim().slice(0, max);
}

function contextKey(context = {}) {
  return {
    siteCode: text(context.siteCode, 160).toLowerCase() || "main",
    user: text(context.user, 160) || "session-user",
  };
}

function ensureTables() {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS june_ics_calendars (
      site_code TEXT NOT NULL,
      username TEXT NOT NULL COLLATE NOCASE,
      calendar_name TEXT,
      calendar_url_enc TEXT NOT NULL,
      enabled INTEGER NOT NULL DEFAULT 1,
      last_fetched_at TEXT,
      last_error TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      PRIMARY KEY (site_code, username)
    )
  `).run();
}

function sourceRow(context) {
  ensureTables();
  const { siteCode, user } = contextKey(context);
  return db.prepare(`
    SELECT site_code, username, calendar_name, calendar_url_enc, enabled, last_fetched_at, last_error, created_at, updated_at
    FROM june_ics_calendars
    WHERE site_code = ? AND username = ?
  `).get(siteCode, user) || null;
}

function isPrivateAddress(address) {
  const version = net.isIP(address);
  if (version === 4) {
    const [a, b] = address.split(".").map(Number);
    return (
      a === 0 || a === 10 || a === 127 ||
      (a === 100 && b >= 64 && b <= 127) ||
      (a === 169 && b === 254) ||
      (a === 172 && b >= 16 && b <= 31) ||
      (a === 192 && b === 0) ||
      (a === 192 && b === 168) ||
      (a === 198 && (b === 18 || b === 19)) ||
      a >= 224
    );
  }
  if (version === 6) {
    const normalized = String(address || "").toLowerCase();
    return normalized === "::" || normalized === "::1" || normalized.startsWith("fc") || normalized.startsWith("fd") || normalized.startsWith("fe80") || normalized.includes("::ffff:127.");
  }
  return true;
}

async function validateIcsUrl(raw) {
  let parsed;
  try {
    parsed = new URL(text(raw, 4000));
  } catch {
    throw new Error("Enter a valid private ICS calendar link.");
  }
  if (parsed.protocol === "webcal:") parsed.protocol = "https:";
  if (parsed.protocol !== "https:") throw new Error("Calendar links must use HTTPS.");
  if (parsed.username || parsed.password) throw new Error("Calendar links must not contain a username or password.");
  if (parsed.port && parsed.port !== "443") throw new Error("Calendar links must use the standard HTTPS port.");
  const host = String(parsed.hostname || "").toLowerCase();
  if (!host || host === "localhost" || host.endsWith(".localhost")) throw new Error("That calendar address is not allowed.");
  if (net.isIP(host)) {
    if (isPrivateAddress(host)) throw new Error("Private network calendar addresses are not allowed.");
  } else {
    let addresses;
    try {
      addresses = await dns.lookup(host, { all: true, verbatim: true });
    } catch {
      throw new Error("Ironlog could not resolve that calendar address.");
    }
    if (!addresses.length || addresses.some((entry) => isPrivateAddress(entry.address))) {
      throw new Error("That calendar address is not allowed.");
    }
  }
  return parsed;
}

async function fetchIcsDocument(rawUrl) {
  let url = await validateIcsUrl(rawUrl);
  for (let redirects = 0; redirects <= MAX_REDIRECTS; redirects += 1) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);
    try {
      const response = await fetch(url, {
        redirect: "manual",
        headers: { Accept: "text/calendar, text/plain;q=0.8, */*;q=0.1", "User-Agent": "IRONLOG-June-Calendar/1.0" },
        signal: controller.signal,
      });
      if ([301, 302, 303, 307, 308].includes(response.status)) {
        const location = response.headers.get("location");
        if (!location) throw new Error("The calendar host returned an incomplete redirect.");
        url = await validateIcsUrl(new URL(location, url).toString());
        continue;
      }
      if (!response.ok) throw new Error(`The calendar host returned HTTP ${response.status}.`);
      const declaredSize = Number(response.headers.get("content-length") || 0);
      if (Number.isFinite(declaredSize) && declaredSize > MAX_ICS_BYTES) throw new Error("The calendar feed is too large.");
      const body = Buffer.from(await response.arrayBuffer());
      if (body.length > MAX_ICS_BYTES) throw new Error("The calendar feed is too large.");
      const document = body.toString("utf8");
      if (!/BEGIN:VCALENDAR/i.test(document) || !/BEGIN:VEVENT/i.test(document)) {
        throw new Error("That link did not return an iCalendar (.ics) feed.");
      }
      return document;
    } catch (error) {
      if (error?.name === "AbortError") throw new Error("The calendar host took too long to respond.");
      throw error;
    } finally {
      clearTimeout(timeout);
    }
  }
  throw new Error("The calendar redirected too many times.");
}

function unfoldedLines(ics) {
  const out = [];
  for (const line of String(ics || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n")) {
    if (/^[ \t]/.test(line) && out.length) out[out.length - 1] += line.slice(1);
    else out.push(line);
  }
  return out;
}

function property(line) {
  const splitAt = line.indexOf(":");
  if (splitAt <= 0) return null;
  const left = line.slice(0, splitAt);
  const [name, ...rawParams] = left.split(";");
  const params = {};
  for (const part of rawParams) {
    const eq = part.indexOf("=");
    if (eq > 0) params[part.slice(0, eq).toUpperCase()] = part.slice(eq + 1).replace(/^"|"$/g, "");
  }
  return { name: name.toUpperCase(), params, value: line.slice(splitAt + 1) };
}

function unescapeIcs(value) {
  return String(value || "").replace(/\\n/gi, " ").replace(/\\,/g, ",").replace(/\\;/g, ";").replace(/\\\\/g, "\\");
}

function dateValue(raw) {
  const value = String(raw || "").trim();
  const allDay = /^\d{8}$/.test(value);
  const match = value.match(/^(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?)?(Z)?$/);
  if (!match) return null;
  const [, year, month, day, hour = "00", minute = "00", second = "00", utc] = match;
  const timestamp = Date.UTC(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second));
  return {
    timestamp,
    allDay,
    date: `${year}-${month}-${day}`,
    time: allDay ? null : `${hour}:${minute}`,
    utc: Boolean(utc),
  };
}

function parseRrule(value) {
  const out = {};
  for (const part of String(value || "").split(";")) {
    const [key, raw] = part.split("=");
    if (key && raw) out[key.toUpperCase()] = raw.toUpperCase();
  }
  return out;
}

function parseEvents(ics) {
  const events = [];
  let active = null;
  for (const line of unfoldedLines(ics)) {
    if (line.toUpperCase() === "BEGIN:VEVENT") {
      active = { exdates: [] };
      continue;
    }
    if (line.toUpperCase() === "END:VEVENT") {
      if (active?.start) events.push(active);
      active = null;
      continue;
    }
    if (!active) continue;
    const entry = property(line);
    if (!entry) continue;
    if (entry.name === "DTSTART") active.start = dateValue(entry.value);
    else if (entry.name === "DTEND") active.end = dateValue(entry.value);
    else if (entry.name === "SUMMARY") active.summary = unescapeIcs(entry.value);
    else if (entry.name === "LOCATION") active.location = unescapeIcs(entry.value);
    else if (entry.name === "UID") active.uid = text(entry.value, 280);
    else if (entry.name === "RRULE") active.rrule = parseRrule(entry.value);
    else if (entry.name === "EXDATE") active.exdates.push(...String(entry.value).split(",").map(dateValue).filter(Boolean));
    else if (entry.name === "STATUS" && String(entry.value).toUpperCase() === "CANCELLED") active.cancelled = true;
  }
  return events.filter((event) => !event.cancelled && event.start);
}

function daysBetween(start, end) {
  return Math.floor((end - start) / 86_400_000);
}

function monthDifference(start, end) {
  const a = new Date(start);
  const b = new Date(end);
  return (b.getUTCFullYear() - a.getUTCFullYear()) * 12 + b.getUTCMonth() - a.getUTCMonth();
}

function weekdayCode(timestamp) {
  return ["SU", "MO", "TU", "WE", "TH", "FR", "SA"][new Date(timestamp).getUTCDay()];
}

function isRecurringOccurrence(event, timestamp, windowEnd) {
  const rule = event.rrule;
  if (!rule) return timestamp === event.start.timestamp;
  const interval = Math.max(1, Number(rule.INTERVAL || 1) || 1);
  const until = dateValue(rule.UNTIL)?.timestamp || Number.POSITIVE_INFINITY;
  if (timestamp < event.start.timestamp || timestamp > until || timestamp > windowEnd) return false;
  const deltaDays = daysBetween(event.start.timestamp, timestamp);
  if (deltaDays < 0) return false;
  if (rule.FREQ === "DAILY") return deltaDays % interval === 0;
  if (rule.FREQ === "WEEKLY") {
    const selectedDays = String(rule.BYDAY || weekdayCode(event.start.timestamp)).split(",");
    return Math.floor(deltaDays / 7) % interval === 0 && selectedDays.includes(weekdayCode(timestamp));
  }
  if (rule.FREQ === "MONTHLY") {
    const start = new Date(event.start.timestamp);
    const current = new Date(timestamp);
    return monthDifference(event.start.timestamp, timestamp) % interval === 0 && current.getUTCDate() === start.getUTCDate();
  }
  if (rule.FREQ === "YEARLY") {
    const start = new Date(event.start.timestamp);
    const current = new Date(timestamp);
    return (current.getUTCFullYear() - start.getUTCFullYear()) % interval === 0 && current.getUTCMonth() === start.getUTCMonth() && current.getUTCDate() === start.getUTCDate();
  }
  return false;
}

function eventDuration(event) {
  const fallback = event.start.allDay ? 86_400_000 : 3_600_000;
  const duration = Number(event?.end?.timestamp || 0) - Number(event?.start?.timestamp || 0);
  return duration > 0 && duration <= 14 * 86_400_000 ? duration : fallback;
}

function formatEvent(event, startTimestamp) {
  const start = new Date(startTimestamp);
  const end = new Date(startTimestamp + eventDuration(event));
  const stamp = (date) => date.toISOString().replace(".000Z", "Z");
  const startDate = start.toISOString().slice(0, 10);
  const startTime = event.start.allDay ? null : start.toISOString().slice(11, 16);
  return {
    uid: text(event.uid, 280) || `${startTimestamp}-${text(event.summary, 80)}`,
    summary: text(event.summary, 180) || "(No title)",
    location: text(event.location, 160) || null,
    all_day: Boolean(event.start.allDay),
    start_at: stamp(start),
    end_at: stamp(end),
    start_date: startDate,
    start_time: startTime,
  };
}

/**
 * Occurrences in a window. By default: still-running or upcoming events in the
 * next WINDOW_DAYS (June's briefing). With a range (the calendar view): every
 * event touching [from, to], past ones included.
 */
function upcomingEvents(ics, range = null) {
  const now = range ? Date.parse(`${range.from}T00:00:00Z`) : Date.now();
  const start = new Date(now);
  start.setUTCHours(0, 0, 0, 0);
  const windowEnd = range ? Date.parse(`${range.to}T23:59:59Z`) : start.getTime() + WINDOW_DAYS * 86_400_000;
  const limit = range ? 400 : 18;
  const all = [];
  for (const event of parseEvents(ics).slice(0, 500)) {
    const excluded = new Set((event.exdates || []).map((item) => item.timestamp));
    if (!event.rrule) {
      if (event.start.timestamp + eventDuration(event) >= now && event.start.timestamp <= windowEnd && !excluded.has(event.start.timestamp)) {
        all.push(formatEvent(event, event.start.timestamp));
      }
      continue;
    }
    const countLimit = Math.max(1, Math.min(400, Number(event.rrule.COUNT || 400) || 400));
    let emitted = 0;
    // Never walk a daily loop from an event that began years ago.  The
    // recurrence test remains anchored to DTSTART, but scanning starts just
    // before the useful window.
    const firstCandidate = Math.max(event.start.timestamp, start.getTime() - 86_400_000);
    const daysToSkip = Math.max(0, Math.ceil((firstCandidate - event.start.timestamp) / 86_400_000));
    const firstOccurrenceDay = event.start.timestamp + daysToSkip * 86_400_000;
    for (let timestamp = firstOccurrenceDay; timestamp <= windowEnd && emitted < countLimit; timestamp += 86_400_000) {
      if (!isRecurringOccurrence(event, timestamp, windowEnd)) continue;
      emitted += 1;
      if (timestamp + eventDuration(event) >= now && !excluded.has(timestamp)) all.push(formatEvent(event, timestamp));
    }
  }
  return all.sort((a, b) => a.start_at.localeCompare(b.start_at)).slice(0, limit);
}

function safeRefreshError(error) {
  const message = text(error?.message || error, 240);
  if (/HTTP 401|HTTP 403/i.test(message)) return "The calendar link is no longer authorised. Create a new private ICS link and save it again.";
  if (/too long|too large/i.test(message)) return message;
  return "Ironlog could not refresh this calendar. Check that the ICS link is still valid and private.";
}

export function getIcsCalendarStatus(context) {
  const row = sourceRow(context);
  if (!row || Number(row.enabled) !== 1) {
    return { state: "not_connected", detail: "Add a private ICS link to show your calendar in Ironlog.", name: null, last_fetched_at: null };
  }
  if (!decryptJuneConnectorSecret(row.calendar_url_enc)) {
    return { state: "needs_reconnect", detail: "The saved calendar link needs to be added again.", name: text(row.calendar_name, 120) || "My calendar", last_fetched_at: row.last_fetched_at || null };
  }
  return {
    state: "connected",
    detail: row.last_error ? "Calendar saved. The latest refresh needs attention." : "Private ICS calendar connected.",
    name: text(row.calendar_name, 120) || "My calendar",
    last_fetched_at: row.last_fetched_at || null,
    last_error: text(row.last_error, 220) || null,
  };
}

export async function saveIcsCalendar(context, { url, name }) {
  const source = await validateIcsUrl(url);
  const { siteCode, user } = contextKey(context);
  ensureTables();
  let encrypted;
  try {
    encrypted = encryptJuneConnectorSecret(source.toString());
  } catch {
    throw new Error("Calendar storage is not configured on the server yet. Refresh Ironlog after the latest deployment, then try again.");
  }
  const label = text(name, 120) || "My calendar";
  db.prepare(`
    INSERT INTO june_ics_calendars (site_code, username, calendar_name, calendar_url_enc, enabled, last_fetched_at, last_error, created_at, updated_at)
    VALUES (?, ?, ?, ?, 1, NULL, NULL, datetime('now'), datetime('now'))
    ON CONFLICT(site_code, username) DO UPDATE SET
      calendar_name = excluded.calendar_name,
      calendar_url_enc = excluded.calendar_url_enc,
      enabled = 1,
      last_fetched_at = NULL,
      last_error = NULL,
      updated_at = datetime('now')
  `).run(siteCode, user, label, encrypted);
  return getIcsCalendarStatus(context);
}

export function removeIcsCalendar(context) {
  const { siteCode, user } = contextKey(context);
  ensureTables();
  const result = db.prepare("DELETE FROM june_ics_calendars WHERE site_code = ? AND username = ?").run(siteCode, user);
  return { removed: Number(result.changes || 0) > 0, calendar: getIcsCalendarStatus(context) };
}

export async function getIcsCalendarEvents(context, range = null) {
  const status = getIcsCalendarStatus(context);
  if (status.state !== "connected") return { ...status, events: [], next_step: "Add a private ICS link before refreshing your calendar." };
  const row = sourceRow(context);
  const url = decryptJuneConnectorSecret(row?.calendar_url_enc);
  if (!url) return { ...status, state: "needs_reconnect", events: [], next_step: "Add the calendar link again." };
  try {
    const document = await fetchIcsDocument(url);
    const events = upcomingEvents(document, range);
    const { siteCode, user } = contextKey(context);
    db.prepare(`
      UPDATE june_ics_calendars
      SET last_fetched_at = datetime('now'), last_error = NULL, updated_at = datetime('now')
      WHERE site_code = ? AND username = ?
    `).run(siteCode, user);
    return { ...getIcsCalendarStatus(context), events, period_days: WINDOW_DAYS };
  } catch (error) {
    const detail = safeRefreshError(error);
    const { siteCode, user } = contextKey(context);
    db.prepare(`UPDATE june_ics_calendars SET last_error = ?, updated_at = datetime('now') WHERE site_code = ? AND username = ?`).run(detail, siteCode, user);
    return { ...getIcsCalendarStatus(context), events: [], error: detail };
  }
}

export async function getIcsCalendarOverview(context) {
  const data = await getIcsCalendarEvents(context);
  return {
    ...data,
    review_only: true,
    source: "private_ics",
    next_step: data.state === "connected"
      ? "June can brief you on this calendar. Booking and changes are not enabled."
      : data.next_step || "Add a private ICS calendar link first.",
  };
}
export const __test = { upcomingEvents };
