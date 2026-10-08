// IRONLOG/api/utils/juneInternalCalendar.js
// A private, Ironlog-owned calendar for June. It is deliberately separate
// from external ICS and Outlook sources, which remain read-only.
import { db } from "../db/client.js";

const EVENT_CATEGORIES = new Set(["meeting", "maintenance", "shutdown", "reminder", "follow_up", "other"]);

function text(value, max = 500) {
  return String(value ?? "").trim().slice(0, max);
}

function hasOwn(value, key) {
  return Object.prototype.hasOwnProperty.call(value || {}, key);
}

function contextKey(context = {}) {
  return {
    siteCode: text(context.siteCode, 160).toLowerCase() || "main",
    user: text(context.user, 160) || "session-user",
  };
}

function isYmd(value) {
  const raw = text(value, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return false;
  const date = new Date(`${raw}T00:00:00.000Z`);
  return !Number.isNaN(date.getTime()) && date.toISOString().slice(0, 10) === raw;
}

function isTime(value) {
  return /^([01]\d|2[0-3]):[0-5]\d$/.test(text(value, 5));
}

function boolean(value, fallback = false) {
  if (value === true || value === 1 || value === "1") return true;
  if (value === false || value === 0 || value === "0") return false;
  return fallback;
}

function nullableText(value, max = 500) {
  const cleaned = text(value, max);
  return cleaned || null;
}

function positiveId(value) {
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed > 0 ? parsed : null;
}

function ensureTables(dbConn = db) {
  dbConn.prepare(`
    CREATE TABLE IF NOT EXISTS june_internal_calendar_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL,
      username TEXT NOT NULL COLLATE NOCASE,
      title TEXT NOT NULL,
      event_date TEXT NOT NULL,
      start_time TEXT,
      end_time TEXT,
      all_day INTEGER NOT NULL DEFAULT 0,
      category TEXT NOT NULL DEFAULT 'meeting',
      asset_code TEXT,
      work_order_id INTEGER,
      notes TEXT,
      status TEXT NOT NULL DEFAULT 'scheduled',
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      cancelled_at TEXT,
      cancelled_by TEXT
    )
  `).run();
  dbConn.prepare("CREATE INDEX IF NOT EXISTS idx_june_internal_calendar_events_owner_date ON june_internal_calendar_events (site_code, username, status, event_date)").run();
}

function publicEvent(row) {
  if (!row) return null;
  return {
    id: Number(row.id),
    title: text(row.title, 180),
    event_date: text(row.event_date, 10),
    start_time: nullableText(row.start_time, 5),
    end_time: nullableText(row.end_time, 5),
    all_day: Boolean(Number(row.all_day)),
    category: text(row.category, 32) || "meeting",
    asset_code: nullableText(row.asset_code, 40),
    work_order_id: positiveId(row.work_order_id),
    notes: nullableText(row.notes, 1200),
    status: text(row.status, 24) || "scheduled",
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
    cancelled_at: row.cancelled_at || null,
  };
}

function eventRow(context, id, dbConn = db) {
  ensureTables(dbConn);
  const { siteCode, user } = contextKey(context);
  return dbConn.prepare(`
    SELECT * FROM june_internal_calendar_events
    WHERE id = ? AND site_code = ? AND username = ?
    LIMIT 1
  `).get(positiveId(id) || -1, siteCode, user) || null;
}

// This is exported so June can show a validated proposal before the browser
// explicitly approves any write. It never touches the database.
export function prepareInternalCalendarEvent(input = {}, current = null) {
  const patch = input && typeof input === "object" ? input : {};
  const base = current ? publicEvent(current) : {};
  const title = hasOwn(patch, "title") ? text(patch.title, 180) : text(base.title, 180);
  const eventDate = hasOwn(patch, "event_date") ? text(patch.event_date, 10) : text(base.event_date, 10);
  const allDay = hasOwn(patch, "all_day") ? boolean(patch.all_day, false) : Boolean(base.all_day);
  const startTime = hasOwn(patch, "start_time") ? nullableText(patch.start_time, 5) : nullableText(base.start_time, 5);
  const endTime = hasOwn(patch, "end_time") ? nullableText(patch.end_time, 5) : nullableText(base.end_time, 5);
  const categoryRaw = hasOwn(patch, "category") ? text(patch.category, 32).toLowerCase() : text(base.category || "meeting", 32).toLowerCase();
  const assetCode = hasOwn(patch, "asset_code") ? nullableText(patch.asset_code, 40)?.toUpperCase() || null : nullableText(base.asset_code, 40)?.toUpperCase() || null;
  const workOrderId = hasOwn(patch, "work_order_id") ? positiveId(patch.work_order_id) : positiveId(base.work_order_id);
  const notes = hasOwn(patch, "notes") ? nullableText(patch.notes, 1200) : nullableText(base.notes, 1200);

  if (!title) throw new Error("June needs a calendar title before she can prepare it.");
  if (!isYmd(eventDate)) throw new Error("Use a valid calendar date in YYYY-MM-DD format.");
  if (!EVENT_CATEGORIES.has(categoryRaw)) throw new Error("Choose a valid calendar category.");
  if (!allDay) {
    if (startTime && !isTime(startTime)) throw new Error("Start time must use 24-hour HH:MM format.");
    if (endTime && !isTime(endTime)) throw new Error("End time must use 24-hour HH:MM format.");
    if (startTime && endTime && endTime <= startTime) throw new Error("End time must be later than start time.");
  }
  return {
    title,
    event_date: eventDate,
    start_time: allDay ? null : startTime,
    end_time: allDay ? null : endTime,
    all_day: allDay,
    category: categoryRaw,
    asset_code: assetCode,
    work_order_id: workOrderId,
    notes,
  };
}

export function getInternalCalendarEvent(context, id, { dbConn = db, includeCancelled = false } = {}) {
  const row = eventRow(context, id, dbConn);
  if (!row || (!includeCancelled && String(row.status).toLowerCase() === "cancelled")) return null;
  return publicEvent(row);
}

export function listInternalCalendarEvents(context, { fromDate, toDate, limit = 18, includeCancelled = false, dbConn = db } = {}) {
  ensureTables(dbConn);
  const { siteCode, user } = contextKey(context);
  const today = new Date().toISOString().slice(0, 10);
  const start = isYmd(fromDate) ? String(fromDate) : today;
  const end = isYmd(toDate) ? String(toDate) : "9999-12-31";
  const take = Math.max(1, Math.min(100, Number(limit) || 18));
  const rows = dbConn.prepare(`
    SELECT * FROM june_internal_calendar_events
    WHERE site_code = ? AND username = ?
      AND event_date >= ? AND event_date <= ?
      AND (? = 1 OR status <> 'cancelled')
    ORDER BY event_date ASC, CASE WHEN all_day = 1 THEN '00:00' ELSE COALESCE(start_time, '23:59') END ASC, id ASC
    LIMIT ?
  `).all(siteCode, user, start, end, includeCancelled ? 1 : 0, take);
  return rows.map(publicEvent);
}

export function createInternalCalendarEvent(context, input, { actor, dbConn = db } = {}) {
  ensureTables(dbConn);
  const event = prepareInternalCalendarEvent(input);
  const { siteCode, user } = contextKey(context);
  const result = dbConn.prepare(`
    INSERT INTO june_internal_calendar_events
      (site_code, username, title, event_date, start_time, end_time, all_day, category, asset_code, work_order_id, notes, status, created_by, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'scheduled', ?, datetime('now'), datetime('now'))
  `).run(siteCode, user, event.title, event.event_date, event.start_time, event.end_time, event.all_day ? 1 : 0, event.category, event.asset_code, event.work_order_id, event.notes, text(actor || user, 160) || user);
  return getInternalCalendarEvent(context, result.lastInsertRowid, { dbConn });
}

export function updateInternalCalendarEvent(context, id, patch, { actor, dbConn = db } = {}) {
  const current = eventRow(context, id, dbConn);
  if (!current || String(current.status).toLowerCase() === "cancelled") throw new Error("That Ironlog calendar entry is no longer available to update.");
  const event = prepareInternalCalendarEvent(patch, current);
  const { siteCode, user } = contextKey(context);
  dbConn.prepare(`
    UPDATE june_internal_calendar_events
    SET title = ?, event_date = ?, start_time = ?, end_time = ?, all_day = ?, category = ?, asset_code = ?, work_order_id = ?, notes = ?, updated_at = datetime('now')
    WHERE id = ? AND site_code = ? AND username = ? AND status <> 'cancelled'
  `).run(event.title, event.event_date, event.start_time, event.end_time, event.all_day ? 1 : 0, event.category, event.asset_code, event.work_order_id, event.notes, positiveId(id), siteCode, user);
  void actor;
  return getInternalCalendarEvent(context, id, { dbConn });
}

export function cancelInternalCalendarEvent(context, id, { actor, dbConn = db } = {}) {
  const current = eventRow(context, id, dbConn);
  if (!current || String(current.status).toLowerCase() === "cancelled") throw new Error("That Ironlog calendar entry is no longer available to cancel.");
  const { siteCode, user } = contextKey(context);
  dbConn.prepare(`
    UPDATE june_internal_calendar_events
    SET status = 'cancelled', cancelled_at = datetime('now'), cancelled_by = ?, updated_at = datetime('now')
    WHERE id = ? AND site_code = ? AND username = ? AND status <> 'cancelled'
  `).run(text(actor || user, 160) || user, positiveId(id), siteCode, user);
  return getInternalCalendarEvent(context, id, { dbConn, includeCancelled: true });
}

export function getInternalCalendarOverview(context, { days = 21, dbConn = db } = {}) {
  const windowDays = Math.max(1, Math.min(90, Number(days) || 21));
  const start = new Date();
  const end = new Date(start);
  end.setUTCDate(end.getUTCDate() + windowDays);
  return {
    source: "internal_ironlog",
    review_only: false,
    writable_with_confirmation: true,
    window_days: windowDays,
    events: listInternalCalendarEvents(context, {
      fromDate: start.toISOString().slice(0, 10),
      toDate: end.toISOString().slice(0, 10),
      limit: 24,
      dbConn,
    }),
    next_step: "June can prepare an internal calendar change, then the administrator must confirm it in Ironlog.",
  };
}
