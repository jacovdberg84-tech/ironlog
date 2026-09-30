// IRONLOG/api/utils/techShift.js — shift timeline and timesheet rows.
//
// A shift report is built from the technician's activity segments between shift
// start and end (so a night shift over midnight is one report). On submit the
// labour part of that timeline becomes rows in the mechanics timesheet, one per
// work order, tagged with the work order number.

import { LABOUR_STATES, STATE_LABELS, localDay, segmentHours } from "./techActivity.js";

function clip(seg, from, to) {
  const start = from && seg.start < from ? from : seg.start;
  const end = to && seg.end > to ? to : seg.end;
  return end > start ? { ...seg, start, end } : null;
}

/** Local clock time (UTC+2) as HH:MM. */
export function localClock(iso) {
  return new Intl.DateTimeFormat("en-GB", { timeZone: "Africa/Maputo", hour: "2-digit", minute: "2-digit", hour12: false }).format(new Date(iso));
}

function sourceCategory(source) {
  const s = String(source || "").toLowerCase();
  if (s === "service") return "Service";
  if (s === "manual" || s === "maintenance") return "Maintenance";
  return "Breakdown";
}

/**
 * Timeline blocks for the shift window, oldest first. `workOrders` is a Map
 * id -> { asset_code, asset_name, job, source, status }.
 */
export function buildShiftTimeline(segments, { from, to, workOrders = new Map() } = {}) {
  const out = [];
  for (const s of segments) {
    const c = clip(s, from, to);
    if (!c) continue;
    const w = workOrders.get(Number(c.work_order_id)) || {};
    out.push({
      work_order_id: Number(c.work_order_id),
      asset_code: w.asset_code || null,
      asset_name: w.asset_name || null,
      job: w.job || null,
      source: w.source || null,
      state: c.state,
      state_label: STATE_LABELS[c.state] || c.state,
      labour: LABOUR_STATES.has(c.state),
      start: c.start,
      end: c.end,
      open: Boolean(s.open && c.end === s.end),
      hours: Number(segmentHours(c).toFixed(2)),
      wo_status: w.status || null,
    });
  }
  return out.sort((a, b) => String(a.start).localeCompare(String(b.start)));
}

// Work within this many hours of the shift start belongs to the shift's day (a
// night shift over midnight stays on one date). Later work — a report left open
// into the next day — goes on the day it was done.
const SHIFT_DAY_HOURS = 16;

/**
 * Timesheet rows from a shift timeline: labour time only, one row per work
 * order and day, from its first start to last finish.
 */
/** The timesheet date for a block of work in a shift (see SHIFT_DAY_HOURS). */
export function shiftWorkDate(blockStart, { day, shiftStart = null } = {}) {
  const startMs = shiftStart ? Date.parse(shiftStart) : NaN;
  const late = Number.isFinite(startMs) && Date.parse(blockStart) - startMs >= SHIFT_DAY_HOURS * 3600000;
  return late ? localDay(blockStart) : day;
}

export function shiftTimesheetRows(timeline, { technicianName, day, shiftStart = null } = {}) {
  const byKey = new Map();
  for (const t of timeline) {
    if (!t.labour || !t.asset_code || t.hours <= 0) continue;
    const workDate = shiftWorkDate(t.start, { day, shiftStart });
    const key = `${t.work_order_id}|${workDate}`;
    const e = byKey.get(key) || { ...t, workDate, hours: 0, first: t.start, last: t.end };
    e.hours += t.hours;
    if (t.start < e.first) e.first = t.start;
    if (t.end > e.last) e.last = t.end;
    byKey.set(key, e);
  }
  return [...byKey.values()]
    .sort((a, b) => String(a.first).localeCompare(String(b.first)))
    .map((e) => ({
      work_date: e.workDate,
      technician_name: technicianName,
      hours: Number(e.hours.toFixed(2)),
      asset_code: e.asset_code,
      reason: `WO #${e.work_order_id}: ${e.job || "Work order"}`.slice(0, 300),
      work_order_id: e.work_order_id,
      time_started: localClock(e.first),
      time_finished: localClock(e.last),
      category: sourceCategory(e.source),
    }))
    .filter((r) => r.hours >= 0.01);
}

export const SHIFT_REMIND_HOURS = 12;

/**
 * Shift reports still open after `hours` (a technician forgot to submit).
 * Oldest first, with how long each has been open.
 */
export function overdueShifts(db, { site = null, hours = SHIFT_REMIND_HOURS, now = new Date() } = {}) {
  const has = db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'tech_shifts'`).get();
  if (!has) return [];
  const cutoff = new Date(now.getTime() - hours * 3600000).toISOString();
  const hasUsers = db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'users'`).get();
  const rows = db.prepare(`
    SELECT s.id, s.username, s.started_at, ${hasUsers ? "u.full_name" : "NULL"} AS full_name
    FROM tech_shifts s
    ${hasUsers ? "LEFT JOIN users u ON LOWER(u.username) = LOWER(s.username)" : ""}
    WHERE s.status = 'open' AND s.started_at <= ?
      ${site ? "AND LOWER(TRIM(COALESCE(s.site_code, 'main'))) = ?" : ""}
    ORDER BY s.started_at ASC
  `).all(...(site ? [cutoff, String(site).toLowerCase()] : [cutoff]));
  return rows.map((r) => ({
    id: r.id,
    username: r.username,
    name: String(r.full_name || "").trim() || r.username,
    started_at: r.started_at,
    hours_open: Number(((now.getTime() - Date.parse(r.started_at)) / 3600000).toFixed(1)),
  }));
}
