// IRONLOG/api/utils/techShift.js — shift timeline and timesheet rows.
//
// A shift report is built from the technician's activity segments between shift
// start and end (so a night shift over midnight is one report). On submit the
// labour part of that timeline becomes rows in the mechanics timesheet, one per
// work order, tagged with the work order number.

import { LABOUR_STATES, STATE_LABELS, segmentHours } from "./techActivity.js";

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

/**
 * Timesheet rows from a shift timeline: labour time only, one row per work
 * order, from its first start to last finish in the shift.
 */
export function shiftTimesheetRows(timeline, { technicianName, day } = {}) {
  const byWo = new Map();
  for (const t of timeline) {
    if (!t.labour || !t.asset_code || t.hours <= 0) continue;
    const e = byWo.get(t.work_order_id) || { ...t, hours: 0, first: t.start, last: t.end };
    e.hours += t.hours;
    if (t.start < e.first) e.first = t.start;
    if (t.end > e.last) e.last = t.end;
    byWo.set(t.work_order_id, e);
  }
  return [...byWo.values()]
    .sort((a, b) => String(a.first).localeCompare(String(b.first)))
    .map((e) => ({
      work_date: day,
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
