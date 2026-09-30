// IRONLOG/api/utils/techOtherTime.js — technician time that is not on a machine.
//
// Workshop housekeeping, training, safety meetings, tool repairs, travel and
// waiting for work have no work order. The technician runs a timer with a
// reason; the blocks show on the shift report and, on submit, go to the
// mechanics timesheet under the general WORKSHOP code (never a machine), so
// paid hours are complete without inflating any machine's cost.

import { localClock, shiftWorkDate } from "./techShift.js";

export const OTHER_ASSET_CODE = "WORKSHOP";
export const OTHER_CATEGORY = "Workshop";

export const OTHER_REASONS = {
  housekeeping: "Workshop housekeeping",
  training: "Training / toolbox talk",
  safety_meeting: "Safety meeting",
  tools: "Tool and workshop repairs",
  travel: "Travel",
  waiting: "Waiting for work",
  other: "Other",
};

export function ensureOtherTimeSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS tech_other_time (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT NOT NULL,
      site_code TEXT NOT NULL DEFAULT 'main',
      reason TEXT NOT NULL,
      note TEXT,
      started_at TEXT NOT NULL,
      ended_at TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_tech_other_user ON tech_other_time(username, started_at);
  `);
}

export function runningOther(db, username) {
  return db.prepare(`SELECT * FROM tech_other_time WHERE LOWER(username) = LOWER(?) AND ended_at IS NULL ORDER BY id DESC LIMIT 1`).get(username) || null;
}

/** Stops the running block (if any) at `at`; returns it. */
export function stopOther(db, username, at) {
  const cur = runningOther(db, username);
  if (!cur) return null;
  const end = at > cur.started_at ? at : cur.started_at;
  db.prepare(`UPDATE tech_other_time SET ended_at = ? WHERE id = ?`).run(end, cur.id);
  return { ...cur, ended_at: end };
}

/** Starts a block (stopping any running one first). */
export function startOther(db, { username, site = "main", reason, note = null, at }) {
  const key = OTHER_REASONS[reason] ? reason : "other";
  stopOther(db, username, at);
  const ins = db.prepare(`INSERT INTO tech_other_time (username, site_code, reason, note, started_at) VALUES (?, ?, ?, ?, ?)`)
    .run(username, site, key, note ? String(note).slice(0, 300) : null, at);
  return { id: Number(ins.lastInsertRowid), reason: key, started_at: at };
}

/** Blocks overlapping [from, to], clipped to it; a running block runs to `to`. */
export function otherBlocks(db, username, { from, to }) {
  return db.prepare(`
    SELECT * FROM tech_other_time
    WHERE LOWER(username) = LOWER(?) AND started_at < ? AND (ended_at IS NULL OR ended_at > ?)
    ORDER BY started_at, id
  `).all(username, to, from).map((r) => {
    const start = r.started_at < from ? from : r.started_at;
    const endRaw = r.ended_at || to;
    const end = endRaw > to ? to : endRaw;
    const ms = Date.parse(end) - Date.parse(start);
    return {
      id: r.id,
      reason: r.reason,
      reason_label: OTHER_REASONS[r.reason] || r.reason,
      note: r.note,
      start,
      end,
      open: !r.ended_at,
      hours: Number((Number.isFinite(ms) && ms > 0 ? ms / 3600000 : 0).toFixed(2)),
    };
  });
}

export function otherHours(blocks) {
  return Number(blocks.reduce((s, b) => s + b.hours, 0).toFixed(2));
}

/** Timesheet rows for other time: one per reason and day, under the WORKSHOP code. */
export function otherTimesheetRows(blocks, { technicianName, day, shiftStart = null } = {}) {
  const byKey = new Map();
  for (const b of blocks) {
    if (b.hours <= 0) continue;
    const workDate = shiftWorkDate(b.start, { day, shiftStart });
    const key = `${b.reason}|${workDate}`;
    const e = byKey.get(key) || { reason: b.reason, label: b.reason_label, workDate, hours: 0, first: b.start, last: b.end, notes: [] };
    e.hours += b.hours;
    if (b.start < e.first) e.first = b.start;
    if (b.end > e.last) e.last = b.end;
    if (b.note && !e.notes.includes(b.note)) e.notes.push(b.note);
    byKey.set(key, e);
  }
  return [...byKey.values()]
    .sort((a, b) => String(a.first).localeCompare(String(b.first)))
    .map((e) => ({
      work_date: e.workDate,
      technician_name: technicianName,
      hours: Number(e.hours.toFixed(2)),
      asset_code: OTHER_ASSET_CODE,
      reason: `Other time — ${e.label}${e.notes.length ? `: ${e.notes.join("; ")}` : ""}`.slice(0, 300),
      work_order_id: null,
      time_started: localClock(e.first),
      time_finished: localClock(e.last),
      category: OTHER_CATEGORY,
    }))
    .filter((r) => r.hours >= 0.01);
}
