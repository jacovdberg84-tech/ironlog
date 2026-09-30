// IRONLOG/api/utils/techActivity.js — technician activity on work orders.
//
// The portal stores what a technician does as an append-only list of events
// (start, pause, resume, waiting for parts/operations, testing, complete).
// Time is derived from those events and never typed twice:
//   active + testing           = labour
//   paused, waiting_parts/ops  = not labour
// Work order status itself stays in work_orders and only changes through the
// normal status route.

export const TECH_ACTIONS = ["start", "pause", "resume", "waiting_parts", "waiting_ops", "testing", "complete"];
export const LABOUR_STATES = new Set(["active", "testing"]);
export const RUNNING_STATES = new Set(["active", "testing"]);

// Action -> resulting state, and which states it may follow.
const RULES = {
  start: { to: "active", from: ["idle", "paused", "waiting_parts", "waiting_ops", "done"] },
  resume: { to: "active", from: ["paused", "waiting_parts", "waiting_ops", "testing"] },
  pause: { to: "paused", from: ["active", "testing", "waiting_parts", "waiting_ops"] },
  waiting_parts: { to: "waiting_parts", from: ["active", "testing", "paused", "waiting_ops", "idle"] },
  waiting_ops: { to: "waiting_ops", from: ["active", "testing", "paused", "waiting_parts", "idle"] },
  testing: { to: "testing", from: ["active", "paused", "waiting_parts", "waiting_ops"] },
  complete: { to: "done", from: ["active", "testing", "paused", "waiting_parts", "waiting_ops"] },
};

export const STATE_LABELS = {
  idle: "Not started",
  active: "Working",
  paused: "Paused",
  waiting_parts: "Waiting for parts",
  waiting_ops: "Waiting for operations",
  testing: "Testing",
  done: "Completed",
};

export function ensureTechSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS tech_activity_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      work_order_id INTEGER NOT NULL,
      username TEXT NOT NULL,
      action TEXT NOT NULL,
      at TEXT NOT NULL,
      note TEXT,
      auto INTEGER NOT NULL DEFAULT 0,
      client_event_id TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_tech_events_wo ON tech_activity_events(work_order_id, at);
    CREATE INDEX IF NOT EXISTS idx_tech_events_user ON tech_activity_events(username, at);

    CREATE TABLE IF NOT EXISTS tech_findings (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      work_order_id INTEGER,
      asset_id INTEGER,
      username TEXT NOT NULL,
      kind TEXT NOT NULL DEFAULT 'finding',
      text TEXT NOT NULL,
      at TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_tech_findings_wo ON tech_findings(work_order_id);

    CREATE TABLE IF NOT EXISTS work_order_photos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      work_order_id INTEGER NOT NULL,
      username TEXT,
      file_path TEXT NOT NULL,
      caption TEXT,
      at TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_wo_photos_wo ON work_order_photos(work_order_id);

    CREATE TABLE IF NOT EXISTS work_order_technicians (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      work_order_id INTEGER NOT NULL,
      username TEXT NOT NULL,
      role TEXT NOT NULL DEFAULT 'helper',
      added_by TEXT,
      added_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(work_order_id, username)
    );

    CREATE TABLE IF NOT EXISTS tech_shifts (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT NOT NULL,
      site_code TEXT NOT NULL DEFAULT 'main',
      started_at TEXT NOT NULL,
      ended_at TEXT,
      status TEXT NOT NULL DEFAULT 'open',
      findings TEXT,
      unplanned TEXT,
      safety TEXT,
      outstanding TEXT,
      handover TEXT,
      next_shift TEXT,
      submitted_at TEXT,
      timesheet_rows INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_tech_shifts_user ON tech_shifts(username, started_at);

    -- One row per offline-capable write, so a retried submission returns the
    -- first result instead of saving twice.
    CREATE TABLE IF NOT EXISTS tech_client_events (
      client_event_id TEXT PRIMARY KEY,
      username TEXT,
      kind TEXT NOT NULL,
      result_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
  `);
}

/** The state after an action, or an error message when it does not follow. */
export function nextState(current, action) {
  const rule = RULES[action];
  if (!rule) return { error: `unknown action '${action}'` };
  const from = current || "idle";
  if (!rule.from.includes(from)) {
    return { error: `cannot ${action.replace(/_/g, " ")} while ${STATE_LABELS[from]?.toLowerCase() || from}` };
  }
  return { state: rule.to };
}

/** Current state per technician on a work order, from its events (oldest first). */
export function statesByUser(events) {
  const out = new Map();
  for (const e of events) {
    const cur = out.get(e.username) || "idle";
    const r = nextState(cur, e.action);
    if (!r.error) out.set(e.username, r.state);
  }
  return out;
}

/**
 * Time segments from events: each accepted event starts a segment in its new
 * state that runs to the same technician's next event on that work order, or to
 * `until` (now / shift end) while still open. Events must be sorted by time.
 */
export function buildSegments(events, { until = new Date().toISOString() } = {}) {
  const segments = [];
  const open = new Map(); // key wo|user -> { state, start }
  for (const e of events) {
    const key = `${e.work_order_id}|${e.username}`;
    const cur = open.get(key);
    const r = nextState(cur ? cur.state : "idle", e.action);
    if (r.error) continue;
    if (cur && cur.state !== "done" && cur.state !== "idle") {
      segments.push({ work_order_id: e.work_order_id, username: e.username, state: cur.state, start: cur.start, end: e.at });
    }
    open.set(key, { state: r.state, start: e.at });
  }
  for (const [key, cur] of open) {
    if (cur.state === "done" || cur.state === "idle") continue;
    const [wo, user] = key.split("|");
    const end = cur.start > until ? cur.start : until;
    segments.push({ work_order_id: Number(wo), username: user, state: cur.state, start: cur.start, end, open: true });
  }
  return segments.sort((a, b) => String(a.start).localeCompare(String(b.start)));
}

export function segmentHours(seg) {
  const ms = Date.parse(seg.end) - Date.parse(seg.start);
  return Number.isFinite(ms) && ms > 0 ? ms / 3600000 : 0;
}

/** Labour hours (active + testing) from segments, optionally clipped to a window. */
export function labourHours(segments, { from = null, to = null } = {}) {
  let total = 0;
  for (const s of segments) {
    if (!LABOUR_STATES.has(s.state)) continue;
    const start = from && s.start < from ? from : s.start;
    const end = to && s.end > to ? to : s.end;
    if (end > start) total += segmentHours({ start, end });
  }
  return Number(total.toFixed(2));
}

/**
 * Accepts a client timestamp for offline capture when it is sane (not more than
 * 5 minutes in the future, not older than 72 hours); otherwise uses server time.
 */
export function eventTime(clientAt, now = new Date()) {
  const t = Date.parse(String(clientAt || ""));
  const n = now.getTime();
  if (!Number.isFinite(t) || t > n + 5 * 60000 || t < n - 72 * 3600000) return now.toISOString();
  return new Date(t).toISOString();
}

/** Local calendar day (Mozambique / South Africa, UTC+2) for an ISO time. */
export function localDay(iso = new Date().toISOString()) {
  return new Intl.DateTimeFormat("en-CA", { timeZone: "Africa/Maputo", year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date(iso));
}

/** Start of a local day as ISO (UTC+2, no daylight saving). */
export function localDayStartIso(day) {
  return new Date(`${day}T00:00:00+02:00`).toISOString();
}
