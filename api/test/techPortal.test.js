import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";
import { buildSegments, eventTime, labourHours, nextState, statesByUser } from "../utils/techActivity.js";
import { buildShiftTimeline, shiftTimesheetRows } from "../utils/techShift.js";

// ---------------------------------------------------------------- pure rules

const ev = (wo, user, action, at) => ({ work_order_id: wo, username: user, action, at });

test("activity states follow the allowed order", () => {
  assert.deepEqual(nextState("idle", "start"), { state: "active" });
  assert.deepEqual(nextState("active", "waiting_parts"), { state: "waiting_parts" });
  assert.deepEqual(nextState("waiting_parts", "resume"), { state: "active" });
  assert.deepEqual(nextState("active", "complete"), { state: "done" });
  assert.ok(nextState("idle", "pause").error, "cannot pause a job that never started");
  assert.ok(nextState("idle", "complete").error);
  assert.ok(nextState("done", "resume").error);
  assert.ok(nextState("active", "dance").error);
  const states = statesByUser([ev(1, "jose", "start", "a"), ev(1, "jose", "pause", "b"), ev(1, "maria", "start", "c")]);
  assert.equal(states.get("jose"), "paused");
  assert.equal(states.get("maria"), "active");
});

test("labour counts working and testing time only", () => {
  const segs = buildSegments([
    ev(1, "jose", "start", "2026-09-30T06:00:00.000Z"),
    ev(1, "jose", "waiting_parts", "2026-09-30T07:00:00.000Z"),
    ev(1, "jose", "resume", "2026-09-30T09:00:00.000Z"),
    ev(1, "jose", "testing", "2026-09-30T10:00:00.000Z"),
    ev(1, "jose", "complete", "2026-09-30T10:30:00.000Z"),
  ], { until: "2026-09-30T12:00:00.000Z" });
  assert.equal(labourHours(segs), 2.5, "1 h + 1 h + 0.5 h testing; 2 h waiting excluded");
  assert.ok(segs.every((s) => !s.open), "a completed job has no open time");
});

test("an unfinished job keeps running to now, and ignores out-of-order events", () => {
  const segs = buildSegments([
    ev(1, "jose", "pause", "2026-09-30T05:00:00.000Z"),
    ev(1, "jose", "start", "2026-09-30T06:00:00.000Z"),
  ], { until: "2026-09-30T08:00:00.000Z" });
  assert.equal(segs.length, 1);
  assert.equal(segs[0].open, true);
  assert.equal(labourHours(segs), 2);
});

test("a night shift across midnight is one report and one timesheet row per job", () => {
  const segs = buildSegments([
    ev(7, "jose", "start", "2026-09-29T20:00:00.000Z"), // 22:00 local
    ev(7, "jose", "pause", "2026-09-29T23:00:00.000Z"),
    ev(8, "jose", "start", "2026-09-29T23:00:00.000Z"),
    ev(8, "jose", "waiting_ops", "2026-09-30T01:00:00.000Z"),
    ev(7, "jose", "resume", "2026-09-30T02:00:00.000Z"),
    ev(7, "jose", "complete", "2026-09-30T03:30:00.000Z"), // 05:30 local
  ], { until: "2026-09-30T04:00:00.000Z" });
  const wos = new Map([
    [7, { asset_code: "T01AM", job: "Hydraulic hose", source: "breakdown" }],
    [8, { asset_code: "W200AM", job: "250 h service", source: "service" }],
  ]);
  const timeline = buildShiftTimeline(segs, { from: "2026-09-29T19:30:00.000Z", to: "2026-09-30T04:00:00.000Z", workOrders: wos });
  assert.equal(timeline.filter((t) => t.labour).reduce((s, t) => s + t.hours, 0), 6.5);
  const rows = shiftTimesheetRows(timeline, { technicianName: "Jose M", day: "2026-09-29" });
  assert.deepEqual(rows.map((r) => [r.asset_code, r.hours, r.time_started, r.time_finished, r.category]), [
    ["T01AM", 4.5, "22:00", "05:30", "Breakdown"],
    ["W200AM", 2, "01:00", "03:00", "Service"],
  ]);
  assert.ok(rows.every((r) => r.work_date === "2026-09-29" && r.reason.startsWith("WO #")));
});

test("a report left open into the next day puts that day's work on its own date", () => {
  const segs = buildSegments([
    ev(9, "jose", "start", "2026-09-29T05:00:00.000Z"), // 07:00 local, day 1
    ev(9, "jose", "pause", "2026-09-29T08:00:00.000Z"),
    ev(9, "jose", "resume", "2026-09-30T06:00:00.000Z"), // next morning, report never submitted
    ev(9, "jose", "complete", "2026-09-30T08:00:00.000Z"),
  ], { until: "2026-09-30T09:00:00.000Z" });
  const wos = new Map([[9, { asset_code: "L05AM", job: "Bucket teeth", source: "manual" }]]);
  const timeline = buildShiftTimeline(segs, { from: "2026-09-29T04:30:00.000Z", to: "2026-09-30T09:00:00.000Z", workOrders: wos });
  const rows = shiftTimesheetRows(timeline, { technicianName: "Jose M", day: "2026-09-29", shiftStart: "2026-09-29T04:30:00.000Z" });
  assert.deepEqual(rows.map((r) => [r.work_date, r.hours, r.category]), [["2026-09-29", 3, "Maintenance"], ["2026-09-30", 2, "Maintenance"]]);
});

test("shift reports open for more than 12 hours are listed for the foreman", async () => {
  const { overdueShifts } = await import("../utils/techShift.js");
  const Database = (await import("better-sqlite3")).default;
  const mem = new Database(":memory:");
  mem.exec(`
    CREATE TABLE tech_shifts (id INTEGER PRIMARY KEY, username TEXT, site_code TEXT, started_at TEXT, status TEXT);
    CREATE TABLE users (username TEXT, full_name TEXT);
    INSERT INTO users VALUES ('jose', 'Jose Mabunda');
    INSERT INTO tech_shifts VALUES
      (1, 'jose', 'main', '2026-09-29T05:00:00.000Z', 'open'),
      (2, 'maria', 'main', '2026-09-30T05:00:00.000Z', 'open'),
      (3, 'pedro', 'main', '2026-09-28T05:00:00.000Z', 'submitted'),
      (4, 'ana', 'other', '2026-09-28T05:00:00.000Z', 'open');
  `);
  const rows = overdueShifts(mem, { site: "main", now: new Date("2026-09-30T09:00:00.000Z") });
  assert.deepEqual(rows.map((r) => [r.name, r.hours_open]), [["Jose Mabunda", 28]]);
});

test("offline timestamps are accepted only within a sane window", () => {
  const now = new Date("2026-09-30T10:00:00.000Z");
  assert.equal(eventTime("2026-09-30T08:00:00.000Z", now), "2026-09-30T08:00:00.000Z");
  assert.equal(eventTime("2026-10-02T08:00:00.000Z", now), now.toISOString(), "future is refused");
  assert.equal(eventTime("2026-09-20T08:00:00.000Z", now), now.toISOString(), "older than 72 h is refused");
  assert.equal(eventTime("nonsense", now), now.toISOString());
});

// ---------------------------------------------------------------- API flow

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-tech-portal-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_ROOT = tempDir;

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { default: authRoutes } = await import("../routes/auth.routes.js");
const { default: workOrderRoutes } = await import("../routes/workorders.routes.js");
const { default: breakdownRoutes } = await import("../routes/breakdowns.routes.js");
const { default: maintenanceRoutes } = await import("../routes/maintenance.routes.js");
const { default: techRoutes } = await import("../routes/tech.routes.js");
const app = Fastify({ logger: false });
await app.register(authRoutes, { prefix: "/api/auth" });
await app.register(workOrderRoutes, { prefix: "/api/workorders" });
await app.register(breakdownRoutes, { prefix: "/api/breakdowns" });
await app.register(maintenanceRoutes, { prefix: "/api/maintenance" });
await app.register(techRoutes, { prefix: "/api/tech" });
await app.ready();
// Some routes finish background writes after a response; clean up when the process ends.
process.on("exit", () => {
  try { db.close(); } catch { /* already closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

db.exec(`
  INSERT INTO users (username, full_name, role, active) VALUES
    ('jose', 'Jose Mabunda', 'artisan', 1),
    ('maria', 'Maria Chissano', 'artisan', 1),
    ('pedro', 'Pedro Other', 'artisan', 1),
    ('foreman', 'Foreman', 'plant_manager', 1);
  INSERT INTO assets (id, asset_code, asset_name, category, active) VALUES (1, 'T01AM', 'Tipper 01', 'tipper_truck', 1);
  INSERT INTO breakdowns (id, asset_id, breakdown_date, status, description, component, critical)
    VALUES (1, 1, '2026-09-30', 'OPEN', 'Hose burst', 'Hydraulics', 1);
  INSERT INTO work_orders (id, asset_id, source, reference_id, status, opened_at, assigned_artisan_name, assigned_at)
    VALUES (1, 1, 'breakdown', 1, 'assigned', datetime('now'), 'Jose Mabunda', datetime('now'));
  INSERT INTO parts (id, part_code, part_name) VALUES (1, 'HOSE-12', 'Hydraulic hose 1/2in');
  INSERT INTO stock_movements (part_id, quantity, movement_type, reference) VALUES (1, 4, 'in', 'grn:1');
`);

const as = (user, role = "artisan") => ({ "x-user-name": user, "x-user-role": role, "x-user-roles": role, "x-site-code": "main" });
const call = async (user, method, url, payload, role) => {
  const res = await app.inject({ method, url: `/api/tech${url}`, headers: as(user, role), payload });
  return { code: res.statusCode, body: res.json() };
};
const act = (user, action, extra = {}) => call(user, "POST", "/workorders/1/action", { action, ...extra });

test("technician flow: today, start, finding, part, wait, resume, complete, shift", async (t) => {
  // Operators and stores cannot use the portal; another artisan cannot see the job.
  assert.equal((await call("op", "GET", "/today", undefined, "operator")).code, 403);
  assert.equal((await call("pedro", "GET", "/workorders/1")).code, 403);

  const t0 = await call("jose", "GET", "/today");
  assert.equal(t0.code, 200, JSON.stringify(t0.body));
  let today = t0.body;
  assert.equal(today.counts.urgent, 1, "critical breakdown is urgent");
  assert.equal(today.groups.urgent[0].job, "Hydraulics — Hose burst");
  assert.equal(today.groups.urgent[0].role, "lead");

  assert.equal((await call("jose", "POST", "/shift/start", { client_event_id: "s1", at: new Date(Date.now() - 4 * 3600000).toISOString() })).code, 200);

  // Start moves the WO to in progress through the normal status route.
  const start = await act("jose", "start", { client_event_id: "e1", at: new Date(Date.now() - 3 * 3600000).toISOString() });
  assert.equal(start.code, 200, JSON.stringify(start.body));
  assert.equal(db.prepare("SELECT status FROM work_orders WHERE id = 1").get().status, "in_progress");

  // The same offline event sent twice is saved once.
  const again = await act("jose", "start", { client_event_id: "e1" });
  assert.equal(again.body.duplicate, true);
  assert.equal(db.prepare("SELECT COUNT(*) AS n FROM tech_activity_events WHERE action = 'start'").get().n, 1);

  // A helper added by the foreman logs their own time; artisans cannot add helpers.
  assert.equal((await call("jose", "POST", "/workorders/1/helpers", { username: "maria" })).code, 403);
  assert.equal((await call("foreman", "POST", "/workorders/1/helpers", { username: "nobody" }, "plant_manager")).code, 400, "unknown user");
  assert.equal((await call("foreman", "POST", "/workorders/1/helpers", { username: "jose" }, "plant_manager")).code, 400, "the lead is not a helper");
  assert.equal((await call("foreman", "POST", "/workorders/1/helpers", { username: "Maria Chissano" }, "plant_manager")).code, 200, "by full name");
  assert.equal((await act("maria", "start", { at: new Date(Date.now() - 2 * 3600000).toISOString() })).code, 200);

  const finding = await call("jose", "POST", "/findings", { work_order_id: 1, text: "Hose chafed on chassis clamp", client_event_id: "f1" });
  assert.equal(finding.code, 200);
  assert.equal(db.prepare("SELECT repair_progress FROM work_orders WHERE id = 1").get().repair_progress, "Hose chafed on chassis clamp");

  const search = (await call("jose", "GET", "/parts/search?q=HOSE")).body;
  assert.equal(search.rows[0].on_hand, 4);

  const reqPart = await call("jose", "POST", "/workorders/1/parts-request", { part_code: "HOSE-12", qty: 1, client_event_id: "p1" });
  assert.equal(reqPart.code, 200, JSON.stringify(reqPart.body));
  const pr = db.prepare("SELECT part_name, work_order_id, status, requested_by FROM maintenance_parts_requests").get();
  assert.deepEqual(pr, { part_name: "Hydraulic hose 1/2in", work_order_id: 1, status: "requested", requested_by: "jose" });
  // Technicians request; they never issue stock.
  assert.equal(db.prepare("SELECT SUM(quantity) AS q FROM stock_movements").get().q, 4);

  assert.equal((await act("jose", "waiting_parts", { note: "Hose from stores", at: new Date(Date.now() - 90 * 60000).toISOString() })).code, 200);
  today = (await call("jose", "GET", "/today")).body;
  assert.equal(today.counts.waiting, 1);
  assert.equal((await act("jose", "resume", { at: new Date(Date.now() - 60 * 60000).toISOString() })).code, 200);

  // Rules: cannot complete without saying what was done; a helper cannot complete the WO.
  assert.equal((await act("jose", "complete", {})).code, 400);
  assert.equal((await act("jose", "pause", {})).code, 200);
  assert.equal((await act("jose", "pause", {})).code, 409, "already paused");
  assert.equal((await act("jose", "resume", {})).code, 200);

  const board = (await call("foreman", "GET", "/workorders/1/helpers", undefined, "plant_manager")).body;
  assert.equal(board.can_edit, true);
  assert.deepEqual(board.helpers.map((h) => [h.username, h.name, h.state]), [["maria", "Maria Chissano", "active"]]);
  assert.equal((await call("jose", "GET", "/workorders/1/helpers")).body.can_edit, false);

  const view = (await call("jose", "GET", "/workorders/1")).body;
  assert.equal(view.me.role, "lead");
  assert.equal(view.team.helpers[0].username, "maria");
  assert.equal(view.parts.requests.length, 1);
  assert.ok(view.time.total_hours > 3, `total ${view.time.total_hours}`);

  const done = await act("jose", "complete", { completion_notes: "Replaced hose and clamp, tested", client_event_id: "c1" });
  assert.equal(done.code, 200, JSON.stringify(done.body));
  const wo = db.prepare("SELECT status, labor_hours, completion_notes FROM work_orders WHERE id = 1").get();
  assert.equal(wo.status, "completed");
  assert.equal(wo.completion_notes, "Replaced hose and clamp, tested");
  // Jose 2.5 h (3 h minus 30 min waiting for parts) + Maria 2 h (stopped by the lead's completion).
  assert.ok(Math.abs(wo.labor_hours - 4.5) < 0.05, `labour ${wo.labor_hours}`);
  assert.equal(db.prepare("SELECT UPPER(status) AS s FROM breakdowns WHERE id = 1").get().s, "CLOSED");

  // Shift report: timeline from activity, handover, then timesheet rows.
  const shift = (await call("jose", "PUT", "/shift", { handover: "T01AM back in service", safety: "none" })).body.shift;
  assert.ok(shift.timeline.some((x) => x.work_order_id === 1 && x.labour));
  const submit = await call("jose", "POST", "/shift/submit", { client_event_id: "sub1" });
  assert.equal(submit.code, 200, JSON.stringify(submit.body));
  assert.equal(submit.body.timesheet_rows, 1);
  const ts = db.prepare("SELECT technician_name, asset_code, hours, job_card_no, created_by FROM mechanic_labor_entries").get();
  assert.equal(ts.technician_name, "Jose Mabunda");
  assert.equal(ts.asset_code, "T01AM");
  assert.equal(ts.job_card_no, "1");
  assert.equal(ts.created_by, "portal:jose");
  assert.ok(Math.abs(ts.hours - 2.5) < 0.05, `timesheet ${ts.hours}`);
  // Submitting twice (offline retry) does not double the timesheet.
  assert.equal((await call("jose", "POST", "/shift/submit", { client_event_id: "sub1" })).body.duplicate, true);
  assert.equal(db.prepare("SELECT COUNT(*) AS n FROM mechanic_labor_entries").get().n, 1);

  const list = (await call("foreman", "GET", "/shifts", undefined, "plant_manager")).body;
  assert.equal(list.rows[0].handover, "T01AM back in service");
  // Maria never tapped Start shift: her first job opened one.
  assert.equal(db.prepare("SELECT COUNT(*) AS n FROM tech_shifts WHERE username = 'maria' AND status = 'open'").get().n, 1);
  assert.equal((await call("pedro", "GET", "/shifts")).body.rows.length, 0, "technicians only see their own reports");

  const week = (await call("jose", "GET", "/week")).body;
  assert.equal(week.jobs_completed, 1);
  assert.equal(week.shifts_submitted, 1);

  const asset = (await call("jose", "GET", "/assets/t01am")).body;
  assert.equal(asset.asset.asset_code, "T01AM");
  assert.equal(asset.history[0].id, 1);
});

test("edge cases: reassignment, failed request retry, helper finishing, abandoned job", async (t) => {
  t.after(async () => {
    await app.close();
  });
  db.exec(`
    INSERT INTO work_orders (id, asset_id, source, status, opened_at, assigned_artisan_name, assigned_at, job_description)
    VALUES (2, 1, 'manual', 'assigned', datetime('now'), 'pedro', datetime('now'), 'Check brakes');
  `);
  const a2 = (user, action, extra = {}) => call(user, "POST", "/workorders/2/action", { action, ...extra });
  assert.equal((await a2("pedro", "start", { at: new Date(Date.now() - 60 * 60000).toISOString() })).code, 200);

  // A stores request with nothing to request fails, and the same offline id may be retried.
  const bad = await call("pedro", "POST", "/workorders/2/parts-request", { qty: 1, client_event_id: "pr-x" });
  assert.equal(bad.code, 400);
  const good = await call("pedro", "POST", "/workorders/2/parts-request", { part_name: "Brake pads", qty: 2, client_event_id: "pr-x" });
  assert.equal(good.code, 200, JSON.stringify(good.body));

  // Helper finishes their own part; the work order stays open.
  await call("foreman", "POST", "/workorders/2/helpers", { username: "maria" }, "plant_manager");
  assert.equal((await a2("maria", "start")).code, 200);
  assert.equal((await a2("maria", "complete", { note: "Adjusted rear brakes" })).code, 200);
  assert.equal(db.prepare("SELECT status FROM work_orders WHERE id = 2").get().status, "in_progress");

  // Reassigned to Jose: Pedro loses the job, Jose leads it.
  db.prepare("UPDATE work_orders SET assigned_artisan_name = 'jose' WHERE id = 2").run();
  assert.equal((await call("pedro", "GET", "/workorders/2")).code, 403);
  assert.equal((await call("jose", "GET", "/workorders/2")).body.me.role, "lead");

  // Pick up an unassigned job: first technician wins, the board records who took it.
  db.exec(`INSERT INTO work_orders (id, asset_id, source, status, opened_at, job_description) VALUES (3, 1, 'manual', 'open', datetime('now'), 'Grease pins')`);
  const avail = (await call("maria", "GET", "/available")).body.rows;
  assert.deepEqual(avail.map((r) => r.id), [3], "only open, unassigned jobs");
  assert.equal((await call("maria", "GET", "/today")).body.available_count, 1);
  const claim = await call("maria", "POST", "/workorders/3/claim", {});
  assert.equal(claim.code, 200, JSON.stringify(claim.body));
  const taken = await call("pedro", "POST", "/workorders/3/claim", {});
  assert.equal(taken.code, 409, "second technician is refused");
  const w3 = db.prepare("SELECT status, assigned_artisan_name, assigned_by, assigned_at FROM work_orders WHERE id = 3").get();
  assert.equal(w3.status, "assigned");
  assert.equal(w3.assigned_artisan_name, "maria");
  assert.equal(w3.assigned_by, "maria");
  assert.ok(w3.assigned_at);
  assert.equal((await call("maria", "GET", "/available")).body.rows.length, 0);
  assert.equal((await call("maria", "POST", "/workorders/3/action", { action: "start" })).code, 200, "she can start it now");
  assert.equal(db.prepare("SELECT status FROM work_orders WHERE id = 3").get().status, "in_progress");
  assert.equal((await call("maria", "POST", "/workorders/1/claim", {})).code, 409, "a finished or assigned job cannot be taken");
  assert.equal((await call("op", "POST", "/workorders/3/claim", {}, "operator")).code, 403);

  // Pedro's job was left running: ending his shift stops the clock.
  const sub = await call("pedro", "POST", "/shift/submit", {});
  assert.equal(sub.code, 200, JSON.stringify(sub.body));
  assert.deepEqual(sub.body.paused_jobs, [2]);
  const last = db.prepare("SELECT action, note FROM tech_activity_events WHERE work_order_id = 2 AND username = 'pedro' ORDER BY id DESC LIMIT 1").get();
  assert.deepEqual(last, { action: "pause", note: "Shift ended" });
});
