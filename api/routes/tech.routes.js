// IRONLOG/api/routes/tech.routes.js — technician portal API (/api/tech).
//
// A thin layer over the existing systems. Reads compose work orders, stock,
// parts requests, breakdowns, services and history. Writes that already have a
// route (work order status, parts requests, breakdowns) go through that route
// via app.inject, so its rules and permissions apply unchanged. What is new is
// stored in the tech_* tables: activity events, findings, photos, helpers and
// shift reports.
//
// Offline-capable writes carry client_event_id; a retried submission returns the
// first result instead of saving twice (tech_client_events).

import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import multipart from "@fastify/multipart";
import { db } from "../db/client.js";
import { getRoles, getSiteCode } from "../utils/request.js";
import { getAssetCurrentHoursInfo } from "../utils/assetMeterHours.js";
import { normalizeUploadedPhoto } from "../utils/imagePdf.js";
import { getDataRoot } from "../utils/storagePaths.js";
import { writeAudit } from "../utils/audit.js";
import { technicianMatchesUser } from "../utils/technicianIdentity.js";
import {
  STATE_LABELS,
  RUNNING_STATES,
  TECH_ACTIONS,
  buildSegments,
  ensureTechSchema,
  eventTime,
  labourHours,
  localDay,
  localDayStartIso,
  nextState,
  segmentHours,
} from "../utils/techActivity.js";
import {
  eventsByUser,
  eventsFor,
  hasColumn,
  hasTable,
  isUrgent,
  jobLine,
  myStates,
  myWorkOrders,
  stockInfo,
  workOrderLabourHours,
  workOrderParts,
} from "../utils/techPortal.js";
import { SHIFT_REMIND_HOURS, buildShiftTimeline, shiftTimesheetRows } from "../utils/techShift.js";

const PORTAL_ROLES = ["artisan", "supervisor", "workshop_admin", "admin", "plant_manager", "site_manager"];
const LEAD_ROLES = ["supervisor", "workshop_admin", "admin", "plant_manager", "site_manager"];

function userOf(req) {
  return String(req.headers["x-user-name"] || "").trim();
}

function roleOk(req, list) {
  return getRoles(req, { fallback: "admin" }).some((r) => list.includes(r));
}

function fullName(username) {
  if (!hasTable(db, "users")) return username;
  const r = db.prepare(`SELECT full_name FROM users WHERE LOWER(username) = LOWER(?)`).get(username);
  return String(r?.full_name || "").trim() || username;
}

/** Headers that let an internal call act as the same signed-in user. */
function forwardHeaders(req) {
  const out = { "content-type": "application/json" };
  for (const h of ["authorization", "x-user-name", "x-user-role", "x-user-roles", "x-site-code", "x-user-permissions"]) {
    if (req.headers[h]) out[h] = req.headers[h];
  }
  return out;
}

/** Runs fn once per client_event_id; later calls with the same id get the first result. */
async function once(req, kind, fn) {
  const id = String(req.body?.client_event_id || "").trim().slice(0, 80);
  if (!id) return fn();
  const prior = db.prepare(`SELECT result_json FROM tech_client_events WHERE client_event_id = ?`).get(id);
  if (prior) return { ...(prior.result_json ? JSON.parse(prior.result_json) : { ok: true }), duplicate: true };
  db.prepare(`INSERT OR IGNORE INTO tech_client_events (client_event_id, username, kind) VALUES (?, ?, ?)`).run(id, userOf(req), kind);
  try {
    const result = await fn();
    if (result?.ok !== false) {
      db.prepare(`UPDATE tech_client_events SET result_json = ? WHERE client_event_id = ?`).run(JSON.stringify(result), id);
    } else {
      // A refused action may be retried after the problem is fixed.
      db.prepare(`DELETE FROM tech_client_events WHERE client_event_id = ?`).run(id);
    }
    return result;
  } catch (err) {
    db.prepare(`DELETE FROM tech_client_events WHERE client_event_id = ?`).run(id);
    throw err;
  }
}

export default async function techRoutes(app) {
  ensureTechSchema(db);
  await app.register(multipart, { limits: { fileSize: 15 * 1024 * 1024, files: 1 } });
  const photoDir = path.join(getDataRoot(), "uploads", "work-order-photos");
  fs.mkdirSync(photoDir, { recursive: true });

  app.addHook("preHandler", async (req, reply) => {
    if (!roleOk(req, PORTAL_ROLES)) return reply.code(403).send({ ok: false, error: "The technician portal is for workshop staff." });
  });

  /** Lead, helper, or a workshop lead role; null when the job is not theirs. */
  function accessTo(req, woId) {
    const wo = db.prepare(`SELECT id, asset_id, status, source, reference_id, assigned_artisan_name, labor_hours FROM work_orders WHERE id = ?`).get(Number(woId));
    if (!wo) return { error: 404 };
    const me = userOf(req);
    const lead = technicianMatchesUser(db, wo.assigned_artisan_name, me);
    const helper = !lead && Boolean(db.prepare(`SELECT 1 FROM work_order_technicians WHERE work_order_id = ? AND LOWER(username) = LOWER(?)`).get(wo.id, me));
    const manager = roleOk(req, LEAD_ROLES);
    if (!lead && !helper && !manager) return { error: 403 };
    return { wo, lead, helper, manager };
  }

  async function changeStatus(req, woId, payload) {
    const res = await app.inject({ method: "POST", url: `/api/workorders/${woId}/status`, headers: forwardHeaders(req), payload });
    const body = res.json();
    return res.statusCode >= 400 ? { ok: false, status: res.statusCode, error: body?.error || "Status change refused" } : { ok: true, ...body };
  }

  function insertEvent({ woId, username, action, at, note = null, auto = 0, clientEventId = null }) {
    db.prepare(`
      INSERT INTO tech_activity_events (work_order_id, username, action, at, note, auto, client_event_id)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `).run(Number(woId), username, action, at, note, auto ? 1 : 0, clientEventId);
  }

  /** Pauses any other job this technician is still running (one running job at a time). */
  function pauseOtherJobs(username, exceptWoId, at) {
    const since = new Date(Date.parse(at) - 14 * 86400000).toISOString();
    const recent = eventsByUser(db, username, since);
    const woIds = [...new Set(recent.map((e) => e.work_order_id))].filter((id) => id !== Number(exceptWoId));
    const { byWo } = myStates(db, username, woIds);
    const paused = [];
    for (const [id, st] of byWo) {
      if (RUNNING_STATES.has(st.mine)) {
        insertEvent({ woId: id, username, action: "pause", at, note: "Paused: started another job", auto: 1 });
        paused.push(id);
      }
    }
    return paused;
  }

  // ---------------------------------------------------------------- Today
  app.get("/today", async (req) => {
    const me = userOf(req);
    const today = localDay();
    const dayStart = localDayStartIso(today);
    const rows = myWorkOrders(db, me, { finishedSince: dayStart });
    const ids = rows.map((r) => r.id);
    const { byWo, events } = myStates(db, me, ids);
    const nowIso = new Date().toISOString();
    const mySegs = buildSegments(eventsByUser(db, me, new Date(Date.now() - 14 * 86400000).toISOString()), { until: nowIso });
    const cards = rows.map((w) => {
      const parts = workOrderParts(db, w.id);
      const st = byWo.get(w.id)?.mine || "idle";
      const finished = ["completed", "approved", "closed"].includes(String(w.status || "").toLowerCase());
      const myToday = labourHours(mySegs.filter((s) => s.work_order_id === w.id), { from: dayStart });
      return {
        id: w.id,
        asset_code: w.asset_code,
        asset_name: w.asset_name,
        job: jobLine(w),
        source: w.source,
        status: w.status,
        priority: w.priority || null,
        critical: Boolean(w.breakdown_critical),
        due_date: w.due_date || w.ets_repair_date || null,
        role: w.role,
        my_state: finished ? "done" : st,
        my_state_label: STATE_LABELS[finished ? "done" : st],
        my_hours_today: myToday,
        service: w.source === "service" ? { name: w.service_name, interval: w.interval_hours } : null,
        parts: parts.summary,
        progress: w.repair_progress || null,
        done_at: w.done_at || null,
        _finished: finished,
        _urgent: isUrgent(w, today),
      };
    });
    const groups = { urgent: [], planned: [], waiting: [], completed: [] };
    for (const c of cards) {
      if (c._finished || c.my_state === "done") groups.completed.push(c);
      else if (["waiting_parts", "waiting_ops"].includes(c.my_state) || (c.parts.waiting && c.my_state !== "active" && c.my_state !== "testing")) groups.waiting.push(c);
      else if (c._urgent) groups.urgent.push(c);
      else groups.planned.push(c);
      delete c._finished;
      delete c._urgent;
    }
    const running = mySegs.find((s) => s.open && RUNNING_STATES.has(s.state));
    const currentCard = running ? cards.find((c) => c.id === running.work_order_id) : null;
    const shift = db.prepare(`SELECT id, started_at, status FROM tech_shifts WHERE LOWER(username) = LOWER(?) AND status = 'open' ORDER BY id DESC LIMIT 1`).get(me) || null;
    const lastSubmitted = db.prepare(`SELECT id, started_at, submitted_at FROM tech_shifts WHERE LOWER(username) = LOWER(?) AND status = 'submitted' ORDER BY id DESC LIMIT 1`).get(me) || null;
    // What changed for me: newly assigned jobs and parts that arrived.
    const notifications = [];
    for (const w of rows) {
      if (w.assigned_at && Date.parse(`${String(w.assigned_at).replace(" ", "T")}Z`) > Date.now() - 24 * 3600000 && w.role === "lead" && !["completed", "approved", "closed"].includes(String(w.status).toLowerCase())) {
        notifications.push({ kind: "assigned", wo_id: w.id, asset_code: w.asset_code, job: jobLine(w), text: `New job: WO #${w.id} ${w.asset_code} — ${jobLine(w)}` });
      }
    }
    if (ids.length && hasTable(db, "maintenance_parts_requests")) {
      const marks = ids.map(() => "?").join(", ");
      for (const r of db.prepare(`
        SELECT work_order_id, part_name, part_code, updated_at FROM maintenance_parts_requests
        WHERE work_order_id IN (${marks}) AND LOWER(status) = 'received' AND updated_at >= ?
      `).all(...ids, new Date(Date.now() - 48 * 3600000).toISOString())) {
        notifications.push({ kind: "parts", wo_id: r.work_order_id, part: r.part_name || r.part_code, text: `Part arrived for WO #${r.work_order_id}: ${r.part_name || r.part_code}` });
      }
    }
    // A shift open for longer than a working day was probably never submitted.
    const shiftHoursOpen = shift ? (Date.now() - Date.parse(shift.started_at)) / 3600000 : 0;
    return {
      ok: true,
      user: { username: me, name: fullName(me) },
      day: today,
      shift: shift ? { ...shift, hours_open: Number(shiftHoursOpen.toFixed(1)), overdue: shiftHoursOpen >= SHIFT_REMIND_HOURS } : null,
      last_submitted_shift: lastSubmitted,
      current: currentCard
        ? { ...currentCard, state: running.state, since: running.start, running_minutes: Math.round(segmentHours({ start: running.start, end: nowIso }) * 60) }
        : null,
      groups,
      counts: {
        urgent: groups.urgent.length,
        planned: groups.planned.length,
        waiting: groups.waiting.length,
        completed: groups.completed.length,
        waiting_parts: cards.filter((c) => c.parts.waiting).length,
      },
      hours_today: labourHours(mySegs, { from: dayStart }),
      notifications: notifications.slice(0, 6),
      events_seen: events.length,
    };
  });

  // ---------------------------------------------------------------- Work order
  app.get("/workorders/:id", async (req, reply) => {
    const acc = accessTo(req, req.params.id);
    if (acc.error) return reply.code(acc.error).send({ ok: false, error: acc.error === 404 ? "Work order not found" : "This job is not assigned to you" });
    const id = acc.wo.id;
    const me = userOf(req);
    const w = db.prepare(`
      SELECT w.*, a.asset_code, a.asset_name, a.category,
        b.description AS breakdown_description, b.component AS breakdown_component, b.critical AS breakdown_critical,
        b.breakdown_date, b.start_at AS breakdown_start, b.parts_status, b.ets_repair_date,
        mp.service_name, mp.interval_hours, mp.last_service_hours
      FROM work_orders w JOIN assets a ON a.id = w.asset_id
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
      LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id AND w.source = 'service'
      WHERE w.id = ?
    `).get(id);
    const meter = getAssetCurrentHoursInfo(w.asset_id);
    const events = eventsFor(db, [id]);
    const nowIso = new Date().toISOString();
    const segs = buildSegments(events, { until: nowIso });
    const states = {};
    for (const s of segs) states[s.username] = s.open ? s.state : states[s.username] || "done";
    for (const e of events) if (!(e.username in states)) states[e.username] = "idle";
    const myState = (() => {
      const st = new Map();
      for (const e of events.filter((x) => x.username.toLowerCase() === me.toLowerCase())) {
        const r = nextState(st.get("s") || "idle", e.action);
        if (!r.error) st.set("s", r.state);
      }
      return st.get("s") || "idle";
    })();
    const helpers = db.prepare(`SELECT username, added_by, added_at FROM work_order_technicians WHERE work_order_id = ? ORDER BY id`).all(id);
    const findings = db.prepare(`SELECT id, username, kind, text, at FROM tech_findings WHERE work_order_id = ? ORDER BY at DESC, id DESC`).all(id);
    const photos = db.prepare(`SELECT id, username, file_path, caption, at FROM work_order_photos WHERE work_order_id = ? ORDER BY at DESC, id DESC`).all(id)
      .map((p) => ({ ...p, url: `/${p.file_path}` }));
    const history = db.prepare(`
      SELECT w2.id, w2.source, w2.status, w2.opened_at, COALESCE(w2.completed_at, w2.closed_at) AS done_at,
        w2.completion_notes, w2.job_description, b2.component AS breakdown_component, b2.description AS breakdown_description,
        mp2.service_name
      FROM work_orders w2
      LEFT JOIN breakdowns b2 ON b2.id = w2.reference_id AND w2.source = 'breakdown'
      LEFT JOIN maintenance_plans mp2 ON mp2.id = w2.reference_id AND w2.source = 'service'
      WHERE w2.asset_id = ? AND w2.id <> ? AND LOWER(COALESCE(w2.status, '')) IN ('completed', 'approved', 'closed')
      ORDER BY COALESCE(w2.completed_at, w2.closed_at, w2.opened_at) DESC LIMIT 8
    `).all(w.asset_id, id).map((h) => ({ id: h.id, source: h.source, job: jobLine(h), done_at: h.done_at, notes: h.completion_notes || null }));
    const lead = String(w.assigned_artisan_name || "");
    const activity = [
      ...events.map((e) => ({ at: e.at, who: e.username, kind: "event", text: `${e.action.replace(/_/g, " ")}${e.note ? ` — ${e.note}` : ""}`, auto: Boolean(e.auto) })),
      ...findings.map((f) => ({ at: f.at, who: f.username, kind: f.kind, text: f.text })),
    ].sort((a, b) => String(b.at).localeCompare(String(a.at))).slice(0, 40);
    let service = null;
    if (w.source === "service" && w.interval_hours) {
      const nextDue = Number(w.last_service_hours || 0) + Number(w.interval_hours || 0);
      service = { name: w.service_name, interval: Number(w.interval_hours), next_due: nextDue, remaining: Number((nextDue - Number(meter.hours || 0)).toFixed(1)) };
    }
    return {
      ok: true,
      wo: {
        id: w.id,
        status: w.status,
        source: w.source,
        priority: w.priority,
        job: jobLine(w),
        job_description: w.job_description || null,
        progress: w.repair_progress || null,
        completion_notes: w.completion_notes || null,
        opened_at: w.opened_at,
        started_at: w.started_at,
        completed_at: w.completed_at,
        due_date: w.due_date || w.ets_repair_date || null,
        labor_hours: Number(w.labor_hours || 0),
      },
      asset: { id: w.asset_id, code: w.asset_code, name: w.asset_name, category: w.category, meter },
      breakdown: w.source === "breakdown"
        ? { id: w.reference_id, component: w.breakdown_component, description: w.breakdown_description, critical: Boolean(w.breakdown_critical), since: w.breakdown_start || w.breakdown_date, parts_status: w.parts_status, return_date: w.ets_repair_date }
        : null,
      service,
      team: { lead, lead_name: lead ? fullName(lead) : null, helpers: helpers.map((h) => ({ ...h, name: fullName(h.username) })), states },
      me: { username: me, role: acc.lead ? "lead" : acc.helper ? "helper" : "manager", state: myState, state_label: STATE_LABELS[myState] },
      time: {
        my_hours: labourHours(segs.filter((s) => s.username.toLowerCase() === me.toLowerCase())),
        total_hours: labourHours(segs),
        running_since: segs.find((s) => s.open && s.username.toLowerCase() === me.toLowerCase())?.start || null,
      },
      parts: workOrderParts(db, id),
      findings,
      photos,
      history,
      activity,
    };
  });

  // Start / pause / resume / waiting / testing / complete.
  app.post("/workorders/:id/action", async (req, reply) => {
    const action = String(req.body?.action || "").trim();
    if (!TECH_ACTIONS.includes(action)) return reply.code(400).send({ ok: false, error: `action must be one of ${TECH_ACTIONS.join(", ")}` });
    const result = await once(req, `action:${action}`, async () => {
      const acc = accessTo(req, req.params.id);
      if (acc.error) return { ok: false, status: acc.error, error: acc.error === 404 ? "Work order not found" : "This job is not assigned to you" };
      const me = userOf(req);
      const woId = acc.wo.id;
      const at = eventTime(req.body?.at);
      const note = String(req.body?.note || "").trim().slice(0, 500) || null;
      const events = eventsFor(db, [woId]).filter((e) => e.username.toLowerCase() === me.toLowerCase());
      let cur = "idle";
      for (const e of events) { const r = nextState(cur, e.action); if (!r.error) cur = r.state; }
      const next = nextState(cur, action);
      if (next.error) return { ok: false, status: 409, error: next.error };
      const woStatus = String(acc.wo.status || "").toLowerCase();
      const isLead = acc.lead || (acc.manager && !acc.helper);

      if (action === "start" || action === "resume") {
        if (["completed", "approved", "closed"].includes(woStatus)) return { ok: false, status: 409, error: "This job is already finished." };
        if (woStatus === "open") return { ok: false, status: 409, error: "This job is not assigned yet. Ask your foreman to assign it." };
      }
      // The work order's own status only moves through the normal route.
      if (action === "start" && woStatus === "assigned") {
        if (!isLead) return { ok: false, status: 409, error: "The lead technician starts the job first." };
        const r = await changeStatus(req, woId, { status: "in_progress" });
        if (!r.ok) return r;
      }
      let statusResult = null;
      if (action === "complete" && isLead) {
        const notes = String(req.body?.completion_notes || note || "").trim();
        if (!notes) return { ok: false, status: 400, error: "Say what was done before completing the job." };
        // Everyone still working on it stops now.
        for (const [who, st] of Object.entries(stateMap(woId))) {
          if (who.toLowerCase() !== me.toLowerCase() && !["idle", "done"].includes(st)) {
            insertEvent({ woId, username: who, action: "complete", at, note: "Job completed by lead", auto: 1 });
          }
        }
        insertEvent({ woId, username: me, action, at, note, clientEventId: req.body?.client_event_id || null });
        // Labour from the timer, unless someone already entered hours on the job.
        const manual = Number(acc.wo.labor_hours || 0) > 0;
        const typed = Number(req.body?.labor_hours);
        const hours = Number.isFinite(typed) && typed > 0 ? typed : manual ? null : workOrderLabourHours(db, woId, at);
        const payload = { status: "completed", completion_notes: notes };
        if (hours != null && hours > 0) payload.labor_hours = hours;
        statusResult = await changeStatus(req, woId, payload);
        if (!statusResult.ok) {
          db.prepare(`DELETE FROM tech_activity_events WHERE work_order_id = ? AND at = ? AND action = 'complete'`).run(woId, at);
          return statusResult;
        }
      } else {
        if (action === "start" || action === "resume") {
          pauseOtherJobs(me, woId, at);
          // The first job of the day opens the shift when Start shift was not tapped.
          if (!db.prepare(`SELECT 1 FROM tech_shifts WHERE LOWER(username) = LOWER(?) AND status = 'open'`).get(me)) {
            db.prepare(`INSERT INTO tech_shifts (username, site_code, started_at) VALUES (?, ?, ?)`).run(me, getSiteCode(req), at);
          }
        }
        insertEvent({ woId, username: me, action, at, note, clientEventId: req.body?.client_event_id || null });
      }
      // Waiting reasons are the job's latest progress (daily report shows it).
      if (["waiting_parts", "waiting_ops"].includes(action) || (note && action !== "complete")) {
        const label = action === "waiting_parts" ? "Waiting for parts" : action === "waiting_ops" ? "Waiting for operations" : null;
        const text = [label, note].filter(Boolean).join(": ");
        if (text && hasColumn(db, "work_orders", "repair_progress")) {
          db.prepare(`UPDATE work_orders SET repair_progress = ?, repair_progress_at = datetime('now') WHERE id = ?`).run(text, woId);
        }
      }
      writeAudit(db, req, { module: "tech", action: `tech.${action}`, entity_type: "work_order", entity_id: String(woId), payload: { at, note } });
      return { ok: true, work_order_id: woId, action, state: next.state, state_label: STATE_LABELS[next.state], at, status_change: statusResult };
    });
    if (result?.ok === false) return reply.code(result.status || 400).send(result);
    return result;
  });

  function stateMap(woId) {
    const out = {};
    for (const e of eventsFor(db, [woId])) {
      const r = nextState(out[e.username] || "idle", e.action);
      if (!r.error) out[e.username] = r.state;
    }
    return out;
  }

  // Findings (also used from the asset view without a work order).
  app.post("/findings", async (req, reply) => {
    const result = await once(req, "finding", async () => {
      const text = String(req.body?.text || "").trim().slice(0, 2000);
      if (!text) return { ok: false, status: 400, error: "Write what you found." };
      const kind = ["finding", "safety", "unplanned", "note"].includes(req.body?.kind) ? req.body.kind : "finding";
      let woId = Number(req.body?.work_order_id || 0) || null;
      let assetId = Number(req.body?.asset_id || 0) || null;
      if (woId) {
        const acc = accessTo(req, woId);
        if (acc.error) return { ok: false, status: acc.error, error: "This job is not assigned to you" };
        assetId = acc.wo.asset_id;
      }
      const at = eventTime(req.body?.at);
      const ins = db.prepare(`INSERT INTO tech_findings (work_order_id, asset_id, username, kind, text, at) VALUES (?, ?, ?, ?, ?, ?)`)
        .run(woId, assetId, userOf(req), kind, text, at);
      if (woId && kind === "finding" && hasColumn(db, "work_orders", "repair_progress")) {
        db.prepare(`UPDATE work_orders SET repair_progress = ?, repair_progress_at = datetime('now') WHERE id = ?`).run(text.slice(0, 300), woId);
      }
      return { ok: true, id: Number(ins.lastInsertRowid), at };
    });
    if (result?.ok === false) return reply.code(result.status || 400).send(result);
    return result;
  });

  // Photos (multipart, field "file"; client_event_id and caption in the query string).
  app.post("/workorders/:id/photos", async (req, reply) => {
    const acc = accessTo(req, req.params.id);
    if (acc.error) return reply.code(acc.error).send({ ok: false, error: "This job is not assigned to you" });
    const clientId = String(req.query?.client_event_id || "").trim().slice(0, 80);
    if (clientId) {
      const prior = db.prepare(`SELECT result_json FROM tech_client_events WHERE client_event_id = ?`).get(clientId);
      if (prior?.result_json) return { ...JSON.parse(prior.result_json), duplicate: true };
    }
    const part = await req.file();
    if (!part) return reply.code(400).send({ ok: false, error: "Attach a photo" });
    let buf;
    try {
      buf = await normalizeUploadedPhoto(await part.toBuffer());
    } catch {
      return reply.code(400).send({ ok: false, error: "Could not read that image" });
    }
    const name = `wo_${acc.wo.id}_${Date.now()}_${crypto.randomBytes(4).toString("hex")}.jpg`;
    await fs.promises.writeFile(path.join(photoDir, name), buf);
    const rel = `uploads/work-order-photos/${name}`;
    const at = eventTime(req.query?.at);
    const caption = String(req.query?.caption || "").trim().slice(0, 200) || null;
    const ins = db.prepare(`INSERT INTO work_order_photos (work_order_id, username, file_path, caption, at) VALUES (?, ?, ?, ?, ?)`)
      .run(acc.wo.id, userOf(req), rel, caption, at);
    const result = { ok: true, id: Number(ins.lastInsertRowid), url: `/${rel}`, at };
    if (clientId) db.prepare(`INSERT OR REPLACE INTO tech_client_events (client_event_id, username, kind, result_json) VALUES (?, ?, 'photo', ?)`).run(clientId, userOf(req), JSON.stringify(result));
    return result;
  });

  // Request a part from stores (the normal parts-request route does the work).
  app.post("/workorders/:id/parts-request", async (req, reply) => {
    const result = await once(req, "parts_request", async () => {
      const acc = accessTo(req, req.params.id);
      if (acc.error) return { ok: false, status: acc.error, error: "This job is not assigned to you" };
      const res = await app.inject({
        method: "POST",
        url: "/api/maintenance/parts-requests",
        headers: forwardHeaders(req),
        payload: {
          work_order_id: acc.wo.id,
          asset_id: acc.wo.asset_id,
          part_code: String(req.body?.part_code || "").trim(),
          part_name: String(req.body?.part_name || "").trim(),
          qty: Number(req.body?.qty || 1),
          urgency: String(req.body?.urgency || "normal"),
          notes: String(req.body?.notes || "").trim(),
        },
      });
      const body = res.json();
      if (res.statusCode >= 400 || body?.ok === false) return { ok: false, status: res.statusCode >= 400 ? res.statusCode : 400, error: body?.error || "Stores request failed" };
      return { ok: true, id: body.id };
    });
    if (result?.ok === false) return reply.code(result.status || 400).send(result);
    return result;
  });

  // Store items for the part picker (code, name, on hand, bin).
  app.get("/parts/search", async (req) => {
    const q = String(req.query?.q || "").trim();
    if (q.length < 2) return { ok: true, rows: [] };
    const rows = db.prepare(`
      SELECT id, part_code, part_name FROM parts
      WHERE part_code LIKE ? OR part_name LIKE ?
      ORDER BY CASE WHEN part_code LIKE ? THEN 0 ELSE 1 END, part_code LIMIT 20
    `).all(`%${q}%`, `%${q}%`, `${q}%`);
    const stock = stockInfo(db, rows.map((r) => r.id));
    return { ok: true, rows: rows.map((r) => ({ part_code: r.part_code, part_name: r.part_name, on_hand: stock.get(r.id)?.on_hand ?? 0, bin: stock.get(r.id)?.bin || null })) };
  });

  // Helpers on a job: list (anyone who can see the job), add/remove (foreman / supervisor).
  app.get("/workorders/:id/helpers", async (req, reply) => {
    const acc = accessTo(req, req.params.id);
    if (acc.error) return reply.code(acc.error).send({ ok: false, error: acc.error === 404 ? "Work order not found" : "This job is not assigned to you" });
    const nowIso = new Date().toISOString();
    const segs = buildSegments(eventsFor(db, [acc.wo.id]), { until: nowIso });
    const states = stateMap(acc.wo.id);
    const helpers = db.prepare(`SELECT username, added_by, added_at FROM work_order_technicians WHERE work_order_id = ? ORDER BY id`).all(acc.wo.id)
      .map((h) => ({
        ...h,
        name: fullName(h.username),
        state: states[h.username] || "idle",
        state_label: STATE_LABELS[states[h.username] || "idle"],
        hours: labourHours(segs.filter((x) => x.username.toLowerCase() === h.username.toLowerCase())),
      }));
    return { ok: true, lead: acc.wo.assigned_artisan_name || null, helpers, can_edit: roleOk(req, LEAD_ROLES) };
  });

  app.post("/workorders/:id/helpers", async (req, reply) => {
    if (!roleOk(req, LEAD_ROLES)) return reply.code(403).send({ ok: false, error: "Only a foreman or supervisor adds helpers." });
    const wanted = String(req.body?.username || "").trim();
    if (!wanted) return reply.code(400).send({ ok: false, error: "username is required" });
    const wo = db.prepare(`SELECT id, assigned_artisan_name FROM work_orders WHERE id = ?`).get(Number(req.params.id));
    if (!wo) return reply.code(404).send({ ok: false, error: "Work order not found" });
    // A helper is a real user (picked by username or full name), so their time and jobs line up.
    const user = hasTable(db, "users")
      ? db.prepare(`SELECT username FROM users WHERE COALESCE(active, 1) = 1 AND (LOWER(username) = LOWER(?) OR LOWER(TRIM(full_name)) = LOWER(?)) LIMIT 1`).get(wanted, wanted)
      : null;
    if (!user) return reply.code(400).send({ ok: false, error: `No active user '${wanted}'. Add the technician in Workshop technicians first.` });
    const username = user.username;
    if (technicianMatchesUser(db, wo.assigned_artisan_name, username)) {
      return reply.code(400).send({ ok: false, error: `${fullName(username)} already leads this job.` });
    }
    db.prepare(`INSERT OR IGNORE INTO work_order_technicians (work_order_id, username, added_by) VALUES (?, ?, ?)`).run(wo.id, username, userOf(req));
    writeAudit(db, req, { module: "tech", action: "tech.helper_add", entity_type: "work_order", entity_id: String(wo.id), payload: { username } });
    return { ok: true };
  });

  app.delete("/workorders/:id/helpers/:username", async (req, reply) => {
    if (!roleOk(req, LEAD_ROLES)) return reply.code(403).send({ ok: false, error: "Only a foreman or supervisor removes helpers." });
    db.prepare(`DELETE FROM work_order_technicians WHERE work_order_id = ? AND LOWER(username) = LOWER(?)`).run(Number(req.params.id), String(req.params.username));
    return { ok: true };
  });

  // ---------------------------------------------------------------- Shift report
  function openShift(me) {
    return db.prepare(`SELECT * FROM tech_shifts WHERE LOWER(username) = LOWER(?) AND status = 'open' ORDER BY id DESC LIMIT 1`).get(me) || null;
  }

  function shiftView(shift, me) {
    const until = shift.ended_at || new Date().toISOString();
    const events = eventsByUser(db, me, new Date(Date.parse(shift.started_at) - 14 * 86400000).toISOString());
    const woIds = [...new Set(events.map((e) => e.work_order_id))];
    const wos = new Map(woIds.length ? db.prepare(`
      SELECT w.id, w.source, w.status, w.job_description, a.asset_code, a.asset_name,
        b.component AS breakdown_component, b.description AS breakdown_description, mp.service_name
      FROM work_orders w JOIN assets a ON a.id = w.asset_id
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
      LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id AND w.source = 'service'
      WHERE w.id IN (${woIds.map(() => "?").join(", ")})
    `).all(...woIds).map((w) => [w.id, { ...w, job: jobLine(w) }]) : []);
    const findings = db.prepare(`
      SELECT f.id, f.work_order_id, f.kind, f.text, f.at FROM tech_findings f
      WHERE LOWER(f.username) = LOWER(?) AND f.at >= ? AND f.at <= ? ORDER BY f.at
    `).all(me, shift.started_at, until);
    return {
      ...shift,
      timeline: buildShiftTimeline(buildSegments(events, { until }), { from: shift.started_at, to: until, workOrders: wos }),
      findings_logged: findings,
      hours: labourHours(buildSegments(events, { until }), { from: shift.started_at, to: until }),
    };
  }

  app.get("/shift", async (req) => {
    const me = userOf(req);
    const shift = openShift(me);
    return { ok: true, shift: shift ? shiftView(shift, me) : null };
  });

  app.post("/shift/start", async (req) => {
    return once(req, "shift_start", async () => {
      const me = userOf(req);
      const existing = openShift(me);
      if (existing) return { ok: true, shift_id: existing.id, already_open: true };
      const at = eventTime(req.body?.at);
      const ins = db.prepare(`INSERT INTO tech_shifts (username, site_code, started_at) VALUES (?, ?, ?)`).run(me, getSiteCode(req), at);
      return { ok: true, shift_id: Number(ins.lastInsertRowid), started_at: at };
    });
  });

  const SHIFT_FIELDS = ["findings", "unplanned", "safety", "outstanding", "handover", "next_shift"];

  app.put("/shift", async (req, reply) => {
    const me = userOf(req);
    const shift = openShift(me);
    if (!shift) return reply.code(404).send({ ok: false, error: "No open shift. Start your shift first." });
    const sets = [];
    const vals = [];
    for (const f of SHIFT_FIELDS) {
      if (Object.prototype.hasOwnProperty.call(req.body || {}, f)) {
        sets.push(`${f} = ?`);
        vals.push(String(req.body[f] ?? "").slice(0, 4000) || null);
      }
    }
    if (sets.length) db.prepare(`UPDATE tech_shifts SET ${sets.join(", ")} WHERE id = ?`).run(...vals, shift.id);
    return { ok: true, shift: shiftView(openShift(me), me) };
  });

  app.post("/shift/submit", async (req, reply) => {
    const result = await once(req, "shift_submit", async () => {
      const me = userOf(req);
      const shift = openShift(me);
      if (!shift) return { ok: false, status: 404, error: "No open shift to submit." };
      const at = eventTime(req.body?.at);
      for (const f of SHIFT_FIELDS) {
        if (Object.prototype.hasOwnProperty.call(req.body || {}, f)) {
          db.prepare(`UPDATE tech_shifts SET ${f} = ? WHERE id = ?`).run(String(req.body[f] ?? "").slice(0, 4000) || null, shift.id);
        }
      }
      // Jobs still running when the shift ends are paused, so time stops.
      const events = eventsByUser(db, me, new Date(Date.parse(shift.started_at) - 14 * 86400000).toISOString());
      const running = buildSegments(events, { until: at }).filter((s) => s.open && RUNNING_STATES.has(s.state));
      for (const s of running) insertEvent({ woId: s.work_order_id, username: me, action: "pause", at, note: "Shift ended", auto: 1 });
      db.prepare(`UPDATE tech_shifts SET ended_at = ?, status = 'submitted', submitted_at = ? WHERE id = ?`).run(at, at, shift.id);
      const view = shiftView({ ...shift, ended_at: at, status: "submitted" }, me);
      // The shift's job time goes to the mechanics timesheet (decided: job + timesheet).
      const rows = shiftTimesheetRows(view.timeline, { technicianName: fullName(me), day: localDay(shift.started_at), shiftStart: shift.started_at });
      let written = 0;
      if (rows.length && hasTable(db, "mechanic_labor_entries")) {
        // Same additive columns the timesheet screen adds on first use.
        for (const col of ["category", "time_started", "time_finished", "job_card_no"]) {
          if (!hasColumn(db, "mechanic_labor_entries", col)) db.prepare(`ALTER TABLE mechanic_labor_entries ADD COLUMN ${col} TEXT`).run();
        }
        const hasJobCard = hasColumn(db, "mechanic_labor_entries", "job_card_no");
        const hasTimes = hasColumn(db, "mechanic_labor_entries", "time_started");
        const hasCategory = hasColumn(db, "mechanic_labor_entries", "category");
        for (const r of rows) {
          const cols = ["work_date", "technician_name", "hours", "asset_code", "reason", "site_code", "created_by"];
          const vals = [r.work_date, r.technician_name, r.hours, r.asset_code, r.reason, getSiteCode(req), `portal:${me}`];
          if (hasJobCard) { cols.push("job_card_no"); vals.push(String(r.work_order_id)); }
          if (hasTimes) { cols.push("time_started", "time_finished"); vals.push(r.time_started, r.time_finished); }
          if (hasCategory) { cols.push("category"); vals.push(r.category); }
          db.prepare(`INSERT INTO mechanic_labor_entries (${cols.join(", ")}) VALUES (${cols.map(() => "?").join(", ")})`).run(...vals);
          written += 1;
        }
      }
      db.prepare(`UPDATE tech_shifts SET timesheet_rows = ? WHERE id = ?`).run(written, shift.id);
      writeAudit(db, req, { module: "tech", action: "tech.shift_submit", entity_type: "tech_shift", entity_id: String(shift.id), payload: { hours: view.hours, timesheet_rows: written } });
      return { ok: true, shift_id: shift.id, hours: view.hours, timesheet_rows: written, paused_jobs: running.map((s) => s.work_order_id) };
    });
    if (result?.ok === false) return reply.code(result.status || 400).send(result);
    return result;
  });

  // Submitted shift reports (foremen and supervisors review them).
  app.get("/shifts", async (req, reply) => {
    const manager = roleOk(req, LEAD_ROLES);
    const me = userOf(req);
    const days = Math.min(31, Math.max(1, Number(req.query?.days || 7)));
    const since = new Date(Date.now() - days * 86400000).toISOString();
    const rows = db.prepare(`
      SELECT * FROM tech_shifts WHERE status = 'submitted' AND started_at >= ?
      ${manager ? "" : "AND LOWER(username) = LOWER(?)"}
      ORDER BY started_at DESC LIMIT 100
    `).all(...(manager ? [since] : [since, me]));
    return { ok: true, rows: rows.map((s) => ({ ...shiftView(s, s.username), name: fullName(s.username) })) };
  });

  // ---------------------------------------------------------------- My week
  app.get("/week", async (req) => {
    const me = userOf(req);
    const today = localDay();
    const days = [];
    for (let i = 6; i >= 0; i -= 1) days.push(localDay(new Date(Date.now() - i * 86400000).toISOString()));
    const from = localDayStartIso(days[0]);
    const nowIso = new Date().toISOString();
    const segs = buildSegments(eventsByUser(db, me, new Date(Date.parse(from) - 14 * 86400000).toISOString()), { until: nowIso });
    const perDay = days.map((d, i) => ({
      day: d,
      hours: labourHours(segs, { from: localDayStartIso(d), to: i === days.length - 1 ? nowIso : localDayStartIso(days[i + 1]) }),
    }));
    const mine = myWorkOrders(db, me, { finishedSince: from });
    const { byWo } = myStates(db, me, mine.map((w) => w.id));
    const finished = mine.filter((w) => ["completed", "approved", "closed"].includes(String(w.status).toLowerCase()));
    const open = mine.filter((w) => !finished.includes(w));
    const waiting = open.filter((w) => ["waiting_parts", "waiting_ops"].includes(byWo.get(w.id)?.mine));
    const shifts = db.prepare(`SELECT id, status, started_at, submitted_at FROM tech_shifts WHERE LOWER(username) = LOWER(?) AND started_at >= ? ORDER BY started_at`).all(me, from);
    return {
      ok: true,
      days: perDay,
      hours: Number(perDay.reduce((s, d) => s + d.hours, 0).toFixed(2)),
      jobs_completed: finished.length,
      jobs_open: open.length,
      jobs_waiting: waiting.length,
      shifts_submitted: shifts.filter((s) => s.status === "submitted").length,
      shift_open: shifts.find((s) => s.status === "open") || null,
      today,
    };
  });

  // ---------------------------------------------------------------- Asset (QR)
  app.get("/assets/:code", async (req, reply) => {
    const a = db.prepare(`SELECT id, asset_code, asset_name, category FROM assets WHERE UPPER(asset_code) = UPPER(?)`).get(String(req.params.code || ""));
    if (!a) return reply.code(404).send({ ok: false, error: "Machine not found" });
    const me = userOf(req);
    const active = db.prepare(`
      SELECT w.id, w.source, w.status, w.assigned_artisan_name, w.priority, w.job_description, w.opened_at,
        b.component AS breakdown_component, b.description AS breakdown_description, mp.service_name
      FROM work_orders w
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
      LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id AND w.source = 'service'
      WHERE w.asset_id = ? AND LOWER(COALESCE(w.status, '')) NOT IN ('completed', 'approved', 'closed')
      ORDER BY w.id DESC
    `).all(a.id).map((w) => ({ id: w.id, source: w.source, status: w.status, job: jobLine(w), lead: w.assigned_artisan_name, mine: technicianMatchesUser(db, w.assigned_artisan_name, me) }));
    const breakdown = db.prepare(`
      SELECT id, component, description, breakdown_date, start_at, parts_status, ets_repair_date, critical FROM breakdowns
      WHERE asset_id = ? AND UPPER(TRIM(COALESCE(status, 'OPEN'))) <> 'CLOSED' ORDER BY id DESC LIMIT 1
    `).get(a.id) || null;
    const plans = db.prepare(`SELECT id, service_name, interval_hours, last_service_hours FROM maintenance_plans WHERE asset_id = ? AND active = 1`).all(a.id);
    const meter = getAssetCurrentHoursInfo(a.id);
    const next = plans
      .map((p) => ({ service: p.service_name, interval: Number(p.interval_hours), due_at: Number(p.last_service_hours || 0) + Number(p.interval_hours || 0) }))
      .map((p) => ({ ...p, remaining: Number((p.due_at - Number(meter.hours || 0)).toFixed(1)) }))
      .sort((x, y) => x.remaining - y.remaining)[0] || null;
    const history = db.prepare(`
      SELECT w.id, w.source, COALESCE(w.completed_at, w.closed_at) AS done_at, w.completion_notes, w.job_description,
        b.component AS breakdown_component, b.description AS breakdown_description, mp.service_name
      FROM work_orders w
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
      LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id AND w.source = 'service'
      WHERE w.asset_id = ? AND LOWER(COALESCE(w.status, '')) IN ('completed', 'approved', 'closed')
      ORDER BY COALESCE(w.completed_at, w.closed_at) DESC LIMIT 6
    `).all(a.id).map((h) => ({ id: h.id, job: jobLine(h), done_at: h.done_at, notes: h.completion_notes || null }));
    const findings = db.prepare(`SELECT id, username, kind, text, at FROM tech_findings WHERE asset_id = ? ORDER BY at DESC LIMIT 6`).all(a.id);
    return { ok: true, asset: { ...a, meter }, active, breakdown, next_service: next, history, findings };
  });

  // Report a breakdown from the portal (normal breakdown route; creates its WO).
  app.post("/breakdowns", async (req, reply) => {
    const result = await once(req, "breakdown", async () => {
      const res = await app.inject({
        method: "POST",
        url: "/api/breakdowns",
        headers: forwardHeaders(req),
        payload: {
          asset_code: String(req.body?.asset_code || "").trim(),
          breakdown_date: localDay(),
          component: String(req.body?.component || "").trim(),
          description: String(req.body?.description || "").trim(),
          critical: Boolean(req.body?.critical),
        },
      });
      const body = res.json();
      if (res.statusCode >= 400) return { ok: false, status: res.statusCode, error: body?.error || "Could not report the breakdown" };
      return { ok: true, breakdown_id: body.breakdown_id, work_order_id: body.primary_work_order_id };
    });
    if (result?.ok === false) return reply.code(result.status || 400).send(result);
    return result;
  });
}
