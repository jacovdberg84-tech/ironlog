// IRONLOG/api/routes/june.routes.js
// June is an admin-only GPT-Live assistant. The browser never receives the
// project API key: it only receives the SDP answer for a short-lived WebRTC
// conversation. June's private Ironlog tools all route through this gateway.
import crypto from "node:crypto";
import { db } from "../db/client.js";
import { getRoles, getSiteCode, getUser } from "../utils/request.js";
import {
  beginOutlookAuthorization,
  completeOutlookAuthorization,
  disconnectOutlook,
  getOutlookCalendarOverview,
  getOutlookConnectionStatus,
  getOutlookPriorityEmails,
} from "../utils/juneOutlook.js";
import {
  getIcsCalendarEvents,
  getIcsCalendarOverview,
  getIcsCalendarStatus,
  removeIcsCalendar,
  saveIcsCalendar,
} from "../utils/juneIcsCalendar.js";
import {
  finishJuneAvatarTurn,
  getJuneAvatarStatus,
  interruptJuneAvatar,
  sendJuneAvatarAudio,
  startJuneAvatar,
  stopJuneAvatar,
} from "../utils/juneAvatar.js";

const OPENAI_LIVE_URL = "https://api.openai.com/v1/live/sessions";
const LIVE_MODEL = "gpt-live-1";
const LIVE_VOICE = "gleam";
const LIVE_SESSION_TIMEOUT_MS = 50_000;
const LIVE_TICKET_TTL_MS = 3 * 60_000;
const LIVE_TICKET_LIMIT = 24;
// A reverse proxy can replace 5xx application responses with its own generic
// error page. Keep an upstream Live failure in the 4xx range so the authenticated
// admin receives Ironlog's useful, safe error message instead.
const LIVE_UPSTREAM_FAILURE_STATUS = 424;
const liveSessionTickets = new Map();
let lastLiveAttempt = null;

const JUNE_LIVE_INSTRUCTIONS = [
  "You are June, the private executive assistant for the Ironlog administrator.",
  "Jaco prefers direct answers, not corporate politeness. Speak with sharp, calm confidence; be practical, decisive, and concise.",
  "Use dry wit and an occasional light tease when it suits the moment. It must feel friendly and earned, never cruel, personal, or distracting.",
  "You may point out when Jaco is overloading his day or piling unrelated requests together. Say what should be prioritised, then move on.",
  "If a request contradicts verified facts, challenge it clearly: state the conflict, give the evidence you have, and recommend the sensible next step.",
  "Never tease during safety matters, incidents, injuries, financial or people-sensitive topics, frustration, or urgent operational decisions. In those cases be steady, respectful, and direct.",
  "Keep normal replies to one or two short sentences. For troubleshooting, give one concrete next step and wait for the answer.",
  "Use the backend whenever the user asks about their calendar, email, weather, Ironlog, Borris, tasks, KPIs, equipment, or a draft.",
  "Never claim that a calendar or email account is connected unless the tool result says it is.",
  "Do not say that a task, work order, requisition, service plan, report, or external message has been created. June only prepares review-only drafts.",
  "If a tool reports that approval is required, clearly say what the administrator must review next.",
  "Backchannel policy: use minimal, natural acknowledgements without talking over Jaco.",
  "Interruption policy: if interrupted, stop speaking and listen to the administrator's correction.",
  "Delegation policy: delegate before answering anything that depends on current Ironlog data, connected tools, or careful engineering analysis. Do not guess while waiting for the result.",
].join(" ");

const JUNE_BACKEND_INSTRUCTIONS = [
  "You are the secure planning backend for June inside Ironlog.",
  "Use the available tools for current business facts instead of inventing values.",
  "Private Ironlog functions are read-only or review-only draft functions. Never imply that a record was created, a requisition was issued, or a message was sent.",
  "For weather requests, use web search when it will improve the answer. Keep the final findings concise and operational.",
  "For calendar or email requests, first call the relevant connector function. If it is not connected, explain the connection requirement without pretending to access data.",
  "For engineering questions, use June's Borris engineering tool and identify uncertainty or missing source data.",
].join(" ");

const TOOL_DEFINITIONS = [
  {
    type: "function",
    name: "june_get_task_brief",
    description: "Get the signed-in administrator's open, overdue, or all Ironlog tasks.",
    parameters: {
      type: "object",
      properties: { scope: { type: "string", enum: ["mine", "overdue", "open", "all"] } },
      required: [],
      additionalProperties: false,
    },
  },
  {
    type: "function",
    name: "june_get_ironlog_brief",
    description: "Get the current Ironlog operational brief: open breakdowns, PM due items, and open work orders.",
    parameters: {
      type: "object",
      properties: { as_of: { type: "string", description: "Optional YYYY-MM-DD date." } },
      required: [],
      additionalProperties: false,
    },
  },
  {
    type: "function",
    name: "june_borris_engineering_brief",
    description: "Get factual asset, maintenance, and breakdown context for Borris-style engineering analysis.",
    parameters: {
      type: "object",
      properties: {
        asset_code: { type: "string", description: "Ironlog fleet number, for example A300AM." },
        as_of: { type: "string", description: "Optional YYYY-MM-DD date." },
      },
      required: ["asset_code"],
      additionalProperties: false,
    },
  },
  {
    type: "function",
    name: "june_draft_task",
    description: "Prepare a review-only Ironlog task draft. This does not create a task.",
    parameters: {
      type: "object",
      properties: {
        title: { type: "string" },
        description: { type: "string" },
        priority: { type: "string", enum: ["high", "medium", "low"] },
        due_date: { type: "string", description: "Optional YYYY-MM-DD date." },
      },
      required: ["title"],
      additionalProperties: false,
    },
  },
  {
    type: "function",
    name: "june_calendar_overview",
    description: "Get the next meetings from June's private ICS calendar or authorised Outlook calendar. This is read-only.",
    parameters: { type: "object", properties: {}, required: [], additionalProperties: false },
  },
  {
    type: "function",
    name: "june_email_priorities",
    description: "Check whether June has an authorised email connection before discussing important messages.",
    parameters: { type: "object", properties: {}, required: [], additionalProperties: false },
  },
  {
    type: "function",
    name: "june_connector_status",
    description: "Show which June gateway services are available and which need to be connected.",
    parameters: { type: "object", properties: {}, required: [], additionalProperties: false },
  },
  // GPT-Live's Responses backend executes this hosted tool itself. June uses
  // it for weather and other public current information; private data stays in
  // the Ironlog functions above.
  { type: "web_search" },
];

function isDate(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "").trim());
}

function todayYmd() {
  return new Date().toISOString().slice(0, 10);
}

function safeText(value, max = 500) {
  return String(value || "").trim().slice(0, max);
}

function number(value, fallback = 0, digits = 1) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return fallback;
  const factor = 10 ** digits;
  return Math.round(parsed * factor) / factor;
}

function hasAdminRole(req) {
  // Never apply the application's legacy local-mode default here. June must
  // fail closed unless the authenticated request carries the admin role.
  return getRoles(req).includes("admin");
}

function requireJuneAdmin(req, reply) {
  if (hasAdminRole(req)) return true;
  reply.code(403).send({ ok: false, error: "June is available to administrators only." });
  return false;
}

function safeRows(sql, params = []) {
  try {
    return db.prepare(sql).all(...params);
  } catch {
    return [];
  }
}

function safeRow(sql, params = []) {
  try {
    return db.prepare(sql).get(...params) || null;
  } catch {
    return null;
  }
}

function connectorStatus(context = {}) {
  const outlook = getOutlookConnectionStatus(context);
  const icsCalendar = getIcsCalendarStatus(context);
  const calendar = icsCalendar.state === "connected" || outlook.state !== "connected" ? icsCalendar : outlook;
  return {
    calendar: {
      state: calendar.state,
      detail: calendar.detail,
    },
    email: {
      state: outlook.state,
      detail: outlook.detail,
    },
    weather: { state: "ready", detail: "June can use live web lookup for weather questions." },
    ironlog: { state: "connected", detail: "Read-only operational facts and review-only drafts are available." },
    borris: { state: "connected", detail: "Asset, PM, and breakdown context is available for engineering analysis." },
  };
}

function getTaskBrief({ siteCode, user, scope }) {
  const selectedScope = ["mine", "overdue", "open", "all"].includes(scope) ? scope : "mine";
  let where = "WHERE t.site_code = ?";
  const params = [siteCode];
  if (selectedScope === "mine") {
    where += " AND LOWER(TRIM(COALESCE(t.assigned_to, ''))) = LOWER(TRIM(?))";
    params.push(user);
  } else if (selectedScope === "overdue") {
    where += " AND t.due_date IS NOT NULL AND date(t.due_date) < date('now') AND LOWER(TRIM(COALESCE(t.status, 'open'))) <> 'done'";
  } else if (selectedScope === "open") {
    where += " AND LOWER(TRIM(COALESCE(t.status, 'open'))) <> 'done'";
  }
  const rows = safeRows(`
    SELECT t.id, t.title, t.description, t.status, t.priority, t.project, t.assigned_to, t.due_date
    FROM tasks t
    ${where}
    ORDER BY
      CASE LOWER(COALESCE(t.priority, 'medium')) WHEN 'high' THEN 1 WHEN 'medium' THEN 2 ELSE 3 END,
      CASE WHEN t.due_date IS NOT NULL AND date(t.due_date) < date('now') THEN 0 ELSE 1 END,
      t.due_date ASC, t.created_at DESC
    LIMIT 12
  `, params).map((row) => ({
    id: Number(row.id || 0),
    title: safeText(row.title, 180),
    status: safeText(row.status || "open", 32),
    priority: safeText(row.priority || "medium", 32),
    project: safeText(row.project, 100),
    due_date: safeText(row.due_date, 10) || null,
    assigned_to: safeText(row.assigned_to, 100) || null,
    description: safeText(row.description, 300) || null,
  }));
  return {
    review_only: true,
    scope: selectedScope,
    user,
    task_count: rows.length,
    tasks: rows,
  };
}

function getPmDueRows(asOf) {
  return safeRows(`
    SELECT
      a.asset_code,
      a.asset_name,
      mp.service_name,
      mp.interval_hours,
      mp.last_service_hours,
      COALESCE((
        SELECT SUM(dh.hours_run)
        FROM daily_hours dh
        WHERE dh.asset_id = mp.asset_id
          AND dh.is_used = 1
          AND dh.hours_run > 0
          AND dh.work_date <= ?
      ), 0) AS current_hours
    FROM maintenance_plans mp
    JOIN assets a ON a.id = mp.asset_id
    WHERE mp.active = 1
  `, [asOf]).map((row) => {
    const due = Number(row.last_service_hours || 0) + Number(row.interval_hours || 0);
    const remaining = due - Number(row.current_hours || 0);
    return {
      asset_code: safeText(row.asset_code, 40),
      asset_name: safeText(row.asset_name, 120),
      service_name: safeText(row.service_name || "Service", 120),
      next_due_hours: number(due),
      remaining_hours: number(remaining),
      status: remaining < 0 ? "overdue" : remaining <= 50 ? "due_soon" : "planned",
    };
  }).filter((row) => row.status !== "planned")
    .sort((a, b) => a.remaining_hours - b.remaining_hours || a.asset_code.localeCompare(b.asset_code))
    .slice(0, 12);
}

function getIronlogBrief(asOf) {
  const breakdowns = safeRows(`
    SELECT b.id, a.asset_code, a.asset_name, b.breakdown_date, b.component, b.description,
      COALESCE(b.critical, 0) AS critical, COALESCE(b.downtime_total_hours, 0) AS downtime_hours,
      b.parts_status, b.ets_repair_date
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    WHERE UPPER(TRIM(COALESCE(b.status, ''))) = 'OPEN'
    ORDER BY COALESCE(b.critical, 0) DESC, b.breakdown_date ASC, a.asset_code ASC
    LIMIT 12
  `).map((row) => ({
    id: Number(row.id || 0),
    asset_code: safeText(row.asset_code, 40),
    asset_name: safeText(row.asset_name, 120),
    breakdown_date: safeText(row.breakdown_date, 10) || null,
    component: safeText(row.component, 140) || null,
    description: safeText(row.description, 400) || null,
    critical: Boolean(row.critical),
    downtime_hours: number(row.downtime_hours),
    parts_status: safeText(row.parts_status, 100) || null,
    ets_repair_date: safeText(row.ets_repair_date, 10) || null,
  }));
  const pm_due = getPmDueRows(asOf);
  const workOrderCount = Number(safeRow(`
    SELECT COUNT(*) AS count
    FROM work_orders
    WHERE LOWER(TRIM(COALESCE(status, 'open'))) NOT IN ('closed', 'completed', 'approved')
  `)?.count || 0);
  return {
    as_of: asOf,
    review_only: true,
    metrics: {
      open_breakdowns: breakdowns.length,
      critical_breakdowns: breakdowns.filter((row) => row.critical).length,
      pm_overdue: pm_due.filter((row) => row.status === "overdue").length,
      pm_due_soon: pm_due.filter((row) => row.status === "due_soon").length,
      open_work_orders: workOrderCount,
    },
    open_breakdowns: breakdowns,
    pm_due,
  };
}

function getBorrisEngineeringBrief(assetCode, asOf) {
  const code = safeText(assetCode, 40).toUpperCase();
  const asset = safeRow(`
    SELECT id, asset_code, asset_name, category
    FROM assets
    WHERE UPPER(TRIM(asset_code)) = ?
    LIMIT 1
  `, [code]);
  if (!asset) return { found: false, asset_code: code, message: `No Ironlog asset found for ${code}.` };
  const runHours = Number(safeRow(`
    SELECT COALESCE(SUM(hours_run), 0) AS hours
    FROM daily_hours
    WHERE asset_id = ? AND is_used = 1 AND hours_run > 0 AND work_date <= ?
  `, [asset.id, asOf])?.hours || 0);
  const openBreakdowns = safeRows(`
    SELECT component, description, breakdown_date, COALESCE(critical, 0) AS critical,
      COALESCE(downtime_total_hours, 0) AS downtime_hours, parts_status, ets_repair_date
    FROM breakdowns
    WHERE asset_id = ? AND UPPER(TRIM(COALESCE(status, ''))) = 'OPEN'
    ORDER BY COALESCE(critical, 0) DESC, breakdown_date ASC
    LIMIT 8
  `, [asset.id]).map((row) => ({
    component: safeText(row.component, 140) || null,
    description: safeText(row.description, 400) || null,
    breakdown_date: safeText(row.breakdown_date, 10) || null,
    critical: Boolean(row.critical),
    downtime_hours: number(row.downtime_hours),
    parts_status: safeText(row.parts_status, 100) || null,
    ets_repair_date: safeText(row.ets_repair_date, 10) || null,
  }));
  const servicePlans = safeRows(`
    SELECT service_name, interval_hours, last_service_hours
    FROM maintenance_plans
    WHERE asset_id = ? AND active = 1
    ORDER BY interval_hours ASC, id ASC
    LIMIT 10
  `, [asset.id]).map((row) => {
    const due = Number(row.last_service_hours || 0) + Number(row.interval_hours || 0);
    return {
      service_name: safeText(row.service_name || "Service", 120),
      interval_hours: number(row.interval_hours),
      next_due_hours: number(due),
      remaining_hours: number(due - runHours),
    };
  });
  return {
    found: true,
    review_only: true,
    as_of: asOf,
    asset: {
      asset_code: safeText(asset.asset_code, 40),
      asset_name: safeText(asset.asset_name, 120),
      category: safeText(asset.category, 80) || null,
      accumulated_run_hours: number(runHours),
    },
    open_breakdowns: openBreakdowns,
    service_plans: servicePlans,
    note: "This is factual Ironlog context for June's Borris-style analysis. Confirm diagnoses against the OEM manual and technician findings.",
  };
}

function buildTaskDraft(args, user) {
  const priority = ["high", "medium", "low"].includes(String(args.priority || "").toLowerCase())
    ? String(args.priority).toLowerCase()
    : "medium";
  return {
    state: "review_only",
    type: "task",
    title: safeText(args.title, 180),
    description: safeText(args.description, 1200) || null,
    priority,
    due_date: isDate(args.due_date) ? String(args.due_date) : null,
    proposed_assignee: user,
    next_step: "Review the draft, then create it manually in the Ironlog Tasks panel.",
  };
}

async function executeGatewayTool(name, args, context) {
  const tool = safeText(name, 80);
  const safeArgs = args && typeof args === "object" && !Array.isArray(args) ? args : {};
  if (tool === "june_get_task_brief") {
    return getTaskBrief({ siteCode: context.siteCode, user: context.user, scope: safeText(safeArgs.scope, 20).toLowerCase() });
  }
  if (tool === "june_get_ironlog_brief") {
    return getIronlogBrief(isDate(safeArgs.as_of) ? String(safeArgs.as_of) : todayYmd());
  }
  if (tool === "june_borris_engineering_brief") {
    return getBorrisEngineeringBrief(safeArgs.asset_code, isDate(safeArgs.as_of) ? String(safeArgs.as_of) : todayYmd());
  }
  if (tool === "june_draft_task") return buildTaskDraft(safeArgs, context.user);
  if (tool === "june_calendar_overview") {
    return getIcsCalendarStatus(context).state === "connected"
      ? getIcsCalendarOverview(context)
      : getOutlookCalendarOverview(context);
  }
  if (tool === "june_email_priorities") return getOutlookPriorityEmails(context);
  if (tool === "june_connector_status") return {
    review_only: true,
    connectors: connectorStatus(context),
    outlook: getOutlookConnectionStatus(context),
    ics_calendar: getIcsCalendarStatus(context),
  };
  return { error: `Unsupported June tool: ${tool}` };
}

function liveBackendModel() {
  // June's voice session needs a Responses-delegation model. Do not inherit
  // Borris's text-only model setting here: production Borris is deliberately
  // pinned to a lighter model, which can make the Live session fail before it
  // has a chance to start.
  return safeText(process.env.JUNE_BACKEND_MODEL || "gpt-5.6-terra", 100);
}

function safetyIdentifier(req) {
  const value = `${getUser(req)}:${getSiteCode(req)}`.toLowerCase();
  return crypto.createHash("sha256").update(value).digest("hex").slice(0, 48);
}

function cleanupLiveTickets(now = Date.now()) {
  for (const [id, ticket] of liveSessionTickets) {
    if (Number(ticket?.expires_at || 0) <= now) liveSessionTickets.delete(id);
  }
}

function liveSessionError(message, statusCode = LIVE_UPSTREAM_FAILURE_STATUS) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function recordLiveAttempt({ state, message = "", statusCode = 0 }) {
  lastLiveAttempt = {
    state: safeText(state, 32) || "unknown",
    message: safeText(message, 400),
    status_code: Number(statusCode || 0) || 0,
    recorded_at: new Date().toISOString(),
  };
}

async function createOpenAiLiveSession({ apiKey, payload, log, safetyId }) {
  const abortController = new AbortController();
  const requestTimeout = setTimeout(() => abortController.abort(), LIVE_SESSION_TIMEOUT_MS);
  try {
    const response = await fetch(OPENAI_LIVE_URL, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
        "OpenAI-Safety-Identifier": safetyId,
      },
      body: JSON.stringify(payload),
      signal: abortController.signal,
    });
    const text = await response.text();
    let data = {};
    try { data = JSON.parse(text); } catch {}
    if (!response.ok) {
      const upstreamError = safeText(data?.error?.message || text, 300);
      log.warn({ status: response.status, error: upstreamError }, "June Live session request failed");
      throw liveSessionError(
        upstreamError
          ? `OpenAI could not start June's live session: ${upstreamError}`
          : "OpenAI could not start June's live session. Confirm that this project has GPT-Live access.",
        // Keep this distinct from Ironlog's own login status and from a proxy
        // gateway error. The browser needs the JSON message, not an HTML 502.
        LIVE_UPSTREAM_FAILURE_STATUS,
      );
    }
    const answer = String(data?.transport?.sdp || "").trim();
    if (!answer) throw liveSessionError("June received an incomplete live session response.");
    return {
      session_id: safeText(data?.session?.id, 160),
      sdp: answer,
      model: LIVE_MODEL,
      voice: LIVE_VOICE,
    };
  } catch (error) {
    if (error?.statusCode) throw error;
    log.warn({ error: safeText(error?.message || error, 300) }, "June Live network request failed");
    if (error?.name === "AbortError") {
      throw liveSessionError("June's live-session request timed out. Please retry; if it continues, check the server's connection to OpenAI.", 504);
    }
    throw liveSessionError("June could not contact the live voice service.");
  } finally {
    clearTimeout(requestTimeout);
  }
}

export default async function juneRoutes(app) {
  app.get("/status", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    const context = { siteCode: getSiteCode(req), user: getUser(req) };
    return {
      ok: true,
      gateway_name: "Emma Tool Gateway",
      live_ready: Boolean(String(process.env.OPENAI_API_KEY || "").trim()),
      live_start_mode: "secure_async_poll",
      model: LIVE_MODEL,
      voice: LIVE_VOICE,
      backend_model: liveBackendModel(),
      last_live_attempt: lastLiveAttempt,
      avatar: getJuneAvatarStatus(),
      connectors: connectorStatus(context),
      outlook: getOutlookConnectionStatus(context),
      ics_calendar: getIcsCalendarStatus(context),
    };
  });

  app.get("/calendar/ics/status", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    return { ok: true, calendar: getIcsCalendarStatus({ siteCode: getSiteCode(req), user: getUser(req) }) };
  });

  app.post("/calendar/ics", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    try {
      const calendar = await saveIcsCalendar(
        { siteCode: getSiteCode(req), user: getUser(req) },
        { url: req.body?.url, name: req.body?.name },
      );
      return { ok: true, calendar };
    } catch (error) {
      return reply.code(400).send({ ok: false, error: safeText(error?.message || "The ICS calendar could not be saved.", 350) });
    }
  });

  app.get("/calendar/ics/events", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    const calendar = await getIcsCalendarEvents({ siteCode: getSiteCode(req), user: getUser(req) });
    return { ok: !calendar.error, calendar };
  });

  app.post("/calendar/ics/remove", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    return { ok: true, ...removeIcsCalendar({ siteCode: getSiteCode(req), user: getUser(req) }) };
  });

  // LemonSlice renders June's visual only. The authenticated browser keeps
  // OpenAI's existing WebRTC audio and delegates response audio to this
  // server-side tunnel; no provider key or LiveKit publish token reaches it.
  app.post("/avatar/session", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    try {
      const session = await startJuneAvatar({ owner: safetyIdentifier(req), log: req.log });
      return { ok: true, session };
    } catch (error) {
      return reply.code(Number(error?.statusCode || 503)).send({ ok: false, error: safeText(error?.message || "June's live visual could not start.", 350) });
    }
  });

  app.post("/avatar/session/:id/audio", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    try {
      return sendJuneAvatarAudio({ id: req.params?.id, owner: safetyIdentifier(req), audio: req.body?.audio });
    } catch (error) {
      return reply.code(Number(error?.statusCode || 503)).send({ ok: false, error: safeText(error?.message || "June's visual-audio stream could not continue.", 350) });
    }
  });

  app.post("/avatar/session/:id/end-turn", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    try {
      return finishJuneAvatarTurn({ id: req.params?.id, owner: safetyIdentifier(req) });
    } catch (error) {
      return reply.code(Number(error?.statusCode || 503)).send({ ok: false, error: safeText(error?.message || "June's visual could not finish this response.", 350) });
    }
  });

  app.post("/avatar/session/:id/interrupt", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    try {
      return interruptJuneAvatar({ id: req.params?.id, owner: safetyIdentifier(req) });
    } catch (error) {
      return reply.code(Number(error?.statusCode || 503)).send({ ok: false, error: safeText(error?.message || "June's visual could not be interrupted.", 350) });
    }
  });

  app.post("/avatar/session/:id/stop", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    return stopJuneAvatar({ id: req.params?.id, owner: safetyIdentifier(req) });
  });

  // Starts the user-authorised Microsoft OAuth flow. This endpoint remains
  // authenticated: only the callback is public, and it is protected by an
  // expiring, single-use state value bound to this administrator.
  app.get("/outlook/connect", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    try {
      const connection = beginOutlookAuthorization({ siteCode: getSiteCode(req), user: getUser(req) });
      return { ok: true, ...connection };
    } catch (error) {
      return reply.code(503).send({ ok: false, error: safeText(error?.message || "Outlook is not configured on the server.", 350) });
    }
  });

  // Microsoft redirects here after consent. Auth cannot be required on this
  // callback because it arrives from Microsoft's browser redirect, not from a
  // fetch carrying Ironlog's Bearer token. The stored state is the authority.
  app.get("/outlook/callback", async (req, reply) => {
    try {
      const result = await completeOutlookAuthorization({
        state: req.query?.state,
        code: req.query?.code,
        error: req.query?.error,
        errorDescription: req.query?.error_description,
      });
      if (result.redirect_url) return reply.redirect(result.redirect_url);
      return reply.code(400).type("text/plain; charset=utf-8").send(result.message || "Outlook connection could not be completed.");
    } catch (error) {
      req.log.warn({ error: safeText(error?.message || error, 350) }, "June Outlook callback failed");
      return reply.code(502).type("text/plain; charset=utf-8").send("Outlook could not be connected right now. Return to Ironlog and try again.");
    }
  });

  app.post("/outlook/disconnect", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    const result = disconnectOutlook({ siteCode: getSiteCode(req), user: getUser(req) });
    return { ok: true, ...result };
  });

  // Receives a browser-created SDP offer and immediately returns a short-lived
  // ticket. The OpenAI call happens in the background, avoiding proxy timeouts
  // while GPT-Live allocates the WebRTC session. The SDP answer remains bound
  // to the authenticated administrator who created the ticket.
  app.post("/live/session", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    // SDP is line-oriented and must be forwarded exactly as the browser
    // generated it. In particular, do not trim the terminal CRLF: some SDP
    // parsers treat a trimmed offer as an unexpected EOF.
    const sdp = typeof req.body?.sdp === "string" ? req.body.sdp : "";
    if (!sdp.trim() || sdp.length > 750000) {
      return reply.code(400).send({ ok: false, error: "A valid WebRTC offer is required." });
    }
    const apiKey = String(process.env.OPENAI_API_KEY || "").trim();
    if (!apiKey) {
      return reply.code(503).send({ ok: false, error: "June Live is not configured on the server." });
    }
    cleanupLiveTickets();
    if (liveSessionTickets.size >= LIVE_TICKET_LIMIT) {
      return reply.code(429).send({ ok: false, error: "June is starting several sessions. Please retry in a moment." });
    }
    const owner = safetyIdentifier(req);
    const payload = {
      session: {
        model: LIVE_MODEL,
        instructions: JUNE_LIVE_INSTRUCTIONS,
        audio: { output: { voice: LIVE_VOICE } },
        store: false,
        delegation: {
          type: "responses",
          responses: {
            model: liveBackendModel(),
            instructions: JUNE_BACKEND_INSTRUCTIONS,
            tools: TOOL_DEFINITIONS,
            tool_choice: "auto",
          },
        },
      },
      transport: { type: "webrtc", sdp },
    };
    const ticket = crypto.randomUUID();
    const pending = {
      owner,
      state: "pending",
      status_code: 202,
      created_at: Date.now(),
      expires_at: Date.now() + LIVE_TICKET_TTL_MS,
      result: null,
      error: "",
    };
    liveSessionTickets.set(ticket, pending);
    recordLiveAttempt({ state: "pending" });
    void createOpenAiLiveSession({ apiKey, payload, log: req.log, safetyId: owner })
      .then((result) => {
        pending.state = "ready";
        pending.status_code = 200;
        pending.result = result;
        recordLiveAttempt({ state: "ready", statusCode: 200 });
      })
      .catch((error) => {
        pending.state = "failed";
        pending.status_code = Number(error?.statusCode || LIVE_UPSTREAM_FAILURE_STATUS);
        pending.error = safeText(error?.message || "June could not start a live voice session.", 400);
        recordLiveAttempt({
          state: "failed",
          statusCode: pending.status_code,
          message: pending.error,
        });
      });
    return reply.code(202).send({ ok: true, state: "pending", ticket, retry_after_ms: 800 });
  });

  app.get("/live/session/:ticket", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    cleanupLiveTickets();
    const ticketId = safeText(req.params?.ticket, 100);
    const ticket = liveSessionTickets.get(ticketId);
    if (!ticket || ticket.owner !== safetyIdentifier(req)) {
      return reply.code(404).send({ ok: false, error: "June session request was not found. Please start again." });
    }
    if (ticket.state === "pending") {
      return reply.code(202).send({ ok: true, state: "pending", retry_after_ms: 800 });
    }
    if (ticket.state === "failed") {
      return reply.code(ticket.status_code || 502).send({ ok: false, state: "failed", error: ticket.error || "June could not start a live voice session." });
    }
    return { ok: true, state: "ready", ...ticket.result };
  });

  // This gateway is the only executor for private June functions. Its tool set
  // intentionally contains no write operation: live conversations can inform
  // and draft, but a person approves every operational change in Ironlog.
  app.post("/gateway/execute", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    const name = safeText(req.body?.name, 80);
    const args = req.body?.arguments;
    try {
      const result = await executeGatewayTool(name, args, {
        siteCode: getSiteCode(req),
        user: getUser(req),
      });
      return { ok: true, tool: name, result };
    } catch (error) {
      req.log.warn({ tool: name, error: safeText(error?.message || error, 350) }, "June gateway tool failed");
      return { ok: false, tool: name, result: { error: safeText(error?.message || "June could not complete that Outlook request.", 350) } };
    }
  });
}

