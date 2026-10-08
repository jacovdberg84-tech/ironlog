// IRONLOG/api/routes/june.routes.js
// June is an admin-only GPT-Live assistant. The browser never receives the
// project API key: it only receives the SDP answer for a short-lived WebRTC
// conversation. June's private Ironlog tools all route through this gateway.
import crypto from "node:crypto";
import { db } from "../db/client.js";
import { getAssetHoursInfoAsOf } from "../utils/assetMeterHours.js";
import { buildJuneMaintenanceSchedule, buildJuneMaintenanceScheduleWorkbook } from "../utils/juneMaintenanceSchedule.js";
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
  cancelInternalCalendarEvent,
  createInternalCalendarEvent,
  getInternalCalendarEvent,
  getInternalCalendarOverview,
  listInternalCalendarEvents,
  prepareInternalCalendarEvent,
  updateInternalCalendarEvent,
} from "../utils/juneInternalCalendar.js";
import {
  finishJuneAvatarTurn,
  getJuneAvatarStatus,
  interruptJuneAvatar,
  sendJuneAvatarAudio,
  startJuneAvatar,
  stopJuneAvatar,
} from "../utils/juneAvatar.js";
import { makeJuneLiveAnswerBrowserCompatible } from "../utils/juneLiveSdp.js";
import { clearJuneMemory, juneMemoryInstructions, recentJuneTurns, saveJuneTurns } from "../utils/juneMemory.js";

const OPENAI_LIVE_URL = "https://api.openai.com/v1/live/sessions";
const LIVE_MODEL = "gpt-live-1";
const LIVE_VOICE = "gleam";
const LIVE_SESSION_TIMEOUT_MS = 50_000;
const LIVE_TICKET_TTL_MS = 3 * 60_000;
const LIVE_TICKET_LIMIT = 24;
const SCHEDULE_REPORT_TTL_MS = 15 * 60_000;
const SCHEDULE_REPORT_LIMIT = 36;
const CALENDAR_APPROVAL_TTL_MS = 10 * 60_000;
const CALENDAR_APPROVAL_LIMIT = 32;
// A reverse proxy can replace 5xx application responses with its own generic
// error page. Keep an upstream Live failure in the 4xx range so the authenticated
// admin receives Ironlog's useful, safe error message instead.
const LIVE_UPSTREAM_FAILURE_STATUS = 424;
const liveSessionTickets = new Map();
const maintenanceScheduleReports = new Map();
const internalCalendarApprovals = new Map();
let lastLiveAttempt = null;

const JUNE_LIVE_INSTRUCTIONS = [
  "You are June, the private executive assistant for the Ironlog administrator.",
  "Personality: you are friendly, warm and playful, with a quick sense of humour. You enjoy the banter of a busy mining workshop and sound like a trusted colleague, not a call-centre bot. Jaco prefers direct answers, not corporate politeness: be practical, decisive and concise.",
  "Workshop language is fine. You may swear casually where it fits naturally, the way people talk on a workshop floor (for example 'damn', 'bloody', 'hell', 'shit', 'crap', 'bugger', 'what a pain in the arse'). Use it for colour, not in every sentence. Never swear at or about a person as an insult, never use slurs or sexual language, and never put swear words into drafts, emails or documents.",
  "Tease Jaco now and then when it is earned, laugh with him, celebrate wins, and call out vague asks, impossible timing or an overloaded day with a grin. Do not become rude, insulting, or performatively sarcastic.",
  "You may point out when Jaco is overloading his day or piling unrelated requests together. Say what should be prioritised, then move on.",
  "If a request contradicts verified facts, challenge it clearly: state the conflict, give the evidence you have, and recommend the sensible next step.",
  "Never tease or swear during safety matters, incidents, injuries, financial or people-sensitive topics, real frustration, or urgent operational decisions. In those cases drop the jokes and be steady, respectful, and direct.",
  "You remember earlier conversations when a memory recap is included below. Greet Jaco naturally and pick up open threads without making him repeat himself.",
  "Keep normal replies to one or two short sentences. For troubleshooting, give one concrete next step and wait for the answer.",
  "Use the backend whenever the user asks about their calendar, email, weather, Ironlog, Borris, tasks, KPIs, equipment, or a draft.",
  "For stores questions—stock on hand, shortages, a part lookup, or items on order—use the Stores briefing tool. It is read-only: never promise that stock was issued, ordered, received, or adjusted.",
  "When the administrator asks for a maintenance schedule for named equipment, prepare the review-only schedule through the backend. Once it reports an Excel file is ready, say it is ready to review and download on screen; never read an internal report identifier aloud.",
  "Ironlog has its own private internal calendar. Outlook and ICS are read-only external sources. For an internal calendar create, move, update, or cancel request, use the preparation tool, repeat the exact change, and tell Jaco to press the on-screen Confirm button. Never claim it was changed until Ironlog reports that confirmation succeeded.",
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
  "For stores questions, use the dedicated read-only Stores tool. Report the on-hand quantity, minimum, shortage, and open-order status exactly as returned; do not invent a receipt date or stock allocation.",
  "For a maintenance schedule request, use the dedicated schedule-draft tool. Its meter and service data are factual Ironlog data; describe forecast dates as planning estimates, never as completed work.",
  "For internal Ironlog calendar changes, prepare a precise proposal only. Do not create, move, or cancel an entry yourself; the authenticated administrator must use the confirmation card in Ironlog. External Outlook and ICS calendars stay read-only.",
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
    name: "june_get_stores_brief",
    description: "Get read-only Ironlog Stores facts: stock on hand, below-minimum items, matching parts, and open part orders. Use this for stock, spares, shortages, or parts-on-order questions.",
    parameters: {
      type: "object",
      properties: {
        scope: { type: "string", enum: ["summary", "low_stock", "on_order", "part_lookup"] },
        query: { type: "string", description: "Optional part code, part description, asset code, or work-order reference to look up." },
      },
      required: [],
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
    name: "june_draft_maintenance_schedule",
    description: "Prepare a review-only maintenance schedule and authenticated Excel download for selected Ironlog equipment. It uses actual meter readings and configured service intervals; it never creates work orders or changes maintenance records.",
    parameters: {
      type: "object",
      properties: {
        asset_codes: {
          type: "array",
          items: { type: "string" },
          description: "Fleet numbers to include, for example GS04AM, G01AM, and F500AM.",
        },
        horizon_days: { type: "number", description: "Planning horizon in days. Use 30 unless the administrator asks for another period." },
        as_of: { type: "string", description: "Optional YYYY-MM-DD date for the meter and schedule snapshot." },
      },
      required: ["asset_codes"],
      additionalProperties: false,
    },
  },
  {
    type: "function",
    name: "june_calendar_overview",
    description: "Get June's upcoming private Ironlog schedule entries plus any connected read-only Outlook or ICS calendar events.",
    parameters: { type: "object", properties: {}, required: [], additionalProperties: false },
  },
  {
    type: "function",
    name: "june_list_internal_calendar_events",
    description: "List upcoming entries from June's private internal Ironlog calendar. This does not access Outlook or ICS.",
    parameters: {
      type: "object",
      properties: {
        from_date: { type: "string", description: "Optional YYYY-MM-DD start date." },
        to_date: { type: "string", description: "Optional YYYY-MM-DD end date." },
      },
      required: [],
      additionalProperties: false,
    },
  },
  {
    type: "function",
    name: "june_prepare_internal_calendar_change",
    description: "Prepare a private internal Ironlog calendar add, update/move, or cancellation for the administrator to explicitly confirm in the browser. This does not create or change anything by itself.",
    parameters: {
      type: "object",
      properties: {
        action: { type: "string", enum: ["create", "update", "cancel"] },
        event_id: { type: "number", description: "Required for update or cancel; obtain it by listing the internal calendar when needed." },
        title: { type: "string" },
        event_date: { type: "string", description: "YYYY-MM-DD." },
        start_time: { type: "string", description: "Optional 24-hour HH:MM." },
        end_time: { type: "string", description: "Optional 24-hour HH:MM." },
        all_day: { type: "boolean" },
        category: { type: "string", enum: ["meeting", "maintenance", "shutdown", "reminder", "follow_up", "other"] },
        asset_code: { type: "string" },
        work_order_id: { type: "number" },
        notes: { type: "string" },
      },
      required: ["action"],
      additionalProperties: false,
    },
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
    ironlog_calendar: {
      state: "ready",
      detail: "Private Ironlog schedule is available. June can prepare changes for explicit confirmation.",
    },
    email: {
      state: outlook.state,
      detail: outlook.detail,
    },
    weather: { state: "ready", detail: "June can use live web lookup for weather questions." },
    ironlog: { state: "connected", detail: "Read-only operational facts, review-only drafts, and maintenance schedule exports are available." },
    borris: { state: "connected", detail: "Asset, PM, and breakdown context is available for engineering analysis." },
    stores: { state: "connected", detail: "Read-only stock, shortages, and parts-on-order facts are available." },
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
      mp.asset_id,
      a.asset_code,
      a.asset_name,
      mp.service_name,
      mp.interval_hours,
      mp.last_service_hours
    FROM maintenance_plans mp
    JOIN assets a ON a.id = mp.asset_id
    WHERE mp.active = 1
  `).map((row) => {
    // A cumulative sum of production hours is not an hour meter. Use the
    // same trusted daily-closing meter logic as Borris and Maintenance.
    const meter = getAssetHoursInfoAsOf(row.asset_id, asOf, db);
    const currentHours = Number(meter.hours || 0);
    const due = Number(row.last_service_hours || 0) + Number(row.interval_hours || 0);
    const remaining = due - currentHours;
    return {
      asset_code: safeText(row.asset_code, 40),
      asset_name: safeText(row.asset_name, 120),
      service_name: safeText(row.service_name || "Service", 120),
      current_hours: number(currentHours),
      meter_source: safeText(meter.source || "unknown", 32),
      meter_date: meter.latest_work_date || null,
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
  // Keep June's engineering brief on the same trusted meter source as Borris.
  // In particular, never describe summed daily run-hours as the machine meter.
  const meter = getAssetHoursInfoAsOf(asset.id, asOf, db);
  const currentMeter = Number(meter.hours || 0);
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
      remaining_hours: number(due - currentMeter),
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
      current_meter: number(currentMeter),
      meter_source: safeText(meter.source || "unknown", 32),
      meter_date: meter.latest_work_date || null,
    },
    open_breakdowns: openBreakdowns,
    service_plans: servicePlans,
    note: "This is factual Ironlog context for June's Borris-style analysis. Confirm diagnoses against the OEM manual and technician findings.",
  };
}

function storesScope(value) {
  const scope = safeText(value, 32).toLowerCase();
  return ["summary", "low_stock", "on_order", "part_lookup"].includes(scope) ? scope : "summary";
}

// June's Stores access is deliberately read-only. This is direct database
// access rather than a browser/API relay so the live tool gets a compact,
// factual answer while normal Stores permissions remain untouched.
function getStoresBrief({ siteCode, scope, query }) {
  const selectedScope = storesScope(scope);
  const search = safeText(query, 120).toUpperCase();
  const like = `%${search}%`;
  const partWhere = search ? "WHERE UPPER(TRIM(p.part_code)) LIKE ? OR UPPER(TRIM(p.part_name)) LIKE ?" : "";
  const partParams = search ? [like, like] : [];
  const partRows = safeRows(`
    SELECT
      p.part_code,
      p.part_name,
      COALESCE(p.min_stock, 0) AS min_stock,
      COALESCE(p.critical, 0) AS critical,
      COALESCE(p.unit_cost, 0) AS unit_cost,
      COALESCE(SUM(sm.quantity), 0) AS on_hand
    FROM parts p
    LEFT JOIN stock_movements sm ON sm.part_id = p.id
    ${partWhere}
    GROUP BY p.id
    ORDER BY p.critical DESC, p.part_code ASC
    LIMIT 30
  `, partParams).map((row) => {
    const onHand = number(row.on_hand, 0, 2);
    const minStock = number(row.min_stock, 0, 2);
    return {
      part_code: safeText(row.part_code, 80),
      part_name: safeText(row.part_name, 180),
      on_hand: onHand,
      min_stock: minStock,
      shortage: number(Math.max(0, minStock - onHand), 0, 2),
      below_min: onHand < minStock,
      critical: Boolean(Number(row.critical || 0)),
      unit_cost: number(row.unit_cost, 0, 2),
    };
  });
  const lowStockRows = safeRows(`
    SELECT
      p.part_code,
      p.part_name,
      COALESCE(p.min_stock, 0) AS min_stock,
      COALESCE(p.critical, 0) AS critical,
      COALESCE(SUM(sm.quantity), 0) AS on_hand
    FROM parts p
    LEFT JOIN stock_movements sm ON sm.part_id = p.id
    GROUP BY p.id
    HAVING COALESCE(SUM(sm.quantity), 0) < COALESCE(p.min_stock, 0)
    ORDER BY p.critical DESC, (COALESCE(p.min_stock, 0) - COALESCE(SUM(sm.quantity), 0)) DESC, p.part_code ASC
    LIMIT 20
  `).map((row) => {
    const onHand = number(row.on_hand, 0, 2);
    const minStock = number(row.min_stock, 0, 2);
    return {
      part_code: safeText(row.part_code, 80),
      part_name: safeText(row.part_name, 180),
      on_hand: onHand,
      min_stock: minStock,
      shortage: number(Math.max(0, minStock - onHand), 0, 2),
      critical: Boolean(Number(row.critical || 0)),
    };
  });
  const partSummary = safeRow(`
    SELECT
      COUNT(*) AS total_parts,
      COALESCE(SUM(CASE WHEN balances.on_hand < balances.min_stock THEN 1 ELSE 0 END), 0) AS below_min,
      COALESCE(SUM(CASE WHEN balances.on_hand < balances.min_stock AND balances.critical = 1 THEN 1 ELSE 0 END), 0) AS critical_below_min
    FROM (
      SELECT p.id, COALESCE(p.min_stock, 0) AS min_stock, COALESCE(p.critical, 0) AS critical, COALESCE(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      GROUP BY p.id
    ) balances
  `) || {};
  const orderRows = safeRows(`
    SELECT
      o.id,
      o.part_code,
      o.part_name,
      o.qty,
      o.unit_cost,
      o.currency,
      o.status,
      o.order_date,
      o.expected_arrival_date,
      o.current_location,
      o.supplier_name,
      o.po_number,
      o.requisition_number
    FROM stores_part_orders o
    WHERE LOWER(TRIM(COALESCE(o.status, 'on_order'))) NOT IN ('arrived', 'cancelled')
      AND LOWER(TRIM(COALESCE(o.site_code, 'main'))) = ?
      AND (
        ? = ''
        OR UPPER(COALESCE(o.part_code, '')) LIKE ?
        OR UPPER(COALESCE(o.part_name, '')) LIKE ?
        OR UPPER(COALESCE(o.po_number, '')) LIKE ?
        OR UPPER(COALESCE(o.requisition_number, '')) LIKE ?
      )
    ORDER BY COALESCE(o.expected_arrival_date, '9999-12-31') ASC, o.order_date DESC, o.id DESC
    LIMIT 20
  `, [safeText(siteCode, 160).toLowerCase() || "main", search, like, like, like, like]).map((row) => ({
    id: Number(row.id || 0),
    part_code: safeText(row.part_code, 80) || null,
    part_name: safeText(row.part_name, 180),
    qty: number(row.qty, 0, 2),
    unit_cost: number(row.unit_cost, 0, 2),
    currency: safeText(row.currency || "USD", 12),
    status: safeText(row.status || "on_order", 40),
    order_date: safeText(row.order_date, 10) || null,
    expected_arrival_date: safeText(row.expected_arrival_date, 10) || null,
    current_location: safeText(row.current_location, 120) || null,
    supplier_name: safeText(row.supplier_name, 120) || null,
    po_number: safeText(row.po_number, 80) || null,
    requisition_number: safeText(row.requisition_number, 80) || null,
  }));

  return {
    review_only: true,
    scope: selectedScope,
    query: search || null,
    summary: {
      total_parts: Number(partSummary.total_parts || 0),
      below_min: Number(partSummary.below_min || 0),
      critical_below_min: Number(partSummary.critical_below_min || 0),
      matching_parts: partRows.length,
      open_part_orders: orderRows.length,
    },
    matching_parts: selectedScope === "low_stock" ? [] : partRows,
    low_stock: selectedScope === "part_lookup" ? partRows.filter((row) => row.below_min) : lowStockRows,
    parts_on_order: selectedScope === "low_stock" ? [] : orderRows,
    note: "Stores data is read-only for June. Confirm actual issue, receipt, allocation, and supplier ETAs with Stores before acting.",
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

function cleanupMaintenanceScheduleReports(now = Date.now()) {
  for (const [id, report] of maintenanceScheduleReports) {
    if (Number(report?.expires_at || 0) <= now) maintenanceScheduleReports.delete(id);
  }
}

function keepMaintenanceScheduleReport(owner, schedule) {
  cleanupMaintenanceScheduleReports();
  if (maintenanceScheduleReports.size >= SCHEDULE_REPORT_LIMIT) {
    const oldest = [...maintenanceScheduleReports.entries()]
      .sort((a, b) => Number(a[1]?.created_at || 0) - Number(b[1]?.created_at || 0))[0];
    if (oldest) maintenanceScheduleReports.delete(oldest[0]);
  }
  const reportId = crypto.randomUUID();
  maintenanceScheduleReports.set(reportId, {
    owner,
    schedule,
    created_at: Date.now(),
    expires_at: Date.now() + SCHEDULE_REPORT_TTL_MS,
  });
  return reportId;
}

function maintenanceScheduleFilename(schedule = {}) {
  const asOf = isDate(schedule.as_of) ? schedule.as_of : todayYmd();
  return `IRONLOG_June_Maintenance_Schedule_${asOf}.xlsx`;
}

function buildMaintenanceScheduleDraft(args, context) {
  const codes = Array.isArray(args.asset_codes) ? args.asset_codes : [args.asset_codes];
  const schedule = buildJuneMaintenanceSchedule({
    assetCodes: codes,
    asOf: isDate(args.as_of) ? String(args.as_of) : todayYmd(),
    horizonDays: args.horizon_days,
    dbConn: db,
  });
  const reportId = keepMaintenanceScheduleReport(context.owner, schedule);
  const rows = schedule.rows.map((row) => ({
    asset_code: row.asset_code,
    equipment: row.asset_name,
    next_service: row.service_name,
    current_meter: row.current_meter,
    meter_unit: row.meter_unit,
    remaining: row.remaining,
    forecast_due_date: row.forecast_due_date,
    status: row.schedule_status,
  }));
  return {
    state: "review_only",
    type: "maintenance_schedule",
    as_of: schedule.as_of,
    horizon_days: schedule.horizon_days,
    summary: schedule.summary,
    missing_asset_codes: schedule.missing_asset_codes,
    schedule: rows,
    download: {
      report_id: reportId,
      filename: maintenanceScheduleFilename(schedule),
      label: "Download maintenance schedule (Excel)",
      expires_in_minutes: Math.round(SCHEDULE_REPORT_TTL_MS / 60_000),
    },
    next_step: "Review the Excel schedule before creating any work orders, service records, or requisitions.",
  };
}

function cleanupInternalCalendarApprovals(now = Date.now()) {
  for (const [id, approval] of internalCalendarApprovals) {
    if (Number(approval?.expires_at || 0) <= now) internalCalendarApprovals.delete(id);
  }
}

function calendarApprovalSummary(action, event) {
  const when = event?.all_day
    ? `${event.event_date} (all day)`
    : `${event.event_date}${event?.start_time ? ` ${event.start_time}` : ""}${event?.end_time ? `–${event.end_time}` : ""}`;
  if (action === "cancel") return `Cancel “${event.title}” on ${when}.`;
  if (action === "update") return `Update “${event.title}” to ${when}.`;
  return `Add “${event.title}” on ${when}.`;
}

function keepInternalCalendarApproval(context, action, event, eventId = null) {
  cleanupInternalCalendarApprovals();
  if (internalCalendarApprovals.size >= CALENDAR_APPROVAL_LIMIT) {
    const oldest = [...internalCalendarApprovals.entries()]
      .sort((a, b) => Number(a[1]?.created_at || 0) - Number(b[1]?.created_at || 0))[0];
    if (oldest) internalCalendarApprovals.delete(oldest[0]);
  }
  const token = crypto.randomUUID();
  internalCalendarApprovals.set(token, {
    owner: context.owner,
    context: { siteCode: context.siteCode, user: context.user },
    action,
    event,
    event_id: eventId,
    created_at: Date.now(),
    expires_at: Date.now() + CALENDAR_APPROVAL_TTL_MS,
  });
  return token;
}

function calendarAction(value) {
  const action = safeText(value, 20).toLowerCase();
  return ["create", "update", "cancel"].includes(action) ? action : "";
}

function calendarEventId(value) {
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed > 0 ? parsed : null;
}

function prepareInternalCalendarChange(args, context) {
  const action = calendarAction(args?.action);
  if (!action) throw new Error("Choose whether June should add, update, or cancel the internal calendar entry.");
  const eventId = calendarEventId(args?.event_id);
  let existing = null;
  if (action !== "create") {
    if (!eventId) throw new Error("June needs the calendar entry to update or cancel. Ask her to list your internal schedule first.");
    existing = getInternalCalendarEvent(context, eventId);
    if (!existing) throw new Error("That Ironlog calendar entry was not found or has already been cancelled.");
  }
  const event = action === "cancel"
    ? existing
    : prepareInternalCalendarEvent(args, existing);
  const token = keepInternalCalendarApproval(context, action, event, eventId);
  return {
    state: "pending_confirmation",
    type: "internal_calendar_change",
    action,
    event,
    approval: {
      token,
      expires_in_minutes: Math.round(CALENDAR_APPROVAL_TTL_MS / 60_000),
    },
    action_summary: calendarApprovalSummary(action, event),
    next_step: "Review the proposed internal calendar change, then use the Confirm button in Ironlog. Nothing has been changed yet.",
  };
}

function applyInternalCalendarApproval(token, owner) {
  cleanupInternalCalendarApprovals();
  const key = safeText(token, 100);
  const approval = internalCalendarApprovals.get(key);
  if (!approval || approval.owner !== owner) return null;
  // Consume before writing so a retry cannot create duplicate meetings.
  internalCalendarApprovals.delete(key);
  const options = { actor: approval.context.user };
  let event;
  if (approval.action === "create") event = createInternalCalendarEvent(approval.context, approval.event, options);
  else if (approval.action === "update") event = updateInternalCalendarEvent(approval.context, approval.event_id, approval.event, options);
  else if (approval.action === "cancel") event = cancelInternalCalendarEvent(approval.context, approval.event_id, options);
  else throw new Error("The proposed internal calendar change was invalid.");
  return {
    action: approval.action,
    event,
    message: approval.action === "cancel"
      ? `Cancelled “${event.title}” in your Ironlog schedule.`
      : approval.action === "update"
        ? `Updated “${event.title}” in your Ironlog schedule.`
        : `Added “${event.title}” to your Ironlog schedule.`,
  };
}

async function getJuneCalendarOverview(context) {
  const internal = getInternalCalendarOverview(context);
  const ics = getIcsCalendarStatus(context);
  const outlook = getOutlookConnectionStatus(context);
  const external = ics.state === "connected"
    ? await getIcsCalendarOverview(context)
    : await getOutlookCalendarOverview(context);
  return {
    internal_calendar: internal,
    external_calendar: external,
    outlook_state: outlook.state,
    ics_state: ics.state,
    note: "Ironlog schedule entries can be changed only through an explicit confirmation. Outlook and ICS entries remain read-only.",
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
  if (tool === "june_get_stores_brief") {
    return getStoresBrief({ siteCode: context.siteCode, scope: safeArgs.scope, query: safeArgs.query });
  }
  if (tool === "june_draft_task") return buildTaskDraft(safeArgs, context.user);
  if (tool === "june_draft_maintenance_schedule") return buildMaintenanceScheduleDraft(safeArgs, context);
  if (tool === "june_calendar_overview") return getJuneCalendarOverview(context);
  if (tool === "june_list_internal_calendar_events") return {
    source: "internal_ironlog",
    events: listInternalCalendarEvents(context, { fromDate: safeArgs.from_date, toDate: safeArgs.to_date, limit: 24 }),
    next_step: "Use a calendar-change preparation only after the administrator has reviewed the relevant entry.",
  };
  if (tool === "june_prepare_internal_calendar_change") return prepareInternalCalendarChange(safeArgs, context);
  if (tool === "june_email_priorities") return getOutlookPriorityEmails(context);
  if (tool === "june_connector_status") return {
    review_only: true,
    connectors: connectorStatus(context),
    outlook: getOutlookConnectionStatus(context),
    ics_calendar: getIcsCalendarStatus(context),
    internal_calendar: getInternalCalendarOverview(context),
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

/** First name for June's memory recap ("Jaco"), from the user record. */
function memoryName(username) {
  try {
    const row = db.prepare(`SELECT full_name FROM users WHERE LOWER(username) = LOWER(?)`).get(String(username || ""));
    const first = String(row?.full_name || "").trim().split(/\s+/)[0];
    return first || String(username || "the administrator");
  } catch {
    return String(username || "the administrator");
  }
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
    const rawAnswer = String(data?.transport?.sdp || "");
    if (!rawAnswer.trim()) throw liveSessionError("June received an incomplete live session response.");
    // Do not make standard Chromium browsers opt into OpenAI's experimental
    // SNAP/WARP SDP extension. The original browser offer tells us whether
    // SNAP was actually negotiated; otherwise use standard SCTP port syntax.
    const answer = makeJuneLiveAnswerBrowserCompatible({ offer: payload?.transport?.sdp, answer: rawAnswer });
    if (answer !== rawAnswer) log.warn("June converted a SNAP-only SDP answer for browser compatibility");
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

/** A from/to date range for the calendar view, or null. */
function calendarRange(query = {}) {
  const ymd = (v) => (/^\d{4}-\d{2}-\d{2}$/.test(String(v || "")) ? String(v) : null);
  const from = ymd(query.from);
  const to = ymd(query.to);
  if (!from || !to || to < from) return null;
  const days = (Date.parse(`${to}T00:00:00Z`) - Date.parse(`${from}T00:00:00Z`)) / 86_400_000;
  return days <= 62 ? { from, to } : null;
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
      internal_calendar: getInternalCalendarOverview(context),
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
    // ?from=YYYY-MM-DD&to=YYYY-MM-DD (the calendar view, up to 62 days); none = upcoming.
    const range = calendarRange(req.query);
    const calendar = await getIcsCalendarEvents({ siteCode: getSiteCode(req), user: getUser(req) }, range);
    return { ok: !calendar.error, calendar };
  });

  app.post("/calendar/ics/remove", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    return { ok: true, ...removeIcsCalendar({ siteCode: getSiteCode(req), user: getUser(req) }) };
  });

  // Internal calendar entries belong to Ironlog, never to the external ICS
  // or Outlook source. A browser-side confirm endpoint is the only writer.
  app.get("/calendar/internal/events", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    const requestedDays = Math.max(1, Math.min(90, Number(req.query?.days) || 21));
    const range = calendarRange(req.query);
    const start = range ? new Date(`${range.from}T00:00:00Z`) : new Date();
    const end = range ? new Date(`${range.to}T00:00:00Z`) : new Date(start);
    if (!range) end.setUTCDate(end.getUTCDate() + requestedDays);
    const context = { siteCode: getSiteCode(req), user: getUser(req) };
    return {
      ok: true,
      calendar: {
        source: "internal_ironlog",
        writable_with_confirmation: true,
        window_days: requestedDays,
        events: listInternalCalendarEvents(context, {
          fromDate: start.toISOString().slice(0, 10),
          toDate: end.toISOString().slice(0, 10),
          limit: range ? 100 : 32,
        }),
      },
    };
  });

  app.post("/calendar/internal/approvals/:token/confirm", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    try {
      const result = applyInternalCalendarApproval(req.params?.token, safetyIdentifier(req));
      if (!result) return reply.code(404).send({ ok: false, error: "This calendar confirmation expired or belongs to another administrator. Ask June to prepare it again." });
      return { ok: true, ...result };
    } catch (error) {
      return reply.code(400).send({ ok: false, error: safeText(error?.message || "Ironlog could not apply that calendar change.", 350) });
    }
  });

  app.post("/calendar/internal/approvals/:token/discard", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    cleanupInternalCalendarApprovals();
    const token = safeText(req.params?.token, 100);
    const approval = internalCalendarApprovals.get(token);
    if (!approval || approval.owner !== safetyIdentifier(req)) {
      return reply.code(404).send({ ok: false, error: "This calendar proposal is no longer available." });
    }
    internalCalendarApprovals.delete(token);
    return { ok: true, discarded: true };
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
  // June's memory of recent conversations (saved by the browser, recapped at session start).
  app.post("/memory/turns", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    const saved = saveJuneTurns({ siteCode: getSiteCode(req), user: getUser(req) }, req.body?.turns);
    return { ok: true, saved };
  });
  app.get("/memory", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    const turns = recentJuneTurns({ siteCode: getSiteCode(req), user: getUser(req) });
    return { ok: true, turns, count: turns.length };
  });
  app.post("/memory/clear", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    return { ok: true, cleared: clearJuneMemory({ siteCode: getSiteCode(req), user: getUser(req) }) };
  });

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
    const context = { siteCode: getSiteCode(req), user: getUser(req) };
    const memory = juneMemoryInstructions(context, { name: memoryName(getUser(req)) });
    const payload = {
      session: {
        model: LIVE_MODEL,
        instructions: memory ? `${JUNE_LIVE_INSTRUCTIONS}\n\n${memory}` : JUNE_LIVE_INSTRUCTIONS,
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

  // A schedule report is generated by June, but it stays inside Ironlog's
  // normal authenticated download flow. The opaque report id is short-lived
  // and bound to the same administrator who asked June to prepare it.
  app.get("/maintenance-schedule/:reportId.xlsx", async (req, reply) => {
    if (!requireJuneAdmin(req, reply)) return;
    cleanupMaintenanceScheduleReports();
    const reportId = safeText(req.params?.reportId, 100);
    const report = maintenanceScheduleReports.get(reportId);
    if (!report || report.owner !== safetyIdentifier(req)) {
      return reply.code(404).send({ ok: false, error: "This June maintenance schedule has expired. Ask June to prepare it again." });
    }
    try {
      const workbook = await buildJuneMaintenanceScheduleWorkbook(report.schedule);
      return reply
        .header("Content-Disposition", `attachment; filename=${maintenanceScheduleFilename(report.schedule)}`)
        .type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .send(workbook);
    } catch (error) {
      req.log.warn({ error: safeText(error?.message || error, 350) }, "June maintenance schedule export failed");
      return reply.code(500).send({ ok: false, error: "June could not prepare the Excel schedule. Please try again." });
    }
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
        owner: safetyIdentifier(req),
      });
      return { ok: true, tool: name, result };
    } catch (error) {
      req.log.warn({ tool: name, error: safeText(error?.message || error, 350) }, "June gateway tool failed");
      return { ok: false, tool: name, result: { error: safeText(error?.message || "June could not complete that Outlook request.", 350) } };
    }
  });
}

