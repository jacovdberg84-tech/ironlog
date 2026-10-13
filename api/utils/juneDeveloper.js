// IRONLOG/api/utils/juneDeveloper.js — June and the Developer Centre.
//
// June can tell the administrator what Claude is working on, and turn what
// they say into a request or a reply for Claude. She only ever PREPARES: the
// browser shows an editable card and the administrator presses Send, which
// goes through the normal admin-only /api/developer endpoints. June never
// sends, merges or approves anything herself.
import { db as defaultDb } from "../db/client.js";
import {
  MODELS,
  STATUS_TEXT,
  ensureDeveloperTables,
  githubClient,
  githubSettings,
  refreshDeveloperRow,
  validateRequest,
} from "./developerClaude.js";

// Statuses where the next move is the administrator's.
const NEEDS_YOU = new Set(["code_ready", "in_review", "checks_failing", "needs_you", "deploy_failed", "not_started"]);
const IN_PROGRESS = new Set(["sending", "waiting", "working", "checks_running", "merged", "deploying"]);
const FINISHED = new Set(["deployed", "closed", "failed"]);

const site = (context) => String(context?.siteCode || "main").toLowerCase();

function detailOf(row) {
  try { return row.status_json ? JSON.parse(row.status_json) : {}; } catch { return {}; }
}

/** Claude's note without markdown symbols, short enough to speak. */
export function spokenNote(note, max = 600) {
  return String(note || "")
    .replace(/```[\s\S]*?```/g, " ")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .replace(/^#+\s*/gm, "")
    .replace(/^\s*-\s*\[x\]\s*/gim, "Done: ")
    .replace(/^\s*-\s*\[ \]\s*/gm, "To do: ")
    .replace(/[*_`>]/g, "")
    .replace(/-{3,}/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, max);
}

function summary(row) {
  const d = detailOf(row);
  return {
    request_id: row.id,
    title: row.title,
    status: row.status,
    status_text: STATUS_TEXT[row.status] || row.status,
    model: MODELS[row.model]?.label || row.model,
    priority: row.priority,
    issue_number: row.issue_number,
    asked_by: row.created_by,
    asked_at: row.created_at,
    needs_you: NEEDS_YOU.has(row.status),
    claude_note: spokenNote(d.claude_note),
    files_changed: d.changes?.files?.length || 0,
    pull_request: d.pull_request ? { number: d.pull_request.number, merged: Boolean(d.pull_request.merged_at) } : null,
    failed_checks: d.checks?.failed?.length ? d.checks.failed : undefined,
  };
}

/** Bring unfinished requests up to date from GitHub (newest five), when connected. */
async function refreshActive(context, { dbConn, fetchImpl }) {
  const s = githubSettings();
  if (!s.configured) return false;
  const gh = githubClient({ token: s.token, repo: s.repo, fetchImpl });
  const rows = dbConn.prepare(`SELECT * FROM developer_requests WHERE site_code = ? AND issue_number IS NOT NULL ORDER BY id DESC LIMIT 20`)
    .all(site(context)).filter((r) => !FINISHED.has(r.status)).slice(0, 5);
  for (const row of rows) await refreshDeveloperRow({ dbConn, row, gh, repo: s.repo });
  return true;
}

/** What Claude is working on, for "June, what's Claude up to?". */
export async function developerStatusForJune(context, { dbConn = defaultDb, fetchImpl, refresh = true } = {}) {
  ensureDeveloperTables(dbConn);
  const live = refresh ? await refreshActive(context, { dbConn, fetchImpl }).catch(() => false) : false;
  const rows = dbConn.prepare(`SELECT * FROM developer_requests WHERE site_code = ? ORDER BY id DESC LIMIT 10`).all(site(context));
  const items = rows.map(summary);
  return {
    review_only: true,
    github_connected: githubSettings().configured,
    refreshed_from_github: live,
    counts: {
      waiting_on_you: items.filter((i) => NEEDS_YOU.has(i.status)).length,
      in_progress: items.filter((i) => IN_PROGRESS.has(i.status)).length,
      live: items.filter((i) => i.status === "deployed").length,
    },
    requests: items,
    note: "Requests are reviewed and merged on GitHub by the administrator. June can prepare a new request or a reply for them to send.",
  };
}

/** Find the request the administrator means: id, issue number (#61) or words from the title. */
export function findDeveloperRequest(context, ref, { dbConn = defaultDb } = {}) {
  ensureDeveloperTables(dbConn);
  const rows = dbConn.prepare(`SELECT * FROM developer_requests WHERE site_code = ? ORDER BY id DESC LIMIT 50`).all(site(context));
  const text = String(ref ?? "").trim();
  if (!text) return rows.find((r) => !FINISHED.has(r.status)) || null;
  const num = Number(text.replace(/^#/, ""));
  if (Number.isInteger(num) && num > 0) {
    return rows.find((r) => r.issue_number === num) || rows.find((r) => r.id === num) || null;
  }
  const words = text.toLowerCase().split(/\s+/).filter((w) => w.length > 2);
  const scored = rows
    .map((r) => ({ r, score: words.filter((w) => r.title.toLowerCase().includes(w)).length }))
    .filter((x) => x.score > 0)
    .sort((a, b) => b.score - a.score || Number(FINISHED.has(a.r.status)) - Number(FINISHED.has(b.r.status)) || b.r.id - a.r.id);
  return scored[0]?.r || null;
}

/** A new request for Claude, shown on screen for the administrator to edit and send. */
export function prepareDeveloperRequest(args = {}) {
  const draft = validateRequest({
    title: args.title,
    request_text: args.request_text,
    priority: args.priority,
    model: args.model,
  });
  if (draft.error) return { ready: false, error: draft.error };
  return {
    ready: true,
    kind: "developer_request",
    draft,
    shown_on_screen: true,
    next_step: "The request is on screen. The administrator can edit it, then press Send to Claude. Nothing has been sent.",
  };
}

/** A reply to Claude on an existing request, shown on screen to edit and send. */
export function prepareDeveloperReply(context, args = {}, { dbConn = defaultDb } = {}) {
  const message = String(args.message || "").replace(/\r\n/g, "\n").trim().slice(0, 6000);
  if (message.length < 3) return { ready: false, error: "Say what Claude should do." };
  const row = findDeveloperRequest(context, args.request, { dbConn });
  if (!row) return { ready: false, error: "I could not find that request. Ask which one, or list the requests first." };
  if (!row.issue_number) return { ready: false, error: "That request never reached GitHub, so Claude cannot get a reply." };
  if (FINISHED.has(row.status)) return { ready: false, error: `"${row.title}" is already finished (${STATUS_TEXT[row.status] || row.status}). Prepare a new request instead.` };
  return {
    ready: true,
    kind: "developer_reply",
    draft: { request_id: row.id, title: row.title, issue_number: row.issue_number, message },
    shown_on_screen: true,
    next_step: "The reply is on screen. The administrator can edit it, then press Send reply. Nothing has been sent.",
  };
}

/** Open one request on the Developer tab. */
export function openDeveloperRequest(context, args = {}, { dbConn = defaultDb } = {}) {
  const row = findDeveloperRequest(context, args.request, { dbConn });
  if (!row) return { found: false, error: "I could not find that request." };
  return { found: true, shown_on_screen: true, request: summary(row) };
}

/** One line for the daily briefing. */
export function developerBriefing(context, { dbConn = defaultDb } = {}) {
  ensureDeveloperTables(dbConn);
  const rows = dbConn.prepare(`SELECT * FROM developer_requests WHERE site_code = ? ORDER BY id DESC LIMIT 20`).all(site(context));
  const needsYou = rows.filter((r) => NEEDS_YOU.has(r.status));
  const working = rows.filter((r) => IN_PROGRESS.has(r.status));
  if (!needsYou.length && !working.length) return null;
  return {
    waiting_on_you: needsYou.map((r) => ({ title: r.title, status: STATUS_TEXT[r.status] || r.status })),
    claude_working_on: working.map((r) => r.title),
  };
}
