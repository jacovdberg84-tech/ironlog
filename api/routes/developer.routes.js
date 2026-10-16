// IRONLOG/api/routes/developer.routes.js — Developer tab: requests to Claude.
// Admin only. Each request becomes a GitHub issue for the Claude GitHub
// Action; see api/utils/developerClaude.js for the full flow.
import { db } from "../db/client.js";
import { getRoles, getSiteCode, getUser } from "../utils/request.js";
import {
  MODELS,
  STATUS_TEXT,
  buildIssue,
  ensureDeveloperTables,
  githubClient,
  githubSettings,
  inspectRequest,
  refreshDeveloperRow,
  validateRequest,
} from "../utils/developerClaude.js";
import { readAttachments } from "../utils/developerAttachments.js";

const FINISHED = new Set(["deployed", "closed", "failed"]);

function requireAdmin(req, reply) {
  if (getRoles(req).includes("admin")) return true;
  reply.code(403).send({ ok: false, error: "The Developer tab is for administrators." });
  return false;
}

function publicRow(row) {
  let detail = null;
  try { detail = row.status_json ? JSON.parse(row.status_json) : null; } catch { detail = null; }
  let attachments = [];
  try { attachments = row.attachments_json ? JSON.parse(row.attachments_json) : []; } catch { attachments = []; }
  return {
    id: row.id,
    title: row.title,
    request_text: row.request_text,
    priority: row.priority,
    model: row.model,
    model_label: MODELS[row.model]?.label || row.model,
    created_by: row.created_by,
    created_at: row.created_at,
    issue_number: row.issue_number,
    issue_url: row.issue_url,
    status: row.status,
    status_text: STATUS_TEXT[row.status] || row.status,
    checked_at: row.checked_at,
    attachments,
    detail,
  };
}

export default async function developerRoutes(app, opts = {}) {
  ensureDeveloperTables(db);
  const fetchImpl = opts.fetchImpl || globalThis.fetch;
  const client = () => {
    const s = githubSettings();
    return s.configured ? { s, gh: githubClient({ token: s.token, repo: s.repo, fetchImpl }) } : { s, gh: null };
  };
  const rowById = (id, site) => db.prepare(`SELECT * FROM developer_requests WHERE id = ? AND site_code = ?`).get(id, site);

  const refresh = (row, gh, repo) => refreshDeveloperRow({ dbConn: db, row, gh, repo });

  // What is set up, so the tab can show a checklist instead of failing later.
  app.get("/status", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const { s, gh } = client();
    const out = { ok: true, repo: s.repo, github_configured: s.configured, models: MODELS, checks: {} };
    if (!gh) return out;
    try {
      const repo = await gh.get(`/repos/{repo}`);
      out.checks.repo = { ok: true, default_branch: repo.default_branch, private: repo.private };
    } catch (err) {
      out.checks.repo = { ok: false, error: String(err.message || err) };
      return out;
    }
    try {
      await gh.get(`/repos/{repo}/contents/.github/workflows/claude.yml`);
      out.checks.workflow = { ok: true };
    } catch (err) {
      out.checks.workflow = { ok: false, error: err.status === 404 ? "The Claude workflow file is not in the repository yet." : String(err.message || err) };
    }
    return out;
  });

  app.get("/requests", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const rows = db.prepare(`SELECT * FROM developer_requests WHERE site_code = ? ORDER BY id DESC LIMIT 50`).all(getSiteCode(req));
    return { ok: true, rows: rows.map(publicRow) };
  });

  // Refresh the requests still in progress (at most 8, newest first).
  app.post("/requests/refresh", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const { s, gh } = client();
    if (!gh) return { ok: false, error: "GitHub is not connected on the server." };
    const open = db.prepare(`SELECT * FROM developer_requests WHERE site_code = ? AND issue_number IS NOT NULL ORDER BY id DESC LIMIT 30`)
      .all(getSiteCode(req)).filter((row) => !FINISHED.has(row.status)).slice(0, 8);
    for (const row of open) await refresh(row, gh, s.repo);
    const rows = db.prepare(`SELECT * FROM developer_requests WHERE site_code = ? ORDER BY id DESC LIMIT 50`).all(getSiteCode(req));
    return { ok: true, refreshed: open.length, rows: rows.map(publicRow) };
  });

  app.get("/requests/:id", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    let row = rowById(Number(req.params.id), getSiteCode(req));
    if (!row) return reply.code(404).send({ ok: false, error: "Request not found." });
    const { s, gh } = client();
    if (gh && req.query?.refresh !== "0") row = await refresh(row, gh, s.repo);
    return { ok: true, row: publicRow(row) };
  });

  app.post("/requests", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const input = validateRequest(req.body || {});
    if (input.error) return reply.code(400).send({ ok: false, error: input.error });
    const docs = await readAttachments(req.body?.attachments);
    if (docs.error) return reply.code(400).send({ ok: false, error: docs.error });
    const { s, gh } = client();
    if (!gh) return reply.code(409).send({ ok: false, error: "GitHub is not connected on the server yet. See the setup checklist." });
    const site = getSiteCode(req);
    const createdBy = getUser(req) || "admin";
    const attachmentNames = docs.attachments.map((d) => d.name);
    const id = Number(db.prepare(`
      INSERT INTO developer_requests (site_code, title, request_text, priority, model, created_by, attachments_json)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `).run(site, input.title, input.request_text, input.priority, input.model, createdBy, attachmentNames.length ? JSON.stringify(attachmentNames) : null).lastInsertRowid);
    try {
      const issue = await gh.post(`/repos/{repo}/issues`, buildIssue({ ...input, created_by: createdBy, attachments: docs.attachments }));
      db.prepare(`UPDATE developer_requests SET issue_number = ?, issue_url = ?, status = 'waiting', checked_at = datetime('now') WHERE id = ?`)
        .run(issue.number, issue.html_url, id);
    } catch (err) {
      db.prepare(`UPDATE developer_requests SET status = 'failed', status_json = ? WHERE id = ?`)
        .run(JSON.stringify({ refresh_error: String(err.message || err).slice(0, 300) }), id);
      return reply.code(424).send({ ok: false, error: `GitHub did not accept the request: ${err.message || err}`, row: publicRow(rowById(id, site)) });
    }
    return { ok: true, row: publicRow(rowById(id, site)), repo: s.repo };
  });

  // Send Claude a follow-up: on the pull request once there is one, else on the issue.
  app.post("/requests/:id/reply", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const row = rowById(Number(req.params.id), getSiteCode(req));
    if (!row?.issue_number) return reply.code(404).send({ ok: false, error: "Request not found." });
    const message = String(req.body?.message || "").trim().slice(0, 6000);
    if (message.length < 3) return reply.code(400).send({ ok: false, error: "Write a message for Claude first." });
    const { s, gh } = client();
    if (!gh) return reply.code(409).send({ ok: false, error: "GitHub is not connected on the server." });
    const detail = row.status_json ? JSON.parse(row.status_json) : {};
    const target = detail.pull_request?.state === "open" ? detail.pull_request.number : row.issue_number;
    await gh.post(`/repos/{repo}/issues/${target}/comments`, { body: `@claude ${message}\n\n_Sent by ${getUser(req) || "admin"} from the IronLog Developer tab._` });
    const updated = await refresh(row, gh, s.repo);
    return { ok: true, posted_on: target, row: publicRow(updated) };
  });

  // Claude pushes a branch; this opens the pull request you review and merge on GitHub.
  app.post("/requests/:id/open-pr", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    let row = rowById(Number(req.params.id), getSiteCode(req));
    if (!row?.issue_number) return reply.code(404).send({ ok: false, error: "Request not found." });
    const { s, gh } = client();
    if (!gh) return reply.code(409).send({ ok: false, error: "GitHub is not connected on the server." });
    row = await refresh(row, gh, s.repo);
    const detail = row.status_json ? JSON.parse(row.status_json) : {};
    if (detail.pull_request) return { ok: true, row: publicRow(row), already: true };
    if (!detail.branch?.name) return reply.code(409).send({ ok: false, error: "Claude has not pushed any code yet." });
    const repo = await gh.get(`/repos/{repo}`);
    const pr = await gh.post(`/repos/{repo}/pulls`, {
      title: row.title,
      head: detail.branch.name,
      base: repo.default_branch || "main",
      body: [
        `Closes #${row.issue_number}`,
        "",
        `Requested by ${row.created_by} from the IronLog Developer tab (${MODELS[row.model]?.label || row.model}).`,
        "",
        "Claude's summary is on the issue. Review the changes, check the tests are green, then merge.",
      ].join("\n"),
    });
    // Same labels as the issue, so an @claude follow-up on the PR uses the same model.
    await gh.post(`/repos/{repo}/issues/${pr.number}/labels`, { labels: ["ironlog-request", `model:${row.model}`] }).catch(() => {});
    row = await refresh(row, gh, s.repo);
    return { ok: true, row: publicRow(row) };
  });

  // Stop a request: close the issue so Claude is no longer asked to work on it.
  app.post("/requests/:id/close", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const row = rowById(Number(req.params.id), getSiteCode(req));
    if (!row) return reply.code(404).send({ ok: false, error: "Request not found." });
    const { s, gh } = client();
    if (gh && row.issue_number) {
      const reason = String(req.body?.reason || "").trim().slice(0, 500);
      if (reason) await gh.post(`/repos/{repo}/issues/${row.issue_number}/comments`, { body: `Closed from IronLog: ${reason}` });
      await gh.patch(`/repos/{repo}/issues/${row.issue_number}`, { state: "closed", state_reason: "not_planned" });
      const updated = await refresh(row, gh, s.repo);
      return { ok: true, row: publicRow(updated) };
    }
    db.prepare(`UPDATE developer_requests SET status = 'closed' WHERE id = ?`).run(row.id);
    return { ok: true, row: publicRow(rowById(row.id, getSiteCode(req))) };
  });
}
