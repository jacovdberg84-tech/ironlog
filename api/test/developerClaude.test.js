import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-developer-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_GITHUB_TOKEN = "test-token";
process.env.IRONLOG_GITHUB_REPO = "acme/ironlog";
const { db } = await import("../db/client.js");
const { default: developerRoutes } = await import("../routes/developer.routes.js");
const { default: aiEngineerRoutes } = await import("../routes/aiEngineer.routes.js");
const dc = await import("../utils/developerClaude.js");
process.on("exit", () => { try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

// A small stand-in for the GitHub REST API, enough for one request's journey.
const gh = {
  calls: [],
  issues: new Map(),
  comments: new Map(),
  refs: [],
  pulls: [],
  checkRuns: [],
  deployRuns: [],
};
async function fakeFetch(url, opts = {}) {
  const u = new URL(url);
  const p = u.pathname.replace("/repos/acme/ironlog", "");
  const method = opts.method || "GET";
  const body = opts.body ? JSON.parse(opts.body) : null;
  gh.calls.push({ method, path: p + u.search, body, auth: opts.headers?.Authorization });
  const json = (data, status = 200) => ({ ok: status < 300, status, text: async () => JSON.stringify(data) });
  if (method === "GET" && p === "") return json({ default_branch: "main", private: true });
  if (method === "GET" && p === "/contents/.github/workflows/claude.yml") return json({ name: "claude.yml" });
  if (method === "POST" && p === "/issues") {
    const n = 100 + gh.issues.size;
    gh.issues.set(n, { number: n, html_url: `https://github.com/acme/ironlog/issues/${n}`, state: "open", created_at: new Date().toISOString(), ...body });
    gh.comments.set(n, []);
    return json(gh.issues.get(n), 201);
  }
  let m = p.match(/^\/issues\/(\d+)$/);
  if (m && method === "GET") return json(gh.issues.get(Number(m[1])));
  if (m && method === "PATCH") { Object.assign(gh.issues.get(Number(m[1])), body); return json(gh.issues.get(Number(m[1]))); }
  m = p.match(/^\/issues\/(\d+)\/comments$/);
  if (m && method === "GET") return json(gh.comments.get(Number(m[1])) || []);
  if (m && method === "POST") { (gh.comments.get(Number(m[1])) || []).push({ body: body.body, user: { login: "admin", type: "User" } }); return json({}, 201); }
  if (p.match(/^\/issues\/\d+\/labels$/)) return json([]);
  m = p.match(/^\/git\/matching-refs\/heads\/claude\/issue-(\d+)-$/);
  if (m) return json(gh.refs.filter((r) => r.ref.includes(`/claude/issue-${m[1]}-`)));
  if (p.startsWith("/compare/")) return json({ ahead_by: 2, files: [{ filename: "web/app/fuel-lube.js", status: "modified", additions: 12, deletions: 3 }] });
  if (method === "GET" && p === "/pulls") return json(gh.pulls.filter((x) => u.searchParams.get("head") === `acme:${x.head_ref}`));
  if (method === "POST" && p === "/pulls") {
    const pr = { number: 7, html_url: "https://github.com/acme/ironlog/pull/7", state: "open", merged_at: null, ...body, head_ref: body.head, head: { sha: "abc" } };
    gh.pulls.push(pr);
    return json(pr, 201);
  }
  if (p.startsWith("/commits/abc/check-runs")) return json({ check_runs: gh.checkRuns });
  if (p.startsWith("/actions/workflows/deploy.yml/runs")) return json({ workflow_runs: gh.deployRuns });
  return json({ message: `Not found: ${method} ${p}` }, 404);
}

const app = Fastify();
await app.register(developerRoutes, { prefix: "/api/developer", fetchImpl: fakeFetch });
await app.register(aiEngineerRoutes, { prefix: "/api/ai-engineer" });
const admin = { "x-user-name": "jaco", "x-user-role": "admin", "x-user-roles": "admin" };
const call = (method, url, payload, headers = admin) => app.inject({ method, url: `/api/developer${url}`, headers, payload });
const claudeSays = (n, body) => gh.comments.get(n).push({ body, user: { login: "claude[bot]", type: "Bot" }, created_at: new Date().toISOString() });

test("only administrators can use the Developer tab", async () => {
  const res = await call("GET", "/requests", undefined, { "x-user-role": "workshop_admin", "x-user-roles": "workshop_admin" });
  assert.equal(res.statusCode, 403);
});

test("setup status checks the GitHub key and the Claude workflow", async () => {
  const res = await call("GET", "/status");
  assert.equal(res.json().github_configured, true);
  assert.equal(res.json().checks.repo.ok, true);
  assert.equal(res.json().checks.workflow.ok, true);
  assert.equal(JSON.stringify(res.json()).includes("test-token"), false, "the token is never sent to the browser");
});

test("a request travels from IronLog to Claude and back to a live deploy", async () => {
  assert.equal((await call("POST", "/requests", { title: "x", request_text: "short" })).statusCode, 400);

  // 1. Sent: a GitHub issue that mentions @claude, labelled with the model.
  const created = await call("POST", "/requests", { title: "Fuel cost column", request_text: "Add a fuel cost column to the Daily Log grid.", priority: "high", model: "sonnet" });
  assert.equal(created.statusCode, 200);
  const row = created.json().row;
  assert.equal(row.issue_number, 100);
  assert.equal(row.status, "waiting");
  const issue = gh.calls.find((c) => c.method === "POST" && c.path === "/issues");
  assert.equal(issue.auth, "Bearer test-token");
  assert.match(issue.body.body, /^@claude/);
  assert.match(issue.body.body, /Add a fuel cost column/);
  assert.match(issue.body.body, /CLAUDE\.md/);
  assert.deepEqual(issue.body.labels, ["ironlog-request", "model:sonnet", "priority:high"]);
  assert.equal(issue.body.title, "[IronLog] Fuel cost column");

  // 2. Claude starts work.
  claudeSays(100, "**Claude Code is working…**\n- [ ] Read the Daily Log grid");
  let r = (await call("GET", `/requests/${row.id}`)).json().row;
  assert.equal(r.status, "working");
  assert.match(r.detail.claude_note, /working/);

  // 3. Claude finishes on its branch: code ready.
  gh.comments.get(100)[0].body = "**Claude finished @jaco's task**\nAdded the column and a test. npm test passed.";
  gh.refs = [{ ref: "refs/heads/claude/issue-100-20261013-0900", object: { sha: "abc" } }];
  r = (await call("GET", `/requests/${row.id}`)).json().row;
  assert.equal(r.status, "code_ready");
  assert.equal(r.detail.branch.name, "claude/issue-100-20261013-0900");
  assert.equal(r.detail.changes.files[0].file, "web/app/fuel-lube.js");

  // 4. Open the pull request from IronLog.
  r = (await call("POST", `/requests/${row.id}/open-pr`, {})).json().row;
  const pr = gh.calls.find((c) => c.method === "POST" && c.path === "/pulls");
  assert.equal(pr.body.head, "claude/issue-100-20261013-0900");
  assert.equal(pr.body.base, "main");
  assert.match(pr.body.body, /Closes #100/);
  assert.ok(gh.calls.some((c) => c.path === "/issues/7/labels" && c.body.labels.includes("model:sonnet")));
  assert.equal(r.status, "in_review");

  // 5. A reply goes to the pull request once it exists.
  gh.comments.set(7, []);
  await call("POST", `/requests/${row.id}/reply`, { message: "Please also show it on the PDF." });
  const reply = gh.calls.filter((c) => c.method === "POST" && c.path === "/issues/7/comments").pop();
  assert.match(reply.body.body, /^@claude Please also show it on the PDF\./);

  // 6. Checks fail, then pass.
  gh.checkRuns = [{ name: "deploy", status: "completed", conclusion: "failure" }, { name: "claude", status: "completed", conclusion: "failure" }];
  r = (await call("GET", `/requests/${row.id}`)).json().row;
  assert.equal(r.status, "checks_failing");
  assert.deepEqual(r.detail.checks.failed, ["deploy"], "Claude's own run is not counted as a check");
  gh.checkRuns = [{ name: "tests", status: "completed", conclusion: "success" }];
  assert.equal((await call("GET", `/requests/${row.id}`)).json().row.status, "in_review");

  // 7. Merged on GitHub, deploying, then live.
  Object.assign(gh.pulls[0], { state: "closed", merged_at: "2026-10-13T10:00:00Z", merge_commit_sha: "def" });
  assert.equal((await call("GET", `/requests/${row.id}`)).json().row.status, "merged");
  gh.deployRuns = [{ status: "in_progress", conclusion: null, html_url: "https://github.com/acme/ironlog/actions/runs/1" }];
  assert.equal((await call("GET", `/requests/${row.id}`)).json().row.status, "deploying");
  gh.deployRuns = [{ status: "completed", conclusion: "success", html_url: "https://github.com/acme/ironlog/actions/runs/1" }];
  r = (await call("GET", `/requests/${row.id}`)).json().row;
  assert.equal(r.status, "deployed");
  assert.equal(r.status_text, "Live");
});

test("closing a request closes its issue", async () => {
  const row = (await call("POST", "/requests", { title: "Not needed", request_text: "Change the login page colour." })).json().row;
  const closed = (await call("POST", `/requests/${row.id}/close`, { reason: "Changed my mind" })).json().row;
  assert.equal(closed.status, "closed");
  assert.equal(gh.issues.get(row.issue_number).state, "closed");
});

test("a request Claude never picks up is flagged after 15 minutes", () => {
  const issue = { state: "open", created_at: "2026-10-13T08:00:00Z" };
  assert.equal(dc.deriveStatus({ issue, now: Date.parse("2026-10-13T08:05:00Z") }).status, "waiting");
  assert.equal(dc.deriveStatus({ issue, now: Date.parse("2026-10-13T08:20:00Z") }).status, "not_started");
  // A question from Claude with no code means it is waiting on you.
  const asked = dc.deriveStatus({ issue, comments: [{ body: "Claude finished. Which screen do you mean?", user: { login: "claude[bot]", type: "Bot" } }] });
  assert.equal(asked.status, "needs_you");
});

test("the old server-side merge is switched off so deploys keep working", async () => {
  const res = await app.inject({ method: "POST", url: "/api/ai-engineer/requests/1/merge", headers: admin, payload: { confirm: true } });
  assert.equal(res.statusCode, 410);
  assert.match(res.json().error, /pull request/);
});
