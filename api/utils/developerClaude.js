// IRONLOG/api/utils/developerClaude.js — Developer tab: Claude as IronLog's coder.
//
// A request written in IronLog becomes a GitHub issue that mentions @claude.
// The Claude GitHub Action (.github/workflows/claude.yml) does the work on a
// branch; IronLog follows it from there: Claude's notes, the branch, the pull
// request, its checks, the merge and the deploy. Nothing here edits the live
// server: every change goes through a pull request and the normal deploy.
//
// Server settings (api/.env):
//   IRONLOG_GITHUB_TOKEN  fine-grained token for the IronLog repo
//                         (Issues: read/write, Pull requests: read/write,
//                          Contents: read, Actions: read, Checks: read)
//   IRONLOG_GITHUB_REPO   owner/name, default jacovdberg84-tech/ironlog
import { db as defaultDb } from "../db/client.js";

export const DEFAULT_REPO = "jacovdberg84-tech/ironlog";
export const MODELS = {
  opus: { label: "Claude Opus", hint: "Best for bigger or trickier jobs" },
  sonnet: { label: "Claude Sonnet", hint: "Quicker and cheaper for small changes" },
};
const PRIORITIES = ["low", "medium", "high", "urgent"];
const GH_API = "https://api.github.com";

export function githubSettings(env = process.env) {
  const token = String(env.IRONLOG_GITHUB_TOKEN || "").trim();
  const repo = String(env.IRONLOG_GITHUB_REPO || DEFAULT_REPO).trim();
  return { token, repo, configured: Boolean(token && /^[\w.-]+\/[\w.-]+$/.test(repo)) };
}

export function ensureDeveloperTables(dbConn = defaultDb) {
  dbConn.exec(`
    CREATE TABLE IF NOT EXISTS developer_requests (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL DEFAULT 'main',
      title TEXT NOT NULL,
      request_text TEXT NOT NULL,
      priority TEXT NOT NULL DEFAULT 'medium',
      model TEXT NOT NULL DEFAULT 'opus',
      created_by TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      issue_number INTEGER,
      issue_url TEXT,
      status TEXT NOT NULL DEFAULT 'sending',
      status_json TEXT,
      checked_at TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_developer_requests_issue ON developer_requests(issue_number);
  `);
}

/** A small GitHub REST client; fetch is injectable for tests. */
export function githubClient({ token, repo, fetchImpl = globalThis.fetch }) {
  async function call(method, path, body) {
    const res = await fetchImpl(`${GH_API}${path.replace("{repo}", repo)}`, {
      method,
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${token}`,
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "IronLog-Developer",
        ...(body ? { "Content-Type": "application/json" } : {}),
      },
      body: body ? JSON.stringify(body) : undefined,
    });
    const text = await res.text();
    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch { data = { raw: text }; }
    if (!res.ok) {
      const message = data?.message || `GitHub returned HTTP ${res.status}`;
      throw Object.assign(new Error(message), { status: res.status });
    }
    return data;
  }
  return {
    get: (path) => call("GET", path),
    post: (path, body) => call("POST", path, body),
    patch: (path, body) => call("PATCH", path, body),
  };
}

function clean(value, max) {
  return String(value ?? "").replace(/\r\n/g, "\n").trim().slice(0, max);
}

/** The GitHub issue Claude works from. */
export function buildIssue({ title, request_text, priority, model, created_by }) {
  const body = [
    "@claude",
    "",
    "## What we need",
    request_text,
    "",
    "## Details",
    `- Asked by: ${created_by} (from the IronLog Developer tab)`,
    `- Priority: ${priority}`,
    `- Model: ${MODELS[model]?.label || model}`,
    "",
    "## How to work",
    "- Follow CLAUDE.md in the repository root.",
    "- Work on a branch, keep the change focused on this request, and add or update tests in api/test.",
    "- Run `npm test` in api/ before you finish and say whether it passed.",
    "- When you are done, summarise in plain language what changed and anything that still needs a person to check.",
    "- If something is unclear, ask in this issue instead of guessing.",
  ].join("\n");
  return {
    title: `[IronLog] ${title}`,
    body,
    labels: ["ironlog-request", `model:${model}`, `priority:${priority}`],
  };
}

export function validateRequest(input = {}) {
  const title = clean(input.title, 120);
  const request_text = clean(input.request_text ?? input.request, 8000);
  const priority = PRIORITIES.includes(String(input.priority || "").toLowerCase()) ? String(input.priority).toLowerCase() : "medium";
  const model = MODELS[String(input.model || "").toLowerCase()] ? String(input.model).toLowerCase() : "opus";
  if (!title) return { error: "Give the change a short title." };
  if (request_text.length < 10) return { error: "Describe the change in a sentence or two." };
  return { title, request_text, priority, model };
}

const isClaude = (user) => /claude/i.test(String(user?.login || "")) && (user?.type === "Bot" || /\[bot\]$/.test(String(user?.login || "")));

/** Plain text from Claude's markdown/HTML comment, for the IronLog view. */
export function plainComment(body, max = 4000) {
  return String(body || "")
    .replace(/<img[^>]*>/gi, "")
    .replace(/<\/?(details|summary|br|p|div|span|sub|sup)[^>]*>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/\[([^\]]+)\]\((https?:[^)]+)\)/g, "$1")
    .replace(/\n{3,}/g, "\n\n")
    .trim()
    .slice(0, max);
}

function checksSummary(runs = []) {
  const real = runs.filter((r) => !/claude/i.test(String(r.name || "")));
  if (!real.length) return { state: "none", total: 0, failed: [] };
  const failed = real.filter((r) => ["failure", "timed_out", "cancelled", "action_required"].includes(r.conclusion)).map((r) => r.name);
  const pending = real.some((r) => r.status !== "completed");
  return { state: failed.length ? "failing" : pending ? "running" : "passing", total: real.length, failed };
}

/**
 * Where a request stands, worked out from what GitHub reports. Pure: given
 * the issue, comments, Claude's branch, the pull request, checks and deploy.
 */
export function deriveStatus({ issue, comments = [], branch = null, compare = null, pr = null, checks = null, deploy = null, now = Date.now() }) {
  const claudeNotes = comments.filter((c) => isClaude(c.user));
  const latest = claudeNotes[claudeNotes.length - 1] || null;
  const latestText = latest ? plainComment(latest.body) : "";
  const working = latest ? /is working|working on|in progress|\[ \]/i.test(latestText) && !/finished|completed|done/i.test(latestText.slice(0, 200)) : false;
  const ageMin = issue?.created_at ? (now - Date.parse(issue.created_at)) / 60000 : 0;

  let status = "waiting";
  if (pr?.merged_at) {
    status = !deploy ? "merged" : deploy.status !== "completed" ? "deploying" : deploy.conclusion === "success" ? "deployed" : "deploy_failed";
  } else if (pr && pr.state === "open") {
    status = checks?.state === "failing" ? "checks_failing" : checks?.state === "running" ? "checks_running" : "in_review";
  } else if (issue?.state === "closed") {
    status = "closed";
  } else if (branch && (compare?.ahead_by ?? 1) > 0) {
    status = working ? "working" : "code_ready";
  } else if (latest) {
    status = working ? "working" : "needs_you";
  } else if (ageMin > 15) {
    status = "not_started";
  }
  return {
    status,
    claude_note: latestText,
    claude_note_at: latest?.updated_at || latest?.created_at || null,
    branch: branch ? { name: branch.name, url: branch.url } : null,
    changes: compare ? {
      commits: compare.ahead_by ?? null,
      files: (compare.files || []).slice(0, 30).map((f) => ({ file: f.filename, status: f.status, added: f.additions, removed: f.deletions })),
    } : null,
    pull_request: pr ? { number: pr.number, url: pr.html_url, state: pr.state, merged_at: pr.merged_at || null, draft: Boolean(pr.draft) } : null,
    checks,
    deploy: deploy ? { status: deploy.status, conclusion: deploy.conclusion, url: deploy.html_url } : null,
  };
}

/** Fetch everything GitHub knows about one request and work out its status. */
export async function inspectRequest(client, { repo, issueNumber }) {
  const [owner] = repo.split("/");
  const issue = await client.get(`/repos/{repo}/issues/${issueNumber}`);
  const comments = await client.get(`/repos/{repo}/issues/${issueNumber}/comments?per_page=100`).catch(() => []);
  // Claude's action names its branch claude/issue-<number>-<timestamp>.
  const refs = await client.get(`/repos/{repo}/git/matching-refs/heads/claude/issue-${issueNumber}-`).catch(() => []);
  const ref = Array.isArray(refs) && refs.length ? refs[refs.length - 1] : null;
  const branchName = ref ? String(ref.ref).replace(/^refs\/heads\//, "") : null;
  const branch = branchName ? { name: branchName, url: `https://github.com/${repo}/tree/${encodeURIComponent(branchName)}`, sha: ref.object?.sha } : null;
  let compare = null;
  let pr = null;
  if (branch) {
    compare = await client.get(`/repos/{repo}/compare/main...${encodeURIComponent(branch.name)}`).catch(() => null);
    const prs = await client.get(`/repos/{repo}/pulls?state=all&head=${encodeURIComponent(`${owner}:${branch.name}`)}`).catch(() => []);
    pr = Array.isArray(prs) && prs.length ? prs[0] : null;
  }
  let checks = null;
  if (pr && !pr.merged_at) {
    const runs = await client.get(`/repos/{repo}/commits/${pr.head?.sha}/check-runs?per_page=50`).catch(() => null);
    checks = runs ? checksSummary(runs.check_runs || []) : null;
  }
  let deploy = null;
  if (pr?.merged_at && pr.merge_commit_sha) {
    const runs = await client.get(`/repos/{repo}/actions/workflows/deploy.yml/runs?head_sha=${pr.merge_commit_sha}&per_page=1`).catch(() => null);
    deploy = runs?.workflow_runs?.[0] || null;
  }
  return {
    issue: { number: issue.number, url: issue.html_url, state: issue.state },
    branch_sha: branch?.sha || null,
    ...deriveStatus({ issue, comments, branch, compare, pr, checks, deploy }),
  };
}

export const STATUS_TEXT = {
  sending: "Sending to GitHub",
  waiting: "Waiting for Claude to start",
  not_started: "Claude has not started",
  working: "Claude is working",
  needs_you: "Claude replied",
  code_ready: "Code ready for a pull request",
  in_review: "Pull request ready for review",
  checks_running: "Checks running",
  checks_failing: "Checks failing",
  merged: "Merged",
  deploying: "Deploying",
  deployed: "Live",
  deploy_failed: "Deploy failed",
  closed: "Closed",
  failed: "Could not send",
};
