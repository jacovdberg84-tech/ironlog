// IRONLOG/web/app/developer.js — Developer tab: requests to Claude.
// Part of the main app; index.html loads these files in order and they share one global scope.
//
// A request becomes a GitHub issue for Claude (see api/utils/developerClaude.js).
// This tab follows each one through: Claude working → code ready → pull
// request → checks → merged → live, and lets the admin reply to Claude.

let developerJobs = [];
let developerSelectedId = 0;
let developerBusy = false;
let developerTimer = null;
let developerSetupOpen = false;

const DEV_ACTIVE = new Set(["sending", "waiting", "working", "code_ready", "in_review", "checks_running", "checks_failing", "merged", "deploying", "needs_you"]);
const DEV_TONE = {
  waiting: "info", working: "info", sending: "info", checks_running: "info", deploying: "info", merged: "info",
  code_ready: "ready", in_review: "ready", needs_you: "warn",
  deployed: "good",
  not_started: "bad", checks_failing: "bad", deploy_failed: "bad", failed: "bad",
  closed: "muted",
};
// The pipeline shown for every request, and the status that completes each step.
const DEV_STEPS = [
  { key: "sent", label: "Sent", done: () => true },
  { key: "working", label: "Claude working", done: (d) => Boolean(d?.claude_note || d?.branch) },
  { key: "code", label: "Code ready", done: (d) => Boolean(d?.branch) },
  { key: "pr", label: "Pull request", done: (d) => Boolean(d?.pull_request) },
  { key: "checks", label: "Checks", done: (d) => d?.checks?.state === "passing" || Boolean(d?.pull_request?.merged_at) },
  { key: "merged", label: "Merged", done: (d) => Boolean(d?.pull_request?.merged_at) },
  { key: "live", label: "Live", done: (d) => d?.deploy?.conclusion === "success" },
];

function developerMsg(text, isError = false) {
  const el = qs("devMsg");
  if (!el) return;
  el.className = isError ? "message-error" : "muted";
  el.textContent = String(text || "");
}

async function developerRequest(path, opts = {}) {
  return fetchJson(`${API}/api/developer${path}`, opts);
}

function developerPostJson(path, body) {
  return developerRequest(path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body || {}),
  });
}

function devAgo(ts) {
  const t = Date.parse(String(ts || "").replace(" ", "T") + (String(ts || "").includes("T") ? "" : "Z"));
  if (!Number.isFinite(t)) return "";
  const m = Math.max(0, Math.round((Date.now() - t) / 60000));
  if (m < 1) return "just now";
  if (m < 60) return `${m} min ago`;
  if (m < 1440) return `${Math.round(m / 60)} h ago`;
  return new Date(t).toLocaleDateString(undefined, { day: "numeric", month: "short" });
}

// Builds the setup checklist from the /status answer. "ready" means GitHub is
// connected and Claude's workflow is in the repository.
function developerSetupState(data) {
  const items = [];
  items.push(data.github_configured
    ? { ok: data.checks?.repo?.ok !== false, text: data.checks?.repo?.ok === false ? `GitHub key is set but GitHub refused it: ${data.checks.repo.error}` : `Connected to GitHub (${data.repo})` }
    : { ok: false, text: "Connect GitHub: add a repository secret IRONLOG_GITHUB_TOKEN (a fine-grained token for the IronLog repo with Issues and Pull requests read/write; Contents, Actions and Checks read). The next deploy passes it to the server." });
  if (data.github_configured && data.checks?.repo?.ok) {
    items.push(data.checks?.workflow?.ok
      ? { ok: true, text: "Claude's GitHub workflow is in the repository" }
      : { ok: false, text: data.checks?.workflow?.error || "Claude's GitHub workflow is missing" });
  }
  items.push({ ok: null, text: "Give Claude a key on GitHub: install the Claude GitHub app on the repository, then add CLAUDE_CODE_OAUTH_TOKEN (your Claude subscription) or ANTHROPIC_API_KEY (pay as you go) as a repository secret." });
  return { items, ready: items.every((i) => i.ok !== false) };
}

// When everything is ready the orange box shrinks to one green line; clicking it opens the checklist.
function showDeveloperSetup(ready) {
  const box = qs("devSetup");
  const line = qs("devReady");
  if (line) {
    line.hidden = !ready;
    line.setAttribute("aria-expanded", String(developerSetupOpen));
  }
  if (box) box.hidden = ready && !developerSetupOpen;
}

function toggleDeveloperSetup() {
  developerSetupOpen = !developerSetupOpen;
  showDeveloperSetup(!qs("devReady")?.hidden);
}

async function loadDeveloperSetup() {
  const list = qs("devChecklist");
  if (!qs("devSetup") || !list) return;
  try {
    const data = await developerRequest("/status");
    const { items, ready } = developerSetupState(data);
    list.innerHTML = items.map((i) => `<li class="${i.ok === true ? "ok" : i.ok === false ? "todo" : "info"}">${i.ok === true ? "✓" : i.ok === false ? "•" : "ℹ"} ${escapeHtml(i.text)}</li>`).join("");
    showDeveloperSetup(ready);
  } catch (err) {
    list.innerHTML = `<li class="todo">• ${escapeHtml(err.message || String(err))}</li>`;
    showDeveloperSetup(false);
  }
}

function renderDeveloperJobs() {
  const list = qs("devJobList");
  if (!list) return;
  if (!developerJobs.length) {
    list.innerHTML = `<p class="muted">No requests yet. Your first one will show up here.</p>`;
    return;
  }
  list.innerHTML = developerJobs.map((job) => {
    const id = Number(job.id || 0);
    const active = id === developerSelectedId ? " active" : "";
    return `<button type="button" class="dev-job${active}" data-dev-open="${id}">
      <span class="dev-job-top"><strong>${escapeHtml(job.title || "Untitled")}</strong><span class="dev-chip tone-${DEV_TONE[job.status] || "info"}">${escapeHtml(job.status_text || job.status)}</span></span>
      <small>${job.issue_number ? `#${job.issue_number} · ` : ""}${escapeHtml(job.model_label || "")} · ${escapeHtml(job.priority || "medium")} · ${escapeHtml(devAgo(job.created_at))}</small>
    </button>`;
  }).join("");
}

function devLink(url, text) {
  return url ? `<a href="${escapeHtml(url)}" target="_blank" rel="noopener">${escapeHtml(text)} ↗</a>` : "";
}

function renderDeveloperDetail(job) {
  const view = qs("devView");
  const empty = qs("devEmpty");
  if (!view || !job) return;
  empty.hidden = true;
  view.hidden = false;
  const d = job.detail || {};
  qs("devDetailTitle").textContent = job.title || "Request";
  qs("devDetailMeta").textContent = [
    job.issue_number ? `Issue #${job.issue_number}` : null,
    job.model_label,
    `asked by ${job.created_by}`,
    job.checked_at ? `checked ${devAgo(job.checked_at)}` : null,
  ].filter(Boolean).join(" · ");
  const chip = qs("devDetailChip");
  chip.className = `dev-chip tone-${DEV_TONE[job.status] || "info"}`;
  chip.textContent = job.status_text || job.status;

  const failedAt = job.status === "checks_failing" ? "checks" : job.status === "deploy_failed" ? "live" : null;
  let current = true;
  qs("devSteps").innerHTML = DEV_STEPS.map((step) => {
    const done = job.status !== "failed" && step.done(d);
    const cls = failedAt === step.key ? "bad" : done ? "done" : current ? "now" : "";
    if (!done) current = false;
    return `<li class="${cls}"><span></span>${escapeHtml(step.label)}</li>`;
  }).join("");

  qs("devLinks").innerHTML = [
    devLink(job.issue_url, "Issue"),
    devLink(d.branch?.url, "Claude's branch"),
    devLink(d.pull_request?.url, `Pull request #${d.pull_request?.number || ""}`),
    devLink(d.deploy?.url, "Deploy run"),
  ].filter(Boolean).join("");

  const actions = [];
  if (job.status === "code_ready" && !d.pull_request) {
    actions.push(`<button type="button" class="btn-primary" data-dev-action="open-pr">Open pull request for review</button>`);
  }
  if (d.pull_request && !d.pull_request.merged_at && d.pull_request.state === "open") {
    actions.push(`<a class="btn btn-primary" href="${escapeHtml(d.pull_request.url)}/files" target="_blank" rel="noopener">Review &amp; merge on GitHub ↗</a>`);
  }
  qs("devActions").innerHTML = actions.join(" ");

  const hints = {
    waiting: "Claude has the request and should start within a minute or two.",
    not_started: "Claude has not picked this up. Check the setup list above: the Claude GitHub app, its key, and the workflow file.",
    checks_failing: `These checks failed: ${(d.checks?.failed || []).join(", ") || "see the pull request"}. Reply below and ask Claude to fix them.`,
    deploy_failed: "The merge went through but the deploy failed. Open the deploy run for the details.",
    failed: d.refresh_error ? `GitHub said: ${d.refresh_error}` : "This request could not be sent to GitHub.",
  };
  qs("devNote").textContent = d.claude_note || hints[job.status] || "No note from Claude yet.";
  if (hints[job.status] && d.claude_note) qs("devNote").textContent = `${hints[job.status]}\n\n${d.claude_note}`;

  const files = d.changes?.files || [];
  qs("devFiles").innerHTML = files.length
    ? `<details class="dev-files" open><summary>${files.length} file${files.length === 1 ? "" : "s"} changed${d.changes.commits ? ` · ${d.changes.commits} commit${d.changes.commits === 1 ? "" : "s"}` : ""}</summary><ul>${
      files.map((f) => `<li><code>${escapeHtml(f.file)}</code> <span class="dev-add">+${Number(f.added || 0)}</span> <span class="dev-del">−${Number(f.removed || 0)}</span></li>`).join("")
    }</ul></details>`
    : "";
  const finished = ["deployed", "closed", "failed"].includes(job.status);
  qs("devReplyText").disabled = finished || !job.issue_number;
  qs("devReplyBtn").disabled = finished || !job.issue_number;
  qs("devCloseBtn").hidden = finished || Boolean(d.pull_request?.merged_at);
  qs("devRequestText").textContent = job.request_text || "";
}

async function openDeveloperJob(id, { refresh = true } = {}) {
  const jobId = Number(id || 0);
  if (!jobId) return;
  developerSelectedId = jobId;
  renderDeveloperJobs();
  const cached = developerJobs.find((j) => Number(j.id) === jobId);
  if (cached) renderDeveloperDetail(cached);
  const data = await developerRequest(`/requests/${jobId}${refresh ? "" : "?refresh=0"}`);
  if (data?.row) {
    developerJobs = developerJobs.map((j) => (Number(j.id) === jobId ? data.row : j));
    renderDeveloperJobs();
    if (developerSelectedId === jobId) renderDeveloperDetail(data.row);
  }
}

async function loadDeveloperJobs({ refresh = false } = {}) {
  const data = refresh ? await developerPostJson("/requests/refresh", {}).catch(() => developerRequest("/requests")) : await developerRequest("/requests");
  developerJobs = Array.isArray(data?.rows) ? data.rows : developerJobs;
  renderDeveloperJobs();
  const selected = developerJobs.find((j) => Number(j.id) === developerSelectedId);
  if (selected) renderDeveloperDetail(selected);
}

function devSelectedModel() {
  return document.querySelector('#tab-developer input[name="devModel"]:checked')?.value || "opus";
}

async function createDeveloperJob() {
  if (developerBusy) return;
  const title = String(qs("devTitle")?.value || "").trim();
  const request = String(qs("devRequest")?.value || "").trim();
  if (!title) return developerMsg("Give the change a short title.", true);
  if (request.length < 10) return developerMsg("Describe the change in a sentence or two.", true);
  developerBusy = true;
  qs("devCreateBtn").disabled = true;
  developerMsg("Sending to Claude…");
  try {
    const res = await developerPostJson("/requests", {
      title,
      request_text: request,
      priority: qs("devPriority")?.value || "medium",
      model: devSelectedModel(),
    });
    qs("devTitle").value = "";
    qs("devRequest").value = "";
    developerSelectedId = Number(res?.row?.id || 0);
    developerMsg(`Sent. Claude is on it as issue #${res?.row?.issue_number}. This tab updates by itself.`);
    await loadDeveloperJobs();
    startDeveloperPolling();
  } catch (err) {
    developerMsg(err.message || String(err), true);
    await loadDeveloperJobs().catch(() => {});
  } finally {
    developerBusy = false;
    qs("devCreateBtn").disabled = false;
  }
}

async function developerAction(action) {
  const id = developerSelectedId;
  if (!id || developerBusy) return;
  developerBusy = true;
  try {
    if (action === "open-pr") {
      developerMsg("Opening the pull request…");
      await developerPostJson(`/requests/${id}/open-pr`, {});
      developerMsg("Pull request opened. Review it on GitHub, then merge to put it live.");
    } else if (action === "reply") {
      const message = String(qs("devReplyText")?.value || "").trim();
      if (message.length < 3) return developerMsg("Write a message for Claude first.", true);
      developerMsg("Sending your reply to Claude…");
      await developerPostJson(`/requests/${id}/reply`, { message });
      qs("devReplyText").value = "";
      developerMsg("Reply sent. Claude will pick it up shortly.");
      startDeveloperPolling();
    } else if (action === "close") {
      const reason = prompt("Close this request? Claude will stop working on it.\n\nReason (optional):", "");
      if (reason === null) return;
      await developerPostJson(`/requests/${id}/close`, { reason });
      developerMsg("Request closed.");
    }
    await openDeveloperJob(id, { refresh: false });
    await loadDeveloperJobs();
  } catch (err) {
    developerMsg(err.message || String(err), true);
  } finally {
    developerBusy = false;
  }
}

function developerTabVisible() {
  return !document.hidden && qs("tab-developer")?.classList.contains("show");
}

function startDeveloperPolling() {
  clearInterval(developerTimer);
  developerTimer = setInterval(() => {
    if (!developerTabVisible()) return;
    if (!developerJobs.some((j) => DEV_ACTIVE.has(j.status))) return clearInterval(developerTimer);
    loadDeveloperJobs({ refresh: true }).catch(() => {});
  }, 30000);
}

function refreshDeveloperTab() {
  loadDeveloperJobs({ refresh: true })
    .then(() => {
      loadDeveloperSetup();
      if (!developerSelectedId && developerJobs[0]) openDeveloperJob(developerJobs[0].id, { refresh: false }).catch(() => {});
      startDeveloperPolling();
    })
    .catch((err) => {
      developerMsg(err.message || String(err), true);
      loadDeveloperSetup();
    });
}

function initDeveloperTab() {
  if (window.__developerTabReady) return;
  window.__developerTabReady = true;
  qs("devCreateBtn")?.addEventListener("click", () => createDeveloperJob());
  qs("devRefreshBtn")?.addEventListener("click", () => refreshDeveloperTab());
  qs("devReady")?.addEventListener("click", () => toggleDeveloperSetup());
  qs("devJobList")?.addEventListener("click", (event) => {
    const button = event.target?.closest?.("[data-dev-open]");
    if (button) openDeveloperJob(button.getAttribute("data-dev-open")).catch((err) => developerMsg(err.message || String(err), true));
  });
  qs("devActions")?.addEventListener("click", (event) => {
    const button = event.target?.closest?.("[data-dev-action]");
    if (button && !button.disabled) developerAction(button.getAttribute("data-dev-action"));
  });
  qs("devReplyBtn")?.addEventListener("click", () => developerAction("reply"));
  qs("devCloseBtn")?.addEventListener("click", () => developerAction("close"));
}
