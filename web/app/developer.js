// IRONLOG/web/app/developer.js — Developer tab coding agent.
// Part of the main app; index.html loads these files in order and they share one global scope.

let developerJobs = [];
let developerSelectedId = 0;
let developerBusy = false;

function developerMsg(text, isError = false) {
  const el = qs("devMsg");
  if (!el) return;
  el.className = isError ? "message-error" : "muted";
  el.textContent = String(text || "");
}

function developerStatusLabel(status) {
  const labels = {
    draft: "Draft",
    planned: "Plan ready",
    approved: "Approved",
    executed: "Sandbox ready",
    patch_ready: "Code ready",
    patch_applied: "Applied in sandbox",
    merged: "Merged",
    rejected: "Stopped",
  };
  const key = String(status || "draft").toLowerCase();
  return labels[key] || key;
}

async function developerRequest(path, opts = {}) {
  return fetchJson(`${API}/api/ai-engineer${path}`, opts);
}

async function loadDeveloperAgentStatus() {
  const el = qs("devAgentStatus");
  if (!el) return;
  try {
    const data = await developerRequest("/status");
    const model = data.coder_model || "coder";
    const rounds = Number(data.max_tool_rounds || 0);
    if (data.coder_ready) {
      el.textContent = `Coder ready on ${model}. It can read and search the app, then write a diff (${rounds} tool steps). Live files change only after you merge.`;
    } else {
      el.textContent = "The coder is not configured. Set OPENAI_API_KEY, or an OpenAI-compatible LLM base URL, on the server.";
    }
  } catch (err) {
    el.textContent = `Coder status unavailable: ${err.message || err}`;
  }
}

function renderDeveloperJobs() {
  const list = qs("devJobList");
  if (!list) return;
  if (!developerJobs.length) {
    list.innerHTML = `<p class="muted">No coding jobs yet.</p>`;
    return;
  }
  list.innerHTML = developerJobs.map((job) => {
    const id = Number(job.id || 0);
    const active = id === developerSelectedId ? " active" : "";
    return `<button type="button" class="dev-job${active}" data-dev-open="${id}">
      <strong>#${id} ${escapeHtml(job.title || "Untitled")}</strong>
      <small>${escapeHtml(developerStatusLabel(job.status))} · ${escapeHtml(job.priority || "medium")}</small>
    </button>`;
  }).join("");
}

function pickDeveloperRun(runs) {
  const order = ["merge", "apply_patch", "generate_patch", "execute", "plan"];
  for (const runType of order) {
    const hit = runs.find((run) => String(run?.run_type || "") === runType && run?.details);
    if (hit) return hit;
  }
  return runs[0] || null;
}

function renderDeveloperPreview(row, runs) {
  const run = pickDeveloperRun(runs);
  const details = run?.details || {};
  const proposal = details.proposal || {};
  const lines = [
    `Job #${Number(row?.id || 0)} · ${developerStatusLabel(row?.status)}`,
    `Priority: ${row?.priority || "-"}`,
    `Asked by: ${row?.requested_by || "-"}`,
    row?.approved_by ? `Approved by: ${row.approved_by}` : "",
    row?.rejection_reason ? `Stopped: ${row.rejection_reason}` : "",
    "",
    String(row?.request_text || ""),
    "",
    run ? `Latest step: ${run.run_type} (${run.status})` : "No steps yet.",
    run?.summary ? String(run.summary) : "",
    "",
  ];
  const files = proposal.target_files || details.target_files || [];
  if (files.length) {
    lines.push("Files:");
    for (const file of files) lines.push(`- ${file}`);
    lines.push("");
  }
  const steps = details.implementation_steps || proposal.implementation_steps || [];
  if (steps.length) {
    lines.push("Plan:");
    for (const step of steps) lines.push(`- ${step}`);
    lines.push("");
  }
  const criteria = details.acceptance_criteria || proposal.acceptance_criteria || [];
  if (criteria.length) {
    lines.push("Done when:");
    for (const item of criteria) lines.push(`- ${item}`);
    lines.push("");
  }
  const markdown = details.markdown_preview || proposal.markdown_preview;
  if (markdown) {
    lines.push(String(markdown));
    lines.push("");
  }
  const diff = details.diff_preview || proposal.diff_preview;
  if (diff) {
    lines.push("Diff:");
    lines.push(String(diff));
    lines.push("");
  } else if (details.raw_preview) {
    lines.push("Coder output:");
    lines.push(String(details.raw_preview));
    lines.push("");
  }
  if (Array.isArray(details.tool_trace) && details.tool_trace.length) {
    lines.push("Tools used:");
    for (const trace of details.tool_trace.slice(-12)) lines.push(`- ${trace.tool}`);
  }
  return lines.filter((line, index, all) => line !== "" || all[index - 1] !== "").join("\n").trim();
}

function renderDeveloperActions(row) {
  const actions = qs("devActions");
  const rejectWrap = qs("devRejectWrap");
  if (!actions) return;
  const status = String(row?.status || "").toLowerCase();
  const id = Number(row?.id || 0);
  const buttons = [];
  if (status === "draft") buttons.push(`<button type="button" data-dev-action="plan" data-dev-id="${id}">Plan</button>`);
  if (["planned", "draft"].includes(status)) buttons.push(`<button type="button" class="btn-primary" data-dev-action="approve" data-dev-id="${id}">Approve plan</button>`);
  if (["approved", "executed"].includes(status)) buttons.push(`<button type="button" class="btn-primary" data-dev-action="code" data-dev-id="${id}">Write the code</button>`);
  if (["patch_ready", "patch_applied", "merged"].includes(status)) buttons.push(`<button type="button" data-dev-action="apply-patch" data-dev-id="${id}" ${status === "merged" ? "disabled" : ""}>Apply in sandbox</button>`);
  if (status === "patch_applied") buttons.push(`<button type="button" class="btn-primary" data-dev-action="merge" data-dev-id="${id}">Merge into IronLog</button>`);
  if (status && status !== "merged" && status !== "rejected") buttons.push(`<button type="button" data-dev-action="reject" data-dev-id="${id}">Stop this job</button>`);
  actions.innerHTML = buttons.join(" ");
  if (rejectWrap) rejectWrap.hidden = !id || status === "merged" || status === "rejected";
}

async function openDeveloperJob(id) {
  const jobId = Number(id || 0);
  if (!jobId) return;
  developerSelectedId = jobId;
  renderDeveloperJobs();
  const title = qs("devDetailTitle");
  const meta = qs("devDetailMeta");
  const preview = qs("devPreview");
  if (preview) preview.textContent = "Loading job…";
  const data = await developerRequest(`/requests/${jobId}`);
  const row = data.row || {};
  const runs = Array.isArray(data.runs) ? data.runs : [];
  if (title) title.textContent = `#${jobId} ${row.title || "Coding job"}`;
  if (meta) meta.textContent = developerStatusLabel(row.status);
  if (preview) preview.textContent = renderDeveloperPreview(row, runs);
  renderDeveloperActions(row);
}

async function loadDeveloperJobs() {
  const data = await developerRequest("/requests");
  developerJobs = Array.isArray(data.rows) ? data.rows : [];
  renderDeveloperJobs();
  if (developerSelectedId && developerJobs.some((job) => Number(job.id) === developerSelectedId)) {
    await openDeveloperJob(developerSelectedId);
  }
}

async function developerPost(id, action, payload) {
  return developerRequest(`/requests/${id}/${action}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload || {}),
  });
}

async function runDeveloperAction(id, action) {
  if (developerBusy) return;
  const jobId = Number(id || 0);
  if (!jobId) return;
  developerBusy = true;
  developerSelectedId = jobId;
  try {
    if (action === "code") {
      developerMsg("Preparing the sandbox, then the coder will write the change. This can take a few minutes.");
      const prepared = await developerPost(jobId, "execute");
      if (prepared?.ok === false) throw new Error(prepared?.run?.summary || prepared?.error || "Sandbox setup failed");
      const coded = await developerPost(jobId, "generate-patch");
      if (coded?.ok === false) throw new Error(coded?.run?.summary || coded?.error || "The coder did not produce a diff");
      developerMsg(`Job #${jobId}: code is ready to review.`);
    } else if (action === "merge") {
      developerMsg("Merging the sandbox change into IronLog…");
      const merged = await developerPost(jobId, "merge", { confirm: true });
      if (merged?.ok === false) throw new Error(merged?.run?.summary || merged?.error || "Merge failed");
      developerMsg(`Job #${jobId} is merged.`);
    } else if (action === "reject") {
      const reason = String(qs("devRejectReason")?.value || "").trim();
      if (!reason) {
        developerMsg("Add a reason before stopping this job.", true);
        return;
      }
      await developerPost(jobId, "reject", { reason });
      developerMsg(`Job #${jobId} stopped.`);
    } else {
      developerMsg(`Running ${action}…`);
      const result = await developerPost(jobId, action);
      if (result?.ok === false) throw new Error(result?.run?.summary || result?.error || `${action} failed`);
      developerMsg(`Job #${jobId}: ${action} complete.`);
    }
    await loadDeveloperJobs();
  } catch (err) {
    developerMsg(err.message || String(err), true);
    await openDeveloperJob(jobId).catch(() => {});
  } finally {
    developerBusy = false;
  }
}

async function createDeveloperJob() {
  if (developerBusy) return;
  const title = String(qs("devTitle")?.value || "").trim();
  const requestText = String(qs("devRequest")?.value || "").trim();
  const priority = String(qs("devPriority")?.value || "medium").trim().toLowerCase();
  if (!title) return developerMsg("Title is required.", true);
  if (!requestText) return developerMsg("Describe the change first.", true);
  developerBusy = true;
  developerMsg("Saving the job and asking the planner…");
  try {
    const created = await developerRequest("/requests", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ title, request_text: requestText, priority }),
    });
    const id = Number(created?.row?.id || 0);
    if (!id) throw new Error("The job was not created.");
    developerSelectedId = id;
    const planned = await developerPost(id, "plan");
    if (planned?.ok === false) throw new Error(planned?.run?.summary || planned?.error || "Planning failed");
    if (qs("devRequest")) qs("devRequest").value = "";
    developerMsg(`Job #${id} is planned. Approve it before the coder writes code.`);
    await loadDeveloperJobs();
  } catch (err) {
    developerMsg(err.message || String(err), true);
    await loadDeveloperJobs().catch(() => {});
  } finally {
    developerBusy = false;
  }
}

function refreshDeveloperTab() {
  loadDeveloperAgentStatus().catch(() => {});
  loadDeveloperJobs().catch((err) => developerMsg(err.message || String(err), true));
}

function initDeveloperTab() {
  if (window.__developerTabReady) return;
  window.__developerTabReady = true;
  qs("devCreateBtn")?.addEventListener("click", () => {
    createDeveloperJob().catch((err) => developerMsg(err.message || String(err), true));
  });
  qs("devRefreshBtn")?.addEventListener("click", () => refreshDeveloperTab());
  qs("devJobList")?.addEventListener("click", (event) => {
    const button = event.target?.closest?.("[data-dev-open]");
    if (!button) return;
    openDeveloperJob(button.getAttribute("data-dev-open")).catch((err) => developerMsg(err.message || String(err), true));
  });
  qs("devActions")?.addEventListener("click", (event) => {
    const button = event.target?.closest?.("[data-dev-action]");
    if (!button || button.disabled) return;
    runDeveloperAction(button.getAttribute("data-dev-id"), button.getAttribute("data-dev-action"))
      .catch((err) => developerMsg(err.message || String(err), true));
  });
}
