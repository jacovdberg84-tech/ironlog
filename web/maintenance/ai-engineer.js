// IRONLOG/web/maintenance/ai-engineer.js — AI engineer requests.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

function aiEngineerMsg(text, isError = false) {
  const el = document.getElementById("aiEngMsg");
  if (!el) return;
  el.className = isError ? "message-error" : "muted";
  el.textContent = String(text || "");
}

function aiEngineerActionButtons(row) {
  const id = Number(row?.id || 0);
  const status = String(row?.status || "draft").toLowerCase();
  const planBtn = `<button type="button" data-aie-plan="${id}">Plan</button>`;
  const approveBtn = `<button type="button" data-aie-approve="${id}" ${["approved", "executed", "patch_ready", "patch_applied", "merged"].includes(status) ? "disabled" : ""}>Approve</button>`;
  const executeBtn = `<button type="button" data-aie-execute="${id}" ${!["approved", "planned", "executed"].includes(status) ? "disabled" : ""}>Execute</button>`;
  const genPatchBtn = `<button type="button" data-aie-gpatch="${id}" ${!["executed", "patch_ready", "patch_applied", "merged"].includes(status) ? "disabled" : ""}>Gen Patch</button>`;
  const applyPatchBtn = `<button type="button" data-aie-apatch="${id}" ${!["patch_ready", "patch_applied", "merged"].includes(status) ? "disabled" : ""}>Apply Patch</button>`;
  const mergeBtn = `<button type="button" data-aie-merge="${id}" ${status !== "patch_applied" ? "disabled" : ""}>Merge</button>`;
  const rejectBtn = `<button type="button" data-aie-reject="${id}" ${status === "merged" ? "disabled" : ""}>Reject</button>`;
  const viewBtn = `<button type="button" data-aie-view="${id}">View</button>`;
  return `${planBtn} ${approveBtn} ${executeBtn} ${genPatchBtn} ${applyPatchBtn} ${mergeBtn} ${rejectBtn} ${viewBtn}`;
}

async function loadAiEngineerRequests() {
  const body = document.getElementById("aiEngBody");
  if (!body) return;
  body.innerHTML = `<tr><td colspan="7" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/ai-engineer/requests`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load requests");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    if (!rows.length) {
      body.innerHTML = `<tr><td colspan="7" class="muted">No requests yet.</td></tr>`;
      return;
    }
    body.innerHTML = rows.map((r) => `
      <tr>
        <td>#${Number(r.id || 0)}</td>
        <td>${esc(r.status || "-")}</td>
        <td>${esc(r.priority || "-")}</td>
        <td>${esc(r.title || "-")}</td>
        <td>${esc(r.requested_by || "-")}</td>
        <td>${esc(r.created_at || "-")}</td>
        <td>${aiEngineerActionButtons(r)}</td>
      </tr>
    `).join("");
  } catch (e) {
    body.innerHTML = `<tr><td colspan="7" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

async function createAiEngineerRequest() {
  const title = String(document.getElementById("aiEngTitle")?.value || "").trim();
  const requestText = String(document.getElementById("aiEngRequestText")?.value || "").trim();
  const priority = String(document.getElementById("aiEngPriority")?.value || "medium").trim().toLowerCase();
  if (!title) return aiEngineerMsg("Title is required.", true);
  if (!requestText) return aiEngineerMsg("Request details are required.", true);
  aiEngineerMsg("Creating request...");
  try {
    const res = await fetch(`${API}/ai-engineer/requests`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({
        title,
        request_text: requestText,
        priority,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to create request");
    aiEngineerMsg(`Created AI engineer request #${Number(data?.row?.id || 0)}.`);
    const txt = document.getElementById("aiEngRequestText");
    if (txt) txt.value = "";
    await loadAiEngineerRequests();
  } catch (e) {
    aiEngineerMsg(e.message || String(e), true);
  }
}

function pickAiEngineerPreviewRun(runs) {
  const order = ["merge", "apply_patch", "generate_patch", "execute", "plan"];
  for (const runType of order) {
    const hit = runs.find((r) => String(r?.run_type || "") === runType && r?.details);
    if (hit) return hit;
  }
  return runs[0] || null;
}

function renderAiEngineerRunPreview({ row, runs }) {
  const previewRun = pickAiEngineerPreviewRun(runs);
  const details = previewRun?.details || null;
  const lines = [
    `Request #${Number(row?.id || 0)}`,
    `Status: ${row?.status || "-"}`,
    `Priority: ${row?.priority || "-"}`,
    `Title: ${row?.title || "-"}`,
    `Latest run: ${previewRun ? `${previewRun.run_type} (${previewRun.status})` : "none"}`,
    ``,
    `${previewRun?.summary || ""}`,
    ``,
  ];

  if (details?.agent) {
    lines.push(`Agent: ${details.agent}`);
    lines.push("");
  }
  if (details?.model) {
    lines.push(`Model: ${details.model}`);
    lines.push("");
  }
  if (details?.patch_file) {
    lines.push(`Patch file: ${details.patch_file}`);
    lines.push("");
  }
  if (details?.proposal?.target_files?.length) {
    lines.push("Target files:");
    for (const f of details.proposal.target_files) lines.push(`- ${f}`);
    lines.push("");
  } else if (Array.isArray(details?.target_files) && details.target_files.length) {
    lines.push("Target files:");
    for (const f of details.target_files) lines.push(`- ${f}`);
    lines.push("");
  }
  if (details?.proposal?.markdown_preview) {
    lines.push("Proposal preview:");
    lines.push(details.proposal.markdown_preview);
    lines.push("");
  } else if (details?.markdown_preview) {
    lines.push("Proposal preview:");
    lines.push(String(details.markdown_preview));
    lines.push("");
  }
  if (Array.isArray(details?.implementation_steps) && details.implementation_steps.length) {
    lines.push("Implementation steps:");
    for (const s of details.implementation_steps) lines.push(`- ${s}`);
    lines.push("");
  } else if (Array.isArray(details?.proposal?.implementation_steps) && details.proposal.implementation_steps.length) {
    lines.push("Implementation steps:");
    for (const s of details.proposal.implementation_steps) lines.push(`- ${s}`);
    lines.push("");
  }
  if (Array.isArray(details?.acceptance_criteria) && details.acceptance_criteria.length) {
    lines.push("Acceptance criteria:");
    for (const s of details.acceptance_criteria) lines.push(`- ${s}`);
    lines.push("");
  }
  if (details?.gates) {
    lines.push("Gates:");
    for (const [k, v] of Object.entries(details.gates)) lines.push(`- ${k}: ${v ? "pass" : "fail"}`);
    lines.push("");
  }
  if (details?.commit_hash) {
    lines.push(`Commit: ${details.commit_hash}`);
    lines.push("");
  }
  if (details?.branch) {
    lines.push(`Branch: ${details.branch}`);
    lines.push("");
  }
  if (details?.changed_files?.length) {
    lines.push("Changed files:");
    for (const f of details.changed_files) lines.push(`- ${f}`);
    lines.push("");
  }
  if (details?.merged_files?.length) {
    lines.push("Merged files:");
    for (const f of details.merged_files) lines.push(`- ${f}`);
    lines.push("");
  }
  if (details?.diff_preview) {
    lines.push("Diff preview:");
    lines.push(String(details.diff_preview));
    lines.push("");
  } else if (details?.proposal?.diff_preview) {
    lines.push("Diff preview:");
    lines.push(String(details.proposal.diff_preview));
    lines.push("");
  } else if (details?.raw_preview) {
    lines.push("Model output preview:");
    lines.push(String(details.raw_preview));
    lines.push("");
  } else if (details) {
    lines.push(JSON.stringify(details, null, 2));
  }
  if (Array.isArray(details?.tool_trace) && details.tool_trace.length) {
    lines.push("Agent tool trace:");
    for (const t of details.tool_trace.slice(-12)) {
      lines.push(`- ${t.tool} @ ${t.at}`);
    }
    lines.push("");
  }
  return lines.join("\n");
}

async function runAiEngineerAction(id, action, payload = null) {
  const reqId = Number(id || 0);
  if (!reqId) return;
  aiEngineerMsg(`Running ${action}...`);
  try {
    const res = await fetch(`${API}/ai-engineer/requests/${reqId}/${action}`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: payload ? JSON.stringify(payload) : "{}",
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || `Failed to ${action}`);
    if (data?.run?.details || data?.row) {
      const pre = document.getElementById("aiEngRunPreview");
      if (pre) {
        pre.textContent = renderAiEngineerRunPreview({
          row: data.row || { id: reqId },
          runs: data.run ? [data.run] : [],
        });
      }
    }
    await loadAiEngineerRequests();
    if (["generate-patch", "apply-patch", "merge", "execute", "plan"].includes(action)) {
      await viewAiEngineerRequest(reqId);
    }
    if (data?.ok === false) {
      aiEngineerMsg(data?.run?.summary || data?.error || `${action} failed`, true);
      return;
    }
    aiEngineerMsg(`Request #${reqId}: ${action} complete.`);
  } catch (e) {
    aiEngineerMsg(e.message || String(e), true);
  }
}

async function viewAiEngineerRequest(id) {
  const reqId = Number(id || 0);
  if (!reqId) return;
  aiEngineerMsg(`Loading request #${reqId} details...`);
  try {
    const res = await fetch(`${API}/ai-engineer/requests/${reqId}`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load request details");
    const row = data?.row || {};
    const runs = Array.isArray(data?.runs) ? data.runs : [];
    const pre = document.getElementById("aiEngRunPreview");
    if (pre) pre.textContent = renderAiEngineerRunPreview({ row, runs });
    aiEngineerMsg(`Loaded request #${reqId}.`);
  } catch (e) {
    aiEngineerMsg(e.message || String(e), true);
  }
}
