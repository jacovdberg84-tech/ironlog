// IRONLOG/web/app/compliance.js — Audit trail, approvals, legal documents.
// Part of the main app; index.html loads these files in order and they share one global scope.

async function loadAuditLogs() {
  const module = (qs("auditModule")?.value || "").trim();
  const action = (qs("auditAction")?.value || "").trim();
  const entity_type = (qs("auditEntityType")?.value || "").trim();
  const entity_id = (qs("auditEntityId")?.value || "").trim();
  const username = (qs("auditUsername")?.value || "").trim();
  const site_code = (qs("auditSiteCode")?.value || "").trim();
  const source_app = (qs("auditSourceApp")?.value || "").trim();
  const start = (qs("auditStart")?.value || "").trim();
  const end = (qs("auditEnd")?.value || "").trim();
  const limit = Number(qs("auditLimit")?.value || 200);
  const list = qs("auditList");
  const detail = qs("auditDetail");
  if (!list) return;
  if (detail) detail.textContent = "";

  setStatus("Loading audit trail...");
  setSkeleton("auditList", 2);

  const q = new URLSearchParams();
  if (module) q.set("module", module);
  if (action) q.set("action", action);
  if (entity_type) q.set("entity_type", entity_type);
  if (entity_id) q.set("entity_id", entity_id);
  if (username) q.set("username", username);
  if (site_code) q.set("site_code", site_code);
  if (source_app) q.set("source_app", source_app);
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  if (Number.isFinite(limit) && limit > 0) q.set("limit", String(Math.trunc(limit)));

  const url = `${API}/api/audit/timeline${q.toString() ? `?${q.toString()}` : ""}`;
  const data = await fetchJson(url);
  const rows = Array.isArray(data.rows) ? data.rows : [];

  list.innerHTML = "";
  rows.forEach((r) => {
    const row = item(
      `<b>${r.created_at || "-"}</b> — ${r.module}.${r.action} ` +
      `<span class="pill blue">${r.role || "-"}</span> <span class="pill">${r.source_app || "-"}</span>` +
      `<br><small>${r.username || "-"} | site=${r.site_code || "-"} | ${r.entity_type || "-"}:${r.entity_id || "-"}</small>`
    );
    row.style.cursor = "pointer";
    row.addEventListener("click", async () => {
      if (!detail) return;
      detail.textContent = "Loading detail...";
      try {
        const d = await fetchJson(`${API}/api/audit/timeline/${Number(r.id || 0)}`);
        detail.textContent = JSON.stringify(d?.row || {}, null, 2);
      } catch (e) {
        detail.textContent = String(e.message || e);
      }
    });
    list.appendChild(row);
  });
  if (!rows.length) list.appendChild(item("<small>No audit records found.</small>"));

  const page = data?.pagination || {};
  setStatus(`Audit trail ready${page?.has_more ? " (more available)" : ""}.`);
}

function canManageLegalDocs() {
  const roles = getSessionRoles();
  return roles.includes("admin") || roles.includes("supervisor");
}

function canApproveRequests() {
  const roles = getSessionRoles();
  return roles.includes("admin") || roles.includes("supervisor");
}

function isTodayStamp(v) {
  const s = String(v || "").trim();
  if (!s) return false;
  const today = new Date().toISOString().slice(0, 10);
  return s.startsWith(today);
}

async function loadApprovalRequests() {
  const list = qs("approvalList");
  if (!list) return;

  const status = (qs("approvalStatus")?.value || "").trim();
  const module = (qs("approvalModule")?.value || "").trim();
  const action = (qs("approvalAction")?.value || "").trim();

  setStatus("Loading approvals...");
  setSkeleton("approvalList", 2);

  const q = new URLSearchParams();
  if (status) q.set("status", status);
  if (module) q.set("module", module);
  if (action) q.set("action", action);

  const data = await fetchJson(`${API}/api/approvals${q.toString() ? `?${q.toString()}` : ""}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  const approver = canApproveRequests();

  setText("approvalAllCount", rows.length);
  const pendingCount = rows.filter((r) => String(r.status || "").toLowerCase() === "pending").length;
  const approvedTodayCount = rows.filter(
    (r) => String(r.status || "").toLowerCase() === "approved" && isTodayStamp(r.approved_at)
  ).length;
  const rejectedTodayCount = rows.filter(
    (r) => String(r.status || "").toLowerCase() === "rejected" && isTodayStamp(r.rejected_at)
  ).length;
  setText("approvalPendingCount", pendingCount);
  setText("approvalApprovedTodayCount", approvedTodayCount);
  setText("approvalRejectedTodayCount", rejectedTodayCount);
  const approvalKpiStrip = qs("approvalKpiStrip");
  if (approvalKpiStrip) {
    const currentFilter = status;
    Array.from(approvalKpiStrip.querySelectorAll("[data-approval-kpi-filter]")).forEach((el) => {
      if (!(el instanceof HTMLElement)) return;
      const filter = String(el.getAttribute("data-approval-kpi-filter") || "");
      el.classList.toggle("pill-active", filter === currentFilter);
    });
  }

  list.innerHTML = "";
  rows.forEach((r) => {
    const st = String(r.status || "").toLowerCase();
    const statusPill =
      st === "approved"
        ? "<span class='pill blue'>approved</span>"
        : st === "rejected"
        ? "<span class='pill red'>rejected</span>"
        : "<span class='pill orange'>pending</span>";
    const payloadTxt = r.payload ? JSON.stringify(r.payload) : "{}";
    const actionBtns =
      approver && st === "pending"
        ? `<br><button data-approval-approve-id="${r.id}" style="margin-top:8px;">Approve</button><button data-approval-reject-id="${r.id}" style="margin-top:8px;">Reject</button>`
        : "";
    list.appendChild(
      item(
        `<b>#${r.id}</b> ${statusPill} — ${r.module || "-"} . ${r.action || "-"}` +
          `<br><small>Requested: ${r.requested_by || "-"} (${r.requested_role || "-"}) @ ${r.created_at || "-"}</small>` +
          `<br><small>Entity: ${r.entity_type || "-"}:${r.entity_id || "-"}</small>` +
          `<br><small>${payloadTxt}</small>` +
          actionBtns
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No approval requests found.</small>"));
  setStatus("Approvals ready.");
}

async function decideApprovalRequest(id, decision) {
  if (!canApproveRequests()) {
    alert("Only admin/supervisor can approve or reject requests.");
    return;
  }
  const reqId = Number(id || 0);
  const d = String(decision || "").trim().toLowerCase();
  if (!reqId || !["approve", "reject"].includes(d)) return;

  const note = (qs("approvalDecisionNote")?.value || "").trim();
  setStatus(`${d === "approve" ? "Approving" : "Rejecting"} request #${reqId}...`);
  try {
    const res = await fetchJson(`${API}/api/approvals/${reqId}/${d}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ note: note || undefined }),
    });
    setText("approvalResult", JSON.stringify(res, null, 2));
    await Promise.all([loadApprovalRequests().catch(() => {}), loadDashboard().catch(() => {})]);
    setStatus(`Approval request #${reqId} ${d}d.`);
  } catch (e) {
    setText("approvalResult", String(e.message || e));
    setStatus("Approval decision failed.");
  }
}

function legalAllowedTransitions(currentStatus) {
  const s = String(currentStatus || "draft").toLowerCase();
  const transitions = {
    draft: ["pending_approval", "superseded"],
    rejected: ["pending_approval", "superseded"],
    pending_approval: ["approved", "rejected", "superseded"],
    approved: ["superseded"],
    superseded: [],
  };
  return transitions[s] || [];
}

async function loadLegalDepartments() {
  const depEl = qs("legalDepartment");
  const filterEl = qs("legalFilterDepartment");
  if (!depEl && !filterEl) return;

  const data = await fetchJson(`${API}/api/legal/departments`);
  const deps = Array.isArray(data.departments) ? data.departments : [];

  if (depEl) {
    depEl.innerHTML = deps.map((d) => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join("");
  }
  if (filterEl) {
    const options = [`<option value="">All departments</option>`]
      .concat(deps.map((d) => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`));
    filterEl.innerHTML = options.join("");
  }
}

async function loadLegalDocs() {
  const dep = (qs("legalFilterDepartment")?.value || "").trim();
  const status = (qs("legalFilterStatus")?.value || "").trim();
  const qText = (qs("legalSearch")?.value || "").trim();
  const includeInactive = qs("legalIncludeInactive")?.checked ? "1" : "0";
  const list = qs("legalList");
  if (!list) return;

  setStatus("Loading legal library...");
  setSkeleton("legalList", 2);

  const q = new URLSearchParams();
  if (dep) q.set("department", dep);
  if (status) q.set("status", status);
  if (qText) q.set("q", qText);
  q.set("include_inactive", includeInactive);

  const data = await fetchJson(`${API}/api/legal?${q.toString()}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];

  list.innerHTML = "";
  rows.forEach((r) => {
    const allowed = legalAllowedTransitions(r.status);
    const statusPill =
      r.status === "approved"
        ? "<span class='pill blue'>approved</span>"
        : r.status === "pending_approval"
        ? "<span class='pill orange'>pending</span>"
        : r.status === "rejected"
        ? "<span class='pill red'>rejected</span>"
        : r.status === "superseded"
        ? "<span class='pill orange'>superseded</span>"
        : "<span class='pill'>draft</span>";

    const archiveBtn = canManageLegalDocs()
      ? `<button data-legal-archive-id="${r.id}" data-legal-active="${r.active ? 0 : 1}" style="margin-top:8px;">${r.active ? "Archive" : "Reactivate"}</button>`
      : "";
    const statusLabel = {
      pending_approval: "Submit",
      approved: "Approve",
      rejected: "Reject",
      superseded: "Supersede",
    };
    const statusBtns = canManageLegalDocs() && Number(r.active) === 1
      ? allowed
          .map(
            (next) =>
              `<button data-legal-status-id="${r.id}" data-legal-status="${next}" style="margin-top:8px;">${statusLabel[next] || next}</button>`
          )
          .join("")
      : "";
    list.appendChild(
      item(
        `<b>${r.department || "-"}</b> — ${r.title || "-"} ${statusPill} ${r.active ? "" : "<span class='pill red'>ARCHIVED</span>"}` +
        `<br><small>Type: ${r.doc_type || "-"} | Version: ${r.version || "-"} | Owner: ${r.owner || "-"} | Uploaded: ${r.created_at || "-"}</small>` +
        `<br><small>Effective: ${r.effective_date || "-"} | Expiry: ${r.expiry_date || "-"} | Approved by: ${r.approved_by || "-"} ${r.approved_at ? `@ ${r.approved_at}` : ""}</small>` +
        (r.approval_note ? `<br><small>Note: ${r.approval_note}</small>` : "") +
        `<br><button data-legal-download-id="${r.id}" style="margin-top:8px;">Download</button>` +
        `<button data-legal-actions-id="${r.id}" style="margin-top:8px;">History</button>` +
        `${statusBtns}${archiveBtn}`
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No documents found.</small>"));
  setStatus("Legal library ready.");
}

async function loadLegalExpiry() {
  const days = Number(qs("legalExpiryDays")?.value || 90);
  const dep = (qs("legalFilterDepartment")?.value || "").trim();
  const status = (qs("legalFilterStatus")?.value || "").trim() || "approved";

  const q = new URLSearchParams();
  if (Number.isFinite(days) && days > 0) q.set("days", String(Math.trunc(days)));
  if (dep) q.set("department", dep);
  if (status) q.set("status", status);

  const data = await fetchJson(`${API}/api/legal/expiry?${q.toString()}`);
  const s = data?.summary || {};
  const setText = (id, v) => {
    const el = qs(id);
    if (el) el.textContent = String(v ?? 0);
  };
  setText("legalExpiredCount", Number(s.expired || 0));
  setText("legalDue30Count", Number(s.due_30 || 0));
  setText("legalDue60Count", Number(s.due_60 || 0));
  setText("legalDue90Count", Number(s.due_90 || 0));
}

function openLegalCompliancePdf(download = false) {
  const days = Number(qs("legalExpiryDays")?.value || 90);
  const dep = (qs("legalFilterDepartment")?.value || "").trim();
  const status = (qs("legalFilterStatus")?.value || "").trim() || "approved";
  const q = new URLSearchParams();
  if (Number.isFinite(days) && days > 0) q.set("days", String(Math.trunc(days)));
  if (dep) q.set("department", dep);
  if (status) q.set("status", status);
  if (download) q.set("download", "1");
  return openAuthedReport(`${API}/api/reports/legal-compliance.pdf?${q.toString()}`, {
    download,
    filename: "IRONLOG_Legal_Compliance.pdf",
  });
}

async function setLegalStatus(id, status) {
  if (!canManageLegalDocs()) {
    alert("Only admin/supervisor can change legal document status.");
    return;
  }
  const docId = Number(id || 0);
  if (!docId) return;
  const note = (qs("legalActionNote")?.value || "").trim() || undefined;
  const supEl = qs("legalSupersedesId");
  const supHint = qs("legalSupersedeHint");

  const payload = { status, note };
  if (status === "superseded") {
    if (supEl) {
      supEl.disabled = false;
      if (supHint) supHint.style.display = "";
      const v = (supEl.value || "").trim();
      if (!v) {
        setStatus("Enter 'Supersedes Doc ID' and click Supersede again.");
        supEl.focus();
        return;
      }
      payload.supersedes_document_id = Number(v);
    }
  } else if (supEl) {
    supEl.value = "";
    supEl.disabled = true;
    if (supHint) supHint.style.display = "none";
  }

  const out = qs("legalActionResult");
  setStatus(`Applying status '${status}'...`);
  try {
    const res = await fetchJson(`${API}/api/legal/${docId}/status`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (out) out.textContent = JSON.stringify(res, null, 2);
    const noteEl = qs("legalActionNote");
    if (noteEl) noteEl.value = "";
    if (supEl) {
      supEl.value = "";
      supEl.disabled = true;
    }
    if (supHint) supHint.style.display = "none";
    await loadLegalDocs().catch(() => {});
    setStatus("Legal status updated.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Legal status update failed.");
  }
}

async function showLegalActions(id) {
  const docId = Number(id || 0);
  if (!docId) return;
  const out = qs("legalActionResult");
  setStatus("Loading legal action history...");
  try {
    const data = await fetchJson(`${API}/api/legal/${docId}/actions`);
    if (out) out.textContent = JSON.stringify(data.actions || [], null, 2);
    setStatus("Legal action history loaded.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Legal action history failed.");
  }
}

async function uploadLegalDoc() {
  if (!canManageLegalDocs()) {
    alert("Only admin/supervisor can upload legal documents.");
    return;
  }

  const fileEl = qs("legalFile");
  const file = fileEl?.files?.[0];
  if (!file) return alert("Choose a file first.");

  const fd = new FormData();
  let department = (qs("legalDepartment")?.value || "").trim();
  if (!department) {
    try {
      await loadLegalDepartments();
      department = (qs("legalDepartment")?.value || "").trim();
    } catch (_) {
      // Keep graceful fallback below.
    }
  }
  if (!department) {
    const depEl = qs("legalDepartment");
    const firstOpt = depEl?.querySelector("option");
    department = String(firstOpt?.value || firstOpt?.textContent || "").trim();
  }
  if (!department) {
    return alert("Select a department before upload.");
  }

  let title = (qs("legalTitle")?.value || "").trim();
  if (!title) {
    title = String(file.name || "Untitled")
      .replace(/\.[^.]+$/, "")
      .trim();
    const titleEl = qs("legalTitle");
    if (titleEl) titleEl.value = title;
  }
  if (!title) {
    return alert("Enter a document title before upload.");
  }

  fd.append("file", file);
  fd.append("department", department);
  fd.append("title", title);
  fd.append("doc_type", (qs("legalDocType")?.value || "").trim());
  fd.append("version", (qs("legalVersion")?.value || "").trim());
  fd.append("owner", (qs("legalOwner")?.value || "").trim());
  fd.append("effective_date", (qs("legalEffectiveDate")?.value || "").trim());
  fd.append("expiry_date", (qs("legalExpiryDate")?.value || "").trim());

  const resultEl = qs("legalUploadResult");
  setStatus("Uploading legal document...");
  try {
    const res = await fetch(`${API}/api/legal/upload`, {
      method: "POST",
      headers: authHeaders(),
      body: fd,
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Upload failed");
    if (resultEl) resultEl.textContent = JSON.stringify(data, null, 2);
    await loadLegalDocs().catch(() => {});
    setStatus("Legal document uploaded.");
  } catch (e) {
    if (resultEl) resultEl.textContent = String(e.message || e);
    setStatus("Legal upload failed.");
  }
}

function downloadLegalDoc(id) {
  const docId = Number(id || 0);
  if (!docId) return;
  return downloadAuthedFile(`${API}/api/legal/${docId}/download`, `IRONLOG_Legal_Document_${docId}`);
}

async function archiveLegalDoc(id, active) {
  if (!canManageLegalDocs()) {
    alert("Only admin/supervisor can archive/reactivate documents.");
    return;
  }
  const docId = Number(id || 0);
  if (!docId) return;

  setStatus(active ? "Reactivating document..." : "Archiving document...");
  try {
    const res = await fetchJson(`${API}/api/legal/${docId}/archive`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ active }),
    });
    await loadLegalDocs().catch(() => {});
    setStatus(active ? "Document reactivated." : "Document archived.");
    return res;
  } catch (e) {
    setStatus("Document archive action failed.");
    alert(e.message || e);
  }
}

/* =========================
   TABS
========================= */
