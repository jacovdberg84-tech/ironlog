const API = "/api";
let closingWorkOrderId = null;
let closingWorkOrderSource = "";
const ROLE_KEY = "ironlog_session_role";
const USER_KEY = "ironlog_session_user";

function getSessionRole() {
  return String(localStorage.getItem(ROLE_KEY) || "admin").trim().toLowerCase() || "admin";
}

function getSessionUser() {
  return String(localStorage.getItem(USER_KEY) || "admin").trim() || "admin";
}

function authHeaders(extra = {}) {
  return {
    ...extra,
    "x-user-role": getSessionRole(),
    "x-user-name": getSessionUser(),
  };
}

function canRoleTransition(role, currentStatus, nextStatus) {
  const r = String(role || "").toLowerCase();
  if (r === "admin" || r === "supervisor") return true;
  if (r === "artisan") {
    const allowed = {
      assigned: ["in_progress"],
      in_progress: ["completed", "assigned"],
      completed: ["in_progress"],
    };
    return (allowed[currentStatus] || []).includes(nextStatus);
  }
  return false;
}

function canRoleClose(role) {
  return ["admin", "supervisor"].includes(String(role || "").toLowerCase());
}

function rolePermissionText(role) {
  const r = String(role || "").toLowerCase();
  if (r === "admin") {
    return "Admin: full control (status transitions, approvals, close, issue parts).";
  }
  if (r === "supervisor") {
    return "Supervisor: full workflow control (assign/start/approve/close) and issue parts.";
  }
  if (r === "artisan") {
    return "Artisan: execution flow only (start, complete, return to progress) and can request close approval. Cannot approve/close or issue parts.";
  }
  if (r === "stores") {
    return "Stores: issue/allocate parts only. Cannot change work order statuses or close.";
  }
  if (r === "operator") {
    return "Operator: read-only work order visibility in this view.";
  }
  return "Role permissions are limited in this view.";
}

function sourceLabel(source) {
  const s = String(source || "").toLowerCase();
  if (s === "service") return "Service";
  if (s === "breakdown") return "Breakdown";
  if (s === "manual") return "Manual";
  return s || "Unknown";
}

function statusClass(status) {
  const s = String(status || "").toLowerCase();
  if (s === "closed") return "status-ok";
  if (s === "open") return "status-overdue";
  if (s === "in_progress" || s === "assigned") return "status-soon";
  if (s === "completed") return "status-completed";
  if (s === "approved") return "status-approved";
  return "status-ok";
}

function woAgeHours(openedAt) {
  const t = Date.parse(String(openedAt || ""));
  if (!Number.isFinite(t)) return 0;
  return Math.max(0, Math.floor((Date.now() - t) / 3600000));
}

function woPriority(status, openedAt) {
  const s = String(status || "").toLowerCase();
  const age = woAgeHours(openedAt);
  if (s === "completed" && age > 48) return "P1";
  if (s === "in_progress" && age > 72) return "P1";
  if ((s === "open" || s === "assigned") && age > 72) return "P1";
  if (s === "completed" && age > 24) return "P2";
  if (s === "in_progress" && age > 48) return "P2";
  if ((s === "open" || s === "assigned") && age > 48) return "P2";
  return "P3";
}

function statusActionButtons(wo) {
  const role = getSessionRole();
  const s = String(wo?.status || "").toLowerCase();
  if (s === "closed") return "";

  const map = {
    open: [{ label: "Assign", to: "assigned" }, { label: "Start", to: "in_progress" }],
    assigned: [{ label: "Start", to: "in_progress" }, { label: "Reopen", to: "open" }],
    in_progress: [{ label: "Complete", to: "completed" }, { label: "Unassign", to: "assigned" }],
    completed: [{ label: "Approve", to: "approved" }, { label: "Back To Progress", to: "in_progress" }],
    approved: [{ label: "Reopen Completed", to: "completed" }],
  };

  const actions = (map[s] || []).filter((a) => canRoleTransition(role, s, a.to));
  return actions
    .map(
      (a) =>
        `<button data-set-status-id="${wo.id}" data-set-status="${a.to}" style="margin-top:8px;">${a.label}</button>`
    )
    .join("");
}

function workOrderCard(wo) {
  const role = getSessionRole();
  const canClose = String(wo.status || "").toLowerCase() !== "closed" && canRoleClose(role);
  const canRequestClose =
    role === "artisan" &&
    ["completed", "approved"].includes(String(wo.status || "").toLowerCase());
  const ageHours = woAgeHours(wo.opened_at);
  const p = woPriority(wo.status, wo.opened_at);
  const pClass = p === "P1" ? "pri-p1" : p === "P2" ? "pri-p2" : "pri-p3";
  return `
    <div class="card" data-wo-id="${wo.id}">
      <div><strong>WO #${wo.id}</strong></div>
      <div><strong>Asset:</strong> ${wo.asset_code || "-"} - ${wo.asset_name || "-"}</div>
      <div><strong>Source:</strong> ${sourceLabel(wo.source)}</div>
      <div><strong>Reference:</strong> ${wo.reference_id ?? "-"}</div>
      <div><strong>Opened:</strong> ${wo.opened_at || "-"}</div>
      <div><strong>Age:</strong> ${ageHours}h <span class="pill ${pClass}">${p}</span></div>
      <div><strong>Closed:</strong> ${wo.closed_at || "-"}</div>
      <div class="${statusClass(wo.status)}">${String(wo.status || "unknown").toUpperCase()}</div>
      ${statusActionButtons(wo)}
      <button data-pdf-id="${wo.id}" style="margin-top:8px;">Open PDF</button>
      <button data-pdf-download-id="${wo.id}" style="margin-top:8px;">Download PDF</button>
      <button data-view-id="${wo.id}" style="margin-top:8px;">View Detail</button>
      ${canRequestClose ? `<button data-request-close-id="${wo.id}" data-request-close-source="${String(wo.source || "").toLowerCase()}" style="margin-top:8px;">Request Close Approval</button>` : ""}
      ${canClose ? `<button data-close-id="${wo.id}" data-close-source="${String(wo.source || "").toLowerCase()}" style="margin-top:8px;">Close Work Order</button>` : ""}
    </div>
  `;
}

function renderParts(parts) {
  if (!Array.isArray(parts) || !parts.length) {
    return `<div class="muted">No parts issued.</div>`;
  }

  return `
    <div class="list">
      ${parts.map((p) => `
        <div class="item">
          <div><strong>${p.part_code || "-"}</strong> - ${p.part_name || "-"}</div>
          <div>Qty: ${p.quantity ?? "-"} | Type: ${p.movement_type || "-"}</div>
        </div>
      `).join("")}
    </div>
  `;
}

function renderBreakdown(breakdown) {
  if (!breakdown) return `<div class="muted">No linked breakdown.</div>`;

  return `
    <div class="item">
      <div><strong>ID:</strong> ${breakdown.id}</div>
      <div><strong>Date:</strong> ${breakdown.breakdown_date || "-"}</div>
      <div><strong>Critical:</strong> ${breakdown.critical ? "Yes" : "No"}</div>
      <div><strong>Description:</strong> ${breakdown.description || "-"}</div>
    </div>
  `;
}

function renderDetail(payload) {
  const wo = payload?.work_order;
  const breakdown = payload?.breakdown;
  const parts = payload?.parts_issued;

  if (!wo) {
    return `<div class="message-error">Work order detail not available.</div>`;
  }

  return `
    <div class="row" style="gap:20px; align-items:flex-start;">
      <div style="min-width:280px; flex:1;">
        <h4 style="margin:0 0 8px 0;">Core</h4>
        <div class="item">
          <div><strong>WO #:</strong> ${wo.id}</div>
          <div><strong>Asset:</strong> ${wo.asset_code || "-"} - ${wo.asset_name || "-"}</div>
          <div><strong>Source:</strong> ${sourceLabel(wo.source)}</div>
          <div><strong>Reference:</strong> ${wo.reference_id ?? "-"}</div>
          <div><strong>Status:</strong> ${String(wo.status || "").toUpperCase()}</div>
          <div><strong>Opened:</strong> ${wo.opened_at || "-"}</div>
          <div><strong>Closed:</strong> ${wo.closed_at || "-"}</div>
        </div>
      </div>

      <div style="min-width:280px; flex:1;">
        <h4 style="margin:0 0 8px 0;">Linked Breakdown</h4>
        ${renderBreakdown(breakdown)}
      </div>
    </div>

    <div style="margin-top:12px;">
      <h4 style="margin:0 0 8px 0;">Issued Parts</h4>
      ${renderParts(parts)}
    </div>
  `;
}

function isTodayStamp(value) {
  const v = String(value || "").trim();
  if (!v) return false;
  const today = new Date().toISOString().slice(0, 10);
  return v.startsWith(today);
}

function updateKpiStrip(rows) {
  const strip = document.getElementById("woKpiStrip");
  if (!strip) return;

  const list = Array.isArray(rows) ? rows : [];
  const count = (status) =>
    list.filter((r) => String(r.status || "").toLowerCase() === status).length;

  const all = list.length;
  const open = count("open");
  const inProgress = count("in_progress");
  const awaitingApproval = count("completed");
  const approvedToday = list.filter(
    (r) => String(r.status || "").toLowerCase() === "approved" && isTodayStamp(r.closed_at || r.opened_at)
  ).length;
  const closedToday = list.filter(
    (r) => String(r.status || "").toLowerCase() === "closed" && isTodayStamp(r.closed_at)
  ).length;

  strip.innerHTML = `
    <button class="pill blue" data-kpi-filter="">All: ${all}</button>
    <button class="pill red" data-kpi-filter="open">Open: ${open}</button>
    <button class="pill orange" data-kpi-filter="in_progress">In Progress: ${inProgress}</button>
    <button class="pill orange" data-kpi-filter="completed">Awaiting Approval: ${awaitingApproval}</button>
    <button class="pill blue" data-kpi-filter="approved">Approved Today: ${approvedToday}</button>
    <button class="pill blue" data-kpi-filter="closed">Closed Today: ${closedToday}</button>
  `;
}

async function fetchWorkOrders() {
  const statusEl = document.getElementById("woStatus");
  const sourceEl = document.getElementById("woSource");
  const searchEl = document.getElementById("woSearch");
  const listEl = document.getElementById("woList");
  const msgEl = document.getElementById("woMessage");

  if (!statusEl || !sourceEl || !searchEl || !listEl || !msgEl) return;

  const status = String(statusEl.value || "").trim();
  const source = String(sourceEl.value || "").trim().toLowerCase();
  const q = String(searchEl.value || "").trim().toLowerCase();

  listEl.innerHTML = `<div class="skeleton-block"></div><div class="skeleton-block"></div>`;
  msgEl.className = "";
  msgEl.textContent = "";

  try {
    const res = await fetch(`${API}/workorders`, { headers: authHeaders() });
    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || "Failed to load work orders");
    }

    const rows = Array.isArray(data) ? data : [];
    updateKpiStrip(rows);

    const filtered = rows.filter((r) => {
      const statusOk = !status || String(r.status || "").toLowerCase() === status.toLowerCase();
      if (!statusOk) return false;

      const sourceOk = !source || String(r.source || "").toLowerCase() === source;
      if (!sourceOk) return false;

      if (!q) return true;
      const hay = `${r.id} ${r.asset_code || ""} ${r.asset_name || ""}`.toLowerCase();
      return hay.includes(q);
    });

    listEl.innerHTML = filtered.length
      ? filtered.map(workOrderCard).join("")
      : "<div>No work orders found for current filters.</div>";

    const requested = getRequestedWorkOrderId();
    if (requested && filtered.some((r) => Number(r.id) === requested)) {
      loadWorkOrderDetail(requested).catch(() => {});
      setTimeout(() => scrollToWorkOrderCard(requested), 0);
    }

    msgEl.className = "message-success";
    msgEl.textContent = `Showing ${filtered.length} work order(s).`;
  } catch (err) {
    console.error("Load work orders error:", err);
    listEl.innerHTML = `<div style="color:#ff8080;">Error loading work orders: ${err.message}</div>`;
    msgEl.className = "message-error";
    msgEl.textContent = err.message;
  }
}

function openCloseModal(id) {
  const woId = Number(id || 0);
  if (!woId) return;

  closingWorkOrderId = woId;
  const modal = document.getElementById("woCloseModal");
  const title = document.getElementById("woCloseModalTitle");
  const notesEl = document.getElementById("woCloseNotes");
  const artisanEl = document.getElementById("woCloseArtisan");
  const supervisorEl = document.getElementById("woCloseSupervisor");
  const msgEl = document.getElementById("woCloseModalMsg");

  if (!modal || !title || !notesEl || !artisanEl || !supervisorEl || !msgEl) return;

  title.textContent = `#${woId}`;
  notesEl.value = "";
  artisanEl.value = "";
  supervisorEl.value = "";
  msgEl.className = "";
  msgEl.textContent = "";
  modal.style.display = "flex";
}

function openCloseModalForRow(id, source) {
  closingWorkOrderSource = String(source || "").toLowerCase();
  openCloseModal(id);

  const msgEl = document.getElementById("woCloseModalMsg");
  if (closingWorkOrderSource === "service" && msgEl) {
    msgEl.className = "";
    msgEl.textContent = "Artisan name and completion notes are required for service work orders.";
  }
}

function closeCloseModal() {
  closingWorkOrderId = null;
  closingWorkOrderSource = "";
  const modal = document.getElementById("woCloseModal");
  if (modal) modal.style.display = "none";
}

async function submitCloseWorkOrder() {
  const woId = Number(closingWorkOrderId || 0);
  if (!woId) return;

  const notesEl = document.getElementById("woCloseNotes");
  const artisanEl = document.getElementById("woCloseArtisan");
  const supervisorEl = document.getElementById("woCloseSupervisor");
  const msgEl = document.getElementById("woCloseModalMsg");
  const confirmBtn = document.getElementById("woCloseConfirmBtn");

  const completion_notes = String(notesEl?.value || "").trim();
  const artisan_name = String(artisanEl?.value || "").trim();
  const supervisor_name = String(supervisorEl?.value || "").trim();

  if (closingWorkOrderSource === "service" && !artisan_name) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = "Artisan name is required for service work orders.";
    }
    return;
  }
  if (closingWorkOrderSource === "service" && !completion_notes) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = "Completion notes are required for service work orders.";
    }
    return;
  }

  if (confirmBtn) confirmBtn.disabled = true;
  if (msgEl) {
    msgEl.className = "";
    msgEl.textContent = "Saving...";
  }

  try {
    const isRequest = getSessionRole() === "artisan";
    const endpoint = isRequest
      ? `${API}/workorders/${woId}/request-close`
      : `${API}/workorders/${woId}/close`;
    const res = await fetch(endpoint, {
      method: "POST",
      headers: {
        ...authHeaders(),
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        completion_notes,
        artisan_name,
        supervisor_name
      })
    });
    const data = await res.json();
    if (!res.ok) {
      throw new Error(data.error || "Failed to close work order");
    }
    closeCloseModal();
    await fetchWorkOrders();
    const detailEl = document.getElementById("woDetail");
    if (detailEl) {
      detailEl.innerHTML = isRequest
        ? `<div class="message-success">Close approval requested for work order #${woId} (Request #${data.request_id || "-"})</div>`
        : `<div class="message-success">Work order #${woId} completed and closed.</div>`;
    }
  } catch (err) {
    console.error("Close work order error:", err);
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message;
    }
  } finally {
    if (confirmBtn) confirmBtn.disabled = false;
  }
}

async function loadWorkOrderDetail(id) {
  const woId = Number(id || 0);
  const detailEl = document.getElementById("woDetail");
  if (!woId || !detailEl) return;

  detailEl.innerHTML = `<div class="skeleton-block"></div>`;

  try {
    const res = await fetch(`${API}/workorders/${woId}`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) {
      throw new Error(data.error || "Failed to load work order detail");
    }
    detailEl.innerHTML = renderDetail(data);
  } catch (err) {
    console.error("Load work order detail error:", err);
    detailEl.innerHTML = `<div class="message-error">${err.message}</div>`;
  }
}

async function setWorkOrderStatus(id, status) {
  const woId = Number(id || 0);
  const next = String(status || "").trim().toLowerCase();
  if (!woId || !next) return;

  try {
    const res = await fetch(`${API}/workorders/${woId}/status`, {
      method: "POST",
      headers: {
        ...authHeaders(),
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ status: next })
    });
    const data = await res.json();
    if (!res.ok) {
      throw new Error(data.error || "Failed to update status");
    }
    await fetchWorkOrders();
    await loadWorkOrderDetail(woId);
  } catch (err) {
    console.error("Set status error:", err);
    alert(`Could not update status: ${err.message}`);
  }
}

function openWorkOrderPdf(id) {
  const woId = Number(id || 0);
  if (!woId) return;
  window.open(`${API}/reports/workorder/${woId}.pdf`, "_blank");
}

function downloadWorkOrderPdf(id) {
  const woId = Number(id || 0);
  if (!woId) return;
  window.open(`${API}/reports/workorder/${woId}.pdf?download=1`, "_blank");
}

function getRequestedWorkOrderId() {
  try {
    const q = new URLSearchParams(window.location.search);
    const id = Number(q.get("wo") || 0);
    return Number.isFinite(id) && id > 0 ? id : null;
  } catch {
    return null;
  }
}

function scrollToWorkOrderCard(woId) {
  const id = Number(woId || 0);
  if (!id) return;
  const card = document.querySelector(`.card[data-wo-id="${id}"]`);
  if (!card) return;
  card.classList.add("wo-highlight");
  card.scrollIntoView({ behavior: "smooth", block: "center" });
  setTimeout(() => card.classList.remove("wo-highlight"), 2600);
}

document.addEventListener("DOMContentLoaded", () => {
  const refreshBtn = document.getElementById("woRefreshBtn");
  const awaitingApprovalBtn = document.getElementById("woAwaitingApprovalBtn");
  const statusEl = document.getElementById("woStatus");
  const sourceEl = document.getElementById("woSource");
  const searchEl = document.getElementById("woSearch");
  const listEl = document.getElementById("woList");
  const closeConfirmBtn = document.getElementById("woCloseConfirmBtn");
  const closeCancelBtn = document.getElementById("woCloseCancelBtn");
  const closeModal = document.getElementById("woCloseModal");
  const kpiStrip = document.getElementById("woKpiStrip");
  const role = getSessionRole();

  const roleBadge = document.getElementById("woRoleBadge");
  if (roleBadge) roleBadge.textContent = `Role: ${role}`;
  const legend = document.getElementById("woPermissionLegend");
  if (legend) legend.textContent = rolePermissionText(role);

  if (!["admin", "supervisor"].includes(role) && awaitingApprovalBtn) {
    awaitingApprovalBtn.style.display = "none";
  }

  if (refreshBtn) refreshBtn.addEventListener("click", fetchWorkOrders);
  if (awaitingApprovalBtn && statusEl) {
    awaitingApprovalBtn.addEventListener("click", () => {
      statusEl.value = "completed";
      fetchWorkOrders();
    });
  }
  if (statusEl) statusEl.addEventListener("change", fetchWorkOrders);
  if (sourceEl) sourceEl.addEventListener("change", fetchWorkOrders);
  if (searchEl) searchEl.addEventListener("input", fetchWorkOrders);
  const requested = getRequestedWorkOrderId();
  if (requested && searchEl) searchEl.value = String(requested);
  if (closeConfirmBtn) closeConfirmBtn.addEventListener("click", submitCloseWorkOrder);
  if (closeCancelBtn) closeCancelBtn.addEventListener("click", closeCloseModal);
  if (closeModal) {
    closeModal.addEventListener("click", (evt) => {
      if (evt.target === closeModal) closeCloseModal();
    });
  }
  if (kpiStrip && statusEl) {
    kpiStrip.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const filter = target.getAttribute("data-kpi-filter");
      if (filter == null) return;
      statusEl.value = filter;
      fetchWorkOrders();
    });
  }

  if (listEl) {
    listEl.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const pdfId = target.getAttribute("data-pdf-id");
      const pdfDownloadId = target.getAttribute("data-pdf-download-id");
      const viewId = target.getAttribute("data-view-id");
      const id = target.getAttribute("data-close-id");
      const setStatusId = target.getAttribute("data-set-status-id");
      const setStatus = target.getAttribute("data-set-status");
      const rowSource = target.getAttribute("data-close-source");
      const requestCloseId = target.getAttribute("data-request-close-id");
      const requestCloseSource = target.getAttribute("data-request-close-source");
      if (pdfId) {
        openWorkOrderPdf(pdfId);
        return;
      }
      if (pdfDownloadId) {
        downloadWorkOrderPdf(pdfDownloadId);
        return;
      }
      if (viewId) {
        loadWorkOrderDetail(viewId);
        return;
      }
      if (setStatusId && setStatus) {
        setWorkOrderStatus(setStatusId, setStatus);
        return;
      }
      if (requestCloseId) {
        openCloseModalForRow(requestCloseId, requestCloseSource);
        return;
      }
      if (id) openCloseModalForRow(id, rowSource);
    });
  }

  fetchWorkOrders();
});
