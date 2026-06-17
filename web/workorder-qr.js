(function () {
  const A = window.IronlogAuth;
  function qs(id) { return document.getElementById(id); }
  function setText(id, value) { const el = qs(id); if (el) el.textContent = value; }

  function headers(extra) {
    return A ? A.authHeaders(extra || {}) : {};
  }

  function getRole() {
    return A ? A.getSessionRole() : String(localStorage.getItem("ironlog_session_role") || "artisan").trim().toLowerCase();
  }

  function getUser() {
    return A ? A.getSessionUser() : String(localStorage.getItem("ironlog_session_user") || "").trim();
  }

  async function fetchJson(url, options) {
    if (A) return A.fetchJson(url, options);
    const res = await fetch(url, options);
    const text = await res.text();
    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch { data = null; }
    if (!res.ok) throw new Error(data?.error || data?.message || text || `Request failed (${res.status})`);
    return data || {};
  }

  function statusClass(s) {
    const v = String(s || "").toLowerCase();
    if (v === "open") return "status open";
    if (v === "assigned" || v === "in_progress") return "status in_progress";
    if (v === "completed" || v === "approved" || v === "closed") return "status completed";
    return "status";
  }

  function getWoId() {
    const q = new URL(window.location.href).searchParams;
    const id = Number(q.get("wo_id") || 0);
    return Number.isFinite(id) && id > 0 ? id : 0;
  }

  let currentStatus = "";
  let assignedTechnician = "";

  function isAssignedToMe() {
    const me = getUser().trim().toLowerCase();
    const assigned = String(assignedTechnician || "").trim().toLowerCase();
    if (!me || !assigned) return false;
    if (me === assigned) return true;
    return false;
  }

  function renderActions() {
    const startBtn = qs("woStartBtn");
    const completeBtn = qs("woCompleteBtn");
    const notesWrap = qs("woNotesWrap");
    const st = String(currentStatus || "").toLowerCase();
    const role = getRole();
    const mine = isAssignedToMe();

    if (startBtn) {
      startBtn.style.display =
        ((role === "artisan" && st === "assigned" && mine) || (["admin", "supervisor"].includes(role) && st === "assigned"))
          ? "inline-block"
          : "none";
    }
    if (completeBtn) completeBtn.style.display =
      (role === "artisan" && st === "in_progress" && mine) || (["admin", "supervisor"].includes(role) && st === "in_progress")
        ? "inline-block"
        : "none";
    if (notesWrap) notesWrap.style.display = completeBtn && completeBtn.style.display !== "none" ? "block" : "none";

    const hint = qs("woActionHint");
    if (!hint) return;
    if (st === "assigned" && !mine && role === "artisan") {
      hint.textContent = `This job is assigned to ${assignedTechnician || "another technician"}.`;
      return;
    }
    if (st === "assigned" && mine) hint.textContent = "Tap Start job when you begin work.";
    else if (st === "assigned" && ["admin", "supervisor"].includes(role)) {
      hint.textContent = "Technician assigned. Tap Start job to move this work order in progress.";
    }
    else if (st === "in_progress" && (mine || ["admin", "supervisor"].includes(role))) {
      hint.textContent = "Add completion notes, then submit for supervisor approval.";
    } else if (st === "completed") hint.textContent = "Job submitted. Waiting for supervisor approval.";
    else if (st === "approved") hint.textContent = "Approved. Supervisor will close the work order.";
    else hint.textContent = role === "artisan" ? "No technician action available for this status." : "";
  }

  async function ensureSession() {
    if (!A) return;
    const user = await A.trySession();
    if (!user && !getUser()) {
      const woId = getWoId();
      const ret = woId ? `?wo_id=${woId}` : "";
      window.location.href = `./technician-terminal.html${ret}`;
      return false;
    }
    return true;
  }

  async function loadWoProfile() {
    const ok = await ensureSession();
    if (ok === false) return;
    const woId = getWoId();
    if (!woId) { setText("woQrSub", "Missing wo_id in QR URL."); return; }
    setText("woQrSub", `Loading WO #${woId}...`);
    const detail = await fetchJson(`/api/workorders/${woId}`, { headers: headers() });
    const wo = detail?.work_order || {};
    setText("woQrSub", `Work order #${woId} loaded`);
    setText("woId", String(wo.id || woId));
    currentStatus = String(wo.status || "");
    assignedTechnician = String(wo.assigned_artisan_name || "");
    const stEl = qs("woStatus");
    if (stEl) { stEl.textContent = currentStatus.toUpperCase() || "-"; stEl.className = statusClass(currentStatus); }
    setText("woSource", String(wo.source || "-"));
    setText("woAsset", `${String(wo.asset_code || "-")} - ${String(wo.asset_name || "-")}`);
    setText("woMakeModel", `${String(detail?.asset?.make || wo.make || "-")} / ${String(detail?.asset?.model || wo.model || "-")}`);
    setText("woOpened", String(wo.opened_at || "-"));
    setText("woAssigned", assignedTechnician || "Unassigned");
    if (qs("woArtisan")) qs("woArtisan").value = getUser();
    renderActions();
  }

  async function startJob() {
    const woId = getWoId();
    if (!woId) return;
    setText("woActionMsg", "Starting job...");
    await fetchJson(`/api/workorders/${woId}/status`, {
      method: "POST",
      headers: headers({ "Content-Type": "application/json" }),
      body: JSON.stringify({ status: "in_progress", artisan_name: getUser() }),
    });
    setText("woActionMsg", "Job started.");
    await loadWoProfile();
  }

  async function completeJob() {
    const woId = getWoId();
    if (!woId) return;
    const notes = String(qs("woNotes")?.value || "").trim();
    const artisan = String(qs("woArtisan")?.value || getUser()).trim();
    if (!notes) throw new Error("Completion notes are required.");
    setText("woActionMsg", "Submitting completion...");
    await fetchJson(`/api/workorders/${woId}/status`, {
      method: "POST",
      headers: headers({ "Content-Type": "application/json" }),
      body: JSON.stringify({ status: "completed", completion_notes: notes, artisan_name: artisan }),
    });
    setText("woActionMsg", "Job marked complete. Waiting for supervisor approval.");
    await loadWoProfile();
  }

  qs("woQrRefresh")?.addEventListener("click", () => loadWoProfile().catch((e) => setText("woQrSub", `Error: ${e.message || e}`)));
  qs("woStartBtn")?.addEventListener("click", () => startJob().catch((e) => setText("woActionMsg", `Error: ${e.message || e}`)));
  qs("woCompleteBtn")?.addEventListener("click", () => completeJob().catch((e) => setText("woActionMsg", `Error: ${e.message || e}`)));
  loadWoProfile().catch((e) => setText("woQrSub", `Error: ${e.message || e}`));
})();
