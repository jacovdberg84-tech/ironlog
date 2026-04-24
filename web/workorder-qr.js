(function () {
  function qs(id) { return document.getElementById(id); }
  function setText(id, value) { const el = qs(id); if (el) el.textContent = value; }
  function getRole() { return String(localStorage.getItem("ironlog_session_role") || "artisan").trim().toLowerCase() || "artisan"; }
  function getUser() { return String(localStorage.getItem("ironlog_session_user") || "qr-user").trim() || "qr-user"; }
  function headers(extra) { return { ...(extra || {}), "x-user-role": getRole(), "x-user-name": getUser() }; }
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
  async function fetchJson(url, options) {
    const res = await fetch(url, options);
    const text = await res.text();
    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch { data = null; }
    if (!res.ok) throw new Error(data?.error || data?.message || text || `Request failed (${res.status})`);
    return data || {};
  }

  async function loadWoProfile() {
    const woId = getWoId();
    if (!woId) { setText("woQrSub", "Missing wo_id in QR URL."); return; }
    setText("woQrSub", `Loading WO #${woId}...`);
    const data = await fetchJson(`/api/workorders/${woId}/qr-profile`, { headers: headers() });
    const p = data?.live_preview || data?.stored?.qr_payload || {};
    setText("woQrSub", `Work order #${woId} loaded`);
    setText("woId", String(p?.work_order?.id || woId));
    const st = String(p?.work_order?.status || "-");
    const stEl = qs("woStatus");
    if (stEl) { stEl.textContent = st.toUpperCase(); stEl.className = statusClass(st); }
    setText("woSource", String(p?.work_order?.source || "-"));
    setText("woAsset", `${String(p?.asset?.asset_code || "-")} - ${String(p?.asset?.asset_name || "-")}`);
    setText("woMakeModel", `${String(p?.asset?.make || "-")} / ${String(p?.asset?.model || "-")}`);
    setText("woOpened", String(p?.work_order?.opened_at || "-"));
  }

  async function submitUpdate() {
    const woId = getWoId();
    if (!woId) return;
    const nextStatus = String(qs("woNextStatus")?.value || "").trim().toLowerCase();
    const notes = String(qs("woNotes")?.value || "").trim();
    const artisan = String(qs("woArtisan")?.value || getUser()).trim();
    setText("woActionMsg", "Submitting update...");
    await fetchJson(`/api/workorders/${woId}/status`, {
      method: "POST",
      headers: headers({ "Content-Type": "application/json" }),
      body: JSON.stringify({ status: nextStatus }),
    });
    if (nextStatus === "completed") {
      await fetchJson(`/api/workorders/${woId}/request-close`, {
        method: "POST",
        headers: headers({ "Content-Type": "application/json" }),
        body: JSON.stringify({ completion_notes: notes || null, artisan_name: artisan || null }),
      });
    }
    setText("woActionMsg", `Update submitted: ${nextStatus}${nextStatus === "completed" ? " (approval requested)" : ""} ✅`);
    await loadWoProfile();
  }

  qs("woQrRefresh")?.addEventListener("click", () => loadWoProfile().catch((e) => setText("woQrSub", `Error: ${e.message || e}`)));
  qs("woSubmitUpdate")?.addEventListener("click", () => submitUpdate().catch((e) => setText("woActionMsg", `Error: ${e.message || e}`)));
  loadWoProfile().catch((e) => setText("woQrSub", `Error: ${e.message || e}`));
})();
