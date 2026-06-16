(function () {
  const LDV_OFFLINE_QUEUE_KEY = "ironlog_ldv_prestart_offline_queue_v1";
  const LDV_CONTEXT_CACHE_KEY = "ironlog_ldv_prestart_context_cache_v1";

  function qs(id) {
    return document.getElementById(id);
  }
  function txt(id, value) {
    const el = qs(id);
    if (el) el.textContent = value;
  }
  function msg(text, type) {
    const box = qs("msg");
    if (!box) return;
    if (!text) {
      box.className = "";
      box.innerHTML = "";
      return;
    }
    box.className = `msg ${type === "ok" ? "ok" : type === "warn" ? "warn" : "err"}`;
    box.textContent = text;
  }
  function getJsonStore(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return fallback;
      const parsed = JSON.parse(raw);
      return parsed == null ? fallback : parsed;
    } catch {
      return fallback;
    }
  }
  function setJsonStore(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch {
      /* ignore storage failures */
    }
  }
  function showSyncState(sync) {
    const el = qs("syncState");
    if (!el) return;
    if (!sync || sync.synced !== true) {
      el.style.display = "none";
      el.textContent = "";
      return;
    }
    const mode = String(sync.mode || "updated");
    const action = mode === "inserted" ? "created" : "updated";
    el.textContent =
      `Synced to Daily Input (${action}) - ${String(sync.work_date || "")}: ` +
      `open ${Number(sync.opening_km || 0).toFixed(1)} km, ` +
      `close ${Number(sync.closing_km || 0).toFixed(1)} km, ` +
      `run ${Number(sync.run_km || 0).toFixed(1)} km.`;
    el.style.display = "block";
  }
  async function fetchJson(url, options) {
    const res = await fetch(url, options);
    const t = await res.text();
    let data = {};
    try {
      data = t ? JSON.parse(t) : {};
    } catch {
      data = {};
    }
    if (!res.ok) throw new Error(data?.error || data?.message || t || `Request failed (${res.status})`);
    return data || {};
  }
  function getAssetCodeFromUrl() {
    const url = new URL(window.location.href);
    return String(url.searchParams.get("asset_code") || "").trim().toUpperCase();
  }
  function todayYmd() {
    return new Date().toISOString().slice(0, 10);
  }
  function readChecklist() {
    return {
      brakes_ok: Boolean(qs("chkBrakes")?.checked),
      lights_ok: Boolean(qs("chkLights")?.checked),
      tyres_ok: Boolean(qs("chkTyres")?.checked),
      oil_coolant_ok: Boolean(qs("chkOilCoolant")?.checked),
      leaks_damage_ok: Boolean(qs("chkLeaks")?.checked),
      safety_items_ok: Boolean(qs("chkSafety")?.checked),
    };
  }
  function applyChecklist(checklist) {
    const byKey = {};
    for (const item of Array.isArray(checklist) ? checklist : []) {
      byKey[String(item?.key || "")] = Boolean(item?.ok);
    }
    if (qs("chkBrakes")) qs("chkBrakes").checked = byKey.brakes_ok === true;
    if (qs("chkLights")) qs("chkLights").checked = byKey.lights_ok === true;
    if (qs("chkTyres")) qs("chkTyres").checked = byKey.tyres_ok === true;
    if (qs("chkOilCoolant")) qs("chkOilCoolant").checked = byKey.oil_coolant_ok === true;
    if (qs("chkLeaks")) qs("chkLeaks").checked = byKey.leaks_damage_ok === true;
    if (qs("chkSafety")) qs("chkSafety").checked = byKey.safety_items_ok === true;
  }

  let currentAssetCode = "";
  let currentDate = todayYmd();
  let previousKm = null;
  let currentCheckId = 0;

  function queueKey(payload) {
    return `${String(payload.asset_code || "").toUpperCase()}::${String(payload.check_date || "")}`;
  }
  function getOfflineQueue() {
    const rows = getJsonStore(LDV_OFFLINE_QUEUE_KEY, []);
    return Array.isArray(rows) ? rows : [];
  }
  function saveOfflineQueue(rows) {
    setJsonStore(LDV_OFFLINE_QUEUE_KEY, Array.isArray(rows) ? rows : []);
  }
  function upsertOfflineQueue(payload) {
    const key = queueKey(payload);
    const queue = getOfflineQueue().filter((x) => String(x?.key || "") !== key);
    queue.push({ key, created_at: new Date().toISOString(), payload });
    saveOfflineQueue(queue);
    return queue.length;
  }
  function readContextCache(assetCode) {
    const code = String(assetCode || "").trim().toUpperCase();
    if (!code) return null;
    const cache = getJsonStore(LDV_CONTEXT_CACHE_KEY, {});
    if (!cache || typeof cache !== "object") return null;
    return cache[code] || null;
  }
  function writeContextCache(assetCode, data) {
    const code = String(assetCode || "").trim().toUpperCase();
    if (!code || !data || typeof data !== "object") return;
    const cache = getJsonStore(LDV_CONTEXT_CACHE_KEY, {});
    if (!cache || typeof cache !== "object") return;
    cache[code] = { cached_at: new Date().toISOString(), data };
    setJsonStore(LDV_CONTEXT_CACHE_KEY, cache);
  }
  function refreshOfflineBanner() {
    const el = qs("offlineState");
    if (!el) return;
    const queued = getOfflineQueue().length;
    if (navigator.onLine) {
      el.textContent = queued ? `Online. ${queued} pre-start submission(s) waiting to sync.` : "Online.";
      return;
    }
    el.textContent = queued ? `Offline mode. ${queued} submission(s) queued.` : "Offline mode.";
  }

  async function loadContext() {
    currentAssetCode = getAssetCodeFromUrl();
    if (!currentAssetCode) {
      txt("sub", "Missing asset_code in URL.");
      return;
    }
    txt("sub", `Loading ${currentAssetCode}...`);
    msg("");
    showSyncState(null);
    const q = new URLSearchParams();
    q.set("asset_code", currentAssetCode);
    q.set("check_date", currentDate);
    let data = null;
    try {
      data = await fetchJson(`/api/maintenance/vehicle-ldv-checks/prestart-context?${q.toString()}`);
      writeContextCache(currentAssetCode, data);
    } catch (err) {
      const cached = readContextCache(currentAssetCode);
      if (!cached?.data) throw err;
      data = cached.data;
      msg(`Offline: showing cached pre-start context from ${String(cached.cached_at || "earlier")}.`, "warn");
    }
    const asset = data?.asset || {};
    previousKm = data?.previous_odometer_km == null ? null : Number(data.previous_odometer_km);

    txt("sub", "Complete pre-start before operating vehicle.");
    txt("assetCode", String(asset.asset_code || currentAssetCode));
    txt("assetName", String(asset.asset_name || "-"));
    txt("checkDate", String(data?.check_date || currentDate));
    txt("prevKm", previousKm == null ? "-" : `${previousKm.toFixed(1)} km`);

    const existing = data?.existing_prestart || null;
    if (existing) {
      currentCheckId = Number(existing.id || 0);
      if (qs("odometerKm") && existing.odometer_km != null) qs("odometerKm").value = String(existing.odometer_km);
      if (qs("inspectorName")) qs("inspectorName").value = String(existing.inspector_name || "");
      if (qs("notes")) qs("notes").value = String(existing.notes || "");
      applyChecklist(existing.checklist);
      msg("Pre-start already captured for today. You can update and resubmit if needed.", "ok");
    }
    const openPdfBtn = qs("openPdfBtn");
    if (openPdfBtn) {
      if (currentCheckId > 0) {
        openPdfBtn.style.display = "inline-block";
        openPdfBtn.href = `/api/reports/vehicle-ldv-check/${encodeURIComponent(String(currentCheckId))}.pdf`;
      } else {
        openPdfBtn.style.display = "none";
        openPdfBtn.href = "#";
      }
    }

    const openQrBtn = qs("openQrBtn");
    if (openQrBtn) openQrBtn.href = `./asset-qr.html?asset_code=${encodeURIComponent(currentAssetCode)}`;
    refreshOfflineBanner();
  }

  async function syncOfflineQueue() {
    if (!navigator.onLine) return;
    const queue = getOfflineQueue();
    if (!queue.length) return;
    const remaining = [];
    let sent = 0;
    for (const row of queue) {
      const payload = row?.payload || null;
      if (!payload || typeof payload !== "object") continue;
      try {
        await fetchJson("/api/maintenance/vehicle-ldv-checks/prestart", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        });
        sent += 1;
      } catch {
        remaining.push(row);
      }
    }
    saveOfflineQueue(remaining);
    refreshOfflineBanner();
    if (sent > 0 && !remaining.length) msg(`Synced ${sent} queued LDV pre-start(s).`, "ok");
    else if (sent > 0) msg(`Synced ${sent} queued LDV pre-start(s). ${remaining.length} still queued.`, "warn");
  }

  async function submitPrestart() {
    msg("");
    const odometerRaw = String(qs("odometerKm")?.value || "").trim();
    if (!odometerRaw) throw new Error("Enter current odometer KM.");
    const odometer = Number(odometerRaw);
    if (!Number.isFinite(odometer) || odometer < 0) throw new Error("Odometer must be a valid number >= 0.");
    if (previousKm != null && odometer < previousKm) {
      const ok = window.confirm(
        `KM ${odometer.toFixed(1)} is less than previous (${previousKm.toFixed(1)}). Submit pre-start anyway?`
      );
      if (!ok) return;
    } else if (odometer > 500000 || (previousKm != null && odometer > previousKm * 1.25 + 500)) {
      const ok = window.confirm(
        `KM ${odometer.toFixed(1)} looks unusually high. Submit pre-start anyway?`
      );
      if (!ok) return;
    }

    const checklist = readChecklist();
    const allChecked = Object.values(checklist).every(Boolean);
    if (!allChecked) throw new Error("Complete all pre-start checks before submitting.");

    const body = {
      asset_code: currentAssetCode,
      check_date: currentDate,
      odometer_km: odometer,
      inspector_name: String(qs("inspectorName")?.value || "").trim(),
      notes: String(qs("notes")?.value || "").trim(),
      checklist,
    };
    if (!navigator.onLine) {
      const queued = upsertOfflineQueue(body);
      refreshOfflineBanner();
      showSyncState(null);
      msg(`Offline: pre-start saved on device. Queue size ${queued}. It will sync automatically.`, "warn");
      return;
    }

    const data = await fetchJson("/api/maintenance/vehicle-ldv-checks/prestart", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const savedKm = data?.odometer_km == null ? odometer : Number(data.odometer_km);
    currentCheckId = Number(data?.id || 0);
    previousKm = Number.isFinite(savedKm) ? savedKm : previousKm;
    txt("prevKm", previousKm == null ? "-" : `${previousKm.toFixed(1)} km`);
    const openPdfBtn = qs("openPdfBtn");
    if (openPdfBtn && currentCheckId > 0) {
      openPdfBtn.style.display = "inline-block";
      openPdfBtn.href = `/api/reports/vehicle-ldv-check/${encodeURIComponent(String(currentCheckId))}.pdf`;
    }
    showSyncState(data?.daily_input_sync || null);
    msg(data?.message || "Pre-start submitted successfully. KM saved to IRONLOG.", data?.km_review_needed ? "err" : "ok");
    await syncOfflineQueue();
  }

  async function uploadPhoto() {
    if (!currentCheckId) {
      throw new Error("Submit pre-start first so a check record exists.");
    }
    const file = qs("photoInput")?.files?.[0];
    if (!file) throw new Error("Select a photo first.");
    const fd = new FormData();
    fd.append("file", file);
    const res = await fetch(`/api/maintenance/vehicle-ldv-checks/${encodeURIComponent(String(currentCheckId))}/photo?caption=${encodeURIComponent("Pre-start photo")}`, {
      method: "POST",
      body: fd,
    });
    const t = await res.text();
    let data = {};
    try {
      data = t ? JSON.parse(t) : {};
    } catch {
      data = {};
    }
    if (!res.ok) throw new Error(data?.error || t || `Upload failed (${res.status})`);
    msg("Photo uploaded to this pre-start check.", "ok");
  }

  qs("refreshBtn")?.addEventListener("click", () => {
    loadContext().catch((e) => msg(String(e.message || e), "err"));
  });
  qs("saveBtn")?.addEventListener("click", () => {
    submitPrestart().catch((e) => msg(String(e.message || e), "err"));
  });
  qs("uploadPhotoBtn")?.addEventListener("click", () => {
    uploadPhoto().catch((e) => msg(String(e.message || e), "err"));
  });
  window.addEventListener("online", () => {
    refreshOfflineBanner();
    syncOfflineQueue().catch(() => {});
  });
  window.addEventListener("offline", refreshOfflineBanner);

  loadContext().catch((e) => msg(String(e.message || e), "err"));
  refreshOfflineBanner();
  syncOfflineQueue().catch(() => {});
})();
