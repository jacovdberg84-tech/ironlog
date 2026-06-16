(function () {
  const SAFETY_OFFLINE_QUEUE_KEY = "ironlog_safety_offline_queue_v1";
  const SAFETY_CONTEXT_CACHE_KEY = "ironlog_safety_context_cache_v1";

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
      box.textContent = "";
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
      /* localStorage may be unavailable */
    }
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
  function getItemCodeFromUrl() {
    return String(new URL(window.location.href).searchParams.get("item_code") || "").trim().toUpperCase();
  }
  function getAssetCodeFromUrl() {
    return String(new URL(window.location.href).searchParams.get("asset_code") || "").trim().toUpperCase();
  }
  function getTemplateKeyFromUrl() {
    return String(new URL(window.location.href).searchParams.get("template_key") || "").trim().toLowerCase();
  }
  function todayYmd() {
    return new Date().toISOString().slice(0, 10);
  }
  function esc(s) {
    return String(s || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/"/g, "&quot;");
  }

  let currentItemCode = "";
  let currentDate = todayYmd();
  let templateItems = [];

  function queueKey(payload) {
    return `${String(payload.item_code || "").toUpperCase()}::${String(payload.inspection_date || "")}`;
  }
  function getOfflineQueue() {
    const rows = getJsonStore(SAFETY_OFFLINE_QUEUE_KEY, []);
    return Array.isArray(rows) ? rows : [];
  }
  function saveOfflineQueue(rows) {
    const clean = Array.isArray(rows) ? rows : [];
    setJsonStore(SAFETY_OFFLINE_QUEUE_KEY, clean);
  }
  function upsertOfflineQueue(payload) {
    const key = queueKey(payload);
    const queue = getOfflineQueue().filter((x) => String(x?.key || "") !== key);
    queue.push({
      key,
      created_at: new Date().toISOString(),
      payload,
    });
    saveOfflineQueue(queue);
    return queue.length;
  }
  function readContextCache(itemCode) {
    const code = String(itemCode || "").trim().toUpperCase();
    if (!code) return null;
    const cache = getJsonStore(SAFETY_CONTEXT_CACHE_KEY, {});
    if (!cache || typeof cache !== "object") return null;
    return cache[code] || null;
  }
  function writeContextCache(itemCode, data) {
    const code = String(itemCode || "").trim().toUpperCase();
    if (!code || !data || typeof data !== "object") return;
    const cache = getJsonStore(SAFETY_CONTEXT_CACHE_KEY, {});
    if (!cache || typeof cache !== "object") return;
    cache[code] = {
      cached_at: new Date().toISOString(),
      data,
    };
    setJsonStore(SAFETY_CONTEXT_CACHE_KEY, cache);
  }
  function validateChecklistForSubmit(checklist) {
    const rows = Array.isArray(checklist) ? checklist : [];
    const unanswered = rows.filter((r) => r?.ok == null);
    if (unanswered.length) {
      return `Complete all checklist items (${unanswered.length} unanswered).`;
    }
    const failedMissingNote = rows.find((r) => r?.ok === false && !String(r?.note || "").trim());
    if (failedMissingNote) {
      return `Add a note for failed item: ${String(failedMissingNote.label || failedMissingNote.key || "Checklist item")}`;
    }
    return "";
  }
  function refreshOfflineBanner() {
    const host = qs("offlineState");
    if (!host) return;
    const queued = getOfflineQueue().length;
    const online = navigator.onLine;
    if (online) {
      host.textContent = queued
        ? `Online. ${queued} offline submission(s) waiting to sync.`
        : "Online.";
      return;
    }
    host.textContent = queued
      ? `Offline mode. ${queued} submission(s) queued for sync.`
      : "Offline mode.";
  }

  function renderChecklist(checklist) {
    const root = qs("checklistRoot");
    if (!root) return;
    const byKey = {};
    for (const row of Array.isArray(checklist) ? checklist : []) {
      byKey[String(row.key || "")] = row;
    }
    root.innerHTML = templateItems.map((it) => {
      const row = byKey[it.key] || {};
      let val = "";
      if (row.ok === true) val = "ok";
      else if (row.ok === false) val = "fail";
      else if (row.ok == null && row.ok !== undefined) val = "na";
      const note = String(row.note || "");
      return `
        <div class="chk-row" data-key="${esc(it.key)}">
          <div>
            <div class="chk-label">${esc(it.label)}</div>
            <input class="chk-note" type="text" data-note-key="${esc(it.key)}" placeholder="Note (required if Fail)" value="${esc(note)}" />
          </div>
          <div class="chk-opts">
            <label><input type="radio" name="chk-${esc(it.key)}" value="ok" ${val === "ok" ? "checked" : ""} /> OK</label>
            <label><input type="radio" name="chk-${esc(it.key)}" value="fail" ${val === "fail" ? "checked" : ""} /> Fail</label>
            <label><input type="radio" name="chk-${esc(it.key)}" value="na" ${val === "na" ? "checked" : ""} /> N/A</label>
          </div>
        </div>
      `;
    }).join("");
  }

  function readChecklist() {
    return templateItems.map((it) => {
      const sel = document.querySelector(`input[name="chk-${CSS.escape(it.key)}"]:checked`);
      const val = sel ? String(sel.value) : "";
      let ok = null;
      if (val === "ok") ok = true;
      else if (val === "fail") ok = false;
      const noteEl = document.querySelector(`input[data-note-key="${CSS.escape(it.key)}"]`);
      const note = String(noteEl?.value || "").trim() || null;
      return { key: it.key, label: it.label, ok, note };
    });
  }

  async function loadContext() {
    currentItemCode = getItemCodeFromUrl();
    if (!currentItemCode) {
      txt("sub", "Missing item_code in URL. Scan the equipment QR label.");
      return;
    }
    txt("sub", `Loading ${currentItemCode}...`);
    msg("");
    const q = new URLSearchParams();
    q.set("item_code", currentItemCode);
    const assetCode = getAssetCodeFromUrl();
    const templateKey = getTemplateKeyFromUrl();
    if (assetCode) q.set("asset_code", assetCode);
    if (templateKey) q.set("template_key", templateKey);
    q.set("inspection_date", currentDate);
    let data = null;
    try {
      data = await fetchJson(`/api/safety/inspection-context?${q.toString()}`);
      writeContextCache(currentItemCode, data);
    } catch (err) {
      const cached = readContextCache(currentItemCode);
      if (!cached?.data) throw err;
      data = cached.data;
      msg(
        `Offline: showing cached checklist from ${String(cached.cached_at || "an earlier session")}.`,
        "warn"
      );
    }
    const item = data?.item || {};
    const template = data?.template || {};
    templateItems = Array.isArray(template.items) ? template.items : [];

    txt("pageTitle", String(template.title || "Safety Inspection"));
    txt("sub", "Complete all checklist items, then submit.");
    txt("itemCode", String(item.item_code || currentItemCode));
    txt("itemName", String(item.item_name || "-"));
    txt("itemLocation", String(item.location || "-"));
    txt("inspDate", String(data?.inspection_date || currentDate));

    const existing = data?.existing_inspection;
    if (qs("inspectorName")) {
      qs("inspectorName").value = String(existing?.inspector_name || "");
    }
    if (qs("notes")) {
      qs("notes").value = String(existing?.notes || "");
    }
    renderChecklist(data?.checklist || []);
    if (existing) {
      msg("Inspection already captured for this date. You can update and resubmit.", "ok");
    }
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
        await fetchJson("/api/safety/inspections", {
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
    if (sent > 0 && !remaining.length) {
      msg(`Synced ${sent} offline inspection(s).`, "ok");
    } else if (sent > 0) {
      msg(`Synced ${sent} inspection(s). ${remaining.length} still queued.`, "warn");
    }
  }

  async function submitInspection() {
    msg("");
    const checklist = readChecklist();
    const validationError = validateChecklistForSubmit(checklist);
    if (validationError) throw new Error(validationError);
    const body = {
      item_code: currentItemCode,
      inspection_date: currentDate,
      inspector_name: String(qs("inspectorName")?.value || "").trim(),
      notes: String(qs("notes")?.value || "").trim(),
      checklist,
    };
    if (!navigator.onLine) {
      const queued = upsertOfflineQueue(body);
      refreshOfflineBanner();
      msg(`Offline: inspection saved on this device. Queue size ${queued}. It will sync automatically when online.`, "warn");
      return;
    }

    const data = await fetchJson("/api/safety/inspections", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const st = String(data?.status || "pass");
    const label = st === "fail" ? "Failed - action required" : st === "attention" ? "Submitted with notes" : "Passed";
    msg(`Inspection saved: ${label}.`, "ok");
    await syncOfflineQueue();
  }

  qs("refreshBtn")?.addEventListener("click", () => loadContext().catch((e) => msg(e.message || String(e), "err")));
  qs("saveBtn")?.addEventListener("click", () => submitInspection().catch((e) => msg(e.message || String(e), "err")));
  window.addEventListener("online", () => {
    refreshOfflineBanner();
    syncOfflineQueue().catch(() => {});
  });
  window.addEventListener("offline", refreshOfflineBanner);

  loadContext().catch((e) => {
    txt("sub", "Load failed.");
    msg(e.message || String(e), "err");
  });
  refreshOfflineBanner();
  syncOfflineQueue().catch(() => {});
})();
