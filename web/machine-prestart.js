(function () {
  const MACHINE_OFFLINE_QUEUE_KEY = "ironlog_machine_prestart_offline_queue_v1";
  const MACHINE_CONTEXT_CACHE_KEY = "ironlog_machine_prestart_context_cache_v1";

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
  function safeDomId(key) {
    return `chk_${String(key || "").replace(/[^a-zA-Z0-9_]/g, "_")}`;
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
  function getCheckDateFromUrlOrToday() {
    const url = new URL(window.location.href);
    const d = String(url.searchParams.get("check_date") || "").trim();
    return /^\d{4}-\d{2}-\d{2}$/.test(d) ? d : todayYmd();
  }

  let currentAssetCode = "";
  let currentDate = getCheckDateFromUrlOrToday();
  let currentCheckId = 0;
  let currentPdfUrl = "";

  function queueKey(payload) {
    return `${String(payload.asset_code || "").toUpperCase()}::${String(payload.check_date || "")}`;
  }
  function getOfflineQueue() {
    const rows = getJsonStore(MACHINE_OFFLINE_QUEUE_KEY, []);
    return Array.isArray(rows) ? rows : [];
  }
  function saveOfflineQueue(rows) {
    setJsonStore(MACHINE_OFFLINE_QUEUE_KEY, Array.isArray(rows) ? rows : []);
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
    const cache = getJsonStore(MACHINE_CONTEXT_CACHE_KEY, {});
    if (!cache || typeof cache !== "object") return null;
    return cache[code] || null;
  }
  function writeContextCache(assetCode, data) {
    const code = String(assetCode || "").trim().toUpperCase();
    if (!code || !data || typeof data !== "object") return;
    const cache = getJsonStore(MACHINE_CONTEXT_CACHE_KEY, {});
    if (!cache || typeof cache !== "object") return;
    cache[code] = { cached_at: new Date().toISOString(), data };
    setJsonStore(MACHINE_CONTEXT_CACHE_KEY, cache);
  }
  function refreshOfflineBanner() {
    const el = qs("offlineState");
    if (!el) return;
    const queued = getOfflineQueue().length;
    if (navigator.onLine) {
      el.textContent = queued ? `${PC.t("online")} ${PC.t("queued", { n: queued })}` : PC.t("online");
      return;
    }
    el.textContent = queued ? `${PC.t("offline")} ${PC.t("queued", { n: queued })}` : PC.t("offline");
  }

  const PC = window.IronlogPrestart;

  function updateProgress(p) {
    const el = qs("pcProgress");
    if (!el) return;
    el.innerHTML = `<strong>${PC.t("progress", { a: p.answered, t: p.total })}</strong>${p.faults ? ` · <span class="pc-fault-count">${PC.t("faultCount", { n: p.faults })}</span>` : ""}`;
    const btn = qs("saveBtn");
    if (btn) btn.textContent = p.faults ? PC.t("submitFaults") : PC.t("submit");
    if (p.answered === p.total && qs("msg")?.classList.contains("err")) msg("");
  }

  let currentTemplate = null;

  function renderChecklist(template) {
    currentTemplate = template || currentTemplate;
    if (!currentTemplate) return;
    const pt = PC.lang() === "pt";
    txt("pageTitle", String((pt && currentTemplate.title_pt) || currentTemplate.title || PC.t("machinePrestart")));
    txt("sub", PC.t("tapEvery"));
    // Tippers log km, other machines hours; both go in the same reading field.
    const meterLabel = qs("meterLabel");
    if (meterLabel) meterLabel.dataset.i18n = currentTemplate.meter === "km" ? "odometerOpt" : "hourMeter";
    PC.translatePage();
    PC.render(qs("checklistRoot"), currentTemplate.sections || [], updateProgress);
    updateProgress(PC.progress(qs("checklistRoot")));
    refreshOfflineBanner();
  }

  function applyChecklist(checklist) {
    PC.apply(qs("checklistRoot"), checklist);
  }

  function showDone(kind, title, text) {
    const done = qs("doneArea");
    const form = qs("formArea");
    if (!done || !form) return;
    done.className = `card pc-done is-${kind}`;
    txt("doneMark", kind === "fault" ? "!" : "✓");
    txt("doneTitle", title);
    txt("doneText", text);
    const pdf = qs("openPdfBtn");
    if (pdf) {
      // Signed link from the server: opens this check's PDF without a login.
      pdf.hidden = !currentPdfUrl;
      pdf.href = currentPdfUrl || "#";
    }
    form.hidden = true;
    done.hidden = false;
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function showForm() {
    if (qs("doneArea")) qs("doneArea").hidden = true;
    if (qs("formArea")) qs("formArea").hidden = false;
  }

  function syncDateInput() {
    const inp = qs("checkDateInput");
    if (inp) inp.value = currentDate;
  }

  async function loadContext() {
    currentAssetCode = getAssetCodeFromUrl();
    if (!currentAssetCode) {
      txt("sub", "Missing asset_code in URL.");
      txt("assetCode", "-");
      return;
    }
    currentDate = qs("checkDateInput")?.value || currentDate || getCheckDateFromUrlOrToday();
    syncDateInput();
    txt("sub", PC.t("loading"));
    msg("");
    const q = new URLSearchParams();
    q.set("asset_code", currentAssetCode);
    q.set("check_date", currentDate);
    let data = null;
    try {
      data = await fetchJson(`/api/maintenance/machine-prestart/context?${q.toString()}`);
      writeContextCache(currentAssetCode, data);
    } catch (err) {
      const cached = readContextCache(currentAssetCode);
      if (!cached?.data && /no machine pre-start template/i.test(String(err.message || ""))) {
        txt("sub", "");
        txt("assetCode", currentAssetCode);
        throw new Error(PC.t("noTemplate"));
      }
      if (!cached?.data) throw err;
      data = cached.data;
      msg(`Offline: showing cached machine pre-start context from ${String(cached.cached_at || "earlier")}.`, "warn");
    }
    const asset = data?.asset || {};
    const template = data?.template || null;
    if (!template) throw new Error("No template returned from server.");

    txt("assetCode", String(asset.asset_code || currentAssetCode));
    txt("assetName", String(asset.asset_name || ""));
    document.title = `${String(asset.asset_code || currentAssetCode)} pre-start`;
    showForm();

    renderChecklist(template);
    PC.renderNotice(qs("noticeHost"), data?.notice || null);

    currentCheckId = 0;
    currentPdfUrl = "";
    const existing = data?.existing_check || null;
    if (existing?.id) {
      currentCheckId = Number(existing.id || 0);
      currentPdfUrl = String(existing.pdf_url || "");
      if (qs("smuHours") && existing.smu_hours != null) qs("smuHours").value = String(existing.smu_hours);
      if (qs("inspectorName")) qs("inspectorName").value = String(existing.inspector_name || "");
      const saved = PC.splitSavedNotes(existing.notes);
      if (qs("notes")) qs("notes").value = saved.notes;
      applyChecklist(existing.checklist);
      PC.applyFaultComments(qs("checklistRoot"), saved.comments);
      msg(PC.t("already"), "ok");
    } else {
      if (qs("smuHours")) qs("smuHours").value = "";
      if (qs("notes")) qs("notes").value = "";
      applyChecklist([]);
      PC.rememberOperator(qs("inspectorName"));
      msg("");
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
        await fetchJson("/api/maintenance/machine-prestart", {
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
    if (sent > 0 && !remaining.length) msg(`Synced ${sent} queued machine pre-start(s).`, "ok");
    else if (sent > 0) msg(`Synced ${sent} queued machine pre-start(s). ${remaining.length} still queued.`, "warn");
  }

  async function submitPrestart() {
    msg("");
    const root = qs("checklistRoot");
    const answers = PC.read(root);
    if (!root?.querySelector(".pc-item")) throw new Error("Checklist failed to load. Tap Reload.");
    root.querySelectorAll(".pc-item.is-missing").forEach((el) => el.classList.remove("is-missing"));
    if (answers.unanswered.length) {
      root.querySelectorAll(".pc-item:not([data-state=ok]):not([data-state=fault])").forEach((el) => el.classList.add("is-missing"));
      PC.firstUnanswered(root)?.scrollIntoView({ behavior: "smooth", block: "center" });
      throw new Error(PC.t("moreChecks", { n: answers.unanswered.length }));
    }
    const inspector_name = String(qs("inspectorName")?.value || "").trim();
    if (!inspector_name) {
      qs("inspectorName")?.focus();
      throw new Error(PC.t("enterName"));
    }

    const smuRaw = String(qs("smuHours")?.value || "").trim();
    let smu_hours = null;
    if (smuRaw) {
      const smu = Number(smuRaw);
      if (!Number.isFinite(smu) || smu < 0) throw new Error(PC.t("badNumber"));
      smu_hours = smu;
    }

    const notice_ack = PC.noticeAck(qs("noticeHost"));
    const body = {
      asset_code: currentAssetCode,
      check_date: currentDate,
      notice_ack,
      smu_hours,
      inspector_name,
      notes: String(qs("notes")?.value || "").trim(),
      checklist: answers.checklist,
      faults: answers.faults,
      lang: PC.lang(),
    };
    const faultCount = Object.keys(answers.faults).length;
    try { localStorage.setItem("ironlog_prestart_operator", inspector_name); } catch {}
    if (!navigator.onLine) {
      upsertOfflineQueue(body);
      refreshOfflineBanner();
      showDone(faultCount ? "fault" : "ok", PC.t("savedPhone"), `${PC.t("savedPhoneText")}${faultCount ? PC.t("offlineFault") : ""}`);
      return;
    }

    const btn = qs("saveBtn");
    if (btn) btn.disabled = true;
    try {
      const data = await fetchJson("/api/maintenance/machine-prestart", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      currentCheckId = Number(data?.id || 0);
      currentPdfUrl = String(data?.pdf_url || "");
      let photoNote = "";
      if (qs("photoInput")?.files?.[0]) {
        try {
          await uploadPhoto();
          photoNote = PC.t("photoAttached");
        } catch (e) {
          photoNote = `${PC.t("photoNot")}${e.message || e}`;
        }
      }
      if (Number(data?.faults || 0) > 0) {
        const list = Object.keys(answers.faults).map((k) => root.querySelector(`.pc-item[data-key="${CSS.escape(k)}"] .pc-label`)?.firstChild?.textContent || k).join(", ");
        const wo = data.fault_work_order_id ? PC.t("woRef", { id: data.fault_work_order_id }) : "";
        showDone("fault", PC.t("faultsTitle"), `${PC.t("faultsText", { n: data.faults, wo, list })}${photoNote}`);
      } else {
        showDone("ok", PC.t("doneTitle"), `${PC.t("doneText", { asset: currentAssetCode, date: currentDate })}${photoNote}`);
      }
      await syncOfflineQueue();
    } finally {
      if (btn) btn.disabled = false;
    }
  }

  async function uploadPhoto() {
    if (!currentCheckId) {
      throw new Error("Submit the pre-start first so a check record exists.");
    }
    const file = qs("photoInput")?.files?.[0];
    if (!file) throw new Error("Select a photo first.");
    const fd = new FormData();
    fd.append("file", file);
    const res = await fetch(`/api/maintenance/vehicle-ldv-checks/${encodeURIComponent(String(currentCheckId))}/photo?caption=${encodeURIComponent("Machine pre-start photo")}`, {
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
  }

  qs("checkDateInput")?.addEventListener("change", () => {
    currentDate = String(qs("checkDateInput")?.value || "").trim() || todayYmd();
    loadContext().catch((e) => msg(String(e.message || e), "err"));
  });

  qs("refreshBtn")?.addEventListener("click", () => {
    loadContext().catch((e) => msg(String(e.message || e), "err"));
  });
  qs("saveBtn")?.addEventListener("click", () => {
    submitPrestart().catch((e) => msg(String(e.message || e), "err"));
  });
  qs("editAgainBtn")?.addEventListener("click", showForm);
  window.addEventListener("online", () => {
    refreshOfflineBanner();
    syncOfflineQueue().catch(() => {});
  });
  window.addEventListener("offline", refreshOfflineBanner);

  PC.bindLanguageToggle(() => {
    renderChecklist();
    PC.renderNotice(qs("noticeHost"), qs("noticeHost")?.__notice || null);
  });
  loadContext().catch((e) => msg(String(e.message || e), "err"));
  refreshOfflineBanner();
  syncOfflineQueue().catch(() => {});
})();
