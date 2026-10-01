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
  const PC = window.IronlogPrestart;
  const LDV_CHECKS = [{
    title: "",
    items: [
      { key: "brakes_ok", label: "Brakes", label_pt: "Travões" },
      { key: "lights_ok", label: "Lights and indicators", label_pt: "Luzes e piscas" },
      { key: "tyres_ok", label: "Tyres (pressure and damage)", label_pt: "Pneus (pressão e danos)" },
      { key: "oil_coolant_ok", label: "Oil and coolant levels", label_pt: "Níveis de óleo e líquido de arrefecimento" },
      { key: "leaks_damage_ok", label: "No leaks or body damage", label_pt: "Sem fugas ou danos na carroçaria" },
      { key: "safety_items_ok", label: "Safety items (triangle, extinguisher, first aid)", label_pt: "Equipamento de segurança (triângulo, extintor, primeiros socorros)" },
    ],
  }];

  function updateProgress(p) {
    const el = qs("pcProgress");
    if (!el) return;
    el.innerHTML = `<strong>${PC.t("progress", { a: p.answered, t: p.total })}</strong>${p.faults ? ` · <span class="pc-fault-count">${PC.t("faultCount", { n: p.faults })}</span>` : ""}`;
    const btn = qs("saveBtn");
    if (btn) btn.textContent = p.faults ? PC.t("submitFaults") : PC.t("submit");
    if (p.answered === p.total && qs("msg")?.classList.contains("err")) msg("");
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

  let currentAssetCode = "";
  let currentDate = todayYmd();
  let previousKm = null;
  let currentCheckId = 0;
  let currentPdfUrl = "";

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
      el.textContent = queued ? `${PC.t("online")} ${PC.t("queued", { n: queued })}` : PC.t("online");
      return;
    }
    el.textContent = queued ? `${PC.t("offline")} ${PC.t("queued", { n: queued })}` : PC.t("offline");
  }

  async function loadContext() {
    currentAssetCode = getAssetCodeFromUrl();
    if (!currentAssetCode) {
      txt("sub", "Missing asset_code in URL.");
      return;
    }
    txt("sub", PC.t("loading"));
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

    txt("sub", PC.t("tapEvery"));
    txt("assetCode", String(asset.asset_code || currentAssetCode));
    txt("assetName", String(asset.asset_name || ""));
    document.title = `${String(asset.asset_code || currentAssetCode)} pre-start`;
    PC.render(qs("checklistRoot"), LDV_CHECKS, updateProgress);
    applyChecklist([]);
    PC.renderNotice(qs("noticeHost"), data?.notice || null);
    showForm();
    txt("checkDate", String(data?.check_date || currentDate));
    txt("prevKm", previousKm == null ? "-" : `${previousKm.toFixed(1)} km`);

    const existing = data?.existing_prestart || null;
    if (existing) {
      currentCheckId = Number(existing.id || 0);
      currentPdfUrl = String(existing.pdf_url || "");
      if (qs("odometerKm") && existing.odometer_km != null) qs("odometerKm").value = String(existing.odometer_km);
      if (qs("inspectorName")) qs("inspectorName").value = String(existing.inspector_name || "");
      const saved = PC.splitSavedNotes(existing.notes);
      if (qs("notes")) qs("notes").value = saved.notes;
      applyChecklist(existing.checklist);
      PC.applyFaultComments(qs("checklistRoot"), saved.comments);
      msg(PC.t("already"), "ok");
    } else {
      PC.rememberOperator(qs("inspectorName"));
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
    if (!odometerRaw) throw new Error(PC.t("enterKm"));
    const odometer = Number(odometerRaw);
    if (!Number.isFinite(odometer) || odometer < 0) throw new Error(PC.t("badNumber"));
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

    const root = qs("checklistRoot");
    const answers = PC.read(root);
    root?.querySelectorAll(".pc-item.is-missing").forEach((el) => el.classList.remove("is-missing"));
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

    const notice_ack = PC.noticeAck(qs("noticeHost"));
    const body = {
      asset_code: currentAssetCode,
      check_date: currentDate,
      notice_ack,
      odometer_km: odometer,
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
      showSyncState(null);
      showDone(faultCount ? "fault" : "ok", PC.t("savedPhone"), `${PC.t("savedPhoneText")}${faultCount ? PC.t("offlineFault") : ""}`);
      return;
    }

    const btn = qs("saveBtn");
    if (btn) btn.disabled = true;
    try {
      const data = await fetchJson("/api/maintenance/vehicle-ldv-checks/prestart", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const savedKm = data?.odometer_km == null ? odometer : Number(data.odometer_km);
      currentCheckId = Number(data?.id || 0);
      currentPdfUrl = String(data?.pdf_url || "");
      previousKm = Number.isFinite(savedKm) ? savedKm : previousKm;
      txt("prevKm", previousKm == null ? "-" : `${previousKm.toFixed(1)} km`);
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
        showDone("fault", PC.t("faultsTitle"), `${PC.t("faultsText", { n: data.faults, wo, list })}${data?.km_review_needed ? ` ${PC.t("kmText")}` : ""}${photoNote}`);
      } else if (data?.km_review_needed) {
        showDone("fault", PC.t("kmTitle"), `${PC.t("kmText")}${photoNote}`);
      } else {
        showDone("ok", PC.t("doneTitle"), `${PC.t("doneText", { asset: currentAssetCode, date: currentDate })}${photoNote}`);
      }
      showSyncState(data?.daily_input_sync || null);
      await syncOfflineQueue();
    } finally {
      if (btn) btn.disabled = false;
    }
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
  }

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
    PC.render(qs("checklistRoot"), LDV_CHECKS, updateProgress);
    PC.renderNotice(qs("noticeHost"), qs("noticeHost")?.__notice || null);
    refreshOfflineBanner();
    if (qs("sub") && qs("checklistRoot")?.querySelector(".pc-item")) txt("sub", PC.t("tapEvery"));
  });
  loadContext().catch((e) => msg(String(e.message || e), "err"));
  refreshOfflineBanner();
  syncOfflineQueue().catch(() => {});
})();
