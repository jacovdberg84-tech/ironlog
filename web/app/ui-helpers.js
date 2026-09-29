// IRONLOG/web/app/ui-helpers.js — Toasts, status, formatting helpers, thresholds, offline queues.
// Part of the main app; index.html loads these files in order and they share one global scope.

let lastToastKey = "";
let lastToastAt = 0;

function statusToastTone(message) {
  const text = String(message || "").toLowerCase();
  if (/error|failed|cannot|blocked|invalid|denied/.test(text)) return "error";
  if (/warning|with issues|queued for sync/.test(text)) return "warning";
  if (/saved|created|updated|uploaded|downloaded|removed|added|sent|copied|complete\s*✅|successfully/.test(text)) return "success";
  return "";
}

function showToast(message, tone = "info") {
  const region = qs("toastRegion");
  const text = String(message || "").trim();
  if (!region || !text) return;
  const key = `${tone}:${text}`;
  const now = Date.now();
  if (key === lastToastKey && now - lastToastAt < 1500) return;
  lastToastKey = key;
  lastToastAt = now;
  const toast = document.createElement("div");
  toast.className = `app-toast ${tone}`;
  toast.setAttribute("role", tone === "error" ? "alert" : "status");
  toast.textContent = text;
  region.replaceChildren(toast);
  const delay = tone === "error" ? 7000 : 4200;
  window.setTimeout(() => {
    toast.classList.add("is-leaving");
    window.setTimeout(() => toast.remove(), 180);
  }, delay);
}

function setStatus(msg) {
  const el = qs("status");
  if (!el) return;
  const translated = translateStatusMessage(msg, getLang());
  el.textContent = translated;
  const tone = statusToastTone(translated);
  if (tone) showToast(translated, tone);
}
function setText(id, value) {
  const el = qs(id);
  if (!el) return;
  el.textContent = value;
}
function fmtMoney(v) {
  const n = Number(v || 0);
  if (!Number.isFinite(n)) return "0.00";
  return n.toFixed(2);
}

async function transitionWorkOrderStatus(id, toStatus) {
  const woId = Number(id || 0);
  const status = String(toStatus || "").trim().toLowerCase();
  if (!woId || !status) return;
  await fetchJson(`${API}/api/workorders/${woId}/status`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ status }),
  });
}

async function nudgeSupervisor(woId) {
  const id = Number(woId || 0);
  if (!id) return;
  const note = "SLA escalation nudge from dashboard";
  await fetchJson(`${API}/api/dashboard/workorders/${id}/nudge`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ note }),
  });
}
function setHtml(id, html) {
  const el = qs(id);
  if (!el) return;
  el.innerHTML = html;
}

function setSkeleton(id, blocks = 1) {
  const el = qs(id);
  if (!el) return;
  el.innerHTML = Array.from({ length: blocks })
    .map(() => `<div class="skeleton-block"></div>`)
    .join("");
}
// --- MUST be first: dashboard list helper ---
function item(html) {
  const d = document.createElement("div");
  d.className = "item";
  d.innerHTML = html;
  return d;
}
function setSpeedo(needleEl, valEl, pct, opts) {
  if (!valEl) return;

  const face = needleEl ? needleEl.parentElement : null; // .speedo-face
  const clearKpiClasses = (el) => {
    if (!el) return;
    el.classList.remove("kpi-good", "kpi-warn", "kpi-bad");
  };
  const barFill =
    valEl.id === "gAvailVal"
      ? qs("gAvailBarFill")
      : (valEl.id === "gUtilVal" ? qs("gUtilBarFill") : null);
  const clearBarClasses = () => {
    if (!barFill) return;
    barFill.classList.remove("kpi-good", "kpi-warn", "kpi-bad");
  };

  // N/A state
  if (pct == null || Number.isNaN(pct)) {
    clearKpiClasses(needleEl);
    clearKpiClasses(face);
    clearBarClasses();
    if (barFill) barFill.style.width = "0%";
    if (needleEl) needleEl.style.transform = "translateX(-50%) rotate(-90deg)";
    valEl.textContent = "N/A";
    return;
  }

  const clamped = Math.max(0, Math.min(100, Number(pct)));
  const deg = -90 + (clamped * 180) / 100;

  // KPI lighting thresholds
  // KPI lighting thresholds (configurable via setSpeedo opts)
  const _goodAt = Number(opts?.goodAt ?? 85);
  const _warnAt = Number(opts?.warnAt ?? 60);
  let kpiClass = "kpi-bad";
  if (clamped >= _goodAt) kpiClass = "kpi-good";
  else if (clamped >= _warnAt) kpiClass = "kpi-warn";

  clearKpiClasses(needleEl);
  clearKpiClasses(face);
  if (needleEl) needleEl.classList.add(kpiClass);
  if (face) face.classList.add(kpiClass);
  clearBarClasses();
  if (barFill) {
    barFill.classList.add(kpiClass);
    barFill.style.width = `${clamped.toFixed(2)}%`;
  }

  // Needle sweep on first render (per needle)
  if (needleEl && !needleEl.dataset.swept) {
    needleEl.dataset.swept = "1";
    needleEl.style.transform = "translateX(-50%) rotate(-90deg)";
    // next frame -> sweep to target
    requestAnimationFrame(() => {
      needleEl.style.transform = `translateX(-50%) rotate(${deg}deg)`;
    });
  } else if (needleEl) {
    needleEl.style.transform = `translateX(-50%) rotate(${deg}deg)`;
  }

  valEl.textContent = clamped.toFixed(2) + "%";
}

function getThresholds() {
  const safeNum = (k, def) => {
    const v = Number(localStorage.getItem(k));
    return Number.isFinite(v) && v >= 0 && v <= 100 ? v : def;
  };
  return {
    availTarget: safeNum("th_avail_target", 85),
    availCrit:   safeNum("th_avail_crit",   70),
    utilTarget:  safeNum("th_util_target",  70),
    utilCrit:    safeNum("th_util_crit",    55),
  };
}

function populateThresholdInputs() {
  const th = getThresholds();
  const set = (id, v) => { const el = qs(id); if (el) el.value = v; };
  set("thAvailTarget", th.availTarget);
  set("thAvailCrit",   th.availCrit);
  set("thUtilTarget",  th.utilTarget);
  set("thUtilCrit",    th.utilCrit);
}

function saveThresholdsFromUI() {
  const getNum = (id, def) => {
    const v = Number(qs(id)?.value);
    return Number.isFinite(v) && v >= 0 && v <= 100 ? v : def;
  };
  localStorage.setItem("th_avail_target", getNum("thAvailTarget", 85));
  localStorage.setItem("th_avail_crit",   getNum("thAvailCrit",   70));
  localStorage.setItem("th_util_target",  getNum("thUtilTarget",  70));
  localStorage.setItem("th_util_crit",    getNum("thUtilCrit",    55));
  setStatus("Thresholds saved.");
  loadDashboard().catch(() => {});
}

function getLdvPrestartThresholds() {
  const safeNum = (k, def) => {
    const v = Number(localStorage.getItem(k));
    return Number.isFinite(v) && v >= 0 && v <= 100 ? v : def;
  };
  const greenAt = safeNum("th_ldv_green_at", 95);
  const warnAtRaw = safeNum("th_ldv_warn_at", 80);
  const warnAt = Math.min(warnAtRaw, greenAt);
  return { greenAt, warnAt };
}

function populateLdvPrestartThresholdInputs() {
  const th = getLdvPrestartThresholds();
  const set = (id, v) => { const el = qs(id); if (el) el.value = v; };
  set("ldvGreenAt", th.greenAt);
  set("ldvWarnAt", th.warnAt);
}

function saveLdvPrestartThresholdsFromUI() {
  const getNum = (id, def) => {
    const v = Number(qs(id)?.value);
    return Number.isFinite(v) && v >= 0 && v <= 100 ? v : def;
  };
  const greenAt = getNum("ldvGreenAt", 95);
  const warnAt = Math.min(getNum("ldvWarnAt", 80), greenAt);
  localStorage.setItem("th_ldv_green_at", String(greenAt));
  localStorage.setItem("th_ldv_warn_at", String(warnAt));
  setStatus("LDV compliance thresholds saved.");
  loadDashboard().catch(() => {});
}

function updateKpiAlertBanner(availPct, utilPct) {
  const banner = qs("kpiAlertBanner");
  if (!banner) return;
  const th = getThresholds();
  const issues = [];
  if (availPct != null && !Number.isNaN(Number(availPct))) {
    const a = Number(availPct);
    if (a < th.availCrit) {
      issues.push({ label: "AVAILABILITY CRITICAL", value: a, target: th.availTarget, cls: "kpi-alert-crit" });
    } else if (a < th.availTarget) {
      issues.push({ label: "AVAILABILITY BELOW TARGET", value: a, target: th.availTarget, cls: "kpi-alert-warn" });
    }
  }
  if (utilPct != null && !Number.isNaN(Number(utilPct))) {
    const u = Number(utilPct);
    if (u < th.utilCrit) {
      issues.push({ label: "UTILIZATION CRITICAL", value: u, target: th.utilTarget, cls: "kpi-alert-crit" });
    } else if (u < th.utilTarget) {
      issues.push({ label: "UTILIZATION BELOW TARGET", value: u, target: th.utilTarget, cls: "kpi-alert-warn" });
    }
  }
  if (!issues.length) {
    banner.style.display = "none";
    banner.innerHTML = "";
    return;
  }
  banner.style.display = "";
  banner.innerHTML = issues
    .map(
      (i) =>
        `<div class="kpi-alert-item ${i.cls}">` +
        `<span class="kpi-alert-icon">${i.cls === "kpi-alert-crit" ? "\u26D4" : "\u26A0\uFE0F"}</span>` +
        `<span class="kpi-alert-text"><b>${escapeHtml(i.label)}</b> \u2014 ${Number(i.value).toFixed(1)}% (target ${i.target}%)</span>` +
        `</div>`
    )
    .join("");
}

/* =========================
   OFFLINE QUEUE STORAGE
========================= */

const OFFLINE_KEY = "ironlog_offline_queue";

function getQueue() {
  return JSON.parse(localStorage.getItem(OFFLINE_KEY) || "[]");
}

function saveQueue(queue) {
  localStorage.setItem(OFFLINE_KEY, JSON.stringify(queue));
  refreshNetBanner();
}

/* =========================
   NET BANNER UI
========================= */

function setNetBanner(state, queuedCount) {
  const banner = qs("netBanner");
  const dot = qs("netDot");
  const text = qs("netText");
  const q = qs("qCount");
  const btn = qs("syncNow");
  if (!banner || !dot || !text || !q || !btn) return;

  const online = navigator.onLine;

  banner.classList.remove("offline", "syncing");
  if (!online) banner.classList.add("offline");
  if (state === "syncing") banner.classList.add("syncing");

  if (!online) text.textContent = "OFFLINE";
  else if (state === "syncing") text.textContent = "SYNCING...";
  else text.textContent = "ONLINE";

  const n = Number(queuedCount || 0);
  if (n > 0) {
    q.style.display = "";
    q.textContent = `Queued: ${n}`;
  } else {
    q.style.display = "none";
  }

  if (online && n > 0) btn.style.display = "";
  else btn.style.display = "none";
}

function getQueuedHoursCount() {
  const queue = getQueue();
  return queue.filter((q) => q.type === "HOURS").length;
}

const QR_OFFLINE_QUEUE_CONFIG = {
  safety: {
    storageKey: "ironlog_safety_offline_queue_v1",
    label: "Safety inspection",
    endpoint: "/api/safety/inspections",
    codeField: "item_code",
    dateField: "inspection_date",
  },
  ldv: {
    storageKey: "ironlog_ldv_prestart_offline_queue_v1",
    label: "LDV pre-start",
    endpoint: "/api/maintenance/vehicle-ldv-checks/prestart",
    codeField: "asset_code",
    dateField: "check_date",
  },
  machine: {
    storageKey: "ironlog_machine_prestart_offline_queue_v1",
    label: "Machine pre-start",
    endpoint: "/api/maintenance/machine-prestart",
    codeField: "asset_code",
    dateField: "check_date",
  },
};

function readNamedOfflineQueue(storageKey) {
  try {
    const rows = JSON.parse(localStorage.getItem(storageKey) || "[]");
    return Array.isArray(rows) ? rows : [];
  } catch {
    return [];
  }
}

function saveNamedOfflineQueue(storageKey, rows) {
  localStorage.setItem(storageKey, JSON.stringify(Array.isArray(rows) ? rows : []));
}

function summarizeQrOfflineQueues() {
  const groups = {};
  const items = [];
  for (const [type, cfg] of Object.entries(QR_OFFLINE_QUEUE_CONFIG)) {
    const rows = readNamedOfflineQueue(cfg.storageKey);
    groups[type] = rows.length;
    for (const row of rows) {
      const payload = row?.payload || {};
      items.push({
        type,
        label: cfg.label,
        key: String(row?.key || ""),
        code: String(payload[cfg.codeField] || "-").toUpperCase(),
        date: String(payload[cfg.dateField] || "-"),
        queued_at: String(row?.created_at || ""),
        inspector: String(payload.inspector_name || "").trim(),
      });
    }
  }
  items.sort((a, b) => String(b.queued_at).localeCompare(String(a.queued_at)));
  return { groups, items };
}

function getTotalQueuedCount() {
  const qr = summarizeQrOfflineQueues();
  const qrTotal = Object.values(qr.groups).reduce((sum, n) => sum + Number(n || 0), 0);
  return getQueuedHoursCount() + qrTotal;
}

function refreshNetBanner() {
  setNetBanner("idle", getTotalQueuedCount());
}

function setOfflineQueueAdminResult(text) {
  const pre = qs("offlineQueueResult");
  if (pre) pre.textContent = String(text || "");
}

function renderOfflineQueueAdminPanel() {
  const summaryHost = qs("offlineQueueSummary");
  const listHost = qs("offlineQueueList");
  if (!summaryHost || !listHost) return;

  const hoursCount = getQueuedHoursCount();
  const { groups, items } = summarizeQrOfflineQueues();
  const total = hoursCount + items.length;
  const online = navigator.onLine;

  summaryHost.innerHTML = `
    <span class="pill ${online ? "green" : "orange"}">${online ? "Online" : "Offline"}</span>
    <span class="pill blue">Total queued: ${total}</span>
    <span class="pill">Safety: ${Number(groups.safety || 0)}</span>
    <span class="pill">LDV: ${Number(groups.ldv || 0)}</span>
    <span class="pill">Machine: ${Number(groups.machine || 0)}</span>
    <span class="pill">Daily hours: ${hoursCount}</span>
  `;

  if (!total) {
    listHost.innerHTML = `<div class="muted small">No offline submissions on this device.</div>`;
    return;
  }

  const rows = [];
  if (hoursCount) {
    const hoursItems = getQueue().filter((q) => q.type === "HOURS");
    for (const row of hoursItems) {
      const payload = row?.payload || {};
      rows.push(`
        <div class="item" style="display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap; padding:8px 0; border-bottom:1px solid rgba(148,163,184,0.2);">
          <div>
            <strong>Daily hours</strong>
            <div class="muted small">${escapeHtml(String(payload.asset_code || "-"))} · ${escapeHtml(String(payload.work_date || "-"))}</div>
          </div>
          <span class="pill orange" style="font-size:0.65rem;">QUEUED</span>
        </div>
      `);
    }
  }

  for (const row of items) {
    const when = row.queued_at ? new Date(row.queued_at).toLocaleString() : "—";
    rows.push(`
      <div class="item" style="display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap; padding:8px 0; border-bottom:1px solid rgba(148,163,184,0.2);">
        <div>
          <strong>${escapeHtml(row.label)}</strong>
          <div class="muted small">${escapeHtml(row.code)} · ${escapeHtml(row.date)}${row.inspector ? ` · ${escapeHtml(row.inspector)}` : ""}</div>
          <div class="muted small">Queued: ${escapeHtml(when)}</div>
        </div>
        <span class="pill orange" style="font-size:0.65rem;">QUEUED</span>
      </div>
    `);
  }

  listHost.innerHTML = rows.join("");
}

async function syncQrOfflineQueue(type) {
  const cfg = QR_OFFLINE_QUEUE_CONFIG[type];
  if (!cfg) return { synced: 0, failed: 0, remaining: 0 };
  const queue = readNamedOfflineQueue(cfg.storageKey);
  if (!queue.length) return { synced: 0, failed: 0, remaining: 0 };

  const remaining = [];
  let synced = 0;
  for (const row of queue) {
    const payload = row?.payload || null;
    if (!payload || typeof payload !== "object") continue;
    try {
      await fetchJson(`${API}${cfg.endpoint}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      synced += 1;
    } catch {
      remaining.push(row);
    }
  }
  saveNamedOfflineQueue(cfg.storageKey, remaining);
  return { synced, failed: queue.length - synced - remaining.length, remaining: remaining.length };
}

async function syncAllOfflineQueues() {
  if (!navigator.onLine) {
    return { ok: false, reason: "offline", synced: 0, remaining: getTotalQueuedCount() };
  }

  const totalBefore = getTotalQueuedCount();
  if (!totalBefore) {
    renderOfflineQueueAdminPanel();
    refreshNetBanner();
    return { ok: true, synced: 0, remaining: 0 };
  }

  setNetBanner("syncing", totalBefore);
  setStatus(`Syncing offline queue (${totalBefore})...`);

  let synced = 0;
  const hoursResult = await syncOfflineHoursQueue({ quiet: true });
  if (hoursResult?.synced) synced += Number(hoursResult.synced || 0);

  for (const type of Object.keys(QR_OFFLINE_QUEUE_CONFIG)) {
    const result = await syncQrOfflineQueue(type);
    synced += Number(result.synced || 0);
  }

  const remaining = getTotalQueuedCount();
  renderOfflineQueueAdminPanel();
  refreshNetBanner();

  if (remaining) {
    setOfflineQueueAdminResult(`Sync finished: ${synced} sent, ${remaining} still queued.`);
    setStatus(`Sync finished: ${synced} sent, ${remaining} still queued.`);
    return { ok: false, synced, remaining };
  }

  setOfflineQueueAdminResult(`Sync finished: all ${synced} queued item(s) sent.`);
  setStatus(`Sync finished: all ${synced} queued item(s) sent.`);
  return { ok: true, synced, remaining: 0 };
}

function initOfflineQueueAdminPanel() {
  if (!qs("adminOfflineQueueCard")) return;
  renderOfflineQueueAdminPanel();
  qs("offlineQueueRefreshBtn")?.addEventListener("click", () => {
    renderOfflineQueueAdminPanel();
    setOfflineQueueAdminResult("Queue list refreshed.");
  });
  qs("offlineQueueSyncBtn")?.addEventListener("click", () => {
    syncAllOfflineQueues().catch((e) => {
      setOfflineQueueAdminResult(String(e.message || e));
      setStatus("Sync error: " + (e.message || e));
      refreshNetBanner();
    });
  });
}

async function disableLegacyServiceWorkers() {
  if (!("serviceWorker" in navigator)) return;
  try {
    const regs = await navigator.serviceWorker.getRegistrations();
    if (!Array.isArray(regs) || !regs.length) return;
    for (const reg of regs) {
      await reg.unregister();
    }
  } catch {
    // non-fatal: app keeps working if SW API is blocked
  }
}

/* =========================
   OFFLINE QUEUE (HOURS ONLY)
========================= */

function hoursQueueKey(payload) {
  return `${payload.work_date}::${payload.asset_code}`;
}

function queueHours(payload) {
  const queue = getQueue();
  const key = hoursQueueKey(payload);

  // Deduplicate: keep only the latest entry per asset/day
  const filtered = queue.filter((q) => {
    if (q.type !== "HOURS") return true;
    return q.key !== key;
  });

  filtered.push({
    type: "HOURS",
    key,
    endpoint: "/api/hours",
    payload,
    timestamp: Date.now(),
  });

  saveQueue(filtered);
}

async function syncOfflineHoursQueue(opts = {}) {
  const quiet = Boolean(opts.quiet);
  if (!navigator.onLine) return { ok: false, reason: "offline" };

  const queue = getQueue();
  const hoursItems = queue.filter((q) => q.type === "HOURS");
  if (!hoursItems.length) return { ok: true, synced: 0 };

  if (!quiet) {
    setNetBanner("syncing", getTotalQueuedCount());
    setStatus(`Syncing offline queue (${hoursItems.length})...`);
  }

  const remaining = queue.filter((q) => q.type !== "HOURS");
  const failed = [];

  for (const item of hoursItems) {
    try {
      await fetchJson(`${API}${item.endpoint}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(item.payload),
      });
    } catch (e) {
      failed.push({ key: item.key, error: e.message || String(e) });
      remaining.push(item);
    }
  }

  saveQueue(remaining);

  if (failed.length) {
    if (!quiet) setStatus(`Sync finished: ${hoursItems.length - failed.length} ok, ${failed.length} failed.`);
    refreshNetBanner();
    return { ok: false, synced: hoursItems.length - failed.length, failed };
  }

  if (!quiet) setStatus("Sync finished: all queued hours synced ✅");
  refreshNetBanner();
  return { ok: true, synced: hoursItems.length };
}

async function postHoursWithOffline(payload) {
  if (!navigator.onLine) {
    queueHours(payload);
    return { queued: true };
  }

  return await fetchJson(`${API}/api/hours`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
}

/* =========================
   DASHBOARD / TELEMATICS
========================= */

/** Start-up: Offline banner and queued-data sync. Called once from init() in init.js. */
function wireOfflineSync() {
  // Net banner
  refreshNetBanner();
  window.addEventListener("offline", () => {
    refreshNetBanner();
    renderOfflineQueueAdminPanel();
  });

  window.addEventListener("online", async () => {
    refreshNetBanner();
    renderOfflineQueueAdminPanel();
    if (getTotalQueuedCount() === 0) return;
    try {
      await syncAllOfflineQueues();
    } catch (e) {
      setStatus("Sync error: " + (e.message || e));
      refreshNetBanner();
    }
  });

  qs("syncNow")?.addEventListener("click", async () => {
    if (!navigator.onLine) return alert("Still offline.");
    try {
      await syncAllOfflineQueues();
    } catch (e) {
      setStatus("Sync error: " + (e.message || e));
      refreshNetBanner();
    }
  });
}
