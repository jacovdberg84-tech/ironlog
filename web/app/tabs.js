// IRONLOG/web/app/tabs.js — Tab switching, in-app help, section toggles.
// Part of the main app; index.html loads these files in order and they share one global scope.

function switchTab(key) {
  const k = String(key || "").trim();
  if (!k) return;
  if (k === "maintenance") {
    location.href = "maintenance.html";
    return;
  }
  document.querySelectorAll(".panel").forEach((p) => p.classList.remove("show"));
  const panel = qs(`tab-${k}`);
  if (panel) panel.classList.add("show");
  const tabSelect = qs("tabSelect");
  if (tabSelect && tabSelect.value !== k) tabSelect.value = k;
  if (k === "Breakdowns") {
    const today = new Date().toISOString().slice(0, 10);
    if (qs("boEnsureDate") && !qs("boEnsureDate").value) qs("boEnsureDate").value = today;
    if (qs("boSlipDate") && !qs("boSlipDate").value) qs("boSlipDate").value = today;
    initBoTyreRows();
    updateBoSlipFormVisibility();
    refreshBreakdownOpsPanels();
    loadBoSlipSavedList().catch(() => {});
  }
  updateSidebarActiveState(k);
  if (k === "finance") {
    try {
      initFinanceTab();
      loadFinanceSiteAllocation().catch(() => {});
    } catch (_) {}
  }
  if (k === "assets") {
    loadAssetsFleet().catch(() => {});
  }
  if (k === "mywork" && typeof loadMyWork === "function") {
    loadMyWork().catch(() => {});
  }
  if (k === "workshop") {
    loadWorkshopDocuments().catch(() => {});
  }
  if (k === "vehicle") {
    loadChecklistHub().catch(() => {});
    loadClHistory().catch(() => {});
  }
  if (k === "admin") {
    renderOfflineQueueAdminPanel();
    ensureAdminTabOptions().catch(() => {});
  }
  if (k === "telematics") {
    loadTelematicsTab().catch(() => {});
  }
  if (k === "stock") {
    loadStoresPartOrders().catch(() => {});
  }
  if (k === "parts-tracking") {
    loadPartsTrackingTab().catch(() => {});
  }
  if (k === "fuel") {
    initFamsFuelCatchupDates();
    loadFamsFuelStatus().catch(() => {});
  }
  if (k === "lube") {
    loadLubeUsage().catch(() => {});
  }
  if (k === "cartrack") {
    loadCartrackTrackingTab({ refresh: true }).catch(() => {});
    setTimeout(() => cartrackMap?.invalidateSize?.(), 120);
  } else {
    stopCartrackTrackPolling();
  }
  try {
    if (typeof window.updateIronmindHelpFabContext === "function") window.updateIronmindHelpFabContext();
  } catch (_) {}
}

/** Current sidebar panel key (e.g. daily, dash, Breakdowns). */
function getCurrentDashboardTabKey() {
  const panel =
    document.querySelector("#mainContent .panel.show") || document.querySelector(".panel.show");
  if (panel && panel.id && panel.id.startsWith("tab-")) return panel.id.slice(4);
  const sel = qs("tabSelect");
  return sel && sel.value ? sel.value : "dash";
}

const IRONLOG_HELP_OPENERS = {
  mywork: "Need help with your My Work list?",
  dash: "Need help with the dashboard?",
  daily: "Need help with daily inputs?",
  assets: "Need help with assets?",
  workshop: "Need help with the Workshop Library?",
  fuel: "Need help with fuel logging and benchmarks?",
  lube: "Need help with lube?",
  stock: "Need help with stores / stock?",
  "parts-tracking": "Need help with parts tracking and off-site repairs?",
  legal: "Need help with legal documents?",
  uploads: "Need help with CSV uploads?",
  reports: "Need help with reports?",
  approvals: "Need help with approvals?",
  procurement: "Need help with supply flow?",
  operations: "Need help with site operations?",
  dispatch: "Need help with dispatch?",
  quality: "Need help with data quality?",
  audit: "Need help with the audit trail?",
  vehicle: "Need help with daily checklists?",
  telematics: "Need help with telematics units and faults?",
  cartrack: "Need help with live Cartrack fleet tracking?",
  admin: "Need help with user admin?",
  docs: "Need help with AI documents?",
  ironmind:
    "Borris analyses fleet data here — use this Help button only for how to use IRONLOG screens.",
  finance: "Need help with finance?",
  enterprise: "Need help with enterprise views?",
  exec: "Need help with executive dashboards?",
  tasks: "Need help with tasks?",
  Breakdowns: "Need help with breakdowns?",
  breakdowns: "Need help with breakdowns?",
};

function getIronlogHelpOpenerForTab(tabKey) {
  const k = String(tabKey || "").trim();
  if (IRONLOG_HELP_OPENERS[k]) return IRONLOG_HELP_OPENERS[k];
  const lower = k.toLowerCase();
  if (IRONLOG_HELP_OPENERS[lower]) return IRONLOG_HELP_OPENERS[lower];
  return `Need help with this section (${k})?`;
}

let ironmindHelpHistory = [];

window.updateIronmindHelpFabContext = function updateIronmindHelpFabContext() {
  const openerEl = qs("ironmindHelpOpener");
  const tab = getCurrentDashboardTabKey();
  const text = getIronlogHelpOpenerForTab(tab);
  if (openerEl) openerEl.textContent = text;
  const ctxEl = qs("ironmindHelpContextLabel");
  if (ctxEl) ctxEl.textContent = tab;
};

async function sendIronmindHelpMessage() {
  const input = qs("ironmindHelpInput");
  const out = qs("ironmindHelpAnswer");
  const q = (input?.value || "").trim();
  if (!q) return;
  setStatus("Getting help…");
  const tab = getCurrentDashboardTabKey();
  try {
    const data = await fetchJson(`${API}/api/ironmind/help`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        question: q,
        context_key: tab,
        history: ironmindHelpHistory.slice(-6),
      }),
    });
    const ans = String(data.answer || "");
    if (out) {
      const block = `Q: ${q}\n\n${ans}`;
      out.textContent = out.textContent ? `${out.textContent}\n\n---\n\n${block}` : block;
    }
    ironmindHelpHistory.push({ question: q, answer: ans });
    if (input) input.value = "";
    setStatus(`Help ready (${data.mode === "live_ai" ? "AI" : "guide"}).`);
  } catch (e) {
    setStatus("Help error: " + (e.message || e));
  }
}

function initIronmindHelpUi() {
  const fab = qs("ironmindHelpFab");
  const modal = qs("ironmindHelpModal");
  const backdrop = qs("ironmindHelpBackdrop");
  const closeBtn = qs("ironmindHelpClose");
  const sendBtn = qs("ironmindHelpSend");
  const clearBtn = qs("ironmindHelpClear");
  const input = qs("ironmindHelpInput");

  function closeModal() {
    modal?.classList.remove("show");
    modal?.setAttribute("aria-hidden", "true");
  }

  fab?.addEventListener("click", () => {
    window.updateIronmindHelpFabContext?.();
    if (qs("ironmindHelpAnswer")) qs("ironmindHelpAnswer").textContent = "";
    ironmindHelpHistory = [];
    modal?.classList.add("show");
    modal?.setAttribute("aria-hidden", "false");
    setTimeout(() => input?.focus(), 50);
  });
  backdrop?.addEventListener("click", closeModal);
  closeBtn?.addEventListener("click", closeModal);
  clearBtn?.addEventListener("click", () => {
    if (qs("ironmindHelpAnswer")) qs("ironmindHelpAnswer").textContent = "";
    ironmindHelpHistory = [];
  });
  sendBtn?.addEventListener("click", () => sendIronmindHelpMessage().catch(() => {}));
  input?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      sendIronmindHelpMessage().catch(() => {});
    }
  });

  window.updateIronmindHelpFabContext?.();
}

function initTabs() {
  const tabSelect = qs("tabSelect");
  if (!tabSelect) return;
  tabSelect.addEventListener("change", () => switchTab(tabSelect.value));
  const urlTab = String(new URLSearchParams(window.location.search).get("tab") || "").trim();
  if (urlTab && tabSelect.querySelector(`option[value="${urlTab}"]`)) {
    tabSelect.value = urlTab;
  }
  if (!document.querySelector(".panel.show")) {
    switchTab(tabSelect.value || "dash");
  }
}

function initSectionCollapseToggles() {
  document.querySelectorAll("button.sectionToggleBtn[data-section-body][data-storage-key]").forEach((btn) => {
    const bodyId = btn.getAttribute("data-section-body");
    const key = btn.getAttribute("data-storage-key");
    if (!bodyId || !key) return;
    const body = document.getElementById(bodyId);
    if (!(body instanceof HTMLElement)) return;

    function applyHidden(hidden) {
      body.style.display = hidden ? "none" : "";
      btn.textContent = hidden ? "Show" : "Hide";
      btn.setAttribute("aria-expanded", hidden ? "false" : "true");
    }

    applyHidden(localStorage.getItem(key) === "1");

    btn.addEventListener("click", () => {
      const willHide = body.style.display !== "none";
      applyHidden(willHide);
      localStorage.setItem(key, willHide ? "1" : "0");
    });
  });
}

/* =========================
   UPLOADS
========================= */
