// IRONLOG/web/maintenance/core.js — Config, sidebar and view switching, session labels, fetch wrapper, shared formatting helpers.
// Part of maintenance.html; the page loads these files in order and they share one global scope.
const API = "/api";
const { getSessionUser, getSessionRoles, getSessionSite, authHeaders } = window.IronlogSession;
const selectedDuePlanIds = new Set();
let lastSyncPull = { last_id: 0, events: [] };
const MAINT_DUE_THRESHOLD_KEY = "ironlog_maintenance_due_threshold_hours";
const INSIGHTS_THRESHOLDS_KEY = "ironlog_maintenance_insights_thresholds";
let reportBuilderMeta = [];
let reportBuilderTemplates = [];
let reportSubscriptionEditId = 0;
let reportSubscriptionsCache = [];
const TYRE_POSITIONS = [
  { key: "front_left", label: "Front Left", surveyCode: "LF", surveyOrder: 1 },
  { key: "front_right", label: "Front Right", surveyCode: "RF", surveyOrder: 2 },
  { key: "rear_right_inner", label: "Rear Right Inner", surveyCode: "RM", surveyOrder: 3 },
  { key: "rear_right_outer", label: "Rear Right Outer", surveyCode: "RR", surveyOrder: 4 },
  { key: "rear_left_outer", label: "Rear Left Outer", surveyCode: "LR", surveyOrder: 5 },
  { key: "rear_left_inner", label: "Rear Left Inner", surveyCode: "LM", surveyOrder: 6 },
];

function initMaintSidebar() {
  const sidebar = document.getElementById("sidebar");
  const toggle = document.getElementById("sidebarToggle");
  const overlay = document.getElementById("sidebarOverlay");
  
  if (!sidebar) return;
  
  // Toggle handler
  toggle?.addEventListener("click", () => {
    sidebar.classList.toggle("collapsed");
    overlay?.classList.toggle("active");
  });
  
  // Close sidebar on mobile when clicking overlay
  overlay?.addEventListener("click", () => {
    sidebar.classList.remove("collapsed");
    overlay.classList.remove("active");
  });
  
  // Navigation item clicks
  sidebar.querySelectorAll(".nav-item").forEach((item) => {
    item.addEventListener("click", (e) => {
      e.preventDefault();
      const section = item.dataset.section;
      if (!section) return;
      
      scrollToSection(section);
      
      // Close mobile sidebar
      if (window.innerWidth <= 1024) {
        sidebar.classList.add("collapsed");
        overlay?.classList.remove("active");
      }
    });
  });
  
  // Handle window resize
  window.addEventListener("resize", () => {
    if (window.innerWidth > 1024) {
      sidebar.classList.remove("collapsed");
      overlay?.classList.remove("active");
    }
  });
}

function initMaintCollapsibles() {
  document.querySelectorAll(".maintenance-card.collapsible").forEach((card) => {
    const header = card.querySelector(".maintenance-card-header");
    const toggle = card.querySelector(".maintenance-card-toggle");
    
    if (!header || !toggle) return;
    
    const toggleCollapse = () => {
      const isCollapsed = card.classList.toggle("collapsed");
      card.dataset.collapsed = isCollapsed;
      const cardTitle = card.querySelector("h3")?.textContent?.trim() || "";
      localStorage.setItem(`maint_card_${cardTitle}`, isCollapsed ? "1" : "0");
    };
    
    header.addEventListener("click", (e) => {
      e.stopPropagation();
      toggleCollapse();
    });
    
    toggle.addEventListener("click", (e) => {
      e.stopPropagation();
      toggleCollapse();
    });
    
    // Restore saved state
    const cardTitle = card.querySelector("h3")?.textContent?.trim() || "";
    const savedState = localStorage.getItem(`maint_card_${cardTitle}`);
    if (savedState === "1") {
      card.classList.add("collapsed");
      card.dataset.collapsed = "true";
    }
  });
}

function viewForSection(section) {
  const viewMap = {
    "maintenance": "main",
    "service-templates": "templates",
    "service-history": "service-history",
    "maintenance-insights": "insights",
    "mechanics-cost": "mc",
    "manager-inspections": "mi",
    "artisan-inspections": "ai",
    "weekly-forum": "wf",
    "tyre-inspections": "tyre",
    "undercarriage-inspections": "uc",
    "weekly-inspections": "wi",
    "asset-kpi": "kpi",
    "reliability": "rel",
    "histogram": "hist",
    "sync-admin": "sync",
    "parts-to-order": "pto",
  };
  return viewMap[String(section || "").trim()] || "";
}

let currentMaintenanceSection = "maintenance";

const MAINT_EXTERNAL_SECTIONS = {
  "work-orders": "workorders.html",
  "breakdowns": "breakdown-ops.html",
};

function isMaintenanceSection(section) {
  const s = String(section || "").trim();
  return Boolean(viewForSection(s) || MAINT_EXTERNAL_SECTIONS[s]);
}

function scrollToSection(section) {
  const s = String(section || "").trim();
  const externalHref = MAINT_EXTERNAL_SECTIONS[s];
  if (externalHref) {
    location.href = externalHref;
    return;
  }
  const targetView = viewForSection(section);
  if (targetView) {
    if (section) currentMaintenanceSection = section;
    setTopView(targetView, section);
    
    setTimeout(() => {
      let targetEl = null;
      
      if (targetView === "main") {
        if (section === "service-history") {
          targetEl = document.getElementById("serviceRecordsCard");
        }
        if (!targetEl) {
          const plansSection = document.querySelector(".panel h3");
          targetEl = plansSection || document.getElementById("mainContent");
        }
      } else {
        const sectionMap = {
          "mi": "managerInspectionsSection",
          "ai": "artisanInspectionsSection",
          "wf": "weeklyForumSection",
          "tyre": "tyreInspectionsSection",
          "uc": "undercarriageInspectionsSection",
          "wi": "weeklyInspectionSection",
          "kpi": "assetKpiSection",
          "rel": "reliabilitySection",
          "hist": "histogramSection",
          "sync": "syncAdminSection",
          "insights": "maintenanceInsightsCard",
          "mc": "mechanicsCostSection",
          "templates": "serviceTemplatesSection",
          "pto": "partsToOrderSection",
        };
        targetEl = document.getElementById(sectionMap[targetView]);
      }
      
      if (targetEl) {
        targetEl.scrollIntoView({ behavior: "smooth", block: "start" });
      } else {
        const mainContent = document.getElementById("mainContent");
        if (mainContent) {
          mainContent.scrollIntoView({ behavior: "smooth", block: "start" });
        }
      }
    }, 50);
  }
}

function refreshTopViewData(view) {
  switch (view) {
    case "mi":
      loadManagerInspections().catch(() => {});
      break;
    case "ai":
      loadArtisanInspections().catch(() => {});
      break;
    case "wf":
      loadWeeklyForumSummary().catch(() => {});
      loadWeeklyForumReviews().catch(() => {});
      loadWeeklyForumActions().catch(() => {});
      loadWeeklyForumInputs().catch(() => {});
      break;
    case "tyre":
      loadTyreInspections();
      loadTyreLifecycle().catch(() => {});
      break;
    case "uc":
      if (typeof window.initUndercarriage === "function") window.initUndercarriage();
      break;
    case "wi":
      loadWeeklyInspectionCalendar().catch(() => {});
      break;
    case "kpi":
      loadAssetKpiWeekly().catch(() => {});
      break;
    case "rel":
      loadReliabilityAssets().catch(() => {});
      loadReliabilityMetrics().catch(() => {});
      break;
    case "insights":
      loadMaintenanceInsights().catch(() => {});
      loadReportSubscriptions().catch(() => {});
      break;
    case "service-history":
      loadBackfillHistory().catch(() => {});
      loadHistory().catch(() => {});
      break;
    case "sync":
      syncLoadState().catch(() => {});
      break;
    case "pto":
      initPartsToOrderSection().catch(() => {});
      break;
    case "mc":
      initMechanicsCostSection().catch(() => {});
      break;
    case "templates":
      initServiceTemplateSection().catch(() => {});
      break;
    default:
      break;
  }
}

function roleDisplayName(role) {
  const labels = {
    admin: "Admin",
    plant_manager: "Plant Manager",
    workshop_admin: "Workshop Admin",
    storeman: "Stores",
    stores: "Stores",
    plant_clerk: "Plant Clerk",
  };
  const key = String(role || "").trim().toLowerCase();
  return labels[key] || key.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase()) || "User";
}

function renderSessionRoleLabel() {
  const label = document.getElementById("sessionRolesBadge");
  if (!label) return;
  const roles = getSessionRoles().map(roleDisplayName);
  label.className = "session-role";
  label.textContent = roles.join(" · ");
  label.title = `Active session ${roles.length === 1 ? "role" : "roles"}: ${roles.join(", ")}`;
}

const __nativeFetch = window.fetch.bind(window);
window.fetch = (input, init = {}) => {
  const reqUrl = typeof input === "string" ? input : String(input?.url || "");
  const sameApi = reqUrl.startsWith("/api/") || reqUrl.startsWith(`${API}/`);
  if (!sameApi) return __nativeFetch(input, init);
  const headers = new Headers(init?.headers || {});
  Object.entries(authHeaders()).forEach(([k, v]) => {
    if (v != null && v !== "") headers.set(k, String(v));
  });
  return __nativeFetch(input, { ...init, headers });
};


function esc(v) {
  return String(v ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function normImgSrc(v) {
  const s = String(v || "").trim();
  if (!s) return "";
  if (s.startsWith("http://") || s.startsWith("https://") || s.startsWith("/")) return s;
  return `/${s.replace(/\\/g, "/")}`;
}


function fmtMoney(v) {
  const n = Number(v || 0);
  return Number.isFinite(n) ? n.toFixed(2) : "0.00";
}

function fmtPct(v) {
  if (v == null || v === "" || !Number.isFinite(Number(v))) return "—";
  return `${Number(v).toFixed(1)}%`;
}


function escAttr(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
