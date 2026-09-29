// IRONLOG/web/maintenance/views.js — Maintenance hub modes and top-level views.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

function maintenanceHubModeFor(view, section) {
  if (view === "insights") return "insights";
  if (view === "service-history") return "service-history";
  return "plans";
}

function applyMaintenanceHubMode(mode) {
  const planSection = document.querySelector("section.panel.page-section");
  const showCard = (id, visible) => {
    const el = document.getElementById(id);
    if (el) el.style.display = visible ? "" : "none";
  };
  const modes = {
    plans: {
      planSection: true,
      plans: true,
      serviceRecords: false,
      serviceSchedule: false,
      due: false,
      insights: false,
    },
    "service-history": {
      planSection: false,
      plans: false,
      serviceRecords: true,
      serviceSchedule: true,
      due: false,
      insights: false,
    },
    insights: {
      planSection: false,
      plans: false,
      serviceRecords: false,
      serviceSchedule: false,
      due: false,
      insights: true,
    },
  };
  const cfg = modes[mode] || modes.plans;

  if (planSection) planSection.style.display = cfg.planSection ? "block" : "none";
  showCard("maintenancePlansCard", cfg.plans);
  showCard("serviceRecordsCard", cfg.serviceRecords);
  showCard("serviceScheduleCard", cfg.serviceSchedule);
  showCard("dueServicesCard", cfg.due);
  showCard("maintenanceInsightsCard", cfg.insights);
  showCard("inspectproStatusCard", false);
  showCard("rsgProfilesCard", false);
}

function setTopView(view, section = "") {
  const main = document.getElementById("mainMaintenanceCards");
  const mi = document.getElementById("managerInspectionsSection");
  const ai = document.getElementById("artisanInspectionsSection");
  const wf = document.getElementById("weeklyForumSection");
  const tyre = document.getElementById("tyreInspectionsSection");
  const uc = document.getElementById("undercarriageInspectionsSection");
  const wi = document.getElementById("weeklyInspectionSection");
  const kpi = document.getElementById("assetKpiSection");
  const rel = document.getElementById("reliabilitySection");
  const hist = document.getElementById("histogramSection");
  const sync = document.getElementById("syncAdminSection");
  const pto = document.getElementById("partsToOrderSection");
  const mc = document.getElementById("mechanicsCostSection");
  const st = document.getElementById("serviceTemplatesSection");
  const planSection = document.querySelector("section.panel.page-section");

  if (section) currentMaintenanceSection = section;

  // Hide all sections first (only mainMaintenanceCards — sibling sections live in the same page-section)
  if (main) main.style.display = "none";
  if (mi) mi.style.display = "none";
  if (ai) ai.style.display = "none";
  if (wf) wf.style.display = "none";
  if (tyre) tyre.style.display = "none";
  if (uc) uc.style.display = "none";
  if (wi) wi.style.display = "none";
  if (kpi) kpi.style.display = "none";
  if (rel) rel.style.display = "none";
  if (hist) hist.style.display = "none";
  if (sync) sync.style.display = "none";
  if (pto) pto.style.display = "none";
  if (mc) mc.style.display = "none";
  if (st) st.style.display = "none";
  if (planSection) planSection.style.display = "none";

  // Show the selected section
  switch (view) {
    case "main":
    case "service-history":
    case "insights":
      if (main) main.style.display = "block";
      applyMaintenanceHubMode(maintenanceHubModeFor(view, section || currentMaintenanceSection));
      break;
    case "mi":
      if (mi) mi.style.display = "block";
      break;
    case "ai":
      if (ai) ai.style.display = "block";
      break;
    case "wf":
      if (wf) wf.style.display = "block";
      break;
    case "tyre":
      if (tyre) tyre.style.display = "block";
      break;
    case "uc":
      if (uc) uc.style.display = "block";
      break;
    case "wi":
      if (wi) wi.style.display = "block";
      break;
    case "kpi":
      if (kpi) kpi.style.display = "block";
      break;
    case "rel":
      if (rel) rel.style.display = "block";
      break;
    case "hist":
      if (hist) hist.style.display = "block";
      break;
    case "sync":
      if (sync) sync.style.display = "block";
      break;
    case "pto":
      if (pto) pto.style.display = "block";
      break;
    case "mc":
      if (mc) mc.style.display = "block";
      break;
    case "templates":
      if (st) st.style.display = "block";
      break;
  }

  // Update active state in sidebar
  const sidebar = document.getElementById("sidebar");
  if (sidebar) {
    sidebar.querySelectorAll(".nav-item").forEach((item) => {
      const navSection = item.dataset.section;
      let isActive = false;
      switch (view) {
        case "main":
          isActive = navSection === "maintenance";
          break;
        case "service-history":
          isActive = navSection === "service-history";
          break;
        case "insights":
          isActive = navSection === "maintenance-insights";
          break;
        case "mi":
          isActive = navSection === "manager-inspections";
          break;
        case "ai":
          isActive = navSection === "artisan-inspections";
          break;
        case "wf":
          isActive = navSection === "weekly-forum";
          break;
        case "tyre":
          isActive = navSection === "tyre-inspections";
          break;
        case "uc":
          isActive = navSection === "undercarriage-inspections";
          break;
        case "wi":
          isActive = navSection === "weekly-inspections";
          break;
        case "kpi":
          isActive = navSection === "asset-kpi";
          break;
        case "rel":
          isActive = navSection === "reliability";
          break;
        case "hist":
          isActive = navSection === "histogram";
          break;
        case "sync":
          isActive = navSection === "sync-admin";
          break;
        case "pto":
          isActive = navSection === "parts-to-order";
          break;
        case "mc":
          isActive = navSection === "mechanics-cost";
          break;
        case "templates":
          isActive = navSection === "service-templates";
          break;
      }
      item.classList.toggle("active", isActive);
    });
  }

  // Ensure section data refreshes when opened via sidebar navigation.
  refreshTopViewData(view);
}
