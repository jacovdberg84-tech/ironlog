// IRONLOG/web/app-nav.js — the one IronLog menu, shared by every desk page.
//
// The main screen (index.html), Maintenance, Work Orders, Breakdowns and the API documentation
// all show this menu, built from NAV_SECTIONS, so a screen is added or renamed in one place. It
// also holds which screens each role sees (roleTabs / effectiveTabs); app/core.js uses the same
// rules. The server still checks every request; the menu only hides what someone cannot use.
// Phone forms and shared terminals (pre-starts, stores terminal, technician portal) keep their
// own simple screens.
(function (global) {
  const NAV_SECTIONS = [
  {
    id: "home", title: "Home",
    items: [
      { label: "My Work", title: "What needs your attention today", tab: "mywork", icon: "<path d=\"M3 11l9-8 9 8\"></path><path d=\"M5 10v10h14V10\"></path><path d=\"M9 20v-6h6v6\"></path>" },
      { label: "Technician portal", title: "Your jobs, job timer, parts and shift report", href: "technician-terminal.html", nav: "tech-portal", key: "tech-portal", id: "navTechPortal", hidden: true, icon: "<path d=\"M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.5 2.5-2.4-.6-.6-2.4z\"></path>" },
      { label: "Dashboard", title: "Fleet availability, utilisation and alerts at a glance", tab: "dash", icon: "<rect x=\"3\" y=\"3\" width=\"7\" height=\"7\"></rect><rect x=\"14\" y=\"3\" width=\"7\" height=\"7\"></rect><rect x=\"14\" y=\"14\" width=\"7\" height=\"7\"></rect><rect x=\"3\" y=\"14\" width=\"7\" height=\"7\"></rect>" },
      { label: "Tasks", title: "Your to-do list and team tasks", tab: "tasks", id: "navTasks", icon: "<path d=\"M9 11l3 3L22 4\"></path><path d=\"M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11\"></path>" },
    ],
  },
  {
    id: "workshop", title: "Workshop",
    items: [
      { label: "Maintenance", title: "Services due, service history and workshop planning", href: "maintenance.html", nav: "maintenance", key: "maintenance", icon: "<path d=\"M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z\"></path>" },
      { label: "Work Orders", title: "Assign, repair and close workshop jobs", href: "workorders.html", nav: "maintenance", key: "workorders", icon: "<path d=\"M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z\"></path><polyline points=\"14 2 14 8 20 8\"></polyline><line x1=\"16\" y1=\"13\" x2=\"8\" y2=\"13\"></line><line x1=\"16\" y1=\"17\" x2=\"8\" y2=\"17\"></line>" },
      { label: "Breakdowns", title: "Log and follow up machine breakdowns", href: "breakdown-ops.html", nav: "maintenance", key: "breakdowns", icon: "<path d=\"M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z\"></path><line x1=\"12\" y1=\"9\" x2=\"12\" y2=\"13\"></line><line x1=\"12\" y1=\"17\" x2=\"12.01\" y2=\"17\"></line>" },
      { label: "Inspections", title: "Pre-start and equipment checklists", tab: "vehicle", icon: "<path d=\"M9 11l3 3L22 4\"></path><path d=\"M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11\"></path>" },
      { label: "Workshop Library", title: "Manuals, parts books and workshop documents", tab: "workshop", icon: "<path d=\"M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z\"></path>" },
    ],
  },
  {
    id: "fleet", title: "Fleet",
    items: [
      { label: "Assets", title: "Machine and vehicle register", tab: "assets", icon: "<path d=\"M6 2L3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z\"></path><line x1=\"3\" y1=\"6\" x2=\"21\" y2=\"6\"></line><path d=\"M16 10a4 4 0 0 1-8 0\"></path>" },
      { label: "Daily Log", title: "Capture daily hours and downtime", tab: "daily", icon: "<path d=\"M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z\"></path><polyline points=\"14 2 14 8 20 8\"></polyline><line x1=\"16\" y1=\"13\" x2=\"8\" y2=\"13\"></line><line x1=\"16\" y1=\"17\" x2=\"8\" y2=\"17\"></line>" },
      { label: "Fuel", title: "Fuel issues and consumption", tab: "fuel", icon: "<path d=\"M3 22h18\"></path><path d=\"M5 22V9a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v13\"></path><path d=\"M11 9V6a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v3\"></path><path d=\"M17 13h2a2 2 0 0 1 2 2v2a2 2 0 0 1-2 2h-2\"></path>" },
      { label: "Oils", title: "Oils issued to machines, and oil stock", tab: "lube", icon: "<path d=\"M10 2v8L8 12l-2-2V2\"></path><path d=\"M14 2v8l2 2 2-2V2\"></path><rect x=\"7\" y=\"12\" width=\"10\" height=\"10\"></rect>" },
      { label: "Telematics", title: "Telematics units and fault codes", tab: "telematics", icon: "<rect x=\"2\" y=\"7\" width=\"20\" height=\"14\" rx=\"2\"></rect><path d=\"M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2\"></path><circle cx=\"12\" cy=\"14\" r=\"2\"></circle><path d=\"M12 12v-2\"></path>" },
      { label: "Fleet Track", title: "Live vehicle positions", tab: "cartrack", icon: "<path d=\"M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z\"></path><circle cx=\"12\" cy=\"10\" r=\"3\"></circle>" },
    ],
  },
  {
    id: "stores", title: "Stores",
    items: [
      { label: "Stock Control", title: "Stock on hand, receiving and issuing parts", tab: "stock", icon: "<path d=\"M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z\"></path><polyline points=\"3.27 6.96 12 12.01 20.73 6.96\"></polyline><line x1=\"12\" y1=\"22.08\" x2=\"12\" y2=\"12\"></line>" },
      { label: "Parts Orders", title: "Parts on order and off-site repairs", tab: "parts-tracking", icon: "<path d=\"M16.5 9.4l-9-5.19\"></path><path d=\"M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z\"></path><polyline points=\"3.27 6.96 12 12.01 20.73 6.96\"></polyline><line x1=\"12\" y1=\"22.08\" x2=\"12\" y2=\"12\"></line>" },
      { label: "Purchasing", title: "Requisitions, approvals and supplier orders", tab: "procurement", icon: "<line x1=\"12\" y1=\"1\" x2=\"12\" y2=\"23\"></line><path d=\"M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6\"></path>" },
    ],
  },
  {
    id: "site", title: "Site & Compliance",
    items: [
      { label: "Site Operations", title: "Daily site production and operations log", tab: "operations", icon: "<path d=\"M2 20h20\"></path><path d=\"M5 20V10l7-7 7 7v10\"></path><path d=\"M9 20v-6h6v6\"></path>" },
      { label: "Dispatch", title: "Machine and vehicle dispatch", tab: "dispatch", icon: "<rect x=\"1\" y=\"3\" width=\"15\" height=\"13\"></rect><polygon points=\"16 8 20 8 23 11 23 16 16 16 16 8\"></polygon><circle cx=\"5.5\" cy=\"18.5\" r=\"2.5\"></circle><circle cx=\"18.5\" cy=\"18.5\" r=\"2.5\"></circle>" },
      { label: "Legal Docs", title: "Licences, certificates and expiry dates", tab: "legal", icon: "<path d=\"M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z\"></path><polyline points=\"14 2 14 8 20 8\"></polyline><line x1=\"16\" y1=\"13\" x2=\"8\" y2=\"13\"></line><line x1=\"16\" y1=\"17\" x2=\"8\" y2=\"17\"></line><polyline points=\"10 9 9 9 8 9\"></polyline>" },
      { label: "Approvals", title: "Items waiting for your approval", tab: "approvals", icon: "<polyline points=\"9 11 12 14 22 4\"></polyline><path d=\"M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11\"></path>" },
      { label: "Data Quality", title: "Find missing or inconsistent data", tab: "quality", icon: "<path d=\"M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z\"></path>" },
    ],
  },
  {
    id: "reports", title: "Reports",
    items: [
      { label: "Reports & Exports", title: "PDF and Excel reports", tab: "reports", icon: "<path d=\"M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z\"></path><polyline points=\"14 2 14 8 20 8\"></polyline><line x1=\"16\" y1=\"13\" x2=\"8\" y2=\"13\"></line><line x1=\"16\" y1=\"17\" x2=\"8\" y2=\"17\"></line>" },
      { label: "Finance", title: "Costs, budgets and month-end", tab: "finance", icon: "<line x1=\"12\" y1=\"1\" x2=\"12\" y2=\"23\"></line><path d=\"M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6\"></path>" },
      { label: "Executive Overview", title: "Executive command centre and board pack", tab: "exec", icon: "<polyline points=\"3 12 9 6 13 10 21 3\"></polyline><polyline points=\"16 3 21 3 21 8\"></polyline>" },
      { label: "Enterprise", title: "Companies, sites, currencies and integrations", tab: "enterprise", icon: "<rect x=\"3\" y=\"3\" width=\"7\" height=\"7\"></rect><rect x=\"14\" y=\"3\" width=\"7\" height=\"7\"></rect><rect x=\"3\" y=\"14\" width=\"7\" height=\"7\"></rect><rect x=\"14\" y=\"14\" width=\"7\" height=\"7\"></rect>" },
    ],
  },
  {
    id: "help", title: "Help",
    items: [
      { label: "Ask Borris", title: "Ask the AI assistant about your fleet", tab: "ironmind", icon: "<path d=\"M12 2a4 4 0 0 1 4 4c0 1.5-.8 2.8-2 3.5V11h2a3 3 0 0 1 3 3v1H5v-1a3 3 0 0 1 3-3h2V9.5A4 4 0 0 1 12 2z\"></path><path d=\"M9 14l2 2 4-4\"></path>" },
      { label: "AI Documents", title: "Draft documents with AI", tab: "docs", icon: "<path d=\"M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z\"></path>" },
    ],
  },
  {
    id: "admin", title: "Admin",
    items: [
      { label: "Developer", title: "Coding agent for IronLog changes", tab: "developer", icon: "<polyline points=\"16 18 22 12 16 6\"></polyline><polyline points=\"8 6 2 12 8 18\"></polyline>" },
      { label: "Uploads & Go-live", title: "Go-live import from Excel and CSV uploads", tab: "uploads", icon: "<path d=\"M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4\"></path><polyline points=\"17 8 12 3 7 8\"></polyline><line x1=\"12\" y1=\"3\" x2=\"12\" y2=\"15\"></line>" },
      { label: "Audit Trail", title: "Who changed what, and when", tab: "audit", icon: "<circle cx=\"12\" cy=\"12\" r=\"10\"></circle><polyline points=\"12 6 12 12 16 14\"></polyline>" },
      { label: "Users & Access", title: "Manage users, roles and permissions", tab: "admin", id: "navAdmin", icon: "<path d=\"M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2\"></path><circle cx=\"9\" cy=\"7\" r=\"4\"></circle><path d=\"M23 21v-2a4 4 0 0 0-3-3.87\"></path><path d=\"M16 3.13a4 4 0 0 1 0 7.75\"></path>" },
      { label: "API Documentation", title: "How other systems connect to IronLog", href: "api-docs.html", key: "api-docs", adminOnly: true, icon: "<polyline points=\"16 18 22 12 16 6\"></polyline><polyline points=\"8 6 2 12 8 18\"></polyline><line x1=\"14\" y1=\"4\" x2=\"10\" y2=\"20\"></line>" },
    ],
  },
  ];

  /** Screens in use on the live site (all opened October 2026; a screen leaves only on purpose). */
  const PRODUCTION_SITE_TABS = [
    "mywork", "tasks", "dash", "daily", "maintenance", "assets", "telematics", "cartrack", "workshop", "fuel", "lube", "vehicle",
    "stock", "parts-tracking", "reports", "finance", "procurement", "approvals", "operations", "dispatch", "quality", "legal",
    "audit", "docs", "uploads", "enterprise", "exec", "ironmind", "developer",
  ];

  const ALL_TABS = [
    "dash", "daily", "assets", "telematics", "cartrack", "workshop", "maintenance", "fuel", "lube", "stock", "parts-tracking", "legal",
    "uploads", "reports", "finance", "enterprise", "exec", "Breakdowns", "approvals", "procurement", "operations", "dispatch", "quality",
    "audit", "ironmind", "docs", "vehicle", "tasks", "admin",
  ];

  const ROLE_TABS = {
    plant_clerk: ["dash", "daily", "assets", "fuel", "lube", "vehicle", "reports", "docs", "tasks"],
    workshop_admin: ["dash", "daily", "assets", "telematics", "workshop", "maintenance", "stock", "parts-tracking", "reports", "Breakdowns", "approvals", "procurement", "ironmind", "docs", "vehicle", "tasks"],
    operator: ["dash", "daily", "workshop", "fuel", "lube", "legal", "operations", "ironmind", "docs", "vehicle", "tasks", "telematics", "cartrack"],
    artisan: ["workshop", "maintenance", "Breakdowns", "vehicle", "tasks"],
    stores: ["dash", "stock", "parts-tracking", "workshop", "reports", "procurement", "tasks"],
    storeman: ["dash", "stock", "parts-tracking", "workshop", "reports", "procurement", "tasks"],
    procurement: ["dash", "stock", "parts-tracking", "workshop", "reports", "finance", "procurement", "operations", "quality", "docs", "tasks"],
    plant_manager: ["dash", "daily", "assets", "telematics", "cartrack", "workshop", "maintenance", "fuel", "lube", "stock", "parts-tracking", "reports", "finance", "procurement", "operations", "dispatch", "quality", "audit", "ironmind", "docs", "vehicle", "tasks"],
    site_manager: ["dash", "daily", "assets", "telematics", "cartrack", "workshop", "maintenance", "fuel", "lube", "stock", "parts-tracking", "reports", "finance", "procurement", "operations", "dispatch", "quality", "audit", "docs", "tasks"],
    executive: ["dash", "workshop", "reports", "finance", "operations", "quality", "audit", "docs", "tasks"],
    supervisor: ["dash", "daily", "assets", "telematics", "cartrack", "workshop", "maintenance", "fuel", "lube", "stock", "parts-tracking", "legal", "uploads", "reports", "finance", "enterprise", "exec", "Breakdowns", "approvals", "procurement", "operations", "dispatch", "quality", "audit", "ironmind", "docs", "vehicle", "tasks"],
  };

  /** Roles that use the technician portal. */
  const TECH_PORTAL_ROLES = ["artisan", "supervisor", "workshop_admin", "admin", "plant_manager", "site_manager"];

  function roleTabs(role) {
    return [...(ROLE_TABS[String(role || "").toLowerCase()] || ALL_TABS)];
  }

  /**
   * The screens a person sees: their roles' screens (or a per-user list an admin saved), limited to
   * the screens in use, without add-on modules outside the licence. My Work comes first.
   */
  function effectiveTabs({ roles = [], override = null, unlicensed = new Set() } = {}) {
    let list = Array.isArray(override) && override.length ? [...override] : Array.from(new Set(roles.flatMap((r) => roleTabs(r))));
    const production = new Set(PRODUCTION_SITE_TABS);
    list = list.filter((t) => production.has(t));
    // Borris stays reachable even for saved per-user lists created before it was restored.
    if (!list.includes("ironmind")) list.push("ironmind");
    // Tasks back the My Work list; per-user section lists cannot assign it, so always allow it.
    if (!list.includes("tasks")) list.push("tasks");
    if (roles.some((r) => ["admin", "supervisor", "executive"].includes(r))) {
      if (!list.includes("enterprise")) list.push("enterprise");
      if (!list.includes("exec")) list.push("exec");
    }
    // Saved per-user lists from before Finance existed.
    if (roles.some((r) => ["admin", "supervisor", "procurement", "plant_manager", "site_manager", "executive"].includes(r)) && !list.includes("finance")) list.push("finance");
    if (roles.includes("admin") && !list.includes("admin")) list.push("admin");
    if (roles.includes("admin") && !list.includes("developer")) list.push("developer");
    if (unlicensed && unlicensed.size) list = list.filter((t) => !unlicensed.has(t));
    return ["mywork", ...list.filter((t) => t !== "mywork")];
  }

  /** Whether one menu item shows for these screens and roles. */
  function itemVisible(it, allowed, roles) {
    if (it.adminOnly || it.tab === "admin" || it.tab === "developer") return roles.includes("admin");
    if (it.key === "tech-portal") return roles.some((r) => TECH_PORTAL_ROLES.includes(r));
    if (it.nav === "maintenance") return allowed.has("maintenance");
    if (it.tab) return allowed.has(it.tab);
    return true;
  }

  const esc = (s) => String(s ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

  /**
   * The menu's HTML. mode "tabs" (the main screen): screens switch in place (href "#", data-tab).
   * mode "links" (other pages): every screen is a link back to the main screen (index.html?tab=…).
   */
  function sectionsHtml({ mode = "links", active = "" } = {}) {
    return NAV_SECTIONS.map((sec) => `
      <div class="nav-section" data-nav-section="${esc(sec.id)}">
        <div class="nav-section-title">${esc(sec.title)}</div>
        ${sec.items.map((it) => {
          const href = it.href || (mode === "tabs" ? "#" : `index.html?tab=${encodeURIComponent(it.tab)}`);
          const attrs = [
            `href="${esc(href)}"`,
            `class="nav-item${active && (it.key === active || it.tab === active) ? " active" : ""}"`,
            mode === "tabs" && it.tab ? `data-tab="${esc(it.tab)}"` : "",
            it.nav ? `data-nav="${esc(it.nav)}"` : "",
            it.key ? `data-nav-key="${esc(it.key)}"` : "",
            it.tab ? `data-nav-tab="${esc(it.tab)}"` : "",
            it.id ? `id="${esc(it.id)}"` : "",
            `title="${esc(it.title)}"`,
            it.hidden ? `style="display:none"` : "",
          ].filter(Boolean).join(" ");
          return `<a ${attrs}><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">${it.icon}</svg><span>${esc(it.label)}</span></a>`;
        }).join("")}
      </div>`).join("");
  }

  /** Hides items this person cannot use, and sections left empty. */
  function applyVisibility(root, allowed, roles) {
    const byKey = new Map();
    NAV_SECTIONS.forEach((sec) => sec.items.forEach((it) => byKey.set(it.key || it.tab, it)));
    root.querySelectorAll(".nav-item[data-nav-key], .nav-item[data-nav-tab]").forEach((el) => {
      const it = byKey.get(el.dataset.navKey || el.dataset.navTab);
      if (it) el.style.display = itemVisible(it, allowed, roles) ? "" : "none";
    });
    root.querySelectorAll(".nav-section[data-nav-section]").forEach((sec) => {
      const shown = Array.from(sec.querySelectorAll(":scope > .nav-item")).some((el) => el.style.display !== "none");
      sec.style.display = shown ? "" : "none";
    });
  }

  // ---- Pages other than the main screen.

  const session = () => global.IronlogSession || null;

  function sessionRoles() {
    if (session()) return session().getSessionRoles();
    try {
      const v = JSON.parse(localStorage.getItem("ironlog_session_roles") || "[]");
      if (Array.isArray(v) && v.length) return v.map((r) => String(r).toLowerCase());
    } catch {}
    return [String(localStorage.getItem("ironlog_session_role") || "admin").toLowerCase()];
  }

  function savedOverride() {
    try {
      const v = JSON.parse(localStorage.getItem("ironlog_allowed_tabs") || "null");
      return Array.isArray(v) && v.length ? v : null;
    } catch {
      return null;
    }
  }

  function requestHeaders() {
    if (session()) return session().authHeaders();
    const tok = String(localStorage.getItem("ironlog_auth_token") || "").trim();
    return {
      ...(tok ? { Authorization: `Bearer ${tok}` } : {}),
      "x-site-code": String(localStorage.getItem("ironlog_session_site") || "main").toLowerCase(),
    };
  }

  /** Applies the person's screens, then again once the licence is known. */
  function refreshVisibility(root) {
    const roles = sessionRoles();
    const show = (unlicensed) => applyVisibility(root, new Set(effectiveTabs({ roles, override: savedOverride(), unlicensed })), roles);
    show(new Set());
    fetch("/api/licence/state", { headers: requestHeaders() })
      .then((r) => (r.ok ? r.json() : null))
      .then((s) => {
        if (!s?.module_tabs) return;
        const on = new Set(s.modules || []);
        show(new Set(Object.entries(s.module_tabs).filter(([m]) => !on.has(m)).flatMap(([, tabs]) => tabs)));
      })
      .catch(() => {});
  }

  const MENU_ICON = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="3" y1="6" x2="21" y2="6"></line><line x1="3" y1="12" x2="21" y2="12"></line><line x1="3" y1="18" x2="21" y2="18"></line></svg>';
  const MOON_ICON = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path></svg>';

  /**
   * Gives a page the IronLog top bar and menu: the page's own content moves into the main area.
   * active = the menu key of this page (workorders, breakdowns, api-docs); subtitle = its name.
   */
  function mountPage({ active = "", subtitle = "" } = {}) {
    const body = document.body;
    if (body.dataset.ironlogNav === "1") return;
    body.dataset.ironlogNav = "1";
    const content = Array.from(body.childNodes).filter((n) => !(n.nodeType === 1 && n.tagName === "SCRIPT"));
    const main = document.createElement("div");
    main.className = "main-content";
    main.id = "mainContent";
    content.forEach((n) => main.appendChild(n));
    const overlay = document.createElement("div");
    overlay.className = "sidebar-overlay";
    overlay.id = "sidebarOverlay";
    const top = document.createElement("header");
    top.className = "topbar";
    top.innerHTML = `
      <div class="topbar-left">
        <button class="sidebar-toggle" id="sidebarToggle" type="button" aria-label="Show or hide the menu">${MENU_ICON}</button>
        <a class="brand" href="index.html" title="IronLog home">
          <div class="logo"><img src="./ironlog-logo.png" alt="IRONLOG" /></div>
          <div><div class="title">IRONLOG</div><div class="subtitle">${esc(subtitle)}</div></div>
        </a>
      </div>
      <div class="topbar-right">
        <button id="ironlogNavTheme" class="icon-btn" type="button" title="Light or dark screen">${MOON_ICON}</button>
      </div>`;
    const nav = document.createElement("nav");
    nav.className = "sidebar";
    nav.id = "sidebar";
    nav.innerHTML = `<div class="sidebar-content">${sectionsHtml({ mode: "links", active })}</div>`;
    body.classList.add("has-sidebar", "ironlog-nav-page");
    body.prepend(overlay, top, nav, main);
    if (localStorage.getItem("ironlog_sidebar_collapsed") === "1" && window.innerWidth > 1024) nav.classList.add("collapsed");
    const close = () => { nav.classList.remove("mobile-open"); overlay.classList.remove("active"); };
    top.querySelector("#sidebarToggle").addEventListener("click", () => {
      if (window.innerWidth <= 1024) {
        const open = !nav.classList.contains("mobile-open");
        nav.classList.toggle("mobile-open", open);
        overlay.classList.toggle("active", open);
        return;
      }
      const collapsed = nav.classList.toggle("collapsed");
      try { localStorage.setItem("ironlog_sidebar_collapsed", collapsed ? "1" : "0"); } catch {}
    });
    overlay.addEventListener("click", close);
    top.querySelector("#ironlogNavTheme").addEventListener("click", () => {
      const dark = document.documentElement.getAttribute("data-theme") !== "dark";
      if (dark) document.documentElement.setAttribute("data-theme", "dark");
      else document.documentElement.removeAttribute("data-theme");
      try { localStorage.setItem("ironlog-theme", dark ? "dark" : "light"); } catch {}
    });
    refreshVisibility(nav);
  }

  /**
   * The Maintenance page keeps its own top bar; the shared menu goes into its sidebar and the
   * page's own sections sit under Workshop → Maintenance.
   */
  function mountIntoSidebar(sidebar, { active = "maintenance" } = {}) {
    const holder = sidebar.querySelector(".sidebar-content") || sidebar;
    if (holder.dataset.ironlogNav === "1") return;
    holder.dataset.ironlogNav = "1";
    const own = Array.from(holder.children);
    holder.innerHTML = sectionsHtml({ mode: "links", active });
    const anchor = holder.querySelector(`.nav-item[data-nav-key="${active}"]`);
    // The page's own sections carry the "active" mark; the shared item just says "you are here".
    anchor?.classList.replace("active", "is-page");
    if (anchor && own.length) {
      const sub = document.createElement("div");
      sub.className = "nav-subsections";
      own.forEach((el) => sub.appendChild(el));
      anchor.after(sub);
    }
    refreshVisibility(holder);
  }

  // The main screen marks where its menu goes; it is filled in as soon as this file loads,
  // before app/core.js wires up the clicks.
  if (typeof document !== "undefined") {
    document.querySelectorAll('[data-ironlog-nav="tabs"]').forEach((el) => {
      el.insertAdjacentHTML("beforebegin", sectionsHtml({ mode: "tabs" }));
      el.remove();
    });
  }

  global.IronlogNav = { NAV_SECTIONS, PRODUCTION_SITE_TABS, roleTabs, effectiveTabs, itemVisible, sectionsHtml, applyVisibility, mountPage, mountIntoSidebar };
})(typeof window !== "undefined" ? window : globalThis);
