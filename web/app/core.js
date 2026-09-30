// IRONLOG/web/app/core.js — Config, session, auth, fetch helpers, navigation and role visibility, login.
// Part of the main app; index.html loads these files in order and they share one global scope.
const API = /^https?:$/.test(window.location?.protocol || "")
  ? window.location.origin
  : "http://localhost:3001";

// Safe getElementById
const qs = (id) => document.getElementById(id) || null;

const ROLE_KEY = "ironlog_session_role";
const ROLES_KEY = "ironlog_session_roles";
const USER_KEY = "ironlog_session_user";
const SITE_KEY = "ironlog_session_site";
const TOKEN_KEY = "ironlog_auth_token";
const TABS_OVERRIDE_KEY = "ironlog_allowed_tabs";
const SLA_OPEN_SAME_TAB_KEY = "ironlog_sla_open_same_tab";
const LOC_DEFAULT_PREFIX = "ironlog_default_location_";
const MAINT_CHILD_TABS = new Set(["Breakdowns", "ironmind"]);
/** Production site nav — only tabs in active use (telematics pilot + go-live). */
const PRODUCTION_SITE_TABS = [
  "mywork",
  "tasks",
  "dash",
  "daily",
  "maintenance",
  "assets",
  "telematics",
  "cartrack",
  "workshop",
  "fuel",
  "lube",
  "vehicle",
  "stock",
  "parts-tracking",
  "reports",
  "ironmind",
];
const PRODUCTION_NAV_ENABLED = true;
/** Dashboard cards hidden during production UI cleanup (remove class `hidden` in index.html to restore). */
function isDashSectionVisible(id) {
  const el = qs(id);
  return Boolean(el && !el.classList.contains("hidden"));
}
const DEFAULT_ROLE = "admin";
const DEFAULT_USER = "admin";
// Fail closed until the public server config confirms local legacy mode.
let LOGIN_GATE_ENABLED = true;
const DEFAULT_SITE = "main";
const LANG_KEY = "ironlog_lang";
const SIDEBAR_COLLAPSED_KEY = "ironlog_sidebar_collapsed";
const DEFAULT_LANG = "en";
const TASK_WORKSPACE_COLLAPSED_KEY = "ironlog_task_workspace_collapsed";
const TASK_SAVED_VIEWS_KEY = "ironlog_task_saved_views";
const DASHBOARD_DETAIL_KEY = "ironlog_dashboard_detail";
const WORKSHOP_LIBRARY_SETTINGS_KEY = "ironlog_workshop_library_settings";
const DEFAULT_WORKSHOP_LIBRARY_SETTINGS = Object.freeze({
  siteUrl: "https://iron-library.base44.app",
  apiBaseUrl: "",
  faultsPath: "faults",
  manualsPath: "manuals",
  repairsPath: "repairs",
});

// Sidebar navigation
function initSidebar() {
  const sidebar = qs("sidebar");
  const toggle = qs("sidebarToggle");
  const overlay = qs("sidebarOverlay");
  
  if (!sidebar) return;
  
  // Restore collapsed state
  if (localStorage.getItem(SIDEBAR_COLLAPSED_KEY) === "1") {
    sidebar.classList.add("collapsed");
    document.body.classList.add("sidebar-collapsed");
  }
  
  // Toggle handler
  toggle?.addEventListener("click", () => {
    const isCollapsed = sidebar.classList.toggle("collapsed");
    document.body.classList.toggle("sidebar-collapsed", isCollapsed);
    localStorage.setItem(SIDEBAR_COLLAPSED_KEY, isCollapsed ? "1" : "0");
    
    // On mobile, toggle overlay
    if (window.innerWidth <= 1024) {
      sidebar.classList.toggle("mobile-open", !isCollapsed);
      overlay?.classList.toggle("active", !isCollapsed);
    }
  });
  
  // Close sidebar on mobile when clicking overlay
  overlay?.addEventListener("click", () => {
    sidebar.classList.remove("mobile-open");
    overlay.classList.remove("active");
  });
  
  // Navigation item clicks (event delegation so dynamic task links also work)
  sidebar.addEventListener("click", (e) => {
    const item = e.target?.closest?.(".nav-item");
    if (!item || !sidebar.contains(item)) return;
    const href = String(item.getAttribute("href") || "").trim();
    const tab = item.dataset.tab;
    const isExternalPage = href && href !== "#" && /\.html(?:[?#]|$)/i.test(href);
    if (!tab || isExternalPage) {
      if (isExternalPage) {
        e.preventDefault();
        location.href = href;
        if (window.innerWidth <= 1024) {
          sidebar.classList.remove("mobile-open");
          overlay?.classList.remove("active");
        }
      }
      return;
    }
    e.preventDefault();
    const taskView = String(item.dataset.taskView || "").trim();
    const taskProject = String(item.dataset.taskProject || "").trim();
    const taskAssigned = String(item.dataset.taskAssigned || "").trim();
    const taskStatus = String(item.dataset.taskStatus || "").trim();
    const taskPriority = String(item.dataset.taskPriority || "").trim();
    const activeKey = String(item.dataset.activeKey || "").trim();

    switchTab(tab);
    if (tab === "tasks" && taskView) {
      setTaskSidebarView(taskView, {
        project: taskProject,
        assigned: taskAssigned,
        status: taskStatus,
        priority: taskPriority,
        activeKey,
        refresh: true
      });
    }

    // Close mobile sidebar
    if (window.innerWidth <= 1024) {
      sidebar.classList.remove("mobile-open");
      overlay?.classList.remove("active");
    }
  });
  
  // Handle window resize
  window.addEventListener("resize", () => {
    if (window.innerWidth > 1024) {
      sidebar.classList.remove("mobile-open");
      overlay?.classList.remove("active");
    }
  });

  initTaskWorkspaceSidebar();
}

function updateSidebarActiveState(activeTab) {
  const sidebar = qs("sidebar");
  if (!sidebar) return;
  
  sidebar.querySelectorAll(".nav-item").forEach((item) => {
    const activeKey = String(item.dataset.activeKey || "").trim();
    if (activeKey && activeTab === "tasks" && currentTaskSidebarActiveKey) {
      item.classList.toggle("active", activeKey === currentTaskSidebarActiveKey);
      return;
    }
    item.classList.toggle("active", item.dataset.tab === activeTab);
  });
  
  // Sync mobile nav
  const mobileSelect = qs("tabSelect");
  if (mobileSelect) {
    mobileSelect.value = activeTab;
  }
  
  // Update user display
  const userDisplay = qs("sessionUserDisplay");
  if (userDisplay) userDisplay.textContent = getSessionUser();
  const siteDisplay = qs("sessionSiteDisplay");
  if (siteDisplay) siteDisplay.textContent = getSessionSite();
  updateSidebarProfile();
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

function updateSidebarProfile() {
  const user = getSessionUser() || "User";
  const nameEl = qs("sidebarUserName");
  const roleEl = qs("sidebarRoleName");
  const initialEl = qs("sidebarUserInitial");
  if (nameEl) nameEl.textContent = user;
  if (roleEl) roleEl.textContent = getSessionRoles().map(roleDisplayName).join(" · ");
  if (initialEl) initialEl.textContent = user.charAt(0).toUpperCase() || "U";
}

function updateDashboardGreeting() {
  const welcome = qs("dashboardWelcome");
  const today = qs("dashboardToday");
  if (welcome) welcome.textContent = `Welcome, ${getSessionUser() || "User"}`;
  if (today) {
    today.textContent = new Intl.DateTimeFormat(undefined, {
      weekday: "long", day: "numeric", month: "long", year: "numeric"
    }).format(new Date());
  }
}

function initDashboardActionHub() {
  updateDashboardGreeting();
  document.querySelectorAll(".dashboard-quick-action").forEach((button) => {
    button.addEventListener("click", () => {
      const tab = String(button.dataset.goTab || "").trim();
      const href = String(button.dataset.goHref || "").trim();
      if (tab) switchTab(tab);
      else if (href) window.location.href = href;
    });
  });

  const grid = qs("dashboardGrid");
  const detailToggle = qs("dashboardDetailToggle");
  const setDetailed = (detailed) => {
    grid?.classList.toggle("show-details", Boolean(detailed));
    detailToggle?.setAttribute("aria-expanded", detailed ? "true" : "false");
    if (detailToggle) detailToggle.textContent = detailed ? "Show focused view" : "Show operational detail";
  };
  setDetailed(localStorage.getItem(DASHBOARD_DETAIL_KEY) === "1");
  detailToggle?.addEventListener("click", () => {
    const detailed = !grid?.classList.contains("show-details");
    setDetailed(detailed);
    localStorage.setItem(DASHBOARD_DETAIL_KEY, detailed ? "1" : "0");
  });
}

function isMaintenanceChildTab(tabKey) {
  return MAINT_CHILD_TABS.has(String(tabKey || "").trim());
}

function isAllowedDashboardTab(tabKey, allowed) {
  const k = String(tabKey || "").trim();
  if (!k) return false;
  return allowed.has(k);
}

function isBareChildTabEmbed() {
  return new URLSearchParams(window.location.search).get("bare") === "1";
}

function applyBareChildTabView() {
  if (!isBareChildTabEmbed()) return false;
  const tab = String(new URLSearchParams(window.location.search).get("tab") || "Breakdowns").trim();
  if (!tab || !isMaintenanceChildTab(tab)) return false;
  document.body.classList.add("bare-child-tab");
  document.querySelector(".sidebar")?.style.setProperty("display", "none");
  document.getElementById("sidebarOverlay")?.style.setProperty("display", "none");
  document.querySelector(".topbar")?.style.setProperty("display", "none");
  document.querySelector(".mobile-nav")?.style.setProperty("display", "none");
  document.querySelector("#mainContent > .card.stack-12")?.style.setProperty("display", "none");
  document.querySelectorAll(".panel").forEach((p) => p.classList.remove("show"));
  const panel = qs(`tab-${tab}`);
  if (panel) panel.classList.add("show");
  switchTab(tab);
  return true;
}

function resolveInitialTabFromUrl() {
  if (isBareChildTabEmbed()) return;
  const urlTab = String(new URLSearchParams(window.location.search).get("tab") || "").trim();
  if (!urlTab || !document.getElementById(`tab-${urlTab}`)) return;
  const allowed = new Set(getEffectiveAllowedTabs());
  if (!isAllowedDashboardTab(urlTab, allowed)) return;
  switchTab(urlTab);
}

const I18N = {
  en: {
    statusReady: "Ready.",
    docsTitle: "AI Assisted Site Documents",
    docsSubtitle: "Generate a draft from your standard header. No chat, only generate and approve (Yes/No).",
    docsHeaderTitle: "Standard Header",
    docsDraftTitle: "Draft Document",
  },
  af: {
    statusReady: "Gereed.",
    docsTitle: "KI Ondersteunde Terrein Dokumente",
    docsSubtitle: "Genereer 'n konsep vanaf jou standaard-opskrif. Geen chat nie, net genereer en goedkeur (Ja/Nee).",
    docsHeaderTitle: "Standaard Opskrif",
    docsDraftTitle: "Konsep Dokument",
  },
  zu: {
    statusReady: "Kulungele.",
    docsTitle: "Amadokhumenti Esiza Asekelwa yi-AI",
    docsSubtitle: "Dala idrafti usebenzisa i-header ejwayelekile. Akukho chat, dala bese uvuma (Yebo/Cha).",
    docsHeaderTitle: "I-Header Ejwayelekile",
    docsDraftTitle: "Idrafti Yedokhumenti",
  },
  pt: {
    statusReady: "Pronto.",
    docsTitle: "Documentos do Site Assistidos por IA",
    docsSubtitle: "Gerar um rascunho a partir do cabecalho padrao. Sem chat, apenas gerar e aprovar (Sim/Nao).",
    docsHeaderTitle: "Cabecalho Padrao",
    docsDraftTitle: "Rascunho do Documento",
  },
};
const UI_STRINGS = {
  af: {
    "User": "Gebruiker",
    "Role": "Rol",
    "Site": "Terrein",
    "Language": "Taal",
    "Date": "Datum",
    "Scheduled hrs": "Geskeduleerde ure",
    "Apply Role": "Pas Rol Toe",
    "Refresh": "Herlaai",
    "Section": "Afdeling",
    "Status:": "Status:",
    "Maintenance": "Instandhouding",
    "📊 Dashboard": "📊 Paneelbord",
    "📝 Daily Input": "📝 Daaglikse Invoer",
    "🛠️ Assets": "🛠️ Bates",
    "⛽ Fuel": "⛽ Brandstof",
    "🧴 Lube": "🧴 Smeermiddel",
    "📦 Stores": "📦 Stoor",
    "📚 Legal Docs": "📚 Regsdokumente",
    "📤 CSV Uploads": "📤 CSV Oplaai",
    "📄 Reports": "📄 Verslae",
    "✅ Approvals": "✅ Goedkeurings",
    "🏗️ Supply Flow": "🏗️ Verskaffingsvloei",
    "🏭 Operations": "🏭 Operasies",
    "🚚 Dispatch": "🚚 Versending",
    "🧪 Data Quality": "🧪 Datakwaliteit",
    "🔎 Audit Trail": "🔎 Ouditspoor",
    "🚧 Breakdown Ops": "🚧 Staking Operasies",
    "🌐 AI Documents": "🌐 KI Dokumente",
    "Reports": "Verslae",
    "Daily Input": "Daaglikse Invoer",
    "Assets": "Bates",
    "Fuel Input & OEM Benchmark": "Brandstof Invoer & OEM Maatstaf",
    "CSV Uploads": "CSV Oplaai",
    "Department Legal Documents": "Departement Regsdokumente",
    "Approval Workflows": "Goedkeurings Werkvloeie",
    "Audit Trail": "Ouditspoor",
    "Download Fuel CSV Template": "Laai Brandstof CSV Sjabloon Af",
    "Download Stores CSV Template": "Laai Stoor CSV Sjabloon Af",
    "Open Daily PDF": "Open Daaglikse PDF",
    "Open Weekly PDF": "Open Weeklikse PDF",
    "Open Lube PDF": "Open Smeermiddel PDF",
    "Open Stock Monitor PDF": "Open Voorraad Monitor PDF",
    "Download Stock Monitor PDF": "Laai Voorraad Monitor PDF Af",
    "Show DOWN only": "Wys net AF",
    "Loading dashboard...": "Laai paneelbord...",
    "Dashboard ready.": "Paneelbord gereed.",
    "Upload complete.": "Oplaai voltooi.",
    "Upload failed.": "Oplaai het misluk.",
    "Saved successfully.": "Suksesvol gestoor.",
    "Dashboard error": "Paneelbord fout",
    "Fuel benchmark error": "Brandstof maatstaf fout",
    "Legal load error": "Regs laai fout",
  },
  pt: {
    "User": "Utilizador",
    "Role": "Funcao",
    "Site": "Local",
    "Language": "Idioma",
    "Date": "Data",
    "Scheduled hrs": "Horas programadas",
    "Apply Role": "Aplicar Funcao",
    "Refresh": "Atualizar",
    "Section": "Secao",
    "Status:": "Estado:",
    "Maintenance": "Manutencao",
    "📊 Dashboard": "📊 Painel",
    "📝 Daily Input": "📝 Entrada Diaria",
    "🛠️ Assets": "🛠️ Ativos",
    "⛽ Fuel": "⛽ Combustivel",
    "🧴 Lube": "🧴 Lubrificante",
    "📦 Stores": "📦 Armazem",
    "📚 Legal Docs": "📚 Documentos Legais",
    "📤 CSV Uploads": "📤 Upload CSV",
    "📄 Reports": "📄 Relatorios",
    "✅ Approvals": "✅ Aprovacoes",
    "🏗️ Supply Flow": "🏗️ Fluxo de Suprimentos",
    "🏭 Operations": "🏭 Operacoes",
    "🚚 Dispatch": "🚚 Expedicao",
    "🧪 Data Quality": "🧪 Qualidade de Dados",
    "🔎 Audit Trail": "🔎 Trilha de Auditoria",
    "🚧 Breakdown Ops": "🚧 Operacoes de Avaria",
    "🌐 AI Documents": "🌐 Documentos IA",
    "Reports": "Relatorios",
    "Daily Input": "Entrada Diaria",
    "Assets": "Ativos",
    "Fuel Input & OEM Benchmark": "Entrada de Combustivel e Referencia OEM",
    "CSV Uploads": "Upload CSV",
    "Department Legal Documents": "Documentos Legais do Departamento",
    "Approval Workflows": "Fluxos de Aprovacao",
    "Audit Trail": "Trilha de Auditoria",
    "Download Fuel CSV Template": "Baixar Modelo CSV de Combustivel",
    "Download Stores CSV Template": "Baixar Modelo CSV de Armazem",
    "Open Daily PDF": "Abrir PDF Diario",
    "Open Weekly PDF": "Abrir PDF Semanal",
    "Open Lube PDF": "Abrir PDF de Lubrificante",
    "Open Stock Monitor PDF": "Abrir PDF de Stock",
    "Download Stock Monitor PDF": "Baixar PDF de Stock",
    "Show DOWN only": "Mostrar apenas PARADO",

    "Availability": "Disponibilidade",
    "Utilization": "Utilizacao",
    "Alerts": "Alertas",
    "Major Downtime": "Maior Paragem",
    "Downtime Reasons": "Razoes de Paragem",
    "Critical Low Stock": "Estoque Baixo Critico",
    "Open Work Orders": "Ordens de Servico em Aberto",
    "WO SLA Escalations": "Escalacoes de SLA (WO)",
    "Reliability (MTBF / LTTR)": "Confiabilidade (MTBF / LTTR)",
    "Cost Trend (12 Months)": "Tendencia de Custo (12 Meses)",
    "Cost Engine (Daily)": "Motor de Custo (Diario)",
    "Stock Monitor": "Monitor de Estoque",
    "Cost Setup": "Configuracao de Custo",

    "Lube Usage": "Uso de Lubrificante",
    "Issue Lube (to Equipment / Work Order)": "Fornecer Lubrificante (para Equipamento / Ordem de Servico)",
    "Lube Minimums & Reorder Alerts": "Minimos de Lubrificante e Alertas de Reposicao",
    "Receive Lube Stock (Top-up)": "Receber Estoque de Lubrificante (Reposicao)",
    "Lube Analytics (Type + Stock)": "Analitica de Lubrificante (Tipo + Estoque)",

    "Daily Input": "Entrada Diaria",

    "Load Defaults": "Carregar Padrões",
    "Save Defaults": "Salvar Padrões",
    "Load Lube": "Carregar Lubrificante",
    "Check Lube Stock": "Verificar Estoque de Lubrificante",
    "Save Lube": "Salvar Lubrificante",
    "Load Day": "Carregar Dia",
    "Copy Yesterday": "Copiar Ontem",
    "Apply": "Aplicar",
    "Save Day": "Salvar Dia",
    "Run Shift Self-Check": "Executar Auto-Check do Turno",
    "Export Self-Check TXT": "Exportar Auto-Check TXT",

    "Load Lube Analytics": "Carregar Analitica",
    "Set Minimum": "Definir Minimo",
    "Refresh Alerts": "Atualizar Alertas",

    "Receive Stock": "Receber Estoque",
    "Load Analytics": "Carregar Analitica",
    "Save Mapping": "Salvar Mapeamento",
    "Load Mappings": "Carregar Mapeamentos",

    "Download Daily Excel": "Baixar Excel Diario",
    "Refresh Reliability": "Atualizar Confiabilidade",
    "Open Daily PDF": "Abrir PDF Diario",

    "Loading dashboard...": "A carregar painel...",
    "Dashboard ready.": "Painel pronto.",
    "Upload complete.": "Upload concluido.",
    "Upload failed.": "Falha no upload.",
    "Saved successfully.": "Guardado com sucesso.",
    "Dashboard error": "Erro do painel",
    "Fuel benchmark error": "Erro de referencia de combustivel",
    "Legal load error": "Erro ao carregar legal",
    "Action (adjust_movement/close_work_order)": "Ação (ajuste_movement/fechar_ordem)",
    "Action (optional)": "Ação (opcional)",
    "Action note (used for submit/approve/reject/supersede)": "Nota de ação (usada para enviar/aprovar/rejeitar/substituir)",
    "Actual tonnes": "Toneladas reais",
    "Amount produced": "Quantidade produzida",
    "Approved by": "Aprovado por",
    "Asset code": "Código do ativo",
    "Asset code (e.g. A300AM)": "Código do ativo (ex.: A300AM)",
    "Asset code (optional)": "Código do ativo (opcional)",
    "Asset Downtime/hr": "Paragem do ativo/h",
    "Asset Fuel/L": "Combustível do ativo (L)",
    "Asset name / unit": "Nome do ativo / unidade",
    "Category": "Categoria",
    "Client": "Cliente",
    "Client delivered to": "Cliente entregue a",
    "Contractor name": "Nome do empreiteiro",
    "Controls / PPE": "Controles / EPI",
    "Counted qty": "Quantidade contada",
    "Cycle count reason (optional)": "Motivo da contagem (opcional)",
    "Decision note (optional)": "Nota da decisão (opcional)",
    "Department": "Departamento",
    "Description": "Descrição",
    "Doc type": "Tipo de documento",
    "Document title": "Título do documento",
    "Downtime/hr default": "Paragem/h padrão",
    "Draft ID": "ID do rascunho",
    "Driver": "Motorista",
    "Entity type": "Tipo de entidade",
    "Exception note": "Nota de exceção",
    "Exception owner": "Responsável da exceção",
    "Extra notes / requirements": "Notas / requisitos adicionais",
    "Filter part code...": "Filtrar código da peça...",
    "Fuel/L default": "Combustível (L) padrão",
    "Hazards / risks": "Perigos / riscos",
    "Header ID": "ID do cabeçalho",
    "Header profile name": "Nome do perfil do cabeçalho",
    "Hours filled": "Horas preenchidas",
    "Issued by": "Emitido por",
    "KM per hour factor": "Fator KM por hora",
    "Labor/hr default": "Mão de obra/h padrão",
    "Location code": "Código da localização",
    "Location code (e.g. MAIN)": "Código da localização (ex.: MAIN)",
    "Location name (optional)": "Nome da localização (opcional)",
    "Lube stock no": "Nº do stock de lubrificante",
    "Lube/Oil type (optional)": "Tipo de lubrificante/óleo (opcional)",
    "Lube/Q default": "Lubrificante/Q padrão",
    "Manual override chain (optional)": "Cadeia de override manual (opcional)",
    "Min stock": "Stock mínimo",
    "Module (optional)": "Módulo (opcional)",
    "Module (stock/workorders)": "Módulo (stock/ordens)",
    "New min": "Novo mínimo",
    "Notes (optional)": "Notas (opcional)",
    "Oil type key (exact)": "Chave do tipo de óleo (exato)",
    "Owner": "Responsável",
    "Part code": "Código da peça",
    "Part description": "Descrição da peça",
    "Part Unit Cost": "Custo unitário da peça",
    "PO Number (optional)": "Nº da PO (opcional)",
    "POD link / file path (optional)": "Link POD / caminho do ficheiro (opcional)",
    "POD ref number": "Referência POD",
    "Prepared by": "Preparado por",
    "Product delivered": "Produto entregue",
    "Product type": "Tipo de produto",
    "Qty requested": "Quantidade solicitada",
    "Reference (e.g. delivery note)": "Referência (ex.: guia de entrega)",
    "Reference (optional)": "Referência (opcional)",
    "Re-open reason (required when reopening closed day)": "Motivo para reabrir (obrigatório ao reabrir um dia fechado)",
    "Req value (R)": "Valor solicitado (R)",
    "Resolution note": "Nota de resolução",
    "Revision (e.g. Rev 1)": "Revisão (ex.: Rev 1)",
    "Scope / objective": "Escopo / objetivo",
    "Search title/type/owner...": "Pesquisar título/tipo/responsável...",
    "Shift (Day/Night)": "Turno (Dia/Noite)",
    "site code": "código do site",
    "Site name": "Nome do site",
    "Source / notes (optional)": "Fonte / notas (opcional)",
    "Stock code": "Código do stock",
    "Supersedes Doc ID": "Substitui ID do documento",
    "Supervisor sign-off name": "Nome para assinatura do supervisor",
    "Supplier (optional)": "Fornecedor (opcional)",
    "Target tonnes": "Toneladas alvo",
    "Tier 1 chain (comma names)": "Cadeia do nível 1 (nomes separados por vírgula)",
    "Tier 1 max value": "Valor máximo do nível 1",
    "Tier 2 chain (comma names)": "Cadeia do nível 2 (nomes separados por vírgula)",
    "Tier 2 max value": "Valor máximo do nível 2",
    "Tier 3 chain (> Tier 2 max)": "Cadeia do nível 3 (> nível 2 máx.)",
    "Title": "Título",
    "Tonnes moved": "Toneladas movimentadas",
    "Trip ID": "ID da viagem",
    "Trip no (optional)": "Nº da viagem (opcional)",
    "Truck reg": "Matrícula do camião",
    "Trucks delivered": "Camiões entregues",
    "Trucks loaded": "Camiões carregados",
    "Type of product produced": "Tipo de produto produzido",
    "username": "nome de utilizador",
    "Variance note (if any)": "Nota de variação (se houver)",
    "Version": "Versão",
    "Weighbridge amount": "Valor da balança",
    "Work order ID": "ID da ordem de serviço",
    "Work order ID (optional)": "ID da ordem de serviço (opcional)",
  },
  zu: {
    "User": "Umsebenzisi",
    "Role": "Indima",
    "Site": "Isiza",
    "Language": "Ulimi",
    "Date": "Usuku",
    "Scheduled hrs": "Amahora ahleliwe",
    "Apply Role": "Sebenzisa Indima",
    "Refresh": "Vuselela",
    "Section": "Isigaba",
    "Status:": "Isimo:",
    "Maintenance": "Ukunakekelwa",
    "📊 Dashboard": "📊 Ideshibhodi",
    "📝 Daily Input": "📝 Ukufaka Kwansuku Zonke",
    "🛠️ Assets": "🛠️ Impahla",
    "⛽ Fuel": "⛽ Uphethiloli",
    "🧴 Lube": "🧴 Uwoyela",
    "📦 Stores": "📦 Isitolo",
    "📚 Legal Docs": "📚 Imibhalo Yomthetho",
    "📤 CSV Uploads": "📤 Ukulayisha i-CSV",
    "📄 Reports": "📄 Imibiko",
    "✅ Approvals": "✅ Ukuvunywa",
    "🏗️ Supply Flow": "🏗️ Ukugeleza Kokuhlinzeka",
    "🏭 Operations": "🏭 Ukusebenza",
    "🚚 Dispatch": "🚚 Ukuthunyelwa",
    "🧪 Data Quality": "🧪 Ikhwalithi Yedatha",
    "🔎 Audit Trail": "🔎 Umkhondo Wokuhlola",
    "🚧 Breakdown Ops": "🚧 Ukusebenza Kokuphuka",
    "🌐 AI Documents": "🌐 Imibhalo ye-AI",
    "Loading dashboard...": "Ideshibhodi iyalayisha...",
    "Dashboard ready.": "Ideshibhodi isilungile.",
    "Upload complete.": "Ukulayisha kuqediwe.",
    "Upload failed.": "Ukulayisha kwehlulekile.",
    "Saved successfully.": "Kugcinwe ngempumelelo.",
    "Dashboard error": "Iphutha ledashibhodi",
    "Fuel benchmark error": "Iphutha lebhentshimakhi likaphethiloli",
    "Legal load error": "Iphutha lokulayisha okomthetho",
  },
};

function getLang() {
  const v = String(localStorage.getItem(LANG_KEY) || DEFAULT_LANG).trim().toLowerCase();
  return I18N[v] ? v : DEFAULT_LANG;
}
function setLang(v) {
  localStorage.setItem(LANG_KEY, I18N[v] ? v : DEFAULT_LANG);
}
function t(key) {
  const lang = getLang();
  return I18N[lang]?.[key] || I18N[DEFAULT_LANG]?.[key] || key;
}
function trUI(text, lang = getLang()) {
  const src = String(text || "");
  return UI_STRINGS[lang]?.[src] || src;
}
function translateStatusMessage(msg, lang = getLang()) {
  const raw = String(msg || "");
  const direct = trUI(raw, lang);
  if (direct !== raw) return direct;
  const idx = raw.indexOf(":");
  if (idx > 0) {
    const head = raw.slice(0, idx).trim();
    const tail = raw.slice(idx + 1);
    const headT = trUI(head, lang);
    if (headT !== head) return `${headT}:${tail}`;
  }
  return raw;
}
function applyGlobalPageTranslation() {
  const lang = getLang();
  document.querySelectorAll("[placeholder]").forEach((el) => {
    if (!el.dataset.i18nPlaceholder) {
      el.dataset.i18nPlaceholder = el.getAttribute("placeholder") || "";
    }
    const base = el.dataset.i18nPlaceholder || "";
    el.setAttribute("placeholder", trUI(base, lang));
  });

  document.querySelectorAll("option").forEach((opt) => {
    if (!opt.dataset.i18nLabel) opt.dataset.i18nLabel = opt.textContent || "";
    opt.textContent = trUI(opt.dataset.i18nLabel || "", lang);
  });

  // Translate a limited set of visible UI elements by exact-string match,
  // to avoid expensive full DOM text-node sweeps.
  const translateBySelector = (selector) => {
    document.querySelectorAll(selector).forEach((el) => {
      const src = String(el.dataset.i18nSrc || el.textContent || "").trim();
      if (!src) return;
      if (!el.dataset.i18nSrc) el.dataset.i18nSrc = src;
      const next = trUI(el.dataset.i18nSrc || "", lang);
      if (next && next !== src) el.textContent = next;
    });
  };

  translateBySelector("h1,h2,h3,h4");
  translateBySelector("button");
}

function getSessionRole() {
  return String(localStorage.getItem(ROLE_KEY) || DEFAULT_ROLE).trim().toLowerCase() || DEFAULT_ROLE;
}
function normalizeRoles(input, fallbackRole = DEFAULT_ROLE) {
  const base = Array.isArray(input) ? input : [];
  const out = Array.from(
    new Set(
      base
        .map((r) => String(r || "").trim().toLowerCase())
        .filter((r) => ["admin", "supervisor", "stores", "artisan", "operator"].includes(r))
    )
  );
  if (out.length) return out;
  const fb = String(fallbackRole || DEFAULT_ROLE).trim().toLowerCase() || DEFAULT_ROLE;
  return [fb];
}
function getSessionRoles() {
  const primary = getSessionRole();
  try {
    const raw = localStorage.getItem(ROLES_KEY);
    if (!raw) return [primary];
    return normalizeRoles(JSON.parse(raw), primary);
  } catch {
    return [primary];
  }
}
function renderSessionRolesBadge() {
  const badge = qs("sessionRolesBadge");
  if (!badge) return;
  const roles = getSessionRoles();
  const labels = roles.map(roleDisplayName);
  badge.className = "session-role";
  badge.textContent = labels.join(" · ");
  badge.title = `Active session ${roles.length === 1 ? "role" : "roles"}: ${labels.join(", ")}`;
}
function getSessionUser() {
  return String(localStorage.getItem(USER_KEY) || DEFAULT_USER).trim() || DEFAULT_USER;
}
function getSessionSite() {
  return String(localStorage.getItem(SITE_KEY) || DEFAULT_SITE).trim().toLowerCase() || DEFAULT_SITE;
}

function getAuthToken() {
  return String(localStorage.getItem(TOKEN_KEY) || sessionStorage.getItem(TOKEN_KEY) || "").trim();
}
function setAuthToken(t, remember = true) {
  if (t) {
    if (remember) {
      localStorage.setItem(TOKEN_KEY, t);
      sessionStorage.removeItem(TOKEN_KEY);
    } else {
      sessionStorage.setItem(TOKEN_KEY, t);
      localStorage.removeItem(TOKEN_KEY);
    }
  } else {
    localStorage.removeItem(TOKEN_KEY);
    sessionStorage.removeItem(TOKEN_KEY);
  }
}
function clearAuthSession() {
  setAuthToken("");
  localStorage.removeItem(TABS_OVERRIDE_KEY);
  localStorage.removeItem(ROLES_KEY);
}

function setSessionContext(user, role, site, roles = null) {
  const rolePrimary = String(role || DEFAULT_ROLE).trim().toLowerCase() || DEFAULT_ROLE;
  const roleList = normalizeRoles(Array.isArray(roles) ? roles : [rolePrimary], rolePrimary);
  localStorage.setItem(USER_KEY, String(user || DEFAULT_USER).trim() || DEFAULT_USER);
  localStorage.setItem(ROLE_KEY, rolePrimary);
  localStorage.setItem(ROLES_KEY, JSON.stringify(roleList));
  localStorage.setItem(SITE_KEY, String(site || DEFAULT_SITE).trim().toLowerCase() || DEFAULT_SITE);
}

function defaultLocationForRole(role) {
  const r = String(role || "").trim().toLowerCase();
  if (r === "stores") return "MAIN";
  if (r === "artisan") return "WORKSHOP";
  if (r === "operator") return "LUBE";
  if (r === "supervisor") return "MAIN";
  return "MAIN";
}

function getRoleDefaultLocation(role) {
  const r = String(role || "").trim().toLowerCase() || DEFAULT_ROLE;
  const key = `${LOC_DEFAULT_PREFIX}${r}`;
  const saved = String(localStorage.getItem(key) || "").trim().toUpperCase();
  return saved || defaultLocationForRole(r);
}

function setRoleDefaultLocation(role, locationCode) {
  const r = String(role || "").trim().toLowerCase() || DEFAULT_ROLE;
  const key = `${LOC_DEFAULT_PREFIX}${r}`;
  const v = String(locationCode || "").trim().toUpperCase();
  if (!v) return;
  localStorage.setItem(key, v);
}

function applyDefaultLocationsToInputs() {
  const def = getRoleDefaultLocation(getSessionRole());
  ["msLocation", "saLocation", "mlLocation"].forEach((id) => {
    const el = qs(id);
    if (!el) return;
    const current = String(el.value || "").trim();
    if (!current) el.value = def;
    el.placeholder = el.placeholder || "Location code";
  });
  loadBinCodeOptionsForLocation(qs("saLocation")?.value || "", "saBinCodeOptions").catch(() => {});
  loadBinCodeOptionsForLocation(qs("msLocation")?.value || "", "msBinCodeOptions").catch(() => {});
}

async function loadBinCodeOptionsForLocation(locationCode, listId) {
  const binList = qs(listId);
  if (!binList) return;
  const loc = String(locationCode || "").trim().toUpperCase();
  binList.innerHTML = "";
  if (!loc) return;
  try {
    const data = await fetchJson(`${API}/api/stock/bins?location_code=${encodeURIComponent(loc)}&active=1`);
    const rows = Array.isArray(data?.rows) ? data.rows : [];
    rows.forEach((b) => {
      const code = String(b.bin_code || "").trim();
      if (!code) return;
      const opt = document.createElement("option");
      opt.value = code;
      opt.textContent = `${code}${b.bin_name ? ` - ${b.bin_name}` : ""}`;
      binList.appendChild(opt);
    });
  } catch {
    // Optional enhancer: bin list is best-effort
  }
}

function getSlaOpenSameTab() {
  return String(localStorage.getItem(SLA_OPEN_SAME_TAB_KEY) || "0") === "1";
}

function setSlaOpenSameTab(v) {
  localStorage.setItem(SLA_OPEN_SAME_TAB_KEY, v ? "1" : "0");
}

function authHeaders(extra = {}) {
  const roles = getSessionRoles();
  const h = {
    ...extra,
    "x-user-name": getSessionUser(),
    "x-user-role": getSessionRole(),
    "x-user-roles": roles.join(","),
    "x-site-code": getSessionSite(),
  };
  const tok = getAuthToken();
  if (tok) h.Authorization = `Bearer ${tok}`;
  return h;
}

// --- Ensure fetchJson exists (paste-safe) ---
async function fetchJson(url, opts) {
  const nextOpts = { ...(opts || {}) };
  const headers = new Headers(nextOpts.headers || {});
  const roles = getSessionRoles();
  headers.set("x-user-name", getSessionUser());
  headers.set("x-user-role", getSessionRole());
  headers.set("x-user-roles", roles.join(","));
  headers.set("x-site-code", getSessionSite());
  const tok = getAuthToken();
  if (tok) headers.set("Authorization", `Bearer ${tok}`);
  if (typeof nextOpts.body === "string" && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }
  nextOpts.headers = headers;

  const res = await fetch(url, nextOpts);
  const text = await res.text();

  let data;
  try {
    data = JSON.parse(text);
  } catch {
    data = { raw: text };
  }

  if (!res.ok) {
    if (res.status === 401 && LOGIN_GATE_ENABLED && !String(url).includes("/api/auth/login")) {
      // No valid sign-in (expired, or signed out in another tab of this browser):
      // bring back the sign-in screen instead of showing a raw "login required".
      promptSignInAgain();
      throw Object.assign(new Error("Your sign-in has ended. Please sign in again, then repeat this step."), { status: 401 });
    }
    if ([502,503,504,524].includes(res.status)) throw new Error('Ironlog took too long to respond or is temporarily unavailable. Please retry shortly.');
    const message = typeof data.error === 'string' ? data.error : typeof data.message === 'string' ? data.message : '';
    const err = new Error(message || `Request failed (${res.status}). Please retry or contact your administrator.`);
    err.status = res.status;
    throw err;
  }
  return data;
}

async function loadAuthConfig() {
  try {
    const res = await fetch(`${API}/api/auth/config`, { cache: "no-store" });
    if (!res.ok) throw new Error(`Auth config failed (${res.status})`);
    const data = await res.json();
    LOGIN_GATE_ENABLED = data?.auth_required !== false;
  } catch (e) {
    LOGIN_GATE_ENABLED = true;
    console.warn("Unable to load auth config; sign-in remains required.", e);
  }
}

/** Opens a PDF (or other binary) in a new tab using the same auth headers as API calls. */
async function openAuthedPdf(url, fallbackName = "ironlog-report.pdf") {
  // Open the tab while handling the user's click. Some embedded browsers block tabs
  // created after fetch completes, but still allow this initial blank tab.
  const preview = window.open("", "_blank");
  if (preview) preview.opener = null;
  const res = await fetch(url, { headers: authHeaders() });
  const blob = await res.blob();
  if (!res.ok) {
    if (res.status === 401 && LOGIN_GATE_ENABLED && !String(url).includes("/api/auth/login")) {
      // No valid sign-in (expired, or signed out in another tab of this browser):
      // bring back the sign-in screen instead of showing a raw "login required".
      promptSignInAgain();
      throw Object.assign(new Error("Your sign-in has ended. Please sign in again, then repeat this step."), { status: 401 });
    }
    let msg = await blob.text().catch(() => "");
    try {
      const j = JSON.parse(msg);
      msg = j.error || j.message || msg;
    } catch {}
    throw new Error(msg || `Request failed (${res.status})`);
  }
  const blobUrl = URL.createObjectURL(blob);
  if (preview && !preview.closed) {
    preview.location.replace(blobUrl);
    setTimeout(() => URL.revokeObjectURL(blobUrl), 120000);
    return true;
  }
  // Embedded browsers do not always permit popup control. Open the authenticated
  // blob in the current tab rather than leaving the user with an empty new tab.
  if (typeof window.location?.assign === "function") {
    window.location.assign(blobUrl);
  } else {
    const anchor = document.createElement("a");
    anchor.href = blobUrl;
    anchor.download = fallbackName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
  }
  setTimeout(() => URL.revokeObjectURL(blobUrl), 120000);
  return false;
}

/** Downloads a protected report/export using the current IRONLOG login token. */
async function downloadAuthedFile(url, fallbackName = "ironlog-report") {
  try {
    const res = await fetch(url, { headers: authHeaders() });
    const blob = await res.blob();
    if (!res.ok) {
      if (res.status === 401 && LOGIN_GATE_ENABLED) {
        const had = Boolean(getAuthToken());
        clearAuthSession();
        if (had) {
          showLoginGate(true);
          updateAuthChrome();
        }
      }
      let msg = await blob.text().catch(() => "");
      try {
        const parsed = JSON.parse(msg);
        msg = parsed.error || parsed.message || msg;
      } catch {}
      throw new Error(msg || `Request failed (${res.status})`);
    }
    const disposition = String(res.headers.get("content-disposition") || "");
    const encoded = disposition.match(/filename\*=UTF-8''([^;]+)/i)?.[1];
    const plain = disposition.match(/filename="?([^";]+)"?/i)?.[1];
    const name = encoded ? decodeURIComponent(encoded) : (plain || fallbackName);
    const blobUrl = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = blobUrl;
    anchor.download = name;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(blobUrl), 120000);
    return true;
  } catch (err) {
    alert(`Could not download report: ${err.message || err}`);
    return false;
  }
}

/** Opens or downloads a protected report without losing the active IRONLOG session. */
function openAuthedReport(url, { download = false, filename = "ironlog-report.pdf" } = {}) {
  if (download) return downloadAuthedFile(url, filename);
  return openAuthedPdf(url, filename).catch((err) => {
    alert(`Could not open report: ${err.message || err}`);
    return false;
  });
}

function getRoleAllowedTabs(role) {
  const r = String(role || "").toLowerCase();
  if (r === "plant_clerk") return ["dash", "daily", "assets", "fuel", "lube", "vehicle", "reports", "docs", "tasks"];
  if (r === "workshop_admin") return ["dash", "daily", "assets", "telematics", "workshop", "maintenance", "stock", "parts-tracking", "reports", "Breakdowns", "approvals", "procurement", "ironmind", "docs", "vehicle", "tasks"];
  if (r === "operator") return ["dash", "daily", "workshop", "fuel", "lube", "legal", "operations", "ironmind", "docs", "vehicle", "tasks", "telematics", "cartrack"];
  if (r === "artisan") return ["workshop", "maintenance", "Breakdowns", "vehicle", "tasks"];
  if (r === "stores" || r === "storeman") return ["dash", "stock", "parts-tracking", "workshop", "reports", "procurement", "tasks"];
  if (r === "procurement") return ["dash", "stock", "parts-tracking", "workshop", "reports", "finance", "procurement", "operations", "quality", "docs", "tasks"];
  if (r === "plant_manager") return ["dash", "daily", "assets", "telematics", "cartrack", "workshop", "maintenance", "fuel", "lube", "stock", "parts-tracking", "reports", "finance", "procurement", "operations", "dispatch", "quality", "audit", "ironmind", "docs", "vehicle", "tasks"];
  if (r === "site_manager") return ["dash", "daily", "assets", "telematics", "cartrack", "workshop", "maintenance", "fuel", "lube", "stock", "parts-tracking", "reports", "finance", "procurement", "operations", "dispatch", "quality", "audit", "docs", "tasks"];
  if (r === "executive") return ["dash", "workshop", "reports", "finance", "operations", "quality", "audit", "docs", "tasks"];
  if (r === "supervisor") return ["dash", "daily", "assets", "telematics", "cartrack", "workshop", "maintenance", "fuel", "lube", "stock", "parts-tracking", "legal", "uploads", "reports", "finance", "enterprise", "exec", "Breakdowns", "approvals", "procurement", "operations", "dispatch", "quality", "audit", "ironmind", "docs", "vehicle", "tasks"];
  return [
    "dash",
    "daily",
    "assets",
    "telematics",
    "cartrack",
    "workshop",
    "maintenance",
    "fuel",
    "lube",
    "stock",
    "parts-tracking",
    "legal",
    "uploads",
    "reports",
    "finance",
    "enterprise",
    "exec",
    "Breakdowns",
    "approvals",
    "procurement",
    "operations",
    "dispatch",
    "quality",
    "audit",
    "ironmind",
    "docs",
    "vehicle",
    "tasks",
    "admin",
  ];
}

function getEffectiveAllowedTabs() {
  const role = getSessionRole();
  const roles = getSessionRoles();
  let list;
  const raw = localStorage.getItem(TABS_OVERRIDE_KEY);
  if (raw) {
    try {
      const arr = JSON.parse(raw);
      if (Array.isArray(arr) && arr.length) list = arr;
    } catch {}
  }
  if (!list) {
    list = Array.from(new Set(roles.flatMap((r) => getRoleAllowedTabs(r))));
  }
  if (PRODUCTION_NAV_ENABLED) {
    const production = new Set(PRODUCTION_SITE_TABS);
    list = list.filter((t) => production.has(t));
    // Borris stays reachable even for saved per-user tab lists created before it was restored.
    if (!list.includes("ironmind")) list = [...list, "ironmind"];
    // Tasks back the My Work list; per-user section lists cannot assign it, so always allow it.
    if (!list.includes("tasks")) list = [...list, "tasks"];
  } else {
    // Keep task workspace reachable even when older saved tab overrides exist.
    if (!list.includes("tasks")) list = [...list, "tasks"];
    // Workshop Library is public-facing and should stay reachable even for older saved tab overrides.
    if (!list.includes("workshop")) list = [...list, "workshop"];
    // Backward compatibility for older saved tab overrides created before Finance tab existed.
    if (roles.some((r) => ["admin", "supervisor", "stores", "procurement", "plant_manager", "site_manager", "executive"].includes(r)) && !list.includes("finance")) {
      list = [...list, "finance"];
    }
    // Enterprise + executive tabs visibility guards (introduced in big-out roll)
    if (roles.some((r) => ["admin", "supervisor", "executive"].includes(r))) {
      if (!list.includes("enterprise")) list = [...list, "enterprise"];
      if (!list.includes("exec")) list = [...list, "exec"];
    }
  }
  // "admin" is not an assignable section in the multiselect; always allow the User admin tab for these roles
  if (roles.includes("admin") && !list.includes("admin")) list = [...list, "admin"];
  // My Work is the home screen for every signed-in user; keep it first so it is the landing tab.
  list = ["mywork", ...list.filter((t) => t !== "mywork")];
  return list;
}

/** Puts the sections a role uses most straight after Home. */
const NAV_SECTION_PRIORITY = {
  storeman: ["stores"],
  stores: ["stores"],
  procurement: ["stores"],
  workshop_admin: ["workshop", "stores", "fleet"],
  plant_manager: ["workshop", "fleet"],
  site_manager: ["workshop", "fleet"],
  plant_clerk: ["fleet"],
};

function orderNavSectionsForRoles(sidebar, roles) {
  const sections = Array.from(sidebar.querySelectorAll(".nav-section[data-nav-section]"));
  if (!sections.length) return;
  const parent = sections[0].parentElement;
  if (!sections[0].dataset.navDefaultIndex) sections.forEach((el, i) => { el.dataset.navDefaultIndex = String(i); });
  const preferred = [];
  (roles.includes("admin") ? [] : roles).forEach((r) => {
    (NAV_SECTION_PRIORITY[r] || []).forEach((id) => { if (!preferred.includes(id)) preferred.push(id); });
  });
  const rank = (el) => {
    const id = el.dataset.navSection;
    if (id === "home") return -1;
    const i = preferred.indexOf(id);
    return i >= 0 ? i : 100 + Number(el.dataset.navDefaultIndex || 0);
  };
  sections.sort((a, b) => rank(a) - rank(b)).forEach((el) => parent.appendChild(el));
}

function applyRoleVisibility() {
  const role = getSessionRole();
  const roles = getSessionRoles();
  renderSessionRolesBadge();
  const allowedList = getEffectiveAllowedTabs();
  const allowed = new Set(allowedList);
  document.querySelectorAll(".dashboard-quick-action[data-required-tab]").forEach((button) => {
    button.hidden = !allowed.has(String(button.dataset.requiredTab || ""));
  });
  const tabSelect = qs("tabSelect");
  if (tabSelect) {
    Array.from(tabSelect.options).forEach((opt) => {
      if (opt.value === "admin") {
        opt.hidden = !roles.includes("admin");
        return;
      }
      opt.hidden = !isAllowedDashboardTab(opt.value, allowed);
    });
    tabSelect.querySelectorAll("optgroup").forEach((group) => {
      group.hidden = !Array.from(group.children).some((opt) => !opt.hidden);
    });
  }
  
  // Update sidebar nav items visibility
  const sidebar = qs("sidebar");
  if (sidebar) {
    sidebar.querySelectorAll(".nav-item").forEach((item) => {
      const tab = item.dataset.tab;
      const navKey = item.dataset.nav || tab;
      if (!navKey) return;
      if (tab === "admin") {
        item.style.display = roles.includes("admin") ? "" : "none";
        return;
      }
      if (navKey === "maintenance") {
        item.style.display = allowed.has("maintenance") ? "" : "none";
        return;
      }
      if (!tab) return;
      item.style.display = isAllowedDashboardTab(tab, allowed) ? "" : "none";
    });
    
    const techPortal = qs("navTechPortal");
    if (techPortal) {
      const workshop = ["artisan", "supervisor", "workshop_admin", "admin", "plant_manager", "site_manager"];
      techPortal.style.display = roles.some((r) => workshop.includes(String(r).toLowerCase())) ? "" : "none";
    }
    const taskWorkspaceHeader = sidebar.querySelector(".task-workspace-header");
    if (taskWorkspaceHeader) taskWorkspaceHeader.style.display = allowed.has("tasks") ? "" : "none";
    orderNavSectionsForRoles(sidebar, roles);

    // Hide nav sections that are empty
    sidebar.querySelectorAll(".nav-section").forEach((section) => {
      const visibleItems = section.querySelectorAll(".nav-item:not([style*='display: none'])");
      section.style.display = visibleItems.length === 0 ? "none" : "";
    });
  }

  const reopenBtn = qs("reopenOperationsDay");
  if (reopenBtn) reopenBtn.style.display = roles.some((r) => ["admin", "plant_manager", "workshop_admin"].includes(r)) ? "" : "none";

  const activePanel = document.querySelector(".panel.show");
  const activeKey = String(activePanel?.id || "").replace(/^tab-/, "");
  const urlParams = new URLSearchParams(window.location.search);
  const urlTab = String(urlParams.get("tab") || "").trim();
  const urlAssetCode = String(urlParams.get("asset_code") || "").trim().toUpperCase();
  const urlItemCode = String(urlParams.get("item_code") || "").trim().toUpperCase();
  if (urlAssetCode && allowed.has("vehicle")) {
    clPendingAssetCode = urlAssetCode;
  }
  if (urlItemCode && allowed.has("vehicle")) {
    clPendingSafetyItemCode = urlItemCode;
  }
  let preferredTab = urlTab && isAllowedDashboardTab(urlTab, allowed) ? urlTab : "";
  if (!preferredTab && (urlAssetCode || urlItemCode) && allowed.has("vehicle")) {
    preferredTab = "vehicle";
  }
  if (isBareChildTabEmbed()) {
    const bareTab = String(urlParams.get("tab") || "Breakdowns").trim();
    if (bareTab) switchTab(bareTab);
    updateSidebarActiveState(bareTab);
    return;
  }
  if (!activeKey || !isAllowedDashboardTab(activeKey, allowed)) {
    const target = preferredTab || allowedList[0];
    if (target) {
      switchTab(target);
      updateSidebarActiveState(target);
      return;
    }
  } else if (preferredTab && activeKey !== preferredTab) {
    switchTab(preferredTab);
    updateSidebarActiveState(preferredTab);
    return;
  } else if (tabSelect) {
    tabSelect.value = activeKey;
  }

  updateSidebarActiveState(activeKey);
}

function initGlobalSearch() {
  const searchInput = qs("globalSearch");
  if (!searchInput) return;
  
  let debounceTimer;
  
  searchInput.addEventListener("input", () => {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => {
      const query = String(searchInput.value || "").trim();
      if (!query) return;
      
      // Map common search terms to tabs
      const tabMappings = {
        "dashboard": "dash",
        "daily": "daily",
        "assets": "assets",
        "workshop": "workshop",
        "library": "workshop",
        "faults": "workshop",
        "manuals": "workshop",
        "repairs": "workshop",
        "fuel": "fuel",
        "lube": "lube",
        "stores": "stock",
        "parts": "parts-tracking",
        "parts tracking": "parts-tracking",
        "offsite": "parts-tracking",
        "legal": "legal",
        "reports": "reports",
        "finance": "finance",
        "budget": "finance",
        "forecast": "finance",
        "journal": "finance",
        "ssot": "finance",
        "approvals": "approvals",
        "supply": "procurement",
        "operations": "operations",
        "dispatch": "dispatch",
        "quality": "quality",
        "audit": "audit",
        "vehicle": "vehicle",
        "checklist": "vehicle",
        "checklists": "vehicle",
        "prestart": "vehicle",
        "admin": "admin",
        "ai": "docs",
        "ironmind": "ironmind"
      };
      
      const lowerQuery = query.toLowerCase();
      
      // Check for tab matches first
      for (const [key, tab] of Object.entries(tabMappings)) {
        if (lowerQuery.includes(key)) {
          switchTab(tab);
          searchInput.value = "";
          setStatus(`Navigated to ${tab}`);
          return;
        }
      }
      
      // Asset code search - navigate to assets tab
      if (query.match(/^[A-Z]{2,}\d{0,4}$/i) || query.match(/^[A-Z0-9-]+$/)) {
        switchTab("assets");
        setStatus(`Search: ${query} - check Assets tab`);
        searchInput.value = "";
        return;
      }
    }, 300);
  });
  
  searchInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      const query = String(searchInput.value || "").trim();
      if (query) {
        setStatus(`Search: "${query}"`);
      }
    }
    if (e.key === "Escape") {
      searchInput.value = "";
      searchInput.blur();
    }
  });
}

function getStoredWorkshopLibrarySettings() {
  try {
    const raw = localStorage.getItem(WORKSHOP_LIBRARY_SETTINGS_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    const siteUrl = String(parsed?.siteUrl || "").trim() || DEFAULT_WORKSHOP_LIBRARY_SETTINGS.siteUrl;
    const apiBaseUrl = String(parsed?.apiBaseUrl || "").trim();
    const faultsPath = String(parsed?.faultsPath || "").trim() || DEFAULT_WORKSHOP_LIBRARY_SETTINGS.faultsPath;
    const manualsPath = String(parsed?.manualsPath || "").trim() || DEFAULT_WORKSHOP_LIBRARY_SETTINGS.manualsPath;
    const repairsPath = String(parsed?.repairsPath || "").trim() || DEFAULT_WORKSHOP_LIBRARY_SETTINGS.repairsPath;
    return {
      ...DEFAULT_WORKSHOP_LIBRARY_SETTINGS,
      siteUrl,
      apiBaseUrl,
      faultsPath,
      manualsPath,
      repairsPath,
    };
  } catch {
    return { ...DEFAULT_WORKSHOP_LIBRARY_SETTINGS };
  }
}

function getWorkshopLibrarySettingsFromInputs() {
  const saved = getStoredWorkshopLibrarySettings();
  return {
    siteUrl: String(qs("workshopLibraryUrl")?.value || saved.siteUrl || "").trim(),
    apiBaseUrl: String(qs("workshopApiBaseUrl")?.value || saved.apiBaseUrl || "").trim(),
    faultsPath: String(qs("workshopFaultsEndpoint")?.value || saved.faultsPath || DEFAULT_WORKSHOP_LIBRARY_SETTINGS.faultsPath).trim(),
    manualsPath: String(qs("workshopManualsEndpoint")?.value || saved.manualsPath || DEFAULT_WORKSHOP_LIBRARY_SETTINGS.manualsPath).trim(),
    repairsPath: String(qs("workshopRepairsEndpoint")?.value || saved.repairsPath || DEFAULT_WORKSHOP_LIBRARY_SETTINGS.repairsPath).trim(),
  };
}

function hydrateWorkshopLibraryInputs(settings = getStoredWorkshopLibrarySettings()) {
  if (qs("workshopLibraryUrl")) qs("workshopLibraryUrl").value = String(settings.siteUrl || "");
  if (qs("workshopApiBaseUrl")) qs("workshopApiBaseUrl").value = String(settings.apiBaseUrl || "");
  if (qs("workshopFaultsEndpoint")) qs("workshopFaultsEndpoint").value = String(settings.faultsPath || DEFAULT_WORKSHOP_LIBRARY_SETTINGS.faultsPath);
  if (qs("workshopManualsEndpoint")) qs("workshopManualsEndpoint").value = String(settings.manualsPath || DEFAULT_WORKSHOP_LIBRARY_SETTINGS.manualsPath);
  if (qs("workshopRepairsEndpoint")) qs("workshopRepairsEndpoint").value = String(settings.repairsPath || DEFAULT_WORKSHOP_LIBRARY_SETTINGS.repairsPath);
}

function normalizeWorkshopUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (/^https?:\/\//i.test(raw)) return raw;
  if (/^www\./i.test(raw)) return `https://${raw}`;
  if (/^(localhost|127\.0\.0\.1)(:\d+)?(\/|$)/i.test(raw)) return `http://${raw}`;
  if (/^[a-z0-9.-]+\.[a-z]{2,}(\/|$)/i.test(raw)) return `https://${raw}`;
  if (raw.startsWith("/")) {
    try {
      return new URL(raw, API).toString();
    } catch {
      return raw;
    }
  }
  return raw;
}

function resolveWorkshopEndpointUrl(endpoint, apiBaseUrl) {
  const target = String(endpoint || "").trim();
  if (!target) return "";
  const direct = normalizeWorkshopUrl(target);
  if (/^https?:\/\//i.test(direct)) return direct;
  const base = normalizeWorkshopUrl(apiBaseUrl);
  if (!base || !/^https?:\/\//i.test(base)) return "";
  try {
    const baseUrl = base.endsWith("/") ? base : `${base}/`;
    return new URL(target.replace(/^\//, ""), baseUrl).toString();
  } catch {
    return "";
  }
}

function getWorkshopLibraryTarget(kind, settings = getWorkshopLibrarySettingsFromInputs()) {
  if (kind === "site") return normalizeWorkshopUrl(settings.siteUrl);
  if (kind === "faults") return resolveWorkshopEndpointUrl(settings.faultsPath, settings.apiBaseUrl);
  if (kind === "manuals") return resolveWorkshopEndpointUrl(settings.manualsPath, settings.apiBaseUrl);
  if (kind === "repairs") return resolveWorkshopEndpointUrl(settings.repairsPath, settings.apiBaseUrl);
  return "";
}

function saveWorkshopLibrarySettings() {
  const settings = getWorkshopLibrarySettingsFromInputs();
  localStorage.setItem(WORKSHOP_LIBRARY_SETTINGS_KEY, JSON.stringify(settings));
  const out = qs("workshopApiResult");
  if (out) {
    out.textContent =
      `Workshop Library links saved.\n\n` +
      `Library: ${settings.siteUrl || "(not set)"}\n` +
      `API base: ${settings.apiBaseUrl || "(not set)"}\n` +
      `Faults: ${settings.faultsPath || "(not set)"}\n` +
      `Manuals: ${settings.manualsPath || "(not set)"}\n` +
      `Repairs: ${settings.repairsPath || "(not set)"}`;
  }
  setStatus("Workshop Library links saved.");
}

function openWorkshopLibraryTarget(kind) {
  const labelMap = {
    site: "Workshop Library",
    faults: "Faults endpoint",
    manuals: "Manuals endpoint",
    repairs: "Repairs endpoint",
  };
  const url = getWorkshopLibraryTarget(kind);
  if (!url || !/^https?:\/\//i.test(url)) {
    alert(`Set a valid ${labelMap[kind] || "Workshop Library"} URL first.`);
    return false;
  }
  const opened = window.open(url, "_blank", "noopener,noreferrer");
  setStatus(`${labelMap[kind] || "Workshop Library"} opened.`);
  return Boolean(opened);
}

function openWorkshopLibraryOnTabActivate() {
  hydrateWorkshopLibraryInputs();
  const settings = getWorkshopLibrarySettingsFromInputs();
  const url = getWorkshopLibraryTarget("site", settings);
  const statusEl = qs("workshopLibraryStatus");
  if (!url || !/^https?:\/\//i.test(url)) {
    if (statusEl) {
      statusEl.textContent = "Workshop Library URL is not configured.";
    }
    setStatus("Workshop Library URL is not configured.");
    return;
  }
  const opened = openWorkshopLibraryTarget("site");
  if (statusEl) {
    statusEl.textContent = opened
      ? `Opened ${url} in a new tab. Use the button below if you need to open it again.`
      : `Could not open a new tab (popup blocked?). Open manually: ${url}`;
  }
}

function summarizeWorkshopApiPayload(payload, rawText) {
  if (Array.isArray(payload)) {
    return {
      summary: `${payload.length} item(s)`,
      preview: JSON.stringify(payload.slice(0, 2), null, 2),
    };
  }
  if (payload && typeof payload === "object") {
    if (Array.isArray(payload.items)) {
      return {
        summary: `${payload.items.length} item(s) in items`,
        preview: JSON.stringify(payload.items.slice(0, 2), null, 2),
      };
    }
    if (Array.isArray(payload.rows)) {
      return {
        summary: `${payload.rows.length} row(s)`,
        preview: JSON.stringify(payload.rows.slice(0, 2), null, 2),
      };
    }
    const keys = Object.keys(payload);
    return {
      summary: keys.length ? `keys: ${keys.slice(0, 8).join(", ")}` : "object response",
      preview: JSON.stringify(payload, null, 2),
    };
  }
  return {
    summary: "text response",
    preview: String(rawText || "").trim(),
  };
}

function clipWorkshopPreview(value, maxLen = 1600) {
  const text = String(value || "").trim();
  if (!text) return "(empty response)";
  return text.length > maxLen ? `${text.slice(0, maxLen)}\n...truncated...` : text;
}

async function checkWorkshopLibraryApiOverview() {
  const out = qs("workshopApiResult");
  if (!out) return;
  const settings = getWorkshopLibrarySettingsFromInputs();
  const targets = [
    { label: "Faults", url: getWorkshopLibraryTarget("faults", settings) },
    { label: "Manuals", url: getWorkshopLibraryTarget("manuals", settings) },
    { label: "Repairs", url: getWorkshopLibraryTarget("repairs", settings) },
  ].filter((entry) => entry.url);

  if (!targets.length) {
    out.textContent = "Add at least one Workshop Library API endpoint before checking the API overview.";
    setStatus("Workshop Library API not configured.");
    return;
  }

  out.textContent = "Checking Workshop Library endpoints...";
  setStatus("Checking Workshop Library API...");

  const results = await Promise.all(
    targets.map(async ({ label, url }) => {
      try {
        const res = await fetch(url, {
          headers: { Accept: "application/json,text/plain,*/*" },
        });
        const rawText = await res.text();
        let payload = rawText;
        try {
          payload = JSON.parse(rawText);
        } catch {}
        const summary = summarizeWorkshopApiPayload(payload, rawText);
        return {
          label,
          url,
          ok: res.ok,
          status: res.status,
          statusText: res.statusText || "",
          summary: summary.summary,
          preview: clipWorkshopPreview(summary.preview),
        };
      } catch (error) {
        return {
          label,
          url,
          ok: false,
          status: "fetch-error",
          statusText: "",
          summary: "Request failed",
          preview:
            `${error?.message || error}\n\n` +
            "If the endpoint is public but this still fails in the browser, check CORS for the IRONLOG origin.",
        };
      }
    })
  );

  out.textContent = results
    .map(
      (result) =>
        `[${result.label}] ${result.ok ? "OK" : "ERROR"} (${result.status}${result.statusText ? ` ${result.statusText}` : ""})\n` +
        `${result.url}\n` +
        `Summary: ${result.summary}\n\n` +
        `${result.preview}`
    )
    .join("\n\n------------------------------\n\n");

  setStatus("Workshop Library API overview loaded.");
}

function initWorkshopLibraryTab() {
  const panel = qs("tab-workshop");
  if (!panel) return;

  hydrateWorkshopLibraryInputs();

  qs("saveWorkshopLinksBtn")?.addEventListener("click", saveWorkshopLibrarySettings);
  qs("checkWorkshopApiBtn")?.addEventListener("click", () =>
    checkWorkshopLibraryApiOverview().catch((e) =>
      setStatus("Workshop Library API error: " + (e?.message || e))
    )
  );

  panel.addEventListener("click", (e) => {
    const target = e.target instanceof HTMLElement ? e.target.closest("[data-workshop-open]") : null;
    if (!target) return;
    e.preventDefault();
    const kind = String(target.getAttribute("data-workshop-open") || "").trim();
    if (!kind) return;
    openWorkshopLibraryTarget(kind);
  });
}

function initReportCardCollapsible() {
  document.querySelectorAll(".report-card.collapsible, .dash-card.collapsible").forEach((card) => {
    const header = card.querySelector(".report-card-header") || card.querySelector(".dash-card-header");
    const toggle = card.querySelector(".report-card-toggle");
    
    if (!header || !toggle) return;
    
    const toggleCollapse = () => {
      const isCollapsed = card.classList.toggle("collapsed");
      card.dataset.collapsed = isCollapsed;
      localStorage.setItem(`report_card_${card.querySelector("h3")?.textContent?.trim() || ""}`, isCollapsed ? "1" : "0");
    };
    
    header.addEventListener("click", (e) => {
      if (e.target.closest(".report-card-toggle")) {
        toggleCollapse();
      } else {
        toggleCollapse();
      }
    });
    
    toggle.addEventListener("click", (e) => {
      e.stopPropagation();
      toggleCollapse();
    });
    
    const cardName = card.querySelector("h3")?.textContent?.trim() || "";
    const savedState = localStorage.getItem(`report_card_${cardName}`);
    if (savedState === "1") {
      card.classList.add("collapsed");
      card.dataset.collapsed = "true";
    }
  });
}

function initSettingsDropdown() {
  const dropdown = qs("settingsDropdown");
  const btn = qs("settingsBtn");
  if (!dropdown || !btn) return;
  
  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    dropdown.classList.toggle("active");
  });
  
  document.addEventListener("click", (e) => {
    if (!dropdown.contains(e.target)) {
      dropdown.classList.remove("active");
    }
  });
  
  // Sync dropdown values with header values
  const syncValues = () => {
    const userEl = qs("sessionUser");
    const roleEl = qs("sessionRole");
    const siteEl = qs("sessionSite");
    const langEl = qs("languageSelect");
    if (userEl) userEl.value = getSessionUser();
    if (roleEl) roleEl.value = getSessionRole();
    if (siteEl) siteEl.value = getSessionSite();
    if (langEl) langEl.value = getLang();
  };
  
  btn.addEventListener("mouseenter", syncValues);
}

function initSessionControls() {
  const userEl = qs("sessionUser");
  const roleEl = qs("sessionRole");
  const siteEl = qs("sessionSite");
  const langEl = qs("languageSelect");
  if (userEl) userEl.value = getSessionUser();
  if (roleEl) roleEl.value = getSessionRole();
  if (siteEl) siteEl.value = getSessionSite();
  if (langEl) langEl.value = getLang();
  renderSessionRolesBadge();

  qs("applySessionRole")?.addEventListener("click", async () => {
    const u = String(userEl?.value || "").trim() || DEFAULT_USER;
    const r = String(roleEl?.value || "").trim().toLowerCase() || DEFAULT_ROLE;
    const s = String(siteEl?.value || "").trim().toLowerCase() || DEFAULT_SITE;
    setSessionContext(u, r, s);
    applyRoleVisibility();
    applyDefaultLocationsToInputs();
    try {
      const me = await fetchJson(`${API}/api/auth/me`);
      if (me.user?.id != null) applySessionFromMeUser(me.user);
      else localStorage.removeItem(TABS_OVERRIDE_KEY);
      applyRoleVisibility();
      setStatus(`Session: ${me.user?.username || u} (${me.user?.role || r}) @ ${s}`);
    } catch {
      setStatus(`Session applied: ${u} (${r}) @ ${s}`);
    }
  });

  const slaOpenSameTabEl = qs("slaOpenSameTab");
  if (slaOpenSameTabEl) {
    slaOpenSameTabEl.checked = getSlaOpenSameTab();
    slaOpenSameTabEl.addEventListener("change", () => {
      setSlaOpenSameTab(Boolean(slaOpenSameTabEl.checked));
    });
  }

  langEl?.addEventListener("change", () => {
    setLang(langEl.value);
    applyI18n();
    applyGlobalPageTranslation();
    setStatus(t("statusReady"));
  });

  qs("logoutBtn")?.addEventListener("click", () =>
    logoutAuth().catch((e) => setStatus("Logout error: " + e.message))
  );
}

/** Session gone while the page was open: show the sign-in screen with a short note. */
var signedInOnThisPage = false; // var: read by functions that may run before this line

function promptSignInAgain() {
  const alreadyShown = qs("loginOverlay")?.style.display === "flex";
  const wasSignedIn = signedInOnThisPage;
  clearAuthSession();
  showLoginGate(true);
  updateAuthChrome();
  const msg = qs("loginError");
  if (msg && !alreadyShown && wasSignedIn) msg.textContent = "Your sign-in has ended (for example you signed out in another tab). Please sign in again.";
}

// Signing out in another tab of this browser (e.g. the Technician portal) removes
// the shared sign-in; react straight away rather than on the next save.
window.addEventListener("storage", (e) => {
  if (e.key === TOKEN_KEY && e.oldValue && !e.newValue && LOGIN_GATE_ENABLED) promptSignInAgain();
});

function showLoginGate(on) {
  const el = qs("loginOverlay");
  if (!el) return;
  el.style.display = on ? "flex" : "none";
  document.body.classList.toggle("login-locked", Boolean(on));
}

function updateAuthChrome() {
  const tok = getAuthToken();
  signedInOnThisPage = Boolean(tok);
  const userEl = qs("sessionUser");
  const roleEl = qs("sessionRole");
  const applyBtn = qs("applySessionRole");
  const logoutBtn = qs("logoutBtn");
  if (userEl) userEl.disabled = Boolean(tok);
  if (roleEl) roleEl.disabled = Boolean(tok);
  if (applyBtn) applyBtn.style.display = tok ? "none" : "";
  if (logoutBtn) logoutBtn.style.display = tok ? "" : "none";
}

/** Reloads the My Work home screen once the session is known (defined in my-work.js). */
function refreshMyWorkIfVisible() {
  if (!qs("tab-mywork")?.classList.contains("show")) return;
  if (typeof loadMyWork === "function") loadMyWork().catch(() => {});
}

function applySessionFromMeUser(user) {
  if (!user) return;
  const u = String(user.username || DEFAULT_USER).trim() || DEFAULT_USER;
  const r = String(user.role || DEFAULT_ROLE).trim().toLowerCase() || DEFAULT_ROLE;
  const roles = normalizeRoles(user.roles, r);
  const allowedLoc = Array.isArray(user.allowed_locations) ? user.allowed_locations.map((x) => String(x || "").trim().toLowerCase()).filter(Boolean) : [];
  const currentSite = getSessionSite();
  const nextSite = allowedLoc.length ? (allowedLoc.includes(currentSite) ? currentSite : allowedLoc[0]) : currentSite;
  setSessionContext(u, r, nextSite, roles);
  if (user.allowed_tabs && Array.isArray(user.allowed_tabs) && user.allowed_tabs.length) {
    localStorage.setItem(TABS_OVERRIDE_KEY, JSON.stringify(user.allowed_tabs));
  } else {
    localStorage.removeItem(TABS_OVERRIDE_KEY);
  }
  renderSessionRolesBadge();
  updateDashboardGreeting();
}

function initReportsHub() {
  const panel = qs("tab-reports");
  if (!panel) return;
  const cards = Array.from(panel.querySelectorAll(".reports-grid > .report-card[data-report-category]"));
  const search = qs("reportHubSearch");
  const count = qs("reportHubCount");
  let activeFilter = "all";

  const render = () => {
    const term = String(search?.value || "").trim().toLowerCase();
    let visible = 0;
    cards.forEach((card) => {
      const categories = String(card.getAttribute("data-report-category") || "").split(/\s+/);
      const matchesCategory = activeFilter === "all" || categories.includes(activeFilter);
      const matchesSearch = !term || String(card.textContent || "").toLowerCase().includes(term);
      card.hidden = !(matchesCategory && matchesSearch);
      if (!card.hidden) visible += 1;
    });
    if (count) count.textContent = `${visible} report group${visible === 1 ? "" : "s"}`;
  };

  search?.addEventListener("input", render);
  panel.querySelectorAll("[data-report-filter]").forEach((button) => {
    button.addEventListener("click", () => {
      activeFilter = String(button.getAttribute("data-report-filter") || "all");
      panel.querySelectorAll("[data-report-filter]").forEach((item) => item.classList.toggle("is-active", item === button));
      render();
    });
  });
  render();
}

async function tryInitialSession() {
  if (!LOGIN_GATE_ENABLED) {
    showLoginGate(false);
    updateAuthChrome();
    const tok = getAuthToken();
    if (tok) {
      try {
        const data = await fetchJson(`${API}/api/auth/me`);
        if (data?.user?.id != null) applySessionFromMeUser(data.user);
      } catch {
        clearAuthSession();
      }
    }
    return;
  }

  const tok = getAuthToken();
  if (!tok) {
    showLoginGate(true);
    updateAuthChrome();
    return;
  }

  let res;
  let data = {};
  try {
    res = await fetch(`${API}/api/auth/me`, {
      headers: new Headers({ Authorization: `Bearer ${tok}` }),
    });
    data = await res.json();
  } catch {
    res = { status: 0 };
  }
  if (res.status === 401) {
    clearAuthSession();
    showLoginGate(true);
    updateAuthChrome();
    return;
  }
  if (data.ok && data.user && data.user.id != null) {
    applySessionFromMeUser(data.user);
    showLoginGate(false);
  } else {
    clearAuthSession();
    showLoginGate(true);
  }
  updateAuthChrome();
}

async function submitLoginForm() {
  const u = String(qs("loginUsername")?.value || "").trim();
  const p = String(qs("loginPassword")?.value || "");
  const setupCode = String(qs("loginSetupCode")?.value || "").trim();
  const setupPassword = String(qs("loginNewPassword")?.value || "").trim();
  const remember = qs("loginRemember")?.checked !== false;
  const errEl = qs("loginError");
  if (errEl) errEl.textContent = "";
  if (!u) {
    if (errEl) errEl.textContent = "Enter username.";
    return;
  }
  if (!p) {
    if (setupCode && setupPassword.length >= 6) {
      try {
        const setupRes = await fetch(`${API}/api/auth/setup-password`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ username: u, setup_code: setupCode, new_password: setupPassword }),
        });
        const setupData = await setupRes.json().catch(() => ({}));
        if (!setupRes.ok) {
          if (errEl) errEl.textContent = setupData.error || setupData.message || "Setup code failed.";
          return;
        }
        if (qs("loginSetupCode")) qs("loginSetupCode").value = "";
        if (qs("loginNewPassword")) qs("loginNewPassword").value = "";
        if (errEl) errEl.textContent = "Password created. Enter your password and sign in.";
      } catch (e) {
        if (errEl) errEl.textContent = String(e.message || e);
      }
      return;
    }
    if (errEl) errEl.textContent = "Enter password, or use setup code with a new password.";
    return;
  }
  try {
    const res = await fetch(`${API}/api/auth/login`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username: u, password: p }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
      if (data.error === "password_not_set") {
        if (!setupCode || setupPassword.length < 6) {
          if (errEl) errEl.textContent = "First-time setup: enter setup code and new password (6+ chars).";
          return;
        }
        const setupRes = await fetch(`${API}/api/auth/setup-password`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ username: u, setup_code: setupCode, new_password: setupPassword }),
        });
        const setupData = await setupRes.json().catch(() => ({}));
        if (!setupRes.ok) {
          if (errEl) errEl.textContent = setupData.error || setupData.message || "Setup code failed.";
          return;
        }
        if (qs("loginSetupCode")) qs("loginSetupCode").value = "";
        if (qs("loginNewPassword")) qs("loginNewPassword").value = "";
        if (errEl) errEl.textContent = "Password created. Please click Sign in again.";
        return;
      }
      if (errEl) errEl.textContent = data.message || data.error || "Login failed.";
      return;
    }
    setAuthToken(data.token, remember);
    applySessionFromMeUser(data.user);
    if (qs("loginPassword")) qs("loginPassword").value = "";
    showLoginGate(false);
    updateAuthChrome();
    initSessionControls();
    applyRoleVisibility();
    refreshMyWorkIfVisible();
    applyDefaultLocationsToInputs();
    setStatus(`Signed in as ${data.user?.username || u}`);
  } catch (e) {
    if (errEl) errEl.textContent = String(e.message || e);
  }
}

async function logoutAuth() {
  try {
    if (getAuthToken()) {
      await fetch(`${API}/api/auth/logout`, {
        method: "POST",
        headers: new Headers(authHeaders()),
      });
    }
  } catch {}
  clearAuthSession();
  updateAuthChrome();
  initSessionControls();
  applyRoleVisibility();
  await tryInitialSession();
}

/** Start-up: Login form controls. Called once from init() in init.js. */
function wireLoginControls() {
  qs("loginSubmit")?.addEventListener("click", () => submitLoginForm());
  qs("loginPassword")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") submitLoginForm();
  });
}

/**
 * Background panels load for every user, but some need a manager's rights.
 * When the server refuses (403), hide the panel quietly instead of showing an error.
 */
function hideWhenForbidden(err, elOrId) {
  if (Number(err?.status) !== 403) return false;
  const el = typeof elOrId === "string" ? qs(elOrId) : elOrId;
  const panel = el?.closest?.(".dash-card, .card, .panel-card, section:not(.panel)") || el;
  if (panel) panel.style.display = "none";
  return true;
}
