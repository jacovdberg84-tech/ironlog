// IRONLOG/web/app.js
const API = window.location?.origin || "http://localhost:3001";

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
/** Set true to require password sign-in before using the dashboard. */
const LOGIN_GATE_ENABLED = false;
const DEFAULT_SITE = "main";
const LANG_KEY = "ironlog_lang";
const SIDEBAR_COLLAPSED_KEY = "ironlog_sidebar_collapsed";
const DEFAULT_LANG = "en";
const TASK_WORKSPACE_COLLAPSED_KEY = "ironlog_task_workspace_collapsed";
const TASK_SAVED_VIEWS_KEY = "ironlog_task_saved_views";
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
  const isAdmin = roles.includes("admin");
  const isSupervisor = roles.includes("supervisor");
  const tone = isAdmin ? "red" : isSupervisor ? "orange" : "blue";
  badge.className = `pill ${tone}`;
  badge.textContent = `Roles: ${roles.join(", ")}`;
  badge.title = `Active session roles: ${roles.join(", ")}`;
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
    if (res.status === 401 && LOGIN_GATE_ENABLED) {
      const had = Boolean(getAuthToken());
      clearAuthSession();
      if (had) {
        showLoginGate(true);
        updateAuthChrome();
      }
    }
    throw new Error(data.error || data.message || text || `Request failed (${res.status})`);
  }
  return data;
}

/** Opens a PDF (or other binary) in a new tab using the same auth headers as API calls. */
async function openAuthedPdf(url) {
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
      const j = JSON.parse(msg);
      msg = j.error || j.message || msg;
    } catch {}
    throw new Error(msg || `Request failed (${res.status})`);
  }
  const u = URL.createObjectURL(blob);
  window.open(u, "_blank", "noopener,noreferrer");
  setTimeout(() => URL.revokeObjectURL(u), 120000);
}

function getRoleAllowedTabs(role) {
  const r = String(role || "").toLowerCase();
  if (r === "operator") return ["dash", "daily", "workshop", "fuel", "lube", "legal", "operations", "ironmind", "docs", "vehicle", "tasks", "telematics", "cartrack"];
  if (r === "artisan") return ["workshop", "maintenance", "Breakdowns", "vehicle", "tasks"];
  if (r === "stores" || r === "storeman") return ["dash", "maintenance", "stock", "parts-tracking", "workshop", "uploads", "reports", "finance", "legal", "procurement", "operations", "dispatch", "quality", "ironmind", "docs", "vehicle", "tasks", "telematics", "cartrack"];
  if (r === "procurement") return ["dash", "stock", "parts-tracking", "workshop", "reports", "finance", "procurement", "operations", "quality", "docs", "tasks"];
  if (r === "plant_manager") return ["dash", "daily", "assets", "telematics", "cartrack", "workshop", "maintenance", "fuel", "lube", "stock", "parts-tracking", "reports", "finance", "procurement", "operations", "dispatch", "quality", "audit", "docs", "tasks"];
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
    // IronMind stays reachable even for saved per-user tab lists created before it was restored.
    if (!list.includes("ironmind")) list = [...list, "ironmind"];
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
  if (roles.some((r) => ["admin", "supervisor"].includes(r)) && !list.includes("admin")) list = [...list, "admin"];
  return list;
}

function applyRoleVisibility() {
  const role = getSessionRole();
  const roles = getSessionRoles();
  renderSessionRolesBadge();
  const allowedList = getEffectiveAllowedTabs();
  const allowed = new Set(allowedList);
  const tabSelect = qs("tabSelect");
  if (tabSelect) {
    Array.from(tabSelect.options).forEach((opt) => {
      if (opt.value === "admin") {
        opt.hidden = !roles.some((r) => ["admin", "supervisor"].includes(r));
        return;
      }
      opt.hidden = !isAllowedDashboardTab(opt.value, allowed);
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
        item.style.display = roles.some((r) => ["admin", "supervisor"].includes(r)) ? "" : "none";
        return;
      }
      if (navKey === "maintenance") {
        item.style.display = allowed.has("maintenance") ? "" : "none";
        return;
      }
      if (!tab) return;
      item.style.display = isAllowedDashboardTab(tab, allowed) ? "" : "none";
    });
    
    // Hide nav sections that are empty
    sidebar.querySelectorAll(".nav-section").forEach((section) => {
      const visibleItems = section.querySelectorAll(".nav-item:not([style*='display: none'])");
      section.style.display = visibleItems.length === 0 ? "none" : "";
    });
  }

  const reopenBtn = qs("reopenOperationsDay");
  if (reopenBtn) reopenBtn.style.display = roles.some((r) => ["admin", "supervisor"].includes(r)) ? "" : "none";

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

function showLoginGate(on) {
  const el = qs("loginOverlay");
  if (!el) return;
  el.style.display = on ? "flex" : "none";
  document.body.classList.toggle("login-locked", Boolean(on));
}

function updateAuthChrome() {
  const tok = getAuthToken();
  const userEl = qs("sessionUser");
  const roleEl = qs("sessionRole");
  const applyBtn = qs("applySessionRole");
  const logoutBtn = qs("logoutBtn");
  if (userEl) userEl.disabled = Boolean(tok);
  if (roleEl) roleEl.disabled = Boolean(tok);
  if (applyBtn) applyBtn.style.display = tok ? "none" : "";
  if (logoutBtn) logoutBtn.style.display = tok ? "" : "none";
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

let __adminTabKeysLoaded = false;

const ADMIN_ROLE_OPTIONS = [
  { value: "admin", label: "admin — system administrator" },
  { value: "supervisor", label: "supervisor — workshop supervisor" },
  { value: "artisan", label: "artisan — technician (workshop terminal)" },
  { value: "operator", label: "operator — field operator" },
  { value: "storeman", label: "storeman — stores / stock" },
  { value: "stores", label: "stores — stores (legacy)" },
  { value: "procurement", label: "procurement — purchasing" },
  { value: "finance", label: "finance — finance & costing" },
  { value: "executive", label: "executive — executive reports" },
  { value: "site_manager", label: "site_manager — site manager" },
  { value: "plant_manager", label: "plant_manager — plant manager" },
  { value: "quality_manager", label: "quality_manager — quality" },
  { value: "hr_manager", label: "hr_manager — HR" },
];

function populateAdminRolesSelect(rolesSel, selectedValues = null) {
  if (!rolesSel) return;
  const selected = new Set(
    (selectedValues != null
      ? selectedValues
      : Array.from(rolesSel.selectedOptions || []).map((o) => o.value)
    ).map((v) => String(v || "").trim().toLowerCase()).filter(Boolean)
  );
  rolesSel.innerHTML = "";
  ADMIN_ROLE_OPTIONS.forEach(({ value, label }) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = label;
    opt.selected = selected.has(value);
    rolesSel.appendChild(opt);
  });
  if (!selected.size) {
    const op = rolesSel.querySelector('option[value="operator"]');
    if (op) op.selected = true;
  }
}

async function ensureAdminTabOptions() {
  const sel = qs("adminUserTabs");
  const rolesSel = qs("adminRoles");
  if (!sel && !rolesSel) return;

  if (rolesSel && rolesSel.options.length <= 1) {
    populateAdminRolesSelect(rolesSel);
  }

  if (__adminTabKeysLoaded) return;
  try {
    const data = await fetchJson(`${API}/api/auth/tabs`);
    const keys = Array.isArray(data.keys) ? data.keys : [];
    if (sel) {
      sel.innerHTML = "";
      keys.forEach((k) => {
        const opt = document.createElement("option");
        opt.value = k;
        opt.textContent = k;
        sel.appendChild(opt);
      });
    }
    const apiRoles = Array.isArray(data.roles) ? data.roles.map((r) => String(r || "").trim().toLowerCase()).filter(Boolean) : [];
    if (rolesSel && apiRoles.length) {
      const known = new Set(ADMIN_ROLE_OPTIONS.map((r) => r.value));
      const selectedBefore = Array.from(rolesSel.selectedOptions || []).map((o) => o.value);
      rolesSel.innerHTML = "";
      apiRoles.forEach((r) => {
        const meta = ADMIN_ROLE_OPTIONS.find((x) => x.value === r);
        const opt = document.createElement("option");
        opt.value = r;
        opt.textContent = meta ? meta.label : r;
        if (selectedBefore.includes(r) || (!selectedBefore.length && r === "operator")) opt.selected = true;
        rolesSel.appendChild(opt);
      });
      apiRoles.filter((r) => !known.has(r)).forEach((r) => {
        const opt = document.createElement("option");
        opt.value = r;
        opt.textContent = r;
        rolesSel.appendChild(opt);
      });
    } else if (rolesSel) {
      populateAdminRolesSelect(rolesSel);
    }
    __adminTabKeysLoaded = true;
  } catch {
    if (rolesSel && rolesSel.options.length <= 1) populateAdminRolesSelect(rolesSel);
  }
}

async function loadAdminUsers() {
  const pre = qs("adminUsersResult");
  const tbody = qs("adminUsersTbody");
  if (!tbody) return;
  setStatus("Loading users…");
  try {
    await ensureAdminTabOptions();
    const data = await fetchJson(`${API}/api/auth/users`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const knownLocations = new Set(["main"]);
    rows.forEach((r) => {
      if (Array.isArray(r.allowed_locations)) {
        r.allowed_locations.forEach((loc) => {
          const v = String(loc || "").trim().toLowerCase();
          if (v) knownLocations.add(v);
        });
      }
    });
    const locList = qs("adminAllowedLocationsList");
    if (locList) {
      locList.innerHTML = "";
      Array.from(knownLocations).sort().forEach((loc) => {
        const opt = document.createElement("option");
        opt.value = loc;
        locList.appendChild(opt);
      });
    }
    tbody.innerHTML = "";
    rows.forEach((r) => {
      const tr = document.createElement("tr");
      const rolesText = Array.isArray(r.roles) && r.roles.length ? r.roles.join(", ") : String(r.role || "operator");
      const locText = Array.isArray(r.allowed_locations) && r.allowed_locations.length ? r.allowed_locations.join(", ") : "all";
      tr.innerHTML = `<td>${escapeHtml(r.username)}</td><td>${escapeHtml(r.full_name || "")}</td><td>${escapeHtml(r.department || "")}</td><td>${escapeHtml(rolesText)}</td><td>${escapeHtml(locText)}</td><td>${r.active ? "yes" : "no"}</td><td>${r.has_password ? "yes" : "no"}</td><td>${r.has_pin ? "yes" : "no"}</td>`;
      tr.style.cursor = "pointer";
      tr.addEventListener("click", () => {
        if (qs("adminUsername")) qs("adminUsername").value = r.username;
        if (qs("adminFullName")) qs("adminFullName").value = r.full_name || "";
        if (qs("adminDepartment")) qs("adminDepartment").value = r.department || "";
        if (qs("adminAllowedLocations")) {
          const loc = Array.isArray(r.allowed_locations) ? r.allowed_locations.join(",") : "";
          qs("adminAllowedLocations").value = loc;
        }
        const rolesSel = qs("adminRoles");
        if (rolesSel) {
          const selectedRoles = Array.isArray(r.roles) && r.roles.length ? r.roles : [String(r.role || "operator")];
          Array.from(rolesSel.options).forEach((o) => {
            o.selected = selectedRoles.includes(o.value);
          });
        }
        if (qs("adminPassword")) qs("adminPassword").value = "";
        if (qs("adminPin")) qs("adminPin").value = "";
        if (qs("adminClearPin")) qs("adminClearPin").checked = false;
        const tabsSel = qs("adminUserTabs");
        if (tabsSel) {
          if (Array.isArray(r.allowed_tabs) && r.allowed_tabs.length) {
            Array.from(tabsSel.options).forEach((o) => {
              o.selected = r.allowed_tabs.includes(o.value);
            });
          } else {
            Array.from(tabsSel.options).forEach((o) => {
              o.selected = false;
            });
          }
        }
      });
      tbody.appendChild(tr);
    });
    if (pre) pre.textContent = JSON.stringify({ count: rows.length }, null, 2);
    setStatus("Users loaded.");
  } catch (e) {
    if (pre) pre.textContent = String(e.message || e);
    setStatus("Failed to load users.");
  }
}

function escapeHtml(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function applyAdminArtisanPreset() {
  const rolesSel = qs("adminRoles");
  const tabsSel = qs("adminUserTabs");
  if (rolesSel) {
    Array.from(rolesSel.options).forEach((o) => {
      o.selected = o.value === "artisan";
    });
  }
  const artisanTabs = getRoleAllowedTabs("artisan");
  if (tabsSel) {
    Array.from(tabsSel.options).forEach((o) => {
      o.selected = artisanTabs.includes(o.value);
    });
  }
  if (qs("adminDepartment") && !qs("adminDepartment").value) {
    qs("adminDepartment").value = "Workshop";
  }
  setStatus("Artisan preset applied — set a 4–6 digit PIN, then save user.");
}

async function saveAdminUser() {
  const username = String(qs("adminUsername")?.value || "").trim();
  const password = String(qs("adminPassword")?.value || "");
  const pin = String(qs("adminPin")?.value || "").replace(/\D/g, "");
  const clearPin = qs("adminClearPin")?.checked === true;
  const full_name = String(qs("adminFullName")?.value || "").trim();
  const department = String(qs("adminDepartment")?.value || "").trim();
  const allowedLocationsRaw = String(qs("adminAllowedLocations")?.value || "").trim();
  const issueSetup = qs("adminIssueSetupCode")?.checked !== false;
  const rolesSel = qs("adminRoles");
  const roles = rolesSel ? Array.from(rolesSel.selectedOptions).map((o) => String(o.value || "").trim().toLowerCase()).filter(Boolean) : [];
  if (!roles.length) return alert("Select at least one role.");
  const tabsSel = qs("adminUserTabs");
  const allowed_tabs = tabsSel ? Array.from(tabsSel.selectedOptions).map((o) => o.value) : [];
  if (!username) return alert("Username is required.");
  const body = {
    username,
    full_name: full_name || null,
    department: department || null,
    roles,
    role: roles[0],
    allowed_tabs,
    allowed_locations: allowedLocationsRaw || null,
    issue_setup_code: issueSetup,
  };
  if (password) body.password = password;
  if (clearPin) body.pin = "";
  else if (pin) {
    if (pin.length < 4 || pin.length > 6) return alert("PIN must be 4–6 digits.");
    body.pin = pin;
  }
  setStatus("Saving user…");
  try {
    const saved = await fetchJson(`${API}/api/auth/users`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const pre = qs("adminUsersResult");
    if (saved?.setup_code) {
      if (pre) pre.textContent = `Setup code for ${username}: ${saved.setup_code}\nExpires: ${saved.setup_code_expires_at || "7 days"}`;
      setStatus("User saved. Share setup code with user.");
    } else {
      setStatus("User saved.");
    }
    await loadAdminUsers();
  } catch (e) {
    setStatus("Save user failed: " + (e.message || e));
  }
}

let safetyTplItems = [];
let safetyReportSelectedCodes = new Set();
let lastSafetyQrUrl = "";

function setSafetyAdminResult(text) {
  const pre = qs("safetyAdminResult");
  if (pre) pre.textContent = String(text || "");
}

function renderSafetyTemplateEditor(items) {
  safetyTplItems = Array.isArray(items) ? items.map((it) => ({
    key: String(it.key || ""),
    label: String(it.label || ""),
  })) : [];
  const host = qs("safetyTplItems");
  if (!host) return;
  if (!safetyTplItems.length) {
    host.innerHTML = `<div class="muted small">No checklist rows.</div>`;
    return;
  }
  host.innerHTML = safetyTplItems.map((it, idx) => `
    <div class="item" style="display:flex; gap:8px; flex-wrap:wrap; align-items:center; margin-bottom:6px;">
      <input type="text" class="w-140" data-safety-tpl-key="${idx}" value="${escapeHtml(it.key)}" placeholder="key" />
      <input type="text" style="flex:1; min-width:200px;" data-safety-tpl-label="${idx}" value="${escapeHtml(it.label)}" placeholder="Label" />
      <button type="button" class="btn btn-secondary btn-sm" data-safety-tpl-remove="${idx}">Remove</button>
    </div>
  `).join("");
}

async function loadSafetyTemplatesSelect(preferredKey) {
  const data = await fetchJson(`${API}/api/safety/templates`);
  const templates = Array.isArray(data.templates) ? data.templates : [];
  const selects = ["safetyTplSelect", "safetyItemTemplate", "safetyPdfType"].map((id) => qs(id)).filter(Boolean);
  const prevKey = String(preferredKey || qs("safetyTplSelect")?.value || "").trim();
  selects.forEach((sel) => {
    const keepAll = sel.id === "safetyPdfType";
    const opts = templates.map((t) =>
      `<option value="${escapeHtml(t.template_key)}">${escapeHtml(t.title || t.template_key)}</option>`
    );
    sel.innerHTML = keepAll
      ? `<option value="">All types</option>${opts.join("")}`
      : opts.join("");
    if (prevKey && templates.some((t) => t.template_key === prevKey)) {
      sel.value = prevKey;
    }
  });
  renderSafetyCategoriesList(templates);
  return templates;
}

function renderSafetyCategoriesList(templates) {
  const host = qs("safetyCategoriesList");
  if (!host) return;
  const rows = Array.isArray(templates) ? templates : [];
  if (!rows.length) {
    host.innerHTML = `<div class="muted small">No categories yet — add one above.</div>`;
    return;
  }
  host.innerHTML = rows
    .map(
      (t) => `
    <div class="item safety-category-row" style="display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap; align-items:center;">
      <div>
        <strong>${escapeHtml(t.title || t.template_key)}</strong>
        <div class="muted small"><code>${escapeHtml(t.template_key)}</code> · ${Number(t.item_count || 0)} item(s) · ${Number(t.items?.length || 0)} checklist row(s)</div>
      </div>
      <button type="button" class="btn btn-secondary btn-sm" data-safety-edit-category="${escapeHtml(t.template_key)}">Edit checklist</button>
    </div>`
    )
    .join("");
}

async function addSafetyCategory() {
  const title = String(qs("safetyCategoryTitle")?.value || "").trim();
  const template_key = String(qs("safetyCategoryKey")?.value || "").trim();
  if (!title) return alert("Enter a category name (e.g. Cutting equipment).");
  setStatus("Adding safety category…");
  const body = { title };
  if (template_key) body.template_key = template_key;
  const data = await fetchJson(`${API}/api/safety/templates`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const key = String(data?.template?.template_key || template_key || "").trim();
  if (qs("safetyCategoryTitle")) qs("safetyCategoryTitle").value = "";
  if (qs("safetyCategoryKey")) qs("safetyCategoryKey").value = "";
  await loadSafetyTemplatesSelect(key);
  if (qs("safetyTplSelect") && key) qs("safetyTplSelect").value = key;
  if (qs("safetyItemTemplate") && key) qs("safetyItemTemplate").value = key;
  await loadSafetyTemplateEditor();
  setSafetyAdminResult(`Added category "${title}" (${key || "saved"}). Edit checklist rows below.`);
  setStatus("Safety category added.");
}

async function loadSafetyTemplateEditor() {
  const key = String(qs("safetyTplSelect")?.value || "fire_extinguisher").trim();
  setStatus("Loading safety template…");
  const data = await fetchJson(`${API}/api/safety/templates/${encodeURIComponent(key)}`);
  renderSafetyTemplateEditor(data?.template?.items || []);
  setSafetyAdminResult(`Loaded template: ${key}`);
  setStatus("Safety template loaded.");
}

async function saveSafetyTemplateEditor() {
  const key = String(qs("safetyTplSelect")?.value || "fire_extinguisher").trim();
  const items = safetyTplItems.map((_, idx) => {
    const keyInp = document.querySelector(`input[data-safety-tpl-key="${idx}"]`);
    const labelInp = document.querySelector(`input[data-safety-tpl-label="${idx}"]`);
    return {
      key: String(keyInp?.value || "").trim(),
      label: String(labelInp?.value || "").trim(),
    };
  }).filter((r) => r.key && r.label);
  if (!items.length) return alert("Add at least one checklist row.");
  setStatus("Saving safety template…");
  const data = await fetchJson(`${API}/api/safety/templates/${encodeURIComponent(key)}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ items }),
  });
  renderSafetyTemplateEditor(data?.template?.items || items);
  await loadSafetyTemplatesSelect(key);
  setSafetyAdminResult(`Saved template ${key} (${items.length} rows).`);
  setStatus("Safety template saved.");
}

async function loadSafetyItemsList() {
  const host = qs("safetyItemsList");
  if (!host) return;
  const data = await fetchJson(`${API}/api/safety/items`);
  const rows = Array.isArray(data.items) ? data.items : [];
  if (!rows.length) {
    safetyReportSelectedCodes = new Set();
    host.innerHTML = `<div class="muted small">No safety items registered yet.</div>`;
    return;
  }
  const validCodes = new Set(rows.map((r) => String(r.item_code || "").trim().toUpperCase()).filter(Boolean));
  safetyReportSelectedCodes = new Set([...safetyReportSelectedCodes].filter((c) => validCodes.has(c)));
  host.innerHTML = rows.map((r) => `
    <div class="item" style="display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap; align-items:center;">
      <div>
        <label style="display:flex; align-items:flex-start; gap:8px;">
          <input type="checkbox" data-safety-report-select="${escapeHtml(r.item_code)}" ${safetyReportSelectedCodes.has(String(r.item_code || "").trim().toUpperCase()) ? "checked" : ""} />
          <span>
            <strong>${escapeHtml(r.item_code)}</strong> — ${escapeHtml(r.item_name || r.template_title || "")}
            <div class="muted small">
              ${escapeHtml(r.template_title || r.template_key || "")}${r.location ? ` · ${escapeHtml(r.location)}` : ""}
              ${String(r.latest_status || "").toLowerCase() === "fail" ? ` · <span class="pill pill-red">FLAGGED</span>` : ""}
            </div>
          </span>
        </label>
      </div>
      <div class="row stack-10">
        <button type="button" class="btn btn-secondary btn-sm" data-safety-item-pdf="${escapeHtml(r.item_code)}" title="Individual inspection PDF">PDF</button>
        <button type="button" class="btn btn-secondary btn-sm" data-safety-use-qr="${escapeHtml(r.item_code)}">QR</button>
        <button type="button" class="btn btn-secondary btn-sm" data-safety-open-insp="${escapeHtml(r.item_code)}">Inspect</button>
        <button type="button" class="btn btn-secondary btn-sm" data-safety-remove-item="${Number(r.id)}">Remove</button>
      </div>
    </div>
  `).join("");
}

async function addSafetyEquipmentItem() {
  const item_code = String(qs("safetyItemCode")?.value || "").trim();
  const template_key = String(qs("safetyItemTemplate")?.value || "fire_extinguisher").trim();
  const item_name = String(qs("safetyItemName")?.value || "").trim();
  const location = String(qs("safetyItemLocation")?.value || "").trim();
  if (!item_code) return alert("Item code is required (e.g. FE-WS-01).");
  setStatus("Adding safety item…");
  await fetchJson(`${API}/api/safety/items`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ item_code, template_key, item_name, location }),
  });
  if (qs("safetyItemCode")) qs("safetyItemCode").value = "";
  if (qs("safetyItemName")) qs("safetyItemName").value = "";
  if (qs("safetyItemLocation")) qs("safetyItemLocation").value = "";
  await loadSafetyItemsList();
  await loadSafetyTemplatesSelect(template_key);
  setSafetyAdminResult(`Added ${item_code.toUpperCase()}.`);
  setStatus("Safety item added.");
}

async function removeSafetyEquipmentItem(id) {
  const rowId = Number(id || 0);
  if (!rowId) return;
  if (!window.confirm("Remove this safety item from the register?")) return;
  setStatus("Removing safety item…");
  await fetchJson(`${API}/api/safety/items/${rowId}`, { method: "DELETE" });
  await loadSafetyItemsList();
  await loadSafetyTemplatesSelect();
  setStatus("Safety item removed.");
}

async function buildSafetyQrImageData(itemCode) {
  const code = String(itemCode || "").trim().toUpperCase();
  if (!code) throw new Error("Item code is required.");
  const res = await fetchJson(`${API}/api/safety/items/${encodeURIComponent(code)}/qr-profile/refresh`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({}),
  });
  const scanValue = String(res?.qr_payload?.scan_url || "").trim();
  if (!scanValue) throw new Error("No QR scan URL generated.");
  const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=420x420&data=${encodeURIComponent(scanValue)}`;
  return { qrUrl, qrText: String(res?.qr_text || ""), scanValue };
}

async function generateSafetyQr() {
  const code = String(qs("safetyQrItemCode")?.value || qs("safetyItemCode")?.value || "").trim();
  if (!code) return alert("Enter a safety item code.");
  setStatus(`Generating QR for ${code}…`);
  const { qrUrl, qrText } = await buildSafetyQrImageData(code);
  lastSafetyQrUrl = qrUrl;
  const prev = qs("safetyQrPreview");
  const img = qs("safetyQrImg");
  const txt = qs("safetyQrText");
  if (img) img.src = qrUrl;
  if (txt) txt.textContent = qrText;
  if (prev) prev.style.display = "block";
  setSafetyAdminResult(`QR ready for ${code.toUpperCase()}.`);
  setStatus(`QR generated for ${code.toUpperCase()}.`);
}

function printSafetyQr() {
  if (!lastSafetyQrUrl) {
    alert("Generate a QR first.");
    return;
  }
  const code = String(qs("safetyQrItemCode")?.value || "").trim().toUpperCase();
  openQrLabelSheetPrintWindow(
    [{ code: code || "Safety item", qrUrl: lastSafetyQrUrl }],
    readQrSheetLayout("safety"),
    "IRONLOG Safety QR Label Sheet"
  );
  setStatus("Safety QR print sheet opened.");
}

async function printAllSafetyQrSheet() {
  setStatus("Building safety QR label sheet…");
  const data = await fetchJson(`${API}/api/safety/items`);
  const rows = Array.isArray(data.items) ? data.items : [];
  if (!rows.length) return alert("No safety items registered yet.");
  const labels = [];
  for (let i = 0; i < rows.length; i += 1) {
    const code = String(rows[i]?.item_code || "").trim();
    if (!code) continue;
    try {
      const { qrUrl } = await buildSafetyQrImageData(code);
      labels.push({ code, qrUrl });
      setStatus(`Preparing safety label ${i + 1}/${rows.length}: ${code}`);
      await new Promise((resolve) => setTimeout(resolve, 80));
    } catch {
      /* skip failed rows */
    }
  }
  if (!labels.length) throw new Error("Could not prepare any safety QR labels.");
  openQrLabelSheetPrintWindow(labels, readQrSheetLayout("safety"), "IRONLOG Safety QR Label Sheet");
  setStatus(`Safety QR sheet ready (${labels.length} labels) ✅`);
}

async function openSafetyRegisterPdf(blank = false) {
  const template_key = String(qs("safetyPdfType")?.value || "").trim();
  const date = String(qs("safetyPdfDate")?.value || "").trim() || new Date().toISOString().slice(0, 10);
  const q = new URLSearchParams();
  if (template_key) q.set("template_key", template_key);
  q.set("date", date);
  if (blank) q.set("blank", "1");
  setStatus(blank ? "Opening blank safety sheet…" : "Opening safety register PDF…");
  await openAuthedPdf(`${API}/api/safety/register.pdf?${q.toString()}`);
  setStatus("Safety PDF opened.");
}

async function openSafetyItemInspectionPdf(itemCode) {
  const code = String(itemCode || "").trim().toUpperCase();
  if (!code) return;
  const date =
    String(qs("safetyReportEndDate")?.value || "").trim() ||
    String(qs("safetyPdfDate")?.value || "").trim() ||
    new Date().toISOString().slice(0, 10);
  const q = new URLSearchParams();
  q.set("item_code", code);
  q.set("date", date);
  setStatus(`Opening safety PDF for ${code}…`);
  await openAuthedPdf(`${API}/api/safety/inspections/item.pdf?${q.toString()}`);
  setStatus(`Safety PDF opened for ${code}.`);
}

async function openSafetyInspectionReportPdf(selectedOnly = false) {
  const start = String(qs("safetyReportStartDate")?.value || "").trim() || new Date().toISOString().slice(0, 10);
  const end = String(qs("safetyReportEndDate")?.value || "").trim() || start;
  const template_key = String(qs("safetyPdfType")?.value || "").trim();
  const selectedCodes = [...safetyReportSelectedCodes];
  if (selectedOnly && !selectedCodes.length) {
    alert("Select at least one equipment item in the register list first.");
    return;
  }
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  if (template_key) q.set("template_key", template_key);
  if (selectedOnly && selectedCodes.length) q.set("item_codes", selectedCodes.join(","));
  setStatus(selectedOnly ? "Opening selected safety inspection report PDF…" : "Opening safety inspection report PDF…");
  await openAuthedPdf(`${API}/api/safety/inspections/report.pdf?${q.toString()}`);
  setStatus("Safety inspection report opened.");
}

async function initSafetyAdminPanel() {
  if (!qs("adminSafetyCard")) return;
  const pdfDate = qs("safetyPdfDate");
  if (pdfDate && !pdfDate.value) pdfDate.value = new Date().toISOString().slice(0, 10);
  const reportStart = qs("safetyReportStartDate");
  const reportEnd = qs("safetyReportEndDate");
  if (reportEnd && !reportEnd.value) reportEnd.value = new Date().toISOString().slice(0, 10);
  if (reportStart && !reportStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    reportStart.value = d.toISOString().slice(0, 10);
  }
  applySafetyQrSheetPreset();
  try {
  await loadSafetyTemplatesSelect();
  await loadSafetyTemplateEditor();
    await loadSafetyItemsList();
  } catch (e) {
    setSafetyAdminResult(String(e.message || e));
  }
}

function setTelemAdminResult(text) {
  const pre = qs("telemAdminResult");
  if (pre) pre.textContent = String(text || "");
}

function telemLinkStatusPill(status) {
  const st = String(status || "offline").toLowerCase();
  if (st === "live") return `<span class="pill green" style="font-size:0.65rem;">LIVE</span>`;
  if (st === "stale") return `<span class="pill amber" style="font-size:0.65rem;">STALE</span>`;
  if (st === "inactive") return `<span class="pill" style="font-size:0.65rem;">INACTIVE</span>`;
  return `<span class="pill" style="font-size:0.65rem;">OFFLINE</span>`;
}

async function loadTelematicsAdminDevices() {
  const host = qs("telemDevicesList");
  if (!host) return;
  const showInactive = qs("telemShowInactive")?.checked === true;
  setStatus("Loading telematics units…");
  try {
    const q = showInactive ? "?all=1" : "";
    const data = await fetchJson(`${API}/api/telematics/devices${q}`);
    const rows = Array.isArray(data.devices) ? data.devices : [];
    if (!rows.length) {
      host.innerHTML = `<div class="muted small">No telematics units registered yet. Add one above.</div>`;
      setStatus("No telematics units.");
      return;
    }
    host.innerHTML = rows.map((r) => {
      const active = Number(r.active) === 1 || r.active === true;
      const meter = r.engine_hours == null ? "—" : `${Number(r.engine_hours).toFixed(1)} h`;
      const lastSeen = String(r.recorded_at || r.snapshot_updated_at || r.updated_at || "—");
      return `
        <div class="item" style="display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap; align-items:center; opacity:${active ? "1" : "0.65"};">
          <div>
            <strong>${escapeHtml(r.asset_code)}</strong> ${telemLinkStatusPill(active ? r.link_status : "inactive")}
            <div class="muted small">${escapeHtml(r.asset_name || "")} · ${escapeHtml(r.unit_model || "FSC")} · S/N ${escapeHtml(r.device_serial || "")}</div>
            <div class="muted small">Meter: ${escapeHtml(meter)} · Last seen: ${escapeHtml(lastSeen)}</div>
          </div>
          <div class="row stack-10">
            <button type="button" class="btn btn-secondary btn-sm" data-telem-edit="${Number(r.id)}" data-telem-asset="${escapeHtml(r.asset_code)}" data-telem-serial="${escapeHtml(r.device_serial || "")}" data-telem-model="${escapeHtml(r.unit_model || "FSC650")}" data-telem-ext="${escapeHtml(r.external_id || "")}">Replace unit</button>
            ${active ? `<button type="button" class="btn btn-secondary btn-sm" data-telem-deactivate="${Number(r.id)}" data-telem-asset-label="${escapeHtml(r.asset_code)}">Deactivate</button>` : ""}
          </div>
        </div>
      `;
    }).join("");
    setStatus(`Telematics units loaded (${rows.length}).`);
  } catch (e) {
    host.innerHTML = `<div class="muted small">Load failed: ${escapeHtml(e.message || String(e))}</div>`;
    setStatus("Telematics load failed.");
  }
}

function fillTelematicsDeviceForm({ assetCode, deviceSerial, unitModel, externalId, replaceFaulty = false }) {
  if (qs("telemAssetCode")) qs("telemAssetCode").value = String(assetCode || "");
  if (qs("telemDeviceSerial")) qs("telemDeviceSerial").value = String(deviceSerial || "");
  if (qs("telemUnitModel")) qs("telemUnitModel").value = String(unitModel || "FSC650");
  if (qs("telemExternalId")) qs("telemExternalId").value = String(externalId || "");
  if (qs("telemReplaceFaulty")) qs("telemReplaceFaulty").checked = Boolean(replaceFaulty);
  qs("telemDeviceSerial")?.focus();
}

async function saveTelematicsDevice() {
  const asset_code = String(qs("telemAssetCode")?.value || "").trim();
  const device_serial = String(qs("telemDeviceSerial")?.value || "").trim();
  const unit_model = String(qs("telemUnitModel")?.value || "FSC650").trim();
  const external_id = String(qs("telemExternalId")?.value || "").trim();
  const replace_faulty = qs("telemReplaceFaulty")?.checked === true;
  if (!asset_code || !device_serial) return alert("Asset code and device serial are required.");
  setStatus("Saving telematics unit…");
  const data = await fetchJson(`${API}/api/telematics/devices`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({
      asset_code,
      device_serial,
      unit_model,
      external_id: external_id || device_serial,
      replace_faulty,
    }),
  });
  const msg = data.replaced
    ? `Replaced unit on ${asset_code.toUpperCase()} → serial ${device_serial}.`
    : data.created
      ? `Registered ${device_serial} on ${asset_code.toUpperCase()}.`
      : `Updated ${asset_code.toUpperCase()}.`;
  setTelemAdminResult(msg);
  if (qs("telemReplaceFaulty")) qs("telemReplaceFaulty").checked = false;
  await loadTelematicsAdminDevices();
  loadTelematicsFleet().catch(() => {});
  setStatus(msg);
}

async function deactivateTelematicsDeviceAdmin(id, assetLabel) {
  const deviceId = Number(id || 0);
  if (!deviceId) return;
  const label = String(assetLabel || "this asset").trim();
  if (!window.confirm(`Deactivate telematics unit for ${label}? Daily hours will unlock until a new unit is registered and reports.`)) return;
  setStatus("Deactivating telematics unit…");
  const data = await fetchJson(`${API}/api/telematics/devices/${deviceId}/deactivate`, {
    method: "POST",
    headers: authHeaders(),
  });
  setTelemAdminResult(`Deactivated ${data.device_serial || "unit"} on ${data.asset_code || label}.`);
  await loadTelematicsAdminDevices();
  loadTelematicsFleet().catch(() => {});
  setStatus("Telematics unit deactivated.");
}

async function initTelematicsAdminPanel() {
  if (!qs("adminTelematicsCard")) return;
  await loadTelematicsAdminDevices().catch((e) => setTelemAdminResult(String(e.message || e)));
}

function applyMdmPolicyCheckboxes(policies) {
  const p = policies && typeof policies === "object" ? policies : {};
  const asset = Array.isArray(p.asset) ? p.asset : [];
  const part = Array.isArray(p.part_stock_intake) ? p.part_stock_intake : [];
  const setChk = (id, key, list) => {
    const el = qs(id);
    if (el) el.checked = list.includes(key);
  };
  setChk("mdmPolAssetDept", "department_code", asset);
  setChk("mdmPolAssetCc", "cost_center_code", asset);
  setChk("mdmPolAssetOwner", "data_owner_username", asset);
  setChk("mdmPolPartDept", "department_code", part);
  setChk("mdmPolPartSup", "default_supplier_code", part);
  setChk("mdmPolPartOwner", "data_owner_username", part);
}

async function saveMdmPolicies() {
  setStatus("Saving policies…");
  try {
    const policies = {
      asset: [],
      part_stock_intake: [],
    };
    if (qs("mdmPolAssetDept")?.checked) policies.asset.push("department_code");
    if (qs("mdmPolAssetCc")?.checked) policies.asset.push("cost_center_code");
    if (qs("mdmPolAssetOwner")?.checked) policies.asset.push("data_owner_username");
    if (qs("mdmPolPartDept")?.checked) policies.part_stock_intake.push("department_code");
    if (qs("mdmPolPartSup")?.checked) policies.part_stock_intake.push("default_supplier_code");
    if (qs("mdmPolPartOwner")?.checked) policies.part_stock_intake.push("data_owner_username");
    const res = await fetchJson(`${API}/api/masterdata/policies`, {
      method: "PUT",
      body: JSON.stringify({ policies }),
    });
    applyMdmPolicyCheckboxes(res.policies);
    setStatus("Policies saved.");
  } catch (e) {
    setStatus("Policy save failed: " + (e.message || e));
  }
}

async function loadMasterDataGovernance() {
  const sumEl = qs("mdmSummaryOut");
  const dBody = qs("mdmDeptTbody");
  const cBody = qs("mdmCcTbody");
  const sBody = qs("mdmSupTbody");
  if (!dBody || !cBody || !sBody) return;
  setStatus("Loading master data…");
  try {
    const [sum, dep, cc, sup, pol] = await Promise.all([
      fetchJson(`${API}/api/masterdata/summary`),
      fetchJson(`${API}/api/masterdata/departments`),
      fetchJson(`${API}/api/masterdata/cost-centers`),
      fetchJson(`${API}/api/masterdata/suppliers`),
      fetchJson(`${API}/api/masterdata/policies`),
    ]);
    applyMdmPolicyCheckboxes(pol.policies);
    if (sumEl) {
      sumEl.textContent = JSON.stringify(
        { site: sum?.site_code, counts: sum?.counts },
        null,
        2
      );
    }
    const dRows = Array.isArray(dep.rows) ? dep.rows : [];
    dBody.innerHTML = dRows
      .map(
        (r) =>
          `<tr><td>${escapeHtml(r.code)}</td><td>${escapeHtml(r.name || "")}</td><td>${escapeHtml(r.owner_username || "—")}</td><td>${r.active ? "yes" : "no"}</td></tr>`
      )
      .join("");
    const cRows = Array.isArray(cc.rows) ? cc.rows : [];
    cBody.innerHTML = cRows
      .map(
        (r) =>
          `<tr><td>${escapeHtml(r.code)}</td><td>${escapeHtml(r.name || "")}</td><td>${escapeHtml(r.department_code || "—")}</td><td>${r.active ? "yes" : "no"}</td></tr>`
      )
      .join("");
    const sRows = Array.isArray(sup.rows) ? sup.rows : [];
    sBody.innerHTML = sRows
      .map(
        (r) =>
          `<tr><td>${escapeHtml(r.supplier_code)}</td><td>${escapeHtml(r.name || "")}</td><td>${escapeHtml(r.contact_email || "—")}</td><td>${r.active ? "yes" : "no"}</td></tr>`
      )
      .join("");
    setStatus("Master data loaded.");
  } catch (e) {
    if (sumEl) sumEl.textContent = String(e.message || e);
    setStatus("Master data load failed.");
  }
}

async function saveMdmDepartment() {
  const code = String(qs("mdmDeptCode")?.value || "").trim();
  const name = String(qs("mdmDeptName")?.value || "").trim();
  const owner_username = String(qs("mdmDeptOwner")?.value || "").trim() || null;
  if (!code || !name) return alert("Department code and name are required.");
  setStatus("Saving department…");
  try {
    await fetchJson(`${API}/api/masterdata/departments`, {
      method: "POST",
      body: JSON.stringify({ code, name, owner_username }),
    });
    if (qs("mdmDeptCode")) qs("mdmDeptCode").value = "";
    if (qs("mdmDeptName")) qs("mdmDeptName").value = "";
    if (qs("mdmDeptOwner")) qs("mdmDeptOwner").value = "";
    await loadMasterDataGovernance();
    setStatus("Department saved.");
  } catch (e) {
    setStatus("Department save failed: " + (e.message || e));
  }
}

async function saveMdmCostCenter() {
  const code = String(qs("mdmCcCode")?.value || "").trim();
  const name = String(qs("mdmCcName")?.value || "").trim();
  const department_code = String(qs("mdmCcDept")?.value || "").trim() || null;
  if (!code || !name) return alert("Cost center code and name are required.");
  setStatus("Saving cost center…");
  try {
    await fetchJson(`${API}/api/masterdata/cost-centers`, {
      method: "POST",
      body: JSON.stringify({ code, name, department_code }),
    });
    if (qs("mdmCcCode")) qs("mdmCcCode").value = "";
    if (qs("mdmCcName")) qs("mdmCcName").value = "";
    if (qs("mdmCcDept")) qs("mdmCcDept").value = "";
    await loadMasterDataGovernance();
    setStatus("Cost center saved.");
  } catch (e) {
    setStatus("Cost center save failed: " + (e.message || e));
  }
}

async function saveMdmSupplier() {
  const supplier_code = String(qs("mdmSupCode")?.value || "").trim();
  const name = String(qs("mdmSupName")?.value || "").trim();
  const contact_email = String(qs("mdmSupEmail")?.value || "").trim() || null;
  if (!supplier_code || !name) return alert("Supplier code and name are required.");
  setStatus("Saving supplier…");
  try {
    await fetchJson(`${API}/api/masterdata/suppliers`, {
      method: "POST",
      body: JSON.stringify({ supplier_code, name, contact_email }),
    });
    if (qs("mdmSupCode")) qs("mdmSupCode").value = "";
    if (qs("mdmSupName")) qs("mdmSupName").value = "";
    if (qs("mdmSupEmail")) qs("mdmSupEmail").value = "";
    await loadMasterDataGovernance();
    setStatus("Supplier saved.");
  } catch (e) {
    setStatus("Supplier save failed: " + (e.message || e));
  }
}

async function submitChangePassword() {
  const old_password = String(qs("chPwdOld")?.value || "");
  const new_password = String(qs("chPwdNew")?.value || "").trim();
  const out = qs("chPwdResult");
  if (out) out.textContent = "";
  if (new_password.length < 6) return alert("New password must be at least 6 characters.");
  setStatus("Updating password…");
  try {
    await fetchJson(`${API}/api/auth/change-password`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ old_password, new_password }),
    });
    if (qs("chPwdOld")) qs("chPwdOld").value = "";
    if (qs("chPwdNew")) qs("chPwdNew").value = "";
    if (out) out.textContent = "Password updated.";
    setStatus("Password updated.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Password change failed.");
  }
}

let smtpHasStoredPassword = false;

async function loadSmtpSettings() {
  const out = qs("smtpSettingsResult");
  try {
    const data = await fetchJson(`${API}/api/reports/smtp-settings`);
    const s = data?.settings || {};
    smtpHasStoredPassword = Boolean(s.has_password);
    if (qs("smtpHost")) qs("smtpHost").value = String(s.host || "");
    if (qs("smtpPort")) qs("smtpPort").value = String(Number(s.port || 587));
    if (qs("smtpSecure")) qs("smtpSecure").value = Number(s.secure || 0) === 1 ? "1" : "0";
    if (qs("smtpUsername")) qs("smtpUsername").value = String(s.username || "");
    if (qs("smtpPassword")) qs("smtpPassword").value = "";
    if (qs("smtpFromEmail")) qs("smtpFromEmail").value = String(s.from_email || "");
    if (qs("smtpFromName")) qs("smtpFromName").value = String(s.from_name || "");
    if (out) {
      out.textContent = `Loaded SMTP settings.\nPassword set: ${s.has_password ? "yes" : "no"}\nUpdated by: ${s.updated_by || "-"}\nUpdated at: ${s.updated_at || "-"}`;
    }
    await loadSmtpSubscriptionOptions();
    setStatus("SMTP settings loaded.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("SMTP settings load failed.");
  }
}

async function loadSmtpSubscriptionOptions() {
  const sel = qs("smtpSubscriptionId");
  if (!sel) return;
  const current = String(sel.value || "");
  const data = await fetchJson(`${API}/api/reports/subscriptions`);
  const rows = Array.isArray(data?.subscriptions) ? data.subscriptions : [];
  sel.innerHTML = '<option value="">Select subscription...</option>';
  rows.forEach((r) => {
    const id = Number(r.id || 0);
    if (!id) return;
    const opt = document.createElement("option");
    opt.value = String(id);
    const channel = String(r.channel || "").toLowerCase();
    const active = Number(r.active || 0) === 1 ? "active" : "paused";
    opt.textContent = `${r.name || "Subscription"} (#${id}) - ${channel} - ${active}`;
    sel.appendChild(opt);
  });
  if (current && Array.from(sel.options).some((o) => o.value === current)) sel.value = current;
}

async function saveSmtpSettings() {
  const out = qs("smtpSettingsResult");
  const body = {
    host: String(qs("smtpHost")?.value || "").trim(),
    port: Number(qs("smtpPort")?.value || 587),
    secure: Number(qs("smtpSecure")?.value || 0) === 1 ? 1 : 0,
    username: String(qs("smtpUsername")?.value || "").trim(),
    password: String(qs("smtpPassword")?.value || ""),
    from_email: String(qs("smtpFromEmail")?.value || "").trim(),
    from_name: String(qs("smtpFromName")?.value || "").trim(),
  };
  if (!body.host) return alert("SMTP host is required.");
  if (!body.username) return alert("SMTP username is required.");
  if (!body.from_email) return alert("From email is required.");
  if (!body.password && !smtpHasStoredPassword) {
    return alert("SMTP password is required on first setup. Enter the password, then Save SMTP.");
  }
  setStatus("Saving SMTP settings…");
  try {
    const data = await fetchJson(`${API}/api/reports/smtp-settings`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (qs("smtpPassword")) qs("smtpPassword").value = "";
    smtpHasStoredPassword = Boolean(data?.settings?.has_password);
    if (out) out.textContent = JSON.stringify({ ok: true, settings: data?.settings || {} }, null, 2);
    setStatus("SMTP settings saved.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("SMTP save failed.");
  }
}

async function testSmtpSettings() {
  const out = qs("smtpSettingsResult");
  const to = String(qs("smtpTestTo")?.value || "").trim();
  if (!to) return alert("Enter a test recipient email.");
  setStatus("Sending SMTP test email…");
  try {
    const data = await fetchJson(`${API}/api/reports/smtp-settings/test`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ to }),
    });
    if (out) out.textContent = JSON.stringify(data, null, 2);
    setStatus("SMTP test email sent.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("SMTP test failed.");
  }
}

function renderPushNotifyDevices(devices) {
  const tbody = qs("pushNotifyDevicesTbody");
  if (!tbody) return;
  const rows = Array.isArray(devices) ? devices : [];
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="4" class="muted small">No devices registered yet. Technicians must sign into IRONLOG Notify on their phone.</td></tr>';
    return;
  }
  tbody.innerHTML = rows
    .map((d) => {
      const user = escapeHtml(String(d.username || ""));
      const label = escapeHtml(String(d.device_label || "—"));
      const platform = escapeHtml(String(d.platform || "android"));
      const seen = escapeHtml(String(d.last_seen_at || "—"));
      return `<tr><td>${user}</td><td>${label}</td><td>${platform}</td><td>${seen}</td></tr>`;
    })
    .join("");
}

function renderPushNotifyRecent(recent) {
  const tbody = qs("pushNotifyRecentTbody");
  if (!tbody) return;
  const rows = Array.isArray(recent) ? recent : [];
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="muted small">No notifications sent yet.</td></tr>';
    return;
  }
  tbody.innerHTML = rows
    .map((r) => {
      const when = escapeHtml(String(r.sent_at || "—"));
      const user = escapeHtml(String(r.username || "—"));
      const kindVal = escapeHtml(String(r.kind || "—"));
      const title = escapeHtml(String(r.title || "—"));
      const ok = r.success ? "Yes" : "No";
      const err = r.error ? ` title="${escapeHtml(String(r.error))}"` : "";
      return `<tr><td>${when}</td><td>${user}</td><td>${kindVal}</td><td>${title}</td><td${err}>${ok}</td></tr>`;
    })
    .join("");
}

function populatePushNotifyUserOptions(devices) {
  const sel = qs("pushNotifyUsername");
  if (!sel) return;
  const current = String(sel.value || "");
  const rows = Array.isArray(devices) ? devices : [];
  const seen = new Set();
  sel.innerHTML = '<option value="">Select technician…</option>';
  rows.forEach((d) => {
    const username = String(d.username || "").trim();
    if (!username || seen.has(username.toLowerCase())) return;
    seen.add(username.toLowerCase());
    const opt = document.createElement("option");
    opt.value = username;
    const label = String(d.device_label || "").trim();
    opt.textContent = label ? `${username} (${label})` : username;
    sel.appendChild(opt);
  });
  if (current && Array.from(sel.options).some((o) => o.value === current)) sel.value = current;
}

async function loadPushNotificationSettings() {
  const out = qs("pushNotifyResult");
  const statusEl = qs("pushNotifyStatus");
  try {
    const data = await fetchJson(`${API}/api/notifications/admin`);
    const enabled = Boolean(data.push_enabled);
    if (statusEl) {
      statusEl.textContent = enabled
        ? "Server push is configured and ready."
        : "Push is not configured on the server (Firebase service account missing). Devices can register but alerts will not send.";
      statusEl.className = enabled ? "muted small" : "muted small";
      statusEl.style.color = enabled ? "" : "#b45309";
    }
    const apkLink = qs("pushNotifyApkLink");
    if (apkLink && data.apk_url) apkLink.href = data.apk_url;
    const expoLink = qs("pushNotifyExpoLink");
    if (expoLink && data.expo_install_url) {
      expoLink.href = data.expo_install_url;
      expoLink.textContent = "Open Expo install page";
    }
    populatePushNotifyUserOptions(data.devices);
    renderPushNotifyDevices(data.devices);
    renderPushNotifyRecent(data.recent);
    if (out) {
      out.textContent =
        `Push enabled: ${enabled ? "yes" : "no"}\n` +
        `Registered devices: ${Array.isArray(data.devices) ? data.devices.length : 0}\n` +
        `APK: ${data.apk_url || "(not set)"}`;
    }
    setStatus("Push notification settings loaded.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    if (statusEl) statusEl.textContent = "Could not load push notification settings.";
    setStatus("Push notification load failed.");
  }
}

async function sendPushNotificationTest() {
  const out = qs("pushNotifyResult");
  const username = String(qs("pushNotifyUsername")?.value || "").trim();
  if (!username) return alert("Select a technician with a registered device.");
  setStatus("Sending push test…");
  try {
    const data = await fetchJson(`${API}/api/notifications/test`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username }),
    });
    if (out) out.textContent = JSON.stringify(data, null, 2);
    setStatus(data.sent ? "Push test sent." : "Push test completed (check result — device may be unregistered).");
    await loadPushNotificationSettings();
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Push test failed.");
  }
}

async function sendPushNotificationManual() {
  const out = qs("pushNotifyResult");
  const username = String(qs("pushNotifyUsername")?.value || "").trim();
  const title = String(qs("pushNotifyTitle")?.value || "").trim();
  const body = String(qs("pushNotifyBody")?.value || "").trim();
  const woId = String(qs("pushNotifyWoId")?.value || "").trim();
  if (!username) return alert("Select a technician with a registered device.");
  if (!title) return alert("Enter a notification title.");
  if (!body) return alert("Enter a notification message.");
  setStatus("Sending push notification…");
  try {
    const payload = { username, title, body };
    if (woId) payload.wo_id = woId;
    const data = await fetchJson(`${API}/api/notifications/send`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (out) out.textContent = JSON.stringify(data, null, 2);
    setStatus(data.sent ? "Push notification sent." : "Send completed (check result).");
    await loadPushNotificationSettings();
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Push send failed.");
  }
}

let pdfReportSitesCache = [];

function updatePdfReportLogoPreview(available, customOnly = false) {
  const img = qs("pdfReportLogoPreview");
  const removeBtn = qs("removePdfReportLogoBtn");
  if (!img) return;
  if (available) {
    const headers = new Headers(authHeaders());
    const tok = getAuthToken();
    if (tok) headers.set("Authorization", `Bearer ${tok}`);
    fetch(`${API}/api/reports/pdf-settings/logo?t=${Date.now()}`, { headers })
      .then((res) => {
        if (!res.ok) throw new Error("Logo not found");
        return res.blob();
      })
      .then((blob) => {
        const url = URL.createObjectURL(blob);
        if (img.dataset.blobUrl) URL.revokeObjectURL(img.dataset.blobUrl);
        img.dataset.blobUrl = url;
        img.src = url;
        img.hidden = false;
      })
      .catch(() => {
        img.hidden = true;
        img.removeAttribute("src");
      });
  } else {
    if (img.dataset.blobUrl) {
      URL.revokeObjectURL(img.dataset.blobUrl);
      delete img.dataset.blobUrl;
    }
    img.hidden = true;
    img.removeAttribute("src");
  }
  if (removeBtn) removeBtn.disabled = !customOnly;
}

async function uploadPdfReportLogo() {
  const file = qs("pdfReportLogoFile")?.files?.[0];
  if (!file) return alert("Choose a logo image first (PNG or JPEG).");
  const out = qs("pdfReportSettingsResult");
  const fd = new FormData();
  fd.append("file", file);
  const headers = new Headers(authHeaders());
  headers.delete("Content-Type");
  setStatus("Uploading PDF logo…");
  try {
    const res = await fetch(`${API}/api/reports/pdf-settings/logo`, { method: "POST", headers, body: fd });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || data.message || "Upload failed");
    if (qs("pdfReportLogoFile")) qs("pdfReportLogoFile").value = "";
    updatePdfReportLogoPreview(Boolean(data?.company_logo_available), Boolean(data?.company_logo_custom));
    if (out) out.textContent = data?.message || "PDF logo uploaded.";
    setStatus("PDF logo uploaded.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("PDF logo upload failed.");
  }
}

async function removePdfReportLogo() {
  const out = qs("pdfReportSettingsResult");
  if (!confirm("Remove the uploaded PDF company logo?")) return;
  setStatus("Removing PDF logo…");
  try {
    const data = await fetchJson(`${API}/api/reports/pdf-settings/logo`, { method: "DELETE" });
    updatePdfReportLogoPreview(Boolean(data?.company_logo_available), Boolean(data?.company_logo_custom));
    if (out) out.textContent = data?.message || "PDF logo removed.";
    setStatus("PDF logo removed.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("PDF logo remove failed.");
  }
}

async function loadPdfReportSettings() {
  const out = qs("pdfReportSettingsResult");
  try {
    const data = await fetchJson(`${API}/api/reports/pdf-settings`);
    pdfReportSitesCache = Array.isArray(data?.sites) ? data.sites : [];
    const companies = Array.isArray(data?.companies) ? data.companies : [];
    const companySel = qs("pdfReportCompanyCode");
    if (companySel) {
      const currentCompany = String(data?.company_code || "");
      companySel.innerHTML = `<option value="">Custom / manual company name</option>${companies.map((c) =>
        `<option value="${escapeHtml(String(c.company_code || ""))}">${escapeHtml(String(c.company_name || c.company_code || ""))}</option>`
      ).join("")}`;
      companySel.value = currentCompany && Array.from(companySel.options).some((o) => o.value === currentCompany)
        ? currentCompany
        : "";
    }
    if (qs("pdfReportCompanyName")) qs("pdfReportCompanyName").value = String(data?.company_name || "");
    const siteSel = qs("pdfReportSiteCode");
    if (siteSel) {
      const currentSite = String(data?.site_code || "");
      const companyFilter = String(data?.company_code || companySel?.value || "").trim();
      const sites = companyFilter
        ? pdfReportSitesCache.filter((s) => String(s.company_code || "") === companyFilter)
        : pdfReportSitesCache;
      siteSel.innerHTML = `<option value="">Custom / manual site name</option>${sites.map((s) =>
        `<option value="${escapeHtml(String(s.site_code || ""))}">${escapeHtml(String(s.site_name || s.site_code || ""))}</option>`
      ).join("")}`;
      siteSel.value = currentSite && Array.from(siteSel.options).some((o) => o.value === currentSite) ? currentSite : "";
    }
    if (qs("pdfReportSiteName")) qs("pdfReportSiteName").value = String(data?.site_name || "");
    updatePdfReportLogoPreview(
      Boolean(data?.company_logo_available),
      Boolean(data?.company_logo_custom)
    );
    const parts = [];
    if (data?.company_name) parts.push(`Company: ${data.company_name}`);
    if (data?.site_name) parts.push(`Site: ${data.site_name}`);
    if (data?.company_logo_custom) parts.push("Custom logo: yes");
    else if (data?.company_logo_available) parts.push("Logo: default");
    if (out) out.textContent = parts.length ? `Current PDF header — ${parts.join(" · ")}` : "No PDF branding configured yet.";
    setStatus("PDF report branding loaded.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("PDF report branding load failed.");
  }
}

async function savePdfReportSettings() {
  const out = qs("pdfReportSettingsResult");
  const company_code = String(qs("pdfReportCompanyCode")?.value || "").trim();
  let company_name = String(qs("pdfReportCompanyName")?.value || "").trim();
  const site_code = String(qs("pdfReportSiteCode")?.value || "").trim();
  let site_name = String(qs("pdfReportSiteName")?.value || "").trim();
  if (company_code && !company_name) {
    const opt = qs("pdfReportCompanyCode")?.selectedOptions?.[0];
    company_name = String(opt?.textContent || company_code).trim();
  }
  if (site_code && !site_name) {
    const opt = qs("pdfReportSiteCode")?.selectedOptions?.[0];
    site_name = String(opt?.textContent || site_code).trim();
  }
  if (!company_name && !company_code && !site_name && !site_code) {
    return alert("Enter a company name and/or site name.");
  }
  setStatus("Saving PDF report branding…");
  try {
    const data = await fetchJson(`${API}/api/reports/pdf-settings`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ company_code, company_name, site_code, site_name }),
    });
    if (qs("pdfReportCompanyName")) qs("pdfReportCompanyName").value = String(data?.company_name || company_name);
    if (qs("pdfReportSiteName")) qs("pdfReportSiteName").value = String(data?.site_name || site_name);
    const parts = [];
    if (data?.company_name || company_name) parts.push(data?.company_name || company_name);
    if (data?.site_name || site_name) parts.push(`Site: ${data?.site_name || site_name}`);
    if (out) out.textContent = parts.length ? `Saved. PDF header will show: ${parts.join(" · ")}` : "Saved.";
    setStatus("PDF report branding saved.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("PDF report branding save failed.");
  }
}

function onPdfReportCompanyCodeChange() {
  const code = String(qs("pdfReportCompanyCode")?.value || "").trim();
  if (code) {
    const opt = qs("pdfReportCompanyCode")?.selectedOptions?.[0];
    const name = String(opt?.textContent || code).trim();
    if (qs("pdfReportCompanyName")) qs("pdfReportCompanyName").value = name;
  }
  const siteSel = qs("pdfReportSiteCode");
  if (!siteSel) return;
  const currentSite = String(siteSel.value || "");
  const sites = code
    ? pdfReportSitesCache.filter((s) => String(s.company_code || "") === code)
    : pdfReportSitesCache;
  siteSel.innerHTML = `<option value="">Custom / manual site name</option>${sites.map((s) =>
    `<option value="${escapeHtml(String(s.site_code || ""))}">${escapeHtml(String(s.site_name || s.site_code || ""))}</option>`
  ).join("")}`;
  if (currentSite && Array.from(siteSel.options).some((o) => o.value === currentSite)) {
    siteSel.value = currentSite;
  } else {
    siteSel.value = "";
    if (qs("pdfReportSiteName")) qs("pdfReportSiteName").value = "";
  }
}

async function onPdfReportSiteCodeChange() {
  const code = String(qs("pdfReportSiteCode")?.value || "").trim();
  if (!code) return;
  const opt = qs("pdfReportSiteCode")?.selectedOptions?.[0];
  const name = String(opt?.textContent || code).trim();
  if (qs("pdfReportSiteName")) qs("pdfReportSiteName").value = name;
}

async function sendSubscriptionNowFromAdmin() {
  const out = qs("smtpSettingsResult");
  const id = Number(qs("smtpSubscriptionId")?.value || 0);
  if (!id) return alert("Select a subscription first.");
  setStatus("Sending subscription now...");
  try {
    const data = await fetchJson(`${API}/api/reports/subscriptions/${id}/send-now`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    if (out) out.textContent = JSON.stringify(data, null, 2);
    setStatus("Subscription sent.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Subscription send failed.");
  }
}

function backupOptionLabel(f) {
  const mb = Number(f?.bytes || 0) / (1024 * 1024);
  const size = Number.isFinite(mb) ? `${mb.toFixed(1)} MB` : "-";
  return `${String(f?.name || "-")} (${size})`;
}

async function loadBackupFiles() {
  const out = qs("backupRestoreResult");
  const sel = qs("backupFileSelect");
  if (!sel) return;
  const current = String(sel.value || "");
  const data = await fetchJson(`${API}/api/admin/backups/list`);
  const files = Array.isArray(data?.files) ? data.files : [];
  const executeEnabled = Boolean(data?.execute_restore_enabled);
  const restartCmdSet = Boolean(data?.restart_command_set);
  sel.innerHTML = files.length
    ? `<option value="">Select backup file...</option>${files
        .map((f) => `<option value="${esc(String(f.name || ""))}">${esc(backupOptionLabel(f))}</option>`)
        .join("")}`
    : `<option value="">No backups found</option>`;
  if (current && Array.from(sel.options).some((o) => o.value === current)) sel.value = current;
  if (out) {
    out.textContent =
      `Backup dir: ${data?.backup_dir || "-"}\n` +
      `DB path: ${data?.db_path || "-"}\n` +
      `Backups found: ${files.length}\n` +
      `Execute restore enabled: ${executeEnabled ? "yes" : "no"}\n` +
      `Restart command set: ${restartCmdSet ? "yes" : "no"}`;
  }
  const execBtn = qs("executeBackupRestoreBtn");
  if (execBtn) execBtn.disabled = !(executeEnabled && restartCmdSet);
}

async function createBackupNow() {
  const out = qs("backupRestoreResult");
  setStatus("Creating manual backup...");
  const data = await fetchJson(`${API}/api/admin/backups/create`, {
    method: "POST",
    body: JSON.stringify({}),
  });
  if (out) out.textContent = JSON.stringify(data, null, 2);
  await loadBackupFiles();
  setStatus("Backup created.");
}

async function previewBackupRestore() {
  const out = qs("backupRestoreResult");
  const backup_name = String(qs("backupFileSelect")?.value || "").trim();
  if (!backup_name) return alert("Select a backup file first.");
  setStatus("Loading restore preview...");
  const data = await fetchJson(`${API}/api/admin/backups/restore/preview`, {
    method: "POST",
    body: JSON.stringify({ backup_name }),
  });
  if (out) out.textContent = JSON.stringify(data, null, 2);
  setStatus("Restore preview ready.");
}

async function stageBackupRestore() {
  const out = qs("backupRestoreResult");
  const backup_name = String(qs("backupFileSelect")?.value || "").trim();
  const notes = String(qs("backupRestoreNotes")?.value || "").trim();
  if (!backup_name) return alert("Select a backup file first.");
  const confirmed = window.confirm(
    `Stage restore plan for ${backup_name}?\n\nThis does not auto-copy yet; it prepares safe restore steps.`
  );
  if (!confirmed) return;
  setStatus("Staging restore plan...");
  const data = await fetchJson(`${API}/api/admin/backups/restore/apply`, {
    method: "POST",
    body: JSON.stringify({ backup_name, confirm_text: "RESTORE", notes }),
  });
  if (out) out.textContent = JSON.stringify(data, null, 2);
  await loadBackupFiles();
  setStatus("Restore plan staged.");
}

async function executeBackupRestoreNow() {
  const out = qs("backupRestoreResult");
  const backup_name = String(qs("backupFileSelect")?.value || "").trim();
  const notes = String(qs("backupRestoreNotes")?.value || "").trim();
  if (!backup_name) return alert("Select a backup file first.");
  const confirmed = window.confirm(
    `Execute restore now for ${backup_name}?\n\nThis will restart the API process immediately.`
  );
  if (!confirmed) return;
  setStatus("Executing restore + restart...");
  const data = await fetchJson(`${API}/api/admin/backups/restore/execute`, {
    method: "POST",
    body: JSON.stringify({ backup_name, confirm_text: "RESTORE_NOW", notes }),
  });
  if (out) out.textContent = JSON.stringify(data, null, 2);
  setStatus("Restore execute requested. Reconnect after restart.");
}

/** Daily checklists — LDV + machine pre-start + safety (QR mirror) */
let clHubData = null;
let clPoisonedBaselineNote = "";
let clSafetyHubData = null;
let clSelectedAssetCode = "";
let clSelectedKind = "";
let clSelectedProfileId = "";
let clCurrentCheckId = 0;
let clPreviousKm = null;
let clPreviousSmu = null;
let clPendingAssetCode = String(
  new URLSearchParams(window.location.search).get("asset_code") || ""
)
  .trim()
  .toUpperCase();
let clPendingSafetyItemCode = String(
  new URLSearchParams(window.location.search).get("item_code") || ""
)
  .trim()
  .toUpperCase();

const CL_LDV_ITEMS = [
  { key: "brakes_ok", label: "Brakes OK" },
  { key: "lights_ok", label: "Lights OK" },
  { key: "tyres_ok", label: "Tyres OK" },
  { key: "oil_coolant_ok", label: "Oil/Coolant OK" },
  { key: "leaks_damage_ok", label: "No leaks or visible damage" },
  { key: "safety_items_ok", label: "Safety items in place" },
];

const CL_PT_HELP_KEY = "ironlog_cl_pt_help";

/** English → Portuguese (Mozambique field Portuguese) for checklist labels & common notes. */
const CL_PT_GLOSSARY = {
  "Pre-start checks (all required)": "Verificações pré-arranque (todas obrigatórias)",
  "Brakes OK": "Travões OK",
  "Lights OK": "Luzes OK",
  "Tyres OK": "Pneus OK",
  "Oil/Coolant OK": "Óleo / líquido de arrefecimento OK",
  "No leaks or visible damage": "Sem fugas ou danos visíveis",
  "Safety items in place": "Equipamento de segurança no lugar",
  "Fluid levels": "Níveis de fluidos",
  "Undercarriage": "Trem de rodagem / esteiras",
  "Safety & cab": "Segurança e cabine",
  "Safety": "Segurança",
  "Machine & tyres": "Máquina e pneus",
  "Body & running gear": "Caixa basculante e rodagem",
  "Fluids & fuel systems": "Fluidos e sistemas de combustível",
  "Dispensing equipment": "Equipamento de distribuição",
  "Vehicle": "Veículo",
  "Moldboard & tyres": "Lâmina e pneus",
  "Engine oil level OK": "Nível de óleo do motor OK",
  "Engine oil OK": "Óleo do motor OK",
  "Coolant level OK": "Nível de líquido de arrefecimento OK",
  "Coolant OK": "Líquido de arrefecimento OK",
  "Hydraulic oil level OK": "Nível de óleo hidráulico OK",
  "Hydraulic oil OK": "Óleo hidráulico OK",
  "Hydraulic / transmission oil OK": "Óleo hidráulico / transmissão OK",
  "Swing / slew gear oil OK (if applicable)": "Óleo da giratória OK (se aplicável)",
  "Final drives / track gear oil OK": "Redutores finais / óleo das esteiras OK",
  "Final drives OK": "Redutores finais OK",
  "Track tension OK": "Tensão das esteiras OK",
  "Rollers & idlers OK (no seized / flat spots)": "Rolos e rodas guia OK (sem bloqueios / zonas planas)",
  "Sprocket teeth OK": "Dentes da coroa OK",
  "No abnormal cuts, cracks, or leaks on tracks": "Sem cortes, fissuras ou fugas anormais nas esteiras",
  "Fire extinguisher present & charged": "Extintor presente e carregado",
  "Fire extinguisher OK": "Extintor OK",
  "Seat belt OK": "Cinto de segurança OK",
  "Mirrors / cameras clean & working": "Espelhos / câmaras limpos e a funcionar",
  "Horn & emergency stop OK": "Buzina e paragem de emergência OK",
  "Horn & E-stop OK": "Buzina e paragem de emergência OK",
  "Windows / guards intact": "Janelas / proteções intactas",
  "Blade / ripper pins & hydraulics OK": "Pinos da lâmina / ripper e hidráulica OK",
  "Rollers & sprocket OK": "Rolos e coroa OK",
  "Transmission / axle oils OK": "Óleos de transmissão / eixos OK",
  "Brake fluid OK": "Líquido de travões OK",
  "Tyres — pressure & damage OK": "Pneus — pressão e danos OK",
  "Tyres OK": "Pneus OK",
  "Centre articulation / pins OK": "Articulação central / pinos OK",
  "Bucket & linkage OK": "Caçamba e linkage OK",
  "Lights & beacon OK": "Luzes e baliza OK",
  "Horn & reversing alarm OK": "Buzina e alarme de marcha-atrás OK",
  "Body / hoist & tail door OK": "Caixa, elevador e porta traseira OK",
  "Steering free play OK": "Folga da direção OK",
  "Product tank level OK": "Nível do tanque de produto OK",
  "No fuel, oil, or hydraulic leaks": "Sem fugas de combustível, óleo ou hidráulica",
  "Hoses, reels & nozzles OK (no damage/leaks)": "Mangueiras, carretéis e bicos OK (sem danos/fugas)",
  "Meters / pumps OK": "Medidores / bombas OK",
  "Bonding & earthing OK": "Ligação equipotencial e aterramento OK",
  "Spill kit present & accessible": "Kit de derrame presente e acessível",
  "Brakes & steering OK": "Travões e direção OK",
  "Hazchem / no smoking signage OK": "Sinalização Hazchem / proibido fumar OK",
  "Circle drive oil OK (if applicable)": "Óleo do círculo OK (se aplicável)",
  "Moldboard & linkage OK": "Lâmina e linkage OK",
  "Horn OK": "Buzina OK",
  "leak": "fuga",
  "leaks": "fugas",
  "damage": "dano",
  "broken": "partido",
  "low oil": "óleo baixo",
  "low coolant": "líquido de arrefecimento baixo",
  "not working": "não funciona",
  "needs attention": "precisa de atenção",
  "flat tyre": "pneu furado",
  "warning light": "luz de aviso",
};

let clPtGlossaryReverse = null;

function isClPtHelpOn() {
  if (getLang() === "pt") return true;
  return localStorage.getItem(CL_PT_HELP_KEY) === "1";
}

function setClPtHelp(on) {
  localStorage.setItem(CL_PT_HELP_KEY, on ? "1" : "0");
}

function clBuildPtGlossaryReverse() {
  if (clPtGlossaryReverse) return clPtGlossaryReverse;
  clPtGlossaryReverse = {};
  for (const [en, pt] of Object.entries(CL_PT_GLOSSARY)) {
    if (pt && !clPtGlossaryReverse[pt.toLowerCase()]) {
      clPtGlossaryReverse[pt.toLowerCase()] = en;
    }
  }
  return clPtGlossaryReverse;
}

function clEnToPt(text) {
  const src = String(text || "").trim();
  if (!src) return "";
  if (CL_PT_GLOSSARY[src]) return CL_PT_GLOSSARY[src];
  let out = src;
  const phrases = Object.keys(CL_PT_GLOSSARY).sort((a, b) => b.length - a.length);
  for (const en of phrases) {
    if (en.length < 4) continue;
    const re = new RegExp(en.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
    out = out.replace(re, CL_PT_GLOSSARY[en]);
  }
  return out;
}

function clPtToEn(text) {
  const src = String(text || "").trim();
  if (!src) return "";
  const reverse = clBuildPtGlossaryReverse();
  if (reverse[src.toLowerCase()]) return reverse[src.toLowerCase()];
  let out = src;
  const phrases = Object.entries(reverse).sort((a, b) => b[0].length - a[0].length);
  for (const [pt, en] of phrases) {
    if (pt.length < 4) continue;
    const re = new RegExp(pt.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
    out = out.replace(re, en);
  }
  return out;
}

function clTranslateFreeText(text, direction = "pt-en") {
  const raw = String(text || "").trim();
  if (!raw) return "";
  return direction === "en-pt" ? clEnToPt(raw) : clPtToEn(raw);
}

function clLabelHtml(englishLabel) {
  const en = String(englishLabel || "");
  if (!isClPtHelpOn()) return escapeHtml(en);
  const pt = clEnToPt(en);
  if (!pt || pt.toLowerCase() === en.toLowerCase()) return escapeHtml(en);
  return `${escapeHtml(en)}<small class="cl-label-pt">${escapeHtml(pt)}</small>`;
}

function clSectionTitleHtml(title) {
  return clLabelHtml(title);
}

async function refreshClChecklistLabels() {
  if (!clSelectedAssetCode) return;
  await selectChecklistAsset(clSelectedAssetCode).catch(() => {});
}

function runClPtTranslate() {
  const out = qs("clPtTranslateOut");
  const input = String(qs("clPtTranslateIn")?.value || "").trim();
  const dir = String(qs("clPtTranslateDir")?.value || "pt-en");
  if (!input) {
    if (out) out.textContent = "Enter text to translate.";
    return;
  }
  const translated = clTranslateFreeText(input, dir);
  if (out) {
    out.textContent = translated || "(No translation found — try shorter phrases or checklist terms.)";
  }
}

function clCheckDate() {
  return qs("clCheckDate")?.value || todayLocalYmd();
}

function clSafeDomId(key) {
  return `cl_chk_${String(key || "").replace(/[^a-zA-Z0-9_]/g, "_")}`;
}

function clSetFormMsg(text, ok) {
  const el = qs("clFormMsg");
  if (!el) return;
  el.textContent = String(text || "");
  el.style.color = ok === true ? "#15803d" : ok === false ? "#b91c1c" : "";
}

function clShowSyncState(sync) {
  const el = qs("clSyncState");
  if (!el) return;
  if (!sync) {
    el.classList.add("hidden");
    el.textContent = "";
    return;
  }
  if (sync.skipped && sync.reason === "unusual_km") {
    el.textContent = String(sync.message || "Pre-start saved. Daily input not updated — unusual KM.");
    el.classList.remove("hidden");
    return;
  }
  if (sync.synced !== true) {
    el.classList.add("hidden");
    el.textContent = "";
    return;
  }
  const mode = String(sync.mode || "updated");
  const action = mode === "inserted" ? "created" : "updated";
  const workDate = String(sync.work_date || "");
  if (sync.unit === "hours" || sync.opening_hours != null || sync.closing_hours != null) {
    el.textContent =
      `Synced to Daily Input (${action}) — ${workDate}: ` +
      `open ${Number(sync.opening_hours || 0).toFixed(1)} hrs, ` +
      `close ${Number(sync.closing_hours || 0).toFixed(1)} hrs, ` +
      `run ${Number(sync.run_hours || 0).toFixed(1)} hrs.`;
  } else {
    el.textContent =
      `Synced to Daily Input (${action}) — ${workDate}: ` +
      `open ${Number(sync.opening_km || 0).toFixed(1)} km, ` +
      `close ${Number(sync.closing_km || 0).toFixed(1)} km, ` +
      `run ${Number(sync.run_km || 0).toFixed(1)} km.`;
  }
  el.classList.remove("hidden");
}

function renderClHubSummary(data) {
  const el = qs("clHubSummary");
  if (!el) return;
  const s = data?.summary || {};
  const ldvDone = Number(s.ldv_compliant || 0);
  const ldvTotal = Number(s.ldv_total || 0);
  const macDone = Number(s.machine_compliant || 0);
  const macTotal = Number(s.machine_total || 0);
  const commentCount = Number(data?.comments?.length || 0) + Number(clSafetyHubData?.comments?.length || 0);
  el.innerHTML = `
    <span class="pill blue">LDV: ${ldvDone}/${ldvTotal}</span>
    <span class="pill blue">Machines: ${macDone}/${macTotal}</span>
    ${clSafetyHubData?.summary ? (() => {
      const s = clSafetyHubData.summary;
      const done = Number(s.completed ?? s.compliant ?? 0);
      const total = Number(s.total || 0);
      const flagged = Number(s.flagged || 0);
      const flaggedNote = flagged ? ` · ${flagged} flagged` : "";
      return `<span class="pill blue">Safety: ${done}/${total} done${flaggedNote}</span>`;
    })() : ""}
    ${commentCount ? `<span class="pill amber">${commentCount} comment${commentCount === 1 ? "" : "s"}</span>` : ""}
    <button type="button" id="clPtHelpPill" class="pill pill-btn${isClPtHelpOn() ? " active" : ""}" title="Show Portuguese under checklist lines">PT ⇄ EN${isClPtHelpOn() ? " ON" : ""}</button>
    <span class="pill">${escapeHtml(String(data?.check_date || clCheckDate()))}</span>
  `;
  qs("clPtHelpPill")?.addEventListener("click", () => {
    setClPtHelp(!isClPtHelpOn());
    renderClHubSummary(data);
    refreshClChecklistLabels().catch(() => {});
  });
}

function clCommentKindLabel(kind) {
  if (kind === "machine") return "Machine";
  if (kind === "safety") return "Safety";
  return "LDV";
}

function renderClDayComments() {
  const host = qs("clDayComments");
  if (!host) return;

  const rows = [
    ...(Array.isArray(clHubData?.comments) ? clHubData.comments : []),
    ...(Array.isArray(clSafetyHubData?.comments) ? clSafetyHubData.comments : []),
  ];

  if (!rows.length) {
    host.classList.add("hidden");
    host.innerHTML = "";
    return;
  }

  host.classList.remove("hidden");
  host.innerHTML = `
    <div class="checklist-day-comments-head">
      <h4>Comments for ${escapeHtml(String(clHubData?.check_date || clCheckDate()))}</h4>
      <span class="muted small">${rows.length} note${rows.length === 1 ? "" : "s"}</span>
    </div>
    <div class="checklist-day-comments-list"></div>
  `;

  const list = host.querySelector(".checklist-day-comments-list");
  rows.forEach((row) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "checklist-day-comment";
    const code = row.kind === "safety"
      ? String(row.item_code || "")
      : String(row.asset_code || "");
    const name = row.kind === "safety"
      ? String(row.item_name || row.location || "")
      : String(row.asset_name || "");
    const inspector = String(row.inspector_name || "").trim();
    const when = String(row.updated_at || "").replace("T", " ").slice(0, 16);
    btn.innerHTML = `
      <div class="checklist-day-comment-meta">
        <span class="pill">${escapeHtml(clCommentKindLabel(row.kind))}</span>
        <strong>${escapeHtml(code || "—")}</strong>
        ${name ? `<span class="muted small">${escapeHtml(name)}</span>` : ""}
        ${inspector ? `<span class="muted small">· ${escapeHtml(inspector)}</span>` : ""}
        ${when ? `<span class="muted small">· ${escapeHtml(when)}</span>` : ""}
      </div>
      <p class="checklist-day-comment-text">${escapeHtml(String(row.notes || ""))}</p>
    `;
    btn.addEventListener("click", () => {
      if (row.kind === "safety") {
        const itemCode = String(row.item_code || "").trim();
        if (itemCode) {
          window.location.href = `./safety-inspection.html?item_code=${encodeURIComponent(itemCode)}`;
        }
        return;
      }
      selectChecklistAsset(String(row.asset_code || "")).catch((e) => {
        setStatus("Checklist open error: " + (e.message || e));
      });
    });
    list.appendChild(btn);
  });
}

function renderClAssetChip(asset, selectedCode) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "checklist-asset-chip";
  if (asset.status === "compliant") btn.classList.add("compliant");
  if (asset.asset_code === selectedCode) btn.classList.add("selected");
  btn.dataset.assetCode = asset.asset_code;
  btn.dataset.kind = asset.kind || "";
  if (asset.profile_id) btn.dataset.profileId = asset.profile_id;
  const pill =
    asset.status === "compliant"
      ? "<span class='pill green'>DONE</span>"
      : "<span class='pill orange'>PENDING</span>";
  btn.innerHTML = `
    <div class="chip-code">${escapeHtml(asset.asset_code)}</div>
    <div class="chip-status">${pill}</div>
    <div class="chip-name">${escapeHtml(asset.asset_name || "")}</div>
  `;
  return btn;
}

function renderClHubSections(data) {
  const host = qs("clHubSections");
  if (!host) return;
  host.innerHTML = "";
  const selected = clSelectedAssetCode;

  const ldv = Array.isArray(data?.ldv) ? data.ldv : [];
  if (ldv.length) {
    const sec = document.createElement("div");
    sec.className = "checklist-hub-section";
    sec.innerHTML = `<h4>LDV Pre-Start (V01–V15)</h4>`;
    const grid = document.createElement("div");
    grid.className = "checklist-asset-grid";
    ldv.forEach((a) => grid.appendChild(renderClAssetChip(a, selected)));
    sec.appendChild(grid);
    host.appendChild(sec);
  }

  const groups = Array.isArray(data?.machine_groups) ? data.machine_groups : [];
  groups.forEach((g) => {
    if (!g.assets?.length) return;
    const sec = document.createElement("div");
    sec.className = "checklist-hub-section";
    sec.innerHTML = `<h4>${escapeHtml(g.title || g.profile_id || "Machine")}</h4>`;
    const grid = document.createElement("div");
    grid.className = "checklist-asset-grid";
    g.assets.forEach((a) => grid.appendChild(renderClAssetChip(a, selected)));
    sec.appendChild(grid);
    host.appendChild(sec);
  });

  const safetyItems = Array.isArray(clSafetyHubData?.items) ? clSafetyHubData.items : [];
  if (safetyItems.length) {
    const sec = document.createElement("div");
    sec.className = "checklist-hub-section";
    sec.innerHTML = `<h4>Safety equipment</h4>`;
    const grid = document.createElement("div");
    grid.className = "checklist-asset-grid";
    safetyItems.forEach((it) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "checklist-asset-chip checklist-safety-chip";
      const st = String(it.status || "pending").toLowerCase();
      if (st === "compliant" || st === "done") btn.classList.add("compliant");
      if (st === "flagged") btn.classList.add("flagged");
      if (st === "attention") btn.classList.add("attention");
      let pill = "<span class='pill orange'>PENDING</span>";
      if (st === "compliant" || st === "done") {
        pill = "<span class='pill green'>DONE</span>";
      } else if (st === "flagged") {
        pill = "<span class='pill pill-red'>FLAGGED</span>";
      } else if (st === "attention") {
        pill = "<span class='pill amber'>ATTENTION</span>";
      }
      btn.dataset.itemCode = String(it.item_code || "");
      btn.innerHTML = `
        <div class="chip-code">${escapeHtml(it.item_code)}</div>
        <div class="chip-status">${pill}</div>
        <div class="chip-name">${escapeHtml(it.item_name || it.template_title || "")}${it.location ? ` · ${escapeHtml(it.location)}` : ""}</div>
      `;
      grid.appendChild(btn);
    });
    sec.appendChild(grid);
    host.appendChild(sec);
  }

  if (!ldv.length && !groups.length && !safetyItems.length) {
    host.innerHTML = `<div class="muted small">No checklist assets configured. Add LDV codes, machine categories, or safety equipment in User Admin.</div>`;
  }
}

async function loadChecklistHub() {
  const host = qs("clHubSections");
  if (host && !clHubData) host.innerHTML = `<div class="muted small">Loading checklists…</div>`;
  const date = clCheckDate();
  try {
    const [data, safety] = await Promise.all([
      fetchJson(`${API}/api/maintenance/checklist-hub?date=${encodeURIComponent(date)}`),
      fetchJson(`${API}/api/safety/hub?date=${encodeURIComponent(date)}`).catch(() => null),
    ]);
    clHubData = data;
    clSafetyHubData = safety;
    renderClHubSummary(data);
    renderClDayComments();
    renderClHubSections(data);
    if (clPendingSafetyItemCode) {
      const code = clPendingSafetyItemCode;
      clPendingSafetyItemCode = "";
      window.location.href = `./safety-inspection.html?item_code=${encodeURIComponent(code)}`;
      return;
    }
    if (clPendingAssetCode) {
      const code = clPendingAssetCode;
      clPendingAssetCode = "";
      await selectChecklistAsset(code).catch(() => {});
    }
  } catch (e) {
    if (host) host.innerHTML = `<div class="muted small">Checklist load error: ${escapeHtml(e.message || e)}</div>`;
    setStatus("Checklist load error: " + (e.message || e));
  }
}

function findClHubAsset(assetCode) {
  const code = String(assetCode || "").trim().toUpperCase();
  if (!code || !clHubData) return null;
  const ldv = (clHubData.ldv || []).find((a) => a.asset_code === code);
  if (ldv) return ldv;
  for (const g of clHubData.machine_groups || []) {
    const hit = (g.assets || []).find((a) => a.asset_code === code);
    if (hit) return hit;
  }
  return null;
}

function renderClLdvChecklist(checklist) {
  const root = qs("clChecklistRoot");
  if (!root) return;
  const byKey = {};
  (Array.isArray(checklist) ? checklist : []).forEach((c) => {
    byKey[String(c.key)] = Boolean(c.ok);
  });
  root.innerHTML = `<div class="cl-sec"><div class="cl-sec-title">${clSectionTitleHtml("Pre-start checks (all required)")}</div>`;
  const sec = root.querySelector(".cl-sec");
  CL_LDV_ITEMS.forEach((it) => {
    const id = clSafeDomId(it.key);
    const row = document.createElement("div");
    row.className = "cl-check-row";
    row.innerHTML = `<input type="checkbox" id="${id}" data-key="${escapeHtml(it.key)}" ${byKey[it.key] ? "checked" : ""} /><label for="${id}">${clLabelHtml(it.label)}</label>`;
    sec.appendChild(row);
  });
}

function renderClMachineChecklist(template, checklist) {
  const root = qs("clChecklistRoot");
  if (!root) return;
  root.innerHTML = "";
  const byKey = {};
  (Array.isArray(checklist) ? checklist : []).forEach((c) => {
    byKey[String(c.key)] = Boolean(c.ok);
  });
  for (const sec of template?.sections || []) {
    const wrap = document.createElement("div");
    wrap.className = "cl-sec";
    const title = document.createElement("div");
    title.className = "cl-sec-title";
    title.innerHTML = clSectionTitleHtml(String(sec.title || ""));
    wrap.appendChild(title);
    for (const it of sec.items || []) {
      const key = String(it.key || "").trim();
      if (!key) continue;
      const id = clSafeDomId(key);
      const row = document.createElement("div");
      row.className = "cl-check-row";
      row.innerHTML = `<input type="checkbox" id="${id}" data-key="${escapeHtml(key)}" ${byKey[key] ? "checked" : ""} /><label for="${id}">${clLabelHtml(it.label || key)}</label>`;
      wrap.appendChild(row);
    }
    root.appendChild(wrap);
  }
}

function readClChecklistObject() {
  const out = {};
  qs("clChecklistRoot")?.querySelectorAll("input[type=checkbox][data-key]").forEach((el) => {
    const k = String(el.dataset.key || "").trim();
    if (!k) return;
    out[k] = Boolean(el.checked);
  });
  return out;
}

async function selectChecklistAsset(assetCode) {
  const code = String(assetCode || "").trim().toUpperCase();
  if (!code) return;
  const hubAsset = findClHubAsset(code);
  if (!hubAsset) {
    clSetFormMsg(`Asset ${code} is not in today's checklist hub. Refresh or check asset category.`, false);
    return;
  }
  clSelectedAssetCode = code;
  clSelectedKind = hubAsset.kind === "machine" ? "machine" : "ldv";
  clSelectedProfileId = String(hubAsset.profile_id || "");
  renderClHubSections(clHubData);
  qs("clFormPanel")?.classList.remove("hidden");
  clSetFormMsg("Loading checklist…", null);
  clShowSyncState(null);

  if (clSelectedKind === "ldv") {
    qs("clLdvFields")?.classList.remove("hidden");
    qs("clMachineFields")?.classList.add("hidden");
    updateClLdvSupervisorPanel();
    const data = await fetchJson(
      `${API}/api/maintenance/vehicle-ldv-checks/prestart-context?asset_code=${encodeURIComponent(code)}&check_date=${encodeURIComponent(clCheckDate())}`
    );
    const asset = data?.asset || {};
    clPreviousKm = data?.previous_odometer_km == null ? null : Number(data.previous_odometer_km);
    if (qs("clFormTitle")) qs("clFormTitle").textContent = `LDV Pre-Start — ${code}`;
    if (qs("clFormSubtitle")) qs("clFormSubtitle").textContent = String(asset.asset_name || "");
    if (qs("clFormMeta")) {
      qs("clFormMeta").innerHTML = `<span class="pill blue">${escapeHtml(code)}</span><span class="pill">${escapeHtml(clCheckDate())}</span>`;
    }
    if (qs("clPrevKm")) {
      qs("clPrevKm").textContent = clPreviousKm == null ? "—" : `${clPreviousKm.toFixed(1)} km`;
    }
    if (qs("clCorrectOpeningKm")) {
      qs("clCorrectOpeningKm").value = "";
      qs("clCorrectOpeningKm").placeholder = "Leave blank — auto from last good reading";
    }
    if (data?.baseline_poisoned && Number(data?.raw_previous_odometer_km) > 0) {
      clPoisonedBaselineNote = `Bad prior KM (${Number(data.raw_previous_odometer_km).toFixed(1)}) ignored — use Supervisor correction if today's reading is wrong.`;
    } else {
      clPoisonedBaselineNote = "";
    }
    const existing = data?.existing_prestart || null;
    clCurrentCheckId = Number(existing?.id || 0);
    if (qs("clOdometer")) qs("clOdometer").value = existing?.odometer_km != null ? String(existing.odometer_km) : "";
    if (qs("clInspector")) qs("clInspector").value = existing?.inspector_name || getSessionUser() || "";
    if (qs("clNotes")) qs("clNotes").value = existing?.notes || "";
    renderClLdvChecklist(existing?.checklist || []);
    const baseMsg = existing ? "Pre-start exists for this date — update if needed." : "";
    const msg = [clPoisonedBaselineNote, baseMsg].filter(Boolean).join(" ");
    clSetFormMsg(msg, clPoisonedBaselineNote ? false : existing ? true : null);
  } else {
    qs("clLdvFields")?.classList.add("hidden");
    qs("clMachineFields")?.classList.remove("hidden");
    updateClLdvSupervisorPanel();
    updateClMachineSupervisorPanel();
    const data = await fetchJson(
      `${API}/api/maintenance/machine-prestart/context?asset_code=${encodeURIComponent(code)}&check_date=${encodeURIComponent(clCheckDate())}`
    );
    const asset = data?.asset || {};
    const template = data?.template || {};
    clPreviousSmu = data?.previous_smu_hours == null ? null : Number(data.previous_smu_hours);
    if (qs("clFormTitle")) qs("clFormTitle").textContent = String(template.title || "Machine pre-start");
    if (qs("clFormSubtitle")) qs("clFormSubtitle").textContent = `${code} — ${asset.asset_name || ""}`;
    if (qs("clFormMeta")) {
      qs("clFormMeta").innerHTML = `<span class="pill blue">${escapeHtml(code)}</span><span class="pill">${escapeHtml(String(data?.profile_id || ""))}</span><span class="pill">${escapeHtml(clCheckDate())}</span>`;
    }
    if (qs("clPrevSmu")) {
      qs("clPrevSmu").textContent = clPreviousSmu == null ? "—" : `${clPreviousSmu.toFixed(1)} hrs`;
    }
    const existing = data?.existing_check || null;
    clCurrentCheckId = Number(existing?.id || 0);
    if (qs("clSmuHours")) qs("clSmuHours").value = existing?.smu_hours != null ? String(existing.smu_hours) : "";
    if (qs("clInspector")) qs("clInspector").value = existing?.inspector_name || getSessionUser() || "";
    if (qs("clNotes")) qs("clNotes").value = existing?.notes || "";
    renderClMachineChecklist(template, existing?.checklist || []);
    clSetFormMsg(existing ? "Pre-start exists for this date — update if needed." : "", existing ? true : null);
  }

  const pdfBtn = qs("clPdfBtn");
  const uploadBtn = qs("clUploadPhotoBtn");
  if (pdfBtn) {
    if (clCurrentCheckId > 0) {
      pdfBtn.classList.remove("hidden");
      pdfBtn.dataset.checkId = String(clCurrentCheckId);
    } else {
      pdfBtn.classList.add("hidden");
      pdfBtn.dataset.checkId = "";
    }
  }
  if (uploadBtn) uploadBtn.disabled = clCurrentCheckId <= 0;
  qs("clFormPanel")?.scrollIntoView({ behavior: "smooth", block: "start" });
  setStatus(`Checklist loaded for ${code}`);
}

function clKmLooksUnusual(odometerKm, previousKm) {
  const odo = Number(odometerKm);
  const prev = previousKm == null ? null : Number(previousKm);
  if (!Number.isFinite(odo) || odo < 0) return false;
  if (prev != null && Number.isFinite(prev) && odo < prev) return true;
  if (odo > 500000) return true;
  if (prev != null && Number.isFinite(prev) && odo > prev * 1.25 + 500) return true;
  return false;
}

async function submitChecklistForm() {
  if (!clSelectedAssetCode) return alert("Select an asset from the checklist sections first.");
  clSetFormMsg("", null);
  clShowSyncState(null);
  const checklist = readClChecklistObject();
  const inspector_name = String(qs("clInspector")?.value || "").trim();
  const notes = String(qs("clNotes")?.value || "").trim();

  if (clSelectedKind === "ldv") {
    const odoRaw = String(qs("clOdometer")?.value || "").trim();
    if (!odoRaw) throw new Error("Enter current odometer KM.");
    const odometer_km = Number(odoRaw);
    if (!Number.isFinite(odometer_km) || odometer_km < 0) throw new Error("Odometer must be a valid number ≥ 0.");
    if (clKmLooksUnusual(odometer_km, clPreviousKm)) {
      const prevTxt = clPreviousKm == null ? "the previous reading" : `${clPreviousKm.toFixed(1)} km`;
      const ok = window.confirm(
        `KM ${odometer_km.toFixed(1)} looks unusual compared with ${prevTxt}. Submit pre-start anyway? Daily input will not be updated until a supervisor reviews.`
      );
      if (!ok) return;
    }
    const data = await fetchJson(`${API}/api/maintenance/vehicle-ldv-checks/prestart`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({
        asset_code: clSelectedAssetCode,
        check_date: clCheckDate(),
        odometer_km,
        inspector_name,
        notes,
        checklist,
      }),
    });
    clCurrentCheckId = Number(data?.id || 0);
    clPreviousKm = data?.odometer_km != null ? Number(data.odometer_km) : clPreviousKm;
    if (qs("clPrevKm")) {
      qs("clPrevKm").textContent = clPreviousKm == null ? "—" : `${clPreviousKm.toFixed(1)} km`;
    }
    clShowSyncState(data?.daily_input_sync || null);
    clSetFormMsg(data?.message || "LDV pre-start saved.", data?.km_review_needed ? false : true);
  } else {
    const smuRaw = String(qs("clSmuHours")?.value || "").trim();
    let smu_hours = null;
    if (smuRaw) {
      const smu = Number(smuRaw);
      if (!Number.isFinite(smu) || smu < 0) throw new Error("SMU hours must be a valid number ≥ 0.");
      if (clPreviousSmu != null && smu < clPreviousSmu) {
        throw new Error(`SMU cannot be less than previous hours (${clPreviousSmu.toFixed(1)}).`);
      }
      smu_hours = smu;
    }
    const data = await fetchJson(`${API}/api/maintenance/machine-prestart`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({
        asset_code: clSelectedAssetCode,
        check_date: clCheckDate(),
        smu_hours,
        inspector_name,
        notes,
        checklist,
      }),
    });
    clCurrentCheckId = Number(data?.id || 0);
    clShowSyncState(data?.daily_input_sync || null);
    clSetFormMsg(data?.message || "Machine pre-start saved.", true);
  }

  const pdfBtn = qs("clPdfBtn");
  if (pdfBtn && clCurrentCheckId > 0) {
    pdfBtn.classList.remove("hidden");
    pdfBtn.dataset.checkId = String(clCurrentCheckId);
  }
  qs("clUploadPhotoBtn") && (qs("clUploadPhotoBtn").disabled = clCurrentCheckId <= 0);
  setStatus("Checklist saved ✅");
  await loadChecklistHub();
  renderClHubSections(clHubData);
  loadClHistory().catch(() => {});
}

function canCorrectChecklistMeter() {
  return getSessionRoles().some((r) => ["admin", "supervisor", "plant_manager", "site_manager"].includes(r));
}

function updateClLdvSupervisorPanel() {
  const panel = qs("clLdvSupervisorPanel");
  if (!panel) return;
  panel.classList.toggle("hidden", !canCorrectChecklistMeter() || clSelectedKind !== "ldv");
}

function updateClMachineSupervisorPanel() {
  const panel = qs("clMachineSupervisorPanel");
  if (!panel) return;
  panel.classList.toggle("hidden", !canCorrectChecklistMeter() || clSelectedKind !== "machine");
}

async function applyMachineHoursCorrection() {
  if (!clSelectedAssetCode) return alert("Select a machine asset first.");
  const closing_hours = Number(String(qs("clCorrectClosingHours")?.value || "").trim());
  if (!Number.isFinite(closing_hours) || closing_hours < 0) return alert("Enter the correct closing hours.");
  const openingRaw = String(qs("clCorrectOpeningHours")?.value || "").trim();
  const body = {
    asset_code: clSelectedAssetCode,
    work_date: clCheckDate(),
    closing_hours,
    notes: `Corrected via Checklists tab by ${getSessionUser() || "supervisor"}`,
  };
  if (openingRaw) body.opening_hours = Number(openingRaw);
  setStatus("Applying hours correction…");
  const data = await fetchJson(`${API}/api/maintenance/machine-prestart/hours-correction`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify(body),
  });
  if (qs("clSmuHours")) qs("clSmuHours").value = String(data.closing_hours ?? closing_hours);
  if (qs("clPrevSmu")) {
    qs("clPrevSmu").textContent =
      data.opening_hours != null ? `${Number(data.opening_hours).toFixed(1)} hrs` : "—";
  }
  clSetFormMsg(data.message || "Hours corrected.", true);
  setStatus(`Hours corrected for ${clSelectedAssetCode} ✅`);
  await loadChecklistHub();
  renderClHubSections(clHubData);
  await selectChecklistAsset(clSelectedAssetCode).catch(() => {});
  loadClHistory().catch(() => {});
}

async function applyLdvKmCorrection() {
  if (!clSelectedAssetCode) return alert("Select an LDV asset first.");
  const closing_km = Number(String(qs("clCorrectClosingKm")?.value || "").trim());
  if (!Number.isFinite(closing_km) || closing_km < 0) return alert("Enter the correct closing KM.");
  const openingRaw = String(qs("clCorrectOpeningKm")?.value || "").trim();
  const body = {
    asset_code: clSelectedAssetCode,
    work_date: clCheckDate(),
    closing_km,
    notes: `Corrected via Checklists tab by ${getSessionUser() || "supervisor"}`,
  };
  if (openingRaw) body.opening_km = Number(openingRaw);
  setStatus("Applying KM correction…");
  const data = await fetchJson(`${API}/api/maintenance/vehicle-ldv-checks/prestart-correction`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify(body),
  });
  const correctedKm = Number(data.closing_km ?? closing_km);
  clPreviousKm = data.previous_odometer_km != null ? Number(data.previous_odometer_km) : clPreviousKm;
  if (qs("clOdometer")) qs("clOdometer").value = String(correctedKm);
  if (qs("clPrevKm")) {
    qs("clPrevKm").textContent =
      clPreviousKm == null ? "—" : `${clPreviousKm.toFixed(1)} km`;
  }
  if (qs("clCorrectClosingKm")) qs("clCorrectClosingKm").value = "";
  if (qs("clCorrectOpeningKm")) {
    qs("clCorrectOpeningKm").value = data.opening_km != null ? String(data.opening_km) : "";
  }
  clCurrentCheckId = Number(data.check_id || clCurrentCheckId || 0);
  const purgeNote =
    data?.purged_baselines?.daily_rows || data?.purged_baselines?.check_rows
      ? ` Cleared ${Number(data.purged_baselines.daily_rows || 0)} bad daily row(s) and ${Number(data.purged_baselines.check_rows || 0)} stray check row(s).`
      : "";
  clSetFormMsg((data.message || "KM corrected.") + purgeNote, true);
  setStatus(`KM corrected for ${clSelectedAssetCode} ✅`);
  await loadChecklistHub();
  renderClHubSections(clHubData);
  await selectChecklistAsset(clSelectedAssetCode).catch(() => {});
  loadClHistory().catch(() => {});
}

async function uploadChecklistPhoto() {
  if (!clCurrentCheckId) return alert("Submit the checklist first.");
  const file = qs("clPhotoFile")?.files?.[0];
  if (!file) return alert("Choose a photo file.");
  const fd = new FormData();
  fd.append("file", file);
  const headers = new Headers(authHeaders());
  headers.delete("Content-Type");
  const caption = clSelectedKind === "ldv" ? "Pre-start photo" : "Machine pre-start photo";
  const res = await fetch(
    `${API}/api/maintenance/vehicle-ldv-checks/${clCurrentCheckId}/photo?caption=${encodeURIComponent(caption)}`,
    { method: "POST", headers, body: fd }
  );
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || data.message || "Upload failed");
  if (qs("clPhotoFile")) qs("clPhotoFile").value = "";
  clSetFormMsg("Photo uploaded to this checklist.", true);
  setStatus("Checklist photo uploaded ✅");
  loadClHistory().catch(() => {});
}

function clCheckModeLabel(mode) {
  const m = String(mode || "").trim();
  if (m === "prestart") return "LDV Pre-Start";
  if (m.startsWith("machine_prestart_")) {
    const profileId = m.replace("machine_prestart_", "");
    const titles = {
      excavator: "Excavator",
      dozer: "Dozer",
      wheel_loader: "Wheel loader",
      haul_truck: "Haul truck",
      fuel_truck: "Fuel truck",
      grader: "Grader",
      mobile_crane: "Mobile crane",
      crusher: "Crusher",
      mobile_screen: "Mobile screen",
      generator: "Generator",
      backhoe_loader: "Backhoe / loader",
    };
    return titles[profileId] || profileId.replace(/_/g, " ");
  }
  if (m === "ldv_general") return "Vehicle Check";
  return m || "Checklist";
}

function clInitHistoryDates() {
  const end = qs("clHistEnd");
  const start = qs("clHistStart");
  const today = todayLocalYmd();
  if (end && !end.value) end.value = today;
  if (start && !start.value) {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    start.value = d.toISOString().slice(0, 10);
  }
}

async function clResolveHistAssetId() {
  const code = String(qs("clHistAsset")?.value || "").trim().toUpperCase();
  if (!code) return 0;
  try {
    const data = await fetchJson(`${API}/api/assets`);
    const list = Array.isArray(data) ? data : [];
    const hit = list.find((a) => String(a.asset_code || "").toUpperCase() === code);
    return hit ? Number(hit.id) : 0;
  } catch {
    return 0;
  }
}

function clOpenCheckPdf(checkId, download = false) {
  const id = Number(checkId || 0);
  if (!id) return alert("No check selected.");
  const q = download ? "?download=1" : "";
  window.open(`${API}/api/reports/vehicle-ldv-check/${id}.pdf${q}`, "_blank");
}

async function clOpenRangePdf(download = false) {
  const start = String(qs("clHistStart")?.value || "").trim();
  const end = String(qs("clHistEnd")?.value || "").trim();
  if (!start || !end) return alert("Select From and To dates first.");
  const assetId = await clResolveHistAssetId();
  const q = new URLSearchParams({ start, end, with_photos: "1" });
  if (assetId > 0) q.set("asset_id", String(assetId));
  if (download) q.set("download", "1");
  window.open(`${API}/api/reports/vehicle-ldv-checks.pdf?${q.toString()}`, "_blank");
}

function renderClHistory(rows) {
  const list = qs("clHistoryList");
  const summary = qs("clHistorySummary");
  if (!list) return;
  list.innerHTML = "";
  const items = Array.isArray(rows) ? rows : [];
  if (summary) {
    summary.textContent = items.length
      ? `${items.length} check(s) in selected period — photos included in PDF exports.`
      : "No checks found for the selected filters.";
  }
  if (!items.length) {
    list.innerHTML = `<div class="muted small">No checklist records found. Try a wider date range or clear filters.</div>`;
    return;
  }
  const table = document.createElement("div");
  table.className = "cl-history-table";
  table.innerHTML = `
    <div class="cl-history-head">
      <span>Date</span><span>Asset</span><span>Type</span><span>Inspector</span><span>Photos</span><span>Actions</span>
    </div>
  `;
  items.forEach((r) => {
    const row = document.createElement("div");
    row.className = "cl-history-row";
    const photoCount = Number(r.photo_count ?? (r.photos || []).length ?? 0);
    const meter =
      r.odometer_km != null
        ? `${Number(r.odometer_km).toFixed(0)} km`
        : r.smu_hours != null
          ? `${Number(r.smu_hours).toFixed(1)} h SMU`
          : "";
    row.innerHTML = `
      <span>${escapeHtml(r.check_date || "-")}</span>
      <span><b>${escapeHtml(r.asset_code || "-")}</b><br><small class="muted">${escapeHtml(r.asset_name || "")}${meter ? ` · ${escapeHtml(meter)}` : ""}</small></span>
      <span><span class="pill">${escapeHtml(clCheckModeLabel(r.check_mode))}</span></span>
      <span>${escapeHtml(r.inspector_name || "-")}</span>
      <span>${photoCount > 0 ? `<span class="pill green">${photoCount}</span>` : `<span class="muted">0</span>`}</span>
      <span class="cl-history-actions">
        <button type="button" class="btn btn-secondary btn-sm" data-cl-view="${Number(r.id)}">View</button>
        <button type="button" class="btn btn-secondary btn-sm" data-cl-pdf="${Number(r.id)}">PDF</button>
        <button type="button" class="btn btn-secondary btn-sm" data-cl-dlpdf="${Number(r.id)}">Save</button>
      </span>
    `;
    row.querySelector("[data-cl-view]")?.addEventListener("click", () => {
      if (qs("clCheckDate")) qs("clCheckDate").value = String(r.check_date || clCheckDate());
      selectChecklistAsset(String(r.asset_code || "")).catch((e) => clSetFormMsg(String(e.message || e), false));
    });
    row.querySelector("[data-cl-pdf]")?.addEventListener("click", () => clOpenCheckPdf(r.id, false));
    row.querySelector("[data-cl-dlpdf]")?.addEventListener("click", () => clOpenCheckPdf(r.id, true));
    table.appendChild(row);
  });
  list.appendChild(table);
}

async function loadClHistory() {
  const start = String(qs("clHistStart")?.value || "").trim();
  const end = String(qs("clHistEnd")?.value || "").trim();
  const checkMode = String(qs("clHistType")?.value || "").trim();
  const assetCode = String(qs("clHistAsset")?.value || "").trim().toUpperCase();
  if (!start || !end) {
    clInitHistoryDates();
    return loadClHistory();
  }
  const list = qs("clHistoryList");
  if (list) list.innerHTML = `<div class="muted small">Loading history…</div>`;
  try {
    const q = new URLSearchParams({ start, end });
    if (checkMode) q.set("check_mode", checkMode);
    if (assetCode) q.set("asset_code", assetCode);
    const data = await fetchJson(`${API}/api/maintenance/vehicle-ldv-checks?${q.toString()}`);
    renderClHistory(data?.rows || []);
  } catch (e) {
    if (list) list.innerHTML = `<div class="muted small">History load error: ${escapeHtml(e.message || e)}</div>`;
  }
}

function initChecklistTab() {
  if (window.__clTabInit) return;
  window.__clTabInit = true;
  const d = qs("clCheckDate");
  if (d && !d.value) d.value = todayLocalYmd();
  clInitHistoryDates();
  const ins = qs("clInspector");
  if (ins && !ins.value) ins.value = getSessionUser();

  qs("clRefreshHub")?.addEventListener("click", () => loadChecklistHub().catch((e) => setStatus(String(e.message || e))));
  qs("clPtTranslateBtn")?.addEventListener("click", runClPtTranslate);
  qs("clPtTranslateIn")?.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === "Enter") runClPtTranslate();
  });
  qs("clCheckDate")?.addEventListener("change", () => {
    clSelectedAssetCode = "";
    qs("clFormPanel")?.classList.add("hidden");
    loadChecklistHub().catch(() => {});
  });
  qs("clHubSections")?.addEventListener("click", (e) => {
    const safetyChip = e.target.closest(".checklist-safety-chip");
    if (safetyChip) {
      const code = String(safetyChip.dataset.itemCode || "").trim();
      if (!code) return;
      window.location.href = `./safety-inspection.html?item_code=${encodeURIComponent(code)}`;
      return;
    }
    const chip = e.target.closest(".checklist-asset-chip");
    if (!chip) return;
    const code = chip.dataset.assetCode;
    if (!code) return;
    selectChecklistAsset(code).catch((err) => clSetFormMsg(String(err.message || err), false));
  });
  qs("clSubmitBtn")?.addEventListener("click", () =>
    submitChecklistForm().catch((e) => clSetFormMsg(String(e.message || e), false))
  );
  qs("clUploadPhotoBtn")?.addEventListener("click", () =>
    uploadChecklistPhoto().catch((e) => clSetFormMsg(String(e.message || e), false))
  );
  qs("clPdfBtn")?.addEventListener("click", () => {
    const id = Number(qs("clPdfBtn")?.dataset.checkId || clCurrentCheckId || 0);
    if (!id) return alert("Submit the checklist first.");
    window.open(`${API}/api/reports/vehicle-ldv-check/${id}.pdf`, "_blank");
  });
  qs("clHistLoad")?.addEventListener("click", () => loadClHistory().catch((e) => setStatus(String(e.message || e))));
  qs("clHistOpenRangePdf")?.addEventListener("click", () => clOpenRangePdf(false));
  qs("clHistDownloadRangePdf")?.addEventListener("click", () => clOpenRangePdf(true));
  qs("clApplyKmCorrection")?.addEventListener("click", () =>
    applyLdvKmCorrection().catch((e) => clSetFormMsg(String(e.message || e), false))
  );
  qs("clApplyHoursCorrection")?.addEventListener("click", () =>
    applyMachineHoursCorrection().catch((e) => clSetFormMsg(String(e.message || e), false))
  );
}

function openChecklistTabForAsset(assetCode) {
  const code = String(assetCode || "").trim().toUpperCase();
  if (!code) return;
  clPendingAssetCode = code;
  switchTab("vehicle");
}

/** LDV vehicle check — photos + fractional damage pins */
const vcMarkerDrafts = new Map();
let vcActiveCheckId = null;

function vcImgUrl(filePath) {
  const n = normalizeImageSrc(String(filePath || ""));
  if (!n) return "";
  return /^https?:\/\//i.test(n) ? n : `${API}${n}`;
}

async function loadVcAssetSelect() {
  const sel = qs("vcAsset");
  if (!sel) return;
  try {
    const data = await fetchJson(`${API}/api/assets`);
    const list = Array.isArray(data) ? data : [];
    const cur = sel.value;
    sel.innerHTML = '<option value="">Select vehicle…</option>';
    list.forEach((a) => {
      if (Number(a.archived) === 1) return;
      const o = document.createElement("option");
      o.value = String(a.id);
      o.textContent = `${a.asset_code} — ${a.asset_name || ""}`;
      sel.appendChild(o);
    });
    if (cur) sel.value = cur;
  } catch (e) {
    setStatus("Vehicle list: " + (e.message || e));
  }
}

async function vcCreateCheck() {
  const asset_id = Number(qs("vcAsset")?.value || 0);
  const check_date = qs("vcDate")?.value || new Date().toISOString().slice(0, 10);
  const vehicle_registration = String(qs("vcReg")?.value || "").trim() || null;
  const odoEl = qs("vcOdo");
  const odometer_km = odoEl && String(odoEl.value).trim() !== "" ? Number(odoEl.value) : null;
  const inspector_name = String(qs("vcInspector")?.value || "").trim() || null;
  const notes = String(qs("vcNotes")?.value || "").trim() || null;
  if (!asset_id) return alert("Select a vehicle asset.");
  setStatus("Creating vehicle check…");
  try {
    const res = await fetchJson(`${API}/api/maintenance/vehicle-ldv-checks`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({
        asset_id,
        check_date,
        vehicle_registration,
        odometer_km: odometer_km != null && Number.isFinite(odometer_km) ? odometer_km : null,
        inspector_name,
        notes,
      }),
    });
    vcActiveCheckId = Number(res.id);
    const lab = qs("vcCheckIdLabel");
    if (lab) lab.textContent = `Check #${vcActiveCheckId}`;
    const up = qs("vcUploadPhoto");
    if (up) up.disabled = false;
    vcMarkerDrafts.clear();
    const ed = qs("vcPhotoEditor");
    if (ed) ed.innerHTML = "";
    setStatus(`Vehicle check #${vcActiveCheckId} started — add photos, then click photo to pin damage.`);
    await vcLoadChecksList();
  } catch (e) {
    setStatus("Vehicle check failed: " + (e.message || e));
  }
}

async function vcUploadPhoto() {
  if (!vcActiveCheckId) return alert("Start a check first.");
  const file = qs("vcPhotoFile")?.files?.[0];
  if (!file) return alert("Choose a photo file.");
  setStatus("Uploading photo…");
  try {
    const fd = new FormData();
    fd.append("file", file);
    const headers = new Headers(authHeaders());
    headers.delete("Content-Type");
    const res = await fetch(`${API}/api/maintenance/vehicle-ldv-checks/${vcActiveCheckId}/photo`, {
      method: "POST",
      headers,
      body: fd,
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || data.message || "Upload failed");
    const pf = qs("vcPhotoFile");
    if (pf) pf.value = "";
    await vcReloadCheckPhotos();
    setStatus("Photo uploaded — click image to add damage pins.");
  } catch (e) {
    setStatus("Photo upload failed: " + (e.message || e));
  }
}

async function vcReloadCheckPhotos() {
  if (!vcActiveCheckId) return;
  const data = await fetchJson(`${API}/api/maintenance/vehicle-ldv-checks?check_id=${vcActiveCheckId}`);
  const row = (data.rows || [])[0];
  if (!row) return;
  vcMarkerDrafts.clear();
  (row.photos || []).forEach((p) => {
    vcMarkerDrafts.set(Number(p.id), JSON.parse(JSON.stringify(p.markers || [])));
  });
  renderVcPhotoEditor(row.photos || []);
}

function renderVcPhotoEditor(photos) {
  const host = qs("vcPhotoEditor");
  if (!host) return;
  host.innerHTML = "";
  (photos || []).forEach((p) => {
    const pid = Number(p.id);
    const wrap = document.createElement("div");
    wrap.className = "vehicle-pin-wrap card stack-8";

    const top = document.createElement("div");
    top.className = "row";
    const lbl = document.createElement("span");
    lbl.className = "muted";
    lbl.textContent = `Photo #${pid}`;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = "Save damage pins";
    btn.addEventListener("click", () =>
      vcSavePins(pid).catch((err) => setStatus(String(err.message || err)))
    );
    top.appendChild(lbl);
    top.appendChild(btn);

    const stage = document.createElement("div");
    stage.className = "vehicle-pin-stage";
    stage.dataset.photoId = String(pid);
    const img = document.createElement("img");
    img.className = "vehicle-pin-img";
    img.alt = "Vehicle photo";
    img.src = vcImgUrl(p.file_path);
    stage.appendChild(img);
    stage.addEventListener("click", (e) => vcOnPhotoClick(e, pid));

    const leg = document.createElement("div");
    leg.className = "vehicle-pin-legend muted";
    leg.dataset.forPin = String(pid);

    wrap.appendChild(top);
    wrap.appendChild(stage);
    wrap.appendChild(leg);
    host.appendChild(wrap);
    vcRedrawPins(pid);
  });
}

function vcOnPhotoClick(e, photoId) {
  if (e.target.classList && e.target.classList.contains("vehicle-pin-dot")) return;
  const stage = e.currentTarget;
  const img = stage.querySelector("img");
  if (!img) return;
  const rect = img.getBoundingClientRect();
  const x = (e.clientX - rect.left) / rect.width;
  const y = (e.clientY - rect.top) / rect.height;
  if (x < 0 || x > 1 || y < 0 || y > 1) return;
  const label = window.prompt("Short label for this damage (e.g. Dent, Scratch)", "Damage");
  if (label === null) return;
  const arr = vcMarkerDrafts.get(photoId) || [];
  arr.push({
    x,
    y,
    label: String(label || "Damage").slice(0, 120),
    note: "",
  });
  vcMarkerDrafts.set(photoId, arr);
  vcRedrawPins(photoId);
}

function vcRedrawPins(photoId) {
  const stage = document.querySelector(`.vehicle-pin-stage[data-photo-id="${photoId}"]`);
  if (!stage) return;
  const markers = vcMarkerDrafts.get(photoId) || [];
  stage.querySelectorAll(".vehicle-pin-dot").forEach((d) => d.remove());
  markers.forEach((m, idx) => {
    const dot = document.createElement("div");
    dot.className = "vehicle-pin-dot";
    dot.title = String(m.label || "Damage");
    dot.style.left = `${(Number(m.x) * 100).toFixed(4)}%`;
    dot.style.top = `${(Number(m.y) * 100).toFixed(4)}%`;
    dot.addEventListener("click", (ev) => {
      ev.stopPropagation();
      const arr = vcMarkerDrafts.get(photoId) || [];
      if (!arr[idx]) return;
      const current = arr[idx];
      const nextLabel = window.prompt(
        "Edit pin label. Type /remove to delete this pin.",
        String(current.label || "Damage")
      );
      if (nextLabel === null) return;
      if (String(nextLabel).trim().toLowerCase() === "/remove") {
        arr.splice(idx, 1);
        vcMarkerDrafts.set(photoId, arr);
        vcRedrawPins(photoId);
        return;
      }
      const nextNote = window.prompt(
        "Optional note for this pin (blank allowed).",
        String(current.note || "")
      );
      if (nextNote === null) return;
      arr[idx] = {
        ...current,
        label: String(nextLabel || "Damage").slice(0, 120),
        note: String(nextNote || "").slice(0, 500),
      };
      vcMarkerDrafts.set(photoId, arr);
      vcRedrawPins(photoId);
    });
    stage.appendChild(dot);
  });
  const leg = document.querySelector(`.vehicle-pin-legend[data-for-pin="${photoId}"]`);
  if (leg) {
    leg.innerHTML = markers.length
      ? `${markers.map((m, i) => `<span style="margin-right:14px">${i + 1}. ${escapeHtml(m.label || "Damage")}${m.note ? ` (${escapeHtml(m.note)})` : ""}</span>`).join("")}<span style="margin-left:10px">Tip: click a red pin to edit, or type /remove.</span>`
      : "Click the photo to add damage pins.";
  }
}

async function vcSavePins(photoId) {
  const markers = vcMarkerDrafts.get(photoId) || [];
  await fetchJson(`${API}/api/maintenance/vehicle-ldv-checks/photos/${photoId}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({ markers }),
  });
  setStatus("Damage pins saved.");
}

async function vcLoadChecksList() {
  const el = qs("vcChecksList");
  const pre = qs("vcResult");
  if (!el) return;
  try {
    const data = await fetchJson(`${API}/api/maintenance/vehicle-ldv-checks`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    el.innerHTML = "";
    if (!rows.length) {
      el.innerHTML = `<div class="muted">No vehicle checks yet.</div>`;
      if (pre) pre.textContent = "";
      return;
    }
    rows.slice(0, 40).forEach((r) => {
      const div = document.createElement("div");
      div.className = "list-row";
      div.style.cssText = "cursor:pointer;padding:8px 0;border-bottom:1px solid var(--line);";
      const n = (r.photos || []).length;
      div.innerHTML = `<strong>#${r.id}</strong> ${escapeHtml(r.asset_code || "")} — ${escapeHtml(r.check_date || "")} ${r.vehicle_registration ? `(${escapeHtml(r.vehicle_registration)})` : ""} <span class="muted">${n} photo(s)</span>`;
      div.addEventListener("click", async () => {
        vcActiveCheckId = Number(r.id);
        const lab = qs("vcCheckIdLabel");
        if (lab) lab.textContent = `Check #${vcActiveCheckId}`;
        const up = qs("vcUploadPhoto");
        if (up) up.disabled = false;
        setStatus(`Loaded check #${vcActiveCheckId}`);
        await vcReloadCheckPhotos();
      });
      el.appendChild(div);
    });
    if (pre) pre.textContent = JSON.stringify({ count: rows.length }, null, 2);
  } catch (e) {
    if (pre) pre.textContent = String(e.message || e);
    setStatus("Failed to load vehicle checks.");
  }
}

function vcOpenPdf(download = false) {
  if (!vcActiveCheckId) return alert("Load or create a check first.");
  const q = download ? "?download=1" : "";
  window.open(`${API}/api/reports/vehicle-ldv-check/${vcActiveCheckId}.pdf${q}`, "_blank");
}

function vcOpenBulkPdf(download = false) {
  const start = String(qs("vcStart")?.value || "").trim();
  const end = String(qs("vcEnd")?.value || "").trim();
  const assetId = Number(qs("vcAsset")?.value || 0);
  if (!start || !end) return alert("Select From and To dates first.");
  const q = new URLSearchParams({
    start,
    end,
    with_photos: "1",
  });
  if (assetId > 0) q.set("asset_id", String(assetId));
  if (download) q.set("download", "1");
  window.open(`${API}/api/reports/vehicle-ldv-checks.pdf?${q.toString()}`, "_blank");
}

function initVehicleCheckTab() {
  initChecklistTab();
  if (window.__vcTabInit) return;
  window.__vcTabInit = true;
  const d = qs("vcDate");
  if (d && !d.value) d.value = new Date().toISOString().slice(0, 10);
  const dStart = qs("vcStart");
  const dEnd = qs("vcEnd");
  if (dStart && !dStart.value) dStart.value = new Date(Date.now() - 1000 * 60 * 60 * 24 * 30).toISOString().slice(0, 10);
  if (dEnd && !dEnd.value) dEnd.value = new Date().toISOString().slice(0, 10);
  const ins = qs("vcInspector");
  if (ins && !ins.value) ins.value = getSessionUser();
  qs("vcCreateCheck")?.addEventListener("click", () => vcCreateCheck().catch((e) => setStatus(String(e.message || e))));
  qs("vcUploadPhoto")?.addEventListener("click", () => vcUploadPhoto().catch((e) => setStatus(String(e.message || e))));
  qs("vcLoadChecks")?.addEventListener("click", () => vcLoadChecksList().catch((e) => setStatus(String(e.message || e))));
  qs("vcOpenPdf")?.addEventListener("click", () => vcOpenPdf(false));
  qs("vcDownloadPdf")?.addEventListener("click", () => vcOpenPdf(true));
  qs("vcOpenBulkPdf")?.addEventListener("click", () => vcOpenBulkPdf(false));
  qs("vcDownloadBulkPdf")?.addEventListener("click", () => vcOpenBulkPdf(true));
  loadVcAssetSelect().catch(() => {});
  vcLoadChecksList().catch(() => {});
}

function setStatus(msg) {
  const el = qs("status");
  if (!el) return;
  el.textContent = translateStatusMessage(msg, getLang());
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

const TELEMATICS_FAULT_POLL_MS = 45000;
let telematicsFaultPollTimer = null;

function telematicsStatusClass(status) {
  const s = String(status || "offline");
  if (s === "live") return "pill green";
  if (s === "stale") return "pill amber";
  return "pill";
}

function buildTelematicsUnitItemHtml(r) {
  const status = String(r.link_status || "offline");
  const statusClass = telematicsStatusClass(status);
  const runHrs = r.run_seconds_today != null ? (Number(r.run_seconds_today) / 3600).toFixed(2) : "-";
  const idleHrs = r.idle_seconds_today != null ? (Number(r.idle_seconds_today) / 3600).toFixed(2) : "-";
  const faults = Number(r.active_fault_count || 0);
  return (
    `<div class="fuel-item-head"><b>${escapeHtml(r.asset_code || "-")}</b> — ${escapeHtml(r.unit_model || "FSC")} <span class="${statusClass}">${status.toUpperCase()}</span>${faults > 0 ? ` <span class="pill red">${faults} FAULT${faults === 1 ? "" : "S"}</span>` : ""}</div>` +
    `<small class="fuel-item-desc">${escapeHtml(r.asset_name || "")}</small>` +
    `<small class="fuel-item-meta">Serial: ${escapeHtml(r.device_serial || "-")} | Meter: ${r.engine_hours == null ? "-" : Number(r.engine_hours).toFixed(1)} h | Run today: ${runHrs} h | Idle today: ${idleHrs} h | Ignition: ${Number(r.ignition_on) === 1 ? "ON" : "OFF"}</small>` +
    `<small class="fuel-item-meta muted">Last seen: ${escapeHtml(r.recorded_at || r.updated_at || r.snapshot_updated_at || "-")}</small>`
  );
}

function renderTelematicsFleetList(container, fleet) {
  if (!container) return;
  container.innerHTML = "";
  const rows = Array.isArray(fleet) ? fleet : [];
  if (!rows.length) {
    container.appendChild(item("<small>No telematics devices registered yet.</small>"));
    return;
  }
  rows.forEach((r) => {
    container.appendChild(item(buildTelematicsUnitItemHtml(r)));
  });
}

function renderTelematicsActiveFaultsList(container, faults) {
  if (!container) return;
  container.innerHTML = "";
  const rows = Array.isArray(faults) ? faults : [];
  if (!rows.length) {
    container.appendChild(item("<small>No active faults — all units clear.</small>"));
    return;
  }
  rows.forEach((f) => {
    const sev = String(f.severity || "warning").toLowerCase();
    const sevClass = sev === "critical" || sev === "error" || sev === "severe" ? "pill red" : "pill amber";
    container.appendChild(
      item(
        `<div class="fuel-item-head"><b>${escapeHtml(f.asset_code || "-")}</b> <span class="${sevClass}">${escapeHtml(sev.toUpperCase())}</span> <code>${escapeHtml(f.fault_code || "-")}</code></div>` +
        `<small class="fuel-item-desc">${escapeHtml(f.description || "No description")}</small>` +
        `<small class="fuel-item-meta muted">${escapeHtml(f.unit_model || "FSC")} · ${escapeHtml(f.event_time || "-")}</small>`
      )
    );
  });
}

function updateTelematicsFaultFloat(summary) {
  const banner = qs("telematicsFaultFloat");
  const titleEl = qs("telematicsFaultFloatTitle");
  const textEl = qs("telematicsFaultFloatText");
  if (!banner || !titleEl || !textEl) return;

  const faults = Array.isArray(summary?.faults) ? summary.faults : [];
  const unitsWithFaults = Number(summary?.units_with_faults || 0);
  const faultCount = Number(summary?.fault_count || faults.length || 0);
  const hasFaults = faultCount > 0 || unitsWithFaults > 0;

  if (!hasFaults) {
    banner.classList.add("hidden");
    document.body.classList.remove("telematics-fault-visible");
    return;
  }

  const assetCodes = Array.from(new Set(faults.map((f) => f.asset_code).filter(Boolean)));
  const preview = faults.slice(0, 3).map((f) => {
    const code = f.asset_code || "?";
    const fc = f.fault_code || "?";
    return `${code}: ${fc}`;
  });
  const more = faults.length > 3 ? ` (+${faults.length - 3} more)` : "";

  titleEl.textContent =
    unitsWithFaults === 1
      ? "Active machine fault"
      : `${unitsWithFaults || assetCodes.length} machine(s) with active faults`;
  textEl.textContent = preview.length
    ? `${preview.join(" · ")}${more}`
    : `${faultCount} active fault signal(s) on telematics units`;

  banner.classList.remove("hidden");
  document.body.classList.add("telematics-fault-visible");
}

async function refreshTelematicsFaultBanner() {
  const allowed = getEffectiveAllowedTabs();
  if (!allowed.includes("telematics") && !allowed.includes("dash")) return;
  try {
    const data = await fetchJson(`${API}/api/telematics/faults/active`);
    updateTelematicsFaultFloat(data);
  } catch (_) {
    /* silent — banner stays as last known state */
  }
}

function initTelematicsFaultBanner() {
  qs("telematicsFaultFloatView")?.addEventListener("click", () => {
    switchTab("telematics");
    loadTelematicsTab().catch(() => {});
  });
  refreshTelematicsFaultBanner().catch(() => {});
  if (telematicsFaultPollTimer) clearInterval(telematicsFaultPollTimer);
  telematicsFaultPollTimer = setInterval(() => {
    refreshTelematicsFaultBanner().catch(() => {});
  }, TELEMATICS_FAULT_POLL_MS);
}

async function loadTelematicsTab() {
  const unitsList = qs("telematicsActiveUnitsList");
  const faultsList = qs("telematicsActiveFaultsList");
  if (!unitsList && !faultsList) return;

  setStatus("Loading telematics…");
  if (unitsList) setSkeleton("telematicsActiveUnitsList", 3);
  if (faultsList) setSkeleton("telematicsActiveFaultsList", 2);

  try {
    const [fleetData, faultData] = await Promise.all([
      fetchJson(`${API}/api/telematics/fleet`),
      fetchJson(`${API}/api/telematics/faults/active`),
    ]);
    const fleet = Array.isArray(fleetData?.fleet) ? fleetData.fleet : [];
    renderTelematicsFleetList(unitsList, fleet);
    renderTelematicsActiveFaultsList(faultsList, faultData?.faults || []);
    updateTelematicsFaultFloat(faultData);

    const liveCount = fleet.filter((r) => String(r.link_status) === "live").length;
    const faultUnits = Number(faultData?.units_with_faults || 0);
    const setText = (id, v) => {
      const el = qs(id);
      if (el) el.textContent = String(v);
    };
    setText("telematicsKpiUnits", fleet.length);
    setText("telematicsKpiLive", liveCount);
    setText("telematicsKpiFaults", faultUnits);
    setText("telematicsFaultsBadge", faultData?.fault_count || faultData?.faults?.length || 0);
    qs("telematicsFaultsCard")?.classList.toggle("kpi-alert-crit", faultUnits > 0);
    qs("telematicsKpiFaultsPill")?.classList.toggle("kpi-pill-red", faultUnits > 0);
    setStatus(`Telematics loaded — ${fleet.length} unit(s), ${faultUnits} with faults.`);
  } catch (e) {
    if (unitsList) {
      unitsList.innerHTML = "";
      unitsList.appendChild(item(`<small>Telematics unavailable: ${escapeHtml(e.message || String(e))}</small>`));
    }
    setStatus(`Telematics error: ${e.message || e}`);
  }
}

async function loadTelematicsFleet() {
  const list = qs("telematicsFleetList");
  if (!list) return;
  try {
    const data = await fetchJson(`${API}/api/telematics/fleet`);
    renderTelematicsFleetList(list, data?.fleet);
    if (data?.active_fault_count != null || data?.units_with_faults != null) {
      updateTelematicsFaultFloat({
        fault_count: data.active_fault_count,
        units_with_faults: data.units_with_faults,
        faults: [],
      });
    }
  } catch (e) {
    list.innerHTML = "";
    list.appendChild(item(`<small>Telematics unavailable: ${escapeHtml(e.message || String(e))}</small>`));
  }
}

let cartrackDashboardFleetCache = { fleet: [], speedingToday: [] };

function isGpsFleetVehicleInUse(v) {
  if (!v?.has_gps) return false;
  if (v.is_speeding) return true;
  if (v.gps_source === "unitech") {
    if (v.position_stale) return false;
    return Number(v.speed_kmh || 0) > 0;
  }
  return Number(v.ignition_on) === 1;
}

function renderCartrackFleetIgnitionCell(v) {
  if (v.gps_source === "unitech") {
    const spd = Number(v.speed_kmh || 0);
    if (spd > 0) return '<span class="pill pill-green">Moving</span>';
    return '<span class="pill">Idle</span>';
  }
  const ign = Number(v.ignition_on) === 1;
  return ign ? '<span class="pill pill-green">ON</span>' : '<span class="pill">OFF</span>';
}

function renderCartrackFleetTable(fleet, speedingToday, { showAll = false } = {}) {
  const host = qs("cartrackFleetHost");
  if (!host) return;
  const all = Array.isArray(fleet) ? fleet : [];
  if (!all.length) {
    host.innerHTML = `<div class="cartrack-empty muted small">No GPS vehicles synced yet. If Test connection succeeds but shows 0 vehicles, ask Cartrack to assign your fleet to the API user. Then click Sync now.</div>`;
    return;
  }
  const inUse = all.filter(isGpsFleetVehicleInUse);
  const display = (showAll ? all : inUse).slice().sort((a, b) => {
    if (Boolean(b.is_speeding) !== Boolean(a.is_speeding)) {
      return Number(b.is_speeding) - Number(a.is_speeding);
    }
    return Number(b.speed_kmh || 0) - Number(a.speed_kmh || 0);
  });
  const note = qs("cartrackDashFleetFilterNote");
  if (note) {
    note.textContent = showAll
      ? `Showing all ${all.length} tracked vehicle(s).`
      : inUse.length
        ? `Showing ${inUse.length} in use (ignition on / moving). ${all.length - inUse.length} parked — open Fleet Track for full list.`
        : `No vehicles in use right now. Tick “Show all” or open Fleet Track for the full list.`;
  }
  if (!display.length) {
    host.innerHTML = `<div class="cartrack-empty muted small">No vehicles in use right now (${all.length} tracked, ignition off / stationary). Tick <strong>Show all tracked vehicles</strong> above or open <strong>Fleet Track</strong> for the full list.</div>`;
    return;
  }
  const speedMap = new Map();
  (speedingToday || []).forEach((e) => {
    const k = e.asset_code || e.registration;
    speedMap.set(k, (speedMap.get(k) || 0) + 1);
  });
  const rows = display
    .map((v) => {
      const code = cartrackVehicleLabel(v);
      const sub = cartrackVehicleSubLabel(v);
      const spd = Number(v.speed_kmh || 0);
      const speedEvents = speedMap.get(v.asset_code) || speedMap.get(v.registration) || speedMap.get(code) || 0;
      const rowCls = speedEvents > 0 ? "cartrack-row--alert" : "";
      const batteryPills = renderCartrackBatteryPillsHtml(v);
      return `<tr class="${rowCls}">
        <td>
          <span class="cartrack-vehicle-code">${escapeHtml(code)}</span>
          ${sub ? `<div class="muted mini">${escapeHtml(sub)}</div>` : ""}
        </td>
        <td>${cartrackSourcePillHtml(v)} ${escapeHtml(v.vehicle_name || v.registration || "—")}</td>
        <td class="cartrack-col-num">${spd.toFixed(0)}</td>
        <td class="cartrack-col-status">${renderCartrackFleetIgnitionCell(v)}</td>
        <td class="cartrack-col-battery">${batteryPills || '<span class="muted">—</span>'}</td>
        <td class="cartrack-col-num">${speedEvents ? `<span class="pill pill-red">${speedEvents}</span>` : "—"}</td>
        <td class="cartrack-col-sync muted">${escapeHtml(String(v.synced_at || "").slice(0, 16))}</td>
      </tr>`;
    })
    .join("");
  host.innerHTML = `
    <div class="cartrack-table-scroll">
      <table class="cartrack-fleet-table">
        <thead>
          <tr>
            <th>Vehicle</th>
            <th>Name / reg</th>
            <th class="cartrack-col-num">Speed</th>
            <th>Ignition</th>
            <th>Battery / power</th>
            <th class="cartrack-col-num">Speeding</th>
            <th>Synced</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

async function loadCartrackFleet() {
  const host = qs("cartrackFleetHost");
  const hint = qs("cartrackFleetHint");
  if (!host) return;
  try {
    const data = await fetchJson(`${API}/api/cartrack/fleet`);
    const s = data?.summary || {};
    cartrackDashboardFleetCache = {
      fleet: data.fleet || [],
      speedingToday: data.speeding_today || [],
    };
    setText("cartrackKpiTotal", Number(s.total_vehicles || 0));
    setText("cartrackKpiLive", Number(s.ignition_on || 0));
    setText("cartrackKpiSpeeding", Number(s.speeding_today || 0));
    if (hint) {
      const ct = Number(s.cartrack_vehicles ?? s.total_vehicles ?? 0);
      const ut = Number(s.unitech_vehicles || 0);
      const parts = [];
      if (data.configured) {
        parts.push(`GPS connected (${data.base_url || "MZ"})`);
        if (ut) parts.push(`${ut} Unitech Afungi`);
        parts.push(`Last sync: ${s.last_sync || "—"}`);
      } else {
        parts.push("GPS not configured — add credentials in User Admin → GPS fleet.");
      }
      parts.push("Dashboard table: in use only unless Show all is ticked.");
      hint.textContent = parts.join(" · ");
    }
    renderCartrackFleetTable(data.fleet || [], data.speeding_today || [], {
      showAll: Boolean(qs("cartrackDashShowAll")?.checked),
    });
    loadCartrackSpeedingEvents(todayLocalYmd(), { useCache: true }).catch(() => {});
  } catch (e) {
    host.innerHTML = `<div class="cartrack-empty muted small">Cartrack: ${escapeHtml(e.message || String(e))}</div>`;
  }
}

async function syncCartrackNow() {
  setStatus("Syncing Cartrack fleet…");
  const today = todayLocalYmd();
  await fetchJson(`${API}/api/cartrack/sync`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({ start_date: today, end_date: today }),
  });
  await loadCartrackFleet();
  loadCartrackSpeedingEvents(todayLocalYmd(), { useCache: true }).catch(() => {});
  refreshCartrackSpeedFloat({ refresh: false }).catch(() => {});
  setStatus("Cartrack sync complete ✅");
}

function yesterdayYmd() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function getCartrackSpeedReportDate() {
  const el = qs("cartrackSpeedReportDate") || qs("cartrackTrackSpeedReportDate");
  const v = String(el?.value || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  return todayLocalYmd();
}

function initCartrackSpeedReportDates() {
  const today = todayLocalYmd();
  for (const id of ["cartrackSpeedReportDate", "cartrackTrackSpeedReportDate"]) {
    const el = qs(id);
    if (el && !el.value) el.value = today;
  }
}

function syncCartrackSpeedReportDateInputs(date) {
  const d = String(date || getCartrackSpeedReportDate()).slice(0, 10);
  for (const id of ["cartrackSpeedReportDate", "cartrackTrackSpeedReportDate"]) {
    const el = qs(id);
    if (el) el.value = d;
  }
  return d;
}

function renderCartrackSpeedingEventsTable(host, events, date) {
  if (!host) return;
  if (!events.length) {
    host.innerHTML = `<div class="cartrack-empty muted small">No speeding events recorded for ${escapeHtml(date)}. Events appear when GPS sync detects speed over the limit (Cartrack ≥ alert threshold, Unitech &gt; 60 km/h).</div>`;
    return;
  }
  const rows = events
    .map((e) => {
      const vehicle = e.asset_code || e.registration || "—";
      const reg = e.registration && e.registration !== vehicle ? e.registration : "";
      const time = String(e.event_time || "").slice(0, 16);
      const speed = e.speed_kmh != null ? `${Number(e.speed_kmh).toFixed(0)} km/h` : "—";
      const limit = e.speed_limit_kmh != null ? `${Number(e.speed_limit_kmh).toFixed(0)} km/h` : "—";
      const type = e.event_type_label || e.event_type || "";
      return `<tr>
        <td class="cartrack-col-sync">${escapeHtml(time)}</td>
        <td>
          <span class="cartrack-vehicle-code">${escapeHtml(vehicle)}</span>
          ${reg ? `<div class="muted mini">${escapeHtml(reg)}</div>` : ""}
        </td>
        <td class="cartrack-col-num">${escapeHtml(speed)}</td>
        <td class="cartrack-col-num">${escapeHtml(limit)}</td>
        <td class="muted mini">${escapeHtml(type)}</td>
      </tr>`;
    })
    .join("");
  host.innerHTML = `
    <div class="cartrack-table-scroll">
      <table class="cartrack-fleet-table cartrack-speeding-events-table">
        <thead>
          <tr>
            <th>Time</th>
            <th>Vehicle</th>
            <th class="cartrack-col-num">Speed</th>
            <th class="cartrack-col-num">Limit</th>
            <th>Type</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    <p class="muted mini cartrack-speeding-events-foot">${events.length} event(s) on ${escapeHtml(date)} — click <strong>Speeding PDF</strong> for a printable report.</p>
  `;
}

async function loadCartrackSpeedingEvents(date, {
  hostId = "cartrackSpeedingEventsHost",
  countId = "cartrackSpeedingEventsCount",
  panelId = "cartrackSpeedingEventsPanel",
  useCache = true,
} = {}) {
  const host = qs(hostId);
  if (!host) return [];
  const reportDate = syncCartrackSpeedReportDateInputs(date);
  const cached = useCache && reportDate === todayLocalYmd() ? cartrackDashboardFleetCache.speedingToday : null;
  try {
    const events = cached?.length
      ? cached
      : (await fetchJson(
          `${API}/api/cartrack/events?start=${encodeURIComponent(reportDate)}&end=${encodeURIComponent(reportDate)}&speeding_only=1`
        ))?.rows || [];
    renderCartrackSpeedingEventsTable(host, events, reportDate);
    const countEl = qs(countId);
    if (countEl) countEl.textContent = events.length ? `(${events.length})` : "";
    const panel = qs(panelId);
    if (panel && events.length) panel.open = true;
    return events;
  } catch (e) {
    host.innerHTML = `<div class="cartrack-empty muted small">Could not load speeding log: ${escapeHtml(e.message || String(e))}</div>`;
    return [];
  }
}

function openCartrackSpeedingEventsPanel() {
  const panel = qs("cartrackSpeedingEventsPanel");
  if (!panel) return;
  panel.open = true;
  panel.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

async function openCartrackMorningPdf(date) {
  const reportDate = syncCartrackSpeedReportDateInputs(date);
  setStatus(`Opening speeding PDF for ${reportDate}…`);
  try {
    await openAuthedPdf(
      `${API}/api/cartrack/morning-report.pdf?date=${encodeURIComponent(reportDate)}&_=${Date.now()}`
    );
    setStatus(`Speeding PDF opened (${reportDate})`);
  } catch (e) {
    setStatus(`Speeding PDF error: ${e.message || e}`);
  }
}

async function emailCartrackMorningReport(date) {
  const reportDate = syncCartrackSpeedReportDateInputs(date);
  setStatus(`Sending speeding report for ${reportDate}…`);
  const data = await fetchJson(`${API}/api/cartrack/morning-report/send`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({ date: reportDate }),
  });
  setStatus(`Speeding report emailed (${data.summary?.total_speeding_events ?? 0} events) ✅`);
}

function setCartrackAdminResult(text, ok) {
  const el = qs("cartrackAdminResult");
  if (!el) return;
  el.textContent = String(text || "");
  el.style.color = ok === true ? "#15803d" : ok === false ? "#b91c1c" : "";
}

async function loadCartrackAdminSettings() {
  if (!qs("adminCartrackCard")) return;
  try {
    const data = await fetchJson(`${API}/api/cartrack/settings`);
    const s = data?.settings || {};
    if (qs("cartrackBaseUrl")) qs("cartrackBaseUrl").value = s.base_url || "";
    if (qs("cartrackUsername")) qs("cartrackUsername").value = s.username || "";
    if (qs("cartrackMorningRecipients")) qs("cartrackMorningRecipients").value = s.morning_recipients || "";
    if (qs("cartrackMorningEnabled")) qs("cartrackMorningEnabled").checked = s.morning_enabled !== false;
    const [hh, mm] = String(s.morning_time || "06:00").split(":");
    if (qs("cartrackMorningHour")) qs("cartrackMorningHour").value = hh || "6";
    if (qs("cartrackMorningMinute")) qs("cartrackMorningMinute").value = mm || "0";
    if (qs("cartrackSpeedAlertKmh")) qs("cartrackSpeedAlertKmh").value = String(s.speed_alert_kmh ?? 100);
    setCartrackAdminResult(
      s.configured ? `Configured (${s.source}). Updated ${s.updated_at || "—"}.` : "Not configured yet.",
      s.configured ? true : null
    );
  } catch (e) {
    setCartrackAdminResult(String(e.message || e), false);
  }
}

async function saveCartrackAdminSettings() {
  setCartrackAdminResult("Saving…", null);
  const body = {
    base_url: qs("cartrackBaseUrl")?.value,
    username: qs("cartrackUsername")?.value,
    morning_recipients: qs("cartrackMorningRecipients")?.value,
    morning_enabled: Boolean(qs("cartrackMorningEnabled")?.checked),
    morning_hour: Number(qs("cartrackMorningHour")?.value || 6),
    morning_minute: Number(qs("cartrackMorningMinute")?.value || 0),
    speed_alert_kmh: Number(qs("cartrackSpeedAlertKmh")?.value || 100),
  };
  const pass = String(qs("cartrackPassword")?.value || "").trim();
  if (pass) body.password = pass;
  const data = await fetchJson(`${API}/api/cartrack/settings`, {
    method: "PUT",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify(body),
  });
  if (qs("cartrackPassword")) qs("cartrackPassword").value = "";
  setCartrackAdminResult("Cartrack settings saved.", true);
  loadCartrackFleet().catch(() => {});
  return data;
}

async function testCartrackConnection() {
  setCartrackAdminResult("Testing connection…", null);
  try {
    const data = await fetchJson(`${API}/api/cartrack/test-connection`, { method: "POST" });
    setCartrackAdminResult(data.message || "Connected.", true);
  } catch (e) {
    setCartrackAdminResult(String(e.message || e), false);
  }
}

async function runCartrackMorningNow() {
  setCartrackAdminResult("Running morning report…", null);
  try {
    const data = await fetchJson(`${API}/api/cartrack/morning-report/run`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({ send_email: true }),
    });
    const n = data?.summary?.total_speeding_events ?? 0;
    setCartrackAdminResult(`Report for ${data.report_date}: ${n} speeding event(s).${data.emailed ? " Emailed." : ""}`, true);
    loadCartrackFleet().catch(() => {});
  } catch (e) {
    setCartrackAdminResult(String(e.message || e), false);
  }
}

async function initCartrackAdminPanel() {
  if (!qs("adminCartrackCard")) return;
  await loadCartrackAdminSettings().catch(() => {});
  await loadUnitechAdminSettings().catch(() => {});
  await loadGpsVehicleLinksAdmin().catch(() => {});
}

function setGpsVehicleLinksResult(text, ok = null) {
  const el = qs("gpsVehicleLinksResult");
  if (!el) return;
  el.textContent = String(text || "");
  el.style.color = ok === true ? "#15803d" : ok === false ? "#b91c1c" : "";
}

async function ensureGpsLinkAssetDatalist() {
  const datalist = qs("gpsLinkAssetCodeList");
  if (!datalist || datalist.dataset.loaded === "1") return;
  try {
    const assets = await fetchJson(`${API}/api/assets?include_archived=0`);
    const rows = Array.isArray(assets) ? assets : assets?.assets || [];
    datalist.innerHTML = rows
      .map((a) => {
        const code = String(a.asset_code || "").trim();
        const name = String(a.asset_name || "").trim();
        if (!code) return "";
        return `<option value="${escapeHtml(code)}">${escapeHtml(name ? `${code} — ${name}` : code)}</option>`;
      })
      .join("");
    datalist.dataset.loaded = "1";
  } catch {
    /* optional */
  }
}

function renderGpsVehicleLinksTable(links) {
  const host = qs("gpsVehicleLinksList");
  if (!host) return;
  if (!links?.length) {
    host.innerHTML = `<div class="muted small">No mappings yet. Add one above or link from the suggestions list.</div>`;
    return;
  }
  host.innerHTML = `
    <div class="cartrack-table-scroll">
      <table class="cartrack-fleet-table">
        <thead>
          <tr>
            <th>Registration</th>
            <th>Fleet code</th>
            <th>Source</th>
            <th>Notes</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          ${links.map((l) => `<tr>
            <td><code>${escapeHtml(l.registration)}</code></td>
            <td><strong>${escapeHtml(l.asset_code)}</strong></td>
            <td>${escapeHtml(l.gps_source || "any")}</td>
            <td class="muted">${escapeHtml(l.notes || "—")}</td>
            <td><button type="button" class="btn btn-secondary btn-sm" data-gps-link-delete="${escapeHtml(l.registration)}">Remove</button></td>
          </tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderGpsVehicleLinkSuggestions(suggestions) {
  const host = qs("gpsVehicleLinkSuggestions");
  if (!host) return;
  if (!suggestions?.length) {
    host.innerHTML = `<div class="muted small">All synced vehicles are mapped or already match a fleet code.</div>`;
    return;
  }
  host.innerHTML = `
    <div class="cartrack-table-scroll">
      <table class="cartrack-fleet-table">
        <thead>
          <tr>
            <th>Registration</th>
            <th>Current label</th>
            <th>Source</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          ${suggestions.map((s) => `<tr>
            <td><code>${escapeHtml(s.registration)}</code></td>
            <td>${escapeHtml(s.current_label || s.registration)}</td>
            <td>${escapeHtml(s.gps_source || "—")}</td>
            <td><button type="button" class="btn btn-primary btn-sm" data-gps-link-prefill="${escapeHtml(s.registration)}" data-gps-link-source="${escapeHtml(s.gps_source || "any")}">Link</button></td>
          </tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
}

async function loadGpsVehicleLinksAdmin() {
  if (!qs("adminGpsVehicleLinksCard")) return;
  await ensureGpsLinkAssetDatalist();
  try {
    const data = await fetchJson(`${API}/api/cartrack/vehicle-links`);
    renderGpsVehicleLinksTable(data?.links || []);
    renderGpsVehicleLinkSuggestions(data?.suggestions || []);
    setGpsVehicleLinksResult(
      `${(data?.links || []).length} mapping(s), ${(data?.suggestions || []).length} unmapped vehicle(s).`,
      true
    );
  } catch (e) {
    setGpsVehicleLinksResult(String(e.message || e), false);
  }
}

async function saveGpsVehicleLink() {
  const registration = String(qs("gpsLinkRegistration")?.value || "").trim();
  const asset_code = String(qs("gpsLinkAssetCode")?.value || "").trim();
  if (!registration || !asset_code) {
    setGpsVehicleLinksResult("Registration and fleet code are required.", false);
    return;
  }
  setGpsVehicleLinksResult("Saving…", null);
  try {
    const data = await fetchJson(`${API}/api/cartrack/vehicle-links`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({
        registration,
        asset_code,
        gps_source: qs("gpsLinkSource")?.value || "any",
        notes: qs("gpsLinkNotes")?.value || "",
      }),
    });
    setGpsVehicleLinksResult(`Saved ${data.link?.registration} → ${data.link?.asset_code}. Fleet updated.`, true);
    if (qs("gpsLinkRegistration")) qs("gpsLinkRegistration").value = "";
    if (qs("gpsLinkAssetCode")) qs("gpsLinkAssetCode").value = "";
    if (qs("gpsLinkNotes")) qs("gpsLinkNotes").value = "";
    await loadGpsVehicleLinksAdmin();
    loadCartrackFleet().catch(() => {});
    refreshCartrackSpeedFloat({ refresh: false }).catch(() => {});
  } catch (e) {
    setGpsVehicleLinksResult(String(e.message || e), false);
  }
}

async function deleteGpsVehicleLink(registration) {
  const reg = String(registration || "").trim();
  if (!reg) return;
  setGpsVehicleLinksResult("Removing…", null);
  try {
    await fetchJson(`${API}/api/cartrack/vehicle-links/${encodeURIComponent(reg)}`, {
      method: "DELETE",
      headers: authHeaders(),
    });
    setGpsVehicleLinksResult(`Removed mapping for ${reg}.`, true);
    await loadGpsVehicleLinksAdmin();
    loadCartrackFleet().catch(() => {});
  } catch (e) {
    setGpsVehicleLinksResult(String(e.message || e), false);
  }
}

async function applyGpsVehicleLinks() {
  setGpsVehicleLinksResult("Re-applying mappings…", null);
  try {
    const data = await fetchJson(`${API}/api/cartrack/vehicle-links/apply`, { method: "POST" });
    setGpsVehicleLinksResult(`Mappings re-applied to ${data.applied?.updated ?? 0} snapshot row(s).`, true);
    loadCartrackFleet().catch(() => {});
    refreshCartrackSpeedFloat({ refresh: false }).catch(() => {});
  } catch (e) {
    setGpsVehicleLinksResult(String(e.message || e), false);
  }
}

function prefillGpsVehicleLinkForm(registration, gpsSource) {
  if (qs("gpsLinkRegistration")) qs("gpsLinkRegistration").value = registration || "";
  if (qs("gpsLinkSource") && gpsSource) qs("gpsLinkSource").value = gpsSource;
  qs("gpsLinkAssetCode")?.focus();
}

function setUnitechAdminResult(text, ok = null) {
  const el = qs("unitechAdminResult");
  if (!el) return;
  el.textContent = String(text || "");
  el.style.color = ok === true ? "#15803d" : ok === false ? "#b91c1c" : "";
}

async function loadUnitechAdminSettings() {
  if (!qs("adminUnitechCard")) return;
  try {
    const data = await fetchJson(`${API}/api/unitech/settings`);
    const s = data?.settings || {};
    if (qs("unitechFeedLabel")) qs("unitechFeedLabel").value = s.feed_label || "Afungi (Unitech)";
    if (qs("unitechEnabled")) qs("unitechEnabled").checked = s.enabled !== false;
    setUnitechAdminResult(
      s.configured
        ? `Configured (${s.source}). KML URL saved. Updated ${s.updated_at || "—"}.`
        : "Paste your Unitech GpsGate KML feed URL and save.",
      s.configured ? true : null
    );
  } catch (e) {
    setUnitechAdminResult(String(e.message || e), false);
  }
}

async function saveUnitechAdminSettings() {
  setUnitechAdminResult("Saving…", null);
  const body = {
    feed_label: qs("unitechFeedLabel")?.value,
    enabled: Boolean(qs("unitechEnabled")?.checked),
  };
  const url = String(qs("unitechKmlFeedUrl")?.value || "").trim();
  if (url) body.kml_feed_url = url;
  const data = await fetchJson(`${API}/api/unitech/settings`, {
    method: "PUT",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify(body),
  });
  if (qs("unitechKmlFeedUrl")) qs("unitechKmlFeedUrl").value = "";
  setUnitechAdminResult("Unitech settings saved.", true);
  loadCartrackFleet().catch(() => {});
  return data;
}

async function testUnitechConnection() {
  setUnitechAdminResult("Testing KML feed…", null);
  try {
    const data = await fetchJson(`${API}/api/unitech/test-connection`, { method: "POST" });
    setUnitechAdminResult(data.message || "Connected.", true);
  } catch (e) {
    setUnitechAdminResult(String(e.message || e), false);
  }
}

const CARTRACK_TRACK_POLL_MS = 45000;
const CARTRACK_MAP_DEFAULT = { lat: -18.665695, lng: 35.529562, zoom: 6 };
let cartrackLeafletPromise = null;
let cartrackMap = null;
let cartrackMarkers = new Map();
let cartrackTrackPollTimer = null;
let cartrackTrackFleetCache = [];

function ensureLeafletLoaded() {
  if (window.L) return Promise.resolve();
  if (cartrackLeafletPromise) return cartrackLeafletPromise;
  cartrackLeafletPromise = new Promise((resolve, reject) => {
    if (!document.querySelector('link[data-leaflet-css]')) {
      const link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css";
      link.setAttribute("data-leaflet-css", "1");
      document.head.appendChild(link);
    }
    const script = document.createElement("script");
    script.src = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js";
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("Could not load map library"));
    document.head.appendChild(script);
  });
  return cartrackLeafletPromise;
}

function cartrackMarkerColor(v) {
  if (v.is_speeding) return "#dc2626";
  if (v.gps_source === "unitech") {
    if (v.position_stale) return "#a16207";
    return "#ea580c";
  }
  if (Number(v.ignition_on) === 1) return "#16a34a";
  return "#64748b";
}

function cartrackVehicleKey(v) {
  const id = String(v.registration || v.asset_code || "").trim();
  if (!id) return "";
  return v.gps_source === "unitech" ? `unitech:${id}` : id;
}

function cartrackSourcePillHtml(v) {
  if (v.gps_source === "unitech") {
    return `<span class="pill pill-orange cartrack-source-pill">${escapeHtml(v.gps_provider || "Unitech")}</span>`;
  }
  if (v.gps_source === "cartrack") {
    return `<span class="pill pill-blue cartrack-source-pill">Cartrack</span>`;
  }
  return "";
}

function normalizeGpsRegLabel(reg) {
  return String(reg || "").trim().toUpperCase().replace(/\s+/g, "");
}

function cartrackVehicleLabel(v) {
  const fleet = String(v.asset_code || "").trim();
  const reg = String(v.registration || "").trim();
  if (fleet) return fleet;
  return reg || "—";
}

function cartrackVehicleSubLabel(v) {
  const fleet = String(v.asset_code || "").trim();
  const reg = String(v.registration || "").trim();
  const name = String(v.vehicle_name || "").trim();
  if (fleet && reg && normalizeGpsRegLabel(fleet) !== normalizeGpsRegLabel(reg)) {
    return reg;
  }
  return name || reg || "";
}

function formatCartrackTelemetryLine(v) {
  const parts = [];
  if (v.ev_battery_pct != null) parts.push(`EV batt ${Math.round(v.ev_battery_pct)}%`);
  if (v.tracker_battery_pct != null) parts.push(`Tracker ${Math.round(v.tracker_battery_pct)}%`);
  if (v.supply_voltage_v != null) parts.push(`${Number(v.supply_voltage_v).toFixed(1)} V`);
  if (v.fuel_pct != null) parts.push(`Fuel ${Math.round(v.fuel_pct)}%`);
  if (v.charging_status) parts.push(String(v.charging_status));
  return parts.join(" · ");
}

function renderCartrackBatteryPillsHtml(v) {
  const pills = [];
  const low = cartrackBatteryLow(v);
  if (v.ev_battery_pct != null) {
    pills.push(`<span class="pill pill-blue cartrack-batt-pill">EV ${Math.round(v.ev_battery_pct)}%</span>`);
  }
  if (v.tracker_battery_pct != null) {
    const tracker = Math.round(Number(v.tracker_battery_pct));
    const cls = low && tracker > 0 && tracker < 25 ? "pill-red" : "pill-blue";
    pills.push(`<span class="pill ${cls} cartrack-batt-pill">Tracker ${tracker}%</span>`);
  }
  if (v.supply_voltage_v != null) {
    const volts = Number(v.supply_voltage_v);
    const vCls = low && volts > 0 ? "pill-orange" : "pill-gray";
    pills.push(`<span class="pill ${vCls} cartrack-batt-pill">${volts.toFixed(1)} V</span>`);
  }
  if (v.fuel_pct != null) {
    pills.push(`<span class="pill pill-gray cartrack-batt-pill">Fuel ${Math.round(v.fuel_pct)}%</span>`);
  }
  if (v.charging_status) {
    pills.push(`<span class="pill pill-green cartrack-batt-pill">${escapeHtml(String(v.charging_status))}</span>`);
  }
  return pills.join("");
}

function cartrackBatteryLow(v) {
  const tracker = Number(v.tracker_battery_pct);
  if (Number.isFinite(tracker) && tracker > 0 && tracker < 25) return true;
  const volts = Number(v.supply_voltage_v);
  if (!Number.isFinite(volts) || volts <= 0) return false;
  if (volts < 24) return volts < 11.5;
  return volts < 22;
}

function buildCartrackPopupHtml(v) {
  const isUnitech = v.gps_source === "unitech";
  const ign = isUnitech
    ? "—"
    : (Number(v.ignition_on) === 1 ? "ON" : "OFF");
  const spd = Number(v.speed_kmh || 0).toFixed(0);
  const odo = v.odometer_km != null ? `${Number(v.odometer_km).toFixed(0)} km` : "—";
  const when = String(v.last_event_at || v.synced_at || "").slice(0, 16) || "—";
  const lat = Number(v.latitude);
  const lng = Number(v.longitude);
  const mapsUrl = Number.isFinite(lat) && Number.isFinite(lng)
    ? `https://www.google.com/maps?q=${lat},${lng}`
    : "";
  const telemetry = formatCartrackTelemetryLine(v);
  const battLow = cartrackBatteryLow(v);
  const batteryPills = renderCartrackBatteryPillsHtml(v);
  const staleNote = v.position_stale
    ? `<div class="cartrack-stale-note">Position may be stale (${Number(v.position_age_days || 0).toFixed(0)} day(s) old)</div>`
    : "";
  return `
    <div class="cartrack-popup">
      <strong>${escapeHtml(cartrackVehicleLabel(v))}</strong>
      <div>${cartrackSourcePillHtml(v)} ${escapeHtml(cartrackVehicleSubLabel(v) || v.vehicle_name || v.registration || "")}</div>
      <div>Speed: <b>${spd}</b> km/h${isUnitech && v.road_speed_limit != null ? ` · Site max <b>${Number(v.road_speed_limit)}</b> km/h` : ""}${isUnitech ? "" : ` · Ignition: <b>${ign}</b>`}</div>
      ${isUnitech ? "" : `<div>Odometer: ${escapeHtml(odo)}</div>`}
      ${batteryPills ? `<div class="cartrack-track-item-battery ${battLow ? "cartrack-batt-low" : ""}">${batteryPills}</div>` : telemetry ? `<div class="${battLow ? "cartrack-batt-low" : ""}">Battery / power: <b>${escapeHtml(telemetry)}</b></div>` : ""}
      ${staleNote}
      <div class="muted">Last: ${escapeHtml(when)}</div>
      ${mapsUrl ? `<a href="${mapsUrl}" target="_blank" rel="noopener noreferrer">Open in Maps</a>` : ""}
    </div>
  `;
}

function ensureCartrackMap() {
  const host = qs("cartrackMap");
  if (!host || !window.L) return null;
  if (cartrackMap) {
    cartrackMap.invalidateSize();
    return cartrackMap;
  }
  cartrackMap = window.L.map(host, { zoomControl: true }).setView(
    [CARTRACK_MAP_DEFAULT.lat, CARTRACK_MAP_DEFAULT.lng],
    CARTRACK_MAP_DEFAULT.zoom
  );
  window.L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
  }).addTo(cartrackMap);
  return cartrackMap;
}

function fitCartrackMapToFleet(fleet) {
  if (!cartrackMap || !window.L) return;
  const pts = (fleet || [])
    .filter((v) => v.has_gps)
    .map((v) => [Number(v.latitude), Number(v.longitude)]);
  if (!pts.length) {
    cartrackMap.setView([CARTRACK_MAP_DEFAULT.lat, CARTRACK_MAP_DEFAULT.lng], CARTRACK_MAP_DEFAULT.zoom);
    return;
  }
  if (pts.length === 1) {
    cartrackMap.setView(pts[0], 14);
    return;
  }
  cartrackMap.fitBounds(window.L.latLngBounds(pts), { padding: [40, 40], maxZoom: 14 });
}

function updateCartrackMapMarkers(fleet) {
  const map = ensureCartrackMap();
  if (!map || !window.L) return;
  const seen = new Set();
  for (const v of fleet || []) {
    const key = cartrackVehicleKey(v);
    if (!key) continue;
    seen.add(key);
    if (!v.has_gps) {
      const old = cartrackMarkers.get(key);
      if (old) {
        map.removeLayer(old);
        cartrackMarkers.delete(key);
      }
      continue;
    }
    const lat = Number(v.latitude);
    const lng = Number(v.longitude);
    const color = cartrackMarkerColor(v);
    let marker = cartrackMarkers.get(key);
    const icon = window.L.divIcon({
      className: "cartrack-map-marker-wrap",
      html: `<span class="cartrack-map-marker" style="background:${color}"></span>`,
      iconSize: [18, 18],
      iconAnchor: [9, 9],
    });
    if (!marker) {
      marker = window.L.marker([lat, lng], { icon }).addTo(map);
      cartrackMarkers.set(key, marker);
    } else {
      marker.setLatLng([lat, lng]);
      marker.setIcon(icon);
    }
    marker.bindPopup(buildCartrackPopupHtml(v));
    marker._cartrackVehicle = v;
  }
  for (const [key, marker] of cartrackMarkers.entries()) {
    if (!seen.has(key)) {
      map.removeLayer(marker);
      cartrackMarkers.delete(key);
    }
  }
}

function renderCartrackTrackList(fleet, filterText = "") {
  const host = qs("cartrackTrackList");
  if (!host) return;
  const q = String(filterText || "").trim().toLowerCase();
  const rows = (fleet || []).filter((v) => {
    if (!q) return true;
    const hay = `${v.asset_code || ""} ${v.registration || ""} ${v.vehicle_name || ""}`.toLowerCase();
    return hay.includes(q);
  });
  if (!rows.length) {
    host.innerHTML = `<div class="cartrack-track-empty muted small">${q ? "No vehicles match your search." : "No vehicles synced yet. Use Refresh now or sync from the dashboard."}</div>`;
    return;
  }
  host.innerHTML = rows
    .map((v) => {
      const key = cartrackVehicleKey(v);
      const label = cartrackVehicleLabel(v);
      const sub = cartrackVehicleSubLabel(v);
      const ign = Number(v.ignition_on) === 1;
      const spd = Number(v.speed_kmh || 0).toFixed(0);
      const batteryPills = renderCartrackBatteryPillsHtml(v);
      const battLow = cartrackBatteryLow(v);
      const cls = [
        "cartrack-track-item",
        v.is_speeding ? "cartrack-track-item--alert" : "",
        !v.has_gps ? "cartrack-track-item--nogps" : "",
        battLow ? "cartrack-track-item--battlow" : "",
        v.position_stale ? "cartrack-track-item--stale" : "",
        v.gps_source === "unitech" ? "cartrack-track-item--unitech" : "",
      ].filter(Boolean).join(" ");
      const ignPill = v.gps_source === "unitech"
        ? ""
        : (ign ? '<span class="pill pill-green">ON</span>' : '<span class="pill">OFF</span>');
      return `<button type="button" class="${cls}" data-cartrack-key="${escapeHtml(key)}">
        <span class="cartrack-track-item-code">${escapeHtml(label)}</span>
        <span class="cartrack-track-item-meta">${cartrackSourcePillHtml(v)} ${escapeHtml(sub || v.vehicle_name || v.registration || "—")}</span>
        <span class="cartrack-track-item-stats">
          ${ignPill}
          <span>${spd} km/h</span>
          ${v.is_speeding ? '<span class="pill pill-red">Speeding</span>' : ""}
          ${!v.has_gps ? '<span class="pill">No GPS</span>' : ""}
          ${v.position_stale ? '<span class="pill pill-orange">Stale GPS</span>' : ""}
        </span>
        ${batteryPills ? `<span class="cartrack-track-item-battery">${batteryPills}</span>` : ""}
      </button>`;
    })
    .join("");
}

function focusCartrackVehicle(key) {
  const marker = cartrackMarkers.get(String(key || ""));
  if (!marker || !cartrackMap) return;
  cartrackMap.setView(marker.getLatLng(), Math.max(cartrackMap.getZoom(), 14));
  marker.openPopup();
  qs("cartrackTrackList")?.querySelectorAll(".cartrack-track-item").forEach((el) => {
    el.classList.toggle("active", el.getAttribute("data-cartrack-key") === key);
  });
}

function setCartrackMapStatus(text) {
  const el = qs("cartrackMapStatus");
  if (el) el.textContent = String(text || "");
}

function stopCartrackTrackPolling() {
  if (cartrackTrackPollTimer) {
    clearInterval(cartrackTrackPollTimer);
    cartrackTrackPollTimer = null;
  }
}

function startCartrackTrackPolling() {
  stopCartrackTrackPolling();
  if (!qs("cartrackAutoRefresh")?.checked) return;
  cartrackTrackPollTimer = setInterval(() => {
    if (document.querySelector("#tab-cartrack.show")) {
      loadCartrackTrackingTab({ refresh: true, quiet: true }).catch(() => {});
    }
  }, CARTRACK_TRACK_POLL_MS);
}

async function loadCartrackTrackingTab({ refresh = true, quiet = false } = {}) {
  if (!qs("tab-cartrack")) return;
  if (!quiet) setStatus("Loading fleet map…");
  setCartrackMapStatus("Loading positions…");
  try {
    await ensureLeafletLoaded();
    ensureCartrackMap();
    const q = refresh ? "refresh=1" : "refresh=0";
    const data = await fetchJson(`${API}/api/cartrack/live?${q}`);
    const fleet = data?.fleet || [];
    cartrackTrackFleetCache = fleet;
    const s = data?.summary || {};
    setText("cartrackTrackKpiTotal", Number(s.total_vehicles || 0));
    setText("cartrackTrackKpiGps", Number(s.with_gps || 0));
    setText("cartrackTrackKpiLive", Number(s.ignition_on || 0));
    setText("cartrackTrackKpiSpeeding", Number(s.speeding_today || 0));
    const syncBits = [String(s.last_sync || "—").slice(0, 16) || "—"];
    if (Number(s.unitech_vehicles || 0) > 0) syncBits.push(`${s.unitech_vehicles} Unitech`);
    if (Number(s.cartrack_vehicles || 0) > 0) syncBits.push(`${s.cartrack_vehicles} Cartrack`);
    setText("cartrackTrackKpiSync", syncBits.join(" · "));
    const search = qs("cartrackTrackSearch")?.value || "";
    renderCartrackTrackList(fleet, search);
    updateCartrackMapMarkers(fleet);
    if (!cartrackMap?._cartrackFittedOnce && fleet.some((v) => v.has_gps)) {
      fitCartrackMapToFleet(fleet);
      cartrackMap._cartrackFittedOnce = true;
    }
    const statusBits = [];
    if (!data.configured) statusBits.push("Cartrack not configured");
    else statusBits.push(`${Number(s.with_gps || 0)} on map`);
    if (data.sync_error) statusBits.push(`sync: ${data.sync_error}`);
    setCartrackMapStatus(statusBits.join(" · ") || "Ready");
    if (!quiet) setStatus(`Fleet map updated — ${Number(s.with_gps || 0)} vehicle(s) with GPS.`);
    startCartrackTrackPolling();
    loadCartrackSpeedingEvents(getCartrackSpeedReportDate(), {
      hostId: "cartrackTrackSpeedingEventsHost",
      countId: "cartrackTrackSpeedingEventsCount",
      panelId: "cartrackTrackSpeedingEventsPanel",
      useCache: getCartrackSpeedReportDate() === todayLocalYmd(),
    }).catch(() => {});
  } catch (e) {
    setCartrackMapStatus(String(e.message || e));
    if (!quiet) setStatus(`Fleet map error: ${e.message || e}`);
  }
}

function initCartrackTrackingTab() {
  qs("cartrackTrackRefreshBtn")?.addEventListener("click", () =>
    loadCartrackTrackingTab({ refresh: true }).catch((e) => setStatus(String(e.message || e)))
  );
  qs("cartrackTrackFitBtn")?.addEventListener("click", () => fitCartrackMapToFleet(cartrackTrackFleetCache));
  qs("cartrackOpenMapBtn")?.addEventListener("click", () => switchTab("cartrack"));
  qs("cartrackAutoRefresh")?.addEventListener("change", () => startCartrackTrackPolling());
  qs("cartrackTrackSearch")?.addEventListener("input", (e) => {
    renderCartrackTrackList(cartrackTrackFleetCache, e.target?.value || "");
  });
  qs("cartrackTrackList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-cartrack-key]");
    if (!btn) return;
    focusCartrackVehicle(btn.getAttribute("data-cartrack-key"));
  });
}

const CARTRACK_SPEED_FLOAT_POLL_MS = 60000;
let cartrackSpeedFloatPollTimer = null;
let cartrackSpeedFloatFleet = [];

function setCartrackSpeedFloatExpanded(expanded) {
  const root = qs("cartrackSpeedFloat");
  if (!root) return;
  root.classList.toggle("is-collapsed", !expanded);
  root.classList.toggle("is-expanded", expanded);
}

function renderCartrackSpeedFloatList(fleet) {
  const host = qs("cartrackSpeedFloatList");
  if (!host) return;
  const rows = [...(fleet || [])].sort((a, b) => {
    const sa = Number(a.speed_kmh || 0);
    const sb = Number(b.speed_kmh || 0);
    if (sb !== sa) return sb - sa;
    return Number(b.ignition_on) - Number(a.ignition_on);
  });
  if (!rows.length) {
    host.innerHTML = `<div class="cartrack-speed-float-empty muted small">No Cartrack vehicles synced.</div>`;
    return;
  }
  host.innerHTML = rows
    .map((v) => {
      const label = cartrackVehicleLabel(v);
      const sub = cartrackVehicleSubLabel(v);
      const spd = Number(v.speed_kmh || 0);
      const ign = Number(v.ignition_on) === 1;
      const limit = v.road_speed_limit != null ? Number(v.road_speed_limit) : null;
      const cls = [
        "cartrack-speed-float-row",
        v.is_speeding ? "cartrack-speed-float-row--alert" : "",
        ign ? "cartrack-speed-float-row--live" : "",
      ].filter(Boolean).join(" ");
      const extras = [];
      if (ign) extras.push("IGN");
      if (v.is_idling) extras.push("Idle");
      if (v.gps_source === "unitech" && limit) extras.push(`Max ${limit} km/h`);
      else if (limit) extras.push(`Limit ${limit}`);
      else if (v.speed_alert_kmh) extras.push(`Alert ≥${v.speed_alert_kmh}`);
      const batteryPills = renderCartrackBatteryPillsHtml(v);
      const battLow = cartrackBatteryLow(v);
      return `<div class="${cls}${battLow ? " cartrack-speed-float-row--battlow" : ""}">
        <span class="cartrack-speed-float-code">${escapeHtml(label)}</span>
        <span class="cartrack-speed-float-spd">${spd.toFixed(0)}<small> km/h</small></span>
        ${batteryPills ? `<span class="cartrack-speed-float-battery">${batteryPills}</span>` : `<span class="cartrack-speed-float-meta">${escapeHtml(extras.join(" · ") || sub || v.registration || "")}</span>`}
      </div>`;
    })
    .join("");
}

function updateCartrackSpeedFloat(data) {
  const root = qs("cartrackSpeedFloat");
  const badge = qs("cartrackSpeedFloatBadge");
  const updated = qs("cartrackSpeedFloatUpdated");
  if (!root) return;

  const allowed = getEffectiveAllowedTabs();
  if (!allowed.includes("cartrack") && !allowed.includes("dash")) {
    root.classList.add("hidden");
    return;
  }

  const fleet = Array.isArray(data?.fleet) ? data.fleet : [];
  cartrackSpeedFloatFleet = fleet;
  if (!data?.configured || !fleet.length) {
    root.classList.add("hidden");
    return;
  }

  root.classList.remove("hidden");
  const moving = fleet.filter((v) => Number(v.speed_kmh || 0) > 3 || Number(v.ignition_on) === 1).length;
  const speeding = fleet.filter((v) => v.is_speeding).length;
  if (badge) {
    badge.textContent = String(moving);
    badge.classList.toggle("cartrack-speed-float-badge--alert", speeding > 0);
  }
  if (updated) {
    const sync = String(data?.summary?.last_sync || "").slice(11, 16) || "—";
    updated.textContent = `${fleet.length} vehicles · ${moving} active · updated ${sync}`;
  }
  renderCartrackSpeedFloatList(fleet);
}

async function refreshCartrackSpeedFloat({ refresh = true } = {}) {
  const allowed = getEffectiveAllowedTabs();
  if (!allowed.includes("cartrack") && !allowed.includes("dash")) return;
  try {
    const q = refresh ? "refresh=1" : "refresh=0";
    const data = await fetchJson(`${API}/api/cartrack/live?${q}`);
    updateCartrackSpeedFloat(data);
  } catch (_) {
    /* silent */
  }
}

function initCartrackSpeedFloat() {
  const root = qs("cartrackSpeedFloat");
  if (!root) return;

  qs("cartrackSpeedFloatTab")?.addEventListener("click", () => {
    const expanded = root.classList.contains("is-expanded");
    setCartrackSpeedFloatExpanded(!expanded);
    if (!expanded) refreshCartrackSpeedFloat({ refresh: true }).catch(() => {});
  });
  qs("cartrackSpeedFloatMinBtn")?.addEventListener("click", () => setCartrackSpeedFloatExpanded(false));
  qs("cartrackSpeedFloatMapBtn")?.addEventListener("click", () => {
    switchTab("cartrack");
    loadCartrackTrackingTab({ refresh: true }).catch(() => {});
  });
  qs("cartrackSpeedFloatRefreshBtn")?.addEventListener("click", () =>
    refreshCartrackSpeedFloat({ refresh: true }).catch(() => {})
  );

  refreshCartrackSpeedFloat({ refresh: true }).catch(() => {});
  if (cartrackSpeedFloatPollTimer) clearInterval(cartrackSpeedFloatPollTimer);
  cartrackSpeedFloatPollTimer = setInterval(() => {
    refreshCartrackSpeedFloat({ refresh: true }).catch(() => {});
  }, CARTRACK_SPEED_FLOAT_POLL_MS);
}

async function loadDashboard() {
  const dateEl = qs("date");
  const scheduledEl = qs("scheduled");
  const date = dateEl ? dateEl.value : new Date().toISOString().slice(0, 10);
  const scheduled = scheduledEl ? scheduledEl.value || 10 : 10;

  setStatus("Loading dashboard...");
  setSkeleton("downtimeList", 2);
  setSkeleton("downtimeReasonsList", 2);
  setSkeleton("ldvPrestartList", 2);
  setSkeleton("stockList", 2);
  setSkeleton("woList", 2);
  setSkeleton("riskBoardList", 2);
  if (isDashSectionVisible("dashSlaCard")) setSkeleton("slaList", 2);
  if (isDashSectionVisible("dashSyncDiagCard")) setSkeleton("syncDiagList", 2);
  if (isDashSectionVisible("dashCostTrendCard")) setSkeleton("costTrendList", 2);
  setSkeleton("costList", 2);
  setSkeleton("lubeList", 2);
  if (isDashSectionVisible("dashStockMonitorCard")) setSkeleton("stockMonitorList", 2);
  setSkeleton("telematicsFleetList", 2);
  setSkeleton("cartrackFleetHost", 2);

  const data = await fetchJson(`${API}/api/dashboard?date=${date}&scheduled=${scheduled}`);
  loadTelematicsFleet().catch(() => {});
  loadCartrackFleet().catch(() => {});

  const sqDateEl = qs("sqDate");
  if (sqDateEl && !sqDateEl.value) sqDateEl.value = date;

  const _kpiTh = getThresholds();
  setSpeedo(qs("availNeedle"), qs("gAvailVal"), data?.kpi?.availability, { goodAt: _kpiTh.availTarget, warnAt: _kpiTh.availCrit });
  setSpeedo(qs("utilNeedle"), qs("gUtilVal"), data?.kpi?.utilization, { goodAt: _kpiTh.utilTarget, warnAt: _kpiTh.utilCrit });
  updateKpiAlertBanner(
    data?.kpi?.availability_mtd ?? data?.kpi?.availability,
    data?.kpi?.utilization_mtd ?? data?.kpi?.utilization
  );

  const mtdRange =
    data.kpi?.mtd_start && data.kpi?.mtd_end
      ? `${data.kpi.mtd_start} → ${data.kpi.mtd_end}`
      : "";
  const k = data.kpi || {};
  const siteTag = k.site_code ? ` · Site: ${k.site_code}` : "";
  const gaugeNote =
    k.gauge_basis === "mtd"
      ? "Gauges use MTD (no planned hour-meter time on the selected day). "
      : `Gauges use selected day (${data.date || ""}). `;
  setText(
    "availMeta",
    `${gaugeNote}${siteTag}` +
      (mtdRange
        ? `MTD ${mtdRange} · Distinct assets (MTD): ${k.used_assets ?? "—"} | Planned−down (MTD) hrs: ${k.available_hours ?? "—"} | Downtime (MTD): ${k.downtime_hours ?? "—"}`
        : `Distinct assets (MTD): ${k.used_assets ?? "—"} | Planned−down (MTD) hrs: ${k.available_hours ?? "—"} | Downtime (MTD): ${k.downtime_hours ?? "—"}`)
  );
  setText(
    "utilMeta",
    `Day planned hrs: ${Number(k.scheduled_hours_day ?? 0).toFixed(1)} · Day run hrs: ${Number(k.run_hours ?? 0).toFixed(1)}. ` +
      (mtdRange
        ? `MTD ${mtdRange} · Run (MTD): ${Number(k.run_hours_mtd ?? k.run_hours ?? 0).toFixed(1)} | Planned (MTD): ${Number(k.utilization_base_hours || 0).toFixed(1)} | MTD util: ${k.utilization_mtd != null ? `${Number(k.utilization_mtd).toFixed(2)}%` : "—"} | Scheduled/asset (header): ${data.scheduled_hours_per_asset}`
        : `Run (MTD): ${Number(k.run_hours_mtd ?? k.run_hours ?? 0).toFixed(1)} | Planned (MTD): ${Number(k.utilization_base_hours || 0).toFixed(1)} | Scheduled/asset: ${data.scheduled_hours_per_asset}`)
  );
  const debugToggle = qs("kpiDebugToggle");
  const debugList = qs("kpiDebugList");
  if (debugList) {
    const show = Boolean(debugToggle?.checked);
    debugList.style.display = show ? "" : "none";
    debugList.innerHTML = "";
    if (show) {
      const rows = Array.isArray(data.per_asset_kpi) ? data.per_asset_kpi : [];
      rows.forEach((r) => {
        const mode = String(r.utilization_mode || "hours").toLowerCase();
        const meterTxt = mode === "km"
          ? `Meter: ${Number(r.meter_run_value || 0).toFixed(2)} km`
          : `Meter: ${Number(r.meter_run_value || r.run_hours || 0).toFixed(2)} h`;
        debugList.appendChild(
          item(
            `<b>${r.asset_code || `ID ${r.asset_id}`}</b>` +
            `<br><small>Mode: ${mode.toUpperCase()}${mode === "km" ? ` (km/h factor ${Number(r.km_per_hour_factor || 10).toFixed(2)})` : ""} | ${meterTxt}</small>` +
            `<br><small>Sched: ${Number(r.scheduled_hours || 0).toFixed(2)} | Down: ${Number(r.downtime_hours || 0).toFixed(2)} | Avail: ${Number(r.available_hours || 0).toFixed(2)} | Run(H): ${Number(r.run_hours || 0).toFixed(2)}${r.contributes_to_kpi === false ? " | KPI: EXCLUDED" : ""}</small>`
          )
        );
      });
      if (!rows.length) {
        debugList.appendChild(
          item(
            "<small>No hour-meter production assets for this site/date (per-asset rows = selected day; gauges = selected day; MTD figures in subtitles).</small>"
          )
        );
      }
    }
  }

  setText("aLowStock", data.alerts.low_stock);
  setText("aOverdue", data.alerts.overdue_maintenance);
  setText("aOpenWO", data.alerts.open_work_orders);

  const downtimeList = qs("downtimeList");
  if (downtimeList) {
    downtimeList.innerHTML = "";
    (data.major_downtime || []).forEach((r) => {
      downtimeList.appendChild(
        item(
          `<b>${r.asset_code}</b> – ${r.downtime_hours}h ${
            r.critical ? " <span class='pill red'>CRIT</span>" : ""
          }<br><small>${r.description}</small>`
        )
      );
    });
    if (!data.major_downtime?.length) downtimeList.appendChild(item("<small>No downtime recorded for this date.</small>"));
  }

  const reasonsList = qs("downtimeReasonsList");
  if (reasonsList) {
    reasonsList.innerHTML = "";
    (data.downtime_reasons || []).forEach((r) => {
      reasonsList.appendChild(
        item(`<b>${r.reason}</b> – ${r.hours_down}h<br><small>Incidents: ${r.incidents}</small>`)
      );
    });
    if (!data.downtime_reasons?.length) {
      reasonsList.appendChild(item("<small>No downtime reasons logged for this date.</small>"));
    }
  }

  const ldvPrestartList = qs("ldvPrestartList");
  if (ldvPrestartList) {
    ldvPrestartList.innerHTML = "";
    const badge = qs("ldvPrestartBadge");
    const card = qs("ldvPrestartCard");
    const compliantPill = qs("ldvPrestartCompliantPill");
    const missingPill = qs("ldvPrestartMissingPill");
    const pctPill = qs("ldvPrestartPctPill");
    const setComplianceTone = (pct) => {
      const p = Number(pct || 0);
      const th = getLdvPrestartThresholds();
      const tone = p >= th.greenAt ? "green" : p >= th.warnAt ? "orange" : "red";
      if (badge) {
        badge.className = `dash-card-badge${tone === "red" ? " dash-card-badge-red" : ""}`;
        badge.textContent = tone === "green" ? "On Track" : tone === "orange" ? "Attention" : "Critical";
      }
      if (card) {
        card.style.borderColor = tone === "green"
          ? "rgba(34,197,94,0.55)"
          : tone === "orange"
            ? "rgba(245,158,11,0.55)"
            : "rgba(239,68,68,0.55)";
      }
      if (compliantPill) compliantPill.className = `kpi-pill ${tone === "green" ? "kpi-pill-green" : "kpi-pill-blue"}`;
      if (missingPill) missingPill.className = `kpi-pill ${tone === "red" ? "kpi-pill-red" : "kpi-pill-orange"}`;
      if (pctPill) pctPill.className = `kpi-pill ${tone === "green" ? "kpi-pill-green" : tone === "orange" ? "kpi-pill-orange" : "kpi-pill-red"}`;
    };
    try {
      const ps = await fetchJson(`${API}/api/dashboard/ldv-prestart/compliance?date=${encodeURIComponent(date)}`);
      const summary = ps?.summary || {};
      setText("ldvPrestartCompliant", Number(summary.compliant || 0));
      setText("ldvPrestartMissing", Number(summary.missing || 0));
      setText("ldvPrestartPct", `${Number(summary.pct || 0).toFixed(1)}%`);
      setComplianceTone(Number(summary.pct || 0));
      const rows = Array.isArray(ps?.rows) ? ps.rows : [];
      const attentionRows = rows.filter((r) => r.status !== "compliant");
      const renderRows = attentionRows.length ? attentionRows : rows.slice(0, 5);
      renderRows.forEach((r) => {
        const statusPill = r.status === "compliant"
          ? "<span class='pill green'>OK</span>"
          : "<span class='pill orange'>PENDING</span>";
        ldvPrestartList.appendChild(
          item(
            `<b>${escapeHtml(String(r.asset_code || "-"))}</b> ${statusPill}` +
            `<br><small>${escapeHtml(String(r.reason || ""))}</small>`
          )
        );
      });
      if (!rows.length) {
        ldvPrestartList.appendChild(item("<small>No LDV assets found for V01AM-V15AM.</small>"));
      }
    } catch (e) {
      setText("ldvPrestartCompliant", "0");
      setText("ldvPrestartMissing", "0");
      setText("ldvPrestartPct", "-");
      setComplianceTone(0);
      if (badge) {
        badge.className = "dash-card-badge dash-card-badge-red";
        badge.textContent = "Unavailable";
      }
      ldvPrestartList.appendChild(
        item(`<small>Pre-start compliance unavailable: ${escapeHtml(String(e.message || e))}</small>`)
      );
    }
  }

  const stockList = qs("stockList");
  if (stockList) {
    stockList.innerHTML = "";
    (data.critical_low_stock || []).forEach((r) => {
      stockList.appendChild(
        item(`<b>${r.part_code}</b> – ${r.on_hand} on hand<br><small>${r.part_name} | Min: ${r.min_stock}</small>`)
      );
    });
    if (!data.critical_low_stock?.length) stockList.appendChild(item("<small>No critical low stock.</small>"));
  }

  const woList = qs("woList");
  if (woList) {
    woList.innerHTML = "";
    const isStrictOpenWO = (r) => {
      const norm = String(r?.status || "").trim().toLowerCase().replace(/\s+/g, "_");
      const completedAt = String(r?.completed_at || "").trim();
      const closedAt = String(r?.closed_at || "").trim();
      return ["open", "assigned", "in_progress"].includes(norm) && !completedAt && !closedAt;
    };
    const openRows = (data.open_work_orders || []).filter(isStrictOpenWO);
    setText("aOpenWO", openRows.length);
    openRows.forEach((r) => {
      woList.appendChild(
        item(`<b>WO #${r.id}</b> – ${r.asset_code}<br><small>${r.source} | ${r.status} | ${r.opened_at}</small>`)
      );
    });
    if (!openRows.length) woList.appendChild(item("<small>No open work orders.</small>"));
  }

  const riskBoardList = qs("riskBoardList");
  if (riskBoardList) {
    riskBoardList.innerHTML = "";
    try {
      const rb = await fetchJson(`${API}/api/ironmind/risk-board?date=${encodeURIComponent(date)}&limit=8`);
      const rows = Array.isArray(rb?.items) ? rb.items : [];
      rows.forEach((r) => {
        const reasons = Array.isArray(r.reasons) ? r.reasons.slice(0, 2).join(" | ") : "";
        riskBoardList.appendChild(
          item(
            `<b>${escapeHtml(r.asset_code || "-")}</b> - Risk ${Number(r.risk_score || 0).toFixed(0)}/100` +
            ` <span class="pill orange">Conf ${Number(r.confidence || 0).toFixed(0)}%</span>` +
            (reasons ? `<br><small>${escapeHtml(reasons)}</small>` : "") +
            `<br><button data-ironmind-risk-asset="${escapeHtml(r.asset_code || "")}">Open Asset History</button> ` +
            `<button data-ironmind-risk-wo="${escapeHtml(r.asset_code || "")}">Create WO</button>`
          )
        );
      });
      if (!rows.length) riskBoardList.appendChild(item("<small>No risk-board data yet. Refresh IronMind insight first.</small>"));
    } catch (e) {
      riskBoardList.appendChild(item(`<small>Risk board unavailable: ${escapeHtml(e.message || String(e))}</small>`));
    }
  }

  const slaList = qs("slaList");
  if (isDashSectionVisible("dashSlaCard") && slaList) {
  const sla = data.workorder_sla || {};
  const slaSummary = sla.summary || {};
  setText("slaOpen24", Number(slaSummary.open_gt_24h || 0));
  setText("slaProgress48", Number(slaSummary.in_progress_gt_48h || 0));
  setText("slaCompleted12", Number(slaSummary.completed_gt_12h || 0));
  if (slaList) {
    slaList.innerHTML = "";
    (sla.breaches || []).forEach((r) => {
      const p = String(r.priority || "P3").toUpperCase();
      const pClass = p === "P1" ? "pri-p1" : p === "P2" ? "pri-p2" : "pri-p3";
      const s = String(r.status || "").toLowerCase();
      const actionBtn =
        s === "open"
          ? `<button data-sla-set-id="${r.id}" data-sla-set-status="assigned">Assign Now</button>`
          : s === "assigned"
          ? `<button data-sla-set-id="${r.id}" data-sla-set-status="in_progress">Start Now</button>`
          : s === "completed"
          ? `<button data-sla-set-id="${r.id}" data-sla-set-status="approved">Approve Now</button>`
          : "";
      slaList.appendChild(
        item(
          `<b>WO #${r.id}</b> - ${r.asset_code} (${r.status})` +
          ` <span class="pill ${pClass}">${p}</span>` +
          `<br><small>Age: ${Number(r.age_hours || 0)}h | Source: ${r.source || "-"} | Opened: ${r.opened_at || "-"}</small>` +
          `<br>${actionBtn} <button data-sla-nudge-id="${r.id}">Nudge Supervisor</button> <button data-sla-open-id="${r.id}">Open WO</button>`
        )
      );
    });
    if (!sla.breaches?.length) slaList.appendChild(item("<small>No SLA breaches right now.</small>"));
  }
  }

  if (isDashSectionVisible("dashSlaCard") && slaList && !slaList.dataset.bound) {
    slaList.dataset.bound = "1";
    slaList.addEventListener("click", async (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const setId = target.getAttribute("data-sla-set-id");
      const setStatus = target.getAttribute("data-sla-set-status");
      const nudgeId = target.getAttribute("data-sla-nudge-id");
      const openId = target.getAttribute("data-sla-open-id");

      try {
        if (setId && setStatus) {
          setStatus(`Updating WO #${setId} -> ${setStatus}...`);
          await transitionWorkOrderStatus(setId, setStatus);
          await loadDashboard();
          setStatus(`WO #${setId} moved to ${setStatus}.`);
          return;
        }
        if (nudgeId) {
          setStatus(`Sending supervisor nudge for WO #${nudgeId}...`);
          await nudgeSupervisor(nudgeId);
          setStatus(`Nudge sent for WO #${nudgeId}.`);
          return;
        }
        if (openId) {
          const url = `/web/workorders.html?wo=${encodeURIComponent(openId)}`;
          if (getSlaOpenSameTab()) {
            window.location.href = url;
          } else {
            window.open(url, "_blank");
          }
        }
      } catch (e) {
        setStatus(`SLA action failed: ${e.message || e}`);
      }
    });
  }

  const syncDiagList = qs("syncDiagList");
  const syncDiagTrend = qs("syncDiagTrend");
  if (isDashSectionVisible("dashSyncDiagCard") && syncDiagList) {
    syncDiagList.innerHTML = "";
    try {
      const sd = await fetchJson(`${API}/api/sync/diagnostics`);
      const tableRows = Array.isArray(sd?.outbox_unsynced_by_table) ? sd.outbox_unsynced_by_table : [];
      const errRows = Array.isArray(sd?.outbox_error_breakdown) ? sd.outbox_error_breakdown : [];
      const peers = Array.isArray(sd?.checkpoints) ? sd.checkpoints : [];
      const unsyncedTotal = tableRows.reduce((sum, r) => sum + Number(r?.count || 0), 0);
      const outboxErrors = errRows
        .filter((r) => String(r?.error_text || "").trim() !== "" && String(r?.error_text || "") !== "(none)")
        .reduce((sum, r) => sum + Number(r?.count || 0), 0);
      const peerCount = new Set(peers.map((r) => String(r?.peer_name || "").trim()).filter(Boolean)).size;
      setText("sdOutboxUnsynced", String(unsyncedTotal));
      setText("sdOutboxErrors", String(outboxErrors));
      setText("sdPeers", String(peerCount));

      // Keep a short local trend history to show direction.
      const trendKey = "ironlog_sync_diag_trend_v1";
      let trendRows = [];
      try {
        const raw = localStorage.getItem(trendKey);
        const parsed = raw ? JSON.parse(raw) : [];
        trendRows = Array.isArray(parsed) ? parsed : [];
      } catch {
        trendRows = [];
      }
      trendRows.push({
        t: new Date().toISOString(),
        unsynced: Number(unsyncedTotal || 0),
        errors: Number(outboxErrors || 0),
      });
      trendRows = trendRows.slice(-12);
      try {
        localStorage.setItem(trendKey, JSON.stringify(trendRows));
      } catch {
        // ignore localStorage write failures
      }

      tableRows.slice(0, 5).forEach((r) => {
        syncDiagList.appendChild(
          item(`<b>${escapeHtml(r.table_name || "-")}</b> — ${Number(r.count || 0)} pending`)
        );
      });
      errRows
        .filter((r) => String(r?.error_text || "").trim() !== "" && String(r?.error_text || "") !== "(none)")
        .slice(0, 3)
        .forEach((r) => {
          syncDiagList.appendChild(
            item(`<small>Error: ${escapeHtml(String(r.error_text || "-"))} (${Number(r.count || 0)})</small>`)
          );
        });
      if (!tableRows.length) {
        syncDiagList.appendChild(item("<small>No outbox backlog detected.</small>"));
      }

      if (syncDiagTrend) {
        const maxUnsynced = Math.max(1, ...trendRows.map((r) => Number(r.unsynced || 0)));
        const maxErrors = Math.max(1, ...trendRows.map((r) => Number(r.errors || 0)));
        const pointsUnsynced = trendRows
          .map((r, i) => {
            const x = trendRows.length <= 1 ? 0 : (i / (trendRows.length - 1)) * 100;
            const y = 100 - (Number(r.unsynced || 0) / maxUnsynced) * 100;
            return `${x.toFixed(2)},${y.toFixed(2)}`;
          })
          .join(" ");
        const pointsErrors = trendRows
          .map((r, i) => {
            const x = trendRows.length <= 1 ? 0 : (i / (trendRows.length - 1)) * 100;
            const y = 100 - (Number(r.errors || 0) / maxErrors) * 100;
            return `${x.toFixed(2)},${y.toFixed(2)}`;
          })
          .join(" ");
        syncDiagTrend.innerHTML = `
          <div class="row" style="justify-content:space-between; align-items:center; margin-bottom:6px;">
            <small class="muted">Backlog trend (last ${trendRows.length} samples)</small>
            <small class="muted">Unsynced <span style="color:#2563eb;">●</span> Errors <span style="color:#dc2626;">●</span></small>
          </div>
          <svg viewBox="0 0 100 100" preserveAspectRatio="none" style="width:100%; height:90px; background:#f8fafc; border:1px solid #e5e7eb; border-radius:8px;">
            <polyline fill="none" stroke="#2563eb" stroke-width="2.2" points="${pointsUnsynced}"></polyline>
            <polyline fill="none" stroke="#dc2626" stroke-width="2.2" points="${pointsErrors}"></polyline>
          </svg>
        `;
      }
    } catch (e) {
      setText("sdOutboxUnsynced", "0");
      setText("sdOutboxErrors", "0");
      setText("sdPeers", "0");
      const msg = String(e?.message || "");
      if (msg.toLowerCase().includes("403")) {
        syncDiagList.appendChild(item("<small>Sync diagnostics available to admin/supervisor roles.</small>"));
      } else {
        syncDiagList.appendChild(item(`<small>Sync diagnostics unavailable: ${escapeHtml(msg || "unknown error")}</small>`));
      }
      if (syncDiagTrend) syncDiagTrend.innerHTML = "";
    }
  }

  const refreshSyncBtn = qs("refreshSyncDiagnostics");
  if (isDashSectionVisible("dashSyncDiagCard") && refreshSyncBtn && !refreshSyncBtn.dataset.bound) {
    refreshSyncBtn.dataset.bound = "1";
    refreshSyncBtn.addEventListener("click", () => {
      loadDashboard().catch((e) => setStatus(`Dashboard reload failed: ${e.message || e}`));
    });
  }

  const relDays = Number(qs("relDays")?.value || 30);
  const relRange = getLastNDaysRange(date, relDays);
  const rel = await fetchJson(
    `${API}/api/dashboard/reliability?start=${encodeURIComponent(relRange.start)}&end=${encodeURIComponent(relRange.end)}`
  );
  setText("relMtbf", rel.mtbf_hours == null ? "-" : Number(rel.mtbf_hours).toFixed(2));
  setText("relLttr", rel.lttr_hours == null ? "-" : Number(rel.lttr_hours).toFixed(2));
  setText("relFailures", String(Number(rel.failure_count || 0)));
  setText("relWindow", `Window: ${relRange.start} to ${relRange.end} (${Math.max(1, relDays)} days)`);
  const relTrend = await fetchJson(
    `${API}/api/dashboard/reliability/trend?weeks=12&end=${encodeURIComponent(date)}`
  );
  const relPoints = Array.isArray(relTrend?.points) ? relTrend.points : [];
  const relChart = qs("relTrendChart");
  const relList = qs("relTrendList");
  if (relChart) {
    relChart.innerHTML = "";
    const maxMtbf = Math.max(1, ...relPoints.map((p) => Number(p.mtbf_hours || 0)));
    relPoints.forEach((p) => {
      const bar = document.createElement("div");
      bar.className = "cost-bar";
      const h = Math.max(6, Math.round((Number(p.mtbf_hours || 0) / maxMtbf) * 100));
      bar.style.height = `${h}px`;
      bar.title = `${p.start} to ${p.end} | MTBF ${p.mtbf_hours ?? "-"} | LTTR ${p.lttr_hours ?? "-"} | Failures ${p.failure_count || 0}`;
      bar.innerHTML =
        `<span class="cost-bar-value">${p.mtbf_hours == null ? "-" : Number(p.mtbf_hours).toFixed(1)}</span>` +
        `<span class="cost-bar-label">${p.label || ""}</span>`;
      relChart.appendChild(bar);
    });
    if (!relPoints.length) relChart.appendChild(item("<small>No reliability trend data.</small>"));
  }
  if (relList) {
    relList.innerHTML = "";
    relPoints.slice(-6).reverse().forEach((p) => {
      relList.appendChild(
        item(
          `<b>${p.start} to ${p.end}</b>` +
          `<br><small>MTBF: ${p.mtbf_hours == null ? "-" : Number(p.mtbf_hours).toFixed(2)} | LTTR: ${p.lttr_hours == null ? "-" : Number(p.lttr_hours).toFixed(2)} | Failures: ${Number(p.failure_count || 0)}</small>`
        )
      );
    });
  }

  let trendRows = [];
  if (isDashSectionVisible("dashCostTrendCard")) {
  const trend = await fetchJson(`${API}/api/dashboard/cost/trend?months=12`);
  trendRows = Array.isArray(trend.rows) ? trend.rows : [];
  const mom = trend.mom || {};
  setText("ctCurrentMonth", trend.latest?.month || "-");
  setText("ctCurrentTotal", fmtMoney(trend.latest?.total_cost || 0));
  if (mom.variance == null) {
    setText("ctMoM", "N/A");
  } else {
    const pct = mom.variance_pct == null ? "" : ` (${Number(mom.variance_pct).toFixed(1)}%)`;
    setText("ctMoM", `${Number(mom.variance) >= 0 ? "+" : ""}${fmtMoney(mom.variance)}${pct}`);
  }
  const trendList = qs("costTrendList");
  const trendChart = qs("costTrendChart");
  if (trendChart) {
    trendChart.innerHTML = "";
    const maxCost = trendRows.reduce((m, r) => Math.max(m, Number(r.total_cost || 0)), 0);
    trendRows.forEach((r, idx) => {
      const total = Number(r.total_cost || 0);
      const h = maxCost > 0 ? Math.max(8, Math.round((total / maxCost) * 92)) : 8;
      const bar = document.createElement("div");
      const prev = idx > 0 ? Number(trendRows[idx - 1]?.total_cost || 0) : null;
      let trendClass = "neutral";
      if (prev != null && Number.isFinite(prev)) {
        if (total > prev) trendClass = "up";
        else if (total < prev) trendClass = "down";
      }
      bar.className = `cost-bar ${trendClass}`;
      bar.style.height = `${h}px`;
      bar.title = `${r.month}: ${fmtMoney(total)}`;

      const monthLabel = document.createElement("span");
      monthLabel.className = "cost-bar-label";
      monthLabel.textContent = String(r.month || "").slice(5);

      const valueLabel = document.createElement("span");
      valueLabel.className = "cost-bar-value";
      valueLabel.textContent = fmtMoney(total);

      bar.appendChild(monthLabel);
      bar.appendChild(valueLabel);
      trendChart.appendChild(bar);
    });
    if (!trendRows.length) trendChart.innerHTML = "<small class='muted'>No trend data.</small>";
  }
  if (trendList) {
    trendList.innerHTML = "";
    trendRows.slice().reverse().forEach((r) => {
      trendList.appendChild(
        item(
          `<b>${r.month}</b> - ${fmtMoney(r.total_cost)}` +
          `<br><small>Fuel ${fmtMoney(r.fuel_cost)} | Lube ${fmtMoney(r.lube_cost)} | Parts ${fmtMoney(r.parts_cost)} | Labor ${fmtMoney(r.labor_cost)} | Down ${fmtMoney(r.downtime_cost)}</small>`
        )
      );
    });
    if (!trendRows.length) trendList.appendChild(item("<small>No monthly cost trend data.</small>"));
  }
  }

  const costs = data.cost_engine || {};
  setText("cTotalCost", fmtMoney(costs.total_cost));
  setText("cCostPerHour", costs.cost_per_run_hour == null ? "N/A" : fmtMoney(costs.cost_per_run_hour));
  setText("cLaborHours", Number(costs.labor_hours || 0).toFixed(1));
  setText("cFuelCost", fmtMoney(costs.fuel_cost));
  setText("cLubeCost", fmtMoney(costs.lube_cost));
  setText("cPartsCost", fmtMoney(costs.parts_cost));
  setText("cLaborCost", fmtMoney(costs.labor_cost));
  setText("cDowntimeCost", fmtMoney(costs.downtime_cost));

  const costList = qs("costList");
  if (costList) {
    costList.innerHTML = "";
    (costs.top_asset_costs || []).forEach((r) => {
      costList.appendChild(
        item(
          `<b>${r.asset_code}</b> - ${fmtMoney(r.total_cost)}` +
          `<br><small>${r.asset_name || ""} | Fuel ${fmtMoney(r.fuel_cost)} | Lube ${fmtMoney(r.lube_cost)} | Parts ${fmtMoney(r.parts_cost)} | Labor ${fmtMoney(r.labor_cost)} | Down ${fmtMoney(r.downtime_cost)}</small>`
        )
      );
    });
    if (!costs.top_asset_costs?.length) costList.appendChild(item("<small>No cost activity for this date.</small>"));
  }

  const lube = data.lube_usage || {};
  setText("lubeQtyTotal", Number(lube.qty_total || 0).toFixed(1));
  setText("lubeEntries", Array.isArray(lube.rows) ? lube.rows.length : 0);
  setText("lubeAssets", Array.isArray(lube.rows) ? lube.rows.length : 0);
  setText("cLubeCost", fmtMoney(lube.total_lube_cost != null ? lube.total_lube_cost : data?.cost_engine?.lube_cost));
  lubeUsageCache = {
    rows: Array.isArray(lube.rows) ? lube.rows.map((r) => ({
      ...r,
      qty_total: Number(r.qty ?? 0),
      total_lube_cost: Number(r.total_lube_cost ?? r.lube_cost ?? 0),
    })) : [],
  };
  renderLubeUsageTable(lubeUsageCache);

  if (isDashSectionVisible("dashStockMonitorCard")) {
    await loadStockMonitor().catch(() => {});
  }
  await loadIronmindInsight({ silent: true }).catch(() => {});
  await loadIronmindHealth().catch(() => {});
  await loadIronmindSettings().catch(() => {});
  await loadIronmindHistory({ silent: true }).catch(() => {});

  setStatus("Dashboard ready.");
}

async function loadIronmindInsight(options = {}) {
  const silent = Boolean(options.silent);
  const summaryEl = qs("ironmindSummary");
  const metaEl = qs("ironmindMeta");

  try {
    const res = await fetchJson(`${API}/api/ironmind/latest?report_type=daily_admin`);
    const report = res?.report || null;
    if (!report) {
      if (summaryEl) {
        const emptySections = parseIronmindSections(
          ["IRONMIND DAILY INSIGHT", "", "Repairs Needed", "- Insufficient data",
           "", "Operational Risks", "- Insufficient data",
           "", "Suggestions", "- Insufficient data",
           "", "Data Gaps", "- Insufficient data",
           "", "Data Anomalies", "- Insufficient data"].join("\n")
        );
        renderIronmindSections(summaryEl, emptySections);
      }
      if (metaEl) metaEl.textContent = "No report generated yet.";
      if (!silent) setStatus("IRONMIND insight not available yet.");
      return;
    }

    renderIronmindReport(report);
    if (!silent) setStatus("IRONMIND insight loaded.");
  } catch (err) {
    if (metaEl) metaEl.textContent = "Insight unavailable right now.";
    if (!silent) setStatus("IRONMIND load error: " + err.message);
    throw err;
  }
}

async function loadIronmindHealth() {
  const healthEl = qs("ironmindHealth");
  if (!healthEl) return;
  try {
    const data = await fetchJson(`${API}/api/ironmind/health`);
    const live = Boolean(data?.live_enabled);
    const provider = String(data?.provider || "none");
    const model = String(data?.model || "");
    const mode = String(data?.last_ask_mode || "unknown");
    const err = String(data?.last_ask_error || "").trim();
    const left = live ? `OpenAI: Connected (${provider}${model ? `/${model}` : ""})` : `OpenAI: Not connected (${provider})`;
    const right = `Last ask mode: ${mode}${err ? ` | Last error: ${err}` : ""}`;
    healthEl.textContent = `${left} | ${right}`;
    healthEl.className = live ? "status-ok" : "status-overdue";
  } catch (e) {
    healthEl.textContent = `OpenAI status unavailable: ${e.message || e}`;
    healthEl.className = "status-overdue";
  }
}

async function loadIronmindSettings() {
  const meta = qs("ironmindSettingsMeta");
  try {
    const res = await fetchJson(`${API}/api/ironmind/settings`);
    const s = res?.settings || {};
    const setVal = (id, v) => {
      const el = qs(id);
      if (el && Number.isFinite(Number(v))) el.value = Number(v);
    };
    setVal("ironmindMaxRunGlobal", s.max_daily_run_hours);
    setVal("ironmindMaxRunLdv", s.max_daily_run_hours_ldv);
    setVal("ironmindMaxRunTruck", s.max_daily_run_hours_truck);
    setVal("ironmindMaxRunHeavy", s.max_daily_run_hours_heavy);
    if (meta) meta.textContent = "Thresholds loaded.";
  } catch {
    if (meta) meta.textContent = "Threshold load failed.";
  }
}

async function saveIronmindSettings() {
  const meta = qs("ironmindSettingsMeta");
  const readNum = (id, d) => {
    const n = Number(qs(id)?.value);
    return Number.isFinite(n) && n > 0 ? n : d;
  };
  const payload = {
    max_daily_run_hours: readNum("ironmindMaxRunGlobal", 24),
    max_daily_run_hours_ldv: readNum("ironmindMaxRunLdv", 16),
    max_daily_run_hours_truck: readNum("ironmindMaxRunTruck", 18),
    max_daily_run_hours_heavy: readNum("ironmindMaxRunHeavy", 24),
  };
  setStatus("Saving IronMind thresholds...");
  try {
    await fetchJson(`${API}/api/ironmind/settings`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (meta) meta.textContent = "Saved. Click Refresh Insight to apply now.";
    setStatus("IronMind thresholds saved.");
  } catch (err) {
    if (meta) meta.textContent = `Save failed: ${err.message || err}`;
    setStatus("IronMind threshold save failed.");
  }
}

function summarizeIronmindText(text) {
  const oneLine = String(text || "")
    .replace(/\s+/g, " ")
    .replace(/^IRONMIND DAILY INSIGHT\s*/i, "")
    .trim();
  if (!oneLine) return "No summary text.";
  return oneLine.length > 150 ? `${oneLine.slice(0, 147)}...` : oneLine;
}

function toYmd(date) {
  const d = new Date(date);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function buildRecentYmds(days) {
  const out = [];
  const now = new Date();
  for (let i = 0; i < days; i += 1) {
    const d = new Date(now);
    d.setDate(now.getDate() - i);
    out.push(toYmd(d));
  }
  return out;
}

function renderIronmindReport(report) {
  const summaryEl = qs("ironmindSummary");
  const metaEl = qs("ironmindMeta");
  if (summaryEl) {
    const sections = parseIronmindSections(String(report?.summary || ""));
    renderIronmindSections(summaryEl, sections);
  }
  if (metaEl) {
    const created = report?.created_at ? String(report.created_at).replace("T", " ").slice(0, 16) : "-";
    metaEl.textContent = `Report date: ${report?.report_date || "-"} | Updated: ${created}`;
  }
}

async function loadIronmindHistory(options = {}) {
  const silent = Boolean(options.silent);
  const listEl = qs("ironmindHistoryList");
  const includeMissing = Boolean(qs("ironmindShowMissingDays")?.checked);
  if (!listEl) return;
  try {
    const res = await fetchJson(`${API}/api/ironmind/history?report_type=daily_admin&days=7`);
    const rowsRaw = Array.isArray(res?.reports) ? res.reports : [];
    const rowsByDate = new Map(rowsRaw.map((r) => [String(r.report_date || "").trim(), r]));
    const rows = includeMissing
      ? buildRecentYmds(7).map((ymd) => {
          if (rowsByDate.has(ymd)) return rowsByDate.get(ymd);
          return {
            id: 0,
            report_date: ymd,
            report_type: "daily_admin",
            created_at: "-",
            summary: "IRONMIND DAILY INSIGHT\n\nRepairs Needed\n- Insufficient data\n\nOperational Risks\n- Insufficient data\n\nSuggestions\n- Insufficient data\n\nData Gaps\n- Insufficient data",
            synthetic_missing: true,
          };
        })
      : rowsRaw;
    listEl.innerHTML = "";
    if (!rows.length) {
      listEl.appendChild(item("<small>No IRONMIND history yet.</small>"));
      if (!silent) setStatus("No IRONMIND history found.");
      return;
    }

    rows.forEach((r) => {
      const created = r?.created_at && r.created_at !== "-" ? String(r.created_at).replace("T", " ").slice(0, 16) : "-";
      const preview = summarizeIronmindText(r?.summary || "");
      const previewClass = r?.synthetic_missing ? "ironmind-history-preview missing" : "ironmind-history-preview";
      const updatedText = r?.synthetic_missing ? "No generated report" : `Updated ${escapeHtml(created)}`;
      const node = item(
        `<div class="ironmind-history-item">` +
          `<div class="ironmind-history-meta"><b>${escapeHtml(r.report_date || "-")}</b> · ${updatedText}</div>` +
          `<div class="${previewClass}">${escapeHtml(preview)}</div>` +
          `<button class="ironmind-history-open" data-ironmind-history-id="${Number(r.id || 0)}">${r?.synthetic_missing ? "Open placeholder" : "Open report"}</button>` +
        `</div>`
      );
      node.dataset.ironmindRow = JSON.stringify(r);
      listEl.appendChild(node);
    });
    if (!silent) setStatus("IRONMIND history loaded.");
  } catch (err) {
    listEl.innerHTML = `<small class="muted">History unavailable right now.</small>`;
    if (!silent) setStatus("IRONMIND history error: " + err.message);
    throw err;
  }
}

async function refreshIronmindInsight() {
  const btn = qs("ironmindRefreshBtn");
  const contextNotes = String(qs("ironmindContext")?.value || "").trim();
  const detailMode = Boolean(qs("ironmindDetailMode")?.checked);
  if (btn) btn.disabled = true;
  setStatus("Refreshing IRONMIND insight...");
  try {
    await fetchJson(`${API}/api/ironmind/run`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        force: true,
        report_date: todayLocalYmd(),
        report_type: "daily_admin",
        context_notes: contextNotes || undefined,
        detail_mode: detailMode,
      }),
    });
    await loadIronmindInsight({ silent: true });
    await loadIronmindHistory({ silent: true });
    setStatus("IRONMIND insight refreshed.");
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function askIronmindQuestion() {
  const input = qs("ironmindAskInput");
  const out = qs("ironmindAskResult");
  const date = qs("date")?.value || todayLocalYmd();
  const contextNotes = String(qs("ironmindContext")?.value || "").trim();
  window.__ironmindAskHistory = Array.isArray(window.__ironmindAskHistory) ? window.__ironmindAskHistory : [];
  const question = String(input?.value || "").trim();
  if (!question) {
    if (out) out.innerHTML = `<small class="muted">Type a question first.</small>`;
    return;
  }
  if (out) out.innerHTML = `<small class="muted">Asking IRONMIND...</small>`;
  try {
    const res = await fetchJson(`${API}/api/ironmind/ask`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        question,
        date,
        history: window.__ironmindAskHistory.slice(-6),
        ...(contextNotes ? { context_notes: contextNotes } : {}),
      }),
    });
    const short = String(res?.short_answer || res?.answer || res?.message || "No answer returned.");
    const safe = escapeHtml(short).replace(/\n/g, "<br>");
    window.__ironmindAskHistory.push({ question, answer: short });
    window.__ironmindAskHistory = window.__ironmindAskHistory.slice(-12);
    saveIronmindAskHistoryLocal();
    renderIronmindAskHistory();
    const sid = getIronmindSessionId();
    fetchJson(`${API}/api/ironmind/ask-memory`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: sid, question, answer: short }),
    }).catch(() => {});
    loadIronmindHealth().catch(() => {});
    setStatus("IRONMIND question answered.");
  } catch (e) {
    if (out) out.innerHTML = `<small class="muted">Question failed: ${escapeHtml(e.message || String(e))}</small>`;
    setStatus("IRONMIND ask error: " + (e.message || e));
  }
}

function ironmindMemoryStorageKey() {
  return `ironmind_chat_${String(getSessionUser() || "user").toLowerCase()}_${String(getSessionSite() || "main").toLowerCase()}`;
}

function getIronmindSessionId() {
  const key = `${ironmindMemoryStorageKey()}_session`;
  let sid = String(localStorage.getItem(key) || "").trim();
  if (!sid) {
    sid = `sess-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
    localStorage.setItem(key, sid);
  }
  return sid;
}

function saveIronmindAskHistoryLocal() {
  try {
    localStorage.setItem(ironmindMemoryStorageKey(), JSON.stringify(window.__ironmindAskHistory || []));
  } catch {}
}

function loadIronmindAskHistoryLocal() {
  try {
    const parsed = JSON.parse(String(localStorage.getItem(ironmindMemoryStorageKey()) || "[]"));
    if (Array.isArray(parsed)) return parsed;
  } catch {}
  return [];
}

function renderIronmindAskHistory() {
  const out = qs("ironmindAskResult");
  if (!out) return;
  const rows = Array.isArray(window.__ironmindAskHistory) ? window.__ironmindAskHistory : [];
  if (!rows.length) {
    out.innerHTML = `<small class="muted">Ask a question for short operational answers.</small>`;
    return;
  }
  out.innerHTML = rows.map((h) => `
    <div style="margin-bottom:10px;">
      <div><b>You:</b> ${escapeHtml(String(h.question || ""))}</div>
      <div><b>IronMind:</b> ${escapeHtml(String(h.answer || "")).replace(/\n/g, "<br>")}</div>
    </div>
  `).join("");
}

async function hydrateIronmindAskMemory() {
  window.__ironmindAskHistory = loadIronmindAskHistoryLocal().slice(-12);
  renderIronmindAskHistory();
  try {
    const sid = getIronmindSessionId();
    const data = await fetchJson(`${API}/api/ironmind/ask-memory?session_id=${encodeURIComponent(sid)}&limit=30`);
    const rows = Array.isArray(data?.rows) ? data.rows : [];
    if (rows.length) {
      window.__ironmindAskHistory = rows.map((r) => ({
        question: String(r.question || ""),
        answer: String(r.answer || ""),
      })).slice(-12);
      saveIronmindAskHistoryLocal();
      renderIronmindAskHistory();
    }
  } catch {}
}

async function resetIronmindAskMemory() {
  const out = qs("ironmindAskResult");
  const key = ironmindMemoryStorageKey();
  try {
    localStorage.removeItem(key);
  } catch {}
  window.__ironmindAskHistory = [];
  renderIronmindAskHistory();
  try {
    const sid = getIronmindSessionId();
    await fetchJson(`${API}/api/ironmind/ask-memory?session_id=${encodeURIComponent(sid)}`, {
      method: "DELETE",
    });
  } catch {}
  if (out) out.innerHTML = `<small class="muted">Memory cleared.</small>`;
  setStatus("IronMind chat memory reset.");
}

async function generateIronmindRsgPlan(createWo = false) {
  const out = qs("ironmindRsgResult");
  const assetCode = String(qs("ironmindRsgAssetCode")?.value || "").trim().toUpperCase();
  const serviceHours = Number(qs("ironmindRsgHours")?.value || 2000);
  if (!assetCode) {
    if (out) out.textContent = "Asset code is required.";
    return;
  }
  const body = {
    asset_code: assetCode,
    service_hours: Number.isFinite(serviceHours) && serviceHours > 0 ? serviceHours : 2000,
  };
  if (out) out.textContent = "";
  setStatus(createWo ? "Creating RSG work order..." : "Generating RSG plan...");
  try {
    const endpoint = createWo ? "/api/ironmind/rsg/create-wo" : "/api/ironmind/rsg/plan";
    const res = await fetchJson(`${API}${endpoint}`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify(body),
    });
    if (out) out.textContent = JSON.stringify(res, null, 2);
    if (createWo && Number(res?.work_order_id || 0) > 0) {
      setStatus(`RSG WO created (#${Number(res.work_order_id)}).`);
      await loadDashboard().catch(() => {});
    } else {
      setStatus("RSG plan generated.");
    }
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus(`RSG ${createWo ? "WO creation" : "plan"} failed.`);
  }
}

async function previewIronmindRsgPdf() {
  const out = qs("ironmindRsgResult");
  const assetCode = String(qs("ironmindRsgAssetCode")?.value || "").trim().toUpperCase();
  const serviceHours = Number(qs("ironmindRsgHours")?.value || 2000);
  if (!assetCode) {
    if (out) out.textContent = "Asset code is required.";
    return;
  }
  const hours = Number.isFinite(serviceHours) && serviceHours > 0 ? serviceHours : 2000;
  setStatus("Opening RSG PDF preview...");
  try {
    const url = `${API}/api/ironmind/rsg/preview.pdf?asset_code=${encodeURIComponent(assetCode)}&service_hours=${encodeURIComponent(hours)}`;
    const win = window.open(url, "_blank", "noopener");
    if (!win) {
      // Popup blocked fallback
      window.location.href = url;
    }
    setStatus("RSG PDF preview opened.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("RSG PDF preview failed.");
  }
}

async function downloadIronmindRsgPdf() {
  const out = qs("ironmindRsgResult");
  const assetCode = String(qs("ironmindRsgAssetCode")?.value || "").trim().toUpperCase();
  const serviceHours = Number(qs("ironmindRsgHours")?.value || 2000);
  if (!assetCode) {
    if (out) out.textContent = "Asset code is required.";
    return;
  }
  const hours = Number.isFinite(serviceHours) && serviceHours > 0 ? serviceHours : 2000;
  setStatus("Preparing RSG PDF download...");
  try {
    const url = `${API}/api/ironmind/rsg/preview.pdf?asset_code=${encodeURIComponent(assetCode)}&service_hours=${encodeURIComponent(hours)}&download=1`;
    const a = document.createElement("a");
    a.href = url;
    a.target = "_blank";
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setStatus("RSG PDF download started.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("RSG PDF download failed.");
  }
}

function parseIronmindSections(text) {
  const defs = [
    { key: "repairs", name: "Repairs Needed" },
    { key: "risks", name: "Operational Risks" },
    { key: "suggestions", name: "Suggestions" },
    { key: "data_gaps", name: "Data Gaps" },
    { key: "anomalies", name: "Data Anomalies" },
  ];
  const src = String(text || "");
  return defs.map((sec, i) => {
    const next = defs[i + 1];
    const start = src.indexOf(sec.name);
    if (start === -1) return { key: sec.key, name: sec.name, items: [] };
    const end = next ? src.indexOf(next.name, start + sec.name.length) : src.length;
    const block = src.slice(start + sec.name.length, end === -1 ? src.length : end);
    const items = block
      .split("\n")
      .map((l) => l.trim())
      .filter((l) => l.startsWith("-"))
      .map((l) => l.slice(1).trim())
      .filter(Boolean);
    return { key: sec.key, name: sec.name, items: items.length ? items : ["Insufficient data"] };
  });
}

function renderIronmindSections(summaryEl, sections) {
  if (!summaryEl) return;
  const navMap = {
    repairs: { label: "\u2192 View Assets", key: "repairs" },
    risks: { label: "\u2192 View Stock", key: "risks" },
  };
  // Pattern: UPPERCASE asset code at start of item before ":"
  const assetPat = /^([A-Z][A-Z0-9_-]{1,9}):\s/;
  let html = `<div class="ironmind-header-line">IRONMIND DAILY INSIGHT</div>`;
  for (const sec of sections) {
    const nav = navMap[sec.key];
    const drillBtn = nav
      ? `<button class="ironmind-drill" data-ironmind-drill="${sec.key}">${escapeHtml(nav.label)}</button>`
      : "";
    html += `<div class="ironmind-title-row"><span class="ironmind-section-name">${escapeHtml(sec.name)}</span>${drillBtn}</div>`;
    html += `<ul class="ironmind-items">`;
    for (const itm of sec.items) {
      const m = sec.key === "repairs" ? itm.match(assetPat) : null;
      if (m) {
        const code = m[1];
        const rest = escapeHtml(itm.slice(m[0].length));
        html += `<li class="ironmind-item">- <button class="ironmind-asset-link" data-ironmind-asset="${escapeHtml(code)}">${escapeHtml(code)}</button>: ${rest}</li>`;
      } else {
        html += `<li class="ironmind-item">- ${escapeHtml(itm)}</li>`;
      }
    }
    html += `</ul>`;
  }
  summaryEl.innerHTML = html;
}

function ironmindDrillDown(sectionKey) {
  if (sectionKey === "repairs") {
    switchTab("assets");
  } else if (sectionKey === "risks") {
    switchTab("stock");
  }
}

async function ironmindGoToAsset(assetCode) {
  switchTab("assets");
  await loadAssetsFleet().catch(() => {});
  const code = String(assetCode || "").trim();
  if (!code) return;
  const exists = assetsFleetCache.some((c) => c.asset_code === code);
  if (exists) {
    await selectAssetCard(code, { loadHistory: true, scroll: true });
  }
}

function getDefaultLubeRange() {
  const end = new Date();
  const start = new Date();
  start.setDate(end.getDate() - 29);
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

let lubeUsageCache = null;

function isPartLikeOilType(v) {
  const t = String(v || "").trim();
  if (!t) return true;
  const lower = t.toLowerCase();
  const knownOilWords = ["oil", "lube", "lub", "hyd", "hydraulic", "grease", "coolant", "atf", "engine"];
  if (knownOilWords.some((w) => lower.includes(w))) return false;
  // Typical stock/part-like key: uppercase-ish code with digits/hyphen/underscore and no spaces.
  return /^[a-z0-9][a-z0-9\-_/.]{2,24}$/i.test(t) && !/\s/.test(t);
}

function renderLubeUsageTable(payload) {
  const lubeList = qs("lubeList");
  if (!lubeList) return;
  const assetFilter = String(qs("lubeFilterAsset")?.value || "").trim().toLowerCase();
  const oilTypeFilter = String(qs("lubeFilterOilType")?.value || "").trim().toLowerCase();
  const hidePartLike = Boolean(qs("lubeHidePartLike")?.checked);

  const detailLines = Array.isArray(payload?.lines) ? payload.lines : [];
  let flat = [];

  if (detailLines.length) {
    flat = detailLines
      .filter((r) => {
        const code = String(r.asset_code || "").toLowerCase();
        const part = String(r.part_code || "").toLowerCase();
        const typ = String(r.lube_type || "").toLowerCase();
        const name = String(r.part_name || "").toLowerCase();
        if (assetFilter && !code.includes(assetFilter)) return false;
        if (oilTypeFilter && !typ.includes(oilTypeFilter) && !part.includes(oilTypeFilter) && !name.includes(oilTypeFilter)) {
          return false;
        }
        if (hidePartLike && isPartLikeOilType(r.part_code) && !typ.includes("oil") && !name.includes("oil")) return false;
        return true;
      })
      .map((r) => ({
        usage_date: String(r.usage_date || ""),
        asset_code: String(r.asset_code || ""),
        asset_name: String(r.asset_name || ""),
        part_code: String(r.part_code || ""),
        part_name: String(r.part_name || ""),
        lube_type: String(r.lube_type || "lube"),
        smr: r.smr != null && Number.isFinite(Number(r.smr)) ? Number(r.smr) : null,
        qty: Number(r.quantity || 0),
        cost: Number(r.line_cost || 0),
        source: String(r.source || ""),
        work_order_id: r.work_order_id != null ? Number(r.work_order_id) : null,
      }));
  } else {
    const dataRows = Array.isArray(payload?.rows) ? payload.rows : [];
    for (const r of dataRows) {
      const assetCode = String(r.asset_code || "");
      const assetName = String(r.asset_name || "");
      if (assetFilter && !assetCode.toLowerCase().includes(assetFilter)) continue;
      const byType = Array.isArray(r.by_oil_type) ? r.by_oil_type : [];
      const normalized = byType.length
        ? byType.map((x) => ({
            oil_type: String(x.oil_type || "UNSPECIFIED"),
            qty: Number(x.qty ?? x.qty_total ?? 0),
            cost: Number(x.lube_cost ?? x.total_lube_cost ?? 0),
          }))
        : [{
            oil_type: "UNSPECIFIED",
            qty: Number(r.qty ?? r.qty_total ?? 0),
            cost: Number(r.lube_cost ?? r.total_lube_cost ?? 0),
          }];

      for (const t of normalized) {
        if (hidePartLike && isPartLikeOilType(t.oil_type)) continue;
        if (oilTypeFilter && !String(t.oil_type || "").toLowerCase().includes(oilTypeFilter)) continue;
        flat.push({
          usage_date: "",
          asset_code: assetCode,
          asset_name: assetName,
          part_code: t.oil_type,
          part_name: t.oil_type,
          lube_type: t.oil_type,
          smr: null,
          qty: t.qty,
          cost: t.cost,
          source: "summary",
          work_order_id: null,
        });
      }
    }
  }

  if (!flat.length) {
    lubeList.innerHTML = `<small>No lube rows match your filters.</small>`;
    return;
  }

  const visibleAssets = new Set(flat.map((r) => r.asset_code)).size;
  const visibleQty = flat.reduce((s, r) => s + Number(r.qty || 0), 0);
  const visibleCost = flat.reduce((s, r) => s + Number(r.cost || 0), 0);

  lubeList.innerHTML = `
    <div style="overflow:auto;">
      <table class="gridTable" style="min-width:1320px;">
        <thead>
          <tr>
            <th>Date</th>
            <th>Lube part no</th>
            <th>Description</th>
            <th>Type</th>
            <th>Plant no</th>
            <th>Machine</th>
            <th style="text-align:right;">Machine hrs</th>
            <th style="text-align:right;">Qty</th>
            <th style="text-align:right;">Cost</th>
            <th>Source</th>
            <th>WO #</th>
          </tr>
        </thead>
        <tbody>
          ${flat.map((r) => `
            <tr>
              <td style="padding:10px 8px;">${escapeHtml(r.usage_date || "-")}</td>
              <td style="padding:10px 8px;">${escapeHtml(r.part_code || "-")}</td>
              <td style="padding:10px 8px;">${escapeHtml(r.part_name || "-")}</td>
              <td style="padding:10px 8px;">${escapeHtml(r.lube_type || "-")}</td>
              <td style="padding:10px 8px;">${escapeHtml(r.asset_code || "-")}</td>
              <td style="padding:10px 8px;">${escapeHtml(r.asset_name || r.asset_code || "-")}</td>
              <td style="padding:10px 8px; text-align:right;">${r.smr != null ? Number(r.smr).toFixed(1) : "-"}</td>
              <td style="padding:10px 8px; text-align:right;">${Number(r.qty || 0).toFixed(1)}</td>
              <td style="padding:10px 8px; text-align:right;">${fmtMoney(r.cost || 0)}</td>
              <td style="padding:10px 8px;">${escapeHtml(r.source || "-")}</td>
              <td style="padding:10px 8px;">${r.work_order_id ? escapeHtml(String(r.work_order_id)) : "-"}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
    <div class="muted" style="margin-top:8px;">
      Showing ${flat.length} line(s) across ${visibleAssets} machine(s) | qty ${visibleQty.toFixed(1)} | cost ${fmtMoney(visibleCost)}
    </div>
  `;
}

async function loadLubeUsage() {
  const start = qs("lubeStart")?.value || "";
  const end = qs("lubeEnd")?.value || "";
  if (!start || !end) {
    alert("Select start and end dates.");
    return;
  }

  setStatus("Loading lube usage...");
  setSkeleton("lubeList", 2);
  const data = await fetchJson(`${API}/api/dashboard/lube?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`);

  setText("lubeQtyTotal", Number(data.summary?.line_qty_total ?? data.summary?.qty_total ?? 0).toFixed(1));
  setText("lubeEntries", Number(data.summary?.line_entries ?? data.summary?.entries ?? 0));
  setText("lubeAssets", Number(data.summary?.line_assets ?? data.summary?.assets ?? 0));
  setText("cLubeCost", fmtMoney(data.summary?.line_total_cost ?? data.summary?.total_lube_cost ?? 0));
  lubeUsageCache = data;
  renderLubeUsageTable(data);

  setStatus("Lube usage ready.");
}

async function saveFuelLog() {
  const meterRaw = (qs("fuelHoursRun")?.value || "").trim();
  const meter_unit = String(qs("fuelMeterUnit")?.value || "hours").trim().toLowerCase() === "km" ? "km" : "hours";
  const cc = String(qs("fuelCostCenter")?.value || "").trim();
  const basePayload = {
    asset_code: (qs("fuelAsset")?.value || "").trim(),
    log_date: (qs("fuelDate")?.value || "").trim() || undefined,
    liters: Number(qs("fuelLiters")?.value || 0),
    meter_run_value: meterRaw === "" ? undefined : Number(meterRaw),
    meter_unit,
    hours_run: meter_unit === "hours" && meterRaw !== "" ? Number(meterRaw) : undefined,
    source: (qs("fuelSource")?.value || "").trim() || undefined,
    cost_center_code: cc || undefined,
  };

  setStatus("Saving fuel log...");
  try {
    let res;
    try {
      res = await fetchJson(`${API}/api/dashboard/fuel/log`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(basePayload),
      });
    } catch (e) {
      const msg = String(e?.message || e || "");
      const isDup = /possible_duplicate_recent/i.test(msg) || /duplicate/i.test(msg);
      if (!isDup) throw e;
      const ok = confirm("Possible duplicate fuel input detected (same values in last 60 seconds).\nSave it anyway?");
      if (!ok) {
        setStatus("Fuel save cancelled (duplicate protection).");
        return;
      }
      res = await fetchJson(`${API}/api/dashboard/fuel/log`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...basePayload, force_duplicate: true }),
      });
    }
    setText("fuelInputResult", JSON.stringify(res, null, 2));
    setStatus("Fuel log saved.");
    await Promise.all([
      loadDashboard().catch(() => {}),
      loadFuelBenchmark().catch(() => {}),
    ]);
  } catch (e) {
    setText("fuelInputResult", String(e.message || e));
    setStatus("Fuel log save failed.");
  }
}

function parseFuelMassText(text) {
  const raw = String(text || "").trim();
  if (!raw) return [];
  const lines = raw.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const rows = [];
  for (const line of lines) {
    const parts = line.includes("\t") ? line.split("\t") : line.split(",");
    const [asset_code, log_date, liters, meter_unit, meter_run_value, source] = parts.map((p) => String(p || "").trim());
    if (!asset_code || !log_date || !liters) continue;
    rows.push({
      asset_code,
      log_date,
      liters: Number(liters),
      meter_unit: String(meter_unit || "").toLowerCase() === "km" ? "km" : "hours",
      meter_run_value: meter_run_value === "" ? undefined : Number(meter_run_value),
      source: source || undefined,
    });
  }
  return rows;
}

async function importFuelMassPaste() {
  const txt = qs("fuelMassPaste")?.value || "";
  const out = qs("fuelMassResult");
  if (out) out.textContent = "";
  const rows = parseFuelMassText(txt);
  if (!rows.length) return alert("No valid rows found. Paste at least: asset_code,log_date,liters.");
  setStatus(`Importing ${rows.length} fuel rows...`);
  let ok = 0;
  let fail = 0;
  const errs = [];
  for (const r of rows) {
    try {
      await fetchJson(`${API}/api/dashboard/fuel/log`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          asset_code: r.asset_code,
          log_date: r.log_date,
          liters: r.liters,
          meter_unit: r.meter_unit,
          meter_run_value: r.meter_run_value,
          source: r.source,
          force_duplicate: true,
        }),
      });
      ok += 1;
    } catch (e) {
      fail += 1;
      errs.push({ row: r, error: String(e.message || e) });
    }
  }
  if (out) out.textContent = JSON.stringify({ ok, fail, errors: errs.slice(0, 30) }, null, 2);
  setStatus(`Fuel mass import done: ${ok} ok, ${fail} failed.`);
  await Promise.all([loadDashboard().catch(() => {}), loadFuelBenchmark().catch(() => {})]);
}

async function loadFuelBaseline() {
  const asset_code = (qs("fuelBaseAsset")?.value || "").trim();
  if (!asset_code) return alert("Enter/select asset code first.");

  setStatus("Loading OEM baseline...");
  try {
    const data = await fetchJson(
      `${API}/api/dashboard/fuel/baseline?asset_code=${encodeURIComponent(asset_code)}`
    );
    const mode = String(data.asset?.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
    const baseline = mode === "km"
      ? Number(data.asset?.baseline_fuel_km_per_l || 2)
      : Number(data.asset?.baseline_fuel_l_per_hour || 5);
    const input = qs("fuelBaseValue");
    if (input) input.value = baseline.toFixed(3);
    const unitEl = qs("fuelBaseUnitLabel");
    if (unitEl) unitEl.textContent = mode === "km" ? "km/L" : "L/hr";
    const mu = qs("fuelMeterUnit");
    if (mu) mu.value = mode === "km" ? "km" : "hours";
    const meterInput = qs("fuelHoursRun");
    if (meterInput) meterInput.placeholder = mode === "km" ? "Distance since fill (km)" : "Hours since fill";
    setText(
      "fuelBaselineResult",
      JSON.stringify(
        {
          asset_code: data.asset?.asset_code,
          asset_name: data.asset?.asset_name,
          metric_mode: mode,
          baseline_value: baseline,
        },
        null,
        2
      )
    );
    setStatus("OEM baseline loaded.");
  } catch (e) {
    setText("fuelBaselineResult", String(e.message || e));
    setStatus("OEM baseline load failed.");
  }
}

async function syncFuelUnitFromAsset(assetCode, target = "input") {
  const code = String(assetCode || "").trim();
  if (!code) return;
  try {
    const data = await fetchJson(`${API}/api/dashboard/fuel/baseline?asset_code=${encodeURIComponent(code)}`);
    const mode = String(data.asset?.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
    if (target === "input" || target === "both") {
      const mu = qs("fuelMeterUnit");
      if (mu) mu.value = mode === "km" ? "km" : "hours";
      const meterInput = qs("fuelHoursRun");
      if (meterInput) meterInput.placeholder = mode === "km" ? "Distance since fill (km)" : "Hours since fill";
    }
    if (target === "baseline" || target === "both") {
      const unitEl = qs("fuelBaseUnitLabel");
      if (unitEl) unitEl.textContent = mode === "km" ? "km/L" : "L/hr";
    }
  } catch {}
}

async function loadCostSettings() {
  setStatus("Loading cost defaults...");
  const data = await fetchJson(`${API}/api/dashboard/cost/settings`);
  const s = data?.settings || {};
  const put = (id, v) => {
    const el = qs(id);
    if (el && v != null) el.value = Number(v).toFixed(2);
  };
  put("costFuelDefault", s.fuel_cost_per_liter_default ?? 1.5);
  put("costLubeDefault", s.lube_cost_per_qty_default ?? 4.0);
  put("costLaborDefault", s.labor_cost_per_hour_default ?? 35.0);
  put("costDowntimeDefault", s.downtime_cost_per_hour_default ?? 120.0);
  setStatus("Cost defaults ready.");
}

async function saveCostSettings() {
  const payload = {
    fuel_cost_per_liter_default: Number(qs("costFuelDefault")?.value || 0),
    lube_cost_per_qty_default: Number(qs("costLubeDefault")?.value || 0),
    labor_cost_per_hour_default: Number(qs("costLaborDefault")?.value || 0),
    downtime_cost_per_hour_default: Number(qs("costDowntimeDefault")?.value || 0),
  };
  setStatus("Saving cost defaults...");
  try {
    const res = await fetchJson(`${API}/api/dashboard/cost/settings`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("costSetupResult", JSON.stringify(res, null, 2));
    setStatus("Cost defaults saved.");
    await loadDashboard().catch(() => {});
  } catch (e) {
    setText("costSetupResult", String(e.message || e));
    setStatus("Cost defaults save failed.");
  }
}

async function saveCostAssetRates() {
  const payload = {
    asset_code: (qs("costAssetCode")?.value || "").trim(),
    fuel_cost_per_liter: (qs("costAssetFuel")?.value || "").trim(),
    downtime_cost_per_hour: (qs("costAssetDowntime")?.value || "").trim(),
    utilization_mode: (qs("costAssetUtilMode")?.value || "").trim(),
    km_per_hour_factor: (qs("costAssetKmFactor")?.value || "").trim(),
  };
  setStatus("Saving asset cost rates...");
  try {
    const res = await fetchJson(`${API}/api/dashboard/cost/asset-rates`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("costSetupResult", JSON.stringify(res, null, 2));
    setStatus("Asset cost rates saved.");
    await loadDashboard().catch(() => {});
  } catch (e) {
    setText("costSetupResult", String(e.message || e));
    setStatus("Asset cost rates save failed.");
  }
}

async function saveCostPartRate() {
  const payload = {
    part_code: (qs("costPartCode")?.value || "").trim(),
    unit_cost: Number(qs("costPartUnit")?.value || 0),
  };
  setStatus("Saving part unit cost...");
  try {
    const res = await fetchJson(`${API}/api/dashboard/cost/part-cost`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("costSetupResult", JSON.stringify(res, null, 2));
    setStatus("Part unit cost saved.");
    await loadDashboard().catch(() => {});
  } catch (e) {
    setText("costSetupResult", String(e.message || e));
    setStatus("Part unit cost save failed.");
  }
}

async function saveFuelBaseline() {
  const mode = String(qs("fuelMeterUnit")?.value || "hours").trim().toLowerCase() === "km" ? "km" : "hours";
  const payload = {
    asset_code: (qs("fuelBaseAsset")?.value || "").trim(),
    metric_mode: mode,
    ...(mode === "km"
      ? { baseline_fuel_km_per_l: Number(qs("fuelBaseValue")?.value || 0) }
      : { baseline_fuel_l_per_hour: Number(qs("fuelBaseValue")?.value || 0) }),
  };
  if (!payload.asset_code) return alert("Enter/select asset code first.");

  setStatus("Saving OEM baseline...");
  try {
    const res = await fetchJson(`${API}/api/dashboard/fuel/baseline`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("fuelBaselineResult", JSON.stringify(res, null, 2));
    setStatus("OEM baseline saved.");
    await loadFuelBenchmark().catch(() => {});
  } catch (e) {
    setText("fuelBaselineResult", String(e.message || e));
    setStatus("OEM baseline save failed.");
  }
}

function fuelSvgPeriodCompareLines(currentSeries, previousSeries, currentRange, previousRange) {
  const n = Math.max(
    Array.isArray(currentSeries) ? currentSeries.length : 0,
    Array.isArray(previousSeries) ? previousSeries.length : 0,
    1,
  );
  const cur = Array.isArray(currentSeries) ? currentSeries : [];
  const prev = Array.isArray(previousSeries) ? previousSeries : [];
  const allVals = [];
  for (let i = 0; i < n; i++) {
    allVals.push(Number(cur[i]?.liters || 0), Number(prev[i]?.liters || 0));
  }
  const max = Math.max(...allVals, 1);
  // Wider + taller for quarter-length ranges; scroll horizontally when needed
  const pxPerDay = n > 60 ? 14 : n > 31 ? 18 : n > 14 ? 22 : 28;
  const width = Math.max(720, Math.min(2200, 72 + n * pxPerDay));
  const height = n > 45 ? 460 : 420;
  const left = 52;
  const right = 18;
  const top = 28;
  const bottom = 42;
  const plotW = width - left - right;
  const plotH = height - top - bottom;
  const xAt = (i) => left + (n === 1 ? plotW / 2 : (i * plotW) / (n - 1));
  const yAt = (v) => top + plotH - (Number(v || 0) / max) * plotH;
  const showValueLabels = n <= 14;
  const dotR = n > 60 ? 2 : n > 31 ? 2.5 : 3.5;

  const linePath = (series, color) => {
    const pts = [];
    for (let i = 0; i < n; i++) {
      const v = Number(series[i]?.liters || 0);
      pts.push({ x: xAt(i), y: yAt(v), v });
    }
    const d = pts.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x.toFixed(1)} ${p.y.toFixed(1)}`).join(" ");
    const dots = pts.map((p) => {
      const label = showValueLabels
        ? `<text x="${p.x.toFixed(1)}" y="${(p.y - 8).toFixed(1)}" text-anchor="middle" font-size="10" fill="#334155">${p.v.toFixed(0)}</text>`
        : "";
      return `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="${dotR}" fill="${color}"/>${label}`;
    }).join("");
    return `<path d="${d}" fill="none" stroke="${color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>${dots}`;
  };

  const labelStep = n > 60 ? Math.ceil(n / 10) : n > 31 ? Math.ceil(n / 9) : n > 14 ? Math.ceil(n / 8) : 1;
  const xLabels = [];
  for (let i = 0; i < n; i++) {
    if (i % labelStep !== 0 && i !== n - 1 && i !== 0) continue;
    const rawDate = String(cur[i]?.date || prev[i]?.date || "").slice(0, 10);
    const label = /^\d{4}-\d{2}-\d{2}$/.test(rawDate)
      ? `${rawDate.slice(5, 7)}-${rawDate.slice(8, 10)}`
      : `D${i + 1}`;
    xLabels.push(`<text x="${xAt(i).toFixed(1)}" y="${top + plotH + 22}" text-anchor="middle" font-size="11" fill="#64748b">${label}</text>`);
  }

  const yTicks = [0, 0.25, 0.5, 0.75, 1].map((t) => {
    const v = max * t;
    const y = yAt(v);
    return `
      <line x1="${left}" y1="${y.toFixed(1)}" x2="${width - right}" y2="${y.toFixed(1)}" stroke="#e2e8f0" stroke-width="1"/>
      <text x="${left - 8}" y="${(y + 4).toFixed(1)}" text-anchor="end" font-size="11" fill="#64748b">${v.toFixed(0)}</text>
    `;
  }).join("");

  const curLabel = currentRange?.start && currentRange?.end
    ? `${currentRange.start} → ${currentRange.end}`
    : "Current";
  const prevLabel = previousRange?.start && previousRange?.end
    ? `${previousRange.start} → ${previousRange.end}`
    : "Previous";

  return `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="100%" height="${height}" role="img" aria-label="Fuel usage current versus previous period" preserveAspectRatio="xMinYMid meet">
      <rect x="0" y="0" width="${width}" height="${height}" fill="#ffffff" rx="8"/>
      <text x="${left}" y="16" font-size="12" fill="#475569">${escapeHtml(curLabel)} vs ${escapeHtml(prevLabel)} (liters per day)</text>
      ${yTicks}
      <line x1="${left}" y1="${top + plotH}" x2="${width - right}" y2="${top + plotH}" stroke="#cbd5e1" stroke-width="1"/>
      <line x1="${left}" y1="${top}" x2="${left}" y2="${top + plotH}" stroke="#cbd5e1" stroke-width="1"/>
      ${linePath(prev, "#94a3b8")}
      ${linePath(cur, "#2563eb")}
      ${xLabels}
    </svg>
  `;
}

function mountFuelCompareSvg(host, svgMarkup) {
  if (!host) return false;
  const markup = String(svgMarkup || "").trim();
  if (!markup) return false;
  const chartH = "420px";
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(markup, "image/svg+xml");
    if (doc.querySelector("parsererror")) throw new Error("svg parse error");
    const svg = doc.documentElement;
    if (!svg || String(svg.tagName || "").toLowerCase() !== "svg") throw new Error("svg root missing");
    const node = host.ownerDocument.importNode(svg, true);
    const h = Number(node.getAttribute("height") || 420) || 420;
    node.setAttribute("width", "100%");
    node.setAttribute("height", String(h));
    node.style.display = "block";
    node.style.width = "100%";
    node.style.height = `${h}px`;
    node.style.minHeight = chartH;
    host.replaceChildren(node);
    return true;
  } catch {
    host.innerHTML = markup;
    const svg = host.querySelector("svg");
    if (svg) {
      const h = Number(svg.getAttribute("height") || 420) || 420;
      svg.style.display = "block";
      svg.style.width = "100%";
      svg.style.height = `${h}px`;
    }
    return Boolean(svg);
  }
}

function fuelEquipChartRows(rows) {
  return (Array.isArray(rows) ? rows : [])
    .filter((r) => String(r.metric_mode || "hours").toLowerCase() !== "km")
    .filter((r) => Number(r.actual_lph || 0) > 0 || Number(r.oem_lph || 0) > 0)
    .map((r) => {
      const oem = Number(r.oem_lph || 0);
      const threshold = Number(
        r.excessive_threshold_lph != null
          ? r.excessive_threshold_lph
          : r.threshold_lph != null
            ? r.threshold_lph
            : oem > 0
              ? oem * 1.15
              : 0
      );
      return {
        asset_code: String(r.asset_code || "").trim(),
        label: String(r.asset_name || r.asset_code || "").trim() || String(r.asset_code || "Unknown"),
        actual: Number(r.actual_lph || 0),
        oem,
        threshold,
      };
    })
    .sort((a, b) => a.label.localeCompare(b.label));
}

function getSelectedFuelEquipmentCodes() {
  const host = qs("fuelEquipFilterList");
  if (!host) return null;
  const boxes = [...host.querySelectorAll('input[type="checkbox"][data-fuel-equip]')];
  if (!boxes.length) return [];
  return boxes.filter((b) => b.checked).map((b) => String(b.getAttribute("data-fuel-equip") || "").trim().toLowerCase());
}

/**
 * Derive a normalised equipment type from an asset label by stripping trailing
 * unit identifiers: "#2", "No.3", "(2)", serial numbers, etc.
 * Returns e.g. "CAT 950K" from "CAT 950K #2" or "Komatsu PC210" from "Komatsu PC210 No.3".
 */
function fuelEquipTypeGroup(label) {
  let s = String(label || "").trim();
  // Strip trailing patterns like "#1", "No.1", "No 1", "(1)", "- 1", "Unit 1", serial-like all-caps codes
  s = s
    .replace(/\s*[-–]?\s*#\s*\d+\s*$/i, "")
    .replace(/\s+No\.?\s*\d+\s*$/i, "")
    .replace(/\s*\(\s*\d+\s*\)\s*$/i, "")
    .replace(/\s+Unit\s+\d+\s*$/i, "")
    .replace(/\s+[A-Z]{2,}\d{4,}\s*$/i, "")  // strip trailing serial-like codes
    .trim();
  // Normalise multiple spaces
  return s.replace(/\s+/g, " ") || label;
}

function renderFuelEquipmentFilter(rows, preserveSelection = true) {
  const host = qs("fuelEquipFilterList");
  if (!host) return;
  const all = fuelEquipChartRows(rows);
  if (!all.length) {
    host.className = "fuel-equip-filter-list muted";
    host.textContent = "No machine (L/hr) equipment in this period.";
    // Clear type dropdown too
    const typeSel = qs("fuelEquipTypeFilter");
    if (typeSel) typeSel.innerHTML = `<option value="">Filter by type…</option>`;
    return;
  }
  const prev = preserveSelection ? new Set(getSelectedFuelEquipmentCodes() || []) : null;
  const selectAllByDefault = !prev || prev.size === 0;
  host.className = "fuel-equip-filter-list";
  host.innerHTML = all.map((r) => {
    const code = r.asset_code;
    const checked = selectAllByDefault || prev.has(code.toLowerCase()) ? "checked" : "";
    const title = escapeHtml(`${r.label} (${code})`);
    const typeGroup = fuelEquipTypeGroup(r.label);
    return `<label class="fuel-equip-filter-item" title="${title}" data-fuel-type="${escapeHtml(typeGroup)}">
      <input type="checkbox" data-fuel-equip="${escapeHtml(code)}" ${checked} />
      <span>${escapeHtml(r.label)}</span>
    </label>`;
  }).join("");

  // Populate type dropdown — unique groups sorted, only show if >1 group or >1 member in a group
  const typeSel = qs("fuelEquipTypeFilter");
  if (typeSel) {
    const groups = new Map();
    all.forEach((r) => {
      const g = fuelEquipTypeGroup(r.label);
      if (!groups.has(g)) groups.set(g, []);
      groups.get(g).push(r.asset_code);
    });
    // Only offer groups that have ≥2 members (singles are selected individually)
    const multiGroups = [...groups.entries()].filter(([, codes]) => codes.length >= 2).sort((a, b) => a[0].localeCompare(b[0]));
    typeSel.innerHTML = `<option value="">Filter by type…</option>` +
      multiGroups.map(([g]) => `<option value="${escapeHtml(g)}">${escapeHtml(g)} (${groups.get(g).length})</option>`).join("");
  }
}

function fuelSvgEquipmentConsumption(points) {
  const rows = Array.isArray(points) ? points : [];
  const n = Math.max(rows.length, 1);
  const vals = [];
  for (const r of rows) {
    vals.push(Number(r.actual || 0), Number(r.oem || 0), Number(r.threshold || 0));
  }
  const rawMax = Math.max(...vals, 1);
  const max = Math.ceil(rawMax / 5) * 5 || 5;
  const slot = Math.max(70, Math.min(110, 900 / Math.max(n, 1)));
  const width = Math.max(720, 80 + n * slot);
  const height = 440;
  const left = 52;
  const right = 20;
  const top = 36;
  const bottom = 88;
  const plotW = width - left - right;
  const plotH = height - top - bottom;
  const band = plotW / n;
  const barW = Math.min(42, Math.max(18, band * 0.45));
  const xCenter = (i) => left + band * i + band / 2;
  const yAt = (v) => top + plotH - (Number(v || 0) / max) * plotH;

  const bars = rows.map((r, i) => {
    const x = xCenter(i) - barW / 2;
    const y = yAt(r.actual);
    const h = Math.max(0, top + plotH - y);
    const labelY = h > 18 ? y + 14 : y - 6;
    const labelFill = h > 18 ? "#ffffff" : "#111827";
    return `
      <rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(1)}" height="${h.toFixed(1)}" fill="#4b5563" rx="2"/>
      <text x="${xCenter(i).toFixed(1)}" y="${labelY.toFixed(1)}" text-anchor="middle" font-size="11" font-weight="600" fill="${labelFill}">${Number(r.actual || 0).toFixed(0)}</text>
    `;
  }).join("");

  const linePath = (key, color) => {
    const pts = rows.map((r, i) => ({ x: xCenter(i), y: yAt(r[key]), v: Number(r[key] || 0) }));
    const d = pts.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x.toFixed(1)} ${p.y.toFixed(1)}`).join(" ");
    const dots = pts.map((p) => `
      <circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="4" fill="${color}" stroke="#fff" stroke-width="1"/>
      <text x="${p.x.toFixed(1)}" y="${(p.y - 10).toFixed(1)}" text-anchor="middle" font-size="11" font-weight="600" fill="#111827">${p.v.toFixed(0)}</text>
    `).join("");
    return `<path d="${d}" fill="none" stroke="${color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>${dots}`;
  };

  const xLabels = rows.map((r, i) => {
    const label = String(r.label || r.asset_code || "").slice(0, 18);
    return `<text x="${xCenter(i).toFixed(1)}" y="${top + plotH + 18}" text-anchor="end" font-size="11" fill="#334155" transform="rotate(-32 ${xCenter(i).toFixed(1)} ${top + plotH + 18})">${escapeHtml(label)}</text>`;
  }).join("");

  const yTicks = [];
  const step = max <= 20 ? 5 : max <= 50 ? 5 : 10;
  for (let v = 0; v <= max + 0.001; v += step) {
    const y = yAt(v);
    yTicks.push(`
      <line x1="${left}" y1="${y.toFixed(1)}" x2="${width - right}" y2="${y.toFixed(1)}" stroke="#e5e7eb" stroke-width="1"/>
      <text x="${left - 8}" y="${(y + 4).toFixed(1)}" text-anchor="end" font-size="11" fill="#64748b">${v}</text>
    `);
  }

  return `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="100%" height="${height}" role="img" aria-label="Equipment fuel consumptions" preserveAspectRatio="xMinYMid meet">
      <rect x="0" y="0" width="${width}" height="${height}" fill="#ffffff" rx="8"/>
      <text x="${left}" y="18" font-size="14" font-weight="700" fill="#111827">Equipment Fuel Consumptions</text>
      <text x="${left}" y="34" font-size="11" fill="#64748b">L/hr — Actual (bar) · OEM (green) · Threshold (red)</text>
      ${yTicks.join("")}
      <line x1="${left}" y1="${top + plotH}" x2="${width - right}" y2="${top + plotH}" stroke="#94a3b8" stroke-width="1"/>
      <line x1="${left}" y1="${top}" x2="${left}" y2="${top + plotH}" stroke="#94a3b8" stroke-width="1"/>
      ${bars}
      ${linePath("oem", "#16a34a")}
      ${linePath("threshold", "#dc2626")}
      ${xLabels}
    </svg>
  `;
}

function aggregateFuelByType(points) {
  const groups = new Map();
  for (const r of points) {
    const g = fuelEquipTypeGroup(r.label);
    if (!groups.has(g)) groups.set(g, []);
    groups.get(g).push(r);
  }
  const avg = (arr, key) => arr.reduce((s, r) => s + Number(r[key] || 0), 0) / arr.length;
  return [...groups.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([g, members]) => ({
      asset_code: g,
      label: members.length > 1 ? `${g} (avg ${members.length})` : g,
      actual: avg(members, "actual"),
      oem: avg(members, "oem"),
      threshold: avg(members, "threshold"),
      count: members.length,
    }));
}

function renderFuelEquipmentChart(rows) {
  const host = qs("fuelEquipChart");
  const summaryEl = qs("fuelEquipChartSummary");
  const wrap = qs("fuelEquipChartWrap");
  if (!host) return;

  const all = fuelEquipChartRows(rows);
  if (!all.length) {
    if (summaryEl) summaryEl.textContent = "No machine (L/hr) rows to chart for this period.";
    host.className = "fuel-equip-chart muted";
    host.textContent = "No L/hr equipment data.";
    return;
  }

  const selected = new Set(getSelectedFuelEquipmentCodes() || []);
  const points = selected.size
    ? all.filter((r) => selected.has(String(r.asset_code || "").toLowerCase()))
    : [];

  if (!points.length) {
    if (summaryEl) summaryEl.textContent = "Select one or more machines in the equipment filter to show the chart.";
    host.className = "fuel-equip-chart muted";
    host.textContent = "No equipment selected.";
    if (wrap) wrap.style.display = "";
    return;
  }

  const viewMode = String(qs("fuelEquipViewMode")?.value || "individual");
  const chartPoints = viewMode === "type" ? aggregateFuelByType(points) : points;

  if (summaryEl) {
    if (viewMode === "type") {
      const typeCount = chartPoints.length;
      const unitCount = points.length;
      summaryEl.textContent = `Showing ${typeCount} type${typeCount !== 1 ? "s" : ""} averaged from ${unitCount} unit${unitCount !== 1 ? "s" : ""}. Actual L/hr (bar) vs OEM (green) and threshold (red).`;
    } else {
      summaryEl.textContent = `Showing ${points.length} of ${all.length} machines (L/hr). Actual bars vs OEM (green) and threshold (red).`;
    }
  }
  if (wrap) wrap.style.display = "";
  host.className = "fuel-equip-chart";
  const svgMarkup = fuelSvgEquipmentConsumption(chartPoints);
  const mounted = mountFuelCompareSvg(host, svgMarkup);
  if (!mounted) {
    host.className = "fuel-equip-chart muted";
    host.textContent = "Chart could not be rendered.";
  }
}

function refreshFuelEquipmentChartFromStore() {
  renderFuelEquipmentFilter(window.__fuelBenchmarkChartRows || [], true);
  renderFuelEquipmentChart(window.__fuelBenchmarkChartRows || []);
}

function renderFuelPeriodCompareChart(data) {
  const host = qs("fuelPeriodCompareChart");
  const summaryEl = qs("fuelPeriodCompareSummary");
  const wrap = qs("fuelPeriodCompareWrap");
  if (!host) return;

  if (!data?.ok) {
    if (summaryEl) summaryEl.textContent = "Could not load period comparison.";
    host.className = "fuel-period-compare-chart muted";
    host.textContent = "No comparison data.";
    return;
  }

  const current = data.current || {};
  const previous = data.previous || {};
  const delta = data.delta || {};
  const curTotal = Number(current.total_liters || 0);
  const prevTotal = Number(previous.total_liters || 0);
  const deltaLiters = Number(delta.liters || 0);
  const deltaPct = delta.pct;
  const sign = deltaLiters >= 0 ? "+" : "";

  if (summaryEl) {
    summaryEl.innerHTML =
      `<strong>Current:</strong> ${curTotal.toFixed(1)} L (${escapeHtml(current.start || "")} to ${escapeHtml(current.end || "")})` +
      ` · <strong>Previous:</strong> ${prevTotal.toFixed(1)} L (${escapeHtml(previous.start || "")} to ${escapeHtml(previous.end || "")})` +
      ` · <strong>Change:</strong> <span class="${deltaLiters > 0 ? "pill red" : deltaLiters < 0 ? "pill blue" : "pill"}">${sign}${deltaLiters.toFixed(1)} L` +
      `${deltaPct == null ? "" : ` (${sign}${Number(deltaPct).toFixed(1)}%)`}</span></span>`;
  }

  if (wrap) wrap.style.display = "";
  host.className = "fuel-period-compare-chart";
  const svgMarkup = fuelSvgPeriodCompareLines(
    current.series,
    previous.series,
    { start: current.start, end: current.end },
    { start: previous.start, end: previous.end },
  );
  const mounted = mountFuelCompareSvg(host, svgMarkup);
  if (!mounted) {
    host.className = "fuel-period-compare-chart muted";
    host.textContent = "Chart could not be rendered for this period.";
  }
}

function setFuelPeriodDates(start, end) {
  if (qs("fuelStart")) qs("fuelStart").value = start;
  if (qs("fuelEnd")) qs("fuelEnd").value = end;
}

function applyFuelPeriodPreset(kind) {
  const today = new Date();
  const y = today.getUTCFullYear();
  const m = today.getUTCMonth() + 1;
  const d = today.getUTCDate();
  const pad = (n) => String(n).padStart(2, "0");
  const todayStr = `${y}-${pad(m)}-${pad(d)}`;
  const clip = (end) => (end > todayStr ? todayStr : end);
  let start = todayStr;
  let end = todayStr;
  if (kind === "q1") {
    start = `${y}-01-01`;
    end = clip(`${y}-03-31`);
  } else if (kind === "q2") {
    start = `${y}-04-01`;
    end = clip(`${y}-06-30`);
  } else if (kind === "q3") {
    start = `${y}-07-01`;
    end = clip(`${y}-09-30`);
  } else if (kind === "ytd") {
    start = `${y}-01-01`;
    end = todayStr;
  } else if (kind === "mtd") {
    start = `${y}-${pad(m)}-01`;
    end = todayStr;
  }
  setFuelPeriodDates(start, end);
  return loadFuelBenchmark();
}

async function loadFuelBenchmark() {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const mode = String(qs("fuelModeFilter")?.value || "").trim();
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  const duplicatesOnly = Boolean(qs("fuelDupOnly")?.checked);
  const runToken = Date.now() + Math.random();
  window.__fuelBenchmarkRunToken = runToken;
  if (!start || !end) return alert("Select start and end dates.");

  setStatus("Loading fuel benchmark...");
  setSkeleton("fuelBenchmarkList", 2);
  const compareUrl =
    `${API}/api/dashboard/fuel/period-compare?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&mode=${encodeURIComponent(mode)}&asset_code=${encodeURIComponent(assetCode)}`;
  const [data, compareData] = await Promise.all([
    duplicatesOnly
      ? fetchJson(
        `${API}/api/dashboard/fuel/duplicates?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&mode=${encodeURIComponent(mode)}&asset_code=${encodeURIComponent(assetCode)}`
      )
      : fetchJson(
        `${API}/api/dashboard/fuel?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}&mode=${encodeURIComponent(mode)}&asset_code=${encodeURIComponent(assetCode)}`
      ),
    duplicatesOnly ? Promise.resolve(null) : fetchJson(compareUrl).catch(() => null),
  ]);
  if (window.__fuelBenchmarkRunToken !== runToken) return;

  const rawRows = Array.isArray(data.rows) ? data.rows : [];
  const benchmarkRows = rawRows;
  const displayRows = benchmarkRows;

  const benchmarkSummary = duplicatesOnly
    ? {
        fuel_liters: Number(data.summary?.fuel_liters || 0),
        duplicate_rows: Number(data.summary?.duplicate_rows || 0),
      }
    : (data.summary && data.summary.avg_lph != null
      ? {
          fuel_liters: Number(data.summary.fuel_liters || 0),
          hours_run: Number(data.summary.hours_run || 0),
          km_run: Number(data.summary.km_run || 0),
          avg_lph: data.summary.avg_lph == null ? null : Number(data.summary.avg_lph),
          avg_km_per_l: data.summary.avg_km_per_l == null ? null : Number(data.summary.avg_km_per_l),
          excessive_count: Number(data.summary.excessive_count || data.summary.excessive || 0),
        }
      : benchmarkRows.reduce(
        (acc, r) => {
          const mode = String(r.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
          const fuel = Number(r.fuel_liters || 0);
          const hours = Number(r.hours_run || 0);
          const km = Number(r.km_run || 0);
          acc.fuel_liters += fuel;
          if (mode === "km") {
            acc.km_run += km;
            if (km > 0) acc.km_fuel += fuel;
          } else {
            acc.hours_run += hours;
            if (hours > 0) acc.hours_fuel += fuel;
          }
          if (Number(r.is_excessive || false)) acc.excessive_count += 1;
          return acc;
        },
        { fuel_liters: 0, hours_run: 0, km_run: 0, hours_fuel: 0, km_fuel: 0, excessive_count: 0 }
      ));
  if (!duplicatesOnly && benchmarkSummary.avg_lph === undefined) {
    benchmarkSummary.avg_lph =
      benchmarkSummary.hours_run > 0
        ? benchmarkSummary.hours_fuel / benchmarkSummary.hours_run
        : null;
    benchmarkSummary.avg_km_per_l =
      benchmarkSummary.km_fuel > 0
        ? benchmarkSummary.km_run / benchmarkSummary.km_fuel
        : null;
  }

  if (duplicatesOnly) {
    setText("fbFuelTotal", Number(benchmarkSummary.fuel_liters || 0).toFixed(2));
    setText("fbHoursTotal", "-");
    setText("fbKmTotal", "-");
    setText("fbAvgLph", "-");
    setText("fbAvgKmpl", "-");
    setText("fbExcessive", Number(benchmarkSummary.duplicate_rows || 0));
    const compareHost = qs("fuelPeriodCompareChart");
    const compareSummary = qs("fuelPeriodCompareSummary");
    if (compareSummary) compareSummary.textContent = "Period comparison is hidden while viewing duplicates.";
    if (compareHost) {
      compareHost.className = "fuel-period-compare-chart muted";
      compareHost.textContent = "Switch off duplicates filter to view period comparison.";
    }
    window.__fuelBenchmarkChartRows = [];
    const equipHost = qs("fuelEquipChart");
    const equipFilter = qs("fuelEquipFilterList");
    const equipSummary = qs("fuelEquipChartSummary");
    if (equipSummary) equipSummary.textContent = "Equipment chart is hidden while viewing duplicates.";
    if (equipFilter) {
      equipFilter.className = "fuel-equip-filter-list muted";
      equipFilter.textContent = "Switch off duplicates filter to select equipment.";
    }
    if (equipHost) {
      equipHost.className = "fuel-equip-chart muted";
      equipHost.textContent = "Switch off duplicates filter to view equipment chart.";
    }
  } else {
    // Pills must reflect only the currently selected date range + filters.
    const s = data.summary || {};
    setText("fbFuelTotal", Number(s.fuel_liters || 0).toFixed(2));
    setText("fbHoursTotal", Number(s.hours_run || 0).toFixed(2));
    setText("fbKmTotal", Number(s.km_run || 0).toFixed(2));
    setText("fbAvgLph", s.avg_lph == null ? "-" : Number(s.avg_lph).toFixed(3));
    setText("fbAvgKmpl", s.avg_km_per_l == null ? "-" : Number(s.avg_km_per_l).toFixed(3));
    setText("fbExcessive", Number(s.excessive_count || 0));
    renderFuelPeriodCompareChart(compareData);
    window.__fuelBenchmarkChartRows = benchmarkRows;
    renderFuelEquipmentFilter(benchmarkRows, false);
    renderFuelEquipmentChart(benchmarkRows);
  }

  const list = qs("fuelBenchmarkList");
  if (list) {
    const isRowFlagged = (r) => {
      const mode = String(r?.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
      if (mode === "km") {
        const actual = Number(r?.actual_km_per_l);
        const benchmark = Number(r?.oem_km_per_l);
        if (!Number.isFinite(actual) || actual <= 0) return false; // no entries -> OK
        if (!Number.isFinite(benchmark) || benchmark <= 0) return false;
        return actual < benchmark;
      }
      const actual = Number(r?.actual_lph);
      const benchmark = Number(r?.oem_lph);
      if (!Number.isFinite(actual) || actual <= 0) return false; // no entries -> OK
      if (!Number.isFinite(benchmark) || benchmark <= 0) return false;
      return actual > benchmark;
    };
    list.innerHTML = "";
    displayRows.forEach((r) => {
      if (duplicatesOnly) {
        const runVal = Number(r.meter_run_value ?? r.hours_run ?? 0);
        const runTxt = String(r.metric_mode || "hours") === "km"
          ? `Meter: ${runVal.toFixed(2)} km`
          : `Meter: ${runVal.toFixed(2)} hours`;
        list.appendChild(
          item(
            `<div class="fuel-item-head"><b>${r.asset_code}</b> — ${r.log_date} <span class='pill red'>DUPLICATE x${Number(r.duplicate_count || 0)}</span></div>` +
            `<small class="fuel-item-desc">${r.asset_name || ""}</small>` +
            `<small class="fuel-item-meta">Fuel: ${Number(r.liters || 0).toFixed(2)}L | ${runTxt} | Source: ${r.source || "-"}</small>` +
            `<br><button data-fuel-delete="${Number(r.id || 0)}">Delete this entry</button>`
          )
        );
        return;
      }
      const isKm = String(r.metric_mode || "hours") === "km";
      const hiredTag = r.is_hired ? " <span class='pill' style='font-size:0.65rem;'>HIRED</span>" : "";
      const archTag = Number(r.archived) ? " <span class='pill' style='font-size:0.65rem;'>ARCH</span>" : "";
      const modeTag = `<span class='pill blue' style='font-size:0.65rem;'>${isKm ? "km/L" : "L/hr"}</span>`;
      const runSrc = r.run_source && r.run_source !== "none"
        ? ` | Run: ${isKm ? `${Number(r.km_run || 0).toFixed(2)} km` : `${Number(r.hours_run || 0).toFixed(2)} h`} (${String(r.run_source).replace(/_/g, " ")})`
        : "";
      const flag = isRowFlagged(r)
        ? `<span class='pill red'>${isKm ? "UNDER BENCHMARK" : "EXCESSIVE"}</span>`
        : "<span class='pill blue'>OK</span>";
      const machineKey = String(r.asset_code || "").replace(/[^A-Za-z0-9_-]/g, "_");
      list.appendChild(
        item(
          `<div class="fuel-item-head"><b>${r.asset_code}</b>${hiredTag}${archTag} ${modeTag} — ${
            isKm
              ? `${r.actual_km_per_l == null ? "-" : Number(r.actual_km_per_l).toFixed(3)} km/L`
              : `${r.actual_lph == null ? "-" : Number(r.actual_lph).toFixed(3)} L/hr`
          } ${flag}</div>` +
          `<small class="fuel-item-desc">${r.asset_name || ""}</small>` +
          `<small class="fuel-item-meta">${
            isKm
              ? `OEM: ${Number(r.oem_km_per_l || 0).toFixed(3)} km/L | Fuel: ${Number(r.fuel_liters || 0).toFixed(2)}L | Distance: ${Number(r.km_run || 0).toFixed(2)} km${runSrc}`
              : `OEM: ${Number(r.oem_lph || 0).toFixed(3)} L/hr | Fuel: ${Number(r.fuel_liters || 0).toFixed(2)}L | Hours: ${Number(r.hours_run || 0).toFixed(2)}${runSrc}`
          }</small>` +
          `<br><button data-fuel-machine="${String(r.asset_code || "").replace(/"/g, "&quot;")}">Open machine history</button> ` +
          `<button data-fuel-machine-pdf="${String(r.asset_code || "").replace(/"/g, "&quot;")}">Machine PDF</button>` +
          `<div class="fuel-inline-history" id="fuel-inline-${machineKey}"></div>`
        )
      );
    });
    if (!displayRows.length) {
      list.appendChild(item(`<small>${duplicatesOnly ? "No duplicate fuel entries found in this period." : "No fuel benchmark data in this period."}</small>`));
    }
  }

  if (duplicatesOnly) {
    setStatus(`Duplicate filter ready (${Number(data.summary?.duplicate_rows || 0)} rows in ${Number(data.summary?.duplicate_groups || 0)} groups).`);
  } else {
    const machineCount = displayRows.length;
    setStatus(machineCount ? `Fuel benchmark ready (${machineCount} machine${machineCount === 1 ? "" : "s"}).` : "Fuel benchmark ready.");
  }
}

function fuelJanToDateRange() {
  const now = new Date();
  const year = now.getFullYear();
  return {
    start: `${year}-01-01`,
    end: now.toISOString().slice(0, 10),
  };
}

function fuelReconExpectedLiters(row) {
  const mode = String(row?.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
  const fuel = Number(row?.fuel_liters || 0);
  if (mode === "km") {
    const kmRun = Number(row?.km_run || 0);
    const baseKmpl = Number(row?.oem_km_per_l || 0);
    const expected = kmRun > 0 && baseKmpl > 0 ? (kmRun / baseKmpl) : 0;
    return { mode, fuel, run: kmRun, expected };
  }
  const hrsRun = Number(row?.hours_run || 0);
  const baseLph = Number(row?.oem_lph || 0);
  const expected = hrsRun > 0 && baseLph > 0 ? (hrsRun * baseLph) : 0;
  return { mode, fuel, run: hrsRun, expected };
}

async function runFuelReconciliation(useJanToDate = true) {
  let start = (qs("fuelStart")?.value || "").trim();
  let end = (qs("fuelEnd")?.value || "").trim();
  if (useJanToDate) {
    const ytd = fuelJanToDateRange();
    start = ytd.start;
    end = ytd.end;
    if (qs("fuelStart")) qs("fuelStart").value = start;
    if (qs("fuelEnd")) qs("fuelEnd").value = end;
  }
  if (!start || !end) return alert("Select start and end dates.");
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  const fuelPrice = Math.max(0, Number(qs("fuelReconPrice")?.value || 0));
  setStatus("Running fuel reconciliation...");
  setSkeleton("fuelReconList", 2);
  const data = await fetchJson(
    `${API}/api/dashboard/fuel?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}&mode=&asset_code=${encodeURIComponent(assetCode)}`
  );
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  const reconRows = rows
    .map((r) => {
      const calc = fuelReconExpectedLiters(r);
      const variance = calc.fuel - calc.expected;
      const allowed = calc.expected * Math.max(0, tolerance);
      const unexplained = Math.max(0, variance - allowed);
      const variancePct = calc.expected > 0 ? (variance / calc.expected) * 100 : null;
      const reasons = [];
      if (calc.run <= 0 && calc.fuel > 0) reasons.push("Fuel captured with no run basis");
      if (Number(r?.fill_count || 0) < 2) reasons.push("Low sample count (<2 fills)");
      if (Number(r?.is_excessive || false)) reasons.push("Outside benchmark tolerance");
      if (variancePct != null && variancePct > 50) reasons.push("Variance above +50%");
      return {
        ...r,
        recon_mode: calc.mode,
        expected_liters: Number(calc.expected.toFixed(2)),
        variance_liters: Number(variance.toFixed(2)),
        variance_pct: variancePct == null ? null : Number(variancePct.toFixed(1)),
        unexplained_liters: Number(unexplained.toFixed(2)),
        anomaly_reasons: reasons,
      };
    })
    .filter((r) => r.fuel_liters > 0)
    .sort((a, b) => Number(b.unexplained_liters || 0) - Number(a.unexplained_liters || 0));

  const totals = reconRows.reduce((acc, r) => {
    acc.actual += Number(r.fuel_liters || 0);
    acc.expected += Number(r.expected_liters || 0);
    acc.unexplained += Number(r.unexplained_liters || 0);
    if ((r.anomaly_reasons || []).length) acc.anomalies += 1;
    return acc;
  }, { actual: 0, expected: 0, unexplained: 0, anomalies: 0 });
  const missingValue = totals.unexplained * fuelPrice;
  const summary = qs("fuelReconSummary");
  if (summary) {
    summary.className = "muted";
    summary.innerHTML =
      `<b>Recon period:</b> ${escapeHtml(start)} to ${escapeHtml(end)} | ` +
      `<b>Actual:</b> ${totals.actual.toFixed(2)} L | ` +
      `<b>Expected:</b> ${totals.expected.toFixed(2)} L | ` +
      `<b>Estimated missing (unexplained):</b> ${totals.unexplained.toFixed(2)} L | ` +
      `<b>Estimated value:</b> ${fmtMoney(missingValue)} | ` +
      `<b>Anomaly assets:</b> ${totals.anomalies}`;
  }
  const list = qs("fuelReconList");
  if (list) {
    list.innerHTML = "";
    const top = reconRows.filter((r) => Number(r.unexplained_liters || 0) > 0 || (r.anomaly_reasons || []).length).slice(0, 25);
    if (!top.length) {
      list.appendChild(item("<small>No anomalies found for selected scope.</small>"));
    } else {
      top.forEach((r) => {
        const reasons = (r.anomaly_reasons || []).join("; ") || "Review";
        const modeLabel = r.recon_mode === "km" ? "km/L" : "L/hr";
        list.appendChild(
          item(
            `<div class="fuel-item-head"><b>${escapeHtml(r.asset_code || "-")}</b> — ${escapeHtml(r.asset_name || "")}</div>` +
            `<small class="fuel-item-meta">Mode: ${modeLabel} | Actual: ${Number(r.fuel_liters || 0).toFixed(2)}L | Expected: ${Number(r.expected_liters || 0).toFixed(2)}L | Variance: ${Number(r.variance_liters || 0).toFixed(2)}L | Unexplained: ${Number(r.unexplained_liters || 0).toFixed(2)}L</small>` +
            `<small class="fuel-item-meta">Reasons: ${escapeHtml(reasons)}</small>`
          )
        );
      });
    }
  }
  setStatus("Fuel reconciliation ready.");
}

function openFuelBenchmarkPdf(download = false) {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const modeFilter = String(qs("fuelModeFilter")?.value || "").trim();
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  if (!start || !end) return alert("Select start and end dates.");

  const mode = download ? "&download=1" : "";
  const url =
    `${API}/api/reports/fuel-benchmark.pdf?start=${encodeURIComponent(start)}` +
    `&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}&mode=${encodeURIComponent(modeFilter)}&asset_code=${encodeURIComponent(assetCode)}${mode}`;
  window.open(url, "_blank");
}

function openFuelBenchmarkXlsx() {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const modeFilter = String(qs("fuelModeFilter")?.value || "").trim();
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  if (!start || !end) return alert("Select start and end dates.");

  const url =
    `${API}/api/reports/fuel-benchmark.xlsx?start=${encodeURIComponent(start)}` +
    `&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}&mode=${encodeURIComponent(modeFilter)}&asset_code=${encodeURIComponent(assetCode)}`;
  window.open(url, "_blank");
}

function openFuelReconciliationPdf(download = false) {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const modeFilter = String(qs("fuelModeFilter")?.value || "").trim();
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  const fuelPrice = Math.max(0, Number(qs("fuelReconPrice")?.value || 0));
  if (!start || !end) return alert("Select start and end dates.");
  const mode = download ? "&download=1" : "";
  const url =
    `${API}/api/reports/fuel-reconciliation.pdf?start=${encodeURIComponent(start)}` +
    `&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}` +
    `&mode=${encodeURIComponent(modeFilter)}&asset_code=${encodeURIComponent(assetCode)}` +
    `&fuel_price=${encodeURIComponent(fuelPrice)}${mode}`;
  window.open(url, "_blank");
}

function openFuelReconciliationXlsx() {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const modeFilter = String(qs("fuelModeFilter")?.value || "").trim();
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  const fuelPrice = Math.max(0, Number(qs("fuelReconPrice")?.value || 0));
  if (!start || !end) return alert("Select start and end dates.");
  const url =
    `${API}/api/reports/fuel-reconciliation.xlsx?start=${encodeURIComponent(start)}` +
    `&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}` +
    `&mode=${encodeURIComponent(modeFilter)}&asset_code=${encodeURIComponent(assetCode)}` +
    `&fuel_price=${encodeURIComponent(fuelPrice)}`;
  window.open(url, "_blank");
}

async function downloadExecutivePackExcel() {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  if (!start || !end) return alert("Select start and end dates.");
  setStatus("Generating executive pack...");
  try {
    const q = new URLSearchParams();
    q.set("start", start);
    q.set("end", end);
    q.set("scheduled", "10");
    q.set("near_due_hours", "50");
    const res = await fetch(`${API}/api/reports/executive-pack.xlsx?${q.toString()}`, { headers: authHeaders() });
    if (!res.ok) {
      const txt = await res.text();
      throw new Error(txt || `Executive pack request failed (${res.status})`);
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `IRONLOG_Executive_Pack_${start}_to_${end}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    setStatus("Executive pack ready.");
  } catch (e) {
    setStatus("Executive pack error: " + (e.message || e));
    alert(`Executive pack error: ${e.message || e}`);
  }
}

function openFuelMachineHistoryPdf(assetCode, download = false) {
  const code = String(assetCode || "").trim();
  const start = (qs("fuelStart")?.value || "").trim() || (qs("fuelSnapStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim() || (qs("fuelSnapEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  if (!code) return alert("Select a machine first.");
  if (!start || !end) return alert("Select start and end dates.");
  const mode = download ? "&download=1" : "";
  const url =
    `${API}/api/reports/fuel-machine-history.pdf?asset_code=${encodeURIComponent(code)}` +
    `&start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}${mode}`;
  window.open(url, "_blank");
}

function fuelPeriodRange(anchorDate, period) {
  const end = new Date(`${anchorDate}T00:00:00`);
  if (Number.isNaN(end.getTime())) return null;
  const start = new Date(end);
  if (period === "daily") {
    // same day
  } else if (period === "weekly") {
    start.setDate(start.getDate() - 6);
  } else if (period === "monthly") {
    start.setDate(start.getDate() - 29);
  } else {
    return null;
  }
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

function maxDateStr(a, b) {
  return String(a || "") > String(b || "") ? String(a) : String(b);
}

async function loadFuelSnapshots() {
  const startInput = (qs("fuelSnapStart")?.value || "").trim();
  const endInput =
    (qs("fuelSnapEnd")?.value || "").trim() ||
    (qs("date")?.value || "").trim() ||
    new Date().toISOString().slice(0, 10);
  const startBase = startInput || endInput;
  if (!startBase || !endInput) return alert("Select snapshot start and end dates.");
  if (startBase > endInput) return alert("Snapshot start date cannot be after end date.");

  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const dailyBase = fuelPeriodRange(endInput, "daily");
  const weeklyBase = fuelPeriodRange(endInput, "weekly");
  const monthlyBase = fuelPeriodRange(endInput, "monthly");
  const ranges = [
    { key: "daily", label: "Daily", range: { start: maxDateStr(startBase, dailyBase.start), end: endInput } },
    { key: "weekly", label: "Weekly (7d)", range: { start: maxDateStr(startBase, weeklyBase.start), end: endInput } },
    { key: "monthly", label: "Monthly (30d)", range: { start: maxDateStr(startBase, monthlyBase.start), end: endInput } },
  ];

  setStatus("Loading fuel snapshots...");
  setSkeleton("fuelSnapshotsList", 3);
  const results = await Promise.all(
    ranges.map(async (p) => {
      const data = await fetchJson(
        `${API}/api/dashboard/fuel?start=${encodeURIComponent(p.range.start)}&end=${encodeURIComponent(p.range.end)}&tolerance=${encodeURIComponent(tolerance)}`
      );
      return { ...p, data };
    })
  );

  setText("fsDailyEx", Number(results.find((r) => r.key === "daily")?.data?.summary?.excessive_count || 0));
  setText("fsWeeklyEx", Number(results.find((r) => r.key === "weekly")?.data?.summary?.excessive_count || 0));
  setText("fsMonthlyEx", Number(results.find((r) => r.key === "monthly")?.data?.summary?.excessive_count || 0));

  const list = qs("fuelSnapshotsList");
  if (!list) return;
  list.innerHTML = "";

  for (const res of results) {
    const s = res.data?.summary || {};
    const rowTop = (res.data?.rows || []).slice(0, 5)
      .map((r) => {
        const flag = r.is_excessive ? "<span class='pill red'>EXCESSIVE</span>" : "<span class='pill blue'>OK</span>";
        const metric = String(r.metric_mode || "hours") === "km"
          ? `${r.actual_km_per_l == null ? "-" : Number(r.actual_km_per_l).toFixed(3)} km/L`
          : `${r.actual_lph == null ? "-" : Number(r.actual_lph).toFixed(3)} L/hr`;
        const runTxt = String(r.metric_mode || "hours") === "km"
          ? `Distance ${Number(r.km_run || 0).toFixed(1)}km`
          : `Hours ${Number(r.hours_run || 0).toFixed(1)}`;
        return `<small><b>${r.asset_code}</b> ${metric} ${flag} | Fuel ${Number(r.fuel_liters || 0).toFixed(1)}L | ${runTxt}</small>`;
      })
      .join("<br>");
    list.appendChild(
      item(
        `<b>${res.label}</b> <span class="pill">${res.range.start} to ${res.range.end}</span>` +
        `<br><small>Fuel: ${Number(s.fuel_liters || 0).toFixed(2)}L | Hours: ${Number(s.hours_run || 0).toFixed(2)} | Distance: ${Number(s.km_run || 0).toFixed(2)}km | Avg(L/hr): ${s.avg_lph == null ? "-" : Number(s.avg_lph).toFixed(3)} | Avg(km/L): ${s.avg_km_per_l == null ? "-" : Number(s.avg_km_per_l).toFixed(3)} | Excessive: ${Number(s.excessive_count || 0)}</small>` +
        (rowTop ? `<br>${rowTop}` : "<br><small>No data in this period.</small>")
      )
    );
  }
  setStatus("Fuel snapshots ready.");
}
async function loadFuelMachineDailyInline(assetCode, mountEl) {
  if (!mountEl) return;
  // Inline machine history should follow the Fuel Benchmark date window first.
  const start = (qs("fuelStart")?.value || "").trim() || (qs("fuelSnapStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim() || (qs("fuelSnapEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  if (!assetCode || !start || !end) return;

  mountEl.innerHTML = "<small>Loading machine history...</small>";
  const data = await fetchJson(
    `${API}/api/dashboard/fuel/daily?asset_code=${encodeURIComponent(assetCode)}&start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}`
  );
  const rows = Array.isArray(data.rows) ? data.rows : [];
  mountEl.setAttribute("data-code", String(assetCode));
  const mode = String(data.summary?.metric_mode || "hours");
  const isInlineFlagged = (r) => {
    if (r?.invalid_delta) return false;
    if (mode === "km") {
      const actual = Number(r?.actual_km_per_l);
      const benchmark = Number(r?.oem_km_per_l);
      if (!Number.isFinite(actual) || actual <= 0) return false; // no entries -> OK
      if (!Number.isFinite(benchmark) || benchmark <= 0) return false;
      return actual < benchmark;
    }
    const actual = Number(r?.actual_lph);
    const benchmark = Number(r?.oem_lph);
    if (!Number.isFinite(actual) || actual <= 0) return false; // no entries -> OK
    if (!Number.isFinite(benchmark) || benchmark <= 0) return false;
    return actual > benchmark;
  };
  const top = `<div class="fuel-inline-summary"><small><b>Fill days:</b> ${Number(data.summary?.days || 0)} | <b>Fuel:</b> ${Number(data.summary?.fuel_liters || 0).toFixed(2)}L | <b>Fill ${mode === "km" ? "distance" : "hours"}:</b> ${Number(mode === "km" ? (data.summary?.km_run || 0) : (data.summary?.hours_run || 0)).toFixed(2)} | <b>Avg:</b> ${mode === "km" ? (data.summary?.avg_km_per_l == null ? "-" : Number(data.summary.avg_km_per_l).toFixed(3) + " km/L") : (data.summary?.avg_lph == null ? "-" : Number(data.summary.avg_lph).toFixed(3) + " L/hr")} | <b>${mode === "km" ? "Under benchmark days" : "Over benchmark days"}:</b> ${Number(data.summary?.excessive_days || 0)}</small></div>`;
  const tableRows = rows.map((r) => {
    const flagged = isInlineFlagged(r);
    const invalid = Boolean(r?.invalid_delta);
    const statusClass = invalid ? "fh-status-excessive" : (flagged ? "fh-status-excessive" : "fh-status-ok");
    const statusText = invalid ? "INVALID DELTA" : (flagged ? (mode === "km" ? "UNDER BENCHMARK" : "EXCESSIVE") : "OK");
    const meterUnit = mode === "km" ? "km" : "hrs";
    const openMeterValue = r.open_meter_value == null ? "" : Number(r.open_meter_value).toFixed(2);
    const closeMeterValue = r.close_meter_value == null ? "" : Number(r.close_meter_value).toFixed(2);
    return (
      `<tr>` +
      `<td class="fh-col-date">${r.log_date}</td>` +
      `<td class="fh-col-num">${Number(r.fuel_liters || 0).toFixed(2)}</td>` +
      `<td class="fh-col-num"><input data-fuel-open-input="1" class="w-110" type="number" step="0.01" min="0" value="${openMeterValue}"> ${meterUnit}</td>` +
      `<td class="fh-col-num"><input data-fuel-close-input="1" class="w-110" type="number" step="0.01" min="0" value="${closeMeterValue}"> ${meterUnit}</td>` +
      `<td class="fh-col-num">${invalid ? "-" : Number((mode === "km" ? r.km_run : r.hours_run) || 0).toFixed(2)}</td>` +
      `<td class="fh-col-num">${invalid ? "-" : (mode === "km" ? (r.actual_km_per_l == null ? "-" : Number(r.actual_km_per_l).toFixed(3)) : (r.actual_lph == null ? "-" : Number(r.actual_lph).toFixed(3)))}</td>` +
      `<td class="fh-col-status"><span class="fh-status ${statusClass}">${statusText}</span></td>` +
      `<td class="fh-col-action"><button data-fuel-save="${Number(r.id || 0)}">Save</button> <button data-fuel-delete="${Number(r.id || 0)}">Delete</button></td>` +
      `</tr>`
    );
  }).join("");
  mountEl.innerHTML =
    `${top}<br>` +
    (tableRows
      ? `<div class="fuel-history-table-wrap"><table class="fuel-history-table"><colgroup><col style="width:14%"><col style="width:12%"><col style="width:14%"><col style="width:14%"><col style="width:14%"><col style="width:10%"><col style="width:10%"><col style="width:12%"></colgroup><thead><tr><th class="fh-col-date">Date</th><th class="fh-col-num">Fuel (L)</th><th class="fh-col-num">Open ${mode === "km" ? "km" : "hrs"}</th><th class="fh-col-num">Close ${mode === "km" ? "km" : "hrs"}</th><th class="fh-col-num">${mode === "km" ? "Distance Between Fills (km)" : "Hours Between Fills"}</th><th class="fh-col-num">${mode === "km" ? "km/L" : "L/hr"}</th><th class="fh-col-status">Status</th><th class="fh-col-action">Action</th></tr></thead><tbody>${tableRows}</tbody></table></div>`
      : "<small>No filled days found for this machine in selected range.</small>");
}

async function deleteFuelLogEntry(logId) {
  const id = Number(logId || 0);
  if (!Number.isInteger(id) || id <= 0) return;
  const ok = confirm("Delete this fuel input entry?");
  if (!ok) return;
  await fetchJson(`${API}/api/dashboard/fuel/log/${id}`, { method: "DELETE" });
}

async function loadStockMonitor() {
  const filter = (qs("stockPartFilter")?.value || "").trim();
  const page = window.stockMonitorPage || 1;
  const pageSize = 20;
  const q = filter ? `?part_code=${encodeURIComponent(filter)}` : "";
  const data = await fetchJson(`${API}/api/stock/monitor${q}`);

  setText("smBelowMin", Number(data.summary?.below_min || 0));
  setText("smCriticalBelow", Number(data.summary?.critical_below_min || 0));
  setText("smTotalParts", Number(data.summary?.total_parts || 0));

  const list = qs("stockMonitorList");
  if (!list) return;
  list.innerHTML = "";
  const rows = data.rows || [];
  const start = (page - 1) * pageSize;
  const end = start + pageSize;
  rows.slice(start, end).forEach((r) => {
    list.appendChild(
      item(
        `<b>${r.part_code}</b> – ${Number(r.on_hand || 0).toFixed(1)} on hand ${
          r.below_min ? "<span class='pill red'>LOW</span>" : ""
        }<br><small>${r.part_name || ""} | Min: ${Number(r.min_stock || 0).toFixed(1)}</small>`
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No parts found for current filter.</small>"));

  // Update paging info
  const pageInfo = qs("stockPageInfo");
  if (pageInfo) {
    const total = rows.length;
    const totalPages = Math.max(1, Math.ceil(total / pageSize));
    pageInfo.textContent = `Page ${page} of ${totalPages}`;
  }
}

// Paging controls
window.stockMonitorPage = 1;
function updateStockMonitorPage(delta) {
  const rows = window.lastStockMonitorRows || [];
  const pageSize = 20;
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  window.stockMonitorPage = Math.max(1, Math.min(window.stockMonitorPage + delta, totalPages));
  loadStockMonitor();
}

// Live filter
const stockPartFilter = qs("stockPartFilter");
if (stockPartFilter) {
  stockPartFilter.addEventListener("input", () => {
    window.stockMonitorPage = 1;
    loadStockMonitor();
  });
}

const prevBtn = qs("prevStockPage");
if (prevBtn) prevBtn.onclick = () => updateStockMonitorPage(-1);
const nextBtn = qs("nextStockPage");
if (nextBtn) nextBtn.onclick = () => updateStockMonitorPage(1);

// Save last rows for paging
const origLoadStockMonitor = loadStockMonitor;
loadStockMonitor = async function() {
  const filter = (qs("stockPartFilter")?.value || "").trim();
  const q = filter ? `?part_code=${encodeURIComponent(filter)}` : "";
  const data = await fetchJson(`${API}/api/stock/monitor${q}`);
  window.lastStockMonitorRows = data.rows || [];
  // Call original logic
  await origLoadStockMonitor.apply(this, arguments);
};

let stockPageData = { rows: [], recent: [], summary: null };

function filterStockDisplayRows(rows) {
  const q = (qs("spFilter")?.value || "").trim().toLowerCase();
  const arr = Array.isArray(rows) ? rows : [];
  if (!q) return arr;
  return arr.filter(
    (r) =>
      String(r.part_code || "").toLowerCase().includes(q) ||
      String(r.part_name || "").toLowerCase().includes(q)
  );
}

function renderStockInventoryTable(rows) {
  const host = qs("spList");
  if (!host) return;
  if (!rows.length) {
    host.innerHTML = `<div class="stores-inventory-empty muted small">No parts found for current filter.</div>`;
    return;
  }

  const bodyRows = rows
    .map((r) => {
      const onHand = Number(r.on_hand || 0);
      const min = Number(r.min_stock || 0);
      const unit = Number(r.unit_cost || 0);
      const value = Number(r.stock_value ?? onHand * unit);
      const below = Boolean(r.below_min);
      const critical = Boolean(r.critical);
      const rowCls = [
        "stores-inv-row",
        below ? "stores-inv-row--low" : "",
        critical && !below ? "stores-inv-row--critical" : "",
      ]
        .filter(Boolean)
        .join(" ");
      const status = below
        ? `<span class="stores-inv-status stores-inv-status--low">Low</span>`
        : critical
          ? `<span class="stores-inv-status stores-inv-status--watch">Critical</span>`
          : `<span class="stores-inv-status stores-inv-status--ok">OK</span>`;
      const flag = below
        ? `<span class="stores-inv-flag" title="Below minimum stock">!</span>`
        : `<span class="stores-inv-flag stores-inv-flag--clear" aria-hidden="true"></span>`;

      return `<tr class="${rowCls}">
        <td class="stores-inv-col-flag">${flag}</td>
        <td class="stores-inv-col-code"><span class="stores-inv-code">${escapeHtml(r.part_code || "")}</span></td>
        <td class="stores-inv-col-desc">${escapeHtml(r.part_name || "—")}</td>
        <td class="stores-inv-col-num">${onHand.toFixed(1)}</td>
        <td class="stores-inv-col-num">${min.toFixed(1)}</td>
        <td class="stores-inv-col-num">$${unit.toFixed(2)}</td>
        <td class="stores-inv-col-num stores-inv-col-value">$${value.toFixed(2)}</td>
        <td class="stores-inv-col-status">${status}</td>
      </tr>`;
    })
    .join("");

  host.innerHTML = `
    <div class="stores-inventory-scroll">
      <table class="stores-inventory-table">
        <thead>
          <tr>
            <th class="stores-inv-col-flag" scope="col" title="Reorder needed">!</th>
            <th class="stores-inv-col-code" scope="col">Part code</th>
            <th class="stores-inv-col-desc" scope="col">Description</th>
            <th class="stores-inv-col-num" scope="col">On hand</th>
            <th class="stores-inv-col-num" scope="col">Min stock</th>
            <th class="stores-inv-col-num" scope="col">Unit price</th>
            <th class="stores-inv-col-num" scope="col">Stock value</th>
            <th class="stores-inv-col-status" scope="col">Status</th>
          </tr>
        </thead>
        <tbody>${bodyRows}</tbody>
      </table>
    </div>
    <div class="stores-inventory-foot muted small">${rows.length} item${rows.length === 1 ? "" : "s"} shown</div>
  `;
}

function renderStockRecentTable(recent) {
  const host = qs("spRecent");
  if (!host) return;
  const rows = Array.isArray(recent) ? recent : [];
  if (!rows.length) {
    host.innerHTML = `<div class="stores-inventory-empty muted small">No stock movements yet.</div>`;
    return;
  }

  const bodyRows = rows
    .map((r) => {
      const qty = Number(r.quantity || 0);
      const qtyCls = qty >= 0 ? "stores-mv-qty--in" : "stores-mv-qty--out";
      return `<tr>
        <td class="stores-mv-col-date">${escapeHtml(String(r.created_at || "").slice(0, 16))}</td>
        <td class="stores-mv-col-code"><span class="stores-inv-code">${escapeHtml(r.part_code || "")}</span></td>
        <td class="stores-mv-col-desc">${escapeHtml(r.part_name || "—")}</td>
        <td class="stores-mv-col-num ${qtyCls}">${qty >= 0 ? "+" : ""}${qty.toFixed(1)}</td>
        <td class="stores-mv-col-type">${escapeHtml(r.movement_type || "—")}</td>
        <td class="stores-mv-col-loc">${escapeHtml(r.location_code || "NO-LOC")}</td>
        <td class="stores-mv-col-ref">${escapeHtml(r.reference || "—")}</td>
      </tr>`;
    })
    .join("");

  host.innerHTML = `
    <div class="stores-inventory-scroll stores-movements-scroll">
      <table class="stores-inventory-table stores-movements-table">
        <thead>
          <tr>
            <th scope="col">Date</th>
            <th scope="col">Part code</th>
            <th scope="col">Description</th>
            <th class="stores-inv-col-num" scope="col">Qty</th>
            <th scope="col">Type</th>
            <th scope="col">Location</th>
            <th scope="col">Reference</th>
          </tr>
        </thead>
        <tbody>${bodyRows}</tbody>
      </table>
    </div>
  `;
}

function refreshStockInventoryDisplay() {
  const onlyLow = Boolean(qs("spOnlyLow")?.checked);
  let baseRows = filterStockDisplayRows(stockPageData.rows);
  if (onlyLow) baseRows = baseRows.filter((r) => Boolean(r.below_min));
  renderStockInventoryTable(sortStockRows(baseRows));
}

function sortStockRows(rows) {
  const mode = (qs("spSort")?.value || "critical_then_low").trim();
  const arr = Array.isArray(rows) ? [...rows] : [];

  if (mode === "on_hand_asc") {
    return arr.sort((a, b) => Number(a.on_hand || 0) - Number(b.on_hand || 0));
  }
  if (mode === "on_hand_desc") {
    return arr.sort((a, b) => Number(b.on_hand || 0) - Number(a.on_hand || 0));
  }
  if (mode === "part_code_desc") {
    return arr.sort((a, b) => String(b.part_code || "").localeCompare(String(a.part_code || "")));
  }
  if (mode === "part_code_asc") {
    return arr.sort((a, b) => String(a.part_code || "").localeCompare(String(b.part_code || "")));
  }

  // default: critical first, then below min, then lowest on hand
  return arr.sort((a, b) => {
    const c = Number(Boolean(b.critical)) - Number(Boolean(a.critical));
    if (c !== 0) return c;
    const low = Number(Boolean(b.below_min)) - Number(Boolean(a.below_min));
    if (low !== 0) return low;
    return Number(a.on_hand || 0) - Number(b.on_hand || 0);
  });
}

async function loadStockOnHandPage() {
  const filter = (qs("spFilter")?.value || "").trim();
  const q = filter ? `?part_code=${encodeURIComponent(filter)}` : "";

  setStatus("Loading stock on hand...");
  setSkeleton("spList", 2);
  setSkeleton("spRecent", 2);

  const data = await fetchJson(`${API}/api/stock/monitor${q}`);
  stockPageData = {
    rows: Array.isArray(data.rows) ? data.rows : [],
    recent: Array.isArray(data.recent) ? data.recent : [],
    summary: data.summary || null,
  };

  setText("spTotalParts", Number(data.summary?.total_parts || 0));
  setText("spBelowMin", Number(data.summary?.below_min || 0));
  setText("spCriticalBelow", Number(data.summary?.critical_below_min || 0));
  setText("spTotalOnHand", Number(data.summary?.total_on_hand || 0).toFixed(1));
  setText("spTotalValue", `$${Number(data.summary?.total_stock_value || 0).toFixed(2)}`);

  refreshStockInventoryDisplay();
  renderStockRecentTable(stockPageData.recent);

  setStatus("Stock on hand ready.");
}

function exportStockOnHandCsv() {
  const onlyLow = Boolean(qs("spOnlyLow")?.checked);
  const baseRows = Array.isArray(stockPageData.rows) ? stockPageData.rows : [];
  const rows = onlyLow ? baseRows.filter((r) => Boolean(r.below_min)) : baseRows;
  if (!rows.length) return alert("Load stock data first.");

  const header = "part_code,part_name,on_hand,min_stock,unit_cost,stock_value,critical,below_min";
  const lines = rows.map((r) =>
    [
      r.part_code || "",
      `"${String(r.part_name || "").replace(/"/g, '""')}"`,
      Number(r.on_hand || 0),
      Number(r.min_stock || 0),
      Number(r.unit_cost || 0),
      Number(r.stock_value || 0),
      r.critical ? 1 : 0,
      r.below_min ? 1 : 0,
    ].join(",")
  );
  const csv = [header, ...lines].join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "stock_on_hand.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  setStatus("Stock CSV exported.");
}

/** Last loaded stock movements report (period ledger). */
let stockMovementsReportData = { rows: [], date_from: "", date_to: "", part_filter: "", truncated: false };

function ensureStockMovementsReportDates() {
  const fromEl = qs("smrDateFrom");
  const toEl = qs("smrDateTo");
  if (!fromEl || !toEl) return;
  if (!fromEl.value || !toEl.value) {
    const today = new Date();
    const start = new Date(today.getFullYear(), today.getMonth(), 1);
    const pad = (n) => String(n).padStart(2, "0");
    const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
    if (!fromEl.value) fromEl.value = ymd(start);
    if (!toEl.value) toEl.value = ymd(today);
  }
}

function normalizeStockReportPartInput(raw) {
  let s = String(raw || "").trim().replace(/\s+/g, " ");
  if (!s) return "";
  const dashIdx = s.indexOf(" - ");
  if (dashIdx > 0) s = s.slice(0, dashIdx).trim();
  return s;
}

async function loadStockMovementsReport() {
  ensureStockMovementsReportDates();
  const date_from = (qs("smrDateFrom")?.value || "").trim();
  const date_to = (qs("smrDateTo")?.value || "").trim();
  const part_code = normalizeStockReportPartInput(qs("smrPartFilter")?.value || "");
  if (!date_from || !date_to) return alert("Choose From and To dates.");

  const q = new URLSearchParams();
  q.set("date_from", date_from);
  q.set("date_to", date_to);
  if (part_code) q.set("part_code", part_code);

  setStatus("Loading stock movements report...");
  setSkeleton("smrList", 2);

  const data = await fetchJson(`${API}/api/stock/movements-report?${q.toString()}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  const summary = data.summary || {};

  stockMovementsReportData = {
    rows,
    date_from: data.date_from || date_from,
    date_to: data.date_to || date_to,
    part_filter: data.part_filter || part_code || "",
    truncated: Boolean(data.truncated),
    total_matching: Number(data.total_matching || rows.length),
    row_limit: Number(data.row_limit || 0),
  };

  setText("smrMoveCount", String(summary.movement_count ?? "-"));
  setText("smrQtyIn", summary.qty_in != null ? Number(summary.qty_in).toFixed(2) : "-");
  setText("smrQtyOut", summary.qty_out != null ? Number(summary.qty_out).toFixed(2) : "-");
  setText("smrNetQty", summary.net_qty != null ? Number(summary.net_qty).toFixed(2) : "-");

  const trunc = qs("smrTruncNote");
  const truncText = qs("smrTruncText");
  if (trunc && truncText) {
    if (stockMovementsReportData.truncated) {
      trunc.style.display = "";
      truncText.textContent = `Showing the latest ${rows.length} of ${stockMovementsReportData.total_matching} movements in this period (export CSV includes loaded rows only). Increase precision with a narrower date range or part filter.`;
    } else {
      trunc.style.display = "none";
      truncText.textContent = "";
    }
  }

  const list = qs("smrList");
  if (list) {
    list.innerHTML = "";
    rows.forEach((r) => {
      const qty = Number(r.quantity || 0);
      const loc = r.location_code || "—";
      const bin = r.bin_code ? String(r.bin_code) : "";
      list.appendChild(
        item(
          `<b>${r.part_code || ""}</b> — ${qty.toFixed(2)} (${r.movement_type || ""})` +
            `<br><small>${r.movement_at || ""} | ${loc}${bin ? " / " + bin : ""} | ${r.reference || "—"}</small>` +
            `<br><small>${r.part_name || ""}</small>`
        )
      );
    });
    if (!rows.length) list.appendChild(item("<small>No movements in this period for the current filter.</small>"));
  }

  setStatus("Stock movements report ready.");
}

function exportStockMovementsReportCsv() {
  const { rows, date_from, date_to } = stockMovementsReportData;
  if (!rows.length) return alert("Load the stock movements report first.");

  const header =
    "movement_at,part_code,part_name,movement_type,quantity,reference,location_code,bin_code";
  const lines = rows.map((r) =>
    [
      r.movement_at || "",
      r.part_code || "",
      `"${String(r.part_name || "").replace(/"/g, '""')}"`,
      r.movement_type || "",
      Number(r.quantity || 0),
      `"${String(r.reference || "").replace(/"/g, '""')}"`,
      r.location_code || "",
      r.bin_code || "",
    ].join(",")
  );
  const csv = [header, ...lines].join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `stock_movements_${date_from || "from"}_${date_to || "to"}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  setStatus("Stock movements CSV exported.");
}

function openStockMovementsReportPdf() {
  ensureStockMovementsReportDates();
  const date_from = (qs("smrDateFrom")?.value || "").trim();
  const date_to = (qs("smrDateTo")?.value || "").trim();
  const part_code = normalizeStockReportPartInput(qs("smrPartFilter")?.value || "");
  if (!date_from || !date_to) return alert("Choose From and To dates.");
  const q = new URLSearchParams();
  q.set("date_from", date_from);
  q.set("date_to", date_to);
  if (part_code) q.set("part_code", part_code);
  window.open(`${API}/api/reports/stock-movements.pdf?${q.toString()}`, "_blank");
}

function openStockOnHandPdf() {
  const filter = (qs("spFilter")?.value || "").trim();
  const q = filter ? `?part_code=${encodeURIComponent(filter)}` : "";
  window.open(`${API}/api/reports/stock-monitor.pdf${q}`, "_blank");
}

let storesPartOrdersCache = [];

function ensureStoresPartOrderDates() {
  const fromEl = qs("spoDateFrom");
  const toEl = qs("spoDateTo");
  const orderEl = qs("spoOrderDate");
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), 1);
  const pad = (n) => String(n).padStart(2, "0");
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  if (fromEl && !fromEl.value) fromEl.value = ymd(start);
  if (toEl && !toEl.value) toEl.value = ymd(today);
  if (orderEl && !orderEl.value) orderEl.value = ymd(today);
}

function spoStatusLabel(status) {
  const s = String(status || "").toLowerCase();
  if (s === "on_order") return "On order";
  if (s === "in_transit") return "In transit";
  if (s === "arrived") return "Arrived";
  if (s === "cancelled") return "Cancelled";
  return s || "—";
}

function moneyUsd(n) {
  return Number(n || 0).toFixed(2);
}

function renderStoresPartOrdersSummary(summary) {
  const s = summary || {};
  setText("spoOnOrderValue", moneyUsd(s.on_order?.value));
  setText("spoInTransitValue", moneyUsd(s.in_transit?.value));
  setText("spoArrivedValue", moneyUsd(s.arrived?.value));
  setText("spoPendingValue", moneyUsd(s.total_pending));
  setText("spoForecastValue", moneyUsd(s.total_forecast));
}

function handleStoresPartOrderReceipt(data) {
  const receipt = data?.stock_receipt;
  if (!receipt) return;
  if (receipt.received) {
    setStatus(
      `Received ${receipt.qty} × ${receipt.part_code} into store inventory (${receipt.on_hand_after} on hand).`
    );
    loadStockOnHandPage().catch(() => {});
    loadStockMonitor().catch(() => {});
    return;
  }
  if (receipt.already) {
    setStatus("Purchase already received into store inventory.");
    return;
  }
  if (receipt.error) {
    alert(`Status saved, but stock was not updated: ${receipt.error}`);
  }
}

function spoAttrVal(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;");
}

function renderStoresPartOrdersTable(rows) {
  const host = qs("spoList");
  if (!host) return;
  if (!Array.isArray(rows) || !rows.length) {
    host.innerHTML = `<div class="muted small">No purchases in this period. Add a line above.</div>`;
    return;
  }
  host.innerHTML = `
    <table class="gridTable spo-purchases-table" style="min-width:1200px;">
      <thead>
        <tr>
          <th>Order date</th>
          <th>Part</th>
          <th style="text-align:right;">Qty</th>
          <th style="text-align:right;">Unit $</th>
          <th style="text-align:right;">Line $</th>
          <th>Supplier</th>
          <th>PO</th>
          <th>Req #</th>
          <th>ETA</th>
          <th>Status</th>
          <th>Inventory</th>
          <th></th>
        </tr>
      </thead>
      <tbody>
        ${rows.map((r) => {
          const id = Number(r.id || 0);
          const cancelled = String(r.status || "").toLowerCase() === "cancelled";
          const partLabel = r.part_code
            ? `<strong>${String(r.part_code).replace(/</g, "&lt;")}</strong><br><small class="muted">${String(r.part_name || "").replace(/</g, "&lt;")}</small>`
            : String(r.part_name || "").replace(/</g, "&lt;");
          const statusOpts = ["on_order", "in_transit", "arrived", "cancelled"]
            .map((st) => `<option value="${st}"${String(r.status || "").toLowerCase() === st ? " selected" : ""}>${spoStatusLabel(st)}</option>`)
            .join("");
          const inStore = Boolean(r.in_store_inventory || r.stock_movement_id);
          const isArrived = String(r.status || "").toLowerCase() === "arrived";
          const inventoryCell = inStore
            ? `<span class="pill green" title="Received into store stock">In store</span>`
            : isArrived
              ? `<button type="button" class="btn btn-secondary btn-sm" data-spo-receive="${id}">Receive to store</button>`
              : "—";
          const dis = cancelled ? " disabled" : "";
          return `
            <tr data-spo-row="${id}">
              <td>${String(r.order_date || "").replace(/</g, "&lt;")}</td>
              <td>${partLabel}</td>
              <td style="text-align:right;">${Number(r.qty || 0)}</td>
              <td style="text-align:right;">${moneyUsd(r.unit_cost)}</td>
              <td style="text-align:right;"><strong>${moneyUsd(r.line_total)}</strong></td>
              <td>
                <input type="text" class="spo-inline-input w-full" data-spo-field="supplier_name" value="${spoAttrVal(r.supplier_name || "")}" placeholder="Supplier"${dis} />
              </td>
              <td>
                <input type="text" class="spo-inline-input w-full" data-spo-field="po_number" value="${spoAttrVal(r.po_number || "")}" placeholder="PO when issued"${dis} />
              </td>
              <td>
                <input type="text" class="spo-inline-input w-full" data-spo-field="requisition_number" value="${spoAttrVal(r.requisition_number || "")}" placeholder="Req #"${dis} />
              </td>
              <td>
                <input type="date" class="spo-inline-input w-full" data-spo-field="expected_arrival_date" value="${spoAttrVal(r.expected_arrival_date || "")}"${dis} />
              </td>
              <td>
                <select data-spo-status="${id}" class="w-full" style="min-width:120px;"${inStore ? " disabled title=\"Already in store inventory\"" : dis}>${statusOpts}</select>
              </td>
              <td>${inventoryCell}</td>
              <td class="spo-row-actions">
                ${cancelled
                  ? `<span class="muted small">Cancelled</span>`
                  : `<button type="button" class="btn btn-primary btn-sm" data-spo-save="${id}">Save</button>
                     <button type="button" class="btn btn-secondary btn-sm" data-spo-del="${id}">Cancel</button>`}
              </td>
            </tr>
          `;
        }).join("")}
      </tbody>
    </table>
    <p class="muted small" style="margin-top:8px;">Edit Req # or PO on each line, then click <strong>Save</strong>. Change status to Arrived when goods land to post into store inventory.</p>
  `;
}

function readStoresPartOrderRowPatch(rowEl) {
  if (!rowEl) return null;
  const read = (field) => {
    const el = rowEl.querySelector(`[data-spo-field="${field}"]`);
    if (!el) return null;
    const v = String(el.value || "").trim();
    return v || null;
  };
  const statusEl = rowEl.querySelector("select[data-spo-status]");
  const patch = {
    supplier_name: read("supplier_name"),
    po_number: read("po_number"),
    requisition_number: read("requisition_number"),
    expected_arrival_date: read("expected_arrival_date"),
  };
  if (statusEl && !statusEl.disabled) {
    patch.status = String(statusEl.value || "").trim();
  }
  return patch;
}

async function patchStoresPartOrder(id, patch) {
  const n = Number(id || 0);
  if (!n) return null;
  const data = await fetchJson(`${API}/api/stock/part-orders/${n}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(patch || {}),
  });
  handleStoresPartOrderReceipt(data);
  return data;
}

async function saveStoresPartOrderRow(id) {
  const host = qs("spoList");
  const rowEl = host?.querySelector(`tr[data-spo-row="${Number(id)}"]`);
  const patch = readStoresPartOrderRowPatch(rowEl);
  if (!patch) return;
  await patchStoresPartOrder(id, patch);
  setStatus("Purchase line saved.");
  await loadStoresPartOrders();
}

async function loadStoresPartOrders() {
  ensureStoresPartOrderDates();
  const start = (qs("spoDateFrom")?.value || "").trim();
  const end = (qs("spoDateTo")?.value || "").trim();
  const status = (qs("spoFilterStatus")?.value || "").trim();
  if (!start || !end) return alert("Choose period from and to dates.");

  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  if (status) q.set("status", status);

  setStatus("Loading parts purchases...");
  setSkeleton("spoList", 2);
  const data = await fetchJson(`${API}/api/stock/part-orders?${q.toString()}`);
  storesPartOrdersCache = Array.isArray(data.rows) ? data.rows : [];
  renderStoresPartOrdersSummary(data.summary);
  renderStoresPartOrdersTable(storesPartOrdersCache);
  setStatus("Parts purchases loaded.");
}

function clearStoresPartOrderForm() {
  ["spoPartCode", "spoPartName", "spoSupplier", "spoPoNumber", "spoRequisitionNumber", "spoNotes"].forEach((id) => {
    const el = qs(id);
    if (el) el.value = "";
  });
  if (qs("spoQty")) qs("spoQty").value = "1";
  if (qs("spoUnitCost")) qs("spoUnitCost").value = "0";
  if (qs("spoStatus")) qs("spoStatus").value = "on_order";
  if (qs("spoExpectedDate")) qs("spoExpectedDate").value = "";
  ensureStoresPartOrderDates();
  const msg = qs("spoFormMsg");
  if (msg) msg.textContent = "";
}

async function saveStoresPartOrder() {
  ensureStoresPartOrderDates();
  const msg = qs("spoFormMsg");
  const part_code = normalizeStockReportPartInput(qs("spoPartCode")?.value || "");
  const part_name = String(qs("spoPartName")?.value || "").trim();
  const qty = Number(qs("spoQty")?.value || 1);
  const unit_cost = Number(qs("spoUnitCost")?.value || 0);
  const supplier_name = String(qs("spoSupplier")?.value || "").trim();
  const po_number = String(qs("spoPoNumber")?.value || "").trim();
  const requisition_number = String(qs("spoRequisitionNumber")?.value || "").trim();
  const order_date = (qs("spoOrderDate")?.value || "").trim();
  const expected_arrival_date = (qs("spoExpectedDate")?.value || "").trim();
  const status = String(qs("spoStatus")?.value || "on_order").trim();
  const notes = String(qs("spoNotes")?.value || "").trim();

  if (!part_code && !part_name) {
    if (msg) msg.textContent = "Enter a part code or description.";
    return;
  }
  if (!order_date) {
    if (msg) msg.textContent = "Order date is required.";
    return;
  }
  if (!Number.isFinite(qty) || qty <= 0) {
    if (msg) msg.textContent = "Quantity must be greater than zero.";
    return;
  }

  if (msg) msg.textContent = "Saving...";
  const data = await fetchJson(`${API}/api/stock/part-orders`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      part_code,
      part_name,
      qty,
      unit_cost,
      supplier_name,
      po_number: po_number || null,
      requisition_number: requisition_number || null,
      order_date,
      expected_arrival_date: expected_arrival_date || null,
      status,
      notes,
    }),
  });
  handleStoresPartOrderReceipt(data);
  if (msg) msg.textContent = data?.stock_receipt?.received
    ? "Purchase saved and received into store."
    : "Purchase saved.";
  clearStoresPartOrderForm();
  await loadStoresPartOrders();
}

async function updateStoresPartOrderStatus(id, status) {
  const n = Number(id || 0);
  if (!n || !status) return;
  const host = qs("spoList");
  const rowEl = host?.querySelector(`tr[data-spo-row="${n}"]`);
  const patch = readStoresPartOrderRowPatch(rowEl) || {};
  patch.status = status;
  await patchStoresPartOrder(n, patch);
  await loadStoresPartOrders();
}

async function receiveStoresPartOrderToInventory(id) {
  const n = Number(id || 0);
  if (!n) return;
  const data = await fetchJson(`${API}/api/stock/part-orders/${n}/receive`, { method: "POST" });
  handleStoresPartOrderReceipt(data);
  await loadStoresPartOrders();
}

async function cancelStoresPartOrder(id) {
  const n = Number(id || 0);
  if (!n) return;
  if (!confirm("Cancel this purchase line?")) return;
  await fetchJson(`${API}/api/stock/part-orders/${n}`, { method: "DELETE" });
  await loadStoresPartOrders();
}

function buildStoresPartOrdersExportQuery() {
  ensureStoresPartOrderDates();
  const start = (qs("spoDateFrom")?.value || "").trim();
  const end = (qs("spoDateTo")?.value || "").trim();
  const status = (qs("spoFilterStatus")?.value || "").trim();
  if (!start || !end) throw new Error("Choose period from and to dates.");
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  if (status) q.set("status", status);
  return { start, end, q };
}

function openStoresPartOrdersPdf(download = false) {
  const { q } = buildStoresPartOrdersExportQuery();
  if (download) q.set("download", "1");
  openAuthedPdf(`${API}/api/reports/part-orders.pdf?${q.toString()}`).catch((e) =>
    setStatus("Parts purchases PDF error: " + (e.message || e))
  );
}

async function exportStoresPartOrdersXlsx() {
  const { start, end, q } = buildStoresPartOrdersExportQuery();
  setStatus("Generating Excel...");
  const res = await fetch(`${API}/api/reports/part-orders.xlsx?${q.toString()}`, { headers: authHeaders() });
  if (!res.ok) {
    let msg = await res.text().catch(() => "");
    try {
      const j = JSON.parse(msg);
      msg = j.error || j.message || msg;
    } catch {}
    throw new Error(msg || `Export failed (${res.status})`);
  }
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `IRONLOG_Parts_Purchases_${start}_${end}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  setStatus("Parts purchases Excel downloaded.");
}

/* ========== Parts Tracking tab (parts + off-site repairs) ========== */
let ptPartsCache = [];
let ptOffsiteCache = [];

function ptTodayYmd() {
  return new Date().toISOString().slice(0, 10);
}

function ensurePartsTrackingDates() {
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), 1);
  const pad = (n) => String(n).padStart(2, "0");
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  if (qs("ptPartsFrom") && !qs("ptPartsFrom").value) qs("ptPartsFrom").value = ymd(start);
  if (qs("ptPartsTo") && !qs("ptPartsTo").value) qs("ptPartsTo").value = ymd(today);
  if (qs("ptOrderDate") && !qs("ptOrderDate").value) qs("ptOrderDate").value = ymd(today);
  if (qs("ptOffSent") && !qs("ptOffSent").value) qs("ptOffSent").value = ymd(today);
}

function ptIsOverdueEta(eta, statusDone) {
  if (statusDone) return false;
  const d = String(eta || "").slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return false;
  return d < ptTodayYmd();
}

function updatePartsTrackingKpis() {
  const partsPending = (ptPartsCache || [])
    .filter((r) => ["on_order", "in_transit"].includes(String(r.status || "").toLowerCase()))
    .reduce((s, r) => s + Number(r.line_total || 0), 0);
  const partsOverdue = (ptPartsCache || []).filter((r) =>
    ptIsOverdueEta(r.expected_arrival_date, String(r.status || "").toLowerCase() === "arrived" || String(r.status || "").toLowerCase() === "cancelled")
  ).length;
  const offEst = (ptOffsiteCache || []).reduce((s, r) => s + Number(r.estimated_cost || 0), 0);
  const offAct = (ptOffsiteCache || []).reduce((s, r) => s + Number(r.actual_cost || 0), 0);
  const offOverdue = (ptOffsiteCache || []).filter((r) =>
    ptIsOverdueEta(r.expected_return_date, String(r.repair_status || "").toLowerCase() === "returned")
  ).length;
  setText("ptPartsPending", moneyUsd(partsPending));
  setText("ptPartsOverdue", String(partsOverdue));
  setText("ptOffsiteEst", moneyUsd(offEst));
  setText("ptOffsiteActual", moneyUsd(offAct));
  setText("ptOffsiteOverdue", String(offOverdue));
}

function clearPtPartsForm() {
  ["ptPartCode", "ptPartName", "ptInvoice", "ptLocation", "ptEta", "ptSupplier", "ptPo", "ptNotes"].forEach((id) => {
    if (qs(id)) qs(id).value = "";
  });
  if (qs("ptQty")) qs("ptQty").value = "1";
  if (qs("ptUnitCost")) qs("ptUnitCost").value = "0";
  if (qs("ptStatus")) qs("ptStatus").value = "on_order";
  if (qs("ptOrderDate")) qs("ptOrderDate").value = ptTodayYmd();
  const msg = qs("ptPartsMsg");
  if (msg) msg.textContent = "";
}

function renderPtPartsTable(rows) {
  const host = qs("ptPartsList");
  if (!host) return;
  const q = String(qs("ptPartsSearch")?.value || "").trim().toLowerCase();
  let list = Array.isArray(rows) ? rows : [];
  if (q) {
    list = list.filter((r) => {
      const hay = `${r.part_code || ""} ${r.part_name || ""} ${r.invoice_number || ""} ${r.current_location || ""} ${r.supplier_name || ""} ${r.po_number || ""}`.toLowerCase();
      return hay.includes(q);
    });
  }
  if (!list.length) {
    host.innerHTML = `<div class="muted small">No purchase lines for this filter.</div>`;
    return;
  }
  host.innerHTML = `
    <table class="gridTable" style="min-width:1280px;">
      <thead>
        <tr>
          <th>Status</th>
          <th>Part</th>
          <th style="text-align:right;">Qty</th>
          <th style="text-align:right;">Unit $</th>
          <th style="text-align:right;">Line $</th>
          <th>Invoice #</th>
          <th>Ordered</th>
          <th>Location</th>
          <th>ETA on site</th>
          <th>Supplier / PO</th>
          <th></th>
        </tr>
      </thead>
      <tbody>
        ${list.map((r) => {
          const id = Number(r.id || 0);
          const cancelled = String(r.status || "").toLowerCase() === "cancelled";
          const overdue = ptIsOverdueEta(r.expected_arrival_date, cancelled || String(r.status || "").toLowerCase() === "arrived");
          const partLabel = r.part_code
            ? `<strong>${escapeHtml(String(r.part_code))}</strong><br><small class="muted">${escapeHtml(String(r.part_name || ""))}</small>`
            : escapeHtml(String(r.part_name || ""));
          const statusOpts = ["on_order", "in_transit", "arrived", "cancelled"]
            .map((st) => `<option value="${st}"${String(r.status || "").toLowerCase() === st ? " selected" : ""}>${spoStatusLabel(st)}</option>`)
            .join("");
          const dis = cancelled ? " disabled" : "";
          return `
            <tr data-pt-row="${id}" class="${overdue ? "pt-row-overdue" : ""}">
              <td><select data-pt-status="${id}" class="w-full" style="min-width:110px;"${dis}>${statusOpts}</select></td>
              <td>${partLabel}</td>
              <td style="text-align:right;">${Number(r.qty || 0)}</td>
              <td style="text-align:right;">${moneyUsd(r.unit_cost)}</td>
              <td style="text-align:right;"><strong>${moneyUsd(r.line_total)}</strong></td>
              <td><input type="text" class="spo-inline-input w-full" data-pt-field="invoice_number" value="${spoAttrVal(r.invoice_number || "")}"${dis} /></td>
              <td>${escapeHtml(String(r.order_date || ""))}</td>
              <td><input type="text" class="spo-inline-input w-full" data-pt-field="current_location" value="${spoAttrVal(r.current_location || "")}"${dis} /></td>
              <td><input type="date" class="spo-inline-input w-full" data-pt-field="expected_arrival_date" value="${spoAttrVal(r.expected_arrival_date || "")}"${dis} /></td>
              <td>
                <input type="text" class="spo-inline-input w-full" data-pt-field="supplier_name" value="${spoAttrVal(r.supplier_name || "")}" placeholder="Supplier"${dis} />
                <input type="text" class="spo-inline-input w-full" data-pt-field="po_number" value="${spoAttrVal(r.po_number || "")}" placeholder="PO"${dis} style="margin-top:4px;" />
              </td>
              <td>
                ${cancelled ? `<span class="muted small">Cancelled</span>` : `
                  <button type="button" class="btn btn-primary btn-sm" data-pt-save="${id}">Save</button>
                  ${String(r.status || "").toLowerCase() === "arrived" && !r.in_store_inventory
                    ? `<button type="button" class="btn btn-secondary btn-sm" data-pt-receive="${id}">Receive</button>`
                    : ""}
                  <button type="button" class="btn btn-secondary btn-sm" data-pt-del="${id}">Cancel</button>
                `}
              </td>
            </tr>`;
        }).join("")}
      </tbody>
    </table>
  `;
}

function readPtPartsRowPatch(rowEl) {
  if (!rowEl) return null;
  const read = (field) => {
    const el = rowEl.querySelector(`[data-pt-field="${field}"]`);
    if (!el) return null;
    return String(el.value || "").trim() || null;
  };
  const statusEl = rowEl.querySelector("select[data-pt-status]");
  const patch = {
    invoice_number: read("invoice_number"),
    current_location: read("current_location"),
    expected_arrival_date: read("expected_arrival_date"),
    supplier_name: read("supplier_name"),
    po_number: read("po_number"),
  };
  if (statusEl && !statusEl.disabled) patch.status = String(statusEl.value || "").trim();
  return patch;
}

async function loadPtPartsOrders() {
  ensurePartsTrackingDates();
  const start = (qs("ptPartsFrom")?.value || "").trim();
  const end = (qs("ptPartsTo")?.value || "").trim();
  const status = (qs("ptPartsStatus")?.value || "").trim();
  const q = new URLSearchParams();
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  if (status) q.set("status", status);
  setSkeleton("ptPartsList", 2);
  const data = await fetchJson(`${API}/api/stock/part-orders?${q.toString()}`);
  ptPartsCache = Array.isArray(data?.rows) ? data.rows : [];
  renderPtPartsTable(ptPartsCache);
  updatePartsTrackingKpis();
}

async function savePtPartsOrder() {
  ensurePartsTrackingDates();
  const msg = qs("ptPartsMsg");
  const body = {
    part_code: normalizeStockReportPartInput(qs("ptPartCode")?.value || ""),
    part_name: String(qs("ptPartName")?.value || "").trim(),
    qty: Number(qs("ptQty")?.value || 1),
    unit_cost: Number(qs("ptUnitCost")?.value || 0),
    invoice_number: String(qs("ptInvoice")?.value || "").trim() || null,
    current_location: String(qs("ptLocation")?.value || "").trim() || null,
    order_date: (qs("ptOrderDate")?.value || "").trim(),
    expected_arrival_date: (qs("ptEta")?.value || "").trim() || null,
    status: String(qs("ptStatus")?.value || "on_order").trim(),
    supplier_name: String(qs("ptSupplier")?.value || "").trim() || null,
    po_number: String(qs("ptPo")?.value || "").trim() || null,
    notes: String(qs("ptNotes")?.value || "").trim() || null,
  };
  if (!body.part_code && !body.part_name) {
    if (msg) msg.textContent = "Enter a part code or description.";
    return;
  }
  if (!body.order_date) {
    if (msg) msg.textContent = "Order date is required.";
    return;
  }
  if (msg) msg.textContent = "Saving…";
  const data = await fetchJson(`${API}/api/stock/part-orders`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  handleStoresPartOrderReceipt(data);
  if (msg) msg.textContent = "Part line saved.";
  clearPtPartsForm();
  await loadPtPartsOrders();
  loadStoresPartOrders().catch(() => {});
}

function ptOffStatusLabel(s) {
  const v = String(s || "").toLowerCase();
  const map = {
    sent_offsite: "Sent offsite",
    diagnosis: "Diagnosis",
    in_repair: "In repair",
    waiting_parts: "Waiting parts",
    ready_return: "Ready return",
    returned: "Returned",
  };
  return map[v] || v || "—";
}

function clearPtOffForm() {
  ["ptOffAsset", "ptOffAttachment", "ptOffLocation", "ptOffInvoice", "ptOffVendor", "ptOffNotes", "ptOffEta"].forEach((id) => {
    if (qs(id)) qs(id).value = "";
  });
  if (qs("ptOffEstCost")) qs("ptOffEstCost").value = "0";
  if (qs("ptOffActCost")) qs("ptOffActCost").value = "0";
  if (qs("ptOffStatus")) qs("ptOffStatus").value = "sent_offsite";
  if (qs("ptOffSent")) qs("ptOffSent").value = ptTodayYmd();
  if (qs("ptOffEditId")) qs("ptOffEditId").value = "";
  const msg = qs("ptOffMsg");
  if (msg) msg.textContent = "";
}

function renderPtOffsiteTable(rows) {
  const host = qs("ptOffList");
  if (!host) return;
  const statusFilter = String(qs("ptOffStatusFilter")?.value || "").trim().toLowerCase();
  let list = Array.isArray(rows) ? rows : [];
  if (statusFilter) list = list.filter((r) => String(r.repair_status || "").toLowerCase() === statusFilter);
  if (!list.length) {
    host.innerHTML = `<div class="muted small">No off-site repairs for this filter.</div>`;
    return;
  }
  host.innerHTML = `
    <table class="gridTable" style="min-width:1280px;">
      <thead>
        <tr>
          <th>Asset / attachment</th>
          <th>Status</th>
          <th>Vendor</th>
          <th>Invoice #</th>
          <th>Sent</th>
          <th>Location</th>
          <th>ETA on site</th>
          <th style="text-align:right;">Est. $</th>
          <th style="text-align:right;">Actual $</th>
          <th>Notes</th>
          <th></th>
        </tr>
      </thead>
      <tbody>
        ${list.map((r) => {
          const id = Number(r.id || 0);
          const overdue = ptIsOverdueEta(r.expected_return_date, String(r.repair_status || "").toLowerCase() === "returned");
          const attach = String(r.attachment_name || "").trim();
          const title = `<strong>${escapeHtml(String(r.asset_code || ""))}</strong>${attach ? `<br><small class="muted">${escapeHtml(attach)}</small>` : ""}`;
          const statusOpts = ["sent_offsite", "diagnosis", "in_repair", "waiting_parts", "ready_return", "returned"]
            .map((st) => `<option value="${st}"${String(r.repair_status || "").toLowerCase() === st ? " selected" : ""}>${ptOffStatusLabel(st)}</option>`)
            .join("");
          return `
            <tr data-pt-off-row="${id}" class="${overdue ? "pt-row-overdue" : ""}">
              <td>${title}<div class="muted small">${escapeHtml(String(r.asset_name || ""))}</div></td>
              <td><select data-pt-off-field="repair_status" class="w-full" style="min-width:120px;">${statusOpts}</select></td>
              <td><input type="text" class="spo-inline-input w-full" data-pt-off-field="vendor" value="${spoAttrVal(r.vendor || "")}" /></td>
              <td><input type="text" class="spo-inline-input w-full" data-pt-off-field="invoice_number" value="${spoAttrVal(r.invoice_number || "")}" /></td>
              <td><input type="date" class="spo-inline-input w-full" data-pt-off-field="sent_date" value="${spoAttrVal(r.sent_date || "")}" /></td>
              <td><input type="text" class="spo-inline-input w-full" data-pt-off-field="current_location" value="${spoAttrVal(r.current_location || "")}" /></td>
              <td><input type="date" class="spo-inline-input w-full" data-pt-off-field="expected_return_date" value="${spoAttrVal(r.expected_return_date || "")}" /></td>
              <td><input type="number" min="0" step="0.01" class="spo-inline-input w-full" data-pt-off-field="estimated_cost" value="${spoAttrVal(r.estimated_cost != null ? r.estimated_cost : "")}" /></td>
              <td><input type="number" min="0" step="0.01" class="spo-inline-input w-full" data-pt-off-field="actual_cost" value="${spoAttrVal(r.actual_cost != null ? r.actual_cost : "")}" /></td>
              <td>
                <input type="text" class="spo-inline-input w-full" data-pt-off-field="notes" value="${spoAttrVal(r.notes || "")}" />
                <input type="hidden" data-pt-off-field="attachment_name" value="${spoAttrVal(r.attachment_name || "")}" />
                <input type="hidden" data-pt-off-field="breakdown_id" value="${spoAttrVal(r.breakdown_id || "")}" />
              </td>
              <td><button type="button" class="btn btn-primary btn-sm" data-pt-off-save="${id}">Save</button></td>
            </tr>`;
        }).join("")}
      </tbody>
    </table>
  `;
}

async function loadPtOffsiteRepairs() {
  ensurePartsTrackingDates();
  const include = qs("ptOffIncludeReturned")?.checked ? "1" : "0";
  setSkeleton("ptOffList", 2);
  const data = await fetchJson(`${API}/api/breakdown-ops/offsite-repairs?include_closed=${include}`);
  ptOffsiteCache = Array.isArray(data?.rows) ? data.rows : [];
  renderPtOffsiteTable(ptOffsiteCache);
  updatePartsTrackingKpis();
}

async function savePtOffsiteRepair() {
  ensurePartsTrackingDates();
  const msg = qs("ptOffMsg");
  const asset_code = String(qs("ptOffAsset")?.value || "").trim();
  if (!asset_code) {
    if (msg) msg.textContent = "Asset code is required.";
    return;
  }
  const sent_date = (qs("ptOffSent")?.value || "").trim();
  if (!sent_date) {
    if (msg) msg.textContent = "Date sent is required.";
    return;
  }
  const body = {
    asset_code,
    attachment_name: String(qs("ptOffAttachment")?.value || "").trim() || null,
    repair_status: String(qs("ptOffStatus")?.value || "sent_offsite").trim(),
    sent_date,
    expected_return_date: (qs("ptOffEta")?.value || "").trim() || null,
    current_location: String(qs("ptOffLocation")?.value || "").trim() || null,
    invoice_number: String(qs("ptOffInvoice")?.value || "").trim() || null,
    vendor: String(qs("ptOffVendor")?.value || "").trim() || null,
    estimated_cost: Number(qs("ptOffEstCost")?.value || 0) || 0,
    actual_cost: Number(qs("ptOffActCost")?.value || 0) || 0,
    notes: String(qs("ptOffNotes")?.value || "").trim() || null,
  };
  if (msg) msg.textContent = "Saving…";
  await fetchJson(`${API}/api/breakdown-ops/offsite-repairs`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (msg) msg.textContent = "Off-site repair saved.";
  clearPtOffForm();
  await loadPtOffsiteRepairs();
}

function readPtOffRowPatch(rowEl) {
  if (!rowEl) return null;
  const read = (field) => {
    const el = rowEl.querySelector(`[data-pt-off-field="${field}"]`);
    if (!el) return null;
    return String(el.value || "").trim();
  };
  const estimated_cost = read("estimated_cost");
  const actual_cost = read("actual_cost");
  const breakdownRaw = read("breakdown_id");
  return {
    repair_status: read("repair_status") || "sent_offsite",
    vendor: read("vendor") || null,
    invoice_number: read("invoice_number") || null,
    sent_date: read("sent_date"),
    current_location: read("current_location") || null,
    expected_return_date: read("expected_return_date") || null,
    estimated_cost: estimated_cost === "" ? null : Number(estimated_cost),
    actual_cost: actual_cost === "" ? null : Number(actual_cost),
    notes: read("notes") || null,
    attachment_name: read("attachment_name") || null,
    breakdown_id: breakdownRaw ? Number(breakdownRaw) : null,
  };
}

async function savePtOffsiteRow(id) {
  const host = qs("ptOffList");
  const rowEl = host?.querySelector(`tr[data-pt-off-row="${Number(id)}"]`);
  const patch = readPtOffRowPatch(rowEl);
  if (!patch || !patch.sent_date) {
    alert("Sent date is required.");
    return;
  }
  await fetchJson(`${API}/api/breakdown-ops/offsite-repairs/${Number(id)}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(patch),
  });
  setStatus("Off-site repair updated.");
  await loadPtOffsiteRepairs();
}

async function loadPartsTrackingTab() {
  ensurePartsTrackingDates();
  await Promise.all([
    loadPtPartsOrders().catch((e) => setStatus("Parts tracking error: " + (e.message || e))),
    loadPtOffsiteRepairs().catch((e) => setStatus("Off-site tracking error: " + (e.message || e))),
  ]);
}

async function loadAuditLogs() {
  const module = (qs("auditModule")?.value || "").trim();
  const action = (qs("auditAction")?.value || "").trim();
  const entity_type = (qs("auditEntityType")?.value || "").trim();
  const entity_id = (qs("auditEntityId")?.value || "").trim();
  const username = (qs("auditUsername")?.value || "").trim();
  const site_code = (qs("auditSiteCode")?.value || "").trim();
  const source_app = (qs("auditSourceApp")?.value || "").trim();
  const start = (qs("auditStart")?.value || "").trim();
  const end = (qs("auditEnd")?.value || "").trim();
  const limit = Number(qs("auditLimit")?.value || 200);
  const list = qs("auditList");
  const detail = qs("auditDetail");
  if (!list) return;
  if (detail) detail.textContent = "";

  setStatus("Loading audit trail...");
  setSkeleton("auditList", 2);

  const q = new URLSearchParams();
  if (module) q.set("module", module);
  if (action) q.set("action", action);
  if (entity_type) q.set("entity_type", entity_type);
  if (entity_id) q.set("entity_id", entity_id);
  if (username) q.set("username", username);
  if (site_code) q.set("site_code", site_code);
  if (source_app) q.set("source_app", source_app);
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  if (Number.isFinite(limit) && limit > 0) q.set("limit", String(Math.trunc(limit)));

  const url = `${API}/api/audit/timeline${q.toString() ? `?${q.toString()}` : ""}`;
  const data = await fetchJson(url);
  const rows = Array.isArray(data.rows) ? data.rows : [];

  list.innerHTML = "";
  rows.forEach((r) => {
    const row = item(
      `<b>${r.created_at || "-"}</b> — ${r.module}.${r.action} ` +
      `<span class="pill blue">${r.role || "-"}</span> <span class="pill">${r.source_app || "-"}</span>` +
      `<br><small>${r.username || "-"} | site=${r.site_code || "-"} | ${r.entity_type || "-"}:${r.entity_id || "-"}</small>`
    );
    row.style.cursor = "pointer";
    row.addEventListener("click", async () => {
      if (!detail) return;
      detail.textContent = "Loading detail...";
      try {
        const d = await fetchJson(`${API}/api/audit/timeline/${Number(r.id || 0)}`);
        detail.textContent = JSON.stringify(d?.row || {}, null, 2);
      } catch (e) {
        detail.textContent = String(e.message || e);
      }
    });
    list.appendChild(row);
  });
  if (!rows.length) list.appendChild(item("<small>No audit records found.</small>"));

  const page = data?.pagination || {};
  setStatus(`Audit trail ready${page?.has_more ? " (more available)" : ""}.`);
}

function canManageLegalDocs() {
  const roles = getSessionRoles();
  return roles.includes("admin") || roles.includes("supervisor");
}

function canApproveRequests() {
  const roles = getSessionRoles();
  return roles.includes("admin") || roles.includes("supervisor");
}

function isTodayStamp(v) {
  const s = String(v || "").trim();
  if (!s) return false;
  const today = new Date().toISOString().slice(0, 10);
  return s.startsWith(today);
}

async function loadApprovalRequests() {
  const list = qs("approvalList");
  if (!list) return;

  const status = (qs("approvalStatus")?.value || "").trim();
  const module = (qs("approvalModule")?.value || "").trim();
  const action = (qs("approvalAction")?.value || "").trim();

  setStatus("Loading approvals...");
  setSkeleton("approvalList", 2);

  const q = new URLSearchParams();
  if (status) q.set("status", status);
  if (module) q.set("module", module);
  if (action) q.set("action", action);

  const data = await fetchJson(`${API}/api/approvals${q.toString() ? `?${q.toString()}` : ""}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  const approver = canApproveRequests();

  setText("approvalAllCount", rows.length);
  const pendingCount = rows.filter((r) => String(r.status || "").toLowerCase() === "pending").length;
  const approvedTodayCount = rows.filter(
    (r) => String(r.status || "").toLowerCase() === "approved" && isTodayStamp(r.approved_at)
  ).length;
  const rejectedTodayCount = rows.filter(
    (r) => String(r.status || "").toLowerCase() === "rejected" && isTodayStamp(r.rejected_at)
  ).length;
  setText("approvalPendingCount", pendingCount);
  setText("approvalApprovedTodayCount", approvedTodayCount);
  setText("approvalRejectedTodayCount", rejectedTodayCount);
  const approvalKpiStrip = qs("approvalKpiStrip");
  if (approvalKpiStrip) {
    const currentFilter = status;
    Array.from(approvalKpiStrip.querySelectorAll("[data-approval-kpi-filter]")).forEach((el) => {
      if (!(el instanceof HTMLElement)) return;
      const filter = String(el.getAttribute("data-approval-kpi-filter") || "");
      el.classList.toggle("pill-active", filter === currentFilter);
    });
  }

  list.innerHTML = "";
  rows.forEach((r) => {
    const st = String(r.status || "").toLowerCase();
    const statusPill =
      st === "approved"
        ? "<span class='pill blue'>approved</span>"
        : st === "rejected"
        ? "<span class='pill red'>rejected</span>"
        : "<span class='pill orange'>pending</span>";
    const payloadTxt = r.payload ? JSON.stringify(r.payload) : "{}";
    const actionBtns =
      approver && st === "pending"
        ? `<br><button data-approval-approve-id="${r.id}" style="margin-top:8px;">Approve</button><button data-approval-reject-id="${r.id}" style="margin-top:8px;">Reject</button>`
        : "";
    list.appendChild(
      item(
        `<b>#${r.id}</b> ${statusPill} — ${r.module || "-"} . ${r.action || "-"}` +
          `<br><small>Requested: ${r.requested_by || "-"} (${r.requested_role || "-"}) @ ${r.created_at || "-"}</small>` +
          `<br><small>Entity: ${r.entity_type || "-"}:${r.entity_id || "-"}</small>` +
          `<br><small>${payloadTxt}</small>` +
          actionBtns
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No approval requests found.</small>"));
  setStatus("Approvals ready.");
}

async function decideApprovalRequest(id, decision) {
  if (!canApproveRequests()) {
    alert("Only admin/supervisor can approve or reject requests.");
    return;
  }
  const reqId = Number(id || 0);
  const d = String(decision || "").trim().toLowerCase();
  if (!reqId || !["approve", "reject"].includes(d)) return;

  const note = (qs("approvalDecisionNote")?.value || "").trim();
  setStatus(`${d === "approve" ? "Approving" : "Rejecting"} request #${reqId}...`);
  try {
    const res = await fetchJson(`${API}/api/approvals/${reqId}/${d}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ note: note || undefined }),
    });
    setText("approvalResult", JSON.stringify(res, null, 2));
    await Promise.all([loadApprovalRequests().catch(() => {}), loadDashboard().catch(() => {})]);
    setStatus(`Approval request #${reqId} ${d}d.`);
  } catch (e) {
    setText("approvalResult", String(e.message || e));
    setStatus("Approval decision failed.");
  }
}

function legalAllowedTransitions(currentStatus) {
  const s = String(currentStatus || "draft").toLowerCase();
  const transitions = {
    draft: ["pending_approval", "superseded"],
    rejected: ["pending_approval", "superseded"],
    pending_approval: ["approved", "rejected", "superseded"],
    approved: ["superseded"],
    superseded: [],
  };
  return transitions[s] || [];
}

async function loadLegalDepartments() {
  const depEl = qs("legalDepartment");
  const filterEl = qs("legalFilterDepartment");
  if (!depEl && !filterEl) return;

  const data = await fetchJson(`${API}/api/legal/departments`);
  const deps = Array.isArray(data.departments) ? data.departments : [];

  if (depEl) {
    depEl.innerHTML = deps.map((d) => `<option value="${d}">${d}</option>`).join("");
  }
  if (filterEl) {
    const options = [`<option value="">All departments</option>`]
      .concat(deps.map((d) => `<option value="${d}">${d}</option>`));
    filterEl.innerHTML = options.join("");
  }
}

async function loadLegalDocs() {
  const dep = (qs("legalFilterDepartment")?.value || "").trim();
  const status = (qs("legalFilterStatus")?.value || "").trim();
  const qText = (qs("legalSearch")?.value || "").trim();
  const includeInactive = qs("legalIncludeInactive")?.checked ? "1" : "0";
  const list = qs("legalList");
  if (!list) return;

  setStatus("Loading legal library...");
  setSkeleton("legalList", 2);

  const q = new URLSearchParams();
  if (dep) q.set("department", dep);
  if (status) q.set("status", status);
  if (qText) q.set("q", qText);
  q.set("include_inactive", includeInactive);

  const data = await fetchJson(`${API}/api/legal?${q.toString()}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];

  list.innerHTML = "";
  rows.forEach((r) => {
    const allowed = legalAllowedTransitions(r.status);
    const statusPill =
      r.status === "approved"
        ? "<span class='pill blue'>approved</span>"
        : r.status === "pending_approval"
        ? "<span class='pill orange'>pending</span>"
        : r.status === "rejected"
        ? "<span class='pill red'>rejected</span>"
        : r.status === "superseded"
        ? "<span class='pill orange'>superseded</span>"
        : "<span class='pill'>draft</span>";

    const archiveBtn = canManageLegalDocs()
      ? `<button data-legal-archive-id="${r.id}" data-legal-active="${r.active ? 0 : 1}" style="margin-top:8px;">${r.active ? "Archive" : "Reactivate"}</button>`
      : "";
    const statusLabel = {
      pending_approval: "Submit",
      approved: "Approve",
      rejected: "Reject",
      superseded: "Supersede",
    };
    const statusBtns = canManageLegalDocs() && Number(r.active) === 1
      ? allowed
          .map(
            (next) =>
              `<button data-legal-status-id="${r.id}" data-legal-status="${next}" style="margin-top:8px;">${statusLabel[next] || next}</button>`
          )
          .join("")
      : "";
    list.appendChild(
      item(
        `<b>${r.department || "-"}</b> — ${r.title || "-"} ${statusPill} ${r.active ? "" : "<span class='pill red'>ARCHIVED</span>"}` +
        `<br><small>Type: ${r.doc_type || "-"} | Version: ${r.version || "-"} | Owner: ${r.owner || "-"} | Uploaded: ${r.created_at || "-"}</small>` +
        `<br><small>Effective: ${r.effective_date || "-"} | Expiry: ${r.expiry_date || "-"} | Approved by: ${r.approved_by || "-"} ${r.approved_at ? `@ ${r.approved_at}` : ""}</small>` +
        (r.approval_note ? `<br><small>Note: ${r.approval_note}</small>` : "") +
        `<br><button data-legal-download-id="${r.id}" style="margin-top:8px;">Download</button>` +
        `<button data-legal-actions-id="${r.id}" style="margin-top:8px;">History</button>` +
        `${statusBtns}${archiveBtn}`
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No documents found.</small>"));
  setStatus("Legal library ready.");
}

async function loadLegalExpiry() {
  const days = Number(qs("legalExpiryDays")?.value || 90);
  const dep = (qs("legalFilterDepartment")?.value || "").trim();
  const status = (qs("legalFilterStatus")?.value || "").trim() || "approved";

  const q = new URLSearchParams();
  if (Number.isFinite(days) && days > 0) q.set("days", String(Math.trunc(days)));
  if (dep) q.set("department", dep);
  if (status) q.set("status", status);

  const data = await fetchJson(`${API}/api/legal/expiry?${q.toString()}`);
  const s = data?.summary || {};
  const setText = (id, v) => {
    const el = qs(id);
    if (el) el.textContent = String(v ?? 0);
  };
  setText("legalExpiredCount", Number(s.expired || 0));
  setText("legalDue30Count", Number(s.due_30 || 0));
  setText("legalDue60Count", Number(s.due_60 || 0));
  setText("legalDue90Count", Number(s.due_90 || 0));
}

function openLegalCompliancePdf(download = false) {
  const days = Number(qs("legalExpiryDays")?.value || 90);
  const dep = (qs("legalFilterDepartment")?.value || "").trim();
  const status = (qs("legalFilterStatus")?.value || "").trim() || "approved";
  const q = new URLSearchParams();
  if (Number.isFinite(days) && days > 0) q.set("days", String(Math.trunc(days)));
  if (dep) q.set("department", dep);
  if (status) q.set("status", status);
  if (download) q.set("download", "1");
  window.open(`${API}/api/reports/legal-compliance.pdf?${q.toString()}`, "_blank");
}

async function setLegalStatus(id, status) {
  if (!canManageLegalDocs()) {
    alert("Only admin/supervisor can change legal document status.");
    return;
  }
  const docId = Number(id || 0);
  if (!docId) return;
  const note = (qs("legalActionNote")?.value || "").trim() || undefined;
  const supEl = qs("legalSupersedesId");
  const supHint = qs("legalSupersedeHint");

  const payload = { status, note };
  if (status === "superseded") {
    if (supEl) {
      supEl.disabled = false;
      if (supHint) supHint.style.display = "";
      const v = (supEl.value || "").trim();
      if (!v) {
        setStatus("Enter 'Supersedes Doc ID' and click Supersede again.");
        supEl.focus();
        return;
      }
      payload.supersedes_document_id = Number(v);
    }
  } else if (supEl) {
    supEl.value = "";
    supEl.disabled = true;
    if (supHint) supHint.style.display = "none";
  }

  const out = qs("legalActionResult");
  setStatus(`Applying status '${status}'...`);
  try {
    const res = await fetchJson(`${API}/api/legal/${docId}/status`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (out) out.textContent = JSON.stringify(res, null, 2);
    const noteEl = qs("legalActionNote");
    if (noteEl) noteEl.value = "";
    if (supEl) {
      supEl.value = "";
      supEl.disabled = true;
    }
    if (supHint) supHint.style.display = "none";
    await loadLegalDocs().catch(() => {});
    setStatus("Legal status updated.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Legal status update failed.");
  }
}

async function showLegalActions(id) {
  const docId = Number(id || 0);
  if (!docId) return;
  const out = qs("legalActionResult");
  setStatus("Loading legal action history...");
  try {
    const data = await fetchJson(`${API}/api/legal/${docId}/actions`);
    if (out) out.textContent = JSON.stringify(data.actions || [], null, 2);
    setStatus("Legal action history loaded.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Legal action history failed.");
  }
}

async function uploadLegalDoc() {
  if (!canManageLegalDocs()) {
    alert("Only admin/supervisor can upload legal documents.");
    return;
  }

  const fileEl = qs("legalFile");
  const file = fileEl?.files?.[0];
  if (!file) return alert("Choose a file first.");

  const fd = new FormData();
  let department = (qs("legalDepartment")?.value || "").trim();
  if (!department) {
    try {
      await loadLegalDepartments();
      department = (qs("legalDepartment")?.value || "").trim();
    } catch (_) {
      // Keep graceful fallback below.
    }
  }
  if (!department) {
    const depEl = qs("legalDepartment");
    const firstOpt = depEl?.querySelector("option");
    department = String(firstOpt?.value || firstOpt?.textContent || "").trim();
  }
  if (!department) {
    return alert("Select a department before upload.");
  }

  let title = (qs("legalTitle")?.value || "").trim();
  if (!title) {
    title = String(file.name || "Untitled")
      .replace(/\.[^.]+$/, "")
      .trim();
    const titleEl = qs("legalTitle");
    if (titleEl) titleEl.value = title;
  }
  if (!title) {
    return alert("Enter a document title before upload.");
  }

  fd.append("file", file);
  fd.append("department", department);
  fd.append("title", title);
  fd.append("doc_type", (qs("legalDocType")?.value || "").trim());
  fd.append("version", (qs("legalVersion")?.value || "").trim());
  fd.append("owner", (qs("legalOwner")?.value || "").trim());
  fd.append("effective_date", (qs("legalEffectiveDate")?.value || "").trim());
  fd.append("expiry_date", (qs("legalExpiryDate")?.value || "").trim());

  const resultEl = qs("legalUploadResult");
  setStatus("Uploading legal document...");
  try {
    const res = await fetch(`${API}/api/legal/upload`, {
      method: "POST",
      headers: authHeaders(),
      body: fd,
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Upload failed");
    if (resultEl) resultEl.textContent = JSON.stringify(data, null, 2);
    await loadLegalDocs().catch(() => {});
    setStatus("Legal document uploaded.");
  } catch (e) {
    if (resultEl) resultEl.textContent = String(e.message || e);
    setStatus("Legal upload failed.");
  }
}

function downloadLegalDoc(id) {
  const docId = Number(id || 0);
  if (!docId) return;
  window.open(`${API}/api/legal/${docId}/download`, "_blank");
}

async function archiveLegalDoc(id, active) {
  if (!canManageLegalDocs()) {
    alert("Only admin/supervisor can archive/reactivate documents.");
    return;
  }
  const docId = Number(id || 0);
  if (!docId) return;

  setStatus(active ? "Reactivating document..." : "Archiving document...");
  try {
    const res = await fetchJson(`${API}/api/legal/${docId}/archive`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ active }),
    });
    await loadLegalDocs().catch(() => {});
    setStatus(active ? "Document reactivated." : "Document archived.");
    return res;
  } catch (e) {
    setStatus("Document archive action failed.");
    alert(e.message || e);
  }
}

/* =========================
   TABS
========================= */

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
  if (k === "workshop") {
    openWorkshopLibraryOnTabActivate();
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
    "IronMind analyses fleet data here — use this Help button only for how to use IRONLOG screens.",
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

async function doUpload() {
  const endpointEl = qs("uploadEndpoint");
  const fileEl = qs("uploadFile");
  const resultEl = qs("uploadResult");
  if (!endpointEl || !fileEl || !resultEl) return;

  const endpoint = endpointEl.value;
  const file = fileEl.files[0];
  if (!file) return alert("Choose a CSV file first.");

  const fd = new FormData();
  fd.append("file", file);

  setStatus("Uploading CSV...");
  try {
    const res = await fetchJson(`${API}${endpoint}`, { method: "POST", body: fd });
    resultEl.textContent = JSON.stringify(res, null, 2);
    setStatus("Upload complete.");
    await loadDashboard().catch(() => {});
  } catch (e) {
    resultEl.textContent = String(e.message || e);
    setStatus("Upload failed.");
  }
}

async function importFamsFuelFile() {
  const fileEl = qs("fuelFamsFile");
  const modeEl = qs("fuelFamsConflictMode");
  const resultEl = qs("fuelFamsResult");
  if (!fileEl || !resultEl) return;

  const file = fileEl.files[0];
  if (!file) return alert("Choose a FAMS CSV file first.");

  const fd = new FormData();
  fd.append("file", file);
  const conflictMode = String(modeEl?.value || "skip").trim().toLowerCase();
  const mode = ["skip", "overwrite"].includes(conflictMode) ? conflictMode : "skip";

  setStatus("Importing FAMS fuel file...");
  resultEl.textContent = "";
  try {
    const res = await fetchJson(`${API}/api/upload/fuel?on_conflict=${encodeURIComponent(mode)}`, { method: "POST", body: fd });
    resultEl.textContent = JSON.stringify(res, null, 2);
    setStatus("FAMS fuel import complete.");
    await loadDashboard().catch(() => {});
    loadFamsFuelStatus().catch(() => {});
  } catch (e) {
    resultEl.textContent = String(e.message || e);
    setStatus("FAMS fuel import failed.");
  }
}

function renderFamsFuelStatus(data) {
  const status = String(data?.status || (data?.enabled ? "idle" : "disabled"));
  const labelMap = {
    connected: "Connected",
    error: "Error",
    disabled: "Disabled",
    idle: "Idle",
    syncing: "Syncing…",
  };
  setText("fuelFamsStatusLabel", labelMap[status] || status);
  setText("fuelFamsLastSuccess", data?.last_success_at ? String(data.last_success_at).replace("T", " ").slice(0, 19) : "—");
  setText("fuelFamsLastAttempt", data?.last_attempt_at ? String(data.last_attempt_at).replace("T", " ").slice(0, 19) : "—");
  setText("fuelFamsReceived", data?.last_received != null ? String(data.last_received) : "—");
  setText("fuelFamsImported", data?.last_imported != null ? String(data.last_imported) : "—");
  setText("fuelFamsSkipped", data?.last_skipped != null ? String(data.last_skipped) : "—");
  const unmatched = data?.unmatched_open != null ? data.unmatched_open : data?.last_unmatched;
  setText("fuelFamsUnmatched", unmatched != null ? String(unmatched) : "—");
  const msg = qs("fuelFamsStatusMsg");
  if (msg) {
    if (!data?.enabled) {
      msg.textContent = "Auto sync is off. Set FAMS_ENABLED=true and FAMS_AUTH in api/.env, then restart the API.";
    } else if (!data?.configured) {
      msg.textContent = "FAMS_AUTH is missing on the server.";
    } else if (data?.last_error) {
      msg.textContent = `Last error: ${data.last_error}`;
    } else if (data?.last_range_start && data?.last_range_end) {
      msg.textContent = `Last range ${data.last_range_start} → ${data.last_range_end}. Fuel only — hours stay on QR pre-start.`;
    } else {
      msg.textContent = data?.note || "";
    }
  }
}

async function loadFamsFuelStatus() {
  try {
    const data = await fetchJson(`${API}/api/dashboard/fuel/fams/status`);
    renderFamsFuelStatus(data);
    return data;
  } catch (e) {
    setText("fuelFamsStatusLabel", "Error");
    const msg = qs("fuelFamsStatusMsg");
    if (msg) msg.textContent = e.message || String(e);
    throw e;
  }
}

async function syncFamsFuelNow() {
  const resultEl = qs("fuelFamsResult");
  setStatus("Syncing FAMS fuel…");
  if (resultEl) resultEl.textContent = "Syncing…";
  try {
    const res = await fetchJson(`${API}/api/dashboard/fuel/fams/sync`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({ force: true }),
    });
    if (resultEl) resultEl.textContent = JSON.stringify(res, null, 2);
    if (res?.status) renderFamsFuelStatus(res.status);
    else await loadFamsFuelStatus().catch(() => {});
    if (res?.ok) {
      setStatus(
        `FAMS sync done: ${Number(res.imported || 0)} imported, ${Number(res.skipped || 0)} skipped, ${Number(res.unmatched || 0)} unmatched.`
      );
      await loadDashboard().catch(() => {});
    } else {
      setStatus(`FAMS sync failed: ${res?.error || res?.reason || "unknown"}`);
    }
  } catch (e) {
    if (resultEl) resultEl.textContent = String(e.message || e);
    setStatus("FAMS sync failed.");
    loadFamsFuelStatus().catch(() => {});
  }
}

async function repairFuelMeterChain() {
  const resultEl = qs("fuelFamsResult");
  const assetCode = (qs("fuelRepairAssetCode")?.value || "").trim();
  if (resultEl) resultEl.textContent = "";
  setStatus("Repairing meter chain...");
  try {
    const body = assetCode ? { asset_code: assetCode } : {};
    const res = await fetchJson(`${API}/api/dashboard/fuel/repair-meter-chain`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify(body),
    });
    if (resultEl) resultEl.textContent = JSON.stringify(res, null, 2);
    setStatus(`Meter chain repair complete. Rows repaired: ${Number(res?.repaired_rows || 0)}`);
    await loadDashboard().catch(() => {});
  } catch (e) {
    if (resultEl) resultEl.textContent = String(e.message || e);
    setStatus("Meter chain repair failed.");
  }
}

async function clearFuelFromDate() {
  const resultEl = qs("fuelFamsResult");
  const fromDate = (qs("fuelClearFromDate")?.value || "").trim();
  const assetCode = (qs("fuelClearAssetCode")?.value || "").trim();
  const clearDailyHours = Boolean(qs("fuelClearDailyHours")?.checked);
  if (!fromDate) return alert("Select a from date first.");

  if (resultEl) resultEl.textContent = "";
  setStatus("Checking clear impact...");
  try {
    const preview = await previewFuelClearImpact(fromDate, assetCode, clearDailyHours);
    const ok = confirm(
      `About to clear fuel from ${fromDate}` +
      `${assetCode ? ` for ${assetCode}` : " for all assets"}.\n` +
      `Fuel logs to delete: ${Number(preview?.deleted_logs || 0)}\n` +
      `Affected days: ${Number(preview?.affected_days || 0)}\n` +
      (clearDailyHours ? "Daily opening/closing/run meters on affected days will also be cleared.\n" : "") +
      "\nContinue?"
    );
    if (!ok) {
      setStatus("Fuel clear cancelled.");
      return;
    }

    setStatus("Clearing fuel data...");
    const res = await fetchJson(`${API}/api/dashboard/fuel/clear-from-date`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({
        from_date: fromDate,
        ...(assetCode ? { asset_code: assetCode } : {}),
        clear_daily_hours: clearDailyHours,
      }),
    });
    if (resultEl) resultEl.textContent = JSON.stringify(res, null, 2);
    setStatus(`Fuel clear complete. Deleted logs: ${Number(res?.deleted_logs || 0)}`);
    await Promise.all([loadDashboard().catch(() => {}), loadFuelBenchmark().catch(() => {})]);
  } catch (e) {
    if (resultEl) resultEl.textContent = String(e.message || e);
    setStatus("Fuel clear failed. Check API version and permissions.");
  }
}

async function previewFuelClearImpact(fromDate, assetCode, clearDailyHours) {
  const result = await fetchJson(`${API}/api/dashboard/fuel/clear-from-date/preview`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({
      from_date: fromDate,
      ...(assetCode ? { asset_code: assetCode } : {}),
      clear_daily_hours: clearDailyHours,
    }),
  });
  const previewEl = qs("fuelClearPreviewResult");
  if (previewEl) previewEl.textContent = JSON.stringify(result, null, 2);
  return result;
}

async function runFuelClearPreview() {
  const fromDate = (qs("fuelClearFromDate")?.value || "").trim();
  const assetCode = (qs("fuelClearAssetCode")?.value || "").trim();
  const clearDailyHours = Boolean(qs("fuelClearDailyHours")?.checked);
  if (!fromDate) return alert("Select a from date first.");
  setStatus("Loading clear preview...");
  const preview = await previewFuelClearImpact(fromDate, assetCode, clearDailyHours);
  setStatus(
    `Preview ready. Logs: ${Number(preview?.deleted_logs || 0)}, Days: ${Number(preview?.affected_days || 0)}`
  );
}

async function editFuelMachineHours(logId, openValue, closeValue) {
  const id = Number(logId || 0);
  if (!Number.isInteger(id) || id <= 0) return;
  const openDefault = Number.isFinite(Number(openValue)) ? String(Number(openValue)) : "";
  const closeDefault = Number.isFinite(Number(closeValue)) ? String(Number(closeValue)) : "";
  const openRaw = prompt("Enter opening meter value", openDefault);
  if (openRaw == null) return;
  const closeRaw = prompt("Enter closing meter value", closeDefault);
  if (closeRaw == null) return;

  const opening = Number(String(openRaw).trim());
  const closing = Number(String(closeRaw).trim());
  if (!Number.isFinite(opening) || opening < 0) return alert("Opening meter must be >= 0.");
  if (!Number.isFinite(closing) || closing < 0) return alert("Closing meter must be >= 0.");
  if (closing < opening) return alert("Closing meter must be greater than or equal to opening meter.");

  await fetchJson(`${API}/api/dashboard/fuel/machine-hours`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({
      fuel_log_id: id,
      opening_meter: opening,
      closing_meter: closing,
    }),
  });
}

async function saveFuelMachineHoursInline(saveBtn) {
  const id = Number(saveBtn?.getAttribute("data-fuel-save") || 0);
  if (!Number.isInteger(id) || id <= 0) return;
  const row = saveBtn.closest("tr");
  if (!row) return;
  const openEl = row.querySelector('input[data-fuel-open-input="1"]');
  const closeEl = row.querySelector('input[data-fuel-close-input="1"]');
  const opening = Number(String(openEl?.value || "").trim());
  const closing = Number(String(closeEl?.value || "").trim());

  if (!Number.isFinite(opening) || opening < 0) return alert("Opening hours must be >= 0.");
  if (!Number.isFinite(closing) || closing < 0) return alert("Closing hours must be >= 0.");
  if (closing < opening) return alert("Closing hours must be greater than or equal to opening hours.");

  await fetchJson(`${API}/api/dashboard/fuel/machine-hours`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({
      fuel_log_id: id,
      opening_meter: opening,
      closing_meter: closing,
    }),
  });
}

function downloadStoresCsvTemplate() {
  const today = new Date().toISOString().slice(0, 10);
  const lines = [
    "part_code,quantity,allocation_date,asset_code,work_order_id,issued_by,notes",
    `FLT-001,2,${today},A300AM,,Storeman A,Planned PM kit`,
    `BLT-009,1,${today},,41,Storeman B,Issued against WO 41`
  ];
  const csv = lines.join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "stores_alloc_template.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  setStatus("Stores CSV template downloaded.");
}

function downloadFuelBaselineCsvTemplate() {
  const lines = [
    "asset_code,baseline_fuel_l_per_hour",
    "A300AM,7.25",
    "A301AM,7.10",
    "E500AM,18.50"
  ];
  const csv = lines.join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "fuel_baseline_template.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  setStatus("Fuel baseline CSV template downloaded.");
}

function downloadFuelCsvTemplate() {
  const today = new Date().toISOString().slice(0, 10);
  const lines = [
    "asset_code,log_date,liters,source,meter_unit,meter_run_value,hours_run",
    `A300AM,${today},180,bowser,hours,10.0,10.0`,
    `A301AM,${today},220,bowser,hours,9.5,9.5`,
    `LDV01,${today},60,pump_1,km,480,`
  ];
  const csv = lines.join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "fuel_import_template.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  setStatus("Fuel CSV template downloaded.");
}

function saveCsvFile(csv, filename) {
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function csvCell(v) {
  const s = String(v == null ? "" : v);
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

async function downloadDailyHoursCsvTemplate() {
  const date = qs("date")?.value || todayLocalYmd();
  const header = "asset_code,asset_name,work_date,scheduled_hours,opening_hours,closing_hours,hours_run,is_used,operator,notes";

  setStatus("Building hours CSV template...");
  let assets = [];
  try {
    assets = await fetchJson(`${API}/api/assets?include_archived=0`);
  } catch {
    assets = [];
  }

  const active = (Array.isArray(assets) ? assets : [])
    .filter((a) => a.active !== 0 && a.active !== false)
    .sort((a, b) => String(a.asset_code || "").localeCompare(String(b.asset_code || "")));

  let lines;
  if (active.length) {
    lines = [header, ...active.map((a) => {
      const standby = a.is_standby ? 1 : 0;
      const sched = standby ? 0 : 10;
      const used = standby ? 0 : 1;
      return [
        csvCell(a.asset_code),
        csvCell(a.asset_name),
        date,
        sched,
        "",
        "",
        "",
        used,
        "",
        "",
      ].join(",");
    })];
  } else {
    // Fallback sample if the asset list could not be loaded.
    lines = [
      header,
      `A300AM,Excavator 300,${date},10,4500.0,4510.0,10.0,1,J Smith,`,
      `A301AM,Excavator 301,${date},10,3200.5,,9.5,1,,`,
      `A302AM,Dozer 302,${date},0,,,0,0,,Standby`,
    ];
  }

  saveCsvFile(lines.join("\n"), `daily_hours_template_${date}.csv`);
  setStatus(active.length
    ? `Daily hours template downloaded (${active.length} asset(s), date ${date}).`
    : "Daily hours CSV template downloaded (sample rows).");
}

async function uploadDailyHoursCsv(file) {
  if (!file) return;
  const fd = new FormData();
  fd.append("file", file);

  setStatus("Uploading hours CSV...");
  try {
    const res = await fetchJson(`${API}/api/upload/hours`, { method: "POST", body: fd });
    const imported = Number(res?.imported || 0);
    const synced = Number(res?.synced_asset_hours || 0);
    setStatus(`Hours CSV uploaded — ${imported} row(s) processed, ${synced} asset(s) synced.`);
    await loadDailyInput().catch(() => {});
  } catch (e) {
    setStatus("Hours CSV upload failed: " + (e.message || e));
    alert("Hours CSV upload failed: " + (e.message || e));
  }
}

async function downloadDailyMatrixCsvTemplate() {
  const date = qs("date")?.value || todayLocalYmd();
  setStatus("Building meter-matrix template...");
  let assets = [];
  try {
    assets = await fetchJson(`${API}/api/assets?include_archived=0`);
  } catch {
    assets = [];
  }
  const codes = (Array.isArray(assets) ? assets : [])
    .filter((a) => a.active !== 0 && a.active !== false)
    .map((a) => String(a.asset_code || "").trim())
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));

  const cols = codes.length ? codes : ["A300AM", "A301AM", "LDV01"];
  const header = ["date", ...cols].map(csvCell).join(",");
  // Two example date rows (blank cells to fill in with cumulative meter readings).
  const prev = prevDateStr(date);
  const lines = [
    header,
    [prev, ...cols.map(() => "")].join(","),
    [date, ...cols.map(() => "")].join(","),
  ];
  saveCsvFile(lines.join("\n"), `meter_matrix_template_${date}.csv`);
  setStatus(codes.length
    ? `Meter-matrix template downloaded (${codes.length} asset columns).`
    : "Meter-matrix template downloaded (sample columns).");
}

async function uploadDailyMatrixCsv(file) {
  if (!file) return;
  const scheduledRaw = Number(qs("bulkSched")?.value);
  const scheduled = Number.isFinite(scheduledRaw) && scheduledRaw > 0 && scheduledRaw <= 24 ? scheduledRaw : 10;
  const fd = new FormData();
  fd.append("file", file);

  setStatus("Uploading meter matrix...");
  try {
    const res = await fetchJson(`${API}/api/upload/hours-matrix?scheduled=${encodeURIComponent(scheduled)}`, {
      method: "POST",
      body: fd,
    });
    const rows = Number(res?.rows_written || 0);
    const assetsMatched = Number(res?.assets_matched || 0);
    const dates = Number(res?.dates || 0);
    const resets = Number(res?.meter_resets || 0);
    const unknown = Array.isArray(res?.unknown_codes) ? res.unknown_codes : [];
    let msg = `Meter matrix imported — ${rows} day-row(s) across ${assetsMatched} asset(s) over ${dates} date(s).`;
    if (resets) msg += ` ${resets} meter reset(s) handled.`;
    if (unknown.length) msg += ` Skipped unknown codes: ${unknown.join(", ")}.`;
    setStatus(msg);
    if (unknown.length) {
      alert(`Import complete, but these header codes were not found and were skipped:\n\n${unknown.join(", ")}`);
    }
    await loadDailyInput().catch(() => {});
  } catch (e) {
    setStatus("Meter matrix upload failed: " + (e.message || e));
    alert("Meter matrix upload failed: " + (e.message || e));
  }
}

/* =========================
   REPORTS
========================= */

function getLast7Range(endDate) {
  const end = new Date(endDate + "T00:00:00");
  const start = new Date(end);
  start.setDate(start.getDate() - 6);
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

function getLastNDaysRange(endDate, days) {
  const end = new Date(`${endDate}T00:00:00`);
  const span = Math.max(1, Number(days || 30));
  const start = new Date(end);
  start.setDate(start.getDate() - (span - 1));
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

function openDailyXlsx() {
  const date = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const scheduled = qs("scheduled")?.value || 10;
  const ts = Date.now();
  window.open(`${API}/api/reports/daily.xlsx?date=${date}&scheduled=${scheduled}&_ts=${ts}`, "_blank");
}

/** GM weekly pack: Maintenance & Engineering KPIs (same date field as daily / weekly PDF). */
function openGmWeeklyXlsx() {
  const end = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const scheduled = qs("scheduled")?.value || 10;
  window.open(
    `${API}/api/reports/gm-weekly.xlsx?end=${encodeURIComponent(end)}&forecast_days=30&scheduled=${scheduled}`,
    "_blank",
  );
}

function downloadCostMonthlyXlsx() {
  const month = (qs("costMonth")?.value || "").trim();
  if (!month) {
    alert("Select a month first.");
    return;
  }
  window.open(`${API}/api/reports/cost-monthly.xlsx?month=${encodeURIComponent(month)}`, "_blank");
}

function monthlyFleetCostPdfUrl(download = false) {
  const month = (qs("costMonth")?.value || "").trim();
  if (!month) return null;
  const scheduled = qs("scheduled")?.value || 10;
  const q = new URLSearchParams({ month, scheduled: String(scheduled) });
  if (download) q.set("download", "1");
  return `${API}/api/reports/monthly.pdf?${q.toString()}`;
}

function openMonthlyFleetCostPdf() {
  const url = monthlyFleetCostPdfUrl(false);
  if (!url) return alert("Select a cost month first.");
  window.open(url, "_blank");
}

function downloadMonthlyFleetCostPdf() {
  const url = monthlyFleetCostPdfUrl(true);
  if (!url) return alert("Select a cost month first.");
  window.open(url, "_blank");
}

function downloadMaintenanceCostByEquipmentXlsx() {
  const month = (qs("costMonth")?.value || "").trim();
  const start = (qs("maintCostStart")?.value || "").trim();
  const end = (qs("maintCostEnd")?.value || "").trim();
  if (!month && (!start || !end)) {
    alert("Select a month or a start/end range first.");
    return;
  }
  const q = month
    ? `month=${encodeURIComponent(month)}`
    : `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  window.open(`${API}/api/reports/maintenance-cost-by-equipment.xlsx?${q}`, "_blank");
}

function downloadMtdOpeningHoursXlsx() {
  const month = (qs("mtdOpeningMonth")?.value || "").trim();
  const q = month ? `?month=${encodeURIComponent(month)}` : "";
  window.open(`${API}/api/reports/mtd-opening-hours.xlsx${q}`, "_blank");
}

function openMaintenanceCostByEquipmentPdf(download = false) {
  const month = (qs("costMonth")?.value || "").trim();
  const start = (qs("maintCostStart")?.value || "").trim();
  const end = (qs("maintCostEnd")?.value || "").trim();
  if (!month && (!start || !end)) {
    alert("Select a month or a start/end range first.");
    return;
  }
  const qBase = month
    ? `month=${encodeURIComponent(month)}`
    : `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  const q = `${qBase}${download ? "&download=1" : ""}`;
  window.open(`${API}/api/reports/maintenance-cost-by-equipment.pdf?${q}`, "_blank");
}

function downloadMaintenanceExecutivePptx() {
  const month = (qs("costMonth")?.value || "").trim();
  const start = (qs("maintCostStart")?.value || "").trim();
  const end = (qs("maintCostEnd")?.value || "").trim();
  if (!month && (!start || !end)) {
    alert("Select a month or a start/end range first.");
    return;
  }
  const qCore = month
    ? `month=${encodeURIComponent(month)}`
    : `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  const site = encodeURIComponent(getSessionSite());
  const q = `${qCore}&site_code=${site}`;
  window.open(`${API}/api/reports/maintenance-exec.pptx?${q}`, "_blank");
}

function downloadGMUpcomingCostsPptx() {
  const month = (qs("costMonth")?.value || "").trim();
  const start = (qs("maintCostStart")?.value || "").trim();
  const end = (qs("maintCostEnd")?.value || "").trim();
  if (!month && (!start || !end)) {
    alert("Select a month or a start/end range first.");
    return;
  }
  const qCore = month
    ? `month=${encodeURIComponent(month)}`
    : `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  const site = encodeURIComponent(getSessionSite());
  const q = `${qCore}&site_code=${site}`;
  window.open(`${API}/api/reports/gm-upcoming-costs.pptx?${q}`, "_blank");
}

function downloadGMBudgetMeetingDocx() {
  const month = (qs("costMonth")?.value || qs("plantHireBudgetMonth")?.value || "").trim();
  if (!month) {
    alert("Select a month on Reports (cost month) or Assets → Plant Hire budget month.");
    return;
  }
  const site = encodeURIComponent(getSessionSite() || "main");
  const ts = Date.now();
  window.open(
    `${API}/api/reports/gm-budget-meeting.docx?month=${encodeURIComponent(month)}&site_code=${site}&_ts=${ts}`,
    "_blank",
  );
  setStatus("Budget meeting Word export started.");
}

async function saveRainDay() {
  const rainDate = (qs("rainDayDate")?.value || "").trim();
  if (!rainDate) return alert("Pick a rain day first.");
  setStatus("Saving rain day...");
  try {
    const res = await fetchJson(`${API}/api/reports/rain-days`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ date: rainDate }),
    });
    setText("rainDaysResult", JSON.stringify(res, null, 2));
    setStatus("Rain day saved.");
    loadRainDays({ silentNoRange: true }).catch(() => {});
  } catch (e) {
    setText("rainDaysResult", String(e.message || e));
    setStatus("Save rain day failed.");
  }
}

async function removeRainDay() {
  const rainDate = (qs("rainDayDate")?.value || "").trim();
  if (!rainDate) return alert("Pick a rain day first.");
  setStatus("Removing rain day...");
  try {
    const res = await fetchJson(`${API}/api/reports/rain-days/${encodeURIComponent(rainDate)}`, {
      method: "DELETE",
    });
    setText("rainDaysResult", JSON.stringify(res, null, 2));
    setStatus("Rain day removed.");
    loadRainDays({ silentNoRange: true }).catch(() => {});
  } catch (e) {
    setText("rainDaysResult", String(e.message || e));
    setStatus("Remove rain day failed.");
  }
}

function getRainDaysRangeQuery() {
  const month = (qs("costMonth")?.value || "").trim();
  const start = (qs("maintCostStart")?.value || "").trim();
  const end = (qs("maintCostEnd")?.value || "").trim();
  let q = "";
  if (month) {
    const d = new Date(`${month}-01T00:00:00`);
    const y = d.getFullYear();
    const m = d.getMonth();
    const startMonth = `${month}-01`;
    const endMonth = new Date(y, m + 1, 0).toISOString().slice(0, 10);
    q = `start=${encodeURIComponent(startMonth)}&end=${encodeURIComponent(endMonth)}`;
  } else if (start && end) {
    q = `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  }
  return q;
}

function renderRainDaysWidget(res) {
  const widget = qs("rainDaysWidget");
  if (!widget) return;
  const rows = Array.isArray(res?.rows) ? res.rows : [];
  const start = String(res?.start || "");
  const end = String(res?.end || "");
  if (!rows.length) {
    widget.innerHTML = `<div class="muted">No rain days recorded for ${start || "selected"}${end ? ` to ${end}` : ""}.</div>`;
    return;
  }
  widget.innerHTML = `
    <div><strong>Rain days (${rows.length})</strong> <span class="muted">${start} to ${end}</span></div>
    <div style="display:flex; flex-wrap:wrap; gap:8px; margin-top:8px;">
      ${rows.map((r) => `<span class="pill blue">🌧️ ${escapeHtml(String(r.rain_date || ""))}</span>`).join("")}
    </div>
    ${rows.some((r) => String(r.notes || "").trim())
      ? `<div class="muted" style="margin-top:8px;">Notes: ${rows
          .filter((r) => String(r.notes || "").trim())
          .map((r) => `${escapeHtml(String(r.rain_date || ""))} (${escapeHtml(String(r.notes || ""))})`)
          .join(" | ")}</div>`
      : ""}
  `;
}

async function loadRainDays(opts = {}) {
  const q = getRainDaysRangeQuery();
  if (!q) {
    if (!opts.silentNoRange) alert("Select month or start/end first.");
    return;
  }
  setStatus("Loading rain days...");
  try {
    const res = await fetchJson(`${API}/api/reports/rain-days?${q}`);
    setText("rainDaysResult", JSON.stringify(res, null, 2));
    renderRainDaysWidget(res);
    setStatus("Rain days loaded.");
  } catch (e) {
    setText("rainDaysResult", String(e.message || e));
    renderRainDaysWidget({ rows: [], start: "", end: "" });
    setStatus("Load rain days failed.");
  }
}

function openDailyPdf() {
  const date = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const scheduled = qs("scheduled")?.value || 10;
  const site = encodeURIComponent(getSessionSite());
  const ts = Date.now();
  window.open(
    `${API}/api/reports/daily.pdf?date=${encodeURIComponent(date)}&scheduled=${encodeURIComponent(scheduled)}&site_code=${site}&_ts=${ts}`,
    "_blank"
  );
}

function openWeeklyPdf() {
  const date = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const scheduled = qs("scheduled")?.value || 10;
  const r = getLast7Range(date);
  window.open(`${API}/api/reports/weekly.pdf?start=${r.start}&end=${r.end}&scheduled=${scheduled}`, "_blank");
}

function openLubePdf() {
  const start = qs("lubeStart")?.value || "";
  const end = qs("lubeEnd")?.value || "";
  if (!start || !end) return alert("Select lube period first.");
  window.open(`${API}/api/reports/lube.pdf?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`, "_blank");
}

function downloadLubeUsageXlsx() {
  const start = qs("lubeStart")?.value || "";
  const end = qs("lubeEnd")?.value || "";
  if (!start || !end) return alert("Select lube period first.");
  const monthRaw = String(qs("lubeStockMonth")?.value || "").trim();
  const month = /^\d{4}-\d{2}$/.test(monthRaw) ? monthRaw : start.slice(0, 7);
  const loc = String(qs("lubeStockLoc")?.value || "").trim().toUpperCase();
  const q = new URLSearchParams({ start, end, month });
  if (loc) q.set("location_code", loc);
  window.open(`${API}/api/reports/lube-usage-by-asset.xlsx?${q}`, "_blank");
}

function renderLubeMonthStockTable(data) {
  const el = qs("lubeMonthStockList");
  if (!el) return;
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  const loc = data?.location_code ? ` · ${escapeHtml(String(data.location_code))}` : "";
  const head = `<div class="muted mini" style="margin-bottom:6px;">Store balances${loc} — opening as-of <strong>${escapeHtml(String(data?.opening_as_of || ""))}</strong>, closing <strong>${escapeHtml(String(data?.closing_as_of || ""))}</strong> (${escapeHtml(String(data?.month || ""))})</div>`;
  if (!rows.length) {
    el.innerHTML = head + "<small>No lube stock rows returned.</small>";
    return;
  }
  const body = rows
    .map(
      (r) =>
        `<tr><td>${escapeHtml(String(r.part_code || ""))}</td><td>${escapeHtml(String(r.part_name || ""))}</td>` +
        `<td class="num">${Number(r.min_stock || 0).toFixed(1)}</td>` +
        `<td class="num">${Number(r.opening_qty || 0).toFixed(2)}</td>` +
        `<td class="num">${Number(r.closing_qty || 0).toFixed(2)}</td>` +
        `<td class="num">${Number(r.net_month_movement || 0).toFixed(2)}</td></tr>`
    )
    .join("");
  el.innerHTML =
    head +
    `<div style="overflow:auto;"><table class="gridTable" style="min-width:720px;"><thead><tr>` +
    `<th>Stock code</th><th>Description</th><th>Min</th><th>Opening</th><th>Closing</th><th>Net (month)</th>` +
    `</tr></thead><tbody>${body}</tbody></table></div>`;
}

async function loadLubeMonthStock() {
  const monthRaw = String(qs("lubeStockMonth")?.value || "").trim();
  const month = /^\d{4}-\d{2}$/.test(monthRaw) ? monthRaw : "";
  if (!month) {
    alert("Pick a month for store opening / closing balances.");
    return;
  }
  const loc = String(qs("lubeStockLoc")?.value || "").trim().toUpperCase();
  const q = new URLSearchParams({ month });
  if (loc) q.set("location_code", loc);
  setStatus("Loading lube month stock...");
  try {
    const data = await fetchJson(`${API}/api/stock/lube-month-stock?${q}`);
    renderLubeMonthStockTable(data);
    setStatus("Lube month stock ready.");
  } catch (e) {
    setStatus("Lube month stock failed: " + (e.message || e));
    renderLubeMonthStockTable({ rows: [], opening_as_of: "", closing_as_of: "", month: "" });
  }
}

function openStockMonitorPdf() {
  const filter = (qs("stockPartFilter")?.value || "").trim();
  const q = filter ? `?part_code=${encodeURIComponent(filter)}` : "";
  window.open(`${API}/api/reports/stock-monitor.pdf${q}`, "_blank");
}

function downloadStockMonitorPdf() {
  const filter = (qs("stockPartFilter")?.value || "").trim();
  const q = filter
    ? `?part_code=${encodeURIComponent(filter)}&download=1`
    : "?download=1";
  window.open(`${API}/api/reports/stock-monitor.pdf${q}`, "_blank");
}

function downloadAssetHistoryPdf() {
  const asset_code = getSelectedAssetCode();
  if (!asset_code) {
    alert("Select a fleet card first.");
    return;
  }
  const start = qs("histStart")?.value || "";
  const end = qs("histEnd")?.value || "";
  const url =
    `${API}/api/reports/asset-history/${encodeURIComponent(asset_code)}.pdf` +
    `?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&download=1`;
  window.open(url, "_blank");
}

function openOperationsPdf(download = false) {
  const start = (qs("opFrom")?.value || "").trim();
  const end = (qs("opTo")?.value || "").trim();
  if (!start || !end) {
    alert("Select operations date range first.");
    return;
  }
  const q = `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}${download ? "&download=1" : ""}`;
  window.open(`${API}/api/reports/operations.pdf?${q}`, "_blank");
}

function downloadOperationsXlsx() {
  const start = (qs("opFrom")?.value || "").trim();
  const end = (qs("opTo")?.value || "").trim();
  if (!start || !end) {
    alert("Select operations date range first.");
    return;
  }
  const q = `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  window.open(`${API}/api/reports/operations.xlsx?${q}`, "_blank");
}

/* =========================
   ACTIONS
========================= */

async function loadBreakdownOpsOpen() {
  const list = qs("boOpenList");
  if (!list) return;
  const d = (qs("boOpenDate")?.value || "").trim();
  const q = d ? `?date=${encodeURIComponent(d)}` : "";
  setSkeleton("boOpenList", 1);
  try {
    const data = await fetchJson(`${API}/api/breakdowns/open-all${q}`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    list.innerHTML = "";
    if (!rows.length) {
      list.appendChild(item("<small>No open incidents for this filter.</small>"));
      return;
    }
    rows.forEach((r) => {
      const bid = Number(r.id || 0);
      const wo = r.primary_work_order_id != null ? Number(r.primary_work_order_id) : "";
      const desc = escapeHtml(r.description || "");
      const code = escapeHtml(r.asset_code || "");
      const woSt = escapeHtml(String(r.primary_work_order_status || ""));
      list.appendChild(
        item(
          `<div><b>#${bid}</b> <span class="pill blue">${code}</span> <span class="pill orange">OPEN</span></div>` +
            `<small>${desc}</small><br/>` +
            `<small>WO: ${wo ? `#${wo} (${woSt})` : "—"} | Start: ${escapeHtml(String(r.start_at || r.breakdown_date || "—"))}</small><br/>` +
            `<button type="button" class="bo-copy-wo" data-wo="${wo}">Copy WO #</button> ` +
            `<button type="button" class="bo-close-bdn" data-id="${bid}">Close incident</button>`
        )
      );
    });
  } catch (e) {
    list.innerHTML = "";
    list.appendChild(item(`<span class="message-error">${escapeHtml(e.message || String(e))}</span>`));
  }
}

async function loadBreakdownOpsRecent() {
  const list = qs("boRecentList");
  if (!list) return;
  setSkeleton("boRecentList", 1);
  try {
    const rows = await fetchJson(`${API}/api/breakdowns`);
    const slice = (Array.isArray(rows) ? rows : []).slice(0, 20);
    list.innerHTML = "";
    if (!slice.length) {
      list.appendChild(item("<small>No breakdowns found.</small>"));
      return;
    }
    slice.forEach((r) => {
      const st = String(r.status || "").toUpperCase();
      list.appendChild(
        item(
          `<b>#${r.id}</b> ${escapeHtml(r.asset_code || "")} <span class="pill ${st === "OPEN" ? "orange" : "blue"}">${escapeHtml(r.status || "")}</span> ` +
            `${escapeHtml(r.breakdown_date || "")}<br/><small>${escapeHtml(r.description || "")}</small>`
        )
      );
    });
  } catch (e) {
    list.innerHTML = "";
    list.appendChild(item(`<span class="message-error">${escapeHtml(e.message || String(e))}</span>`));
  }
}

function refreshBreakdownOpsPanels() {
  loadBreakdownOpsOpen().catch(() => {});
  loadBreakdownOpsRecent().catch(() => {});
}

function initBoTyreRows() {
  const tb = qs("boTyreTbody");
  if (!tb || tb.dataset.ready === "1") return;
  tb.dataset.ready = "1";
  for (let i = 0; i < 10; i++) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td><input class="w-90" id="boT${i}_pos" placeholder="e.g. FL" /></td>
      <td><input class="w-120" id="boT${i}_sout" /></td>
      <td><input class="w-120" id="boT${i}_sin" /></td>
      <td><input class="w-90" id="boT${i}_tread" /></td>
      <td><input class="w-140" id="boT${i}_reason" /></td>
      <td><input class="w-90" type="number" step="0.1" min="0" id="boT${i}_hu" /></td>
      <td><input class="w-90" type="number" step="0.1" min="0" id="boT${i}_hf" /></td>
      <td><input class="w-120" id="boT${i}_part" list="partCodeOptions" /></td>
      <td><input class="w-90" type="number" step="0.01" min="0" id="boT${i}_cost" placeholder="ov." /></td>
      <td><input class="w-120" id="boT${i}_make" list="partCodeOptions" placeholder="Make code" /></td>`;
    tb.appendChild(tr);
  }
}

function updateBoSlipFormVisibility() {
  const t = qs("boSlipType")?.value || "hose_failure";
  const map = {
    hose_failure: "boWrapHose",
    get_change: "boWrapGet",
    component_change: "boWrapComp",
    tyre_change: "boWrapTyre",
  };
  Object.values(map).forEach((id) => {
    const el = qs(id);
    if (el) el.style.display = "none";
  });
  const showId = map[t];
  if (showId && qs(showId)) qs(showId).style.display = "";
  if (t === "tyre_change") initBoTyreRows();
}

function numOrUndef(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

const BO_SLIP_PIC_MAX = 4;
const BO_SLIP_PIC_MAX_BYTES = 512 * 1024;
let boSlipPicturesPayload = [];

function clearBoSlipPhotosUi() {
  boSlipPicturesPayload = [];
  const inp = qs("boSlipPhotosInput");
  if (inp) inp.value = "";
  const prev = qs("boSlipPhotosPreview");
  if (prev) prev.innerHTML = "";
}

function renderBoSlipPhotosPreview() {
  const prev = qs("boSlipPhotosPreview");
  if (!prev) return;
  prev.innerHTML = "";
  boSlipPicturesPayload.forEach((p, idx) => {
    const wrap = document.createElement("div");
    wrap.style.cssText =
      "position:relative;border:1px solid #cbd5e1;border-radius:6px;padding:4px;background:#f8fafc;";
    const img = document.createElement("img");
    img.src = `data:${p.mime};base64,${p.data_base64}`;
    img.alt = "";
    img.style.cssText = "max-width:120px;max-height:90px;display:block;";
    const rm = document.createElement("button");
    rm.type = "button";
    rm.textContent = "Remove";
    rm.style.cssText = "font-size:11px;margin-top:4px;";
    rm.addEventListener("click", () => {
      boSlipPicturesPayload = boSlipPicturesPayload.filter((_, i) => i !== idx);
      renderBoSlipPhotosPreview();
    });
    wrap.appendChild(img);
    wrap.appendChild(rm);
    prev.appendChild(wrap);
  });
}

async function boSlipReadPictureFile(file) {
  const mime = String(file.type || "").toLowerCase();
  if (mime !== "image/jpeg" && mime !== "image/png") {
    alert(`${file.name}: only JPEG or PNG images are supported.`);
    return null;
  }
  if (file.size > BO_SLIP_PIC_MAX_BYTES) {
    alert(`${file.name} is larger than 512 KB. Choose a smaller file or compress the image.`);
    return null;
  }
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => {
      const s = String(fr.result || "");
      const i = s.indexOf(",");
      const b64 = i >= 0 ? s.slice(i + 1) : "";
      resolve(b64 ? { mime, data_base64: b64 } : null);
    };
    fr.onerror = () => reject(new Error("Read failed"));
    fr.readAsDataURL(file);
  });
}

async function onBoSlipPhotosInputChange(ev) {
  const input = ev.target;
  const files = input?.files ? Array.from(input.files) : [];
  if (!files.length) return;
  for (const f of files) {
    if (boSlipPicturesPayload.length >= BO_SLIP_PIC_MAX) {
      alert(`Maximum ${BO_SLIP_PIC_MAX} pictures.`);
      break;
    }
    try {
      const pic = await boSlipReadPictureFile(f);
      if (pic) boSlipPicturesPayload.push(pic);
    } catch {
      setStatus("Could not read one of the pictures.");
    }
  }
  input.value = "";
  renderBoSlipPhotosPreview();
}

function collectBoTyreRows() {
  const tyres = [];
  for (let i = 0; i < 10; i++) {
    const position = String(qs(`boT${i}_pos`)?.value || "").trim();
    const serial_removed = String(qs(`boT${i}_sout`)?.value || "").trim();
    const serial_new = String(qs(`boT${i}_sin`)?.value || "").trim();
    const tread_left = String(qs(`boT${i}_tread`)?.value || "").trim();
    const reason = String(qs(`boT${i}_reason`)?.value || "").trim();
    const hours_in_use = numOrUndef(qs(`boT${i}_hu`)?.value);
    const hours_fitted = numOrUndef(qs(`boT${i}_hf`)?.value);
    const part_code = String(qs(`boT${i}_part`)?.value || "").trim();
    const cost_manual = numOrUndef(qs(`boT${i}_cost`)?.value);
    const tyre_make_part_code = String(qs(`boT${i}_make`)?.value || "").trim();
    if (
      !position &&
      !serial_removed &&
      !serial_new &&
      !reason &&
      !part_code &&
      !tyre_make_part_code &&
      hours_in_use == null &&
      hours_fitted == null
    ) {
      continue;
    }
    tyres.push({
      position,
      serial_removed,
      serial_new,
      tread_left,
      reason,
      hours_in_use,
      hours_fitted,
      part_code,
      cost_manual,
      tyre_make_part_code,
    });
  }
  return tyres;
}

function boSlipQtyOrUndef(el) {
  const n = Number(el?.value);
  if (!Number.isFinite(n) || n <= 0) return undefined;
  return n;
}

async function pullBoSlipFromAsset() {
  const asset_code = String(qs("boSlipAsset")?.value || "").trim();
  if (!asset_code) {
    alert("Enter asset code first.");
    return;
  }
  const slip_type = String(qs("boSlipType")?.value || "").trim();
  setStatus("Loading asset slip hints...");
  try {
    const data = await fetchJson(
      `${API}/api/breakdown-ops/slip-asset-hints?asset_code=${encodeURIComponent(asset_code)}`
    );
    if (slip_type === "get_change" && data.get_change) {
      const h = data.get_change;
      const gh = qs("boGetHours");
      if (gh && h.hours_fitted != null && Number.isFinite(Number(h.hours_fitted))) gh.value = String(h.hours_fitted);
      if (qs("boGetPart")) qs("boGetPart").value = h.part_code || "";
      if (qs("boGetPartQty")) qs("boGetPartQty").value = String(h.part_qty != null ? h.part_qty : 1);
      if (qs("boGetSupplier")) qs("boGetSupplier").value = h.supplier || "";
      if (qs("boGetDateChg")) qs("boGetDateChg").value = h.date_changed || "";
      if (qs("boGetDescPart")) qs("boGetDescPart").value = h.description_part_code || "";
      if (qs("boGetDescPartQty")) qs("boGetDescPartQty").value = String(h.description_part_qty != null ? h.description_part_qty : 1);
      if (h.notes && qs("boGetNotes")) qs("boGetNotes").value = h.notes;
      setStatus(`G.E.T. fields filled from ${data.get_change_source || "asset"}.`);
    } else if (slip_type === "hose_failure" && data.hose_failure) {
      const h = data.hose_failure;
      if (qs("boHoseDateFitted")) qs("boHoseDateFitted").value = h.date_fitted || "";
      if (qs("boHoseReason")) qs("boHoseReason").value = h.reason_fitted || "";
      if (qs("boHosePreventable")) qs("boHosePreventable").checked = Boolean(h.preventable);
      if (qs("boHosePart")) qs("boHosePart").value = h.hose_part_code || "";
      if (qs("boHoseQty")) qs("boHoseQty").value = String(h.hose_qty != null ? h.hose_qty : 1);
      if (qs("boOilPart")) qs("boOilPart").value = h.oil_loss_part_code || "";
      if (qs("boOilQty")) qs("boOilQty").value = String(h.oil_loss_qty != null ? h.oil_loss_qty : 1);
      if (h.notes && qs("boHoseNotes")) qs("boHoseNotes").value = h.notes;
      setStatus(`Hose fields filled from ${data.hose_failure_source || "asset"}.`);
    } else {
      setStatus(
        slip_type === "get_change"
          ? "No GET data: add a GET change slip on the asset (Assets → history) or save a G.E.T. slip first."
          : slip_type === "hose_failure"
            ? "No prior hose failure slip for this asset yet."
            : "Pull from asset applies to G.E.T. or Hose failure slips."
      );
    }
  } catch (e) {
    setStatus("Asset hints failed: " + (e.message || e));
  }
}

async function saveBoSlipReport() {
  const slip_type = String(qs("boSlipType")?.value || "").trim();
  const asset_code = String(qs("boSlipAsset")?.value || "").trim();
  const report_date = String(qs("boSlipDate")?.value || "").trim();
  if (!asset_code || !report_date) {
    alert("Asset code and report date are required.");
    return;
  }
  let body = { slip_type, asset_code, report_date };
  if (slip_type === "hose_failure") {
    Object.assign(body, {
      date_fitted: String(qs("boHoseDateFitted")?.value || "").trim(),
      reason_fitted: String(qs("boHoseReason")?.value || "").trim(),
      preventable: Boolean(qs("boHosePreventable")?.checked),
      hose_part_code: String(qs("boHosePart")?.value || "").trim(),
      oil_loss_part_code: String(qs("boOilPart")?.value || "").trim(),
      hose_qty: boSlipQtyOrUndef(qs("boHoseQty")),
      oil_loss_qty: boSlipQtyOrUndef(qs("boOilQty")),
      hose_cost_manual: numOrUndef(qs("boHoseCostOv")?.value),
      oil_cost_manual: numOrUndef(qs("boOilCostOv")?.value),
      notes: String(qs("boHoseNotes")?.value || "").trim() || undefined,
    });
  } else if (slip_type === "get_change") {
    Object.assign(body, {
      hours_fitted: numOrUndef(qs("boGetHours")?.value),
      part_code: String(qs("boGetPart")?.value || "").trim(),
      part_qty: boSlipQtyOrUndef(qs("boGetPartQty")),
      supplier: String(qs("boGetSupplier")?.value || "").trim(),
      date_changed: String(qs("boGetDateChg")?.value || "").trim(),
      description_part_code: String(qs("boGetDescPart")?.value || "").trim(),
      description_part_qty: boSlipQtyOrUndef(qs("boGetDescPartQty")),
      notes: String(qs("boGetNotes")?.value || "").trim() || undefined,
    });
  } else if (slip_type === "component_change") {
    Object.assign(body, {
      date_changed: String(qs("boCompDate")?.value || "").trim(),
      hours_in_service: numOrUndef(qs("boCompHrsSvc")?.value),
      reason: String(qs("boCompReason")?.value || "").trim(),
      component_type: String(qs("boCompType")?.value || "").trim(),
      part_code: String(qs("boCompPart")?.value || "").trim(),
      cost_manual: numOrUndef(qs("boCompCostOv")?.value),
      notes: String(qs("boCompNotes")?.value || "").trim() || undefined,
    });
  } else if (slip_type === "tyre_change") {
    const tyres = collectBoTyreRows();
    if (!tyres.length) {
      alert("Enter at least one tyre line.");
      return;
    }
    Object.assign(body, { tyres, notes: String(qs("boTyreNotes")?.value || "").trim() || undefined });
  } else {
    alert("Unknown slip type.");
    return;
  }

  if (boSlipPicturesPayload.length) {
    body.pictures = boSlipPicturesPayload.map((p) => ({
      mime: p.mime,
      data_base64: p.data_base64,
    }));
  }

  setStatus("Saving slip report...");
  try {
    const res = await fetchJson(`${API}/api/breakdown-ops/slips`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    setText("boSlipResult", JSON.stringify(res, null, 2));
    setStatus("Slip saved.");
    clearBoSlipPhotosUi();
    if (res.id) {
      window.open(`${API}/api/breakdown-ops/slips/${res.id}/pdf`, "_blank");
    }
    await loadBoSlipSavedList();
  } catch (e) {
    setText("boSlipResult", String(e.message || e));
    setStatus("Slip save failed.");
  }
}

async function loadBoSlipSavedList() {
  const list = qs("boSlipSavedList");
  if (!list) return;
  list.innerHTML = `<div class="muted">Loading…</div>`;
  try {
    const data = await fetchJson(`${API}/api/breakdown-ops/slips`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    list.innerHTML = "";
    if (!rows.length) {
      list.appendChild(item("<small>No slip reports saved yet.</small>"));
      return;
    }
    rows.slice(0, 40).forEach((r) => {
      list.appendChild(renderBoSlipSavedRow(r));
    });
  } catch (e) {
    list.innerHTML = "";
    list.appendChild(item(`<span class="message-error">${escapeHtml(e.message || String(e))}</span>`));
  }
}

function renderBoSlipSavedRow(r) {
  const el = document.createElement("div");
  el.className = "bo-slip-row";
  const label = escapeHtml(String(r.slip_type || "").replace(/_/g, " "));
  el.innerHTML = `
    <div class="bo-slip-row-main">
      <span class="bo-slip-id">#${Number(r.id)}</span>
      <span class="pill blue">${label}</span>
      <span class="bo-slip-asset">${escapeHtml(r.asset_code || "—")}</span>
      <span class="bo-slip-date">${escapeHtml(r.report_date || "")}</span>
    </div>
    <div class="bo-slip-row-actions">
      <button type="button" class="bo-slip-pdf btn btn-secondary btn-sm" data-id="${Number(r.id)}">Open PDF</button>
    </div>
  `;
  return el;
}

function openBoSlipPdf(id) {
  const n = Number(id || 0);
  if (!n) return;
  window.open(`${API}/api/breakdown-ops/slips/${n}/pdf`, "_blank");
}

async function ensureOpenBreakdownOps() {
  const asset_code = (qs("boEnsureAsset")?.value || "").trim();
  const breakdown_date = (qs("boEnsureDate")?.value || "").trim() || new Date().toISOString().slice(0, 10);
  if (!asset_code) return alert("Enter asset code.");
  setStatus("Ensuring open breakdown...");
  try {
    const res = await fetchJson(`${API}/api/breakdowns/ensure-open`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ asset_code, breakdown_date, description: "Down - Breakdown Ops" }),
    });
    setText("boEnsureResult", JSON.stringify(res, null, 2));
    if (res.primary_work_order_id && qs("iWo")) qs("iWo").value = String(res.primary_work_order_id);
    setStatus("Ensure open complete.");
    await loadBreakdownOpsOpen();
    await loadBreakdownOpsRecent();
  } catch (e) {
    setText("boEnsureResult", String(e.message || e));
    setStatus("Ensure open failed.");
  }
}

let lastBoRepairWoId = null;

async function createRepairWorkOrderOps() {
  const asset_code = (qs("boRepairAsset")?.value || "").trim();
  const component = (qs("boRepairComponent")?.value || "").trim();
  const description = (qs("boRepairDescription")?.value || "").trim();
  const inspectionRaw = (qs("boRepairInspectionId")?.value || "").trim();
  const inspection_id = inspectionRaw ? Number(inspectionRaw) : 0;
  if (!asset_code) return alert("Enter asset code.");
  if (!description) return alert("Enter repair description.");
  setStatus("Creating repair work order…");
  try {
    const body = { asset_code, description };
    if (component) body.component = component;
    if (Number.isFinite(inspection_id) && inspection_id > 0) body.inspection_id = inspection_id;
    const res = await fetchJson(`${API}/api/workorders/repair`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    lastBoRepairWoId = Number(res.work_order_id || 0) || null;
    setText("boRepairResult", JSON.stringify(res, null, 2));
    const openBtn = qs("boRepairOpenWo");
    if (openBtn && lastBoRepairWoId) openBtn.style.display = "";
    setStatus(
      res.already_exists
        ? `WO #${lastBoRepairWoId} already linked to inspection.`
        : `Repair WO #${lastBoRepairWoId} created (no breakdown).`
    );
  } catch (e) {
    setText("boRepairResult", String(e.message || e));
    setStatus("Repair WO create failed.");
  }
}

async function pullBreakdownOpsLiveHours() {
  const hint = qs("boLiveHoursHint");
  const code = (qs("sqAsset")?.value || "").trim();
  const asOf =
    (qs("sqDate")?.value || "").trim() || (qs("date")?.value || "").trim() || new Date().toISOString().slice(0, 10);
  if (!code) {
    alert("Enter asset code on the short breakdown line first.");
    return;
  }
  if (hint) hint.textContent = "Loading…";
  try {
    const rows = await fetchJson(`${API}/api/assets?include_archived=0`);
    const arr = Array.isArray(rows) ? rows : [];
    const a = arr.find((x) => String(x.asset_code || "").toUpperCase() === code.toUpperCase());
    if (!a) throw new Error("Asset not found");
    const q = asOf ? `?as_of=${encodeURIComponent(asOf)}` : "";
    const data = await fetchJson(`${API}/api/maintenance/asset/${a.id}/live-hours${q}`);
    const h = Number(data.current_hours || 0);
    const src = String(data.current_hours_source || "");
    if (hint) {
      hint.textContent = `Meter (as of ${asOf}): ${h.toFixed(1)} h (${src}).`;
    }
    setStatus("Live meter loaded.");
  } catch (e) {
    if (hint) hint.textContent = e.message || String(e);
  }
}

async function closeBreakdownFromOps(breakdownId) {
  const id = Number(breakdownId || 0);
  if (!id) return;
  if (!confirm(`Close breakdown #${id}? Component work orders must be closed first.`)) return;
  setStatus("Closing breakdown...");
  try {
    await fetchJson(`${API}/api/breakdowns/${id}/close`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({}),
    });
    setStatus("Breakdown closed.");
    await loadBreakdownOpsOpen();
    await loadBreakdownOpsRecent();
    await loadDashboard().catch(() => {});
  } catch (e) {
    alert(e.message || String(e));
    setStatus("Close failed.");
  }
}

async function createBreakdown() {
  const date = (qs("bDate")?.value || qs("date")?.value || new Date().toISOString().slice(0, 10)).trim();
  const payload = {
    asset_code: (qs("bAsset")?.value || "").trim(),
    breakdown_date: date,
    description: (qs("bDesc")?.value || "").trim(),
    downtime_hours: Number(qs("bDown")?.value || 0),
    critical: !!qs("bCrit")?.checked,
    parts_ordered_date: String(qs("bPartsOrderedDate")?.value || "").trim() || null,
    parts_status: String(qs("bPartsStatus")?.value || "").trim() || null,
    parts_received_date: String(qs("bPartsReceivedDate")?.value || "").trim() || null,
    ets_repair_date: String(qs("bEtsRepairDate")?.value || "").trim() || null,
  };

  setStatus("Creating breakdown...");
  try {
    const res = await fetchJson(`${API}/api/breakdowns`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("breakdownResult", JSON.stringify(res, null, 2));
    setStatus("Breakdown created.");
    await loadDashboard().catch(() => {});
    refreshBreakdownOpsPanels();
  } catch (e) {
    setText("breakdownResult", String(e.message || e));
    setStatus("Breakdown failed.");
  }
}

/** Short breakdown: closed incident + downtime logs + parts + oils in one POST. */
function collectShortBreakdownParts() {
  const parts = [];
  document.querySelectorAll("#sqPartsRows .sq-part-row").forEach((row) => {
    const part_code = String(row.querySelector(".sq-part-code")?.value || "").trim();
    const quantity = Number(row.querySelector(".sq-part-qty")?.value || 0);
    if (part_code && Number.isFinite(quantity) && quantity > 0) {
      parts.push({ part_code, quantity });
    }
  });
  return parts;
}

function addShortBreakdownPartRow() {
  const container = qs("sqPartsRows");
  if (!container) return;
  const row = document.createElement("div");
  row.className = "row sq-part-row";
  row.innerHTML = `
    <input class="sq-part-code w-180" list="partCodeOptions" placeholder="Part code" />
    <input class="sq-part-qty w-70" type="number" min="1" step="1" placeholder="Qty" />
    <button type="button" class="sq-part-remove" title="Remove part line">Remove</button>
  `;
  container.appendChild(row);
}

function bindShortBreakdownPartsUi() {
  qs("sqAddPart")?.addEventListener("click", () => addShortBreakdownPartRow());
  qs("sqPartsRows")?.addEventListener("click", (ev) => {
    const btn = ev.target?.closest?.(".sq-part-remove");
    if (!btn) return;
    const row = btn.closest(".sq-part-row");
    const container = qs("sqPartsRows");
    if (!row || !container) return;
    if (container.querySelectorAll(".sq-part-row").length <= 1) {
      row.querySelector(".sq-part-code").value = "";
      row.querySelector(".sq-part-qty").value = "";
      return;
    }
    row.remove();
  });
}

async function submitShortBreakdown() {
  const headerDate = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const breakdown_date = (qs("sqDate")?.value || headerDate).trim();
  const asset_code = (qs("sqAsset")?.value || "").trim();
  const description = (qs("sqDesc")?.value || "").trim();
  const td = (qs("sqTimeDown")?.value || "").trim();
  const tu = (qs("sqTimeUp")?.value || "").trim();
  const comp = (qs("sqComponent")?.value || "").trim();

  const parts = collectShortBreakdownParts();

  const oils = [];
  const oilType = (qs("sqOilType")?.value || "").trim();
  const oilQty = Number(qs("sqOilQty")?.value || 0);
  if (oilType && Number.isFinite(oilQty) && oilQty > 0) oils.push({ oil_type: oilType, quantity: oilQty });

  if (!asset_code || !description) {
    alert("Asset code and description are required.");
    return;
  }

  const payload = {
    asset_code,
    breakdown_date,
    description,
    critical: !!qs("sqCrit")?.checked,
    parts,
    oils,
  };
  if (comp) payload.component = comp;

  if (td && tu) {
    payload.time_down = td;
    payload.time_up = tu;
  } else if (!td && !tu) {
    const h = Number(qs("sqHours")?.value);
    if (Number.isNaN(h) || h <= 0 || h > 24) {
      alert("Enter both time down and time up, or a single hours-down value (0–24) for the event date.");
      return;
    }
    payload.hours_down = h;
  } else {
    alert("Provide both time down and time up, or clear both and use hours down.");
    return;
  }

  setStatus("Logging short breakdown...");
  try {
    const res = await fetchJson(`${API}/api/breakdowns/short-complete`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("shortBreakdownResult", JSON.stringify(res, null, 2));
    setStatus("Short breakdown logged.");
    await loadDashboard().catch(() => {});
    refreshBreakdownOpsPanels();
  } catch (e) {
    setText("shortBreakdownResult", String(e.message || e));
    setStatus("Short breakdown failed.");
  }
}

async function issuePart() {
  const woId = (qs("iWo")?.value || "").trim();
  const payload = {
    part_code: (qs("iPart")?.value || "").trim(),
    quantity: Number(qs("iQty")?.value || 1),
  };

  setStatus("Issuing part...");
  try {
    const res = await fetchJson(`${API}/api/workorders/${woId}/issue`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("issueResult", JSON.stringify(res, null, 2));
    setStatus("Part issued.");
    await loadDashboard().catch(() => {});
  } catch (e) {
    setText("issueResult", String(e.message || e));
    setStatus("Issue failed.");
  }
}

async function allocateStore() {
  const payload = {
    part_code: (qs("saPart")?.value || "").trim(),
    quantity: Number(qs("saQty")?.value || 0),
    location_code: (qs("saLocation")?.value || "").trim() || undefined,
    bin_code: (qs("saBin")?.value || "").trim() || undefined,
    asset_code: (qs("saAsset")?.value || "").trim() || undefined,
    work_order_id: (qs("saWo")?.value || "").trim() ? Number((qs("saWo")?.value || "").trim()) : undefined,
    allocation_date: (qs("saDate")?.value || "").trim() || undefined,
    issued_by: (qs("saIssuedBy")?.value || "").trim() || undefined,
    cost_center_code: (qs("saCostCenter")?.value || "").trim() || undefined,
    notes: (qs("saNotes")?.value || "").trim() || undefined,
  };

  setStatus("Allocating stores...");
  try {
    const res = await fetchJson(`${API}/api/stock/allocate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("storeAllocResult", JSON.stringify(res, null, 2));
    setStatus("Stores allocated.");
    await Promise.all([
      loadStoreAllocations().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
  } catch (e) {
    setText("storeAllocResult", String(e.message || e));
    setStatus("Stores allocation failed.");
  }
}

async function loadStoreAllocations() {
  const list = qs("storeAllocList");
  if (!list) return;

  list.innerHTML = "";
  setSkeleton("storeAllocList", 2);

  const rows = await fetchJson(`${API}/api/stock/allocations`);
  const data = Array.isArray(rows?.rows) ? rows.rows : [];

  list.innerHTML = "";
  data.slice(0, 20).forEach((r) => {
    const ref = r.work_order_id ? `WO #${r.work_order_id}` : r.asset_code;
    const unitCost = Number(r.unit_cost || 0);
    const lineValue = Number((unitCost * Number(r.quantity || 0)).toFixed(2));
    list.appendChild(
      item(
        `<b>${r.allocation_date}</b> — ${r.part_code} x ${Number(r.quantity || 0).toFixed(1)}<br>` +
        `<small>${ref} | ${r.location_code || "NO-LOC"}${r.bin_code ? `/${r.bin_code}` : ""} | CC: ${r.cost_center_code || "-"} | Unit: $${unitCost.toFixed(2)} | Value: $${lineValue.toFixed(2)} | ${r.issued_by || "No issuer"}${r.notes ? ` | ${r.notes}` : ""}</small>`
      )
    );
  });

  if (!data.length) {
    list.appendChild(item("<small>No store allocations yet.</small>"));
  }
}

function applyAssetCostCenterToInputs(assetCode) {
  const code = String(assetCode || "").trim();
  const a = code && window.__assetsByCode ? window.__assetsByCode[code] : null;
  const cc = a?.cost_center_code ? String(a.cost_center_code) : "";
  if (qs("fuelCostCenter")) qs("fuelCostCenter").value = cc;
  if (qs("mlCostCenter")) qs("mlCostCenter").value = cc;
}

async function loadCostCenterOptions() {
  const list = qs("costCenterOptions");
  if (!list) return;
  try {
    const data = await fetchJson(`${API}/api/masterdata/cost-centers`);
    const rows = Array.isArray(data?.rows) ? data.rows : [];
    list.innerHTML = "";
    rows.forEach((r) => {
      const code = String(r.code || "").trim();
      if (!code) return;
      const opt = document.createElement("option");
      opt.value = code;
      opt.label = String(r.name || code);
      list.appendChild(opt);
    });
  } catch {}
}

async function populateAssetAllocSelect() {
  const sel = qs("assetAllocSelect");
  if (!sel) return;
  const current = String(sel.value || "");
  try {
    const assets = await fetchJson(`${API}/api/assets?include_archived=0`);
    const rows = Array.isArray(assets) ? assets : [];
    window.__assetsByCode = {};
    sel.innerHTML = '<option value="">Select asset…</option>';
    rows.forEach((a) => {
      const code = String(a.asset_code || "").trim();
      if (!code) return;
      window.__assetsByCode[code] = a;
      const opt = document.createElement("option");
      opt.value = code;
      const cc = a.cost_center_code ? ` · ${a.cost_center_code}` : "";
      opt.textContent = `${code} — ${a.asset_name || ""}${cc}`;
      sel.appendChild(opt);
    });
    if (current && window.__assetsByCode[current]) sel.value = current;
    loadAssetAllocationForm();
  } catch (e) {
    setStatus("Asset allocation list failed: " + (e.message || e));
  }
}

function loadAssetAllocationForm() {
  const code = String(qs("assetAllocSelect")?.value || "").trim();
  const a = code && window.__assetsByCode ? window.__assetsByCode[code] : null;
  if (qs("assetAllocSite")) qs("assetAllocSite").value = a?.site_code ? String(a.site_code) : (getSessionSite() || "main");
  if (qs("assetAllocCostCenter")) qs("assetAllocCostCenter").value = a?.cost_center_code ? String(a.cost_center_code) : "";
  if (qs("assetAllocDepartment")) qs("assetAllocDepartment").value = a?.department_code ? String(a.department_code) : "";
}

async function saveAssetAllocation() {
  const code = String(qs("assetAllocSelect")?.value || "").trim();
  const out = qs("assetAllocResult");
  if (!code) return alert("Select an asset first.");
  const body = {
    site_code: String(qs("assetAllocSite")?.value || "").trim() || null,
    cost_center_code: String(qs("assetAllocCostCenter")?.value || "").trim() || null,
    department_code: String(qs("assetAllocDepartment")?.value || "").trim() || null,
  };
  setStatus("Saving asset allocation…");
  try {
    const res = await fetchJson(`${API}/api/assets/${encodeURIComponent(code)}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (out) out.textContent = JSON.stringify(res?.asset || res, null, 2);
    await Promise.all([
      populateAssetAllocSelect(),
      loadCodePickers(),
      loadCostCenterOptions(),
    ]);
    setStatus(`Allocation saved for ${code}.`);
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Asset allocation save failed.");
  }
}

async function loadCodePickers() {
  const assetList = qs("assetCodeOptions");
  const partList = qs("partCodeOptions");
  const locationList = qs("locationCodeOptions");
  if (!assetList && !partList && !locationList) return;

  if (assetList) {
    try {
      const assets = await fetchJson(`${API}/api/assets?include_archived=0`);
      window.__assetsByCode = window.__assetsByCode || {};
      assetList.innerHTML = "";
      (Array.isArray(assets) ? assets : []).forEach((a) => {
        const code = String(a.asset_code || "").trim();
        if (!code) return;
        window.__assetsByCode[code] = a;
        const hiredTag = isHiredAsset(a) ? " [HIRED]" : "";
        const opt = document.createElement("option");
        opt.value = code;
        opt.textContent = `${code}${hiredTag} - ${a.asset_name || ""}`;
        assetList.appendChild(opt);
      });
    } catch {}
  }
  await loadCostCenterOptions();

  if (partList) {
    try {
      const parts = await fetchJson(`${API}/api/stock/onhand`);
      partList.innerHTML = "";
      const map = {};
      (Array.isArray(parts) ? parts : []).forEach((p) => {
        const code = String(p.part_code || "").trim();
        if (!code) return;
        map[String(code).toUpperCase()] = String(p.part_name || "").trim();
        const opt = document.createElement("option");
        opt.value = code;
        opt.textContent = `${code} - ${p.part_name || ""}`;
        partList.appendChild(opt);
      });
      window.__partNameByCode = map;
    } catch {}
  }

  if (locationList) {
    try {
      const locations = await fetchJson(`${API}/api/stock/locations?active=1`);
      const rows = Array.isArray(locations?.rows) ? locations.rows : [];
      locationList.innerHTML = "";
      rows.forEach((l) => {
        const code = String(l.location_code || "").trim();
        if (!code) return;
        const locText = Array.isArray(r.allowed_locations) && r.allowed_locations.length ? r.allowed_locations.join(", ") : "all";
        tr.innerHTML = `<td>${escapeHtml(r.username)}</td><td>${escapeHtml(r.full_name || "")}</td><td>${escapeHtml(r.department || "")}</td><td>${escapeHtml(rolesText)}</td><td>${escapeHtml(locText)}</td><td>${r.active ? "yes" : "no"}</td><td>${r.has_password ? "yes" : "no"}</td>`;
        opt.value = code;
        opt.textContent = `${code}${l.location_name ? ` - ${l.location_name}` : ""}`;
        locationList.appendChild(opt);
      });
    } catch {}
  }

  applyDefaultLocationsToInputs();
}

function updateManualStockCostRowVisibility() {
  const t = String(qs("msType")?.value || "in").trim().toLowerCase();
  const row = qs("msCostRow");
  if (row) row.style.display = t === "in" ? "" : "none";
}

function updateManualStockPartDesc() {
  const code = String(qs("msPart")?.value || "").trim().toUpperCase();
  const descEl = qs("msPartDesc");
  if (!descEl) return;
  const map = window.__partNameByCode || {};
  const name = code && map && map[code] ? String(map[code]) : "";
  if (name) {
    descEl.value = name;
    descEl.disabled = true;
    return;
  }
  descEl.disabled = false;
  if (!code) {
    descEl.value = "";
    return;
  }
  descEl.value = descEl.value || "";
  fetchPartNameByCode(code).then((n) => {
    const now = String(qs("msPart")?.value || "").trim().toUpperCase();
    if (now !== code) return;
    if (n) {
      descEl.value = n;
      descEl.disabled = true;
    } else {
      descEl.disabled = false;
    }
  });
}

// Fallback lookup (covers cases where code pickers haven't loaded yet)
const __partNameFetchCache = new Map();
let __partNameFetchSeq = 0;
async function fetchPartNameByCode(code) {
  const c = String(code || "").trim().toUpperCase();
  if (!c) return "";
  if (__partNameFetchCache.has(c)) return __partNameFetchCache.get(c) || "";
  const mySeq = ++__partNameFetchSeq;
  try {
    const data = await fetchJson(`${API}/api/stock/control-summary?part_code=${encodeURIComponent(c)}`);
    const name = String(data?.part?.part_name || "").trim();
    if (mySeq === __partNameFetchSeq) {
      __partNameFetchCache.set(c, name || "");
      if (!window.__partNameByCode) window.__partNameByCode = {};
      window.__partNameByCode[c] = name || "";
    }
    return name || "";
  } catch {
    __partNameFetchCache.set(c, "");
    return "";
  }
}

function updateManualLubePartDesc() {
  const code = String(qs("mlPart")?.value || "").trim().toUpperCase();
  const descEl = qs("mlPartDesc");
  if (!descEl) return;
  const map = window.__partNameByCode || {};
  const name = code && map && map[code] ? String(map[code]) : "";
  descEl.value = name || "";
  if (!descEl.value && code) {
    fetchPartNameByCode(code).then((n) => {
      const now = String(qs("mlPart")?.value || "").trim().toUpperCase();
      if (now !== code) return;
      descEl.value = n || "";
    });
  }
}

function updateLubeMinPartDesc() {
  const code = String(qs("lubeMinPart")?.value || "").trim().toUpperCase();
  const descEl = qs("lubeMinPartDesc");
  if (!descEl) return;
  const map = window.__partNameByCode || {};
  const name = code && map && map[code] ? String(map[code]) : "";
  descEl.value = name || "";
  if (!descEl.value && code) {
    fetchPartNameByCode(code).then((n) => {
      const now = String(qs("lubeMinPart")?.value || "").trim().toUpperCase();
      if (now !== code) return;
      descEl.value = n || "";
    });
  }
}

function updateReceiveLubePartDesc() {
  const code = String(qs("lrPart")?.value || "").trim().toUpperCase();
  const descEl = qs("lrPartDesc");
  if (!descEl) return;
  const map = window.__partNameByCode || {};
  const name = code && map && map[code] ? String(map[code]) : "";
  descEl.value = name || (descEl.value || "");
  // If code is unknown, allow manual description entry (needed to create new stock items)
  descEl.disabled = !!name;
}

async function setThisLubeMinimum() {
  const part_code = (qs("mlPart")?.value || "").trim();
  const min_stock = Number(qs("mlMinInput")?.value || 0);
  if (!part_code) return alert("Enter a lube stock number first.");
  if (!Number.isFinite(min_stock) || min_stock < 0) return alert("Minimum must be >= 0.");
  setStatus("Saving lube minimum...");
  try {
    const res = await fetchJson(`${API}/api/stock/part-minimum`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ part_code, min_stock }),
    });
    setText("manualLubeResult", JSON.stringify(res, null, 2));
    await Promise.all([
      loadLubeStockOnHand().catch(() => {}),
      loadStockOnHandPage().catch(() => {}),
      loadInventoryControl().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
    setStatus("Lube minimum updated.");
  } catch (e) {
    setText("manualLubeResult", String(e.message || e));
    setStatus("Failed to set lube minimum.");
  }
}

async function receiveLubeStock() {
  const part_code = (qs("lrPart")?.value || "").trim();
  const location_code = (qs("lrLocation")?.value || "").trim() || "LUBE";
  const quantity = Number(qs("lrQty")?.value || 0);
  const reference = (qs("lrRef")?.value || "").trim() || "lube_receive";
  const part_name = (qs("lrPartDesc")?.value || "").trim();
  if (!part_code) return alert("Enter lube stock number.");
  if (!Number.isFinite(quantity) || quantity <= 0) return alert("Quantity must be > 0.");

  setStatus("Receiving lube stock...");
  try {
    const res = await fetchJson(`${API}/api/stock/movement`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        part_code,
        movement_type: "in",
        quantity,
        reference,
        location_code,
        part_name: part_name || undefined,
        create_if_missing: true,
      }),
    });
    setText("receiveLubeResult", JSON.stringify(res, null, 2));
    await Promise.all([
      loadLubeStockOnHand().catch(() => {}),
      loadStockOnHandPage().catch(() => {}),
      loadInventoryControl().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
    setStatus("Lube stock received.");
  } catch (e) {
    setText("receiveLubeResult", String(e.message || e));
    setStatus("Receive lube failed.");
  }
}

async function saveManualStock() {
  const movement_type = String(qs("msType")?.value || "in").trim().toLowerCase();
  const part_code = String(qs("msPart")?.value || "").trim().toUpperCase();
  const part_name = String(qs("msPartDesc")?.value || "").trim();
  const rawCost = String(qs("msUnitCost")?.value || "").trim();
  const unit_cost =
    movement_type === "in" && rawCost !== "" && Number.isFinite(Number(rawCost)) && Number(rawCost) > 0
      ? Number(rawCost)
      : undefined;
  const cost_currency = String(qs("msCostCurrency")?.value || "USD").trim().toUpperCase();

  const payload = {
    part_code,
    location_code: (qs("msLocation")?.value || "").trim() || undefined,
    bin_code: (qs("msBin")?.value || "").trim() || undefined,
    movement_type,
    quantity: Number(qs("msQty")?.value || 0),
    reference: (qs("msRef")?.value || "").trim() || undefined,
    cost_center_code: (qs("msCostCenter")?.value || "").trim() || undefined,
    ...(movement_type === "in"
      ? { create_if_missing: true, ...(part_name ? { part_name } : {}) }
      : {}),
    ...(movement_type === "in" && unit_cost != null
      ? { unit_cost, cost_currency: ["USD", "ZAR", "MZN"].includes(cost_currency) ? cost_currency : "USD" }
      : {}),
  };

  setStatus("Saving manual stock entry...");
  try {
    const res = await fetchJson(`${API}/api/stock/movement`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("manualStockResult", JSON.stringify(res, null, 2));
    setStatus("Manual stock saved.");
    if (part_code && part_name) {
      __partNameFetchCache.set(part_code, part_name);
      if (!window.__partNameByCode) window.__partNameByCode = {};
      window.__partNameByCode[part_code] = part_name;
    }
    // Clear form for fast consecutive entries.
    if (qs("msPart")) qs("msPart").value = "";
    if (qs("msPartDesc")) qs("msPartDesc").value = "";
    if (qs("msQty")) qs("msQty").value = "1";
    if (qs("msRef")) qs("msRef").value = "";
    if (qs("msBin")) qs("msBin").value = "";
    if (qs("msCostCenter")) qs("msCostCenter").value = "";
    if (qs("msUnitCost")) qs("msUnitCost").value = "";
    if (qs("msType")) qs("msType").value = "in";
    if (qs("msCostCurrency")) qs("msCostCurrency").value = "USD";
    updateManualStockCostRowVisibility();
    qs("msPart")?.focus();
    await loadDashboard().catch(() => {});
  } catch (e) {
    setText("manualStockResult", String(e.message || e));
    setStatus("Manual stock save failed.");
  }
}

async function loadInventoryControl() {
  const part_code = (qs("icPartCode")?.value || "").trim();
  const q = part_code ? `?part_code=${encodeURIComponent(part_code)}` : "";
  setStatus("Loading inventory control...");
  setSkeleton("icLubeLowList", 1);
  try {
    const data = await fetchJson(`${API}/api/stock/control-summary${q}`);
    const summary = data.summary || {};
    const part = data.part || null;
    const lubeRows = Array.isArray(data.low_lube_rows) ? data.low_lube_rows : [];

    setText("icBelowMinTotal", Number(summary.below_min_total || 0));
    setText("icLubeLowCount", Number(summary.lube_below_min_count || 0));
    setText("icOnHand", part ? Number(part.on_hand || 0).toFixed(1) : "-");
    setText("icMinStock", part ? Number(part.min_stock || 0).toFixed(1) : "-");
    if (part && qs("icMinInput")) qs("icMinInput").value = Number(part.min_stock || 0).toFixed(1);
    if (part && qs("icCountedQty")) qs("icCountedQty").value = Number(part.on_hand || 0).toFixed(1);

    const list = qs("icLubeLowList");
    if (list) {
      list.innerHTML = "";
      lubeRows.forEach((r) => {
        list.appendChild(
          item(
            `<b>${r.part_code}</b> — ${Number(r.on_hand || 0).toFixed(1)} on hand <span class='pill red'>LOW</span>` +
            `<br><small>${r.part_name || ""} | Min ${Number(r.min_stock || 0).toFixed(1)} | Short ${Number(r.shortage || 0).toFixed(1)}</small>`
          )
        );
      });
      if (!lubeRows.length) list.appendChild(item("<small>No low lube items right now.</small>"));
    }

    if (part) {
      setText("inventoryControlResult", JSON.stringify(part, null, 2));
      setStatus(`Inventory control ready for ${part.part_code}.`);
    } else {
      setText("inventoryControlResult", JSON.stringify(summary, null, 2));
      setStatus("Inventory control summary ready.");
    }
  } catch (e) {
    setText("inventoryControlResult", String(e.message || e));
    setStatus("Inventory control load failed.");
  }
}

async function saveInventoryPartMinimum() {
  const part_code = (qs("icPartCode")?.value || "").trim();
  const min_stock = Number(qs("icMinInput")?.value || 0);
  if (!part_code) return alert("Enter part code first.");
  if (!Number.isFinite(min_stock) || min_stock < 0) return alert("Minimum must be >= 0.");
  setStatus("Saving part minimum...");
  try {
    const res = await fetchJson(`${API}/api/stock/part-minimum`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ part_code, min_stock }),
    });
    setText("inventoryControlResult", JSON.stringify(res, null, 2));
    await Promise.all([
      loadInventoryControl().catch(() => {}),
      loadStockOnHandPage().catch(() => {}),
      loadLubeStockOnHand().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
    setStatus("Part minimum updated.");
  } catch (e) {
    setText("inventoryControlResult", String(e.message || e));
    setStatus("Failed to update part minimum.");
  }
}

async function submitInventoryCycleCount() {
  const part_code = (qs("icPartCode")?.value || "").trim();
  const counted_qty = Number(qs("icCountedQty")?.value || 0);
  const reason = (qs("icCountReason")?.value || "").trim() || "cycle_count";
  if (!part_code) return alert("Enter part code first.");
  if (!Number.isFinite(counted_qty) || counted_qty < 0) return alert("Counted qty must be >= 0.");

  setStatus("Submitting cycle count...");
  try {
    const res = await fetchJson(`${API}/api/stock/cycle-count`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ part_code, counted_qty, reason }),
    });
    setText("inventoryControlResult", JSON.stringify(res, null, 2));
    await Promise.all([
      loadInventoryControl().catch(() => {}),
      loadApprovalRequests().catch(() => {}),
    ]);
    setStatus(res.no_change ? "Cycle count matched on-hand (no request)." : "Cycle count request submitted.");
  } catch (e) {
    setText("inventoryControlResult", String(e.message || e));
    setStatus("Cycle count submit failed.");
  }
}

async function saveManualLube() {
  const part_code = (qs("mlPart")?.value || "").trim();
  const qtyRequested = Number(qs("mlQty")?.value || 0);
  if (part_code && Number.isFinite(lubeStockMatch.on_hand) && qtyRequested > Number(lubeStockMatch.on_hand)) {
    const warn = `Requested ${qtyRequested.toFixed(1)} exceeds available ${Number(lubeStockMatch.on_hand).toFixed(1)} for ${lubeStockMatch.part_code || part_code}.`;
    setText("mlQtyWarn", warn);
    setStatus("Cannot save lube: insufficient stock.");
    return;
  }
  const mlCc = String(qs("mlCostCenter")?.value || "").trim();
  const payload = {
    asset_code: (qs("mlAsset")?.value || "").trim(),
    log_date: (qs("mlDate")?.value || "").trim() || undefined,
    part_code: part_code || undefined,
    location_code: (qs("mlLocation")?.value || "").trim() || undefined,
    oil_type: (qs("mlType")?.value || "").trim() || undefined,
    quantity: Number(qs("mlQty")?.value || 0),
    cost_center_code: mlCc || undefined,
  };

  setStatus(part_code ? "Issuing lube stock..." : "Saving manual lube entry...");
  try {
    const endpoint = part_code ? `${API}/api/stock/lube-issue` : `${API}/api/stock/lube-log`;
    const res = await fetchJson(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("manualLubeResult", JSON.stringify(res, null, 2));
    setStatus("Manual lube saved.");
    await Promise.all([
      loadDashboard().catch(() => {}),
      loadLubeUsage().catch(() => {}),
    ]);
  } catch (e) {
    setText("manualLubeResult", String(e.message || e));
    setStatus("Manual lube save failed.");
  }
}

async function loadLocations() {
  const showInactive = Boolean(qs("locShowInactive")?.checked);
  setStatus("Loading locations...");
  setSkeleton("locList", 1);
  try {
    const data = await fetchJson(`${API}/api/stock/locations?active=${showInactive ? "0" : "1"}`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const list = qs("locList");
    if (list) {
      list.innerHTML = "";
      rows.forEach((l) => {
        const active = Number(l.active || 0) === 1;
        const tone = active ? "" : "opacity:0.65;";
        list.appendChild(
          item(
            `<div style="display:flex;gap:10px;align-items:center;${tone}">` +
              `<div style="min-width:76px;"><b>${l.location_code}</b></div>` +
              `<div style="flex:1;">${l.location_name || "<span class='muted'>(no name)</span>"}</div>` +
              `<span class="pill ${active ? "blue" : "orange"}">${active ? "ACTIVE" : "INACTIVE"}</span>` +
            `</div>`
          )
        );
      });
      if (!rows.length) list.appendChild(item("<small>No locations found.</small>"));
    }
    setText("locResult", JSON.stringify({ count: rows.length }, null, 2));
    await loadCodePickers().catch(() => {});
    setStatus("Locations ready.");
  } catch (e) {
    setText("locResult", String(e.message || e));
    setStatus("Locations load failed.");
  }
}

async function saveLocation() {
  const location_code = String(qs("locCode")?.value || "").trim().toUpperCase();
  const location_name = String(qs("locName")?.value || "").trim() || undefined;
  const active = String(qs("locActive")?.value || "1") === "1";
  if (!location_code) return alert("Enter location code.");

  setStatus("Saving location...");
  try {
    const res = await fetchJson(`${API}/api/stock/locations`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ location_code, location_name, active }),
    });
    setText("locResult", JSON.stringify(res, null, 2));
    await Promise.all([loadLocations().catch(() => {}), loadCodePickers().catch(() => {})]);
    setStatus("Location saved.");
  } catch (e) {
    setText("locResult", String(e.message || e));
    setStatus("Location save failed.");
  }
}

async function saveStockBin() {
  const location_code = String(qs("sbLocationCode")?.value || "").trim().toUpperCase();
  const bin_code = String(qs("sbBinCode")?.value || "").trim().toUpperCase();
  const bin_name = String(qs("sbBinName")?.value || "").trim() || undefined;
  if (!location_code || !bin_code) return alert("Location and bin code are required.");
  setStatus("Saving stock bin...");
  const res = await fetchJson(`${API}/api/stock/bins`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ location_code, bin_code, bin_name, active: true }),
  });
  setText("sbResult", JSON.stringify(res, null, 2));
  await loadStockBins();
  setStatus("Stock bin saved.");
}

async function loadStockBins() {
  const list = qs("sbBinsList");
  if (!list) return;
  const location_code = String(qs("sbLocationCode")?.value || "").trim().toUpperCase();
  const q = location_code ? `?location_code=${encodeURIComponent(location_code)}&active=1` : "?active=1";
  const data = await fetchJson(`${API}/api/stock/bins${q}`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  list.innerHTML = "";
  rows.forEach((r) => {
    list.appendChild(item(`<b>${r.location_code || "-"}/${r.bin_code || "-"}</b><br><small>${r.bin_name || ""}</small>`));
  });
  if (!rows.length) list.appendChild(item("<small>No bins found.</small>"));
}

async function saveStockMinMax() {
  const part_code = String(qs("sbPartCode")?.value || "").trim();
  const location_code = String(qs("sbLocationCode")?.value || "").trim().toUpperCase();
  const bin_code = String(qs("sbBinCode")?.value || "").trim().toUpperCase() || undefined;
  const min_qty = Number(qs("sbMinQty")?.value || 0);
  const max_qty = Number(qs("sbMaxQty")?.value || 0);
  const reorder_qty_raw = String(qs("sbReorderQty")?.value || "").trim();
  const reorder_qty = reorder_qty_raw === "" ? undefined : Number(reorder_qty_raw);
  if (!part_code || !location_code) return alert("Part code and location are required.");
  if (!Number.isFinite(min_qty) || min_qty < 0 || !Number.isFinite(max_qty) || max_qty < 0) return alert("Min and max must be >= 0.");
  const res = await fetchJson(`${API}/api/stock/min-max`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ part_code, location_code, bin_code, min_qty, max_qty, reorder_qty }),
  });
  setText("sbResult", JSON.stringify(res, null, 2));
  await loadReplenishmentSuggestions();
  setStatus("Min-max policy saved.");
}

async function loadStockDepth() {
  const list = qs("sbDepthList");
  if (!list) return;
  const part_code = String(qs("sbPartCode")?.value || "").trim();
  const location_code = String(qs("sbLocationCode")?.value || "").trim().toUpperCase();
  const bin_code = String(qs("sbBinCode")?.value || "").trim().toUpperCase();
  const q = new URLSearchParams();
  if (part_code) q.set("part_code", part_code);
  if (location_code) q.set("location_code", location_code);
  if (bin_code) q.set("bin_code", bin_code);
  const data = await fetchJson(`${API}/api/stock/depth${q.toString() ? `?${q.toString()}` : ""}`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  list.innerHTML = "";
  rows.slice(0, 120).forEach((r) => {
    list.appendChild(item(
      `<b>${r.part_code || "-"}</b> ${r.location_code || "-"}/${r.bin_code || "-"}`
      + `<br><small>On hand: ${Number(r.on_hand || 0).toFixed(2)} | On order: ${Number(r.on_order || 0).toFixed(2)} | Available: ${Number(r.available || 0).toFixed(2)}</small>`
    ));
  });
  if (!rows.length) list.appendChild(item("<small>No depth rows found.</small>"));
}

let stockReplenishmentRowsCache = [];

async function loadReplenishmentSuggestions() {
  const list = qs("sbReplenishmentList");
  if (!list) return;
  const location_code = String(qs("sbLocationCode")?.value || "").trim().toUpperCase();
  const bin_code = String(qs("sbBinCode")?.value || "").trim().toUpperCase();
  const q = new URLSearchParams();
  if (location_code) q.set("location_code", location_code);
  if (bin_code) q.set("bin_code", bin_code);
  const data = await fetchJson(`${API}/api/stock/replenishment-suggestions${q.toString() ? `?${q.toString()}` : ""}`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  stockReplenishmentRowsCache = rows;
  list.innerHTML = "";
  rows.slice(0, 120).forEach((r) => {
    list.appendChild(item(
      `<b>${r.part_code || "-"}</b> ${r.location_code || "-"}/${r.bin_code || "-"} <span class="pill red">REPLENISH</span>`
      + `<br><small>On hand: ${Number(r.on_hand || 0).toFixed(2)} | Min: ${Number(r.min_qty || 0).toFixed(2)} | Suggest: ${Number(r.suggested_order_qty || 0).toFixed(2)}</small>`
    ));
  });
  if (!rows.length) list.appendChild(item("<small>No replenishment suggestions.</small>"));
}

function exportReplenishmentSuggestionsCsv() {
  const rows = Array.isArray(stockReplenishmentRowsCache) ? stockReplenishmentRowsCache : [];
  if (!rows.length) throw new Error("Load replenishment suggestions first.");
  const header = [
    "part_code",
    "part_name",
    "location_code",
    "bin_code",
    "on_hand",
    "min_qty",
    "max_qty",
    "shortage_qty",
    "suggested_order_qty",
  ];
  const esc = (v) => {
    const s = String(v ?? "");
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g, "\"\"")}"`;
    return s;
  };
  const csv = [header.join(",")].concat(rows.map((r) => [
    r.part_code || "",
    r.part_name || "",
    r.location_code || "",
    r.bin_code || "",
    Number(r.on_hand || 0).toFixed(2),
    Number(r.min_qty || 0).toFixed(2),
    Number(r.max_qty || 0).toFixed(2),
    Number(r.shortage_qty || 0).toFixed(2),
    Number(r.suggested_order_qty || 0).toFixed(2),
  ].map(esc).join(","))).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "replenishment_suggestions.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function createCycleSession() {
  const location_code = String(qs("sbLocationCode")?.value || "").trim().toUpperCase();
  const bin_code = String(qs("sbBinCode")?.value || "").trim().toUpperCase() || undefined;
  if (!location_code) return alert("Location is required to create a cycle session.");
  const res = await fetchJson(`${API}/api/stock/cycle-sessions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      location_code,
      bin_code,
      planned_date: new Date().toISOString().slice(0, 10),
      notes: "Created from stock tab",
    }),
  });
  setText("sbResult", JSON.stringify(res, null, 2));
  await loadCycleSessions();
  setStatus("Cycle session created.");
}

async function loadCycleSessions() {
  const list = qs("sbCycleSessionsList");
  if (!list) return;
  const data = await fetchJson(`${API}/api/stock/cycle-sessions`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  list.innerHTML = "";
  rows.slice(0, 80).forEach((r) => {
    const id = Number(r.id || 0);
    list.appendChild(item(
      `<b>Session #${id}</b> ${r.location_code || "-"}/${r.bin_code || "-"} <span class="pill blue">${r.status || "-"}</span>`
      + `<br><small>Lines: ${Number(r.line_count || 0)} | Variance abs: ${Number(r.variance_abs || 0).toFixed(2)} | Planned: ${r.planned_date || "-"}</small>`
      + `<br><button data-sb-cs-submit="${id}" style="margin-top:8px;">Submit</button>`
      + ` <button data-sb-cs-approve="${id}" style="margin-top:8px;">Approve</button>`
      + ` <button data-sb-cs-countone="${id}" style="margin-top:8px;">Add One Part Count</button>`
    ));
  });
  if (!rows.length) list.appendChild(item("<small>No cycle sessions.</small>"));
}

async function addOnePartCountToSession(sessionId) {
  const id = Number(sessionId || 0);
  if (!id) return;
  const part_code = String(qs("sbPartCode")?.value || "").trim();
  if (!part_code) return alert("Enter/select part code first.");
  const countedRaw = prompt("Counted quantity for this part:", "0");
  if (countedRaw == null) return;
  const counted_qty = Number(countedRaw);
  if (!Number.isFinite(counted_qty) || counted_qty < 0) return alert("Counted qty must be >= 0.");
  const res = await fetchJson(`${API}/api/stock/cycle-sessions/${id}/lines/upsert`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ lines: [{ part_code, counted_qty, reason: "session_count" }] }),
  });
  setText("sbResult", JSON.stringify(res, null, 2));
  await loadCycleSessions();
}

async function submitCycleSession(sessionId) {
  const id = Number(sessionId || 0);
  if (!id) return;
  const res = await fetchJson(`${API}/api/stock/cycle-sessions/${id}/submit`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("sbResult", JSON.stringify(res, null, 2));
  await loadCycleSessions();
  setStatus(`Cycle session #${id} submitted.`);
}

async function approveCycleSession(sessionId) {
  const id = Number(sessionId || 0);
  if (!id) return;
  const ok = confirm(`Approve cycle session #${id} and post adjustment movements?`);
  if (!ok) return;
  const res = await fetchJson(`${API}/api/stock/cycle-sessions/${id}/approve`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("sbResult", JSON.stringify(res, null, 2));
  await Promise.all([loadCycleSessions().catch(() => {}), loadStockDepth().catch(() => {}), loadStockOnHandPage().catch(() => {})]);
  setStatus(`Cycle session #${id} approved.`);
}

let lubeStockMatch = { part_code: null, on_hand: null };

function updateLubeQtyWarning() {
  const warnEl = qs("mlQtyWarn");
  const part_code = (qs("mlPart")?.value || "").trim();
  const qty = Number(qs("mlQty")?.value || 0);
  if (!warnEl) return;
  if (!part_code || !Number.isFinite(lubeStockMatch.on_hand) || !Number.isFinite(qty) || qty <= 0) {
    warnEl.textContent = "";
    return;
  }
  const available = Number(lubeStockMatch.on_hand || 0);
  if (qty > available) {
    warnEl.textContent = `Warning: requested ${qty.toFixed(1)} is above available ${available.toFixed(1)} for ${lubeStockMatch.part_code || part_code}.`;
  } else {
    warnEl.textContent = "";
  }
}

async function loadLubeStockOnHand() {
  const qText = (qs("mlPart")?.value || "").trim() || (qs("mlType")?.value || "").trim();
  const list = qs("lubeStockList");
  if (list) setSkeleton("lubeStockList", 1);
  try {
    // Issue lube uses the dedicated LUBE store by default
    const location_code = "LUBE";
    const q = qText ? `?q=${encodeURIComponent(qText)}&location_code=${encodeURIComponent(location_code)}` : `?location_code=${encodeURIComponent(location_code)}`;
    const data = await fetchJson(`${API}/api/stock/lube-onhand${q}`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const exact = data.exact || (rows.length ? rows[0] : null);

    lubeStockMatch = {
      part_code: exact?.part_code || null,
      on_hand: exact != null ? Number(exact.on_hand || 0) : null,
    };

    const quick = qs("mlLubeQuickLine");
    if (quick) {
      const oils = rows
        .filter((r) => Number(r.on_hand || 0) > 0)
        .slice(0, 8)
        .map((r) => `${r.part_code}: ${Number(r.on_hand || 0).toFixed(0)}`)
        .join(" | ");
      quick.textContent = oils ? `LUBE store available: ${oils}` : "LUBE store available: -";
    }

    setText("mlAvailableQty", exact ? Number(exact.on_hand || 0).toFixed(1) : "-");
    setText("mlAvailablePart", exact ? `${exact.part_code || "-"} ${exact.part_name ? `(${exact.part_name})` : ""}` : "-");
    const partEl = qs("mlPart");
    const typeText = (qs("mlType")?.value || "").trim();
    if (partEl && !String(partEl.value || "").trim() && typeText && exact?.part_code) {
      partEl.value = String(exact.part_code);
    }
    updateManualLubePartDesc();

    if (list) {
      list.innerHTML = "";
      rows.slice(0, 8).forEach((r) => {
        list.appendChild(
          item(
            `<b>${r.part_code}</b> — ${Number(r.on_hand || 0).toFixed(1)} on hand` +
            `${r.below_min ? " <span class='pill red'>LOW</span>" : ""}` +
            `<br><small>${r.part_name || ""} | Min: ${Number(r.min_stock || 0).toFixed(1)}</small>`
          )
        );
      });
      if (!rows.length) list.appendChild(item("<small>No lube stock items found for this filter.</small>"));
    }
    updateLubeQtyWarning();
  } catch (e) {
    lubeStockMatch = { part_code: null, on_hand: null };
    const quick = qs("mlLubeQuickLine");
    if (quick) quick.textContent = "";
    setText("mlAvailableQty", "-");
    setText("mlAvailablePart", "-");
    updateLubeQtyWarning();
    const msg = String(e.message || e);
    if (list) list.innerHTML = `<div class="item"><small>Lube stock load failed: ${msg}</small></div>`;
    setStatus("Lube stock load failed: " + msg);
  }
}

async function setLubeMinimumStock() {
  const minStock = Number(qs("lubeMinStockValue")?.value || 210);
  if (!Number.isFinite(minStock) || minStock < 0) {
    alert("Minimum lube stock must be a valid number >= 0.");
    return;
  }
  setStatus(`Setting minimum lube stock to ${minStock}...`);
  try {
    const res = await fetchJson(`${API}/api/stock/lube-minimums`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ min_stock: minStock }),
    });
    setText("manualLubeResult", JSON.stringify(res, null, 2));
    await Promise.all([
      loadLubeStockOnHand().catch(() => {}),
      loadLubeReorderAlerts().catch(() => {}),
      loadStockOnHandPage().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
    setStatus(`Minimum lube stock set to ${Number(res.min_stock || minStock)} for ${Number(res.updated_count || 0)} item(s).`);
  } catch (e) {
    setText("manualLubeResult", String(e.message || e));
    setStatus("Failed to set minimum lube stock.");
  }
}

async function setSingleLubeMinimum() {
  const part_code = (qs("lubeMinPart")?.value || "").trim();
  const min_stock = Number(qs("lubeMinValue")?.value || 0);
  if (!part_code) return alert("Enter a lube stock number first.");
  if (!Number.isFinite(min_stock) || min_stock < 0) return alert("Minimum must be >= 0.");
  setStatus("Saving lube minimum...");
  try {
    const res = await fetchJson(`${API}/api/stock/part-minimum`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ part_code, min_stock }),
    });
    setText("lubeMinResult", JSON.stringify(res, null, 2));
    await Promise.all([
      loadLubeStockOnHand().catch(() => {}),
      loadLubeReorderAlerts().catch(() => {}),
      loadStockOnHandPage().catch(() => {}),
      loadInventoryControl().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
    setStatus("Lube minimum updated.");
  } catch (e) {
    setText("lubeMinResult", String(e.message || e));
    setStatus("Failed to set lube minimum.");
  }
}

async function loadLubeReorderAlerts() {
  const list = qs("lubeReorderList");
  if (list) setSkeleton("lubeReorderList", 1);
  try {
    const data = await fetchJson(`${API}/api/stock/lube-onhand`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const flagged = rows
      .map((r) => ({
        ...r,
        on_hand: Number(r.on_hand || 0),
        min_stock: Number(r.min_stock || 0),
      }))
      .filter((r) => {
        const min = r.min_stock;
        if (!Number.isFinite(min) || min <= 0) return false;
        const near = min + Math.max(1, min * 0.1);
        return r.on_hand <= near;
      })
      .sort((a, b) => (a.on_hand - a.min_stock) - (b.on_hand - b.min_stock))
      .slice(0, 30);

    if (list) {
      list.innerHTML = "";
      flagged.forEach((r) => {
        const low = r.on_hand <= r.min_stock;
        const pill = low ? "<span class='pill red'>REORDER</span>" : "<span class='pill orange'>NEAR MIN</span>";
        list.appendChild(
          item(
            `<b>${r.part_code}</b> ${pill} — On hand ${r.on_hand.toFixed(1)} | Min ${r.min_stock.toFixed(1)}` +
              `<br><small>${r.part_name || ""}</small>`
          )
        );
      });
      if (!flagged.length) list.appendChild(item("<small>No lube items near/below minimum.</small>"));
    }
  } catch (e) {
    if (list) list.innerHTML = `<div class="item"><small>${String(e.message || e)}</small></div>`;
  }
}

async function loadLubeAnalytics() {
  const months = Number(qs("lubeMonths")?.value || 6);
  setStatus("Loading lube analytics...");
  setSkeleton("lubeAnalyticsList", 2);
  const data = await fetchJson(`${API}/api/dashboard/lube/analytics?months=${encodeURIComponent(months)}`);
  const summary = data.summary || {};
  setText("laTypes", Number(summary.oils || 0));
  setText("laQty", Number(summary.qty_total || 0).toFixed(1));
  setText("laLowRisk", Number(summary.low_risk_count || 0));

  const list = qs("lubeAnalyticsList");
  if (!list) return;
  const trend = Array.isArray(data.trend) ? data.trend : [];
  const monthSet = Array.from(new Set(trend.map((t) => String(t.month || "")))).filter(Boolean);
  const forecast = Array.isArray(data.forecast) ? data.forecast : [];

  list.innerHTML = "";
  forecast.slice(0, 20).forEach((r) => {
    const perMonth = monthSet
      .map((m) => {
        const hit = trend.find((t) => t.month === m && String(t.oil_key || "") === String(r.oil_key || ""));
        return `${m}: ${Number(hit?.qty || 0).toFixed(1)}`;
      })
      .join(" | ");
    list.appendChild(
      item(
        `<b>${r.oil_key || "-"}</b> ${r.low_risk ? "<span class='pill red'>LOW RISK</span>" : "<span class='pill blue'>OK</span>"}` +
          `<br><small>Total ${Number(r.qty_total || 0).toFixed(1)} | Avg/day ${Number(r.avg_daily_use || 0).toFixed(2)} | On hand ${r.on_hand == null ? "-" : Number(r.on_hand).toFixed(1)} | Min ${r.min_stock == null ? "-" : Number(r.min_stock).toFixed(1)} | Days to min ${r.days_to_min == null ? "-" : Number(r.days_to_min).toFixed(1)}</small>` +
          `<br><small>${perMonth || "No monthly trend data."}</small>` +
          `<br><button data-map-oil-key="${String(r.oil_key || "").replace(/"/g, "&quot;")}" data-map-part-code="${String((r.part_code || r.mapped_part_code || "")).replace(/"/g, "&quot;")}" style="margin-top:8px;">Map this</button>`
      )
    );
  });
  if (!forecast.length) list.appendChild(item("<small>No lube analytics found for this period.</small>"));
  setStatus("Lube analytics ready.");
}

async function loadLubeMappings() {
  const list = qs("lubeMapList");
  if (!list) return;
  setSkeleton("lubeMapList", 1);
  const data = await fetchJson(`${API}/api/dashboard/lube/mappings`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  rows.forEach((r) => {
    list.appendChild(
      item(
        `<b>${r.oil_key || "-"}</b> -> ${r.part_code || "-"}` +
        `<br><small>${r.updated_by || "-"} @ ${r.updated_at || "-"}</small>`
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No lube mappings yet.</small>"));
}

async function saveLubeMapping() {
  const oil_key = (qs("lubeMapOilKey")?.value || "").trim();
  const part_code = (qs("lubeMapPartCode")?.value || "").trim();
  if (!oil_key || !part_code) {
    alert("Enter oil key and stock code.");
    return;
  }
  setStatus("Saving lube mapping...");
  const res = await fetchJson(`${API}/api/dashboard/lube/mappings`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ oil_key, part_code }),
  });
  setText("manualLubeResult", JSON.stringify(res, null, 2));
  await Promise.all([
    loadLubeMappings().catch(() => {}),
    loadLubeAnalytics().catch(() => {}),
  ]);
  setStatus("Lube mapping saved.");
}

async function createRequisition() {
  const payload = {
    part_code: (qs("prPartCode")?.value || "").trim(),
    qty_requested: Number(qs("prQty")?.value || 0),
    estimated_value: (qs("prValue")?.value || "").trim() === "" ? undefined : Number(qs("prValue")?.value || 0),
    needed_by_date: (qs("prNeedBy")?.value || "").trim() || undefined,
    bill_to: (qs("prBillTo")?.value || "workshop").trim(),
    request_type: (qs("prRequestType")?.value || "site").trim(),
    supplier_name: (qs("prSupplier")?.value || "").trim() || undefined,
    po_number: (qs("prPo")?.value || "").trim() || undefined,
    notes: (qs("prNotes")?.value || "").trim() || undefined,
  };
  setStatus("Creating requisition...");
  const res = await fetchJson(`${API}/api/procurement/requisitions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await loadRequisitions();
  setStatus("Requisition created.");
}

async function requestRequisitionApproval(id) {
  const reqId = Number(id || 0);
  if (!reqId) return;
  setStatus(`Submitting requisition #${reqId} for approval...`);
  const res = await fetchJson(`${API}/api/procurement/requisitions/${reqId}/request-approval`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await Promise.all([loadRequisitions().catch(() => {}), loadApprovalRequests().catch(() => {})]);
  setStatus(`Requisition #${reqId} sent for approval.`);
}

async function requestRequisitionReceive(id) {
  const reqId = Number(id || 0);
  if (!reqId) return;
  const qtyRaw = prompt("Receive quantity:");
  if (qtyRaw == null) return;
  const qty_receive = Number(qtyRaw);
  if (!Number.isFinite(qty_receive) || qty_receive <= 0) {
    alert("Receive quantity must be > 0.");
    return;
  }
  const reference = prompt("Reference (GRN/Invoice/PO):", `requisition:${reqId}`) || `requisition:${reqId}`;
  setStatus(`Submitting receive request for requisition #${reqId}...`);
  const res = await fetchJson(`${API}/api/procurement/requisitions/${reqId}/request-receive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ qty_receive, reference }),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await Promise.all([loadRequisitions().catch(() => {}), loadApprovalRequests().catch(() => {})]);
  setStatus(`Receive request submitted for requisition #${reqId}.`);
}

async function requestRequisitionReceiveFull(id, qtyOutstanding) {
  const reqId = Number(id || 0);
  const qty_receive = Number(qtyOutstanding || 0);
  if (!reqId) return;
  if (!Number.isFinite(qty_receive) || qty_receive <= 0) {
    alert("No outstanding quantity to receive.");
    return;
  }
  const reference = prompt("Reference (GRN/Invoice/PO):", `requisition:${reqId}:full`) || `requisition:${reqId}:full`;
  setStatus(`Submitting full receive request for requisition #${reqId}...`);
  const res = await fetchJson(`${API}/api/procurement/requisitions/${reqId}/request-receive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ qty_receive, reference }),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await Promise.all([loadRequisitions().catch(() => {}), loadApprovalRequests().catch(() => {})]);
  setStatus(`Full receive request submitted for requisition #${reqId}.`);
}

async function requestRequisitionReceiveHalf(id, qtyOutstanding) {
  const reqId = Number(id || 0);
  const outstanding = Number(qtyOutstanding || 0);
  if (!reqId) return;
  if (!Number.isFinite(outstanding) || outstanding <= 0) {
    alert("No outstanding quantity to receive.");
    return;
  }
  const qty_receive = Number((outstanding * 0.5).toFixed(2));
  const reference = prompt("Reference (GRN/Invoice/PO):", `requisition:${reqId}:half`) || `requisition:${reqId}:half`;
  setStatus(`Submitting 50% receive request for requisition #${reqId}...`);
  const res = await fetchJson(`${API}/api/procurement/requisitions/${reqId}/request-receive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ qty_receive, reference }),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await Promise.all([loadRequisitions().catch(() => {}), loadApprovalRequests().catch(() => {})]);
  setStatus(`50% receive request submitted for requisition #${reqId}.`);
}

async function duplicateRequisitionFromRow(rowJson) {
  let row = null;
  try {
    row = JSON.parse(String(rowJson || "{}"));
  } catch {
    row = null;
  }
  if (!row || !row.part_code) return;

  const qtyDefault = Number(row.qty_requested || 1);
  const qtyRaw = prompt("Duplicate requisition quantity:", String(qtyDefault));
  if (qtyRaw == null) return;
  const qty_requested = Number(qtyRaw);
  if (!Number.isFinite(qty_requested) || qty_requested <= 0) {
    alert("Quantity must be > 0.");
    return;
  }
  const needBy = prompt("Needed by date (YYYY-MM-DD, optional):", String(row.needed_by_date || "")) || "";
  const payload = {
    part_code: String(row.part_code || "").trim(),
    qty_requested,
    needed_by_date: needBy.trim() || undefined,
    supplier_name: String(row.supplier_name || "").trim() || undefined,
    po_number: String(row.po_number || "").trim() || undefined,
    notes: `Duplicate of REQ #${row.id}${row.notes ? ` | ${row.notes}` : ""}`,
  };
  setStatus("Creating duplicate requisition...");
  const res = await fetchJson(`${API}/api/procurement/requisitions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await loadRequisitions();
  setStatus(`Duplicate requisition created from REQ #${row.id}.`);
}

function getProcurementChainConfig() {
  const fallback = {
    tier1Max: 5000,
    tier1Chain: "supervisor",
    tier2Max: 25000,
    tier2Chain: "supervisor,manager",
    tier3Chain: "supervisor,manager,finance,admin",
  };
  try {
    const raw = localStorage.getItem("ironlog.procurement.chainConfig");
    if (!raw) return fallback;
    const parsed = JSON.parse(raw);
    return {
      tier1Max: Number(parsed?.tier1Max || fallback.tier1Max),
      tier1Chain: String(parsed?.tier1Chain || fallback.tier1Chain),
      tier2Max: Number(parsed?.tier2Max || fallback.tier2Max),
      tier2Chain: String(parsed?.tier2Chain || fallback.tier2Chain),
      tier3Chain: String(parsed?.tier3Chain || fallback.tier3Chain),
    };
  } catch {
    return fallback;
  }
}

function setProcurementChainInputsFromConfig() {
  const cfg = getProcurementChainConfig();
  if (qs("prTier1Max")) qs("prTier1Max").value = String(cfg.tier1Max);
  if (qs("prTier1Chain")) qs("prTier1Chain").value = cfg.tier1Chain;
  if (qs("prTier2Max")) qs("prTier2Max").value = String(cfg.tier2Max);
  if (qs("prTier2Chain")) qs("prTier2Chain").value = cfg.tier2Chain;
  if (qs("prTier3Chain")) qs("prTier3Chain").value = cfg.tier3Chain;
}

function saveProcurementChainConfig() {
  const tier1Max = Number(qs("prTier1Max")?.value || 0);
  const tier2Max = Number(qs("prTier2Max")?.value || 0);
  const tier1Chain = String(qs("prTier1Chain")?.value || "").trim();
  const tier2Chain = String(qs("prTier2Chain")?.value || "").trim();
  const tier3Chain = String(qs("prTier3Chain")?.value || "").trim();
  if (!Number.isFinite(tier1Max) || tier1Max < 0) throw new Error("Tier 1 max value is invalid.");
  if (!Number.isFinite(tier2Max) || tier2Max < 0) throw new Error("Tier 2 max value is invalid.");
  if (tier2Max < tier1Max) throw new Error("Tier 2 max must be greater than or equal to Tier 1 max.");
  if (!tier1Chain || !tier2Chain || !tier3Chain) throw new Error("All tier chains are required.");
  localStorage.setItem("ironlog.procurement.chainConfig", JSON.stringify({ tier1Max, tier1Chain, tier2Max, tier2Chain, tier3Chain }));
  setStatus("Approval chain rules saved.");
}

function pickApprovalChainForValue(value) {
  const cfg = getProcurementChainConfig();
  const v = Number(value || 0);
  if (!Number.isFinite(v) || v <= cfg.tier1Max) return cfg.tier1Chain;
  if (v <= cfg.tier2Max) return cfg.tier2Chain;
  return cfg.tier3Chain;
}

function procurementTierMetaByValue(value) {
  const cfg = getProcurementChainConfig();
  const v = Number(value || 0);
  if (!Number.isFinite(v) || v <= cfg.tier1Max) return { label: "TIER 1", cls: "tier-1" };
  if (v <= cfg.tier2Max) return { label: "TIER 2", cls: "tier-2" };
  return { label: "TIER 3", cls: "tier-3" };
}

function updateProcurementChainPreview() {
  const badge = qs("prChainTierBadge");
  const setBadge = (label, cls) => {
    if (!badge) return;
    badge.className = "pill tier-badge";
    if (cls) badge.classList.add(cls);
    badge.textContent = label;
  };
  const override = String(qs("prApproverChain")?.value || "").trim();
  if (override) {
    setBadge("MANUAL", "tier-manual");
    setText("prChainPreview", `Manual override: ${override}`);
    return;
  }
  const cfg = getProcurementChainConfig();
  const valueRaw = String(qs("prValue")?.value || "").trim();
  const valueNum = valueRaw === "" ? 0 : Number(valueRaw);
  const chain = pickApprovalChainForValue(valueNum);
  const displayValue = Number.isFinite(valueNum) ? valueNum.toFixed(2) : "0.00";
  const tier = procurementTierMetaByValue(valueNum);
  setBadge(tier.label, tier.cls);
  setText("prChainPreview", `Value R${displayValue} -> ${chain}`);
}

async function launchApprovalRouteForRequisition(reqId) {
  const row = procurementRowsCache.find((r) => Number(r.id) === Number(reqId));
  const estimatedValue = Number(row?.estimated_value || 0);
  const chainInput = qs("prApproverChain");
  const typedChain = String(chainInput?.value || "").trim();
  const valueBasedChain = pickApprovalChainForValue(estimatedValue);
  const namesRaw = typedChain || valueBasedChain || prompt("Approvers in order (comma separated names):", "approver1,approver2");
  if (!namesRaw) return;
  const approvers = namesRaw
    .split(",")
    .map((x) => String(x || "").trim())
    .filter(Boolean)
    .map((name) => ({ name }));
  await fetchJson(`${API}/api/procurement/requisitions/${reqId}/approvers`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ approvers }),
  });
  if (chainInput && typedChain) {
    chainInput.value = namesRaw;
  }
  const res = await fetchJson(`${API}/api/procurement/requisitions/${reqId}/send-approval`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
}

async function approveCurrentStepForRequisition(reqId) {
  const who = prompt("Approver name:", getSessionUser()) || getSessionUser();
  const comment = prompt("Approval comment (optional):", "") || "";
  const res = await fetchJson(`${API}/api/procurement/requisitions/${reqId}/approve`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ approver_name: who, comment }),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
}

async function advanceRequisitionStage(reqId, currentStatus) {
  const id = Number(reqId || 0);
  const s = String(currentStatus || "").toLowerCase();
  if (!id || !s) return;
  setStatus(`Advancing requisition #${id} from ${s}...`);
  if (s === "draft") {
    const res = await fetchJson(`${API}/api/procurement/requisitions/${id}/finalize`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    setText("procurementResult", JSON.stringify(res, null, 2));
  } else if (s === "finalized") {
    const res = await fetchJson(`${API}/api/procurement/requisitions/${id}/post`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    setText("procurementResult", JSON.stringify(res, null, 2));
  } else if (s === "posted") {
    await launchApprovalRouteForRequisition(id);
  } else if (s === "approval_in_progress") {
    await approveCurrentStepForRequisition(id);
  } else {
    setStatus(`No advance action defined for status '${s}'.`);
    return;
  }
  await loadRequisitions();
  setStatus(`Requisition #${id} advanced.`);
}

function supplyFlowAdvanceButton(reqId, status) {
  const s = String(status || "").toLowerCase();
  const canAdvance = ["draft", "finalized", "posted", "approval_in_progress"].includes(s);
  if (canAdvance) {
    return `<button data-pr-advance-id="${reqId}" data-pr-advance-status="${s}" style="margin-top:8px;">Advance Stage</button>`;
  }
  let tip = "No advance action for this stage.";
  if (s === "approved_all") tip = "PO Ready is complete. Next action is Receive.";
  if (s === "approved") tip = "Receiving actions are available for this stage.";
  if (s === "received") tip = "Requisition fully received.";
  return `<button class="btn-disabled" title="${tip}" disabled style="margin-top:8px;">Advance Stage</button>`;
}

let procurementKpiFilter = "all";
let procurementRowsCache = [];

function setProcurementKpiFilter(filter) {
  procurementKpiFilter = String(filter || "all");
  qs("prKpiAll")?.classList.toggle("pill-active", procurementKpiFilter === "all");
  qs("prKpiApprovedOpen")?.classList.toggle("pill-active", procurementKpiFilter === "approved_open");
  qs("prKpiInFlow")?.classList.toggle("pill-active", procurementKpiFilter === "in_flow");
}

async function loadRequisitions() {
  const list = qs("procurementList");
  if (!list) return;
  const status = (qs("prStatusFilter")?.value || "").trim();
  const tierFilter = (qs("prTierFilter")?.value || "").trim().toLowerCase();
  const q = status ? `?status=${encodeURIComponent(status)}` : "";
  setStatus("Loading requisitions...");
  setSkeleton("procurementList", 2);
  const data = await fetchJson(`${API}/api/procurement/requisitions${q}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  procurementRowsCache = rows;
  setText("prAllCount", rows.length);
  setText(
    "prApprovedOpenCount",
    rows.filter((r) => String(r.status || "").toLowerCase() === "approved" && Number(r.qty_outstanding || 0) > 0).length
  );
  setText(
    "prInFlowCount",
    rows.filter((r) => {
      const s = String(r.status || "").toLowerCase();
      return ["draft", "finalized", "posted", "approval_in_progress", "approved_all"].includes(s);
    }).length
  );
  renderSupplyFlowBoard(rows);
  let displayRows = rows;
  if (procurementKpiFilter === "approved_open") {
    displayRows = rows.filter((r) => String(r.status || "").toLowerCase() === "approved" && Number(r.qty_outstanding || 0) > 0);
  } else if (procurementKpiFilter === "in_flow") {
    displayRows = rows.filter((r) => {
      const s = String(r.status || "").toLowerCase();
      return ["draft", "finalized", "posted", "approval_in_progress", "approved_all"].includes(s);
    });
  } else if (procurementKpiFilter === "receive_set") {
    displayRows = rows.filter((r) => ["approved", "received"].includes(String(r.status || "").toLowerCase()));
  }
  if (tierFilter) {
    displayRows = displayRows.filter((r) => procurementTierMetaByValue(r.estimated_value).cls === tierFilter.replace("_", "-"));
  }
  list.innerHTML = "";
  displayRows.forEach((r) => {
    const s = String(r.status || "").toLowerCase();
    const stageLabel = supplyFlowStageLabel(s);
    const tier = procurementTierMetaByValue(r.estimated_value);
    const tierBadge = `<span class="pill tier-badge ${tier.cls}">${tier.label}</span>`;
    const finalizeBtn = s === "draft" ? `<button data-pr-finalize-id="${r.id}" style="margin-top:8px;">Finalize Request</button>` : "";
    const postBtn = s === "finalized" ? `<button data-pr-post-id="${r.id}" style="margin-top:8px;">Buyer Review & Post</button>` : "";
    const routeBtn = s === "posted" ? `<button data-pr-route-id="${r.id}" style="margin-top:8px;">Launch Approval Route</button>` : "";
    const approveBtn = s === "approval_in_progress" ? `<button data-pr-approve-id="${r.id}" style="margin-top:8px;">Approve Current Step</button>` : "";
    const submitBtn = s === "draft" ? `<button data-pr-submit-id="${r.id}" style="margin-top:8px;">Fast Track to Approval</button>` : "";
    const receiveBtn = ["approved", "approved_all"].includes(s) ? `<button data-pr-receive-id="${r.id}" style="margin-top:8px;">Request Receive</button>` : "";
    const createPoBtn = ["approved_all", "approved", "po_ready", "partially_received", "received"].includes(s)
      ? `<button data-pr-create-po-id="${r.id}" style="margin-top:8px;">Create / Open PO</button>`
      : "";
    const receiveHalfBtn = ["approved", "approved_all"].includes(s) && Number(r.qty_outstanding || 0) > 0
      ? `<button data-pr-receive-half-id="${r.id}" data-pr-outstanding="${Number(r.qty_outstanding || 0)}" style="margin-top:8px;">Receive 50%</button>`
      : "";
    const receiveFullBtn = ["approved", "approved_all"].includes(s) && Number(r.qty_outstanding || 0) > 0
      ? `<button data-pr-receive-full-id="${r.id}" data-pr-outstanding="${Number(r.qty_outstanding || 0)}" style="margin-top:8px;">Receive Full</button>`
      : "";
    const advanceBtn = supplyFlowAdvanceButton(r.id, s);
    const dupPayload = String(JSON.stringify({
      id: r.id,
      part_code: r.part_code,
      qty_requested: r.qty_requested,
      needed_by_date: r.needed_by_date,
      supplier_name: r.supplier_name,
      po_number: r.po_number,
      notes: r.notes,
    })).replace(/"/g, "&quot;");
    list.appendChild(
      item(
        `<b>REQ #${r.id}</b> <span class="pill blue">${stageLabel}</span> - ${r.part_code || "-"} (${r.part_name || "-"})` +
          `<br>${tierBadge}` +
          `<br><small>Site Req No: ${r.site_request_no || "-"} | Bill to: ${r.bill_to || "-"} | Type: ${r.request_type || "-"}</small>` +
          `<br><small>Requested: ${Number(r.qty_requested || 0).toFixed(1)} | Received: ${Number(r.qty_received || 0).toFixed(1)} | Outstanding: ${Number(r.qty_outstanding || 0).toFixed(1)} | Need by: ${r.needed_by_date || "-"}</small>` +
          `<br><small>Req Value: ${r.estimated_value == null ? "-" : Number(r.estimated_value).toFixed(2)}</small>` +
          `<br><small>Supplier: ${r.supplier_name || "-"} | PO: ${r.po_number || "-"}</small>` +
          `<br><small>Requester: ${r.requester || "-"} | ${r.created_at || "-"}</small>` +
          (r.latest_approval_id ? `<br><small>Approval: #${r.latest_approval_id} (${r.latest_approval_status || "-"})</small>` : "") +
          (r.notes ? `<br><small>${r.notes}</small>` : "") +
          `<br>${advanceBtn} ${finalizeBtn} ${postBtn} ${routeBtn} ${approveBtn} ${submitBtn} ${receiveBtn} ${receiveHalfBtn} ${receiveFullBtn} ${createPoBtn} <button data-pr-duplicate="${dupPayload}" style="margin-top:8px;">Duplicate</button> ${r.latest_approval_id ? `<button data-pr-open-approval-id="${r.latest_approval_id}" style="margin-top:8px;">Open Approval</button>` : ""}`
      )
    );
  });
  if (!displayRows.length) list.appendChild(item("<small>No requisitions found.</small>"));
  setStatus("Requisitions ready.");
}

function supplyFlowStageLabel(status) {
  const s = String(status || "").toLowerCase();
  if (s === "draft") return "Plan";
  if (s === "finalized") return "Review";
  if (s === "posted") return "Approval Route";
  if (s === "approval_in_progress") return "Approvals";
  if (s === "approved_all") return "PO Ready";
  if (s === "approved" || s === "received") return "Receive";
  return s || "Unknown";
}

function renderSupplyFlowBoard(rows) {
  const lanes = {
    plan: qs("sfPlan"),
    review: qs("sfReview"),
    route: qs("sfRoute"),
    approve: qs("sfApprove"),
    po_ready: qs("sfPoReady"),
    receive: qs("sfReceive"),
  };
  const laneCounts = {
    plan: 0,
    review: 0,
    route: 0,
    approve: 0,
    po_ready: 0,
    receive: 0,
  };
  Object.values(lanes).forEach((el) => {
    if (el) el.innerHTML = "";
  });

  const add = (laneKey, html) => {
    const lane = lanes[laneKey];
    if (lane) lane.appendChild(item(html));
  };

  (Array.isArray(rows) ? rows : []).slice(0, 80).forEach((r) => {
    const s = String(r.status || "").toLowerCase();
    const tier = procurementTierMetaByValue(r.estimated_value);
    const snippet =
      `<b>#${r.id}</b> ${r.part_code || "-"} <span class="pill tier-badge ${tier.cls}">${tier.label}</span><br><small>Out: ${Number(r.qty_outstanding || 0).toFixed(1)} | ${r.bill_to || "-"}</small>` +
      `<br>${supplyFlowAdvanceButton(r.id, s)}`;
    if (s === "draft") {
      laneCounts.plan += 1;
      add("plan", snippet);
    } else if (s === "finalized") {
      laneCounts.review += 1;
      add("review", snippet);
    } else if (s === "posted") {
      laneCounts.route += 1;
      add("route", snippet);
    } else if (s === "approval_in_progress") {
      laneCounts.approve += 1;
      add("approve", snippet);
    } else if (s === "approved_all") {
      laneCounts.po_ready += 1;
      add("po_ready", snippet);
    } else {
      laneCounts.receive += 1;
      add("receive", snippet);
    }
  });

  Object.entries(lanes).forEach(([k, el]) => {
    if (el && !el.children.length) {
      el.appendChild(item(`<small>No requests in ${k.replace("_", " ")}.</small>`));
    }
  });

  setText("sfCountPlan", String(laneCounts.plan));
  setText("sfCountReview", String(laneCounts.review));
  setText("sfCountRoute", String(laneCounts.route));
  setText("sfCountApprove", String(laneCounts.approve));
  setText("sfCountPoReady", String(laneCounts.po_ready));
  setText("sfCountReceive", String(laneCounts.receive));
}

let procurementLastJournalBatchId = "";

async function loadPurchaseOrders() {
  const list = qs("prPoList");
  if (!list) return;
  const status = String(qs("prPoStatusFilter")?.value || "").trim();
  const q = status ? `?status=${encodeURIComponent(status)}` : "";
  setStatus("Loading purchase orders...");
  list.innerHTML = "";
  const data = await fetchJson(`${API}/api/procurement/purchase-orders${q}`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  rows.forEach((r) => {
    const poId = Number(r.id || 0);
    list.appendChild(item(
      `<b>PO #${poId}</b> <span class="pill blue">${String(r.status || "-")}</span> - ${r.po_number || "-"}`
      + `<br><small>Supplier: ${r.supplier_name || r.supplier_code || "-"}</small>`
      + `<br><small>Subtotal: ${Number(r.subtotal || 0).toFixed(2)} ${r.currency || "USD"} | Req: ${r.requisition_id || "-"}</small>`
      + `<br><button data-pr-po-open="${poId}" style="margin-top:8px;">Open PO</button>`
      + ` <button data-pr-po-approve="${poId}" style="margin-top:8px;">Approve</button>`
      + ` <button data-pr-po-send="${poId}" style="margin-top:8px;">Mark Sent</button>`
    ));
  });
  if (!rows.length) list.appendChild(item("<small>No purchase orders found.</small>"));
  setStatus("Purchase orders ready.");
}

async function createPoFromRequisition(reqId) {
  const id = Number(reqId || 0);
  if (!id) return;
  setStatus(`Creating PO from requisition #${id}...`);
  const res = await fetchJson(`${API}/api/procurement/requisitions/${id}/create-po`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  if (res?.po_id) {
    const poEl = qs("prActivePoId");
    if (poEl) poEl.value = String(res.po_id);
  }
  await Promise.all([loadRequisitions().catch(() => {}), loadPurchaseOrders().catch(() => {})]);
  setStatus(`PO created from requisition #${id}.`);
}

async function openPurchaseOrder(poId) {
  const id = Number(poId || 0);
  if (!id) return;
  const res = await fetchJson(`${API}/api/procurement/purchase-orders/${id}/detail`);
  const po = res?.po || {};
  const lines = Array.isArray(res?.lines) ? res.lines : [];
  const receipts = Array.isArray(res?.receipts) ? res.receipts : [];
  setText("procurementResult", JSON.stringify(res, null, 2));
  const poEl = qs("prActivePoId");
  if (poEl) poEl.value = String(id);
  const recJson = qs("prReceiveLinesJson");
  if (recJson) {
    const sample = lines
      .filter((l) => Number(l.quantity_ordered || 0) > Number(l.quantity_received || 0))
      .slice(0, 3)
      .map((l) => ({ po_line_id: Number(l.id), quantity_received: Number((Number(l.quantity_ordered || 0) - Number(l.quantity_received || 0)).toFixed(2)) }));
    recJson.value = JSON.stringify(sample.length ? sample : [{ po_line_id: Number(lines[0]?.id || 0), quantity_received: 1 }]);
  }
  const invJson = qs("prInvoiceLinesJson");
  if (invJson) {
    const sample = lines.slice(0, 3).map((l) => ({
      po_line_id: Number(l.id || 0),
      quantity_invoiced: Number((Number(l.quantity_received || l.quantity_ordered || 0)).toFixed(2)),
      unit_price: Number(l.unit_price || 0),
    }));
    invJson.value = JSON.stringify(sample.length ? sample : [{ po_line_id: Number(lines[0]?.id || 0), quantity_invoiced: 1, unit_price: 0 }]);
  }
  setStatus(`PO ${po.po_number || id} loaded (${lines.length} lines, ${receipts.length} receipts).`);
}

async function approvePurchaseOrder(poId) {
  const id = Number(poId || 0);
  if (!id) return;
  const res = await fetchJson(`${API}/api/procurement/purchase-orders/${id}/approve`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await loadPurchaseOrders();
  setStatus(`PO #${id} approved.`);
}

async function sendPurchaseOrder(poId) {
  const id = Number(poId || 0);
  if (!id) return;
  const res = await fetchJson(`${API}/api/procurement/purchase-orders/${id}/send`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await loadPurchaseOrders();
  setStatus(`PO #${id} marked sent.`);
}

async function postPoReceipt() {
  const poId = Number(qs("prActivePoId")?.value || 0);
  if (!poId) throw new Error("PO ID is required.");
  const location_code = String(qs("prReceiveLocationCode")?.value || "MAIN").trim().toUpperCase() || "MAIN";
  const raw = String(qs("prReceiveLinesJson")?.value || "").trim();
  if (!raw) throw new Error("Receipt lines JSON is required.");
  let lines = [];
  try {
    lines = JSON.parse(raw);
  } catch {
    throw new Error("Receipt lines JSON is invalid.");
  }
  const payload = { location_code, lines };
  const res = await fetchJson(`${API}/api/procurement/purchase-orders/${poId}/receive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await Promise.all([loadPurchaseOrders().catch(() => {}), loadRequisitions().catch(() => {})]);
  setStatus(`Receipt posted for PO #${poId}.`);
}

async function capturePoInvoice() {
  const poId = Number(qs("prActivePoId")?.value || 0);
  if (!poId) throw new Error("PO ID is required.");
  const invoice_number = String(qs("prInvoiceNumber")?.value || "").trim();
  if (!invoice_number) throw new Error("Invoice number is required.");
  const raw = String(qs("prInvoiceLinesJson")?.value || "").trim();
  if (!raw) throw new Error("Invoice lines JSON is required.");
  let lines = [];
  try {
    lines = JSON.parse(raw);
  } catch {
    throw new Error("Invoice lines JSON is invalid.");
  }
  const res = await fetchJson(`${API}/api/procurement/purchase-orders/${poId}/invoices`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ invoice_number, lines }),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  setStatus(`Invoice captured for PO #${poId}.`);
}

async function runPoThreeWayMatch() {
  const poId = Number(qs("prActivePoId")?.value || 0);
  if (!poId) throw new Error("PO ID is required.");
  const res = await fetchJson(`${API}/api/procurement/purchase-orders/${poId}/three-way-match`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ quantity_tolerance: 0, price_tolerance_pct: 5, total_tolerance: 1 }),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await loadProcurementExceptions();
  setStatus(`3-way match complete for PO #${poId}.`);
}

async function loadProcurementExceptions() {
  const list = qs("prExceptionsList");
  if (!list) return;
  const status = String(qs("prExceptionStatus")?.value || "open").trim().toLowerCase();
  list.innerHTML = "";
  const data = await fetchJson(`${API}/api/procurement/exceptions?status=${encodeURIComponent(status)}`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  rows.forEach((r) => {
    list.appendChild(item(
      `<b>EX #${Number(r.id || 0)}</b> <span class="pill ${String(r.severity || "").toLowerCase() === "high" ? "red" : "orange"}">${r.severity || "-"}</span> <span class="pill blue">${r.status || "-"}</span>`
      + `<br><small>PO: ${r.po_number || r.po_id || "-"} | Type: ${r.exception_type || "-"}</small>`
      + `<br><small>${r.details ? JSON.stringify(r.details) : "-"}</small>`
      + (String(r.status || "").toLowerCase() === "open" ? `<br><button data-pr-ex-resolve="${Number(r.id || 0)}" style="margin-top:8px;">Resolve</button>` : "")
    ));
  });
  if (!rows.length) list.appendChild(item("<small>No exceptions found.</small>"));
}

async function resolveProcurementException(exId) {
  const id = Number(exId || 0);
  if (!id) return;
  const res = await fetchJson(`${API}/api/procurement/exceptions/${id}/resolve`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await loadProcurementExceptions();
  setStatus(`Exception #${id} resolved.`);
}

async function buildProcurementJournals() {
  const start = String(qs("prJournalStart")?.value || "").trim();
  const end = String(qs("prJournalEnd")?.value || "").trim();
  if (!start || !end) throw new Error("Journal start and end dates are required.");
  const batch_id = String(qs("prJournalBatchId")?.value || "").trim() || undefined;
  const default_cost_center_code = String(qs("prJournalCc")?.value || "").trim() || undefined;
  const res = await fetchJson(`${API}/api/procurement/journals/build`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ start, end, batch_id, default_cost_center_code }),
  });
  procurementLastJournalBatchId = String(res?.batch_id || "");
  setText("prJournalLastBatch", procurementLastJournalBatchId || "-");
  setText("procurementResult", JSON.stringify(res, null, 2));
  setStatus(`Journal batch built: ${procurementLastJournalBatchId || "-"}.`);
}

function currentProcurementJournalBatch() {
  const typed = String(qs("prJournalBatchId")?.value || "").trim();
  return typed || procurementLastJournalBatchId;
}

function exportProcurementJournalsCsv() {
  const batch = currentProcurementJournalBatch();
  if (!batch) throw new Error("Build journals first or enter batch ID.");
  window.open(`${API}/api/procurement/journals/export.csv?batch_id=${encodeURIComponent(batch)}`, "_blank");
}

function exportProcurementJournalsXlsx() {
  const batch = currentProcurementJournalBatch();
  if (!batch) throw new Error("Build journals first or enter batch ID.");
  window.open(`${API}/api/procurement/journals/export.xlsx?batch_id=${encodeURIComponent(batch)}`, "_blank");
}

function getSiteOpsFrom() {
  return (qs("opFrom")?.value || "").trim();
}

function getSiteOpsTo() {
  return (qs("opTo")?.value || "").trim();
}

async function loadSiteZones() {
  const data = await fetchJson(`${API}/api/operations/site/zones`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  const zoneSelect = qs("opSiteZone");
  const zoneList = qs("siteZoneList");
  if (zoneSelect) {
    zoneSelect.innerHTML = `<option value="">Select zone</option>`;
    rows
      .filter((r) => Number(r.active || 0) === 1)
      .forEach((r) => {
        const opt = document.createElement("option");
        opt.value = String(r.id || "");
        opt.textContent = `${r.name || "Zone"} (#${r.id})`;
        zoneSelect.appendChild(opt);
      });
  }
  if (zoneList) {
    zoneList.innerHTML = "";
    if (!rows.length) {
      zoneList.appendChild(item("<small>No zones configured.</small>"));
    } else {
      rows.forEach((r) => {
        zoneList.appendChild(item(`<b>#${r.id}</b> ${r.name || "-"} <span class="pill ${Number(r.active || 0) ? "blue" : "orange"}">${Number(r.active || 0) ? "active" : "inactive"}</span>`));
      });
    }
  }
}

async function saveSiteZone() {
  const name = String(qs("opZoneName")?.value || "").trim();
  if (!name) {
    alert("Zone name is required.");
    return;
  }
  const active = Boolean(qs("opZoneActive")?.checked);
  const res = await fetchJson(`${API}/api/operations/site/zones`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, active }),
  });
  setText("siteDailyResult", JSON.stringify(res, null, 2));
  await loadSiteZones();
  setStatus("Site zone saved.");
}

async function saveSiteDailyEntry() {
  const payload = {
    op_date: (qs("opSiteDate")?.value || "").trim(),
    material_type: (qs("opSiteMaterial")?.value || "").trim(),
    zone_id: (qs("opSiteZone")?.value || "").trim() === "" ? undefined : Number(qs("opSiteZone")?.value || 0),
    planned_tonnage: (qs("opSitePlanned")?.value || "").trim() === "" ? undefined : Number(qs("opSitePlanned")?.value || 0),
    actual_tonnage: (qs("opSiteActual")?.value || "").trim() === "" ? undefined : Number(qs("opSiteActual")?.value || 0),
    loads_count: (qs("opSiteLoads")?.value || "").trim() === "" ? undefined : Number(qs("opSiteLoads")?.value || 0),
    avg_cycle_time: (qs("opSiteCycle")?.value || "").trim() === "" ? undefined : Number(qs("opSiteCycle")?.value || 0),
    operator_name: (qs("opSiteOperator")?.value || "").trim() || undefined,
    notes: (qs("opSiteNotes")?.value || "").trim() || undefined,
  };
  if (!payload.op_date || !payload.material_type) {
    alert("Date and material type are required.");
    return;
  }
  const res = await fetchJson(`${API}/api/operations/site/daily`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("siteDailyResult", JSON.stringify(res, null, 2));
  if (qs("opSiteDailyId") && res?.id) qs("opSiteDailyId").value = String(res.id);
  await loadSiteDailyEntries();
  await loadSiteDashboard();
  setStatus("Site daily entry saved.");
}

async function loadSiteDailyEntries() {
  const list = qs("siteDailyList");
  if (!list) return;
  const from = getSiteOpsFrom();
  const to = getSiteOpsTo();
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (to) q.set("to", to);
  const data = await fetchJson(`${API}/api/operations/site/daily${q.toString() ? `?${q.toString()}` : ""}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  if (!rows.length) {
    list.appendChild(item("<small>No site daily entries found.</small>"));
    return;
  }
  rows.forEach((r) => {
    list.appendChild(
      item(
        `<b>#${r.id}</b> ${r.op_date || "-"} <span class="pill blue">${String(r.shift || "-").toUpperCase()}</span> <span class="pill">${r.material_type || "-"}</span>` +
          `<br><small>Zone: ${r.zone_name || "-"} | Planned: ${Number(r.planned_tonnage || 0).toFixed(2)} | Actual: ${Number(r.actual_tonnage || 0).toFixed(2)} | Loads: ${Number(r.loads_count || 0)}</small>` +
          `<br><small>Operator: ${r.operator_name || "-"}${r.notes ? ` | ${r.notes}` : ""}</small>`
      )
    );
  });
}

async function saveSiteEquipmentUsage() {
  const dailyId = Number(qs("opSiteDailyId")?.value || 0);
  const assetId = Number(qs("opSiteEqAssetId")?.value || 0);
  const role = String(qs("opSiteEqRole")?.value || "").trim();
  const hours = (qs("opSiteEqHours")?.value || "").trim() === "" ? undefined : Number(qs("opSiteEqHours")?.value || 0);
  if (!dailyId || !assetId || !role) {
    alert("Daily ID, Asset ID, and role are required.");
    return;
  }
  const res = await fetchJson(`${API}/api/operations/site/daily/${dailyId}/equipment`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ asset_id: assetId, role, hours_used: hours }),
  });
  setText("siteDailyResult", JSON.stringify(res, null, 2));
  await loadSiteEquipmentUsage();
  setStatus("Equipment linked to production.");
}

async function loadSiteEquipmentUsage() {
  const dailyId = Number(qs("opSiteDailyId")?.value || 0);
  const list = qs("siteEquipmentUsageList");
  if (!list) return;
  if (!dailyId) {
    list.innerHTML = "";
    list.appendChild(item("<small>Select a daily entry ID to view equipment usage.</small>"));
    return;
  }
  const data = await fetchJson(`${API}/api/operations/site/daily/${dailyId}/equipment`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  if (!rows.length) {
    list.appendChild(item("<small>No equipment usage linked yet.</small>"));
    return;
  }
  rows.forEach((r) => {
    list.appendChild(item(`<b>#${r.id}</b> Asset ${r.asset_code || r.asset_id} (${r.asset_name || "-"}) | Role: ${r.role || "-"} | Hours: ${Number(r.hours_used || 0).toFixed(2)}`));
  });
}

async function saveSiteTarget() {
  const payload = {
    target_date: (qs("opTargetDate")?.value || "").trim(),
    material_type: (qs("opTargetMaterial")?.value || "").trim(),
    target_tonnage: (qs("opTargetTonnage")?.value || "").trim() === "" ? undefined : Number(qs("opTargetTonnage")?.value || 0),
  };
  if (!payload.target_date || !payload.material_type || payload.target_tonnage == null) {
    alert("Target date, material, and tonnage are required.");
    return;
  }
  const res = await fetchJson(`${API}/api/operations/site/targets`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("siteDailyResult", JSON.stringify(res, null, 2));
  await loadSiteTargets();
  await loadSiteDashboard();
  setStatus("Site target saved.");
}

async function loadSiteTargets() {
  const list = qs("siteTargetList");
  if (!list) return;
  const from = getSiteOpsFrom();
  const to = getSiteOpsTo();
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (to) q.set("to", to);
  const data = await fetchJson(`${API}/api/operations/site/targets${q.toString() ? `?${q.toString()}` : ""}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  if (!rows.length) {
    list.appendChild(item("<small>No targets found.</small>"));
    return;
  }
  rows.forEach((r) => {
    list.appendChild(item(`<b>${r.target_date}</b> ${r.material_type || "-"} <span class="pill blue">${Number(r.target_tonnage || 0).toFixed(2)} t</span>`));
  });
}

async function saveSiteDelay() {
  const payload = {
    delay_date: (qs("opDelayDate")?.value || "").trim(),
    delay_type: (qs("opDelayType")?.value || "").trim(),
    hours_lost: (qs("opDelayHours")?.value || "").trim() === "" ? undefined : Number(qs("opDelayHours")?.value || 0),
    impact_tonnage: (qs("opDelayImpact")?.value || "").trim() === "" ? undefined : Number(qs("opDelayImpact")?.value || 0),
    notes: (qs("opDelayNotes")?.value || "").trim() || undefined,
  };
  if (!payload.delay_date || !payload.delay_type || payload.hours_lost == null) {
    alert("Delay date, type and hours lost are required.");
    return;
  }
  const res = await fetchJson(`${API}/api/operations/site/delays`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("siteDailyResult", JSON.stringify(res, null, 2));
  await loadSiteDelays();
  await loadSiteDashboard();
  setStatus("Operational delay saved.");
}

async function loadSiteDelays() {
  const list = qs("siteDelayList");
  if (!list) return;
  const from = getSiteOpsFrom();
  const to = getSiteOpsTo();
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (to) q.set("to", to);
  const data = await fetchJson(`${API}/api/operations/site/delays${q.toString() ? `?${q.toString()}` : ""}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  if (!rows.length) {
    list.appendChild(item("<small>No operational delays found.</small>"));
    return;
  }
  rows.forEach((r) => {
    list.appendChild(item(`<b>${r.delay_date}</b> <span class="pill orange">${r.delay_type}</span> | Hours lost: ${Number(r.hours_lost || 0).toFixed(2)} | Impact: ${Number(r.impact_tonnage || 0).toFixed(2)}t${r.notes ? `<br><small>${r.notes}</small>` : ""}`));
  });
}

async function loadSiteDashboard() {
  const date = (qs("opSiteDashDate")?.value || qs("opSiteDate")?.value || "").trim();
  if (!date) return;
  const data = await fetchJson(`${API}/api/operations/site/dashboard?date=${encodeURIComponent(date)}`);
  const today = data?.today || {};
  const week = data?.week || {};
  const losses = data?.losses || {};
  setText("opSiteKpiTodayTons", Number(today.total_tons_produced || 0).toFixed(2));
  setText("opSiteKpiAchieved", Number(today.achieved_pct || 0).toFixed(1));
  setText("opSiteKpiLoads", String(Number(today.loads_moved || 0)));
  setText("opSiteKpiZones", String(Number(today.active_zones || 0)));
  setText("opSiteKpiWeekTotal", Number(week.total_production || 0).toFixed(2));
  setText("opSiteKpiBreakdownLoss", Number(losses.breakdown_hours || 0).toFixed(2));
  setText("opSiteKpiOpsLoss", Number(losses.operational_delay_hours || 0).toFixed(2));

  const list = qs("siteDashboardList");
  if (list) {
    list.innerHTML = "";
    const best = week.best_day ? `${week.best_day.date} (${Number(week.best_day.tons || 0).toFixed(2)}t)` : "-";
    const worst = week.worst_day ? `${week.worst_day.date} (${Number(week.worst_day.tons || 0).toFixed(2)}t)` : "-";
    list.appendChild(item(`<b>Best day:</b> ${best}`));
    list.appendChild(item(`<b>Worst day:</b> ${worst}`));
    list.appendChild(item(`<b>Today target:</b> ${Number(today.target_tonnage || 0).toFixed(2)}t | <b>Shortfall:</b> ${Math.max(0, Number(today.target_tonnage || 0) - Number(today.total_tons_produced || 0)).toFixed(2)}t`));
  }
}

async function saveOperationEntry() {
  const payload = {
    op_date: (qs("opDate")?.value || "").trim() || undefined,
    tonnes_moved: (qs("opTonnesMoved")?.value || "").trim() === "" ? undefined : Number(qs("opTonnesMoved")?.value || 0),
    product_type: (qs("opProductType")?.value || "").trim() || undefined,
    product_produced: (qs("opProductProduced")?.value || "").trim() === "" ? undefined : Number(qs("opProductProduced")?.value || 0),
    trucks_loaded: (qs("opTrucksLoaded")?.value || "").trim() === "" ? undefined : Number(qs("opTrucksLoaded")?.value || 0),
    loads_count: (qs("opLoadsCount")?.value || "").trim() === "" ? undefined : Number(qs("opLoadsCount")?.value || 0),
    crusher_feed_tonnes: (qs("opCrusherFeedTonnes")?.value || "").trim() === "" ? undefined : Number(qs("opCrusherFeedTonnes")?.value || 0),
    crusher_output_tonnes: (qs("opCrusherOutputTonnes")?.value || "").trim() === "" ? undefined : Number(qs("opCrusherOutputTonnes")?.value || 0),
    crusher_hours: (qs("opCrusherHours")?.value || "").trim() === "" ? undefined : Number(qs("opCrusherHours")?.value || 0),
    crusher_downtime_hours: (qs("opCrusherDowntime")?.value || "").trim() === "" ? undefined : Number(qs("opCrusherDowntime")?.value || 0),
    weighbridge_amount: (qs("opWeighbridgeAmount")?.value || "").trim() === "" ? undefined : Number(qs("opWeighbridgeAmount")?.value || 0),
    trucks_delivered: (qs("opTrucksDelivered")?.value || "").trim() === "" ? undefined : Number(qs("opTrucksDelivered")?.value || 0),
    product_delivered: (qs("opProductDelivered")?.value || "").trim() === "" ? undefined : Number(qs("opProductDelivered")?.value || 0),
    client_delivered_to: (qs("opClientDeliveredTo")?.value || "").trim() || undefined,
    notes: (qs("opNotes")?.value || "").trim() || undefined,
  };
  setStatus("Saving operations entry...");
  const res = await fetchJson(`${API}/api/operations`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("operationsResult", JSON.stringify(res, null, 2));
  await loadOperations();
  setStatus("Operations entry saved.");
}

async function loadOperationsClosingForDate(opDate) {
  const date = String(opDate || "").trim();
  if (!date) return;
  const data = await fetchJson(`${API}/api/operations/closing/${encodeURIComponent(date)}`);
  const row = data?.row || null;
  if (qs("opCloseShift")) qs("opCloseShift").value = row?.shift_name || "";
  if (qs("opCloseSupervisor")) qs("opCloseSupervisor").value = row?.supervisor_name || "";
  if (qs("opCloseVariance")) qs("opCloseVariance").value = row?.variance_note || "";
  if (qs("opCloseChkWeighbridge")) qs("opCloseChkWeighbridge").checked = Boolean(row?.checklist_weighbridge_reconciled);
  if (qs("opCloseChkTrucks")) qs("opCloseChkTrucks").checked = Boolean(row?.checklist_trucks_reconciled);
  if (qs("opCloseChkClient")) qs("opCloseChkClient").checked = Boolean(row?.checklist_client_confirmed);
  const status = String(row?.status || "open").toUpperCase();
  setText("opCloseStatusPill", `Day Status: ${status}`);
  if (row) setText("operationsClosingResult", JSON.stringify(row, null, 2));
  await loadOperationsClosingHistoryForDate(date);
}

async function loadOperationsClosingHistoryForDate(opDate) {
  const date = String(opDate || "").trim();
  const list = qs("operationsClosingHistory");
  if (!date || !list) return;
  const data = await fetchJson(`${API}/api/operations/closing/${encodeURIComponent(date)}/history`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  if (!rows.length) {
    list.appendChild(item("<small>No closure history for this date.</small>"));
    return;
  }
  rows.forEach((r) => {
    const p = r.payload && typeof r.payload === "object" ? r.payload : {};
    const reason = p.reason ? ` | Reason: ${p.reason}` : "";
    const supervisor = p.supervisor_name ? ` | Supervisor: ${p.supervisor_name}` : "";
    list.appendChild(
      item(
        `<b>${r.action}</b> <span class="pill blue">${r.role || "-"}</span>` +
          `<br><small>User: ${r.username || "-"} | At: ${r.created_at || "-"}</small>` +
          `<br><small>Status: ${p.status || "-"}${supervisor}${reason}</small>`
      )
    );
  });
}

async function saveOperationsClosing(closeDay) {
  const op_date = (qs("opDate")?.value || "").trim();
  if (!op_date) {
    alert("Select an operations date first.");
    return;
  }
  const payload = {
    op_date,
    shift_name: (qs("opCloseShift")?.value || "").trim() || undefined,
    supervisor_name: (qs("opCloseSupervisor")?.value || "").trim() || undefined,
    variance_note: (qs("opCloseVariance")?.value || "").trim() || undefined,
    checklist_weighbridge_reconciled: Boolean(qs("opCloseChkWeighbridge")?.checked),
    checklist_trucks_reconciled: Boolean(qs("opCloseChkTrucks")?.checked),
    checklist_client_confirmed: Boolean(qs("opCloseChkClient")?.checked),
    close_day: Boolean(closeDay),
  };
  if (closeDay) {
    if (!payload.supervisor_name) {
      alert("Supervisor sign-off name is required to close day.");
      return;
    }
    if (!(payload.checklist_weighbridge_reconciled && payload.checklist_trucks_reconciled && payload.checklist_client_confirmed)) {
      alert("Complete all checklist items before closing day.");
      return;
    }
  }
  setStatus(closeDay ? "Closing operations day..." : "Saving operations closing draft...");
  const res = await fetchJson(`${API}/api/operations/closing`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("operationsClosingResult", JSON.stringify(res?.row || res, null, 2));
  await loadOperationsClosingForDate(op_date);
  setStatus(closeDay ? "Operations day closed." : "Operations closing draft saved.");
}

async function reopenOperationsDay() {
  const roles = getSessionRoles();
  if (!roles.some((r) => ["admin", "supervisor"].includes(r))) {
    alert("Only admin or supervisor can re-open a closed day.");
    return;
  }
  const op_date = (qs("opDate")?.value || "").trim();
  if (!op_date) {
    alert("Select an operations date first.");
    return;
  }
  const reopen_reason = (qs("opReopenReason")?.value || "").trim();
  if (!reopen_reason) {
    alert("Re-open reason is required.");
    return;
  }
  const payload = {
    op_date,
    reopen_day: true,
    reopen_reason,
    close_day: false,
    shift_name: (qs("opCloseShift")?.value || "").trim() || undefined,
    supervisor_name: (qs("opCloseSupervisor")?.value || "").trim() || undefined,
    variance_note: (qs("opCloseVariance")?.value || "").trim() || undefined,
    checklist_weighbridge_reconciled: Boolean(qs("opCloseChkWeighbridge")?.checked),
    checklist_trucks_reconciled: Boolean(qs("opCloseChkTrucks")?.checked),
    checklist_client_confirmed: Boolean(qs("opCloseChkClient")?.checked),
  };
  setStatus("Re-opening operations day...");
  const res = await fetchJson(`${API}/api/operations/closing`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("operationsClosingResult", JSON.stringify(res?.row || res, null, 2));
  await loadOperationsClosingForDate(op_date);
  setStatus("Operations day re-opened.");
}

async function loadOperations() {
  const list = qs("operationsList");
  if (!list) return;
  const from = (qs("opFrom")?.value || "").trim();
  const to = (qs("opTo")?.value || "").trim();
  const params = [];
  if (from) params.push(`from=${encodeURIComponent(from)}`);
  if (to) params.push(`to=${encodeURIComponent(to)}`);
  const q = params.length ? `?${params.join("&")}` : "";
  setStatus("Loading operations...");
  setSkeleton("operationsList", 2);
  const data = await fetchJson(`${API}/api/operations${q}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  let tonnes = 0;
  let produced = 0;
  let loaded = 0;
  let loadCycles = 0;
  let crusherFeed = 0;
  let crusherOutput = 0;
  const clientTotalsDelivered = new Map();
  const clientTotalsTrucks = new Map();
  const clientTotalsTonnes = new Map();
  rows.forEach((r) => {
    tonnes += Number(r.tonnes_moved || 0);
    produced += Number(r.product_produced || 0);
    loaded += Number(r.trucks_loaded || 0);
    loadCycles += Number(r.loads_count || 0);
    crusherFeed += Number(r.crusher_feed_tonnes || 0);
    crusherOutput += Number(r.crusher_output_tonnes || 0);
    const client = String(r.client_delivered_to || "").trim() || "Unspecified";
    const deliveredQty = Number(r.product_delivered || 0);
    const trucksQty = Number(r.trucks_delivered || 0);
    const tonnesQty = Number(r.tonnes_moved || 0);
    clientTotalsDelivered.set(client, Number(clientTotalsDelivered.get(client) || 0) + deliveredQty);
    clientTotalsTrucks.set(client, Number(clientTotalsTrucks.get(client) || 0) + trucksQty);
    clientTotalsTonnes.set(client, Number(clientTotalsTonnes.get(client) || 0) + tonnesQty);
    list.appendChild(
      item(
        `<b>${r.op_date || "-"}</b> <span class="pill blue">${r.product_type || "product"}</span>` +
          `<br><small>Tonnes moved: ${Number(r.tonnes_moved || 0).toFixed(2)} | Produced: ${Number(r.product_produced || 0).toFixed(2)}</small>` +
          `<br><small>Trucks loaded: ${Number(r.trucks_loaded || 0)} | Load cycles: ${Number(r.loads_count || 0)}</small>` +
          `<br><small>Crusher feed: ${Number(r.crusher_feed_tonnes || 0).toFixed(2)}t | Crusher output: ${Number(r.crusher_output_tonnes || 0).toFixed(2)}t | Crusher h: ${Number(r.crusher_hours || 0).toFixed(2)} | Downtime h: ${Number(r.crusher_downtime_hours || 0).toFixed(2)}</small>` +
          `${r.notes ? `<br><small>Notes: ${r.notes}</small>` : ""}`
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No operations entries found.</small>"));
  setText("opKpiTonnes", tonnes.toFixed(2));
  setText("opKpiProduced", produced.toFixed(2));
  setText("opKpiLoaded", String(loaded));
  setText("opKpiLoads", String(loadCycles));
  const crusherPerf = crusherFeed > 0 ? (crusherOutput / crusherFeed) * 100 : 0;
  setText("opKpiCrusherPerf", crusherPerf.toFixed(1));
  const metric = String(qs("opClientMetric")?.value || "delivered").toLowerCase();
  const metricMap =
    metric === "trucks" ? clientTotalsTrucks :
    metric === "tonnes" ? clientTotalsTonnes :
    clientTotalsDelivered;
  const metricLabel =
    metric === "trucks" ? "Trucks" :
    metric === "tonnes" ? "Tonnes" :
    "Delivered";
  const topN = Math.max(1, Number(qs("opClientTopN")?.value || 8));
  const topClients = Array.from(metricMap.entries())
    .map(([client, qty]) => ({ client, qty: Number(qty || 0) }))
    .sort((a, b) => b.qty - a.qty)
    .slice(0, topN);
  const maxQty = Math.max(1, ...topClients.map((x) => x.qty));
  const chart = qs("opClientChart");
  const perfList = qs("opClientPerfList");
  if (chart) {
    chart.innerHTML = "";
    if (!topClients.length) {
      chart.appendChild(item("<small>No client delivery data in selected range.</small>"));
    } else {
      topClients.forEach((x) => {
        const h = Math.max(6, Math.round((x.qty / maxQty) * 100));
        const bar = document.createElement("div");
        bar.className = "cost-bar op-client-bar";
        bar.style.height = `${h}px`;
        bar.title = `${x.client}: ${x.qty.toFixed(2)} (${metricLabel})`;
        bar.innerHTML =
          `<span class="cost-bar-value">${x.qty.toFixed(1)}</span>` +
          `<span class="cost-bar-label">${x.client.length > 10 ? `${x.client.slice(0, 10)}...` : x.client}</span>`;
        chart.appendChild(bar);
      });
    }
  }
  if (perfList) {
    perfList.innerHTML = "";
    if (!topClients.length) {
      perfList.appendChild(item("<small>No client totals available.</small>"));
    } else {
      topClients.forEach((x, i) => {
        perfList.appendChild(item(`<b>#${i + 1}</b> ${x.client}<br><small>${metricLabel}: ${x.qty.toFixed(2)}</small>`));
      });
    }
  }
  const opDate = (qs("opDate")?.value || "").trim();
  if (opDate) {
    loadOperationsClosingForDate(opDate).catch(() => {});
  }
  setStatus("Operations ready.");
}

async function createDispatchTrip() {
  const payload = {
    op_date: (qs("dpDate")?.value || "").trim() || undefined,
    trip_no: (qs("dpTripNo")?.value || "").trim() || undefined,
    truck_reg: (qs("dpTruckReg")?.value || "").trim(),
    driver_name: (qs("dpDriver")?.value || "").trim() || undefined,
    product_type: (qs("dpProduct")?.value || "").trim() || undefined,
    client_name: (qs("dpClient")?.value || "").trim() || undefined,
    target_tonnes: (qs("dpTargetTonnes")?.value || "").trim() === "" ? undefined : Number(qs("dpTargetTonnes")?.value || 0),
    actual_tonnes: (qs("dpActualTonnes")?.value || "").trim() === "" ? undefined : Number(qs("dpActualTonnes")?.value || 0),
    notes: (qs("dpNotes")?.value || "").trim() || undefined,
  };
  if (!payload.truck_reg) {
    alert("Truck reg is required.");
    return;
  }
  setStatus("Creating dispatch trip...");
  const res = await fetchJson(`${API}/api/dispatch/trips`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("dispatchResult", JSON.stringify(res, null, 2));
  await loadDispatchTrips();
  setStatus("Dispatch trip created.");
}

async function updateDispatchTripStatus(id, status) {
  const requirePod = Boolean(qs("dpRequirePodDelivered")?.checked);
  const res = await fetchJson(`${API}/api/dispatch/trips/${id}/status`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ status, require_pod_for_delivered: requirePod ? 1 : 0 }),
  });
  setText("dispatchResult", JSON.stringify(res?.row || res, null, 2));
  await loadDispatchTrips();
}

async function saveDispatchPod() {
  const tripId = Number(qs("dpActionTripId")?.value || 0);
  if (!tripId) {
    alert("Trip ID is required for POD.");
    return;
  }
  const pod_ref = String(qs("dpPodRef")?.value || "").trim();
  if (!pod_ref) {
    alert("POD ref is required.");
    return;
  }
  const pod_link = String(qs("dpPodLink")?.value || "").trim() || undefined;
  const res = await fetchJson(`${API}/api/dispatch/trips/${tripId}/pod`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ pod_ref, pod_link }),
  });
  setText("dispatchResult", JSON.stringify(res?.row || res, null, 2));
  await loadDispatchTrips();
}

async function createDispatchException() {
  const trip_id = Number(qs("dpActionTripId")?.value || 0);
  if (!trip_id) {
    alert("Trip ID is required for exception.");
    return;
  }
  const exception_type = String(qs("dpExType")?.value || "").trim();
  const severity = String(qs("dpExSeverity")?.value || "medium").trim();
  const owner_name = String(qs("dpExOwner")?.value || "").trim() || undefined;
  const note = String(qs("dpExNote")?.value || "").trim() || undefined;
  const res = await fetchJson(`${API}/api/dispatch/exceptions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ trip_id, exception_type, severity, owner_name, note }),
  });
  setText("dispatchResult", JSON.stringify(res, null, 2));
  await loadDispatchExceptions();
  await loadDispatchTrips();
}

async function resolveDispatchException(id, status) {
  const resolution_note = String(qs("dpExResolveNote")?.value || "").trim() || undefined;
  const res = await fetchJson(`${API}/api/dispatch/exceptions/${id}/resolve`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ status, resolution_note }),
  });
  setText("dispatchResult", JSON.stringify(res?.row || res, null, 2));
  await loadDispatchExceptions();
  await loadDispatchTrips();
}

async function loadDispatchExceptions() {
  const list = qs("dispatchExceptionsList");
  if (!list) return;
  const from = (qs("dpFrom")?.value || "").trim();
  const to = (qs("dpTo")?.value || "").trim();
  const onlyOpen = Boolean(qs("dpExOnlyOpen")?.checked);
  const status = onlyOpen ? "open" : (qs("dpExStatusFilter")?.value || "").trim();
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (to) q.set("to", to);
  if (status) q.set("status", status);
  const data = await fetchJson(`${API}/api/dispatch/exceptions?${q.toString()}`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  let crit = 0;
  let high = 0;
  let med = 0;
  let low = 0;
  list.innerHTML = "";
  rows.forEach((r) => {
    const s = String(r.status || "").toLowerCase();
    const sev = String(r.severity || "").toLowerCase();
    if (s === "open") {
      if (sev === "critical") crit += 1;
      else if (sev === "high") high += 1;
      else if (sev === "medium") med += 1;
      else low += 1;
    }
    const badgeClass = s === "open" ? "red" : s === "resolved" ? "blue" : "";
    const actions = s === "open"
      ? `<button data-dp-ex-id="${r.id}" data-dp-ex-next="resolved">Resolve</button> <button data-dp-ex-id="${r.id}" data-dp-ex-next="waived">Waive</button>`
      : "";
    list.appendChild(
      item(
        `<b>EX #${r.id}</b> <span class="pill ${badgeClass}">${r.status}</span> <span class="pill">${r.exception_type}</span> <span class="pill">${r.severity}</span>` +
        `<br><small>Trip #${r.trip_id} | ${r.op_date || "-"} | Truck ${r.truck_reg || "-"} | Client ${r.client_name || "-"}</small>` +
        `<br><small>Owner: ${r.owner_name || "-"} | Note: ${r.note || "-"}</small>` +
        `${r.resolution_note ? `<br><small>Resolution: ${r.resolution_note}</small>` : ""}` +
        `<br>${actions}`
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No exceptions found.</small>"));
  setText("dpExKpiCritical", String(crit));
  setText("dpExKpiHigh", String(high));
  setText("dpExKpiMedium", String(med));
  setText("dpExKpiLow", String(low));
}

async function loadQualityCenter() {
  const list = qs("qualityList");
  if (!list) return;
  const from = (qs("qFrom")?.value || "").trim();
  const to = (qs("qTo")?.value || "").trim();
  const sev = (qs("qSeverityFilter")?.value || "").trim().toLowerCase();
  const typ = (qs("qTypeFilter")?.value || "").trim();
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (to) q.set("to", to);
  setStatus("Loading data quality...");
  setSkeleton("qualityList", 2);
  const data = await fetchJson(`${API}/api/quality?${q.toString()}`);
  const summary = data?.summary || {};
  setText("qTotal", String(Number(summary.total || 0)));
  setText("qHigh", String(Number(summary.high || 0)));
  setText("qMedium", String(Number(summary.medium || 0)));
  setText("qLow", String(Number(summary.low || 0)));
  setText("qDaily", String(Number(summary.daily_issues || 0)));
  setText("qPodGaps", String(Number(summary.dispatch_pod_gaps || 0)));
  setText("qOpenExceptions", String(Number(summary.exceptions_open || 0)));
  setText("qPendingApprovals", String(Number(summary.approvals_pending || 0)));

  let rows = Array.isArray(data?.rows) ? data.rows : [];
  if (sev) rows = rows.filter((r) => String(r.severity || "").toLowerCase() === sev);
  if (typ) rows = rows.filter((r) => String(r.type || "") === typ);

  list.innerHTML = "";
  rows.slice(0, 500).forEach((r) => {
    const s = String(r.severity || "low").toLowerCase();
    const badgeClass = s === "high" ? "red" : s === "medium" ? "orange" : "";
    const fixBtn = `<button data-q-fix="1" data-q-type="${String(r.type || "")}" data-q-asset="${String(r.asset_code || "")}" data-q-entity="${String(r.entity_id || "")}" data-q-date="${String(r.date || "")}">Open Fix</button>`;
    const resolveBtn = String(r.type || "") === "dispatch_delivered_no_pod"
      ? `<button data-q-resolve="pod" data-q-entity="${String(r.entity_id || "")}" data-q-date="${String(r.date || "")}">Resolve & Refresh</button>`
      : String(r.type || "") === "dispatch_exception_open"
      ? `<button data-q-resolve="exception_resolve" data-q-entity="${String(r.entity_id || "")}" data-q-date="${String(r.date || "")}">Resolve & Refresh</button> <button data-q-resolve="exception_waive" data-q-entity="${String(r.entity_id || "")}" data-q-date="${String(r.date || "")}">Waive & Refresh</button>`
      : "";
    list.appendChild(
      item(
        `<b>${r.type}</b> <span class="pill ${badgeClass}">${s}</span>` +
        `<br><small>Date: ${r.date || "-"} | Asset/Entity: ${r.asset_code || r.entity_id || "-"}</small>` +
        `<br><small>${r.details || "-"}</small>` +
        `<br>${fixBtn} ${resolveBtn}`
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No quality issues for selected filters.</small>"));
  setText("qualityResult", JSON.stringify({ from: data?.from, to: data?.to, shown: rows.length }, null, 2));
  setStatus("Data quality ready.");
}

async function resolveQualityIssueNow(mode, entityId, issueDate) {
  const id = Number(entityId || 0);
  if (!id) return;
  if (mode === "pod") {
    const pod_ref = prompt(`Enter POD ref for delivered trip #${id}:`, "") || "";
    if (!pod_ref.trim()) return;
    const pod_link = prompt("Optional POD link/path:", "") || "";
    await fetchJson(`${API}/api/dispatch/trips/${id}/pod`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pod_ref: pod_ref.trim(), pod_link: pod_link.trim() || undefined }),
    });
    setStatus(`POD saved for trip #${id}. Rechecking quality...`);
    if (qs("qFrom") && issueDate && !qs("qFrom").value) qs("qFrom").value = String(issueDate);
    if (qs("qTo") && issueDate && !qs("qTo").value) qs("qTo").value = String(issueDate);
    await loadQualityCenter();
    return;
  }
  if (mode === "exception_resolve" || mode === "exception_waive") {
    const resolution_note = prompt("Resolution note:", "") || "";
    const status = mode === "exception_waive" ? "waived" : "resolved";
    await fetchJson(`${API}/api/dispatch/exceptions/${id}/resolve`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ status, resolution_note: resolution_note.trim() || undefined }),
    });
    setStatus(`Exception #${id} set to ${status}. Rechecking quality...`);
    await loadQualityCenter();
  }
}

function openQualityFix(type, assetCode, entityId, issueDate) {
  const t = String(type || "");
  if (t === "dispatch_delivered_no_pod") {
    switchTab("dispatch");
    if (qs("dpActionTripId")) qs("dpActionTripId").value = String(entityId || "");
    if (qs("dpStatusFilter")) qs("dpStatusFilter").value = "delivered";
    if (qs("dpFrom") && issueDate) qs("dpFrom").value = String(issueDate);
    if (qs("dpTo") && issueDate) qs("dpTo").value = String(issueDate);
    loadDispatchTrips().catch(() => {});
    setStatus(`Opened Dispatch fix for trip #${entityId}. Save POD and re-check.`);
    return;
  }
  if (t === "dispatch_exception_open") {
    switchTab("dispatch");
    if (qs("dpExStatusFilter")) qs("dpExStatusFilter").value = "open";
    if (qs("dpFrom") && issueDate) qs("dpFrom").value = String(issueDate);
    if (qs("dpTo") && issueDate) qs("dpTo").value = String(issueDate);
    loadDispatchExceptions().catch(() => {});
    setStatus(`Opened Dispatch exceptions. Resolve exception #${entityId}.`);
    return;
  }
  if (t === "approval_pending") {
    switchTab("approvals");
    if (qs("approvalStatus")) qs("approvalStatus").value = "pending";
    loadApprovalRequests().catch(() => {});
    setStatus(`Opened Approvals. Review pending approval #${entityId}.`);
    return;
  }
  if (t.startsWith("daily_")) {
    switchTab("daily");
    if (qs("date") && issueDate) qs("date").value = String(issueDate);
    loadDailyInput().catch(() => {});
    setStatus(`Opened Daily Input for ${issueDate}. Check asset ${assetCode || "-"}.`);
    return;
  }
  setStatus("No quick-fix route for this issue type yet.");
}

function dispatchVarianceMeta(targetTonnes, actualTonnes, tolerancePct) {
  const target = Number(targetTonnes || 0);
  const actual = Number(actualTonnes || 0);
  if (!Number.isFinite(target) || target <= 0) {
    return { pct: 0, cls: "var-warn", label: "NO TARGET", breached: false };
  }
  const pct = ((actual - target) / target) * 100;
  const absPct = Math.abs(pct);
  const t = Math.max(0, Number(tolerancePct || 0));
  if (absPct <= t * 0.5) return { pct, cls: "var-good", label: "OK", breached: false };
  if (absPct <= t) return { pct, cls: "var-warn", label: "WARN", breached: false };
  return { pct, cls: "var-breach", label: "BREACH", breached: true };
}

async function loadDispatchTrips() {
  const list = qs("dispatchList");
  if (!list) return;
  const from = (qs("dpFrom")?.value || "").trim();
  const to = (qs("dpTo")?.value || "").trim();
  const status = (qs("dpStatusFilter")?.value || "").trim();
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (to) q.set("to", to);
  if (status) q.set("status", status);
  setStatus("Loading dispatch trips...");
  setSkeleton("dispatchList", 2);
  const [kpi, trips] = await Promise.all([
    fetchJson(`${API}/api/dispatch/kpi?${q.toString()}`),
    fetchJson(`${API}/api/dispatch/trips?${q.toString()}`),
  ]);
  setText("dpKpiTrips", String(Number(kpi?.total_trips || 0)));
  setText("dpKpiDeliveredTonnes", Number(kpi?.delivered_tonnes || 0).toFixed(2));
  setText("dpKpiTurnaround", Number(kpi?.avg_turnaround_hours || 0).toFixed(2));
  setText("dpKpiQueued", String(Number(kpi?.by_status?.queued || 0)));
  setText("dpKpiLoading", String(Number(kpi?.by_status?.loading || 0)));
  setText("dpKpiTransit", String(Number(kpi?.by_status?.in_transit || 0)));
  setText("dpKpiDelivered", String(Number(kpi?.by_status?.delivered || 0)));
  setText("dpKpiReturned", String(Number(kpi?.by_status?.returned || 0)));
  setText("dpKpiPodPct", kpi?.delivered_with_pod_pct == null ? "-" : `${Number(kpi.delivered_with_pod_pct).toFixed(1)}%`);
  setText("dpKpiExceptionsOpen", String(Number(kpi?.exceptions_open || 0)));

  const rows = Array.isArray(trips?.rows) ? trips.rows : [];
  const tolerancePct = Number(qs("dpVarTolerance")?.value || 10);
  const onlyBreaches = Boolean(qs("dpOnlyBreaches")?.checked);
  list.innerHTML = "";
  const lanes = {
    queued: qs("dpLaneQueued"),
    loading: qs("dpLaneLoading"),
    in_transit: qs("dpLaneTransit"),
    delivered: qs("dpLaneDelivered"),
    returned: qs("dpLaneReturned"),
  };
  Object.values(lanes).forEach((el) => { if (el) el.innerHTML = ""; });
  const addLane = (key, html) => {
    if (lanes[key]) lanes[key].appendChild(item(html));
  };
  let varianceBreaches = 0;

  rows.forEach((r) => {
    const s = String(r.status || "queued");
    const variance = dispatchVarianceMeta(r.target_tonnes, r.actual_tonnes, tolerancePct);
    if (onlyBreaches && !variance.breached) return;
    if (variance.breached) varianceBreaches += 1;
    const actions = []
      .concat(s !== "loading" ? [`<button data-dp-status-id="${r.id}" data-dp-next="loading">Loading</button>`] : [])
      .concat(s !== "in_transit" ? [`<button data-dp-status-id="${r.id}" data-dp-next="in_transit">In Transit</button>`] : [])
      .concat(s !== "delivered" ? [`<button data-dp-status-id="${r.id}" data-dp-next="delivered">Delivered</button>`] : [])
      .concat(s !== "returned" ? [`<button data-dp-status-id="${r.id}" data-dp-next="returned">Returned</button>`] : [])
      .join(" ");
    const html =
      `<b>Trip #${r.id}</b> <span class="pill blue">${s}</span> ${r.trip_no ? `<span class="pill">${r.trip_no}</span>` : ""}` +
      ` <span class="pill ${variance.cls}">${variance.label} ${Number(variance.pct || 0).toFixed(1)}%</span>` +
      `<br><small>${r.op_date || "-"} | Truck: ${r.truck_reg || "-"} | Driver: ${r.driver_name || "-"}</small>` +
      `<br><small>Product: ${r.product_type || "-"} | Client: ${r.client_name || "-"}</small>` +
      `<br><small>Target: ${Number(r.target_tonnes || 0).toFixed(2)} | Actual: ${Number(r.actual_tonnes || 0).toFixed(2)} | POD: ${r.pod_ref || "-"}</small>` +
      `<br>${actions}`;
    list.appendChild(item(html));
    addLane(s, html);
  });
  if (!rows.length) {
    list.appendChild(item("<small>No dispatch trips found.</small>"));
  }
  Object.entries(lanes).forEach(([k, el]) => {
    if (el && !el.children.length) el.appendChild(item(`<small>No ${k.replace("_", " ")} trips.</small>`));
  });
  setText("dpKpiVarianceBreach", String(varianceBreaches));
  loadDispatchExceptions().catch(() => {});
  setStatus("Dispatch ready.");
}

/* =========================
   DAILY INPUT (GRID)
========================= */

let dailyRows = [];
let dailyShowDownOnly = false;
let dailyPrestartRows = [];
let dailyPrestartMeta = {
  deduction_hours_per_check: 0.5,
  production_deduction_hours: 0,
  production_deduction_count: 0,
};

function fmt(n) {
  if (n === null || n === undefined || Number.isNaN(n)) return "";
  return Number(n).toFixed(1).replace(/\.0$/, "");
}
function toNum(v) {
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  if (s === "") return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}
function calcRun(opening, closing) {
  if (opening == null || closing == null) return 0;
  const run = closing - opening;
  return Number.isFinite(run) ? run : 0;
}
function daySummary() {
  const prod = dailyRows.filter((r) => r.is_used).length;
  const standby = dailyRows.filter((r) => !r.is_used).length;
  const bad = dailyRows.filter((r) => r.error).length;
  return `Production: ${prod} | Standby: ${standby} | Errors: ${bad}`;
}
function prevDateStr(dateStr) {
  const d = new Date(dateStr + "T00:00:00");
  d.setDate(d.getDate() - 1);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function todayLocalYmd() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function validateDailyRows() {
  for (const r of dailyRows) {
    r.error = null;
    r.warning = null;

    r.hours_run = calcRun(r.opening_hours, r.closing_hours);

    if (r.is_down) {
      const dh = r.down_hours != null ? Number(r.down_hours) : Number(r.scheduled_hours || 0);
      if (!Number.isFinite(dh) || dh < 0) {
        r.error = "DOWN HOURS INVALID — MUST BE >= 0.";
        continue;
      }
      if (r.scheduled_hours != null && dh > Number(r.scheduled_hours || 0)) {
        r.error = "DOWN HOURS TOO HIGH — MUST BE <= SCHEDULED HOURS.";
        continue;
      }
      if (r.opening_hours != null && (r.closing_hours == null || r.closing_hours === "")) {
        r.closing_hours = r.opening_hours;
        r.hours_run = 0;
      }
      continue;
    }

    if (r.telematics_locked) {
      if (!r.is_used && r.hours_run > 0) {
        r.error = "STANDBY SELECTED — HOURS NOT ALLOWED.";
        continue;
      }
      if (r.is_used && (r.scheduled_hours == null || r.scheduled_hours === 0)) {
        r.error = "PRODUCTION SELECTED — SCHEDULED HOURS IS 0.";
      } else if (r.is_used && r.hours_run === 0) {
        r.warning = "TELEMATICS — METER WILL SYNC FROM FSC ON SAVE.";
      }
      continue;
    }

    if (!r.is_used && r.hours_run > 0) {
      r.error = "STANDBY SELECTED — HOURS NOT ALLOWED.";
      continue;
    }

    if (r.is_used && r.hours_run === 0) {
      r.error = "PRODUCTION SELECTED — NO HOURS. CHECK CLOSING HOURMETER.";
      continue;
    }

    if (r.is_used && (r.scheduled_hours == null || r.scheduled_hours === 0)) {
      r.error = "PRODUCTION SELECTED — SCHEDULED HOURS IS 0.";
      continue;
    }

    if (r.is_used && r.opening_hours == null) {
      r.warning = "OPENING HOURS MISSING — CHECK YESTERDAY CLOSING.";
    }

    if (r.opening_hours != null && r.closing_hours != null && r.closing_hours < r.opening_hours) {
      r.error = "HOURMETER MISMATCH — CLOSING LOWER THAN OPENING.";
    }
  }
}

/* -------- KPI Preview -------- */

function calcDailyPreviewKpis() {
  const used = dailyRows.filter(
    (r) => r.is_used && !r.is_master_standby && String(r.input_unit || "hours").toLowerCase() !== "km"
  );
  const usedCount = used.length;

  let totalScheduled = 0;
  let totalRun = 0;
  let totalDowntime = 0;

  let totalPrestart = 0;
  let prestartCount = 0;
  const prestartPerCheck = Number(dailyPrestartMeta.deduction_hours_per_check || 0.5);
  const prestartCodes = new Set(
    dailyPrestartRows.map((r) => String(r.asset_code || "").trim()).filter(Boolean)
  );

  for (const r of used) {
    const scheduled = Math.max(0, Number(r.scheduled_hours || 0));
    const runRaw = Math.max(0, Number(r.hours_run || 0));
    const runEff = Math.min(runRaw, scheduled);
    const downRaw = r.is_down
      ? (Number.isFinite(Number(r.down_hours)) ? Number(r.down_hours) : scheduled)
      : 0;
    const downEff = Math.min(Math.max(0, downRaw), scheduled);

    totalScheduled += scheduled;
    totalRun += runEff;
    totalDowntime += downEff;

    if (prestartCodes.has(r.asset_code)) {
      totalPrestart += prestartPerCheck;
      prestartCount += 1;
    }
  }

  const utilization = totalScheduled > 0 ? (totalRun / totalScheduled) * 100 : null;
  const available = Math.max(0, totalScheduled - totalDowntime - totalPrestart);
  const availability = totalScheduled > 0 ? (available / totalScheduled) * 100 : null;

  return {
    usedCount,
    totalScheduled,
    totalRun,
    totalDowntime,
    totalPrestart,
    prestartCount,
    availability,
    utilization,
  };
}

function renderDailyPreview() {
  setText("dailySummary", daySummary());

  const k = calcDailyPreviewKpis();
  const th = getThresholds();

  setText("kUsed", `Used: ${k.usedCount}`);
  setText("kSched", `Scheduled: ${k.totalScheduled.toFixed(1).replace(/\.0$/, "")}`);
  setText("kRun", `Run: ${k.totalRun.toFixed(1).replace(/\.0$/, "")}`);

  setSpeedo(qs("pAvailNeedle"), qs("pAvailVal"), k.availability, {
    goodAt: th.availTarget,
    warnAt: th.availCrit,
  });
  setSpeedo(qs("pUtilNeedle"), qs("pUtilVal"), k.utilization, {
    goodAt: th.utilTarget,
    warnAt: th.utilCrit,
  });

  if (k.totalScheduled === 0) setText("kNote", "Preview waiting for scheduled/run hours. Standby excluded.");
  else {
    const prestartNote = k.prestartCount > 0
      ? ` Pre-start: ${k.prestartCount} production check(s), ${k.totalPrestart.toFixed(1).replace(/\.0$/, "")} hr deducted from availability.`
      : "";
    setText(
      "kNote",
      `Preview uses production rows only (standby excluded). Down hours: ${k.totalDowntime.toFixed(1).replace(/\.0$/, "")}.${prestartNote}`
    );
  }
}

function renderDailyPrestartSection() {
  const panel = qs("dailyPrestartPanel");
  if (!panel) return;

  const perCheck = Number(dailyPrestartMeta.deduction_hours_per_check || 0.5);
  const perCheckLabel = perCheck.toFixed(1).replace(/\.0$/, "");

  if (!dailyPrestartRows.length) {
    panel.innerHTML = `
      <div class="daily-prestart-header">
        <strong>Daily pre-start checks</strong>
        <span class="muted small">None logged for this date</span>
      </div>`;
    panel.classList.add("is-empty");
    return;
  }

  panel.classList.remove("is-empty");
  const prodCount = Number(dailyPrestartMeta.production_deduction_count || 0);
  const prodHours = Number(dailyPrestartMeta.production_deduction_hours || 0);
  const chips = dailyPrestartRows.map((r) => {
    const code = String(r.asset_code || "").trim();
    const type = String(r.check_type || "Pre-start");
    const inspector = r.inspector_name ? ` · ${String(r.inspector_name).trim()}` : "";
    return `<span class="daily-prestart-chip" title="${type}${inspector}"><b>${code}</b> ${type}</span>`;
  }).join("");

  panel.innerHTML = `
    <div class="daily-prestart-header">
      <strong>Daily pre-start checks</strong>
      <span class="muted small">${dailyPrestartRows.length} completed · ${perCheckLabel} hr each</span>
    </div>
    <div class="daily-prestart-list">${chips}</div>
    <div class="daily-prestart-foot muted small">
      ${prodCount > 0
        ? `${prodCount} production asset pre-start(s) deduct ${prodHours.toFixed(1).replace(/\.0$/, "")} hr from preview availability.`
        : "Pre-starts shown for reference; none apply to production hour-based availability today."}
    </div>`;
}

/* -------- DOWN helper -------- */

async function logDownRowToBreakdowns(date, r) {
  try {
    const downDesc = r.down_reason ? `DOWN — ${r.down_reason}` : "DOWN";
    const b = await fetchJson(`${API}/api/breakdowns/ensure-open`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        asset_code: r.asset_code,
        breakdown_date: date,
        start_date: String(r.breakdown_start_date || "").trim() || date,
        description: downDesc,
        critical: false,
        parts_ordered_date: String(r.parts_ordered_date || "").trim() || null,
        parts_status: String(r.parts_status || "").trim() || null,
        parts_received_date: String(r.parts_received_date || "").trim() || null,
        ets_repair_date: String(r.ets_repair_date || "").trim() || null,
      }),
    });

    const breakdownId = b.breakdown_id || b.breakdownId || b.id;
    if (!breakdownId) return;

    const notes = r.down_reason
      ? `Auto from Daily Input (DOWN) — ${r.down_reason}`
      : "Auto from Daily Input (DOWN)";

    const downHoursRaw = r.down_hours != null ? Number(r.down_hours) : Number(r.scheduled_hours || 0);
    const downHours = Number.isFinite(downHoursRaw) ? Math.max(0, downHoursRaw) : 0;

    await fetchJson(`${API}/api/breakdowns/${breakdownId}/downtime`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        log_date: date,
        hours_down: downHours,
        notes,
      }),
    });
  } catch {
    // swallow
  }
}

function renderDailyTable() {
  const body = qs("dailyBody");
  if (!body) return;

  body.innerHTML = "";
  const rowsToRender = dailyShowDownOnly ? dailyRows.filter((r) => !!r.is_down) : dailyRows;

  // Create a container for card-style rows
  const container = document.createElement("div");
  container.className = "daily-rows-container";

  for (const r of rowsToRender) {
    const rowClass = [
      "daily-row-card",
      r.is_master_standby ? "standby" : "",
      r.is_down ? "down" : "",
      r.error ? "has-error" : "",
      r.warning ? "has-warning" : ""
    ].filter(Boolean).join(" ");

    const card = document.createElement("div");
    card.className = rowClass;

    // Asset Header
    const assetHeader = document.createElement("div");
    assetHeader.className = "daily-row-header";
    assetHeader.innerHTML = `
      <div class="daily-asset-info">
        <span class="daily-asset-code">${r.asset_code}</span>
        <button class="daily-btn-icon daily-asset-qr" title="Download QR PNG for ${r.asset_code}">QR</button>
        <span class="daily-asset-name">${r.asset_name || ""}</span>
        ${r.telematics_locked ? '<span class="pill blue" style="margin-left:8px">TELEMATICS</span>' : ""}
      </div>
      <div class="daily-status-badge ${r.is_down ? 'status-down' : r.error ? 'status-error' : r.warning ? 'status-warning' : 'status-ok'}">
        ${r.is_down ? '⬇ DOWN' : r.error ? '✕ ' + r.error : r.warning ? '⚠ ' + r.warning : r.is_used ? '▶ PRODUCTION' : '⏸ STANDBY'}
      </div>
    `;
    card.appendChild(assetHeader);

    // Main Content Grid
    const contentGrid = document.createElement("div");
    contentGrid.className = "daily-row-content";

    // Left Column - Hours
    const hoursCol = document.createElement("div");
    hoursCol.className = "daily-hours-col";
    hoursCol.innerHTML = `
      <div class="daily-hours-grid">
        <div class="daily-hour-field">
          <label>Scheduled</label>
          <input type="number" step="0.5" min="0" max="24" value="${fmt(r.scheduled_hours)}" class="daily-input" />
        </div>
        <div class="daily-hour-field">
          <label>Opening</label>
          <input type="number" step="0.1" value="${fmt(r.opening_hours)}" class="daily-input readonly" readonly title="${r.opening_from_date ? `Auto-filled from ${r.opening_from_date}` : 'Auto-filled from yesterday'}" />
        </div>
        <div class="daily-hour-field">
          <label>Closing</label>
          <input type="number" step="0.1" value="${fmt(r.closing_hours)}" class="daily-input ${r.is_down || r.telematics_locked ? 'disabled readonly' : ''}" ${r.is_down || r.telematics_locked ? 'readonly' : ''} title="${r.telematics_locked ? 'Hourmeter from FSC telematics' : ''}" />
        </div>
        <div class="daily-hour-field">
          <label>Run</label>
          <div class="daily-run-display">${fmt(r.hours_run)} ${String(r.input_unit || "hours").toLowerCase() === "km" ? "km" : "hrs"}</div>
        </div>
      </div>
    `;
    contentGrid.appendChild(hoursCol);

    // Right Column - Controls
    const controlsCol = document.createElement("div");
    controlsCol.className = "daily-controls-col";

    // Production Toggle
    const prodToggle = document.createElement("div");
    prodToggle.className = "daily-toggle-group";
    prodToggle.innerHTML = `
      <label class="daily-toggle ${r.is_used ? 'active' : ''} ${r.is_master_standby ? 'disabled' : ''}">
        <input type="checkbox" ${r.is_used ? 'checked' : ''} ${r.is_master_standby ? 'disabled' : ''} />
        <span class="toggle-track"><span class="toggle-thumb"></span></span>
        <span class="toggle-label">PROD</span>
      </label>
      <label class="daily-toggle down ${r.is_down ? 'active' : ''} ${r.is_master_standby || (r.is_down && r.down_lock) ? 'disabled' : ''}">
        <input type="checkbox" ${r.is_down ? 'checked' : ''} ${r.is_master_standby || (r.is_down && r.down_lock) ? 'disabled' : ''} />
        <span class="toggle-track"><span class="toggle-thumb"></span></span>
        <span class="toggle-label">DOWN</span>
      </label>
    `;
    controlsCol.appendChild(prodToggle);

    // Unit selector
    const unitSelect = document.createElement("div");
    unitSelect.className = "daily-unit-select";
    unitSelect.innerHTML = `
      <select class="daily-select ${r.is_down && r.down_lock ? 'disabled' : ''}" ${r.is_down && r.down_lock ? 'disabled' : ''}>
        <option value="hours" ${String(r.input_unit || "hours").toLowerCase() === "hours" ? 'selected' : ''}>HRS</option>
        <option value="km" ${String(r.input_unit || "hours").toLowerCase() === "km" ? 'selected' : ''}>KM</option>
      </select>
      <button class="daily-btn-icon reset-unit" title="Reset to suggested (${r.suggested_input_unit || 'hours'})" ${r.is_down && r.down_lock ? 'disabled' : ''}>↺</button>
    `;
    controlsCol.appendChild(unitSelect);

    contentGrid.appendChild(controlsCol);

    // Down Details Row (shown when down)
    if (r.is_down) {
      const downDetails = document.createElement("div");
      downDetails.className = "daily-down-details";
      
      const startYmd = String(r.breakdown_start_date || "").trim().match(/^(\d{4}-\d{2}-\d{2})/)?.[1] || "";
      const calcDaysDown = (() => {
        if (!startYmd) return null;
        const endYmd = String(qs("date")?.value || todayLocalYmd());
        const s = Date.parse(`${startYmd}T00:00:00`);
        const e = Date.parse(`${endYmd}T00:00:00`);
        if (!Number.isFinite(s) || !Number.isFinite(e) || e < s) return null;
        return Math.floor((e - s) / 86400000) + 1;
      })();

      downDetails.innerHTML = `
        <div class="down-details-grid">
          <div class="down-field">
            <label>Reason</label>
            <select class="daily-select ${r.is_down && r.down_lock ? 'disabled' : ''}" ${r.is_down && r.down_lock ? 'disabled' : ''}>
              <option value="">Select reason...</option>
              <option value="Mechanical" ${r.down_reason === "Mechanical" ? 'selected' : ''}>Mechanical</option>
              <option value="Electrical" ${r.down_reason === "Electrical" ? 'selected' : ''}>Electrical</option>
              <option value="Hydraulics" ${r.down_reason === "Hydraulics" ? 'selected' : ''}>Hydraulics</option>
              <option value="Tyres/Undercarriage" ${r.down_reason === "Tyres/Undercarriage" ? 'selected' : ''}>Tyres/Undercarriage</option>
              <option value="Waiting parts" ${r.down_reason === "Waiting parts" ? 'selected' : ''}>Waiting parts</option>
              <option value="No operator" ${r.down_reason === "No operator" ? 'selected' : ''}>No operator</option>
              <option value="Weather/Access" ${r.down_reason === "Weather/Access" ? 'selected' : ''}>Weather/Access</option>
            </select>
          </div>
          <div class="down-field">
            <label>Down Hours</label>
            <input type="number" step="0.5" min="0" max="24" value="${fmt(r.down_hours != null ? r.down_hours : Number(r.scheduled_hours || 0))}" class="daily-input" placeholder="0" />
          </div>
          <div class="down-field">
            <label>Date down</label>
            <input type="date" value="${startYmd}" class="daily-input" data-down-field="start" />
          </div>
          <div class="down-field">
            <label>Date parts ordered</label>
            <input type="date" value="${String(r.parts_ordered_date || "").trim().match(/^(\d{4}-\d{2}-\d{2})/)?.[1] || ""}" class="daily-input" data-down-field="parts_ordered" />
          </div>
          <div class="down-field">
            <label>Status of parts</label>
            <select class="daily-select" data-down-field="parts_status">
              <option value="">—</option>
              <option value="Not ordered" ${r.parts_status === "Not ordered" ? "selected" : ""}>Not ordered</option>
              <option value="Ordered" ${r.parts_status === "Ordered" ? "selected" : ""}>Ordered</option>
              <option value="In transit" ${r.parts_status === "In transit" ? "selected" : ""}>In transit</option>
              <option value="Partial" ${r.parts_status === "Partial" ? "selected" : ""}>Partial</option>
              <option value="Received" ${r.parts_status === "Received" ? "selected" : ""}>Received</option>
              <option value="Waiting OEM" ${r.parts_status === "Waiting OEM" ? "selected" : ""}>Waiting OEM</option>
            </select>
          </div>
          <div class="down-field">
            <label>Received date</label>
            <input type="date" value="${String(r.parts_received_date || "").trim().match(/^(\d{4}-\d{2}-\d{2})/)?.[1] || ""}" class="daily-input" data-down-field="parts_received" />
          </div>
          <div class="down-field">
            <label>ETS Repair</label>
            <input type="date" value="${String(r.ets_repair_date || "").trim().match(/^(\d{4}-\d{2}-\d{2})/)?.[1] || ""}" class="daily-input" data-down-field="ets_repair" />
          </div>
          <div class="down-field full">
            <label>Comment</label>
            <input type="text" value="${r.breakdown_comment || ''}" class="daily-input" placeholder="Breakdown comment..." data-down-field="comment" />
          </div>
        </div>
        ${startYmd ? `<div class="down-meta">Date down: ${startYmd}${calcDaysDown != null ? ` | Total days down: ${calcDaysDown}` : ""}</div>` : ""}
        ${r.is_down && r.down_lock ? `<div class="down-lock-notice">Locked until WO is repaired (${r.lock_wo_status || "-"})</div>` : ""}
      `;
      contentGrid.appendChild(downDetails);
    }

    // Footer with opening source
    if (r.opening_from_date) {
      const footer = document.createElement("div");
      footer.className = "daily-row-footer";
      footer.innerHTML = `<span class="opening-source">Opening from: ${r.opening_from_date}</span>`;
      contentGrid.appendChild(footer);
    }

    card.appendChild(contentGrid);
    container.appendChild(card);

    // Add event listeners
    const inputs = card.querySelectorAll("input, select");

    inputs.forEach(input => {
      input.addEventListener("change", () => {
        const type = input.type;
        const cls = input.className;

        if (type === "number") {
          if (cls.includes("readonly")) {
            r.opening_hours = toNum(input.value);
          } else if (input.closest(".down-field") && input.previousElementSibling?.textContent === "Down Hours") {
            r.down_hours = toNum(input.value) ?? 0;
          } else if (input.previousElementSibling?.textContent === "Scheduled") {
            r.scheduled_hours = toNum(input.value) ?? 0;
          } else if (input.previousElementSibling?.textContent === "Closing") {
            r.closing_hours = toNum(input.value);
          }
        } else if (type === "date") {
          const downField = String(input.getAttribute("data-down-field") || "").trim();
          if (downField === "start" || input.previousElementSibling?.textContent === "Date down" || input.previousElementSibling?.textContent === "Started") {
            r.breakdown_start_date = String(input.value || "").trim();
          } else if (downField === "parts_ordered") {
            r.parts_ordered_date = String(input.value || "").trim();
          } else if (downField === "parts_received") {
            r.parts_received_date = String(input.value || "").trim();
          } else if (downField === "ets_repair") {
            r.ets_repair_date = String(input.value || "").trim();
          }
        } else if (type === "text") {
          if (input.getAttribute("data-down-field") === "comment" || input.previousElementSibling?.textContent === "Comment") {
            r.breakdown_comment = String(input.value || "").trim();
          }
        } else if (input.tagName === "SELECT") {
          if (cls.includes("daily-select") && !cls.includes("disabled")) {
            if (input.getAttribute("data-down-field") === "parts_status") {
              r.parts_status = String(input.value || "").trim();
            } else if (input.closest(".down-field")) {
              r.down_reason = String(input.value || "");
            } else if (input.value === "hours" || input.value === "km") {
              r.input_unit = input.value;
            }
          }
        }

        validateDailyRows();
        renderDailyTable();
        renderDailyPreview();
      });
    });

    // Toggle listeners
    const prodCheckbox = prodToggle.querySelector('input[type="checkbox"]:first-child');
    const downCheckbox = prodToggle.querySelector('input[type="checkbox"]:last-child');

    prodCheckbox?.addEventListener("change", () => {
      r.is_used = prodCheckbox.checked;
      if (r.is_master_standby) r.is_used = false;
      validateDailyRows();
      renderDailyTable();
      renderDailyPreview();
    });

    downCheckbox?.addEventListener("change", () => {
      r.is_down = downCheckbox.checked;
      if (r.is_down) {
        if (r.opening_hours != null) r.closing_hours = r.opening_hours;
        r.hours_run = 0;
        if (r.down_hours == null) r.down_hours = Number(r.scheduled_hours || 0);
      } else {
        r.down_reason = "";
        r.down_hours = null;
        r.breakdown_comment = "";
        r.breakdown_start_date = "";
        r.parts_ordered_date = "";
        r.parts_status = "";
        r.parts_received_date = "";
        r.ets_repair_date = "";
      }
      validateDailyRows();
      renderDailyTable();
      renderDailyPreview();
    });

    // Reset unit button
    const resetBtn = unitSelect.querySelector(".reset-unit");
    resetBtn?.addEventListener("click", () => {
      const suggested = String(r.suggested_input_unit || "hours").toLowerCase() === "km" ? "km" : "hours";
      r.input_unit = suggested;
      validateDailyRows();
      renderDailyTable();
      renderDailyPreview();
    });

    const assetQrBtn = assetHeader.querySelector(".daily-asset-qr");
    assetQrBtn?.addEventListener("click", (ev) => {
      ev.preventDefault();
      downloadAssetQrPng(r.asset_code).catch((e) => setStatus("QR download error: " + e.message));
    });
  }

  body.appendChild(container);
}

async function loadDailyInput() {
  const date = qs("date")?.value || todayLocalYmd();
  const y = prevDateStr(date);
  setStatus("Loading daily input...");
  setText("dailyResult", "");

  const assets = await fetchJson(`${API}/api/assets?include_archived=0`);

  let existing = [];
  try {
    existing = await fetchJson(`${API}/api/hours/${date}`);
  } catch {
    existing = [];
  }

  const existingByCode = new Map();
  for (const r of existing) existingByCode.set(r.asset_code, r);

  let yRows = [];
  try {
    yRows = await fetchJson(`${API}/api/hours/${y}`);
  } catch {
    yRows = [];
  }
  const yByCode = new Map();
  for (const r of yRows) yByCode.set(r.asset_code, r);

  dailyRows = [];

  let openBreakdownByAsset = new Map();
  try {
    const openData = await fetchJson(`${API}/api/breakdowns/open-all?date=${encodeURIComponent(date)}`);
    const rows = Array.isArray(openData?.rows) ? openData.rows : [];
    for (const bd of rows) {
      const code = String(bd.asset_code || "").trim();
      if (!code) continue;
      if (!openBreakdownByAsset.has(code)) openBreakdownByAsset.set(code, bd);
    }
  } catch {
    openBreakdownByAsset = new Map();
  }

  try {
    const ps = await fetchJson(`${API}/api/maintenance/prestart/daily-summary?date=${encodeURIComponent(date)}`);
    dailyPrestartRows = Array.isArray(ps?.rows) ? ps.rows : [];
    dailyPrestartMeta = {
      deduction_hours_per_check: Number(ps?.deduction_hours_per_check ?? 0.5),
      production_deduction_hours: Number(ps?.production_deduction?.hours ?? 0),
      production_deduction_count: Number(ps?.production_deduction?.count ?? 0),
    };
  } catch {
    dailyPrestartRows = [];
    dailyPrestartMeta = {
      deduction_hours_per_check: 0.5,
      production_deduction_hours: 0,
      production_deduction_count: 0,
    };
  }

  const parseDownReasonFromDesc = (desc) => {
    const d = String(desc || "").trim();
    if (!d) return "";
    const m = d.match(/DOWN\s*[-—:]\s*(.+)$/i);
    if (m && m[1]) return String(m[1]).trim();
    const m2 = d.match(/^DOWN\s*(.+)$/i);
    if (m2 && m2[1]) return String(m2[1]).trim();
    // Fallback for descriptions like: "DOWN � Hydraulics"
    const m3 = d.match(/^DOWN\s*[^A-Za-z0-9]*\s*(.+)$/i);
    if (m3 && m3[1]) return String(m3[1]).trim();
    return "";
  };

  for (const a of assets.filter((x) => x.active !== 0 && x.active !== false)) {
    const ex = existingByCode.get(a.asset_code);
    const masterStandby = !!a.is_standby;
    const forceOpenFromYesterday =
      ex &&
      ex.opening_hours != null &&
      (ex.closing_hours == null || ex.closing_hours === "") &&
      (ex.hours_run == null || Number(ex.hours_run) === 0);

    const row = {
      asset_code: a.asset_code,
      asset_name: a.asset_name,
      is_master_standby: masterStandby,

      is_used: ex
        ? !!ex.is_used
        : (yByCode.has(a.asset_code) ? !!yByCode.get(a.asset_code)?.is_used : !masterStandby),
      input_unit: ex?.input_unit
        ? String(ex.input_unit).toLowerCase()
        : (String(a.category || "").toLowerCase().includes("truck") || String(a.category || "").toLowerCase().includes("vehicle") ? "km" : "hours"),
      suggested_input_unit: ex?.input_unit
        ? String(ex.input_unit).toLowerCase()
        : (String(a.category || "").toLowerCase().includes("truck") || String(a.category || "").toLowerCase().includes("vehicle") ? "km" : "hours"),
      input_unit_locked: Boolean(ex?.input_unit_locked) || Boolean(ex?.telematics_locked),
      telematics_locked: Boolean(ex?.telematics_locked),
      meter_source: ex?.meter_source || "manual",

      scheduled_hours: ex ? toNum(ex.scheduled_hours) : null,
      opening_hours: ex ? toNum(ex.opening_hours) : null,
      opening_from_date: null,
      closing_hours: ex ? toNum(ex.closing_hours) : null,
      hours_run: ex ? toNum(ex.hours_run) ?? 0 : 0,

      is_down: false,
      down_reason: "",
      down_hours: null,
      down_lock: false,
      breakdown_start_date: "",
      breakdown_comment: ex?.notes ? String(ex.notes) : "",
      parts_ordered_date: "",
      parts_status: "",
      parts_received_date: "",
      ets_repair_date: "",

      error: null,
      warning: null,
    };

    if (row.is_master_standby) row.is_used = false;

    if (forceOpenFromYesterday) row.opening_hours = null;
    const yr = yByCode.get(row.asset_code);
    if (yr) {
      const yClose = toNum(yr.closing_hours);
      if ((row.opening_hours == null || forceOpenFromYesterday) && yClose != null) {
        row.opening_hours = yClose;
        row.opening_from_date = y;
      }
      if (row.scheduled_hours == null) {
        const ySched = toNum(yr.scheduled_hours);
        if (ySched != null) row.scheduled_hours = ySched;
      }
    }
    if (row.opening_hours == null || row.scheduled_hours == null) {
      try {
        const d = await fetchJson(
          `${API}/api/hours/defaults?asset_code=${encodeURIComponent(row.asset_code)}&work_date=${date}`
        );
        if (d?.suggested_input_unit) {
          const suggestedUnit = String(d.suggested_input_unit).toLowerCase() === "km" ? "km" : "hours";
          row.suggested_input_unit = suggestedUnit;
          row.input_unit = suggestedUnit;
        }
        if (!ex && typeof d?.suggested_is_used === "boolean") row.is_used = Boolean(d.suggested_is_used);
        if (typeof d?.input_unit_locked === "boolean") row.input_unit_locked = d.input_unit_locked;
        if ((row.opening_hours == null || forceOpenFromYesterday) && d.suggested_opening_hours != null) {
          row.opening_hours = Number(d.suggested_opening_hours);
          row.opening_from_date = String(d.suggested_opening_from_date || "").trim() || null;
        }
        if (row.scheduled_hours == null && d.suggested_scheduled_hours != null) row.scheduled_hours = Number(d.suggested_scheduled_hours);
      } catch {}
    }

    // Opening must match previous closing — refresh when prior day was corrected after save
    if (!row.telematics_locked && !row.is_master_standby && row.opening_hours != null) {
      let expectedOpen = null;
      let fromDate = null;
      const yClose = yr ? toNum(yr.closing_hours) : null;
      if (yClose != null) {
        expectedOpen = yClose;
        fromDate = y;
      } else {
        try {
          const d = await fetchJson(
            `${API}/api/hours/defaults?asset_code=${encodeURIComponent(row.asset_code)}&work_date=${date}`
          );
          if (d?.suggested_opening_hours != null) {
            expectedOpen = Number(d.suggested_opening_hours);
            fromDate = String(d.suggested_opening_from_date || "").trim() || null;
          }
        } catch {}
      }
      if (
        expectedOpen != null &&
        Math.abs(Number(row.opening_hours) - expectedOpen) > 0.0001
      ) {
        row.opening_hours = expectedOpen;
        row.opening_from_date = fromDate;
        if (!row.warning) row.warning = "Opening updated from corrected previous closing";
      }
    }

    if (row.scheduled_hours == null) row.scheduled_hours = 0;
    if (row.is_master_standby) row.is_used = false;
    row.hours_run = calcRun(row.opening_hours, row.closing_hours);

    // Apply carry-forward lock from open breakdown
    const bd = openBreakdownByAsset.get(row.asset_code);
    if (bd) {
      row.is_down = true;
      row.down_lock = true;
      row.down_reason = parseDownReasonFromDesc(bd.description);
      row.lock_wo_status = bd.primary_work_order_status || "";
      row.breakdown_start_date = String(bd.breakdown_date || bd.start_at || "").trim();
      row.parts_ordered_date = String(bd.parts_ordered_date || "").trim();
      row.parts_status = String(bd.parts_status || "").trim();
      row.parts_received_date = String(bd.parts_received_date || "").trim();
      row.ets_repair_date = String(bd.ets_repair_date || "").trim();
      const downForDate = Number(bd.hours_down_for_date);
      row.down_hours = Number.isFinite(downForDate) && downForDate >= 0
        ? downForDate
        : Number(row.scheduled_hours || 0);
      if (row.opening_hours != null) {
        row.closing_hours = row.opening_hours;
        row.hours_run = 0;
      }
    }

    if (row.is_down && !String(row.breakdown_start_date || "").trim()) {
      row.breakdown_start_date = date;
    }

    dailyRows.push(row);
  }

  validateDailyRows();
  renderDailyPrestartSection();
  renderDailyTable();
  renderDailyPreview();
  setStatus("Daily input loaded.");
}

/* -------- Copy Yesterday + Bulk Scheduled -------- */

async function copyYesterdayToToday() {
  const today = qs("date")?.value || todayLocalYmd();
  const y = prevDateStr(today);

  setStatus(`Copying from ${y}...`);

  let yRows = [];
  try {
    yRows = await fetchJson(`${API}/api/hours/${y}`);
  } catch {
    yRows = [];
  }

  const yByCode = new Map();
  for (const r of yRows) yByCode.set(r.asset_code, r);

  for (const r of dailyRows) {
    const yr = yByCode.get(r.asset_code);
    if (!yr) continue;

    if (r.is_master_standby) {
      r.is_used = false;
      r.scheduled_hours = 0;
      r.opening_hours = null;
      r.closing_hours = null;
      r.hours_run = 0;
      r.is_down = false;
      r.down_reason = "";
      r.down_lock = false;
      continue;
    }

    r.scheduled_hours = toNum(yr.scheduled_hours) ?? r.scheduled_hours ?? 0;
    r.input_unit = String(yr.input_unit || r.input_unit || "hours").toLowerCase() === "km" ? "km" : "hours";

    const yClose = toNum(yr.closing_hours);
    if (yClose != null) r.opening_hours = yClose;

    r.closing_hours = null;
    r.hours_run = 0;

    r.is_used = !!yr.is_used;
    // Preserve carry-forward lock state from open breakdowns
    if (!r.down_lock) {
      r.is_down = false;
      r.down_reason = "";
    }
  }

  validateDailyRows();
  renderDailyTable();
  renderDailyPreview();
  setStatus(`Copied yesterday (${y}) ✅`);
}

function applyBulkScheduled() {
  const v = toNum(qs("bulkSched")?.value);
  if (v == null || v < 0 || v > 24) {
    alert("Bulk scheduled must be between 0 and 24.");
    return;
  }

  for (const r of dailyRows) {
    if (r.is_master_standby) continue;
    if (!r.is_used) continue;
    r.scheduled_hours = v;
  }

  validateDailyRows();
  renderDailyTable();
  renderDailyPreview();
  setStatus(`Bulk scheduled applied: ${v}h`);
}

async function saveDailyInput() {
  const date = qs("date")?.value || todayLocalYmd();

  validateDailyRows();
  renderDailyPreview();

  const errors = dailyRows.filter((r) => r.error);
  if (errors.length) {
    setText(
      "dailyResult",
      "Cannot save yet. Fix these rows first:\n\n" +
        errors
          .slice(0, 30)
          .map((e) => `${e.asset_code}: ${e.error}`)
          .join("\n") +
        (errors.length > 30 ? `\n...and ${errors.length - 30} more` : "") +
        "\n\nTips:\n- Use KM unit for vehicle distance rows.\n- Standby rows must have run = 0.\n- Production rows must have scheduled > 0."
    );
    setStatus("Save blocked: fix errors.");
    // focus the errors so the user can actually see them
const out = qs("dailyResult");
if (out) {
  out.scrollIntoView({ behavior: "smooth", block: "start" });
}
renderDailyTable(); // re-render so errorRow highlighting appears
    return;
  }

  setStatus("Saving daily input...");
  setText("dailyResult", "");

  const results = [];
  for (const r of dailyRows) {
    if (r.is_down && r.opening_hours != null) {
      r.closing_hours = r.opening_hours;
      r.hours_run = 0;
    }

    const payload = {
      asset_code: r.asset_code,
      work_date: date,
      is_used: r.is_used,
      input_unit: String(r.input_unit || "hours").toLowerCase(),
      scheduled_hours: r.scheduled_hours ?? 0,
      opening_hours: r.opening_hours,
      closing_hours: r.closing_hours,
      hours_run: r.hours_run,
      notes: r.is_down ? (String(r.breakdown_comment || "").trim() || null) : null,
    };

    try {
      const res = await postHoursWithOffline(payload);
      if (res && res.queued) results.push({ asset_code: r.asset_code, ok: true, queued: true });
      else results.push({ asset_code: r.asset_code, ok: true, res });
    } catch (e) {
      results.push({ asset_code: r.asset_code, ok: false, error: e.message || String(e) });
    }
  }

  for (const r of dailyRows) {
    if (!r.is_down) continue;
    await logDownRowToBreakdowns(date, r);
  }

  const failed = results.filter((x) => !x.ok);
  const queued = results.filter((x) => x.queued).length;

  setText("dailyResult", JSON.stringify({ saved: results.length - failed.length, failed }, null, 2));

  if (failed.length) setStatus(`Saved with issues: ${failed.length} failed.`);
  else if (queued) setStatus(`Saved offline: ${queued} queued for sync ✅`);
  else setStatus("Saved successfully.");

  refreshNetBanner();

  await loadDashboard().catch(() => {});
  await loadDailyInput().catch(() => {});
}

async function buildAssetQrImageData(assetCode) {
  const code = String(assetCode || "").trim();
  if (!code) throw new Error("Asset code is required.");
  const res = await fetchJson(`${API}/api/assets/${encodeURIComponent(code)}/qr-profile/refresh`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({}),
  });
  const qrText = String(res?.qr_text || "").trim();
  const scanValue = String(res?.qr_payload?.scan_url || qrText || "").trim();
  if (!scanValue) throw new Error("No QR value generated.");
  const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=420x420&data=${encodeURIComponent(scanValue)}`;
  return { qrUrl, qrText, scanValue, payload: res?.qr_payload || {} };
}

async function downloadAssetQrPng(assetCode) {
  const code = String(assetCode || "").trim();
  if (!code) {
    alert("Asset code missing.");
    return;
  }
  setStatus(`Preparing QR image for ${code}...`);
  const { qrUrl } = await buildAssetQrImageData(code);
  const response = await fetch(qrUrl);
  if (!response.ok) throw new Error(`QR image fetch failed (${response.status})`);
  const blob = await response.blob();
  const objUrl = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = objUrl;
  a.download = `${code}_qr.png`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(objUrl);
  setStatus(`QR image downloaded for ${code} ✅`);
}

async function downloadAllVisibleDailyQrs() {
  if (typeof window.JSZip !== "function") {
    alert("ZIP library failed to load. Check internet connection and try again.");
    return;
  }
  const rowsToUse = dailyShowDownOnly ? dailyRows.filter((r) => !!r.is_down) : dailyRows;
  const codes = Array.from(new Set(rowsToUse.map((r) => String(r.asset_code || "").trim()).filter(Boolean)));
  if (!codes.length) {
    alert("No visible assets to export.");
    return;
  }

  setStatus(`Preparing ${codes.length} visible QR images for ZIP...`);
  const zip = new window.JSZip();
  let ok = 0;
  let fail = 0;
  for (let i = 0; i < codes.length; i += 1) {
    const code = codes[i];
    try {
      const { qrUrl } = await buildAssetQrImageData(code);
      const response = await fetch(qrUrl);
      if (!response.ok) throw new Error(`fetch ${response.status}`);
      const blob = await response.blob();
      zip.file(`${code}_qr.png`, blob);
      ok += 1;
      setStatus(`Collecting QR ${i + 1}/${codes.length}: ${code}`);
      await new Promise((resolve) => setTimeout(resolve, 120));
    } catch {
      fail += 1;
    }
  }
  if (!ok) throw new Error("No QR images could be added to ZIP.");

  setStatus("Building ZIP file...");
  const zipBlob = await zip.generateAsync({ type: "blob" });
  const dt = new Date();
  const stamp = `${dt.getFullYear()}${String(dt.getMonth() + 1).padStart(2, "0")}${String(dt.getDate()).padStart(2, "0")}_${String(dt.getHours()).padStart(2, "0")}${String(dt.getMinutes()).padStart(2, "0")}`;
  const objUrl = URL.createObjectURL(zipBlob);
  const a = document.createElement("a");
  a.href = objUrl;
  a.download = `ironlog_visible_qr_${stamp}.zip`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(objUrl);
  setStatus(`QR ZIP ready: ${ok} included${fail ? `, ${fail} failed` : ""} ✅`);
}

async function printVisibleDailyQrSheet() {
  const rowsToUse = dailyShowDownOnly ? dailyRows.filter((r) => !!r.is_down) : dailyRows;
  const codes = Array.from(new Set(rowsToUse.map((r) => String(r.asset_code || "").trim()).filter(Boolean)));
  if (!codes.length) {
    alert("No visible assets to print.");
    return;
  }

  setStatus(`Building printable QR sheet for ${codes.length} asset(s)...`);
  const labels = [];
  for (let i = 0; i < codes.length; i += 1) {
    const code = codes[i];
    try {
      const { qrUrl } = await buildAssetQrImageData(code);
      labels.push({ code, qrUrl });
      setStatus(`Preparing label ${i + 1}/${codes.length}: ${code}`);
      await new Promise((resolve) => setTimeout(resolve, 80));
    } catch {
      // skip bad rows but continue
    }
  }

  if (!labels.length) throw new Error("Could not prepare any QR labels.");

  openQrLabelSheetPrintWindow(labels, readQrSheetLayout("daily"), "IRONLOG QR Label Sheet");
  setStatus(`Printable QR sheet ready (${labels.length} labels) ✅`);
}

const QR_SHEET_PRESETS = {
  small: { cols: 5, qr: 22, cell: 32, gap: 2 },
  medium: { cols: 4, qr: 28, cell: 45, gap: 4 },
  large: { cols: 3, qr: 35, cell: 55, gap: 5 },
  avery_3474: { cols: 4, qr: 23, cell: 34, gap: 2 },
  avery_l7163: { cols: 2, qr: 35, cell: 43, gap: 4 },
};

function qrSheetFieldIds(scope) {
  if (scope === "safety") {
    return {
      preset: "safetyQrPreset",
      cols: "safetyQrCols",
      qr: "safetyQrSizeMm",
      cell: "safetyQrCellMm",
      gap: "safetyQrGapMm",
    };
  }
  return {
    preset: "qrPreset",
    cols: "qrCols",
    qr: "qrSizeMm",
    cell: "qrCellMm",
    gap: "qrGapMm",
  };
}

function readQrSheetLayout(scope) {
  const ids = qrSheetFieldIds(scope);
  const defaults = scope === "safety"
    ? { cols: 5, qr: 22, cell: 32, gap: 2 }
    : { cols: 4, qr: 28, cell: 45, gap: 4 };
  return {
    cols: Math.max(1, Math.min(8, Number(toNum(qs(ids.cols)?.value) ?? defaults.cols))),
    qrSizeMm: Math.max(12, Math.min(60, Number(toNum(qs(ids.qr)?.value) ?? defaults.qr))),
    cellMm: Math.max(20, Math.min(80, Number(toNum(qs(ids.cell)?.value) ?? defaults.cell))),
    gapMm: Math.max(0, Math.min(20, Number(toNum(qs(ids.gap)?.value) ?? defaults.gap))),
  };
}

function openQrLabelSheetPrintWindow(labels, layout, sheetTitle) {
  const safeLabels = Array.isArray(labels) ? labels.filter((l) => l?.qrUrl && l?.code) : [];
  if (!safeLabels.length) throw new Error("No QR labels to print.");
  const { cols, qrSizeMm, cellMm, gapMm } = layout || readQrSheetLayout("daily");
  const win = window.open("", "_blank", "width=1100,height=800");
  if (!win) {
    alert("Pop-up blocked. Allow pop-ups and try again.");
    return;
  }
  const cells = safeLabels
    .map(
      (l) => `
      <div class="cell">
        <img src="${l.qrUrl}" alt="${escapeHtml(l.code)} QR" />
        <div class="code">${escapeHtml(l.code)}</div>
      </div>
    `
    )
    .join("");
  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>${escapeHtml(sheetTitle || "IRONLOG QR Label Sheet")}</title>
  <style>
    @page { size: A4 portrait; margin: 8mm; }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: Arial, sans-serif; color: #111; }
    .sheet { padding: 8mm; }
    .head { margin-bottom: 6mm; font-size: 12px; }
    .grid {
      display: grid;
      grid-template-columns: repeat(${cols}, 1fr);
      gap: ${Math.max(1, Math.round(gapMm * 1.2))}mm ${gapMm}mm;
    }
    .cell {
      border: 1px solid #bbb;
      border-radius: 4px;
      padding: 3mm 2mm;
      text-align: center;
      min-height: ${cellMm}mm;
      break-inside: avoid;
    }
    .cell img {
      width: ${qrSizeMm}mm;
      height: ${qrSizeMm}mm;
      image-rendering: pixelated;
      display: block;
      margin: 0 auto 2mm;
    }
    .code {
      font-size: 10px;
      font-weight: 700;
      letter-spacing: 0.2px;
      overflow-wrap: anywhere;
    }
    @media print {
      .no-print { display: none; }
      body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    }
  </style>
</head>
<body>
  <div class="sheet">
    <div class="head">${escapeHtml(sheetTitle || "IRONLOG QR Label Sheet")} | Total: ${safeLabels.length} | Layout: ${cols} cols, QR ${qrSizeMm}mm, Cell ${cellMm}mm, Gap ${gapMm}mm | Generated: ${new Date().toISOString()}</div>
    <div class="grid">${cells}</div>
    <div class="no-print" style="margin-top:10px;font-size:12px;color:#555;">Use browser print scaling at 100% for label alignment.</div>
  </div>
  <script>window.onload = () => { window.focus(); window.print(); };</script>
</body>
</html>`;
  win.document.open();
  win.document.write(html);
  win.document.close();
}

function applyQrSheetPresetScope(scope) {
  const ids = qrSheetFieldIds(scope);
  const preset = String(qs(ids.preset)?.value || "custom");
  const target = QR_SHEET_PRESETS[preset];
  if (!target) return;
  const setVal = (id, value) => {
    const el = qs(id);
    if (el) el.value = String(value);
  };
  setVal(ids.cols, target.cols);
  setVal(ids.qr, target.qr);
  setVal(ids.cell, target.cell);
  setVal(ids.gap, target.gap);
}

function applyQrSheetPreset() {
  applyQrSheetPresetScope("daily");
}

function applySafetyQrSheetPreset() {
  applyQrSheetPresetScope("safety");
}

async function generateDailyAssetQr() {
  const assetCode = String(qs("dailyQrAssetCode")?.value || "").trim();
  if (!assetCode) {
    alert("Enter/select an asset code first.");
    return;
  }
  setStatus(`Generating QR for ${assetCode}...`);
  const { qrUrl, qrText, payload } = await buildAssetQrImageData(assetCode);
  const img = qs("dailyQrImg");
  if (img) img.src = qrUrl;
  const preview = qs("dailyQrPreview");
  if (preview) preview.style.display = "block";

  const service = payload?.next_service_due;
  const out = {
    asset_code: payload?.asset?.asset_code || assetCode,
    status: payload?.status || "UNKNOWN",
    next_service_due: service
      ? `${service.service_name} @ ${service.next_due_hours}h (${service.remaining_hours}h remaining)`
      : "No active maintenance plan",
    fuel_liters_last_30_days: payload?.fuel?.liters_last_30_days ?? 0,
    last_inspection_date: payload?.inspections?.last_inspection_date || null,
    generated_at: payload?.generated_at || new Date().toISOString(),
    qr_text: qrText,
  };
  setText("dailyQrText", JSON.stringify(out, null, 2));
  setStatus(`QR saved for ${assetCode} ✅`);
}

function printDailyAssetQr() {
  const imgSrc = String(qs("dailyQrImg")?.src || "").trim();
  const payloadText = String(qs("dailyQrText")?.textContent || "").trim();
  if (!imgSrc || !payloadText) {
    alert("Generate a QR first, then print.");
    return;
  }

  let payload = null;
  try {
    payload = JSON.parse(payloadText);
  } catch {
    payload = null;
  }
  if (!payload) {
    alert("QR payload is invalid. Generate the QR again.");
    return;
  }

  const machine = String(payload.asset_code || "Unknown");
  const status = String(payload.status || "UNKNOWN");
  const nextService = String(payload.next_service_due || "No active maintenance plan");
  const fuel = `${Number(payload.fuel_liters_last_30_days || 0).toFixed(1)} L (30 days)`;
  const inspection = String(payload.last_inspection_date || "No inspection date");
  const generated = String(payload.generated_at || "");

  const win = window.open("", "_blank", "width=900,height=700");
  if (!win) {
    alert("Pop-up blocked. Allow pop-ups and try again.");
    return;
  }

  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>IRONLOG QR - ${machine}</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; color: #111; }
    .sheet { border: 1px solid #222; border-radius: 10px; padding: 18px; max-width: 760px; }
    h1 { margin: 0 0 10px; font-size: 22px; }
    .meta { margin: 0 0 14px; font-size: 14px; }
    .grid { display: grid; grid-template-columns: 240px 1fr; gap: 16px; align-items: start; }
    img { width: 220px; height: 220px; border: 1px solid #999; }
    .row { margin: 0 0 8px; font-size: 14px; }
    .label { font-weight: 700; }
    .raw { margin-top: 14px; padding: 10px; border: 1px dashed #999; white-space: pre-wrap; font-size: 12px; }
    @media print {
      body { margin: 0; }
      .sheet { border: 0; border-radius: 0; padding: 8mm; max-width: none; }
    }
  </style>
</head>
<body>
  <div class="sheet">
    <h1>IRONLOG Machine QR</h1>
    <div class="meta">Generated: ${generated || new Date().toISOString()}</div>
    <div class="grid">
      <div><img src="${imgSrc}" alt="Machine QR code" /></div>
      <div>
        <div class="row"><span class="label">Machine:</span> ${machine}</div>
        <div class="row"><span class="label">Status:</span> ${status}</div>
        <div class="row"><span class="label">Next service:</span> ${nextService}</div>
        <div class="row"><span class="label">Fuel used:</span> ${fuel}</div>
        <div class="row"><span class="label">Last inspection:</span> ${inspection}</div>
      </div>
    </div>
    <div class="raw">${String(payload.qr_text || "").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</div>
  </div>
  <script>window.onload = () => { window.focus(); window.print(); };</script>
</body>
</html>`;

  win.document.open();
  win.document.write(html);
  win.document.close();
}

async function runShiftSelfCheck() {
  const date = qs("date")?.value || todayLocalYmd();
  const out = qs("shiftSelfCheckResult");
  if (out) out.textContent = "Running checks...";
  const checks = [];
  async function step(name, fn) {
    try {
      await fn();
      checks.push({ name, ok: true, msg: "OK" });
    } catch (e) {
      checks.push({ name, ok: false, msg: e.message || String(e) });
    }
  }

  await step("API health", async () => {
    const r = await fetchJson(`${API}/health`);
    if (!r?.ok) throw new Error("Health endpoint not OK");
  });
  await step("Daily rows load", async () => {
    const r = await fetchJson(`${API}/api/hours/${date}`);
    if (!Array.isArray(r)) throw new Error("Daily rows response invalid");
  });
  await step("Dashboard KPI load", async () => {
    const r = await fetchJson(`${API}/api/dashboard?date=${date}&scheduled=${qs("scheduled")?.value || 10}`);
    if (!r?.kpi) throw new Error("Missing KPI block");
  });
  await step("Dispatch KPI load", async () => {
    const r = await fetchJson(`${API}/api/dispatch/kpi?from=${encodeURIComponent(date)}&to=${encodeURIComponent(date)}`);
    if (!r?.ok) throw new Error("Dispatch KPI not OK");
  });
  await step("Operations load", async () => {
    const r = await fetchJson(`${API}/api/operations?from=${encodeURIComponent(date)}&to=${encodeURIComponent(date)}`);
    if (!r?.ok) throw new Error("Operations response not OK");
  });

  const okCount = checks.filter((c) => c.ok).length;
  const failCount = checks.length - okCount;
  const lines = [
    `Shift Self-Check (${date})`,
    `Passed: ${okCount} | Failed: ${failCount}`,
    "",
    ...checks.map((c) => `${c.ok ? "PASS" : "FAIL"} - ${c.name}: ${c.msg}`),
  ];
  if (out) out.textContent = lines.join("\n");
  setStatus(failCount ? `Self-check finished with ${failCount} failure(s).` : "Self-check passed.");
}

function exportShiftSelfCheckTxt() {
  const date = qs("date")?.value || todayLocalYmd();
  const content = String(qs("shiftSelfCheckResult")?.textContent || "").trim();
  if (!content) {
    alert("Run Shift Self-Check first, then export.");
    return;
  }
  const exportedBy = getSessionUser();
  const exportedRole = getSessionRole();
  const exportedAt = new Date().toISOString();
  const header = [
    "IRONLOG Shift Self-Check Export",
    `Shift date: ${date}`,
    `Exported by: ${exportedBy}`,
    `Role: ${exportedRole}`,
    `Exported at: ${exportedAt}`,
    "",
  ].join("\n");
  const blob = new Blob([header + content + "\n"], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `IRONLOG_ShiftSelfCheck_${date}.txt`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  setStatus("Shift self-check TXT exported.");
}

/* =========================
   ASSETS TAB (History + Archive)
========================= */

function normBool(v) {
  return v === true || v === 1 || v === "1";
}

function isHiredAsset(a) {
  const c = String(a?.category || "").toLowerCase();
  return c.includes("contractor hire") || c.includes("contractor") || c.includes("hire");
}

function isKmFuelAsset(a) {
  const mode = String(a?.utilization_mode || "").trim().toLowerCase();
  if (mode === "km") return true;
  if (mode === "hours") return false;
  const cat = String(a?.category || "").toLowerCase();
  const code = String(a?.asset_code || "").toUpperCase();
  const name = String(a?.asset_name || "").toLowerCase();
  if (code.startsWith("BMP")) return true;
  if (/^V\d{2}AM$/.test(code) || /^T\d{2}AM$/.test(code) || code.startsWith("PTT") || code.startsWith("LDV")) return true;
  const keys = ["truck", "vehicle", "ldv", "pickup", "bakkie", "tipper", "dump", "haul", "spinner"];
  if (keys.some((k) => cat.includes(k) || name.includes(k))) return true;
  if (name.includes("toyota") && name.includes("hilux")) return true;
  return false;
}

function makeContractorAssetCode() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `HIRE-${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

async function saveContractorAsset() {
  const contractor = String(qs("caContractor")?.value || "").trim();
  const codeInput = String(qs("caCode")?.value || "").trim().toUpperCase();
  const assetName = String(qs("caName")?.value || "").trim();
  const categoryInput = String(qs("caCategory")?.value || "").trim() || "Contractor Hire";
  const isStandby = !!qs("caStandby")?.checked;
  const out = qs("contractorAssetResult");

  if (!assetName) {
    alert("Enter contractor asset name / unit.");
    return;
  }

  const asset_code = codeInput || makeContractorAssetCode();
  const category = contractor ? `${categoryInput} (${contractor})` : categoryInput;
  const cost_center_code = String(qs("caCostCenter")?.value || "").trim() || undefined;
  const payload = {
    asset_code,
    asset_name: assetName,
    category,
    active: 1,
    is_standby: isStandby ? 1 : 0,
    site_code: getSessionSite() || "main",
    cost_center_code,
  };

  setStatus(`Adding contractor asset ${asset_code}...`);
  try {
    const res = await fetchJson(`${API}/api/assets`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (out) out.textContent = JSON.stringify({ ok: true, asset_code, id: res?.id || null }, null, 2);
    if (isKmFuelAsset({ category, asset_code, asset_name: assetName })) {
      await fetchJson(`${API}/api/dashboard/cost/asset-rates`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ asset_code, utilization_mode: "km" }),
      }).catch(() => {});
    }
    const hireMode = String(qs("caHireBilling")?.value || "").trim();
    if (hireMode) {
      await fetchJson(`${API}/api/assets/${encodeURIComponent(asset_code)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          hire_billing_mode: hireMode,
          hire_rate_per_hour: qs("caHireRateHour")?.value || null,
          hire_fixed_monthly: qs("caHireFixedMonthly")?.value || null,
        }),
      }).catch(() => {});
    }
    if (qs("caCode")) qs("caCode").value = "";
    if (qs("caName")) qs("caName").value = "";
    if (qs("caStandby")) qs("caStandby").checked = false;
    setStatus(`Contractor asset ${asset_code} added.`);
    await Promise.all([
      loadAssetsFleet().catch(() => {}),
      loadCodePickers().catch(() => {}),
      loadDashboard().catch(() => {}),
      loadPlantHirePanel().catch(() => {}),
    ]);
    await selectAssetCard(asset_code, { loadHistory: false, scroll: true }).catch(() => {});
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Failed to add contractor asset.");
  }
}

let assetsFleetCache = [];
let assetsSelectedCode = "";

function getSelectedAssetCode() {
  return String(assetsSelectedCode || qs("histAsset")?.value || "").trim();
}

function syncAssetsArchiveLabel() {
  const label = qs("assetsArchiveSelectedLabel");
  if (!label) return;
  const code = getSelectedAssetCode();
  label.textContent = code || "—";
}

function fleetStatusPill(status) {
  const st = String(status || "UNKNOWN").toUpperCase();
  if (st === "PRODUCTION") return "<span class='pill green'>PROD</span>";
  if (st === "STANDBY") return "<span class='pill orange'>STBY</span>";
  if (st === "DOWN") return "<span class='pill red'>DOWN</span>";
  return "<span class='pill'>—</span>";
}

function ensureAssetHistoryDateRange() {
  const endEl = qs("histEnd");
  const startEl = qs("histStart");
  const today = todayLocalYmd();
  if (endEl && !endEl.value) endEl.value = today;
  if (startEl && !startEl.value) {
    const d = new Date();
    d.setMonth(d.getMonth() - 12);
    startEl.value = d.toISOString().slice(0, 10);
  }
}

function renderAssetFleetGrid(cards) {
  const grid = qs("assetFleetGrid");
  if (!grid) return;
  const filter = String(qs("assetsFleetFilter")?.value || "").trim().toLowerCase();
  const filtered = (cards || []).filter((c) => {
    if (!filter) return true;
    const hay = `${c.asset_code} ${c.asset_name} ${c.category || ""}`.toLowerCase();
    return hay.includes(filter);
  });

  grid.innerHTML = "";
  if (!filtered.length) {
    grid.innerHTML = `<div class="muted small">${cards?.length ? "No machines match your search." : "No assets found."}</div>`;
    return;
  }

  for (const c of filtered) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "asset-fleet-card";
    if (Number(c.archived)) btn.classList.add("archived");
    if (c.asset_code === assetsSelectedCode) btn.classList.add("selected");
    btn.dataset.assetCode = c.asset_code;
    const hired = isHiredAsset(c);
    btn.innerHTML = `
      <div class="asset-fleet-code">${escapeHtml(c.asset_code)}${hired ? " <span class='pill' style='font-size:0.65rem;'>HIRED</span>" : ""}${Number(c.archived) ? " <span class='pill red' style='font-size:0.65rem;'>ARCH</span>" : ""}</div>
      <div class="asset-fleet-name">${escapeHtml(c.asset_name || "")}</div>
      <div class="asset-fleet-meta">
        ${fleetStatusPill(c.status)}
        ${c.category ? `<span class="pill">${escapeHtml(String(c.category))}</span>` : ""}
        <span class="pill blue">${Number(c.current_hours || 0).toFixed(0)}h</span>
        ${Number(c.fuel_liters_30d) > 0 ? `<span class="pill green">${Number(c.fuel_liters_30d).toFixed(0)}L/30d</span>` : ""}
      </div>
    `;
    grid.appendChild(btn);
  }
}

function downloadAssetsCostCentersXlsx() {
  const includeArchived = qs("showArchived")?.checked ? 1 : 0;
  window.open(
    `${API}/api/assets/cost-centers.xlsx?include_archived=${includeArchived}&_ts=${Date.now()}`,
    "_blank",
  );
  setStatus("Asset cost center register export started.");
}

let plantHireRegisterCache = [];

function fillPlantHireRateFields(row) {
  if (qs("plantHireBillingMode")) {
    qs("plantHireBillingMode").value = String(row?.hire_billing_mode || "");
  }
  if (qs("plantHireRateHour")) {
    qs("plantHireRateHour").value =
      row?.hire_rate_per_hour != null && row.hire_rate_per_hour !== "" ? String(row.hire_rate_per_hour) : "";
  }
  if (qs("plantHireFixedMonthly")) {
    qs("plantHireFixedMonthly").value =
      row?.hire_fixed_monthly != null && row.hire_fixed_monthly !== "" ? String(row.hire_fixed_monthly) : "";
  }
}

function syncPlantHireAssetLabel(code) {
  const label = qs("plantHireAssetLabel");
  if (!label) return;
  const c = String(code || "").trim();
  if (!c) {
    label.textContent = "—";
    return;
  }
  const row = plantHireRegisterCache.find((r) => r.asset_code === c);
  label.textContent = row ? `${c} — ${row.asset_name || ""}` : c;
}

async function loadPlantHirePanel() {
  const monthEl = qs("plantHireBudgetMonth");
  if (monthEl && !monthEl.value) {
    monthEl.value = (qs("costMonth")?.value || "").trim() || new Date().toISOString().slice(0, 7);
  }
  try {
    const data = await fetchJson(`${API}/api/assets/hire-register`);
    plantHireRegisterCache = Array.isArray(data?.rows) ? data.rows : [];
    const sel = qs("plantHireAssetSelect");
    if (sel) {
      const prev = String(sel.value || getSelectedAssetCode() || "");
      sel.innerHTML = `<option value="">— Select hired asset —</option>`;
      plantHireRegisterCache.forEach((r) => {
        const opt = document.createElement("option");
        opt.value = r.asset_code;
        const mode = String(r.utilization_mode || "").trim().toLowerCase();
        const modeLbl = mode === "km" ? "km" : mode === "hours" ? "hrs" : "?";
        const arch = Number(r.archived) ? " [arch]" : "";
        opt.textContent = `${r.asset_code} — ${r.asset_name || ""} (${modeLbl}${arch})`;
        sel.appendChild(opt);
      });
      const pick = prev && plantHireRegisterCache.some((r) => r.asset_code === prev) ? prev : "";
      if (pick) sel.value = pick;
    }
    const code = String(qs("plantHireAssetSelect")?.value || getSelectedAssetCode() || "").trim();
    if (code) {
      const row = plantHireRegisterCache.find((r) => r.asset_code === code);
      fillPlantHireRateFields(row || {});
      syncPlantHireAssetLabel(code);
    }
    await loadPlantHireBudgetStatus().catch(() => {});
  } catch (e) {
    const out = qs("plantHireRatesResult");
    if (out) out.textContent = String(e.message || e);
  }
}

async function loadPlantHireBudgetStatus() {
  const month = (qs("plantHireBudgetMonth")?.value || "").trim();
  if (!month) return;
  const site = encodeURIComponent(getSessionSite() || "main");
  const opStatus = qs("operatingBudgetStatus");
  const phStatus = qs("plantHireBudgetStatus");
  try {
    const [opRes, phRes] = await Promise.all([
      fetchJson(`${API}/api/finance/operating-budget?period=${encodeURIComponent(month)}&site_code=${site}`),
      fetchJson(`${API}/api/finance/plant-hire-budget?period=${encodeURIComponent(month)}&site_code=${site}`),
    ]);
    const opAmt = Number(opRes?.budget_amount || 0);
    const phAmt = Number(phRes?.budget_amount || 0);
    if (qs("operatingBudgetAmount") && document.activeElement !== qs("operatingBudgetAmount")) {
      qs("operatingBudgetAmount").value = opAmt > 0 ? String(opAmt) : "";
    }
    if (qs("plantHireBudgetAmount") && document.activeElement !== qs("plantHireBudgetAmount")) {
      qs("plantHireBudgetAmount").value = phAmt > 0 ? String(phAmt) : "";
    }
    if (opStatus) {
      let opText = opAmt > 0
        ? `Saved operating budget for ${month}: ${fmtMoney(opAmt)} (one total for all expense categories)`
        : `No operating budget saved for ${month} yet — enter your total monthly expenses above.`;
      if (opAmt <= 0 && phAmt > 0) {
        opText += ` If you entered total expenses under plant hire by mistake, copy that amount here and save.`;
      }
      opStatus.textContent = opText;
    }
    if (phStatus) {
      phStatus.textContent = phAmt > 0
        ? `Saved plant hire income target for ${month}: ${fmtMoney(phAmt)}`
        : `No plant hire income target for ${month} (optional — contractor plant income only).`;
    }
  } catch (e) {
    if (opStatus) opStatus.textContent = String(e.message || e);
    if (phStatus) phStatus.textContent = String(e.message || e);
  }
}

async function saveOperatingBudget() {
  const period = (qs("plantHireBudgetMonth")?.value || "").trim();
  const budget_amount = Number(qs("operatingBudgetAmount")?.value || 0);
  if (!period) return alert("Select a budget month.");
  if (!Number.isFinite(budget_amount) || budget_amount < 0) return alert("Enter a valid budget amount.");
  setStatus("Saving operating budget...");
  try {
    const res = await fetchJson(`${API}/api/finance/operating-budget`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        period,
        site_code: getSessionSite() || "main",
        budget_amount,
      }),
    });
    setStatus(`Operating budget saved for ${period}.`);
    if (qs("operatingBudgetStatus")) {
      qs("operatingBudgetStatus").textContent =
        `Saved operating budget for ${period}: ${fmtMoney(res.budget_amount)} (one total for all expense categories)`;
    }
  } catch (e) {
    setStatus("Operating budget save failed.");
    alert(String(e.message || e));
  }
}

async function savePlantHireBudget() {
  const period = (qs("plantHireBudgetMonth")?.value || "").trim();
  const budget_amount = Number(qs("plantHireBudgetAmount")?.value || 0);
  if (!period) return alert("Select a budget month.");
  if (!Number.isFinite(budget_amount) || budget_amount < 0) return alert("Enter a valid budget amount.");
  setStatus("Saving plant hire budget...");
  try {
    const res = await fetchJson(`${API}/api/finance/plant-hire-budget`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        period,
        site_code: getSessionSite() || "main",
        budget_amount,
      }),
    });
    setStatus(`Plant hire income target saved for ${period}.`);
    if (qs("plantHireBudgetStatus")) {
      qs("plantHireBudgetStatus").textContent = `Saved plant hire income target for ${period}: ${fmtMoney(res.budget_amount)}`;
    }
  } catch (e) {
    setStatus("Plant hire budget save failed.");
    alert(String(e.message || e));
  }
}

async function savePlantHireRates() {
  const asset_code = String(qs("plantHireAssetSelect")?.value || getSelectedAssetCode() || "").trim();
  if (!asset_code) return alert("Select a hired asset first.");
  const hire_billing_mode = String(qs("plantHireBillingMode")?.value || "").trim();
  const payload = {
    hire_billing_mode: hire_billing_mode || null,
    hire_rate_per_hour: qs("plantHireRateHour")?.value || null,
    hire_fixed_monthly: qs("plantHireFixedMonthly")?.value || null,
  };
  setStatus(`Saving hire rates for ${asset_code}...`);
  try {
    const res = await fetchJson(`${API}/api/assets/${encodeURIComponent(asset_code)}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const out = qs("plantHireRatesResult");
    if (out) out.textContent = JSON.stringify(res, null, 2);
    setStatus(`Hire rates saved for ${asset_code}.`);
    await loadPlantHirePanel().catch(() => {});
  } catch (e) {
    const out = qs("plantHireRatesResult");
    if (out) out.textContent = String(e.message || e);
    setStatus("Hire rates save failed.");
  }
}

async function loadAssetsFleet() {
  const showArchived = !!qs("showArchived")?.checked;
  const url = `${API}/api/assets/fleet-summary?include_archived=${showArchived ? 1 : 0}`;
  const grid = qs("assetFleetGrid");
  if (grid && !assetsFleetCache.length) {
    grid.innerHTML = `<div class="muted small">Loading fleet…</div>`;
  }
  try {
    const data = await fetchJson(url);
    assetsFleetCache = Array.isArray(data?.cards) ? data.cards : [];
    renderAssetFleetGrid(assetsFleetCache);
    await populateHistoryAssets().catch(() => {});
    await loadPlantHirePanel().catch(() => {});
    if (assetsSelectedCode && !assetsFleetCache.some((c) => c.asset_code === assetsSelectedCode)) {
      assetsSelectedCode = "";
      qs("assetDetailPanel")?.classList.add("hidden");
    } else if (assetsSelectedCode) {
      renderAssetFleetGrid(assetsFleetCache);
    }
    syncAssetsArchiveLabel();
  } catch (e) {
    if (grid) grid.innerHTML = `<div class="muted small">Fleet load error: ${escapeHtml(e.message || e)}</div>`;
    setStatus("Fleet load error: " + (e.message || e));
  }
}

async function loadAssetDetailHeader(asset_code) {
  const header = qs("assetDetailHeader");
  const title = qs("assetDetailTitle");
  const subtitle = qs("assetDetailSubtitle");
  if (!header) return;
  header.innerHTML = `<div class="skeleton-block"></div>`;
  try {
    const data = await fetchJson(`${API}/api/assets/${encodeURIComponent(asset_code)}/qr-profile`);
    const p = data?.live_preview || {};
    const card = assetsFleetCache.find((c) => c.asset_code === asset_code);
    if (title) title.textContent = asset_code;
    if (subtitle) {
      subtitle.textContent = [
        card?.asset_name || p.asset?.asset_name,
        card?.category || p.asset?.category,
      ]
        .filter(Boolean)
        .join(" · ");
    }
    const nextSvc = p.next_service_due;
    const insp = p.inspections?.last_inspection_date;
    const hours = p.meter?.current_hours ?? card?.current_hours ?? 0;
    const fuel30 = p.fuel?.liters_last_30_days ?? card?.fuel_liters_30d ?? 0;
    header.innerHTML = `
      ${fleetStatusPill(p.status || card?.status)}
      <span class="pill blue">${Number(hours).toFixed(1)} h meter</span>
      <span class="pill green">${Number(fuel30).toFixed(1)} L fuel (30d)</span>
      ${
        nextSvc
          ? `<span class="pill orange">Next ${escapeHtml(nextSvc.service_name || "service")}: ${Number(nextSvc.remaining_hours ?? 0).toFixed(0)} h</span>`
          : ""
      }
      ${insp ? `<span class="pill blue">Last inspection ${escapeHtml(String(insp))}</span>` : ""}
      ${Number(card?.archived) ? `<span class="pill red">Archived</span>` : ""}
    `;
  } catch (_) {
    const card = assetsFleetCache.find((c) => c.asset_code === asset_code);
    if (title) title.textContent = asset_code;
    if (subtitle) subtitle.textContent = card?.asset_name || "";
    header.innerHTML = card
      ? `${fleetStatusPill(card.status)} <span class="pill blue">${Number(card.current_hours || 0).toFixed(1)} h</span>`
      : "";
  }
}

async function selectAssetCard(asset_code, opts = {}) {
  const code = String(asset_code || "").trim();
  if (!code) return;
  assetsSelectedCode = code;
  const sel = qs("histAsset");
  if (sel) {
    const exists = Array.from(sel.options).some((o) => o.value === code);
    if (exists) sel.value = code;
  }
  syncAssetsArchiveLabel();
  renderAssetFleetGrid(assetsFleetCache);
  qs("assetDetailPanel")?.classList.remove("hidden");
  if (qs("plantHireAssetSelect")) {
    const row = plantHireRegisterCache.find((r) => r.asset_code === code);
    if (row) {
      qs("plantHireAssetSelect").value = code;
      fillPlantHireRateFields(row);
    }
    syncPlantHireAssetLabel(code);
  }
  ensureAssetHistoryDateRange();
  await loadAssetDetailHeader(code);
  if (opts.loadHistory !== false) {
    await loadAssetHistory().catch((e) => setStatus("History error: " + (e.message || e)));
  }
  if (opts.scroll) {
    qs("assetDetailPanel")?.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

async function populateHistoryAssets() {
  const sel = qs("histAsset");
  if (!sel) return;

  const showArchived = !!qs("showArchived")?.checked;
  const url = `${API}/api/assets?include_archived=${showArchived ? 1 : 0}`;

  let assets = [];
  try {
    assets = await fetchJson(url);
  } catch (e) {
    // don’t blank the dropdown forever
    setStatus("Assets load error: " + (e.message || e));
    return;
  }

  const current = sel.value;
  sel.innerHTML = "";

  if (!assets || !assets.length) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "No assets found";
    sel.appendChild(opt);
    return;
  }

  for (const a of assets) {
    const opt = document.createElement("option");
    opt.value = a.asset_code;

    const archived = normBool(a.archived);
    const hired = isHiredAsset(a);
    const tag = archived ? " (ARCHIVED)" : "";
    const hiredTag = hired ? " [HIRED]" : "";
    opt.textContent = `${a.asset_code}${hiredTag} — ${a.asset_name}${tag}`;

    sel.appendChild(opt);
  }

  const keep = assetsSelectedCode || current;
  if (keep) {
    const exists = Array.from(sel.options).some((o) => o.value === keep);
    if (exists) sel.value = keep;
  }
  syncAssetsArchiveLabel();
}

function formatOpsSlipHistoryExtra(ev) {
  const d = ev.details || {};
  const slipId = Number(d.slip_id || 0);
  const st = String(d.slip_type || "");
  const s = d.summary || {};
  const bits = [];
  if (s.parse_error) bits.push("Summary could not be read from saved slip.");
  else {
    const pic = Number(s.pictures_attached || 0);
    if (pic > 0) bits.push(`${pic} attached picture(s) on PDF slip`);
    if (st === "hose_failure") {
      if (s.hose_part_code) bits.push(`Hose part: ${s.hose_part_code}`);
      if (s.oil_loss_part_code) bits.push(`Oil loss part: ${s.oil_loss_part_code}`);
      if (s.reason) bits.push(String(s.reason));
      if (s.preventable) bits.push("Tagged preventable");
    } else if (st === "get_change") {
      if (s.part_code) bits.push(`Part: ${s.part_code}`);
      if (s.date_changed) bits.push(`Changed: ${s.date_changed}`);
    } else if (st === "component_change") {
      if (s.component_type) bits.push(`Component: ${s.component_type}`);
      if (s.part_code) bits.push(`Part: ${s.part_code}`);
      if (s.reason) bits.push(String(s.reason));
    } else if (st === "tyre_change") {
      if (s.tyre_lines) bits.push(`${s.tyre_lines} tyre line(s)`);
      if (Array.isArray(s.positions) && s.positions.length) bits.push(`Positions: ${s.positions.join(", ")}`);
    }
  }
  const lines = bits.length
    ? bits.map((b) => `<small>${escapeHtml(b)}</small>`).join("<br>")
    : `<small>Operational slip (Breakdown Ops).</small>`;
  const who = d.created_by ? `<br><small>${escapeHtml(String(d.created_by))}</small>` : "";
  const btn = slipId
    ? `<br><button type="button" class="btn small" data-ops-slip-pdf="${slipId}">Open slip PDF</button>`
    : "";
  return `<br>${lines}${who}${btn}`;
}

function pillForType(t) {
  if (t === "breakdown") return "<span class='pill red'>BD</span>";
  if (t === "service") return "<span class='pill blue'>SV</span>";
  if (t === "fuel") return "<span class='pill green'>FUEL</span>";
  if (t === "lube") return "<span class='pill' style='background:#0d9488;color:#fff;'>LUBE</span>";
  if (t === "inspection") return "<span class='pill blue'>INSP</span>";
  if (t === "get_slip") return "<span class='pill orange'>GET</span>";
  if (t === "component_slip") return "<span class='pill orange'>COMP</span>";
  if (t === "damage_report") return "<span class='pill red'>DMG</span>";
  if (t === "tyre_change") return "<span class='pill orange'>TY CHG</span>";
  if (t === "tyre_inspection") return "<span class='pill blue'>TY INSP</span>";
  if (t === "undercarriage_inspection") return "<span class='pill' style='background:#334155;color:#fff;'>UC</span>";
  if (t === "work_order") return "<span class='pill blue'>WO</span>";
  if (t === "ops_slip") return "<span class='pill' style='background:#5b21b6;color:#fff;'>OPS</span>";
  return "<span class='pill'>EV</span>";
}

function normalizeImageSrc(v) {
  const raw = String(v || "").trim();
  if (!raw) return "";
  if (/^https?:\/\//i.test(raw)) return raw;
  if (raw.startsWith("/")) return raw;
  const normalized = raw.replace(/\\/g, "/");
  const lower = normalized.toLowerCase();
  const uploadsIdx = lower.indexOf("/uploads/");
  if (uploadsIdx >= 0) return normalized.slice(uploadsIdx);
  if (/^[a-z]:\//i.test(normalized)) return "";
  return "/" + normalized;
}

function buildPhotoDebugBadge(photoPath) {
  const hasPath = String(photoPath || "").trim().length > 0;
  if (!hasPath) return "<span class='pill orange'>No photo linked</span>";
  return "<span class='pill blue'>Photo path linked</span>";
}

async function loadAssetHistory() {
  const asset_code = getSelectedAssetCode();
  if (!asset_code) return alert("Select a fleet card first.");

  const start = qs("histStart")?.value || "";
  const end = qs("histEnd")?.value || "";

  setStatus(`Loading history for ${asset_code}...`);
  const data = await fetchJson(
    `${API}/api/assets/${encodeURIComponent(asset_code)}/history?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`
  );

  const list = qs("historyList");
  const summaryEl = qs("historySummary");
  if (!list) return;

  list.innerHTML = "";
  if (summaryEl) summaryEl.innerHTML = `<div class="skeleton-block"></div>`;
  const summary = data.summary || {};
  const counts = summary.counts || {};
  const totals = summary.totals || {};
  if (summaryEl) {
    summaryEl.innerHTML = `
      <div class="pill blue">Events: ${Number(counts.events_total || 0)}</div>
      <div class="pill red">Breakdowns: ${Number(counts.breakdowns || 0)}</div>
      <div class="pill blue">Work Orders: ${Number(counts.work_orders || 0)}</div>
      <div class="pill" style="background:#5b21b6;color:#fff;">Ops slips: ${Number(counts.ops_slips || 0)}</div>
      <div class="pill blue">GET Slips: ${Number(counts.get_slips || 0)}</div>
      <div class="pill blue">Component Slips: ${Number(counts.component_slips || 0)}</div>
      <div class="pill red">Damage Reports: ${Number(counts.damage_reports || 0)}</div>
      <div class="pill orange">Tyre Changes: ${Number(counts.tyre_changes || 0)}</div>
      <div class="pill blue">Tyre Inspections: ${Number(counts.tyre_inspections || 0)}</div>
      <div class="pill" style="background:#334155;color:#fff;">Undercarriage: ${Number(counts.undercarriage_inspections || 0)}</div>
      <div class="pill green">Fuel logs: ${Number(counts.fuel_logs || 0)}</div>
      <div class="pill" style="background:#0d9488;color:#fff;">Lube logs: ${Number(counts.lube_logs || 0)}</div>
      <div class="pill blue">Inspections: ${Number(counts.inspections || 0)}</div>
      <div class="pill green">Fuel total: ${Number(totals.fuel_liters_total || 0).toFixed(1)} L</div>
      <div class="pill orange">Parts Qty: ${Number(totals.parts_qty_total || 0).toFixed(1)}</div>
      <div class="pill orange">Oil Qty: ${Number(totals.oil_qty_total || 0).toFixed(1)}</div>
      <div class="pill red">Parts Cost: ${Number(totals.parts_cost_total || 0).toFixed(2)}</div>
      <div class="pill red">Oil Cost: ${Number(totals.oil_cost_total || 0).toFixed(2)}</div>
      <div class="pill red">Maintenance Cost: ${Number(summary.maintenance_cost_total || 0).toFixed(2)}</div>
    `;
  }

  (data.history || []).forEach((ev) => {
    const wo = ev.work_order_id ? ` <small>WO #${ev.work_order_id}</small>` : "";

    const rawPhoto = ev.details?.photo || "";
    const image = normalizeImageSrc(rawPhoto);
    const debugBadge = buildPhotoDebugBadge(rawPhoto);
    const unresolvedWindowsPath = rawPhoto && !image;
    const debugPrefix = unresolvedWindowsPath
      ? `${debugBadge} <span class='pill red'>Windows path unresolved</span>`
      : debugBadge;
    const photoBlock = image
      ? `<div style="margin-top:8px;">${debugPrefix}<br><img src="${image}" alt="event photo" style="max-width:180px; max-height:120px; border-radius:8px; border:1px solid var(--line); object-fit:cover; margin-top:6px;" onload="this.dataset.loaded='1'; this.previousSibling && this.previousSibling.remove && this.previousSibling.remove();" onerror="this.insertAdjacentHTML('beforebegin','<span class=&quot;pill red&quot;>Photo file missing / blocked</span><br>'); this.style.display='none';" /></div>`
      : `<div style="margin-top:8px;">${debugPrefix}</div>`;

    const extra =
      ev.type === "breakdown"
        ? (() => {
            const logs = ev.details?.downtime_logs || [];
            const lines = logs.length
              ? `<div style="margin-top:6px; padding-left:10px; border-left:2px solid #7d2a2a;">
                   ${logs
                     .map(
                       (l) =>
                         `<div><small><b>${l.log_date}</b> — ${l.hours_down}h${
                           l.notes ? ` | ${l.notes}` : ""
                         }</small></div>`
                     )
                     .join("")}
                 </div>`
              : `<br><small>No downtime log lines recorded.</small>`;

            return `<br><small>Downtime total: ${ev.details?.downtime_hours ?? 0}h ${
              ev.details?.critical ? " | CRIT" : ""
            }</small>${lines}${photoBlock}`;
          })()
        : ev.type === "get_slip"
        ? `<br><small>Items: ${(ev.details?.items || []).length}</small>`
        : ev.type === "component_slip"
        ? `<br><small>${ev.details?.serial_out || ""} → ${ev.details?.serial_in || ""}</small>`
        : ev.type === "work_order"
        ? `<br><small>Status: ${ev.details?.status || ""}</small>${photoBlock}`
        : ev.type === "damage_report"
        ? `<br><small>${ev.details?.notes || "No notes recorded."}</small>${photoBlock}`
        : ev.type === "tyre_change"
        ? `<br><small>${ev.details?.serial_out || "-"} → ${ev.details?.serial_in || "-"}</small>${
            ev.details?.hours_at_change != null
              ? `<br><small>Hours at change: ${Number(ev.details.hours_at_change).toFixed(1)}</small>`
              : ""
          }${ev.details?.notes ? `<br><small>${ev.details.notes}</small>` : ""}${photoBlock}`
        : ev.type === "tyre_inspection"
        ? `<br><small>Condition: ${ev.details?.condition || "-"}</small>${
            ev.details?.pressure != null ? `<br><small>Pressure: ${ev.details.pressure}</small>` : ""
          }${
            ev.details?.tread_depth != null ? `<br><small>Tread: ${ev.details.tread_depth}</small>` : ""
          }${ev.details?.notes ? `<br><small>${ev.details.notes}</small>` : ""}${photoBlock}`
        : ev.type === "undercarriage_inspection"
        ? `<br><small>SMU: ${ev.details?.smu ?? "-"} | Max wear: ${ev.details?.worst_wear_pct ?? "-"}%${
            ev.details?.worst_component ? ` (${ev.details.worst_component})` : ""
          }</small>${
            ev.details?.pdf_url
              ? `<br><small><a href="${ev.details.pdf_url}" target="_blank" rel="noopener">Open PDF</a></small>`
              : ""
          }${ev.details?.notes ? `<br><small>${ev.details.notes}</small>` : ""}`
        : ev.type === "ops_slip"
        ? formatOpsSlipHistoryExtra(ev)
        : ev.type === "fuel"
        ? `<br><small>${Number(ev.details?.liters || 0).toFixed(1)} L${ev.details?.source ? ` · ${escapeHtml(String(ev.details.source))}` : ""}</small>`
        : ev.type === "lube"
        ? `<br><small>${Number(ev.details?.quantity || 0).toFixed(1)} L${ev.details?.oil_type ? ` · ${escapeHtml(String(ev.details.oil_type))}` : ""}</small>`
        : ev.type === "inspection"
        ? `<br><small>${ev.details?.inspector ? `Inspector: ${escapeHtml(String(ev.details.inspector))}` : "Inspection recorded"}${ev.details?.notes ? `<br>${escapeHtml(String(ev.details.notes))}` : ""}</small>`
        : "";

    list.appendChild(item(`${pillForType(ev.type)} <b>${ev.date}</b> — ${ev.title}${wo}${extra}`));
  });

  if (!data.history?.length) list.appendChild(item("<small>No history found for this range.</small>"));

  setStatus("History loaded ✅");
}

async function archiveSelectedAsset() {
  const code = getSelectedAssetCode();
  if (!code) return alert("Select a fleet card first.");

  const reason = prompt(`Archive ${code}.\nReason (optional):`, "Scrapped / Not in use");
  if (reason === null) return;

  setStatus(`Archiving ${code}...`);
  await fetchJson(`${API}/api/assets/${encodeURIComponent(code)}/archive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ archived: true, reason: String(reason || "").trim() }),
  });

  setStatus(`Archived ${code} ✅`);
  await loadAssetsFleet().catch(() => {});
  await loadDashboard().catch(() => {});
}

async function unarchiveSelectedAsset() {
  const code = getSelectedAssetCode();
  if (!code) return alert("Select a fleet card first.");

  if (!confirm(`Unarchive ${code}?`)) return;

  setStatus(`Unarchiving ${code}...`);
  await fetchJson(`${API}/api/assets/${encodeURIComponent(code)}/archive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ archived: false }),
  });

  setStatus(`Unarchived ${code} ✅`);
  await loadAssetsFleet().catch(() => {});
  await loadDashboard().catch(() => {});
}

/* =========================
   INIT
========================= */

async function init() {
  await disableLegacyServiceWorkers();
  await tryInitialSession();
  initSidebar();
  initTabs();
  initWorkshopLibraryTab();
  initIronmindHelpUi();
  initSectionCollapseToggles();
  initSessionControls();
  initVehicleCheckTab();
  initSettingsDropdown();
  initGlobalSearch();
  initReportCardCollapsible();
  initTasks();
  initTelematicsFaultBanner();
  initCartrackSpeedFloat();
  applyRoleVisibility();
  if (!isBareChildTabEmbed()) {
    resolveInitialTabFromUrl();
  }
  applyBareChildTabView();
  applyI18n();
  applyGlobalPageTranslation();

  const dateEl = qs("date");
  if (dateEl) dateEl.value = todayLocalYmd();
  const mtdMonthEl = qs("mtdOpeningMonth");
  if (mtdMonthEl && !mtdMonthEl.value) {
    mtdMonthEl.value = (dateEl?.value || todayLocalYmd()).slice(0, 7);
  }

  qs("refresh")?.addEventListener("click", () =>
    loadDashboard().catch((e) => setStatus("Dashboard error: " + e.message))
  );
  qs("kpiDebugToggle")?.addEventListener("change", () =>
    loadDashboard().catch((e) => setStatus("Dashboard error: " + e.message))
  );
  qs("loadReliability")?.addEventListener("click", () =>
    loadDashboard().catch((e) => setStatus("Dashboard error: " + e.message))
  );
  qs("ironmindRefreshBtn")?.addEventListener("click", () =>
    refreshIronmindInsight().catch((e) => setStatus("IRONMIND refresh error: " + e.message))
  );
  setInterval(() => {
    loadIronmindHealth().catch(() => {});
  }, 30000);
  qs("ironmindSaveSettingsBtn")?.addEventListener("click", () =>
    saveIronmindSettings().catch((e) => setStatus("IronMind settings error: " + (e.message || e)))
  );
  qs("ironmindAskBtn")?.addEventListener("click", () =>
    askIronmindQuestion().catch((e) => setStatus("IRONMIND ask error: " + e.message))
  );
  qs("ironmindResetMemoryBtn")?.addEventListener("click", () =>
    resetIronmindAskMemory().catch((e) => setStatus("IRONMIND reset memory error: " + e.message))
  );
  qs("ironmindAskInput")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      askIronmindQuestion().catch((err) => setStatus("IRONMIND ask error: " + err.message));
    }
  });
  hydrateIronmindAskMemory().catch(() => {});
  qs("saveThresholds")?.addEventListener("click", () => saveThresholdsFromUI());
  qs("saveLdvThresholds")?.addEventListener("click", () => saveLdvPrestartThresholdsFromUI());
  qs("ironmindSummary")?.addEventListener("click", (e) => {
    const el = e.target instanceof HTMLElement ? e.target : null;
    if (!el) return;
    const drillKey = el.dataset.ironmindDrill;
    const assetCode = el.dataset.ironmindAsset;
    if (drillKey) ironmindDrillDown(drillKey);
    if (assetCode) ironmindGoToAsset(assetCode).catch(() => {});
  });
  qs("ironmindHistoryList")?.addEventListener("click", (e) => {
    const el = e.target instanceof HTMLElement ? e.target.closest("button[data-ironmind-history-id]") : null;
    if (!el) return;
    const rowEl = el.closest(".item");
    if (!rowEl?.dataset?.ironmindRow) return;
    try {
      const row = JSON.parse(rowEl.dataset.ironmindRow);
      renderIronmindReport(row);
      setStatus(`Opened IRONMIND report for ${row?.report_date || "-"}.`);
    } catch (_) {
      setStatus("Unable to open selected IRONMIND report.");
    }
  });
  qs("riskBoardList")?.addEventListener("click", (e) => {
    const target = e.target instanceof HTMLElement ? e.target : null;
    if (!target) return;
    const openBtn = target.closest("button[data-ironmind-risk-asset]");
    if (openBtn) {
      const code = String(openBtn.getAttribute("data-ironmind-risk-asset") || "").trim();
      if (!code) return;
      ironmindGoToAsset(code).catch(() => {});
      return;
    }
    const woBtn = target.closest("button[data-ironmind-risk-wo]");
    if (woBtn) {
      const code = String(woBtn.getAttribute("data-ironmind-risk-wo") || "").trim();
      if (!code) return;
      (async () => {
        const downDesc = "IRONMIND predicted risk work order";
        const res = await fetchJson(`${API}/api/breakdowns/ensure-open`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            asset_code: code,
            breakdown_date: date,
            description: downDesc,
            critical: false,
          }),
        });
        const woId = Number(res?.primary_work_order_id || 0);
        setStatus(woId > 0 ? `WO #${woId} ready for ${code}.` : `Work order ensured for ${code}.`);
        await loadDashboard().catch(() => {});
      })().catch((err) => setStatus("Create WO failed: " + (err.message || err)));
    }
  });
  qs("ironmindShowMissingDays")?.addEventListener("change", () => {
    loadIronmindHistory({ silent: true }).catch(() => {});
  });
  qs("ironmindReloadHistory")?.addEventListener("click", () => {
    loadIronmindHistory().catch((e) => setStatus("IRONMIND history error: " + e.message));
  });
  qs("ironmindRsgPlanBtn")?.addEventListener("click", () => {
    generateIronmindRsgPlan(false).catch((e) => setStatus("RSG plan error: " + (e.message || e)));
  });
  qs("ironmindRsgPreviewPdfBtn")?.addEventListener("click", () => {
    previewIronmindRsgPdf().catch((e) => setStatus("RSG preview error: " + (e.message || e)));
  });
  qs("ironmindRsgDownloadPdfBtn")?.addEventListener("click", () => {
    downloadIronmindRsgPdf().catch((e) => setStatus("RSG download error: " + (e.message || e)));
  });
  qs("ironmindRsgCreateWoBtn")?.addEventListener("click", () => {
    generateIronmindRsgPlan(true).catch((e) => setStatus("RSG create WO error: " + (e.message || e)));
  });
  qs("saveDocHeaderBtn")?.addEventListener("click", () =>
    saveDocHeader().catch((e) => setStatus("Header save error: " + e.message))
  );
  qs("loadDocHeadersBtn")?.addEventListener("click", () =>
    loadDocHeaders().catch((e) => setStatus("Header load error: " + e.message))
  );
  qs("generateDocDraftBtn")?.addEventListener("click", () =>
    generateDocDraft().catch((e) => setStatus("Draft generate error: " + e.message))
  );
  qs("generateDocDraftFromRequestBtn")?.addEventListener("click", () =>
    generateDocDraftFromRequest().catch((e) => setStatus("Draft request generate error: " + e.message))
  );
  qs("aiSmartRunBtn")?.addEventListener("click", () =>
    runAiSmart().catch((e) => setStatus("Smart AI error: " + e.message))
  );
  qs("askJakesBtn")?.addEventListener("click", () =>
    askJakes().catch((e) => setStatus("Ask Jakes error: " + e.message))
  );
  qs("askJakesPresetHydraulics")?.addEventListener("click", () =>
    applyAskJakesPreset("hydraulics")
  );
  qs("askJakesPresetStarting")?.addEventListener("click", () =>
    applyAskJakesPreset("starting")
  );
  qs("askJakesPresetOverheat")?.addEventListener("click", () =>
    applyAskJakesPreset("overheat")
  );
  qs("askJakesUseAsNotesBtn")?.addEventListener("click", () =>
    useAskJakesAnswerAsNotes()
  );
  qs("speakDocDraftBtn")?.addEventListener("click", () =>
    speakDocDraft()
  );
  qs("stopSpeakDocDraftBtn")?.addEventListener("click", () =>
    stopSpeakingDocDraft()
  );
  qs("docApproveYesBtn")?.addEventListener("click", () =>
    decideDocDraft(true).catch((e) => setStatus("Draft decision error: " + e.message))
  );
  qs("docApproveNoBtn")?.addEventListener("click", () =>
    decideDocDraft(false).catch((e) => setStatus("Draft decision error: " + e.message))
  );
  qs("openDocDraftPdfBtn")?.addEventListener("click", () =>
    openDocDraftPdf(false)
  );
  qs("downloadDocDraftPdfBtn")?.addEventListener("click", () =>
    openDocDraftPdf(true)
  );
  qs("openDocDraftWordBtn")?.addEventListener("click", () =>
    openDocDraftWord(false)
  );
  qs("downloadDocDraftWordBtn")?.addEventListener("click", () =>
    openDocDraftWord(true)
  );
  qs("openDocRegisterPdfBtn")?.addEventListener("click", () =>
    openDocRegisterPdf(false)
  );
  qs("downloadDocRegisterPdfBtn")?.addEventListener("click", () =>
    openDocRegisterPdf(true)
  );
  qs("openDocRegisterWordBtn")?.addEventListener("click", () =>
    openDocRegisterWord(false)
  );
  qs("downloadDocRegisterWordBtn")?.addEventListener("click", () =>
    openDocRegisterWord(true)
  );
  qs("loadDocDraftsBtn")?.addEventListener("click", () =>
    loadDocDrafts().catch((e) => setStatus("Draft list error: " + e.message))
  );
  qs("docDraftsCurrentOnly")?.addEventListener("change", () =>
    loadDocDrafts().catch((e) => setStatus("Draft list error: " + e.message))
  );
  qs("loadLube")?.addEventListener("click", () =>
    loadLubeUsage().catch((e) => setStatus("Lube error: " + e.message))
  );
  qs("lubeFilterAsset")?.addEventListener("input", () => {
    if (lubeUsageCache) renderLubeUsageTable(lubeUsageCache);
  });
  qs("lubeFilterOilType")?.addEventListener("input", () => {
    if (lubeUsageCache) renderLubeUsageTable(lubeUsageCache);
  });
  qs("lubeHidePartLike")?.addEventListener("change", () => {
    if (lubeUsageCache) renderLubeUsageTable(lubeUsageCache);
  });
  qs("loadLubeAnalytics")?.addEventListener("click", () =>
    loadLubeAnalytics().catch((e) => setStatus("Lube analytics error: " + e.message))
  );
  qs("createRequisition")?.addEventListener("click", () =>
    createRequisition().catch((e) => setStatus("Requisition create error: " + e.message))
  );
  qs("prLoadPoList")?.addEventListener("click", () =>
    loadPurchaseOrders().catch((e) => setStatus("PO load error: " + e.message))
  );
  qs("prPoStatusFilter")?.addEventListener("change", () =>
    loadPurchaseOrders().catch((e) => setStatus("PO load error: " + e.message))
  );
  qs("prPostReceiptBtn")?.addEventListener("click", () =>
    postPoReceipt().catch((e) => setStatus("PO receipt error: " + e.message))
  );
  qs("prCaptureInvoiceBtn")?.addEventListener("click", () =>
    capturePoInvoice().catch((e) => setStatus("Invoice capture error: " + e.message))
  );
  qs("prRunMatchBtn")?.addEventListener("click", () =>
    runPoThreeWayMatch().catch((e) => setStatus("3-way match error: " + e.message))
  );
  qs("prLoadExceptionsBtn")?.addEventListener("click", () =>
    loadProcurementExceptions().catch((e) => setStatus("Exception load error: " + e.message))
  );
  qs("prExceptionStatus")?.addEventListener("change", () =>
    loadProcurementExceptions().catch((e) => setStatus("Exception load error: " + e.message))
  );
  qs("prBuildJournalBtn")?.addEventListener("click", () =>
    buildProcurementJournals().catch((e) => setStatus("Journal build error: " + e.message))
  );
  qs("prExportJournalCsvBtn")?.addEventListener("click", () => {
    try {
      exportProcurementJournalsCsv();
    } catch (e) {
      setStatus("Journal CSV export error: " + (e.message || e));
    }
  });
  qs("prExportJournalXlsxBtn")?.addEventListener("click", () => {
    try {
      exportProcurementJournalsXlsx();
    } catch (e) {
      setStatus("Journal XLSX export error: " + (e.message || e));
    }
  });
  qs("loadRequisitions")?.addEventListener("click", () =>
    loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message))
  );
  qs("prStatusFilter")?.addEventListener("change", () => {
    setProcurementKpiFilter("all");
    loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message));
  });
  qs("prTierFilter")?.addEventListener("change", () => {
    loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message));
  });
  qs("prKpiAll")?.addEventListener("click", () => {
    setProcurementKpiFilter("all");
    loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message));
  });
  qs("prKpiApprovedOpen")?.addEventListener("click", () => {
    const statusEl = qs("prStatusFilter");
    if (statusEl) statusEl.value = "";
    setProcurementKpiFilter("approved_open");
    loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message));
  });
  qs("prKpiInFlow")?.addEventListener("click", () => {
    const statusEl = qs("prStatusFilter");
    if (statusEl) statusEl.value = "";
    setProcurementKpiFilter("in_flow");
    loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message));
  });
  qs("prSaveChainConfig")?.addEventListener("click", () => {
    try {
      saveProcurementChainConfig();
      updateProcurementChainPreview();
    } catch (e) {
      setStatus(`Save chain rules failed: ${e.message || e}`);
    }
  });
  ["prValue", "prTier1Max", "prTier1Chain", "prTier2Max", "prTier2Chain", "prTier3Chain", "prApproverChain"].forEach((id) => {
    qs(id)?.addEventListener("input", updateProcurementChainPreview);
  });
  qs("loadLubeMaps")?.addEventListener("click", () =>
    loadLubeMappings().catch((e) => setStatus("Lube mapping load error: " + e.message))
  );
  qs("saveLubeMap")?.addEventListener("click", () =>
    saveLubeMapping().catch((e) => setStatus("Lube mapping save error: " + e.message))
  );
  qs("saveFuelLog")?.addEventListener("click", () =>
    saveFuelLog().catch((e) => setStatus("Fuel log error: " + e.message))
  );
  qs("fuelMassImportBtn")?.addEventListener("click", () =>
    importFuelMassPaste().catch((e) => setStatus("Fuel mass import error: " + e.message))
  );
  qs("fuelAsset")?.addEventListener("change", () => {
    applyAssetCostCenterToInputs(qs("fuelAsset")?.value);
    syncFuelUnitFromAsset(qs("fuelAsset")?.value, "input").catch(() => {});
  });
  qs("mlAsset")?.addEventListener("change", () => {
    applyAssetCostCenterToInputs(qs("mlAsset")?.value);
  });
  qs("fuelMeterUnit")?.addEventListener("change", () => {
    const mode = String(qs("fuelMeterUnit")?.value || "hours").toLowerCase() === "km" ? "km" : "hours";
    const meterInput = qs("fuelHoursRun");
    if (meterInput) meterInput.placeholder = mode === "km" ? "Distance since fill (km)" : "Hours since fill";
  });
  qs("loadFuelBaseline")?.addEventListener("click", () =>
    loadFuelBaseline().catch((e) => setStatus("Fuel baseline error: " + e.message))
  );
  qs("fuelBaseAsset")?.addEventListener("change", () => {
    syncFuelUnitFromAsset(qs("fuelBaseAsset")?.value, "both").catch(() => {});
  });
  qs("saveFuelBaseline")?.addEventListener("click", () =>
    saveFuelBaseline().catch((e) => setStatus("Fuel baseline error: " + e.message))
  );
  qs("createDispatchTrip")?.addEventListener("click", () =>
    createDispatchTrip().catch((e) => setStatus("Dispatch create error: " + e.message))
  );
  qs("saveDispatchPod")?.addEventListener("click", () =>
    saveDispatchPod().catch((e) => setStatus("Dispatch POD error: " + e.message))
  );
  qs("createDispatchException")?.addEventListener("click", () =>
    createDispatchException().catch((e) => setStatus("Dispatch exception error: " + e.message))
  );
  qs("loadDispatchExceptions")?.addEventListener("click", () =>
    loadDispatchExceptions().catch((e) => setStatus("Dispatch exceptions load error: " + e.message))
  );
  qs("loadDispatchTrips")?.addEventListener("click", () =>
    loadDispatchTrips().catch((e) => setStatus("Dispatch load error: " + e.message))
  );
  qs("dpStatusFilter")?.addEventListener("change", () =>
    loadDispatchTrips().catch((e) => setStatus("Dispatch load error: " + e.message))
  );
  qs("dpVarTolerance")?.addEventListener("change", () =>
    loadDispatchTrips().catch((e) => setStatus("Dispatch load error: " + e.message))
  );
  qs("dpOnlyBreaches")?.addEventListener("change", () =>
    loadDispatchTrips().catch((e) => setStatus("Dispatch load error: " + e.message))
  );
  qs("dpExStatusFilter")?.addEventListener("change", () =>
    loadDispatchExceptions().catch((e) => setStatus("Dispatch exceptions load error: " + e.message))
  );
  qs("dpExOnlyOpen")?.addEventListener("change", () =>
    loadDispatchExceptions().catch((e) => setStatus("Dispatch exceptions load error: " + e.message))
  );
  qs("loadQualityCenter")?.addEventListener("click", () =>
    loadQualityCenter().catch((e) => setStatus("Quality center load error: " + e.message))
  );
  qs("qSeverityFilter")?.addEventListener("change", () =>
    loadQualityCenter().catch((e) => setStatus("Quality center load error: " + e.message))
  );
  qs("qTypeFilter")?.addEventListener("change", () =>
    loadQualityCenter().catch((e) => setStatus("Quality center load error: " + e.message))
  );
  qs("saveOperationEntry")?.addEventListener("click", () =>
    saveOperationEntry().catch((e) => setStatus("Operations save error: " + e.message))
  );
  qs("saveSiteDailyEntry")?.addEventListener("click", () =>
    saveSiteDailyEntry().catch((e) => setStatus("Site daily save error: " + e.message))
  );
  qs("loadSiteDailyEntries")?.addEventListener("click", () =>
    loadSiteDailyEntries().catch((e) => setStatus("Site daily load error: " + e.message))
  );
  qs("saveSiteEquipmentUsage")?.addEventListener("click", () =>
    saveSiteEquipmentUsage().catch((e) => setStatus("Site equipment link error: " + e.message))
  );
  qs("loadSiteEquipmentUsage")?.addEventListener("click", () =>
    loadSiteEquipmentUsage().catch((e) => setStatus("Site equipment load error: " + e.message))
  );
  qs("saveSiteTarget")?.addEventListener("click", () =>
    saveSiteTarget().catch((e) => setStatus("Site target save error: " + e.message))
  );
  qs("loadSiteTargets")?.addEventListener("click", () =>
    loadSiteTargets().catch((e) => setStatus("Site target load error: " + e.message))
  );
  qs("saveSiteDelay")?.addEventListener("click", () =>
    saveSiteDelay().catch((e) => setStatus("Site delay save error: " + e.message))
  );
  qs("loadSiteDelays")?.addEventListener("click", () =>
    loadSiteDelays().catch((e) => setStatus("Site delay load error: " + e.message))
  );
  qs("saveSiteZone")?.addEventListener("click", () =>
    saveSiteZone().catch((e) => setStatus("Site zone save error: " + e.message))
  );
  qs("loadSiteZones")?.addEventListener("click", () =>
    loadSiteZones().catch((e) => setStatus("Site zone load error: " + e.message))
  );
  qs("loadSiteDashboard")?.addEventListener("click", () =>
    loadSiteDashboard().catch((e) => setStatus("Site dashboard load error: " + e.message))
  );
  qs("saveOperationsClosingDraft")?.addEventListener("click", () =>
    saveOperationsClosing(false).catch((e) => setStatus("Operations closing error: " + e.message))
  );
  qs("closeOperationsDay")?.addEventListener("click", () =>
    saveOperationsClosing(true).catch((e) => setStatus("Operations close day error: " + e.message))
  );
  qs("reopenOperationsDay")?.addEventListener("click", () =>
    reopenOperationsDay().catch((e) => setStatus("Operations reopen day error: " + e.message))
  );
  qs("opDate")?.addEventListener("change", () =>
    loadOperationsClosingForDate((qs("opDate")?.value || "").trim()).catch((e) => setStatus("Operations closing load error: " + e.message))
  );
  qs("loadOperations")?.addEventListener("click", () =>
    loadOperations().catch((e) => setStatus("Operations load error: " + e.message))
  );
  qs("opClientMetric")?.addEventListener("change", () =>
    loadOperations().catch((e) => setStatus("Operations load error: " + e.message))
  );
  qs("opClientTopN")?.addEventListener("change", () =>
    loadOperations().catch((e) => setStatus("Operations load error: " + e.message))
  );
  qs("loadCostSettings")?.addEventListener("click", () =>
    loadCostSettings().catch((e) => setStatus("Cost settings error: " + e.message))
  );
  qs("saveCostSettings")?.addEventListener("click", () =>
    saveCostSettings().catch((e) => setStatus("Cost settings error: " + e.message))
  );
  qs("saveCostAssetRates")?.addEventListener("click", () =>
    saveCostAssetRates().catch((e) => setStatus("Cost asset rates error: " + e.message))
  );
  qs("saveCostPartRate")?.addEventListener("click", () =>
    saveCostPartRate().catch((e) => setStatus("Cost part rate error: " + e.message))
  );
  qs("loadFuelBenchmark")?.addEventListener("click", () =>
    loadFuelBenchmark().catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelPresetQ1")?.addEventListener("click", () =>
    applyFuelPeriodPreset("q1").catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelPresetQ2")?.addEventListener("click", () =>
    applyFuelPeriodPreset("q2").catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelPresetQ3")?.addEventListener("click", () =>
    applyFuelPeriodPreset("q3").catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelPresetYtd")?.addEventListener("click", () =>
    applyFuelPeriodPreset("ytd").catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelPresetMtd")?.addEventListener("click", () =>
    applyFuelPeriodPreset("mtd").catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelDupOnly")?.addEventListener("change", () =>
    loadFuelBenchmark().catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelEquipSelectAll")?.addEventListener("click", () => {
    const host = qs("fuelEquipFilterList");
    host?.querySelectorAll('input[type="checkbox"][data-fuel-equip]')?.forEach((b) => { b.checked = true; });
    renderFuelEquipmentChart(window.__fuelBenchmarkChartRows || []);
  });
  qs("fuelEquipClear")?.addEventListener("click", () => {
    const host = qs("fuelEquipFilterList");
    host?.querySelectorAll('input[type="checkbox"][data-fuel-equip]')?.forEach((b) => { b.checked = false; });
    renderFuelEquipmentChart(window.__fuelBenchmarkChartRows || []);
  });
  qs("fuelEquipTypeFilter")?.addEventListener("change", (evt) => {
    const type = String(evt.target?.value || "").trim();
    if (!type) return;
    const host = qs("fuelEquipFilterList");
    if (!host) return;
    // Uncheck all, then check only the matching type
    host.querySelectorAll('input[type="checkbox"][data-fuel-equip]').forEach((b) => { b.checked = false; });
    host.querySelectorAll(`label[data-fuel-type="${CSS.escape(type)}"] input[type="checkbox"]`).forEach((b) => { b.checked = true; });
    // Reset the select so it can be used again next time
    evt.target.value = "";
    renderFuelEquipmentChart(window.__fuelBenchmarkChartRows || []);
  });
  qs("fuelEquipViewMode")?.addEventListener("change", () => {
    renderFuelEquipmentChart(window.__fuelBenchmarkChartRows || []);
  });
  qs("fuelEquipFilterList")?.addEventListener("change", (evt) => {
    const t = evt.target;
    if (!(t instanceof HTMLInputElement) || t.type !== "checkbox") return;
    if (!t.hasAttribute("data-fuel-equip")) return;
    renderFuelEquipmentChart(window.__fuelBenchmarkChartRows || []);
  });
  qs("loadFuelSnapshots")?.addEventListener("click", () =>
    loadFuelSnapshots().catch((e) => setStatus("Fuel snapshots error: " + e.message))
  );
  qs("fuelBenchmarkList")?.addEventListener("click", (evt) => {
    const pdfBtn = evt.target?.closest?.("button[data-fuel-machine-pdf]");
    if (pdfBtn) {
      const code = String(pdfBtn.getAttribute("data-fuel-machine-pdf") || "").trim();
      if (!code) return;
      openFuelMachineHistoryPdf(code, false);
      return;
    }

    const saveBtn = evt.target?.closest?.("button[data-fuel-save]");
    if (saveBtn) {
      const rowEl = saveBtn.closest(".item");
      const mountEl = rowEl?.querySelector?.(".fuel-inline-history");
      const code = String(mountEl?.getAttribute?.("data-code") || "");
      saveFuelMachineHoursInline(saveBtn)
        .then(() => Promise.all([
          loadFuelBenchmark().catch(() => {}),
          code && mountEl ? loadFuelMachineDailyInline(code, mountEl).catch(() => {}) : Promise.resolve(),
        ]))
        .then(() => setStatus("Machine hours updated."))
        .catch((e) => setStatus("Machine hours update failed: " + (e.message || e)));
      return;
    }

    const delBtn = evt.target?.closest?.("button[data-fuel-delete]");
    if (delBtn) {
      const logId = Number(delBtn.getAttribute("data-fuel-delete") || 0);
      const rowEl = delBtn.closest(".item");
      const mountEl = rowEl?.querySelector?.(".fuel-inline-history");
      const code = String(mountEl?.getAttribute?.("data-code") || "");
      deleteFuelLogEntry(logId)
        .then(() => Promise.all([
          loadFuelBenchmark().catch(() => {}),
          code && mountEl ? loadFuelMachineDailyInline(code, mountEl).catch(() => {}) : Promise.resolve(),
        ]))
        .then(() => setStatus("Fuel input deleted."))
        .catch((e) => setStatus("Delete fuel input failed: " + (e.message || e)));
      return;
    }

    const btn = evt.target?.closest?.("button[data-fuel-machine]");
    if (!btn) return;
    const code = String(btn.getAttribute("data-fuel-machine") || "").trim();
    if (!code) return;
    const rowEl = btn.closest(".item");
    const mountEl = rowEl?.querySelector?.(".fuel-inline-history");
    if (!mountEl) return;
    const opened = mountEl.getAttribute("data-opened") === "1";
    const openedCode = String(mountEl.getAttribute("data-code") || "");
    if (opened && openedCode === code) {
      mountEl.innerHTML = "";
      mountEl.setAttribute("data-opened", "0");
      mountEl.setAttribute("data-code", "");
      return;
    }
    mountEl.setAttribute("data-opened", "1");
    mountEl.setAttribute("data-code", code);
    loadFuelMachineDailyInline(code, mountEl).catch((e) => {
      mountEl.innerHTML = `<small>Machine history error: ${String(e.message || e)}</small>`;
      setStatus("Machine fuel consumption error: " + (e.message || e));
    });
  });
  qs("openFuelBenchmarkPdf")?.addEventListener("click", () => openFuelBenchmarkPdf(false));
  qs("downloadFuelBenchmarkPdf")?.addEventListener("click", () => openFuelBenchmarkPdf(true));
  qs("downloadFuelBenchmarkXlsx")?.addEventListener("click", () => openFuelBenchmarkXlsx());
  qs("openFuelReconPdf")?.addEventListener("click", () => openFuelReconciliationPdf(false));
  qs("downloadFuelReconPdf")?.addEventListener("click", () => openFuelReconciliationPdf(true));
  qs("downloadFuelReconXlsx")?.addEventListener("click", () => openFuelReconciliationXlsx());
  qs("runFuelReconYtd")?.addEventListener("click", () =>
    runFuelReconciliation(true).catch((e) => setStatus("Fuel reconciliation error: " + e.message))
  );
  qs("runFuelReconRange")?.addEventListener("click", () =>
    runFuelReconciliation(false).catch((e) => setStatus("Fuel reconciliation error: " + e.message))
  );
  qs("downloadExecutivePackXlsx")?.addEventListener("click", () => {
    downloadExecutivePackExcel().catch((e) => setStatus("Executive pack error: " + e.message));
  });
  qs("loadStockMonitor")?.addEventListener("click", () =>
    loadStockMonitor().catch((e) => setStatus("Stock monitor error: " + e.message))
  );
  qs("spLoad")?.addEventListener("click", () =>
    loadStockOnHandPage().catch((e) => setStatus("Stock on hand error: " + e.message))
  );
  qs("spSort")?.addEventListener("change", () => refreshStockInventoryDisplay());
  qs("spOnlyLow")?.addEventListener("change", () => refreshStockInventoryDisplay());
  qs("spFilter")?.addEventListener("input", () => {
    if (stockPageData.rows.length) refreshStockInventoryDisplay();
  });
  qs("spExportCsv")?.addEventListener("click", exportStockOnHandCsv);
  qs("spOpenPdf")?.addEventListener("click", openStockOnHandPdf);
  qs("smrLoad")?.addEventListener("click", () =>
    loadStockMovementsReport().catch((e) => setStatus("Stock movements report error: " + e.message))
  );
  qs("smrExportCsv")?.addEventListener("click", exportStockMovementsReportCsv);
  qs("smrOpenPdf")?.addEventListener("click", openStockMovementsReportPdf);
  qs("spoLoad")?.addEventListener("click", () =>
    loadStoresPartOrders().catch((e) => setStatus("Parts purchases error: " + e.message))
  );
  qs("spoSave")?.addEventListener("click", () =>
    saveStoresPartOrder().catch((e) => {
      const msg = qs("spoFormMsg");
      if (msg) msg.textContent = e.message || String(e);
    })
  );
  qs("spoClear")?.addEventListener("click", clearStoresPartOrderForm);
  qs("spoFilterStatus")?.addEventListener("change", () => {
    loadStoresPartOrders().catch(() => {});
  });
  qs("spoList")?.addEventListener("change", (evt) => {
    const sel = evt.target?.closest?.("select[data-spo-status]");
    if (!sel) return;
    const id = Number(sel.getAttribute("data-spo-status") || 0);
    const status = String(sel.value || "").trim();
    if (!id) return;
    updateStoresPartOrderStatus(id, status).catch((e) => alert(e.message || String(e)));
  });
  qs("spoList")?.addEventListener("click", (evt) => {
    const saveBtn = evt.target?.closest?.("button[data-spo-save]");
    if (saveBtn) {
      saveStoresPartOrderRow(Number(saveBtn.getAttribute("data-spo-save") || 0)).catch((e) =>
        alert(e.message || String(e))
      );
      return;
    }
    const recvBtn = evt.target?.closest?.("button[data-spo-receive]");
    if (recvBtn) {
      receiveStoresPartOrderToInventory(Number(recvBtn.getAttribute("data-spo-receive") || 0)).catch((e) =>
        alert(e.message || String(e))
      );
      return;
    }
    const btn = evt.target?.closest?.("button[data-spo-del]");
    if (!btn) return;
    cancelStoresPartOrder(Number(btn.getAttribute("data-spo-del") || 0)).catch((e) => alert(e.message || String(e)));
  });
  qs("spoExportXlsx")?.addEventListener("click", () =>
    exportStoresPartOrdersXlsx().catch((e) => setStatus("Parts purchases Excel error: " + e.message))
  );
  qs("spoOpenPdf")?.addEventListener("click", () => openStoresPartOrdersPdf(false));
  qs("spoDownloadPdf")?.addEventListener("click", () => openStoresPartOrdersPdf(true));

  qs("ptPartsLoad")?.addEventListener("click", () =>
    loadPtPartsOrders().catch((e) => setStatus("Parts tracking error: " + e.message))
  );
  qs("ptPartsSave")?.addEventListener("click", () =>
    savePtPartsOrder().catch((e) => {
      const msg = qs("ptPartsMsg");
      if (msg) msg.textContent = e.message || String(e);
    })
  );
  qs("ptPartsClear")?.addEventListener("click", clearPtPartsForm);
  qs("ptPartsStatus")?.addEventListener("change", () => loadPtPartsOrders().catch(() => {}));
  qs("ptPartsSearch")?.addEventListener("input", () => renderPtPartsTable(ptPartsCache));
  qs("ptPartsList")?.addEventListener("change", (evt) => {
    const sel = evt.target?.closest?.("select[data-pt-status]");
    if (!sel) return;
    const id = Number(sel.getAttribute("data-pt-status") || 0);
    if (!id) return;
    const rowEl = qs("ptPartsList")?.querySelector(`tr[data-pt-row="${id}"]`);
    const patch = readPtPartsRowPatch(rowEl) || {};
    patch.status = String(sel.value || "").trim();
    patchStoresPartOrder(id, patch)
      .then(() => loadPtPartsOrders())
      .catch((e) => alert(e.message || String(e)));
  });
  qs("ptPartsList")?.addEventListener("click", (evt) => {
    const saveBtn = evt.target?.closest?.("button[data-pt-save]");
    if (saveBtn) {
      const id = Number(saveBtn.getAttribute("data-pt-save") || 0);
      const rowEl = qs("ptPartsList")?.querySelector(`tr[data-pt-row="${id}"]`);
      const patch = readPtPartsRowPatch(rowEl);
      if (!patch) return;
      patchStoresPartOrder(id, patch)
        .then(() => {
          setStatus("Part line saved.");
          return loadPtPartsOrders();
        })
        .catch((e) => alert(e.message || String(e)));
      return;
    }
    const recvBtn = evt.target?.closest?.("button[data-pt-receive]");
    if (recvBtn) {
      receiveStoresPartOrderToInventory(Number(recvBtn.getAttribute("data-pt-receive") || 0))
        .then(() => loadPtPartsOrders())
        .catch((e) => alert(e.message || String(e)));
      return;
    }
    const delBtn = evt.target?.closest?.("button[data-pt-del]");
    if (delBtn) {
      cancelStoresPartOrder(Number(delBtn.getAttribute("data-pt-del") || 0))
        .then(() => loadPtPartsOrders())
        .catch((e) => alert(e.message || String(e)));
    }
  });
  qs("ptOffLoad")?.addEventListener("click", () =>
    loadPtOffsiteRepairs().catch((e) => setStatus("Off-site tracking error: " + e.message))
  );
  qs("ptOffSave")?.addEventListener("click", () =>
    savePtOffsiteRepair().catch((e) => {
      const msg = qs("ptOffMsg");
      if (msg) msg.textContent = e.message || String(e);
    })
  );
  qs("ptOffClear")?.addEventListener("click", clearPtOffForm);
  qs("ptOffStatusFilter")?.addEventListener("change", () => renderPtOffsiteTable(ptOffsiteCache));
  qs("ptOffIncludeReturned")?.addEventListener("change", () => loadPtOffsiteRepairs().catch(() => {}));
  qs("ptOffList")?.addEventListener("click", (evt) => {
    const btn = evt.target?.closest?.("button[data-pt-off-save]");
    if (!btn) return;
    savePtOffsiteRow(Number(btn.getAttribute("data-pt-off-save") || 0)).catch((e) =>
      alert(e.message || String(e))
    );
  });
  qs("loadAudit")?.addEventListener("click", () =>
    loadAuditLogs().catch((e) => setStatus("Audit error: " + e.message))
  );
  qs("loginSubmit")?.addEventListener("click", () => submitLoginForm());
  qs("loginPassword")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") submitLoginForm();
  });
  qs("loadAdminUsersBtn")?.addEventListener("click", () =>
    loadAdminUsers().catch((e) => setStatus("Admin users error: " + e.message))
  );
  qs("mdmLoadBtn")?.addEventListener("click", () =>
    loadMasterDataGovernance().catch((e) => setStatus("Master data error: " + e.message))
  );
  qs("mdmDeptSaveBtn")?.addEventListener("click", () =>
    saveMdmDepartment().catch((e) => setStatus("Department error: " + e.message))
  );
  qs("mdmCcSaveBtn")?.addEventListener("click", () =>
    saveMdmCostCenter().catch((e) => setStatus("Cost center error: " + e.message))
  );
  qs("mdmSupSaveBtn")?.addEventListener("click", () =>
    saveMdmSupplier().catch((e) => setStatus("Supplier error: " + e.message))
  );
  qs("mdmPolicySaveBtn")?.addEventListener("click", () =>
    saveMdmPolicies().catch((e) => setStatus("Policy error: " + e.message))
  );
  qs("adminArtisanPresetBtn")?.addEventListener("click", applyAdminArtisanPreset);
  qs("saveAdminUserBtn")?.addEventListener("click", () =>
    saveAdminUser().catch((e) => setStatus("Save user error: " + e.message))
  );
  qs("chPwdSubmit")?.addEventListener("click", () =>
    submitChangePassword().catch((e) => setStatus("Password error: " + e.message))
  );
  qs("loadSmtpSettingsBtn")?.addEventListener("click", () =>
    loadSmtpSettings().catch((e) => setStatus("SMTP load error: " + e.message))
  );
  qs("saveSmtpSettingsBtn")?.addEventListener("click", () =>
    saveSmtpSettings().catch((e) => setStatus("SMTP save error: " + e.message))
  );
  qs("testSmtpSettingsBtn")?.addEventListener("click", () =>
    testSmtpSettings().catch((e) => setStatus("SMTP test error: " + e.message))
  );
  qs("loadPushNotifyBtn")?.addEventListener("click", () =>
    loadPushNotificationSettings().catch((e) => setStatus("Push load error: " + e.message))
  );
  qs("sendPushNotifyTestBtn")?.addEventListener("click", () =>
    sendPushNotificationTest().catch((e) => setStatus("Push test error: " + e.message))
  );
  qs("sendPushNotifyBtn")?.addEventListener("click", () =>
    sendPushNotificationManual().catch((e) => setStatus("Push send error: " + e.message))
  );
  qs("loadPdfReportSettingsBtn")?.addEventListener("click", () =>
    loadPdfReportSettings().catch((e) => setStatus("PDF site load error: " + e.message))
  );
  qs("savePdfReportSettingsBtn")?.addEventListener("click", () =>
    savePdfReportSettings().catch((e) => setStatus("PDF site save error: " + e.message))
  );
  qs("uploadPdfReportLogoBtn")?.addEventListener("click", () =>
    uploadPdfReportLogo().catch((e) => setStatus("PDF logo upload error: " + e.message))
  );
  qs("removePdfReportLogoBtn")?.addEventListener("click", () =>
    removePdfReportLogo().catch((e) => setStatus("PDF logo remove error: " + e.message))
  );
  qs("pdfReportSiteCode")?.addEventListener("change", () => onPdfReportSiteCodeChange());
  qs("pdfReportCompanyCode")?.addEventListener("change", () => onPdfReportCompanyCodeChange());
  qs("sendSubscriptionNowBtn")?.addEventListener("click", () =>
    sendSubscriptionNowFromAdmin().catch((e) => setStatus("Subscription send error: " + e.message))
  );
  qs("refreshBackupsBtn")?.addEventListener("click", () =>
    loadBackupFiles().catch((e) => setStatus("Backups load error: " + e.message))
  );
  qs("createBackupNowBtn")?.addEventListener("click", () =>
    createBackupNow().catch((e) => setStatus("Backup create error: " + e.message))
  );
  qs("previewBackupRestoreBtn")?.addEventListener("click", () =>
    previewBackupRestore().catch((e) => setStatus("Backup preview error: " + e.message))
  );
  qs("stageBackupRestoreBtn")?.addEventListener("click", () =>
    stageBackupRestore().catch((e) => setStatus("Restore stage error: " + e.message))
  );
  qs("executeBackupRestoreBtn")?.addEventListener("click", () =>
    executeBackupRestoreNow().catch((e) => setStatus("Restore execute error: " + e.message))
  );
  qs("loadApprovals")?.addEventListener("click", () =>
    loadApprovalRequests().catch((e) => setStatus("Approvals error: " + e.message))
  );
  qs("approvalStatus")?.addEventListener("change", () =>
    loadApprovalRequests().catch((e) => setStatus("Approvals error: " + e.message))
  );
  qs("legalUploadBtn")?.addEventListener("click", () =>
    uploadLegalDoc().catch((e) => setStatus("Legal upload error: " + e.message))
  );
  qs("loadLegalBtn")?.addEventListener("click", () =>
    loadLegalDocs().catch((e) => setStatus("Legal load error: " + e.message))
  );
  qs("loadLegalExpiryBtn")?.addEventListener("click", () =>
    loadLegalExpiry().catch((e) => setStatus("Legal expiry error: " + e.message))
  );
  qs("openLegalCompliancePdf")?.addEventListener("click", () => openLegalCompliancePdf(false));
  qs("downloadLegalCompliancePdf")?.addEventListener("click", () => openLegalCompliancePdf(true));
  qs("doUpload")?.addEventListener("click", () =>
    doUpload().catch((e) => setStatus("Upload error: " + e.message))
  );
  qs("fuelFamsUploadBtn")?.addEventListener("click", () =>
    importFamsFuelFile().catch((e) => setStatus("FAMS import error: " + e.message))
  );
  qs("fuelFamsSyncNowBtn")?.addEventListener("click", () =>
    syncFamsFuelNow().catch((e) => setStatus("FAMS sync error: " + e.message))
  );
  qs("fuelFamsRefreshStatusBtn")?.addEventListener("click", () =>
    loadFamsFuelStatus().catch((e) => setStatus("FAMS status error: " + e.message))
  );
  qs("fuelRepairMeterChainBtn")?.addEventListener("click", () =>
    repairFuelMeterChain().catch((e) => setStatus("Meter chain repair error: " + e.message))
  );
  qs("fuelClearFromDateBtn")?.addEventListener("click", () =>
    clearFuelFromDate().catch((e) => setStatus("Fuel clear error: " + e.message))
  );
  qs("fuelClearPreviewBtn")?.addEventListener("click", () =>
    runFuelClearPreview().catch((e) => setStatus("Fuel clear preview error: " + e.message))
  );
  qs("downloadFuelTemplate")?.addEventListener("click", downloadFuelCsvTemplate);
  qs("downloadStoreTemplate")?.addEventListener("click", downloadStoresCsvTemplate);
  qs("downloadFuelBaselineTemplate")?.addEventListener("click", downloadFuelBaselineCsvTemplate);

  qs("openDaily")?.addEventListener("click", openDailyPdf);
  qs("openWeekly")?.addEventListener("click", openWeeklyPdf);
  qs("openLubePdf")?.addEventListener("click", openLubePdf);
  qs("openLubePdfFromLube")?.addEventListener("click", openLubePdf);
  qs("downloadLubeUsageXlsx")?.addEventListener("click", downloadLubeUsageXlsx);
  qs("loadLubeMonthStock")?.addEventListener("click", () =>
    loadLubeMonthStock().catch((e) => setStatus("Lube month stock error: " + e.message))
  );
  qs("openStockMonitorPdf")?.addEventListener("click", openStockMonitorPdf);
  qs("downloadStockMonitorPdf")?.addEventListener("click", downloadStockMonitorPdf);
  qs("openOperationsPdf")?.addEventListener("click", () => openOperationsPdf(false));
  qs("downloadOperationsPdf")?.addEventListener("click", () => openOperationsPdf(true));
  qs("downloadOperationsXlsx")?.addEventListener("click", downloadOperationsXlsx);
  qs("openDailyXlsx")?.addEventListener("click", openDailyXlsx);
  qs("openGmWeeklyXlsx")?.addEventListener("click", openGmWeeklyXlsx);
  qs("downloadCostMonthlyXlsx")?.addEventListener("click", downloadCostMonthlyXlsx);
  qs("openMonthlyFleetCostPdf")?.addEventListener("click", openMonthlyFleetCostPdf);
  qs("downloadMonthlyFleetCostPdf")?.addEventListener("click", downloadMonthlyFleetCostPdf);
  qs("downloadMtdOpeningHoursXlsx")?.addEventListener("click", downloadMtdOpeningHoursXlsx);
  qs("downloadMaintenanceCostByEquipmentXlsx")?.addEventListener("click", downloadMaintenanceCostByEquipmentXlsx);
  qs("openMaintenanceCostByEquipmentPdf")?.addEventListener("click", () => openMaintenanceCostByEquipmentPdf(false));
  qs("downloadMaintenanceCostByEquipmentPdf")?.addEventListener("click", () => openMaintenanceCostByEquipmentPdf(true));
  qs("downloadMaintenanceExecutivePptx")?.addEventListener("click", downloadMaintenanceExecutivePptx);
  qs("downloadGMUpcomingCostsPptx")?.addEventListener("click", downloadGMUpcomingCostsPptx);
  qs("downloadGMBudgetMeetingDocx")?.addEventListener("click", downloadGMBudgetMeetingDocx);
  qs("saveRainDayBtn")?.addEventListener("click", () => saveRainDay().catch((e) => setStatus("Rain day save error: " + e.message)));
  qs("removeRainDayBtn")?.addEventListener("click", () => removeRainDay().catch((e) => setStatus("Rain day remove error: " + e.message)));
  qs("loadRainDaysBtn")?.addEventListener("click", () => loadRainDays().catch((e) => setStatus("Rain day load error: " + e.message)));

  qs("boRefreshOpen")?.addEventListener("click", () =>
    loadBreakdownOpsOpen().catch((e) => setStatus("Open list error: " + e.message))
  );
  qs("boRefreshRecent")?.addEventListener("click", () =>
    loadBreakdownOpsRecent().catch((e) => setStatus("Recent list error: " + e.message))
  );
  qs("boEnsureOpen")?.addEventListener("click", () =>
    ensureOpenBreakdownOps().catch((e) => setStatus("Ensure open error: " + e.message))
  );
  qs("boRepairCreateWo")?.addEventListener("click", () =>
    createRepairWorkOrderOps().catch((e) => setStatus("Repair WO error: " + e.message))
  );
  qs("boRepairOpenWo")?.addEventListener("click", () => {
    if (lastBoRepairWoId) window.open(`/web/workorders.html?wo=${encodeURIComponent(String(lastBoRepairWoId))}`, "_blank");
  });
  qs("boPullLiveHours")?.addEventListener("click", () =>
    pullBreakdownOpsLiveHours().catch((e) => setStatus("Live hours error: " + e.message))
  );
  qs("boOpenList")?.addEventListener("click", (ev) => {
    const w = ev.target?.closest?.(".bo-copy-wo");
    if (w) {
      const wo = w.getAttribute("data-wo");
      if (wo && qs("iWo")) qs("iWo").value = String(wo);
      setStatus(`Copied WO #${wo} for parts issue.`);
      return;
    }
    const c = ev.target?.closest?.(".bo-close-bdn");
    if (c) closeBreakdownFromOps(c.getAttribute("data-id")).catch(() => {});
  });

  qs("boSlipType")?.addEventListener("change", updateBoSlipFormVisibility);
  qs("boSlipPhotosInput")?.addEventListener("change", (e) =>
    onBoSlipPhotosInputChange(e).catch((err) => setStatus(String(err.message || err)))
  );
  qs("boSlipPhotosClear")?.addEventListener("click", () => {
    clearBoSlipPhotosUi();
    setStatus("Slip pictures cleared.");
  });
  qs("boSlipPullAsset")?.addEventListener("click", () =>
    pullBoSlipFromAsset().catch((e) => setStatus("Pull from asset error: " + (e.message || e)))
  );
  qs("boSlipSave")?.addEventListener("click", () =>
    saveBoSlipReport().catch((e) => setStatus("Slip save error: " + e.message))
  );
  qs("boSlipLoadList")?.addEventListener("click", () =>
    loadBoSlipSavedList().catch((e) => setStatus("Slip list error: " + e.message))
  );
  qs("boSlipSavedList")?.addEventListener("click", (ev) => {
    const b = ev.target?.closest?.(".bo-slip-pdf");
    if (b) openBoSlipPdf(b.getAttribute("data-id"));
  });

  qs("makeBreakdown")?.addEventListener("click", () =>
    createBreakdown().catch((e) => setStatus("Breakdown error: " + e.message))
  );
  bindShortBreakdownPartsUi();
  qs("sqSubmit")?.addEventListener("click", () =>
    submitShortBreakdown().catch((e) => setStatus("Short breakdown error: " + e.message))
  );
  qs("issuePart")?.addEventListener("click", () =>
    issuePart().catch((e) => setStatus("Issue error: " + e.message))
  );
  qs("allocateStore")?.addEventListener("click", () =>
    allocateStore().catch((e) => setStatus("Stores allocation error: " + e.message))
  );
  qs("refreshAllocations")?.addEventListener("click", () =>
    loadStoreAllocations().catch((e) => setStatus("Allocation list error: " + e.message))
  );
  qs("saveManualStock")?.addEventListener("click", () =>
    saveManualStock().catch((e) => setStatus("Manual stock error: " + e.message))
  );
  qs("msPart")?.addEventListener("input", updateManualStockPartDesc);
  qs("msPart")?.addEventListener("change", updateManualStockPartDesc);
  qs("msType")?.addEventListener("change", () => {
    updateManualStockCostRowVisibility();
  });
  qs("msType")?.addEventListener("input", () => {
    updateManualStockCostRowVisibility();
  });
  qs("mlPart")?.addEventListener("input", updateManualLubePartDesc);
  qs("mlPart")?.addEventListener("change", updateManualLubePartDesc);
  // Lube minimums moved to separate card
  qs("lubeMinPart")?.addEventListener("input", updateLubeMinPartDesc);
  qs("lubeMinPart")?.addEventListener("change", updateLubeMinPartDesc);
  qs("lubeMinSetOne")?.addEventListener("click", () =>
    setSingleLubeMinimum().catch((e) => setStatus("Lube min error: " + e.message))
  );
  qs("lubeMinRefresh")?.addEventListener("click", () =>
    loadLubeReorderAlerts().catch((e) => setStatus("Lube alerts error: " + e.message))
  );
  qs("receiveLube")?.addEventListener("click", () =>
    receiveLubeStock().catch((e) => setStatus("Receive lube error: " + e.message))
  );
  qs("lrPart")?.addEventListener("input", updateReceiveLubePartDesc);
  qs("lrPart")?.addEventListener("change", updateReceiveLubePartDesc);
  qs("icLoad")?.addEventListener("click", () =>
    loadInventoryControl().catch((e) => setStatus("Inventory control error: " + e.message))
  );
  qs("icSaveMin")?.addEventListener("click", () =>
    saveInventoryPartMinimum().catch((e) => setStatus("Part minimum error: " + e.message))
  );
  qs("icSubmitCount")?.addEventListener("click", () =>
    submitInventoryCycleCount().catch((e) => setStatus("Cycle count error: " + e.message))
  );
  qs("icPartCode")?.addEventListener("change", () =>
    loadInventoryControl().catch((e) => setStatus("Inventory control error: " + e.message))
  );
  qs("saveManualLube")?.addEventListener("click", () =>
    saveManualLube().catch((e) => setStatus("Manual lube error: " + e.message))
  );
  ["msLocation", "saLocation", "mlLocation"].forEach((id) => {
    qs(id)?.addEventListener("change", () => {
      const v = String(qs(id)?.value || "").trim().toUpperCase();
      if (!v) return;
      setRoleDefaultLocation(getSessionRole(), v);
      applyDefaultLocationsToInputs();
      if (id === "saLocation") {
        const binInput = qs("saBin");
        if (binInput) binInput.value = "";
        loadBinCodeOptionsForLocation(v, "saBinCodeOptions").catch(() => {});
      }
      if (id === "msLocation") {
        const binInput = qs("msBin");
        if (binInput) binInput.value = "";
        loadBinCodeOptionsForLocation(v, "msBinCodeOptions").catch(() => {});
      }
    });
  });
  qs("locLoad")?.addEventListener("click", () =>
    loadLocations().catch((e) => setStatus("Locations error: " + e.message))
  );
  qs("locShowInactive")?.addEventListener("change", () =>
    loadLocations().catch((e) => setStatus("Locations error: " + e.message))
  );
  qs("locSave")?.addEventListener("click", () =>
    saveLocation().catch((e) => setStatus("Location save error: " + e.message))
  );
  qs("sbSaveBinBtn")?.addEventListener("click", () =>
    saveStockBin().catch((e) => setStatus("Bin save error: " + e.message))
  );
  qs("sbLoadBinsBtn")?.addEventListener("click", () =>
    loadStockBins().catch((e) => setStatus("Bins load error: " + e.message))
  );
  qs("sbSaveMinMaxBtn")?.addEventListener("click", () =>
    saveStockMinMax().catch((e) => setStatus("Min-max save error: " + e.message))
  );
  qs("sbLoadDepthBtn")?.addEventListener("click", () =>
    loadStockDepth().catch((e) => setStatus("Depth load error: " + e.message))
  );
  qs("sbLoadReplenishmentBtn")?.addEventListener("click", () =>
    loadReplenishmentSuggestions().catch((e) => setStatus("Replenishment load error: " + e.message))
  );
  qs("sbExportReplenishmentCsvBtn")?.addEventListener("click", () => {
    try {
      exportReplenishmentSuggestionsCsv();
      setStatus("Replenishment CSV exported.");
    } catch (e) {
      setStatus("Replenishment export error: " + (e.message || e));
    }
  });
  qs("sbCreateCycleSessionBtn")?.addEventListener("click", () =>
    createCycleSession().catch((e) => setStatus("Cycle session create error: " + e.message))
  );
  qs("sbLoadCycleSessionsBtn")?.addEventListener("click", () =>
    loadCycleSessions().catch((e) => setStatus("Cycle sessions load error: " + e.message))
  );
  qs("loadLubeStock")?.addEventListener("click", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlPart")?.addEventListener("change", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlPart")?.addEventListener("input", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlType")?.addEventListener("change", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlType")?.addEventListener("input", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlQty")?.addEventListener("input", updateLubeQtyWarning);
  qs("setLubeMin210")?.addEventListener("click", () =>
    setLubeMinimumStock().catch((e) => setStatus("Lube minimum error: " + e.message))
  );

  // Daily
  qs("loadDaily")?.addEventListener("click", () =>
    loadDailyInput().catch((e) => setStatus("Daily load error: " + e.message))
  );
  qs("saveDaily")?.addEventListener("click", () =>
    saveDailyInput().catch((e) => setStatus("Daily save error: " + e.message))
  );
  qs("runShiftSelfCheck")?.addEventListener("click", () =>
    runShiftSelfCheck().catch((e) => setStatus("Self-check error: " + e.message))
  );
  qs("exportShiftSelfCheck")?.addEventListener("click", exportShiftSelfCheckTxt);

  qs("copyYesterday")?.addEventListener("click", () =>
    copyYesterdayToToday().catch((e) => setStatus("Copy yesterday error: " + e.message))
  );
  qs("dailyHoursCsvBtn")?.addEventListener("click", () => qs("dailyHoursCsvFile")?.click());
  qs("dailyHoursCsvTemplate")?.addEventListener("click", downloadDailyHoursCsvTemplate);
  qs("dailyHoursCsvFile")?.addEventListener("change", (e) => {
    const file = e.target?.files?.[0];
    if (file) uploadDailyHoursCsv(file).finally(() => { e.target.value = ""; });
  });
  qs("dailyMatrixCsvBtn")?.addEventListener("click", () => qs("dailyMatrixCsvFile")?.click());
  qs("dailyMatrixCsvTemplate")?.addEventListener("click", downloadDailyMatrixCsvTemplate);
  qs("dailyMatrixCsvFile")?.addEventListener("change", (e) => {
    const file = e.target?.files?.[0];
    if (file) uploadDailyMatrixCsv(file).finally(() => { e.target.value = ""; });
  });
  qs("applyBulkSched")?.addEventListener("click", applyBulkScheduled);
  qs("dailyDownOnly")?.addEventListener("change", () => {
    dailyShowDownOnly = !!qs("dailyDownOnly")?.checked;
    renderDailyTable();
  });
  qs("dailyQrGenerate")?.addEventListener("click", () =>
    generateDailyAssetQr().catch((e) => setStatus("QR generate error: " + e.message))
  );
  qs("dailyQrPrint")?.addEventListener("click", printDailyAssetQr);
  qs("dailyQrDownloadVisible")?.addEventListener("click", () =>
    downloadAllVisibleDailyQrs().catch((e) => setStatus("Bulk QR download error: " + e.message))
  );
  qs("dailyQrPrintVisible")?.addEventListener("click", () =>
    printVisibleDailyQrSheet().catch((e) => setStatus("QR sheet print error: " + e.message))
  );
  qs("qrPreset")?.addEventListener("change", applyQrSheetPreset);
  ["qrCols", "qrSizeMm", "qrCellMm", "qrGapMm"].forEach((id) => {
    qs(id)?.addEventListener("input", () => {
      const preset = qs("qrPreset");
      if (preset && preset.value !== "custom") preset.value = "custom";
    });
  });
  applyQrSheetPreset();

  qs("safetyTplLoadBtn")?.addEventListener("click", () =>
    loadSafetyTemplateEditor().catch((e) => setStatus("Safety template error: " + e.message))
  );
  qs("safetyTplSelect")?.addEventListener("change", () =>
    loadSafetyTemplateEditor().catch(() => {})
  );
  qs("safetyCategoryAddBtn")?.addEventListener("click", () =>
    addSafetyCategory().catch((e) => setStatus("Add category error: " + e.message))
  );
  qs("safetyCategoriesList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-safety-edit-category]");
    if (!btn) return;
    const key = String(btn.getAttribute("data-safety-edit-category") || "").trim();
    if (qs("safetyTplSelect")) qs("safetyTplSelect").value = key;
    loadSafetyTemplateEditor().catch(() => {});
  });
  qs("safetyTplSaveBtn")?.addEventListener("click", () =>
    saveSafetyTemplateEditor().catch((e) => setStatus("Safety template save error: " + e.message))
  );
  qs("safetyTplAddRowBtn")?.addEventListener("click", () => {
    safetyTplItems.push({ key: `item_${safetyTplItems.length + 1}`, label: "New checklist item" });
    renderSafetyTemplateEditor(safetyTplItems);
  });
  qs("safetyTplItems")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-safety-tpl-remove]");
    if (!btn) return;
    const idx = Number(btn.getAttribute("data-safety-tpl-remove"));
    safetyTplItems.splice(idx, 1);
    renderSafetyTemplateEditor(safetyTplItems);
  });
  qs("safetyItemAddBtn")?.addEventListener("click", () =>
    addSafetyEquipmentItem().catch((e) => setStatus("Add safety item error: " + e.message))
  );
  qs("safetyItemsList")?.addEventListener("click", (e) => {
    const qrBtn = e.target.closest("button[data-safety-use-qr]");
    if (qrBtn) {
      const code = String(qrBtn.getAttribute("data-safety-use-qr") || "");
      if (qs("safetyQrItemCode")) qs("safetyQrItemCode").value = code;
      generateSafetyQr().catch((err) => setStatus("QR error: " + err.message));
      return;
    }
    const inspBtn = e.target.closest("button[data-safety-open-insp]");
    if (inspBtn) {
      const code = String(inspBtn.getAttribute("data-safety-open-insp") || "");
      if (code) window.open(`./safety-inspection.html?item_code=${encodeURIComponent(code)}`, "_blank");
      return;
    }
    const pdfBtn = e.target.closest("button[data-safety-item-pdf]");
    if (pdfBtn) {
      openSafetyItemInspectionPdf(pdfBtn.getAttribute("data-safety-item-pdf"))
        .catch((err) => setStatus("Safety PDF error: " + err.message));
      return;
    }
    const rmBtn = e.target.closest("button[data-safety-remove-item]");
    if (rmBtn) {
      removeSafetyEquipmentItem(Number(rmBtn.getAttribute("data-safety-remove-item") || 0))
        .catch((err) => setStatus("Remove error: " + err.message));
    }
  });
  qs("safetyItemsList")?.addEventListener("change", (e) => {
    const chk = e.target.closest("input[data-safety-report-select]");
    if (!chk) return;
    const code = String(chk.getAttribute("data-safety-report-select") || "").trim().toUpperCase();
    if (!code) return;
    if (chk.checked) safetyReportSelectedCodes.add(code);
    else safetyReportSelectedCodes.delete(code);
  });
  qs("safetyQrGenerate")?.addEventListener("click", () =>
    generateSafetyQr().catch((e) => setStatus("Safety QR error: " + e.message))
  );
  qs("safetyQrPrint")?.addEventListener("click", printSafetyQr);
  qs("safetyQrPrintSheet")?.addEventListener("click", () =>
    printAllSafetyQrSheet().catch((e) => setStatus("Safety QR sheet error: " + e.message))
  );
  qs("safetyQrPreset")?.addEventListener("change", applySafetyQrSheetPreset);
  ["safetyQrCols", "safetyQrSizeMm", "safetyQrCellMm", "safetyQrGapMm"].forEach((id) => {
    qs(id)?.addEventListener("input", () => {
      const preset = qs("safetyQrPreset");
      if (preset && preset.value !== "custom") preset.value = "custom";
    });
  });
  qs("safetyPdfRegisterBtn")?.addEventListener("click", () =>
    openSafetyRegisterPdf(false).catch((e) => setStatus("Safety PDF error: " + e.message))
  );
  qs("safetyPdfBlankBtn")?.addEventListener("click", () =>
    openSafetyRegisterPdf(true).catch((e) => setStatus("Safety PDF error: " + e.message))
  );
  qs("safetyInspectionReportAllBtn")?.addEventListener("click", () =>
    openSafetyInspectionReportPdf(false).catch((e) => setStatus("Safety report error: " + e.message))
  );
  qs("safetyInspectionReportSelectedBtn")?.addEventListener("click", () =>
    openSafetyInspectionReportPdf(true).catch((e) => setStatus("Safety report error: " + e.message))
  );
  initSafetyAdminPanel().catch(() => {});
  initOfflineQueueAdminPanel();
  initTelematicsAdminPanel().catch(() => {});
  initCartrackAdminPanel().catch(() => {});
  initCartrackTrackingTab();

  qs("cartrackSaveSettingsBtn")?.addEventListener("click", () =>
    saveCartrackAdminSettings().catch((e) => setCartrackAdminResult(String(e.message || e), false))
  );
  qs("cartrackTestBtn")?.addEventListener("click", () => testCartrackConnection().catch(() => {}));
  qs("cartrackRunMorningBtn")?.addEventListener("click", () => runCartrackMorningNow().catch(() => {}));
  qs("unitechSaveSettingsBtn")?.addEventListener("click", () =>
    saveUnitechAdminSettings().catch((e) => setUnitechAdminResult(String(e.message || e), false))
  );
  qs("unitechTestBtn")?.addEventListener("click", () => testUnitechConnection().catch(() => {}));
  qs("gpsLinkSaveBtn")?.addEventListener("click", () => saveGpsVehicleLink().catch(() => {}));
  qs("gpsLinkRefreshBtn")?.addEventListener("click", () => loadGpsVehicleLinksAdmin().catch(() => {}));
  qs("gpsLinkApplyBtn")?.addEventListener("click", () => applyGpsVehicleLinks().catch(() => {}));
  qs("gpsVehicleLinksList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-gps-link-delete]");
    if (!btn) return;
    deleteGpsVehicleLink(btn.getAttribute("data-gps-link-delete")).catch(() => {});
  });
  qs("gpsVehicleLinkSuggestions")?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-gps-link-prefill]");
    if (!btn) return;
    prefillGpsVehicleLinkForm(btn.getAttribute("data-gps-link-prefill"), btn.getAttribute("data-gps-link-source"));
  });
  qs("cartrackSyncBtn")?.addEventListener("click", () => syncCartrackNow().catch((e) => setStatus(String(e.message || e))));
  qs("cartrackDashShowAll")?.addEventListener("change", () => {
    renderCartrackFleetTable(
      cartrackDashboardFleetCache.fleet,
      cartrackDashboardFleetCache.speedingToday,
      { showAll: Boolean(qs("cartrackDashShowAll")?.checked) }
    );
  });
  qs("cartrackMorningPdfBtn")?.addEventListener("click", () =>
    openCartrackMorningPdf().catch((e) => setStatus(String(e.message || e)))
  );
  qs("cartrackMorningEmailBtn")?.addEventListener("click", () =>
    emailCartrackMorningReport().catch((e) => setStatus(String(e.message || e)))
  );
  qs("cartrackTrackMorningPdfBtn")?.addEventListener("click", () =>
    openCartrackMorningPdf().catch((e) => setStatus(String(e.message || e)))
  );
  const onCartrackSpeedReportDateChange = (e) => {
    const d = syncCartrackSpeedReportDateInputs(e.target?.value);
    loadCartrackSpeedingEvents(d, { useCache: d === todayLocalYmd() }).catch(() => {});
    loadCartrackSpeedingEvents(d, {
      hostId: "cartrackTrackSpeedingEventsHost",
      countId: "cartrackTrackSpeedingEventsCount",
      panelId: "cartrackTrackSpeedingEventsPanel",
      useCache: d === todayLocalYmd(),
    }).catch(() => {});
  };
  qs("cartrackSpeedReportDate")?.addEventListener("change", onCartrackSpeedReportDateChange);
  qs("cartrackTrackSpeedReportDate")?.addEventListener("change", onCartrackSpeedReportDateChange);
  qs("cartrackKpiSpeeding")?.closest(".kpi-pill")?.addEventListener("click", () => {
    openCartrackSpeedingEventsPanel();
    loadCartrackSpeedingEvents(todayLocalYmd(), { useCache: true }).catch(() => {});
  });
  initCartrackSpeedReportDates();

  qs("telemSaveDeviceBtn")?.addEventListener("click", () =>
    saveTelematicsDevice().catch((e) => setTelemAdminResult(String(e.message || e)))
  );
  qs("telemRefreshDevicesBtn")?.addEventListener("click", () =>
    loadTelematicsAdminDevices().catch((e) => setTelemAdminResult(String(e.message || e)))
  );
  qs("telemShowInactive")?.addEventListener("change", () =>
    loadTelematicsAdminDevices().catch(() => {})
  );
  qs("telematicsRefreshBtn")?.addEventListener("click", () =>
    loadTelematicsTab().catch((e) => setStatus("Telematics refresh error: " + (e.message || e)))
  );
  qs("telemDevicesList")?.addEventListener("click", (e) => {
    const editBtn = e.target.closest("button[data-telem-edit]");
    if (editBtn) {
      fillTelematicsDeviceForm({
        assetCode: editBtn.getAttribute("data-telem-asset"),
        deviceSerial: "",
        unitModel: editBtn.getAttribute("data-telem-model") || "FSC650",
        externalId: "",
        replaceFaulty: true,
      });
      setTelemAdminResult(`Enter new serial for ${editBtn.getAttribute("data-telem-asset")} (was ${editBtn.getAttribute("data-telem-serial")}).`);
      return;
    }
    const deactBtn = e.target.closest("button[data-telem-deactivate]");
    if (deactBtn) {
      deactivateTelematicsDeviceAdmin(
        deactBtn.getAttribute("data-telem-deactivate"),
        deactBtn.getAttribute("data-telem-asset-label")
      ).catch((err) => setTelemAdminResult(String(err.message || err)));
    }
  });

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

  // Assets
  qs("loadHistory")?.addEventListener("click", () =>
    loadAssetHistory().catch((e) => setStatus("History error: " + e.message))
  );
  qs("downloadHistoryPdf")?.addEventListener("click", downloadAssetHistoryPdf);

  qs("historyList")?.addEventListener("click", async (e) => {
    const btn = e.target.closest("button[data-ops-slip-pdf]");
    if (!btn) return;
    e.preventDefault();
    const id = btn.getAttribute("data-ops-slip-pdf");
    if (!id) return;
    try {
      setStatus("Opening slip PDF...");
      await openAuthedPdf(`${API}/api/breakdown-ops/slips/${encodeURIComponent(id)}/pdf`);
      setStatus("PDF opened ✅");
    } catch (err) {
      setStatus("PDF error: " + (err.message || err));
    }
  });

  qs("showArchived")?.addEventListener("change", () => {
    loadAssetsFleet().catch(() => {});
  });

  qs("downloadAssetsCostCentersXlsx")?.addEventListener("click", () => downloadAssetsCostCentersXlsx());

  qs("saveOperatingBudget")?.addEventListener("click", () =>
    saveOperatingBudget().catch((e) => setStatus("Operating budget error: " + e.message))
  );
  qs("savePlantHireBudget")?.addEventListener("click", () =>
    savePlantHireBudget().catch((e) => setStatus("Plant hire budget error: " + e.message))
  );
  qs("savePlantHireRates")?.addEventListener("click", () =>
    savePlantHireRates().catch((e) => setStatus("Plant hire rates error: " + e.message))
  );
  qs("plantHireBudgetMonth")?.addEventListener("change", () => loadPlantHireBudgetStatus().catch(() => {}));
  qs("plantHireAssetSelect")?.addEventListener("change", () => {
    const code = String(qs("plantHireAssetSelect")?.value || "").trim();
    const row = plantHireRegisterCache.find((r) => r.asset_code === code);
    fillPlantHireRateFields(row || {});
    syncPlantHireAssetLabel(code);
  });

  qs("assetsFleetFilter")?.addEventListener("input", () => {
    renderAssetFleetGrid(assetsFleetCache);
  });

  qs("assetFleetGrid")?.addEventListener("click", (e) => {
    const card = e.target.closest(".asset-fleet-card");
    if (!card) return;
    const code = card.dataset.assetCode;
    if (!code) return;
    selectAssetCard(code, { loadHistory: true, scroll: true }).catch((err) =>
      setStatus("Asset select error: " + (err.message || err))
    );
  });

  qs("btnArchiveAsset")?.addEventListener("click", () =>
    archiveSelectedAsset().catch((e) => setStatus("Archive error: " + e.message))
  );

  qs("btnUnarchiveAsset")?.addEventListener("click", () =>
    unarchiveSelectedAsset().catch((e) => setStatus("Unarchive error: " + e.message))
  );
  qs("saveContractorAsset")?.addEventListener("click", () =>
    saveContractorAsset().catch((e) => setStatus("Contractor asset save error: " + e.message))
  );
  qs("assetAllocSelect")?.addEventListener("change", () => loadAssetAllocationForm());
  qs("saveAssetAllocationBtn")?.addEventListener("click", () =>
    saveAssetAllocation().catch((e) => setStatus("Asset allocation error: " + e.message))
  );
  qs("refreshAssetAllocationBtn")?.addEventListener("click", () =>
    populateAssetAllocSelect().catch((e) => setStatus("Asset allocation refresh error: " + e.message))
  );
  populateAssetAllocSelect().catch(() => {});

  loadAssetsFleet().catch(() => {});
  const saDate = qs("saDate");
  if (saDate) saDate.value = new Date().toISOString().slice(0, 10);
  const mlDate = qs("mlDate");
  if (mlDate) mlDate.value = new Date().toISOString().slice(0, 10);
  const fuelDate = qs("fuelDate");
  if (fuelDate) fuelDate.value = new Date().toISOString().slice(0, 10);
  const lubeRange = getDefaultLubeRange();
  const lubeStart = qs("lubeStart");
  const lubeEnd = qs("lubeEnd");
  if (lubeStart && !lubeStart.value) lubeStart.value = lubeRange.start;
  if (lubeEnd && !lubeEnd.value) lubeEnd.value = lubeRange.end;
  const lubeStockMonth = qs("lubeStockMonth");
  if (lubeStockMonth && !lubeStockMonth.value) {
    lubeStockMonth.value = (lubeStart?.value || lubeRange.end).slice(0, 7);
  }
  const fuelStart = qs("fuelStart");
  const fuelEnd = qs("fuelEnd");
  if (fuelStart && !fuelStart.value) fuelStart.value = lubeRange.start;
  if (fuelEnd && !fuelEnd.value) fuelEnd.value = lubeRange.end;
  const fuelTolerance = qs("fuelTolerance");
  if (fuelTolerance && !fuelTolerance.value) fuelTolerance.value = "0.15";
  const fuelSnapStart = qs("fuelSnapStart");
  const fuelSnapEnd = qs("fuelSnapEnd");
  if (fuelSnapStart && !fuelSnapStart.value) fuelSnapStart.value = fuelStart?.value || lubeRange.start;
  if (fuelSnapEnd && !fuelSnapEnd.value) fuelSnapEnd.value = fuelEnd?.value || date.value;
  const opDate = qs("opDate");
  if (opDate && !opDate.value) opDate.value = new Date().toISOString().slice(0, 10);
  const opSiteDate = qs("opSiteDate");
  if (opSiteDate && !opSiteDate.value) opSiteDate.value = new Date().toISOString().slice(0, 10);
  const opSiteDashDate = qs("opSiteDashDate");
  if (opSiteDashDate && !opSiteDashDate.value) opSiteDashDate.value = opSiteDate?.value || new Date().toISOString().slice(0, 10);
  const opDelayDate = qs("opDelayDate");
  if (opDelayDate && !opDelayDate.value) opDelayDate.value = new Date().toISOString().slice(0, 10);
  const opTargetDate = qs("opTargetDate");
  if (opTargetDate && !opTargetDate.value) opTargetDate.value = new Date().toISOString().slice(0, 10);
  const opFrom = qs("opFrom");
  const opTo = qs("opTo");
  if (opFrom && !opFrom.value) opFrom.value = new Date(Date.now() - 1000 * 60 * 60 * 24 * 30).toISOString().slice(0, 10);
  if (opTo && !opTo.value) opTo.value = new Date().toISOString().slice(0, 10);
  const dpDate = qs("dpDate");
  if (dpDate && !dpDate.value) dpDate.value = new Date().toISOString().slice(0, 10);
  const dpFrom = qs("dpFrom");
  const dpTo = qs("dpTo");
  if (dpFrom && !dpFrom.value) dpFrom.value = new Date(Date.now() - 1000 * 60 * 60 * 24 * 7).toISOString().slice(0, 10);
  if (dpTo && !dpTo.value) dpTo.value = new Date().toISOString().slice(0, 10);
  const qFrom = qs("qFrom");
  const qTo = qs("qTo");
  if (qFrom && !qFrom.value) qFrom.value = new Date(Date.now() - 1000 * 60 * 60 * 24 * 30).toISOString().slice(0, 10);
  if (qTo && !qTo.value) qTo.value = new Date().toISOString().slice(0, 10);
  const costMonth = qs("costMonth");
  if (costMonth && !costMonth.value) costMonth.value = new Date().toISOString().slice(0, 7);
  loadStoreAllocations().catch(() => {});
  loadStockOnHandPage().catch(() => {});
  loadInventoryControl().catch(() => {});
  loadLocations().catch(() => {});
  loadStockBins().catch(() => {});
  loadStockDepth().catch(() => {});
  loadReplenishmentSuggestions().catch(() => {});
  loadCycleSessions().catch(() => {});
  loadLubeStockOnHand().catch(() => {});
  loadLubeReorderAlerts().catch(() => {});
  applyDefaultLocationsToInputs();
  loadBinCodeOptionsForLocation(qs("saLocation")?.value || "", "saBinCodeOptions").catch(() => {});
  loadBinCodeOptionsForLocation(qs("msLocation")?.value || "", "msBinCodeOptions").catch(() => {});
  updateManualStockCostRowVisibility();
  updateManualStockPartDesc();
  updateManualLubePartDesc();
  updateReceiveLubePartDesc();
  updateLubeMinPartDesc();
  loadLubeAnalytics().catch(() => {});
  loadLubeMappings().catch(() => {});
  setProcurementChainInputsFromConfig();
  updateProcurementChainPreview();
  setProcurementKpiFilter("all");
  const prJournalStart = qs("prJournalStart");
  const prJournalEnd = qs("prJournalEnd");
  if (prJournalStart && !prJournalStart.value) prJournalStart.value = new Date(Date.now() - 1000 * 60 * 60 * 24 * 30).toISOString().slice(0, 10);
  if (prJournalEnd && !prJournalEnd.value) prJournalEnd.value = new Date().toISOString().slice(0, 10);
  loadRequisitions().catch(() => {});
  loadPurchaseOrders().catch(() => {});
  loadProcurementExceptions().catch(() => {});
  loadOperations().catch(() => {});
  loadSiteZones().catch(() => {});
  loadSiteDailyEntries().catch(() => {});
  loadSiteTargets().catch(() => {});
  loadSiteDelays().catch(() => {});
  loadSiteDashboard().catch(() => {});
  loadDispatchTrips().catch(() => {});
  loadQualityCenter().catch(() => {});
  // Fuel endpoints are expensive on large datasets; keep initial page load responsive.
  // Users can load these manually from the Fuel tab buttons.
  if (getSessionRoles().some((r) => ["admin", "supervisor"].includes(r))) {
    loadCostSettings().catch(() => {});
  }
  loadLegalDepartments().catch(() => {});
  loadLegalDocs().catch(() => {});
  loadLegalExpiry().catch(() => {});
  if (getSessionRoles().some((r) => ["admin", "supervisor"].includes(r))) {
    loadAuditLogs().catch(() => {});
    loadApprovalRequests().catch(() => {});
    loadSmtpSettings().catch(() => {});
    loadPushNotificationSettings().catch(() => {});
    loadPdfReportSettings().catch(() => {});
    loadBackupFiles().catch(() => {});
  }
  loadCodePickers().catch(() => {});
  populateThresholdInputs();
  populateLdvPrestartThresholdInputs();
  loadDashboard().catch((e) => setStatus("Dashboard error: " + e.message));

  const legalList = qs("legalList");
  if (legalList) {
    legalList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const dl = target.getAttribute("data-legal-download-id");
      const ar = target.getAttribute("data-legal-archive-id");
      const active = target.getAttribute("data-legal-active");
      const stId = target.getAttribute("data-legal-status-id");
      const st = target.getAttribute("data-legal-status");
      const actionsId = target.getAttribute("data-legal-actions-id");
      if (dl) {
        downloadLegalDoc(dl);
        return;
      }
      if (actionsId) {
        showLegalActions(actionsId);
        return;
      }
      if (stId && st) {
        setLegalStatus(stId, st);
        return;
      }
      if (ar && active != null) {
        archiveLegalDoc(ar, Number(active) === 1);
      }
    });
  }

  const approvalList = qs("approvalList");
  if (approvalList) {
    approvalList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const aId = target.getAttribute("data-approval-approve-id");
      const rId = target.getAttribute("data-approval-reject-id");
      if (aId) {
        decideApprovalRequest(aId, "approve");
        return;
      }
      if (rId) {
        decideApprovalRequest(rId, "reject");
      }
    });
  }

  const approvalKpiStrip = qs("approvalKpiStrip");
  if (approvalKpiStrip) {
    approvalKpiStrip.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const btn = target.closest("[data-approval-kpi-filter]");
      if (!(btn instanceof HTMLElement)) return;
      const filter = btn.getAttribute("data-approval-kpi-filter");
      if (filter == null) return;
      const statusEl = qs("approvalStatus");
      if (statusEl) statusEl.value = filter;
      loadApprovalRequests().catch((e) => setStatus("Approvals error: " + e.message));
    });
  }

  const lubeAnalyticsList = qs("lubeAnalyticsList");
  if (lubeAnalyticsList) {
    lubeAnalyticsList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const btn = target.closest("[data-map-oil-key]");
      if (!(btn instanceof HTMLElement)) return;
      const oilKey = String(btn.getAttribute("data-map-oil-key") || "").trim();
      const partCode = String(btn.getAttribute("data-map-part-code") || "").trim();
      const oilEl = qs("lubeMapOilKey");
      const partEl = qs("lubeMapPartCode");
      if (oilEl) oilEl.value = oilKey;
      if (partEl) partEl.value = partCode;
      setStatus("Mapping fields pre-filled from selected lube row.");
    });
  }

  const procurementList = qs("procurementList");
  if (procurementList) {
    procurementList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const advanceId = target.getAttribute("data-pr-advance-id");
      const advanceStatus = target.getAttribute("data-pr-advance-status");
      const submitId = target.getAttribute("data-pr-submit-id");
      const finalizeId = target.getAttribute("data-pr-finalize-id");
      const postId = target.getAttribute("data-pr-post-id");
      const routeId = target.getAttribute("data-pr-route-id");
      const approveId = target.getAttribute("data-pr-approve-id");
      const receiveId = target.getAttribute("data-pr-receive-id");
      const receiveHalfId = target.getAttribute("data-pr-receive-half-id");
      const receiveFullId = target.getAttribute("data-pr-receive-full-id");
      const createPoId = target.getAttribute("data-pr-create-po-id");
      const outstanding = target.getAttribute("data-pr-outstanding");
      const duplicateJson = target.getAttribute("data-pr-duplicate");
      const openApprovalId = target.getAttribute("data-pr-open-approval-id");
      if (advanceId && advanceStatus) {
        advanceRequisitionStage(advanceId, advanceStatus).catch((e) => setStatus(`Advance failed: ${e.message || e}`));
        return;
      }
      if (finalizeId) {
        fetchJson(`${API}/api/procurement/requisitions/${finalizeId}/finalize`, { method: "POST", headers: { "Content-Type": "application/json" } })
          .then((res) => {
            setText("procurementResult", JSON.stringify(res, null, 2));
            return loadRequisitions();
          })
          .catch((e) => setStatus(`Finalize failed: ${e.message || e}`));
        return;
      }
      if (postId) {
        fetchJson(`${API}/api/procurement/requisitions/${postId}/post`, { method: "POST", headers: { "Content-Type": "application/json" } })
          .then((res) => {
            setText("procurementResult", JSON.stringify(res, null, 2));
            return loadRequisitions();
          })
          .catch((e) => setStatus(`Post failed: ${e.message || e}`));
        return;
      }
      if (routeId) {
        launchApprovalRouteForRequisition(routeId)
          .then(() => loadRequisitions())
          .catch((e) => setStatus(`Route failed: ${e.message || e}`));
        return;
      }
      if (approveId) {
        approveCurrentStepForRequisition(approveId)
          .then(() => loadRequisitions())
          .catch((e) => setStatus(`Approve failed: ${e.message || e}`));
        return;
      }
      if (submitId) {
        fetchJson(`${API}/api/procurement/requisitions/${submitId}/finalize`, { method: "POST", headers: { "Content-Type": "application/json" } })
          .then(() => fetchJson(`${API}/api/procurement/requisitions/${submitId}/post`, { method: "POST", headers: { "Content-Type": "application/json" } }))
          .then(() =>
            fetchJson(`${API}/api/procurement/requisitions/${submitId}/approvers`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ approvers: [{ name: "approver1" }] }),
            })
          )
          .then(() => fetchJson(`${API}/api/procurement/requisitions/${submitId}/send-approval`, { method: "POST", headers: { "Content-Type": "application/json" } }))
          .then((res) => {
            setText("procurementResult", JSON.stringify(res, null, 2));
            return loadRequisitions();
          })
          .catch((e) => setStatus(`Quick send failed: ${e.message || e}`));
        return;
      }
      if (receiveId) {
        requestRequisitionReceive(receiveId);
        return;
      }
      if (receiveHalfId) {
        requestRequisitionReceiveHalf(receiveHalfId, Number(outstanding || 0));
        return;
      }
      if (receiveFullId) {
        requestRequisitionReceiveFull(receiveFullId, Number(outstanding || 0));
        return;
      }
      if (createPoId) {
        createPoFromRequisition(createPoId).catch((e) => setStatus(`Create PO failed: ${e.message || e}`));
        return;
      }
      if (duplicateJson) {
        duplicateRequisitionFromRow(duplicateJson);
        return;
      }
      if (openApprovalId) {
        const statusEl = qs("approvalStatus");
        const moduleEl = qs("approvalModule");
        const actionEl = qs("approvalAction");
        if (statusEl) statusEl.value = "";
        if (moduleEl) moduleEl.value = "procurement";
        if (actionEl) actionEl.value = "";
        switchTab("approvals");
        loadApprovalRequests().catch(() => {});
        setStatus(`Showing approvals. Latest request id: #${openApprovalId}`);
      }
    });
  }

  const prPoList = qs("prPoList");
  if (prPoList) {
    prPoList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const openId = target.getAttribute("data-pr-po-open");
      const approveId = target.getAttribute("data-pr-po-approve");
      const sendId = target.getAttribute("data-pr-po-send");
      if (openId) {
        openPurchaseOrder(openId).catch((e) => setStatus(`PO detail error: ${e.message || e}`));
        return;
      }
      if (approveId) {
        approvePurchaseOrder(approveId).catch((e) => setStatus(`PO approve error: ${e.message || e}`));
        return;
      }
      if (sendId) {
        sendPurchaseOrder(sendId).catch((e) => setStatus(`PO send error: ${e.message || e}`));
      }
    });
  }

  const prExList = qs("prExceptionsList");
  if (prExList) {
    prExList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const resolveId = target.getAttribute("data-pr-ex-resolve");
      if (resolveId) {
        resolveProcurementException(resolveId).catch((e) => setStatus(`Exception resolve error: ${e.message || e}`));
      }
    });
  }

  const sbCycleSessionsList = qs("sbCycleSessionsList");
  if (sbCycleSessionsList) {
    sbCycleSessionsList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const submitId = target.getAttribute("data-sb-cs-submit");
      const approveId = target.getAttribute("data-sb-cs-approve");
      const countOneId = target.getAttribute("data-sb-cs-countone");
      if (submitId) {
        submitCycleSession(submitId).catch((e) => setStatus(`Cycle submit error: ${e.message || e}`));
        return;
      }
      if (approveId) {
        approveCycleSession(approveId).catch((e) => setStatus(`Cycle approve error: ${e.message || e}`));
        return;
      }
      if (countOneId) {
        addOnePartCountToSession(countOneId).catch((e) => setStatus(`Cycle line upsert error: ${e.message || e}`));
      }
    });
  }

  ["sfPlan", "sfReview", "sfRoute", "sfApprove", "sfPoReady", "sfReceive"].forEach((laneId) => {
    const lane = qs(laneId);
    if (!lane) return;
    lane.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const advanceId = target.getAttribute("data-pr-advance-id");
      const advanceStatus = target.getAttribute("data-pr-advance-status");
      if (!advanceId || !advanceStatus) return;
      advanceRequisitionStage(advanceId, advanceStatus).catch((e) => setStatus(`Advance failed: ${e.message || e}`));
    });
  });

  const dispatchList = qs("dispatchList");
  if (dispatchList) {
    dispatchList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const id = target.getAttribute("data-dp-status-id");
      const next = target.getAttribute("data-dp-next");
      if (!id || !next) return;
      updateDispatchTripStatus(id, next).catch((e) => setStatus(`Dispatch status update failed: ${e.message || e}`));
    });
  }
  const dispatchExceptionsList = qs("dispatchExceptionsList");
  if (dispatchExceptionsList) {
    dispatchExceptionsList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const id = target.getAttribute("data-dp-ex-id");
      const next = target.getAttribute("data-dp-ex-next");
      if (!id || !next) return;
      resolveDispatchException(id, next).catch((e) => setStatus(`Dispatch exception update failed: ${e.message || e}`));
    });
  }
  const qualityList = qs("qualityList");
  if (qualityList) {
    qualityList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const resolveBtn = target.closest("[data-q-resolve]");
      if (resolveBtn instanceof HTMLElement) {
        const mode = resolveBtn.getAttribute("data-q-resolve");
        const entity = resolveBtn.getAttribute("data-q-entity");
        const date = resolveBtn.getAttribute("data-q-date");
        resolveQualityIssueNow(mode, entity, date).catch((e) => setStatus(`Quality resolve failed: ${e.message || e}`));
        return;
      }
      const btn = target.closest("[data-q-fix]");
      if (!(btn instanceof HTMLElement)) return;
      const type = btn.getAttribute("data-q-type");
      const asset = btn.getAttribute("data-q-asset");
      const entity = btn.getAttribute("data-q-entity");
      const date = btn.getAttribute("data-q-date");
      openQualityFix(type, asset, entity, date);
    });
  }

  ["sfCountPlan", "sfCountReview", "sfCountRoute", "sfCountApprove", "sfCountPoReady", "sfCountReceive"].forEach((id) => {
    const btn = qs(id);
    if (!btn) return;
    btn.addEventListener("click", () => {
      const statusEl = qs("prStatusFilter");
      if (id === "sfCountReceive") {
        if (statusEl) statusEl.value = "";
        setProcurementKpiFilter("receive_set");
      } else {
        const status = String(btn.getAttribute("data-sf-status") || "").trim();
        if (statusEl) statusEl.value = status;
        setProcurementKpiFilter("all");
      }
      loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message));
    });
  });

  loadDocHeaders().catch(() => {});
  loadDocDrafts().catch(() => {});
}

function applyI18n() {
  const map = {
    docsTitle: t("docsTitle"),
    docsSubtitle: t("docsSubtitle"),
    docsHeaderTitle: t("docsHeaderTitle"),
    docsDraftTitle: t("docsDraftTitle"),
  };
  Object.entries(map).forEach(([id, text]) => {
    const el = qs(id);
    if (el) el.textContent = text;
  });
  const statusEl = qs("status");
  if (statusEl && (!statusEl.textContent || statusEl.textContent.trim() === "Ready.")) {
    statusEl.textContent = t("statusReady");
  }
}

async function saveDocHeader() {
  const payload = {
    name: qs("docHeaderName")?.value || "",
    site_name: qs("docHeaderSite")?.value || "",
    department: qs("docHeaderDepartment")?.value || "",
    prepared_by: qs("docHeaderPreparedBy")?.value || "",
    approved_by: qs("docHeaderApprovedBy")?.value || "",
    revision: qs("docHeaderRevision")?.value || "",
  };
  const res = await fetchJson(`${API}/api/docs/headers`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setStatus(`Header saved (#${res.id}).`);
  await loadDocHeaders();
}

async function loadDocHeaders() {
  const listEl = qs("docHeaderList");
  if (!listEl) return;
  const res = await fetchJson(`${API}/api/docs/headers`);
  if (!res.rows?.length) {
    listEl.innerHTML = `<div class="item"><small>No header profiles yet.</small></div>`;
    return;
  }
  listEl.innerHTML = "";
  res.rows.forEach((r) => {
    const d = document.createElement("div");
    d.className = "item";
    d.innerHTML = `<b>#${r.id} ${r.name}</b> - ${r.site_name || "-"} / ${r.department || "-"} <small>(Rev: ${r.revision || "-"})</small>`;
    d.addEventListener("click", () => {
      const idEl = qs("docHeaderId");
      if (idEl) idEl.value = String(r.id);
    });
    listEl.appendChild(d);
  });
  const idEl = qs("docHeaderId");
  if (idEl && !Number(idEl.value || 0) && res.rows[0]?.id) {
    idEl.value = String(res.rows[0].id);
  }
}

async function generateDocDraft() {
  const payload = {
    header_id: Number(qs("docHeaderId")?.value || 0),
    doc_type: qs("docType")?.value || "SOP",
    title: qs("docTitle")?.value || "",
    language: qs("docLanguage")?.value || getLang(),
    scope: qs("docScope")?.value || "",
    hazards: qs("docHazards")?.value || "",
    controls: qs("docControls")?.value || "",
    extra_notes: qs("docInputs")?.value || "",
  };
  const res = await fetchJson(`${API}/api/docs/draft-generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const out = qs("docDraftOutput");
  if (out) out.textContent = res.draft_text || "";
  const idEl = qs("docDraftId");
  if (idEl) idEl.value = String(res.id || "");
  setStatus(`Draft generated (#${res.id}).`);
  await loadDocDrafts();
}

async function generateDocDraftFromRequest(requestArg) {
  const requestText = String(requestArg || qs("aiSmartPrompt")?.value || "").trim();
  if (!requestText) return alert("Enter what document you want first.");
  let headerId = Number(qs("docHeaderId")?.value || 0);
  if (!headerId) {
    try {
      const hdr = await fetchJson(`${API}/api/docs/headers`);
      if (Array.isArray(hdr.rows) && hdr.rows[0]?.id) {
        headerId = Number(hdr.rows[0].id);
        const idEl = qs("docHeaderId");
        if (idEl) idEl.value = String(headerId);
      }
    } catch (_) {
      // Backend will still fallback to latest header when possible.
    }
  }
  const payload = {
    request_text: requestText,
    header_id: headerId,
    doc_type: qs("docType")?.value || "SOP",
    title: qs("docTitle")?.value || requestText.slice(0, 80),
    language: qs("docLanguage")?.value || getLang(),
    scope: qs("docScope")?.value || "",
    hazards: qs("docHazards")?.value || "",
    controls: qs("docControls")?.value || "",
  };
  const res = await fetchJson(`${API}/api/docs/draft-generate-request`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const out = qs("docDraftOutput");
  if (out) {
    const ctx = Array.isArray(res.related_docs) && res.related_docs.length
      ? `\n\n[Related docs used]\n${res.related_docs.map((d) => `#${d.id} ${d.title}`).join("\n")}`
      : "";
    out.textContent = (res.draft_text || "") + ctx;
  }
  const idEl = qs("docDraftId");
  if (idEl) idEl.value = String(res.id || "");
  setStatus(`Draft generated from request (#${res.id})${res.ai_used ? " with AI" : " (template fallback)"}.`);
  await loadDocDrafts();
}

function inferDocTypeFromPrompt(prompt) {
  const s = String(prompt || "").toLowerCase();
  if (s.includes("checklist")) return "Checklist";
  if (s.includes("method statement")) return "Method Statement";
  if (s.includes("site instruction")) return "Site Instruction";
  if (s.includes("risk")) return "Risk Note";
  if (s.includes("sop") || s.includes("procedure")) return "SOP";
  return "";
}

function parseMachineProblemFromPrompt(prompt) {
  const src = String(prompt || "").trim();
  if (!src) return { machine: "", problem: "" };
  const m = src.match(/^(.+?)\s+(?:has|have|with|showing|shows|no)\s+(.+)$/i);
  if (m) {
    const machine = String(m[1] || "").replace(/\s+$/, "").trim();
    const problem = src.slice(machine.length).replace(/^\s*(has|have|with|showing|shows)?\s*/i, "").trim();
    return { machine, problem };
  }
  return { machine: "", problem: src };
}

async function runAiSmart() {
  const out = qs("askJakesOutput");
  const smartPrompt = String(qs("aiSmartPrompt")?.value || "").trim();

  if (!smartPrompt) {
    alert("Enter a question or document request first.");
    return;
  }

  if (out) {
    out.textContent = "⏳ Thinking...";
    out.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }

  const faultKeywords = /(fault|error|not working|no\s+hydraulic|no\s+hydraulics|no\s+start|won't start|wont start|leak|overheat|pressure|engine|starter|battery|transmission|brake)/i;
  const docKeywords = /(sop|checklist|method statement|site instruction|risk note|document|procedure|policy|template)/i;

  try {
    if (faultKeywords.test(smartPrompt) && !docKeywords.test(smartPrompt)) {
      const parsed = parseMachineProblemFromPrompt(smartPrompt);
      await askJakes({ machine: parsed.machine, problem: parsed.problem, context: "" });
      return;
    }

    const inferred = inferDocTypeFromPrompt(smartPrompt);
    const titleEl = qs("docTitle");
    if (titleEl && !String(titleEl.value || "").trim()) titleEl.value = smartPrompt.slice(0, 80);
    if (inferred) {
      const typeEl = qs("docType");
      if (typeEl) typeEl.value = inferred;
    }
    await generateDocDraftFromRequest(smartPrompt);
    if (out) {
      out.textContent = "Document draft generated — see the Draft Output box below.";
    }
  } catch (err) {
    if (out) out.textContent = "❌ Error: " + (err.message || String(err));
    setStatus("Smart AI error: " + err.message);
  }
}

function applyAskJakesPreset(type) {
  const machineEl = qs("askJakesMachine");
  const problemEl = qs("askJakesProblem");
  const contextEl = qs("askJakesContext");
  if (!machineEl || !problemEl || !contextEl) return;

  if (type === "hydraulics") {
    machineEl.value = machineEl.value || "CAT 950 Loader";
    problemEl.value = "No hydraulics";
    contextEl.value = "Engine starts, steering weak, no bucket lift.";
    return;
  }
  if (type === "starting") {
    machineEl.value = machineEl.value || "CAT 950 Loader";
    problemEl.value = "Will not start";
    contextEl.value = "Battery indicator low, starter clicking.";
    return;
  }
  if (type === "overheat") {
    machineEl.value = machineEl.value || "CAT 950 Loader";
    problemEl.value = "Engine overheating";
    contextEl.value = "Temperature rises under load, fan noise normal.";
  }
}

function useAskJakesAnswerAsNotes() {
  const answer = String(qs("askJakesOutput")?.textContent || "").trim();
  if (!answer) {
    alert("Ask Jakes first to get an answer.");
    return;
  }
  const notesEl = qs("docInputs");
  if (!notesEl) return;
  const existing = String(notesEl.value || "").trim();
  notesEl.value = existing ? `${existing}\n\nAsk Jakes notes:\n${answer}` : `Ask Jakes notes:\n${answer}`;
  setStatus("Ask Jakes answer copied to draft notes.");
}

async function askJakes(override = {}) {
  const machine = String(override.machine || "").trim();
  const problem = String(override.problem || "").trim();
  const context = String(override.context || "").trim();
  const out = qs("askJakesOutput");

  if (!problem) {
    if (out) out.textContent = "❌ Please describe the machine problem.";
    return;
  }

  try {
    const res = await fetchJson(`${API}/api/docs/ai/ask`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ machine, problem, context }),
    });

    if (out) {
      out.textContent = String(res.answer || "No answer returned.");
      out.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }
    setStatus("Ask Jakes answered.");
  } catch (err) {
    if (out) out.textContent = "❌ Error: " + (err.message || String(err));
    setStatus("Ask Jakes error: " + err.message);
  }
}

function speakDocDraft() {
  const text = String(qs("docDraftOutput")?.textContent || "").trim();
  if (!text) {
    alert("Generate a draft first.");
    return;
  }
  if (!("speechSynthesis" in window)) {
    alert("Speech is not supported in this browser.");
    return;
  }
  const lang = String(qs("docLanguage")?.value || getLang() || "en").toLowerCase();
  const voiceLang = lang === "af" ? "af-ZA" : lang === "zu" ? "zu-ZA" : lang === "pt" ? "pt-PT" : "en-US";
  window.speechSynthesis.cancel();
  const utter = new SpeechSynthesisUtterance(text.slice(0, 12000));
  utter.lang = voiceLang;
  utter.rate = 1;
  utter.pitch = 1;
  window.speechSynthesis.speak(utter);
  setStatus("Speaking draft...");
}

function stopSpeakingDocDraft() {
  if ("speechSynthesis" in window) {
    window.speechSynthesis.cancel();
    setStatus("Speech stopped.");
  }
}

async function openDocDraftPdf(download = false) {
  const id = Number(qs("docDraftId")?.value || 0);
  if (!id) return alert("Enter/select a Draft ID first.");
  try {
    const check = await fetchJson(`${API}/api/docs/drafts`);
    const row = Array.isArray(check.rows) ? check.rows.find((r) => Number(r.id) === id) : null;
    const decision = String(row?.decision || "").toLowerCase();
    if (decision !== "approved") {
      setStatus(`Draft #${id} is '${decision || "pending"}'. Approve it (Yes) before PDF export.`);
      alert("Only approved documents can be exported to PDF.");
      return;
    }
    const url = `${API}/api/docs/drafts/${id}.pdf${download ? "?download=1" : ""}`;
    window.open(url, "_blank");
  } catch (e) {
    setStatus("Open PDF failed: " + (e.message || e));
  }
}

function openDocRegisterPdf(download = false) {
  const currentOnly = Boolean(qs("docRegisterCurrentOnly")?.checked);
  const params = new URLSearchParams();
  if (download) params.set("download", "1");
  if (currentOnly) params.set("current_only", "1");
  const q = params.toString();
  const url = `${API}/api/docs/register.pdf${q ? `?${q}` : ""}`;
  window.open(url, "_blank");
}

async function openDocDraftWord(download = false) {
  const id = Number(qs("docDraftId")?.value || 0);
  if (!id) return alert("Enter/select a Draft ID first.");
  try {
    const check = await fetchJson(`${API}/api/docs/drafts`);
    const row = Array.isArray(check.rows) ? check.rows.find((r) => Number(r.id) === id) : null;
    const decision = String(row?.decision || "").toLowerCase();
    if (decision !== "approved") {
      setStatus(`Draft #${id} is '${decision || "pending"}'. Approve it (Yes) before Word export.`);
      alert("Only approved documents can be exported to Word.");
      return;
    }
    const url = `${API}/api/docs/drafts/${id}.docx${download ? "?download=1" : ""}`;
    window.open(url, "_blank");
  } catch (e) {
    setStatus("Open Word failed: " + (e.message || e));
  }
}

function openDocRegisterWord(download = false) {
  const currentOnly = Boolean(qs("docRegisterCurrentOnly")?.checked);
  const params = new URLSearchParams();
  if (download) params.set("download", "1");
  if (currentOnly) params.set("current_only", "1");
  const q = params.toString();
  const url = `${API}/api/docs/register.docx${q ? `?${q}` : ""}`;
  window.open(url, "_blank");
}

async function decideDocDraft(approved) {
  const id = Number(qs("docDraftId")?.value || 0);
  if (!id) return alert("Enter/select a Draft ID first.");
  const res = await fetchJson(`${API}/api/docs/drafts/${id}/decision`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ approved }),
  });
  setStatus(`Draft #${res.id} marked ${res.decision}.`);
  await loadDocDrafts();
}

async function loadDocDrafts() {
  const listEl = qs("docDraftsList");
  if (!listEl) return;
  const res = await fetchJson(`${API}/api/docs/drafts`);
  if (!res.rows?.length) {
    listEl.innerHTML = `<div class="item"><small>No drafts yet.</small></div>`;
    return;
  }
  listEl.innerHTML = "";
  const rows = Array.isArray(res.rows) ? res.rows : [];
  const supersededByApproved = new Set(
    rows
      .filter((x) => String(x?.decision || "").toLowerCase() === "approved" && Number(x?.supersedes_draft_id || 0) > 0)
      .map((x) => Number(x.supersedes_draft_id))
  );
  const currentOnly = Boolean(qs("docDraftsCurrentOnly")?.checked);
  const visibleRows = currentOnly
    ? rows.filter((r) => {
      const decision = String(r?.decision || "").toLowerCase();
      return decision === "approved" && !supersededByApproved.has(Number(r?.id || 0));
    })
    : rows;
  if (!visibleRows.length) {
    listEl.innerHTML = `<div class="item"><small>${currentOnly ? "No current approved drafts." : "No drafts yet."}</small></div>`;
    return;
  }
  visibleRows.forEach((r) => {
    const d = document.createElement("div");
    d.className = "item";
    const decision = String(r.decision || "").toLowerCase();
    const canPdf = decision === "approved";
    const isCurrentApproved = decision === "approved" && !supersededByApproved.has(Number(r.id));
    const stateBadge = isCurrentApproved
      ? `<span class="pill blue">Current</span>`
      : `<span class="pill orange">Historical</span>`;
    const rev = r.revision_no ? `Rev ${r.revision_no}` : "Rev -";
    const supersedes = r.supersedes_draft_id ? ` | supersedes #${r.supersedes_draft_id}` : "";
    d.innerHTML = `<b>#${r.id}</b> ${r.doc_type} - ${r.title} <small>[${r.language}]</small> <span class="pill">${r.decision}</span> ${stateBadge} <small>${rev}${supersedes} | Header: ${r.header_name || "-"}</small> <button data-doc-open-pdf="${r.id}" ${canPdf ? "" : "disabled"} title="${canPdf ? "Open final PDF" : "Approve first"}">Open PDF</button>`;
    d.addEventListener("click", () => {
      const idEl = qs("docDraftId");
      if (idEl) idEl.value = String(r.id);
    });
    d.querySelector("button[data-doc-open-pdf]")?.addEventListener("click", (evt) => {
      evt.stopPropagation();
      const idEl = qs("docDraftId");
      if (idEl) idEl.value = String(r.id);
      openDocDraftPdf(false);
    });
    listEl.appendChild(d);
  });
}

function initDarkMode() {
  const toggle = document.getElementById("darkModeToggle");
  if (!toggle) return;
  
  const saved = localStorage.getItem("ironlog-theme");
  if (saved === "dark") {
    document.documentElement.setAttribute("data-theme", "dark");
  }
  
  toggle.addEventListener("click", () => {
    const isDark = document.documentElement.getAttribute("data-theme") === "dark";
    if (isDark) {
      document.documentElement.removeAttribute("data-theme");
      localStorage.setItem("ironlog-theme", "light");
    } else {
      document.documentElement.setAttribute("data-theme", "dark");
      localStorage.setItem("ironlog-theme", "dark");
    }
  });
}

// Task Management
let currentProjectFilter = "";
let currentTaskId = null;
let currentTaskView = "all";
let currentTaskSidebarActiveKey = "";
let teamMembers = [];

function getSavedTaskViews() {
  try {
    const raw = localStorage.getItem(TASK_SAVED_VIEWS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function persistSavedTaskViews(views) {
  localStorage.setItem(TASK_SAVED_VIEWS_KEY, JSON.stringify(Array.isArray(views) ? views : []));
}

function renderTeamMemberInputs() {
  const datalist = qs("teamMemberList");
  const authorSelect = qs("taskCommentAuthor");
  const mentions = qs("taskCommentMentions");
  const quickAssign = qs("taskAssignQuickPicks");
  const members = Array.isArray(teamMembers) ? teamMembers : [];
  const sorted = [...members].sort((a, b) => String(a.username || "").localeCompare(String(b.username || "")));
  if (datalist) {
    datalist.innerHTML = sorted
      .map((m) => `<option value="${escapeHtml(m.username)}">${escapeHtml(m.full_name || m.username)}</option>`)
      .join("");
  }
  if (authorSelect) {
    const me = getSessionUser();
    authorSelect.innerHTML = sorted
      .map((m) => `<option value="${escapeHtml(m.username)}">${escapeHtml(m.full_name || m.username)}</option>`)
      .join("");
    if (sorted.some((m) => m.username === me)) authorSelect.value = me;
  }
  if (mentions) {
    if (!sorted.length) {
      mentions.innerHTML = "";
      return;
    }
    mentions.innerHTML = `Tag team: ${sorted
      .slice(0, 8)
      .map(
        (m) =>
          `<button type="button" class="btn btn-secondary btn-sm" data-mention-user="${escapeHtml(m.username)}" style="margin:2px 4px 2px 0;padding:2px 8px;">@${escapeHtml(m.username)}</button>`
      )
      .join("")}`;
    mentions.querySelectorAll("[data-mention-user]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const ta = qs("newComment");
        if (!ta) return;
        ta.value = `${ta.value || ""}${ta.value ? " " : ""}@${btn.dataset.mentionUser} `;
        ta.focus();
      });
    });
  }
  if (quickAssign) {
    if (!sorted.length) {
      quickAssign.innerHTML = "";
      return;
    }
    quickAssign.innerHTML = `Quick assign: ${sorted
      .slice(0, 8)
      .map(
        (m) =>
          `<button type="button" class="btn btn-secondary btn-sm" data-assign-user="${escapeHtml(m.username)}" style="margin:2px 4px 2px 0;padding:2px 8px;">${escapeHtml(m.full_name || m.username)}</button>`
      )
      .join("")}`;
    quickAssign.querySelectorAll("[data-assign-user]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const assigned = qs("taskAssigned");
        if (!assigned) return;
        assigned.value = String(btn.dataset.assignUser || "");
        assigned.focus();
      });
    });
  }
}

async function loadTeamMembers() {
  try {
    const res = await fetchJson(`${API}/api/auth/team`);
    teamMembers = Array.isArray(res.rows) ? res.rows : [];
  } catch {
    teamMembers = [{ username: getSessionUser(), full_name: null, role: getSessionRole() }];
  }
  renderTeamMemberInputs();
}

function renderTaskWorkspaceSavedViews() {
  const wrap = qs("taskCustomViewSidebarLinks");
  if (!wrap) return;
  const views = getSavedTaskViews();
  if (!views.length) {
    wrap.innerHTML = `<div class="nav-item nav-subitem" style="pointer-events:none;"><span>No saved views</span></div>`;
    return;
  }
  wrap.innerHTML = views
    .map(
      (v) => `
      <a href="#" class="nav-item nav-subitem" data-tab="tasks" data-task-view="${escapeHtml(v.view || "all")}" data-task-assigned="${escapeHtml(v.assigned || "")}" data-task-project="${escapeHtml(v.project || "")}" data-task-priority="${escapeHtml(v.priority || "")}" data-task-status="${escapeHtml(v.status || "")}" data-active-key="tasks:saved:${escapeHtml(v.id)}">
        <span class="saved-view-label">${escapeHtml(v.name || "Saved View")}</span>
        <span class="saved-view-actions">
          <button type="button" class="task-saved-view-action" data-task-view-action="rename" data-task-view-id="${escapeHtml(v.id)}" title="Rename view">Rename</button>
          <button type="button" class="task-saved-view-action" data-task-view-action="delete" data-task-view-id="${escapeHtml(v.id)}" title="Delete view">Delete</button>
        </span>
      </a>
    `
    )
    .join("");
}

function renameSavedTaskView(viewId) {
  const id = String(viewId || "").trim();
  if (!id) return;
  const views = getSavedTaskViews();
  const item = views.find((v) => String(v.id) === id);
  if (!item) return;
  const nextName = String(prompt("Rename saved view:", item.name || "Saved View") || "").trim();
  if (!nextName) return;
  item.name = nextName;
  persistSavedTaskViews(views);
  renderTaskWorkspaceSavedViews();
}

function deleteSavedTaskView(viewId) {
  const id = String(viewId || "").trim();
  if (!id) return;
  const views = getSavedTaskViews();
  const filtered = views.filter((v) => String(v.id) !== id);
  persistSavedTaskViews(filtered);
  if (currentTaskSidebarActiveKey === `tasks:saved:${id}`) {
    currentTaskSidebarActiveKey = "";
    updateSidebarActiveState("tasks");
  }
  renderTaskWorkspaceSavedViews();
}

function renderTaskWorkspaceProjectLinks(projects = []) {
  const wrap = qs("taskProjectSidebarLinks");
  if (!wrap) return;
  if (!Array.isArray(projects) || !projects.length) {
    wrap.innerHTML = `<div class="nav-item nav-subitem" style="pointer-events:none;"><span>No projects</span></div>`;
    return;
  }
  wrap.innerHTML = projects
    .map((p) => {
      const color = escapeHtml(p.color || "#3b82f6");
      const name = escapeHtml(p.name || "");
      return `<a href="#" class="nav-item nav-subitem project-link" data-tab="tasks" data-task-view="all" data-task-project="${name}" data-active-key="tasks:project:${name}" style="--project-dot-color:${color};"><span>${name}</span></a>`;
    })
    .join("");
}

function setTaskApiIndicator(state, label) {
  const el = qs("taskApiIndicator");
  if (!el) return;
  el.classList.remove("api-indicator-ok", "api-indicator-down", "api-indicator-unknown");
  if (state === "ok") {
    el.classList.add("api-indicator-ok");
    el.textContent = label || "API online";
    return;
  }
  if (state === "down") {
    el.classList.add("api-indicator-down");
    el.textContent = label || "API unavailable";
    return;
  }
  el.classList.add("api-indicator-unknown");
  el.textContent = label || "API checking...";
}

function initTaskWorkspaceSidebar() {
  const toggle = qs("taskWorkspaceToggle");
  const links = qs("taskWorkspaceLinks");
  if (!toggle || !links) return;

  const apply = (collapsed) => {
    links.style.display = collapsed ? "none" : "";
    toggle.setAttribute("aria-expanded", collapsed ? "false" : "true");
  };

  apply(localStorage.getItem(TASK_WORKSPACE_COLLAPSED_KEY) === "1");
  toggle.addEventListener("click", () => {
    const collapsed = links.style.display !== "none";
    apply(collapsed);
    localStorage.setItem(TASK_WORKSPACE_COLLAPSED_KEY, collapsed ? "1" : "0");
  });

  links.addEventListener("click", (e) => {
    const btn = e.target?.closest?.(".task-saved-view-action");
    if (!btn) return;
    e.preventDefault();
    e.stopPropagation();
    const action = String(btn.dataset.taskViewAction || "").trim();
    const id = String(btn.dataset.taskViewId || "").trim();
    if (action === "rename") renameSavedTaskView(id);
    if (action === "delete") deleteSavedTaskView(id);
  });
}

function normalizeDateOnly(value) {
  return String(value || "").trim().slice(0, 10);
}

function taskMatchesSidebarView(task, view) {
  const v = String(view || "all").trim().toLowerCase();
  const due = normalizeDateOnly(task?.due_date);
  const today = new Date().toISOString().slice(0, 10);
  if (v === "today") return due === today;
  if (v === "upcoming") return !!due && due > today;
  if (v === "overdue") return !!due && due < today && String(task?.status || "").toLowerCase() !== "done";
  if (v === "inbox") return !String(task?.project || "").trim();
  return true;
}

function setTaskSidebarView(view, options = {}) {
  const v = String(view || "all").trim().toLowerCase();
  const assignedEl = qs("taskFilterAssigned");
  const statusEl = qs("taskFilterStatus");
  const priorityEl = qs("taskFilterPriority");
  currentTaskView = v || "all";

  if (statusEl && options.status !== undefined) statusEl.value = String(options.status || "");
  if (priorityEl && options.priority !== undefined) priorityEl.value = String(options.priority || "");
  if (options.project !== undefined) currentProjectFilter = String(options.project || "");

  if (assignedEl) {
    if (options.assigned !== undefined) {
      assignedEl.value = String(options.assigned || "");
    } else if (currentTaskView === "mine") {
      assignedEl.value = getSessionUser();
    } else if (options.clearAssigned !== false) {
      assignedEl.value = "";
    }
  }

  if (options.activeKey) {
    currentTaskSidebarActiveKey = String(options.activeKey);
  } else if (currentTaskView === "mine") {
    currentTaskSidebarActiveKey = "tasks:mine";
  } else if (["inbox", "today", "upcoming", "overdue"].includes(currentTaskView)) {
    currentTaskSidebarActiveKey = `tasks:${currentTaskView}`;
  } else if (currentProjectFilter) {
    currentTaskSidebarActiveKey = `tasks:project:${currentProjectFilter}`;
  } else {
    currentTaskSidebarActiveKey = "";
  }

  updateSidebarActiveState("tasks");
  if (options.refresh) loadTasks();
}

async function loadProjects() {
  try {
    const res = await fetchJson(`${API}/api/projects`);
    if (res.projects) {
      const tabsEl = qs("projectTabs");
      const projectSelect = qs("taskProject");
      const projectsList = qs("projectsList");
      
      // Render project tabs
      tabsEl.innerHTML = `<button class="project-tab ${!currentProjectFilter ? 'active' : ''}" data-project="">All Tasks</button>` +
        res.projects.map(p => `
          <button class="project-tab ${currentProjectFilter === p.name ? 'active' : ''}" data-project="${escapeHtml(p.name)}" style="--project-color:${escapeHtml(p.color || '#3b82f6')}">
            <span class="project-dot"></span>
            ${escapeHtml(p.name)}
            <span class="project-count">${p.task_count || 0}</span>
          </button>
        `).join("");
      
      // Render project select dropdown
      projectSelect.innerHTML = `<option value="">No Project</option>` +
        res.projects.map(p => `<option value="${escapeHtml(p.name)}">${escapeHtml(p.name)}</option>`).join("");
      
      // Render projects list in sidebar
      projectsList.innerHTML = res.projects.map(p => `
        <div class="project-item" style="border-left:3px solid ${escapeHtml(p.color || '#3b82f6')};">
          <div style="display:flex;justify-content:space-between;align-items:center;">
            <strong style="font-size:13px;">${escapeHtml(p.name)}</strong>
            <span class="muted" style="font-size:11px;">${p.task_count || 0} tasks</span>
          </div>
          ${p.description ? `<small class="muted">${escapeHtml(p.description)}</small>` : ""}
        </div>
      `).join("");
      
      // Add click handlers for project tabs
      tabsEl.querySelectorAll(".project-tab").forEach(tab => {
        tab.addEventListener("click", () => {
          currentProjectFilter = tab.dataset.project;
          currentTaskSidebarActiveKey = currentProjectFilter ? `tasks:project:${currentProjectFilter}` : "";
          loadProjects();
          loadTasks();
        });
      });
      renderTaskWorkspaceProjectLinks(res.projects);
    }
  } catch (err) {
    console.error("Failed to load projects", err);
  }
}

async function loadTasks() {
  const listEl = qs("tasksList");
  if (!listEl) return;
  
  const status = qs("taskFilterStatus")?.value || "";
  const priority = qs("taskFilterPriority")?.value || "";
  const assigned = qs("taskFilterAssigned")?.value || "";
  
  let url = `${API}/api/tasks?`;
  if (status) url += `status=${encodeURIComponent(status)}&`;
  if (priority) url += `priority=${encodeURIComponent(priority)}&`;
  if (assigned) url += `assigned=${encodeURIComponent(assigned)}&`;
  if (currentProjectFilter) url += `project=${encodeURIComponent(currentProjectFilter)}&`;
  
  try {
    const res = await fetchJson(url);
    setTaskApiIndicator("ok", "API online");
    const filteredTasks = (res.tasks || []).filter((task) => taskMatchesSidebarView(task, currentTaskView));
    if (!filteredTasks.length) {
      listEl.innerHTML = `<div class="item"><small class="muted">No tasks found.</small></div>`;
      return;
    }
    
    listEl.innerHTML = filteredTasks.map(task => {
      const priorityClass = task.priority === "high" ? "pill-red" : task.priority === "medium" ? "pill-orange" : "";
      const statusClass = task.status === "done" ? "pill-green" : task.status === "in_progress" ? "pill-blue" : "pill-gray";
      const dueClass = task.due_date && task.status !== "done" ? (new Date(task.due_date) < new Date() ? "text-danger" : "") : "";
      const checked = task.status === "done" ? "checked" : "";
      const strikethrough = task.status === "done" ? "text-decoration:line-through;opacity:0.6" : "";
      
      return `<div class="item task-item ${currentTaskId === task.id ? 'active' : ''}" data-task-id="${task.id}" style="border-left:3px solid ${task.priority === 'high' ? '#dc2626' : task.priority === 'medium' ? '#d97706' : '#16a34a'}; padding-left:12px; margin-bottom:8px; cursor:pointer;">
        <div style="display:flex; align-items:flex-start; gap:12px;">
          <input type="checkbox" ${checked} data-task-toggle="${task.id}" style="margin-top:4px;" onclick="event.stopPropagation();" />
          <div style="flex:1;">
            <div style="font-weight:500; ${strikethrough}">${escapeHtml(task.title)}</div>
            ${task.description ? `<small class="muted">${escapeHtml(task.description.substring(0, 80))}${task.description.length > 80 ? '...' : ''}</small><br/>` : ""}
            <div style="margin-top:6px; display:flex; gap:8px; flex-wrap:wrap; align-items:center;">
              <span class="pill ${statusClass}" data-task-status="${task.id}">${task.status.replace("_", " ")}</span>
              <span class="pill ${priorityClass}">${task.priority}</span>
              ${task.project ? `<span class="pill pill-blue">${escapeHtml(task.project)}</span>` : ""}
              ${task.assigned_to ? `<span class="muted" style="font-size:11px;">${escapeHtml(task.assigned_to)}</span>` : ""}
              ${task.due_date ? `<span class="muted ${dueClass}" style="font-size:11px;">${task.due_date}</span>` : ""}
              ${task.comments_count > 0 ? `<span class="muted" style="font-size:11px;">💬 ${task.comments_count}</span>` : ""}
              <button class="btn btn-secondary btn-sm" data-task-edit="${task.id}" style="padding:2px 8px; font-size:11px;" onclick="event.stopPropagation();">Edit</button>
              <button class="btn btn-secondary btn-sm" data-task-delete="${task.id}" style="padding:2px 8px; font-size:11px; color:#dc2626;" onclick="event.stopPropagation();">X</button>
            </div>
          </div>
        </div>
      </div>`;
    }).join("");
    
    // Add click handlers for viewing task details
    listEl.querySelectorAll(".task-item").forEach(item => {
      item.addEventListener("click", async (e) => {
        const id = parseInt(item.dataset.taskId);
        currentTaskId = id;
        await loadTaskDetail(id);
        loadTasks();
      });
    });
    
    // Add event listeners
    listEl.querySelectorAll("[data-task-toggle]").forEach(cb => {
      cb.addEventListener("change", async (e) => {
        const id = e.target.dataset.taskToggle;
        const done = e.target.checked;
        await fetchJson(`${API}/api/tasks/${id}`, {
          method: "PUT",
          body: JSON.stringify({ status: done ? "done" : "open" })
        });
        loadTasks();
        loadTasksStats();
      });
    });
    
    listEl.querySelectorAll("[data-task-status]").forEach(el => {
      el.addEventListener("click", async (e) => {
        e.stopPropagation();
        const id = e.target.dataset.taskStatus;
        const current = e.target.textContent.trim().replace(" ", "_");
        const statuses = ["open", "in_progress", "done"];
        const currentIdx = statuses.indexOf(current);
        const next = statuses[(currentIdx + 1) % statuses.length];
        await fetchJson(`${API}/api/tasks/${id}`, {
          method: "PUT",
          body: JSON.stringify({ status: next })
        });
        loadTasks();
        loadTasksStats();
      });
    });
    
    listEl.querySelectorAll("[data-task-edit]").forEach(btn => {
      btn.addEventListener("click", async (e) => {
        const id = e.target.dataset.taskEdit;
        const res = await fetchJson(`${API}/api/tasks/${id}`);
        if (res.task) {
          qs("taskTitle").value = res.task.title || "";
          qs("taskDescription").value = res.task.description || "";
          qs("taskProject").value = res.task.project || "";
          qs("taskPriority").value = res.task.priority || "medium";
          qs("taskAssigned").value = res.task.assigned_to || "";
          qs("taskDueDate").value = res.task.due_date || "";
          qs("taskTitle").dataset.editId = id;
          qs("createTaskBtn").textContent = "Update Task";
        }
      });
    });
    
    listEl.querySelectorAll("[data-task-delete]").forEach(btn => {
      btn.addEventListener("click", async (e) => {
        if (!confirm("Delete this task?")) return;
        const id = e.target.dataset.taskDelete;
        await fetchJson(`${API}/api/tasks/${id}`, { method: "DELETE" });
        if (currentTaskId === parseInt(id)) {
          currentTaskId = null;
          qs("taskDetailPanel").style.display = "none";
        }
        loadTasks();
        loadTasksStats();
      });
    });
  } catch (err) {
    setTaskApiIndicator("down", "API unavailable");
    listEl.innerHTML = `<div class="item"><small class="muted">Error loading tasks.</small></div>`;
  }
}

async function loadTaskDetail(taskId) {
  const panel = qs("taskDetailPanel");
  const content = qs("taskDetailContent");
  const comments = qs("taskComments");
  
  if (!panel || !content) return;
  
  try {
    const res = await fetchJson(`${API}/api/tasks/${taskId}`);
    if (res.task) {
      const task = res.task;
      const statusClass = task.status === "done" ? "pill-green" : task.status === "in_progress" ? "pill-blue" : "pill-gray";
      const priorityClass = task.priority === "high" ? "pill-red" : task.priority === "medium" ? "pill-orange" : "";
      const isOverdue = task.due_date && task.status !== "done" && new Date(task.due_date) < new Date();
      
      content.innerHTML = `
        <h3 style="margin:0 0 8px 0;">${escapeHtml(task.title)}</h3>
        <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;">
          <span class="pill ${statusClass}">${task.status.replace("_", " ")}</span>
          <span class="pill ${priorityClass}">${task.priority}</span>
          ${task.project ? `<span class="pill pill-blue">${escapeHtml(task.project)}</span>` : ""}
          ${isOverdue ? `<span class="pill pill-red">Overdue</span>` : ""}
        </div>
        ${task.description ? `<p style="margin:0 0 12px 0;color:var(--muted);">${escapeHtml(task.description)}</p>` : ""}
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:13px;">
          ${task.assigned_to ? `<div><span class="muted">Assigned:</span> ${escapeHtml(task.assigned_to)}</div>` : ""}
          ${task.due_date ? `<div><span class="muted">Due:</span> <span class="${isOverdue ? 'text-danger' : ''}">${task.due_date}</span></div>` : ""}
          <div><span class="muted">Created:</span> ${task.created_at ? task.created_at.split("T")[0] : ""}</div>
        </div>
      `;
      
      const commentRows = Array.isArray(res.comments) ? res.comments : [];
      comments.innerHTML = commentRows.length
        ? commentRows.map(c => `
            <div class="comment-item">
              <div class="comment-header">
                <strong>${escapeHtml(c.author || "User")}</strong>
                <span class="muted">${new Date(c.created_at).toLocaleString()}</span>
              </div>
              <p style="margin:4px 0 0 0;font-size:13px;">${escapeHtml(c.comment)}</p>
              <button class="btn btn-link btn-sm" data-delete-comment="${c.id}" style="color:#dc2626;font-size:11px;padding:0;">Delete</button>
            </div>
          `).join("")
        : `<small class="muted">No comments yet.</small>`;
      
      comments.querySelectorAll("[data-delete-comment]").forEach(btn => {
        btn.addEventListener("click", async () => {
          if (!confirm("Delete this comment?")) return;
          await fetchJson(`${API}/api/comments/${btn.dataset.deleteComment}`, { method: "DELETE" });
          loadTaskDetail(taskId);
        });
      });
      
      qs("addCommentBtn").onclick = async () => {
        const text = qs("newComment")?.value?.trim();
        if (!text) return;
        const author = String(qs("taskCommentAuthor")?.value || getSessionUser()).trim() || getSessionUser();
        await fetchJson(`${API}/api/tasks/${taskId}/comments`, {
          method: "POST",
          body: JSON.stringify({ comment: text, author })
        });
        qs("newComment").value = "";
        loadTaskDetail(taskId);
      };
      
      panel.style.display = "block";
    }
  } catch (err) {
    console.error("Failed to load task detail", err);
  }
}

async function loadTasksStats() {
  const statsEl = qs("tasksStats");
  if (!statsEl) return;
  
  try {
    const res = await fetchJson(`${API}/api/tasks/stats/summary`);
    if (res.ok) {
      setTaskApiIndicator("ok", "API online");
      statsEl.innerHTML = `
        <span class="kpi-pill"><strong>Total:</strong> ${res.total}</span>
        <span class="kpi-pill kpi-pill-blue"><strong>Open:</strong> ${res.open}</span>
        <span class="kpi-pill kpi-pill-orange"><strong>In Progress:</strong> ${res.in_progress}</span>
        <span class="kpi-pill kpi-pill-green"><strong>Done:</strong> ${res.done}</span>
        ${res.overdue ? `<span class="kpi-pill kpi-pill-red"><strong>Overdue:</strong> ${res.overdue}</span>` : ""}
      `;
    }
  } catch (err) {
    setTaskApiIndicator("down", "API unavailable");
    console.error("Failed to load tasks stats", err);
  }
}

function initTasks() {
  setTaskApiIndicator("unknown", "API checking...");
  const createBtn = qs("createTaskBtn");
  const loadBtn = qs("loadTasksBtn");
  const myTasksBtn = qs("myTasksBtn");
  const saveViewBtn = qs("saveTaskViewBtn");
  const closeDetailBtn = qs("closeTaskDetail");
  const createProjectBtn = qs("createProjectBtn");
  const statusFilterEl = qs("taskFilterStatus");
  const priorityFilterEl = qs("taskFilterPriority");
  const assignedFilterEl = qs("taskFilterAssigned");

  const refreshFilterInputs = () => {
    currentTaskView = "all";
    currentTaskSidebarActiveKey = "";
    updateSidebarActiveState("tasks");
    loadTasks();
  };

  statusFilterEl?.addEventListener("change", refreshFilterInputs);
  priorityFilterEl?.addEventListener("change", refreshFilterInputs);
  assignedFilterEl?.addEventListener("change", refreshFilterInputs);
  
  createBtn?.addEventListener("click", async () => {
    const title = qs("taskTitle")?.value?.trim();
    if (!title) {
      alert("Task title is required");
      return;
    }
    
    const editId = qs("taskTitle").dataset.editId;
    const data = {
      title,
      description: qs("taskDescription")?.value?.trim() || null,
      project: qs("taskProject")?.value?.trim() || null,
      priority: qs("taskPriority")?.value || "medium",
      assigned_to: qs("taskAssigned")?.value?.trim() || null,
      due_date: qs("taskDueDate").value || null
    };
    
    try {
      const notifyAssignee = qs("taskNotifyAssignee")?.checked !== false;
      const actor = getSessionUser();
      const assignee = String(data.assigned_to || "").trim();
      const watchers = String(qs("taskWatchers")?.value || "")
        .split(",")
        .map((w) => w.trim())
        .filter(Boolean);
      let savedTaskId = null;
      if (editId) {
        const updated = await fetchJson(`${API}/api/tasks/${editId}`, { method: "PUT", body: JSON.stringify(data) });
        savedTaskId = Number(updated?.task?.id || editId || 0);
        delete qs("taskTitle").dataset.editId;
        createBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="16"></line><line x1="8" y1="12" x2="16" y2="12"></line></svg> Create Task`;
      } else {
        const created = await fetchJson(`${API}/api/tasks`, { method: "POST", body: JSON.stringify(data) });
        savedTaskId = Number(created?.task?.id || 0);
      }

      if (notifyAssignee && savedTaskId && assignee) {
        const msg = assignee === actor
          ? `@${assignee} self-assigned this task.`
          : `@${assignee} assigned by @${actor}.`;
        await fetchJson(`${API}/api/tasks/${savedTaskId}/comments`, {
          method: "POST",
          body: JSON.stringify({ comment: msg, author: actor })
        });
      }
      if (notifyAssignee && savedTaskId && watchers.length) {
        const uniqueWatchers = Array.from(new Set(watchers)).filter((w) => w !== assignee);
        if (uniqueWatchers.length) {
          const watchMsg = `Watchers added by @${actor}: ${uniqueWatchers.map((w) => `@${w}`).join(" ")}`;
          await fetchJson(`${API}/api/tasks/${savedTaskId}/comments`, {
            method: "POST",
            body: JSON.stringify({ comment: watchMsg, author: actor })
          });
        }
      }
      
      qs("taskTitle").value = "";
      qs("taskDescription").value = "";
      qs("taskProject").value = "";
      qs("taskPriority").value = "medium";
      qs("taskAssigned").value = "";
      qs("taskWatchers").value = "";
      qs("taskDueDate").value = "";
      if (qs("taskNotifyAssignee")) qs("taskNotifyAssignee").checked = true;
      
      loadTasks();
      loadTasksStats();
      loadProjects();
      setStatus(editId ? "Task updated." : "Task created.");
    } catch (err) {
      alert(`Failed to save task: ${err?.message || err}`);
      setStatus("Task save failed.");
    }
  });
  
  loadBtn?.addEventListener("click", loadTasks);
  
  myTasksBtn?.addEventListener("click", async () => {
    const username = prompt("Enter your username:");
    if (!username) return;
    qs("taskFilterAssigned").value = username;
    currentProjectFilter = "";
    currentTaskView = "all";
    currentTaskSidebarActiveKey = "";
    updateSidebarActiveState("tasks");
    await loadProjects();
    loadTasks();
  });

  saveViewBtn?.addEventListener("click", () => {
    const name = String(prompt("Saved view name:") || "").trim();
    if (!name) return;
    const status = qs("taskFilterStatus")?.value || "";
    const priority = qs("taskFilterPriority")?.value || "";
    const assigned = qs("taskFilterAssigned")?.value || "";
    const existing = getSavedTaskViews();
    const view = {
      id: `${Date.now()}`,
      name,
      view: currentTaskView || "all",
      status,
      priority,
      assigned,
      project: currentProjectFilter || ""
    };
    existing.push(view);
    persistSavedTaskViews(existing.slice(-12));
    renderTaskWorkspaceSavedViews();
  });
  
  closeDetailBtn?.addEventListener("click", () => {
    const content = qs("taskDetailContent");
    if (content) content.innerHTML = `<small class="muted">Select a task from the list to open shared comments and collaboration tools.</small>`;
    const comments = qs("taskComments");
    if (comments) comments.innerHTML = `<small class="muted">No task selected.</small>`;
    currentTaskId = null;
    loadTasks();
  });
  
  createProjectBtn?.addEventListener("click", async () => {
    const name = qs("newProjectName")?.value?.trim();
    if (!name) {
      alert("Project name is required");
      return;
    }
    
    try {
      await fetchJson(`${API}/api/projects`, {
        method: "POST",
        body: JSON.stringify({ name, description: "", color: "#3b82f6" })
      });
      qs("newProjectName").value = "";
      loadProjects();
    } catch (err) {
      alert("Failed to create project");
    }
  });
  
  loadTasks();
  loadTasksStats();
  loadProjects();
  loadTeamMembers();
  renderTaskWorkspaceSavedViews();
}
/* =====================================================================
   FINANCE INTEGRATION (summarized journal posting, period lock,
   budget vs actual, rolling forecast, SSOT report + KPI definitions)
===================================================================== */

let financeLastRunId = null;
let financeLastForecastBatchId = null;
let financeSiteAllocCache = null;

function fmtMoney(n) {
  const v = Number(n || 0);
  return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function currentFinancePeriod() {
  return new Date().toISOString().slice(0, 7);
}

function monthBoundsFromPeriod(period) {
  const p = String(period || "").trim();
  if (!/^\d{4}-\d{2}$/.test(p)) return { start: "", end: "" };
  const [y, m] = p.split("-").map((x) => Number(x));
  const start = `${p}-01`;
  const end = new Date(Date.UTC(y, m, 0)).toISOString().slice(0, 10);
  return { start, end };
}

function initFinanceTab() {
  const period = currentFinancePeriod();
  const { start, end } = monthBoundsFromPeriod(period);
  const setIfEmpty = (id, val) => {
    const el = qs(id);
    if (el && !String(el.value || "").trim()) el.value = val;
  };
  setIfEmpty("finMasterPeriod", period);
  setIfEmpty("finBvaPeriod", period);
  setIfEmpty("finPeriodInput", period);
  setIfEmpty("finSsotPeriod", period);
  setIfEmpty("finForecastStart", period);
  setIfEmpty("finRunStart", start);
  setIfEmpty("finRunEnd", end);
  financeSyncPeriodFields(false);
  loadFinanceSiteFilterFromEntity().catch(() => {});
}

async function loadFinanceSiteFilterFromEntity() {
  const sel = qs("finSiteFilter");
  if (!sel) return;
  try {
    const res = await fetchJson(`${API}/api/entity/sites`);
    const rows = Array.isArray(res.rows) ? res.rows : [];
    if (!rows.length) return;
    const prev = String(sel.value || "");
    sel.innerHTML = `<option value="">All sites</option>${rows
      .map((r) => {
        const code = String(r.site_code || "").trim();
        const name = String(r.site_name || code).trim();
        return code ? `<option value="${escapeHtml(code)}">${escapeHtml(name)} (${escapeHtml(code)})</option>` : "";
      })
      .filter(Boolean)
      .join("")}`;
    if (prev) sel.value = prev;
  } catch {}
}

function financeSyncPeriodFields(showAlert) {
  const period = String(qs("finMasterPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) {
    if (showAlert) alert("Enter reporting period as YYYY-MM.");
    return;
  }
  const { start, end } = monthBoundsFromPeriod(period);
  const ids = ["finBvaPeriod", "finPeriodInput", "finSsotPeriod", "finForecastStart"];
  ids.forEach((id) => {
    const el = qs(id);
    if (el) el.value = period;
  });
  if (qs("finRunStart")) qs("finRunStart").value = start;
  if (qs("finRunEnd")) qs("finRunEnd").value = end;
  if (showAlert) setStatus(`Finance period set to ${period} (${start} → ${end}).`);
}

function finCatAmount(categories, name) {
  const row = (Array.isArray(categories) ? categories : []).find((c) => String(c.category) === name);
  return row ? Number(row.actual || 0) : 0;
}

async function loadFinanceSiteFilterOptions(sites) {
  const sel = qs("finSiteFilter");
  if (!sel) return;
  const prev = String(sel.value || "");
  const codes = Array.from(
    new Set((sites || []).map((s) => String(s.site_code || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));
  sel.innerHTML = `<option value="">All sites</option>${codes
    .map((c) => {
      const name = (sites || []).find((s) => String(s.site_code) === c)?.site_name || c;
      return `<option value="${escapeHtml(c)}">${escapeHtml(name)} (${escapeHtml(c)})</option>`;
    })
    .join("")}`;
  if (prev && codes.includes(prev)) sel.value = prev;
}

async function loadFinanceSiteAllocation() {
  const period = String(qs("finMasterPeriod")?.value || "").trim();
  const site = String(qs("finSiteFilter")?.value || "").trim();
  const msg = qs("finSiteAllocMsg");
  const body = qs("finSiteAllocBody");
  const totalsEl = qs("finSiteAllocTotals");
  const warnEl = qs("finSiteAllocWarn");
  if (!/^\d{4}-\d{2}$/.test(period)) {
    alert("Enter reporting period as YYYY-MM.");
    return;
  }
  if (msg) msg.textContent = "Loading site allocation…";
  if (body) body.innerHTML = `<tr><td colspan="11" class="muted">Loading…</td></tr>`;
  const q = new URLSearchParams({ period });
  if (site) q.set("site_code", site);
  try {
    const res = await fetchJson(`${API}/api/finance/site-allocation?${q.toString()}`);
    financeSiteAllocCache = res;
    const sites = Array.isArray(res.sites) ? res.sites : [];
    await loadFinanceSiteFilterOptions(sites);
    const t = res.totals || {};
    if (totalsEl) {
      totalsEl.innerHTML = `
        <div class="kpi-card kpi-util">
          <div class="kpi-card-header"><div class="kpi-icon">B</div><div class="kpi-title">Budget</div></div>
          <div class="kpi-big-value">${fmtMoney(t.budget)}</div>
          <div class="kpi-meta">Period ${escapeHtml(period)}</div>
        </div>
        <div class="kpi-card kpi-scheduled">
          <div class="kpi-card-header"><div class="kpi-icon">A</div><div class="kpi-title">Actual</div></div>
          <div class="kpi-big-value">${fmtMoney(t.actual)}</div>
          <div class="kpi-meta">${escapeHtml(res.period_start || "")} → ${escapeHtml(res.period_end || "")}</div>
        </div>
        <div class="kpi-card kpi-alerts">
          <div class="kpi-card-header"><div class="kpi-icon">V</div><div class="kpi-title">Variance</div></div>
          <div class="kpi-big-value">${fmtMoney(t.variance)}</div>
          <div class="kpi-meta">Actual minus budget</div>
        </div>
      `;
    }
    if (warnEl) {
      const miss = Number(res.assets_missing_site || 0);
      warnEl.innerHTML = miss > 0
        ? `<div class="message-error" style="padding:8px 10px; border-radius:8px;">${miss} active asset(s) have no <code>site_code</code> — their costs appear under <strong>Unassigned</strong>. Update assets to improve site splits.</div>`
        : (res.has_asset_site_column === false
          ? `<div class="message-error" style="padding:8px 10px; border-radius:8px;">Asset <code>site_code</code> column missing — restart API to apply schema update.</div>`
          : "");
    }
    if (body) {
      body.innerHTML = sites.length
        ? sites.map((s) => `<tr>
            <td>${escapeHtml(s.site_name || s.site_code || "-")}</td>
            <td><code>${escapeHtml(s.site_code || "")}</code></td>
            <td class="num">${fmtMoney(s.budget_total)}</td>
            <td class="num">${fmtMoney(s.total_actual)}</td>
            <td class="num">${fmtMoney(s.variance)}</td>
            <td class="num">${s.variance_pct == null ? "-" : Number(s.variance_pct).toFixed(1) + "%"}</td>
            <td class="num">${fmtMoney(finCatAmount(s.categories, "parts"))}</td>
            <td class="num">${fmtMoney(finCatAmount(s.categories, "labor"))}</td>
            <td class="num">${fmtMoney(finCatAmount(s.categories, "fuel"))}</td>
            <td class="num">${fmtMoney(finCatAmount(s.categories, "lube"))}</td>
            <td class="num">${fmtMoney(finCatAmount(s.categories, "downtime"))}</td>
          </tr>`).join("")
        : `<tr><td colspan="11" class="muted">No costs in this period${site ? ` for site ${escapeHtml(site)}` : ""}.</td></tr>`;
    }
    if (msg) {
      msg.textContent = `Loaded ${sites.length} site row(s) for ${period}.`;
      msg.className = "message-success";
    }
    financeSyncPeriodFields(false);
  } catch (e) {
    financeSiteAllocCache = null;
    if (msg) {
      msg.textContent = `Load failed: ${e.message}`;
      msg.className = "message-error";
    }
    if (body) body.innerHTML = `<tr><td colspan="11" class="message-error">${escapeHtml(e.message)}</td></tr>`;
  }
}

function financeExportSiteAllocation() {
  const period = String(qs("finMasterPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) {
    alert("Enter reporting period as YYYY-MM.");
    return;
  }
  const site = String(qs("finSiteFilter")?.value || "").trim();
  const q = new URLSearchParams({ period });
  if (site) q.set("site_code", site);
  window.open(`${API}/api/finance/site-allocation/export.xlsx?${q.toString()}`, "_blank");
}

async function financeBuildSummarizedRun() {
  const start = String(qs("finRunStart")?.value || "").trim();
  const end = String(qs("finRunEnd")?.value || "").trim();
  if (!start || !end) { alert("Select start and end dates."); return; }
  const categories = [];
  if (qs("finCatParts")?.checked) categories.push("parts");
  if (qs("finCatLabor")?.checked) categories.push("labor");
  if (qs("finCatDowntime")?.checked) categories.push("downtime");
  if (qs("finCatFuel")?.checked) categories.push("fuel");
  if (qs("finCatLube")?.checked) categories.push("lube");
  if (qs("finCatGrn")?.checked) categories.push("procurement_grn");
  if (qs("finCatAp")?.checked) categories.push("procurement_ap");
  if (!categories.length) { alert("Select at least one category."); return; }
  setStatus("Building summarized journal run...");
  try {
    const res = await fetchJson(`${API}/api/procurement/journals/summarize`, {
      method: "POST",
      body: JSON.stringify({
        start,
        end,
        categories,
        default_cost_center_code: String(qs("finRunDefaultCC")?.value || "").trim() || undefined,
        currency: String(qs("finRunCurrency")?.value || "USD").trim() || "USD",
      }),
    });
    const el = qs("finRunMsg");
    if (el) el.textContent = `Run ${res.run?.run_number} built | lines=${res.run?.line_count} | debit=${fmtMoney(res.run?.total_debit)} | credit=${fmtMoney(res.run?.total_credit)} | balanced=${res.balanced}`;
    financeLastRunId = Number(res.run?.id || 0) || null;
    const rid = qs("finActiveRunId");
    if (rid && financeLastRunId) rid.value = String(financeLastRunId);
    await loadFinanceRuns();
    setStatus("Summarized journal run built.");
  } catch (e) {
    setStatus(`Build failed: ${e.message}`);
    alert(`Build failed: ${e.message}`);
  }
}

async function loadFinanceRuns() {
  const list = qs("finRunList");
  if (!list) return;
  const status = String(qs("finRunStatusFilter")?.value || "").trim();
  const q = status ? `?status=${encodeURIComponent(status)}` : "";
  try {
    const data = await fetchJson(`${API}/api/procurement/journals/runs${q}`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    if (!rows.length) {
      list.innerHTML = "<em>No runs found.</em>";
      return;
    }
    list.innerHTML = `
      <table class="table">
        <thead><tr>
          <th>ID</th><th>Run #</th><th>Period</th><th>Range</th><th>Status</th>
          <th class="num">Debit</th><th class="num">Credit</th><th class="num">Lines</th><th>Created</th>
        </tr></thead>
        <tbody>
          ${rows.map((r) => `
            <tr>
              <td><a href="#" data-fin-run-id="${r.id}">${r.id}</a></td>
              <td>${escapeHtml(r.run_number)}</td>
              <td>${escapeHtml(r.period || "")}</td>
              <td>${escapeHtml(r.start_date || "")} → ${escapeHtml(r.end_date || "")}</td>
              <td>${escapeHtml(r.status || "")}</td>
              <td class="num">${fmtMoney(r.total_debit)}</td>
              <td class="num">${fmtMoney(r.total_credit)}</td>
              <td class="num">${Number(r.line_count || 0)}</td>
              <td>${escapeHtml((r.created_at || "").slice(0, 16))}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
    list.querySelectorAll("[data-fin-run-id]").forEach((a) => {
      a.addEventListener("click", (ev) => {
        ev.preventDefault();
        const id = Number(a.getAttribute("data-fin-run-id")) || 0;
        if (!id) return;
        financeLastRunId = id;
        const el = qs("finActiveRunId");
        if (el) el.value = String(id);
        loadFinanceRunDetail().catch(() => {});
      });
    });
  } catch (e) {
    list.innerHTML = `<em class="muted">Failed: ${escapeHtml(e.message)}</em>`;
  }
}

async function loadFinanceRunDetail() {
  const id = Number(qs("finActiveRunId")?.value || 0);
  if (!id) { alert("Enter a run ID first."); return; }
  const el = qs("finRunDetail");
  if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/procurement/journals/runs/${id}`);
    const run = res.run || {};
    const byCat = Array.isArray(res.by_category) ? res.by_category : [];
    const bySite = Array.isArray(res.by_site) ? res.by_site : [];
    el.innerHTML = `
      <div class="kpi-pills">
        <span class="kpi-pill"><strong>Run:</strong> ${escapeHtml(run.run_number || "")}</span>
        <span class="kpi-pill"><strong>Status:</strong> ${escapeHtml(run.status || "")}</span>
        <span class="kpi-pill"><strong>Debit:</strong> ${fmtMoney(run.total_debit)}</span>
        <span class="kpi-pill"><strong>Credit:</strong> ${fmtMoney(run.total_credit)}</span>
        <span class="kpi-pill"><strong>Lines:</strong> ${Number(run.line_count || 0)}</span>
      </div>
      <h5 style="margin-top:12px;">By site</h5>
      <table class="table">
        <thead><tr><th>Site</th><th class="num">Lines</th><th class="num">Debit</th><th class="num">Credit</th></tr></thead>
        <tbody>
          ${bySite.length ? bySite.map((s) => `<tr>
            <td><code>${escapeHtml(s.site_code || "")}</code></td>
            <td class="num">${Number(s.lines || 0)}</td>
            <td class="num">${fmtMoney(s.debit_total)}</td>
            <td class="num">${fmtMoney(s.credit_total)}</td>
          </tr>`).join("") : `<tr><td colspan="4" class="muted">No site breakdown on lines.</td></tr>`}
        </tbody>
      </table>
      <h5 style="margin-top:12px;">By category</h5>
      <table class="table">
        <thead><tr><th>Category</th><th class="num">Lines</th><th class="num">Debit</th><th class="num">Credit</th></tr></thead>
        <tbody>
          ${byCat.map((c) => `<tr>
            <td>${escapeHtml(c.category)}</td>
            <td class="num">${Number(c.lines || 0)}</td>
            <td class="num">${fmtMoney(c.debit_total)}</td>
            <td class="num">${fmtMoney(c.credit_total)}</td>
          </tr>`).join("")}
        </tbody>
      </table>
    `;
  } catch (e) {
    el.innerHTML = `<em class="muted">Load failed: ${escapeHtml(e.message)}</em>`;
  }
}

function financeExportRunCsv() {
  const id = Number(qs("finActiveRunId")?.value || 0);
  if (!id) { alert("Enter a run ID first."); return; }
  window.open(`${API}/api/procurement/journals/runs/${id}/export.csv`, "_blank");
}
function financeExportRunXlsx() {
  const id = Number(qs("finActiveRunId")?.value || 0);
  if (!id) { alert("Enter a run ID first."); return; }
  window.open(`${API}/api/procurement/journals/runs/${id}/export.xlsx`, "_blank");
}

async function financeMarkExported() {
  const id = Number(qs("finActiveRunId")?.value || 0);
  if (!id) { alert("Enter a run ID first."); return; }
  try {
    await fetchJson(`${API}/api/procurement/journals/runs/${id}/mark-exported`, { method: "POST", body: JSON.stringify({}) });
    setStatus("Run marked exported.");
    await loadFinanceRuns();
    await loadFinanceRunDetail().catch(() => {});
  } catch (e) { alert(e.message); }
}
async function financeMarkPosted() {
  const id = Number(qs("finActiveRunId")?.value || 0);
  if (!id) { alert("Enter a run ID first."); return; }
  const ref = String(qs("finPostedRef")?.value || "").trim();
  if (!confirm(`Mark run ${id} as posted? This locks it from further changes.`)) return;
  try {
    await fetchJson(`${API}/api/procurement/journals/runs/${id}/mark-posted`, {
      method: "POST",
      body: JSON.stringify({ posted_reference: ref || null }),
    });
    setStatus("Run marked posted.");
    await loadFinanceRuns();
    await loadFinanceRunDetail().catch(() => {});
  } catch (e) { alert(e.message); }
}
async function financeReverseRun() {
  const id = Number(qs("finActiveRunId")?.value || 0);
  if (!id) { alert("Enter a run ID first."); return; }
  const reason = String(qs("finReverseReason")?.value || "").trim();
  if (!reason) { alert("Reason required."); return; }
  if (!confirm(`Reverse run ${id}? Reason: ${reason}`)) return;
  try {
    await fetchJson(`${API}/api/procurement/journals/runs/${id}/reverse`, {
      method: "POST",
      body: JSON.stringify({ reason }),
    });
    setStatus("Run reversed.");
    await loadFinanceRuns();
    await loadFinanceRunDetail().catch(() => {});
  } catch (e) { alert(e.message); }
}

async function loadFinanceChecklist() {
  const period = String(qs("finPeriodInput")?.value || "").trim();
  const msg = qs("finChecklistMsg");
  const tbl = qs("finChecklistTable");
  if (!period || !/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  try {
    const res = await fetchJson(`${API}/api/finance/periods/${period}/checklist`);
    const items = Array.isArray(res.items) ? res.items : [];
    const lock = res.lock || {};
    if (msg) msg.textContent = `Period ${period} | lock status: ${lock.status || "open"} | ${items.length} items`;
    if (tbl) {
      tbl.innerHTML = `
        <table class="table">
          <thead><tr><th>Code</th><th>Task</th><th>Status</th><th>Updated</th><th>Actions</th></tr></thead>
          <tbody>
            ${items.map((it) => `
              <tr>
                <td><code>${escapeHtml(it.code)}</code></td>
                <td>${escapeHtml(it.label || "")}</td>
                <td>${escapeHtml(it.status || "")}</td>
                <td>${escapeHtml((it.updated_at || "").slice(0, 16))}</td>
                <td>
                  <button type="button" data-chk-mark-done="${escapeHtml(it.code)}">Done</button>
                  <button type="button" data-chk-mark-skip="${escapeHtml(it.code)}">Skip</button>
                  <button type="button" data-chk-mark-pending="${escapeHtml(it.code)}">Reset</button>
                </td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      `;
      tbl.querySelectorAll("[data-chk-mark-done]").forEach((b) => b.addEventListener("click", () => financeChecklistUpdate(period, b.getAttribute("data-chk-mark-done"), "done")));
      tbl.querySelectorAll("[data-chk-mark-skip]").forEach((b) => b.addEventListener("click", () => financeChecklistUpdate(period, b.getAttribute("data-chk-mark-skip"), "skipped")));
      tbl.querySelectorAll("[data-chk-mark-pending]").forEach((b) => b.addEventListener("click", () => financeChecklistUpdate(period, b.getAttribute("data-chk-mark-pending"), "pending")));
    }
  } catch (e) {
    if (msg) msg.textContent = `Load failed: ${e.message}`;
  }
}

async function financeChecklistUpdate(period, code, status) {
  try {
    await fetchJson(`${API}/api/finance/periods/${period}/checklist/${code}`, {
      method: "POST",
      body: JSON.stringify({ status }),
    });
    await loadFinanceChecklist();
  } catch (e) { alert(e.message); }
}

async function financeClosePeriod() {
  const period = String(qs("finPeriodInput")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  const force = qs("finClosePeriodForce")?.checked || false;
  if (!confirm(`Close period ${period}? force=${force}`)) return;
  try {
    const res = await fetchJson(`${API}/api/finance/periods/${period}/close`, {
      method: "POST",
      body: JSON.stringify({ force }),
    });
    setStatus(`Period ${period} closed.`);
    await loadFinanceChecklist();
  } catch (e) { alert(e.message); }
}
async function financeReopenPeriod() {
  const period = String(qs("finPeriodInput")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  const reason = String(qs("finReopenReason")?.value || "").trim();
  if (!reason) { alert("Reason required to reopen"); return; }
  if (!confirm(`Reopen period ${period}? Reason: ${reason}`)) return;
  try {
    await fetchJson(`${API}/api/finance/periods/${period}/reopen`, {
      method: "POST",
      body: JSON.stringify({ reason }),
    });
    setStatus(`Period ${period} reopened.`);
    await loadFinanceChecklist();
  } catch (e) { alert(e.message); }
}

async function loadFinanceBudgetVsActual() {
  const period = String(qs("finBvaPeriod")?.value || "").trim();
  const dim = String(qs("finBvaDimension")?.value || "site_code").trim();
  const site = String(qs("finSiteFilter")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  try {
    const q = new URLSearchParams({ period, dimension: dim });
    if (site) q.set("site_code", site);
    const res = await fetchJson(`${API}/api/finance/budgets-vs-actual?${q.toString()}`);
    const rows = Array.isArray(res.rows) ? res.rows : [];
    const total = res.total || {};
    const totalsEl = qs("finBvaTotals");
    if (totalsEl) {
      totalsEl.innerHTML = `
        <div class="kpi-pills">
          <span class="kpi-pill"><strong>Budget:</strong> ${fmtMoney(total.budget)}</span>
          <span class="kpi-pill"><strong>Actual:</strong> ${fmtMoney(total.actual)}</span>
          <span class="kpi-pill"><strong>Variance:</strong> ${fmtMoney(total.variance)}</span>
        </div>
      `;
    }
    const dimLabel = dim === "site_code" ? "Site" : dim === "cost_center_code" ? "Cost Center" : dim;
    const tbl = qs("finBvaTable");
    if (tbl) {
      tbl.innerHTML = `
        <table class="table">
          <thead><tr><th>${escapeHtml(dimLabel)}</th><th class="num">Budget</th><th class="num">Actual</th><th class="num">Variance</th><th class="num">Variance %</th></tr></thead>
          <tbody>
            ${rows.map((r) => `<tr>
              <td>${escapeHtml(r.dimension_key || "(none)")}</td>
              <td class="num">${fmtMoney(r.budget)}</td>
              <td class="num">${fmtMoney(r.actual)}</td>
              <td class="num">${fmtMoney(r.variance)}</td>
              <td class="num">${r.variance_pct == null ? "-" : Number(r.variance_pct).toFixed(1) + "%"}</td>
            </tr>`).join("")}
          </tbody>
        </table>
      `;
    }
  } catch (e) { alert(e.message); }
}

async function financeSaveBudgets() {
  const raw = String(qs("finBudgetJson")?.value || "").trim();
  if (!raw) { alert("Paste budget rows JSON first"); return; }
  let rows;
  try { rows = JSON.parse(raw); } catch { alert("Invalid JSON"); return; }
  if (!Array.isArray(rows)) { alert("JSON must be an array"); return; }
  try {
    const res = await fetchJson(`${API}/api/finance/budgets/upsert`, {
      method: "POST",
      body: JSON.stringify({ rows }),
    });
    setStatus(`Saved ${res.saved || 0} budget rows.`);
    await loadFinanceBudgetList();
  } catch (e) { alert(e.message); }
}

async function loadFinanceBudgetList() {
  const period = String(qs("finBvaPeriod")?.value || "").trim();
  const q = period ? `?period=${encodeURIComponent(period)}` : "";
  try {
    const res = await fetchJson(`${API}/api/finance/budgets${q}`);
    const rows = Array.isArray(res.rows) ? res.rows : [];
    const el = qs("finBudgetList");
    if (!el) return;
    if (!rows.length) {
      el.innerHTML = "<em>No budgets.</em>";
      return;
    }
    el.innerHTML = `
      <table class="table">
        <thead><tr><th>Period</th><th>Site</th><th>Cost Center</th><th>Equipment</th><th>Category</th><th class="num">Budget</th><th>Currency</th></tr></thead>
        <tbody>
          ${rows.map((r) => `<tr>
            <td>${escapeHtml(r.period)}</td>
            <td>${escapeHtml(r.site_code || "")}</td>
            <td>${escapeHtml(r.cost_center_code || "")}</td>
            <td>${escapeHtml(r.equipment_type || "")}</td>
            <td>${escapeHtml(r.category)}</td>
            <td class="num">${fmtMoney(r.budget_amount)}</td>
            <td>${escapeHtml(r.currency || "USD")}</td>
          </tr>`).join("")}
        </tbody>
      </table>
    `;
  } catch (e) {
    const el = qs("finBudgetList");
    if (el) el.innerHTML = `<em class="muted">Failed: ${escapeHtml(e.message)}</em>`;
  }
}

async function financeRebuildForecast() {
  const start = String(qs("finForecastStart")?.value || "").trim();
  const months = Math.max(1, Math.min(6, Number(qs("finForecastMonths")?.value || 3)));
  if (!/^\d{4}-\d{2}$/.test(start)) { alert("Enter start period as YYYY-MM"); return; }
  try {
    const res = await fetchJson(`${API}/api/finance/forecast/rebuild`, {
      method: "POST",
      body: JSON.stringify({ start_period: start, months }),
    });
    financeLastForecastBatchId = res.batch_id || null;
    setStatus(`Forecast rebuilt | batch ${res.batch_id} | ${res.saved || 0} rows.`);
    await loadFinanceForecast();
  } catch (e) { alert(e.message); }
}

async function loadFinanceForecast() {
  const msg = qs("finForecastMsg");
  const totalsEl = qs("finForecastTotals");
  const tbl = qs("finForecastTable");
  try {
    const q = financeLastForecastBatchId ? `?batch_id=${encodeURIComponent(financeLastForecastBatchId)}` : "";
    const res = await fetchJson(`${API}/api/finance/forecast${q}`);
    const rows = Array.isArray(res.rows) ? res.rows : [];
    const totals = Array.isArray(res.totals_by_period) ? res.totals_by_period : [];
    if (msg) msg.textContent = `${rows.length} forecast rows across ${totals.length} months`;
    if (totalsEl) {
      totalsEl.innerHTML = totals.map((t) => `
        <span class="kpi-pill"><strong>${escapeHtml(t.period)}:</strong> ${fmtMoney(t.forecast)} (baseline ${fmtMoney(t.baseline)} + uplift ${fmtMoney(t.uplift)})</span>
      `).join(" ");
    }
    if (tbl) {
      tbl.innerHTML = `
        <table class="table">
          <thead><tr><th>Period</th><th>Site</th><th>Cost Center</th><th>Equipment</th><th>Category</th>
            <th class="num">Baseline</th><th class="num">Uplift</th><th class="num">Forecast</th></tr></thead>
          <tbody>
            ${rows.map((r) => `<tr>
              <td>${escapeHtml(r.period)}</td>
              <td>${escapeHtml(r.site_code || "")}</td>
              <td>${escapeHtml(r.cost_center_code || "")}</td>
              <td>${escapeHtml(r.equipment_type || "")}</td>
              <td>${escapeHtml(r.category)}</td>
              <td class="num">${fmtMoney(r.baseline_amount)}</td>
              <td class="num">${fmtMoney(r.uplift_amount)}</td>
              <td class="num">${fmtMoney(r.forecast_amount)}</td>
            </tr>`).join("")}
          </tbody>
        </table>
      `;
    }
  } catch (e) {
    if (msg) msg.textContent = `Load failed: ${e.message}`;
  }
}

async function loadFinanceSsot() {
  const period = String(qs("finSsotPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  try {
    const res = await fetchJson(`${API}/api/finance/reports/ssot?period=${encodeURIComponent(period)}`);
    const kpi = res.kpi || {};
    const kel = qs("finSsotKpi");
    if (kel) {
      kel.innerHTML = `
        <div class="kpi-pills">
          <span class="kpi-pill"><strong>Availability:</strong> ${kpi.availability == null ? "-" : kpi.availability + "%"}</span>
          <span class="kpi-pill"><strong>Utilization:</strong> ${kpi.utilization == null ? "-" : kpi.utilization + "%"}</span>
          <span class="kpi-pill"><strong>MTBF:</strong> ${kpi.mtbf == null ? "-" : kpi.mtbf + " h"}</span>
          <span class="kpi-pill"><strong>MTTR:</strong> ${kpi.mttr == null ? "-" : kpi.mttr + " h"}</span>
          <span class="kpi-pill"><strong>Cost/Asset-hr:</strong> ${kpi.cost_per_asset_hour == null ? "-" : fmtMoney(kpi.cost_per_asset_hour)}</span>
          <span class="kpi-pill"><strong>Run hrs:</strong> ${kpi.run_hours}</span>
          <span class="kpi-pill"><strong>Down hrs:</strong> ${kpi.downtime_hours}</span>
          <span class="kpi-pill"><strong>Total Cost:</strong> ${fmtMoney(kpi.total_cost)}</span>
        </div>
      `;
    }
    const actuals = Array.isArray(res.actuals) ? res.actuals : [];
    const ael = qs("finSsotActuals");
    if (ael) {
      ael.innerHTML = `
        <h5>Actuals Breakdown</h5>
        <table class="table">
          <thead><tr><th>Site</th><th>Cost Center</th><th>Equipment</th><th>Category</th><th class="num">Actual</th></tr></thead>
          <tbody>
            ${actuals.map((r) => `<tr>
              <td>${escapeHtml(r.site_code || "")}</td>
              <td>${escapeHtml(r.cost_center_code || "")}</td>
              <td>${escapeHtml(r.equipment_type || "")}</td>
              <td>${escapeHtml(r.category)}</td>
              <td class="num">${fmtMoney(r.actual_amount)}</td>
            </tr>`).join("")}
          </tbody>
        </table>
      `;
    }
    const defs = Array.isArray(res.kpi_definitions) ? res.kpi_definitions : [];
    const dEl = qs("finKpiDefs");
    if (dEl) {
      dEl.innerHTML = `
        <table class="table">
          <thead><tr><th>Code</th><th>Label</th><th>Unit</th><th>Formula</th><th>Source Tables</th></tr></thead>
          <tbody>
            ${defs.map((d) => `<tr>
              <td><code>${escapeHtml(d.code)}</code></td>
              <td>${escapeHtml(d.label)}</td>
              <td>${escapeHtml(d.unit)}</td>
              <td><code>${escapeHtml(d.formula || "")}</code></td>
              <td>${escapeHtml((d.source_tables || []).join(", "))}</td>
            </tr>`).join("")}
          </tbody>
        </table>
      `;
    }
  } catch (e) { alert(e.message); }
}
function financeExportSsotCsv() {
  const period = String(qs("finSsotPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  window.open(`${API}/api/finance/reports/ssot/export.csv?period=${encodeURIComponent(period)}`, "_blank");
}
function financeExportSsotXlsx() {
  const period = String(qs("finSsotPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  window.open(`${API}/api/finance/reports/ssot/export.xlsx?period=${encodeURIComponent(period)}`, "_blank");
}

async function loadFinanceKpiDefs() {
  const el = qs("finKpiDefs");
  if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/finance/kpis/definitions`);
    const defs = Array.isArray(res.kpis) ? res.kpis : [];
    el.innerHTML = `
      <table class="table">
        <thead><tr><th>Code</th><th>Label</th><th>Unit</th><th>Formula</th><th>Source Tables</th><th>Exclusions</th></tr></thead>
        <tbody>
          ${defs.map((d) => `<tr>
            <td><code>${escapeHtml(d.code)}</code></td>
            <td>${escapeHtml(d.label)}</td>
            <td>${escapeHtml(d.unit)}</td>
            <td><code>${escapeHtml(d.formula || "")}</code></td>
            <td>${escapeHtml((d.source_tables || []).join(", "))}</td>
            <td>${escapeHtml((d.exclusions || []).join(", "))}</td>
          </tr>`).join("")}
        </tbody>
      </table>
    `;
  } catch (e) { el.innerHTML = `<em class="muted">Failed: ${escapeHtml(e.message)}</em>`; }
}

function bindFinanceHandlers() {
  const bind = (id, ev, fn) => { const el = qs(id); if (el) el.addEventListener(ev, fn); };
  bind("finLoadSiteAllocBtn", "click", () => loadFinanceSiteAllocation().catch((e) => alert(e.message)));
  bind("finExportSiteAllocBtn", "click", financeExportSiteAllocation);
  bind("finSyncPeriodsBtn", "click", () => financeSyncPeriodFields(true));
  bind("finBuildRunBtn", "click", financeBuildSummarizedRun);
  bind("finLoadRunsBtn", "click", () => loadFinanceRuns().catch(() => {}));
  bind("finLoadRunDetailBtn", "click", () => loadFinanceRunDetail().catch(() => {}));
  bind("finExportRunCsvBtn", "click", financeExportRunCsv);
  bind("finExportRunXlsxBtn", "click", financeExportRunXlsx);
  bind("finMarkExportedBtn", "click", financeMarkExported);
  bind("finMarkPostedBtn", "click", financeMarkPosted);
  bind("finReverseRunBtn", "click", financeReverseRun);
  bind("finLoadChecklistBtn", "click", loadFinanceChecklist);
  bind("finClosePeriodBtn", "click", financeClosePeriod);
  bind("finReopenPeriodBtn", "click", financeReopenPeriod);
  bind("finLoadBvaBtn", "click", loadFinanceBudgetVsActual);
  bind("finSaveBudgetBtn", "click", financeSaveBudgets);
  bind("finLoadBudgetListBtn", "click", loadFinanceBudgetList);
  bind("finRebuildForecastBtn", "click", financeRebuildForecast);
  bind("finLoadForecastBtn", "click", loadFinanceForecast);
  bind("finLoadSsotBtn", "click", loadFinanceSsot);
  bind("finExportSsotCsvBtn", "click", financeExportSsotCsv);
  bind("finExportSsotXlsxBtn", "click", financeExportSsotXlsx);
  bind("finLoadKpiDefsBtn", "click", loadFinanceKpiDefs);
  const runStatusSel = qs("finRunStatusFilter");
  if (runStatusSel) runStatusSel.addEventListener("change", () => loadFinanceRuns().catch(() => {}));
}

/* =====================================================================
   ENTERPRISE (entity/integrations/governance/evidence)
===================================================================== */

async function entSaveCompany() {
  const payload = {
    company_code: qs("entCompanyCode")?.value,
    company_name: qs("entCompanyName")?.value,
    base_currency: qs("entCompanyBaseCcy")?.value,
    reporting_currency: qs("entCompanyReportCcy")?.value,
    tax_region: qs("entCompanyTaxRegion")?.value,
  };
  if (!payload.company_code || !payload.company_name) { alert("company code + name required"); return; }
  try { await fetchJson(`${API}/api/entity/companies`, { method: "POST", body: JSON.stringify(payload) }); setStatus("Company saved."); entLoadCompanies(); }
  catch (e) { alert(e.message); }
}

async function entLoadCompanies() {
  const el = qs("entCompanyList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/entity/companies`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No companies.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>Code</th><th>Name</th><th>Base</th><th>Reporting</th><th>Active</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${escapeHtml(r.company_code)}</td><td>${escapeHtml(r.company_name)}</td><td>${escapeHtml(r.base_currency || "")}</td><td>${escapeHtml(r.reporting_currency || "")}</td><td>${r.active ? "yes" : "no"}</td></tr>`).join("")}</tbody></table>`;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function entSaveSite() {
  const payload = {
    site_code: qs("entSiteCode")?.value,
    company_code: qs("entSiteCompany")?.value,
    site_name: qs("entSiteName")?.value,
    local_currency: qs("entSiteLocalCcy")?.value,
    region: qs("entSiteRegion")?.value,
  };
  if (!payload.site_code || !payload.company_code || !payload.site_name) { alert("site_code, company_code, site_name required"); return; }
  try { await fetchJson(`${API}/api/entity/sites`, { method: "POST", body: JSON.stringify(payload) }); setStatus("Site saved."); entLoadSites(); }
  catch (e) { alert(e.message); }
}

async function entLoadSites() {
  const el = qs("entSiteList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/entity/sites`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No sites.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>Site</th><th>Company</th><th>Name</th><th>Local Ccy</th><th>Region</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${escapeHtml(r.site_code)}</td><td>${escapeHtml(r.company_code)}</td><td>${escapeHtml(r.site_name)}</td><td>${escapeHtml(r.local_currency || "")}</td><td>${escapeHtml(r.region || "")}</td></tr>`).join("")}</tbody></table>`;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function entLoadTree() {
  const el = qs("entTree"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/entity/tree`);
    const tree = res.tree || [];
    if (!tree.length) { el.innerHTML = "<em>No entities.</em>"; return; }
    el.innerHTML = tree.map((c) => `<div><strong>${escapeHtml(c.company_code)}</strong> - ${escapeHtml(c.company_name)}<ul>${(c.sites || []).map((s) => `<li>${escapeHtml(s.site_code)} - ${escapeHtml(s.site_name)}</li>`).join("")}</ul></div>`).join("");
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function entSaveCcyRates() {
  const raw = String(qs("entCcyJson")?.value || "").trim();
  if (!raw) { alert("paste rows JSON"); return; }
  let rows; try { rows = JSON.parse(raw); } catch { alert("invalid JSON"); return; }
  try { const res = await fetchJson(`${API}/api/entity/currency/rates/upsert`, { method: "POST", body: JSON.stringify({ rows }) }); setStatus(`Saved ${res.saved || 0} rates.`); entLoadCcyRates(); }
  catch (e) { alert(e.message); }
}

async function entLoadCcyRates() {
  const el = qs("entCcyList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/entity/currency/rates`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No rates.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>From</th><th>To</th><th>Rate</th><th>Effective</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${escapeHtml(r.from_currency)}</td><td>${escapeHtml(r.to_currency)}</td><td>${Number(r.rate).toFixed(4)}</td><td>${escapeHtml(r.effective_date)}</td></tr>`).join("")}</tbody></table>`;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function entSaveTax() {
  const payload = {
    tax_code: qs("entTaxCode")?.value,
    label: qs("entTaxLabel")?.value,
    rate_pct: Number(qs("entTaxRate")?.value || 0),
    region: qs("entTaxRegion")?.value,
  };
  if (!payload.tax_code || !payload.label) { alert("tax_code and label required"); return; }
  try { await fetchJson(`${API}/api/entity/tax/profiles/upsert`, { method: "POST", body: JSON.stringify(payload) }); setStatus("Tax profile saved."); entLoadTax(); }
  catch (e) { alert(e.message); }
}

async function entLoadTax() {
  const el = qs("entTaxList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/entity/tax/profiles`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No tax profiles.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>Code</th><th>Label</th><th>Rate %</th><th>Region</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${escapeHtml(r.tax_code)}</td><td>${escapeHtml(r.label)}</td><td>${Number(r.rate_pct).toFixed(2)}</td><td>${escapeHtml(r.region || "")}</td></tr>`).join("")}</tbody></table>`;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function intSaveConnection() {
  const payload = {
    connection_code: qs("intConnCode")?.value,
    connector_key: qs("intConnKey")?.value,
    label: qs("intConnLabel")?.value,
  };
  if (!payload.connection_code || !payload.label) { alert("connection_code + label required"); return; }
  try { await fetchJson(`${API}/api/integrations/connections/upsert`, { method: "POST", body: JSON.stringify(payload) }); setStatus("Connection saved."); intLoadConnections(); }
  catch (e) { alert(e.message); }
}

async function intLoadConnections() {
  const el = qs("intConnList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/integrations/connections`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No connections.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>Code</th><th>Connector</th><th>Label</th><th>Active</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${escapeHtml(r.connection_code)}</td><td>${escapeHtml(r.connector_key)}</td><td>${escapeHtml(r.label)}</td><td>${r.active ? "yes" : "no"}</td></tr>`).join("")}</tbody></table>`;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function intEnqueueJob() {
  const connection_code = qs("intJobConnCode")?.value;
  const connector_key = qs("intJobConnector")?.value;
  const rawPayload = String(qs("intJobPayload")?.value || "").trim();
  const idempotency_key = String(qs("intJobIdem")?.value || "").trim() || null;
  if (!connection_code) { alert("connection_code required"); return; }
  let payload = {};
  if (rawPayload) { try { payload = JSON.parse(rawPayload); } catch { alert("invalid payload JSON"); return; } }
  try {
    const res = await fetchJson(`${API}/api/integrations/jobs/enqueue`, {
      method: "POST",
      body: JSON.stringify({ connection_code, connector_key, payload, idempotency_key }),
    });
    setStatus(`Job ${res.id} ${res.status}${res.duplicate ? " (duplicate)" : ""}`);
    const el = qs("intActiveJobId"); if (el) el.value = String(res.id || "");
    intLoadJobs();
  } catch (e) { alert(e.message); }
}

async function intRunJob() {
  const id = Number(qs("intActiveJobId")?.value || 0);
  if (!id) { alert("job id required"); return; }
  try { const res = await fetchJson(`${API}/api/integrations/jobs/${id}/run`, { method: "POST", body: JSON.stringify({}) }); setStatus(`Run ${id} ${res.status}`); intLoadJobs(); }
  catch (e) { alert(e.message); }
}

async function intRetryJob() {
  const id = Number(qs("intActiveJobId")?.value || 0);
  if (!id) { alert("job id required"); return; }
  try { await fetchJson(`${API}/api/integrations/jobs/${id}/retry-now`, { method: "POST", body: JSON.stringify({}) }); setStatus("Retry queued."); intLoadJobs(); }
  catch (e) { alert(e.message); }
}

async function intCancelJob() {
  const id = Number(qs("intActiveJobId")?.value || 0);
  if (!id) { alert("job id required"); return; }
  const reason = prompt("Cancel reason?") || "cancelled";
  try { await fetchJson(`${API}/api/integrations/jobs/${id}/cancel`, { method: "POST", body: JSON.stringify({ reason }) }); setStatus(`Cancelled ${id}.`); intLoadJobs(); }
  catch (e) { alert(e.message); }
}

async function intLoadJobs() {
  const el = qs("intJobsTable"); if (!el) return;
  const status = String(qs("intJobStatusFilter")?.value || "").trim();
  const q = status ? `?status=${encodeURIComponent(status)}` : "";
  try {
    const res = await fetchJson(`${API}/api/integrations/jobs${q}`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No jobs.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>ID</th><th>Connector</th><th>Status</th><th>Attempts</th><th>Error</th><th>Created</th></tr></thead><tbody>${rows.map((r) => `<tr><td><a href="#" data-int-job-id="${r.id}">${r.id}</a></td><td>${escapeHtml(r.connector_key)}</td><td>${escapeHtml(r.status)}</td><td>${r.attempts}/${r.max_attempts}</td><td>${escapeHtml(r.error_message || "")}</td><td>${escapeHtml((r.created_at || "").slice(0, 16))}</td></tr>`).join("")}</tbody></table>`;
    el.querySelectorAll("[data-int-job-id]").forEach((a) => a.addEventListener("click", (ev) => { ev.preventDefault(); const input = qs("intActiveJobId"); if (input) input.value = a.getAttribute("data-int-job-id"); }));
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function intLoadMonitor() {
  const el = qs("intMonitorSummary"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/integrations/monitoring/summary`);
    el.innerHTML = `
      <div class="kpi-pills">
        ${(res.by_status || []).map((s) => `<span class="kpi-pill"><strong>${escapeHtml(s.status)}:</strong> ${Number(s.c || 0)}</span>`).join(" ")}
        <span class="kpi-pill"><strong>Dead-letter open:</strong> ${Number(res.dead_letter_open || 0)}</span>
        ${res.oldest_queued ? `<span class="kpi-pill"><strong>Oldest queued:</strong> #${res.oldest_queued.id} (${escapeHtml(res.oldest_queued.connector_key)})</span>` : ""}
      </div>
      ${(res.top_errors || []).length ? `<h5>Top errors</h5><ul>${res.top_errors.map((e) => `<li>${escapeHtml(e.error_message || "")} &times; ${Number(e.c || 0)}</li>`).join("")}</ul>` : ""}
    `;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function intLoadDeadLetter() {
  const el = qs("intDeadLetterTable"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/integrations/dead-letter`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No dead-letter entries.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>ID</th><th>Job</th><th>Connector</th><th>Reason</th><th>Ack</th><th>Actions</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${r.id}</td><td>${r.job_id}</td><td>${escapeHtml(r.connector_key)}</td><td>${escapeHtml(r.reason || "")}</td><td>${escapeHtml(r.acknowledged_at || "")}</td><td>${r.acknowledged_at ? "" : `<button type="button" data-int-dl-ack="${r.id}">Ack</button>`}</td></tr>`).join("")}</tbody></table>`;
    el.querySelectorAll("[data-int-dl-ack]").forEach((b) => b.addEventListener("click", async () => {
      const id = b.getAttribute("data-int-dl-ack");
      try { await fetchJson(`${API}/api/integrations/dead-letter/${id}/acknowledge`, { method: "POST", body: JSON.stringify({}) }); intLoadDeadLetter(); }
      catch (e) { alert(e.message); }
    }));
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function govLoadPolicies() {
  const el = qs("govPolicyList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/governance/policies`);
    const rows = res.rows || [];
    el.innerHTML = `<h5>Policies</h5><table class="table"><thead><tr><th>Code</th><th>Mode</th><th>Restricted</th><th>Precursor</th><th>Window (min)</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${escapeHtml(r.policy_code)}</td><td>${escapeHtml(r.mode)}</td><td>${escapeHtml(r.restricted_action)}</td><td>${escapeHtml(r.precursor_action)}</td><td>${r.window_minutes}</td></tr>`).join("")}</tbody></table>`;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function govLoadViolations() {
  const el = qs("govViolationList"); if (!el) return;
  const ack = qs("govViolationFilter")?.value || "";
  const q = ack ? `?ack=${encodeURIComponent(ack)}` : "";
  try {
    const res = await fetchJson(`${API}/api/governance/violations${q}`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No violations.</em>"; return; }
    el.innerHTML = `<h5>Violations</h5><table class="table"><thead><tr><th>ID</th><th>Policy</th><th>User</th><th>Restricted</th><th>Blocked</th><th>Ack</th><th>Actions</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${r.id}</td><td>${escapeHtml(r.policy_code)}</td><td>${escapeHtml(r.username)}</td><td>${escapeHtml(r.restricted_action)}</td><td>${r.blocked ? "yes" : "no"}</td><td>${escapeHtml(r.acknowledged_at || "")}</td><td>${r.acknowledged_at ? "" : `<button type="button" data-gov-ack="${r.id}">Acknowledge</button>`}</td></tr>`).join("")}</tbody></table>`;
    el.querySelectorAll("[data-gov-ack]").forEach((b) => b.addEventListener("click", async () => {
      const id = b.getAttribute("data-gov-ack");
      const notes = prompt("Resolution notes (optional)") || "";
      try { await fetchJson(`${API}/api/governance/violations/${id}/acknowledge`, { method: "POST", body: JSON.stringify({ notes }) }); govLoadViolations(); }
      catch (e) { alert(e.message); }
    }));
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function govEvaluate() {
  const action = String(qs("govEvalAction")?.value || "").trim();
  const username = String(qs("govEvalUser")?.value || "").trim();
  if (!action) { alert("action required"); return; }
  try {
    const res = await fetchJson(`${API}/api/governance/evaluate`, { method: "POST", body: JSON.stringify({ action, username }) });
    const el = qs("govEvalResult");
    if (el) el.innerHTML = `<span class="kpi-pill"><strong>Blocked:</strong> ${res.blocked ? "YES" : "no"}</span> <span class="kpi-pill"><strong>Violations:</strong> ${(res.violations || []).length}</span>`;
    govLoadViolations();
  } catch (e) { alert(e.message); }
}

async function evBuildPack() {
  const pack_type = qs("evPackType")?.value;
  const period = String(qs("evPackPeriod")?.value || "").trim();
  const site_code = String(qs("evPackSite")?.value || "").trim() || null;
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("period YYYY-MM required"); return; }
  try {
    const res = await fetchJson(`${API}/api/executive/evidence-packs/build`, { method: "POST", body: JSON.stringify({ pack_type, period, site_code }) });
    setStatus(`Pack built ${res.pack_code}`);
    evLoadPacks();
  } catch (e) { alert(e.message); }
}

async function evLoadPacks() {
  const el = qs("evPackList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/executive/evidence-packs`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No packs.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>ID</th><th>Code</th><th>Type</th><th>Period</th><th>Site</th><th>Created</th></tr></thead><tbody>${rows.map((r) => `<tr><td><a href="#" data-ev-view-id="${r.id}">${r.id}</a></td><td>${escapeHtml(r.pack_code)}</td><td>${escapeHtml(r.pack_type)}</td><td>${escapeHtml(r.period)}</td><td>${escapeHtml(r.site_code || "")}</td><td>${escapeHtml((r.created_at || "").slice(0, 16))}</td></tr>`).join("")}</tbody></table>`;
    el.querySelectorAll("[data-ev-view-id]").forEach((a) => a.addEventListener("click", (ev) => { ev.preventDefault(); const input = qs("evPackViewId"); if (input) input.value = a.getAttribute("data-ev-view-id"); }));
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function evLoadPackDetail() {
  const id = Number(qs("evPackViewId")?.value || 0);
  if (!id) { alert("Enter pack id"); return; }
  const el = qs("evPackDetail"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/executive/evidence-packs/${id}`);
    const pack = res.pack || {};
    const summary = pack.summary || {};
    el.innerHTML = `
      <div class="kpi-pills">
        <span class="kpi-pill"><strong>Code:</strong> ${escapeHtml(pack.pack_code || "")}</span>
        <span class="kpi-pill"><strong>Type:</strong> ${escapeHtml(pack.pack_type || "")}</span>
        <span class="kpi-pill"><strong>Integrity:</strong> ${pack.integrity_ok ? "OK" : "FAIL"}</span>
        <span class="kpi-pill"><strong>Hash:</strong> ${escapeHtml((pack.integrity_hash || "").slice(0, 12))}</span>
      </div>
      <h5>Summary</h5>
      <pre class="output">${escapeHtml(JSON.stringify(summary, null, 2))}</pre>
    `;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function execLoadCommand() {
  const period = String(qs("execPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("period YYYY-MM required"); return; }
  try {
    const res = await fetchJson(`${API}/api/executive/command-center?period=${encodeURIComponent(period)}`);
    const el = qs("execCommandKpi"); if (!el) return;
    const ops = res.operational || {};
    const gov = res.governance || {};
    const integ = res.integrations || {};
    el.innerHTML = `
      <div class="kpi-pills">
        <span class="kpi-pill"><strong>Run hrs:</strong> ${Number(ops.run_hours || 0).toFixed(1)}</span>
        <span class="kpi-pill"><strong>Downtime hrs:</strong> ${Number(ops.downtime_hours || 0).toFixed(1)}</span>
        <span class="kpi-pill"><strong>Breakdowns:</strong> ${ops.breakdowns || 0}</span>
        <span class="kpi-pill"><strong>Assets used:</strong> ${ops.used_assets_total || 0}</span>
        <span class="kpi-pill"><strong>Gov violations (open):</strong> ${gov.open || 0}</span>
        <span class="kpi-pill"><strong>Dead-letter:</strong> ${integ.dead_letter_open || 0}</span>
      </div>
    `;
  } catch (e) { alert(e.message); }
}

async function execLoadBoard() {
  const period = String(qs("execPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("period YYYY-MM required"); return; }
  try {
    const res = await fetchJson(`${API}/api/executive/board-pack?period=${encodeURIComponent(period)}`);
    const narr = qs("execNarrative");
    if (narr) narr.innerHTML = `<h5>Highlights</h5><ul>${(res.narrative?.bullets || []).map((b) => `<li>${escapeHtml(b)}</li>`).join("")}</ul>`;
    const blocks = qs("execBoardBlocks");
    if (blocks) {
      blocks.innerHTML = `
        <h5>Financial</h5>
        <pre class="output">${escapeHtml(JSON.stringify(res.financial || {}, null, 2))}</pre>
        <h5>Operational</h5>
        <pre class="output">${escapeHtml(JSON.stringify(res.operational || {}, null, 2))}</pre>
        <h5>Integrations</h5>
        <pre class="output">${escapeHtml(JSON.stringify(res.integrations || {}, null, 2))}</pre>
        <h5>Governance</h5>
        <pre class="output">${escapeHtml(JSON.stringify(res.governance || {}, null, 2))}</pre>
      `;
    }
  } catch (e) { alert(e.message); }
}

function execExportBoardXlsx() {
  const period = String(qs("execPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("period YYYY-MM required"); return; }
  window.open(`${API}/api/executive/board-pack/export.xlsx?period=${encodeURIComponent(period)}`, "_blank");
}

function bindEnterpriseHandlers() {
  const bind = (id, ev, fn) => { const el = qs(id); if (el) el.addEventListener(ev, fn); };
  bind("entSaveCompanyBtn", "click", entSaveCompany);
  bind("entLoadCompaniesBtn", "click", entLoadCompanies);
  bind("entSaveSiteBtn", "click", entSaveSite);
  bind("entLoadSitesBtn", "click", entLoadSites);
  bind("entLoadTreeBtn", "click", entLoadTree);
  bind("entSaveCcyBtn", "click", entSaveCcyRates);
  bind("entLoadCcyBtn", "click", entLoadCcyRates);
  bind("entSaveTaxBtn", "click", entSaveTax);
  bind("entLoadTaxBtn", "click", entLoadTax);
  bind("intSaveConnBtn", "click", intSaveConnection);
  bind("intLoadConnBtn", "click", intLoadConnections);
  bind("intEnqueueBtn", "click", intEnqueueJob);
  bind("intRunJobBtn", "click", intRunJob);
  bind("intRetryJobBtn", "click", intRetryJob);
  bind("intCancelJobBtn", "click", intCancelJob);
  bind("intLoadJobsBtn", "click", intLoadJobs);
  bind("intJobStatusFilter", "change", intLoadJobs);
  bind("intLoadMonitorBtn", "click", intLoadMonitor);
  bind("intLoadDeadLetterBtn", "click", intLoadDeadLetter);
  bind("govLoadPoliciesBtn", "click", govLoadPolicies);
  bind("govLoadViolationsBtn", "click", govLoadViolations);
  bind("govViolationFilter", "change", govLoadViolations);
  bind("govEvalBtn", "click", govEvaluate);
  bind("evBuildPackBtn", "click", evBuildPack);
  bind("evLoadPacksBtn", "click", evLoadPacks);
  bind("evLoadPackDetailBtn", "click", evLoadPackDetail);
  bind("execLoadCommandBtn", "click", execLoadCommand);
  bind("execLoadBoardBtn", "click", execLoadBoard);
  bind("execExportBoardXlsxBtn", "click", execExportBoardXlsx);
}

document.addEventListener("DOMContentLoaded", () => {
  initDarkMode();
  init().catch((e) => console.error(e));
  try {
    initFinanceTab();
    bindFinanceHandlers();
  } catch (e) { console.error("finance bind failed", e); }
  try { bindEnterpriseHandlers(); } catch (e) { console.error("enterprise bind failed", e); }
  try { if (typeof window.bindStoreQrAdmin === "function") window.bindStoreQrAdmin(); } catch (e) { console.error("store QR bind failed", e); }
});
