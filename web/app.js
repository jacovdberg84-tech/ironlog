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
const DEFAULT_ROLE = "admin";
const DEFAULT_USER = "admin";
const DEFAULT_SITE = "main";
const LANG_KEY = "ironlog_lang";
const DEFAULT_LANG = "en";
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
    if (res.status === 401) {
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

function getRoleAllowedTabs(role) {
  const r = String(role || "").toLowerCase();
  if (r === "operator") return ["dash", "daily", "fuel", "lube", "legal", "operations", "docs", "vehicle"];
  if (r === "artisan") return ["dash", "maintenance", "Breakdowns", "reports", "fuel", "lube", "legal", "operations", "dispatch", "docs", "vehicle"];
  if (r === "stores") return ["dash", "maintenance", "stock", "uploads", "reports", "legal", "procurement", "operations", "dispatch", "quality", "docs", "vehicle"];
  if (r === "supervisor") return ["dash", "daily", "assets", "maintenance", "fuel", "lube", "stock", "legal", "uploads", "reports", "Breakdowns", "approvals", "procurement", "operations", "dispatch", "quality", "audit", "docs", "vehicle"];
  return [
    "dash",
    "daily",
    "assets",
    "maintenance",
    "fuel",
    "lube",
    "stock",
    "legal",
    "uploads",
    "reports",
    "Breakdowns",
    "approvals",
    "procurement",
    "operations",
    "dispatch",
    "quality",
    "audit",
    "docs",
    "vehicle",
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
      opt.hidden = !allowed.has(opt.value);
    });
  }

  const maintBtn = qs("openMaintenanceBtn");
  if (maintBtn) maintBtn.style.display = allowed.has("maintenance") ? "" : "none";
  const siteBtn = qs("openSiteOpsBtn");
  if (siteBtn) siteBtn.style.display = allowed.has("operations") ? "" : "none";
  const reopenBtn = qs("reopenOperationsDay");
  if (reopenBtn) reopenBtn.style.display = roles.some((r) => ["admin", "supervisor"].includes(r)) ? "" : "none";

  const activePanel = document.querySelector(".panel.show");
  const activeKey = String(activePanel?.id || "").replace(/^tab-/, "");
  if (!activeKey || !allowed.has(activeKey)) {
    const firstAllowed = allowedList[0];
    if (firstAllowed) {
      switchTab(firstAllowed);
    }
  } else if (tabSelect) {
    tabSelect.value = activeKey;
  }
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
  let res;
  let data = {};
  try {
    res = await fetch(`${API}/api/auth/me`, { headers: new Headers(authHeaders()) });
    data = await res.json();
  } catch {
    res = { status: 0 };
  }
  if (res.status === 401) {
    showLoginGate(true);
    updateAuthChrome();
    return;
  }
  if (data.ok && data.user && data.user.id != null) {
    applySessionFromMeUser(data.user);
  } else if (data.ok && data.user && data.user.id == null) {
    localStorage.removeItem(TABS_OVERRIDE_KEY);
  }
  showLoginGate(false);
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
async function ensureAdminTabOptions() {
  if (__adminTabKeysLoaded) return;
  const sel = qs("adminUserTabs");
  if (!sel) return;
  try {
    const data = await fetchJson(`${API}/api/auth/tabs`);
    const keys = Array.isArray(data.keys) ? data.keys : [];
    sel.innerHTML = "";
    keys.forEach((k) => {
      const opt = document.createElement("option");
      opt.value = k;
      opt.textContent = k;
      sel.appendChild(opt);
    });
    __adminTabKeysLoaded = true;
  } catch {}
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
    tbody.innerHTML = "";
    rows.forEach((r) => {
      const tr = document.createElement("tr");
      const rolesText = Array.isArray(r.roles) && r.roles.length ? r.roles.join(", ") : String(r.role || "operator");
      const locText = Array.isArray(r.allowed_locations) && r.allowed_locations.length ? r.allowed_locations.join(", ") : "all";
      tr.innerHTML = `<td>${escapeHtml(r.username)}</td><td>${escapeHtml(r.full_name || "")}</td><td>${escapeHtml(r.department || "")}</td><td>${escapeHtml(rolesText)}</td><td>${escapeHtml(locText)}</td><td>${r.active ? "yes" : "no"}</td><td>${r.has_password ? "yes" : "no"}</td>`;
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

async function saveAdminUser() {
  const username = String(qs("adminUsername")?.value || "").trim();
  const password = String(qs("adminPassword")?.value || "");
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

function refreshNetBanner() {
  setNetBanner("idle", getQueuedHoursCount());
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

async function syncOfflineHoursQueue() {
  if (!navigator.onLine) return { ok: false, reason: "offline" };

  const queue = getQueue();
  const hoursItems = queue.filter((q) => q.type === "HOURS");
  if (!hoursItems.length) return { ok: true, synced: 0 };

  setNetBanner("syncing", hoursItems.length);
  setStatus(`Syncing offline queue (${hoursItems.length})...`);

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
    setStatus(`Sync finished: ${hoursItems.length - failed.length} ok, ${failed.length} failed.`);
    refreshNetBanner();
    return { ok: false, synced: hoursItems.length - failed.length, failed };
  }

  setStatus("Sync finished: all queued hours synced ✅");
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
   DASHBOARD
========================= */

async function loadDashboard() {
  const dateEl = qs("date");
  const scheduledEl = qs("scheduled");
  const date = dateEl ? dateEl.value : new Date().toISOString().slice(0, 10);
  const scheduled = scheduledEl ? scheduledEl.value || 10 : 10;

  setStatus("Loading dashboard...");
  setSkeleton("downtimeList", 2);
  setSkeleton("downtimeReasonsList", 2);
  setSkeleton("stockList", 2);
  setSkeleton("woList", 2);
  setSkeleton("riskBoardList", 2);
  setSkeleton("slaList", 2);
  setSkeleton("costTrendList", 2);
  setSkeleton("costList", 2);
  setSkeleton("lubeList", 2);
  setSkeleton("stockMonitorList", 2);

  const data = await fetchJson(`${API}/api/dashboard?date=${date}&scheduled=${scheduled}`);

  const sqDateEl = qs("sqDate");
  if (sqDateEl && !sqDateEl.value) sqDateEl.value = date;

  const _kpiTh = getThresholds();
  setSpeedo(qs("availNeedle"), qs("gAvailVal"), data?.kpi?.availability, { goodAt: _kpiTh.availTarget, warnAt: _kpiTh.availCrit });
  setSpeedo(qs("utilNeedle"), qs("gUtilVal"), data?.kpi?.utilization, { goodAt: _kpiTh.utilTarget, warnAt: _kpiTh.utilCrit });
  updateKpiAlertBanner(data?.kpi?.availability, data?.kpi?.utilization);

  const mtdRange =
    data.kpi?.mtd_start && data.kpi?.mtd_end
      ? `${data.kpi.mtd_start} → ${data.kpi.mtd_end}`
      : "";
  setText(
    "availMeta",
    mtdRange
      ? `MTD ${mtdRange} · Used assets: ${data.kpi.used_assets} | Avail hrs: ${data.kpi.available_hours} | Downtime: ${data.kpi.downtime_hours}`
      : `Used assets: ${data.kpi.used_assets} | Avail hrs: ${data.kpi.available_hours} | Downtime: ${data.kpi.downtime_hours}`
  );
  setText(
    "utilMeta",
    mtdRange
      ? `MTD ${mtdRange} · Run hrs: ${data.kpi.run_hours} | Scheduled/asset (header): ${data.scheduled_hours_per_asset}`
      : `Run hrs: ${data.kpi.run_hours} | Scheduled/asset: ${data.scheduled_hours_per_asset}`
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
      if (!rows.length) debugList.appendChild(item("<small>No production assets found for selected date (per-asset rows are for this day only; gauges are month-to-date).</small>"));
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
    (data.open_work_orders || []).forEach((r) => {
      woList.appendChild(
        item(`<b>WO #${r.id}</b> – ${r.asset_code}<br><small>${r.source} | ${r.status} | ${r.opened_at}</small>`)
      );
    });
    if (!data.open_work_orders?.length) woList.appendChild(item("<small>No open work orders.</small>"));
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

  const sla = data.workorder_sla || {};
  const slaSummary = sla.summary || {};
  setText("slaOpen24", Number(slaSummary.open_gt_24h || 0));
  setText("slaProgress48", Number(slaSummary.in_progress_gt_48h || 0));
  setText("slaCompleted12", Number(slaSummary.completed_gt_12h || 0));
  const slaList = qs("slaList");
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

  if (slaList && !slaList.dataset.bound) {
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

  const trend = await fetchJson(`${API}/api/dashboard/cost/trend?months=12`);
  const trendRows = Array.isArray(trend.rows) ? trend.rows : [];
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
  const lubeList = qs("lubeList");
  if (lubeList) {
    lubeList.innerHTML = "";
    (lube.rows || []).forEach((r) => {
      lubeList.appendChild(
        item(`<b>${r.asset_code}</b> – ${Number(r.qty || 0).toFixed(1)} qty<br><small>${r.asset_name || ""}</small>`)
      );
    });
    if (!lube.rows?.length) lubeList.appendChild(item("<small>No lube logs for this date.</small>"));
  }

  await loadStockMonitor().catch(() => {});
  await loadIronmindInsight({ silent: true }).catch(() => {});
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
           "", "Data Gaps", "- Insufficient data"].join("\n")
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
      body: JSON.stringify({ question, date }),
    });
    const short = String(res?.short_answer || "No answer returned.");
    if (out) out.innerHTML = `<div>${escapeHtml(short)}</div>`;
    setStatus("IRONMIND question answered.");
  } catch (e) {
    if (out) out.innerHTML = `<small class="muted">Question failed: ${escapeHtml(e.message || String(e))}</small>`;
    setStatus("IRONMIND ask error: " + (e.message || e));
  }
}

function parseIronmindSections(text) {
  const defs = [
    { key: "repairs", name: "Repairs Needed" },
    { key: "risks", name: "Operational Risks" },
    { key: "suggestions", name: "Suggestions" },
    { key: "data_gaps", name: "Data Gaps" },
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
  const sel = qs("histAsset");
  if (!sel) return;
  if (sel.options.length <= 1) {
    await populateHistoryAssets().catch(() => {});
  }
  const exists = Array.from(sel.options).some((o) => o.value === assetCode);
  if (exists) {
    sel.value = assetCode;
    await loadAssetHistory().catch(() => {});
  }
}

function getDefaultLubeRange() {
  const end = new Date();
  const start = new Date();
  start.setDate(end.getDate() - 29);
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
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

  setText("lubeQtyTotal", Number(data.summary?.qty_total || 0).toFixed(1));
  setText("lubeEntries", Number(data.summary?.entries || 0));
  setText("lubeAssets", Number(data.summary?.assets || 0));

  const lubeList = qs("lubeList");
  if (lubeList) {
    lubeList.innerHTML = "";
    (data.rows || []).forEach((r) => {
      lubeList.appendChild(
        item(
          `<b>${r.asset_code}</b> – ${Number(r.qty_total || 0).toFixed(1)} qty` +
          `<br><small>${r.asset_name || ""} | Entries: ${Number(r.entries || 0)}</small>`
        )
      );
    });
    if (!data.rows?.length) lubeList.appendChild(item("<small>No lube usage in this period.</small>"));
  }

  setStatus("Lube usage ready.");
}

async function saveFuelLog() {
  const meterRaw = (qs("fuelHoursRun")?.value || "").trim();
  const meter_unit = String(qs("fuelMeterUnit")?.value || "hours").trim().toLowerCase() === "km" ? "km" : "hours";
  const basePayload = {
    asset_code: (qs("fuelAsset")?.value || "").trim(),
    log_date: (qs("fuelDate")?.value || "").trim() || undefined,
    liters: Number(qs("fuelLiters")?.value || 0),
    meter_run_value: meterRaw === "" ? undefined : Number(meterRaw),
    meter_unit,
    hours_run: meter_unit === "hours" && meterRaw !== "" ? Number(meterRaw) : undefined,
    source: (qs("fuelSource")?.value || "").trim() || undefined,
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

async function loadFuelBenchmark() {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const mode = String(qs("fuelModeFilter")?.value || "").trim();
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  const duplicatesOnly = Boolean(qs("fuelDupOnly")?.checked);
  if (!start || !end) return alert("Select start and end dates.");

  setStatus("Loading fuel benchmark...");
  setSkeleton("fuelBenchmarkList", 2);
  const data = duplicatesOnly
    ? await fetchJson(
      `${API}/api/dashboard/fuel/duplicates?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&mode=${encodeURIComponent(mode)}&asset_code=${encodeURIComponent(assetCode)}`
    )
    : await fetchJson(
      `${API}/api/dashboard/fuel?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}&mode=${encodeURIComponent(mode)}&asset_code=${encodeURIComponent(assetCode)}`
    );

  if (duplicatesOnly) {
    setText("fbFuelTotal", Number(data.summary?.fuel_liters || 0).toFixed(2));
    setText("fbHoursTotal", "-");
    setText("fbKmTotal", "-");
    setText("fbAvgLph", "-");
    setText("fbAvgKmpl", "-");
    setText("fbExcessive", Number(data.summary?.duplicate_rows || 0));
  } else {
    setText("fbFuelTotal", Number(data.summary?.fuel_liters || 0).toFixed(2));
    setText("fbHoursTotal", Number(data.summary?.hours_run || 0).toFixed(2));
    setText("fbKmTotal", Number(data.summary?.km_run || 0).toFixed(2));
    setText("fbAvgLph", data.summary?.avg_lph == null ? "-" : Number(data.summary.avg_lph).toFixed(3));
    setText("fbAvgKmpl", data.summary?.avg_km_per_l == null ? "-" : Number(data.summary.avg_km_per_l).toFixed(3));
    setText("fbExcessive", Number(data.summary?.excessive_count || 0));
  }

  const list = qs("fuelBenchmarkList");
  if (list) {
    list.innerHTML = "";
    (data.rows || []).forEach((r) => {
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
      const flag = r.is_excessive
        ? "<span class='pill red'>EXCESSIVE</span>"
        : "<span class='pill blue'>OK</span>";
      const machineKey = String(r.asset_code || "").replace(/[^A-Za-z0-9_-]/g, "_");
      list.appendChild(
        item(
          `<div class="fuel-item-head"><b>${r.asset_code}</b> — ${
            String(r.metric_mode || "hours") === "km"
              ? `${r.actual_km_per_l == null ? "-" : Number(r.actual_km_per_l).toFixed(3)} km/L`
              : `${r.actual_lph == null ? "-" : Number(r.actual_lph).toFixed(3)} L/hr`
          } ${flag}</div>` +
          `<small class="fuel-item-desc">${r.asset_name || ""}</small>` +
          `<small class="fuel-item-meta">${
            String(r.metric_mode || "hours") === "km"
              ? `OEM: ${Number(r.oem_km_per_l || 0).toFixed(3)} km/L | Fuel: ${Number(r.fuel_liters || 0).toFixed(2)}L | Distance: ${Number(r.km_run || 0).toFixed(2)} km`
              : `OEM: ${Number(r.oem_lph || 0).toFixed(3)} L/hr | Fuel: ${Number(r.fuel_liters || 0).toFixed(2)}L | Hours: ${Number(r.hours_run || 0).toFixed(2)}`
          }</small>` +
          `<br><button data-fuel-machine="${String(r.asset_code || "").replace(/"/g, "&quot;")}">Open machine history</button> ` +
          `<button data-fuel-machine-pdf="${String(r.asset_code || "").replace(/"/g, "&quot;")}">Machine PDF</button>` +
          `<div class="fuel-inline-history" id="fuel-inline-${machineKey}"></div>`
        )
      );
    });
    if (!data.rows?.length) {
      list.appendChild(item(`<small>${duplicatesOnly ? "No duplicate fuel entries found in this period." : "No fuel benchmark data in this period."}</small>`));
    }
  }

  if (duplicatesOnly) {
    setStatus(`Duplicate filter ready (${Number(data.summary?.duplicate_rows || 0)} rows in ${Number(data.summary?.duplicate_groups || 0)} groups).`);
  } else {
    setStatus("Fuel benchmark ready.");
  }
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
  const top = `<div class="fuel-inline-summary"><small><b>Fill days:</b> ${Number(data.summary?.days || 0)} | <b>Fuel:</b> ${Number(data.summary?.fuel_liters || 0).toFixed(2)}L | <b>Fill ${mode === "km" ? "distance" : "hours"}:</b> ${Number(mode === "km" ? (data.summary?.km_run || 0) : (data.summary?.hours_run || 0)).toFixed(2)} | <b>Avg:</b> ${mode === "km" ? (data.summary?.avg_km_per_l == null ? "-" : Number(data.summary.avg_km_per_l).toFixed(3) + " km/L") : (data.summary?.avg_lph == null ? "-" : Number(data.summary.avg_lph).toFixed(3) + " L/hr")} | <b>Over benchmark days:</b> ${Number(data.summary?.excessive_days || 0)}</small></div>`;
  const tableRows = rows.map((r) => {
    const statusClass = r.is_excessive ? "fh-status-excessive" : "fh-status-ok";
    const statusText = r.is_excessive ? "EXCESSIVE" : "OK";
    return (
      `<tr>` +
      `<td class="fh-col-date">${r.log_date}</td>` +
      `<td class="fh-col-num">${Number(r.fuel_liters || 0).toFixed(2)}</td>` +
      `<td class="fh-col-num">${Number((mode === "km" ? r.km_run : r.hours_run) || 0).toFixed(2)}</td>` +
      `<td class="fh-col-num">${mode === "km" ? (r.actual_km_per_l == null ? "-" : Number(r.actual_km_per_l).toFixed(3)) : (r.actual_lph == null ? "-" : Number(r.actual_lph).toFixed(3))}</td>` +
      `<td class="fh-col-status"><span class="fh-status ${statusClass}">${statusText}</span></td>` +
      `<td class="fh-col-action"><button data-fuel-delete="${Number(r.id || 0)}">Delete</button></td>` +
      `</tr>`
    );
  }).join("");
  mountEl.innerHTML =
    `${top}<br>` +
    (tableRows
      ? `<div class="fuel-history-table-wrap"><table class="fuel-history-table"><colgroup><col style="width:18%"><col style="width:16%"><col style="width:16%"><col style="width:14%"><col style="width:16%"><col style="width:20%"></colgroup><thead><tr><th class="fh-col-date">Date</th><th class="fh-col-num">Fuel (L)</th><th class="fh-col-num">${mode === "km" ? "Distance Between Fills (km)" : "Hours Between Fills"}</th><th class="fh-col-num">${mode === "km" ? "km/L" : "L/hr"}</th><th class="fh-col-status">Status</th><th class="fh-col-action">Action</th></tr></thead><tbody>${tableRows}</tbody></table></div>`
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

  const list = qs("spList");
  if (list) {
    list.innerHTML = "";
    const onlyLow = Boolean(qs("spOnlyLow")?.checked);
    const baseRows = onlyLow
      ? stockPageData.rows.filter((r) => Boolean(r.below_min))
      : stockPageData.rows;
    const sortedRows = sortStockRows(baseRows);
    sortedRows.forEach((r) => {
      const tone = r.below_min
        ? "border-left:4px solid #b04747;background:rgba(176,71,71,0.08);"
        : r.critical
        ? "border-left:4px solid #b08947;background:rgba(176,137,71,0.08);"
        : "";
      list.appendChild(
        item(
          `<div style="${tone}padding-left:8px;">` +
          `<b>${r.part_code}</b> – ${Number(r.on_hand || 0).toFixed(1)} on hand ${
            r.below_min ? "<span class='pill red'>LOW</span>" : ""
          }${r.critical ? " <span class='pill orange'>CRITICAL</span>" : ""}` +
          `<br><small>${r.part_name || ""} | Min: ${Number(r.min_stock || 0).toFixed(1)} | Unit: $${Number(r.unit_cost || 0).toFixed(2)} | Value: $${Number(r.stock_value || 0).toFixed(2)}</small>` +
          `</div>`
        )
      );
    });
    if (!sortedRows.length) {
      list.appendChild(item("<small>No parts found for current filter.</small>"));
    }
  }

  const recent = qs("spRecent");
  if (recent) {
    recent.innerHTML = "";
    stockPageData.recent.forEach((r) => {
      recent.appendChild(
        item(
          `<b>${r.part_code}</b> – ${Number(r.quantity || 0).toFixed(1)} (${r.movement_type})` +
          `<br><small>${r.created_at || ""} | ${r.location_code || "NO-LOC"} | ${r.reference || "-"}</small>`
        )
      );
    });
    if (!stockPageData.recent.length) {
      recent.appendChild(item("<small>No stock movements yet.</small>"));
    }
  }

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

function openStockOnHandPdf() {
  const filter = (qs("spFilter")?.value || "").trim();
  const q = filter ? `?part_code=${encodeURIComponent(filter)}` : "";
  window.open(`${API}/api/reports/stock-monitor.pdf${q}`, "_blank");
}

async function loadAuditLogs() {
  const module = (qs("auditModule")?.value || "").trim();
  const action = (qs("auditAction")?.value || "").trim();
  const entity_type = (qs("auditEntityType")?.value || "").trim();
  const username = (qs("auditUsername")?.value || "").trim();
  const limit = Number(qs("auditLimit")?.value || 200);
  const list = qs("auditList");
  if (!list) return;

  setStatus("Loading audit trail...");
  setSkeleton("auditList", 2);

  const q = new URLSearchParams();
  if (module) q.set("module", module);
  if (action) q.set("action", action);
  if (entity_type) q.set("entity_type", entity_type);
  if (username) q.set("username", username);
  if (Number.isFinite(limit) && limit > 0) q.set("limit", String(Math.trunc(limit)));

  const url = `${API}/api/audit${q.toString() ? `?${q.toString()}` : ""}`;
  const data = await fetchJson(url);
  const rows = Array.isArray(data.rows) ? data.rows : [];

  list.innerHTML = "";
  rows.forEach((r) => {
    list.appendChild(
      item(
        `<b>${r.created_at || "-"}</b> — ${r.module}.${r.action} ` +
        `<span class="pill blue">${r.role || "-"}</span>` +
        `<br><small>${r.username || "-"} | ${r.entity_type || "-"}:${r.entity_id || "-"}</small>` +
        (r.payload ? `<br><small>${JSON.stringify(r.payload)}</small>` : "")
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No audit records found.</small>"));

  setStatus("Audit trail ready.");
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
  document.querySelectorAll(".panel").forEach((p) => p.classList.remove("show"));
  const panel = qs(`tab-${k}`);
  if (panel) panel.classList.add("show");
  const tabSelect = qs("tabSelect");
  if (tabSelect && tabSelect.value !== k) tabSelect.value = k;
}

function initTabs() {
  const tabSelect = qs("tabSelect");
  if (!tabSelect) return;
  tabSelect.addEventListener("change", () => switchTab(tabSelect.value));
  if (!document.querySelector(".panel.show")) {
    switchTab(tabSelect.value || "dash");
  }
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
  const resultEl = qs("fuelFamsResult");
  if (!fileEl || !resultEl) return;

  const file = fileEl.files[0];
  if (!file) return alert("Choose a FAMS CSV file first.");

  const fd = new FormData();
  fd.append("file", file);

  setStatus("Importing FAMS fuel file...");
  resultEl.textContent = "";
  try {
    const res = await fetchJson(`${API}/api/upload/fuel`, { method: "POST", body: fd });
    resultEl.textContent = JSON.stringify(res, null, 2);
    setStatus("FAMS fuel import complete.");
    await loadDashboard().catch(() => {});
  } catch (e) {
    resultEl.textContent = String(e.message || e);
    setStatus("FAMS fuel import failed.");
  }
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
  const q = month
    ? `month=${encodeURIComponent(month)}`
    : `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  window.open(`${API}/api/reports/maintenance-exec.pptx?${q}`, "_blank");
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
  } catch (e) {
    setText("rainDaysResult", String(e.message || e));
    setStatus("Remove rain day failed.");
  }
}

async function loadRainDays() {
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
  } else {
    alert("Select month or start/end first.");
    return;
  }
  setStatus("Loading rain days...");
  try {
    const res = await fetchJson(`${API}/api/reports/rain-days?${q}`);
    setText("rainDaysResult", JSON.stringify(res, null, 2));
    setStatus("Rain days loaded.");
  } catch (e) {
    setText("rainDaysResult", String(e.message || e));
    setStatus("Load rain days failed.");
  }
}

function openDailyPdf() {
  const date = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const scheduled = qs("scheduled")?.value || 10;
  const ts = Date.now();
  window.open(`${API}/api/reports/daily.pdf?date=${date}&scheduled=${scheduled}&_ts=${ts}`, "_blank");
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
  const asset_code = (qs("histAsset")?.value || "").trim();
  if (!asset_code) {
    alert("Select an asset first.");
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

async function createBreakdown() {
  const date = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const payload = {
    asset_code: (qs("bAsset")?.value || "").trim(),
    breakdown_date: date,
    description: (qs("bDesc")?.value || "").trim(),
    downtime_hours: Number(qs("bDown")?.value || 0),
    critical: !!qs("bCrit")?.checked,
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
  } catch (e) {
    setText("breakdownResult", String(e.message || e));
    setStatus("Breakdown failed.");
  }
}

/** Short breakdown: closed incident + downtime logs + parts + oils in one POST. */
async function submitShortBreakdown() {
  const headerDate = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const breakdown_date = (qs("sqDate")?.value || headerDate).trim();
  const asset_code = (qs("sqAsset")?.value || "").trim();
  const description = (qs("sqDesc")?.value || "").trim();
  const td = (qs("sqTimeDown")?.value || "").trim();
  const tu = (qs("sqTimeUp")?.value || "").trim();
  const comp = (qs("sqComponent")?.value || "").trim();

  const parts = [];
  const p1 = (qs("sqPart1")?.value || "").trim();
  const q1 = Number(qs("sqQty1")?.value || 0);
  if (p1 && Number.isFinite(q1) && q1 > 0) parts.push({ part_code: p1, quantity: q1 });
  const p2 = (qs("sqPart2")?.value || "").trim();
  const q2 = Number(qs("sqQty2")?.value || 0);
  if (p2 && Number.isFinite(q2) && q2 > 0) parts.push({ part_code: p2, quantity: q2 });

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
    asset_code: (qs("saAsset")?.value || "").trim() || undefined,
    work_order_id: (qs("saWo")?.value || "").trim() ? Number((qs("saWo")?.value || "").trim()) : undefined,
    allocation_date: (qs("saDate")?.value || "").trim() || undefined,
    issued_by: (qs("saIssuedBy")?.value || "").trim() || undefined,
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
        `<small>${ref} | ${r.location_code || "NO-LOC"} | Unit: $${unitCost.toFixed(2)} | Value: $${lineValue.toFixed(2)} | ${r.issued_by || "No issuer"}${r.notes ? ` | ${r.notes}` : ""}</small>`
      )
    );
  });

  if (!data.length) {
    list.appendChild(item("<small>No store allocations yet.</small>"));
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
      assetList.innerHTML = "";
      (Array.isArray(assets) ? assets : []).forEach((a) => {
        const code = String(a.asset_code || "").trim();
        if (!code) return;
        const hiredTag = isHiredAsset(a) ? " [HIRED]" : "";
        const opt = document.createElement("option");
        opt.value = code;
        opt.textContent = `${code}${hiredTag} - ${a.asset_name || ""}`;
        assetList.appendChild(opt);
      });
    } catch {}
  }

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
    movement_type,
    quantity: Number(qs("msQty")?.value || 0),
    reference: (qs("msRef")?.value || "").trim() || undefined,
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
  const payload = {
    asset_code: (qs("mlAsset")?.value || "").trim(),
    log_date: (qs("mlDate")?.value || "").trim() || undefined,
    part_code: part_code || undefined,
    location_code: (qs("mlLocation")?.value || "").trim() || undefined,
    oil_type: (qs("mlType")?.value || "").trim() || undefined,
    quantity: Number(qs("mlQty")?.value || 0),
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
          `<br>${advanceBtn} ${finalizeBtn} ${postBtn} ${routeBtn} ${approveBtn} ${submitBtn} ${receiveBtn} ${receiveHalfBtn} ${receiveFullBtn} <button data-pr-duplicate="${dupPayload}" style="margin-top:8px;">Duplicate</button> ${r.latest_approval_id ? `<button data-pr-open-approval-id="${r.latest_approval_id}" style="margin-top:8px;">Open Approval</button>` : ""}`
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

  for (const r of used) {
    totalScheduled += Number(r.scheduled_hours || 0);
    totalRun += Number(r.hours_run || 0);
  }

  const ratio = totalScheduled > 0 ? totalRun / totalScheduled : null;
  const pct = ratio == null ? null : ratio * 100;

  return {
    usedCount,
    totalScheduled,
    totalRun,
    availability: pct,
    utilization: pct,
  };
}

function renderDailyPreview() {
  setText("dailySummary", daySummary());

  const k = calcDailyPreviewKpis();

  setText("kUsed", `Used: ${k.usedCount}`);
  setText("kSched", `Scheduled: ${k.totalScheduled.toFixed(1).replace(/\.0$/, "")}`);
  setText("kRun", `Run: ${k.totalRun.toFixed(1).replace(/\.0$/, "")}`);

  setSpeedo(qs("pAvailNeedle"), qs("pAvailVal"), k.availability);
  setSpeedo(qs("pUtilNeedle"), qs("pUtilVal"), k.utilization);

  if (k.totalScheduled === 0) setText("kNote", "Preview waiting for scheduled/run hours. Standby excluded.");
  else setText("kNote", "Preview uses Production rows only. Standby excluded.");
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
        description: downDesc,
        critical: false,
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

  for (const r of rowsToRender) {
    const tr = document.createElement("tr");
    if (r.is_master_standby) tr.classList.add("standbyRow");
    if (r.is_down) tr.classList.add("downRow");
    if (r.error) tr.classList.add("errorRow");

    const tdA = document.createElement("td");
    tdA.innerHTML = `<div class="rowAsset"><b>${r.asset_code}</b><small>${r.asset_name || ""}</small></div>`;
    tr.appendChild(tdA);

    const tdP = document.createElement("td");
    tdP.className = "toggle";

    const prodWrap = document.createElement("label");
    prodWrap.className = "prodToggle";
    const chkProd = document.createElement("input");
    chkProd.type = "checkbox";
    chkProd.checked = !!r.is_used;
    chkProd.disabled = r.is_master_standby;
    const prodTxt = document.createElement("span");
    prodTxt.textContent = "PROD";
    prodWrap.appendChild(chkProd);
    prodWrap.appendChild(prodTxt);

    chkProd.addEventListener("change", () => {
      r.is_used = chkProd.checked;
      if (r.is_master_standby) r.is_used = false;
      validateDailyRows();
      renderDailyTable();
      renderDailyPreview();
    });

    const downWrap = document.createElement("label");
    downWrap.className = "downToggle";
    const chkDown = document.createElement("input");
    chkDown.type = "checkbox";
    chkDown.checked = !!r.is_down;
    chkDown.disabled = r.is_master_standby || (r.is_down && r.down_lock);
    const downTxt = document.createElement("span");
    downTxt.textContent = "DOWN";
    downWrap.appendChild(chkDown);
    downWrap.appendChild(downTxt);

    const reason = document.createElement("select");
    reason.style.marginLeft = "8px";
    reason.disabled = !r.is_down || (r.is_down && r.down_lock);

    const opts = [
      "",
      "Mechanical",
      "Electrical",
      "Hydraulics",
      "Tyres/Undercarriage",
      "Waiting parts",
      "No operator",
      "Weather/Access",
    ];
    for (const o of opts) {
      const op = document.createElement("option");
      op.value = o;
      op.textContent = o === "" ? "Reason..." : o;
      if ((r.down_reason || "") === o) op.selected = true;
      reason.appendChild(op);
    }

    reason.addEventListener("change", () => {
      r.down_reason = String(reason.value || "");
    });

    downWrap.appendChild(reason);
    const unitSel = document.createElement("select");
    unitSel.style.marginLeft = "8px";
    ["hours", "km"].forEach((u) => {
      const op = document.createElement("option");
      op.value = u;
      op.textContent = u.toUpperCase();
      if (String(r.input_unit || "hours").toLowerCase() === u) op.selected = true;
      unitSel.appendChild(op);
    });
    unitSel.disabled = !!(r.is_down && r.down_lock);
    unitSel.addEventListener("change", () => {
      r.input_unit = String(unitSel.value || "hours").toLowerCase();
      validateDailyRows();
      renderDailyTable();
      renderDailyPreview();
    });
    downWrap.appendChild(unitSel);
    const resetUnitBtn = document.createElement("button");
    resetUnitBtn.type = "button";
    resetUnitBtn.className = "miniBtn";
    resetUnitBtn.style.marginLeft = "6px";
    resetUnitBtn.textContent = "↺";
    resetUnitBtn.title = "Reset unit to suggested default";
    resetUnitBtn.disabled = !!(r.is_down && r.down_lock);
    resetUnitBtn.addEventListener("click", () => {
      const suggested = String(r.suggested_input_unit || "hours").toLowerCase() === "km" ? "km" : "hours";
      r.input_unit = suggested;
      unitSel.value = suggested;
      validateDailyRows();
      renderDailyTable();
      renderDailyPreview();
    });
    downWrap.appendChild(resetUnitBtn);

    // Down hrs input (for accurate availability downtime)
    const downHrs = document.createElement("input");
    downHrs.className = "cellInput";
    downHrs.type = "number";
    downHrs.step = "0.5";
    downHrs.min = "0";
    downHrs.max = "24";
    downHrs.style.marginLeft = "8px";
    downHrs.style.width = "88px";
    downHrs.title = "Downtime hours for the day (0..Scheduled). Used for Availability.";
    downHrs.placeholder = "Down hrs";
    downHrs.disabled = !r.is_down; // editable even when locked (only DOWN flag/reason/unit are locked)
    downHrs.value = fmt(r.down_hours != null ? r.down_hours : (r.is_down ? Number(r.scheduled_hours || 0) : 0));
    downHrs.addEventListener("input", () => {
      r.down_hours = toNum(downHrs.value) ?? 0;
      validateDailyRows();
      renderDailyPreview();
    });
    downHrs.addEventListener("change", () => {
      r.down_hours = toNum(downHrs.value) ?? 0;
      validateDailyRows();
      renderDailyTable();
      renderDailyPreview();
    });
    downWrap.appendChild(downHrs);

    chkDown.addEventListener("change", () => {
      r.is_down = chkDown.checked;

      if (r.is_down) {
        if (r.opening_hours != null) r.closing_hours = r.opening_hours;
        r.hours_run = 0;
        if (r.down_hours == null) r.down_hours = Number(r.scheduled_hours || 0);
      } else {
        r.down_reason = "";
        r.down_hours = null;
      }

      reason.disabled = !r.is_down;
      downHrs.disabled = !r.is_down;

      validateDailyRows();
      renderDailyTable();
      renderDailyPreview();
    });

    tdP.appendChild(prodWrap);
    tdP.appendChild(downWrap);
    tr.appendChild(tdP);

    const tdS = document.createElement("td");
    const sIn = document.createElement("input");
    sIn.className = "cellInput";
    sIn.type = "number";
    sIn.step = "0.5";
    sIn.min = "0";
    sIn.max = "24";
    sIn.value = fmt(r.scheduled_hours);

    sIn.addEventListener("input", () => {
      r.scheduled_hours = toNum(sIn.value) ?? 0;
      validateDailyRows();
      renderDailyPreview();
    });
    sIn.addEventListener("change", () => {
      r.scheduled_hours = toNum(sIn.value) ?? 0;
      validateDailyRows();
      renderDailyTable();
      renderDailyPreview();
    });

    tdS.appendChild(sIn);
    tr.appendChild(tdS);

    const tdO = document.createElement("td");
    const oIn = document.createElement("input");
    oIn.className = "cellInput readonly";
    oIn.type = "number";
    oIn.step = "0.1";
    oIn.value = fmt(r.opening_hours);
    oIn.title = r.opening_from_date
      ? `Auto-filled from closing on ${r.opening_from_date}`
      : "Auto-filled from carry-forward closing (if available)";
    oIn.disabled = false;

    oIn.addEventListener("input", () => {
      r.opening_hours = toNum(oIn.value);
      validateDailyRows();
      renderDailyPreview();
    });
    oIn.addEventListener("change", () => {
      r.opening_hours = toNum(oIn.value);
      validateDailyRows();
      renderDailyTable();
      renderDailyPreview();
    });

    tdO.appendChild(oIn);
    tr.appendChild(tdO);

    const tdC = document.createElement("td");
    const cIn = document.createElement("input");
    cIn.className = "cellInput";
    cIn.type = "number";
    cIn.step = "0.1";
    cIn.value = fmt(r.closing_hours);
    cIn.disabled = r.is_down;

    cIn.addEventListener("input", () => {
      r.closing_hours = toNum(cIn.value);
      validateDailyRows();
      renderDailyPreview();
    });
    cIn.addEventListener("change", () => {
      r.closing_hours = toNum(cIn.value);
      validateDailyRows();
      renderDailyTable();
      renderDailyPreview();
    });

    tdC.appendChild(cIn);
    tr.appendChild(tdC);

    const tdR = document.createElement("td");
    tdR.textContent = `${fmt(r.hours_run)}${String(r.input_unit || "hours").toLowerCase() === "km" ? " km" : ""}`;
    tr.appendChild(tdR);

    const tdSt = document.createElement("td");
    if (r.is_down) {
      const lockNote = r.is_down && r.down_lock ? `<br><small class="muted">Locked until WO is repaired (status: ${r.lock_wo_status || "-"})</small>` : "";
      tdSt.innerHTML = `<span class="badge err">DOWN — ${r.down_reason || "No reason"}</span>${lockNote}`;
    }
    else if (r.error) tdSt.innerHTML = `<span class="badge err">${r.error}</span>`;
    else if (r.warning) tdSt.innerHTML = `<span class="badge warn">${r.warning}</span>`;
    else tdSt.innerHTML = `<span class="badge ok">${r.is_used ? "OK — PRODUCTION" : "OK — STANDBY"}</span>`;
    if (r.opening_from_date) {
      tdSt.innerHTML += `<br><span class="pill blue">Opening source: ${r.opening_from_date}</span>`;
    }
    tr.appendChild(tdSt);

    body.appendChild(tr);
  }
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

      is_used: ex ? !!ex.is_used : !masterStandby,
      input_unit: ex?.input_unit
        ? String(ex.input_unit).toLowerCase()
        : (String(a.category || "").toLowerCase().includes("truck") || String(a.category || "").toLowerCase().includes("vehicle") ? "km" : "hours"),
      suggested_input_unit: ex?.input_unit
        ? String(ex.input_unit).toLowerCase()
        : (String(a.category || "").toLowerCase().includes("truck") || String(a.category || "").toLowerCase().includes("vehicle") ? "km" : "hours"),
      input_unit_locked: Boolean(ex?.input_unit_locked),

      scheduled_hours: ex ? toNum(ex.scheduled_hours) : null,
      opening_hours: ex ? toNum(ex.opening_hours) : null,
      opening_from_date: null,
      closing_hours: ex ? toNum(ex.closing_hours) : null,
      hours_run: ex ? toNum(ex.hours_run) ?? 0 : 0,

      is_down: false,
      down_reason: "",
      down_hours: null,
      down_lock: false,

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
        if (typeof d?.input_unit_locked === "boolean") row.input_unit_locked = d.input_unit_locked;
        if ((row.opening_hours == null || forceOpenFromYesterday) && d.suggested_opening_hours != null) {
          row.opening_hours = Number(d.suggested_opening_hours);
          row.opening_from_date = String(d.suggested_opening_from_date || "").trim() || null;
        }
        if (row.scheduled_hours == null && d.suggested_scheduled_hours != null) row.scheduled_hours = Number(d.suggested_scheduled_hours);
      } catch {}
    }

    if (row.scheduled_hours == null) row.scheduled_hours = 0;
    row.hours_run = calcRun(row.opening_hours, row.closing_hours);

    // Apply carry-forward lock from open breakdown
    const bd = openBreakdownByAsset.get(row.asset_code);
    if (bd) {
      row.is_down = true;
      row.down_lock = true;
      row.down_reason = parseDownReasonFromDesc(bd.description);
        row.lock_wo_status = bd.primary_work_order_status || "";
      const downForDate = Number(bd.hours_down_for_date);
      row.down_hours = Number.isFinite(downForDate) && downForDate >= 0
        ? downForDate
        : Number(row.scheduled_hours || 0);
      if (row.opening_hours != null) {
        row.closing_hours = row.opening_hours;
        row.hours_run = 0;
      }
    }

    dailyRows.push(row);
  }

  validateDailyRows();
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
  const payload = {
    asset_code,
    asset_name: assetName,
    category,
    active: 1,
    is_standby: isStandby ? 1 : 0,
  };

  setStatus(`Adding contractor asset ${asset_code}...`);
  try {
    const res = await fetchJson(`${API}/api/assets`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (out) out.textContent = JSON.stringify({ ok: true, asset_code, id: res?.id || null }, null, 2);
    if (qs("caCode")) qs("caCode").value = "";
    if (qs("caName")) qs("caName").value = "";
    if (qs("caStandby")) qs("caStandby").checked = false;
    setStatus(`Contractor asset ${asset_code} added.`);
    await Promise.all([
      populateHistoryAssets().catch(() => {}),
      loadCodePickers().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
    const histSel = qs("histAsset");
    if (histSel) histSel.value = asset_code;
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Failed to add contractor asset.");
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

  if (current) {
    const exists = Array.from(sel.options).some((o) => o.value === current);
    if (exists) sel.value = current;
  }
}

function pillForType(t) {
  if (t === "breakdown") return "<span class='pill red'>BD</span>";
  if (t === "service") return "<span class='pill blue'>SV</span>";
  if (t === "get_slip") return "<span class='pill orange'>GET</span>";
  if (t === "component_slip") return "<span class='pill orange'>COMP</span>";
  if (t === "damage_report") return "<span class='pill red'>DMG</span>";
  if (t === "tyre_change") return "<span class='pill orange'>TY CHG</span>";
  if (t === "tyre_inspection") return "<span class='pill blue'>TY INSP</span>";
  if (t === "work_order") return "<span class='pill blue'>WO</span>";
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
  const asset_code = qs("histAsset")?.value;
  if (!asset_code) return alert("Select an asset.");

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
      <div class="pill blue">GET Slips: ${Number(counts.get_slips || 0)}</div>
      <div class="pill blue">Component Slips: ${Number(counts.component_slips || 0)}</div>
      <div class="pill red">Damage Reports: ${Number(counts.damage_reports || 0)}</div>
      <div class="pill orange">Tyre Changes: ${Number(counts.tyre_changes || 0)}</div>
      <div class="pill blue">Tyre Inspections: ${Number(counts.tyre_inspections || 0)}</div>
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
        : "";

    list.appendChild(item(`${pillForType(ev.type)} <b>${ev.date}</b> — ${ev.title}${wo}${extra}`));
  });

  if (!data.history?.length) list.appendChild(item("<small>No history found for this range.</small>"));

  setStatus("History loaded ✅");
}

async function archiveSelectedAsset() {
  const code = qs("histAsset")?.value;
  if (!code) return alert("Select an asset first.");

  const reason = prompt(`Archive ${code}.\nReason (optional):`, "Scrapped / Not in use");
  if (reason === null) return;

  setStatus(`Archiving ${code}...`);
  await fetchJson(`${API}/api/assets/${encodeURIComponent(code)}/archive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ archived: true, reason: String(reason || "").trim() }),
  });

  setStatus(`Archived ${code} ✅`);
  await populateHistoryAssets().catch(() => {});
  await loadDashboard().catch(() => {});
}

async function unarchiveSelectedAsset() {
  const code = qs("histAsset")?.value;
  if (!code) return alert("Select an asset first.");

  if (!confirm(`Unarchive ${code}?`)) return;

  setStatus(`Unarchiving ${code}...`);
  await fetchJson(`${API}/api/assets/${encodeURIComponent(code)}/archive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ archived: false }),
  });

  setStatus(`Unarchived ${code} ✅`);
  await populateHistoryAssets().catch(() => {});
  await loadDashboard().catch(() => {});
}

/* =========================
   INIT
========================= */

async function init() {
  await disableLegacyServiceWorkers();
  await tryInitialSession();
  initTabs();
  initSessionControls();
  initVehicleCheckTab();
  applyRoleVisibility();
  applyI18n();
  applyGlobalPageTranslation();

  const dateEl = qs("date");
  if (dateEl) dateEl.value = todayLocalYmd();

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
  qs("ironmindAskBtn")?.addEventListener("click", () =>
    askIronmindQuestion().catch((e) => setStatus("IRONMIND ask error: " + e.message))
  );
  qs("ironmindAskInput")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      askIronmindQuestion().catch((err) => setStatus("IRONMIND ask error: " + err.message));
    }
  });
  qs("saveThresholds")?.addEventListener("click", () => saveThresholdsFromUI());
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
  qs("loadLubeAnalytics")?.addEventListener("click", () =>
    loadLubeAnalytics().catch((e) => setStatus("Lube analytics error: " + e.message))
  );
  qs("createRequisition")?.addEventListener("click", () =>
    createRequisition().catch((e) => setStatus("Requisition create error: " + e.message))
  );
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
    syncFuelUnitFromAsset(qs("fuelAsset")?.value, "input").catch(() => {});
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
  qs("fuelDupOnly")?.addEventListener("change", () =>
    loadFuelBenchmark().catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
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
  qs("loadStockMonitor")?.addEventListener("click", () =>
    loadStockMonitor().catch((e) => setStatus("Stock monitor error: " + e.message))
  );
  qs("spLoad")?.addEventListener("click", () =>
    loadStockOnHandPage().catch((e) => setStatus("Stock on hand error: " + e.message))
  );
  qs("spSort")?.addEventListener("change", () =>
    loadStockOnHandPage().catch((e) => setStatus("Stock on hand error: " + e.message))
  );
  qs("spOnlyLow")?.addEventListener("change", () =>
    loadStockOnHandPage().catch((e) => setStatus("Stock on hand error: " + e.message))
  );
  qs("spExportCsv")?.addEventListener("click", exportStockOnHandCsv);
  qs("spOpenPdf")?.addEventListener("click", openStockOnHandPdf);
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
  qs("saveAdminUserBtn")?.addEventListener("click", () =>
    saveAdminUser().catch((e) => setStatus("Save user error: " + e.message))
  );
  qs("chPwdSubmit")?.addEventListener("click", () =>
    submitChangePassword().catch((e) => setStatus("Password error: " + e.message))
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
  qs("downloadFuelTemplate")?.addEventListener("click", downloadFuelCsvTemplate);
  qs("downloadStoreTemplate")?.addEventListener("click", downloadStoresCsvTemplate);
  qs("downloadFuelBaselineTemplate")?.addEventListener("click", downloadFuelBaselineCsvTemplate);

  qs("openDaily")?.addEventListener("click", openDailyPdf);
  qs("openWeekly")?.addEventListener("click", openWeeklyPdf);
  qs("openLubePdf")?.addEventListener("click", openLubePdf);
  qs("openStockMonitorPdf")?.addEventListener("click", openStockMonitorPdf);
  qs("downloadStockMonitorPdf")?.addEventListener("click", downloadStockMonitorPdf);
  qs("openOperationsPdf")?.addEventListener("click", () => openOperationsPdf(false));
  qs("downloadOperationsPdf")?.addEventListener("click", () => openOperationsPdf(true));
  qs("downloadOperationsXlsx")?.addEventListener("click", downloadOperationsXlsx);
  qs("openDailyXlsx")?.addEventListener("click", openDailyXlsx);
  qs("openGmWeeklyXlsx")?.addEventListener("click", openGmWeeklyXlsx);
  qs("downloadCostMonthlyXlsx")?.addEventListener("click", downloadCostMonthlyXlsx);
  qs("downloadMaintenanceCostByEquipmentXlsx")?.addEventListener("click", downloadMaintenanceCostByEquipmentXlsx);
  qs("openMaintenanceCostByEquipmentPdf")?.addEventListener("click", () => openMaintenanceCostByEquipmentPdf(false));
  qs("downloadMaintenanceCostByEquipmentPdf")?.addEventListener("click", () => openMaintenanceCostByEquipmentPdf(true));
  qs("downloadMaintenanceExecutivePptx")?.addEventListener("click", downloadMaintenanceExecutivePptx);
  qs("saveRainDayBtn")?.addEventListener("click", () => saveRainDay().catch((e) => setStatus("Rain day save error: " + e.message)));
  qs("removeRainDayBtn")?.addEventListener("click", () => removeRainDay().catch((e) => setStatus("Rain day remove error: " + e.message)));
  qs("loadRainDaysBtn")?.addEventListener("click", () => loadRainDays().catch((e) => setStatus("Rain day load error: " + e.message)));

  qs("makeBreakdown")?.addEventListener("click", () =>
    createBreakdown().catch((e) => setStatus("Breakdown error: " + e.message))
  );
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
  qs("applyBulkSched")?.addEventListener("click", applyBulkScheduled);
  qs("dailyDownOnly")?.addEventListener("change", () => {
    dailyShowDownOnly = !!qs("dailyDownOnly")?.checked;
    renderDailyTable();
  });

  // Net banner
  refreshNetBanner();
  window.addEventListener("offline", refreshNetBanner);

  window.addEventListener("online", async () => {
    refreshNetBanner();
    if (getQueuedHoursCount() === 0) return;
    try {
      await syncOfflineHoursQueue();
    } catch (e) {
      setStatus("Sync error: " + (e.message || e));
      refreshNetBanner();
    }
  });

  qs("syncNow")?.addEventListener("click", async () => {
    if (!navigator.onLine) return alert("Still offline.");
    try {
      await syncOfflineHoursQueue();
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

  qs("showArchived")?.addEventListener("change", () => {
    populateHistoryAssets().catch(() => {});
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

  populateHistoryAssets().catch(() => {});
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
  loadLubeStockOnHand().catch(() => {});
  loadLubeReorderAlerts().catch(() => {});
  applyDefaultLocationsToInputs();
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
  loadRequisitions().catch(() => {});
  loadOperations().catch(() => {});
  loadSiteZones().catch(() => {});
  loadSiteDailyEntries().catch(() => {});
  loadSiteTargets().catch(() => {});
  loadSiteDelays().catch(() => {});
  loadSiteDashboard().catch(() => {});
  loadDispatchTrips().catch(() => {});
  loadQualityCenter().catch(() => {});
  loadFuelBenchmark().catch(() => {});
  loadFuelSnapshots().catch(() => {});
  loadCostSettings().catch(() => {});
  loadLegalDepartments().catch(() => {});
  loadLegalDocs().catch(() => {});
  loadLegalExpiry().catch(() => {});
  if (getSessionRoles().some((r) => ["admin", "supervisor"].includes(r))) {
    loadAuditLogs().catch(() => {});
    loadApprovalRequests().catch(() => {});
  }
  loadCodePickers().catch(() => {});
  populateThresholdInputs();
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

// Keep startup reliable even if script placement changes
document.addEventListener("DOMContentLoaded", () => {
  init().catch((e) => console.error(e));
});