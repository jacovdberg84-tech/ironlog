/**
 * IRONLOG stores counter terminal (kiosk).
 *
 * People tap their name and enter a PIN. Storemen issue, receive, look up,
 * count and work the workshop's requests; technicians collect parts for
 * their own work orders. A USB/Bluetooth barcode scanner types into the page;
 * a phone can be paired as a camera scanner (store-scanner.html), its scans
 * relayed through /api/stock/terminal/scan.
 */
(function () {
  const A = window.IronlogAuth;
  const API = A.API;
  const esc = A.escapeHtml;
  const LANG_KEY = "ironlog-terminal-lang";
  const PAIR_KEY = "ironlog-terminal-scan-key";
  const IDLE_TECH_MS = 90 * 1000;
  const IDLE_STORES_MS = 10 * 60 * 1000;
  const IDLE_WARN_MS = 15 * 1000;
  const IDLE_MANUAL_MS = 20 * 60 * 1000;
  const DONE_RETURN_MS = 10 * 1000;

  // ------------------------------------------------------------- language
  const PT = {
    "Stores": "Armazém",
    "Sign out": "Sair",
    "Scanner ready: scan a part label, machine or work order": "Leitor pronto: leia uma etiqueta de peça, máquina ou ordem de trabalho",
    "📱 Phone scanner": "📱 Telemóvel como leitor",
    "Tap your name": "Toque no seu nome",
    "Then enter your PIN": "Depois introduza o seu PIN",
    "Find your name": "Procure o seu nome",
    "No one can sign in here yet. A supervisor sets a PIN for each person in User Admin.": "Ainda ninguém pode entrar aqui. O supervisor define um PIN para cada pessoa na Administração de Utilizadores.",
    "Enter your PIN": "Introduza o seu PIN",
    "Clear": "Limpar",
    "Cancel": "Cancelar",
    "Sign in": "Entrar",
    "Wrong PIN, try again": "PIN errado, tente de novo",
    "This account cannot use the stores terminal": "Esta conta não pode usar o terminal do armazém",
    "Storeman": "Armazém",
    "Technician": "Técnico",
    "Hello, {name}": "Olá, {name}",
    "Issue parts": "Entregar peças",
    "To a work order or machine": "Para uma ordem de trabalho ou máquina",
    "Collect parts": "Levantar peças",
    "For one of your jobs": "Para um dos seus trabalhos",
    "{n} open jobs": "{n} trabalhos abertos",
    "Receive delivery": "Receber entrega",
    "Book stock in from an invoice": "Dar entrada de stock de uma fatura",
    "Find stock": "Procurar stock",
    "How many and where": "Quantos e onde",
    "Workshop requests": "Pedidos da oficina",
    "{n} ready in stock": "{n} prontos em stock",
    "Count stock": "Contar stock",
    "Shelf count for one item": "Contagem de um artigo na prateleira",
    "Back": "Voltar",
    "1. Job": "1. Trabalho",
    "2. Parts": "2. Peças",
    "Work order number or machine": "Número da OT ou máquina",
    "No open jobs found": "Nenhum trabalho aberto encontrado",
    "You have no open jobs. Ask your foreman to assign you.": "Não tem trabalhos abertos. Peça ao seu encarregado para o atribuir.",
    "Change": "Mudar",
    "Issue to a machine (no work order)": "Entregar a uma máquina (sem OT)",
    "Machine code": "Código da máquina",
    "Scan or type the part": "Leia ou escreva a peça",
    "Scan a part label or search above": "Leia uma etiqueta ou procure acima",
    "in stock": "em stock",
    "Bin": "Prateleira",
    "Only {n} in stock": "Só {n} em stock",
    "Choose the job first": "Escolha primeiro o trabalho",
    "Issue {n} item(s)": "Entregar {n} artigo(s)",
    "Issue these parts?": "Entregar estas peças?",
    "Yes, issue": "Sim, entregar",
    "Issued": "Entregue",
    "left in stock": "ficam em stock",
    "Done": "Concluído",
    "Signing you out in {n} s": "A terminar a sessão em {n} s",
    "Back to the start in {n} s": "Volta ao início em {n} s",
    "Invoice / GRN number": "Número da fatura / GRN",
    "Supplier (optional)": "Fornecedor (opcional)",
    "Delivery": "Entrega",
    "Receive {n} line(s)": "Receber {n} linha(s)",
    "Enter the invoice or GRN number": "Introduza o número da fatura ou GRN",
    "This delivery may already be in stock": "Esta entrega pode já estar em stock",
    "Receive anyway": "Receber na mesma",
    "Received": "Recebido",
    "now in stock": "agora em stock",
    "Part not found? New parts are added in Stock Control on the office PC.": "Peça não encontrada? Peças novas são criadas no Controlo de Stock no PC do escritório.",
    "Type at least 2 letters": "Escreva pelo menos 2 letras",
    "Nothing found": "Nada encontrado",
    "Issue this part": "Entregar esta peça",
    "Minimum": "Mínimo",
    "No open requests": "Sem pedidos abertos",
    "Machine down": "Máquina parada",
    "In stock": "Em stock",
    "Not in stock": "Sem stock",
    "requested by {name}": "pedido por {name}",
    "Not in stock: order it in Stock Control": "Sem stock: encomende no Controlo de Stock",
    "Scan or search the item to count": "Leia ou procure o artigo a contar",
    "How many are on the shelf?": "Quantos estão na prateleira?",
    "Save count": "Guardar contagem",
    "Count matches the system": "A contagem confere com o sistema",
    "Difference of {n} sent to the supervisor for approval": "Diferença de {n} enviada ao supervisor para aprovação",
    "Quantity": "Quantidade",
    "OK": "OK",
    "Sign in first": "Entre primeiro",
    "Not found: {code}": "Não encontrado: {code}",
    "Scan a part label": "Leia uma etiqueta de peça",
    "That work order is not one of your jobs": "Essa ordem de trabalho não é sua",
    "Added {code}": "Adicionado {code}",
    "Pair a phone as the scanner": "Usar um telemóvel como leitor",
    "Scan this code with the phone's camera. Leave the scanner page open on the phone; every code it reads appears here.": "Leia este código com a câmara do telemóvel. Deixe a página do leitor aberta no telemóvel; cada código lido aparece aqui.",
    "New pairing": "Nova ligação",
    "A new pairing disconnects phones paired before.": "Uma nova ligação desliga os telemóveis ligados antes.",
    "Close": "Fechar",
    "Phone scanner paired": "Telemóvel ligado como leitor",
    "Could not reach IronLog. Check the network.": "Sem ligação ao IronLog. Verifique a rede.",
    "Tap anywhere to stay signed in": "Toque em qualquer lado para continuar",
    "No work order": "Sem OT",
    "New barcode": "Código de barras novo",
    "Which part is in this box? Choose it once; after that this barcode finds the part for everyone.": "Que peça está nesta caixa? Escolha uma vez; depois este código encontra a peça para todos.",
    "Linked: {barcode} → {code}": "Ligado: {barcode} → {code}",
    "This barcode is linked to {code}": "Este código está ligado a {code}",
    "Move it to {code}?": "Passar para {code}?",
    "Move it": "Passar",
    "Not found: {code}. Ask the storeman to link this barcode.": "Não encontrado: {code}. Peça ao fiel de armazém para ligar este código.",
    "Box barcodes": "Códigos de barras da caixa",
    "Link a box barcode": "Ligar código de barras da caixa",
    "Scan the barcode on the box now": "Leia agora o código de barras da caixa",
    "or type it here": "ou escreva aqui",
    "Remove barcode {barcode}?": "Remover o código {barcode}?",
    "Remove": "Remover",
    "Workshop library": "Biblioteca da oficina",
    "Manuals, bulletins and fixes": "Manuais, boletins e soluções",
    "Manuals": "Manuais",
    "Search: machine, model or manual": "Procurar: máquina, modelo ou manual",
    "All": "Todos",
    "Workshop Manual": "Manual de oficina",
    "Service Manual": "Manual de serviço",
    "Parts Manual": "Catálogo de peças",
    "OEM Bulletin": "Boletim do fabricante",
    "Technical Document": "Documento técnico",
    "Workshop Fix": "Solução da oficina",
    "No manuals found": "Nenhum manual encontrado",
    "Ask the manuals": "Pergunte aos manuais",
    "Ask about a fault code, torque, procedure or part. The answer points to the manual pages: always check the page before working.": "Pergunte sobre um código de avaria, binário, procedimento ou peça. A resposta indica as páginas do manual: confirme sempre a página antes de trabalhar.",
    "e.g. B30D fault code 2121": "ex. B30D código de avaria 2121",
    "Ask": "Perguntar",
    "Searching the manuals…": "A procurar nos manuais…",
    "page {n}": "página {n}",
    "Opening the manual… large manuals take a moment.": "A abrir o manual… manuais grandes demoram um pouco.",
    "Could not open the manual": "Não foi possível abrir o manual",
    "Close the manual first": "Feche primeiro o manual",
    "rev {r}": "rev {r}",
    "🖨 Label": "🖨 Etiqueta",
    "Print labels for {code}": "Imprimir etiquetas para {code}",
    "How many labels?": "Quantas etiquetas?",
    "Add to the print list": "Adicionar à lista de impressão",
    "Print here": "Imprimir aqui",
    "Added to the print list ({n} parts waiting). Print them in Stock Control → Setup → Part labels.": "Adicionado à lista ({n} peças em espera). Imprima em Controlo de Stock → Configuração → Etiquetas.",
  };
  let lang = (() => { try { return localStorage.getItem(LANG_KEY) || "en"; } catch { return "en"; } })();
  function T(s, vars) {
    let out = lang === "pt" && PT[s] ? PT[s] : s;
    if (vars) for (const [k, v] of Object.entries(vars)) out = out.split(`{${k}}`).join(String(v));
    return out;
  }
  function applyStaticText() {
    document.querySelectorAll("[data-t]").forEach((el) => { el.textContent = T(el.dataset.t); });
    document.querySelectorAll("[data-lang]").forEach((b) => b.classList.toggle("on", b.dataset.lang === lang));
    document.documentElement.lang = lang;
  }

  // ------------------------------------------------------------- state
  const $ = (id) => document.getElementById(id);
  const main = $("stMain");
  const st = {
    me: null, // { username, label, stores }
    screen: "signin",
    roster: [],
    rosterFilter: "",
    home: null,
    issue: null,
    recv: null,
    find: null,
    count: null,
    requests: null,
    lastActivity: Date.now(),
    doneTimer: null,
  };

  const api = (path, opts) => A.fetchJson(`${API}/api${path}`, opts);
  const post = (path, body) => api(path, { method: "POST", body: JSON.stringify(body) });
  const fmtQty = (n) => {
    const v = Number(n || 0);
    return Number.isInteger(v) ? String(v) : v.toFixed(2).replace(/\.?0+$/, "");
  };

  // ------------------------------------------------------------- feedback
  let audioCtx = null;
  function beep(ok = true) {
    try {
      audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
      const o = audioCtx.createOscillator();
      const g = audioCtx.createGain();
      o.frequency.value = ok ? 1250 : 220;
      o.type = ok ? "sine" : "square";
      g.gain.value = 0.12;
      o.connect(g).connect(audioCtx.destination);
      o.start();
      o.stop(audioCtx.currentTime + (ok ? 0.09 : 0.35));
    } catch { /* no audio */ }
  }
  let toastTimer = null;
  function toast(msg, kind = "") {
    const el = $("stToast");
    el.textContent = msg;
    el.className = `st-toast show ${kind}`;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { el.className = "st-toast"; }, kind === "bad" ? 4000 : 2400);
  }
  function errText(e) {
    if (e && e.status == null && /fetch|network/i.test(String(e.message))) return T("Could not reach IronLog. Check the network.");
    return String(e?.message || e);
  }

  // ------------------------------------------------------------- overlays
  function overlay(html) {
    const o = $("stOverlay");
    o.onclick = null;
    o._keys = null;
    o.innerHTML = html;
    o.classList.remove("hidden");
    return o;
  }
  function closeOverlay() {
    const o = $("stOverlay");
    o.classList.add("hidden");
    o.innerHTML = "";
    o.onclick = null;
    st.captureScan = null;
  }
  const overlayOpen = () => !$("stOverlay").classList.contains("hidden");

  /** Touch number pad. opts: { title, sub, value, decimal, pin, okText, onOk(value) } */
  function numpad(opts) {
    let value = opts.value != null ? String(opts.value) : "";
    let fresh = opts.value != null; // first key press replaces the shown value
    const o = overlay(`
      <div class="st-sheet">
        <h2>${esc(opts.title || "")}</h2>
        ${opts.sub ? `<div class="muted">${esc(opts.sub)}</div>` : ""}
        ${opts.pin ? `<div class="pin-dots" id="npDots"></div>` : `<div class="num-display" id="npShow"></div>`}
        <div class="numpad">
          ${[1, 2, 3, 4, 5, 6, 7, 8, 9].map((d) => `<button type="button" data-k="${d}">${d}</button>`).join("")}
          ${opts.decimal ? `<button type="button" data-k=".">.</button>` : `<button type="button" class="fn" data-k="clear">${esc(T("Clear"))}</button>`}
          <button type="button" data-k="0">0</button>
          <button type="button" class="fn" data-k="back">⌫</button>
        </div>
        <div class="st-error" id="npErr"></div>
        <div class="row">
          <button type="button" class="btn ghost" data-k="cancel">${esc(T("Cancel"))}</button>
          <button type="button" class="btn ok" data-k="ok" id="npOk">${esc(opts.okText || T("OK"))}</button>
        </div>
      </div>`);
    const show = () => {
      if (opts.pin) {
        const n = Math.max(4, value.length);
        $("npDots").innerHTML = Array.from({ length: n }, (_, i) => `<i class="${i < value.length ? "on" : ""}"></i>`).join("");
      } else {
        $("npShow").textContent = value || "0";
      }
      $("npOk").disabled = opts.pin ? value.length < 4 : !(Number(value) > 0 || (opts.allowZero && value !== ""));
    };
    const press = async (k) => {
      if (k === "cancel") return closeOverlay();
      if (k === "ok") {
        if ($("npOk").disabled) return;
        $("npOk").disabled = true;
        try {
          const keep = await opts.onOk(value);
          if (keep === false) { value = ""; show(); return; }
          closeOverlay();
        } catch (e) {
          $("npErr").textContent = errText(e);
          if (opts.pin) { value = ""; show(); }
          $("npOk").disabled = false;
        }
        return;
      }
      $("npErr").textContent = "";
      if (fresh && k !== "back") { value = ""; fresh = false; }
      if (k === "clear") value = "";
      else if (k === "back") { value = value.slice(0, -1); fresh = false; }
      else if (k === ".") { if (!value.includes(".")) value = (value || "0") + "."; }
      else if (value.length < (opts.pin ? 6 : 7)) value += k;
      show();
    };
    o.onclick = (e) => {
      const b = e.target.closest("[data-k]");
      if (b) press(b.dataset.k);
    };
    o._keys = (e) => {
      if (/^[0-9]$/.test(e.key)) press(e.key);
      else if (e.key === "Backspace") press("back");
      else if (e.key === "Enter") press("ok");
      else if (e.key === "Escape") press("cancel");
      else if (e.key === "." && opts.decimal) press(".");
      else return false;
      return true;
    };
    show();
  }

  function confirmSheet({ title, body, yes, yesClass = "ok", onYes }) {
    const o = overlay(`
      <div class="st-sheet wide">
        <h2>${esc(title)}</h2>
        <div>${body}</div>
        <div class="st-error" id="cfErr"></div>
        <div class="row">
          <button type="button" class="btn ghost huge" data-cf="no">${esc(T("Cancel"))}</button>
          <button type="button" class="btn ${yesClass} huge" data-cf="yes">${esc(yes)}</button>
        </div>
      </div>`);
    o.onclick = async (e) => {
      const b = e.target.closest("[data-cf]");
      if (!b) return;
      if (b.dataset.cf === "no") return closeOverlay();
      b.disabled = true;
      try {
        await onYes();
      } catch (err) {
        $("cfErr").textContent = errText(err);
        b.disabled = false;
      }
    };
  }

  // ------------------------------------------------------------- sign in / out
  async function loadRoster() {
    try {
      const data = await A.fetchJson(`${API}/api/auth/pin/roster?terminal=stores`);
      st.roster = Array.isArray(data?.technicians) ? data.technicians : [];
    } catch (e) {
      st.roster = [];
      toast(errText(e), "bad");
    }
  }

  function initials(name) {
    return String(name || "?").split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0].toUpperCase()).join("") || "?";
  }

  function renderSignIn() {
    st.screen = "signin";
    $("stWho").classList.add("hidden");
    $("stSignOut").classList.add("hidden");
    const f = st.rosterFilter.trim().toLowerCase();
    const people = st.roster.filter((p) => !f || `${p.label} ${p.username}`.toLowerCase().includes(f));
    main.innerHTML = `
      <div class="st-signin">
        <h1>${esc(T("Tap your name"))}</h1>
        <div class="muted">${esc(T("Then enter your PIN"))}</div>
        ${st.roster.length > 15 ? `<div class="st-search" style="margin-top:16px"><input id="stRosterFind" type="search" placeholder="${esc(T("Find your name"))}" value="${esc(st.rosterFilter)}" autocomplete="off" /></div>` : ""}
        <div class="st-roster">
          ${people.map((p) => `<button type="button" class="st-person" data-user="${esc(p.username)}"><span class="av">${esc(initials(p.label))}</span>${esc(p.label)}</button>`).join("")}
        </div>
        ${st.roster.length ? "" : `<div class="st-empty">${esc(T("No one can sign in here yet. A supervisor sets a PIN for each person in User Admin."))}</div>`}
      </div>`;
    const find = $("stRosterFind");
    if (find) {
      find.addEventListener("input", () => {
        st.rosterFilter = find.value;
        renderSignIn();
        const again = $("stRosterFind");
        again.focus();
        again.setSelectionRange(again.value.length, again.value.length);
      });
    }
  }

  function askPin(username) {
    const person = st.roster.find((p) => p.username === username);
    numpad({
      title: person?.label || username,
      sub: T("Enter your PIN"),
      pin: true,
      okText: T("Sign in"),
      onOk: async (pin) => {
        try {
          await A.loginPin(username, pin, false);
        } catch (e) {
          if (e.status === 401) throw new Error(T("Wrong PIN, try again"));
          if (e.status === 403) throw new Error(T("This account cannot use the stores terminal"));
          throw e;
        }
        try {
          st.home = await api("/stock/terminal/home");
        } catch (e) {
          await signOut(true);
          if (e.status === 403) throw new Error(T("This account cannot use the stores terminal"));
          throw e;
        }
        st.me = { username, label: person?.label || username, stores: Boolean(st.home.stores) };
        touch();
        renderHome();
      },
    });
  }

  async function signOut(quiet = false) {
    clearTimeout(st.doneTimer);
    try { await A.fetchJson(`${API}/api/auth/logout`, { method: "POST" }); } catch { /* already gone */ }
    A.clearSession();
    st.me = null;
    st.home = st.issue = st.recv = st.find = st.count = st.requests = st.lib = null;
    forgetManuals();
    st.rosterFilter = "";
    $("stIdle").classList.add("hidden");
    if (!quiet) {
      closeOverlay();
      renderSignIn();
    }
  }

  // ------------------------------------------------------------- idle sign-out
  function touch() {
    st.lastActivity = Date.now();
    $("stIdle").classList.add("hidden");
  }
  function idleLimit() {
    // Touches inside the PDF viewer do not reach the page: allow time to read.
    if (st.screen === "manual") return IDLE_MANUAL_MS;
    return st.me?.stores ? IDLE_STORES_MS : IDLE_TECH_MS;
  }
  setInterval(() => {
    if (!st.me) return;
    const left = idleLimit() - (Date.now() - st.lastActivity);
    const banner = $("stIdle");
    if (left <= 0) {
      signOut();
    } else if (left <= IDLE_WARN_MS) {
      banner.textContent = `${T("Signing you out in {n} s", { n: Math.ceil(left / 1000) })} · ${T("Tap anywhere to stay signed in")}`;
      banner.classList.remove("hidden");
    }
  }, 1000);
  ["pointerdown", "keydown"].forEach((ev) => document.addEventListener(ev, touch, true));

  // ------------------------------------------------------------- top bar
  function showWho() {
    $("stWho").classList.remove("hidden");
    $("stSignOut").classList.remove("hidden");
    $("stWhoName").textContent = st.me.label;
    $("stWhoRole").textContent = st.me.stores ? T("Storeman") : T("Technician");
  }

  // ------------------------------------------------------------- home
  async function renderHome() {
    clearTimeout(st.doneTimer);
    st.screen = "home";
    showWho();
    try {
      st.home = await api("/stock/terminal/home");
    } catch (e) {
      if (e.status === 401) return signOut();
    }
    const h = st.home || {};
    const tile = (go, cls, icon, title, sub, badge) => `
      <button type="button" class="st-tile ${cls}" data-go="${go}">
        ${badge ? `<span class="badge">${esc(String(badge))}</span>` : ""}
        <span class="ic">${icon}</span><b>${esc(title)}</b><span>${esc(sub)}</span>
      </button>`;
    const tiles = st.me.stores
      ? [
        tile("issue", "t-issue", "📤", T("Issue parts"), T("To a work order or machine")),
        tile("receive", "t-receive", "📥", T("Receive delivery"), T("Book stock in from an invoice")),
        tile("requests", "t-requests", "🔧", T("Workshop requests"), T("{n} ready in stock", { n: h.requests_in_stock || 0 }), h.requests_waiting || ""),
        tile("find", "t-find", "🔎", T("Find stock"), T("How many and where")),
        tile("count", "t-count", "🔢", T("Count stock"), T("Shelf count for one item")),
        tile("library", "t-library", "📚", T("Workshop library"), T("Manuals, bulletins and fixes")),
      ]
      : [
        tile("issue", "t-issue", "📤", T("Collect parts"), T("{n} open jobs", { n: h.my_jobs || 0 })),
        tile("find", "t-find", "🔎", T("Find stock"), T("How many and where")),
        tile("library", "t-library", "📚", T("Workshop library"), T("Manuals, bulletins and fixes")),
      ];
    main.innerHTML = `
      <div class="st-hello">${esc(T("Hello, {name}", { name: st.me.label }))}</div>
      <div class="st-tiles">${tiles.join("")}</div>`;
  }

  function header(title) {
    return `<div class="st-head"><button type="button" class="btn ghost back" data-go="home">← ${esc(T("Back"))}</button><h2>${esc(title)}</h2><span class="grow"></span></div>`;
  }

  // ------------------------------------------------------------- part search (shared)
  let searchSeq = 0;
  async function searchParts(q) {
    const seq = ++searchSeq;
    if (String(q || "").trim().length < 2) return { seq, rows: null };
    const data = await api(`/stock/parts/search?q=${encodeURIComponent(q.trim())}`);
    return { seq, rows: data.rows || [] };
  }
  function partResults(rows, { showStock = true } = {}) {
    if (rows == null) return `<div class="st-empty">${esc(T("Scan a part label or search above"))}</div>`;
    if (!rows.length) return `<div class="st-empty">${esc(T("Nothing found"))}<br><small>${esc(T("Part not found? New parts are added in Stock Control on the office PC."))}</small></div>`;
    return rows.map((r) => `
      <button type="button" class="st-item" data-part="${esc(r.part_code)}">
        <span class="main"><b>${esc(r.part_code)}</b><small>${esc(r.part_name || "")}${r.bin ? ` · ${esc(T("Bin"))} ${esc(r.bin)}` : ""}</small></span>
        ${showStock ? `<span class="qty ${Number(r.on_hand) > 0 ? "" : "zero"}">${fmtQty(r.on_hand)}</span>` : ""}
      </button>`).join("");
  }
  /** Wire a search box: typing searches, Enter treats the text as a scan first. */
  function wireSearch(inputId, resultsId, onPick, { showStock = true } = {}) {
    const input = $(inputId);
    const box = $(resultsId);
    if (!input || !box) return;
    let rows = null;
    let t = null;
    const run = async () => {
      try {
        const res = await searchParts(input.value);
        if (res.seq !== searchSeq) return;
        rows = res.rows;
        box.innerHTML = partResults(rows, { showStock });
      } catch (e) {
        box.innerHTML = `<div class="st-msg bad">${esc(errText(e))}</div>`;
      }
    };
    input.addEventListener("input", () => { clearTimeout(t); t = setTimeout(run, 280); });
    input.addEventListener("keydown", async (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      const code = input.value.trim();
      if (!code) return;
      const handled = await onScan(code, { quietUnknown: true });
      if (handled) { input.value = ""; box.innerHTML = partResults(null); } else run();
    });
    box.addEventListener("click", (e) => {
      const b = e.target.closest("[data-part]");
      if (!b) return;
      const r = (rows || []).find((x) => x.part_code === b.dataset.part);
      if (r) onPick(r);
    });
    box.innerHTML = partResults(null, { showStock });
  }

  // ------------------------------------------------------------- basket (issue + receive)
  function basketHtml(lines, { checkStock }) {
    if (!lines.length) return `<div class="st-empty">${esc(T("Scan a part label or search above"))}</div>`;
    return lines.map((l, i) => {
      const short = checkStock && Number(l.qty) > Number(l.on_hand || 0);
      return `
        <div class="st-line ${short ? "short" : ""}" data-line="${i}">
          <div class="main"><b>${esc(l.part_code)}</b><small>${esc(l.part_name || "")}</small>
            <small>${short ? `<span style="color:#fca5a5">${esc(T("Only {n} in stock", { n: fmtQty(l.on_hand) }))}</span>` : `${fmtQty(l.on_hand)} ${esc(T("in stock"))}${l.bin ? ` · ${esc(T("Bin"))} ${esc(l.bin)}` : ""}`}</small>
          </div>
          <div class="st-step">
            <button type="button" data-step="-1">−</button>
            <button type="button" class="n" data-qty>${fmtQty(l.qty)}</button>
            <button type="button" data-step="1">+</button>
          </div>
          <button type="button" class="rm" data-rm aria-label="Remove">✕</button>
        </div>`;
    }).join("");
  }
  function addLine(lines, part, qty = 1, extra = {}) {
    const code = String(part.part_code).toUpperCase();
    const have = lines.find((l) => String(l.part_code).toUpperCase() === code && !l.request_id && !extra.request_id);
    if (have) have.qty = Number(have.qty) + qty;
    else lines.push({ part_code: part.part_code, part_name: part.part_name, on_hand: Number(part.on_hand || 0), bin: part.bin || null, qty, ...extra });
  }
  function wireBasket(boxId, lines, rerender) {
    $(boxId).addEventListener("click", (e) => {
      const row = e.target.closest("[data-line]");
      if (!row) return;
      const l = lines[Number(row.dataset.line)];
      if (!l) return;
      if (e.target.closest("[data-rm]")) { lines.splice(Number(row.dataset.line), 1); return rerender(); }
      const step = e.target.closest("[data-step]");
      if (step) {
        l.qty = Math.max(1, Number(l.qty) + Number(step.dataset.step));
        return rerender();
      }
      if (e.target.closest("[data-qty]")) {
        numpad({
          title: l.part_code,
          sub: `${l.part_name || ""} · ${T("Quantity")}`,
          value: fmtQty(l.qty),
          decimal: true,
          onOk: (v) => { l.qty = Number(v); rerender(); },
        });
      }
    });
  }

  // ------------------------------------------------------------- issue
  function startIssue({ job = null, lines = [] } = {}) {
    st.issue = { job, lines, woRows: null, woQuery: "" };
    renderIssue();
  }

  function jobCard(job) {
    const title = job.type === "wo" ? `WO #${job.id} · ${job.asset_code || ""}` : `${job.asset_code} · ${T("No work order")}`;
    return `<div class="st-job"><div class="grow"><b>${esc(title)}</b><span class="muted">${esc(job.label || job.asset_name || "")}</span></div>
      <button type="button" class="btn ghost sm" data-job-change>${esc(T("Change"))}</button></div>`;
  }

  function renderIssue() {
    st.screen = "issue";
    const s = st.issue;
    const title = st.me.stores ? T("Issue parts") : T("Collect parts");
    main.innerHTML = `
      ${header(title)}
      <div class="st-cols">
        <section class="st-panel">
          <h3>${esc(T("1. Job"))}</h3>
          <div id="isJob"></div>
        </section>
        <section class="st-panel">
          <h3>${esc(T("2. Parts"))}</h3>
          <div class="st-search"><input id="isFind" type="search" placeholder="${esc(T("Scan or type the part"))}" autocomplete="off" /></div>
          <div class="st-list" id="isResults"></div>
          <div style="height:14px"></div>
          <div class="st-basket" id="isBasket"></div>
          <div id="isMsg"></div>
          <button type="button" class="btn ok huge st-confirm" id="isGo"></button>
        </section>
      </div>`;
    renderIssueJob();
    renderIssueBasket();
    wireSearch("isFind", "isResults", (r) => { addLine(s.lines, r); renderIssueBasket(); toast(T("Added {code}", { code: r.part_code }), "ok"); });
    wireBasket("isBasket", s.lines, renderIssueBasket);
    $("isGo").addEventListener("click", confirmIssue);
  }

  function renderIssueJob() {
    const s = st.issue;
    const box = $("isJob");
    if (!box) return;
    if (s.job) {
      box.innerHTML = jobCard(s.job);
      box.querySelector("[data-job-change]").onclick = () => { s.job = null; renderIssueJob(); renderIssueBasket(); };
      return;
    }
    box.innerHTML = `
      <div class="st-search"><input id="isWoFind" type="search" placeholder="${esc(T("Work order number or machine"))}" value="${esc(s.woQuery)}" autocomplete="off" /></div>
      <div class="st-list" id="isWoList"><div class="st-empty">…</div></div>
      ${st.me.stores ? `<div style="height:12px"></div><button type="button" class="btn ghost" style="width:100%" id="isToMachine">${esc(T("Issue to a machine (no work order)"))}</button>` : ""}`;
    const input = $("isWoFind");
    let t = null;
    input.addEventListener("input", () => { s.woQuery = input.value; clearTimeout(t); t = setTimeout(loadWos, 280); });
    input.addEventListener("keydown", async (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      if (await onScan(input.value.trim(), { quietUnknown: true })) { input.value = ""; s.woQuery = ""; }
      else loadWos();
    });
    $("isWoList").addEventListener("click", (e) => {
      const b = e.target.closest("[data-wo]");
      if (!b) return;
      const w = (s.woRows || []).find((x) => String(x.id) === b.dataset.wo);
      if (w) setIssueJob({ type: "wo", id: w.id, asset_code: w.asset_code, label: w.job || w.asset_name });
    });
    $("isToMachine")?.addEventListener("click", askMachine);
    loadWos();
  }

  async function loadWos() {
    const s = st.issue;
    const list = $("isWoList");
    if (!s || !list) return;
    try {
      const data = await api(`/stock/terminal/work-orders?q=${encodeURIComponent(s.woQuery || "")}`);
      s.woRows = data.rows || [];
    } catch (e) {
      list.innerHTML = `<div class="st-msg bad">${esc(errText(e))}</div>`;
      return;
    }
    if (!$("isWoList")) return;
    $("isWoList").innerHTML = s.woRows.length
      ? s.woRows.map((w) => `
        <button type="button" class="st-item" data-wo="${w.id}">
          <span class="main"><b>WO #${w.id} · ${esc(w.asset_code)}</b><small>${esc(w.job || w.asset_name || "")}</small>${st.me.stores && w.lead ? `<small>${esc(w.lead)}</small>` : ""}</span>
        </button>`).join("")
      : `<div class="st-empty">${esc(st.me.stores || s.woQuery ? T("No open jobs found") : T("You have no open jobs. Ask your foreman to assign you."))}</div>`;
  }

  function setIssueJob(job) {
    st.issue.job = job;
    renderIssueJob();
    renderIssueBasket();
  }

  function askMachine() {
    const o = overlay(`
      <div class="st-sheet">
        <h2>${esc(T("Issue to a machine (no work order)"))}</h2>
        <label class="st-field"><span>${esc(T("Machine code"))}</span><input id="mcCode" type="text" autocomplete="off" /></label>
        <div class="st-error" id="mcErr"></div>
        <div class="row">
          <button type="button" class="btn ghost" data-mc="no">${esc(T("Cancel"))}</button>
          <button type="button" class="btn ok" data-mc="yes">${esc(T("OK"))}</button>
        </div>
      </div>`);
    const go = async () => {
      const code = $("mcCode").value.trim();
      if (!code) return;
      try {
        const hit = await api(`/stock/terminal/lookup?code=${encodeURIComponent(code)}`);
        if (hit.kind !== "asset") { $("mcErr").textContent = T("Not found: {code}", { code }); return; }
        closeOverlay();
        setIssueJob({ type: "asset", asset_code: hit.asset_code, asset_name: hit.asset_name });
      } catch (e) {
        $("mcErr").textContent = errText(e);
      }
    };
    o.onclick = (e) => {
      const b = e.target.closest("[data-mc]");
      if (!b) return;
      if (b.dataset.mc === "no") closeOverlay(); else go();
    };
    $("mcCode").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); go(); } });
    $("mcCode").focus();
  }

  function issueProblem() {
    const s = st.issue;
    if (!s.job) return T("Choose the job first");
    const need = new Map();
    for (const l of s.lines) need.set(l.part_code, (need.get(l.part_code) || 0) + Number(l.qty));
    for (const l of s.lines) if (need.get(l.part_code) > Number(l.on_hand || 0)) return `${l.part_code}: ${T("Only {n} in stock", { n: fmtQty(l.on_hand) })}`;
    return "";
  }

  function renderIssueBasket() {
    const s = st.issue;
    if (!$("isBasket")) return;
    $("isBasket").innerHTML = basketHtml(s.lines, { checkStock: true });
    const items = s.lines.reduce((n, l) => n + Number(l.qty || 0), 0);
    const problem = s.lines.length ? issueProblem() : "";
    $("isMsg").innerHTML = problem ? `<div class="st-msg ${s.job ? "bad" : "warn"}">${esc(problem)}</div>` : "";
    const go = $("isGo");
    go.textContent = T("Issue {n} item(s)", { n: fmtQty(items) });
    go.disabled = !s.lines.length || Boolean(problem);
  }

  function confirmIssue() {
    const s = st.issue;
    if (issueProblem() || !s.lines.length) return;
    const jobTitle = s.job.type === "wo" ? `WO #${s.job.id} · ${s.job.asset_code}` : s.job.asset_code;
    confirmSheet({
      title: T("Issue these parts?"),
      body: `<p class="muted" style="font-size:1.2rem">${esc(jobTitle)}</p>
        <ul class="st-list" style="list-style:none;padding:0">${s.lines.map((l) => `<li class="st-line"><div class="main"><b>${esc(l.part_code)}</b><small>${esc(l.part_name || "")}</small></div><b style="font-size:1.5rem">× ${fmtQty(l.qty)}</b></li>`).join("")}</ul>`,
      yes: T("Yes, issue"),
      onYes: async () => {
        const body = {
          lines: s.lines.map((l) => ({ part_code: l.part_code, quantity: Number(l.qty), request_id: l.request_id || undefined })),
        };
        if (s.job.type === "wo") body.work_order_id = s.job.id;
        else body.asset_code = s.job.asset_code;
        const res = await post("/stock/terminal/issue", body);
        closeOverlay();
        beep(true);
        renderDone({
          title: T("Issued"),
          sub: res.work_order_id ? `WO #${res.work_order_id} · ${res.asset_code || ""}` : res.asset_code || "",
          lines: res.lines.map((l) => [`${l.part_code} × ${fmtQty(l.quantity)}`, `${fmtQty(l.on_hand_after)} ${T("left in stock")}`]),
        });
      },
    });
  }

  // ------------------------------------------------------------- done
  function renderDone({ title, sub, lines }) {
    st.screen = "done";
    const techOut = !st.me.stores;
    const until = Date.now() + DONE_RETURN_MS;
    main.innerHTML = `
      <div class="st-done">
        <div class="big">✅</div>
        <h2>${esc(title)}</h2>
        <div class="muted" style="font-size:1.2rem">${esc(sub || "")}</div>
        <ul>${(lines || []).map(([a, b]) => `<li><b>${esc(a)}</b><span class="muted">${esc(b)}</span></li>`).join("")}</ul>
        <button type="button" class="btn primary huge" style="max-width:520px" id="doneBtn">${esc(T("Done"))}</button>
        <p class="muted" id="doneCount"></p>
      </div>`;
    const finish = () => (techOut ? signOut() : renderHome());
    $("doneBtn").onclick = finish;
    const tick = () => {
      const left = Math.ceil((until - Date.now()) / 1000);
      if (st.screen !== "done") return;
      if (left <= 0) return finish();
      $("doneCount").textContent = techOut ? T("Signing you out in {n} s", { n: left }) : T("Back to the start in {n} s", { n: left });
      st.doneTimer = setTimeout(tick, 500);
    };
    tick();
  }

  // ------------------------------------------------------------- receive
  function startReceive() {
    st.recv = { reference: "", supplier: "", lines: [] };
    renderReceive();
  }

  function renderReceive() {
    st.screen = "receive";
    const s = st.recv;
    main.innerHTML = `
      ${header(T("Receive delivery"))}
      <div class="st-cols">
        <section class="st-panel">
          <h3>${esc(T("Delivery"))}</h3>
          <label class="st-field"><span>${esc(T("Invoice / GRN number"))}</span><input id="rcRef" type="text" autocomplete="off" value="${esc(s.reference)}" /></label>
          <label class="st-field"><span>${esc(T("Supplier (optional)"))}</span><input id="rcSup" type="text" autocomplete="off" value="${esc(s.supplier)}" /></label>
          <p class="muted">${esc(T("Part not found? New parts are added in Stock Control on the office PC."))}</p>
        </section>
        <section class="st-panel">
          <h3>${esc(T("2. Parts"))}</h3>
          <div class="st-search"><input id="rcFind" type="search" placeholder="${esc(T("Scan or type the part"))}" autocomplete="off" /></div>
          <div class="st-list" id="rcResults"></div>
          <div style="height:14px"></div>
          <div class="st-basket" id="rcBasket"></div>
          <div id="rcMsg"></div>
          <button type="button" class="btn ok huge st-confirm" id="rcGo"></button>
        </section>
      </div>`;
    $("rcRef").addEventListener("input", (e) => { s.reference = e.target.value; renderReceiveBasket(); });
    $("rcSup").addEventListener("input", (e) => { s.supplier = e.target.value; });
    wireSearch("rcFind", "rcResults", (r) => { addLine(s.lines, r); renderReceiveBasket(); toast(T("Added {code}", { code: r.part_code }), "ok"); });
    wireBasket("rcBasket", s.lines, renderReceiveBasket);
    $("rcGo").addEventListener("click", () => submitReceive(false));
    renderReceiveBasket();
  }

  function renderReceiveBasket() {
    const s = st.recv;
    if (!$("rcBasket")) return;
    $("rcBasket").innerHTML = basketHtml(s.lines, { checkStock: false });
    $("rcMsg").innerHTML = s.lines.length && !s.reference.trim() ? `<div class="st-msg warn">${esc(T("Enter the invoice or GRN number"))}</div>` : "";
    $("rcGo").textContent = T("Receive {n} line(s)", { n: s.lines.length });
    $("rcGo").disabled = !s.lines.length || !s.reference.trim();
  }

  async function submitReceive(confirmDuplicates) {
    const s = st.recv;
    const body = {
      reference: s.reference.trim(),
      supplier: s.supplier.trim() || undefined,
      lines: s.lines.map((l) => ({ part_code: l.part_code, quantity: Number(l.qty) })),
      confirm_duplicates: confirmDuplicates || undefined,
    };
    $("rcGo").disabled = true;
    try {
      // Plain fetch: a 409 carries the possible duplicates to show.
      const resp = await fetch(`${API}/api/stock/deliveries`, { method: "POST", headers: A.authHeaders({ "Content-Type": "application/json" }), body: JSON.stringify(body) });
      const res = await resp.json().catch(() => ({}));
      if (!resp.ok) throw Object.assign(new Error(res.error || `Request failed (${resp.status})`), { status: resp.status, data: res });
      closeOverlay();
      beep(true);
      renderDone({
        title: T("Received"),
        sub: [res.reference, res.supplier].filter(Boolean).join(" · "),
        lines: res.lines.map((l) => [`${l.part_code} × ${fmtQty(l.quantity)}`, `${fmtQty(l.on_hand_after)} ${T("now in stock")}`]),
      });
    } catch (e) {
      if (e.status === 409 && !confirmDuplicates) {
        const dups = e.data?.duplicates || [];
        confirmSheet({
          title: T("This delivery may already be in stock"),
          body: dups.length ? `<ul>${dups.map((d) => `<li>${esc(d.text)}</li>`).join("")}</ul>` : `<p>${esc(errText(e))}</p>`,
          yes: T("Receive anyway"),
          yesClass: "bad",
          onYes: () => submitReceive(true),
        });
      } else {
        beep(false);
        if (overlayOpen()) throw e;
        $("rcMsg").innerHTML = `<div class="st-msg bad">${esc(errText(e))}</div>`;
      }
    } finally {
      if ($("rcGo")) renderReceiveBasket();
    }
  }

  // ------------------------------------------------------------- find stock
  function startFind(part = null) {
    st.find = { part };
    renderFind();
  }

  function renderFind() {
    st.screen = "find";
    const s = st.find;
    main.innerHTML = `
      ${header(T("Find stock"))}
      <div class="st-cols">
        <section class="st-panel">
          <div class="st-search"><input id="fsFind" type="search" placeholder="${esc(T("Scan or type the part"))}" autocomplete="off" /></div>
          <div class="st-list" id="fsResults"></div>
        </section>
        <section class="st-panel" id="fsPart"></section>
      </div>`;
    wireSearch("fsFind", "fsResults", (r) => { s.part = r; renderFindPart(); });
    renderFindPart();
  }

  function renderFindPart() {
    const box = $("fsPart");
    const p = st.find?.part;
    if (!box) return;
    if (!p) { box.innerHTML = `<div class="st-empty">${esc(T("Scan a part label or search above"))}</div>`; return; }
    box.innerHTML = `
      <h2 style="margin:0">${esc(p.part_code)}</h2>
      <div class="muted" style="font-size:1.15rem">${esc(p.part_name || "")}</div>
      <div class="st-stock">
        <div class="num" style="color:${Number(p.on_hand) > 0 ? "inherit" : "#fca5a5"}">${fmtQty(p.on_hand)}</div>
        <div class="muted">${esc(T("in stock"))}</div>
        ${p.bin ? `<div class="bin">${esc(T("Bin"))} ${esc(p.bin)}</div>` : ""}
        ${Number(p.min_stock) > 0 ? `<p class="muted">${esc(T("Minimum"))}: ${fmtQty(p.min_stock)}</p>` : ""}
      </div>
      ${Number(p.on_hand) > 0 ? `<button type="button" class="btn primary huge" id="fsIssue">${esc(st.me.stores ? T("Issue this part") : T("Collect parts"))}</button>` : ""}
      ${st.me.stores ? `<button type="button" class="btn ghost" style="width:100%;margin-top:12px" id="fsLabel">${esc(T("🖨 Label"))}</button><div class="st-barcodes" id="fsBarcodes"></div>` : ""}`;
    $("fsIssue")?.addEventListener("click", () => {
      const lines = [];
      addLine(lines, p);
      startIssue({ lines });
    });
    if (st.me.stores) renderPartBarcodes(p);
    $("fsLabel")?.addEventListener("click", () => labelSheet(p));
  }

  /** IronLog's own label for a part: add to the office print list, or print from here. */
  function labelSheet(part) {
    let copies = 1;
    const o = overlay(`
      <div class="st-sheet">
        <h2>${esc(T("Print labels for {code}", { code: part.part_code }))}</h2>
        <p class="muted">${esc(part.part_name || "")}${part.bin ? ` · ${esc(T("Bin"))} ${esc(part.bin)}` : ""}</p>
        <p style="font-weight:600">${esc(T("How many labels?"))}</p>
        <div class="st-step" style="justify-content:center">
          <button type="button" data-lb="minus">−</button><span class="n" id="lbN">1</span><button type="button" data-lb="plus">+</button>
        </div>
        <div class="st-error" id="lbErr"></div>
        <div class="row"><button type="button" class="btn primary" data-lb="queue">${esc(T("Add to the print list"))}</button></div>
        <div class="row">
          <button type="button" class="btn ghost" data-lb="cancel">${esc(T("Cancel"))}</button>
          ${window.IronlogPartLabels ? `<button type="button" class="btn ghost" data-lb="print">${esc(T("Print here"))}</button>` : ""}
        </div>
      </div>`);
    o.onclick = async (e) => {
      const b = e.target.closest("[data-lb]");
      if (!b) return;
      const k = b.dataset.lb;
      if (k === "cancel") return closeOverlay();
      if (k === "minus" || k === "plus") {
        copies = Math.max(1, Math.min(100, copies + (k === "plus" ? 1 : -1)));
        $("lbN").textContent = String(copies);
        return;
      }
      if (k === "print") {
        try {
          window.IronlogPartLabels.print([{ ...part, copies }], { site: A.getSessionSite() });
          closeOverlay();
        } catch (err) { $("lbErr").textContent = errText(err); }
        return;
      }
      b.disabled = true;
      try {
        const res = await post("/stock/labels/queue", { part_code: part.part_code, copies });
        closeOverlay();
        beep(true);
        toast(T("Added to the print list ({n} parts waiting). Print them in Stock Control → Setup → Part labels.", { n: res.in_list }), "ok");
      } catch (err) {
        $("lbErr").textContent = errText(err);
        b.disabled = false;
      }
    };
  }

  // ------------------------------------------------------------- box barcodes
  async function renderPartBarcodes(p) {
    const box = $("fsBarcodes");
    if (!box) return;
    let codes = [];
    try {
      codes = (await api(`/stock/terminal/lookup?code=${encodeURIComponent(`PART:${p.part_code}`)}`)).barcodes || [];
    } catch { /* show the button anyway */ }
    if (st.find?.part !== p || !$("fsBarcodes")) return;
    box.innerHTML = `
      <h3>${esc(T("Box barcodes"))}</h3>
      <div class="st-chips">${codes.map((c) => `<button type="button" class="st-chip" data-unlink="${esc(c)}">${esc(c)} <span aria-hidden="true">✕</span></button>`).join("")}</div>
      <button type="button" class="btn ghost" style="width:100%" id="fsLink">＋ ${esc(T("Link a box barcode"))}</button>`;
    $("fsLink").onclick = () => captureBarcodeFor(p);
    box.querySelectorAll("[data-unlink]").forEach((b) => {
      b.onclick = () => confirmSheet({
        title: T("Remove barcode {barcode}?", { barcode: b.dataset.unlink }),
        body: `<p class="muted">${esc(p.part_code)} · ${esc(p.part_name || "")}</p>`,
        yes: T("Remove"),
        yesClass: "bad",
        onYes: async () => {
          await api(`/stock/terminal/barcodes/${encodeURIComponent(b.dataset.unlink)}`, { method: "DELETE" });
          closeOverlay();
          renderPartBarcodes(p);
        },
      });
    });
  }

  /** Save a barcode → part link, asking before moving a barcode off another part. */
  async function saveBarcodeLink(barcode, part, replace = false) {
    const resp = await fetch(`${API}/api/stock/terminal/barcodes`, {
      method: "POST",
      headers: A.authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ barcode, part_code: part.part_code, replace }),
    });
    const res = await resp.json().catch(() => ({}));
    if (resp.status === 409 && res.linked_to) {
      return new Promise((resolve, reject) => {
        confirmSheet({
          title: T("This barcode is linked to {code}", { code: res.linked_to.part_code }),
          body: `<p class="muted">${esc(res.linked_to.part_name || "")}</p><p style="font-size:1.2rem">${esc(T("Move it to {code}?", { code: part.part_code }))}</p>`,
          yes: T("Move it"),
          onYes: async () => { resolve(await saveBarcodeLink(barcode, part, true)); },
        });
        // Cancel closes the sheet without resolving: treat it as not linked.
        const o = $("stOverlay");
        const cancel = o.querySelector('[data-cf="no"]');
        cancel?.addEventListener("click", () => reject(Object.assign(new Error("cancelled"), { cancelled: true })), { once: true });
      });
    }
    if (!resp.ok) throw Object.assign(new Error(res.error || `Request failed (${resp.status})`), { status: resp.status });
    return res;
  }

  /** An unknown scan at the stores: ask which part the box holds, then carry on with the scan. */
  function linkUnknownBarcode(barcode) {
    overlay(`
      <div class="st-sheet wide">
        <h2>${esc(T("New barcode"))}: ${esc(barcode)}</h2>
        <p class="muted">${esc(T("Which part is in this box? Choose it once; after that this barcode finds the part for everyone."))}</p>
        <div class="st-search"><input id="lkFind" type="search" placeholder="${esc(T("Scan or type the part"))}" autocomplete="off" /></div>
        <div class="st-list" id="lkResults" style="max-height:44vh;overflow-y:auto"></div>
        <div class="st-error" id="lkErr"></div>
        <div class="row"><button type="button" class="btn ghost" id="lkCancel">${esc(T("Cancel"))}</button></div>
      </div>`);
    $("lkCancel").onclick = closeOverlay;
    wireSearch("lkFind", "lkResults", async (part) => {
      try {
        await saveBarcodeLink(barcode, part);
        closeOverlay();
        beep(true);
        toast(T("Linked: {barcode} → {code}", { barcode, code: part.part_code }), "ok");
        await onScan(barcode); // continue: add to the basket, show stock …
      } catch (e) {
        if (e.cancelled) { setTimeout(() => linkUnknownBarcode(barcode), 0); return; }
        $("lkErr") && ($("lkErr").textContent = errText(e));
      }
    });
    setTimeout(() => $("lkFind")?.focus(), 50);
  }

  /** From Find stock: the next scan (scanner, phone or typed) becomes this part's barcode. */
  function captureBarcodeFor(part) {
    const done = () => { st.captureScan = null; };
    const link = async (barcode) => {
      done();
      try {
        await saveBarcodeLink(barcode, part);
        closeOverlay();
        beep(true);
        toast(T("Linked: {barcode} → {code}", { barcode: String(barcode).toUpperCase(), code: part.part_code }), "ok");
      } catch (e) {
        if (!e.cancelled) toast(errText(e), "bad");
        closeOverlay();
      }
      renderPartBarcodes(part);
    };
    const o = overlay(`
      <div class="st-sheet">
        <h2>${esc(T("Scan the barcode on the box now"))}</h2>
        <p class="muted">${esc(part.part_code)} · ${esc(part.part_name || "")}</p>
        <label class="st-field"><span>${esc(T("or type it here"))}</span><input id="cbCode" type="text" autocomplete="off" /></label>
        <div class="row">
          <button type="button" class="btn ghost" data-cb="no">${esc(T("Cancel"))}</button>
          <button type="button" class="btn ok" data-cb="yes">${esc(T("OK"))}</button>
        </div>
      </div>`);
    st.captureScan = (code) => link(code);
    const typed = () => { const v = $("cbCode").value.trim(); if (v) link(v); };
    o.onclick = (e) => {
      const b = e.target.closest("[data-cb]");
      if (!b) return;
      if (b.dataset.cb === "no") { done(); closeOverlay(); } else typed();
    };
    $("cbCode").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); typed(); } });
    $("cbCode").focus();
  }

  /** A part from a lookup (scan) in the shape the screens use. */
  function partFromHit(hit) {
    return { part_code: hit.part_code, part_name: hit.part_name, on_hand: Number(hit.on_hand || 0), bin: hit.bin || null, min_stock: Number(hit.min_stock || 0) };
  }

  // ------------------------------------------------------------- workshop requests
  async function renderRequests() {
    st.screen = "requests";
    main.innerHTML = `${header(T("Workshop requests"))}<div class="st-list" id="rqList"><div class="st-empty">…</div></div>`;
    try {
      st.requests = (await api("/stock/terminal/requests")).rows || [];
    } catch (e) {
      $("rqList").innerHTML = `<div class="st-msg bad">${esc(errText(e))}</div>`;
      return;
    }
    const rows = st.requests.slice().sort((a, b) => Number(b.machine_down) - Number(a.machine_down) || Number(b.in_stock) - Number(a.in_stock));
    $("rqList").innerHTML = rows.length
      ? rows.map((r) => `
        <button type="button" class="st-item" data-req="${r.id}">
          <span class="main">
            <b>${esc(r.asset_code || "")}${r.work_order_id ? ` · WO #${r.work_order_id}` : ""}</b>
            <small>${esc(r.part_code || "")} ${esc(r.part_name || "")}</small>
            <small>${esc(T("requested by {name}", { name: r.requested_by || "?" }))}${r.notes ? ` · ${esc(r.notes)}` : ""}</small>
          </span>
          ${r.machine_down ? `<span class="tag down">${esc(T("Machine down"))}</span>` : ""}
          <span class="tag ${r.in_stock ? "ok" : ""}">${esc(r.in_stock ? T("In stock") : T("Not in stock"))}</span>
          <span class="qty">× ${fmtQty(r.qty)}</span>
        </button>`).join("")
      : `<div class="st-empty">${esc(T("No open requests"))}</div>`;
    $("rqList").onclick = async (e) => {
      const b = e.target.closest("[data-req]");
      if (!b) return;
      const r = rows.find((x) => String(x.id) === b.dataset.req);
      if (!r) return;
      if (!r.in_stock) { toast(T("Not in stock: order it in Stock Control"), "bad"); return; }
      const job = r.work_order_id
        ? { type: "wo", id: r.work_order_id, asset_code: r.asset_code, label: r.asset_name }
        : { type: "asset", asset_code: r.asset_code, asset_name: r.asset_name };
      const lines = [];
      try {
        const hit = await api(`/stock/terminal/lookup?code=${encodeURIComponent(`PART:${r.part_code}`)}`);
        if (hit.kind === "part") lines.push({ ...partFromHit(hit), qty: Number(r.qty) || 1, request_id: r.id });
      } catch { /* the storeman can still add the part by hand */ }
      startIssue({ job, lines });
    };
  }

  // ------------------------------------------------------------- count
  function startCount() {
    st.count = { part: null };
    renderCount();
  }

  function renderCount() {
    st.screen = "count";
    main.innerHTML = `
      ${header(T("Count stock"))}
      <div class="st-cols">
        <section class="st-panel">
          <div class="st-search"><input id="ccFind" type="search" placeholder="${esc(T("Scan or type the part"))}" autocomplete="off" /></div>
          <div class="st-list" id="ccResults"></div>
        </section>
        <section class="st-panel" id="ccPart"></section>
      </div>`;
    // A blind count: the system quantity is not shown while counting.
    wireSearch("ccFind", "ccResults", (r) => countPart(r), { showStock: false });
    renderCountPart();
  }

  function renderCountPart() {
    const box = $("ccPart");
    const p = st.count?.part;
    if (!box) return;
    box.innerHTML = p
      ? `<h2 style="margin:0">${esc(p.part_code)}</h2><div class="muted" style="font-size:1.15rem">${esc(p.part_name || "")}</div>
         <div style="height:16px"></div><button type="button" class="btn primary huge" id="ccAgain">${esc(T("How many are on the shelf?"))}</button>`
      : `<div class="st-empty">${esc(T("Scan or search the item to count"))}</div>`;
    $("ccAgain")?.addEventListener("click", () => countPart(p));
  }

  function countPart(part) {
    st.count.part = part;
    renderCountPart();
    numpad({
      title: part.part_code,
      sub: `${part.part_name || ""} · ${T("How many are on the shelf?")}`,
      decimal: true,
      allowZero: true,
      okText: T("Save count"),
      onOk: async (v) => {
        const res = await post("/stock/cycle-count", { part_code: part.part_code, counted_qty: Number(v || 0), reason: "stores_terminal" });
        beep(true);
        toast(res.no_change ? T("Count matches the system") : T("Difference of {n} sent to the supervisor for approval", { n: fmtQty(res.adjustment_qty ?? (Number(v) - Number(res.on_hand || 0))) }), "ok");
        st.count.part = null;
        renderCountPart();
      },
    });
  }

  // ------------------------------------------------------------- workshop library
  const DOC_TYPES = ["Workshop Manual", "Service Manual", "Parts Manual", "OEM Bulletin", "Technical Document", "Workshop Fix"];
  const pdfCache = new Map(); // document id → blob URL of the last few manuals opened

  function forgetManuals() {
    for (const url of pdfCache.values()) URL.revokeObjectURL(url);
    pdfCache.clear();
  }

  function startLibrary() {
    st.lib = st.lib || { q: "", type: "", docs: null, question: "", answer: null };
    renderLibrary();
  }

  function renderLibrary() {
    st.screen = "library";
    const s = st.lib;
    main.innerHTML = `
      ${header(T("Workshop library"))}
      <div class="st-cols">
        <section class="st-panel">
          <h3>${esc(T("Manuals"))}</h3>
          <div class="st-search"><input id="lbQ" type="search" placeholder="${esc(T("Search: machine, model or manual"))}" value="${esc(s.q)}" autocomplete="off" /></div>
          <div class="st-chips" id="lbTypes">
            ${["", ...DOC_TYPES].map((t) => `<button type="button" class="st-chip ${s.type === t ? "on" : ""}" data-type="${esc(t)}">${esc(t ? T(t) : T("All"))}</button>`).join("")}
          </div>
          <div class="st-list" id="lbDocs"><div class="st-empty">…</div></div>
        </section>
        <section class="st-panel">
          <h3>${esc(T("Ask the manuals"))}</h3>
          <p class="muted">${esc(T("Ask about a fault code, torque, procedure or part. The answer points to the manual pages: always check the page before working."))}</p>
          <div class="st-search"><input id="lbAsk" type="search" placeholder="${esc(T("e.g. B30D fault code 2121"))}" value="${esc(s.question)}" autocomplete="off" /><button type="button" class="btn primary" id="lbAskGo">${esc(T("Ask"))}</button></div>
          <div id="lbAnswer"></div>
        </section>
      </div>`;
    let t = null;
    $("lbQ").addEventListener("input", (e) => { s.q = e.target.value; clearTimeout(t); t = setTimeout(loadLibrary, 280); });
    $("lbTypes").addEventListener("click", (e) => {
      const b = e.target.closest("[data-type]");
      if (!b) return;
      s.type = b.dataset.type;
      $("lbTypes").querySelectorAll("[data-type]").forEach((x) => x.classList.toggle("on", x === b));
      renderLibraryDocs();
    });
    $("lbDocs").addEventListener("click", (e) => {
      const b = e.target.closest("[data-doc]");
      const d = b && (s.docs || []).find((x) => x.id === b.dataset.doc);
      if (d) openManual(d, 1);
    });
    $("lbAsk").addEventListener("input", (e) => { s.question = e.target.value; });
    $("lbAsk").addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); askManuals(); } });
    $("lbAskGo").addEventListener("click", askManuals);
    $("lbAnswer").addEventListener("click", (e) => {
      const b = e.target.closest("[data-src]");
      const src = b && (s.answer?.sources || [])[Number(b.dataset.src)];
      if (!src) return;
      const doc = (s.docs || []).find((x) => x.id === src.document_id) || { id: src.document_id, title: src.title };
      openManual(doc, src.page);
    });
    renderAnswer();
    loadLibrary();
  }

  async function loadLibrary() {
    const s = st.lib;
    try {
      const data = await api(`/workshop/documents?q=${encodeURIComponent(s.q || "")}`);
      s.docs = data.documents || [];
    } catch (e) {
      if ($("lbDocs")) $("lbDocs").innerHTML = `<div class="st-msg bad">${esc(errText(e))}</div>`;
      return;
    }
    renderLibraryDocs();
  }

  function renderLibraryDocs() {
    const s = st.lib;
    const box = $("lbDocs");
    if (!box || !s.docs) return;
    const rows = s.docs.filter((d) => !s.type || d.doc_type === s.type);
    box.innerHTML = rows.length
      ? rows.map((d) => `
        <button type="button" class="st-item" data-doc="${esc(d.id)}">
          <span class="ic-doc">📄</span>
          <span class="main"><b>${esc(d.title)}</b>
            <small>${esc(T(d.doc_type))}${d.manufacturer ? ` · ${esc(d.manufacturer)}` : ""}${d.model ? ` · ${esc(d.model)}` : ""}${d.revision ? ` · ${esc(T("rev {r}", { r: d.revision }))}` : ""}</small>
            ${d.applicability ? `<small>${esc(d.applicability)}</small>` : ""}
          </span>
        </button>`).join("")
      : `<div class="st-empty">${esc(T("No manuals found"))}</div>`;
  }

  async function askManuals() {
    const s = st.lib;
    const q = String($("lbAsk")?.value || "").trim();
    if (!q) return;
    s.question = q;
    s.answer = { loading: true };
    renderAnswer();
    $("lbAskGo").disabled = true;
    try {
      s.answer = await post("/workshop/ask", { question: q });
    } catch (e) {
      s.answer = { error: errText(e) };
    }
    if ($("lbAskGo")) $("lbAskGo").disabled = false;
    renderAnswer();
  }

  function renderAnswer() {
    const box = $("lbAnswer");
    const a = st.lib?.answer;
    if (!box) return;
    if (!a) { box.innerHTML = ""; return; }
    if (a.loading) { box.innerHTML = `<div class="st-empty">${esc(T("Searching the manuals…"))}</div>`; return; }
    if (a.error) { box.innerHTML = `<div class="st-msg bad">${esc(a.error)}</div>`; return; }
    // The answer text; the sources are shown as cards that open the page.
    const text = String(a.short_answer || "").split("\n\nSources (PDF page numbers):")[0];
    const sources = a.sources || [];
    box.innerHTML = `
      <div class="st-answer">${esc(text)}</div>
      <div class="st-list">${sources.map((src, i) => `
        <button type="button" class="st-item" data-src="${i}">
          <span class="tag">${esc(src.citation)}</span>
          <span class="main"><b>${esc(src.title)} · ${esc(T("page {n}", { n: src.page }))}</b><small>${esc(src.excerpt || "")}</small></span>
        </button>`).join("")}</div>`;
  }

  async function openManual(doc, page = 1) {
    st.screen = "manual";
    main.innerHTML = `
      <div class="st-head">
        <button type="button" class="btn ghost back" id="mnBack">← ${esc(T("Workshop library"))}</button>
        <h2 class="st-doc-title">${esc(doc.title || "")}</h2>
        <span class="grow"></span>
        ${page > 1 ? `<span class="muted">${esc(T("page {n}", { n: page }))}</span>` : ""}
      </div>
      <div class="st-pdf" id="mnBox"><div class="st-empty">${esc(T("Opening the manual… large manuals take a moment."))}</div></div>`;
    $("mnBack").onclick = () => renderLibrary();
    let url = pdfCache.get(doc.id);
    try {
      if (!url) {
        const res = await fetch(`${API}/api/workshop/documents/${encodeURIComponent(doc.id)}/file?inline=1`, { headers: A.authHeaders() });
        if (res.status === 401) { signOut(); return; }
        if (!res.ok) throw new Error(T("Could not open the manual"));
        url = URL.createObjectURL(new Blob([await res.blob()], { type: "application/pdf" }));
        pdfCache.set(doc.id, url);
        while (pdfCache.size > 3) {
          const [oldId, oldUrl] = pdfCache.entries().next().value;
          URL.revokeObjectURL(oldUrl);
          pdfCache.delete(oldId);
        }
      }
    } catch (e) {
      if ($("mnBox")) $("mnBox").innerHTML = `<div class="st-msg bad">${esc(errText(e))}</div>`;
      return;
    }
    if (st.screen !== "manual" || !$("mnBox")) return;
    $("mnBox").innerHTML = `<iframe title="${esc(doc.title || "Manual")}" src="${url}#page=${Math.max(1, Number(page) || 1)}"></iframe>`;
  }

  // ------------------------------------------------------------- scans
  /**
   * A code from the barcode scanner, the phone, or Enter in a search box.
   * Returns true when the code was used.
   */
  async function onScan(raw, { quietUnknown = false } = {}) {
    const code = String(raw || "").trim();
    if (!code) return false;
    touch();
    if (!st.me) { toast(T("Sign in first"), "bad"); beep(false); return false; }
    if (st.captureScan) { st.captureScan(code); return true; }
    if (st.screen === "manual") { toast(T("Close the manual first")); return false; }
    if (overlayOpen()) return false;
    let hit;
    try {
      hit = await api(`/stock/terminal/lookup?code=${encodeURIComponent(code)}`);
    } catch (e) {
      if (e.status === 401) { signOut(); return false; }
      toast(errText(e), "bad");
      return false;
    }
    if (hit.kind === "unknown" || hit.kind === "none") {
      if (quietUnknown) return false;
      beep(false);
      // An unknown box barcode: the stores links it to its part once.
      if (st.me.stores) { linkUnknownBarcode(code); return true; }
      toast(T("Not found: {code}. Ask the storeman to link this barcode.", { code }), "bad");
      return false;
    }
    if (hit.kind === "work_order" && !hit.allowed) {
      toast(T("That work order is not one of your jobs"), "bad");
      beep(false);
      return true;
    }
    beep(true);
    const s = st.screen;
    if (s === "issue") {
      if (hit.kind === "part") { addLine(st.issue.lines, partFromHit(hit)); renderIssueBasket(); toast(T("Added {code}", { code: hit.part_code }), "ok"); }
      else if (hit.kind === "work_order") setIssueJob({ type: "wo", id: hit.id, asset_code: hit.asset_code, label: hit.job || hit.asset_name });
      else if (hit.kind === "asset") await jobFromMachine(hit);
      return true;
    }
    if (s === "receive") {
      if (hit.kind !== "part") { toast(T("Scan a part label"), "bad"); return true; }
      addLine(st.recv.lines, partFromHit(hit));
      renderReceiveBasket();
      toast(T("Added {code}", { code: hit.part_code }), "ok");
      return true;
    }
    if (s === "count") {
      if (hit.kind !== "part") { toast(T("Scan a part label"), "bad"); return true; }
      countPart(partFromHit(hit));
      return true;
    }
    if (s === "find" && hit.kind === "part") {
      st.find.part = partFromHit(hit);
      renderFindPart();
      return true;
    }
    // Anywhere else: open the screen the code belongs to.
    if (hit.kind === "part") startFind(partFromHit(hit));
    else if (hit.kind === "work_order") startIssue({ job: { type: "wo", id: hit.id, asset_code: hit.asset_code, label: hit.job || hit.asset_name } });
    else if (hit.kind === "asset") { startIssue(); await jobFromMachine(hit); }
    return true;
  }

  /** A scanned machine: its one open job, or the list of its jobs. */
  async function jobFromMachine(hit) {
    const s = st.issue;
    s.woQuery = hit.asset_code;
    const data = await api(`/stock/terminal/work-orders?q=${encodeURIComponent(hit.asset_code)}`).catch(() => ({ rows: [] }));
    const rows = (data.rows || []).filter((w) => String(w.asset_code).toUpperCase() === String(hit.asset_code).toUpperCase());
    if (rows.length === 1) return setIssueJob({ type: "wo", id: rows[0].id, asset_code: rows[0].asset_code, label: rows[0].job || rows[0].asset_name });
    if (!rows.length && st.me.stores) return setIssueJob({ type: "asset", asset_code: hit.asset_code, asset_name: hit.asset_name });
    s.job = null;
    renderIssueJob();
  }

  // Keyboard-wedge scanners type fast and press Enter. Outside an input,
  // collect those keys and treat them as a scan.
  let wedge = "";
  let wedgeAt = 0;
  document.addEventListener("keydown", (e) => {
    const o = $("stOverlay");
    if (overlayOpen() && typeof o._keys === "function") {
      if (!/^(INPUT|TEXTAREA)$/.test(e.target.tagName) && o._keys(e)) e.preventDefault();
      return;
    }
    if (/^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName)) return;
    const now = Date.now();
    if (now - wedgeAt > 120) wedge = "";
    wedgeAt = now;
    if (e.key === "Enter") {
      const code = wedge;
      wedge = "";
      if (code.length >= 3) { e.preventDefault(); onScan(code); }
      return;
    }
    if (e.key.length === 1) wedge += e.key;
  });
  // Remove the numpad key handler when the overlay closes.
  new MutationObserver(() => { if (!overlayOpen()) $("stOverlay")._keys = null; }).observe($("stOverlay"), { attributes: true, attributeFilter: ["class"] });

  // ------------------------------------------------------------- phone scanner
  function scanKey() {
    try { return localStorage.getItem(PAIR_KEY) || ""; } catch { return ""; }
  }
  function newScanKey() {
    const bytes = new Uint8Array(16);
    crypto.getRandomValues(bytes);
    const key = Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join("");
    try { localStorage.setItem(PAIR_KEY, key); } catch { /* kiosk storage off */ }
    phone.after = null;
    return key;
  }
  const phone = { after: null, ok: false, busy: false };
  async function pollPhone() {
    const key = scanKey();
    if (!key || phone.busy || document.hidden) return;
    phone.busy = true;
    try {
      const res = await fetch(`${API}/api/stock/terminal/scans?key=${key}&after=${phone.after ?? 0}`, { cache: "no-store" });
      const data = await res.json();
      if (phone.after == null) {
        phone.after = data.last || 0; // ignore scans from before this page opened
      } else {
        for (const r of data.rows || []) {
          phone.after = Math.max(phone.after, r.seq);
          await onScan(r.code);
        }
        if (!(data.rows || []).length && data.last < phone.after) phone.after = data.last; // server restarted
      }
      phone.ok = true;
    } catch {
      phone.ok = false;
    } finally {
      phone.busy = false;
      footer();
    }
  }
  setInterval(pollPhone, 1500);

  function footer() {
    const paired = Boolean(scanKey()) && phone.ok;
    $("stScanDot").classList.toggle("on", paired);
    $("stScanText").textContent = paired
      ? `${T("Phone scanner paired")} · ${T("Scanner ready: scan a part label, machine or work order")}`
      : T("Scanner ready: scan a part label, machine or work order");
  }

  function showPairing() {
    const key = scanKey() || newScanKey();
    const url = `${location.origin}/web/store-scanner.html?key=${key}`;
    let svg = "";
    try {
      const qr = window.qrcode(0, "M");
      qr.addData(url);
      qr.make();
      svg = qr.createSvgTag({ cellSize: 6, margin: 2, scalable: true });
    } catch { svg = ""; }
    const o = overlay(`
      <div class="st-sheet">
        <h2>${esc(T("Pair a phone as the scanner"))}</h2>
        <p class="muted">${esc(T("Scan this code with the phone's camera. Leave the scanner page open on the phone; every code it reads appears here."))}</p>
        <div class="st-pair-qr">${svg}</div>
        <div class="st-pair-url">${esc(url)}</div>
        <p class="muted" style="text-align:center">${esc(T("A new pairing disconnects phones paired before."))}</p>
        <div class="row">
          <button type="button" class="btn ghost" data-pr="new">${esc(T("New pairing"))}</button>
          <button type="button" class="btn primary" data-pr="close">${esc(T("Close"))}</button>
        </div>
      </div>`);
    o.onclick = (e) => {
      const b = e.target.closest("[data-pr]");
      if (!b) return;
      if (b.dataset.pr === "new") { newScanKey(); showPairing(); } else closeOverlay();
    };
  }

  // ------------------------------------------------------------- navigation
  document.addEventListener("click", (e) => {
    const go = e.target.closest("[data-go]");
    if (go && st.me) {
      const to = go.dataset.go;
      if (to === "home") renderHome();
      else if (to === "issue") startIssue();
      else if (to === "receive") startReceive();
      else if (to === "find") startFind();
      else if (to === "requests") renderRequests();
      else if (to === "count") startCount();
      else if (to === "library") startLibrary();
      return;
    }
    const person = e.target.closest("[data-user]");
    if (person && !st.me) askPin(person.dataset.user);
  });
  $("stSignOut").addEventListener("click", () => signOut());
  $("stPairBtn").addEventListener("click", showPairing);
  document.querySelectorAll("[data-lang]").forEach((b) => b.addEventListener("click", () => {
    lang = b.dataset.lang;
    try { localStorage.setItem(LANG_KEY, lang); } catch { /* storage off */ }
    applyStaticText();
    footer();
    // Redraw the current screen in the new language.
    if (!st.me) renderSignIn();
    else if (st.screen === "issue") renderIssue();
    else if (st.screen === "receive") renderReceive();
    else if (st.screen === "find") renderFind();
    else if (st.screen === "requests") renderRequests();
    else if (st.screen === "count") renderCount();
    else if (st.screen === "library") renderLibrary();
    else if (st.screen === "manual") renderLibrary();
    else renderHome();
  }));

  function clock() {
    const d = new Date();
    $("stClock").textContent = d.toLocaleString(lang === "pt" ? "pt-PT" : "en-GB", { weekday: "short", day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" });
  }
  setInterval(clock, 15000);

  // Kiosk: no context menu or pinch zoom on the touch screen.
  document.addEventListener("contextmenu", (e) => e.preventDefault());

  async function start() {
    applyStaticText();
    clock();
    footer();
    await A.loadConfig();
    // The terminal always starts signed out: every person signs in with their own PIN.
    if (A.getAuthToken()) await signOut(true);
    await loadRoster();
    renderSignIn();
    pollPhone();
  }
  start();
})();
