// IRONLOG/web/app/inspections.js — Checklist hub, LDV/machine checklists, vehicle check photos.
// Part of the main app; index.html loads these files in order and they share one global scope.

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

// Accents and case are ignored when matching Portuguese ("travoes" finds "Travões").
function clFold(text) {
  return String(text || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}

function clBuildPtGlossaryReverse() {
  if (clPtGlossaryReverse) return clPtGlossaryReverse;
  clPtGlossaryReverse = {};
  for (const [en, pt] of Object.entries(CL_PT_GLOSSARY)) {
    const key = clFold(pt);
    if (pt && !clPtGlossaryReverse[key]) clPtGlossaryReverse[key] = en;
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
  if (reverse[clFold(src)]) return reverse[clFold(src)];
  // Replace phrase by phrase on the accent-folded text.
  let out = clFold(src);
  const phrases = Object.entries(reverse).sort((a, b) => b[0].length - a[0].length);
  for (const [pt, en] of phrases) {
    if (pt.length < 4) continue;
    const re = new RegExp(pt.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g");
    out = out.replace(re, en);
  }
  return out === clFold(src) ? src : out;
}

function clTranslateFreeText(text, direction = "pt-en") {
  const raw = String(text || "").trim();
  if (!raw) return "";
  return direction === "en-pt" ? clEnToPt(raw) : clPtToEn(raw);
}

function clLabelHtml(englishLabel, portugueseLabel = "") {
  const en = String(englishLabel || "");
  if (!isClPtHelpOn()) return escapeHtml(en);
  // The server sends label_pt for machine checklists; the glossary covers the rest.
  const pt = String(portugueseLabel || "") || clEnToPt(en);
  if (!pt || pt.toLowerCase() === en.toLowerCase()) return escapeHtml(en);
  return `${escapeHtml(en)}<small class="cl-label-pt">${escapeHtml(pt)}</small>`;
}

function clSectionTitleHtml(title, titlePt = "") {
  return clLabelHtml(title, titlePt);
}

async function refreshClChecklistLabels() {
  if (!clSelectedAssetCode) return;
  await selectChecklistAsset(clSelectedAssetCode).catch(() => {});
}

async function runClPtTranslate() {
  const out = qs("clPtTranslateOut");
  const input = String(qs("clPtTranslateIn")?.value || "").trim();
  const dir = String(qs("clPtTranslateDir")?.value || "pt-en");
  if (!out) return;
  if (!input) {
    out.textContent = "Enter text to translate.";
    return;
  }
  out.textContent = "Translating…";
  // Full sentences go through the server's AI translation; the checklist
  // glossary is the fallback when AI is not set up or unreachable.
  try {
    const data = await fetchJson(`${API}/api/translate`, {
      method: "POST",
      body: JSON.stringify({ text: input, to: dir === "en-pt" ? "pt" : "en" }),
    });
    if (data?.ok && data.text) {
      out.textContent = data.text;
      return;
    }
  } catch {}
  const translated = clTranslateFreeText(input, dir);
  const unchanged = !translated || clFold(translated) === clFold(input);
  out.textContent = unchanged
    ? "No translation found. Only checklist words are known offline — ask your admin to set up AI translation for full sentences."
    : `${translated}\n(Word list only — check the meaning.)`;
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
    title.innerHTML = clSectionTitleHtml(String(sec.title || ""), sec.title_pt);
    wrap.appendChild(title);
    for (const it of sec.items || []) {
      const key = String(it.key || "").trim();
      if (!key) continue;
      const id = clSafeDomId(key);
      const row = document.createElement("div");
      row.className = "cl-check-row";
      row.innerHTML = `<input type="checkbox" id="${id}" data-key="${escapeHtml(key)}" ${byKey[key] ? "checked" : ""} /><label for="${id}">${clLabelHtml(it.label || key, it.label_pt)}</label>`;
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
  // An unticked check is saved as a fault and opens a repair work order.
  const unticked = Object.values(checklist).filter((v) => !v).length;
  if (unticked && !confirm(`${unticked} check${unticked === 1 ? " is" : "s are"} not ticked. ${unticked === 1 ? "It" : "They"} will be saved as fault${unticked === 1 ? "" : "s"} and sent to the workshop as a repair work order. Continue?`)) return;

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
  return openAuthedReport(`${API}/api/reports/vehicle-ldv-check/${id}.pdf${q}`, {
    download,
    filename: `IRONLOG_Vehicle_Check_${id}.pdf`,
  });
}

async function clOpenRangePdf(download = false) {
  const start = String(qs("clHistStart")?.value || "").trim();
  const end = String(qs("clHistEnd")?.value || "").trim();
  if (!start || !end) return alert("Select From and To dates first.");
  const assetId = await clResolveHistAssetId();
  const q = new URLSearchParams({ start, end, with_photos: "1" });
  if (assetId > 0) q.set("asset_id", String(assetId));
  if (download) q.set("download", "1");
  return openAuthedReport(`${API}/api/reports/vehicle-ldv-checks.pdf?${q.toString()}`, {
    download,
    filename: `IRONLOG_Vehicle_Checks_${start}_to_${end}.pdf`,
  });
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
  qs("clPtTranslateBtn")?.addEventListener("click", () => runClPtTranslate().catch(() => {}));
  qs("clPtTranslateIn")?.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === "Enter") runClPtTranslate().catch(() => {});
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
    void openAuthedReport(`${API}/api/reports/vehicle-ldv-check/${id}.pdf`, {
      filename: `IRONLOG_Vehicle_Check_${id}.pdf`,
    });
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
  return openAuthedReport(`${API}/api/reports/vehicle-ldv-check/${vcActiveCheckId}.pdf${q}`, {
    download,
    filename: `IRONLOG_Vehicle_Check_${vcActiveCheckId}.pdf`,
  });
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
  return openAuthedReport(`${API}/api/reports/vehicle-ldv-checks.pdf?${q.toString()}`, {
    download,
    filename: `IRONLOG_Vehicle_Checks_${start}_to_${end}.pdf`,
  });
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
