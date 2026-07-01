const API = "/api";
let closingWorkOrderId = null;
let closingWorkOrderSource = "";
let assigningWorkOrderId = null;
let completingWorkOrderId = null;
let currentDetailWorkOrderId = null;
let lastWorkOrderDetail = null;
let stockCatalogCache = [];
let technicianOptions = [];
let lastCreatedRepairWoId = null;
const ROLE_KEY = "ironlog_session_role";
const USER_KEY = "ironlog_session_user";

function getSessionRole() {
  return String(localStorage.getItem(ROLE_KEY) || "admin").trim().toLowerCase() || "admin";
}

function getSessionUser() {
  return String(localStorage.getItem(USER_KEY) || "admin").trim() || "admin";
}

function authHeaders(extra = {}) {
  const h = {
    ...extra,
    "x-user-role": getSessionRole(),
    "x-user-name": getSessionUser(),
  };
  const tok = String(localStorage.getItem("ironlog_auth_token") || sessionStorage.getItem("ironlog_auth_token") || "").trim();
  if (tok) h.Authorization = `Bearer ${tok}`;
  return h;
}

function canRoleTransition(role, currentStatus, nextStatus) {
  const r = String(role || "").toLowerCase();
  if (r === "admin" || r === "supervisor") return true;
  if (r === "artisan") {
    const allowed = {
      assigned: ["in_progress"],
      in_progress: ["completed", "assigned"],
      completed: ["in_progress"],
    };
    return (allowed[currentStatus] || []).includes(nextStatus);
  }
  return false;
}

function canRoleClose(role) {
  return ["admin", "supervisor"].includes(String(role || "").toLowerCase());
}

function isSupervisorRole(role) {
  return ["admin", "supervisor"].includes(String(role || "").toLowerCase());
}

function isArtisanRole(role) {
  return String(role || "").toLowerCase() === "artisan";
}

function technicianIdentityKeys(nameOrUsername, roster = []) {
  const keys = new Set();
  const raw = String(nameOrUsername || "").trim().toLowerCase();
  if (!raw) return keys;
  keys.add(raw);
  for (const t of roster) {
    const username = String(t?.username || t || "").trim().toLowerCase();
    const label = String(t?.label || t?.username || t || "").trim().toLowerCase();
    if (raw === username || raw === label) {
      if (username) keys.add(username);
      if (label) keys.add(label);
    }
  }
  return keys;
}

function isAssignedToMe(wo) {
  const me = getSessionUser().trim().toLowerCase();
  const assigned = String(wo?.assigned_artisan_name || "").trim().toLowerCase();
  if (!me || !assigned) return false;
  const meKeys = technicianIdentityKeys(me, technicianOptions);
  const assignedKeys = technicianIdentityKeys(assigned, technicianOptions);
  for (const k of meKeys) {
    if (assignedKeys.has(k)) return true;
  }
  return false;
}

function workflowStepClass(current, step) {
  const order = ["open", "assigned", "in_progress", "completed", "approved", "closed"];
  const curIdx = order.indexOf(String(current || "").toLowerCase());
  const stepIdx = order.indexOf(step);
  if (curIdx < 0 || stepIdx < 0) return "pill";
  if (curIdx > stepIdx) return "pill green";
  if (curIdx === stepIdx) return "pill blue";
  return "pill";
}

function workflowStepsHtml(wo) {
  const s = String(wo?.status || "open").toLowerCase();
  return `
    <div class="row" style="gap:6px; flex-wrap:wrap; margin:8px 0;">
      <span class="${workflowStepClass(s, "open")}" style="font-size:0.65rem;">Open</span>
      <span class="${workflowStepClass(s, "assigned")}" style="font-size:0.65rem;">Assigned</span>
      <span class="${workflowStepClass(s, "in_progress")}" style="font-size:0.65rem;">In progress</span>
      <span class="${workflowStepClass(s, "completed")}" style="font-size:0.65rem;">Awaiting approval</span>
      <span class="${workflowStepClass(s, "approved")}" style="font-size:0.65rem;">Approved</span>
      <span class="${workflowStepClass(s, "closed")}" style="font-size:0.65rem;">Closed</span>
    </div>
  `;
}

function rolePermissionText(role) {
  const r = String(role || "").toLowerCase();
  if (r === "admin") {
    return "Admin: full control (status transitions, approvals, close, issue parts).";
  }
  if (r === "supervisor") {
    return "Supervisor: assign a technician, update repair progress, start/complete jobs, approve, then close work orders.";
  }
  if (r === "artisan") {
    return "Technician: start assigned jobs, save repair progress (daily PDF), complete with notes, then wait for supervisor approval.";
  }
  if (r === "stores") {
    return "Stores: issue/allocate parts only. Cannot change work order statuses or close.";
  }
  if (r === "operator") {
    return "Operator: read-only work order visibility in this view.";
  }
  return "Role permissions are limited in this view.";
}

function sourceLabel(source) {
  const s = String(source || "").toLowerCase();
  if (s === "service") return "Service";
  if (s === "breakdown") return "Breakdown";
  if (s === "inspection") return "Inspection repair";
  if (s === "manual") return "Manual repair";
  return s || "Unknown";
}

function statusClass(status) {
  const s = String(status || "").toLowerCase();
  if (s === "closed") return "status-ok";
  if (s === "open") return "status-overdue";
  if (s === "in_progress" || s === "assigned") return "status-soon";
  if (s === "completed") return "status-completed";
  if (s === "approved") return "status-approved";
  return "status-ok";
}

function woAgeHours(openedAt) {
  const t = Date.parse(String(openedAt || ""));
  if (!Number.isFinite(t)) return 0;
  return Math.max(0, Math.floor((Date.now() - t) / 3600000));
}

function woPriority(status, openedAt) {
  const s = String(status || "").toLowerCase();
  const age = woAgeHours(openedAt);
  if (s === "completed" && age > 48) return "P1";
  if (s === "in_progress" && age > 72) return "P1";
  if ((s === "open" || s === "assigned") && age > 72) return "P1";
  if (s === "completed" && age > 24) return "P2";
  if (s === "in_progress" && age > 48) return "P2";
  if ((s === "open" || s === "assigned") && age > 48) return "P2";
  return "P3";
}

function escapeHtml(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function canEditRepairProgress(wo) {
  const s = String(wo?.status || "").toLowerCase();
  if (!["open", "assigned", "in_progress"].includes(s)) return false;
  const role = getSessionRole();
  if (isSupervisorRole(role)) return true;
  return isArtisanRole(role) && isAssignedToMe(wo) && s !== "open";
}

function canEditRepairCosts(wo) {
  const s = String(wo?.status || "").toLowerCase();
  if (s === "closed") return false;
  if (isSupervisorRole(getSessionRole())) return true;
  return isArtisanRole(getSessionRole()) && isAssignedToMe(wo);
}

function renderRepairCostsPanel(wo) {
  const s = String(wo?.status || "").toLowerCase();
  const canEdit = canEditRepairCosts(wo);
  const laborHours = Number(wo?.labor_hours || 0);
  const laborCost = Number(wo?.labor_cost || 0);
  const totalOil = Number(wo?.total_oil_cost || 0);
  const hasData = laborHours > 0 || laborCost > 0 || totalOil > 0 || Number(wo?.oil_cost || 0) > 0;
  if (!canEdit && !hasData) return "";

  const rate = Number.isFinite(Number(wo?.labor_rate_per_hour))
    ? Number(wo.labor_rate_per_hour)
    : Number(wo?.default_labor_rate || 0);
  const manualOil = Number(wo?.oil_cost || 0);
  const issuedOil = Number(wo?.issued_oil_cost || 0);
  const tech = String(wo?.assigned_artisan_name || "").trim();
  const inputStyle =
    "width:100%;max-width:200px;padding:10px;border-radius:8px;border:1px solid #2b3f63;background:#0b1628;color:#e8eefc;";

  if (!canEdit) {
    return `
    <div class="item" style="margin-top:12px;">
      <h4 style="margin:0 0 8px 0;">Repair costs &amp; labor</h4>
      <div><strong>Repair hours:</strong> ${laborHours > 0 ? laborHours : "—"}</div>
      <div><strong>Labor rate:</strong> ${rate > 0 ? `$${rate.toFixed(2)}/hr` : "—"}</div>
      <div><strong>Labor cost:</strong> $${laborCost.toFixed(2)}</div>
      <div><strong>Manual oil cost:</strong> $${manualOil.toFixed(2)}</div>
      <div><strong>Issued oil (stores):</strong> $${issuedOil.toFixed(2)}</div>
      <div><strong>Total oil:</strong> $${totalOil.toFixed(2)}</div>
      <div><strong>Technician:</strong> ${escapeHtml(tech || "—")}</div>
    </div>`;
  }

  const hoursVal = laborHours > 0 ? laborHours : "";
  const rateVal = rate > 0 ? rate : "";
  const oilVal = manualOil > 0 ? manualOil : "";
  const techField = isSupervisorRole(getSessionRole())
    ? `<label style="display:flex;flex-direction:column;gap:6px;">
        Technician
        <select id="woRepairCostTechnician" style="padding:10px;border-radius:8px;border:1px solid #2b3f63;background:#0b1628;color:#e8eefc;min-width:200px;"></select>
      </label>`
    : `<div><strong>Technician:</strong> ${escapeHtml(tech || getSessionUser())}</div>`;

  return `
    <div class="item" style="margin-top:12px;">
      <h4 style="margin:0 0 8px 0;">Repair costs &amp; labor <span class="muted" style="font-weight:normal;">(feeds Plant Labor &amp; Oil XLSX)</span></h4>
      <div class="row" style="gap:12px;flex-wrap:wrap;align-items:flex-end;">
        <label style="display:flex;flex-direction:column;gap:6px;">
          Repair hours
          <input id="woRepairCostHours" type="number" min="0" step="0.25" value="${hoursVal}" placeholder="e.g. 4.5" style="${inputStyle}" />
        </label>
        <label style="display:flex;flex-direction:column;gap:6px;">
          Labor rate ($/hr)
          <input id="woRepairCostRate" type="number" min="0" step="0.01" value="${rateVal}" placeholder="Default rate" style="${inputStyle}" />
        </label>
        <label style="display:flex;flex-direction:column;gap:6px;">
          Manual oil cost (USD)
          <input id="woRepairCostOil" type="number" min="0" step="0.01" value="${oilVal}" placeholder="Oil not via stores" style="${inputStyle}" />
        </label>
        ${techField}
      </div>
      <div class="muted small" style="margin-top:8px;">
        Issued oil from stores: $${issuedOil.toFixed(2)} · Total oil: $${totalOil.toFixed(2)} · Labor cost: $${laborCost.toFixed(2)}
      </div>
      <div class="row" style="gap:8px;align-items:center;margin-top:8px;">
        <button type="button" data-wo-save-costs="${wo.id}">Save costs</button>
        <span id="woRepairCostsMsg"></span>
      </div>
    </div>`;
}

function renderRepairProgressPanel(wo) {
  const s = String(wo?.status || "").toLowerCase();
  if (!["open", "assigned", "in_progress"].includes(s)) return "";
  const canEdit = canEditRepairProgress(wo);
  const progress = String(wo?.repair_progress || "").trim();
  if (!canEdit && !progress) return "";
  const updated = wo?.repair_progress_at ? `Last updated: ${wo.repair_progress_at}` : "";
  return `
    <div class="item" style="margin-top:12px;">
      <h4 style="margin:0 0 8px 0;">Repair progress <span class="muted" style="font-weight:normal;">(shows on daily PDF)</span></h4>
      ${
        canEdit
          ? `<textarea id="woRepairProgressInput" rows="3" style="width:100%;max-width:640px;padding:10px;border-radius:8px;border:1px solid #2b3f63;background:#0b1628;color:#e8eefc;" placeholder="e.g. Removed pump, waiting on seal kit — ETA tomorrow">${escapeHtml(progress)}</textarea>
      <div class="row" style="gap:8px;align-items:center;margin-top:8px;">
        <button type="button" data-wo-save-progress="${wo.id}">Save progress</button>
        <span class="muted small" id="woRepairProgressMeta">${escapeHtml(updated)}</span>
        <span id="woRepairProgressMsg"></span>
      </div>`
          : `<div>${escapeHtml(progress) || "—"}</div>
      ${updated ? `<div class="muted small">${escapeHtml(updated)}</div>` : ""}`
      }
    </div>
  `;
}

function workflowActionButtons(wo) {
  const role = getSessionRole();
  const s = String(wo?.status || "").toLowerCase();
  if (s === "closed") return "";

  const buttons = [];
  if (isSupervisorRole(role) && ["open", "assigned"].includes(s)) {
    buttons.push(`<button data-assign-id="${wo.id}" style="margin-top:8px;">Assign technician</button>`);
  }
  if (isSupervisorRole(role) && s === "assigned") {
    buttons.push(`<button data-set-status-id="${wo.id}" data-set-status="in_progress" style="margin-top:8px;">Start job</button>`);
  }
  if (isArtisanRole(role) && s === "assigned" && isAssignedToMe(wo)) {
    buttons.push(`<button data-set-status-id="${wo.id}" data-set-status="in_progress" style="margin-top:8px;">Start job</button>`);
  }
  if (isArtisanRole(role) && s === "in_progress" && isAssignedToMe(wo)) {
    buttons.push(`<button data-complete-id="${wo.id}" style="margin-top:8px;">Mark complete</button>`);
  }
  if (isSupervisorRole(role) && s === "in_progress") {
    buttons.push(`<button data-complete-id="${wo.id}" style="margin-top:8px;">Mark complete</button>`);
  }
  if (isSupervisorRole(role) && s === "completed") {
    buttons.push(`<button data-approve-id="${wo.id}" style="margin-top:8px;">Approve job</button>`);
  }
  if (isSupervisorRole(role) && s === "approved") {
    buttons.push(`<button data-close-id="${wo.id}" data-close-source="${String(wo.source || "").toLowerCase()}" style="margin-top:8px;">Close work order</button>`);
  }
  if (isArtisanRole(role) && ["completed", "approved"].includes(s) && isAssignedToMe(wo)) {
    buttons.push(`<button data-request-close-id="${wo.id}" data-request-close-source="${String(wo.source || "").toLowerCase()}" style="margin-top:8px;">Request close approval</button>`);
  }

  return buttons.join("");
}

function workOrderCard(wo) {
  const ageHours = woAgeHours(wo.opened_at);
  const p = woPriority(wo.status, wo.opened_at);
  const pClass = p === "P1" ? "pri-p1" : p === "P2" ? "pri-p2" : "pri-p3";
  return `
    <div class="card" data-wo-id="${wo.id}">
      <div><strong>WO #${wo.id}</strong></div>
      <div><strong>Asset:</strong> ${wo.asset_code || "-"} - ${wo.asset_name || "-"}</div>
      <div><strong>Source:</strong> ${sourceLabel(wo.source)}</div>
      <div><strong>Reference:</strong> ${wo.reference_id ?? "-"}</div>
      <div><strong>Opened:</strong> ${wo.opened_at || "-"}</div>
      <div><strong>Age:</strong> ${ageHours}h <span class="pill ${pClass}">${p}</span></div>
      <div><strong>Closed:</strong> ${wo.closed_at || "-"}</div>
      <div><strong>Technician:</strong> ${wo.assigned_artisan_name || wo.artisan_name || "Unassigned"}</div>
      ${Number(wo.labor_hours || 0) > 0 ? `<div><strong>Repair hours:</strong> ${Number(wo.labor_hours).toFixed(2)}</div>` : ""}
      ${Number(wo.total_oil_cost || wo.oil_cost || 0) > 0 ? `<div><strong>Oil cost:</strong> $${Number(wo.total_oil_cost || wo.oil_cost || 0).toFixed(2)}</div>` : ""}
      ${wo.repair_progress ? `<div><strong>Progress:</strong> ${escapeHtml(String(wo.repair_progress).slice(0, 160))}${String(wo.repair_progress).length > 160 ? "…" : ""}</div>` : ""}
      ${workflowStepsHtml(wo)}
      <div class="${statusClass(wo.status)}">${String(wo.status || "unknown").toUpperCase()}</div>
      ${workflowActionButtons(wo)}
      <button data-pdf-id="${wo.id}" style="margin-top:8px;">Open PDF</button>
      <button data-pdf-download-id="${wo.id}" style="margin-top:8px;">Download PDF</button>
      <button data-view-id="${wo.id}" style="margin-top:8px;">View Detail</button>
      <button data-wo-qr-open="${wo.id}" style="margin-top:8px;">Open QR Page</button>
      <button data-wo-qr-print="${wo.id}" style="margin-top:8px;">Print WO QR</button>
      <button data-wo-qr-png="${wo.id}" style="margin-top:8px;">Download WO QR PNG</button>
      <button data-wo-qr-link="${wo.id}" style="margin-top:8px;">Copy WO Link</button>
    </div>
  `;
}

async function loadInspectionQuality() {
  const el = document.getElementById("woInspectionQuality");
  if (!el) return;
  try {
    el.className = "muted";
    el.textContent = "Loading quality score...";
    const data = await fetchJson(`${API}/workorders/inspection-quality`, { headers: authHeaders() });
    const s = data?.score || {};
    el.className = "";
    el.innerHTML = `
      <div class="row" style="gap:10px; flex-wrap:wrap;">
        <span class="pill blue">Overall: ${Number(s.overall || 0).toFixed(1)} / 100</span>
        <span class="pill">Completeness: ${Number(s.completeness || 0).toFixed(1)}</span>
        <span class="pill">Photo Evidence: ${Number(s.photo_evidence || 0).toFixed(1)}</span>
        <span class="pill">Comment Quality: ${Number(s.comment_quality || 0).toFixed(1)}</span>
        <span class="pill">Repeat Issue Rate: ${Number(s.repeat_issue_rate || 0).toFixed(1)}%</span>
      </div>
    `;
  } catch (err) {
    el.className = "message-error";
    el.textContent = err.message;
  }
}

async function getWoQrData(woId) {
  const id = Number(woId || 0);
  if (!id) throw new Error("Invalid WO id");
  const data = await fetchJson(`${API}/workorders/${id}/qr-profile/refresh`, {
    method: "POST",
    headers: authHeaders({ "Content-Type": "application/json" }),
    body: JSON.stringify({}),
  });
  const scanUrl = String(data?.qr_payload?.scan_url || "").trim();
  const qrText = String(data?.qr_text || "").trim();
  const value = scanUrl || qrText;
  if (!value) throw new Error("No QR value generated");
  const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=420x420&data=${encodeURIComponent(value)}`;
  return { scanUrl: scanUrl || `/web/workorder-qr.html?wo_id=${id}`, qrUrl };
}

async function openWoQrPage(woId) {
  const data = await getWoQrData(woId);
  window.open(data.scanUrl, "_blank");
}

async function downloadWoQrPng(woId) {
  const data = await getWoQrData(woId);
  const res = await fetch(data.qrUrl);
  if (!res.ok) throw new Error(`QR image fetch failed (${res.status})`);
  const blob = await res.blob();
  const obj = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = obj;
  a.download = `WO_${Number(woId)}_qr.png`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(obj);
}

async function printWoQr(woId) {
  const data = await getWoQrData(woId);
  const win = window.open("", "_blank", "width=900,height=700");
  if (!win) throw new Error("Popup blocked");
  const html = `<!doctype html><html><head><meta charset="utf-8"/><title>WO #${Number(woId)} QR</title>
  <style>body{font-family:Arial,sans-serif;margin:20px;color:#111}.sheet{border:1px solid #333;border-radius:8px;padding:16px;max-width:500px}
  img{width:220px;height:220px;border:1px solid #999}.k{font-size:14px;margin-top:8px}@media print{body{margin:0}.sheet{border:0;padding:8mm}}</style></head>
  <body><div class="sheet"><h2 style="margin:0 0 8px;">IRONLOG WO #${Number(woId)}</h2><img src="${data.qrUrl}" alt="WO QR"/><div class="k">${data.scanUrl}</div></div>
  <script>window.onload=()=>{window.focus();window.print();};</script></body></html>`;
  win.document.open();
  win.document.write(html);
  win.document.close();
}

async function copyWoLink(woId) {
  const data = await getWoQrData(woId);
  await navigator.clipboard.writeText(data.scanUrl);
}

async function fetchJson(url, options) {
  const res = await fetch(url, options);
  const text = await res.text();
  let data = null;
  try { data = text ? JSON.parse(text) : null; } catch { data = null; }
  if (!res.ok) throw new Error(data?.error || data?.message || text || `Request failed (${res.status})`);
  return data || {};
}

function renderParts(parts) {
  if (!Array.isArray(parts) || !parts.length) {
    return `<div class="muted">No parts issued.</div>`;
  }

  return `
    <div class="list">
      ${parts.map((p) => `
        <div class="item">
          <div><strong>${p.part_code || "-"}</strong> - ${p.part_name || "-"}</div>
          <div>Qty: ${p.quantity ?? "-"} | Type: ${p.movement_type || "-"}</div>
        </div>
      `).join("")}
    </div>
  `;
}

function isLubeItem(item) {
  const txt = `${String(item?.part_code || "").toLowerCase()} ${String(item?.part_name || "").toLowerCase()}`;
  return /\blube\b|\boil\b|\bgrease\b|\bhydraulic\b/.test(txt);
}

function formatStockOptions(items) {
  const list = Array.isArray(items) ? items : [];
  const sorted = [...list].sort((a, b) => {
    const aL = isLubeItem(a) ? 1 : 0;
    const bL = isLubeItem(b) ? 1 : 0;
    if (aL !== bL) return aL - bL;
    return String(a.part_code || "").localeCompare(String(b.part_code || ""));
  });
  return sorted
    .map((r) => {
      const code = String(r.part_code || "").trim();
      const name = String(r.part_name || "").trim();
      const onHand = Number(r.on_hand || 0);
      const bucket = isLubeItem(r) ? "Lube" : "Part";
      return `<option value="${code.replace(/"/g, "&quot;")}" data-onhand="${onHand}">[${bucket}] ${code} - ${name} (on hand: ${onHand})</option>`;
    })
    .join("");
}

async function ensureStockCatalogLoaded() {
  if (Array.isArray(stockCatalogCache) && stockCatalogCache.length) return stockCatalogCache;
  const res = await fetch(`${API}/stock/onhand`, { headers: authHeaders() });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to load stock catalog");
  const rows = Array.isArray(data) ? data : [];
  stockCatalogCache = rows.filter((r) => Number(r.on_hand || 0) > 0);
  return stockCatalogCache;
}

function canIssueParts(role) {
  return ["admin", "supervisor", "stores"].includes(String(role || "").toLowerCase());
}

function renderIssuePanel(wo) {
  const role = getSessionRole();
  const canIssue = canIssueParts(role);
  if (!wo || !canIssue) {
    return `<div class="muted">Issue from stores is available for Admin, Supervisor and Stores roles.</div>`;
  }
  return `
    <div class="card" style="margin-top:10px;">
      <h4 style="margin:0 0 8px 0;">Issue Parts / Lube From Stores</h4>
      <div class="row" style="gap:10px; align-items:flex-end; flex-wrap:wrap;">
        <label style="min-width:240px; flex:1;">
          Search Store Item
          <input id="woIssueSearch" type="text" placeholder="Type part code or name..." />
        </label>
        <label style="min-width:360px; flex:1;">
          Store Item
          <select id="woIssuePartCode">
            <option value="">Select part or lube...</option>
          </select>
        </label>
        <label style="min-width:140px;">
          Quantity
          <input id="woIssueQty" type="number" min="1" step="1" value="1" />
        </label>
        <button id="woIssueSubmitBtn" data-wo-issue-id="${Number(wo.id)}">Issue to WO</button>
      </div>
      <div id="woIssueHint" class="muted" style="margin-top:8px;">Pick an item from current stores stock.</div>
      <div id="woIssueMsg" style="margin-top:8px;"></div>
    </div>
  `;
}

function refreshIssueOptions(searchTerm = "") {
  const select = document.getElementById("woIssuePartCode");
  if (!select) return;
  const q = String(searchTerm || "").trim().toLowerCase();
  const rows = Array.isArray(stockCatalogCache) ? stockCatalogCache : [];
  const filtered = !q
    ? rows
    : rows.filter((r) => {
        const hay = `${String(r.part_code || "").toLowerCase()} ${String(r.part_name || "").toLowerCase()}`;
        return hay.includes(q);
      });
  select.innerHTML = `<option value="">Select part or lube...</option>${formatStockOptions(filtered)}`;
}

function renderBreakdown(breakdown) {
  if (!breakdown) return `<div class="muted">No linked breakdown.</div>`;

  return `
    <div class="item">
      <div><strong>ID:</strong> ${breakdown.id}</div>
      <div><strong>Date:</strong> ${breakdown.breakdown_date || "-"}</div>
      <div><strong>Critical:</strong> ${breakdown.critical ? "Yes" : "No"}</div>
      <div><strong>Description:</strong> ${breakdown.description || "-"}</div>
    </div>
  `;
}

function renderDetail(payload) {
  const wo = payload?.work_order;
  const breakdown = payload?.breakdown;
  const parts = payload?.parts_issued;
  const lubeIssued = (Array.isArray(parts) ? parts : []).filter((p) => isLubeItem(p));
  const nonLubeIssued = (Array.isArray(parts) ? parts : []).filter((p) => !isLubeItem(p));

  if (!wo) {
    return `<div class="message-error">Work order detail not available.</div>`;
  }

  return `
    <div class="row" style="gap:20px; align-items:flex-start;">
      <div style="min-width:280px; flex:1;">
        <h4 style="margin:0 0 8px 0;">Core</h4>
        <div class="item">
          <div><strong>WO #:</strong> ${wo.id}</div>
          <div><strong>Asset:</strong> ${wo.asset_code || "-"} - ${wo.asset_name || "-"}</div>
          <div><strong>Source:</strong> ${sourceLabel(wo.source)}</div>
          <div><strong>Reference:</strong> ${wo.reference_id ?? "-"}</div>
          <div><strong>Status:</strong> ${String(wo.status || "").toUpperCase()}</div>
          <div><strong>Assigned technician:</strong> ${wo.assigned_artisan_name || "Unassigned"}</div>
          <div><strong>Assigned at:</strong> ${wo.assigned_at || "-"}</div>
          <div><strong>Started:</strong> ${wo.started_at || "-"}</div>
          <div><strong>Completed:</strong> ${wo.completed_at || "-"}</div>
          <div><strong>Technician sign-off:</strong> ${wo.artisan_name || "-"}</div>
          <div><strong>Supervisor sign-off:</strong> ${wo.supervisor_name || "-"}</div>
          ${wo.job_description ? `<div style="margin-top:8px;"><strong>Job description / findings:</strong><pre style="white-space:pre-wrap; margin:4px 0 0 0; font-family:inherit;">${escapeHtml(wo.job_description)}</pre></div>` : ""}
          ${renderRepairProgressPanel(wo)}
          ${renderRepairCostsPanel(wo)}
          <div><strong>Opened:</strong> ${wo.opened_at || "-"}</div>
          <div><strong>Closed:</strong> ${wo.closed_at || "-"}</div>
          ${workflowStepsHtml(wo)}
          ${isSupervisorRole(getSessionRole()) ? `
            <div class="row" style="gap:8px; flex-wrap:wrap; margin-top:10px;">
              ${["open", "assigned"].includes(String(wo.status || "").toLowerCase()) ? `<button type="button" data-assign-id="${wo.id}">Assign technician</button>` : ""}
              ${String(wo.status || "").toLowerCase() === "assigned" ? `<button type="button" data-set-status-id="${wo.id}" data-set-status="in_progress">Start job</button>` : ""}
              ${String(wo.status || "").toLowerCase() === "in_progress" ? `<button type="button" data-complete-id="${wo.id}">Mark complete</button>` : ""}
              ${String(wo.status || "").toLowerCase() === "completed" ? `<button type="button" data-approve-id="${wo.id}">Approve job</button>` : ""}
              ${String(wo.status || "").toLowerCase() === "approved" ? `<button type="button" data-close-id="${wo.id}" data-close-source="${String(wo.source || "").toLowerCase()}">Close work order</button>` : ""}
            </div>
          ` : ""}
          ${isArtisanRole(getSessionRole()) && isAssignedToMe(wo) ? `
            <div class="row" style="gap:8px; flex-wrap:wrap; margin-top:10px;">
              ${String(wo.status || "").toLowerCase() === "assigned" ? `<button type="button" data-set-status-id="${wo.id}" data-set-status="in_progress">Start job</button>` : ""}
              ${String(wo.status || "").toLowerCase() === "in_progress" ? `<button type="button" data-complete-id="${wo.id}">Mark complete</button>` : ""}
            </div>
          ` : ""}
        </div>
      </div>

      <div style="min-width:280px; flex:1;">
        <h4 style="margin:0 0 8px 0;">Linked Breakdown</h4>
        ${renderBreakdown(breakdown)}
      </div>
    </div>

    <div style="margin-top:12px;">
      <h4 style="margin:0 0 8px 0;">Issued Parts</h4>
      ${renderParts(nonLubeIssued)}
    </div>

    <div style="margin-top:12px;">
      <h4 style="margin:0 0 8px 0;">Issued Lube</h4>
      ${renderParts(lubeIssued)}
    </div>

    ${renderIssuePanel(wo)}
  `;
}

function isTodayStamp(value) {
  const v = String(value || "").trim();
  if (!v) return false;
  const today = new Date().toISOString().slice(0, 10);
  return v.startsWith(today);
}

function updateKpiStrip(rows) {
  const strip = document.getElementById("woKpiStrip");
  if (!strip) return;

  const list = Array.isArray(rows) ? rows : [];
  const count = (status) =>
    list.filter((r) => String(r.status || "").toLowerCase() === status).length;

  const all = list.length;
  const open = count("open");
  const inProgress = count("in_progress");
  const awaitingApproval = count("completed");
  const approvedToday = list.filter(
    (r) => String(r.status || "").toLowerCase() === "approved" && isTodayStamp(r.closed_at || r.opened_at)
  ).length;
  const closedToday = list.filter(
    (r) => String(r.status || "").toLowerCase() === "closed" && isTodayStamp(r.closed_at)
  ).length;

  strip.innerHTML = `
    <button class="pill blue" data-kpi-filter="">All: ${all}</button>
    <button class="pill red" data-kpi-filter="open">Open: ${open}</button>
    <button class="pill orange" data-kpi-filter="in_progress">In Progress: ${inProgress}</button>
    <button class="pill orange" data-kpi-filter="completed">Awaiting Approval: ${awaitingApproval}</button>
    <button class="pill blue" data-kpi-filter="approved">Approved Today: ${approvedToday}</button>
    <button class="pill blue" data-kpi-filter="closed">Closed Today: ${closedToday}</button>
  `;
}

async function fetchWorkOrders() {
  const statusEl = document.getElementById("woStatus");
  const sourceEl = document.getElementById("woSource");
  const searchEl = document.getElementById("woSearch");
  const listEl = document.getElementById("woList");
  const msgEl = document.getElementById("woMessage");

  if (!statusEl || !sourceEl || !searchEl || !listEl || !msgEl) return;

  const status = String(statusEl.value || "").trim();
  const source = String(sourceEl.value || "").trim().toLowerCase();
  const q = String(searchEl.value || "").trim().toLowerCase();

  listEl.innerHTML = `<div class="skeleton-block"></div><div class="skeleton-block"></div>`;
  msgEl.className = "";
  msgEl.textContent = "";

  try {
    const res = await fetch(`${API}/workorders`, { headers: authHeaders() });
    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || "Failed to load work orders");
    }

    const rows = Array.isArray(data) ? data : [];
    updateKpiStrip(rows);

    const filtered = rows.filter((r) => {
      const statusOk = !status || String(r.status || "").toLowerCase() === status.toLowerCase();
      if (!statusOk) return false;

      const sourceOk = !source || String(r.source || "").toLowerCase() === source;
      if (!sourceOk) return false;

      if (!q) return true;
      const hay = `${r.id} ${r.asset_code || ""} ${r.asset_name || ""}`.toLowerCase();
      return hay.includes(q);
    });

    listEl.innerHTML = filtered.length
      ? filtered.map(workOrderCard).join("")
      : "<div>No work orders found for current filters.</div>";

    const requested = getRequestedWorkOrderId();
    if (requested && filtered.some((r) => Number(r.id) === requested)) {
      loadWorkOrderDetail(requested).catch(() => {});
      setTimeout(() => scrollToWorkOrderCard(requested), 0);
    }

    msgEl.className = "message-success";
    msgEl.textContent = `Showing ${filtered.length} work order(s).`;
  } catch (err) {
    console.error("Load work orders error:", err);
    listEl.innerHTML = `<div style="color:#ff8080;">Error loading work orders: ${err.message}</div>`;
    msgEl.className = "message-error";
    msgEl.textContent = err.message;
  }
}

function openCloseModal(id) {
  const woId = Number(id || 0);
  if (!woId) return;

  closingWorkOrderId = woId;
  const modal = document.getElementById("woCloseModal");
  const title = document.getElementById("woCloseModalTitle");
  const notesEl = document.getElementById("woCloseNotes");
  const artisanEl = document.getElementById("woCloseArtisan");
  const supervisorEl = document.getElementById("woCloseSupervisor");
  const msgEl = document.getElementById("woCloseModalMsg");

  if (!modal || !title || !notesEl || !artisanEl || !supervisorEl || !msgEl) return;

  title.textContent = `#${woId}`;
  notesEl.value = "";
  artisanEl.value = "";
  supervisorEl.value = "";
  msgEl.className = "";
  msgEl.textContent = "";
  modal.style.display = "flex";
}

function openCloseModalForRow(id, source) {
  closingWorkOrderSource = String(source || "").toLowerCase();
  openCloseModal(id);

  const msgEl = document.getElementById("woCloseModalMsg");
  if (closingWorkOrderSource === "service" && msgEl) {
    msgEl.className = "";
    msgEl.textContent = "Artisan name and completion notes are required for service work orders.";
  }
}

function closeCloseModal() {
  closingWorkOrderId = null;
  closingWorkOrderSource = "";
  const modal = document.getElementById("woCloseModal");
  if (modal) modal.style.display = "none";
}

async function submitCloseWorkOrder() {
  const woId = Number(closingWorkOrderId || 0);
  if (!woId) return;

  const notesEl = document.getElementById("woCloseNotes");
  const artisanEl = document.getElementById("woCloseArtisan");
  const supervisorEl = document.getElementById("woCloseSupervisor");
  const msgEl = document.getElementById("woCloseModalMsg");
  const confirmBtn = document.getElementById("woCloseConfirmBtn");

  const completion_notes = String(notesEl?.value || "").trim();
  const artisan_name = String(artisanEl?.value || "").trim();
  const supervisor_name = String(supervisorEl?.value || "").trim();

  if (closingWorkOrderSource === "service" && !artisan_name) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = "Artisan name is required for service work orders.";
    }
    return;
  }
  if (closingWorkOrderSource === "service" && !completion_notes) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = "Completion notes are required for service work orders.";
    }
    return;
  }

  if (confirmBtn) confirmBtn.disabled = true;
  if (msgEl) {
    msgEl.className = "";
    msgEl.textContent = "Saving...";
  }

  try {
    const isRequest = getSessionRole() === "artisan";
    const endpoint = isRequest
      ? `${API}/workorders/${woId}/request-close`
      : `${API}/workorders/${woId}/close`;
    const res = await fetch(endpoint, {
      method: "POST",
      headers: {
        ...authHeaders(),
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        completion_notes,
        artisan_name,
        supervisor_name
      })
    });
    const data = await res.json();
    if (!res.ok) {
      throw new Error(data.error || "Failed to close work order");
    }
    closeCloseModal();
    await fetchWorkOrders();
    const detailEl = document.getElementById("woDetail");
    if (detailEl) {
      detailEl.innerHTML = isRequest
        ? `<div class="message-success">Close approval requested for work order #${woId} (Request #${data.request_id || "-"})</div>`
        : `<div class="message-success">Work order #${woId} completed and closed.</div>`;
    }
  } catch (err) {
    console.error("Close work order error:", err);
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message;
    }
  } finally {
    if (confirmBtn) confirmBtn.disabled = false;
  }
}

async function loadWorkOrderDetail(id) {
  const woId = Number(id || 0);
  const detailEl = document.getElementById("woDetail");
  if (!woId || !detailEl) return;

  detailEl.innerHTML = `<div class="skeleton-block"></div>`;
  currentDetailWorkOrderId = woId;

  try {
    const res = await fetch(`${API}/workorders/${woId}`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) {
      throw new Error(data.error || "Failed to load work order detail");
    }
    lastWorkOrderDetail = data;
    detailEl.innerHTML = renderDetail(data);
    const techSelect = document.getElementById("woRepairCostTechnician");
    if (techSelect) {
      await loadTechnicians();
      fillTechnicianSelect(techSelect, data.work_order?.assigned_artisan_name || "");
    }
    if (canIssueParts(getSessionRole())) {
      const select = document.getElementById("woIssuePartCode");
      const searchInput = document.getElementById("woIssueSearch");
      if (select) {
        const rows = await ensureStockCatalogLoaded();
        refreshIssueOptions("");
        if (searchInput) {
          searchInput.value = "";
        }
      }
    }
  } catch (err) {
    console.error("Load work order detail error:", err);
    detailEl.innerHTML = `<div class="message-error">${err.message}</div>`;
  }
}

async function issueToWorkOrder() {
  const woId = Number(currentDetailWorkOrderId || 0);
  if (!woId) return;
  const msg = document.getElementById("woIssueMsg");
  const partSelect = document.getElementById("woIssuePartCode");
  const qtyInput = document.getElementById("woIssueQty");
  const hint = document.getElementById("woIssueHint");
  const btn = document.getElementById("woIssueSubmitBtn");
  const part_code = String(partSelect?.value || "").trim();
  const quantity = Number(qtyInput?.value || 0);
  if (!part_code || !Number.isFinite(quantity) || quantity <= 0) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Select a store item and quantity > 0.";
    }
    return;
  }
  if (btn) btn.disabled = true;
  if (msg) {
    msg.className = "";
    msg.textContent = "Issuing...";
  }
  try {
    const res = await fetch(`${API}/workorders/${woId}/issue`, {
      method: "POST",
      headers: { ...authHeaders(), "Content-Type": "application/json" },
      body: JSON.stringify({ part_code, quantity }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Issue failed");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `Issued ${quantity} x ${part_code} to WO #${woId}.`;
    }
    stockCatalogCache = [];
    await loadWorkOrderDetail(woId);
    await fetchWorkOrders();
    if (hint) hint.textContent = "Issued successfully. Stock refreshed.";
  } catch (err) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = err.message || String(err);
    }
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function loadTechnicians() {
  try {
    const data = await fetchJson(`${API}/workorders/technicians`, { headers: authHeaders() });
    if (Array.isArray(data?.technician_users) && data.technician_users.length) {
      technicianOptions = data.technician_users;
    } else {
      technicianOptions = (Array.isArray(data?.technicians) ? data.technicians : []).map((name) => ({
        username: name,
        label: name,
      }));
    }
  } catch {
    technicianOptions = [];
  }
  return technicianOptions;
}

function renderTechnicianRoster() {
  const listEl = document.getElementById("woTechniciansList");
  if (!listEl) return;
  const users = Array.isArray(technicianOptions) ? technicianOptions : [];
  if (!users.length) {
    listEl.innerHTML = `<small class="muted">No technicians yet. Add one below — they will appear on the terminal login screen after you set a PIN.</small>`;
    return;
  }
  listEl.innerHTML = users
    .map((t) => {
      const username = escapeHtml(String(t.username || t.label || t || "").trim());
      const label = escapeHtml(String(t.label || t.username || t || "").trim());
      const sub = label && label.toLowerCase() !== username.toLowerCase() ? `<small class="muted"> (${username})</small>` : "";
      return `<div class="item"><strong>${label || username}</strong>${sub}</div>`;
    })
    .join("");
}

async function loadTechnicianRoster() {
  const msgEl = document.getElementById("woTechniciansMsg");
  try {
    await loadTechnicians();
    renderTechnicianRoster();
    fillTechnicianSelect(document.getElementById("woRepairAssignee"));
    if (msgEl) {
      msgEl.className = "muted";
      msgEl.textContent = `${technicianOptions.length} technician(s) on roster.`;
    }
  } catch (err) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message || String(err);
    }
  }
}

async function loadRepairAssetOptions() {
  const dl = document.getElementById("woRepairAssetList");
  if (!dl) return;
  try {
    const rows = await fetchJson(`${API}/assets?include_archived=0`, { headers: authHeaders() });
    const arr = Array.isArray(rows) ? rows : [];
    dl.innerHTML = arr
      .map((a) => `<option value="${escapeHtml(String(a.asset_code || ""))}"></option>`)
      .join("");
  } catch {
    dl.innerHTML = "";
  }
}

async function createRepairWorkOrder() {
  const msgEl = document.getElementById("woRepairMsg");
  const asset_code = String(document.getElementById("woRepairAsset")?.value || "").trim();
  const component = String(document.getElementById("woRepairComponent")?.value || "").trim();
  const description = String(document.getElementById("woRepairDescription")?.value || "").trim();
  const inspectionRaw = String(document.getElementById("woRepairInspectionId")?.value || "").trim();
  const inspection_id = inspectionRaw ? Number(inspectionRaw) : 0;
  const assigned_artisan_name = String(document.getElementById("woRepairAssignee")?.value || "").trim();

  if (!asset_code) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = "Asset code is required.";
    }
    return;
  }
  if (!description) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = "Repair description is required.";
    }
    return;
  }

  const btn = document.getElementById("woRepairCreateBtn");
  if (btn) btn.disabled = true;
  if (msgEl) {
    msgEl.className = "muted";
    msgEl.textContent = "Creating repair work order…";
  }

  try {
    const body = { asset_code, description };
    if (component) body.component = component;
    if (Number.isFinite(inspection_id) && inspection_id > 0) body.inspection_id = inspection_id;
    if (assigned_artisan_name) body.assigned_artisan_name = assigned_artisan_name;

    const data = await fetchJson(`${API}/workorders/repair`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify(body),
    });

    lastCreatedRepairWoId = Number(data.work_order_id || 0) || null;
    const openBtn = document.getElementById("woRepairOpenWoBtn");
    if (openBtn && lastCreatedRepairWoId) openBtn.style.display = "";

    if (msgEl) {
      msgEl.className = "message-success";
      msgEl.textContent = data.already_exists
        ? `Work order #${lastCreatedRepairWoId} already linked to inspection (status: ${data.status || "open"}).`
        : `Created repair work order #${lastCreatedRepairWoId} (${data.status || "open"}) — no breakdown logged.`;
    }

    await fetchWorkOrders();
    if (lastCreatedRepairWoId) {
      loadWorkOrderDetail(lastCreatedRepairWoId);
      scrollToWorkOrderCard(lastCreatedRepairWoId);
    }
  } catch (err) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message || String(err);
    }
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function saveWorkshopTechnician() {
  const msgEl = document.getElementById("woTechniciansMsg");
  const username = String(document.getElementById("woTechUsername")?.value || "").trim();
  const full_name = String(document.getElementById("woTechFullName")?.value || "").trim();
  const pin = String(document.getElementById("woTechPin")?.value || "").replace(/\D/g, "");
  if (!username) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = "Username is required.";
    }
    return;
  }
  if (!pin || pin.length < 4 || pin.length > 6) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = "Enter a 4–6 digit PIN for workshop login.";
    }
    return;
  }
  const btn = document.getElementById("woTechSaveBtn");
  if (btn) btn.disabled = true;
  if (msgEl) {
    msgEl.className = "muted";
    msgEl.textContent = "Saving technician…";
  }
  try {
    await fetchJson(`${API}/workorders/technicians`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({
        username,
        full_name: full_name || username,
        pin,
      }),
    });
    if (msgEl) {
      msgEl.className = "message-success";
      msgEl.textContent = `Saved ${full_name || username}. They can sign in on the Technician Terminal and appear in Assign technician.`;
    }
    const pinEl = document.getElementById("woTechPin");
    if (pinEl) pinEl.value = "";
    await loadTechnicianRoster();
  } catch (err) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message || String(err);
    }
  } finally {
    if (btn) btn.disabled = false;
  }
}

function fillTechnicianSelect(selectEl, selectedName = "") {
  if (!selectEl) return;
  const selected = String(selectedName || "").trim();
  const users = Array.isArray(technicianOptions) ? technicianOptions : [];
  const merged = [...users];
  if (selected && !merged.some((t) => String(t.username || t).toLowerCase() === selected.toLowerCase() || String(t.label || t).toLowerCase() === selected.toLowerCase())) {
    merged.unshift({ username: selected, label: selected });
  }
  selectEl.innerHTML =
    `<option value="">Select technician...</option>` +
    merged
      .map((t) => {
        const username = String(t.username || t || "").trim();
        const label = String(t.label || t.username || t || "").trim();
        const value = username || label;
        const escVal = value.replace(/"/g, "&quot;");
        const escLabel = label.replace(/</g, "&lt;");
        const sel =
          selected &&
          (selected.toLowerCase() === value.toLowerCase() || selected.toLowerCase() === label.toLowerCase())
            ? " selected"
            : "";
        return `<option value="${escVal}"${sel}>${escLabel}</option>`;
      })
      .join("");
}

function openAssignModal(id) {
  const woId = Number(id || 0);
  if (!woId) return;
  assigningWorkOrderId = woId;
  const modal = document.getElementById("woAssignModal");
  const title = document.getElementById("woAssignModalTitle");
  const select = document.getElementById("woAssignTechnician");
  const manual = document.getElementById("woAssignTechnicianManual");
  const msgEl = document.getElementById("woAssignModalMsg");
  if (!modal || !title || !select) return;
  title.textContent = `#${woId}`;
  fillTechnicianSelect(select);
  if (manual) manual.value = "";
  if (msgEl) {
    msgEl.className = "";
    msgEl.textContent = "";
  }
  modal.style.display = "flex";
}

function closeAssignModal() {
  assigningWorkOrderId = null;
  const modal = document.getElementById("woAssignModal");
  if (modal) modal.style.display = "none";
}

async function submitAssignWorkOrder() {
  const woId = Number(assigningWorkOrderId || 0);
  if (!woId) return;
  const select = document.getElementById("woAssignTechnician");
  const manual = document.getElementById("woAssignTechnicianManual");
  const msgEl = document.getElementById("woAssignModalMsg");
  const confirmBtn = document.getElementById("woAssignConfirmBtn");
  const assigned_artisan_name =
    String(manual?.value || "").trim() || String(select?.value || "").trim();
  if (!assigned_artisan_name) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = "Select or type a technician name.";
    }
    return;
  }
  if (confirmBtn) confirmBtn.disabled = true;
  if (msgEl) {
    msgEl.className = "";
    msgEl.textContent = "Assigning...";
  }
  try {
    await fetchJson(`${API}/workorders/${woId}/assign`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ assigned_artisan_name }),
    });
    closeAssignModal();
    await fetchWorkOrders();
    await loadWorkOrderDetail(woId);
  } catch (err) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message || String(err);
    }
  } finally {
    if (confirmBtn) confirmBtn.disabled = false;
  }
}

function openCompleteModal(id) {
  const woId = Number(id || 0);
  if (!woId) return;
  completingWorkOrderId = woId;
  const modal = document.getElementById("woCompleteModal");
  const title = document.getElementById("woCompleteModalTitle");
  const notesEl = document.getElementById("woCompleteNotes");
  const artisanEl = document.getElementById("woCompleteArtisan");
  const laborEl = document.getElementById("woCompleteLaborHours");
  const oilEl = document.getElementById("woCompleteOilCost");
  const msgEl = document.getElementById("woCompleteModalMsg");
  if (!modal || !title || !notesEl || !artisanEl) return;
  title.textContent = `#${woId}`;
  notesEl.value = "";
  artisanEl.value = getSessionUser();
  const wo = lastWorkOrderDetail?.work_order;
  if (laborEl) {
    laborEl.value =
      wo && Number(wo.id) === woId && Number(wo.labor_hours || 0) > 0 ? String(wo.labor_hours) : "";
  }
  if (oilEl) {
    oilEl.value =
      wo && Number(wo.id) === woId && Number(wo.oil_cost || 0) > 0 ? String(wo.oil_cost) : "";
  }
  if (msgEl) {
    msgEl.className = "";
    msgEl.textContent = "";
  }
  modal.style.display = "flex";
}

function closeCompleteModal() {
  completingWorkOrderId = null;
  const modal = document.getElementById("woCompleteModal");
  if (modal) modal.style.display = "none";
}

async function submitCompleteWorkOrder() {
  const woId = Number(completingWorkOrderId || 0);
  if (!woId) return;
  const notesEl = document.getElementById("woCompleteNotes");
  const artisanEl = document.getElementById("woCompleteArtisan");
  const msgEl = document.getElementById("woCompleteModalMsg");
  const confirmBtn = document.getElementById("woCompleteConfirmBtn");
  const completion_notes = String(notesEl?.value || "").trim();
  const artisan_name = String(artisanEl?.value || getSessionUser()).trim();
  const laborRaw = String(document.getElementById("woCompleteLaborHours")?.value || "").trim();
  const oilRaw = String(document.getElementById("woCompleteOilCost")?.value || "").trim();
  if (!completion_notes) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = "Completion notes are required.";
    }
    return;
  }
  if (confirmBtn) confirmBtn.disabled = true;
  if (msgEl) {
    msgEl.className = "";
    msgEl.textContent = "Submitting...";
  }
  try {
    const extra = { completion_notes, artisan_name };
    if (laborRaw !== "") extra.labor_hours = Math.max(0, Number(laborRaw));
    if (oilRaw !== "") extra.oil_cost = Math.max(0, Number(oilRaw));
    await setWorkOrderStatus(woId, "completed", extra);
    closeCompleteModal();
  } catch (err) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message || String(err);
    }
  } finally {
    if (confirmBtn) confirmBtn.disabled = false;
  }
}

async function approveWorkOrder(id) {
  const woId = Number(id || 0);
  if (!woId) return;
  const supervisor_name = getSessionUser();
  const ok = window.confirm(`Approve completed work on WO #${woId}?`);
  if (!ok) return;
  await setWorkOrderStatus(woId, "approved", { supervisor_name });
}

async function saveRepairCosts(woId) {
  const id = Number(woId || 0);
  if (!id) return;
  const hoursEl = document.getElementById("woRepairCostHours");
  const rateEl = document.getElementById("woRepairCostRate");
  const oilEl = document.getElementById("woRepairCostOil");
  const techEl = document.getElementById("woRepairCostTechnician");
  const msg = document.getElementById("woRepairCostsMsg");
  const body = {};
  const hoursRaw = String(hoursEl?.value ?? "").trim();
  const rateRaw = String(rateEl?.value ?? "").trim();
  const oilRaw = String(oilEl?.value ?? "").trim();
  if (hoursRaw !== "") body.labor_hours = Math.max(0, Number(hoursRaw));
  if (rateRaw !== "") body.labor_rate_per_hour = Math.max(0, Number(rateRaw));
  if (oilRaw !== "") body.oil_cost = Math.max(0, Number(oilRaw));
  if (techEl && isSupervisorRole(getSessionRole())) {
    const tech = String(techEl.value || "").trim();
    if (tech) body.assigned_artisan_name = tech;
  }
  if (msg) {
    msg.className = "";
    msg.textContent = "Saving...";
  }
  try {
    const res = await fetch(`${API}/workorders/${id}/costs`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify(body),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save costs");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = "Costs saved.";
    }
    await fetchWorkOrders();
    if (currentDetailWorkOrderId === id) await loadWorkOrderDetail(id);
  } catch (err) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = err.message || String(err);
    }
  }
}

async function saveRepairProgress(woId) {
  const id = Number(woId || 0);
  if (!id) return;
  const input = document.getElementById("woRepairProgressInput");
  const msg = document.getElementById("woRepairProgressMsg");
  const meta = document.getElementById("woRepairProgressMeta");
  const repair_progress = String(input?.value || "").trim();
  if (msg) {
    msg.className = "";
    msg.textContent = "Saving...";
  }
  try {
    const res = await fetchJson(`${API}/workorders/${id}/progress`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ repair_progress }),
    });
    if (meta) meta.textContent = res.repair_progress_at ? `Last updated: ${res.repair_progress_at}` : "";
    if (msg) {
      msg.className = "message-success";
      msg.textContent = "Progress saved — will appear on the daily PDF.";
    }
    await fetchWorkOrders();
    if (currentDetailWorkOrderId === id) await loadWorkOrderDetail(id);
  } catch (err) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = err.message || String(err);
    }
  }
}

async function setWorkOrderStatus(id, status, extraBody = {}) {
  const woId = Number(id || 0);
  const next = String(status || "").trim().toLowerCase();
  if (!woId || !next) return;

  try {
    const res = await fetch(`${API}/workorders/${woId}/status`, {
      method: "POST",
      headers: {
        ...authHeaders(),
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ status: next, ...extraBody })
    });
    const data = await res.json();
    if (!res.ok) {
      throw new Error(data.error || "Failed to update status");
    }
    await fetchWorkOrders();
    await loadWorkOrderDetail(woId);
    return data;
  } catch (err) {
    console.error("Set status error:", err);
    alert(`Could not update status: ${err.message}`);
    throw err;
  }
}

function openWorkOrderPdf(id) {
  const woId = Number(id || 0);
  if (!woId) return;
  window.open(`${API}/reports/workorder/${woId}.pdf`, "_blank");
}

function downloadWorkOrderPdf(id) {
  const woId = Number(id || 0);
  if (!woId) return;
  window.open(`${API}/reports/workorder/${woId}.pdf?download=1`, "_blank");
}

async function downloadPlantLaborOilReport() {
  const yearEl = document.getElementById("woLaborReportYear");
  const msgEl = document.getElementById("woMessage");
  const year = Number(yearEl?.value) || new Date().getFullYear();
  if (msgEl) {
    msgEl.className = "";
    msgEl.textContent = "Generating report...";
  }
  try {
    const res = await fetch(`${API}/reports/plant-labor-oil.xlsx?year=${encodeURIComponent(year)}`, {
      headers: authHeaders(),
    });
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || `Download failed (${res.status})`);
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `plant-labor-oil-${year}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    if (msgEl) {
      msgEl.className = "message-success";
      msgEl.textContent = `Downloaded Plant Labor & Oil report for ${year}.`;
    }
  } catch (err) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message || String(err);
    }
  }
}

function getRequestedWorkOrderId() {
  try {
    const q = new URLSearchParams(window.location.search);
    const id = Number(q.get("wo") || 0);
    return Number.isFinite(id) && id > 0 ? id : null;
  } catch {
    return null;
  }
}

function scrollToWorkOrderCard(woId) {
  const id = Number(woId || 0);
  if (!id) return;
  const card = document.querySelector(`.card[data-wo-id="${id}"]`);
  if (!card) return;
  card.classList.add("wo-highlight");
  card.scrollIntoView({ behavior: "smooth", block: "center" });
  setTimeout(() => card.classList.remove("wo-highlight"), 2600);
}

document.addEventListener("DOMContentLoaded", () => {
  const refreshBtn = document.getElementById("woRefreshBtn");
  const awaitingApprovalBtn = document.getElementById("woAwaitingApprovalBtn");
  const statusEl = document.getElementById("woStatus");
  const sourceEl = document.getElementById("woSource");
  const searchEl = document.getElementById("woSearch");
  const listEl = document.getElementById("woList");
  const closeConfirmBtn = document.getElementById("woCloseConfirmBtn");
  const closeCancelBtn = document.getElementById("woCloseCancelBtn");
  const closeModal = document.getElementById("woCloseModal");
  const kpiStrip = document.getElementById("woKpiStrip");
  const detailEl = document.getElementById("woDetail");
  const role = getSessionRole();
  try {
    const q = new URLSearchParams(window.location.search);
    const initialSearch = String(q.get("search") || "").trim();
    if (initialSearch && searchEl) {
      searchEl.value = initialSearch;
      if (statusEl) statusEl.value = "";
    }
  } catch {}
  const roleBadge = document.getElementById("woRoleBadge");
  if (roleBadge) roleBadge.textContent = `Role: ${role}`;
  const legend = document.getElementById("woPermissionLegend");
  if (legend) legend.textContent = rolePermissionText(role);

  if (!["admin", "supervisor"].includes(role) && awaitingApprovalBtn) {
    awaitingApprovalBtn.style.display = "none";
  }
  const techPanel = document.getElementById("woTechniciansPanel");
  if (techPanel) {
    if (isSupervisorRole(role)) {
      techPanel.style.display = "";
    } else {
      techPanel.style.display = "none";
    }
  }
  const repairPanel = document.getElementById("woCreateRepairPanel");
  if (repairPanel) {
    repairPanel.style.display = isSupervisorRole(role) ? "" : "none";
  }
  if (isArtisanRole(role) && statusEl) {
    statusEl.value = "";
  }

  if (refreshBtn) refreshBtn.addEventListener("click", fetchWorkOrders);
  if (awaitingApprovalBtn && statusEl) {
    awaitingApprovalBtn.addEventListener("click", () => {
      statusEl.value = "completed";
      fetchWorkOrders();
    });
  }
  if (statusEl) statusEl.addEventListener("change", fetchWorkOrders);
  if (sourceEl) sourceEl.addEventListener("change", fetchWorkOrders);
  if (searchEl) searchEl.addEventListener("input", fetchWorkOrders);
  const requested = getRequestedWorkOrderId();
  if (requested && searchEl) searchEl.value = String(requested);
  if (closeConfirmBtn) closeConfirmBtn.addEventListener("click", submitCloseWorkOrder);
  if (closeCancelBtn) closeCancelBtn.addEventListener("click", closeCloseModal);
  document.getElementById("woAssignConfirmBtn")?.addEventListener("click", () => submitAssignWorkOrder());
  document.getElementById("woAssignCancelBtn")?.addEventListener("click", closeAssignModal);
  document.getElementById("woCompleteConfirmBtn")?.addEventListener("click", () => submitCompleteWorkOrder());
  document.getElementById("woCompleteCancelBtn")?.addEventListener("click", closeCompleteModal);
  document.getElementById("woTechSaveBtn")?.addEventListener("click", () => saveWorkshopTechnician());
  document.getElementById("woTechRefreshBtn")?.addEventListener("click", () => loadTechnicianRoster());
  document.getElementById("woRepairCreateBtn")?.addEventListener("click", () => createRepairWorkOrder());
  document.getElementById("woRepairOpenWoBtn")?.addEventListener("click", () => {
    if (lastCreatedRepairWoId) loadWorkOrderDetail(lastCreatedRepairWoId);
  });
  const laborYearEl = document.getElementById("woLaborReportYear");
  if (laborYearEl && !laborYearEl.value) laborYearEl.value = String(new Date().getFullYear());
  document.getElementById("woLaborReportBtn")?.addEventListener("click", () => downloadPlantLaborOilReport());
  if (closeModal) {
    closeModal.addEventListener("click", (evt) => {
      if (evt.target === closeModal) closeCloseModal();
    });
  }
  if (kpiStrip && statusEl) {
    kpiStrip.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const filter = target.getAttribute("data-kpi-filter");
      if (filter == null) return;
      statusEl.value = filter;
      fetchWorkOrders();
    });
  }

  if (listEl) {
    listEl.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const pdfId = target.getAttribute("data-pdf-id");
      const pdfDownloadId = target.getAttribute("data-pdf-download-id");
      const viewId = target.getAttribute("data-view-id");
      const id = target.getAttribute("data-close-id");
      const setStatusId = target.getAttribute("data-set-status-id");
      const setStatus = target.getAttribute("data-set-status");
      const assignId = target.getAttribute("data-assign-id");
      const completeId = target.getAttribute("data-complete-id");
      const approveId = target.getAttribute("data-approve-id");
      const rowSource = target.getAttribute("data-close-source");
      const requestCloseId = target.getAttribute("data-request-close-id");
      const requestCloseSource = target.getAttribute("data-request-close-source");
      const woQrOpen = target.getAttribute("data-wo-qr-open");
      const woQrPrint = target.getAttribute("data-wo-qr-print");
      const woQrPng = target.getAttribute("data-wo-qr-png");
      const woQrLink = target.getAttribute("data-wo-qr-link");
      if (pdfId) {
        openWorkOrderPdf(pdfId);
        return;
      }
      if (pdfDownloadId) {
        downloadWorkOrderPdf(pdfDownloadId);
        return;
      }
      if (viewId) {
        loadWorkOrderDetail(viewId);
        return;
      }
      if (setStatusId && setStatus) {
        if (setStatus === "completed") {
          openCompleteModal(setStatusId);
        } else {
          setWorkOrderStatus(setStatusId, setStatus).catch(() => {});
        }
        return;
      }
      if (assignId) {
        openAssignModal(assignId);
        return;
      }
      if (completeId) {
        openCompleteModal(completeId);
        return;
      }
      if (approveId) {
        approveWorkOrder(approveId).catch((e) => alert(`Approve failed: ${e.message}`));
        return;
      }
      if (requestCloseId) {
        openCloseModalForRow(requestCloseId, requestCloseSource);
        return;
      }
      if (woQrOpen) {
        openWoQrPage(woQrOpen).catch((e) => alert(`WO QR open failed: ${e.message}`));
        return;
      }
      if (woQrPrint) {
        printWoQr(woQrPrint).catch((e) => alert(`WO QR print failed: ${e.message}`));
        return;
      }
      if (woQrPng) {
        downloadWoQrPng(woQrPng).catch((e) => alert(`WO QR download failed: ${e.message}`));
        return;
      }
      if (woQrLink) {
        copyWoLink(woQrLink).then(() => alert(`WO #${woQrLink} link copied`)).catch((e) => alert(`WO link copy failed: ${e.message}`));
        return;
      }
      if (id) openCloseModalForRow(id, rowSource);
    });
  }
  if (detailEl) {
    detailEl.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const issueId = target.getAttribute("data-wo-issue-id");
      const assignId = target.getAttribute("data-assign-id");
      const completeId = target.getAttribute("data-complete-id");
      const approveId = target.getAttribute("data-approve-id");
      const setStatusId = target.getAttribute("data-set-status-id");
      const setStatus = target.getAttribute("data-set-status");
      const closeId = target.getAttribute("data-close-id");
      const closeSource = target.getAttribute("data-close-source");
      const saveProgressId = target.getAttribute("data-wo-save-progress");
      const saveCostsId = target.getAttribute("data-wo-save-costs");
      if (issueId) {
        currentDetailWorkOrderId = Number(issueId);
        issueToWorkOrder();
        return;
      }
      if (assignId) {
        openAssignModal(assignId);
        return;
      }
      if (completeId) {
        openCompleteModal(completeId);
        return;
      }
      if (approveId) {
        approveWorkOrder(approveId).catch((e) => alert(`Approve failed: ${e.message}`));
        return;
      }
      if (setStatusId && setStatus) {
        setWorkOrderStatus(setStatusId, setStatus).catch(() => {});
        return;
      }
      if (closeId) openCloseModalForRow(closeId, closeSource);
      if (saveProgressId) saveRepairProgress(saveProgressId).catch(() => {});
      if (saveCostsId) saveRepairCosts(saveCostsId).catch(() => {});
    });
    detailEl.addEventListener("input", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      if (target.id === "woIssueSearch") {
        refreshIssueOptions(target.value || "");
      }
    });
  }

  fetchWorkOrders();
  loadInspectionQuality();
  loadTechnicians().catch(() => {});
  if (isSupervisorRole(getSessionRole())) {
    loadTechnicianRoster().catch(() => {});
    loadRepairAssetOptions().catch(() => {});
    fillTechnicianSelect(document.getElementById("woRepairAssignee"));
  }
});
