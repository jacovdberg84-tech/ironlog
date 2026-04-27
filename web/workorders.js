const API = "/api";
let closingWorkOrderId = null;
let closingWorkOrderSource = "";
let currentDetailWorkOrderId = null;
let stockCatalogCache = [];
const ROLE_KEY = "ironlog_session_role";
const USER_KEY = "ironlog_session_user";
let boardState = [];

function getSessionRole() {
  return String(localStorage.getItem(ROLE_KEY) || "admin").trim().toLowerCase() || "admin";
}

function getSessionUser() {
  return String(localStorage.getItem(USER_KEY) || "admin").trim() || "admin";
}

function authHeaders(extra = {}) {
  return {
    ...extra,
    "x-user-role": getSessionRole(),
    "x-user-name": getSessionUser(),
  };
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

function rolePermissionText(role) {
  const r = String(role || "").toLowerCase();
  if (r === "admin") {
    return "Admin: full control (status transitions, approvals, close, issue parts).";
  }
  if (r === "supervisor") {
    return "Supervisor: full workflow control (assign/start/approve/close) and issue parts.";
  }
  if (r === "artisan") {
    return "Artisan: execution flow only (start, complete, return to progress) and can request close approval. Cannot approve/close or issue parts.";
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
  if (s === "manual") return "Manual";
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

function statusActionButtons(wo) {
  const role = getSessionRole();
  const s = String(wo?.status || "").toLowerCase();
  if (s === "closed") return "";

  const map = {
    open: [{ label: "Assign", to: "assigned" }, { label: "Start", to: "in_progress" }],
    assigned: [{ label: "Start", to: "in_progress" }, { label: "Reopen", to: "open" }],
    in_progress: [{ label: "Complete", to: "completed" }, { label: "Unassign", to: "assigned" }],
    completed: [{ label: "Approve", to: "approved" }, { label: "Back To Progress", to: "in_progress" }],
    approved: [{ label: "Reopen Completed", to: "completed" }],
  };

  const actions = (map[s] || []).filter((a) => canRoleTransition(role, s, a.to));
  return actions
    .map(
      (a) =>
        `<button data-set-status-id="${wo.id}" data-set-status="${a.to}" style="margin-top:8px;">${a.label}</button>`
    )
    .join("");
}

function workOrderCard(wo) {
  const role = getSessionRole();
  const canClose = String(wo.status || "").toLowerCase() !== "closed" && canRoleClose(role);
  const canRequestClose =
    role === "artisan" &&
    ["completed", "approved"].includes(String(wo.status || "").toLowerCase());
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
      <div class="${statusClass(wo.status)}">${String(wo.status || "unknown").toUpperCase()}</div>
      ${statusActionButtons(wo)}
      <button data-pdf-id="${wo.id}" style="margin-top:8px;">Open PDF</button>
      <button data-pdf-download-id="${wo.id}" style="margin-top:8px;">Download PDF</button>
      <button data-view-id="${wo.id}" style="margin-top:8px;">View Detail</button>
      <button data-wo-qr-open="${wo.id}" style="margin-top:8px;">Open QR Page</button>
      <button data-wo-qr-print="${wo.id}" style="margin-top:8px;">Print WO QR</button>
      <button data-wo-qr-png="${wo.id}" style="margin-top:8px;">Download WO QR PNG</button>
      <button data-wo-qr-link="${wo.id}" style="margin-top:8px;">Copy WO Link</button>
      ${canRequestClose ? `<button data-request-close-id="${wo.id}" data-request-close-source="${String(wo.source || "").toLowerCase()}" style="margin-top:8px;">Request Close Approval</button>` : ""}
      ${canClose ? `<button data-close-id="${wo.id}" data-close-source="${String(wo.source || "").toLowerCase()}" style="margin-top:8px;">Close Work Order</button>` : ""}
    </div>
  `;
}

function boardPriorityClass(priority) {
  const p = String(priority || "").toUpperCase();
  if (p === "P1") return "pri-p1";
  if (p === "P2") return "pri-p2";
  return "pri-p3";
}

function boardCard(wo) {
  return `
    <div draggable="true" class="card" style="padding:8px; margin-bottom:8px;" data-board-wo-id="${Number(wo.id)}">
      <div><strong>WO #${Number(wo.id)}</strong> <span class="pill ${boardPriorityClass(wo.priority)}">${String(wo.priority || "P3").toUpperCase()}</span></div>
      <div><small>${wo.asset_code || "-"} - ${wo.asset_name || "-"}</small></div>
      <div><small>Due: ${wo.due_date || "-"}</small></div>
      <div><small>Shift: ${wo.shift || "-"}</small></div>
      <div><small>Skill: ${wo.required_skill || "-"}</small></div>
    </div>
  `;
}

function renderScheduleBoard(rows) {
  const boardEl = document.getElementById("woBoard");
  if (!boardEl) return;
  const list = Array.isArray(rows) ? rows : [];
  const shiftOrder = ["day", "night", "unplanned"];
  const shiftLabel = { day: "Day Shift", night: "Night Shift", unplanned: "Unplanned Shift" };
  const grouped = new Map();
  for (const s of shiftOrder) grouped.set(s, []);
  for (const row of list) {
    const sh = String(row.shift || "").trim().toLowerCase();
    grouped.get(sh === "day" || sh === "night" ? sh : "unplanned").push(row);
  }
  const renderLane = (laneRows) => {
    const artisans = [...new Set(laneRows.map((r) => String(r.assigned_artisan_name || "").trim() || "Unassigned"))];
    const ordered = ["Unassigned", ...artisans.filter((a) => a !== "Unassigned").sort((a, b) => a.localeCompare(b))];
    return ordered
      .map((artisan) => {
        const items = laneRows.filter((r) => (String(r.assigned_artisan_name || "").trim() || "Unassigned") === artisan);
        return `
          <div class="card" style="min-width:260px; max-width:260px;" data-board-column="${artisan.replace(/"/g, "&quot;")}">
            <div style="font-weight:700; margin-bottom:8px;">${artisan}</div>
            <div class="muted" style="margin-bottom:8px;">${items.length} WO(s)</div>
            <div data-board-dropzone="${artisan.replace(/"/g, "&quot;")}" style="min-height:120px;">
              ${items.map(boardCard).join("") || `<small class="muted">Drop work orders here.</small>`}
            </div>
          </div>
        `;
      })
      .join("");
  };
  boardEl.innerHTML = shiftOrder
    .map((s) => `
      <div style="width:100%;">
        <div style="font-weight:700; margin:6px 0;">${shiftLabel[s]}</div>
        <div class="row" style="gap:10px; align-items:stretch; overflow:auto;">
          ${renderLane(grouped.get(s) || [])}
        </div>
      </div>
    `)
    .join("");
}

async function loadScheduleBoard() {
  const msgEl = document.getElementById("woBoardMsg");
  const artisan = String(document.getElementById("woBoardArtisan")?.value || "").trim();
  const shift = String(document.getElementById("woBoardShift")?.value || "").trim();
  const priority = String(document.getElementById("woBoardPriority")?.value || "").trim();
  const due_date = String(document.getElementById("woBoardDueDate")?.value || "").trim();
  const q = new URLSearchParams();
  if (artisan) q.set("artisan", artisan);
  if (shift) q.set("shift", shift);
  if (priority) q.set("priority", priority);
  if (due_date) q.set("due_date", due_date);
  try {
    if (msgEl) {
      msgEl.className = "";
      msgEl.textContent = "Loading scheduling board...";
    }
    const data = await fetchJson(`${API}/workorders/schedule/board?${q.toString()}`, { headers: authHeaders() });
    boardState = Array.isArray(data?.rows) ? data.rows : [];
    renderScheduleBoard(boardState);
    if (msgEl) {
      msgEl.className = "message-success";
      msgEl.textContent = `Board loaded (${boardState.length} work orders).`;
    }
  } catch (err) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message;
    }
  }
}

async function autoAssignWorkOrders() {
  const msgEl = document.getElementById("woBoardMsg");
  try {
    if (msgEl) {
      msgEl.className = "";
      msgEl.textContent = "Auto-assigning...";
    }
    const data = await fetchJson(`${API}/workorders/schedule/auto-assign`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({}),
    });
    await loadScheduleBoard();
    await fetchWorkOrders();
    if (msgEl) {
      msgEl.className = "message-success";
      msgEl.textContent = `Auto-assigned ${Number(data?.assigned_count || 0)} work orders.`;
    }
  } catch (err) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message;
    }
  }
}

async function runEscalationCheck() {
  const msgEl = document.getElementById("woBoardMsg");
  try {
    if (msgEl) {
      msgEl.className = "";
      msgEl.textContent = "Checking escalations...";
    }
    const overdueHours = Math.max(1, Number(document.getElementById("woEscHours")?.value || 8));
    const data = await fetchJson(`${API}/workorders/schedule/escalations/check`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ overdue_hours: overdueHours }),
    });
    if (msgEl) {
      msgEl.className = "message-success";
      msgEl.textContent = `Escalations triggered: ${Number(data?.escalated_count || 0)} (alerts ${Number(data?.notification_count || 0)}, threshold ${Number(data?.threshold_hours || 0)}h).`;
    }
    await loadEscalations();
  } catch (err) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message;
    }
  }
}

function esc(s) {
  return String(s ?? "").replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[m]));
}

async function loadAssignmentRules() {
  const listEl = document.getElementById("woRulesList");
  const msgEl = document.getElementById("woRulesMsg");
  if (!listEl) return;
  try {
    const data = await fetchJson(`${API}/workorders/schedule/rules`, { headers: authHeaders() });
    const rows = Array.isArray(data?.rows) ? data.rows : [];
    listEl.innerHTML = rows.length
      ? rows.map((r) => `
        <div class="item row" style="justify-content:space-between; gap:8px; margin-bottom:6px;">
          <div>
            <strong>${esc(r.artisan_name)}</strong> | skill: ${esc(r.skill || "any")} | loc: ${esc(r.location_code || "any")} | shift: ${esc(r.shift || "any")} | max: ${Number(r.max_open_wos || 0)}
          </div>
          <button data-rule-del="${Number(r.id)}" type="button">Delete</button>
        </div>
      `).join("")
      : `<small class="muted">No assignment rules configured.</small>`;
    if (msgEl) msgEl.textContent = `${rows.length} rule(s) loaded.`;
  } catch (err) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message;
    }
  }
}

async function saveAssignmentRule() {
  const msgEl = document.getElementById("woRulesMsg");
  const artisan_name = String(document.getElementById("woRuleArtisan")?.value || "").trim();
  const skill = String(document.getElementById("woRuleSkill")?.value || "").trim();
  const location_code = String(document.getElementById("woRuleLocation")?.value || "").trim();
  const shift = String(document.getElementById("woRuleShift")?.value || "").trim();
  const max_open_wos = Math.max(1, Number(document.getElementById("woRuleMaxLoad")?.value || 8));
  if (!artisan_name) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = "Artisan name is required.";
    }
    return;
  }
  try {
    await fetchJson(`${API}/workorders/schedule/rules`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ artisan_name, skill, location_code, shift, max_open_wos }),
    });
    if (msgEl) {
      msgEl.className = "message-success";
      msgEl.textContent = "Rule saved.";
    }
    await loadAssignmentRules();
  } catch (err) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message;
    }
  }
}

async function deleteAssignmentRule(id) {
  const msgEl = document.getElementById("woRulesMsg");
  try {
    await fetchJson(`${API}/workorders/schedule/rules/${Number(id)}/delete`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({}),
    });
    await loadAssignmentRules();
  } catch (err) {
    if (msgEl) {
      msgEl.className = "message-error";
      msgEl.textContent = err.message;
    }
  }
}

async function loadEscalationConfig() {
  const msg = document.getElementById("woEscMsg");
  try {
    const data = await fetchJson(`${API}/workorders/schedule/escalation-config`, { headers: authHeaders() });
    const c = data?.config || {};
    const set = (id, v) => {
      const el = document.getElementById(id);
      if (el) el.value = String(v ?? "");
    };
    set("woEscHours", c.overdue_hours || 8);
    set("woEscRole1", c.level1_role || "supervisor");
    set("woEscRole2", c.level2_role || "manager");
    set("woEscRole3", c.level3_role || "admin");
    if (msg) {
      msg.className = "muted";
      msg.textContent = "Escalation config loaded.";
    }
  } catch (err) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = err.message;
    }
  }
}

async function saveEscalationConfig() {
  const msg = document.getElementById("woEscMsg");
  const payload = {
    overdue_hours: Math.max(1, Number(document.getElementById("woEscHours")?.value || 8)),
    level1_role: String(document.getElementById("woEscRole1")?.value || "supervisor").trim(),
    level2_role: String(document.getElementById("woEscRole2")?.value || "manager").trim(),
    level3_role: String(document.getElementById("woEscRole3")?.value || "admin").trim(),
  };
  try {
    await fetchJson(`${API}/workorders/schedule/escalation-config`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify(payload),
    });
    if (msg) {
      msg.className = "message-success";
      msg.textContent = "Escalation chain saved.";
    }
  } catch (err) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = err.message;
    }
  }
}

async function loadEscalations() {
  const el = document.getElementById("woEscList");
  if (!el) return;
  try {
    const data = await fetchJson(`${API}/workorders/schedule/escalations`, { headers: authHeaders() });
    const rows = Array.isArray(data?.rows) ? data.rows : [];
    el.innerHTML = rows.length
      ? rows.slice(0, 12).map((r) => `
        <div class="item" style="margin-bottom:6px;">
          WO #${Number(r.work_order_id)} (${esc(r.asset_code || "-")}) | L${Number(r.chain_level || 1)} | overdue>${Number(r.threshold_hours || 0)}h | status: ${esc(r.status || "-")}
        </div>
      `).join("")
      : `<small class="muted">No escalation events yet.</small>`;
  } catch (err) {
    el.innerHTML = `<div class="message-error">${esc(err.message)}</div>`;
  }
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
          <div><strong>Opened:</strong> ${wo.opened_at || "-"}</div>
          <div><strong>Closed:</strong> ${wo.closed_at || "-"}</div>
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
    detailEl.innerHTML = renderDetail(data);
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

async function setWorkOrderStatus(id, status) {
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
      body: JSON.stringify({ status: next })
    });
    const data = await res.json();
    if (!res.ok) {
      throw new Error(data.error || "Failed to update status");
    }
    await fetchWorkOrders();
    await loadWorkOrderDetail(woId);
  } catch (err) {
    console.error("Set status error:", err);
    alert(`Could not update status: ${err.message}`);
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
  const boardRefreshBtn = document.getElementById("woBoardRefreshBtn");
  const autoAssignBtn = document.getElementById("woAutoAssignBtn");
  const escalationsBtn = document.getElementById("woEscalationsBtn");
  const ruleSaveBtn = document.getElementById("woRuleSaveBtn");
  const rulesList = document.getElementById("woRulesList");
  const escSaveBtn = document.getElementById("woEscSaveBtn");

  const roleBadge = document.getElementById("woRoleBadge");
  if (roleBadge) roleBadge.textContent = `Role: ${role}`;
  const legend = document.getElementById("woPermissionLegend");
  if (legend) legend.textContent = rolePermissionText(role);

  if (!["admin", "supervisor"].includes(role) && awaitingApprovalBtn) {
    awaitingApprovalBtn.style.display = "none";
  }

  if (refreshBtn) refreshBtn.addEventListener("click", fetchWorkOrders);
  if (boardRefreshBtn) boardRefreshBtn.addEventListener("click", loadScheduleBoard);
  if (autoAssignBtn) autoAssignBtn.addEventListener("click", autoAssignWorkOrders);
  if (escalationsBtn) escalationsBtn.addEventListener("click", runEscalationCheck);
  if (ruleSaveBtn) ruleSaveBtn.addEventListener("click", saveAssignmentRule);
  if (escSaveBtn) escSaveBtn.addEventListener("click", saveEscalationConfig);
  if (rulesList) {
    rulesList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const delId = target.getAttribute("data-rule-del");
      if (delId) deleteAssignmentRule(delId);
    });
  }
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
        setWorkOrderStatus(setStatusId, setStatus);
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
  const boardEl = document.getElementById("woBoard");
  if (boardEl) {
    boardEl.addEventListener("dragstart", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const id = target.getAttribute("data-board-wo-id");
      if (!id) return;
      evt.dataTransfer?.setData("text/plain", id);
    });
    boardEl.addEventListener("dragover", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      if (target.closest("[data-board-dropzone]")) evt.preventDefault();
    });
    boardEl.addEventListener("drop", async (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const zone = target.closest("[data-board-dropzone]");
      if (!(zone instanceof HTMLElement)) return;
      evt.preventDefault();
      const woId = Number(evt.dataTransfer?.getData("text/plain") || 0);
      const artisan = String(zone.getAttribute("data-board-dropzone") || "").trim();
      if (!woId || !artisan) return;
      try {
        await fetchJson(`${API}/workorders/${woId}/schedule`, {
          method: "POST",
          headers: authHeaders({ "Content-Type": "application/json" }),
          body: JSON.stringify({ assigned_artisan_name: artisan === "Unassigned" ? null : artisan }),
        });
        await loadScheduleBoard();
        await fetchWorkOrders();
      } catch (e) {
        alert(`Reassign failed: ${e.message}`);
      }
    });
  }
  if (detailEl) {
    detailEl.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const issueId = target.getAttribute("data-wo-issue-id");
      if (issueId) {
        currentDetailWorkOrderId = Number(issueId);
        issueToWorkOrder();
      }
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
  loadScheduleBoard();
  loadInspectionQuality();
  loadAssignmentRules();
  loadEscalationConfig();
  loadEscalations();
});
