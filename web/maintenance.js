const API = "/api";
let lastSyncPull = { last_id: 0, events: [] };
const ROLE_KEY = "ironlog_session_role";
const ROLES_KEY = "ironlog_session_roles";
const USER_KEY = "ironlog_session_user";
const SITE_KEY = "ironlog_session_site";
const TOKEN_KEY = "ironlog_auth_token";
const MAINT_DUE_THRESHOLD_KEY = "ironlog_maintenance_due_threshold_hours";
const MAINT_LOCK_KEY = "ironlog_maintenance_access_ok";
const MAINT_LOCK_USER = "BJ van den Berg";
const MAINT_LOCK_PASSWORD = "0mhliac789";

function ensureMaintenanceAccess() {
  if (sessionStorage.getItem(MAINT_LOCK_KEY) === "1") return true;
  const user = String(window.prompt("Maintenance username:") || "").trim();
  const pass = String(window.prompt("Maintenance password:") || "");
  const ok = user === MAINT_LOCK_USER && pass === MAINT_LOCK_PASSWORD;
  if (ok) {
    sessionStorage.setItem(MAINT_LOCK_KEY, "1");
    return true;
  }
  alert("Access denied.");
  location.href = "index.html";
  return false;
}

function getSessionRole() {
  return String(localStorage.getItem(ROLE_KEY) || "admin").trim().toLowerCase() || "admin";
}
function getSessionRoles() {
  try {
    const parsed = JSON.parse(String(localStorage.getItem(ROLES_KEY) || "[]"));
    if (Array.isArray(parsed) && parsed.length) {
      return Array.from(
        new Set(parsed.map((r) => String(r || "").trim().toLowerCase()).filter(Boolean))
      );
    }
  } catch {}
  return [getSessionRole()];
}
function getSessionUser() {
  return String(localStorage.getItem(USER_KEY) || "admin").trim() || "admin";
}
function getSessionSite() {
  return String(localStorage.getItem(SITE_KEY) || "main").trim().toLowerCase() || "main";
}
function getAuthToken() {
  return String(localStorage.getItem(TOKEN_KEY) || sessionStorage.getItem(TOKEN_KEY) || "").trim();
}
function authHeaders(extra = {}) {
  const h = {
    ...extra,
    "x-user-name": getSessionUser(),
    "x-user-role": getSessionRole(),
    "x-user-roles": getSessionRoles().join(","),
    "x-site-code": getSessionSite(),
  };
  const tok = getAuthToken();
  if (tok) h.Authorization = `Bearer ${tok}`;
  return h;
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

function shouldAutoFillLastServiceHours() {
  const useLiveEl = document.getElementById("planUseLiveForLastService");
  return !useLiveEl || Boolean(useLiveEl.checked);
}

function syncLastServiceHoursFromLive() {
  if (!shouldAutoFillLastServiceHours()) return;

  const currentHoursEl = document.getElementById("planCurrentHours");
  const lastServiceEl = document.getElementById("planLastServiceHours");
  if (!currentHoursEl || !lastServiceEl) return;

  const current = Number(currentHoursEl.value || 0);
  lastServiceEl.value = Number.isFinite(current) ? current.toFixed(1) : "0";
}

function planCard(p) {
  return `
    <div class="card">
      <div><strong>${p.asset_code}</strong> - ${p.asset_name}</div>
      <div><strong>Service:</strong> ${p.service_name}</div>
      <div><strong>Interval:</strong> ${Number(p.interval_hours || 0).toFixed(1)} hrs</div>
      <div><strong>Last Service:</strong> ${Number(p.last_service_hours || 0).toFixed(1)} hrs</div>
      <div><strong>Active:</strong> ${Number(p.active || 0) ? "Yes" : "No"}</div>
    </div>
  `;
}

function dueCard(d) {
  let statusClass = "status-ok";
  let statusText = "OK";

  if (d.is_overdue) {
    statusClass = "status-overdue";
    statusText = "OVERDUE";
  } else if (String(d.status || "").toUpperCase() === "ALMOST DUE") {
    statusClass = "status-soon";
    statusText = "ALMOST DUE";
  }

  return `
    <div class="card">
      <div><strong>${d.asset_code}</strong> - ${d.asset_name}</div>
      <div><strong>Service:</strong> ${d.service_name}</div>
      <div><strong>Current Hours:</strong> ${Number(d.current_hours || 0).toFixed(1)}</div>
      <div><strong>Next Due:</strong> ${Number(d.next_due_hours || 0).toFixed(1)}</div>
      <div><strong>Remaining:</strong> ${Number(d.remaining_hours || 0).toFixed(1)}</div>
      <div class="${statusClass}">${statusText}</div>
    </div>
  `;
}

function fmt1(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n.toFixed(1) : "-";
}

function histRow(r) {
  const eq = `${r.asset_code || "-"} - ${r.asset_name || "-"}`;
  const last = r.last_serviced_date || "-";
  const est = r.estimated_service_date || "-";
  const hrsToNext = Number(r.remaining_hours || 0);
  const warn = hrsToNext <= 0 ? "status-overdue" : hrsToNext <= 50 ? "status-soon" : "status-ok";
  const sourceMap = {
    daily_closing: "Daily closing",
    asset_hours: "Asset hours",
    daily_sum: "Daily sum",
  };
  const src = sourceMap[String(r.current_hours_source || "").trim()] || "Unknown";
  return `
    <tr class="${hrsToNext <= 0 ? "downRow" : ""}">
      <td><b>${eq}</b></td>
      <td>${r.service_name || "-"}</td>
      <td>${last}</td>
      <td style="text-align:right;">${fmt1(r.current_hours)}<br><small class="muted">(${src})</small></td>
      <td style="text-align:right;"><span class="${warn}">${fmt1(r.remaining_hours)}</span></td>
      <td style="text-align:right;">${fmt1(r.avg_daily_hours)}</td>
      <td>${est}</td>
    </tr>
  `;
}

function escBackfill(s) {
  return String(s == null ? "" : s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function backfillRow(r) {
  const hours = r.service_hours == null ? "-" : Number(r.service_hours).toFixed(1);
  return `
    <tr data-backfill-id="${Number(r.id || 0)}">
      <td>${escBackfill(r.asset_code || "-")} - ${escBackfill(r.asset_name || "-")}</td>
      <td>${escBackfill(r.service_name || "-")}</td>
      <td>${escBackfill(r.service_date || "-")}</td>
      <td style="text-align:right;">${hours}</td>
      <td>${escBackfill(r.notes || "")}</td>
      <td>
        <button data-backfill-action="edit">Edit</button>
        <button data-backfill-action="delete">Delete</button>
      </td>
    </tr>
  `;
}

function inspectproRow(r) {
  const status = String(r.status || "").toLowerCase();
  const cls = status === "ok" ? "status-ok" : status === "error" ? "status-overdue" : "status-soon";
  const when = String(r.updated_at || r.created_at || "-");
  return `
    <tr>
      <td>${Number(r.id || 0)}</td>
      <td>${escBackfill(when)}</td>
      <td>${escBackfill(r.event_type || "-")}</td>
      <td>${escBackfill(r.asset_code || "-")}</td>
      <td><span class="${cls}">${escBackfill(status || "-")}</span></td>
      <td>${r.target_id == null ? "-" : Number(r.target_id)}</td>
      <td>${escBackfill(r.error_message || "")}</td>
    </tr>
  `;
}

async function loadHistory() {
  const body = document.getElementById("histBody");
  const meta = document.getElementById("histMeta");
  const modeEl = document.getElementById("histViewMode");
  const limitEl = document.getElementById("histClosestLimit");
  const limitWrap = document.getElementById("histClosestLimitWrap");
  if (!body) return;
  body.innerHTML = `<tr><td colspan="7" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/maintenance/history`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load history");
    const mode = String(modeEl?.value || "all").trim().toLowerCase();
    if (limitWrap) limitWrap.style.display = mode === "closest" ? "" : "none";
    const rawRows = Array.isArray(data.rows) ? data.rows : [];
    let rows = rawRows.slice();
    if (mode === "closest") {
      rows.sort((a, b) => Number(a?.remaining_hours || 0) - Number(b?.remaining_hours || 0));
      const n = Number(limitEl?.value || 20);
      const topN = Number.isFinite(n) ? Math.max(1, Math.min(200, Math.trunc(n))) : 20;
      rows = rows.slice(0, topN);
    }
    if (meta) {
      const suffix = mode === "closest" ? ` | Showing closest due (${rows.length}/${rawRows.length})` : "";
      meta.textContent = `As of: ${data.as_of || "-"}${suffix}`;
    }
    body.innerHTML = rows.length ? rows.map(histRow).join("") : `<tr><td colspan="7" class="muted">No history rows.</td></tr>`;
  } catch (e) {
    body.innerHTML = `<tr><td colspan="7" class="message-error">History load error: ${e.message || e}</td></tr>`;
  }
}

async function loadBackfillHistory() {
  const body = document.getElementById("backfillBody");
  if (!body) return;
  body.innerHTML = `<tr><td colspan="6" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/maintenance/history/backfill?limit=20`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load backfill history");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    body.innerHTML = rows.length
      ? rows.map(backfillRow).join("")
      : `<tr><td colspan="6" class="muted">No historical entries yet.</td></tr>`;
  } catch (e) {
    body.innerHTML = `<tr><td colspan="6" class="message-error">${escBackfill(e.message || e)}</td></tr>`;
  }
}

async function loadInspectproStatus() {
  const body = document.getElementById("inspectproStatusBody");
  const meta = document.getElementById("inspectproStatusMeta");
  if (!body) return;
  body.innerHTML = `<tr><td colspan="7" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/integrations/inspectpro/status?limit=20`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load InspectPro status");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    body.innerHTML = rows.length
      ? rows.map(inspectproRow).join("")
      : `<tr><td colspan="7" class="muted">No InspectPro events yet.</td></tr>`;
    if (meta) meta.textContent = `Latest update: ${new Date().toLocaleString()}`;
  } catch (e) {
    body.innerHTML = `<tr><td colspan="7" class="message-error">${escBackfill(e.message || e)}</td></tr>`;
    if (meta) meta.textContent = "Status load failed.";
  }
}

async function loadAssetsForPlan() {
  const select = document.getElementById("planAsset");
  const backfillSelect = document.getElementById("backfillAsset");
  if (!select) return;

  select.innerHTML = `<option value="">Loading assets...</option>`;
  if (backfillSelect) backfillSelect.innerHTML = `<option value="">Loading assets...</option>`;

  try {
    const res = await fetch(`${API}/assets`);
    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || "Failed to load assets");
    }

    const assets = Array.isArray(data)
      ? data
      : Array.isArray(data?.assets)
        ? data.assets
        : Array.isArray(data?.rows)
          ? data.rows
          : Array.isArray(data?.data)
            ? data.data
            : [];

    if (!Array.isArray(assets) || !assets.length) {
      select.innerHTML = `<option value="">No assets found</option>`;
      return;
    }

    const validAssets = assets.filter((a) => {
      const idOk = Number.isInteger(Number(a?.id)) && Number(a.id) > 0;
      if (!idOk) return false;
      const active = Number(a?.active ?? 1) !== 0;
      const archived = Number(a?.archived ?? 0) === 1;
      return active && !archived;
    });

    if (!validAssets.length) {
      select.innerHTML = `<option value="">No valid assets found</option>`;
      return;
    }

    select.innerHTML = `
      <option value="">Select asset</option>
      ${validAssets.map(a => `
        <option value="${a.id}">
          ${a.asset_code || "NO-CODE"} - ${a.asset_name || "Unnamed Asset"}
        </option>
      `).join("")}
    `;
    if (backfillSelect) {
      backfillSelect.innerHTML = `
        <option value="">Select asset</option>
        ${validAssets.map(a => `
          <option value="${a.id}">
            ${a.asset_code || "NO-CODE"} - ${a.asset_name || "Unnamed Asset"}
          </option>
        `).join("")}
      `;
    }
  } catch (err) {
    console.error("Assets load error:", err);
    select.innerHTML = `<option value="">Failed to load assets</option>`;
    if (backfillSelect) backfillSelect.innerHTML = `<option value="">Failed to load assets</option>`;
  }
}

async function saveBackfillHistory() {
  const assetEl = document.getElementById("backfillAsset");
  const serviceEl = document.getElementById("backfillServiceName");
  const dateEl = document.getElementById("backfillServiceDate");
  const hoursEl = document.getElementById("backfillServiceHours");
  const notesEl = document.getElementById("backfillNotes");
  const updPlanEl = document.getElementById("backfillUpdatePlanHours");
  const msgEl = document.getElementById("backfillMsg");
  if (!assetEl || !serviceEl || !dateEl || !hoursEl || !notesEl || !updPlanEl || !msgEl) return;

  const asset_id = Number(assetEl.value || 0);
  const service_name = String(serviceEl.value || "").trim();
  const service_date = String(dateEl.value || "").trim();
  const service_hours_raw = String(hoursEl.value || "").trim();
  const notes = String(notesEl.value || "").trim();
  const update_plan_last_hours = updPlanEl.checked ? 1 : 0;

  if (!asset_id || !service_name || !service_date) {
    msgEl.className = "message-error";
    msgEl.textContent = "Asset, service name, and service date are required.";
    return;
  }

  const payload = {
    asset_id,
    service_name,
    service_date,
    service_hours: service_hours_raw === "" ? null : Number(service_hours_raw),
    notes,
    update_plan_last_hours,
  };

  msgEl.className = "muted";
  msgEl.textContent = "Saving historical service...";
  try {
    const res = await fetch(`${API}/maintenance/history/backfill`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save historical service");

    msgEl.className = "message-success";
    msgEl.textContent = data.plan_last_hours_updated
      ? "Historical service saved and plan last-service hours updated."
      : "Historical service saved.";
    serviceEl.value = "";
    hoursEl.value = "";
    notesEl.value = "";
    await loadHistory();
    await loadPlans();
    await loadDue();
    await loadBackfillHistory();
  } catch (err) {
    msgEl.className = "message-error";
    msgEl.textContent = err.message || String(err);
  }
}

async function editBackfillHistory(id) {
  const iid = Number(id || 0);
  if (!iid) return;
  const service_name = prompt("Service name:");
  if (service_name == null) return;
  const service_date = prompt("Service date (YYYY-MM-DD):");
  if (service_date == null) return;
  const service_hours = prompt("Service hours (blank for none):", "");
  if (service_hours == null) return;
  const notes = prompt("Notes (optional):", "");
  if (notes == null) return;
  const payload = {
    service_name: String(service_name || "").trim(),
    service_date: String(service_date || "").trim(),
    service_hours: String(service_hours || "").trim() === "" ? null : Number(service_hours),
    notes: String(notes || "").trim(),
  };
  const res = await fetch(`${API}/maintenance/history/backfill/${iid}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to update historical entry");
  await loadBackfillHistory();
  await loadHistory();
}

async function deleteBackfillHistory(id) {
  const iid = Number(id || 0);
  if (!iid) return;
  if (!confirm("Delete this historical service entry?")) return;
  const res = await fetch(`${API}/maintenance/history/backfill/${iid}`, { method: "DELETE" });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to delete historical entry");
  await loadBackfillHistory();
  await loadHistory();
}

async function loadLiveHoursForSelectedAsset() {
  const assetEl = document.getElementById("planAsset");
  const currentHoursEl = document.getElementById("planCurrentHours");
  const currentHoursSrcEl = document.getElementById("planCurrentHoursSource");

  if (!assetEl || !currentHoursEl) return;

  const assetId = Number(assetEl.value || 0);

  if (!assetId) {
    currentHoursEl.value = "0";
    if (currentHoursSrcEl) currentHoursSrcEl.textContent = "Source: -";
    syncLastServiceHoursFromLive();
    return;
  }

  currentHoursEl.value = "0";
  currentHoursEl.placeholder = "Loading...";

  try {
    const res = await fetch(`${API}/maintenance/asset/${assetId}/live-hours`);
    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || "Failed to load live hours");
    }

    currentHoursEl.value = Number(data.current_hours || 0).toFixed(1);
    currentHoursEl.placeholder = "";
    const sourceMap = {
      daily_closing: "Daily closing",
      asset_hours: "Asset hours",
      daily_sum: "Daily sum",
    };
    const src = sourceMap[String(data.current_hours_source || "").trim()] || "Unknown";
    if (currentHoursSrcEl) currentHoursSrcEl.textContent = `Source: ${src}`;
    syncLastServiceHoursFromLive();
  } catch (err) {
    console.error("Live hours load error:", err);
    // Don’t overwrite user input with 0 when live hours fails
    currentHoursEl.value = "";
    currentHoursEl.placeholder = "";
    if (currentHoursSrcEl) currentHoursSrcEl.textContent = "Source: -";
  }
}

async function savePlan() {
  const assetEl = document.getElementById("planAsset");
  const serviceNameEl = document.getElementById("planServiceName");
  const intervalEl = document.getElementById("planIntervalHours");
  const lastServiceEl = document.getElementById("planLastServiceHours");
  const activeEl = document.getElementById("planActive");
  const msgEl = document.getElementById("planFormMessage");

  if (!assetEl || !serviceNameEl || !intervalEl || !lastServiceEl || !activeEl || !msgEl) {
    console.error("Plan form elements missing");
    return;
  }

  const payload = {
    asset_id: Number(assetEl.value || 0),
    service_name: serviceNameEl.value.trim(),
    interval_hours: Number(intervalEl.value || 0),
    last_service_hours: Number(lastServiceEl.value || 0),
    active: Number(activeEl.value || 1)
  };

  if (!payload.asset_id || !payload.service_name || payload.interval_hours <= 0) {
    msgEl.className = "message-error";
    msgEl.textContent = "Please select an asset, enter a service name, and set interval hours.";
    return;
  }

  msgEl.className = "";
  msgEl.textContent = "Saving plan...";

  try {
    const res = await fetch(`${API}/maintenance/plans`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || "Failed to save maintenance plan");
    }

    msgEl.className = "message-success";
    msgEl.textContent = "Maintenance plan saved successfully.";

    serviceNameEl.value = "";
    intervalEl.value = "";
    lastServiceEl.value = "0";
    activeEl.value = "1";

    await loadPlans();
    await loadDue();
    await loadHistory();
  } catch (err) {
    console.error("Save plan error:", err);
    msgEl.className = "message-error";
    msgEl.textContent = err.message;
  }
}

async function loadPlans() {
  const container = document.getElementById("plansList");
  container.innerHTML = `<div class="skeleton-block"></div><div class="skeleton-block"></div>`;

  try {
    const res = await fetch(`${API}/maintenance/plans`);
    const data = await res.json();
    console.log("Plans response:", data);

    if (!res.ok) {
      throw new Error(data.error || "Failed to load plans");
    }

    const plans = Array.isArray(data.plans) ? data.plans : [];
    container.innerHTML = plans.length
      ? plans.map(planCard).join("")
      : "<div>No maintenance plans found.</div>";
  } catch (err) {
    console.error("Plans error:", err);
    container.innerHTML = `<div style="color:#ff8080;">Error loading plans: ${err.message}</div>`;
  }
}

async function loadDue() {
  const container = document.getElementById("dueList");
  container.innerHTML = `<div class="skeleton-block"></div><div class="skeleton-block"></div>`;
  const nearDueHours = getDueThresholdHours();

  try {
    const q = new URLSearchParams();
    q.set("near_due_hours", String(nearDueHours));
    const res = await fetch(`${API}/maintenance/due?${q.toString()}`);
    const data = await res.json();
    console.log("Due response:", data);

    if (!res.ok) {
      throw new Error(data.error || "Failed to load due services");
    }
    

    const due = Array.isArray(data.due) ? data.due : [];
    due.sort((a, b) => {
      const getRank = (d) => {
        const remaining = Number(d.remaining_hours || 0);
        if (remaining <= 0) return 1;
        if (remaining <= nearDueHours) return 2;
        return 3;
      };

      const rankDiff = getRank(a) - getRank(b);
      if (rankDiff !== 0) return rankDiff;
      return Number(a.remaining_hours || 0) - Number(b.remaining_hours || 0);
    });
    container.innerHTML = due.length
      ? due.map(dueCard).join("")
      : "<div>No due services found.</div>";
  } catch (err) {
    console.error("Due error:", err);
    container.innerHTML = `<div style="color:#ff8080;">Error loading due services: ${err.message}</div>`;
  }
}

async function openUpcomingServicesPdf(download = false) {
  const nearDueHours = getDueThresholdHours();
  const q = new URLSearchParams();
  q.set("near_due_hours", String(nearDueHours));
  const url = `${API}/maintenance/due-upcoming.pdf?${q.toString()}`;
  try {
    const res = await fetch(url, { headers: authHeaders() });
    if (!res.ok) {
      const t = await res.text();
      throw new Error(t || `PDF request failed (${res.status})`);
    }
    const blob = await res.blob();
    const blobUrl = URL.createObjectURL(blob);
    if (download) {
      const a = document.createElement("a");
      const dateTag = new Date().toISOString().slice(0, 10);
      a.href = blobUrl;
      a.download = `maintenance-upcoming-services-${dateTag}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(blobUrl), 3000);
      return;
    }
    window.open(blobUrl, "_blank");
    setTimeout(() => URL.revokeObjectURL(blobUrl), 15000);
  } catch (err) {
    alert(`Could not open PDF: ${err.message || err}`);
  }
}

function mpWeekRangeLabel() {
  const d = new Date();
  const day = d.getDay();
  const mondayOffset = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + mondayOffset);
  const start = d.toISOString().slice(0, 10);
  d.setDate(d.getDate() + 6);
  const end = d.toISOString().slice(0, 10);
  return { start, end };
}

function mpMonthLabel() {
  return new Date().toISOString().slice(0, 7);
}

function getMpSelectedRange(reportType) {
  const type = String(reportType || "").toLowerCase() === "monthly" ? "monthly" : "weekly";
  if (type === "monthly") {
    const month = String(document.getElementById("mpMonth")?.value || "").trim() || mpMonthLabel();
    return { month };
  }
  const start = String(document.getElementById("mpWeekStart")?.value || "").trim();
  const end = String(document.getElementById("mpWeekEnd")?.value || "").trim();
  if (start && end) return { start, end };
  return mpWeekRangeLabel();
}

async function mpGenerate(reportType) {
  const msg = document.getElementById("mpStatusMsg");
  const type = String(reportType || "").toLowerCase() === "monthly" ? "monthly" : "weekly";
  if (msg) {
    msg.className = "muted";
    msg.textContent = `Generating ${type} presentation...`;
  }
  const body = { period_type: type };
  const sel = getMpSelectedRange(type);
  if (type === "monthly") body.month = sel.month;
  else {
    body.start = sel.start;
    body.end = sel.end;
  }
  try {
    const res = await fetch(`${API}/reports/maintenance-master/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify(body),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Generation failed");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `${type} presentation generated (${data.label}).`;
    }
    await loadMaintenancePackStatus();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Generate error: ${e.message || e}`;
    }
  }
}

function openMaintenancePackLatest(reportType, download = false) {
  const type = String(reportType || "").toLowerCase() === "monthly" ? "monthly" : "weekly";
  const q = new URLSearchParams();
  q.set("period_type", type);
  const sel = getMpSelectedRange(type);
  if (type === "monthly") q.set("month", sel.month);
  else {
    q.set("start", sel.start);
    q.set("end", sel.end);
  }
  if (download) q.set("download", "1");
  window.open(`${API}/reports/maintenance-master/latest.pptx?${q.toString()}`, "_blank");
}

async function loadMaintenancePackStatus() {
  const body = document.getElementById("mpStatusBody");
  const msg = document.getElementById("mpStatusMsg");
  if (!body) return;
  body.innerHTML = `<tr><td colspan="5" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/reports/maintenance-master/status`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Status load failed");
    const weekly = data?.latest?.weekly || null;
    const monthly = data?.latest?.monthly || null;
    const rowHtml = (label, key, r) => `
      <tr>
        <td>${esc(label)}</td>
        <td>${esc(r?.period_start && r?.period_end ? `${r.period_start} to ${r.period_end}` : "-")}</td>
        <td>${esc(r?.generated_at || "-")}</td>
        <td>${esc(r?.status || "not generated")}</td>
        <td style="display:flex; gap:8px;">
          <button type="button" data-mp-gen="${key}">Generate</button>
          <button type="button" data-mp-open="${key}">Open Latest</button>
          <button type="button" data-mp-download="${key}">Download Latest</button>
        </td>
      </tr>
    `;
    body.innerHTML = rowHtml("Weekly", "weekly", weekly) + rowHtml("Monthly", "monthly", monthly);
    if (msg) {
      msg.className = "muted";
      msg.textContent = "Status loaded.";
    }
  } catch (e) {
    body.innerHTML = `<tr><td colspan="5" class="message-error">${esc(e.message || String(e))}</td></tr>`;
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Status error: ${e.message || e}`;
    }
  }
}

function getDueThresholdHours() {
  const input = document.getElementById("dueNearThresholdHours");
  const fromInput = Number(input?.value || 50);
  const fallbackSaved = Number(localStorage.getItem(MAINT_DUE_THRESHOLD_KEY) || 50);
  const v = Number.isFinite(fromInput) && fromInput > 0
    ? fromInput
    : (Number.isFinite(fallbackSaved) && fallbackSaved > 0 ? fallbackSaved : 50);
  return Math.max(1, Math.round(v));
}

function syncDueThresholdInput() {
  const input = document.getElementById("dueNearThresholdHours");
  if (!input) return;
  const saved = Number(localStorage.getItem(MAINT_DUE_THRESHOLD_KEY) || 50);
  const value = Number.isFinite(saved) && saved > 0 ? Math.round(saved) : 50;
  input.value = String(value);
}

async function generateWO() {
  const generateBtn = document.getElementById("generateBtn");
  if (generateBtn && generateBtn.disabled) return;

  try {
    if (generateBtn) {
      generateBtn.disabled = true;
      generateBtn.textContent = "Generating...";
    }

    const res = await fetch(`${API}/maintenance/generate`, { method: "POST" });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to generate work orders");
    alert(`Created ${Number(data.created_count || 0)} work orders`);
    await loadDue();
    await loadHistory();
  } catch (err) {
    console.error("Generate error:", err);
    alert(`Generate failed: ${err.message}`);
  } finally {
    if (generateBtn) {
      generateBtn.disabled = false;
      generateBtn.textContent = "Generate Work Orders";
    }
  }
}

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

function syncSetMsg(text, isError = false) {
  const el = document.getElementById("syncMsg");
  if (!el) return;
  el.className = isError ? "message-error" : "muted";
  el.textContent = text || "";
}

function syncSetOutput(obj) {
  const el = document.getElementById("syncOutput");
  if (!el) return;
  try {
    el.textContent = JSON.stringify(obj ?? {}, null, 2);
  } catch {
    el.textContent = String(obj ?? "");
  }
}

function syncHeaders() {
  return authHeaders({ "Content-Type": "application/json" });
}

function setTopView(view) {
  const main = document.getElementById("mainMaintenanceCards");
  const mi = document.getElementById("managerInspectionsCard");
  const ai = document.getElementById("artisanInspectionsCard");
  const wf = document.getElementById("weeklyForumCard");
  const kpi = document.getElementById("assetKpiCard");
  const hist = document.getElementById("histogramCard");
  const sync = document.getElementById("syncAdminCard");
  const btnMain = document.getElementById("showMainMaintBtn");
  const btnMi = document.getElementById("showManagerInspectionsBtn");
  const btnAi = document.getElementById("showArtisanInspectionsBtn");
  const btnWf = document.getElementById("showWeeklyForumBtn");
  const btnKpi = document.getElementById("showAssetKpiBtn");
  const btnHist = document.getElementById("showHistogramBtn");
  const btnSync = document.getElementById("showSyncAdminBtn");
  if (!main || !mi || !wf || !sync) return;

  main.style.display = view === "main" ? "" : "none";
  mi.style.display = view === "mi" ? "" : "none";
  if (ai) ai.style.display = view === "ai" ? "" : "none";
  wf.style.display = view === "wf" ? "" : "none";
  if (kpi) kpi.style.display = view === "kpi" ? "" : "none";
  if (hist) hist.style.display = view === "hist" ? "" : "none";
  sync.style.display = view === "sync" ? "" : "none";

  const styleBtn = (btn, active) => {
    if (!btn) return;
    btn.style.borderColor = active ? "#3b82f6" : "";
    btn.style.background = active ? "#13233c" : "";
    btn.style.color = active ? "#fff" : "";
  };
  styleBtn(btnMain, view === "main");
  styleBtn(btnMi, view === "mi");
  styleBtn(btnAi, view === "ai");
  styleBtn(btnWf, view === "wf");
  styleBtn(btnKpi, view === "kpi");
  styleBtn(btnHist, view === "hist");
  styleBtn(btnSync, view === "sync");
}

async function loadHistogramEvents() {
  const body = document.getElementById("histEventBody");
  const msg = document.getElementById("histMsg");
  if (!body) return;
  body.innerHTML = `<tr><td colspan="10" class="muted">Loading...</td></tr>`;
  try {
    const q = new URLSearchParams();
    const start = String(document.getElementById("histFilterStart")?.value || "").trim();
    const end = String(document.getElementById("histFilterEnd")?.value || "").trim();
    const location = String(document.getElementById("histFilterLocation")?.value || "").trim();
    const part = String(document.getElementById("histFilterPart")?.value || "").trim();
    const approval = String(document.getElementById("histFilterApproval")?.value || "").trim();
    if (start) q.set("start", start);
    if (end) q.set("end", end);
    if (location) q.set("location", location);
    if (part) q.set("part", part);
    if (approval) q.set("approval", approval);
    const res = await fetch(`${API}/maintenance/histogram/events?${q.toString()}`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load histogram events");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    if (!rows.length) {
      body.innerHTML = `<tr><td colspan="10" class="muted">No events found for selected filters.</td></tr>`;
      if (msg) {
        msg.className = "muted";
        msg.textContent = "No events found.";
      }
      return;
    }
    body.innerHTML = rows.map((r) => `
      <tr>
        <td>${esc(r.event_date || "-")}</td>
        <td>${esc(r.asset_number || "-")}</td>
        <td>${esc(r.location || "-")}</td>
        <td>${esc(r.part_code || "-")}</td>
        <td>${esc(r.part_name || "-")}</td>
        <td>${esc(r.approval_status || "-")}</td>
        <td>${esc(r.approved_by || "-")}</td>
        <td>${esc(r.notes || "-")}</td>
        <td>${esc(r.created_by || "-")}</td>
        <td style="white-space:nowrap;">
          <button type="button" data-hist-edit="${Number(r.id || 0)}">Edit</button>
          <button type="button" data-hist-del="${Number(r.id || 0)}">Delete</button>
        </td>
      </tr>
    `).join("");
    if (msg) {
      msg.className = "muted";
      msg.textContent = `Loaded ${rows.length} event(s).`;
    }
  } catch (e) {
    body.innerHTML = `<tr><td colspan="10" class="message-error">${esc(e.message || String(e))}</td></tr>`;
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Load error: ${e.message || e}`;
    }
  }
}

async function saveHistogramEvent() {
  const msg = document.getElementById("histMsg");
  const event_date = String(document.getElementById("histEventDate")?.value || "").trim();
  const asset_number = String(document.getElementById("histAssetNumber")?.value || "").trim();
  const location = String(document.getElementById("histLocation")?.value || "").trim();
  const part_code = String(document.getElementById("histPartCode")?.value || "").trim();
  const part_name = String(document.getElementById("histPartName")?.value || "").trim();
  const approval_status = String(document.getElementById("histApprovalStatus")?.value || "").trim();
  const approved_by = String(document.getElementById("histApprovedBy")?.value || "").trim();
  const notes = String(document.getElementById("histNotes")?.value || "").trim();
  if (!event_date) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Event date is required.";
    }
    return;
  }
  try {
    const res = await fetch(`${API}/maintenance/histogram/events`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({ event_date, asset_number, location, part_code, part_name, approval_status, approved_by, notes }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save event");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = "Histogram event saved.";
    }
    const clearIds = ["histAssetNumber", "histLocation", "histPartCode", "histPartName", "histApprovedBy", "histNotes"];
    clearIds.forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.value = "";
    });
    const approvalEl = document.getElementById("histApprovalStatus");
    if (approvalEl) approvalEl.value = "";
    await loadHistogramEvents();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Save error: ${e.message || e}`;
    }
  }
}

function openHistogramPdf(download = false) {
  const q = new URLSearchParams();
  q.set("include_all", "1");
  q.set("site_code", getSessionSite());
  if (download) q.set("download", "1");
  window.open(`${API}/maintenance/histogram/events.pdf?${q.toString()}`, "_blank");
}

async function editHistogramEvent(id) {
  const n = Number(id || 0);
  if (!n) return;
  const rows = Array.from(document.querySelectorAll("#histEventBody tr"));
  const row = rows.find((tr) => Number(tr.querySelector("button[data-hist-edit]")?.getAttribute("data-hist-edit") || 0) === n);
  const tds = row ? row.querySelectorAll("td") : [];
  const currentDate = String(tds[0]?.textContent || "").trim();
  const currentAssetNumber = String(tds[1]?.textContent || "").trim();
  const currentLocation = String(tds[2]?.textContent || "").trim();
  const currentPartCode = String(tds[3]?.textContent || "").trim();
  const currentPartName = String(tds[4]?.textContent || "").trim();
  const currentApproval = String(tds[5]?.textContent || "").trim();
  const currentApprovedBy = String(tds[6]?.textContent || "").trim();
  const currentNotes = String(tds[7]?.textContent || "").trim();

  const event_date = String(window.prompt("Event date (YYYY-MM-DD):", currentDate) || "").trim();
  if (!event_date) return;
  const asset_number = String(window.prompt("Asset number:", currentAssetNumber) || "").trim();
  const location = String(window.prompt("Location:", currentLocation) || "").trim();
  const part_code = String(window.prompt("Part code:", currentPartCode) || "").trim();
  const part_name = String(window.prompt("Part name:", currentPartName) || "").trim();
  const approval_status = String(window.prompt("Approval status (Pending/Approved/Rejected):", currentApproval) || "").trim();
  const approved_by = String(window.prompt("Approved by:", currentApprovedBy) || "").trim();
  const notes = String(window.prompt("Notes:", currentNotes) || "").trim();

  const msg = document.getElementById("histMsg");
  try {
    const res = await fetch(`${API}/maintenance/histogram/events/${n}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({ event_date, asset_number, location, part_code, part_name, approval_status, approved_by, notes }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to update event");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = "Histogram event updated.";
    }
    await loadHistogramEvents();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Update error: ${e.message || e}`;
    }
  }
}

async function deleteHistogramEvent(id) {
  const n = Number(id || 0);
  if (!n) return;
  if (!window.confirm("Delete this histogram event?")) return;
  const msg = document.getElementById("histMsg");
  try {
    const res = await fetch(`${API}/maintenance/histogram/events/${n}`, {
      method: "DELETE",
      headers: authHeaders(),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to delete event");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = "Histogram event deleted.";
    }
    await loadHistogramEvents();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Delete error: ${e.message || e}`;
    }
  }
}

function getSyncForm() {
  const peer = String(document.getElementById("syncPeer")?.value || "").trim();
  const since = Number(document.getElementById("syncSinceId")?.value || 0);
  const limit = Number(document.getElementById("syncLimit")?.value || 200);
  return {
    peer: peer || "local-maint-ui",
    since_id: Number.isFinite(since) && since >= 0 ? Math.trunc(since) : 0,
    limit: Number.isFinite(limit) ? Math.max(1, Math.min(5000, Math.trunc(limit))) : 200,
  };
}

async function syncLoadStats() {
  syncSetMsg("Loading sync stats...");
  try {
    const res = await fetch(`${API}/sync/stats`, { headers: syncHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Stats load failed");
    syncSetOutput(data);
    syncSetMsg(`Stats loaded. Unsynced: ${Number(data.unsynced || 0)}.`);
  } catch (e) {
    syncSetMsg(`Sync stats error: ${e.message || e}`, true);
  }
}

async function syncLoadState() {
  const f = getSyncForm();
  syncSetMsg("Loading sync state...");
  try {
    const q = new URLSearchParams();
    q.set("schema_version", "1");
    if (f.peer) q.set("peer", f.peer);
    const res = await fetch(`${API}/sync/state?${q.toString()}`, { headers: syncHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Sync state load failed");
    syncSetOutput(data);
    syncSetMsg(`State loaded${data.peer ? ` for ${data.peer}` : ""}.`);
  } catch (e) {
    syncSetMsg(`Sync state error: ${e.message || e}`, true);
  }
}

async function syncLoadOutbox() {
  const { limit } = getSyncForm();
  syncSetMsg("Loading outbox...");
  try {
    const res = await fetch(`${API}/sync/outbox?limit=${limit}`, { headers: syncHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Outbox load failed");
    syncSetOutput(data);
    syncSetMsg(`Outbox loaded (${Number(data.count || 0)} rows).`);
  } catch (e) {
    syncSetMsg(`Sync outbox error: ${e.message || e}`, true);
  }
}

async function syncPullEvents() {
  const f = getSyncForm();
  syncSetMsg("Pulling sync events...");
  try {
    const q = new URLSearchParams();
    q.set("peer", f.peer);
    q.set("since_id", String(f.since_id));
    q.set("limit", String(f.limit));
    const res = await fetch(`${API}/sync/pull?${q.toString()}`, { headers: syncHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Sync pull failed");
    lastSyncPull = { last_id: Number(data.last_id || 0), events: Array.isArray(data.events) ? data.events : [] };
    const sinceEl = document.getElementById("syncSinceId");
    if (sinceEl) sinceEl.value = String(lastSyncPull.last_id || f.since_id);
    syncSetOutput(data);
    syncSetMsg(`Pulled ${lastSyncPull.events.length} event(s). Last id: ${lastSyncPull.last_id}.`);
  } catch (e) {
    syncSetMsg(`Sync pull error: ${e.message || e}`, true);
  }
}

async function syncApplyLastPull() {
  const f = getSyncForm();
  if (!lastSyncPull.events.length) {
    syncSetMsg("No pulled events to apply yet. Click Pull first.", true);
    return;
  }
  syncSetMsg(`Applying ${lastSyncPull.events.length} event(s)...`);
  try {
    const res = await fetch(`${API}/sync/apply`, {
      method: "POST",
      headers: syncHeaders(),
      body: JSON.stringify({ peer: f.peer, events: lastSyncPull.events }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Sync apply failed");
    syncSetOutput(data);
    syncSetMsg(`Apply complete. Applied: ${Number(data.applied || 0)}, skipped: ${Number(data.skipped || 0)}.`);
  } catch (e) {
    syncSetMsg(`Sync apply error: ${e.message || e}`, true);
  }
}

async function syncApplyLastPullDryRun() {
  const f = getSyncForm();
  if (!lastSyncPull.events.length) {
    syncSetMsg("No pulled events to dry-run. Click Pull first.", true);
    return;
  }
  syncSetMsg(`Dry-run applying ${lastSyncPull.events.length} event(s)...`);
  try {
    const res = await fetch(`${API}/sync/apply`, {
      method: "POST",
      headers: syncHeaders(),
      body: JSON.stringify({ schema_version: 1, dry_run: 1, peer: f.peer, events: lastSyncPull.events }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Sync apply dry-run failed");
    syncSetOutput(data);
    syncSetMsg(`Dry-run done. Would apply: ${Number(data.would_apply || 0)}, skipped: ${Number(data.skipped || 0)}.`);
  } catch (e) {
    syncSetMsg(`Sync apply dry-run error: ${e.message || e}`, true);
  }
}

function syncExportLastPull() {
  const f = getSyncForm();
  if (!lastSyncPull.events.length) {
    syncSetMsg("No pulled events to export yet. Click Pull first.", true);
    return;
  }
  try {
    const payload = {
      schema_version: 1,
      exported_at: new Date().toISOString(),
      peer: f.peer,
      last_id: Number(lastSyncPull.last_id || 0),
      count: Number(lastSyncPull.events.length || 0),
      events: lastSyncPull.events,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const a = document.createElement("a");
    a.href = url;
    a.download = `ironlog_sync_last_pull_${f.peer || "peer"}_${stamp}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    syncSetMsg(`Exported ${lastSyncPull.events.length} event(s) to JSON.`);
  } catch (e) {
    syncSetMsg(`Export failed: ${e.message || e}`, true);
  }
}

async function syncAckLastPull() {
  if (!lastSyncPull.events.length) {
    syncSetMsg("No pulled events to acknowledge yet. Click Pull first.", true);
    return;
  }
  const ids = lastSyncPull.events
    .map((e) => Number(e?.id))
    .filter((n) => Number.isInteger(n) && n > 0);
  if (!ids.length) {
    syncSetMsg("Last pull has no event IDs to acknowledge.", true);
    return;
  }
  syncSetMsg(`Acknowledging ${ids.length} outbox event(s)...`);
  try {
    const res = await fetch(`${API}/sync/outbox/ack`, {
      method: "POST",
      headers: syncHeaders(),
      body: JSON.stringify({ ids }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Outbox ack failed");
    syncSetOutput(data);
    syncSetMsg(`Acknowledged ${Number(data.acknowledged || 0)} outbox event(s).`);
  } catch (e) {
    syncSetMsg(`Outbox ack error: ${e.message || e}`, true);
  }
}

async function syncCheckpointLastPull() {
  const f = getSyncForm();
  const lastId = Number(lastSyncPull.last_id || f.since_id || 0);
  syncSetMsg(`Saving checkpoint at ${lastId}...`);
  try {
    const res = await fetch(`${API}/sync/checkpoint`, {
      method: "POST",
      headers: syncHeaders(),
      body: JSON.stringify({ peer: f.peer, last_outbox_id: lastId }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Checkpoint save failed");
    syncSetOutput(data);
    syncSetMsg(`Checkpoint saved for ${f.peer} at ${lastId}.`);
  } catch (e) {
    syncSetMsg(`Checkpoint error: ${e.message || e}`, true);
  }
}

async function loadAssetsForInspection() {
  const selA = document.getElementById("miAsset");
  const selF = document.getElementById("miFilterAsset");
  const aiA = document.getElementById("aiAsset");
  const aiF = document.getElementById("aiFilterAsset");
  const drA = document.getElementById("drAsset");
  const drF = document.getElementById("drFilterAsset");
  if (!selA || !selF) return;
  try {
    const res = await fetch(`${API}/assets?include_archived=0`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load assets");
    const rows = Array.isArray(data) ? data : [];
    const opts = rows
      .filter((a) => Number(a.active ?? 1) !== 0 && Number(a.archived ?? 0) === 0)
      .map((a) => `<option value="${Number(a.id)}">${esc(a.asset_code)} - ${esc(a.asset_name)}</option>`)
      .join("");
    selA.innerHTML = `<option value="">Select asset</option>${opts}`;
    selF.innerHTML = `<option value="">All assets</option>${opts}`;
    if (aiA) aiA.innerHTML = `<option value="">Select asset</option>${opts}`;
    if (aiF) aiF.innerHTML = `<option value="">All assets</option>${opts}`;
    if (drA) drA.innerHTML = `<option value="">Select asset</option>${opts}`;
    if (drF) drF.innerHTML = `<option value="">All assets</option>${opts}`;
  } catch (e) {
    selA.innerHTML = `<option value="">Assets load failed</option>`;
    selF.innerHTML = `<option value="">Assets load failed</option>`;
    if (aiA) aiA.innerHTML = `<option value="">Assets load failed</option>`;
    if (aiF) aiF.innerHTML = `<option value="">Assets load failed</option>`;
    if (drA) drA.innerHTML = `<option value="">Assets load failed</option>`;
    if (drF) drF.innerHTML = `<option value="">Assets load failed</option>`;
  }
}

async function saveManagerInspection() {
  const asset_id = Number(document.getElementById("miAsset")?.value || 0);
  const inspection_date = String(document.getElementById("miDate")?.value || "").trim();
  const inspector_name = String(document.getElementById("miInspector")?.value || "").trim();
  const notes = String(document.getElementById("miNotes")?.value || "").trim();
  const msg = document.getElementById("miMsg");
  const mhRaw = String(document.getElementById("miMachineHours")?.value || "").trim();
  const machine_hours = mhRaw === "" ? null : Number(mhRaw);
  if (machine_hours != null && !Number.isFinite(machine_hours)) {
    return alert("Machine hours must be a number.");
  }
  const checklist = collectManagerInspectionChecklist();
  const required_parts = collectManagerInspectionParts();
  const create_work_order = document.getElementById("miCreateWoAlways")?.checked === true;
  const create_work_order_on_issues = document.getElementById("miAutoWoOnIssues")?.checked !== false;

  if (!asset_id) return alert("Select an asset.");
  if (!inspection_date) return alert("Select inspection date.");
  if (msg) msg.textContent = "Saving inspection...";
  try {
    const res = await fetch(`${API}/maintenance/inspections`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({
        asset_id,
        inspection_date,
        inspector_name,
        notes,
        machine_hours,
        checklist,
        required_parts,
        create_work_order,
        create_work_order_on_issues,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save inspection");
    const wo = data.work_order_id ? ` Work order #${data.work_order_id} created.` : "";
    if (msg) msg.textContent = `Inspection saved.${wo}`;
    resetManagerInspectionForm();
    await loadManagerInspections();
  } catch (e) {
    if (msg) msg.textContent = `Save error: ${e.message || e}`;
  }
}

async function uploadInspectionPhoto(inspectionId) {
  const fileEl = document.getElementById(`miPhotoFile-${inspectionId}`);
  const capEl = document.getElementById(`miPhotoCaption-${inspectionId}`);
  const file = fileEl?.files?.[0];
  if (!file) return alert("Choose a photo first.");
  const fd = new FormData();
  fd.append("file", file);
  const caption = String(capEl?.value || "").trim();
  const q = caption ? `?caption=${encodeURIComponent(caption)}` : "";
  const res = await fetch(`${API}/maintenance/inspections/${inspectionId}/photo${q}`, {
    method: "POST",
    body: fd,
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Photo upload failed");
  await loadManagerInspections();
}

async function saveDamageReport() {
  const asset_id = Number(document.getElementById("drAsset")?.value || 0);
  const report_date = String(document.getElementById("drDate")?.value || "").trim();
  const damage_time = String(document.getElementById("drTime")?.value || "").trim();
  const inspector_name = String(document.getElementById("drInspector")?.value || "").trim();
  const hour_meter_raw = String(document.getElementById("drHours")?.value || "").trim();
  const damage_location = String(document.getElementById("drLocation")?.value || "").trim();
  const responsible_person = String(document.getElementById("drResponsiblePerson")?.value || "").trim();
  const severity = String(document.getElementById("drSeverity")?.value || "").trim();
  const damage_description = String(document.getElementById("drDescription")?.value || "").trim();
  const immediate_action = String(document.getElementById("drAction")?.value || "").trim();
  const out_of_service = document.getElementById("drOutOfService")?.checked ? 1 : 0;
  const pending_investigation = document.getElementById("drPendingInvestigation")?.checked ? 1 : 0;
  const hse_report_available = document.getElementById("drHseReportAvailable")?.checked ? 1 : 0;
  const msg = document.getElementById("drMsg");

  if (!asset_id) return alert("Select an asset for damage report.");
  if (!report_date) return alert("Select damage report date.");
  if (!damage_location) return alert("Enter damage location.");
  if (!severity) return alert("Select severity.");
  if (!damage_description) return alert("Enter damage description.");
  if (!immediate_action) return alert("Enter immediate action.");

  const hour_meter = hour_meter_raw === "" ? null : Number(hour_meter_raw);
  if (hour_meter != null && !Number.isFinite(hour_meter)) {
    return alert("Machine hours must be numeric.");
  }

  if (msg) msg.textContent = "Saving damage report...";
  try {
    const res = await fetch(`${API}/maintenance/damage-reports`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        asset_id,
        report_date,
        damage_time,
        inspector_name,
        hour_meter,
        damage_location,
        responsible_person,
        severity,
        damage_description,
        immediate_action,
        out_of_service,
        pending_investigation,
        hse_report_available,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save damage report");
    if (msg) msg.textContent = "Damage report saved.";
    document.getElementById("drDescription").value = "";
    document.getElementById("drAction").value = "";
    document.getElementById("drLocation").value = "";
    document.getElementById("drResponsiblePerson").value = "";
    document.getElementById("drHours").value = "";
    document.getElementById("drTime").value = "";
    document.getElementById("drSeverity").value = "";
    document.getElementById("drOutOfService").checked = false;
    document.getElementById("drPendingInvestigation").checked = false;
    document.getElementById("drHseReportAvailable").checked = false;
    await loadDamageReports();
  } catch (e) {
    if (msg) msg.textContent = `Save error: ${e.message || e}`;
  }
}

async function uploadDamagePhoto(reportId) {
  const fileEl = document.getElementById(`drPhotoFile-${reportId}`);
  const capEl = document.getElementById(`drPhotoCaption-${reportId}`);
  const file = fileEl?.files?.[0];
  if (!file) return alert("Choose a photo first.");
  const fd = new FormData();
  fd.append("file", file);
  const caption = String(capEl?.value || "").trim();
  const q = caption ? `?caption=${encodeURIComponent(caption)}` : "";
  const res = await fetch(`${API}/maintenance/damage-reports/${reportId}/photo${q}`, {
    method: "POST",
    body: fd,
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Damage photo upload failed");
  await loadDamageReports();
}

function damageCard(r) {
  const photos = Array.isArray(r.photos) ? r.photos : [];
  const sev = String(r.severity || "").toLowerCase();
  const sevClass = sev === "critical" || sev === "high" ? "status-overdue" : sev === "medium" ? "status-soon" : "status-ok";
  const photoHtml = photos.length
    ? `<div class="row stack-10">${photos.map((p) => {
        const src = normImgSrc(p.file_path);
        return `<div style="display:flex; flex-direction:column; gap:4px;">
          <img src="${esc(src)}" alt="damage photo" style="width:140px; height:100px; object-fit:cover; border:1px solid #d1d5db; border-radius:8px;" />
          <small class="muted">${esc(p.caption || "")}</small>
        </div>`;
      }).join("")}</div>`
    : `<small class="muted">No photos yet.</small>`;

  return `
    <div class="card">
      <div><b>${esc(r.asset_code)}</b> - ${esc(r.asset_name || "")}</div>
      <div><small>Date: ${esc(r.report_date || "-")} ${esc(r.damage_time || "")} | Inspector: ${esc(r.inspector_name || "-")} | Hours: ${esc(r.hour_meter == null ? "-" : String(Number(r.hour_meter).toFixed(1)))}</small></div>
      <div><small>Location: <b>${esc(r.damage_location || "-")}</b> | Responsible: <b>${esc(r.responsible_person || "-")}</b></small></div>
      <div><small>Severity: <span class="${sevClass}">${esc((r.severity || "-").toUpperCase())}</span> | Out of service: <b>${Number(r.out_of_service || 0) ? "YES" : "NO"}</b> | Pending investigation: <b>${Number(r.pending_investigation || 0) ? "YES" : "NO"}</b> | HSE report: <b>${Number(r.hse_report_available || 0) ? "YES" : "NO"}</b></small></div>
      <div style="margin-top:6px;"><small><b>Damage:</b> ${esc(r.damage_description || "")}</small></div>
      <div style="margin-top:4px;"><small><b>Immediate action:</b> ${esc(r.immediate_action || "")}</small></div>
      <div style="margin-top:8px;">${photoHtml}</div>
      <div class="row stack-10" style="margin-top:8px;">
        <button data-dr-open-pdf="${Number(r.id)}">Open PDF</button>
        <button data-dr-download-pdf="${Number(r.id)}">Download PDF</button>
        <input id="drPhotoFile-${Number(r.id)}" type="file" accept="image/*" />
        <input id="drPhotoCaption-${Number(r.id)}" class="w-200" placeholder="Photo caption (optional)" />
        <button data-dr-upload="${Number(r.id)}">Upload Photo</button>
      </div>
    </div>
  `;
}

async function loadDamageReports() {
  const list = document.getElementById("drList");
  if (!list) return;
  list.innerHTML = `<div class="skeleton-block"></div>`;
  const asset_id = String(document.getElementById("drFilterAsset")?.value || "").trim();
  const start = String(document.getElementById("drStart")?.value || "").trim();
  const end = String(document.getElementById("drEnd")?.value || "").trim();
  const responsible_person = String(document.getElementById("drFilterResponsiblePerson")?.value || "").trim();
  const pending_investigation = String(document.getElementById("drFilterPendingInvestigation")?.value || "").trim();
  const hse_report_available = String(document.getElementById("drFilterHseReportAvailable")?.value || "").trim();
  const q = new URLSearchParams();
  if (asset_id) q.set("asset_id", asset_id);
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  if (responsible_person) q.set("responsible_person", responsible_person);
  if (pending_investigation === "0" || pending_investigation === "1") q.set("pending_investigation", pending_investigation);
  if (hse_report_available === "0" || hse_report_available === "1") q.set("hse_report_available", hse_report_available);
  try {
    const res = await fetch(`${API}/maintenance/damage-reports${q.toString() ? `?${q.toString()}` : ""}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load damage reports");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    list.innerHTML = rows.length ? rows.map(damageCard).join("") : `<div class="muted">No damage reports found.</div>`;
  } catch (e) {
    list.innerHTML = `<div class="message-error">Damage report load error: ${esc(e.message || e)}</div>`;
  }
}

function openDamageReportPdf(id, download = false) {
  const n = Number(id || 0);
  if (!n) return;
  const q = download ? "?download=1" : "";
  window.open(`${API}/reports/damage-report/${n}.pdf${q}`, "_blank");
}

function openDamageReportsBulkPdf(download = false, withPhotos = false) {
  const start = String(document.getElementById("drStart")?.value || "").trim();
  const end = String(document.getElementById("drEnd")?.value || "").trim();
  const assetId = String(document.getElementById("drFilterAsset")?.value || "").trim();
  if (!start || !end) {
    alert("Select damage report start and end dates first.");
    return;
  }
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  if (assetId) q.set("asset_id", assetId);
  if (withPhotos) q.set("with_photos", "1");
  if (download) q.set("download", "1");
  window.open(`${API}/reports/damage-reports.pdf?${q.toString()}`, "_blank");
}

function openManagerInspectionPdf(id, download = false) {
  const n = Number(id || 0);
  if (!n) return;
  const q = download ? "?download=1" : "";
  window.open(`${API}/reports/manager-inspection/${n}.pdf${q}`, "_blank");
}

function openManagerInspectionsBulkPdf(download = false, withPhotos = false) {
  const start = String(document.getElementById("miStart")?.value || "").trim();
  const end = String(document.getElementById("miEnd")?.value || "").trim();
  const assetId = String(document.getElementById("miFilterAsset")?.value || "").trim();
  if (!start || !end) {
    alert("Select start and end dates first.");
    return;
  }
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  if (assetId) q.set("asset_id", assetId);
  if (withPhotos) q.set("with_photos", "1");
  if (download) q.set("download", "1");
  window.open(`${API}/reports/manager-inspections.pdf?${q.toString()}`, "_blank");
}

function inspectionCard(r) {
  const photos = Array.isArray(r.photos) ? r.photos : [];
  const photoHtml = photos.length
    ? `<div class="row stack-10">${photos.map((p) => {
        const src = normImgSrc(p.file_path);
        return `<div style="display:flex; flex-direction:column; gap:4px;">
          <img src="${esc(src)}" alt="inspection photo" style="width:140px; height:100px; object-fit:cover; border:1px solid #d1d5db; border-radius:8px;" />
          <small class="muted">${esc(p.caption || "")}</small>
        </div>`;
      }).join("")}</div>`
    : `<small class="muted">No photos yet.</small>`;

  const hrs =
    r.machine_hours != null && Number.isFinite(Number(r.machine_hours))
      ? Number(r.machine_hours).toFixed(1)
      : "—";
  const live =
    r.live_hours_snapshot != null && Number.isFinite(Number(r.live_hours_snapshot))
      ? `${Number(r.live_hours_snapshot).toFixed(1)} (${esc(r.live_hours_source || "—")})`
      : "—";
  const wo =
    r.work_order_id != null && Number(r.work_order_id) > 0
      ? `<b>WO #${Number(r.work_order_id)}</b>`
      : `<span class="muted">No WO</span>`;
  const chk = Array.isArray(r.checklist) ? r.checklist : [];
  const fails = chk.filter((c) => c.ok === false);
  const failLine = fails.length
    ? `<div style="margin-top:4px;"><small class="status-overdue">Checklist fail: ${fails.map((c) => esc(c.label || c.key)).join("; ")}</small></div>`
    : "";
  const parts = Array.isArray(r.required_parts) ? r.required_parts : [];
  const partsLine = parts.length
    ? `<div style="margin-top:4px;"><small><b>Parts:</b> ${parts.map((p) => `${esc(p.part_code)} × ${esc(String(p.qty))}`).join(", ")}</small></div>`
    : "";

  return `
    <div class="card">
      <div><b>${esc(r.asset_code)}</b> - ${esc(r.asset_name || "")}</div>
      <div><small>Date: ${esc(r.inspection_date)} | Inspector: ${esc(r.inspector_name || "-")}</small></div>
      <div><small>Machine hrs: ${esc(hrs)} | Live snapshot: ${live} | ${wo}</small></div>
      ${failLine}
      ${partsLine}
      <div style="margin-top:6px;"><small>${esc(r.notes || "")}</small></div>
      <div style="margin-top:8px;">${photoHtml}</div>
      <div class="row stack-10" style="margin-top:8px;">
        <button data-mi-open-pdf="${Number(r.id)}">Open PDF</button>
        <button data-mi-download-pdf="${Number(r.id)}">Download PDF</button>
        <input id="miPhotoFile-${Number(r.id)}" type="file" accept="image/*" />
        <input id="miPhotoCaption-${Number(r.id)}" class="w-200" placeholder="Photo caption (optional)" />
        <button data-mi-upload="${Number(r.id)}">Upload Photo</button>
      </div>
    </div>
  `;
}

async function loadManagerInspections() {
  const list = document.getElementById("miList");
  if (!list) return;
  list.innerHTML = `<div class="skeleton-block"></div>`;
  const asset_id = String(document.getElementById("miFilterAsset")?.value || "").trim();
  const start = String(document.getElementById("miStart")?.value || "").trim();
  const end = String(document.getElementById("miEnd")?.value || "").trim();
  const q = new URLSearchParams();
  if (asset_id) q.set("asset_id", asset_id);
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  try {
    const res = await fetch(`${API}/maintenance/inspections${q.toString() ? `?${q.toString()}` : ""}`, {
      headers: authHeaders(),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load inspections");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    list.innerHTML = rows.length ? rows.map(inspectionCard).join("") : `<div class="muted">No inspections found.</div>`;
  } catch (e) {
    list.innerHTML = `<div class="message-error">Inspection load error: ${esc(e.message || e)}</div>`;
  }
}

function fmtMoney(v) {
  const n = Number(v || 0);
  return Number.isFinite(n) ? n.toFixed(2) : "0.00";
}

function fmtPct(v) {
  if (v == null || v === "" || !Number.isFinite(Number(v))) return "—";
  return `${Number(v).toFixed(1)}%`;
}

function weeklyForumQueryString() {
  const start = String(document.getElementById("wfStart")?.value || "").trim();
  const end = String(document.getElementById("wfEnd")?.value || "").trim();
  const near = Math.max(1, Number(document.getElementById("wfNearDueHours")?.value || 50));
  const q = new URLSearchParams();
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  q.set("near_due_hours", String(near));
  return q.toString();
}

let wfUpcomingCache = [];
let wfInputsCache = [];
let wfPartsCache = [];
let wfDraftItems = [];
function wfPlanLabel(r) {
  return `${String(r.asset_code || "-")} - ${String(r.asset_name || "-")} | ${String(r.service_name || "-")} (Plan ${Number(r.plan_id || 0)})`;
}
function refreshWeeklyForumPlanOptions() {
  const sel = document.getElementById("wfInputPlan");
  if (!sel) return;
  sel.innerHTML = wfUpcomingCache.length
    ? `<option value="">Select upcoming service</option>${wfUpcomingCache.map((r) => `<option value="${Number(r.plan_id || 0)}">${esc(wfPlanLabel(r))}</option>`).join("")}`
    : `<option value="">No upcoming services loaded</option>`;
}
function refreshWeeklyForumInputsTable() {
  const body = document.getElementById("wfInputsBody");
  if (!body) return;
  body.innerHTML = wfInputsCache.length
    ? wfInputsCache.map((r) => {
        const plan = wfUpcomingCache.find((p) => Number(p.plan_id || 0) === Number(r.plan_id || 0));
        let items = [];
        try {
          const parsed = JSON.parse(String(r.items_json || "[]"));
          if (Array.isArray(parsed)) items = parsed;
        } catch {}
        if (!items.length) {
          const oilCode = String(r.oil_part_code || "").trim();
          const oilQty = Number(r.oil_qty || 0);
          const partCode = String(r.parts_part_code || "").trim();
          const partQty = Number(r.parts_qty || 0);
          if (oilCode && oilQty > 0) items.push({ type: "oil", part_code: oilCode, qty: oilQty });
          if (partCode && partQty > 0) items.push({ type: "part", part_code: partCode, qty: partQty });
        }
        const oils = items.filter((x) => String(x.type || "part").toLowerCase() === "oil");
        const parts = items.filter((x) => String(x.type || "part").toLowerCase() !== "oil");
        const render = (rows) => rows.length
          ? rows.map((x) => `${String(x.part_code || "")} (${fmt1(x.qty)})`).join(", ")
          : "-";
        return `
          <tr>
            <td>${esc(plan ? wfPlanLabel(plan) : `Plan ${Number(r.plan_id || 0)}`)}</td>
            <td>${esc(render(oils))}</td>
            <td>${esc(render(parts))}</td>
            <td>${esc(r.notes || "")}</td>
          </tr>
        `;
      }).join("")
    : `<tr><td colspan="4" class="muted">No manual inputs saved.</td></tr>`;
}
function refreshWeeklyForumPartsDatalist() {
  const html = wfPartsCache.map((p) => {
    const code = String(p.part_code || "").trim();
    const desc = `${code} - ${String(p.part_name || "")} | on hand ${Number(p.on_hand || 0).toFixed(2)} | unit ${Number(p.latest_unit_cost || 0).toFixed(2)}`;
    return `<option value="${esc(code)}">${esc(desc)}</option>`;
  }).join("");
  const dl = document.getElementById("wfPartsList");
  if (dl) dl.innerHTML = html;
  const dlMi = document.getElementById("miPartsList");
  if (dlMi) dlMi.innerHTML = html;
}
function getWfPartByCode(codeIn) {
  const code = String(codeIn || "").trim().toUpperCase();
  if (!code) return null;
  return wfPartsCache.find((p) => String(p.part_code || "").trim().toUpperCase() === code) || null;
}
function getMiPartByCode(codeIn) {
  return getWfPartByCode(codeIn);
}

const MANAGER_INSPECTION_CHECKLIST = [
  { key: "structure", label: "Structure / visible damage" },
  { key: "fluids", label: "Fluids / leaks" },
  { key: "tyres", label: "Tyres / undercarriage" },
  { key: "safety", label: "Safety & access (rails, steps, extinguisher)" },
  { key: "cabin", label: "Cabin / visibility / instruments" },
  { key: "attachments", label: "GET / tools / attachments" },
  { key: "housekeeping", label: "Housekeeping" },
  { key: "noise", label: "Operation / unusual noise" },
];

function renderManagerInspectionChecklist() {
  const host = document.getElementById("miChecklist");
  if (!host) return;
  host.innerHTML = MANAGER_INSPECTION_CHECKLIST.map(
    (row) => `
    <div class="row stack-10" style="align-items:center; flex-wrap:wrap; gap:8px;">
      <span style="min-width:240px; font-size:13px;">${esc(row.label)}</span>
      <span class="row stack-10" style="gap:10px;">
        <label><input type="radio" name="miChk-${esc(row.key)}" value="ok" /> OK</label>
        <label><input type="radio" name="miChk-${esc(row.key)}" value="fail" /> Fail</label>
        <label><input type="radio" name="miChk-${esc(row.key)}" value="na" /> N/A</label>
      </span>
      <input type="text" class="w-200 mi-chk-note" data-mi-chk="${esc(row.key)}" placeholder="Note (optional)" />
    </div>`
  ).join("");
}

function collectManagerInspectionChecklist() {
  return MANAGER_INSPECTION_CHECKLIST.map((row) => {
    const sel = document.querySelector(`input[name="miChk-${row.key}"]:checked`);
    const val = sel ? String(sel.value || "") : "";
    let ok = null;
    if (val === "ok") ok = true;
    else if (val === "fail") ok = false;
    const noteEl = document.querySelector(`input.mi-chk-note[data-mi-chk="${row.key}"]`);
    const note = String(noteEl?.value || "").trim() || null;
    return { key: row.key, label: row.label, ok, note };
  });
}

function addManagerInspectionPartRow(partCode = "", qty = "", note = "") {
  const body = document.getElementById("miPartsBody");
  if (!body) return;
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td><input class="mi-part-code" type="text" list="miPartsList" style="min-width:160px;" value="${esc(partCode)}" /></td>
    <td><input class="mi-part-qty" type="number" step="0.01" min="0" style="width:90px;" value="${esc(qty)}" /></td>
    <td><input class="mi-part-note" type="text" style="min-width:140px;" value="${esc(note)}" /></td>
    <td><button type="button" class="mi-part-remove">Remove</button></td>`;
  body.appendChild(tr);
  tr.querySelector(".mi-part-remove")?.addEventListener("click", () => {
    tr.remove();
  });
}

function collectManagerInspectionParts() {
  const body = document.getElementById("miPartsBody");
  if (!body) return [];
  const out = [];
  body.querySelectorAll("tr").forEach((tr) => {
    const part_code = String(tr.querySelector(".mi-part-code")?.value || "").trim();
    const qty = Number(tr.querySelector(".mi-part-qty")?.value || 0);
    const note = String(tr.querySelector(".mi-part-note")?.value || "").trim() || null;
    if (!part_code || !Number.isFinite(qty) || qty <= 0) return;
    const meta = getMiPartByCode(part_code);
    out.push({
      part_id: meta?.id != null ? Number(meta.id) : null,
      part_code,
      qty,
      note,
    });
  });
  return out;
}

async function pullManagerInspectionLiveHours() {
  const assetId = Number(document.getElementById("miAsset")?.value || 0);
  const inspectionDate = String(document.getElementById("miDate")?.value || "").trim();
  const meta = document.getElementById("miLiveMeta");
  const inp = document.getElementById("miMachineHours");
  if (!assetId) {
    if (meta) meta.textContent = "Select an asset first.";
    return;
  }
  if (meta) meta.textContent = "Loading live hours…";
  try {
    const q = new URLSearchParams();
    if (/^\d{4}-\d{2}-\d{2}$/.test(inspectionDate)) q.set("as_of", inspectionDate);
    const qs = q.toString();
    const url = `${API}/maintenance/asset/${assetId}/live-hours${qs ? `?${qs}` : ""}`;
    const res = await fetch(url, { headers: authHeaders() });
    let data;
    try {
      data = await res.json();
    } catch {
      throw new Error(`Live hours response was not JSON (HTTP ${res.status}).`);
    }
    if (!res.ok) throw new Error(data?.error || "Failed to load live hours");
    const h = Number(data.current_hours ?? 0);
    const src = String(data.current_hours_source || "");
    const asOf = data.as_of ? ` up to ${data.as_of}` : "";
    if (inp) inp.value = Number.isFinite(h) ? String(Number(h).toFixed(1)) : "";
    const srcLabel =
      {
        daily_closing: "Daily closing",
        daily_sum: "Daily sum",
        asset_hours: "Asset hours",
      }[src] || src || "—";
    let line = `Live${asOf}: ${Number.isFinite(h) ? h.toFixed(1) : "—"} h (${srcLabel})`;
    if (Number.isFinite(h) && h <= 0) {
      line +=
        " — No meter or usage found for this asset/date; add Daily Input or enter hours manually.";
    }
    if (meta) meta.textContent = line;
  } catch (e) {
    console.error("pullManagerInspectionLiveHours:", e);
    if (meta) meta.textContent = e.message || String(e);
  }
}

function resetManagerInspectionForm() {
  document.getElementById("miNotes").value = "";
  MANAGER_INSPECTION_CHECKLIST.forEach((row) => {
    document.querySelectorAll(`input[name="miChk-${row.key}"]`).forEach((r) => {
      r.checked = false;
    });
    const ne = document.querySelector(`input.mi-chk-note[data-mi-chk="${row.key}"]`);
    if (ne) ne.value = "";
  });
  const body = document.getElementById("miPartsBody");
  if (body) {
    body.innerHTML = "";
    addManagerInspectionPartRow();
  }
  const auto = document.getElementById("miAutoWoOnIssues");
  if (auto) auto.checked = true;
  const al = document.getElementById("miCreateWoAlways");
  if (al) al.checked = false;
  const lm = document.getElementById("miLiveMeta");
  if (lm) lm.textContent = "";
}

const ARTISAN_INSPECTION_CHECKLIST = [
  { key: "prestart", label: "Pre-start visual condition (machine / plant)" },
  { key: "guards", label: "Guards, covers, and safety devices" },
  { key: "hydraulics", label: "Hydraulic hoses, leaks, and fittings" },
  { key: "electrical", label: "Electrical panels / cabling / lights" },
  { key: "lubrication", label: "Lubrication points / levels" },
  { key: "brakes_steering", label: "Brakes / steering / controls response" },
  { key: "alarms", label: "Alarms, horn, and warning systems" },
  { key: "housekeeping", label: "Housekeeping around machine / plant" },
];

function renderArtisanInspectionChecklist() {
  const host = document.getElementById("aiChecklist");
  if (!host) return;
  host.innerHTML = ARTISAN_INSPECTION_CHECKLIST.map(
    (row) => `
    <div class="row stack-10" style="align-items:center; flex-wrap:wrap; gap:8px;">
      <span style="min-width:240px; font-size:13px;">${esc(row.label)}</span>
      <span class="row stack-10" style="gap:10px;">
        <label><input type="radio" name="aiChk-${esc(row.key)}" value="ok" /> OK</label>
        <label><input type="radio" name="aiChk-${esc(row.key)}" value="fail" /> Fail</label>
        <label><input type="radio" name="aiChk-${esc(row.key)}" value="na" /> N/A</label>
      </span>
      <input type="text" class="w-200 ai-chk-note" data-ai-chk="${esc(row.key)}" placeholder="Note (optional)" />
    </div>`
  ).join("");
}

function collectArtisanInspectionChecklist() {
  return ARTISAN_INSPECTION_CHECKLIST.map((row) => {
    const sel = document.querySelector(`input[name="aiChk-${row.key}"]:checked`);
    const val = sel ? String(sel.value || "") : "";
    let ok = null;
    if (val === "ok") ok = true;
    else if (val === "fail") ok = false;
    const noteEl = document.querySelector(`input.ai-chk-note[data-ai-chk="${row.key}"]`);
    const note = String(noteEl?.value || "").trim() || null;
    return { key: row.key, label: row.label, ok, note };
  });
}

async function pullArtisanInspectionLiveHours() {
  const assetId = Number(document.getElementById("aiAsset")?.value || 0);
  const inspectionDate = String(document.getElementById("aiDate")?.value || "").trim();
  const meta = document.getElementById("aiLiveMeta");
  const inp = document.getElementById("aiMachineHours");
  if (!assetId) {
    if (meta) meta.textContent = "Select an asset first.";
    return;
  }
  if (meta) meta.textContent = "Loading live hours…";
  try {
    const q = new URLSearchParams();
    if (/^\d{4}-\d{2}-\d{2}$/.test(inspectionDate)) q.set("as_of", inspectionDate);
    const qs = q.toString();
    const url = `${API}/maintenance/asset/${assetId}/live-hours${qs ? `?${qs}` : ""}`;
    const res = await fetch(url, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data?.error || "Failed to load live hours");
    const h = Number(data.current_hours ?? 0);
    const src = String(data.current_hours_source || "");
    const asOf = data.as_of ? ` up to ${data.as_of}` : "";
    if (inp) inp.value = Number.isFinite(h) ? String(Number(h).toFixed(1)) : "";
    const srcLabel = {
      daily_closing: "Daily closing",
      daily_sum: "Daily sum",
      asset_hours: "Asset hours",
    }[src] || src || "—";
    if (meta) meta.textContent = `Live${asOf}: ${Number.isFinite(h) ? h.toFixed(1) : "—"} h (${srcLabel})`;
  } catch (e) {
    if (meta) meta.textContent = e.message || String(e);
  }
}

function resetArtisanInspectionForm() {
  const notes = document.getElementById("aiNotes");
  if (notes) notes.value = "";
  const shift = document.getElementById("aiShift");
  if (shift) shift.value = "";
  const formNo = document.getElementById("aiFormNumber");
  if (formNo) formNo.value = generateArtisanFormNumber();
  const live = document.getElementById("aiLiveMeta");
  if (live) live.textContent = "";
  ARTISAN_INSPECTION_CHECKLIST.forEach((row) => {
    document.querySelectorAll(`input[name="aiChk-${row.key}"]`).forEach((r) => {
      r.checked = false;
    });
    const ne = document.querySelector(`input.ai-chk-note[data-ai-chk="${row.key}"]`);
    if (ne) ne.value = "";
  });
}

function generateArtisanFormNumber() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const stamp = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}${pad(d.getHours())}${pad(d.getMinutes())}`;
  const rand = Math.floor(Math.random() * 900 + 100);
  return `AI-${stamp}-${rand}`;
}

async function saveArtisanInspection() {
  const asset_id = Number(document.getElementById("aiAsset")?.value || 0);
  const inspection_date = String(document.getElementById("aiDate")?.value || "").trim();
  const inspector_name = String(document.getElementById("aiInspector")?.value || "").trim();
  const shift = String(document.getElementById("aiShift")?.value || "").trim().toLowerCase();
  const form_number = String(document.getElementById("aiFormNumber")?.value || "").trim();
  const notes = String(document.getElementById("aiNotes")?.value || "").trim();
  const msg = document.getElementById("aiMsg");
  const mhRaw = String(document.getElementById("aiMachineHours")?.value || "").trim();
  const machine_hours = mhRaw === "" ? null : Number(mhRaw);
  if (machine_hours != null && !Number.isFinite(machine_hours)) {
    return alert("Machine hours must be a number.");
  }
  const checklist = collectArtisanInspectionChecklist();

  if (!asset_id) return alert("Select an asset.");
  if (!inspection_date) return alert("Select inspection date.");
  if (msg) msg.textContent = "Saving artisan inspection...";
  try {
    const res = await fetch(`${API}/maintenance/artisan-inspections`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({
        asset_id,
        inspection_date,
        inspector_name,
        shift: shift || null,
        form_number: form_number || null,
        notes,
        machine_hours,
        checklist,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save artisan inspection");
    if (msg) msg.textContent = "Artisan inspection saved.";
    resetArtisanInspectionForm();
    await loadArtisanInspections();
  } catch (e) {
    if (msg) msg.textContent = `Save error: ${e.message || e}`;
  }
}

function openArtisanInspectionPdf(id, download = false) {
  const n = Number(id || 0);
  if (!n) return;
  const q = download ? "?download=1" : "";
  window.open(`${API}/reports/artisan-inspection/${n}.pdf${q}`, "_blank");
}

function artisanInspectionCard(r) {
  const hrs = r.machine_hours != null && Number.isFinite(Number(r.machine_hours))
    ? Number(r.machine_hours).toFixed(1)
    : "—";
  const live = r.live_hours_snapshot != null && Number.isFinite(Number(r.live_hours_snapshot))
    ? `${Number(r.live_hours_snapshot).toFixed(1)} (${esc(r.live_hours_source || "—")})`
    : "—";
  const shift = String(r.shift || "").trim();
  const formNo = String(r.form_number || "").trim();
  const chk = Array.isArray(r.checklist) ? r.checklist : [];
  const fails = chk.filter((c) => c.ok === false);
  const failLine = fails.length
    ? `<div style="margin-top:4px;"><small class="status-overdue">Checklist fail: ${fails.map((c) => esc(c.label || c.key)).join("; ")}</small></div>`
    : "";

  return `
    <div class="card">
      <div><b>${esc(r.asset_code)}</b> - ${esc(r.asset_name || "")}</div>
      <div><small>Date: ${esc(r.inspection_date)}${shift ? ` | Shift: ${esc(shift.toUpperCase())}` : ""} | Artisan: ${esc(r.inspector_name || "-")}</small></div>
      <div><small>Form No: <b>${esc(formNo || "-")}</b></small></div>
      <div><small>Machine hrs: ${esc(hrs)} | Live snapshot: ${live}</small></div>
      ${failLine}
      <div style="margin-top:6px;"><small>${esc(r.notes || "")}</small></div>
      <div class="row stack-10" style="margin-top:8px;">
        <button data-ai-open-pdf="${Number(r.id)}">Open PDF</button>
        <button data-ai-download-pdf="${Number(r.id)}">Download PDF</button>
      </div>
    </div>
  `;
}

function openArtisanBlankFormPdf(download = false) {
  const date = String(document.getElementById("aiDate")?.value || "").trim();
  const assetId = String(document.getElementById("aiAsset")?.value || "").trim();
  const shift = String(document.getElementById("aiShift")?.value || "").trim();
  const inspector = String(document.getElementById("aiInspector")?.value || "").trim();
  const formNoEl = document.getElementById("aiFormNumber");
  if (formNoEl && !String(formNoEl.value || "").trim()) formNoEl.value = generateArtisanFormNumber();
  const formNo = String(formNoEl?.value || "").trim();
  const q = new URLSearchParams();
  if (date) q.set("date", date);
  if (assetId) q.set("asset_id", assetId);
  if (shift) q.set("shift", shift);
  if (inspector) q.set("inspector_name", inspector);
  if (formNo) q.set("form_number", formNo);
  if (download) q.set("download", "1");
  window.open(`${API}/reports/artisan-inspection-form.pdf?${q.toString()}`, "_blank");
}

async function loadArtisanInspections() {
  const list = document.getElementById("aiList");
  if (!list) return;
  list.innerHTML = `<div class="skeleton-block"></div>`;
  const asset_id = String(document.getElementById("aiFilterAsset")?.value || "").trim();
  const start = String(document.getElementById("aiStart")?.value || "").trim();
  const end = String(document.getElementById("aiEnd")?.value || "").trim();
  const q = new URLSearchParams();
  if (asset_id) q.set("asset_id", asset_id);
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  try {
    const res = await fetch(`${API}/maintenance/artisan-inspections${q.toString() ? `?${q.toString()}` : ""}`, {
      headers: authHeaders(),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load artisan inspections");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    list.innerHTML = rows.length ? rows.map(artisanInspectionCard).join("") : `<div class="muted">No artisan inspections found.</div>`;
  } catch (e) {
    list.innerHTML = `<div class="message-error">Artisan inspection load error: ${esc(e.message || e)}</div>`;
  }
}
function refreshWfDraftEditor() {
  const body = document.getElementById("wfItemsEditorBody");
  const oilTotalEl = document.getElementById("wfOilTotal");
  const partsTotalEl = document.getElementById("wfPartsTotal");
  const serviceTotalEl = document.getElementById("wfServiceTotal");
  if (!body) return;
  body.innerHTML = wfDraftItems.length
    ? wfDraftItems.map((it, idx) => `
      <tr>
        <td>${esc(it.type === "oil" ? "Oil" : "Part")}</td>
        <td>${esc(it.part_code)}</td>
        <td>${esc(it.part_name || "-")}</td>
        <td style="text-align:right;">${fmt1(it.qty)}</td>
        <td style="text-align:right;">${fmtMoney(it.unit_cost)}</td>
        <td style="text-align:right;">${fmtMoney(it.line_cost)}</td>
        <td style="text-align:right;">${fmt1(it.on_hand)}</td>
        <td style="text-align:right;"><button type="button" data-wf-item-del="${idx}">Remove</button></td>
      </tr>
    `).join("")
    : `<tr><td colspan="8" class="muted">No items added yet.</td></tr>`;
  const oilTotal = wfDraftItems.filter((x) => x.type === "oil").reduce((s, x) => s + Number(x.line_cost || 0), 0);
  const partsTotal = wfDraftItems.filter((x) => x.type !== "oil").reduce((s, x) => s + Number(x.line_cost || 0), 0);
  const serviceTotal = oilTotal + partsTotal;
  if (oilTotalEl) oilTotalEl.textContent = fmtMoney(oilTotal);
  if (partsTotalEl) partsTotalEl.textContent = fmtMoney(partsTotal);
  if (serviceTotalEl) serviceTotalEl.textContent = fmtMoney(serviceTotal);
}
function hydrateWfDraftFromSaved(planId) {
  const row = wfInputsCache.find((r) => Number(r.plan_id || 0) === Number(planId || 0));
  if (!row) {
    wfDraftItems = [];
    refreshWfDraftEditor();
    return;
  }
  let items = [];
  try {
    const parsed = JSON.parse(String(row.items_json || "[]"));
    if (Array.isArray(parsed)) items = parsed;
  } catch {}
  wfDraftItems = items.map((it) => {
    const part = getWfPartByCode(it.part_code);
    const qty = Math.max(0, Number(it.qty || 0));
    const unit = Number(part?.latest_unit_cost || 0);
    return {
      type: String(it.type || "part").toLowerCase() === "oil" ? "oil" : "part",
      part_code: String(it.part_code || "").trim(),
      part_name: String(part?.part_name || ""),
      qty,
      unit_cost: unit,
      on_hand: Number(part?.on_hand || 0),
      line_cost: qty * unit,
    };
  }).filter((x) => x.part_code && x.qty > 0);
  refreshWfDraftEditor();
}
function addWfDraftItem() {
  const msg = document.getElementById("wfInputMsg");
  const type = String(document.getElementById("wfItemType")?.value || "part").toLowerCase() === "oil" ? "oil" : "part";
  const part_code = String(document.getElementById("wfItemCode")?.value || "").trim();
  const qty = Math.max(0, Number(document.getElementById("wfItemQty")?.value || 0));
  if (!part_code || qty <= 0) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Select a part code and enter a quantity greater than 0.";
    }
    return;
  }
  const part = getWfPartByCode(part_code);
  if (!part) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Part code not found in Stores list. Please use a valid part code.";
    }
    return;
  }
  wfDraftItems.push({
    type,
    part_code: String(part.part_code || "").trim(),
    part_name: String(part.part_name || ""),
    qty,
    unit_cost: Number(part.latest_unit_cost || 0),
    on_hand: Number(part.on_hand || 0),
    line_cost: qty * Number(part.latest_unit_cost || 0),
  });
  const codeEl = document.getElementById("wfItemCode");
  const qtyEl = document.getElementById("wfItemQty");
  if (codeEl) codeEl.value = "";
  if (qtyEl) qtyEl.value = "0";
  if (msg) {
    msg.className = "muted";
    msg.textContent = "Item added.";
  }
  refreshWfDraftEditor();
}

async function loadWeeklyForumSummary() {
  const msg = document.getElementById("wfMsg");
  const kpiBody = document.getElementById("wfKpiBody");
  const upcomingBody = document.getElementById("wfUpcomingBody");
  const startEl = document.getElementById("wfStart");
  const endEl = document.getElementById("wfEnd");
  const nearEl = document.getElementById("wfNearDueHours");
  if (!msg || !kpiBody || !upcomingBody || !startEl || !endEl || !nearEl) return;

  const q = weeklyForumQueryString();

  msg.className = "muted";
  msg.textContent = "Loading weekly forum data...";
  kpiBody.innerHTML = `<tr><td colspan="2" class="muted">Loading...</td></tr>`;
  upcomingBody.innerHTML = `<tr><td colspan="11" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/summary?${q}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load weekly forum summary");

    const kpis = data.kpis || {};
    const costs = data.costs || {};
    const range = data.range || {};
    kpiBody.innerHTML = [
      ["Range", `${range.start || "-"} to ${range.end || "-"}`],
      ["Open Work Orders", Number(kpis.open_work_orders || 0)],
      ["Upcoming Services Flagged", Number(kpis.upcoming_services_flagged || 0)],
      ["Stores parts (excl. oil/lube SKUs)", fmtMoney(costs.stores_parts_cost)],
      ["Oil cost — lube log entries", fmtMoney(costs.stores_oil_from_logs)],
      ["Oil cost — WO stock (oil/lube lines)", fmtMoney(costs.stores_oil_from_work_orders)],
      ["Stores oil total", fmtMoney(costs.stores_oil_cost)],
      ["Maintenance Labor Cost", fmtMoney(costs.maintenance_labor_cost)],
      ["Weekly Total Cost", fmtMoney(costs.weekly_total_cost)],
      ["Upcoming Service Forecast Cost", fmtMoney(costs.upcoming_service_forecast_cost)],
    ].map(([k, v]) => `<tr><td>${esc(k)}</td><td style="text-align:right;">${esc(String(v))}</td></tr>`).join("");

    const rows = Array.isArray(data.upcoming_services) ? data.upcoming_services : [];
    const planSel = document.getElementById("wfInputPlan");
    const prevPlan = Number(planSel?.value || 0);
    wfUpcomingCache = rows;
    refreshWeeklyForumPlanOptions();
    if (planSel && prevPlan && wfUpcomingCache.some((x) => Number(x.plan_id || 0) === prevPlan)) {
      planSel.value = String(prevPlan);
    }
    hydrateWfDraftFromSaved(Number(planSel?.value || 0));
    upcomingBody.innerHTML = rows.length
      ? rows.map((r) => `
        <tr>
          <td>${esc(r.asset_code || "-")} - ${esc(r.asset_name || "-")}</td>
          <td>${esc(r.service_name || "-")}</td>
          <td style="text-align:right;">${fmt1(r.current_hours)}</td>
          <td style="text-align:right;">${fmt1(r.next_due_hours)}</td>
          <td style="text-align:right;">${fmt1(r.remaining_hours)}</td>
          <td>${esc(r.status || "-")}</td>
          <td style="text-align:right;">${fmt1(r?.forecast?.avg_oil_qty)}</td>
          <td style="text-align:right;">${fmtMoney(r?.forecast?.avg_oil_cost)}</td>
          <td style="text-align:right;">${fmt1(r?.forecast?.avg_parts_qty)}</td>
          <td style="text-align:right;">${fmtMoney(r?.forecast?.avg_parts_cost)}</td>
          <td style="text-align:right;">${fmtMoney(r?.forecast?.est_service_kit_cost)}</td>
        </tr>
      `).join("")
      : `<tr><td colspan="11" class="muted">No upcoming services within threshold.</td></tr>`;

    msg.className = "message-success";
    msg.textContent = "Weekly forum summary loaded.";
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = `Load error: ${e.message || e}`;
    kpiBody.innerHTML = `<tr><td colspan="2" class="message-error">${esc(e.message || String(e))}</td></tr>`;
    upcomingBody.innerHTML = `<tr><td colspan="11" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

let akpLastResponse = null;

function akpCategoryNorm(cat) {
  const t = String(cat ?? "").trim();
  return t || "Uncategorized";
}

function akpPct(num, den) {
  return den > 0 && Number.isFinite(num) ? Number(((num / den) * 100).toFixed(1)) : null;
}

function akpRollupCategoriesFromAssets(assetRows) {
  const catMap = new Map();
  for (const a of assetRows) {
    const catKey = akpCategoryNorm(a.category);
    if (!catMap.has(catKey)) {
      catMap.set(catKey, {
        category: catKey,
        scheduled_hours: 0,
        run_hours: 0,
        downtime_hours: 0,
        available_hours: 0,
        asset_ids: new Set(),
      });
    }
    const c = catMap.get(catKey);
    c.scheduled_hours += Number(a.scheduled_hours || 0);
    c.run_hours += Number(a.run_hours || 0);
    c.downtime_hours += Number(a.downtime_hours || 0);
    c.available_hours += Number(a.available_hours || 0);
    c.asset_ids.add(a.asset_id);
  }
  const rows = Array.from(catMap.values()).map((c) => {
    const sched = c.scheduled_hours;
    const avail = c.available_hours;
    const run = c.run_hours;
    return {
      category: c.category,
      asset_count: c.asset_ids.size,
      scheduled_hours: Number(sched.toFixed(2)),
      run_hours: Number(run.toFixed(2)),
      downtime_hours: Number(c.downtime_hours.toFixed(2)),
      available_hours: Number(avail.toFixed(2)),
      availability_pct: akpPct(avail, sched),
      utilization_pct: akpPct(run, avail),
    };
  });
  rows.sort((x, y) => {
    if (x.utilization_pct == null && y.utilization_pct == null) {
      return String(x.category || "").localeCompare(String(y.category || ""));
    }
    if (x.utilization_pct == null) return 1;
    if (y.utilization_pct == null) return -1;
    return y.utilization_pct - x.utilization_pct;
  });
  return rows;
}

function akpFleetFromAssets(assetRows) {
  const fleet_sched = assetRows.reduce((s, r) => s + Number(r.scheduled_hours || 0), 0);
  const fleet_avail = assetRows.reduce((s, r) => s + Number(r.available_hours || 0), 0);
  const fleet_run = assetRows.reduce((s, r) => s + Number(r.run_hours || 0), 0);
  const fleet_down = assetRows.reduce((s, r) => s + Number(r.downtime_hours || 0), 0);
  return {
    scheduled_hours: Number(fleet_sched.toFixed(2)),
    available_hours: Number(fleet_avail.toFixed(2)),
    run_hours: Number(fleet_run.toFixed(2)),
    downtime_hours: Number(fleet_down.toFixed(2)),
    availability_pct: akpPct(fleet_avail, fleet_sched),
    utilization_pct: akpPct(fleet_run, fleet_avail),
  };
}

function refreshAkpCategoryFilterOptions(data, previousValue) {
  const sel = document.getElementById("akpCategoryFilter");
  if (!sel) return;
  const cats = Array.isArray(data?.by_category) ? data.by_category : [];
  const keys = cats.map((c) => String(c.category ?? "Uncategorized"));
  const prev = previousValue != null ? String(previousValue) : String(sel.value || "");
  sel.innerHTML = `<option value="">All types</option>${keys
    .map((k) => `<option value="${escAttr(k)}">${esc(k)}</option>`)
    .join("")}`;
  if (prev && keys.includes(prev)) sel.value = prev;
  else sel.value = "";
}

function escAttr(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function renderAssetKpiTables(data) {
  const fleetEl = document.getElementById("akpFleetSummary");
  const catBody = document.getElementById("akpCategoryBody");
  const assetBody = document.getElementById("akpAssetBody");
  const filterSel = document.getElementById("akpCategoryFilter");
  if (!catBody || !assetBody || !data) return;

  const filterRaw = String(filterSel?.value || "").trim();
  const allAssets = Array.isArray(data.by_asset) ? data.by_asset : [];
  const filteredAssets = filterRaw
    ? allAssets.filter((a) => akpCategoryNorm(a.category) === filterRaw)
    : allAssets;

  const cats = filterRaw
    ? akpRollupCategoriesFromAssets(filteredAssets)
    : Array.isArray(data.by_category) ? data.by_category : [];

  const fleet = filterRaw ? akpFleetFromAssets(filteredAssets) : data.fleet || {};
  const days = Number(data.days_in_range || 0);
  if (fleetEl) {
    const label = filterRaw ? `<strong>Filtered (${esc(filterRaw)}):</strong>` : "<strong>All assets in range:</strong>";
    fleetEl.innerHTML = `${label} scheduled ${fmt1(fleet.scheduled_hours)} h, available ${fmt1(fleet.available_hours)} h, run ${fmt1(fleet.run_hours)} h, downtime ${fmt1(fleet.downtime_hours)} h — availability ${fmtPct(fleet.availability_pct)}, utilization ${fmtPct(fleet.utilization_pct)} <span class="muted">(${days} calendar days)</span>`;
  }

  catBody.innerHTML = cats.length
    ? cats.map((r) => `
        <tr>
          <td>${esc(r.category || "—")}</td>
          <td style="text-align:right;">${Number(r.asset_count || 0)}</td>
          <td style="text-align:right;">${fmt1(r.scheduled_hours)}</td>
          <td style="text-align:right;">${fmt1(r.available_hours)}</td>
          <td style="text-align:right;">${fmt1(r.run_hours)}</td>
          <td style="text-align:right;">${fmt1(r.downtime_hours)}</td>
          <td style="text-align:right;">${fmtPct(r.availability_pct)}</td>
          <td style="text-align:right;">${fmtPct(r.utilization_pct)}</td>
        </tr>
      `).join("")
    : `<tr><td colspan="8" class="muted">${filterRaw ? "No assets in this type for the range." : "No production daily hours in range (check Daily Input / dates)."}</td></tr>`;

  assetBody.innerHTML = filteredAssets.length
    ? filteredAssets.map((r) => `
        <tr>
          <td>${esc(r.asset_code || "—")} — ${esc(r.asset_name || "")}</td>
          <td>${esc(r.category || "—")}</td>
          <td>${esc(r.utilization_mode || "—")}</td>
          <td style="text-align:right;">${Number(r.days_with_data || 0)} / ${Number(r.days_in_range || 0)}</td>
          <td style="text-align:right;">${fmt1(r.scheduled_hours)}</td>
          <td style="text-align:right;">${fmt1(r.available_hours)}</td>
          <td style="text-align:right;">${fmt1(r.run_hours)}</td>
          <td style="text-align:right;">${fmt1(r.downtime_hours)}</td>
          <td style="text-align:right;">${fmtPct(r.availability_pct)}</td>
          <td style="text-align:right;">${fmtPct(r.utilization_pct)}</td>
        </tr>
      `).join("")
    : `<tr><td colspan="10" class="muted">${filterRaw ? "No rows for this type." : "No rows."}</td></tr>`;
}

async function loadAssetKpiWeekly() {
  const msg = document.getElementById("akpMsg");
  const catBody = document.getElementById("akpCategoryBody");
  const assetBody = document.getElementById("akpAssetBody");
  const fleetEl = document.getElementById("akpFleetSummary");
  const start = String(document.getElementById("akpStart")?.value || "").trim();
  const end = String(document.getElementById("akpEnd")?.value || "").trim();
  const schedEl = document.getElementById("akpScheduled");
  const sched = Math.max(0.5, Number(schedEl?.value || 10));
  const filterSel = document.getElementById("akpCategoryFilter");
  const prevFilter = String(filterSel?.value || "");
  if (!msg || !catBody || !assetBody) return;
  if (!start || !end) {
    msg.className = "message-error";
    msg.textContent = "Choose start and end dates.";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Loading KPI…";
  catBody.innerHTML = `<tr><td colspan="8" class="muted">Loading…</td></tr>`;
  assetBody.innerHTML = `<tr><td colspan="10" class="muted">Loading…</td></tr>`;
  if (fleetEl) fleetEl.textContent = "";
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  q.set("scheduled", String(sched));
  try {
    const res = await fetch(`${API}/dashboard/asset-kpi/weekly?${q.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load asset KPI");
    akpLastResponse = data;
    refreshAkpCategoryFilterOptions(data, prevFilter);
    renderAssetKpiTables(data);
    msg.className = "message-success";
    const fil = String(document.getElementById("akpCategoryFilter")?.value || "").trim();
    msg.textContent = fil
      ? `Loaded ${start} → ${end}, showing type “${fil}”.`
      : `Loaded ${start} → ${end}. Higher utilization = more run hours per available hour.`;
  } catch (e) {
    akpLastResponse = null;
    msg.className = "message-error";
    msg.textContent = `Error: ${e.message || e}`;
    catBody.innerHTML = `<tr><td colspan="8" class="message-error">${esc(e.message || String(e))}</td></tr>`;
    assetBody.innerHTML = `<tr><td colspan="10" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

async function openWeeklyForumPdf(download = false) {
  const q = weeklyForumQueryString();
  const url = `${API}/maintenance/weekly-forum.pdf?${q}${download ? "&download=1" : ""}`;
  try {
    const res = await fetch(url, { headers: authHeaders() });
    if (!res.ok) {
      const txt = await res.text();
      throw new Error(txt || `PDF request failed (${res.status})`);
    }
    const blob = await res.blob();
    const blobUrl = URL.createObjectURL(blob);
    if (download) {
      const a = document.createElement("a");
      const dateTag = new Date().toISOString().slice(0, 10);
      a.href = blobUrl;
      a.download = `weekly-forum-${dateTag}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(blobUrl), 3000);
      return;
    }
    window.open(blobUrl, "_blank");
    setTimeout(() => URL.revokeObjectURL(blobUrl), 10000);
  } catch (e) {
    alert(`Weekly Forum PDF error: ${e.message || e}`);
  }
}

function wfStatusLabel(s) {
  const v = String(s || "open").toLowerCase();
  if (v === "in_progress") return "In Progress";
  if (v === "blocked") return "Blocked";
  if (v === "done") return "Done";
  return "Open";
}

async function loadWeeklyForumActions() {
  const body = document.getElementById("wfActionBody");
  const start = String(document.getElementById("wfStart")?.value || "").trim();
  const end = String(document.getElementById("wfEnd")?.value || "").trim();
  if (!body) return;
  body.innerHTML = `<tr><td colspan="8" class="muted">Loading...</td></tr>`;
  const q = new URLSearchParams();
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/actions?${q.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load actions");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    body.innerHTML = rows.length
      ? rows.map((r) => `
        <tr>
          <td>${Number(r.id || 0)}</td>
          <td>${esc(r.action_date || "-")}</td>
          <td>${esc(r.department || "-")}</td>
          <td>${esc(r.action_item || "-")}</td>
          <td>${esc(r.owner_name || "-")}</td>
          <td>${esc(r.due_date || "-")}</td>
          <td>
            <select data-wf-action-status="${Number(r.id || 0)}">
              <option value="open" ${String(r.status) === "open" ? "selected" : ""}>Open</option>
              <option value="in_progress" ${String(r.status) === "in_progress" ? "selected" : ""}>In Progress</option>
              <option value="blocked" ${String(r.status) === "blocked" ? "selected" : ""}>Blocked</option>
              <option value="done" ${String(r.status) === "done" ? "selected" : ""}>Done</option>
            </select>
          </td>
          <td>${esc(r.notes || "")}</td>
        </tr>
      `).join("")
      : `<tr><td colspan="8" class="muted">No actions logged for selected range.</td></tr>`;
  } catch (e) {
    body.innerHTML = `<tr><td colspan="8" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

async function saveWeeklyForumAction() {
  const msg = document.getElementById("wfActionMsg");
  const action_date = String(document.getElementById("wfActionDate")?.value || "").trim();
  const department = String(document.getElementById("wfActionDept")?.value || "").trim();
  const owner_name = String(document.getElementById("wfActionOwner")?.value || "").trim();
  const due_date = String(document.getElementById("wfActionDue")?.value || "").trim();
  const status = String(document.getElementById("wfActionStatus")?.value || "open").trim().toLowerCase();
  const action_item = String(document.getElementById("wfActionItem")?.value || "").trim();
  const notes = String(document.getElementById("wfActionNotes")?.value || "").trim();
  if (!msg) return;

  if (!action_date || !department || !owner_name || !action_item) {
    msg.className = "message-error";
    msg.textContent = "Action date, department, owner, and action item are required.";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Saving action...";
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/actions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        action_date,
        department,
        action_item,
        owner_name,
        due_date: due_date || null,
        status,
        notes: notes || null,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save action");
    msg.className = "message-success";
    msg.textContent = "Action saved.";
    document.getElementById("wfActionItem").value = "";
    document.getElementById("wfActionNotes").value = "";
    await loadWeeklyForumActions();
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = e.message || String(e);
  }
}

async function updateWeeklyForumActionStatus(id, status) {
  const n = Number(id || 0);
  if (!n) return;
  await fetch(`${API}/maintenance/weekly-forum/actions/${n}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ status }),
  });
}

async function loadWeeklyForumParts() {
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/parts`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load parts list");
    wfPartsCache = Array.isArray(data.rows) ? data.rows : [];
    refreshWeeklyForumPartsDatalist();
    hydrateWfDraftFromSaved(Number(document.getElementById("wfInputPlan")?.value || 0));
  } catch {
    wfPartsCache = [];
    refreshWeeklyForumPartsDatalist();
  }
}

async function loadWeeklyForumInputs() {
  const body = document.getElementById("wfInputsBody");
  if (body) body.innerHTML = `<tr><td colspan="4" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/forecast-inputs`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load manual inputs");
    wfInputsCache = Array.isArray(data.rows) ? data.rows : [];
    refreshWeeklyForumInputsTable();
    const planNow = Number(document.getElementById("wfInputPlan")?.value || 0);
    if (planNow) hydrateWfDraftFromSaved(planNow);
  } catch (e) {
    if (body) body.innerHTML = `<tr><td colspan="4" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

async function saveWeeklyForumInput() {
  const msg = document.getElementById("wfInputMsg");
  const plan_id = Number(document.getElementById("wfInputPlan")?.value || 0);
  const notes = String(document.getElementById("wfInputNotes")?.value || "").trim();
  const items = wfDraftItems.map((x) => ({
    type: x.type === "oil" ? "oil" : "part",
    part_code: String(x.part_code || "").trim(),
    qty: Math.max(0, Number(x.qty || 0)),
  })).filter((x) => x.part_code && x.qty > 0);
  if (!msg) return;
  if (!plan_id) {
    msg.className = "message-error";
    msg.textContent = "Select an upcoming service plan first.";
    return;
  }
  if (!items.length) {
    msg.className = "message-error";
    msg.textContent = "Enter at least one oil or part line: part_code,qty";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Saving manual input...";
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/forecast-inputs`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ plan_id, items, notes: notes || null }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save input");
    msg.className = "message-success";
    msg.textContent = "Manual input saved. Forecast cost now uses stores pricing.";
    await loadWeeklyForumInputs();
    await loadWeeklyForumSummary();
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = e.message || String(e);
  }
}

function parseLines(text) {
  return String(text || "")
    .split("\n")
    .map((s) => s.trim())
    .filter(Boolean);
}

function parseCsvLine(line) {
  const out = [];
  let cur = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    const nxt = line[i + 1];
    if (ch === '"' && inQuotes && nxt === '"') {
      cur += '"';
      i += 1;
      continue;
    }
    if (ch === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (ch === "," && !inQuotes) {
      out.push(cur);
      cur = "";
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out.map((s) => String(s || "").trim());
}

function parsePackedItems(s, fallbackUnit) {
  return String(s || "")
    .split("||")
    .map((x) => String(x || "").trim())
    .filter(Boolean)
    .map((chunk) => {
      const [nameRaw, qtyRaw, unitRaw, hintRaw] = chunk.split("|").map((v) => String(v || "").trim());
      const qty = Number(qtyRaw || 0);
      return {
        name: nameRaw,
        qty: Number.isFinite(qty) ? qty : 0,
        unit: unitRaw || fallbackUnit,
        part_hint: hintRaw || nameRaw,
      };
    })
    .filter((x) => x.name && x.qty > 0);
}

function parseRsgItems(text, defaultUnit) {
  return parseLines(text).map((line) => {
    const [nameRaw, qtyRaw, unitRaw, hintRaw] = line.split("|").map((s) => String(s || "").trim());
    const qty = Number(qtyRaw || 0);
    return {
      name: nameRaw,
      qty: Number.isFinite(qty) ? qty : 0,
      unit: unitRaw || defaultUnit,
      part_hint: hintRaw || nameRaw,
    };
  }).filter((x) => x.name && x.qty > 0);
}

function itemLines(items = []) {
  return (Array.isArray(items) ? items : [])
    .map((x) => `${String(x?.name || "").trim()}|${Number(x?.qty || 0)}|${String(x?.unit || "").trim()}|${String(x?.part_hint || x?.name || "").trim()}`)
    .join("\n");
}

function slug(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function rsgProfileRow(r) {
  return `
    <tr>
      <td>${esc(r.key || "-")}</td>
      <td>${esc(r.make || "-")}</td>
      <td>${esc((r.modelContains || []).join(", ") || "-")}</td>
      <td style="text-align:right;">${Number(r.service_hours || 0)}</td>
      <td>${esc(r.title || "-")}</td>
      <td style="text-align:right;">${Array.isArray(r.oils) ? r.oils.length : 0}</td>
      <td style="text-align:right;">${Array.isArray(r.filters) ? r.filters.length : 0}</td>
      <td><button data-rsg-edit="${esc(r.key || "")}">Edit</button></td>
    </tr>
  `;
}

let rsgProfilesCache = [];
async function loadRsgProfiles() {
  const body = document.getElementById("rsgProfilesBody");
  if (!body) return;
  body.innerHTML = `<tr><td colspan="8" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/ironmind/rsg/profiles`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load RSG profiles");
    rsgProfilesCache = Array.isArray(data.rows) ? data.rows : [];
    body.innerHTML = rsgProfilesCache.length
      ? rsgProfilesCache.map(rsgProfileRow).join("")
      : `<tr><td colspan="8" class="muted">No saved profiles yet.</td></tr>`;
  } catch (e) {
    body.innerHTML = `<tr><td colspan="8" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

function fillRsgProfileForm(profileKey) {
  const p = rsgProfilesCache.find((x) => String(x.key || "") === String(profileKey || ""));
  if (!p) return;
  const set = (id, v) => {
    const el = document.getElementById(id);
    if (el) el.value = v == null ? "" : String(v);
  };
  set("rsgProfileKey", p.key || "");
  set("rsgMake", p.make || "");
  set("rsgModelMatch", Array.isArray(p.modelContains) ? p.modelContains.join(",") : "");
  set("rsgServiceHours", Number(p.service_hours || 0) || 0);
  set("rsgTitle", p.title || "");
  set("rsgTasks", parseLines((p.tasks || []).join("\n")).join("\n"));
  set("rsgChecks", parseLines((p.checks || []).join("\n")).join("\n"));
  set("rsgPostChecks", parseLines((p.post_service_checks || []).join("\n")).join("\n"));
  set("rsgSafety", parseLines((p.safety || []).join("\n")).join("\n"));
  set("rsgOils", itemLines(p.oils || []));
  set("rsgFilters", itemLines(p.filters || []));
}

async function saveRsgProfile() {
  const msg = document.getElementById("rsgProfileMsg");
  const make = String(document.getElementById("rsgMake")?.value || "").trim();
  const model_match = String(document.getElementById("rsgModelMatch")?.value || "").trim();
  const service_hours = Math.max(1, Number(document.getElementById("rsgServiceHours")?.value || 0));
  const title = String(document.getElementById("rsgTitle")?.value || "").trim();
  let profile_key = String(document.getElementById("rsgProfileKey")?.value || "").trim().toLowerCase();
  const tasks = parseLines(document.getElementById("rsgTasks")?.value || "");
  const checks = parseLines(document.getElementById("rsgChecks")?.value || "");
  const post_service_checks = parseLines(document.getElementById("rsgPostChecks")?.value || "");
  const safety = parseLines(document.getElementById("rsgSafety")?.value || "");
  const oils = parseRsgItems(document.getElementById("rsgOils")?.value || "", "L");
  const filters = parseRsgItems(document.getElementById("rsgFilters")?.value || "", "ea");
  if (!profile_key) {
    profile_key = `${slug(make)}-${slug(model_match || "model")}-${Number(service_hours || 0)}`;
    const keyEl = document.getElementById("rsgProfileKey");
    if (keyEl) keyEl.value = profile_key;
  }
  if (!profile_key || !title || !service_hours) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Profile key, service hours, and title are required.";
    }
    return;
  }
  if (msg) {
    msg.className = "muted";
    msg.textContent = "Saving profile...";
  }
  try {
    const res = await fetch(`${API}/ironmind/rsg/profiles`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        profile_key,
        make,
        model_match,
        service_hours,
        title,
        tasks,
        checks,
        post_service_checks,
        safety,
        oils,
        filters,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save profile");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `Profile saved: ${profile_key}`;
    }
    await loadRsgProfiles();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = e.message || String(e);
    }
  }
}

function downloadRsgCsvTemplate() {
  window.open(`${API}/ironmind/rsg/profiles/template.csv`, "_blank");
}

async function importRsgProfilesCsv() {
  const msg = document.getElementById("rsgProfileMsg");
  const file = document.getElementById("rsgCsvFile")?.files?.[0];
  if (!file) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Choose a CSV file first.";
    }
    return;
  }
  if (msg) {
    msg.className = "muted";
    msg.textContent = "Parsing CSV and importing...";
  }
  const txt = await file.text();
  const lines = String(txt || "").split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  if (lines.length < 2) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "CSV has no data rows.";
    }
    return;
  }
  const header = parseCsvLine(lines[0]).map((h) => h.toLowerCase());
  const idx = (name) => header.indexOf(String(name).toLowerCase());
  const get = (arr, name) => {
    const i = idx(name);
    return i >= 0 ? String(arr[i] || "").trim() : "";
  };
  const profiles = [];
  for (let i = 1; i < lines.length; i += 1) {
    const cols = parseCsvLine(lines[i]);
    const profile_key = get(cols, "profile_key");
    const service_hours = Number(get(cols, "service_hours") || 0);
    const title = get(cols, "title");
    if (!profile_key || !service_hours || !title) continue;
    profiles.push({
      profile_key,
      make: get(cols, "make"),
      model_match: get(cols, "model_match"),
      service_hours,
      title,
      tasks: parseLines(get(cols, "tasks").split("|").join("\n")),
      checks: parseLines(get(cols, "checks").split("|").join("\n")),
      post_service_checks: parseLines(get(cols, "post_service_checks").split("|").join("\n")),
      safety: parseLines(get(cols, "safety").split("|").join("\n")),
      oils: parsePackedItems(get(cols, "oils"), "L"),
      filters: parsePackedItems(get(cols, "filters"), "ea"),
    });
  }
  try {
    const res = await fetch(`${API}/ironmind/rsg/profiles/import`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ profiles }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Import failed");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `CSV import complete: ${Number(data.upserted || 0)} profile(s) upserted.`;
    }
    await loadRsgProfiles();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = e.message || String(e);
    }
  }
}

document.addEventListener("DOMContentLoaded", () => {
  if (!ensureMaintenanceAccess()) return;
  console.log("Maintenance UI loaded");

  const generateBtn = document.getElementById("generateBtn");
  const savePlanBtn = document.getElementById("savePlanBtn");
  const saveBackfillBtn = document.getElementById("saveBackfillBtn");
  const inspectproRefreshStatusBtn = document.getElementById("inspectproRefreshStatusBtn");
  const backfillBody = document.getElementById("backfillBody");
  const useLiveEl = document.getElementById("planUseLiveForLastService");

  syncDueThresholdInput();
  if (generateBtn) {
    generateBtn.addEventListener("click", generateWO);
  }
  document.getElementById("applyDueThresholdBtn")?.addEventListener("click", async () => {
    const v = getDueThresholdHours();
    localStorage.setItem(MAINT_DUE_THRESHOLD_KEY, String(v));
    await loadDue();
  });
  document.getElementById("openUpcomingServicesPdfBtn")?.addEventListener("click", () => openUpcomingServicesPdf(false));
  document.getElementById("downloadUpcomingServicesPdfBtn")?.addEventListener("click", () => openUpcomingServicesPdf(true));

  if (savePlanBtn) {
    savePlanBtn.addEventListener("click", savePlan);
  }
  if (saveBackfillBtn) {
    saveBackfillBtn.addEventListener("click", saveBackfillHistory);
  }
  if (inspectproRefreshStatusBtn) {
    inspectproRefreshStatusBtn.addEventListener("click", () => {
      loadInspectproStatus().catch(() => {});
    });
  }
  if (backfillBody) {
    backfillBody.addEventListener("click", (e) => {
      const btn = e.target instanceof HTMLElement ? e.target.closest("button[data-backfill-action]") : null;
      if (!btn) return;
      const tr = btn.closest("tr[data-backfill-id]");
      const id = Number(tr?.getAttribute("data-backfill-id") || 0);
      const action = String(btn.getAttribute("data-backfill-action") || "");
      if (!id || !action) return;
      if (action === "edit") {
        editBackfillHistory(id).catch((err) => alert(err.message || err));
      } else if (action === "delete") {
        deleteBackfillHistory(id).catch((err) => alert(err.message || err));
      }
    });
  }

  const assetEl = document.getElementById("planAsset");
  if (assetEl) {
    assetEl.addEventListener("change", loadLiveHoursForSelectedAsset);
  }

  if (useLiveEl) {
    useLiveEl.addEventListener("change", syncLastServiceHoursFromLive);
  }

  loadAssetsForPlan().then(() => {
    loadLiveHoursForSelectedAsset();
  });

  loadPlans();
  loadDue();
  loadHistory();
  loadBackfillHistory();
  loadInspectproStatus();
  setInterval(() => {
    loadInspectproStatus().catch(() => {});
  }, 20000);
  const miDate = document.getElementById("miDate");
  const aiDate = document.getElementById("aiDate");
  const aiFormNo = document.getElementById("aiFormNumber");
  const drDate = document.getElementById("drDate");
  if (miDate && !miDate.value) miDate.value = new Date().toISOString().slice(0, 10);
  if (aiDate && !aiDate.value) aiDate.value = new Date().toISOString().slice(0, 10);
  if (aiFormNo && !aiFormNo.value) aiFormNo.value = generateArtisanFormNumber();
  if (drDate && !drDate.value) drDate.value = new Date().toISOString().slice(0, 10);
  const miStart = document.getElementById("miStart");
  const miEnd = document.getElementById("miEnd");
  const aiStart = document.getElementById("aiStart");
  const aiEnd = document.getElementById("aiEnd");
  const drStart = document.getElementById("drStart");
  const drEnd = document.getElementById("drEnd");
  if (miStart && !miStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    miStart.value = d.toISOString().slice(0, 10);
  }
  if (miEnd && !miEnd.value) miEnd.value = new Date().toISOString().slice(0, 10);
  if (aiStart && !aiStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    aiStart.value = d.toISOString().slice(0, 10);
  }
  if (aiEnd && !aiEnd.value) aiEnd.value = new Date().toISOString().slice(0, 10);
  if (drStart && !drStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    drStart.value = d.toISOString().slice(0, 10);
  }
  if (drEnd && !drEnd.value) drEnd.value = new Date().toISOString().slice(0, 10);
  const wfStart = document.getElementById("wfStart");
  const wfEnd = document.getElementById("wfEnd");
  if (wfStart && !wfStart.value) {
    const d = new Date();
    const day = d.getDay();
    const mondayOffset = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + mondayOffset);
    wfStart.value = d.toISOString().slice(0, 10);
  }
  if (wfEnd && !wfEnd.value) {
    const d = new Date();
    const day = d.getDay();
    const mondayOffset = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + mondayOffset + 4);
    wfEnd.value = d.toISOString().slice(0, 10);
  }
  const akpStart = document.getElementById("akpStart");
  const akpEnd = document.getElementById("akpEnd");
  if (akpStart && !akpStart.value && wfStart?.value) akpStart.value = wfStart.value;
  if (akpEnd && !akpEnd.value && wfEnd?.value) akpEnd.value = wfEnd.value;
  if (akpStart && !akpStart.value) {
    const d = new Date();
    const day = d.getDay();
    const mondayOffset = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + mondayOffset);
    akpStart.value = d.toISOString().slice(0, 10);
  }
  if (akpEnd && !akpEnd.value) {
    const d = new Date();
    const day = d.getDay();
    const mondayOffset = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + mondayOffset + 4);
    akpEnd.value = d.toISOString().slice(0, 10);
  }
  const wfActionDate = document.getElementById("wfActionDate");
  if (wfActionDate && !wfActionDate.value) wfActionDate.value = new Date().toISOString().slice(0, 10);
  const histViewMode = document.getElementById("histViewMode");
  const histClosestLimitWrap = document.getElementById("histClosestLimitWrap");
  if (histClosestLimitWrap) histClosestLimitWrap.style.display = String(histViewMode?.value || "all") === "closest" ? "" : "none";
  const histEventDate = document.getElementById("histEventDate");
  if (histEventDate && !histEventDate.value) histEventDate.value = new Date().toISOString().slice(0, 10);
  const histFilterStart = document.getElementById("histFilterStart");
  const histFilterEnd = document.getElementById("histFilterEnd");
  if (histFilterStart && !histFilterStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    histFilterStart.value = d.toISOString().slice(0, 10);
  }
  if (histFilterEnd && !histFilterEnd.value) histFilterEnd.value = new Date().toISOString().slice(0, 10);
  const mpWeekStart = document.getElementById("mpWeekStart");
  const mpWeekEnd = document.getElementById("mpWeekEnd");
  const mpMonth = document.getElementById("mpMonth");
  if (mpWeekStart && !mpWeekStart.value) {
    const w = mpWeekRangeLabel();
    mpWeekStart.value = w.start;
  }
  if (mpWeekEnd && !mpWeekEnd.value) {
    const w = mpWeekRangeLabel();
    mpWeekEnd.value = w.end;
  }
  if (mpMonth && !mpMonth.value) mpMonth.value = mpMonthLabel();
  renderManagerInspectionChecklist();
  renderArtisanInspectionChecklist();
  const miPartsBody = document.getElementById("miPartsBody");
  if (miPartsBody && !miPartsBody.querySelector("tr")) addManagerInspectionPartRow();
  document.getElementById("miPullLiveHoursBtn")?.addEventListener("click", () =>
    pullManagerInspectionLiveHours().catch((e) => console.error(e))
  );
  document.getElementById("miAddPartRowBtn")?.addEventListener("click", () => addManagerInspectionPartRow());
  const miReloadHours = () => {
    const id = Number(document.getElementById("miAsset")?.value || 0);
    if (!id) {
      const meta = document.getElementById("miLiveMeta");
      const inp = document.getElementById("miMachineHours");
      if (meta) meta.textContent = "";
      if (inp) inp.value = "";
      return;
    }
    pullManagerInspectionLiveHours().catch((e) => console.error(e));
  };
  document.getElementById("miAsset")?.addEventListener("change", miReloadHours);
  document.getElementById("miDate")?.addEventListener("change", miReloadHours);
  const aiReloadHours = () => {
    const id = Number(document.getElementById("aiAsset")?.value || 0);
    if (!id) {
      const meta = document.getElementById("aiLiveMeta");
      const inp = document.getElementById("aiMachineHours");
      if (meta) meta.textContent = "";
      if (inp) inp.value = "";
      return;
    }
    pullArtisanInspectionLiveHours().catch((e) => console.error(e));
  };
  document.getElementById("aiPullLiveHoursBtn")?.addEventListener("click", () =>
    pullArtisanInspectionLiveHours().catch((e) => console.error(e))
  );
  document.getElementById("aiAsset")?.addEventListener("change", aiReloadHours);
  document.getElementById("aiDate")?.addEventListener("change", aiReloadHours);
  document.getElementById("openAiBlankPdfBtn")?.addEventListener("click", () => openArtisanBlankFormPdf(false));
  document.getElementById("downloadAiBlankPdfBtn")?.addEventListener("click", () => openArtisanBlankFormPdf(true));
  document.getElementById("saveMiBtn")?.addEventListener("click", saveManagerInspection);
  document.getElementById("saveAiBtn")?.addEventListener("click", saveArtisanInspection);
  document.getElementById("saveDrBtn")?.addEventListener("click", saveDamageReport);
  document.getElementById("loadMiBtn")?.addEventListener("click", loadManagerInspections);
  document.getElementById("loadAiBtn")?.addEventListener("click", loadArtisanInspections);
  document.getElementById("loadDrBtn")?.addEventListener("click", loadDamageReports);
  document.getElementById("openDrBulkPdfBtn")?.addEventListener("click", () => openDamageReportsBulkPdf(false));
  document.getElementById("downloadDrBulkPdfBtn")?.addEventListener("click", () => openDamageReportsBulkPdf(true));
  document.getElementById("openDrBulkPdfPhotosBtn")?.addEventListener("click", () => openDamageReportsBulkPdf(false, true));
  document.getElementById("downloadDrBulkPdfPhotosBtn")?.addEventListener("click", () => openDamageReportsBulkPdf(true, true));
  document.getElementById("openMiBulkPdfBtn")?.addEventListener("click", () => openManagerInspectionsBulkPdf(false));
  document.getElementById("downloadMiBulkPdfBtn")?.addEventListener("click", () => openManagerInspectionsBulkPdf(true));
  document.getElementById("openMiBulkPdfPhotosBtn")?.addEventListener("click", () => openManagerInspectionsBulkPdf(false, true));
  document.getElementById("downloadMiBulkPdfPhotosBtn")?.addEventListener("click", () => openManagerInspectionsBulkPdf(true, true));
  document.getElementById("miList")?.addEventListener("click", (evt) => {
    const openPdf = evt.target?.closest?.("button[data-mi-open-pdf]");
    if (openPdf) {
      const id = Number(openPdf.getAttribute("data-mi-open-pdf") || 0);
      if (id) openManagerInspectionPdf(id, false);
      return;
    }
    const dlPdf = evt.target?.closest?.("button[data-mi-download-pdf]");
    if (dlPdf) {
      const id = Number(dlPdf.getAttribute("data-mi-download-pdf") || 0);
      if (id) openManagerInspectionPdf(id, true);
      return;
    }
    const btn = evt.target?.closest?.("button[data-mi-upload]");
    if (!btn) return;
    const id = Number(btn.getAttribute("data-mi-upload") || 0);
    if (!id) return;
    uploadInspectionPhoto(id).catch((e) => alert(`Photo upload failed: ${e.message || e}`));
  });
  document.getElementById("aiList")?.addEventListener("click", (evt) => {
    const openPdf = evt.target?.closest?.("button[data-ai-open-pdf]");
    if (openPdf) {
      const id = Number(openPdf.getAttribute("data-ai-open-pdf") || 0);
      if (id) openArtisanInspectionPdf(id, false);
      return;
    }
    const dlPdf = evt.target?.closest?.("button[data-ai-download-pdf]");
    if (dlPdf) {
      const id = Number(dlPdf.getAttribute("data-ai-download-pdf") || 0);
      if (id) openArtisanInspectionPdf(id, true);
    }
  });
  document.getElementById("drList")?.addEventListener("click", (evt) => {
    const openPdf = evt.target?.closest?.("button[data-dr-open-pdf]");
    if (openPdf) {
      const id = Number(openPdf.getAttribute("data-dr-open-pdf") || 0);
      if (id) openDamageReportPdf(id, false);
      return;
    }
    const dlPdf = evt.target?.closest?.("button[data-dr-download-pdf]");
    if (dlPdf) {
      const id = Number(dlPdf.getAttribute("data-dr-download-pdf") || 0);
      if (id) openDamageReportPdf(id, true);
      return;
    }
    const btn = evt.target?.closest?.("button[data-dr-upload]");
    if (!btn) return;
    const id = Number(btn.getAttribute("data-dr-upload") || 0);
    if (!id) return;
    uploadDamagePhoto(id).catch((e) => alert(`Damage photo upload failed: ${e.message || e}`));
  });
  loadAssetsForInspection().catch(() => {});
  loadManagerInspections().catch(() => {});
  loadArtisanInspections().catch(() => {});
  loadDamageReports().catch(() => {});
  loadWeeklyForumSummary().catch(() => {});
  loadWeeklyForumActions().catch(() => {});
  loadWeeklyForumParts().catch(() => {});
  loadWeeklyForumInputs().catch(() => {});
  loadMaintenancePackStatus().catch(() => {});
  loadRsgProfiles().catch(() => {});
  document.getElementById("syncStatsBtn")?.addEventListener("click", syncLoadStats);
  document.getElementById("syncStateBtn")?.addEventListener("click", syncLoadState);
  document.getElementById("syncOutboxBtn")?.addEventListener("click", syncLoadOutbox);
  document.getElementById("syncPullBtn")?.addEventListener("click", syncPullEvents);
  document.getElementById("syncExportLastPullBtn")?.addEventListener("click", syncExportLastPull);
  document.getElementById("syncApplyBtn")?.addEventListener("click", syncApplyLastPull);
  document.getElementById("syncApplyDryRunBtn")?.addEventListener("click", syncApplyLastPullDryRun);
  document.getElementById("syncAckBtn")?.addEventListener("click", syncAckLastPull);
  document.getElementById("syncCheckpointBtn")?.addEventListener("click", syncCheckpointLastPull);
  document.getElementById("showMainMaintBtn")?.addEventListener("click", () => setTopView("main"));
  document.getElementById("showManagerInspectionsBtn")?.addEventListener("click", () => setTopView("mi"));
  document.getElementById("showArtisanInspectionsBtn")?.addEventListener("click", () => setTopView("ai"));
  document.getElementById("showWeeklyForumBtn")?.addEventListener("click", () => setTopView("wf"));
  document.getElementById("showAssetKpiBtn")?.addEventListener("click", () => setTopView("kpi"));
  document.getElementById("showHistogramBtn")?.addEventListener("click", () => setTopView("hist"));
  document.getElementById("showSyncAdminBtn")?.addEventListener("click", () => setTopView("sync"));
  document.getElementById("histViewMode")?.addEventListener("change", () => loadHistory());
  document.getElementById("histClosestLimit")?.addEventListener("change", () => loadHistory());
  document.getElementById("saveHistogramBtn")?.addEventListener("click", () => saveHistogramEvent());
  document.getElementById("loadHistogramBtn")?.addEventListener("click", () => loadHistogramEvents());
  document.getElementById("openHistogramPdfBtn")?.addEventListener("click", () => openHistogramPdf(false));
  document.getElementById("downloadHistogramPdfBtn")?.addEventListener("click", () => openHistogramPdf(true));
  document.getElementById("histEventBody")?.addEventListener("click", (evt) => {
    const editBtn = evt.target?.closest?.("button[data-hist-edit]");
    if (editBtn) {
      editHistogramEvent(Number(editBtn.getAttribute("data-hist-edit") || 0));
      return;
    }
    const delBtn = evt.target?.closest?.("button[data-hist-del]");
    if (delBtn) {
      deleteHistogramEvent(Number(delBtn.getAttribute("data-hist-del") || 0));
    }
  });
  document.getElementById("loadAssetKpiBtn")?.addEventListener("click", () => loadAssetKpiWeekly());
  document.getElementById("mpRefreshStatusBtn")?.addEventListener("click", () => loadMaintenancePackStatus());
  document.getElementById("mpStatusBody")?.addEventListener("click", (evt) => {
    const gen = evt.target?.closest?.("button[data-mp-gen]");
    if (gen) {
      mpGenerate(String(gen.getAttribute("data-mp-gen") || "weekly"));
      return;
    }
    const open = evt.target?.closest?.("button[data-mp-open]");
    if (open) {
      openMaintenancePackLatest(String(open.getAttribute("data-mp-open") || "weekly"), false);
      return;
    }
    const dl = evt.target?.closest?.("button[data-mp-download]");
    if (dl) {
      openMaintenancePackLatest(String(dl.getAttribute("data-mp-download") || "weekly"), true);
    }
  });
  document.getElementById("akpCategoryFilter")?.addEventListener("change", () => {
    if (akpLastResponse) renderAssetKpiTables(akpLastResponse);
  });
  document.getElementById("loadWeeklyForumBtn")?.addEventListener("click", loadWeeklyForumSummary);
  document.getElementById("saveRsgProfileBtn")?.addEventListener("click", saveRsgProfile);
  document.getElementById("loadRsgProfilesBtn")?.addEventListener("click", loadRsgProfiles);
  document.getElementById("downloadRsgCsvTemplateBtn")?.addEventListener("click", downloadRsgCsvTemplate);
  document.getElementById("importRsgCsvBtn")?.addEventListener("click", importRsgProfilesCsv);
  document.getElementById("saveWfActionBtn")?.addEventListener("click", saveWeeklyForumAction);
  document.getElementById("loadWfActionsBtn")?.addEventListener("click", loadWeeklyForumActions);
  document.getElementById("saveWfInputBtn")?.addEventListener("click", saveWeeklyForumInput);
  document.getElementById("loadWfInputsBtn")?.addEventListener("click", loadWeeklyForumInputs);
  document.getElementById("wfAddItemBtn")?.addEventListener("click", addWfDraftItem);
  document.getElementById("wfInputPlan")?.addEventListener("change", (evt) => {
    const planId = Number(evt.target?.value || 0);
    hydrateWfDraftFromSaved(planId);
  });
  document.getElementById("wfItemsEditorBody")?.addEventListener("click", (evt) => {
    const btn = evt.target?.closest?.("button[data-wf-item-del]");
    if (!btn) return;
    const idx = Number(btn.getAttribute("data-wf-item-del") || -1);
    if (idx < 0 || idx >= wfDraftItems.length) return;
    wfDraftItems.splice(idx, 1);
    refreshWfDraftEditor();
  });
  document.getElementById("openWeeklyForumPdfBtn")?.addEventListener("click", () => openWeeklyForumPdf(false));
  document.getElementById("downloadWeeklyForumPdfBtn")?.addEventListener("click", () => openWeeklyForumPdf(true));
  document.getElementById("wfActionBody")?.addEventListener("change", (evt) => {
    const sel = evt.target?.closest?.("select[data-wf-action-status]");
    if (!sel) return;
    const id = Number(sel.getAttribute("data-wf-action-status") || 0);
    const status = String(sel.value || "open");
    updateWeeklyForumActionStatus(id, status)
      .then(() => loadWeeklyForumActions())
      .catch((e) => alert(`Failed to update status: ${e.message || e}`));
  });
  document.getElementById("rsgProfilesBody")?.addEventListener("click", (evt) => {
    const btn = evt.target?.closest?.("button[data-rsg-edit]");
    if (!btn) return;
    fillRsgProfileForm(btn.getAttribute("data-rsg-edit") || "");
  });
  setTopView("main");
  loadHistogramEvents().catch(() => {});
});