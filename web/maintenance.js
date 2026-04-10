const API = "/api";
let lastSyncPull = { last_id: 0, events: [] };
const ROLE_KEY = "ironlog_session_role";
const ROLES_KEY = "ironlog_session_roles";
const USER_KEY = "ironlog_session_user";
const SITE_KEY = "ironlog_session_site";
const TOKEN_KEY = "ironlog_auth_token";
const MAINT_DUE_THRESHOLD_KEY = "ironlog_maintenance_due_threshold_hours";

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
  if (!body) return;
  body.innerHTML = `<tr><td colspan="7" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/maintenance/history`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load history");
    if (meta) meta.textContent = `As of: ${data.as_of || "-"}`;
    const rows = Array.isArray(data.rows) ? data.rows : [];
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
  const wf = document.getElementById("weeklyForumCard");
  const sync = document.getElementById("syncAdminCard");
  const btnMain = document.getElementById("showMainMaintBtn");
  const btnMi = document.getElementById("showManagerInspectionsBtn");
  const btnWf = document.getElementById("showWeeklyForumBtn");
  const btnSync = document.getElementById("showSyncAdminBtn");
  if (!main || !mi || !wf || !sync) return;

  main.style.display = view === "main" ? "" : "none";
  mi.style.display = view === "mi" ? "" : "none";
  wf.style.display = view === "wf" ? "" : "none";
  sync.style.display = view === "sync" ? "" : "none";

  const styleBtn = (btn, active) => {
    if (!btn) return;
    btn.style.borderColor = active ? "#3b82f6" : "";
    btn.style.background = active ? "#13233c" : "";
    btn.style.color = active ? "#fff" : "";
  };
  styleBtn(btnMain, view === "main");
  styleBtn(btnMi, view === "mi");
  styleBtn(btnWf, view === "wf");
  styleBtn(btnSync, view === "sync");
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
    if (drA) drA.innerHTML = `<option value="">Select asset</option>${opts}`;
    if (drF) drF.innerHTML = `<option value="">All assets</option>${opts}`;
  } catch (e) {
    selA.innerHTML = `<option value="">Assets load failed</option>`;
    selF.innerHTML = `<option value="">Assets load failed</option>`;
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
  if (!asset_id) return alert("Select an asset.");
  if (!inspection_date) return alert("Select inspection date.");
  if (msg) msg.textContent = "Saving inspection...";
  try {
    const res = await fetch(`${API}/maintenance/inspections`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ asset_id, inspection_date, inspector_name, notes }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save inspection");
    if (msg) msg.textContent = "Inspection saved.";
    document.getElementById("miNotes").value = "";
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

  return `
    <div class="card">
      <div><b>${esc(r.asset_code)}</b> - ${esc(r.asset_name || "")}</div>
      <div><small>Date: ${esc(r.inspection_date)} | Inspector: ${esc(r.inspector_name || "-")}</small></div>
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
    const res = await fetch(`${API}/maintenance/inspections${q.toString() ? `?${q.toString()}` : ""}`);
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
  const dl = document.getElementById("wfPartsList");
  if (!dl) return;
  dl.innerHTML = wfPartsCache.map((p) => {
    const code = String(p.part_code || "").trim();
    const desc = `${code} - ${String(p.part_name || "")} | on hand ${Number(p.on_hand || 0).toFixed(2)} | unit ${Number(p.latest_unit_cost || 0).toFixed(2)}`;
    return `<option value="${esc(code)}">${esc(desc)}</option>`;
  }).join("");
}
function getWfPartByCode(codeIn) {
  const code = String(codeIn || "").trim().toUpperCase();
  if (!code) return null;
  return wfPartsCache.find((p) => String(p.part_code || "").trim().toUpperCase() === code) || null;
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
  const drDate = document.getElementById("drDate");
  if (miDate && !miDate.value) miDate.value = new Date().toISOString().slice(0, 10);
  if (drDate && !drDate.value) drDate.value = new Date().toISOString().slice(0, 10);
  const miStart = document.getElementById("miStart");
  const miEnd = document.getElementById("miEnd");
  const drStart = document.getElementById("drStart");
  const drEnd = document.getElementById("drEnd");
  if (miStart && !miStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    miStart.value = d.toISOString().slice(0, 10);
  }
  if (miEnd && !miEnd.value) miEnd.value = new Date().toISOString().slice(0, 10);
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
  const wfActionDate = document.getElementById("wfActionDate");
  if (wfActionDate && !wfActionDate.value) wfActionDate.value = new Date().toISOString().slice(0, 10);
  document.getElementById("saveMiBtn")?.addEventListener("click", saveManagerInspection);
  document.getElementById("saveDrBtn")?.addEventListener("click", saveDamageReport);
  document.getElementById("loadMiBtn")?.addEventListener("click", loadManagerInspections);
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
  loadDamageReports().catch(() => {});
  loadWeeklyForumSummary().catch(() => {});
  loadWeeklyForumActions().catch(() => {});
  loadWeeklyForumParts().catch(() => {});
  loadWeeklyForumInputs().catch(() => {});
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
  document.getElementById("showWeeklyForumBtn")?.addEventListener("click", () => setTopView("wf"));
  document.getElementById("showSyncAdminBtn")?.addEventListener("click", () => setTopView("sync"));
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
});