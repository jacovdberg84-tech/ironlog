// IRONLOG/web/maintenance/plans.js — Maintenance plans, services due, service history and backfill.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

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

function isLdvAssetCode(code) {
  return /^V(0[1-9]|1[0-5])AM$/i.test(String(code || "").trim());
}

function syncPlanServiceTypeFromDropdown() {
  const typeEl = document.getElementById("planServiceType");
  const nameEl = document.getElementById("planServiceName");
  const intervalEl = document.getElementById("planIntervalHours");
  if (!typeEl || !nameEl || !intervalEl) return;
  const hours = Number(typeEl.value || 0);
  if (hours > 0) {
    nameEl.value = String(hours);
    intervalEl.value = String(hours);
  } else {
    nameEl.value = "";
    intervalEl.value = "0";
  }
}

function refreshPlanServiceTypeDropdown() {
  const assetEl = document.getElementById("planAsset");
  const typeEl = document.getElementById("planServiceType");
  if (!typeEl) return;
  const opt = assetEl?.selectedOptions?.[0];
  const code = String(opt?.textContent || "").split(" - ")[0]?.trim() || "";
  const ldv = isLdvAssetCode(code);
  typeEl.querySelectorAll("option[data-ldv-only]").forEach((o) => {
    o.hidden = !ldv;
    o.disabled = !ldv;
  });
  if (!ldv && typeEl.value === "10000") typeEl.value = "";
  syncPlanServiceTypeFromDropdown();
}

function groupPlansByAsset(plans) {
  const map = new Map();
  for (const p of plans || []) {
    const assetId = Number(p.asset_id || 0);
    if (!assetId) continue;
    if (!map.has(assetId)) {
      map.set(assetId, {
        asset_id: assetId,
        asset_code: p.asset_code,
        asset_name: p.asset_name,
        plans: [],
      });
    }
    map.get(assetId).plans.push(p);
  }
  for (const group of map.values()) {
    group.plans.sort((a, b) => Number(a.interval_hours || 0) - Number(b.interval_hours || 0));
    const nextPlan = group.plans.find((p) => p.is_next_for_asset) || group.plans[0];
    group.next = nextPlan || null;
    group.current_hours = nextPlan?.current_hours;
    group.last_service_hours = nextPlan?.last_service_hours_snapped ?? nextPlan?.last_service_hours;
    group.next_service_name = nextPlan?.next_service_name || nextPlan?.service_name;
    group.next_due_hours = nextPlan?.next_due_hours;
    group.remaining_hours = nextPlan?.remaining_hours;
    group.schedule_mode = nextPlan?.schedule_mode;
    group.near_due_threshold = nextPlan?.near_due_threshold ?? 50;
    group.meter_unit = nextPlan?.meter_unit || "hours";
    group.status = nextPlan?.status || null;
  }
  return Array.from(map.values()).sort((a, b) =>
    String(a.asset_code || "").localeCompare(String(b.asset_code || ""))
  );
}

function plansTableRow(group) {
  const assetId = Number(group.asset_id || 0);
  const nextPlan = group.next;
  const planId = Number(nextPlan?.id ?? nextPlan?.plan_id ?? 0);
  const remaining = Number(group.remaining_hours ?? 0);
  const near = Number(group.near_due_threshold ?? 50);
  const unit = String(group.meter_unit || "hours") === "km" ? "km" : "h";
  const intervals = [...new Set((group.plans || [])
    .map((p) => Number(p.interval_hours || 0))
    .filter((interval) => interval > 0))]
    .sort((a, b) => a - b);
  const cycle = intervals.length
    ? `${intervals.map((interval) => `${interval.toFixed(0)}${unit}`).join(" / ")}${intervals.length > 1 ? " alternating" : " interval"}`
    : "—";
  const remClass = remaining <= 0 ? "status-overdue" : remaining <= near ? "status-soon" : "";
  const statusPill = group.status === "OVERDUE"
    ? `<br><span class="maintenance-status is-overdue">Overdue</span>`
    : group.status === "ALMOST DUE"
      ? `<br><span class="maintenance-status is-soon">Almost due</span>`
      : "";
  return `
    <tr data-plan-asset-id="${assetId}" data-plan-id="${planId}">
      <td style="text-align:center;">
        ${planId > 0 ? `<input type="checkbox" class="due-plan-select" data-plan-id="${planId}" title="Select for work order / PDF" ${selectedDuePlanIds.has(planId) ? "checked" : ""} />` : ""}
      </td>
      <td><b>${escBackfill(group.asset_code || "-")}</b><br><small class="muted">${escBackfill(group.asset_name || "")}</small></td>
      <td style="text-align:right;">${fmt1(group.current_hours)}<br><small class="muted">${unit}</small></td>
      <td style="text-align:right;">${fmt1(group.last_service_hours)}<br><small class="muted">${unit}</small></td>
      <td><strong>${escBackfill(group.next_service_name || "-")}</strong>${statusPill}</td>
      <td style="text-align:right;">${fmt1(group.next_due_hours)}<br><small class="muted">${unit}</small></td>
      <td style="text-align:right;"><span class="${remClass}">${fmt1(group.remaining_hours)}</span><br><small class="muted">${unit}</small></td>
      <td><small class="muted">${cycle}</small></td>
      <td>
        ${planId > 0 ? `<button type="button" data-plan-rebase-id="${planId}" title="Use only after the service has been completed">Mark serviced</button>` : ""}
      </td>
    </tr>
  `;
}

function updateDueSelectionCount() {
  const el = document.getElementById("dueSelectionCount");
  if (el) el.textContent = `${selectedDuePlanIds.size} selected`;
}

function bindDuePlanSelectCheckboxes(root) {
  (root || document).querySelectorAll(".due-plan-select").forEach((el) => {
    if (el.dataset.boundDueSelect === "1") return;
    el.dataset.boundDueSelect = "1";
    el.addEventListener("change", () => {
      const pid = Number(el.getAttribute("data-plan-id") || 0);
      if (!pid) return;
      if (el.checked) selectedDuePlanIds.add(pid);
      else selectedDuePlanIds.delete(pid);
      updateDueSelectionCount();
    });
  });
  updateDueSelectionCount();
}

function clearDuePlanSelection() {
  selectedDuePlanIds.clear();
  document.querySelectorAll(".due-plan-select").forEach((el) => { el.checked = false; });
  updateDueSelectionCount();
}

function selectAllDuePlansInTable() {
  document.querySelectorAll("#plansList .due-plan-select").forEach((el) => {
    const pid = Number(el.getAttribute("data-plan-id") || 0);
    if (!pid) return;
    el.checked = true;
    selectedDuePlanIds.add(pid);
  });
  updateDueSelectionCount();
}

function planIntervalForMatch(p) {
  const iv = Number(p?.interval_hours || 0);
  if (iv > 0) return iv;
  const m = String(p?.service_name || "").match(/(\d+(?:\.\d+)?)/);
  return m ? Number(m[1]) : 0;
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

  const unit = String(d.meter_unit || "hours") === "km" ? "km" : "h";

  return `
    <div class="card">
      <label style="display:flex; align-items:center; gap:6px; margin-bottom:6px;">
        <input type="checkbox" class="due-plan-select" data-plan-id="${Number(d.plan_id || 0)}" ${selectedDuePlanIds.has(Number(d.plan_id || 0)) ? "checked" : ""} />
        <small>Select for work order</small>
      </label>
      <div><strong>${d.asset_code}</strong> - ${d.asset_name}</div>
      <div><strong>Next service:</strong> ${d.service_name}${String(d.schedule_mode || "") === "rotating" ? ` <span class="muted">(${Number(d.next_service_interval || d.interval_hours || 0).toFixed(0)}${unit})</span>` : ""}</div>
      <div><strong>Current:</strong> ${Number(d.current_hours || 0).toFixed(1)} ${unit}</div>
      <div><strong>Next Due:</strong> ${Number(d.next_due_hours || 0).toFixed(1)} ${unit}</div>
      <div><strong>Remaining:</strong> ${Number(d.remaining_hours || 0).toFixed(1)} ${unit}${Number(d.near_due_threshold || 0) > 0 ? ` <span class="muted">(almost due ≤ ${Number(d.near_due_threshold).toFixed(0)}${unit})</span>` : ""}</div>
      <div><strong>Est. service date:</strong> ${d.estimated_service_date || "—"}</div>
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
  const near = Number(r.near_due_threshold ?? 50);
  const unit = String(r.meter_unit || "hours") === "km" ? "km" : "h";
  const warn = hrsToNext <= 0 ? "status-overdue" : hrsToNext <= near ? "status-soon" : "status-ok";
  const sourceMap = {
    daily_closing: "Daily closing",
    asset_hours: "Asset hours",
    daily_sum: "Daily sum",
  };
  const src = sourceMap[String(r.current_hours_source || "").trim()] || "Unknown";
  const assetCode = String(r.asset_code || "").trim();
  const assetId = Number(r.asset_id || 0);
  const lastSvcHrs = r.last_service_hours != null ? fmt1(r.last_service_hours) : null;
  const lastSrc = String(r.last_service_source || "").trim();
  const srcLabel = lastSrc === "backfill" ? " <span class='maintenance-source'>Service record</span>" : lastSrc === "work_order" ? " <span class='maintenance-source'>Work order</span>" : "";
  const nextDue = r.next_due_hours != null ? `<br><small class="muted">Due @ ${fmt1(r.next_due_hours)}h</small>` : "";
  return `
    <tr class="${hrsToNext <= 0 ? "downRow" : ""}">
      <td><b>${escBackfill(eq)}</b>${lastSvcHrs != null ? `<br><small class="muted">Last svc ${lastSvcHrs}h</small>` : ""}</td>
      <td>${escBackfill(r.service_name || "-")}${nextDue}</td>
      <td>${escBackfill(last)}${srcLabel}</td>
      <td style="text-align:right;">${fmt1(r.current_hours)}<br><small class="muted">(${escBackfill(src)}, ${unit})</small></td>
      <td style="text-align:right;"><span class="${warn}">${fmt1(r.remaining_hours)}</span>${r.status === "ALMOST DUE" ? `<br><small class="maintenance-status is-soon">Almost due</small>` : ""}</td>
      <td style="text-align:right;">${fmt1(r.avg_daily_hours)}</td>
      <td>
        ${escBackfill(est)}
        ${assetCode ? `<div style="margin-top:6px;"><button type="button" data-hist-view-records="${assetCode.replace(/"/g, "")}" data-hist-asset-id="${assetId}">View records</button></div>` : ""}
      </td>
    </tr>
  `;
}

function filterHistoryRows(rows) {
  const code = String(document.getElementById("histAssetFilter")?.value || "").trim().toUpperCase();
  if (!code) return rows;
  return rows.filter((r) => String(r.asset_code || "").trim().toUpperCase() === code);
}

function openServiceRecordsForAsset(assetCode, assetId = 0) {
  const code = String(assetCode || "").trim().toUpperCase();
  window.__serviceRecordsPendingFilter = {
    code,
    assetId: Number(assetId || 0),
  };
  applyServiceRecordsAssetFilter();
  scrollToSection("service-history");
  const card = document.getElementById("serviceRecordsCard");
  if (card) {
    card.setAttribute("data-collapsed", "false");
    card.querySelector(".maintenance-card-body")?.removeAttribute("hidden");
  }
  setTimeout(() => loadBackfillHistory().catch(() => {}), 80);
}

function applyServiceRecordsAssetFilter() {
  const pending = window.__serviceRecordsPendingFilter;
  if (!pending) return;
  const listSel = document.getElementById("backfillListAsset");
  if (!listSel) return;
  if (pending.assetId > 0) {
    listSel.value = String(pending.assetId);
    return;
  }
  const code = String(pending.code || "").trim().toUpperCase();
  const opt = Array.from(listSel.options).find(
    (o) => String(o.getAttribute("data-asset-code") || "").toUpperCase() === code
  );
  if (opt) listSel.value = opt.value;
}

function hideBackfillEditPanel() {
  document.getElementById("backfillEditPanel")?.classList.add("hidden");
  const idEl = document.getElementById("backfillEditId");
  if (idEl) idEl.value = "";
}

function showBackfillEditPanel(row) {
  const panel = document.getElementById("backfillEditPanel");
  if (!panel || !row) return;
  const idEl = document.getElementById("backfillEditId");
  if (idEl) idEl.value = String(row.id || "");
  const nameEl = document.getElementById("backfillEditServiceName");
  if (nameEl) nameEl.value = String(row.service_name || "");
  const dateEl = document.getElementById("backfillEditServiceDate");
  if (dateEl) dateEl.value = String(row.service_date || "");
  const hoursEl = document.getElementById("backfillEditServiceHours");
  if (hoursEl) hoursEl.value = row.service_hours == null ? "" : String(row.service_hours);
  const notesEl = document.getElementById("backfillEditNotes");
  if (notesEl) notesEl.value = String(row.notes || "");
  panel.classList.remove("hidden");
  panel.scrollIntoView({ behavior: "smooth", block: "nearest" });
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
  const source = String(r.record_source || "backfill");
  const isWo = source === "work_order";
  const srcBadge = isWo ? ` <span class="maintenance-source">Work order #${Number(r.id || 0)}</span>` : "";
  const actions = isWo
    ? `<span class="muted">Closed work order</span>`
    : `<button data-backfill-action="edit">Edit</button>
        <button data-backfill-action="delete">Delete</button>`;
  return `
    <tr data-backfill-id="${Number(r.id || 0)}" data-record-source="${escBackfill(source)}">
      <td>${escBackfill(r.asset_code || "-")} - ${escBackfill(r.asset_name || "-")}</td>
      <td>${escBackfill(r.service_name || "-")}${srcBadge}</td>
      <td>${escBackfill(r.service_date || "-")}</td>
      <td style="text-align:right;">${hours}</td>
      <td>${escBackfill(r.notes || "")}</td>
      <td>${actions}</td>
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
    let rows = filterHistoryRows(rawRows.slice());
    if (mode === "closest") {
      rows.sort((a, b) => Number(a?.remaining_hours || 0) - Number(b?.remaining_hours || 0));
      const n = Number(limitEl?.value || 20);
      const topN = Number.isFinite(n) ? Math.max(1, Math.min(200, Math.trunc(n))) : 20;
      rows = rows.slice(0, topN);
    }
    if (meta) {
      const filt = String(document.getElementById("histAssetFilter")?.value || "").trim();
      const suffix = mode === "closest" ? ` | Closest due (${rows.length})` : filt ? ` | Filtered: ${filt}` : "";
      meta.textContent = `As of: ${data.as_of || "-"} | ${rawRows.length} plan(s)${suffix}`;
    }
    body.innerHTML = rows.length ? rows.map(histRow).join("") : `<tr><td colspan="7" class="muted">No history rows.</td></tr>`;
  } catch (e) {
    body.innerHTML = `<tr><td colspan="7" class="message-error">History load error: ${e.message || e}</td></tr>`;
  }
}


async function loadBackfillHistory() {
  const body = document.getElementById("backfillBody");
  const meta = document.getElementById("backfillListMeta");
  if (!body) return;
  applyServiceRecordsAssetFilter();
  body.innerHTML = `<tr><td colspan="6" class="muted">Loading...</td></tr>`;
  try {
    const assetId = Number(document.getElementById("backfillListAsset")?.value || 0);
    const limit = 200;
    const q = new URLSearchParams({ limit: String(limit), include_work_orders: "1" });
    if (assetId > 0) q.set("asset_id", String(assetId));
    const res = await fetch(`${API}/maintenance/history/backfill?${q.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load service records");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    window.__backfillRowsById = new Map(rows.map((r) => [Number(r.id), r]));
    body.innerHTML = rows.length
      ? rows.map(backfillRow).join("")
      : `<tr><td colspan="6" class="muted">No service records found for this filter.</td></tr>`;
    if (meta) {
      meta.textContent = `${rows.length} record(s)${assetId > 0 ? " for selected asset" : ""}`;
    }
  } catch (e) {
    body.innerHTML = `<tr><td colspan="6" class="message-error">${escBackfill(e.message || e)}</td></tr>`;
    if (meta) meta.textContent = "Load failed";
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

function populateMaintAssetSelects(validAssets) {
  const optionsHtml = `
    <option value="">Select asset</option>
    ${validAssets.map(a => `
      <option value="${a.id}">
        ${a.asset_code || "NO-CODE"} - ${a.asset_name || "Unnamed Asset"}
      </option>
    `).join("")}
  `;
  const filterOptionsHtml = `
    <option value="">All assets</option>
    ${validAssets.map(a => `
      <option value="${a.id}" data-asset-code="${escBackfill(a.asset_code || "")}">
        ${a.asset_code || "NO-CODE"} - ${a.asset_name || "Unnamed Asset"}
      </option>
    `).join("")}
  `;
  const histFilterOptionsHtml = `
    <option value="">All equipment</option>
    ${validAssets.map(a => `
      <option value="${escBackfill(a.asset_code || "")}">
        ${a.asset_code || "NO-CODE"} - ${a.asset_name || "Unnamed Asset"}
      </option>
    `).join("")}
  `;
  const select = document.getElementById("planAsset");
  const backfillSelect = document.getElementById("backfillAsset");
  const backfillListSelect = document.getElementById("backfillListAsset");
  const histAssetFilter = document.getElementById("histAssetFilter");
  if (select) select.innerHTML = optionsHtml;
  if (backfillSelect) backfillSelect.innerHTML = optionsHtml;
  if (backfillListSelect) backfillListSelect.innerHTML = filterOptionsHtml;
  if (histAssetFilter) histAssetFilter.innerHTML = histFilterOptionsHtml;
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

    populateMaintAssetSelects(validAssets);
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

function startEditBackfillHistory(id) {
  const iid = Number(id || 0);
  if (!iid) return;
  const row = window.__backfillRowsById?.get(iid);
  if (!row) return;
  showBackfillEditPanel(row);
}

async function saveBackfillEditPanel() {
  const iid = Number(document.getElementById("backfillEditId")?.value || 0);
  if (!iid) return;
  const payload = {
    service_name: String(document.getElementById("backfillEditServiceName")?.value || "").trim(),
    service_date: String(document.getElementById("backfillEditServiceDate")?.value || "").trim(),
    service_hours: (() => {
      const raw = String(document.getElementById("backfillEditServiceHours")?.value || "").trim();
      return raw === "" ? null : Number(raw);
    })(),
    notes: String(document.getElementById("backfillEditNotes")?.value || "").trim(),
  };
  if (!payload.service_name || !payload.service_date) {
    alert("Service name and date are required.");
    return;
  }
  const res = await fetch(`${API}/maintenance/history/backfill/${iid}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to update service record");
  hideBackfillEditPanel();
  await loadBackfillHistory();
  await loadHistory();
  await loadPlans();
  await loadDue();
}

async function deleteBackfillHistory(id) {
  const iid = Number(id || 0);
  if (!iid) return;
  if (!confirm("Delete this service record? This cannot be undone.")) return;
  const res = await fetch(`${API}/maintenance/history/backfill/${iid}`, { method: "DELETE" });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to delete historical entry");
  if (Number(document.getElementById("backfillEditId")?.value || 0) === iid) hideBackfillEditPanel();
  await loadBackfillHistory();
  await loadHistory();
  await loadPlans();
  await loadDue();
}

async function loadLiveHoursForSelectedAsset() {
  const assetEl = document.getElementById("planAsset");
  const currentHoursEl = document.getElementById("planCurrentHours");
  const currentHoursSrcEl = document.getElementById("planCurrentHoursSource");
  const nextPreviewEl = document.getElementById("planNextServicePreview");

  if (!assetEl || !currentHoursEl) return;

  const assetId = Number(assetEl.value || 0);

  if (!assetId) {
    currentHoursEl.value = "0";
    if (currentHoursSrcEl) currentHoursSrcEl.textContent = "Source: -";
    if (nextPreviewEl) nextPreviewEl.textContent = "";
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
    if (nextPreviewEl) {
      const ns = data.next_service;
      if (ns && data.rotating_schedule) {
        nextPreviewEl.textContent = `Next service: ${ns.service_name} @ ${Number(ns.next_due_hours || 0).toFixed(0)}h (${Number(ns.remaining_hours || 0).toFixed(0)}h remaining)`;
      } else if (ns) {
        nextPreviewEl.textContent = `Next due: ${ns.service_name} @ ${Number(ns.next_due_hours || 0).toFixed(0)}h`;
      } else {
        nextPreviewEl.textContent = "No active plans yet for this asset.";
      }
    }
    syncLastServiceHoursFromLive();
  } catch (err) {
    console.error("Live hours load error:", err);
    // Don’t overwrite user input with 0 when live hours fails
    currentHoursEl.value = "";
    currentHoursEl.placeholder = "";
    if (currentHoursSrcEl) currentHoursSrcEl.textContent = "Source: -";
    if (nextPreviewEl) nextPreviewEl.textContent = "";
  }
}

async function savePlan() {
  const assetEl = document.getElementById("planAsset");
  const typeEl = document.getElementById("planServiceType");
  const serviceNameEl = document.getElementById("planServiceName");
  const intervalEl = document.getElementById("planIntervalHours");
  const lastServiceEl = document.getElementById("planLastServiceHours");
  const activeEl = document.getElementById("planActive");
  const msgEl = document.getElementById("planFormMessage");

  syncPlanServiceTypeFromDropdown();

  if (!assetEl || !typeEl || !serviceNameEl || !intervalEl || !lastServiceEl || !activeEl || !msgEl) {
    console.error("Plan form elements missing");
    return;
  }

  const payload = {
    asset_id: Number(assetEl.value || 0),
    service_name: serviceNameEl.value.trim(),
    interval_hours: Number(intervalEl.value || 0),
    last_service_hours: Number(lastServiceEl.value || 0),
    active: Number(activeEl.value || 1),
  };

  if (!payload.asset_id || !payload.service_name || payload.interval_hours <= 0) {
    msgEl.className = "message-error";
    msgEl.textContent = "Please select an asset and a service type from the dropdown.";
    return;
  }

  const existing = (__maintenancePlansCache || []).find(
    (p) =>
      Number(p.asset_id) === payload.asset_id &&
      planIntervalForMatch(p) === payload.interval_hours,
  );

  msgEl.className = "";
  msgEl.textContent = existing ? "Updating existing schedule…" : "Adding service type…";

  try {
    const res = await fetch(
      existing
        ? `${API}/maintenance/plans/${Number(existing.id ?? existing.plan_id ?? 0)}`
        : `${API}/maintenance/plans`,
      {
        method: existing ? "PUT" : "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      },
    );

    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || "Failed to save maintenance plan");
    }

    msgEl.className = "message-success";
    msgEl.textContent = existing
      ? "Service schedule updated."
      : "Service type added to asset.";

    typeEl.value = "";
    syncPlanServiceTypeFromDropdown();
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

async function deleteMaintenancePlan(planId) {
  const id = Number(planId || 0);
  if (!id) return;
  if (!confirm("Remove this service type from the asset? (Use this to delete duplicate entries.)")) return;
  const res = await fetch(`${API}/maintenance/plans/${id}`, { method: "DELETE" });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to delete plan");
  await loadPlans();
  await loadDue();
  await loadHistory();
}

let __maintenancePlansCache = [];
let __maintenancePlanGroupsCache = [];
let maintenancePlanFilter = "all";

function maintenancePlanGroupBucket(group) {
  const remaining = Number(group?.remaining_hours ?? 0);
  const near = Number(group?.near_due_threshold ?? 50);
  if (remaining <= 0) return "overdue";
  if (remaining <= near) return "due";
  return "upcoming";
}

async function openProtectedPdf(url, { download = false, filename = "IRONLOG-report.pdf" } = {}) {
  const preview = download ? null : window.open("", "_blank");
  if (preview) preview.opener = null;
  try {
    const res = await fetch(url, { headers: authHeaders(), cache: "no-store" });
    if (!res.ok) {
      const body = await res.json().catch(() => ({}));
      throw new Error(res.status === 401
        ? "Your session has expired. Please sign in again."
        : body.error || body.message || `PDF request failed (${res.status})`);
    }
    const blob = await res.blob();
    if (!String(blob.type).toLowerCase().includes("application/pdf")) {
      throw new Error("The server did not return a PDF.");
    }
    const blobUrl = URL.createObjectURL(blob);
    try {
      if (download) {
        const link = document.createElement("a");
        link.href = blobUrl;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        link.remove();
      } else if (preview && !preview.closed) {
        preview.location.replace(blobUrl);
      } else if (typeof window.location?.assign === "function") {
        // Embedded browsers may block controlled popups. Keep the PDF accessible
        // by opening the authenticated blob in the current tab instead.
        window.location.assign(blobUrl);
      } else {
        const link = document.createElement("a");
        link.href = blobUrl;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        link.remove();
      }
    } finally {
      setTimeout(() => URL.revokeObjectURL(blobUrl), 120000);
    }
  } catch (err) {
    if (preview && !preview.closed) preview.close();
    alert(`Could not ${download ? "download" : "open"} PDF: ${err.message || err}`);
  }
}

function renderMaintenancePlanningKpis(groups) {
  const list = Array.isArray(groups) ? groups : [];
  const counts = {
    all: list.length,
    overdue: list.filter((g) => maintenancePlanGroupBucket(g) === "overdue").length,
    due: list.filter((g) => maintenancePlanGroupBucket(g) === "due").length,
    upcoming: list.filter((g) => maintenancePlanGroupBucket(g) === "upcoming").length,
  };
  document.querySelectorAll("#maintenancePlanningKpis [data-maint-plan-filter]").forEach((button) => {
    const key = String(button.getAttribute("data-maint-plan-filter") || "all");
    const value = button.querySelector("strong");
    if (value) value.textContent = String(counts[key] ?? 0);
    button.classList.toggle("is-active", key === maintenancePlanFilter);
  });
}

function renderMaintenancePlanQueue() {
  const container = document.getElementById("plansList");
  if (!container) return;
  const search = String(document.getElementById("maintenancePlanSearch")?.value || "").trim().toLowerCase();
  const groups = __maintenancePlanGroupsCache.filter((group) => {
    if (maintenancePlanFilter !== "all" && maintenancePlanGroupBucket(group) !== maintenancePlanFilter) return false;
    if (!search) return true;
    const services = (group.plans || []).map((p) => p.service_name || p.interval_hours || "").join(" ");
    return `${group.asset_code || ""} ${group.asset_name || ""} ${services}`.toLowerCase().includes(search);
  });
  container.innerHTML = groups.length
    ? groups.map(plansTableRow).join("")
    : `<tr><td colspan="9" class="muted">No service schedules match this view.</td></tr>`;
  bindDuePlanSelectCheckboxes(container);
  const meta = document.getElementById("maintenancePlanQueueMeta");
  if (meta) meta.textContent = `Showing ${groups.length} of ${__maintenancePlanGroupsCache.length} planned asset(s).`;
  document.querySelectorAll(".maintenance-filter-buttons [data-maint-plan-filter]").forEach((button) => {
    button.classList.toggle("is-active", button.getAttribute("data-maint-plan-filter") === maintenancePlanFilter);
  });
  renderMaintenancePlanningKpis(__maintenancePlanGroupsCache);
}

function setMaintenancePlanFilter(filter) {
  const next = ["all", "overdue", "due", "upcoming"].includes(String(filter)) ? String(filter) : "all";
  maintenancePlanFilter = next;
  renderMaintenancePlanQueue();
}

async function loadPlans() {
  const container = document.getElementById("plansList");
  if (!container) return;
  container.innerHTML = `<tr><td colspan="9" class="muted">Loading...</td></tr>`;

  try {
    const res = await fetch(`${API}/maintenance/plans`);
    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || "Failed to load plans");
    }

    const plans = (Array.isArray(data.plans) ? data.plans : []).map((p) => ({
      ...p,
      id: Number(p?.id ?? p?.plan_id ?? 0) || 0,
    }));
    __maintenancePlansCache = plans;
    const groups = groupPlansByAsset(plans);
    __maintenancePlanGroupsCache = groups;
    renderMaintenancePlanQueue();
  } catch (err) {
    console.error("Plans error:", err);
    container.innerHTML = `<tr><td colspan="9" class="message-error">Error loading plans: ${escBackfill(err.message)}</td></tr>`;
  }
}

async function rebasePlanLastServiceHours(id, fallback = {}) {
  const planId = Number(id || 0);
  const url = planId
    ? `${API}/maintenance/plans/${planId}/rebase-last-service`
    : `${API}/maintenance/plans/0/rebase-last-service`;
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      plan_id: planId || null,
      asset_id: Number(fallback.asset_id || 0) || null,
      service_name: String(fallback.service_name || "").trim() || null
    })
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to rebase last service hours");
  await loadPlans();
  await loadDue();
  await loadHistory();
  alert(
    `Rebased ${data.asset_code || ""} ${data.service_name || ""} to ${Number(data.last_service_hours || 0).toFixed(1)} hours (${data.source || "source unknown"}).`
  );
}

async function loadDue() {
  const container = document.getElementById("dueList");
  if (!container) return;
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
    const validPlanIds = new Set(due.map((x) => Number(x.plan_id || 0)).filter((x) => x > 0));
    Array.from(selectedDuePlanIds).forEach((id) => {
      if (!validPlanIds.has(id)) selectedDuePlanIds.delete(id);
    });
    due.sort((a, b) => {
      const getRank = (d) => {
        const remaining = Number(d.remaining_hours || 0);
        const near = Number(d.near_due_threshold || nearDueHours);
        if (remaining <= 0) return 1;
        if (remaining <= near) return 2;
        return 3;
      };

      const rankDiff = getRank(a) - getRank(b);
      if (rankDiff !== 0) return rankDiff;
      return Number(a.remaining_hours || 0) - Number(b.remaining_hours || 0);
    });
    container.innerHTML = due.length
      ? due.map(dueCard).join("")
      : "<div>No due services found.</div>";
    bindDuePlanSelectCheckboxes(container);
  } catch (err) {
    console.error("Due error:", err);
    container.innerHTML = `<div style="color:#ff8080;">Error loading due services: ${err.message}</div>`;
  }
}

async function openUpcomingServicesPdf(download = false, includeAll = false) {
  const nearDueHours = getDueThresholdHours();
  const q = new URLSearchParams();
  q.set("near_due_hours", String(nearDueHours));
  const selectedPlans = Array.from(selectedDuePlanIds);
  if (!includeAll && selectedPlans.length) q.set("plan_ids", selectedPlans.join(","));
  else if (!includeAll) q.set("within_hours", String(nearDueHours));
  q.set("_", String(Date.now()));
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

async function downloadProtectedXlsxFile(url, filename) {
  const res = await fetch(url, { headers: authHeaders(), cache: "no-store" });
  if (!res.ok) {
    let message = await res.text().catch(() => "");
    try {
      const body = JSON.parse(message);
      message = body.error || body.message || message;
    } catch {}
    throw new Error(res.status === 401
      ? "Your session has expired. Please sign in again."
      : message || `Export failed (${res.status})`);
  }
  const blob = await res.blob();
  const blobUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = blobUrl;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(blobUrl), 120000);
}

async function downloadMaintenancePlansXlsx() {
  const nearDueHours = getDueThresholdHours();
  const dateTag = new Date().toISOString().slice(0, 10);
  try {
    await downloadProtectedXlsxFile(
      `${API}/maintenance/plans.xlsx?near_due_hours=${encodeURIComponent(String(nearDueHours))}`,
      `IRONLOG_Maintenance_Plans_${dateTag}.xlsx`,
    );
  } catch (err) {
    alert(`Could not download maintenance plans: ${err.message || err}`);
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

async function openMaintenancePackLatest(reportType, download = false) {
  const type = String(reportType || "").toLowerCase() === "monthly" ? "monthly" : "weekly";
  const msg = document.getElementById("mpStatusMsg");
  const q = new URLSearchParams();
  q.set("period_type", type);
  const sel = getMpSelectedRange(type);
  if (type === "monthly") q.set("month", sel.month);
  else {
    q.set("start", sel.start);
    q.set("end", sel.end);
  }
  if (download) q.set("download", "1");
  const popup = download ? null : window.open("", "_blank");
  const url = `${API}/reports/maintenance-master/latest.pptx?${q.toString()}`;
  try {
    if (msg) {
      msg.className = "muted";
      msg.textContent = `${download ? "Downloading" : "Opening"} ${type} presentation...`;
    }
    const res = await fetch(url, { headers: authHeaders() });
    const blob = await res.blob();
    if (!res.ok) {
      let detail = await blob.text().catch(() => "");
      try {
        const parsed = JSON.parse(detail);
        detail = parsed.error || parsed.message || detail;
      } catch { /* use response text */ }
      throw new Error(detail || `Presentation request failed (${res.status})`);
    }
    const blobUrl = URL.createObjectURL(blob);
    if (download) {
      const disposition = String(res.headers.get("content-disposition") || "");
      const match = disposition.match(/filename="?([^";]+)"?/i);
      const anchor = document.createElement("a");
      anchor.href = blobUrl;
      anchor.download = match?.[1] || `IRONLOG_Maintenance_Master_${type}.pptx`;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      setTimeout(() => URL.revokeObjectURL(blobUrl), 120000);
    } else if (popup) {
      popup.location.href = blobUrl;
      setTimeout(() => URL.revokeObjectURL(blobUrl), 120000);
    } else {
      window.open(blobUrl, "_blank");
      setTimeout(() => URL.revokeObjectURL(blobUrl), 120000);
    }
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `${type} presentation ${download ? "downloaded" : "opened"}.`;
    }
  } catch (e) {
    if (popup && !popup.closed) popup.close();
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Presentation error: ${e.message || e}`;
    }
  }
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

    const nearDueHours = getDueThresholdHours();
    const selectedPlans = Array.from(selectedDuePlanIds);
    const res = await fetch(`${API}/maintenance/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        near_due_hours: nearDueHours,
        plan_ids: selectedPlans
      })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to generate work orders");
    const mode = selectedPlans.length ? "selected plans" : "overdue plans";
    const skipped = Array.isArray(data.skipped) ? data.skipped : [];
    const skippedMsg = skipped.length
      ? `\n\nAlready open: ${skipped.map((r) => `${r.asset_code || "asset"} — WO #${r.work_order_id}`).join(", ")}`
      : "";
    alert(`Created ${Number(data.created_count || 0)} work orders (${mode})${skippedMsg}`);
    selectedDuePlanIds.clear();
    await loadPlans();
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
