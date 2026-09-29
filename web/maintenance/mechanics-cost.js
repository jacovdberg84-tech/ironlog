// IRONLOG/web/maintenance/mechanics-cost.js — Mechanics cost and timesheets.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

let mcTechnicianOptions = [];
let mcDayRowsCache = [];
let mcDraftRows = [];
let mcWorkOrderOptions = [];

const MC_CATEGORIES = ["Startup", "Breakdown", "Service", "Maintenance", "Inspection", "Other"];

function mcEmptyDraftRow(technician_name = "") {
  return {
    technician_name: String(technician_name || "").trim(),
    hours: "",
    asset_code: "",
    category: "",
    reason: "",
    time_started: "",
    time_finished: "",
    job_card_no: "",
    smr: "",
  };
}

function mcTechnicianLabel(t) {
  return String(t?.label || t?.username || t || "").trim();
}

function mcTechnicianValue(t) {
  return String(t?.username || t?.label || t || "").trim();
}

function mcTechnicianSelectOptions(selected = "") {
  const picked = String(selected || "").trim().toLowerCase();
  const users = Array.isArray(mcTechnicianOptions) ? mcTechnicianOptions : [];
  const merged = [...users];
  if (
    picked &&
    !merged.some((t) => {
      const v = mcTechnicianValue(t).toLowerCase();
      const l = mcTechnicianLabel(t).toLowerCase();
      return v === picked || l === picked;
    })
  ) {
    merged.unshift({ username: selected, label: selected });
  }
  return (
    `<option value="">Select…</option>` +
    merged
      .map((t) => {
        const value = mcTechnicianValue(t);
        const label = mcTechnicianLabel(t) || value;
        const sel =
          picked && (picked === value.toLowerCase() || picked === label.toLowerCase()) ? " selected" : "";
        return `<option value="${esc(value)}"${sel}>${esc(label)}</option>`;
      })
      .join("")
  );
}

function mcCategorySelectOptions(selected = "") {
  const current = String(selected || "").trim();
  const known = [...MC_CATEGORIES];
  if (current && !known.some((category) => category.toLowerCase() === current.toLowerCase())) {
    known.unshift(current);
  }
  return [
    `<option value="">Select…</option>`,
    ...known.map((category) => `<option value="${esc(category)}"${category.toLowerCase() === current.toLowerCase() ? " selected" : ""}>${esc(category)}</option>`),
  ].join("");
}

function mcDisplayDate(ymd) {
  const value = String(ymd || "").trim();
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return value || "—";
  const date = new Date(`${value}T12:00:00`);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat("en-GB", { day: "2-digit", month: "short", year: "numeric" }).format(date);
}

function mcHoursFromTimes(start, finish) {
  const parse = (value) => {
    const match = String(value || "").trim().match(/^(\d{1,2}):(\d{2})$/);
    if (!match) return null;
    const hours = Number(match[1]);
    const minutes = Number(match[2]);
    if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) return null;
    return hours * 60 + minutes;
  };
  const startMinutes = parse(start);
  const finishMinutes = parse(finish);
  if (startMinutes == null || finishMinutes == null || finishMinutes <= startMinutes) return null;
  return Number(((finishMinutes - startMinutes) / 60).toFixed(2));
}

function renderMcTechChips() {
  const el = document.getElementById("mcTechChips");
  const dl = document.getElementById("mcTechList");
  if (!el) return;
  const users = Array.isArray(mcTechnicianOptions) ? mcTechnicianOptions : [];
  if (dl) {
    dl.innerHTML = users
      .map((t) => `<option value="${esc(mcTechnicianValue(t))}">${esc(mcTechnicianLabel(t))}</option>`)
      .join("");
  }
  if (!users.length) {
    el.innerHTML = `<span class="muted small">No roster loaded — type the mechanic's name directly in the sheet.</span>`;
    return;
  }
  el.innerHTML = users
    .map((t) => {
      const label = mcTechnicianLabel(t);
      const value = mcTechnicianValue(t);
      return `<button type="button" class="mc-tech-chip" data-mc-chip-tech="${esc(value)}" title="Add row for ${esc(label)}">+ ${esc(label)}</button>`;
    })
    .join("");
}

function refreshMcDraftEditor() {
  const body = document.getElementById("mcDraftBody");
  if (!body) return;
  if (!mcDraftRows.length) {
    body.innerHTML = `<tr><td colspan="11" class="muted">Click <strong>+ Row per technician</strong> or a technician chip above to start.</td></tr>`;
    return;
  }
  const workDate = mcDisplayDate(document.getElementById("mcWorkDate")?.value);
  body.innerHTML = mcDraftRows
    .map(
      (row, idx) => `
    <tr data-mc-draft-idx="${idx}">
      <td>
        <span class="muted small">${esc(workDate)}</span>
      </td>
      <td>
        <input data-mc-draft-field="asset_code" data-mc-draft-idx="${idx}" type="text" list="mcAssetList"
          value="${esc(row.asset_code || "")}" placeholder="A300AM" style="width:110px;" autocomplete="off" />
      </td>
      <td>
        <input data-mc-draft-field="hours" data-mc-draft-idx="${idx}" type="number" min="0" step="0.25"
          value="${esc(row.hours !== "" && row.hours != null ? String(row.hours) : "")}" placeholder="0.00" style="width:88px; text-align:right;" />
      </td>
      <td>
        <select data-mc-draft-field="category" data-mc-draft-idx="${idx}" style="min-width:125px;">
          ${mcCategorySelectOptions(row.category)}
        </select>
      </td>
      <td style="min-width:260px;">
        <input data-mc-draft-field="reason" data-mc-draft-idx="${idx}" type="text"
          value="${esc(row.reason || "")}" placeholder="Description of work carried out" style="width:100%; min-width:240px;" />
      </td>
      <td>
        <input data-mc-draft-field="time_started" data-mc-draft-idx="${idx}" type="time"
          value="${esc(row.time_started || "")}" style="width:112px;" />
      </td>
      <td>
        <input data-mc-draft-field="time_finished" data-mc-draft-idx="${idx}" type="time"
          value="${esc(row.time_finished || "")}" style="width:112px;" />
      </td>
      <td>
        ${mcTechnicianOptions.length
          ? `<select data-mc-draft-field="technician_name" data-mc-draft-idx="${idx}" style="min-width:145px;">
              ${mcTechnicianSelectOptions(row.technician_name)}
            </select>`
          : `<input data-mc-draft-field="technician_name" data-mc-draft-idx="${idx}" type="text" list="mcTechList"
              value="${esc(row.technician_name || "")}" placeholder="Technician" style="min-width:145px;" autocomplete="off" />`}
      </td>
      <td>
        <input data-mc-draft-field="job_card_no" data-mc-draft-idx="${idx}" type="text" list="mcJobCardList"
          value="${esc(row.job_card_no || "")}" placeholder="WO #123" style="width:120px;" autocomplete="off" />
      </td>
      <td>
        <input data-mc-draft-field="smr" data-mc-draft-idx="${idx}" type="number" min="0" step="0.1"
          value="${esc(row.smr !== "" && row.smr != null ? String(row.smr) : "")}" placeholder="0.0" style="width:96px; text-align:right;" />
      </td>
      <td style="text-align:right;">
        <button type="button" data-mc-draft-del="${idx}">Remove</button>
      </td>
    </tr>`,
    )
    .join("");
}

function syncMcDraftFromDom() {
  document.querySelectorAll("#mcDraftBody tr[data-mc-draft-idx]").forEach((tr) => {
    const idx = Number(tr.getAttribute("data-mc-draft-idx") || -1);
    if (idx < 0 || idx >= mcDraftRows.length) return;
    tr.querySelectorAll("[data-mc-draft-field]").forEach((el) => {
      const field = String(el.getAttribute("data-mc-draft-field") || "");
      if (!field) return;
      mcDraftRows[idx][field] = String(el.value || "").trim();
    });
  });
}

function addMcDraftRow(technician_name = "") {
  mcDraftRows.push(mcEmptyDraftRow(technician_name));
  refreshMcDraftEditor();
}

function addMcDraftRowPerTechnician() {
  const users = Array.isArray(mcTechnicianOptions) ? mcTechnicianOptions : [];
  if (!users.length) {
    addMcDraftRow();
    return;
  }
  users.forEach((t) => mcDraftRows.push(mcEmptyDraftRow(mcTechnicianValue(t))));
  refreshMcDraftEditor();
}

function addMcDraftBlankRows(count = 5) {
  const n = Math.max(1, Math.min(50, Number(count) || 5));
  for (let i = 0; i < n; i += 1) mcDraftRows.push(mcEmptyDraftRow());
  refreshMcDraftEditor();
}

function clearMcDraft() {
  mcDraftRows = [];
  refreshMcDraftEditor();
  const msg = document.getElementById("mcFormMsg");
  if (msg) {
    msg.className = "muted";
    msg.textContent = "";
  }
}

function loadMcSavedIntoDraft() {
  if (!mcDayRowsCache.length) {
    const msg = document.getElementById("mcFormMsg");
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "No saved entries for this date to load.";
    }
    return;
  }
  mcDraftRows = mcDayRowsCache.map((r) => ({
    technician_name: String(r.technician_name || ""),
    hours: r.hours != null ? String(r.hours) : "",
    asset_code: String(r.asset_code || ""),
    category: String(r.category || ""),
    reason: String(r.reason || ""),
    time_started: String(r.time_started || ""),
    time_finished: String(r.time_finished || ""),
    job_card_no: String(r.job_card_no || ""),
    smr: r.smr != null ? String(r.smr) : "",
  }));
  refreshMcDraftEditor();
  const msg = document.getElementById("mcFormMsg");
  if (msg) {
    msg.className = "muted";
    msg.textContent = `Loaded ${mcDraftRows.length} saved row(s) into the sheet. Edit and use Save all with Replace checked to overwrite the day.`;
  }
}

function readMcDraftEntriesForSave() {
  syncMcDraftFromDom();
  const defaultRate = Number(document.getElementById("mcDefaultRate")?.value || 0);
  const entries = [];
  const errors = [];
  mcDraftRows.forEach((row, idx) => {
    const technician_name = String(row.technician_name || "").trim();
    const hours = Math.max(0, Number(row.hours || 0));
    const asset_code = String(row.asset_code || "").trim().toUpperCase();
    const category = String(row.category || "").trim();
    const reason = String(row.reason || "").trim();
    const time_started = String(row.time_started || "").trim();
    const time_finished = String(row.time_finished || "").trim();
    const job_card_no = String(row.job_card_no || "").trim();
    const smrRaw = String(row.smr ?? "").trim();
    const smr = smrRaw === "" ? null : Number(smrRaw);
    const empty = !technician_name && !asset_code && !category && !reason && !hours && !time_started && !time_finished && !job_card_no && smrRaw === "";
    if (empty) return;
    const line = idx + 1;
    if (!technician_name) errors.push(`Row ${line}: technician required`);
    if (!Number.isFinite(hours) || hours <= 0) errors.push(`Row ${line}: hours must be > 0`);
    if (!asset_code) errors.push(`Row ${line}: plant / asset required`);
    if (!category) errors.push(`Row ${line}: category required`);
    if (!reason) errors.push(`Row ${line}: reason required`);
    if (smrRaw !== "" && (!Number.isFinite(smr) || smr < 0)) errors.push(`Row ${line}: SMR must be 0 or more`);
    if (!technician_name || !Number.isFinite(hours) || hours <= 0 || !asset_code || !category || !reason || (smrRaw !== "" && (!Number.isFinite(smr) || smr < 0))) return;
    const entry = {
      technician_name,
      hours,
      asset_code,
      category,
      reason,
      time_started: time_started || null,
      time_finished: time_finished || null,
      job_card_no: job_card_no || null,
      smr: smrRaw === "" ? null : smr,
    };
    if (defaultRate > 0) entry.labor_rate_per_hour = defaultRate;
    entries.push(entry);
  });
  return { entries, errors };
}

function mcYearValue() {
  const y = Number(document.getElementById("mcReportYear")?.value || 0);
  return Number.isFinite(y) && y >= 2020 ? y : new Date().getFullYear();
}

function syncMcDateBounds() {
  const year = mcYearValue();
  const dateEl = document.getElementById("mcWorkDate");
  const yearEl = document.getElementById("mcReportYear");
  if (yearEl && !yearEl.value) yearEl.value = String(year);
  if (!dateEl) return;
  dateEl.min = `${year}-01-01`;
  dateEl.max = `${year}-12-31`;
  const cur = String(dateEl.value || "").trim();
  if (!cur || cur < dateEl.min || cur > dateEl.max) {
    const today = new Date().toISOString().slice(0, 10);
    dateEl.value = today >= dateEl.min && today <= dateEl.max ? today : dateEl.min;
  }
}

async function loadMcTechnicians() {
  try {
    const res = await fetch(`${API}/workorders/technicians`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load technicians");
    if (Array.isArray(data?.technician_users) && data.technician_users.length) {
      mcTechnicianOptions = data.technician_users;
    } else {
      mcTechnicianOptions = (Array.isArray(data?.technicians) ? data.technicians : []).map((name) => ({
        username: name,
        label: name,
      }));
    }
  } catch {
    mcTechnicianOptions = [];
  }
  renderMcTechChips();
  refreshMcDraftEditor();
}

async function loadMcAssetDatalist() {
  const dl = document.getElementById("mcAssetList");
  if (!dl) return;
  try {
    const res = await fetch(`${API}/assets?include_archived=0`, { headers: authHeaders() });
    const rows = await res.json();
    if (!res.ok) throw new Error(rows.error || "Failed to load assets");
    const arr = Array.isArray(rows) ? rows : [];
    dl.innerHTML = arr
      .map((a) => `<option value="${esc(String(a.asset_code || ""))}">${esc(String(a.asset_name || ""))}</option>`)
      .join("");
  } catch {
    dl.innerHTML = "";
  }
}

async function loadMcJobCardDatalist() {
  const dl = document.getElementById("mcJobCardList");
  if (!dl) return;
  try {
    const year = mcYearValue();
    const params = new URLSearchParams({ from_date: `${year}-01-01`, to_date: `${year}-12-31` });
    const res = await fetch(`${API}/workorders?${params.toString()}`, { headers: authHeaders() });
    const rows = await res.json();
    if (!res.ok) throw new Error(rows.error || "Failed to load work orders");
    mcWorkOrderOptions = Array.isArray(rows) ? rows : [];
    dl.innerHTML = mcWorkOrderOptions
      .map((wo) => {
        const id = Number(wo?.id || 0);
        if (!id) return "";
        const label = `WO #${id}`;
        const detail = `${wo.asset_code || "-"} — ${wo.source || "work order"}`;
        return `<option value="${esc(label)}">${esc(detail)}</option>`;
      })
      .join("");
  } catch {
    mcWorkOrderOptions = [];
    dl.innerHTML = "";
  }
}

async function saveMcAllRows() {
  const msg = document.getElementById("mcFormMsg");
  syncMcDateBounds();
  const work_date = String(document.getElementById("mcWorkDate")?.value || "").trim();
  const replace = Boolean(document.getElementById("mcReplaceDay")?.checked);
  const { entries, errors } = readMcDraftEntriesForSave();

  if (!work_date) {
    if (msg) { msg.className = "message-error"; msg.textContent = "Work date is required."; }
    return;
  }
  if (errors.length) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = errors.slice(0, 4).join(" · ") + (errors.length > 4 ? ` (+${errors.length - 4} more)` : "");
    }
    return;
  }
  if (!entries.length) {
    if (msg) { msg.className = "message-error"; msg.textContent = "Add at least one complete row before saving."; }
    return;
  }
  if (replace && mcDayRowsCache.length) {
    const ok = window.confirm(
      `Replace all ${mcDayRowsCache.length} saved entr${mcDayRowsCache.length === 1 ? "y" : "ies"} for ${work_date} with ${entries.length} row(s) from the sheet?`,
    );
    if (!ok) return;
  }

  if (msg) {
    msg.className = "muted";
    msg.textContent = `Saving ${entries.length} row(s)…`;
  }
  try {
    const res = await fetch(`${API}/maintenance/mechanic-labor/batch`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ work_date, entries, mode: replace ? "replace" : "append" }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Batch save failed");
    if (msg) {
      msg.className = "message-success";
      const skipped = Number(data.skipped || 0);
      msg.textContent = `Saved ${Number(data.saved || entries.length)} row(s) for ${work_date}.${skipped ? ` Skipped ${skipped} blank row(s).` : ""}`;
    }
    clearMcDraft();
    const replaceEl = document.getElementById("mcReplaceDay");
    if (replaceEl) replaceEl.checked = false;
    await loadMcDayEntries();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = e.message || String(e);
    }
  }
}

function renderMcDaySummary(totals) {
  const el = document.getElementById("mcDaySummary");
  if (!el) return;
  const hours = Number(totals?.hours || 0);
  const cost = Number(totals?.labor_cost || 0);
  const entries = Number(totals?.entries || 0);
  el.innerHTML = `
    <span class="maintenance-inline-metric">Entries <strong>${entries}</strong></span>
    <span class="maintenance-inline-metric">Total hours <strong>${hours.toFixed(2)}</strong></span>
    <span class="maintenance-inline-metric">Labor cost <strong>$${cost.toFixed(2)}</strong></span>
  `;
}

function renderMcDayTable(rows) {
  const body = document.getElementById("mcDayBody");
  if (!body) return;
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) {
    body.innerHTML = `<tr><td colspan="13" class="muted">No labor entries for this date.</td></tr>`;
    return;
  }
  body.innerHTML = list
    .map((r) => {
      const id = Number(r.id || 0);
      return `<tr data-mc-id="${id}">
        <td>${esc(mcDisplayDate(r.work_date))}</td>
        <td>${esc(r.asset_code || "")}</td>
        <td style="text-align:right;">${Number(r.hours || 0).toFixed(2)}</td>
        <td>${esc(r.category || "—")}</td>
        <td>${esc(r.reason || "")}</td>
        <td>${esc(r.time_started || "—")}</td>
        <td>${esc(r.time_finished || "—")}</td>
        <td>${esc(r.technician_name || "")}</td>
        <td>${esc(r.job_card_no || "—")}</td>
        <td style="text-align:right;">${r.smr != null && r.smr !== "" ? Number(r.smr).toFixed(1) : "—"}</td>
        <td style="text-align:right;">$${Number(r.labor_rate_per_hour || 0).toFixed(2)}</td>
        <td style="text-align:right;">$${Number(r.labor_cost || 0).toFixed(2)}</td>
        <td>
          <button type="button" data-mc-add-sheet="${id}">Add to sheet</button>
          <button type="button" data-mc-delete="${id}">Delete</button>
        </td>
      </tr>`;
    })
    .join("");
}

async function loadMcDayEntries() {
  syncMcDateBounds();
  const date = String(document.getElementById("mcWorkDate")?.value || "").trim();
  const body = document.getElementById("mcDayBody");
  if (!date) {
    if (body) body.innerHTML = `<tr><td colspan="13" class="muted">Pick a work date.</td></tr>`;
    return;
  }
  if (body) body.innerHTML = `<tr><td colspan="13" class="muted">Loading…</td></tr>`;
  try {
    const res = await fetch(`${API}/maintenance/mechanic-labor?date=${encodeURIComponent(date)}`, {
      headers: authHeaders(),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load entries");
    const rateEl = document.getElementById("mcDefaultRate");
    if (rateEl && data?.default_labor_rate != null) {
      rateEl.value = String(Number(data.default_labor_rate));
    }
    renderMcDayTable(data.rows);
    mcDayRowsCache = Array.isArray(data.rows) ? data.rows : [];
    renderMcDaySummary(data.totals);
  } catch (e) {
    if (body) body.innerHTML = `<tr><td colspan="13" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

async function saveMcDefaultRate() {
  const msg = document.getElementById("mcFormMsg");
  const rate = Math.max(0, Number(document.getElementById("mcDefaultRate")?.value || 0));
  if (!Number.isFinite(rate) || rate <= 0) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Enter a valid default labor rate.";
    }
    return;
  }
  if (msg) {
    msg.className = "muted";
    msg.textContent = "Saving default rate…";
  }
  try {
    const res = await fetch(`${API}/maintenance/mechanic-labor/settings`, {
      method: "PUT",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ labor_rate_per_hour: rate }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save rate");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `Default labor rate saved: $${rate.toFixed(2)}/hr`;
    }
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = e.message || String(e);
    }
  }
}

function addMcDraftRowFromSaved(row) {
  if (!row) return;
  mcDraftRows.push({
    technician_name: String(row.technician_name || ""),
    hours: row.hours != null ? String(row.hours) : "",
    asset_code: String(row.asset_code || ""),
    category: String(row.category || ""),
    reason: String(row.reason || ""),
    time_started: String(row.time_started || ""),
    time_finished: String(row.time_finished || ""),
    job_card_no: String(row.job_card_no || ""),
    smr: row.smr != null ? String(row.smr) : "",
  });
  refreshMcDraftEditor();
  document.getElementById("mcDraftBody")?.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

async function deleteMcEntry(id) {
  const entryId = Number(id || 0);
  if (!entryId) return;
  const ok = window.confirm(`Delete labor entry #${entryId}?`);
  if (!ok) return;
  const res = await fetch(`${API}/maintenance/mechanic-labor/${entryId}`, {
    method: "DELETE",
    headers: authHeaders(),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Delete failed");
  await loadMcDayEntries();
}

async function downloadMcYearXlsx() {
  const year = mcYearValue();
  const msg = document.getElementById("mcFormMsg");
  if (msg) {
    msg.className = "muted";
    msg.textContent = `Generating ${year} workbook…`;
  }
  try {
    const res = await fetch(`${API}/maintenance/mechanic-labor.xlsx?year=${encodeURIComponent(year)}`, {
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
    a.download = `mechanics-cost-${year}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `Downloaded mechanics cost report for ${year}.`;
    }
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = e.message || String(e);
    }
  }
}

async function downloadMcTimesheetXlsx() {
  const year = mcYearValue();
  const msg = document.getElementById("mcFormMsg");
  if (msg) {
    msg.className = "muted";
    msg.textContent = `Downloading ${year} mechanics timesheet…`;
  }
  try {
    const qs = new URLSearchParams({ year: String(year) });
    const res = await fetch(`${API}/maintenance/mechanic-labor/timesheet.xlsx?${qs.toString()}`, {
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
    a.download = `mechanics-timesheet-${year}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `Downloaded mechanics timesheet for ${year} (saved entries only).`;
    }
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = e.message || String(e);
    }
  }
}

async function uploadMcTimesheetFile() {
  const msg = document.getElementById("mcTimesheetUploadMsg") || document.getElementById("mcFormMsg");
  const input = document.getElementById("mcTimesheetFile");
  const file = input?.files?.[0];
  if (!file) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Choose an .xlsx or .csv timesheet file first.";
    }
    return;
  }
  const replaceDates = Boolean(document.getElementById("mcTimesheetReplaceDates")?.checked);
  const mode = replaceDates ? "replace_dates" : "append";
  if (msg) {
    msg.className = "muted";
    msg.textContent = `Uploading ${file.name}…`;
  }
  try {
    const form = new FormData();
    form.append("file", file, file.name);
    const res = await fetch(`${API}/maintenance/mechanic-labor/timesheet/import?mode=${encodeURIComponent(mode)}`, {
      method: "POST",
      headers: authHeaders(),
      body: form,
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || `Upload failed (${res.status})`);
    const warn = Array.isArray(data.warnings) && data.warnings.length
      ? ` (${data.warnings.length} row warning(s))`
      : "";
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `Imported ${Number(data.imported || 0)} row(s) across ${Number(data.date_count || 0)} date(s)${replaceDates ? `; replaced ${Number(data.deleted || 0)} existing` : ""}${warn}.`;
    }
    if (input) input.value = "";
    await loadMcDayEntries();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = e.message || String(e);
    }
  }
}

async function initMechanicsCostSection() {
  syncMcDateBounds();
  await Promise.all([loadMcTechnicians(), loadMcAssetDatalist(), loadMcJobCardDatalist()]);
  await loadMcDayEntries();
}
