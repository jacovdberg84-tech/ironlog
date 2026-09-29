// IRONLOG/web/maintenance/weekly-inspections.js — Weekly inspection calendar and roster.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

let wiCalendarCache = null;
let wiSelectedDate = "";

function wiTodayYmd() {
  return new Date().toISOString().slice(0, 10);
}

function wiMonthInput() {
  return document.getElementById("wiMonth");
}

function wiCurrentMonth() {
  const raw = String(wiMonthInput()?.value || "").trim();
  if (/^\d{4}-\d{2}$/.test(raw)) return raw;
  return new Date().toISOString().slice(0, 7);
}

function wiSetMonth(ym) {
  const inp = wiMonthInput();
  if (inp) inp.value = String(ym || "").trim();
}

function wiShiftMonth(delta) {
  const [y, m] = wiCurrentMonth().split("-").map(Number);
  const d = new Date(y, m - 1 + Number(delta || 0), 1);
  wiSetMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
}

function wiSetMsg(text, isError = false) {
  const msg = document.getElementById("wiMsg");
  if (!msg) return;
  msg.className = isError ? "message-error" : "muted";
  msg.textContent = String(text || "");
}

function wiMonthTitle(ym) {
  const [y, m] = String(ym || "").split("-").map(Number);
  if (!y || !m) return String(ym || "");
  return new Date(y, m - 1, 1).toLocaleDateString(undefined, { month: "long", year: "numeric" });
}

function wiNextStatus(status) {
  const s = String(status || "pending").toLowerCase();
  if (s === "pending") return "done";
  if (s === "done") return "skipped";
  return "pending";
}

function wiStatusLabel(status) {
  const s = String(status || "pending").toLowerCase();
  if (s === "done") return "Released";
  if (s === "skipped") return "Skipped";
  return "Pending";
}

function wiFormatMinutes(mins) {
  const n = Math.max(0, Number(mins || 0));
  if (n < 60) return `${n}m`;
  const h = Math.floor(n / 60);
  const m = n % 60;
  return m ? `${h}h ${m}m` : `${h}h`;
}

function wiFormatDayTitle(ymd) {
  const d = new Date(`${String(ymd).trim()}T12:00:00`);
  if (Number.isNaN(d.getTime())) return String(ymd || "");
  return d.toLocaleDateString(undefined, { weekday: "long", day: "numeric", month: "long" });
}

function wiSlotChipClass(status) {
  const s = String(status || "pending").toLowerCase();
  if (s === "done") return "wi-slot-chip--released";
  if (s === "skipped") return "wi-slot-chip--skipped";
  return "wi-slot-chip--pending";
}

function wiSlotChipClassForDate(status, ymd) {
  const base = wiSlotChipClass(status);
  const s = String(status || "pending").toLowerCase();
  if (s !== "done" && String(ymd || "") < wiTodayYmd()) return `${base} wi-slot-chip--overdue`;
  return base;
}

function wiDayHasNonCompliance(date, slots) {
  const ymd = String(date || "");
  if (!ymd || ymd >= wiTodayYmd()) return false;
  return (slots || []).some((slot) => String(slot?.status || "pending").toLowerCase() !== "done");
}

function wiFormatShortDate(ymd) {
  const d = new Date(`${String(ymd).trim()}T12:00:00`);
  if (Number.isNaN(d.getTime())) return String(ymd || "");
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

function wiUpcomingStatusMeta(status, date) {
  const s = String(status || "pending").toLowerCase();
  if (s === "done") return { label: "Released", cls: "wi-upcoming-badge--done" };
  if (String(date || "") < wiTodayYmd()) return { label: "Non-compliant", cls: "wi-upcoming-badge--overdue" };
  if (s === "skipped") return { label: "Skipped", cls: "wi-upcoming-badge--skipped" };
  return { label: "Pending", cls: "wi-upcoming-badge--pending" };
}

function renderWiUpcomingSidebar(data) {
  const list = document.getElementById("wiUpcomingList");
  if (!list) return;
  const today = wiTodayYmd();
  const items = [];
  const weeks = Array.isArray(data?.calendar_weeks) ? data.calendar_weeks : [];
  for (const row of weeks) {
    for (const cell of row || []) {
      if (!cell?.date) continue;
      for (const slot of cell.slots || []) {
        items.push({
          date: cell.date,
          asset_code: slot.asset_code,
          asset_name: slot.asset_name,
          status: slot.status,
          est_minutes: slot.est_minutes,
        });
      }
    }
  }
  if (!items.length) {
    list.innerHTML = `<div class="empty">No inspections scheduled this month.</div>`;
    return;
  }
  items.sort((a, b) => {
    const aOver = a.date < today && String(a.status || "").toLowerCase() !== "done";
    const bOver = b.date < today && String(b.status || "").toLowerCase() !== "done";
    if (aOver !== bOver) return aOver ? -1 : 1;
    return String(a.date).localeCompare(String(b.date)) || String(a.asset_code).localeCompare(String(b.asset_code));
  });
  const shown = items.slice(0, 24);
  list.innerHTML = shown.map((item) => {
    const meta = wiUpcomingStatusMeta(item.status, item.date);
    const title = item.asset_name ? esc(item.asset_name) : esc(item.asset_code);
    const sub = item.asset_name ? esc(item.asset_code) : "";
    return `
      <button type="button" class="wi-upcoming-item" data-wi-open-day="${esc(item.date)}">
        <div class="wi-upcoming-item-head">
          <strong>${title}</strong>
          <span class="wi-upcoming-badge ${meta.cls}">${esc(meta.label)}</span>
        </div>
        <div class="wi-upcoming-item-meta muted mini">
          <span>${esc(wiFormatShortDate(item.date))}</span>
          ${sub ? `<span>· ${sub}</span>` : ""}
          <span>· ${esc(wiFormatMinutes(item.est_minutes))}</span>
        </div>
      </button>
    `;
  }).join("");
  if (items.length > shown.length) {
    list.innerHTML += `<div class="muted mini" style="padding:8px 4px;">+${items.length - shown.length} more this month</div>`;
  }
}

function populateWiDayAssetSelect(assets) {
  const sel = document.getElementById("wiDayAssetSelect");
  if (!sel) return;
  const rows = Array.isArray(assets) ? assets : [];
  if (!rows.length) {
    sel.innerHTML = `<option value="">Add equipment to roster first</option>`;
    return;
  }
  sel.innerHTML = `<option value="">Select equipment</option>${rows.map((a) =>
    `<option value="${Number(a.asset_id)}">${esc(a.asset_code)} — ${esc(a.asset_name)}</option>`
  ).join("")}`;
}

function renderWeeklyInspectionCompliance(compliance) {
  const card = document.getElementById("wiComplianceCard");
  if (!card) return;
  const c = compliance || {};
  const gaps = Array.isArray(c.weekly_gaps) ? c.weekly_gaps.length : 0;
  card.innerHTML = `
    <div class="wi-kpi">
      <div class="wi-kpi-value">${Number(c.done_count ?? 0)} / ${Number(c.total_slots ?? 0)}</div>
      <div class="wi-kpi-label">Released</div>
      <div class="wi-kpi-meta">${Number(c.pending_count ?? 0)} pending · ${Number(c.skipped_count ?? 0)} skipped</div>
    </div>
    <div class="wi-kpi wi-kpi--bad">
      <div class="wi-kpi-value">${Number(c.not_released_count ?? 0)}</div>
      <div class="wi-kpi-label">Overdue</div>
      <div class="wi-kpi-meta">Past date, not released</div>
    </div>
    <div class="wi-kpi${gaps ? " wi-kpi--warn" : ""}">
      <div class="wi-kpi-value">${gaps}</div>
      <div class="wi-kpi-label">Missing weekly visit</div>
      <div class="wi-kpi-meta">Roster machine with no slot that week</div>
    </div>
    <div class="wi-kpi">
      <div class="wi-kpi-value">${wiFormatMinutes(c.est_minutes_total)}</div>
      <div class="wi-kpi-label">Planned time</div>
      <div class="wi-kpi-meta">${wiFormatMinutes(c.est_minutes_released)} released</div>
    </div>
  `;
}

function renderWeeklyInspectionWeeklyGaps(compliance) {
  const wrap = document.getElementById("wiWeeklyGaps");
  if (!wrap) return;
  const gaps = Array.isArray(compliance?.weekly_gaps) ? compliance.weekly_gaps : [];
  if (!gaps.length) {
    wrap.innerHTML = "";
    return;
  }
  const lines = gaps.slice(0, 12).map((g) =>
    `<span class="wi-gap-chip">${esc(g.asset_code)} · week ${esc(String(g.week_start || "").slice(5))}</span>`
  ).join("");
  const more = gaps.length > 12 ? `<span class="muted mini">+${gaps.length - 12} more</span>` : "";
  wrap.innerHTML = `
    <div class="wi-weekly-gaps-box">
      <strong>Missing weekly workshop visit</strong>
      <div class="wi-gap-chips">${lines}${more}</div>
    </div>
  `;
}

function renderWeeklyInspectionAssetList(assets) {
  const list = document.getElementById("wiAssetList");
  if (!list) return;
  const rows = Array.isArray(assets) ? assets : [];
  if (!rows.length) {
    list.innerHTML = `<div class="empty">No equipment on the workshop roster yet. Add machines above, then click calendar days to schedule them.</div>`;
    return;
  }
  list.innerHTML = rows.map((a) => `
    <div class="item" style="display:flex; justify-content:space-between; align-items:center; gap:10px; flex-wrap:wrap;">
      <div>
        <strong>${esc(a.asset_code)}</strong> — ${esc(a.asset_name)}
        <div class="muted mini">Default est. ${wiFormatMinutes(a.est_minutes || 30)}</div>
        ${a.notes ? `<div class="muted mini">${esc(a.notes)}</div>` : ""}
      </div>
      <div class="row stack-10 wi-no-print" style="align-items:center;">
        <label class="mini muted">
          Est.
          <input type="number" min="5" step="5" class="w-70" data-wi-est="${Number(a.id)}" value="${Number(a.est_minutes || 30)}" />
        </label>
        <button type="button" data-wi-save-est="${Number(a.id)}">Save</button>
        <button type="button" data-wi-remove="${Number(a.id)}">Remove</button>
      </div>
    </div>
  `).join("");
}

function renderWiDaySlotList(date) {
  const list = document.getElementById("wiDaySlotList");
  if (!list) return;
  const weeks = Array.isArray(wiCalendarCache?.calendar_weeks) ? wiCalendarCache.calendar_weeks : [];
  let slots = [];
  for (const row of weeks) {
    const cell = (row || []).find((c) => String(c?.date) === String(date));
    if (cell) {
      slots = Array.isArray(cell.slots) ? cell.slots : [];
      break;
    }
  }
  if (!slots.length) {
    list.innerHTML = `<div class="empty">No equipment scheduled for this day yet.</div>`;
    return;
  }
  list.innerHTML = slots.map((slot) => `
    <div class="wi-day-slot-row">
      <button
        type="button"
        class="wi-slot-chip ${wiSlotChipClassForDate(slot.status, wiSelectedDate)}"
        data-wi-slot-status="${Number(slot.id)}"
        data-wi-status="${esc(String(slot.status || "pending"))}"
        title="Click to cycle: Pending → Released → Skipped"
      >${esc(slot.asset_code)} · ${wiFormatMinutes(slot.est_minutes)} · ${esc(wiStatusLabel(slot.status))}</button>
      <button type="button" class="wi-slot-remove" data-wi-slot-remove="${Number(slot.id)}">Remove</button>
    </div>
  `).join("");
}

function openWiDayPanel(date) {
  wiSelectedDate = String(date || "");
  const panel = document.getElementById("wiDayPanel");
  const title = document.getElementById("wiDayPanelTitle");
  if (!panel || !wiSelectedDate) return;
  if (title) title.textContent = wiFormatDayTitle(wiSelectedDate);
  const estInp = document.getElementById("wiDayEst");
  if (estInp && !estInp.value) estInp.value = "30";
  const copyInp = document.getElementById("wiCopyToDate");
  if (copyInp) {
    const d = new Date(`${wiSelectedDate}T12:00:00`);
    d.setDate(d.getDate() + 7);
    copyInp.min = wiTodayYmd();
    copyInp.value = d.toISOString().slice(0, 10);
  }
  populateWiDayAssetSelect(wiCalendarCache?.assets || []);
  renderWiDaySlotList(wiSelectedDate);
  panel.style.display = "block";
  panel.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function closeWiDayPanel() {
  wiSelectedDate = "";
  const panel = document.getElementById("wiDayPanel");
  if (panel) panel.style.display = "none";
}

function renderWeeklyInspectionCalendar(data) {
  const wrap = document.getElementById("wiCalendarWrap");
  const title = document.getElementById("wiMonthTitle");
  if (!wrap) return;
  const weeks = Array.isArray(data?.calendar_weeks) ? data.calendar_weeks : [];
  if (title) title.textContent = wiMonthTitle(data?.month || wiCurrentMonth());
  renderWeeklyInspectionCompliance(data?.compliance);
  renderWeeklyInspectionWeeklyGaps(data?.compliance);
  renderWeeklyInspectionAssetList(data?.assets || []);
  renderWiUpcomingSidebar(data);
  populateWiDayAssetSelect(data?.assets || []);
  if (!weeks.length) {
    wrap.innerHTML = `<div class="empty">No calendar data for this month.</div>`;
    return;
  }
  const dayNames = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
  const head = dayNames.map((d) => `<div class="wi-month-dow">${d}</div>`).join("");
  const body = weeks.map((row) => {
    const cells = (row || []).map((cell) => {
      if (!cell?.in_month || !cell?.date) {
        return `<div class="wi-month-cell wi-month-cell--pad"></div>`;
      }
      const slots = Array.isArray(cell.slots) ? cell.slots : [];
      const chips = slots.map((slot) => `
        <button
          type="button"
          class="wi-slot-chip ${wiSlotChipClassForDate(slot.status, cell.date)}"
          data-wi-slot-status="${Number(slot.id)}"
          data-wi-status="${esc(String(slot.status || "pending"))}"
          title="${esc(wiStatusLabel(slot.status))}"
        >${esc(slot.asset_code)}</button>
      `).join("");
      const todayCls = cell.is_today ? " wi-month-cell--today" : "";
      const selectedCls = String(cell.date) === wiSelectedDate ? " wi-month-cell--selected" : "";
      const nonCompliantCls = wiDayHasNonCompliance(cell.date, slots) ? " wi-month-cell--noncompliant" : "";
      return `
        <div
          class="wi-month-cell${todayCls}${selectedCls}${nonCompliantCls}"
          role="button"
          tabindex="0"
          data-wi-open-day="${esc(cell.date)}"
        >
          <div class="wi-month-daynum">${Number(cell.day || 0)}</div>
          <div class="wi-month-chips">${chips}</div>
        </div>
      `;
    }).join("");
    return `<div class="wi-month-row">${cells}</div>`;
  }).join("");
  wrap.innerHTML = `
    <div class="wi-month-grid">
      <div class="wi-month-head">${head}</div>
      ${body}
    </div>
  `;
  if (wiSelectedDate) renderWiDaySlotList(wiSelectedDate);
}

async function loadWeeklyInspectionCalendar() {
  const wrap = document.getElementById("wiCalendarWrap");
  if (wrap) wrap.innerHTML = `<div class="empty">Loading calendar...</div>`;
  wiSetMsg("Loading workshop calendar...");
  const q = new URLSearchParams();
  q.set("month", wiCurrentMonth());
  try {
    const res = await fetch(`${API}/maintenance/weekly-inspections/calendar?${q.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load calendar");
    wiCalendarCache = data;
    renderWeeklyInspectionCalendar(data);
    const assetCount = Array.isArray(data.assets) ? data.assets.length : 0;
    const slotCount = Number(data?.compliance?.total_slots ?? 0);
    const gaps = Array.isArray(data?.compliance?.weekly_gaps) ? data.compliance.weekly_gaps.length : 0;
    wiSetMsg(`${assetCount} roster machine${assetCount === 1 ? "" : "s"} · ${slotCount} scheduled visit${slotCount === 1 ? "" : "s"}${gaps ? ` · ${gaps} missing weekly visit${gaps === 1 ? "" : "s"}` : ""}.`);
  } catch (e) {
    wiCalendarCache = null;
    if (wrap) wrap.innerHTML = `<div class="empty">Load failed: ${esc(e.message || String(e))}</div>`;
    wiSetMsg(`Load failed: ${e.message || e}`, true);
  }
}

async function cycleWeeklyInspectionSlot(slotId, currentStatus) {
  const next = wiNextStatus(currentStatus);
  const inspector_name = String(document.getElementById("wiInspector")?.value || "").trim();
  wiSetMsg("Saving...");
  try {
    const res = await fetch(`${API}/maintenance/weekly-inspections/entries`, {
      method: "PUT",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ slot_id: Number(slotId), status: next, inspector_name }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to update status");
    await loadWeeklyInspectionCalendar();
    if (wiSelectedDate) openWiDayPanel(wiSelectedDate);
    wiSetMsg(`Marked ${wiStatusLabel(next).toLowerCase()}.`);
  } catch (e) {
    wiSetMsg(`Update failed: ${e.message || e}`, true);
  }
}

async function addWeeklyInspectionSlot() {
  const planned_date = String(wiSelectedDate || "").trim();
  const asset_id = Number(document.getElementById("wiDayAssetSelect")?.value || 0);
  const est_minutes = Math.max(5, Number(document.getElementById("wiDayEst")?.value || 30) || 30);
  if (!planned_date) {
    wiSetMsg("Click a calendar day first.", true);
    return;
  }
  if (!asset_id) {
    wiSetMsg("Select equipment to schedule.", true);
    return;
  }
  wiSetMsg("Adding to day...");
  try {
    const res = await fetch(`${API}/maintenance/weekly-inspections/slots`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ planned_date, asset_id, est_minutes }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to add equipment to day");
    await loadWeeklyInspectionCalendar();
    openWiDayPanel(planned_date);
    wiSetMsg(`Added to ${wiFormatDayTitle(planned_date)}.`);
  } catch (e) {
    wiSetMsg(`Add failed: ${e.message || e}`, true);
  }
}

async function copyWeeklyInspectionDay() {
  const from_date = String(wiSelectedDate || "").trim();
  const to_date = String(document.getElementById("wiCopyToDate")?.value || "").trim();
  if (!from_date) {
    wiSetMsg("Click a calendar day first.", true);
    return;
  }
  if (!to_date) {
    wiSetMsg("Choose the target date to copy to.", true);
    return;
  }
  if (from_date === to_date) {
    wiSetMsg("Choose a different target date.", true);
    return;
  }
  wiSetMsg("Copying day...");
  try {
    const res = await fetch(`${API}/maintenance/weekly-inspections/slots/copy-day`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ from_date, to_date }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to copy day");
    const targetMonth = to_date.slice(0, 7);
    if (targetMonth !== wiCurrentMonth()) wiSetMonth(targetMonth);
    await loadWeeklyInspectionCalendar();
    openWiDayPanel(to_date);
    const skipped = Number(data.skipped || 0);
    wiSetMsg(
      `Copied ${Number(data.copied || 0)} item${Number(data.copied || 0) === 1 ? "" : "s"} to ${wiFormatDayTitle(to_date)}${skipped ? ` (${skipped} already scheduled)` : ""}.`,
    );
  } catch (e) {
    wiSetMsg(`Copy failed: ${e.message || e}`, true);
  }
}

function toggleWiCompliancePanel() {
  const card = document.getElementById("wiComplianceCard");
  const gaps = document.getElementById("wiWeeklyGaps");
  if (!card) return;
  const hidden = card.style.display === "none";
  card.style.display = hidden ? "" : "none";
  if (gaps) gaps.style.display = hidden ? "" : "none";
}

async function removeWeeklyInspectionSlot(slotId) {
  const id = Number(slotId || 0);
  if (!id) return;
  wiSetMsg("Removing...");
  try {
    const res = await fetch(`${API}/maintenance/weekly-inspections/slots/${id}`, {
      method: "DELETE",
      headers: authHeaders(),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to remove slot");
    await loadWeeklyInspectionCalendar();
    if (wiSelectedDate) openWiDayPanel(wiSelectedDate);
    wiSetMsg("Removed from day.");
  } catch (e) {
    wiSetMsg(`Remove failed: ${e.message || e}`, true);
  }
}

async function saveWeeklyInspectionAssetEst(id) {
  const rowId = Number(id || 0);
  if (!rowId) return;
  const inp = document.querySelector(`input[data-wi-est="${rowId}"]`);
  const est_minutes = Math.max(5, Number(inp?.value || 30) || 30);
  wiSetMsg("Saving estimate...");
  try {
    const res = await fetch(`${API}/maintenance/weekly-inspections/assets/${rowId}`, {
      method: "PUT",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ est_minutes }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save estimate");
    await loadWeeklyInspectionCalendar();
    wiSetMsg("Default estimate saved.");
  } catch (e) {
    wiSetMsg(`Save failed: ${e.message || e}`, true);
  }
}

async function addWeeklyInspectionAsset() {
  const asset_id = Number(document.getElementById("wiAssetSelect")?.value || 0);
  const notes = String(document.getElementById("wiAssetNotes")?.value || "").trim();
  const est_minutes = Math.max(5, Number(document.getElementById("wiAssetEst")?.value || 30) || 30);
  if (!asset_id) {
    wiSetMsg("Select equipment to add.", true);
    return;
  }
  wiSetMsg("Adding to roster...");
  try {
    const res = await fetch(`${API}/maintenance/weekly-inspections/assets`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ asset_id, notes, est_minutes }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to add equipment");
    const notesInp = document.getElementById("wiAssetNotes");
    if (notesInp) notesInp.value = "";
    await loadWeeklyInspectionCalendar();
    wiSetMsg("Equipment added to workshop roster.");
  } catch (e) {
    wiSetMsg(`Add failed: ${e.message || e}`, true);
  }
}

async function removeWeeklyInspectionAsset(id) {
  const rowId = Number(id || 0);
  if (!rowId) return;
  if (!window.confirm("Remove this equipment from the workshop roster?")) return;
  wiSetMsg("Removing...");
  try {
    const res = await fetch(`${API}/maintenance/weekly-inspections/assets/${rowId}`, {
      method: "DELETE",
      headers: authHeaders(),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to remove equipment");
    await loadWeeklyInspectionCalendar();
    wiSetMsg("Equipment removed from roster.");
  } catch (e) {
    wiSetMsg(`Remove failed: ${e.message || e}`, true);
  }
}

async function clearWeeklyInspectionRoster() {
  const rosterCount = Array.isArray(wiCalendarCache?.assets) ? wiCalendarCache.assets.length : 0;
  const slotCount = Number(wiCalendarCache?.compliance?.total_slots ?? 0);
  if (!rosterCount && !slotCount) {
    wiSetMsg("Roster is already empty.");
    return;
  }
  const msg = [
    "Clear the entire workshop roster and start fresh?",
    rosterCount ? `${rosterCount} roster item${rosterCount === 1 ? "" : "s"} will be removed.` : "",
    slotCount ? `${slotCount} scheduled calendar visit${slotCount === 1 ? "" : "s"} will also be cleared.` : "",
  ].filter(Boolean).join("\n\n");
  if (!window.confirm(msg)) return;
  wiSetMsg("Clearing roster...");
  try {
    const res = await fetch(`${API}/maintenance/weekly-inspections/roster?clear_slots=1`, {
      method: "DELETE",
      headers: authHeaders(),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to clear roster");
    closeWiDayPanel();
    wiCalendarCache = { ...wiCalendarCache, assets: [], calendar_weeks: [], compliance: { ...wiCalendarCache?.compliance, weekly_gaps: [], day_agenda: [], total_slots: 0, done_count: 0 } };
    const assetList = document.getElementById("wiAssetList");
    if (assetList) {
      assetList.innerHTML = `<div class="empty">No equipment on the workshop roster yet. Add machines above, then click calendar days to schedule them.</div>`;
    }
    await loadWeeklyInspectionCalendar();
    wiSetMsg(
      `Roster cleared (${Number(data.roster_cleared || 0)} item${Number(data.roster_cleared || 0) === 1 ? "" : "s"}${Number(data.slots_cleared || 0) ? `, ${Number(data.slots_cleared)} calendar visit${Number(data.slots_cleared) === 1 ? "" : "s"} removed` : ""}).`,
    );
  } catch (e) {
    wiSetMsg(`Clear failed: ${e.message || e}`, true);
  }
}

async function fetchWeeklyInspectionPdfBlob() {
  const q = new URLSearchParams();
  q.set("month", wiCurrentMonth());
  q.set("_", String(Date.now()));
  const res = await fetch(`${API}/maintenance/weekly-inspections.pdf?${q.toString()}`, {
    headers: authHeaders(),
    cache: "no-store",
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(txt || `PDF request failed (${res.status})`);
  }
  return res.blob();
}

async function openWeeklyInspectionPdf(download = false) {
  try {
    wiSetMsg("Generating PDF...");
    const blob = await fetchWeeklyInspectionPdfBlob();
    const blobUrl = URL.createObjectURL(blob);
    if (download) {
      const a = document.createElement("a");
      a.href = blobUrl;
      a.download = `IRONLOG_Workshop_Inspections-${wiCurrentMonth()}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(blobUrl), 3000);
      wiSetMsg("PDF downloaded.");
      return;
    }
    window.open(blobUrl, "_blank");
    setTimeout(() => URL.revokeObjectURL(blobUrl), 60000);
    wiSetMsg("PDF opened.");
  } catch (e) {
    wiSetMsg(`PDF error: ${e.message || e}`, true);
  }
}

async function printWeeklyInspectionCalendar() {
  let blobUrl = "";
  let iframe = null;
  try {
    wiSetMsg("Preparing branded PDF for print...");
    const blob = await fetchWeeklyInspectionPdfBlob();
    blobUrl = URL.createObjectURL(blob);
    iframe = document.createElement("iframe");
    iframe.setAttribute("title", "Workshop inspection print preview");
    iframe.style.cssText = "position:fixed;right:0;bottom:0;width:0;height:0;border:0;";
    iframe.src = blobUrl;
    document.body.appendChild(iframe);
    iframe.onload = () => {
      setTimeout(() => {
        try {
          iframe.contentWindow?.focus();
          iframe.contentWindow?.print();
          wiSetMsg("Print dialog opened (landscape PDF with company branding).");
        } catch (e) {
          window.open(blobUrl, "_blank");
          wiSetMsg("Opened PDF in a new tab — use the browser print button there.", true);
        }
      }, 400);
    };
    window.addEventListener("afterprint", () => {
      iframe?.remove();
      if (blobUrl) URL.revokeObjectURL(blobUrl);
    }, { once: true });
    setTimeout(() => {
      iframe?.remove();
      if (blobUrl) URL.revokeObjectURL(blobUrl);
    }, 120000);
  } catch (e) {
    iframe?.remove();
    if (blobUrl) URL.revokeObjectURL(blobUrl);
    wiSetMsg(`Print failed: ${e.message || e}`, true);
  }
}
