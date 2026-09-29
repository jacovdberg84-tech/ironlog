// IRONLOG/web/app/daily-input.js — Daily input grid, prestart section, asset QR sheets, shift self-check.
// Part of the main app; index.html loads these files in order and they share one global scope.

let dailyRows = [];
let dailyShowDownOnly = false;
let dailyScheduledOverride = null;
let dailyPrestartRows = [];
let dailyPrestartMeta = {
  deduction_hours_per_check: 0.5,
  production_deduction_hours: 0,
  production_deduction_count: 0,
};

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
function getDayScheduledHours() {
  const value = toNum(qs("scheduled")?.value);
  return value != null && value > 0 && value <= 24 ? value : 10;
}

function applyDayScheduledHours(value = getDayScheduledHours()) {
  const scheduled = Number(value);
  if (!Number.isFinite(scheduled) || scheduled <= 0 || scheduled > 24) return false;
  dailyScheduledOverride = scheduled;
  for (const row of dailyRows) {
    if (row.is_master_standby) continue;
    if (row.is_used || row.is_down) row.scheduled_hours = scheduled;
  }
  if (dailyRows.length) {
    validateDailyRows();
    renderDailyTable();
    renderDailyPreview();
  }
  return true;
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
      if (r.incident_mode === "short") {
        if (r.down_lock) {
          r.error = "SHORT BREAKDOWN UNAVAILABLE — THIS ASSET ALREADY HAS AN OPEN REPAIR.";
          continue;
        }
        const shortDuration = updateDailyShortBreakdownHours(
          qs("date")?.value || todayLocalYmd(),
          r,
        );
        if (shortDuration == null) {
          r.error = "SHORT BREAKDOWN — ENTER A VALID TIME DOWN AND TIME UP.";
          continue;
        }
      }
      const dh = r.down_hours != null ? Number(r.down_hours) : Number(r.scheduled_hours || 0);
      if (!Number.isFinite(dh) || dh < 0) {
        r.error = "DOWN HOURS INVALID — MUST BE >= 0.";
        continue;
      }
      if (r.scheduled_hours != null && dh > Number(r.scheduled_hours || 0)) {
        r.error = "DOWN HOURS TOO HIGH — MUST BE <= SCHEDULED HOURS.";
        continue;
      }
      if (r.opening_hours != null && r.closing_hours != null && r.closing_hours < r.opening_hours) {
        r.error = "HOURMETER MISMATCH — CLOSING LOWER THAN OPENING.";
        continue;
      }
      // A machine can work part of the shift before it goes down. Keep those
      // meter inputs editable and only mark it as production when it ran.
      r.is_used = Number(r.hours_run || 0) > 0;
      continue;
    }

    if (r.telematics_locked) {
      if (!r.is_used && r.hours_run > 0) {
        r.error = "STANDBY SELECTED — HOURS NOT ALLOWED.";
        continue;
      }
      if (r.is_used && (r.scheduled_hours == null || r.scheduled_hours === 0)) {
        r.error = "PRODUCTION SELECTED — SCHEDULED HOURS IS 0.";
      } else if (r.is_used && r.hours_run === 0) {
        r.warning = "TELEMATICS — METER WILL SYNC FROM FSC ON SAVE.";
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
    (r) => (r.is_used || r.is_down) && !r.is_master_standby && String(r.input_unit || "hours").toLowerCase() !== "km"
  );
  const usedCount = used.length;

  let totalScheduled = 0;
  let totalRun = 0;
  let totalDowntime = 0;

  let totalPrestart = 0;
  let prestartCount = 0;
  const prestartPerCheck = Number(dailyPrestartMeta.deduction_hours_per_check || 0.5);
  const prestartCodes = new Set(
    dailyPrestartRows.map((r) => String(r.asset_code || "").trim()).filter(Boolean)
  );

  for (const r of used) {
    const scheduled = Math.max(0, Number(r.scheduled_hours || 0));
    const runRaw = Math.max(0, Number(r.hours_run || 0));
    const runEff = Math.min(runRaw, scheduled);
    const downRaw = r.is_down
      ? (Number.isFinite(Number(r.down_hours)) ? Number(r.down_hours) : scheduled)
      : 0;
    const downEff = Math.min(Math.max(0, downRaw), scheduled);

    totalScheduled += scheduled;
    totalRun += runEff;
    totalDowntime += downEff;

    if (prestartCodes.has(r.asset_code)) {
      totalPrestart += prestartPerCheck;
      prestartCount += 1;
    }
  }

  const utilization = totalScheduled > 0 ? (totalRun / totalScheduled) * 100 : null;
  const available = Math.max(0, totalScheduled - totalDowntime - totalPrestart);
  const availability = totalScheduled > 0 ? (available / totalScheduled) * 100 : null;

  return {
    usedCount,
    totalScheduled,
    totalRun,
    totalDowntime,
    totalPrestart,
    prestartCount,
    availability,
    utilization,
  };
}

function renderDailyPreview() {
  setText("dailySummary", daySummary());

  const k = calcDailyPreviewKpis();
  const th = getThresholds();

  setText("kUsed", `Used: ${k.usedCount}`);
  setText("kSched", `Scheduled: ${k.totalScheduled.toFixed(1).replace(/\.0$/, "")}`);
  setText("kRun", `Run: ${k.totalRun.toFixed(1).replace(/\.0$/, "")}`);

  setSpeedo(qs("pAvailNeedle"), qs("pAvailVal"), k.availability, {
    goodAt: th.availTarget,
    warnAt: th.availCrit,
  });
  setSpeedo(qs("pUtilNeedle"), qs("pUtilVal"), k.utilization, {
    goodAt: th.utilTarget,
    warnAt: th.utilCrit,
  });

  if (k.totalScheduled === 0) setText("kNote", "Preview waiting for scheduled/run hours. Standby excluded.");
  else {
    const prestartNote = k.prestartCount > 0
      ? ` Pre-start: ${k.prestartCount} production check(s), ${k.totalPrestart.toFixed(1).replace(/\.0$/, "")} hr deducted from availability.`
      : "";
    setText(
      "kNote",
      `Preview uses production rows only (standby excluded). Down hours: ${k.totalDowntime.toFixed(1).replace(/\.0$/, "")}.${prestartNote}`
    );
  }
}

function renderDailyPrestartSection() {
  const panel = qs("dailyPrestartPanel");
  if (!panel) return;

  const perCheck = Number(dailyPrestartMeta.deduction_hours_per_check || 0.5);
  const perCheckLabel = perCheck.toFixed(1).replace(/\.0$/, "");

  if (!dailyPrestartRows.length) {
    panel.innerHTML = `
      <div class="daily-prestart-header">
        <strong>Daily pre-start checks</strong>
        <span class="muted small">None logged for this date</span>
      </div>`;
    panel.classList.add("is-empty");
    return;
  }

  panel.classList.remove("is-empty");
  const prodCount = Number(dailyPrestartMeta.production_deduction_count || 0);
  const prodHours = Number(dailyPrestartMeta.production_deduction_hours || 0);
  const chips = dailyPrestartRows.map((r) => {
    const code = String(r.asset_code || "").trim();
    const type = String(r.check_type || "Pre-start");
    const inspector = r.inspector_name ? ` · ${String(r.inspector_name).trim()}` : "";
    return `<span class="daily-prestart-chip" title="${type}${inspector}"><b>${code}</b> ${type}</span>`;
  }).join("");

  panel.innerHTML = `
    <div class="daily-prestart-header">
      <strong>Daily pre-start checks</strong>
      <span class="muted small">${dailyPrestartRows.length} completed · ${perCheckLabel} hr each</span>
    </div>
    <div class="daily-prestart-list">${chips}</div>
    <div class="daily-prestart-foot muted small">
      ${prodCount > 0
        ? `${prodCount} production asset pre-start(s) deduct ${prodHours.toFixed(1).replace(/\.0$/, "")} hr from preview availability.`
        : "Pre-starts shown for reference; none apply to production hour-based availability today."}
    </div>`;
}

/* -------- DOWN helper -------- */

function dailyOperationsDate(captureDate) {
  return prevDateStr(String(captureDate || todayLocalYmd()).slice(0, 10));
}

function dailyShortBreakdownDurationHours(captureDate, timeDown, timeUp) {
  const down = String(timeDown || "").trim();
  const up = String(timeUp || "").trim();
  if (!/^\d{2}:\d{2}$/.test(down) || !/^\d{2}:\d{2}$/.test(up)) return null;
  const operationsDate = dailyOperationsDate(captureDate);
  const start = new Date(`${operationsDate}T${down}:00`);
  const end = new Date(`${operationsDate}T${up}:00`);
  const elapsed = (end.getTime() - start.getTime()) / 3600000;
  if (!Number.isFinite(elapsed) || elapsed <= 0 || elapsed > 24) return null;
  return Number(elapsed.toFixed(4));
}

function updateDailyShortBreakdownHours(captureDate, row) {
  const elapsed = dailyShortBreakdownDurationHours(
    captureDate,
    row.short_time_down,
    row.short_time_up,
  );
  if (elapsed != null) row.down_hours = elapsed;
  return elapsed;
}

async function saveDailyOffsiteProgress(captureDate, r, breakdownId) {
  if (!r.offsite_enabled) return null;

  const sentDate = String(r.offsite_sent_date || "").trim() || dailyOperationsDate(captureDate);
  const payload = {
    asset_code: r.asset_code,
    breakdown_id: Number(breakdownId || r.breakdown_id || 0) || null,
    repair_status: String(r.offsite_status || "sent_offsite").trim() || "sent_offsite",
    sent_date: sentDate,
    expected_return_date: String(r.offsite_expected_return_date || r.ets_repair_date || "").trim() || null,
    actual_return_date: String(r.offsite_actual_return_date || "").trim() || null,
    vendor: String(r.offsite_vendor || "").trim() || null,
    current_location: String(r.offsite_location || "").trim() || null,
    repair_reason: String(r.down_reason || "").trim() || null,
    notes: String(r.offsite_progress || r.repair_progress || r.breakdown_comment || "").trim() || null,
  };

  if (Number(r.offsite_repair_id || 0) > 0) {
    await fetchJson(`${API}/api/breakdown-ops/offsite-repairs/${Number(r.offsite_repair_id)}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    return Number(r.offsite_repair_id);
  }

  const created = await fetchJson(`${API}/api/breakdown-ops/offsite-repairs`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const id = Number(created?.id || 0) || null;
  if (id) r.offsite_repair_id = id;
  return id;
}

async function logDownRowToBreakdowns(date, r) {
  const operationsDate = dailyOperationsDate(date);
  const downDesc = r.down_reason ? `BREAKDOWN — ${r.down_reason}` : "BREAKDOWN";
  const b = await fetchJson(`${API}/api/breakdowns/ensure-open`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      asset_code: r.asset_code,
      breakdown_date: operationsDate,
      start_date: String(r.breakdown_start_date || "").trim() || operationsDate,
      description: downDesc,
      component: String(r.breakdown_component || "").trim() || null,
      critical: Boolean(r.breakdown_critical),
      parts_ordered_date: String(r.parts_ordered_date || "").trim() || null,
      parts_status: String(r.parts_status || "").trim() || null,
      parts_received_date: String(r.parts_received_date || "").trim() || null,
      ets_repair_date: String(r.ets_repair_date || "").trim() || null,
    }),
  });

  const breakdownId = Number(b.breakdown_id || b.breakdownId || b.id || 0);
  if (!breakdownId) throw new Error(`${r.asset_code}: breakdown was not created`);

  const notes = r.down_reason
    ? `Daily Log breakdown — ${r.down_reason}`
    : "Daily Log breakdown";
  const downHoursRaw = r.down_hours != null ? Number(r.down_hours) : Number(r.scheduled_hours || 0);
  const downHours = Number.isFinite(downHoursRaw) ? Math.max(0, downHoursRaw) : 0;

  await fetchJson(`${API}/api/breakdowns/${breakdownId}/downtime`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      log_date: operationsDate,
      hours_down: downHours,
      notes,
    }),
  });

  await fetchJson(`${API}/api/breakdowns/${breakdownId}/daily-progress`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      repair_progress: String(r.repair_progress || "").trim(),
      scheduled_repair_date: String(r.ets_repair_date || "").trim() || null,
    }),
  });

  const offsiteRepairId = await saveDailyOffsiteProgress(date, r, breakdownId);
  r.breakdown_id = breakdownId;
  r.work_order_id = Number(b.primary_work_order_id || r.work_order_id || 0) || null;
  return { breakdownId, offsiteRepairId };
}

async function logDailyShortBreakdown(date, r) {
  if (Number(r.short_breakdown_id || 0) > 0) {
    return { breakdownId: Number(r.short_breakdown_id), workOrderId: Number(r.work_order_id || 0) || null };
  }

  const operationsDate = dailyOperationsDate(date);
  const duration = updateDailyShortBreakdownHours(date, r);
  if (duration == null) {
    throw new Error(`${r.asset_code}: enter a valid time down and time up for the short breakdown`);
  }

  const descriptionParts = [
    r.down_reason ? `BREAKDOWN — ${r.down_reason}` : "BREAKDOWN",
    String(r.breakdown_comment || "").trim(),
  ].filter(Boolean);
  const result = await fetchJson(`${API}/api/breakdowns/short-complete`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      asset_code: r.asset_code,
      breakdown_date: operationsDate,
      description: descriptionParts.join(" — "),
      component: String(r.breakdown_component || "").trim() || null,
      critical: Boolean(r.breakdown_critical),
      time_down: `${operationsDate}T${r.short_time_down}`,
      time_up: `${operationsDate}T${r.short_time_up}`,
    }),
  });

  r.short_breakdown_id = Number(result?.breakdown_id || 0) || null;
  r.work_order_id = Number(result?.work_order_id || 0) || null;
  return { breakdownId: r.short_breakdown_id, workOrderId: r.work_order_id };
}

function renderDailyTable() {
  const body = qs("dailyBody");
  if (!body) return;

  body.innerHTML = "";
  const rowsToRender = dailyShowDownOnly ? dailyRows.filter((r) => !!r.is_down) : dailyRows;

  // Create a container for card-style rows
  const container = document.createElement("div");
  container.className = "daily-rows-container";

  for (const r of rowsToRender) {
    const rowClass = [
      "daily-row-card",
      r.is_master_standby ? "standby" : "",
      r.is_down ? "down" : "",
      r.error ? "has-error" : "",
      r.warning ? "has-warning" : ""
    ].filter(Boolean).join(" ");

    const card = document.createElement("div");
    card.className = rowClass;

    // Asset Header
    const assetHeader = document.createElement("div");
    assetHeader.className = "daily-row-header";
    assetHeader.innerHTML = `
      <div class="daily-asset-info">
        <span class="daily-asset-code">${r.asset_code}</span>
        <button class="daily-btn-icon daily-asset-qr" title="Download QR PNG for ${r.asset_code}">QR</button>
        <span class="daily-asset-name">${r.asset_name || ""}</span>
        ${r.telematics_locked ? '<span class="pill blue" style="margin-left:8px">TELEMATICS</span>' : ""}
      </div>
      <div class="daily-status-badge ${r.is_down ? 'status-down' : r.error ? 'status-error' : r.warning ? 'status-warning' : 'status-ok'}">
        ${r.is_down ? '⬇ DOWN' : r.error ? '✕ ' + r.error : r.warning ? '⚠ ' + r.warning : r.is_used ? '▶ PRODUCTION' : '⏸ STANDBY'}
      </div>
    `;
    card.appendChild(assetHeader);

    // Main Content Grid
    const contentGrid = document.createElement("div");
    contentGrid.className = "daily-row-content";

    // Left Column - Hours
    const hoursCol = document.createElement("div");
    hoursCol.className = "daily-hours-col";
    hoursCol.innerHTML = `
      <div class="daily-hours-grid">
        <div class="daily-hour-field">
          <label>Scheduled</label>
          <input type="number" step="0.5" min="0" max="24" value="${fmt(r.scheduled_hours)}" class="daily-input" />
        </div>
        <div class="daily-hour-field">
          <label>Opening</label>
          <input type="number" step="0.1" value="${fmt(r.opening_hours)}" class="daily-input readonly" readonly title="${r.opening_from_date ? `Auto-filled from ${r.opening_from_date}` : 'Auto-filled from yesterday'}" />
        </div>
        <div class="daily-hour-field">
          <label>Closing</label>
          <input type="number" step="0.1" value="${fmt(r.closing_hours)}" class="daily-input ${r.telematics_locked ? 'disabled readonly' : ''}" ${r.telematics_locked ? 'readonly' : ''} title="${r.telematics_locked ? 'Hourmeter from FSC telematics' : 'Enter the actual closing meter even when this machine had downtime'}" />
        </div>
        <div class="daily-hour-field">
          <label>Run</label>
          <div class="daily-run-display">${fmt(r.hours_run)} ${String(r.input_unit || "hours").toLowerCase() === "km" ? "km" : "hrs"}</div>
        </div>
      </div>
    `;
    contentGrid.appendChild(hoursCol);

    // Right Column - Controls
    const controlsCol = document.createElement("div");
    controlsCol.className = "daily-controls-col";

    // One clear operating status replaces separate production/down toggles.
    const statusField = document.createElement("div");
    statusField.className = "daily-status-field";
    const statusValue = r.is_down ? "maintenance" : r.is_used ? "production" : "standby";
    statusField.innerHTML = `
      <label>Machine status</label>
      <select class="daily-select daily-status-select status-${statusValue}" data-daily-status
        ${r.is_master_standby || (r.is_down && r.down_lock) ? "disabled" : ""}>
        <option value="production" ${statusValue === "production" ? "selected" : ""}>Production</option>
        <option value="standby" ${statusValue === "standby" ? "selected" : ""}>Standby</option>
        <option value="maintenance" ${statusValue === "maintenance" ? "selected" : ""}>Breakdown / repair</option>
      </select>
    `;
    controlsCol.appendChild(statusField);

    // Unit selector
    const unitSelect = document.createElement("div");
    unitSelect.className = "daily-unit-select";
    unitSelect.innerHTML = `
      <select class="daily-select">
        <option value="hours" ${String(r.input_unit || "hours").toLowerCase() === "hours" ? 'selected' : ''}>HRS</option>
        <option value="km" ${String(r.input_unit || "hours").toLowerCase() === "km" ? 'selected' : ''}>KM</option>
      </select>
      <button class="daily-btn-icon reset-unit" title="Reset to suggested (${r.suggested_input_unit || 'hours'})">↺</button>
    `;
    controlsCol.appendChild(unitSelect);

    contentGrid.appendChild(controlsCol);

    // Down Details Row (shown when down)
    if (r.is_down) {
      const downDetails = document.createElement("div");
      downDetails.className = "daily-down-details";
      
      const startYmd = String(r.breakdown_start_date || "").trim().match(/^(\d{4}-\d{2}-\d{2})/)?.[1] || "";
      const calcDaysDown = (() => {
        if (!startYmd) return null;
        const endYmd = String(qs("date")?.value || todayLocalYmd());
        const s = Date.parse(`${startYmd}T00:00:00`);
        const e = Date.parse(`${endYmd}T00:00:00`);
        if (!Number.isFinite(s) || !Number.isFinite(e) || e < s) return null;
        return Math.floor((e - s) / 86400000) + 1;
      })();

      downDetails.innerHTML = `
        <div class="daily-down-title">
          <strong>Daily incident update</strong>
          <span>One place for downtime, repair progress, parts, dates, and off-site tracking</span>
        </div>
        <div class="down-details-grid">
          <div class="down-field">
            <label>Incident handling</label>
            <select class="daily-select" data-down-field="incident_mode" ${r.down_lock ? "disabled" : ""}>
              <option value="ongoing" ${r.incident_mode !== "short" ? "selected" : ""}>Ongoing repair — keep WO open</option>
              <option value="short" ${r.incident_mode === "short" ? "selected" : ""}>Short breakdown — repaired in shift</option>
            </select>
          </div>
          <div class="down-field">
            <label>Reason</label>
            <select class="daily-select">
              <option value="">Select reason...</option>
              <option value="Mechanical" ${r.down_reason === "Mechanical" ? 'selected' : ''}>Mechanical</option>
              <option value="Electrical" ${r.down_reason === "Electrical" ? 'selected' : ''}>Electrical</option>
              <option value="Hydraulics" ${r.down_reason === "Hydraulics" ? 'selected' : ''}>Hydraulics</option>
              <option value="Tyres/Undercarriage" ${r.down_reason === "Tyres/Undercarriage" ? 'selected' : ''}>Tyres/Undercarriage</option>
              <option value="Waiting parts" ${r.down_reason === "Waiting parts" ? 'selected' : ''}>Waiting parts</option>
              <option value="Planned maintenance" ${r.down_reason === "Planned maintenance" ? 'selected' : ''}>Planned maintenance</option>
              <option value="Service" ${r.down_reason === "Service" ? 'selected' : ''}>Service</option>
              <option value="No operator" ${r.down_reason === "No operator" ? 'selected' : ''}>No operator</option>
              <option value="Weather/Access" ${r.down_reason === "Weather/Access" ? 'selected' : ''}>Weather/Access</option>
            </select>
          </div>
          <div class="down-field">
            <label>Down Hours</label>
            <input type="number" step="0.5" min="0" max="24" value="${fmt(r.down_hours != null ? r.down_hours : Number(r.scheduled_hours || 0))}" class="daily-input" placeholder="0" />
          </div>
          <div class="down-field">
            <label>Component / Area</label>
            <input type="text" value="${escapeHtml(String(r.breakdown_component || ""))}" class="daily-input" placeholder="Engine, hydraulics, tyres..." data-down-field="component" />
          </div>
          <div class="down-field">
            <label>Priority</label>
            <select class="daily-select" data-down-field="critical">
              <option value="normal" ${!r.breakdown_critical ? "selected" : ""}>Normal</option>
              <option value="critical" ${r.breakdown_critical ? "selected" : ""}>Critical</option>
            </select>
          </div>
          <div class="down-field">
            <label>Date down</label>
            <input type="date" value="${startYmd}" class="daily-input" data-down-field="start" />
          </div>
          <div class="down-field">
            <label>Date parts ordered</label>
            <input type="date" value="${String(r.parts_ordered_date || "").trim().match(/^(\d{4}-\d{2}-\d{2})/)?.[1] || ""}" class="daily-input" data-down-field="parts_ordered" />
          </div>
          <div class="down-field">
            <label>Status of parts</label>
            <select class="daily-select" data-down-field="parts_status">
              <option value="">—</option>
              <option value="Not ordered" ${r.parts_status === "Not ordered" ? "selected" : ""}>Not ordered</option>
              <option value="Ordered" ${r.parts_status === "Ordered" ? "selected" : ""}>Ordered</option>
              <option value="In transit" ${r.parts_status === "In transit" ? "selected" : ""}>In transit</option>
              <option value="Partial" ${r.parts_status === "Partial" ? "selected" : ""}>Partial</option>
              <option value="Received" ${r.parts_status === "Received" ? "selected" : ""}>Received</option>
              <option value="Waiting OEM" ${r.parts_status === "Waiting OEM" ? "selected" : ""}>Waiting OEM</option>
            </select>
          </div>
          <div class="down-field">
            <label>Received date</label>
            <input type="date" value="${String(r.parts_received_date || "").trim().match(/^(\d{4}-\d{2}-\d{2})/)?.[1] || ""}" class="daily-input" data-down-field="parts_received" />
          </div>
          <div class="down-field">
            <label>Target repair date</label>
            <input type="date" value="${String(r.ets_repair_date || "").trim().match(/^(\d{4}-\d{2}-\d{2})/)?.[1] || ""}" class="daily-input" data-down-field="ets_repair" />
          </div>
          <div class="down-field full">
            <label>Fault / action taken</label>
            <input type="text" value="${escapeHtml(String(r.breakdown_comment || ""))}" class="daily-input" placeholder="Describe the fault, inspection, repair or action required..." data-down-field="comment" />
          </div>
        </div>
        ${r.incident_mode === "short" ? `
          <div class="daily-short-breakdown-panel">
            <div class="daily-progress-heading">
              <strong>Short breakdown timing</strong>
              <span>Closes the incident and its work order when this Daily Log is saved.</span>
            </div>
            <div class="down-details-grid">
              <div class="down-field">
                <label>Time down</label>
                <input type="time" value="${escapeHtml(String(r.short_time_down || ""))}" class="daily-input" data-down-field="short_time_down" />
              </div>
              <div class="down-field">
                <label>Time up</label>
                <input type="time" value="${escapeHtml(String(r.short_time_up || ""))}" class="daily-input" data-down-field="short_time_up" />
              </div>
              <div class="down-field">
                <label>Calculated downtime</label>
                <div class="daily-run-display">${fmt(dailyShortBreakdownDurationHours(qs("date")?.value || todayLocalYmd(), r.short_time_down, r.short_time_up))} hrs</div>
              </div>
            </div>
          </div>` : ""}
        <div class="daily-progress-panel">
          <div class="daily-progress-heading">
            <strong>Repair progress and next plan</strong>
            <span>${r.work_order_id ? `Linked WO #${r.work_order_id}` : "A linked work order will be created when you save"}</span>
          </div>
          <div class="down-details-grid">
            <div class="down-field full">
              <label>Progress / next action</label>
              <textarea rows="2" class="daily-input" placeholder="What was completed today, what is waiting, and the next action..." data-down-field="repair_progress">${escapeHtml(String(r.repair_progress || ""))}</textarea>
            </div>
            <div class="down-field full daily-offsite-toggle-field">
              <label class="daily-offsite-toggle">
                <input type="checkbox" data-down-field="offsite_enabled" ${r.offsite_enabled ? "checked" : ""} />
                <span>Asset is off site for repair</span>
              </label>
              <span class="muted small">Use this only when the machine or component has left site.</span>
            </div>
          </div>
          ${r.offsite_enabled ? `
            <div class="daily-offsite-panel">
              <div class="daily-progress-heading">
                <strong>Off-site repair tracking</strong>
                <span>Saved to the linked off-site repair record</span>
              </div>
              <div class="down-details-grid">
                <div class="down-field">
                  <label>Repair stage</label>
                  <select class="daily-select" data-down-field="offsite_status">
                    <option value="sent_offsite" ${r.offsite_status === "sent_offsite" ? "selected" : ""}>Sent off site</option>
                    <option value="diagnosis" ${r.offsite_status === "diagnosis" ? "selected" : ""}>Diagnosis</option>
                    <option value="in_repair" ${r.offsite_status === "in_repair" ? "selected" : ""}>In repair</option>
                    <option value="waiting_parts" ${r.offsite_status === "waiting_parts" ? "selected" : ""}>Waiting parts</option>
                    <option value="ready_return" ${r.offsite_status === "ready_return" ? "selected" : ""}>Ready to return</option>
                    <option value="returned" ${r.offsite_status === "returned" ? "selected" : ""}>Returned</option>
                  </select>
                </div>
                <div class="down-field">
                  <label>Sent date</label>
                  <input type="date" value="${String(r.offsite_sent_date || "").trim().match(/^(\d{4}-\d{2}-\d{2})/)?.[1] || ""}" class="daily-input" data-down-field="offsite_sent" />
                </div>
                <div class="down-field">
                  <label>Expected return</label>
                  <input type="date" value="${String(r.offsite_expected_return_date || "").trim().match(/^(\d{4}-\d{2}-\d{2})/)?.[1] || ""}" class="daily-input" data-down-field="offsite_expected" />
                </div>
                <div class="down-field">
                  <label>Supplier / repairer</label>
                  <input type="text" value="${escapeHtml(String(r.offsite_vendor || ""))}" class="daily-input" placeholder="Workshop or supplier" data-down-field="offsite_vendor" />
                </div>
                <div class="down-field">
                  <label>Current location</label>
                  <input type="text" value="${escapeHtml(String(r.offsite_location || ""))}" class="daily-input" placeholder="Workshop, city, depot..." data-down-field="offsite_location" />
                </div>
                <div class="down-field">
                  <label>Actual return</label>
                  <input type="date" value="${String(r.offsite_actual_return_date || "").trim().match(/^(\d{4}-\d{2}-\d{2})/)?.[1] || ""}" class="daily-input" data-down-field="offsite_actual" />
                </div>
                <div class="down-field full">
                  <label>Off-site progress / next action</label>
                  <textarea rows="2" class="daily-input" placeholder="Supplier update, parts status, collection plan..." data-down-field="offsite_progress">${escapeHtml(String(r.offsite_progress || ""))}</textarea>
                </div>
              </div>
            </div>` : ""}
        </div>
        ${startYmd ? `<div class="down-meta">Date down: ${startYmd}${calcDaysDown != null ? ` | Total days down: ${calcDaysDown}` : ""}</div>` : ""}
        ${r.is_down && r.down_lock ? `<div class="down-lock-notice">Linked WO is ${escapeHtml(String(r.lock_wo_status || "open").replace(/_/g, " "))}. Meter inputs and the planning section remain editable here.</div>` : ""}
      `;
      contentGrid.appendChild(downDetails);
    }

    // Footer with opening source
    if (r.opening_from_date) {
      const footer = document.createElement("div");
      footer.className = "daily-row-footer";
      footer.innerHTML = `<span class="opening-source">Opening from: ${r.opening_from_date}</span>`;
      contentGrid.appendChild(footer);
    }

    card.appendChild(contentGrid);
    container.appendChild(card);

    // Add event listeners
    const inputs = card.querySelectorAll("input, select, textarea");

    inputs.forEach(input => {
      input.addEventListener("change", () => {
        const type = input.type;
        const cls = input.className;

        if (type === "number") {
          if (cls.includes("readonly")) {
            r.opening_hours = toNum(input.value);
          } else if (input.closest(".down-field") && input.previousElementSibling?.textContent === "Down Hours") {
            r.down_hours = toNum(input.value) ?? 0;
          } else if (input.previousElementSibling?.textContent === "Scheduled") {
            r.scheduled_hours = toNum(input.value) ?? 0;
          } else if (input.previousElementSibling?.textContent === "Closing") {
            r.closing_hours = toNum(input.value);
          }
        } else if (type === "date") {
          const downField = String(input.getAttribute("data-down-field") || "").trim();
          if (downField === "start" || input.previousElementSibling?.textContent === "Date down" || input.previousElementSibling?.textContent === "Started") {
            r.breakdown_start_date = String(input.value || "").trim();
          } else if (downField === "parts_ordered") {
            r.parts_ordered_date = String(input.value || "").trim();
          } else if (downField === "parts_received") {
            r.parts_received_date = String(input.value || "").trim();
          } else if (downField === "ets_repair") {
            r.ets_repair_date = String(input.value || "").trim();
          } else if (downField === "offsite_sent") {
            r.offsite_sent_date = String(input.value || "").trim();
          } else if (downField === "offsite_expected") {
            r.offsite_expected_return_date = String(input.value || "").trim();
          } else if (downField === "offsite_actual") {
            r.offsite_actual_return_date = String(input.value || "").trim();
          }
        } else if (type === "time") {
          if (input.getAttribute("data-down-field") === "short_time_down") {
            r.short_time_down = String(input.value || "").trim();
          } else if (input.getAttribute("data-down-field") === "short_time_up") {
            r.short_time_up = String(input.value || "").trim();
          }
          updateDailyShortBreakdownHours(qs("date")?.value || todayLocalYmd(), r);
        } else if (type === "text") {
          if (input.getAttribute("data-down-field") === "component") {
            r.breakdown_component = String(input.value || "").trim();
          } else if (input.getAttribute("data-down-field") === "comment" || input.previousElementSibling?.textContent === "Comment") {
            r.breakdown_comment = String(input.value || "").trim();
          } else if (input.getAttribute("data-down-field") === "offsite_vendor") {
            r.offsite_vendor = String(input.value || "").trim();
          } else if (input.getAttribute("data-down-field") === "offsite_location") {
            r.offsite_location = String(input.value || "").trim();
          }
        } else if (type === "checkbox") {
          if (input.getAttribute("data-down-field") === "offsite_enabled") {
            r.offsite_enabled = Boolean(input.checked);
            if (r.offsite_enabled && !r.offsite_sent_date) {
              r.offsite_sent_date = dailyOperationsDate(qs("date")?.value || todayLocalYmd());
            }
          }
        } else if (input.tagName === "TEXTAREA") {
          if (input.getAttribute("data-down-field") === "repair_progress") {
            r.repair_progress = String(input.value || "").trim();
          } else if (input.getAttribute("data-down-field") === "offsite_progress") {
            r.offsite_progress = String(input.value || "").trim();
          }
        } else if (input.tagName === "SELECT") {
          if (input.hasAttribute("data-daily-status")) {
            const nextStatus = String(input.value || "standby");
            r.is_down = nextStatus === "maintenance";
            r.is_used = nextStatus !== "standby";
            if (r.is_down) {
              if (r.down_hours == null) r.down_hours = Number(r.scheduled_hours || 10);
              if (!r.breakdown_start_date) r.breakdown_start_date = dailyOperationsDate(qs("date")?.value || todayLocalYmd());
            } else {
              r.down_reason = "";
              r.down_hours = null;
              r.breakdown_comment = "";
              r.breakdown_component = "";
              r.breakdown_critical = false;
              r.breakdown_start_date = "";
              r.parts_ordered_date = "";
              r.parts_status = "";
              r.parts_received_date = "";
              r.ets_repair_date = "";
              r.incident_mode = "ongoing";
              r.short_time_down = "";
              r.short_time_up = "";
              r.short_breakdown_id = null;
              r.repair_progress = "";
              r.offsite_enabled = false;
              r.offsite_status = "sent_offsite";
              r.offsite_sent_date = "";
              r.offsite_expected_return_date = "";
              r.offsite_actual_return_date = "";
              r.offsite_vendor = "";
              r.offsite_location = "";
              r.offsite_progress = "";
            }
          } else if (cls.includes("daily-select") && !cls.includes("disabled")) {
            if (input.getAttribute("data-down-field") === "parts_status") {
              r.parts_status = String(input.value || "").trim();
            } else if (input.getAttribute("data-down-field") === "critical") {
              r.breakdown_critical = input.value === "critical";
            } else if (input.getAttribute("data-down-field") === "incident_mode") {
              r.incident_mode = input.value === "short" ? "short" : "ongoing";
              if (r.incident_mode === "short") {
                r.offsite_enabled = false;
                r.offsite_status = "sent_offsite";
                r.offsite_sent_date = "";
                r.offsite_expected_return_date = "";
                r.offsite_actual_return_date = "";
                r.offsite_vendor = "";
                r.offsite_location = "";
                r.offsite_progress = "";
              }
            } else if (input.getAttribute("data-down-field") === "offsite_status") {
              r.offsite_status = String(input.value || "sent_offsite").trim();
            } else if (input.closest(".down-field")) {
              r.down_reason = String(input.value || "");
            } else if (input.value === "hours" || input.value === "km") {
              r.input_unit = input.value;
            }
          }
        }

        validateDailyRows();
        renderDailyTable();
        renderDailyPreview();
      });
    });

    // Reset unit button
    const resetBtn = unitSelect.querySelector(".reset-unit");
    resetBtn?.addEventListener("click", () => {
      const suggested = String(r.suggested_input_unit || "hours").toLowerCase() === "km" ? "km" : "hours";
      r.input_unit = suggested;
      validateDailyRows();
      renderDailyTable();
      renderDailyPreview();
    });

    const assetQrBtn = assetHeader.querySelector(".daily-asset-qr");
    assetQrBtn?.addEventListener("click", (ev) => {
      ev.preventDefault();
      downloadAssetQrPng(r.asset_code).catch((e) => setStatus("QR download error: " + e.message));
    });
  }

  body.appendChild(container);
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
    const openData = await fetchJson(`${API}/api/breakdowns/open-all?date=${encodeURIComponent(dailyOperationsDate(date))}`);
    const rows = Array.isArray(openData?.rows) ? openData.rows : [];
    for (const bd of rows) {
      const code = String(bd.asset_code || "").trim();
      if (!code) continue;
      if (!openBreakdownByAsset.has(code)) openBreakdownByAsset.set(code, bd);
    }
  } catch {
    openBreakdownByAsset = new Map();
  }

  let openOffsiteByBreakdown = new Map();
  let openOffsiteByAsset = new Map();
  try {
    const offsiteData = await fetchJson(`${API}/api/breakdown-ops/offsite-repairs?include_closed=0`);
    const rows = Array.isArray(offsiteData?.rows) ? offsiteData.rows : [];
    for (const repair of rows) {
      const breakdownId = Number(repair.breakdown_id || 0);
      const assetCode = String(repair.asset_code || "").trim();
      if (breakdownId && !openOffsiteByBreakdown.has(breakdownId)) {
        openOffsiteByBreakdown.set(breakdownId, repair);
      }
      if (assetCode && !openOffsiteByAsset.has(assetCode)) {
        openOffsiteByAsset.set(assetCode, repair);
      }
    }
  } catch {
    openOffsiteByBreakdown = new Map();
    openOffsiteByAsset = new Map();
  }

  try {
    const ps = await fetchJson(`${API}/api/maintenance/prestart/daily-summary?date=${encodeURIComponent(date)}`);
    dailyPrestartRows = Array.isArray(ps?.rows) ? ps.rows : [];
    dailyPrestartMeta = {
      deduction_hours_per_check: Number(ps?.deduction_hours_per_check ?? 0.5),
      production_deduction_hours: Number(ps?.production_deduction?.hours ?? 0),
      production_deduction_count: Number(ps?.production_deduction?.count ?? 0),
    };
  } catch {
    dailyPrestartRows = [];
    dailyPrestartMeta = {
      deduction_hours_per_check: 0.5,
      production_deduction_hours: 0,
      production_deduction_count: 0,
    };
  }

  const parseDownReasonFromDesc = (desc) => {
    const d = String(desc || "").trim();
    if (!d) return "";
    const m = d.match(/(?:DOWN|BREAKDOWN)\s*[-—:]\s*(.+)$/i);
    if (m && m[1]) return String(m[1]).trim();
    const m2 = d.match(/^(?:DOWN|BREAKDOWN)\s*(.+)$/i);
    if (m2 && m2[1]) return String(m2[1]).trim();
    // Fallback for descriptions like: "BREAKDOWN — Hydraulics"
    const m3 = d.match(/^(?:DOWN|BREAKDOWN)\s*[^A-Za-z0-9]*\s*(.+)$/i);
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

      is_used: ex
        ? !!ex.is_used
        : (yByCode.has(a.asset_code) ? !!yByCode.get(a.asset_code)?.is_used : !masterStandby),
      input_unit: ex?.input_unit
        ? String(ex.input_unit).toLowerCase()
        : (String(a.category || "").toLowerCase().includes("truck") || String(a.category || "").toLowerCase().includes("vehicle") ? "km" : "hours"),
      suggested_input_unit: ex?.input_unit
        ? String(ex.input_unit).toLowerCase()
        : (String(a.category || "").toLowerCase().includes("truck") || String(a.category || "").toLowerCase().includes("vehicle") ? "km" : "hours"),
      input_unit_locked: Boolean(ex?.input_unit_locked) || Boolean(ex?.telematics_locked),
      telematics_locked: Boolean(ex?.telematics_locked),
      meter_source: ex?.meter_source || "manual",

      scheduled_hours: ex ? toNum(ex.scheduled_hours) : null,
      opening_hours: ex ? toNum(ex.opening_hours) : null,
      opening_from_date: null,
      closing_hours: ex ? toNum(ex.closing_hours) : null,
      hours_run: ex ? toNum(ex.hours_run) ?? 0 : 0,

      is_down: false,
      down_reason: "",
      down_hours: null,
      down_lock: false,
      breakdown_start_date: "",
      breakdown_comment: ex?.notes ? String(ex.notes) : "",
      breakdown_component: "",
      breakdown_critical: false,
      parts_ordered_date: "",
      parts_status: "",
      parts_received_date: "",
      ets_repair_date: "",
      incident_mode: "ongoing",
      short_time_down: "",
      short_time_up: "",
      short_breakdown_id: null,
      breakdown_id: null,
      work_order_id: null,
      repair_progress: "",
      repair_due_date: "",
      offsite_enabled: false,
      offsite_repair_id: null,
      offsite_status: "sent_offsite",
      offsite_sent_date: "",
      offsite_expected_return_date: "",
      offsite_actual_return_date: "",
      offsite_vendor: "",
      offsite_location: "",
      offsite_progress: "",

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
        if (!ex && typeof d?.suggested_is_used === "boolean") row.is_used = Boolean(d.suggested_is_used);
        if (typeof d?.input_unit_locked === "boolean") row.input_unit_locked = d.input_unit_locked;
        if ((row.opening_hours == null || forceOpenFromYesterday) && d.suggested_opening_hours != null) {
          row.opening_hours = Number(d.suggested_opening_hours);
          row.opening_from_date = String(d.suggested_opening_from_date || "").trim() || null;
        }
        if (row.scheduled_hours == null && d.suggested_scheduled_hours != null) row.scheduled_hours = Number(d.suggested_scheduled_hours);
      } catch {}
    }

    // Opening must match previous closing — refresh when prior day was corrected after save
    if (!row.telematics_locked && !row.is_master_standby && row.opening_hours != null) {
      let expectedOpen = null;
      let fromDate = null;
      const yClose = yr ? toNum(yr.closing_hours) : null;
      if (yClose != null) {
        expectedOpen = yClose;
        fromDate = y;
      } else {
        try {
          const d = await fetchJson(
            `${API}/api/hours/defaults?asset_code=${encodeURIComponent(row.asset_code)}&work_date=${date}`
          );
          if (d?.suggested_opening_hours != null) {
            expectedOpen = Number(d.suggested_opening_hours);
            fromDate = String(d.suggested_opening_from_date || "").trim() || null;
          }
        } catch {}
      }
      if (
        expectedOpen != null &&
        Math.abs(Number(row.opening_hours) - expectedOpen) > 0.0001
      ) {
        row.opening_hours = expectedOpen;
        row.opening_from_date = fromDate;
        if (!row.warning) row.warning = "Opening updated from corrected previous closing";
      }
    }

    if (row.scheduled_hours == null) row.scheduled_hours = row.is_master_standby ? 0 : getDayScheduledHours();
    if (row.is_master_standby) row.is_used = false;
    row.hours_run = calcRun(row.opening_hours, row.closing_hours);

    // Apply carry-forward lock from open breakdown
    const bd = openBreakdownByAsset.get(row.asset_code);
    if (bd) {
      row.is_down = true;
      row.down_lock = true;
      row.down_reason = parseDownReasonFromDesc(bd.description);
      row.lock_wo_status = bd.primary_work_order_status || "";
      row.breakdown_start_date = String(bd.breakdown_date || bd.start_at || "").trim();
      row.breakdown_component = String(bd.component || "").trim();
      row.breakdown_critical = Boolean(bd.critical);
      row.parts_ordered_date = String(bd.parts_ordered_date || "").trim();
      row.parts_status = String(bd.parts_status || "").trim();
      row.parts_received_date = String(bd.parts_received_date || "").trim();
      row.ets_repair_date = String(bd.ets_repair_date || "").trim();
      row.breakdown_id = Number(bd.id || 0) || null;
      row.work_order_id = Number(bd.primary_work_order_id || 0) || null;
      row.incident_mode = "ongoing";
      row.repair_progress = String(bd.repair_progress || "").trim();
      row.repair_due_date = String(bd.repair_due_date || "").trim();
      if (!row.ets_repair_date && row.repair_due_date) {
        row.ets_repair_date = row.repair_due_date;
      }
      const downForDate = Number(bd.hours_down_for_date);
      row.down_hours = Number.isFinite(downForDate) && downForDate >= 0
        ? downForDate
        : Number(row.scheduled_hours || 0);
    }

    const offsite = row.breakdown_id
      ? (openOffsiteByBreakdown.get(Number(row.breakdown_id)) || openOffsiteByAsset.get(row.asset_code))
      : null;
    if (offsite) {
      row.offsite_enabled = true;
      row.offsite_repair_id = Number(offsite.id || 0) || null;
      row.offsite_status = String(offsite.repair_status || "sent_offsite").trim();
      row.offsite_sent_date = String(offsite.sent_date || "").trim();
      row.offsite_expected_return_date = String(offsite.expected_return_date || "").trim();
      row.offsite_actual_return_date = String(offsite.actual_return_date || "").trim();
      row.offsite_vendor = String(offsite.vendor || "").trim();
      row.offsite_location = String(offsite.current_location || "").trim();
      row.offsite_progress = String(offsite.notes || "").trim();
      if (!row.ets_repair_date && row.offsite_expected_return_date) {
        row.ets_repair_date = row.offsite_expected_return_date;
      }
    }

    if (row.is_down && !String(row.breakdown_start_date || "").trim()) {
      row.breakdown_start_date = dailyOperationsDate(date);
    }

    if (dailyScheduledOverride != null && !row.is_master_standby && (row.is_used || row.is_down)) {
      row.scheduled_hours = dailyScheduledOverride;
      if (row.is_down && row.down_hours == null) row.down_hours = dailyScheduledOverride;
    }

    dailyRows.push(row);
  }

  validateDailyRows();
  renderDailyPrestartSection();
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
    const payload = {
      asset_code: r.asset_code,
      work_date: date,
      is_used: r.is_down ? Number(r.hours_run || 0) > 0 : r.is_used,
      input_unit: String(r.input_unit || "hours").toLowerCase(),
      scheduled_hours: r.scheduled_hours ?? 0,
      opening_hours: r.opening_hours,
      closing_hours: r.closing_hours,
      hours_run: r.hours_run,
      notes: r.is_down ? (String(r.breakdown_comment || "").trim() || null) : null,
    };

    try {
      const res = await postHoursWithOffline(payload);
      if (res && res.queued) results.push({ asset_code: r.asset_code, ok: true, queued: true });
      else results.push({ asset_code: r.asset_code, ok: true, res });
    } catch (e) {
      results.push({ asset_code: r.asset_code, ok: false, error: e.message || String(e) });
    }
  }

  const incidentFailures = [];
  for (const r of dailyRows) {
    if (!r.is_down) continue;
    try {
      if (r.incident_mode === "short") {
        await logDailyShortBreakdown(date, r);
      } else {
        await logDownRowToBreakdowns(date, r);
      }
    } catch (e) {
      incidentFailures.push({ asset_code: r.asset_code, error: e.message || String(e) });
    }
  }

  const failed = results.filter((x) => !x.ok);
  const queued = results.filter((x) => x.queued).length;

  setText("dailyResult", JSON.stringify({
    saved: results.length - failed.length,
    failed,
    incident_updates_failed: incidentFailures,
  }, null, 2));

  if (failed.length || incidentFailures.length) setStatus(`Saved with issues: ${failed.length + incidentFailures.length} failed.`);
  else if (queued) setStatus(`Saved offline: ${queued} queued for sync ✅`);
  else setStatus("Saved successfully.");

  refreshNetBanner();

  await loadDashboard().catch(() => {});
  await loadDailyInput().catch(() => {});
}

async function buildAssetQrImageData(assetCode) {
  const code = String(assetCode || "").trim();
  if (!code) throw new Error("Asset code is required.");
  const res = await fetchJson(`${API}/api/assets/${encodeURIComponent(code)}/qr-profile/refresh`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({}),
  });
  const qrText = String(res?.qr_text || "").trim();
  const scanValue = String(res?.qr_payload?.scan_url || qrText || "").trim();
  if (!scanValue) throw new Error("No QR value generated.");
  const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=420x420&data=${encodeURIComponent(scanValue)}`;
  return { qrUrl, qrText, scanValue, payload: res?.qr_payload || {} };
}

async function downloadAssetQrPng(assetCode) {
  const code = String(assetCode || "").trim();
  if (!code) {
    alert("Asset code missing.");
    return;
  }
  setStatus(`Preparing QR image for ${code}...`);
  const { qrUrl } = await buildAssetQrImageData(code);
  const response = await fetch(qrUrl);
  if (!response.ok) throw new Error(`QR image fetch failed (${response.status})`);
  const blob = await response.blob();
  const objUrl = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = objUrl;
  a.download = `${code}_qr.png`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(objUrl);
  setStatus(`QR image downloaded for ${code} ✅`);
}

async function downloadAllVisibleDailyQrs() {
  if (typeof window.JSZip !== "function") {
    alert("ZIP library failed to load. Check internet connection and try again.");
    return;
  }
  const rowsToUse = dailyShowDownOnly ? dailyRows.filter((r) => !!r.is_down) : dailyRows;
  const codes = Array.from(new Set(rowsToUse.map((r) => String(r.asset_code || "").trim()).filter(Boolean)));
  if (!codes.length) {
    alert("No visible assets to export.");
    return;
  }

  setStatus(`Preparing ${codes.length} visible QR images for ZIP...`);
  const zip = new window.JSZip();
  let ok = 0;
  let fail = 0;
  for (let i = 0; i < codes.length; i += 1) {
    const code = codes[i];
    try {
      const { qrUrl } = await buildAssetQrImageData(code);
      const response = await fetch(qrUrl);
      if (!response.ok) throw new Error(`fetch ${response.status}`);
      const blob = await response.blob();
      zip.file(`${code}_qr.png`, blob);
      ok += 1;
      setStatus(`Collecting QR ${i + 1}/${codes.length}: ${code}`);
      await new Promise((resolve) => setTimeout(resolve, 120));
    } catch {
      fail += 1;
    }
  }
  if (!ok) throw new Error("No QR images could be added to ZIP.");

  setStatus("Building ZIP file...");
  const zipBlob = await zip.generateAsync({ type: "blob" });
  const dt = new Date();
  const stamp = `${dt.getFullYear()}${String(dt.getMonth() + 1).padStart(2, "0")}${String(dt.getDate()).padStart(2, "0")}_${String(dt.getHours()).padStart(2, "0")}${String(dt.getMinutes()).padStart(2, "0")}`;
  const objUrl = URL.createObjectURL(zipBlob);
  const a = document.createElement("a");
  a.href = objUrl;
  a.download = `ironlog_visible_qr_${stamp}.zip`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(objUrl);
  setStatus(`QR ZIP ready: ${ok} included${fail ? `, ${fail} failed` : ""} ✅`);
}

async function printVisibleDailyQrSheet() {
  const rowsToUse = dailyShowDownOnly ? dailyRows.filter((r) => !!r.is_down) : dailyRows;
  const codes = Array.from(new Set(rowsToUse.map((r) => String(r.asset_code || "").trim()).filter(Boolean)));
  if (!codes.length) {
    alert("No visible assets to print.");
    return;
  }

  setStatus(`Building printable QR sheet for ${codes.length} asset(s)...`);
  const labels = [];
  for (let i = 0; i < codes.length; i += 1) {
    const code = codes[i];
    try {
      const { qrUrl } = await buildAssetQrImageData(code);
      labels.push({ code, qrUrl });
      setStatus(`Preparing label ${i + 1}/${codes.length}: ${code}`);
      await new Promise((resolve) => setTimeout(resolve, 80));
    } catch {
      // skip bad rows but continue
    }
  }

  if (!labels.length) throw new Error("Could not prepare any QR labels.");

  openQrLabelSheetPrintWindow(labels, readQrSheetLayout("daily"), "IRONLOG QR Label Sheet");
  setStatus(`Printable QR sheet ready (${labels.length} labels) ✅`);
}

const QR_SHEET_PRESETS = {
  small: { cols: 5, qr: 22, cell: 32, gap: 2 },
  medium: { cols: 4, qr: 28, cell: 45, gap: 4 },
  large: { cols: 3, qr: 35, cell: 55, gap: 5 },
  vinyl_45up: {
    cols: 5,
    qr: 22,
    cell: 29.9,
    gap: 2,
    cellWidthMm: 38.5,
    rowGapMm: 0,
    pageMarginTopMm: 13.95,
    pageMarginSideMm: 4.75,
    exactSheet: true,
  },
  avery_3474: { cols: 4, qr: 23, cell: 34, gap: 2 },
  avery_l7163: { cols: 2, qr: 35, cell: 43, gap: 4 },
};

function qrSheetFieldIds(scope) {
  if (scope === "safety") {
    return {
      preset: "safetyQrPreset",
      cols: "safetyQrCols",
      qr: "safetyQrSizeMm",
      cell: "safetyQrCellMm",
      gap: "safetyQrGapMm",
    };
  }
  return {
    preset: "qrPreset",
    cols: "qrCols",
    qr: "qrSizeMm",
    cell: "qrCellMm",
    gap: "qrGapMm",
  };
}

function readQrSheetLayout(scope) {
  const ids = qrSheetFieldIds(scope);
  const preset = String(qs(ids.preset)?.value || "custom");
  const presetLayout = QR_SHEET_PRESETS[preset] || {};
  const defaults = scope === "safety"
    ? { cols: 5, qr: 22, cell: 32, gap: 2 }
    : { cols: 4, qr: 28, cell: 45, gap: 4 };
  return {
    preset,
    cols: Math.max(1, Math.min(8, Number(toNum(qs(ids.cols)?.value) ?? defaults.cols))),
    qrSizeMm: Math.max(12, Math.min(60, Number(toNum(qs(ids.qr)?.value) ?? defaults.qr))),
    cellMm: Math.max(20, Math.min(80, Number(toNum(qs(ids.cell)?.value) ?? defaults.cell))),
    gapMm: Math.max(0, Math.min(20, Number(toNum(qs(ids.gap)?.value) ?? defaults.gap))),
    cellWidthMm: Number(presetLayout.cellWidthMm) || null,
    rowGapMm: Number.isFinite(Number(presetLayout.rowGapMm)) ? Number(presetLayout.rowGapMm) : null,
    pageMarginTopMm: Number(presetLayout.pageMarginTopMm) || null,
    pageMarginSideMm: Number(presetLayout.pageMarginSideMm) || null,
    exactSheet: presetLayout.exactSheet === true,
  };
}

function openQrLabelSheetPrintWindow(labels, layout, sheetTitle) {
  const safeLabels = Array.isArray(labels) ? labels.filter((l) => l?.qrUrl && l?.code) : [];
  if (!safeLabels.length) throw new Error("No QR labels to print.");
  const {
    cols,
    qrSizeMm,
    cellMm,
    gapMm,
    cellWidthMm,
    rowGapMm,
    pageMarginTopMm,
    pageMarginSideMm,
    exactSheet,
  } = layout || readQrSheetLayout("daily");
  const verticalGapMm = rowGapMm == null ? Math.max(1, Math.round(gapMm * 1.2)) : rowGapMm;
  const pageMargin = exactSheet
    ? `${pageMarginTopMm}mm ${pageMarginSideMm}mm`
    : "8mm";
  const columnTemplate = cellWidthMm
    ? `repeat(${cols}, ${cellWidthMm}mm)`
    : `repeat(${cols}, 1fr)`;
  const win = window.open("", "_blank", "width=1100,height=800");
  if (!win) {
    alert("Pop-up blocked. Allow pop-ups and try again.");
    return;
  }
  const cells = safeLabels
    .map(
      (l) => `
      <div class="cell">
        <img src="${l.qrUrl}" alt="${escapeHtml(l.code)} QR" />
        <div class="code">${escapeHtml(l.code)}</div>
      </div>
    `
    )
    .join("");
  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>${escapeHtml(sheetTitle || "IRONLOG QR Label Sheet")}</title>
  <style>
    @page { size: A4 portrait; margin: ${pageMargin}; }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: Arial, sans-serif; color: #111; }
    .sheet { padding: ${exactSheet ? "0" : "8mm"}; }
    .head { margin-bottom: 6mm; font-size: 12px; ${exactSheet ? "display:none;" : ""} }
    .grid {
      display: grid;
      grid-template-columns: ${columnTemplate};
      gap: ${verticalGapMm}mm ${gapMm}mm;
    }
    .cell {
      border: ${exactSheet ? "0" : "1px solid #bbb"};
      border-radius: ${exactSheet ? "0" : "4px"};
      padding: ${exactSheet ? "1.2mm 1mm 0.7mm" : "3mm 2mm"};
      text-align: center;
      min-height: ${cellMm}mm;
      height: ${exactSheet ? `${cellMm}mm` : "auto"};
      break-inside: avoid;
    }
    .cell img {
      width: ${qrSizeMm}mm;
      height: ${qrSizeMm}mm;
      image-rendering: pixelated;
      display: block;
      margin: 0 auto ${exactSheet ? "0.5mm" : "2mm"};
    }
    .code {
      font-size: ${exactSheet ? "9px" : "10px"};
      font-weight: 700;
      letter-spacing: 0.2px;
      overflow-wrap: anywhere;
    }
    @media print {
      .no-print { display: none; }
      body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    }
  </style>
</head>
<body>
  <div class="sheet">
    <div class="head">${escapeHtml(sheetTitle || "IRONLOG QR Label Sheet")} | Total: ${safeLabels.length} | Layout: ${cols} cols, QR ${qrSizeMm}mm, Cell ${cellMm}mm, Gap ${gapMm}mm | Generated: ${new Date().toISOString()}</div>
    <div class="grid">${cells}</div>
    <div class="no-print" style="margin-top:10px;font-size:12px;color:#555;">Use browser print scaling at 100% for label alignment.</div>
  </div>
  <script>window.onload = () => { window.focus(); window.print(); };</script>
</body>
</html>`;
  win.document.open();
  win.document.write(html);
  win.document.close();
}

function applyQrSheetPresetScope(scope) {
  const ids = qrSheetFieldIds(scope);
  const preset = String(qs(ids.preset)?.value || "custom");
  const target = QR_SHEET_PRESETS[preset];
  if (!target) return;
  const setVal = (id, value) => {
    const el = qs(id);
    if (el) el.value = String(value);
  };
  setVal(ids.cols, target.cols);
  setVal(ids.qr, target.qr);
  setVal(ids.cell, target.cell);
  setVal(ids.gap, target.gap);
}

function applyQrSheetPreset() {
  applyQrSheetPresetScope("daily");
}

function applySafetyQrSheetPreset() {
  applyQrSheetPresetScope("safety");
}

async function generateDailyAssetQr() {
  const assetCode = String(qs("dailyQrAssetCode")?.value || "").trim();
  if (!assetCode) {
    alert("Enter/select an asset code first.");
    return;
  }
  setStatus(`Generating QR for ${assetCode}...`);
  const { qrUrl, qrText, payload } = await buildAssetQrImageData(assetCode);
  const img = qs("dailyQrImg");
  if (img) img.src = qrUrl;
  const preview = qs("dailyQrPreview");
  if (preview) preview.style.display = "block";

  const service = payload?.next_service_due;
  const out = {
    asset_code: payload?.asset?.asset_code || assetCode,
    status: payload?.status || "UNKNOWN",
    next_service_due: service
      ? `${service.service_name} @ ${service.next_due_hours}h (${service.remaining_hours}h remaining)`
      : "No active maintenance plan",
    fuel_liters_last_30_days: payload?.fuel?.liters_last_30_days ?? 0,
    last_inspection_date: payload?.inspections?.last_inspection_date || null,
    generated_at: payload?.generated_at || new Date().toISOString(),
    qr_text: qrText,
  };
  setText("dailyQrText", JSON.stringify(out, null, 2));
  setStatus(`QR saved for ${assetCode} ✅`);
}

function printDailyAssetQr() {
  const imgSrc = String(qs("dailyQrImg")?.src || "").trim();
  const payloadText = String(qs("dailyQrText")?.textContent || "").trim();
  if (!imgSrc || !payloadText) {
    alert("Generate a QR first, then print.");
    return;
  }

  let payload = null;
  try {
    payload = JSON.parse(payloadText);
  } catch {
    payload = null;
  }
  if (!payload) {
    alert("QR payload is invalid. Generate the QR again.");
    return;
  }

  const machine = String(payload.asset_code || "Unknown");
  const status = String(payload.status || "UNKNOWN");
  const nextService = String(payload.next_service_due || "No active maintenance plan");
  const fuel = `${Number(payload.fuel_liters_last_30_days || 0).toFixed(1)} L (30 days)`;
  const inspection = String(payload.last_inspection_date || "No inspection date");
  const generated = String(payload.generated_at || "");

  const win = window.open("", "_blank", "width=900,height=700");
  if (!win) {
    alert("Pop-up blocked. Allow pop-ups and try again.");
    return;
  }

  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>IRONLOG QR - ${machine}</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; color: #111; }
    .sheet { border: 1px solid #222; border-radius: 10px; padding: 18px; max-width: 760px; }
    h1 { margin: 0 0 10px; font-size: 22px; }
    .meta { margin: 0 0 14px; font-size: 14px; }
    .grid { display: grid; grid-template-columns: 240px 1fr; gap: 16px; align-items: start; }
    img { width: 220px; height: 220px; border: 1px solid #999; }
    .row { margin: 0 0 8px; font-size: 14px; }
    .label { font-weight: 700; }
    .raw { margin-top: 14px; padding: 10px; border: 1px dashed #999; white-space: pre-wrap; font-size: 12px; }
    @media print {
      body { margin: 0; }
      .sheet { border: 0; border-radius: 0; padding: 8mm; max-width: none; }
    }
  </style>
</head>
<body>
  <div class="sheet">
    <h1>IRONLOG Machine QR</h1>
    <div class="meta">Generated: ${generated || new Date().toISOString()}</div>
    <div class="grid">
      <div><img src="${imgSrc}" alt="Machine QR code" /></div>
      <div>
        <div class="row"><span class="label">Machine:</span> ${machine}</div>
        <div class="row"><span class="label">Status:</span> ${status}</div>
        <div class="row"><span class="label">Next service:</span> ${nextService}</div>
        <div class="row"><span class="label">Fuel used:</span> ${fuel}</div>
        <div class="row"><span class="label">Last inspection:</span> ${inspection}</div>
      </div>
    </div>
    <div class="raw">${String(payload.qr_text || "").replace(/</g, "&lt;").replace(/>/g, "&gt;")}</div>
  </div>
  <script>window.onload = () => { window.focus(); window.print(); };</script>
</body>
</html>`;

  win.document.open();
  win.document.write(html);
  win.document.close();
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
