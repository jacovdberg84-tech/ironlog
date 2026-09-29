// IRONLOG/web/maintenance/insights.js — Maintenance insights, forecasts and governance signals.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

function insightsRowsTable(headers, rows) {
  if (!Array.isArray(rows) || !rows.length) return `<div class="muted">No data in selected range.</div>`;
  const head = `<tr>${headers.map((h) => `<th>${escBackfill(h)}</th>`).join("")}</tr>`;
  const body = rows
    .map((r) => `<tr>${r.map((v) => {
      if (v && typeof v === "object" && typeof v.__html === "string") return `<td>${v.__html}</td>`;
      return `<td>${escBackfill(v == null ? "-" : String(v))}</td>`;
    }).join("")}</tr>`)
    .join("");
  return `<div style="overflow:auto;"><table class="gridTable" style="min-width:640px;"><thead>${head}</thead><tbody>${body}</tbody></table></div>`;
}

function renderStandardizedInsightDashboards(data) {
  const host = document.getElementById("insightsStandardDashboards");
  if (!host) return;

  const costRows = Array.isArray(data?.maintenance_cost) ? data.maintenance_cost : [];
  const downtimeRows = Array.isArray(data?.downtime?.by_component) ? data.downtime.by_component : [];
  const partRows = Array.isArray(data?.parts_planning?.suggestions) ? data.parts_planning.suggestions : [];

  const totalMaintenanceCost = costRows.reduce((sum, r) => sum + Number(r?.total_cost || 0), 0);
  const totalPartsCost = costRows.reduce((sum, r) => sum + Number(r?.parts_cost || 0), 0);
  const totalDowntimeHours = Number(data?.downtime?.total_hours || 0)
    || downtimeRows.reduce((sum, r) => sum + Number(r?.downtime_hours || 0), 0);
  const laborTotals = data?.downtime?.labor?.totals || {};
  const repairLaborHours = Number(laborTotals.repair_labor_hours || 0);
  const repairLaborCost = Number(laborTotals.repair_labor_cost || 0);
  const needsLaborInputCount = Number(laborTotals.needs_input_count || 0);
  const totalDowntimeIncidents = downtimeRows.reduce((sum, r) => sum + Number(r?.incidents || 0), 0);
  const totalSuggestedPartsQty = partRows.reduce((sum, r) => sum + Number(r?.suggested_qty || 0), 0);
  const totalPartsGapQty = partRows.reduce((sum, r) => sum + Number(r?.gap_qty || 0), 0);
  const totalUpcomingCost = Number(data?.parts_planning?.total_upcoming_cost || 0);
  const needsManualCount = Number(data?.parts_planning?.needs_manual_input_count || 0);

  const topCostAsset = costRows[0];
  const topDowntimeComponent = downtimeRows[0];
  const topPartsDemand = partRows[0];

  const gapSeverityClass = totalSuggestedPartsQty <= 0
    ? "kpi-good"
    : ((totalPartsGapQty / totalSuggestedPartsQty) > 0.25 ? "kpi-bad" : "kpi-warn");
  const gapPct = totalSuggestedPartsQty > 0
    ? Math.min(100, Math.max(0, (totalPartsGapQty / totalSuggestedPartsQty) * 100))
    : 0;

  host.innerHTML = `
    <div class="kpi-card kpi-util">
      <div class="kpi-card-header">
        <div class="kpi-icon">R</div>
        <div class="kpi-title">Maintenance Cost Dashboard</div>
      </div>
      <div class="kpi-big-value">${fmtMoney(totalMaintenanceCost)}</div>
      <div class="kpi-meta">Total maintenance spend (${Number(costRows.length || 0)} assets)</div>
      <div class="kpi-meta" style="margin-top:6px;">Labor, parts, lube, and outsourced — excludes fuel (running cost)</div>
      <div class="kpi-meta" style="margin-top:6px;">Parts consumption cost: ${fmtMoney(totalPartsCost)}</div>
      <div class="kpi-meta" style="margin-top:6px;">
        Top cost asset:
        ${topCostAsset ? `${escBackfill(topCostAsset.asset_code || "-")} (${fmtMoney(topCostAsset.total_cost)})` : "N/A"}
      </div>
    </div>
    <div class="kpi-card kpi-alerts">
      <div class="kpi-card-header">
        <div class="kpi-icon">H</div>
        <div class="kpi-title">Downtime Dashboard</div>
      </div>
      <div class="kpi-big-value">${fmt1(totalDowntimeHours)}</div>
      <div class="kpi-meta">Total machine downtime hours (${Number(totalDowntimeIncidents || 0)} incidents)</div>
      <div class="kpi-meta" style="margin-top:6px;">Actual repair labor: ${fmt1(repairLaborHours)} hrs (${fmtMoney(repairLaborCost)})</div>
      ${needsLaborInputCount > 0
        ? `<div class="kpi-meta message-error" style="margin-top:6px;">${needsLaborInputCount} breakdown(s) need repair labor hours</div>`
        : ""}
      <div class="kpi-meta" style="margin-top:6px;">
        Top component:
        ${topDowntimeComponent ? `${escBackfill(topDowntimeComponent.component || "-")} (${fmt1(topDowntimeComponent.downtime_hours)}h)` : "N/A"}
      </div>
      <div class="kpi-meta" style="margin-top:6px;">Standardized view: hours, incidents, and top contributor</div>
    </div>
    <div class="kpi-card kpi-avail">
      <div class="kpi-card-header">
        <div class="kpi-icon">P</div>
        <div class="kpi-title">Parts Consumption Dashboard</div>
      </div>
      <div class="kpi-big-value">${fmt1(totalSuggestedPartsQty)}</div>
      <div class="kpi-meta">Suggested parts quantity (${Number(partRows.length || 0)} SKU suggestions)</div>
      <div class="kpi-meta" style="margin-top:6px;">Upcoming service forecast: ${fmtMoney(totalUpcomingCost)}</div>
      ${needsManualCount > 0 ? `<div class="kpi-meta" style="margin-top:6px;">${needsManualCount} service(s) need manual cost input</div>` : ""}
      <div class="kpi-progress">
        <div class="kpi-progress-bar ${gapSeverityClass}" style="width:${gapPct.toFixed(1)}%;"></div>
      </div>
      <div class="kpi-meta">Stock gap: ${fmt1(totalPartsGapQty)} (${gapPct.toFixed(1)}%)</div>
      <div class="kpi-meta" style="margin-top:6px;">
        Top suggested part:
        ${topPartsDemand ? `${escBackfill(topPartsDemand.part_name || "-")} (${fmt1(topPartsDemand.suggested_qty)})` : "N/A"}
      </div>
    </div>
  `;
}

function populateInsightsPlanSelect(forecastRows) {
  const sel = document.getElementById("insightsInputPlan");
  if (!sel) return;
  const rows = Array.isArray(forecastRows) ? forecastRows : [];
  sel.innerHTML = `<option value="">Select plan...</option>${rows.map((r) => {
    const pid = Number(r.plan_id || 0);
    const label = `${r.asset_code || "-"} | ${r.service_name || "-"} (${Number(r.remaining_hours || 0).toFixed(0)}h)`;
    return `<option value="${pid}">${escBackfill(label)}</option>`;
  }).join("")}`;
}

function getInsightsPartByCode(code) {
  return getWfPartByCode(code);
}

function refreshInsightsDraftEditor() {
  const body = document.getElementById("insightsDraftBody");
  if (!body) return;
  body.innerHTML = insightsDraftItems.length
    ? insightsRowsTable(
      ["Type", "Part", "Qty", "Unit $", "Line $", ""],
      insightsDraftItems.map((it, idx) => [
        it.type || "part",
        `${it.part_code || "-"}${it.part_name ? ` - ${it.part_name}` : ""}`,
        Number(it.qty || 0).toFixed(2),
        fmtMoney(it.unit_cost || 0),
        fmtMoney(it.line_cost || 0),
        { __html: `<button type="button" data-insights-remove-item="${idx}">Remove</button>` },
      ]),
    )
    : `<small class="muted">No manual parts/oil lines yet.</small>`;
}

function hydrateInsightsDraftFromSaved(planId) {
  const row = insightsInputsCache.find((r) => Number(r.plan_id || 0) === Number(planId || 0));
  const notesEl = document.getElementById("insightsInputNotes");
  const laborEl = document.getElementById("insightsInputLabor");
  if (!row) {
    insightsDraftItems = [];
    if (notesEl) notesEl.value = "";
    if (laborEl) laborEl.value = "0";
    refreshInsightsDraftEditor();
    return;
  }
  if (notesEl) notesEl.value = String(row.notes || "");
  if (laborEl) laborEl.value = String(Number(row.labor_total || 0));
  let items = [];
  try {
    const parsed = JSON.parse(String(row.items_json || "[]"));
    if (Array.isArray(parsed)) items = parsed;
  } catch {}
  insightsDraftItems = items.map((it) => {
    const part = getInsightsPartByCode(it.part_code);
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
  refreshInsightsDraftEditor();
}

function addInsightsDraftItem() {
  const msg = document.getElementById("insightsInputMsg");
  const type = String(document.getElementById("insightsItemType")?.value || "part").toLowerCase() === "oil" ? "oil" : "part";
  const part_code = String(document.getElementById("insightsItemCode")?.value || "").trim();
  const qty = Math.max(0, Number(document.getElementById("insightsItemQty")?.value || 0));
  if (!part_code || qty <= 0) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Select a part code and enter a quantity greater than 0.";
    }
    return;
  }
  const part = getInsightsPartByCode(part_code);
  if (!part) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Part code not found in Stores list.";
    }
    return;
  }
  insightsDraftItems.push({
    type,
    part_code: String(part.part_code || "").trim(),
    part_name: String(part.part_name || ""),
    qty,
    unit_cost: Number(part.latest_unit_cost || 0),
    on_hand: Number(part.on_hand || 0),
    line_cost: qty * Number(part.latest_unit_cost || 0),
  });
  const codeEl = document.getElementById("insightsItemCode");
  const qtyEl = document.getElementById("insightsItemQty");
  if (codeEl) codeEl.value = "";
  if (qtyEl) qtyEl.value = "0";
  if (msg) {
    msg.className = "muted";
    msg.textContent = "Item added.";
  }
  refreshInsightsDraftEditor();
}

let insightsManualUiBound = false;
let insightsRepairLaborUiBound = false;
let insightsRepairLaborCache = [];
function bindInsightsManualForecastUi() {
  const host = document.getElementById("insightsParts");
  if (!host || insightsManualUiBound) return;
  insightsManualUiBound = true;
  host.addEventListener("change", (evt) => {
    if (evt.target?.id === "insightsInputPlan") {
      hydrateInsightsDraftFromSaved(Number(evt.target?.value || 0));
    }
  });
  host.addEventListener("click", (evt) => {
    if (evt.target?.closest?.("#insightsAddItemBtn")) {
      addInsightsDraftItem();
      return;
    }
    if (evt.target?.closest?.("#insightsSaveInputBtn")) {
      saveInsightsForecastInput();
      return;
    }
    const btn = evt.target?.closest?.("button[data-insights-remove-item]");
    if (!btn) return;
    const idx = Number(btn.getAttribute("data-insights-remove-item") || -1);
    if (idx < 0 || idx >= insightsDraftItems.length) return;
    insightsDraftItems.splice(idx, 1);
    refreshInsightsDraftEditor();
  });
}

async function loadInsightsPartsCatalog() {
  if (wfPartsCache.length) {
    refreshInsightsPartsDatalist();
    return;
  }
  await loadWeeklyForumParts();
  refreshInsightsPartsDatalist();
}

function refreshInsightsPartsDatalist() {
  const list = document.getElementById("insightsPartsList");
  if (!list) return;
  list.innerHTML = wfPartsCache.map((p) => {
    const code = String(p.part_code || "").trim();
    const name = String(p.part_name || "").trim();
    return `<option value="${escBackfill(code)}">${escBackfill(name)}</option>`;
  }).join("");
}

async function loadInsightsForecastInputs() {
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/forecast-inputs`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load saved templates");
    insightsInputsCache = Array.isArray(data.rows) ? data.rows : [];
    const planNow = Number(document.getElementById("insightsInputPlan")?.value || 0);
    if (planNow) hydrateInsightsDraftFromSaved(planNow);
  } catch {
    insightsInputsCache = [];
  }
}

function bindInsightsRepairLaborUi() {
  const host = document.getElementById("insightsDowntime");
  if (!host || insightsRepairLaborUiBound) return;
  insightsRepairLaborUiBound = true;
  host.addEventListener("change", (evt) => {
    if (evt.target?.id !== "insightsRepairLaborBreakdown") return;
    const bid = Number(evt.target?.value || 0);
    const inc = insightsRepairLaborCache.find((r) => Number(r.breakdown_id || 0) === bid);
    const laborEl = document.getElementById("insightsRepairLaborHours");
    const notesEl = document.getElementById("insightsRepairLaborNotes");
    if (laborEl) laborEl.value = inc ? Number(inc.actual_labor_hours || 0) : 0;
    if (notesEl) notesEl.value = inc ? String(inc.labor_notes || "") : "";
  });
  host.addEventListener("click", (evt) => {
    if (!evt.target?.closest?.("#insightsSaveRepairLaborBtn")) return;
    saveBreakdownRepairLabor();
  });
}

async function saveBreakdownRepairLabor() {
  const msg = document.getElementById("insightsRepairLaborMsg");
  const breakdown_id = Number(document.getElementById("insightsRepairLaborBreakdown")?.value || 0);
  const labor_hours = Math.max(0, Number(document.getElementById("insightsRepairLaborHours")?.value || 0));
  const notes = String(document.getElementById("insightsRepairLaborNotes")?.value || "").trim();
  if (!msg) return;
  if (!breakdown_id) {
    msg.className = "message-error";
    msg.textContent = "Select a breakdown incident first.";
    return;
  }
  if (labor_hours <= 0) {
    msg.className = "message-error";
    msg.textContent = "Enter actual repair labor hours (technician time, not machine downtime).";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Saving repair labor...";
  try {
    const res = await fetch(`${API}/maintenance/breakdowns/${breakdown_id}/repair-labor`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ labor_hours, notes: notes || null }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save repair labor");
    msg.className = "message-success";
    msg.textContent = `Saved ${Number(data.labor_hours || labor_hours).toFixed(2)} repair labor hrs for ${data.asset_code || "breakdown"}.`;
    await loadMaintenanceInsights();
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = e.message || String(e);
  }
}

async function saveInsightsForecastInput() {
  const msg = document.getElementById("insightsInputMsg");
  const plan_id = Number(document.getElementById("insightsInputPlan")?.value || 0);
  const notes = String(document.getElementById("insightsInputNotes")?.value || "").trim();
  const labor_total = Math.max(0, Number(document.getElementById("insightsInputLabor")?.value || 0));
  const items = insightsDraftItems.map((x) => ({
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
  if (!items.length && labor_total <= 0) {
    msg.className = "message-error";
    msg.textContent = "Add at least one part/oil line or enter a labor total.";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Saving service cost template...";
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/forecast-inputs`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ plan_id, items, labor_total, notes: notes || null }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save template");
    msg.className = "message-success";
    msg.textContent = "Template saved. Refresh insights to update forecast totals.";
    await loadInsightsForecastInputs();
    await loadMaintenanceInsights();
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = e.message || String(e);
  }
}

function renderMaintenanceInsights(data) {
  const predictiveEl = document.getElementById("insightsPredictive");
  const partsEl = document.getElementById("insightsParts");
  const downtimeEl = document.getElementById("insightsDowntime");
  const slaEl = document.getElementById("insightsSla");
  const costEl = document.getElementById("insightsCost");
  const trendsEl = document.getElementById("insightsTrends");
  if (!predictiveEl || !partsEl || !downtimeEl || !slaEl || !costEl || !trendsEl) return;
  renderStandardizedInsightDashboards(data);

  const riskRows = Array.isArray(data?.predictive?.at_risk_plans) ? data.predictive.at_risk_plans : [];
  const failRows = Array.isArray(data?.predictive?.repeated_checklist_failures) ? data.predictive.repeated_checklist_failures : [];
  const fuelRows = Array.isArray(data?.predictive?.fuel_anomalies) ? data.predictive.fuel_anomalies : [];
  predictiveEl.innerHTML = `
    <div class="muted">At-risk plans: ${riskRows.length} | Repeat checklist failures: ${failRows.length} | Fuel anomalies: ${fuelRows.length}</div>
    ${insightsRowsTable(
      ["Asset", "Service", "Remaining Hrs", "Risk"],
      riskRows.slice(0, 10).map((r) => [
        {
          __html: `${r.asset_code ? `<button type="button" data-insights-asset-code="${escBackfill(r.asset_code)}">${escBackfill(r.asset_code)}</button>` : "-"}${r.asset_name ? ` - ${escBackfill(r.asset_name)}` : ""}`,
        },
        r.service_name || "-",
        Number(r.remaining_hours || 0).toFixed(1),
        r.risk || "-",
      ])
    )}
  `;

  const partRows = Array.isArray(data?.parts_planning?.suggestions) ? data.parts_planning.suggestions : [];
  const forecastRows = Array.isArray(data?.parts_planning?.upcoming_cost_forecasts) ? data.parts_planning.upcoming_cost_forecasts : [];
  insightsForecastCache = forecastRows;
  partsEl.innerHTML = `
    <div class="muted">
      Upcoming services in horizon: ${Number(data?.parts_planning?.upcoming_service_count || 0)}
      | Forecast total: ${fmtMoney(data?.parts_planning?.total_upcoming_cost || 0)}
      ${Number(data?.parts_planning?.needs_manual_input_count || 0) > 0
        ? ` | <span class="message-error">${Number(data.parts_planning.needs_manual_input_count)} need manual cost input</span>`
        : ""}
    </div>
    <h5 style="margin:12px 0 6px;">Upcoming maintenance cost by equipment</h5>
    ${insightsRowsTable(
      ["Asset", "Service", "Remaining Hrs", "Status", "Kit $", "Labor $", "Total $", "Source"],
      forecastRows.slice(0, 15).map((r) => [
        `${r.asset_code || "-"}${r.asset_name ? ` - ${r.asset_name}` : ""}`,
        r.service_name || "-",
        Number(r.remaining_hours || 0).toFixed(1),
        r.status || "-",
        fmtMoney(r?.forecast?.est_service_kit_cost || 0),
        fmtMoney(r?.forecast?.est_labor_cost || 0),
        fmtMoney(r?.forecast?.est_total_cost || 0),
        String(r?.forecast?.cost_source || "-").replace(/_/g, " "),
      ])
    )}
    <h5 style="margin:14px 0 6px;">Parts demand (from prior service history)</h5>
    ${insightsRowsTable(
      ["Part", "Suggested Qty", "Est Cost", "On Hand", "Gap", "Linked Services"],
      partRows.slice(0, 12).map((r) => [
        r.part_name || "-",
        Number(r.suggested_qty || 0).toFixed(1),
        fmtMoney(r.est_cost || 0),
        Number(r.on_hand || 0).toFixed(1),
        Number(r.gap_qty || 0).toFixed(1),
        Array.isArray(r.linked_services) ? r.linked_services.join(", ") : "-",
      ])
    )}
    <hr class="hr-soft" style="margin:14px 0;" />
    <h5 style="margin:0 0 8px;">Manual service cost template (saved for next events)</h5>
    <small class="muted" style="display:block; margin-bottom:8px;">
      Use when no closed service history exists for a plan. Parts priced from Stores; labor is a flat total for the job.
    </small>
    <div class="row stack-10">
      <label>
        Upcoming service plan
        <select id="insightsInputPlan"></select>
      </label>
      <label>
        Total labor ($)
        <input id="insightsInputLabor" type="number" min="0" step="0.01" value="0" />
      </label>
    </div>
    <div class="row stack-10">
      <label>
        Type
        <select id="insightsItemType">
          <option value="oil">Oil</option>
          <option value="part">Part</option>
        </select>
      </label>
      <label style="flex:1;">
        Store part code
        <input id="insightsItemCode" list="insightsPartsList" placeholder="Search part code or name" />
      </label>
      <label>
        Qty
        <input id="insightsItemQty" type="number" min="0" step="0.1" value="0" />
      </label>
      <button type="button" id="insightsAddItemBtn">Add item</button>
    </div>
    <datalist id="insightsPartsList"></datalist>
    <div id="insightsDraftBody" style="margin-top:8px;"></div>
    <div class="row stack-10" style="margin-top:8px;">
      <label style="flex:1;">
        Notes
        <input id="insightsInputNotes" placeholder="Optional notes for this service template" />
      </label>
      <button type="button" id="insightsSaveInputBtn">Save template</button>
    </div>
    <div id="insightsInputMsg" class="muted" style="margin-top:6px;"></div>
  `;
  populateInsightsPlanSelect(forecastRows);
  bindInsightsManualForecastUi();
  loadInsightsForecastInputs().catch(() => {});
  loadInsightsPartsCatalog().catch(() => {});

  const compRows = Array.isArray(data?.downtime?.by_component) ? data.downtime.by_component : [];
  const teamRows = Array.isArray(data?.downtime?.by_team) ? data.downtime.by_team : [];
  const laborIncidents = Array.isArray(data?.downtime?.labor?.incidents) ? data.downtime.labor.incidents : [];
  const laborTotals = data?.downtime?.labor?.totals || {};
  insightsRepairLaborCache = laborIncidents;
  const laborRate = Number(data?.range?.labor_cost_per_hour || 0);
  const needsLaborCount = Number(laborTotals.needs_input_count || 0);
  const laborIncidentOptions = laborIncidents.map((r) => {
    const label = `#${r.breakdown_id} ${r.asset_code || "-"} — down ${Number(r.downtime_hours || 0).toFixed(1)}h`;
    return `<option value="${Number(r.breakdown_id || 0)}">${escBackfill(label)}</option>`;
  }).join("");
  downtimeEl.innerHTML = `
    <div class="muted">
      Machine down: ${Number(laborTotals.downtime_hours || 0).toFixed(2)} hrs
      | Actual repair labor: ${Number(laborTotals.repair_labor_hours || 0).toFixed(2)} hrs (${fmtMoney(laborTotals.repair_labor_cost || 0)})
      ${needsLaborCount > 0 ? ` | <span class="message-error">${needsLaborCount} need labor input</span>` : ""}
    </div>
    <small class="muted" style="display:block; margin:6px 0 10px;">
      Downtime is machine-not-available time. Repair labor is technician hours actually spent — they are not the same (e.g. 200h down may be 12h repair labor).
    </small>
    <h5 style="margin:0 0 6px;">Breakdown downtime vs repair labor</h5>
    ${insightsRowsTable(
      ["Asset", "Breakdown", "Down Hrs", "Repair Hrs", "Repair $", "Source", "Status"],
      laborIncidents.slice(0, 15).map((r) => [
        `${r.asset_code || "-"}${r.asset_name ? ` - ${r.asset_name}` : ""}`,
        `#${r.breakdown_id}${r.description ? ` — ${String(r.description).slice(0, 40)}` : ""}`,
        Number(r.downtime_hours || 0).toFixed(2),
        r.needs_labor_input
          ? { __html: `<span class="message-error">${Number(r.actual_labor_hours || 0).toFixed(2)}</span>` }
          : Number(r.actual_labor_hours || 0).toFixed(2),
        fmtMoney(r.repair_labor_cost || 0),
        String(r.labor_source || "-").replace(/_/g, " "),
        r.status || "-",
      ])
    )}
    <hr class="hr-soft" style="margin:14px 0;" />
    <h5 style="margin:0 0 8px;">Enter actual repair labor hours</h5>
    <div class="row stack-10">
      <label style="flex:1;">
        Breakdown incident
        <select id="insightsRepairLaborBreakdown">
          <option value="">Select breakdown...</option>
          ${laborIncidentOptions}
        </select>
      </label>
      <label>
        Repair labor (hrs)
        <input id="insightsRepairLaborHours" type="number" min="0" step="0.25" value="0" />
      </label>
      <label style="flex:1;">
        Notes
        <input id="insightsRepairLaborNotes" placeholder="Optional — e.g. 2 techs × 6 hrs" />
      </label>
      <button type="button" id="insightsSaveRepairLaborBtn">Save labor</button>
    </div>
    <div id="insightsRepairLaborMsg" class="muted" style="margin-top:6px;">
      ${laborRate > 0 ? `Default labor rate when WO rate missing: ${fmtMoney(laborRate)}/hr.` : ""}
    </div>
    <h5 style="margin:14px 0 6px;">By component</h5>
    ${insightsRowsTable(
      ["Component", "Incidents", "Downtime Hrs"],
      compRows.slice(0, 8).map((r) => [r.component || "-", Number(r.incidents || 0), Number(r.downtime_hours || 0).toFixed(2)])
    )}
    <div style="margin-top:10px;"></div>
    <h5 style="margin:0 0 6px;">By team</h5>
    ${insightsRowsTable(
      ["Team", "Incidents", "Downtime Hrs"],
      teamRows.slice(0, 8).map((r) => [r.team || "-", Number(r.incidents || 0), Number(r.downtime_hours || 0).toFixed(2)])
    )}
  `;
  bindInsightsRepairLaborUi();

  const s = data?.sla || {};
  slaEl.innerHTML = insightsRowsTable(
    ["Metric", "Hours"],
    [
      ["Open -> Assign", s.avg_open_to_assign_hours == null ? "-" : Number(s.avg_open_to_assign_hours).toFixed(2)],
      ["Open -> Complete", s.avg_open_to_complete_hours == null ? "-" : Number(s.avg_open_to_complete_hours).toFixed(2)],
      ["Complete -> Approve", s.avg_complete_to_approve_hours == null ? "-" : Number(s.avg_complete_to_approve_hours).toFixed(2)],
      ["Open -> Close", s.avg_open_to_close_hours == null ? "-" : Number(s.avg_open_to_close_hours).toFixed(2)],
      ["Service Work Orders", Number(s.work_orders || 0)],
    ]
  );

  const costRows = Array.isArray(data?.maintenance_cost) ? data.maintenance_cost : [];
  const costLaborRate = Number(data?.range?.labor_cost_per_hour || 0);
  const laborHint = costLaborRate > 0
    ? `<div class="muted mini" style="margin:0 0 8px;">Labor = service WO hours + actual breakdown repair hours (not machine downtime). Default repair rate ${fmtMoney(costLaborRate)}/hr when WO rate is missing. Fuel excluded.</div>`
    : `<div class="muted mini" style="margin:0 0 8px;">Labor = service WO hours + actual breakdown repair hours (not machine downtime). Fuel excluded.</div>`;
  costEl.innerHTML = `${laborHint}${insightsRowsTable(
    ["Asset", "Jobs", "Down Hrs", "Repair Hrs", "WO Labor", "Repair Labor", "Parts", "Lube", "Total"],
    costRows.slice(0, 15).map((r) => [
      `${r.asset_code || "-"} - ${r.asset_name || "-"}`,
      Number(r.service_jobs || 0),
      Number(r.downtime_hours || 0).toFixed(2),
      Number(r.repair_labor_hours || 0).toFixed(2),
      fmtMoney(r.wo_labor_cost || 0),
      fmtMoney(r.repair_labor_cost || 0),
      fmtMoney(r.parts_cost || 0),
      fmtMoney(r.lube_cost || 0),
      fmtMoney(r.total_cost || 0),
    ])
  )}`;

  const downtimeDaily = Array.isArray(data?.trends?.downtime_daily) ? data.trends.downtime_daily : [];
  const laborDaily = Array.isArray(data?.trends?.labor_daily) ? data.trends.labor_daily : [];
  const toMiniBars = (title, rows, valueKey, color) => {
    const latest = rows.slice(-14);
    const maxVal = Math.max(1, ...latest.map((r) => Number(r?.[valueKey] || 0)));
    const bars = latest
      .map((r) => {
        const day = String(r?.day || "-").slice(5);
        const val = Number(r?.[valueKey] || 0);
        const pct = Math.max(4, Math.min(100, (val / maxVal) * 100));
        return `<div style="display:grid; grid-template-columns:56px 1fr 60px; align-items:center; gap:8px; margin:4px 0;">
          <small class="muted">${escBackfill(day)}</small>
          <div style="background:#e5e7eb; border-radius:6px; height:10px; overflow:hidden;">
            <div style="width:${pct.toFixed(1)}%; height:100%; background:${color};"></div>
          </div>
          <small>${escBackfill(val.toFixed(2))}</small>
        </div>`;
      })
      .join("");
    return `<div class="card" style="padding:10px;"><div style="font-weight:600; margin-bottom:6px;">${escBackfill(title)}</div>${bars || `<small class="muted">No trend data.</small>`}</div>`;
  };
  trendsEl.innerHTML = `
    <div class="form-grid">
      ${toMiniBars("Downtime Hours (Last 14 days)", downtimeDaily, "downtime_hours", "#dc2626")}
      ${toMiniBars("Labor Cost (Last 14 days)", laborDaily, "labor_cost", "#2563eb")}
    </div>
  `;
}

function renderGovernanceSignals(data) {
  const msgEl = document.getElementById("governanceSignalsMsg");
  const scoreEl = document.getElementById("insightsDataQualityScore");
  const qualityEl = document.getElementById("governanceQuality");
  const anomaliesEl = document.getElementById("governanceAnomalies");
  if (!msgEl || !qualityEl || !anomaliesEl) return;
  const q = data?.quality || {};
  const a = data?.anomalies || {};
  const missing = Array.isArray(q.missing_meter_readings) ? q.missing_meter_readings : [];
  const inconsistent = Array.isArray(q.inconsistent_statuses) ? q.inconsistent_statuses : [];
  const stale = Array.isArray(q.stale_plans) ? q.stale_plans : [];
  const spikes = Array.isArray(a.fuel_spikes) ? a.fuel_spikes : [];
  const duplicates = Array.isArray(a.fuel_duplicates) ? a.fuel_duplicates : [];
  const jumps = Array.isArray(a.suspicious_meter_jumps) ? a.suspicious_meter_jumps : [];
  const issueScore = (
    (missing.length * 3)
    + (inconsistent.length * 2)
    + (stale.length * 2)
    + (spikes.length * 1.5)
    + (duplicates.length * 1.5)
    + (jumps.length * 2)
  );
  const score = Math.max(0, Math.min(100, Math.round(100 - issueScore)));
  const scoreClass = score >= 85 ? "kpi-good" : (score >= 65 ? "kpi-warn" : "kpi-bad");
  if (scoreEl) {
    scoreEl.innerHTML = `
      <div class="kpi-card" style="padding:14px; border-radius:12px;">
        <div class="kpi-card-header" style="margin-bottom:8px;">
          <div class="kpi-icon">Q</div>
          <div class="kpi-title">Data Quality Score</div>
        </div>
        <div class="kpi-big-value" style="font-size:34px;">${score}</div>
        <div class="kpi-progress">
          <div class="kpi-progress-bar ${scoreClass}" style="width:${score}%;"></div>
        </div>
        <div class="kpi-meta">
          Calculated from missing meter readings, stale plans, inconsistent statuses, and anomaly signals.
        </div>
      </div>
    `;
  }
  qualityEl.innerHTML = `
    <div class="muted">Missing meter: ${missing.length} | Inconsistent WO status: ${inconsistent.length} | Stale plans: ${stale.length}</div>
    ${insightsRowsTable(
      ["Asset", "Issue", "Detail"],
      [
        ...missing.slice(0, 10).map((r) => [`${r.asset_code || "-"} - ${r.asset_name || "-"}`, "Missing meter", `Current ${Number(r.current_hours || 0).toFixed(1)} (${r.source || "-"})`]),
        ...inconsistent.slice(0, 10).map((r) => [`WO #${Number(r.work_order_id || 0)} (${r.asset_code || "-"})`, "Inconsistent status", `${r.status || "-"} | completed_at=${r.completed_at || "-"} | closed_at=${r.closed_at || "-"}`]),
        ...stale.slice(0, 10).map((r) => [`${r.asset_code || "-"} - ${r.asset_name || "-"}`, "Stale plan", `${r.service_name || "-"} | remaining ${Number(r.remaining_hours || 0).toFixed(1)}h`]),
      ]
    )}
  `;
  anomaliesEl.innerHTML = `
    <div class="muted">Fuel spikes: ${spikes.length} | Fuel duplicates: ${duplicates.length} | Meter jumps: ${jumps.length}</div>
    ${insightsRowsTable(
      ["Signal", "Asset", "Detail"],
      [
        ...spikes.slice(0, 10).map((r) => ["Fuel spike", `${r.asset_code || "-"} - ${r.asset_name || "-"}`, `${r.log_date || "-"} | ${Number(r.liters || 0).toFixed(1)}L vs avg ${Number(r.avg_liters || 0).toFixed(1)}L (${Number(r.spike_ratio || 0).toFixed(2)}x)`]),
        ...duplicates.slice(0, 10).map((r) => ["Duplicate fuel", `${r.asset_code || "-"} - ${r.asset_name || "-"}`, `${r.log_date || "-"} | ${Number(r.liters || 0).toFixed(1)}L | count ${Number(r.duplicate_count || 0)}`]),
        ...jumps.slice(0, 10).map((r) => ["Meter jump", `${r.asset_code || "-"} - ${r.asset_name || "-"}`, `${r.input_date || "-"} | ${Number(r.previous_meter || 0).toFixed(1)} -> ${Number(r.current_meter || 0).toFixed(1)} (jump ${Number(r.jump || 0).toFixed(1)})`]),
      ]
    )}
  `;
}

function getInsightsThresholds() {
  try {
    const raw = JSON.parse(String(localStorage.getItem(INSIGHTS_THRESHOLDS_KEY) || "{}"));
    return {
      near_due_hours: Math.max(1, Number(raw.near_due_hours || 50)),
      predictive_horizon_hours: Math.max(1, Number(raw.predictive_horizon_hours || 100)),
      checklist_fail_threshold: Math.max(1, Number(raw.checklist_fail_threshold || 2)),
      fuel_variance_threshold: Math.max(0, Number(raw.fuel_variance_threshold || 15)),
    };
  } catch {
    return {
      near_due_hours: 50,
      predictive_horizon_hours: 100,
      checklist_fail_threshold: 2,
      fuel_variance_threshold: 15,
    };
  }
}

function persistInsightsThresholds() {
  const nearEl = document.getElementById("insightsNearDueHours");
  const horizonEl = document.getElementById("insightsPredictiveHorizonHours");
  const failEl = document.getElementById("insightsChecklistFailThreshold");
  const fuelEl = document.getElementById("insightsFuelVarianceThreshold");
  const payload = {
    near_due_hours: Math.max(1, Number(nearEl?.value || 50)),
    predictive_horizon_hours: Math.max(1, Number(horizonEl?.value || 100)),
    checklist_fail_threshold: Math.max(1, Number(failEl?.value || 2)),
    fuel_variance_threshold: Math.max(0, Number(fuelEl?.value || 15)),
  };
  localStorage.setItem(INSIGHTS_THRESHOLDS_KEY, JSON.stringify(payload));
  return payload;
}

function buildMaintenanceInsightsExportQuery() {
  const start = String(document.getElementById("insightsStart")?.value || "").trim();
  const end = String(document.getElementById("insightsEnd")?.value || "").trim();
  if (!start || !end) throw new Error("Select insights start and end dates.");
  const t = persistInsightsThresholds();
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  q.set("near_due_hours", String(t.near_due_hours));
  q.set("predictive_horizon_hours", String(t.predictive_horizon_hours));
  q.set("checklist_fail_threshold", String(t.checklist_fail_threshold));
  q.set("fuel_variance_threshold", String(t.fuel_variance_threshold));
  return { start, end, q };
}

async function openMaintenanceInsightsXlsx() {
  const { start, end, q } = buildMaintenanceInsightsExportQuery();
  await downloadProtectedXlsxFile(
    `${API}/maintenance/insights.xlsx?${q.toString()}`,
    `IRONLOG_Maintenance_Insights_${start}_to_${end}.xlsx`,
  );
}

async function downloadInsightsPartsDemandXlsx() {
  const { start, end, q } = buildMaintenanceInsightsExportQuery();
  await downloadProtectedXlsxFile(
    `${API}/maintenance/insights/parts-demand.xlsx?${q.toString()}`,
    `IRONLOG_Parts_Demand_${start}_to_${end}.xlsx`,
  );
}

async function downloadInsightsCostPerMachineXlsx() {
  const { start, end, q } = buildMaintenanceInsightsExportQuery();
  await downloadProtectedXlsxFile(
    `${API}/maintenance/insights/cost-per-machine.xlsx?${q.toString()}`,
    `IRONLOG_Maintenance_Cost_Per_Machine_${start}_to_${end}.xlsx`,
  );
}

function openMaintenanceInsightsPdf(download = false) {
  const start = String(document.getElementById("insightsStart")?.value || "").trim();
  const end = String(document.getElementById("insightsEnd")?.value || "").trim();
  if (!start || !end) return alert("Select insights start and end dates.");
  const t = persistInsightsThresholds();
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  q.set("near_due_hours", String(t.near_due_hours));
  q.set("predictive_horizon_hours", String(t.predictive_horizon_hours));
  q.set("checklist_fail_threshold", String(t.checklist_fail_threshold));
  q.set("fuel_variance_threshold", String(t.fuel_variance_threshold));
  if (download) q.set("download", "1");
  return openProtectedPdf(`${API}/maintenance/insights.pdf?${q.toString()}`, {
    download,
    filename: `IRONLOG_Maintenance_Insights_${start}_to_${end}.pdf`,
  });
}

async function loadMaintenanceInsights() {
  const msgEl = document.getElementById("insightsMsg");
  const startEl = document.getElementById("insightsStart");
  const endEl = document.getElementById("insightsEnd");
  const nearEl = document.getElementById("insightsNearDueHours");
  const horizonEl = document.getElementById("insightsPredictiveHorizonHours");
  const failEl = document.getElementById("insightsChecklistFailThreshold");
  const fuelEl = document.getElementById("insightsFuelVarianceThreshold");
  if (!msgEl || !startEl || !endEl || !nearEl || !horizonEl || !failEl || !fuelEl) return;
  const start = String(startEl.value || "").trim();
  const end = String(endEl.value || "").trim();
  const thresholds = persistInsightsThresholds();
  const near_due_hours = Math.max(1, Number(thresholds.near_due_hours || 50));
  const predictive_horizon_hours = Math.max(near_due_hours, Number(thresholds.predictive_horizon_hours || 100));
  const checklist_fail_threshold = Math.max(1, Number(thresholds.checklist_fail_threshold || 2));
  const fuel_variance_threshold = Math.max(0, Number(thresholds.fuel_variance_threshold || 15));
  if (!start || !end) {
    msgEl.className = "message-error";
    msgEl.textContent = "Select insights start and end dates.";
    return;
  }
  msgEl.className = "muted";
  msgEl.textContent = "Loading maintenance insights...";
  try {
    const q = new URLSearchParams();
    q.set("start", start);
    q.set("end", end);
    q.set("near_due_hours", String(near_due_hours));
    q.set("predictive_horizon_hours", String(predictive_horizon_hours));
    q.set("checklist_fail_threshold", String(checklist_fail_threshold));
    q.set("fuel_variance_threshold", String(fuel_variance_threshold));
    const res = await fetch(`${API}/maintenance/insights?${q.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load maintenance insights");
    renderMaintenanceInsights(data);
    msgEl.className = "message-success";
    msgEl.textContent = `Insights loaded for ${data?.range?.start || start} to ${data?.range?.end || end}.`;
  } catch (e) {
    msgEl.className = "message-error";
    msgEl.textContent = `Insights error: ${e.message || e}`;
  }
}

async function loadGovernanceSignals() {
  const msgEl = document.getElementById("governanceSignalsMsg");
  const start = String(document.getElementById("insightsStart")?.value || "").trim();
  const end = String(document.getElementById("insightsEnd")?.value || "").trim();
  if (!msgEl) return;
  if (!start || !end) {
    msgEl.className = "message-error";
    msgEl.textContent = "Select insights start and end dates.";
    return;
  }
  msgEl.className = "muted";
  msgEl.textContent = "Loading governance signals...";
  try {
    const q = new URLSearchParams();
    q.set("start", start);
    q.set("end", end);
    const res = await fetch(`${API}/maintenance/governance/signals?${q.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load governance signals");
    renderGovernanceSignals(data);
    msgEl.className = "message-success";
    msgEl.textContent = `Governance signals loaded for ${data?.range?.start || start} to ${data?.range?.end || end}.`;
  } catch (e) {
    msgEl.className = "message-error";
    msgEl.textContent = `Governance error: ${e.message || e}`;
  }
}

let insightsForecastCache = [];
let insightsDraftItems = [];
let insightsInputsCache = [];
