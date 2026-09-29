// IRONLOG/web/app/site-ops.js — Site operations, closing, dispatch, data quality.
// Part of the main app; index.html loads these files in order and they share one global scope.

function getSiteOpsFrom() {
  return (qs("opFrom")?.value || "").trim();
}

function getSiteOpsTo() {
  return (qs("opTo")?.value || "").trim();
}

async function loadSiteZones() {
  const data = await fetchJson(`${API}/api/operations/site/zones`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  const zoneSelect = qs("opSiteZone");
  const zoneList = qs("siteZoneList");
  if (zoneSelect) {
    zoneSelect.innerHTML = `<option value="">Select zone</option>`;
    rows
      .filter((r) => Number(r.active || 0) === 1)
      .forEach((r) => {
        const opt = document.createElement("option");
        opt.value = String(r.id || "");
        opt.textContent = `${r.name || "Zone"} (#${r.id})`;
        zoneSelect.appendChild(opt);
      });
  }
  if (zoneList) {
    zoneList.innerHTML = "";
    if (!rows.length) {
      zoneList.appendChild(item("<small>No zones configured.</small>"));
    } else {
      rows.forEach((r) => {
        zoneList.appendChild(item(`<b>#${r.id}</b> ${r.name || "-"} <span class="pill ${Number(r.active || 0) ? "blue" : "orange"}">${Number(r.active || 0) ? "active" : "inactive"}</span>`));
      });
    }
  }
}

async function saveSiteZone() {
  const name = String(qs("opZoneName")?.value || "").trim();
  if (!name) {
    alert("Zone name is required.");
    return;
  }
  const active = Boolean(qs("opZoneActive")?.checked);
  const res = await fetchJson(`${API}/api/operations/site/zones`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, active }),
  });
  setText("siteDailyResult", JSON.stringify(res, null, 2));
  await loadSiteZones();
  setStatus("Site zone saved.");
}

async function saveSiteDailyEntry() {
  const payload = {
    op_date: (qs("opSiteDate")?.value || "").trim(),
    material_type: (qs("opSiteMaterial")?.value || "").trim(),
    zone_id: (qs("opSiteZone")?.value || "").trim() === "" ? undefined : Number(qs("opSiteZone")?.value || 0),
    planned_tonnage: (qs("opSitePlanned")?.value || "").trim() === "" ? undefined : Number(qs("opSitePlanned")?.value || 0),
    actual_tonnage: (qs("opSiteActual")?.value || "").trim() === "" ? undefined : Number(qs("opSiteActual")?.value || 0),
    loads_count: (qs("opSiteLoads")?.value || "").trim() === "" ? undefined : Number(qs("opSiteLoads")?.value || 0),
    avg_cycle_time: (qs("opSiteCycle")?.value || "").trim() === "" ? undefined : Number(qs("opSiteCycle")?.value || 0),
    operator_name: (qs("opSiteOperator")?.value || "").trim() || undefined,
    notes: (qs("opSiteNotes")?.value || "").trim() || undefined,
  };
  if (!payload.op_date || !payload.material_type) {
    alert("Date and material type are required.");
    return;
  }
  const res = await fetchJson(`${API}/api/operations/site/daily`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("siteDailyResult", JSON.stringify(res, null, 2));
  if (qs("opSiteDailyId") && res?.id) qs("opSiteDailyId").value = String(res.id);
  await loadSiteDailyEntries();
  await loadSiteDashboard();
  setStatus("Site daily entry saved.");
}

async function loadSiteDailyEntries() {
  const list = qs("siteDailyList");
  if (!list) return;
  const from = getSiteOpsFrom();
  const to = getSiteOpsTo();
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (to) q.set("to", to);
  const data = await fetchJson(`${API}/api/operations/site/daily${q.toString() ? `?${q.toString()}` : ""}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  if (!rows.length) {
    list.appendChild(item("<small>No site daily entries found.</small>"));
    return;
  }
  rows.forEach((r) => {
    list.appendChild(
      item(
        `<b>#${r.id}</b> ${r.op_date || "-"} <span class="pill blue">${String(r.shift || "-").toUpperCase()}</span> <span class="pill">${r.material_type || "-"}</span>` +
          `<br><small>Zone: ${r.zone_name || "-"} | Planned: ${Number(r.planned_tonnage || 0).toFixed(2)} | Actual: ${Number(r.actual_tonnage || 0).toFixed(2)} | Loads: ${Number(r.loads_count || 0)}</small>` +
          `<br><small>Operator: ${r.operator_name || "-"}${r.notes ? ` | ${r.notes}` : ""}</small>`
      )
    );
  });
}

async function saveSiteEquipmentUsage() {
  const dailyId = Number(qs("opSiteDailyId")?.value || 0);
  const assetId = Number(qs("opSiteEqAssetId")?.value || 0);
  const role = String(qs("opSiteEqRole")?.value || "").trim();
  const hours = (qs("opSiteEqHours")?.value || "").trim() === "" ? undefined : Number(qs("opSiteEqHours")?.value || 0);
  if (!dailyId || !assetId || !role) {
    alert("Daily ID, Asset ID, and role are required.");
    return;
  }
  const res = await fetchJson(`${API}/api/operations/site/daily/${dailyId}/equipment`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ asset_id: assetId, role, hours_used: hours }),
  });
  setText("siteDailyResult", JSON.stringify(res, null, 2));
  await loadSiteEquipmentUsage();
  setStatus("Equipment linked to production.");
}

async function loadSiteEquipmentUsage() {
  const dailyId = Number(qs("opSiteDailyId")?.value || 0);
  const list = qs("siteEquipmentUsageList");
  if (!list) return;
  if (!dailyId) {
    list.innerHTML = "";
    list.appendChild(item("<small>Select a daily entry ID to view equipment usage.</small>"));
    return;
  }
  const data = await fetchJson(`${API}/api/operations/site/daily/${dailyId}/equipment`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  if (!rows.length) {
    list.appendChild(item("<small>No equipment usage linked yet.</small>"));
    return;
  }
  rows.forEach((r) => {
    list.appendChild(item(`<b>#${r.id}</b> Asset ${r.asset_code || r.asset_id} (${r.asset_name || "-"}) | Role: ${r.role || "-"} | Hours: ${Number(r.hours_used || 0).toFixed(2)}`));
  });
}

async function saveSiteTarget() {
  const payload = {
    target_date: (qs("opTargetDate")?.value || "").trim(),
    material_type: (qs("opTargetMaterial")?.value || "").trim(),
    target_tonnage: (qs("opTargetTonnage")?.value || "").trim() === "" ? undefined : Number(qs("opTargetTonnage")?.value || 0),
  };
  if (!payload.target_date || !payload.material_type || payload.target_tonnage == null) {
    alert("Target date, material, and tonnage are required.");
    return;
  }
  const res = await fetchJson(`${API}/api/operations/site/targets`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("siteDailyResult", JSON.stringify(res, null, 2));
  await loadSiteTargets();
  await loadSiteDashboard();
  setStatus("Site target saved.");
}

async function loadSiteTargets() {
  const list = qs("siteTargetList");
  if (!list) return;
  const from = getSiteOpsFrom();
  const to = getSiteOpsTo();
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (to) q.set("to", to);
  const data = await fetchJson(`${API}/api/operations/site/targets${q.toString() ? `?${q.toString()}` : ""}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  if (!rows.length) {
    list.appendChild(item("<small>No targets found.</small>"));
    return;
  }
  rows.forEach((r) => {
    list.appendChild(item(`<b>${r.target_date}</b> ${r.material_type || "-"} <span class="pill blue">${Number(r.target_tonnage || 0).toFixed(2)} t</span>`));
  });
}

async function saveSiteDelay() {
  const payload = {
    delay_date: (qs("opDelayDate")?.value || "").trim(),
    delay_type: (qs("opDelayType")?.value || "").trim(),
    hours_lost: (qs("opDelayHours")?.value || "").trim() === "" ? undefined : Number(qs("opDelayHours")?.value || 0),
    impact_tonnage: (qs("opDelayImpact")?.value || "").trim() === "" ? undefined : Number(qs("opDelayImpact")?.value || 0),
    notes: (qs("opDelayNotes")?.value || "").trim() || undefined,
  };
  if (!payload.delay_date || !payload.delay_type || payload.hours_lost == null) {
    alert("Delay date, type and hours lost are required.");
    return;
  }
  const res = await fetchJson(`${API}/api/operations/site/delays`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("siteDailyResult", JSON.stringify(res, null, 2));
  await loadSiteDelays();
  await loadSiteDashboard();
  setStatus("Operational delay saved.");
}

async function loadSiteDelays() {
  const list = qs("siteDelayList");
  if (!list) return;
  const from = getSiteOpsFrom();
  const to = getSiteOpsTo();
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (to) q.set("to", to);
  const data = await fetchJson(`${API}/api/operations/site/delays${q.toString() ? `?${q.toString()}` : ""}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  if (!rows.length) {
    list.appendChild(item("<small>No operational delays found.</small>"));
    return;
  }
  rows.forEach((r) => {
    list.appendChild(item(`<b>${r.delay_date}</b> <span class="pill orange">${r.delay_type}</span> | Hours lost: ${Number(r.hours_lost || 0).toFixed(2)} | Impact: ${Number(r.impact_tonnage || 0).toFixed(2)}t${r.notes ? `<br><small>${r.notes}</small>` : ""}`));
  });
}

async function loadSiteDashboard() {
  const date = (qs("opSiteDashDate")?.value || qs("opSiteDate")?.value || "").trim();
  if (!date) return;
  const data = await fetchJson(`${API}/api/operations/site/dashboard?date=${encodeURIComponent(date)}`);
  const today = data?.today || {};
  const week = data?.week || {};
  const losses = data?.losses || {};
  setText("opSiteKpiTodayTons", Number(today.total_tons_produced || 0).toFixed(2));
  setText("opSiteKpiAchieved", Number(today.achieved_pct || 0).toFixed(1));
  setText("opSiteKpiLoads", String(Number(today.loads_moved || 0)));
  setText("opSiteKpiZones", String(Number(today.active_zones || 0)));
  setText("opSiteKpiWeekTotal", Number(week.total_production || 0).toFixed(2));
  setText("opSiteKpiBreakdownLoss", Number(losses.breakdown_hours || 0).toFixed(2));
  setText("opSiteKpiOpsLoss", Number(losses.operational_delay_hours || 0).toFixed(2));

  const list = qs("siteDashboardList");
  if (list) {
    list.innerHTML = "";
    const best = week.best_day ? `${week.best_day.date} (${Number(week.best_day.tons || 0).toFixed(2)}t)` : "-";
    const worst = week.worst_day ? `${week.worst_day.date} (${Number(week.worst_day.tons || 0).toFixed(2)}t)` : "-";
    list.appendChild(item(`<b>Best day:</b> ${best}`));
    list.appendChild(item(`<b>Worst day:</b> ${worst}`));
    list.appendChild(item(`<b>Today target:</b> ${Number(today.target_tonnage || 0).toFixed(2)}t | <b>Shortfall:</b> ${Math.max(0, Number(today.target_tonnage || 0) - Number(today.total_tons_produced || 0)).toFixed(2)}t`));
  }
}

async function saveOperationEntry() {
  const payload = {
    op_date: (qs("opDate")?.value || "").trim() || undefined,
    tonnes_moved: (qs("opTonnesMoved")?.value || "").trim() === "" ? undefined : Number(qs("opTonnesMoved")?.value || 0),
    product_type: (qs("opProductType")?.value || "").trim() || undefined,
    product_produced: (qs("opProductProduced")?.value || "").trim() === "" ? undefined : Number(qs("opProductProduced")?.value || 0),
    trucks_loaded: (qs("opTrucksLoaded")?.value || "").trim() === "" ? undefined : Number(qs("opTrucksLoaded")?.value || 0),
    loads_count: (qs("opLoadsCount")?.value || "").trim() === "" ? undefined : Number(qs("opLoadsCount")?.value || 0),
    crusher_feed_tonnes: (qs("opCrusherFeedTonnes")?.value || "").trim() === "" ? undefined : Number(qs("opCrusherFeedTonnes")?.value || 0),
    crusher_output_tonnes: (qs("opCrusherOutputTonnes")?.value || "").trim() === "" ? undefined : Number(qs("opCrusherOutputTonnes")?.value || 0),
    crusher_hours: (qs("opCrusherHours")?.value || "").trim() === "" ? undefined : Number(qs("opCrusherHours")?.value || 0),
    crusher_downtime_hours: (qs("opCrusherDowntime")?.value || "").trim() === "" ? undefined : Number(qs("opCrusherDowntime")?.value || 0),
    weighbridge_amount: (qs("opWeighbridgeAmount")?.value || "").trim() === "" ? undefined : Number(qs("opWeighbridgeAmount")?.value || 0),
    trucks_delivered: (qs("opTrucksDelivered")?.value || "").trim() === "" ? undefined : Number(qs("opTrucksDelivered")?.value || 0),
    product_delivered: (qs("opProductDelivered")?.value || "").trim() === "" ? undefined : Number(qs("opProductDelivered")?.value || 0),
    client_delivered_to: (qs("opClientDeliveredTo")?.value || "").trim() || undefined,
    notes: (qs("opNotes")?.value || "").trim() || undefined,
  };
  setStatus("Saving operations entry...");
  const res = await fetchJson(`${API}/api/operations`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("operationsResult", JSON.stringify(res, null, 2));
  await loadOperations();
  setStatus("Operations entry saved.");
}

async function loadOperationsClosingForDate(opDate) {
  const date = String(opDate || "").trim();
  if (!date) return;
  const data = await fetchJson(`${API}/api/operations/closing/${encodeURIComponent(date)}`);
  const row = data?.row || null;
  if (qs("opCloseShift")) qs("opCloseShift").value = row?.shift_name || "";
  if (qs("opCloseSupervisor")) qs("opCloseSupervisor").value = row?.supervisor_name || "";
  if (qs("opCloseVariance")) qs("opCloseVariance").value = row?.variance_note || "";
  if (qs("opCloseChkWeighbridge")) qs("opCloseChkWeighbridge").checked = Boolean(row?.checklist_weighbridge_reconciled);
  if (qs("opCloseChkTrucks")) qs("opCloseChkTrucks").checked = Boolean(row?.checklist_trucks_reconciled);
  if (qs("opCloseChkClient")) qs("opCloseChkClient").checked = Boolean(row?.checklist_client_confirmed);
  const status = String(row?.status || "open").toUpperCase();
  setText("opCloseStatusPill", `Day Status: ${status}`);
  if (row) setText("operationsClosingResult", JSON.stringify(row, null, 2));
  await loadOperationsClosingHistoryForDate(date);
}

async function loadOperationsClosingHistoryForDate(opDate) {
  const date = String(opDate || "").trim();
  const list = qs("operationsClosingHistory");
  if (!date || !list) return;
  const data = await fetchJson(`${API}/api/operations/closing/${encodeURIComponent(date)}/history`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  if (!rows.length) {
    list.appendChild(item("<small>No closure history for this date.</small>"));
    return;
  }
  rows.forEach((r) => {
    const p = r.payload && typeof r.payload === "object" ? r.payload : {};
    const reason = p.reason ? ` | Reason: ${p.reason}` : "";
    const supervisor = p.supervisor_name ? ` | Supervisor: ${p.supervisor_name}` : "";
    list.appendChild(
      item(
        `<b>${r.action}</b> <span class="pill blue">${r.role || "-"}</span>` +
          `<br><small>User: ${r.username || "-"} | At: ${r.created_at || "-"}</small>` +
          `<br><small>Status: ${p.status || "-"}${supervisor}${reason}</small>`
      )
    );
  });
}

async function saveOperationsClosing(closeDay) {
  const op_date = (qs("opDate")?.value || "").trim();
  if (!op_date) {
    alert("Select an operations date first.");
    return;
  }
  const payload = {
    op_date,
    shift_name: (qs("opCloseShift")?.value || "").trim() || undefined,
    supervisor_name: (qs("opCloseSupervisor")?.value || "").trim() || undefined,
    variance_note: (qs("opCloseVariance")?.value || "").trim() || undefined,
    checklist_weighbridge_reconciled: Boolean(qs("opCloseChkWeighbridge")?.checked),
    checklist_trucks_reconciled: Boolean(qs("opCloseChkTrucks")?.checked),
    checklist_client_confirmed: Boolean(qs("opCloseChkClient")?.checked),
    close_day: Boolean(closeDay),
  };
  if (closeDay) {
    if (!payload.supervisor_name) {
      alert("Supervisor sign-off name is required to close day.");
      return;
    }
    if (!(payload.checklist_weighbridge_reconciled && payload.checklist_trucks_reconciled && payload.checklist_client_confirmed)) {
      alert("Complete all checklist items before closing day.");
      return;
    }
  }
  setStatus(closeDay ? "Closing operations day..." : "Saving operations closing draft...");
  const res = await fetchJson(`${API}/api/operations/closing`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("operationsClosingResult", JSON.stringify(res?.row || res, null, 2));
  await loadOperationsClosingForDate(op_date);
  setStatus(closeDay ? "Operations day closed." : "Operations closing draft saved.");
}

async function reopenOperationsDay() {
  const roles = getSessionRoles();
  if (!roles.some((r) => ["admin", "supervisor"].includes(r))) {
    alert("Only admin or supervisor can re-open a closed day.");
    return;
  }
  const op_date = (qs("opDate")?.value || "").trim();
  if (!op_date) {
    alert("Select an operations date first.");
    return;
  }
  const reopen_reason = (qs("opReopenReason")?.value || "").trim();
  if (!reopen_reason) {
    alert("Re-open reason is required.");
    return;
  }
  const payload = {
    op_date,
    reopen_day: true,
    reopen_reason,
    close_day: false,
    shift_name: (qs("opCloseShift")?.value || "").trim() || undefined,
    supervisor_name: (qs("opCloseSupervisor")?.value || "").trim() || undefined,
    variance_note: (qs("opCloseVariance")?.value || "").trim() || undefined,
    checklist_weighbridge_reconciled: Boolean(qs("opCloseChkWeighbridge")?.checked),
    checklist_trucks_reconciled: Boolean(qs("opCloseChkTrucks")?.checked),
    checklist_client_confirmed: Boolean(qs("opCloseChkClient")?.checked),
  };
  setStatus("Re-opening operations day...");
  const res = await fetchJson(`${API}/api/operations/closing`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("operationsClosingResult", JSON.stringify(res?.row || res, null, 2));
  await loadOperationsClosingForDate(op_date);
  setStatus("Operations day re-opened.");
}

async function loadOperations() {
  const list = qs("operationsList");
  if (!list) return;
  const from = (qs("opFrom")?.value || "").trim();
  const to = (qs("opTo")?.value || "").trim();
  const params = [];
  if (from) params.push(`from=${encodeURIComponent(from)}`);
  if (to) params.push(`to=${encodeURIComponent(to)}`);
  const q = params.length ? `?${params.join("&")}` : "";
  setStatus("Loading operations...");
  setSkeleton("operationsList", 2);
  const data = await fetchJson(`${API}/api/operations${q}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  let tonnes = 0;
  let produced = 0;
  let loaded = 0;
  let loadCycles = 0;
  let crusherFeed = 0;
  let crusherOutput = 0;
  const clientTotalsDelivered = new Map();
  const clientTotalsTrucks = new Map();
  const clientTotalsTonnes = new Map();
  rows.forEach((r) => {
    tonnes += Number(r.tonnes_moved || 0);
    produced += Number(r.product_produced || 0);
    loaded += Number(r.trucks_loaded || 0);
    loadCycles += Number(r.loads_count || 0);
    crusherFeed += Number(r.crusher_feed_tonnes || 0);
    crusherOutput += Number(r.crusher_output_tonnes || 0);
    const client = String(r.client_delivered_to || "").trim() || "Unspecified";
    const deliveredQty = Number(r.product_delivered || 0);
    const trucksQty = Number(r.trucks_delivered || 0);
    const tonnesQty = Number(r.tonnes_moved || 0);
    clientTotalsDelivered.set(client, Number(clientTotalsDelivered.get(client) || 0) + deliveredQty);
    clientTotalsTrucks.set(client, Number(clientTotalsTrucks.get(client) || 0) + trucksQty);
    clientTotalsTonnes.set(client, Number(clientTotalsTonnes.get(client) || 0) + tonnesQty);
    list.appendChild(
      item(
        `<b>${r.op_date || "-"}</b> <span class="pill blue">${r.product_type || "product"}</span>` +
          `<br><small>Tonnes moved: ${Number(r.tonnes_moved || 0).toFixed(2)} | Produced: ${Number(r.product_produced || 0).toFixed(2)}</small>` +
          `<br><small>Trucks loaded: ${Number(r.trucks_loaded || 0)} | Load cycles: ${Number(r.loads_count || 0)}</small>` +
          `<br><small>Crusher feed: ${Number(r.crusher_feed_tonnes || 0).toFixed(2)}t | Crusher output: ${Number(r.crusher_output_tonnes || 0).toFixed(2)}t | Crusher h: ${Number(r.crusher_hours || 0).toFixed(2)} | Downtime h: ${Number(r.crusher_downtime_hours || 0).toFixed(2)}</small>` +
          `${r.notes ? `<br><small>Notes: ${r.notes}</small>` : ""}`
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No operations entries found.</small>"));
  setText("opKpiTonnes", tonnes.toFixed(2));
  setText("opKpiProduced", produced.toFixed(2));
  setText("opKpiLoaded", String(loaded));
  setText("opKpiLoads", String(loadCycles));
  const crusherPerf = crusherFeed > 0 ? (crusherOutput / crusherFeed) * 100 : 0;
  setText("opKpiCrusherPerf", crusherPerf.toFixed(1));
  const metric = String(qs("opClientMetric")?.value || "delivered").toLowerCase();
  const metricMap =
    metric === "trucks" ? clientTotalsTrucks :
    metric === "tonnes" ? clientTotalsTonnes :
    clientTotalsDelivered;
  const metricLabel =
    metric === "trucks" ? "Trucks" :
    metric === "tonnes" ? "Tonnes" :
    "Delivered";
  const topN = Math.max(1, Number(qs("opClientTopN")?.value || 8));
  const topClients = Array.from(metricMap.entries())
    .map(([client, qty]) => ({ client, qty: Number(qty || 0) }))
    .sort((a, b) => b.qty - a.qty)
    .slice(0, topN);
  const maxQty = Math.max(1, ...topClients.map((x) => x.qty));
  const chart = qs("opClientChart");
  const perfList = qs("opClientPerfList");
  if (chart) {
    chart.innerHTML = "";
    if (!topClients.length) {
      chart.appendChild(item("<small>No client delivery data in selected range.</small>"));
    } else {
      topClients.forEach((x) => {
        const h = Math.max(6, Math.round((x.qty / maxQty) * 100));
        const bar = document.createElement("div");
        bar.className = "cost-bar op-client-bar";
        bar.style.height = `${h}px`;
        bar.title = `${x.client}: ${x.qty.toFixed(2)} (${metricLabel})`;
        bar.innerHTML =
          `<span class="cost-bar-value">${x.qty.toFixed(1)}</span>` +
          `<span class="cost-bar-label">${x.client.length > 10 ? `${x.client.slice(0, 10)}...` : x.client}</span>`;
        chart.appendChild(bar);
      });
    }
  }
  if (perfList) {
    perfList.innerHTML = "";
    if (!topClients.length) {
      perfList.appendChild(item("<small>No client totals available.</small>"));
    } else {
      topClients.forEach((x, i) => {
        perfList.appendChild(item(`<b>#${i + 1}</b> ${x.client}<br><small>${metricLabel}: ${x.qty.toFixed(2)}</small>`));
      });
    }
  }
  const opDate = (qs("opDate")?.value || "").trim();
  if (opDate) {
    loadOperationsClosingForDate(opDate).catch(() => {});
  }
  setStatus("Operations ready.");
}

async function createDispatchTrip() {
  const payload = {
    op_date: (qs("dpDate")?.value || "").trim() || undefined,
    trip_no: (qs("dpTripNo")?.value || "").trim() || undefined,
    truck_reg: (qs("dpTruckReg")?.value || "").trim(),
    driver_name: (qs("dpDriver")?.value || "").trim() || undefined,
    product_type: (qs("dpProduct")?.value || "").trim() || undefined,
    client_name: (qs("dpClient")?.value || "").trim() || undefined,
    target_tonnes: (qs("dpTargetTonnes")?.value || "").trim() === "" ? undefined : Number(qs("dpTargetTonnes")?.value || 0),
    actual_tonnes: (qs("dpActualTonnes")?.value || "").trim() === "" ? undefined : Number(qs("dpActualTonnes")?.value || 0),
    notes: (qs("dpNotes")?.value || "").trim() || undefined,
  };
  if (!payload.truck_reg) {
    alert("Truck reg is required.");
    return;
  }
  setStatus("Creating dispatch trip...");
  const res = await fetchJson(`${API}/api/dispatch/trips`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("dispatchResult", JSON.stringify(res, null, 2));
  await loadDispatchTrips();
  setStatus("Dispatch trip created.");
}

async function updateDispatchTripStatus(id, status) {
  const requirePod = Boolean(qs("dpRequirePodDelivered")?.checked);
  const res = await fetchJson(`${API}/api/dispatch/trips/${id}/status`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ status, require_pod_for_delivered: requirePod ? 1 : 0 }),
  });
  setText("dispatchResult", JSON.stringify(res?.row || res, null, 2));
  await loadDispatchTrips();
}

async function saveDispatchPod() {
  const tripId = Number(qs("dpActionTripId")?.value || 0);
  if (!tripId) {
    alert("Trip ID is required for POD.");
    return;
  }
  const pod_ref = String(qs("dpPodRef")?.value || "").trim();
  if (!pod_ref) {
    alert("POD ref is required.");
    return;
  }
  const pod_link = String(qs("dpPodLink")?.value || "").trim() || undefined;
  const res = await fetchJson(`${API}/api/dispatch/trips/${tripId}/pod`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ pod_ref, pod_link }),
  });
  setText("dispatchResult", JSON.stringify(res?.row || res, null, 2));
  await loadDispatchTrips();
}

async function createDispatchException() {
  const trip_id = Number(qs("dpActionTripId")?.value || 0);
  if (!trip_id) {
    alert("Trip ID is required for exception.");
    return;
  }
  const exception_type = String(qs("dpExType")?.value || "").trim();
  const severity = String(qs("dpExSeverity")?.value || "medium").trim();
  const owner_name = String(qs("dpExOwner")?.value || "").trim() || undefined;
  const note = String(qs("dpExNote")?.value || "").trim() || undefined;
  const res = await fetchJson(`${API}/api/dispatch/exceptions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ trip_id, exception_type, severity, owner_name, note }),
  });
  setText("dispatchResult", JSON.stringify(res, null, 2));
  await loadDispatchExceptions();
  await loadDispatchTrips();
}

async function resolveDispatchException(id, status) {
  const resolution_note = String(qs("dpExResolveNote")?.value || "").trim() || undefined;
  const res = await fetchJson(`${API}/api/dispatch/exceptions/${id}/resolve`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ status, resolution_note }),
  });
  setText("dispatchResult", JSON.stringify(res?.row || res, null, 2));
  await loadDispatchExceptions();
  await loadDispatchTrips();
}

async function loadDispatchExceptions() {
  const list = qs("dispatchExceptionsList");
  if (!list) return;
  const from = (qs("dpFrom")?.value || "").trim();
  const to = (qs("dpTo")?.value || "").trim();
  const onlyOpen = Boolean(qs("dpExOnlyOpen")?.checked);
  const status = onlyOpen ? "open" : (qs("dpExStatusFilter")?.value || "").trim();
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (to) q.set("to", to);
  if (status) q.set("status", status);
  const data = await fetchJson(`${API}/api/dispatch/exceptions?${q.toString()}`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  let crit = 0;
  let high = 0;
  let med = 0;
  let low = 0;
  list.innerHTML = "";
  rows.forEach((r) => {
    const s = String(r.status || "").toLowerCase();
    const sev = String(r.severity || "").toLowerCase();
    if (s === "open") {
      if (sev === "critical") crit += 1;
      else if (sev === "high") high += 1;
      else if (sev === "medium") med += 1;
      else low += 1;
    }
    const badgeClass = s === "open" ? "red" : s === "resolved" ? "blue" : "";
    const actions = s === "open"
      ? `<button data-dp-ex-id="${r.id}" data-dp-ex-next="resolved">Resolve</button> <button data-dp-ex-id="${r.id}" data-dp-ex-next="waived">Waive</button>`
      : "";
    list.appendChild(
      item(
        `<b>EX #${r.id}</b> <span class="pill ${badgeClass}">${r.status}</span> <span class="pill">${r.exception_type}</span> <span class="pill">${r.severity}</span>` +
        `<br><small>Trip #${r.trip_id} | ${r.op_date || "-"} | Truck ${r.truck_reg || "-"} | Client ${r.client_name || "-"}</small>` +
        `<br><small>Owner: ${r.owner_name || "-"} | Note: ${r.note || "-"}</small>` +
        `${r.resolution_note ? `<br><small>Resolution: ${r.resolution_note}</small>` : ""}` +
        `<br>${actions}`
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No exceptions found.</small>"));
  setText("dpExKpiCritical", String(crit));
  setText("dpExKpiHigh", String(high));
  setText("dpExKpiMedium", String(med));
  setText("dpExKpiLow", String(low));
}

async function loadQualityCenter() {
  const list = qs("qualityList");
  if (!list) return;
  const from = (qs("qFrom")?.value || "").trim();
  const to = (qs("qTo")?.value || "").trim();
  const sev = (qs("qSeverityFilter")?.value || "").trim().toLowerCase();
  const typ = (qs("qTypeFilter")?.value || "").trim();
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (to) q.set("to", to);
  setStatus("Loading data quality...");
  setSkeleton("qualityList", 2);
  const data = await fetchJson(`${API}/api/quality?${q.toString()}`);
  const summary = data?.summary || {};
  setText("qTotal", String(Number(summary.total || 0)));
  setText("qHigh", String(Number(summary.high || 0)));
  setText("qMedium", String(Number(summary.medium || 0)));
  setText("qLow", String(Number(summary.low || 0)));
  setText("qDaily", String(Number(summary.daily_issues || 0)));
  setText("qPodGaps", String(Number(summary.dispatch_pod_gaps || 0)));
  setText("qOpenExceptions", String(Number(summary.exceptions_open || 0)));
  setText("qPendingApprovals", String(Number(summary.approvals_pending || 0)));

  let rows = Array.isArray(data?.rows) ? data.rows : [];
  if (sev) rows = rows.filter((r) => String(r.severity || "").toLowerCase() === sev);
  if (typ) rows = rows.filter((r) => String(r.type || "") === typ);

  list.innerHTML = "";
  rows.slice(0, 500).forEach((r) => {
    const s = String(r.severity || "low").toLowerCase();
    const badgeClass = s === "high" ? "red" : s === "medium" ? "orange" : "";
    const fixBtn = `<button data-q-fix="1" data-q-type="${String(r.type || "")}" data-q-asset="${String(r.asset_code || "")}" data-q-entity="${String(r.entity_id || "")}" data-q-date="${String(r.date || "")}">Open Fix</button>`;
    const resolveBtn = String(r.type || "") === "dispatch_delivered_no_pod"
      ? `<button data-q-resolve="pod" data-q-entity="${String(r.entity_id || "")}" data-q-date="${String(r.date || "")}">Resolve & Refresh</button>`
      : String(r.type || "") === "dispatch_exception_open"
      ? `<button data-q-resolve="exception_resolve" data-q-entity="${String(r.entity_id || "")}" data-q-date="${String(r.date || "")}">Resolve & Refresh</button> <button data-q-resolve="exception_waive" data-q-entity="${String(r.entity_id || "")}" data-q-date="${String(r.date || "")}">Waive & Refresh</button>`
      : "";
    list.appendChild(
      item(
        `<b>${r.type}</b> <span class="pill ${badgeClass}">${s}</span>` +
        `<br><small>Date: ${r.date || "-"} | Asset/Entity: ${r.asset_code || r.entity_id || "-"}</small>` +
        `<br><small>${r.details || "-"}</small>` +
        `<br>${fixBtn} ${resolveBtn}`
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No quality issues for selected filters.</small>"));
  setText("qualityResult", JSON.stringify({ from: data?.from, to: data?.to, shown: rows.length }, null, 2));
  setStatus("Data quality ready.");
}

async function resolveQualityIssueNow(mode, entityId, issueDate) {
  const id = Number(entityId || 0);
  if (!id) return;
  if (mode === "pod") {
    const pod_ref = prompt(`Enter POD ref for delivered trip #${id}:`, "") || "";
    if (!pod_ref.trim()) return;
    const pod_link = prompt("Optional POD link/path:", "") || "";
    await fetchJson(`${API}/api/dispatch/trips/${id}/pod`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pod_ref: pod_ref.trim(), pod_link: pod_link.trim() || undefined }),
    });
    setStatus(`POD saved for trip #${id}. Rechecking quality...`);
    if (qs("qFrom") && issueDate && !qs("qFrom").value) qs("qFrom").value = String(issueDate);
    if (qs("qTo") && issueDate && !qs("qTo").value) qs("qTo").value = String(issueDate);
    await loadQualityCenter();
    return;
  }
  if (mode === "exception_resolve" || mode === "exception_waive") {
    const resolution_note = prompt("Resolution note:", "") || "";
    const status = mode === "exception_waive" ? "waived" : "resolved";
    await fetchJson(`${API}/api/dispatch/exceptions/${id}/resolve`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ status, resolution_note: resolution_note.trim() || undefined }),
    });
    setStatus(`Exception #${id} set to ${status}. Rechecking quality...`);
    await loadQualityCenter();
  }
}

function openQualityFix(type, assetCode, entityId, issueDate) {
  const t = String(type || "");
  if (t === "dispatch_delivered_no_pod") {
    switchTab("dispatch");
    if (qs("dpActionTripId")) qs("dpActionTripId").value = String(entityId || "");
    if (qs("dpStatusFilter")) qs("dpStatusFilter").value = "delivered";
    if (qs("dpFrom") && issueDate) qs("dpFrom").value = String(issueDate);
    if (qs("dpTo") && issueDate) qs("dpTo").value = String(issueDate);
    loadDispatchTrips().catch(() => {});
    setStatus(`Opened Dispatch fix for trip #${entityId}. Save POD and re-check.`);
    return;
  }
  if (t === "dispatch_exception_open") {
    switchTab("dispatch");
    if (qs("dpExStatusFilter")) qs("dpExStatusFilter").value = "open";
    if (qs("dpFrom") && issueDate) qs("dpFrom").value = String(issueDate);
    if (qs("dpTo") && issueDate) qs("dpTo").value = String(issueDate);
    loadDispatchExceptions().catch(() => {});
    setStatus(`Opened Dispatch exceptions. Resolve exception #${entityId}.`);
    return;
  }
  if (t === "approval_pending") {
    switchTab("approvals");
    if (qs("approvalStatus")) qs("approvalStatus").value = "pending";
    loadApprovalRequests().catch(() => {});
    setStatus(`Opened Approvals. Review pending approval #${entityId}.`);
    return;
  }
  if (t.startsWith("daily_")) {
    switchTab("daily");
    if (qs("date") && issueDate) qs("date").value = String(issueDate);
    loadDailyInput().catch(() => {});
    setStatus(`Opened Daily Input for ${issueDate}. Check asset ${assetCode || "-"}.`);
    return;
  }
  setStatus("No quick-fix route for this issue type yet.");
}

function dispatchVarianceMeta(targetTonnes, actualTonnes, tolerancePct) {
  const target = Number(targetTonnes || 0);
  const actual = Number(actualTonnes || 0);
  if (!Number.isFinite(target) || target <= 0) {
    return { pct: 0, cls: "var-warn", label: "NO TARGET", breached: false };
  }
  const pct = ((actual - target) / target) * 100;
  const absPct = Math.abs(pct);
  const t = Math.max(0, Number(tolerancePct || 0));
  if (absPct <= t * 0.5) return { pct, cls: "var-good", label: "OK", breached: false };
  if (absPct <= t) return { pct, cls: "var-warn", label: "WARN", breached: false };
  return { pct, cls: "var-breach", label: "BREACH", breached: true };
}

async function loadDispatchTrips() {
  const list = qs("dispatchList");
  if (!list) return;
  const from = (qs("dpFrom")?.value || "").trim();
  const to = (qs("dpTo")?.value || "").trim();
  const status = (qs("dpStatusFilter")?.value || "").trim();
  const q = new URLSearchParams();
  if (from) q.set("from", from);
  if (to) q.set("to", to);
  if (status) q.set("status", status);
  setStatus("Loading dispatch trips...");
  setSkeleton("dispatchList", 2);
  const [kpi, trips] = await Promise.all([
    fetchJson(`${API}/api/dispatch/kpi?${q.toString()}`),
    fetchJson(`${API}/api/dispatch/trips?${q.toString()}`),
  ]);
  setText("dpKpiTrips", String(Number(kpi?.total_trips || 0)));
  setText("dpKpiDeliveredTonnes", Number(kpi?.delivered_tonnes || 0).toFixed(2));
  setText("dpKpiTurnaround", Number(kpi?.avg_turnaround_hours || 0).toFixed(2));
  setText("dpKpiQueued", String(Number(kpi?.by_status?.queued || 0)));
  setText("dpKpiLoading", String(Number(kpi?.by_status?.loading || 0)));
  setText("dpKpiTransit", String(Number(kpi?.by_status?.in_transit || 0)));
  setText("dpKpiDelivered", String(Number(kpi?.by_status?.delivered || 0)));
  setText("dpKpiReturned", String(Number(kpi?.by_status?.returned || 0)));
  setText("dpKpiPodPct", kpi?.delivered_with_pod_pct == null ? "-" : `${Number(kpi.delivered_with_pod_pct).toFixed(1)}%`);
  setText("dpKpiExceptionsOpen", String(Number(kpi?.exceptions_open || 0)));

  const rows = Array.isArray(trips?.rows) ? trips.rows : [];
  const tolerancePct = Number(qs("dpVarTolerance")?.value || 10);
  const onlyBreaches = Boolean(qs("dpOnlyBreaches")?.checked);
  list.innerHTML = "";
  const lanes = {
    queued: qs("dpLaneQueued"),
    loading: qs("dpLaneLoading"),
    in_transit: qs("dpLaneTransit"),
    delivered: qs("dpLaneDelivered"),
    returned: qs("dpLaneReturned"),
  };
  Object.values(lanes).forEach((el) => { if (el) el.innerHTML = ""; });
  const addLane = (key, html) => {
    if (lanes[key]) lanes[key].appendChild(item(html));
  };
  let varianceBreaches = 0;

  rows.forEach((r) => {
    const s = String(r.status || "queued");
    const variance = dispatchVarianceMeta(r.target_tonnes, r.actual_tonnes, tolerancePct);
    if (onlyBreaches && !variance.breached) return;
    if (variance.breached) varianceBreaches += 1;
    const actions = []
      .concat(s !== "loading" ? [`<button data-dp-status-id="${r.id}" data-dp-next="loading">Loading</button>`] : [])
      .concat(s !== "in_transit" ? [`<button data-dp-status-id="${r.id}" data-dp-next="in_transit">In Transit</button>`] : [])
      .concat(s !== "delivered" ? [`<button data-dp-status-id="${r.id}" data-dp-next="delivered">Delivered</button>`] : [])
      .concat(s !== "returned" ? [`<button data-dp-status-id="${r.id}" data-dp-next="returned">Returned</button>`] : [])
      .join(" ");
    const html =
      `<b>Trip #${r.id}</b> <span class="pill blue">${s}</span> ${r.trip_no ? `<span class="pill">${r.trip_no}</span>` : ""}` +
      ` <span class="pill ${variance.cls}">${variance.label} ${Number(variance.pct || 0).toFixed(1)}%</span>` +
      `<br><small>${r.op_date || "-"} | Truck: ${r.truck_reg || "-"} | Driver: ${r.driver_name || "-"}</small>` +
      `<br><small>Product: ${r.product_type || "-"} | Client: ${r.client_name || "-"}</small>` +
      `<br><small>Target: ${Number(r.target_tonnes || 0).toFixed(2)} | Actual: ${Number(r.actual_tonnes || 0).toFixed(2)} | POD: ${r.pod_ref || "-"}</small>` +
      `<br>${actions}`;
    list.appendChild(item(html));
    addLane(s, html);
  });
  if (!rows.length) {
    list.appendChild(item("<small>No dispatch trips found.</small>"));
  }
  Object.entries(lanes).forEach(([k, el]) => {
    if (el && !el.children.length) el.appendChild(item(`<small>No ${k.replace("_", " ")} trips.</small>`));
  });
  setText("dpKpiVarianceBreach", String(varianceBreaches));
  loadDispatchExceptions().catch(() => {});
  setStatus("Dispatch ready.");
}

/* =========================
   DAILY INPUT (GRID)
========================= */
