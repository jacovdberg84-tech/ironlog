// IRONLOG/web/app/telematics.js — Telematics, Cartrack fleet/map/speeding, GPS links, Unitech.
// Part of the main app; index.html loads these files in order and they share one global scope.

const TELEMATICS_FAULT_POLL_MS = 45000;
let telematicsFaultPollTimer = null;

function telematicsStatusClass(status) {
  const s = String(status || "offline");
  if (s === "live") return "pill green";
  if (s === "stale") return "pill amber";
  return "pill";
}

function buildTelematicsUnitItemHtml(r) {
  const status = String(r.link_status || "offline");
  const statusClass = telematicsStatusClass(status);
  const runHrs = r.run_seconds_today != null ? (Number(r.run_seconds_today) / 3600).toFixed(2) : "-";
  const idleHrs = r.idle_seconds_today != null ? (Number(r.idle_seconds_today) / 3600).toFixed(2) : "-";
  const faults = Number(r.active_fault_count || 0);
  return (
    `<div class="fuel-item-head"><b>${escapeHtml(r.asset_code || "-")}</b> — ${escapeHtml(r.unit_model || "FSC")} <span class="${statusClass}">${status.toUpperCase()}</span>${faults > 0 ? ` <span class="pill red">${faults} FAULT${faults === 1 ? "" : "S"}</span>` : ""}</div>` +
    `<small class="fuel-item-desc">${escapeHtml(r.asset_name || "")}</small>` +
    `<small class="fuel-item-meta">Serial: ${escapeHtml(r.device_serial || "-")} | Meter: ${r.engine_hours == null ? "-" : Number(r.engine_hours).toFixed(1)} h | Run today: ${runHrs} h | Idle today: ${idleHrs} h | Ignition: ${Number(r.ignition_on) === 1 ? "ON" : "OFF"}</small>` +
    `<small class="fuel-item-meta muted">Last seen: ${escapeHtml(r.recorded_at || r.updated_at || r.snapshot_updated_at || "-")}</small>`
  );
}

function renderTelematicsFleetList(container, fleet) {
  if (!container) return;
  container.innerHTML = "";
  const rows = Array.isArray(fleet) ? fleet : [];
  if (!rows.length) {
    container.appendChild(item("<small>No telematics devices registered yet.</small>"));
    return;
  }
  rows.forEach((r) => {
    container.appendChild(item(buildTelematicsUnitItemHtml(r)));
  });
}

function renderTelematicsActiveFaultsList(container, faults) {
  if (!container) return;
  container.innerHTML = "";
  const rows = Array.isArray(faults) ? faults : [];
  if (!rows.length) {
    container.appendChild(item("<small>No active faults — all units clear.</small>"));
    return;
  }
  rows.forEach((f) => {
    const sev = String(f.severity || "warning").toLowerCase();
    const sevClass = sev === "critical" || sev === "error" || sev === "severe" ? "pill red" : "pill amber";
    container.appendChild(
      item(
        `<div class="fuel-item-head"><b>${escapeHtml(f.asset_code || "-")}</b> <span class="${sevClass}">${escapeHtml(sev.toUpperCase())}</span> <code>${escapeHtml(f.fault_code || "-")}</code></div>` +
        `<small class="fuel-item-desc">${escapeHtml(f.description || "No description")}</small>` +
        `<small class="fuel-item-meta muted">${escapeHtml(f.unit_model || "FSC")} · ${escapeHtml(f.event_time || "-")}</small>`
      )
    );
  });
}

function updateTelematicsFaultFloat(summary) {
  const banner = qs("telematicsFaultFloat");
  const titleEl = qs("telematicsFaultFloatTitle");
  const textEl = qs("telematicsFaultFloatText");
  if (!banner || !titleEl || !textEl) return;

  const faults = Array.isArray(summary?.faults) ? summary.faults : [];
  const unitsWithFaults = Number(summary?.units_with_faults || 0);
  const faultCount = Number(summary?.fault_count || faults.length || 0);
  const hasFaults = faultCount > 0 || unitsWithFaults > 0;

  if (!hasFaults) {
    banner.classList.add("hidden");
    document.body.classList.remove("telematics-fault-visible");
    return;
  }

  const assetCodes = Array.from(new Set(faults.map((f) => f.asset_code).filter(Boolean)));
  const preview = faults.slice(0, 3).map((f) => {
    const code = f.asset_code || "?";
    const fc = f.fault_code || "?";
    return `${code}: ${fc}`;
  });
  const more = faults.length > 3 ? ` (+${faults.length - 3} more)` : "";

  titleEl.textContent =
    unitsWithFaults === 1
      ? "Active machine fault"
      : `${unitsWithFaults || assetCodes.length} machine(s) with active faults`;
  textEl.textContent = preview.length
    ? `${preview.join(" · ")}${more}`
    : `${faultCount} active fault signal(s) on telematics units`;

  banner.classList.remove("hidden");
  document.body.classList.add("telematics-fault-visible");
}

async function refreshTelematicsFaultBanner() {
  const allowed = getEffectiveAllowedTabs();
  if (!allowed.includes("telematics") && !allowed.includes("dash")) return;
  try {
    const data = await fetchJson(`${API}/api/telematics/faults/active`);
    updateTelematicsFaultFloat(data);
  } catch (_) {
    /* silent — banner stays as last known state */
  }
}

function initTelematicsFaultBanner() {
  qs("telematicsFaultFloatView")?.addEventListener("click", () => {
    switchTab("telematics");
    loadTelematicsTab().catch(() => {});
  });
  refreshTelematicsFaultBanner().catch(() => {});
  if (telematicsFaultPollTimer) clearInterval(telematicsFaultPollTimer);
  telematicsFaultPollTimer = setInterval(() => {
    refreshTelematicsFaultBanner().catch(() => {});
  }, TELEMATICS_FAULT_POLL_MS);
}

async function loadTelematicsTab() {
  const unitsList = qs("telematicsActiveUnitsList");
  const faultsList = qs("telematicsActiveFaultsList");
  if (!unitsList && !faultsList) return;

  setStatus("Loading telematics…");
  if (unitsList) setSkeleton("telematicsActiveUnitsList", 3);
  if (faultsList) setSkeleton("telematicsActiveFaultsList", 2);

  try {
    const [fleetData, faultData] = await Promise.all([
      fetchJson(`${API}/api/telematics/fleet`),
      fetchJson(`${API}/api/telematics/faults/active`),
    ]);
    const fleet = Array.isArray(fleetData?.fleet) ? fleetData.fleet : [];
    renderTelematicsFleetList(unitsList, fleet);
    renderTelematicsActiveFaultsList(faultsList, faultData?.faults || []);
    updateTelematicsFaultFloat(faultData);

    const liveCount = fleet.filter((r) => String(r.link_status) === "live").length;
    const faultUnits = Number(faultData?.units_with_faults || 0);
    const setText = (id, v) => {
      const el = qs(id);
      if (el) el.textContent = String(v);
    };
    setText("telematicsKpiUnits", fleet.length);
    setText("telematicsKpiLive", liveCount);
    setText("telematicsKpiFaults", faultUnits);
    setText("telematicsFaultsBadge", faultData?.fault_count || faultData?.faults?.length || 0);
    qs("telematicsFaultsCard")?.classList.toggle("kpi-alert-crit", faultUnits > 0);
    qs("telematicsKpiFaultsPill")?.classList.toggle("kpi-pill-red", faultUnits > 0);
    setStatus(`Telematics loaded — ${fleet.length} unit(s), ${faultUnits} with faults.`);
  } catch (e) {
    if (unitsList) {
      unitsList.innerHTML = "";
      unitsList.appendChild(item(`<small>Telematics unavailable: ${escapeHtml(e.message || String(e))}</small>`));
    }
    setStatus(`Telematics error: ${e.message || e}`);
  }
}

async function loadTelematicsFleet() {
  const list = qs("telematicsFleetList");
  if (!list) return;
  try {
    const data = await fetchJson(`${API}/api/telematics/fleet`);
    renderTelematicsFleetList(list, data?.fleet);
    if (data?.active_fault_count != null || data?.units_with_faults != null) {
      updateTelematicsFaultFloat({
        fault_count: data.active_fault_count,
        units_with_faults: data.units_with_faults,
        faults: [],
      });
    }
  } catch (e) {
    list.innerHTML = "";
    list.appendChild(item(`<small>Telematics unavailable: ${escapeHtml(e.message || String(e))}</small>`));
  }
}

let cartrackDashboardFleetCache = { fleet: [], speedingToday: [] };

function isGpsFleetVehicleInUse(v) {
  if (!v?.has_gps) return false;
  if (v.is_speeding) return true;
  if (v.gps_source === "unitech") {
    if (v.position_stale) return false;
    return Number(v.speed_kmh || 0) > 0;
  }
  return Number(v.ignition_on) === 1;
}

function renderCartrackFleetIgnitionCell(v) {
  if (v.gps_source === "unitech") {
    const spd = Number(v.speed_kmh || 0);
    if (spd > 0) return '<span class="pill pill-green">Moving</span>';
    return '<span class="pill">Idle</span>';
  }
  const ign = Number(v.ignition_on) === 1;
  return ign ? '<span class="pill pill-green">ON</span>' : '<span class="pill">OFF</span>';
}

function renderCartrackFleetTable(fleet, speedingToday, { showAll = false } = {}) {
  const host = qs("cartrackFleetHost");
  if (!host) return;
  const all = Array.isArray(fleet) ? fleet : [];
  if (!all.length) {
    host.innerHTML = `<div class="cartrack-empty muted small">No GPS vehicles synced yet. If Test connection succeeds but shows 0 vehicles, ask Cartrack to assign your fleet to the API user. Then click Sync now.</div>`;
    return;
  }
  const inUse = all.filter(isGpsFleetVehicleInUse);
  const display = (showAll ? all : inUse).slice().sort((a, b) => {
    if (Boolean(b.is_speeding) !== Boolean(a.is_speeding)) {
      return Number(b.is_speeding) - Number(a.is_speeding);
    }
    return Number(b.speed_kmh || 0) - Number(a.speed_kmh || 0);
  });
  const note = qs("cartrackDashFleetFilterNote");
  if (note) {
    note.textContent = showAll
      ? `Showing all ${all.length} tracked vehicle(s).`
      : inUse.length
        ? `Showing ${inUse.length} in use (ignition on / moving). ${all.length - inUse.length} parked — open Fleet Track for full list.`
        : `No vehicles in use right now. Tick “Show all” or open Fleet Track for the full list.`;
  }
  if (!display.length) {
    host.innerHTML = `<div class="cartrack-empty muted small">No vehicles in use right now (${all.length} tracked, ignition off / stationary). Tick <strong>Show all tracked vehicles</strong> above or open <strong>Fleet Track</strong> for the full list.</div>`;
    return;
  }
  const speedMap = new Map();
  (speedingToday || []).forEach((e) => {
    const k = e.asset_code || e.registration;
    speedMap.set(k, (speedMap.get(k) || 0) + 1);
  });
  const rows = display
    .map((v) => {
      const code = cartrackVehicleLabel(v);
      const sub = cartrackVehicleSubLabel(v);
      const spd = Number(v.speed_kmh || 0);
      const speedEvents = speedMap.get(v.asset_code) || speedMap.get(v.registration) || speedMap.get(code) || 0;
      const rowCls = speedEvents > 0 ? "cartrack-row--alert" : "";
      const batteryPills = renderCartrackBatteryPillsHtml(v);
      return `<tr class="${rowCls}">
        <td>
          <span class="cartrack-vehicle-code">${escapeHtml(code)}</span>
          ${sub ? `<div class="muted mini">${escapeHtml(sub)}</div>` : ""}
        </td>
        <td>${cartrackSourcePillHtml(v)} ${escapeHtml(v.vehicle_name || v.registration || "—")}</td>
        <td class="cartrack-col-num">${spd.toFixed(0)}</td>
        <td class="cartrack-col-status">${renderCartrackFleetIgnitionCell(v)}</td>
        <td class="cartrack-col-battery">${batteryPills || '<span class="muted">—</span>'}</td>
        <td class="cartrack-col-num">${speedEvents ? `<span class="pill pill-red">${speedEvents}</span>` : "—"}</td>
        <td class="cartrack-col-sync muted">${escapeHtml(String(v.synced_at || "").slice(0, 16))}</td>
      </tr>`;
    })
    .join("");
  host.innerHTML = `
    <div class="cartrack-table-scroll">
      <table class="cartrack-fleet-table">
        <thead>
          <tr>
            <th>Vehicle</th>
            <th>Name / reg</th>
            <th class="cartrack-col-num">Speed</th>
            <th>Ignition</th>
            <th>Battery / power</th>
            <th class="cartrack-col-num">Speeding</th>
            <th>Synced</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

async function loadCartrackFleet() {
  const host = qs("cartrackFleetHost");
  const hint = qs("cartrackFleetHint");
  if (!host) return;
  try {
    const data = await fetchJson(`${API}/api/cartrack/fleet`);
    const s = data?.summary || {};
    cartrackDashboardFleetCache = {
      fleet: data.fleet || [],
      speedingToday: data.speeding_today || [],
    };
    setText("cartrackKpiTotal", Number(s.total_vehicles || 0));
    setText("cartrackKpiLive", Number(s.ignition_on || 0));
    setText("cartrackKpiSpeeding", Number(s.speeding_today || 0));
    if (hint) {
      const ct = Number(s.cartrack_vehicles ?? s.total_vehicles ?? 0);
      const ut = Number(s.unitech_vehicles || 0);
      const parts = [];
      if (data.configured) {
        parts.push(`GPS connected (${data.base_url || "MZ"})`);
        if (ut) parts.push(`${ut} Unitech Afungi`);
        parts.push(`Last sync: ${s.last_sync || "—"}`);
      } else {
        parts.push("GPS not configured — add credentials in User Admin → GPS fleet.");
      }
      parts.push("Dashboard table: in use only unless Show all is ticked.");
      hint.textContent = parts.join(" · ");
    }
    renderCartrackFleetTable(data.fleet || [], data.speeding_today || [], {
      showAll: Boolean(qs("cartrackDashShowAll")?.checked),
    });
    loadCartrackSpeedingEvents(todayLocalYmd(), { useCache: true }).catch(() => {});
  } catch (e) {
    host.innerHTML = `<div class="cartrack-empty muted small">Cartrack: ${escapeHtml(e.message || String(e))}</div>`;
  }
}

async function syncCartrackNow() {
  setStatus("Syncing Cartrack fleet…");
  const today = todayLocalYmd();
  await fetchJson(`${API}/api/cartrack/sync`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({ start_date: today, end_date: today }),
  });
  await loadCartrackFleet();
  loadCartrackSpeedingEvents(todayLocalYmd(), { useCache: true }).catch(() => {});
  refreshCartrackSpeedFloat({ refresh: false }).catch(() => {});
  setStatus("Cartrack sync complete ✅");
}

function yesterdayYmd() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function getCartrackSpeedReportDate() {
  const el = qs("cartrackSpeedReportDate") || qs("cartrackTrackSpeedReportDate");
  const v = String(el?.value || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  return todayLocalYmd();
}

function initCartrackSpeedReportDates() {
  const today = todayLocalYmd();
  for (const id of ["cartrackSpeedReportDate", "cartrackTrackSpeedReportDate"]) {
    const el = qs(id);
    if (el && !el.value) el.value = today;
  }
}

function syncCartrackSpeedReportDateInputs(date) {
  const d = String(date || getCartrackSpeedReportDate()).slice(0, 10);
  for (const id of ["cartrackSpeedReportDate", "cartrackTrackSpeedReportDate"]) {
    const el = qs(id);
    if (el) el.value = d;
  }
  return d;
}

function renderCartrackSpeedingEventsTable(host, events, date) {
  if (!host) return;
  if (!events.length) {
    host.innerHTML = `<div class="cartrack-empty muted small">No speeding events recorded for ${escapeHtml(date)}. Events appear when GPS sync detects speed over the limit (Cartrack ≥ alert threshold, Unitech &gt; 60 km/h).</div>`;
    return;
  }
  const rows = events
    .map((e) => {
      const vehicle = e.asset_code || e.registration || "—";
      const reg = e.registration && e.registration !== vehicle ? e.registration : "";
      const time = String(e.event_time || "").slice(0, 16);
      const speed = e.speed_kmh != null ? `${Number(e.speed_kmh).toFixed(0)} km/h` : "—";
      const limit = e.speed_limit_kmh != null ? `${Number(e.speed_limit_kmh).toFixed(0)} km/h` : "—";
      const type = e.event_type_label || e.event_type || "";
      return `<tr>
        <td class="cartrack-col-sync">${escapeHtml(time)}</td>
        <td>
          <span class="cartrack-vehicle-code">${escapeHtml(vehicle)}</span>
          ${reg ? `<div class="muted mini">${escapeHtml(reg)}</div>` : ""}
        </td>
        <td class="cartrack-col-num">${escapeHtml(speed)}</td>
        <td class="cartrack-col-num">${escapeHtml(limit)}</td>
        <td class="muted mini">${escapeHtml(type)}</td>
      </tr>`;
    })
    .join("");
  host.innerHTML = `
    <div class="cartrack-table-scroll">
      <table class="cartrack-fleet-table cartrack-speeding-events-table">
        <thead>
          <tr>
            <th>Time</th>
            <th>Vehicle</th>
            <th class="cartrack-col-num">Speed</th>
            <th class="cartrack-col-num">Limit</th>
            <th>Type</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    <p class="muted mini cartrack-speeding-events-foot">${events.length} event(s) on ${escapeHtml(date)} — click <strong>Speeding PDF</strong> for a printable report.</p>
  `;
}

async function loadCartrackSpeedingEvents(date, {
  hostId = "cartrackSpeedingEventsHost",
  countId = "cartrackSpeedingEventsCount",
  panelId = "cartrackSpeedingEventsPanel",
  useCache = true,
} = {}) {
  const host = qs(hostId);
  if (!host) return [];
  const reportDate = syncCartrackSpeedReportDateInputs(date);
  const cached = useCache && reportDate === todayLocalYmd() ? cartrackDashboardFleetCache.speedingToday : null;
  try {
    const events = cached?.length
      ? cached
      : (await fetchJson(
          `${API}/api/cartrack/events?start=${encodeURIComponent(reportDate)}&end=${encodeURIComponent(reportDate)}&speeding_only=1`
        ))?.rows || [];
    renderCartrackSpeedingEventsTable(host, events, reportDate);
    const countEl = qs(countId);
    if (countEl) countEl.textContent = events.length ? `(${events.length})` : "";
    const panel = qs(panelId);
    if (panel && events.length) panel.open = true;
    return events;
  } catch (e) {
    host.innerHTML = `<div class="cartrack-empty muted small">Could not load speeding log: ${escapeHtml(e.message || String(e))}</div>`;
    return [];
  }
}

function openCartrackSpeedingEventsPanel() {
  const panel = qs("cartrackSpeedingEventsPanel");
  if (!panel) return;
  panel.open = true;
  panel.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

async function openCartrackMorningPdf(date) {
  const reportDate = syncCartrackSpeedReportDateInputs(date);
  setStatus(`Opening speeding PDF for ${reportDate}…`);
  try {
    await openAuthedPdf(
      `${API}/api/cartrack/morning-report.pdf?date=${encodeURIComponent(reportDate)}&_=${Date.now()}`
    );
    setStatus(`Speeding PDF opened (${reportDate})`);
  } catch (e) {
    setStatus(`Speeding PDF error: ${e.message || e}`);
  }
}

async function emailCartrackMorningReport(date) {
  const reportDate = syncCartrackSpeedReportDateInputs(date);
  setStatus(`Sending speeding report for ${reportDate}…`);
  const data = await fetchJson(`${API}/api/cartrack/morning-report/send`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({ date: reportDate }),
  });
  setStatus(`Speeding report emailed (${data.summary?.total_speeding_events ?? 0} events) ✅`);
}

function setCartrackAdminResult(text, ok) {
  const el = qs("cartrackAdminResult");
  if (!el) return;
  el.textContent = String(text || "");
  el.style.color = ok === true ? "#15803d" : ok === false ? "#b91c1c" : "";
}

async function loadCartrackAdminSettings() {
  if (!qs("adminCartrackCard")) return;
  try {
    const data = await fetchJson(`${API}/api/cartrack/settings`);
    const s = data?.settings || {};
    if (qs("cartrackBaseUrl")) qs("cartrackBaseUrl").value = s.base_url || "";
    if (qs("cartrackUsername")) qs("cartrackUsername").value = s.username || "";
    if (qs("cartrackMorningRecipients")) qs("cartrackMorningRecipients").value = s.morning_recipients || "";
    if (qs("cartrackMorningEnabled")) qs("cartrackMorningEnabled").checked = s.morning_enabled !== false;
    const [hh, mm] = String(s.morning_time || "06:00").split(":");
    if (qs("cartrackMorningHour")) qs("cartrackMorningHour").value = hh || "6";
    if (qs("cartrackMorningMinute")) qs("cartrackMorningMinute").value = mm || "0";
    if (qs("cartrackSpeedAlertKmh")) qs("cartrackSpeedAlertKmh").value = String(s.speed_alert_kmh ?? 100);
    setCartrackAdminResult(
      s.configured ? `Configured (${s.source}). Updated ${s.updated_at || "—"}.` : "Not configured yet.",
      s.configured ? true : null
    );
  } catch (e) {
    setCartrackAdminResult(String(e.message || e), false);
  }
}

async function saveCartrackAdminSettings() {
  setCartrackAdminResult("Saving…", null);
  const body = {
    base_url: qs("cartrackBaseUrl")?.value,
    username: qs("cartrackUsername")?.value,
    morning_recipients: qs("cartrackMorningRecipients")?.value,
    morning_enabled: Boolean(qs("cartrackMorningEnabled")?.checked),
    morning_hour: Number(qs("cartrackMorningHour")?.value || 6),
    morning_minute: Number(qs("cartrackMorningMinute")?.value || 0),
    speed_alert_kmh: Number(qs("cartrackSpeedAlertKmh")?.value || 100),
  };
  const pass = String(qs("cartrackPassword")?.value || "").trim();
  if (pass) body.password = pass;
  const data = await fetchJson(`${API}/api/cartrack/settings`, {
    method: "PUT",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify(body),
  });
  if (qs("cartrackPassword")) qs("cartrackPassword").value = "";
  setCartrackAdminResult("Cartrack settings saved.", true);
  loadCartrackFleet().catch(() => {});
  return data;
}

async function testCartrackConnection() {
  setCartrackAdminResult("Testing connection…", null);
  try {
    const data = await fetchJson(`${API}/api/cartrack/test-connection`, { method: "POST" });
    setCartrackAdminResult(data.message || "Connected.", true);
  } catch (e) {
    setCartrackAdminResult(String(e.message || e), false);
  }
}

async function runCartrackMorningNow() {
  setCartrackAdminResult("Running morning report…", null);
  try {
    const data = await fetchJson(`${API}/api/cartrack/morning-report/run`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({ send_email: true }),
    });
    const n = data?.summary?.total_speeding_events ?? 0;
    setCartrackAdminResult(`Report for ${data.report_date}: ${n} speeding event(s).${data.emailed ? " Emailed." : ""}`, true);
    loadCartrackFleet().catch(() => {});
  } catch (e) {
    setCartrackAdminResult(String(e.message || e), false);
  }
}

async function initCartrackAdminPanel() {
  if (!qs("adminCartrackCard")) return;
  await loadCartrackAdminSettings().catch(() => {});
  await loadUnitechAdminSettings().catch(() => {});
  await loadGpsVehicleLinksAdmin().catch(() => {});
}

function setGpsVehicleLinksResult(text, ok = null) {
  const el = qs("gpsVehicleLinksResult");
  if (!el) return;
  el.textContent = String(text || "");
  el.style.color = ok === true ? "#15803d" : ok === false ? "#b91c1c" : "";
}

async function ensureGpsLinkAssetDatalist() {
  const datalist = qs("gpsLinkAssetCodeList");
  if (!datalist || datalist.dataset.loaded === "1") return;
  try {
    const assets = await fetchJson(`${API}/api/assets?include_archived=0`);
    const rows = Array.isArray(assets) ? assets : assets?.assets || [];
    datalist.innerHTML = rows
      .map((a) => {
        const code = String(a.asset_code || "").trim();
        const name = String(a.asset_name || "").trim();
        if (!code) return "";
        return `<option value="${escapeHtml(code)}">${escapeHtml(name ? `${code} — ${name}` : code)}</option>`;
      })
      .join("");
    datalist.dataset.loaded = "1";
  } catch {
    /* optional */
  }
}

function renderGpsVehicleLinksTable(links) {
  const host = qs("gpsVehicleLinksList");
  if (!host) return;
  if (!links?.length) {
    host.innerHTML = `<div class="muted small">No mappings yet. Add one above or link from the suggestions list.</div>`;
    return;
  }
  host.innerHTML = `
    <div class="cartrack-table-scroll">
      <table class="cartrack-fleet-table">
        <thead>
          <tr>
            <th>Registration</th>
            <th>Fleet code</th>
            <th>Source</th>
            <th>Notes</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          ${links.map((l) => `<tr>
            <td><code>${escapeHtml(l.registration)}</code></td>
            <td><strong>${escapeHtml(l.asset_code)}</strong></td>
            <td>${escapeHtml(l.gps_source || "any")}</td>
            <td class="muted">${escapeHtml(l.notes || "—")}</td>
            <td><button type="button" class="btn btn-secondary btn-sm" data-gps-link-delete="${escapeHtml(l.registration)}">Remove</button></td>
          </tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
}

function renderGpsVehicleLinkSuggestions(suggestions) {
  const host = qs("gpsVehicleLinkSuggestions");
  if (!host) return;
  if (!suggestions?.length) {
    host.innerHTML = `<div class="muted small">All synced vehicles are mapped or already match a fleet code.</div>`;
    return;
  }
  host.innerHTML = `
    <div class="cartrack-table-scroll">
      <table class="cartrack-fleet-table">
        <thead>
          <tr>
            <th>Registration</th>
            <th>Current label</th>
            <th>Source</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          ${suggestions.map((s) => `<tr>
            <td><code>${escapeHtml(s.registration)}</code></td>
            <td>${escapeHtml(s.current_label || s.registration)}</td>
            <td>${escapeHtml(s.gps_source || "—")}</td>
            <td><button type="button" class="btn btn-primary btn-sm" data-gps-link-prefill="${escapeHtml(s.registration)}" data-gps-link-source="${escapeHtml(s.gps_source || "any")}">Link</button></td>
          </tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
}

async function loadGpsVehicleLinksAdmin() {
  if (!qs("adminGpsVehicleLinksCard")) return;
  await ensureGpsLinkAssetDatalist();
  try {
    const data = await fetchJson(`${API}/api/cartrack/vehicle-links`);
    renderGpsVehicleLinksTable(data?.links || []);
    renderGpsVehicleLinkSuggestions(data?.suggestions || []);
    setGpsVehicleLinksResult(
      `${(data?.links || []).length} mapping(s), ${(data?.suggestions || []).length} unmapped vehicle(s).`,
      true
    );
  } catch (e) {
    setGpsVehicleLinksResult(String(e.message || e), false);
  }
}

async function saveGpsVehicleLink() {
  const registration = String(qs("gpsLinkRegistration")?.value || "").trim();
  const asset_code = String(qs("gpsLinkAssetCode")?.value || "").trim();
  if (!registration || !asset_code) {
    setGpsVehicleLinksResult("Registration and fleet code are required.", false);
    return;
  }
  setGpsVehicleLinksResult("Saving…", null);
  try {
    const data = await fetchJson(`${API}/api/cartrack/vehicle-links`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({
        registration,
        asset_code,
        gps_source: qs("gpsLinkSource")?.value || "any",
        notes: qs("gpsLinkNotes")?.value || "",
      }),
    });
    setGpsVehicleLinksResult(`Saved ${data.link?.registration} → ${data.link?.asset_code}. Fleet updated.`, true);
    if (qs("gpsLinkRegistration")) qs("gpsLinkRegistration").value = "";
    if (qs("gpsLinkAssetCode")) qs("gpsLinkAssetCode").value = "";
    if (qs("gpsLinkNotes")) qs("gpsLinkNotes").value = "";
    await loadGpsVehicleLinksAdmin();
    loadCartrackFleet().catch(() => {});
    refreshCartrackSpeedFloat({ refresh: false }).catch(() => {});
  } catch (e) {
    setGpsVehicleLinksResult(String(e.message || e), false);
  }
}

async function deleteGpsVehicleLink(registration) {
  const reg = String(registration || "").trim();
  if (!reg) return;
  setGpsVehicleLinksResult("Removing…", null);
  try {
    await fetchJson(`${API}/api/cartrack/vehicle-links/${encodeURIComponent(reg)}`, {
      method: "DELETE",
      headers: authHeaders(),
    });
    setGpsVehicleLinksResult(`Removed mapping for ${reg}.`, true);
    await loadGpsVehicleLinksAdmin();
    loadCartrackFleet().catch(() => {});
  } catch (e) {
    setGpsVehicleLinksResult(String(e.message || e), false);
  }
}

async function applyGpsVehicleLinks() {
  setGpsVehicleLinksResult("Re-applying mappings…", null);
  try {
    const data = await fetchJson(`${API}/api/cartrack/vehicle-links/apply`, { method: "POST" });
    setGpsVehicleLinksResult(`Mappings re-applied to ${data.applied?.updated ?? 0} snapshot row(s).`, true);
    loadCartrackFleet().catch(() => {});
    refreshCartrackSpeedFloat({ refresh: false }).catch(() => {});
  } catch (e) {
    setGpsVehicleLinksResult(String(e.message || e), false);
  }
}

function prefillGpsVehicleLinkForm(registration, gpsSource) {
  if (qs("gpsLinkRegistration")) qs("gpsLinkRegistration").value = registration || "";
  if (qs("gpsLinkSource") && gpsSource) qs("gpsLinkSource").value = gpsSource;
  qs("gpsLinkAssetCode")?.focus();
}

function setUnitechAdminResult(text, ok = null) {
  const el = qs("unitechAdminResult");
  if (!el) return;
  el.textContent = String(text || "");
  el.style.color = ok === true ? "#15803d" : ok === false ? "#b91c1c" : "";
}

async function loadUnitechAdminSettings() {
  if (!qs("adminUnitechCard")) return;
  try {
    const data = await fetchJson(`${API}/api/unitech/settings`);
    const s = data?.settings || {};
    if (qs("unitechFeedLabel")) qs("unitechFeedLabel").value = s.feed_label || "Afungi (Unitech)";
    if (qs("unitechEnabled")) qs("unitechEnabled").checked = s.enabled !== false;
    setUnitechAdminResult(
      s.configured
        ? `Configured (${s.source}). KML URL saved. Updated ${s.updated_at || "—"}.`
        : "Paste your Unitech GpsGate KML feed URL and save.",
      s.configured ? true : null
    );
  } catch (e) {
    setUnitechAdminResult(String(e.message || e), false);
  }
}

async function saveUnitechAdminSettings() {
  setUnitechAdminResult("Saving…", null);
  const body = {
    feed_label: qs("unitechFeedLabel")?.value,
    enabled: Boolean(qs("unitechEnabled")?.checked),
  };
  const url = String(qs("unitechKmlFeedUrl")?.value || "").trim();
  if (url) body.kml_feed_url = url;
  const data = await fetchJson(`${API}/api/unitech/settings`, {
    method: "PUT",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify(body),
  });
  if (qs("unitechKmlFeedUrl")) qs("unitechKmlFeedUrl").value = "";
  setUnitechAdminResult("Unitech settings saved.", true);
  loadCartrackFleet().catch(() => {});
  return data;
}

async function testUnitechConnection() {
  setUnitechAdminResult("Testing KML feed…", null);
  try {
    const data = await fetchJson(`${API}/api/unitech/test-connection`, { method: "POST" });
    setUnitechAdminResult(data.message || "Connected.", true);
  } catch (e) {
    setUnitechAdminResult(String(e.message || e), false);
  }
}

const CARTRACK_TRACK_POLL_MS = 45000;
const CARTRACK_MAP_DEFAULT = { lat: -18.665695, lng: 35.529562, zoom: 6 };
let cartrackLeafletPromise = null;
let cartrackMap = null;
let cartrackMarkers = new Map();
let cartrackTrackPollTimer = null;
let cartrackTrackFleetCache = [];

function ensureLeafletLoaded() {
  if (window.L) return Promise.resolve();
  if (cartrackLeafletPromise) return cartrackLeafletPromise;
  cartrackLeafletPromise = new Promise((resolve, reject) => {
    if (!document.querySelector('link[data-leaflet-css]')) {
      const link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css";
      link.setAttribute("data-leaflet-css", "1");
      document.head.appendChild(link);
    }
    const script = document.createElement("script");
    script.src = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js";
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("Could not load map library"));
    document.head.appendChild(script);
  });
  return cartrackLeafletPromise;
}

function cartrackMarkerColor(v) {
  if (v.is_speeding) return "#dc2626";
  if (v.gps_source === "unitech") {
    if (v.position_stale) return "#a16207";
    return "#ea580c";
  }
  if (Number(v.ignition_on) === 1) return "#16a34a";
  return "#64748b";
}

function cartrackVehicleKey(v) {
  const id = String(v.registration || v.asset_code || "").trim();
  if (!id) return "";
  return v.gps_source === "unitech" ? `unitech:${id}` : id;
}

function cartrackSourcePillHtml(v) {
  if (v.gps_source === "unitech") {
    return `<span class="pill pill-orange cartrack-source-pill">${escapeHtml(v.gps_provider || "Unitech")}</span>`;
  }
  if (v.gps_source === "cartrack") {
    return `<span class="pill pill-blue cartrack-source-pill">Cartrack</span>`;
  }
  return "";
}

function normalizeGpsRegLabel(reg) {
  return String(reg || "").trim().toUpperCase().replace(/\s+/g, "");
}

function cartrackVehicleLabel(v) {
  const fleet = String(v.asset_code || "").trim();
  const reg = String(v.registration || "").trim();
  if (fleet) return fleet;
  return reg || "—";
}

function cartrackVehicleSubLabel(v) {
  const fleet = String(v.asset_code || "").trim();
  const reg = String(v.registration || "").trim();
  const name = String(v.vehicle_name || "").trim();
  if (fleet && reg && normalizeGpsRegLabel(fleet) !== normalizeGpsRegLabel(reg)) {
    return reg;
  }
  return name || reg || "";
}

function formatCartrackTelemetryLine(v) {
  const parts = [];
  if (v.ev_battery_pct != null) parts.push(`EV batt ${Math.round(v.ev_battery_pct)}%`);
  if (v.tracker_battery_pct != null) parts.push(`Tracker ${Math.round(v.tracker_battery_pct)}%`);
  if (v.supply_voltage_v != null) parts.push(`${Number(v.supply_voltage_v).toFixed(1)} V`);
  if (v.fuel_pct != null) parts.push(`Fuel ${Math.round(v.fuel_pct)}%`);
  if (v.charging_status) parts.push(String(v.charging_status));
  return parts.join(" · ");
}

function renderCartrackBatteryPillsHtml(v) {
  const pills = [];
  const low = cartrackBatteryLow(v);
  if (v.ev_battery_pct != null) {
    pills.push(`<span class="pill pill-blue cartrack-batt-pill">EV ${Math.round(v.ev_battery_pct)}%</span>`);
  }
  if (v.tracker_battery_pct != null) {
    const tracker = Math.round(Number(v.tracker_battery_pct));
    const cls = low && tracker > 0 && tracker < 25 ? "pill-red" : "pill-blue";
    pills.push(`<span class="pill ${cls} cartrack-batt-pill">Tracker ${tracker}%</span>`);
  }
  if (v.supply_voltage_v != null) {
    const volts = Number(v.supply_voltage_v);
    const vCls = low && volts > 0 ? "pill-orange" : "pill-gray";
    pills.push(`<span class="pill ${vCls} cartrack-batt-pill">${volts.toFixed(1)} V</span>`);
  }
  if (v.fuel_pct != null) {
    pills.push(`<span class="pill pill-gray cartrack-batt-pill">Fuel ${Math.round(v.fuel_pct)}%</span>`);
  }
  if (v.charging_status) {
    pills.push(`<span class="pill pill-green cartrack-batt-pill">${escapeHtml(String(v.charging_status))}</span>`);
  }
  return pills.join("");
}

function cartrackBatteryLow(v) {
  const tracker = Number(v.tracker_battery_pct);
  if (Number.isFinite(tracker) && tracker > 0 && tracker < 25) return true;
  const volts = Number(v.supply_voltage_v);
  if (!Number.isFinite(volts) || volts <= 0) return false;
  if (volts < 24) return volts < 11.5;
  return volts < 22;
}

function buildCartrackPopupHtml(v) {
  const isUnitech = v.gps_source === "unitech";
  const ign = isUnitech
    ? "—"
    : (Number(v.ignition_on) === 1 ? "ON" : "OFF");
  const spd = Number(v.speed_kmh || 0).toFixed(0);
  const odo = v.odometer_km != null ? `${Number(v.odometer_km).toFixed(0)} km` : "—";
  const when = String(v.last_event_at || v.synced_at || "").slice(0, 16) || "—";
  const lat = Number(v.latitude);
  const lng = Number(v.longitude);
  const mapsUrl = Number.isFinite(lat) && Number.isFinite(lng)
    ? `https://www.google.com/maps?q=${lat},${lng}`
    : "";
  const telemetry = formatCartrackTelemetryLine(v);
  const battLow = cartrackBatteryLow(v);
  const batteryPills = renderCartrackBatteryPillsHtml(v);
  const staleNote = v.position_stale
    ? `<div class="cartrack-stale-note">Position may be stale (${Number(v.position_age_days || 0).toFixed(0)} day(s) old)</div>`
    : "";
  return `
    <div class="cartrack-popup">
      <strong>${escapeHtml(cartrackVehicleLabel(v))}</strong>
      <div>${cartrackSourcePillHtml(v)} ${escapeHtml(cartrackVehicleSubLabel(v) || v.vehicle_name || v.registration || "")}</div>
      <div>Speed: <b>${spd}</b> km/h${isUnitech && v.road_speed_limit != null ? ` · Site max <b>${Number(v.road_speed_limit)}</b> km/h` : ""}${isUnitech ? "" : ` · Ignition: <b>${ign}</b>`}</div>
      ${isUnitech ? "" : `<div>Odometer: ${escapeHtml(odo)}</div>`}
      ${batteryPills ? `<div class="cartrack-track-item-battery ${battLow ? "cartrack-batt-low" : ""}">${batteryPills}</div>` : telemetry ? `<div class="${battLow ? "cartrack-batt-low" : ""}">Battery / power: <b>${escapeHtml(telemetry)}</b></div>` : ""}
      ${staleNote}
      <div class="muted">Last: ${escapeHtml(when)}</div>
      ${mapsUrl ? `<a href="${mapsUrl}" target="_blank" rel="noopener noreferrer">Open in Maps</a>` : ""}
    </div>
  `;
}

function ensureCartrackMap() {
  const host = qs("cartrackMap");
  if (!host || !window.L) return null;
  if (cartrackMap) {
    cartrackMap.invalidateSize();
    return cartrackMap;
  }
  cartrackMap = window.L.map(host, { zoomControl: true }).setView(
    [CARTRACK_MAP_DEFAULT.lat, CARTRACK_MAP_DEFAULT.lng],
    CARTRACK_MAP_DEFAULT.zoom
  );
  window.L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
  }).addTo(cartrackMap);
  return cartrackMap;
}

function fitCartrackMapToFleet(fleet) {
  if (!cartrackMap || !window.L) return;
  const pts = (fleet || [])
    .filter((v) => v.has_gps)
    .map((v) => [Number(v.latitude), Number(v.longitude)]);
  if (!pts.length) {
    cartrackMap.setView([CARTRACK_MAP_DEFAULT.lat, CARTRACK_MAP_DEFAULT.lng], CARTRACK_MAP_DEFAULT.zoom);
    return;
  }
  if (pts.length === 1) {
    cartrackMap.setView(pts[0], 14);
    return;
  }
  cartrackMap.fitBounds(window.L.latLngBounds(pts), { padding: [40, 40], maxZoom: 14 });
}

function updateCartrackMapMarkers(fleet) {
  const map = ensureCartrackMap();
  if (!map || !window.L) return;
  const seen = new Set();
  for (const v of fleet || []) {
    const key = cartrackVehicleKey(v);
    if (!key) continue;
    seen.add(key);
    if (!v.has_gps) {
      const old = cartrackMarkers.get(key);
      if (old) {
        map.removeLayer(old);
        cartrackMarkers.delete(key);
      }
      continue;
    }
    const lat = Number(v.latitude);
    const lng = Number(v.longitude);
    const color = cartrackMarkerColor(v);
    let marker = cartrackMarkers.get(key);
    const icon = window.L.divIcon({
      className: "cartrack-map-marker-wrap",
      html: `<span class="cartrack-map-marker" style="background:${color}"></span>`,
      iconSize: [18, 18],
      iconAnchor: [9, 9],
    });
    if (!marker) {
      marker = window.L.marker([lat, lng], { icon }).addTo(map);
      cartrackMarkers.set(key, marker);
    } else {
      marker.setLatLng([lat, lng]);
      marker.setIcon(icon);
    }
    marker.bindPopup(buildCartrackPopupHtml(v));
    marker._cartrackVehicle = v;
  }
  for (const [key, marker] of cartrackMarkers.entries()) {
    if (!seen.has(key)) {
      map.removeLayer(marker);
      cartrackMarkers.delete(key);
    }
  }
}

function renderCartrackTrackList(fleet, filterText = "") {
  const host = qs("cartrackTrackList");
  if (!host) return;
  const q = String(filterText || "").trim().toLowerCase();
  const rows = (fleet || []).filter((v) => {
    if (!q) return true;
    const hay = `${v.asset_code || ""} ${v.registration || ""} ${v.vehicle_name || ""}`.toLowerCase();
    return hay.includes(q);
  });
  if (!rows.length) {
    host.innerHTML = `<div class="cartrack-track-empty muted small">${q ? "No vehicles match your search." : "No vehicles synced yet. Use Refresh now or sync from the dashboard."}</div>`;
    return;
  }
  host.innerHTML = rows
    .map((v) => {
      const key = cartrackVehicleKey(v);
      const label = cartrackVehicleLabel(v);
      const sub = cartrackVehicleSubLabel(v);
      const ign = Number(v.ignition_on) === 1;
      const spd = Number(v.speed_kmh || 0).toFixed(0);
      const batteryPills = renderCartrackBatteryPillsHtml(v);
      const battLow = cartrackBatteryLow(v);
      const cls = [
        "cartrack-track-item",
        v.is_speeding ? "cartrack-track-item--alert" : "",
        !v.has_gps ? "cartrack-track-item--nogps" : "",
        battLow ? "cartrack-track-item--battlow" : "",
        v.position_stale ? "cartrack-track-item--stale" : "",
        v.gps_source === "unitech" ? "cartrack-track-item--unitech" : "",
      ].filter(Boolean).join(" ");
      const ignPill = v.gps_source === "unitech"
        ? ""
        : (ign ? '<span class="pill pill-green">ON</span>' : '<span class="pill">OFF</span>');
      return `<button type="button" class="${cls}" data-cartrack-key="${escapeHtml(key)}">
        <span class="cartrack-track-item-code">${escapeHtml(label)}</span>
        <span class="cartrack-track-item-meta">${cartrackSourcePillHtml(v)} ${escapeHtml(sub || v.vehicle_name || v.registration || "—")}</span>
        <span class="cartrack-track-item-stats">
          ${ignPill}
          <span>${spd} km/h</span>
          ${v.is_speeding ? '<span class="pill pill-red">Speeding</span>' : ""}
          ${!v.has_gps ? '<span class="pill">No GPS</span>' : ""}
          ${v.position_stale ? '<span class="pill pill-orange">Stale GPS</span>' : ""}
        </span>
        ${batteryPills ? `<span class="cartrack-track-item-battery">${batteryPills}</span>` : ""}
      </button>`;
    })
    .join("");
}

function focusCartrackVehicle(key) {
  const marker = cartrackMarkers.get(String(key || ""));
  if (!marker || !cartrackMap) return;
  cartrackMap.setView(marker.getLatLng(), Math.max(cartrackMap.getZoom(), 14));
  marker.openPopup();
  qs("cartrackTrackList")?.querySelectorAll(".cartrack-track-item").forEach((el) => {
    el.classList.toggle("active", el.getAttribute("data-cartrack-key") === key);
  });
}

function setCartrackMapStatus(text) {
  const el = qs("cartrackMapStatus");
  if (el) el.textContent = String(text || "");
}

function stopCartrackTrackPolling() {
  if (cartrackTrackPollTimer) {
    clearInterval(cartrackTrackPollTimer);
    cartrackTrackPollTimer = null;
  }
}

function startCartrackTrackPolling() {
  stopCartrackTrackPolling();
  if (!qs("cartrackAutoRefresh")?.checked) return;
  cartrackTrackPollTimer = setInterval(() => {
    if (document.querySelector("#tab-cartrack.show")) {
      loadCartrackTrackingTab({ refresh: true, quiet: true }).catch(() => {});
    }
  }, CARTRACK_TRACK_POLL_MS);
}

async function loadCartrackTrackingTab({ refresh = true, quiet = false } = {}) {
  if (!qs("tab-cartrack")) return;
  if (!quiet) setStatus("Loading fleet map…");
  setCartrackMapStatus("Loading positions…");
  try {
    await ensureLeafletLoaded();
    ensureCartrackMap();
    const q = refresh ? "refresh=1" : "refresh=0";
    const data = await fetchJson(`${API}/api/cartrack/live?${q}`);
    const fleet = data?.fleet || [];
    cartrackTrackFleetCache = fleet;
    const s = data?.summary || {};
    setText("cartrackTrackKpiTotal", Number(s.total_vehicles || 0));
    setText("cartrackTrackKpiGps", Number(s.with_gps || 0));
    setText("cartrackTrackKpiLive", Number(s.ignition_on || 0));
    setText("cartrackTrackKpiSpeeding", Number(s.speeding_today || 0));
    const syncBits = [String(s.last_sync || "—").slice(0, 16) || "—"];
    if (Number(s.unitech_vehicles || 0) > 0) syncBits.push(`${s.unitech_vehicles} Unitech`);
    if (Number(s.cartrack_vehicles || 0) > 0) syncBits.push(`${s.cartrack_vehicles} Cartrack`);
    setText("cartrackTrackKpiSync", syncBits.join(" · "));
    const search = qs("cartrackTrackSearch")?.value || "";
    renderCartrackTrackList(fleet, search);
    updateCartrackMapMarkers(fleet);
    if (!cartrackMap?._cartrackFittedOnce && fleet.some((v) => v.has_gps)) {
      fitCartrackMapToFleet(fleet);
      cartrackMap._cartrackFittedOnce = true;
    }
    const statusBits = [];
    if (!data.configured) statusBits.push("Cartrack not configured");
    else statusBits.push(`${Number(s.with_gps || 0)} on map`);
    if (data.sync_error) statusBits.push(`sync: ${data.sync_error}`);
    setCartrackMapStatus(statusBits.join(" · ") || "Ready");
    if (!quiet) setStatus(`Fleet map updated — ${Number(s.with_gps || 0)} vehicle(s) with GPS.`);
    startCartrackTrackPolling();
    loadCartrackSpeedingEvents(getCartrackSpeedReportDate(), {
      hostId: "cartrackTrackSpeedingEventsHost",
      countId: "cartrackTrackSpeedingEventsCount",
      panelId: "cartrackTrackSpeedingEventsPanel",
      useCache: getCartrackSpeedReportDate() === todayLocalYmd(),
    }).catch(() => {});
  } catch (e) {
    setCartrackMapStatus(String(e.message || e));
    if (!quiet) setStatus(`Fleet map error: ${e.message || e}`);
  }
}

function initCartrackTrackingTab() {
  qs("cartrackTrackRefreshBtn")?.addEventListener("click", () =>
    loadCartrackTrackingTab({ refresh: true }).catch((e) => setStatus(String(e.message || e)))
  );
  qs("cartrackTrackFitBtn")?.addEventListener("click", () => fitCartrackMapToFleet(cartrackTrackFleetCache));
  qs("cartrackOpenMapBtn")?.addEventListener("click", () => switchTab("cartrack"));
  qs("cartrackAutoRefresh")?.addEventListener("change", () => startCartrackTrackPolling());
  qs("cartrackTrackSearch")?.addEventListener("input", (e) => {
    renderCartrackTrackList(cartrackTrackFleetCache, e.target?.value || "");
  });
  qs("cartrackTrackList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-cartrack-key]");
    if (!btn) return;
    focusCartrackVehicle(btn.getAttribute("data-cartrack-key"));
  });
}

const CARTRACK_SPEED_FLOAT_POLL_MS = 60000;
let cartrackSpeedFloatPollTimer = null;
let cartrackSpeedFloatFleet = [];

function setCartrackSpeedFloatExpanded(expanded) {
  const root = qs("cartrackSpeedFloat");
  if (!root) return;
  root.classList.toggle("is-collapsed", !expanded);
  root.classList.toggle("is-expanded", expanded);
}

function renderCartrackSpeedFloatList(fleet) {
  const host = qs("cartrackSpeedFloatList");
  if (!host) return;
  const rows = [...(fleet || [])].sort((a, b) => {
    const sa = Number(a.speed_kmh || 0);
    const sb = Number(b.speed_kmh || 0);
    if (sb !== sa) return sb - sa;
    return Number(b.ignition_on) - Number(a.ignition_on);
  });
  if (!rows.length) {
    host.innerHTML = `<div class="cartrack-speed-float-empty muted small">No Cartrack vehicles synced.</div>`;
    return;
  }
  host.innerHTML = rows
    .map((v) => {
      const label = cartrackVehicleLabel(v);
      const sub = cartrackVehicleSubLabel(v);
      const spd = Number(v.speed_kmh || 0);
      const ign = Number(v.ignition_on) === 1;
      const limit = v.road_speed_limit != null ? Number(v.road_speed_limit) : null;
      const cls = [
        "cartrack-speed-float-row",
        v.is_speeding ? "cartrack-speed-float-row--alert" : "",
        ign ? "cartrack-speed-float-row--live" : "",
      ].filter(Boolean).join(" ");
      const extras = [];
      if (ign) extras.push("IGN");
      if (v.is_idling) extras.push("Idle");
      if (v.gps_source === "unitech" && limit) extras.push(`Max ${limit} km/h`);
      else if (limit) extras.push(`Limit ${limit}`);
      else if (v.speed_alert_kmh) extras.push(`Alert ≥${v.speed_alert_kmh}`);
      const batteryPills = renderCartrackBatteryPillsHtml(v);
      const battLow = cartrackBatteryLow(v);
      return `<div class="${cls}${battLow ? " cartrack-speed-float-row--battlow" : ""}">
        <span class="cartrack-speed-float-code">${escapeHtml(label)}</span>
        <span class="cartrack-speed-float-spd">${spd.toFixed(0)}<small> km/h</small></span>
        ${batteryPills ? `<span class="cartrack-speed-float-battery">${batteryPills}</span>` : `<span class="cartrack-speed-float-meta">${escapeHtml(extras.join(" · ") || sub || v.registration || "")}</span>`}
      </div>`;
    })
    .join("");
}

function updateCartrackSpeedFloat(data) {
  const root = qs("cartrackSpeedFloat");
  const badge = qs("cartrackSpeedFloatBadge");
  const updated = qs("cartrackSpeedFloatUpdated");
  if (!root) return;

  const allowed = getEffectiveAllowedTabs();
  if (!allowed.includes("cartrack") && !allowed.includes("dash")) {
    root.classList.add("hidden");
    return;
  }

  const fleet = Array.isArray(data?.fleet) ? data.fleet : [];
  cartrackSpeedFloatFleet = fleet;
  if (!data?.configured || !fleet.length) {
    root.classList.add("hidden");
    return;
  }

  root.classList.remove("hidden");
  const moving = fleet.filter((v) => Number(v.speed_kmh || 0) > 3 || Number(v.ignition_on) === 1).length;
  const speeding = fleet.filter((v) => v.is_speeding).length;
  if (badge) {
    badge.textContent = String(moving);
    badge.classList.toggle("cartrack-speed-float-badge--alert", speeding > 0);
  }
  if (updated) {
    const sync = String(data?.summary?.last_sync || "").slice(11, 16) || "—";
    updated.textContent = `${fleet.length} vehicles · ${moving} active · updated ${sync}`;
  }
  renderCartrackSpeedFloatList(fleet);
}

async function refreshCartrackSpeedFloat({ refresh = true } = {}) {
  const allowed = getEffectiveAllowedTabs();
  if (!allowed.includes("cartrack") && !allowed.includes("dash")) return;
  try {
    const q = refresh ? "refresh=1" : "refresh=0";
    const data = await fetchJson(`${API}/api/cartrack/live?${q}`);
    updateCartrackSpeedFloat(data);
  } catch (_) {
    /* silent */
  }
}

function initCartrackSpeedFloat() {
  const root = qs("cartrackSpeedFloat");
  if (!root) return;

  qs("cartrackSpeedFloatTab")?.addEventListener("click", () => {
    const expanded = root.classList.contains("is-expanded");
    setCartrackSpeedFloatExpanded(!expanded);
    if (!expanded) refreshCartrackSpeedFloat({ refresh: true }).catch(() => {});
  });
  qs("cartrackSpeedFloatMinBtn")?.addEventListener("click", () => setCartrackSpeedFloatExpanded(false));
  qs("cartrackSpeedFloatMapBtn")?.addEventListener("click", () => {
    switchTab("cartrack");
    loadCartrackTrackingTab({ refresh: true }).catch(() => {});
  });
  qs("cartrackSpeedFloatRefreshBtn")?.addEventListener("click", () =>
    refreshCartrackSpeedFloat({ refresh: true }).catch(() => {})
  );

  refreshCartrackSpeedFloat({ refresh: true }).catch(() => {});
  if (cartrackSpeedFloatPollTimer) clearInterval(cartrackSpeedFloatPollTimer);
  cartrackSpeedFloatPollTimer = setInterval(() => {
    refreshCartrackSpeedFloat({ refresh: true }).catch(() => {});
  }, CARTRACK_SPEED_FLOAT_POLL_MS);
}

/** Start-up: Telematics, Cartrack and GPS link controls. Called once from init() in init.js. */
function wireTelematicsControls() {
  initCartrackAdminPanel().catch(() => {});
  initCartrackTrackingTab();

  qs("cartrackSaveSettingsBtn")?.addEventListener("click", () =>
    saveCartrackAdminSettings().catch((e) => setCartrackAdminResult(String(e.message || e), false))
  );
  qs("cartrackTestBtn")?.addEventListener("click", () => testCartrackConnection().catch(() => {}));
  qs("cartrackRunMorningBtn")?.addEventListener("click", () => runCartrackMorningNow().catch(() => {}));
  qs("unitechSaveSettingsBtn")?.addEventListener("click", () =>
    saveUnitechAdminSettings().catch((e) => setUnitechAdminResult(String(e.message || e), false))
  );
  qs("unitechTestBtn")?.addEventListener("click", () => testUnitechConnection().catch(() => {}));
  qs("gpsLinkSaveBtn")?.addEventListener("click", () => saveGpsVehicleLink().catch(() => {}));
  qs("gpsLinkRefreshBtn")?.addEventListener("click", () => loadGpsVehicleLinksAdmin().catch(() => {}));
  qs("gpsLinkApplyBtn")?.addEventListener("click", () => applyGpsVehicleLinks().catch(() => {}));
  qs("gpsVehicleLinksList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-gps-link-delete]");
    if (!btn) return;
    deleteGpsVehicleLink(btn.getAttribute("data-gps-link-delete")).catch(() => {});
  });
  qs("gpsVehicleLinkSuggestions")?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-gps-link-prefill]");
    if (!btn) return;
    prefillGpsVehicleLinkForm(btn.getAttribute("data-gps-link-prefill"), btn.getAttribute("data-gps-link-source"));
  });
  qs("cartrackSyncBtn")?.addEventListener("click", () => syncCartrackNow().catch((e) => setStatus(String(e.message || e))));
  qs("cartrackDashShowAll")?.addEventListener("change", () => {
    renderCartrackFleetTable(
      cartrackDashboardFleetCache.fleet,
      cartrackDashboardFleetCache.speedingToday,
      { showAll: Boolean(qs("cartrackDashShowAll")?.checked) }
    );
  });
  qs("cartrackMorningPdfBtn")?.addEventListener("click", () =>
    openCartrackMorningPdf().catch((e) => setStatus(String(e.message || e)))
  );
  qs("cartrackMorningEmailBtn")?.addEventListener("click", () =>
    emailCartrackMorningReport().catch((e) => setStatus(String(e.message || e)))
  );
  qs("cartrackTrackMorningPdfBtn")?.addEventListener("click", () =>
    openCartrackMorningPdf().catch((e) => setStatus(String(e.message || e)))
  );
  const onCartrackSpeedReportDateChange = (e) => {
    const d = syncCartrackSpeedReportDateInputs(e.target?.value);
    loadCartrackSpeedingEvents(d, { useCache: d === todayLocalYmd() }).catch(() => {});
    loadCartrackSpeedingEvents(d, {
      hostId: "cartrackTrackSpeedingEventsHost",
      countId: "cartrackTrackSpeedingEventsCount",
      panelId: "cartrackTrackSpeedingEventsPanel",
      useCache: d === todayLocalYmd(),
    }).catch(() => {});
  };
  qs("cartrackSpeedReportDate")?.addEventListener("change", onCartrackSpeedReportDateChange);
  qs("cartrackTrackSpeedReportDate")?.addEventListener("change", onCartrackSpeedReportDateChange);
  qs("cartrackKpiSpeeding")?.closest(".kpi-pill")?.addEventListener("click", () => {
    openCartrackSpeedingEventsPanel();
    loadCartrackSpeedingEvents(todayLocalYmd(), { useCache: true }).catch(() => {});
  });
  initCartrackSpeedReportDates();

  qs("telemSaveDeviceBtn")?.addEventListener("click", () =>
    saveTelematicsDevice().catch((e) => setTelemAdminResult(String(e.message || e)))
  );
  qs("telemRefreshDevicesBtn")?.addEventListener("click", () =>
    loadTelematicsAdminDevices().catch((e) => setTelemAdminResult(String(e.message || e)))
  );
  qs("telemShowInactive")?.addEventListener("change", () =>
    loadTelematicsAdminDevices().catch(() => {})
  );
  qs("telematicsRefreshBtn")?.addEventListener("click", () =>
    loadTelematicsTab().catch((e) => setStatus("Telematics refresh error: " + (e.message || e)))
  );
  qs("telemDevicesList")?.addEventListener("click", (e) => {
    const editBtn = e.target.closest("button[data-telem-edit]");
    if (editBtn) {
      fillTelematicsDeviceForm({
        assetCode: editBtn.getAttribute("data-telem-asset"),
        deviceSerial: "",
        unitModel: editBtn.getAttribute("data-telem-model") || "FSC650",
        externalId: "",
        replaceFaulty: true,
      });
      setTelemAdminResult(`Enter new serial for ${editBtn.getAttribute("data-telem-asset")} (was ${editBtn.getAttribute("data-telem-serial")}).`);
      return;
    }
    const deactBtn = e.target.closest("button[data-telem-deactivate]");
    if (deactBtn) {
      deactivateTelematicsDeviceAdmin(
        deactBtn.getAttribute("data-telem-deactivate"),
        deactBtn.getAttribute("data-telem-asset-label")
      ).catch((err) => setTelemAdminResult(String(err.message || err)));
    }
  });
}
