// IRONLOG/web/maintenance/tyres.js — Tyre inspections.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

async function loadAssetsForInspection() {
  const selA = document.getElementById("miAsset");
  const selF = document.getElementById("miFilterAsset");
  const aiA = document.getElementById("aiAsset");
  const aiF = document.getElementById("aiFilterAsset");
  const drA = document.getElementById("drAsset");
  const drF = document.getElementById("drFilterAsset");
  const tyreA = document.getElementById("tyreAsset");
  const tyreF = document.getElementById("tyreFilterAsset");
  const ucA = document.getElementById("ucAsset");
  const ucF = document.getElementById("ucFilterAsset");
  const ucQr = document.getElementById("ucQrAsset");
  const tiQr = document.getElementById("tiQrAsset");
  const wiA = document.getElementById("wiAssetSelect");
  const ptoA = document.getElementById("ptoAsset");
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
    if (tyreA) tyreA.innerHTML = `<option value="">Select asset</option>${opts}`;
    if (tyreF) tyreF.innerHTML = `<option value="">All assets</option>${opts}`;
    if (ucA) ucA.innerHTML = `<option value="">Select asset</option>${opts}`;
    if (ucF) ucF.innerHTML = `<option value="">All assets</option>${opts}`;
    if (ucQr) ucQr.innerHTML = `<option value="">Select asset</option>${opts}`;
    if (tiQr) tiQr.innerHTML = `<option value="">Select asset</option>${opts}`;
    if (wiA) wiA.innerHTML = `<option value="">Select asset</option>${opts}`;
    if (ptoA) ptoA.innerHTML = `<option value="">Any / not linked</option>${opts}`;
  } catch (e) {
    selA.innerHTML = `<option value="">Assets load failed</option>`;
    selF.innerHTML = `<option value="">Assets load failed</option>`;
    if (aiA) aiA.innerHTML = `<option value="">Assets load failed</option>`;
    if (aiF) aiF.innerHTML = `<option value="">Assets load failed</option>`;
    if (drA) drA.innerHTML = `<option value="">Assets load failed</option>`;
    if (drF) drF.innerHTML = `<option value="">Assets load failed</option>`;
    if (tyreA) tyreA.innerHTML = `<option value="">Assets load failed</option>`;
    if (tyreF) tyreF.innerHTML = `<option value="">Assets load failed</option>`;
    if (ucA) ucA.innerHTML = `<option value="">Assets load failed</option>`;
    if (ucF) ucF.innerHTML = `<option value="">Assets load failed</option>`;
    if (ucQr) ucQr.innerHTML = `<option value="">Assets load failed</option>`;
    if (wiA) wiA.innerHTML = `<option value="">Assets load failed</option>`;
    if (ptoA) ptoA.innerHTML = `<option value="">Assets load failed</option>`;
  }
}

function tyreInputId(positionKey, field) {
  return `tyre_${positionKey}_${field}`;
}

function tyreStatusBadge(status) {
  const s = String(status || "unknown").toLowerCase();
  if (s === "replace") return `<span class="badge err">Replace</span>`;
  if (s === "warn") return `<span class="badge warn">End of life soon</span>`;
  if (s === "ok") return `<span class="badge ok">OK</span>`;
  return `<span class="badge">Unknown</span>`;
}

function renderTyreLifecyclePanel(data) {
  const panel = document.getElementById("tyreLifecyclePanel");
  const kpis = document.getElementById("tyreLifecycleKpis");
  const alertsEl = document.getElementById("tyreLifecycleAlerts");
  const tableEl = document.getElementById("tyreLifecycleTable");
  const historyEl = document.getElementById("tyreChangeHistory");
  if (!panel || !kpis || !tableEl || !historyEl) return;

  if (!data?.asset?.id) {
    panel.style.display = "none";
    return;
  }
  panel.style.display = "block";
  const s = data.summary || {};
  const thresholds = data.thresholds || {};
  const warnInp = document.getElementById("tyreWarnTread");
  const minInp = document.getElementById("tyreMinTread");
  if (warnInp && thresholds.warn_tread_mm != null) warnInp.value = Number(thresholds.warn_tread_mm);
  if (minInp && thresholds.min_tread_mm != null) minInp.value = Number(thresholds.min_tread_mm);

  kpis.innerHTML = `
    <div class="kpi-card kpi-util">
      <div class="kpi-card-header"><div class="kpi-icon">T</div><div class="kpi-title">Active Tyres</div></div>
      <div class="kpi-big-value">${Number(s.active_tyres || 0)}</div>
      <div class="kpi-meta">${esc(data.asset.asset_code || "-")} — ${esc(data.asset.asset_name || "")}</div>
    </div>
    <div class="kpi-card kpi-alerts">
      <div class="kpi-card-header"><div class="kpi-icon">!</div><div class="kpi-title">End-of-life flags</div></div>
      <div class="kpi-big-value">${Number(s.replace_count || 0) + Number(s.warn_count || 0)}</div>
      <div class="kpi-meta">${Number(s.replace_count || 0)} replace · ${Number(s.warn_count || 0)} warn</div>
    </div>
    <div class="kpi-card kpi-avail">
      <div class="kpi-card-header"><div class="kpi-icon">$</div><div class="kpi-title">Fleet tyre $/hr</div></div>
      <div class="kpi-big-value">${s.fleet_cost_per_hour != null ? Number(s.fleet_cost_per_hour).toFixed(2) : "—"}</div>
      <div class="kpi-meta">Sum of position cost/hr · avg ${s.avg_cost_per_hour != null ? Number(s.avg_cost_per_hour).toFixed(2) : "—"}</div>
    </div>
  `;

  const alerts = Array.isArray(data.alerts) ? data.alerts : [];
  alertsEl.innerHTML = alerts.length
    ? `<div class="message-error" style="padding:8px 10px; border-radius:8px;">${alerts.map((a) =>
        `<div><strong>${esc(a.position_label || a.position_key)}</strong>: ${esc(a.tread_alert || a.lifecycle_status)}</div>`
      ).join("")}</div>`
    : `<div class="muted mini">No end-of-life tread alerts on active tyres.</div>`;

  const positions = Array.isArray(data.positions) ? data.positions : [];
  tableEl.innerHTML = positions.length
    ? `<table class="gridTable" style="min-width:980px;">
        <thead><tr>
          <th>Position</th><th>Serial</th><th>Status</th><th>Tread</th>
          <th>Installed</th><th>Hrs on tyre</th><th>Cost</th><th>$/hr</th><th>Last insp.</th>
        </tr></thead>
        <tbody>${positions.map((p) => `
          <tr>
            <td>${esc(p.position_label || p.position_key)}</td>
            <td>${esc(p.serial_number || "-")}</td>
            <td>${tyreStatusBadge(p.lifecycle_status)}</td>
            <td>${p.tread_depth == null ? "-" : Number(p.tread_depth).toFixed(1)}</td>
            <td>${esc(p.install_date || "-")}</td>
            <td>${Number(p.hours_on_tyre || 0).toFixed(1)}</td>
            <td>${Number(p.tyre_cost || 0).toFixed(2)}</td>
            <td>${p.cost_per_hour == null ? "-" : Number(p.cost_per_hour).toFixed(4)}</td>
            <td>${esc(p.last_inspection_date || "-")}</td>
          </tr>`).join("")}</tbody>
      </table>`
    : `<div class="empty">No active tyre installs yet — save an inspection to start lifecycle tracking.</div>`;

  const history = Array.isArray(data.change_history) ? data.change_history : [];
  historyEl.innerHTML = history.length
    ? `<table class="gridTable" style="min-width:920px;">
        <thead><tr>
          <th>Position</th><th>Serial</th><th>Installed</th><th>Removed</th>
          <th>Hrs on tyre</th><th>Cost</th><th>$/hr</th><th>Reason</th>
        </tr></thead>
        <tbody>${history.map((h) => `
          <tr>
            <td>${esc(h.position_label || h.position_key)}</td>
            <td>${esc(h.serial_number || "-")}</td>
            <td>${esc(h.install_date || "-")}</td>
            <td>${esc(h.removed_date || "-")}</td>
            <td>${Number(h.hours_on_tyre || 0).toFixed(1)}</td>
            <td>${Number(h.tyre_cost || 0).toFixed(2)}</td>
            <td>${h.cost_per_hour == null ? "-" : Number(h.cost_per_hour).toFixed(4)}</td>
            <td>${esc(String(h.removed_reason || "-").replace(/_/g, " "))}</td>
          </tr>`).join("")}</tbody>
      </table>`
    : `<div class="empty">No tyre changes recorded yet.</div>`;
}

async function loadTyreLifecycle() {
  const assetId = Number(document.getElementById("tyreAsset")?.value || 0);
  const panel = document.getElementById("tyreLifecyclePanel");
  if (!assetId) {
    if (panel) panel.style.display = "none";
    return;
  }
  try {
    const res = await fetch(`${API}/maintenance/tyre-inspections/lifecycle?asset_id=${assetId}`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load tyre lifecycle");
    renderTyreLifecyclePanel(data);
    prefillTyreFormFromLifecycle(data);
  } catch {
    if (panel) panel.style.display = "none";
  }
}

function prefillTyreFormFromLifecycle(data) {
  const positions = Array.isArray(data?.positions) ? data.positions : [];
  const byKey = new Map(positions.map((p) => [String(p.position_key || "").toLowerCase(), p]));
  TYRE_POSITIONS.forEach((p) => {
    const row = byKey.get(String(p.key).toLowerCase());
    if (!row) return;
    const serialEl = document.getElementById(tyreInputId(p.key, "serial"));
    const changedEl = document.getElementById(tyreInputId(p.key, "changed"));
    const costEl = document.getElementById(tyreInputId(p.key, "cost"));
    const treadEl = document.getElementById(tyreInputId(p.key, "rtd_outer"));
    const pressureEl = document.getElementById(tyreInputId(p.key, "pressure"));
    if (serialEl && !serialEl.value && row.serial_number) serialEl.value = row.serial_number;
    if (changedEl && !changedEl.value && row.install_date) changedEl.value = row.install_date;
    if (costEl && !costEl.value && row.tyre_cost > 0) costEl.value = Number(row.tyre_cost).toFixed(2);
    if (treadEl && !treadEl.value && row.tread_depth != null) treadEl.value = Number(row.tread_depth);
    if (pressureEl && !pressureEl.value && row.pressure != null) pressureEl.value = Number(row.pressure);
  });
}

function initTyreLayout() {
  const grid = document.getElementById("tyreGrid");
  if (!grid) return;
  grid.innerHTML = TYRE_POSITIONS.map((p) => `
    <div class="card tyre-pos-card" style="padding:10px;">
      <div style="font-weight:600; margin-bottom:8px;">${esc(p.label)} <span class="muted mini">(${esc(p.surveyCode)})</span></div>
      <div class="form-grid" style="grid-template-columns:repeat(2,minmax(0,1fr)); gap:8px;">
        <label>Tyre make<input id="${tyreInputId(p.key, "make")}" type="text" placeholder="e.g. Michelin" /></label>
        <label>Brand / size<input id="${tyreInputId(p.key, "brand")}" type="text" placeholder="e.g. 23.5R25" /></label>
        <label>Description<input id="${tyreInputId(p.key, "desc")}" type="text" placeholder="Model / pattern" /></label>
        <label>Serial number<input id="${tyreInputId(p.key, "serial")}" type="text" placeholder="Serial" /></label>
        <label>Pressure cold<input id="${tyreInputId(p.key, "pressure")}" type="number" min="0" step="0.1" placeholder="kPa" /></label>
        <label>Pressure recom.<input id="${tyreInputId(p.key, "pressure_recom")}" type="number" min="0" step="0.1" placeholder="kPa" /></label>
        <label>Pressure hot<input id="${tyreInputId(p.key, "pressure_hot")}" type="number" min="0" step="0.1" placeholder="kPa" /></label>
        <label>OTD (mm)<input id="${tyreInputId(p.key, "otd")}" type="number" min="0" step="0.1" placeholder="Original tread" /></label>
        <label>RTD outer (mm)<input id="${tyreInputId(p.key, "rtd_outer")}" type="number" min="0" step="0.1" placeholder="Outer RTD" /></label>
        <label>RTD inner (mm)<input id="${tyreInputId(p.key, "rtd_inner")}" type="number" min="0" step="0.1" placeholder="Inner RTD" /></label>
        <label>Date last changed<input id="${tyreInputId(p.key, "changed")}" type="date" /></label>
        <label>Purchase price<input id="${tyreInputId(p.key, "cost")}" type="number" min="0" step="0.01" placeholder="Cost" /></label>
      </div>
    </div>
  `).join("");
}

function tyreReadOptionalNumber(id) {
  const raw = String(document.getElementById(id)?.value || "").trim();
  if (raw === "") return null;
  const n = Number(raw);
  return Number.isFinite(n) ? n : null;
}

function collectTyrePositionRows() {
  return TYRE_POSITIONS.map((p) => {
    const tyre_cost = Number(document.getElementById(tyreInputId(p.key, "cost"))?.value || 0) || 0;
    const rtd_outer = tyreReadOptionalNumber(tyreInputId(p.key, "rtd_outer"));
    const rtd_inner = tyreReadOptionalNumber(tyreInputId(p.key, "rtd_inner"));
    return {
      position_key: p.key,
      position_label: p.label,
      survey_code: p.surveyCode,
      tyre_make: String(document.getElementById(tyreInputId(p.key, "make"))?.value || "").trim(),
      brand_number: String(document.getElementById(tyreInputId(p.key, "brand"))?.value || "").trim(),
      tyre_description: String(document.getElementById(tyreInputId(p.key, "desc"))?.value || "").trim(),
      pressure: tyreReadOptionalNumber(tyreInputId(p.key, "pressure")),
      pressure_recommended: tyreReadOptionalNumber(tyreInputId(p.key, "pressure_recom")),
      pressure_hot: tyreReadOptionalNumber(tyreInputId(p.key, "pressure_hot")),
      original_tread_depth: tyreReadOptionalNumber(tyreInputId(p.key, "otd")),
      rtd_outer,
      rtd_inner,
      tread_depth: rtd_outer,
      serial_number: String(document.getElementById(tyreInputId(p.key, "serial"))?.value || "").trim(),
      last_changed_date: String(document.getElementById(tyreInputId(p.key, "changed"))?.value || "").trim(),
      tyre_cost,
    };
  });
}

function clearTyreForm() {
  TYRE_POSITIONS.forEach((p) => {
    ["make", "brand", "desc", "pressure", "pressure_recom", "pressure_hot", "otd", "rtd_outer", "rtd_inner", "serial", "changed", "cost"].forEach((field) => {
      const el = document.getElementById(tyreInputId(p.key, field));
      if (el) el.value = "";
    });
  });
}

function tyreSurveyMonth() {
  const raw = String(document.getElementById("tyreSurveyMonth")?.value || "").trim();
  if (/^\d{4}-\d{2}$/.test(raw)) return raw;
  return new Date().toISOString().slice(0, 7);
}

async function openTyreSurveyExport(kind) {
  const msg = document.getElementById("tyreMsg");
  const month = tyreSurveyMonth();
  const q = new URLSearchParams();
  q.set("month", month);
  const path = kind === "xlsx"
    ? `${API}/maintenance/tyre-inspections/survey.xlsx?${q.toString()}`
    : `${API}/maintenance/tyre-inspections/survey.pdf?${q.toString()}`;
  if (msg) msg.textContent = `Generating survey ${kind.toUpperCase()}...`;
  try {
    const res = await fetch(path, { headers: authHeaders(), cache: "no-store" });
    if (!res.ok) {
      const txt = await res.text();
      throw new Error(txt || `Survey ${kind} failed (${res.status})`);
    }
    const blob = await res.blob();
    const blobUrl = URL.createObjectURL(blob);
    if (kind === "xlsx") {
      const a = document.createElement("a");
      a.href = blobUrl;
      a.download = `IRONLOG_Tyre_Survey_${month}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    } else {
      window.open(blobUrl, "_blank");
    }
    setTimeout(() => URL.revokeObjectURL(blobUrl), 15000);
    if (msg) msg.textContent = `Survey ${kind.toUpperCase()} ready for ${month}.`;
  } catch (e) {
    if (msg) msg.textContent = `Survey export failed: ${e.message || e}`;
  }
}

function renderTyreList(rows) {
  const list = document.getElementById("tyreList");
  if (!list) return;
  if (!rows.length) {
    list.innerHTML = `<div class="empty">No tyre inspections saved for this filter.</div>`;
    return;
  }
  list.innerHTML = rows.map((r) => `
    <div class="item">
      <div class="title">${esc(r.asset_code || "-")} - ${esc(r.asset_name || "")}</div>
      <div class="meta">
        Inspection ${esc(r.inspection_date || "-")} | Machine hours ${Number(r.running_hours || 0).toFixed(1)} |
        Fleet tyre $/hr <b>${Number(r.cost_per_running_hour || 0).toFixed(4)}</b>
        | Positions ${(Array.isArray(r.tyres) ? r.tyres : []).length}
      </div>
      <div style="overflow:auto; margin-top:8px;">
        <table class="gridTable" style="min-width:980px;">
          <thead>
            <tr>
              <th>Position</th><th>Make</th><th>Status</th><th>Pressure</th><th>RTD out/in</th><th>Serial</th>
              <th>Installed</th><th>Hrs on tyre</th><th>Cost</th><th>$/hr</th>
            </tr>
          </thead>
          <tbody>
            ${(Array.isArray(r.tyres) ? r.tyres : []).map((t) => `
              <tr>
                <td>${esc(t.position_label || "-")}</td>
                <td>${esc(t.tyre_make || "-")}</td>
                <td>${tyreStatusBadge(t.lifecycle_status)}</td>
                <td>${t.pressure == null ? "-" : Number(t.pressure).toFixed(0)}</td>
                <td>${t.rtd_outer == null && t.rtd_inner == null ? "-" : `${t.rtd_outer == null ? "-" : Number(t.rtd_outer).toFixed(1)} / ${t.rtd_inner == null ? "-" : Number(t.rtd_inner).toFixed(1)}`}</td>
                <td>${esc(t.serial_number || "-")}</td>
                <td>${esc(t.install_date || t.last_changed_date || "-")}</td>
                <td>${t.hours_on_tyre == null ? "-" : Number(t.hours_on_tyre).toFixed(1)}</td>
                <td>${Number(t.tyre_cost || 0).toFixed(2)}</td>
                <td>${t.cost_per_hour == null ? "-" : Number(t.cost_per_hour).toFixed(4)}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    </div>
  `).join("");
}

async function loadTyreInspections() {
  const list = document.getElementById("tyreList");
  if (list) list.innerHTML = `<div class="empty">Loading tyre inspections...</div>`;
  const filter = String(document.getElementById("tyreFilterAsset")?.value || "").trim();
  const q = new URLSearchParams();
  if (filter) q.set("asset_id", filter);
  try {
    const res = await fetch(`${API}/maintenance/tyre-inspections?${q.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load tyre inspections");
    const rows = Array.isArray(data?.rows) ? data.rows : [];
    renderTyreList(rows);
  } catch (e) {
    if (list) list.innerHTML = `<div class="empty">Load failed: ${esc(e.message || String(e))}</div>`;
  }
}

async function pullTyreLiveHours() {
  const assetId = Number(document.getElementById("tyreAsset")?.value || 0);
  const out = document.getElementById("tyreLiveHoursMeta");
  if (!assetId) {
    if (out) out.textContent = "Select an asset first.";
    return;
  }
  if (out) out.textContent = "Loading live hours...";
  try {
    const res = await fetch(`${API}/maintenance/asset/${assetId}/live-hours`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load live hours");
    const current = Number(data?.current_hours ?? 0);
    const inp = document.getElementById("tyreRunningHours");
    if (inp) inp.value = Number.isFinite(current) ? current.toFixed(1) : "";
    if (out) out.textContent = `Live hours: ${Number.isFinite(current) ? current.toFixed(1) : "0.0"} (${data?.source || "-"})`;
  } catch (e) {
    if (out) out.textContent = `Live hours error: ${e.message || e}`;
  }
}

async function saveTyreInspection() {
  const msg = document.getElementById("tyreMsg");
  const assetSel = document.getElementById("tyreAsset");
  const asset_id = Number(assetSel?.value || 0);
  const inspection_date = String(document.getElementById("tyreInspectionDate")?.value || "").trim();
  const runningRaw = String(document.getElementById("tyreRunningHours")?.value || "").trim();
  const running_hours = runningRaw === "" ? 0 : Number(runningRaw);
  if (!asset_id) {
    if (msg) msg.textContent = "Select an asset.";
    return;
  }
  if (!inspection_date) {
    if (msg) msg.textContent = "Select inspection date.";
    return;
  }
  if (!Number.isFinite(running_hours) || running_hours < 0) {
    if (msg) msg.textContent = "Running hours must be a valid positive number.";
    return;
  }
  const tyres = collectTyrePositionRows();
  const badNumber = tyres.some((t) => (t.pressure != null && !Number.isFinite(t.pressure))
    || (t.tread_depth != null && !Number.isFinite(t.tread_depth))
    || !Number.isFinite(Number(t.tyre_cost || 0)));
  if (badNumber) {
    if (msg) msg.textContent = "Please correct tyre pressure, tread, and cost values.";
    return;
  }
  const total_tyre_cost = Number(tyres.reduce((sum, t) => sum + Number(t.tyre_cost || 0), 0).toFixed(2));
  const divisor = running_hours > 0 ? running_hours : 1;
  const cost_per_running_hour = Number((total_tyre_cost / divisor).toFixed(4));
  if (msg) {
    msg.className = "muted";
    msg.textContent = "Saving...";
  }
  try {
    const res = await fetch(`${API}/maintenance/tyre-inspections`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({
        asset_id,
        inspection_date,
        running_hours: Number(running_hours.toFixed(1)),
        tyres,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save tyre inspection");
    await loadTyreInspections();
    await loadTyreLifecycle();
    const alertCount = Array.isArray(data.alerts) ? data.alerts.length : 0;
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `Saved. Fleet tyre $/hr ${Number(data.cost_per_running_hour || 0).toFixed(4)}${alertCount ? ` · ${alertCount} end-of-life flag(s)` : ""}.`;
    }
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Save failed: ${e.message || e}`;
    }
  }
}
