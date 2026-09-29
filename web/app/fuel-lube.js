// IRONLOG/web/app/fuel-lube.js — Lube usage, fuel log/benchmarks, cost settings, shift scenarios.
// Part of the main app; index.html loads these files in order and they share one global scope.

function getDefaultLubeRange() {
  const end = new Date();
  const start = new Date();
  start.setDate(end.getDate() - 29);
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

let lubeUsageCache = null;

function isPartLikeOilType(v) {
  const t = String(v || "").trim();
  if (!t) return true;
  const lower = t.toLowerCase();
  const knownOilWords = ["oil", "lube", "lub", "hyd", "hydraulic", "grease", "coolant", "atf", "engine"];
  if (knownOilWords.some((w) => lower.includes(w))) return false;
  // Typical stock/part-like key: uppercase-ish code with digits/hyphen/underscore and no spaces.
  return /^[a-z0-9][a-z0-9\-_/.]{2,24}$/i.test(t) && !/\s/.test(t);
}

function renderLubeUsageTable(payload) {
  const lubeList = qs("lubeList");
  if (!lubeList) return;
  const assetFilter = String(qs("lubeFilterAsset")?.value || "").trim().toLowerCase();
  const oilTypeFilter = String(qs("lubeFilterOilType")?.value || "").trim().toLowerCase();
  const hidePartLike = Boolean(qs("lubeHidePartLike")?.checked);

  const detailLines = Array.isArray(payload?.lines) ? payload.lines : [];
  let flat = [];

  if (detailLines.length) {
    flat = detailLines
      .filter((r) => {
        const code = String(r.asset_code || "").toLowerCase();
        const part = String(r.part_code || "").toLowerCase();
        const typ = String(r.lube_type || "").toLowerCase();
        const name = String(r.part_name || "").toLowerCase();
        if (assetFilter && !code.includes(assetFilter)) return false;
        if (oilTypeFilter && !typ.includes(oilTypeFilter) && !part.includes(oilTypeFilter) && !name.includes(oilTypeFilter)) {
          return false;
        }
        if (hidePartLike && isPartLikeOilType(r.part_code) && !typ.includes("oil") && !name.includes("oil")) return false;
        return true;
      })
      .map((r) => ({
        usage_date: String(r.usage_date || ""),
        asset_code: String(r.asset_code || ""),
        asset_name: String(r.asset_name || ""),
        part_code: String(r.part_code || ""),
        part_name: String(r.part_name || ""),
        lube_type: String(r.lube_type || "lube"),
        smr: r.smr != null && Number.isFinite(Number(r.smr)) ? Number(r.smr) : null,
        qty: Number(r.quantity || 0),
        cost: Number(r.line_cost || 0),
        source: String(r.source || ""),
        work_order_id: r.work_order_id != null ? Number(r.work_order_id) : null,
      }));
  } else {
    const dataRows = Array.isArray(payload?.rows) ? payload.rows : [];
    for (const r of dataRows) {
      const assetCode = String(r.asset_code || "");
      const assetName = String(r.asset_name || "");
      if (assetFilter && !assetCode.toLowerCase().includes(assetFilter)) continue;
      const byType = Array.isArray(r.by_oil_type) ? r.by_oil_type : [];
      const normalized = byType.length
        ? byType.map((x) => ({
            oil_type: String(x.oil_type || "UNSPECIFIED"),
            qty: Number(x.qty ?? x.qty_total ?? 0),
            cost: Number(x.lube_cost ?? x.total_lube_cost ?? 0),
          }))
        : [{
            oil_type: "UNSPECIFIED",
            qty: Number(r.qty ?? r.qty_total ?? 0),
            cost: Number(r.lube_cost ?? r.total_lube_cost ?? 0),
          }];

      for (const t of normalized) {
        if (hidePartLike && isPartLikeOilType(t.oil_type)) continue;
        if (oilTypeFilter && !String(t.oil_type || "").toLowerCase().includes(oilTypeFilter)) continue;
        flat.push({
          usage_date: "",
          asset_code: assetCode,
          asset_name: assetName,
          part_code: t.oil_type,
          part_name: t.oil_type,
          lube_type: t.oil_type,
          smr: null,
          qty: t.qty,
          cost: t.cost,
          source: "summary",
          work_order_id: null,
        });
      }
    }
  }

  if (!flat.length) {
    lubeList.innerHTML = `<small>No lube rows match your filters.</small>`;
    return;
  }

  const visibleAssets = new Set(flat.map((r) => r.asset_code)).size;
  const visibleQty = flat.reduce((s, r) => s + Number(r.qty || 0), 0);
  const visibleCost = flat.reduce((s, r) => s + Number(r.cost || 0), 0);

  lubeList.innerHTML = `
    <div style="overflow:auto;">
      <table class="gridTable" style="min-width:1320px;">
        <thead>
          <tr>
            <th>Date</th>
            <th>Lube part no</th>
            <th>Description</th>
            <th>Type</th>
            <th>Plant no</th>
            <th>Machine</th>
            <th style="text-align:right;">Machine hrs</th>
            <th style="text-align:right;">Qty</th>
            <th style="text-align:right;">Cost</th>
            <th>Source</th>
            <th>WO #</th>
          </tr>
        </thead>
        <tbody>
          ${flat.map((r) => `
            <tr>
              <td style="padding:10px 8px;">${escapeHtml(r.usage_date || "-")}</td>
              <td style="padding:10px 8px;">${escapeHtml(r.part_code || "-")}</td>
              <td style="padding:10px 8px;">${escapeHtml(r.part_name || "-")}</td>
              <td style="padding:10px 8px;">${escapeHtml(r.lube_type || "-")}</td>
              <td style="padding:10px 8px;">${escapeHtml(r.asset_code || "-")}</td>
              <td style="padding:10px 8px;">${escapeHtml(r.asset_name || r.asset_code || "-")}</td>
              <td style="padding:10px 8px; text-align:right;">${r.smr != null ? Number(r.smr).toFixed(1) : "-"}</td>
              <td style="padding:10px 8px; text-align:right;">${Number(r.qty || 0).toFixed(1)}</td>
              <td style="padding:10px 8px; text-align:right;">${fmtMoney(r.cost || 0)}</td>
              <td style="padding:10px 8px;">${escapeHtml(r.source || "-")}</td>
              <td style="padding:10px 8px;">${r.work_order_id ? escapeHtml(String(r.work_order_id)) : "-"}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
    <div class="muted" style="margin-top:8px;">
      Showing ${flat.length} line(s) across ${visibleAssets} machine(s) | qty ${visibleQty.toFixed(1)} | cost ${fmtMoney(visibleCost)}
    </div>
  `;
}

async function loadLubeUsage() {
  const start = qs("lubeStart")?.value || "";
  const end = qs("lubeEnd")?.value || "";
  if (!start || !end) {
    alert("Select start and end dates.");
    return;
  }

  setStatus("Loading lube usage...");
  setSkeleton("lubeList", 2);
  const data = await fetchJson(`${API}/api/dashboard/lube?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`);

  setText("lubeQtyTotal", Number(data.summary?.line_qty_total ?? data.summary?.qty_total ?? 0).toFixed(1));
  setText("lubeEntries", Number(data.summary?.line_entries ?? data.summary?.entries ?? 0));
  setText("lubeAssets", Number(data.summary?.line_assets ?? data.summary?.assets ?? 0));
  setText("cLubeCost", fmtMoney(data.summary?.line_total_cost ?? data.summary?.total_lube_cost ?? 0));
  lubeUsageCache = data;
  renderLubeUsageTable(data);

  setStatus("Lube usage ready.");
}

async function saveFuelLog() {
  const meterRaw = (qs("fuelHoursRun")?.value || "").trim();
  const meter_unit = String(qs("fuelMeterUnit")?.value || "hours").trim().toLowerCase() === "km" ? "km" : "hours";
  const cc = String(qs("fuelCostCenter")?.value || "").trim();
  const basePayload = {
    asset_code: (qs("fuelAsset")?.value || "").trim(),
    log_date: (qs("fuelDate")?.value || "").trim() || undefined,
    liters: Number(qs("fuelLiters")?.value || 0),
    meter_run_value: meterRaw === "" ? undefined : Number(meterRaw),
    meter_unit,
    hours_run: meter_unit === "hours" && meterRaw !== "" ? Number(meterRaw) : undefined,
    source: (qs("fuelSource")?.value || "").trim() || undefined,
    cost_center_code: cc || undefined,
  };

  setStatus("Saving fuel log...");
  try {
    let res;
    try {
      res = await fetchJson(`${API}/api/dashboard/fuel/log`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(basePayload),
      });
    } catch (e) {
      const msg = String(e?.message || e || "");
      const isDup = /possible_duplicate_recent/i.test(msg) || /duplicate/i.test(msg);
      if (!isDup) throw e;
      const ok = confirm("Possible duplicate fuel input detected (same values in last 60 seconds).\nSave it anyway?");
      if (!ok) {
        setStatus("Fuel save cancelled (duplicate protection).");
        return;
      }
      res = await fetchJson(`${API}/api/dashboard/fuel/log`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...basePayload, force_duplicate: true }),
      });
    }
    setText("fuelInputResult", JSON.stringify(res, null, 2));
    setStatus("Fuel log saved.");
    await Promise.all([
      loadDashboard().catch(() => {}),
      loadFuelBenchmark().catch(() => {}),
    ]);
  } catch (e) {
    setText("fuelInputResult", String(e.message || e));
    setStatus("Fuel log save failed.");
  }
}

function parseFuelMassText(text) {
  const raw = String(text || "").trim();
  if (!raw) return [];
  const lines = raw.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const rows = [];
  for (const line of lines) {
    const parts = line.includes("\t") ? line.split("\t") : line.split(",");
    const [asset_code, log_date, liters, meter_unit, meter_run_value, source] = parts.map((p) => String(p || "").trim());
    if (!asset_code || !log_date || !liters) continue;
    rows.push({
      asset_code,
      log_date,
      liters: Number(liters),
      meter_unit: String(meter_unit || "").toLowerCase() === "km" ? "km" : "hours",
      meter_run_value: meter_run_value === "" ? undefined : Number(meter_run_value),
      source: source || undefined,
    });
  }
  return rows;
}

async function importFuelMassPaste() {
  const txt = qs("fuelMassPaste")?.value || "";
  const out = qs("fuelMassResult");
  if (out) out.textContent = "";
  const rows = parseFuelMassText(txt);
  if (!rows.length) return alert("No valid rows found. Paste at least: asset_code,log_date,liters.");
  setStatus(`Importing ${rows.length} fuel rows...`);
  let ok = 0;
  let fail = 0;
  const errs = [];
  for (const r of rows) {
    try {
      await fetchJson(`${API}/api/dashboard/fuel/log`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          asset_code: r.asset_code,
          log_date: r.log_date,
          liters: r.liters,
          meter_unit: r.meter_unit,
          meter_run_value: r.meter_run_value,
          source: r.source,
          force_duplicate: true,
        }),
      });
      ok += 1;
    } catch (e) {
      fail += 1;
      errs.push({ row: r, error: String(e.message || e) });
    }
  }
  if (out) out.textContent = JSON.stringify({ ok, fail, errors: errs.slice(0, 30) }, null, 2);
  setStatus(`Fuel mass import done: ${ok} ok, ${fail} failed.`);
  await Promise.all([loadDashboard().catch(() => {}), loadFuelBenchmark().catch(() => {})]);
}

async function loadFuelBaseline() {
  const asset_code = (qs("fuelBaseAsset")?.value || "").trim();
  if (!asset_code) return alert("Enter/select asset code first.");

  setStatus("Loading OEM baseline...");
  try {
    const data = await fetchJson(
      `${API}/api/dashboard/fuel/baseline?asset_code=${encodeURIComponent(asset_code)}`
    );
    const mode = String(data.asset?.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
    const baseline = mode === "km"
      ? Number(data.asset?.baseline_fuel_km_per_l || 2)
      : Number(data.asset?.baseline_fuel_l_per_hour || 5);
    const input = qs("fuelBaseValue");
    if (input) input.value = baseline.toFixed(3);
    const unitEl = qs("fuelBaseUnitLabel");
    if (unitEl) unitEl.textContent = mode === "km" ? "km/L" : "L/hr";
    const mu = qs("fuelMeterUnit");
    if (mu) mu.value = mode === "km" ? "km" : "hours";
    const meterInput = qs("fuelHoursRun");
    if (meterInput) meterInput.placeholder = mode === "km" ? "Distance since fill (km)" : "Hours since fill";
    setText(
      "fuelBaselineResult",
      JSON.stringify(
        {
          asset_code: data.asset?.asset_code,
          asset_name: data.asset?.asset_name,
          metric_mode: mode,
          baseline_value: baseline,
        },
        null,
        2
      )
    );
    setStatus("OEM baseline loaded.");
  } catch (e) {
    setText("fuelBaselineResult", String(e.message || e));
    setStatus("OEM baseline load failed.");
  }
}

async function syncFuelUnitFromAsset(assetCode, target = "input") {
  const code = String(assetCode || "").trim();
  if (!code) return;
  try {
    const data = await fetchJson(`${API}/api/dashboard/fuel/baseline?asset_code=${encodeURIComponent(code)}`);
    const mode = String(data.asset?.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
    if (target === "input" || target === "both") {
      const mu = qs("fuelMeterUnit");
      if (mu) mu.value = mode === "km" ? "km" : "hours";
      const meterInput = qs("fuelHoursRun");
      if (meterInput) meterInput.placeholder = mode === "km" ? "Distance since fill (km)" : "Hours since fill";
    }
    if (target === "baseline" || target === "both") {
      const unitEl = qs("fuelBaseUnitLabel");
      if (unitEl) unitEl.textContent = mode === "km" ? "km/L" : "L/hr";
    }
  } catch {}
}

async function loadCostSettings() {
  setStatus("Loading cost defaults...");
  const data = await fetchJson(`${API}/api/dashboard/cost/settings`);
  const s = data?.settings || {};
  const put = (id, v) => {
    const el = qs(id);
    if (el && v != null) el.value = Number(v).toFixed(2);
  };
  put("costFuelDefault", s.fuel_cost_per_liter_default ?? 1.5);
  put("costLubeDefault", s.lube_cost_per_qty_default ?? 4.0);
  put("costLaborDefault", s.labor_cost_per_hour_default ?? 35.0);
  put("costDowntimeDefault", s.downtime_cost_per_hour_default ?? 120.0);
  setStatus("Cost defaults ready.");
}

async function saveCostSettings() {
  const payload = {
    fuel_cost_per_liter_default: Number(qs("costFuelDefault")?.value || 0),
    lube_cost_per_qty_default: Number(qs("costLubeDefault")?.value || 0),
    labor_cost_per_hour_default: Number(qs("costLaborDefault")?.value || 0),
    downtime_cost_per_hour_default: Number(qs("costDowntimeDefault")?.value || 0),
  };
  setStatus("Saving cost defaults...");
  try {
    const res = await fetchJson(`${API}/api/dashboard/cost/settings`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("costSetupResult", JSON.stringify(res, null, 2));
    setStatus("Cost defaults saved.");
    await loadDashboard().catch(() => {});
  } catch (e) {
    setText("costSetupResult", String(e.message || e));
    setStatus("Cost defaults save failed.");
  }
}

async function saveCostAssetRates() {
  const payload = {
    asset_code: (qs("costAssetCode")?.value || "").trim(),
    fuel_cost_per_liter: (qs("costAssetFuel")?.value || "").trim(),
    downtime_cost_per_hour: (qs("costAssetDowntime")?.value || "").trim(),
    utilization_mode: (qs("costAssetUtilMode")?.value || "").trim(),
    km_per_hour_factor: (qs("costAssetKmFactor")?.value || "").trim(),
  };
  setStatus("Saving asset cost rates...");
  try {
    const res = await fetchJson(`${API}/api/dashboard/cost/asset-rates`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("costSetupResult", JSON.stringify(res, null, 2));
    setStatus("Asset cost rates saved.");
    await loadDashboard().catch(() => {});
  } catch (e) {
    setText("costSetupResult", String(e.message || e));
    setStatus("Asset cost rates save failed.");
  }
}

async function saveCostPartRate() {
  const payload = {
    part_code: (qs("costPartCode")?.value || "").trim(),
    unit_cost: Number(qs("costPartUnit")?.value || 0),
  };
  setStatus("Saving part unit cost...");
  try {
    const res = await fetchJson(`${API}/api/dashboard/cost/part-cost`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("costSetupResult", JSON.stringify(res, null, 2));
    setStatus("Part unit cost saved.");
    await loadDashboard().catch(() => {});
  } catch (e) {
    setText("costSetupResult", String(e.message || e));
    setStatus("Part unit cost save failed.");
  }
}

async function saveFuelBaseline() {
  const mode = String(qs("fuelMeterUnit")?.value || "hours").trim().toLowerCase() === "km" ? "km" : "hours";
  const payload = {
    asset_code: (qs("fuelBaseAsset")?.value || "").trim(),
    metric_mode: mode,
    ...(mode === "km"
      ? { baseline_fuel_km_per_l: Number(qs("fuelBaseValue")?.value || 0) }
      : { baseline_fuel_l_per_hour: Number(qs("fuelBaseValue")?.value || 0) }),
  };
  if (!payload.asset_code) return alert("Enter/select asset code first.");

  setStatus("Saving OEM baseline...");
  try {
    const res = await fetchJson(`${API}/api/dashboard/fuel/baseline`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("fuelBaselineResult", JSON.stringify(res, null, 2));
    setStatus("OEM baseline saved.");
    await loadFuelBenchmark().catch(() => {});
  } catch (e) {
    setText("fuelBaselineResult", String(e.message || e));
    setStatus("OEM baseline save failed.");
  }
}

function fuelSvgPeriodCompareLines(currentSeries, previousSeries, currentRange, previousRange) {
  const n = Math.max(
    Array.isArray(currentSeries) ? currentSeries.length : 0,
    Array.isArray(previousSeries) ? previousSeries.length : 0,
    1,
  );
  const cur = Array.isArray(currentSeries) ? currentSeries : [];
  const prev = Array.isArray(previousSeries) ? previousSeries : [];
  const allVals = [];
  for (let i = 0; i < n; i++) {
    allVals.push(Number(cur[i]?.liters || 0), Number(prev[i]?.liters || 0));
  }
  const max = Math.max(...allVals, 1);
  // Wider + taller for quarter-length ranges; scroll horizontally when needed
  const pxPerDay = n > 60 ? 14 : n > 31 ? 18 : n > 14 ? 22 : 28;
  const width = Math.max(720, Math.min(2200, 72 + n * pxPerDay));
  const height = n > 45 ? 460 : 420;
  const left = 52;
  const right = 18;
  const top = 28;
  const bottom = 42;
  const plotW = width - left - right;
  const plotH = height - top - bottom;
  const xAt = (i) => left + (n === 1 ? plotW / 2 : (i * plotW) / (n - 1));
  const yAt = (v) => top + plotH - (Number(v || 0) / max) * plotH;
  const showValueLabels = n <= 14;
  const dotR = n > 60 ? 2 : n > 31 ? 2.5 : 3.5;

  const linePath = (series, color) => {
    const pts = [];
    for (let i = 0; i < n; i++) {
      const v = Number(series[i]?.liters || 0);
      pts.push({ x: xAt(i), y: yAt(v), v });
    }
    const d = pts.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x.toFixed(1)} ${p.y.toFixed(1)}`).join(" ");
    const dots = pts.map((p) => {
      const label = showValueLabels
        ? `<text x="${p.x.toFixed(1)}" y="${(p.y - 8).toFixed(1)}" text-anchor="middle" font-size="10" fill="#334155">${p.v.toFixed(0)}</text>`
        : "";
      return `<circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="${dotR}" fill="${color}"/>${label}`;
    }).join("");
    return `<path d="${d}" fill="none" stroke="${color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>${dots}`;
  };

  const labelStep = n > 60 ? Math.ceil(n / 10) : n > 31 ? Math.ceil(n / 9) : n > 14 ? Math.ceil(n / 8) : 1;
  const xLabels = [];
  for (let i = 0; i < n; i++) {
    if (i % labelStep !== 0 && i !== n - 1 && i !== 0) continue;
    const rawDate = String(cur[i]?.date || prev[i]?.date || "").slice(0, 10);
    const label = /^\d{4}-\d{2}-\d{2}$/.test(rawDate)
      ? `${rawDate.slice(5, 7)}-${rawDate.slice(8, 10)}`
      : `D${i + 1}`;
    xLabels.push(`<text x="${xAt(i).toFixed(1)}" y="${top + plotH + 22}" text-anchor="middle" font-size="11" fill="#64748b">${label}</text>`);
  }

  const yTicks = [0, 0.25, 0.5, 0.75, 1].map((t) => {
    const v = max * t;
    const y = yAt(v);
    return `
      <line x1="${left}" y1="${y.toFixed(1)}" x2="${width - right}" y2="${y.toFixed(1)}" stroke="#e2e8f0" stroke-width="1"/>
      <text x="${left - 8}" y="${(y + 4).toFixed(1)}" text-anchor="end" font-size="11" fill="#64748b">${v.toFixed(0)}</text>
    `;
  }).join("");

  const curLabel = currentRange?.start && currentRange?.end
    ? `${currentRange.start} → ${currentRange.end}`
    : "Current";
  const prevLabel = previousRange?.start && previousRange?.end
    ? `${previousRange.start} → ${previousRange.end}`
    : "Previous";

  return `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="100%" height="${height}" role="img" aria-label="Fuel usage current versus previous period" preserveAspectRatio="xMinYMid meet">
      <rect x="0" y="0" width="${width}" height="${height}" fill="#ffffff" rx="8"/>
      <text x="${left}" y="16" font-size="12" fill="#475569">${escapeHtml(curLabel)} vs ${escapeHtml(prevLabel)} (liters per day)</text>
      ${yTicks}
      <line x1="${left}" y1="${top + plotH}" x2="${width - right}" y2="${top + plotH}" stroke="#cbd5e1" stroke-width="1"/>
      <line x1="${left}" y1="${top}" x2="${left}" y2="${top + plotH}" stroke="#cbd5e1" stroke-width="1"/>
      ${linePath(prev, "#94a3b8")}
      ${linePath(cur, "#2563eb")}
      ${xLabels}
    </svg>
  `;
}

function mountFuelCompareSvg(host, svgMarkup) {
  if (!host) return false;
  const markup = String(svgMarkup || "").trim();
  if (!markup) return false;
  const chartH = "420px";
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(markup, "image/svg+xml");
    if (doc.querySelector("parsererror")) throw new Error("svg parse error");
    const svg = doc.documentElement;
    if (!svg || String(svg.tagName || "").toLowerCase() !== "svg") throw new Error("svg root missing");
    const node = host.ownerDocument.importNode(svg, true);
    const h = Number(node.getAttribute("height") || 420) || 420;
    node.setAttribute("width", "100%");
    node.setAttribute("height", String(h));
    node.style.display = "block";
    node.style.width = "100%";
    node.style.height = `${h}px`;
    node.style.minHeight = chartH;
    host.replaceChildren(node);
    return true;
  } catch {
    host.innerHTML = markup;
    const svg = host.querySelector("svg");
    if (svg) {
      const h = Number(svg.getAttribute("height") || 420) || 420;
      svg.style.display = "block";
      svg.style.width = "100%";
      svg.style.height = `${h}px`;
    }
    return Boolean(svg);
  }
}

function fuelEquipChartRows(rows) {
  return (Array.isArray(rows) ? rows : [])
    .filter((r) => String(r.metric_mode || "hours").toLowerCase() !== "km")
    .filter((r) => Number(r.actual_lph || 0) > 0 || Number(r.oem_lph || 0) > 0)
    .map((r) => {
      const oem = Number(r.oem_lph || 0);
      const threshold = Number(
        r.excessive_threshold_lph != null
          ? r.excessive_threshold_lph
          : r.threshold_lph != null
            ? r.threshold_lph
            : oem > 0
              ? oem * 1.15
              : 0
      );
      return {
        asset_code: String(r.asset_code || "").trim(),
        label: String(r.asset_name || r.asset_code || "").trim() || String(r.asset_code || "Unknown"),
        actual: Number(r.actual_lph || 0),
        oem,
        threshold,
      };
    })
    .sort((a, b) => a.label.localeCompare(b.label));
}

function getSelectedFuelEquipmentCodes() {
  const host = qs("fuelEquipFilterList");
  if (!host) return null;
  const boxes = [...host.querySelectorAll('input[type="checkbox"][data-fuel-equip]')];
  if (!boxes.length) return [];
  return boxes.filter((b) => b.checked).map((b) => String(b.getAttribute("data-fuel-equip") || "").trim().toLowerCase());
}

/**
 * Derive a normalised equipment type from an asset label by stripping trailing
 * unit identifiers: "#2", "No.3", "(2)", serial numbers, etc.
 * Returns e.g. "CAT 950K" from "CAT 950K #2" or "Komatsu PC210" from "Komatsu PC210 No.3".
 */
function fuelEquipTypeGroup(label) {
  let s = String(label || "").trim();
  // Strip trailing patterns like "#1", "No.1", "No 1", "(1)", "- 1", "Unit 1", serial-like all-caps codes
  s = s
    .replace(/\s*[-–]?\s*#\s*\d+\s*$/i, "")
    .replace(/\s+No\.?\s*\d+\s*$/i, "")
    .replace(/\s*\(\s*\d+\s*\)\s*$/i, "")
    .replace(/\s+Unit\s+\d+\s*$/i, "")
    .replace(/\s+[A-Z]{2,}\d{4,}\s*$/i, "")  // strip trailing serial-like codes
    .trim();
  // Normalise multiple spaces
  return s.replace(/\s+/g, " ") || label;
}

function renderFuelEquipmentFilter(rows, preserveSelection = true) {
  const host = qs("fuelEquipFilterList");
  if (!host) return;
  const all = fuelEquipChartRows(rows);
  if (!all.length) {
    host.className = "fuel-equip-filter-list muted";
    host.textContent = "No machine (L/hr) equipment in this period.";
    // Clear type dropdown too
    const typeSel = qs("fuelEquipTypeFilter");
    if (typeSel) typeSel.innerHTML = `<option value="">Filter by type…</option>`;
    return;
  }
  const prev = preserveSelection ? new Set(getSelectedFuelEquipmentCodes() || []) : null;
  const selectAllByDefault = !prev || prev.size === 0;
  host.className = "fuel-equip-filter-list";
  host.innerHTML = all.map((r) => {
    const code = r.asset_code;
    const checked = selectAllByDefault || prev.has(code.toLowerCase()) ? "checked" : "";
    const title = escapeHtml(`${r.label} (${code})`);
    const typeGroup = fuelEquipTypeGroup(r.label);
    return `<label class="fuel-equip-filter-item" title="${title}" data-fuel-type="${escapeHtml(typeGroup)}">
      <input type="checkbox" data-fuel-equip="${escapeHtml(code)}" ${checked} />
      <span>${escapeHtml(r.label)}</span>
    </label>`;
  }).join("");

  // Populate type dropdown — unique groups sorted, only show if >1 group or >1 member in a group
  const typeSel = qs("fuelEquipTypeFilter");
  if (typeSel) {
    const groups = new Map();
    all.forEach((r) => {
      const g = fuelEquipTypeGroup(r.label);
      if (!groups.has(g)) groups.set(g, []);
      groups.get(g).push(r.asset_code);
    });
    // Only offer groups that have ≥2 members (singles are selected individually)
    const multiGroups = [...groups.entries()].filter(([, codes]) => codes.length >= 2).sort((a, b) => a[0].localeCompare(b[0]));
    typeSel.innerHTML = `<option value="">Filter by type…</option>` +
      multiGroups.map(([g]) => `<option value="${escapeHtml(g)}">${escapeHtml(g)} (${groups.get(g).length})</option>`).join("");
  }
}

function fuelSvgEquipmentConsumption(points) {
  const rows = Array.isArray(points) ? points : [];
  const n = Math.max(rows.length, 1);
  const vals = [];
  for (const r of rows) {
    vals.push(Number(r.actual || 0), Number(r.oem || 0), Number(r.threshold || 0));
  }
  const rawMax = Math.max(...vals, 1);
  const max = Math.ceil(rawMax / 5) * 5 || 5;
  const slot = Math.max(70, Math.min(110, 900 / Math.max(n, 1)));
  const width = Math.max(720, 80 + n * slot);
  const height = 440;
  const left = 52;
  const right = 20;
  const top = 36;
  const bottom = 88;
  const plotW = width - left - right;
  const plotH = height - top - bottom;
  const band = plotW / n;
  const barW = Math.min(42, Math.max(18, band * 0.45));
  const xCenter = (i) => left + band * i + band / 2;
  const yAt = (v) => top + plotH - (Number(v || 0) / max) * plotH;

  const bars = rows.map((r, i) => {
    const x = xCenter(i) - barW / 2;
    const y = yAt(r.actual);
    const h = Math.max(0, top + plotH - y);
    const labelY = h > 18 ? y + 14 : y - 6;
    const labelFill = h > 18 ? "#ffffff" : "#111827";
    return `
      <rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW.toFixed(1)}" height="${h.toFixed(1)}" fill="#4b5563" rx="2"/>
      <text x="${xCenter(i).toFixed(1)}" y="${labelY.toFixed(1)}" text-anchor="middle" font-size="11" font-weight="600" fill="${labelFill}">${Number(r.actual || 0).toFixed(0)}</text>
    `;
  }).join("");

  const linePath = (key, color) => {
    const pts = rows.map((r, i) => ({ x: xCenter(i), y: yAt(r[key]), v: Number(r[key] || 0) }));
    const d = pts.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x.toFixed(1)} ${p.y.toFixed(1)}`).join(" ");
    const dots = pts.map((p) => `
      <circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="4" fill="${color}" stroke="#fff" stroke-width="1"/>
      <text x="${p.x.toFixed(1)}" y="${(p.y - 10).toFixed(1)}" text-anchor="middle" font-size="11" font-weight="600" fill="#111827">${p.v.toFixed(0)}</text>
    `).join("");
    return `<path d="${d}" fill="none" stroke="${color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>${dots}`;
  };

  const xLabels = rows.map((r, i) => {
    const label = String(r.label || r.asset_code || "").slice(0, 18);
    return `<text x="${xCenter(i).toFixed(1)}" y="${top + plotH + 18}" text-anchor="end" font-size="11" fill="#334155" transform="rotate(-32 ${xCenter(i).toFixed(1)} ${top + plotH + 18})">${escapeHtml(label)}</text>`;
  }).join("");

  const yTicks = [];
  const step = max <= 20 ? 5 : max <= 50 ? 5 : 10;
  for (let v = 0; v <= max + 0.001; v += step) {
    const y = yAt(v);
    yTicks.push(`
      <line x1="${left}" y1="${y.toFixed(1)}" x2="${width - right}" y2="${y.toFixed(1)}" stroke="#e5e7eb" stroke-width="1"/>
      <text x="${left - 8}" y="${(y + 4).toFixed(1)}" text-anchor="end" font-size="11" fill="#64748b">${v}</text>
    `);
  }

  return `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="100%" height="${height}" role="img" aria-label="Equipment fuel consumptions" preserveAspectRatio="xMinYMid meet">
      <rect x="0" y="0" width="${width}" height="${height}" fill="#ffffff" rx="8"/>
      <text x="${left}" y="18" font-size="14" font-weight="700" fill="#111827">Equipment Fuel Consumptions</text>
      <text x="${left}" y="34" font-size="11" fill="#64748b">L/hr — Actual (bar) · OEM (green) · Threshold (red)</text>
      ${yTicks.join("")}
      <line x1="${left}" y1="${top + plotH}" x2="${width - right}" y2="${top + plotH}" stroke="#94a3b8" stroke-width="1"/>
      <line x1="${left}" y1="${top}" x2="${left}" y2="${top + plotH}" stroke="#94a3b8" stroke-width="1"/>
      ${bars}
      ${linePath("oem", "#16a34a")}
      ${linePath("threshold", "#dc2626")}
      ${xLabels}
    </svg>
  `;
}

function aggregateFuelByType(points) {
  const groups = new Map();
  for (const r of points) {
    const g = fuelEquipTypeGroup(r.label);
    if (!groups.has(g)) groups.set(g, []);
    groups.get(g).push(r);
  }
  const avg = (arr, key) => arr.reduce((s, r) => s + Number(r[key] || 0), 0) / arr.length;
  return [...groups.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([g, members]) => ({
      asset_code: g,
      label: members.length > 1 ? `${g} (avg ${members.length})` : g,
      actual: avg(members, "actual"),
      oem: avg(members, "oem"),
      threshold: avg(members, "threshold"),
      count: members.length,
    }));
}

function renderFuelEquipmentChart(rows) {
  const host = qs("fuelEquipChart");
  const summaryEl = qs("fuelEquipChartSummary");
  const wrap = qs("fuelEquipChartWrap");
  if (!host) return;

  const all = fuelEquipChartRows(rows);
  if (!all.length) {
    if (summaryEl) summaryEl.textContent = "No machine (L/hr) rows to chart for this period.";
    host.className = "fuel-equip-chart muted";
    host.textContent = "No L/hr equipment data.";
    return;
  }

  const selected = new Set(getSelectedFuelEquipmentCodes() || []);
  const points = selected.size
    ? all.filter((r) => selected.has(String(r.asset_code || "").toLowerCase()))
    : [];

  if (!points.length) {
    if (summaryEl) summaryEl.textContent = "Select one or more machines in the equipment filter to show the chart.";
    host.className = "fuel-equip-chart muted";
    host.textContent = "No equipment selected.";
    if (wrap) wrap.style.display = "";
    return;
  }

  const viewMode = String(qs("fuelEquipViewMode")?.value || "individual");
  const chartPoints = viewMode === "type" ? aggregateFuelByType(points) : points;

  if (summaryEl) {
    if (viewMode === "type") {
      const typeCount = chartPoints.length;
      const unitCount = points.length;
      summaryEl.textContent = `Showing ${typeCount} type${typeCount !== 1 ? "s" : ""} averaged from ${unitCount} unit${unitCount !== 1 ? "s" : ""}. Actual L/hr (bar) vs OEM (green) and threshold (red).`;
    } else {
      summaryEl.textContent = `Showing ${points.length} of ${all.length} machines (L/hr). Actual bars vs OEM (green) and threshold (red).`;
    }
  }
  if (wrap) wrap.style.display = "";
  host.className = "fuel-equip-chart";
  const svgMarkup = fuelSvgEquipmentConsumption(chartPoints);
  const mounted = mountFuelCompareSvg(host, svgMarkup);
  if (!mounted) {
    host.className = "fuel-equip-chart muted";
    host.textContent = "Chart could not be rendered.";
  }
}

function refreshFuelEquipmentChartFromStore() {
  renderFuelEquipmentFilter(window.__fuelBenchmarkChartRows || [], true);
  renderFuelEquipmentChart(window.__fuelBenchmarkChartRows || []);
}

function renderFuelPeriodCompareChart(data) {
  const host = qs("fuelPeriodCompareChart");
  const summaryEl = qs("fuelPeriodCompareSummary");
  const wrap = qs("fuelPeriodCompareWrap");
  if (!host) return;

  if (!data?.ok) {
    if (summaryEl) summaryEl.textContent = "Could not load period comparison.";
    host.className = "fuel-period-compare-chart muted";
    host.textContent = "No comparison data.";
    return;
  }

  const current = data.current || {};
  const previous = data.previous || {};
  const delta = data.delta || {};
  const curTotal = Number(current.total_liters || 0);
  const prevTotal = Number(previous.total_liters || 0);
  const deltaLiters = Number(delta.liters || 0);
  const deltaPct = delta.pct;
  const sign = deltaLiters >= 0 ? "+" : "";

  if (summaryEl) {
    summaryEl.innerHTML =
      `<strong>Current:</strong> ${curTotal.toFixed(1)} L (${escapeHtml(current.start || "")} to ${escapeHtml(current.end || "")})` +
      ` · <strong>Previous:</strong> ${prevTotal.toFixed(1)} L (${escapeHtml(previous.start || "")} to ${escapeHtml(previous.end || "")})` +
      ` · <strong>Change:</strong> <span class="${deltaLiters > 0 ? "pill red" : deltaLiters < 0 ? "pill blue" : "pill"}">${sign}${deltaLiters.toFixed(1)} L` +
      `${deltaPct == null ? "" : ` (${sign}${Number(deltaPct).toFixed(1)}%)`}</span></span>`;
  }

  if (wrap) wrap.style.display = "";
  host.className = "fuel-period-compare-chart";
  const svgMarkup = fuelSvgPeriodCompareLines(
    current.series,
    previous.series,
    { start: current.start, end: current.end },
    { start: previous.start, end: previous.end },
  );
  const mounted = mountFuelCompareSvg(host, svgMarkup);
  if (!mounted) {
    host.className = "fuel-period-compare-chart muted";
    host.textContent = "Chart could not be rendered for this period.";
  }
}

function setFuelPeriodDates(start, end) {
  if (qs("fuelStart")) qs("fuelStart").value = start;
  if (qs("fuelEnd")) qs("fuelEnd").value = end;
}

function applyFuelPeriodPreset(kind) {
  const today = new Date();
  const y = today.getUTCFullYear();
  const m = today.getUTCMonth() + 1;
  const d = today.getUTCDate();
  const pad = (n) => String(n).padStart(2, "0");
  const todayStr = `${y}-${pad(m)}-${pad(d)}`;
  const clip = (end) => (end > todayStr ? todayStr : end);
  let start = todayStr;
  let end = todayStr;
  if (kind === "q1") {
    start = `${y}-01-01`;
    end = clip(`${y}-03-31`);
  } else if (kind === "q2") {
    start = `${y}-04-01`;
    end = clip(`${y}-06-30`);
  } else if (kind === "q3") {
    start = `${y}-07-01`;
    end = clip(`${y}-09-30`);
  } else if (kind === "ytd") {
    start = `${y}-01-01`;
    end = todayStr;
  } else if (kind === "mtd") {
    start = `${y}-${pad(m)}-01`;
    end = todayStr;
  }
  setFuelPeriodDates(start, end);
  return loadFuelBenchmark();
}

async function loadFuelBenchmark() {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const mode = String(qs("fuelModeFilter")?.value || "").trim();
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  const duplicatesOnly = Boolean(qs("fuelDupOnly")?.checked);
  const runToken = Date.now() + Math.random();
  window.__fuelBenchmarkRunToken = runToken;
  if (!start || !end) return alert("Select start and end dates.");

  setStatus("Loading fuel benchmark...");
  setSkeleton("fuelBenchmarkList", 2);
  const compareUrl =
    `${API}/api/dashboard/fuel/period-compare?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&mode=${encodeURIComponent(mode)}&asset_code=${encodeURIComponent(assetCode)}`;
  const [data, compareData] = await Promise.all([
    duplicatesOnly
      ? fetchJson(
        `${API}/api/dashboard/fuel/duplicates?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&mode=${encodeURIComponent(mode)}&asset_code=${encodeURIComponent(assetCode)}`
      )
      : fetchJson(
        `${API}/api/dashboard/fuel?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}&mode=${encodeURIComponent(mode)}&asset_code=${encodeURIComponent(assetCode)}`
      ),
    duplicatesOnly ? Promise.resolve(null) : fetchJson(compareUrl).catch(() => null),
  ]);
  if (window.__fuelBenchmarkRunToken !== runToken) return;

  const rawRows = Array.isArray(data.rows) ? data.rows : [];
  const benchmarkRows = rawRows;
  const displayRows = benchmarkRows;

  const benchmarkSummary = duplicatesOnly
    ? {
        fuel_liters: Number(data.summary?.fuel_liters || 0),
        duplicate_rows: Number(data.summary?.duplicate_rows || 0),
      }
    : (data.summary && data.summary.avg_lph != null
      ? {
          fuel_liters: Number(data.summary.fuel_liters || 0),
          hours_run: Number(data.summary.hours_run || 0),
          km_run: Number(data.summary.km_run || 0),
          avg_lph: data.summary.avg_lph == null ? null : Number(data.summary.avg_lph),
          avg_km_per_l: data.summary.avg_km_per_l == null ? null : Number(data.summary.avg_km_per_l),
          excessive_count: Number(data.summary.excessive_count || data.summary.excessive || 0),
        }
      : benchmarkRows.reduce(
        (acc, r) => {
          const mode = String(r.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
          const fuel = Number(r.fuel_liters || 0);
          const hours = Number(r.hours_run || 0);
          const km = Number(r.km_run || 0);
          acc.fuel_liters += fuel;
          if (mode === "km") {
            acc.km_run += km;
            if (km > 0) acc.km_fuel += fuel;
          } else {
            acc.hours_run += hours;
            if (hours > 0) acc.hours_fuel += fuel;
          }
          if (Number(r.is_excessive || false)) acc.excessive_count += 1;
          return acc;
        },
        { fuel_liters: 0, hours_run: 0, km_run: 0, hours_fuel: 0, km_fuel: 0, excessive_count: 0 }
      ));
  if (!duplicatesOnly && benchmarkSummary.avg_lph === undefined) {
    benchmarkSummary.avg_lph =
      benchmarkSummary.hours_run > 0
        ? benchmarkSummary.hours_fuel / benchmarkSummary.hours_run
        : null;
    benchmarkSummary.avg_km_per_l =
      benchmarkSummary.km_fuel > 0
        ? benchmarkSummary.km_run / benchmarkSummary.km_fuel
        : null;
  }

  if (duplicatesOnly) {
    setText("fbFuelTotal", Number(benchmarkSummary.fuel_liters || 0).toFixed(2));
    setText("fbHoursTotal", "-");
    setText("fbKmTotal", "-");
    setText("fbAvgLph", "-");
    setText("fbAvgKmpl", "-");
    setText("fbExcessive", Number(benchmarkSummary.duplicate_rows || 0));
    const compareHost = qs("fuelPeriodCompareChart");
    const compareSummary = qs("fuelPeriodCompareSummary");
    if (compareSummary) compareSummary.textContent = "Period comparison is hidden while viewing duplicates.";
    if (compareHost) {
      compareHost.className = "fuel-period-compare-chart muted";
      compareHost.textContent = "Switch off duplicates filter to view period comparison.";
    }
    window.__fuelBenchmarkChartRows = [];
    const equipHost = qs("fuelEquipChart");
    const equipFilter = qs("fuelEquipFilterList");
    const equipSummary = qs("fuelEquipChartSummary");
    if (equipSummary) equipSummary.textContent = "Equipment chart is hidden while viewing duplicates.";
    if (equipFilter) {
      equipFilter.className = "fuel-equip-filter-list muted";
      equipFilter.textContent = "Switch off duplicates filter to select equipment.";
    }
    if (equipHost) {
      equipHost.className = "fuel-equip-chart muted";
      equipHost.textContent = "Switch off duplicates filter to view equipment chart.";
    }
  } else {
    // Pills must reflect only the currently selected date range + filters.
    const s = data.summary || {};
    setText("fbFuelTotal", Number(s.fuel_liters || 0).toFixed(2));
    setText("fbHoursTotal", Number(s.hours_run || 0).toFixed(2));
    setText("fbKmTotal", Number(s.km_run || 0).toFixed(2));
    setText("fbAvgLph", s.avg_lph == null ? "-" : Number(s.avg_lph).toFixed(3));
    setText("fbAvgKmpl", s.avg_km_per_l == null ? "-" : Number(s.avg_km_per_l).toFixed(3));
    setText("fbExcessive", Number(s.excessive_count || 0));
    renderFuelPeriodCompareChart(compareData);
    window.__fuelBenchmarkChartRows = benchmarkRows;
    renderFuelEquipmentFilter(benchmarkRows, false);
    renderFuelEquipmentChart(benchmarkRows);
  }

  const list = qs("fuelBenchmarkList");
  if (list) {
    const isRowFlagged = (r) => {
      const mode = String(r?.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
      if (mode === "km") {
        const actual = Number(r?.actual_km_per_l);
        const benchmark = Number(r?.oem_km_per_l);
        if (!Number.isFinite(actual) || actual <= 0) return false; // no entries -> OK
        if (!Number.isFinite(benchmark) || benchmark <= 0) return false;
        return actual < benchmark;
      }
      const actual = Number(r?.actual_lph);
      const benchmark = Number(r?.oem_lph);
      if (!Number.isFinite(actual) || actual <= 0) return false; // no entries -> OK
      if (!Number.isFinite(benchmark) || benchmark <= 0) return false;
      return actual > benchmark;
    };
    list.innerHTML = "";
    displayRows.forEach((r) => {
      if (duplicatesOnly) {
        const runVal = Number(r.meter_run_value ?? r.hours_run ?? 0);
        const runTxt = String(r.metric_mode || "hours") === "km"
          ? `Meter: ${runVal.toFixed(2)} km`
          : `Meter: ${runVal.toFixed(2)} hours`;
        list.appendChild(
          item(
            `<div class="fuel-item-head"><b>${r.asset_code}</b> — ${r.log_date} <span class='pill red'>DUPLICATE x${Number(r.duplicate_count || 0)}</span></div>` +
            `<small class="fuel-item-desc">${r.asset_name || ""}</small>` +
            `<small class="fuel-item-meta">Fuel: ${Number(r.liters || 0).toFixed(2)}L | ${runTxt} | Source: ${r.source || "-"}</small>` +
            `<br><button data-fuel-delete="${Number(r.id || 0)}">Delete this entry</button>`
          )
        );
        return;
      }
      const isKm = String(r.metric_mode || "hours") === "km";
      const hiredTag = r.is_hired ? " <span class='pill' style='font-size:0.65rem;'>HIRED</span>" : "";
      const archTag = Number(r.archived) ? " <span class='pill' style='font-size:0.65rem;'>ARCH</span>" : "";
      const modeTag = `<span class='pill blue' style='font-size:0.65rem;'>${isKm ? "km/L" : "L/hr"}</span>`;
      const runSrc = r.run_source && r.run_source !== "none"
        ? ` | Run: ${isKm ? `${Number(r.km_run || 0).toFixed(2)} km` : `${Number(r.hours_run || 0).toFixed(2)} h`} (${String(r.run_source).replace(/_/g, " ")})`
        : "";
      const flag = isRowFlagged(r)
        ? `<span class='pill red'>${isKm ? "UNDER BENCHMARK" : "EXCESSIVE"}</span>`
        : "<span class='pill blue'>OK</span>";
      const machineKey = String(r.asset_code || "").replace(/[^A-Za-z0-9_-]/g, "_");
      list.appendChild(
        item(
          `<div class="fuel-item-head"><b>${r.asset_code}</b>${hiredTag}${archTag} ${modeTag} — ${
            isKm
              ? `${r.actual_km_per_l == null ? "-" : Number(r.actual_km_per_l).toFixed(3)} km/L`
              : `${r.actual_lph == null ? "-" : Number(r.actual_lph).toFixed(3)} L/hr`
          } ${flag}</div>` +
          `<small class="fuel-item-desc">${r.asset_name || ""}</small>` +
          `<small class="fuel-item-meta">${
            isKm
              ? `OEM: ${Number(r.oem_km_per_l || 0).toFixed(3)} km/L | Fuel: ${Number(r.fuel_liters || 0).toFixed(2)}L | Distance: ${Number(r.km_run || 0).toFixed(2)} km${runSrc}`
              : `OEM: ${Number(r.oem_lph || 0).toFixed(3)} L/hr | Fuel: ${Number(r.fuel_liters || 0).toFixed(2)}L | Hours: ${Number(r.hours_run || 0).toFixed(2)}${runSrc}`
          }</small>` +
          `<br><button data-fuel-machine="${String(r.asset_code || "").replace(/"/g, "&quot;")}">Open machine history</button> ` +
          `<button data-fuel-machine-pdf="${String(r.asset_code || "").replace(/"/g, "&quot;")}">Machine PDF</button>` +
          `<div class="fuel-inline-history" id="fuel-inline-${machineKey}"></div>`
        )
      );
    });
    if (!displayRows.length) {
      list.appendChild(item(`<small>${duplicatesOnly ? "No duplicate fuel entries found in this period." : "No fuel benchmark data in this period."}</small>`));
    }
  }

  if (duplicatesOnly) {
    setStatus(`Duplicate filter ready (${Number(data.summary?.duplicate_rows || 0)} rows in ${Number(data.summary?.duplicate_groups || 0)} groups).`);
  } else {
    const machineCount = displayRows.length;
    setStatus(machineCount ? `Fuel benchmark ready (${machineCount} machine${machineCount === 1 ? "" : "s"}).` : "Fuel benchmark ready.");
  }
}

function shiftScenarioRequestParams() {
  const start = String(qs("fuelStart")?.value || "").trim();
  const end = String(qs("fuelEnd")?.value || "").trim();
  const baseHours = Number(qs("shiftScenarioBaseHours")?.value || 11);
  const scenarioHours = Number(qs("shiftScenarioHours")?.value || 8);
  const assetCode = String(qs("shiftScenarioAsset")?.value || "").trim();
  if (!start || !end) {
    alert("Select the Fuel Benchmark start and end dates first.");
    return null;
  }
  if (!Number.isFinite(baseHours) || baseHours < 0.25 || baseHours > 24 || !Number.isFinite(scenarioHours) || scenarioHours < 0.25 || scenarioHours > 24) {
    alert("Shift hours must be between 0.25 and 24.");
    return null;
  }
  const query = new URLSearchParams({
    start,
    end,
    base_hours: String(baseHours),
    scenario_hours: String(scenarioHours),
  });
  if (assetCode) query.set("asset_code", assetCode);
  return { start, end, baseHours, scenarioHours, assetCode, query };
}

function shiftScenarioMoney(value) {
  const amount = Number(value || 0);
  return `$${amount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function renderShiftScenario(data) {
  const baseHours = Number(data?.base_hours || 11);
  const scenarioHours = Number(data?.scenario_hours || 8);
  const fleet = data?.fleet || {};
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  const excluded = Array.isArray(data?.excluded) ? data.excluded : [];
  const period = qs("shiftScenarioPeriod");
  if (period) period.textContent = `Analysis period: ${data?.start || "-"} to ${data?.end || "-"}${data?.asset_code ? ` · ${data.asset_code}` : " · All eligible equipment"}.`;

  setText("shiftScenarioBaseFuelHead", `${baseHours}h fuel`);
  setText("shiftScenarioScenarioFuelHead", `${scenarioHours}h fuel`);
  setText("shiftScenarioBaseCostHead", `${baseHours}h cost / h`);
  setText("shiftScenarioScenarioCostHead", `${scenarioHours}h cost / h`);

  const summary = qs("shiftScenarioSummary");
  if (summary) {
    const change = Number(fleet.cost_per_operating_hour_change || 0);
    const costDirection = change > 0.004 ? "higher" : change < -0.004 ? "lower" : "unchanged";
    summary.innerHTML = [
      `<span class="kpi-pill"><strong>${baseHours}h fuel:</strong> ${Number(fleet.base_fuel_liters || 0).toFixed(1)} L</span>`,
      `<span class="kpi-pill"><strong>${scenarioHours}h fuel:</strong> ${Number(fleet.scenario_fuel_liters || 0).toFixed(1)} L</span>`,
      `<span class="kpi-pill kpi-pill-green"><strong>Fuel saving:</strong> ${Number(fleet.fuel_liters_saved || 0).toFixed(1)} L</span>`,
      `<span class="kpi-pill"><strong>${baseHours}h cost / h:</strong> ${shiftScenarioMoney(fleet.base_cost_per_operating_hour)}</span>`,
      `<span class="kpi-pill"><strong>${scenarioHours}h cost / h:</strong> ${shiftScenarioMoney(fleet.scenario_cost_per_operating_hour)}</span>`,
      `<span class="kpi-pill ${change > 0.004 ? "kpi-pill-amber" : "kpi-pill-green"}"><strong>Cost / h:</strong> ${shiftScenarioMoney(Math.abs(change))} ${costDirection}</span>`,
    ].join("");
  }
  const assumptions = qs("shiftScenarioAssumptions");
  if (assumptions) {
    assumptions.textContent = `Included ${rows.length} equipment unit${rows.length === 1 ? "" : "s"} across ${Number(fleet.active_shifts || 0).toFixed(0)} active shifts. Fuel follows logged L/hr; non-fuel recorded cost is held constant per active shift.${excluded.length ? ` ${excluded.length} asset${excluded.length === 1 ? "" : "s"} excluded because the logged basis is incomplete or kilometre-based.` : ""}`;
  }

  const list = qs("shiftScenarioList");
  if (list) {
    list.innerHTML = rows.map((row) => {
      const change = Number(row.cost_per_operating_hour_change || 0);
      const changeLabel = `${change > 0 ? "+" : ""}${shiftScenarioMoney(change)}`;
      return `<tr style="border-bottom:1px solid #d9e2f3;">
        <td style="padding:9px;"><strong>${escapeHtml(row.asset_code || "-")}</strong></td>
        <td style="padding:9px;">${escapeHtml(row.asset_name || "-")}</td>
        <td style="padding:9px;">${Number(row.actual_liters_per_hour || 0).toFixed(3)}</td>
        <td style="padding:9px;">${Number(row.base_fuel_liters_per_shift || 0).toFixed(1)} L</td>
        <td style="padding:9px;">${Number(row.scenario_fuel_liters_per_shift || 0).toFixed(1)} L</td>
        <td style="padding:9px; color:#0f766e; font-weight:600;">${Number(row.fuel_liters_saved_per_shift || 0).toFixed(1)} L</td>
        <td style="padding:9px;">${shiftScenarioMoney(row.base_cost_per_operating_hour)}</td>
        <td style="padding:9px;">${shiftScenarioMoney(row.scenario_cost_per_operating_hour)}</td>
        <td style="padding:9px; color:${change > 0.004 ? "#b45309" : "#0f766e"}; font-weight:600;">${changeLabel}</td>
      </tr>`;
    }).join("") || `<tr><td colspan="9" class="muted" style="padding:12px;">No equipment had both operating-hours and fuel data for this period.</td></tr>`;
  }
}

async function loadShiftScenario() {
  const request = shiftScenarioRequestParams();
  if (!request) return;
  setStatus("Calculating shift scenario...");
  const data = await fetchJson(`${API}/api/dashboard/shift-scenario?${request.query.toString()}`);
  renderShiftScenario(data);
  setStatus(data.rows?.length ? `Shift scenario ready (${data.rows.length} equipment units).` : "Shift scenario ready. Add logged hours and fuel to include equipment.");
}

async function downloadShiftScenarioXlsx() {
  const request = shiftScenarioRequestParams();
  if (!request) return;
  setStatus("Preparing shift scenario Excel...");
  const ok = await downloadAuthedFile(
    `${API}/api/dashboard/shift-scenario.xlsx?${request.query.toString()}`,
    `IRONLOG_Shift_Scenario_${request.start}_to_${request.end}.xlsx`,
  );
  if (ok) setStatus("Shift scenario Excel downloaded.");
}

function fuelJanToDateRange() {
  const now = new Date();
  const year = now.getFullYear();
  return {
    start: `${year}-01-01`,
    end: now.toISOString().slice(0, 10),
  };
}

function fuelReconExpectedLiters(row) {
  const mode = String(row?.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
  const fuel = Number(row?.fuel_liters || 0);
  if (mode === "km") {
    const kmRun = Number(row?.km_run || 0);
    const baseKmpl = Number(row?.oem_km_per_l || 0);
    const expected = kmRun > 0 && baseKmpl > 0 ? (kmRun / baseKmpl) : 0;
    return { mode, fuel, run: kmRun, expected };
  }
  const hrsRun = Number(row?.hours_run || 0);
  const baseLph = Number(row?.oem_lph || 0);
  const expected = hrsRun > 0 && baseLph > 0 ? (hrsRun * baseLph) : 0;
  return { mode, fuel, run: hrsRun, expected };
}

async function runFuelReconciliation(useJanToDate = true) {
  let start = (qs("fuelStart")?.value || "").trim();
  let end = (qs("fuelEnd")?.value || "").trim();
  if (useJanToDate) {
    const ytd = fuelJanToDateRange();
    start = ytd.start;
    end = ytd.end;
    if (qs("fuelStart")) qs("fuelStart").value = start;
    if (qs("fuelEnd")) qs("fuelEnd").value = end;
  }
  if (!start || !end) return alert("Select start and end dates.");
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  const fuelPrice = Math.max(0, Number(qs("fuelReconPrice")?.value || 0));
  setStatus("Running fuel reconciliation...");
  setSkeleton("fuelReconList", 2);
  const data = await fetchJson(
    `${API}/api/dashboard/fuel?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}&mode=&asset_code=${encodeURIComponent(assetCode)}`
  );
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  const reconRows = rows
    .map((r) => {
      const calc = fuelReconExpectedLiters(r);
      const variance = calc.fuel - calc.expected;
      const allowed = calc.expected * Math.max(0, tolerance);
      const unexplained = Math.max(0, variance - allowed);
      const variancePct = calc.expected > 0 ? (variance / calc.expected) * 100 : null;
      const reasons = [];
      if (calc.run <= 0 && calc.fuel > 0) reasons.push("Fuel captured with no run basis");
      if (Number(r?.fill_count || 0) < 2) reasons.push("Low sample count (<2 fills)");
      if (Number(r?.is_excessive || false)) reasons.push("Outside benchmark tolerance");
      if (variancePct != null && variancePct > 50) reasons.push("Variance above +50%");
      return {
        ...r,
        recon_mode: calc.mode,
        expected_liters: Number(calc.expected.toFixed(2)),
        variance_liters: Number(variance.toFixed(2)),
        variance_pct: variancePct == null ? null : Number(variancePct.toFixed(1)),
        unexplained_liters: Number(unexplained.toFixed(2)),
        anomaly_reasons: reasons,
      };
    })
    .filter((r) => r.fuel_liters > 0)
    .sort((a, b) => Number(b.unexplained_liters || 0) - Number(a.unexplained_liters || 0));

  const totals = reconRows.reduce((acc, r) => {
    acc.actual += Number(r.fuel_liters || 0);
    acc.expected += Number(r.expected_liters || 0);
    acc.unexplained += Number(r.unexplained_liters || 0);
    if ((r.anomaly_reasons || []).length) acc.anomalies += 1;
    return acc;
  }, { actual: 0, expected: 0, unexplained: 0, anomalies: 0 });
  const missingValue = totals.unexplained * fuelPrice;
  const summary = qs("fuelReconSummary");
  if (summary) {
    summary.className = "muted";
    summary.innerHTML =
      `<b>Recon period:</b> ${escapeHtml(start)} to ${escapeHtml(end)} | ` +
      `<b>Actual:</b> ${totals.actual.toFixed(2)} L | ` +
      `<b>Expected:</b> ${totals.expected.toFixed(2)} L | ` +
      `<b>Estimated missing (unexplained):</b> ${totals.unexplained.toFixed(2)} L | ` +
      `<b>Estimated value:</b> ${fmtMoney(missingValue)} | ` +
      `<b>Anomaly assets:</b> ${totals.anomalies}`;
  }
  const list = qs("fuelReconList");
  if (list) {
    list.innerHTML = "";
    const top = reconRows.filter((r) => Number(r.unexplained_liters || 0) > 0 || (r.anomaly_reasons || []).length).slice(0, 25);
    if (!top.length) {
      list.appendChild(item("<small>No anomalies found for selected scope.</small>"));
    } else {
      top.forEach((r) => {
        const reasons = (r.anomaly_reasons || []).join("; ") || "Review";
        const modeLabel = r.recon_mode === "km" ? "km/L" : "L/hr";
        list.appendChild(
          item(
            `<div class="fuel-item-head"><b>${escapeHtml(r.asset_code || "-")}</b> — ${escapeHtml(r.asset_name || "")}</div>` +
            `<small class="fuel-item-meta">Mode: ${modeLabel} | Actual: ${Number(r.fuel_liters || 0).toFixed(2)}L | Expected: ${Number(r.expected_liters || 0).toFixed(2)}L | Variance: ${Number(r.variance_liters || 0).toFixed(2)}L | Unexplained: ${Number(r.unexplained_liters || 0).toFixed(2)}L</small>` +
            `<small class="fuel-item-meta">Reasons: ${escapeHtml(reasons)}</small>`
          )
        );
      });
    }
  }
  setStatus("Fuel reconciliation ready.");
}

function openFuelBenchmarkPdf(download = false) {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const modeFilter = String(qs("fuelModeFilter")?.value || "").trim();
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  if (!start || !end) return alert("Select start and end dates.");

  const mode = download ? "&download=1" : "";
  const url =
    `${API}/api/reports/fuel-benchmark.pdf?start=${encodeURIComponent(start)}` +
    `&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}&mode=${encodeURIComponent(modeFilter)}&asset_code=${encodeURIComponent(assetCode)}${mode}`;
  return openAuthedReport(url, { download, filename: `IRONLOG_Fuel_Benchmark_${start}_to_${end}.pdf` });
}

function openFuelBenchmarkXlsx() {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const modeFilter = String(qs("fuelModeFilter")?.value || "").trim();
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  if (!start || !end) return alert("Select start and end dates.");

  const url =
    `${API}/api/reports/fuel-benchmark.xlsx?start=${encodeURIComponent(start)}` +
    `&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}&mode=${encodeURIComponent(modeFilter)}&asset_code=${encodeURIComponent(assetCode)}`;
  return downloadAuthedFile(url, `IRONLOG_Fuel_Benchmark_${start}_to_${end}.xlsx`);
}

function openFuelReconciliationPdf(download = false) {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const modeFilter = String(qs("fuelModeFilter")?.value || "").trim();
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  const fuelPrice = Math.max(0, Number(qs("fuelReconPrice")?.value || 0));
  if (!start || !end) return alert("Select start and end dates.");
  const mode = download ? "&download=1" : "";
  const url =
    `${API}/api/reports/fuel-reconciliation.pdf?start=${encodeURIComponent(start)}` +
    `&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}` +
    `&mode=${encodeURIComponent(modeFilter)}&asset_code=${encodeURIComponent(assetCode)}` +
    `&fuel_price=${encodeURIComponent(fuelPrice)}${mode}`;
  return openAuthedReport(url, { download, filename: `IRONLOG_Fuel_Reconciliation_${start}_to_${end}.pdf` });
}

function openFuelReconciliationXlsx() {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const modeFilter = String(qs("fuelModeFilter")?.value || "").trim();
  const assetCode = String(qs("fuelAssetFilter")?.value || "").trim();
  const fuelPrice = Math.max(0, Number(qs("fuelReconPrice")?.value || 0));
  if (!start || !end) return alert("Select start and end dates.");
  const url =
    `${API}/api/reports/fuel-reconciliation.xlsx?start=${encodeURIComponent(start)}` +
    `&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}` +
    `&mode=${encodeURIComponent(modeFilter)}&asset_code=${encodeURIComponent(assetCode)}` +
    `&fuel_price=${encodeURIComponent(fuelPrice)}`;
  return downloadAuthedFile(url, `IRONLOG_Fuel_Reconciliation_${start}_to_${end}.xlsx`);
}

async function downloadExecutivePackExcel() {
  const start = (qs("fuelStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim();
  if (!start || !end) return alert("Select start and end dates.");
  setStatus("Generating executive pack...");
  try {
    const q = new URLSearchParams();
    q.set("start", start);
    q.set("end", end);
    q.set("scheduled", "10");
    q.set("near_due_hours", "50");
    const res = await fetch(`${API}/api/reports/executive-pack.xlsx?${q.toString()}`, { headers: authHeaders() });
    if (!res.ok) {
      const txt = await res.text();
      throw new Error(txt || `Executive pack request failed (${res.status})`);
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `IRONLOG_Executive_Pack_${start}_to_${end}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
    setStatus("Executive pack ready.");
  } catch (e) {
    setStatus("Executive pack error: " + (e.message || e));
    alert(`Executive pack error: ${e.message || e}`);
  }
}

function openFuelMachineHistoryPdf(assetCode, download = false) {
  const code = String(assetCode || "").trim();
  const start = (qs("fuelStart")?.value || "").trim() || (qs("fuelSnapStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim() || (qs("fuelSnapEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  if (!code) return alert("Select a machine first.");
  if (!start || !end) return alert("Select start and end dates.");
  const mode = download ? "&download=1" : "";
  const url =
    `${API}/api/reports/fuel-machine-history.pdf?asset_code=${encodeURIComponent(code)}` +
    `&start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}${mode}`;
  return openAuthedReport(url, { download, filename: `IRONLOG_Fuel_History_${code}_${start}_to_${end}.pdf` });
}

function fuelPeriodRange(anchorDate, period) {
  const end = new Date(`${anchorDate}T00:00:00`);
  if (Number.isNaN(end.getTime())) return null;
  const start = new Date(end);
  if (period === "daily") {
    // same day
  } else if (period === "weekly") {
    start.setDate(start.getDate() - 6);
  } else if (period === "monthly") {
    start.setDate(start.getDate() - 29);
  } else {
    return null;
  }
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

function maxDateStr(a, b) {
  return String(a || "") > String(b || "") ? String(a) : String(b);
}

async function loadFuelSnapshots() {
  const startInput = (qs("fuelSnapStart")?.value || "").trim();
  const endInput =
    (qs("fuelSnapEnd")?.value || "").trim() ||
    (qs("date")?.value || "").trim() ||
    new Date().toISOString().slice(0, 10);
  const startBase = startInput || endInput;
  if (!startBase || !endInput) return alert("Select snapshot start and end dates.");
  if (startBase > endInput) return alert("Snapshot start date cannot be after end date.");

  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  const dailyBase = fuelPeriodRange(endInput, "daily");
  const weeklyBase = fuelPeriodRange(endInput, "weekly");
  const monthlyBase = fuelPeriodRange(endInput, "monthly");
  const ranges = [
    { key: "daily", label: "Daily", range: { start: maxDateStr(startBase, dailyBase.start), end: endInput } },
    { key: "weekly", label: "Weekly (7d)", range: { start: maxDateStr(startBase, weeklyBase.start), end: endInput } },
    { key: "monthly", label: "Monthly (30d)", range: { start: maxDateStr(startBase, monthlyBase.start), end: endInput } },
  ];

  setStatus("Loading fuel snapshots...");
  setSkeleton("fuelSnapshotsList", 3);
  const results = await Promise.all(
    ranges.map(async (p) => {
      const data = await fetchJson(
        `${API}/api/dashboard/fuel?start=${encodeURIComponent(p.range.start)}&end=${encodeURIComponent(p.range.end)}&tolerance=${encodeURIComponent(tolerance)}`
      );
      return { ...p, data };
    })
  );

  setText("fsDailyEx", Number(results.find((r) => r.key === "daily")?.data?.summary?.excessive_count || 0));
  setText("fsWeeklyEx", Number(results.find((r) => r.key === "weekly")?.data?.summary?.excessive_count || 0));
  setText("fsMonthlyEx", Number(results.find((r) => r.key === "monthly")?.data?.summary?.excessive_count || 0));

  const list = qs("fuelSnapshotsList");
  if (!list) return;
  list.innerHTML = "";

  for (const res of results) {
    const s = res.data?.summary || {};
    const rowTop = (res.data?.rows || []).slice(0, 5)
      .map((r) => {
        const flag = r.is_excessive ? "<span class='pill red'>EXCESSIVE</span>" : "<span class='pill blue'>OK</span>";
        const metric = String(r.metric_mode || "hours") === "km"
          ? `${r.actual_km_per_l == null ? "-" : Number(r.actual_km_per_l).toFixed(3)} km/L`
          : `${r.actual_lph == null ? "-" : Number(r.actual_lph).toFixed(3)} L/hr`;
        const runTxt = String(r.metric_mode || "hours") === "km"
          ? `Distance ${Number(r.km_run || 0).toFixed(1)}km`
          : `Hours ${Number(r.hours_run || 0).toFixed(1)}`;
        return `<small><b>${r.asset_code}</b> ${metric} ${flag} | Fuel ${Number(r.fuel_liters || 0).toFixed(1)}L | ${runTxt}</small>`;
      })
      .join("<br>");
    list.appendChild(
      item(
        `<b>${res.label}</b> <span class="pill">${res.range.start} to ${res.range.end}</span>` +
        `<br><small>Fuel: ${Number(s.fuel_liters || 0).toFixed(2)}L | Hours: ${Number(s.hours_run || 0).toFixed(2)} | Distance: ${Number(s.km_run || 0).toFixed(2)}km | Avg(L/hr): ${s.avg_lph == null ? "-" : Number(s.avg_lph).toFixed(3)} | Avg(km/L): ${s.avg_km_per_l == null ? "-" : Number(s.avg_km_per_l).toFixed(3)} | Excessive: ${Number(s.excessive_count || 0)}</small>` +
        (rowTop ? `<br>${rowTop}` : "<br><small>No data in this period.</small>")
      )
    );
  }
  setStatus("Fuel snapshots ready.");
}
async function loadFuelMachineDailyInline(assetCode, mountEl) {
  if (!mountEl) return;
  // Inline machine history should follow the Fuel Benchmark date window first.
  const start = (qs("fuelStart")?.value || "").trim() || (qs("fuelSnapStart")?.value || "").trim();
  const end = (qs("fuelEnd")?.value || "").trim() || (qs("fuelSnapEnd")?.value || "").trim();
  const tolerance = Number(qs("fuelTolerance")?.value || 0.15);
  if (!assetCode || !start || !end) return;

  mountEl.innerHTML = "<small>Loading machine history...</small>";
  const data = await fetchJson(
    `${API}/api/dashboard/fuel/daily?asset_code=${encodeURIComponent(assetCode)}&start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&tolerance=${encodeURIComponent(tolerance)}`
  );
  const rows = Array.isArray(data.rows) ? data.rows : [];
  mountEl.setAttribute("data-code", String(assetCode));
  const mode = String(data.summary?.metric_mode || "hours");
  const isInlineFlagged = (r) => {
    if (r?.invalid_delta) return false;
    if (mode === "km") {
      const actual = Number(r?.actual_km_per_l);
      const benchmark = Number(r?.oem_km_per_l);
      if (!Number.isFinite(actual) || actual <= 0) return false; // no entries -> OK
      if (!Number.isFinite(benchmark) || benchmark <= 0) return false;
      return actual < benchmark;
    }
    const actual = Number(r?.actual_lph);
    const benchmark = Number(r?.oem_lph);
    if (!Number.isFinite(actual) || actual <= 0) return false; // no entries -> OK
    if (!Number.isFinite(benchmark) || benchmark <= 0) return false;
    return actual > benchmark;
  };
  const top = `<div class="fuel-inline-summary"><small><b>Fill days:</b> ${Number(data.summary?.days || 0)} | <b>Fuel:</b> ${Number(data.summary?.fuel_liters || 0).toFixed(2)}L | <b>Fill ${mode === "km" ? "distance" : "hours"}:</b> ${Number(mode === "km" ? (data.summary?.km_run || 0) : (data.summary?.hours_run || 0)).toFixed(2)} | <b>Avg:</b> ${mode === "km" ? (data.summary?.avg_km_per_l == null ? "-" : Number(data.summary.avg_km_per_l).toFixed(3) + " km/L") : (data.summary?.avg_lph == null ? "-" : Number(data.summary.avg_lph).toFixed(3) + " L/hr")} | <b>${mode === "km" ? "Under benchmark days" : "Over benchmark days"}:</b> ${Number(data.summary?.excessive_days || 0)}</small></div>`;
  const tableRows = rows.map((r) => {
    const flagged = isInlineFlagged(r);
    const invalid = Boolean(r?.invalid_delta);
    const statusClass = invalid ? "fh-status-excessive" : (flagged ? "fh-status-excessive" : "fh-status-ok");
    const statusText = invalid ? "INVALID DELTA" : (flagged ? (mode === "km" ? "UNDER BENCHMARK" : "EXCESSIVE") : "OK");
    const meterUnit = mode === "km" ? "km" : "hrs";
    const openMeterValue = r.open_meter_value == null ? "" : Number(r.open_meter_value).toFixed(2);
    const closeMeterValue = r.close_meter_value == null ? "" : Number(r.close_meter_value).toFixed(2);
    return (
      `<tr>` +
      `<td class="fh-col-date">${r.log_date}</td>` +
      `<td class="fh-col-num">${Number(r.fuel_liters || 0).toFixed(2)}</td>` +
      `<td class="fh-col-num"><input data-fuel-open-input="1" class="w-110" type="number" step="0.01" min="0" value="${openMeterValue}"> ${meterUnit}</td>` +
      `<td class="fh-col-num"><input data-fuel-close-input="1" class="w-110" type="number" step="0.01" min="0" value="${closeMeterValue}"> ${meterUnit}</td>` +
      `<td class="fh-col-num">${invalid ? "-" : Number((mode === "km" ? r.km_run : r.hours_run) || 0).toFixed(2)}</td>` +
      `<td class="fh-col-num">${invalid ? "-" : (mode === "km" ? (r.actual_km_per_l == null ? "-" : Number(r.actual_km_per_l).toFixed(3)) : (r.actual_lph == null ? "-" : Number(r.actual_lph).toFixed(3)))}</td>` +
      `<td class="fh-col-status"><span class="fh-status ${statusClass}">${statusText}</span></td>` +
      `<td class="fh-col-action"><button data-fuel-save="${Number(r.id || 0)}">Save</button> <button data-fuel-delete="${Number(r.id || 0)}">Delete</button></td>` +
      `</tr>`
    );
  }).join("");
  mountEl.innerHTML =
    `${top}<br>` +
    (tableRows
      ? `<div class="fuel-history-table-wrap"><table class="fuel-history-table"><colgroup><col style="width:14%"><col style="width:12%"><col style="width:14%"><col style="width:14%"><col style="width:14%"><col style="width:10%"><col style="width:10%"><col style="width:12%"></colgroup><thead><tr><th class="fh-col-date">Date</th><th class="fh-col-num">Fuel (L)</th><th class="fh-col-num">Open ${mode === "km" ? "km" : "hrs"}</th><th class="fh-col-num">Close ${mode === "km" ? "km" : "hrs"}</th><th class="fh-col-num">${mode === "km" ? "Distance Between Fills (km)" : "Hours Between Fills"}</th><th class="fh-col-num">${mode === "km" ? "km/L" : "L/hr"}</th><th class="fh-col-status">Status</th><th class="fh-col-action">Action</th></tr></thead><tbody>${tableRows}</tbody></table></div>`
      : "<small>No filled days found for this machine in selected range.</small>");
}

async function deleteFuelLogEntry(logId) {
  const id = Number(logId || 0);
  if (!Number.isInteger(id) || id <= 0) return;
  const ok = confirm("Delete this fuel input entry?");
  if (!ok) return;
  await fetchJson(`${API}/api/dashboard/fuel/log/${id}`, { method: "DELETE" });
}
