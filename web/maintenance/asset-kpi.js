// IRONLOG/web/maintenance/asset-kpi.js — Asset KPI, reliability and executive exports.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

let akpLastResponse = null;
let akpLastTrendSeries = [];
let akpLastMeta = null;

function akpCategoryNorm(cat) {
  const t = String(cat ?? "").trim();
  return t || "Uncategorized";
}

function akpPct(num, den) {
  return den > 0 && Number.isFinite(num) ? Number(((num / den) * 100).toFixed(1)) : null;
}

function akpRollupCategoriesFromAssets(assetRows) {
  const catMap = new Map();
  for (const a of assetRows) {
    const catKey = akpCategoryNorm(a.category);
    if (!catMap.has(catKey)) {
      catMap.set(catKey, {
        category: catKey,
        scheduled_hours: 0,
        run_hours: 0,
        downtime_hours: 0,
        available_hours: 0,
        asset_ids: new Set(),
      });
    }
    const c = catMap.get(catKey);
    c.scheduled_hours += Number(a.scheduled_hours || 0);
    c.run_hours += Number(a.run_hours || 0);
    c.downtime_hours += Number(a.downtime_hours || 0);
    c.available_hours += Number(a.available_hours || 0);
    c.asset_ids.add(a.asset_id);
  }
  const rows = Array.from(catMap.values()).map((c) => {
    const sched = c.scheduled_hours;
    const avail = c.available_hours;
    const run = c.run_hours;
    return {
      category: c.category,
      asset_count: c.asset_ids.size,
      scheduled_hours: Number(sched.toFixed(2)),
      run_hours: Number(run.toFixed(2)),
      downtime_hours: Number(c.downtime_hours.toFixed(2)),
      available_hours: Number(avail.toFixed(2)),
      availability_pct: akpPct(avail, sched),
      utilization_pct: akpPct(run, sched),
    };
  });
  rows.sort((x, y) => {
    if (x.utilization_pct == null && y.utilization_pct == null) {
      return String(x.category || "").localeCompare(String(y.category || ""));
    }
    if (x.utilization_pct == null) return 1;
    if (y.utilization_pct == null) return -1;
    return y.utilization_pct - x.utilization_pct;
  });
  return rows;
}

function akpFleetFromAssets(assetRows) {
  const fleet_sched = assetRows.reduce((s, r) => s + Number(r.scheduled_hours || 0), 0);
  const fleet_avail = assetRows.reduce((s, r) => s + Number(r.available_hours || 0), 0);
  const fleet_run = assetRows.reduce((s, r) => s + Number(r.run_hours || 0), 0);
  const fleet_down = assetRows.reduce((s, r) => s + Number(r.downtime_hours || 0), 0);
  return {
    scheduled_hours: Number(fleet_sched.toFixed(2)),
    available_hours: Number(fleet_avail.toFixed(2)),
    run_hours: Number(fleet_run.toFixed(2)),
    downtime_hours: Number(fleet_down.toFixed(2)),
    availability_pct: akpPct(fleet_avail, fleet_sched),
    utilization_pct: akpPct(fleet_run, fleet_sched),
  };
}

function akpSelectedDayRanges(start, end) {
  const ranges = [];
  const cur = new Date(`${start}T00:00:00`);
  const last = new Date(`${end}T00:00:00`);
  while (cur <= last) {
    const day = cur.toISOString().slice(0, 10);
    ranges.push({
      start: day,
      end: day,
      label: day.slice(5, 10),
    });
    cur.setDate(cur.getDate() + 1);
  }
  return ranges;
}

async function akpLoadTrendSeries(start, end, sched) {
  const ranges = akpSelectedDayRanges(start, end);
  const out = [];
  for (const r of ranges) {
    const q = new URLSearchParams();
    q.set("start", r.start);
    q.set("end", r.end);
    q.set("scheduled", String(sched));
    const res = await fetch(`${API}/dashboard/asset-kpi/weekly?${q.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load KPI trend");
    out.push({ ...r, data });
  }
  return out;
}

function akpFilteredFleet(data, filterRaw) {
  if (!filterRaw) return data?.fleet || {};
  const assets = Array.isArray(data?.by_asset) ? data.by_asset.filter((a) => akpCategoryNorm(a.category) === filterRaw) : [];
  return akpFleetFromAssets(assets);
}

function akpSeriesForAsset(assetId) {
  return (akpLastTrendSeries || []).map((row) => {
    const assetRow = (row.data?.by_asset || []).find((a) => Number(a.asset_id || 0) === Number(assetId || 0));
    const pctFromRow = (r, key) => {
      if (!r) return null;
      const v = r[key];
      if (v == null || v === "") return null;
      const n = Number(v);
      return Number.isFinite(n) ? n : null;
    };
    return {
      label: row.label,
      scheduled_hours: Number(assetRow?.scheduled_hours || 0),
      run_hours: Number(assetRow?.run_hours || 0),
      downtime_hours: Number(assetRow?.downtime_hours || 0),
      // No by_asset row (e.g. daily standby) must not read as 0% — that was misleading on trend charts.
      availability_pct: pctFromRow(assetRow, "availability_pct"),
      utilization_pct: pctFromRow(assetRow, "utilization_pct"),
    };
  });
}

function akpSvgBarChart(fleet) {
  const scheduled = Number(fleet?.scheduled_hours || 0);
  const run = Number(fleet?.run_hours || 0);
  const down = Number(fleet?.downtime_hours || 0);
  const max = Math.max(scheduled, run, down, 1);
  const bars = [
    { label: "Scheduled", value: scheduled, color: "#2563eb" },
    { label: "Run", value: run, color: "#16a34a" },
    { label: "Downtime", value: down, color: "#dc2626" },
  ];
  return `
    <svg viewBox="0 0 420 220" width="100%" height="220" role="img" aria-label="Scheduled versus run versus downtime">
      <line x1="40" y1="180" x2="390" y2="180" stroke="#cbd5e1" stroke-width="1"/>
      ${bars.map((b, i) => {
        const h = Math.max(2, (b.value / max) * 130);
        const x = 70 + i * 110;
        const y = 180 - h;
        return `
          <rect x="${x}" y="${y}" width="54" height="${h}" rx="8" fill="${b.color}"/>
          <text x="${x + 27}" y="${y - 8}" text-anchor="middle" font-size="12" fill="#334155">${fmt1(b.value)}h</text>
          <text x="${x + 27}" y="198" text-anchor="middle" font-size="12" fill="#475569">${b.label}</text>
        `;
      }).join("")}
    </svg>
  `;
}

function akpSvgTrend(series, metricKey, color, label) {
  if (!series.length) return `<div class="muted">No data available for selected days.</div>`;
  const width = 420;
  const height = 220;
  const left = 40;
  const bottom = 28;
  const top = 20;
  const plotW = width - left - 16;
  const plotH = height - top - bottom;
  const vals = series.map((s) => {
    const raw = s[metricKey];
    if (raw == null || raw === "") return null;
    const n = Number(raw);
    return Number.isFinite(n) ? n : null;
  });
  const finite = vals.filter((v) => v != null);
  const max = Math.max(...finite, 1);
  const xAt = (i) => left + (series.length === 1 ? plotW / 2 : (i * plotW) / (series.length - 1));
  const yAt = (v) => top + plotH - (v / max) * plotH;
  const points = vals.map((v, i) => ({
    x: xAt(i),
    y: v == null ? null : yAt(v),
    v,
    label: series[i].label,
  }));
  const pathParts = [];
  let segmentOpen = false;
  for (let i = 0; i < points.length; i++) {
    const p = points[i];
    if (p.v == null) {
      segmentOpen = false;
      continue;
    }
    pathParts.push(`${segmentOpen ? "L" : "M"} ${p.x} ${p.y}`);
    segmentOpen = true;
  }
  const path = pathParts.join(" ");
  return `
    <svg viewBox="0 0 ${width} ${height}" width="100%" height="220" role="img" aria-label="${label}">
      <line x1="${left}" y1="${top + plotH}" x2="${width - 8}" y2="${top + plotH}" stroke="#cbd5e1" stroke-width="1"/>
      <line x1="${left}" y1="${top}" x2="${left}" y2="${top + plotH}" stroke="#cbd5e1" stroke-width="1"/>
      ${path ? `<path d="${path}" fill="none" stroke="${color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>` : ""}
      ${points.map((p) => `
        ${p.v != null && p.y != null ? `<circle cx="${p.x}" cy="${p.y}" r="4" fill="${color}"/>` : ""}
        <text x="${p.x}" y="${top + plotH + 16}" text-anchor="middle" font-size="11" fill="#475569">${esc(p.label)}</text>
        <text x="${p.x}" y="${(p.y != null ? p.y : top + plotH) - 10}" text-anchor="middle" font-size="11" fill="#334155">${p.v != null ? fmtPct(p.v) : "—"}</text>
      `).join("")}
    </svg>
  `;
}

function akpRenderWorstAssets(assetRows) {
  const rows = [...assetRows]
    .filter((r) => r.scheduled_hours > 0)
    .sort((a, b) => {
      const au = a.utilization_pct == null ? 999 : a.utilization_pct;
      const bu = b.utilization_pct == null ? 999 : b.utilization_pct;
      if (au !== bu) return au - bu;
      const aa = a.availability_pct == null ? 999 : a.availability_pct;
      const ba = b.availability_pct == null ? 999 : b.availability_pct;
      if (aa !== ba) return aa - ba;
      return Number(b.downtime_hours || 0) - Number(a.downtime_hours || 0);
    })
    .slice(0, 5);
  if (!rows.length) return `<div class="muted">No asset rows available for this selection.</div>`;
  return `<div class="stack-10">${rows.map((r, idx) => `
    <div style="padding:10px 12px;border:1px solid #e5e7eb;border-radius:10px;background:#fff;">
      <div style="display:flex;justify-content:space-between;gap:12px;align-items:center;">
        <strong>${idx + 1}. ${esc(r.asset_code || "—")} ${r.asset_name ? `- ${esc(r.asset_name)}` : ""}</strong>
        <span class="muted">${esc(akpCategoryNorm(r.category))}</span>
      </div>
      <div class="muted" style="margin-top:4px;">
        Scheduled ${fmt1(r.scheduled_hours)}h | Run ${fmt1(r.run_hours)}h | Downtime ${fmt1(r.downtime_hours)}h
      </div>
      <div style="margin-top:4px;">
        Availability <strong>${fmtPct(r.availability_pct)}</strong> | Utilization <strong>${fmtPct(r.utilization_pct)}</strong>
      </div>
    </div>
  `).join("")}</div>`;
}

function akpMiniBars(series) {
  const totals = series.reduce((acc, s) => {
    acc.scheduled += Number(s.scheduled_hours || 0);
    acc.run += Number(s.run_hours || 0);
    acc.down += Number(s.downtime_hours || 0);
    return acc;
  }, { scheduled: 0, run: 0, down: 0 });
  return akpSvgBarChart({
    scheduled_hours: totals.scheduled,
    run_hours: totals.run,
    downtime_hours: totals.down,
  });
}

function akpDebugStrip(r) {
  const d = r?.debug || {};
  const reported = Number(d.reported_days || 0);
  const noRow = Number(d.no_daily_row_days || 0);
  const used = Number(d.used_flag_days || 0);
  const standby = Number(d.standby_flag_days || 0);
  const fallback = Number(d.fallback_schedule_days || 0);
  const logged = Number(d.logged_downtime_days || 0);
  const imputed = Number(d.imputed_open_breakdown_days || 0);
  return `
    <div class="muted" style="margin:8px 0 6px 0;">
      KPI debug: reported <b>${reported}</b>, no daily row <b>${noRow}</b>, used-flag days <b>${used}</b>, standby-flag days <b>${standby}</b>,
      fallback schedule days <b>${fallback}</b>, logged downtime days <b>${logged}</b>, open-breakdown imputed days <b>${imputed}</b>.
    </div>
  `;
}

function renderAssetKpiVisuals(data) {
  const hoursEl = document.getElementById("akpHoursChart");
  const availEl = document.getElementById("akpAvailabilityTrend");
  const utilEl = document.getElementById("akpUtilizationTrend");
  const worstEl = document.getElementById("akpWorstAssets");
  const filterRaw = String(document.getElementById("akpCategoryFilter")?.value || "").trim();
  if (!hoursEl || !availEl || !utilEl || !worstEl || !data) return;
  const allAssets = Array.isArray(data.by_asset) ? data.by_asset : [];
  const filteredAssets = filterRaw ? allAssets.filter((a) => akpCategoryNorm(a.category) === filterRaw) : allAssets;
  const fleet = filterRaw ? akpFleetFromAssets(filteredAssets) : data.fleet || {};
  hoursEl.innerHTML = akpSvgBarChart(fleet);
  worstEl.innerHTML = akpRenderWorstAssets(filteredAssets);
  const trendSeries = (akpLastTrendSeries || []).map((row) => {
    const fleetRow = akpFilteredFleet(row.data, filterRaw);
    const pctOrNull = (v) => {
      if (v == null || v === "") return null;
      const n = Number(v);
      return Number.isFinite(n) ? n : null;
    };
    return {
      label: row.label,
      availability_pct: pctOrNull(fleetRow.availability_pct),
      utilization_pct: pctOrNull(fleetRow.utilization_pct),
    };
  });
  availEl.innerHTML = akpSvgTrend(trendSeries, "availability_pct", "#2563eb", "Availability trend by selected day");
  utilEl.innerHTML = akpSvgTrend(trendSeries, "utilization_pct", "#16a34a", "Utilization trend by selected day");
}

/* ---- Availability / Utilization bar chart by type or asset ---- */

function akpKpiChartRows(data) {
  if (!data) return [];
  const filterRaw = String(document.getElementById("akpCategoryFilter")?.value || "").trim();
  const viewMode = String(document.getElementById("akpChartViewMode")?.value || "category");
  const allAssets = Array.isArray(data.by_asset) ? data.by_asset : [];
  const filteredAssets = filterRaw
    ? allAssets.filter((a) => akpCategoryNorm(a.category) === filterRaw)
    : allAssets;

  if (viewMode === "asset") {
    return filteredAssets.map((a) => ({
      label: `${String(a.asset_code || "")} ${String(a.asset_name || "")}`.trim() || "—",
      availability_pct: a.availability_pct != null ? Number(a.availability_pct) : null,
      utilization_pct: a.utilization_pct != null ? Number(a.utilization_pct) : null,
    })).sort((a, b) => a.label.localeCompare(b.label));
  }

  const cats = filterRaw
    ? akpRollupCategoriesFromAssets(filteredAssets)
    : Array.isArray(data.by_category) ? data.by_category : [];
  return cats.map((c) => ({
    label: String(c.category || "—"),
    availability_pct: c.availability_pct != null ? Number(c.availability_pct) : null,
    utilization_pct: c.utilization_pct != null ? Number(c.utilization_pct) : null,
  }));
}

function akpSvgKpiBarChart(rows, availTarget, utilTarget) {
  const n = rows.length;
  if (!n) return `<div class="muted">No data rows for the current filter.</div>`;
  const max = 100;
  // Tighter spacing: fixed 52px slot, bars fill 70% of slot split between the pair
  const slot = Math.max(44, Math.min(72, 700 / Math.max(n, 1)));
  const barPair = Math.min(26, Math.max(10, slot * 0.38));
  const gap = 4;
  const width = Math.max(600, 80 + n * slot * 2);
  const height = 380;
  const left = 52;
  const right = 20;
  const top = 36;
  const bottom = 72;
  const plotW = width - left - right;
  const plotH = height - top - bottom;
  const band = plotW / n;
  const xCenter = (i) => left + band * i + band / 2;
  const yAt = (v) => top + plotH - (Math.max(0, Math.min(100, Number(v || 0))) / max) * plotH;

  const bars = rows.map((r, i) => {
    const av = r.availability_pct;
    const ut = r.utilization_pct;
    const xA = xCenter(i) - barPair - gap / 2;
    const xU = xCenter(i) + gap / 2;
    const avBar = av != null
      ? (() => {
          const y = yAt(av); const h = Math.max(0, top + plotH - y);
          const ly = h > 16 ? y + 13 : y - 5; const lf = h > 16 ? "#fff" : "#111827";
          return `<rect x="${xA.toFixed(1)}" y="${y.toFixed(1)}" width="${barPair.toFixed(1)}" height="${h.toFixed(1)}" fill="#2563eb" rx="2"/>
          <text x="${(xA + barPair / 2).toFixed(1)}" y="${ly.toFixed(1)}" text-anchor="middle" font-size="10" font-weight="600" fill="${lf}">${av.toFixed(1)}</text>`;
        })()
      : `<text x="${(xA + barPair / 2).toFixed(1)}" y="${(top + plotH + 14).toFixed(1)}" text-anchor="middle" font-size="10" fill="#94a3b8">—</text>`;
    const utBar = ut != null
      ? (() => {
          const y = yAt(ut); const h = Math.max(0, top + plotH - y);
          const ly = h > 16 ? y + 13 : y - 5; const lf = h > 16 ? "#fff" : "#111827";
          return `<rect x="${xU.toFixed(1)}" y="${y.toFixed(1)}" width="${barPair.toFixed(1)}" height="${h.toFixed(1)}" fill="#16a34a" rx="2"/>
          <text x="${(xU + barPair / 2).toFixed(1)}" y="${ly.toFixed(1)}" text-anchor="middle" font-size="10" font-weight="600" fill="${lf}">${ut.toFixed(1)}</text>`;
        })()
      : "";
    const labelX = xCenter(i);
    const labelY = top + plotH + 16;
    const labelText = String(r.label || "").slice(0, 22);
    return `${avBar}${utBar}<text x="${labelX.toFixed(1)}" y="${labelY.toFixed(1)}" text-anchor="end" font-size="10" fill="#334155" transform="rotate(-32 ${labelX.toFixed(1)} ${labelY.toFixed(1)})">${esc(labelText)}</text>`;
  }).join("");

  const yTicks = [];
  for (let v = 0; v <= 100; v += 20) {
    const y = yAt(v);
    yTicks.push(`<line x1="${left}" y1="${y.toFixed(1)}" x2="${width - right}" y2="${y.toFixed(1)}" stroke="#e5e7eb" stroke-width="1"/>
      <text x="${left - 8}" y="${(y + 4).toFixed(1)}" text-anchor="end" font-size="11" fill="#64748b">${v}</text>`);
  }

  const targetLine = (val, color) => {
    if (val == null || !Number.isFinite(val) || val < 0 || val > 100) return "";
    const y = yAt(val);
    return `<line x1="${left}" y1="${y.toFixed(1)}" x2="${(width - right).toFixed(1)}" y2="${y.toFixed(1)}" stroke="${color}" stroke-width="2" stroke-dasharray="6 4"/>
      <text x="${(width - right + 4).toFixed(1)}" y="${(y + 4).toFixed(1)}" font-size="10" fill="${color}" font-weight="600">${val}%</text>`;
  };

  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}" width="100%" height="${height}" role="img" aria-label="Asset KPI chart" preserveAspectRatio="xMinYMid meet" style="display:block;width:100%;height:${height}px;">
    <rect x="0" y="0" width="${width}" height="${height}" fill="#ffffff" rx="8"/>
    <text x="${left}" y="16" font-size="13" font-weight="700" fill="#111827">Availability &amp; Utilization %</text>
    <text x="${left}" y="30" font-size="11" fill="#64748b">Blue = Availability · Green = Utilization · Dashed = Targets</text>
    ${yTicks.join("")}
    <line x1="${left}" y1="${top + plotH}" x2="${width - right}" y2="${top + plotH}" stroke="#94a3b8" stroke-width="1"/>
    <line x1="${left}" y1="${top}" x2="${left}" y2="${top + plotH}" stroke="#94a3b8" stroke-width="1"/>
    ${bars}
    ${targetLine(availTarget, "#1d4ed8")}
    ${targetLine(utilTarget, "#15803d")}
  </svg>`;
}

function renderAkpKpiChart(data) {
  const host = document.getElementById("akpKpiChart");
  const summaryEl = document.getElementById("akpKpiChartSummary");
  if (!host) return;
  const rows = akpKpiChartRows(data);
  if (!rows.length) {
    host.className = "muted";
    host.innerHTML = "No data for current filter. Load KPI data first.";
    if (summaryEl) summaryEl.textContent = "No data.";
    return;
  }
  const availTarget = parseFloat(document.getElementById("akpAvailTarget")?.value);
  const utilTarget = parseFloat(document.getElementById("akpUtilTarget")?.value);
  const viewMode = String(document.getElementById("akpChartViewMode")?.value || "category");
  host.className = "";
  host.innerHTML = akpSvgKpiBarChart(rows, availTarget, utilTarget);
  if (summaryEl) {
    summaryEl.textContent = `${rows.length} ${viewMode === "asset" ? "asset" : "equipment type"}${rows.length !== 1 ? "s" : ""} · Availability target: ${Number.isFinite(availTarget) ? availTarget + "%" : "—"} · Utilization target: ${Number.isFinite(utilTarget) ? utilTarget + "%" : "—"}`;
  }
}

function refreshAkpCategoryFilterOptions(data, previousValue) {
  const sel = document.getElementById("akpCategoryFilter");
  if (!sel) return;
  const cats = Array.isArray(data?.by_category) ? data.by_category : [];
  const keys = cats.map((c) => String(c.category ?? "Uncategorized"));
  const prev = previousValue != null ? String(previousValue) : String(sel.value || "");
  sel.innerHTML = `<option value="">All types</option>${keys
    .map((k) => `<option value="${escAttr(k)}">${esc(k)}</option>`)
    .join("")}`;
  if (prev && keys.includes(prev)) sel.value = prev;
  else sel.value = "";
}

function akpSelectedAssetCodes() {
  const sel = document.getElementById("akpAssetFilter");
  if (!sel) return [];
  return Array.from(sel.selectedOptions || [])
    .map((o) => String(o.value || "").trim())
    .filter(Boolean);
}

// Populate the equipment multi-select from loaded KPI data, preserving any
// existing selection. Leave empty (nothing selected) = all equipment.
function refreshAkpAssetFilterOptions(data) {
  const sel = document.getElementById("akpAssetFilter");
  if (!sel) return;
  const prevSelected = new Set(akpSelectedAssetCodes());
  const assets = Array.isArray(data?.by_asset) ? data.by_asset : [];
  const opts = assets
    .map((a) => ({ code: String(a.asset_code || "").trim(), name: String(a.asset_name || "").trim() }))
    .filter((a) => a.code)
    .sort((a, b) => a.code.localeCompare(b.code));
  sel.innerHTML = opts
    .map((a) => {
      const label = a.name ? `${a.code} — ${a.name}` : a.code;
      const selAttr = prevSelected.has(a.code) ? " selected" : "";
      return `<option value="${escAttr(a.code)}"${selAttr}>${esc(label)}</option>`;
    })
    .join("");
}


function renderAssetKpiTables(data) {
  const fleetEl = document.getElementById("akpFleetSummary");
  const catBody = document.getElementById("akpCategoryBody");
  const assetBody = document.getElementById("akpAssetBody");
  const filterSel = document.getElementById("akpCategoryFilter");
  if (!catBody || !assetBody || !data) return;

  const filterRaw = String(filterSel?.value || "").trim();
  const allAssets = Array.isArray(data.by_asset) ? data.by_asset : [];
  const filteredAssets = filterRaw
    ? allAssets.filter((a) => akpCategoryNorm(a.category) === filterRaw)
    : allAssets;

  const cats = filterRaw
    ? akpRollupCategoriesFromAssets(filteredAssets)
    : Array.isArray(data.by_category) ? data.by_category : [];

  const fleet = filterRaw ? akpFleetFromAssets(filteredAssets) : data.fleet || {};
  const days = Number(data.days_in_range || 0);
  if (fleetEl) {
    const label = filterRaw ? `<strong>Filtered (${esc(filterRaw)}):</strong>` : "<strong>All assets in range:</strong>";
    fleetEl.innerHTML = `${label} scheduled ${fmt1(fleet.scheduled_hours)} h, available ${fmt1(fleet.available_hours)} h, run ${fmt1(fleet.run_hours)} h, downtime ${fmt1(fleet.downtime_hours)} h — availability ${fmtPct(fleet.availability_pct)}, utilization ${fmtPct(fleet.utilization_pct)} <span class="muted">(${days} calendar days)</span>`;
  }

  catBody.innerHTML = cats.length
    ? cats.map((r) => `
        <tr>
          <td>${esc(r.category || "—")}</td>
          <td style="text-align:right;">${Number(r.asset_count || 0)}</td>
          <td style="text-align:right;">${fmt1(r.scheduled_hours)}</td>
          <td style="text-align:right;">${fmt1(r.available_hours)}</td>
          <td style="text-align:right;">${fmt1(r.run_hours)}</td>
          <td style="text-align:right;">${fmt1(r.downtime_hours)}</td>
          <td style="text-align:right;">${fmtPct(r.availability_pct)}</td>
          <td style="text-align:right;">${fmtPct(r.utilization_pct)}</td>
        </tr>
      `).join("")
    : `<tr><td colspan="8" class="muted">${filterRaw ? "No assets in this type for the range." : "No production daily hours in range (check Daily Input / dates)."}</td></tr>`;

  assetBody.innerHTML = filteredAssets.length
    ? filteredAssets.map((r) => {
        const trend = akpSeriesForAsset(r.asset_id);
        return `
          <tr>
            <td>${esc(r.asset_code || "—")} — ${esc(r.asset_name || "")}</td>
            <td>${esc(r.category || "—")}</td>
            <td>${esc(r.utilization_mode || "—")}</td>
            <td style="text-align:right;">${Number(r.days_with_data || 0)} / ${Number(r.days_in_range || 0)}</td>
            <td style="text-align:right;">${fmt1(r.scheduled_hours)}</td>
            <td style="text-align:right;">${fmt1(r.available_hours)}</td>
            <td style="text-align:right;">${fmt1(r.run_hours)}</td>
            <td style="text-align:right;">${fmt1(r.downtime_hours)}</td>
            <td style="text-align:right;">${fmtPct(r.availability_pct)}</td>
            <td style="text-align:right;">${fmtPct(r.utilization_pct)}</td>
          </tr>
          <tr>
            <td colspan="10" style="background:#fafafa;padding:12px;">
              ${akpDebugStrip(r)}
              <div class="form-grid">
                <div>
                  <div class="muted" style="margin-bottom:6px;">Asset hours bar view</div>
                  ${akpMiniBars(trend)}
                </div>
                <div>
                  <div class="muted" style="margin-bottom:6px;">Availability trend</div>
                  ${akpSvgTrend(trend, "availability_pct", "#2563eb", `Availability trend for ${esc(r.asset_code || "asset")}`)}
                </div>
                <div>
                  <div class="muted" style="margin-bottom:6px;">Utilization trend</div>
                  ${akpSvgTrend(trend, "utilization_pct", "#16a34a", `Utilization trend for ${esc(r.asset_code || "asset")}`)}
                </div>
              </div>
            </td>
          </tr>
        `;
      }).join("")
    : `<tr><td colspan="10" class="muted">${filterRaw ? "No rows for this type." : "No rows."}</td></tr>`;
  renderAssetKpiVisuals(data);
  renderAkpKpiChart(data);
}

let relAssetCatalog = [];
let relLastMeta = null;
let relKpiScheduledHours = 10;

function relFmtHours(v) {
  if (v == null || v === "") return "-";
  const n = Number(v);
  return Number.isFinite(n) ? n.toFixed(2) : "-";
}

function relSelectedAssetIds() {
  const sel = document.getElementById("relAssetSelect");
  if (!sel) return [];
  return Array.from(sel.selectedOptions || [])
    .map((o) => Number(o.value || 0))
    .filter((n) => n > 0);
}

function matchReliabilityAssetKpiScope() {
  const msg = document.getElementById("relMsg");
  const sel = document.getElementById("relAssetSelect");
  const start = String(document.getElementById("relStart")?.value || "").trim();
  const end = String(document.getElementById("relEnd")?.value || "").trim();
  if (!sel || !start || !end) return;

  if (!akpLastMeta || !akpLastResponse || akpLastMeta.start !== start || akpLastMeta.end !== end) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Load Asset KPI for this same date range first, then match its equipment scope.";
    }
    return;
  }

  const selectedKpiCodes = akpSelectedAssetCodes();
  const sourceAssets = selectedKpiCodes.length
    ? selectedKpiCodes
    : (Array.isArray(akpLastResponse.by_asset) ? akpLastResponse.by_asset.map((a) => a.asset_code) : []);
  const wanted = new Set(sourceAssets.map((code) => String(code || "").trim().toUpperCase()).filter(Boolean));
  if (!wanted.size) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "The Asset KPI report has no equipment in this date range to match.";
    }
    return;
  }

  const codeById = new Map(relAssetCatalog.map((asset) => [
    Number(asset.id || asset.asset_id || 0),
    String(asset.asset_code || "").trim().toUpperCase(),
  ]));
  let matched = 0;
  Array.from(sel.options).forEach((option) => {
    const match = wanted.has(codeById.get(Number(option.value || 0)) || "");
    option.selected = match;
    if (match) matched += 1;
  });
  relLastMeta = null;
  relKpiScheduledHours = Math.max(0.5, Number(akpLastMeta.sched || 10) || 10);
  if (msg) {
    msg.className = matched === wanted.size ? "message-success" : "message-error";
    msg.textContent = matched === wanted.size
      ? `Matched ${matched} Asset KPI equipment item(s). Load MTBF / LTTR to apply the shared scope.`
      : `Matched ${matched} of ${wanted.size} Asset KPI equipment item(s). Clear the Reliability category filter and try again.`;
  }
}

function relRefreshCategoryOptions(assets, keepValue = "") {
  const catSel = document.getElementById("relCategoryFilter");
  if (!catSel) return;
  const cats = Array.from(
    new Set((assets || []).map((a) => String(a.category || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));
  const prev = keepValue || String(catSel.value || "");
  catSel.innerHTML = `<option value="">All categories</option>${cats.map((c) => `<option value="${esc(c)}">${esc(c)}</option>`).join("")}`;
  if (prev && cats.includes(prev)) catSel.value = prev;
}

function relRenderAssetOptions(assets) {
  const sel = document.getElementById("relAssetSelect");
  if (!sel) return;
  const cat = String(document.getElementById("relCategoryFilter")?.value || "").trim();
  const filtered = (assets || []).filter((a) => !cat || String(a.category || "").trim() === cat);
  const prev = new Set(relSelectedAssetIds());
  sel.innerHTML = filtered.map((a) => {
    const id = Number(a.id || a.asset_id || 0);
    const code = String(a.asset_code || "NO-CODE");
    const name = String(a.asset_name || "");
    const selected = prev.has(id) ? " selected" : "";
    return `<option value="${id}"${selected}>${esc(code)} — ${esc(name)}</option>`;
  }).join("");
}

async function loadReliabilityAssets() {
  const sel = document.getElementById("relAssetSelect");
  if (!sel) return;
  try {
    const res = await fetch(`${API}/assets`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load assets");
    const assets = Array.isArray(data)
      ? data
      : Array.isArray(data?.assets)
        ? data.assets
        : Array.isArray(data?.rows)
          ? data.rows
          : [];
    relAssetCatalog = assets.filter((a) => {
      const idOk = Number.isInteger(Number(a?.id)) && Number(a.id) > 0;
      return idOk && Number(a?.active ?? 1) !== 0 && Number(a?.archived ?? 0) !== 1;
    });
    relRefreshCategoryOptions(relAssetCatalog);
    relRenderAssetOptions(relAssetCatalog);
  } catch (e) {
    sel.innerHTML = `<option value="">Failed to load assets</option>`;
    const msg = document.getElementById("relMsg");
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Asset list error: ${e.message || e}`;
    }
  }
}

function renderReliabilityReport(data) {
  const summaryEl = document.getElementById("relSummary");
  const body = document.getElementById("relAssetBody");
  const incBody = document.getElementById("relIncidentBody");
  const s = data?.summary || {};
  const recordedDowntime = Number(s.recorded_downtime_hours || 0);
  const totalDowntime = Number(s.downtime_hours || 0);
  const downtimeMeta = totalDowntime > recordedDowntime + 0.01
    ? `Asset KPI daily basis · ${relFmtHours(recordedDowntime)} h directly recorded`
    : "Same daily basis as Asset KPI";
  if (summaryEl) {
    summaryEl.innerHTML = `
      <div class="kpi-card kpi-util">
        <div class="kpi-card-header"><div class="kpi-icon">F</div><div class="kpi-title">Failures</div></div>
        <div class="kpi-big-value">${Number(s.failure_count || 0)}</div>
        <div class="kpi-meta">Incidents with downtime in period</div>
      </div>
      <div class="kpi-card kpi-scheduled">
        <div class="kpi-card-header"><div class="kpi-icon">O</div><div class="kpi-title">Operating hours</div></div>
        <div class="kpi-big-value">${relFmtHours(s.operating_hours)}</div>
        <div class="kpi-meta">Daily input run hours</div>
      </div>
        <div class="kpi-card kpi-alerts">
        <div class="kpi-card-header"><div class="kpi-icon">D</div><div class="kpi-title">Downtime hours</div></div>
        <div class="kpi-big-value">${relFmtHours(s.downtime_hours)}</div>
        <div class="kpi-meta">${downtimeMeta}</div>
      </div>
      <div class="kpi-card kpi-avail">
        <div class="kpi-card-header"><div class="kpi-icon">M</div><div class="kpi-title">MTBF</div></div>
        <div class="kpi-big-value">${relFmtHours(s.mtbf_hours)}</div>
        <div class="kpi-meta">Mean time between failures (h)</div>
      </div>
      <div class="kpi-card kpi-run">
        <div class="kpi-card-header"><div class="kpi-icon">L</div><div class="kpi-title">LTTR</div></div>
        <div class="kpi-big-value">${relFmtHours(s.lttr_hours)}</div>
        <div class="kpi-meta">Lost time to repair (h)</div>
      </div>
    `;
  }
  if (body) {
    const rows = Array.isArray(data?.by_asset) ? data.by_asset : [];
    body.innerHTML = rows.length
    ? rows.map((r) => `
      <tr>
        <td>${esc(r.asset_code || "")}</td>
        <td>${esc(r.asset_name || "")}</td>
        <td>${esc(r.category || "")}</td>
        <td style="text-align:right;">${Number(r.failure_count || 0)}</td>
        <td style="text-align:right;">${relFmtHours(r.operating_hours)}</td>
        <td style="text-align:right;">${relFmtHours(r.downtime_hours)}</td>
        <td style="text-align:right;">${relFmtHours(r.mtbf_hours)}</td>
        <td style="text-align:right;">${relFmtHours(r.lttr_hours)}</td>
      </tr>
    `).join("")
      : `<tr><td colspan="8" class="muted">No assets in scope for this filter.</td></tr>`;
  }

  if (incBody) {
    const incidents = Array.isArray(data?.incidents) ? data.incidents : [];
    const srcLabel = (src) => {
      if (src === "downtime_logs") return "Daily logs";
      if (src === "breakdown_header") return "Breakdown total";
      if (src === "work_order") return "Work order";
      return src || "-";
    };
    incBody.innerHTML = incidents.length
      ? incidents.map((r) => `
        <tr>
          <td>${esc(r.asset_code || "")}</td>
          <td>${Number(r.breakdown_id || 0) || "-"}</td>
          <td>${esc(r.breakdown_date || "")}</td>
          <td>${r.work_order_id ? `#${Number(r.work_order_id)}` : "-"}</td>
          <td style="text-align:right;">${relFmtHours(r.downtime_hours)}</td>
          <td>${esc(srcLabel(r.downtime_source))}</td>
          <td style="text-align:right;">${relFmtHours(r.log_downtime_in_period)}</td>
          <td style="text-align:right;">${relFmtHours(r.header_downtime_hours)}</td>
          <td>${esc(String(r.description || "").slice(0, 80))}</td>
        </tr>
      `).join("")
      : `<tr><td colspan="9" class="muted">No breakdown incidents with downtime in this period.</td></tr>`;
  }
}

async function loadReliabilityMetrics() {
  const msg = document.getElementById("relMsg");
  const body = document.getElementById("relAssetBody");
  const start = String(document.getElementById("relStart")?.value || "").trim();
  const end = String(document.getElementById("relEnd")?.value || "").trim();
  const category = String(document.getElementById("relCategoryFilter")?.value || "").trim();
  const assetIds = relSelectedAssetIds();
  if (!start || !end) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Choose start and end dates.";
    }
    return;
  }
  if (msg) {
    msg.className = "muted";
    msg.textContent = "Loading MTBF / LTTR…";
  }
  if (body) body.innerHTML = `<tr><td colspan="8" class="muted">Loading…</td></tr>`;
  const incBody = document.getElementById("relIncidentBody");
  if (incBody) incBody.innerHTML = `<tr><td colspan="9" class="muted">Loading…</td></tr>`;
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  q.set("scheduled", String(relKpiScheduledHours));
  if (category) q.set("category", category);
  if (assetIds.length) q.set("asset_ids", assetIds.join(","));
  try {
    const res = await fetch(`${API}/maintenance/reliability?${q.toString()}`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load reliability metrics");
    relLastMeta = { start, end, category, asset_ids: assetIds.join(","), scheduled: relKpiScheduledHours };
    renderReliabilityReport(data);
    if (msg) {
      msg.className = "message-success";
      const scope = assetIds.length
        ? `${assetIds.length} selected asset(s)`
        : category
          ? `all assets in “${category}”`
          : "all active assets";
      msg.textContent = `Loaded ${start} → ${end} for ${scope} (${Number(data.asset_filter_count || 0)} assets).`;
    }
  } catch (e) {
    relLastMeta = null;
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Error: ${e.message || e}`;
    }
    if (body) body.innerHTML = `<tr><td colspan="8" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

function exportReliabilityToExcel() {
  if (!relLastMeta) {
    alert("Load MTBF / LTTR data first.");
    return;
  }
  const q = new URLSearchParams();
  q.set("start", relLastMeta.start);
  q.set("end", relLastMeta.end);
  q.set("scheduled", String(relLastMeta.scheduled || 10));
  if (relLastMeta.category) q.set("category", relLastMeta.category);
  if (relLastMeta.asset_ids) q.set("asset_ids", relLastMeta.asset_ids);
  return downloadProtectedXlsxFile(
    `${API}/maintenance/reliability.xlsx?${q.toString()}`,
    `IRONLOG_Reliability_${relLastMeta.start}_to_${relLastMeta.end}.xlsx`,
  ).catch((err) => alert(`Could not download reliability export: ${err.message || err}`));
}

function exportReliabilityExecutiveToExcel() {
  if (!relLastMeta) {
    alert("Load MTBF / LTTR data first.");
    return;
  }
  const q = new URLSearchParams();
  q.set("start", relLastMeta.start);
  q.set("end", relLastMeta.end);
  q.set("scheduled", String(relLastMeta.scheduled || 10));
  if (relLastMeta.category) q.set("category", relLastMeta.category);
  if (relLastMeta.asset_ids) q.set("asset_ids", relLastMeta.asset_ids);
  return downloadProtectedXlsxFile(
    `${API}/maintenance/reliability-executive.xlsx?${q.toString()}`,
    `IRONLOG_MTBF_LTTR_Executive_${relLastMeta.start}_to_${relLastMeta.end}.xlsx`,
  ).catch((err) => alert(`Could not download executive reliability export: ${err.message || err}`));
}

async function loadAssetKpiWeekly() {
  const msg = document.getElementById("akpMsg");
  const catBody = document.getElementById("akpCategoryBody");
  const assetBody = document.getElementById("akpAssetBody");
  const fleetEl = document.getElementById("akpFleetSummary");
  const start = String(document.getElementById("akpStart")?.value || "").trim();
  const end = String(document.getElementById("akpEnd")?.value || "").trim();
  const schedEl = document.getElementById("akpScheduled");
  const sched = Math.max(0.5, Number(schedEl?.value || 10));
  const filterSel = document.getElementById("akpCategoryFilter");
  const prevFilter = String(filterSel?.value || "");
  if (!msg || !catBody || !assetBody) return;
  if (!start || !end) {
    msg.className = "message-error";
    msg.textContent = "Choose start and end dates.";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Loading KPI…";
  catBody.innerHTML = `<tr><td colspan="8" class="muted">Loading…</td></tr>`;
  assetBody.innerHTML = `<tr><td colspan="10" class="muted">Loading…</td></tr>`;
  if (fleetEl) fleetEl.textContent = "";
  const selectedCodes = akpSelectedAssetCodes();
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  q.set("scheduled", String(sched));
  if (selectedCodes.length) q.set("asset_codes", selectedCodes.join(","));
  try {
    const res = await fetch(`${API}/dashboard/asset-kpi/weekly?${q.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load asset KPI");
    akpLastResponse = data;
    akpLastMeta = { start, end, sched, asset_codes: selectedCodes };
    akpLastTrendSeries = await akpLoadTrendSeries(start, end, sched);
    refreshAkpCategoryFilterOptions(data, prevFilter);
    // Only repopulate the equipment picker from a full (unfiltered) load,
    // so the option list keeps every asset even after a filtered load.
    if (!selectedCodes.length) refreshAkpAssetFilterOptions(data);
    renderAssetKpiTables(data);
    msg.className = "message-success";
    const fil = String(document.getElementById("akpCategoryFilter")?.value || "").trim();
    msg.textContent = fil
      ? `Loaded ${start} → ${end}, showing type “${fil}”.`
      : `Loaded ${start} → ${end}. Higher utilization = more run hours per available hour.`;
  } catch (e) {
    akpLastResponse = null;
    akpLastTrendSeries = [];
    akpLastMeta = null;
    msg.className = "message-error";
    msg.textContent = `Error: ${e.message || e}`;
    catBody.innerHTML = `<tr><td colspan="8" class="message-error">${esc(e.message || String(e))}</td></tr>`;
    assetBody.innerHTML = `<tr><td colspan="10" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

async function downloadAssetKpiExport(url, filename, label) {
  const msg = document.getElementById("akpMsg");
  if (msg) {
    msg.className = "muted";
    msg.textContent = `Preparing ${label}...`;
  }
  try {
    const res = await fetch(url, { headers: authHeaders() });
    if (!res.ok) {
      let detail = await res.text().catch(() => "");
      try {
        const data = JSON.parse(detail);
        detail = data.error || data.message || detail;
      } catch {}
      throw new Error(detail || `${label} request failed (${res.status})`);
    }
    const blobUrl = URL.createObjectURL(await res.blob());
    const a = document.createElement("a");
    a.href = blobUrl;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(blobUrl), 5000);
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `${label} downloaded.`;
    }
  } catch (e) {
    const detail = e?.message || String(e);
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `${label} error: ${detail}`;
    }
    alert(`${label} error: ${detail}`);
  }
}

async function exportAssetKpiToExcel() {
  if (!akpLastMeta) {
    alert("Load KPI data first before exporting.");
    return;
  }
  const q = new URLSearchParams();
  q.set("start", akpLastMeta.start);
  q.set("end", akpLastMeta.end);
  q.set("scheduled", String(akpLastMeta.sched));
  // Prefer the current picker selection; fall back to whatever was last loaded.
  const selectedCodes = akpSelectedAssetCodes();
  const codes = selectedCodes.length ? selectedCodes : (akpLastMeta.asset_codes || []);
  if (codes.length) q.set("asset_codes", codes.join(","));
  await downloadAssetKpiExport(
    `${API}/dashboard/asset-kpi.xlsx?${q.toString()}`,
    `IRONLOG_Asset_KPI_${akpLastMeta.start}_to_${akpLastMeta.end}.xlsx`,
    "Asset KPI Excel export",
  );
}

async function exportAssetKpiDocx() {
  if (!akpLastMeta) {
    alert("Load KPI data first before exporting.");
    return;
  }
  const q = new URLSearchParams();
  q.set("start", akpLastMeta.start);
  q.set("end", akpLastMeta.end);
  q.set("scheduled", String(akpLastMeta.sched));
  const selectedCodes = akpSelectedAssetCodes();
  const codes = selectedCodes.length ? selectedCodes : (akpLastMeta.asset_codes || []);
  if (codes.length) q.set("asset_codes", codes.join(","));
  // Pass current chart settings
  const viewMode = String(document.getElementById("akpChartViewMode")?.value || "category");
  const categoryFilter = String(document.getElementById("akpCategoryFilter")?.value || "").trim();
  const availTarget = String(document.getElementById("akpAvailTarget")?.value || "85").trim();
  const utilTarget = String(document.getElementById("akpUtilTarget")?.value || "75").trim();
  q.set("view", viewMode);
  if (categoryFilter) q.set("category", categoryFilter);
  q.set("avail_target", availTarget);
  q.set("util_target", utilTarget);
  await downloadAssetKpiExport(
    `${API}/dashboard/asset-kpi.docx?${q.toString()}`,
    `IRONLOG_Asset_KPI_${akpLastMeta.start}_to_${akpLastMeta.end}.docx`,
    "Asset KPI Word export",
  );
}

async function exportExecutivePackFromAssetKpi() {
  const start = String(document.getElementById("akpStart")?.value || "").trim();
  const end = String(document.getElementById("akpEnd")?.value || "").trim();
  const scheduled = Math.max(0.5, Number(document.getElementById("akpScheduled")?.value || 10));
  const msg = document.getElementById("akpMsg");
  if (!start || !end) {
    alert("Choose start and end dates first.");
    return;
  }
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  q.set("scheduled", String(scheduled));
  q.set("near_due_hours", "50");
  if (msg) {
    msg.className = "muted";
    msg.textContent = "Generating executive pack...";
  }
  try {
    const res = await fetch(`${API}/reports/executive-pack.xlsx?${q.toString()}`, { headers: authHeaders() });
    if (!res.ok) {
      const txt = await res.text();
      throw new Error(txt || `Executive pack request failed (${res.status})`);
    }
    const blob = await res.blob();
    const blobUrl = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = blobUrl;
    a.download = `IRONLOG_Executive_Pack_${start}_to_${end}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(blobUrl), 5000);
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `Executive pack downloaded for ${start} to ${end}.`;
    }
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Executive pack error: ${e.message || e}`;
    }
    alert(`Executive pack error: ${e.message || e}`);
  }
}

async function exportExecutiveKpiPackBySite() {
  const periodType = String(document.getElementById("kpiPackPeriodType")?.value || "weekly").trim().toLowerCase();
  const month = String(document.getElementById("kpiPackMonth")?.value || "").trim();
  const start = String(document.getElementById("akpStart")?.value || "").trim();
  const end = String(document.getElementById("akpEnd")?.value || "").trim();
  const scheduled = Math.max(0.5, Number(document.getElementById("akpScheduled")?.value || 10));
  const siteCodesRaw = String(document.getElementById("kpiPackSiteCodes")?.value || "main").trim();
  const msg = document.getElementById("akpMsg");
  if (periodType === "weekly" && (!start || !end)) {
    alert("Weekly KPI pack needs start and end dates.");
    return;
  }
  if (periodType === "monthly" && !month) {
    alert("Monthly KPI pack needs a month.");
    return;
  }
  const siteCodes = siteCodesRaw || "main";
  const q = new URLSearchParams();
  q.set("period_type", periodType);
  q.set("site_codes", siteCodes);
  q.set("scheduled", String(scheduled));
  q.set("near_due_hours", "50");
  if (periodType === "weekly") {
    q.set("start", start);
    q.set("end", end);
  } else {
    q.set("month", month);
  }
  if (msg) {
    msg.className = "muted";
    msg.textContent = "Generating executive KPI pack by site...";
  }
  try {
    const res = await fetch(`${API}/reports/executive-kpi-pack.xlsx?${q.toString()}`, { headers: authHeaders() });
    if (!res.ok) {
      const txt = await res.text();
      throw new Error(txt || `Executive KPI pack request failed (${res.status})`);
    }
    const blob = await res.blob();
    const blobUrl = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const dateTag = new Date().toISOString().slice(0, 10);
    a.href = blobUrl;
    a.download = `IRONLOG_Executive_KPI_Pack_${periodType}_${dateTag}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(blobUrl), 5000);
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `Executive KPI pack downloaded (${periodType}, sites: ${siteCodes}).`;
    }
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Executive KPI pack error: ${e.message || e}`;
    }
    alert(`Executive KPI pack error: ${e.message || e}`);
  }
}
