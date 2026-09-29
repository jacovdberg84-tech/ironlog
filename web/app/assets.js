// IRONLOG/web/app/assets.js — Assets register, contractors, plant hire, asset history.
// Part of the main app; index.html loads these files in order and they share one global scope.

function normBool(v) {
  return v === true || v === 1 || v === "1";
}

function isHiredAsset(a) {
  const c = String(a?.category || "").toLowerCase();
  return c.includes("contractor hire") || c.includes("contractor") || c.includes("hire");
}

function isKmFuelAsset(a) {
  const mode = String(a?.utilization_mode || "").trim().toLowerCase();
  if (mode === "km") return true;
  if (mode === "hours") return false;
  const cat = String(a?.category || "").toLowerCase();
  const code = String(a?.asset_code || "").toUpperCase();
  const name = String(a?.asset_name || "").toLowerCase();
  if (code.startsWith("BMP")) return true;
  if (/^V\d{2}AM$/.test(code) || /^T\d{2}AM$/.test(code) || code.startsWith("PTT") || code.startsWith("LDV")) return true;
  const keys = ["truck", "vehicle", "ldv", "pickup", "bakkie", "tipper", "dump", "haul", "spinner"];
  if (keys.some((k) => cat.includes(k) || name.includes(k))) return true;
  if (name.includes("toyota") && name.includes("hilux")) return true;
  return false;
}

function makeContractorAssetCode() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `HIRE-${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

async function saveContractorAsset() {
  const contractor = String(qs("caContractor")?.value || "").trim();
  const codeInput = String(qs("caCode")?.value || "").trim().toUpperCase();
  const assetName = String(qs("caName")?.value || "").trim();
  const categoryInput = String(qs("caCategory")?.value || "").trim() || "Contractor Hire";
  const isStandby = !!qs("caStandby")?.checked;
  const out = qs("contractorAssetResult");

  if (!assetName) {
    alert("Enter contractor asset name / unit.");
    return;
  }

  const asset_code = codeInput || makeContractorAssetCode();
  const category = contractor ? `${categoryInput} (${contractor})` : categoryInput;
  const cost_center_code = String(qs("caCostCenter")?.value || "").trim() || undefined;
  const payload = {
    asset_code,
    asset_name: assetName,
    category,
    active: 1,
    is_standby: isStandby ? 1 : 0,
    site_code: getSessionSite() || "main",
    cost_center_code,
  };

  setStatus(`Adding contractor asset ${asset_code}...`);
  try {
    const res = await fetchJson(`${API}/api/assets`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (out) out.textContent = JSON.stringify({ ok: true, asset_code, id: res?.id || null }, null, 2);
    if (isKmFuelAsset({ category, asset_code, asset_name: assetName })) {
      await fetchJson(`${API}/api/dashboard/cost/asset-rates`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ asset_code, utilization_mode: "km" }),
      }).catch(() => {});
    }
    const hireMode = String(qs("caHireBilling")?.value || "").trim();
    if (hireMode) {
      await fetchJson(`${API}/api/assets/${encodeURIComponent(asset_code)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          hire_billing_mode: hireMode,
          hire_rate_per_hour: qs("caHireRateHour")?.value || null,
          hire_fixed_monthly: qs("caHireFixedMonthly")?.value || null,
        }),
      }).catch(() => {});
    }
    if (qs("caCode")) qs("caCode").value = "";
    if (qs("caName")) qs("caName").value = "";
    if (qs("caStandby")) qs("caStandby").checked = false;
    setStatus(`Contractor asset ${asset_code} added.`);
    await Promise.all([
      loadAssetsFleet().catch(() => {}),
      loadCodePickers().catch(() => {}),
      loadDashboard().catch(() => {}),
      loadPlantHirePanel().catch(() => {}),
    ]);
    await selectAssetCard(asset_code, { loadHistory: false, scroll: true }).catch(() => {});
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Failed to add contractor asset.");
  }
}

let assetsFleetCache = [];
let assetsSelectedCode = "";

function getSelectedAssetCode() {
  return String(assetsSelectedCode || qs("histAsset")?.value || "").trim();
}

function syncAssetsArchiveLabel() {
  const label = qs("assetsArchiveSelectedLabel");
  if (!label) return;
  const code = getSelectedAssetCode();
  label.textContent = code || "—";
}

function fleetStatusPill(status) {
  const st = String(status || "UNKNOWN").toUpperCase();
  if (st === "PRODUCTION") return "<span class='pill green'>PROD</span>";
  if (st === "STANDBY") return "<span class='pill orange'>STBY</span>";
  if (st === "DOWN") return "<span class='pill red'>DOWN</span>";
  return "<span class='pill'>—</span>";
}

function ensureAssetHistoryDateRange() {
  const endEl = qs("histEnd");
  const startEl = qs("histStart");
  const today = todayLocalYmd();
  if (endEl && !endEl.value) endEl.value = today;
  if (startEl && !startEl.value) {
    const d = new Date();
    d.setMonth(d.getMonth() - 12);
    startEl.value = d.toISOString().slice(0, 10);
  }
}

function renderAssetFleetGrid(cards) {
  const grid = qs("assetFleetGrid");
  if (!grid) return;
  const filter = String(qs("assetsFleetFilter")?.value || "").trim().toLowerCase();
  const statusFilter = String(qs("assetsFleetStatus")?.value || "").trim().toUpperCase();
  const liveCards = (cards || []).filter((c) => !Number(c.archived));
  setText("assetsCountAll", liveCards.length);
  setText("assetsCountProduction", liveCards.filter((c) => String(c.status || "").toUpperCase() === "PRODUCTION").length);
  setText("assetsCountStandby", liveCards.filter((c) => String(c.status || "").toUpperCase() === "STANDBY").length);
  setText("assetsCountDown", liveCards.filter((c) => String(c.status || "").toUpperCase() === "DOWN").length);
  const filtered = (cards || []).filter((c) => {
    if (statusFilter && String(c.status || "").toUpperCase() !== statusFilter) return false;
    if (!filter) return true;
    const hay = `${c.asset_code} ${c.asset_name} ${c.category || ""}`.toLowerCase();
    return hay.includes(filter);
  });

  grid.innerHTML = "";
  if (!filtered.length) {
    grid.innerHTML = `<div class="muted small">${cards?.length ? "No machines match your search." : "No assets found."}</div>`;
    return;
  }

  for (const c of filtered) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "asset-fleet-card";
    btn.classList.add(`status-${String(c.status || "unknown").toLowerCase()}`);
    if (Number(c.archived)) btn.classList.add("archived");
    if (c.asset_code === assetsSelectedCode) btn.classList.add("selected");
    btn.dataset.assetCode = c.asset_code;
    const hired = isHiredAsset(c);
    btn.innerHTML = `
      <div class="asset-fleet-card-head"><div class="asset-fleet-code">${escapeHtml(c.asset_code)}</div>${fleetStatusPill(c.status)}</div>
      <div class="asset-fleet-name">${escapeHtml(c.asset_name || "")}</div>
      <div class="asset-fleet-meta">
        ${c.category ? `<span class="pill">${escapeHtml(String(c.category))}</span>` : ""}
        ${hired ? "<span class='pill'>HIRED</span>" : ""}${Number(c.archived) ? "<span class='pill red'>ARCHIVED</span>" : ""}
      </div>
      <div class="asset-fleet-card-metrics"><span><small>Meter</small><strong>${Number(c.current_hours || 0).toFixed(0)} h</strong></span><span><small>Fuel 30d</small><strong>${Number(c.fuel_liters_30d || 0).toFixed(0)} L</strong></span></div>
      <div class="asset-fleet-open">Open machine profile →</div>
    `;
    grid.appendChild(btn);
  }
}

function downloadAssetsCostCentersXlsx() {
  const includeArchived = qs("showArchived")?.checked ? 1 : 0;
  void downloadAuthedFile(
    `${API}/api/assets/cost-centers.xlsx?include_archived=${includeArchived}&_ts=${Date.now()}`,
    "IRONLOG_Asset_Cost_Centers.xlsx",
  );
  setStatus("Asset cost center register export started.");
}

let plantHireRegisterCache = [];

function fillPlantHireRateFields(row) {
  if (qs("plantHireBillingMode")) {
    qs("plantHireBillingMode").value = String(row?.hire_billing_mode || "");
  }
  if (qs("plantHireRateHour")) {
    qs("plantHireRateHour").value =
      row?.hire_rate_per_hour != null && row.hire_rate_per_hour !== "" ? String(row.hire_rate_per_hour) : "";
  }
  if (qs("plantHireFixedMonthly")) {
    qs("plantHireFixedMonthly").value =
      row?.hire_fixed_monthly != null && row.hire_fixed_monthly !== "" ? String(row.hire_fixed_monthly) : "";
  }
}

function syncPlantHireAssetLabel(code) {
  const label = qs("plantHireAssetLabel");
  if (!label) return;
  const c = String(code || "").trim();
  if (!c) {
    label.textContent = "—";
    return;
  }
  const row = plantHireRegisterCache.find((r) => r.asset_code === c);
  label.textContent = row ? `${c} — ${row.asset_name || ""}` : c;
}

async function loadPlantHirePanel() {
  const monthEl = qs("plantHireBudgetMonth");
  if (monthEl && !monthEl.value) {
    monthEl.value = (qs("costMonth")?.value || "").trim() || new Date().toISOString().slice(0, 7);
  }
  try {
    const data = await fetchJson(`${API}/api/assets/hire-register`);
    plantHireRegisterCache = Array.isArray(data?.rows) ? data.rows : [];
    const sel = qs("plantHireAssetSelect");
    if (sel) {
      const prev = String(sel.value || getSelectedAssetCode() || "");
      sel.innerHTML = `<option value="">— Select hired asset —</option>`;
      plantHireRegisterCache.forEach((r) => {
        const opt = document.createElement("option");
        opt.value = r.asset_code;
        const mode = String(r.utilization_mode || "").trim().toLowerCase();
        const modeLbl = mode === "km" ? "km" : mode === "hours" ? "hrs" : "?";
        const arch = Number(r.archived) ? " [arch]" : "";
        opt.textContent = `${r.asset_code} — ${r.asset_name || ""} (${modeLbl}${arch})`;
        sel.appendChild(opt);
      });
      const pick = prev && plantHireRegisterCache.some((r) => r.asset_code === prev) ? prev : "";
      if (pick) sel.value = pick;
    }
    const code = String(qs("plantHireAssetSelect")?.value || getSelectedAssetCode() || "").trim();
    if (code) {
      const row = plantHireRegisterCache.find((r) => r.asset_code === code);
      fillPlantHireRateFields(row || {});
      syncPlantHireAssetLabel(code);
    }
    await loadPlantHireBudgetStatus().catch(() => {});
  } catch (e) {
    const out = qs("plantHireRatesResult");
    if (out) out.textContent = String(e.message || e);
  }
}

async function loadPlantHireBudgetStatus() {
  const month = (qs("plantHireBudgetMonth")?.value || "").trim();
  if (!month) return;
  const site = encodeURIComponent(getSessionSite() || "main");
  const opStatus = qs("operatingBudgetStatus");
  const phStatus = qs("plantHireBudgetStatus");
  try {
    const [opRes, phRes] = await Promise.all([
      fetchJson(`${API}/api/finance/operating-budget?period=${encodeURIComponent(month)}&site_code=${site}`),
      fetchJson(`${API}/api/finance/plant-hire-budget?period=${encodeURIComponent(month)}&site_code=${site}`),
    ]);
    const opAmt = Number(opRes?.budget_amount || 0);
    const phAmt = Number(phRes?.budget_amount || 0);
    if (qs("operatingBudgetAmount") && document.activeElement !== qs("operatingBudgetAmount")) {
      qs("operatingBudgetAmount").value = opAmt > 0 ? String(opAmt) : "";
    }
    if (qs("plantHireBudgetAmount") && document.activeElement !== qs("plantHireBudgetAmount")) {
      qs("plantHireBudgetAmount").value = phAmt > 0 ? String(phAmt) : "";
    }
    if (opStatus) {
      let opText = opAmt > 0
        ? `Saved operating budget for ${month}: ${fmtMoney(opAmt)} (one total for all expense categories)`
        : `No operating budget saved for ${month} yet — enter your total monthly expenses above.`;
      if (opAmt <= 0 && phAmt > 0) {
        opText += ` If you entered total expenses under plant hire by mistake, copy that amount here and save.`;
      }
      opStatus.textContent = opText;
    }
    if (phStatus) {
      phStatus.textContent = phAmt > 0
        ? `Saved plant hire income target for ${month}: ${fmtMoney(phAmt)}`
        : `No plant hire income target for ${month} (optional — contractor plant income only).`;
    }
  } catch (e) {
    if (opStatus) opStatus.textContent = String(e.message || e);
    if (phStatus) phStatus.textContent = String(e.message || e);
  }
}

async function saveOperatingBudget() {
  const period = (qs("plantHireBudgetMonth")?.value || "").trim();
  const budget_amount = Number(qs("operatingBudgetAmount")?.value || 0);
  if (!period) return alert("Select a budget month.");
  if (!Number.isFinite(budget_amount) || budget_amount < 0) return alert("Enter a valid budget amount.");
  setStatus("Saving operating budget...");
  try {
    const res = await fetchJson(`${API}/api/finance/operating-budget`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        period,
        site_code: getSessionSite() || "main",
        budget_amount,
      }),
    });
    setStatus(`Operating budget saved for ${period}.`);
    if (qs("operatingBudgetStatus")) {
      qs("operatingBudgetStatus").textContent =
        `Saved operating budget for ${period}: ${fmtMoney(res.budget_amount)} (one total for all expense categories)`;
    }
  } catch (e) {
    setStatus("Operating budget save failed.");
    alert(String(e.message || e));
  }
}

async function savePlantHireBudget() {
  const period = (qs("plantHireBudgetMonth")?.value || "").trim();
  const budget_amount = Number(qs("plantHireBudgetAmount")?.value || 0);
  if (!period) return alert("Select a budget month.");
  if (!Number.isFinite(budget_amount) || budget_amount < 0) return alert("Enter a valid budget amount.");
  setStatus("Saving plant hire budget...");
  try {
    const res = await fetchJson(`${API}/api/finance/plant-hire-budget`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        period,
        site_code: getSessionSite() || "main",
        budget_amount,
      }),
    });
    setStatus(`Plant hire income target saved for ${period}.`);
    if (qs("plantHireBudgetStatus")) {
      qs("plantHireBudgetStatus").textContent = `Saved plant hire income target for ${period}: ${fmtMoney(res.budget_amount)}`;
    }
  } catch (e) {
    setStatus("Plant hire budget save failed.");
    alert(String(e.message || e));
  }
}

async function savePlantHireRates() {
  const asset_code = String(qs("plantHireAssetSelect")?.value || getSelectedAssetCode() || "").trim();
  if (!asset_code) return alert("Select a hired asset first.");
  const hire_billing_mode = String(qs("plantHireBillingMode")?.value || "").trim();
  const payload = {
    hire_billing_mode: hire_billing_mode || null,
    hire_rate_per_hour: qs("plantHireRateHour")?.value || null,
    hire_fixed_monthly: qs("plantHireFixedMonthly")?.value || null,
  };
  setStatus(`Saving hire rates for ${asset_code}...`);
  try {
    const res = await fetchJson(`${API}/api/assets/${encodeURIComponent(asset_code)}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const out = qs("plantHireRatesResult");
    if (out) out.textContent = JSON.stringify(res, null, 2);
    setStatus(`Hire rates saved for ${asset_code}.`);
    await loadPlantHirePanel().catch(() => {});
  } catch (e) {
    const out = qs("plantHireRatesResult");
    if (out) out.textContent = String(e.message || e);
    setStatus("Hire rates save failed.");
  }
}

async function loadAssetsFleet() {
  const showArchived = !!qs("showArchived")?.checked;
  const url = `${API}/api/assets/fleet-summary?include_archived=${showArchived ? 1 : 0}`;
  const grid = qs("assetFleetGrid");
  if (grid && !assetsFleetCache.length) {
    grid.innerHTML = `<div class="muted small">Loading fleet…</div>`;
  }
  try {
    const data = await fetchJson(url);
    assetsFleetCache = Array.isArray(data?.cards) ? data.cards : [];
    renderAssetFleetGrid(assetsFleetCache);
    await populateHistoryAssets().catch(() => {});
    await loadPlantHirePanel().catch(() => {});
    if (assetsSelectedCode && !assetsFleetCache.some((c) => c.asset_code === assetsSelectedCode)) {
      assetsSelectedCode = "";
      qs("assetDetailPanel")?.classList.add("hidden");
    } else if (assetsSelectedCode) {
      renderAssetFleetGrid(assetsFleetCache);
    }
    syncAssetsArchiveLabel();
  } catch (e) {
    if (grid) grid.innerHTML = `<div class="muted small">Fleet load error: ${escapeHtml(e.message || e)}</div>`;
    setStatus("Fleet load error: " + (e.message || e));
  }
}

async function loadAssetDetailHeader(asset_code) {
  const header = qs("assetDetailHeader");
  const title = qs("assetDetailTitle");
  const subtitle = qs("assetDetailSubtitle");
  if (!header) return;
  header.innerHTML = `<div class="skeleton-block"></div>`;
  try {
    const data = await fetchJson(`${API}/api/assets/${encodeURIComponent(asset_code)}/qr-profile`);
    const p = data?.live_preview || {};
    const card = assetsFleetCache.find((c) => c.asset_code === asset_code);
    if (title) title.textContent = asset_code;
    if (subtitle) {
      subtitle.textContent = [
        card?.asset_name || p.asset?.asset_name,
        card?.category || p.asset?.category,
      ]
        .filter(Boolean)
        .join(" · ");
    }
    const nextSvc = p.next_service_due;
    const insp = p.inspections?.last_inspection_date;
    const hours = p.meter?.current_hours ?? card?.current_hours ?? 0;
    const fuel30 = p.fuel?.liters_last_30_days ?? card?.fuel_liters_30d ?? 0;
    header.innerHTML = `
      <div class="asset-profile-status">${fleetStatusPill(p.status || card?.status)}${Number(card?.archived) ? `<span class="pill red">Archived</span>` : ""}</div>
      <div class="asset-profile-metrics">
        <div><span>Current meter</span><strong>${Number(hours).toFixed(1)} h</strong></div>
        <div><span>Fuel used · 30 days</span><strong>${Number(fuel30).toFixed(1)} L</strong></div>
        <div class="${nextSvc && Number(nextSvc.remaining_hours ?? 0) <= 50 ? "is-warning" : ""}"><span>Next service</span><strong>${nextSvc ? `${escapeHtml(nextSvc.service_name || "Service")} · ${Number(nextSvc.remaining_hours ?? 0).toFixed(0)} h` : "Not scheduled"}</strong></div>
        <div><span>Last inspection</span><strong>${insp ? escapeHtml(String(insp)) : "No inspection"}</strong></div>
      </div>
    `;
  } catch (_) {
    const card = assetsFleetCache.find((c) => c.asset_code === asset_code);
    if (title) title.textContent = asset_code;
    if (subtitle) subtitle.textContent = card?.asset_name || "";
    header.innerHTML = card
      ? `${fleetStatusPill(card.status)} <span class="pill blue">${Number(card.current_hours || 0).toFixed(1)} h</span>`
      : "";
  }
}

function loadAssetDetailsForm(assetCode) {
  const code = String(assetCode || "").trim();
  const asset = assetsFleetCache.find((card) => String(card.asset_code || "") === code);
  if (qs("assetDetailsName")) qs("assetDetailsName").value = asset?.asset_name || "";
  if (qs("assetDetailsCategory")) qs("assetDetailsCategory").value = asset?.category || "";
  if (qs("assetDetailsStatus")) {
    qs("assetDetailsStatus").textContent = asset
      ? `Editing ${code}. Leave the class blank only when the asset does not need one.`
      : "Select a fleet card above to edit its details.";
  }
}

async function saveAssetDetails() {
  const code = getSelectedAssetCode();
  const description = String(qs("assetDetailsName")?.value || "").trim();
  const category = String(qs("assetDetailsCategory")?.value || "").trim();
  const status = qs("assetDetailsStatus");
  if (!code) return alert("Select a fleet card first.");
  if (!description) return alert("Enter an equipment description.");

  if (status) status.textContent = `Saving details for ${code}…`;
  try {
    await fetchJson(`${API}/api/assets/${encodeURIComponent(code)}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ asset_name: description, category: category || null }),
    });
    await loadAssetsFleet();
    await selectAssetCard(code, { loadHistory: false, scroll: false });
    if (status) status.textContent = `Equipment details saved for ${code}.`;
    setStatus(`Equipment details saved for ${code}.`);
  } catch (e) {
    if (status) status.textContent = String(e.message || e);
    setStatus("Equipment details save failed.");
  }
}

async function selectAssetCard(asset_code, opts = {}) {
  const code = String(asset_code || "").trim();
  if (!code) return;
  assetsSelectedCode = code;
  const sel = qs("histAsset");
  if (sel) {
    const exists = Array.from(sel.options).some((o) => o.value === code);
    if (exists) sel.value = code;
  }
  syncAssetsArchiveLabel();
  renderAssetFleetGrid(assetsFleetCache);
  qs("assetDetailPanel")?.classList.remove("hidden");
  if (qs("plantHireAssetSelect")) {
    const row = plantHireRegisterCache.find((r) => r.asset_code === code);
    if (row) {
      qs("plantHireAssetSelect").value = code;
      fillPlantHireRateFields(row);
    }
    syncPlantHireAssetLabel(code);
  }
  ensureAssetHistoryDateRange();
  loadAssetDetailsForm(code);
  await loadAssetDetailHeader(code);
  if (opts.loadHistory !== false) {
    await loadAssetHistory().catch((e) => setStatus("History error: " + (e.message || e)));
  }
  if (opts.scroll) {
    qs("assetDetailPanel")?.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

async function populateHistoryAssets() {
  const sel = qs("histAsset");
  if (!sel) return;

  const showArchived = !!qs("showArchived")?.checked;
  const url = `${API}/api/assets?include_archived=${showArchived ? 1 : 0}`;

  let assets = [];
  try {
    assets = await fetchJson(url);
  } catch (e) {
    // don’t blank the dropdown forever
    setStatus("Assets load error: " + (e.message || e));
    return;
  }

  const current = sel.value;
  sel.innerHTML = "";

  if (!assets || !assets.length) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "No assets found";
    sel.appendChild(opt);
    return;
  }

  for (const a of assets) {
    const opt = document.createElement("option");
    opt.value = a.asset_code;

    const archived = normBool(a.archived);
    const hired = isHiredAsset(a);
    const tag = archived ? " (ARCHIVED)" : "";
    const hiredTag = hired ? " [HIRED]" : "";
    opt.textContent = `${a.asset_code}${hiredTag} — ${a.asset_name}${tag}`;

    sel.appendChild(opt);
  }

  const keep = assetsSelectedCode || current;
  if (keep) {
    const exists = Array.from(sel.options).some((o) => o.value === keep);
    if (exists) sel.value = keep;
  }
  syncAssetsArchiveLabel();
}

function formatOpsSlipHistoryExtra(ev) {
  const d = ev.details || {};
  const slipId = Number(d.slip_id || 0);
  const st = String(d.slip_type || "");
  const s = d.summary || {};
  const bits = [];
  if (s.parse_error) bits.push("Summary could not be read from saved slip.");
  else {
    const pic = Number(s.pictures_attached || 0);
    if (pic > 0) bits.push(`${pic} attached picture(s) on PDF slip`);
    if (st === "hose_failure") {
      if (s.hose_part_code) bits.push(`Hose part: ${s.hose_part_code}`);
      if (s.oil_loss_part_code) bits.push(`Oil loss part: ${s.oil_loss_part_code}`);
      if (s.reason) bits.push(String(s.reason));
      if (s.preventable) bits.push("Tagged preventable");
    } else if (st === "get_change") {
      if (s.part_code) bits.push(`Part: ${s.part_code}`);
      if (s.date_changed) bits.push(`Changed: ${s.date_changed}`);
    } else if (st === "component_change") {
      if (s.component_type) bits.push(`Component: ${s.component_type}`);
      if (s.part_code) bits.push(`Part: ${s.part_code}`);
      if (s.reason) bits.push(String(s.reason));
    } else if (st === "tyre_change") {
      if (s.tyre_lines) bits.push(`${s.tyre_lines} tyre line(s)`);
      if (Array.isArray(s.positions) && s.positions.length) bits.push(`Positions: ${s.positions.join(", ")}`);
    }
  }
  const lines = bits.length
    ? bits.map((b) => `<small>${escapeHtml(b)}</small>`).join("<br>")
    : `<small>Operational slip (Breakdown Ops).</small>`;
  const who = d.created_by ? `<br><small>${escapeHtml(String(d.created_by))}</small>` : "";
  const btn = slipId
    ? `<br><button type="button" class="btn small" data-ops-slip-pdf="${slipId}">Open slip PDF</button>`
    : "";
  return `<br>${lines}${who}${btn}`;
}

function pillForType(t) {
  if (t === "breakdown") return "<span class='pill red'>BD</span>";
  if (t === "service") return "<span class='pill blue'>SV</span>";
  if (t === "fuel") return "<span class='pill green'>FUEL</span>";
  if (t === "lube") return "<span class='pill' style='background:#0d9488;color:#fff;'>LUBE</span>";
  if (t === "inspection") return "<span class='pill blue'>INSP</span>";
  if (t === "get_slip") return "<span class='pill orange'>GET</span>";
  if (t === "component_slip") return "<span class='pill orange'>COMP</span>";
  if (t === "damage_report") return "<span class='pill red'>DMG</span>";
  if (t === "tyre_change") return "<span class='pill orange'>TY CHG</span>";
  if (t === "tyre_inspection") return "<span class='pill blue'>TY INSP</span>";
  if (t === "undercarriage_inspection") return "<span class='pill' style='background:#334155;color:#fff;'>UC</span>";
  if (t === "work_order") return "<span class='pill blue'>WO</span>";
  if (t === "ops_slip") return "<span class='pill' style='background:#5b21b6;color:#fff;'>OPS</span>";
  return "<span class='pill'>EV</span>";
}

function normalizeImageSrc(v) {
  const raw = String(v || "").trim();
  if (!raw) return "";
  if (/^https?:\/\//i.test(raw)) return raw;
  if (raw.startsWith("/")) return raw;
  const normalized = raw.replace(/\\/g, "/");
  const lower = normalized.toLowerCase();
  const uploadsIdx = lower.indexOf("/uploads/");
  if (uploadsIdx >= 0) return normalized.slice(uploadsIdx);
  if (/^[a-z]:\//i.test(normalized)) return "";
  return "/" + normalized;
}

function buildPhotoDebugBadge(photoPath) {
  const hasPath = String(photoPath || "").trim().length > 0;
  if (!hasPath) return "<span class='pill orange'>No photo linked</span>";
  return "<span class='pill blue'>Photo path linked</span>";
}

async function loadAssetHistory() {
  const asset_code = getSelectedAssetCode();
  if (!asset_code) return alert("Select a fleet card first.");

  const start = qs("histStart")?.value || "";
  const end = qs("histEnd")?.value || "";

  setStatus(`Loading history for ${asset_code}...`);
  const data = await fetchJson(
    `${API}/api/assets/${encodeURIComponent(asset_code)}/history?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`
  );

  const list = qs("historyList");
  const summaryEl = qs("historySummary");
  if (!list) return;

  list.innerHTML = "";
  if (summaryEl) summaryEl.innerHTML = `<div class="skeleton-block"></div>`;
  const summary = data.summary || {};
  const counts = summary.counts || {};
  const totals = summary.totals || {};
  if (summaryEl) {
    summaryEl.innerHTML = `
      <div class="asset-history-metric"><span>Events</span><strong>${Number(counts.events_total || 0)}</strong></div>
      <div class="asset-history-metric is-danger"><span>Breakdowns</span><strong>${Number(counts.breakdowns || 0)}</strong></div>
      <div class="asset-history-metric"><span>Work orders</span><strong>${Number(counts.work_orders || 0)}</strong></div>
      <div class="asset-history-metric"><span>Inspections</span><strong>${Number(counts.inspections || 0)}</strong></div>
      <div class="asset-history-metric"><span>Fuel used</span><strong>${Number(totals.fuel_liters_total || 0).toFixed(1)} L</strong></div>
      <div class="asset-history-metric"><span>Parts cost</span><strong>$${Number(totals.parts_cost_total || 0).toFixed(2)}</strong></div>
      <div class="asset-history-metric"><span>Maintenance cost</span><strong>$${Number(summary.maintenance_cost_total || 0).toFixed(2)}</strong></div>
    `;
  }

  (data.history || []).forEach((ev) => {
    const wo = ev.work_order_id ? ` <small>WO #${ev.work_order_id}</small>` : "";

    const rawPhoto = ev.details?.photo || "";
    const image = normalizeImageSrc(rawPhoto);
    const debugBadge = buildPhotoDebugBadge(rawPhoto);
    const unresolvedWindowsPath = rawPhoto && !image;
    const debugPrefix = unresolvedWindowsPath
      ? `${debugBadge} <span class='pill red'>Windows path unresolved</span>`
      : debugBadge;
    const photoBlock = image
      ? `<div style="margin-top:8px;">${debugPrefix}<br><img src="${image}" alt="event photo" style="max-width:180px; max-height:120px; border-radius:8px; border:1px solid var(--line); object-fit:cover; margin-top:6px;" onload="this.dataset.loaded='1'; this.previousSibling && this.previousSibling.remove && this.previousSibling.remove();" onerror="this.insertAdjacentHTML('beforebegin','<span class=&quot;pill red&quot;>Photo file missing / blocked</span><br>'); this.style.display='none';" /></div>`
      : `<div style="margin-top:8px;">${debugPrefix}</div>`;

    const extra =
      ev.type === "breakdown"
        ? (() => {
            const logs = ev.details?.downtime_logs || [];
            const lines = logs.length
              ? `<div style="margin-top:6px; padding-left:10px; border-left:2px solid #7d2a2a;">
                   ${logs
                     .map(
                       (l) =>
                         `<div><small><b>${l.log_date}</b> — ${l.hours_down}h${
                           l.notes ? ` | ${l.notes}` : ""
                         }</small></div>`
                     )
                     .join("")}
                 </div>`
              : `<br><small>No downtime log lines recorded.</small>`;

            return `<br><small>Downtime total: ${ev.details?.downtime_hours ?? 0}h ${
              ev.details?.critical ? " | CRIT" : ""
            }</small>${lines}${photoBlock}`;
          })()
        : ev.type === "get_slip"
        ? `<br><small>Items: ${(ev.details?.items || []).length}</small>`
        : ev.type === "component_slip"
        ? `<br><small>${ev.details?.serial_out || ""} → ${ev.details?.serial_in || ""}</small>`
        : ev.type === "work_order"
        ? `<br><small>Status: ${ev.details?.status || ""}</small>${photoBlock}`
        : ev.type === "damage_report"
        ? `<br><small>${ev.details?.notes || "No notes recorded."}</small>${photoBlock}`
        : ev.type === "tyre_change"
        ? `<br><small>${ev.details?.serial_out || "-"} → ${ev.details?.serial_in || "-"}</small>${
            ev.details?.hours_at_change != null
              ? `<br><small>Hours at change: ${Number(ev.details.hours_at_change).toFixed(1)}</small>`
              : ""
          }${ev.details?.notes ? `<br><small>${ev.details.notes}</small>` : ""}${photoBlock}`
        : ev.type === "tyre_inspection"
        ? `<br><small>Condition: ${ev.details?.condition || "-"}</small>${
            ev.details?.pressure != null ? `<br><small>Pressure: ${ev.details.pressure}</small>` : ""
          }${
            ev.details?.tread_depth != null ? `<br><small>Tread: ${ev.details.tread_depth}</small>` : ""
          }${ev.details?.notes ? `<br><small>${ev.details.notes}</small>` : ""}${photoBlock}`
        : ev.type === "undercarriage_inspection"
        ? `<br><small>SMU: ${ev.details?.smu ?? "-"} | Max wear: ${ev.details?.worst_wear_pct ?? "-"}%${
            ev.details?.worst_component ? ` (${ev.details.worst_component})` : ""
          }</small>${
            ev.details?.pdf_url
              ? `<br><small><button type="button" class="link-button" data-open-authed-pdf="${escapeHtml(ev.details.pdf_url)}">Open PDF</button></small>`
              : ""
          }${ev.details?.notes ? `<br><small>${ev.details.notes}</small>` : ""}`
        : ev.type === "ops_slip"
        ? formatOpsSlipHistoryExtra(ev)
        : ev.type === "fuel"
        ? `<br><small>${Number(ev.details?.liters || 0).toFixed(1)} L${ev.details?.source ? ` · ${escapeHtml(String(ev.details.source))}` : ""}</small>`
        : ev.type === "lube"
        ? `<br><small>${Number(ev.details?.quantity || 0).toFixed(1)} L${ev.details?.oil_type ? ` · ${escapeHtml(String(ev.details.oil_type))}` : ""}</small>`
        : ev.type === "inspection"
        ? `<br><small>${ev.details?.inspector ? `Inspector: ${escapeHtml(String(ev.details.inspector))}` : "Inspection recorded"}${ev.details?.notes ? `<br>${escapeHtml(String(ev.details.notes))}` : ""}</small>`
        : "";

    list.appendChild(item(`${pillForType(ev.type)} <b>${ev.date}</b> — ${ev.title}${wo}${extra}`));
  });

  if (!data.history?.length) list.appendChild(item("<small>No history found for this range.</small>"));

  setStatus("History loaded ✅");
}

async function archiveSelectedAsset() {
  const code = getSelectedAssetCode();
  if (!code) return alert("Select a fleet card first.");

  const reason = prompt(`Archive ${code}.\nReason (optional):`, "Scrapped / Not in use");
  if (reason === null) return;

  setStatus(`Archiving ${code}...`);
  await fetchJson(`${API}/api/assets/${encodeURIComponent(code)}/archive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ archived: true, reason: String(reason || "").trim() }),
  });

  setStatus(`Archived ${code} ✅`);
  await loadAssetsFleet().catch(() => {});
  await loadDashboard().catch(() => {});
}

async function unarchiveSelectedAsset() {
  const code = getSelectedAssetCode();
  if (!code) return alert("Select a fleet card first.");

  if (!confirm(`Unarchive ${code}?`)) return;

  setStatus(`Unarchiving ${code}...`);
  await fetchJson(`${API}/api/assets/${encodeURIComponent(code)}/archive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ archived: false }),
  });

  setStatus(`Unarchived ${code} ✅`);
  await loadAssetsFleet().catch(() => {});
  await loadDashboard().catch(() => {});
}

/* =========================
   INIT
========================= */
