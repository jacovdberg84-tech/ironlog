// IRONLOG/web/app/inventory.js — Part issues, store allocations, manual stock, bins, cycle counts, lube stock.
// Part of the main app; index.html loads these files in order and they share one global scope.

async function issuePart() {
  const woId = (qs("iWo")?.value || "").trim();
  const payload = {
    part_code: (qs("iPart")?.value || "").trim(),
    quantity: Number(qs("iQty")?.value || 1),
  };

  setStatus("Issuing part...");
  try {
    const res = await fetchJson(`${API}/api/workorders/${woId}/issue`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("issueResult", JSON.stringify(res, null, 2));
    setStatus("Part issued.");
    await loadDashboard().catch(() => {});
  } catch (e) {
    setText("issueResult", String(e.message || e));
    setStatus("Issue failed.");
  }
}

async function allocateStore() {
  const payload = {
    part_code: (qs("saPart")?.value || "").trim(),
    quantity: Number(qs("saQty")?.value || 0),
    location_code: (qs("saLocation")?.value || "").trim() || undefined,
    bin_code: (qs("saBin")?.value || "").trim() || undefined,
    asset_code: (qs("saAsset")?.value || "").trim() || undefined,
    work_order_id: (qs("saWo")?.value || "").trim() ? Number((qs("saWo")?.value || "").trim()) : undefined,
    allocation_date: (qs("saDate")?.value || "").trim() || undefined,
    issued_by: (qs("saIssuedBy")?.value || "").trim() || undefined,
    cost_center_code: (qs("saCostCenter")?.value || "").trim() || undefined,
    notes: (qs("saNotes")?.value || "").trim() || undefined,
  };

  setStatus("Allocating stores...");
  try {
    const res = await fetchJson(`${API}/api/stock/allocate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("storeAllocResult", JSON.stringify(res, null, 2));
    setStatus("Stores allocated.");
    await Promise.all([
      loadStoreAllocations().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
  } catch (e) {
    setText("storeAllocResult", String(e.message || e));
    setStatus("Stores allocation failed.");
  }
}

async function loadStoreAllocations() {
  const list = qs("storeAllocList");
  if (!list) return;

  list.innerHTML = "";
  setSkeleton("storeAllocList", 2);

  const rows = await fetchJson(`${API}/api/stock/allocations`);
  const data = Array.isArray(rows?.rows) ? rows.rows : [];

  list.innerHTML = "";
  data.slice(0, 20).forEach((r) => {
    const ref = r.work_order_id ? `WO #${r.work_order_id}` : r.asset_code;
    const unitCost = Number(r.unit_cost || 0);
    const lineValue = Number((unitCost * Number(r.quantity || 0)).toFixed(2));
    list.appendChild(
      item(
        `<b>${r.allocation_date}</b> — ${r.part_code} x ${Number(r.quantity || 0).toFixed(1)}<br>` +
        `<small>${ref} | ${r.location_code || "NO-LOC"}${r.bin_code ? `/${r.bin_code}` : ""} | CC: ${r.cost_center_code || "-"} | Unit: $${unitCost.toFixed(2)} | Value: $${lineValue.toFixed(2)} | ${r.issued_by || "No issuer"}${r.notes ? ` | ${r.notes}` : ""}</small>`
      )
    );
  });

  if (!data.length) {
    list.appendChild(item("<small>No store allocations yet.</small>"));
  }
}

function applyAssetCostCenterToInputs(assetCode) {
  const code = String(assetCode || "").trim();
  const a = code && window.__assetsByCode ? window.__assetsByCode[code] : null;
  const cc = a?.cost_center_code ? String(a.cost_center_code) : "";
  if (qs("fuelCostCenter")) qs("fuelCostCenter").value = cc;
  if (qs("mlCostCenter")) qs("mlCostCenter").value = cc;
}

async function loadCostCenterOptions() {
  const list = qs("costCenterOptions");
  if (!list) return;
  try {
    const data = await fetchJson(`${API}/api/masterdata/cost-centers`);
    const rows = Array.isArray(data?.rows) ? data.rows : [];
    list.innerHTML = "";
    rows.forEach((r) => {
      const code = String(r.code || "").trim();
      if (!code) return;
      const opt = document.createElement("option");
      opt.value = code;
      opt.label = String(r.name || code);
      list.appendChild(opt);
    });
  } catch {}
}

async function populateAssetAllocSelect() {
  const sel = qs("assetAllocSelect");
  if (!sel) return;
  const current = String(sel.value || "");
  try {
    const assets = await fetchJson(`${API}/api/assets?include_archived=0`);
    const rows = Array.isArray(assets) ? assets : [];
    window.__assetsByCode = {};
    sel.innerHTML = '<option value="">Select asset…</option>';
    rows.forEach((a) => {
      const code = String(a.asset_code || "").trim();
      if (!code) return;
      window.__assetsByCode[code] = a;
      const opt = document.createElement("option");
      opt.value = code;
      const cc = a.cost_center_code ? ` · ${a.cost_center_code}` : "";
      opt.textContent = `${code} — ${a.asset_name || ""}${cc}`;
      sel.appendChild(opt);
    });
    if (current && window.__assetsByCode[current]) sel.value = current;
    loadAssetAllocationForm();
  } catch (e) {
    setStatus("Asset allocation list failed: " + (e.message || e));
  }
}

function loadAssetAllocationForm() {
  const code = String(qs("assetAllocSelect")?.value || "").trim();
  const a = code && window.__assetsByCode ? window.__assetsByCode[code] : null;
  if (qs("assetAllocSite")) qs("assetAllocSite").value = a?.site_code ? String(a.site_code) : (getSessionSite() || "main");
  if (qs("assetAllocCostCenter")) qs("assetAllocCostCenter").value = a?.cost_center_code ? String(a.cost_center_code) : "";
  if (qs("assetAllocDepartment")) qs("assetAllocDepartment").value = a?.department_code ? String(a.department_code) : "";
}

async function saveAssetAllocation() {
  const code = String(qs("assetAllocSelect")?.value || "").trim();
  const out = qs("assetAllocResult");
  if (!code) return alert("Select an asset first.");
  const body = {
    site_code: String(qs("assetAllocSite")?.value || "").trim() || null,
    cost_center_code: String(qs("assetAllocCostCenter")?.value || "").trim() || null,
    department_code: String(qs("assetAllocDepartment")?.value || "").trim() || null,
  };
  setStatus("Saving asset allocation…");
  try {
    const res = await fetchJson(`${API}/api/assets/${encodeURIComponent(code)}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (out) out.textContent = JSON.stringify(res?.asset || res, null, 2);
    await Promise.all([
      populateAssetAllocSelect(),
      loadCodePickers(),
      loadCostCenterOptions(),
    ]);
    setStatus(`Allocation saved for ${code}.`);
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Asset allocation save failed.");
  }
}

async function loadCodePickers() {
  const assetList = qs("assetCodeOptions");
  const partList = qs("partCodeOptions");
  const locationList = qs("locationCodeOptions");
  if (!assetList && !partList && !locationList) return;

  if (assetList) {
    try {
      const assets = await fetchJson(`${API}/api/assets?include_archived=0`);
      window.__assetsByCode = window.__assetsByCode || {};
      assetList.innerHTML = "";
      (Array.isArray(assets) ? assets : []).forEach((a) => {
        const code = String(a.asset_code || "").trim();
        if (!code) return;
        window.__assetsByCode[code] = a;
        const hiredTag = isHiredAsset(a) ? " [HIRED]" : "";
        const opt = document.createElement("option");
        opt.value = code;
        opt.textContent = `${code}${hiredTag} - ${a.asset_name || ""}`;
        assetList.appendChild(opt);
      });
    } catch {}
  }
  await loadCostCenterOptions();

  if (partList) {
    try {
      const parts = await fetchJson(`${API}/api/stock/onhand`);
      partList.innerHTML = "";
      const map = {};
      (Array.isArray(parts) ? parts : []).forEach((p) => {
        const code = String(p.part_code || "").trim();
        if (!code) return;
        map[String(code).toUpperCase()] = String(p.part_name || "").trim();
        const opt = document.createElement("option");
        opt.value = code;
        opt.textContent = `${code} - ${p.part_name || ""}`;
        partList.appendChild(opt);
      });
      window.__partNameByCode = map;
    } catch {}
  }

  if (locationList) {
    try {
      const locations = await fetchJson(`${API}/api/stock/locations?active=1`);
      const rows = Array.isArray(locations?.rows) ? locations.rows : [];
      locationList.innerHTML = "";
      rows.forEach((l) => {
        const code = String(l.location_code || "").trim();
        if (!code) return;
        const locText = Array.isArray(r.allowed_locations) && r.allowed_locations.length ? r.allowed_locations.join(", ") : "all";
        tr.innerHTML = `<td>${escapeHtml(r.username)}</td><td>${escapeHtml(r.full_name || "")}</td><td>${escapeHtml(r.department || "")}</td><td>${escapeHtml(rolesText)}</td><td>${escapeHtml(locText)}</td><td>${r.active ? "yes" : "no"}</td><td>${r.has_password ? "yes" : "no"}</td>`;
        opt.value = code;
        opt.textContent = `${code}${l.location_name ? ` - ${l.location_name}` : ""}`;
        locationList.appendChild(opt);
      });
    } catch {}
  }

  applyDefaultLocationsToInputs();
}

function updateManualStockCostRowVisibility() {
  const t = String(qs("msType")?.value || "in").trim().toLowerCase();
  const row = qs("msCostRow");
  if (row) row.style.display = t === "in" ? "" : "none";
}

function updateManualStockPartDesc() {
  const code = String(qs("msPart")?.value || "").trim().toUpperCase();
  const descEl = qs("msPartDesc");
  if (!descEl) return;
  const map = window.__partNameByCode || {};
  const name = code && map && map[code] ? String(map[code]) : "";
  if (name) {
    descEl.value = name;
    descEl.disabled = true;
    return;
  }
  descEl.disabled = false;
  if (!code) {
    descEl.value = "";
    return;
  }
  descEl.value = descEl.value || "";
  fetchPartNameByCode(code).then((n) => {
    const now = String(qs("msPart")?.value || "").trim().toUpperCase();
    if (now !== code) return;
    if (n) {
      descEl.value = n;
      descEl.disabled = true;
    } else {
      descEl.disabled = false;
    }
  });
}

// Fallback lookup (covers cases where code pickers haven't loaded yet)
const __partNameFetchCache = new Map();
let __partNameFetchSeq = 0;
async function fetchPartNameByCode(code) {
  const c = String(code || "").trim().toUpperCase();
  if (!c) return "";
  if (__partNameFetchCache.has(c)) return __partNameFetchCache.get(c) || "";
  const mySeq = ++__partNameFetchSeq;
  try {
    const data = await fetchJson(`${API}/api/stock/control-summary?part_code=${encodeURIComponent(c)}`);
    const name = String(data?.part?.part_name || "").trim();
    if (mySeq === __partNameFetchSeq) {
      __partNameFetchCache.set(c, name || "");
      if (!window.__partNameByCode) window.__partNameByCode = {};
      window.__partNameByCode[c] = name || "";
    }
    return name || "";
  } catch {
    __partNameFetchCache.set(c, "");
    return "";
  }
}

function updateManualLubePartDesc() {
  const code = String(qs("mlPart")?.value || "").trim().toUpperCase();
  const descEl = qs("mlPartDesc");
  if (!descEl) return;
  const map = window.__partNameByCode || {};
  const name = code && map && map[code] ? String(map[code]) : "";
  descEl.value = name || "";
  if (!descEl.value && code) {
    fetchPartNameByCode(code).then((n) => {
      const now = String(qs("mlPart")?.value || "").trim().toUpperCase();
      if (now !== code) return;
      descEl.value = n || "";
    });
  }
}

function updateLubeMinPartDesc() {
  const code = String(qs("lubeMinPart")?.value || "").trim().toUpperCase();
  const descEl = qs("lubeMinPartDesc");
  if (!descEl) return;
  const map = window.__partNameByCode || {};
  const name = code && map && map[code] ? String(map[code]) : "";
  descEl.value = name || "";
  if (!descEl.value && code) {
    fetchPartNameByCode(code).then((n) => {
      const now = String(qs("lubeMinPart")?.value || "").trim().toUpperCase();
      if (now !== code) return;
      descEl.value = n || "";
    });
  }
}

function updateReceiveLubePartDesc() {
  const code = String(qs("lrPart")?.value || "").trim().toUpperCase();
  const descEl = qs("lrPartDesc");
  if (!descEl) return;
  const map = window.__partNameByCode || {};
  const name = code && map && map[code] ? String(map[code]) : "";
  descEl.value = name || (descEl.value || "");
  // If code is unknown, allow manual description entry (needed to create new stock items)
  descEl.disabled = !!name;
}

async function setThisLubeMinimum() {
  const part_code = (qs("mlPart")?.value || "").trim();
  const min_stock = Number(qs("mlMinInput")?.value || 0);
  if (!part_code) return alert("Enter a lube stock number first.");
  if (!Number.isFinite(min_stock) || min_stock < 0) return alert("Minimum must be >= 0.");
  setStatus("Saving lube minimum...");
  try {
    const res = await fetchJson(`${API}/api/stock/part-minimum`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ part_code, min_stock }),
    });
    setText("manualLubeResult", JSON.stringify(res, null, 2));
    await Promise.all([
      loadLubeStockOnHand().catch(() => {}),
      loadStockOnHandPage().catch(() => {}),
      loadInventoryControl().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
    setStatus("Lube minimum updated.");
  } catch (e) {
    setText("manualLubeResult", String(e.message || e));
    setStatus("Failed to set lube minimum.");
  }
}

async function receiveLubeStock() {
  const part_code = (qs("lrPart")?.value || "").trim();
  const location_code = (qs("lrLocation")?.value || "").trim() || "LUBE";
  const quantity = Number(qs("lrQty")?.value || 0);
  const reference = (qs("lrRef")?.value || "").trim() || "lube_receive";
  const part_name = (qs("lrPartDesc")?.value || "").trim();
  if (!part_code) return alert("Enter lube stock number.");
  if (!Number.isFinite(quantity) || quantity <= 0) return alert("Quantity must be > 0.");

  setStatus("Receiving lube stock...");
  try {
    const res = await fetchJson(`${API}/api/stock/movement`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        part_code,
        movement_type: "in",
        quantity,
        reference,
        location_code,
        part_name: part_name || undefined,
        create_if_missing: true,
      }),
    });
    setText("receiveLubeResult", JSON.stringify(res, null, 2));
    await Promise.all([
      loadLubeStockOnHand().catch(() => {}),
      loadStockOnHandPage().catch(() => {}),
      loadInventoryControl().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
    setStatus("Lube stock received.");
  } catch (e) {
    setText("receiveLubeResult", String(e.message || e));
    setStatus("Receive lube failed.");
  }
}

async function saveManualStock() {
  const movement_type = String(qs("msType")?.value || "in").trim().toLowerCase();
  const part_code = String(qs("msPart")?.value || "").trim().toUpperCase();
  const part_name = String(qs("msPartDesc")?.value || "").trim();
  const rawCost = String(qs("msUnitCost")?.value || "").trim();
  const unit_cost =
    movement_type === "in" && rawCost !== "" && Number.isFinite(Number(rawCost)) && Number(rawCost) > 0
      ? Number(rawCost)
      : undefined;
  const cost_currency = String(qs("msCostCurrency")?.value || "USD").trim().toUpperCase();

  const payload = {
    part_code,
    location_code: (qs("msLocation")?.value || "").trim() || undefined,
    bin_code: (qs("msBin")?.value || "").trim() || undefined,
    movement_type,
    quantity: Number(qs("msQty")?.value || 0),
    reference: (qs("msRef")?.value || "").trim() || undefined,
    cost_center_code: (qs("msCostCenter")?.value || "").trim() || undefined,
    ...(movement_type === "in"
      ? { create_if_missing: true, ...(part_name ? { part_name } : {}) }
      : {}),
    ...(movement_type === "in" && unit_cost != null
      ? { unit_cost, cost_currency: ["USD", "ZAR", "MZN"].includes(cost_currency) ? cost_currency : "USD" }
      : {}),
  };

  setStatus("Saving manual stock entry...");
  try {
    const res = await fetchJson(`${API}/api/stock/movement`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("manualStockResult", JSON.stringify(res, null, 2));
    setStatus("Manual stock saved.");
    if (part_code && part_name) {
      __partNameFetchCache.set(part_code, part_name);
      if (!window.__partNameByCode) window.__partNameByCode = {};
      window.__partNameByCode[part_code] = part_name;
    }
    // Clear form for fast consecutive entries.
    if (qs("msPart")) qs("msPart").value = "";
    if (qs("msPartDesc")) qs("msPartDesc").value = "";
    if (qs("msQty")) qs("msQty").value = "1";
    if (qs("msRef")) qs("msRef").value = "";
    if (qs("msBin")) qs("msBin").value = "";
    if (qs("msCostCenter")) qs("msCostCenter").value = "";
    if (qs("msUnitCost")) qs("msUnitCost").value = "";
    if (qs("msType")) qs("msType").value = "out";
    if (qs("msCostCurrency")) qs("msCostCurrency").value = "USD";
    updateManualStockCostRowVisibility();
    qs("msPart")?.focus();
    await loadDashboard().catch(() => {});
  } catch (e) {
    setText("manualStockResult", String(e.message || e));
    setStatus("Manual stock save failed.");
  }
}

async function loadInventoryControl() {
  const part_code = (qs("icPartCode")?.value || "").trim();
  const q = part_code ? `?part_code=${encodeURIComponent(part_code)}` : "";
  setStatus("Loading inventory control...");
  setSkeleton("icLubeLowList", 1);
  try {
    const data = await fetchJson(`${API}/api/stock/control-summary${q}`);
    const summary = data.summary || {};
    const part = data.part || null;
    const lubeRows = Array.isArray(data.low_lube_rows) ? data.low_lube_rows : [];

    setText("icBelowMinTotal", Number(summary.below_min_total || 0));
    setText("icLubeLowCount", Number(summary.lube_below_min_count || 0));
    setText("icOnHand", part ? Number(part.on_hand || 0).toFixed(1) : "-");
    setText("icMinStock", part ? Number(part.min_stock || 0).toFixed(1) : "-");
    if (part && qs("icMinInput")) qs("icMinInput").value = Number(part.min_stock || 0).toFixed(1);
    if (part && qs("icCountedQty")) qs("icCountedQty").value = Number(part.on_hand || 0).toFixed(1);

    const list = qs("icLubeLowList");
    if (list) {
      list.innerHTML = "";
      lubeRows.forEach((r) => {
        list.appendChild(
          item(
            `<b>${r.part_code}</b> — ${Number(r.on_hand || 0).toFixed(1)} on hand <span class='pill red'>LOW</span>` +
            `<br><small>${r.part_name || ""} | Min ${Number(r.min_stock || 0).toFixed(1)} | Short ${Number(r.shortage || 0).toFixed(1)}</small>`
          )
        );
      });
      if (!lubeRows.length) list.appendChild(item("<small>No low lube items right now.</small>"));
    }

    if (part) {
      setText("inventoryControlResult", JSON.stringify(part, null, 2));
      setStatus(`Inventory control ready for ${part.part_code}.`);
    } else {
      setText("inventoryControlResult", JSON.stringify(summary, null, 2));
      setStatus("Inventory control summary ready.");
    }
  } catch (e) {
    setText("inventoryControlResult", String(e.message || e));
    setStatus("Inventory control load failed.");
  }
}

async function saveInventoryPartMinimum() {
  const part_code = (qs("icPartCode")?.value || "").trim();
  const min_stock = Number(qs("icMinInput")?.value || 0);
  if (!part_code) return alert("Enter part code first.");
  if (!Number.isFinite(min_stock) || min_stock < 0) return alert("Minimum must be >= 0.");
  setStatus("Saving part minimum...");
  try {
    const res = await fetchJson(`${API}/api/stock/part-minimum`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ part_code, min_stock }),
    });
    setText("inventoryControlResult", JSON.stringify(res, null, 2));
    await Promise.all([
      loadInventoryControl().catch(() => {}),
      loadStockOnHandPage().catch(() => {}),
      loadLubeStockOnHand().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
    setStatus("Part minimum updated.");
  } catch (e) {
    setText("inventoryControlResult", String(e.message || e));
    setStatus("Failed to update part minimum.");
  }
}

async function submitInventoryCycleCount() {
  const part_code = (qs("icPartCode")?.value || "").trim();
  const counted_qty = Number(qs("icCountedQty")?.value || 0);
  const reason = (qs("icCountReason")?.value || "").trim() || "cycle_count";
  if (!part_code) return alert("Enter part code first.");
  if (!Number.isFinite(counted_qty) || counted_qty < 0) return alert("Counted qty must be >= 0.");

  setStatus("Submitting cycle count...");
  try {
    const res = await fetchJson(`${API}/api/stock/cycle-count`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ part_code, counted_qty, reason }),
    });
    setText("inventoryControlResult", JSON.stringify(res, null, 2));
    await Promise.all([
      loadInventoryControl().catch(() => {}),
      loadApprovalRequests().catch(() => {}),
    ]);
    setStatus(res.no_change ? "Cycle count matched on-hand (no request)." : "Cycle count request submitted.");
  } catch (e) {
    setText("inventoryControlResult", String(e.message || e));
    setStatus("Cycle count submit failed.");
  }
}

async function saveManualLube() {
  const part_code = (qs("mlPart")?.value || "").trim();
  const qtyRequested = Number(qs("mlQty")?.value || 0);
  if (part_code && Number.isFinite(lubeStockMatch.on_hand) && qtyRequested > Number(lubeStockMatch.on_hand)) {
    const warn = `Requested ${qtyRequested.toFixed(1)} exceeds available ${Number(lubeStockMatch.on_hand).toFixed(1)} for ${lubeStockMatch.part_code || part_code}.`;
    setText("mlQtyWarn", warn);
    setStatus("Cannot save lube: insufficient stock.");
    return;
  }
  const mlCc = String(qs("mlCostCenter")?.value || "").trim();
  const payload = {
    asset_code: (qs("mlAsset")?.value || "").trim(),
    log_date: (qs("mlDate")?.value || "").trim() || undefined,
    part_code: part_code || undefined,
    location_code: (qs("mlLocation")?.value || "").trim() || undefined,
    oil_type: (qs("mlType")?.value || "").trim() || undefined,
    quantity: Number(qs("mlQty")?.value || 0),
    cost_center_code: mlCc || undefined,
  };

  setStatus(part_code ? "Issuing lube stock..." : "Saving manual lube entry...");
  try {
    const endpoint = part_code ? `${API}/api/stock/lube-issue` : `${API}/api/stock/lube-log`;
    const res = await fetchJson(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("manualLubeResult", JSON.stringify(res, null, 2));
    setStatus("Manual lube saved.");
    await Promise.all([
      loadDashboard().catch(() => {}),
      loadLubeUsage().catch(() => {}),
    ]);
  } catch (e) {
    setText("manualLubeResult", String(e.message || e));
    setStatus("Manual lube save failed.");
  }
}

async function loadLocations() {
  const showInactive = Boolean(qs("locShowInactive")?.checked);
  setStatus("Loading locations...");
  setSkeleton("locList", 1);
  try {
    const data = await fetchJson(`${API}/api/stock/locations?active=${showInactive ? "0" : "1"}`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const list = qs("locList");
    if (list) {
      list.innerHTML = "";
      rows.forEach((l) => {
        const active = Number(l.active || 0) === 1;
        const tone = active ? "" : "opacity:0.65;";
        list.appendChild(
          item(
            `<div style="display:flex;gap:10px;align-items:center;${tone}">` +
              `<div style="min-width:76px;"><b>${l.location_code}</b></div>` +
              `<div style="flex:1;">${l.location_name || "<span class='muted'>(no name)</span>"}</div>` +
              `<span class="pill ${active ? "blue" : "orange"}">${active ? "ACTIVE" : "INACTIVE"}</span>` +
            `</div>`
          )
        );
      });
      if (!rows.length) list.appendChild(item("<small>No locations found.</small>"));
    }
    setText("locResult", JSON.stringify({ count: rows.length }, null, 2));
    await loadCodePickers().catch(() => {});
    setStatus("Locations ready.");
  } catch (e) {
    setText("locResult", String(e.message || e));
    setStatus("Locations load failed.");
  }
}

async function saveLocation() {
  const location_code = String(qs("locCode")?.value || "").trim().toUpperCase();
  const location_name = String(qs("locName")?.value || "").trim() || undefined;
  const active = String(qs("locActive")?.value || "1") === "1";
  if (!location_code) return alert("Enter location code.");

  setStatus("Saving location...");
  try {
    const res = await fetchJson(`${API}/api/stock/locations`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ location_code, location_name, active }),
    });
    setText("locResult", JSON.stringify(res, null, 2));
    await Promise.all([loadLocations().catch(() => {}), loadCodePickers().catch(() => {})]);
    setStatus("Location saved.");
  } catch (e) {
    setText("locResult", String(e.message || e));
    setStatus("Location save failed.");
  }
}

async function saveStockBin() {
  const location_code = String(qs("sbLocationCode")?.value || "").trim().toUpperCase();
  const bin_code = String(qs("sbBinCode")?.value || "").trim().toUpperCase();
  const bin_name = String(qs("sbBinName")?.value || "").trim() || undefined;
  if (!location_code || !bin_code) return alert("Location and bin code are required.");
  setStatus("Saving stock bin...");
  const res = await fetchJson(`${API}/api/stock/bins`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ location_code, bin_code, bin_name, active: true }),
  });
  setText("sbResult", JSON.stringify(res, null, 2));
  await loadStockBins();
  setStatus("Stock bin saved.");
}

async function loadStockBins() {
  const list = qs("sbBinsList");
  if (!list) return;
  const location_code = String(qs("sbLocationCode")?.value || "").trim().toUpperCase();
  const q = location_code ? `?location_code=${encodeURIComponent(location_code)}&active=1` : "?active=1";
  const data = await fetchJson(`${API}/api/stock/bins${q}`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  list.innerHTML = "";
  rows.forEach((r) => {
    list.appendChild(item(`<b>${r.location_code || "-"}/${r.bin_code || "-"}</b><br><small>${r.bin_name || ""}</small>`));
  });
  if (!rows.length) list.appendChild(item("<small>No bins found.</small>"));
}

async function saveStockMinMax() {
  const part_code = String(qs("sbPartCode")?.value || "").trim();
  const location_code = String(qs("sbLocationCode")?.value || "").trim().toUpperCase();
  const bin_code = String(qs("sbBinCode")?.value || "").trim().toUpperCase() || undefined;
  const min_qty = Number(qs("sbMinQty")?.value || 0);
  const max_qty = Number(qs("sbMaxQty")?.value || 0);
  const reorder_qty_raw = String(qs("sbReorderQty")?.value || "").trim();
  const reorder_qty = reorder_qty_raw === "" ? undefined : Number(reorder_qty_raw);
  if (!part_code || !location_code) return alert("Part code and location are required.");
  if (!Number.isFinite(min_qty) || min_qty < 0 || !Number.isFinite(max_qty) || max_qty < 0) return alert("Min and max must be >= 0.");
  const res = await fetchJson(`${API}/api/stock/min-max`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ part_code, location_code, bin_code, min_qty, max_qty, reorder_qty }),
  });
  setText("sbResult", JSON.stringify(res, null, 2));
  await loadReplenishmentSuggestions();
  setStatus("Min-max policy saved.");
}

async function loadStockDepth() {
  const list = qs("sbDepthList");
  if (!list) return;
  const part_code = String(qs("sbPartCode")?.value || "").trim();
  const location_code = String(qs("sbLocationCode")?.value || "").trim().toUpperCase();
  const bin_code = String(qs("sbBinCode")?.value || "").trim().toUpperCase();
  const q = new URLSearchParams();
  if (part_code) q.set("part_code", part_code);
  if (location_code) q.set("location_code", location_code);
  if (bin_code) q.set("bin_code", bin_code);
  const data = await fetchJson(`${API}/api/stock/depth${q.toString() ? `?${q.toString()}` : ""}`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  list.innerHTML = "";
  rows.slice(0, 120).forEach((r) => {
    list.appendChild(item(
      `<b>${r.part_code || "-"}</b> ${r.location_code || "-"}/${r.bin_code || "-"}`
      + `<br><small>On hand: ${Number(r.on_hand || 0).toFixed(2)} | On order: ${Number(r.on_order || 0).toFixed(2)} | Available: ${Number(r.available || 0).toFixed(2)}</small>`
    ));
  });
  if (!rows.length) list.appendChild(item("<small>No depth rows found.</small>"));
}

let stockReplenishmentRowsCache = [];

async function loadReplenishmentSuggestions() {
  const list = qs("sbReplenishmentList");
  if (!list) return;
  const location_code = String(qs("sbLocationCode")?.value || "").trim().toUpperCase();
  const bin_code = String(qs("sbBinCode")?.value || "").trim().toUpperCase();
  const q = new URLSearchParams();
  if (location_code) q.set("location_code", location_code);
  if (bin_code) q.set("bin_code", bin_code);
  const data = await fetchJson(`${API}/api/stock/replenishment-suggestions${q.toString() ? `?${q.toString()}` : ""}`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  stockReplenishmentRowsCache = rows;
  list.innerHTML = "";
  rows.slice(0, 120).forEach((r) => {
    list.appendChild(item(
      `<b>${r.part_code || "-"}</b> ${r.location_code || "-"}/${r.bin_code || "-"} <span class="pill red">REPLENISH</span>`
      + `<br><small>On hand: ${Number(r.on_hand || 0).toFixed(2)} | Min: ${Number(r.min_qty || 0).toFixed(2)} | Suggest: ${Number(r.suggested_order_qty || 0).toFixed(2)}</small>`
    ));
  });
  if (!rows.length) list.appendChild(item("<small>No replenishment suggestions.</small>"));
}

function exportReplenishmentSuggestionsCsv() {
  const rows = Array.isArray(stockReplenishmentRowsCache) ? stockReplenishmentRowsCache : [];
  if (!rows.length) throw new Error("Load replenishment suggestions first.");
  const header = [
    "part_code",
    "part_name",
    "location_code",
    "bin_code",
    "on_hand",
    "min_qty",
    "max_qty",
    "shortage_qty",
    "suggested_order_qty",
  ];
  const esc = (v) => {
    const s = String(v ?? "");
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g, "\"\"")}"`;
    return s;
  };
  const csv = [header.join(",")].concat(rows.map((r) => [
    r.part_code || "",
    r.part_name || "",
    r.location_code || "",
    r.bin_code || "",
    Number(r.on_hand || 0).toFixed(2),
    Number(r.min_qty || 0).toFixed(2),
    Number(r.max_qty || 0).toFixed(2),
    Number(r.shortage_qty || 0).toFixed(2),
    Number(r.suggested_order_qty || 0).toFixed(2),
  ].map(esc).join(","))).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "replenishment_suggestions.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function createCycleSession() {
  const location_code = String(qs("sbLocationCode")?.value || "").trim().toUpperCase();
  const bin_code = String(qs("sbBinCode")?.value || "").trim().toUpperCase() || undefined;
  if (!location_code) return alert("Location is required to create a cycle session.");
  const res = await fetchJson(`${API}/api/stock/cycle-sessions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      location_code,
      bin_code,
      planned_date: new Date().toISOString().slice(0, 10),
      notes: "Created from stock tab",
    }),
  });
  setText("sbResult", JSON.stringify(res, null, 2));
  await loadCycleSessions();
  setStatus("Cycle session created.");
}

async function loadCycleSessions() {
  const list = qs("sbCycleSessionsList");
  if (!list) return;
  const data = await fetchJson(`${API}/api/stock/cycle-sessions`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  list.innerHTML = "";
  rows.slice(0, 80).forEach((r) => {
    const id = Number(r.id || 0);
    list.appendChild(item(
      `<b>Session #${id}</b> ${r.location_code || "-"}/${r.bin_code || "-"} <span class="pill blue">${r.status || "-"}</span>`
      + `<br><small>Lines: ${Number(r.line_count || 0)} | Variance abs: ${Number(r.variance_abs || 0).toFixed(2)} | Planned: ${r.planned_date || "-"}</small>`
      + `<br><button data-sb-cs-submit="${id}" style="margin-top:8px;">Submit</button>`
      + ` <button data-sb-cs-approve="${id}" style="margin-top:8px;">Approve</button>`
      + ` <button data-sb-cs-countone="${id}" style="margin-top:8px;">Add One Part Count</button>`
    ));
  });
  if (!rows.length) list.appendChild(item("<small>No cycle sessions.</small>"));
}

async function addOnePartCountToSession(sessionId) {
  const id = Number(sessionId || 0);
  if (!id) return;
  const part_code = String(qs("sbPartCode")?.value || "").trim();
  if (!part_code) return alert("Enter/select part code first.");
  const countedRaw = prompt("Counted quantity for this part:", "0");
  if (countedRaw == null) return;
  const counted_qty = Number(countedRaw);
  if (!Number.isFinite(counted_qty) || counted_qty < 0) return alert("Counted qty must be >= 0.");
  const res = await fetchJson(`${API}/api/stock/cycle-sessions/${id}/lines/upsert`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ lines: [{ part_code, counted_qty, reason: "session_count" }] }),
  });
  setText("sbResult", JSON.stringify(res, null, 2));
  await loadCycleSessions();
}

async function submitCycleSession(sessionId) {
  const id = Number(sessionId || 0);
  if (!id) return;
  const res = await fetchJson(`${API}/api/stock/cycle-sessions/${id}/submit`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("sbResult", JSON.stringify(res, null, 2));
  await loadCycleSessions();
  setStatus(`Cycle session #${id} submitted.`);
}

async function approveCycleSession(sessionId) {
  const id = Number(sessionId || 0);
  if (!id) return;
  const ok = confirm(`Approve cycle session #${id} and post adjustment movements?`);
  if (!ok) return;
  const res = await fetchJson(`${API}/api/stock/cycle-sessions/${id}/approve`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("sbResult", JSON.stringify(res, null, 2));
  await Promise.all([loadCycleSessions().catch(() => {}), loadStockDepth().catch(() => {}), loadStockOnHandPage().catch(() => {})]);
  setStatus(`Cycle session #${id} approved.`);
}

let lubeStockMatch = { part_code: null, on_hand: null };

function updateLubeQtyWarning() {
  const warnEl = qs("mlQtyWarn");
  const part_code = (qs("mlPart")?.value || "").trim();
  const qty = Number(qs("mlQty")?.value || 0);
  if (!warnEl) return;
  if (!part_code || !Number.isFinite(lubeStockMatch.on_hand) || !Number.isFinite(qty) || qty <= 0) {
    warnEl.textContent = "";
    return;
  }
  const available = Number(lubeStockMatch.on_hand || 0);
  if (qty > available) {
    warnEl.textContent = `Warning: requested ${qty.toFixed(1)} is above available ${available.toFixed(1)} for ${lubeStockMatch.part_code || part_code}.`;
  } else {
    warnEl.textContent = "";
  }
}

async function loadLubeStockOnHand() {
  const qText = (qs("mlPart")?.value || "").trim() || (qs("mlType")?.value || "").trim();
  const list = qs("lubeStockList");
  if (list) setSkeleton("lubeStockList", 1);
  try {
    // Issue lube uses the dedicated LUBE store by default
    const location_code = "LUBE";
    const q = qText ? `?q=${encodeURIComponent(qText)}&location_code=${encodeURIComponent(location_code)}` : `?location_code=${encodeURIComponent(location_code)}`;
    const data = await fetchJson(`${API}/api/stock/lube-onhand${q}`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const exact = data.exact || (rows.length ? rows[0] : null);

    lubeStockMatch = {
      part_code: exact?.part_code || null,
      on_hand: exact != null ? Number(exact.on_hand || 0) : null,
    };

    const quick = qs("mlLubeQuickLine");
    if (quick) {
      const oils = rows
        .filter((r) => Number(r.on_hand || 0) > 0)
        .slice(0, 8)
        .map((r) => `${r.part_code}: ${Number(r.on_hand || 0).toFixed(0)}`)
        .join(" | ");
      quick.textContent = oils ? `LUBE store available: ${oils}` : "LUBE store available: -";
    }

    setText("mlAvailableQty", exact ? Number(exact.on_hand || 0).toFixed(1) : "-");
    setText("mlAvailablePart", exact ? `${exact.part_code || "-"} ${exact.part_name ? `(${exact.part_name})` : ""}` : "-");
    const partEl = qs("mlPart");
    const typeText = (qs("mlType")?.value || "").trim();
    if (partEl && !String(partEl.value || "").trim() && typeText && exact?.part_code) {
      partEl.value = String(exact.part_code);
    }
    updateManualLubePartDesc();

    if (list) {
      list.innerHTML = "";
      rows.slice(0, 8).forEach((r) => {
        list.appendChild(
          item(
            `<b>${r.part_code}</b> — ${Number(r.on_hand || 0).toFixed(1)} on hand` +
            `${r.below_min ? " <span class='pill red'>LOW</span>" : ""}` +
            `<br><small>${r.part_name || ""} | Min: ${Number(r.min_stock || 0).toFixed(1)}</small>`
          )
        );
      });
      if (!rows.length) list.appendChild(item("<small>No lube stock items found for this filter.</small>"));
    }
    updateLubeQtyWarning();
  } catch (e) {
    lubeStockMatch = { part_code: null, on_hand: null };
    const quick = qs("mlLubeQuickLine");
    if (quick) quick.textContent = "";
    setText("mlAvailableQty", "-");
    setText("mlAvailablePart", "-");
    updateLubeQtyWarning();
    const msg = String(e.message || e);
    if (list) list.innerHTML = `<div class="item"><small>Lube stock load failed: ${escapeHtml(msg)}</small></div>`;
    setStatus("Lube stock load failed: " + msg);
  }
}

async function setLubeMinimumStock() {
  const minStock = Number(qs("lubeMinStockValue")?.value || 210);
  if (!Number.isFinite(minStock) || minStock < 0) {
    alert("Minimum lube stock must be a valid number >= 0.");
    return;
  }
  setStatus(`Setting minimum lube stock to ${minStock}...`);
  try {
    const res = await fetchJson(`${API}/api/stock/lube-minimums`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ min_stock: minStock }),
    });
    setText("manualLubeResult", JSON.stringify(res, null, 2));
    await Promise.all([
      loadLubeStockOnHand().catch(() => {}),
      loadLubeReorderAlerts().catch(() => {}),
      loadStockOnHandPage().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
    setStatus(`Minimum lube stock set to ${Number(res.min_stock || minStock)} for ${Number(res.updated_count || 0)} item(s).`);
  } catch (e) {
    setText("manualLubeResult", String(e.message || e));
    setStatus("Failed to set minimum lube stock.");
  }
}

async function setSingleLubeMinimum() {
  const part_code = (qs("lubeMinPart")?.value || "").trim();
  const min_stock = Number(qs("lubeMinValue")?.value || 0);
  if (!part_code) return alert("Enter a lube stock number first.");
  if (!Number.isFinite(min_stock) || min_stock < 0) return alert("Minimum must be >= 0.");
  setStatus("Saving lube minimum...");
  try {
    const res = await fetchJson(`${API}/api/stock/part-minimum`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ part_code, min_stock }),
    });
    setText("lubeMinResult", JSON.stringify(res, null, 2));
    await Promise.all([
      loadLubeStockOnHand().catch(() => {}),
      loadLubeReorderAlerts().catch(() => {}),
      loadStockOnHandPage().catch(() => {}),
      loadInventoryControl().catch(() => {}),
      loadDashboard().catch(() => {}),
    ]);
    setStatus("Lube minimum updated.");
  } catch (e) {
    setText("lubeMinResult", String(e.message || e));
    setStatus("Failed to set lube minimum.");
  }
}

async function loadLubeReorderAlerts() {
  const list = qs("lubeReorderList");
  if (list) setSkeleton("lubeReorderList", 1);
  try {
    const data = await fetchJson(`${API}/api/stock/lube-onhand`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const flagged = rows
      .map((r) => ({
        ...r,
        on_hand: Number(r.on_hand || 0),
        min_stock: Number(r.min_stock || 0),
      }))
      .filter((r) => {
        const min = r.min_stock;
        if (!Number.isFinite(min) || min <= 0) return false;
        const near = min + Math.max(1, min * 0.1);
        return r.on_hand <= near;
      })
      .sort((a, b) => (a.on_hand - a.min_stock) - (b.on_hand - b.min_stock))
      .slice(0, 30);

    if (list) {
      list.innerHTML = "";
      flagged.forEach((r) => {
        const low = r.on_hand <= r.min_stock;
        const pill = low ? "<span class='pill red'>REORDER</span>" : "<span class='pill orange'>NEAR MIN</span>";
        list.appendChild(
          item(
            `<b>${escapeHtml(r.part_code)}</b> ${pill} — On hand ${r.on_hand.toFixed(1)} | Min ${r.min_stock.toFixed(1)}` +
              `<br><small>${escapeHtml(r.part_name || "")}</small>`
          )
        );
      });
      if (!flagged.length) list.appendChild(item("<small>No lube items near/below minimum.</small>"));
    }
  } catch (e) {
    if (list) list.innerHTML = `<div class="item"><small>${escapeHtml(String(e.message || e))}</small></div>`;
  }
}

async function loadLubeAnalytics() {
  const months = Number(qs("lubeMonths")?.value || 6);
  setStatus("Loading lube analytics...");
  setSkeleton("lubeAnalyticsList", 2);
  const data = await fetchJson(`${API}/api/dashboard/lube/analytics?months=${encodeURIComponent(months)}`);
  const summary = data.summary || {};
  setText("laTypes", Number(summary.oils || 0));
  setText("laQty", Number(summary.qty_total || 0).toFixed(1));
  setText("laLowRisk", Number(summary.low_risk_count || 0));

  const list = qs("lubeAnalyticsList");
  if (!list) return;
  const trend = Array.isArray(data.trend) ? data.trend : [];
  const monthSet = Array.from(new Set(trend.map((t) => String(t.month || "")))).filter(Boolean);
  const forecast = Array.isArray(data.forecast) ? data.forecast : [];

  list.innerHTML = "";
  forecast.slice(0, 20).forEach((r) => {
    const perMonth = monthSet
      .map((m) => {
        const hit = trend.find((t) => t.month === m && String(t.oil_key || "") === String(r.oil_key || ""));
        return `${m}: ${Number(hit?.qty || 0).toFixed(1)}`;
      })
      .join(" | ");
    list.appendChild(
      item(
        `<b>${r.oil_key || "-"}</b> ${r.low_risk ? "<span class='pill red'>LOW RISK</span>" : "<span class='pill blue'>OK</span>"}` +
          `<br><small>Total ${Number(r.qty_total || 0).toFixed(1)} | Avg/day ${Number(r.avg_daily_use || 0).toFixed(2)} | On hand ${r.on_hand == null ? "-" : Number(r.on_hand).toFixed(1)} | Min ${r.min_stock == null ? "-" : Number(r.min_stock).toFixed(1)} | Days to min ${r.days_to_min == null ? "-" : Number(r.days_to_min).toFixed(1)}</small>` +
          `<br><small>${perMonth || "No monthly trend data."}</small>` +
          `<br><button data-map-oil-key="${String(r.oil_key || "").replace(/"/g, "&quot;")}" data-map-part-code="${String((r.part_code || r.mapped_part_code || "")).replace(/"/g, "&quot;")}" style="margin-top:8px;">Map this</button>`
      )
    );
  });
  if (!forecast.length) list.appendChild(item("<small>No lube analytics found for this period.</small>"));
  setStatus("Lube analytics ready.");
}

async function loadLubeMappings() {
  const list = qs("lubeMapList");
  if (!list) return;
  setSkeleton("lubeMapList", 1);
  const data = await fetchJson(`${API}/api/dashboard/lube/mappings`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  list.innerHTML = "";
  rows.forEach((r) => {
    list.appendChild(
      item(
        `<b>${r.oil_key || "-"}</b> -> ${r.part_code || "-"}` +
        `<br><small>${r.updated_by || "-"} @ ${r.updated_at || "-"}</small>`
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No lube mappings yet.</small>"));
}

async function saveLubeMapping() {
  const oil_key = (qs("lubeMapOilKey")?.value || "").trim();
  const part_code = (qs("lubeMapPartCode")?.value || "").trim();
  if (!oil_key || !part_code) {
    alert("Enter oil key and stock code.");
    return;
  }
  setStatus("Saving lube mapping...");
  const res = await fetchJson(`${API}/api/dashboard/lube/mappings`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ oil_key, part_code }),
  });
  setText("manualLubeResult", JSON.stringify(res, null, 2));
  await Promise.all([
    loadLubeMappings().catch(() => {}),
    loadLubeAnalytics().catch(() => {}),
  ]);
  setStatus("Lube mapping saved.");
}

/** Start-up: Part issue, store allocation, manual stock, bins and cycle count controls. Called once from init() in init.js. */
function wireInventoryControls() {
  qs("issuePart")?.addEventListener("click", () =>
    issuePart().catch((e) => setStatus("Issue error: " + e.message))
  );
  qs("allocateStore")?.addEventListener("click", () =>
    allocateStore().catch((e) => setStatus("Stores allocation error: " + e.message))
  );
  qs("refreshAllocations")?.addEventListener("click", () =>
    loadStoreAllocations().catch((e) => setStatus("Allocation list error: " + e.message))
  );
  qs("saveManualStock")?.addEventListener("click", () =>
    saveManualStock().catch((e) => setStatus("Manual stock error: " + e.message))
  );
  qs("msPart")?.addEventListener("input", updateManualStockPartDesc);
  qs("msPart")?.addEventListener("change", updateManualStockPartDesc);
  qs("msType")?.addEventListener("change", () => {
    updateManualStockCostRowVisibility();
  });
  qs("msType")?.addEventListener("input", () => {
    updateManualStockCostRowVisibility();
  });
  qs("mlPart")?.addEventListener("input", updateManualLubePartDesc);
  qs("mlPart")?.addEventListener("change", updateManualLubePartDesc);
  // Lube minimums moved to separate card
  qs("lubeMinPart")?.addEventListener("input", updateLubeMinPartDesc);
  qs("lubeMinPart")?.addEventListener("change", updateLubeMinPartDesc);
  qs("lubeMinSetOne")?.addEventListener("click", () =>
    setSingleLubeMinimum().catch((e) => setStatus("Lube min error: " + e.message))
  );
  qs("lubeMinRefresh")?.addEventListener("click", () =>
    loadLubeReorderAlerts().catch((e) => setStatus("Lube alerts error: " + e.message))
  );
  qs("receiveLube")?.addEventListener("click", () =>
    receiveLubeStock().catch((e) => setStatus("Receive lube error: " + e.message))
  );
  qs("lrPart")?.addEventListener("input", updateReceiveLubePartDesc);
  qs("lrPart")?.addEventListener("change", updateReceiveLubePartDesc);
  qs("icLoad")?.addEventListener("click", () =>
    loadInventoryControl().catch((e) => setStatus("Inventory control error: " + e.message))
  );
  qs("icSaveMin")?.addEventListener("click", () =>
    saveInventoryPartMinimum().catch((e) => setStatus("Part minimum error: " + e.message))
  );
  qs("icSubmitCount")?.addEventListener("click", () =>
    submitInventoryCycleCount().catch((e) => setStatus("Cycle count error: " + e.message))
  );
  qs("icPartCode")?.addEventListener("change", () =>
    loadInventoryControl().catch((e) => setStatus("Inventory control error: " + e.message))
  );
  qs("saveManualLube")?.addEventListener("click", () =>
    saveManualLube().catch((e) => setStatus("Manual lube error: " + e.message))
  );
  ["msLocation", "saLocation", "mlLocation"].forEach((id) => {
    qs(id)?.addEventListener("change", () => {
      const v = String(qs(id)?.value || "").trim().toUpperCase();
      if (!v) return;
      setRoleDefaultLocation(getSessionRole(), v);
      applyDefaultLocationsToInputs();
      if (id === "saLocation") {
        const binInput = qs("saBin");
        if (binInput) binInput.value = "";
        loadBinCodeOptionsForLocation(v, "saBinCodeOptions").catch(() => {});
      }
      if (id === "msLocation") {
        const binInput = qs("msBin");
        if (binInput) binInput.value = "";
        loadBinCodeOptionsForLocation(v, "msBinCodeOptions").catch(() => {});
      }
    });
  });
  qs("locLoad")?.addEventListener("click", () =>
    loadLocations().catch((e) => setStatus("Locations error: " + e.message))
  );
  qs("locShowInactive")?.addEventListener("change", () =>
    loadLocations().catch((e) => setStatus("Locations error: " + e.message))
  );
  qs("locSave")?.addEventListener("click", () =>
    saveLocation().catch((e) => setStatus("Location save error: " + e.message))
  );
  qs("sbSaveBinBtn")?.addEventListener("click", () =>
    saveStockBin().catch((e) => setStatus("Bin save error: " + e.message))
  );
  qs("sbLoadBinsBtn")?.addEventListener("click", () =>
    loadStockBins().catch((e) => setStatus("Bins load error: " + e.message))
  );
  qs("sbSaveMinMaxBtn")?.addEventListener("click", () =>
    saveStockMinMax().catch((e) => setStatus("Min-max save error: " + e.message))
  );
  qs("sbLoadDepthBtn")?.addEventListener("click", () =>
    loadStockDepth().catch((e) => setStatus("Depth load error: " + e.message))
  );
  qs("sbLoadReplenishmentBtn")?.addEventListener("click", () =>
    loadReplenishmentSuggestions().catch((e) => setStatus("Replenishment load error: " + e.message))
  );
  qs("sbExportReplenishmentCsvBtn")?.addEventListener("click", () => {
    try {
      exportReplenishmentSuggestionsCsv();
      setStatus("Replenishment CSV exported.");
    } catch (e) {
      setStatus("Replenishment export error: " + (e.message || e));
    }
  });
  qs("sbCreateCycleSessionBtn")?.addEventListener("click", () =>
    createCycleSession().catch((e) => setStatus("Cycle session create error: " + e.message))
  );
  qs("sbLoadCycleSessionsBtn")?.addEventListener("click", () =>
    loadCycleSessions().catch((e) => setStatus("Cycle sessions load error: " + e.message))
  );
  qs("loadLubeStock")?.addEventListener("click", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlPart")?.addEventListener("change", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlPart")?.addEventListener("input", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlType")?.addEventListener("change", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlType")?.addEventListener("input", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlQty")?.addEventListener("input", updateLubeQtyWarning);
  qs("setLubeMin210")?.addEventListener("click", () =>
    setLubeMinimumStock().catch((e) => setStatus("Lube minimum error: " + e.message))
  );
}

/** Start-up: Lube analytics list actions. Called once from init() in init.js. */
function wireLubeAnalyticsList() {
  const lubeAnalyticsList = qs("lubeAnalyticsList");
  if (lubeAnalyticsList) {
    lubeAnalyticsList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const btn = target.closest("[data-map-oil-key]");
      if (!(btn instanceof HTMLElement)) return;
      const oilKey = String(btn.getAttribute("data-map-oil-key") || "").trim();
      const partCode = String(btn.getAttribute("data-map-part-code") || "").trim();
      const oilEl = qs("lubeMapOilKey");
      const partEl = qs("lubeMapPartCode");
      if (oilEl) oilEl.value = oilKey;
      if (partEl) partEl.value = partCode;
      setStatus("Mapping fields pre-filled from selected lube row.");
    });
  }
}

/** Start-up: Cycle count session list actions. Called once from init() in init.js. */
function wireCycleCountList() {
  const sbCycleSessionsList = qs("sbCycleSessionsList");
  if (sbCycleSessionsList) {
    sbCycleSessionsList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const submitId = target.getAttribute("data-sb-cs-submit");
      const approveId = target.getAttribute("data-sb-cs-approve");
      const countOneId = target.getAttribute("data-sb-cs-countone");
      if (submitId) {
        submitCycleSession(submitId).catch((e) => setStatus(`Cycle submit error: ${e.message || e}`));
        return;
      }
      if (approveId) {
        approveCycleSession(approveId).catch((e) => setStatus(`Cycle approve error: ${e.message || e}`));
        return;
      }
      if (countOneId) {
        addOnePartCountToSession(countOneId).catch((e) => setStatus(`Cycle line upsert error: ${e.message || e}`));
      }
    });
  }
}

// ---- Reverse a receipt (delivery captured twice or by mistake) --------------
async function loadReceiptsForReversal() {
  const host = qs("rrList");
  if (!host) return;
  const part = (qs("rrPart")?.value || "").trim().split(" - ")[0];
  const days = qs("rrDays")?.value || "30";
  const onlyDup = Boolean(qs("rrOnlyDup")?.checked);
  host.innerHTML = `<p class="muted small">Loading receipts…</p>`;
  let rows;
  try {
    const q = new URLSearchParams({ days });
    if (part) q.set("part_code", part);
    rows = (await fetchJson(`${API}/api/stock/receipts?${q}`)).rows || [];
  } catch (e) {
    host.innerHTML = `<p class="muted small">${escapeHtml(e.message || String(e))}</p>`;
    return;
  }
  if (onlyDup) rows = rows.filter((r) => r.possible_duplicate_of && !r.reversed_by);
  if (!rows.length) {
    host.innerHTML = `<p class="muted small">${onlyDup ? "No possible duplicates found." : "No receipts in this period."}</p>`;
    return;
  }
  host.innerHTML = rows.map((r) => {
    const dup = r.possible_duplicate_of && !r.reversed_by;
    const action = r.reversed_by
      ? `<span class="muted small">Reversed (movement #${r.reversed_by})</span>`
      : r.pending_request_id
        ? `<span class="muted small">Waiting for approval (request #${r.pending_request_id})</span>`
        : `<button type="button" class="btn" data-rr-reverse="${r.id}" data-rr-label="${escapeHtml(`${r.part_code} +${Number(r.quantity)} (${r.reference || "no reference"})`)}">Reverse</button>`;
    return `
      <div class="rr-row${dup ? " rr-dup" : ""}">
        <div class="rr-main">
          <div><b>${escapeHtml(r.part_code)}</b> <span class="rr-qty">+${Number(r.quantity).toFixed(1)}</span></div>
          <div class="muted small">${escapeHtml(r.part_name || "")}</div>
          <div class="muted small">Received ${escapeHtml(String(r.created_at || "").slice(0, 16))} · #${r.id} · ref ${escapeHtml(r.reference || "—")}${r.location_code ? ` · ${escapeHtml([r.location_code, r.bin_code].filter(Boolean).join(" / "))}` : ""}</div>
          ${dup ? `<div class="rr-flag">Possible duplicate of receipt #${r.possible_duplicate_of}</div>` : ""}
        </div>
        <div class="rr-action">${action}</div>
      </div>`;
  }).join("");
}

async function requestReceiptReversal(id, label) {
  const reason = prompt(`Reverse receipt #${id}: ${label}\n\nWhy? (e.g. delivery captured twice)`);
  if (reason == null) return;
  if (!reason.trim()) return alert("Please give a reason.");
  try {
    const res = await fetchJson(`${API}/api/stock/movements/${id}/reverse`, { method: "POST", body: JSON.stringify({ reason: reason.trim() }) });
    setStatus(`Reversal of receipt #${id} sent for approval (request #${res.request_id}).`);
  } catch (e) {
    alert(e.message || String(e));
  }
  await loadReceiptsForReversal();
  if (typeof loadApprovalRequests === "function") loadApprovalRequests().catch(() => {});
}

(function initReceiptReversal() {
  qs("rrLoad")?.addEventListener("click", () => loadReceiptsForReversal());
  qs("rrOnlyDup")?.addEventListener("change", () => loadReceiptsForReversal());
  qs("rrPart")?.addEventListener("keydown", (e) => { if (e.key === "Enter") loadReceiptsForReversal(); });
  qs("rrList")?.addEventListener("click", (e) => {
    const b = e.target.closest("[data-rr-reverse]");
    if (b) requestReceiptReversal(b.dataset.rrReverse, b.dataset.rrLabel);
  });
})();
