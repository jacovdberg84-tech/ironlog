// IRONLOG/web/app/breakdowns.js — Breakdown ops, operational slips, short breakdowns.
// Part of the main app; index.html loads these files in order and they share one global scope.

async function loadBreakdownOpsOpen() {
  const list = qs("boOpenList");
  if (!list) return;
  const d = (qs("boOpenDate")?.value || "").trim();
  const q = d ? `?date=${encodeURIComponent(d)}` : "";
  setSkeleton("boOpenList", 1);
  try {
    const data = await fetchJson(`${API}/api/breakdowns/open-all${q}`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    list.innerHTML = "";
    if (!rows.length) {
      list.appendChild(item("<small>No open incidents for this filter.</small>"));
      return;
    }
    rows.forEach((r) => {
      const bid = Number(r.id || 0);
      const wo = r.primary_work_order_id != null ? Number(r.primary_work_order_id) : "";
      const desc = escapeHtml(r.description || "");
      const code = escapeHtml(r.asset_code || "");
      const woSt = escapeHtml(String(r.primary_work_order_status || ""));
      list.appendChild(
        item(
          `<div><b>#${bid}</b> <span class="pill blue">${code}</span> <span class="pill orange">OPEN</span></div>` +
            `<small>${desc}</small><br/>` +
            `<small>WO: ${wo ? `#${wo} (${woSt})` : "—"} | Start: ${escapeHtml(String(r.start_at || r.breakdown_date || "—"))}</small><br/>` +
            `<button type="button" class="bo-copy-wo" data-wo="${wo}">Copy WO #</button> ` +
            `<button type="button" class="bo-close-bdn" data-id="${bid}">Close incident</button>`
        )
      );
    });
  } catch (e) {
    list.innerHTML = "";
    list.appendChild(item(`<span class="message-error">${escapeHtml(e.message || String(e))}</span>`));
  }
}

async function loadBreakdownOpsRecent() {
  const list = qs("boRecentList");
  if (!list) return;
  setSkeleton("boRecentList", 1);
  try {
    const rows = await fetchJson(`${API}/api/breakdowns`);
    const slice = (Array.isArray(rows) ? rows : []).slice(0, 20);
    list.innerHTML = "";
    if (!slice.length) {
      list.appendChild(item("<small>No breakdowns found.</small>"));
      return;
    }
    slice.forEach((r) => {
      const st = String(r.status || "").toUpperCase();
      list.appendChild(
        item(
          `<b>#${r.id}</b> ${escapeHtml(r.asset_code || "")} <span class="pill ${st === "OPEN" ? "orange" : "blue"}">${escapeHtml(r.status || "")}</span> ` +
            `${escapeHtml(r.breakdown_date || "")}<br/><small>${escapeHtml(r.description || "")}</small>`
        )
      );
    });
  } catch (e) {
    list.innerHTML = "";
    list.appendChild(item(`<span class="message-error">${escapeHtml(e.message || String(e))}</span>`));
  }
}

function refreshBreakdownOpsPanels() {
  loadBreakdownOpsOpen().catch(() => {});
  loadBreakdownOpsRecent().catch(() => {});
}

function initBoTyreRows() {
  const tb = qs("boTyreTbody");
  if (!tb || tb.dataset.ready === "1") return;
  tb.dataset.ready = "1";
  for (let i = 0; i < 10; i++) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td><input class="w-90" id="boT${i}_pos" placeholder="e.g. FL" /></td>
      <td><input class="w-120" id="boT${i}_sout" /></td>
      <td><input class="w-120" id="boT${i}_sin" /></td>
      <td><input class="w-90" id="boT${i}_tread" /></td>
      <td><input class="w-140" id="boT${i}_reason" /></td>
      <td><input class="w-90" type="number" step="0.1" min="0" id="boT${i}_hu" /></td>
      <td><input class="w-90" type="number" step="0.1" min="0" id="boT${i}_hf" /></td>
      <td><input class="w-120" id="boT${i}_part" list="partCodeOptions" /></td>
      <td><input class="w-90" type="number" step="0.01" min="0" id="boT${i}_cost" placeholder="ov." /></td>
      <td><input class="w-120" id="boT${i}_make" list="partCodeOptions" placeholder="Make code" /></td>`;
    tb.appendChild(tr);
  }
}

function updateBoSlipFormVisibility() {
  const t = qs("boSlipType")?.value || "hose_failure";
  const map = {
    hose_failure: "boWrapHose",
    get_change: "boWrapGet",
    component_change: "boWrapComp",
    tyre_change: "boWrapTyre",
  };
  Object.values(map).forEach((id) => {
    const el = qs(id);
    if (el) el.style.display = "none";
  });
  const showId = map[t];
  if (showId && qs(showId)) qs(showId).style.display = "";
  if (t === "tyre_change") initBoTyreRows();
}

function numOrUndef(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

const BO_SLIP_PIC_MAX = 4;
const BO_SLIP_PIC_MAX_BYTES = 512 * 1024;
let boSlipPicturesPayload = [];

function clearBoSlipPhotosUi() {
  boSlipPicturesPayload = [];
  const inp = qs("boSlipPhotosInput");
  if (inp) inp.value = "";
  const prev = qs("boSlipPhotosPreview");
  if (prev) prev.innerHTML = "";
}

function renderBoSlipPhotosPreview() {
  const prev = qs("boSlipPhotosPreview");
  if (!prev) return;
  prev.innerHTML = "";
  boSlipPicturesPayload.forEach((p, idx) => {
    const wrap = document.createElement("div");
    wrap.style.cssText =
      "position:relative;border:1px solid #cbd5e1;border-radius:6px;padding:4px;background:#f8fafc;";
    const img = document.createElement("img");
    img.src = `data:${p.mime};base64,${p.data_base64}`;
    img.alt = "";
    img.style.cssText = "max-width:120px;max-height:90px;display:block;";
    const rm = document.createElement("button");
    rm.type = "button";
    rm.textContent = "Remove";
    rm.style.cssText = "font-size:11px;margin-top:4px;";
    rm.addEventListener("click", () => {
      boSlipPicturesPayload = boSlipPicturesPayload.filter((_, i) => i !== idx);
      renderBoSlipPhotosPreview();
    });
    wrap.appendChild(img);
    wrap.appendChild(rm);
    prev.appendChild(wrap);
  });
}

async function boSlipReadPictureFile(file) {
  const mime = String(file.type || "").toLowerCase();
  if (mime !== "image/jpeg" && mime !== "image/png") {
    alert(`${file.name}: only JPEG or PNG images are supported.`);
    return null;
  }
  if (file.size > BO_SLIP_PIC_MAX_BYTES) {
    alert(`${file.name} is larger than 512 KB. Choose a smaller file or compress the image.`);
    return null;
  }
  return new Promise((resolve, reject) => {
    const fr = new FileReader();
    fr.onload = () => {
      const s = String(fr.result || "");
      const i = s.indexOf(",");
      const b64 = i >= 0 ? s.slice(i + 1) : "";
      resolve(b64 ? { mime, data_base64: b64 } : null);
    };
    fr.onerror = () => reject(new Error("Read failed"));
    fr.readAsDataURL(file);
  });
}

async function onBoSlipPhotosInputChange(ev) {
  const input = ev.target;
  const files = input?.files ? Array.from(input.files) : [];
  if (!files.length) return;
  for (const f of files) {
    if (boSlipPicturesPayload.length >= BO_SLIP_PIC_MAX) {
      alert(`Maximum ${BO_SLIP_PIC_MAX} pictures.`);
      break;
    }
    try {
      const pic = await boSlipReadPictureFile(f);
      if (pic) boSlipPicturesPayload.push(pic);
    } catch {
      setStatus("Could not read one of the pictures.");
    }
  }
  input.value = "";
  renderBoSlipPhotosPreview();
}

function collectBoTyreRows() {
  const tyres = [];
  for (let i = 0; i < 10; i++) {
    const position = String(qs(`boT${i}_pos`)?.value || "").trim();
    const serial_removed = String(qs(`boT${i}_sout`)?.value || "").trim();
    const serial_new = String(qs(`boT${i}_sin`)?.value || "").trim();
    const tread_left = String(qs(`boT${i}_tread`)?.value || "").trim();
    const reason = String(qs(`boT${i}_reason`)?.value || "").trim();
    const hours_in_use = numOrUndef(qs(`boT${i}_hu`)?.value);
    const hours_fitted = numOrUndef(qs(`boT${i}_hf`)?.value);
    const part_code = String(qs(`boT${i}_part`)?.value || "").trim();
    const cost_manual = numOrUndef(qs(`boT${i}_cost`)?.value);
    const tyre_make_part_code = String(qs(`boT${i}_make`)?.value || "").trim();
    if (
      !position &&
      !serial_removed &&
      !serial_new &&
      !reason &&
      !part_code &&
      !tyre_make_part_code &&
      hours_in_use == null &&
      hours_fitted == null
    ) {
      continue;
    }
    tyres.push({
      position,
      serial_removed,
      serial_new,
      tread_left,
      reason,
      hours_in_use,
      hours_fitted,
      part_code,
      cost_manual,
      tyre_make_part_code,
    });
  }
  return tyres;
}

function boSlipQtyOrUndef(el) {
  const n = Number(el?.value);
  if (!Number.isFinite(n) || n <= 0) return undefined;
  return n;
}

async function pullBoSlipFromAsset() {
  const asset_code = String(qs("boSlipAsset")?.value || "").trim();
  if (!asset_code) {
    alert("Enter asset code first.");
    return;
  }
  const slip_type = String(qs("boSlipType")?.value || "").trim();
  setStatus("Loading asset slip hints...");
  try {
    const data = await fetchJson(
      `${API}/api/breakdown-ops/slip-asset-hints?asset_code=${encodeURIComponent(asset_code)}`
    );
    if (slip_type === "get_change" && data.get_change) {
      const h = data.get_change;
      const gh = qs("boGetHours");
      if (gh && h.hours_fitted != null && Number.isFinite(Number(h.hours_fitted))) gh.value = String(h.hours_fitted);
      if (qs("boGetPart")) qs("boGetPart").value = h.part_code || "";
      if (qs("boGetPartQty")) qs("boGetPartQty").value = String(h.part_qty != null ? h.part_qty : 1);
      if (qs("boGetSupplier")) qs("boGetSupplier").value = h.supplier || "";
      if (qs("boGetDateChg")) qs("boGetDateChg").value = h.date_changed || "";
      if (qs("boGetDescPart")) qs("boGetDescPart").value = h.description_part_code || "";
      if (qs("boGetDescPartQty")) qs("boGetDescPartQty").value = String(h.description_part_qty != null ? h.description_part_qty : 1);
      if (h.notes && qs("boGetNotes")) qs("boGetNotes").value = h.notes;
      setStatus(`G.E.T. fields filled from ${data.get_change_source || "asset"}.`);
    } else if (slip_type === "hose_failure" && data.hose_failure) {
      const h = data.hose_failure;
      if (qs("boHoseDateFitted")) qs("boHoseDateFitted").value = h.date_fitted || "";
      if (qs("boHoseReason")) qs("boHoseReason").value = h.reason_fitted || "";
      if (qs("boHosePreventable")) qs("boHosePreventable").checked = Boolean(h.preventable);
      if (qs("boHosePart")) qs("boHosePart").value = h.hose_part_code || "";
      if (qs("boHoseQty")) qs("boHoseQty").value = String(h.hose_qty != null ? h.hose_qty : 1);
      if (qs("boOilPart")) qs("boOilPart").value = h.oil_loss_part_code || "";
      if (qs("boOilQty")) qs("boOilQty").value = String(h.oil_loss_qty != null ? h.oil_loss_qty : 1);
      if (h.notes && qs("boHoseNotes")) qs("boHoseNotes").value = h.notes;
      setStatus(`Hose fields filled from ${data.hose_failure_source || "asset"}.`);
    } else {
      setStatus(
        slip_type === "get_change"
          ? "No GET data: add a GET change slip on the asset (Assets → history) or save a G.E.T. slip first."
          : slip_type === "hose_failure"
            ? "No prior hose failure slip for this asset yet."
            : "Pull from asset applies to G.E.T. or Hose failure slips."
      );
    }
  } catch (e) {
    setStatus("Asset hints failed: " + (e.message || e));
  }
}

async function saveBoSlipReport() {
  const slip_type = String(qs("boSlipType")?.value || "").trim();
  const asset_code = String(qs("boSlipAsset")?.value || "").trim();
  const report_date = String(qs("boSlipDate")?.value || "").trim();
  if (!asset_code || !report_date) {
    alert("Asset code and report date are required.");
    return;
  }
  let body = { slip_type, asset_code, report_date };
  if (slip_type === "hose_failure") {
    Object.assign(body, {
      date_fitted: String(qs("boHoseDateFitted")?.value || "").trim(),
      reason_fitted: String(qs("boHoseReason")?.value || "").trim(),
      preventable: Boolean(qs("boHosePreventable")?.checked),
      hose_part_code: String(qs("boHosePart")?.value || "").trim(),
      oil_loss_part_code: String(qs("boOilPart")?.value || "").trim(),
      hose_qty: boSlipQtyOrUndef(qs("boHoseQty")),
      oil_loss_qty: boSlipQtyOrUndef(qs("boOilQty")),
      hose_cost_manual: numOrUndef(qs("boHoseCostOv")?.value),
      oil_cost_manual: numOrUndef(qs("boOilCostOv")?.value),
      notes: String(qs("boHoseNotes")?.value || "").trim() || undefined,
    });
  } else if (slip_type === "get_change") {
    Object.assign(body, {
      hours_fitted: numOrUndef(qs("boGetHours")?.value),
      part_code: String(qs("boGetPart")?.value || "").trim(),
      part_qty: boSlipQtyOrUndef(qs("boGetPartQty")),
      supplier: String(qs("boGetSupplier")?.value || "").trim(),
      date_changed: String(qs("boGetDateChg")?.value || "").trim(),
      description_part_code: String(qs("boGetDescPart")?.value || "").trim(),
      description_part_qty: boSlipQtyOrUndef(qs("boGetDescPartQty")),
      notes: String(qs("boGetNotes")?.value || "").trim() || undefined,
    });
  } else if (slip_type === "component_change") {
    Object.assign(body, {
      date_changed: String(qs("boCompDate")?.value || "").trim(),
      hours_in_service: numOrUndef(qs("boCompHrsSvc")?.value),
      reason: String(qs("boCompReason")?.value || "").trim(),
      component_type: String(qs("boCompType")?.value || "").trim(),
      part_code: String(qs("boCompPart")?.value || "").trim(),
      cost_manual: numOrUndef(qs("boCompCostOv")?.value),
      notes: String(qs("boCompNotes")?.value || "").trim() || undefined,
    });
  } else if (slip_type === "tyre_change") {
    const tyres = collectBoTyreRows();
    if (!tyres.length) {
      alert("Enter at least one tyre line.");
      return;
    }
    Object.assign(body, { tyres, notes: String(qs("boTyreNotes")?.value || "").trim() || undefined });
  } else {
    alert("Unknown slip type.");
    return;
  }

  if (boSlipPicturesPayload.length) {
    body.pictures = boSlipPicturesPayload.map((p) => ({
      mime: p.mime,
      data_base64: p.data_base64,
    }));
  }

  setStatus("Saving slip report...");
  try {
    const res = await fetchJson(`${API}/api/breakdown-ops/slips`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    setText("boSlipResult", JSON.stringify(res, null, 2));
    setStatus("Slip saved.");
    clearBoSlipPhotosUi();
    if (res.id) {
      await openAuthedReport(`${API}/api/breakdown-ops/slips/${res.id}/pdf`, {
        filename: `IRONLOG_Operational_Slip_${res.id}.pdf`,
      });
    }
    await loadBoSlipSavedList();
  } catch (e) {
    setText("boSlipResult", String(e.message || e));
    setStatus("Slip save failed.");
  }
}

async function loadBoSlipSavedList() {
  const list = qs("boSlipSavedList");
  if (!list) return;
  list.innerHTML = `<div class="muted">Loading…</div>`;
  try {
    const data = await fetchJson(`${API}/api/breakdown-ops/slips`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    list.innerHTML = "";
    if (!rows.length) {
      list.appendChild(item("<small>No slip reports saved yet.</small>"));
      return;
    }
    rows.slice(0, 40).forEach((r) => {
      list.appendChild(renderBoSlipSavedRow(r));
    });
  } catch (e) {
    list.innerHTML = "";
    list.appendChild(item(`<span class="message-error">${escapeHtml(e.message || String(e))}</span>`));
  }
}

function renderBoSlipSavedRow(r) {
  const el = document.createElement("div");
  el.className = "bo-slip-row";
  const label = escapeHtml(String(r.slip_type || "").replace(/_/g, " "));
  el.innerHTML = `
    <div class="bo-slip-row-main">
      <span class="bo-slip-id">#${Number(r.id)}</span>
      <span class="pill blue">${label}</span>
      <span class="bo-slip-asset">${escapeHtml(r.asset_code || "—")}</span>
      <span class="bo-slip-date">${escapeHtml(r.report_date || "")}</span>
    </div>
    <div class="bo-slip-row-actions">
      <button type="button" class="bo-slip-pdf btn btn-secondary btn-sm" data-id="${Number(r.id)}">Open PDF</button>
    </div>
  `;
  return el;
}

function openBoSlipPdf(id) {
  const n = Number(id || 0);
  if (!n) return;
  return openAuthedReport(`${API}/api/breakdown-ops/slips/${n}/pdf`, {
    filename: `IRONLOG_Operational_Slip_${n}.pdf`,
  });
}

async function ensureOpenBreakdownOps() {
  const asset_code = (qs("boEnsureAsset")?.value || "").trim();
  const breakdown_date = (qs("boEnsureDate")?.value || "").trim() || new Date().toISOString().slice(0, 10);
  if (!asset_code) return alert("Enter asset code.");
  setStatus("Ensuring open breakdown...");
  try {
    const res = await fetchJson(`${API}/api/breakdowns/ensure-open`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ asset_code, breakdown_date, description: "Down - Breakdown Ops" }),
    });
    setText("boEnsureResult", JSON.stringify(res, null, 2));
    if (res.primary_work_order_id && qs("iWo")) qs("iWo").value = String(res.primary_work_order_id);
    setStatus("Ensure open complete.");
    await loadBreakdownOpsOpen();
    await loadBreakdownOpsRecent();
  } catch (e) {
    setText("boEnsureResult", String(e.message || e));
    setStatus("Ensure open failed.");
  }
}

let lastBoRepairWoId = null;

async function createRepairWorkOrderOps() {
  const asset_code = (qs("boRepairAsset")?.value || "").trim();
  const component = (qs("boRepairComponent")?.value || "").trim();
  const description = (qs("boRepairDescription")?.value || "").trim();
  const inspectionRaw = (qs("boRepairInspectionId")?.value || "").trim();
  const inspection_id = inspectionRaw ? Number(inspectionRaw) : 0;
  if (!asset_code) return alert("Enter asset code.");
  if (!description) return alert("Enter repair description.");
  setStatus("Creating repair work order…");
  try {
    const body = { asset_code, description };
    if (component) body.component = component;
    if (Number.isFinite(inspection_id) && inspection_id > 0) body.inspection_id = inspection_id;
    const res = await fetchJson(`${API}/api/workorders/repair`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    lastBoRepairWoId = Number(res.work_order_id || 0) || null;
    setText("boRepairResult", JSON.stringify(res, null, 2));
    const openBtn = qs("boRepairOpenWo");
    if (openBtn && lastBoRepairWoId) openBtn.style.display = "";
    setStatus(
      res.already_exists
        ? `WO #${lastBoRepairWoId} already linked to inspection.`
        : `Repair WO #${lastBoRepairWoId} created (no breakdown).`
    );
  } catch (e) {
    setText("boRepairResult", String(e.message || e));
    setStatus("Repair WO create failed.");
  }
}

async function pullBreakdownOpsLiveHours() {
  const hint = qs("boLiveHoursHint");
  const code = (qs("sqAsset")?.value || "").trim();
  const asOf =
    (qs("sqDate")?.value || "").trim() || (qs("date")?.value || "").trim() || new Date().toISOString().slice(0, 10);
  if (!code) {
    alert("Enter asset code on the short breakdown line first.");
    return;
  }
  if (hint) hint.textContent = "Loading…";
  try {
    const rows = await fetchJson(`${API}/api/assets?include_archived=0`);
    const arr = Array.isArray(rows) ? rows : [];
    const a = arr.find((x) => String(x.asset_code || "").toUpperCase() === code.toUpperCase());
    if (!a) throw new Error("Asset not found");
    const q = asOf ? `?as_of=${encodeURIComponent(asOf)}` : "";
    const data = await fetchJson(`${API}/api/maintenance/asset/${a.id}/live-hours${q}`);
    const h = Number(data.current_hours || 0);
    const src = String(data.current_hours_source || "");
    if (hint) {
      hint.textContent = `Meter (as of ${asOf}): ${h.toFixed(1)} h (${src}).`;
    }
    setStatus("Live meter loaded.");
  } catch (e) {
    if (hint) hint.textContent = e.message || String(e);
  }
}

async function closeBreakdownFromOps(breakdownId) {
  const id = Number(breakdownId || 0);
  if (!id) return;
  if (!confirm(`Close breakdown #${id}? Component work orders must be closed first.`)) return;
  setStatus("Closing breakdown...");
  try {
    await fetchJson(`${API}/api/breakdowns/${id}/close`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({}),
    });
    setStatus("Breakdown closed.");
    await loadBreakdownOpsOpen();
    await loadBreakdownOpsRecent();
    await loadDashboard().catch(() => {});
  } catch (e) {
    alert(e.message || String(e));
    setStatus("Close failed.");
  }
}

async function createBreakdown() {
  const date = (qs("bDate")?.value || qs("date")?.value || new Date().toISOString().slice(0, 10)).trim();
  const payload = {
    asset_code: (qs("bAsset")?.value || "").trim(),
    breakdown_date: date,
    time_down: qs("bTime")?.value ? `${date}T${qs("bTime").value}` : null,
    description: (qs("bDesc")?.value || "").trim(),
    downtime_hours: Number(qs("bDown")?.value || 0),
    critical: !!qs("bCrit")?.checked,
    parts_ordered_date: String(qs("bPartsOrderedDate")?.value || "").trim() || null,
    parts_status: String(qs("bPartsStatus")?.value || "").trim() || null,
    parts_received_date: String(qs("bPartsReceivedDate")?.value || "").trim() || null,
    ets_repair_date: String(qs("bEtsRepairDate")?.value || "").trim() || null,
  };

  setStatus("Creating breakdown...");
  try {
    const res = await fetchJson(`${API}/api/breakdowns`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("breakdownResult", JSON.stringify(res, null, 2));
    setStatus("Breakdown created.");
    await loadDashboard().catch(() => {});
    refreshBreakdownOpsPanels();
  } catch (e) {
    setText("breakdownResult", String(e.message || e));
    setStatus("Breakdown failed.");
  }
}

/** Short breakdown: closed incident + downtime logs + parts + oils in one POST. */
function collectShortBreakdownParts() {
  const parts = [];
  document.querySelectorAll("#sqPartsRows .sq-part-row").forEach((row) => {
    const part_code = String(row.querySelector(".sq-part-code")?.value || "").trim();
    const quantity = Number(row.querySelector(".sq-part-qty")?.value || 0);
    if (part_code && Number.isFinite(quantity) && quantity > 0) {
      parts.push({ part_code, quantity });
    }
  });
  return parts;
}

function addShortBreakdownPartRow() {
  const container = qs("sqPartsRows");
  if (!container) return;
  const row = document.createElement("div");
  row.className = "row sq-part-row";
  row.innerHTML = `
    <input class="sq-part-code w-180" list="partCodeOptions" placeholder="Part code" />
    <input class="sq-part-qty w-70" type="number" min="1" step="1" placeholder="Qty" />
    <button type="button" class="sq-part-remove" title="Remove part line">Remove</button>
  `;
  container.appendChild(row);
}

function bindShortBreakdownPartsUi() {
  qs("sqAddPart")?.addEventListener("click", () => addShortBreakdownPartRow());
  qs("sqPartsRows")?.addEventListener("click", (ev) => {
    const btn = ev.target?.closest?.(".sq-part-remove");
    if (!btn) return;
    const row = btn.closest(".sq-part-row");
    const container = qs("sqPartsRows");
    if (!row || !container) return;
    if (container.querySelectorAll(".sq-part-row").length <= 1) {
      row.querySelector(".sq-part-code").value = "";
      row.querySelector(".sq-part-qty").value = "";
      return;
    }
    row.remove();
  });
}

async function submitShortBreakdown() {
  const headerDate = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const breakdown_date = (qs("sqDate")?.value || headerDate).trim();
  const asset_code = (qs("sqAsset")?.value || "").trim();
  const description = (qs("sqDesc")?.value || "").trim();
  const td = (qs("sqTimeDown")?.value || "").trim();
  const tu = (qs("sqTimeUp")?.value || "").trim();
  const comp = (qs("sqComponent")?.value || "").trim();

  const parts = collectShortBreakdownParts();

  const oils = [];
  const oilType = (qs("sqOilType")?.value || "").trim();
  const oilQty = Number(qs("sqOilQty")?.value || 0);
  if (oilType && Number.isFinite(oilQty) && oilQty > 0) oils.push({ oil_type: oilType, quantity: oilQty });

  if (!asset_code || !description) {
    alert("Asset code and description are required.");
    return;
  }

  const payload = {
    asset_code,
    breakdown_date,
    description,
    critical: !!qs("sqCrit")?.checked,
    parts,
    oils,
  };
  if (comp) payload.component = comp;

  if (td && tu) {
    payload.time_down = td;
    payload.time_up = tu;
  } else if (!td && !tu) {
    const h = Number(qs("sqHours")?.value);
    if (Number.isNaN(h) || h <= 0 || h > 24) {
      alert("Enter both time down and time up, or a single hours-down value (0–24) for the event date.");
      return;
    }
    payload.hours_down = h;
  } else {
    alert("Provide both time down and time up, or clear both and use hours down.");
    return;
  }

  setStatus("Logging short breakdown...");
  try {
    const res = await fetchJson(`${API}/api/breakdowns/short-complete`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    setText("shortBreakdownResult", JSON.stringify(res, null, 2));
    setStatus("Short breakdown logged.");
    await loadDashboard().catch(() => {});
    refreshBreakdownOpsPanels();
  } catch (e) {
    setText("shortBreakdownResult", String(e.message || e));
    setStatus("Short breakdown failed.");
  }
}

/** Start-up: Breakdown ops, slip report and short breakdown controls. Called once from init() in init.js. */
function wireBreakdownControls() {
  qs("boRefreshOpen")?.addEventListener("click", () =>
    loadBreakdownOpsOpen().catch((e) => setStatus("Open list error: " + e.message))
  );
  qs("boRefreshRecent")?.addEventListener("click", () =>
    loadBreakdownOpsRecent().catch((e) => setStatus("Recent list error: " + e.message))
  );
  qs("boEnsureOpen")?.addEventListener("click", () =>
    ensureOpenBreakdownOps().catch((e) => setStatus("Ensure open error: " + e.message))
  );
  qs("boRepairCreateWo")?.addEventListener("click", () =>
    createRepairWorkOrderOps().catch((e) => setStatus("Repair WO error: " + e.message))
  );
  qs("boRepairOpenWo")?.addEventListener("click", () => {
    if (lastBoRepairWoId) window.open(`/web/workorders.html?wo=${encodeURIComponent(String(lastBoRepairWoId))}`, "_blank");
  });
  qs("boPullLiveHours")?.addEventListener("click", () =>
    pullBreakdownOpsLiveHours().catch((e) => setStatus("Live hours error: " + e.message))
  );
  qs("boOpenList")?.addEventListener("click", (ev) => {
    const w = ev.target?.closest?.(".bo-copy-wo");
    if (w) {
      const wo = w.getAttribute("data-wo");
      if (wo && qs("iWo")) qs("iWo").value = String(wo);
      setStatus(`Copied WO #${wo} for parts issue.`);
      return;
    }
    const c = ev.target?.closest?.(".bo-close-bdn");
    if (c) closeBreakdownFromOps(c.getAttribute("data-id")).catch(() => {});
  });

  qs("boSlipType")?.addEventListener("change", updateBoSlipFormVisibility);
  qs("boSlipPhotosInput")?.addEventListener("change", (e) =>
    onBoSlipPhotosInputChange(e).catch((err) => setStatus(String(err.message || err)))
  );
  qs("boSlipPhotosClear")?.addEventListener("click", () => {
    clearBoSlipPhotosUi();
    setStatus("Slip pictures cleared.");
  });
  qs("boSlipPullAsset")?.addEventListener("click", () =>
    pullBoSlipFromAsset().catch((e) => setStatus("Pull from asset error: " + (e.message || e)))
  );
  qs("boSlipSave")?.addEventListener("click", () =>
    saveBoSlipReport().catch((e) => setStatus("Slip save error: " + e.message))
  );
  qs("boSlipLoadList")?.addEventListener("click", () =>
    loadBoSlipSavedList().catch((e) => setStatus("Slip list error: " + e.message))
  );
  qs("boSlipSavedList")?.addEventListener("click", (ev) => {
    const b = ev.target?.closest?.(".bo-slip-pdf");
    if (b) openBoSlipPdf(b.getAttribute("data-id"));
  });

  qs("makeBreakdown")?.addEventListener("click", () =>
    createBreakdown().catch((e) => setStatus("Breakdown error: " + e.message))
  );
  bindShortBreakdownPartsUi();
  qs("sqSubmit")?.addEventListener("click", () =>
    submitShortBreakdown().catch((e) => setStatus("Short breakdown error: " + e.message))
  );
}
