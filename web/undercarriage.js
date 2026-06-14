(function () {
  const API = "/api";
  let ucSchema = [];
  let ucChecklistItems = [];
  let ucTrackSagPoints = ["A", "B", "C", "D"];
  let ucWearBands = [];
  let ucMobileMode = false;

  function esc(v) {
    return String(v == null ? "" : v)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;");
  }

  function authHeaders(extra = {}) {
    if (typeof window.authHeaders === "function") return window.authHeaders(extra);
    const h = { ...extra };
    try {
      h["x-site-code"] = localStorage.getItem("ironlog_session_site") || "main";
      h["x-user-name"] = localStorage.getItem("ironlog_session_user") || "system";
      h["x-user-role"] = localStorage.getItem("ironlog_session_role") || "admin";
      h["x-user-roles"] = localStorage.getItem("ironlog_session_roles") || "admin";
      const tok = localStorage.getItem("ironlog_auth_token");
      if (tok) h.Authorization = `Bearer ${tok}`;
    } catch {}
    return h;
  }

  function ucInputId(key, field) {
    return `uc_${key}_${field}`;
  }

  function ucReadNumber(id) {
    const raw = String(document.getElementById(id)?.value || "").trim();
    if (raw === "") return null;
    const n = Number(raw);
    return Number.isFinite(n) ? n : null;
  }

  function calcWearPct(measurement, base, wearLimit, wearDirection) {
    const m = measurement;
    const b = base;
    const w = wearLimit;
    if (m == null || b == null || w == null) return null;
    let pct;
    if (String(wearDirection || "down").toLowerCase() === "up") {
      const denom = w - b;
      if (!Number.isFinite(denom) || denom <= 0) return null;
      pct = ((m - b) / denom) * 100;
    } else {
      const denom = b - w;
      if (!Number.isFinite(denom) || denom <= 0) return null;
      pct = ((b - m) / denom) * 100;
    }
    return Number(Math.max(0, pct).toFixed(2));
  }

  function wearBandStyle(pct) {
    const n = pct == null ? null : Number(pct);
    if (n == null) return { bg: "#f1f5f9", label: "—" };
    if (n <= 75) return { bg: "#22c55e", label: "Good", color: "#fff" };
    if (n <= 100) return { bg: "#eab308", label: "Monitor", color: "#1e293b" };
    if (n <= 120) return { bg: "#f97316", label: "Plan replace", color: "#fff" };
    return { bg: "#ef4444", label: "Replace", color: "#fff" };
  }

  function ucGroups() {
    const order = [
      "bushings", "links", "pins", "track_shoe", "carrier_rollers", "track_rollers", "grouser_height",
    ];
    const labels = {
      bushings: "Bushings",
      links: "Links",
      pins: "Pins",
      track_shoe: "Track Shoe",
      carrier_rollers: "Carrier Rollers",
      track_rollers: "Track Rollers (1–12)",
      grouser_height: "Grouser Height",
    };
    const map = new Map();
    for (const row of ucSchema) {
      const g = row.group || "other";
      if (!map.has(g)) map.set(g, []);
      map.get(g).push(row);
    }
    return order.filter((g) => map.has(g)).map((g) => ({ key: g, label: labels[g] || g, rows: map.get(g) }));
  }

  function renderWearPreview(key, row) {
    const m = ucReadNumber(ucInputId(key, "measurement"));
    const b = ucReadNumber(ucInputId(key, "base"));
    const w = ucReadNumber(ucInputId(key, "limit"));
    const pct = calcWearPct(m, b, w, row.wear_direction);
    const band = wearBandStyle(pct);
    const el = document.getElementById(ucInputId(key, "wear_preview"));
    if (!el) return;
    el.innerHTML = pct == null
      ? `<span class="uc-wear-chip muted">Enter measurement</span>`
      : `<span class="uc-wear-chip" style="background:${band.bg};color:${band.color || "#fff"}">${Number(pct).toFixed(1)}% · ${band.label}</span>`;
  }

  function renderMeasurementRow(row) {
    const key = row.key;
    return `
      <div class="uc-measure-row" data-uc-key="${esc(key)}">
        <div class="uc-measure-label">
          <strong>${esc(row.label)}</strong>
          <span class="muted mini">${esc(row.side)}</span>
        </div>
        <label>Meas (mm)<input id="${ucInputId(key, "measurement")}" type="number" step="0.1" inputmode="decimal" class="uc-num-input" /></label>
        <label>Base<input id="${ucInputId(key, "base")}" type="number" step="0.1" inputmode="decimal" class="uc-num-input" /></label>
        <label>Limit<input id="${ucInputId(key, "limit")}" type="number" step="0.1" inputmode="decimal" class="uc-num-input" /></label>
        <div id="${ucInputId(key, "wear_preview")}" class="uc-wear-preview"></div>
      </div>
    `;
  }

  function renderForm() {
    const mount = document.getElementById("ucFormMount");
    if (!mount || !ucSchema.length) return;

    const groups = ucGroups();
    mount.innerHTML = `
      <div class="uc-legend">
        <span class="uc-wear-chip" style="background:#22c55e">0–75% Good</span>
        <span class="uc-wear-chip" style="background:#eab308;color:#1e293b">76–100%</span>
        <span class="uc-wear-chip" style="background:#f97316">101–120%</span>
        <span class="uc-wear-chip" style="background:#ef4444">&gt;120%</span>
      </div>
      ${groups.map((g) => `
        <details class="uc-group" open="${g.key === "bushings" ? "open" : ""}">
          <summary>${esc(g.label)}</summary>
          <div class="uc-group-body">
            ${g.rows.map((r) => renderMeasurementRow(r)).join("")}
          </div>
        </details>
      `).join("")}
      <details class="uc-group">
        <summary>Track sag (mm)</summary>
        <div class="uc-group-body uc-track-sag-grid">
          ${ucTrackSagPoints.map((p) => `
            <label>${esc(p)}<input id="uc_sag_${p}" type="number" step="0.1" inputmode="decimal" class="uc-num-input" /></label>
          `).join("")}
        </div>
      </details>
      <details class="uc-group">
        <summary>Condition checklist</summary>
        <div class="uc-group-body">
          ${ucChecklistItems.map((item) => `
            <div class="uc-check-row">
              <div class="uc-check-label">${esc(item.label)}</div>
              <label><input id="uc_check_${item.key}_lh" type="checkbox" /> LH issue</label>
              <label><input id="uc_check_${item.key}_rh" type="checkbox" /> RH issue</label>
            </div>
          `).join("")}
          <label style="margin-top:10px;">General condition
            <select id="ucGeneralCondition">
              <option value="">—</option>
              <option value="good">Good</option>
              <option value="fair">Fair</option>
              <option value="poor">Poor</option>
            </select>
          </label>
          <label style="margin-top:8px;">Comments<textarea id="ucComments" rows="3" placeholder="Observations"></textarea></label>
        </div>
      </details>
    `;

    mount.querySelectorAll(".uc-num-input").forEach((el) => {
      el.addEventListener("input", () => {
        const row = el.closest(".uc-measure-row");
        const key = row?.getAttribute("data-uc-key");
        const schemaRow = ucSchema.find((s) => s.key === key);
        if (schemaRow) renderWearPreview(key, schemaRow);
      });
    });
    for (const row of ucSchema) renderWearPreview(row.key, row);
  }

  function collectMeasurements() {
    return ucSchema.map((row) => ({
      key: row.key,
      measurement: ucReadNumber(ucInputId(row.key, "measurement")),
      base: ucReadNumber(ucInputId(row.key, "base")),
      wear_limit: ucReadNumber(ucInputId(row.key, "limit")),
    }));
  }

  function collectTrackSag() {
    const out = {};
    for (const p of ucTrackSagPoints) out[p] = ucReadNumber(`uc_sag_${p}`);
    return out;
  }

  function collectChecklist() {
    const items = {};
    for (const item of ucChecklistItems) {
      items[item.key] = {
        lh: Boolean(document.getElementById(`uc_check_${item.key}_lh`)?.checked),
        rh: Boolean(document.getElementById(`uc_check_${item.key}_rh`)?.checked),
      };
    }
    return {
      items,
      general_condition: String(document.getElementById("ucGeneralCondition")?.value || "").trim(),
      comments: String(document.getElementById("ucComments")?.value || "").trim(),
    };
  }

  function fillMeasurements(rows) {
    const list = Array.isArray(rows) ? rows : [];
    const byKey = new Map(list.map((r) => [String(r.key || "").toLowerCase(), r]));
    for (const row of ucSchema) {
      const src = byKey.get(String(row.key).toLowerCase()) || {};
      const meas = document.getElementById(ucInputId(row.key, "measurement"));
      const base = document.getElementById(ucInputId(row.key, "base"));
      const limit = document.getElementById(ucInputId(row.key, "limit"));
      if (meas) meas.value = src.measurement ?? "";
      if (base) base.value = src.base ?? "";
      if (limit) limit.value = src.wear_limit ?? "";
      renderWearPreview(row.key, row);
    }
  }

  function fillFromLatest(data) {
    if (!data) return;
    fillMeasurements(data.measurements || []);
    const sag = data.track_sag || {};
    for (const p of ucTrackSagPoints) {
      const el = document.getElementById(`uc_sag_${p}`);
      if (el) el.value = sag[p] ?? "";
    }
    const fields = [
      ["ucSerialNo", data.serial_no],
      ["ucModel", data.model],
      ["ucJobNo", data.job_no],
      ["ucPlanner", data.planner],
      ["ucUnitAssembly", data.unit_assembly],
      ["ucYardNo", data.yard_no],
      ["ucWorkOrderNo", data.work_order_no],
      ["ucComponentGroup", data.component_group],
      ["ucGroupId", data.group_id],
      ["ucComponentSerial", data.component_serial_no],
      ["ucPartNo", data.part_no],
      ["ucCostCenter", data.cost_center],
    ];
    for (const [id, val] of fields) {
      const el = document.getElementById(id);
      if (el && val != null && String(el.value || "").trim() === "") el.value = val;
    }
    // Prefill base/limit from previous; clear measurements for new reading
    for (const row of ucSchema) {
      const meas = document.getElementById(ucInputId(row.key, "measurement"));
      if (meas) meas.value = "";
      renderWearPreview(row.key, row);
    }
  }

  async function loadTemplate() {
    const res = await fetch(`${API}/maintenance/undercarriage-inspections/template`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load template");
    ucSchema = Array.isArray(data.components) ? data.components : [];
    ucChecklistItems = Array.isArray(data.checklist_items) ? data.checklist_items : [];
    ucTrackSagPoints = Array.isArray(data.track_sag_points) ? data.track_sag_points : ucTrackSagPoints;
    ucWearBands = Array.isArray(data.wear_bands) ? data.wear_bands : [];
    renderForm();
  }

  async function loadLatestForAsset(assetId) {
    if (!assetId) return;
    try {
      const res = await fetch(`${API}/maintenance/undercarriage-inspections/latest?asset_id=${assetId}`, { headers: authHeaders() });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Failed to load previous inspection");
      fillFromLatest(data.row);
    } catch {}
  }

  async function pullLiveSmu() {
    const assetId = Number(document.getElementById("ucAsset")?.value || 0);
    const out = document.getElementById("ucSmuMeta");
    if (!assetId) {
      if (out) out.textContent = "Select a machine first.";
      return;
    }
    if (out) out.textContent = "Loading live hours…";
    try {
      const res = await fetch(`${API}/maintenance/asset/${assetId}/live-hours`, { headers: authHeaders() });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Failed to load live hours");
      const hrs = Number(data.hours ?? data.total_hours ?? 0);
      const smu = document.getElementById("ucSmu");
      if (smu) smu.value = Number.isFinite(hrs) ? hrs.toFixed(1) : "";
      if (out) out.textContent = `Live hours: ${Number.isFinite(hrs) ? hrs.toFixed(1) : "—"} (${data.source || "snapshot"})`;
    } catch (e) {
      if (out) out.textContent = `Live hours failed: ${e.message || e}`;
    }
  }

  function renderSavedList(rows) {
    const list = document.getElementById("ucSavedList");
    if (!list) return;
    if (!rows.length) {
      list.innerHTML = `<div class="empty">No undercarriage inspections saved yet.</div>`;
      return;
    }
    list.innerHTML = rows.map((r) => {
      const summary = r.summary || {};
      const worst = summary.worst_wear_pct != null ? `${Number(summary.worst_wear_pct).toFixed(1)}%` : "—";
      const pdf = `${API}/maintenance/undercarriage-inspections/${Number(r.id)}.pdf`;
      const xlsx = `${API}/maintenance/undercarriage-inspections/report.xlsx?inspection_id=${Number(r.id)}`;
      return `
        <div class="item">
          <div class="title">${esc(r.asset_code || "-")} — ${esc(r.inspection_date || "")}</div>
          <div class="meta">
            SMU ${r.smu ?? "—"} | Inspector ${esc(r.inspector_name || "-")} |
            Worst wear <b>${worst}</b>${summary.worst_component ? ` (${esc(summary.worst_component)})` : ""}
          </div>
          <div class="row stack-10" style="margin-top:8px; flex-wrap:wrap;">
            <a class="btn" href="${esc(pdf)}" target="_blank" rel="noopener">PDF</a>
            <button type="button" class="btn" data-uc-xlsx="${Number(r.id)}">XLSX</button>
          </div>
        </div>
      `;
    }).join("");
  }

  async function loadSavedList() {
    const list = document.getElementById("ucSavedList");
    if (list) list.innerHTML = `<div class="empty">Loading…</div>`;
    const assetId = String(document.getElementById("ucFilterAsset")?.value || "").trim();
    const q = new URLSearchParams();
    if (assetId) q.set("asset_id", assetId);
    try {
      const res = await fetch(`${API}/maintenance/undercarriage-inspections?${q.toString()}`, { headers: authHeaders() });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Load failed");
      renderSavedList(Array.isArray(data.rows) ? data.rows : []);
    } catch (e) {
      if (list) list.innerHTML = `<div class="empty">${esc(e.message || String(e))}</div>`;
    }
  }

  async function saveInspection() {
    const msg = document.getElementById("ucMsg");
    const asset_id = Number(document.getElementById("ucAsset")?.value || 0);
    const inspection_date = String(document.getElementById("ucInspectionDate")?.value || "").trim();
    const smuRaw = String(document.getElementById("ucSmu")?.value || "").trim();
    if (!asset_id) {
      if (msg) msg.textContent = "Select a machine.";
      return;
    }
    if (!inspection_date) {
      if (msg) msg.textContent = "Inspection date is required.";
      return;
    }
    if (msg) msg.textContent = "Saving…";
    const body = {
      asset_id,
      inspection_date,
      inspector_name: String(document.getElementById("ucInspector")?.value || "").trim(),
      smu: smuRaw === "" ? null : Number(smuRaw),
      job_no: String(document.getElementById("ucJobNo")?.value || "").trim(),
      planner: String(document.getElementById("ucPlanner")?.value || "").trim(),
      serial_no: String(document.getElementById("ucSerialNo")?.value || "").trim(),
      unit_assembly: String(document.getElementById("ucUnitAssembly")?.value || "").trim(),
      model: String(document.getElementById("ucModel")?.value || "").trim(),
      yard_no: String(document.getElementById("ucYardNo")?.value || "").trim(),
      work_order_no: String(document.getElementById("ucWorkOrderNo")?.value || "").trim(),
      component_group: String(document.getElementById("ucComponentGroup")?.value || "").trim(),
      group_id: String(document.getElementById("ucGroupId")?.value || "").trim(),
      component_serial_no: String(document.getElementById("ucComponentSerial")?.value || "").trim(),
      part_no: String(document.getElementById("ucPartNo")?.value || "").trim(),
      cost_center: String(document.getElementById("ucCostCenter")?.value || "").trim(),
      measurements: collectMeasurements(),
      track_sag: collectTrackSag(),
      checklist: collectChecklist(),
    };
    try {
      const res = await fetch(`${API}/maintenance/undercarriage-inspections`, {
        method: "POST",
        headers: { ...authHeaders(), "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Save failed");
      if (msg) {
        const worst = data.summary?.worst_wear_pct;
        msg.textContent = worst != null
          ? `Saved. Highest wear ${Number(worst).toFixed(1)}% on ${data.summary?.worst_component || "component"}.`
          : "Undercarriage inspection saved to machine history.";
      }
      if (data.pdf_url) window.open(data.pdf_url, "_blank");
      loadSavedList();
    } catch (e) {
      if (msg) msg.textContent = `Save failed: ${e.message || e}`;
    }
  }

  async function exportXlsx(inspectionId) {
    const q = inspectionId ? `?inspection_id=${inspectionId}` : `?month=${encodeURIComponent(String(document.getElementById("ucReportMonth")?.value || "").slice(0, 7))}`;
    const res = await fetch(`${API}/maintenance/undercarriage-inspections/report.xlsx${q}`, { headers: authHeaders() });
    if (!res.ok) throw new Error(await res.text() || "Export failed");
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `IRONLOG_Undercarriage${inspectionId ? `_${inspectionId}` : ""}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  async function initUndercarriage(opts = {}) {
    ucMobileMode = Boolean(opts.mobile);
    document.body.classList.toggle("uc-mobile-mode", ucMobileMode);

    const dateEl = document.getElementById("ucInspectionDate");
    if (dateEl && !dateEl.value) dateEl.value = new Date().toISOString().slice(0, 10);
    const monthEl = document.getElementById("ucReportMonth");
    if (monthEl && !monthEl.value) monthEl.value = new Date().toISOString().slice(0, 7);

    try {
      await loadTemplate();
    } catch (e) {
      const mount = document.getElementById("ucFormMount");
      if (mount) mount.innerHTML = `<div class="empty">${esc(e.message || String(e))}</div>`;
      return;
    }

    const assetFromUrl = String(opts.assetId || opts.assetCode || "").trim();
    if (assetFromUrl && document.getElementById("ucAsset")) {
      const sel = document.getElementById("ucAsset");
      if (/^\d+$/.test(assetFromUrl)) sel.value = assetFromUrl;
      else {
        const opt = Array.from(sel.options).find((o) => o.textContent.startsWith(assetFromUrl));
        if (opt) sel.value = opt.value;
      }
      loadLatestForAsset(Number(sel.value || 0));
    }

    loadSavedList();
  }

  function bindUndercarriageEvents() {
    document.getElementById("ucAsset")?.addEventListener("change", (evt) => {
      loadLatestForAsset(Number(evt.target?.value || 0));
    });
    document.getElementById("ucPullSmuBtn")?.addEventListener("click", () => pullLiveSmu());
    document.getElementById("ucSaveBtn")?.addEventListener("click", () => saveInspection());
    document.getElementById("ucLoadSavedBtn")?.addEventListener("click", () => loadSavedList());
    document.getElementById("ucReportXlsxBtn")?.addEventListener("click", () => exportXlsx(0).catch((e) => alert(e.message)));
    document.getElementById("ucSavedList")?.addEventListener("click", (evt) => {
      const btn = evt.target?.closest?.("button[data-uc-xlsx]");
      if (!btn) return;
      exportXlsx(Number(btn.getAttribute("data-uc-xlsx") || 0)).catch((e) => alert(e.message));
    });
  }

  window.initUndercarriage = initUndercarriage;
  window.bindUndercarriageEvents = bindUndercarriageEvents;
})();
