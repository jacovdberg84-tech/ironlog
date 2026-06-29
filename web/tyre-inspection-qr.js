(function () {
  const API = "/api";
  const TYRE_POSITIONS = [
    { key: "front_left", label: "Front Left", surveyCode: "LF" },
    { key: "front_right", label: "Front Right", surveyCode: "RF" },
    { key: "rear_right_inner", label: "Rear Right Inner", surveyCode: "RM" },
    { key: "rear_right_outer", label: "Rear Right Outer", surveyCode: "RR" },
    { key: "rear_left_outer", label: "Rear Left Outer", surveyCode: "LR" },
    { key: "rear_left_inner", label: "Rear Left Inner", surveyCode: "LM" },
  ];

  let mobileCtx = {
    assetId: 0,
    assetCode: "",
    thresholds: { warn_tread_mm: 8, min_tread_mm: 3 },
    lifecycleByKey: new Map(),
  };

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
      h["x-user-name"] = localStorage.getItem("ironlog_session_user") || "field";
      h["x-user-role"] = localStorage.getItem("ironlog_session_role") || "operator";
      h["x-user-roles"] = localStorage.getItem("ironlog_session_roles") || "operator";
      const tok = localStorage.getItem("ironlog_auth_token");
      if (tok) h.Authorization = `Bearer ${tok}`;
    } catch {}
    return h;
  }

  function tiInputId(key, field) {
    return `ti_${key}_${field}`;
  }

  function readOptionalNumber(id) {
    const raw = String(document.getElementById(id)?.value || "").trim();
    if (raw === "") return null;
    const n = Number(raw);
    return Number.isFinite(n) ? n : null;
  }

  function treadStatus(rtd) {
    const t = rtd == null ? null : Number(rtd);
    if (t == null || !Number.isFinite(t)) return { cls: "", label: "" };
    const warn = Number(mobileCtx.thresholds.warn_tread_mm || 8);
    const min = Number(mobileCtx.thresholds.min_tread_mm || 3);
    if (t <= min) return { cls: "bad", label: "Replace" };
    if (t <= warn) return { cls: "warn", label: "Monitor" };
    return { cls: "ok", label: "OK" };
  }

  function updatePositionStatus(key) {
    const rtd = readOptionalNumber(tiInputId(key, "rtd_outer"));
    const st = treadStatus(rtd);
    const el = document.getElementById(`ti_status_${key}`);
    if (!el) return;
    if (!st.label) {
      el.textContent = "";
      el.className = "ti-status";
      return;
    }
    el.textContent = st.label;
    el.className = `ti-status ${st.cls}`;
  }

  function renderPositionCards() {
    const mount = document.getElementById("tiPositions");
    if (!mount) return;
    mount.innerHTML = TYRE_POSITIONS.map((p) => `
      <details class="ti-pos" open>
        <summary>
          <span>${esc(p.label)}</span>
          <span style="display:flex;align-items:center;gap:6px;">
            <span class="ti-pos-code">${esc(p.surveyCode)}</span>
            <span class="ti-status" id="ti_status_${esc(p.key)}"></span>
          </span>
        </summary>
        <div class="ti-pos-body">
          <div class="ti-pos-grid">
            <label class="span-2">Serial
              <input id="${tiInputId(p.key, "serial")}" type="text" placeholder="Tyre serial" />
            </label>
            <label>Pressure cold (kPa)
              <input id="${tiInputId(p.key, "pressure")}" type="number" min="0" step="1" inputmode="decimal" />
            </label>
            <label>Pressure recom.
              <input id="${tiInputId(p.key, "pressure_recom")}" type="number" min="0" step="1" inputmode="decimal" />
            </label>
            <label>RTD outer (mm)
              <input id="${tiInputId(p.key, "rtd_outer")}" type="number" min="0" step="0.1" inputmode="decimal" data-ti-rtd="${esc(p.key)}" />
            </label>
            <label>RTD inner (mm)
              <input id="${tiInputId(p.key, "rtd_inner")}" type="number" min="0" step="0.1" inputmode="decimal" />
            </label>
            <label>OTD (mm)
              <input id="${tiInputId(p.key, "otd")}" type="number" min="0" step="0.1" inputmode="decimal" />
            </label>
            <label>Purchase price
              <input id="${tiInputId(p.key, "cost")}" type="number" min="0" step="0.01" inputmode="decimal" />
            </label>
            <label class="span-2">Make / brand
              <input id="${tiInputId(p.key, "make")}" type="text" placeholder="e.g. Michelin 23.5R25" />
            </label>
            <label class="span-2">Last changed
              <input id="${tiInputId(p.key, "changed")}" type="date" />
            </label>
          </div>
        </div>
      </details>
    `).join("");

    mount.querySelectorAll("[data-ti-rtd]").forEach((el) => {
      el.addEventListener("input", () => updatePositionStatus(el.getAttribute("data-ti-rtd") || ""));
    });
  }

  function prefillFromLifecycle() {
    TYRE_POSITIONS.forEach((p) => {
      const row = mobileCtx.lifecycleByKey.get(String(p.key).toLowerCase());
      if (!row) return;
      const set = (field, val) => {
        const el = document.getElementById(tiInputId(p.key, field));
        if (!el || el.value) return;
        if (val != null && val !== "") el.value = val;
      };
      set("serial", row.serial_number);
      set("changed", row.install_date || row.last_changed_date);
      set("cost", row.tyre_cost > 0 ? Number(row.tyre_cost).toFixed(2) : "");
      set("rtd_outer", row.tread_depth ?? row.rtd_outer);
      set("rtd_inner", row.rtd_inner);
      set("pressure", row.pressure);
      set("pressure_recom", row.pressure_recommended);
      set("otd", row.original_tread_depth);
      set("make", [row.tyre_make, row.brand_number].filter(Boolean).join(" "));
      updatePositionStatus(p.key);
    });
  }

  function collectTyreRows() {
    return TYRE_POSITIONS.map((p) => {
      const tyre_cost = Number(document.getElementById(tiInputId(p.key, "cost"))?.value || 0) || 0;
      const rtd_outer = readOptionalNumber(tiInputId(p.key, "rtd_outer"));
      const rtd_inner = readOptionalNumber(tiInputId(p.key, "rtd_inner"));
      const makeBrand = String(document.getElementById(tiInputId(p.key, "make"))?.value || "").trim();
      const parts = makeBrand.split(/\s+/).filter(Boolean);
      return {
        position_key: p.key,
        position_label: p.label,
        survey_code: p.surveyCode,
        tyre_make: parts[0] || "",
        brand_number: parts.slice(1).join(" ") || "",
        tyre_description: "",
        pressure: readOptionalNumber(tiInputId(p.key, "pressure")),
        pressure_recommended: readOptionalNumber(tiInputId(p.key, "pressure_recom")),
        pressure_hot: null,
        original_tread_depth: readOptionalNumber(tiInputId(p.key, "otd")),
        rtd_outer,
        rtd_inner,
        tread_depth: rtd_outer,
        serial_number: String(document.getElementById(tiInputId(p.key, "serial"))?.value || "").trim(),
        last_changed_date: String(document.getElementById(tiInputId(p.key, "changed"))?.value || "").trim(),
        tyre_cost,
      };
    });
  }

  async function fetchJson(url, options) {
    const res = await fetch(url, options);
    const text = await res.text();
    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch { data = null; }
    if (!res.ok) throw new Error(data?.error || text || `Request failed (${res.status})`);
    return data || {};
  }

  async function loadAssetContext(assetCode) {
    const code = String(assetCode || "").trim();
    if (!code) throw new Error("Missing asset_code in QR link.");
    const data = await fetchJson(`${API}/assets/${encodeURIComponent(code)}/tyre-qr-profile`);
    const payload = data?.live_preview || data?.stored?.qr_payload || {};
    const asset = payload?.asset || {};
    const assetId = Number(asset.id || 0);
    if (!assetId) throw new Error(`Asset ${code} not found.`);

    mobileCtx.assetId = assetId;
    mobileCtx.assetCode = String(asset.asset_code || code);

    const setText = (id, v) => {
      const el = document.getElementById(id);
      if (el) el.textContent = v;
    };
    setText("tiMachineCode", mobileCtx.assetCode);
    setText("tiMachineName", String(asset.asset_name || "—"));
    const hrs = payload?.meter?.current_hours;
    setText("tiMeter", hrs != null ? `${Number(hrs).toFixed(1)} h` : "—");

    const locked = document.getElementById("tiLockedNote");
    if (locked) locked.style.display = "block";

    const dateEl = document.getElementById("tiDate");
    if (dateEl && !dateEl.value) dateEl.value = new Date().toISOString().slice(0, 10);

    const hoursEl = document.getElementById("tiHours");
    if (hoursEl && !hoursEl.value && hrs != null) {
      hoursEl.value = Number(hrs).toFixed(1);
    }

    try {
      const insp = localStorage.getItem("ironlog_session_user");
      const inspector = document.getElementById("tiInspector");
      if (inspector && !inspector.value && insp) inspector.value = insp;
    } catch {}

    try {
      const life = await fetchJson(`${API}/maintenance/tyre-inspections/lifecycle?asset_id=${assetId}`, {
        headers: authHeaders(),
      });
      mobileCtx.thresholds = life?.thresholds || mobileCtx.thresholds;
      mobileCtx.lifecycleByKey = new Map();
      for (const row of life?.positions || []) {
        const k = String(row.position_key || "").toLowerCase();
        if (k) mobileCtx.lifecycleByKey.set(k, row);
      }
      prefillFromLifecycle();
    } catch {
      prefillFromLifecycle();
    }
  }

  async function pullLiveHours() {
    const meta = document.getElementById("tiHoursMeta");
    if (!mobileCtx.assetId) {
      if (meta) meta.textContent = "Machine not loaded.";
      return;
    }
    if (meta) meta.textContent = "Loading live hours…";
    try {
      const data = await fetchJson(`${API}/maintenance/asset/${mobileCtx.assetId}/live-hours`, {
        headers: authHeaders(),
      });
      const current = Number(data?.current_hours ?? 0);
      const inp = document.getElementById("tiHours");
      if (inp) inp.value = Number.isFinite(current) ? current.toFixed(1) : "";
      if (meta) meta.textContent = `Live hours: ${Number.isFinite(current) ? current.toFixed(1) : "0.0"} (${data?.source || "-"})`;
    } catch (e) {
      if (meta) meta.textContent = `Live hours error: ${e.message || e}`;
    }
  }

  async function saveTyreInspectionMobile() {
    const msg = document.getElementById("tiMsg");
    const btn = document.getElementById("tiSaveBtn");
    const asset_id = mobileCtx.assetId;
    const inspection_date = String(document.getElementById("tiDate")?.value || "").trim();
    const inspector_name = String(document.getElementById("tiInspector")?.value || "").trim();
    const notes = String(document.getElementById("tiNotes")?.value || "").trim();
    const runningRaw = String(document.getElementById("tiHours")?.value || "").trim();
    const running_hours = runningRaw === "" ? 0 : Number(runningRaw);

    if (!asset_id) {
      if (msg) { msg.className = "ti-msg err"; msg.textContent = "Open this form from a machine QR code."; }
      return;
    }
    if (!inspection_date) {
      if (msg) { msg.className = "ti-msg err"; msg.textContent = "Select inspection date."; }
      return;
    }
    if (!Number.isFinite(running_hours) || running_hours < 0) {
      if (msg) { msg.className = "ti-msg err"; msg.textContent = "Enter valid machine hours."; }
      return;
    }

    const tyres = collectTyreRows();
    const hasReading = tyres.some((t) =>
      t.serial_number
      || t.rtd_outer != null
      || t.rtd_inner != null
      || t.pressure != null
      || Number(t.tyre_cost || 0) > 0,
    );
    if (!hasReading) {
      if (msg) { msg.className = "ti-msg err"; msg.textContent = "Enter at least one tyre reading (serial, tread, or pressure)."; }
      return;
    }

    if (msg) { msg.className = "ti-msg"; msg.textContent = "Uploading to IRONLOG…"; }
    if (btn) btn.disabled = true;
    try {
      const data = await fetchJson(`${API}/maintenance/tyre-inspections`, {
        method: "POST",
        headers: authHeaders({ "Content-Type": "application/json" }),
        body: JSON.stringify({
          asset_id,
          inspection_date,
          inspector_name: inspector_name || null,
          running_hours: Number(running_hours.toFixed(1)),
          notes: notes || null,
          tyres,
        }),
      });
      const alerts = Array.isArray(data.alerts) ? data.alerts.length : 0;
      if (msg) {
        msg.className = "ti-msg ok";
        msg.textContent = `Saved to IRONLOG for ${mobileCtx.assetCode}.${alerts ? ` ${alerts} tyre(s) flagged for attention.` : ""}`;
      }
      try {
        if (inspector_name) localStorage.setItem("ironlog_session_user", inspector_name);
      } catch {}
    } catch (e) {
      if (msg) {
        msg.className = "ti-msg err";
        msg.textContent = e.message || String(e);
      }
    } finally {
      if (btn) btn.disabled = false;
    }
  }

  async function initTyreInspectionMobile() {
    renderPositionCards();
    const params = new URLSearchParams(location.search);
    const assetCode = params.get("asset_code") || "";
    await loadAssetContext(assetCode);
    document.getElementById("tiPullHoursBtn")?.addEventListener("click", () => pullLiveHours().catch(() => {}));
    document.getElementById("tiSaveBtn")?.addEventListener("click", () => saveTyreInspectionMobile());
  }

  function assetCodeFromSelect(selectEl) {
    if (!selectEl) return "";
    const opt = selectEl.options[selectEl.selectedIndex];
    if (!opt) return "";
    return String(opt.textContent || "").split(" - ")[0].trim();
  }

  async function buildTyreQrImageData(assetCode) {
    const code = String(assetCode || "").trim();
    if (!code) throw new Error("Select a machine first.");
    const res = await fetch(`${API}/assets/${encodeURIComponent(code)}/tyre-qr-profile/refresh`, {
      method: "POST",
      headers: { ...authHeaders(), "Content-Type": "application/json" },
      body: JSON.stringify({}),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "QR generation failed");
    const scanValue = String(data?.qr_payload?.scan_url || "").trim();
    if (!scanValue) throw new Error("No scan URL generated.");
    const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=420x420&data=${encodeURIComponent(scanValue)}`;
    return { qrUrl, qrText: String(data.qr_text || ""), scanValue, payload: data.qr_payload || {} };
  }

  function openTyreQrLabelSheet(labels) {
    const safeLabels = Array.isArray(labels) ? labels.filter((l) => l?.qrUrl && l?.code) : [];
    if (!safeLabels.length) throw new Error("No QR labels to print.");
    const win = window.open("", "_blank", "width=1100,height=800");
    if (!win) {
      alert("Pop-up blocked. Allow pop-ups and try again.");
      return;
    }
    const cells = safeLabels.map((l) => `
      <div class="cell">
        <img src="${l.qrUrl}" alt="${esc(l.code)} QR" />
        <div class="code">${esc(l.code)}</div>
        <div class="sub">Tyre inspection</div>
      </div>
    `).join("");
    win.document.write(`<!doctype html><html><head><meta charset="utf-8" /><title>IRONLOG Tyre QR Labels</title>
      <style>
        @page { size: A4 portrait; margin: 8mm; }
        body { margin: 0; font-family: Arial, sans-serif; color: #111; }
        .sheet { padding: 8mm; }
        .head { margin-bottom: 6mm; font-size: 12px; }
        .grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 4mm; }
        .cell { border: 1px solid #bbb; border-radius: 4px; padding: 3mm 2mm; text-align: center; min-height: 48mm; break-inside: avoid; }
        .cell img { width: 32mm; height: 32mm; display: block; margin: 0 auto 2mm; }
        .code { font-size: 11px; font-weight: 700; }
        .sub { font-size: 9px; color: #475569; margin-top: 2px; }
      </style></head><body>
      <div class="sheet">
        <div class="head">IRONLOG Tyre Inspection QR Labels | ${safeLabels.length} label(s) | Stick near wheel arch / tyre bay</div>
        <div class="grid">${cells}</div>
      </div>
      <script>window.onload = () => { window.focus(); window.print(); };</script>
      </body></html>`);
    win.document.close();
  }

  async function previewTyreQr() {
    const code = assetCodeFromSelect(document.getElementById("tiQrAsset"))
      || assetCodeFromSelect(document.getElementById("tyreAsset"));
    const msg = document.getElementById("tiQrMsg");
    const img = document.getElementById("tiQrPreview");
    const meta = document.getElementById("tiQrMeta");
    if (!code) {
      if (msg) msg.textContent = "Select a machine for the QR label.";
      return;
    }
    if (msg) msg.textContent = "Generating QR…";
    try {
      const { qrUrl, scanValue, payload } = await buildTyreQrImageData(code);
      if (img) {
        img.src = qrUrl;
        img.style.display = "block";
      }
      if (meta) {
        meta.innerHTML = `
          <div><b>${esc(code)}</b> — opens tyre inspection for this machine only</div>
          <div class="muted mini" style="margin-top:4px; word-break:break-all;">${esc(scanValue)}</div>
          <div class="muted mini" style="margin-top:4px;">Meter ${payload?.meter?.current_hours ?? "—"}h · Last inspection ${esc(payload?.tyres?.last_inspection_date || "none")}</div>
        `;
      }
      if (msg) msg.textContent = "QR ready — print or download and stick near the tyres.";
    } catch (e) {
      if (msg) msg.textContent = `QR failed: ${e.message || e}`;
    }
  }

  async function downloadTyreQrPng() {
    const code = assetCodeFromSelect(document.getElementById("tiQrAsset"))
      || assetCodeFromSelect(document.getElementById("tyreAsset"));
    if (!code) return alert("Select a machine first.");
    const { qrUrl } = await buildTyreQrImageData(code);
    const response = await fetch(qrUrl);
    if (!response.ok) throw new Error("QR image fetch failed");
    const blob = await response.blob();
    const objUrl = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = objUrl;
    a.download = `${code}_tyre_inspection_qr.png`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(objUrl), 10000);
  }

  async function printTyreQrLabel() {
    const code = assetCodeFromSelect(document.getElementById("tiQrAsset"))
      || assetCodeFromSelect(document.getElementById("tyreAsset"));
    if (!code) return alert("Select a machine first.");
    const { qrUrl } = await buildTyreQrImageData(code);
    openTyreQrLabelSheet([{ code, qrUrl }]);
  }

  async function printAllTyreQrLabels() {
    const sel = document.getElementById("tiQrAsset") || document.getElementById("tyreAsset");
    if (!sel) throw new Error("Asset list not loaded.");
    const codes = Array.from(sel.options)
      .map((o) => String(o.textContent || "").split(" - ")[0].trim())
      .filter((c) => c && c !== "Select asset" && c !== "Select machine");
    if (!codes.length) throw new Error("No machines in list.");
    const msg = document.getElementById("tiQrMsg");
    if (msg) msg.textContent = `Generating ${codes.length} QR labels…`;
    const labels = [];
    for (const code of codes) {
      try {
        const { qrUrl } = await buildTyreQrImageData(code);
        labels.push({ code, qrUrl });
      } catch {}
    }
    openTyreQrLabelSheet(labels);
    if (msg) msg.textContent = `Printed sheet with ${labels.length} label(s).`;
  }

  function bindTyreQrAdmin() {
    document.getElementById("tiQrPreviewBtn")?.addEventListener("click", () => previewTyreQr().catch((e) => alert(e.message)));
    document.getElementById("tiQrDownloadBtn")?.addEventListener("click", () => downloadTyreQrPng().catch((e) => alert(e.message)));
    document.getElementById("tiQrPrintBtn")?.addEventListener("click", () => printTyreQrLabel().catch((e) => alert(e.message)));
    document.getElementById("tiQrPrintAllBtn")?.addEventListener("click", () => printAllTyreQrLabels().catch((e) => alert(e.message)));
    const demo = document.getElementById("tiQrOpenMobileDemo");
    const sel = document.getElementById("tiQrAsset");
    const syncDemoHref = () => {
      if (!demo) return;
      const code = assetCodeFromSelect(sel) || assetCodeFromSelect(document.getElementById("tyreAsset"));
      demo.href = code
        ? `./tyre-inspection-mobile.html?asset_code=${encodeURIComponent(code)}`
        : "./tyre-inspection-mobile.html";
    };
    sel?.addEventListener("change", syncDemoHref);
    syncDemoHref();
  }

  function syncTyreQrAssetSelect() {
    const src = document.getElementById("tyreAsset");
    const dst = document.getElementById("tiQrAsset");
    if (!src || !dst) return;
    dst.innerHTML = src.innerHTML;
  }

  window.initTyreInspectionMobile = initTyreInspectionMobile;
  window.bindTyreQrAdmin = bindTyreQrAdmin;
  window.syncTyreQrAssetSelect = syncTyreQrAssetSelect;
})();
