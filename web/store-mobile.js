(function () {
  const API = "/api";
  const LABEL_QUEUE_KEY = "ironlog_store_label_queue";

  let state = {
    site: "main",
    stockRows: [],
    pendingOrders: [],
    locations: [],
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
      h["x-site-code"] = localStorage.getItem("ironlog_session_site") || state.site || "main";
      h["x-user-name"] = localStorage.getItem("ironlog_session_user") || "storeman";
      h["x-user-role"] = localStorage.getItem("ironlog_session_role") || "storeman";
      h["x-user-roles"] = localStorage.getItem("ironlog_session_roles") || "storeman";
      const tok = localStorage.getItem("ironlog_auth_token");
      if (tok) h.Authorization = `Bearer ${tok}`;
    } catch {}
    return h;
  }

  async function fetchJson(url, options) {
    const res = await fetch(url, options);
    const text = await res.text();
    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch { data = null; }
    if (!res.ok) throw new Error(data?.error || text || `Request failed (${res.status})`);
    return data || {};
  }

  function readLabelQueue() {
    try {
      const raw = localStorage.getItem(LABEL_QUEUE_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr : [];
    } catch {
      return [];
    }
  }

  function writeLabelQueue(items) {
    localStorage.setItem(LABEL_QUEUE_KEY, JSON.stringify(items));
    updateLabelCount();
    renderLabelQueue();
  }

  function addToLabelQueue(item) {
    const q = readLabelQueue();
    q.unshift({
      id: `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      part_code: String(item.part_code || "").trim(),
      part_name: String(item.part_name || "").trim(),
      qty: Number(item.qty || 0),
      location_code: String(item.location_code || "").trim(),
      bin_code: String(item.bin_code || "").trim(),
      received_at: new Date().toISOString(),
      reference: String(item.reference || "").trim(),
    });
    writeLabelQueue(q.slice(0, 50));
  }

  function updateLabelCount() {
    const n = readLabelQueue().length;
    const el = document.getElementById("stLabelCount");
    if (el) el.textContent = n ? `(${n})` : "";
  }

  function setActiveTab(tab) {
    const id = String(tab || "stock");
    document.querySelectorAll("[data-st-tab]").forEach((btn) => {
      btn.classList.toggle("is-active", btn.getAttribute("data-st-tab") === id);
    });
    document.querySelectorAll(".st-panel").forEach((p) => p.classList.remove("is-active"));
    const panel = document.getElementById(`stPanel${id.charAt(0).toUpperCase()}${id.slice(1)}`);
    if (panel) panel.classList.add("is-active");
    if (id === "receive") loadPendingOrders().catch(() => {});
    if (id === "labels") renderLabelQueue();
  }

  function filterStockRows() {
    const q = String(document.getElementById("stStockFilter")?.value || "").trim().toLowerCase();
    const onlyLow = Boolean(document.getElementById("stOnlyLow")?.checked);
    return state.stockRows.filter((r) => {
      if (onlyLow && !r.below_min) return false;
      if (!q) return true;
      return String(r.part_code || "").toLowerCase().includes(q)
        || String(r.part_name || "").toLowerCase().includes(q);
    });
  }

  function renderStockTable() {
    const body = document.getElementById("stStockBody");
    const msg = document.getElementById("stStockMsg");
    if (!body) return;
    const rows = filterStockRows();
    if (!rows.length) {
      body.innerHTML = `<tr><td colspan="5" class="muted">No parts match filter.</td></tr>`;
      if (msg) msg.textContent = `${state.stockRows.length} parts in catalogue.`;
      return;
    }
    body.innerHTML = rows.slice(0, 500).map((r) => {
      const low = Boolean(r.below_min);
      const badge = low
        ? `<span class="st-badge low">Low</span>`
        : `<span class="st-badge ok">OK</span>`;
      return `<tr class="${low ? "low" : ""}">
        <td><strong>${esc(r.part_code)}</strong></td>
        <td>${esc(r.part_name || "—")}</td>
        <td>${Number(r.on_hand || 0).toFixed(1)}</td>
        <td>${Number(r.min_stock || 0).toFixed(1)}</td>
        <td>${badge}</td>
      </tr>`;
    }).join("");
    if (msg) {
      msg.textContent = `Showing ${Math.min(rows.length, 500)} of ${rows.length} (catalogue ${state.stockRows.length}).`;
    }
  }

  function updateKpis() {
    const parts = state.stockRows.length;
    const low = state.stockRows.filter((r) => r.below_min).length;
    const pending = state.pendingOrders.length;
    const set = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
    set("stKpiParts", parts);
    set("stKpiLow", low);
    set("stKpiPending", pending);
  }

  async function loadStock() {
    const msg = document.getElementById("stStockMsg");
    if (msg) msg.textContent = "Loading stock…";
    const data = await fetchJson(`${API}/stock/monitor`, { headers: authHeaders() });
    state.stockRows = Array.isArray(data.rows) ? data.rows : [];
    renderStockTable();
    updateKpis();
  }

  async function loadPendingOrders() {
    const data = await fetchJson(`${API}/stock/part-orders?status=`, { headers: authHeaders() });
    const rows = Array.isArray(data.rows) ? data.rows : [];
    state.pendingOrders = rows.filter((r) => {
      const st = String(r.status || "").toLowerCase();
      return (st === "on_order" || st === "in_transit") && !r.in_store_inventory;
    });
    renderPendingOrders();
    updateKpis();
  }

  function renderPendingOrders() {
    const host = document.getElementById("stPendingList");
    if (!host) return;
    if (!state.pendingOrders.length) {
      host.innerHTML = `<div class="st-card"><div class="meta">No purchase lines waiting to arrive.</div></div>`;
      return;
    }
    host.innerHTML = state.pendingOrders.map((r) => {
      const id = Number(r.id || 0);
      const st = String(r.status || "on_order").replace(/_/g, " ");
      const hasCode = Boolean(String(r.part_code || "").trim());
      return `<div class="st-card" data-pending-id="${id}">
        <h3>${esc(r.part_code || r.part_name || `Line #${id}`)}</h3>
        <div class="meta">${esc(r.part_name || "")}</div>
        <div class="meta" style="margin-top:4px;">
          Qty <b>${Number(r.qty || 0)}</b>
          · ${esc(st)}
          ${r.po_number ? ` · PO ${esc(r.po_number)}` : ""}
          ${r.requisition_number ? ` · Req ${esc(r.requisition_number)}` : ""}
        </div>
        ${!hasCode ? `<div class="meta" style="color:var(--warn);margin-top:6px;">Add a part code on this line in IRONLOG before receiving.</div>` : ""}
        <div class="st-btn-row">
          <button type="button" class="st-btn primary" data-receive-order="${id}" ${!hasCode ? "disabled" : ""}>Receive to stock</button>
        </div>
      </div>`;
    }).join("");

    host.querySelectorAll("[data-receive-order]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const id = Number(btn.getAttribute("data-receive-order") || 0);
        receivePurchaseLine(id).catch((e) => showReceiveMsg(e.message, true));
      });
    });
  }

  function showReceiveMsg(text, isErr) {
    const el = document.getElementById("stReceiveMsg");
    if (!el) return;
    el.className = `st-msg ${isErr ? "err" : "ok"}`;
    el.textContent = text;
  }

  async function receivePurchaseLine(orderId) {
    const id = Number(orderId || 0);
    if (!id) return;
    showReceiveMsg("Receiving…", false);
    const data = await fetchJson(`${API}/stock/part-orders/${id}/receive`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({}),
    });
    const receipt = data.stock_receipt || {};
    if (receipt.received) {
      addToLabelQueue({
        part_code: receipt.part_code,
        part_name: state.pendingOrders.find((r) => Number(r.id) === id)?.part_name || "",
        qty: receipt.qty,
        reference: `PO line #${id}`,
      });
      showReceiveMsg(
        `Received ${receipt.qty} × ${receipt.part_code} (${receipt.on_hand_after} on hand). Added to label queue.`,
        false,
      );
    } else if (receipt.already) {
      showReceiveMsg("Already in store inventory.", false);
    } else {
      throw new Error(receipt.error || data.error || "Receive failed");
    }
    await Promise.all([loadPendingOrders(), loadStock()]);
  }

  async function quickReceive() {
    const part_code = String(document.getElementById("stRcvCode")?.value || "").trim();
    const part_name = String(document.getElementById("stRcvName")?.value || "").trim();
    const qty = Number(document.getElementById("stRcvQty")?.value || 0);
    const location_code = String(document.getElementById("stRcvLocation")?.value || "").trim();
    const bin_code = String(document.getElementById("stRcvBin")?.value || "").trim();
    const unit_cost = document.getElementById("stRcvCost")?.value;
    const cost_currency = String(document.getElementById("stRcvCurrency")?.value || "USD");
    const reference = String(document.getElementById("stRcvRef")?.value || "").trim() || "store_mobile_receive";
    const received_by = String(document.getElementById("stRcvBy")?.value || "").trim();

    if (!part_code) throw new Error("Part code is required.");
    if (!Number.isFinite(qty) || qty <= 0) throw new Error("Qty must be greater than zero.");

    const existing = state.stockRows.find((r) => String(r.part_code) === part_code);
    const body = {
      part_code,
      quantity: qty,
      movement_type: "in",
      reference,
      location_code: location_code || undefined,
      bin_code: bin_code || undefined,
      create_if_missing: !existing,
      part_name: existing ? undefined : (part_name || part_code),
    };
    if (unit_cost !== "" && unit_cost != null) {
      body.unit_cost = Number(unit_cost);
      body.cost_currency = cost_currency;
    }

    showReceiveMsg("Posting stock IN…", false);
    const data = await fetchJson(`${API}/stock/movement`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify(body),
    });

    try {
      if (received_by) localStorage.setItem("ironlog_session_user", received_by);
    } catch {}

    addToLabelQueue({
      part_code,
      part_name: part_name || existing?.part_name || part_code,
      qty,
      location_code,
      bin_code,
      reference,
    });

    showReceiveMsg(
      `Received ${qty} × ${part_code}. On hand now ${Number(data.on_hand_after ?? 0).toFixed(1)}. Added to label queue.`,
      false,
    );

    document.getElementById("stRcvCode").value = "";
    document.getElementById("stRcvQty").value = "1";
    document.getElementById("stRcvName").value = "";
    document.getElementById("stRcvRef").value = "";

    await loadStock();
  }

  function renderLabelQueue() {
    const host = document.getElementById("stLabelQueue");
    if (!host) return;
    const q = readLabelQueue();
    updateLabelCount();
    if (!q.length) {
      host.innerHTML = `<div class="st-card"><div class="meta">No labels queued. Receive parts first.</div></div>`;
      return;
    }
    host.innerHTML = q.map((item) => `
      <div class="st-label-item">
        <div>
          <strong>${esc(item.part_code)}</strong>
          <div class="meta">${esc(item.part_name || "")}</div>
          <div class="meta">Qty ${Number(item.qty || 0)} · ${esc(item.location_code || "—")} ${item.bin_code ? `/ ${esc(item.bin_code)}` : ""}</div>
        </div>
        <button type="button" class="st-btn" style="width:auto;padding:6px 10px;" data-remove-label="${esc(item.id)}">×</button>
      </div>
    `).join("");

    host.querySelectorAll("[data-remove-label]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const id = btn.getAttribute("data-remove-label");
        writeLabelQueue(readLabelQueue().filter((x) => x.id !== id));
      });
    });
  }

  function partLabelQrUrl(partCode) {
    const origin = window.location.origin;
    const site = encodeURIComponent(state.site || "main");
    const code = encodeURIComponent(partCode);
    const scan = `${origin}/web/store-mobile.html?site=${site}&part_code=${code}`;
    return `https://api.qrserver.com/v1/create-qr-code/?size=120x120&data=${encodeURIComponent(scan)}`;
  }

  function printLabelSheet(items) {
    const labels = Array.isArray(items) ? items.filter((x) => x.part_code) : [];
    if (!labels.length) throw new Error("No labels to print.");
    const win = window.open("", "_blank", "width=1100,height=800");
    if (!win) {
      alert("Pop-up blocked. Allow pop-ups and try again.");
      return;
    }
    const cells = labels.map((l) => `
      <div class="cell">
        <img src="${partLabelQrUrl(l.part_code)}" alt="QR" />
        <div class="code">${esc(l.part_code)}</div>
        <div class="desc">${esc(String(l.part_name || "").slice(0, 42))}</div>
        <div class="qty">Qty: ${Number(l.qty || 0)}</div>
        <div class="loc">${esc([l.location_code, l.bin_code].filter(Boolean).join(" / ") || "STORE")}</div>
        <div class="date">${esc(String(l.received_at || "").slice(0, 10))}</div>
      </div>
    `).join("");
    win.document.write(`<!doctype html><html><head><meta charset="utf-8" /><title>Store shelf labels</title>
      <style>
        @page { size: A4 portrait; margin: 8mm; }
        body { margin: 0; font-family: Arial, sans-serif; color: #111; }
        .sheet { padding: 8mm; }
        .head { margin-bottom: 6mm; font-size: 12px; }
        .grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 4mm; }
        .cell { border: 1px solid #bbb; border-radius: 4px; padding: 3mm; text-align: center; min-height: 52mm; break-inside: avoid; }
        .cell img { width: 28mm; height: 28mm; display: block; margin: 0 auto 2mm; }
        .code { font-size: 13px; font-weight: 800; letter-spacing: 0.02em; }
        .desc { font-size: 9px; color: #334155; margin-top: 2px; min-height: 2.4em; }
        .qty, .loc, .date { font-size: 9px; color: #475569; margin-top: 2px; }
      </style></head><body>
      <div class="sheet">
        <div class="head">IRONLOG Store shelf labels · ${labels.length} label(s) · ${esc(state.site)}</div>
        <div class="grid">${cells}</div>
      </div>
      <script>window.onload = () => { window.focus(); window.print(); };</script>
      </body></html>`);
    win.document.close();
  }

  async function loadLocations() {
    try {
      const data = await fetchJson(`${API}/stock/locations?active=1`, { headers: authHeaders() });
      state.locations = Array.isArray(data.rows) ? data.rows : [];
      const sel = document.getElementById("stRcvLocation");
      if (sel) {
        sel.innerHTML = `<option value="">— optional —</option>` + state.locations
          .map((l) => `<option value="${esc(l.location_code)}">${esc(l.location_code)}${l.location_name ? ` — ${esc(l.location_name)}` : ""}</option>`)
          .join("");
      }
    } catch {}
  }

  async function loadStoreProfile() {
    const data = await fetchJson(`${API}/stock/store-qr-profile?site=${encodeURIComponent(state.site)}`, {
      headers: authHeaders(),
    });
    const p = data?.live_preview || data?.stored?.qr_payload || {};
    state.site = String(p.site_code || state.site || "main");
    const siteEl = document.getElementById("stSiteCode");
    if (siteEl) siteEl.textContent = state.site;
    try {
      localStorage.setItem("ironlog_session_site", state.site);
    } catch {}
    const set = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
    set("stKpiParts", p.inventory?.part_count ?? "—");
    set("stKpiLow", p.inventory?.below_min ?? "—");
    set("stKpiPending", p.pending_arrivals ?? "—");
  }

  async function initStoreMobile() {
    const params = new URLSearchParams(location.search);
    state.site = String(params.get("site") || "main").trim().toLowerCase() || "main";
    const partCodeFocus = String(params.get("part_code") || "").trim();

    document.querySelectorAll("[data-st-tab]").forEach((btn) => {
      btn.addEventListener("click", () => setActiveTab(btn.getAttribute("data-st-tab")));
    });
    document.getElementById("stStockFilter")?.addEventListener("input", renderStockTable);
    document.getElementById("stOnlyLow")?.addEventListener("change", renderStockTable);
    document.getElementById("stQuickReceiveBtn")?.addEventListener("click", () => quickReceive().catch((e) => showReceiveMsg(e.message, true)));
    document.getElementById("stClearLabelsBtn")?.addEventListener("click", () => {
      writeLabelQueue([]);
      const el = document.getElementById("stLabelMsg");
      if (el) { el.className = "st-msg"; el.textContent = "Label queue cleared."; }
    });
    document.getElementById("stPrintLabelsBtn")?.addEventListener("click", () => {
      try {
        const q = readLabelQueue();
        printLabelSheet(q);
        const el = document.getElementById("stLabelMsg");
        if (el) { el.className = "st-msg ok"; el.textContent = `Printing ${q.length} label(s)…`; }
      } catch (e) {
        const el = document.getElementById("stLabelMsg");
        if (el) { el.className = "st-msg err"; el.textContent = e.message || String(e); }
      }
    });

    try {
      const user = localStorage.getItem("ironlog_session_user");
      const by = document.getElementById("stRcvBy");
      if (by && user && !by.value) by.value = user;
    } catch {}

    await loadStoreProfile();
    await Promise.all([loadStock(), loadLocations(), loadPendingOrders()]);
    updateLabelCount();
    renderLabelQueue();

    if (partCodeFocus) {
      const filter = document.getElementById("stStockFilter");
      if (filter) filter.value = partCodeFocus;
      setActiveTab("stock");
      renderStockTable();
    }
  }

  async function buildStoreQrImageData(siteCode) {
    const site = String(siteCode || "main").trim().toLowerCase() || "main";
    const res = await fetch(`${API}/stock/store-qr-profile/refresh`, {
      method: "POST",
      headers: { ...authHeaders(), "Content-Type": "application/json" },
      body: JSON.stringify({ site }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "QR generation failed");
    const scanValue = String(data?.qr_payload?.scan_url || "").trim();
    if (!scanValue) throw new Error("No scan URL generated.");
    const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=420x420&data=${encodeURIComponent(scanValue)}`;
    return { qrUrl, qrText: String(data.qr_text || ""), scanValue, payload: data.qr_payload || {} };
  }

  async function previewStoreQr() {
    const site = String(document.getElementById("storeQrSite")?.value || "main").trim().toLowerCase() || "main";
    const msg = document.getElementById("storeQrMsg");
    const img = document.getElementById("storeQrPreview");
    const meta = document.getElementById("storeQrMeta");
    if (msg) msg.textContent = "Generating store QR…";
    const { qrUrl, scanValue, payload } = await buildStoreQrImageData(site);
    if (img) { img.src = qrUrl; img.style.display = "block"; }
    if (meta) {
      meta.innerHTML = `
        <div><b>Stores terminal — ${esc(site)}</b></div>
        <div class="muted small" style="margin-top:4px;word-break:break-all;">${esc(scanValue)}</div>
        <div class="muted small" style="margin-top:4px;">
          ${Number(payload?.inventory?.part_count || 0)} parts ·
          ${Number(payload?.inventory?.below_min || 0)} below min ·
          ${Number(payload?.pending_arrivals || 0)} pending arrivals
        </div>
      `;
    }
    if (msg) msg.textContent = "QR ready — print and mount at the stores desk. One QR for all storemen.";
  }

  async function downloadStoreQrPng() {
    const site = String(document.getElementById("storeQrSite")?.value || "main").trim().toLowerCase() || "main";
    const { qrUrl } = await buildStoreQrImageData(site);
    const response = await fetch(qrUrl);
    if (!response.ok) throw new Error("QR image fetch failed");
    const blob = await response.blob();
    const objUrl = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = objUrl;
    a.download = `IRONLOG_Stores_QR_${site}.png`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(objUrl), 10000);
  }

  function printStoreQrLabel() {
    const site = String(document.getElementById("storeQrSite")?.value || "main").trim().toLowerCase() || "main";
    buildStoreQrImageData(site).then(({ qrUrl }) => {
      const win = window.open("", "_blank", "width=600,height=700");
      if (!win) return alert("Pop-up blocked.");
      win.document.write(`<!doctype html><html><head><title>Store QR</title>
        <style>body{font-family:Arial;text-align:center;padding:24px;} img{width:280px;height:280px;} .t{margin-top:12px;font-size:18px;font-weight:700;}</style></head>
        <body><img src="${qrUrl}" alt="Store QR" /><div class="t">IRONLOG Stores — ${esc(site)}</div>
        <div style="margin-top:8px;font-size:13px;color:#475569;">Scan for inventory · receive · labels</div>
        <script>window.onload=()=>{window.focus();window.print();}</script></body></html>`);
      win.document.close();
    }).catch((e) => alert(e.message));
  }

  function bindStoreQrAdmin() {
    document.getElementById("storeQrPreviewBtn")?.addEventListener("click", () => previewStoreQr().catch((e) => alert(e.message)));
    document.getElementById("storeQrDownloadBtn")?.addEventListener("click", () => downloadStoreQrPng().catch((e) => alert(e.message)));
    document.getElementById("storeQrPrintBtn")?.addEventListener("click", () => printStoreQrLabel());
    const demo = document.getElementById("storeQrOpenMobile");
    const siteEl = document.getElementById("storeQrSite");
    const syncDemo = () => {
      if (!demo) return;
      const site = String(siteEl?.value || "main").trim().toLowerCase() || "main";
      demo.href = `./store-mobile.html?site=${encodeURIComponent(site)}`;
    };
    siteEl?.addEventListener("change", syncDemo);
    siteEl?.addEventListener("input", syncDemo);
    syncDemo();
  }

  window.initStoreMobile = initStoreMobile;
  window.bindStoreQrAdmin = bindStoreQrAdmin;
})();
