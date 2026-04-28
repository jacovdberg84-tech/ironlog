(function () {
  function qs(id) { return document.getElementById(id); }
  function esc(v) {
    return String(v == null ? "" : v)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  async function fetchJson(url, options) {
    const res = await fetch(url, options);
    const text = await res.text();
    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch { data = null; }
    if (!res.ok) {
      const msg = data?.error || data?.message || text || `Request failed (${res.status})`;
      throw new Error(msg);
    }
    return data || {};
  }

  function parseQuery() {
    const u = new URL(window.location.href);
    const assetCode = String(u.searchParams.get("asset_code") || "").trim();
    const viewRaw = String(u.searchParams.get("view") || "history").trim().toLowerCase();
    const view = ["inspections", "history", "workorders"].includes(viewRaw) ? viewRaw : "history";
    return { assetCode, view };
  }

  function toDate(v) {
    const s = String(v || "");
    return s ? s.slice(0, 10) : "-";
  }

  function inspectionLineFromChecklist(list) {
    const rows = Array.isArray(list) ? list : [];
    const fail = rows.filter((r) => r?.ok === false).length;
    const pass = rows.filter((r) => r?.ok === true).length;
    if (!rows.length) return "Checklist: -";
    return `Checklist: ${pass} pass, ${fail} fail`;
  }

  function inspectionCard(r, kind) {
    const type = kind === "manager"
      ? String(r.inspection_type || "manager").replaceAll("_", " ")
      : "artisan";
    const notes = String(r.notes || "").trim();
    return `
      <div class="item">
        <div class="title">${esc(toDate(r.inspection_date))} - ${esc(type)}</div>
        <div class="meta">Inspector: ${esc(r.inspector_name || "-")} | Hours: ${r.machine_hours != null ? esc(Number(r.machine_hours).toFixed(1)) : "-"}</div>
        <div class="small" style="margin-top:4px;">${esc(inspectionLineFromChecklist(r.checklist))}</div>
        ${notes ? `<div class="small" style="margin-top:4px;">Notes: ${esc(notes)}</div>` : ""}
      </div>
    `;
  }

  function historyCard(ev) {
    const d = ev?.details || {};
    const lines = [];
    if (d.status) lines.push(`Status: ${d.status}`);
    if (d.source) lines.push(`Source: ${d.source}`);
    if (d.downtime_hours != null) lines.push(`Downtime: ${Number(d.downtime_hours || 0).toFixed(1)}h`);
    if (d.notes) lines.push(`Notes: ${String(d.notes).slice(0, 220)}`);
    return `
      <div class="item">
        <div class="title">${esc(toDate(ev.date))} - ${esc(ev.title || ev.type || "event")}</div>
        <div class="meta">Type: ${esc(ev.type || "-")} ${ev.work_order_id ? `| WO #${esc(ev.work_order_id)}` : ""}</div>
        ${lines.length ? `<div class="small" style="margin-top:4px;">${esc(lines.join(" | "))}</div>` : ""}
      </div>
    `;
  }

  function workOrderCard(ev) {
    const d = ev?.details || {};
    const status = String(d.status || "").toLowerCase();
    const badgeClass = status === "closed" ? "ok" : (status === "open" || status === "assigned" || status === "in_progress") ? "warn" : "bad";
    const woId = Number(ev.work_order_id || 0);
    return `
      <div class="item">
        <div class="title">WO #${woId || "-"} - ${esc(toDate(ev.date))}</div>
        <div class="meta">
          <span class="pill ${badgeClass}">${esc(String(d.status || "unknown").toUpperCase())}</span>
          <span class="pill">${esc(d.source || "-")}</span>
        </div>
        <div class="small" style="margin-top:4px;">Opened: ${esc(String(d.opened_at || "-"))} | Closed: ${esc(String(d.closed_at || "-"))}</div>
      </div>
    `;
  }

  async function loadPage() {
    const { assetCode, view } = parseQuery();
    if (!assetCode) {
      qs("sub").textContent = "Missing asset_code in URL.";
      return;
    }
    const titleMap = {
      inspections: "Asset QR - Inspections",
      history: "Asset QR - Service History",
      workorders: "Asset QR - Work Orders",
    };
    qs("viewTitle").textContent = titleMap[view] || "Asset QR Detail";
    qs("sub").textContent = `Loading ${assetCode}...`;
    qs("backQrBtn").href = `./asset-qr.html?asset_code=${encodeURIComponent(assetCode)}`;

    const [qr, history] = await Promise.all([
      fetchJson(`/api/assets/${encodeURIComponent(assetCode)}/qr-profile`),
      fetchJson(`/api/assets/${encodeURIComponent(assetCode)}/history`),
    ]);
    const asset = history?.asset || qr?.live_preview?.asset || {};
    const assetId = Number(qr?.live_preview?.asset?.id || 0);
    qs("assetLabel").textContent = `${asset.asset_code || assetCode} - ${asset.asset_name || ""}`.trim();

    const listEl = qs("list");
    let html = "";
    let total = 0;

    if (view === "inspections") {
      let managerRows = [];
      let artisanRows = [];
      if (assetId > 0) {
        const [mgr, art] = await Promise.all([
          fetchJson(`/api/maintenance/inspections?asset_id=${assetId}`),
          fetchJson(`/api/maintenance/artisan-inspections?asset_id=${assetId}`),
        ]);
        managerRows = Array.isArray(mgr?.rows) ? mgr.rows : [];
        artisanRows = Array.isArray(art?.rows) ? art.rows : [];
      }
      const rows = [
        ...managerRows.map((r) => ({ kind: "manager", row: r })),
        ...artisanRows.map((r) => ({ kind: "artisan", row: r })),
      ].sort((a, b) => String(b.row?.inspection_date || "").localeCompare(String(a.row?.inspection_date || "")));
      total = rows.length;
      html = rows.map((x) => inspectionCard(x.row, x.kind)).join("");
    } else if (view === "workorders") {
      const rows = (Array.isArray(history?.history) ? history.history : []).filter((e) => String(e?.type || "") === "work_order");
      total = rows.length;
      html = rows.map(workOrderCard).join("");
    } else {
      const serviceTypes = new Set([
        "breakdown", "get_slip", "component_slip", "ops_slip", "damage_report", "tyre_change", "tyre_inspection",
      ]);
      const rows = (Array.isArray(history?.history) ? history.history : []).filter((e) => serviceTypes.has(String(e?.type || "")));
      total = rows.length;
      html = rows.map(historyCard).join("");
    }

    qs("totalCount").textContent = String(total);
    listEl.innerHTML = html || `<div class="empty">No ${esc(view)} data found for this asset yet.</div>`;
    qs("sub").textContent = `${assetCode} loaded`;
  }

  qs("refreshBtn")?.addEventListener("click", () => {
    loadPage().catch((e) => {
      qs("sub").textContent = `Error: ${e.message || e}`;
    });
  });

  loadPage().catch((e) => {
    qs("sub").textContent = `Error: ${e.message || e}`;
  });
})();
