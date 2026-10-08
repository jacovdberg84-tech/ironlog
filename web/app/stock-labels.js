// IRONLOG/web/app/stock-labels.js — Part labels card (Stock Control → Setup).
// Part of the main app; index.html loads these files in order and they share one global scope.
//
// The print list lives on the server (the stores terminal adds to it too);
// printing uses web/part-labels.js.

let plRows = [];
let plHitsRows = [];

function plMsg(html, kind = "") {
  const el = qs("plMsg");
  if (!el) return;
  el.className = `dlv-msg ${kind}`;
  el.innerHTML = html;
}

function plLayout() {
  return qs("plLayout")?.value || IronlogPartLabels.savedLayout();
}

function plSyncSkip() {
  const L = IronlogPartLabels.LAYOUTS[plLayout()];
  const box = qs("plSkipBox");
  if (box) box.hidden = Boolean(L?.roll);
  const skip = qs("plSkip");
  if (skip && L && !L.roll) skip.max = String(L.cols * L.rows);
}

async function loadLabelQueue() {
  const data = await fetchJson(`${API}/api/stock/labels/queue`);
  plRows = Array.isArray(data.rows) ? data.rows : [];
  renderLabelQueue();
}

function renderLabelQueue() {
  const host = qs("plList");
  if (!host) return;
  const total = plRows.reduce((n, r) => n + Number(r.copies || 0), 0);
  if (qs("plPrint")) {
    qs("plPrint").textContent = total ? `Print ${total} label${total === 1 ? "" : "s"}` : "Print labels";
    qs("plPrint").disabled = !total;
  }
  if (!plRows.length) {
    host.innerHTML = `<div class="muted small" style="padding:10px 0">The print list is empty. Add parts above, or from the stores terminal (Find stock → Label).</div>`;
    return;
  }
  host.innerHTML = `<table class="dlv-table">
    <thead><tr><th>Part</th><th>Bin</th><th class="dlv-num">In stock</th><th class="dlv-num">Labels</th><th></th></tr></thead>
    <tbody>${plRows.map((r) => `
      <tr data-pl-id="${r.queue_id}">
        <td><b>${escapeHtml(r.part_code)}</b><br><small class="muted">${escapeHtml(r.part_name || "")}</small></td>
        <td>${escapeHtml(r.bin || "—")}</td>
        <td class="dlv-num">${escapeHtml(String(r.on_hand ?? ""))}</td>
        <td class="dlv-num"><input class="pl-copies" data-pl-copies type="number" min="1" max="100" value="${Number(r.copies) || 1}" /></td>
        <td><button type="button" class="btn btn-secondary btn-sm" data-pl-remove title="Remove from the list">✕</button></td>
      </tr>`).join("")}</tbody></table>`;
}

async function addToLabelQueue(parts) {
  if (!parts.length) return 0;
  const res = await fetchJson(`${API}/api/stock/labels/queue`, {
    method: "POST",
    body: JSON.stringify({ parts: parts.map((p) => ({ part_code: p.part_code, copies: p.copies || 1 })) }),
  });
  await loadLabelQueue();
  return res.added || 0;
}

async function addLabelSet(set, label) {
  plMsg("Finding parts…");
  const data = await fetchJson(`${API}/api/stock/labels/parts?set=${set}${set === "received" ? "&days=7" : ""}`);
  const rows = (data.rows || []).filter((r) => !plRows.some((q) => q.part_code === r.part_code));
  if (!rows.length) return plMsg(`No more parts ${label}.`);
  const n = await addToLabelQueue(rows);
  plMsg(`Added ${n} part${n === 1 ? "" : "s"} ${label}. Set how many labels each needs, then print.`, "ok");
}

let plSearchTimer = null;
let plSearchSeq = 0;
async function plSearch() {
  const q = String(qs("plFind")?.value || "").trim();
  const hits = qs("plHits");
  const seq = ++plSearchSeq;
  if (q.length < 2) { hits.hidden = true; return; }
  const data = await fetchJson(`${API}/api/stock/parts/search?q=${encodeURIComponent(q)}`);
  if (seq !== plSearchSeq) return;
  plHitsRows = data.rows || [];
  hits.innerHTML = plHitsRows.length
    ? plHitsRows.map((r, i) => `<button type="button" class="dlv-hit" data-pl-hit="${i}"><b>${escapeHtml(r.part_code)}</b> ${escapeHtml(r.part_name || "")}${r.bin ? ` <span class="muted">· ${escapeHtml(r.bin)}</span>` : ""}</button>`).join("")
    : `<div class="muted small" style="padding:8px 10px">No part found.</div>`;
  hits.hidden = false;
}

function printLabelQueue() {
  if (!plRows.length) return;
  const layout = plLayout();
  IronlogPartLabels.saveLayout(layout);
  const count = IronlogPartLabels.print(plRows, {
    layout,
    skip: Math.max(0, Number(qs("plSkip")?.value || 1) - 1),
    site: getSessionSite(),
  });
  plMsg(`Sent ${count} label${count === 1 ? "" : "s"} to the printer. Print at 100% (no "fit to page"). When they have printed correctly, clear the list.`, "ok");
}

/** Start-up: Part labels card. Called once from init() in init.js. */
function wirePartLabels() {
  const card = qs("partLabelsCard");
  if (!card || !window.IronlogPartLabels) return;
  const sel = qs("plLayout");
  if (sel) {
    sel.innerHTML = Object.entries(IronlogPartLabels.LAYOUTS).map(([k, L]) => `<option value="${k}">${escapeHtml(L.name)}</option>`).join("");
    sel.value = IronlogPartLabels.savedLayout();
    sel.addEventListener("change", () => { IronlogPartLabels.saveLayout(sel.value); plSyncSkip(); });
  }
  plSyncSkip();
  let loaded = false;
  const ensureLoaded = () => {
    if (loaded) return;
    loaded = true;
    loadLabelQueue().catch((e) => plMsg(escapeHtml(e.message || String(e)), "bad"));
  };
  qs("stockTabs")?.addEventListener("click", (e) => {
    if (e.target.closest('[data-stock-tab-btn="setup"]')) { loaded = false; ensureLoaded(); }
  });
  if (!card.hidden) ensureLoaded();

  qs("plFind")?.addEventListener("input", () => {
    clearTimeout(plSearchTimer);
    plSearchTimer = setTimeout(() => plSearch().catch((e) => plMsg(escapeHtml(e.message || String(e)), "bad")), 250);
  });
  qs("plHits")?.addEventListener("click", async (e) => {
    const b = e.target.closest("[data-pl-hit]");
    if (!b) return;
    const r = plHitsRows[Number(b.dataset.plHit)];
    qs("plHits").hidden = true;
    qs("plFind").value = "";
    try {
      await addToLabelQueue([r]);
      plMsg(`Added ${escapeHtml(r.part_code)}.`, "ok");
    } catch (err) { plMsg(escapeHtml(err.message || String(err)), "bad"); }
  });
  document.addEventListener("click", (e) => { if (!e.target.closest(".pl-search") && qs("plHits")) qs("plHits").hidden = true; });
  qs("plAddNoBarcode")?.addEventListener("click", () => addLabelSet("no_barcode", "in stock without a box barcode").catch((e) => plMsg(escapeHtml(e.message || String(e)), "bad")));
  qs("plAddReceived")?.addEventListener("click", () => addLabelSet("received", "received in the last 7 days").catch((e) => plMsg(escapeHtml(e.message || String(e)), "bad")));
  qs("plList")?.addEventListener("change", async (e) => {
    const input = e.target.closest("[data-pl-copies]");
    if (!input) return;
    const id = input.closest("[data-pl-id]").dataset.plId;
    const copies = Math.max(1, Math.min(100, Math.round(Number(input.value) || 1)));
    input.value = copies;
    try {
      await fetchJson(`${API}/api/stock/labels/queue/${id}`, { method: "PUT", body: JSON.stringify({ copies }) });
      const row = plRows.find((r) => String(r.queue_id) === id);
      if (row) row.copies = copies;
      renderLabelQueue();
    } catch (err) { plMsg(escapeHtml(err.message || String(err)), "bad"); }
  });
  qs("plList")?.addEventListener("click", async (e) => {
    const b = e.target.closest("[data-pl-remove]");
    if (!b) return;
    const id = b.closest("[data-pl-id]").dataset.plId;
    try {
      await fetchJson(`${API}/api/stock/labels/queue/${id}`, { method: "DELETE" });
      await loadLabelQueue();
    } catch (err) { plMsg(escapeHtml(err.message || String(err)), "bad"); }
  });
  qs("plPrint")?.addEventListener("click", () => {
    try { printLabelQueue(); } catch (err) { plMsg(escapeHtml(err.message || String(err)), "bad"); }
  });
  qs("plClear")?.addEventListener("click", async () => {
    if (!plRows.length || !confirm("Clear the label print list?")) return;
    try {
      await fetchJson(`${API}/api/stock/labels/queue`, { method: "DELETE" });
      await loadLabelQueue();
      plMsg("List cleared.");
    } catch (err) { plMsg(escapeHtml(err.message || String(err)), "bad"); }
  });
}
