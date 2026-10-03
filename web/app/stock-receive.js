// IRONLOG/web/app/stock-receive.js — Stock Control tabs and "Receive a delivery".
// Part of the main app; index.html loads these files in order and they share one global scope.
//
// Tabs: each card in #tab-stock carries data-stock-tab; only the open tab's
// cards show. The Parts Orders cards (workshop requests, requisitions, parts on
// order, off-site repairs) are moved into "Requests & orders" while it is open
// and back to their own page when that page is opened from the menu.
//
// Receive: one delivery (supplier, invoice/GRN, date, currency, store) with many
// lines; part search shows on hand and bin; a new code asks for its description.
// The server checks for duplicates and asks before saving them.

const STOCK_TAB_KEY = "ironlog_stock_tab";
const STOCK_TABS = ["onhand", "receive", "issue", "requests", "counts", "setup", "reports"];

function currentStockTab() {
  try {
    const saved = localStorage.getItem(STOCK_TAB_KEY);
    if (STOCK_TABS.includes(saved)) return saved;
  } catch (_) { /* storage unavailable */ }
  return "onhand";
}

/** The Parts Orders page's cards, moved between that page and the Requests tab. */
function partsOrderCards() {
  const home = qs("tab-parts-tracking");
  const host = qs("stockRequestsHost");
  return { home, host, cards: [...(home?.children || []), ...(host?.children || [])].filter((el) => el.classList?.contains("dash-card")) };
}

function moveRequestCardsTo(where) {
  const { home, host, cards } = partsOrderCards();
  const target = where === "stock" ? host : home;
  if (!target) return;
  cards.forEach((c) => { if (c.parentElement !== target) target.appendChild(c); });
}

function showStockTab(tab) {
  const k = STOCK_TABS.includes(tab) ? tab : "onhand";
  try { localStorage.setItem(STOCK_TAB_KEY, k); } catch (_) { /* ignore */ }
  document.querySelectorAll("#tab-stock [data-stock-tab]").forEach((el) => {
    el.hidden = el.dataset.stockTab !== k;
  });
  document.querySelectorAll("#stockTabs [data-stock-tab-btn]").forEach((b) => {
    const on = b.dataset.stockTabBtn === k;
    b.classList.toggle("on", on);
    b.setAttribute("aria-selected", String(on));
  });
  if (k === "requests") {
    moveRequestCardsTo("stock");
    if (typeof loadPartsTrackingTab === "function") loadPartsTrackingTab().catch(() => {});
  }
  if (k === "receive") {
    ensureDeliveryForm();
    loadRecentDeliveries().catch(() => {});
  }
}

// ---------------------------------------------------------------- Receive

let dlvLineSeq = 0;
let dlvSearchTimer = null;

const dlvMoney = (n) => (Number.isFinite(n) ? n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : "");

function dlvLineHtml(id) {
  return `
    <tr data-dlv-line="${id}">
      <td class="dlv-part">
        <input class="dlv-code w-full" data-dlv-code placeholder="Code or name, e.g. rimula" autocomplete="off" aria-label="Part" />
        <input class="dlv-name w-full" data-dlv-name placeholder="New part: description" hidden aria-label="Description for new part" />
        <div class="dlv-hits" data-dlv-hits hidden></div>
      </td>
      <td data-dlv-onhand class="muted">–</td>
      <td><input class="dlv-bin" data-dlv-bin list="msBinCodeOptions" placeholder="Bin" aria-label="Bin" /></td>
      <td class="dlv-num"><input class="dlv-qty" data-dlv-qty type="number" min="0" step="any" inputmode="decimal" aria-label="Quantity" /></td>
      <td class="dlv-num"><input class="dlv-cost" data-dlv-cost type="number" min="0" step="0.01" inputmode="decimal" placeholder="0.00" aria-label="Unit cost" /></td>
      <td class="dlv-num" data-dlv-linetotal></td>
      <td><button type="button" class="dlv-del" data-dlv-del aria-label="Remove line">×</button></td>
    </tr>`;
}

function addDeliveryLine(prefill = null) {
  const body = qs("dlvLines");
  if (!body) return null;
  dlvLineSeq += 1;
  body.insertAdjacentHTML("beforeend", dlvLineHtml(dlvLineSeq));
  const row = body.lastElementChild;
  if (prefill) pickDeliveryPart(row, prefill);
  return row;
}

function ensureDeliveryForm() {
  if (qs("dlvDate") && !qs("dlvDate").value) qs("dlvDate").value = new Date().toISOString().slice(0, 10);
  const body = qs("dlvLines");
  if (body && !body.children.length) {
    addDeliveryLine();
    addDeliveryLine();
    addDeliveryLine();
  }
  updateDeliveryTotals();
}

function deliveryRows() {
  return [...document.querySelectorAll("#dlvLines [data-dlv-line]")];
}

function pickDeliveryPart(row, p, { focus = true } = {}) {
  row.querySelector("[data-dlv-code]").value = p.part_code;
  row.dataset.known = "1";
  row.dataset.code = p.part_code;
  const name = row.querySelector("[data-dlv-name]");
  name.hidden = true;
  name.value = "";
  row.querySelector("[data-dlv-onhand]").innerHTML = `${escapeHtml(String(p.on_hand ?? 0))}<br><small class="muted">${escapeHtml(p.part_name || "")}</small>`;
  const bin = row.querySelector("[data-dlv-bin]");
  if (!bin.value && p.bin) bin.value = String(p.bin).split(" / ").pop();
  const hits = row.querySelector("[data-dlv-hits]");
  hits.hidden = true;
  hits.innerHTML = "";
  if (focus) row.querySelector("[data-dlv-qty]").focus();
}

/** A typed code that is not in stores: ask for its description (new part). */
function markDeliveryNewPart(row) {
  const code = String(row.querySelector("[data-dlv-code]").value || "").trim().toUpperCase();
  row.dataset.known = "";
  row.dataset.code = code;
  const name = row.querySelector("[data-dlv-name]");
  name.hidden = !code;
  row.querySelector("[data-dlv-onhand]").innerHTML = code ? `<span class="pill">NEW</span>` : "–";
}

async function searchDeliveryPart(row) {
  const input = row.querySelector("[data-dlv-code]");
  const hits = row.querySelector("[data-dlv-hits]");
  const q = String(input.value || "").trim();
  row.dataset.known = "";
  if (q.length < 2) {
    hits.hidden = true;
    return;
  }
  const data = await fetchJson(`${API}/api/stock/parts/search?q=${encodeURIComponent(q)}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  const exact = rows.find((r) => String(r.part_code).toUpperCase() === q.toUpperCase());
  if (exact && rows.length === 1) return pickDeliveryPart(row, exact);
  hits.innerHTML = rows.map((r, i) => `
      <button type="button" class="dlv-hit" data-dlv-hit="${i}">
        <b>${escapeHtml(r.part_code)}</b> ${escapeHtml(r.part_name || "")}
        <span class="muted small">${escapeHtml(String(r.on_hand))} on hand${r.bin ? ` · ${escapeHtml(r.bin)}` : ""}</span>
      </button>`).join("")
    + `<button type="button" class="dlv-hit dlv-hit-new" data-dlv-hit="new">+ New part “${escapeHtml(q.toUpperCase())}”</button>`;
  hits.__rows = rows;
  hits.hidden = false;
}

function updateDeliveryTotals() {
  let total = 0;
  let lines = 0;
  for (const row of deliveryRows()) {
    const qty = Number(row.querySelector("[data-dlv-qty]").value);
    const cost = Number(row.querySelector("[data-dlv-cost]").value);
    const cell = row.querySelector("[data-dlv-linetotal]");
    const has = String(row.querySelector("[data-dlv-code]").value || "").trim();
    if (has && qty > 0) lines += 1;
    if (qty > 0 && cost > 0) {
      total += qty * cost;
      cell.textContent = dlvMoney(qty * cost);
    } else cell.textContent = "";
  }
  const cur = qs("dlvCurrency")?.value || "USD";
  if (qs("dlvTotal")) qs("dlvTotal").textContent = lines ? `${lines} line${lines === 1 ? "" : "s"}${total ? ` · ${cur} ${dlvMoney(total)}` : ""}` : "";
  if (qs("dlvSave")) qs("dlvSave").textContent = lines ? `Receive ${lines} line${lines === 1 ? "" : "s"}` : "Receive delivery";
}

function deliveryMsg(html, kind = "") {
  const el = qs("dlvMsg");
  if (!el) return;
  el.className = `dlv-msg ${kind}`;
  el.innerHTML = html;
}

function collectDelivery() {
  const lines = deliveryRows().map((row) => ({
    part_code: String(row.querySelector("[data-dlv-code]").value || "").trim().toUpperCase(),
    part_name: row.dataset.known ? undefined : String(row.querySelector("[data-dlv-name]").value || "").trim() || undefined,
    quantity: row.querySelector("[data-dlv-qty]").value === "" ? "" : Number(row.querySelector("[data-dlv-qty]").value),
    unit_cost: row.querySelector("[data-dlv-cost]").value === "" ? undefined : Number(row.querySelector("[data-dlv-cost]").value),
    bin_code: String(row.querySelector("[data-dlv-bin]").value || "").trim() || undefined,
  })).filter((l) => l.part_code || l.quantity !== "");
  return {
    supplier: String(qs("dlvSupplier")?.value || "").trim(),
    reference: String(qs("dlvRef")?.value || "").trim(),
    received_date: String(qs("dlvDate")?.value || "").trim(),
    currency: String(qs("dlvCurrency")?.value || "USD"),
    location_code: String(qs("dlvLocation")?.value || "MAIN").trim().toUpperCase() || "MAIN",
    notes: String(qs("dlvNotes")?.value || "").trim(),
    lines,
  };
}

/** POST that keeps the server's reply on a refusal (the duplicate list). */
async function postDelivery(payload) {
  const headers = {
    "Content-Type": "application/json",
    "x-user-name": getSessionUser(),
    "x-user-role": getSessionRole(),
    "x-user-roles": getSessionRoles().join(","),
    "x-site-code": getSessionSite(),
  };
  const tok = getAuthToken();
  if (tok) headers.Authorization = `Bearer ${tok}`;
  const res = await fetch(`${API}/api/stock/deliveries`, { method: "POST", headers, body: JSON.stringify(payload) });
  let data = {};
  try { data = await res.json(); } catch (_) { data = {}; }
  if (res.status === 401 && typeof promptSignInAgain === "function" && LOGIN_GATE_ENABLED) promptSignInAgain();
  return { status: res.status, data };
}

async function saveDelivery(confirm = false) {
  const btn = qs("dlvSave");
  const payload = collectDelivery();
  if (!payload.reference) {
    deliveryMsg("Enter the invoice or GRN number.", "bad");
    qs("dlvRef")?.focus();
    return;
  }
  if (!payload.lines.length) {
    deliveryMsg("Add at least one line.", "bad");
    return;
  }
  if (btn) btn.disabled = true;
  deliveryMsg("Saving…");
  try {
    const { status, data } = await postDelivery({ ...payload, confirm_duplicates: confirm });
    if (status === 409 && data.needs_confirmation) {
      deliveryMsg(`
        <b>Check before saving: this may already be in stock.</b>
        <ul>${(data.duplicates || []).map((d) => `<li>${escapeHtml(d.text)}</li>`).join("")}</ul>
        <div class="dlv-actions"><button type="button" class="btn btn-secondary" data-dlv-cancel>Don't save</button>
        <button type="button" class="btn btn-primary" data-dlv-confirm>It is a new delivery: receive it</button></div>`, "warn");
      return;
    }
    if (status >= 400 || !data.ok) {
      const list = Array.isArray(data.problems) && data.problems.length > 1
        ? `<ul>${data.problems.map((p) => `<li>${escapeHtml(p)}</li>`).join("")}</ul>` : "";
      deliveryMsg(`${escapeHtml(data.error || `Could not save (${status})`)}${list}<br><small>Nothing was received.</small>`, "bad");
      return;
    }
    deliveryMsg(`
      <b>Received ${data.line_count} line${data.line_count === 1 ? "" : "s"} on ${escapeHtml(data.reference)}${data.value_usd ? ` · USD ${dlvMoney(data.value_usd)}` : ""}.</b>
      <ul>${data.lines.map((l) => `<li>${escapeHtml(l.part_code)} ${escapeHtml(l.part_name || "")}: +${l.quantity} → ${l.on_hand_after} on hand${l.new_part ? " (new part)" : ""}</li>`).join("")}</ul>`, "ok");
    clearDeliveryForm(true);
    loadRecentDeliveries().catch(() => {});
    if (typeof loadDashboard === "function") loadDashboard().catch(() => {});
  } catch (e) {
    deliveryMsg(escapeHtml(e.message || String(e)), "bad");
  } finally {
    if (btn) btn.disabled = false;
  }
}

function clearDeliveryForm(keepMessage = false) {
  ["dlvSupplier", "dlvRef", "dlvNotes"].forEach((id) => { if (qs(id)) qs(id).value = ""; });
  if (qs("dlvLines")) qs("dlvLines").innerHTML = "";
  if (!keepMessage) deliveryMsg("");
  ensureDeliveryForm();
  qs("dlvSupplier")?.focus();
}

async function loadRecentDeliveries() {
  const host = qs("dlvRecent");
  if (!host) return;
  const data = await fetchJson(`${API}/api/stock/deliveries?days=30`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  host.classList.toggle("muted", !rows.length);
  host.innerHTML = rows.length
    ? `<table class="dlv-table"><thead><tr><th>Date</th><th>Invoice / GRN</th><th>Supplier</th><th>Lines</th><th class="dlv-num">Value (USD)</th><th>By</th></tr></thead><tbody>${rows.map((d) => `
        <tr title="${escapeHtml((d.lines || []).map((l) => `${l.part_code} +${l.quantity}`).join(", "))}">
          <td>${escapeHtml(d.received_date || String(d.created_at || "").slice(0, 10))}</td>
          <td><b>${escapeHtml(d.reference || "")}</b></td>
          <td>${escapeHtml(d.supplier || "")}</td>
          <td>${escapeHtml((d.lines || []).map((l) => `${l.part_code} ×${l.quantity}`).join(", "))}</td>
          <td class="dlv-num">${d.value_usd ? dlvMoney(Number(d.value_usd)) : ""}</td>
          <td>${escapeHtml(d.received_by || "")}</td>
        </tr>`).join("")}</tbody></table>`
    : "No deliveries received in the last 30 days.";
}

/** "Receive" on a stock card: open the Receive tab with that part on a line. */
function stockReceivePart(code, name) {
  showStockTab("receive");
  const empty = deliveryRows().find((r) => !String(r.querySelector("[data-dlv-code]").value || "").trim());
  const row = empty || addDeliveryLine();
  if (row) pickDeliveryPart(row, { part_code: code, part_name: name, on_hand: "", bin: null });
  updateDeliveryTotals();
}

/** Start-up: Stock Control tabs and the delivery form. Called once from init() in init.js. */
function wireStockReceive() {
  const tabs = qs("stockTabs");
  if (!tabs) return;
  tabs.addEventListener("click", (e) => {
    const b = e.target.closest("[data-stock-tab-btn]");
    if (b) showStockTab(b.dataset.stockTabBtn);
  });
  showStockTab(currentStockTab());

  // Parts Orders opened from the menu: its cards go back to their own page.
  document.querySelectorAll('[data-tab="parts-tracking"]').forEach((el) => el.addEventListener("click", () => moveRequestCardsTo("home"), true));
  qs("tabSelect")?.addEventListener("change", (e) => { if (e.target.value === "parts-tracking") moveRequestCardsTo("home"); }, true);
  document.querySelectorAll('[data-tab="stock"]').forEach((el) => el.addEventListener("click", () => {
    if (currentStockTab() === "requests") setTimeout(() => moveRequestCardsTo("stock"), 0);
  }));

  const card = qs("deliveryCard");
  qs("dlvAddLine")?.addEventListener("click", () => addDeliveryLine()?.querySelector("[data-dlv-code]").focus());
  qs("dlvClear")?.addEventListener("click", () => clearDeliveryForm());
  qs("dlvSave")?.addEventListener("click", () => saveDelivery(false));
  qs("dlvCurrency")?.addEventListener("change", updateDeliveryTotals);
  qs("dlvMsg")?.addEventListener("click", (e) => {
    if (e.target.closest("[data-dlv-confirm]")) saveDelivery(true);
    if (e.target.closest("[data-dlv-cancel]")) deliveryMsg("Not saved.", "");
  });
  card?.addEventListener("input", (e) => {
    const row = e.target.closest("[data-dlv-line]");
    if (!row) return;
    if (e.target.matches("[data-dlv-code]")) {
      clearTimeout(dlvSearchTimer);
      dlvSearchTimer = setTimeout(() => searchDeliveryPart(row).catch(() => {}), 250);
    }
    updateDeliveryTotals();
  });
  card?.addEventListener("click", (e) => {
    const hit = e.target.closest("[data-dlv-hit]");
    if (hit) {
      const row = hit.closest("[data-dlv-line]");
      const hits = row.querySelector("[data-dlv-hits]");
      if (hit.dataset.dlvHit === "new") {
        hits.hidden = true;
        markDeliveryNewPart(row);
        row.querySelector("[data-dlv-name]").focus();
      } else {
        pickDeliveryPart(row, hits.__rows[Number(hit.dataset.dlvHit)]);
      }
      updateDeliveryTotals();
      return;
    }
    const del = e.target.closest("[data-dlv-del]");
    if (del) {
      del.closest("[data-dlv-line]").remove();
      if (!deliveryRows().length) addDeliveryLine();
      updateDeliveryTotals();
    }
  });
  card?.addEventListener("keydown", (e) => {
    // Enter on the last cost field adds a line, like a spreadsheet.
    if (e.key === "Enter" && e.target.matches("[data-dlv-cost]")) {
      e.preventDefault();
      const row = e.target.closest("[data-dlv-line]");
      const next = row.nextElementSibling || addDeliveryLine();
      next.querySelector("[data-dlv-code]").focus();
    }
  });
  card?.addEventListener("focusout", (e) => {
    // Leaving a code that matched nothing: treat it as a new part.
    if (!e.target.matches("[data-dlv-code]")) return;
    const row = e.target.closest("[data-dlv-line]");
    setTimeout(() => {
      const hits = row.querySelector("[data-dlv-hits]");
      if (document.activeElement === e.target || hits.contains(document.activeElement)) return;
      hits.hidden = true;
      const typed = String(e.target.value || "").trim().toUpperCase();
      if (row.dataset.known || !typed) return;
      const exact = (hits.__rows || []).find((r) => String(r.part_code).toUpperCase() === typed);
      if (exact) pickDeliveryPart(row, exact, { focus: false });
      else markDeliveryNewPart(row);
    }, 200);
  });
}
