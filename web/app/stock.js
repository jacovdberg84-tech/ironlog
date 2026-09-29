// IRONLOG/web/app/stock.js — Stock monitor, stock reports, stores part orders, parts tracking.
// Part of the main app; index.html loads these files in order and they share one global scope.

async function loadStockMonitor() {
  const filter = (qs("stockPartFilter")?.value || "").trim();
  const page = window.stockMonitorPage || 1;
  const pageSize = 20;
  const q = filter ? `?part_code=${encodeURIComponent(filter)}` : "";
  const data = await fetchJson(`${API}/api/stock/monitor${q}`);

  setText("smBelowMin", Number(data.summary?.below_min || 0));
  setText("smCriticalBelow", Number(data.summary?.critical_below_min || 0));
  setText("smTotalParts", Number(data.summary?.total_parts || 0));

  const list = qs("stockMonitorList");
  if (!list) return;
  list.innerHTML = "";
  const rows = data.rows || [];
  const start = (page - 1) * pageSize;
  const end = start + pageSize;
  rows.slice(start, end).forEach((r) => {
    list.appendChild(
      item(
        `<b>${r.part_code}</b> – ${Number(r.on_hand || 0).toFixed(1)} on hand ${
          r.below_min ? "<span class='pill red'>LOW</span>" : ""
        }<br><small>${r.part_name || ""} | Min: ${Number(r.min_stock || 0).toFixed(1)}</small>`
      )
    );
  });
  if (!rows.length) list.appendChild(item("<small>No parts found for current filter.</small>"));

  // Update paging info
  const pageInfo = qs("stockPageInfo");
  if (pageInfo) {
    const total = rows.length;
    const totalPages = Math.max(1, Math.ceil(total / pageSize));
    pageInfo.textContent = `Page ${page} of ${totalPages}`;
  }
}

// Paging controls
window.stockMonitorPage = 1;
function updateStockMonitorPage(delta) {
  const rows = window.lastStockMonitorRows || [];
  const pageSize = 20;
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  window.stockMonitorPage = Math.max(1, Math.min(window.stockMonitorPage + delta, totalPages));
  loadStockMonitor();
}

// Live filter
const stockPartFilter = qs("stockPartFilter");
if (stockPartFilter) {
  stockPartFilter.addEventListener("input", () => {
    window.stockMonitorPage = 1;
    loadStockMonitor();
  });
}

const prevBtn = qs("prevStockPage");
if (prevBtn) prevBtn.onclick = () => updateStockMonitorPage(-1);
const nextBtn = qs("nextStockPage");
if (nextBtn) nextBtn.onclick = () => updateStockMonitorPage(1);

// Save last rows for paging
const origLoadStockMonitor = loadStockMonitor;
loadStockMonitor = async function() {
  const filter = (qs("stockPartFilter")?.value || "").trim();
  const q = filter ? `?part_code=${encodeURIComponent(filter)}` : "";
  const data = await fetchJson(`${API}/api/stock/monitor${q}`);
  window.lastStockMonitorRows = data.rows || [];
  // Call original logic
  await origLoadStockMonitor.apply(this, arguments);
};

let stockPageData = { rows: [], recent: [], summary: null };

const STOCK_CATEGORY_OPTIONS = [
  ["part", "Parts"],
  ["component", "Components"],
  ["get", "G.E.T"],
  ["tyre", "Tyres"],
  ["oil", "Oils & lubricants"],
];

function filterStockDisplayRows(rows) {
  const q = (qs("spFilter")?.value || "").trim().toLowerCase();
  const category = (qs("spCategory")?.value || "").trim();
  let arr = Array.isArray(rows) ? rows : [];
  if (category) arr = arr.filter((r) => String(r.stock_category || "part") === category);
  if (!q) return arr;
  return arr.filter(
    (r) =>
      String(r.part_code || "").toLowerCase().includes(q) ||
      String(r.part_name || "").toLowerCase().includes(q)
  );
}

function renderStockInventoryTable(rows) {
  const host = qs("spList");
  if (!host) return;
  if (!rows.length) {
    host.innerHTML = `<div class="stores-inventory-empty muted small">No parts found for current filter.</div>`;
    return;
  }

  const cards = rows
    .map((r) => {
      const onHand = Number(r.on_hand || 0);
      const min = Number(r.min_stock || 0);
      const unit = Number(r.unit_cost || 0);
      const value = Number(r.stock_value ?? onHand * unit);
      const below = Boolean(r.below_min);
      const critical = Boolean(r.critical);
      const cardCls = [
        "stores-stock-card",
        below ? "is-low" : "",
        critical && !below ? "is-critical" : "",
      ]
        .filter(Boolean)
        .join(" ");
      const status = below
        ? `<span class="stores-inv-status stores-inv-status--low">Low</span>`
        : critical
          ? `<span class="stores-inv-status stores-inv-status--watch">Critical</span>`
          : `<span class="stores-inv-status stores-inv-status--ok">OK</span>`;
      const shortage = Math.max(0, min - onHand);
      const category = String(r.stock_category || "part");
      const categoryControl = canEditStockCategory()
        ? `<label class="stores-stock-category"><span class="sr-only">Category for ${escapeHtml(r.part_code || "")}</span>
            <select data-stock-category-select title="${r.stock_category_source === "manual" ? "Set by stores" : "Set automatically — change if wrong"}">
              ${STOCK_CATEGORY_OPTIONS.map(([key, label]) => `<option value="${key}"${key === category ? " selected" : ""}>${label}</option>`).join("")}
            </select></label>`
        : `<span class="stores-stock-category">${escapeHtml(r.stock_category_label || "Parts")}</span>`;
      return `<article class="${cardCls}" data-stock-code="${spoAttrVal(r.part_code || "")}" data-stock-name="${spoAttrVal(r.part_name || "")}">
        <header>
          <div><span class="stores-inv-code">${escapeHtml(r.part_code || "")}</span><div class="stores-stock-name">${escapeHtml(r.part_name || "—")}</div>${categoryControl}</div>
          ${status}
        </header>
        <div class="stores-stock-metrics">
          <div><span>On hand</span><strong>${onHand.toFixed(1)}</strong></div>
          <div><span>Minimum</span><strong>${min.toFixed(1)}</strong></div>
          <div><span>${below ? "Short by" : "Buffer"}</span><strong>${below ? shortage.toFixed(1) : Math.max(0, onHand - min).toFixed(1)}</strong></div>
          <div><span>Stock value</span><strong>$${value.toFixed(2)}</strong></div>
        </div>
        <footer>
          <span class="muted small">Unit price $${unit.toFixed(2)}</span>
          <div class="stores-stock-actions">
            <button type="button" class="btn btn-secondary btn-sm" data-stock-action="receive">Receive</button>
            <button type="button" class="btn btn-primary btn-sm" data-stock-action="issue"${onHand <= 0 ? " disabled" : ""}>Issue</button>
            <button type="button" class="btn btn-secondary btn-sm" data-stock-action="count">Count</button>
          </div>
        </footer>
      </article>`;
    })
    .join("");

  host.innerHTML = `
    <div class="stores-stock-grid">${cards}</div>
    <div class="stores-inventory-foot muted small">${rows.length} item${rows.length === 1 ? "" : "s"} shown</div>
  `;
}

function canEditStockCategory() {
  return getSessionRoles().some((r) => ["admin", "supervisor", "stores", "storeman", "workshop_admin"].includes(r));
}

async function saveStockCategory(select) {
  const card = select.closest("[data-stock-code]");
  const code = String(card?.dataset?.stockCode || "").trim();
  if (!code) return;
  select.disabled = true;
  try {
    const res = await fetchJson(`${API}/api/stock/part-category`, {
      method: "POST",
      body: JSON.stringify({ part_code: code, category: select.value }),
    });
    const row = stockPageData.rows.find((r) => r.part_code === code);
    if (row) {
      row.stock_category = res.stock_category;
      row.stock_category_label = res.stock_category_label;
      row.stock_category_source = res.stock_category_source;
    }
    setStatus(`${code} is now in ${res.stock_category_label}.`);
    if (qs("spCategory")?.value) refreshStockInventoryDisplay();
  } catch (e) {
    setStatus(`Could not change the category for ${code}: ${e.message || e}`);
    const row = stockPageData.rows.find((r) => r.part_code === code);
    if (row) select.value = row.stock_category || "part";
  } finally {
    select.disabled = false;
  }
}

function openStockAction(card, action) {
  const code = String(card?.dataset?.stockCode || "").trim();
  const name = String(card?.dataset?.stockName || "").trim();
  if (!code) return;
  let target = null;
  if (action === "receive") {
    if (qs("msPart")) qs("msPart").value = code;
    if (qs("msPartDesc")) qs("msPartDesc").value = name;
    if (qs("msType")) qs("msType").value = "in";
    if (qs("msQty")) qs("msQty").value = "1";
    updateManualStockCostRowVisibility();
    target = qs("msPart")?.closest(".dash-card");
    setStatus(`Ready to receive ${code}.`);
  } else if (action === "issue") {
    if (qs("saPart")) qs("saPart").value = code;
    if (qs("saQty")) qs("saQty").value = "1";
    target = qs("saPart")?.closest(".dash-card");
    setStatus(`Ready to issue ${code}. Select an asset or work order.`);
  } else if (action === "count") {
    if (qs("icPartCode")) qs("icPartCode").value = code;
    target = qs("icPartCode")?.closest(".dash-card");
    loadInventoryControl().catch((e) => setStatus(`Inventory count load failed: ${e.message}`));
    setStatus(`Ready to count ${code}.`);
  }
  if (!target) return;
  target.dataset.collapsed = "false";
  target.scrollIntoView({ behavior: "smooth", block: "center" });
  setTimeout(() => target.querySelector("input:not([disabled]), select:not([disabled])")?.focus(), 350);
}

function openUtilityWorkflow(cardId) {
  const card = document.getElementById(String(cardId || ""));
  if (!card) return;
  card.classList.remove("collapsed");
  card.dataset.collapsed = "false";
  card.scrollIntoView({ behavior: "smooth", block: "start" });
  setTimeout(() => card.querySelector("input:not([disabled]), select:not([disabled]), button.btn-primary")?.focus(), 350);
}

function renderStockRecentTable(recent) {
  const host = qs("spRecent");
  if (!host) return;
  const rows = Array.isArray(recent) ? recent : [];
  if (!rows.length) {
    host.innerHTML = `<div class="stores-inventory-empty muted small">No stock movements yet.</div>`;
    return;
  }

  const bodyRows = rows
    .map((r) => {
      const qty = Number(r.quantity || 0);
      const qtyCls = qty >= 0 ? "stores-mv-qty--in" : "stores-mv-qty--out";
      return `<tr>
        <td class="stores-mv-col-date">${escapeHtml(String(r.created_at || "").slice(0, 16))}</td>
        <td class="stores-mv-col-code"><span class="stores-inv-code">${escapeHtml(r.part_code || "")}</span></td>
        <td class="stores-mv-col-desc">${escapeHtml(r.part_name || "—")}</td>
        <td class="stores-mv-col-num ${qtyCls}">${qty >= 0 ? "+" : ""}${qty.toFixed(1)}</td>
        <td class="stores-mv-col-type">${escapeHtml(r.movement_type || "—")}</td>
        <td class="stores-mv-col-loc">${escapeHtml(r.location_code || "NO-LOC")}</td>
        <td class="stores-mv-col-ref">${escapeHtml(r.reference || "—")}</td>
      </tr>`;
    })
    .join("");

  host.innerHTML = `
    <div class="stores-inventory-scroll stores-movements-scroll">
      <table class="stores-inventory-table stores-movements-table">
        <thead>
          <tr>
            <th scope="col">Date</th>
            <th scope="col">Part code</th>
            <th scope="col">Description</th>
            <th class="stores-inv-col-num" scope="col">Qty</th>
            <th scope="col">Type</th>
            <th scope="col">Location</th>
            <th scope="col">Reference</th>
          </tr>
        </thead>
        <tbody>${bodyRows}</tbody>
      </table>
    </div>
  `;
}

function refreshStockInventoryDisplay() {
  const onlyLow = Boolean(qs("spOnlyLow")?.checked);
  let baseRows = filterStockDisplayRows(stockPageData.rows);
  if (onlyLow) baseRows = baseRows.filter((r) => Boolean(r.below_min));
  renderStockInventoryTable(sortStockRows(baseRows));
}

function sortStockRows(rows) {
  const mode = (qs("spSort")?.value || "critical_then_low").trim();
  const arr = Array.isArray(rows) ? [...rows] : [];

  if (mode === "on_hand_asc") {
    return arr.sort((a, b) => Number(a.on_hand || 0) - Number(b.on_hand || 0));
  }
  if (mode === "on_hand_desc") {
    return arr.sort((a, b) => Number(b.on_hand || 0) - Number(a.on_hand || 0));
  }
  if (mode === "part_code_desc") {
    return arr.sort((a, b) => String(b.part_code || "").localeCompare(String(a.part_code || "")));
  }
  if (mode === "part_code_asc") {
    return arr.sort((a, b) => String(a.part_code || "").localeCompare(String(b.part_code || "")));
  }

  // default: critical first, then below min, then lowest on hand
  return arr.sort((a, b) => {
    const c = Number(Boolean(b.critical)) - Number(Boolean(a.critical));
    if (c !== 0) return c;
    const low = Number(Boolean(b.below_min)) - Number(Boolean(a.below_min));
    if (low !== 0) return low;
    return Number(a.on_hand || 0) - Number(b.on_hand || 0);
  });
}

async function loadStockOnHandPage() {
  const filter = (qs("spFilter")?.value || "").trim();
  const q = filter ? `?part_code=${encodeURIComponent(filter)}` : "";

  setStatus("Loading stock on hand...");
  setSkeleton("spList", 2);
  setSkeleton("spRecent", 2);

  const data = await fetchJson(`${API}/api/stock/monitor${q}`);
  stockPageData = {
    rows: Array.isArray(data.rows) ? data.rows : [],
    recent: Array.isArray(data.recent) ? data.recent : [],
    summary: data.summary || null,
  };

  setText("spTotalParts", Number(data.summary?.total_parts || 0));
  setText("spBelowMin", Number(data.summary?.below_min || 0));
  setText("spCriticalBelow", Number(data.summary?.critical_below_min || 0));
  setText("spTotalOnHand", Number(data.summary?.total_on_hand || 0).toFixed(1));
  setText("spTotalValue", Number(data.summary?.total_stock_value || 0).toFixed(2));

  refreshStockInventoryDisplay();
  renderStockRecentTable(stockPageData.recent);

  setStatus("Stock on hand ready.");
}

function exportStockOnHandCsv() {
  const onlyLow = Boolean(qs("spOnlyLow")?.checked);
  const baseRows = Array.isArray(stockPageData.rows) ? stockPageData.rows : [];
  const rows = onlyLow ? baseRows.filter((r) => Boolean(r.below_min)) : baseRows;
  if (!rows.length) return alert("Load stock data first.");

  const header = "part_code,part_name,category,on_hand,min_stock,unit_cost,stock_value,critical,below_min";
  const lines = rows.map((r) =>
    [
      r.part_code || "",
      `"${String(r.part_name || "").replace(/"/g, '""')}"`,
      `"${String(r.stock_category_label || "Parts").replace(/"/g, '""')}"`,
      Number(r.on_hand || 0),
      Number(r.min_stock || 0),
      Number(r.unit_cost || 0),
      Number(r.stock_value || 0),
      r.critical ? 1 : 0,
      r.below_min ? 1 : 0,
    ].join(",")
  );
  const csv = [header, ...lines].join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "stock_on_hand.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  setStatus("Stock CSV exported.");
}

/** Last loaded stock movements report (period ledger). */
let stockMovementsReportData = { rows: [], date_from: "", date_to: "", part_filter: "", truncated: false };

function ensureStockMovementsReportDates() {
  const fromEl = qs("smrDateFrom");
  const toEl = qs("smrDateTo");
  if (!fromEl || !toEl) return;
  if (!fromEl.value || !toEl.value) {
    const today = new Date();
    const start = new Date(today.getFullYear(), today.getMonth(), 1);
    const pad = (n) => String(n).padStart(2, "0");
    const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
    if (!fromEl.value) fromEl.value = ymd(start);
    if (!toEl.value) toEl.value = ymd(today);
  }
}

function normalizeStockReportPartInput(raw) {
  let s = String(raw || "").trim().replace(/\s+/g, " ");
  if (!s) return "";
  const dashIdx = s.indexOf(" - ");
  if (dashIdx > 0) s = s.slice(0, dashIdx).trim();
  return s;
}

async function loadStockMovementsReport() {
  ensureStockMovementsReportDates();
  const date_from = (qs("smrDateFrom")?.value || "").trim();
  const date_to = (qs("smrDateTo")?.value || "").trim();
  const part_code = normalizeStockReportPartInput(qs("smrPartFilter")?.value || "");
  if (!date_from || !date_to) return alert("Choose From and To dates.");

  const q = new URLSearchParams();
  q.set("date_from", date_from);
  q.set("date_to", date_to);
  if (part_code) q.set("part_code", part_code);

  setStatus("Loading stock movements report...");
  setSkeleton("smrList", 2);

  const data = await fetchJson(`${API}/api/stock/movements-report?${q.toString()}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  const summary = data.summary || {};

  stockMovementsReportData = {
    rows,
    date_from: data.date_from || date_from,
    date_to: data.date_to || date_to,
    part_filter: data.part_filter || part_code || "",
    truncated: Boolean(data.truncated),
    total_matching: Number(data.total_matching || rows.length),
    row_limit: Number(data.row_limit || 0),
  };

  setText("smrMoveCount", String(summary.movement_count ?? "-"));
  setText("smrQtyIn", summary.qty_in != null ? Number(summary.qty_in).toFixed(2) : "-");
  setText("smrQtyOut", summary.qty_out != null ? Number(summary.qty_out).toFixed(2) : "-");
  setText("smrNetQty", summary.net_qty != null ? Number(summary.net_qty).toFixed(2) : "-");

  const trunc = qs("smrTruncNote");
  const truncText = qs("smrTruncText");
  if (trunc && truncText) {
    if (stockMovementsReportData.truncated) {
      trunc.style.display = "";
      truncText.textContent = `Showing the latest ${rows.length} of ${stockMovementsReportData.total_matching} movements in this period (export CSV includes loaded rows only). Increase precision with a narrower date range or part filter.`;
    } else {
      trunc.style.display = "none";
      truncText.textContent = "";
    }
  }

  const list = qs("smrList");
  if (list) {
    list.innerHTML = "";
    rows.forEach((r) => {
      const qty = Number(r.quantity || 0);
      const loc = r.location_code || "—";
      const bin = r.bin_code ? String(r.bin_code) : "";
      list.appendChild(
        item(
          `<b>${r.part_code || ""}</b> — ${qty.toFixed(2)} (${r.movement_type || ""})` +
            `<br><small>${r.movement_at || ""} | ${loc}${bin ? " / " + bin : ""} | ${r.reference || "—"}</small>` +
            `<br><small>${r.part_name || ""}</small>`
        )
      );
    });
    if (!rows.length) list.appendChild(item("<small>No movements in this period for the current filter.</small>"));
  }

  setStatus("Stock movements report ready.");
}

function exportStockMovementsReportCsv() {
  const { rows, date_from, date_to } = stockMovementsReportData;
  if (!rows.length) return alert("Load the stock movements report first.");

  const header =
    "movement_at,part_code,part_name,movement_type,quantity,reference,location_code,bin_code";
  const lines = rows.map((r) =>
    [
      r.movement_at || "",
      r.part_code || "",
      `"${String(r.part_name || "").replace(/"/g, '""')}"`,
      r.movement_type || "",
      Number(r.quantity || 0),
      `"${String(r.reference || "").replace(/"/g, '""')}"`,
      r.location_code || "",
      r.bin_code || "",
    ].join(",")
  );
  const csv = [header, ...lines].join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `stock_movements_${date_from || "from"}_${date_to || "to"}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  setStatus("Stock movements CSV exported.");
}

function openStockMovementsReportPdf() {
  ensureStockMovementsReportDates();
  const date_from = (qs("smrDateFrom")?.value || "").trim();
  const date_to = (qs("smrDateTo")?.value || "").trim();
  const part_code = normalizeStockReportPartInput(qs("smrPartFilter")?.value || "");
  if (!date_from || !date_to) return alert("Choose From and To dates.");
  const q = new URLSearchParams();
  q.set("date_from", date_from);
  q.set("date_to", date_to);
  if (part_code) q.set("part_code", part_code);
  return openAuthedReport(`${API}/api/reports/stock-movements.pdf?${q.toString()}`, {
    filename: `IRONLOG_Stock_Movements_${date_from}_to_${date_to}.pdf`,
  });
}

function openStockOnHandPdf() {
  const filter = (qs("spFilter")?.value || "").trim();
  const q = filter ? `?part_code=${encodeURIComponent(filter)}` : "";
  return openAuthedReport(`${API}/api/reports/stock-monitor.pdf${q}`, {
    filename: "IRONLOG_Stock_On_Hand.pdf",
  });
}

function ensureGmStockReportDate() {
  const input = qs("gmStockReportDate");
  if (input && !input.value) input.value = todayLocalYmd();
}

function updateGmStockReportHelp() {
  const period = String(qs("gmStockReportPeriod")?.value || "monthly").trim().toLowerCase();
  const label = qs("gmStockReportDateLabel");
  const help = qs("gmStockReportHelp");
  if (period === "weekly") {
    if (label) label.textContent = "Week ending";
    if (help) help.textContent = "Weekly uses the seven days ending on the selected date.";
    return;
  }
  if (label) label.textContent = "Month containing";
  if (help) help.textContent = "Monthly uses the full calendar month containing the selected date.";
}

async function downloadGmStockReportXlsx() {
  ensureGmStockReportDate();
  const period = String(qs("gmStockReportPeriod")?.value || "monthly").trim().toLowerCase() === "weekly"
    ? "weekly"
    : "monthly";
  const reportDate = String(qs("gmStockReportDate")?.value || "").trim();
  if (!reportDate) return alert("Choose a report date.");
  const query = new URLSearchParams({ period, report_date: reportDate });
  setStatus("Preparing GM stock Excel...");
  const ok = await downloadAuthedFile(
    `${API}/api/stock/gm-stock-report.xlsx?${query.toString()}`,
    `IRONLOG_${period === "weekly" ? "Weekly" : "Monthly"}_Stock_Report_${reportDate}.xlsx`,
  );
  if (ok) setStatus("GM stock Excel downloaded.");
}

let storesPartOrdersCache = [];

function ensureStoresPartOrderDates() {
  const fromEl = qs("spoDateFrom");
  const toEl = qs("spoDateTo");
  const orderEl = qs("spoOrderDate");
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), 1);
  const pad = (n) => String(n).padStart(2, "0");
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  if (fromEl && !fromEl.value) fromEl.value = ymd(start);
  if (toEl && !toEl.value) toEl.value = ymd(today);
  if (orderEl && !orderEl.value) orderEl.value = ymd(today);
}

function spoStatusLabel(status) {
  const s = String(status || "").toLowerCase();
  if (s === "on_order") return "On order";
  if (s === "in_transit") return "In transit";
  if (s === "arrived") return "Arrived";
  if (s === "cancelled") return "Cancelled";
  return s || "—";
}

function moneyUsd(n) {
  return Number(n || 0).toFixed(2);
}

function renderStoresPartOrdersSummary(summary) {
  const s = summary || {};
  setText("spoOnOrderValue", moneyUsd(s.on_order?.value));
  setText("spoInTransitValue", moneyUsd(s.in_transit?.value));
  setText("spoArrivedValue", moneyUsd(s.arrived?.value));
  setText("spoPendingValue", moneyUsd(s.total_pending));
  setText("spoForecastValue", moneyUsd(s.total_forecast));
}

function handleStoresPartOrderReceipt(data) {
  const receipt = data?.stock_receipt;
  if (!receipt) return;
  if (receipt.received) {
    setStatus(
      `Received ${receipt.qty} × ${receipt.part_code} into store inventory (${receipt.on_hand_after} on hand).`
    );
    loadStockOnHandPage().catch(() => {});
    loadStockMonitor().catch(() => {});
    return;
  }
  if (receipt.already) {
    setStatus("Purchase already received into store inventory.");
    return;
  }
  if (receipt.error) {
    alert(`Status saved, but stock was not updated: ${receipt.error}`);
  }
}

function spoAttrVal(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;");
}

function renderStoresPartOrdersTable(rows) {
  const host = qs("spoList");
  if (!host) return;
  if (!Array.isArray(rows) || !rows.length) {
    host.innerHTML = `<div class="muted small">No purchases in this period. Add a line above.</div>`;
    return;
  }
  host.innerHTML = `
    <table class="gridTable spo-purchases-table" style="min-width:1200px;">
      <thead>
        <tr>
          <th>Order date</th>
          <th>Part</th>
          <th style="text-align:right;">Qty</th>
          <th style="text-align:right;">Unit $</th>
          <th style="text-align:right;">Line $</th>
          <th>Supplier</th>
          <th>PO</th>
          <th>Req #</th>
          <th>ETA</th>
          <th>Status</th>
          <th>Inventory</th>
          <th></th>
        </tr>
      </thead>
      <tbody>
        ${rows.map((r) => {
          const id = Number(r.id || 0);
          const cancelled = String(r.status || "").toLowerCase() === "cancelled";
          const partLabel = r.part_code
            ? `<strong>${String(r.part_code).replace(/</g, "&lt;")}</strong><br><small class="muted">${String(r.part_name || "").replace(/</g, "&lt;")}</small>`
            : String(r.part_name || "").replace(/</g, "&lt;");
          const statusOpts = ["on_order", "in_transit", "arrived", "cancelled"]
            .map((st) => `<option value="${st}"${String(r.status || "").toLowerCase() === st ? " selected" : ""}>${spoStatusLabel(st)}</option>`)
            .join("");
          const inStore = Boolean(r.in_store_inventory || r.stock_movement_id);
          const isArrived = String(r.status || "").toLowerCase() === "arrived";
          const inventoryCell = inStore
            ? `<span class="pill green" title="Received into store stock">In store</span>`
            : isArrived
              ? `<button type="button" class="btn btn-secondary btn-sm" data-spo-receive="${id}">Receive to store</button>`
              : "—";
          const dis = cancelled ? " disabled" : "";
          return `
            <tr data-spo-row="${id}">
              <td>${String(r.order_date || "").replace(/</g, "&lt;")}</td>
              <td>${partLabel}</td>
              <td style="text-align:right;">${Number(r.qty || 0)}</td>
              <td style="text-align:right;">${moneyUsd(r.unit_cost)}</td>
              <td style="text-align:right;"><strong>${moneyUsd(r.line_total)}</strong></td>
              <td>
                <input type="text" class="spo-inline-input w-full" data-spo-field="supplier_name" value="${spoAttrVal(r.supplier_name || "")}" placeholder="Supplier"${dis} />
              </td>
              <td>
                <input type="text" class="spo-inline-input w-full" data-spo-field="po_number" value="${spoAttrVal(r.po_number || "")}" placeholder="PO when issued"${dis} />
              </td>
              <td>
                <input type="text" class="spo-inline-input w-full" data-spo-field="requisition_number" value="${spoAttrVal(r.requisition_number || "")}" placeholder="Req #"${dis} />
              </td>
              <td>
                <input type="date" class="spo-inline-input w-full" data-spo-field="expected_arrival_date" value="${spoAttrVal(r.expected_arrival_date || "")}"${dis} />
              </td>
              <td>
                <select data-spo-status="${id}" class="w-full" style="min-width:120px;"${inStore ? " disabled title=\"Already in store inventory\"" : dis}>${statusOpts}</select>
              </td>
              <td>${inventoryCell}</td>
              <td class="spo-row-actions">
                ${cancelled
                  ? `<span class="muted small">Cancelled</span>`
                  : `<button type="button" class="btn btn-primary btn-sm" data-spo-save="${id}">Save</button>
                     <button type="button" class="btn btn-secondary btn-sm" data-spo-del="${id}">Cancel</button>`}
              </td>
            </tr>
          `;
        }).join("")}
      </tbody>
    </table>
    <p class="muted small" style="margin-top:8px;">Edit Req # or PO on each line, then click <strong>Save</strong>. Change status to Arrived when goods land to post into store inventory.</p>
  `;
}

function readStoresPartOrderRowPatch(rowEl) {
  if (!rowEl) return null;
  const read = (field) => {
    const el = rowEl.querySelector(`[data-spo-field="${field}"]`);
    if (!el) return null;
    const v = String(el.value || "").trim();
    return v || null;
  };
  const statusEl = rowEl.querySelector("select[data-spo-status]");
  const patch = {
    supplier_name: read("supplier_name"),
    po_number: read("po_number"),
    requisition_number: read("requisition_number"),
    expected_arrival_date: read("expected_arrival_date"),
  };
  if (statusEl && !statusEl.disabled) {
    patch.status = String(statusEl.value || "").trim();
  }
  return patch;
}

async function patchStoresPartOrder(id, patch) {
  const n = Number(id || 0);
  if (!n) return null;
  const data = await fetchJson(`${API}/api/stock/part-orders/${n}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(patch || {}),
  });
  handleStoresPartOrderReceipt(data);
  return data;
}

async function saveStoresPartOrderRow(id) {
  const host = qs("spoList");
  const rowEl = host?.querySelector(`tr[data-spo-row="${Number(id)}"]`);
  const patch = readStoresPartOrderRowPatch(rowEl);
  if (!patch) return;
  await patchStoresPartOrder(id, patch);
  setStatus("Purchase line saved.");
  await loadStoresPartOrders();
}

async function loadStoresPartOrders() {
  ensureStoresPartOrderDates();
  const start = (qs("spoDateFrom")?.value || "").trim();
  const end = (qs("spoDateTo")?.value || "").trim();
  const status = (qs("spoFilterStatus")?.value || "").trim();
  if (!start || !end) return alert("Choose period from and to dates.");

  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  if (status) q.set("status", status);

  setStatus("Loading parts purchases...");
  setSkeleton("spoList", 2);
  const data = await fetchJson(`${API}/api/stock/part-orders?${q.toString()}`);
  storesPartOrdersCache = Array.isArray(data.rows) ? data.rows : [];
  renderStoresPartOrdersSummary(data.summary);
  renderStoresPartOrdersTable(storesPartOrdersCache);
  setStatus("Parts purchases loaded.");
}

function clearStoresPartOrderForm() {
  ["spoPartCode", "spoPartName", "spoSupplier", "spoPoNumber", "spoRequisitionNumber", "spoNotes"].forEach((id) => {
    const el = qs(id);
    if (el) el.value = "";
  });
  if (qs("spoQty")) qs("spoQty").value = "1";
  if (qs("spoUnitCost")) qs("spoUnitCost").value = "0";
  if (qs("spoStatus")) qs("spoStatus").value = "on_order";
  if (qs("spoExpectedDate")) qs("spoExpectedDate").value = "";
  ensureStoresPartOrderDates();
  const msg = qs("spoFormMsg");
  if (msg) msg.textContent = "";
}

async function saveStoresPartOrder() {
  ensureStoresPartOrderDates();
  const msg = qs("spoFormMsg");
  const part_code = normalizeStockReportPartInput(qs("spoPartCode")?.value || "");
  const part_name = String(qs("spoPartName")?.value || "").trim();
  const qty = Number(qs("spoQty")?.value || 1);
  const unit_cost = Number(qs("spoUnitCost")?.value || 0);
  const supplier_name = String(qs("spoSupplier")?.value || "").trim();
  const po_number = String(qs("spoPoNumber")?.value || "").trim();
  const requisition_number = String(qs("spoRequisitionNumber")?.value || "").trim();
  const order_date = (qs("spoOrderDate")?.value || "").trim();
  const expected_arrival_date = (qs("spoExpectedDate")?.value || "").trim();
  const status = String(qs("spoStatus")?.value || "on_order").trim();
  const notes = String(qs("spoNotes")?.value || "").trim();

  if (!part_code && !part_name) {
    if (msg) msg.textContent = "Enter a part code or description.";
    return;
  }
  if (!order_date) {
    if (msg) msg.textContent = "Order date is required.";
    return;
  }
  if (!Number.isFinite(qty) || qty <= 0) {
    if (msg) msg.textContent = "Quantity must be greater than zero.";
    return;
  }

  if (msg) msg.textContent = "Saving...";
  const data = await fetchJson(`${API}/api/stock/part-orders`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      part_code,
      part_name,
      qty,
      unit_cost,
      supplier_name,
      po_number: po_number || null,
      requisition_number: requisition_number || null,
      order_date,
      expected_arrival_date: expected_arrival_date || null,
      status,
      notes,
    }),
  });
  handleStoresPartOrderReceipt(data);
  if (msg) msg.textContent = data?.stock_receipt?.received
    ? "Purchase saved and received into store."
    : "Purchase saved.";
  clearStoresPartOrderForm();
  await loadStoresPartOrders();
}

async function updateStoresPartOrderStatus(id, status) {
  const n = Number(id || 0);
  if (!n || !status) return;
  const host = qs("spoList");
  const rowEl = host?.querySelector(`tr[data-spo-row="${n}"]`);
  const patch = readStoresPartOrderRowPatch(rowEl) || {};
  patch.status = status;
  await patchStoresPartOrder(n, patch);
  await loadStoresPartOrders();
}

async function receiveStoresPartOrderToInventory(id) {
  const n = Number(id || 0);
  if (!n) return;
  const data = await fetchJson(`${API}/api/stock/part-orders/${n}/receive`, { method: "POST" });
  handleStoresPartOrderReceipt(data);
  await loadStoresPartOrders();
}

async function cancelStoresPartOrder(id) {
  const n = Number(id || 0);
  if (!n) return;
  if (!confirm("Cancel this purchase line?")) return;
  await fetchJson(`${API}/api/stock/part-orders/${n}`, { method: "DELETE" });
  await loadStoresPartOrders();
}

function buildStoresPartOrdersExportQuery() {
  ensureStoresPartOrderDates();
  const start = (qs("spoDateFrom")?.value || "").trim();
  const end = (qs("spoDateTo")?.value || "").trim();
  const status = (qs("spoFilterStatus")?.value || "").trim();
  if (!start || !end) throw new Error("Choose period from and to dates.");
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  if (status) q.set("status", status);
  return { start, end, q };
}

function openStoresPartOrdersPdf(download = false) {
  const { q } = buildStoresPartOrdersExportQuery();
  if (download) q.set("download", "1");
  openAuthedPdf(`${API}/api/reports/part-orders.pdf?${q.toString()}`).catch((e) =>
    setStatus("Parts purchases PDF error: " + (e.message || e))
  );
}

async function exportStoresPartOrdersXlsx() {
  const { start, end, q } = buildStoresPartOrdersExportQuery();
  setStatus("Generating Excel...");
  const res = await fetch(`${API}/api/reports/part-orders.xlsx?${q.toString()}`, { headers: authHeaders() });
  if (!res.ok) {
    let msg = await res.text().catch(() => "");
    try {
      const j = JSON.parse(msg);
      msg = j.error || j.message || msg;
    } catch {}
    throw new Error(msg || `Export failed (${res.status})`);
  }
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `IRONLOG_Parts_Purchases_${start}_${end}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  setStatus("Parts purchases Excel downloaded.");
}

/* ========== Parts Tracking tab (parts + off-site repairs) ========== */
let ptPartsCache = [];
let ptOffsiteCache = [];

function ptTodayYmd() {
  return new Date().toISOString().slice(0, 10);
}

function ensurePartsTrackingDates() {
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth(), 1);
  const pad = (n) => String(n).padStart(2, "0");
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  if (qs("ptPartsFrom") && !qs("ptPartsFrom").value) qs("ptPartsFrom").value = ymd(start);
  if (qs("ptPartsTo") && !qs("ptPartsTo").value) qs("ptPartsTo").value = ymd(today);
  if (qs("ptOrderDate") && !qs("ptOrderDate").value) qs("ptOrderDate").value = ymd(today);
  if (qs("ptOffSent") && !qs("ptOffSent").value) qs("ptOffSent").value = ymd(today);
}

function ptIsOverdueEta(eta, statusDone) {
  if (statusDone) return false;
  const d = String(eta || "").slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return false;
  return d < ptTodayYmd();
}

function updatePartsTrackingKpis() {
  const partsPending = (ptPartsCache || [])
    .filter((r) => ["on_order", "in_transit"].includes(String(r.status || "").toLowerCase()))
    .reduce((s, r) => s + Number(r.line_total || 0), 0);
  const partsOverdue = (ptPartsCache || []).filter((r) =>
    ptIsOverdueEta(r.expected_arrival_date, String(r.status || "").toLowerCase() === "arrived" || String(r.status || "").toLowerCase() === "cancelled")
  ).length;
  const offEst = (ptOffsiteCache || []).reduce((s, r) => s + Number(r.estimated_cost || 0), 0);
  const offAct = (ptOffsiteCache || []).reduce((s, r) => s + Number(r.actual_cost || 0), 0);
  const offOverdue = (ptOffsiteCache || []).filter((r) =>
    ptIsOverdueEta(r.expected_return_date, String(r.repair_status || "").toLowerCase() === "returned")
  ).length;
  setText("ptPartsPending", moneyUsd(partsPending));
  setText("ptPartsOverdue", String(partsOverdue));
  setText("ptOffsiteEst", moneyUsd(offEst));
  setText("ptOffsiteActual", moneyUsd(offAct));
  setText("ptOffsiteOverdue", String(offOverdue));
}

function clearPtPartsForm() {
  ["ptPartCode", "ptPartName", "ptInvoice", "ptLocation", "ptEta", "ptSupplier", "ptPo", "ptNotes", "ptAsset", "ptWorkOrderId", "ptBreakdownId", "ptOffsiteRepairId", "ptResponsible"].forEach((id) => {
    if (qs(id)) qs(id).value = "";
  });
  if (qs("ptQty")) qs("ptQty").value = "1";
  if (qs("ptUnitCost")) qs("ptUnitCost").value = "0";
  if (qs("ptStatus")) qs("ptStatus").value = "on_order";
  if (qs("ptOrderDate")) qs("ptOrderDate").value = ptTodayYmd();
  const msg = qs("ptPartsMsg");
  if (msg) msg.textContent = "";
}

function renderPtPartsTable(rows) {
  const host = qs("ptPartsList");
  if (!host) return;
  const q = String(qs("ptPartsSearch")?.value || "").trim().toLowerCase();
  let list = Array.isArray(rows) ? rows : [];
  if (q) {
    list = list.filter((r) => {
      const hay = `${r.part_code || ""} ${r.part_name || ""} ${r.invoice_number || ""} ${r.current_location || ""} ${r.supplier_name || ""} ${r.po_number || ""}`.toLowerCase();
      return hay.includes(q);
    });
  }
  if (!list.length) {
    host.innerHTML = `<div class="muted small">No purchase lines for this filter.</div>`;
    return;
  }
  host.innerHTML = `<div class="parts-order-grid">
    ${list.map((r) => {
          const id = Number(r.id || 0);
          const cancelled = String(r.status || "").toLowerCase() === "cancelled";
          const overdue = ptIsOverdueEta(r.expected_arrival_date, cancelled || String(r.status || "").toLowerCase() === "arrived");
          const partLabel = r.part_code
            ? `<strong>${escapeHtml(String(r.part_code))}</strong><br><small class="muted">${escapeHtml(String(r.part_name || ""))}</small>`
            : escapeHtml(String(r.part_name || ""));
          const statusOpts = ["on_order", "in_transit", "arrived", "cancelled"]
            .map((st) => `<option value="${st}"${String(r.status || "").toLowerCase() === st ? " selected" : ""}>${spoStatusLabel(st)}</option>`)
            .join("");
          const dis = cancelled ? " disabled" : "";
          const links = [r.asset_code ? `Asset ${r.asset_code}` : "", r.work_order_id ? `WO #${r.work_order_id}` : "", r.breakdown_id ? `BD #${r.breakdown_id}` : "", r.offsite_repair_id ? `Offsite #${r.offsite_repair_id}` : ""].filter(Boolean);
          return `
            <article data-pt-row="${id}" class="parts-order-card ${overdue ? "is-overdue" : ""}">
              <header><div>${partLabel}<div class="muted small">Qty ${Number(r.qty || 0)} · ${moneyUsd(r.line_total)} total</div></div><div><span class="pill ${overdue ? "red" : "blue"}">${overdue ? "OVERDUE" : spoStatusLabel(r.status)}</span></div></header>
              <div class="parts-order-links">${links.length ? links.map((x) => `<span class="pill">${escapeHtml(x)}</span>`).join(" ") : `<span class="muted small">Not linked to equipment yet</span>`}</div>
              <div class="parts-order-fields">
                <label>Status<select data-pt-status="${id}"${dis}>${statusOpts}</select></label>
                <label>Asset<input data-pt-field="asset_code" list="assetCodeOptions" value="${spoAttrVal(r.asset_code || "")}"${dis} /></label>
                <label>WO #<input type="number" min="1" data-pt-field="work_order_id" value="${spoAttrVal(r.work_order_id || "")}"${dis} /></label>
                <label>Breakdown #<input type="number" min="1" data-pt-field="breakdown_id" value="${spoAttrVal(r.breakdown_id || "")}"${dis} /></label>
                <label>Offsite #<input type="number" min="1" data-pt-field="offsite_repair_id" value="${spoAttrVal(r.offsite_repair_id || "")}"${dis} /></label>
                <label>Responsible<input data-pt-field="responsible_person" value="${spoAttrVal(r.responsible_person || "")}"${dis} /></label>
                <label>Supplier<input data-pt-field="supplier_name" value="${spoAttrVal(r.supplier_name || "")}"${dis} /></label>
                <label>PO #<input data-pt-field="po_number" value="${spoAttrVal(r.po_number || "")}"${dis} /></label>
                <label>Invoice #<input data-pt-field="invoice_number" value="${spoAttrVal(r.invoice_number || "")}"${dis} /></label>
                <label>Location<input data-pt-field="current_location" value="${spoAttrVal(r.current_location || "")}"${dis} /></label>
                <label>Ordered<input type="date" value="${spoAttrVal(r.order_date || "")}" disabled /></label>
                <label>ETA on site<input type="date" data-pt-field="expected_arrival_date" value="${spoAttrVal(r.expected_arrival_date || "")}"${dis} /></label>
              </div>
              <footer><span class="muted small">${r.in_store_inventory ? "Received into Stores inventory" : escapeHtml(r.notes || "No progress note")}</span><div>
                ${cancelled ? `<span class="muted small">Cancelled</span>` : `
                  <button type="button" class="btn btn-primary btn-sm" data-pt-save="${id}">Save</button>
                  ${String(r.status || "").toLowerCase() === "arrived" && !r.in_store_inventory
                    ? `<button type="button" class="btn btn-secondary btn-sm" data-pt-receive="${id}">Receive</button>`
                    : ""}
                  <button type="button" class="btn btn-secondary btn-sm" data-pt-del="${id}">Cancel</button>
                `}
              </div></footer>
            </article>`;
        }).join("")}
    </div>`;
}

function readPtPartsRowPatch(rowEl) {
  if (!rowEl) return null;
  const read = (field) => {
    const el = rowEl.querySelector(`[data-pt-field="${field}"]`);
    if (!el) return null;
    return String(el.value || "").trim() || null;
  };
  const statusEl = rowEl.querySelector("select[data-pt-status]");
  const patch = {
    invoice_number: read("invoice_number"),
    current_location: read("current_location"),
    expected_arrival_date: read("expected_arrival_date"),
    supplier_name: read("supplier_name"),
    po_number: read("po_number"),
    asset_code: read("asset_code"),
    work_order_id: read("work_order_id") ? Number(read("work_order_id")) : null,
    breakdown_id: read("breakdown_id") ? Number(read("breakdown_id")) : null,
    offsite_repair_id: read("offsite_repair_id") ? Number(read("offsite_repair_id")) : null,
    responsible_person: read("responsible_person"),
  };
  if (statusEl && !statusEl.disabled) patch.status = String(statusEl.value || "").trim();
  return patch;
}

async function loadPtPartsOrders() {
  ensurePartsTrackingDates();
  const start = (qs("ptPartsFrom")?.value || "").trim();
  const end = (qs("ptPartsTo")?.value || "").trim();
  const status = (qs("ptPartsStatus")?.value || "").trim();
  const q = new URLSearchParams();
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  if (status) q.set("status", status);
  setSkeleton("ptPartsList", 2);
  const data = await fetchJson(`${API}/api/stock/part-orders?${q.toString()}`);
  ptPartsCache = Array.isArray(data?.rows) ? data.rows : [];
  renderPtPartsTable(ptPartsCache);
  updatePartsTrackingKpis();
}

async function savePtPartsOrder() {
  ensurePartsTrackingDates();
  const msg = qs("ptPartsMsg");
  const body = {
    part_code: normalizeStockReportPartInput(qs("ptPartCode")?.value || ""),
    part_name: String(qs("ptPartName")?.value || "").trim(),
    qty: Number(qs("ptQty")?.value || 1),
    unit_cost: Number(qs("ptUnitCost")?.value || 0),
    invoice_number: String(qs("ptInvoice")?.value || "").trim() || null,
    current_location: String(qs("ptLocation")?.value || "").trim() || null,
    order_date: (qs("ptOrderDate")?.value || "").trim(),
    expected_arrival_date: (qs("ptEta")?.value || "").trim() || null,
    status: String(qs("ptStatus")?.value || "on_order").trim(),
    supplier_name: String(qs("ptSupplier")?.value || "").trim() || null,
    po_number: String(qs("ptPo")?.value || "").trim() || null,
    notes: String(qs("ptNotes")?.value || "").trim() || null,
    asset_code: String(qs("ptAsset")?.value || "").trim() || null,
    work_order_id: Number(qs("ptWorkOrderId")?.value || 0) || null,
    breakdown_id: Number(qs("ptBreakdownId")?.value || 0) || null,
    offsite_repair_id: Number(qs("ptOffsiteRepairId")?.value || 0) || null,
    responsible_person: String(qs("ptResponsible")?.value || "").trim() || null,
  };
  if (!body.part_code && !body.part_name) {
    if (msg) msg.textContent = "Enter a part code or description.";
    return;
  }
  if (!body.order_date) {
    if (msg) msg.textContent = "Order date is required.";
    return;
  }
  if (msg) msg.textContent = "Saving…";
  const data = await fetchJson(`${API}/api/stock/part-orders`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  handleStoresPartOrderReceipt(data);
  if (msg) msg.textContent = "Part line saved.";
  clearPtPartsForm();
  await loadPtPartsOrders();
  loadStoresPartOrders().catch(() => {});
}

function ptOffStatusLabel(s) {
  const v = String(s || "").toLowerCase();
  const map = {
    sent_offsite: "Sent offsite",
    diagnosis: "Diagnosis",
    in_repair: "In repair",
    waiting_parts: "Waiting parts",
    ready_return: "Ready return",
    returned: "Returned",
  };
  return map[v] || v || "—";
}

function clearPtOffForm() {
  ["ptOffAsset", "ptOffAttachment", "ptOffLocation", "ptOffInvoice", "ptOffVendor", "ptOffNotes", "ptOffEta", "ptOffReason", "ptOffResponsible", "ptOffQuote"].forEach((id) => {
    if (qs(id)) qs(id).value = "";
  });
  if (qs("ptOffEstCost")) qs("ptOffEstCost").value = "0";
  if (qs("ptOffActCost")) qs("ptOffActCost").value = "0";
  if (qs("ptOffStatus")) qs("ptOffStatus").value = "sent_offsite";
  if (qs("ptOffApproval")) qs("ptOffApproval").value = "not_required";
  if (qs("ptOffSent")) qs("ptOffSent").value = ptTodayYmd();
  if (qs("ptOffEditId")) qs("ptOffEditId").value = "";
  const msg = qs("ptOffMsg");
  if (msg) msg.textContent = "";
}

function renderPtOffsiteTable(rows) {
  const host = qs("ptOffList");
  if (!host) return;
  const statusFilter = String(qs("ptOffStatusFilter")?.value || "").trim().toLowerCase();
  let list = Array.isArray(rows) ? rows : [];
  if (statusFilter) list = list.filter((r) => String(r.repair_status || "").toLowerCase() === statusFilter);
  if (!list.length) {
    host.innerHTML = `<div class="muted small">No off-site repairs for this filter.</div>`;
    return;
  }
  const approvalLabel = (value) => ({ not_required: "Not required", quote_required: "Quote required", awaiting_approval: "Awaiting approval", approved: "Approved", declined: "Declined" }[value] || value || "Not required");
  host.innerHTML = `<div class="offsite-workflow-grid">
    ${list.map((r) => {
          const id = Number(r.id || 0);
          const overdue = ptIsOverdueEta(r.expected_return_date, String(r.repair_status || "").toLowerCase() === "returned");
          const attach = String(r.attachment_name || "").trim();
          const statusOpts = ["sent_offsite", "diagnosis", "in_repair", "waiting_parts", "ready_return", "returned"]
            .map((st) => `<option value="${st}"${String(r.repair_status || "").toLowerCase() === st ? " selected" : ""}>${ptOffStatusLabel(st)}</option>`)
            .join("");
          const approvalOpts = ["not_required", "quote_required", "awaiting_approval", "approved", "declined"]
            .map((st) => `<option value="${st}"${String(r.approval_status || "not_required").toLowerCase() === st ? " selected" : ""}>${approvalLabel(st)}</option>`)
            .join("");
          return `
            <article data-pt-off-row="${id}" class="offsite-workflow-card ${overdue ? "is-overdue" : ""}">
              <header><div><strong>${escapeHtml(String(r.asset_code || ""))}</strong> · ${escapeHtml(String(r.asset_name || ""))}${attach ? `<div class="muted small">${escapeHtml(attach)}</div>` : ""}</div><span class="pill ${overdue ? "red" : "blue"}">${overdue ? "OVERDUE" : `${Number(r.days_offsite || 0)} days offsite`}</span></header>
              <div class="offsite-workflow-fields">
                <label>Status<select data-pt-off-field="repair_status">${statusOpts}</select></label>
                <label>Approval<select data-pt-off-field="approval_status">${approvalOpts}</select></label>
                <label>Responsible<input data-pt-off-field="responsible_person" value="${spoAttrVal(r.responsible_person || "")}" /></label>
                <label>Vendor<input data-pt-off-field="vendor" value="${spoAttrVal(r.vendor || "")}" /></label>
                <label>Location<input data-pt-off-field="current_location" value="${spoAttrVal(r.current_location || "")}" /></label>
                <label>Sent<input type="date" data-pt-off-field="sent_date" value="${spoAttrVal(r.sent_date || "")}" /></label>
                <label>Expected return<input type="date" data-pt-off-field="expected_return_date" value="${spoAttrVal(r.expected_return_date || "")}" /></label>
                <label>Actual return<input type="date" data-pt-off-field="actual_return_date" value="${spoAttrVal(r.actual_return_date || "")}" /></label>
                <label>Quote #<input data-pt-off-field="quote_number" value="${spoAttrVal(r.quote_number || "")}" /></label>
                <label>Invoice #<input data-pt-off-field="invoice_number" value="${spoAttrVal(r.invoice_number || "")}" /></label>
                <label>Estimated cost<input type="number" min="0" step="0.01" data-pt-off-field="estimated_cost" value="${spoAttrVal(r.estimated_cost ?? "")}" /></label>
                <label>Actual cost<input type="number" min="0" step="0.01" data-pt-off-field="actual_cost" value="${spoAttrVal(r.actual_cost ?? "")}" /></label>
              </div>
              <label class="offsite-wide-field">Repair reason<input data-pt-off-field="repair_reason" value="${spoAttrVal(r.repair_reason || "")}" /></label>
              <label class="offsite-wide-field">Progress / next action<textarea rows="2" data-pt-off-field="notes">${escapeHtml(r.notes || "")}</textarea></label>
                <input type="hidden" data-pt-off-field="attachment_name" value="${spoAttrVal(r.attachment_name || "")}" />
                <input type="hidden" data-pt-off-field="breakdown_id" value="${spoAttrVal(r.breakdown_id || "")}" />
              <footer><span class="muted small">${r.return_confirmed_by ? `Return confirmed by ${escapeHtml(r.return_confirmed_by)}` : `Updated ${escapeHtml(r.updated_at || "—")}`}</span><div><button type="button" class="btn btn-secondary btn-sm" data-pt-off-history="${id}">History</button> <button type="button" class="btn btn-primary btn-sm" data-pt-off-save="${id}">Save progress</button></div></footer>
              <div class="offsite-history" data-pt-off-history-host="${id}" hidden></div>
            </article>`;
        }).join("")}
    </div>`;
}

async function loadPtOffsiteRepairs() {
  ensurePartsTrackingDates();
  const include = qs("ptOffIncludeReturned")?.checked ? "1" : "0";
  setSkeleton("ptOffList", 2);
  const data = await fetchJson(`${API}/api/breakdown-ops/offsite-repairs?include_closed=${include}`);
  ptOffsiteCache = Array.isArray(data?.rows) ? data.rows : [];
  renderPtOffsiteTable(ptOffsiteCache);
  updatePartsTrackingKpis();
}

async function savePtOffsiteRepair() {
  ensurePartsTrackingDates();
  const msg = qs("ptOffMsg");
  const asset_code = String(qs("ptOffAsset")?.value || "").trim();
  if (!asset_code) {
    if (msg) msg.textContent = "Asset code is required.";
    return;
  }
  const sent_date = (qs("ptOffSent")?.value || "").trim();
  if (!sent_date) {
    if (msg) msg.textContent = "Date sent is required.";
    return;
  }
  const body = {
    asset_code,
    attachment_name: String(qs("ptOffAttachment")?.value || "").trim() || null,
    repair_status: String(qs("ptOffStatus")?.value || "sent_offsite").trim(),
    sent_date,
    expected_return_date: (qs("ptOffEta")?.value || "").trim() || null,
    current_location: String(qs("ptOffLocation")?.value || "").trim() || null,
    invoice_number: String(qs("ptOffInvoice")?.value || "").trim() || null,
    vendor: String(qs("ptOffVendor")?.value || "").trim() || null,
    estimated_cost: Number(qs("ptOffEstCost")?.value || 0) || 0,
    actual_cost: Number(qs("ptOffActCost")?.value || 0) || 0,
    notes: String(qs("ptOffNotes")?.value || "").trim() || null,
    repair_reason: String(qs("ptOffReason")?.value || "").trim() || null,
    responsible_person: String(qs("ptOffResponsible")?.value || "").trim() || null,
    approval_status: String(qs("ptOffApproval")?.value || "not_required").trim(),
    quote_number: String(qs("ptOffQuote")?.value || "").trim() || null,
  };
  if (!body.repair_reason) {
    if (msg) msg.textContent = "Repair reason is required.";
    return;
  }
  if (msg) msg.textContent = "Saving…";
  await fetchJson(`${API}/api/breakdown-ops/offsite-repairs`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (msg) msg.textContent = "Off-site repair saved.";
  clearPtOffForm();
  await loadPtOffsiteRepairs();
}

function readPtOffRowPatch(rowEl) {
  if (!rowEl) return null;
  const read = (field) => {
    const el = rowEl.querySelector(`[data-pt-off-field="${field}"]`);
    if (!el) return null;
    return String(el.value || "").trim();
  };
  const estimated_cost = read("estimated_cost");
  const actual_cost = read("actual_cost");
  const breakdownRaw = read("breakdown_id");
  return {
    repair_status: read("repair_status") || "sent_offsite",
    vendor: read("vendor") || null,
    invoice_number: read("invoice_number") || null,
    sent_date: read("sent_date"),
    current_location: read("current_location") || null,
    expected_return_date: read("expected_return_date") || null,
    actual_return_date: read("actual_return_date") || null,
    estimated_cost: estimated_cost === "" ? null : Number(estimated_cost),
    actual_cost: actual_cost === "" ? null : Number(actual_cost),
    notes: read("notes") || null,
    repair_reason: read("repair_reason") || null,
    responsible_person: read("responsible_person") || null,
    approval_status: read("approval_status") || "not_required",
    quote_number: read("quote_number") || null,
    attachment_name: read("attachment_name") || null,
    breakdown_id: breakdownRaw ? Number(breakdownRaw) : null,
  };
}

async function savePtOffsiteRow(id) {
  const host = qs("ptOffList");
  const rowEl = host?.querySelector(`[data-pt-off-row="${Number(id)}"]`);
  const patch = readPtOffRowPatch(rowEl);
  if (!patch || !patch.sent_date) {
    alert("Sent date is required.");
    return;
  }
  await fetchJson(`${API}/api/breakdown-ops/offsite-repairs/${Number(id)}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(patch),
  });
  setStatus("Off-site repair updated.");
  await loadPtOffsiteRepairs();
}

async function loadPartsTrackingTab() {
  ensurePartsTrackingDates();
  await Promise.all([
    loadPtPartsOrders().catch((e) => setStatus("Parts tracking error: " + (e.message || e))),
    loadPtOffsiteRepairs().catch((e) => setStatus("Off-site tracking error: " + (e.message || e))),
  ]);
}

/** Start-up: Stock monitor, stock reports and stores part order controls. Called once from init() in init.js. */
function wireStockControls() {
  qs("loadStockMonitor")?.addEventListener("click", () =>
    loadStockMonitor().catch((e) => setStatus("Stock monitor error: " + e.message))
  );
  qs("spLoad")?.addEventListener("click", () =>
    loadStockOnHandPage().catch((e) => setStatus("Stock on hand error: " + e.message))
  );
  ensureGmStockReportDate();
  updateGmStockReportHelp();
  qs("gmStockReportPeriod")?.addEventListener("change", updateGmStockReportHelp);
  qs("downloadGmStockReportXlsx")?.addEventListener("click", () => downloadGmStockReportXlsx());
  qs("spSort")?.addEventListener("change", () => refreshStockInventoryDisplay());
  qs("spOnlyLow")?.addEventListener("change", () => refreshStockInventoryDisplay());
  qs("spCategory")?.addEventListener("change", () => refreshStockInventoryDisplay());
  qs("spList")?.addEventListener("change", (evt) => {
    const select = evt.target?.closest?.("select[data-stock-category-select]");
    if (select) saveStockCategory(select);
  });
  qs("spList")?.addEventListener("click", (evt) => {
    const btn = evt.target?.closest?.("button[data-stock-action]");
    if (!btn) return;
    openStockAction(btn.closest("[data-stock-code]"), btn.getAttribute("data-stock-action"));
  });
  qs("spFilter")?.addEventListener("input", () => {
    if (stockPageData.rows.length) refreshStockInventoryDisplay();
  });
  qs("spExportCsv")?.addEventListener("click", exportStockOnHandCsv);
  qs("spOpenPdf")?.addEventListener("click", openStockOnHandPdf);
  qs("smrLoad")?.addEventListener("click", () =>
    loadStockMovementsReport().catch((e) => setStatus("Stock movements report error: " + e.message))
  );
  qs("smrExportCsv")?.addEventListener("click", exportStockMovementsReportCsv);
  qs("smrOpenPdf")?.addEventListener("click", openStockMovementsReportPdf);
  qs("spoLoad")?.addEventListener("click", () =>
    loadStoresPartOrders().catch((e) => setStatus("Parts purchases error: " + e.message))
  );
  qs("spoSave")?.addEventListener("click", () =>
    saveStoresPartOrder().catch((e) => {
      const msg = qs("spoFormMsg");
      if (msg) msg.textContent = e.message || String(e);
    })
  );
  qs("spoClear")?.addEventListener("click", clearStoresPartOrderForm);
  qs("spoFilterStatus")?.addEventListener("change", () => {
    loadStoresPartOrders().catch(() => {});
  });
  qs("spoList")?.addEventListener("change", (evt) => {
    const sel = evt.target?.closest?.("select[data-spo-status]");
    if (!sel) return;
    const id = Number(sel.getAttribute("data-spo-status") || 0);
    const status = String(sel.value || "").trim();
    if (!id) return;
    updateStoresPartOrderStatus(id, status).catch((e) => alert(e.message || String(e)));
  });
  qs("spoList")?.addEventListener("click", (evt) => {
    const saveBtn = evt.target?.closest?.("button[data-spo-save]");
    if (saveBtn) {
      saveStoresPartOrderRow(Number(saveBtn.getAttribute("data-spo-save") || 0)).catch((e) =>
        alert(e.message || String(e))
      );
      return;
    }
    const recvBtn = evt.target?.closest?.("button[data-spo-receive]");
    if (recvBtn) {
      receiveStoresPartOrderToInventory(Number(recvBtn.getAttribute("data-spo-receive") || 0)).catch((e) =>
        alert(e.message || String(e))
      );
      return;
    }
    const btn = evt.target?.closest?.("button[data-spo-del]");
    if (!btn) return;
    cancelStoresPartOrder(Number(btn.getAttribute("data-spo-del") || 0)).catch((e) => alert(e.message || String(e)));
  });
  qs("spoExportXlsx")?.addEventListener("click", () =>
    exportStoresPartOrdersXlsx().catch((e) => setStatus("Parts purchases Excel error: " + e.message))
  );
  qs("spoOpenPdf")?.addEventListener("click", () => openStoresPartOrdersPdf(false));
  qs("spoDownloadPdf")?.addEventListener("click", () => openStoresPartOrdersPdf(true));

  qs("ptPartsLoad")?.addEventListener("click", () =>
    loadPtPartsOrders().catch((e) => setStatus("Parts tracking error: " + e.message))
  );
  qs("ptPartsSave")?.addEventListener("click", () =>
    savePtPartsOrder().catch((e) => {
      const msg = qs("ptPartsMsg");
      if (msg) msg.textContent = e.message || String(e);
    })
  );
  qs("ptPartsClear")?.addEventListener("click", clearPtPartsForm);
  qs("ptPartsStatus")?.addEventListener("change", () => loadPtPartsOrders().catch(() => {}));
  qs("ptPartsSearch")?.addEventListener("input", () => renderPtPartsTable(ptPartsCache));
  qs("ptPartsList")?.addEventListener("change", (evt) => {
    const sel = evt.target?.closest?.("select[data-pt-status]");
    if (!sel) return;
    const id = Number(sel.getAttribute("data-pt-status") || 0);
    if (!id) return;
    const rowEl = qs("ptPartsList")?.querySelector(`[data-pt-row="${id}"]`);
    const patch = readPtPartsRowPatch(rowEl) || {};
    patch.status = String(sel.value || "").trim();
    patchStoresPartOrder(id, patch)
      .then(() => loadPtPartsOrders())
      .catch((e) => alert(e.message || String(e)));
  });
  qs("ptPartsList")?.addEventListener("click", (evt) => {
    const saveBtn = evt.target?.closest?.("button[data-pt-save]");
    if (saveBtn) {
      const id = Number(saveBtn.getAttribute("data-pt-save") || 0);
      const rowEl = qs("ptPartsList")?.querySelector(`[data-pt-row="${id}"]`);
      const patch = readPtPartsRowPatch(rowEl);
      if (!patch) return;
      patchStoresPartOrder(id, patch)
        .then(() => {
          setStatus("Part line saved.");
          return loadPtPartsOrders();
        })
        .catch((e) => alert(e.message || String(e)));
      return;
    }
    const recvBtn = evt.target?.closest?.("button[data-pt-receive]");
    if (recvBtn) {
      receiveStoresPartOrderToInventory(Number(recvBtn.getAttribute("data-pt-receive") || 0))
        .then(() => loadPtPartsOrders())
        .catch((e) => alert(e.message || String(e)));
      return;
    }
    const delBtn = evt.target?.closest?.("button[data-pt-del]");
    if (delBtn) {
      cancelStoresPartOrder(Number(delBtn.getAttribute("data-pt-del") || 0))
        .then(() => loadPtPartsOrders())
        .catch((e) => alert(e.message || String(e)));
    }
  });
  qs("ptOffLoad")?.addEventListener("click", () =>
    loadPtOffsiteRepairs().catch((e) => setStatus("Off-site tracking error: " + e.message))
  );
  qs("ptOffSave")?.addEventListener("click", () =>
    savePtOffsiteRepair().catch((e) => {
      const msg = qs("ptOffMsg");
      if (msg) msg.textContent = e.message || String(e);
    })
  );
  qs("ptOffClear")?.addEventListener("click", clearPtOffForm);
  qs("ptOffStatusFilter")?.addEventListener("change", () => renderPtOffsiteTable(ptOffsiteCache));
  qs("ptOffIncludeReturned")?.addEventListener("change", () => loadPtOffsiteRepairs().catch(() => {}));
  qs("ptOffList")?.addEventListener("click", (evt) => {
    const historyBtn = evt.target?.closest?.("button[data-pt-off-history]");
    if (historyBtn) {
      togglePtOffsiteHistory(Number(historyBtn.getAttribute("data-pt-off-history") || 0));
      return;
    }
    const btn = evt.target?.closest?.("button[data-pt-off-save]");
    if (!btn) return;
    savePtOffsiteRow(Number(btn.getAttribute("data-pt-off-save") || 0)).catch((e) =>
      alert(e.message || String(e))
    );
  });
}
