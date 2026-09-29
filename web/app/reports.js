// IRONLOG/web/app/reports.js — Report downloads (daily, weekly, GM, cost, lube, stock, operations).
// Part of the main app; index.html loads these files in order and they share one global scope.

function getLast7Range(endDate) {
  const end = new Date(endDate + "T00:00:00");
  const start = new Date(end);
  start.setDate(start.getDate() - 6);
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

async function syncFamsFuelSelectedDates() {
  const range = selectedFamsFuelSyncRange();
  if (!range) return;
  return syncFamsFuelNow(range);
}

async function previewFamsFuelDuplicates() {
  const resultEl = qs("fuelFamsResult");
  const range = selectedFamsFuelSyncRange();
  if (!range) return;
  setStatus(`Checking FAMS duplicates for ${range.startDate} to ${range.endDate}…`);
  const res = await fetchJson(`${API}/api/dashboard/fuel/fams/duplicates/preview`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({ start_date: range.startDate, end_date: range.endDate }),
  });
  if (resultEl) resultEl.textContent = JSON.stringify(res, null, 2);
  setStatus(`${Number(res.found || 0)} exact FAMS duplicate pair(s) found. Review the result, then remove confirmed duplicates if needed.`);
}

async function removeFamsFuelDuplicates() {
  const resultEl = qs("fuelFamsResult");
  const range = selectedFamsFuelSyncRange();
  if (!range) return;
  const approved = confirm(
    `Remove only exact, confirmed FAMS duplicate pairs from ${range.startDate} to ${range.endDate}? The FAMS-tagged transaction will be kept.`,
  );
  if (!approved) return;
  setStatus(`Removing confirmed FAMS duplicates for ${range.startDate} to ${range.endDate}…`);
  const res = await fetchJson(`${API}/api/dashboard/fuel/fams/duplicates/remove`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({ start_date: range.startDate, end_date: range.endDate }),
  });
  if (resultEl) resultEl.textContent = JSON.stringify(res, null, 2);
  setStatus(`${Number(res.removed || 0)} confirmed FAMS duplicate row(s) removed.`);
  await loadDashboard().catch(() => {});
}

function defaultAmlWeeklyEndDate(date = new Date()) {
  const result = new Date(date);
  const daysSinceFriday = (result.getDay() + 2) % 7;
  result.setDate(result.getDate() - daysSinceFriday);
  return result.toISOString().slice(0, 10);
}

async function downloadAmlWeeklyCheckSheet() {
  const weekEnding = String(qs("amlWeeklyEnd")?.value || "").trim();
  const status = qs("amlWeeklyExportStatus");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(weekEnding)) {
    alert("Select the AML week-ending date first.");
    return;
  }
  if (status) status.textContent = "Preparing the protected AML weekly template…";
  setStatus("Preparing AML Weekly Check Sheet…");
  const downloaded = await downloadAuthedFile(
    `${API}/api/reports/aml-weekly-check-sheet.xlsx?week_ending=${encodeURIComponent(weekEnding)}`,
    `AML_Weekly_Check_Sheet_${weekEnding}.xlsx`,
  );
  if (status) {
    status.textContent = downloaded
      ? "Downloaded. Review the completed fields, then submit the sheet as normal."
      : "The AML weekly sheet could not be downloaded. Please retry or contact your administrator.";
  }
  if (downloaded) setStatus("AML Weekly Check Sheet downloaded.");
}

function getLastNDaysRange(endDate, days) {
  const end = new Date(`${endDate}T00:00:00`);
  const span = Math.max(1, Number(days || 30));
  const start = new Date(end);
  start.setDate(start.getDate() - (span - 1));
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

function openDailyXlsx() {
  const date = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const scheduled = qs("scheduled")?.value || 10;
  const ts = Date.now();
  return downloadAuthedFile(
    `${API}/api/reports/daily.xlsx?date=${date}&scheduled=${scheduled}&_ts=${ts}`,
    `IRONLOG_Daily_${date}.xlsx`,
  );
}

/** GM weekly pack: Maintenance & Engineering KPIs (same date field as daily / weekly PDF). */
function openGmWeeklyXlsx() {
  const end = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const scheduled = qs("scheduled")?.value || 10;
  return downloadAuthedFile(
    `${API}/api/reports/gm-weekly.xlsx?end=${encodeURIComponent(end)}&forecast_days=30&scheduled=${scheduled}`,
    `IRONLOG_GM_Weekly_${end}.xlsx`,
  );
}

function downloadCostMonthlyXlsx() {
  const month = (qs("costMonth")?.value || "").trim();
  if (!month) {
    alert("Select a month first.");
    return;
  }
  downloadAuthedFile(`${API}/api/reports/cost-monthly.xlsx?month=${encodeURIComponent(month)}`, `fleet-cost-${month}.xlsx`);
}

function monthlyFleetCostPdfUrl(download = false) {
  const month = (qs("costMonth")?.value || "").trim();
  if (!month) return null;
  const scheduled = qs("scheduled")?.value || 10;
  const q = new URLSearchParams({ month, scheduled: String(scheduled) });
  if (download) q.set("download", "1");
  return `${API}/api/reports/monthly.pdf?${q.toString()}`;
}

async function openMonthlyFleetCostPdf() {
  const url = monthlyFleetCostPdfUrl(false);
  if (!url) return alert("Select a cost month first.");
  try { await openAuthedPdf(url); } catch (err) { alert(`Could not open Monthly Fleet Cost PDF: ${err.message || err}`); }
}

async function togglePtOffsiteHistory(id) {
  const host = qs("ptOffList")?.querySelector(`[data-pt-off-history-host="${Number(id)}"]`);
  if (!host) return;
  if (!host.hidden) { host.hidden = true; return; }
  host.hidden = false;
  host.innerHTML = `<div class="muted small">Loading history…</div>`;
  try {
    const data = await fetchJson(`${API}/api/breakdown-ops/offsite-repairs/${Number(id)}/history`);
    const rows = Array.isArray(data?.rows) ? data.rows : [];
    host.innerHTML = rows.length
      ? rows.map((r) => `<div class="offsite-history-row"><strong>${escapeHtml(ptOffStatusLabel(r.repair_status))}</strong><span>${escapeHtml(r.approval_status || "")}</span><span>${escapeHtml(r.changed_by || "user")} · ${escapeHtml(r.changed_at || "")}</span><div>${escapeHtml(r.notes || "No comment")}</div></div>`).join("")
      : `<div class="muted small">No progress history recorded yet.</div>`;
  } catch (e) {
    host.innerHTML = `<div class="message-error">${escapeHtml(e.message || String(e))}</div>`;
  }
}

function downloadMonthlyFleetCostPdf() {
  const url = monthlyFleetCostPdfUrl(true);
  if (!url) return alert("Select a cost month first.");
  downloadAuthedFile(url, "monthly-fleet-cost.pdf");
}

function downloadMaintenanceCostByEquipmentXlsx() {
  const month = (qs("costMonth")?.value || "").trim();
  const start = (qs("maintCostStart")?.value || "").trim();
  const end = (qs("maintCostEnd")?.value || "").trim();
  if (!month && (!start || !end)) {
    alert("Select a month or a start/end range first.");
    return;
  }
  const q = month
    ? `month=${encodeURIComponent(month)}`
    : `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  downloadAuthedFile(`${API}/api/reports/maintenance-cost-by-equipment.xlsx?${q}`, "maintenance-cost-by-equipment.xlsx");
}

function downloadMtdOpeningHoursXlsx() {
  const month = (qs("mtdOpeningMonth")?.value || "").trim();
  const q = month ? `?month=${encodeURIComponent(month)}` : "";
  downloadAuthedFile(`${API}/api/reports/mtd-opening-hours.xlsx${q}`, "mtd-opening-hours.xlsx");
}

async function openMaintenanceCostByEquipmentPdf(download = false) {
  const month = (qs("costMonth")?.value || "").trim();
  const start = (qs("maintCostStart")?.value || "").trim();
  const end = (qs("maintCostEnd")?.value || "").trim();
  if (!month && (!start || !end)) {
    alert("Select a month or a start/end range first.");
    return;
  }
  const qBase = month
    ? `month=${encodeURIComponent(month)}`
    : `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  const q = `${qBase}${download ? "&download=1" : ""}`;
  const url = `${API}/api/reports/maintenance-cost-by-equipment.pdf?${q}`;
  if (download) downloadAuthedFile(url, "maintenance-cost-by-equipment.pdf");
  else {
    try { await openAuthedPdf(url); } catch (err) { alert(`Could not open Maintenance Cost PDF: ${err.message || err}`); }
  }
}

function downloadMaintenanceExecutivePptx() {
  const month = (qs("costMonth")?.value || "").trim();
  const start = (qs("maintCostStart")?.value || "").trim();
  const end = (qs("maintCostEnd")?.value || "").trim();
  if (!month && (!start || !end)) {
    alert("Select a month or a start/end range first.");
    return;
  }
  const qCore = month
    ? `month=${encodeURIComponent(month)}`
    : `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  const site = encodeURIComponent(getSessionSite());
  const q = `${qCore}&site_code=${site}`;
  downloadAuthedFile(`${API}/api/reports/maintenance-exec.pptx?${q}`, "maintenance-executive.pptx");
}

function downloadGMUpcomingCostsPptx() {
  const month = (qs("costMonth")?.value || "").trim();
  const start = (qs("maintCostStart")?.value || "").trim();
  const end = (qs("maintCostEnd")?.value || "").trim();
  if (!month && (!start || !end)) {
    alert("Select a month or a start/end range first.");
    return;
  }
  const qCore = month
    ? `month=${encodeURIComponent(month)}`
    : `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  const site = encodeURIComponent(getSessionSite());
  const q = `${qCore}&site_code=${site}`;
  downloadAuthedFile(`${API}/api/reports/gm-upcoming-costs.pptx?${q}`, "gm-upcoming-costs.pptx");
}

function downloadGMBudgetMeetingDocx() {
  const month = (qs("costMonth")?.value || qs("plantHireBudgetMonth")?.value || "").trim();
  if (!month) {
    alert("Select a month on Reports (cost month) or Assets → Plant Hire budget month.");
    return;
  }
  const site = encodeURIComponent(getSessionSite() || "main");
  const ts = Date.now();
  downloadAuthedFile(
    `${API}/api/reports/gm-budget-meeting.docx?month=${encodeURIComponent(month)}&site_code=${site}&_ts=${ts}`,
    `gm-budget-meeting-${month}.docx`,
  );
  setStatus("Budget meeting Word export started.");
}

async function saveRainDay() {
  const rainDate = (qs("rainDayDate")?.value || "").trim();
  if (!rainDate) return alert("Pick a rain day first.");
  setStatus("Saving rain day...");
  try {
    const res = await fetchJson(`${API}/api/reports/rain-days`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ date: rainDate }),
    });
    setText("rainDaysResult", JSON.stringify(res, null, 2));
    setStatus("Rain day saved.");
    loadRainDays({ silentNoRange: true }).catch(() => {});
  } catch (e) {
    setText("rainDaysResult", String(e.message || e));
    setStatus("Save rain day failed.");
  }
}

async function removeRainDay() {
  const rainDate = (qs("rainDayDate")?.value || "").trim();
  if (!rainDate) return alert("Pick a rain day first.");
  setStatus("Removing rain day...");
  try {
    const res = await fetchJson(`${API}/api/reports/rain-days/${encodeURIComponent(rainDate)}`, {
      method: "DELETE",
    });
    setText("rainDaysResult", JSON.stringify(res, null, 2));
    setStatus("Rain day removed.");
    loadRainDays({ silentNoRange: true }).catch(() => {});
  } catch (e) {
    setText("rainDaysResult", String(e.message || e));
    setStatus("Remove rain day failed.");
  }
}

function getRainDaysRangeQuery() {
  const month = (qs("costMonth")?.value || "").trim();
  const start = (qs("maintCostStart")?.value || "").trim();
  const end = (qs("maintCostEnd")?.value || "").trim();
  let q = "";
  if (month) {
    const d = new Date(`${month}-01T00:00:00`);
    const y = d.getFullYear();
    const m = d.getMonth();
    const startMonth = `${month}-01`;
    const endMonth = new Date(y, m + 1, 0).toISOString().slice(0, 10);
    q = `start=${encodeURIComponent(startMonth)}&end=${encodeURIComponent(endMonth)}`;
  } else if (start && end) {
    q = `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  }
  return q;
}

function renderRainDaysWidget(res) {
  const widget = qs("rainDaysWidget");
  if (!widget) return;
  const rows = Array.isArray(res?.rows) ? res.rows : [];
  const start = String(res?.start || "");
  const end = String(res?.end || "");
  if (!rows.length) {
    widget.innerHTML = `<div class="muted">No rain days recorded for ${start || "selected"}${end ? ` to ${end}` : ""}.</div>`;
    return;
  }
  widget.innerHTML = `
    <div><strong>Rain days (${rows.length})</strong> <span class="muted">${start} to ${end}</span></div>
    <div style="display:flex; flex-wrap:wrap; gap:8px; margin-top:8px;">
      ${rows.map((r) => `<span class="pill blue">🌧️ ${escapeHtml(String(r.rain_date || ""))}</span>`).join("")}
    </div>
    ${rows.some((r) => String(r.notes || "").trim())
      ? `<div class="muted" style="margin-top:8px;">Notes: ${rows
          .filter((r) => String(r.notes || "").trim())
          .map((r) => `${escapeHtml(String(r.rain_date || ""))} (${escapeHtml(String(r.notes || ""))})`)
          .join(" | ")}</div>`
      : ""}
  `;
}

async function loadRainDays(opts = {}) {
  const q = getRainDaysRangeQuery();
  if (!q) {
    if (!opts.silentNoRange) alert("Select month or start/end first.");
    return;
  }
  setStatus("Loading rain days...");
  try {
    const res = await fetchJson(`${API}/api/reports/rain-days?${q}`);
    setText("rainDaysResult", JSON.stringify(res, null, 2));
    renderRainDaysWidget(res);
    setStatus("Rain days loaded.");
  } catch (e) {
    setText("rainDaysResult", String(e.message || e));
    renderRainDaysWidget({ rows: [], start: "", end: "" });
    setStatus("Load rain days failed.");
  }
}

async function openDailyPdf() {
  const date = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const scheduled = qs("scheduled")?.value || 10;
  const site = encodeURIComponent(getSessionSite());
  const ts = Date.now();
  try {
    await openAuthedPdf(
      `${API}/api/reports/daily.pdf?date=${encodeURIComponent(date)}&scheduled=${encodeURIComponent(scheduled)}&site_code=${site}&_ts=${ts}`,
    );
  } catch (err) {
    alert(`Could not open Daily PDF: ${err.message || err}`);
  }
}

async function openWeeklyPdf() {
  const date = qs("date")?.value || new Date().toISOString().slice(0, 10);
  const scheduled = qs("scheduled")?.value || 10;
  const r = getLast7Range(date);
  try {
    await openAuthedPdf(`${API}/api/reports/weekly.pdf?start=${r.start}&end=${r.end}&scheduled=${scheduled}`);
  } catch (err) {
    alert(`Could not open Weekly PDF: ${err.message || err}`);
  }
}

async function openLubePdf() {
  const start = qs("lubeStart")?.value || "";
  const end = qs("lubeEnd")?.value || "";
  if (!start || !end) return alert("Select lube period first.");
  try {
    await openAuthedPdf(`${API}/api/reports/lube.pdf?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`);
  } catch (err) {
    alert(`Could not open Lube PDF: ${err.message || err}`);
  }
}

async function downloadLubeUsageXlsx() {
  const start = qs("lubeStart")?.value || "";
  const end = qs("lubeEnd")?.value || "";
  if (!start || !end) return alert("Select lube period first.");
  const monthRaw = String(qs("lubeStockMonth")?.value || "").trim();
  const month = /^\d{4}-\d{2}$/.test(monthRaw) ? monthRaw : start.slice(0, 7);
  const loc = String(qs("lubeStockLoc")?.value || "").trim().toUpperCase();
  const q = new URLSearchParams({ start, end, month });
  if (loc) q.set("location_code", loc);
  setStatus("Preparing lube usage Excel export…");
  const downloaded = await downloadAuthedFile(
    `${API}/api/reports/lube-usage-by-asset.xlsx?${q.toString()}`,
    `IRONLOG_Lube_Usage_${start}_to_${end}.xlsx`,
  );
  if (downloaded) setStatus("Lube usage Excel downloaded.");
}

function renderLubeMonthStockTable(data) {
  const el = qs("lubeMonthStockList");
  if (!el) return;
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  const loc = data?.location_code ? ` · ${escapeHtml(String(data.location_code))}` : "";
  const head = `<div class="muted mini" style="margin-bottom:6px;">Store balances${loc} — opening as-of <strong>${escapeHtml(String(data?.opening_as_of || ""))}</strong>, closing <strong>${escapeHtml(String(data?.closing_as_of || ""))}</strong> (${escapeHtml(String(data?.month || ""))})</div>`;
  if (!rows.length) {
    el.innerHTML = head + "<small>No lube stock rows returned.</small>";
    return;
  }
  const body = rows
    .map(
      (r) =>
        `<tr><td>${escapeHtml(String(r.part_code || ""))}</td><td>${escapeHtml(String(r.part_name || ""))}</td>` +
        `<td class="num">${Number(r.min_stock || 0).toFixed(1)}</td>` +
        `<td class="num">${Number(r.opening_qty || 0).toFixed(2)}</td>` +
        `<td class="num">${Number(r.closing_qty || 0).toFixed(2)}</td>` +
        `<td class="num">${Number(r.net_month_movement || 0).toFixed(2)}</td></tr>`
    )
    .join("");
  el.innerHTML =
    head +
    `<div style="overflow:auto;"><table class="gridTable" style="min-width:720px;"><thead><tr>` +
    `<th>Stock code</th><th>Description</th><th>Min</th><th>Opening</th><th>Closing</th><th>Net (month)</th>` +
    `</tr></thead><tbody>${body}</tbody></table></div>`;
}

async function loadLubeMonthStock() {
  const monthRaw = String(qs("lubeStockMonth")?.value || "").trim();
  const month = /^\d{4}-\d{2}$/.test(monthRaw) ? monthRaw : "";
  if (!month) {
    alert("Pick a month for store opening / closing balances.");
    return;
  }
  const loc = String(qs("lubeStockLoc")?.value || "").trim().toUpperCase();
  const q = new URLSearchParams({ month });
  if (loc) q.set("location_code", loc);
  setStatus("Loading lube month stock...");
  try {
    const data = await fetchJson(`${API}/api/stock/lube-month-stock?${q}`);
    renderLubeMonthStockTable(data);
    setStatus("Lube month stock ready.");
  } catch (e) {
    setStatus("Lube month stock failed: " + (e.message || e));
    renderLubeMonthStockTable({ rows: [], opening_as_of: "", closing_as_of: "", month: "" });
  }
}

async function openStockMonitorPdf() {
  const filter = (qs("stockPartFilter")?.value || "").trim();
  const q = filter ? `?part_code=${encodeURIComponent(filter)}` : "";
  try { await openAuthedPdf(`${API}/api/reports/stock-monitor.pdf${q}`); } catch (err) { alert(`Could not open Stock PDF: ${err.message || err}`); }
}

function downloadStockMonitorPdf() {
  const filter = (qs("stockPartFilter")?.value || "").trim();
  const q = filter
    ? `?part_code=${encodeURIComponent(filter)}&download=1`
    : "?download=1";
  downloadAuthedFile(`${API}/api/reports/stock-monitor.pdf${q}`, "stock-monitor.pdf");
}

async function downloadAssetHistoryPdf() {
  const asset_code = getSelectedAssetCode();
  if (!asset_code) {
    alert("Select a fleet card first.");
    return;
  }
  const start = qs("histStart")?.value || "";
  const end = qs("histEnd")?.value || "";
  const url =
    `${API}/api/reports/asset-history/${encodeURIComponent(asset_code)}.pdf` +
    `?start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}&download=1`;
  const filename = `machine-history-${asset_code}-${start || "all"}-to-${end || "today"}.pdf`;
  setStatus(`Preparing machine history PDF for ${asset_code}...`);
  try {
    await downloadAuthedFile(url, filename);
    setStatus(`Machine history PDF downloaded for ${asset_code}.`);
  } catch (err) {
    setStatus("Machine history PDF failed: " + (err.message || err));
    alert(`Could not download machine history PDF: ${err.message || err}`);
  }
}

function openOperationsPdf(download = false) {
  const start = (qs("opFrom")?.value || "").trim();
  const end = (qs("opTo")?.value || "").trim();
  if (!start || !end) {
    alert("Select operations date range first.");
    return;
  }
  const q = `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}${download ? "&download=1" : ""}`;
  return openAuthedReport(`${API}/api/reports/operations.pdf?${q}`, {
    download,
    filename: `IRONLOG_Operations_${start}_to_${end}.pdf`,
  });
}

function downloadOperationsXlsx() {
  const start = (qs("opFrom")?.value || "").trim();
  const end = (qs("opTo")?.value || "").trim();
  if (!start || !end) {
    alert("Select operations date range first.");
    return;
  }
  const q = `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  return downloadAuthedFile(`${API}/api/reports/operations.xlsx?${q}`, `IRONLOG_Operations_${start}_to_${end}.xlsx`);
}

/* =========================
   ACTIONS
========================= */
