// IRONLOG/web/app/finance.js — Finance.
// Part of the main app; index.html loads these files in order and they share one global scope.

let financeLastRunId = null;
let financeLastForecastBatchId = null;
let financeSiteAllocCache = null;

function fmtMoney(n) {
  const v = Number(n || 0);
  return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function currentFinancePeriod() {
  return new Date().toISOString().slice(0, 7);
}

function monthBoundsFromPeriod(period) {
  const p = String(period || "").trim();
  if (!/^\d{4}-\d{2}$/.test(p)) return { start: "", end: "" };
  const [y, m] = p.split("-").map((x) => Number(x));
  const start = `${p}-01`;
  const end = new Date(Date.UTC(y, m, 0)).toISOString().slice(0, 10);
  return { start, end };
}

function initFinanceTab() {
  const period = currentFinancePeriod();
  const { start, end } = monthBoundsFromPeriod(period);
  const setIfEmpty = (id, val) => {
    const el = qs(id);
    if (el && !String(el.value || "").trim()) el.value = val;
  };
  setIfEmpty("finMasterPeriod", period);
  setIfEmpty("finBvaPeriod", period);
  setIfEmpty("finPeriodInput", period);
  setIfEmpty("finSsotPeriod", period);
  setIfEmpty("finForecastStart", period);
  setIfEmpty("finRunStart", start);
  setIfEmpty("finRunEnd", end);
  financeSyncPeriodFields(false);
  loadFinanceSiteFilterFromEntity().catch(() => {});
}

async function loadFinanceSiteFilterFromEntity() {
  const sel = qs("finSiteFilter");
  if (!sel) return;
  try {
    const res = await fetchJson(`${API}/api/entity/sites`);
    const rows = Array.isArray(res.rows) ? res.rows : [];
    if (!rows.length) return;
    const prev = String(sel.value || "");
    sel.innerHTML = `<option value="">All sites</option>${rows
      .map((r) => {
        const code = String(r.site_code || "").trim();
        const name = String(r.site_name || code).trim();
        return code ? `<option value="${escapeHtml(code)}">${escapeHtml(name)} (${escapeHtml(code)})</option>` : "";
      })
      .filter(Boolean)
      .join("")}`;
    if (prev) sel.value = prev;
  } catch {}
}

function financeSyncPeriodFields(showAlert) {
  const period = String(qs("finMasterPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) {
    if (showAlert) alert("Enter reporting period as YYYY-MM.");
    return;
  }
  const { start, end } = monthBoundsFromPeriod(period);
  const ids = ["finBvaPeriod", "finPeriodInput", "finSsotPeriod", "finForecastStart"];
  ids.forEach((id) => {
    const el = qs(id);
    if (el) el.value = period;
  });
  if (qs("finRunStart")) qs("finRunStart").value = start;
  if (qs("finRunEnd")) qs("finRunEnd").value = end;
  if (showAlert) setStatus(`Finance period set to ${period} (${start} → ${end}).`);
}

function finCatAmount(categories, name) {
  const row = (Array.isArray(categories) ? categories : []).find((c) => String(c.category) === name);
  return row ? Number(row.actual || 0) : 0;
}

async function loadFinanceSiteFilterOptions(sites) {
  const sel = qs("finSiteFilter");
  if (!sel) return;
  const prev = String(sel.value || "");
  const codes = Array.from(
    new Set((sites || []).map((s) => String(s.site_code || "").trim()).filter(Boolean))
  ).sort((a, b) => a.localeCompare(b));
  sel.innerHTML = `<option value="">All sites</option>${codes
    .map((c) => {
      const name = (sites || []).find((s) => String(s.site_code) === c)?.site_name || c;
      return `<option value="${escapeHtml(c)}">${escapeHtml(name)} (${escapeHtml(c)})</option>`;
    })
    .join("")}`;
  if (prev && codes.includes(prev)) sel.value = prev;
}

async function loadFinanceSiteAllocation() {
  const period = String(qs("finMasterPeriod")?.value || "").trim();
  const site = String(qs("finSiteFilter")?.value || "").trim();
  const msg = qs("finSiteAllocMsg");
  const body = qs("finSiteAllocBody");
  const totalsEl = qs("finSiteAllocTotals");
  const warnEl = qs("finSiteAllocWarn");
  if (!/^\d{4}-\d{2}$/.test(period)) {
    alert("Enter reporting period as YYYY-MM.");
    return;
  }
  if (msg) msg.textContent = "Loading site allocation…";
  if (body) body.innerHTML = `<tr><td colspan="11" class="muted">Loading…</td></tr>`;
  const q = new URLSearchParams({ period });
  if (site) q.set("site_code", site);
  try {
    const res = await fetchJson(`${API}/api/finance/site-allocation?${q.toString()}`);
    financeSiteAllocCache = res;
    const sites = Array.isArray(res.sites) ? res.sites : [];
    await loadFinanceSiteFilterOptions(sites);
    const t = res.totals || {};
    if (totalsEl) {
      totalsEl.innerHTML = `
        <div class="kpi-card kpi-util">
          <div class="kpi-card-header"><div class="kpi-icon">B</div><div class="kpi-title">Budget</div></div>
          <div class="kpi-big-value">${fmtMoney(t.budget)}</div>
          <div class="kpi-meta">Period ${escapeHtml(period)}</div>
        </div>
        <div class="kpi-card kpi-scheduled">
          <div class="kpi-card-header"><div class="kpi-icon">A</div><div class="kpi-title">Actual</div></div>
          <div class="kpi-big-value">${fmtMoney(t.actual)}</div>
          <div class="kpi-meta">${escapeHtml(res.period_start || "")} → ${escapeHtml(res.period_end || "")}</div>
        </div>
        <div class="kpi-card kpi-alerts">
          <div class="kpi-card-header"><div class="kpi-icon">V</div><div class="kpi-title">Variance</div></div>
          <div class="kpi-big-value">${fmtMoney(t.variance)}</div>
          <div class="kpi-meta">Actual minus budget</div>
        </div>
      `;
    }
    if (warnEl) {
      const miss = Number(res.assets_missing_site || 0);
      warnEl.innerHTML = miss > 0
        ? `<div class="message-error" style="padding:8px 10px; border-radius:8px;">${miss} active asset(s) have no <code>site_code</code> — their costs appear under <strong>Unassigned</strong>. Update assets to improve site splits.</div>`
        : (res.has_asset_site_column === false
          ? `<div class="message-error" style="padding:8px 10px; border-radius:8px;">Asset <code>site_code</code> column missing — restart API to apply schema update.</div>`
          : "");
    }
    if (body) {
      body.innerHTML = sites.length
        ? sites.map((s) => `<tr>
            <td>${escapeHtml(s.site_name || s.site_code || "-")}</td>
            <td><code>${escapeHtml(s.site_code || "")}</code></td>
            <td class="num">${fmtMoney(s.budget_total)}</td>
            <td class="num">${fmtMoney(s.total_actual)}</td>
            <td class="num">${fmtMoney(s.variance)}</td>
            <td class="num">${s.variance_pct == null ? "-" : Number(s.variance_pct).toFixed(1) + "%"}</td>
            <td class="num">${fmtMoney(finCatAmount(s.categories, "parts"))}</td>
            <td class="num">${fmtMoney(finCatAmount(s.categories, "labor"))}</td>
            <td class="num">${fmtMoney(finCatAmount(s.categories, "fuel"))}</td>
            <td class="num">${fmtMoney(finCatAmount(s.categories, "lube"))}</td>
            <td class="num">${fmtMoney(finCatAmount(s.categories, "downtime"))}</td>
          </tr>`).join("")
        : `<tr><td colspan="11" class="muted">No costs in this period${site ? ` for site ${escapeHtml(site)}` : ""}.</td></tr>`;
    }
    if (msg) {
      msg.textContent = `Loaded ${sites.length} site row(s) for ${period}.`;
      msg.className = "message-success";
    }
    financeSyncPeriodFields(false);
  } catch (e) {
    financeSiteAllocCache = null;
    if (msg) {
      msg.textContent = `Load failed: ${e.message}`;
      msg.className = "message-error";
    }
    if (body) body.innerHTML = `<tr><td colspan="11" class="message-error">${escapeHtml(e.message)}</td></tr>`;
  }
}

function financeExportSiteAllocation() {
  const period = String(qs("finMasterPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) {
    alert("Enter reporting period as YYYY-MM.");
    return;
  }
  const site = String(qs("finSiteFilter")?.value || "").trim();
  const q = new URLSearchParams({ period });
  if (site) q.set("site_code", site);
  return downloadAuthedFile(`${API}/api/finance/site-allocation/export.xlsx?${q.toString()}`, `IRONLOG_Site_Allocation_${period}.xlsx`);
}

async function financeBuildSummarizedRun() {
  const start = String(qs("finRunStart")?.value || "").trim();
  const end = String(qs("finRunEnd")?.value || "").trim();
  if (!start || !end) { alert("Select start and end dates."); return; }
  const categories = [];
  if (qs("finCatParts")?.checked) categories.push("parts");
  if (qs("finCatLabor")?.checked) categories.push("labor");
  if (qs("finCatDowntime")?.checked) categories.push("downtime");
  if (qs("finCatFuel")?.checked) categories.push("fuel");
  if (qs("finCatLube")?.checked) categories.push("lube");
  if (qs("finCatGrn")?.checked) categories.push("procurement_grn");
  if (qs("finCatAp")?.checked) categories.push("procurement_ap");
  if (!categories.length) { alert("Select at least one category."); return; }
  setStatus("Building summarized journal run...");
  try {
    const res = await fetchJson(`${API}/api/procurement/journals/summarize`, {
      method: "POST",
      body: JSON.stringify({
        start,
        end,
        categories,
        default_cost_center_code: String(qs("finRunDefaultCC")?.value || "").trim() || undefined,
        currency: String(qs("finRunCurrency")?.value || "USD").trim() || "USD",
      }),
    });
    const el = qs("finRunMsg");
    if (el) el.textContent = `Run ${res.run?.run_number} built | lines=${res.run?.line_count} | debit=${fmtMoney(res.run?.total_debit)} | credit=${fmtMoney(res.run?.total_credit)} | balanced=${res.balanced}`;
    financeLastRunId = Number(res.run?.id || 0) || null;
    const rid = qs("finActiveRunId");
    if (rid && financeLastRunId) rid.value = String(financeLastRunId);
    await loadFinanceRuns();
    setStatus("Summarized journal run built.");
  } catch (e) {
    setStatus(`Build failed: ${e.message}`);
    alert(`Build failed: ${e.message}`);
  }
}

async function loadFinanceRuns() {
  const list = qs("finRunList");
  if (!list) return;
  const status = String(qs("finRunStatusFilter")?.value || "").trim();
  const q = status ? `?status=${encodeURIComponent(status)}` : "";
  try {
    const data = await fetchJson(`${API}/api/procurement/journals/runs${q}`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    if (!rows.length) {
      list.innerHTML = "<em>No runs found.</em>";
      return;
    }
    list.innerHTML = `
      <table class="table">
        <thead><tr>
          <th>ID</th><th>Run #</th><th>Period</th><th>Range</th><th>Status</th>
          <th class="num">Debit</th><th class="num">Credit</th><th class="num">Lines</th><th>Created</th>
        </tr></thead>
        <tbody>
          ${rows.map((r) => `
            <tr>
              <td><a href="#" data-fin-run-id="${r.id}">${r.id}</a></td>
              <td>${escapeHtml(r.run_number)}</td>
              <td>${escapeHtml(r.period || "")}</td>
              <td>${escapeHtml(r.start_date || "")} → ${escapeHtml(r.end_date || "")}</td>
              <td>${escapeHtml(r.status || "")}</td>
              <td class="num">${fmtMoney(r.total_debit)}</td>
              <td class="num">${fmtMoney(r.total_credit)}</td>
              <td class="num">${Number(r.line_count || 0)}</td>
              <td>${escapeHtml((r.created_at || "").slice(0, 16))}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `;
    list.querySelectorAll("[data-fin-run-id]").forEach((a) => {
      a.addEventListener("click", (ev) => {
        ev.preventDefault();
        const id = Number(a.getAttribute("data-fin-run-id")) || 0;
        if (!id) return;
        financeLastRunId = id;
        const el = qs("finActiveRunId");
        if (el) el.value = String(id);
        loadFinanceRunDetail().catch(() => {});
      });
    });
  } catch (e) {
    list.innerHTML = `<em class="muted">Failed: ${escapeHtml(e.message)}</em>`;
  }
}

async function loadFinanceRunDetail() {
  const id = Number(qs("finActiveRunId")?.value || 0);
  if (!id) { alert("Enter a run ID first."); return; }
  const el = qs("finRunDetail");
  if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/procurement/journals/runs/${id}`);
    const run = res.run || {};
    const byCat = Array.isArray(res.by_category) ? res.by_category : [];
    const bySite = Array.isArray(res.by_site) ? res.by_site : [];
    el.innerHTML = `
      <div class="kpi-pills">
        <span class="kpi-pill"><strong>Run:</strong> ${escapeHtml(run.run_number || "")}</span>
        <span class="kpi-pill"><strong>Status:</strong> ${escapeHtml(run.status || "")}</span>
        <span class="kpi-pill"><strong>Debit:</strong> ${fmtMoney(run.total_debit)}</span>
        <span class="kpi-pill"><strong>Credit:</strong> ${fmtMoney(run.total_credit)}</span>
        <span class="kpi-pill"><strong>Lines:</strong> ${Number(run.line_count || 0)}</span>
      </div>
      <h5 style="margin-top:12px;">By site</h5>
      <table class="table">
        <thead><tr><th>Site</th><th class="num">Lines</th><th class="num">Debit</th><th class="num">Credit</th></tr></thead>
        <tbody>
          ${bySite.length ? bySite.map((s) => `<tr>
            <td><code>${escapeHtml(s.site_code || "")}</code></td>
            <td class="num">${Number(s.lines || 0)}</td>
            <td class="num">${fmtMoney(s.debit_total)}</td>
            <td class="num">${fmtMoney(s.credit_total)}</td>
          </tr>`).join("") : `<tr><td colspan="4" class="muted">No site breakdown on lines.</td></tr>`}
        </tbody>
      </table>
      <h5 style="margin-top:12px;">By category</h5>
      <table class="table">
        <thead><tr><th>Category</th><th class="num">Lines</th><th class="num">Debit</th><th class="num">Credit</th></tr></thead>
        <tbody>
          ${byCat.map((c) => `<tr>
            <td>${escapeHtml(c.category)}</td>
            <td class="num">${Number(c.lines || 0)}</td>
            <td class="num">${fmtMoney(c.debit_total)}</td>
            <td class="num">${fmtMoney(c.credit_total)}</td>
          </tr>`).join("")}
        </tbody>
      </table>
    `;
  } catch (e) {
    el.innerHTML = `<em class="muted">Load failed: ${escapeHtml(e.message)}</em>`;
  }
}

function financeExportRunCsv() {
  const id = Number(qs("finActiveRunId")?.value || 0);
  if (!id) { alert("Enter a run ID first."); return; }
  return downloadAuthedFile(`${API}/api/procurement/journals/runs/${id}/export.csv`, `IRONLOG_Journal_Run_${id}.csv`);
}
function financeExportRunXlsx() {
  const id = Number(qs("finActiveRunId")?.value || 0);
  if (!id) { alert("Enter a run ID first."); return; }
  return downloadAuthedFile(`${API}/api/procurement/journals/runs/${id}/export.xlsx`, `IRONLOG_Journal_Run_${id}.xlsx`);
}

async function financeMarkExported() {
  const id = Number(qs("finActiveRunId")?.value || 0);
  if (!id) { alert("Enter a run ID first."); return; }
  try {
    await fetchJson(`${API}/api/procurement/journals/runs/${id}/mark-exported`, { method: "POST", body: JSON.stringify({}) });
    setStatus("Run marked exported.");
    await loadFinanceRuns();
    await loadFinanceRunDetail().catch(() => {});
  } catch (e) { alert(e.message); }
}
async function financeMarkPosted() {
  const id = Number(qs("finActiveRunId")?.value || 0);
  if (!id) { alert("Enter a run ID first."); return; }
  const ref = String(qs("finPostedRef")?.value || "").trim();
  if (!confirm(`Mark run ${id} as posted? This locks it from further changes.`)) return;
  try {
    await fetchJson(`${API}/api/procurement/journals/runs/${id}/mark-posted`, {
      method: "POST",
      body: JSON.stringify({ posted_reference: ref || null }),
    });
    setStatus("Run marked posted.");
    await loadFinanceRuns();
    await loadFinanceRunDetail().catch(() => {});
  } catch (e) { alert(e.message); }
}
async function financeReverseRun() {
  const id = Number(qs("finActiveRunId")?.value || 0);
  if (!id) { alert("Enter a run ID first."); return; }
  const reason = String(qs("finReverseReason")?.value || "").trim();
  if (!reason) { alert("Reason required."); return; }
  if (!confirm(`Reverse run ${id}? Reason: ${reason}`)) return;
  try {
    await fetchJson(`${API}/api/procurement/journals/runs/${id}/reverse`, {
      method: "POST",
      body: JSON.stringify({ reason }),
    });
    setStatus("Run reversed.");
    await loadFinanceRuns();
    await loadFinanceRunDetail().catch(() => {});
  } catch (e) { alert(e.message); }
}

async function loadFinanceChecklist() {
  const period = String(qs("finPeriodInput")?.value || "").trim();
  const msg = qs("finChecklistMsg");
  const tbl = qs("finChecklistTable");
  if (!period || !/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  try {
    const res = await fetchJson(`${API}/api/finance/periods/${period}/checklist`);
    const items = Array.isArray(res.items) ? res.items : [];
    const lock = res.lock || {};
    if (msg) msg.textContent = `Period ${period} | lock status: ${lock.status || "open"} | ${items.length} items`;
    if (tbl) {
      tbl.innerHTML = `
        <table class="table">
          <thead><tr><th>Code</th><th>Task</th><th>Status</th><th>Updated</th><th>Actions</th></tr></thead>
          <tbody>
            ${items.map((it) => `
              <tr>
                <td><code>${escapeHtml(it.code)}</code></td>
                <td>${escapeHtml(it.label || "")}</td>
                <td>${escapeHtml(it.status || "")}</td>
                <td>${escapeHtml((it.updated_at || "").slice(0, 16))}</td>
                <td>
                  <button type="button" data-chk-mark-done="${escapeHtml(it.code)}">Done</button>
                  <button type="button" data-chk-mark-skip="${escapeHtml(it.code)}">Skip</button>
                  <button type="button" data-chk-mark-pending="${escapeHtml(it.code)}">Reset</button>
                </td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      `;
      tbl.querySelectorAll("[data-chk-mark-done]").forEach((b) => b.addEventListener("click", () => financeChecklistUpdate(period, b.getAttribute("data-chk-mark-done"), "done")));
      tbl.querySelectorAll("[data-chk-mark-skip]").forEach((b) => b.addEventListener("click", () => financeChecklistUpdate(period, b.getAttribute("data-chk-mark-skip"), "skipped")));
      tbl.querySelectorAll("[data-chk-mark-pending]").forEach((b) => b.addEventListener("click", () => financeChecklistUpdate(period, b.getAttribute("data-chk-mark-pending"), "pending")));
    }
  } catch (e) {
    if (msg) msg.textContent = `Load failed: ${e.message}`;
  }
}

async function financeChecklistUpdate(period, code, status) {
  try {
    await fetchJson(`${API}/api/finance/periods/${period}/checklist/${code}`, {
      method: "POST",
      body: JSON.stringify({ status }),
    });
    await loadFinanceChecklist();
  } catch (e) { alert(e.message); }
}

async function financeClosePeriod() {
  const period = String(qs("finPeriodInput")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  const force = qs("finClosePeriodForce")?.checked || false;
  if (!confirm(`Close period ${period}? force=${force}`)) return;
  try {
    const res = await fetchJson(`${API}/api/finance/periods/${period}/close`, {
      method: "POST",
      body: JSON.stringify({ force }),
    });
    setStatus(`Period ${period} closed.`);
    await loadFinanceChecklist();
  } catch (e) { alert(e.message); }
}
async function financeReopenPeriod() {
  const period = String(qs("finPeriodInput")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  const reason = String(qs("finReopenReason")?.value || "").trim();
  if (!reason) { alert("Reason required to reopen"); return; }
  if (!confirm(`Reopen period ${period}? Reason: ${reason}`)) return;
  try {
    await fetchJson(`${API}/api/finance/periods/${period}/reopen`, {
      method: "POST",
      body: JSON.stringify({ reason }),
    });
    setStatus(`Period ${period} reopened.`);
    await loadFinanceChecklist();
  } catch (e) { alert(e.message); }
}

async function loadFinanceBudgetVsActual() {
  const period = String(qs("finBvaPeriod")?.value || "").trim();
  const dim = String(qs("finBvaDimension")?.value || "site_code").trim();
  const site = String(qs("finSiteFilter")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  try {
    const q = new URLSearchParams({ period, dimension: dim });
    if (site) q.set("site_code", site);
    const res = await fetchJson(`${API}/api/finance/budgets-vs-actual?${q.toString()}`);
    const rows = Array.isArray(res.rows) ? res.rows : [];
    const total = res.total || {};
    const totalsEl = qs("finBvaTotals");
    if (totalsEl) {
      totalsEl.innerHTML = `
        <div class="kpi-pills">
          <span class="kpi-pill"><strong>Budget:</strong> ${fmtMoney(total.budget)}</span>
          <span class="kpi-pill"><strong>Actual:</strong> ${fmtMoney(total.actual)}</span>
          <span class="kpi-pill"><strong>Variance:</strong> ${fmtMoney(total.variance)}</span>
        </div>
      `;
    }
    const dimLabel = dim === "site_code" ? "Site" : dim === "cost_center_code" ? "Cost Center" : dim;
    const tbl = qs("finBvaTable");
    if (tbl) {
      tbl.innerHTML = `
        <table class="table">
          <thead><tr><th>${escapeHtml(dimLabel)}</th><th class="num">Budget</th><th class="num">Actual</th><th class="num">Variance</th><th class="num">Variance %</th></tr></thead>
          <tbody>
            ${rows.map((r) => `<tr>
              <td>${escapeHtml(r.dimension_key || "(none)")}</td>
              <td class="num">${fmtMoney(r.budget)}</td>
              <td class="num">${fmtMoney(r.actual)}</td>
              <td class="num">${fmtMoney(r.variance)}</td>
              <td class="num">${r.variance_pct == null ? "-" : Number(r.variance_pct).toFixed(1) + "%"}</td>
            </tr>`).join("")}
          </tbody>
        </table>
      `;
    }
  } catch (e) { alert(e.message); }
}

async function financeSaveBudgets() {
  const raw = String(qs("finBudgetJson")?.value || "").trim();
  if (!raw) { alert("Paste budget rows JSON first"); return; }
  let rows;
  try { rows = JSON.parse(raw); } catch { alert("Invalid JSON"); return; }
  if (!Array.isArray(rows)) { alert("JSON must be an array"); return; }
  try {
    const res = await fetchJson(`${API}/api/finance/budgets/upsert`, {
      method: "POST",
      body: JSON.stringify({ rows }),
    });
    setStatus(`Saved ${res.saved || 0} budget rows.`);
    await loadFinanceBudgetList();
  } catch (e) { alert(e.message); }
}

async function loadFinanceBudgetList() {
  const period = String(qs("finBvaPeriod")?.value || "").trim();
  const q = period ? `?period=${encodeURIComponent(period)}` : "";
  try {
    const res = await fetchJson(`${API}/api/finance/budgets${q}`);
    const rows = Array.isArray(res.rows) ? res.rows : [];
    const el = qs("finBudgetList");
    if (!el) return;
    if (!rows.length) {
      el.innerHTML = "<em>No budgets.</em>";
      return;
    }
    el.innerHTML = `
      <table class="table">
        <thead><tr><th>Period</th><th>Site</th><th>Cost Center</th><th>Equipment</th><th>Category</th><th class="num">Budget</th><th>Currency</th></tr></thead>
        <tbody>
          ${rows.map((r) => `<tr>
            <td>${escapeHtml(r.period)}</td>
            <td>${escapeHtml(r.site_code || "")}</td>
            <td>${escapeHtml(r.cost_center_code || "")}</td>
            <td>${escapeHtml(r.equipment_type || "")}</td>
            <td>${escapeHtml(r.category)}</td>
            <td class="num">${fmtMoney(r.budget_amount)}</td>
            <td>${escapeHtml(r.currency || "USD")}</td>
          </tr>`).join("")}
        </tbody>
      </table>
    `;
  } catch (e) {
    const el = qs("finBudgetList");
    if (el) el.innerHTML = `<em class="muted">Failed: ${escapeHtml(e.message)}</em>`;
  }
}

async function financeRebuildForecast() {
  const start = String(qs("finForecastStart")?.value || "").trim();
  const months = Math.max(1, Math.min(6, Number(qs("finForecastMonths")?.value || 3)));
  if (!/^\d{4}-\d{2}$/.test(start)) { alert("Enter start period as YYYY-MM"); return; }
  try {
    const res = await fetchJson(`${API}/api/finance/forecast/rebuild`, {
      method: "POST",
      body: JSON.stringify({ start_period: start, months }),
    });
    financeLastForecastBatchId = res.batch_id || null;
    setStatus(`Forecast rebuilt | batch ${res.batch_id} | ${res.saved || 0} rows.`);
    await loadFinanceForecast();
  } catch (e) { alert(e.message); }
}

async function loadFinanceForecast() {
  const msg = qs("finForecastMsg");
  const totalsEl = qs("finForecastTotals");
  const tbl = qs("finForecastTable");
  try {
    const q = financeLastForecastBatchId ? `?batch_id=${encodeURIComponent(financeLastForecastBatchId)}` : "";
    const res = await fetchJson(`${API}/api/finance/forecast${q}`);
    const rows = Array.isArray(res.rows) ? res.rows : [];
    const totals = Array.isArray(res.totals_by_period) ? res.totals_by_period : [];
    if (msg) msg.textContent = `${rows.length} forecast rows across ${totals.length} months`;
    if (totalsEl) {
      totalsEl.innerHTML = totals.map((t) => `
        <span class="kpi-pill"><strong>${escapeHtml(t.period)}:</strong> ${fmtMoney(t.forecast)} (baseline ${fmtMoney(t.baseline)} + uplift ${fmtMoney(t.uplift)})</span>
      `).join(" ");
    }
    if (tbl) {
      tbl.innerHTML = `
        <table class="table">
          <thead><tr><th>Period</th><th>Site</th><th>Cost Center</th><th>Equipment</th><th>Category</th>
            <th class="num">Baseline</th><th class="num">Uplift</th><th class="num">Forecast</th></tr></thead>
          <tbody>
            ${rows.map((r) => `<tr>
              <td>${escapeHtml(r.period)}</td>
              <td>${escapeHtml(r.site_code || "")}</td>
              <td>${escapeHtml(r.cost_center_code || "")}</td>
              <td>${escapeHtml(r.equipment_type || "")}</td>
              <td>${escapeHtml(r.category)}</td>
              <td class="num">${fmtMoney(r.baseline_amount)}</td>
              <td class="num">${fmtMoney(r.uplift_amount)}</td>
              <td class="num">${fmtMoney(r.forecast_amount)}</td>
            </tr>`).join("")}
          </tbody>
        </table>
      `;
    }
  } catch (e) {
    if (msg) msg.textContent = `Load failed: ${e.message}`;
  }
}

async function loadFinanceSsot() {
  const period = String(qs("finSsotPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  try {
    const res = await fetchJson(`${API}/api/finance/reports/ssot?period=${encodeURIComponent(period)}`);
    const kpi = res.kpi || {};
    const kel = qs("finSsotKpi");
    if (kel) {
      kel.innerHTML = `
        <div class="kpi-pills">
          <span class="kpi-pill"><strong>Availability:</strong> ${kpi.availability == null ? "-" : kpi.availability + "%"}</span>
          <span class="kpi-pill"><strong>Utilization:</strong> ${kpi.utilization == null ? "-" : kpi.utilization + "%"}</span>
          <span class="kpi-pill"><strong>MTBF:</strong> ${kpi.mtbf == null ? "-" : kpi.mtbf + " h"}</span>
          <span class="kpi-pill"><strong>MTTR:</strong> ${kpi.mttr == null ? "-" : kpi.mttr + " h"}</span>
          <span class="kpi-pill"><strong>Cost/Asset-hr:</strong> ${kpi.cost_per_asset_hour == null ? "-" : fmtMoney(kpi.cost_per_asset_hour)}</span>
          <span class="kpi-pill"><strong>Run hrs:</strong> ${kpi.run_hours}</span>
          <span class="kpi-pill"><strong>Down hrs:</strong> ${kpi.downtime_hours}</span>
          <span class="kpi-pill"><strong>Total Cost:</strong> ${fmtMoney(kpi.total_cost)}</span>
        </div>
      `;
    }
    const actuals = Array.isArray(res.actuals) ? res.actuals : [];
    const ael = qs("finSsotActuals");
    if (ael) {
      ael.innerHTML = `
        <h5>Actuals Breakdown</h5>
        <table class="table">
          <thead><tr><th>Site</th><th>Cost Center</th><th>Equipment</th><th>Category</th><th class="num">Actual</th></tr></thead>
          <tbody>
            ${actuals.map((r) => `<tr>
              <td>${escapeHtml(r.site_code || "")}</td>
              <td>${escapeHtml(r.cost_center_code || "")}</td>
              <td>${escapeHtml(r.equipment_type || "")}</td>
              <td>${escapeHtml(r.category)}</td>
              <td class="num">${fmtMoney(r.actual_amount)}</td>
            </tr>`).join("")}
          </tbody>
        </table>
      `;
    }
    const defs = Array.isArray(res.kpi_definitions) ? res.kpi_definitions : [];
    const dEl = qs("finKpiDefs");
    if (dEl) {
      dEl.innerHTML = `
        <table class="table">
          <thead><tr><th>Code</th><th>Label</th><th>Unit</th><th>Formula</th><th>Source Tables</th></tr></thead>
          <tbody>
            ${defs.map((d) => `<tr>
              <td><code>${escapeHtml(d.code)}</code></td>
              <td>${escapeHtml(d.label)}</td>
              <td>${escapeHtml(d.unit)}</td>
              <td><code>${escapeHtml(d.formula || "")}</code></td>
              <td>${escapeHtml((d.source_tables || []).join(", "))}</td>
            </tr>`).join("")}
          </tbody>
        </table>
      `;
    }
  } catch (e) { alert(e.message); }
}
function financeExportSsotCsv() {
  const period = String(qs("finSsotPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  return downloadAuthedFile(`${API}/api/finance/reports/ssot/export.csv?period=${encodeURIComponent(period)}`, `IRONLOG_SSOT_${period}.csv`);
}
function financeExportSsotXlsx() {
  const period = String(qs("finSsotPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("Enter period as YYYY-MM"); return; }
  return downloadAuthedFile(`${API}/api/finance/reports/ssot/export.xlsx?period=${encodeURIComponent(period)}`, `IRONLOG_SSOT_${period}.xlsx`);
}

async function loadFinanceKpiDefs() {
  const el = qs("finKpiDefs");
  if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/finance/kpis/definitions`);
    const defs = Array.isArray(res.kpis) ? res.kpis : [];
    el.innerHTML = `
      <table class="table">
        <thead><tr><th>Code</th><th>Label</th><th>Unit</th><th>Formula</th><th>Source Tables</th><th>Exclusions</th></tr></thead>
        <tbody>
          ${defs.map((d) => `<tr>
            <td><code>${escapeHtml(d.code)}</code></td>
            <td>${escapeHtml(d.label)}</td>
            <td>${escapeHtml(d.unit)}</td>
            <td><code>${escapeHtml(d.formula || "")}</code></td>
            <td>${escapeHtml((d.source_tables || []).join(", "))}</td>
            <td>${escapeHtml((d.exclusions || []).join(", "))}</td>
          </tr>`).join("")}
        </tbody>
      </table>
    `;
  } catch (e) { el.innerHTML = `<em class="muted">Failed: ${escapeHtml(e.message)}</em>`; }
}

function bindFinanceHandlers() {
  const bind = (id, ev, fn) => { const el = qs(id); if (el) el.addEventListener(ev, fn); };
  bind("finLoadSiteAllocBtn", "click", () => loadFinanceSiteAllocation().catch((e) => alert(e.message)));
  bind("finExportSiteAllocBtn", "click", financeExportSiteAllocation);
  bind("finSyncPeriodsBtn", "click", () => financeSyncPeriodFields(true));
  bind("finBuildRunBtn", "click", financeBuildSummarizedRun);
  bind("finLoadRunsBtn", "click", () => loadFinanceRuns().catch(() => {}));
  bind("finLoadRunDetailBtn", "click", () => loadFinanceRunDetail().catch(() => {}));
  bind("finExportRunCsvBtn", "click", financeExportRunCsv);
  bind("finExportRunXlsxBtn", "click", financeExportRunXlsx);
  bind("finMarkExportedBtn", "click", financeMarkExported);
  bind("finMarkPostedBtn", "click", financeMarkPosted);
  bind("finReverseRunBtn", "click", financeReverseRun);
  bind("finLoadChecklistBtn", "click", loadFinanceChecklist);
  bind("finClosePeriodBtn", "click", financeClosePeriod);
  bind("finReopenPeriodBtn", "click", financeReopenPeriod);
  bind("finLoadBvaBtn", "click", loadFinanceBudgetVsActual);
  bind("finSaveBudgetBtn", "click", financeSaveBudgets);
  bind("finLoadBudgetListBtn", "click", loadFinanceBudgetList);
  bind("finRebuildForecastBtn", "click", financeRebuildForecast);
  bind("finLoadForecastBtn", "click", loadFinanceForecast);
  bind("finLoadSsotBtn", "click", loadFinanceSsot);
  bind("finExportSsotCsvBtn", "click", financeExportSsotCsv);
  bind("finExportSsotXlsxBtn", "click", financeExportSsotXlsx);
  bind("finLoadKpiDefsBtn", "click", loadFinanceKpiDefs);
  const runStatusSel = qs("finRunStatusFilter");
  if (runStatusSel) runStatusSel.addEventListener("change", () => loadFinanceRuns().catch(() => {}));
}

/* =====================================================================
   ENTERPRISE (entity/integrations/governance/evidence)
===================================================================== */
