// IRONLOG/web/app/enterprise.js — Enterprise, integrations, governance, evidence packs, executive.
// Part of the main app; index.html loads these files in order and they share one global scope.

async function entSaveCompany() {
  const payload = {
    company_code: qs("entCompanyCode")?.value,
    company_name: qs("entCompanyName")?.value,
    base_currency: qs("entCompanyBaseCcy")?.value,
    reporting_currency: qs("entCompanyReportCcy")?.value,
    tax_region: qs("entCompanyTaxRegion")?.value,
  };
  if (!payload.company_code || !payload.company_name) { alert("company code + name required"); return; }
  try { await fetchJson(`${API}/api/entity/companies`, { method: "POST", body: JSON.stringify(payload) }); setStatus("Company saved."); entLoadCompanies(); }
  catch (e) { alert(e.message); }
}

async function entLoadCompanies() {
  const el = qs("entCompanyList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/entity/companies`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No companies.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>Code</th><th>Name</th><th>Base</th><th>Reporting</th><th>Active</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${escapeHtml(r.company_code)}</td><td>${escapeHtml(r.company_name)}</td><td>${escapeHtml(r.base_currency || "")}</td><td>${escapeHtml(r.reporting_currency || "")}</td><td>${r.active ? "yes" : "no"}</td></tr>`).join("")}</tbody></table>`;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function entSaveSite() {
  const payload = {
    site_code: qs("entSiteCode")?.value,
    company_code: qs("entSiteCompany")?.value,
    site_name: qs("entSiteName")?.value,
    local_currency: qs("entSiteLocalCcy")?.value,
    region: qs("entSiteRegion")?.value,
  };
  if (!payload.site_code || !payload.company_code || !payload.site_name) { alert("site_code, company_code, site_name required"); return; }
  try { await fetchJson(`${API}/api/entity/sites`, { method: "POST", body: JSON.stringify(payload) }); setStatus("Site saved."); entLoadSites(); }
  catch (e) { alert(e.message); }
}

async function entLoadSites() {
  const el = qs("entSiteList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/entity/sites`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No sites.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>Site</th><th>Company</th><th>Name</th><th>Local Ccy</th><th>Region</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${escapeHtml(r.site_code)}</td><td>${escapeHtml(r.company_code)}</td><td>${escapeHtml(r.site_name)}</td><td>${escapeHtml(r.local_currency || "")}</td><td>${escapeHtml(r.region || "")}</td></tr>`).join("")}</tbody></table>`;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function entLoadTree() {
  const el = qs("entTree"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/entity/tree`);
    const tree = res.tree || [];
    if (!tree.length) { el.innerHTML = "<em>No entities.</em>"; return; }
    el.innerHTML = tree.map((c) => `<div><strong>${escapeHtml(c.company_code)}</strong> - ${escapeHtml(c.company_name)}<ul>${(c.sites || []).map((s) => `<li>${escapeHtml(s.site_code)} - ${escapeHtml(s.site_name)}</li>`).join("")}</ul></div>`).join("");
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function entSaveCcyRates() {
  const raw = String(qs("entCcyJson")?.value || "").trim();
  if (!raw) { alert("paste rows JSON"); return; }
  let rows; try { rows = JSON.parse(raw); } catch { alert("invalid JSON"); return; }
  try { const res = await fetchJson(`${API}/api/entity/currency/rates/upsert`, { method: "POST", body: JSON.stringify({ rows }) }); setStatus(`Saved ${res.saved || 0} rates.`); entLoadCcyRates(); }
  catch (e) { alert(e.message); }
}

async function entLoadCcyRates() {
  const el = qs("entCcyList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/entity/currency/rates`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No rates.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>From</th><th>To</th><th>Rate</th><th>Effective</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${escapeHtml(r.from_currency)}</td><td>${escapeHtml(r.to_currency)}</td><td>${Number(r.rate).toFixed(4)}</td><td>${escapeHtml(r.effective_date)}</td></tr>`).join("")}</tbody></table>`;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function entSaveTax() {
  const payload = {
    tax_code: qs("entTaxCode")?.value,
    label: qs("entTaxLabel")?.value,
    rate_pct: Number(qs("entTaxRate")?.value || 0),
    region: qs("entTaxRegion")?.value,
  };
  if (!payload.tax_code || !payload.label) { alert("tax_code and label required"); return; }
  try { await fetchJson(`${API}/api/entity/tax/profiles/upsert`, { method: "POST", body: JSON.stringify(payload) }); setStatus("Tax profile saved."); entLoadTax(); }
  catch (e) { alert(e.message); }
}

async function entLoadTax() {
  const el = qs("entTaxList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/entity/tax/profiles`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No tax profiles.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>Code</th><th>Label</th><th>Rate %</th><th>Region</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${escapeHtml(r.tax_code)}</td><td>${escapeHtml(r.label)}</td><td>${Number(r.rate_pct).toFixed(2)}</td><td>${escapeHtml(r.region || "")}</td></tr>`).join("")}</tbody></table>`;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function intSaveConnection() {
  const payload = {
    connection_code: qs("intConnCode")?.value,
    connector_key: qs("intConnKey")?.value,
    label: qs("intConnLabel")?.value,
  };
  if (!payload.connection_code || !payload.label) { alert("connection_code + label required"); return; }
  try { await fetchJson(`${API}/api/integrations/connections/upsert`, { method: "POST", body: JSON.stringify(payload) }); setStatus("Connection saved."); intLoadConnections(); }
  catch (e) { alert(e.message); }
}

async function intLoadConnections() {
  const el = qs("intConnList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/integrations/connections`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No connections.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>Code</th><th>Connector</th><th>Label</th><th>Active</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${escapeHtml(r.connection_code)}</td><td>${escapeHtml(r.connector_key)}</td><td>${escapeHtml(r.label)}</td><td>${r.active ? "yes" : "no"}</td></tr>`).join("")}</tbody></table>`;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function intEnqueueJob() {
  const connection_code = qs("intJobConnCode")?.value;
  const connector_key = qs("intJobConnector")?.value;
  const rawPayload = String(qs("intJobPayload")?.value || "").trim();
  const idempotency_key = String(qs("intJobIdem")?.value || "").trim() || null;
  if (!connection_code) { alert("connection_code required"); return; }
  let payload = {};
  if (rawPayload) { try { payload = JSON.parse(rawPayload); } catch { alert("invalid payload JSON"); return; } }
  try {
    const res = await fetchJson(`${API}/api/integrations/jobs/enqueue`, {
      method: "POST",
      body: JSON.stringify({ connection_code, connector_key, payload, idempotency_key }),
    });
    setStatus(`Job ${res.id} ${res.status}${res.duplicate ? " (duplicate)" : ""}`);
    const el = qs("intActiveJobId"); if (el) el.value = String(res.id || "");
    intLoadJobs();
  } catch (e) { alert(e.message); }
}

async function intRunJob() {
  const id = Number(qs("intActiveJobId")?.value || 0);
  if (!id) { alert("job id required"); return; }
  try { const res = await fetchJson(`${API}/api/integrations/jobs/${id}/run`, { method: "POST", body: JSON.stringify({}) }); setStatus(`Run ${id} ${res.status}`); intLoadJobs(); }
  catch (e) { alert(e.message); }
}

async function intRetryJob() {
  const id = Number(qs("intActiveJobId")?.value || 0);
  if (!id) { alert("job id required"); return; }
  try { await fetchJson(`${API}/api/integrations/jobs/${id}/retry-now`, { method: "POST", body: JSON.stringify({}) }); setStatus("Retry queued."); intLoadJobs(); }
  catch (e) { alert(e.message); }
}

async function intCancelJob() {
  const id = Number(qs("intActiveJobId")?.value || 0);
  if (!id) { alert("job id required"); return; }
  const reason = prompt("Cancel reason?") || "cancelled";
  try { await fetchJson(`${API}/api/integrations/jobs/${id}/cancel`, { method: "POST", body: JSON.stringify({ reason }) }); setStatus(`Cancelled ${id}.`); intLoadJobs(); }
  catch (e) { alert(e.message); }
}

async function intLoadJobs() {
  const el = qs("intJobsTable"); if (!el) return;
  const status = String(qs("intJobStatusFilter")?.value || "").trim();
  const q = status ? `?status=${encodeURIComponent(status)}` : "";
  try {
    const res = await fetchJson(`${API}/api/integrations/jobs${q}`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No jobs.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>ID</th><th>Connector</th><th>Status</th><th>Attempts</th><th>Error</th><th>Created</th></tr></thead><tbody>${rows.map((r) => `<tr><td><a href="#" data-int-job-id="${r.id}">${r.id}</a></td><td>${escapeHtml(r.connector_key)}</td><td>${escapeHtml(r.status)}</td><td>${r.attempts}/${r.max_attempts}</td><td>${escapeHtml(r.error_message || "")}</td><td>${escapeHtml((r.created_at || "").slice(0, 16))}</td></tr>`).join("")}</tbody></table>`;
    el.querySelectorAll("[data-int-job-id]").forEach((a) => a.addEventListener("click", (ev) => { ev.preventDefault(); const input = qs("intActiveJobId"); if (input) input.value = a.getAttribute("data-int-job-id"); }));
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function intLoadMonitor() {
  const el = qs("intMonitorSummary"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/integrations/monitoring/summary`);
    el.innerHTML = `
      <div class="kpi-pills">
        ${(res.by_status || []).map((s) => `<span class="kpi-pill"><strong>${escapeHtml(s.status)}:</strong> ${Number(s.c || 0)}</span>`).join(" ")}
        <span class="kpi-pill"><strong>Dead-letter open:</strong> ${Number(res.dead_letter_open || 0)}</span>
        ${res.oldest_queued ? `<span class="kpi-pill"><strong>Oldest queued:</strong> #${res.oldest_queued.id} (${escapeHtml(res.oldest_queued.connector_key)})</span>` : ""}
      </div>
      ${(res.top_errors || []).length ? `<h5>Top errors</h5><ul>${res.top_errors.map((e) => `<li>${escapeHtml(e.error_message || "")} &times; ${Number(e.c || 0)}</li>`).join("")}</ul>` : ""}
    `;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function intLoadDeadLetter() {
  const el = qs("intDeadLetterTable"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/integrations/dead-letter`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No dead-letter entries.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>ID</th><th>Job</th><th>Connector</th><th>Reason</th><th>Ack</th><th>Actions</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${r.id}</td><td>${r.job_id}</td><td>${escapeHtml(r.connector_key)}</td><td>${escapeHtml(r.reason || "")}</td><td>${escapeHtml(r.acknowledged_at || "")}</td><td>${r.acknowledged_at ? "" : `<button type="button" data-int-dl-ack="${r.id}">Ack</button>`}</td></tr>`).join("")}</tbody></table>`;
    el.querySelectorAll("[data-int-dl-ack]").forEach((b) => b.addEventListener("click", async () => {
      const id = b.getAttribute("data-int-dl-ack");
      try { await fetchJson(`${API}/api/integrations/dead-letter/${id}/acknowledge`, { method: "POST", body: JSON.stringify({}) }); intLoadDeadLetter(); }
      catch (e) { alert(e.message); }
    }));
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function govLoadPolicies() {
  const el = qs("govPolicyList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/governance/policies`);
    const rows = res.rows || [];
    el.innerHTML = `<h5>Policies</h5><table class="table"><thead><tr><th>Code</th><th>Mode</th><th>Restricted</th><th>Precursor</th><th>Window (min)</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${escapeHtml(r.policy_code)}</td><td>${escapeHtml(r.mode)}</td><td>${escapeHtml(r.restricted_action)}</td><td>${escapeHtml(r.precursor_action)}</td><td>${r.window_minutes}</td></tr>`).join("")}</tbody></table>`;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function govLoadViolations() {
  const el = qs("govViolationList"); if (!el) return;
  const ack = qs("govViolationFilter")?.value || "";
  const q = ack ? `?ack=${encodeURIComponent(ack)}` : "";
  try {
    const res = await fetchJson(`${API}/api/governance/violations${q}`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No violations.</em>"; return; }
    el.innerHTML = `<h5>Violations</h5><table class="table"><thead><tr><th>ID</th><th>Policy</th><th>User</th><th>Restricted</th><th>Blocked</th><th>Ack</th><th>Actions</th></tr></thead><tbody>${rows.map((r) => `<tr><td>${r.id}</td><td>${escapeHtml(r.policy_code)}</td><td>${escapeHtml(r.username)}</td><td>${escapeHtml(r.restricted_action)}</td><td>${r.blocked ? "yes" : "no"}</td><td>${escapeHtml(r.acknowledged_at || "")}</td><td>${r.acknowledged_at ? "" : `<button type="button" data-gov-ack="${r.id}">Acknowledge</button>`}</td></tr>`).join("")}</tbody></table>`;
    el.querySelectorAll("[data-gov-ack]").forEach((b) => b.addEventListener("click", async () => {
      const id = b.getAttribute("data-gov-ack");
      const notes = prompt("Resolution notes (optional)") || "";
      try { await fetchJson(`${API}/api/governance/violations/${id}/acknowledge`, { method: "POST", body: JSON.stringify({ notes }) }); govLoadViolations(); }
      catch (e) { alert(e.message); }
    }));
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function govEvaluate() {
  const action = String(qs("govEvalAction")?.value || "").trim();
  const username = String(qs("govEvalUser")?.value || "").trim();
  if (!action) { alert("action required"); return; }
  try {
    const res = await fetchJson(`${API}/api/governance/evaluate`, { method: "POST", body: JSON.stringify({ action, username }) });
    const el = qs("govEvalResult");
    if (el) el.innerHTML = `<span class="kpi-pill"><strong>Blocked:</strong> ${res.blocked ? "YES" : "no"}</span> <span class="kpi-pill"><strong>Violations:</strong> ${(res.violations || []).length}</span>`;
    govLoadViolations();
  } catch (e) { alert(e.message); }
}

async function evBuildPack() {
  const pack_type = qs("evPackType")?.value;
  const period = String(qs("evPackPeriod")?.value || "").trim();
  const site_code = String(qs("evPackSite")?.value || "").trim() || null;
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("period YYYY-MM required"); return; }
  try {
    const res = await fetchJson(`${API}/api/executive/evidence-packs/build`, { method: "POST", body: JSON.stringify({ pack_type, period, site_code }) });
    setStatus(`Pack built ${res.pack_code}`);
    evLoadPacks();
  } catch (e) { alert(e.message); }
}

async function evLoadPacks() {
  const el = qs("evPackList"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/executive/evidence-packs`);
    const rows = res.rows || [];
    if (!rows.length) { el.innerHTML = "<em>No packs.</em>"; return; }
    el.innerHTML = `<table class="table"><thead><tr><th>ID</th><th>Code</th><th>Type</th><th>Period</th><th>Site</th><th>Created</th></tr></thead><tbody>${rows.map((r) => `<tr><td><a href="#" data-ev-view-id="${r.id}">${r.id}</a></td><td>${escapeHtml(r.pack_code)}</td><td>${escapeHtml(r.pack_type)}</td><td>${escapeHtml(r.period)}</td><td>${escapeHtml(r.site_code || "")}</td><td>${escapeHtml((r.created_at || "").slice(0, 16))}</td></tr>`).join("")}</tbody></table>`;
    el.querySelectorAll("[data-ev-view-id]").forEach((a) => a.addEventListener("click", (ev) => { ev.preventDefault(); const input = qs("evPackViewId"); if (input) input.value = a.getAttribute("data-ev-view-id"); }));
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function evLoadPackDetail() {
  const id = Number(qs("evPackViewId")?.value || 0);
  if (!id) { alert("Enter pack id"); return; }
  const el = qs("evPackDetail"); if (!el) return;
  try {
    const res = await fetchJson(`${API}/api/executive/evidence-packs/${id}`);
    const pack = res.pack || {};
    const summary = pack.summary || {};
    el.innerHTML = `
      <div class="kpi-pills">
        <span class="kpi-pill"><strong>Code:</strong> ${escapeHtml(pack.pack_code || "")}</span>
        <span class="kpi-pill"><strong>Type:</strong> ${escapeHtml(pack.pack_type || "")}</span>
        <span class="kpi-pill"><strong>Integrity:</strong> ${pack.integrity_ok ? "OK" : "FAIL"}</span>
        <span class="kpi-pill"><strong>Hash:</strong> ${escapeHtml((pack.integrity_hash || "").slice(0, 12))}</span>
      </div>
      <h5>Summary</h5>
      <pre class="output">${escapeHtml(JSON.stringify(summary, null, 2))}</pre>
    `;
  } catch (e) { el.innerHTML = `<em class="muted">${escapeHtml(e.message)}</em>`; }
}

async function execLoadCommand() {
  const period = String(qs("execPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("period YYYY-MM required"); return; }
  try {
    const res = await fetchJson(`${API}/api/executive/command-center?period=${encodeURIComponent(period)}`);
    const el = qs("execCommandKpi"); if (!el) return;
    const ops = res.operational || {};
    const gov = res.governance || {};
    const integ = res.integrations || {};
    el.innerHTML = `
      <div class="kpi-pills">
        <span class="kpi-pill"><strong>Run hrs:</strong> ${Number(ops.run_hours || 0).toFixed(1)}</span>
        <span class="kpi-pill"><strong>Downtime hrs:</strong> ${Number(ops.downtime_hours || 0).toFixed(1)}</span>
        <span class="kpi-pill"><strong>Breakdowns:</strong> ${ops.breakdowns || 0}</span>
        <span class="kpi-pill"><strong>Assets used:</strong> ${ops.used_assets_total || 0}</span>
        <span class="kpi-pill"><strong>Gov violations (open):</strong> ${gov.open || 0}</span>
        <span class="kpi-pill"><strong>Dead-letter:</strong> ${integ.dead_letter_open || 0}</span>
      </div>
    `;
  } catch (e) { alert(e.message); }
}

async function execLoadBoard() {
  const period = String(qs("execPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("period YYYY-MM required"); return; }
  try {
    const res = await fetchJson(`${API}/api/executive/board-pack?period=${encodeURIComponent(period)}`);
    const narr = qs("execNarrative");
    if (narr) narr.innerHTML = `<h5>Highlights</h5><ul>${(res.narrative?.bullets || []).map((b) => `<li>${escapeHtml(b)}</li>`).join("")}</ul>`;
    const blocks = qs("execBoardBlocks");
    if (blocks) {
      blocks.innerHTML = `
        <h5>Financial</h5>
        <pre class="output">${escapeHtml(JSON.stringify(res.financial || {}, null, 2))}</pre>
        <h5>Operational</h5>
        <pre class="output">${escapeHtml(JSON.stringify(res.operational || {}, null, 2))}</pre>
        <h5>Integrations</h5>
        <pre class="output">${escapeHtml(JSON.stringify(res.integrations || {}, null, 2))}</pre>
        <h5>Governance</h5>
        <pre class="output">${escapeHtml(JSON.stringify(res.governance || {}, null, 2))}</pre>
      `;
    }
  } catch (e) { alert(e.message); }
}

function execExportBoardXlsx() {
  const period = String(qs("execPeriod")?.value || "").trim();
  if (!/^\d{4}-\d{2}$/.test(period)) { alert("period YYYY-MM required"); return; }
  return downloadAuthedFile(`${API}/api/executive/board-pack/export.xlsx?period=${encodeURIComponent(period)}`, `IRONLOG_Board_Pack_${period}.xlsx`);
}

function bindEnterpriseHandlers() {
  const bind = (id, ev, fn) => { const el = qs(id); if (el) el.addEventListener(ev, fn); };
  bind("entSaveCompanyBtn", "click", entSaveCompany);
  bind("entLoadCompaniesBtn", "click", entLoadCompanies);
  bind("entSaveSiteBtn", "click", entSaveSite);
  bind("entLoadSitesBtn", "click", entLoadSites);
  bind("entLoadTreeBtn", "click", entLoadTree);
  bind("entSaveCcyBtn", "click", entSaveCcyRates);
  bind("entLoadCcyBtn", "click", entLoadCcyRates);
  bind("entSaveTaxBtn", "click", entSaveTax);
  bind("entLoadTaxBtn", "click", entLoadTax);
  bind("intSaveConnBtn", "click", intSaveConnection);
  bind("intLoadConnBtn", "click", intLoadConnections);
  bind("intEnqueueBtn", "click", intEnqueueJob);
  bind("intRunJobBtn", "click", intRunJob);
  bind("intRetryJobBtn", "click", intRetryJob);
  bind("intCancelJobBtn", "click", intCancelJob);
  bind("intLoadJobsBtn", "click", intLoadJobs);
  bind("intJobStatusFilter", "change", intLoadJobs);
  bind("intLoadMonitorBtn", "click", intLoadMonitor);
  bind("intLoadDeadLetterBtn", "click", intLoadDeadLetter);
  bind("govLoadPoliciesBtn", "click", govLoadPolicies);
  bind("govLoadViolationsBtn", "click", govLoadViolations);
  bind("govViolationFilter", "change", govLoadViolations);
  bind("govEvalBtn", "click", govEvaluate);
  bind("evBuildPackBtn", "click", evBuildPack);
  bind("evLoadPacksBtn", "click", evLoadPacks);
  bind("evLoadPackDetailBtn", "click", evLoadPackDetail);
  bind("execLoadCommandBtn", "click", execLoadCommand);
  bind("execLoadBoardBtn", "click", execLoadBoard);
  bind("execExportBoardXlsxBtn", "click", execExportBoardXlsx);
}

document.addEventListener("DOMContentLoaded", () => {
  document.addEventListener("click", (event) => {
    const button = event.target instanceof Element ? event.target.closest("[data-open-authed-pdf]") : null;
    const url = String(button?.getAttribute("data-open-authed-pdf") || "").trim();
    if (!url) return;
    event.preventDefault();
    void openAuthedReport(url, { filename: "IRONLOG_Report.pdf" });
  });
  initDarkMode();
  init().catch((e) => console.error(e));
  try {
    initFinanceTab();
    bindFinanceHandlers();
  } catch (e) { console.error("finance bind failed", e); }
  try { bindEnterpriseHandlers(); } catch (e) { console.error("enterprise bind failed", e); }
  try { if (typeof window.bindStoreQrAdmin === "function") window.bindStoreQrAdmin(); } catch (e) { console.error("store QR bind failed", e); }
});
