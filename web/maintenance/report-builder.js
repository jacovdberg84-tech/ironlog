// IRONLOG/web/maintenance/report-builder.js — Custom report builder and report subscriptions.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

function reportBuilderSelectedColumns() {
  return Array.from(document.querySelectorAll("input[data-report-builder-col]:checked"))
    .map((el) => String(el.getAttribute("data-report-builder-col") || "").trim())
    .filter(Boolean);
}

function renderReportBuilderColumns(columns, selected = []) {
  const wrap = document.getElementById("reportBuilderColumns");
  if (!wrap) return;
  const selectedSet = new Set((Array.isArray(selected) ? selected : []).map((x) => String(x || "").trim()));
  const list = Array.isArray(columns) ? columns : [];
  wrap.innerHTML = list.length
    ? list.map((col) => `
      <label style="display:flex; align-items:center; gap:6px; border:1px solid #e5e7eb; border-radius:8px; padding:6px 8px;">
        <input type="checkbox" data-report-builder-col="${escBackfill(col)}" ${selectedSet.has(col) ? "checked" : ""} />
        <span>${escBackfill(col)}</span>
      </label>
    `).join("")
    : `<small class="muted">No columns available for this dataset.</small>`;
}

function renderReportBuilderPreview(data) {
  const out = document.getElementById("reportBuilderPreview");
  if (!out) return;
  const cols = Array.isArray(data?.columns) ? data.columns : [];
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  out.innerHTML = `
    <div class="muted">Rows: ${Number(data?.count || rows.length || 0)} (limit ${Number(data?.limit || 0)})</div>
    ${insightsRowsTable(cols, rows.map((r) => cols.map((c) => r?.[c] ?? "-")))}
  `;
}

function reportBuilderCurrentConfig() {
  const dataset = String(document.getElementById("reportBuilderDataset")?.value || "").trim();
  const name = String(document.getElementById("reportBuilderName")?.value || "").trim();
  const columns = reportBuilderSelectedColumns();
  const filters = {
    start: String(document.getElementById("reportBuilderStart")?.value || "").trim(),
    end: String(document.getElementById("reportBuilderEnd")?.value || "").trim(),
    asset_id: Number(document.getElementById("reportBuilderAssetId")?.value || 0) || 0,
    status: String(document.getElementById("reportBuilderStatus")?.value || "").trim(),
    limit: Number(document.getElementById("reportBuilderLimit")?.value || 100) || 100,
  };
  return { dataset, name, columns, filters };
}

async function loadReportBuilderMeta() {
  const datasetEl = document.getElementById("reportBuilderDataset");
  const msg = document.getElementById("reportBuilderMsg");
  if (!datasetEl || !msg) return;
  try {
    const res = await fetch(`${API}/reports/custom-builder/meta`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load report builder metadata");
    reportBuilderMeta = Array.isArray(data?.datasets) ? data.datasets : [];
    datasetEl.innerHTML = reportBuilderMeta.length
      ? reportBuilderMeta.map((d) => `<option value="${escBackfill(d.key)}">${escBackfill(d.key)}</option>`).join("")
      : `<option value="">No datasets</option>`;
    const first = reportBuilderMeta[0];
    renderReportBuilderColumns(first?.columns || [], first?.columns?.slice(0, 4) || []);
    msg.className = "muted";
    msg.textContent = "Builder ready.";
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = `Builder meta error: ${e.message || e}`;
  }
}

async function loadReportBuilderTemplates() {
  const tplEl = document.getElementById("reportBuilderTemplate");
  const msg = document.getElementById("reportBuilderMsg");
  if (!tplEl || !msg) return;
  try {
    const res = await fetch(`${API}/reports/custom-builder/templates`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load templates");
    reportBuilderTemplates = Array.isArray(data?.templates) ? data.templates : [];
    tplEl.innerHTML = `<option value="">(None)</option>${reportBuilderTemplates
      .map((t) => `<option value="${Number(t.id || 0)}">${escBackfill(t.name || "Template")} - ${escBackfill(t.dataset || "")}</option>`)
      .join("")}`;
    msg.className = "muted";
    msg.textContent = `Templates loaded: ${reportBuilderTemplates.length}`;
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = `Template load error: ${e.message || e}`;
  }
}

function applyReportBuilderTemplate(id) {
  const tpl = reportBuilderTemplates.find((t) => Number(t.id || 0) === Number(id || 0));
  if (!tpl) return;
  const dsEl = document.getElementById("reportBuilderDataset");
  const nameEl = document.getElementById("reportBuilderName");
  const startEl = document.getElementById("reportBuilderStart");
  const endEl = document.getElementById("reportBuilderEnd");
  const assetEl = document.getElementById("reportBuilderAssetId");
  const statusEl = document.getElementById("reportBuilderStatus");
  const limitEl = document.getElementById("reportBuilderLimit");
  if (dsEl) dsEl.value = String(tpl.dataset || "");
  if (nameEl) nameEl.value = String(tpl.name || "");
  if (startEl) startEl.value = String(tpl?.filters?.start || "");
  if (endEl) endEl.value = String(tpl?.filters?.end || "");
  if (assetEl) assetEl.value = String(Number(tpl?.filters?.asset_id || 0) || "");
  if (statusEl) statusEl.value = String(tpl?.filters?.status || "");
  if (limitEl) limitEl.value = String(Number(tpl?.filters?.limit || 100) || 100);
  const ds = reportBuilderMeta.find((d) => String(d.key || "") === String(tpl.dataset || ""));
  renderReportBuilderColumns(ds?.columns || [], Array.isArray(tpl.columns) ? tpl.columns : []);
}

async function runReportBuilderPreview() {
  const msg = document.getElementById("reportBuilderMsg");
  if (!msg) return;
  const cfg = reportBuilderCurrentConfig();
  if (!cfg.dataset) {
    msg.className = "message-error";
    msg.textContent = "Select dataset.";
    return;
  }
  if (!cfg.columns.length) {
    msg.className = "message-error";
    msg.textContent = "Select at least one column.";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Running preview...";
  try {
    const res = await fetch(`${API}/reports/custom-builder/preview`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(cfg),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Preview failed");
    renderReportBuilderPreview(data);
    msg.className = "message-success";
    msg.textContent = `Preview ready (${Number(data?.count || 0)} rows).`;
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = `Preview error: ${e.message || e}`;
  }
}

async function exportReportBuilderXlsx() {
  const msg = document.getElementById("reportBuilderMsg");
  if (!msg) return;
  const cfg = reportBuilderCurrentConfig();
  if (!cfg.dataset) {
    msg.className = "message-error";
    msg.textContent = "Select dataset.";
    return;
  }
  if (!cfg.columns.length) {
    msg.className = "message-error";
    msg.textContent = "Select at least one column.";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Generating XLSX...";
  try {
    const res = await fetch(`${API}/reports/custom-builder/export.xlsx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(cfg),
    });
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || "XLSX export failed");
    }
    const blob = await res.blob();
    const now = new Date().toISOString().slice(0, 10);
    const safeDataset = String(cfg.dataset || "report").replace(/[^a-z0-9_-]/gi, "_");
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `IRONLOG_Custom_${safeDataset}_${now}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    msg.className = "message-success";
    msg.textContent = "XLSX export downloaded.";
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = `XLSX export error: ${e.message || e}`;
  }
}

async function saveReportBuilderTemplate() {
  const msg = document.getElementById("reportBuilderMsg");
  const tplId = Number(document.getElementById("reportBuilderTemplate")?.value || 0);
  if (!msg) return;
  const cfg = reportBuilderCurrentConfig();
  if (!cfg.name) {
    msg.className = "message-error";
    msg.textContent = "Template name is required.";
    return;
  }
  if (!cfg.dataset || !cfg.columns.length) {
    msg.className = "message-error";
    msg.textContent = "Select dataset and at least one column.";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Saving template...";
  try {
    const res = await fetch(`${API}/reports/custom-builder/templates`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id: tplId || undefined, ...cfg }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Save failed");
    await loadReportBuilderTemplates();
    const sel = document.getElementById("reportBuilderTemplate");
    if (sel && Number(data?.id || 0) > 0) sel.value = String(Number(data.id));
    msg.className = "message-success";
    msg.textContent = "Template saved.";
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = `Save error: ${e.message || e}`;
  }
}

async function deleteReportBuilderTemplate() {
  const msg = document.getElementById("reportBuilderMsg");
  const sel = document.getElementById("reportBuilderTemplate");
  const id = Number(sel?.value || 0);
  if (!msg || !sel) return;
  if (!id) {
    msg.className = "message-error";
    msg.textContent = "Select template to delete.";
    return;
  }
  if (!confirm("Delete selected report template?")) return;
  msg.className = "muted";
  msg.textContent = "Deleting template...";
  try {
    const res = await fetch(`${API}/reports/custom-builder/templates/${id}`, { method: "DELETE" });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Delete failed");
    await loadReportBuilderTemplates();
    msg.className = "message-success";
    msg.textContent = "Template deleted.";
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = `Delete error: ${e.message || e}`;
  }
}

function subscriptionScheduleLabel(s) {
  const f = String(s.schedule_frequency || "weekly");
  const t = String(s.send_time || "07:00");
  if (f === "daily") return `daily @ ${t}`;
  if (f === "monthly") return `monthly day ${Number(s.day_of_month || 1)} @ ${t}`;
  return `weekly day ${Number(s.day_of_week || 1)} @ ${t}`;
}

function fillSubscriptionForm(s) {
  reportSubscriptionEditId = Number(s?.id || 0);
  const set = (id, v) => {
    const el = document.getElementById(id);
    if (!el) return;
    if (el.type === "checkbox") el.checked = Boolean(v);
    else el.value = v == null ? "" : String(v);
  };
  set("subName", s?.name || "");
  set("subReportType", s?.report_type || "fuel_benchmark_xlsx");
  set("subChannel", s?.channel || "email");
  set("subAttachFormat", s?.filters?.attach_format || "pdf");
  set("subRecipients", Array.isArray(s?.recipients) ? s.recipients.join(",") : "");
  set("subFrequency", s?.schedule_frequency || "weekly");
  set("subSendTime", s?.send_time || "07:00");
  set("subDayOfWeek", Number(s?.day_of_week ?? 1));
  set("subDayOfMonth", Number(s?.day_of_month ?? 1));
  set("subActive", Number(s?.active ?? 1) === 1);
  const startEl = document.getElementById("insightsStart");
  const endEl = document.getElementById("insightsEnd");
  if (s?.filters?.start && startEl) startEl.value = String(s.filters.start);
  if (s?.filters?.end && endEl) endEl.value = String(s.filters.end);
  const setNum = (id, v) => {
    const el = document.getElementById(id);
    if (el && v != null && v !== "") el.value = String(v);
  };
  setNum("insightsNearDueHours", s?.filters?.near_due_hours);
  setNum("insightsPredictiveHorizonHours", s?.filters?.predictive_horizon_hours);
  setNum("insightsChecklistFailThreshold", s?.filters?.checklist_fail_threshold);
  setNum("insightsFuelVarianceThreshold", s?.filters?.fuel_variance_threshold);
}

async function loadReportSubscriptionLogs() {
  const body = document.getElementById("subLogBody");
  if (!body) return;
  try {
    const res = await fetch(`${API}/reports/subscriptions/logs?limit=40`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load logs");
    const rows = Array.isArray(data?.rows) ? data.rows : [];
    body.innerHTML = rows.length
      ? rows.map((r) => `
        <tr>
          <td>${escBackfill(String(r.created_at || "-"))}</td>
          <td>${Number(r.subscription_id || 0)}</td>
          <td>${escBackfill(String(r.report_type || "-"))}</td>
          <td>${escBackfill(String(r.channel || "-"))}</td>
          <td>${escBackfill(String(r.status || "-"))}</td>
          <td>${escBackfill(String(r.detail || ""))}</td>
        </tr>
      `).join("")
      : `<tr><td colspan="6" class="muted">No delivery logs yet.</td></tr>`;
  } catch (e) {
    body.innerHTML = `<tr><td colspan="6" class="message-error">${escBackfill(e.message || e)}</td></tr>`;
  }
}

async function loadReportSubscriptions() {
  const body = document.getElementById("subBody");
  const msg = document.getElementById("subMsg");
  if (!body || !msg) return;
  body.innerHTML = `<tr><td colspan="8" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/reports/subscriptions`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load subscriptions");
    const rows = Array.isArray(data?.subscriptions) ? data.subscriptions : [];
    reportSubscriptionsCache = rows.slice();
    body.innerHTML = rows.length
      ? rows.map((s) => `
        <tr>
          <td>${escBackfill(String(s.name || ""))}${Number(s.active || 0) ? "" : ` <small class="muted">(paused)</small>`}</td>
          <td>${escBackfill(String(s.report_type || ""))}</td>
          <td>${escBackfill(String(s.channel || ""))}</td>
          <td>${escBackfill((Array.isArray(s.recipients) ? s.recipients : []).join(", "))}</td>
          <td>${escBackfill(subscriptionScheduleLabel(s))}</td>
          <td>${escBackfill(String(s.next_run_at || "-"))}</td>
          <td>${escBackfill(String(s.last_sent_at || "-"))}</td>
          <td>
            <button type="button" data-sub-edit="${Number(s.id || 0)}">Edit</button>
            <button type="button" data-sub-send="${Number(s.id || 0)}">Send Now</button>
            <button type="button" data-sub-del="${Number(s.id || 0)}">Delete</button>
          </td>
        </tr>
      `).join("")
      : `<tr><td colspan="8" class="muted">No subscriptions saved.</td></tr>`;
    msg.className = "message-success";
    msg.textContent = `Subscriptions loaded: ${rows.length}`;
    await loadReportSubscriptionLogs();
  } catch (e) {
    body.innerHTML = `<tr><td colspan="8" class="message-error">${escBackfill(e.message || e)}</td></tr>`;
    msg.className = "message-error";
    msg.textContent = `Subscriptions error: ${e.message || e}`;
  }
}

function subscriptionInsightsThresholds() {
  const nearEl = document.getElementById("insightsNearDueHours");
  const horizonEl = document.getElementById("insightsPredictiveHorizonHours");
  const failEl = document.getElementById("insightsChecklistFailThreshold");
  const fuelEl = document.getElementById("insightsFuelVarianceThreshold");
  const saved = getInsightsThresholds();
  return {
    near_due_hours: Math.max(1, Number(nearEl?.value || saved.near_due_hours || 50)),
    predictive_horizon_hours: Math.max(
      1,
      Number(horizonEl?.value || saved.predictive_horizon_hours || 100),
    ),
    checklist_fail_threshold: Math.max(1, Number(failEl?.value || saved.checklist_fail_threshold || 2)),
    fuel_variance_threshold: Math.max(0, Number(fuelEl?.value || saved.fuel_variance_threshold || 15)),
  };
}

function subscriptionPayloadFromForm() {
  const isChecked = (id) => Boolean(document.getElementById(id)?.checked);
  const value = (id) => String(document.getElementById(id)?.value || "").trim();
  const thresholds = subscriptionInsightsThresholds();
  const start = String(document.getElementById("insightsStart")?.value || "").trim();
  const end = String(document.getElementById("insightsEnd")?.value || "").trim();
  const reportType = value("subReportType");
  const rollingReport = reportType === "maintenance_insights_xlsx" || reportType === "fuel_benchmark_xlsx";
  return {
    id: reportSubscriptionEditId || undefined,
    name: value("subName"),
    report_type: reportType,
    channel: value("subChannel"),
    recipients: value("subRecipients"),
    schedule_frequency: value("subFrequency"),
    send_time: value("subSendTime") || "07:00",
    day_of_week: Number(value("subDayOfWeek") || 1),
    day_of_month: Number(value("subDayOfMonth") || 1),
    active: isChecked("subActive") ? 1 : 0,
    filters: {
      start,
      end,
      period_mode: rollingReport ? "rolling" : "fixed",
      site_codes: String(document.getElementById("kpiPackSiteCodes")?.value || "main").trim() || "main",
      period_type: String(document.getElementById("kpiPackPeriodType")?.value || "weekly").trim().toLowerCase(),
      attach_format: value("subAttachFormat") || "pdf",
      ...thresholds,
    },
  };
}

async function saveReportSubscription() {
  const msg = document.getElementById("subMsg");
  if (!msg) return;
  msg.className = "muted";
  msg.textContent = "Saving subscription...";
  try {
    const payload = subscriptionPayloadFromForm();
    const res = await fetch(`${API}/reports/subscriptions`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify(payload),
    });
    let data = {};
    try {
      data = await res.json();
    } catch {
      throw new Error(`Save failed (${res.status})`);
    }
    if (!res.ok) throw new Error(data.error || data.message || "Failed to save subscription");
    reportSubscriptionEditId = Number(data?.id || 0);
    msg.className = "message-success";
    msg.textContent = "Subscription saved.";
    await loadReportSubscriptions();
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = `Save failed: ${e.message || e}`;
  }
}

async function sendSubscriptionNow(id) {
  const msg = document.getElementById("subMsg");
  if (!msg) return;
  msg.className = "muted";
  msg.textContent = "Sending subscription...";
  try {
    const res = await fetch(`${API}/reports/subscriptions/${Number(id || 0)}/send-now`, { method: "POST" });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Send failed");
    msg.className = "message-success";
    msg.textContent = data.attachment_count
      ? `Sent: ${data.status || "ok"} (${data.attachment_count} attachment${data.attachment_count === 1 ? "" : "s"})`
      : `Sent: ${data.status || "ok"}`;
    await loadReportSubscriptions();
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = `Send failed: ${e.message || e}`;
  }
}

async function deleteSubscription(id) {
  const msg = document.getElementById("subMsg");
  if (!msg) return;
  if (!confirm("Delete this subscription?")) return;
  msg.className = "muted";
  msg.textContent = "Deleting subscription...";
  try {
    const res = await fetch(`${API}/reports/subscriptions/${Number(id || 0)}`, { method: "DELETE" });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Delete failed");
    msg.className = "message-success";
    msg.textContent = "Subscription deleted.";
    await loadReportSubscriptions();
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = `Delete failed: ${e.message || e}`;
  }
}
