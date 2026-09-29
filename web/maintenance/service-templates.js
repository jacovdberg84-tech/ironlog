// IRONLOG/web/maintenance/service-templates.js — Service templates and service planner.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

let serviceTemplatesCache = [];
let servicePlannerCache = [];
let serviceTemplateUiBound = false;

function serviceTemplateMessage(id, message, isError = false) {
  const target = document.getElementById(id);
  if (!target) return;
  target.className = isError ? "message-error" : "message-success";
  target.textContent = message || "";
}

function moneyLabel(value, complete = true) {
  if (!complete || value == null || !Number.isFinite(Number(value))) return "Price required";
  return `$${Number(value).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function populateServiceTemplateSelects(assets = []) {
  const templateSelect = document.getElementById("serviceTemplateAssignmentTemplate");
  if (templateSelect) {
    const selected = String(templateSelect.value || "");
    templateSelect.innerHTML = `<option value="">Select a template…</option>${serviceTemplatesCache
      .filter((template) => template.active)
      .map((template) => `<option value="${Number(template.id)}">${esc(template.template_key)} r${Number(template.revision_number)} — ${esc(template.name)}</option>`)
      .join("")}`;
    if ([...templateSelect.options].some((option) => option.value === selected)) templateSelect.value = selected;
  }
  const assetSelect = document.getElementById("serviceTemplateAssignmentAsset");
  if (assetSelect && assets.length) {
    const selected = String(assetSelect.value || "");
    assetSelect.innerHTML = `<option value="">No exact asset</option>${assets.map((asset) =>
      `<option value="${Number(asset.id)}">${esc(asset.asset_code || "")} — ${esc(asset.asset_name || "")}</option>`
    ).join("")}`;
    if ([...assetSelect.options].some((option) => option.value === selected)) assetSelect.value = selected;
  }
}

async function loadServiceTemplates() {
  const list = document.getElementById("serviceTemplatesList");
  try {
    const [templatesResponse, assetsResponse] = await Promise.all([
      fetch(`${API}/maintenance/service-templates`, { headers: authHeaders(), cache: "no-store" }),
      fetch(`${API}/assets?include_archived=0`, { headers: authHeaders(), cache: "no-store" }),
    ]);
    const templatesData = await templatesResponse.json();
    const assetsData = await assetsResponse.json();
    if (!templatesResponse.ok) throw new Error(templatesData.error || "Unable to load templates");
    serviceTemplatesCache = Array.isArray(templatesData.templates) ? templatesData.templates : [];
    const assets = Array.isArray(assetsData.assets) ? assetsData.assets : (Array.isArray(assetsData) ? assetsData : []);
    populateServiceTemplateSelects(assets);
    if (list) {
      list.innerHTML = serviceTemplatesCache.length ? serviceTemplatesCache.map((template) => `
        <tr>
          <td>${esc(template.template_key || "-")}</td>
          <td>r${Number(template.revision_number || 1)}</td>
          <td><strong>${esc(template.name || "-")}</strong><br><span class="muted mini">${esc(template.manufacturer || template.asset_category || "General")}</span></td>
          <td>${Number(template.service_interval_hours || 0).toLocaleString()} ${esc(template.meter_unit || "hours")}</td>
          <td>${Number(template.item_count || 0)}</td>
          <td>${Number(template.assignment_count || 0)}</td>
          <td>${template.active ? "Active" : "Superseded"}</td>
        </tr>`).join("") : `<tr><td colspan="7" class="muted">No service templates yet. Create one from a confirmed service kit or OEM list.</td></tr>`;
    }
  } catch (error) {
    if (list) list.innerHTML = `<tr><td colspan="7" class="message-error">${esc(error.message || String(error))}</td></tr>`;
  }
}

function renderServicePlanner() {
  const list = document.getElementById("servicePlannerList");
  const message = document.getElementById("servicePlannerMessage");
  if (!list) return;
  if (!servicePlannerCache.length) {
    list.innerHTML = `<tr><td colspan="9" class="muted">No due or near-due services match this view.</td></tr>`;
    if (message) message.textContent = "No planner actions are needed at this threshold.";
    return;
  }
  list.innerHTML = servicePlannerCache.map((row) => {
    const estimate = row.latest_estimate;
    const template = row.template_status === "matched"
      ? `${esc(row.template_name || "Matched template")} r${Number(row.template_revision || 1)}`
      : row.template_status === "ambiguous" ? "Assignment needs review" : "No matching template";
    const stock = row.template_status === "matched"
      ? (row.stock_available ? "Available" : "Short / review")
      : "—";
    let action = `<button type="button" data-service-planner-action="estimate" data-plan-id="${Number(row.plan_id)}" ${row.template_status === "matched" ? "" : "disabled"}>Create estimate</button>`;
    if (estimate?.status === "draft") action = `<button type="button" data-service-planner-action="approve" data-estimate-id="${Number(estimate.id)}" ${Number(estimate.pricing_complete) === 1 ? "" : "disabled title=\"Price required before approval\""}>Approve estimate</button>`;
    if (estimate?.status === "approved") action = `<button type="button" data-service-planner-action="convert" data-estimate-id="${Number(estimate.id)}">Create work order</button>`;
    if (estimate?.status === "converted") action = `Work order #${Number(estimate.work_order_id || 0) || "created"}`;
    return `<tr>
      <td><strong>${esc(row.asset_code || "-")}</strong><br><span class="muted mini">${esc(row.asset_name || "")}</span></td>
      <td>${esc(row.service_name || "Service")}<br><span class="muted mini">${Number(row.current_hours || 0).toFixed(1)} ${esc(row.meter_unit || "hours")}</span></td>
      <td>${Number(row.remaining_hours || 0).toFixed(1)} remaining<br><span class="muted mini">${esc(row.status || "OK")}</span></td>
      <td>${template}</td>
      <td>${moneyLabel(row.estimate_total, Boolean(row.pricing_complete))}</td>
      <td>${stock}</td>
      <td>${estimate ? esc(estimate.status) : "Not estimated"}</td>
      <td><span class="muted mini">${esc(row.borris_advisory?.action || "review")} · ${esc(row.borris_advisory?.confidence || "needs review")}</span></td>
      <td>${action}</td>
    </tr>`;
  }).join("");
  if (message) message.textContent = `${servicePlannerCache.length} service(s) shown. Borris is advisory only and cannot create, approve, reserve or issue anything.`;
}

async function loadServicePlanner() {
  const list = document.getElementById("servicePlannerList");
  if (list) list.innerHTML = `<tr><td colspan="9" class="muted">Loading costed planner…</td></tr>`;
  try {
    const includeAll = document.getElementById("servicePlannerIncludeAll")?.checked ? "1" : "0";
    const response = await fetch(`${API}/maintenance/service-planner?include_all=${includeAll}`, { headers: authHeaders(), cache: "no-store" });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Unable to load service planner");
    servicePlannerCache = Array.isArray(data.rows) ? data.rows : [];
    renderServicePlanner();
  } catch (error) {
    if (list) list.innerHTML = `<tr><td colspan="9" class="message-error">${esc(error.message || String(error))}</td></tr>`;
  }
}

function parseServiceTemplateItems() {
  return String(document.getElementById("serviceTemplateItems")?.value || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [description = "", quantity = "0", unit_of_measure = "ea", item_type = "part", stock_part_code = ""] = line.split("|").map((value) => value.trim());
      return { description, quantity_required: Number(quantity), unit_of_measure, item_type, stock_part_code };
    });
}

async function saveServiceTemplate() {
  try {
    const payload = {
      template_key: String(document.getElementById("serviceTemplateKey")?.value || "").trim(),
      name: String(document.getElementById("serviceTemplateName")?.value || "").trim(),
      service_interval_hours: Number(document.getElementById("serviceTemplateInterval")?.value || 0),
      meter_unit: String(document.getElementById("serviceTemplateMeter")?.value || "hours"),
      default_labour_hours: Number(document.getElementById("serviceTemplateLabourHours")?.value || 0),
      default_labour_rate: Number(document.getElementById("serviceTemplateLabourRate")?.value || 0),
      manufacturer: String(document.getElementById("serviceTemplateMake")?.value || "").trim(),
      model: String(document.getElementById("serviceTemplateModel")?.value || "").trim(),
      asset_category: String(document.getElementById("serviceTemplateCategory")?.value || "").trim(),
      description: String(document.getElementById("serviceTemplateDescription")?.value || "").trim(),
      items: parseServiceTemplateItems(),
    };
    const response = await fetch(`${API}/maintenance/service-templates`, {
      method: "POST", headers: authHeaders({ "Content-Type": "application/json" }), body: JSON.stringify(payload),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Unable to save template");
    serviceTemplateMessage("serviceTemplateMessage", `Template ${payload.template_key} saved as revision 1.`);
    ["serviceTemplateKey", "serviceTemplateName", "serviceTemplateInterval", "serviceTemplateDescription", "serviceTemplateItems"].forEach((id) => {
      const field = document.getElementById(id); if (field) field.value = "";
    });
    await Promise.all([loadServiceTemplates(), loadServicePlanner()]);
  } catch (error) {
    serviceTemplateMessage("serviceTemplateMessage", error.message || String(error), true);
  }
}

async function saveServiceTemplateAssignment() {
  try {
    const templateId = Number(document.getElementById("serviceTemplateAssignmentTemplate")?.value || 0);
    if (!templateId) throw new Error("Select a template first");
    const payload = {
      asset_id: Number(document.getElementById("serviceTemplateAssignmentAsset")?.value || 0) || null,
      manufacturer: String(document.getElementById("serviceTemplateAssignmentMake")?.value || "").trim(),
      model: String(document.getElementById("serviceTemplateAssignmentModel")?.value || "").trim(),
      asset_category: String(document.getElementById("serviceTemplateAssignmentCategory")?.value || "").trim(),
      priority: Number(document.getElementById("serviceTemplateAssignmentPriority")?.value || 100),
    };
    const response = await fetch(`${API}/maintenance/service-templates/${templateId}/assignments`, {
      method: "POST", headers: authHeaders({ "Content-Type": "application/json" }), body: JSON.stringify(payload),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Unable to save assignment");
    serviceTemplateMessage("serviceTemplateAssignmentMessage", "Template assignment saved.");
    await Promise.all([loadServiceTemplates(), loadServicePlanner()]);
  } catch (error) {
    serviceTemplateMessage("serviceTemplateAssignmentMessage", error.message || String(error), true);
  }
}

async function runServicePlannerAction(button) {
  const action = String(button?.dataset?.servicePlannerAction || "");
  const planId = Number(button?.dataset?.planId || 0);
  const estimateId = Number(button?.dataset?.estimateId || 0);
  try {
    button.disabled = true;
    let response;
    if (action === "estimate") {
      response = await fetch(`${API}/maintenance/service-estimates`, {
        method: "POST", headers: authHeaders({ "Content-Type": "application/json" }), body: JSON.stringify({ plan_id: planId }),
      });
    } else if (action === "approve") {
      response = await fetch(`${API}/maintenance/service-estimates/${estimateId}/approve`, {
        method: "POST", headers: authHeaders({ "Content-Type": "application/json" }), body: "{}",
      });
    } else if (action === "convert") {
      const reserveStock = window.confirm("Reserve available planned materials for this approved work order? Select Cancel to create the work order without reserving stock.");
      response = await fetch(`${API}/maintenance/service-estimates/${estimateId}/convert`, {
        method: "POST", headers: authHeaders({ "Content-Type": "application/json" }), body: JSON.stringify({ reserve_stock: reserveStock }),
      });
    } else return;
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Planner action failed");
    await loadServicePlanner();
  } catch (error) {
    alert(error.message || String(error));
  } finally {
    if (button?.isConnected) button.disabled = false;
  }
}

let templateProposalsCache = [];

function proposalSourceLabel(p) {
  if (p.source === "history") return `${p.history_services} past service${p.history_services === 1 ? "" : "s"}`;
  const oils = p.oils_from ? `, oils as on ${p.oils_from}` : " — add oils";
  if (p.source === "kit_code") return `Service kit by stock code${oils}`;
  if (p.source === "kit_model") return `Model service kit${oils}`;
  return "No history or kit found";
}

function renderTemplateProposals() {
  const list = document.getElementById("templateProposalsList");
  if (!list) return;
  if (!templateProposalsCache.length) {
    list.innerHTML = `<tr><td colspan="8" class="muted">No active maintenance plans found.</td></tr>`;
    return;
  }
  list.innerHTML = templateProposalsCache.map((p) => {
    const usable = p.items.length > 0;
    const tick = usable && !(p.existing && p.existing.asset_specific);
    const items = usable
      ? p.items.map((it) => `${esc(it.part_code)} — ${esc(it.description)} × ${Number(it.quantity_required).toLocaleString()} ${esc(it.unit_of_measure)}`).join("<br>")
      : `<span class="muted">Nothing to propose — create this one by hand.</span>`;
    const missing = p.estimate.missing_prices.length ? `<br><span class="message-error mini">No store price: ${esc(p.estimate.missing_prices.join(", "))}</span>` : "";
    const current = p.existing
      ? (p.existing.ambiguous ? "Two templates tie — review" : `${esc(p.existing.name)}${p.existing.asset_specific ? "" : " <span class=\"muted mini\">(shared)</span>"}`)
      : `<span class="muted">None</span>`;
    return `<tr>
      <td><input type="checkbox" data-proposal-key="${esc(p.key)}"${usable ? "" : " disabled"}${tick ? " checked" : ""} aria-label="Create template for ${esc(p.asset_code)} ${p.interval}" /></td>
      <td><strong>${esc(p.asset_code)}</strong><br><span class="muted mini">${esc(p.asset_name || "")}</span></td>
      <td>${Number(p.interval).toLocaleString()} ${p.meter_unit === "km" ? "km" : "h"}<br><span class="muted mini">${esc(p.service_name || "")}</span></td>
      <td>${esc(proposalSourceLabel(p))}</td>
      <td class="mini">${items}</td>
      <td>${p.labour_hours} h × ${moneyLabel(p.labour_rate)}</td>
      <td>${usable ? moneyLabel(p.estimate.total, !p.estimate.missing_prices.length) : "—"}${missing}</td>
      <td>${current}</td>
    </tr>`;
  }).join("");
  syncTemplateProposalButtons();
}

function syncTemplateProposalButtons() {
  const ticked = document.querySelectorAll("#templateProposalsList input[data-proposal-key]:checked").length;
  const btn = document.getElementById("createTemplateProposalsBtn");
  if (btn) {
    btn.disabled = ticked === 0;
    btn.textContent = ticked ? `Create ${ticked} ticked template${ticked === 1 ? "" : "s"}` : "Create ticked templates";
  }
}

async function buildTemplateProposals() {
  const msg = document.getElementById("templateProposalsMessage");
  if (msg) msg.textContent = "Building proposals from service history…";
  try {
    const res = await fetch(`${API}/maintenance/service-templates/suggestions`, { headers: authHeaders(), cache: "no-store" });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Unable to build proposals");
    templateProposalsCache = Array.isArray(data.proposals) ? data.proposals : [];
    const fromHistory = templateProposalsCache.filter((p) => p.source === "history").length;
    const fromKit = templateProposalsCache.filter((p) => p.source === "kit_code" || p.source === "kit_model").length;
    const none = templateProposalsCache.filter((p) => !p.items.length).length;
    if (msg) msg.textContent = `${templateProposalsCache.length} machine services: ${fromHistory} from service history, ${fromKit} from a matching service kit, ${none} with nothing to propose. Machines that already have their own template are left unticked.`;
    renderTemplateProposals();
  } catch (error) {
    if (msg) msg.textContent = error.message || String(error);
  }
}

async function createTemplateProposals() {
  const keys = Array.from(document.querySelectorAll("#templateProposalsList input[data-proposal-key]:checked")).map((el) => el.dataset.proposalKey);
  if (!keys.length) return;
  if (!confirm(`Create ${keys.length} service template${keys.length === 1 ? "" : "s"}? A machine's current template for the same interval is replaced (kept in history).`)) return;
  const msg = document.getElementById("templateProposalsMessage");
  try {
    const res = await fetch(`${API}/maintenance/service-templates/suggestions/apply`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ keys }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Unable to create templates");
    await Promise.all([loadServiceTemplates(), loadServicePlanner(), buildTemplateProposals()]);
    if (msg) msg.textContent = `Created ${data.created.length} template${data.created.length === 1 ? "" : "s"}. Service costs now use them. ${msg.textContent}`;
  } catch (error) {
    if (msg) msg.textContent = error.message || String(error);
  }
}

async function initServiceTemplateSection() {
  if (!serviceTemplateUiBound) {
    serviceTemplateUiBound = true;
    document.getElementById("refreshServicePlannerBtn")?.addEventListener("click", () => loadServicePlanner());
    document.getElementById("servicePlannerIncludeAll")?.addEventListener("change", () => loadServicePlanner());
    document.getElementById("saveServiceTemplateBtn")?.addEventListener("click", () => saveServiceTemplate());
    document.getElementById("buildTemplateProposalsBtn")?.addEventListener("click", () => buildTemplateProposals());
    document.getElementById("createTemplateProposalsBtn")?.addEventListener("click", () => createTemplateProposals());
    document.getElementById("templateProposalsList")?.addEventListener("change", () => syncTemplateProposalButtons());
    document.getElementById("templateProposalsAll")?.addEventListener("change", (event) => {
      document.querySelectorAll("#templateProposalsList input[data-proposal-key]:not(:disabled)").forEach((el) => { el.checked = event.target.checked; });
      syncTemplateProposalButtons();
    });
    document.getElementById("saveServiceTemplateAssignmentBtn")?.addEventListener("click", () => saveServiceTemplateAssignment());
    document.getElementById("servicePlannerList")?.addEventListener("click", (event) => {
      const button = event.target instanceof HTMLElement ? event.target.closest("button[data-service-planner-action]") : null;
      if (button) runServicePlannerAction(button);
    });
  }
  await Promise.all([loadServiceTemplates(), loadServicePlanner()]);
}
