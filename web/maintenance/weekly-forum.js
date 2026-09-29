// IRONLOG/web/maintenance/weekly-forum.js — Weekly maintenance forum: summary, drafts, actions, reviews and inputs.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

function weeklyForumQueryString() {
  const start = String(document.getElementById("wfStart")?.value || "").trim();
  const end = String(document.getElementById("wfEnd")?.value || "").trim();
  const near = Math.max(1, Number(document.getElementById("wfNearDueHours")?.value || 50));
  const q = new URLSearchParams();
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  q.set("near_due_hours", String(near));
  return q.toString();
}

/** Sets wfStart / wfEnd to a full calendar month (this month or previous). */
function wfApplyMonthPreset(which) {
  const sEl = document.getElementById("wfStart");
  const eEl = document.getElementById("wfEnd");
  if (!sEl || !eEl) return;
  const now = new Date();
  let y = now.getFullYear();
  let m = now.getMonth();
  if (which === "last") {
    m -= 1;
    if (m < 0) {
      m = 11;
      y -= 1;
    }
  }
  const start = new Date(y, m, 1);
  const end = new Date(y, m + 1, 0);
  sEl.value = start.toISOString().slice(0, 10);
  eEl.value = end.toISOString().slice(0, 10);
  const mpStart = document.getElementById("mpWeekStart");
  const mpEnd = document.getElementById("mpWeekEnd");
  if (mpStart) mpStart.value = sEl.value;
  if (mpEnd) mpEnd.value = eEl.value;
  loadWeeklyForumReviews().catch(() => {});
}

let wfUpcomingCache = [];
let wfInputsCache = [];
let wfPartsCache = [];
let wfDraftItems = [];

function wfPlanLabel(r) {
  return `${String(r.asset_code || "-")} - ${String(r.asset_name || "-")} | ${String(r.service_name || "-")} (Plan ${Number(r.plan_id || 0)})`;
}
function refreshWeeklyForumPlanOptions() {
  const sel = document.getElementById("wfInputPlan");
  if (!sel) return;
  sel.innerHTML = wfUpcomingCache.length
    ? `<option value="">Select upcoming service</option>${wfUpcomingCache.map((r) => `<option value="${Number(r.plan_id || 0)}">${esc(wfPlanLabel(r))}</option>`).join("")}`
    : `<option value="">No upcoming services loaded</option>`;
}
function refreshWeeklyForumInputsTable() {
  const body = document.getElementById("wfInputsBody");
  if (!body) return;
  body.innerHTML = wfInputsCache.length
    ? wfInputsCache.map((r) => {
        const plan = wfUpcomingCache.find((p) => Number(p.plan_id || 0) === Number(r.plan_id || 0));
        let items = [];
        try {
          const parsed = JSON.parse(String(r.items_json || "[]"));
          if (Array.isArray(parsed)) items = parsed;
        } catch {}
        if (!items.length) {
          const oilCode = String(r.oil_part_code || "").trim();
          const oilQty = Number(r.oil_qty || 0);
          const partCode = String(r.parts_part_code || "").trim();
          const partQty = Number(r.parts_qty || 0);
          if (oilCode && oilQty > 0) items.push({ type: "oil", part_code: oilCode, qty: oilQty });
          if (partCode && partQty > 0) items.push({ type: "part", part_code: partCode, qty: partQty });
        }
        const oils = items.filter((x) => String(x.type || "part").toLowerCase() === "oil");
        const parts = items.filter((x) => String(x.type || "part").toLowerCase() !== "oil");
        const render = (rows) => rows.length
          ? rows.map((x) => `${String(x.part_code || "")} (${fmt1(x.qty)})`).join(", ")
          : "-";
        return `
          <tr>
            <td>${esc(plan ? wfPlanLabel(plan) : `Plan ${Number(r.plan_id || 0)}`)}</td>
            <td>${esc(render(oils))}</td>
            <td>${esc(render(parts))}</td>
            <td style="text-align:right;">${fmtMoney(Number(r.all_in_total || 0))}</td>
            <td>${esc(r.notes || "")}</td>
          </tr>
        `;
      }).join("")
    : `<tr><td colspan="5" class="muted">No manual inputs saved.</td></tr>`;
}
function refreshWeeklyForumPartsDatalist() {
  const html = wfPartsCache.map((p) => {
    const code = String(p.part_code || "").trim();
    const desc = `${code} - ${String(p.part_name || "")} | on hand ${Number(p.on_hand || 0).toFixed(2)} | unit ${Number(p.latest_unit_cost || 0).toFixed(2)}`;
    return `<option value="${esc(code)}">${esc(desc)}</option>`;
  }).join("");
  const dl = document.getElementById("wfPartsList");
  if (dl) dl.innerHTML = html;
  const dlMi = document.getElementById("miPartsList");
  if (dlMi) dlMi.innerHTML = html;
}
function getWfPartByCode(codeIn) {
  const code = String(codeIn || "").trim().toUpperCase();
  if (!code) return null;
  return wfPartsCache.find((p) => String(p.part_code || "").trim().toUpperCase() === code) || null;
}

function refreshWfDraftEditor() {
  const body = document.getElementById("wfItemsEditorBody");
  const oilTotalEl = document.getElementById("wfOilTotal");
  const partsTotalEl = document.getElementById("wfPartsTotal");
  const serviceTotalEl = document.getElementById("wfServiceTotal");
  if (!body) return;
  body.innerHTML = wfDraftItems.length
    ? wfDraftItems.map((it, idx) => `
      <tr>
        <td>${esc(it.type === "oil" ? "Oil" : "Part")}</td>
        <td>${esc(it.part_code)}</td>
        <td>${esc(it.part_name || "-")}</td>
        <td style="text-align:right;">${fmt1(it.qty)}</td>
        <td style="text-align:right;">${fmtMoney(it.unit_cost)}</td>
        <td style="text-align:right;">${fmtMoney(it.line_cost)}</td>
        <td style="text-align:right;">${fmt1(it.on_hand)}</td>
        <td style="text-align:right;"><button type="button" data-wf-item-del="${idx}">Remove</button></td>
      </tr>
    `).join("")
    : `<tr><td colspan="8" class="muted">No items added yet.</td></tr>`;
  const oilTotal = wfDraftItems.filter((x) => x.type === "oil").reduce((s, x) => s + Number(x.line_cost || 0), 0);
  const partsTotal = wfDraftItems.filter((x) => x.type !== "oil").reduce((s, x) => s + Number(x.line_cost || 0), 0);
  const serviceTotal = oilTotal + partsTotal;
  if (oilTotalEl) oilTotalEl.textContent = fmtMoney(oilTotal);
  if (partsTotalEl) partsTotalEl.textContent = fmtMoney(partsTotal);
  if (serviceTotalEl) serviceTotalEl.textContent = fmtMoney(serviceTotal);
}
function hydrateWfDraftFromSaved(planId) {
  const row = wfInputsCache.find((r) => Number(r.plan_id || 0) === Number(planId || 0));
  if (!row) {
    wfDraftItems = [];
    const laborEl = document.getElementById("wfInputLabor");
    const allInEl = document.getElementById("wfInputAllInTotal");
    if (laborEl) laborEl.value = "0";
    if (allInEl) allInEl.value = "0";
    refreshWfDraftEditor();
    return;
  }
  const laborEl = document.getElementById("wfInputLabor");
  const allInEl = document.getElementById("wfInputAllInTotal");
  if (laborEl) laborEl.value = String(Number(row.labor_total || 0));
  if (allInEl) allInEl.value = String(Number(row.all_in_total || 0));
  let items = [];
  try {
    const parsed = JSON.parse(String(row.items_json || "[]"));
    if (Array.isArray(parsed)) items = parsed;
  } catch {}
  wfDraftItems = items.map((it) => {
    const part = getWfPartByCode(it.part_code);
    const qty = Math.max(0, Number(it.qty || 0));
    const unit = Number(part?.latest_unit_cost || 0);
    return {
      type: String(it.type || "part").toLowerCase() === "oil" ? "oil" : "part",
      part_code: String(it.part_code || "").trim(),
      part_name: String(part?.part_name || ""),
      qty,
      unit_cost: unit,
      on_hand: Number(part?.on_hand || 0),
      line_cost: qty * unit,
    };
  }).filter((x) => x.part_code && x.qty > 0);
  refreshWfDraftEditor();
}
function addWfDraftItem() {
  const msg = document.getElementById("wfInputMsg");
  const type = String(document.getElementById("wfItemType")?.value || "part").toLowerCase() === "oil" ? "oil" : "part";
  const part_code = String(document.getElementById("wfItemCode")?.value || "").trim();
  const qty = Math.max(0, Number(document.getElementById("wfItemQty")?.value || 0));
  if (!part_code || qty <= 0) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Select a part code and enter a quantity greater than 0.";
    }
    return;
  }
  const part = getWfPartByCode(part_code);
  if (!part) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Part code not found in Stores list. Please use a valid part code.";
    }
    return;
  }
  wfDraftItems.push({
    type,
    part_code: String(part.part_code || "").trim(),
    part_name: String(part.part_name || ""),
    qty,
    unit_cost: Number(part.latest_unit_cost || 0),
    on_hand: Number(part.on_hand || 0),
    line_cost: qty * Number(part.latest_unit_cost || 0),
  });
  const codeEl = document.getElementById("wfItemCode");
  const qtyEl = document.getElementById("wfItemQty");
  if (codeEl) codeEl.value = "";
  if (qtyEl) qtyEl.value = "0";
  if (msg) {
    msg.className = "muted";
    msg.textContent = "Item added.";
  }
  refreshWfDraftEditor();
}

async function loadWeeklyForumSummary() {
  const msg = document.getElementById("wfMsg");
  const kpiBody = document.getElementById("wfKpiBody");
  const upcomingBody = document.getElementById("wfUpcomingBody");
  const periodBody = document.getElementById("wfPeriodActualsBody");
  const startEl = document.getElementById("wfStart");
  const endEl = document.getElementById("wfEnd");
  const nearEl = document.getElementById("wfNearDueHours");
  if (!msg || !kpiBody || !upcomingBody || !periodBody || !startEl || !endEl || !nearEl) return;

  const q = weeklyForumQueryString();

  msg.className = "muted";
  msg.textContent = "Loading weekly forum data...";
  kpiBody.innerHTML = `<tr><td colspan="2" class="muted">Loading...</td></tr>`;
  periodBody.innerHTML = `<tr><td colspan="8" class="muted">Loading...</td></tr>`;
  upcomingBody.innerHTML = `<tr><td colspan="11" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/summary?${q}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load weekly forum summary");

    const kpis = data.kpis || {};
    const costs = data.costs || {};
    const range = data.range || {};
    kpiBody.innerHTML = [
      ["Range", `${range.start || "-"} to ${range.end || "-"}`],
      ["Open Work Orders", Number(kpis.open_work_orders || 0)],
      ["Upcoming Services Flagged", Number(kpis.upcoming_services_flagged || 0)],
      ["Stores parts (excl. oil/lube SKUs)", fmtMoney(costs.stores_parts_cost)],
      ["Oil cost — lube log entries", fmtMoney(costs.stores_oil_from_logs)],
      ["Oil cost — WO stock (oil/lube lines)", fmtMoney(costs.stores_oil_from_work_orders)],
      ["Stores oil total", fmtMoney(costs.stores_oil_cost)],
      ["Maintenance Labor Cost", fmtMoney(costs.maintenance_labor_cost)],
      ["Weekly Total Cost", fmtMoney(costs.weekly_total_cost)],
      ["Upcoming Service Forecast Cost", fmtMoney(costs.upcoming_service_forecast_cost)],
    ].map(([k, v]) => `<tr><td>${esc(k)}</td><td style="text-align:right;">${esc(String(v))}</td></tr>`).join("");

    const rows = Array.isArray(data.upcoming_services) ? data.upcoming_services : [];
    const planSel = document.getElementById("wfInputPlan");
    const prevPlan = Number(planSel?.value || 0);
    wfUpcomingCache = rows;
    refreshWeeklyForumPlanOptions();
    if (planSel && prevPlan && wfUpcomingCache.some((x) => Number(x.plan_id || 0) === prevPlan)) {
      planSel.value = String(prevPlan);
    }
    hydrateWfDraftFromSaved(Number(planSel?.value || 0));

    const actuals = Array.isArray(data.period_actuals_by_asset) ? data.period_actuals_by_asset : [];
    periodBody.innerHTML = actuals.length
      ? actuals
          .map(
            (r) => `
        <tr>
          <td>${esc(r.asset_code || "-")} - ${esc(r.asset_name || "-")}</td>
          <td style="text-align:right;">${fmtMoney(r.parts_cost)}</td>
          <td style="text-align:right;">${fmtMoney(r.lubes_logs_cost)}</td>
          <td style="text-align:right;">${fmtMoney(r.lubes_work_order_cost)}</td>
          <td style="text-align:right;">${fmtMoney(r.lubes_total_cost)}</td>
          <td style="text-align:right;">${fmtMoney(r.labor_cost)}</td>
          <td style="text-align:right;">${fmtMoney(r.period_total_cost)}</td>
          <td style="text-align:right;">${esc(String(r.closed_work_orders ?? 0))}</td>
        </tr>
      `
          )
          .join("")
      : `<tr><td colspan="8" class="muted">No equipment consumption in this range.</td></tr>`;

    upcomingBody.innerHTML = rows.length
      ? rows.map((r) => `
        <tr>
          <td>${esc(r.asset_code || "-")} - ${esc(r.asset_name || "-")}</td>
          <td>${esc(r.service_name || "-")}</td>
          <td style="text-align:right;">${fmt1(r.current_hours)}</td>
          <td style="text-align:right;">${fmt1(r.next_due_hours)}</td>
          <td style="text-align:right;">${fmt1(r.remaining_hours)}</td>
          <td>${esc(r.status || "-")}</td>
          <td style="text-align:right;">${fmt1(r?.forecast?.avg_oil_qty)}</td>
          <td style="text-align:right;">${fmtMoney(r?.forecast?.avg_oil_cost)}</td>
          <td style="text-align:right;">${fmt1(r?.forecast?.avg_parts_qty)}</td>
          <td style="text-align:right;">${fmtMoney(r?.forecast?.avg_parts_cost)}</td>
          <td style="text-align:right;">${fmtMoney(r?.forecast?.est_service_kit_cost)}</td>
        </tr>
      `).join("")
      : `<tr><td colspan="11" class="muted">No upcoming services within threshold.</td></tr>`;

    msg.className = "message-success";
    msg.textContent = "Weekly forum summary loaded.";
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = `Load error: ${e.message || e}`;
    kpiBody.innerHTML = `<tr><td colspan="2" class="message-error">${esc(e.message || String(e))}</td></tr>`;
    periodBody.innerHTML = `<tr><td colspan="8" class="message-error">${esc(e.message || String(e))}</td></tr>`;
    upcomingBody.innerHTML = `<tr><td colspan="11" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}


async function openWeeklyForumPdf(download = false) {
  const q = weeklyForumQueryString();
  const url = `${API}/maintenance/weekly-forum.pdf?${q}${download ? "&download=1" : ""}`;
  try {
    const res = await fetch(url, { headers: authHeaders() });
    if (!res.ok) {
      const txt = await res.text();
      throw new Error(txt || `PDF request failed (${res.status})`);
    }
    const blob = await res.blob();
    const blobUrl = URL.createObjectURL(blob);
    if (download) {
      const a = document.createElement("a");
      const dateTag = new Date().toISOString().slice(0, 10);
      a.href = blobUrl;
      a.download = `weekly-forum-${dateTag}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(blobUrl), 3000);
      return;
    }
    window.open(blobUrl, "_blank");
    setTimeout(() => URL.revokeObjectURL(blobUrl), 10000);
  } catch (e) {
    alert(`Weekly Forum PDF error: ${e.message || e}`);
  }
}

function wfStatusLabel(s) {
  const v = String(s || "open").toLowerCase();
  if (v === "in_progress") return "In Progress";
  if (v === "blocked") return "Blocked";
  if (v === "done") return "Done";
  return "Open";
}

async function loadWeeklyForumActions() {
  const body = document.getElementById("wfActionBody");
  const start = String(document.getElementById("wfStart")?.value || "").trim();
  const end = String(document.getElementById("wfEnd")?.value || "").trim();
  if (!body) return;
  body.innerHTML = `<tr><td colspan="8" class="muted">Loading...</td></tr>`;
  const q = new URLSearchParams();
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/actions?${q.toString()}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load actions");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    body.innerHTML = rows.length
      ? rows.map((r) => `
        <tr>
          <td>${Number(r.id || 0)}</td>
          <td>${esc(r.action_date || "-")}</td>
          <td>${esc(r.department || "-")}</td>
          <td>${esc(r.action_item || "-")}</td>
          <td>${esc(r.owner_name || "-")}</td>
          <td>${esc(r.due_date || "-")}</td>
          <td>
            <select data-wf-action-status="${Number(r.id || 0)}">
              <option value="open" ${String(r.status) === "open" ? "selected" : ""}>Open</option>
              <option value="in_progress" ${String(r.status) === "in_progress" ? "selected" : ""}>In Progress</option>
              <option value="blocked" ${String(r.status) === "blocked" ? "selected" : ""}>Blocked</option>
              <option value="done" ${String(r.status) === "done" ? "selected" : ""}>Done</option>
            </select>
          </td>
          <td>${esc(r.notes || "")}</td>
        </tr>
      `).join("")
      : `<tr><td colspan="8" class="muted">No actions logged for selected range.</td></tr>`;
  } catch (e) {
    body.innerHTML = `<tr><td colspan="8" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

async function saveWeeklyForumAction() {
  const msg = document.getElementById("wfActionMsg");
  const action_date = String(document.getElementById("wfActionDate")?.value || "").trim();
  const department = String(document.getElementById("wfActionDept")?.value || "").trim();
  const owner_name = String(document.getElementById("wfActionOwner")?.value || "").trim();
  const due_date = String(document.getElementById("wfActionDue")?.value || "").trim();
  const status = String(document.getElementById("wfActionStatus")?.value || "open").trim().toLowerCase();
  const action_item = String(document.getElementById("wfActionItem")?.value || "").trim();
  const notes = String(document.getElementById("wfActionNotes")?.value || "").trim();
  if (!msg) return;

  if (!action_date || !department || !owner_name || !action_item) {
    msg.className = "message-error";
    msg.textContent = "Action date, department, owner, and action item are required.";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Saving action...";
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/actions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        action_date,
        department,
        action_item,
        owner_name,
        due_date: due_date || null,
        status,
        notes: notes || null,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save action");
    msg.className = "message-success";
    msg.textContent = "Action saved.";
    document.getElementById("wfActionItem").value = "";
    document.getElementById("wfActionNotes").value = "";
    await loadWeeklyForumActions();
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = e.message || String(e);
  }
}

function weeklyForumReviewRange() {
  return {
    start: String(document.getElementById("wfStart")?.value || "").trim(),
    end: String(document.getElementById("wfEnd")?.value || "").trim(),
  };
}

async function loadWeeklyForumReviews() {
  const body = document.getElementById("wfReviewBody");
  if (!body) return;
  const { start, end } = weeklyForumReviewRange();
  if (!start || !end) {
    body.innerHTML = `<tr><td colspan="5" class="muted">Select a start and end date first.</td></tr>`;
    return;
  }
  body.innerHTML = `<tr><td colspan="5" class="muted">Loading...</td></tr>`;
  const q = new URLSearchParams({ start, end });
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/review-notes?${q.toString()}`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load weekly review notes");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    body.innerHTML = rows.length
      ? rows.map((r) => `
        <tr>
          <td>${esc(r.area || "-")}</td>
          <td>${esc(r.weekly_finding || "-")}</td>
          <td>${esc(r.action_owner || "-")}</td>
          <td>${esc(r.due_date || "-")}</td>
          <td><button type="button" class="btn btn-secondary" data-wf-review-delete="${Number(r.id || 0)}">Remove</button></td>
        </tr>
      `).join("")
      : `<tr><td colspan="5" class="muted">No weekly review inputs saved for this period yet.</td></tr>`;
  } catch (e) {
    body.innerHTML = `<tr><td colspan="5" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

async function saveWeeklyForumReview() {
  const msg = document.getElementById("wfReviewMsg");
  if (!msg) return;
  const { start, end } = weeklyForumReviewRange();
  const area = String(document.getElementById("wfReviewArea")?.value || "").trim();
  const weekly_finding = String(document.getElementById("wfReviewFinding")?.value || "").trim();
  const action_owner = String(document.getElementById("wfReviewActionOwner")?.value || "").trim();
  const due_date = String(document.getElementById("wfReviewDue")?.value || "").trim();
  if (!start || !end || !area || !weekly_finding || !action_owner) {
    msg.className = "message-error";
    msg.textContent = "Start date, end date, area, weekly finding, and action / owner are required.";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Saving weekly review...";
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/review-notes`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({ start, end, area, weekly_finding, action_owner, due_date: due_date || null }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save weekly review");
    msg.className = "message-success";
    msg.textContent = "Weekly review saved. It will appear in the next Weekly Forum presentation.";
    document.getElementById("wfReviewFinding").value = "";
    document.getElementById("wfReviewActionOwner").value = "";
    document.getElementById("wfReviewDue").value = "";
    await loadWeeklyForumReviews();
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = e.message || String(e);
  }
}

async function deleteWeeklyForumReview(id) {
  const n = Number(id || 0);
  if (!n) return;
  const msg = document.getElementById("wfReviewMsg");
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/review-notes/${n}`, {
      method: "DELETE",
      headers: authHeaders(),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to remove weekly review");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = "Weekly review removed.";
    }
    await loadWeeklyForumReviews();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = e.message || String(e);
    }
  }
}

async function updateWeeklyForumActionStatus(id, status) {
  const n = Number(id || 0);
  if (!n) return;
  await fetch(`${API}/maintenance/weekly-forum/actions/${n}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ status }),
  });
}


async function loadWeeklyForumParts() {
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/parts`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load parts list");
    wfPartsCache = Array.isArray(data.rows) ? data.rows : [];
    refreshWeeklyForumPartsDatalist();
    hydrateWfDraftFromSaved(Number(document.getElementById("wfInputPlan")?.value || 0));
  } catch {
    wfPartsCache = [];
    refreshWeeklyForumPartsDatalist();
  }
}

async function loadWeeklyForumInputs() {
  const body = document.getElementById("wfInputsBody");
  if (body) body.innerHTML = `<tr><td colspan="4" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/forecast-inputs`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load manual inputs");
    wfInputsCache = Array.isArray(data.rows) ? data.rows : [];
    refreshWeeklyForumInputsTable();
    const planNow = Number(document.getElementById("wfInputPlan")?.value || 0);
    if (planNow) hydrateWfDraftFromSaved(planNow);
  } catch (e) {
    if (body) body.innerHTML = `<tr><td colspan="5" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

async function saveWeeklyForumInput() {
  const msg = document.getElementById("wfInputMsg");
  const plan_id = Number(document.getElementById("wfInputPlan")?.value || 0);
  const notes = String(document.getElementById("wfInputNotes")?.value || "").trim();
  const labor_total = Math.max(0, Number(document.getElementById("wfInputLabor")?.value || 0));
  const all_in_total = Math.max(0, Number(document.getElementById("wfInputAllInTotal")?.value || 0));
  const items = wfDraftItems.map((x) => ({
    type: x.type === "oil" ? "oil" : "part",
    part_code: String(x.part_code || "").trim(),
    qty: Math.max(0, Number(x.qty || 0)),
  })).filter((x) => x.part_code && x.qty > 0);
  if (!msg) return;
  if (!plan_id) {
    msg.className = "message-error";
    msg.textContent = "Select an upcoming service plan first.";
    return;
  }
  if (!items.length && labor_total <= 0 && all_in_total <= 0) {
    msg.className = "message-error";
    msg.textContent = "Enter store items, labor, or an all-in planned cost.";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Saving manual input...";
  try {
    const res = await fetch(`${API}/maintenance/weekly-forum/forecast-inputs`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ plan_id, items, labor_total, all_in_total, notes: notes || null }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save input");
    msg.className = "message-success";
    msg.textContent = all_in_total > 0
      ? "All-in planned cost saved. The presentation pack will use it as the service total."
      : "Manual input saved. Forecast cost now uses stores pricing.";
    await loadWeeklyForumInputs();
    await loadWeeklyForumSummary();
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = e.message || String(e);
  }
}
