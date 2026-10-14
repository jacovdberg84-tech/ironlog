// IRONLOG/web/maintenance/care-maintenance.js — Care & maintenance plans.
// Part of maintenance.html; the page loads these files in order and they share one global scope.
//
// One plan per C&M period: every machine with the work it needs, the parts
// (priced) and the repair hours. Each machine gets a normal work order when
// work starts, and the plan shows hours and parts used from it next to the estimate.

const CM_EDIT_ROLES = new Set(["admin", "supervisor", "workshop_admin", "plant_manager", "site_manager", "workshop_manager"]);
const CM_STATUS_PILL = { not_started: "pill-gray", waiting_parts: "pill-orange", in_progress: "pill-blue", done: "pill-green", parked: "pill-gray" };
const CM_PART_STATE = {
  issued: ["Issued", "pill-green"],
  in_stock: ["In stores", "pill-green"],
  requested: ["Requested", "pill-orange"],
  ordered: ["Ordered", "pill-blue"],
  received: ["Received", "pill-green"],
  short: ["Not in stores", "pill-red"],
};
let cmPlans = [];
let cmDetail = null;
let cmOpenRow = 0;
let cmAssets = [];
let cmWired = false;

const cmCanEdit = () => getSessionRoles().some((r) => CM_EDIT_ROLES.has(r));
const cmMoney = (v) => `$${Number(v || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const cmHours = (v) => `${Number(v || 0).toLocaleString("en-US", { maximumFractionDigits: 1 })} h`;
const cmEl = (id) => document.getElementById(id);

async function cmFetch(path, opts = {}) {
  const init = { ...opts, headers: { ...(opts.body ? { "Content-Type": "application/json" } : {}), ...(opts.headers || {}) } };
  if (opts.body && typeof opts.body !== "string") init.body = JSON.stringify(opts.body);
  const res = await fetch(`${API}/maintenance/cm${path}`, init);
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.ok === false) throw new Error(data.error || (res.status === 403 ? "You don't have access to change this plan." : `Request failed (${res.status})`));
  return data;
}

function cmMsg(text, isError = false) {
  const el = cmEl("cmMsg");
  if (!el) return;
  el.textContent = text || "";
  el.className = isError ? "message-error" : "muted";
}

async function initCareMaintenanceSection() {
  if (!cmWired) wireCareMaintenance();
  document.querySelectorAll("#careMaintenanceSection .cm-edit-only").forEach((el) => el.classList.toggle("hidden", !cmCanEdit()));
  try {
    if (!cmAssets.length) {
      const res = await fetch(`${API}/assets?include_archived=0`);
      const rows = await res.json().catch(() => []);
      cmAssets = (Array.isArray(rows) ? rows : []).filter((a) => Number(a.active ?? 1) !== 0 && Number(a.archived ?? 0) === 0);
    }
    if (typeof loadWeeklyForumParts === "function" && !wfPartsCache.length) loadWeeklyForumParts().catch(() => {});
    await cmLoadPlans();
  } catch (e) {
    cmMsg(e.message || String(e), true);
  }
}

async function cmLoadPlans(selectId = 0) {
  const data = await cmFetch("/plans");
  cmPlans = data.plans || [];
  const sel = cmEl("cmPlanSelect");
  const keep = selectId || Number(sel?.value || 0) || cmPlans[0]?.id || 0;
  if (sel) {
    sel.innerHTML = cmPlans.length
      ? cmPlans.map((p) => `<option value="${p.id}">${esc(p.name)}${p.status === "closed" ? " (closed)" : ""} · ${p.machines} machines</option>`).join("")
      : `<option value="">No plans yet</option>`;
    if (keep) sel.value = String(keep);
  }
  cmEl("cmEmpty")?.classList.toggle("hidden", cmPlans.length > 0);
  cmEl("cmPlanBody")?.classList.toggle("hidden", cmPlans.length === 0);
  if (keep) await cmLoadPlan(keep);
}

async function cmLoadPlan(id = Number(cmEl("cmPlanSelect")?.value || 0)) {
  if (!id) return;
  cmDetail = await cmFetch(`/plans/${id}`);
  cmRender();
}

function cmRender() {
  if (!cmDetail) return;
  const { plan, totals, machines } = cmDetail;
  const dates = [plan.start_date, plan.end_date].filter(Boolean).join(" to ");
  cmEl("cmPlanInfo").textContent = `${dates ? `${dates} · ` : ""}Labour at ${cmMoney(plan.labour_rate_used)}/h${plan.status === "closed" ? " · Closed" : ""}`;
  const closeBtn = cmEl("cmClosePlanBtn");
  if (closeBtn) closeBtn.textContent = plan.status === "closed" ? "Reopen plan" : "Close plan";
  const pct = (a, b) => (b > 0 ? ` (${Math.round((a / b) * 100)}%)` : "");
  const done = totals.by_status.done || 0;
  cmEl("cmTiles").innerHTML = [
    ["Machines", `${totals.machines}`, `${done} done · ${totals.by_status.in_progress || 0} in progress · ${totals.by_status.waiting_parts || 0} waiting for parts`],
    ["Planned cost", cmMoney(totals.planned_total), `Parts ${cmMoney(totals.planned_parts)} · Labour ${cmMoney(totals.planned_labour)}`],
    ["Spent so far", cmMoney(totals.actual_total), `Parts ${cmMoney(totals.actual_parts)} · Labour ${cmMoney(totals.actual_labour)}${pct(totals.actual_total, totals.planned_total)}`],
    ["Repair hours", `${cmHours(totals.actual_hours)} of ${cmHours(totals.planned_hours)}`, `Hours booked against the plan${pct(totals.actual_hours, totals.planned_hours)}`],
    ["Parts still needed", `${totals.parts_short}`, totals.parts_short ? "Lines not yet in stores or not yet issued" : "Everything is in stores or issued"],
  ].map(([label, value, sub]) => `<div class="cm-tile"><div class="cm-tile-label">${esc(label)}</div><div class="cm-tile-value">${esc(value)}</div><div class="cm-tile-sub">${esc(sub)}</div></div>`).join("");

  // Machines that can still be added.
  const onPlan = new Set(machines.map((m) => m.asset_id));
  const cat = cmEl("cmCategory")?.value || "";
  const cats = [...new Set(cmAssets.map((a) => String(a.category || "").trim()).filter(Boolean))].sort();
  const catSel = cmEl("cmCategory");
  if (catSel) {
    catSel.innerHTML = `<option value="">All types</option>${cats.map((c) => `<option value="${esc(c)}">${esc(c)}</option>`).join("")}`;
    catSel.value = cats.includes(cat) ? cat : "";
  }
  cmRenderAssetOptions(onPlan);

  const filter = cmEl("cmStatusFilter")?.value || "";
  const list = machines.filter((m) => !filter || m.status === filter);
  const body = cmEl("cmMachinesBody");
  if (!machines.length) {
    body.innerHTML = `<tr><td colspan="8" class="muted">No machines on this plan yet. Add them above.</td></tr>`;
    return;
  }
  body.innerHTML = list.map((m) => cmMachineRow(m)).join("") || `<tr><td colspan="8" class="muted">No machines with this status.</td></tr>`;
}

function cmRenderAssetOptions(onPlan = new Set((cmDetail?.machines || []).map((m) => m.asset_id))) {
  const cat = cmEl("cmCategory")?.value || "";
  const free = cmAssets.filter((a) => !onPlan.has(Number(a.id)) && (!cat || String(a.category || "").trim() === cat));
  const sel = cmEl("cmAssetSelect");
  if (sel) {
    sel.innerHTML = free.length
      ? `<option value="">Choose a machine</option>${free.map((a) => `<option value="${a.id}">${esc(a.asset_code)} - ${esc(a.asset_name)}</option>`).join("")}`
      : `<option value="">All machines are on the plan</option>`;
  }
  const all = cmEl("cmAddAllBtn");
  if (all) {
    all.textContent = `Add all ${cat ? `${cat} ` : ""}(${free.length})`;
    all.disabled = !free.length;
  }
}

function cmMachineRow(m) {
  const edit = cmCanEdit();
  const statusCell = edit
    ? `<select class="cm-status" data-cm-row="${m.id}">${Object.entries(cmDetail.statuses).map(([k, v]) => `<option value="${k}"${k === m.status ? " selected" : ""}>${esc(v)}</option>`).join("")}</select>`
    : `<span class="pill ${CM_STATUS_PILL[m.status] || "pill-gray"}">${esc(m.status_label)}</span>`;
  const used = m.planned.hours > 0 ? Math.min(100, Math.round((m.actual.hours / m.planned.hours) * 100)) : 0;
  const over = m.planned.hours > 0 && m.actual.hours > m.planned.hours;
  const wo = m.work_order
    ? `<a href="workorders.html?wo=${m.work_order.id}" target="_blank" rel="noopener">WO #${m.work_order.id}</a><div class="muted small">${esc(m.work_order.status)}${m.work_order.technician ? ` · ${esc(m.work_order.technician)}` : ""}</div>`
    : `<span class="muted">None yet</span>`;
  const parts = m.parts.length
    ? (m.parts_short ? `<span class="pill pill-orange">${m.parts_short} of ${m.parts.length} needed</span>` : `<span class="pill pill-green">${m.parts.length} ready</span>`)
    : `<span class="muted">None listed</span>`;
  const open = cmOpenRow === m.id;
  return `
    <tr class="cm-row${open ? " cm-row-open" : ""}">
      <td><button type="button" class="link-btn cm-toggle" data-cm-open="${m.id}" aria-expanded="${open}">${open ? "▾" : "▸"} <strong>${esc(m.asset_code)}</strong></button>
        <div class="muted small">${esc(m.asset_name || "")}${m.priority === "high" ? ' · <span class="pill pill-red">High</span>' : ""}</div></td>
      <td class="cm-scope">${m.scope ? esc(m.scope.split("\n")[0]) : '<span class="muted">Not planned yet</span>'}</td>
      <td>${statusCell}</td>
      <td><div>${cmHours(m.actual.hours)} / ${cmHours(m.planned.hours)}</div>
        <div class="cm-bar${over ? " cm-bar-over" : ""}" title="${m.hours_used_pct ?? 0}% of planned hours"><span style="width:${used}%"></span></div></td>
      <td style="text-align:right;">${cmMoney(m.planned.total)}</td>
      <td style="text-align:right;">${cmMoney(m.actual.total)}</td>
      <td>${parts}</td>
      <td>${wo}</td>
    </tr>
    ${open ? `<tr class="cm-editor-row"><td colspan="8">${cmEditor(m)}</td></tr>` : ""}`;
}

function cmPartLine(p = {}, edit = cmCanEdit()) {
  const state = p.state ? CM_PART_STATE[p.state] : null;
  const lineTotal = Number(p.qty || 0) * Number(p.unit_cost || 0);
  if (!edit) {
    return `<tr><td>${esc(p.part_code || "—")}</td><td>${esc(p.description)}</td><td>${Number(p.qty)}</td><td>${cmMoney(p.unit_cost)}</td><td>${cmMoney(lineTotal)}</td>
      <td>${state ? `<span class="pill ${state[1]}">${state[0]}</span>` : ""}${p.on_hand != null && p.state !== "issued" ? ` <span class="muted small">${Number(p.on_hand)} on hand</span>` : ""}</td><td></td></tr>`;
  }
  return `<tr class="cm-part" data-request-id="${p.request_id || ""}">
    <td><input class="cm-p-code" list="wfPartsList" value="${escAttr(p.part_code || "")}" placeholder="Stores code" /></td>
    <td><input class="cm-p-desc" value="${escAttr(p.description || "")}" placeholder="Part description" /></td>
    <td><input class="cm-p-qty" type="number" min="0.1" step="0.1" value="${escAttr(p.qty ?? 1)}" /></td>
    <td><input class="cm-p-cost" type="number" min="0" step="0.01" value="${escAttr(p.unit_cost ?? "")}" placeholder="0.00" /></td>
    <td class="cm-p-total">${cmMoney(lineTotal)}</td>
    <td>${state ? `<span class="pill ${state[1]}">${state[0]}</span>` : '<span class="muted small">New</span>'}${p.on_hand != null && p.state !== "issued" ? ` <span class="muted small">${Number(p.on_hand)} on hand</span>` : ""}</td>
    <td><button type="button" class="link-btn cm-p-remove" title="Remove this part">✕</button></td>
  </tr>`;
}

function cmEditor(m) {
  const edit = cmCanEdit();
  const issued = m.issued_parts.length
    ? `<p class="muted small">Issued by stores on the work order: ${m.issued_parts.map((i) => `${esc(i.part_code)} × ${i.qty} (${cmMoney(i.cost)})`).join(", ")}</p>`
    : "";
  return `
    <div class="cm-editor" data-cm-editor="${m.id}">
      <div class="cm-editor-grid">
        <label class="cm-wide">What needs doing
          <textarea class="cm-scope-in" rows="3" ${edit ? "" : "readonly"} placeholder="Example: Reseal boom cylinder, replace bucket teeth, full 1000h service">${esc(m.scope)}</textarea></label>
        <label>Repair hours planned<input class="cm-hours-in" type="number" min="0" step="0.5" value="${escAttr(m.planned.hours)}" ${edit ? "" : "readonly"} /></label>
        <label>Priority<select class="cm-priority-in" ${edit ? "" : "disabled"}>
          ${["high", "normal", "low"].map((p) => `<option value="${p}"${p === m.priority ? " selected" : ""}>${p[0].toUpperCase()}${p.slice(1)}</option>`).join("")}</select></label>
        <label class="cm-wide">Notes<input class="cm-notes-in" value="${escAttr(m.notes)}" ${edit ? "" : "readonly"} placeholder="Anything else to remember" /></label>
      </div>
      <h4>Parts needed</h4>
      <div class="table-wrap"><table class="cm-parts-table">
        <thead><tr><th>Stores code</th><th>Description</th><th>Qty</th><th>Unit cost (USD)</th><th>Line total</th><th>Stores</th><th></th></tr></thead>
        <tbody class="cm-parts-body">${m.parts.map((p) => cmPartLine(p, edit)).join("") || (edit ? "" : `<tr><td colspan="7" class="muted">No parts listed.</td></tr>`)}</tbody>
      </table></div>
      ${issued}
      <div class="cm-editor-foot">
        <div class="muted small">Estimate: parts ${cmMoney(m.planned.parts_cost)} + labour ${cmMoney(m.planned.labour_cost)} = <strong>${cmMoney(m.planned.total)}</strong>
          · Used so far: ${cmHours(m.actual.hours)}, ${cmMoney(m.actual.total)}</div>
        ${edit ? `<div class="row cm-actions">
          <button type="button" class="btn-secondary cm-add-part">+ Add part</button>
          <button type="button" class="btn-primary cm-save">Save machine</button>
          ${m.work_order && !m.work_order.finished
            ? `<a class="btn btn-secondary" href="workorders.html?wo=${m.work_order.id}" target="_blank" rel="noopener">Open WO #${m.work_order.id}</a>`
            : `<button type="button" class="btn-secondary cm-make-wo" title="Opens a work order with this work and these parts, so technicians can book time on it">Create work order</button>`}
          <button type="button" class="btn-secondary cm-request" ${m.parts_short ? "" : "disabled"} title="Adds parts the stores don't have to Parts to Order">Order missing parts</button>
          <button type="button" class="link-btn cm-remove">Remove from plan</button>
        </div>` : ""}
      </div>
    </div>`;
}

function cmEditorLines(editor) {
  return [...editor.querySelectorAll(".cm-part")].map((tr) => ({
    part_code: tr.querySelector(".cm-p-code").value.trim(),
    description: tr.querySelector(".cm-p-desc").value.trim(),
    qty: tr.querySelector(".cm-p-qty").value,
    unit_cost: tr.querySelector(".cm-p-cost").value,
    request_id: Number(tr.dataset.requestId || 0) || null,
  })).filter((l) => l.part_code || l.description);
}

async function cmSaveMachine(editor, { quiet = false } = {}) {
  const id = Number(editor.dataset.cmEditor);
  await cmFetch(`/machines/${id}`, {
    method: "PATCH",
    body: {
      scope: editor.querySelector(".cm-scope-in").value,
      planned_hours: editor.querySelector(".cm-hours-in").value,
      priority: editor.querySelector(".cm-priority-in").value,
      notes: editor.querySelector(".cm-notes-in").value,
    },
  });
  await cmFetch(`/machines/${id}/parts`, { method: "PUT", body: { lines: cmEditorLines(editor) } });
  if (!quiet) cmMsg("Saved.");
}

function wireCareMaintenance() {
  cmWired = true;
  const section = cmEl("careMaintenanceSection");
  cmEl("cmPlanSelect")?.addEventListener("change", () => { cmOpenRow = 0; cmLoadPlan().catch((e) => cmMsg(e.message, true)); });
  cmEl("cmStatusFilter")?.addEventListener("change", () => cmRender());
  cmEl("cmCategory")?.addEventListener("change", () => cmRenderAssetOptions());
  cmEl("cmNewPlanBtn")?.addEventListener("click", () => {
    cmEl("cmNewPlanForm")?.classList.toggle("hidden");
    const n = cmEl("cmNewName");
    if (n && !n.value) n.value = `C&M ${new Date().toLocaleDateString("en-GB", { month: "long", year: "numeric" })}`;
    n?.focus();
  });
  cmEl("cmCreatePlanBtn")?.addEventListener("click", async () => {
    try {
      const r = await cmFetch("/plans", {
        method: "POST",
        body: { name: cmEl("cmNewName").value, start_date: cmEl("cmNewStart").value, end_date: cmEl("cmNewEnd").value, labour_rate: cmEl("cmNewRate").value },
      });
      cmEl("cmNewPlanForm").classList.add("hidden");
      cmEl("cmNewName").value = "";
      await cmLoadPlans(r.id);
      cmMsg("Plan created. Now add the machines.");
    } catch (e) { cmMsg(e.message, true); }
  });
  cmEl("cmClosePlanBtn")?.addEventListener("click", async () => {
    if (!cmDetail) return;
    const closing = cmDetail.plan.status !== "closed";
    if (closing && !confirm("Close this plan? You can still view it and reopen it later.")) return;
    try {
      await cmFetch(`/plans/${cmDetail.plan.id}`, { method: "PATCH", body: { status: closing ? "closed" : "active" } });
      await cmLoadPlans(cmDetail.plan.id);
    } catch (e) { cmMsg(e.message, true); }
  });
  cmEl("cmExportBtn")?.addEventListener("click", async () => {
    if (!cmDetail) return;
    try {
      const res = await fetch(`${API}/maintenance/cm/plans/${cmDetail.plan.id}/export.csv`);
      if (!res.ok) throw new Error("Could not download the plan.");
      const url = URL.createObjectURL(await res.blob());
      const a = Object.assign(document.createElement("a"), { href: url, download: `care-maintenance-${cmDetail.plan.name.replace(/[^A-Za-z0-9-]+/g, "_")}.csv` });
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch (e) { cmMsg(e.message, true); }
  });
  const addMachines = async (ids) => {
    if (!cmDetail || !ids.length) return;
    try {
      const r = await cmFetch(`/plans/${cmDetail.plan.id}/machines`, { method: "POST", body: { asset_ids: ids } });
      await cmLoadPlans(cmDetail.plan.id);
      cmMsg(`${r.added} machine${r.added === 1 ? "" : "s"} added.`);
    } catch (e) { cmMsg(e.message, true); }
  };
  cmEl("cmAddMachineBtn")?.addEventListener("click", () => {
    const id = Number(cmEl("cmAssetSelect")?.value || 0);
    if (!id) return cmMsg("Choose a machine first.", true);
    addMachines([id]);
  });
  cmEl("cmAddAllBtn")?.addEventListener("click", () => {
    const ids = [...(cmEl("cmAssetSelect")?.options || [])].map((o) => Number(o.value)).filter(Boolean);
    if (ids.length > 5 && !confirm(`Add ${ids.length} machines to the plan?`)) return;
    addMachines(ids);
  });

  section?.addEventListener("click", async (ev) => {
    const t = ev.target.closest("button");
    if (!t) return;
    const editor = t.closest(".cm-editor");
    try {
      if (t.dataset.cmOpen) {
        const id = Number(t.dataset.cmOpen);
        cmOpenRow = cmOpenRow === id ? 0 : id;
        cmRender();
        if (cmOpenRow) {
          const ed = section.querySelector(`[data-cm-editor="${cmOpenRow}"]`);
          if (ed && !ed.querySelector(".cm-part") && cmCanEdit()) ed.querySelector(".cm-parts-body").insertAdjacentHTML("beforeend", cmPartLine({ qty: 1 }));
        }
        return;
      }
      if (!editor) return;
      const id = Number(editor.dataset.cmEditor);
      if (t.classList.contains("cm-add-part")) {
        editor.querySelector(".cm-parts-body").insertAdjacentHTML("beforeend", cmPartLine({ qty: 1 }));
        editor.querySelector(".cm-part:last-child .cm-p-code")?.focus();
      } else if (t.classList.contains("cm-p-remove")) {
        t.closest("tr").remove();
      } else if (t.classList.contains("cm-save")) {
        await cmSaveMachine(editor);
        await cmLoadPlan();
      } else if (t.classList.contains("cm-make-wo")) {
        await cmSaveMachine(editor, { quiet: true });
        const r = await cmFetch(`/machines/${id}/work-order`, { method: "POST" });
        await cmLoadPlan();
        cmMsg(`Work order #${r.work_order_id} created. Technicians can now book time on it, and stores can issue the parts to it.`);
      } else if (t.classList.contains("cm-request")) {
        await cmSaveMachine(editor, { quiet: true });
        const r = await cmFetch(`/machines/${id}/request-parts`, { method: "POST" });
        await cmLoadPlan();
        cmMsg(r.created.length
          ? `${r.created.length} part${r.created.length === 1 ? "" : "s"} added to Parts to Order.`
          : "Nothing to order: those parts are already in stores or requested.");
      } else if (t.classList.contains("cm-remove")) {
        const m = cmDetail.machines.find((x) => x.id === id);
        if (!confirm(`Remove ${m?.asset_code || "this machine"} from the plan? Its estimate and parts list are deleted. Any work order stays.`)) return;
        await cmFetch(`/machines/${id}`, { method: "DELETE" });
        cmOpenRow = 0;
        await cmLoadPlans(cmDetail.plan.id);
      }
    } catch (e) { cmMsg(e.message, true); }
  });

  section?.addEventListener("change", async (ev) => {
    const sel = ev.target.closest(".cm-status");
    if (sel) {
      try {
        await cmFetch(`/machines/${sel.dataset.cmRow}`, { method: "PATCH", body: { status: sel.value } });
        await cmLoadPlan();
      } catch (e) { cmMsg(e.message, true); }
      return;
    }
    // A stores code fills in the description and latest price.
    const code = ev.target.closest(".cm-p-code");
    if (code && typeof getWfPartByCode === "function") {
      const p = getWfPartByCode(code.value);
      const tr = code.closest("tr");
      if (p) {
        tr.querySelector(".cm-p-desc").value = p.part_name || "";
        const cost = tr.querySelector(".cm-p-cost");
        if (!cost.value && Number(p.latest_unit_cost) > 0) cost.value = Number(p.latest_unit_cost).toFixed(2);
      }
    }
  });
  section?.addEventListener("input", (ev) => {
    const tr = ev.target.closest(".cm-part");
    if (!tr) return;
    const total = Number(tr.querySelector(".cm-p-qty").value || 0) * Number(tr.querySelector(".cm-p-cost").value || 0);
    tr.querySelector(".cm-p-total").textContent = cmMoney(total);
  });
}
