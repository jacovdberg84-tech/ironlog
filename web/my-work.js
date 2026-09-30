/**
 * My Work — the signed-in user's home screen (tab-mywork).
 * Loaded after app.js and uses its globals (API, fetchJson, escapeHtml, switchTab, ...).
 */
const MY_WORK_MAINTENANCE_ROLES = ["admin", "supervisor", "workshop_admin", "plant_manager", "site_manager"];
let myWorkLoading = false;

function myWorkAllowed(tab) {
  try {
    return getEffectiveAllowedTabs().includes(tab);
  } catch {
    return false;
  }
}

function myWorkGo(target) {
  if (/\.html(?:[?#]|$)/.test(target)) {
    location.href = target;
    return;
  }
  switchTab(target);
}

function openMyTasks(taskId) {
  switchTab("tasks");
  setTaskSidebarView("mine", { refresh: true });
  if (taskId) {
    currentTaskId = taskId;
    loadTaskDetail(taskId).catch(() => {});
  }
}

function myWorkDate(value) {
  const s = String(value || "").slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return "";
  const d = new Date(`${s}T00:00:00`);
  return d.toLocaleDateString(undefined, { day: "numeric", month: "short" });
}

function myWorkPill(text, tone) {
  return `<span class="mywork-pill mywork-pill-${tone}">${escapeHtml(text)}</span>`;
}

function renderMyWorkCard({ key, title, count, meta, items, empty, action, tone = "neutral" }) {
  const clear = !count;
  const list = items.length
    ? `<ul class="mywork-list">${items.map((it) => `
        <li${it.taskId ? ` data-mywork-task="${Number(it.taskId)}" tabindex="0" role="button"` : ""}>
          <div class="mywork-item-main">
            <strong>${escapeHtml(it.primary)}</strong>
            ${it.secondary ? `<span>${escapeHtml(it.secondary)}</span>` : ""}
          </div>
          ${it.badge || ""}
        </li>`).join("")}</ul>`
    : `<p class="mywork-empty">${escapeHtml(empty)}</p>`;
  const more = count > items.length ? `<p class="mywork-more">+${count - items.length} more</p>` : "";
  const button = action
    ? `<button type="button" class="mywork-action" data-mywork-go="${escapeHtml(action.target)}">${escapeHtml(action.label)} →</button>`
    : "";
  return `
    <article class="mywork-card ${clear ? "is-clear" : `tone-${tone}`}" data-mywork-card="${key}">
      <header>
        <h3>${escapeHtml(title)}</h3>
        <span class="mywork-count">${clear ? "✓" : count}</span>
      </header>
      ${meta ? `<p class="mywork-meta">${escapeHtml(meta)}</p>` : ""}
      ${list}
      ${more}
      ${button}
    </article>`;
}

function myWorkCards(sections, due) {
  const cards = [];
  const t = sections.tasks;
  if (t) {
    const bits = [];
    if (t.overdue) bits.push(`${t.overdue} overdue`);
    if (t.due_today) bits.push(`${t.due_today} due today`);
    cards.push({
      key: "tasks",
      title: "My tasks",
      count: t.count,
      tone: t.overdue ? "danger" : "info",
      meta: bits.join(" · "),
      empty: "No open tasks assigned to you.",
      items: t.items.map((x) => ({
        taskId: x.id,
        primary: x.title,
        secondary: [x.project, x.due_date ? `Due ${myWorkDate(x.due_date)}` : ""].filter(Boolean).join(" · "),
        badge: x.due_date && x.due_date < (t.today || "")
          ? myWorkPill("Overdue", "danger")
          : x.priority === "high" ? myWorkPill("High", "warn") : "",
      })),
      action: myWorkAllowed("tasks") ? { label: "Open my tasks", target: "tasks" } : null,
    });
  }
  const b = sections.open_breakdowns;
  if (b) {
    cards.push({
      key: "breakdowns",
      title: "Machines broken down",
      count: b.count,
      tone: "danger",
      empty: "No open breakdowns.",
      meta: b.open_breakdown_records > b.count ? `${b.open_breakdown_records} open breakdown records` : "",
      items: b.items.map((x) => ({
        primary: `${x.asset_code} — ${x.description || "Breakdown"}${x.open_breakdowns > 1 ? ` (+${x.open_breakdowns - 1} more)` : ""}`,
        secondary: [myWorkDate(x.breakdown_date) && `Since ${myWorkDate(x.breakdown_date)}`, x.parts_status && `Parts: ${x.parts_status}`].filter(Boolean).join(" · "),
        badge: x.critical ? myWorkPill("Critical", "danger") : "",
      })),
      action: myWorkAllowed("maintenance") ? { label: "Open breakdowns", target: "breakdown-ops.html" } : null,
    });
  }
  if (due) {
    const overdue = due.filter((x) => x.is_overdue).sort((x, y) => Number(x.remaining_hours) - Number(y.remaining_hours));
    const almost = due.filter((x) => x.is_almost_due && !x.is_overdue).length;
    cards.push({
      key: "services",
      title: "Services overdue",
      count: overdue.length,
      tone: "warn",
      meta: almost ? `${almost} more due soon` : "",
      empty: "No services overdue.",
      items: overdue.slice(0, 5).map((x) => ({
        primary: `${x.asset_code} — ${x.service_name || "Service"}`,
        secondary: `${Math.abs(Number(x.remaining_hours || 0)).toFixed(0)} ${x.meter_unit === "km" ? "km" : "h"} overdue`,
      })),
      action: myWorkAllowed("maintenance") ? { label: "Open maintenance", target: "maintenance.html" } : null,
    });
  }
  const w = sections.open_work_orders;
  if (w) {
    cards.push({
      key: "workorders",
      title: "Open work orders",
      count: w.count,
      tone: "info",
      meta: w.unassigned ? `${w.unassigned} not assigned to an artisan` : "",
      empty: "No open work orders.",
      items: w.items.map((x) => ({
        primary: `WO ${x.id} — ${x.asset_code}`,
        secondary: [x.source, x.assigned_artisan_name || "Unassigned", myWorkDate(x.opened_at) && `Opened ${myWorkDate(x.opened_at)}`].filter(Boolean).join(" · "),
        badge: myWorkPill(String(x.status || "open").replace(/_/g, " "), "neutral"),
      })),
      action: myWorkAllowed("maintenance") ? { label: "Open work orders", target: "workorders.html" } : null,
    });
  }
  const os = sections.open_shifts;
  if (os && os.count) {
    cards.push({
      key: "shifts",
      title: "Shift reports not submitted",
      count: os.count,
      tone: "warn",
      meta: "Open for more than 12 hours. Job time reaches the timesheet only when the report is submitted.",
      empty: "",
      items: os.items.map((x) => ({
        primary: x.name,
        secondary: `Shift started ${myWorkDate(x.started_at) || x.started_at} · open ${Math.round(x.hours_open)} h`,
        badge: myWorkPill("Not submitted", "warn"),
      })),
      action: { label: "Open technician portal", target: "technician-terminal.html?tab=shift" },
    });
  }
  const wp = sections.waiting_parts;
  if (wp) {
    const bits = [];
    if (wp.machines_down) bits.push(`${wp.machines_down} machine${wp.machines_down === 1 ? "" : "s"} down`);
    if (wp.in_stock) bits.push(`${wp.in_stock} can be issued from stock`);
    if (wp.not_ordered) bits.push(`${wp.not_ordered} to order`);
    cards.push({
      key: "waitingparts",
      title: "Workshop waiting on parts",
      count: wp.count,
      tone: wp.machines_down ? "danger" : "warn",
      meta: bits.join(" · "),
      empty: "The workshop is not waiting on any parts.",
      items: wp.items.map((x) => ({
        primary: `${x.asset_code || "No machine"} — ${x.kind === "request" ? `${x.part_name || x.part_code} × ${x.qty}` : "part not listed yet"}`,
        secondary: [x.work_order_id ? `WO ${x.work_order_id}` : "", x.kind === "request" ? (x.status === "ordered" ? "Ordered" : "Not ordered yet") : `Parts ${String(x.status || "").toLowerCase()}`, myWorkDate(x.since) && `Since ${myWorkDate(x.since)}`].filter(Boolean).join(" · "),
        badge: x.in_stock ? myWorkPill("In stock", "success") : x.machine_down ? myWorkPill("Machine down", "danger") : "",
      })),
      action: myWorkAllowed("parts-tracking") ? { label: "Open stores queue", target: "parts-tracking" } : null,
    });
  }
  const ls = sections.low_stock;
  if (ls) {
    cards.push({
      key: "lowstock",
      title: "Low stock",
      count: ls.count,
      tone: "warn",
      empty: "All parts are above minimum stock.",
      items: ls.items.map((x) => ({
        primary: `${x.part_code} — ${x.part_name || ""}`,
        secondary: `On hand ${x.on_hand} · minimum ${x.min_stock}`,
        badge: x.critical ? myWorkPill("Critical", "danger") : "",
      })),
      action: myWorkAllowed("stock") ? { label: "Open stock control", target: "stock" } : null,
    });
  }
  const po = sections.parts_on_order;
  if (po) {
    cards.push({
      key: "partsorders",
      title: "Parts on order",
      count: po.count,
      tone: "info",
      empty: "No parts waiting to be received.",
      items: po.items.map((x) => ({
        primary: `${x.part_code} — ${x.part_name || ""}`,
        secondary: [`Qty ${x.quantity}`, x.expected_date ? `Expected ${myWorkDate(x.expected_date)}` : "No expected date"].join(" · "),
        badge: x.critical ? myWorkPill("Critical", "danger") : "",
      })),
      action: myWorkAllowed("stock") ? { label: "Open stock control", target: "stock" } : null,
    });
  }
  return cards;
}

async function loadMyWork() {
  const grid = qs("myWorkGrid");
  if (!grid || myWorkLoading) return;
  myWorkLoading = true;
  const subtitle = qs("myWorkSubtitle");
  const greeting = qs("myWorkGreeting");
  const user = getSessionUser();
  if (greeting) greeting.textContent = user ? `Hi ${user}, here's what needs attention` : "What needs your attention";
  if (!grid.children.length) grid.innerHTML = `<div class="skeleton-block"></div><div class="skeleton-block"></div><div class="skeleton-block"></div>`;
  try {
    const roles = getSessionRoles();
    const wantsDue = roles.some((r) => MY_WORK_MAINTENANCE_ROLES.includes(r)) && myWorkAllowed("maintenance");
    const [work, dueRes] = await Promise.all([
      fetchJson(`${API}/api/my-work`),
      wantsDue ? fetchJson(`${API}/api/maintenance/due`).catch(() => null) : Promise.resolve(null),
    ]);
    const sections = work.sections || {};
    if (sections.tasks) sections.tasks.today = work.today;
    const due = dueRes && Array.isArray(dueRes.due) ? dueRes.due : null;
    const cards = myWorkCards(sections, due);
    grid.innerHTML = cards.map(renderMyWorkCard).join("");
    loadCostingGapsCard(grid).catch(() => {});
    const open = cards.reduce((n, c) => n + (c.key === "services" || c.key === "partsorders" ? 0 : c.count), 0);
    if (subtitle) {
      const today = new Date().toLocaleDateString(undefined, { weekday: "long", day: "numeric", month: "long" });
      subtitle.textContent = open ? `${today} · ${open} item${open === 1 ? "" : "s"} to follow up` : `${today} · You're all caught up`;
    }
  } catch (err) {
    grid.innerHTML = `<p class="mywork-error">Could not load your work list: ${escapeHtml(err.message || String(err))}</p>`;
    if (subtitle) subtitle.textContent = "";
  } finally {
    myWorkLoading = false;
  }
}

(function initMyWork() {
  const grid = qs("myWorkGrid");
  if (!grid) return;
  grid.addEventListener("click", (e) => {
    const go = e.target.closest("[data-mywork-go]");
    if (go) {
      const target = go.dataset.myworkGo;
      if (target === "tasks") openMyTasks();
      else myWorkGo(target);
      return;
    }
    const task = e.target.closest("[data-mywork-task]");
    if (task) openMyTasks(Number(task.dataset.myworkTask));
  });
  grid.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" && e.key !== " ") return;
    const task = e.target.closest("[data-mywork-task]");
    if (!task) return;
    e.preventDefault();
    openMyTasks(Number(task.dataset.myworkTask));
  });
  qs("myWorkRefresh")?.addEventListener("click", () => loadMyWork());
})();

/* ---------- Costing gaps + Borris costing assistant ---------- */

const COSTING_GAP_ROLES = ["admin", "supervisor", "workshop_admin", "plant_manager", "site_manager"];
let myWorkCostingGaps = [];
let costingGapsExpanded = false;

function costingGapOpenTarget(g) {
  if (g.type === "service_no_labour" && g.work_order_id) return `workorders.html?wo=${g.work_order_id}`;
  if (g.type === "part_zero_cost") return myWorkAllowed("stock") ? "stock" : "";
  if (g.type === "labour_rate_missing") return myWorkAllowed("admin") ? "admin" : "";
  if (g.type === "service_unpriced") return "maintenance.html";
  return "";
}

function renderCostingGapsCard(gaps) {
  const shown = costingGapsExpanded ? gaps : gaps.slice(0, 5);
  const counts = [];
  const n = (t) => gaps.filter((g) => g.type === t).length;
  if (n("service_unpriced")) counts.push(`${n("service_unpriced")} service${n("service_unpriced") === 1 ? "" : "s"} unpriced`);
  if (n("part_zero_cost")) counts.push(`${n("part_zero_cost")} part${n("part_zero_cost") === 1 ? "" : "s"} at $0`);
  if (n("service_no_labour")) counts.push(`${n("service_no_labour")} job${n("service_no_labour") === 1 ? "" : "s"} without labour`);
  const items = shown.map((g) => {
    const target = costingGapOpenTarget(g);
    return `<li class="costing-gap sev-${escapeHtml(g.severity)}">
      <div class="mywork-item-main">
        <strong>${escapeHtml(g.title)}</strong>
        <span>${escapeHtml(g.detail)}</span>
        <div class="costing-gap-actions">
          ${g.borris ? `<button type="button" class="btn-primary" data-costing-assist="${escapeHtml(g.key)}">Work on it with Borris</button>` : ""}
          ${target ? `<button type="button" data-mywork-go="${escapeHtml(target)}">Open</button>` : ""}
          <button type="button" data-costing-dismiss="${escapeHtml(g.key)}">Not needed</button>
        </div>
      </div>
    </li>`;
  }).join("");
  const more = gaps.length > 5
    ? `<button type="button" class="mywork-more costing-more" data-costing-more="1">${costingGapsExpanded ? "Show fewer" : `Show all ${gaps.length}`}</button>`
    : "";
  return `
    <article class="mywork-card mywork-card-wide ${gaps.length ? "tone-warn" : "is-clear"}" data-mywork-card="costing">
      <header>
        <h3>Costing gaps</h3>
        <span class="mywork-count">${gaps.length ? gaps.length : "✓"}</span>
      </header>
      <p class="mywork-meta">${escapeHtml(counts.join(" · ") || "Every upcoming service and part in use has a price.")}</p>
      ${gaps.length ? `<ul class="mywork-list costing-gap-list">${items}</ul>` : ""}
      ${more}
    </article>`;
}

async function loadCostingGapsCard(grid) {
  const roles = getSessionRoles();
  if (!roles.some((r) => COSTING_GAP_ROLES.includes(r)) || !myWorkAllowed("maintenance")) return;
  try {
    const data = await fetchJson(`${API}/api/maintenance/costing-gaps`);
    myWorkCostingGaps = Array.isArray(data.gaps) ? data.gaps : [];
  } catch {
    return;
  }
  grid.querySelector('[data-mywork-card="costing"]')?.remove();
  grid.insertAdjacentHTML("beforeend", renderCostingGapsCard(myWorkCostingGaps));
}

/* The Borris panel: evidence-based proposal the person can edit and apply. */
function borrisPanel() {
  let panel = qs("borrisCostingPanel");
  if (panel) return panel;
  document.body.insertAdjacentHTML("beforeend", `
    <div id="borrisCostingBackdrop" class="bw-backdrop" hidden></div>
    <section id="borrisCostingPanel" class="bw-panel" role="dialog" aria-modal="true" aria-labelledby="bwTitle" hidden>
      <header class="bw-head">
        <div><span class="bw-kicker">Borris · costing</span><h3 id="bwTitle">Costing gap</h3></div>
        <button type="button" data-bw-close>Close</button>
      </header>
      <div id="bwBody" class="bw-body"></div>
    </section>`);
  panel = qs("borrisCostingPanel");
  qs("borrisCostingBackdrop").addEventListener("click", closeBorrisPanel);
  panel.addEventListener("click", onBorrisPanelClick);
  panel.addEventListener("input", onBorrisPanelInput);
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") closeBorrisPanel(); });
  return panel;
}

function openBorrisPanel(title) {
  const panel = borrisPanel();
  qs("bwTitle").textContent = title;
  panel.hidden = false;
  qs("borrisCostingBackdrop").hidden = false;
}

function closeBorrisPanel() {
  assistSeq += 1; // stops polling and ignores late answers for the closed gap
  if (qs("borrisCostingPanel")) qs("borrisCostingPanel").hidden = true;
  if (qs("borrisCostingBackdrop")) qs("borrisCostingBackdrop").hidden = true;
}

let borrisCurrent = null;
// Each opened gap gets a number; answers for an older number are ignored, so a
// slow reply for one machine can never land in another machine's panel.
let assistSeq = 0;
let borrisPoll = null;
let storePartsCache = null;

function money(n) {
  return `$${Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function renderServiceProposal(data) {
  const p = data.proposal;
  const plan = data.plan;
  const unit = plan.meter_unit === "km" ? "km" : "h";
  const modeNote = p.mode === "borris"
    ? `<span class="bw-badge">Borris proposal</span>`
    : `<span class="bw-badge is-history">From service history</span>`;
  const rows = p.lines.map((l, i) => `
    <tr>
      <td><strong>${escapeHtml(l.part_code)}</strong><div class="bw-sub">${escapeHtml(l.part_name || "")}</div>${l.why ? `<div class="bw-why">${escapeHtml(l.why)}</div>` : ""}</td>
      <td><input type="number" min="0" step="0.01" value="${l.qty}" data-bw-qty="${i}" aria-label="Quantity for ${escapeHtml(l.part_code)}" /></td>
      <td class="${l.unit_cost > 0 ? "" : "bw-missing"}">${l.unit_cost > 0 ? money(l.unit_cost) : "No price"}</td>
      <td data-bw-line="${i}">${money(l.line_cost)}</td>
      <td><button type="button" class="bw-remove" data-bw-remove="${i}" aria-label="Remove ${escapeHtml(l.part_code)}">✕</button></td>
    </tr>`).join("");
  const src = p.sources;
  return `
    <p class="bw-lead">${escapeHtml(plan.asset_code)} ${escapeHtml(plan.asset_name || "")} · ${escapeHtml(plan.service_name || "")} (${Number(plan.interval_hours)} ${unit})</p>
    <div class="bw-mode">${modeNote}</div>
    ${p.summary ? `<p class="bw-summary">${escapeHtml(p.summary)}</p>` : ""}
    ${p.borris_note ? `<p class="bw-note">Borris: ${escapeHtml(p.borris_note)}</p>` : ""}
    ${p.lines.length ? `
      <table class="bw-table">
        <thead><tr><th>Part</th><th>Qty</th><th>Store price</th><th>Line</th><th></th></tr></thead>
        <tbody>${rows}</tbody>
      </table>` : `<p class="bw-note">No parts proposed yet.</p>`}
    ${p.unpriced_codes.length ? `<p class="bw-warn">No store price for ${escapeHtml(p.unpriced_codes.join(", "))} — those lines count as $0 until priced.</p>` : ""}
    ${p.rejected_codes.length ? `<p class="bw-note">Left out (not in stores): ${escapeHtml(p.rejected_codes.join(", "))}.</p>` : ""}
    <div class="bw-labour">
      <label>Labour hours <input type="number" min="0" step="0.5" value="${p.labour.hours}" data-bw-hours="1" /></label>
      <span>× ${money(p.labour.rate)}/h = <strong data-bw-labour-total>${money(p.labour.total)}</strong></span>
    </div>
    <p class="bw-total">Service estimate: <strong data-bw-total>${money(p.totals.total)}</strong></p>
    <div class="bw-addpart">
      <input list="bwPartOptions" data-bw-add-code placeholder="Add a store part (code or name)" aria-label="Store part to add" />
      <input type="number" min="0" step="0.01" value="1" data-bw-add-qty aria-label="Quantity to add" />
      <button type="button" data-bw-add="1">Add part</button>
    </div>
    <datalist id="bwPartOptions"></datalist>
    ${p.questions.length ? `<div class="bw-questions"><strong>Borris needs to know:</strong><ul>${p.questions.map((q) => `<li>${escapeHtml(q)}</li>`).join("")}</ul></div>` : ""}
    <div class="bw-notes">
      <label for="bwNotes"><strong>Tell Borris what you know</strong> <span class="bw-sub">Filters, oils and litres, part numbers, manual page or supplier kit. Saved for this service.</span></label>
      <textarea id="bwNotes" rows="3" data-bw-notes placeholder="e.g. 1000 h: engine oil 18 L 15W40, oil filter LF9009, fuel filter FF5488, hydraulic return filter 1x">${escapeHtml(data.notes_draft ?? data.planner_notes?.notes ?? "")}</textarea>
      <div class="bw-notes-foot">
        <button type="button" data-bw-save-notes="1">Save and ask Borris again</button>
        ${data.planner_notes?.updated_at ? `<span class="bw-sub">Saved ${escapeHtml(String(data.planner_notes.updated_at).slice(0, 16))}${data.planner_notes.updated_by ? ` by ${escapeHtml(data.planner_notes.updated_by)}` : ""}</span>` : ""}
      </div>
    </div>
    <div class="bw-borris" data-bw-borris>${borrisStatusHtml(data.borris)}</div>
    <p class="bw-sources">Evidence: ${src.own_services} past service${src.own_services === 1 ? "" : "s"} on this machine${src.peer_assets.length ? `; ${src.peer_services} on ${escapeHtml(src.peer_assets.join(", "))}` : ""}${src.manual.length ? `; manual: ${escapeHtml(src.manual.join("; "))}` : "; no manual pages matched"}.</p>
    <div class="bw-actions">
      <button type="button" class="btn-primary" data-bw-apply="service">Apply to service cost</button>
      <button type="button" data-costing-dismiss="${escapeHtml(data.key)}">Not needed</button>
    </div>
    <p class="bw-msg" data-bw-msg role="status"></p>
    ${borrisFollowUpHtml()}`;
}

function renderPartProposal(data) {
  const p = data.proposal;
  const list = (rows, label) => rows.length
    ? `<p class="bw-sub">${label}</p><ul class="bw-list">${rows.map((r) => `<li>${escapeHtml(r.part_code || r.source || "")} ${r.part_name ? `— ${escapeHtml(r.part_name)}` : ""} · ${money(r.unit_cost)}${r.date ? ` · ${escapeHtml(String(r.date).slice(0, 10))}` : ""}</li>`).join("")}</ul>`
    : "";
  return `
    <p class="bw-lead">${escapeHtml(p.part_code)}</p>
    <div class="bw-mode"><span class="bw-badge is-history">From purchase records</span></div>
    <p class="bw-summary">${escapeHtml(p.summary)}</p>
    ${list(p.purchases, "Purchase prices on record")}
    ${list(p.similar, "Similar store items (a guide only)")}
    <div class="bw-labour">
      <label>Unit cost (USD) <input type="number" min="0" step="0.01" value="${p.unit_cost ?? ""}" data-bw-unit-cost="1" /></label>
    </div>
    <div class="bw-actions">
      <button type="button" class="btn-primary" data-bw-apply="part">Save store price</button>
      <button type="button" data-costing-dismiss="${escapeHtml(data.key)}">Not needed</button>
    </div>
    <p class="bw-msg" data-bw-msg role="status"></p>`;
}

function borrisFollowUpHtml() {
  return `
    <details class="bw-followup">
      <summary>Ask Borris about this</summary>
      <textarea data-bw-question rows="2" placeholder="e.g. Which fuel filter does the 950GC use at 500 h?"></textarea>
      <button type="button" data-bw-ask="1">Ask</button>
      <div class="bw-answer" data-bw-answer></div>
    </details>`;
}

function borrisStatusHtml(b) {
  const st = b?.status || "off";
  if (st === "off") return `<span class="bw-sub">Borris (AI) is not set up on this server — use the records and your notes.</span>`;
  if (st === "queued") return `<span class="bw-spinner" aria-hidden="true"></span> Borris is waiting his turn${b.ahead ? ` (${b.ahead} ahead)` : ""}. You can keep working; his proposal appears here.`;
  if (st === "working") return `<span class="bw-spinner" aria-hidden="true"></span> Borris is working on a proposal${b.seconds ? ` (${b.seconds} s)` : ""}. You can keep working or close this; it appears here when ready.`;
  if (st === "busy") return `<span class="bw-sub">${escapeHtml(b.error || "Borris is busy. Try again in a minute.")}</span>`;
  if (st === "failed") return `<span class="bw-warn-inline">Borris could not answer (${escapeHtml(b.error || "no reply")}).</span> Tell him more above and ask again, or add the parts yourself.`;
  if (st === "done" && b.proposal) {
    const p = b.proposal;
    return `<div class="bw-ready"><strong>Borris's proposal is ready:</strong> ${p.lines.length} part${p.lines.length === 1 ? "" : "s"}, ${escapeHtml(money(p.totals.total))} with labour.
      ${p.summary ? `<div class="bw-sub">${escapeHtml(p.summary)}</div>` : ""}
      <button type="button" class="btn-primary" data-bw-use-borris="1">Use Borris's proposal</button></div>`;
  }
  return "";
}

function setBorrisStatus(b) {
  if (!borrisCurrent) return;
  borrisCurrent.borris = b;
  const el = qs("borrisCostingPanel")?.querySelector("[data-bw-borris]");
  if (el) el.innerHTML = borrisStatusHtml(b);
}

function pollBorris(key, seq) {
  clearTimeout(borrisPoll);
  borrisPoll = setTimeout(async () => {
    if (seq !== assistSeq) return;
    try {
      const res = await fetchJson(`${API}/api/maintenance/costing-gaps/assist/status?key=${encodeURIComponent(key)}`);
      if (seq !== assistSeq) return;
      const b = { ...res.borris, proposal: res.proposal };
      setBorrisStatus(b);
      if (["queued", "working"].includes(b.status)) pollBorris(key, seq);
    } catch {
      if (seq === assistSeq) pollBorris(key, seq);
    }
  }, 5000);
}

async function openCostingAssist(key, { refresh = false } = {}) {
  const seq = ++assistSeq;
  const gap = myWorkCostingGaps.find((g) => g.key === key);
  openBorrisPanel(gap?.title || "Costing gap");
  const body = qs("bwBody");
  body.innerHTML = `<div class="bw-thinking">Reading the service history, sister machines, store prices and manuals…</div>`;
  try {
    const data = await fetchJson(`${API}/api/maintenance/costing-gaps/assist`, { method: "POST", body: JSON.stringify({ key, refresh }) });
    if (seq !== assistSeq) return; // another gap was opened meanwhile
    borrisCurrent = JSON.parse(JSON.stringify(data));
    body.innerHTML = data.kind === "part" ? renderPartProposal(data) : renderServiceProposal(data);
    if (data.kind === "service") {
      fillStorePartOptions().catch(() => {});
      if (["queued", "working"].includes(data.borris?.status)) pollBorris(key, seq);
    }
  } catch (err) {
    if (seq === assistSeq) body.innerHTML = `<p class="bw-warn">Could not get a proposal: ${escapeHtml(err.message || String(err))}</p>`;
  }
}

async function fillStorePartOptions() {
  if (!storePartsCache) {
    const rows = await fetchJson(`${API}/api/stock/onhand`);
    storePartsCache = Array.isArray(rows) ? rows : [];
  }
  const list = qs("bwPartOptions");
  if (list && !list.children.length) {
    list.innerHTML = storePartsCache.map((r) => `<option value="${escapeHtml(r.part_code)}">${escapeHtml(`${r.part_name || ""} · ${money(r.unit_cost)}${Number(r.on_hand) > 0 ? ` · ${r.on_hand} in stock` : ""}`)}</option>`).join("");
  }
}

function rerenderService() {
  const notes = qs("borrisCostingPanel")?.querySelector("[data-bw-notes]");
  if (notes) borrisCurrent.notes_draft = notes.value;
  qs("bwBody").innerHTML = renderServiceProposal(borrisCurrent);
  recalcServiceProposal();
  fillStorePartOptions().catch(() => {});
}

function addPartToProposal() {
  const panel = qs("borrisCostingPanel");
  const raw = String(panel.querySelector("[data-bw-add-code]")?.value || "").trim();
  const qty = Number(panel.querySelector("[data-bw-add-qty]")?.value || 0);
  if (!raw) return bwMsg("Pick a store part to add.", false);
  if (!(qty > 0)) return bwMsg("Enter a quantity above zero.", false);
  const part = (storePartsCache || []).find((r) => String(r.part_code).toLowerCase() === raw.toLowerCase());
  if (!part) return bwMsg(`${raw} is not in stores. Add it to stock first, or tell Borris in the notes.`, false);
  const p = borrisCurrent.proposal;
  const existing = p.lines.find((l) => l.part_code === part.part_code);
  const unit = Number(part.unit_cost || 0);
  if (existing) {
    existing.qty = Number((existing.qty + qty).toFixed(2));
    existing.line_cost = Number((existing.qty * unit).toFixed(2));
  } else {
    p.lines.push({ part_code: part.part_code, part_name: part.part_name, type: part.stock_category === "oil" ? "oil" : "part", qty, unit_cost: unit, line_cost: Number((qty * unit).toFixed(2)), why: "added by you" });
  }
  p.unpriced_codes = p.lines.filter((l) => !(l.unit_cost > 0)).map((l) => l.part_code);
  rerenderService();
}

async function saveNotesAndAskAgain() {
  const d = borrisCurrent;
  const notes = String(qs("borrisCostingPanel").querySelector("[data-bw-notes]")?.value || "");
  const seq = assistSeq;
  const res = await fetchJson(`${API}/api/maintenance/costing-gaps/notes`, { method: "POST", body: JSON.stringify({ plan_id: d.plan.plan_id, notes }) });
  if (seq !== assistSeq) return;
  d.planner_notes = res.planner_notes;
  d.notes_draft = notes;
  bwMsg(notes.trim() ? "Notes saved. Borris is reading them." : "Notes cleared.", true);
  setBorrisStatus(res.borris);
  if (["queued", "working"].includes(res.borris?.status)) pollBorris(d.key, seq);
}

function useBorrisProposal() {
  const p = borrisCurrent?.borris?.proposal;
  if (!p) return;
  borrisCurrent.proposal = JSON.parse(JSON.stringify(p));
  rerenderService();
}

function recalcServiceProposal() {
  const p = borrisCurrent?.proposal;
  if (!p) return;
  const parts = p.lines.reduce((s, l) => s + l.line_cost, 0);
  p.labour.total = Number((p.labour.hours * p.labour.rate).toFixed(2));
  p.totals = { parts, labour: p.labour.total, total: Number((parts + p.labour.total).toFixed(2)) };
  const panel = qs("borrisCostingPanel");
  p.lines.forEach((l, i) => { const el = panel.querySelector(`[data-bw-line="${i}"]`); if (el) el.textContent = money(l.line_cost); });
  const lt = panel.querySelector("[data-bw-labour-total]");
  if (lt) lt.textContent = money(p.labour.total);
  const tt = panel.querySelector("[data-bw-total]");
  if (tt) tt.textContent = money(p.totals.total);
}

function onBorrisPanelInput(e) {
  if (e.target.dataset.bwNotes != null && borrisCurrent) borrisCurrent.notes_draft = e.target.value;
  const p = borrisCurrent?.proposal;
  if (!p) return;
  const qtyIdx = e.target.dataset.bwQty;
  if (qtyIdx != null) {
    const l = p.lines[Number(qtyIdx)];
    l.qty = Math.max(0, Number(e.target.value || 0));
    l.line_cost = Number((l.qty * l.unit_cost).toFixed(2));
    recalcServiceProposal();
  }
  if (e.target.dataset.bwHours) {
    p.labour.hours = Math.max(0, Number(e.target.value || 0));
    recalcServiceProposal();
  }
}

function bwMsg(text, ok) {
  const el = qs("borrisCostingPanel")?.querySelector("[data-bw-msg]");
  if (!el) return;
  el.textContent = text;
  el.className = `bw-msg ${ok ? "is-ok" : "is-err"}`;
}

async function applyBorrisProposal(kind) {
  const data = borrisCurrent;
  if (!data) return;
  const user = getSessionUser();
  const today = new Date().toISOString().slice(0, 10);
  if (kind === "service") {
    const p = data.proposal;
    const items = p.lines.filter((l) => l.qty > 0).map((l) => ({ type: l.type, part_code: l.part_code, qty: l.qty }));
    if (!items.length && !(p.labour.total > 0)) return bwMsg("Add at least one part or labour hours first.", false);
    await fetchJson(`${API}/api/maintenance/weekly-forum/forecast-inputs`, {
      method: "POST",
      body: JSON.stringify({
        plan_id: data.plan.plan_id,
        items,
        labor_total: p.labour.total,
        notes: `${p.mode === "borris" ? "Borris proposal" : "Costing assistant estimate"} approved by ${user || "user"} on ${today}. ${p.summary || ""}`.slice(0, 500),
      }),
    });
    bwMsg(`Saved: ${data.plan.asset_code} ${data.plan.service_name} is now costed at ${money(p.totals.total)}.`, true);
  } else {
    const input = qs("borrisCostingPanel").querySelector("[data-bw-unit-cost]");
    const unit_cost = Number(input?.value);
    if (!(unit_cost > 0)) return bwMsg("Enter the unit cost from the invoice or quote.", false);
    await fetchJson(`${API}/api/dashboard/cost/part-cost`, { method: "POST", body: JSON.stringify({ part_code: data.proposal.part_code, unit_cost }) });
    bwMsg(`Saved: ${data.proposal.part_code} now costs ${money(unit_cost)}.`, true);
  }
  qs("borrisCostingPanel").querySelectorAll("[data-bw-apply]").forEach((b) => { b.disabled = true; });
  loadMyWork();
}

async function askBorrisFollowUp() {
  const panel = qs("borrisCostingPanel");
  const q = String(panel.querySelector("[data-bw-question]")?.value || "").trim();
  const out = panel.querySelector("[data-bw-answer]");
  if (!q || !out) return;
  out.textContent = "Borris is thinking…";
  const d = borrisCurrent || {};
  const context = d.kind === "service"
    ? `Costing gap: ${d.plan?.asset_code} ${d.plan?.asset_name} ${d.plan?.service_name} (${d.plan?.interval_hours}). Current proposal: ${(d.proposal?.lines || []).map((l) => `${l.part_code} x${l.qty}`).join(", ") || "none"}; labour ${d.proposal?.labour?.hours} h.`
    : `Costing gap: store part ${d.proposal?.part_code} has no price.`;
  try {
    const res = await fetchJson(`${API}/api/ironmind/ask`, { method: "POST", body: JSON.stringify({ question: q, asset_code: d.plan?.asset_code || "", context_notes: context }) });
    out.textContent = res.short_answer || "No answer.";
  } catch (err) {
    out.textContent = `Borris could not answer: ${err.message || err}`;
  }
}

async function dismissCostingGapUi(key) {
  const reason = window.prompt("Why is this not needed? (optional — e.g. 'machine sold', 'priced by contractor')", "");
  if (reason === null) return;
  await fetchJson(`${API}/api/maintenance/costing-gaps/dismiss`, { method: "POST", body: JSON.stringify({ key, reason }) });
  closeBorrisPanel();
  loadMyWork();
}

function onBorrisPanelClick(e) {
  if (e.target.closest("[data-bw-close]")) return closeBorrisPanel();
  const rm = e.target.closest("[data-bw-remove]");
  if (rm && borrisCurrent?.proposal) {
    borrisCurrent.proposal.lines.splice(Number(rm.dataset.bwRemove), 1);
    rerenderService();
    return;
  }
  const apply = e.target.closest("[data-bw-apply]");
  if (apply) {
    apply.disabled = true;
    applyBorrisProposal(apply.dataset.bwApply).catch((err) => { apply.disabled = false; bwMsg(err.message || String(err), false); });
    return;
  }
  const dis = e.target.closest("[data-costing-dismiss]");
  if (dis) return void dismissCostingGapUi(dis.dataset.costingDismiss).catch((err) => bwMsg(err.message || String(err), false));
  if (e.target.closest("[data-bw-ask]")) return void askBorrisFollowUp();
  if (e.target.closest("[data-bw-add]")) return addPartToProposal();
  if (e.target.closest("[data-bw-use-borris]")) return useBorrisProposal();
  const save = e.target.closest("[data-bw-save-notes]");
  if (save) {
    save.disabled = true;
    saveNotesAndAskAgain().catch((err) => bwMsg(err.message || String(err), false)).finally(() => { save.disabled = false; });
  }
}

(function initCostingGaps() {
  const grid = qs("myWorkGrid");
  if (!grid) return;
  grid.addEventListener("click", (e) => {
    const assist = e.target.closest("[data-costing-assist]");
    if (assist) return void openCostingAssist(assist.dataset.costingAssist);
    const dis = e.target.closest("[data-costing-dismiss]");
    if (dis) return void dismissCostingGapUi(dis.dataset.costingDismiss).catch(() => {});
    if (e.target.closest("[data-costing-more]")) {
      costingGapsExpanded = !costingGapsExpanded;
      grid.querySelector('[data-mywork-card="costing"]')?.remove();
      grid.insertAdjacentHTML("beforeend", renderCostingGapsCard(myWorkCostingGaps));
    }
  });
})();
