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
  if (target.endsWith(".html")) {
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
