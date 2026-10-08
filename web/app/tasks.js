// IRONLOG/web/app/tasks.js — Task workspace.
// Part of the main app; index.html loads these files in order and they share one global scope.
//
// One list with quick filters, a side panel for the open task (details and
// conversation), and a short "New task" form. Comments are signed by the
// signed-in user; @name tags someone and puts the task in their inbox.

let teamMembers = [];
let twView = "mine";
let twProject = "";
let twSearch = "";
let twTasks = [];
let twOpenId = null;
let twProjects = [];
let twNewPriority = "medium";
// Kept for core.js: the sidebar marks the active Tasks link with this key.
let currentTaskSidebarActiveKey = "";

const TW_VIEW_KEY = "ironlog_task_view";
const TW_STATUS = { open: "Open", in_progress: "In progress", done: "Done" };

function twMember(username) {
  const u = String(username || "").toLowerCase();
  return teamMembers.find((m) => String(m.username || "").toLowerCase() === u) || null;
}
function twName(username) {
  if (!username) return "";
  const m = twMember(username);
  return m?.full_name || username;
}
function twInitials(username) {
  return String(twName(username) || "?").split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0].toUpperCase()).join("") || "?";
}
function twToday() {
  return new Date().toISOString().slice(0, 10);
}
function twDueText(due, status) {
  if (!due) return "";
  const d = String(due).slice(0, 10);
  const today = twToday();
  if (status !== "done" && d < today) return `<span class="tw-due late">Overdue · ${escapeHtml(d)}</span>`;
  if (d === today) return `<span class="tw-due soon">Due today</span>`;
  return `<span class="tw-due">Due ${escapeHtml(d)}</span>`;
}
function twAgo(ts) {
  const t = new Date(String(ts || "").replace(" ", "T") + (String(ts || "").includes("Z") ? "" : "Z"));
  const s = Math.max(0, (Date.now() - t.getTime()) / 1000);
  if (!Number.isFinite(s)) return "";
  if (s < 60) return "just now";
  if (s < 3600) return `${Math.round(s / 60)} min ago`;
  if (s < 86400) return `${Math.round(s / 3600)} h ago`;
  return t.toLocaleDateString(undefined, { day: "numeric", month: "short" });
}
/** Comment text with @mentions highlighted. */
function twRich(text) {
  return escapeHtml(text).replace(/(^|[^\w@])@([A-Za-z0-9._-]{2,40})/g, (m, pre, name) => (twMember(name) ? `${pre}<span class="tw-at">@${name}</span>` : m)).replace(/\n/g, "<br>");
}

async function loadTeamMembers() {
  try {
    const res = await fetchJson(`${API}/api/auth/team`);
    teamMembers = Array.isArray(res.rows) ? res.rows : [];
  } catch {
    teamMembers = [{ username: getSessionUser(), full_name: null, role: getSessionRole() }];
  }
}

function twPeopleOptions(selected = "", { blank = "Nobody yet" } = {}) {
  const sorted = [...teamMembers].sort((a, b) => String(a.full_name || a.username).localeCompare(String(b.full_name || b.username)));
  return `<option value="">${escapeHtml(blank)}</option>${sorted.map((m) => `<option value="${escapeHtml(m.username)}" ${String(m.username).toLowerCase() === String(selected || "").toLowerCase() ? "selected" : ""}>${escapeHtml(m.full_name || m.username)}</option>`).join("")}`;
}

async function loadProjects() {
  try {
    const res = await fetchJson(`${API}/api/projects`);
    twProjects = Array.isArray(res.projects) ? res.projects : [];
  } catch {
    twProjects = [];
  }
  const opts = twProjects.map((p) => `<option value="${escapeHtml(p.name)}">${escapeHtml(p.name)}</option>`).join("");
  if (qs("twProject")) {
    qs("twProject").innerHTML = `<option value="">All projects</option>${opts}`;
    qs("twProject").value = twProject;
  }
  if (qs("twNewProject")) qs("twNewProject").innerHTML = `<option value="">No project</option>${opts}`;
}

async function loadTasks() {
  const list = qs("twList");
  if (!list) return;
  const p = new URLSearchParams({ view: twView });
  if (twSearch) p.set("q", twSearch);
  if (twProject) p.set("project", twProject);
  try {
    const res = await fetchJson(`${API}/api/tasks?${p}`);
    twTasks = res.tasks || [];
  } catch (e) {
    list.innerHTML = `<div class="tw-empty">Could not load tasks: ${escapeHtml(e.message || String(e))}</div>`;
    return;
  }
  renderTaskList();
  loadTasksStats();
}

function renderTaskList() {
  const list = qs("twList");
  if (!list) return;
  document.querySelectorAll("[data-tw-view]").forEach((b) => b.classList.toggle("on", b.dataset.twView === twView));
  if (!twTasks.length) {
    const empty = {
      mine: "Nothing assigned to you. 🎉",
      created: "You have not assigned any open tasks.",
      watching: "You are not following any open tasks.",
      overdue: "No overdue tasks.",
      done: "No finished tasks yet.",
    }[twView] || "No tasks found.";
    list.innerHTML = `<div class="tw-empty">${escapeHtml(twSearch ? "No tasks match your search." : empty)}</div>`;
    return;
  }
  list.innerHTML = twTasks.map((t) => {
    const done = t.status === "done";
    return `
      <div class="tw-item ${twOpenId === t.id ? "on" : ""} ${done ? "done" : ""}" data-tw-open="${t.id}">
        <button type="button" class="tw-check ${done ? "on" : ""}" data-tw-done="${t.id}" title="${done ? "Reopen" : "Mark done"}" aria-label="${done ? "Reopen" : "Mark done"}">✓</button>
        <div class="tw-item-main">
          <div class="tw-item-title"><span class="tw-pri ${escapeHtml(t.priority || "medium")}" title="${escapeHtml(t.priority || "medium")} priority"></span>${escapeHtml(t.title)}</div>
          <div class="tw-meta">
            ${t.assigned_to ? `<span class="tw-who"><span class="tw-av">${escapeHtml(twInitials(t.assigned_to))}</span>${escapeHtml(twName(t.assigned_to))}</span>` : `<span class="tw-who muted">Unassigned</span>`}
            ${twDueText(t.due_date, t.status)}
            ${t.status === "in_progress" ? `<span class="tw-tag prog">In progress</span>` : ""}
            ${t.project ? `<span class="tw-tag">${escapeHtml(t.project)}</span>` : ""}
            ${t.comments_count ? `<span class="tw-count" title="Comments">💬 ${t.comments_count}</span>` : ""}
            ${t.watching ? `<span class="tw-count" title="You follow this task">★</span>` : ""}
          </div>
        </div>
        <span class="tw-id">#${t.id}</span>
      </div>`;
  }).join("");
}

async function loadTasksStats() {
  const el = qs("twSummary");
  if (!el) return;
  try {
    const [stats, mine] = await Promise.all([
      fetchJson(`${API}/api/tasks/stats/summary`),
      fetchJson(`${API}/api/tasks?view=mine`),
    ]);
    const myCount = (mine.tasks || []).length;
    el.textContent = `${myCount} open for you · ${stats.open + stats.in_progress} open in the team${stats.overdue ? ` · ${stats.overdue} overdue` : ""}`;
  } catch {
    el.textContent = "";
  }
}

async function openTask(id) {
  twOpenId = Number(id);
  renderTaskList();
  const panel = qs("twDetail");
  if (!panel) return;
  panel.hidden = false;
  panel.innerHTML = `<div class="tw-empty">Loading…</div>`;
  let res;
  try {
    res = await fetchJson(`${API}/api/tasks/${twOpenId}`);
  } catch (e) {
    panel.innerHTML = `<div class="tw-empty">${escapeHtml(e.message || String(e))}</div>`;
    return;
  }
  // Opening the task reads its inbox items.
  fetchJson(`${API}/api/collab/inbox/read`, { method: "POST", body: JSON.stringify({ task_id: twOpenId }) }).then(() => window.CollabInbox?.refresh()).catch(() => {});
  renderTaskDetail(res);
}

function renderTaskDetail(res) {
  const panel = qs("twDetail");
  const t = res.task;
  const me = getSessionUser();
  const staff = (getSessionRoles?.() || [getSessionRole()]).some((r) => ["admin", "supervisor", "workshop_admin", "plant_manager", "site_manager"].includes(r));
  const canDelete = staff || String(t.created_by || "").toLowerCase() === String(me).toLowerCase();
  const comments = res.comments || [];
  const watchers = res.watchers || [];
  panel.innerHTML = `
    <div class="tw-d-head">
      <span class="tw-id">#${t.id}</span>
      <button type="button" class="tw-x" data-tw-close aria-label="Close">✕</button>
    </div>
    <h3 class="tw-d-title" contenteditable="true" spellcheck="true" data-tw-title>${escapeHtml(t.title)}</h3>
    <div class="tw-seg tw-status">
      ${Object.entries(TW_STATUS).map(([k, label]) => `<button type="button" data-tw-status="${k}" class="${t.status === k ? "on" : ""}">${label}</button>`).join("")}
    </div>
    <div class="tw-fields">
      <label><span>Assigned to</span><select data-tw-field="assigned_to">${twPeopleOptions(t.assigned_to)}</select></label>
      <label><span>Due</span><input type="date" data-tw-field="due_date" value="${escapeHtml(String(t.due_date || "").slice(0, 10))}" /></label>
      <label><span>Priority</span><select data-tw-field="priority">${["low", "medium", "high"].map((p) => `<option value="${p}" ${t.priority === p ? "selected" : ""}>${p[0].toUpperCase()}${p.slice(1)}</option>`).join("")}</select></label>
      <label><span>Project</span><select data-tw-field="project"><option value="">No project</option>${twProjects.map((p) => `<option ${p.name === t.project ? "selected" : ""}>${escapeHtml(p.name)}</option>`).join("")}</select></label>
    </div>
    <div class="tw-desc" contenteditable="true" data-tw-desc data-placeholder="Add a description…">${escapeHtml(t.description || "")}</div>
    <div class="tw-follow">
      <span class="muted small">Created by ${escapeHtml(twName(t.created_by) || "—")} · Followers:</span>
      ${watchers.map((w) => `<span class="tw-tag">${escapeHtml(twName(w))}</span>`).join("") || `<span class="muted small">none</span>`}
      <button type="button" class="btn btn-secondary btn-sm" data-tw-watch="${res.watching ? "0" : "1"}">${res.watching ? "Unfollow" : "☆ Follow"}</button>
    </div>
    <div class="tw-thread" id="twThread">
      ${comments.length ? comments.map((c) => {
        const mine = String(c.author || "").toLowerCase() === String(me).toLowerCase();
        return `
          <div class="tw-msg ${mine ? "mine" : ""}">
            <span class="tw-av" title="${escapeHtml(twName(c.author))}">${escapeHtml(twInitials(c.author))}</span>
            <div class="tw-bubble">
              <div class="tw-msg-head"><b>${escapeHtml(twName(c.author))}</b><span>${escapeHtml(twAgo(c.created_at))}</span>${mine || staff ? `<button type="button" class="tw-del" data-tw-del-comment="${c.id}" title="Delete">✕</button>` : ""}</div>
              <div>${twRich(c.comment)}</div>
            </div>
          </div>`;
      }).join("") : `<div class="tw-empty small">No messages yet. Ask a question or give an update; type @ to tag someone.</div>`}
    </div>
    <div class="tw-compose">
      <div class="tw-mention-list" id="twMentionList" hidden></div>
      <textarea id="twReply" rows="2" placeholder="Write a message… (@ to tag, Ctrl+Enter to send)"></textarea>
      <button type="button" class="btn btn-primary" id="twSend">Send</button>
    </div>
    ${canDelete ? `<div class="tw-d-foot"><button type="button" class="tw-link danger" data-tw-delete>Delete task</button></div>` : ""}`;
  const thread = qs("twThread");
  thread.scrollTop = thread.scrollHeight;
  wireMentionBox(qs("twReply"), qs("twMentionList"));
  qs("twReply").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); sendTaskComment(); }
  });
  qs("twSend").onclick = sendTaskComment;
}

/** @mention suggestions under a textarea. */
function wireMentionBox(ta, box) {
  if (!ta || !box) return;
  let matches = [];
  const close = () => { box.hidden = true; matches = []; };
  ta.addEventListener("input", () => {
    const before = ta.value.slice(0, ta.selectionStart);
    const m = /(^|\s)@([\w.-]*)$/.exec(before);
    if (!m) return close();
    const q = m[2].toLowerCase();
    matches = teamMembers.filter((p) => String(p.username).toLowerCase().startsWith(q) || String(p.full_name || "").toLowerCase().split(/\s+/).some((w) => w.startsWith(q))).slice(0, 6);
    if (!matches.length) return close();
    box.innerHTML = matches.map((p, i) => `<button type="button" data-i="${i}"><span class="tw-av">${escapeHtml(twInitials(p.username))}</span>${escapeHtml(p.full_name || p.username)} <span class="muted">@${escapeHtml(p.username)}</span></button>`).join("");
    box.hidden = false;
  });
  box.addEventListener("mousedown", (e) => {
    const b = e.target.closest("[data-i]");
    if (!b) return;
    e.preventDefault();
    const p = matches[Number(b.dataset.i)];
    const pos = ta.selectionStart;
    const before = ta.value.slice(0, pos).replace(/@([\w.-]*)$/, `@${p.username} `);
    ta.value = before + ta.value.slice(pos);
    ta.setSelectionRange(before.length, before.length);
    ta.focus();
    close();
  });
  ta.addEventListener("blur", () => setTimeout(close, 150));
}

async function sendTaskComment() {
  const ta = qs("twReply");
  const text = String(ta?.value || "").trim();
  if (!text || !twOpenId) return;
  qs("twSend").disabled = true;
  try {
    await fetchJson(`${API}/api/tasks/${twOpenId}/comments`, { method: "POST", body: JSON.stringify({ comment: text }) });
    await openTask(twOpenId);
    loadTasks();
  } catch (e) {
    alert(`Could not send: ${e.message || e}`);
    if (qs("twSend")) qs("twSend").disabled = false;
  }
}

async function updateOpenTask(patch) {
  if (!twOpenId) return;
  try {
    await fetchJson(`${API}/api/tasks/${twOpenId}`, { method: "PUT", body: JSON.stringify(patch) });
  } catch (e) {
    alert(`Could not save: ${e.message || e}`);
  }
  await openTask(twOpenId);
  loadTasks();
}

function openNewTaskModal() {
  const m = qs("twModal");
  if (!m) return;
  qs("twNewTitle").value = "";
  qs("twNewDesc").value = "";
  qs("twNewDue").value = "";
  qs("twNewAssignee").innerHTML = twPeopleOptions("", { blank: "Nobody yet" });
  qs("twNewFollowers").innerHTML = twPeopleOptions("", { blank: "" }).replace('<option value=""></option>', "");
  qs("twNewProject").value = twProject || "";
  qs("twNewErr").textContent = "";
  twNewPriority = "medium";
  qs("twNewPriority").querySelectorAll("[data-p]").forEach((b) => b.classList.toggle("on", b.dataset.p === "medium"));
  m.hidden = false;
  setTimeout(() => qs("twNewTitle").focus(), 30);
}

async function saveNewTask() {
  const title = qs("twNewTitle").value.trim();
  if (!title) { qs("twNewErr").textContent = "Write what needs to be done."; return; }
  const followers = [...qs("twNewFollowers").selectedOptions].map((o) => o.value).filter(Boolean);
  qs("twNewSave").disabled = true;
  try {
    const res = await fetchJson(`${API}/api/tasks`, {
      method: "POST",
      body: JSON.stringify({
        title,
        description: qs("twNewDesc").value.trim() || null,
        assigned_to: qs("twNewAssignee").value || null,
        due_date: qs("twNewDue").value || null,
        priority: twNewPriority,
        project: qs("twNewProject").value || null,
        watchers: followers,
      }),
    });
    qs("twModal").hidden = true;
    setStatus("Task created.");
    await loadTasks();
    openTask(res.task.id);
  } catch (e) {
    qs("twNewErr").textContent = e.message || String(e);
  } finally {
    qs("twNewSave").disabled = false;
  }
}

/** Sidebar links (core.js) and the floating inbox open the workspace through here. */
function setTaskSidebarView(view, options = {}) {
  const map = { inbox: "mine", today: "mine", upcoming: "open", all: "open" };
  twView = ["mine", "created", "watching", "open", "overdue", "done"].includes(view) ? view : map[view] || "mine";
  if (options.project !== undefined) twProject = String(options.project || "");
  currentTaskSidebarActiveKey = "";
  if (options.refresh) loadTasks();
}
function initTaskWorkspaceSidebar() { /* the workspace no longer has a sidebar sub-menu */ }

/** Open the Tasks page at one task (used by the floating inbox). */
async function showTask(id) {
  if (typeof switchTab === "function") switchTab("tasks");
  if (!twTasks.some((t) => t.id === Number(id))) {
    twView = "open";
    twSearch = "";
    if (qs("twSearch")) qs("twSearch").value = "";
    await loadTasks();
  }
  openTask(id);
}

function initTasks() {
  try { twView = localStorage.getItem(TW_VIEW_KEY) || "mine"; } catch { twView = "mine"; }
  qs("twViews")?.addEventListener("click", (e) => {
    const b = e.target.closest("[data-tw-view]");
    if (!b) return;
    twView = b.dataset.twView;
    try { localStorage.setItem(TW_VIEW_KEY, twView); } catch { /* storage off */ }
    loadTasks();
  });
  let timer = null;
  qs("twSearch")?.addEventListener("input", (e) => {
    twSearch = e.target.value.trim();
    clearTimeout(timer);
    timer = setTimeout(loadTasks, 250);
  });
  qs("twProject")?.addEventListener("change", (e) => { twProject = e.target.value; loadTasks(); });
  qs("twNewBtn")?.addEventListener("click", openNewTaskModal);
  qs("twNewCancel")?.addEventListener("click", () => { qs("twModal").hidden = true; });
  qs("twModal")?.addEventListener("click", (e) => { if (e.target === qs("twModal")) qs("twModal").hidden = true; });
  qs("twNewSave")?.addEventListener("click", saveNewTask);
  qs("twNewTitle")?.addEventListener("keydown", (e) => { if (e.key === "Enter") saveNewTask(); });
  qs("twNewPriority")?.addEventListener("click", (e) => {
    const b = e.target.closest("[data-p]");
    if (!b) return;
    twNewPriority = b.dataset.p;
    qs("twNewPriority").querySelectorAll("[data-p]").forEach((x) => x.classList.toggle("on", x === b));
  });
  qs("twAddProject")?.addEventListener("click", async () => {
    const name = qs("twNewProjectName").value.trim();
    if (!name) return;
    try {
      await fetchJson(`${API}/api/projects`, { method: "POST", body: JSON.stringify({ name, color: "#3b82f6" }) });
      qs("twNewProjectName").value = "";
      await loadProjects();
      qs("twNewProject").value = name;
    } catch (e) {
      qs("twNewErr").textContent = e.message || String(e);
    }
  });

  qs("twList")?.addEventListener("click", async (e) => {
    const done = e.target.closest("[data-tw-done]");
    if (done) {
      e.stopPropagation();
      const t = twTasks.find((x) => x.id === Number(done.dataset.twDone));
      await fetchJson(`${API}/api/tasks/${done.dataset.twDone}`, { method: "PUT", body: JSON.stringify({ status: t?.status === "done" ? "open" : "done" }) });
      if (twOpenId === Number(done.dataset.twDone)) openTask(twOpenId);
      return loadTasks();
    }
    const row = e.target.closest("[data-tw-open]");
    if (row) openTask(row.dataset.twOpen);
  });

  const panel = qs("twDetail");
  panel?.addEventListener("click", async (e) => {
    if (e.target.closest("[data-tw-close]")) {
      twOpenId = null;
      panel.hidden = true;
      return renderTaskList();
    }
    const st = e.target.closest("[data-tw-status]");
    if (st) return updateOpenTask({ status: st.dataset.twStatus });
    const w = e.target.closest("[data-tw-watch]");
    if (w) {
      await fetchJson(`${API}/api/tasks/${twOpenId}/watch`, { method: "POST", body: JSON.stringify({ watch: w.dataset.twWatch === "1" }) });
      openTask(twOpenId);
      return loadTasks();
    }
    const del = e.target.closest("[data-tw-del-comment]");
    if (del) {
      if (!confirm("Delete this message?")) return;
      try {
        await fetchJson(`${API}/api/comments/${del.dataset.twDelComment}`, { method: "DELETE" });
      } catch (err) { alert(err.message || err); }
      return openTask(twOpenId);
    }
    if (e.target.closest("[data-tw-delete]")) {
      if (!confirm("Delete this task and its conversation?")) return;
      try {
        await fetchJson(`${API}/api/tasks/${twOpenId}`, { method: "DELETE" });
      } catch (err) { return alert(err.message || err); }
      twOpenId = null;
      panel.hidden = true;
      loadTasks();
    }
  });
  panel?.addEventListener("change", (e) => {
    const f = e.target.closest("[data-tw-field]");
    if (f) updateOpenTask({ [f.dataset.twField]: f.value || null });
  });
  // Title and description save when you click away.
  panel?.addEventListener("focusout", (e) => {
    const t = twTasks.find((x) => x.id === twOpenId);
    if (e.target.matches("[data-tw-title]")) {
      const v = e.target.innerText.trim();
      if (v && v !== t?.title) updateOpenTask({ title: v });
    }
    if (e.target.matches("[data-tw-desc]")) {
      const v = e.target.innerText.trim();
      if (v !== String(t?.description || "").trim()) updateOpenTask({ description: v || null });
    }
  });
  panel?.addEventListener("keydown", (e) => {
    if (e.target.matches("[data-tw-title]") && e.key === "Enter") { e.preventDefault(); e.target.blur(); }
  });

  Promise.all([loadTeamMembers(), loadProjects()]).then(loadTasks);
}
