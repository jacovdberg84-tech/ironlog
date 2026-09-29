// IRONLOG/web/app/tasks.js — Task workspace.
// Part of the main app; index.html loads these files in order and they share one global scope.

// Task Management
let currentProjectFilter = "";
let currentTaskId = null;
let currentTaskView = "all";
let currentTaskSidebarActiveKey = "";
let teamMembers = [];

function getSavedTaskViews() {
  try {
    const raw = localStorage.getItem(TASK_SAVED_VIEWS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function persistSavedTaskViews(views) {
  localStorage.setItem(TASK_SAVED_VIEWS_KEY, JSON.stringify(Array.isArray(views) ? views : []));
}

function renderTeamMemberInputs() {
  const datalist = qs("teamMemberList");
  const authorSelect = qs("taskCommentAuthor");
  const mentions = qs("taskCommentMentions");
  const quickAssign = qs("taskAssignQuickPicks");
  const members = Array.isArray(teamMembers) ? teamMembers : [];
  const sorted = [...members].sort((a, b) => String(a.username || "").localeCompare(String(b.username || "")));
  if (datalist) {
    datalist.innerHTML = sorted
      .map((m) => `<option value="${escapeHtml(m.username)}">${escapeHtml(m.full_name || m.username)}</option>`)
      .join("");
  }
  if (authorSelect) {
    const me = getSessionUser();
    authorSelect.innerHTML = sorted
      .map((m) => `<option value="${escapeHtml(m.username)}">${escapeHtml(m.full_name || m.username)}</option>`)
      .join("");
    if (sorted.some((m) => m.username === me)) authorSelect.value = me;
  }
  if (mentions) {
    if (!sorted.length) {
      mentions.innerHTML = "";
      return;
    }
    mentions.innerHTML = `Tag team: ${sorted
      .slice(0, 8)
      .map(
        (m) =>
          `<button type="button" class="btn btn-secondary btn-sm" data-mention-user="${escapeHtml(m.username)}" style="margin:2px 4px 2px 0;padding:2px 8px;">@${escapeHtml(m.username)}</button>`
      )
      .join("")}`;
    mentions.querySelectorAll("[data-mention-user]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const ta = qs("newComment");
        if (!ta) return;
        ta.value = `${ta.value || ""}${ta.value ? " " : ""}@${btn.dataset.mentionUser} `;
        ta.focus();
      });
    });
  }
  if (quickAssign) {
    if (!sorted.length) {
      quickAssign.innerHTML = "";
      return;
    }
    quickAssign.innerHTML = `Quick assign: ${sorted
      .slice(0, 8)
      .map(
        (m) =>
          `<button type="button" class="btn btn-secondary btn-sm" data-assign-user="${escapeHtml(m.username)}" style="margin:2px 4px 2px 0;padding:2px 8px;">${escapeHtml(m.full_name || m.username)}</button>`
      )
      .join("")}`;
    quickAssign.querySelectorAll("[data-assign-user]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const assigned = qs("taskAssigned");
        if (!assigned) return;
        assigned.value = String(btn.dataset.assignUser || "");
        assigned.focus();
      });
    });
  }
}

async function loadTeamMembers() {
  try {
    const res = await fetchJson(`${API}/api/auth/team`);
    teamMembers = Array.isArray(res.rows) ? res.rows : [];
  } catch {
    teamMembers = [{ username: getSessionUser(), full_name: null, role: getSessionRole() }];
  }
  renderTeamMemberInputs();
}

function renderTaskWorkspaceSavedViews() {
  const wrap = qs("taskCustomViewSidebarLinks");
  if (!wrap) return;
  const views = getSavedTaskViews();
  if (!views.length) {
    wrap.innerHTML = `<div class="nav-item nav-subitem" style="pointer-events:none;"><span>No saved views</span></div>`;
    return;
  }
  wrap.innerHTML = views
    .map(
      (v) => `
      <a href="#" class="nav-item nav-subitem" data-tab="tasks" data-task-view="${escapeHtml(v.view || "all")}" data-task-assigned="${escapeHtml(v.assigned || "")}" data-task-project="${escapeHtml(v.project || "")}" data-task-priority="${escapeHtml(v.priority || "")}" data-task-status="${escapeHtml(v.status || "")}" data-active-key="tasks:saved:${escapeHtml(v.id)}">
        <span class="saved-view-label">${escapeHtml(v.name || "Saved View")}</span>
        <span class="saved-view-actions">
          <button type="button" class="task-saved-view-action" data-task-view-action="rename" data-task-view-id="${escapeHtml(v.id)}" title="Rename view">Rename</button>
          <button type="button" class="task-saved-view-action" data-task-view-action="delete" data-task-view-id="${escapeHtml(v.id)}" title="Delete view">Delete</button>
        </span>
      </a>
    `
    )
    .join("");
}

function renameSavedTaskView(viewId) {
  const id = String(viewId || "").trim();
  if (!id) return;
  const views = getSavedTaskViews();
  const item = views.find((v) => String(v.id) === id);
  if (!item) return;
  const nextName = String(prompt("Rename saved view:", item.name || "Saved View") || "").trim();
  if (!nextName) return;
  item.name = nextName;
  persistSavedTaskViews(views);
  renderTaskWorkspaceSavedViews();
}

function deleteSavedTaskView(viewId) {
  const id = String(viewId || "").trim();
  if (!id) return;
  const views = getSavedTaskViews();
  const filtered = views.filter((v) => String(v.id) !== id);
  persistSavedTaskViews(filtered);
  if (currentTaskSidebarActiveKey === `tasks:saved:${id}`) {
    currentTaskSidebarActiveKey = "";
    updateSidebarActiveState("tasks");
  }
  renderTaskWorkspaceSavedViews();
}

function renderTaskWorkspaceProjectLinks(projects = []) {
  const wrap = qs("taskProjectSidebarLinks");
  if (!wrap) return;
  if (!Array.isArray(projects) || !projects.length) {
    wrap.innerHTML = `<div class="nav-item nav-subitem" style="pointer-events:none;"><span>No projects</span></div>`;
    return;
  }
  wrap.innerHTML = projects
    .map((p) => {
      const color = escapeHtml(p.color || "#3b82f6");
      const name = escapeHtml(p.name || "");
      return `<a href="#" class="nav-item nav-subitem project-link" data-tab="tasks" data-task-view="all" data-task-project="${name}" data-active-key="tasks:project:${name}" style="--project-dot-color:${color};"><span>${name}</span></a>`;
    })
    .join("");
}

function setTaskApiIndicator(state, label) {
  const el = qs("taskApiIndicator");
  if (!el) return;
  el.classList.remove("api-indicator-ok", "api-indicator-down", "api-indicator-unknown");
  if (state === "ok") {
    el.classList.add("api-indicator-ok");
    el.textContent = label || "API online";
    return;
  }
  if (state === "down") {
    el.classList.add("api-indicator-down");
    el.textContent = label || "API unavailable";
    return;
  }
  el.classList.add("api-indicator-unknown");
  el.textContent = label || "API checking...";
}

function initTaskWorkspaceSidebar() {
  const toggle = qs("taskWorkspaceToggle");
  const links = qs("taskWorkspaceLinks");
  if (!toggle || !links) return;

  const apply = (collapsed) => {
    links.style.display = collapsed ? "none" : "";
    toggle.setAttribute("aria-expanded", collapsed ? "false" : "true");
  };

  apply(localStorage.getItem(TASK_WORKSPACE_COLLAPSED_KEY) !== "0");
  toggle.addEventListener("click", () => {
    const collapsed = links.style.display !== "none";
    apply(collapsed);
    localStorage.setItem(TASK_WORKSPACE_COLLAPSED_KEY, collapsed ? "1" : "0");
  });

  links.addEventListener("click", (e) => {
    const btn = e.target?.closest?.(".task-saved-view-action");
    if (!btn) return;
    e.preventDefault();
    e.stopPropagation();
    const action = String(btn.dataset.taskViewAction || "").trim();
    const id = String(btn.dataset.taskViewId || "").trim();
    if (action === "rename") renameSavedTaskView(id);
    if (action === "delete") deleteSavedTaskView(id);
  });
}

function normalizeDateOnly(value) {
  return String(value || "").trim().slice(0, 10);
}

function taskMatchesSidebarView(task, view) {
  const v = String(view || "all").trim().toLowerCase();
  const due = normalizeDateOnly(task?.due_date);
  const today = new Date().toISOString().slice(0, 10);
  if (v === "today") return due === today;
  if (v === "upcoming") return !!due && due > today;
  if (v === "overdue") return !!due && due < today && String(task?.status || "").toLowerCase() !== "done";
  if (v === "inbox") return !String(task?.project || "").trim();
  return true;
}

function setTaskSidebarView(view, options = {}) {
  const v = String(view || "all").trim().toLowerCase();
  const assignedEl = qs("taskFilterAssigned");
  const statusEl = qs("taskFilterStatus");
  const priorityEl = qs("taskFilterPriority");
  currentTaskView = v || "all";

  if (statusEl && options.status !== undefined) statusEl.value = String(options.status || "");
  if (priorityEl && options.priority !== undefined) priorityEl.value = String(options.priority || "");
  if (options.project !== undefined) currentProjectFilter = String(options.project || "");

  if (assignedEl) {
    if (options.assigned !== undefined) {
      assignedEl.value = String(options.assigned || "");
    } else if (currentTaskView === "mine") {
      assignedEl.value = getSessionUser();
    } else if (options.clearAssigned !== false) {
      assignedEl.value = "";
    }
  }

  if (options.activeKey) {
    currentTaskSidebarActiveKey = String(options.activeKey);
  } else if (currentTaskView === "mine") {
    currentTaskSidebarActiveKey = "tasks:mine";
  } else if (["inbox", "today", "upcoming", "overdue"].includes(currentTaskView)) {
    currentTaskSidebarActiveKey = `tasks:${currentTaskView}`;
  } else if (currentProjectFilter) {
    currentTaskSidebarActiveKey = `tasks:project:${currentProjectFilter}`;
  } else {
    currentTaskSidebarActiveKey = "";
  }

  updateSidebarActiveState("tasks");
  if (options.refresh) loadTasks();
}

async function loadProjects() {
  try {
    const res = await fetchJson(`${API}/api/projects`);
    if (res.projects) {
      const tabsEl = qs("projectTabs");
      const projectSelect = qs("taskProject");
      const projectsList = qs("projectsList");
      
      // Render project tabs
      tabsEl.innerHTML = `<button class="project-tab ${!currentProjectFilter ? 'active' : ''}" data-project="">All Tasks</button>` +
        res.projects.map(p => `
          <button class="project-tab ${currentProjectFilter === p.name ? 'active' : ''}" data-project="${escapeHtml(p.name)}" style="--project-color:${escapeHtml(p.color || '#3b82f6')}">
            <span class="project-dot"></span>
            ${escapeHtml(p.name)}
            <span class="project-count">${p.task_count || 0}</span>
          </button>
        `).join("");
      
      // Render project select dropdown
      projectSelect.innerHTML = `<option value="">No Project</option>` +
        res.projects.map(p => `<option value="${escapeHtml(p.name)}">${escapeHtml(p.name)}</option>`).join("");
      
      // Render projects list in sidebar
      projectsList.innerHTML = res.projects.map(p => `
        <div class="project-item" style="border-left:3px solid ${escapeHtml(p.color || '#3b82f6')};">
          <div style="display:flex;justify-content:space-between;align-items:center;">
            <strong style="font-size:13px;">${escapeHtml(p.name)}</strong>
            <span class="muted" style="font-size:11px;">${p.task_count || 0} tasks</span>
          </div>
          ${p.description ? `<small class="muted">${escapeHtml(p.description)}</small>` : ""}
        </div>
      `).join("");
      
      // Add click handlers for project tabs
      tabsEl.querySelectorAll(".project-tab").forEach(tab => {
        tab.addEventListener("click", () => {
          currentProjectFilter = tab.dataset.project;
          currentTaskSidebarActiveKey = currentProjectFilter ? `tasks:project:${currentProjectFilter}` : "";
          loadProjects();
          loadTasks();
        });
      });
      renderTaskWorkspaceProjectLinks(res.projects);
    }
  } catch (err) {
    console.error("Failed to load projects", err);
  }
}

async function loadTasks() {
  const listEl = qs("tasksList");
  if (!listEl) return;
  
  const status = qs("taskFilterStatus")?.value || "";
  const priority = qs("taskFilterPriority")?.value || "";
  const assigned = qs("taskFilterAssigned")?.value || "";
  
  let url = `${API}/api/tasks?`;
  if (status) url += `status=${encodeURIComponent(status)}&`;
  if (priority) url += `priority=${encodeURIComponent(priority)}&`;
  if (assigned) url += `assigned=${encodeURIComponent(assigned)}&`;
  if (currentProjectFilter) url += `project=${encodeURIComponent(currentProjectFilter)}&`;
  
  try {
    const res = await fetchJson(url);
    setTaskApiIndicator("ok", "API online");
    const filteredTasks = (res.tasks || []).filter((task) => taskMatchesSidebarView(task, currentTaskView));
    if (!filteredTasks.length) {
      listEl.innerHTML = `<div class="item"><small class="muted">No tasks found.</small></div>`;
      return;
    }
    
    listEl.innerHTML = filteredTasks.map(task => {
      const priorityClass = task.priority === "high" ? "pill-red" : task.priority === "medium" ? "pill-orange" : "";
      const statusClass = task.status === "done" ? "pill-green" : task.status === "in_progress" ? "pill-blue" : "pill-gray";
      const dueClass = task.due_date && task.status !== "done" ? (new Date(task.due_date) < new Date() ? "text-danger" : "") : "";
      const checked = task.status === "done" ? "checked" : "";
      const strikethrough = task.status === "done" ? "text-decoration:line-through;opacity:0.6" : "";
      
      return `<div class="item task-item ${currentTaskId === task.id ? 'active' : ''}" data-task-id="${task.id}" style="border-left:3px solid ${task.priority === 'high' ? '#dc2626' : task.priority === 'medium' ? '#d97706' : '#16a34a'}; padding-left:12px; margin-bottom:8px; cursor:pointer;">
        <div style="display:flex; align-items:flex-start; gap:12px;">
          <input type="checkbox" ${checked} data-task-toggle="${task.id}" style="margin-top:4px;" onclick="event.stopPropagation();" />
          <div style="flex:1;">
            <div style="font-weight:500; ${strikethrough}">${escapeHtml(task.title)}</div>
            ${task.description ? `<small class="muted">${escapeHtml(task.description.substring(0, 80))}${task.description.length > 80 ? '...' : ''}</small><br/>` : ""}
            <div style="margin-top:6px; display:flex; gap:8px; flex-wrap:wrap; align-items:center;">
              <span class="pill ${statusClass}" data-task-status="${task.id}">${task.status.replace("_", " ")}</span>
              <span class="pill ${priorityClass}">${task.priority}</span>
              ${task.project ? `<span class="pill pill-blue">${escapeHtml(task.project)}</span>` : ""}
              ${task.assigned_to ? `<span class="muted" style="font-size:11px;">${escapeHtml(task.assigned_to)}</span>` : ""}
              ${task.due_date ? `<span class="muted ${dueClass}" style="font-size:11px;">${task.due_date}</span>` : ""}
              ${task.comments_count > 0 ? `<span class="muted" style="font-size:11px;">💬 ${task.comments_count}</span>` : ""}
              <button class="btn btn-secondary btn-sm" data-task-edit="${task.id}" style="padding:2px 8px; font-size:11px;" onclick="event.stopPropagation();">Edit</button>
              <button class="btn btn-secondary btn-sm" data-task-delete="${task.id}" style="padding:2px 8px; font-size:11px; color:#dc2626;" onclick="event.stopPropagation();">X</button>
            </div>
          </div>
        </div>
      </div>`;
    }).join("");
    
    // Add click handlers for viewing task details
    listEl.querySelectorAll(".task-item").forEach(item => {
      item.addEventListener("click", async (e) => {
        const id = parseInt(item.dataset.taskId);
        currentTaskId = id;
        await loadTaskDetail(id);
        loadTasks();
      });
    });
    
    // Add event listeners
    listEl.querySelectorAll("[data-task-toggle]").forEach(cb => {
      cb.addEventListener("change", async (e) => {
        const id = e.target.dataset.taskToggle;
        const done = e.target.checked;
        await fetchJson(`${API}/api/tasks/${id}`, {
          method: "PUT",
          body: JSON.stringify({ status: done ? "done" : "open" })
        });
        loadTasks();
        loadTasksStats();
      });
    });
    
    listEl.querySelectorAll("[data-task-status]").forEach(el => {
      el.addEventListener("click", async (e) => {
        e.stopPropagation();
        const id = e.target.dataset.taskStatus;
        const current = e.target.textContent.trim().replace(" ", "_");
        const statuses = ["open", "in_progress", "done"];
        const currentIdx = statuses.indexOf(current);
        const next = statuses[(currentIdx + 1) % statuses.length];
        await fetchJson(`${API}/api/tasks/${id}`, {
          method: "PUT",
          body: JSON.stringify({ status: next })
        });
        loadTasks();
        loadTasksStats();
      });
    });
    
    listEl.querySelectorAll("[data-task-edit]").forEach(btn => {
      btn.addEventListener("click", async (e) => {
        const id = e.target.dataset.taskEdit;
        const res = await fetchJson(`${API}/api/tasks/${id}`);
        if (res.task) {
          qs("taskTitle").value = res.task.title || "";
          qs("taskDescription").value = res.task.description || "";
          qs("taskProject").value = res.task.project || "";
          qs("taskPriority").value = res.task.priority || "medium";
          qs("taskAssigned").value = res.task.assigned_to || "";
          qs("taskDueDate").value = res.task.due_date || "";
          qs("taskTitle").dataset.editId = id;
          qs("createTaskBtn").textContent = "Update Task";
        }
      });
    });
    
    listEl.querySelectorAll("[data-task-delete]").forEach(btn => {
      btn.addEventListener("click", async (e) => {
        if (!confirm("Delete this task?")) return;
        const id = e.target.dataset.taskDelete;
        await fetchJson(`${API}/api/tasks/${id}`, { method: "DELETE" });
        if (currentTaskId === parseInt(id)) {
          currentTaskId = null;
          qs("taskDetailPanel").style.display = "none";
        }
        loadTasks();
        loadTasksStats();
      });
    });
  } catch (err) {
    setTaskApiIndicator("down", "API unavailable");
    listEl.innerHTML = `<div class="item"><small class="muted">Error loading tasks.</small></div>`;
  }
}

async function loadTaskDetail(taskId) {
  const panel = qs("taskDetailPanel");
  const content = qs("taskDetailContent");
  const comments = qs("taskComments");
  
  if (!panel || !content) return;
  
  try {
    const res = await fetchJson(`${API}/api/tasks/${taskId}`);
    if (res.task) {
      const task = res.task;
      const statusClass = task.status === "done" ? "pill-green" : task.status === "in_progress" ? "pill-blue" : "pill-gray";
      const priorityClass = task.priority === "high" ? "pill-red" : task.priority === "medium" ? "pill-orange" : "";
      const isOverdue = task.due_date && task.status !== "done" && new Date(task.due_date) < new Date();
      
      content.innerHTML = `
        <h3 style="margin:0 0 8px 0;">${escapeHtml(task.title)}</h3>
        <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:12px;">
          <span class="pill ${statusClass}">${task.status.replace("_", " ")}</span>
          <span class="pill ${priorityClass}">${task.priority}</span>
          ${task.project ? `<span class="pill pill-blue">${escapeHtml(task.project)}</span>` : ""}
          ${isOverdue ? `<span class="pill pill-red">Overdue</span>` : ""}
        </div>
        ${task.description ? `<p style="margin:0 0 12px 0;color:var(--muted);">${escapeHtml(task.description)}</p>` : ""}
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:13px;">
          ${task.assigned_to ? `<div><span class="muted">Assigned:</span> ${escapeHtml(task.assigned_to)}</div>` : ""}
          ${task.due_date ? `<div><span class="muted">Due:</span> <span class="${isOverdue ? 'text-danger' : ''}">${task.due_date}</span></div>` : ""}
          <div><span class="muted">Created:</span> ${task.created_at ? task.created_at.split("T")[0] : ""}</div>
        </div>
      `;
      
      const commentRows = Array.isArray(res.comments) ? res.comments : [];
      comments.innerHTML = commentRows.length
        ? commentRows.map(c => `
            <div class="comment-item">
              <div class="comment-header">
                <strong>${escapeHtml(c.author || "User")}</strong>
                <span class="muted">${new Date(c.created_at).toLocaleString()}</span>
              </div>
              <p style="margin:4px 0 0 0;font-size:13px;">${escapeHtml(c.comment)}</p>
              <button class="btn btn-link btn-sm" data-delete-comment="${c.id}" style="color:#dc2626;font-size:11px;padding:0;">Delete</button>
            </div>
          `).join("")
        : `<small class="muted">No comments yet.</small>`;
      
      comments.querySelectorAll("[data-delete-comment]").forEach(btn => {
        btn.addEventListener("click", async () => {
          if (!confirm("Delete this comment?")) return;
          await fetchJson(`${API}/api/comments/${btn.dataset.deleteComment}`, { method: "DELETE" });
          loadTaskDetail(taskId);
        });
      });
      
      qs("addCommentBtn").onclick = async () => {
        const text = qs("newComment")?.value?.trim();
        if (!text) return;
        const author = String(qs("taskCommentAuthor")?.value || getSessionUser()).trim() || getSessionUser();
        await fetchJson(`${API}/api/tasks/${taskId}/comments`, {
          method: "POST",
          body: JSON.stringify({ comment: text, author })
        });
        qs("newComment").value = "";
        loadTaskDetail(taskId);
      };
      
      panel.style.display = "block";
    }
  } catch (err) {
    console.error("Failed to load task detail", err);
  }
}

async function loadTasksStats() {
  const statsEl = qs("tasksStats");
  if (!statsEl) return;
  
  try {
    const res = await fetchJson(`${API}/api/tasks/stats/summary`);
    if (res.ok) {
      setTaskApiIndicator("ok", "API online");
      statsEl.innerHTML = `
        <span class="kpi-pill"><strong>Total:</strong> ${res.total}</span>
        <span class="kpi-pill kpi-pill-blue"><strong>Open:</strong> ${res.open}</span>
        <span class="kpi-pill kpi-pill-orange"><strong>In Progress:</strong> ${res.in_progress}</span>
        <span class="kpi-pill kpi-pill-green"><strong>Done:</strong> ${res.done}</span>
        ${res.overdue ? `<span class="kpi-pill kpi-pill-red"><strong>Overdue:</strong> ${res.overdue}</span>` : ""}
      `;
    }
  } catch (err) {
    setTaskApiIndicator("down", "API unavailable");
    console.error("Failed to load tasks stats", err);
  }
}

function initTasks() {
  setTaskApiIndicator("unknown", "API checking...");
  const createBtn = qs("createTaskBtn");
  const loadBtn = qs("loadTasksBtn");
  const myTasksBtn = qs("myTasksBtn");
  const saveViewBtn = qs("saveTaskViewBtn");
  const closeDetailBtn = qs("closeTaskDetail");
  const createProjectBtn = qs("createProjectBtn");
  const statusFilterEl = qs("taskFilterStatus");
  const priorityFilterEl = qs("taskFilterPriority");
  const assignedFilterEl = qs("taskFilterAssigned");

  const refreshFilterInputs = () => {
    currentTaskView = "all";
    currentTaskSidebarActiveKey = "";
    updateSidebarActiveState("tasks");
    loadTasks();
  };

  statusFilterEl?.addEventListener("change", refreshFilterInputs);
  priorityFilterEl?.addEventListener("change", refreshFilterInputs);
  assignedFilterEl?.addEventListener("change", refreshFilterInputs);
  
  createBtn?.addEventListener("click", async () => {
    const title = qs("taskTitle")?.value?.trim();
    if (!title) {
      alert("Task title is required");
      return;
    }
    
    const editId = qs("taskTitle").dataset.editId;
    const data = {
      title,
      description: qs("taskDescription")?.value?.trim() || null,
      project: qs("taskProject")?.value?.trim() || null,
      priority: qs("taskPriority")?.value || "medium",
      assigned_to: qs("taskAssigned")?.value?.trim() || null,
      due_date: qs("taskDueDate").value || null
    };
    
    try {
      const notifyAssignee = qs("taskNotifyAssignee")?.checked !== false;
      const actor = getSessionUser();
      const assignee = String(data.assigned_to || "").trim();
      const watchers = String(qs("taskWatchers")?.value || "")
        .split(",")
        .map((w) => w.trim())
        .filter(Boolean);
      let savedTaskId = null;
      if (editId) {
        const updated = await fetchJson(`${API}/api/tasks/${editId}`, { method: "PUT", body: JSON.stringify(data) });
        savedTaskId = Number(updated?.task?.id || editId || 0);
        delete qs("taskTitle").dataset.editId;
        createBtn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="16"></line><line x1="8" y1="12" x2="16" y2="12"></line></svg> Create Task`;
      } else {
        const created = await fetchJson(`${API}/api/tasks`, { method: "POST", body: JSON.stringify(data) });
        savedTaskId = Number(created?.task?.id || 0);
      }

      if (notifyAssignee && savedTaskId && assignee) {
        const msg = assignee === actor
          ? `@${assignee} self-assigned this task.`
          : `@${assignee} assigned by @${actor}.`;
        await fetchJson(`${API}/api/tasks/${savedTaskId}/comments`, {
          method: "POST",
          body: JSON.stringify({ comment: msg, author: actor })
        });
      }
      if (notifyAssignee && savedTaskId && watchers.length) {
        const uniqueWatchers = Array.from(new Set(watchers)).filter((w) => w !== assignee);
        if (uniqueWatchers.length) {
          const watchMsg = `Watchers added by @${actor}: ${uniqueWatchers.map((w) => `@${w}`).join(" ")}`;
          await fetchJson(`${API}/api/tasks/${savedTaskId}/comments`, {
            method: "POST",
            body: JSON.stringify({ comment: watchMsg, author: actor })
          });
        }
      }
      
      qs("taskTitle").value = "";
      qs("taskDescription").value = "";
      qs("taskProject").value = "";
      qs("taskPriority").value = "medium";
      qs("taskAssigned").value = "";
      qs("taskWatchers").value = "";
      qs("taskDueDate").value = "";
      if (qs("taskNotifyAssignee")) qs("taskNotifyAssignee").checked = true;
      
      loadTasks();
      loadTasksStats();
      loadProjects();
      setStatus(editId ? "Task updated." : "Task created.");
    } catch (err) {
      alert(`Failed to save task: ${err?.message || err}`);
      setStatus("Task save failed.");
    }
  });
  
  loadBtn?.addEventListener("click", loadTasks);
  
  myTasksBtn?.addEventListener("click", async () => {
    const username = prompt("Enter your username:");
    if (!username) return;
    qs("taskFilterAssigned").value = username;
    currentProjectFilter = "";
    currentTaskView = "all";
    currentTaskSidebarActiveKey = "";
    updateSidebarActiveState("tasks");
    await loadProjects();
    loadTasks();
  });

  saveViewBtn?.addEventListener("click", () => {
    const name = String(prompt("Saved view name:") || "").trim();
    if (!name) return;
    const status = qs("taskFilterStatus")?.value || "";
    const priority = qs("taskFilterPriority")?.value || "";
    const assigned = qs("taskFilterAssigned")?.value || "";
    const existing = getSavedTaskViews();
    const view = {
      id: `${Date.now()}`,
      name,
      view: currentTaskView || "all",
      status,
      priority,
      assigned,
      project: currentProjectFilter || ""
    };
    existing.push(view);
    persistSavedTaskViews(existing.slice(-12));
    renderTaskWorkspaceSavedViews();
  });
  
  closeDetailBtn?.addEventListener("click", () => {
    const content = qs("taskDetailContent");
    if (content) content.innerHTML = `<small class="muted">Select a task from the list to open shared comments and collaboration tools.</small>`;
    const comments = qs("taskComments");
    if (comments) comments.innerHTML = `<small class="muted">No task selected.</small>`;
    currentTaskId = null;
    loadTasks();
  });
  
  createProjectBtn?.addEventListener("click", async () => {
    const name = qs("newProjectName")?.value?.trim();
    if (!name) {
      alert("Project name is required");
      return;
    }
    
    try {
      await fetchJson(`${API}/api/projects`, {
        method: "POST",
        body: JSON.stringify({ name, description: "", color: "#3b82f6" })
      });
      qs("newProjectName").value = "";
      loadProjects();
    } catch (err) {
      alert("Failed to create project");
    }
  });
  
  loadTasks();
  loadTasksStats();
  loadProjects();
  loadTeamMembers();
  renderTaskWorkspaceSavedViews();
}
/* =====================================================================
   FINANCE INTEGRATION (summarized journal posting, period lock,
   budget vs actual, rolling forecast, SSOT report + KPI definitions)
===================================================================== */
