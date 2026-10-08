// IRONLOG/web/app/collab-inbox.js — floating inbox (bottom right of every page).
// Part of the main app; index.html loads these files in order and they share one global scope.
//
// Shows the signed-in user's collaboration inbox: tasks assigned to them and
// @mentions. The bubble shows the unread count; tap to open or collapse.
// Tapping an item opens the task in the Tasks page.

const CollabInbox = (() => {
  const OPEN_KEY = "ironlog_collab_inbox_open";
  const POLL_MS = 60 * 1000;
  let items = [];
  let unread = 0;
  let open = false;
  let timer = null;
  let root = null;

  const KIND = {
    assigned: (a) => `${a} assigned you a task`,
    mention: (a) => `${a} mentioned you`,
  };

  function ago(ts) {
    const t = new Date(String(ts || "").replace(" ", "T") + "Z");
    const s = Math.max(0, (Date.now() - t.getTime()) / 1000);
    if (!Number.isFinite(s)) return "";
    if (s < 60) return "just now";
    if (s < 3600) return `${Math.round(s / 60)} min`;
    if (s < 86400) return `${Math.round(s / 3600)} h`;
    return t.toLocaleDateString(undefined, { day: "numeric", month: "short" });
  }
  const who = (u) => (typeof twName === "function" ? twName(u) : u) || "Someone";

  function build() {
    if (root) return;
    root = document.createElement("div");
    root.className = "ci";
    root.innerHTML = `
      <div class="ci-panel" id="ciPanel" hidden>
        <div class="ci-head">
          <b>Inbox</b>
          <span class="ci-head-actions">
            <button type="button" class="ci-link" data-ci="readall">Mark all read</button>
            <button type="button" class="ci-x" data-ci="close" aria-label="Collapse inbox">⌄</button>
          </span>
        </div>
        <div class="ci-list" id="ciList"></div>
        <div class="ci-foot"><button type="button" class="ci-link" data-ci="tasks">Open Tasks</button></div>
      </div>
      <button type="button" class="ci-fab" id="ciFab" aria-label="Inbox" title="Inbox: tasks assigned to you and @mentions">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 12h-6l-2 3h-4l-2-3H2"></path><path d="M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"></path></svg>
        <span class="ci-badge" id="ciBadge" hidden>0</span>
      </button>`;
    document.body.appendChild(root);
    root.addEventListener("click", onClick);
  }

  function render() {
    if (!root) return;
    const badge = root.querySelector("#ciBadge");
    badge.hidden = !unread;
    badge.textContent = unread > 99 ? "99+" : String(unread);
    root.querySelector("#ciFab").classList.toggle("has-new", unread > 0);
    root.querySelector("#ciPanel").hidden = !open;
    root.querySelector("#ciFab").setAttribute("aria-expanded", open ? "true" : "false");
    const list = root.querySelector("#ciList");
    list.innerHTML = items.length
      ? items.map((i) => `
        <button type="button" class="ci-item ${i.read_at ? "" : "new"}" data-ci-item="${i.id}" data-task="${i.task_id || ""}">
          <span class="ci-dot"></span>
          <span class="ci-main">
            <span class="ci-line"><b>${escapeHtml((KIND[i.kind] || ((a) => `${a} updated a task`))(who(i.actor)))}</b><span class="ci-time">${escapeHtml(ago(i.created_at))}</span></span>
            <span class="ci-task">#${i.task_id || ""} ${escapeHtml(i.task_title || "")}${i.task_status === "done" ? " · done" : ""}</span>
            ${i.snippet ? `<span class="ci-snip">${escapeHtml(i.snippet)}</span>` : ""}
          </span>
        </button>`).join("")
      : `<div class="ci-empty">Nothing yet. When someone assigns you a task or @mentions you, it shows up here.</div>`;
  }

  async function refresh() {
    if (!getSessionUser()) return;
    try {
      const res = await fetchJson(`${API}/api/collab/inbox?limit=30`);
      items = res.items || [];
      unread = Number(res.unread || 0);
      render();
    } catch { /* offline or signed out: keep what is shown */ }
  }

  async function markRead(body) {
    try {
      const res = await fetchJson(`${API}/api/collab/inbox/read`, { method: "POST", body: JSON.stringify(body) });
      unread = Number(res.unread || 0);
    } catch { /* try again on next refresh */ }
  }

  async function onClick(e) {
    const fab = e.target.closest("#ciFab");
    const act = e.target.closest("[data-ci]")?.dataset.ci;
    const item = e.target.closest("[data-ci-item]");
    if (fab) return setOpen(!open);
    if (act === "close") return setOpen(false);
    if (act === "readall") {
      await markRead({ all: true });
      items = items.map((i) => ({ ...i, read_at: i.read_at || new Date().toISOString() }));
      return render();
    }
    if (act === "tasks") {
      setOpen(false);
      return typeof switchTab === "function" && switchTab("tasks");
    }
    if (item) {
      const taskId = Number(item.dataset.task);
      await markRead({ task_id: taskId });
      items = items.map((i) => (i.task_id === taskId ? { ...i, read_at: i.read_at || new Date().toISOString() } : i));
      setOpen(false);
      if (taskId && typeof showTask === "function") showTask(taskId);
    }
  }

  function setOpen(v) {
    open = Boolean(v);
    try { localStorage.setItem(OPEN_KEY, open ? "1" : "0"); } catch { /* storage off */ }
    render();
    if (open) refresh();
  }

  function start() {
    build();
    try { open = localStorage.getItem(OPEN_KEY) === "1"; } catch { open = false; }
    render();
    refresh();
    clearInterval(timer);
    timer = setInterval(() => { if (!document.hidden) refresh(); }, POLL_MS);
    document.addEventListener("visibilitychange", () => { if (!document.hidden) refresh(); });
    document.addEventListener("keydown", (e) => { if (e.key === "Escape" && open) setOpen(false); });
  }

  return { start, refresh };
})();
window.CollabInbox = CollabInbox;

/** Start-up: floating inbox. Called once from init() in init.js. */
function initCollabInbox() {
  CollabInbox.start();
}
