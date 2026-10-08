// IRONLOG/api/utils/taskCollab.js — who gets told about a task.
//
// Each user has an inbox. A task assigned to you, or an @mention of you in a
// task or comment, puts an item there. Comments are always signed by the
// person signed in; nobody posts as someone else.

const ready = new WeakSet();

export function ensureTaskCollabSchema(db) {
  if (ready.has(db)) return;
  db.exec(`
    CREATE TABLE IF NOT EXISTS task_watchers (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      task_id INTEGER NOT NULL,
      username TEXT NOT NULL,
      added_by TEXT,
      added_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(task_id, username)
    );
    CREATE TABLE IF NOT EXISTS collab_inbox (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT NOT NULL,
      kind TEXT NOT NULL,
      task_id INTEGER,
      comment_id INTEGER,
      actor TEXT,
      task_title TEXT,
      snippet TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      read_at TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_collab_inbox_user ON collab_inbox(username, read_at);
    CREATE INDEX IF NOT EXISTS idx_collab_inbox_task ON collab_inbox(task_id);
  `);
  ready.add(db);
}

/** The signed-in user (the auth hook sets x-user-name from the session token). */
export function requestUser(req) {
  return String(req.headers["x-user-name"] || req.headers["x-user"] || "").trim();
}

/** Active users by lower-case username → the username as stored. */
function userIndex(db) {
  const rows = db.prepare(`SELECT username FROM users WHERE COALESCE(active, 1) = 1`).all();
  return new Map(rows.map((r) => [String(r.username).toLowerCase(), r.username]));
}

export function resolveUser(db, name) {
  const key = String(name || "").trim().replace(/^@/, "").toLowerCase();
  if (!key) return null;
  return userIndex(db).get(key) || null;
}

/** Usernames @mentioned in a text that belong to real, active users. */
export function mentionedUsers(db, text) {
  const users = userIndex(db);
  const out = new Set();
  for (const m of String(text || "").matchAll(/(^|[^\w@])@([A-Za-z0-9._-]{2,40})/g)) {
    const name = m[2].replace(/[._-]+$/, "").toLowerCase();
    if (users.has(name)) out.add(users.get(name));
  }
  return [...out];
}

const snippetOf = (text) => {
  const t = String(text || "").replace(/\s+/g, " ").trim();
  return t.length > 160 ? `${t.slice(0, 157)}…` : t;
};

/** Put an item in someone's inbox. Nobody is told about their own action. */
export function notify(db, { username, kind, task, actor, text = "", commentId = null }) {
  const to = resolveUser(db, username);
  if (!to || !task) return false;
  if (actor && to.toLowerCase() === String(actor).toLowerCase()) return false;
  db.prepare(`
    INSERT INTO collab_inbox (username, kind, task_id, comment_id, actor, task_title, snippet)
    VALUES (?, ?, ?, ?, ?, ?, ?)
  `).run(to, kind, task.id, commentId, actor || null, task.title || "", snippetOf(text));
  return true;
}

/** Inbox items for a task assignment and the @mentions in its text. */
export function notifyTaskChange(db, { task, actor, assignedChanged, text }) {
  let n = 0;
  const assignee = resolveUser(db, task.assigned_to);
  if (assignedChanged && assignee) n += notify(db, { username: assignee, kind: "assigned", task, actor, text: task.description || "" }) ? 1 : 0;
  for (const u of mentionedUsers(db, text)) {
    if (assignedChanged && assignee && u === assignee) continue; // already told
    n += notify(db, { username: u, kind: "mention", task, actor, text }) ? 1 : 0;
  }
  return n;
}

export function setWatchers(db, taskId, names, addedBy) {
  const list = (Array.isArray(names) ? names : String(names || "").split(/[,;\s]+/))
    .map((n) => resolveUser(db, n))
    .filter(Boolean);
  const add = db.prepare(`INSERT OR IGNORE INTO task_watchers (task_id, username, added_by) VALUES (?, ?, ?)`);
  for (const u of new Set(list)) add.run(taskId, u, addedBy || null);
  return [...new Set(list)];
}

export function watchersOf(db, taskId) {
  return db.prepare(`SELECT username FROM task_watchers WHERE task_id = ? ORDER BY username`).all(taskId).map((r) => r.username);
}
