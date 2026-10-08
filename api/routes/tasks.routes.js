import { db } from "../db/client.js";
import { getRoles, getSiteCode } from "../utils/request.js";
import { ensureTaskCollabSchema, mentionedUsers, notify, notifyTaskChange, requestUser, resolveUser, setWatchers, watchersOf } from "../utils/taskCollab.js";

export default async function tasksRoutes(app) {
  // =========================
  // PROJECTS
  // =========================
  
  app.get("/projects", async (req, reply) => {
    const site_code = getSiteCode(req);
    const projects = db.prepare(`
      SELECT p.*, COUNT(t.id) as task_count,
        SUM(CASE WHEN t.status = 'done' THEN 1 ELSE 0 END) as done_count
      FROM projects p
      LEFT JOIN tasks t ON t.project = p.name AND t.site_code = p.site_code
      WHERE p.site_code = ?
      GROUP BY p.id
      ORDER BY p.name ASC
    `).all(site_code);
    return { ok: true, projects };
  });

  app.post("/projects", async (req, reply) => {
    const { name, description, color } = req.body || {};
    const site_code = getSiteCode(req);
    if (!name || !String(name).trim()) {
      return reply.code(400).send({ error: "Project name is required" });
    }
    
    try {
      const result = db.prepare(`
        INSERT INTO projects (name, description, color, site_code, created_by)
        VALUES (?, ?, ?, ?, ?)
      `).run(
        String(name).trim(),
        description || null,
        color || "#3b82f6",
        site_code,
        requestUser(req) || "system"
      );
      
      const project = db.prepare("SELECT * FROM projects WHERE id = ?").get(result.lastInsertRowid);
      return { ok: true, project };
    } catch (err) {
      if (err.message.includes("UNIQUE")) {
        return reply.code(400).send({ error: "Project with this name already exists" });
      }
      throw err;
    }
  });

  app.delete("/projects/:id", async (req, reply) => {
    const id = req.params.id;
    const existing = db.prepare("SELECT * FROM projects WHERE id = ?").get(id);
    if (!existing) {
      return reply.code(404).send({ error: "Project not found" });
    }
    
    db.prepare("DELETE FROM projects WHERE id = ?").run(id);
    return { ok: true, deleted: true };
  });

  // =========================
  // TASKS
  // =========================

  ensureTaskCollabSchema(db);
  const STAFF_ROLES = ["admin", "supervisor", "workshop_admin", "plant_manager", "site_manager"];
  const taskRow = (id) => db.prepare("SELECT * FROM tasks WHERE id = ?").get(id);
  const isStaff = (req) => getRoles(req, { fallback: "admin" }).some((r) => STAFF_ROLES.includes(r));

  // GET /api/tasks?view=&q=&status=&priority=&project=&assigned=
  // view: all | mine (assigned to me) | created (I assigned) | watching | involved | overdue | done | open
  app.get("/tasks", async (req) => {
    const site_code = getSiteCode(req);
    const me = requestUser(req);
    const q = req.query || {};
    const where = ["t.site_code = ?"];
    const params = [site_code];
    const view = String(q.view || (q.my_tasks === "true" ? "mine" : "all")).toLowerCase();
    const meLower = me.toLowerCase();
    const watchingSql = `EXISTS (SELECT 1 FROM task_watchers w WHERE w.task_id = t.id AND LOWER(w.username) = ?)`;
    const mentionedSql = `EXISTS (SELECT 1 FROM collab_inbox i WHERE i.task_id = t.id AND LOWER(i.username) = ?)`;
    if (view === "mine") { where.push("LOWER(COALESCE(t.assigned_to, '')) = ?"); params.push(meLower); }
    else if (view === "created") { where.push("LOWER(COALESCE(t.created_by, '')) = ?"); params.push(meLower); }
    else if (view === "watching") { where.push(watchingSql); params.push(meLower); }
    else if (view === "involved") {
      where.push(`(LOWER(COALESCE(t.assigned_to, '')) = ? OR LOWER(COALESCE(t.created_by, '')) = ? OR ${watchingSql} OR ${mentionedSql})`);
      params.push(meLower, meLower, meLower, meLower);
    } else if (view === "overdue") where.push("t.status != 'done' AND t.due_date IS NOT NULL AND t.due_date < date('now')");
    if (view === "done") where.push("t.status = 'done'");
    else if (view === "open" || (["mine", "created", "watching", "involved", "overdue"].includes(view) && q.include_done !== "1")) where.push("t.status != 'done'");
    if (q.status) { where.push("t.status = ?"); params.push(String(q.status)); }
    if (q.priority) { where.push("t.priority = ?"); params.push(String(q.priority)); }
    if (q.project) { where.push("t.project = ?"); params.push(String(q.project)); }
    if (q.assigned) { where.push("LOWER(COALESCE(t.assigned_to, '')) = LOWER(?)"); params.push(String(q.assigned)); }
    const term = String(q.q || "").trim();
    if (term) {
      where.push("(t.title LIKE ? OR COALESCE(t.description, '') LIKE ? OR COALESCE(t.assigned_to, '') LIKE ? OR COALESCE(t.project, '') LIKE ? OR CAST(t.id AS TEXT) = ?)");
      params.push(`%${term}%`, `%${term}%`, `%${term}%`, `%${term}%`, term.replace(/^#/, ""));
    }
    const tasks = db.prepare(`
      SELECT t.*,
        (SELECT COUNT(*) FROM task_comments c WHERE c.task_id = t.id) AS comments_count,
        (SELECT MAX(c.created_at) FROM task_comments c WHERE c.task_id = t.id) AS last_comment_at,
        ${watchingSql} AS watching
      FROM tasks t
      WHERE ${where.join(" AND ")}
      ORDER BY CASE WHEN t.status = 'done' THEN 1 ELSE 0 END,
        CASE WHEN t.due_date IS NOT NULL AND t.due_date < date('now') AND t.status != 'done' THEN 0 ELSE 1 END,
        CASE t.priority WHEN 'high' THEN 1 WHEN 'medium' THEN 2 WHEN 'low' THEN 3 ELSE 4 END,
        COALESCE(t.due_date, '9999-12-31') ASC, t.updated_at DESC
      LIMIT 500
    `).all(meLower, ...params);
    return { ok: true, tasks: tasks.map((t) => ({ ...t, watching: Boolean(t.watching) })) };
  });

  app.get("/tasks/:id", async (req, reply) => {
    const task = taskRow(req.params.id);
    if (!task) return reply.code(404).send({ error: "Task not found" });
    const comments = db.prepare(`SELECT * FROM task_comments WHERE task_id = ? ORDER BY created_at ASC, id ASC`).all(req.params.id);
    const watchers = watchersOf(db, task.id);
    const me = requestUser(req);
    return { ok: true, task, comments, watchers, watching: watchers.some((w) => w.toLowerCase() === me.toLowerCase()) };
  });

  app.post("/tasks", async (req, reply) => {
    const { title, description, status, priority, project, assigned_to, due_date, watchers } = req.body || {};
    const site_code = getSiteCode(req);
    const me = requestUser(req) || "system";
    if (!title || !String(title).trim()) return reply.code(400).send({ error: "Title is required" });
    const assignee = assigned_to ? (resolveUser(db, assigned_to) || String(assigned_to).trim()) : null;
    const result = db.prepare(`
      INSERT INTO tasks (title, description, status, priority, project, assigned_to, due_date, site_code, created_by)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(String(title).trim(), description || null, status || "open", priority || "medium", project || null, assignee, due_date || null, site_code, me);
    const task = taskRow(result.lastInsertRowid);
    if (watchers) setWatchers(db, task.id, watchers, me);
    notifyTaskChange(db, { task, actor: me, assignedChanged: Boolean(assignee), text: `${task.title}\n${task.description || ""}` });
    return { ok: true, task, watchers: watchersOf(db, task.id) };
  });

  app.put("/tasks/:id", async (req, reply) => {
    const { title, description, status, priority, project, assigned_to, due_date, watchers } = req.body || {};
    const id = req.params.id;
    const me = requestUser(req) || "system";
    const existing = taskRow(id);
    if (!existing) return reply.code(404).send({ error: "Task not found" });

    const updates = [];
    const params = [];
    if (title !== undefined) {
      if (!String(title).trim()) return reply.code(400).send({ error: "Title cannot be empty" });
      updates.push("title = ?");
      params.push(String(title).trim());
    }
    if (description !== undefined) { updates.push("description = ?"); params.push(description); }
    if (status !== undefined) {
      updates.push("status = ?");
      params.push(status);
      if (status === "done" && !existing.completed_at) updates.push("completed_at = datetime('now')");
      else if (status !== "done") updates.push("completed_at = NULL");
    }
    if (priority !== undefined) { updates.push("priority = ?"); params.push(priority); }
    if (project !== undefined) { updates.push("project = ?"); params.push(project); }
    let assignedChanged = false;
    if (assigned_to !== undefined) {
      const assignee = assigned_to ? (resolveUser(db, assigned_to) || String(assigned_to).trim()) : null;
      assignedChanged = Boolean(assignee) && String(assignee).toLowerCase() !== String(existing.assigned_to || "").toLowerCase();
      updates.push("assigned_to = ?");
      params.push(assignee);
    }
    if (due_date !== undefined) { updates.push("due_date = ?"); params.push(due_date); }
    updates.push("updated_at = datetime('now')");
    params.push(id);
    db.prepare(`UPDATE tasks SET ${updates.join(", ")} WHERE id = ?`).run(...params);
    const task = taskRow(id);
    if (Array.isArray(watchers) || typeof watchers === "string") {
      db.prepare(`DELETE FROM task_watchers WHERE task_id = ?`).run(task.id);
      setWatchers(db, task.id, watchers, me);
    }
    // Only newly written text is checked for @mentions, so an edit does not repeat old ones.
    const before = `${existing.title}\n${existing.description || ""}`;
    const after = `${task.title}\n${task.description || ""}`;
    const fresh = mentionedUsers(db, after).filter((u) => !mentionedUsers(db, before).includes(u)).map((u) => `@${u}`).join(" ");
    notifyTaskChange(db, { task, actor: me, assignedChanged, text: fresh ? `${fresh} ${task.title}` : "" });
    return { ok: true, task, watchers: watchersOf(db, task.id) };
  });

  // POST /api/tasks/:id/watch { watch: true|false } — follow or unfollow a task.
  app.post("/tasks/:id/watch", async (req, reply) => {
    const me = requestUser(req);
    const task = taskRow(req.params.id);
    if (!task) return reply.code(404).send({ error: "Task not found" });
    if (!me) return reply.code(401).send({ error: "Sign in first" });
    if (req.body?.watch === false) db.prepare(`DELETE FROM task_watchers WHERE task_id = ? AND LOWER(username) = LOWER(?)`).run(task.id, me);
    else db.prepare(`INSERT OR IGNORE INTO task_watchers (task_id, username, added_by) VALUES (?, ?, ?)`).run(task.id, resolveUser(db, me) || me, me);
    return { ok: true, watchers: watchersOf(db, task.id) };
  });

  app.delete("/tasks/:id", async (req, reply) => {
    const id = req.params.id;
    const existing = taskRow(id);
    if (!existing) return reply.code(404).send({ error: "Task not found" });
    const me = requestUser(req);
    if (!isStaff(req) && String(existing.created_by || "").toLowerCase() !== me.toLowerCase()) {
      return reply.code(403).send({ error: "Only the person who created the task or a supervisor can delete it" });
    }
    db.prepare("DELETE FROM task_comments WHERE task_id = ?").run(id);
    db.prepare("DELETE FROM task_watchers WHERE task_id = ?").run(id);
    db.prepare("DELETE FROM collab_inbox WHERE task_id = ?").run(id);
    db.prepare("DELETE FROM tasks WHERE id = ?").run(id);
    return { ok: true, deleted: true };
  });

  // =========================
  // COMMENTS
  // =========================

  // POST /api/tasks/:id/comments { comment } — signed by the signed-in user.
  app.post("/tasks/:id/comments", async (req, reply) => {
    const task = taskRow(req.params.id);
    if (!task) return reply.code(404).send({ error: "Task not found" });
    const { comment } = req.body || {};
    if (!comment || !String(comment).trim()) return reply.code(400).send({ error: "Comment cannot be empty" });
    const me = requestUser(req) || String(req.body?.author || "").trim() || "anonymous";
    const text = String(comment).trim();
    const result = db.prepare(`INSERT INTO task_comments (task_id, comment, author) VALUES (?, ?, ?)`).run(task.id, text, me);
    db.prepare(`UPDATE tasks SET updated_at = datetime('now') WHERE id = ?`).run(task.id);
    const commentId = Number(result.lastInsertRowid);
    for (const u of mentionedUsers(db, text)) notify(db, { username: u, kind: "mention", task, actor: me, text, commentId });
    return { ok: true, comment: db.prepare("SELECT * FROM task_comments WHERE id = ?").get(commentId) };
  });

  app.delete("/comments/:id", async (req, reply) => {
    const existing = db.prepare("SELECT * FROM task_comments WHERE id = ?").get(req.params.id);
    if (!existing) return reply.code(404).send({ error: "Comment not found" });
    const me = requestUser(req);
    if (!isStaff(req) && String(existing.author || "").toLowerCase() !== me.toLowerCase()) {
      return reply.code(403).send({ error: "You can only delete your own comments" });
    }
    db.prepare("DELETE FROM task_comments WHERE id = ?").run(req.params.id);
    db.prepare("DELETE FROM collab_inbox WHERE comment_id = ?").run(req.params.id);
    return { ok: true, deleted: true };
  });

  // =========================
  // INBOX (assigned to me, @mentions)
  // =========================

  // GET /api/collab/inbox?limit=30
  app.get("/collab/inbox", async (req, reply) => {
    const me = requestUser(req);
    if (!me) return reply.code(401).send({ error: "Sign in first" });
    const limit = Math.min(100, Math.max(1, Number(req.query?.limit) || 30));
    const items = db.prepare(`
      SELECT i.*, t.status AS task_status, t.due_date AS task_due_date
      FROM collab_inbox i
      LEFT JOIN tasks t ON t.id = i.task_id
      WHERE LOWER(i.username) = LOWER(?)
      ORDER BY i.id DESC
      LIMIT ?
    `).all(me, limit);
    const unread = db.prepare(`SELECT COUNT(*) AS n FROM collab_inbox WHERE LOWER(username) = LOWER(?) AND read_at IS NULL`).get(me).n;
    return { ok: true, unread, items };
  });

  // POST /api/collab/inbox/read { ids: [] } | { task_id } | { all: true }
  app.post("/collab/inbox/read", async (req, reply) => {
    const me = requestUser(req);
    if (!me) return reply.code(401).send({ error: "Sign in first" });
    const b = req.body || {};
    let changes = 0;
    if (b.all === true) {
      changes = db.prepare(`UPDATE collab_inbox SET read_at = datetime('now') WHERE LOWER(username) = LOWER(?) AND read_at IS NULL`).run(me).changes;
    } else if (b.task_id) {
      changes = db.prepare(`UPDATE collab_inbox SET read_at = datetime('now') WHERE LOWER(username) = LOWER(?) AND read_at IS NULL AND task_id = ?`).run(me, Number(b.task_id)).changes;
    } else if (Array.isArray(b.ids) && b.ids.length) {
      const ids = b.ids.map(Number).filter(Number.isFinite).slice(0, 200);
      changes = db.prepare(`UPDATE collab_inbox SET read_at = datetime('now') WHERE LOWER(username) = LOWER(?) AND read_at IS NULL AND id IN (${ids.map(() => "?").join(",")})`).run(me, ...ids).changes;
    }
    const unread = db.prepare(`SELECT COUNT(*) AS n FROM collab_inbox WHERE LOWER(username) = LOWER(?) AND read_at IS NULL`).get(me).n;
    return { ok: true, marked: changes, unread };
  });

  // =========================
  // STATS
  // =========================

  app.get("/tasks/stats/summary", async (req, reply) => {
    const site_code = getSiteCode(req);
    
    const stats = db.prepare(`
      SELECT 
        status,
        COUNT(*) as count
      FROM tasks
      WHERE site_code = ?
      GROUP BY status
    `).all(site_code);
    
    const priorities = db.prepare(`
      SELECT 
        priority,
        COUNT(*) as count
      FROM tasks
      WHERE site_code = ? AND status != 'done'
      GROUP BY priority
    `).all(site_code);
    
    const byProject = db.prepare(`
      SELECT 
        COALESCE(project, 'No Project') as project,
        COUNT(*) as total,
        SUM(CASE WHEN status = 'done' THEN 1 ELSE 0 END) as done,
        SUM(CASE WHEN status = 'open' THEN 1 ELSE 0 END) as open,
        SUM(CASE WHEN status = 'in_progress' THEN 1 ELSE 0 END) as in_progress
      FROM tasks
      WHERE site_code = ?
      GROUP BY project
      ORDER BY total DESC
    `).all(site_code);
    
    const overdue = db.prepare(`
      SELECT COUNT(*) as count FROM tasks
      WHERE site_code = ?
        AND status != 'done'
        AND due_date IS NOT NULL
        AND due_date < date('now')
    `).get(site_code);
    
    return {
      ok: true,
      by_status: stats,
      by_priority: priorities,
      by_project: byProject,
      total: stats.reduce((sum, s) => sum + s.count, 0),
      open: stats.find(s => s.status === "open")?.count || 0,
      in_progress: stats.find(s => s.status === "in_progress")?.count || 0,
      done: stats.find(s => s.status === "done")?.count || 0,
      overdue: overdue?.count || 0
    };
  });
}
