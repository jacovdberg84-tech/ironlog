import { db } from "../db/client.js";

function getSiteCode(req) {
  return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
}

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
        req.headers["x-user"] || "system"
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

  app.get("/tasks", async (req, reply) => {
    const site_code = getSiteCode(req);
    const status = req.query?.status;
    const priority = req.query?.priority;
    const assigned = req.query?.assigned;
    const project = req.query?.project;
    const my_tasks = req.query?.my_tasks === "true";
    const user = req.headers["x-user"];
    
    let sql = "SELECT * FROM tasks WHERE site_code = ?";
    const params = [site_code];
    
    if (status) {
      sql += " AND status = ?";
      params.push(status);
    }
    if (priority) {
      sql += " AND priority = ?";
      params.push(priority);
    }
    if (project) {
      sql += " AND project = ?";
      params.push(project);
    }
    if (assigned) {
      sql += " AND assigned_to = ?";
      params.push(assigned);
    }
    if (my_tasks && user) {
      sql += " AND assigned_to = ?";
      params.push(user);
    }
    
    sql += " ORDER BY CASE priority WHEN 'high' THEN 1 WHEN 'medium' THEN 2 WHEN 'low' THEN 3 ELSE 4 END, created_at DESC";
    
    const tasks = db.prepare(sql).all(...params);
    return { ok: true, tasks };
  });

  app.get("/tasks/:id", async (req, reply) => {
    const task = db.prepare("SELECT * FROM tasks WHERE id = ?").get(req.params.id);
    if (!task) {
      return reply.code(404).send({ error: "Task not found" });
    }
    
    const comments = db.prepare(`
      SELECT * FROM task_comments 
      WHERE task_id = ? 
      ORDER BY created_at ASC
    `).all(req.params.id);
    
    return { ok: true, task, comments };
  });

  app.post("/tasks", async (req, reply) => {
    const { title, description, status, priority, project, assigned_to, due_date } = req.body || {};
    const site_code = getSiteCode(req);
    
    if (!title || !String(title).trim()) {
      return reply.code(400).send({ error: "Title is required" });
    }
    
    const result = db.prepare(`
      INSERT INTO tasks (title, description, status, priority, project, assigned_to, due_date, site_code, created_by)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(
      String(title).trim(),
      description || null,
      status || "open",
      priority || "medium",
      project || null,
      assigned_to || null,
      due_date || null,
      site_code,
      req.headers["x-user"] || "system"
    );
    
    const task = db.prepare("SELECT * FROM tasks WHERE id = ?").get(result.lastInsertRowid);
    return { ok: true, task };
  });

  app.put("/tasks/:id", async (req, reply) => {
    const { title, description, status, priority, project, assigned_to, due_date } = req.body || {};
    const id = req.params.id;
    
    const existing = db.prepare("SELECT * FROM tasks WHERE id = ?").get(id);
    if (!existing) {
      return reply.code(404).send({ error: "Task not found" });
    }
    
    const updates = [];
    const params = [];
    
    if (title !== undefined) {
      if (!String(title).trim()) {
        return reply.code(400).send({ error: "Title cannot be empty" });
      }
      updates.push("title = ?");
      params.push(String(title).trim());
    }
    if (description !== undefined) {
      updates.push("description = ?");
      params.push(description);
    }
    if (status !== undefined) {
      updates.push("status = ?");
      params.push(status);
      if (status === "done" && !existing.completed_at) {
        updates.push("completed_at = datetime('now')");
      } else if (status !== "done") {
        updates.push("completed_at = NULL");
      }
    }
    if (priority !== undefined) {
      updates.push("priority = ?");
      params.push(priority);
    }
    if (project !== undefined) {
      updates.push("project = ?");
      params.push(project);
    }
    if (assigned_to !== undefined) {
      updates.push("assigned_to = ?");
      params.push(assigned_to);
    }
    if (due_date !== undefined) {
      updates.push("due_date = ?");
      params.push(due_date);
    }
    
    updates.push("updated_at = datetime('now')");
    params.push(id);
    
    db.prepare(`UPDATE tasks SET ${updates.join(", ")} WHERE id = ?`).run(...params);
    
    const task = db.prepare("SELECT * FROM tasks WHERE id = ?").get(id);
    return { ok: true, task };
  });

  app.delete("/tasks/:id", async (req, reply) => {
    const id = req.params.id;
    const existing = db.prepare("SELECT * FROM tasks WHERE id = ?").get(id);
    if (!existing) {
      return reply.code(404).send({ error: "Task not found" });
    }
    
    db.prepare("DELETE FROM task_comments WHERE task_id = ?").run(id);
    db.prepare("DELETE FROM tasks WHERE id = ?").run(id);
    return { ok: true, deleted: true };
  });

  // =========================
  // COMMENTS
  // =========================

  app.post("/tasks/:id/comments", async (req, reply) => {
    const task_id = req.params.id;
    const { comment } = req.body || {};
    
    const task = db.prepare("SELECT * FROM tasks WHERE id = ?").get(task_id);
    if (!task) {
      return reply.code(404).send({ error: "Task not found" });
    }
    
    if (!comment || !String(comment).trim()) {
      return reply.code(400).send({ error: "Comment cannot be empty" });
    }
    
    const result = db.prepare(`
      INSERT INTO task_comments (task_id, comment, author)
      VALUES (?, ?, ?)
    `).run(
      task_id,
      String(comment).trim(),
      req.headers["x-user"] || "anonymous"
    );
    
    const newComment = db.prepare("SELECT * FROM task_comments WHERE id = ?").get(result.lastInsertRowid);
    return { ok: true, comment: newComment };
  });

  app.delete("/comments/:id", async (req, reply) => {
    const id = req.params.id;
    const existing = db.prepare("SELECT * FROM task_comments WHERE id = ?").get(id);
    if (!existing) {
      return reply.code(404).send({ error: "Comment not found" });
    }
    
    db.prepare("DELETE FROM task_comments WHERE id = ?").run(id);
    return { ok: true, deleted: true };
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
