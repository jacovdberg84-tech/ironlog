import { db } from "../db/client.js";

function getSiteCode(req) {
  return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
}

export default async function tasksRoutes(app) {
  const SITE = "main";

  app.get("/tasks", async (req, reply) => {
    const site_code = getSiteCode(req);
    const status = req.query?.status;
    const priority = req.query?.priority;
    const assigned = req.query?.assigned;
    
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
    if (assigned) {
      sql += " AND assigned_to = ?";
      params.push(assigned);
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
    return { ok: true, task };
  });

  app.post("/tasks", async (req, reply) => {
    const { title, description, status, priority, project, assigned_to, due_date } = req.body || {};
    
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
      SITE,
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
    
    db.prepare("DELETE FROM tasks WHERE id = ?").run(id);
    return { ok: true, deleted: true };
  });

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
    
    return {
      ok: true,
      by_status: stats,
      by_priority: priorities,
      total: stats.reduce((sum, s) => sum + s.count, 0),
      open: stats.find(s => s.status === "open")?.count || 0,
      in_progress: stats.find(s => s.status === "in_progress")?.count || 0,
      done: stats.find(s => s.status === "done")?.count || 0
    };
  });
}
