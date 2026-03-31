import { db } from "../db/client.js";

const VALID_ROLES = ["admin", "supervisor", "stores", "artisan", "operator"];

export default async function authRoutes(app) {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT NOT NULL UNIQUE,
      full_name TEXT,
      role TEXT NOT NULL DEFAULT 'operator',
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE INDEX IF NOT EXISTS idx_users_role
    ON users(role)
  `).run();

  const ensureAdmin = db.prepare(`
    INSERT INTO users (username, full_name, role, active)
    SELECT 'admin', 'System Admin', 'admin', 1
    WHERE NOT EXISTS (SELECT 1 FROM users WHERE username = 'admin')
  `);
  ensureAdmin.run();

  const getByUsername = db.prepare(`
    SELECT id, username, full_name, role, active, created_at
    FROM users
    WHERE username = ?
  `);

  function getRequestRole(req) {
    const raw = String(req.headers["x-user-role"] || "").trim().toLowerCase();
    return VALID_ROLES.includes(raw) ? raw : "admin";
  }

  function getRequestUsername(req) {
    return String(req.headers["x-user-name"] || "").trim();
  }

  function requireAdmin(req, reply) {
    const role = getRequestRole(req);
    if (role !== "admin") {
      reply.code(403).send({ error: "admin role required" });
      return false;
    }
    return true;
  }

  // GET /api/auth/me
  app.get("/me", async (req) => {
    const username = getRequestUsername(req);
    const roleFromHeader = getRequestRole(req);

    if (username) {
      const user = getByUsername.get(username);
      if (user && Number(user.active) === 1) {
        return {
          ok: true,
          user: {
            id: Number(user.id),
            username: user.username,
            full_name: user.full_name,
            role: user.role,
            active: Number(user.active),
          },
          roles: VALID_ROLES,
        };
      }
    }

    return {
      ok: true,
      user: {
        id: null,
        username: username || "session-user",
        full_name: null,
        role: roleFromHeader,
        active: 1,
      },
      roles: VALID_ROLES,
    };
  });

  // GET /api/auth/users (admin only)
  app.get("/users", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const rows = db.prepare(`
      SELECT id, username, full_name, role, active, created_at
      FROM users
      ORDER BY username ASC
    `).all().map((r) => ({
      ...r,
      id: Number(r.id),
      active: Number(r.active),
    }));
    return { ok: true, rows };
  });

  // POST /api/auth/users (admin only)
  // Body: { username, full_name?, role, active? }
  app.post("/users", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const body = req.body || {};
    const username = String(body.username || "").trim();
    const full_name = String(body.full_name || "").trim() || null;
    const role = String(body.role || "").trim().toLowerCase();
    const active = body.active === 0 || body.active === false ? 0 : 1;

    if (!username) return reply.code(400).send({ error: "username is required" });
    if (!VALID_ROLES.includes(role)) {
      return reply.code(400).send({ error: `role must be one of: ${VALID_ROLES.join(", ")}` });
    }

    db.prepare(`
      INSERT INTO users (username, full_name, role, active)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(username) DO UPDATE SET
        full_name = excluded.full_name,
        role = excluded.role,
        active = excluded.active
    `).run(username, full_name, role, active);

    const user = getByUsername.get(username);
    return {
      ok: true,
      user: {
        id: Number(user.id),
        username: user.username,
        full_name: user.full_name,
        role: user.role,
        active: Number(user.active),
      },
    };
  });
}
