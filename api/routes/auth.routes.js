import crypto from "node:crypto";
import { db } from "../db/client.js";

export const VALID_ROLES = ["admin", "supervisor", "stores", "artisan", "operator"];

/** Tab keys that can be assigned per user (UI sections). Admin-only "admin" tab is not listed here. */
export const ASSIGNABLE_TAB_KEYS = [
  "dash",
  "daily",
  "assets",
  "maintenance",
  "fuel",
  "lube",
  "stock",
  "legal",
  "uploads",
  "reports",
  "approvals",
  "procurement",
  "operations",
  "dispatch",
  "quality",
  "audit",
  "docs",
  "Breakdowns",
  "vehicle",
];

const SESSION_DAYS = Math.min(
  365,
  Math.max(1, Number(process.env.IRONLOG_SESSION_DAYS || 14) || 14)
);

function hasColumn(table, name) {
  const rows = db.prepare(`PRAGMA table_info(${table})`).all();
  return rows.some((r) => r.name === name);
}

function hashPassword(plain) {
  const salt = crypto.randomBytes(16);
  const hash = crypto.scryptSync(String(plain), salt, 64);
  return `scrypt$${salt.toString("base64")}$${hash.toString("base64")}`;
}

function verifyPassword(plain, stored) {
  if (!stored || !plain) return false;
  const parts = String(stored).split("$");
  if (parts[0] !== "scrypt" || parts.length !== 3) return false;
  try {
    const salt = Buffer.from(parts[1], "base64");
    const expected = Buffer.from(parts[2], "base64");
    const hash = crypto.scryptSync(String(plain), salt, 64);
    if (hash.length !== expected.length) return false;
    return crypto.timingSafeEqual(hash, expected);
  } catch {
    return false;
  }
}

function parseAllowedTabs(raw) {
  if (raw == null || raw === "") return null;
  try {
    const v = typeof raw === "string" ? JSON.parse(raw) : raw;
    if (!Array.isArray(v)) return null;
    const set = new Set(ASSIGNABLE_TAB_KEYS);
    const out = v.map((x) => String(x)).filter((k) => set.has(k));
    return out.length ? out : null;
  } catch {
    return null;
  }
}

function parseRoles(raw, fallbackRole = "operator") {
  let arr = [];
  if (raw != null && raw !== "") {
    try {
      const v = typeof raw === "string" ? JSON.parse(raw) : raw;
      if (Array.isArray(v)) arr = v.map((x) => String(x || "").trim().toLowerCase());
    } catch {}
  }
  if (!arr.length && fallbackRole) arr = [String(fallbackRole || "").trim().toLowerCase()];
  const out = Array.from(new Set(arr.filter((r) => VALID_ROLES.includes(r))));
  return out.length ? out : ["operator"];
}

function pickPrimaryRole(roles) {
  const list = Array.isArray(roles) ? roles : [];
  for (const r of VALID_ROLES) {
    if (list.includes(r)) return r;
  }
  return "operator";
}

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

  if (!hasColumn("users", "password_hash")) {
    db.prepare(`ALTER TABLE users ADD COLUMN password_hash TEXT`).run();
  }
  if (!hasColumn("users", "department")) {
    db.prepare(`ALTER TABLE users ADD COLUMN department TEXT`).run();
  }
  if (!hasColumn("users", "allowed_tabs")) {
    db.prepare(`ALTER TABLE users ADD COLUMN allowed_tabs TEXT`).run();
  }
  if (!hasColumn("users", "roles_json")) {
    db.prepare(`ALTER TABLE users ADD COLUMN roles_json TEXT`).run();
  }

  db.prepare(`
    CREATE TABLE IF NOT EXISTS auth_sessions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      token TEXT NOT NULL UNIQUE,
      user_id INTEGER NOT NULL,
      expires_at TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
    )
  `).run();

  db.prepare(`CREATE INDEX IF NOT EXISTS idx_auth_sessions_token ON auth_sessions(token)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_auth_sessions_user ON auth_sessions(user_id)`).run();

  const ensureAdmin = db.prepare(`
    INSERT INTO users (username, full_name, role, active)
    SELECT 'admin', 'System Admin', 'admin', 1
    WHERE NOT EXISTS (SELECT 1 FROM users WHERE username = 'admin')
  `);
  ensureAdmin.run();

  const getByUsername = db.prepare(`
    SELECT id, username, full_name, role, active, created_at, password_hash, department, allowed_tabs, roles_json
    FROM users
    WHERE username = ?
  `);

  const insertSession = db.prepare(`
    INSERT INTO auth_sessions (token, user_id, expires_at)
    VALUES (?, ?, datetime('now', ?))
  `);

  const deleteSession = db.prepare(`DELETE FROM auth_sessions WHERE token = ?`);

  function getRequestRoles(req) {
    const fromMany = String(req.headers["x-user-roles"] || "")
      .split(",")
      .map((x) => String(x || "").trim().toLowerCase())
      .filter(Boolean);
    const fromSingle = String(req.headers["x-user-role"] || "")
      .split(",")
      .map((x) => String(x || "").trim().toLowerCase())
      .filter(Boolean);
    const merged = Array.from(new Set([...fromMany, ...fromSingle].filter((r) => VALID_ROLES.includes(r))));
    return merged.length ? merged : ["admin"];
  }

  function getRequestRole(req) {
    return pickPrimaryRole(getRequestRoles(req));
  }

  function getRequestUsername(req) {
    return String(req.headers["x-user-name"] || "").trim();
  }

  function requireAdmin(req, reply) {
    const roles = getRequestRoles(req);
    if (!roles.includes("admin") && !roles.includes("supervisor")) {
      reply.code(403).send({ error: "admin or supervisor role required" });
      return false;
    }
    return true;
  }

  function userPayload(row) {
    if (!row) return null;
    const roles = parseRoles(row.roles_json, row.role);
    return {
      id: Number(row.id),
      username: row.username,
      full_name: row.full_name,
      role: pickPrimaryRole(roles),
      roles,
      active: Number(row.active),
      department: row.department || null,
      allowed_tabs: parseAllowedTabs(row.allowed_tabs),
      has_password: Boolean(row.password_hash && String(row.password_hash).length > 0),
    };
  }

  // GET /api/auth/tabs — section keys + roles (for admin UI; not sensitive)
  app.get("/tabs", async () => ({
    ok: true,
    keys: [...ASSIGNABLE_TAB_KEYS],
    roles: [...VALID_ROLES],
  }));

  // POST /api/auth/login
  app.post("/login", async (req, reply) => {
    const body = req.body || {};
    const username = String(body.username || "").trim();
    const password = String(body.password || "");

    if (!username || !password) {
      return reply.code(400).send({ error: "username and password are required" });
    }

    const row = getByUsername.get(username);
    if (!row || Number(row.active) !== 1) {
      return reply.code(401).send({ error: "invalid username or password" });
    }
    if (!row.password_hash) {
      return reply.code(403).send({
        error: "password_not_set",
        message: "This account has no password yet. An admin must set one under User admin.",
      });
    }
    if (!verifyPassword(password, row.password_hash)) {
      return reply.code(401).send({ error: "invalid username or password" });
    }

    const token = crypto.randomBytes(32).toString("hex");
    const dayMod = `+${SESSION_DAYS} days`;
    insertSession.run(token, row.id, dayMod);

    return {
      ok: true,
      token,
      user: userPayload(row),
    };
  });

  // POST /api/auth/logout
  app.post("/logout", async (req) => {
    const auth = String(req.headers.authorization || "");
    const token = auth.startsWith("Bearer ") ? auth.slice(7).trim() : "";
    if (token) deleteSession.run(token);
    return { ok: true };
  });

  // GET /api/auth/me
  app.get("/me", async (req) => {
    const username = getRequestUsername(req);
    const roleFromHeader = getRequestRole(req);

    if (username) {
      const user = getByUsername.get(username);
      if (user && Number(user.active) === 1) {
        return {
          ok: true,
          user: userPayload(user),
          roles: VALID_ROLES,
          auth_required: process.env.IRONLOG_AUTH_REQUIRED === "1" || String(process.env.IRONLOG_AUTH_REQUIRED).toLowerCase() === "true",
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
        roles: [roleFromHeader],
        active: 1,
        department: null,
        allowed_tabs: null,
        has_password: false,
      },
      roles: VALID_ROLES,
      auth_required: process.env.IRONLOG_AUTH_REQUIRED === "1" || String(process.env.IRONLOG_AUTH_REQUIRED).toLowerCase() === "true",
    };
  });

  // GET /api/auth/users (admin only)
  app.get("/users", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const rows = db
      .prepare(
        `
      SELECT id, username, full_name, role, active, created_at, department, allowed_tabs, roles_json,
        CASE WHEN password_hash IS NOT NULL AND length(trim(password_hash)) > 0 THEN 1 ELSE 0 END AS has_password
      FROM users
      ORDER BY username ASC
    `
      )
      .all()
      .map((r) => ({
        id: Number(r.id),
        username: r.username,
        full_name: r.full_name,
        role: pickPrimaryRole(parseRoles(r.roles_json, r.role)),
        roles: parseRoles(r.roles_json, r.role),
        active: Number(r.active),
        created_at: r.created_at,
        department: r.department || null,
        allowed_tabs: parseAllowedTabs(r.allowed_tabs),
        has_password: Number(r.has_password) === 1,
      }));
    return { ok: true, rows };
  });

  // POST /api/auth/users (admin only)
  // Body: { username, full_name?, role?, roles?: string[], active?, password?, department?, allowed_tabs?: string[] }
  app.post("/users", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const body = req.body || {};
    const username = String(body.username || "").trim();
    const full_name = String(body.full_name || "").trim() || null;
    const department = String(body.department || "").trim() || null;
    const rolesInput = Array.isArray(body.roles)
      ? body.roles
      : (body.role ? [body.role] : []);
    const roles = parseRoles(JSON.stringify(rolesInput), "operator");
    const role = pickPrimaryRole(roles);
    const active = body.active === 0 || body.active === false ? 0 : 1;
    const password = String(body.password || "").trim();

    if (!username) return reply.code(400).send({ error: "username is required" });
    if (!roles.length) {
      return reply.code(400).send({ error: `roles must contain at least one of: ${VALID_ROLES.join(", ")}` });
    }

    let tabsStored = null;
    if (Object.prototype.hasOwnProperty.call(body, "allowed_tabs")) {
      const rawTabs = body.allowed_tabs;
      const parsed =
        rawTabs == null || (Array.isArray(rawTabs) && rawTabs.length === 0)
          ? null
          : parseAllowedTabs(Array.isArray(rawTabs) ? JSON.stringify(rawTabs) : String(rawTabs));
      tabsStored = parsed == null ? null : JSON.stringify(parsed);
    } else {
      const ex = getByUsername.get(username);
      tabsStored = ex ? ex.allowed_tabs : null;
    }

    db.prepare(
      `
      INSERT INTO users (username, full_name, role, active, department, allowed_tabs, roles_json)
      VALUES (?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(username) DO UPDATE SET
        full_name = COALESCE(excluded.full_name, users.full_name),
        role = excluded.role,
        active = excluded.active,
        department = excluded.department,
        allowed_tabs = excluded.allowed_tabs,
        roles_json = excluded.roles_json
    `
    ).run(username, full_name, role, active, department, tabsStored, JSON.stringify(roles));

    if (password) {
      const h = hashPassword(password);
      db.prepare(`UPDATE users SET password_hash = ? WHERE username = ?`).run(h, username);
    }

    const user = getByUsername.get(username);
    return { ok: true, user: userPayload(user) };
  });

  // POST /api/auth/change-password (logged-in user changes own password)
  app.post("/change-password", async (req, reply) => {
    const username = getRequestUsername(req);
    if (!username) return reply.code(401).send({ error: "not authenticated" });
    const body = req.body || {};
    const oldPassword = String(body.old_password || "");
    const newPassword = String(body.new_password || "").trim();
    if (!newPassword || newPassword.length < 6) {
      return reply.code(400).send({ error: "new_password must be at least 6 characters" });
    }
    const row = getByUsername.get(username);
    if (!row) return reply.code(404).send({ error: "user not found" });
    if (row.password_hash) {
      if (!verifyPassword(oldPassword, row.password_hash)) {
        return reply.code(403).send({ error: "current password is incorrect" });
      }
    }
    db.prepare(`UPDATE users SET password_hash = ? WHERE username = ?`).run(hashPassword(newPassword), username);
    return { ok: true };
  });
}
