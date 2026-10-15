// IRONLOG/api/utils/firstAdmin.js — the first administrator of a brand-new company instance.
//
// A new instance starts with no users, and sign-in is required, so nobody could
// get in. deploy/ironlog-instance.sh calls scripts/first-admin.js once the
// instance has started: it adds one admin with a one-time setup code (valid 48
// hours) which that person uses on the sign-in page to choose their password.
// It never touches an instance where an administrator has set a password; the
// built-in 'admin' placeholder has none and cannot sign in, so it does not count.
// Running it again before the code is used issues a fresh code.
import crypto from "node:crypto";

// Same format as auth.routes.js, so the sign-in page accepts the code.
function hashPassword(plain) {
  const salt = crypto.randomBytes(16);
  const hash = crypto.scryptSync(String(plain), salt, 64);
  return `scrypt$${salt.toString("base64")}$${hash.toString("base64")}`;
}

function hasUsableAdmin(db) {
  return db.prepare(`SELECT role, roles_json FROM users WHERE active = 1 AND password_hash IS NOT NULL`).all().some((u) => {
    if (u.role === "admin") return true;
    try {
      return JSON.parse(u.roles_json || "[]").includes("admin");
    } catch {
      return false;
    }
  });
}

/** { ok, setup_code, expires_at } or { ok: false, error }. */
export function createFirstAdmin(db, { username, full_name = "" } = {}) {
  const name = String(username || "").trim();
  if (!/^[A-Za-z0-9._-]{2,60}$/.test(name)) return { ok: false, error: "Username must be 2-60 letters, numbers, dots, dashes or underscores." };
  const table = db.prepare(`SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'users'`).get();
  if (!table) return { ok: false, error: "This instance has not started yet (no users table). Start it first." };
  if (hasUsableAdmin(db)) return { ok: false, error: "This instance already has an administrator. Nothing was changed." };
  if (db.prepare(`SELECT 1 FROM users WHERE username = ? AND password_hash IS NOT NULL`).get(name)) {
    return { ok: false, error: `${name} already has a password. Choose another username.` };
  }
  const code = crypto.randomBytes(4).toString("hex").toUpperCase();
  db.prepare(`
    INSERT INTO users (username, full_name, role, active, roles_json, setup_code_hash, setup_code_expires_at)
    VALUES (?, ?, 'admin', 1, '["admin"]', ?, datetime('now', '+2 days'))
    ON CONFLICT(username) DO UPDATE SET
      full_name = COALESCE(excluded.full_name, users.full_name), role = 'admin', active = 1, roles_json = '["admin"]',
      setup_code_hash = excluded.setup_code_hash, setup_code_expires_at = excluded.setup_code_expires_at
  `).run(name, String(full_name || "").trim() || null, hashPassword(code));
  const row = db.prepare(`SELECT setup_code_expires_at FROM users WHERE username = ?`).get(name);
  return { ok: true, username: name, setup_code: code, expires_at: row.setup_code_expires_at };
}
