// IRONLOG/api/utils/signInSecurity.js — wrong-password lockout and signed-in sessions.
//
// Lockout: LOCK_AFTER wrong passwords, PINs or setup codes for one account within
// WINDOW_MINUTES lock that account for LOCK_MINUTES (even the right password is refused
// meanwhile); IP_LIMIT wrong passwords from one address in the window pause that address,
// which stops someone trying one password against many accounts. PIN mistakes do not count
// towards the address limit: a site's people often share one internet address, and the
// shared workshop terminals would otherwise pause themselves. Unknown usernames count too,
// so the answer never reveals which accounts exist. A good sign-in clears the account's
// count; an administrator can unlock early.
//
// Sessions: each sign-in (password, PIN or Microsoft) is a row in auth_sessions with how
// it signed in, the address and device, and when it was last used (updated at most every
// SEEN_EVERY_MINUTES). Administrators see who is signed in and can sign one session, or a
// person everywhere, out at once.
import crypto from "node:crypto";

export const LOCK_AFTER = 5;
export const WINDOW_MINUTES = 15;
export const LOCK_MINUTES = 15;
export const IP_LIMIT = 50;
const SEEN_EVERY_MINUTES = 5;

const key = (username) => String(username || "").trim().toLowerCase();
const iso = (d) => d.toISOString();

export function ensureSignInTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS login_failures (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username_key TEXT NOT NULL,
      ip TEXT,
      kind TEXT NOT NULL,
      failed_at TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_login_failures_user ON login_failures(username_key, failed_at);
    CREATE INDEX IF NOT EXISTS idx_login_failures_ip ON login_failures(ip, failed_at);
    CREATE TABLE IF NOT EXISTS login_lockouts (
      username_key TEXT PRIMARY KEY,
      locked_until TEXT NOT NULL,
      locked_at TEXT NOT NULL,
      failures INTEGER NOT NULL
    );
  `);
  const have = new Set(db.prepare(`PRAGMA table_info(auth_sessions)`).all().map((c) => c.name));
  if (!have.size) return;
  for (const [col, type] of [["kind", "TEXT"], ["ip", "TEXT"], ["user_agent", "TEXT"], ["last_seen_at", "TEXT"]]) {
    if (!have.has(col)) db.prepare(`ALTER TABLE auth_sessions ADD COLUMN ${col} ${type}`).run();
  }
}

const minutesLeft = (until, now) => Math.max(1, Math.ceil((Date.parse(until) - now.getTime()) / 60000));

/**
 * Whether a sign-in may be tried now: { ok } or { ok:false, minutes, reason: "account"|"address", message }.
 */
export function signInAllowed(db, username, ip, now = new Date(), { kind = "password" } = {}) {
  ensureSignInTables(db);
  const lock = db.prepare(`SELECT locked_until FROM login_lockouts WHERE username_key = ?`).get(key(username));
  if (lock && Date.parse(lock.locked_until) > now.getTime()) {
    const m = minutesLeft(lock.locked_until, now);
    return { ok: false, reason: "account", minutes: m, message: `Too many wrong attempts. This account is locked for ${m} more minute${m === 1 ? "" : "s"}. Try again later, or ask an administrator to unlock it.` };
  }
  if (ip && kind !== "pin") {
    const since = iso(new Date(now.getTime() - WINDOW_MINUTES * 60000));
    const n = db.prepare(`SELECT COUNT(*) AS c, MIN(failed_at) AS first FROM login_failures WHERE ip = ? AND kind != 'pin' AND failed_at > ?`).get(ip, since);
    if (Number(n.c) >= IP_LIMIT) {
      const m = minutesLeft(iso(new Date(Date.parse(n.first) + WINDOW_MINUTES * 60000)), now);
      return { ok: false, reason: "address", minutes: m, message: `Too many wrong attempts from this device. Wait ${m} minute${m === 1 ? "" : "s"} and try again.` };
    }
  }
  return { ok: true };
}

/** Notes a wrong password/PIN/code. Returns { locked, failures } (locked = this failure locked the account). */
export function recordFailure(db, username, ip, kind, now = new Date()) {
  ensureSignInTables(db);
  const k = key(username);
  db.prepare(`INSERT INTO login_failures (username_key, ip, kind, failed_at) VALUES (?, ?, ?, ?)`).run(k, ip || null, kind, iso(now));
  const since = iso(new Date(now.getTime() - WINDOW_MINUTES * 60000));
  const failures = Number(db.prepare(`SELECT COUNT(*) AS c FROM login_failures WHERE username_key = ? AND failed_at > ?`).get(k, since).c);
  let locked = false;
  if (failures >= LOCK_AFTER) {
    db.prepare(`
      INSERT INTO login_lockouts (username_key, locked_until, locked_at, failures) VALUES (?, ?, ?, ?)
      ON CONFLICT(username_key) DO UPDATE SET locked_until = excluded.locked_until, locked_at = excluded.locked_at, failures = excluded.failures
    `).run(k, iso(new Date(now.getTime() + LOCK_MINUTES * 60000)), iso(now), failures);
    db.prepare(`DELETE FROM login_failures WHERE username_key = ?`).run(k);
    locked = true;
  }
  // Keep a day of failures.
  db.prepare(`DELETE FROM login_failures WHERE failed_at < ?`).run(iso(new Date(now.getTime() - 86400000)));
  return { locked, failures };
}

/** A good sign-in clears the account's count (an active lock is checked before this is reached). */
export function recordSuccess(db, username) {
  ensureSignInTables(db);
  db.prepare(`DELETE FROM login_failures WHERE username_key = ?`).run(key(username));
  db.prepare(`DELETE FROM login_lockouts WHERE username_key = ?`).run(key(username));
}

export function unlockAccount(db, username) {
  ensureSignInTables(db);
  const had = db.prepare(`DELETE FROM login_lockouts WHERE username_key = ?`).run(key(username)).changes > 0;
  db.prepare(`DELETE FROM login_failures WHERE username_key = ?`).run(key(username));
  return had;
}

export function lockedAccounts(db, now = new Date()) {
  ensureSignInTables(db);
  return db.prepare(`SELECT username_key AS username, locked_until, locked_at, failures FROM login_lockouts WHERE locked_until > ? ORDER BY locked_at DESC`).all(iso(now));
}

/** The request's address and device, for the session list. */
export function clientOf(req) {
  return {
    ip: String(req?.ip || req?.socket?.remoteAddress || "").slice(0, 64) || null,
    user_agent: String(req?.headers?.["user-agent"] || "").slice(0, 300) || null,
  };
}

/** Starts a session; returns its token. lifetime is an SQLite modifier like "+7 days". */
export function createSession(db, userId, { kind, lifetime, req } = {}) {
  ensureSignInTables(db);
  const token = crypto.randomBytes(32).toString("hex");
  const c = clientOf(req);
  db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at, kind, ip, user_agent, last_seen_at) VALUES (?, ?, datetime('now', ?), ?, ?, ?, datetime('now'))`)
    .run(token, userId, lifetime, kind || null, c.ip, c.user_agent);
  return token;
}

/** Marks a session as used (at most every few minutes, so busy screens do not write on every call). */
export function touchSession(db, token) {
  try {
    db.prepare(`UPDATE auth_sessions SET last_seen_at = datetime('now') WHERE token = ? AND (last_seen_at IS NULL OR last_seen_at < datetime('now', ?))`)
      .run(token, `-${SEEN_EVERY_MINUTES} minutes`);
  } catch { /* columns not added yet on a fresh start */ }
}

/** A short device name from the browser's user agent. */
export function deviceName(ua) {
  const s = String(ua || "");
  if (!s) return "Unknown device";
  const os = /Android/i.test(s) ? "Android" : /iPhone|iPad/i.test(s) ? "iPhone/iPad" : /Windows/i.test(s) ? "Windows" : /Mac OS/i.test(s) ? "Mac" : /Linux/i.test(s) ? "Linux" : "";
  const browser = /Edg\//.test(s) ? "Edge" : /Chrome\//.test(s) ? "Chrome" : /Firefox\//.test(s) ? "Firefox" : /Safari\//.test(s) ? "Safari" : "";
  return [browser, os].filter(Boolean).join(" on ") || s.slice(0, 40);
}

/** Active sessions, newest use first. The token itself is never returned. */
export function listSessions(db, { currentToken = "" } = {}) {
  ensureSignInTables(db);
  return db.prepare(`
    SELECT s.id, s.token, s.kind, s.ip, s.user_agent, s.created_at, s.last_seen_at, s.expires_at, u.username, u.full_name
    FROM auth_sessions s JOIN users u ON u.id = s.user_id
    WHERE datetime(s.expires_at) > datetime('now')
    ORDER BY COALESCE(s.last_seen_at, s.created_at) DESC
  `).all().map(({ token, ...r }) => ({ ...r, device: deviceName(r.user_agent), current: Boolean(currentToken) && token === currentToken }));
}

export function endSession(db, id) {
  return db.prepare(`DELETE FROM auth_sessions WHERE id = ?`).run(Number(id)).changes > 0;
}

/** Signs a person out everywhere (except keepToken, when given). Returns how many sessions ended. */
export function endUserSessions(db, username, { keepToken = "" } = {}) {
  return db.prepare(`DELETE FROM auth_sessions WHERE user_id IN (SELECT id FROM users WHERE LOWER(username) = ?) AND token != ?`)
    .run(key(username), keepToken || "").changes;
}
