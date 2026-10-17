import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import Database from "better-sqlite3";
import Fastify from "fastify";
import * as sec from "../utils/signInSecurity.js";

// Security quick wins: wrong-password lockout and signed-in sessions.
test("five wrong attempts lock the account for 15 minutes; a good sign-in or an admin clears it", () => {
  const d = new Database(":memory:");
  const t0 = new Date("2026-10-17T08:00:00Z");
  for (let i = 1; i <= 4; i++) assert.equal(sec.recordFailure(d, "Jaco", "1.1.1.1", "password", new Date(t0.getTime() + i * 1000)).locked, false);
  assert.equal(sec.signInAllowed(d, "jaco", "1.1.1.1", t0).ok, true);
  assert.equal(sec.recordFailure(d, "jaco", "1.1.1.1", "password", new Date(t0.getTime() + 5000)).locked, true);
  const refused = sec.signInAllowed(d, "JACO", "9.9.9.9", new Date(t0.getTime() + 60000));
  assert.deepEqual([refused.ok, refused.reason, refused.minutes], [false, "account", 15], "locked from any address");
  assert.equal(sec.signInAllowed(d, "jaco", "9.9.9.9", new Date(t0.getTime() + 16 * 60000)).ok, true, "opens again after 15 minutes");
  assert.equal(sec.lockedAccounts(d, new Date(t0.getTime() + 60000)).length, 1);
  assert.equal(sec.unlockAccount(d, "jaco"), true);
  assert.equal(sec.signInAllowed(d, "jaco", "9.9.9.9", new Date(t0.getTime() + 60000)).ok, true);

  // Old mistakes fall out of the 15-minute window.
  for (let i = 0; i < 4; i++) sec.recordFailure(d, "ana", null, "password", t0);
  assert.equal(sec.recordFailure(d, "ana", null, "password", new Date(t0.getTime() + 16 * 60000)).locked, false);

  // One address trying many accounts is paused; PIN typos at a shared terminal are not counted for the address.
  for (let i = 0; i < sec.IP_LIMIT; i++) sec.recordFailure(d, `user${i}`, "6.6.6.6", "password", t0);
  assert.equal(sec.signInAllowed(d, "someone-else", "6.6.6.6", t0).reason, "address");
  for (let i = 0; i < 80; i++) sec.recordFailure(d, `tech${i % 20}`, "10.0.0.5", "pin", t0);
  assert.equal(sec.signInAllowed(d, "tech-new", "10.0.0.5", t0, { kind: "pin" }).ok, true);
  assert.equal(sec.signInAllowed(d, "tech-new", "10.0.0.5", t0).ok, true);
  assert.equal(sec.deviceName("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0 Safari/537.36"), "Chrome on Windows");
});

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-signin-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";

test("sign-in is refused while locked, admins see sessions and sign people out, a new password ends other sessions", async (t) => {
  await import("../db/migrate.js");
  (await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
  const { db } = await import("../db/client.js");
  const { ironlogAuthHook } = await import("../auth/hook.js");
  const realSetInterval = globalThis.setInterval;
  globalThis.setInterval = (...args) => realSetInterval(...args).unref();
  const app = Fastify({ logger: false });
  app.addHook("onRequest", ironlogAuthHook);
  app.decorateRequest("ironlogAllowedSites", null);
  await app.register((await import("../routes/auth.routes.js")).default, { prefix: "/api/auth" });
  await app.ready();
  globalThis.setInterval = realSetInterval;
  t.after(async () => { await app.close(); try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

  const make = (u, role) => db.prepare(`INSERT OR IGNORE INTO users (username, role, roles_json, active) VALUES (?, ?, ?, 1)`).run(u, role, JSON.stringify([role]));
  make("boss", "admin");
  make("sipho", "artisan");
  // Sipho sets a first password from a signed-in session (the app's own hashing).
  const session = (u) => {
    const token = crypto.randomBytes(16).toString("hex");
    db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, (SELECT id FROM users WHERE username = ?), datetime('now', '+1 day'))`).run(token, u);
    return { authorization: `Bearer ${token}` };
  };
  const bossH = session("boss");
  const sipho1 = session("sipho");
  assert.equal((await app.inject({ method: "POST", url: "/api/auth/change-password", headers: sipho1, payload: { new_password: "Correct-horse-1" } })).statusCode, 200);
  assert.equal((await app.inject({ method: "GET", url: "/api/auth/me", headers: sipho1 })).statusCode, 200, "the session that changed the password stays");

  const login = (password) => app.inject({ method: "POST", url: "/api/auth/login", payload: { username: "sipho", password } });
  const ok = await login("Correct-horse-1");
  assert.equal(ok.statusCode, 200, ok.body);
  const siphoToken = ok.json().token;
  for (let i = 0; i < 4; i++) assert.equal((await login("wrong")).statusCode, 401);
  assert.equal((await login("wrong")).statusCode, 401, "the fifth wrong password still says wrong password");
  const locked = await login("Correct-horse-1");
  assert.equal(locked.statusCode, 429, "even the right password waits while locked");
  assert.match(locked.json().message, /locked for 15 more minutes/);
  assert.equal(db.prepare(`SELECT COUNT(*) AS c FROM audit_logs WHERE action = 'login.locked'`).get().c, 1);

  const list = (await app.inject({ method: "GET", url: "/api/auth/sessions", headers: bossH })).json();
  assert.deepEqual(list.locked.map((l) => l.username), ["sipho"]);
  const mine = list.sessions.filter((s) => s.username === "sipho");
  assert.equal(mine.length, 2);
  assert.ok(mine.some((s) => s.kind === "password"), "how they signed in");
  assert.equal(JSON.stringify(list).includes(siphoToken), false, "tokens are never shown");
  assert.equal(list.sessions.find((s) => s.username === "boss").current, true);
  assert.equal((await app.inject({ method: "GET", url: "/api/auth/sessions", headers: sipho1 })).statusCode, 403, "administrators only");

  assert.equal((await app.inject({ method: "POST", url: "/api/auth/users/sipho/unlock", headers: bossH })).json().was_locked, true);
  assert.equal((await login("Correct-horse-1")).statusCode, 200);

  const out = (await app.inject({ method: "POST", url: "/api/auth/users/sipho/sign-out", headers: bossH })).json();
  assert.equal(out.ended, 3);
  assert.equal((await app.inject({ method: "GET", url: "/api/auth/me", headers: { authorization: `Bearer ${siphoToken}` } })).statusCode, 401, "signed out everywhere at once");

  const a = (await login("Correct-horse-1")).json().token;
  const b = (await login("Correct-horse-1")).json().token;
  const change = await app.inject({ method: "POST", url: "/api/auth/change-password", headers: { authorization: `Bearer ${a}` }, payload: { old_password: "Correct-horse-1", new_password: "Correct-horse-2" } });
  assert.equal(change.json().other_sessions_ended, 1);
  assert.equal((await app.inject({ method: "GET", url: "/api/auth/me", headers: { authorization: `Bearer ${b}` } })).statusCode, 401);
  assert.equal((await app.inject({ method: "GET", url: "/api/auth/me", headers: { authorization: `Bearer ${a}` } })).statusCode, 200);

  const sid = (await app.inject({ method: "GET", url: "/api/auth/sessions", headers: bossH })).json().sessions.find((s) => s.username === "sipho").id;
  assert.equal((await app.inject({ method: "DELETE", url: `/api/auth/sessions/${sid}`, headers: bossH })).statusCode, 200);
  assert.equal((await app.inject({ method: "GET", url: "/api/auth/me", headers: { authorization: `Bearer ${a}` } })).statusCode, 401);
  const e = (await login("Correct-horse-2")).json().token;
  assert.equal((await app.inject({ method: "POST", url: "/api/auth/logout-everywhere", headers: { authorization: `Bearer ${e}` } })).json().ended, 1);
});
