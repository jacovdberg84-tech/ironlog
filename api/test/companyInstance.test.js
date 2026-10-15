import test from "node:test";
import assert from "node:assert/strict";
import crypto from "node:crypto";
import Database from "better-sqlite3";
import { dropInheritedSettings } from "../loadEnv.js";
import { createFirstAdmin } from "../utils/firstAdmin.js";

test("a company instance drops IronLog settings it inherited, so AML keys never leak in", () => {
  const env = {
    IRONLOG_ENV_FILE: "/data/ironlog-instances/pilot/instance.env",
    OPENAI_API_KEY: "sk-aml", IRONLOG_GITHUB_TOKEN: "ghp", JUNE_OUTLOOK_CLIENT_SECRET: "x", DB_PATH: "/data/ironlog/api/db/ironlog.db",
    PORT: "3001", SMTP_PASSWORD: "p", CARTRACK_PASSWORD: "c", GH_TOKEN: "g",
    PATH: "/usr/bin", HOME: "/root", NODE_ENV: "production", pm_id: "3",
  };
  const removed = dropInheritedSettings(env);
  assert.deepEqual(Object.keys(env).sort(), ["HOME", "IRONLOG_ENV_FILE", "NODE_ENV", "PATH", "pm_id"]);
  assert.equal(removed.length, 8);
  const aml = { OPENAI_API_KEY: "sk-aml", PORT: "3001" };
  assert.deepEqual(dropInheritedSettings(aml), [], "AML's own server (no instance file) keeps everything");
  assert.equal(aml.OPENAI_API_KEY, "sk-aml");
});

const usersDb = () => {
  const d = new Database(":memory:");
  d.exec(`CREATE TABLE users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL UNIQUE, full_name TEXT, role TEXT NOT NULL DEFAULT 'operator',
    active INTEGER NOT NULL DEFAULT 1, password_hash TEXT, roles_json TEXT, setup_code_hash TEXT, setup_code_expires_at TEXT);
    INSERT INTO users (username, full_name, role) VALUES ('admin', 'System Admin', 'admin');`);
  return d;
};
const codeMatches = (code, stored) => {
  const [, salt, hash] = stored.split("$");
  return crypto.scryptSync(code, Buffer.from(salt, "base64"), 64).equals(Buffer.from(hash, "base64"));
};

test("a new instance gets one first administrator with a one-time setup code", () => {
  const d = usersDb();
  assert.match(createFirstAdmin(d, { username: "bad name!" }).error, /Username must be/);
  assert.match(createFirstAdmin(new Database(":memory:"), { username: "x1" }).error, /has not started yet/);

  const first = createFirstAdmin(d, { username: "pilot.admin", full_name: "Pilot Admin" });
  assert.equal(first.ok, true, "the built-in admin placeholder has no password, so it does not count");
  assert.match(first.setup_code, /^[0-9A-F]{8}$/);
  const row = d.prepare(`SELECT * FROM users WHERE username = 'pilot.admin'`).get();
  assert.equal(row.role, "admin");
  assert.equal(row.roles_json, '["admin"]');
  assert.equal(row.password_hash, null);
  assert.ok(codeMatches(first.setup_code, row.setup_code_hash), "the sign-in page can check the code");

  const again = createFirstAdmin(d, { username: "pilot.admin" });
  assert.equal(again.ok, true, "a re-run before the code is used gives a fresh code");
  assert.notEqual(again.setup_code, first.setup_code);
  assert.equal(d.prepare(`SELECT full_name FROM users WHERE username = 'pilot.admin'`).get().full_name, "Pilot Admin");

  d.prepare(`UPDATE users SET password_hash = 'scrypt$x$y', setup_code_hash = NULL WHERE username = 'pilot.admin'`).run();
  assert.match(createFirstAdmin(d, { username: "someone.else" }).error, /already has an administrator/);
  assert.equal(d.prepare(`SELECT COUNT(*) AS n FROM users`).get().n, 2, "nothing added once an admin can sign in");
});
