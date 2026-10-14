import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-june-task-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
const Fastify = (await import("fastify")).default;
const juneRoutes = (await import("../routes/june.routes.js")).default;
const { db } = await import("../db/client.js");

const admin = { "x-user-role": "admin", "x-user-roles": "admin", "x-user-name": "jaco", "x-site-code": "main" };

test("June's task draft is shown for review, with the assignee matched from a spoken name", async (t) => {
  t.after(() => { try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY, username TEXT, full_name TEXT, active INTEGER DEFAULT 1);
    INSERT INTO users (username, full_name, active) VALUES
      ('jaco', 'Jaco van den Berg', 1), ('jose.m', 'Jose Mabunda', 1), ('ana', 'Ana Tecnica', 1),
      ('maria1', 'Maria Stores', 1), ('maria2', 'Maria Silva', 1), ('old', 'Old Timer', 0);
  `);
  const app = Fastify();
  await app.register(juneRoutes, { prefix: "/api/june" });
  t.after(() => app.close());
  const draft = async (args) => {
    const res = await app.inject({ method: "POST", url: "/api/june/gateway/execute", headers: admin, payload: { name: "june_draft_task", arguments: args } });
    assert.equal(res.statusCode, 200);
    return res.json().result;
  };

  const d = await draft({ title: "Check E502AM oil leak", description: "Oil low five days running.", priority: "high", due_date: "2026-10-15", assignee: "Jose" });
  assert.equal(d.state, "review_only");
  assert.equal(d.shown_on_screen, true);
  assert.equal(d.proposed_assignee, "jose.m", "a first name only one person has");
  assert.equal(d.priority, "high");
  assert.equal(d.due_date, "2026-10-15");
  assert.match(d.next_step, /Nothing has been created/);

  assert.equal((await draft({ title: "x", assignee: "Ana Tecnica" })).proposed_assignee, "ana", "full name");
  assert.equal((await draft({ title: "x", assignee: "@ana" })).proposed_assignee, "ana", "username");
  assert.equal((await draft({ title: "x" })).proposed_assignee, "jaco", "no name: the administrator");
  const unclear = await draft({ title: "x", assignee: "Maria" });
  assert.equal(unclear.proposed_assignee, null, "two Marias: the card asks instead of guessing");
  assert.equal(unclear.assignee_not_found, "Maria");
  assert.equal((await draft({ title: "x", assignee: "Old" })).proposed_assignee, null, "inactive users are not offered");
  assert.equal((await draft({ title: "   " })).state, "error");
  // Nothing was created: drafts never write tasks.
  const tasks = db.prepare(`SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'tasks'`).get();
  if (tasks) assert.equal(db.prepare(`SELECT COUNT(*) AS n FROM tasks`).get().n, 0);
});
