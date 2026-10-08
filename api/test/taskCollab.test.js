import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-collab-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { mentionedUsers } = await import("../utils/taskCollab.js");
const { default: tasksRoutes } = await import("../routes/tasks.routes.js");
const { default: authRoutes } = await import("../routes/auth.routes.js");
const app = Fastify({ logger: false });
await app.register(authRoutes, { prefix: "/api/auth" });
await app.register(tasksRoutes, { prefix: "/api" });
await app.ready();
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

const cols = new Set(db.prepare("PRAGMA table_info(users)").all().map((c) => c.name));
const addUser = db.prepare(`INSERT INTO users (username, full_name, role${cols.has("password_hash") ? ", password_hash" : ""}${cols.has("active") ? ", active" : ""}) VALUES (?, ?, ?${cols.has("password_hash") ? ", 'x'" : ""}${cols.has("active") ? ", 1" : ""})`);
addUser.run("jaco", "Jaco", "admin");
addUser.run("ana", "Ana Tecnica", "artisan");
addUser.run("maria", "Maria Stores", "storeman");

const as = (user, role = "admin") => ({ "x-user-name": user, "x-user-role": role, "x-user-roles": role, "x-site-code": "main" });
const call = async (user, method, url, payload, role) => {
  const res = await app.inject({ method, url: `/api${url}`, headers: as(user, role), payload });
  return { code: res.statusCode, body: res.json() };
};
const inbox = async (user) => (await call(user, "GET", "/collab/inbox")).body;

test("@mentions resolve to real users only", () => {
  assert.deepEqual(mentionedUsers(db, "Hi @Ana and @maria, also @nobody and email x@ana.com").sort(), ["ana", "maria"]);
});

test("assigned and @mentioned people get inbox items; comments are signed by the signed-in user", async (t) => {
  t.after(() => app.close());

  // Jaco creates a task for Ana and mentions Maria.
  const created = await call("jaco", "POST", "/tasks", { title: "Order hoses for E504AM", description: "@maria please check lead time", assigned_to: "ANA", watchers: "maria, nobody" });
  assert.equal(created.code, 200);
  const id = created.body.task.id;
  assert.equal(created.body.task.assigned_to, "ana", "assignee stored as the real username");
  assert.equal(created.body.task.created_by, "jaco");
  assert.deepEqual(created.body.watchers, ["maria"]);

  let a = await inbox("ana");
  assert.equal(a.unread, 1);
  assert.equal(a.items[0].kind, "assigned");
  assert.equal(a.items[0].actor, "jaco");
  let m = await inbox("maria");
  assert.equal(m.items[0].kind, "mention");
  assert.equal((await inbox("jaco")).unread, 0, "nobody is told about their own action");

  // Ana replies and mentions Jaco; "author" in the body is ignored.
  const c = await call("ana", "POST", `/tasks/${id}/comments`, { comment: "On it @jaco, supplier says Friday", author: "maria" }, "artisan");
  assert.equal(c.body.comment.author, "ana");
  const j = await inbox("jaco");
  assert.equal(j.unread, 1);
  assert.equal(j.items[0].kind, "mention");
  assert.match(j.items[0].snippet, /supplier says Friday/);

  // Views: mine / created / watching / involved.
  assert.deepEqual((await call("ana", "GET", "/tasks?view=mine")).body.tasks.map((x) => x.id), [id]);
  assert.deepEqual((await call("jaco", "GET", "/tasks?view=created")).body.tasks.map((x) => x.id), [id]);
  assert.deepEqual((await call("maria", "GET", "/tasks?view=watching")).body.tasks.map((x) => x.id), [id]);
  assert.equal((await call("maria", "GET", "/tasks?view=involved")).body.tasks[0].comments_count, 1);
  assert.deepEqual((await call("ana", "GET", "/tasks?q=hoses")).body.tasks.map((x) => x.id), [id]);

  // Reassigning tells the new assignee only; editing old text does not repeat mentions.
  await call("jaco", "PUT", `/tasks/${id}`, { assigned_to: "maria", description: "@maria please check lead time (urgent)" });
  m = await inbox("maria");
  assert.equal(m.items.filter((i) => i.kind === "assigned").length, 1);
  assert.equal(m.items.filter((i) => i.kind === "mention").length, 1, "no repeat mention");

  // Reading: by task and all.
  assert.equal((await call("maria", "POST", "/collab/inbox/read", { task_id: id })).body.unread, 0);
  a = await call("ana", "POST", "/collab/inbox/read", { all: true });
  assert.equal(a.body.unread, 0);

  // Follow / unfollow.
  assert.deepEqual((await call("ana", "POST", `/tasks/${id}/watch`, { watch: true }, "artisan")).body.watchers.sort(), ["ana", "maria"]);
  assert.deepEqual((await call("ana", "POST", `/tasks/${id}/watch`, { watch: false }, "artisan")).body.watchers, ["maria"]);

  // Deleting someone else's comment or task needs a supervisor.
  const jc = await call("jaco", "POST", `/tasks/${id}/comments`, { comment: "Thanks" });
  assert.equal((await call("ana", "DELETE", `/comments/${jc.body.comment.id}`, undefined, "artisan")).code, 403);
  assert.equal((await call("ana", "DELETE", `/tasks/${id}`, undefined, "artisan")).code, 403);
  assert.equal((await call("jaco", "DELETE", `/tasks/${id}`)).code, 200);
  assert.equal((await inbox("jaco")).items.length, 0, "inbox items go with the task");
});
