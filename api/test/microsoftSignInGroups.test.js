import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import Database from "better-sqlite3";
import Fastify from "fastify";

// Microsoft sign-in stage 3: groups → roles and sites, "open to the whole company",
// "Microsoft sign-in only" users, and the warning before the client secret runs out.
const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-msgroups-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";
process.env.JUNE_CONNECTOR_ENCRYPTION_SECRET = "test-connector-secret-for-microsoft-sign-in-0123456789";

const ms = await import("../utils/microsoftSignIn.js");
const { VALID_ROLES } = await import("../routes/auth.routes.js");

const TENANT = "11111111-2222-3333-4444-555555555555";
const CLIENT = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";
const SECRET = "Abc8Q~very-secret-value-from-entra-1234";
const G_WORKSHOP = "0a0a0a0a-1111-2222-3333-444444444444";
const G_STORES = "0b0b0b0b-1111-2222-3333-444444444444";
const G_NORTH = "0c0c0c0c-1111-2222-3333-444444444444";
const OTHER_GROUP = "0d0d0d0d-1111-2222-3333-444444444444";

function freshDb() {
  const d = new Database(":memory:");
  d.exec(`CREATE TABLE users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, full_name TEXT, role TEXT, active INTEGER DEFAULT 1, roles_json TEXT,
            allowed_locations TEXT, password_hash TEXT, setup_code_hash TEXT, setup_code_expires_at TEXT, pin_hash TEXT);
          INSERT INTO users (username, role, roles_json, password_hash) VALUES ('boss', 'admin', '["admin"]', 'scrypt$x$y'), ('ana', 'plant_manager', '["plant_manager"]', 'scrypt$x$y');`);
  ms.ensureMicrosoftSignInTables(d);
  return d;
}
const checkSites = (list) => (list.every((s) => ["main", "north"].includes(s)) ? { ok: true, sites: list } : { ok: false, error: `Not a site: ${list.join(", ")}` });
const save = (d, groups) => ms.saveGroupRoles(d, groups, { validRoles: VALID_ROLES, checkSites, by: "boss" });
const claims = (extra = {}) => ({ tid: TENANT, oid: "oid-new", email: "Maria.Santos@aml.co.mz", name: "Maria Santos", groups: [], ...extra });
const errCode = (fn) => { try { fn(); return null; } catch (e) { return e.code; } };

test("linked groups are checked before they are saved", () => {
  const d = freshDb();
  const bad = save(d, [{ group_id: "Workshop team", roles: ["artisan"] }, { group_id: G_STORES, roles: ["boss"] }, { group_id: G_NORTH, roles: ["artisan"], sites: ["south"] }]);
  assert.equal(bad.ok, false);
  assert.match(bad.errors[0], /Object ID must look like/);
  assert.match(bad.errors[1], /not a role: boss/);
  assert.match(bad.errors[2], /Not a site: south/);
  assert.equal(ms.listGroupRoles(d).length, 0, "nothing saved when anything is wrong");
  const ok = save(d, [{ group_id: G_WORKSHOP.toUpperCase(), label: "Workshop", roles: ["artisan", "Artisan"] }, { group_id: G_NORTH, label: "North pit", roles: ["supervisor"], sites: "north" }]);
  assert.equal(ok.ok, true);
  assert.deepEqual(ok.groups.find((g) => g.group_id === G_WORKSHOP), { group_id: G_WORKSHOP, label: "Workshop", roles: ["artisan"], sites: [] });
  assert.deepEqual(ok.groups.find((g) => g.group_id === G_NORTH).sites, ["north"]);
});

test("a person's groups give the roles of every linked group, and sites only when all their groups are limited", () => {
  const d = freshDb();
  save(d, [{ group_id: G_WORKSHOP, roles: ["artisan"] }, { group_id: G_STORES, roles: ["storeman"], sites: ["main"] }, { group_id: G_NORTH, roles: ["supervisor"], sites: ["north"] }]);
  assert.deepEqual(ms.groupAccess(d, claims({ groups: [G_STORES, G_NORTH, OTHER_GROUP] })), {
    overage: false, groups: ms.listGroupRoles(d).filter((g) => [G_STORES, G_NORTH].includes(g.group_id)), roles: ["storeman", "supervisor"], sites: ["main", "north"],
  });
  assert.equal(ms.groupAccess(d, claims({ groups: [G_WORKSHOP, G_NORTH] })).sites, null, "a group without sites means all sites");
  assert.equal(ms.groupAccess(d, claims({ _claim_names: { groups: "src1" } })).overage, true, "too many groups for Microsoft to list");
});

test("open to the whole company: people in a linked group get a Microsoft-only user; everyone else is refused", () => {
  const d = freshDb();
  ms.saveMsSettings(d, { tenant_id: TENANT, client_id: CLIENT, client_secret: SECRET, enabled: true });
  assert.match(ms.saveMsSettings(d, { open_to_company: true }).errors[0], /Link at least one Microsoft group/);
  save(d, [{ group_id: G_WORKSHOP, label: "Workshop", roles: ["artisan"] }, { group_id: G_NORTH, roles: ["supervisor"], sites: ["north"] }]);
  const v = { validRoles: VALID_ROLES };

  assert.equal(errCode(() => ms.matchUser(d, claims({ groups: [G_WORKSHOP] }), v)), "not_invited", "closed: an invite is needed");
  assert.equal(ms.saveMsSettings(d, { open_to_company: true }).settings.open_to_company, true);
  assert.equal(errCode(() => ms.matchUser(d, claims({ groups: [OTHER_GROUP] }), v)), "no_group", "in the company but in no linked group");
  assert.equal(errCode(() => ms.matchUser(d, claims({ hasgroups: true }), v)), "too_many_groups");

  const u = ms.matchUser(d, claims({ groups: [G_NORTH] }), v);
  assert.equal(u.created, true);
  assert.deepEqual([u.username, u.full_name, u.role, JSON.parse(u.roles_json), JSON.parse(u.allowed_locations)], ["maria.santos", "Maria Santos", "supervisor", ["supervisor"], ["north"]]);
  assert.deepEqual([u.ms_only, u.ms_roles_from_groups, u.password_hash, u.ms_email], [1, 1, null, "maria.santos@aml.co.mz"]);
  assert.equal(ms.matchUser(d, claims({ groups: [G_NORTH] }), v).created, false, "the second sign-in finds the same user");

  d.prepare(`INSERT INTO users (username, role) VALUES ('joao.m', 'artisan')`).run();
  assert.equal(ms.matchUser(d, claims({ oid: "oid-2", email: "Joao.M@aml.co.mz", groups: [G_WORKSHOP] }), v).username, "joao.m2", "a taken username gets a number");

  save(d, []);
  assert.equal(ms.getMsSettings(d).open_to_company, false, "removing every linked group closes open sign-in");
});

test("roles from groups: reset at every sign-in, refused when no longer in a linked group; hand-set users keep their roles", () => {
  const d = freshDb();
  ms.saveMsSettings(d, { tenant_id: TENANT, client_id: CLIENT, client_secret: SECRET, enabled: true });
  save(d, [{ group_id: G_WORKSHOP, roles: ["artisan"] }, { group_id: G_STORES, roles: ["storeman"], sites: ["main"] }]);
  const v = { validRoles: VALID_ROLES };
  ms.setUserMicrosoftEmail(d, "ana", "ana@aml.co.mz");
  const ana = claims({ oid: "oid-ana", email: "ana@aml.co.mz", groups: [G_STORES] });

  assert.equal(ms.matchUser(d, ana, v).roles_changed, null, "invited by hand: groups are not used");
  assert.deepEqual(JSON.parse(d.prepare(`SELECT roles_json FROM users WHERE username = 'ana'`).get().roles_json), ["plant_manager"]);

  assert.equal(ms.setUserMicrosoftAccess(d, "ana", { roles_from_groups: true }).ok, true);
  const synced = ms.matchUser(d, ana, v);
  assert.deepEqual(synced.roles_changed, { from: { roles: ["plant_manager"], sites: null }, to: { roles: ["storeman"], sites: ["main"] } });
  assert.deepEqual([synced.role, synced.allowed_locations], ["storeman", '["main"]']);
  assert.equal(ms.matchUser(d, ana, v).roles_changed, null, "nothing changed, nothing recorded");
  assert.deepEqual(JSON.parse(ms.matchUser(d, { ...ana, groups: [G_STORES, G_WORKSHOP] }, v).roles_json).sort(), ["artisan", "storeman"]);
  assert.equal(errCode(() => ms.matchUser(d, { ...ana, groups: [OTHER_GROUP] }, v)), "no_group", "taken out of every linked group");
});

test("Microsoft-only needs a Microsoft account and sign-in switched on, removes the password, and never locks out the last admin", () => {
  const d = freshDb();
  assert.match(ms.setUserMicrosoftAccess(d, "ana", { ms_only: true }).error, /Microsoft account \(work email\) first/);
  ms.setUserMicrosoftEmail(d, "ana", "ana@aml.co.mz");
  assert.match(ms.setUserMicrosoftAccess(d, "ana", { ms_only: true }).error, /Switch Microsoft sign-in on/);
  ms.saveMsSettings(d, { tenant_id: TENANT, client_id: CLIENT, client_secret: SECRET, enabled: true });
  d.prepare(`UPDATE users SET pin_hash = 'scrypt$p$q' WHERE username = 'ana'`).run();
  assert.equal(ms.setUserMicrosoftAccess(d, "ana", { ms_only: true }).ok, true);
  const row = d.prepare(`SELECT ms_only, password_hash, pin_hash FROM users WHERE username = 'ana'`).get();
  assert.deepEqual(row, { ms_only: 1, password_hash: null, pin_hash: "scrypt$p$q" }, "password gone, PIN kept for the terminals");

  ms.setUserMicrosoftEmail(d, "boss", "boss@aml.co.mz");
  assert.match(ms.setUserMicrosoftAccess(d, "boss", { ms_only: true }).error, /At least one administrator must keep a password/);
  assert.notEqual(d.prepare(`SELECT password_hash FROM users WHERE username = 'boss'`).get().password_hash, null);
  d.prepare(`INSERT INTO users (username, role, roles_json, password_hash) VALUES ('it', 'admin', '["admin"]', 'scrypt$x$y')`).run();
  assert.equal(ms.setUserMicrosoftAccess(d, "boss", { ms_only: true }).ok, true, "fine once another admin has a password");
});

test("administrators are warned 30 days before the client secret runs out", () => {
  const d = freshDb();
  ms.saveMsSettings(d, { tenant_id: TENANT, client_id: CLIENT, client_secret: SECRET, secret_expires_on: "2026-11-01", enabled: false });
  assert.equal(ms.secretWarning(d, "2026-10-16"), null, "switched off: nothing to warn about");
  ms.saveMsSettings(d, { enabled: true });
  assert.equal(ms.secretWarning(d, "2026-09-01"), null, "more than 30 days left");
  assert.deepEqual([ms.secretWarning(d, "2026-10-16").days_left, /runs out in 16 days/.test(ms.secretWarning(d, "2026-10-16").message)], [16, true]);
  assert.match(ms.secretWarning(d, "2026-11-03").message, /ran out on 2026-11-01/);
});

test("sign-in routes: a Microsoft-only user has no password way in, but the PIN still works; a refused setting changes nothing", async (t) => {
  await import("../db/migrate.js");
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

  const hash = (p) => { const salt = crypto.randomBytes(16); return `scrypt$${salt.toString("base64")}$${crypto.scryptSync(p, salt, 64).toString("base64")}`; };
  db.prepare(`UPDATE users SET password_hash = ?, roles_json = '["admin"]' WHERE username = 'admin'`).run(hash("admin-pass-1"));
  const login = (username, password) => app.inject({ method: "POST", url: "/api/auth/login", payload: { username, password } });
  const token = (await login("admin", "admin-pass-1")).json().token;
  const asAdmin = { authorization: `Bearer ${token}` };
  const saveUser = (payload) => app.inject({ method: "POST", url: "/api/auth/users", headers: asAdmin, payload });

  const created = await saveUser({ username: "tecla", roles: ["artisan"], password: "tecla-pass-1", pin: "4321" });
  assert.equal(created.statusCode, 200);
  const refused = await saveUser({ username: "tecla", full_name: "Changed", roles: ["artisan"], microsoft_email: "tecla@aml.co.mz", microsoft_only: true });
  assert.equal(refused.statusCode, 400);
  assert.match(refused.json().error, /Switch Microsoft sign-in on/);
  const row = db.prepare(`SELECT full_name, ms_email, password_hash FROM users WHERE username = 'tecla'`).get();
  assert.deepEqual([row.full_name, row.ms_email, Boolean(row.password_hash)], [null, null, true], "the refused save left the user as it was");

  ms.saveMsSettings(db, { tenant_id: TENANT, client_id: CLIENT, client_secret: SECRET, enabled: true });
  assert.match((await saveUser({ username: "tecla", roles: ["artisan"], password: "x123456", microsoft_email: "tecla@aml.co.mz", microsoft_only: true })).json().error, /has no password/);
  const ok = await saveUser({ username: "tecla", roles: ["artisan"], microsoft_email: "tecla@aml.co.mz", microsoft_only: true });
  assert.equal(ok.statusCode, 200);
  assert.deepEqual([ok.json().user.microsoft_only, ok.json().user.has_password, ok.json().setup_code], [true, false, null], "no password, no setup code");

  const pw = await login("tecla", "tecla-pass-1");
  assert.deepEqual([pw.statusCode, pw.json().error], [403, "microsoft_only"]);
  assert.equal((await app.inject({ method: "POST", url: "/api/auth/setup-password", payload: { username: "tecla", setup_code: "x", new_password: "abcdef" } })).statusCode, 403);
  assert.equal((await app.inject({ method: "POST", url: "/api/auth/pin/login", payload: { username: "tecla", pin: "4321" } })).statusCode, 200, "PIN still works on the terminals");

  const last = await saveUser({ username: "admin", roles: ["admin"], microsoft_email: "admin@aml.co.mz", microsoft_only: true });
  assert.match(last.json().error, /At least one administrator must keep a password/);
  assert.equal((await login("admin", "admin-pass-1")).statusCode, 200, "the only admin with a password can still sign in");
});
