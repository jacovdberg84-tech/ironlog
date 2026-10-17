import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import Database from "better-sqlite3";
import Fastify from "fastify";
import * as keys from "../utils/apiKeys.js";

// Connecting to other systems, part A: read-only API keys for other systems.
test("a key is shown once, stored only as a hash, and checked by scope, method and end date", () => {
  const d = new Database(":memory:");
  const bad = keys.createApiKey(d, { name: "", scopes: ["everything"] });
  assert.equal(bad.ok, false);
  assert.equal(bad.errors.length, 2);
  assert.match(keys.createApiKey(d, { name: "x", scopes: ["assets:read"], expires_on: "2020-01-01" }, { today: "2026-10-17" }).errors[0], /today or later/);

  const made = keys.createApiKey(d, { name: "Sage", scopes: ["assets:read", "assets:read"], sites: "North" }, { by: "boss", checkSites: (l) => ({ ok: true, sites: l }) });
  assert.equal(made.ok, true);
  assert.match(made.key, /^ilk_[A-Za-z0-9_-]{40}$/);
  assert.deepEqual([made.row.scopes, made.row.sites, made.row.status, made.row.key_prefix], [["assets:read"], ["north"], "active", made.key.slice(0, 12)]);
  const stored = d.prepare(`SELECT * FROM api_keys`).get();
  assert.equal(JSON.stringify(stored).includes(made.key), false, "the key itself is never stored");
  assert.equal(JSON.stringify(keys.listApiKeys(d)).includes(stored.key_hash), false, "nor its hash shown");

  const k = keys.resolveApiKey(d, made.key);
  assert.equal(k.name, "Sage");
  assert.equal(keys.resolveApiKey(d, `${made.key}x`), null);
  assert.equal(keys.resolveApiKey(d, "ilk_short"), null);
  assert.equal(keys.keyAllows(k, "GET", "/api/assets/fleet-summary"), true);
  assert.equal(keys.keyAllows(k, "GET", "/api/assets/E1/components"), true);
  assert.equal(keys.keyAllows(k, "POST", "/api/assets"), false, "read-only");
  assert.equal(keys.keyAllows(k, "GET", "/api/assets/E1/history"), false, "only listed addresses");
  assert.equal(keys.keyAllows(k, "GET", "/api/breakdowns"), false, "only its scopes");
  assert.equal(keys.keyAllows(k, "GET", "/api/auth/users"), false);

  d.prepare(`UPDATE api_keys SET expires_on = '2026-10-16'`).run();
  assert.equal(keys.resolveApiKey(d, made.key, "2026-10-17"), null, "past its end date");
  d.prepare(`UPDATE api_keys SET expires_on = NULL`).run();
  assert.equal(keys.revokeApiKey(d, k.id, "boss"), true);
  assert.equal(keys.revokeApiKey(d, k.id, "boss"), false, "already cancelled");
  assert.equal(keys.resolveApiKey(d, made.key), null);
  assert.equal(keys.listApiKeys(d)[0].status, "cancelled");

  assert.equal(keys.underRateLimit(99, 0, 2), true);
  assert.equal(keys.underRateLimit(99, 1000, 2), true);
  assert.equal(keys.underRateLimit(99, 2000, 2), false, "third call in the minute");
  assert.equal(keys.underRateLimit(99, 61000, 2), true, "a new minute");

  const doc = keys.openApiDocument("https://pilot.ironlogafrica.com");
  assert.equal(doc.openapi, "3.0.3");
  assert.equal(doc.paths["/api/assets/{asset_code}/components"].get.parameters[0].name, "asset_code");
});

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-apikeys-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";

test("another system reads with its key: its scopes, its sites, read-only, until cancelled", async (t) => {
  await import("../db/migrate.js");
  (await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
  const { db } = await import("../db/client.js");
  const { ironlogAuthHook } = await import("../auth/hook.js");
  const { licenceHook } = await import("../routes/licence.routes.js");
  const { makeSiteGuard } = await import("../utils/siteGuard.js");
  const realSetInterval = globalThis.setInterval;
  globalThis.setInterval = (...args) => realSetInterval(...args).unref();
  const app = Fastify({ logger: false });
  app.addHook("onRequest", ironlogAuthHook);
  app.addHook("onRequest", licenceHook);
  app.decorateRequest("ironlogAllowedSites", null);
  app.decorateRequest("ironlogApiKey", null);
  app.addHook("preHandler", makeSiteGuard(db));
  await app.register((await import("../routes/auth.routes.js")).default, { prefix: "/api/auth" });
  await app.register((await import("../routes/assets.routes.js")).default, { prefix: "/api/assets" });
  await app.register((await import("../routes/breakdowns.routes.js")).default, { prefix: "/api/breakdowns" });
  await app.register((await import("../routes/apiKeys.routes.js")).default, { prefix: "/api/api-keys" });
  await app.ready();
  globalThis.setInterval = realSetInterval;
  t.after(async () => { await app.close(); try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

  for (const [code, site] of [["N1", "north"], ["N2", "north"], ["S1", "south"]]) {
    db.prepare(`INSERT INTO assets (asset_code, asset_name, category, active, site_code) VALUES (?, ?, 'Truck', 1, ?)`).run(code, code, site);
  }
  const session = (username, role) => {
    db.prepare(`INSERT INTO users (username, role, roles_json, active) VALUES (?, ?, ?, 1)`).run(username, role, JSON.stringify([role]));
    const token = crypto.randomBytes(16).toString("hex");
    db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, (SELECT id FROM users WHERE username = ?), datetime('now', '+1 day'))`).run(token, username);
    return { authorization: `Bearer ${token}` };
  };
  const boss = session("boss", "admin");
  const foreman = session("foreman", "plant_manager");

  assert.equal((await app.inject({ method: "POST", url: "/api/api-keys", headers: foreman, payload: { name: "x", scopes: ["assets:read"] } })).statusCode, 403, "administrators only");
  const bad = await app.inject({ method: "POST", url: "/api/api-keys", headers: boss, payload: { name: "Sage", scopes: ["assets:read"], sites: ["moon"] } });
  assert.match(bad.json().error, /Not a site: moon/);
  const created = await app.inject({ method: "POST", url: "/api/api-keys", headers: boss, payload: { name: "Sage", scopes: ["assets:read"], sites: ["north"] } });
  assert.equal(created.statusCode, 200, created.body);
  const key = created.json().key;
  const asKey = (extra = {}) => ({ authorization: `Bearer ${key}`, ...extra });

  const fleet = await app.inject({ method: "GET", url: "/api/assets/fleet-summary", headers: asKey() });
  assert.equal(fleet.statusCode, 200, fleet.body);
  assert.deepEqual(fleet.json().cards.map((c) => c.asset_code).sort(), ["N1", "N2"], "only its site");
  assert.equal((await app.inject({ method: "GET", url: "/api/assets/fleet-summary", headers: asKey({ "x-site-code": "south" }) })).statusCode, 403, "not another site");
  assert.equal((await app.inject({ method: "GET", url: "/api/assets/fleet-summary", headers: { "x-api-key": key } })).statusCode, 200, "X-API-Key works too");

  assert.equal((await app.inject({ method: "GET", url: "/api/breakdowns", headers: asKey() })).statusCode, 403, "not its scope");
  assert.equal((await app.inject({ method: "PATCH", url: "/api/assets/N1", headers: asKey({ "x-user-role": "admin" }), payload: { asset_name: "hacked" } })).statusCode, 403, "read-only, whatever role it claims");
  assert.equal(db.prepare(`SELECT asset_name FROM assets WHERE asset_code = 'N1'`).get().asset_name, "N1");
  assert.equal((await app.inject({ method: "GET", url: "/api/api-keys", headers: asKey() })).statusCode, 403, "a key cannot manage keys");
  assert.equal((await app.inject({ method: "GET", url: "/api/auth/users", headers: asKey() })).statusCode, 403);

  const listed = (await app.inject({ method: "GET", url: "/api/api-keys", headers: boss })).json().keys[0];
  assert.equal(listed.use_count >= 2, true);
  assert.ok(listed.last_used_at);
  assert.equal(JSON.stringify(listed).includes(key), false);

  assert.equal((await app.inject({ method: "POST", url: `/api/api-keys/${listed.id}/cancel`, headers: boss })).statusCode, 200);
  assert.equal((await app.inject({ method: "GET", url: "/api/assets/fleet-summary", headers: asKey() })).statusCode, 401, "cancelled keys stop at once");
  assert.equal((await app.inject({ method: "GET", url: "/api/assets/fleet-summary", headers: asKey() })).json().error, "invalid, cancelled or expired API key");
  const audit = db.prepare(`SELECT action FROM audit_logs WHERE module = 'api_keys' ORDER BY id`).all().map((r) => r.action);
  assert.deepEqual(audit, ["api_key.create", "api_key.cancel"]);
  assert.equal((await app.inject({ method: "GET", url: "/api/api-keys/catalogue", headers: foreman })).json().scopes.length, 6, "the documentation is for anyone signed in");
});
