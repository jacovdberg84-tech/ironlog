import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import http from "node:http";
import crypto from "node:crypto";
import Database from "better-sqlite3";
import Fastify from "fastify";

// The webhook helpers open the database (for the secret store), so the test database is set first.
const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-webhooks-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";
const wh = await import("../utils/webhooks.js");

// Connecting to other systems, part D: webhooks.
test("addresses must be public https; signatures are HMAC-SHA256 over time and body", () => {
  assert.equal(wh.checkWebhookUrl("https://erp.example.com/ironlog").ok, true);
  for (const bad of ["http://erp.example.com", "https://localhost/x", "https://10.0.0.5/x", "https://192.168.1.2", "https://[::1]/x", "https://user:pw@erp.example.com", "erp.example.com"]) {
    assert.equal(wh.checkWebhookUrl(bad, {}).ok, false, bad);
  }
  assert.equal(wh.checkWebhookUrl("http://127.0.0.1:9000/x", { IRONLOG_WEBHOOK_ALLOW_PRIVATE: "1" }).ok, true, "on-site systems when switched on");
  assert.deepEqual(["8.8.8.8", "172.16.0.1", "172.32.0.1", "169.254.169.254", "::ffff:10.1.1.1", "fd00::1", "2a00:1450::1"].map(wh.isPrivateAddress), [false, true, false, true, true, true, false]);
  const sig = wh.signWebhook("whsec_x", '{"a":1}', 1700000000);
  assert.equal(sig, `t=1700000000,v1=${crypto.createHmac("sha256", "whsec_x").update('1700000000.{"a":1}').digest("hex")}`);
});

test("the scan sends only what happens after subscribing, per event and site, and stock warns once until restocked", () => {
  const d = new Database(":memory:");
  d.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT, category TEXT, site_code TEXT);
    CREATE TABLE breakdowns (id INTEGER PRIMARY KEY, asset_id INTEGER, breakdown_date TEXT, status TEXT, description TEXT, critical INTEGER, downtime_total_hours REAL, site_code TEXT);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, status TEXT, labor_hours REAL, site_code TEXT);
    CREATE TABLE parts (id INTEGER PRIMARY KEY, part_code TEXT, part_name TEXT, min_stock REAL, unit_cost REAL);
    CREATE TABLE stock_movements (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL);
    INSERT INTO assets VALUES (1, 'A302AM', 'Bell B30D', 'ADT', 'main'), (2, 'S1', 'South truck', 'ADT', 'south');
    INSERT INTO breakdowns VALUES (1, 1, '2026-10-01', 'CLOSED', 'Old one', 0, 5, NULL);
    INSERT INTO parts VALUES (1, '210105', 'LEVEL INDICATOR', 2, 19.5);
    INSERT INTO stock_movements (part_id, quantity) VALUES (1, 5);
  `);
  assert.equal(wh.scanWebhookEvents(d), 0, "history is remembered, not sent");
  const hook = wh.createWebhook(d, { name: "ERP", url: "https://erp.example.com/h", events: ["breakdown.created", "breakdown.closed", "stock.below_minimum"], sites: ["main"] });
  assert.equal(hook.ok, true);
  wh.createWebhook(d, { name: "All sites, work orders", url: "https://bi.example.com/h", events: ["workorder.completed"] });
  assert.equal(wh.scanWebhookEvents(d), 0);

  d.exec(`INSERT INTO breakdowns VALUES (2, 1, '2026-10-17', 'OPEN', 'Hydraulic hose burst', 1, 0, NULL), (3, 2, '2026-10-17', 'OPEN', 'South', 0, 0, NULL);
          INSERT INTO work_orders VALUES (1, 2, 'breakdown', 'completed', 3.5, NULL);`);
  assert.equal(wh.scanWebhookEvents(d), 2, "main breakdown to ERP; south work order to the all-sites hook; south breakdown to nobody");
  const sent = d.prepare(`SELECT event, record_key, payload_json FROM webhook_deliveries ORDER BY id`).all();
  assert.deepEqual(sent.map((s) => [s.event, s.record_key]), [["breakdown.created", "2"], ["workorder.completed", "1"]]);
  const body = JSON.parse(sent[0].payload_json);
  assert.deepEqual([body.event, body.site_code, body.data.asset_code, body.data.description, body.data.critical], ["breakdown.created", "main", "A302AM", "Hydraulic hose burst", true]);
  assert.equal(wh.scanWebhookEvents(d), 0, "each event once");

  d.exec(`UPDATE breakdowns SET status = 'CLOSED', downtime_total_hours = 6 WHERE id = 2`);
  assert.equal(wh.scanWebhookEvents(d), 1);
  assert.equal(JSON.parse(d.prepare(`SELECT payload_json FROM webhook_deliveries WHERE event = 'breakdown.closed'`).get().payload_json).data.downtime_hours, 6);

  d.exec(`INSERT INTO stock_movements (part_id, quantity) VALUES (1, -4)`);
  assert.equal(wh.scanWebhookEvents(d), 1, "1 left, minimum 2");
  assert.equal(wh.scanWebhookEvents(d), 0, "not again while still low");
  d.exec(`INSERT INTO stock_movements (part_id, quantity) VALUES (1, 3)`);
  assert.equal(wh.scanWebhookEvents(d), 0);
  d.exec(`INSERT INTO stock_movements (part_id, quantity) VALUES (1, -3)`);
  assert.equal(wh.scanWebhookEvents(d), 1, "low again after a restock");

  d.exec(`INSERT INTO breakdowns VALUES (4, 1, '2026-10-17', 'OPEN', 'Not licensed', 0, 0, NULL)`);
  assert.equal(wh.scanWebhookEvents(d, { moduleOn: () => false }), 0, "nothing goes out without the licence");
  assert.equal(wh.updateWebhook(d, hook.hook.id, { active: false }).hook.active, false);
  d.exec(`INSERT INTO breakdowns VALUES (5, 1, '2026-10-17', 'OPEN', 'Paused', 0, 0, NULL)`);
  assert.equal(wh.scanWebhookEvents(d), 0, "a paused webhook gets nothing");
});

test("administrators add a webhook, see its secret once, test it, and failed messages retry", async (t) => {
  await import("../db/migrate.js");
  (await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
  const { db } = await import("../db/client.js");
  const { ironlogAuthHook } = await import("../auth/hook.js");
  const realSetInterval = globalThis.setInterval;
  globalThis.setInterval = (...args) => realSetInterval(...args).unref();
  const app = Fastify({ logger: false });
  app.addHook("onRequest", ironlogAuthHook);
  app.decorateRequest("ironlogAllowedSites", null);
  app.decorateRequest("ironlogApiKey", null);
  await app.register((await import("../routes/auth.routes.js")).default, { prefix: "/api/auth" });
  const webhooks = await import("../routes/webhooks.routes.js");
  await app.register(webhooks.default, { prefix: "/api/webhooks" });
  await app.ready();
  globalThis.setInterval = realSetInterval;

  // The other system: records what arrives; answers 500 while `failing`.
  const got = [];
  let failing = false;
  const server = http.createServer((req, res) => {
    let body = "";
    req.on("data", (c) => { body += c; });
    req.on("end", () => { got.push({ headers: req.headers, body }); res.statusCode = failing ? 500 : 200; res.end("ok"); });
  });
  await new Promise((r) => server.listen(0, "127.0.0.1", r));
  const url = `http://127.0.0.1:${server.address().port}/ironlog`;
  process.env.IRONLOG_WEBHOOK_ALLOW_PRIVATE = "1";
  t.after(async () => {
    delete process.env.IRONLOG_WEBHOOK_ALLOW_PRIVATE;
    server.close();
    await app.close();
    try { db.close(); } catch {}
    rmSync(tempDir, { recursive: true, force: true });
  });

  const session = (username, role) => {
    db.prepare(`INSERT INTO users (username, role, roles_json, active) VALUES (?, ?, ?, 1)`).run(username, role, JSON.stringify([role]));
    const token = crypto.randomBytes(16).toString("hex");
    db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, (SELECT id FROM users WHERE username = ?), datetime('now', '+1 day'))`).run(token, username);
    return { authorization: `Bearer ${token}` };
  };
  const boss = session("boss", "admin");
  const foreman = session("foreman", "plant_manager");

  assert.equal((await app.inject({ method: "POST", url: "/api/webhooks", headers: foreman, payload: { name: "x", url, events: ["breakdown.created"] } })).statusCode, 403);
  assert.match((await app.inject({ method: "POST", url: "/api/webhooks", headers: boss, payload: { name: "ERP", url, events: ["breakdown.created"], sites: ["moon"] } })).json().error, /Not a site: moon/);
  const made = await app.inject({ method: "POST", url: "/api/webhooks", headers: boss, payload: { name: "ERP", url, events: ["breakdown.created", "workorder.completed"] } });
  assert.equal(made.statusCode, 200, made.body);
  const { secret, webhook } = made.json();
  assert.match(secret, /^whsec_/);
  const listed = (await app.inject({ method: "GET", url: "/api/webhooks", headers: boss })).json();
  assert.equal(JSON.stringify(listed).includes(secret), false, "the secret is shown only once");
  assert.equal(JSON.stringify(db.prepare(`SELECT * FROM webhooks`).all()).includes(secret), false, "and stored encrypted");
  assert.equal(listed.events.length, 5);

  const ping = (await app.inject({ method: "POST", url: `/api/webhooks/${webhook.id}/test`, headers: boss })).json();
  assert.deepEqual([ping.ok, ping.status_code], [true, 200]);
  const first = got.at(-1);
  assert.equal(first.headers["x-ironlog-event"], "ping");
  const [, ts, v1] = first.headers["x-ironlog-signature"].match(/^t=(\d+),v1=([0-9a-f]{64})$/);
  assert.equal(v1, crypto.createHmac("sha256", secret).update(`${ts}.${first.body}`).digest("hex"), "the other system can check the signature");

  // A real event: history first, then a new breakdown.
  db.prepare(`INSERT INTO assets (asset_code, asset_name, category, active) VALUES ('A302AM', 'Bell B30D', 'ADT', 1)`).run();
  await webhooks.runWebhookPass();
  db.prepare(`INSERT INTO breakdowns (asset_id, breakdown_date, status, description) VALUES ((SELECT id FROM assets WHERE asset_code = 'A302AM'), '2026-10-17', 'OPEN', 'Hose burst')`).run();
  failing = true;
  await webhooks.runWebhookPass();
  const d = (await app.inject({ method: "GET", url: `/api/webhooks/${webhook.id}/deliveries`, headers: boss })).json().deliveries[0];
  assert.deepEqual([d.event, d.status, d.attempts, d.last_status_code], ["breakdown.created", "pending", 1, 500], "kept to try again");
  assert.match(d.last_error, /answered 500/);
  failing = false;
  assert.equal((await app.inject({ method: "POST", url: `/api/webhooks/${webhook.id}/deliveries/${d.id}/retry`, headers: boss })).statusCode, 200);
  const again = (await app.inject({ method: "GET", url: `/api/webhooks/${webhook.id}/deliveries`, headers: boss })).json().deliveries[0];
  assert.deepEqual([again.status, again.attempts], ["delivered", 2]);
  assert.equal(JSON.parse(got.at(-1).body).data.description, "Hose burst");

  const rotated = (await app.inject({ method: "POST", url: `/api/webhooks/${webhook.id}/secret`, headers: boss })).json().secret;
  assert.notEqual(rotated, secret);
  await app.inject({ method: "POST", url: `/api/webhooks/${webhook.id}/test`, headers: boss });
  const last = got.at(-1);
  const [, ts2, sig2] = last.headers["x-ironlog-signature"].match(/^t=(\d+),v1=(.+)$/);
  assert.equal(sig2, crypto.createHmac("sha256", rotated).update(`${ts2}.${last.body}`).digest("hex"), "signed with the new secret");

  assert.equal((await app.inject({ method: "DELETE", url: `/api/webhooks/${webhook.id}`, headers: boss })).statusCode, 200);
  const actions = db.prepare(`SELECT action FROM audit_logs WHERE module = 'webhooks' ORDER BY id`).all().map((r) => r.action);
  assert.deepEqual(actions, ["webhook.create", "webhook.test", "webhook.retry", "webhook.secret", "webhook.test", "webhook.delete"]);
});
