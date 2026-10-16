import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import Fastify from "fastify";

// Two sites, one person limited to "north": nothing from "south" is listed, opened or changed.
const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-sites-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";

await import("../db/migrate.js");
(await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
const { db } = await import("../db/client.js");
const { ironlogAuthHook } = await import("../auth/hook.js");
const { makeSiteGuard } = await import("../utils/siteGuard.js");
const { ensureSiteFollowsMachine, ensureAssetSiteColumn } = await import("../utils/sites.js");
const realSetInterval = globalThis.setInterval;
globalThis.setInterval = (...args) => realSetInterval(...args).unref();
const app = Fastify({ logger: false });
app.addHook("onRequest", ironlogAuthHook);
app.decorateRequest("ironlogAllowedSites", null);
app.addHook("preHandler", makeSiteGuard(db));
await app.register((await import("../routes/auth.routes.js")).default, { prefix: "/api/auth" });
await app.register((await import("../routes/breakdowns.routes.js")).default, { prefix: "/api/breakdowns" });
await app.register((await import("../routes/hours.routes.js")).default, { prefix: "/api/hours" });
await app.register((await import("../routes/workorders.routes.js")).default, { prefix: "/api/workorders" });
await app.register((await import("../routes/maintenance.routes.js")).default, { prefix: "/api/maintenance" });
await app.ready();
globalThis.setInterval = realSetInterval;
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

ensureAssetSiteColumn(db);
ensureSiteFollowsMachine(db);
const asset = (code, site) => db.prepare(`INSERT INTO assets (asset_code, asset_name, category, active, site_code) VALUES (?, ?, 'Truck', 1, ?)`).run(code, code, site).lastInsertRowid;
const north = asset("N1", "north");
const south = asset("S1", "south");
const breakdown = (assetId) => db.prepare(`INSERT INTO breakdowns (asset_id, breakdown_date, status, description) VALUES (?, '2026-10-16', 'OPEN', 'Hose')`).run(assetId).lastInsertRowid;
const bdNorth = breakdown(north);
const bdSouth = breakdown(south);
const wo = (assetId) => db.prepare(`INSERT INTO work_orders (asset_id, source, status) VALUES (?, 'manual', 'open')`).run(assetId).lastInsertRowid;
const woNorth = wo(north);
const woSouth = wo(south);
for (const id of [north, south]) db.prepare(`INSERT INTO daily_hours (asset_id, work_date, hours_run, is_used) VALUES (?, '2026-10-16', 8, 1)`).run(id);
const plan = (assetId) => db.prepare(`INSERT INTO maintenance_plans (asset_id, service_name, interval_hours, last_service_hours, active) VALUES (?, '250h service', 5, 0, 1)`).run(assetId).lastInsertRowid;
plan(north);
const planSouth = plan(south);

const session = (username, role, allowed) => {
  db.prepare(`INSERT INTO users (username, role, roles_json, active, allowed_locations) VALUES (?, ?, ?, 1, ?)`).run(username, role, JSON.stringify([role]), allowed ? JSON.stringify(allowed) : null);
  const token = crypto.randomBytes(16).toString("hex");
  const id = db.prepare(`SELECT id FROM users WHERE username = ?`).get(username).id;
  db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, ?, datetime('now', '+1 day'))`).run(token, id);
  return token;
};
const nina = session("nina", "supervisor", ["north"]);
const boss = session("boss", "admin", null);
const call = (token, method, url, { site, body } = {}) => app.inject({
  method, url, payload: body,
  headers: { authorization: `Bearer ${token}`, ...(site ? { "x-site-code": site } : {}) },
});

test("new records take their machine's site", () => {
  assert.equal(db.prepare(`SELECT site_code FROM breakdowns WHERE id = ?`).get(bdSouth).site_code, "south");
  assert.equal(db.prepare(`SELECT site_code FROM work_orders WHERE id = ?`).get(woNorth).site_code, "north");
  assert.equal(db.prepare(`SELECT site_code FROM daily_hours WHERE asset_id = ?`).get(south).site_code, "south");
});

test("a person limited to north never sees, opens or changes south's records", async () => {
  const list = (await call(nina, "GET", "/api/breakdowns")).json();
  assert.deepEqual(list.map((b) => b.asset_code), ["N1"], "breakdowns list");
  assert.deepEqual((await call(nina, "GET", "/api/breakdowns/open-all")).json().rows.map((b) => b.asset_code), ["N1"]);
  assert.deepEqual((await call(nina, "GET", "/api/hours/2026-10-16")).json().map((r) => r.asset_code), ["N1"], "daily hours");

  assert.equal((await call(nina, "GET", `/api/breakdowns/${bdSouth}`)).statusCode, 404, "open a south breakdown by number");
  assert.equal((await call(nina, "GET", `/api/breakdowns/${bdNorth}`)).statusCode, 200);
  assert.equal((await call(nina, "GET", `/api/workorders/${woSouth}`)).statusCode, 404, "open a south work order");
  assert.equal((await call(nina, "POST", `/api/workorders/${woSouth}/status`, { body: { status: "completed" } })).statusCode, 404, "change a south work order");
  assert.equal(db.prepare(`SELECT status FROM work_orders WHERE id = ?`).get(woSouth).status, "open", "south work order untouched");
  assert.equal((await call(nina, "POST", "/api/hours", { body: { asset_code: "S1", work_date: "2026-10-17", hours_run: 5 } })).statusCode, 404, "enter hours for a south machine");
  assert.equal((await call(nina, "GET", "/api/breakdowns/open?asset_code=S1")).statusCode, 404);
  const due = (await call(nina, "GET", "/api/maintenance/due")).json();
  const dueCodes = [...new Set((due.due || []).map((r) => r.asset_code))];
  assert.deepEqual(dueCodes, ["N1"], "maintenance due list");
  assert.equal((await call(nina, "PUT", `/api/maintenance/plans/${planSouth}`, { body: { service_name: "changed", interval_hours: 500 } })).statusCode, 404, "edit a south maintenance plan");
  assert.equal(db.prepare(`SELECT service_name FROM maintenance_plans WHERE id = ?`).get(planSouth).service_name, "250h service");
  assert.equal((await call(nina, "POST", "/api/maintenance/plans", { body: { asset_id: south, service_name: "x", interval_hours: 100 } })).statusCode, 404, "add a plan to a south machine");
  assert.equal((await call(nina, "GET", "/api/breakdowns", { site: "south" })).statusCode, 403, "switching to south is refused");
  assert.deepEqual((await call(nina, "GET", "/api/breakdowns", { site: '["workshop"]' })).json().map((b) => b.asset_code), ["N1"], "an old garbled site in the browser falls back to her own site");
});

test("an administrator sees the site they choose", async () => {
  assert.deepEqual((await call(boss, "GET", "/api/breakdowns", { site: "south" })).json().map((b) => b.asset_code), ["S1"]);
  assert.deepEqual((await call(boss, "GET", "/api/breakdowns", { site: "north" })).json().map((b) => b.asset_code), ["N1"]);
  assert.equal((await call(boss, "GET", `/api/breakdowns/${bdSouth}`, { site: "north" })).statusCode, 200, "admins may open any site's record");
  assert.equal((await call(boss, "GET", "/api/hours/2026-10-16", { site: "x' OR 1=1 --" })).json().length, 0, "a strange site code matches nothing");
});
