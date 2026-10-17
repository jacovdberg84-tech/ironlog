import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import Fastify from "fastify";

// Asset structure part 3: area, a machine's major components and their life.
const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-components-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";

await import("../db/migrate.js");
(await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
const { db } = await import("../db/client.js");
const { ironlogAuthHook } = await import("../auth/hook.js");
const { makeSiteGuard } = await import("../utils/siteGuard.js");
const comp = await import("../utils/assetComponents.js");
const realSetInterval = globalThis.setInterval;
globalThis.setInterval = (...args) => realSetInterval(...args).unref();
const app = Fastify({ logger: false });
app.addHook("onRequest", ironlogAuthHook);
app.decorateRequest("ironlogAllowedSites", null);
app.addHook("preHandler", makeSiteGuard(db));
await app.register((await import("../routes/auth.routes.js")).default, { prefix: "/api/auth" });
await app.register((await import("../routes/assets.routes.js")).default, { prefix: "/api/assets" });
await app.ready();
globalThis.setInterval = realSetInterval;
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

const asset = (code, site, meter) => {
  const id = db.prepare(`INSERT INTO assets (asset_code, asset_name, category, active, site_code) VALUES (?, ?, 'Excavator', 1, ?)`).run(code, code, site).lastInsertRowid;
  db.prepare(`INSERT INTO daily_hours (asset_id, work_date, opening_hours, closing_hours, hours_run, is_used) VALUES (?, '2026-10-16', ?, ?, 10, 1)`).run(id, meter - 10, meter);
  return id;
};
asset("E502AM", "north", 12000);
asset("S1", "south", 5000);

const session = (username, role, allowed) => {
  db.prepare(`INSERT INTO users (username, role, roles_json, active, allowed_locations) VALUES (?, ?, ?, 1, ?)`).run(username, role, JSON.stringify([role]), allowed ? JSON.stringify(allowed) : null);
  const token = crypto.randomBytes(16).toString("hex");
  const id = db.prepare(`SELECT id FROM users WHERE username = ?`).get(username).id;
  db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, ?, datetime('now', '+1 day'))`).run(token, id);
  return (site) => ({ authorization: `Bearer ${token}`, ...(site ? { "x-site-code": site } : {}) });
};
const foreman = session("foreman", "plant_manager");
const tech = session("tech", "artisan");
const southOnly = session("sf", "plant_manager", ["south"]);
const call = (method, url, headers, payload) => app.inject({ method, url, headers, payload });

test("hours since fitted and life used come from the machine's meter", () => {
  const v = comp.componentView({ component: "ENG", fitted_hours: 2000, expected_life_hours: 12000, status: "fitted" }, 12500, { ENG: "Engine" });
  assert.deepEqual([v.label, v.hours_since_fitted, v.life_used_pct, v.life_left_hours, v.due], ["Engine", 10500, 88, 1500, true]);
  const removed = comp.componentView({ component: "ENG", fitted_hours: 2000, removed_hours: 9000, status: "removed" }, 12500);
  assert.equal(removed.hours_since_fitted, 7000, "a changed component stops counting at its change");
  assert.equal(comp.componentView({ component: "ENG", fitted_hours: null, status: "fitted" }, 100).hours_since_fitted, null);
  assert.match(comp.readComponentFields({ fitted_hours: "-1" }).errors[0], /number of hours/);
  assert.match(comp.readComponentFields({}, { componentCodes: ["ENG"], requireComponent: true }).errors[0], /Choose which component/);
});

test("foremen fit and change components; history stays; technicians only look", async () => {
  const fit = await call("POST", "/api/assets/E502AM/components", foreman("north"), { component: "eng", serial_number: "QSK-001", make: "Cummins", model: "QSK19", fitted_hours: 2000, fitted_date: "2024-01-10", expected_life_hours: 12000 });
  assert.equal(fit.statusCode, 200, fit.body);
  assert.equal((await call("POST", "/api/assets/E502AM/components", foreman("north"), { component: "ROCKET" })).statusCode, 400);
  assert.equal((await call("POST", "/api/assets/E502AM/components", tech("north"), { component: "HYD" })).statusCode, 403);
  const trans = await call("POST", "/api/assets/E502AM/components", foreman("north"), { component: "TRANS", serial_number: "TR-9" });
  assert.equal(trans.statusCode, 200);

  let list = (await call("GET", "/api/assets/E502AM/components", tech("north"))).json();
  assert.equal(list.meter_hours, 12000);
  const eng = list.fitted.find((c) => c.component === "ENG");
  assert.deepEqual([eng.label, eng.hours_since_fitted, eng.life_used_pct, eng.due], ["Engine", 10000, 83, true]);
  const tr = list.fitted.find((c) => c.component === "TRANS");
  assert.deepEqual([tr.fitted_hours, tr.hours_since_fitted, tr.life_used_pct], [12000, 0, null], "no hours given: fitted at today's meter");

  const changeUrl = `/api/assets/E502AM/components/${eng.id}/change`;
  assert.equal((await call("POST", changeUrl, tech("north"), { reason: "failed" })).statusCode, 403);
  assert.match((await call("POST", changeUrl, foreman("north"), {})).json().error, /Say why it was changed/);
  assert.match((await call("POST", changeUrl, foreman("north"), { reason: "x", removed_hours: 1500 })).json().error, /below the hours when it was fitted/);
  assert.match((await call("POST", changeUrl, foreman("north"), { reason: "x", removed_date: "2023-01-01" })).json().error, /before it was fitted/);
  const changed = await call("POST", changeUrl, foreman("north"), { reason: "End of life", removed_hours: 11800, removed_date: "2026-10-01", serial_number: "QSK-002" });
  assert.equal(changed.statusCode, 200, changed.body);
  assert.equal((await call("POST", changeUrl, foreman("north"), { reason: "again" })).statusCode, 409);

  list = (await call("GET", "/api/assets/E502AM/components", foreman("north"))).json();
  const newEng = list.fitted.find((c) => c.component === "ENG");
  assert.deepEqual([newEng.serial_number, newEng.make, newEng.fitted_hours, newEng.hours_since_fitted, newEng.expected_life_hours], ["QSK-002", "Cummins", 11800, 200, 12000],
    "the new engine takes the old one's place, make and expected life");
  assert.deepEqual(list.removed.map((c) => [c.serial_number, c.removed_date, c.hours_since_fitted, c.removed_reason]), [["QSK-001", "2026-10-01", 9800, "End of life"]]);

  const fix = await call("PATCH", `/api/assets/E502AM/components/${newEng.id}`, foreman("north"), { position: "Main", expected_life_hours: "" });
  assert.equal(fix.statusCode, 200);
  assert.equal(db.prepare(`SELECT expected_life_hours FROM asset_components WHERE id = ?`).get(newEng.id).expected_life_hours, null);
});

test("other sites: guarded by address and left out of the end-of-life list", async () => {
  await call("POST", "/api/assets/S1/components", foreman("south"), { component: "UC", fitted_hours: 0, expected_life_hours: 5000 });
  assert.equal((await call("GET", "/api/assets/E502AM/components", southOnly("south"))).statusCode, 404, "a south-only foreman cannot open a north machine");
  assert.equal((await call("POST", "/api/assets/E502AM/components", southOnly("south"), { component: "BRK" })).statusCode, 404);

  db.prepare(`UPDATE asset_components SET expected_life_hours = 10000 WHERE component = 'TRANS'`).run();
  db.prepare(`UPDATE asset_components SET fitted_hours = 3000 WHERE component = 'TRANS'`).run();
  const north = (await call("GET", "/api/assets/component-life", foreman("north"))).json().rows;
  assert.deepEqual(north.map((r) => [r.asset_code, r.component, r.life_used_pct]), [["E502AM", "TRANS", 90]]);
  const south = (await call("GET", "/api/assets/component-life", foreman("south"))).json().rows;
  assert.deepEqual(south.map((r) => [r.asset_code, r.component, r.life_used_pct]), [["S1", "UC", 100]]);
});

test("area is a machine detail like make and model", async () => {
  const res = await call("PATCH", "/api/assets/E502AM", foreman("north"), { area: "Pit" });
  assert.equal(res.statusCode, 200);
  assert.equal(res.json().asset.area, "Pit");
  const card = (await call("GET", "/api/assets/fleet-summary", foreman("north"))).json().cards.find((c) => c.asset_code === "E502AM");
  assert.equal(card.area, "Pit");
  assert.equal((await call("PATCH", "/api/assets/E502AM", foreman("north"), { area: "x".repeat(41) })).statusCode, 400);
});
