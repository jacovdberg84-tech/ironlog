import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import Fastify from "fastify";

// Asset structure part 2: failure codes on breakdowns (component, problem, cause, action).
const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-failcodes-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";

await import("../db/migrate.js");
(await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
const { db } = await import("../db/client.js");
const { ironlogAuthHook } = await import("../auth/hook.js");
const { makeSiteGuard } = await import("../utils/siteGuard.js");
const { ensureAssetSiteColumn } = await import("../utils/sites.js");
const fc = await import("../utils/failureCodes.js");
const realSetInterval = globalThis.setInterval;
globalThis.setInterval = (...args) => realSetInterval(...args).unref();
const app = Fastify({ logger: false });
app.addHook("onRequest", ironlogAuthHook);
app.decorateRequest("ironlogAllowedSites", null);
app.addHook("preHandler", makeSiteGuard(db));
await app.register((await import("../routes/auth.routes.js")).default, { prefix: "/api/auth" });
await app.register((await import("../routes/breakdowns.routes.js")).default, { prefix: "/api/breakdowns" });
await app.ready();
globalThis.setInterval = realSetInterval;
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

ensureAssetSiteColumn(db);
const asset = (code, site) => db.prepare(`INSERT INTO assets (asset_code, asset_name, category, active, site_code) VALUES (?, ?, 'Excavator', 1, ?)`).run(code, code, site).lastInsertRowid;
const breakdown = (assetId, status, component, date = new Date().toISOString().slice(0, 10)) =>
  db.prepare(`INSERT INTO breakdowns (asset_id, breakdown_date, status, description, component, downtime_total_hours) VALUES (?, ?, ?, 'Hose burst on boom', ?, 6)`)
    .run(assetId, date, status, component).lastInsertRowid;
const north = asset("N1", "north");
const south = asset("S1", "south");
const bdNorth = breakdown(north, "CLOSED", "Hydraulics");
const bdNorthOpen = breakdown(north, "OPEN", "Motor");
const bdSouth = breakdown(south, "CLOSED", "Engine");
const bdOld = breakdown(north, "CLOSED", "Brakes", "2025-01-10");

const session = (username, role, allowed) => {
  db.prepare(`INSERT INTO users (username, role, roles_json, active, allowed_locations) VALUES (?, ?, ?, 1, ?)`).run(username, role, JSON.stringify([role]), allowed ? JSON.stringify(allowed) : null);
  const token = crypto.randomBytes(16).toString("hex");
  const id = db.prepare(`SELECT id FROM users WHERE username = ?`).get(username).id;
  db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, ?, datetime('now', '+1 day'))`).run(token, id);
  return (site) => ({ authorization: `Bearer ${token}`, ...(site ? { "x-site-code": site } : {}) });
};
const foreman = session("foreman", "plant_manager");
const tech = session("tech", "artisan");
const southForeman = session("sf", "plant_manager", ["south"]);

test("the lists start with the ISO 14224-style defaults, in English and Portuguese", () => {
  const lists = fc.listFailureCodes(db);
  assert.deepEqual(Object.keys(lists), ["component", "problem", "cause", "action"]);
  assert.ok(lists.component.length >= 20 && lists.cause.some((c) => c.code === "MAINT"));
  assert.ok(Object.values(lists).flat().every((o) => o.label_en && o.label_pt), "every entry has both languages");
  assert.equal(fc.componentFromText(db, "hydraulics"), "HYD");
  assert.equal(fc.componentFromText(db, "Motor"), "ENG", "the Portuguese name works too");
  assert.equal(fc.componentFromText(db, "boom hose"), null);
});

test("the technician suggests, the foreman confirms with all four codes", async () => {
  assert.equal(fc.suggestBreakdownFailureCodes(db, bdNorth, { component: "HYD", problem: "NOPE" }, "tech"), false, "an unknown code suggests nothing");
  assert.equal(fc.suggestBreakdownFailureCodes(db, bdNorth, { component: "HYD", problem: "BURST" }, "tech"), true);

  const put = (id, body, who = foreman("north")) => app.inject({ method: "PUT", url: `/api/breakdowns/${id}/failure-codes`, headers: who, payload: body });
  assert.equal((await put(bdNorth, { cause: "WEAR" }, tech("north"))).statusCode, 403, "technicians suggest in the portal; foremen confirm");
  assert.match((await put(bdNorth, { cause: "SOMETHING" })).json().error, /Cause: SOMETHING is not on the list/);

  const partial = (await put(bdNorth, { cause: "wear" })).json();
  assert.equal(partial.codes.confirmed_at, null, "three of four is not confirmed");
  const full = (await put(bdNorth, { action: "REPL" })).json();
  assert.deepEqual([full.codes.component, full.codes.problem, full.codes.cause, full.codes.action, full.codes.confirmed_by], ["HYD", "BURST", "WEAR", "REPL", "foreman"]);
  assert.ok(full.codes.confirmed_at);
  assert.equal(fc.suggestBreakdownFailureCodes(db, bdNorth, { component: "ENG" }, "tech"), false, "a confirmed breakdown keeps the foreman's codes");
  assert.equal((await put(bdNorth, { action: "" })).json().codes.confirmed_at, null, "emptying one takes the confirmation away");
  await put(bdNorth, { action: "REPL" });

  assert.equal((await put(bdSouth, { component: "ENG" }, foreman("north"))).statusCode, 200, "managers without site limits choose any");
  assert.equal((await put(bdNorth, { cause: "OPER" }, southForeman("south"))).statusCode, 404, "a south-only foreman cannot code a north breakdown");
});

test("waiting for codes: this site, last 90 days, machines back in service first; summary counts", async () => {
  const pending = (await app.inject({ method: "GET", url: "/api/breakdowns/failure-codes/pending", headers: foreman("north") })).json().rows;
  assert.deepEqual(pending.map((r) => r.id), [bdNorthOpen], "confirmed, other-site and older breakdowns are left out");
  assert.equal(pending[0].guess_component, "ENG", "\"Motor\" already names the component");

  const sum = (await app.inject({ method: "GET", url: "/api/breakdowns/failure-codes/summary", headers: foreman("north") })).json();
  assert.deepEqual(sum.totals, { breakdowns: 3, coded: 1, confirmed: 1 });
  assert.deepEqual(sum.by_component[0], { code: "HYD", breakdowns: 1, downtime_hours: 6, label: "Hydraulics" });
  assert.equal(sum.by_cause[0].label, "Normal wear and tear");
});

test("admins and foremen keep the lists; a used code is switched off, never lost", async () => {
  const edit = (rows, who = foreman()) => app.inject({ method: "PUT", url: "/api/breakdowns/failure-codes/cause", headers: who, payload: { rows } });
  const lists = fc.listFailureCodes(db, { all: true });
  assert.equal((await edit(lists.cause, tech())).statusCode, 403);
  assert.match((await edit([{ code: "bad code!", label_en: "x" }])).json().error, /code must be 1-16 letters/);
  assert.match((await edit([{ code: "A", label_en: "x" }, { code: "A", label_en: "y" }])).json().error, /listed twice/);

  const rows = lists.cause.filter((c) => c.code !== "WEAR").concat([{ code: "HEAT", label_en: "Heat stress", label_pt: "Stress térmico" }]);
  const saved = await edit(rows);
  assert.equal(saved.statusCode, 200);
  const after = fc.listFailureCodes(db, { all: true }).cause;
  assert.equal(after.find((c) => c.code === "WEAR").active, false, "left out = switched off");
  assert.equal(after.at(-1).code, "HEAT");
  assert.ok(!fc.listFailureCodes(db).cause.some((c) => c.code === "WEAR"), "not offered any more");
  const north = fc.breakdownFailureCodes(db, bdNorth);
  assert.equal(north.cause, "WEAR", "breakdowns already coded keep it");
  const resave = await app.inject({ method: "PUT", url: `/api/breakdowns/${bdNorth}/failure-codes`, headers: foreman("north"), payload: { cause: "WEAR", action: "REPL" } });
  assert.equal(resave.statusCode, 200, "the breakdown's own switched-off code can be saved again");
  assert.match((await app.inject({ method: "PUT", url: `/api/breakdowns/${bdNorthOpen}/failure-codes`, headers: foreman("north"), payload: { cause: "WEAR" } })).json().error, /not on the list/);
});
