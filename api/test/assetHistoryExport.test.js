import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import Fastify from "fastify";
import ExcelJS from "exceljs";

// Every machine's history on one Excel sheet (Assets → All machine history).
const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-history-xlsx-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";

await import("../db/migrate.js");
(await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
const { db } = await import("../db/client.js");
const { ironlogAuthHook } = await import("../auth/hook.js");
const { makeSiteGuard } = await import("../utils/siteGuard.js");
const { historyRow } = await import("../utils/assetHistory.js");
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

const asset = (code, site, archived = 0) => db.prepare(`INSERT INTO assets (asset_code, asset_name, category, active, site_code, archived) VALUES (?, ?, 'Excavator', 1, ?, ?)`).run(code, `${code} name`, site, archived).lastInsertRowid;
const ex = asset("EX01", "north");
asset("EX02", "north");
asset("OLD1", "north", 1);
const south = asset("S1", "south");
db.prepare(`INSERT INTO breakdowns (asset_id, breakdown_date, status, description, critical) VALUES (?, '2026-11-02', 'OPEN', 'Boom hose burst', 1)`).run(ex);
db.prepare(`INSERT INTO work_orders (asset_id, source, status, opened_at) VALUES (?, 'service', 'open', '2026-11-03 08:00:00')`).run(ex);
db.prepare(`INSERT INTO fuel_logs (asset_id, log_date, liters) VALUES (?, '2026-11-04', 250)`).run(ex);
db.prepare(`INSERT INTO fuel_logs (asset_id, log_date, liters) VALUES (?, '2026-10-01', 100)`).run(ex);
db.prepare(`INSERT INTO oil_logs (asset_id, log_date, quantity) VALUES (?, '2026-11-05', 20)`).run(ex);
db.prepare(`INSERT INTO fuel_logs (asset_id, log_date, liters) VALUES (?, '2026-11-04', 999)`).run(south);

db.prepare(`INSERT INTO users (username, role, roles_json, active) VALUES ('tech', 'artisan', '["artisan"]', 1)`).run();
const token = crypto.randomBytes(16).toString("hex");
db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, (SELECT id FROM users WHERE username = 'tech'), datetime('now', '+1 day'))`).run(token);
const get = (url, site = "north") => app.inject({ method: "GET", url, headers: { authorization: `Bearer ${token}`, "x-site-code": site } });
const sheet = async (res) => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(res.rawPayload);
  assert.equal(wb.worksheets.length, 1, "one sheet");
  const ws = wb.worksheets[0];
  const head = ws.getRow(1).values.slice(1);
  const rows = [];
  ws.eachRow((r, n) => { if (n > 1) rows.push(Object.fromEntries(head.map((h, i) => [h, r.values[i + 1] ?? ""]))); });
  return { head, rows };
};

test("one sheet with every machine on the site and every kind of record", async () => {
  const res = await get("/api/assets/history.xlsx");
  assert.equal(res.statusCode, 200, res.body);
  assert.match(res.headers["content-disposition"], /IRONLOG_Machine_History_north\.xlsx/);
  const { head, rows } = await sheet(res);
  assert.deepEqual(head.slice(0, 6), ["Machine", "Machine name", "Category", "Site", "Date", "Type"]);
  const ex1 = rows.filter((r) => r.Machine === "EX01");
  assert.deepEqual(ex1.map((r) => r.Type), ["Oil / lube", "Fuel", "Work order", "Breakdown", "Fuel"], "newest first");
  assert.equal(ex1.find((r) => r.Type === "Breakdown").Description, "Boom hose burst");
  assert.equal(ex1.find((r) => r.Date === "2026-11-04").Litres, 250);
  assert.equal(rows.find((r) => r.Machine === "EX02").Type, "No records", "a machine with nothing still shows");
  assert.ok(!rows.some((r) => r.Machine === "S1"), "another site's machines stay out");
  assert.ok(!rows.some((r) => r.Machine === "OLD1"), "archived machines only when asked");
  const withArchived = await sheet(await get("/api/assets/history.xlsx?include_archived=1"));
  assert.ok(withArchived.rows.some((r) => r.Machine === "OLD1"));
});

test("a date range narrows the records; bad dates are refused", async () => {
  const res = await get("/api/assets/history.xlsx?start=2026-11-01&end=2026-11-30");
  assert.match(res.headers["content-disposition"], /_2026-11-01_to_2026-11-30\.xlsx/);
  const { rows } = await sheet(res);
  assert.equal(rows.filter((r) => r.Machine === "EX01").length, 4, "the October fuel log is left out");
  assert.equal((await get("/api/assets/history.xlsx?start=01/11/2026")).statusCode, 400);
  assert.equal((await get("/api/assets/history.xlsx?start=2026-12-01&end=2026-11-01")).statusCode, 400);
  assert.equal((await app.inject({ method: "GET", url: "/api/assets/history.xlsx" })).statusCode, 401, "sign-in needed");
});

test("the machine profile timeline is unchanged, and works with technician-portal photos", async () => {
  // The technician portal's work_order_photos table has no photo_stage column; the history used to fail on it.
  (await import("../utils/techActivity.js")).ensureTechSchema(db);
  db.prepare(`INSERT INTO work_order_photos (work_order_id, username, file_path, at) VALUES (1, 'tech', 'uploads/wo1.jpg', '2026-11-03 09:00:00')`).run();
  const res = await get("/api/assets/EX01/history");
  assert.equal(res.statusCode, 200, res.body);
  const one = res.json();
  assert.equal(one.history.find((e) => e.type === "work_order").details.photo, "uploads/wo1.jpg");
  assert.equal((await get("/api/assets/history.xlsx")).statusCode, 200);
  assert.equal(one.summary.counts.events_total, 5);
  assert.equal(one.summary.totals.fuel_liters_total, 350);
  const row = historyRow(one.history.find((e) => e.type === "work_order"));
  assert.deepEqual([row.type, row.status, row.work_order], ["Work order", "open", 1]);
});
