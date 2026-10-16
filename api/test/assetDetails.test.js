import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import Fastify from "fastify";

// Asset structure part 1: what a machine is (make, model, serial, year, purchase, warranty).
const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-assetdetails-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";

await import("../db/migrate.js");
(await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
const { db } = await import("../db/client.js");
const { ironlogAuthHook } = await import("../auth/hook.js");
const details = await import("../utils/assetDetails.js");
const realSetInterval = globalThis.setInterval;
globalThis.setInterval = (...args) => realSetInterval(...args).unref();
const app = Fastify({ logger: false });
app.addHook("onRequest", ironlogAuthHook);
app.decorateRequest("ironlogAllowedSites", null);
await app.register((await import("../routes/auth.routes.js")).default, { prefix: "/api/auth" });
await app.register((await import("../routes/assets.routes.js")).default, { prefix: "/api/assets" });
await app.register((await import("../routes/upload.routes.js")).default, { prefix: "/api/upload" });
await app.register((await import("../routes/reports.routes.js")).default, { prefix: "/api/reports" });
await app.ready();
globalThis.setInterval = realSetInterval;
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

const session = (username, role) => {
  db.prepare(`INSERT INTO users (username, role, roles_json, active) VALUES (?, ?, ?, 1)`).run(username, role, JSON.stringify([role]));
  const token = crypto.randomBytes(16).toString("hex");
  const id = db.prepare(`SELECT id FROM users WHERE username = ?`).get(username).id;
  db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, ?, datetime('now', '+1 day'))`).run(token, id);
  return { authorization: `Bearer ${token}` };
};
const foreman = session("foreman", "plant_manager");
const tech = session("tech", "artisan");

test("machine details are checked: dates, year, money, warranty after purchase", () => {
  const ok = details.readAssetDetails({ make: " Caterpillar ", model: "140K", year_built: "2019", purchase_price_usd: "$285,000.50", purchase_date: "2019-03-01", warranty_end: "", serial_number: null }, { thisYear: 2026 });
  assert.deepEqual(ok, { ok: true, values: { make: "Caterpillar", model: "140K", year_built: 2019, purchase_date: "2019-03-01", purchase_price_usd: 285000.5, warranty_end: null, serial_number: null } });
  assert.deepEqual(details.readAssetDetails({ notes: "x" }), { ok: true, values: {} }, "only the fields sent");
  const bad = details.readAssetDetails({ year_built: "1890", purchase_date: "2019-02-30", purchase_price_usd: "-5", make: "x".repeat(61) }, { thisYear: 2026 });
  assert.equal(bad.ok, false);
  assert.equal(bad.errors.length, 4);
  assert.match(details.readAssetDetails({ purchase_date: "2020-01-01", warranty_end: "2019-12-31" }).errors[0], /Warranty ends before the purchase date/);
  assert.deepEqual(details.warrantyStatus("2026-11-15", "2026-10-16"), { in_warranty: true, days_left: 30, text: "In warranty, 30 days left" });
  assert.equal(details.warrantyStatus("2026-01-31", "2026-10-16").text, "Warranty ended 2026-01-31");
  assert.equal(details.warrantyStatus(null), null);
  assert.equal(details.assetDetailLine({ make: "Caterpillar", model: "140K", serial_number: "R9J01234", year_built: 2019 }), "Caterpillar 140K · S/N R9J01234 · 2019");
});

test("foremen add and edit a machine's details; technicians cannot change machines", async () => {
  const created = await app.inject({ method: "POST", url: "/api/assets", headers: foreman, payload: { asset_code: "G02AM", asset_name: "Motor grader", category: "Grader", make: "Caterpillar", model: "140K", year_built: 2019 } });
  assert.equal(created.statusCode, 201);
  assert.equal((await app.inject({ method: "POST", url: "/api/assets", headers: foreman, payload: { asset_code: "X1", asset_name: "Bad", year_built: 3000 } })).statusCode, 400);
  assert.equal(db.prepare(`SELECT COUNT(*) AS n FROM assets WHERE asset_code = 'X1'`).get().n, 0, "a refused machine is not added");

  const patched = await app.inject({ method: "PATCH", url: "/api/assets/G02AM", headers: foreman, payload: { serial_number: "R9J01234", purchase_date: "2019-03-01", purchase_price_usd: "285000", warranty_end: "2027-03-01", model: "" } });
  assert.equal(patched.statusCode, 200);
  assert.deepEqual(
    [patched.json().asset.make, patched.json().asset.model, patched.json().asset.serial_number, patched.json().asset.purchase_price_usd, patched.json().asset.warranty_end],
    ["Caterpillar", null, "R9J01234", 285000, "2027-03-01"], "a blank clears; untouched fields stay");
  const badPatch = await app.inject({ method: "PATCH", url: "/api/assets/G02AM", headers: foreman, payload: { warranty_end: "2018-01-01" } });
  assert.equal(badPatch.statusCode, 400);
  assert.match(badPatch.json().error, /Warranty ends before the purchase date/, "checked against the saved purchase date");

  const xlsx = await app.inject({ method: "GET", url: "/api/assets/cost-centers.xlsx", headers: foreman });
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(xlsx.rawPayload);
  const sheet = wb.getWorksheet("Equipment");
  const header = sheet.getRow(1).values.slice(1);
  const g02 = sheet.getRow(2).values.slice(1);
  assert.equal(g02[header.indexOf("Serial number")], "R9J01234", "the equipment register spreadsheet carries the details");
  assert.equal(g02[header.indexOf("Purchase price (USD)")], 285000);

  const card = (await app.inject({ method: "GET", url: "/api/assets/fleet-summary", headers: foreman })).json().cards.find((c) => c.asset_code === "G02AM");
  assert.deepEqual([card.make, card.serial_number, card.year_built], ["Caterpillar", "R9J01234", 2019], "fleet cards carry the details");

  assert.equal((await app.inject({ method: "PATCH", url: "/api/assets/G02AM", headers: tech, payload: { serial_number: "HACK" } })).statusCode, 403);
  assert.equal((await app.inject({ method: "POST", url: "/api/assets", headers: tech, payload: { asset_code: "T1", asset_name: "x" } })).statusCode, 403);
  assert.equal((await app.inject({ method: "POST", url: "/api/assets/G02AM/archive", headers: tech, payload: { archived: true } })).statusCode, 403);
  assert.equal(db.prepare(`SELECT serial_number, archived FROM assets WHERE asset_code = 'G02AM'`).get().serial_number, "R9J01234");
});

test("the machines CSV upload takes the details too, and reports a bad value without saving it", async () => {
  const csv = "asset_code,asset_name,category,make,model,serial_number,year_built,warranty_end\n"
    + "E502AM,Excavator 50t,Excavator,Komatsu,PC500,KMT500X,2021,2027-06-30\n"
    + "E503AM,Excavator 50t,Excavator,Komatsu,PC500,KMT501X,1800,\n";
  const boundary = "----ironlogtest";
  const body = `--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="assets.csv"\r\nContent-Type: text/csv\r\n\r\n${csv}\r\n--${boundary}--\r\n`;
  const admin = session("boss", "admin");
  const res = await app.inject({ method: "POST", url: "/api/upload/assets", headers: { ...admin, "content-type": `multipart/form-data; boundary=${boundary}` }, payload: body });
  assert.equal(res.statusCode, 200);
  assert.equal(res.json().warnings.length, 1);
  assert.match(res.json().warnings[0], /Row 3 \(E503AM\).*Year built/);
  assert.deepEqual(db.prepare(`SELECT make, model, serial_number, year_built, warranty_end FROM assets WHERE asset_code = 'E502AM'`).get(),
    { make: "Komatsu", model: "PC500", serial_number: "KMT500X", year_built: 2021, warranty_end: "2027-06-30" });
  assert.deepEqual(db.prepare(`SELECT make, year_built FROM assets WHERE asset_code = 'E503AM'`).get(), { make: null, year_built: null }, "the machine is added, its bad details are not");
});

test("the machine history PDF shows what the machine is", async () => {
  // Added on the live database by the work order routes, which this test does not load.
  const woCols = db.prepare(`PRAGMA table_info(work_orders)`).all().map((c) => c.name);
  if (!woCols.includes("completed_at")) db.prepare(`ALTER TABLE work_orders ADD COLUMN completed_at TEXT`).run();
  const PDFDocument = (await import("pdfkit")).default;
  const written = [];
  const text = PDFDocument.prototype.text;
  PDFDocument.prototype.text = function (s, ...rest) { written.push(String(s)); return text.call(this, s, ...rest); };
  try {
    const res = await app.inject({ method: "GET", url: "/api/reports/asset-history/G02AM.pdf", headers: foreman });
    assert.equal(res.statusCode, 200);
    assert.equal(res.rawPayload.subarray(0, 4).toString(), "%PDF");
  } finally {
    PDFDocument.prototype.text = text;
  }
  const all = written.join("\n");
  assert.match(all, /Make \/ model/);
  assert.match(all, /R9J01234/);
  assert.match(all, /\$285,000\.00|\$285000\.00/);
  assert.match(all, /In warranty|Warranty ended/);
});
