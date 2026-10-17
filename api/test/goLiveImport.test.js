import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import Fastify from "fastify";
import ExcelJS from "exceljs";
import { cellText, readNumber } from "../utils/goLiveImport.js";

// Connecting to other systems, part C: the guided go-live import from one Excel workbook.
test("cells and numbers are read the way people type them", () => {
  assert.equal(cellText(new Date(Date.UTC(2026, 9, 12))), "2026-10-12");
  assert.equal(cellText({ richText: [{ text: "A3" }, { text: "02AM" }] }), "A302AM");
  assert.equal(cellText({ formula: "1+1", result: 2 }), "2");
  assert.deepEqual([readNumber("1 234,5"), readNumber("1,234.5"), readNumber(""), Number.isNaN(readNumber("abc"))], [1234.5, 1234.5, null, true]);
});

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-golive-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";

await import("../db/migrate.js");
(await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
const { db } = await import("../db/client.js");
const { ironlogAuthHook } = await import("../auth/hook.js");
const realSetInterval = globalThis.setInterval;
globalThis.setInterval = (...args) => realSetInterval(...args).unref();
const app = Fastify({ logger: false });
app.addHook("onRequest", ironlogAuthHook);
app.decorateRequest("ironlogAllowedSites", null);
await app.register((await import("../routes/auth.routes.js")).default, { prefix: "/api/auth" });
// The stores and procurement routes make the stock and supplier tables.
await app.register((await import("../routes/stock.routes.js")).default, { prefix: "/api/stock" });
await app.register((await import("../routes/procurement.routes.js")).default, { prefix: "/api/procurement" });
await app.register((await import("../routes/upload.routes.js")).default, { prefix: "/api/upload" });
await app.ready();
globalThis.setInterval = realSetInterval;
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

const session = (username, role, allowed) => {
  db.prepare(`INSERT INTO users (username, role, roles_json, active, allowed_locations) VALUES (?, ?, ?, 1, ?)`).run(username, role, JSON.stringify([role]), allowed ? JSON.stringify(allowed) : null);
  const token = crypto.randomBytes(16).toString("hex");
  db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, (SELECT id FROM users WHERE username = ?), datetime('now', '+1 day'))`).run(token, username);
  return (site = "main") => ({ authorization: `Bearer ${token}`, "x-site-code": site });
};
const plantAdmin = session("pa", "workshop_admin");
const stores = session("st", "stores");
const northOnly = session("no", "workshop_admin", ["north"]);
db.prepare(`INSERT INTO assets (asset_code, asset_name, category, active, site_code) VALUES ('N1', 'North truck', 'ADT', 1, 'north')`).run();

/** A filled-in workbook built from the template the server hands out. */
async function workbook(fill, headers = plantAdmin()) {
  const res = await app.inject({ method: "GET", url: "/api/upload/go-live/template.xlsx", headers });
  assert.equal(res.statusCode, 200, res.body);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(res.rawPayload);
  for (const [sheet, rows] of Object.entries(fill)) rows.forEach((r) => wb.getWorksheet(sheet).addRow(r));
  return Buffer.from(await wb.xlsx.writeBuffer());
}
function upload(url, buffer, headers) {
  const boundary = "----ironlog" + crypto.randomBytes(6).toString("hex");
  const body = Buffer.concat([
    Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="golive.xlsx"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`),
    buffer,
    Buffer.from(`\r\n--${boundary}--\r\n`),
  ]);
  return app.inject({ method: "POST", url, headers: { ...headers, "content-type": `multipart/form-data; boundary=${boundary}` }, payload: body });
}

const GOOD = {
  // Machine code, name, category, site, standby, area, make, model, serial, year, purchase date, price, warranty, cost centre, fuel L/hr, meter
  Machines: [
    ["A302AM", "Bell B30D ADT", "ADT", "", "no", "Pit", "Bell", "B30D", "AEBA830DJ06105875", 2019, new Date(Date.UTC(2019, 4, 2)), "185000", "02/05/2021", "PIT", 9, 10450],
    ["E502AM", "CAT 336 Excavator", "Excavator", "main", "", "Pit", "CAT", "336", "", "", "", "", "", "PIT", "", ""],
  ],
  Parts: [["210105", "LEVEL INDICATOR", "Parts", 19.5, 2, "yes"], ["OIL-15W40", "ENGINE OIL 15W40 (L)", "Oils & lubricants", 3.2, 200, "no"]],
  "Stock on hand": [["210105", 3], ["OIL-15W40", 410]],
  Suppliers: [["000028", "BELL EQUIPMENT SALES SA LTD -P", "zar", 21]],
  Components: [["A302AM", "eng", "", "QSK-1", "Cummins", "QSB6.7", "2024-01-10", 2000, 12000]],
};

test("the template has a read-me and one sheet per kind of record, with the help on each heading", async () => {
  const res = await app.inject({ method: "GET", url: "/api/upload/go-live/template.xlsx", headers: plantAdmin() });
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(res.rawPayload);
  assert.deepEqual(wb.worksheets.map((w) => w.name), ["Read me", "Machines", "Parts", "Stock on hand", "Suppliers", "Components"]);
  assert.equal(wb.getWorksheet("Machines").getRow(1).getCell(1).value, "Machine code *");
  assert.match(String(wb.getWorksheet("Machines").getRow(1).getCell(1).note), /unique/);
  assert.equal(wb.getWorksheet("Machines").rowCount, 1, "no empty rows under the heading");
  assert.deepEqual(wb.getWorksheet("Machines").getCell("E40").dataValidation.formulae, ['"yes,no"'], "yes/no picker on Standby");
  assert.equal((await app.inject({ method: "GET", url: "/api/upload/go-live/template.xlsx", headers: stores() })).statusCode, 403, "administrators and plant administrators only");
});

test("check lists every problem by sheet, row and column and changes nothing", async () => {
  const bad = await workbook({
    Machines: [["", "No code", "ADT"], ["X 1", "Bad code", "ADT"], ["B1", "Moon truck", "ADT", "moon"], ["N1", "Moved", "ADT", "main"], ["B2", "Bad year", "ADT", "", "maybe", "", "", "", "", 1800]],
    Parts: [["P1", "Bolt", "Rockets", -1]],
    "Stock on hand": [["NOPE", 2], ["P1", "lots"]],
    Suppliers: [["S1", "Shop", "Rand"]],
    Components: [["GHOST", "ENG"], ["N1", "ENG"]],
  });
  const res = await upload("/api/upload/go-live/check", bad, plantAdmin());
  assert.equal(res.statusCode, 200, res.body);
  const r = res.json();
  assert.equal(r.ok, false);
  const where = r.errors.map((e) => `${e.sheet}:${e.row}:${e.column}`);
  for (const w of ["Machines:2:Machine code", "Machines:3:Machine code", "Machines:4:Site", "Machines:5:Site", "Machines:6:Standby", "Machines:6:",
    "Parts:2:Category", "Parts:2:Unit cost (USD)", "Stock on hand:2:Part code", "Stock on hand:3:Quantity on hand", "Suppliers:2:Currency", "Components:2:Machine code"]) {
    assert.ok(where.includes(w), `expected a problem at ${w}; got ${where.join(", ")}`);
  }
  assert.match(r.errors.find((e) => e.row === 5 && e.sheet === "Machines").message, /already on site north/);
  assert.equal(db.prepare(`SELECT COUNT(*) AS c FROM assets`).get().c, 1, "nothing written");
  assert.equal((await upload("/api/upload/go-live/import", bad, plantAdmin())).statusCode, 409, "import refuses a file with problems");
  assert.match((await upload("/api/upload/go-live/check", Buffer.from("not excel"), plantAdmin())).json().error, /not an Excel workbook/);
});

test("a clean file imports everything at once, and the same file again changes nothing", async () => {
  const file = await workbook(GOOD);
  const check = (await upload("/api/upload/go-live/check", file, plantAdmin())).json();
  assert.equal(check.ok, true, JSON.stringify(check.errors));
  assert.deepEqual(check.summary.machines, { new: 2, changed: 0, same: 0 });
  assert.deepEqual(check.summary.components, { new: 1, changed: 0, same: 0 });

  const done = await upload("/api/upload/go-live/import", file, plantAdmin());
  assert.equal(done.statusCode, 200, done.body);
  const a = db.prepare(`SELECT * FROM assets WHERE asset_code = 'A302AM'`).get();
  assert.deepEqual([a.site_code, a.make, a.model, a.year_built, a.purchase_date, a.purchase_price_usd, a.warranty_end, a.cost_center_code, a.baseline_fuel_l_per_hour, a.is_standby],
    ["main", "Bell", "B30D", 2019, "2019-05-02", 185000, "2021-05-02", "PIT", 9, 0], "Excel dates and DD/MM/YYYY both read as dates");
  assert.equal(db.prepare(`SELECT total_hours FROM asset_hours WHERE asset_id = ?`).get(a.id).total_hours, 10450);
  const part = db.prepare(`SELECT * FROM parts WHERE part_code = 'OIL-15W40'`).get();
  assert.deepEqual([part.stock_category, part.stock_category_source, part.unit_cost, part.min_stock, part.critical], ["oil", "manual", 3.2, 200, 0]);
  assert.equal(db.prepare(`SELECT SUM(quantity) AS q FROM stock_movements WHERE part_id = ?`).get(part.id).q, 410);
  assert.deepEqual({ ...db.prepare(`SELECT currency, lead_time_days FROM suppliers WHERE supplier_code = '000028'`).get() }, { currency: "ZAR", lead_time_days: 21 });
  const comp = db.prepare(`SELECT * FROM asset_components WHERE asset_id = ?`).get(a.id);
  assert.deepEqual([comp.component, comp.serial_number, comp.fitted_hours, comp.expected_life_hours, comp.created_by], ["ENG", "QSK-1", 2000, 12000, "pa"]);

  const again = (await upload("/api/upload/go-live/check", file, plantAdmin())).json();
  for (const s of ["machines", "parts", "stock", "suppliers", "components"]) assert.equal(again.summary[s].new + again.summary[s].changed, 0, `${s} unchanged on a second upload`);
  await upload("/api/upload/go-live/import", file, plantAdmin());
  assert.equal(db.prepare(`SELECT COUNT(*) AS c FROM stock_movements WHERE reference = 'go_live_opening'`).get().c, 2, "stock is set, not added twice");
  assert.equal(db.prepare(`SELECT COUNT(*) AS c FROM asset_components`).get().c, 1);

  // A recount sets the balance by posting the difference, and says so when stock has moved.
  db.prepare(`INSERT INTO stock_movements (part_id, quantity, movement_type, reference) VALUES (?, -10, 'out', 'work_order:1')`).run(part.id);
  const recount = await workbook({ "Stock on hand": [["OIL-15W40", 395]] });
  const rc = (await upload("/api/upload/go-live/check", recount, plantAdmin())).json();
  assert.match(rc.warnings[0].message, /IronLog has 400 on hand; this posts -5 to make it 395/);
  await upload("/api/upload/go-live/import", recount, plantAdmin());
  assert.equal(db.prepare(`SELECT SUM(quantity) AS q FROM stock_movements WHERE part_id = ?`).get(part.id).q, 395);

  const status = (await app.inject({ method: "GET", url: "/api/upload/go-live/status", headers: plantAdmin() })).json();
  const step = (k) => status.steps.find((s) => s.key === k);
  assert.deepEqual([step("machines").done, step("meters").detail, step("suppliers").done, step("hours").done], [true, "1 of 2", true, false]);
  assert.equal(status.history.length, 3);
  assert.equal(status.history[2].summary.machines.new, 2);
  assert.equal(db.prepare(`SELECT COUNT(*) AS c FROM audit_logs WHERE action = 'go_live.import'`).get().c, 3);
});

test("the filled-in template carries the site's data back out, and site limits hold", async () => {
  const res = await app.inject({ method: "GET", url: "/api/upload/go-live/template.xlsx?filled=1", headers: plantAdmin() });
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(res.rawPayload);
  const codes = [];
  wb.getWorksheet("Machines").eachRow((r, n) => { if (n > 1) codes.push(r.getCell(1).value); });
  assert.deepEqual(codes, ["A302AM", "E502AM"], "only this site's machines");
  assert.equal(wb.getWorksheet("Machines").getRow(2).getCell(16).value, 10450);

  const north = await workbook({ Machines: [["N2", "North loader", "Loader"]] }, northOnly("north"));
  assert.equal((await upload("/api/upload/go-live/check", north, northOnly("north"))).json().ok, true);
  const south = await workbook({ Machines: [["S9", "South loader", "Loader", "main"]] }, northOnly("north"));
  const r = (await upload("/api/upload/go-live/check", south, northOnly("north"))).json();
  assert.match(r.errors[0].message, /may not add machines to site main/);
});
