import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";
import ExcelJS from "exceljs";
import { packSizeFromName, suggestOilType } from "../utils/lubeModel.js";
import { excelDate } from "../utils/xlsxPatch.js";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-lubemodel-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { default: workOrderRoutes } = await import("../routes/workorders.routes.js");
const { default: maintenanceRoutes } = await import("../routes/maintenance.routes.js");
const { default: stockRoutes } = await import("../routes/stock.routes.js");
const app = Fastify({ logger: false });
await app.register(workOrderRoutes, { prefix: "/api/workorders" });
await app.register(maintenanceRoutes, { prefix: "/api/maintenance" });
await app.register(stockRoutes, { prefix: "/api/stock" });
await app.ready();
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

// The sign-in hook gives a storeman the "stores" role too; this test does not load it.
const H = { "x-user-name": "stores1", "x-user-role": "storeman", "x-user-roles": "storeman,stores", "x-site-code": "main" };
const SHEETS = ["1. Control_Sheet", "2.Lubes_Summary", "3.deliveries to site", "4.list_of_site_lube_suppliers", "5.previous month closing stock",
  ...Array.from({ length: 50 }, (_, i) => `INPUT ${i + 1}`)];

/** A small stand-in for the site's workbook: same sheets and input cells. */
async function fakeTemplate() {
  const wb = new ExcelJS.Workbook();
  for (const n of SHEETS) wb.addWorksheet(n);
  const c = wb.getWorksheet("1. Control_Sheet");
  c.getCell("C3").value = "AML";
  c.getCell("C7").value = 2025;
  c.getCell("C8").value = 1;
  c.getCell("E16").value = { formula: 'IF(D16="",C16,D16)' };
  const in1 = wb.getWorksheet("INPUT 1");
  in1.getCell("D6").value = "OLD-PLANT"; // last month's line: must be cleared
  in1.getCell("B6").value = { formula: 'IF(D6="","",$H$4)' };
  const d = wb.getWorksheet("3.deliveries to site");
  d.getCell("E4").value = new Date(Date.UTC(2026, 7, 3));
  d.getCell("K4").value = "Old invoice line";
  d.getCell("O4").value = 1;
  d.getCell("Q4").value = "INV-OLD";
  d.getCell("F5").value = { formula: 'IF(E5="","",MONTH(E5))' };
  wb.getWorksheet("4.list_of_site_lube_suppliers").getCell("A2").value = "Name_of_Supplier";
  wb.getWorksheet("4.list_of_site_lube_suppliers").getCell("A3").value = "Executive Logistics Pemba";
  const o = wb.getWorksheet("5.previous month closing stock");
  ["COOLANT", "ENGINE 10W40", "ENGINE 15W40", "HYDRAULIC"].forEach((t, i) => { o.getCell(`C${4 + i}`).value = t; });
  return Buffer.from(await wb.xlsx.writeBuffer());
}

function multipartBody(name, buf) {
  const boundary = "----ironloglube";
  return {
    boundary,
    body: Buffer.concat([
      Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="${name}"\r\nContent-Type: application/octet-stream\r\n\r\n`),
      buf,
      Buffer.from(`\r\n--${boundary}--\r\n`),
    ]),
  };
}

test("oil types and container sizes are read from item names", () => {
  assert.equal(suggestOilType("Fuchs Titan Truck Plus 15W-40 20L"), "ENGINE 15W40");
  assert.equal(suggestOilType("Fuchs Renolin HO68 208L"), "HYDRAULIC");
  assert.equal(suggestOilType("Fuchs TO-4 SAE50"), "TRANSMISSION 50W");
  assert.equal(suggestOilType("Fuchs EP2 grease 18kg"), "GREASE");
  assert.equal(suggestOilType("Fuchs Titan LS 80w-90"), "GEARBOX/F DRIVE 80W90");
  assert.equal(suggestOilType("Fuchs Frictolin coolant"), "COOLANT");
  assert.equal(suggestOilType("Mystery fluid"), null);
  assert.equal(packSizeFromName("Rimula R4 15W40 20L"), 20);
  assert.equal(packSizeFromName("Renolin HO68 208 LT"), 208);
  assert.equal(packSizeFromName("EP2 grease 18kg"), 18);
  assert.equal(packSizeFromName("ATF"), null);
});

test("the month's issues, deliveries and opening stock fill the model", async (t) => {
  t.after(() => app.close());
  db.exec(`
    INSERT INTO assets (id, asset_code, asset_name, category, active) VALUES (1, 'E502AM', 'CAT 350', 'Excavator', 1), (2, 'A301AM', 'Bell B30D', 'ADT', 1);
    INSERT INTO parts (id, part_code, part_name) VALUES
      (1, 'MLFPT6001', 'Fuchs Titan Truck Plus 15W-40 20L'),
      (2, 'MLFHO68', 'Fuchs Renolin HO68 208L'),
      (3, 'MLFXX', 'Fuchs special fluid');
  `);
  const loc = db.prepare("SELECT id FROM stock_locations WHERE location_code = 'MAIN'").get()
    || { id: db.prepare("INSERT INTO stock_locations (location_code, location_name, active) VALUES ('MAIN', 'Main', 1)").run().lastInsertRowid };

  // August closing stock: 12 pails of 15W40 (240 L) and 2 drums of HO68 (416 L).
  db.exec(`
    INSERT INTO stock_movements (part_id, quantity, movement_type, reference, created_at) VALUES
      (1, 12, 'in', 'opening', '2026-08-20 08:00:00'), (2, 2, 'in', 'opening', '2026-08-20 08:00:00'), (3, 5, 'in', 'opening', '2026-08-20 08:00:00');
  `);
  // A service on E502AM on 2 Sept: the lube issued that day is the 500 h service.
  db.exec(`
    INSERT INTO maintenance_plans (id, asset_id, service_name, interval_hours, last_service_hours, active) VALUES (5, 1, '500', 500, 5000, 1);
    INSERT INTO work_orders (asset_id, source, reference_id, status, opened_at, completed_at) VALUES (1, 'service', 5, 'completed', '2026-09-02 07:00:00', '2026-09-02 15:00:00');
  `);
  const issue = async (part, asset, qty, date) => {
    const r = await app.inject({ method: "POST", url: "/api/stock/lube-issue", headers: H, payload: { part_code: part, asset_code: asset, quantity: qty, log_date: date } });
    assert.equal(r.statusCode, 200, r.body);
  };
  await issue("MLFPT6001", "E502AM", 2, "2026-09-02");
  await issue("MLFHO68", "A301AM", 0.1, "2026-09-02");
  await issue("MLFPT6001", "A301AM", 1, "2026-09-09");
  await issue("MLFXX", "A301AM", 1, "2026-09-09");

  // A delivery in September.
  const dl = await app.inject({ method: "POST", url: "/api/stock/deliveries", headers: H, payload: {
    supplier: "Lubemoz", reference: "INV-3001", currency: "MZN", received_date: "2026-09-05",
    lines: [{ part_code: "MLFPT6001", quantity: 10, unit_cost: 2464 }],
  } });
  assert.equal(dl.statusCode, 200, dl.body);

  // No template yet.
  assert.equal((await app.inject({ method: "GET", url: "/api/stock/lube-model/download?month=2026-09", headers: H })).statusCode, 409);
  const bad = multipartBody("notes.xlsx", Buffer.from(await new ExcelJS.Workbook().xlsx.writeBuffer()));
  const badUp = await app.inject({ method: "POST", url: "/api/stock/lube-model/template", headers: { ...H, "content-type": `multipart/form-data; boundary=${bad.boundary}` }, payload: bad.body });
  assert.equal(badUp.statusCode, 400);
  assert.match(badUp.json().error, /not the lube model/);
  const good = multipartBody("LUBE_TEMPLATE_V14.xlsx", await fakeTemplate());
  const up = await app.inject({ method: "POST", url: "/api/stock/lube-model/template", headers: { ...H, "content-type": `multipart/form-data; boundary=${good.boundary}` }, payload: good.body });
  assert.equal(up.statusCode, 200, up.body);

  // The item with no recognisable oil type is reported, then mapped.
  let prev = (await app.inject({ method: "GET", url: "/api/stock/lube-model/preview?month=2026-09", headers: H })).json();
  assert.deepEqual(prev.unmapped.map((u) => u.part_code), ["MLFXX"]);
  const map = await app.inject({ method: "PUT", url: "/api/stock/lube-model/parts/MLFXX", headers: H, payload: { oil_type: "ATF", pack_size: 5 } });
  assert.equal(map.statusCode, 200, map.body);
  assert.equal((await app.inject({ method: "PUT", url: "/api/stock/lube-model/parts/MLFXX", headers: H, payload: { oil_type: "OLIVE OIL" } })).statusCode, 400);
  prev = (await app.inject({ method: "GET", url: "/api/stock/lube-model/preview?month=2026-09", headers: H })).json();
  assert.equal(prev.unmapped.length, 0, JSON.stringify(prev.unmapped));
  assert.equal(prev.issue_days, 2);
  assert.equal(prev.issue_litres, 40 + 20.8 + 20 + 5);

  const file = await app.inject({ method: "GET", url: "/api/stock/lube-model/download?month=2026-09", headers: H });
  assert.equal(file.statusCode, 200, file.body);
  assert.match(file.headers["content-disposition"], /LUBE_MODEL_SEP26_IRONLOG\.xlsx/);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(file.rawPayload);
  const v = (sheet, ref) => {
    const val = wb.getWorksheet(sheet).getCell(ref).value;
    return val && typeof val === "object" && "richText" in val ? val.richText.map((r) => r.text).join("") : val;
  };
  assert.equal(v("1. Control_Sheet", "C7"), 2026);
  assert.equal(v("1. Control_Sheet", "C8"), 9);
  assert.equal(v("1. Control_Sheet", "F16"), 2);
  assert.equal(v("1. Control_Sheet", "F17"), 9);
  assert.equal(v("1. Control_Sheet", "F18"), null);
  assert.deepEqual(wb.getWorksheet("1. Control_Sheet").getCell("E16").value, { formula: 'IF(D16="",C16,D16)' }, "formulas untouched");
  // Day 2: E502AM 2 pails = 40 L on the 500 h service; A301AM 0.1 drum = 20.8 L top-up.
  const day1 = [6, 7].map((r) => ["D", "E", "F", "H"].map((c) => v("INPUT 1", `${c}${r}`)));
  assert.deepEqual(day1.sort(), [["A301AM", "HYDRAULIC", 20.8, "TOP-UP"], ["E502AM", "ENGINE 15W40", 40, "500HR_SERVICE"]].sort());
  assert.equal(v("INPUT 1", "D8"), null, "last month's lines are cleared");
  assert.deepEqual(wb.getWorksheet("INPUT 1").getCell("B6").value?.formula, 'IF(D6="","",$H$4)');
  assert.equal(v("INPUT 2", "D6") && v("INPUT 2", "D7") ? 2 : 0, 2);
  assert.equal(v("INPUT 3", "D6"), null);
  // Delivery added under the existing line; supplier joins the drop-down list.
  assert.equal(v("3.deliveries to site", "K4"), "Old invoice line");
  assert.equal(v("3.deliveries to site", "K5"), "Fuchs Titan Truck Plus 15W-40 20L");
  assert.equal(v("3.deliveries to site", "L5"), "ENGINE 15W40");
  assert.deepEqual([v("3.deliveries to site", "N5"), v("3.deliveries to site", "O5"), v("3.deliveries to site", "Q5"), v("3.deliveries to site", "R5"), v("3.deliveries to site", "U5")], [20, 10, "INV-3001", 24640, 64]);
  // An Excel date serial (the real template's date format shows it as 05/09/2026).
  assert.equal(v("3.deliveries to site", "E5"), excelDate("2026-09-05"));
  assert.equal(v("4.list_of_site_lube_suppliers", "A4"), "Lubemoz");
  // Opening stock (end of August) by oil type in litres.
  const opening = {};
  for (let r = 4; r < 8; r += 1) if (v("5.previous month closing stock", `N${r}`)) opening[v("5.previous month closing stock", `N${r}`)] = v("5.previous month closing stock", `P${r}`);
  assert.deepEqual(opening, { "ENGINE 15W40": 240, HYDRAULIC: 416, ATF: 25 });
  assert.equal(v("5.previous month closing stock", "B6"), "ENGINE 15W40");

  // Downloading again does not add the delivery twice.
  const again = await app.inject({ method: "GET", url: "/api/stock/lube-model/download?month=2026-09", headers: H });
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(again.rawPayload);
  assert.equal(wb2.getWorksheet("3.deliveries to site").getCell("K6").value, null);
});
