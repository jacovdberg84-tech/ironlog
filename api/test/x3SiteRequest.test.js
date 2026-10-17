import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import Database from "better-sqlite3";
import Fastify from "fastify";
import * as x3 from "../utils/x3SiteRequest.js";

// Connecting to other systems: IronLog requisitions as Sage X3 Site Requests (AML's OZS1 screen).
const SETTINGS = {
  delivery_site: "AML01", purchase_site: "VS120", request_type: "AML-Plant", signature_channel: "AML-Plant", requested_by: "BER03", buyer: "MAL01",
  customer: "AML009", bill_to_customer: "AML009", delivery_address: "001", project: "AML-3001", item_code: "120-00110",
  user_codes: "jaco = BER03\nmario=MAL01", supplier_codes: "Bell Equipment Sales SA Ltd -P = 000028",
};
const REQ = {
  requisition: { id: 42, status: "posted", requester: "Jaco", supplier_name: "BELL EQUIPMENT SALES SA LTD -P", po_number: "2225533", needed_by_date: "2026-10-12" },
  lines: [
    { line_no: 1, product_code: "210105", description: "LEVEL INDICATOR", unit: "UN", quantity: 3, equipment_no: "A302AM", job_card: "WO-881", currency: "ZAR", gross_price: 364.4, discount_type: "Value", discount: 0, net_price: 364.4 },
    { line_no: 2, product_code: "140466", description: "O RING, \"3AO185AON70\"", quantity: 3, equipment_no: "A302AM", currency: "ZAR", gross_price: 28.85 },
  ],
  asset: { serial_number: "AEBA830DJ06105875", meter: 10450.5 },
  supplier_code: "",
};

test("a requisition becomes an X3 Site Request: header from settings and the machine, a line per part", () => {
  const b = x3.buildX3SiteRequests([REQ], SETTINGS, { today: "2026-10-17" });
  assert.equal(b.ok, true, b.errors.join());
  assert.deepEqual(b.warnings, []);
  const [r] = b.records;
  assert.equal(r.ref, "IL000042");
  assert.deepEqual([r.header.ZDELFCY, r.header.ZFCY, r.header.ZREQUESTBY, r.header.ZBUYER, r.header.ZHBPSNUM, r.header.ZSUPPQUOTE],
    ["AML01", "VS120", "BER03", "MAL01", "000028", "2225533"], "requester and supplier names matched to X3 codes, ignoring case");
  assert.deepEqual([r.header.ZAX1, r.header.ZVIN, r.header.ZODOREADING, r.header.ZSITEJOB, r.header.ZDATEREQ, r.header.ZAX3],
    ["A302AM", "AEBA830DJ06105875", "0", "WO-881", "12/10/2026", "AML-3001"], "no ODO unit set: reading 0, as X3 shows it");
  assert.deepEqual([r.lines[1].ZITMREF, r.lines[1].ZOEM, r.lines[1].ZUOM, r.lines[1].ZNETPRI, r.lines[1].ZLCUR, r.lines[1].ZDATEREQ],
    ["120-00110", "140466", "UN", "28.85", "ZAR", "12/10/2026"], "unit from settings, net price from gross, date from the header");

  const withOdo = x3.buildX3SiteRequests([REQ], { ...SETTINGS, odo_unit: "HR" }, { today: "2026-10-17" }).records[0];
  assert.deepEqual([withOdo.header.ZODOUNIT, withOdo.header.ZODOREADING, withOdo.lines[0].ZODOREADING], ["HR", "10450.5", "10450.5"]);

  const file = x3.x3SiteRequestFile(b.records, SETTINGS).split("\r\n");
  assert.equal(file.length, 4, "E, L, L and the closing line break");
  assert.equal(file[0].split(";").length, 1 + x3.X3_HEADER_FIELDS.length);
  assert.match(file[0], /^"E";"2";"AML01";"VS120";"IL000042";"USD";"";"1";"AML-Plant";"AML-Plant";"BER03";"MAL01";"12\/10\/2026";"000028";/);
  assert.match(file[2], /^"L";"120-00110";"140466";"O RING, ""3AO185AON70""";"UN";"3";"A302AM";/, "quotes doubled, commas kept inside quotes");
  assert.equal(file[1].split(";").length, 1 + x3.X3_LINE_FIELDS.length);
});

test("what X3 needs is asked for before a file is made", () => {
  const empty = x3.buildX3SiteRequests([REQ], {}, { today: "2026-10-17" });
  assert.equal(empty.ok, false);
  assert.match(empty.errors[0], /Fill in the X3 settings first: Delivery Site, Purchase Site, Request type, Signature Channel, Buyer, Product\./);

  const odd = { ...REQ, requisition: { ...REQ.requisition, status: "draft", requester: "someone", supplier_name: "Unknown Parts Ltd", x3_site_request_no: "VS120-00400" }, lines: [{ ...REQ.lines[0], equipment_no: "", quantity: 0 }] };
  const b = x3.buildX3SiteRequests([odd], { ...SETTINGS, requested_by: "" }, { today: "2026-10-17" });
  assert.deepEqual(b.errors, [
    "Requisition 42 is still a draft. Finalize it first.",
    "Requisition 42: no X3 code for someone. Add them to the people list or set \"Requested By\".",
    "Requisition 42 line 1: the quantity must be more than 0.",
  ]);
  assert.equal(b.warnings.length, 3, "unknown supplier, no machine, already in X3");
  assert.match(b.warnings[2], /already in X3 as VS120-00400/);
  assert.equal(b.records[0].header.ZDATEREQ, "12/10/2026");
  assert.equal(x3.buildX3SiteRequests([], SETTINGS, { today: "2026-10-17" }).errors[0], "Choose at least one requisition.");
});

test("settings are checked before they are saved", () => {
  const d = new Database(":memory:");
  const bad = x3.saveX3Settings(d, { separator: "#", delivery_site: "AML 01;", user_codes: "jaco BER03", header_marker: "L" });
  assert.equal(bad.ok, false);
  assert.equal(bad.errors.length, 4);
  const ok = x3.saveX3Settings(d, { delivery_site: "AML01", date_format: "DDMMYY", user_codes: "jaco = BER03" }, "buyer");
  assert.equal(ok.ok, true, ok.errors?.join());
  assert.deepEqual([ok.settings.delivery_site, ok.settings.date_format, ok.settings.separator, ok.settings.updated_by], ["AML01", "DDMMYY", ";", "buyer"]);
  assert.equal(x3.formatX3Date("2026-10-12", "DDMMYY"), "121026");
});

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-x3-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";

test("stores download the file for their site; the X3 number is typed back", async (t) => {
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
  // The stores routes make the stock tables the procurement routes build on.
  await app.register((await import("../routes/stock.routes.js")).default, { prefix: "/api/stock" });
  await app.register((await import("../routes/procurement.routes.js")).default, { prefix: "/api/procurement" });
  await app.ready();
  globalThis.setInterval = realSetInterval;
  t.after(async () => { await app.close(); try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

  const session = (username, role) => {
    db.prepare(`INSERT INTO users (username, role, roles_json, active) VALUES (?, ?, ?, 1)`).run(username, role, JSON.stringify([role]));
    const token = crypto.randomBytes(16).toString("hex");
    db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, (SELECT id FROM users WHERE username = ?), datetime('now', '+1 day'))`).run(token, username);
    return (site = "main") => ({ authorization: `Bearer ${token}`, "x-site-code": site });
  };
  const stores = session("jaco", "stores");
  const tech = session("tech", "artisan");

  const assetId = db.prepare(`INSERT INTO assets (asset_code, asset_name, category, active, site_code) VALUES ('A302AM', 'Bell B30D', 'ADT', 1, 'main')`).run().lastInsertRowid;
  try { db.prepare(`UPDATE assets SET serial_number = 'AEBA830DJ06105875' WHERE id = ?`).run(assetId); } catch { /* column added by the asset routes */ }
  const partId = db.prepare(`INSERT INTO parts (part_code, part_name) VALUES ('210105', 'LEVEL INDICATOR')`).run().lastInsertRowid;
  db.prepare(`INSERT INTO suppliers (supplier_code, supplier_name) VALUES ('000028', 'BELL EQUIPMENT SALES SA LTD -P')`).run();
  const mkReq = (site, status) => {
    const id = db.prepare(`INSERT INTO procurement_requisitions (part_id, qty_requested, supplier_name, requester, status, site_code, needed_by_date) VALUES (?, 3, 'BELL EQUIPMENT SALES SA LTD -P', 'jaco', ?, ?, '2026-10-12')`).run(partId, status, site).lastInsertRowid;
    db.prepare(`INSERT INTO procurement_requisition_lines (requisition_id, line_no, product_code, part_id, description, quantity, equipment_no, gross_price) VALUES (?, 1, '210105', ?, 'LEVEL INDICATOR', 3, 'A302AM', 364.4)`).run(id, partId);
    return id;
  };
  const r1 = mkReq("main", "posted");
  const other = mkReq("north", "posted");

  const url = "/api/procurement/x3/site-requests";
  assert.equal((await app.inject({ method: "PUT", url: `${url}/settings`, headers: tech(), payload: SETTINGS })).statusCode, 403);
  const blocked = await app.inject({ method: "POST", url: `${url}/export`, headers: stores(), payload: { ids: [r1] } });
  assert.equal(blocked.statusCode, 409, "settings missing");
  assert.equal((await app.inject({ method: "PUT", url: `${url}/settings`, headers: session("boss", "admin")(), payload: SETTINGS })).statusCode, 200);

  assert.equal((await app.inject({ method: "POST", url: `${url}/check`, headers: stores(), payload: { ids: [r1, other] } })).statusCode, 404, "not another site's requisition");
  const check = (await app.inject({ method: "POST", url: `${url}/check`, headers: stores(), payload: { ids: [r1] } })).json();
  assert.equal(check.ok, true, check.errors.join());
  assert.equal(check.records[0].header.ZHBPSNUM, "000028", "the IronLog supplier code is used");

  assert.equal((await app.inject({ method: "POST", url: `${url}/export`, headers: tech(), payload: { ids: [r1] } })).statusCode, 403);
  const file = await app.inject({ method: "POST", url: `${url}/export`, headers: stores(), payload: { ids: [r1] } });
  assert.equal(file.statusCode, 200, file.body);
  assert.match(file.headers["content-disposition"], new RegExp(`X3_SiteRequests_\\d{8}_REQ${r1}\\.csv`));
  assert.match(file.body, /^"E";"2";"AML01";"VS120";"IL0000\d\d";/);
  assert.match(file.body, /\r\n"L";"120-00110";"210105";"LEVEL INDICATOR";"UN";"3";"A302AM";/);
  assert.ok(db.prepare(`SELECT x3_exported_at FROM procurement_requisitions WHERE id = ?`).get(r1).x3_exported_at);

  assert.equal((await app.inject({ method: "PUT", url: `/api/procurement/requisitions/${r1}/x3-number`, headers: stores(), payload: { x3_site_request_no: "VS120 00353;" } })).statusCode, 400);
  assert.equal((await app.inject({ method: "PUT", url: `/api/procurement/requisitions/${other}/x3-number`, headers: stores(), payload: { x3_site_request_no: "VS120-00353" } })).statusCode, 404);
  assert.equal((await app.inject({ method: "PUT", url: `/api/procurement/requisitions/${r1}/x3-number`, headers: stores(), payload: { x3_site_request_no: "VS120-00353" } })).statusCode, 200);
  const listed = (await app.inject({ method: "GET", url: "/api/procurement/requisitions", headers: stores() })).json().rows.find((r) => r.id === r1);
  assert.equal(listed.x3_site_request_no, "VS120-00353");
  const again = (await app.inject({ method: "POST", url: `${url}/check`, headers: stores(), payload: { ids: [r1] } })).json();
  assert.match(again.warnings[0], /already in X3 as VS120-00353/);
  const audit = db.prepare(`SELECT action FROM audit_logs WHERE module = 'procurement' AND action LIKE 'x3.%' ORDER BY id`).all().map((r) => r.action);
  assert.deepEqual(audit, ["x3.settings", "x3.export", "x3.number"]);
});
