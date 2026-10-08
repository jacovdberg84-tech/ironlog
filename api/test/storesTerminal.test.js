import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-terminal-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { ensureTechSchema } = await import("../utils/techActivity.js");
const { lookupCode } = await import("../routes/stock/terminal.routes.js");
const { isPublicAuthRequest } = await import("../auth/config.js");
const { default: stockRoutes } = await import("../routes/stock.routes.js");
const { default: workOrderRoutes } = await import("../routes/workorders.routes.js");
const { default: breakdownRoutes } = await import("../routes/breakdowns.routes.js");
const app = Fastify({ logger: false });
await app.register(workOrderRoutes, { prefix: "/api/workorders" });
await app.register(breakdownRoutes, { prefix: "/api/breakdowns" });
await app.register(stockRoutes, { prefix: "/api/stock" });
await app.ready();
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

ensureTechSchema(db);
// Columns other route groups add to breakdowns on a live server.
const bdCols = new Set(db.prepare("PRAGMA table_info(breakdowns)").all().map((c) => c.name));
for (const [name, type] of [["site_code", "TEXT"], ["parts_status", "TEXT"], ["critical", "INTEGER"], ["ets_repair_date", "TEXT"], ["primary_work_order_id", "INTEGER"]]) {
  if (!bdCols.has(name)) db.exec(`ALTER TABLE breakdowns ADD COLUMN ${name} ${type}`);
}
db.exec(`
  CREATE TABLE IF NOT EXISTS maintenance_parts_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT, site_code TEXT DEFAULT 'main', asset_id INTEGER, asset_code TEXT,
    part_code TEXT, part_name TEXT NOT NULL, qty REAL NOT NULL DEFAULT 1, urgency TEXT NOT NULL DEFAULT 'normal',
    notes TEXT, work_order_id INTEGER, status TEXT NOT NULL DEFAULT 'requested', requested_by TEXT NOT NULL,
    ordered_by TEXT, status_notes TEXT, created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
  );
  INSERT INTO assets (id, asset_code, asset_name) VALUES (1, 'A303AM', 'Komatsu dozer'), (2, 'LDV12', 'Hilux');
  INSERT INTO parts (id, part_code, part_name) VALUES (1, 'FLT-01', 'Oil filter'), (2, 'MLFPT6001', 'Shell Rimula 15W40 20L');
  INSERT INTO stock_movements (part_id, quantity, movement_type, reference) VALUES (1, 6, 'in', 'opening'), (2, 3, 'in', 'opening');
  INSERT INTO work_orders (id, asset_id, source, status, opened_at, job_description, assigned_artisan_name)
    VALUES (10, 1, 'manual', 'open', datetime('now'), 'Replace oil filter', 'ana'),
           (11, 2, 'manual', 'in_progress', datetime('now'), 'Brakes', 'joao'),
           (12, 2, 'manual', 'completed', datetime('now'), 'Done job', 'ana');
  INSERT INTO work_order_technicians (work_order_id, username) VALUES (11, 'pedro');
  INSERT INTO maintenance_parts_requests (id, asset_id, asset_code, part_code, part_name, qty, work_order_id, requested_by)
    VALUES (5, 1, 'A303AM', 'FLT-01', 'Oil filter', 2, 10, 'ana');
`);

const as = (user, roles) => ({ "x-user-name": user, "x-user-role": roles.split(",")[0], "x-user-roles": roles, "x-site-code": "main" });
const STORES = as("stores1", "storeman,stores");
const call = async (headers, method, url, payload) => {
  const res = await app.inject({ method, url: `/api/stock${url}`, headers, payload });
  return { code: res.statusCode, body: res.json() };
};
const onHand = (id) => db.prepare("SELECT SUM(quantity) AS q FROM stock_movements WHERE part_id = ?").get(id).q;

test("a scanned code is a part, a machine or a work order, links included", () => {
  assert.equal(lookupCode(db, "flt-01").kind, "part");
  assert.equal(lookupCode(db, "https://ironlog.example/web/store-mobile.html?site=main&part_code=MLFPT6001").part_code, "MLFPT6001");
  assert.equal(lookupCode(db, "https://ironlog.example/web/machine-prestart.html?asset_code=a303am").asset_code, "A303AM");
  assert.equal(lookupCode(db, "WO #10").kind, "work_order");
  assert.equal(lookupCode(db, "WO:11").id, 11);
  assert.equal(lookupCode(db, "nothing-here").kind, "unknown");
  assert.equal(lookupCode(db, "").kind, "none");
});

test("stores terminal: home, lookup, issue for storemen and technicians", async (t) => {
  t.after(() => app.close());

  const home = (await call(STORES, "GET", "/terminal/home")).body;
  assert.equal(home.stores, true);
  assert.equal(home.requests_waiting, 1);
  assert.equal(home.requests_in_stock, 1);
  assert.equal((await call(as("op1", "operator"), "GET", "/terminal/home")).code, 403);

  // Technicians see only their own open jobs (lead or helper).
  assert.deepEqual((await call(as("ana", "artisan"), "GET", "/terminal/work-orders")).body.rows.map((w) => w.id), [10]);
  assert.deepEqual((await call(as("pedro", "artisan"), "GET", "/terminal/work-orders")).body.rows.map((w) => w.id), [11]);
  assert.deepEqual((await call(STORES, "GET", "/terminal/work-orders")).body.rows.map((w) => w.id), [11, 10]);
  assert.deepEqual((await call(STORES, "GET", "/terminal/work-orders?q=a303")).body.rows.map((w) => w.id), [10]);

  const part = (await call(STORES, "GET", "/terminal/lookup?code=flt-01")).body;
  assert.equal(part.kind, "part");
  assert.equal(part.on_hand, 6);
  const wo = (await call(as("ana", "artisan"), "GET", "/terminal/lookup?code=WO%2011")).body;
  assert.equal(wo.kind, "work_order");
  assert.equal(wo.allowed, false, "not ana's job");

  const requests = (await call(STORES, "GET", "/terminal/requests")).body.rows;
  assert.equal(requests[0].id, 5);
  assert.equal((await call(as("ana", "artisan"), "GET", "/terminal/requests")).code, 403);

  // A technician cannot collect for someone else's job, or without a job.
  assert.equal((await call(as("ana", "artisan"), "POST", "/terminal/issue", { work_order_id: 11, lines: [{ part_code: "FLT-01", quantity: 1 }] })).code, 403);
  assert.equal((await call(as("ana", "artisan"), "POST", "/terminal/issue", { asset_code: "A303AM", lines: [{ part_code: "FLT-01", quantity: 1 }] })).code, 403);

  // Not enough stock on any line: nothing is issued.
  const short = await call(as("ana", "artisan"), "POST", "/terminal/issue", {
    work_order_id: 10, lines: [{ part_code: "FLT-01", quantity: 2 }, { part_code: "MLFPT6001", quantity: 4 }],
  });
  assert.equal(short.code, 400);
  assert.match(short.body.error, /MLFPT6001: only 3 in stock, 4 asked/);
  assert.equal(onHand(1), 6);

  // The technician collects against their job; the request is closed off.
  const ok = await call(as("ana", "artisan"), "POST", "/terminal/issue", {
    work_order_id: 10, lines: [{ part_code: "FLT-01", quantity: 2, request_id: 5 }, { part_code: "MLFPT6001", quantity: 1 }],
  });
  assert.equal(ok.code, 200, JSON.stringify(ok.body));
  assert.deepEqual(ok.body.lines.map((l) => [l.part_code, l.on_hand_after]), [["FLT-01", 4], ["MLFPT6001", 2]]);
  assert.equal(onHand(1), 4);
  const alloc = db.prepare("SELECT asset_id, work_order_id, issued_by, quantity FROM store_allocations WHERE part_id = 1").get();
  assert.deepEqual({ ...alloc }, { asset_id: 1, work_order_id: 10, issued_by: "ana", quantity: 2 });
  assert.equal(db.prepare("SELECT reference FROM stock_movements WHERE part_id = 1 AND quantity < 0").get().reference, "work_order:10");
  assert.equal(db.prepare("SELECT status FROM maintenance_parts_requests WHERE id = 5").get().status, "received");

  // A helper may collect for the job they help on; stores may issue to a machine.
  assert.equal((await call(as("pedro", "artisan"), "POST", "/terminal/issue", { work_order_id: 11, lines: [{ part_code: "FLT-01", quantity: 1 }] })).code, 200);
  const toMachine = await call(STORES, "POST", "/terminal/issue", { asset_code: "ldv12", lines: [{ part_code: "FLT-01", quantity: 1 }] });
  assert.equal(toMachine.code, 200);
  assert.equal(toMachine.body.asset_code, "LDV12");
  assert.equal(onHand(1), 2);
});

test("a paired phone relays scans to the terminal by its key", async () => {
  const app2 = Fastify({ logger: false });
  await app2.register(stockRoutes, { prefix: "/api/stock" });
  await app2.ready();
  try {
    const key = "0123456789abcdef0123456789abcdef";
    assert.equal(isPublicAuthRequest("/api/stock/terminal/scan", "POST"), true);
    assert.equal(isPublicAuthRequest("/api/stock/terminal/scans", "GET"), true);
    assert.equal(isPublicAuthRequest("/api/stock/terminal/issue", "POST"), false);

    const bad = await app2.inject({ method: "POST", url: "/api/stock/terminal/scan", payload: { key: "short", code: "FLT-01" } });
    assert.equal(bad.statusCode, 400);
    const first = (await app2.inject({ method: "GET", url: `/api/stock/terminal/scans?key=${key}` })).json();
    assert.deepEqual(first.rows, []);
    await app2.inject({ method: "POST", url: "/api/stock/terminal/scan", payload: { key, code: "FLT-01" } });
    await app2.inject({ method: "POST", url: "/api/stock/terminal/scan", payload: { key, code: "A303AM" } });
    const got = (await app2.inject({ method: "GET", url: `/api/stock/terminal/scans?key=${key}&after=${first.last}` })).json();
    assert.deepEqual(got.rows.map((r) => r.code), ["FLT-01", "A303AM"]);
    const none = (await app2.inject({ method: "GET", url: `/api/stock/terminal/scans?key=${key}&after=${got.last}` })).json();
    assert.deepEqual(none.rows, []);
    const other = (await app2.inject({ method: "GET", url: `/api/stock/terminal/scans?key=${"f".repeat(32)}` })).json();
    assert.deepEqual(other.rows, [], "another terminal sees nothing");
  } finally {
    await app2.close();
  }
});

test("a supplier's box barcode is linked to a part once and then found by a scan", async () => {
  const app3 = Fastify({ logger: false });
  await app3.register(stockRoutes, { prefix: "/api/stock" });
  await app3.ready();
  const req = async (headers, method, url, payload) => {
    const res = await app3.inject({ method, url: `/api/stock${url}`, headers, payload });
    return { code: res.statusCode, body: res.json() };
  };
  try {
    assert.equal((await req(STORES, "GET", "/terminal/lookup?code=6001234567890")).body.kind, "unknown");
    // Technicians cannot link; a part or machine code cannot be used as a barcode.
    assert.equal((await req(as("ana", "artisan"), "POST", "/terminal/barcodes", { barcode: "6001234567890", part_code: "FLT-01" })).code, 403);
    assert.match((await req(STORES, "POST", "/terminal/barcodes", { barcode: "a303am", part_code: "FLT-01" })).body.error, /already a part or machine code/);

    const linked = await req(STORES, "POST", "/terminal/barcodes", { barcode: " 6001234567890 ", part_code: "flt-01" });
    assert.equal(linked.code, 200, JSON.stringify(linked.body));
    assert.equal(linked.body.part_code, "FLT-01");
    const hit = (await req(as("ana", "artisan"), "GET", "/terminal/lookup?code=6001234567890")).body;
    assert.equal(hit.kind, "part");
    assert.equal(hit.part_code, "FLT-01");
    assert.deepEqual(hit.barcodes, ["6001234567890"]);
    assert.equal((await req(STORES, "GET", "/parts/search?q=6001234567890")).body.rows[0].part_code, "FLT-01");

    // Linked to another part: refused until replace is confirmed.
    const clash = await req(STORES, "POST", "/terminal/barcodes", { barcode: "6001234567890", part_code: "MLFPT6001" });
    assert.equal(clash.code, 409);
    assert.equal(clash.body.linked_to.part_code, "FLT-01");
    const moved = await req(STORES, "POST", "/terminal/barcodes", { barcode: "6001234567890", part_code: "MLFPT6001", replace: true });
    assert.equal(moved.body.moved_from, "FLT-01");
    assert.equal((await req(STORES, "GET", "/terminal/lookup?code=6001234567890")).body.part_code, "MLFPT6001");

    assert.equal((await req(STORES, "DELETE", "/terminal/barcodes/6001234567890")).code, 200);
    assert.equal((await req(STORES, "GET", "/terminal/lookup?code=6001234567890")).body.kind, "unknown");
    assert.equal((await req(STORES, "DELETE", "/terminal/barcodes/6001234567890")).code, 404);
  } finally {
    await app3.close();
  }
});
