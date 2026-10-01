import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-requisition-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { default: stockRoutes } = await import("../routes/stock.routes.js");
const { default: maintenanceRoutes } = await import("../routes/maintenance.routes.js");
const app = Fastify({ logger: false });
await app.register(maintenanceRoutes, { prefix: "/api/maintenance" });
await app.register(stockRoutes, { prefix: "/api/stock" });
await app.ready();
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

db.exec(`
  INSERT INTO assets (id, asset_code, asset_name, category, active) VALUES (1, 'T01AM', 'Tipper 01', 'tipper_truck', 1);
  INSERT INTO work_orders (id, asset_id, source, status, opened_at) VALUES (7, 1, 'manual', 'in_progress', datetime('now'));
  INSERT INTO parts (id, part_code, part_name) VALUES (1, 'HOSE-12', 'Hydraulic hose 1/2in'), (2, 'FLT-01', 'Oil filter');
  INSERT INTO stock_movements (part_id, quantity, movement_type, reference) VALUES (1, 4, 'in', 'GRN 1'), (2, 6, 'in', 'GRN 2');
`);

const as = (role, user = "stores1") => ({ "x-user-name": user, "x-user-role": role, "x-user-roles": role, "x-site-code": "main" });
const call = async (method, url, role, payload) => {
  const res = await app.inject({ method, url, headers: as(role), payload });
  return { code: res.statusCode, body: res.headers["content-type"]?.includes("pdf") ? res.rawPayload : res.json(), type: res.headers["content-type"] };
};

test("requisition from the stores queue, reprint keeps the number, walk-up form, PDF", async (t) => {
  t.after(() => app.close());

  // Two lines the workshop asked for on WO #7.
  for (const line of [{ part_code: "HOSE-12", qty: 1 }, { part_code: "FLT-01", qty: 2, urgency: "urgent" }]) {
    const r = await app.inject({ method: "POST", url: "/api/maintenance/parts-requests", headers: as("artisan", "jose"), payload: { ...line, work_order_id: 7, asset_id: 1 } });
    assert.equal(r.statusCode, 200, r.body);
  }

  assert.equal((await call("POST", "/api/stock/requisitions", "artisan", { work_order_id: 7 })).code, 403, "stores only");
  const first = await call("POST", "/api/stock/requisitions", "storeman", { work_order_id: 7 });
  assert.equal(first.code, 200, JSON.stringify(first.body));
  assert.equal(first.body.created, true);
  const req = first.body.requisition;
  assert.equal(req.number, `SR-${String(first.body.id).padStart(6, "0")}`);
  assert.equal(req.work_order.id, 7);
  assert.equal(req.asset.asset_code, "T01AM");
  assert.equal(req.requested_by, "jose");
  assert.deepEqual(req.lines.map((l) => [l.part_code, l.qty, l.on_hand]), [["HOSE-12", 1, 4], ["FLT-01", 2, 6]]);

  const again = await call("POST", "/api/stock/requisitions", "storeman", { work_order_id: 7 });
  assert.equal(again.body.created, false);
  assert.equal(again.body.id, first.body.id, "reprinting keeps the same requisition number");
  assert.equal((await call("POST", "/api/stock/requisitions", "storeman", { work_order_id: 999 })).code, 404);

  // Someone asks at the counter.
  assert.equal((await call("POST", "/api/stock/requisitions/walk-up", "storeman", { lines: [{ part_code: "FLT-01", qty: 1 }] })).code, 400, "needs a name");
  assert.equal((await call("POST", "/api/stock/requisitions/walk-up", "storeman", { requested_by: "Carlos", lines: [] })).code, 400, "needs a line");
  assert.equal((await call("POST", "/api/stock/requisitions/walk-up", "storeman", { requested_by: "Carlos", asset_code: "NOPE", lines: [{ part_code: "FLT-01", qty: 1 }] })).code, 404);
  const walk = await call("POST", "/api/stock/requisitions/walk-up", "storeman", {
    requested_by: "Carlos (operator)",
    asset_code: "t01am",
    notes: "Daily service",
    lines: [{ part_name: "flt-01", qty: 1 }, { part_name: "Cable ties", qty: 20 }],
  });
  assert.equal(walk.code, 200, JSON.stringify(walk.body));
  const w = walk.body.requisition;
  assert.equal(w.walk_up, 1);
  assert.equal(w.requested_by, "Carlos (operator)");
  assert.equal(w.asset.asset_code, "T01AM");
  assert.deepEqual(w.lines.map((l) => [l.part_code || l.part_name, l.qty]), [["FLT-01", 1], ["Cable ties", 20]], "a typed stock code is matched to the stock item");
  assert.equal(w.lines[0].description, "Oil filter");
  assert.equal(w.lines[0].on_hand, 6);
  // The walk-up lines are ordinary request lines, so stores see them in the queue too.
  assert.equal(db.prepare("SELECT COUNT(*) AS n FROM maintenance_parts_requests WHERE requisition_id = ?").get(walk.body.id).n, 2);
  assert.match(db.prepare("SELECT notes FROM maintenance_parts_requests WHERE requisition_id = ? LIMIT 1").get(walk.body.id).notes, /Asked at stores by Carlos/);

  const pdf = await call("GET", `/api/stock/requisitions/${first.body.id}.pdf`, "storeman");
  assert.equal(pdf.code, 200);
  assert.match(pdf.type, /application\/pdf/);
  assert.equal(pdf.body.subarray(0, 4).toString(), "%PDF");
  assert.equal((await call("GET", "/api/stock/requisitions/9999.pdf", "storeman")).code, 404);

  const list = (await call("GET", "/api/stock/requisitions", "storeman")).body.rows;
  assert.deepEqual(list.map((r) => [r.id, r.line_count]), [[walk.body.id, 2], [first.body.id, 2]]);
  assert.equal(list[1].printed_count, 1);
});
