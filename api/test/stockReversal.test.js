import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-stock-reversal-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { default: stockRoutes } = await import("../routes/stock.routes.js");
const { default: approvalsRoutes } = await import("../routes/approvals.routes.js");
const app = Fastify({ logger: false });
await app.register(stockRoutes, { prefix: "/api/stock" });
await app.register(approvalsRoutes, { prefix: "/api/approvals" });
await app.ready();

db.exec(`
  INSERT INTO parts (id, part_code, part_name) VALUES (1, 'MLFPT6001', 'Fuchs Truck Plus 15W40 20L'), (2, 'HYD68', 'Hydraulic oil 68');
  INSERT INTO stock_movements (id, part_id, quantity, movement_type, reference, created_at) VALUES
    (1, 1, 10, 'in', 'GRN 55', datetime('now', '-2 hours')),
    (2, 1, 10, 'in', 'GRN 55', datetime('now', '-2 hours', '+3 minutes')),
    (3, 2, 4, 'in', 'GRN 60', datetime('now', '-1 hours')),
    (4, 2, -3, 'out', 'work_order:9', datetime('now', '-30 minutes'));
`);

const as = (role, user = "store") => ({ "x-user-name": user, "x-user-role": role, "x-user-roles": role });
const call = async (method, url, role, payload) => {
  const res = await app.inject({ method, url, headers: as(role), payload });
  return { code: res.statusCode, body: res.json() };
};
const onHand = (id) => db.prepare("SELECT SUM(quantity) AS q FROM stock_movements WHERE part_id = ?").get(id).q;

test("a receipt captured twice is flagged, reversed after approval, and only once", async (t) => {
  t.after(async () => {
    await app.close();
    db.close();
    rmSync(tempDir, { recursive: true, force: true });
  });

  const list = (await call("GET", "/api/stock/receipts?part_code=MLFPT", "storeman")).body.rows;
  assert.deepEqual(list.map((r) => [r.id, r.possible_duplicate_of]), [[2, 1], [1, null]], "second GRN 55 flagged as a repeat of the first");
  assert.equal((await call("GET", "/api/stock/receipts", "artisan")).code, 403);

  assert.equal((await call("POST", "/api/stock/movements/2/reverse", "storeman", {})).code, 400, "needs a reason");
  assert.equal((await call("POST", "/api/stock/movements/4/reverse", "storeman", { reason: "x" })).code, 400, "only receipts");
  assert.equal((await call("POST", "/api/stock/movements/99/reverse", "storeman", { reason: "x" })).code, 404);
  assert.equal((await call("POST", "/api/stock/movements/2/reverse", "operator", { reason: "x" })).code, 403);

  const req = await call("POST", "/api/stock/movements/2/reverse", "storeman", { reason: "Delivery captured twice" });
  assert.equal(req.code, 200, JSON.stringify(req.body));
  assert.equal(req.body.pending_approval, true);
  assert.equal(onHand(1), 20, "nothing changes until approved");
  assert.equal((await call("POST", "/api/stock/movements/2/reverse", "storeman", { reason: "again" })).code, 409, "already waiting");
  assert.equal((await call("GET", "/api/stock/receipts?part_code=MLFPT", "storeman")).body.rows[0].pending_request_id, req.body.request_id);

  assert.equal((await call("POST", `/api/approvals/${req.body.request_id}/approve`, "storeman", {})).code, 403, "stores cannot approve their own reversal");
  const ok = await call("POST", `/api/approvals/${req.body.request_id}/approve`, "admin", {});
  assert.equal(ok.code, 200, JSON.stringify(ok.body));
  assert.equal(ok.body.execution.on_hand_before, 20);
  assert.equal(ok.body.execution.on_hand_after, 10);
  assert.equal(onHand(1), 10);
  const rev = db.prepare("SELECT quantity, movement_type, reference FROM stock_movements WHERE id = ?").get(ok.body.execution.reversal_movement_id);
  assert.deepEqual(rev, { quantity: -10, movement_type: "adjust", reference: "reversal_of:2 GRN 55" });

  const after = (await call("GET", "/api/stock/receipts?part_code=MLFPT", "storeman")).body.rows;
  assert.equal(after[0].reversed_by, ok.body.execution.reversal_movement_id);
  assert.equal(after[1].possible_duplicate_of, null);
  assert.equal((await call("POST", "/api/stock/movements/2/reverse", "storeman", { reason: "again" })).code, 409, "never reversed twice");

  // Stock already issued: approving would go below zero, so it is refused and stays pending.
  const hyd = await call("POST", "/api/stock/movements/3/reverse", "storeman", { reason: "wrong item" });
  assert.equal(hyd.code, 200);
  const refused = await call("POST", `/api/approvals/${hyd.body.request_id}/approve`, "admin", {});
  assert.equal(refused.code, 409);
  assert.match(refused.body.error, /Only 1 of HYD68 left/);
  assert.equal(onHand(2), 1);
  assert.equal(db.prepare("SELECT status FROM approval_requests WHERE id = ?").get(hyd.body.request_id).status, "pending");
});
