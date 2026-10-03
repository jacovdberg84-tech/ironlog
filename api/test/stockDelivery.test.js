import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-delivery-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { default: stockRoutes } = await import("../routes/stock.routes.js");
const app = Fastify({ logger: false });
await app.register(stockRoutes, { prefix: "/api/stock" });
await app.ready();
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

db.exec(`
  INSERT INTO parts (id, part_code, part_name) VALUES (1, 'MLFPT6001', 'Shell Rimula 15W40 20L'), (2, 'FLT-01', 'Oil filter');
  INSERT INTO stock_movements (part_id, quantity, movement_type, reference) VALUES (1, 12, 'in', 'opening'), (2, 6, 'in', 'opening');
`);
const loc = db.prepare("SELECT location_code FROM stock_locations WHERE UPPER(location_code) = 'MAIN'").get();
if (!loc) db.exec(`INSERT INTO stock_locations (location_code, location_name, active) VALUES ('MAIN', 'Main Store', 1)`);

const as = (role) => ({ "x-user-name": "stores1", "x-user-role": role, "x-user-roles": role, "x-site-code": "main" });
const post = async (payload, role = "storeman") => {
  const res = await app.inject({ method: "POST", url: "/api/stock/deliveries", headers: as(role), payload });
  return { code: res.statusCode, body: res.json() };
};
const onHand = (id) => db.prepare("SELECT SUM(quantity) AS q FROM stock_movements WHERE part_id = ?").get(id).q;

test("receive a whole delivery, all lines or none, with duplicate warnings", async (t) => {
  t.after(() => app.close());

  const search = (await app.inject({ method: "GET", url: "/api/stock/parts/search?q=rimula", headers: as("storeman") })).json();
  assert.equal(search.rows[0].part_code, "MLFPT6001");
  assert.equal(search.rows[0].on_hand, 12);

  assert.equal((await post({ reference: "INV-1", lines: [] }, "artisan")).code, 403);
  assert.match((await post({ lines: [{ part_code: "FLT-01", quantity: 1 }] })).body.error, /invoice or GRN/);

  // One bad line stops the whole delivery: nothing is saved.
  const bad = await post({ reference: "INV-2231", lines: [{ part_code: "MLFPT6001", quantity: 4 }, { part_code: "NEW-HOSE", quantity: 2 }] });
  assert.equal(bad.code, 400);
  assert.match(bad.body.error, /NEW-HOSE is a new part, add its description/);
  assert.equal(onHand(1), 12, "nothing received");

  const good = await post({
    supplier: "Lubemoz", reference: "INV-2231", currency: "MZN", received_date: "2026-10-03",
    lines: [
      { part_code: "MLFPT6001", quantity: 4, unit_cost: 2464 },
      { part_code: "FLT-01", quantity: 10 },
      { part_code: "new-hose", part_name: "Hydraulic hose 1/2in", quantity: 2, unit_cost: 640 },
      { part_code: "", quantity: "" },
    ],
  });
  assert.equal(good.code, 200, JSON.stringify(good.body));
  assert.equal(good.body.line_count, 3);
  assert.deepEqual(good.body.lines.map((l) => [l.part_code, l.on_hand_after, l.new_part]), [["MLFPT6001", 16, false], ["FLT-01", 16, false], ["NEW-HOSE", 2, true]]);
  assert.ok(good.body.lines[0].unit_cost_usd > 0, "MZN converted to USD");
  assert.equal(db.prepare("SELECT COUNT(*) AS n FROM stock_movements WHERE delivery_id = ?").get(good.body.delivery_id).n, 3);
  assert.equal(db.prepare("SELECT reference FROM stock_movements WHERE delivery_id = ? LIMIT 1").get(good.body.delivery_id).reference, "INV-2231");

  // The same invoice again: warned, not saved, until confirmed.
  const again = await post({ reference: "inv-2231", lines: [{ part_code: "MLFPT6001", quantity: 4 }] });
  assert.equal(again.code, 409);
  assert.equal(again.body.needs_confirmation, true);
  assert.deepEqual(again.body.duplicates.map((d) => d.kind), ["invoice", "line"]);
  assert.equal(onHand(1), 16);
  const confirmed = await post({ reference: "inv-2231", confirm_duplicates: true, lines: [{ part_code: "MLFPT6001", quantity: 4 }] });
  assert.equal(confirmed.code, 200);
  assert.equal(onHand(1), 20);

  // The same part on two lines is refused; a missing bin is named.
  assert.match((await post({ reference: "INV-9", lines: [{ part_code: "FLT-01", quantity: 1 }, { part_code: "flt-01", quantity: 2 }] })).body.error, /more than one line/);
  assert.match((await post({ reference: "INV-9", lines: [{ part_code: "FLT-01", quantity: 1, bin_code: "Z99" }] })).body.error, /bin Z99 not found/);

  const list = (await app.inject({ method: "GET", url: "/api/stock/deliveries", headers: as("storeman") })).json();
  assert.equal(list.rows.length, 2);
  assert.equal(list.rows[1].supplier, "Lubemoz");
  assert.equal(list.rows[1].lines.length, 3);
});
