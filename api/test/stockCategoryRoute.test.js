import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";
import { ensureStockCategorySchema } from "../utils/stockCategory.js";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-stock-category-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { default: stockRoutes } = await import("../routes/stock.routes.js");
const app = Fastify({ logger: false });
await app.register(stockRoutes, { prefix: "/api/stock" });
await app.ready();
db.exec(`INSERT INTO parts (part_code, part_name, critical, min_stock, unit_cost) VALUES
  ('MLFPT6001', 'Fuchs Truck Plus 15W40', 0, 0, 4.63),
  ('224363', 'Coil directional vale B30', 0, 0, 76.19)`);
// Items created outside the stores screens get their automatic category at the next start-up.
ensureStockCategorySchema(db, { refresh: true });

const as = (role) => ({ "x-user-name": "tester", "x-user-role": role, "x-user-roles": role });

test("stock monitor carries each item's category and stores can correct it", async (t) => {
  t.after(async () => {
    await app.close();
    db.close();
    rmSync(tempDir, { recursive: true, force: true });
  });

  const monitor = async () => JSON.parse((await app.inject({ method: "GET", url: "/api/stock/monitor", headers: as("stores") })).body).rows;
  const byCode = (rows, code) => rows.find((r) => r.part_code === code);
  let rows = await monitor();
  assert.equal(byCode(rows, "MLFPT6001").stock_category, "oil");
  assert.equal(byCode(rows, "MLFPT6001").stock_category_label, "Oils & lubricants");
  assert.equal(byCode(rows, "224363").stock_category, "part");

  const denied = await app.inject({ method: "POST", url: "/api/stock/part-category", headers: as("operator"), payload: { part_code: "224363", category: "component" } });
  assert.equal(denied.statusCode, 403);

  const bad = await app.inject({ method: "POST", url: "/api/stock/part-category", headers: as("stores"), payload: { part_code: "224363", category: "widgets" } });
  assert.equal(bad.statusCode, 400);

  const ok = await app.inject({ method: "POST", url: "/api/stock/part-category", headers: as("stores"), payload: { part_code: "224363", category: "component" } });
  assert.equal(ok.statusCode, 200, ok.body);
  assert.equal(JSON.parse(ok.body).stock_category_source, "manual");
  rows = await monitor();
  assert.equal(byCode(rows, "224363").stock_category, "component");

  const reset = await app.inject({ method: "POST", url: "/api/stock/part-category", headers: as("stores"), payload: { part_code: "224363", category: "auto" } });
  assert.equal(JSON.parse(reset.body).stock_category, "part");

  const cats = JSON.parse((await app.inject({ method: "GET", url: "/api/stock/categories", headers: as("stores") })).body);
  assert.deepEqual(cats.categories.map((c) => c.key), ["oil", "get", "part", "component", "tyre"]);
});
