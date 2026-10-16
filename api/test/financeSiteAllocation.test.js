import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-finance-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;

await import("../db/migrate.js");
(await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
const { db } = await import("../db/client.js");
const realSetInterval = globalThis.setInterval;
globalThis.setInterval = (...args) => realSetInterval(...args).unref();
const app = Fastify({ logger: false });
// Stock adds the part cost columns the allocation reads, as at start-up.
await app.register((await import("../routes/stock.routes.js")).default, { prefix: "/api/stock" });
await app.register((await import("../routes/dashboard.routes.js")).default, { prefix: "/api/dashboard" });
await app.register((await import("../routes/finance.routes.js")).default, { prefix: "/api/finance" });
await app.ready();
globalThis.setInterval = realSetInterval;
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

test("the finance site allocation opens for a month (it used to fail on the downtime rate)", async () => {
  const res = await app.inject({ method: "GET", url: "/api/finance/site-allocation?period=2026-10", headers: { "x-user-role": "admin", "x-user-roles": "admin", "x-site-code": "main" } });
  assert.equal(res.statusCode, 200, res.body);
});
