import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-quality-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;

await import("../db/migrate.js");
(await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
const { db } = await import("../db/client.js");
const realSetInterval = globalThis.setInterval;
globalThis.setInterval = (...args) => realSetInterval(...args).unref();
const app = Fastify({ logger: false });
// The tables the Quality screen reads are made by these areas at start-up.
for (const [file, prefix] of [["hours.routes.js", "/api/hours"], ["workorders.routes.js", "/api/workorders"], ["stock.routes.js", "/api/stock"], ["operations.routes.js", "/api/operations"], ["dispatch.routes.js", "/api/dispatch"], ["quality.routes.js", "/api/quality"]]) {
  await app.register((await import(`../routes/${file}`)).default, { prefix });
}
await app.ready();
globalThis.setInterval = realSetInterval;
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

test("the Quality screen opens and lists pending approvals with who asked", async () => {
  db.prepare(`INSERT INTO approval_requests (module, action, status, requested_by) VALUES ('stock', 'adjust', 'pending', 'joao')`).run();
  const res = await app.inject({ method: "GET", url: "/api/quality?from=2026-09-16&to=2026-10-16", headers: { "x-user-role": "admin", "x-user-roles": "admin", "x-site-code": "main" } });
  assert.equal(res.statusCode, 200, res.body);
  const pending = JSON.stringify(res.json());
  assert.match(pending, /"requester":"joao"/);
});
