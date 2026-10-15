import test from "node:test";
import assert from "node:assert/strict";
import crypto from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-licence-routes-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
const { privateKey, publicKey } = crypto.generateKeyPairSync("ed25519");
process.env.IRONLOG_LICENCE_PUBLIC_KEYS = publicKey.export({ format: "der", type: "spki" }).toString("base64");
const Fastify = (await import("fastify")).default;
const { db } = await import("../db/client.js");
const { signLicence } = await import("../utils/licensing.js");
const licenceRoutes = (await import("../routes/licence.routes.js")).default;
const { licenceHook, clearLicenceCache } = await import("../routes/licence.routes.js");
process.on("exit", () => { try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

db.exec(`CREATE TABLE IF NOT EXISTS assets (id INTEGER PRIMARY KEY, asset_code TEXT, site_code TEXT, active INTEGER DEFAULT 1, archived INTEGER DEFAULT 0);
         INSERT INTO assets (asset_code, site_code) VALUES ('E501AM', 'main');`);

test("installing a key is for administrators; the hook then switches modules off on the server", async (t) => {
  const app = Fastify();
  app.addHook("onRequest", licenceHook);
  await app.register(licenceRoutes, { prefix: "/api/licence" });
  app.get("/api/dashboard/fuel/daily", async () => ({ ok: true }));
  app.post("/api/workorders", async () => ({ ok: true }));
  t.after(() => app.close());
  const as = (role) => ({ "x-user-roles": role, "x-user-name": role, "x-site-code": "main" });

  const before = (await app.inject({ method: "GET", url: "/api/licence/state", headers: as("artisan") })).json();
  assert.equal(before.mode, "legacy", "an existing install starts fully unlocked");
  assert.equal((await app.inject({ method: "GET", url: "/api/dashboard/fuel/daily", headers: as("artisan") })).statusCode, 200);

  const key = signLicence({ licence_id: "L1", licensee: "AML", expires_at: "2099-12-31", sites: { main: { package: "Core", modules: ["stores"] } } }, privateKey);
  assert.equal((await app.inject({ method: "POST", url: "/api/licence/install", headers: as("plant_manager"), payload: { key } })).statusCode, 403);
  const bad = await app.inject({ method: "POST", url: "/api/licence/install", headers: as("admin"), payload: { key: key.slice(0, -4) + "AAAA" } });
  assert.equal(bad.statusCode, 400);
  const ok = await app.inject({ method: "POST", url: "/api/licence/install", headers: as("admin"), payload: { key } });
  assert.equal(ok.statusCode, 200);
  assert.equal(ok.json().state.mode, "licensed");

  const fuel = await app.inject({ method: "GET", url: "/api/dashboard/fuel/daily", headers: as("admin") });
  assert.equal(fuel.statusCode, 403);
  assert.match(fuel.json().error, /not included in this site's IronLog licence/);
  assert.equal((await app.inject({ method: "POST", url: "/api/workorders", headers: as("admin") })).statusCode, 200, "core keeps working");

  const csv = await app.inject({ method: "GET", url: "/api/licence/usage.csv", headers: as("admin") });
  assert.match(csv.body, /main,Core,1,,/);
  assert.equal((await app.inject({ method: "GET", url: "/api/licence/usage", headers: as("storeman") })).statusCode, 403);

  // An expired licence past its grace period: read-only.
  const expired = signLicence({ licence_id: "L2", licensee: "AML", expires_at: "2020-01-01", sites: { "*": { modules: ["fuel"] } } }, privateKey);
  await app.inject({ method: "POST", url: "/api/licence/install", headers: as("admin"), payload: { key: expired } });
  clearLicenceCache();
  const save = await app.inject({ method: "POST", url: "/api/workorders", headers: as("admin") });
  assert.equal(save.statusCode, 423);
  assert.match(save.json().error, /read-only/);
  assert.equal((await app.inject({ method: "GET", url: "/api/dashboard/fuel/daily", headers: as("admin") })).statusCode, 200, "reading still works");
  const audit = db.prepare(`SELECT COUNT(*) AS n FROM audit_logs WHERE module = 'licence' AND action = 'install'`).get().n;
  assert.equal(audit, 2);
});
