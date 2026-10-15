import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-licence-issuer-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_SECRETS_DIR = path.join(tempDir, "secrets");
delete process.env.IRONLOG_LICENCE_PUBLIC_KEYS;
delete process.env.IRONLOG_LICENCE_ISSUERS;
const Fastify = (await import("fastify")).default;
const { db } = await import("../db/client.js");
const { verifyLicence } = await import("../utils/licensing.js");
const licenceRoutes = (await import("../routes/licence.routes.js")).default;
const { licenceHook } = await import("../routes/licence.routes.js");
process.on("exit", () => { try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

db.exec(`CREATE TABLE IF NOT EXISTS assets (id INTEGER PRIMARY KEY, asset_code TEXT, site_code TEXT, active INTEGER DEFAULT 1, archived INTEGER DEFAULT 0);
         INSERT INTO assets (asset_code, site_code) VALUES ('E501AM', 'main');`);

const owner = { "x-user-roles": "admin", "x-user-name": "bj.vandenberg", "x-site-code": "main" };

test("only the IronLog owner, signed in as an admin, can use the issuer", async (t) => {
  const app = Fastify();
  await app.register(licenceRoutes, { prefix: "/api/licence" });
  t.after(() => app.close());
  const otherAdmin = { "x-user-roles": "admin", "x-user-name": "aml.admin" };
  const ownerNotAdmin = { "x-user-roles": "plant_manager", "x-user-name": "bj.vandenberg" };
  assert.equal((await app.inject({ method: "GET", url: "/api/licence/issuer/status", headers: otherAdmin })).json().allowed, false);
  for (const h of [otherAdmin, ownerNotAdmin]) {
    assert.equal((await app.inject({ method: "POST", url: "/api/licence/issuer/setup", headers: h })).statusCode, 403);
    assert.equal((await app.inject({ method: "GET", url: "/api/licence/issuer/catalogue", headers: h })).statusCode, 403);
  }
  assert.equal((await app.inject({ method: "GET", url: "/api/licence/issuer/status", headers: { ...owner, "x-user-name": "BJ.VanDenBerg" } })).json().allowed, true, "username case does not matter");
});

test("the owner sets up a signing key once, edits packages and prices, and issues a key this server accepts", async (t) => {
  const app = Fastify();
  app.addHook("onRequest", licenceHook);
  await app.register(licenceRoutes, { prefix: "/api/licence" });
  t.after(() => app.close());

  let st = (await app.inject({ method: "GET", url: "/api/licence/issuer/status", headers: owner })).json();
  assert.deepEqual([st.allowed, st.has_signing_key], [true, false]);
  assert.equal((await app.inject({ method: "POST", url: "/api/licence/issuer/issue", headers: owner, payload: {} })).statusCode, 400, "no key yet");
  const setup = await app.inject({ method: "POST", url: "/api/licence/issuer/setup", headers: owner });
  assert.equal(setup.statusCode, 200);
  const pub = setup.json().public_key;
  assert.equal((await app.inject({ method: "POST", url: "/api/licence/issuer/setup", headers: owner })).statusCode, 409, "never replaced");
  const file = path.join(tempDir, "secrets", "licence-signing.pem");
  assert.equal(fs.statSync(file).mode & 0o777, 0o600, "only the server account can read it");
  st = (await app.inject({ method: "GET", url: "/api/licence/issuer/status", headers: owner })).json();
  assert.deepEqual([st.has_signing_key, st.public_key, st.built_in], [true, pub, false]);

  let cat = (await app.inject({ method: "GET", url: "/api/licence/issuer/catalogue", headers: owner })).json();
  assert.deepEqual(cat.packages.map((p) => p.name), ["Core", "Professional", "Enterprise"]);
  assert.equal(cat.bands.length, 5);
  assert.equal((await app.inject({ method: "POST", url: "/api/licence/issuer/bands", headers: owner, payload: { label: "Tiny", max_assets: "ten" } })).statusCode, 400);
  await app.inject({ method: "POST", url: "/api/licence/issuer/packages", headers: owner, payload: { name: "Stores only", modules: ["stores", "nonsense"] } });
  cat = (await app.inject({ method: "GET", url: "/api/licence/issuer/catalogue", headers: owner })).json();
  const storesOnly = cat.packages.find((p) => p.name === "Stores only");
  assert.deepEqual(storesOnly.modules, ["stores"]);
  const band150 = cat.bands.find((b) => b.max_assets === 150);
  const unlimited = cat.bands.find((b) => b.max_assets == null);
  const ent = cat.packages.find((p) => p.name === "Enterprise");
  assert.equal((await app.inject({ method: "PUT", url: "/api/licence/issuer/prices", headers: owner, payload: { prices: [{ package_id: storesOnly.id, band_id: band150.id, monthly_price: "-5" }] } })).statusCode, 400);
  await app.inject({ method: "PUT", url: "/api/licence/issuer/prices", headers: owner, payload: { prices: [
    { package_id: storesOnly.id, band_id: band150.id, monthly_price: "450", currency: "usd" },
    { package_id: ent.id, band_id: unlimited.id, monthly_price: "2400" },
  ] } });

  const bad = await app.inject({ method: "POST", url: "/api/licence/issuer/issue", headers: owner, payload: { licensee: "Pilot", expires_at: "2020-01-01", sites: [] } });
  assert.match(bad.json().error, /end date after today/);
  const issued = await app.inject({ method: "POST", url: "/api/licence/issuer/issue", headers: owner, payload: {
    licensee: "Pilot Mining Lda",
    expires_at: "2099-06-30",
    sites: [
      { site_code: "main", name: "Main pit", package_id: storesOnly.id, band_id: band150.id, extra_modules: ["fuel"] },
      { site_code: "*", package_id: ent.id, band_id: unlimited.id },
    ],
  } });
  assert.equal(issued.statusCode, 200, issued.body);
  const r = issued.json();
  assert.match(r.licence_id, /^IL-\d{4}-0001-[0-9A-F]{4}$/);
  assert.deepEqual(r.payload.sites.main, { name: "Main pit", package: "Stores only", asset_limit: 150, modules: ["fuel", "stores"] });
  assert.equal(r.payload.sites["*"].asset_limit, null);
  assert.deepEqual(r.quote.total, { monthly: 2850, currency: "USD" });
  assert.equal(verifyLicence(r.key, [pub]).ok, true);

  // The issuing server accepts its own key straight away.
  const inst = await app.inject({ method: "POST", url: "/api/licence/install", headers: owner, payload: { key: r.key } });
  assert.equal(inst.statusCode, 200, inst.body);
  assert.equal(inst.json().state.licensee, "Pilot Mining Lda");
  assert.deepEqual(inst.json().state.modules, ["fuel", "stores"]);

  const list = (await app.inject({ method: "GET", url: "/api/licence/issuer/catalogue", headers: owner })).json().issued;
  assert.equal(list[0].licence_id, r.licence_id);
  assert.equal("licence_key" in list[0], false, "keys are fetched one at a time");
  assert.equal((await app.inject({ method: "GET", url: `/api/licence/issuer/issued/${list[0].id}/key`, headers: owner })).json().licence_key, r.key);
  const pem = await app.inject({ method: "GET", url: "/api/licence/issuer/signing-key.pem", headers: owner });
  assert.match(pem.body, /BEGIN PRIVATE KEY/);
  assert.equal((await app.inject({ method: "GET", url: "/api/licence/issuer/signing-key.pem", headers: { ...owner, "x-user-name": "aml.admin" } })).statusCode, 403);
  const actions = db.prepare(`SELECT action FROM audit_logs WHERE module = 'licence_issuer' ORDER BY id`).all().map((a) => a.action);
  assert.deepEqual(actions, ["signing_key_created", "package_save", "prices_save", "issue", "signing_key_downloaded"]);
});
