import test from "node:test";
import assert from "node:assert/strict";
import crypto from "node:crypto";
import Database from "better-sqlite3";

process.env.JUNE_CONNECTOR_ENCRYPTION_SECRET = "test-connector-secret-for-microsoft-sign-in-0123456789";
const ms = await import("../utils/microsoftSignIn.js");
const { licenceDecision } = await import("../utils/licensing.js");
const { ensureIssuerTables } = await import("../utils/licenceIssuer.js");

const TENANT = "11111111-2222-3333-4444-555555555555";
const CLIENT = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";
const SECRET = "Abc8Q~very-secret-value-from-entra-1234";
const { privateKey, publicKey } = crypto.generateKeyPairSync("rsa", { modulusLength: 2048 });
const jwk = { ...publicKey.export({ format: "jwk" }), kid: "key-1" };
const other = crypto.generateKeyPairSync("rsa", { modulusLength: 2048 }).privateKey;

function token(claims, { key = privateKey, kid = "key-1" } = {}) {
  const now = Math.floor(Date.now() / 1000);
  const body = { iss: `https://login.microsoftonline.com/${TENANT}/v2.0`, aud: CLIENT, tid: TENANT, oid: "oid-ana", nonce: "n1",
    email: "Ana.Silva@pilotmining.co.mz", name: "Ana Silva", exp: now + 3600, nbf: now - 10, ...claims };
  const h = Buffer.from(JSON.stringify({ alg: "RS256", kid, typ: "JWT" })).toString("base64url");
  const p = Buffer.from(JSON.stringify(body)).toString("base64url");
  const sig = crypto.sign("RSA-SHA256", Buffer.from(`${h}.${p}`), key).toString("base64url");
  return `${h}.${p}.${sig}`;
}
const keys = async () => [jwk];
const opts = { tenantId: TENANT, clientId: CLIENT, nonce: "n1", keys };

function freshDb() {
  const d = new Database(":memory:");
  d.exec(`CREATE TABLE users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, full_name TEXT, role TEXT, active INTEGER DEFAULT 1, roles_json TEXT);
          CREATE TABLE auth_sessions (id INTEGER PRIMARY KEY AUTOINCREMENT, token TEXT UNIQUE, user_id INTEGER, expires_at TEXT, created_at TEXT DEFAULT (datetime('now')));
          INSERT INTO users (username, role) VALUES ('ana', 'plant_manager'), ('joao', 'storeman'), ('old', 'artisan');`);
  ms.ensureMicrosoftSignInTables(d);
  return d;
}

test("settings are checked, the secret is stored encrypted and never returned", () => {
  const d = freshDb();
  const bad = ms.saveMsSettings(d, { tenant_id: "aml.onmicrosoft.com", client_id: CLIENT, client_secret: CLIENT, enabled: true });
  assert.equal(bad.ok, false);
  assert.match(bad.errors.join(" "), /Directory \(tenant\) ID must look like/);
  assert.match(bad.errors.join(" "), /Copy the secret's Value \(not its Secret ID\)/);
  assert.match(ms.saveMsSettings(d, { tenant_id: TENANT, client_id: CLIENT, enabled: true }).errors[0], /before switching Microsoft sign-in on/);
  const ok = ms.saveMsSettings(d, { tenant_id: TENANT.toUpperCase(), client_id: CLIENT, client_secret: SECRET, secret_expires_on: "2028-10-16", enabled: true }, "jaco");
  assert.equal(ok.ok, true);
  assert.deepEqual([ok.settings.enabled, ok.settings.tenant_id, ok.settings.has_secret], [true, TENANT, true]);
  assert.equal(JSON.stringify(ok.settings).includes(SECRET), false, "the secret never comes back");
  const stored = d.prepare(`SELECT client_secret_enc FROM ms_signin_settings`).get().client_secret_enc;
  assert.equal(stored.includes(SECRET), false, "stored encrypted");
  assert.equal(ms.saveMsSettings(d, { enabled: false }).settings.has_secret, true, "a blank secret keeps the saved one");
  assert.equal(ms.secretDaysLeft({ secret_expires_on: "2026-11-15" }, "2026-10-16"), 30);
});

test("the sign-in starts with a one-time state, nonce and PKCE challenge", async () => {
  const d = freshDb();
  assert.throws(() => ms.startSignIn(d, { redirectUri: "https://x/cb" }), /not switched on/);
  ms.saveMsSettings(d, { tenant_id: TENANT, client_id: CLIENT, client_secret: SECRET, enabled: false });
  assert.throws(() => ms.startSignIn(d, { redirectUri: "https://x/cb" }), /not switched on/, "switched off: only the admin's test works");
  const testUrl = new URL(ms.startSignIn(d, { redirectUri: "https://pilot.ironlogafrica.com/api/sso/microsoft/callback", test: true }));
  assert.equal(testUrl.origin + testUrl.pathname, `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/authorize`);
  assert.equal(testUrl.searchParams.get("code_challenge_method"), "S256");
  assert.equal(testUrl.searchParams.get("scope"), "openid profile email");
  const state = testUrl.searchParams.get("state");
  const req = ms.takeRequest(d, state);
  assert.equal(req.test, 1);
  assert.equal(crypto.createHash("sha256").update(req.verifier).digest("base64url"), testUrl.searchParams.get("code_challenge"));
  assert.equal(ms.takeRequest(d, state), null, "a state works once");

  let sent;
  const fetchImpl = async (url, init) => { sent = { url, body: new URLSearchParams(init.body) }; return new Response(JSON.stringify({ id_token: "t.o.k" }), { status: 200 }); };
  assert.equal(await ms.exchangeCode(d, { code: "c1", verifier: req.verifier, redirectUri: req.redirect_uri, fetchImpl }), "t.o.k");
  assert.equal(sent.url, `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`);
  assert.deepEqual([sent.body.get("client_secret"), sent.body.get("code_verifier"), sent.body.get("grant_type")], [SECRET, req.verifier, "authorization_code"]);
  const refuse = async () => new Response(JSON.stringify({ error: "invalid_client", error_description: "AADSTS7000215: Invalid client secret provided.\nTrace ID: x" }), { status: 401 });
  await assert.rejects(ms.exchangeCode(d, { code: "c1", verifier: "v", redirectUri: "r", fetchImpl: refuse }), /Microsoft did not accept the sign-in: AADSTS7000215: Invalid client secret provided\.$/);
});

test("Microsoft's token is fully checked: signature, company, app, time and one-time nonce", async () => {
  const claims = await ms.verifyIdToken(token({}), opts);
  assert.equal(claims.oid, "oid-ana");
  const code = async (t, o = opts) => (await ms.verifyIdToken(t, o).catch((e) => e)).code;
  assert.equal(await code(token({}, { key: other })), "bad_signature", "signed by someone else");
  assert.equal(await code(token({ tid: "99999999-2222-3333-4444-555555555555", iss: "https://login.microsoftonline.com/99999999-2222-3333-4444-555555555555/v2.0" })), "wrong_tenant", "another company's directory");
  assert.equal(await code(token({ aud: "ffffffff-bbbb-cccc-dddd-eeeeeeeeeeee" })), "bad_token", "meant for another app");
  assert.equal(await code(token({ exp: Math.floor(Date.now() / 1000) - 3600 })), "expired");
  assert.equal(await code(token({ nonce: "old" })), "replayed");
  assert.equal(await code("not.a.token"), "bad_token");
  const [h, p, s] = token({}).split(".");
  const forged = Buffer.from(JSON.stringify({ ...JSON.parse(Buffer.from(p, "base64url")), oid: "oid-boss" })).toString("base64url");
  assert.equal(await code(`${h}.${forged}.${s}`), "bad_signature", "changed claims break the signature");
  let refreshed = false;
  const rotating = async (refresh) => { if (refresh) refreshed = true; return refresh ? [jwk] : [{ ...jwk, kid: "old-key" }]; };
  assert.equal((await ms.verifyIdToken(token({}), { ...opts, keys: rotating })).oid, "oid-ana");
  assert.equal(refreshed, true, "an unknown key ID fetches Microsoft's keys again");
});

test("only invited, active people get in; the Microsoft account is linked on first use", () => {
  const d = freshDb();
  const claims = { tid: TENANT, oid: "oid-ana", email: "Ana.Silva@pilotmining.co.mz" };
  assert.equal((() => { try { ms.matchUser(d, claims); } catch (e) { return e.code; } })(), "not_invited");
  assert.equal(ms.setUserMicrosoftEmail(d, "ana", "ana.silva@pilotmining.co.mz").ok, true);
  assert.match(ms.setUserMicrosoftEmail(d, "joao", "ANA.SILVA@pilotmining.co.mz").error, /already the Microsoft account of ana/);
  assert.match(ms.setUserMicrosoftEmail(d, "joao", "joao").error, /work email address/);
  const first = ms.matchUser(d, claims);
  assert.deepEqual([first.username, first.linked_now], ["ana", true]);
  const renamed = ms.matchUser(d, { ...claims, email: "ana.costa@pilotmining.co.mz" });
  assert.equal(renamed.username, "ana", "linked by Microsoft account ID, so a new email still works");
  assert.equal(renamed.linked_now, false);
  const stranger = (() => { try { ms.matchUser(d, { tid: TENANT, oid: "oid-x", email: "ana.silva@pilotmining.co.mz" }); } catch (e) { return e.code; } })();
  assert.equal(stranger, "not_invited", "another account using the same email is refused once linked");
  d.prepare(`UPDATE users SET active = 0 WHERE username = 'ana'`).run();
  assert.equal((() => { try { ms.matchUser(d, claims); } catch (e) { return e.code; } })(), "inactive");
  ms.setUserMicrosoftEmail(d, "ana", "ana.costa@pilotmining.co.mz");
  assert.equal(d.prepare(`SELECT ms_oid FROM users WHERE username = 'ana'`).get().ms_oid, null, "a new invite email unlinks the old account");
});

test("a Microsoft sign-in gets a 12-hour IronLog session through a one-time handoff", () => {
  const d = freshDb();
  const code = ms.createSessionHandoff(d, 1);
  const tok = ms.takeHandoff(d, code);
  assert.match(tok, /^[0-9a-f]{64}$/);
  assert.equal(ms.takeHandoff(d, code), null, "a handoff code works once");
  const hours = d.prepare(`SELECT (julianday(expires_at) - julianday(created_at)) * 24 AS h FROM auth_sessions WHERE token = ?`).get(tok).h;
  assert.ok(Math.abs(hours - 12) < 0.01, `session lasts ${hours} hours`);
});

test("Microsoft sign-in is a licensed module, in Professional and Enterprise", () => {
  assert.equal(licenceDecision({ mode: "licensed", modules: ["stores"] }, { urlPath: "/api/sso/microsoft/start" }).module, "microsoft_signin");
  assert.equal(licenceDecision({ mode: "licensed", modules: ["microsoft_signin"] }, { urlPath: "/api/sso/microsoft/start" }), null);
  const d = new Database(":memory:");
  d.exec(`CREATE TABLE licence_packages (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, description TEXT, modules_json TEXT NOT NULL DEFAULT '[]', sort INTEGER DEFAULT 0, active INTEGER DEFAULT 1);
          INSERT INTO licence_packages (name, modules_json) VALUES ('Core', '[]'), ('Professional', '["stores"]'), ('Enterprise', '["stores","fleet"]');`);
  ensureIssuerTables(d);
  const mods = Object.fromEntries(d.prepare(`SELECT name, modules_json FROM licence_packages`).all().map((r) => [r.name, JSON.parse(r.modules_json)]));
  assert.deepEqual(mods, { Core: [], Professional: ["stores", "microsoft_signin"], Enterprise: ["stores", "fleet", "microsoft_signin", "api_access"] });
  d.prepare(`UPDATE licence_packages SET modules_json = '["stores"]' WHERE name = 'Professional'`).run();
  ensureIssuerTables(d);
  assert.deepEqual(JSON.parse(d.prepare(`SELECT modules_json FROM licence_packages WHERE name = 'Professional'`).get().modules_json), ["stores"], "taken out by the owner, it stays out");
});
