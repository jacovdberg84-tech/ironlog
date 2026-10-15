import test from "node:test";
import assert from "node:assert/strict";
import crypto from "node:crypto";
import Database from "better-sqlite3";
import * as lic from "../utils/licensing.js";

const { privateKey, publicKey } = crypto.generateKeyPairSync("ed25519");
const PUB = publicKey.export({ format: "der", type: "spki" }).toString("base64");
const keys = [PUB];

const freshDb = (withAssets) => {
  const d = new Database(":memory:");
  d.exec(`CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, site_code TEXT, active INTEGER DEFAULT 1, archived INTEGER DEFAULT 0);
          CREATE TABLE users (id INTEGER PRIMARY KEY, username TEXT, active INTEGER DEFAULT 1);`);
  if (withAssets) d.exec(`INSERT INTO assets (asset_code, site_code) VALUES ('E501AM', 'main')`);
  return d;
};

const sample = {
  v: 1,
  licence_id: "LIC-PILOT-001",
  licensee: "Pilot Mining Lda",
  issued_at: "2026-10-15",
  expires_at: "2027-10-14",
  grace_days: 30,
  sites: {
    main: { name: "Main pit", package: "Professional", asset_limit: 2, modules: ["stores", "fuel", "made_up_module"] },
  },
};

test("only keys signed by IronLog are accepted; any change breaks the signature", () => {
  const key = lic.signLicence(sample, privateKey);
  assert.match(key, /^IRONLOG1\.[\w-]+\.[\w-]+$/);
  assert.equal(lic.verifyLicence(key, keys).ok, true);
  const [p, body, sig] = key.split(".");
  const forged = Buffer.from(JSON.stringify({ ...sample, expires_at: "2099-01-01" })).toString("base64url");
  assert.match(lic.verifyLicence(`${p}.${forged}.${sig}`, keys).error, /not issued by IronLog, or it has been changed/);
  const other = crypto.generateKeyPairSync("ed25519").privateKey;
  assert.equal(lic.verifyLicence(lic.signLicence(sample, other), keys).ok, false, "someone else's signing key");
  assert.match(lic.verifyLicence("hello", keys).error, /not an IronLog licence key/);
  assert.equal(lic.verifyLicence(key, []).ok, false, "no trusted keys, nothing verifies");
  assert.equal(lic.verifyLicence(`${p}.${body}`, keys).ok, false);
});

test("an existing install runs fully unlocked; a new one gets a 30-day trial, then grace, then read-only", () => {
  const legacy = lic.licenceState(freshDb(true), { today: "2026-10-15", publicKeys: keys });
  assert.equal(legacy.mode, "legacy");
  assert.deepEqual(legacy.modules, lic.ALL_MODULES);
  assert.deepEqual(legacy.warnings, []);

  const d = freshDb(false);
  d.exec(`CREATE TABLE IF NOT EXISTS licence_state (id INTEGER PRIMARY KEY, licence_key TEXT, installed_at TEXT, installed_by TEXT, legacy INTEGER NOT NULL DEFAULT 0, trial_started_at TEXT NOT NULL)`);
  d.exec(`INSERT INTO licence_state (id, legacy, trial_started_at) VALUES (1, 0, '2026-10-01')`);
  const trial = lic.licenceState(d, { today: "2026-10-15", publicKeys: keys });
  assert.equal(trial.mode, "trial");
  assert.equal(trial.expires_at, "2026-10-31");
  assert.equal(trial.days_left, 16);
  assert.deepEqual(trial.modules, lic.ALL_MODULES, "everything on during the trial");
  assert.match(trial.warnings[0], /trial: 16 days left/);
  assert.equal(lic.licenceState(d, { today: "2026-11-15", publicKeys: keys }).mode, "grace");
  const ro = lic.licenceState(d, { today: "2026-12-01", publicKeys: keys });
  assert.equal(ro.mode, "read_only");
  assert.equal(ro.read_only_from, "2026-12-01");
  assert.equal(lic.licenceState(d, { today: "2026-11-30", publicKeys: keys }).mode, "grace", "the last grace day still works");
});

test("a licence switches modules per site, warns before it ends, then grace, then read-only", () => {
  const d = freshDb(true);
  assert.equal(lic.installLicence(d, "IRONLOG1.bad.key", "jaco", { publicKeys: keys }).ok, false);
  assert.equal(lic.installLicence(d, lic.signLicence(sample, privateKey), "jaco", { publicKeys: keys }).ok, true);
  const s = lic.licenceState(d, { siteCode: "main", today: "2027-01-10", publicKeys: keys });
  assert.equal(s.mode, "licensed");
  assert.deepEqual(s.modules, ["stores", "fuel"], "unknown module names are ignored");
  assert.equal(s.package, "Professional");
  assert.deepEqual(s.warnings, []);
  assert.match(lic.licenceState(d, { siteCode: "main", today: "2027-10-01", publicKeys: keys }).warnings[0], /ends on 2027-10-14 \(13 days\)/);
  assert.equal(lic.licenceState(d, { siteCode: "main", today: "2027-11-13", publicKeys: keys }).mode, "grace");
  assert.equal(lic.licenceState(d, { siteCode: "main", today: "2027-11-14", publicKeys: keys }).mode, "read_only");
  const other = lic.licenceState(d, { siteCode: "north", today: "2027-01-10", publicKeys: keys });
  assert.deepEqual(other.modules, [], "a site not in the licence gets core only");
  assert.match(other.warnings[0], /not in the licence/);
  // A licence with a "*" entry covers every other site.
  lic.installLicence(d, lic.signLicence({ ...sample, sites: { "*": { package: "Core", modules: ["fleet"] } } }, privateKey), "jaco", { publicKeys: keys });
  assert.deepEqual(lic.licenceState(d, { siteCode: "north", today: "2027-01-10", publicKeys: keys }).modules, ["fleet"]);
  // A key that no longer verifies (keys rotated) falls back as if none were installed, and says so.
  const fallback = lic.licenceState(d, { siteCode: "main", today: "2027-01-10", publicKeys: [] });
  assert.equal(fallback.mode, "legacy");
  assert.match(fallback.warnings[0], /installed licence is not valid/);
});

test("the server refuses unlicensed modules and, when read-only, any save", () => {
  const state = { mode: "licensed", modules: ["stores"] };
  assert.equal(lic.licenceDecision(state, { method: "GET", urlPath: "/api/stock/items" }), null);
  const fuel = lic.licenceDecision(state, { method: "GET", urlPath: "/api/dashboard/fuel/daily" });
  assert.equal(fuel.status, 403);
  assert.match(fuel.error, /Fuel and lubricants is not included/);
  assert.equal(lic.licenceDecision(state, { method: "POST", urlPath: "/api/maintenance/cm/plans" }).module, "care_maintenance");
  assert.equal(lic.licenceDecision(state, { method: "POST", urlPath: "/api/maintenance/plans" }), null, "core maintenance is always on");
  assert.equal(lic.licenceDecision(state, { method: "GET", urlPath: "/api/reports/aml-weekly-check-sheet.xlsx" }).module, "aml_reports");
  assert.equal(lic.licenceDecision(state, { method: "GET", urlPath: "/api/dashboard/kpis" }), null);

  const ro = { mode: "read_only", modules: lic.ALL_MODULES };
  assert.equal(lic.licenceDecision(ro, { method: "GET", urlPath: "/api/workorders" }), null, "reading still works");
  assert.equal(lic.licenceDecision(ro, { method: "POST", urlPath: "/api/workorders" }).status, 423);
  assert.equal(lic.licenceDecision(ro, { method: "POST", urlPath: "/api/auth/login" }), null, "people can still sign in");
  assert.equal(lic.licenceDecision(ro, { method: "POST", urlPath: "/api/licence/install" }), null, "and enter a new key");
  assert.equal(lic.licenceDecision(ro, { method: "POST", urlPath: "/api/admin/backups/run-now" }), null, "and back up");
  assert.equal(lic.licenceDecision(ro, { method: "GET", urlPath: "/web/index.html" }), null);
});

test("usage counts active machines per site against the band and flags overruns", () => {
  const d = freshDb(true);
  d.exec(`INSERT INTO assets (asset_code, site_code) VALUES ('D600AM', 'main'), ('L1', 'main'), ('OLD', 'main');
          UPDATE assets SET archived = 1 WHERE asset_code = 'OLD';
          INSERT INTO assets (asset_code, site_code) VALUES ('N1', 'north');
          INSERT INTO users (username) VALUES ('a'), ('b');`);
  lic.installLicence(d, lic.signLicence(sample, privateKey), "jaco", { publicKeys: keys });
  const u = lic.licenceUsage(d, { today: "2027-01-10", publicKeys: keys });
  assert.equal(u.active_users, 2);
  const main = u.sites.find((s) => s.site_code === "main");
  assert.equal(main.active_assets, 3, "archived machines do not count");
  assert.equal(main.asset_limit, 2);
  assert.equal(main.over_band, true);
  assert.match(u.warnings[0], /Site main has 3 active machines; the licence covers 2/);
  assert.equal(u.sites.find((s) => s.site_code === "north").over_band, false);
});

test("the built-in vendor key is a valid Ed25519 public key, and a key it did not sign is refused", () => {
  assert.equal(lic.BUILT_IN_PUBLIC_KEYS.length, 1);
  const k = crypto.createPublicKey({ key: Buffer.from(lic.BUILT_IN_PUBLIC_KEYS[0], "base64"), format: "der", type: "spki" });
  assert.equal(k.asymmetricKeyType, "ed25519");
  assert.equal(lic.verifyLicence(lic.signLicence(sample, privateKey), lic.BUILT_IN_PUBLIC_KEYS).ok, false, "only the owner's signing key passes");
});
