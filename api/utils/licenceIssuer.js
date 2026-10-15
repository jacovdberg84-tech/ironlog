// IRONLOG/api/utils/licenceIssuer.js — the vendor's licence issuer.
//
// Only the IronLog owner (ISSUER_USERNAMES, signed in as an admin) issues
// licences. The signing key is made once on the server that issues them and
// kept in api/secrets/ (git-ignored, so deploys never touch it); it never
// leaves the server except as a download the owner asks for, to keep offline.
// Customer servers check keys with the matching public key, built into
// licensing.js; the issuing server also trusts its own.
//
// Packages, asset bands and prices are data the owner edits here, never code.
// Prices are optional and only used to show a quote when issuing.
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { ALL_MODULES, MODULES, signLicence } from "./licensing.js";

export const ISSUER_USERNAMES = ["bj.vandenberg"];
const DEFAULT_SECRETS_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "secrets");
const KEY_FILE = "licence-signing.pem";
const secretsDir = () => process.env.IRONLOG_SECRETS_DIR || DEFAULT_SECRETS_DIR;

/** Customer servers run with IRONLOG_LICENCE_ISSUER=0: no issuer there, and no trust in a local key. */
export function issuerEnabled(env = process.env) {
  return String(env.IRONLOG_LICENCE_ISSUER ?? "").trim() !== "0";
}

export function isIssuer(username, roles = []) {
  if (!issuerEnabled()) return false;
  const extra = String(process.env.IRONLOG_LICENCE_ISSUERS || "").split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
  const u = String(username || "").trim().toLowerCase();
  return Boolean(u) && roles.includes("admin") && [...ISSUER_USERNAMES, ...extra].includes(u);
}

export function ensureIssuerTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS licence_packages (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL UNIQUE,
      description TEXT,
      modules_json TEXT NOT NULL DEFAULT '[]',
      sort INTEGER NOT NULL DEFAULT 0,
      active INTEGER NOT NULL DEFAULT 1
    );
    CREATE TABLE IF NOT EXISTS licence_bands (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      label TEXT NOT NULL UNIQUE,
      max_assets INTEGER,
      sort INTEGER NOT NULL DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS licence_prices (
      package_id INTEGER NOT NULL,
      band_id INTEGER NOT NULL,
      monthly_price REAL,
      currency TEXT NOT NULL DEFAULT 'USD',
      PRIMARY KEY (package_id, band_id)
    );
    CREATE TABLE IF NOT EXISTS licences_issued (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      licence_id TEXT NOT NULL UNIQUE,
      licensee TEXT NOT NULL,
      expires_at TEXT NOT NULL,
      payload_json TEXT NOT NULL,
      licence_key TEXT NOT NULL,
      quote_json TEXT,
      issued_by TEXT,
      issued_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
  `);
  // Starting points the owner renames, changes or deletes; no prices.
  if (!db.prepare(`SELECT 1 FROM licence_packages LIMIT 1`).get()) {
    const add = db.prepare(`INSERT INTO licence_packages (name, description, modules_json, sort) VALUES (?, ?, ?, ?)`);
    add.run("Core", "Machines, work orders, breakdowns, maintenance, inspections, technician portal", "[]", 1);
    add.run("Professional", "Core plus stores, fuel, finance and care and maintenance planning", JSON.stringify(["stores", "fuel", "finance", "care_maintenance"]), 2);
    add.run("Enterprise", "Everything, including fleet tracking and the AI assistants", JSON.stringify(ALL_MODULES.filter((m) => m !== "aml_reports")), 3);
  }
  if (!db.prepare(`SELECT 1 FROM licence_bands LIMIT 1`).get()) {
    const add = db.prepare(`INSERT INTO licence_bands (label, max_assets, sort) VALUES (?, ?, ?)`);
    [["Up to 50 machines", 50], ["Up to 150 machines", 150], ["Up to 300 machines", 300], ["Up to 600 machines", 600], ["Unlimited", null]]
      .forEach(([label, max], i) => add.run(label, max, i + 1));
  }
}

// ------------------------------------------------------------------ signing key

function keyPath() {
  return path.join(secretsDir(), KEY_FILE);
}

/** The signing key as PEM text, for the owner to keep a copy offline. */
export function signingKeyPem() {
  try {
    return fs.readFileSync(keyPath(), "utf8");
  } catch {
    return null;
  }
}

/** The issuer's signing key, if this server has one. */
export function loadSigningKey() {
  try {
    return crypto.createPrivateKey(fs.readFileSync(keyPath(), "utf8"));
  } catch {
    return null;
  }
}

export function publicKeyOf(privateKey) {
  return crypto.createPublicKey(privateKey).export({ format: "der", type: "spki" }).toString("base64");
}

/** Makes the signing key once. Refuses to replace an existing one: keys already issued would stop working. */
export function createSigningKey() {
  if (!issuerEnabled()) return { ok: false, error: "Licences are not issued from this server." };
  if (loadSigningKey()) return { ok: false, error: "This server already has a signing key. Replacing it would make every licence already issued stop working." };
  const { privateKey } = crypto.generateKeyPairSync("ed25519");
  fs.mkdirSync(secretsDir(), { recursive: true, mode: 0o700 });
  fs.writeFileSync(keyPath(), privateKey.export({ format: "pem", type: "pkcs8" }), { mode: 0o600, flag: "wx" });
  return { ok: true, public_key: publicKeyOf(privateKey) };
}

/** The public key of this server's signing key, so it trusts its own licences. */
export function localPublicKeys() {
  if (!issuerEnabled()) return [];
  const k = loadSigningKey();
  return k ? [publicKeyOf(k)] : [];
}

// ------------------------------------------------------------------ catalogue

export function catalogue(db) {
  ensureIssuerTables(db);
  const packages = db.prepare(`SELECT * FROM licence_packages WHERE active = 1 ORDER BY sort, id`).all()
    .map((p) => ({ id: p.id, name: p.name, description: p.description || "", modules: JSON.parse(p.modules_json || "[]").filter((m) => MODULES[m]) }));
  const bands = db.prepare(`SELECT * FROM licence_bands ORDER BY sort, id`).all();
  const prices = db.prepare(`SELECT * FROM licence_prices`).all();
  return { packages, bands, prices, modules: Object.fromEntries(Object.entries(MODULES).map(([k, m]) => [k, m.label])) };
}

const clean = (v, max = 200) => String(v ?? "").trim().slice(0, max);

export function savePackage(db, body = {}) {
  ensureIssuerTables(db);
  const name = clean(body.name, 60);
  if (!name) return { ok: false, error: "The package needs a name." };
  const modules = (Array.isArray(body.modules) ? body.modules : []).filter((m) => MODULES[m]);
  try {
    if (body.id) {
      db.prepare(`UPDATE licence_packages SET name = ?, description = ?, modules_json = ? WHERE id = ?`).run(name, clean(body.description, 300), JSON.stringify(modules), Number(body.id));
    } else {
      db.prepare(`INSERT INTO licence_packages (name, description, modules_json, sort) VALUES (?, ?, ?, (SELECT COALESCE(MAX(sort), 0) + 1 FROM licence_packages))`).run(name, clean(body.description, 300), JSON.stringify(modules));
    }
  } catch {
    return { ok: false, error: `There is already a package called ${name}.` };
  }
  return { ok: true };
}

export function saveBand(db, body = {}) {
  ensureIssuerTables(db);
  const label = clean(body.label, 60);
  if (!label) return { ok: false, error: "The band needs a name." };
  const max = String(body.max_assets ?? "").trim() === "" ? null : Number(body.max_assets);
  if (max != null && !(Number.isInteger(max) && max > 0)) return { ok: false, error: "Machines in the band must be a whole number, or blank for unlimited." };
  try {
    if (body.id) db.prepare(`UPDATE licence_bands SET label = ?, max_assets = ? WHERE id = ?`).run(label, max, Number(body.id));
    else db.prepare(`INSERT INTO licence_bands (label, max_assets, sort) VALUES (?, ?, (SELECT COALESCE(MAX(sort), 0) + 1 FROM licence_bands))`).run(label, max);
  } catch {
    return { ok: false, error: `There is already a band called ${label}.` };
  }
  return { ok: true };
}

/** Sets prices: [{ package_id, band_id, monthly_price (blank clears), currency }]. */
export function savePrices(db, rows = []) {
  ensureIssuerTables(db);
  const up = db.prepare(`
    INSERT INTO licence_prices (package_id, band_id, monthly_price, currency) VALUES (?, ?, ?, ?)
    ON CONFLICT(package_id, band_id) DO UPDATE SET monthly_price = excluded.monthly_price, currency = excluded.currency
  `);
  for (const r of rows) {
    const price = String(r.monthly_price ?? "").trim() === "" ? null : Number(r.monthly_price);
    if (price != null && !(Number.isFinite(price) && price >= 0)) return { ok: false, error: "Prices must be numbers of 0 or more." };
  }
  db.transaction(() => {
    for (const r of rows) {
      const price = String(r.monthly_price ?? "").trim() === "" ? null : Number(r.monthly_price);
      up.run(Number(r.package_id), Number(r.band_id), price, clean(r.currency || "USD", 3).toUpperCase() || "USD");
    }
  })();
  return { ok: true };
}

// ------------------------------------------------------------------ issuing

const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ""));

/**
 * Builds and signs a licence. body: { licensee, expires_at, grace_days,
 * sites: [{ site_code, name, package_id, band_id, extra_modules: [] }] }.
 */
export function issueLicence(db, body = {}, { by = "", today = new Date().toISOString().slice(0, 10), privateKey = loadSigningKey() } = {}) {
  ensureIssuerTables(db);
  if (!privateKey) return { ok: false, error: "Set up the signing key on this server first." };
  const licensee = clean(body.licensee, 120);
  if (!licensee) return { ok: false, error: "Who is the licence for? Enter the customer's company name." };
  if (!isDate(body.expires_at) || body.expires_at <= today) return { ok: false, error: "Choose an end date after today." };
  const grace = body.grace_days === undefined || body.grace_days === "" ? 30 : Number(body.grace_days);
  if (!Number.isInteger(grace) || grace < 0 || grace > 365) return { ok: false, error: "Grace days must be a whole number from 0 to 365." };
  const sitesIn = Array.isArray(body.sites) ? body.sites : [];
  if (!sitesIn.length) return { ok: false, error: "Add at least one site." };
  const cat = catalogue(db);
  const sites = {};
  const quote = [];
  for (const s of sitesIn) {
    const code = clean(s.site_code, 40).toLowerCase();
    if (!/^[a-z0-9*][a-z0-9_-]*$|^\*$/.test(code)) return { ok: false, error: `Site code "${s.site_code}" is not valid: use letters, numbers, - or _ (as in the customer's IronLog), or * for every site.` };
    if (sites[code]) return { ok: false, error: `Site ${code} is listed twice.` };
    const pkg = cat.packages.find((p) => p.id === Number(s.package_id));
    const band = cat.bands.find((b) => b.id === Number(s.band_id));
    if (!pkg || !band) return { ok: false, error: `Choose a package and a band for site ${code}.` };
    const extras = (Array.isArray(s.extra_modules) ? s.extra_modules : []).filter((m) => MODULES[m]);
    sites[code] = {
      name: clean(s.name, 120) || null,
      package: pkg.name,
      asset_limit: band.max_assets ?? null,
      modules: [...new Set([...pkg.modules, ...extras])].sort(),
    };
    const price = cat.prices.find((p) => p.package_id === pkg.id && p.band_id === band.id);
    quote.push({ site_code: code, package: pkg.name, band: band.label, monthly_price: price?.monthly_price ?? null, currency: price?.currency || "USD" });
  }
  const n = db.prepare(`SELECT COUNT(*) AS n FROM licences_issued WHERE issued_at >= date('now', 'start of year')`).get().n + 1;
  const licenceId = `IL-${today.slice(0, 4)}-${String(n).padStart(4, "0")}-${crypto.randomBytes(2).toString("hex").toUpperCase()}`;
  const payload = { v: 1, licence_id: licenceId, licensee, issued_at: today, expires_at: body.expires_at, grace_days: grace, sites };
  const key = signLicence(payload, privateKey);
  const priced = quote.filter((q) => q.monthly_price != null);
  const total = priced.length === quote.length && new Set(quote.map((q) => q.currency)).size === 1
    ? { monthly: Number(priced.reduce((s, q) => s + q.monthly_price, 0).toFixed(2)), currency: quote[0].currency }
    : null;
  db.prepare(`INSERT INTO licences_issued (licence_id, licensee, expires_at, payload_json, licence_key, quote_json, issued_by) VALUES (?, ?, ?, ?, ?, ?, ?)`)
    .run(licenceId, licensee, body.expires_at, JSON.stringify(payload), key, JSON.stringify({ sites: quote, total }), by || null);
  return { ok: true, licence_id: licenceId, key, payload, quote: { sites: quote, total } };
}

export function listIssued(db) {
  ensureIssuerTables(db);
  return db.prepare(`SELECT id, licence_id, licensee, expires_at, payload_json, quote_json, issued_by, issued_at FROM licences_issued ORDER BY id DESC LIMIT 100`).all()
    .map((r) => ({ ...r, payload: JSON.parse(r.payload_json), quote: r.quote_json ? JSON.parse(r.quote_json) : null, payload_json: undefined, quote_json: undefined }));
}

export function issuedKey(db, id) {
  ensureIssuerTables(db);
  return db.prepare(`SELECT licence_id, licence_key FROM licences_issued WHERE id = ?`).get(Number(id)) || null;
}
