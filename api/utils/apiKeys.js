// IRONLOG/api/utils/apiKeys.js — keys that let another system (an ERP, Power BI, a client portal) read IronLog.
//
// An administrator creates one key per outside system, picks what it may read
// (scopes) and, optionally, which sites. The key is shown once; only its SHA-256
// hash is stored. Keys are read-only: each scope is a short list of GET addresses,
// so a key can never change a record or reach settings, users or admin screens.
// A key can be cancelled at any time; every use updates "last used", and keys are
// limited to API_KEY_RATE_PER_MIN requests a minute. Using keys is the licensed
// "api_access" module (Enterprise).
import crypto from "node:crypto";

export const API_KEY_PREFIX = "ilk_";
export const API_KEY_RATE_PER_MIN = 300;

// What a key may read. Each address is matched exactly (GET only).
export const API_SCOPES = {
  "assets:read": {
    label: "Machines",
    endpoints: [
      { path: "/api/assets", re: /^\/api\/assets$/, text: "All machines with their details (make, model, serial, area, site, cost centre)." },
      { path: "/api/assets/fleet-summary", re: /^\/api\/assets\/fleet-summary$/, text: "Every machine on the site with its status, meter hours and fuel used in the last 30 days." },
      { path: "/api/assets/{asset_code}/components", re: /^\/api\/assets\/[^/]+\/components$/, text: "A machine's major components: serial, hours since fitted, life used, change history." },
      { path: "/api/assets/component-life", re: /^\/api\/assets\/component-life$/, text: "Components on the site at 80% or more of their expected life (?min_pct=)." },
    ],
  },
  "breakdowns:read": {
    label: "Breakdowns and failure codes",
    endpoints: [
      { path: "/api/breakdowns", re: /^\/api\/breakdowns$/, text: "Breakdowns on the site, newest first (?start=YYYY-MM-DD&end=YYYY-MM-DD)." },
      { path: "/api/breakdowns/{id}", re: /^\/api\/breakdowns\/\d+$/, text: "One breakdown with its downtime and work order." },
      { path: "/api/breakdowns/failure-codes", re: /^\/api\/breakdowns\/failure-codes$/, text: "The failure code lists (component, problem, cause, action) in English and Portuguese." },
      { path: "/api/breakdowns/failure-codes/summary", re: /^\/api\/breakdowns\/failure-codes\/summary$/, text: "Coded breakdowns counted by component, problem and cause (?start&end)." },
    ],
  },
  "workorders:read": {
    label: "Work orders",
    endpoints: [
      { path: "/api/workorders", re: /^\/api\/workorders$/, text: "Work orders on the site with status, machine, technician and hours." },
      { path: "/api/workorders/{id}", re: /^\/api\/workorders\/\d+$/, text: "One work order with its parts and labour." },
    ],
  },
  "hours:read": {
    label: "Machine hours and fleet KPIs",
    endpoints: [
      { path: "/api/hours/{date}", re: /^\/api\/hours\/\d{4}-\d{2}-\d{2}$/, text: "Daily input for one day: run hours, opening and closing meters per machine." },
      { path: "/api/kpi/daily/{date}", re: /^\/api\/kpi\/daily\/\d{4}-\d{2}-\d{2}$/, text: "Fleet availability and utilisation for one day." },
      { path: "/api/kpi/range", re: /^\/api\/kpi\/range$/, text: "Fleet availability and utilisation over a period (?start&end)." },
    ],
  },
  "reliability:read": {
    label: "Reliability",
    endpoints: [
      { path: "/api/maintenance/reliability", re: /^\/api\/maintenance\/reliability$/, text: "MTBF, LTTR, availability and maintenance cost per operating hour by machine, model, area, class and failed component (?start&end)." },
    ],
  },
  "stock:read": {
    label: "Stores",
    endpoints: [
      { path: "/api/stock/onhand", re: /^\/api\/stock\/onhand$/, text: "Parts on hand with quantities and unit cost (USD)." },
      { path: "/api/stock/part-orders", re: /^\/api\/stock\/part-orders$/, text: "Parts ordered for machines and where they stand." },
    ],
  },
};

const sha256 = (s) => crypto.createHash("sha256").update(String(s)).digest("hex");

export function ensureApiKeyTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS api_keys (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      key_prefix TEXT NOT NULL,
      key_hash TEXT NOT NULL UNIQUE,
      scopes_json TEXT NOT NULL,
      sites_json TEXT,
      expires_on TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      created_by TEXT,
      last_used_at TEXT,
      last_used_path TEXT,
      use_count INTEGER NOT NULL DEFAULT 0,
      revoked_at TEXT,
      revoked_by TEXT
    );
  `);
}

const parse = (raw) => {
  try {
    const v = JSON.parse(raw || "null");
    return Array.isArray(v) ? v : null;
  } catch {
    return null;
  }
};

function publicRow(r, today = new Date().toISOString().slice(0, 10)) {
  const expired = Boolean(r.expires_on && r.expires_on < today);
  return {
    id: r.id,
    name: r.name,
    key_prefix: r.key_prefix,
    scopes: parse(r.scopes_json) || [],
    sites: parse(r.sites_json),
    expires_on: r.expires_on || null,
    created_at: r.created_at,
    created_by: r.created_by,
    last_used_at: r.last_used_at,
    last_used_path: r.last_used_path,
    use_count: Number(r.use_count || 0),
    revoked_at: r.revoked_at,
    revoked_by: r.revoked_by,
    status: r.revoked_at ? "cancelled" : expired ? "expired" : "active",
  };
}

export function listApiKeys(db) {
  ensureApiKeyTables(db);
  return db.prepare(`SELECT * FROM api_keys ORDER BY revoked_at IS NOT NULL, created_at DESC, id DESC`).all().map((r) => publicRow(r));
}

/**
 * Creates a key: { ok, key (shown once), row } or { ok:false, errors }.
 * checkSites(list) → { ok, sites } | { ok:false, error } (the company's real sites).
 */
export function createApiKey(db, body = {}, { by = "", checkSites = (l) => ({ ok: true, sites: l }), today = new Date().toISOString().slice(0, 10) } = {}) {
  ensureApiKeyTables(db);
  const errors = [];
  const name = String(body.name || "").trim();
  if (!name || name.length > 60) errors.push("Give the key a name of up to 60 characters, e.g. the system it is for (Sage, Power BI).");
  const scopes = [...new Set((Array.isArray(body.scopes) ? body.scopes : []).map((s) => String(s).trim()))];
  const bad = scopes.filter((s) => !API_SCOPES[s]);
  if (!scopes.length || bad.length) errors.push(`Choose what the key may read${bad.length ? ` (unknown: ${bad.join(", ")})` : ""}.`);
  let sites = null;
  const siteList = (Array.isArray(body.sites) ? body.sites : String(body.sites || "").split(",")).map((s) => String(s).trim().toLowerCase()).filter(Boolean);
  if (siteList.length) {
    const r = checkSites(siteList);
    if (!r.ok) errors.push(r.error);
    else sites = r.sites;
  }
  const expires = String(body.expires_on || "").trim();
  if (expires && (!/^\d{4}-\d{2}-\d{2}$/.test(expires) || expires < today)) errors.push("The end date must be a date like 2027-10-17, today or later.");
  if (errors.length) return { ok: false, errors };
  const key = `${API_KEY_PREFIX}${crypto.randomBytes(30).toString("base64url")}`;
  const id = db.prepare(`INSERT INTO api_keys (name, key_prefix, key_hash, scopes_json, sites_json, expires_on, created_by) VALUES (?, ?, ?, ?, ?, ?, ?)`)
    .run(name, key.slice(0, 12), sha256(key), JSON.stringify(scopes), sites ? JSON.stringify(sites) : null, expires || null, by || null).lastInsertRowid;
  return { ok: true, key, row: publicRow(db.prepare(`SELECT * FROM api_keys WHERE id = ?`).get(id), today) };
}

export function revokeApiKey(db, id, by = "") {
  ensureApiKeyTables(db);
  const r = db.prepare(`UPDATE api_keys SET revoked_at = datetime('now'), revoked_by = ? WHERE id = ? AND revoked_at IS NULL`).run(by || null, Number(id));
  return r.changes > 0;
}

/** The active key for a presented secret, or null (unknown, cancelled or past its end date). */
export function resolveApiKey(db, presented, today = new Date().toISOString().slice(0, 10)) {
  const s = String(presented || "").trim();
  if (!s.startsWith(API_KEY_PREFIX) || s.length < 20 || s.length > 100) return null;
  ensureApiKeyTables(db);
  const r = db.prepare(`SELECT * FROM api_keys WHERE key_hash = ?`).get(sha256(s));
  if (!r || r.revoked_at || (r.expires_on && r.expires_on < today)) return null;
  return { ...publicRow(r, today), sites: parse(r.sites_json) };
}

/** True when the key's scopes allow this request (GET only, listed addresses only). */
export function keyAllows(key, method, urlPath) {
  if (!key || String(method).toUpperCase() !== "GET") return false;
  return (key.scopes || []).some((s) => (API_SCOPES[s]?.endpoints || []).some((e) => e.re.test(urlPath)));
}

/** Notes a use (at most once a minute per key, so a busy integration does not write on every call). */
export function touchApiKey(db, key, urlPath) {
  db.prepare(`
    UPDATE api_keys SET use_count = use_count + 1,
      last_used_path = CASE WHEN last_used_at IS NULL OR last_used_at < datetime('now', '-60 seconds') THEN ? ELSE last_used_path END,
      last_used_at = CASE WHEN last_used_at IS NULL OR last_used_at < datetime('now', '-60 seconds') THEN datetime('now') ELSE last_used_at END
    WHERE id = ?
  `).run(urlPath, key.id);
}

const windows = new Map();
/** True when the key is still under its requests-per-minute limit. */
export function underRateLimit(keyId, now = Date.now(), limit = API_KEY_RATE_PER_MIN) {
  const minute = Math.floor(now / 60000);
  const w = windows.get(keyId);
  if (!w || w.minute !== minute) {
    windows.set(keyId, { minute, count: 1 });
    return true;
  }
  w.count += 1;
  return w.count <= limit;
}

/** The scopes and their addresses, for the documentation page and the OpenAPI file. */
export function apiCatalogue() {
  return Object.entries(API_SCOPES).map(([scope, s]) => ({
    scope,
    label: s.label,
    endpoints: s.endpoints.map((e) => ({ method: "GET", path: e.path, description: e.text })),
  }));
}

/** OpenAPI 3 description of what keys can read. */
export function openApiDocument(serverUrl = "") {
  const paths = {};
  for (const s of apiCatalogue()) {
    for (const e of s.endpoints) {
      const params = [...e.path.matchAll(/\{(\w+)\}/g)].map((m) => ({ name: m[1], in: "path", required: true, schema: { type: "string" } }));
      paths[e.path] = {
        get: {
          summary: e.description,
          tags: [s.label],
          "x-ironlog-scope": s.scope,
          parameters: [...params, { name: "X-Site-Code", in: "header", required: false, schema: { type: "string" }, description: "Which site (the key's first site when not given)." }],
          responses: { 200: { description: "OK (JSON)" }, 401: { description: "Missing, unknown, cancelled or expired key" }, 403: { description: "Not allowed for this key's scopes, sites or licence" }, 429: { description: "Too many requests this minute" } },
          security: [{ ironlogKey: [] }],
        },
      };
    }
  }
  return {
    openapi: "3.0.3",
    info: { title: "IronLog API", version: "1", description: "Read-only access for other systems. Create a key under Admin → API keys. Money is US dollars; dates are YYYY-MM-DD." },
    servers: serverUrl ? [{ url: serverUrl }] : [],
    components: { securitySchemes: { ironlogKey: { type: "http", scheme: "bearer", description: "Authorization: Bearer ilk_…" } } },
    paths,
  };
}

