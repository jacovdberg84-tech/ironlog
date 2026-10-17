// IRONLOG/api/utils/licensing.js — what a customer's licence lets them use.
//
// A licence is a key signed by the IronLog vendor (Ed25519). It names the
// licensee, the end date, the grace period and, per site, the package, the
// asset band and the add-on modules. Prices never appear here: packages and
// prices live with the vendor's licence issuer, and the key carries only the
// resulting entitlements.
//
// Modes, worked out on every request (cached briefly):
//   legacy     an install that already held data before licensing: everything on
//   trial      a new install without a licence: everything on for 30 days
//   licensed   a valid key: its modules per site
//   grace      past the end date, within the grace days: as licensed, with warnings
//   read_only  past the grace period: everything can be read and exported, nothing saved
// Over the asset band only warns; adding a machine is never blocked.
import crypto from "node:crypto";

export const TRIAL_DAYS = 30;
export const WARN_DAYS = 30;
const DEFAULT_GRACE_DAYS = 30;
const KEY_PREFIX = "IRONLOG1";

// Core (machines, work orders, breakdowns, maintenance plans, inspections,
// technician portal, tasks, reports) is always included. Add-on modules:
export const MODULES = {
  stores: { label: "Stores and procurement", api: ["/api/stock", "/api/procurement"], tabs: ["stock", "parts-tracking", "procurement"] },
  fuel: { label: "Fuel and lubricants", api: ["/api/dashboard/fuel", "/api/dashboard/lube", "/api/fuel"], tabs: ["fuel", "lube"] },
  fleet: { label: "Fleet tracking and security", api: ["/api/cartrack", "/api/unitech", "/api/telematics"], tabs: ["cartrack", "telematics"] },
  finance: { label: "Finance and costing", api: ["/api/finance", "/api/costing"], tabs: ["finance"] },
  ai: { label: "AI assistants (Borris and June)", api: ["/api/ironmind", "/api/june", "/api/ai-engineer"], tabs: ["ironmind"] },
  care_maintenance: { label: "Care and maintenance planning", api: ["/api/maintenance/cm"], tabs: [] },
  developer: { label: "Developer Centre", api: ["/api/developer"], tabs: ["developer"] },
  aml_reports: { label: "AML customer reports", api: ["/api/reports/aml-weekly-check-sheet"], tabs: [] },
  microsoft_signin: { label: "Sign in with Microsoft 365", api: ["/api/sso/microsoft"], tabs: [] },
  api_access: { label: "API access for other systems", api: ["/api/api-keys", "/api/webhooks"], tabs: [] },
};
export const ALL_MODULES = Object.keys(MODULES);

// Requests that always work, even read-only: signing in, the licence itself, backups.
const ALWAYS_ALLOWED = ["/api/auth/", "/api/licence", "/api/admin/backups"];

export function ensureLicenceTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS licence_state (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      licence_key TEXT,
      installed_at TEXT,
      installed_by TEXT,
      legacy INTEGER NOT NULL DEFAULT 0,
      trial_started_at TEXT NOT NULL DEFAULT (date('now'))
    );
  `);
  if (!db.prepare(`SELECT 1 FROM licence_state WHERE id = 1`).get()) {
    let hasData = false;
    try { hasData = Number(db.prepare(`SELECT COUNT(*) AS n FROM assets`).get()?.n || 0) > 0; } catch {}
    db.prepare(`INSERT INTO licence_state (id, legacy) VALUES (1, ?)`).run(hasData ? 1 : 0);
  }
}

/** Public keys the vendor signs with: built in, plus IRONLOG_LICENCE_PUBLIC_KEYS (comma-separated, base64 SPKI). */
export const BUILT_IN_PUBLIC_KEYS = [
  // IronLog licence signing key, set up on the issuing server on 2026-10-15.
  "MCowBQYDK2VwAyEAEfynEM7DjNmIXX66VEBcCQQb3EL+rHjpxorB2N3LJTM=",
];
export function vendorPublicKeys(env = process.env) {
  const extra = String(env.IRONLOG_LICENCE_PUBLIC_KEYS || "").split(",").map((s) => s.trim()).filter(Boolean);
  return [...BUILT_IN_PUBLIC_KEYS, ...extra];
}

const b64u = (buf) => Buffer.from(buf).toString("base64url");

/** Signs a licence payload (vendor side). privateKey: a KeyObject or PEM. */
export function signLicence(payload, privateKey) {
  const body = Buffer.from(JSON.stringify(payload), "utf8");
  const sig = crypto.sign(null, body, privateKey);
  return `${KEY_PREFIX}.${b64u(body)}.${b64u(sig)}`;
}

/** Checks a licence key's signature and shape; returns { ok, licence } or { ok: false, error }. */
export function verifyLicence(key, publicKeys = vendorPublicKeys()) {
  const parts = String(key || "").trim().split(".");
  if (parts.length !== 3 || parts[0] !== KEY_PREFIX) return { ok: false, error: "That is not an IronLog licence key." };
  let body;
  let sig;
  try {
    body = Buffer.from(parts[1], "base64url");
    sig = Buffer.from(parts[2], "base64url");
  } catch {
    return { ok: false, error: "The licence key is damaged. Copy it again in full." };
  }
  const trusted = publicKeys.some((pk) => {
    try {
      const keyObj = crypto.createPublicKey({ key: Buffer.from(pk, "base64"), format: "der", type: "spki" });
      return crypto.verify(null, body, keyObj, sig);
    } catch {
      return false;
    }
  });
  if (!trusted) return { ok: false, error: "The licence key was not issued by IronLog, or it has been changed." };
  let licence;
  try {
    licence = JSON.parse(body.toString("utf8"));
  } catch {
    return { ok: false, error: "The licence key is damaged. Copy it again in full." };
  }
  if (!licence || !/^\d{4}-\d{2}-\d{2}$/.test(String(licence.expires_at || "")) || typeof licence.sites !== "object") {
    return { ok: false, error: "The licence key is incomplete." };
  }
  return { ok: true, licence };
}

const addDays = (ymd, days) => new Date(Date.parse(`${ymd}T00:00:00Z`) + days * 86400000).toISOString().slice(0, 10);
const daysBetween = (fromYmd, toYmd) => Math.round((Date.parse(`${toYmd}T00:00:00Z`) - Date.parse(`${fromYmd}T00:00:00Z`)) / 86400000);

function siteEntitlement(licence, siteCode) {
  const sites = licence.sites || {};
  const s = sites[siteCode] || sites["*"] || null;
  if (!s) return { package: null, asset_limit: null, modules: [], listed: false };
  const modules = Array.isArray(s.modules) ? s.modules.filter((m) => MODULES[m]) : [];
  const limit = Number(s.asset_limit);
  return { package: s.package || null, asset_limit: Number.isFinite(limit) && limit > 0 ? limit : null, modules, listed: true, name: s.name || null };
}

/**
 * The licence situation for a site today. `today` is YYYY-MM-DD (site time).
 * Returns the mode, the modules switched on, dates and the warnings to show.
 */
export function licenceState(db, { siteCode = "main", today = new Date().toISOString().slice(0, 10), publicKeys } = {}) {
  ensureLicenceTables(db);
  const row = db.prepare(`SELECT * FROM licence_state WHERE id = 1`).get();
  const out = { site_code: siteCode, today, warnings: [] };
  let verified = null;
  if (row.licence_key) {
    verified = verifyLicence(row.licence_key, publicKeys);
    if (!verified.ok) out.warnings.push(`The installed licence is not valid: ${verified.error}`);
  }

  if (verified?.ok) {
    const lic = verified.licence;
    const grace = Number.isInteger(lic.grace_days) && lic.grace_days >= 0 ? lic.grace_days : DEFAULT_GRACE_DAYS;
    const site = siteEntitlement(lic, siteCode);
    const left = daysBetween(today, lic.expires_at);
    const readOnlyFrom = addDays(lic.expires_at, grace + 1);
    Object.assign(out, {
      licence_id: lic.licence_id || null,
      licensee: lic.licensee || "",
      expires_at: lic.expires_at,
      read_only_from: readOnlyFrom,
      days_left: left,
      package: site.package,
      asset_limit: site.asset_limit,
      modules: site.modules,
      sites: Object.keys(lic.sites || {}),
    });
    if (left >= 0) {
      out.mode = "licensed";
      if (left <= WARN_DAYS) out.warnings.push(`The IronLog licence ends on ${lic.expires_at} (${left} day${left === 1 ? "" : "s"}). Renew it to avoid interruption.`);
    } else if (today < readOnlyFrom) {
      out.mode = "grace";
      out.warnings.push(`The IronLog licence ended on ${lic.expires_at}. IronLog becomes read-only on ${readOnlyFrom} unless it is renewed.`);
    } else {
      out.mode = "read_only";
      out.warnings.push(`The IronLog licence ended on ${lic.expires_at}. IronLog is read-only: you can view and export, but not save. Ask your administrator to renew.`);
    }
    if (!site.listed) out.warnings.push(`This site (${siteCode}) is not in the licence, so only the core modules are on.`);
    return out;
  }

  if (row.legacy) {
    return { ...out, mode: "legacy", modules: [...ALL_MODULES], expires_at: null, days_left: null };
  }
  const trialEnd = addDays(row.trial_started_at, TRIAL_DAYS);
  const left = daysBetween(today, trialEnd);
  const readOnlyFrom = addDays(trialEnd, DEFAULT_GRACE_DAYS + 1);
  Object.assign(out, { expires_at: trialEnd, read_only_from: readOnlyFrom, days_left: left, modules: [...ALL_MODULES], trial_started_at: row.trial_started_at });
  if (left >= 0) {
    out.mode = "trial";
    out.warnings.push(`IronLog trial: ${left} day${left === 1 ? "" : "s"} left. Enter a licence key under Admin → Licence.`);
  } else if (today < readOnlyFrom) {
    out.mode = "grace";
    out.warnings.push(`The IronLog trial ended on ${trialEnd}. IronLog becomes read-only on ${readOnlyFrom} unless a licence key is entered.`);
  } else {
    out.mode = "read_only";
    out.warnings.push(`The IronLog trial ended on ${trialEnd}. IronLog is read-only until a licence key is entered under Admin → Licence.`);
  }
  return out;
}

/** Installs a licence key after checking it. */
export function installLicence(db, key, actor = "", { publicKeys } = {}) {
  const v = verifyLicence(key, publicKeys);
  if (!v.ok) return v;
  ensureLicenceTables(db);
  db.prepare(`UPDATE licence_state SET licence_key = ?, installed_at = datetime('now'), installed_by = ? WHERE id = 1`).run(String(key).trim(), actor || null);
  return { ok: true, licence: v.licence };
}

/** The add-on module an API path belongs to, or null for core. */
export function moduleForPath(urlPath) {
  for (const [key, m] of Object.entries(MODULES)) {
    if (m.api.some((p) => urlPath === p || urlPath.startsWith(`${p}/`) || urlPath.startsWith(`${p}.`) || urlPath.startsWith(`${p}?`))) return key;
  }
  return null;
}

/**
 * The server check for one request: null when allowed, else { status, error }.
 * Read-only blocks any change; a module outside the licence is refused.
 */
export function licenceDecision(state, { method = "GET", urlPath = "" } = {}) {
  if (!urlPath.startsWith("/api/") || ALWAYS_ALLOWED.some((p) => urlPath.startsWith(p))) return null;
  const mod = moduleForPath(urlPath);
  if (mod && !(state.modules || []).includes(mod)) {
    return { status: 403, error: `${MODULES[mod].label} is not included in this site's IronLog licence.`, module: mod };
  }
  if (state.mode === "read_only" && !["GET", "HEAD", "OPTIONS"].includes(String(method).toUpperCase())) {
    return { status: 423, error: "IronLog is read-only because the licence has ended. You can view and export, but not save. Ask your administrator to renew." };
  }
  return null;
}

/** Usage for the licence page and the vendor's invoice: machines and users per site, against the bands. */
export function licenceUsage(db, { today, publicKeys } = {}) {
  ensureLicenceTables(db);
  const sites = new Map();
  try {
    for (const r of db.prepare(`
      SELECT LOWER(COALESCE(NULLIF(TRIM(site_code), ''), 'main')) AS site, COUNT(*) AS n
      FROM assets WHERE COALESCE(active, 1) = 1 AND COALESCE(archived, 0) = 0 GROUP BY 1
    `).all()) sites.set(r.site, { site_code: r.site, active_assets: r.n });
  } catch {}
  let users = 0;
  try { users = db.prepare(`SELECT COUNT(*) AS n FROM users WHERE COALESCE(active, 1) = 1`).get().n; } catch {}
  const rows = [...sites.values()].map((s) => {
    const st = licenceState(db, { siteCode: s.site_code, today, publicKeys });
    const over = st.asset_limit != null && s.active_assets > st.asset_limit;
    return { ...s, package: st.package || (st.mode === "legacy" ? "Existing installation" : st.mode === "trial" ? "Trial" : null), asset_limit: st.asset_limit ?? null, over_band: over, modules: st.modules };
  }).sort((a, b) => a.site_code.localeCompare(b.site_code));
  return { generated_at: new Date().toISOString(), active_users: users, sites: rows, warnings: rows.filter((r) => r.over_band).map((r) => `Site ${r.site_code} has ${r.active_assets} active machines; the licence covers ${r.asset_limit}.`) };
}
