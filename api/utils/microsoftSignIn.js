// IRONLOG/api/utils/microsoftSignIn.js — "Sign in with Microsoft" (Microsoft Entra ID) for one company.
//
// Each IronLog (AML, the pilot, later customers) trusts only its own company's
// Microsoft directory: an administrator enters the directory (tenant) ID, the
// app's client ID and its client secret under Admin → Microsoft sign-in. The
// round trip is the standard authorisation-code flow with PKCE; Microsoft's
// signed ID token is checked here (signature against Microsoft's published
// keys, issuer, audience, tenant, expiry, one-time nonce). Nobody gets in by
// having a Microsoft account alone: an IronLog admin invites the person by
// typing their work email on their IronLog user. At the first sign-in the
// permanent Microsoft account ID is remembered, so later email changes don't
// break the link. A Microsoft sign-in gets a normal IronLog session of 12 hours.
//
// Stage 3: an administrator can link Microsoft groups to IronLog roles (and,
// optionally, sites). People marked "roles come from Microsoft groups" get their
// roles and sites reset from their groups at every sign-in, and are refused when
// they are in no linked group. With "open to the whole company" on, anyone in a
// linked group gets an IronLog user at their first sign-in without an invite.
// "Microsoft sign-in only" users have no password (their PIN for the shared
// terminals still works).
// Plan: the "IronLog — Microsoft 365 (Entra ID) sign-in plan" doc.
import crypto from "node:crypto";
import { decryptJuneConnectorSecret, encryptJuneConnectorSecret } from "./juneOutlook.js";

export const MS_SESSION_HOURS = 12;
const REQUEST_MINUTES = 10;
const HANDOFF_SECONDS = 120;
const CLOCK_SKEW_S = 300;
const LOGIN_HOST = "https://login.microsoftonline.com";
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const EMAIL = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

const b64url = (buf) => Buffer.from(buf).toString("base64url");
const random = (n = 32) => b64url(crypto.randomBytes(n));

function hasColumn(db, table, name) {
  return db.prepare(`PRAGMA table_info(${table})`).all().some((r) => r.name === name);
}

export function ensureMicrosoftSignInTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS ms_signin_settings (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      enabled INTEGER NOT NULL DEFAULT 0,
      tenant_id TEXT,
      client_id TEXT,
      client_secret_enc TEXT,
      secret_expires_on TEXT,
      updated_at TEXT,
      updated_by TEXT
    );
    CREATE TABLE IF NOT EXISTS ms_signin_requests (
      state TEXT PRIMARY KEY,
      nonce TEXT NOT NULL,
      verifier TEXT NOT NULL,
      redirect_uri TEXT NOT NULL,
      test INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS ms_signin_handoffs (
      code TEXT PRIMARY KEY,
      token TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS ms_signin_group_roles (
      group_id TEXT PRIMARY KEY,
      label TEXT,
      roles_json TEXT NOT NULL,
      sites_json TEXT,
      updated_at TEXT,
      updated_by TEXT
    );
  `);
  if (!hasColumn(db, "ms_signin_settings", "open_to_company")) db.prepare(`ALTER TABLE ms_signin_settings ADD COLUMN open_to_company INTEGER NOT NULL DEFAULT 0`).run();
  try {
    for (const [col, type] of [["ms_email", "TEXT"], ["ms_oid", "TEXT"], ["ms_tid", "TEXT"], ["ms_linked_at", "TEXT"],
      ["ms_only", "INTEGER NOT NULL DEFAULT 0"], ["ms_roles_from_groups", "INTEGER NOT NULL DEFAULT 0"]]) {
      if (!hasColumn(db, "users", col)) db.prepare(`ALTER TABLE users ADD COLUMN ${col} ${type}`).run();
    }
  } catch {
    // users is created by auth.routes.js; it calls this again once the table exists.
  }
}

/** Settings without the secret. */
export function getMsSettings(db) {
  ensureMicrosoftSignInTables(db);
  const r = db.prepare(`SELECT * FROM ms_signin_settings WHERE id = 1`).get() || {};
  const configured = Boolean(r.tenant_id && r.client_id && r.client_secret_enc);
  return {
    enabled: Boolean(r.enabled) && configured,
    configured,
    tenant_id: r.tenant_id || "",
    client_id: r.client_id || "",
    has_secret: Boolean(r.client_secret_enc),
    secret_expires_on: r.secret_expires_on || "",
    open_to_company: Boolean(r.open_to_company),
    updated_at: r.updated_at || "",
    updated_by: r.updated_by || "",
  };
}

/** Days until the client secret runs out (null when no date was entered). */
export function secretDaysLeft(settings, today = new Date().toISOString().slice(0, 10)) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(settings.secret_expires_on || "")) return null;
  return Math.round((Date.parse(settings.secret_expires_on) - Date.parse(today)) / 864e5);
}

/** Saves what was sent (a blank secret keeps the saved one). Returns { ok, settings } or { ok:false, errors }. */
export function saveMsSettings(db, body = {}, by = "") {
  const cur = getMsSettings(db);
  const tenant = body.tenant_id != null ? String(body.tenant_id).trim().toLowerCase() : cur.tenant_id;
  const client = body.client_id != null ? String(body.client_id).trim().toLowerCase() : cur.client_id;
  const secret = String(body.client_secret || "").trim();
  const expires = body.secret_expires_on != null ? String(body.secret_expires_on).trim() : cur.secret_expires_on;
  const enabled = body.enabled != null ? Boolean(body.enabled) : cur.enabled;
  const open = body.open_to_company != null ? Boolean(body.open_to_company) : cur.open_to_company;
  const errors = [];
  if (open && !listGroupRoles(db).length) errors.push("Link at least one Microsoft group to an IronLog role before opening sign-in to the whole company.");
  if (tenant && !GUID.test(tenant)) errors.push("Directory (tenant) ID must look like 1a2b3c4d-1234-5678-9abc-1234567890ab (from the app's overview page in Entra).");
  if (client && !GUID.test(client)) errors.push("Application (client) ID must look like 1a2b3c4d-1234-5678-9abc-1234567890ab (from the app's overview page in Entra).");
  if (secret && (secret.length < 16 || secret.length > 200 || GUID.test(secret))) errors.push("That does not look like a client secret. Copy the secret's Value (not its Secret ID) under Certificates & secrets.");
  if (expires && !/^\d{4}-\d{2}-\d{2}$/.test(expires)) errors.push("Secret expiry date must be a date like 2028-10-16.");
  if (enabled && !(tenant && client && (secret || cur.has_secret))) errors.push("Enter the directory ID, client ID and client secret before switching Microsoft sign-in on.");
  if (errors.length) return { ok: false, errors };
  const encSecret = secret ? encryptJuneConnectorSecret(secret) : null;
  db.prepare(`
    INSERT INTO ms_signin_settings (id, enabled, tenant_id, client_id, client_secret_enc, secret_expires_on, open_to_company, updated_at, updated_by)
    VALUES (1, ?, ?, ?, ?, ?, ?, datetime('now'), ?)
    ON CONFLICT(id) DO UPDATE SET enabled = excluded.enabled, tenant_id = excluded.tenant_id, client_id = excluded.client_id,
      client_secret_enc = COALESCE(excluded.client_secret_enc, ms_signin_settings.client_secret_enc),
      secret_expires_on = excluded.secret_expires_on, open_to_company = excluded.open_to_company,
      updated_at = excluded.updated_at, updated_by = excluded.updated_by
  `).run(enabled ? 1 : 0, tenant || null, client || null, encSecret, expires || null, open ? 1 : 0, by || null);
  return { ok: true, settings: getMsSettings(db) };
}

function clientSecret(db) {
  const r = db.prepare(`SELECT client_secret_enc FROM ms_signin_settings WHERE id = 1`).get();
  return r?.client_secret_enc ? decryptJuneConnectorSecret(r.client_secret_enc) : null;
}

/** The address Microsoft sends people back to; IT enters exactly this in Entra. */
export function redirectUriFor(base) {
  return `${String(base).replace(/\/+$/, "")}/api/sso/microsoft/callback`;
}

/** Starts a sign-in: remembers the one-time state, nonce and PKCE verifier, returns Microsoft's address. */
export function startSignIn(db, { redirectUri, test = false } = {}) {
  const s = getMsSettings(db);
  if (!s.configured || (!s.enabled && !test)) throw Object.assign(new Error("Microsoft sign-in is not switched on."), { code: "not_enabled" });
  db.prepare(`DELETE FROM ms_signin_requests WHERE created_at < datetime('now', '-${REQUEST_MINUTES} minutes')`).run();
  const state = random();
  const nonce = random();
  const verifier = random(48);
  db.prepare(`INSERT INTO ms_signin_requests (state, nonce, verifier, redirect_uri, test) VALUES (?, ?, ?, ?, ?)`).run(state, nonce, verifier, redirectUri, test ? 1 : 0);
  const q = new URLSearchParams({
    client_id: s.client_id,
    response_type: "code",
    redirect_uri: redirectUri,
    response_mode: "query",
    scope: "openid profile email",
    state,
    nonce,
    code_challenge: b64url(crypto.createHash("sha256").update(verifier).digest()),
    code_challenge_method: "S256",
    prompt: "select_account",
  });
  return `${LOGIN_HOST}/${s.tenant_id}/oauth2/v2.0/authorize?${q}`;
}

/** The pending request for this state, used once (null when unknown or older than 10 minutes). */
export function takeRequest(db, state) {
  const r = db.prepare(`SELECT * FROM ms_signin_requests WHERE state = ? AND created_at >= datetime('now', '-${REQUEST_MINUTES} minutes')`).get(String(state || ""));
  db.prepare(`DELETE FROM ms_signin_requests WHERE state = ?`).run(String(state || ""));
  return r || null;
}

const fail = (code, message) => Object.assign(new Error(message), { code });

/** Swaps the one-time code for Microsoft's ID token. */
export async function exchangeCode(db, { code, verifier, redirectUri, fetchImpl = fetch }) {
  const s = getMsSettings(db);
  const secret = clientSecret(db);
  if (!secret) throw fail("not_configured", "The client secret could not be read. Enter it again under Admin → Microsoft sign-in.");
  const res = await fetchImpl(`${LOGIN_HOST}/${s.tenant_id}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: s.client_id, client_secret: secret, grant_type: "authorization_code",
      code: String(code || ""), redirect_uri: redirectUri, code_verifier: verifier, scope: "openid profile email",
    }).toString(),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !data.id_token) {
    // Microsoft's error codes (never the secret) help IT fix the app registration.
    const why = String(data.error_description || data.error || `HTTP ${res.status}`).split("\n")[0].slice(0, 300);
    throw fail("token_exchange", `Microsoft did not accept the sign-in: ${why}`);
  }
  return data.id_token;
}

let keyCache = { tenant: "", at: 0, keys: [] };
/** Microsoft's published signing keys for the tenant, cached for an hour. */
export async function signingKeys(tenantId, { fetchImpl = fetch, now = Date.now(), refresh = false } = {}) {
  if (!refresh && keyCache.tenant === tenantId && now - keyCache.at < 3600e3 && keyCache.keys.length) return keyCache.keys;
  const res = await fetchImpl(`${LOGIN_HOST}/${tenantId}/discovery/v2.0/keys`);
  const data = await res.json().catch(() => ({}));
  if (!res.ok || !Array.isArray(data.keys)) throw fail("keys", "Could not fetch Microsoft's signing keys.");
  keyCache = { tenant: tenantId, at: now, keys: data.keys };
  return data.keys;
}
export function clearKeyCache() {
  keyCache = { tenant: "", at: 0, keys: [] };
}

/** Checks Microsoft's ID token completely; returns its claims or throws with a code. */
export async function verifyIdToken(idToken, { tenantId, clientId, nonce, keys, now = Date.now() }) {
  const parts = String(idToken || "").split(".");
  if (parts.length !== 3) throw fail("bad_token", "The reply from Microsoft was not a valid token.");
  let header;
  let claims;
  try {
    header = JSON.parse(Buffer.from(parts[0], "base64url").toString("utf8"));
    claims = JSON.parse(Buffer.from(parts[1], "base64url").toString("utf8"));
  } catch {
    throw fail("bad_token", "The reply from Microsoft was not a valid token.");
  }
  if (header.alg !== "RS256") throw fail("bad_token", "The reply from Microsoft used an unexpected signature type.");
  // Microsoft rotates its keys: an unknown key ID fetches the list again once.
  const jwk = (await keys(false)).find((k) => k.kid === header.kid) || (await keys(true)).find((k) => k.kid === header.kid);
  if (!jwk) throw fail("bad_signature", "The reply was not signed by Microsoft.");
  const key = crypto.createPublicKey({ key: { kty: jwk.kty, n: jwk.n, e: jwk.e }, format: "jwk" });
  const ok = crypto.verify("RSA-SHA256", Buffer.from(`${parts[0]}.${parts[1]}`), key, Buffer.from(parts[2], "base64url"));
  if (!ok) throw fail("bad_signature", "The reply was not signed by Microsoft.");
  const t = Math.floor(now / 1000);
  if (claims.tid !== tenantId || claims.iss !== `${LOGIN_HOST}/${tenantId}/v2.0`) throw fail("wrong_tenant", "This Microsoft account belongs to a different company.");
  if (claims.aud !== clientId) throw fail("bad_token", "The reply was meant for a different app.");
  if (!(Number(claims.exp) > t - CLOCK_SKEW_S) || (claims.nbf && Number(claims.nbf) > t + CLOCK_SKEW_S)) throw fail("expired", "The sign-in took too long. Please try again.");
  if (!nonce || claims.nonce !== nonce) throw fail("replayed", "This sign-in link was already used. Please try again.");
  if (!claims.oid) throw fail("bad_token", "The reply from Microsoft did not name the account.");
  return claims;
}

/** The work email Microsoft gives for the account, lower case. */
export function claimEmail(claims) {
  const e = String(claims.email || claims.preferred_username || claims.upn || "").trim().toLowerCase();
  return EMAIL.test(e) ? e : "";
}

// ---- Microsoft groups → IronLog roles and sites

const parseList = (raw) => {
  try {
    const v = JSON.parse(raw || "[]");
    return Array.isArray(v) ? v.map((x) => String(x)) : [];
  } catch {
    return [];
  }
};

/** The linked groups: [{ group_id, label, roles, sites }] (sites empty = all sites). */
export function listGroupRoles(db) {
  ensureMicrosoftSignInTables(db);
  return db.prepare(`SELECT * FROM ms_signin_group_roles ORDER BY COALESCE(label, group_id)`).all()
    .map((r) => ({ group_id: r.group_id, label: r.label || "", roles: parseList(r.roles_json), sites: parseList(r.sites_json) }));
}

/**
 * Replaces the linked groups. validRoles: IronLog's roles; checkSites(list) → { ok, sites } or { ok:false, error }.
 * Returns { ok, groups } or { ok:false, errors }.
 */
export function saveGroupRoles(db, rows, { validRoles = [], checkSites = (l) => ({ ok: true, sites: l.length ? l : null }), by = "" } = {}) {
  ensureMicrosoftSignInTables(db);
  if (!Array.isArray(rows)) return { ok: false, errors: ["Groups must be a list."] };
  if (rows.length > 100) return { ok: false, errors: ["At most 100 linked groups."] };
  const errors = [];
  const clean = [];
  const seen = new Set();
  rows.forEach((r, i) => {
    const n = `Group ${i + 1}`;
    const id = String(r?.group_id || "").trim().toLowerCase();
    if (!GUID.test(id)) return errors.push(`${n}: the group's Object ID must look like 1a2b3c4d-1234-5678-9abc-1234567890ab (Entra → Groups → the group → Overview).`);
    if (seen.has(id)) return errors.push(`${n}: this group is listed twice.`);
    seen.add(id);
    const roles = [...new Set((Array.isArray(r.roles) ? r.roles : [r.roles]).map((x) => String(x || "").trim().toLowerCase()).filter(Boolean))];
    const bad = roles.filter((x) => !validRoles.includes(x));
    if (!roles.length || bad.length) return errors.push(`${n}: choose at least one IronLog role${bad.length ? ` (not a role: ${bad.join(", ")})` : ""}.`);
    const siteList = (Array.isArray(r.sites) ? r.sites : String(r.sites || "").split(",")).map((x) => String(x).trim()).filter(Boolean);
    const sites = siteList.length ? checkSites(siteList) : { ok: true, sites: null };
    if (!sites.ok) return errors.push(`${n}: ${sites.error}`);
    clean.push({ group_id: id, label: String(r.label || "").trim().slice(0, 80), roles, sites: sites.sites || [] });
  });
  if (errors.length) return { ok: false, errors };
  db.transaction(() => {
    db.prepare(`DELETE FROM ms_signin_group_roles`).run();
    const ins = db.prepare(`INSERT INTO ms_signin_group_roles (group_id, label, roles_json, sites_json, updated_at, updated_by) VALUES (?, ?, ?, ?, datetime('now'), ?)`);
    for (const g of clean) ins.run(g.group_id, g.label || null, JSON.stringify(g.roles), g.sites.length ? JSON.stringify(g.sites) : null, by || null);
    // Open sign-in needs at least one linked group.
    if (!clean.length) db.prepare(`UPDATE ms_signin_settings SET open_to_company = 0 WHERE id = 1`).run();
  })();
  return { ok: true, groups: listGroupRoles(db) };
}

/**
 * What the person's Microsoft groups give them: { overage, groups (linked ones they are in), roles, sites }.
 * sites is null (all sites) when any of their groups is not limited to sites. Overage: the person is in
 * too many groups for Microsoft to list them in the sign-in (IT should send only the app's assigned groups).
 */
export function groupAccess(db, claims) {
  const overage = Boolean(claims?._claim_names?.groups || claims?.hasgroups);
  const theirs = new Set((Array.isArray(claims?.groups) ? claims.groups : []).map((g) => String(g).toLowerCase()));
  const groups = listGroupRoles(db).filter((g) => theirs.has(g.group_id));
  const roles = [...new Set(groups.flatMap((g) => g.roles))];
  const sites = !groups.length || groups.some((g) => !g.sites.length) ? null : [...new Set(groups.flatMap((g) => g.sites))].sort();
  return { overage, groups, roles, sites };
}

const TOO_MANY_GROUPS = "You are in too many Microsoft groups for the sign-in to list them. Ask IT to set the IronLog app's groups claim to 'Groups assigned to the application'.";

/** A free username from the work email ("Ana.Silva@aml.co.mz" → "ana.silva", then "ana.silva2" …). */
export function usernameFromEmail(db, email) {
  const base = (String(email).split("@")[0].toLowerCase().replace(/[^a-z0-9._-]+/g, "").replace(/^[._-]+/, "") || "user").slice(0, 50).padEnd(2, "0");
  const taken = (u) => db.prepare(`SELECT 1 FROM users WHERE LOWER(username) = ?`).get(u);
  if (!taken(base)) return base;
  for (let i = 2; i < 1000; i++) if (!taken(`${base}${i}`)) return `${base}${i}`;
  return `${base}-${crypto.randomBytes(3).toString("hex")}`;
}

const primaryRole = (roles, validRoles) => validRoles.find((r) => roles.includes(r)) || roles[0];

/**
 * The IronLog user for these claims: linked, invited (linked now), or — with "open to the whole company" —
 * created from their Microsoft groups. Users whose roles come from groups get their roles and sites reset.
 * Returns the user row plus { linked_now, created, roles_changed }.
 */
export function matchUser(db, claims, { validRoles = [] } = {}) {
  ensureMicrosoftSignInTables(db);
  let user = db.prepare(`SELECT * FROM users WHERE ms_tid = ? AND ms_oid = ?`).get(claims.tid, claims.oid);
  let linkedNow = false;
  let created = false;
  if (!user) {
    const email = claimEmail(claims);
    const invited = email ? db.prepare(`SELECT * FROM users WHERE LOWER(TRIM(ms_email)) = ?`).all(email) : [];
    if (invited.length > 1) throw fail("not_invited", "This Microsoft account has not been invited to IronLog.");
    if (invited.length === 1) {
      user = invited[0];
      if (user.ms_oid && user.ms_oid !== claims.oid) throw fail("not_invited", "This IronLog user is already linked to a different Microsoft account.");
      db.prepare(`UPDATE users SET ms_oid = ?, ms_tid = ?, ms_linked_at = datetime('now') WHERE id = ?`).run(claims.oid, claims.tid, user.id);
      user = { ...user, ms_oid: claims.oid, ms_tid: claims.tid };
      linkedNow = true;
    } else {
      if (!getMsSettings(db).open_to_company || !email) throw fail("not_invited", "This Microsoft account has not been invited to IronLog.");
      const access = groupAccess(db, claims);
      if (access.overage) throw fail("too_many_groups", TOO_MANY_GROUPS);
      if (!access.roles.length) throw fail("no_group", "You are not in a Microsoft group that has access to IronLog. Ask your IronLog administrator.");
      const username = usernameFromEmail(db, email);
      const id = db.prepare(`
        INSERT INTO users (username, full_name, role, active, roles_json, allowed_locations, ms_email, ms_oid, ms_tid, ms_linked_at, ms_only, ms_roles_from_groups)
        VALUES (?, ?, ?, 1, ?, ?, ?, ?, ?, datetime('now'), 1, 1)
      `).run(username, String(claims.name || "").trim().slice(0, 120) || null, primaryRole(access.roles, validRoles), JSON.stringify(access.roles),
        access.sites ? JSON.stringify(access.sites) : null, email, claims.oid, claims.tid).lastInsertRowid;
      user = db.prepare(`SELECT * FROM users WHERE id = ?`).get(id);
      created = true;
    }
  }
  if (Number(user.active) !== 1) throw fail("inactive", "This IronLog user is switched off. Ask your IronLog administrator.");
  let rolesChanged = null;
  if (Number(user.ms_roles_from_groups) === 1 && !created) {
    const access = groupAccess(db, claims);
    if (access.overage) throw fail("too_many_groups", TOO_MANY_GROUPS);
    if (!access.roles.length) throw fail("no_group", "You are no longer in a Microsoft group that has access to IronLog. Ask your IronLog administrator.");
    const before = { roles: parseList(user.roles_json), sites: user.allowed_locations ? parseList(user.allowed_locations) : null };
    const after = { roles: access.roles, sites: access.sites };
    if (JSON.stringify([...before.roles].sort()) !== JSON.stringify([...after.roles].sort()) || JSON.stringify(before.sites) !== JSON.stringify(after.sites)) {
      db.prepare(`UPDATE users SET role = ?, roles_json = ?, allowed_locations = ? WHERE id = ?`)
        .run(primaryRole(after.roles, validRoles), JSON.stringify(after.roles), after.sites ? JSON.stringify(after.sites) : null, user.id);
      rolesChanged = { from: before, to: after };
      user = db.prepare(`SELECT * FROM users WHERE id = ?`).get(user.id);
    }
  }
  return { ...user, linked_now: linkedNow, created, roles_changed: rolesChanged };
}

/** A normal IronLog session (12 hours) plus a one-time handoff code for the browser (2 minutes). */
export function createSessionHandoff(db, userId) {
  const token = crypto.randomBytes(32).toString("hex");
  db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, ?, datetime('now', '+${MS_SESSION_HOURS} hours'))`).run(token, userId);
  db.prepare(`DELETE FROM ms_signin_handoffs WHERE created_at < datetime('now', '-${HANDOFF_SECONDS} seconds')`).run();
  const code = random();
  db.prepare(`INSERT INTO ms_signin_handoffs (code, token) VALUES (?, ?)`).run(code, token);
  return code;
}

/** The session token for a handoff code, once. */
export function takeHandoff(db, code) {
  ensureMicrosoftSignInTables(db);
  const r = db.prepare(`SELECT token FROM ms_signin_handoffs WHERE code = ? AND created_at >= datetime('now', '-${HANDOFF_SECONDS} seconds')`).get(String(code || ""));
  db.prepare(`DELETE FROM ms_signin_handoffs WHERE code = ?`).run(String(code || ""));
  return r?.token || null;
}

/** Sets or clears the work email an admin invites a user with; a new email unlinks the old account. */
export function setUserMicrosoftEmail(db, username, raw) {
  ensureMicrosoftSignInTables(db);
  const email = String(raw || "").trim().toLowerCase();
  if (email && !EMAIL.test(email)) return { ok: false, error: "Microsoft account must be a work email address, like name@company.com." };
  const user = db.prepare(`SELECT id, ms_email FROM users WHERE username = ?`).get(username);
  if (!user) return { ok: false, error: "user not found" };
  if (email) {
    const other = db.prepare(`SELECT username FROM users WHERE LOWER(TRIM(ms_email)) = ? AND id <> ?`).get(email, user.id);
    if (other) return { ok: false, error: `${email} is already the Microsoft account of ${other.username}.` };
  }
  if ((user.ms_email || "").toLowerCase() !== email) {
    db.prepare(`UPDATE users SET ms_email = ?, ms_oid = NULL, ms_tid = NULL, ms_linked_at = NULL WHERE id = ?`).run(email || null, user.id);
  }
  return { ok: true };
}

/**
 * "Microsoft sign-in only" and "roles come from Microsoft groups" on one user. Both need a Microsoft
 * account on the user. Microsoft-only removes the password and any setup code (the PIN stays, for the
 * shared terminals), and needs Microsoft sign-in switched on. At least one active administrator always
 * keeps a password, so the company can still get in if Microsoft sign-in breaks.
 */
export function setUserMicrosoftAccess(db, username, { ms_only, roles_from_groups } = {}) {
  ensureMicrosoftSignInTables(db);
  const user = db.prepare(`SELECT * FROM users WHERE username = ?`).get(username);
  if (!user) return { ok: false, error: "user not found" };
  const only = ms_only == null ? Number(user.ms_only) === 1 : Boolean(ms_only);
  const fromGroups = roles_from_groups == null ? Number(user.ms_roles_from_groups) === 1 : Boolean(roles_from_groups);
  if ((only || fromGroups) && !String(user.ms_email || "").trim() && !user.ms_oid) {
    return { ok: false, error: "Enter the person's Microsoft account (work email) first." };
  }
  if (only && Number(user.ms_only) !== 1) {
    if (!getMsSettings(db).enabled) return { ok: false, error: "Switch Microsoft sign-in on (Admin → Microsoft sign-in) before making anyone Microsoft-only." };
    const roles = parseList(user.roles_json);
    if (roles.includes("admin") || user.role === "admin") {
      const others = db.prepare(`
        SELECT COUNT(*) AS n FROM users
        WHERE id <> ? AND COALESCE(active, 1) = 1 AND COALESCE(ms_only, 0) = 0
          AND password_hash IS NOT NULL AND length(trim(password_hash)) > 0
          AND (role = 'admin' OR roles_json LIKE '%"admin"%')
      `).get(user.id).n;
      if (!others) return { ok: false, error: "At least one administrator must keep a password, in case Microsoft sign-in stops working. Give another administrator a password first." };
    }
  }
  db.prepare(`UPDATE users SET ms_only = ?, ms_roles_from_groups = ? WHERE id = ?`).run(only ? 1 : 0, fromGroups ? 1 : 0, user.id);
  if (only) db.prepare(`UPDATE users SET password_hash = NULL, setup_code_hash = NULL, setup_code_expires_at = NULL WHERE id = ?`).run(user.id);
  return { ok: true, ms_only: only, roles_from_groups: fromGroups };
}

/** For the admin banner: is the client secret about to run out? */
export function secretWarning(db, today) {
  const s = getMsSettings(db);
  const days = secretDaysLeft(s, today);
  if (!s.enabled || days == null || days > 30) return null;
  return {
    days_left: days,
    expires_on: s.secret_expires_on,
    message: days < 0
      ? `The Microsoft sign-in secret ran out on ${s.secret_expires_on}: Microsoft sign-in has stopped working. Create a new secret in Entra and save it under Admin → Microsoft sign-in.`
      : `The Microsoft sign-in secret runs out in ${days} day${days === 1 ? "" : "s"} (${s.secret_expires_on}). Create a new secret in Entra and save it under Admin → Microsoft sign-in before then.`,
  };
}
