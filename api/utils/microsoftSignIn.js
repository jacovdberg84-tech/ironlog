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
  `);
  try {
    for (const [col, type] of [["ms_email", "TEXT"], ["ms_oid", "TEXT"], ["ms_tid", "TEXT"], ["ms_linked_at", "TEXT"]]) {
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
  const errors = [];
  if (tenant && !GUID.test(tenant)) errors.push("Directory (tenant) ID must look like 1a2b3c4d-1234-5678-9abc-1234567890ab (from the app's overview page in Entra).");
  if (client && !GUID.test(client)) errors.push("Application (client) ID must look like 1a2b3c4d-1234-5678-9abc-1234567890ab (from the app's overview page in Entra).");
  if (secret && (secret.length < 16 || secret.length > 200 || GUID.test(secret))) errors.push("That does not look like a client secret. Copy the secret's Value (not its Secret ID) under Certificates & secrets.");
  if (expires && !/^\d{4}-\d{2}-\d{2}$/.test(expires)) errors.push("Secret expiry date must be a date like 2028-10-16.");
  if (enabled && !(tenant && client && (secret || cur.has_secret))) errors.push("Enter the directory ID, client ID and client secret before switching Microsoft sign-in on.");
  if (errors.length) return { ok: false, errors };
  const encSecret = secret ? encryptJuneConnectorSecret(secret) : null;
  db.prepare(`
    INSERT INTO ms_signin_settings (id, enabled, tenant_id, client_id, client_secret_enc, secret_expires_on, updated_at, updated_by)
    VALUES (1, ?, ?, ?, ?, ?, datetime('now'), ?)
    ON CONFLICT(id) DO UPDATE SET enabled = excluded.enabled, tenant_id = excluded.tenant_id, client_id = excluded.client_id,
      client_secret_enc = COALESCE(excluded.client_secret_enc, ms_signin_settings.client_secret_enc),
      secret_expires_on = excluded.secret_expires_on, updated_at = excluded.updated_at, updated_by = excluded.updated_by
  `).run(enabled ? 1 : 0, tenant || null, client || null, encSecret, expires || null, by || null);
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

/** The invited, active IronLog user for these claims; links the Microsoft account on first use. */
export function matchUser(db, claims) {
  ensureMicrosoftSignInTables(db);
  let user = db.prepare(`SELECT * FROM users WHERE ms_tid = ? AND ms_oid = ?`).get(claims.tid, claims.oid);
  if (!user) {
    const email = claimEmail(claims);
    const invited = email ? db.prepare(`SELECT * FROM users WHERE LOWER(TRIM(ms_email)) = ?`).all(email) : [];
    if (invited.length !== 1) throw fail("not_invited", "This Microsoft account has not been invited to IronLog.");
    user = invited[0];
    if (user.ms_oid && user.ms_oid !== claims.oid) throw fail("not_invited", "This IronLog user is already linked to a different Microsoft account.");
    db.prepare(`UPDATE users SET ms_oid = ?, ms_tid = ?, ms_linked_at = datetime('now') WHERE id = ?`).run(claims.oid, claims.tid, user.id);
    user = { ...user, ms_oid: claims.oid, ms_tid: claims.tid, linked_now: true };
  }
  if (Number(user.active) !== 1) throw fail("inactive", "This IronLog user is switched off. Ask your IronLog administrator.");
  return user;
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
