// IRONLOG/api/utils/juneOutlook.js
// Private, user-authorised Microsoft Outlook connector for June. Tokens are
// encrypted at rest and never leave the API server.
import crypto from "node:crypto";
import { db } from "../db/client.js";

const MICROSOFT_LOGIN_BASE = "https://login.microsoftonline.com";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const TOKEN_SKEW_MS = 120_000;
const STATE_TTL_MS = 10 * 60_000;
const GRAPH_TIMEOUT_MS = 20_000;
const OUTLOOK_SCOPES = ["openid", "profile", "offline_access", "User.Read", "Mail.Read", "Calendars.Read"];

function text(value, max = 500) {
  return String(value || "").trim().slice(0, max);
}

function contextValue(value, fallback) {
  return text(value, 160).toLowerCase() || fallback;
}

function connectorKey() {
  // Prefer a dedicated connector key. IRONLOG_AUTH_SECRET is an acceptable
  // fallback for existing secured installations, but no default is allowed.
  const raw = text(process.env.JUNE_CONNECTOR_ENCRYPTION_SECRET || process.env.IRONLOG_AUTH_SECRET, 2000);
  if (raw.length < 32) return null;
  return crypto.createHash("sha256").update(raw).digest();
}

function encrypt(value) {
  const key = connectorKey();
  if (!key) throw new Error("June's Outlook token encryption secret is not configured.");
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  const encrypted = Buffer.concat([cipher.update(String(value || ""), "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return `${iv.toString("base64")}.${tag.toString("base64")}.${encrypted.toString("base64")}`;
}

function decrypt(value) {
  const raw = text(value, 50000);
  const key = connectorKey();
  if (!raw || !key) return null;
  const [ivB64, tagB64, encryptedB64] = raw.split(".");
  if (!ivB64 || !tagB64 || !encryptedB64) return null;
  try {
    const decipher = crypto.createDecipheriv("aes-256-gcm", key, Buffer.from(ivB64, "base64"));
    decipher.setAuthTag(Buffer.from(tagB64, "base64"));
    return Buffer.concat([decipher.update(Buffer.from(encryptedB64, "base64")), decipher.final()]).toString("utf8");
  } catch {
    return null;
  }
}

function config() {
  const clientId = text(process.env.JUNE_OUTLOOK_CLIENT_ID, 300);
  const clientSecret = text(process.env.JUNE_OUTLOOK_CLIENT_SECRET, 2000);
  const tenantId = text(process.env.JUNE_OUTLOOK_TENANT_ID, 300);
  const redirectUri = text(process.env.JUNE_OUTLOOK_REDIRECT_URI, 1200);
  const missing = [];
  if (!clientId) missing.push("client ID");
  if (!clientSecret) missing.push("client secret");
  if (!tenantId) missing.push("tenant ID");
  if (!redirectUri) missing.push("redirect URI");
  if (!connectorKey()) missing.push("token encryption secret");
  return {
    ready: missing.length === 0,
    clientId,
    clientSecret,
    tenantId,
    redirectUri,
    detail: missing.length
      ? `Outlook needs server configuration: ${missing.join(", ")}.`
      : "Connect Outlook to let June read your calendar and prioritise inbox messages.",
  };
}

function ensureTables() {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS june_outlook_connections (
      site_code TEXT NOT NULL,
      username TEXT NOT NULL COLLATE NOCASE,
      tenant_id TEXT NOT NULL,
      account_email TEXT,
      display_name TEXT,
      access_token_enc TEXT NOT NULL,
      refresh_token_enc TEXT NOT NULL,
      token_expires_at_ms INTEGER NOT NULL DEFAULT 0,
      scopes TEXT,
      status TEXT NOT NULL DEFAULT 'connected',
      connected_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      PRIMARY KEY (site_code, username)
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS june_outlook_oauth_states (
      state TEXT PRIMARY KEY,
      site_code TEXT NOT NULL,
      username TEXT NOT NULL COLLATE NOCASE,
      code_verifier_enc TEXT NOT NULL,
      expires_at_ms INTEGER NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
}

function contextKey(context = {}) {
  return {
    siteCode: contextValue(context.siteCode, "main"),
    user: text(context.user, 160) || "session-user",
  };
}

function connectionRow(context) {
  ensureTables();
  const { siteCode, user } = contextKey(context);
  return db.prepare(`
    SELECT site_code, username, tenant_id, account_email, display_name, access_token_enc, refresh_token_enc,
           token_expires_at_ms, scopes, status, connected_at, updated_at
    FROM june_outlook_connections
    WHERE site_code = ? AND username = ?
  `).get(siteCode, user) || null;
}

function publicConnectionStatus(context) {
  const cfg = config();
  if (!cfg.ready) {
    return { state: "not_configured", detail: cfg.detail, provider: "Microsoft Outlook", account: null };
  }
  const row = connectionRow(context);
  if (!row) {
    return { state: "not_connected", detail: cfg.detail, provider: "Microsoft Outlook", account: null };
  }
  if (String(row.status || "").toLowerCase() !== "connected") {
    return {
      state: "needs_reconnect",
      detail: "Outlook needs to be reconnected before June can use it.",
      provider: "Microsoft Outlook",
      account: text(row.account_email, 320) || null,
      display_name: text(row.display_name, 180) || null,
    };
  }
  return {
    state: "connected",
    detail: "June can read your calendar and identify priority inbox messages. She cannot send mail or create meetings.",
    provider: "Microsoft Outlook",
    account: text(row.account_email, 320) || null,
    display_name: text(row.display_name, 180) || null,
    connected_at: row.connected_at || null,
  };
}

function base64Url(value) {
  return Buffer.from(value).toString("base64url");
}

function completionUrl(cfg, state, message = "") {
  const target = new URL("/web/index.html", cfg.redirectUri);
  target.searchParams.set("june_outlook", state);
  if (message) target.searchParams.set("june_outlook_message", text(message, 220));
  return target.toString();
}

function stateFailureUrl(cfg, message) {
  return cfg?.ready ? completionUrl(cfg, "error", message) : null;
}

function cleanProviderError(payload, fallback) {
  return text(payload?.error_description || payload?.error?.message || payload?.error || fallback, 350);
}

async function requestJson(url, options = {}, timeoutMs = GRAPH_TIMEOUT_MS) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, { ...options, signal: controller.signal });
    const raw = await response.text();
    let payload = {};
    try { payload = raw ? JSON.parse(raw) : {}; } catch {}
    if (!response.ok) {
      const error = new Error(cleanProviderError(payload, `Microsoft returned HTTP ${response.status}.`));
      error.statusCode = response.status;
      error.providerPayload = payload;
      throw error;
    }
    return payload;
  } catch (error) {
    if (error?.name === "AbortError") {
      const timeoutError = new Error("Microsoft Outlook took too long to respond.");
      timeoutError.statusCode = 504;
      throw timeoutError;
    }
    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

async function requestToken(cfg, form) {
  const body = new URLSearchParams(form);
  return requestJson(`${MICROSOFT_LOGIN_BASE}/${encodeURIComponent(cfg.tenantId)}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });
}

function markReconnectRequired(context) {
  const { siteCode, user } = contextKey(context);
  ensureTables();
  db.prepare(`
    UPDATE june_outlook_connections
    SET status = 'needs_reconnect', updated_at = datetime('now')
    WHERE site_code = ? AND username = ?
  `).run(siteCode, user);
}

async function refreshAccessToken(context, row) {
  const cfg = config();
  if (!cfg.ready) throw new Error(cfg.detail);
  const refreshToken = decrypt(row?.refresh_token_enc);
  if (!refreshToken) {
    markReconnectRequired(context);
    throw new Error("June cannot unlock the saved Outlook connection. Reconnect Outlook to continue.");
  }
  let token;
  try {
    token = await requestToken(cfg, {
      client_id: cfg.clientId,
      client_secret: cfg.clientSecret,
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      scope: OUTLOOK_SCOPES.join(" "),
    });
  } catch (error) {
    markReconnectRequired(context);
    throw new Error(`June needs Outlook reconnected: ${cleanProviderError(error?.providerPayload, error?.message || "refresh failed")}`);
  }
  const accessToken = text(token?.access_token, 12000);
  if (!accessToken) {
    markReconnectRequired(context);
    throw new Error("Microsoft did not return an Outlook access token. Reconnect Outlook and try again.");
  }
  const expiresAt = Date.now() + Math.max(60, Number(token?.expires_in || 3600)) * 1000;
  const nextRefresh = text(token?.refresh_token, 12000);
  const { siteCode, user } = contextKey(context);
  db.prepare(`
    UPDATE june_outlook_connections
    SET access_token_enc = ?, refresh_token_enc = ?, token_expires_at_ms = ?, status = 'connected', updated_at = datetime('now')
    WHERE site_code = ? AND username = ?
  `).run(encrypt(accessToken), nextRefresh ? encrypt(nextRefresh) : row.refresh_token_enc, expiresAt, siteCode, user);
  return accessToken;
}

async function accessTokenFor(context, { forceRefresh = false } = {}) {
  const status = publicConnectionStatus(context);
  if (status.state !== "connected") throw new Error(status.detail);
  const row = connectionRow(context);
  const accessToken = !forceRefresh ? decrypt(row?.access_token_enc) : null;
  if (accessToken && Number(row?.token_expires_at_ms || 0) > Date.now() + TOKEN_SKEW_MS) return accessToken;
  return refreshAccessToken(context, row);
}

async function graphGet(context, path, headers = {}) {
  const call = async (forceRefresh) => {
    const token = await accessTokenFor(context, { forceRefresh });
    return requestJson(`${GRAPH_BASE}${path}`, {
      headers: { Authorization: `Bearer ${token}`, Accept: "application/json", ...headers },
    });
  };
  try {
    return await call(false);
  } catch (error) {
    if (Number(error?.statusCode) !== 401) throw error;
    return call(true);
  }
}

export function getOutlookConnectionStatus(context) {
  return publicConnectionStatus(context);
}

export function getOutlookSetupSummary() {
  const cfg = config();
  return {
    configured: cfg.ready,
    redirect_uri: cfg.ready ? cfg.redirectUri : null,
    scopes: ["User.Read", "Mail.Read", "Calendars.Read"],
    detail: cfg.detail,
  };
}

export function beginOutlookAuthorization(context) {
  const cfg = config();
  if (!cfg.ready) throw new Error(cfg.detail);
  ensureTables();
  db.prepare("DELETE FROM june_outlook_oauth_states WHERE expires_at_ms <= ?").run(Date.now());
  const state = base64Url(crypto.randomBytes(32));
  const codeVerifier = base64Url(crypto.randomBytes(48));
  const codeChallenge = crypto.createHash("sha256").update(codeVerifier).digest("base64url");
  const { siteCode, user } = contextKey(context);
  db.prepare(`
    INSERT INTO june_outlook_oauth_states (state, site_code, username, code_verifier_enc, expires_at_ms)
    VALUES (?, ?, ?, ?, ?)
  `).run(state, siteCode, user, encrypt(codeVerifier), Date.now() + STATE_TTL_MS);
  const authorize = new URL(`${MICROSOFT_LOGIN_BASE}/${encodeURIComponent(cfg.tenantId)}/oauth2/v2.0/authorize`);
  authorize.searchParams.set("client_id", cfg.clientId);
  authorize.searchParams.set("response_type", "code");
  authorize.searchParams.set("redirect_uri", cfg.redirectUri);
  authorize.searchParams.set("response_mode", "query");
  authorize.searchParams.set("scope", OUTLOOK_SCOPES.join(" "));
  authorize.searchParams.set("state", state);
  authorize.searchParams.set("code_challenge", codeChallenge);
  authorize.searchParams.set("code_challenge_method", "S256");
  authorize.searchParams.set("prompt", "select_account");
  return { authorize_url: authorize.toString(), expires_at: new Date(Date.now() + STATE_TTL_MS).toISOString() };
}

export async function completeOutlookAuthorization({ state, code, error, errorDescription }) {
  const cfg = config();
  if (!cfg.ready) return { ok: false, redirect_url: null, message: cfg.detail };
  ensureTables();
  const safeState = text(state, 300);
  const pending = safeState
    ? db.prepare(`SELECT state, site_code, username, code_verifier_enc, expires_at_ms FROM june_outlook_oauth_states WHERE state = ?`).get(safeState)
    : null;
  if (!pending || Number(pending.expires_at_ms || 0) <= Date.now()) {
    if (safeState) db.prepare("DELETE FROM june_outlook_oauth_states WHERE state = ?").run(safeState);
    return { ok: false, redirect_url: stateFailureUrl(cfg, "That Outlook connection request expired. Start again from June."), message: "That Outlook connection request expired. Start again from June." };
  }
  db.prepare("DELETE FROM june_outlook_oauth_states WHERE state = ?").run(safeState);
  if (error) {
    const message = `Outlook connection was not approved: ${text(errorDescription || error, 220)}`;
    return { ok: false, redirect_url: stateFailureUrl(cfg, message), message };
  }
  const authorizationCode = text(code, 8000);
  const codeVerifier = decrypt(pending.code_verifier_enc);
  if (!authorizationCode || !codeVerifier) {
    const message = "Outlook connection details were incomplete. Start the connection again from June.";
    return { ok: false, redirect_url: stateFailureUrl(cfg, message), message };
  }
  let token;
  try {
    token = await requestToken(cfg, {
      client_id: cfg.clientId,
      client_secret: cfg.clientSecret,
      grant_type: "authorization_code",
      code: authorizationCode,
      redirect_uri: cfg.redirectUri,
      code_verifier: codeVerifier,
      scope: OUTLOOK_SCOPES.join(" "),
    });
  } catch (requestError) {
    const message = `Microsoft could not complete the Outlook connection: ${cleanProviderError(requestError?.providerPayload, requestError?.message || "authorisation failed")}`;
    return { ok: false, redirect_url: stateFailureUrl(cfg, message), message };
  }
  const accessToken = text(token?.access_token, 12000);
  const refreshToken = text(token?.refresh_token, 12000);
  if (!accessToken || !refreshToken) {
    const message = "Microsoft did not grant the long-term Outlook access June needs. Check that offline access is allowed, then reconnect.";
    return { ok: false, redirect_url: stateFailureUrl(cfg, message), message };
  }
  let profile;
  try {
    profile = await requestJson(`${GRAPH_BASE}/me?$select=id,displayName,mail,userPrincipalName`, {
      headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json" },
    });
  } catch (profileError) {
    const message = `Outlook authorised the connection, but June could not read your account profile: ${cleanProviderError(profileError?.providerPayload, profileError?.message || "profile lookup failed")}`;
    return { ok: false, redirect_url: stateFailureUrl(cfg, message), message };
  }
  const accountEmail = text(profile?.mail || profile?.userPrincipalName, 320);
  const displayName = text(profile?.displayName, 180);
  const expiresAt = Date.now() + Math.max(60, Number(token?.expires_in || 3600)) * 1000;
  db.prepare(`
    INSERT INTO june_outlook_connections (
      site_code, username, tenant_id, account_email, display_name, access_token_enc, refresh_token_enc,
      token_expires_at_ms, scopes, status, connected_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'connected', datetime('now'), datetime('now'))
    ON CONFLICT(site_code, username) DO UPDATE SET
      tenant_id = excluded.tenant_id,
      account_email = excluded.account_email,
      display_name = excluded.display_name,
      access_token_enc = excluded.access_token_enc,
      refresh_token_enc = excluded.refresh_token_enc,
      token_expires_at_ms = excluded.token_expires_at_ms,
      scopes = excluded.scopes,
      status = 'connected',
      updated_at = datetime('now')
  `).run(
    contextValue(pending.site_code, "main"),
    text(pending.username, 160) || "session-user",
    cfg.tenantId,
    accountEmail || null,
    displayName || null,
    encrypt(accessToken),
    encrypt(refreshToken),
    expiresAt,
    OUTLOOK_SCOPES.join(" "),
  );
  const message = accountEmail ? `Outlook connected to ${accountEmail}.` : "Outlook connected.";
  return {
    ok: true,
    redirect_url: completionUrl(cfg, "connected", message),
    message,
    connection: publicConnectionStatus({ siteCode: pending.site_code, user: pending.username }),
  };
}

export function disconnectOutlook(context) {
  const { siteCode, user } = contextKey(context);
  ensureTables();
  const result = db.prepare("DELETE FROM june_outlook_connections WHERE site_code = ? AND username = ?").run(siteCode, user);
  return { disconnected: Number(result.changes || 0) > 0, connection: publicConnectionStatus(context) };
}

export async function getOutlookCalendarOverview(context) {
  const status = publicConnectionStatus(context);
  if (status.state !== "connected") {
    return { ...status, review_only: true, next_step: "Connect Outlook before June can read meetings." };
  }
  const start = new Date();
  const end = new Date(start.getTime() + 7 * 24 * 60 * 60 * 1000);
  const params = new URLSearchParams({
    startDateTime: start.toISOString(),
    endDateTime: end.toISOString(),
    "$select": "subject,start,end,location,organizer,isCancelled",
    "$orderby": "start/dateTime",
    "$top": "12",
  });
  const data = await graphGet(context, `/me/calendarView?${params.toString()}`, { Prefer: 'outlook.timezone="South Africa Standard Time"' });
  const meetings = Array.isArray(data?.value) ? data.value : [];
  return {
    state: "connected",
    review_only: true,
    account: status.account,
    period: { start: start.toISOString(), end: end.toISOString() },
    meetings: meetings.filter((item) => !item?.isCancelled).map((item) => ({
      subject: text(item?.subject, 180) || "(No subject)",
      start: text(item?.start?.dateTime, 80),
      end: text(item?.end?.dateTime, 80),
      location: text(item?.location?.displayName, 140) || null,
      organiser: text(item?.organizer?.emailAddress?.name || item?.organizer?.emailAddress?.address, 180) || null,
    })),
    next_step: "June can brief you on these meetings; scheduling remains approval-only and is not enabled.",
  };
}

export async function getOutlookPriorityEmails(context) {
  const status = publicConnectionStatus(context);
  if (status.state !== "connected") {
    return { ...status, review_only: true, next_step: "Connect Outlook before June can read priority inbox messages." };
  }
  const params = new URLSearchParams({
    "$select": "subject,from,receivedDateTime,importance,isRead",
    "$orderby": "receivedDateTime desc",
    "$top": "30",
  });
  const data = await graphGet(context, `/me/messages?${params.toString()}`);
  const messages = Array.isArray(data?.value) ? data.value : [];
  const ranked = messages
    .filter((item) => String(item?.importance || "normal").toLowerCase() === "high" || item?.isRead === false)
    .sort((a, b) => {
      const weight = (item) => String(item?.importance || "normal").toLowerCase() === "high" ? 0 : item?.isRead === false ? 1 : 2;
      return weight(a) - weight(b) || String(b?.receivedDateTime || "").localeCompare(String(a?.receivedDateTime || ""));
    })
    .slice(0, 10)
    .map((item) => ({
      subject: text(item?.subject, 180) || "(No subject)",
      from: text(item?.from?.emailAddress?.name || item?.from?.emailAddress?.address, 180) || "Unknown sender",
      received_at: text(item?.receivedDateTime, 80),
      importance: String(item?.importance || "normal").toLowerCase(),
      unread: item?.isRead === false,
    }));
  return {
    state: "connected",
    review_only: true,
    account: status.account,
    priority_messages: ranked,
    unread_count: messages.filter((item) => item?.isRead === false).length,
    note: "June reads message metadata only for this brief. She does not send mail or expose message bodies.",
  };
}
