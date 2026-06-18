// IRONLOG — Firebase Cloud Messaging (FCM HTTP v1) push dispatcher
import fs from "fs";
import path from "path";
import { db } from "../db/client.js";

let googleAuth = null;
let serviceAccount = null;
let projectId = null;

function loadServiceAccount() {
  if (serviceAccount) return serviceAccount;
  const jsonInline = String(process.env.FCM_SERVICE_ACCOUNT_JSON || "").trim();
  const jsonPath = String(process.env.FCM_SERVICE_ACCOUNT_PATH || "").trim();
  try {
    if (jsonInline) {
      serviceAccount = JSON.parse(jsonInline);
    } else if (jsonPath && fs.existsSync(jsonPath)) {
      serviceAccount = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
    }
  } catch (err) {
    console.error("[push] invalid FCM service account:", err?.message || err);
    serviceAccount = null;
  }
  if (serviceAccount) {
    projectId =
      String(process.env.FCM_PROJECT_ID || serviceAccount.project_id || "").trim() || null;
  }
  return serviceAccount;
}

export function isPushConfigured() {
  return Boolean(loadServiceAccount() && projectId);
}

export function getNotifyApkUrl(req) {
  const envUrl = String(process.env.IRONLOG_NOTIFY_APK_URL || "").trim();
  if (envUrl) return envUrl;
  const proto = String(req?.headers?.["x-forwarded-proto"] || "http").split(",")[0].trim() || "http";
  const host = String(req?.headers?.["x-forwarded-host"] || req?.headers?.host || "").split(",")[0].trim();
  if (host) return `${proto}://${host}/web/downloads/ironlog-notify.apk`;
  return "/web/downloads/ironlog-notify.apk";
}

/** Expo internal build page — scan QR on phone to install without hosting APK. */
export function getNotifyExpoInstallUrl() {
  return (
    String(process.env.IRONLOG_NOTIFY_EXPO_INSTALL_URL || "").trim() ||
    "https://expo.dev/accounts/jakes84/projects/ironlog-notify/builds"
  );
}

async function getAccessToken() {
  const creds = loadServiceAccount();
  if (!creds) return null;
  if (!googleAuth) {
    const mod = await import("google-auth-library");
    const { GoogleAuth } = mod;
    googleAuth = new GoogleAuth({
      credentials: creds,
      scopes: ["https://www.googleapis.com/auth/firebase.messaging"],
    });
  }
  const client = await googleAuth.getClient();
  const res = await client.getAccessToken();
  return res?.token || null;
}

function ensurePushTables() {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS push_device_tokens (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT NOT NULL,
      token TEXT NOT NULL UNIQUE,
      platform TEXT NOT NULL DEFAULT 'android',
      device_label TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      last_seen_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`
    CREATE INDEX IF NOT EXISTS idx_push_tokens_username ON push_device_tokens(username)
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS push_notification_log (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT,
      kind TEXT NOT NULL,
      title TEXT NOT NULL,
      body TEXT NOT NULL,
      data_json TEXT,
      sent_at TEXT NOT NULL DEFAULT (datetime('now')),
      success INTEGER NOT NULL DEFAULT 0,
      error TEXT
    )
  `).run();
}

ensurePushTables();

function logNotification({ username, kind, title, body, data, success, error }) {
  try {
    db.prepare(`
      INSERT INTO push_notification_log (username, kind, title, body, data_json, success, error)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `).run(
      username || null,
      kind,
      title,
      body,
      data ? JSON.stringify(data) : null,
      success ? 1 : 0,
      error || null
    );
  } catch (err) {
    console.error("[push] log failed:", err?.message || err);
  }
}

export function listTokensForUsernames(usernames) {
  const names = Array.from(
    new Set(
      (Array.isArray(usernames) ? usernames : [])
        .map((u) => String(u || "").trim().toLowerCase())
        .filter(Boolean)
    )
  );
  if (!names.length) return [];
  const marks = names.map(() => "?").join(",");
  return db
    .prepare(`
      SELECT username, token, platform, device_label
      FROM push_device_tokens
      WHERE lower(username) IN (${marks})
    `)
    .all(...names);
}

export function listTokensForRoles(roles) {
  const wanted = Array.from(
    new Set(
      (Array.isArray(roles) ? roles : [])
        .map((r) => String(r || "").trim().toLowerCase())
        .filter(Boolean)
    )
  );
  if (!wanted.length) return [];
  const users = db
    .prepare(`
      SELECT username, role, roles_json, active
      FROM users
      WHERE active = 1
    `)
    .all();
  const usernames = users
    .filter((u) => {
      let roles = [];
      try {
        const parsed = u.roles_json ? JSON.parse(u.roles_json) : [];
        if (Array.isArray(parsed)) roles = parsed;
      } catch {}
      if (!roles.length && u.role) roles = [u.role];
      roles = roles.map((r) => String(r || "").trim().toLowerCase());
      return roles.some((r) => wanted.includes(r));
    })
    .map((u) => String(u.username || "").trim().toLowerCase())
    .filter(Boolean);
  return listTokensForUsernames(usernames);
}

async function sendFcmToToken({ token, title, body, data = {} }) {
  if (!isPushConfigured()) {
    return { ok: false, skipped: true, error: "FCM not configured" };
  }
  const accessToken = await getAccessToken();
  if (!accessToken) return { ok: false, error: "FCM auth failed" };

  const payload = {
    message: {
      token,
      notification: { title, body },
      data: Object.fromEntries(
        Object.entries(data).map(([k, v]) => [String(k), String(v ?? "")])
      ),
      android: { priority: "HIGH" },
    },
  };

  const res = await fetch(
    `https://fcm.googleapis.com/v1/projects/${projectId}/messages:send`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    }
  );

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    return { ok: false, error: text || `HTTP ${res.status}` };
  }
  return { ok: true };
}

export async function sendPushToUsernames({ usernames, kind, title, body, data }) {
  if (!isPushConfigured()) return { sent: 0, skipped: true };
  const tokens = listTokensForUsernames(usernames);
  let sent = 0;
  for (const row of tokens) {
    const result = await sendFcmToToken({ token: row.token, title, body, data });
    logNotification({
      username: row.username,
      kind,
      title,
      body,
      data,
      success: result.ok,
      error: result.error,
    });
    if (result.ok) sent += 1;
    if (result.error && /not.+registered|invalid.+token|UNREGISTERED/i.test(result.error)) {
      try {
        db.prepare(`DELETE FROM push_device_tokens WHERE token = ?`).run(row.token);
      } catch {}
    }
  }
  return { sent, total: tokens.length };
}

export async function sendPushToRoles({ roles, kind, title, body, data }) {
  if (!isPushConfigured()) return { sent: 0, skipped: true };
  const tokens = listTokensForRoles(roles);
  let sent = 0;
  for (const row of tokens) {
    const result = await sendFcmToToken({ token: row.token, title, body, data });
    logNotification({
      username: row.username,
      kind,
      title,
      body,
      data,
      success: result.ok,
      error: result.error,
    });
    if (result.ok) sent += 1;
    if (result.error && /not.+registered|invalid.+token|UNREGISTERED/i.test(result.error)) {
      try {
        db.prepare(`DELETE FROM push_device_tokens WHERE token = ?`).run(row.token);
      } catch {}
    }
  }
  return { sent, total: tokens.length };
}

export async function notifyBreakdownCreated({ assetCode, description, breakdownId, workOrderId }) {
  const asset = String(assetCode || "").trim() || "Asset";
  const desc = String(description || "Breakdown reported").trim().slice(0, 120);
  const title = `Breakdown — ${asset}`;
  const body = desc;
  const data = {
    kind: "breakdown",
    breakdown_id: String(breakdownId || ""),
    wo_id: String(workOrderId || ""),
    asset_code: asset,
  };
  return sendPushToRoles({
    roles: ["artisan", "supervisor", "admin"],
    kind: "breakdown",
    title,
    body,
    data,
  });
}

export async function notifyWorkOrderAssigned({ workOrderId, assignedUsername, assetCode, source }) {
  const wo = String(workOrderId || "").trim();
  const user = String(assignedUsername || "").trim();
  if (!user) return { sent: 0 };
  const asset = String(assetCode || "").trim();
  const title = `WO #${wo} assigned to you`;
  const parts = [];
  if (asset) parts.push(asset);
  if (source) parts.push(String(source));
  const body = parts.length ? parts.join(" · ") : "Tap to open work order";
  const data = {
    kind: "work_order_assigned",
    wo_id: wo,
    asset_code: asset,
  };
  return sendPushToUsernames({
    usernames: [user],
    kind: "work_order_assigned",
    title,
    body,
    data,
  });
}
