// IRONLOG — push notification registration + config
import { db } from "../db/client.js";
import {
  getNotifyApkUrl,
  getNotifyExpoInstallUrl,
  isPushConfigured,
  sendPushToUsernames,
} from "../utils/pushNotify.js";

function parseRoles(req) {
  const raw = String(req.headers["x-user-roles"] || req.headers["x-user-role"] || "")
    .split(",")
    .map((r) => r.trim().toLowerCase())
    .filter(Boolean);
  return Array.from(new Set(raw));
}

function requireRoles(req, reply, allowed) {
  const roles = parseRoles(req);
  const ok = roles.some((r) => allowed.includes(r));
  if (!ok) {
    reply.code(403).send({ error: "forbidden" });
    return false;
  }
  return true;
}

function sessionUsername(req) {
  return String(req.headers["x-user-name"] || "").trim().toLowerCase();
}

export default async function notificationsRoutes(app) {
  app.get("/config", async (req) => {
    const expoInstallUrl = getNotifyExpoInstallUrl();
    return {
      ok: true,
      push_enabled: isPushConfigured(),
      apk_url: getNotifyApkUrl(req),
      expo_install_url: expoInstallUrl,
      expo_qr_url: `https://api.qrserver.com/v1/create-qr-code/?size=320x320&data=${encodeURIComponent(expoInstallUrl)}`,
      register_path: "/api/notifications/register",
    };
  });

  app.post("/register", async (req, reply) => {
    const username = sessionUsername(req);
    if (!username) return reply.code(401).send({ error: "login required" });

    const roles = parseRoles(req);
    const allowed = roles.some((r) =>
      ["artisan", "admin", "supervisor", "operator"].includes(r)
    );
    if (!allowed) return reply.code(403).send({ error: "role not allowed for push" });

    const token = String(req.body?.token || "").trim();
    if (!token || token.length < 20) {
      return reply.code(400).send({ error: "valid FCM device token required" });
    }
    const platform = String(req.body?.platform || "android").trim().toLowerCase() || "android";
    const deviceLabel = String(req.body?.device_label || "").trim() || null;

    db.prepare(`
      INSERT INTO push_device_tokens (username, token, platform, device_label, updated_at, last_seen_at)
      VALUES (?, ?, ?, ?, datetime('now'), datetime('now'))
      ON CONFLICT(token) DO UPDATE SET
        username = excluded.username,
        platform = excluded.platform,
        device_label = COALESCE(excluded.device_label, push_device_tokens.device_label),
        updated_at = datetime('now'),
        last_seen_at = datetime('now')
    `).run(username, token, platform, deviceLabel);

    return { ok: true, username, push_enabled: isPushConfigured() };
  });

  app.post("/unregister", async (req, reply) => {
    const username = sessionUsername(req);
    if (!username) return reply.code(401).send({ error: "login required" });
    const token = String(req.body?.token || "").trim();
    if (!token) return reply.code(400).send({ error: "token required" });
    db.prepare(`
      DELETE FROM push_device_tokens
      WHERE token = ? AND lower(username) = lower(?)
    `).run(token, username);
    return { ok: true };
  });

  app.get("/inbox", async (req, reply) => {
    const username = sessionUsername(req);
    if (!username) return reply.code(401).send({ error: "login required" });
    const limit = Math.min(50, Math.max(1, Number(req.query?.limit) || 20));
    const rows = db
      .prepare(`
        SELECT id, kind, title, body, data_json, sent_at, success
        FROM push_notification_log
        WHERE lower(username) = lower(?)
        ORDER BY id DESC
        LIMIT ?
      `)
      .all(username, limit);
    return {
      ok: true,
      items: rows.map((r) => ({
        ...r,
        success: Boolean(r.success),
        data: r.data_json ? JSON.parse(r.data_json) : null,
      })),
    };
  });

  app.get("/admin", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const devices = db
      .prepare(`
        SELECT username, platform, device_label, last_seen_at, updated_at
        FROM push_device_tokens
        ORDER BY last_seen_at DESC
      `)
      .all();
    const recent = db
      .prepare(`
        SELECT id, username, kind, title, body, sent_at, success, error
        FROM push_notification_log
        ORDER BY id DESC
        LIMIT 25
      `)
      .all();
    const expoInstallUrl = getNotifyExpoInstallUrl();
    return {
      ok: true,
      push_enabled: isPushConfigured(),
      apk_url: getNotifyApkUrl(req),
      expo_install_url: expoInstallUrl,
      expo_qr_url: `https://api.qrserver.com/v1/create-qr-code/?size=320x320&data=${encodeURIComponent(expoInstallUrl)}`,
      devices,
      recent: recent.map((r) => ({
        ...r,
        success: Boolean(r.success),
      })),
    };
  });

  app.post("/test", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const username = String(req.body?.username || sessionUsername(req)).trim();
    if (!username) return reply.code(400).send({ error: "username required" });
    const result = await sendPushToUsernames({
      usernames: [username],
      kind: "test",
      title: "IRONLOG test notification",
      body: "Push notifications are working.",
      data: { kind: "test" },
    });
    return { ok: true, ...result, push_enabled: isPushConfigured() };
  });

  app.post("/send", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const single = String(req.body?.username || "").trim();
    const many = Array.isArray(req.body?.usernames)
      ? req.body.usernames.map((u) => String(u || "").trim()).filter(Boolean)
      : [];
    const usernames = many.length ? many : single ? [single] : [];
    const title = String(req.body?.title || "").trim();
    const body = String(req.body?.body || "").trim();
    const woId = String(req.body?.wo_id || "").trim();
    if (!usernames.length) return reply.code(400).send({ error: "username required" });
    if (!title) return reply.code(400).send({ error: "title required" });
    if (!body) return reply.code(400).send({ error: "body required" });
    const data = { kind: "manual" };
    if (woId) data.wo_id = woId;
    const result = await sendPushToUsernames({
      usernames,
      kind: "manual",
      title,
      body,
      data,
    });
    return { ok: true, ...result, push_enabled: isPushConfigured() };
  });
}
