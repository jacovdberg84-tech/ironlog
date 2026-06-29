// IRONLOG outgoing mail — nodemailer + encrypted SMTP settings in SQLite.

import crypto from "node:crypto";
import nodemailer from "nodemailer";
import { db } from "../db/client.js";

function env(name, fallback = "") {
  const v = process.env[name];
  if (v === undefined || v === null || String(v).trim() === "") return fallback;
  return String(v).trim();
}

export function smtpSecret() {
  const raw = env("IRONLOG_SMTP_SECRET", env("IRONLOG_AUTH_SECRET"));
  return raw || "IRONLOG_SMTP_DEFAULT_SECRET_CHANGE_ME";
}

export function encryptSecret(plain) {
  const iv = crypto.randomBytes(12);
  const key = crypto.createHash("sha256").update(smtpSecret()).digest();
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  const encrypted = Buffer.concat([cipher.update(String(plain || ""), "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return `${iv.toString("base64")}.${tag.toString("base64")}.${encrypted.toString("base64")}`;
}

export function decryptSecret(cipherText) {
  const raw = String(cipherText || "").trim();
  if (!raw) return "";
  const [ivB64, tagB64, encB64] = raw.split(".");
  if (!ivB64 || !tagB64 || !encB64) return null;
  try {
    const iv = Buffer.from(ivB64, "base64");
    const tag = Buffer.from(tagB64, "base64");
    const enc = Buffer.from(encB64, "base64");
    const key = crypto.createHash("sha256").update(smtpSecret()).digest();
    const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
    decipher.setAuthTag(tag);
    return Buffer.concat([decipher.update(enc), decipher.final()]).toString("utf8");
  } catch {
    return null;
  }
}

export function ensureSmtpTables() {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS smtp_settings (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      host TEXT,
      port INTEGER,
      secure INTEGER NOT NULL DEFAULT 0,
      username TEXT,
      password_enc TEXT,
      from_email TEXT,
      from_name TEXT,
      updated_by TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`INSERT INTO smtp_settings (id) SELECT 1 WHERE NOT EXISTS (SELECT 1 FROM smtp_settings WHERE id = 1)`).run();
}

export function getSmtpSettingsRow() {
  ensureSmtpTables();
  return db.prepare(`
    SELECT id, host, port, secure, username, password_enc, from_email, from_name, updated_by, updated_at
    FROM smtp_settings
    WHERE id = 1
  `).get() || null;
}

export function smtpPublicPayload(row) {
  const r = row || {};
  return {
    host: String(r.host || ""),
    port: Number(r.port || 587),
    secure: Number(r.secure || 0) === 1 ? 1 : 0,
    username: String(r.username || ""),
    from_email: String(r.from_email || ""),
    from_name: String(r.from_name || ""),
    has_password: Boolean(String(r.password_enc || "").trim()),
    updated_by: String(r.updated_by || ""),
    updated_at: r.updated_at || null,
  };
}

export function formatSmtpError(err) {
  const code = String(err?.code || "").trim();
  const msg = String(err?.response || err?.message || err || "SMTP error").trim();
  const combined = `${code} ${msg}`.trim();
  if (/ECONNREFUSED|ETIMEDOUT|ENOTFOUND|EHOSTUNREACH|ESOCKET/i.test(combined)) {
    return `${msg} — Check host and port. Use port 587 with Secure = No (STARTTLS), or port 465 with Secure = Yes.`;
  }
  if (/535|534|EAUTH|authentication|invalid login|auth failed/i.test(combined)) {
    return `${msg} — Authentication failed. Use the full email as username. For Microsoft 365, enable SMTP AUTH on the mailbox (or use an app password).`;
  }
  if (/certificate|self[- ]signed|UNABLE_TO_VERIFY/i.test(combined)) {
    return `${msg} — TLS certificate problem. Confirm host/port with your IT team.`;
  }
  return msg;
}

export function buildSmtpTransport() {
  const row = getSmtpSettingsRow();
  if (!row) return { error: "SMTP is not configured. Save settings in Admin first." };
  const host = String(row.host || "").trim();
  const username = String(row.username || "").trim();
  const fromEmail = String(row.from_email || "").trim();
  if (!host) return { error: "SMTP host is missing. Save settings again." };
  if (!username) return { error: "SMTP username is missing. Save settings again." };
  if (!fromEmail) return { error: "From email is missing. Save settings again." };
  const enc = String(row.password_enc || "").trim();
  if (!enc) {
    return { error: "SMTP password is not set. Enter the password and click Save SMTP (required on first setup)." };
  }
  const password = decryptSecret(enc);
  if (password === null) {
    return {
      error: "Stored SMTP password could not be decrypted (server secret may have changed). Re-enter the password and click Save SMTP.",
    };
  }
  if (!password) return { error: "SMTP password is empty. Re-enter the password and click Save SMTP." };
  const port = Math.max(1, Number(row.port || 587));
  const secure = Number(row.secure || 0) === 1;
  const transporter = nodemailer.createTransport({
    host,
    port,
    secure,
    requireTLS: !secure && port === 587,
    auth: { user: username, pass: password },
    tls: { minVersion: "TLSv1.2" },
    connectionTimeout: 20000,
    greetingTimeout: 20000,
    socketTimeout: 45000,
  });
  return {
    transporter,
    from: row.from_name ? `"${String(row.from_name).replace(/"/g, "")}" <${fromEmail}>` : fromEmail,
  };
}

export function parseEmailRecipients(raw) {
  return Array.from(new Set(
    String(raw || "")
      .split(",")
      .map((x) => String(x || "").trim())
      .filter(Boolean)
  )).slice(0, 50);
}

export function saveSmtpSettings({
  host,
  port = 587,
  secure = 0,
  username,
  password = "",
  from_email,
  from_name = "",
  updated_by = "system",
}) {
  ensureSmtpTables();
  const existing = getSmtpSettingsRow() || {};
  const hasStoredPassword = Boolean(String(existing.password_enc || "").trim());
  if (!password && !hasStoredPassword) {
    return { ok: false, error: "SMTP password is required on first setup." };
  }
  const passwordEnc = password ? encryptSecret(password) : String(existing.password_enc || "");
  db.prepare(`
    UPDATE smtp_settings
    SET host = ?, port = ?, secure = ?, username = ?, password_enc = ?, from_email = ?, from_name = ?, updated_by = ?, updated_at = datetime('now')
    WHERE id = 1
  `).run(
    String(host || "").trim(),
    Math.max(1, Number(port || 587)),
    Number(secure) === 1 ? 1 : 0,
    String(username || "").trim(),
    passwordEnc,
    String(from_email || "").trim(),
    String(from_name || "").trim() || null,
    String(updated_by || "system").trim() || "system",
  );
  return { ok: true, settings: smtpPublicPayload(getSmtpSettingsRow()) };
}

export function bootstrapSmtpFromEnv({ log = console, force = false } = {}) {
  ensureSmtpTables();
  const host = env("SMTP_HOST");
  if (!host) return { bootstrapped: false, reason: "SMTP_HOST not set" };

  const existing = getSmtpSettingsRow() || {};
  const hasConfig = Boolean(String(existing.host || "").trim() && String(existing.password_enc || "").trim());
  const shouldForce = force || env("SMTP_BOOTSTRAP_FORCE") === "1";
  if (hasConfig && !shouldForce) {
    return { bootstrapped: false, reason: "SMTP already configured in database" };
  }

  const username = env("SMTP_USERNAME");
  const password = env("SMTP_PASSWORD");
  const fromEmail = env("SMTP_FROM_EMAIL", username);
  if (!username || !password || !fromEmail) {
    return { bootstrapped: false, reason: "SMTP_USERNAME, SMTP_PASSWORD, and SMTP_FROM_EMAIL required for bootstrap" };
  }

  const port = Math.max(1, Number(env("SMTP_PORT", "587")) || 587);
  const secure = env("SMTP_SECURE", "0") === "1" ? 1 : 0;
  const fromName = env("SMTP_FROM_NAME", "IRONLOG Reports");

  const out = saveSmtpSettings({
    host,
    port,
    secure,
    username,
    password,
    from_email: fromEmail,
    from_name: fromName,
    updated_by: "env-bootstrap",
  });
  if (!out.ok) return { bootstrapped: false, reason: out.error };

  log.info?.(`[mail] SMTP bootstrapped from environment (${host}:${port})`);
  return { bootstrapped: true, settings: out.settings };
}

export async function sendIronlogMail({ to, subject, text, html, attachments }) {
  const recipients = Array.isArray(to) ? to : parseEmailRecipients(to);
  if (!recipients.length) return { ok: false, error: "No recipients" };

  const smtp = buildSmtpTransport();
  if (smtp.error) return { ok: false, error: smtp.error };

  await smtp.transporter.verify();
  const info = await smtp.transporter.sendMail({
    from: smtp.from,
    to: recipients.join(", "),
    subject: String(subject || "IRONLOG notification"),
    text: text || undefined,
    html: html || undefined,
    attachments: attachments || undefined,
  });
  return { ok: true, messageId: info?.messageId || null, recipients };
}
