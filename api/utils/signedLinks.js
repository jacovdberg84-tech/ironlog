// IRONLOG/api/utils/signedLinks.js — links that open one record without logging in.
//
// Operators do pre-starts from a QR code without an account. Their "Open check
// PDF" link carries a signature for that one check (and an expiry), so the PDF
// opens without a login while other checks stay private: guessing another
// check number gives no valid signature.

import crypto from "node:crypto";

const DAY_MS = 24 * 60 * 60 * 1000;

function linkSecret(db) {
  const fromEnv = String(process.env.IRONLOG_LINK_SECRET || process.env.IRONLOG_AUTH_SECRET || "").trim();
  if (fromEnv) return fromEnv;
  db.exec(`CREATE TABLE IF NOT EXISTS app_secrets (key TEXT PRIMARY KEY, value TEXT NOT NULL, created_at TEXT NOT NULL DEFAULT (datetime('now')))`);
  const row = db.prepare(`SELECT value FROM app_secrets WHERE key = 'link_secret'`).get();
  if (row?.value) return row.value;
  const value = crypto.randomBytes(32).toString("hex");
  db.prepare(`INSERT OR IGNORE INTO app_secrets (key, value) VALUES ('link_secret', ?)`).run(value);
  return db.prepare(`SELECT value FROM app_secrets WHERE key = 'link_secret'`).get().value;
}

function sign(db, purpose, id, exp) {
  return crypto.createHmac("sha256", linkSecret(db)).update(`${purpose}:${Number(id)}:${exp}`).digest("base64url").slice(0, 32);
}

/** Token for one record: "<expiry ms>.<signature>". */
export function signRecordToken(db, purpose, id, { days = 14, now = Date.now() } = {}) {
  const exp = now + days * DAY_MS;
  return `${exp}.${sign(db, purpose, id, exp)}`;
}

export function verifyRecordToken(db, purpose, id, token, { now = Date.now() } = {}) {
  const [expRaw, sig] = String(token || "").split(".");
  const exp = Number(expRaw);
  if (!Number.isFinite(exp) || exp < now || !sig) return false;
  const expected = sign(db, purpose, id, exp);
  if (expected.length !== sig.length) return false;
  return crypto.timingSafeEqual(Buffer.from(expected), Buffer.from(sig));
}

/** Public PDF link for a pre-start check (machine or LDV). */
export function prestartPdfUrl(db, checkId) {
  const id = Number(checkId || 0);
  if (!id) return null;
  return `/api/reports/prestart-check/${id}.pdf?t=${encodeURIComponent(signRecordToken(db, "prestart-pdf", id))}`;
}
