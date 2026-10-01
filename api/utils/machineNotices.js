// IRONLOG/api/utils/machineNotices.js — workshop notices shown to operators.
//
// When a pre-start reports faults, the next operator who scans the machine's QR
// sees a notice at the top of the pre-start ("Report this machine to the
// workshop…") and ticks "I have read this"; the tick is saved with their check.
// A fault work order puts up a standard notice automatically; the admin can
// change its wording or clear it. A notice linked to a work order ends by
// itself when that work order is completed. English + Portuguese.

import { toPortuguese } from "./prestartPortuguese.js";

const FINISHED = ["completed", "approved", "closed", "cancelled"];

export const NOTICE_PRESETS = [
  {
    key: "report",
    en: "Report this machine to the workshop for repairs before you work with it.",
    pt: "Leve esta máquina à oficina para reparação antes de trabalhar com ela.",
  },
  {
    key: "do_not_operate",
    en: "Do not operate this machine. Wait for the workshop to clear it.",
    pt: "Não opere esta máquina. Aguarde que a oficina a liberte.",
  },
  {
    key: "end_of_shift",
    en: "Repair booked: bring the machine to the workshop at the end of your shift.",
    pt: "Reparação marcada: traga a máquina à oficina no fim do seu turno.",
  },
  {
    key: "workshop_coming",
    en: "The workshop will come to the machine. Keep it parked and report to your foreman.",
    pt: "A oficina vem à máquina. Mantenha-a estacionada e informe o seu encarregado.",
  },
];

export function ensureNoticeSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS machine_notices (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      work_order_id INTEGER,
      source TEXT NOT NULL DEFAULT 'auto',
      message_en TEXT NOT NULL,
      message_pt TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      cleared_by TEXT,
      cleared_at TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_machine_notices_asset ON machine_notices(asset_id, active);
    CREATE TABLE IF NOT EXISTS machine_notice_acks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      notice_id INTEGER NOT NULL,
      asset_id INTEGER NOT NULL,
      check_kind TEXT NOT NULL,
      check_id INTEGER,
      operator TEXT,
      acknowledged_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_machine_notice_acks_notice ON machine_notice_acks(notice_id);
  `);
}

function finishedWorkOrder(db, id) {
  if (!id) return false;
  const w = db.prepare(`SELECT status FROM work_orders WHERE id = ?`).get(Number(id));
  return !w || FINISHED.includes(String(w.status || "").trim().toLowerCase().replace(/\s+/g, "_"));
}

/** The machine's current notice (ending it first when its work order is done). */
export function activeNotice(db, assetId) {
  ensureNoticeSchema(db);
  const n = db.prepare(`SELECT * FROM machine_notices WHERE asset_id = ? AND active = 1 ORDER BY id DESC LIMIT 1`).get(Number(assetId));
  if (!n) return null;
  if (n.work_order_id && finishedWorkOrder(db, n.work_order_id)) {
    db.prepare(`UPDATE machine_notices SET active = 0, cleared_by = 'work order completed', cleared_at = datetime('now') WHERE id = ?`).run(n.id);
    return null;
  }
  const acks = db.prepare(`SELECT COUNT(*) AS n, MAX(acknowledged_at) AS last FROM machine_notice_acks WHERE notice_id = ?`).get(n.id);
  return { ...n, ack_count: Number(acks.n || 0), last_ack_at: acks.last || null };
}

/** For the operator screens: just what they need to read. */
export function noticeForOperator(db, assetId) {
  const n = activeNotice(db, assetId);
  return n ? { id: n.id, message_en: n.message_en, message_pt: n.message_pt || null, work_order_id: n.work_order_id || null, since: n.created_at } : null;
}

export function autoNoticeText(assetCode, faults = [], { origin = "prestart", notFit = false } = {}) {
  const labels = faults.map((f) => String(f.label || "").trim()).filter(Boolean);
  const listEn = labels.length ? ` (${labels.join(", ")})` : "";
  const labelsPt = faults.filter((f) => String(f.label || "").trim()).map((f) => f.label_pt || toPortuguese(String(f.label).trim()) || String(f.label).trim());
  const listPt = labelsPt.length ? ` (${labelsPt.join(", ")})` : "";
  const inspection = origin === "inspection";
  const whereEn = inspection ? "found on" : "reported on";
  const atEn = inspection ? "at a workshop inspection" : "at a pre-start";
  const wherePt = inspection ? "encontradas" : "reportadas";
  const atPt = inspection ? "numa inspecção da oficina" : "numa inspecção pré-arranque";
  const doEn = notFit
    ? "Do not operate this machine. Wait for the workshop to clear it."
    : "Report the machine to the workshop for repairs before you work with it.";
  const doPt = notFit
    ? "Não opere esta máquina. Aguarde que a oficina a liberte."
    : "Leve a máquina à oficina para reparação antes de trabalhar com ela.";
  return {
    en: `Faults were ${whereEn} ${assetCode} ${atEn}${listEn}. ${doEn}`,
    pt: `Foram ${wherePt} avarias na ${assetCode} ${atPt}${listPt}. ${doPt}`,
  };
}

/**
 * Puts up the standard notice for a fault work order. A notice the admin wrote
 * is left alone; an automatic one for the same work order is refreshed.
 */
export function autoNoticeForFaults(db, { assetId, assetCode, workOrderId, faults, origin = "prestart", notFit = false }) {
  ensureNoticeSchema(db);
  const cur = activeNotice(db, assetId);
  const text = autoNoticeText(assetCode, faults, { origin, notFit });
  if (cur && cur.source !== "auto") return cur.id;
  if (cur && cur.source === "auto") {
    db.prepare(`UPDATE machine_notices SET message_en = ?, message_pt = ?, work_order_id = COALESCE(?, work_order_id), updated_at = datetime('now') WHERE id = ?`)
      .run(text.en, text.pt, workOrderId || null, cur.id);
    return cur.id;
  }
  return Number(db.prepare(`
    INSERT INTO machine_notices (asset_id, work_order_id, source, message_en, message_pt, created_by)
    VALUES (?, ?, 'auto', ?, ?, 'pre-start fault')
  `).run(Number(assetId), workOrderId || null, text.en, text.pt).lastInsertRowid);
}

/** Admin writes or changes the machine's notice (replaces the current one). */
export function setNotice(db, { assetId, workOrderId = null, messageEn, messagePt = null, user = null }) {
  ensureNoticeSchema(db);
  const en = String(messageEn || "").trim().slice(0, 500);
  if (!en) throw Object.assign(new Error("Write the message for the operators"), { status: 400 });
  const pt = String(messagePt || "").trim().slice(0, 500) || null;
  const tx = db.transaction(() => {
    db.prepare(`UPDATE machine_notices SET active = 0, cleared_by = ?, cleared_at = datetime('now') WHERE asset_id = ? AND active = 1`).run(user ? `replaced by ${user}` : "replaced", Number(assetId));
    return Number(db.prepare(`
      INSERT INTO machine_notices (asset_id, work_order_id, source, message_en, message_pt, created_by)
      VALUES (?, ?, 'admin', ?, ?, ?)
    `).run(Number(assetId), workOrderId || null, en, pt, user).lastInsertRowid);
  });
  return tx();
}

export function clearNotice(db, assetId, user = null) {
  ensureNoticeSchema(db);
  return db.prepare(`UPDATE machine_notices SET active = 0, cleared_by = ?, cleared_at = datetime('now') WHERE asset_id = ? AND active = 1`)
    .run(user || "admin", Number(assetId)).changes;
}

/** Saves the operator's "I have read this" with their pre-start. */
export function recordAck(db, { noticeId, assetId, checkKind, checkId, operator }) {
  ensureNoticeSchema(db);
  const n = db.prepare(`SELECT id, asset_id FROM machine_notices WHERE id = ?`).get(Number(noticeId));
  if (!n || Number(n.asset_id) !== Number(assetId)) return false;
  const dup = checkId
    ? db.prepare(`SELECT 1 FROM machine_notice_acks WHERE notice_id = ? AND check_kind = ? AND check_id = ?`).get(n.id, checkKind, Number(checkId))
    : null;
  if (!dup) {
    db.prepare(`INSERT INTO machine_notice_acks (notice_id, asset_id, check_kind, check_id, operator) VALUES (?, ?, ?, ?, ?)`)
      .run(n.id, Number(assetId), checkKind, checkId ? Number(checkId) : null, String(operator || "").trim() || null);
  }
  return true;
}

export function noticeAcks(db, noticeId) {
  ensureNoticeSchema(db);
  return db.prepare(`SELECT operator, check_kind, check_id, acknowledged_at FROM machine_notice_acks WHERE notice_id = ? ORDER BY id DESC LIMIT 50`).all(Number(noticeId));
}
