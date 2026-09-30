// IRONLOG/api/utils/stockReversal.js — undo a stock receipt entered by mistake
// (for example a delivery captured twice).
//
// Stores ask for a reversal; like any stock adjustment it waits in Approvals.
// On approval an equal and opposite movement is posted with the receipt's
// location, bin and cost, and its reference points back at the receipt
// ("reversal_of:<id>"), so the receipt can never be reversed twice and the
// ledger shows both lines.

export const REVERSAL_ACTION = "reverse_movement";
const REF_PREFIX = "reversal_of:";
const DUPLICATE_WINDOW_MIN = 30;

function hasColumn(db, table, col) {
  return db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
}

function hasTable(db, name) {
  return Boolean(db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?`).get(name));
}

export function reversalReference(movement) {
  const ref = String(movement.reference || "").trim();
  return `${REF_PREFIX}${movement.id}${ref ? ` ${ref}` : ""}`.slice(0, 200);
}

/** Reversal state per receipt id: { reversed_by, pending_request_id }. */
export function reversalStatus(db, ids) {
  const out = new Map(ids.map((id) => [Number(id), { reversed_by: null, pending_request_id: null }]));
  if (!ids.length) return out;
  for (const r of db.prepare(`SELECT id, reference FROM stock_movements WHERE reference LIKE '${REF_PREFIX}%'`).all()) {
    const m = String(r.reference).match(/^reversal_of:(\d+)/);
    if (m && out.has(Number(m[1]))) out.get(Number(m[1])).reversed_by = r.id;
  }
  if (hasTable(db, "approval_requests")) {
    for (const r of db.prepare(`
      SELECT id, entity_id FROM approval_requests
      WHERE module = 'stock' AND action = ? AND LOWER(status) = 'pending'
    `).all(REVERSAL_ACTION)) {
      const e = out.get(Number(r.entity_id));
      if (e) e.pending_request_id = r.id;
    }
  }
  return out;
}

/**
 * Recent receipts (stock in), newest first, with reversal state and a hint when
 * a receipt looks like a repeat of an earlier one: same part and quantity, and
 * the same reference or captured within 30 minutes of it.
 */
export function recentReceipts(db, { partCode = "", days = 30, limit = 200, now = new Date() } = {}) {
  const since = new Date(now.getTime() - days * 86400000).toISOString().replace("T", " ").slice(0, 19);
  const locJoin = hasColumn(db, "stock_movements", "location_id") && hasTable(db, "stock_locations")
    ? "LEFT JOIN stock_locations l ON l.id = sm.location_id" : "";
  const binJoin = hasColumn(db, "stock_movements", "bin_id") && hasTable(db, "stock_bins")
    ? "LEFT JOIN stock_bins b ON b.id = sm.bin_id" : "";
  const rows = db.prepare(`
    SELECT sm.id, sm.part_id, sm.quantity, sm.reference, sm.created_at, p.part_code, p.part_name,
      ${locJoin ? "l.location_code" : "NULL"} AS location_code, ${binJoin ? "b.bin_code" : "NULL"} AS bin_code
    FROM stock_movements sm
    JOIN parts p ON p.id = sm.part_id
    ${locJoin} ${binJoin}
    WHERE LOWER(sm.movement_type) = 'in' AND sm.quantity > 0 AND datetime(sm.created_at) >= datetime(?)
      ${partCode ? "AND (UPPER(p.part_code) LIKE UPPER(?) OR UPPER(p.part_name) LIKE UPPER(?))" : ""}
    ORDER BY sm.id DESC
    LIMIT ?
  `).all(...(partCode ? [since, `%${partCode}%`, `%${partCode}%`, limit] : [since, limit]));
  const status = reversalStatus(db, rows.map((r) => r.id));
  const ms = (t) => Date.parse(`${String(t).replace(" ", "T")}Z`);
  return rows.map((r) => {
    const earlier = rows.find((o) =>
      o.id < r.id && o.part_id === r.part_id && Number(o.quantity) === Number(r.quantity) && !status.get(o.id).reversed_by &&
      ((r.reference && o.reference && String(r.reference).trim().toLowerCase() === String(o.reference).trim().toLowerCase() && !/^manual_entry$/i.test(String(r.reference).trim())) ||
        Math.abs(ms(r.created_at) - ms(o.created_at)) <= DUPLICATE_WINDOW_MIN * 60000));
    const st = status.get(r.id);
    return {
      id: r.id,
      part_code: r.part_code,
      part_name: r.part_name,
      quantity: Number(r.quantity),
      reference: r.reference,
      created_at: r.created_at,
      location_code: r.location_code,
      bin_code: r.bin_code,
      reversed_by: st.reversed_by,
      pending_request_id: st.pending_request_id,
      possible_duplicate_of: st.reversed_by ? null : earlier ? earlier.id : null,
    };
  });
}

function onHand(db, partId) {
  return Number(db.prepare(`SELECT IFNULL(SUM(quantity), 0) AS q FROM stock_movements WHERE part_id = ?`).get(partId).q || 0);
}

/** Checks a receipt can be reversed; returns the movement or throws with a plain reason. */
export function reversibleReceipt(db, movementId) {
  const m = db.prepare(`SELECT sm.*, p.part_code FROM stock_movements sm JOIN parts p ON p.id = sm.part_id WHERE sm.id = ?`).get(Number(movementId));
  if (!m) throw Object.assign(new Error("Stock movement not found"), { status: 404 });
  if (String(m.movement_type).toLowerCase() !== "in" || !(Number(m.quantity) > 0)) {
    throw Object.assign(new Error("Only a stock receipt (stock in) can be reversed"), { status: 400 });
  }
  const st = reversalStatus(db, [m.id]).get(m.id);
  if (st.reversed_by) throw Object.assign(new Error(`This receipt was already reversed (movement #${st.reversed_by})`), { status: 409 });
  return { movement: m, pending_request_id: st.pending_request_id };
}

/** Posts the reversal (called when the approval is approved). */
export function applyReversal(db, movementId, { approvalId = null } = {}) {
  const { movement: m } = reversibleReceipt(db, movementId);
  const qty = Number(m.quantity);
  const before = onHand(db, m.part_id);
  if (before < qty) {
    throw Object.assign(new Error(`Only ${before} of ${m.part_code} left in stock; ${qty} cannot be reversed (some was already issued). Use a stock adjustment for the rest.`), { status: 409 });
  }
  const cols = ["part_id", "quantity", "movement_type", "reference"];
  const vals = [m.part_id, -qty, "adjust", reversalReference(m)];
  for (const c of ["location_id", "bin_id", "cost_center_code", "unit_cost_usd", "cost_currency"]) {
    if (hasColumn(db, "stock_movements", c) && m[c] != null) { cols.push(c); vals.push(m[c]); }
  }
  const ins = db.prepare(`INSERT INTO stock_movements (${cols.join(", ")}) VALUES (${cols.map(() => "?").join(", ")})`).run(...vals);
  return {
    part_code: m.part_code,
    reversed_movement_id: m.id,
    reversal_movement_id: Number(ins.lastInsertRowid),
    quantity: -qty,
    on_hand_before: before,
    on_hand_after: onHand(db, m.part_id),
    approval_id: approvalId,
  };
}
