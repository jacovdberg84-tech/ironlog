// IRONLOG/api/utils/prestartFaults.js — pre-start checks that find a fault.
//
// Operators answer every check with OK or Fault. A check with faults is still
// saved (so the record is honest) and the faults go to the workshop as one
// repair work order per check (source 'prestart', reference_id = check id).
// Resubmitting the same check updates that work order instead of adding another.

/** Labels of checks the operator did not answer (neither OK nor Fault). */
export function unansweredChecks(checklist, rawInput) {
  const src = rawInput && typeof rawInput === "object" && !Array.isArray(rawInput) ? rawInput : {};
  return checklist.filter((c) => typeof src[c.key] !== "boolean").map((c) => plainLabel(c.label));
}

/** Template labels read "Engine oil level OK"; a fault reads better without the "OK". */
export function plainLabel(label) {
  return String(label || "").replace(/\s+OK$/i, "").trim() || String(label || "");
}

/** Failed checks with the operator's comment, if any. */
export function prestartFaultList(checklist, faultComments) {
  const comments = faultComments && typeof faultComments === "object" ? faultComments : {};
  return checklist
    .filter((c) => !c.ok)
    .map((c) => ({ key: c.key, label: plainLabel(c.label), comment: String(comments[c.key] || "").trim().slice(0, 300) }));
}

export function faultSummaryText(faults) {
  return faults.map((f) => (f.comment ? `${f.label} (${f.comment})` : f.label)).join("; ");
}

/** Operator notes plus a "Faults:" line, rebuilt on every submit so it never piles up. */
export function notesWithFaults(notes, faults) {
  const base = String(notes || "")
    .split(" | ")
    .map((x) => x.trim())
    .filter((x) => x && !x.startsWith("Faults: ") && x !== "KM flagged for supervisor review")
    .join(" | ");
  if (!faults.length) return base || null;
  return [base, `Faults: ${faultSummaryText(faults)}`].filter(Boolean).join(" | ");
}

function hasColumn(db, table, col) {
  return db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
}

/**
 * Opens (or updates) the repair work order for a pre-start with faults. When a
 * resubmitted check has no faults left, an untouched open work order is closed.
 * Returns { work_order_id, created, closed } or null when nothing applies.
 */
export function syncPrestartFaultWorkOrder(db, { assetId, checkId, siteCode = "main", checkDate, operator, faults }) {
  const existing = db.prepare(`
    SELECT id, status FROM work_orders
    WHERE source = 'prestart' AND reference_id = ?
    ORDER BY id DESC LIMIT 1
  `).get(Number(checkId));
  const status = String(existing?.status || "").trim().toLowerCase();
  const unfinished = existing && !["completed", "approved", "closed"].includes(status);

  if (!faults.length) {
    if (existing && status === "open") {
      db.prepare(`
        UPDATE work_orders
        SET status = 'closed', closed_at = datetime('now'),
            completion_notes = COALESCE(NULLIF(TRIM(completion_notes), ''), 'Cleared: pre-start resubmitted with no faults')
        WHERE id = ?
      `).run(existing.id);
      return { work_order_id: Number(existing.id), created: false, closed: true };
    }
    return null;
  }

  const who = String(operator || "").trim() || "operator";
  const description = [
    `Pre-start fault${faults.length === 1 ? "" : "s"} reported by ${who} on ${checkDate}:`,
    ...faults.map((f) => `- ${f.label}${f.comment ? `: ${f.comment}` : ""}`),
  ].join("\n");
  const withDescription = hasColumn(db, "work_orders", "job_description");

  if (unfinished) {
    if (withDescription) db.prepare(`UPDATE work_orders SET job_description = ? WHERE id = ?`).run(description, existing.id);
    return { work_order_id: Number(existing.id), created: false, closed: false };
  }
  const ins = db.prepare(`
    INSERT INTO work_orders (asset_id, source, reference_id, status, site_code)
    VALUES (?, 'prestart', ?, 'open', ?)
  `).run(Number(assetId), Number(checkId), String(siteCode || "main"));
  const id = Number(ins.lastInsertRowid);
  if (withDescription) db.prepare(`UPDATE work_orders SET job_description = ? WHERE id = ?`).run(description, id);
  return { work_order_id: id, created: true, closed: false };
}

export function faultMessage(faults, wo) {
  if (!faults.length) return null;
  const list = faults.map((f) => f.label).join(", ");
  return `Pre-start saved with ${faults.length} fault${faults.length === 1 ? "" : "s"} (${list}). `
    + `${wo?.work_order_id ? `The workshop has work order #${wo.work_order_id}. ` : ""}`
    + "Tell your foreman before you operate.";
}
