// IRONLOG/api/utils/prestartFaults.js — pre-start checks that find a fault.
//
// Operators answer every check with OK or Fault. A check with faults is still
// saved (so the record is honest) and the faults go to the workshop as one
// repair work order (source 'prestart', reference_id = first check id). Later
// checks with faults on the same machine join that work order while it is
// still open, so a machine never collects a pile of pre-start jobs.

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

const DONE = ["completed", "approved", "closed"];

/**
 * Which pre-start checks feed which work order (one machine, one open pre-start
 * job). LDV and machine checks share vehicle_ldv_checks, so a check id is unique;
 * check_kind is kept for reading.
 */
export function ensurePrestartLinkTable(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS work_order_prestart_checks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      work_order_id INTEGER NOT NULL,
      check_kind TEXT NOT NULL DEFAULT 'machine',
      check_id INTEGER NOT NULL UNIQUE,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_wo_prestart_checks_wo ON work_order_prestart_checks(work_order_id);
  `);
}

function linkCheck(db, workOrderId, checkKind, checkId) {
  db.prepare(`
    INSERT INTO work_order_prestart_checks (work_order_id, check_kind, check_id) VALUES (?, ?, ?)
    ON CONFLICT(check_id) DO UPDATE SET work_order_id = excluded.work_order_id
  `).run(Number(workOrderId), checkKind, Number(checkId));
}

/** How many different pre-start checks a work order holds. */
function checksOnWorkOrder(db, wo) {
  const linked = db.prepare(`SELECT check_id FROM work_order_prestart_checks WHERE work_order_id = ?`).all(wo.id);
  const ids = new Set(linked.map((r) => Number(r.check_id)));
  // Older work orders only know their first check through reference_id.
  ids.add(Number(wo.reference_id));
  return ids.size;
}

function isUnfinished(status) {
  return !DONE.includes(String(status || "").trim().toLowerCase());
}

function faultLines(faults) {
  return faults.map((f) => `- ${f.label}${f.comment ? `: ${f.comment}` : ""}`);
}

/**
 * Opens (or updates) the repair work order for a pre-start with faults.
 *
 * One machine has at most one open pre-start work order: when a new check
 * finds faults and the machine already has an unfinished pre-start job, the
 * faults are added to that job instead of opening another one. Resubmitting
 * the same check updates its own job; a resubmit with no faults closes the
 * job only when it holds just that one check and nobody has started on it.
 * Returns { work_order_id, created, closed, added } or null when nothing applies.
 */
export function syncPrestartFaultWorkOrder(db, { assetId, checkId, checkKind = "machine", siteCode = "main", checkDate, operator, faults }) {
  ensurePrestartLinkTable(db);
  const kind = checkKind === "ldv" ? "ldv" : "machine";
  const withDescription = hasColumn(db, "work_orders", "job_description");
  const linked = db.prepare(`
    SELECT w.id, w.status, w.reference_id FROM work_order_prestart_checks l
    JOIN work_orders w ON w.id = l.work_order_id
    WHERE l.check_id = ?
  `).get(Number(checkId));
  const existing = linked || db.prepare(`
    SELECT id, status, reference_id FROM work_orders
    WHERE source = 'prestart' AND reference_id = ? AND asset_id = ?
    ORDER BY id DESC LIMIT 1
  `).get(Number(checkId), Number(assetId));
  const status = String(existing?.status || "").trim().toLowerCase();
  const unfinished = existing && isUnfinished(status);
  const shared = existing ? checksOnWorkOrder(db, existing) > 1 : false;
  const who = String(operator || "").trim() || "operator";

  if (!faults.length) {
    if (existing && status === "open" && !shared) {
      db.prepare(`
        UPDATE work_orders
        SET status = 'closed', closed_at = datetime('now'),
            completion_notes = COALESCE(NULLIF(TRIM(completion_notes), ''), 'Cleared: pre-start resubmitted with no faults')
        WHERE id = ?
      `).run(existing.id);
      return { work_order_id: Number(existing.id), created: false, closed: true };
    }
    if (existing && unfinished && shared && withDescription) {
      // Other checks still have faults on this job: just note the all-clear.
      appendToDescription(db, existing.id, `Re-checked by ${who} on ${checkDate}: no faults on this pre-start.`);
    }
    return null;
  }

  const description = [
    `Pre-start fault${faults.length === 1 ? "" : "s"} reported by ${who} on ${checkDate}:`,
    ...faultLines(faults),
  ].join("\n");

  if (unfinished) {
    if (withDescription) {
      if (shared) appendToDescription(db, existing.id, `Updated by ${who} on ${checkDate}:\n${faultLines(faults).join("\n")}`);
      else db.prepare(`UPDATE work_orders SET job_description = ? WHERE id = ?`).run(description, existing.id);
    }
    linkCheck(db, existing.id, kind, checkId);
    return { work_order_id: Number(existing.id), created: false, closed: false };
  }

  // A new check: join the machine's open pre-start job if there is one.
  const open = db.prepare(`
    SELECT id, status, job_description FROM work_orders
    WHERE source = 'prestart' AND asset_id = ?
      AND LOWER(TRIM(COALESCE(status, 'open'))) NOT IN ('completed', 'approved', 'closed')
    ORDER BY id DESC LIMIT 1
  `).get(Number(assetId));
  if (open) {
    if (withDescription) {
      const already = String(open.job_description || "");
      const lines = faultLines(faults).map((l) => (already.includes(l) ? `${l} (again)` : l));
      appendToDescription(db, open.id, `Also reported by ${who} on ${checkDate}:\n${lines.join("\n")}`);
    }
    linkCheck(db, open.id, kind, checkId);
    return { work_order_id: Number(open.id), created: false, closed: false, added: true };
  }

  const ins = db.prepare(`
    INSERT INTO work_orders (asset_id, source, reference_id, status, site_code)
    VALUES (?, 'prestart', ?, 'open', ?)
  `).run(Number(assetId), Number(checkId), String(siteCode || "main"));
  const id = Number(ins.lastInsertRowid);
  if (withDescription) db.prepare(`UPDATE work_orders SET job_description = ? WHERE id = ?`).run(description, id);
  linkCheck(db, id, kind, checkId);
  return { work_order_id: id, created: true, closed: false };
}

function appendToDescription(db, workOrderId, block) {
  const row = db.prepare(`SELECT job_description FROM work_orders WHERE id = ?`).get(Number(workOrderId));
  const current = String(row?.job_description || "").trim();
  db.prepare(`UPDATE work_orders SET job_description = ? WHERE id = ?`).run(current ? `${current}\n\n${block}` : block, Number(workOrderId));
}

export function faultMessage(faults, wo) {
  if (!faults.length) return null;
  const list = faults.map((f) => f.label).join(", ");
  return `Pre-start saved with ${faults.length} fault${faults.length === 1 ? "" : "s"} (${list}). `
    + `${wo?.work_order_id ? `The workshop has work order #${wo.work_order_id}${wo.added ? " (added to the open pre-start job)" : ""}. ` : ""}`
    + "Tell your foreman before you operate.";
}

/**
 * Adds an English version after each Portuguese fault comment on the repair
 * work order ("- Brakes: pedal mole (EN: soft pedal)"). translate(text) resolves
 * to English or null. Runs after the operator's submit has been answered.
 */
export async function addEnglishToFaultComments(db, workOrderId, faults, translate) {
  const withComments = faults.filter((f) => f.comment);
  if (!workOrderId || !withComments.length) return false;
  const row = db.prepare(`SELECT job_description FROM work_orders WHERE id = ?`).get(Number(workOrderId));
  if (!row) return false;
  let desc = String(row.job_description || "");
  let changed = false;
  for (const f of withComments) {
    const line = `- ${f.label}: ${f.comment}`;
    if (!desc.includes(line) || desc.includes(`${line} (EN:`)) continue;
    const en = String((await translate(f.comment)) || "").trim();
    if (!en || en.toLowerCase() === f.comment.toLowerCase()) continue;
    desc = desc.replace(line, `${line} (EN: ${en})`);
    changed = true;
  }
  if (changed) db.prepare(`UPDATE work_orders SET job_description = ? WHERE id = ?`).run(desc, Number(workOrderId));
  return changed;
}
