// IRONLOG/api/utils/workOrderSync.js — keeps breakdowns and their work orders in step.
//
// Every breakdown gets a work order (source 'breakdown', reference_id = breakdown id).
// A breakdown is finished when the machine is back in service; its work order is
// finished once it is completed, signed off (approved) or closed. Either side
// finishing used to leave the other open, so machines stayed "broken down" and
// work orders never left the queue.

const FINISHED = ["completed", "approved", "closed"];
const FINISHED_SQL = FINISHED.map((s) => `'${s}'`).join(", ");
const woStatusSql = (alias) => `REPLACE(TRIM(LOWER(COALESCE(${alias}.status, ''))), ' ', '_')`;

function hasColumn(db, table, col) {
  try {
    return db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
  } catch {
    return false;
  }
}

/**
 * Closes a breakdown once none of its breakdown work orders is still unfinished.
 * Returns true when the breakdown was closed by this call.
 */
export function closeBreakdownIfWorkFinished(db, breakdownId) {
  const id = Number(breakdownId || 0);
  if (!id) return false;
  const endExpr = hasColumn(db, "work_orders", "completed_at")
    ? "COALESCE(w.completed_at, w.closed_at)"
    : "w.closed_at";
  const info = db.prepare(`
    UPDATE breakdowns
    SET status = 'CLOSED',
        end_at = COALESCE(NULLIF(TRIM(end_at), ''), (
          SELECT MAX(${endExpr}) FROM work_orders w
          WHERE w.source = 'breakdown' AND w.reference_id = breakdowns.id
        ), datetime('now'))
    WHERE id = ?
      AND UPPER(TRIM(COALESCE(status, 'OPEN'))) <> 'CLOSED'
      AND EXISTS (SELECT 1 FROM work_orders w WHERE w.source = 'breakdown' AND w.reference_id = breakdowns.id)
      AND NOT EXISTS (
        SELECT 1 FROM work_orders w
        WHERE w.source = 'breakdown' AND w.reference_id = breakdowns.id
          AND ${woStatusSql("w")} NOT IN (${FINISHED_SQL})
      )
  `).run(id);
  return info.changes > 0;
}

/**
 * When a breakdown is closed (machine back in service), close its breakdown work
 * orders that were still open, assigned or in progress. Returns their ids.
 */
export function closeWorkOrdersForClosedBreakdown(db, breakdownId, note = "Closed when the breakdown was closed") {
  const id = Number(breakdownId || 0);
  if (!id) return [];
  const rows = db.prepare(`
    SELECT w.id FROM work_orders w
    WHERE w.source = 'breakdown' AND w.reference_id = ?
      AND ${woStatusSql("w")} NOT IN (${FINISHED_SQL})
  `).all(id);
  if (!rows.length) return [];
  const notes = hasColumn(db, "work_orders", "completion_notes");
  const close = db.prepare(`
    UPDATE work_orders
    SET status = 'closed',
        closed_at = COALESCE(closed_at, (SELECT COALESCE(NULLIF(TRIM(end_at), ''), datetime('now')) FROM breakdowns WHERE id = ?))
        ${notes ? ", completion_notes = COALESCE(NULLIF(TRIM(completion_notes), ''), ?)" : ""}
    WHERE id = ?
  `);
  for (const r of rows) {
    if (notes) close.run(id, note, r.id);
    else close.run(id, r.id);
  }
  return rows.map((r) => r.id);
}

/**
 * One-off repair for records that drifted apart before the two were kept in
 * step. Idempotent: only touches pairs that disagree.
 */
export function repairBreakdownWorkOrderLinks(db) {
  const hasTables = ["breakdowns", "work_orders"].every((t) =>
    db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?`).get(t),
  );
  if (!hasTables) return { breakdowns_closed: [], work_orders_closed: [] };
  const breakdownsClosed = [];
  const workOrdersClosed = [];
  db.transaction(() => {
    const openBreakdowns = db.prepare(`
      SELECT b.id FROM breakdowns b
      WHERE UPPER(TRIM(COALESCE(b.status, 'OPEN'))) <> 'CLOSED'
    `).all();
    for (const b of openBreakdowns) if (closeBreakdownIfWorkFinished(db, b.id)) breakdownsClosed.push(b.id);
    const closedWithOpenWork = db.prepare(`
      SELECT DISTINCT b.id FROM breakdowns b
      JOIN work_orders w ON w.source = 'breakdown' AND w.reference_id = b.id
      WHERE UPPER(TRIM(COALESCE(b.status, ''))) = 'CLOSED'
        AND ${woStatusSql("w")} NOT IN (${FINISHED_SQL})
    `).all();
    for (const b of closedWithOpenWork) workOrdersClosed.push(...closeWorkOrdersForClosedBreakdown(db, b.id, "Closed by data clean-up: the breakdown was already closed"));
  })();
  return { breakdowns_closed: breakdownsClosed, work_orders_closed: workOrdersClosed };
}
