// IRONLOG/api/utils/workOrderReopen.js — undo closing a work order.
//
// Closing a service work order records the service as done: the machine's
// active plans get last_service_hours = current meter, so the service drops off
// the due list. When a work order was closed by mistake, reopening it must undo
// that too, otherwise the machine can never get the work order back.
//
// From now on the plans' previous values are kept on the work order at close
// (plan_hours_before_close, JSON {plan_id: hours}) and put back exactly. For
// work orders closed before that, plans still carrying the value the close set
// go back one interval, which shows the service as due again.

function hasColumn(db, table, col) {
  return db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
}

export function ensureReopenSchema(db) {
  if (!hasColumn(db, "work_orders", "plan_hours_before_close")) {
    db.prepare(`ALTER TABLE work_orders ADD COLUMN plan_hours_before_close TEXT`).run();
  }
}

/** Saves the machine's active plan hours on the work order before a close changes them. */
export function rememberPlanHours(db, woId, assetId) {
  ensureReopenSchema(db);
  const rows = db.prepare(`SELECT id, last_service_hours FROM maintenance_plans WHERE asset_id = ? AND active = 1`).all(Number(assetId));
  const snap = Object.fromEntries(rows.map((r) => [r.id, r.last_service_hours == null ? null : Number(r.last_service_hours)]));
  db.prepare(`UPDATE work_orders SET plan_hours_before_close = ? WHERE id = ?`).run(JSON.stringify(snap), Number(woId));
}

const REOPENABLE = ["closed", "approved", "completed"];

/**
 * Reopens a finished work order: back to "assigned" (or "open" when nobody was
 * assigned) so it can be assigned and worked again; service plans and a closed
 * breakdown are put back as they were.
 */
export function reopenWorkOrder(db, woId, { rolledHours = null } = {}) {
  ensureReopenSchema(db);
  const wo = db.prepare(`SELECT * FROM work_orders WHERE id = ?`).get(Number(woId));
  if (!wo) throw Object.assign(new Error("work order not found"), { status: 404 });
  const from = String(wo.status || "").toLowerCase().replace(/\s+/g, "_");
  if (!REOPENABLE.includes(from)) {
    throw Object.assign(new Error("Only a completed, approved or closed work order can be reopened"), { status: 409 });
  }
  const to = String(wo.assigned_artisan_name || "").trim() ? "assigned" : "open";
  const plans = [];
  let breakdownReopened = false;

  const tx = db.transaction(() => {
    db.prepare(`
      UPDATE work_orders
      SET status = ?, closed_at = NULL, completed_at = NULL, plan_hours_before_close = NULL
      WHERE id = ?
    `).run(to, wo.id);

    if (String(wo.source || "").toLowerCase() === "service") {
      let saved = null;
      try { saved = wo.plan_hours_before_close ? JSON.parse(wo.plan_hours_before_close) : null; } catch { saved = null; }
      if (saved && typeof saved === "object") {
        for (const [planId, hours] of Object.entries(saved)) {
          const p = db.prepare(`SELECT id, last_service_hours FROM maintenance_plans WHERE id = ?`).get(Number(planId));
          if (!p) continue;
          db.prepare(`UPDATE maintenance_plans SET last_service_hours = ? WHERE id = ?`).run(hours, p.id);
          plans.push({ plan_id: p.id, from: p.last_service_hours, to: hours, exact: true });
        }
      } else if (Number(wo.reference_id) > 0) {
        // Closed before plan hours were kept: plans still on the hours the close
        // set go back one interval, so the service shows as due again.
        const own = db.prepare(`SELECT last_service_hours FROM maintenance_plans WHERE id = ?`).get(Number(wo.reference_id));
        const closeHours = rolledHours != null ? Number(rolledHours) : own ? Number(own.last_service_hours) : null;
        if (closeHours != null && Number.isFinite(closeHours)) {
          for (const p of db.prepare(`SELECT id, last_service_hours, interval_hours FROM maintenance_plans WHERE asset_id = ? AND active = 1`).all(wo.asset_id)) {
            if (Number(p.last_service_hours) !== closeHours) continue;
            const back = Math.max(0, closeHours - Number(p.interval_hours || 0));
            db.prepare(`UPDATE maintenance_plans SET last_service_hours = ? WHERE id = ?`).run(back, p.id);
            plans.push({ plan_id: p.id, from: p.last_service_hours, to: back, exact: false });
          }
        }
      }
    }

    if (String(wo.source || "").toLowerCase() === "breakdown" && Number(wo.reference_id) > 0) {
      const res = db.prepare(`
        UPDATE breakdowns SET status = 'OPEN', end_at = NULL
        WHERE id = ? AND UPPER(TRIM(COALESCE(status, ''))) = 'CLOSED'
      `).run(Number(wo.reference_id));
      breakdownReopened = res.changes > 0;
    }
  });
  tx();
  return { id: wo.id, from_status: from, status: to, plans_restored: plans, breakdown_reopened: breakdownReopened };
}
