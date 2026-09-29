// IRONLOG/api/utils/myWork.js — role-aware "My Work" home screen data.
// Overdue services come from /api/maintenance/due on the client so counts
// match the Maintenance page. The low stock query is shared with /api/alerts.

import { summarizeWaiting, workshopWaitingOnParts } from "./partsWaiting.js";

export function queryLowStock(db, limit = 100) {
  return db.prepare(`
    SELECT
      p.part_code,
      p.part_name,
      p.critical,
      p.min_stock,
      IFNULL(SUM(sm.quantity),0) AS on_hand
    FROM parts p
    LEFT JOIN stock_movements sm ON sm.part_id = p.id
    GROUP BY p.id
    HAVING on_hand < p.min_stock
    ORDER BY p.critical DESC, on_hand ASC
    LIMIT ?
  `).all(limit).map((r) => ({ ...r, critical: Boolean(r.critical), on_hand: Number(r.on_hand) }));
}

const MAINTENANCE_ROLES = ["admin", "supervisor", "workshop_admin", "plant_manager", "site_manager"];
const STORES_ROLES = ["admin", "storeman", "stores", "workshop_admin"];
const PREVIEW = 5;

function section(items) {
  return { count: items.length, items: items.slice(0, PREVIEW) };
}

function myTasks(db, { user, site, today }) {
  const rows = db.prepare(`
    SELECT id, title, status, priority, project, due_date
    FROM tasks
    WHERE LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
      AND LOWER(TRIM(COALESCE(assigned_to, ''))) = ?
      AND COALESCE(status, 'open') != 'done'
    ORDER BY
      CASE WHEN due_date IS NULL OR TRIM(due_date) = '' THEN 1 ELSE 0 END,
      due_date ASC,
      CASE priority WHEN 'high' THEN 1 WHEN 'medium' THEN 2 WHEN 'low' THEN 3 ELSE 4 END,
      id DESC
  `).all(site, user);
  const overdue = rows.filter((t) => t.due_date && String(t.due_date) < today).length;
  const dueToday = rows.filter((t) => String(t.due_date || "") === today).length;
  return { ...section(rows), overdue, due_today: dueToday };
}

/** One row per machine that is down (a machine can have several open breakdowns). */
function openBreakdowns(db, { site }) {
  const rows = db.prepare(`
    SELECT b.id, b.asset_id, a.asset_code, a.asset_name, b.breakdown_date, b.description, b.critical, b.parts_status
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    WHERE UPPER(TRIM(COALESCE(b.status, 'OPEN'))) <> 'CLOSED'
      AND LOWER(TRIM(COALESCE(b.site_code, 'main'))) = ?
    ORDER BY b.critical DESC, b.breakdown_date ASC, b.id ASC
  `).all(site);
  const byMachine = new Map();
  for (const r of rows) {
    const m = byMachine.get(r.asset_id);
    if (m) {
      m.open_breakdowns += 1;
      m.critical = m.critical || Boolean(r.critical);
    } else {
      byMachine.set(r.asset_id, { ...r, critical: Boolean(r.critical), open_breakdowns: 1 });
    }
  }
  const machines = [...byMachine.values()].sort((x, y) => Number(y.critical) - Number(x.critical) || String(x.breakdown_date).localeCompare(String(y.breakdown_date)));
  return { ...section(machines), open_breakdown_records: rows.length };
}

function openWorkOrders(db, { site }) {
  const rows = db.prepare(`
    SELECT w.id, a.asset_code, w.source, w.status, w.opened_at, w.assigned_artisan_name, w.priority
    FROM work_orders w
    JOIN assets a ON a.id = w.asset_id
    WHERE REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
      AND (w.completed_at IS NULL OR TRIM(COALESCE(w.completed_at, '')) = '')
      AND (w.closed_at IS NULL OR TRIM(COALESCE(w.closed_at, '')) = '')
      AND LOWER(TRIM(COALESCE(w.site_code, 'main'))) = ?
    ORDER BY w.opened_at ASC, w.id ASC
  `).all(site);
  const unassigned = rows.filter((w) => !String(w.assigned_artisan_name || "").trim()).length;
  return { ...section(rows), unassigned };
}

function partsOnOrder(db) {
  const rows = db.prepare(`
    SELECT po.id, p.part_code, p.part_name, p.critical, po.quantity, po.expected_date, po.status
    FROM parts_orders po
    JOIN parts p ON p.id = po.part_id
    WHERE COALESCE(po.status, '') != 'received'
    ORDER BY p.critical DESC,
      CASE WHEN po.expected_date IS NULL OR TRIM(po.expected_date) = '' THEN 1 ELSE 0 END,
      po.expected_date ASC
  `).all();
  return section(rows.map((r) => ({ ...r, critical: Boolean(r.critical) })));
}

/**
 * Builds the sections relevant to the caller's roles. Each section has a
 * total `count` and up to five preview `items`.
 */
export function buildMyWork(db, { user, roles, site = "main", today }) {
  const has = (list) => roles.some((r) => list.includes(r));
  const ctx = {
    user: String(user || "").trim().toLowerCase(),
    site: String(site || "main").trim().toLowerCase() || "main",
    today,
  };
  const sections = { tasks: myTasks(db, ctx) };
  if (has(MAINTENANCE_ROLES)) {
    sections.open_breakdowns = openBreakdowns(db, ctx);
    sections.open_work_orders = openWorkOrders(db, ctx);
  }
  if (has(STORES_ROLES)) {
    const waiting = workshopWaitingOnParts(db, ctx);
    sections.waiting_parts = { ...section(waiting), ...summarizeWaiting(waiting) };
    sections.low_stock = section(queryLowStock(db, 1000));
    sections.parts_on_order = partsOnOrder(db);
  }
  return { user: ctx.user, site: ctx.site, today, sections };
}
