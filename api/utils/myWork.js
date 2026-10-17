// IRONLOG/api/utils/myWork.js — role-aware "My Work" home screen data.
// Overdue services come from /api/maintenance/due on the client so counts
// match the Maintenance page. The low stock query is shared with /api/alerts.

import { overdueShifts } from "./techShift.js";
import { summarizeWaiting, workshopWaitingOnParts } from "./partsWaiting.js";
import { jobLine, myWorkOrders } from "./techPortal.js";

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
// Who sees what else on the home screen (ease of use, part 2: a home screen per role).
const MANAGER_ROLES = ["admin", "supervisor", "site_manager", "plant_manager", "executive"];
const FOREMAN_ROLES = ["admin", "supervisor", "workshop_admin", "plant_manager", "site_manager"];
const APPROVER_ROLES = ["admin", "supervisor"];
const TECH_ROLES = ["artisan"];
const SCHEDULED_HOURS = 10;
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

const hasTable = (db, t) => Boolean(db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?`).get(t));
const hasCol = (db, t, c) => db.prepare(`PRAGMA table_info(${t})`).all().some((r) => r.name === c);
const addDays = (ymd, n) => new Date(Date.parse(`${ymd}T00:00:00Z`) + n * 86400000).toISOString().slice(0, 10);

/**
 * A manager's figures for the site: availability over the last 7 full days (the same sum as
 * the KPI report, 10 scheduled hours per machine used), machines down now, and parts and
 * labour booked to work orders this month (USD).
 */
function siteFigures(db, { site, today }, machinesDown) {
  const start = addDays(today, -7);
  const end = addDays(today, -1);
  const aSite = hasCol(db, "assets", "site_code") ? "LOWER(TRIM(COALESCE(a.site_code, 'main')))" : "'main'";
  const used = hasTable(db, "daily_hours") ? db.prepare(`
    SELECT COUNT(DISTINCT dh.asset_id || '|' || dh.work_date) AS machine_days
    FROM daily_hours dh JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date BETWEEN ? AND ? AND dh.hours_run > 0 AND ${aSite} = ?
  `).get(start, end, site) : null;
  const available = Number(used?.machine_days || 0) * SCHEDULED_HOURS;
  let downtime = 0;
  if (hasTable(db, "breakdown_downtime_logs")) {
    downtime = Number(db.prepare(`
      SELECT IFNULL(SUM(l.hours_down), 0) AS h
      FROM breakdown_downtime_logs l JOIN breakdowns b ON b.id = l.breakdown_id JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date BETWEEN ? AND ? AND ${aSite} = ?
    `).get(start, end, site)?.h || 0);
  }
  const month = today.slice(0, 7);
  let cost = null;
  if (hasCol(db, "work_orders", "labor_hours")) {
    let rate = 35;
    try {
      const v = Number(db.prepare(`SELECT value FROM cost_settings WHERE key = 'labor_cost_per_hour_default'`).get()?.value);
      if (Number.isFinite(v) && v > 0) rate = v;
    } catch { /* no cost settings yet */ }
    const doneAt = hasCol(db, "work_orders", "completed_at") ? "COALESCE(w.completed_at, w.closed_at)" : "w.closed_at";
    const rateSql = hasCol(db, "work_orders", "labor_rate_per_hour") ? "COALESCE(w.labor_rate_per_hour, ?)" : "?";
    const labour = Number(db.prepare(`
      SELECT IFNULL(SUM(COALESCE(w.labor_hours, 0) * ${rateSql}), 0) AS c
      FROM work_orders w JOIN assets a ON a.id = w.asset_id
      WHERE SUBSTR(${doneAt}, 1, 7) = ? AND ${aSite} = ?
    `).get(rate, month, site)?.c || 0);
    const costCol = hasCol(db, "stock_movements", "unit_cost_usd") ? "COALESCE(m.unit_cost_usd, p.unit_cost, 0)" : "COALESCE(p.unit_cost, 0)";
    const parts = hasCol(db, "parts", "unit_cost") && hasCol(db, "stock_movements", "movement_type") ? Number(db.prepare(`
      SELECT IFNULL(SUM(ABS(m.quantity) * ${costCol}), 0) AS c
      FROM stock_movements m JOIN parts p ON p.id = m.part_id
      JOIN work_orders w ON m.reference = ('work_order:' || w.id) JOIN assets a ON a.id = w.asset_id
      WHERE m.movement_type = 'out' AND SUBSTR(m.created_at, 1, 7) = ? AND ${aSite} = ?
    `).get(month, site)?.c || 0) : 0;
    cost = Math.round((labour + parts) * 100) / 100;
  }
  return {
    period: { start, end },
    availability_pct: available > 0 ? Math.round(((available - downtime) / available) * 1000) / 10 : null,
    machines_down: machinesDown,
    month,
    work_order_cost_usd: cost,
  };
}

/** What is waiting for this person to approve: approval requests and completed work orders. */
function waitingApproval(db, { site }, { requests, workOrders }) {
  const items = [];
  if (requests && hasTable(db, "approval_requests")) {
    for (const r of db.prepare(`SELECT id, module, action, requested_by, created_at FROM approval_requests WHERE LOWER(status) = 'pending' ORDER BY id`).all()) {
      items.push({ kind: "request", id: r.id, what: `${r.module}: ${String(r.action).replace(/_/g, " ")}`, by: r.requested_by, since: r.created_at });
    }
  }
  if (workOrders) {
    const doneAt = hasCol(db, "work_orders", "completed_at") ? "w.completed_at" : "w.closed_at";
    for (const w of db.prepare(`
      SELECT w.id, a.asset_code, w.assigned_artisan_name, ${doneAt} AS done_at
      FROM work_orders w JOIN assets a ON a.id = w.asset_id
      WHERE LOWER(TRIM(COALESCE(w.status, ''))) = 'completed' AND LOWER(TRIM(COALESCE(w.site_code, 'main'))) = ?
      ORDER BY w.id
    `).all(site)) {
      items.push({ kind: "work_order", id: w.id, what: `WO ${w.id} — ${w.asset_code} completed`, by: w.assigned_artisan_name, since: w.done_at });
    }
  }
  items.sort((x, y) => String(x.since || "").localeCompare(String(y.since || "")));
  return { ...section(items), requests: items.filter((i) => i.kind === "request").length, work_orders: items.filter((i) => i.kind === "work_order").length };
}

/** Technicians on shift at the site now, and whether each is busy on a job or free. */
function teamOnShift(db, { site }) {
  if (!hasTable(db, "tech_shifts")) return section([]);
  const hasUsers = hasTable(db, "users");
  const shifts = db.prepare(`
    SELECT s.username, s.started_at, ${hasUsers ? "u.full_name" : "NULL"} AS full_name
    FROM tech_shifts s ${hasUsers ? "LEFT JOIN users u ON LOWER(u.username) = LOWER(s.username)" : ""}
    WHERE s.status = 'open' AND LOWER(TRIM(COALESCE(s.site_code, 'main'))) = ?
    ORDER BY s.started_at
  `).all(site);
  const people = shifts.map((s) => {
    let jobs = [];
    try { jobs = myWorkOrders(db, s.username); } catch { jobs = []; }
    const busy = jobs.find((j) => String(j.status || "").toLowerCase().replace(/\s+/g, "_") === "in_progress");
    return { username: s.username, name: String(s.full_name || "").trim() || s.username, since: s.started_at, busy_on: busy ? `WO ${busy.id} — ${busy.asset_code}` : null, jobs: jobs.length };
  });
  people.sort((x, y) => Number(Boolean(x.busy_on)) - Number(Boolean(y.busy_on)));
  return { ...section(people), free: people.filter((p) => !p.busy_on).length };
}

/** A technician's own open jobs. */
function myJobs(db, { user, today }) {
  let rows = [];
  try { rows = myWorkOrders(db, user); } catch { rows = []; }
  const jobs = rows.map((w) => ({ id: w.id, asset_code: w.asset_code, line: jobLine(w), status: w.status, role: w.role, due_date: w.due_date || null, critical: Boolean(w.breakdown_critical) }));
  return { ...section(jobs), urgent: jobs.filter((j) => j.critical || (j.due_date && j.due_date <= today)).length };
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
    sections.open_shifts = section(overdueShifts(db, { site: ctx.site }));
  }
  if (has(MANAGER_ROLES)) {
    const down = sections.open_breakdowns ? sections.open_breakdowns.count : openBreakdowns(db, ctx).count;
    sections.figures = siteFigures(db, ctx, down);
  }
  if (has(APPROVER_ROLES) || has(FOREMAN_ROLES)) sections.approvals = waitingApproval(db, ctx, { requests: has(APPROVER_ROLES), workOrders: has(FOREMAN_ROLES) });
  if (has(FOREMAN_ROLES)) sections.team = teamOnShift(db, ctx);
  if (has(TECH_ROLES)) sections.my_jobs = myJobs(db, ctx);
  if (has(STORES_ROLES)) {
    const waiting = workshopWaitingOnParts(db, ctx);
    sections.waiting_parts = { ...section(waiting), ...summarizeWaiting(waiting) };
    sections.low_stock = section(queryLowStock(db, 1000));
    sections.parts_on_order = partsOnOrder(db);
  }
  return { user: ctx.user, site: ctx.site, today, sections };
}
