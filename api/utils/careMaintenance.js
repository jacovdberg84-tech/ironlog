// IRONLOG/api/utils/careMaintenance.js — care and maintenance (C&M) plans.
//
// Before the fleet goes on care and maintenance, the planner walks every
// machine and writes down what it needs: the work, the parts (with prices)
// and the repair hours. A plan holds one row per machine with that estimate.
// When work starts, a machine gets a normal work order, so technicians book
// time in the portal and stores issue parts against it as usual; the plan
// then reads the hours and parts actually used from that work order and
// shows them next to the estimate. Nothing here keeps a second copy of time
// or stock.
import { hasColumn, hasTable, stockInfo, workOrderLabourHours } from "./techPortal.js";
import { ensureServiceTemplateSchema } from "./serviceTemplates.js";

export const CM_STATUSES = {
  not_started: "Not started",
  waiting_parts: "Waiting for parts",
  in_progress: "In progress",
  done: "Done - ready to run",
  parked: "Parked (preserved)",
};
export const CM_PRIORITIES = ["high", "normal", "low"];
const FINISHED = ["completed", "approved", "closed"];

const round2 = (n) => Number(Number(n || 0).toFixed(2));
const num = (v, fallback = 0) => (Number.isFinite(Number(v)) && String(v ?? "").trim() !== "" ? Number(v) : fallback);
const clean = (v, max = 500) => String(v ?? "").trim().slice(0, max);

export function ensureCareMaintenanceTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS cm_plans (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL DEFAULT 'main',
      name TEXT NOT NULL,
      start_date TEXT,
      end_date TEXT,
      labour_rate REAL,
      notes TEXT,
      status TEXT NOT NULL DEFAULT 'active',
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS cm_plan_machines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      plan_id INTEGER NOT NULL,
      asset_id INTEGER NOT NULL,
      scope TEXT,
      planned_hours REAL NOT NULL DEFAULT 0,
      status TEXT NOT NULL DEFAULT 'not_started',
      priority TEXT NOT NULL DEFAULT 'normal',
      notes TEXT,
      work_order_id INTEGER,
      updated_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(plan_id, asset_id)
    );
    CREATE INDEX IF NOT EXISTS idx_cm_plan_machines_plan ON cm_plan_machines(plan_id);
    CREATE TABLE IF NOT EXISTS cm_plan_parts (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      machine_row_id INTEGER NOT NULL,
      part_id INTEGER,
      part_code TEXT,
      description TEXT NOT NULL,
      qty REAL NOT NULL DEFAULT 1,
      unit_cost REAL NOT NULL DEFAULT 0,
      request_id INTEGER,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_cm_plan_parts_row ON cm_plan_parts(machine_row_id);
  `);
}

/** The workshop's default labour rate (USD per hour), as used on work orders. */
export function defaultLabourRate(db) {
  try {
    const v = Number(db.prepare(`SELECT value FROM cost_settings WHERE key = 'labor_cost_per_hour_default' LIMIT 1`).get()?.value);
    return Number.isFinite(v) && v > 0 ? v : 35;
  } catch {
    return 35;
  }
}

const planRate = (db, plan) => (Number(plan?.labour_rate) > 0 ? Number(plan.labour_rate) : defaultLabourRate(db));

export function listPlans(db, site) {
  ensureCareMaintenanceTables(db);
  return db.prepare(`
    SELECT p.*, (SELECT COUNT(*) FROM cm_plan_machines m WHERE m.plan_id = p.id) AS machines
    FROM cm_plans p WHERE p.site_code = ? ORDER BY CASE p.status WHEN 'active' THEN 0 ELSE 1 END, p.id DESC
  `).all(site);
}

export function getPlan(db, planId, site) {
  ensureCareMaintenanceTables(db);
  return db.prepare(`SELECT * FROM cm_plans WHERE id = ? AND site_code = ?`).get(Number(planId), site) || null;
}

/** A machine row with its plan, only when the plan belongs to this site. */
export function getMachineRow(db, rowId, site) {
  ensureCareMaintenanceTables(db);
  return db.prepare(`
    SELECT m.*, p.name AS plan_name, p.labour_rate, p.site_code, a.asset_code, a.asset_name
    FROM cm_plan_machines m JOIN cm_plans p ON p.id = m.plan_id LEFT JOIN assets a ON a.id = m.asset_id
    WHERE m.id = ? AND p.site_code = ?
  `).get(Number(rowId), site) || null;
}

function planFields(body) {
  const out = {};
  if (body.name !== undefined) out.name = clean(body.name, 120);
  for (const k of ["start_date", "end_date"]) {
    if (body[k] !== undefined) out[k] = /^\d{4}-\d{2}-\d{2}$/.test(String(body[k] || "")) ? String(body[k]) : null;
  }
  if (body.labour_rate !== undefined) out.labour_rate = num(body.labour_rate, 0) > 0 ? num(body.labour_rate) : null;
  if (body.notes !== undefined) out.notes = clean(body.notes, 2000) || null;
  if (body.status !== undefined) out.status = body.status === "closed" ? "closed" : "active";
  return out;
}

export function createPlan(db, site, body = {}, actor = "") {
  ensureCareMaintenanceTables(db);
  const f = planFields(body);
  if (!f.name) return { ok: false, error: "Give the plan a name, for example \"C&M November 2026\"." };
  const info = db.prepare(`
    INSERT INTO cm_plans (site_code, name, start_date, end_date, labour_rate, notes, created_by)
    VALUES (?, ?, ?, ?, ?, ?, ?)
  `).run(site, f.name, f.start_date ?? null, f.end_date ?? null, f.labour_rate ?? null, f.notes ?? null, actor || null);
  return { ok: true, id: Number(info.lastInsertRowid) };
}

export function updatePlan(db, planId, body = {}) {
  const f = planFields(body);
  if ("name" in f && !f.name) return { ok: false, error: "The plan needs a name." };
  const cols = Object.keys(f);
  if (cols.length) {
    db.prepare(`UPDATE cm_plans SET ${cols.map((c) => `${c} = ?`).join(", ")}, updated_at = datetime('now') WHERE id = ?`)
      .run(...cols.map((c) => f[c]), Number(planId));
  }
  return { ok: true };
}

/** Add machines to a plan; machines already on it are skipped. */
export function addMachines(db, planId, assetIds = [], actor = "") {
  const ids = [...new Set((assetIds || []).map(Number).filter((n) => Number.isInteger(n) && n > 0))];
  const known = ids.length
    ? new Set(db.prepare(`SELECT id FROM assets WHERE id IN (${ids.map(() => "?").join(",")})`).all(...ids).map((r) => r.id))
    : new Set();
  const ins = db.prepare(`INSERT OR IGNORE INTO cm_plan_machines (plan_id, asset_id, updated_by) VALUES (?, ?, ?)`);
  let added = 0;
  db.transaction(() => {
    for (const id of ids) if (known.has(id)) added += ins.run(Number(planId), id, actor || null).changes;
  })();
  return { ok: true, added, skipped: ids.length - added };
}

export function updateMachine(db, rowId, body = {}, actor = "") {
  const sets = {};
  if (body.scope !== undefined) sets.scope = clean(body.scope, 4000) || null;
  if (body.planned_hours !== undefined) sets.planned_hours = Math.max(0, num(body.planned_hours, 0));
  if (body.status !== undefined) {
    if (!CM_STATUSES[body.status]) return { ok: false, error: "Unknown status." };
    sets.status = body.status;
  }
  if (body.priority !== undefined) sets.priority = CM_PRIORITIES.includes(body.priority) ? body.priority : "normal";
  if (body.notes !== undefined) sets.notes = clean(body.notes, 2000) || null;
  const cols = Object.keys(sets);
  if (cols.length) {
    db.prepare(`UPDATE cm_plan_machines SET ${cols.map((c) => `${c} = ?`).join(", ")}, updated_by = ?, updated_at = datetime('now') WHERE id = ?`)
      .run(...cols.map((c) => sets[c]), actor || null, Number(rowId));
  }
  return { ok: true };
}

export function removeMachine(db, rowId) {
  db.transaction(() => {
    db.prepare(`DELETE FROM cm_plan_parts WHERE machine_row_id = ?`).run(Number(rowId));
    db.prepare(`DELETE FROM cm_plan_machines WHERE id = ?`).run(Number(rowId));
  })();
  return { ok: true };
}

/**
 * Replace a machine's parts list. A line whose code is a stores part is linked
 * to it (and takes its price when none is given); other lines are free text,
 * for parts the stores do not carry yet.
 */
export function setParts(db, rowId, lines = []) {
  if (!Array.isArray(lines)) return { ok: false, error: "Send the parts as a list." };
  const byCode = hasTable(db, "parts") ? db.prepare(`SELECT * FROM parts WHERE UPPER(TRIM(part_code)) = UPPER(TRIM(?))`) : null;
  const rows = [];
  for (const raw of lines.slice(0, 300)) {
    const code = clean(raw?.part_code, 80);
    const part = code && byCode ? byCode.get(code) : null;
    const description = clean(raw?.description, 300) || part?.part_name || code;
    if (!description) continue;
    const qty = num(raw?.qty, 1);
    if (!(qty > 0)) return { ok: false, error: `Quantity for "${description}" must be more than 0.` };
    const given = String(raw?.unit_cost ?? "").trim() !== "";
    const unit = given ? Math.max(0, num(raw.unit_cost, 0)) : Number(part?.unit_cost || 0);
    rows.push({ part_id: part?.id ?? null, part_code: part?.part_code || code || null, description, qty, unit_cost: round2(unit), request_id: Number(raw?.request_id) || null });
  }
  const ins = db.prepare(`
    INSERT INTO cm_plan_parts (machine_row_id, part_id, part_code, description, qty, unit_cost, request_id)
    VALUES (?, ?, ?, ?, ?, ?, ?)
  `);
  db.transaction(() => {
    db.prepare(`DELETE FROM cm_plan_parts WHERE machine_row_id = ?`).run(Number(rowId));
    for (const r of rows) ins.run(Number(rowId), r.part_id, r.part_code, r.description, r.qty, r.unit_cost, r.request_id);
  })();
  return { ok: true, lines: rows.length };
}

/** Hours actually booked on a work order: portal time plus hand-entered timesheet rows. */
export function workOrderHours(db, wo, nowIso = new Date().toISOString()) {
  let hours = hasTable(db, "tech_activity_events") ? workOrderLabourHours(db, wo.id, nowIso) : 0;
  // Portal time is already counted above; its timesheet rows are marked portal:<user>.
  if (hasColumn(db, "mechanic_labor_entries", "job_card_no")) {
    hours += Number(db.prepare(`
      SELECT COALESCE(SUM(hours), 0) AS h FROM mechanic_labor_entries
      WHERE TRIM(job_card_no) = ? AND COALESCE(created_by, '') NOT LIKE 'portal:%'
    `).get(String(wo.id)).h || 0);
  }
  // Older jobs closed with typed hours only.
  if (hours <= 0) hours = Number(wo.labor_hours || 0);
  return round2(hours);
}

/** Parts issued against a work order, per part, with their cost. */
function issuedParts(db, woId) {
  if (!hasTable(db, "stock_movements") || !hasTable(db, "parts")) return [];
  const cost = hasColumn(db, "stock_movements", "unit_cost_usd") ? "COALESCE(NULLIF(sm.unit_cost_usd, 0), p.unit_cost, 0)" : "COALESCE(p.unit_cost, 0)";
  return db.prepare(`
    SELECT p.id AS part_id, p.part_code, p.part_name, -SUM(sm.quantity) AS qty, -SUM(sm.quantity * ${cost}) AS cost
    FROM stock_movements sm JOIN parts p ON p.id = sm.part_id
    WHERE sm.reference = ? GROUP BY p.id HAVING qty > 0
  `).all(`work_order:${Number(woId)}`).map((r) => ({ ...r, qty: Number(r.qty), cost: round2(r.cost) }));
}

function openRequests(db, ids) {
  if (!ids.length || !hasTable(db, "maintenance_parts_requests")) return new Map();
  return new Map(db.prepare(`SELECT id, status FROM maintenance_parts_requests WHERE id IN (${ids.map(() => "?").join(",")})`)
    .all(...ids).map((r) => [r.id, String(r.status || "requested").toLowerCase()]));
}

/** One machine row with its estimate, what has been used so far, and parts status. */
function machineView(db, row, rate, nowIso) {
  const parts = db.prepare(`SELECT * FROM cm_plan_parts WHERE machine_row_id = ? ORDER BY id`).all(row.id);
  const wo = row.work_order_id && hasTable(db, "work_orders")
    ? db.prepare(`SELECT * FROM work_orders WHERE id = ?`).get(row.work_order_id)
    : null;
  const issued = wo ? issuedParts(db, wo.id) : [];
  const issuedBy = new Map(issued.map((i) => [i.part_id, i.qty]));
  const stock = stockInfo(db, [...new Set(parts.map((p) => p.part_id).filter(Boolean))]);
  const requests = openRequests(db, parts.map((p) => p.request_id).filter(Boolean));
  const lines = parts.map((p) => {
    const got = p.part_id ? issuedBy.get(p.part_id) || 0 : 0;
    const need = Math.max(0, Number(p.qty) - got);
    const onHand = p.part_id ? stock.get(p.part_id)?.on_hand ?? 0 : null;
    const request = p.request_id ? requests.get(p.request_id) || null : null;
    let state;
    if (need <= 0) state = "issued";
    else if (onHand != null && onHand >= need) state = "in_stock";
    else if (request === "ordered") state = "ordered";
    else if (request === "requested") state = "requested";
    else if (request === "received") state = "received";
    else state = "short";
    return {
      id: p.id, part_id: p.part_id, part_code: p.part_code, description: p.description,
      qty: Number(p.qty), unit_cost: round2(p.unit_cost), line_cost: round2(Number(p.qty) * Number(p.unit_cost)),
      issued: got, on_hand: onHand, request_id: p.request_id, request_status: request, state,
    };
  });
  const plannedParts = round2(lines.reduce((s, l) => s + l.line_cost, 0));
  const plannedLabour = round2(Number(row.planned_hours || 0) * rate);
  const actualHours = wo ? workOrderHours(db, wo, nowIso) : 0;
  const actualParts = round2(issued.reduce((s, i) => s + i.cost, 0));
  const actualLabour = round2(actualHours * rate);
  const plannedHours = Number(row.planned_hours || 0);
  return {
    id: row.id,
    asset_id: row.asset_id,
    asset_code: row.asset_code,
    asset_name: row.asset_name,
    category: row.category,
    scope: row.scope || "",
    priority: row.priority,
    status: row.status,
    status_label: CM_STATUSES[row.status] || row.status,
    notes: row.notes || "",
    updated_at: row.updated_at,
    updated_by: row.updated_by,
    work_order: wo ? {
      id: wo.id,
      status: wo.status,
      finished: FINISHED.includes(String(wo.status || "").toLowerCase()),
      technician: wo.assigned_artisan_name || null,
    } : null,
    parts: lines,
    issued_parts: issued,
    planned: { hours: plannedHours, parts_cost: plannedParts, labour_cost: plannedLabour, total: round2(plannedParts + plannedLabour) },
    actual: { hours: actualHours, parts_cost: actualParts, labour_cost: actualLabour, total: round2(actualParts + actualLabour) },
    hours_used_pct: plannedHours > 0 ? Math.round((actualHours / plannedHours) * 100) : null,
    parts_short: lines.filter((l) => ["short", "requested", "ordered"].includes(l.state)).length,
  };
}

/** The plan with every machine, and totals for the whole plan. */
export function planDetail(db, planId, site, { now = new Date() } = {}) {
  const plan = getPlan(db, planId, site);
  if (!plan) return null;
  const rate = planRate(db, plan);
  const nowIso = now.toISOString();
  const catCol = hasColumn(db, "assets", "category") ? "a.category" : "NULL";
  const rows = db.prepare(`
    SELECT m.*, a.asset_code, a.asset_name, ${catCol} AS category FROM cm_plan_machines m
    LEFT JOIN assets a ON a.id = m.asset_id
    WHERE m.plan_id = ?
    ORDER BY CASE m.priority WHEN 'high' THEN 0 WHEN 'normal' THEN 1 ELSE 2 END, a.asset_code
  `).all(plan.id);
  const machines = rows.map((r) => machineView(db, r, rate, nowIso));
  const sum = (pick) => round2(machines.reduce((s, m) => s + pick(m), 0));
  const byStatus = Object.fromEntries(Object.keys(CM_STATUSES).map((k) => [k, machines.filter((m) => m.status === k).length]));
  return {
    plan: { ...plan, labour_rate_used: rate },
    statuses: CM_STATUSES,
    machines,
    totals: {
      machines: machines.length,
      by_status: byStatus,
      planned_hours: sum((m) => m.planned.hours),
      actual_hours: sum((m) => m.actual.hours),
      planned_parts: sum((m) => m.planned.parts_cost),
      planned_labour: sum((m) => m.planned.labour_cost),
      planned_total: sum((m) => m.planned.total),
      actual_parts: sum((m) => m.actual.parts_cost),
      actual_labour: sum((m) => m.actual.labour_cost),
      actual_total: sum((m) => m.actual.total),
      parts_short: machines.reduce((s, m) => s + m.parts_short, 0),
    },
  };
}

/**
 * Open a work order for a machine on the plan, carrying its scope, planned
 * hours and the stores parts, so technicians see it in the portal and stores
 * can issue against it.
 */
export function createMachineWorkOrder(db, rowId, site, actor = "") {
  const row = getMachineRow(db, rowId, site);
  if (!row) return { ok: false, status: 404, error: "Machine not found on this plan." };
  if (row.work_order_id) {
    const wo = db.prepare(`SELECT id, status FROM work_orders WHERE id = ?`).get(row.work_order_id);
    if (wo && !FINISHED.includes(String(wo.status || "").toLowerCase())) {
      return { ok: false, status: 409, error: `This machine already has WO #${wo.id} open.`, work_order_id: wo.id };
    }
  }
  ensureServiceTemplateSchema(db);
  const parts = db.prepare(`SELECT * FROM cm_plan_parts WHERE machine_row_id = ? ORDER BY id`).all(row.id);
  const rate = planRate(db, row);
  const lines = [`Care & maintenance: ${row.plan_name}`];
  if (row.scope) lines.push(row.scope);
  if (parts.length) lines.push("", "Parts planned:", ...parts.map((p) => `- ${p.qty} x ${p.part_code ? `${p.part_code} ` : ""}${p.description}`));
  let woId;
  db.transaction(() => {
    const cols = ["asset_id", "source", "reference_id", "status"];
    const vals = [row.asset_id, "care_maintenance", row.id, "open"];
    const opt = {
      site_code: site,
      job_description: lines.join("\n"),
      planned_labor_hours: Number(row.planned_hours || 0) || null,
      planned_labor_rate: rate,
      labor_rate_per_hour: rate,
    };
    for (const [c, v] of Object.entries(opt)) {
      if (hasColumn(db, "work_orders", c)) { cols.push(c); vals.push(v); }
    }
    woId = Number(db.prepare(`INSERT INTO work_orders (${cols.join(", ")}) VALUES (${cols.map(() => "?").join(", ")})`).run(...vals).lastInsertRowid);
    const pm = db.prepare(`
      INSERT INTO work_order_planned_materials (work_order_id, part_id, part_code, description, quantity_planned, unit_cost_snapshot, planned_line_total)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `);
    for (const p of parts) {
      if (p.part_id) pm.run(woId, p.part_id, p.part_code, p.description, p.qty, p.unit_cost, round2(p.qty * p.unit_cost));
    }
    // Any parts requests already raised for this machine follow it onto the job.
    const reqIds = parts.map((p) => p.request_id).filter(Boolean);
    if (reqIds.length && hasTable(db, "maintenance_parts_requests")) {
      db.prepare(`UPDATE maintenance_parts_requests SET work_order_id = ? WHERE id IN (${reqIds.map(() => "?").join(",")}) AND work_order_id IS NULL`)
        .run(woId, ...reqIds);
    }
    db.prepare(`UPDATE cm_plan_machines SET work_order_id = ?, updated_by = ?, updated_at = datetime('now') WHERE id = ?`).run(woId, actor || null, row.id);
  })();
  return { ok: true, work_order_id: woId };
}

/**
 * Raise parts requests (Parts to Order) for every line the stores cannot cover
 * yet. Lines that already have an open request are left alone.
 */
export function requestMissingParts(db, rowId, site, actor = "") {
  const row = getMachineRow(db, rowId, site);
  if (!row) return { ok: false, status: 404, error: "Machine not found on this plan." };
  if (!hasTable(db, "maintenance_parts_requests")) return { ok: false, status: 409, error: "Parts requests are not set up on this server yet." };
  const view = machineView(db, { ...row, category: null }, planRate(db, row), new Date().toISOString());
  const todo = view.parts.filter((l) => l.state === "short");
  const ins = db.prepare(`
    INSERT INTO maintenance_parts_requests (site_code, asset_id, asset_code, part_code, part_name, qty, urgency, notes, work_order_id, status, requested_by)
    VALUES (?, ?, ?, ?, ?, ?, 'normal', ?, ?, 'requested', ?)
  `);
  const link = db.prepare(`UPDATE cm_plan_parts SET request_id = ? WHERE id = ?`);
  const created = [];
  db.transaction(() => {
    for (const l of todo) {
      const qty = round2(Math.max(0, l.qty - l.issued - Math.max(0, l.on_hand || 0))) || l.qty;
      const id = Number(ins.run(site, row.asset_id, row.asset_code, l.part_code, l.description, qty,
        `Care & maintenance: ${row.plan_name}`, row.work_order_id || null, actor || "system").lastInsertRowid);
      link.run(id, l.id);
      created.push({ request_id: id, part_code: l.part_code, description: l.description, qty });
    }
  })();
  return { ok: true, created };
}

/** The plan as CSV rows (one per machine), for Excel. */
export function planCsv(detail) {
  const q = (v) => {
    const s = String(v ?? "");
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const head = ["Machine", "Description", "Priority", "Status", "Work order", "Work to do", "Parts planned (USD)", "Hours planned", "Labour planned (USD)", "Total planned (USD)",
    "Hours used", "Parts used (USD)", "Labour used (USD)", "Total used (USD)", "Parts still needed", "Notes"];
  const rows = detail.machines.map((m) => [
    m.asset_code, m.asset_name, m.priority, m.status_label, m.work_order ? `WO #${m.work_order.id} (${m.work_order.status})` : "",
    m.scope, m.planned.parts_cost, m.planned.hours, m.planned.labour_cost, m.planned.total,
    m.actual.hours, m.actual.parts_cost, m.actual.labour_cost, m.actual.total, m.parts_short, m.notes,
  ]);
  const t = detail.totals;
  rows.push(["TOTAL", `${t.machines} machines`, "", "", "", "", t.planned_parts, t.planned_hours, t.planned_labour, t.planned_total, t.actual_hours, t.actual_parts, t.actual_labour, t.actual_total, t.parts_short, ""]);
  return [head, ...rows].map((r) => r.map(q).join(",")).join("\r\n");
}
