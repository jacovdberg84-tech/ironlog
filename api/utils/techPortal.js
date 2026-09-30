// IRONLOG/api/utils/techPortal.js — read-only building blocks for the
// technician portal. Everything here reads the existing work order, stock,
// parts request, breakdown and service tables; the portal adds no second copy.

import { technicianIdentityKeys } from "./technicianIdentity.js";
import { buildSegments, labourHours, statesByUser } from "./techActivity.js";

const FINISHED = ["completed", "approved", "closed"];
const woStatusSql = (a = "w") => `REPLACE(TRIM(LOWER(COALESCE(${a}.status, ''))), ' ', '_')`;

export function hasTable(db, name) {
  return Boolean(db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?`).get(name));
}

export function hasColumn(db, table, col) {
  return hasTable(db, table) && db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
}

/** Keys that identify the signed-in technician on work orders and helper rows. */
export function myKeys(db, username) {
  return [...technicianIdentityKeys(db, username)];
}

/**
 * Work orders the technician leads (assigned) or helps on. Includes unfinished
 * jobs plus jobs finished since `finishedSince` (for "completed today").
 */
export function myWorkOrders(db, username, { finishedSince = null } = {}) {
  const keys = myKeys(db, username);
  if (!keys.length) return [];
  const marks = keys.map(() => "?").join(", ");
  const helperSql = hasTable(db, "work_order_technicians")
    ? `OR w.id IN (SELECT work_order_id FROM work_order_technicians WHERE LOWER(TRIM(username)) IN (${marks}))`
    : "";
  const doneAt = hasColumn(db, "work_orders", "completed_at") ? "COALESCE(w.completed_at, w.closed_at)" : "w.closed_at";
  const rows = db.prepare(`
    SELECT w.id, w.asset_id, w.source, w.reference_id, w.status, w.opened_at, w.assigned_artisan_name,
      w.assigned_at, w.started_at, ${doneAt} AS done_at, w.due_date, w.priority, w.job_description,
      w.repair_progress, w.completion_notes, w.labor_hours,
      a.asset_code, a.asset_name, a.category,
      b.description AS breakdown_description, b.component AS breakdown_component, b.critical AS breakdown_critical,
      b.parts_status, b.ets_repair_date,
      mp.service_name, mp.interval_hours
    FROM work_orders w
    JOIN assets a ON a.id = w.asset_id
    LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
    LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id AND w.source = 'service'
    WHERE (LOWER(TRIM(COALESCE(w.assigned_artisan_name, ''))) IN (${marks}) ${helperSql})
      AND (${woStatusSql()} NOT IN (${FINISHED.map((s) => `'${s}'`).join(", ")})
           OR (? IS NOT NULL AND ${doneAt} >= ?))
    ORDER BY w.id
  `).all(...keys, ...(helperSql ? keys : []), finishedSince, finishedSince);
  return rows.map((r) => ({ ...r, role: keys.includes(String(r.assigned_artisan_name || "").trim().toLowerCase()) ? "lead" : "helper" }));
}

/** One line that says what the job is. */
export function jobLine(w) {
  const source = String(w.source || "").toLowerCase();
  if (source === "breakdown") return [w.breakdown_component, w.breakdown_description].filter(Boolean).join(" — ") || "Breakdown repair";
  if (source === "service") return w.service_name ? `${w.service_name}${/^\d+$/.test(String(w.service_name).trim()) ? " h service" : ""}` : "Scheduled service";
  const lines = String(w.job_description || "").split("\n").map((l) => l.replace(/^-\s*/, "").trim()).filter(Boolean);
  if (source === "prestart" && lines.length > 1) return `Pre-start: ${lines.slice(1).join("; ")}`;
  return lines[lines.length - 1] || "Repair";
}

export function isUrgent(w, today) {
  const p = String(w.priority || "").toLowerCase();
  if (Number(w.breakdown_critical)) return true;
  if (["p1", "high", "critical", "urgent"].includes(p)) return true;
  if (String(w.source || "").toLowerCase() === "breakdown") return true;
  const due = String(w.due_date || "").slice(0, 10);
  return Boolean(due && due < today);
}

/** Events for the given work orders, oldest first. */
export function eventsFor(db, woIds) {
  if (!woIds.length || !hasTable(db, "tech_activity_events")) return [];
  const marks = woIds.map(() => "?").join(", ");
  return db.prepare(`
    SELECT id, work_order_id, username, action, at, note, auto
    FROM tech_activity_events WHERE work_order_id IN (${marks})
    ORDER BY at, id
  `).all(...woIds);
}

/** Every event by one technician since `sinceIso`, oldest first. */
export function eventsByUser(db, username, sinceIso) {
  if (!hasTable(db, "tech_activity_events")) return [];
  return db.prepare(`
    SELECT id, work_order_id, username, action, at, note, auto
    FROM tech_activity_events
    WHERE LOWER(username) = LOWER(?) AND at >= ?
    ORDER BY at, id
  `).all(username, sinceIso);
}

/** This technician's state on each work order (from all events on those orders). */
export function myStates(db, username, woIds) {
  const byWo = new Map();
  const events = eventsFor(db, woIds);
  for (const id of woIds) {
    const states = statesByUser(events.filter((e) => e.work_order_id === id));
    byWo.set(id, { mine: states.get(username) || "idle", all: states });
  }
  return { byWo, events };
}

/** Total labour on a work order from everyone's activity (for WO labour hours). */
export function workOrderLabourHours(db, woId, nowIso = new Date().toISOString()) {
  return labourHours(buildSegments(eventsFor(db, [Number(woId)]), { until: nowIso }));
}

/** Stock on hand and the bin it was last received into, per part id. */
export function stockInfo(db, partIds) {
  const out = new Map();
  if (!partIds.length || !hasTable(db, "stock_movements")) return out;
  const marks = partIds.map(() => "?").join(", ");
  for (const r of db.prepare(`
    SELECT part_id, COALESCE(SUM(quantity), 0) AS on_hand FROM stock_movements
    WHERE part_id IN (${marks}) GROUP BY part_id
  `).all(...partIds)) out.set(r.part_id, { on_hand: Number(r.on_hand), bin: null });
  if (hasColumn(db, "stock_movements", "bin_id") && hasTable(db, "stock_bins")) {
    const locJoin = hasTable(db, "stock_locations") ? "LEFT JOIN stock_locations l ON l.id = b.location_id" : "";
    const locCol = hasTable(db, "stock_locations") ? "l.location_code" : "NULL";
    for (const r of db.prepare(`
      SELECT sm.part_id, b.bin_code, ${locCol} AS location_code
      FROM stock_movements sm
      JOIN stock_bins b ON b.id = sm.bin_id
      ${locJoin}
      WHERE sm.part_id IN (${marks}) AND sm.quantity > 0
      ORDER BY sm.id DESC
    `).all(...partIds)) {
      const e = out.get(r.part_id) || { on_hand: 0, bin: null };
      if (!e.bin) e.bin = [r.location_code, r.bin_code].filter(Boolean).join(" / ");
      out.set(r.part_id, e);
    }
  }
  return out;
}

/**
 * Parts for a work order: planned lines (service kit / estimate), what stores
 * issued, requests to stores, and how each planned line stands.
 */
export function workOrderParts(db, woId) {
  const id = Number(woId);
  const planned = hasTable(db, "work_order_planned_materials")
    ? db.prepare(`
        SELECT pm.id, pm.part_id, COALESCE(p.part_code, pm.part_code) AS part_code, COALESCE(p.part_name, pm.description) AS description,
          pm.quantity_planned, pm.unit_of_measure, pm.item_type
        FROM work_order_planned_materials pm LEFT JOIN parts p ON p.id = pm.part_id
        WHERE pm.work_order_id = ? ORDER BY pm.id
      `).all(id)
    : [];
  const issued = hasTable(db, "stock_movements")
    ? db.prepare(`
        SELECT p.id AS part_id, p.part_code, p.part_name, -SUM(sm.quantity) AS qty
        FROM stock_movements sm JOIN parts p ON p.id = sm.part_id
        WHERE sm.reference = ? GROUP BY p.id HAVING qty > 0 ORDER BY p.part_code
      `).all(`work_order:${id}`)
    : [];
  const requests = hasTable(db, "maintenance_parts_requests")
    ? db.prepare(`
        SELECT id, part_code, part_name, qty, urgency, status, requested_by, notes, status_notes, created_at, updated_at
        FROM maintenance_parts_requests WHERE work_order_id = ?
        ORDER BY CASE LOWER(COALESCE(status, 'requested')) WHEN 'requested' THEN 0 WHEN 'ordered' THEN 1 WHEN 'received' THEN 2 ELSE 3 END, id DESC
      `).all(id)
    : [];
  const partIds = [...new Set([...planned.map((p) => p.part_id), ...issued.map((p) => p.part_id)].filter(Boolean))];
  const stock = stockInfo(db, partIds);
  const issuedBy = new Map(issued.map((i) => [i.part_id, Number(i.qty)]));
  const plannedLines = planned.map((p) => {
    const s = stock.get(p.part_id) || { on_hand: null, bin: null };
    const got = issuedBy.get(p.part_id) || 0;
    const need = Math.max(0, Number(p.quantity_planned || 0) - got);
    return {
      ...p,
      quantity_planned: Number(p.quantity_planned || 0),
      quantity_issued: got,
      on_hand: s.on_hand,
      bin: s.bin,
      state: need <= 0 ? "issued" : s.on_hand != null && s.on_hand >= need ? "in_stock" : "short",
    };
  });
  const issuedLines = issued.map((i) => ({ ...i, qty: Number(i.qty), bin: stock.get(i.part_id)?.bin || null }));
  const openRequests = requests.filter((r) => ["requested", "ordered"].includes(String(r.status || "requested").toLowerCase()));
  return {
    planned: plannedLines,
    issued: issuedLines,
    requests,
    summary: {
      planned: plannedLines.length,
      short: plannedLines.filter((l) => l.state === "short").length,
      open_requests: openRequests.length,
      received: requests.filter((r) => String(r.status || "").toLowerCase() === "received").length,
      waiting: plannedLines.some((l) => l.state === "short") || openRequests.length > 0,
    },
  };
}
