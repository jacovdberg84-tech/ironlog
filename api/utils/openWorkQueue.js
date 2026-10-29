// IRONLOG/api/utils/openWorkQueue.js — the open work-order queue with parts
// availability, for the administrator and Borris to weigh priorities.
//
// Read-only: every query here is a SELECT. Nothing is created, allocated,
// issued, received or closed. Stock on hand is the Stores ledger total (the
// same figure as Stock on hand); parts on order come from Stores part orders
// for the same site and from parts requests already marked ordered.

/** Who may see the whole queue: the leads who already manage work orders. */
export const OPEN_WORK_QUEUE_ROLES = ["admin", "supervisor", "workshop_admin", "plant_manager"];

const FINISHED_WO =["closed", "completed", "approved"];
const WAITING_REQUEST_STATUSES = ["requested", "ordered"];
const OPEN_PART_ORDER_STATUSES = ["on_order", "warehouse_ready", "in_transit"];

const list = (xs) => xs.map((x) => `'${x}'`).join(", ");

function hasTable(db, name) {
  return Boolean(db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?`).get(name));
}

function hasColumn(db, table, col) {
  return db.prepare(`PRAGMA table_info(${table})`).all().some((r) => String(r.name) === col);
}

function daysBetween(fromYmd, toYmd) {
  const a = Date.parse(`${String(fromYmd || "").slice(0, 10)}T00:00:00Z`);
  const b = Date.parse(`${String(toYmd || "").slice(0, 10)}T00:00:00Z`);
  if (!Number.isFinite(a) || !Number.isFinite(b)) return null;
  return Math.max(0, Math.round((b - a) / 86400000));
}

function firstLine(text, max = 160) {
  const line = String(text || "").split(/\r?\n/).map((x) => x.trim()).find(Boolean) || "";
  return line.length > max ? `${line.slice(0, max - 1)}…` : line;
}

/** Stock on hand, minimum and open Stores orders for one part code (case-insensitive). */
function stockLookup(db, site) {
  const hasParts = hasTable(db, "parts");
  const hasMoves = hasParts && hasTable(db, "stock_movements");
  const hasMin = hasParts && hasColumn(db, "parts", "min_stock");
  const partStmt = hasParts
    ? db.prepare(`
        SELECT p.id, p.part_code, p.part_name, ${hasMin ? "COALESCE(p.min_stock, 0)" : "0"} AS min_stock,
               ${hasMoves ? "(SELECT IFNULL(SUM(sm.quantity), 0) FROM stock_movements sm WHERE sm.part_id = p.id)" : "NULL"} AS on_hand
        FROM parts p
        WHERE LOWER(TRIM(p.part_code)) = LOWER(TRIM(?))
        LIMIT 1
      `)
    : null;
  const orderStmt = hasTable(db, "stores_part_orders")
    ? db.prepare(`
        SELECT IFNULL(SUM(COALESCE(o.supplier_outstanding_qty, o.qty, 0)), 0) AS qty,
               MIN(NULLIF(TRIM(o.expected_arrival_date), '')) AS expected
        FROM stores_part_orders o
        WHERE LOWER(TRIM(COALESCE(o.site_code, 'main'))) = ?
          AND LOWER(COALESCE(o.status, 'on_order')) IN (${list(OPEN_PART_ORDER_STATUSES)})
          AND (o.part_id = ? OR LOWER(TRIM(COALESCE(o.part_code, ''))) = LOWER(TRIM(?)))
      `)
    : null;
  const cache = new Map();
  return (code) => {
    const key = String(code || "").trim().toLowerCase();
    if (!key) return null;
    if (cache.has(key)) return cache.get(key);
    const part = partStmt ? partStmt.get(key) : null;
    const order = orderStmt ? orderStmt.get(site, part ? part.id : -1, key) : null;
    const out = {
      known: Boolean(part),
      on_hand: part && part.on_hand != null ? Number(part.on_hand) : null,
      min_stock: part ? Number(part.min_stock || 0) : null,
      on_order_qty: order ? Number(order.qty || 0) : 0,
      expected_arrival: order?.expected || null,
    };
    cache.set(key, out);
    return out;
  };
}

/**
 * Every open work order on the site, with its machine, job, age, status,
 * technician and outstanding parts. Each part shows stock on hand, whether
 * Stores is below minimum and whether it is on order. `parts_state` is
 * 'none' (no parts waiting), 'ready' (stock covers every part), 'on_order'
 * (short, but everything short is on order) or 'blocked' (short and not ordered).
 */
export function buildOpenWorkQueue(db, { site = "main", today } = {}) {
  const siteCode = String(site || "main").trim().toLowerCase() || "main";
  const asOf = /^\d{4}-\d{2}-\d{2}$/.test(String(today || "")) ? String(today) : new Date().toISOString().slice(0, 10);
  if (!hasTable(db, "work_orders")) return { ok: true, as_of: asOf, site: siteCode, summary: summarize([]), work_orders: [] };

  const woCol = (c, fallback = "NULL") => (hasColumn(db, "work_orders", c) ? `w.${c}` : fallback);
  const hasBreakdowns = hasTable(db, "breakdowns");
  const hasPlans = hasTable(db, "maintenance_plans");
  // Breakdown jobs count their age from when the machine went down.
  const bdStart = hasBreakdowns && hasColumn(db, "breakdowns", "start_at") ? "NULLIF(TRIM(b.start_at), ''), " : "";
  const openedExpr = hasBreakdowns
    ? `CASE WHEN w.source = 'breakdown' THEN COALESCE(${bdStart}NULLIF(TRIM(b.breakdown_date), ''), w.opened_at) ELSE w.opened_at END`
    : "w.opened_at";
  const rows = db.prepare(`
    SELECT w.id, w.source, w.status, ${openedExpr} AS opened_at,
           ${woCol("assigned_artisan_name")} AS assigned_artisan_name, ${woCol("artisan_name")} AS artisan_name,
           ${woCol("priority")} AS priority, ${woCol("due_date")} AS due_date,
           ${woCol("job_description")} AS job_description, ${woCol("repair_progress")} AS repair_progress,
           ${hasBreakdowns ? "b.description" : "NULL"} AS breakdown_description,
           ${hasBreakdowns ? "b.component" : "NULL"} AS breakdown_component,
           ${hasBreakdowns ? "b.critical" : "NULL"} AS breakdown_critical,
           ${hasBreakdowns ? "b.parts_status" : "NULL"} AS breakdown_parts_status,
           ${hasPlans ? "mp.service_name" : "NULL"} AS service_name,
           w.asset_id, a.asset_code, a.asset_name, a.category
    FROM work_orders w
    JOIN assets a ON a.id = w.asset_id
    ${hasBreakdowns ? "LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'" : ""}
    ${hasPlans ? "LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id AND w.source = 'service'" : ""}
    WHERE LOWER(TRIM(COALESCE(${woCol("site_code", "'main'")}, 'main'))) = ?
      AND REPLACE(TRIM(LOWER(COALESCE(w.status, 'open'))), ' ', '_') NOT IN (${list(FINISHED_WO)})
      ${hasColumn(db, "work_orders", "merged_into_id") ? "AND w.merged_into_id IS NULL" : ""}
    ORDER BY w.id ASC
  `).all(siteCode);

  const downAssets = new Set(
    hasBreakdowns
      ? db.prepare(`SELECT DISTINCT asset_id FROM breakdowns WHERE UPPER(TRIM(COALESCE(status, 'OPEN'))) <> 'CLOSED'`).all().map((r) => Number(r.asset_id))
      : [],
  );

  const partsByWo = new Map();
  if (hasTable(db, "maintenance_parts_requests") && rows.length) {
    const reqs = db.prepare(`
      SELECT id, work_order_id, part_code, part_name, qty, urgency, status, created_at
      FROM maintenance_parts_requests
      WHERE work_order_id IS NOT NULL
        AND LOWER(COALESCE(status, 'requested')) IN (${list(WAITING_REQUEST_STATUSES)})
      ORDER BY id ASC
    `).all();
    for (const r of reqs) {
      const key = Number(r.work_order_id);
      if (!partsByWo.has(key)) partsByWo.set(key, []);
      partsByWo.get(key).push(r);
    }
  }

  const stock = stockLookup(db, siteCode);
  const workOrders = rows.map((w) => {
    const parts = (partsByWo.get(Number(w.id)) || []).map((r) => {
      const s = stock(r.part_code);
      const qty = Number(r.qty || 0);
      const status = String(r.status || "requested").toLowerCase();
      const inStock = s?.on_hand != null && qty > 0 && s.on_hand >= qty;
      const onOrder = status === "ordered" || Number(s?.on_order_qty || 0) > 0;
      return {
        request_id: Number(r.id),
        part_code: r.part_code || "",
        part_name: r.part_name || "",
        qty,
        urgency: String(r.urgency || "normal").toLowerCase(),
        request_status: status,
        requested_on: String(r.created_at || "").slice(0, 10) || null,
        on_hand: s ? s.on_hand : null,
        min_stock: s ? s.min_stock : null,
        below_min: Boolean(s && s.on_hand != null && s.min_stock > 0 && s.on_hand < s.min_stock),
        in_stock: inStock,
        on_order: onOrder,
        on_order_qty: s ? s.on_order_qty : 0,
        expected_arrival: s ? s.expected_arrival : null,
        not_in_catalogue: !s || !s.known,
      };
    });
    const short = parts.filter((p) => !p.in_stock);
    const partsState = !parts.length ? "none" : !short.length ? "ready" : short.every((p) => p.on_order) ? "on_order" : "blocked";
    const job = firstLine(
      [w.breakdown_component, w.breakdown_description].map((x) => String(x || "").trim()).filter(Boolean).join(" — ")
      || w.service_name || w.job_description || w.source,
    );
    return {
      id: Number(w.id),
      asset_id: Number(w.asset_id),
      asset_code: w.asset_code || "",
      asset_name: w.asset_name || "",
      category: w.category || "",
      source: w.source || "",
      job,
      status: String(w.status || "open").toLowerCase(),
      progress: firstLine(w.repair_progress, 120) || null,
      priority: w.priority ? String(w.priority).toLowerCase() : null,
      critical: Boolean(Number(w.breakdown_critical || 0)),
      machine_down: downAssets.has(Number(w.asset_id)),
      opened_on: String(w.opened_at || "").slice(0, 10) || null,
      age_days: daysBetween(w.opened_at, asOf),
      due_date: w.due_date || null,
      technician: String(w.assigned_artisan_name || w.artisan_name || "").trim() || null,
      breakdown_parts_status: w.breakdown_parts_status || null,
      parts,
      parts_state: partsState,
    };
  });

  // Down machines first, then critical, then parts blockers, then the oldest.
  const stateRank = { blocked: 0, on_order: 1, ready: 2, none: 2 };
  workOrders.sort((x, y) =>
    Number(y.machine_down) - Number(x.machine_down)
    || Number(y.critical) - Number(x.critical)
    || stateRank[x.parts_state] - stateRank[y.parts_state]
    || (y.age_days ?? 0) - (x.age_days ?? 0)
    || x.id - y.id);

  return { ok: true, as_of: asOf, site: siteCode, summary: summarize(workOrders), work_orders: workOrders };
}

function summarize(workOrders) {
  const count = (fn) => workOrders.filter(fn).length;
  return {
    open: workOrders.length,
    machines_down: count((w) => w.machine_down),
    unassigned: count((w) => !w.technician),
    parts_blocked: count((w) => w.parts_state === "blocked"),
    parts_on_order: count((w) => w.parts_state === "on_order"),
    parts_ready: count((w) => w.parts_state === "ready"),
    older_than_14_days: count((w) => (w.age_days ?? 0) > 14),
  };
}

/** Questions that should use the open work-order queue (asked from the Borris screen). */
export function isWorkQueueQuestion(question) {
  return /\b(open|outstanding|current)\s+(work[\s-]?orders?|jobs?|wos?|queue)\b|\bwork[\s-]?order\s+queue\b|\b(prioriti[sz]e|priority assessment)\b.*\b(jobs?|work[\s-]?orders?|queue)\b|\bparts?\s+blockers?\b/i.test(String(question || ""));
}

/** A smaller copy of the queue for Borris' context: the first `limit` jobs, short fields only. */
export function compactQueueForBorris(queue, limit = 60) {
  const rows = Array.isArray(queue?.work_orders) ? queue.work_orders : [];
  return {
    as_of: queue?.as_of || null,
    summary: queue?.summary || null,
    shown: Math.min(rows.length, limit),
    work_orders: rows.slice(0, limit).map((w) => ({
      wo: w.id,
      asset: w.asset_code,
      job: w.job,
      status: w.status,
      age_days: w.age_days,
      technician: w.technician,
      machine_down: w.machine_down,
      critical: w.critical,
      parts_state: w.parts_state,
      parts: w.parts.map((p) => ({
        part: p.part_code || p.part_name,
        qty: p.qty,
        on_hand: p.on_hand,
        below_min: p.below_min,
        on_order: p.on_order,
        eta: p.expected_arrival,
      })),
    })),
  };
}

const PARTS_STATE_TEXT = {
  none: "no parts waiting",
  ready: "parts in stock",
  on_order: "waiting on parts already ordered",
  blocked: "parts short and not ordered",
};

/** Plain-language priority list built from the data alone (used when no AI is connected). */
export function describeOpenWorkQueue(queue, limit = 15) {
  const s = queue?.summary || summarize([]);
  const rows = Array.isArray(queue?.work_orders) ? queue.work_orders : [];
  if (!rows.length) return "There are no open work orders on this site right now.";
  const lines = [
    `Open work orders: ${s.open} (${s.machines_down} on machines that are down, ${s.unassigned} without a technician).`,
    `Parts: ${s.parts_blocked} blocked and not ordered, ${s.parts_on_order} waiting on orders, ${s.parts_ready} ready to issue.`,
    "",
    "Suggested order of work (machines down, then critical, then parts blockers, then oldest):",
  ];
  rows.slice(0, limit).forEach((w, i) => {
    const flags = [w.machine_down ? "machine down" : null, w.critical ? "critical" : null].filter(Boolean).join(", ");
    const shortParts = w.parts.filter((p) => !p.in_stock).map((p) => p.part_code || p.part_name).slice(0, 3);
    lines.push(
      `${i + 1}. WO #${w.id} ${w.asset_code} — ${w.job}${flags ? ` (${flags})` : ""}; ${w.age_days ?? "?"} days open; `
      + `${w.technician ? `with ${w.technician}` : "no technician assigned"}; ${PARTS_STATE_TEXT[w.parts_state]}`
      + `${shortParts.length ? `: ${shortParts.join(", ")}` : ""}.`,
    );
  });
  if (rows.length > limit) lines.push(`…and ${rows.length - limit} more in the list below.`);
  lines.push("", "This is a read-only view. Make any changes in Work Orders or Stores.");
  return lines.join("\n");
}
