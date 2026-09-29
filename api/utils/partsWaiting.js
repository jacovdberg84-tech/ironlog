// IRONLOG/api/utils/partsWaiting.js — what the workshop is waiting on from stores.
//
// Two sources feed the stores queue:
//  - parts requests raised on a work order or machine (maintenance_parts_requests)
//    that are still requested or ordered;
//  - open breakdowns marked as waiting on parts where nobody has listed the part
//    yet, so stores can chase the workshop for a part number.
// Machines that are down come first, then critical/urgent requests, oldest first.

const WAITING_REQUEST_STATUSES = ["requested", "ordered"];
const WAITING_BREAKDOWN_PARTS = ["not ordered", "ordered", "in transit", "partial", "waiting oem"];
const FINISHED_WO = ["completed", "approved", "closed"];

const list = (xs) => xs.map((x) => `'${x}'`).join(", ");

function hasTable(db, name) {
  return Boolean(db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?`).get(name));
}

function onHandByCode(db) {
  if (!hasTable(db, "parts")) return () => null;
  const stmt = hasTable(db, "stock_movements")
    ? db.prepare(`
        SELECT p.part_code, p.part_name, IFNULL(SUM(sm.quantity), 0) AS on_hand
        FROM parts p
        LEFT JOIN stock_movements sm ON sm.part_id = p.id
        WHERE LOWER(p.part_code) = LOWER(?)
        GROUP BY p.id
      `)
    : null;
  return (code) => {
    const c = String(code || "").trim();
    if (!c || !stmt) return null;
    const row = stmt.get(c);
    return row ? Number(row.on_hand) : null;
  };
}

/**
 * Rows the stores team should act on. Each row has `kind` ('request' or
 * 'breakdown'), the machine, the work order and whether the machine is down.
 */
export function workshopWaitingOnParts(db, { site = "main" } = {}) {
  const siteCode = String(site || "main").trim().toLowerCase() || "main";
  const rows = [];
  const onHand = onHandByCode(db);
  const hasBreakdowns = hasTable(db, "breakdowns");
  const downAssets = new Set(
    hasBreakdowns
      ? db.prepare(`SELECT DISTINCT asset_id FROM breakdowns WHERE UPPER(TRIM(COALESCE(status, 'OPEN'))) <> 'CLOSED'`).all().map((r) => Number(r.asset_id))
      : [],
  );

  const coveredWorkOrders = new Set();
  if (hasTable(db, "maintenance_parts_requests")) {
    const reqs = db.prepare(`
      SELECT pr.id, pr.asset_id, COALESCE(a.asset_code, pr.asset_code) AS asset_code, a.asset_name,
             pr.part_code, pr.part_name, pr.qty, pr.urgency, pr.status, pr.work_order_id,
             pr.requested_by, pr.notes, pr.status_notes, pr.created_at,
             w.status AS work_order_status, COALESCE(w.asset_id, pr.asset_id) AS machine_id
      FROM maintenance_parts_requests pr
      LEFT JOIN work_orders w ON w.id = pr.work_order_id
      LEFT JOIN assets a ON a.id = COALESCE(pr.asset_id, w.asset_id)
      WHERE LOWER(COALESCE(pr.status, 'requested')) IN (${list(WAITING_REQUEST_STATUSES)})
        AND COALESCE(pr.site_code, 'main') = ?
        AND (pr.work_order_id IS NULL OR w.id IS NULL
             OR REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') NOT IN (${list(FINISHED_WO)}))
    `).all(siteCode);
    for (const r of reqs) {
      if (r.work_order_id) coveredWorkOrders.add(Number(r.work_order_id));
      const stock = onHand(r.part_code);
      rows.push({
        kind: "request",
        id: r.id,
        asset_id: r.machine_id != null ? Number(r.machine_id) : null,
        asset_code: r.asset_code || "",
        asset_name: r.asset_name || "",
        work_order_id: r.work_order_id != null ? Number(r.work_order_id) : null,
        part_code: r.part_code || "",
        part_name: r.part_name || "",
        qty: Number(r.qty || 0),
        urgency: String(r.urgency || "normal").toLowerCase(),
        status: String(r.status || "requested").toLowerCase(),
        requested_by: r.requested_by || "",
        notes: [r.notes, r.status_notes].map((x) => String(x || "").trim()).filter(Boolean).join(" · "),
        since: r.created_at,
        on_hand: stock,
        in_stock: stock != null && stock >= Number(r.qty || 0) && Number(r.qty || 0) > 0,
        machine_down: r.machine_id != null && downAssets.has(Number(r.machine_id)),
      });
    }
  }

  if (hasBreakdowns) {
    const bds = db.prepare(`
      SELECT b.id, b.asset_id, a.asset_code, a.asset_name, b.component, b.description, b.parts_status,
             b.critical, b.breakdown_date, b.ets_repair_date,
             COALESCE(b.primary_work_order_id, (
               SELECT w.id FROM work_orders w WHERE w.source = 'breakdown' AND w.reference_id = b.id ORDER BY w.id DESC LIMIT 1
             )) AS work_order_id
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      WHERE UPPER(TRIM(COALESCE(b.status, 'OPEN'))) <> 'CLOSED'
        AND LOWER(TRIM(COALESCE(b.parts_status, ''))) IN (${list(WAITING_BREAKDOWN_PARTS)})
        AND LOWER(TRIM(COALESCE(b.site_code, 'main'))) = ?
    `).all(siteCode);
    for (const b of bds) {
      if (b.work_order_id && coveredWorkOrders.has(Number(b.work_order_id))) continue;
      rows.push({
        kind: "breakdown",
        id: b.id,
        asset_id: Number(b.asset_id),
        asset_code: b.asset_code || "",
        asset_name: b.asset_name || "",
        work_order_id: b.work_order_id != null ? Number(b.work_order_id) : null,
        part_code: "",
        part_name: "",
        qty: 0,
        urgency: b.critical ? "critical" : "normal",
        status: String(b.parts_status || ""),
        requested_by: "",
        notes: [b.component, b.description].map((x) => String(x || "").trim()).filter(Boolean).join(" — "),
        since: b.breakdown_date,
        return_date: b.ets_repair_date || null,
        on_hand: null,
        in_stock: false,
        machine_down: true,
      });
    }
  }

  const urgencyRank = { critical: 0, urgent: 1 };
  rows.sort((x, y) =>
    Number(y.machine_down) - Number(x.machine_down)
    || (urgencyRank[x.urgency] ?? 2) - (urgencyRank[y.urgency] ?? 2)
    || String(x.since || "").localeCompare(String(y.since || ""))
    || x.id - y.id);
  return rows;
}

/** Counts for headings: total rows, rows on down machines, and requests stores can fill now. */
export function summarizeWaiting(rows) {
  return {
    total: rows.length,
    machines_down: new Set(rows.filter((r) => r.machine_down).map((r) => r.asset_id)).size,
    in_stock: rows.filter((r) => r.in_stock && r.status === "requested").length,
    not_ordered: rows.filter((r) => r.kind === "request" && r.status === "requested" && !r.in_stock).length,
  };
}
