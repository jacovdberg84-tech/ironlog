// IRONLOG/api/utils/siteGuard.js — people limited to some sites can't open other sites' records.
//
// The sign-in check (auth/hook.js) pins a person with Allowed Locations to one of
// their sites and notes those sites on the request. This guard runs before every
// route (server.js, preHandler): when the address names a record by its number or
// machine code, or the body names a machine, it looks up that record's site and
// answers 404 when the person may not use it. One table of addresses covers every
// route that opens, changes or prints a single record, including routes whose
// queries are prepared elsewhere. Administrators and people without Allowed
// Locations are not limited (docs/sites.md).
import { defaultSite } from "./sites.js";

// [address pattern (after /api), table, column the captured value matches]
export const RECORD_RULES = [
  [/^\/breakdowns\/(\d+)(?:\/|$)/, "breakdowns", "id"],
  [/^\/maintenance\/breakdowns\/(\d+)\//, "breakdowns", "id"],
  [/^\/workorders\/(\d+)(?:\/|$)/, "work_orders", "id"],
  [/^\/tech\/workorders\/(\d+)(?:\/|$)/, "work_orders", "id"],
  [/^\/reports\/workorder\/(\d+)\.pdf$/, "work_orders", "id"],
  [/^\/tech\/assets\/([^/]+)$/, "assets", "asset_code"],
  [/^\/assets\/([^/]+)\/(?:history|hours|qr-profile|tyre-qr-profile|undercarriage-qr-profile)(?:\/refresh)?$/, "assets", "asset_code"],
  [/^\/assets\/([^/]+)(?:\/archive)?$/, "assets", "asset_code"],
  [/^\/reports\/asset-history\/([^/]+)\.pdf$/, "assets", "asset_code"],
  [/^\/maintenance\/asset\/(\d+)\/live-hours$/, "assets", "id"],
  [/^\/breakdown-ops\/slips\/(\d+)(?:\/|$)/, "ops_slip_reports", "id"],
  [/^\/breakdown-ops\/offsite-repairs\/(\d+)(?:\/|$)/, "breakdown_offsite_repairs", "id"],
  [/^\/dispatch\/trips\/(\d+)\//, "dispatch_trips", "id"],
  [/^\/maintenance\/inspections\/(\d+)(?:\/|$)/, "manager_inspections", "id"],
  [/^\/reports\/manager-inspection\/(\d+)\.pdf$/, "manager_inspections", "id"],
  [/^\/maintenance\/undercarriage-inspections\/(\d+)\.pdf$/, "undercarriage_inspections", "id"],
  [/^\/maintenance\/mechanic-labor\/(\d+)$/, "mechanic_labor_entries", "id"],
  [/^\/maintenance\/parts-requests\/(\d+)\//, "maintenance_parts_requests", "id"],
  [/^\/maintenance\/damage-reports\/(\d+)\//, "manager_damage_reports", "id"],
  [/^\/reports\/damage-report\/(\d+)\.pdf$/, "manager_damage_reports", "id"],
  [/^\/maintenance\/vehicle-ldv-checks\/(\d+)\//, "vehicle_ldv_checks", "id"],
  [/^\/reports\/vehicle-ldv-check\/(\d+)\.pdf$/, "vehicle_ldv_checks", "id"],
  [/^\/reports\/artisan-inspection\/(\d+)\.pdf$/, "artisan_inspections", "id"],
  [/^\/tech\/inspections\/(\d+)\/pdf$/, "artisan_inspections", "id"],
  [/^\/procurement\/purchase-orders\/(\d+)\//, "procurement_purchase_orders", "id"],
  [/^\/procurement\/requisitions\/(\d+)\//, "procurement_requisitions", "id"],
  [/^\/stock\/part-orders\/(\d+)(?:\/|$)/, "stores_part_orders", "id"],
  [/^\/tasks\/(\d+)(?:\/|$)/, "tasks", "id"],
  [/^\/projects\/(\d+)$/, "projects", "id"],
];

const hasSiteColumn = new Map();
function siteColumn(db, table) {
  if (!hasSiteColumn.has(table)) {
    let ok = false;
    try {
      ok = db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === "site_code");
    } catch {}
    hasSiteColumn.set(table, ok);
  }
  return hasSiteColumn.get(table);
}

const norm = (s) => String(s ?? "").trim().toLowerCase() || defaultSite();

/** Site of one record (null when there is no such record or the table has no site). */
export function recordSite(db, table, column, value) {
  if (!siteColumn(db, table)) return null;
  try {
    const row = db.prepare(`SELECT site_code FROM ${table} WHERE ${column} = ? COLLATE NOCASE LIMIT 1`).get(value);
    return row ? norm(row.site_code) : null;
  } catch {
    return null;
  }
}

/** Records a request names: by address, and machines, work orders and breakdowns named in the body. */
export function namedRecords(path, body) {
  const out = [];
  const rel = String(path || "").replace(/^\/api/, "");
  for (const [re, table, column] of RECORD_RULES) {
    const m = rel.match(re);
    if (m) {
      out.push({ table, column, value: decodeURIComponent(m[1]) });
      break;
    }
  }
  const b = body && typeof body === "object" && !Array.isArray(body) ? body : null;
  if (b && Number.isFinite(Number(b.asset_id)) && Number(b.asset_id) > 0) out.push({ table: "assets", column: "id", value: Number(b.asset_id) });
  if (b && typeof b.asset_code === "string" && b.asset_code.trim()) out.push({ table: "assets", column: "asset_code", value: b.asset_code.trim() });
  if (b && Number.isFinite(Number(b.work_order_id)) && Number(b.work_order_id) > 0) out.push({ table: "work_orders", column: "id", value: Number(b.work_order_id) });
  if (b && Number.isFinite(Number(b.breakdown_id)) && Number(b.breakdown_id) > 0) out.push({ table: "breakdowns", column: "id", value: Number(b.breakdown_id) });
  return out;
}

/** The first named record outside the allowed sites, or null. */
export function blockedRecord(db, allowedSites, path, body) {
  if (!Array.isArray(allowedSites) || !allowedSites.length) return null;
  for (const r of namedRecords(path, body)) {
    const site = recordSite(db, r.table, r.column, r.value);
    if (site && !allowedSites.includes(site)) return { ...r, site };
  }
  return null;
}

/** preHandler for server.js. */
export function makeSiteGuard(db) {
  return async function siteGuard(req, reply) {
    const allowed = req.ironlogAllowedSites;
    if (!allowed) return;
    const hit = blockedRecord(db, allowed, String(req.url || "").split("?")[0], req.body);
    if (hit) return reply.code(404).send({ ok: false, error: "Not found on your sites." });
  };
}
