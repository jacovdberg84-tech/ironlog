// IRONLOG/api/utils/sites.js — a company's sites, and the one-time tidy-up of site codes.
//
// Every record belongs to a site (site_code, "main" when a company has only
// one). People limited to some sites (users.allowed_locations) only see those;
// administrators and unlimited users choose any site. A site code is a short
// plain word ("main", "north-pit"); anything else — a screen name typed into
// Allowed Locations, or a list saved as text ('["workshop"]') — is a mistake.
import fs from "node:fs";
import path from "node:path";

export const SITE_CODE = /^[a-z0-9][a-z0-9_-]{0,31}$/;

export function defaultSite() {
  const s = String(process.env.IRONLOG_DEFAULT_SITE_CODE || "main").trim().toLowerCase();
  return SITE_CODE.test(s) ? s : "main";
}

/** '["workshop"]', '["[\\"workshop\\"]"]', ' Main ' → ["workshop"] / ["main"]: lists saved as text are opened up. */
export function unwrapSiteList(raw) {
  let items = Array.isArray(raw) ? raw : [raw];
  for (let i = 0; i < 4; i++) {
    const next = [];
    let opened = false;
    for (const it of items) {
      const t = String(it ?? "").trim();
      if (t.startsWith("[")) {
        try {
          const v = JSON.parse(t);
          if (Array.isArray(v)) { next.push(...v); opened = true; continue; }
        } catch {}
      }
      next.push(...t.split(","));
    }
    items = next;
    if (!opened) break;
  }
  return [...new Set(items.map((x) => String(x ?? "").trim().toLowerCase().replace(/^["']+|["']+$/g, "")).filter(Boolean))];
}

function tableHas(db, table, col) {
  try {
    return db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
  } catch {
    return false;
  }
}

/** The company's sites: site profiles, the default site, and clean site codes on its machines. */
export function knownSites(db) {
  const out = new Set([defaultSite()]);
  try {
    for (const r of db.prepare(`SELECT site_code FROM site_profiles WHERE COALESCE(active, 1) = 1`).all()) {
      const s = String(r.site_code || "").trim().toLowerCase();
      if (SITE_CODE.test(s)) out.add(s);
    }
  } catch {}
  if (tableHas(db, "assets", "site_code")) {
    for (const r of db.prepare(`SELECT DISTINCT LOWER(TRIM(site_code)) AS s FROM assets WHERE site_code IS NOT NULL`).all()) {
      if (SITE_CODE.test(r.s || "")) out.add(r.s);
    }
  }
  return [...out].sort();
}

/** Allowed locations an admin typed: { ok, sites|null } or { ok:false, error } naming the real sites. */
export function checkAllowedLocations(db, raw) {
  if (raw == null || String(raw).trim() === "" || (Array.isArray(raw) && !raw.length)) return { ok: true, sites: null };
  const wanted = unwrapSiteList(raw);
  const known = knownSites(db);
  const unknown = wanted.filter((s) => !known.includes(s));
  if (unknown.length) {
    return { ok: false, error: `Allowed locations must be sites. Not a site: ${unknown.join(", ")}. Sites: ${known.join(", ")}. Leave empty for all sites. (Screens are set under Sections.)` };
  }
  return { ok: true, sites: wanted.length ? wanted : null };
}

// Tables whose site_code was written from a garbled user setting (seen at AML, October 2026).
const GARBLED_TABLES = ["work_orders", "tech_shifts", "tech_other_time", "breakdowns", "daily_hours", "fuel_logs", "oil_logs",
  "artisan_inspections", "manager_inspections", "mechanic_labor_entries", "maintenance_parts_requests", "stores_part_orders",
  "stock_deliveries", "tasks", "assets", "vehicle_ldv_checks"];

/** What the tidy-up would change (nothing is written). */
export function planSiteRepair(db, known = knownSites(db)) {
  const def = defaultSite();
  const changes = [];
  for (const t of GARBLED_TABLES) {
    if (!tableHas(db, t, "site_code")) continue;
    for (const r of db.prepare(`SELECT id, site_code FROM ${t} WHERE TRIM(site_code) LIKE '[%'`).all()) {
      const [first] = unwrapSiteList(r.site_code);
      // Kept only when it names a real site; otherwise the default site (AML: "workshop" → "main").
      changes.push({ table: t, id: r.id, from: r.site_code, to: known.includes(first) ? first : def, why: "site saved as a list" });
    }
  }
  if (tableHas(db, "assets", "site_code")) {
    for (const r of db.prepare(`SELECT id, site_code FROM assets WHERE site_code IS NULL OR TRIM(site_code) = ''`).all()) {
      changes.push({ table: "assets", id: r.id, from: r.site_code ?? null, to: def, why: "machine had no site" });
    }
  }
  return changes;
}

/** Users whose allowed locations name things that are not sites, after the record fixes. */
export function planUserRepair(db, known) {
  const out = [];
  if (!tableHas(db, "users", "allowed_locations")) return out;
  for (const u of db.prepare(`SELECT id, username, allowed_locations FROM users WHERE allowed_locations IS NOT NULL AND TRIM(allowed_locations) NOT IN ('', '[]', 'null')`).all()) {
    const list = unwrapSiteList(u.allowed_locations);
    const keep = list.filter((s) => known.includes(s));
    const next = keep.length ? JSON.stringify(keep) : null;
    if (next !== u.allowed_locations) out.push({ id: u.id, username: u.username, from: u.allowed_locations, to: next });
  }
  return out;
}

/**
 * Once per database: copy the database, then put garbled site codes right, give machines with
 * no site the default site, and drop allowed locations that are not sites. Every change is audited.
 */
export async function repairSiteDataOnce(db, { backupDir, audit = () => {}, log = console } = {}) {
  db.exec(`CREATE TABLE IF NOT EXISTS site_repairs (step TEXT PRIMARY KEY, done_at TEXT NOT NULL DEFAULT (datetime('now')), summary TEXT)`);
  if (db.prepare(`SELECT 1 FROM site_repairs WHERE step = 'site_codes_2026_10'`).get()) return null;
  const known = knownSites(db);
  const records = planSiteRepair(db, known);
  const users = planUserRepair(db, known);
  let backup = null;
  if ((records.length || users.length) && backupDir) {
    fs.mkdirSync(backupDir, { recursive: true });
    backup = path.join(backupDir, `ironlog_before_site_repair_${new Date().toISOString().replace(/\D/g, "").slice(0, 14)}.db`);
    await db.backup(backup);
  }
  db.transaction(() => {
    for (const c of records) db.prepare(`UPDATE ${c.table} SET site_code = ? WHERE id = ?`).run(c.to, c.id);
    for (const u of users) db.prepare(`UPDATE users SET allowed_locations = ? WHERE id = ?`).run(u.to, u.id);
    const summary = { records: records.length, users: users.map((u) => u.username), backup };
    db.prepare(`INSERT INTO site_repairs (step, summary) VALUES ('site_codes_2026_10', ?)`).run(JSON.stringify(summary));
  })();
  for (const c of records) audit({ action: "site_repair.record", entity_type: c.table, entity_id: String(c.id), before: { site_code: c.from }, after: { site_code: c.to }, payload: { why: c.why } });
  for (const u of users) audit({ action: "site_repair.user", entity_type: "user", entity_id: u.username, before: { allowed_locations: u.from }, after: { allowed_locations: u.to }, payload: { why: "allowed locations were not sites" } });
  if (records.length || users.length) log.info?.(`[sites] tidied ${records.length} site codes and ${users.length} users' allowed locations; copy first: ${backup}`);
  return { records, users, backup };
}

// ---- Site filters for routes (docs/sites.md). Blank site codes count as the default site.

/** SQL for a column's site, lower case, blank = default site: `${siteSql("a.site_code")} = ?`. */
export function siteSql(column) {
  return `LOWER(COALESCE(NULLIF(TRIM(${column}), ''), '${defaultSite()}'))`;
}

/** The site this request is for (set by the sign-in check; default site when none). */
export function requestSite(req) {
  const s = String(req?.headers?.["x-site-code"] || "").trim().toLowerCase();
  return s || defaultSite();
}

/** True when a record's site is the request's site. */
export function onRequestSite(req, recordSite) {
  const s = String(recordSite ?? "").trim().toLowerCase() || defaultSite();
  return s === requestSite(req);
}

/** Refuses (404) a record from another site; returns false when it did. */
export function requireRequestSite(req, reply, recordSite, what = "Record") {
  if (onRequestSite(req, recordSite)) return true;
  reply.code(404).send({ ok: false, error: `${what} not found on this site.` });
  return false;
}

// Records that belong to a machine take the machine's site when they are created, so a
// breakdown, work order, fuel log or inspection always sits on its machine's site.
export const MACHINE_RECORD_TABLES = ["daily_hours", "fuel_logs", "oil_logs", "work_orders", "breakdowns", "ops_slip_reports",
  "breakdown_offsite_repairs", "stores_part_orders", "manager_inspections", "artisan_inspections", "manager_damage_reports",
  "maintenance_parts_requests", "tyre_inspections", "undercarriage_inspections", "undercarriage_wear_profiles", "vehicle_ldv_checks"];

export function ensureSiteFollowsMachine(db) {
  const made = [];
  for (const t of MACHINE_RECORD_TABLES) {
    if (!tableHas(db, t, "site_code") || !tableHas(db, t, "asset_id")) continue;
    db.exec(`
      CREATE TRIGGER IF NOT EXISTS site_follows_machine_${t}
      AFTER INSERT ON ${t}
      FOR EACH ROW WHEN NEW.asset_id IS NOT NULL
      BEGIN
        UPDATE ${t} SET site_code = (SELECT LOWER(TRIM(site_code)) FROM assets WHERE id = NEW.asset_id)
        WHERE rowid = NEW.rowid
          AND (SELECT NULLIF(TRIM(site_code), '') FROM assets WHERE id = NEW.asset_id) IS NOT NULL;
      END;
    `);
    made.push(t);
  }
  return made;
}

/** " AND <alias>.site_code is the request's site" — safe to place in SQL (the code is checked first). */
export function andSite(alias, req) {
  const s = requestSite(req);
  if (!SITE_CODE.test(s)) return " AND 0";
  return ` AND ${siteSql(alias ? `${alias}.site_code` : "site_code")} = '${s}'`;
}

/** " AND <column> is a machine on the request's site" — for tables joined to machines only by asset_id. */
export function andAssetSite(column, req) {
  const s = requestSite(req);
  if (!SITE_CODE.test(s)) return " AND 0";
  return ` AND ${column} IN (SELECT id FROM assets WHERE ${siteSql("site_code")} = '${s}')`;
}

/** Machines carry a site (the live database has it; this makes sure on a new or test database). */
export function ensureAssetSiteColumn(db) {
  if (!tableHas(db, "assets", "site_code") && tableHas(db, "assets", "id")) {
    try { db.prepare(`ALTER TABLE assets ADD COLUMN site_code TEXT`).run(); } catch {}
  }
}

let knownCache = { at: 0, sites: null };
/** knownSites, re-read at most once a minute (the sign-in check asks on every request). */
export function knownSitesCached(db, now = Date.now()) {
  if (!knownCache.sites || now - knownCache.at > 60000) knownCache = { at: now, sites: knownSites(db) };
  return knownCache.sites;
}
