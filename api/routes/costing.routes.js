// IRONLOG/api/routes/costing.routes.js
// Costing feeds for Excel Power Query (Get Data from Web).

import { db } from "../db/client.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function getRoles(req) {
  const many = String(req.headers["x-user-roles"] || "")
    .split(",")
    .map((x) => String(x || "").trim().toLowerCase())
    .filter(Boolean);
  const one = String(req.headers["x-user-role"] || "")
    .split(",")
    .map((x) => String(x || "").trim().toLowerCase())
    .filter(Boolean);
  return Array.from(new Set([...many, ...one]));
}

function requireRoles(req, reply, allowed) {
  const roles = getRoles(req);
  if (!roles.some((r) => allowed.includes(r))) {
    reply.code(403).send({ ok: false, error: "not allowed" });
    return false;
  }
  return true;
}

function hasTable(table) {
  try {
    const r = db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name=? LIMIT 1`).get(String(table));
    return Boolean(r?.name);
  } catch {
    return false;
  }
}

function tableHasColumn(table, col) {
  try {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name) === col);
  } catch {
    return false;
  }
}

function ensureColumn(table, colName, colDef) {
  if (!hasTable(table)) return;
  if (!tableHasColumn(table, colName)) {
    try {
      db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
    } catch {}
  }
}

function fmtNum(v, digits = 2) {
  const n = Number(v);
  if (!Number.isFinite(n)) return null;
  return Number(n.toFixed(digits));
}

function readLaborRateDefault() {
  try {
    const row = db.prepare(`SELECT value FROM cost_settings WHERE key = 'labor_cost_per_hour_default' LIMIT 1`).get();
    const v = Number(row?.value);
    return Number.isFinite(v) && v > 0 ? v : 35;
  } catch {
    return 35;
  }
}

const MECHANICS_READ_ROLES = [
  "admin",
  "executive",
  "finance",
  "supervisor",
  "plant_manager",
  "site_manager",
  "workshop_manager",
  "stores",
  "storeman",
];

export default async function costingRoutes(app) {
  ensureColumn("mechanic_labor_entries", "category", "category TEXT");
  ensureColumn("mechanic_labor_entries", "time_started", "time_started TEXT");
  ensureColumn("mechanic_labor_entries", "time_finished", "time_finished TEXT");
  ensureColumn("mechanic_labor_entries", "job_card_no", "job_card_no TEXT");
  ensureColumn("mechanic_labor_entries", "smr", "smr REAL");

  // GET /api/costing/mechanics?from=2026-01-01&to=2026-12-31
  // Optional: site_code=main, format=table (bare JSON array for Excel Get Data)
  app.get("/mechanics", async (req, reply) => {
    if (!requireRoles(req, reply, MECHANICS_READ_ROLES)) return;

    const from = String(req.query?.from || "").trim();
    const to = String(req.query?.to || "").trim();
    const siteCode = String(req.query?.site_code || req.headers["x-site-code"] || "main")
      .trim()
      .toLowerCase() || "main";
    const format = String(req.query?.format || "").trim().toLowerCase();

    if (!isDate(from) || !isDate(to)) {
      return reply.code(400).send({ ok: false, error: "from and to (YYYY-MM-DD) are required" });
    }
    if (from > to) {
      return reply.code(400).send({ ok: false, error: "from must be on or before to" });
    }

    if (!hasTable("mechanic_labor_entries")) {
      const empty = [];
      if (format === "table" || format === "array") return reply.send(empty);
      return reply.send({ ok: true, from, to, site_code: siteCode, count: 0, rows: empty });
    }

    const hasDailyHours = hasTable("daily_hours");
    const smrExpr = hasDailyHours
      ? `COALESCE(
          m.smr,
          (
            SELECT dh.closing_hours
            FROM daily_hours dh
            WHERE dh.asset_id = a.id
              AND dh.work_date <= m.work_date
              AND dh.closing_hours IS NOT NULL
            ORDER BY dh.work_date DESC
            LIMIT 1
          )
        )`
      : "m.smr";

    const rows = db.prepare(`
      SELECT
        m.id,
        m.work_date,
        m.asset_code,
        m.hours,
        COALESCE(NULLIF(TRIM(m.category), ''), NULLIF(TRIM(a.category), ''), '') AS category,
        m.reason,
        m.time_started,
        m.time_finished,
        m.technician_name,
        m.job_card_no,
        ${smrExpr} AS smr,
        m.labor_rate_per_hour
      FROM mechanic_labor_entries m
      LEFT JOIN assets a
        ON UPPER(TRIM(a.asset_code)) = UPPER(TRIM(m.asset_code))
      WHERE m.work_date >= ?
        AND m.work_date <= ?
        AND LOWER(TRIM(COALESCE(m.site_code, 'main'))) = ?
      ORDER BY m.work_date ASC, m.technician_name ASC, m.id ASC
    `).all(from, to, siteCode);

    const defaultRate = readLaborRateDefault();
    const mapped = rows.map((r) => {
      const hours = fmtNum(r.hours, 2);
      const smr = r.smr != null && Number.isFinite(Number(r.smr)) ? fmtNum(r.smr, 1) : null;
      const rate = Number.isFinite(Number(r.labor_rate_per_hour)) && Number(r.labor_rate_per_hour) > 0
        ? Number(r.labor_rate_per_hour)
        : defaultRate;
      return {
        Date: String(r.work_date || ""),
        "Plant no": String(r.asset_code || ""),
        "Work Hours": hours ?? 0,
        Category: String(r.category || ""),
        "Description Of Work Carried Out": String(r.reason || ""),
        "Time Started": String(r.time_started || ""),
        "Time finished": String(r.time_finished || ""),
        Technician: String(r.technician_name || ""),
        "Job Card No": String(r.job_card_no || "").trim(),
        SMR: smr,
        labor_rate_per_hour: fmtNum(rate, 2),
        labor_cost: hours != null ? fmtNum(hours * rate, 2) : null,
        entry_id: Number(r.id || 0),
      };
    });

    if (format === "table" || format === "array") {
      return reply.send(mapped);
    }

    return reply.send({
      ok: true,
      from,
      to,
      site_code: siteCode,
      count: mapped.length,
      rows: mapped,
    });
  });
}
