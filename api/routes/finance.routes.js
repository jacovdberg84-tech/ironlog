// IRONLOG/api/routes/finance.routes.js
// Finance integration: month-end lock + close checklist, budget/forecast,
// canonical KPI definitions, and SSOT finance reporting.

import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";

function getRole(req) {
  return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
}
function getUser(req) {
  return String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
}
function getRoles(req) {
  const many = String(req.headers["x-user-roles"] || "")
    .split(",").map((x) => String(x || "").trim().toLowerCase()).filter(Boolean);
  const one = String(req.headers["x-user-role"] || "")
    .split(",").map((x) => String(x || "").trim().toLowerCase()).filter(Boolean);
  const merged = Array.from(new Set([...many, ...one]));
  return merged.length ? merged : ["operator"];
}
function hasAnyRole(req, allowed) {
  const roles = getRoles(req);
  return roles.some((r) => allowed.includes(r));
}
function requireRoles(req, reply, allowed) {
  if (!hasAnyRole(req, allowed)) {
    reply.code(403).send({ error: `role '${getRole(req)}' not allowed` });
    return false;
  }
  return true;
}
function tableHasColumn(table, col) {
  try {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name) === col);
  } catch { return false; }
}
function readCostSetting(key, fallback) {
  try {
    const row = db.prepare(`SELECT value FROM cost_settings WHERE key = ? LIMIT 1`).get(key);
    const v = Number(row?.value);
    return Number.isFinite(v) ? v : fallback;
  } catch { return fallback; }
}
function monthStart(period) {
  const [y, m] = String(period).split("-");
  return `${y}-${String(m).padStart(2, "0")}-01`;
}
function monthEnd(period) {
  const [y, m] = String(period).split("-").map((x) => Number(x));
  const end = new Date(Date.UTC(y, m, 0));
  return end.toISOString().slice(0, 10);
}
function addMonths(period, delta) {
  const [y, m] = String(period).split("-").map((x) => Number(x));
  const d = new Date(Date.UTC(y, m - 1 + delta, 1));
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
}
function ensureLockRow(period) {
  db.prepare(`
    INSERT OR IGNORE INTO finance_period_locks (period, status) VALUES (?, 'open')
  `).run(String(period));
  return db.prepare(`SELECT * FROM finance_period_locks WHERE period = ?`).get(String(period));
}
function isPeriodLocked(period) {
  const row = db.prepare(`SELECT status FROM finance_period_locks WHERE period = ?`).get(String(period || ""));
  const s = String(row?.status || "open").toLowerCase();
  return s === "locked" || s === "closed";
}

/* ============================================================
   DEFAULT MONTH-END CHECKLIST TEMPLATE
============================================================ */
const DEFAULT_CHECKLIST = [
  { code: "meters_up_to_date", label: "All meter readings captured to month end" },
  { code: "daily_hours_locked", label: "Daily hours finalized for every working day" },
  { code: "breakdowns_closed", label: "Open breakdowns from prior months reviewed/closed" },
  { code: "wo_labor_complete", label: "Work orders completed/closed for the month" },
  { code: "parts_issued_posted", label: "All parts issues posted (no pending stock_movements)" },
  { code: "procurement_grn_matched", label: "Goods receipts matched or exceptions queued" },
  { code: "invoices_captured", label: "Supplier invoices captured for closed POs" },
  { code: "fuel_lube_posted", label: "Fuel + lube logs posted for the month" },
  { code: "summarized_journal_run", label: "Summarized journal posting run built and exported" },
  { code: "budget_variance_reviewed", label: "Budget vs actual variance reviewed" },
  { code: "period_locked", label: "Period locked in IRONLOG" },
];

/* ============================================================
   KPI DEFINITIONS REGISTRY (single source of truth)
============================================================ */
const KPI_DEFINITIONS = [
  {
    code: "availability",
    label: "Availability",
    unit: "percent",
    grain: ["daily", "monthly", "asset", "site", "cost_center", "equipment_type"],
    formula: "(available_hours - downtime_hours) / available_hours",
    numerator: "available_hours - downtime_hours",
    denominator: "available_hours",
    source_tables: ["daily_hours", "breakdown_downtime_logs", "assets"],
    exclusions: ["standby assets", "assets archived in period"],
    notes: "available_hours = scheduled hours for used assets in window"
  },
  {
    code: "utilization",
    label: "Utilization",
    unit: "percent",
    grain: ["daily", "monthly", "asset", "site", "cost_center", "equipment_type"],
    formula: "run_hours / available_hours",
    numerator: "run_hours (SUM daily_hours.hours_run)",
    denominator: "available_hours",
    source_tables: ["daily_hours", "assets"],
    exclusions: ["standby assets", "days with hours_run = 0"],
    notes: "Available hours use scheduled value (default 10h per used asset per day)"
  },
  {
    code: "mtbf",
    label: "Mean Time Between Failures",
    unit: "hours",
    grain: ["monthly", "asset", "site", "equipment_type"],
    formula: "SUM(run_hours) / COUNT(breakdowns)",
    numerator: "SUM(run_hours) across active assets in window",
    denominator: "Count of breakdowns in window (status in reported/open/closed)",
    source_tables: ["daily_hours", "breakdowns"],
    exclusions: ["standby assets", "breakdowns with zero downtime"],
    notes: "Returned NULL when denominator = 0"
  },
  {
    code: "mttr",
    label: "Mean Time To Repair",
    unit: "hours",
    grain: ["monthly", "asset", "site", "equipment_type"],
    formula: "SUM(downtime_hours) / COUNT(breakdowns)",
    numerator: "SUM(breakdown_downtime_logs.hours_down) in window",
    denominator: "Count of distinct breakdowns contributing downtime in window",
    source_tables: ["breakdown_downtime_logs", "breakdowns"],
    exclusions: ["imputed downtime rows with hours_down = 0"],
    notes: "Returned NULL when denominator = 0"
  },
  {
    code: "cost_per_asset_hour",
    label: "Cost per Asset Hour",
    unit: "currency_per_hour",
    grain: ["monthly", "asset", "site", "cost_center", "equipment_type"],
    formula: "(parts + labor + fuel + lube + downtime) / run_hours",
    numerator: "SUM of parts + labor + fuel + lube + downtime costs in window",
    denominator: "SUM(daily_hours.hours_run) in window",
    source_tables: ["stock_movements", "parts", "work_orders", "fuel_logs", "oil_logs", "breakdown_downtime_logs", "assets", "cost_settings"],
    exclusions: ["assets with zero run hours in window"],
    notes: "Uses cost_settings defaults when per-asset rates missing"
  },
];

export default async function financeRoutes(app) {
  ensureAuditTable(db);

  /* --- ensure tables owned by this module exist --- */
  db.prepare(`
    CREATE TABLE IF NOT EXISTS finance_period_locks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      period TEXT NOT NULL UNIQUE,
      status TEXT NOT NULL DEFAULT 'open',
      locked_by TEXT,
      locked_at TEXT,
      reopened_by TEXT,
      reopened_at TEXT,
      reopen_reason TEXT,
      closed_by TEXT,
      closed_at TEXT,
      notes TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS finance_period_checklist (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      period TEXT NOT NULL,
      code TEXT NOT NULL,
      label TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT 'pending',
      completed_by TEXT,
      completed_at TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE (period, code)
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS finance_budgets_monthly (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      period TEXT NOT NULL,
      site_code TEXT,
      cost_center_code TEXT,
      equipment_type TEXT,
      category TEXT NOT NULL,
      budget_amount REAL NOT NULL DEFAULT 0,
      currency TEXT DEFAULT 'USD',
      notes TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE (period, site_code, cost_center_code, equipment_type, category)
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_fbm_period ON finance_budgets_monthly(period)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS finance_forecasts_monthly (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      batch_id TEXT NOT NULL,
      period TEXT NOT NULL,
      site_code TEXT,
      cost_center_code TEXT,
      equipment_type TEXT,
      category TEXT NOT NULL,
      baseline_amount REAL NOT NULL DEFAULT 0,
      uplift_amount REAL NOT NULL DEFAULT 0,
      forecast_amount REAL NOT NULL DEFAULT 0,
      currency TEXT DEFAULT 'USD',
      assumptions_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_ffm_period ON finance_forecasts_monthly(period)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_ffm_batch ON finance_forecasts_monthly(batch_id)`).run();

  /* ============================================================
     PERIOD LOCK + CHECKLIST
  ============================================================ */

  app.get("/periods/:period", async (req, reply) => {
    const period = String(req.params?.period || "").trim();
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period must be YYYY-MM" });
    const lock = ensureLockRow(period);
    return { ok: true, period, lock };
  });

  app.get("/periods/:period/checklist", async (req, reply) => {
    const period = String(req.params?.period || "").trim();
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period must be YYYY-MM" });
    const existing = db.prepare(`SELECT code FROM finance_period_checklist WHERE period = ?`).all(period);
    const existSet = new Set(existing.map((r) => String(r.code)));
    const ins = db.prepare(`
      INSERT OR IGNORE INTO finance_period_checklist (period, code, label)
      VALUES (?, ?, ?)
    `);
    const tx = db.transaction(() => {
      for (const item of DEFAULT_CHECKLIST) {
        if (!existSet.has(item.code)) ins.run(period, item.code, item.label);
      }
    });
    tx();
    const items = db.prepare(`
      SELECT id, period, code, label, status, completed_by, completed_at, notes, updated_at
      FROM finance_period_checklist WHERE period = ?
      ORDER BY id ASC
    `).all(period);
    const lock = ensureLockRow(period);
    return { ok: true, period, lock, items };
  });

  app.post("/periods/:period/checklist/:code", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager"])) return;
    const period = String(req.params?.period || "").trim();
    const code = String(req.params?.code || "").trim();
    const status = String(req.body?.status || "done").trim().toLowerCase();
    const notes = req.body?.notes ? String(req.body.notes).trim() : null;
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period must be YYYY-MM" });
    if (!code) return reply.code(400).send({ error: "code required" });
    if (!["pending", "in_progress", "done", "skipped"].includes(status)) {
      return reply.code(400).send({ error: "invalid status" });
    }
    const existing = db.prepare(`SELECT id FROM finance_period_checklist WHERE period = ? AND code = ?`).get(period, code);
    if (!existing) {
      const template = DEFAULT_CHECKLIST.find((d) => d.code === code);
      const label = template ? template.label : code;
      db.prepare(`
        INSERT INTO finance_period_checklist (period, code, label, status, notes)
        VALUES (?, ?, ?, ?, ?)
      `).run(period, code, label, status, notes);
    } else {
      db.prepare(`
        UPDATE finance_period_checklist
        SET status = ?,
            notes = COALESCE(?, notes),
            completed_by = CASE WHEN ? IN ('done', 'skipped') THEN ? ELSE completed_by END,
            completed_at = CASE WHEN ? IN ('done', 'skipped') THEN datetime('now') ELSE completed_at END,
            updated_at = datetime('now')
        WHERE period = ? AND code = ?
      `).run(status, notes, status, getUser(req), status, period, code);
    }
    writeAudit(db, req, {
      module: "finance", action: "checklist.update", entity_type: "finance_period_checklist",
      entity_id: `${period}:${code}`, payload: { period, code, status, notes }
    });
    return { ok: true, period, code, status };
  });

  app.post("/periods/:period/close", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager"])) return;
    const period = String(req.params?.period || "").trim();
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period must be YYYY-MM" });
    const lock = ensureLockRow(period);
    if (String(lock.status).toLowerCase() === "closed") return { ok: true, period, duplicate: true, status: "closed" };
    const force = Boolean(req.body?.force);
    const remaining = db.prepare(`
      SELECT code, label FROM finance_period_checklist
      WHERE period = ? AND status NOT IN ('done', 'skipped')
    `).all(period);
    if (remaining.length && !force) {
      return reply.code(409).send({
        error: "checklist incomplete",
        remaining
      });
    }
    db.prepare(`
      UPDATE finance_period_locks
      SET status = 'closed', closed_by = ?, closed_at = datetime('now'),
          locked_by = ?, locked_at = datetime('now'),
          updated_at = datetime('now')
      WHERE period = ?
    `).run(getUser(req), getUser(req), period);
    writeAudit(db, req, {
      module: "finance", action: "period.close", entity_type: "finance_period_locks",
      entity_id: period, payload: { period, force, remaining: remaining.length }
    });
    return { ok: true, period, status: "closed", remaining };
  });

  app.post("/periods/:period/reopen", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin"])) return;
    const period = String(req.params?.period || "").trim();
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period must be YYYY-MM" });
    const reason = String(req.body?.reason || "").trim();
    if (!reason) return reply.code(400).send({ error: "reason is required" });
    const lock = ensureLockRow(period);
    if (String(lock.status).toLowerCase() === "open") return { ok: true, period, duplicate: true, status: "open" };
    db.prepare(`
      UPDATE finance_period_locks
      SET status = 'open', reopened_by = ?, reopened_at = datetime('now'),
          reopen_reason = ?, closed_at = NULL, closed_by = NULL,
          locked_at = NULL, locked_by = NULL,
          updated_at = datetime('now')
      WHERE period = ?
    `).run(getUser(req), reason, period);
    writeAudit(db, req, {
      module: "finance", action: "period.reopen", entity_type: "finance_period_locks",
      entity_id: period, payload: { period, reason }
    });
    return { ok: true, period, status: "open", reason };
  });

  /* ============================================================
     MONTHLY ACTUALS (derived from source data + cost settings)
  ============================================================ */

  function buildMonthlyActuals(period, filters = {}) {
    const start = monthStart(period);
    const end = monthEnd(period);
    const defaultCostCenter = filters.default_cost_center_code || null;
    const assetsHasCC = tableHasColumn("assets", "cost_center_code");
    const assetsHasSite = tableHasColumn("assets", "site_code");
    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const smHasCreated = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = smHasCreated ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
    const smHasCC = smCols.some((c) => String(c.name) === "cost_center_code");
    const fuelDefault = readCostSetting("fuel_cost_per_liter_default", 1.5);
    const lubeDefault = readCostSetting("lube_cost_per_qty_default", 4.0);
    const laborDefault = readCostSetting("labor_cost_per_hour_default", 35.0);
    const downtimeDefault = readCostSetting("downtime_cost_per_hour_default", 120.0);

    const partsRows = db.prepare(`
      SELECT
        COALESCE(a.category, '') AS equipment_type,
        ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
        ${smHasCC
          ? `COALESCE(NULLIF(TRIM(sm.cost_center_code), ''), ${assetsHasCC ? "NULLIF(TRIM(a.cost_center_code), '')," : ""} ?)`
          : `COALESCE(${assetsHasCC ? "NULLIF(TRIM(a.cost_center_code), ''), " : ""}?)`} AS cost_center_code,
        SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)) AS amount
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      LEFT JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
      LEFT JOIN assets a ON a.id = w.asset_id
      WHERE sm.movement_type = 'out'
        AND ${smDateExpr} BETWEEN DATE(?) AND DATE(?)
      GROUP BY equipment_type, site_code, cost_center_code
    `).all(defaultCostCenter, start, end);

    const woCols = db.prepare(`PRAGMA table_info(work_orders)`).all();
    const woHasCompleted = woCols.some((c) => String(c.name) === "completed_at");
    const woHasClosed = woCols.some((c) => String(c.name) === "closed_at");
    const woHasLaborHours = woCols.some((c) => String(c.name) === "labor_hours");
    const woHasLaborRate = woCols.some((c) => String(c.name) === "labor_rate_per_hour");
    const woHasCC = woCols.some((c) => String(c.name) === "cost_center_code");
    const laborDateExpr = woHasCompleted && woHasClosed
      ? "DATE(COALESCE(w.completed_at, w.closed_at))"
      : (woHasCompleted ? "DATE(w.completed_at)" : "DATE(w.closed_at)");
    const laborRows = woHasLaborHours ? db.prepare(`
      SELECT
        COALESCE(a.category, '') AS equipment_type,
        ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
        ${woHasCC
          ? `COALESCE(NULLIF(TRIM(w.cost_center_code), ''), ${assetsHasCC ? "NULLIF(TRIM(a.cost_center_code), ''), " : ""} ?)`
          : `COALESCE(${assetsHasCC ? "NULLIF(TRIM(a.cost_center_code), ''), " : ""} ?)`} AS cost_center_code,
        SUM(COALESCE(w.labor_hours, 0) * COALESCE(${woHasLaborRate ? "w.labor_rate_per_hour" : "NULL"}, ?)) AS amount
      FROM work_orders w
      LEFT JOIN assets a ON a.id = w.asset_id
      WHERE w.status IN ('completed','approved','closed')
        AND ${laborDateExpr} BETWEEN DATE(?) AND DATE(?)
      GROUP BY equipment_type, site_code, cost_center_code
    `).all(defaultCostCenter, laborDefault, start, end) : [];

    const flCols = db.prepare(`PRAGMA table_info(fuel_logs)`).all();
    const flHasUnit = flCols.some((c) => String(c.name) === "unit_cost_per_liter");
    const fuelRows = db.prepare(`
      SELECT
        COALESCE(a.category, '') AS equipment_type,
        ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
        ${assetsHasCC ? "COALESCE(NULLIF(TRIM(a.cost_center_code), ''), ?)" : "?"} AS cost_center_code,
        SUM(COALESCE(fl.liters, 0) * COALESCE(${flHasUnit ? "fl.unit_cost_per_liter" : "NULL"}, a.fuel_cost_per_liter, ?)) AS amount
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date BETWEEN DATE(?) AND DATE(?)
      GROUP BY equipment_type, site_code, cost_center_code
    `).all(defaultCostCenter, fuelDefault, start, end);

    const olCols = db.prepare(`PRAGMA table_info(oil_logs)`).all();
    const olHasUnit = olCols.some((c) => String(c.name) === "unit_cost");
    const lubeRows = db.prepare(`
      SELECT
        COALESCE(a.category, '') AS equipment_type,
        ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
        ${assetsHasCC ? "COALESCE(NULLIF(TRIM(a.cost_center_code), ''), ?)" : "?"} AS cost_center_code,
        SUM(COALESCE(ol.quantity, 0) * COALESCE(${olHasUnit ? "ol.unit_cost" : "NULL"}, ?)) AS amount
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN DATE(?) AND DATE(?)
      GROUP BY equipment_type, site_code, cost_center_code
    `).all(defaultCostCenter, lubeDefault, start, end);

    const downtimeRows = db.prepare(`
      SELECT
        COALESCE(a.category, '') AS equipment_type,
        ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
        ${assetsHasCC ? "COALESCE(NULLIF(TRIM(a.cost_center_code), ''), ?)" : "?"} AS cost_center_code,
        SUM(COALESCE(l.hours_down, 0) * COALESCE(a.downtime_cost_per_hour, ?)) AS amount
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date BETWEEN DATE(?) AND DATE(?)
      GROUP BY equipment_type, site_code, cost_center_code
    `).all(defaultCostCenter, downtimeDefault, start, end);

    const out = new Map();
    const addRow = (rows, cat) => {
      for (const r of rows) {
        const amount = Number(r.amount || 0);
        if (amount === 0) continue;
        const key = `${cat}|${r.site_code || ""}|${r.cost_center_code || ""}|${r.equipment_type || ""}`;
        const prev = out.get(key) || {
          period,
          site_code: r.site_code || null,
          cost_center_code: r.cost_center_code || null,
          equipment_type: r.equipment_type || null,
          category: cat,
          actual_amount: 0
        };
        prev.actual_amount += amount;
        out.set(key, prev);
      }
    };
    addRow(partsRows, "parts");
    addRow(laborRows, "labor");
    addRow(fuelRows, "fuel");
    addRow(lubeRows, "lube");
    addRow(downtimeRows, "downtime");
    return Array.from(out.values());
  }

  /* ============================================================
     BUDGETS
  ============================================================ */

  app.post("/budgets/upsert", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager"])) return;
    const rows = Array.isArray(req.body?.rows) ? req.body.rows : [];
    if (!rows.length) return reply.code(400).send({ error: "rows required (array)" });
    const ins = db.prepare(`
      INSERT INTO finance_budgets_monthly
        (period, site_code, cost_center_code, equipment_type, category, budget_amount, currency, notes, created_by)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(period, site_code, cost_center_code, equipment_type, category) DO UPDATE SET
        budget_amount = excluded.budget_amount,
        currency = excluded.currency,
        notes = excluded.notes,
        updated_at = datetime('now')
    `);
    let saved = 0;
    const tx = db.transaction(() => {
      for (const r of rows) {
        const period = String(r?.period || "").trim();
        const category = String(r?.category || "").trim().toLowerCase();
        if (!/^\d{4}-\d{2}$/.test(period)) continue;
        if (!category) continue;
        const site = r?.site_code ? String(r.site_code).trim() : "";
        const cc = r?.cost_center_code ? String(r.cost_center_code).trim() : "";
        const eq = r?.equipment_type ? String(r.equipment_type).trim() : "";
        const amount = Number(r?.budget_amount || 0);
        const currency = String(r?.currency || "USD").trim().toUpperCase() || "USD";
        const notes = r?.notes ? String(r.notes).trim() : null;
        ins.run(period, site, cc, eq, category, amount, currency, notes, getUser(req));
        saved += 1;
      }
    });
    tx();
    writeAudit(db, req, {
      module: "finance", action: "budgets.upsert", entity_type: "finance_budgets_monthly",
      entity_id: String(saved), payload: { rows: saved }
    });
    return { ok: true, saved };
  });

  app.get("/budgets", async (req) => {
    const period = req.query?.period ? String(req.query.period).trim() : null;
    const where = [];
    const args = [];
    if (period) { where.push(`period = ?`); args.push(period); }
    const sql = `
      SELECT id, period, site_code, cost_center_code, equipment_type, category,
             budget_amount, currency, notes, created_by, created_at, updated_at
      FROM finance_budgets_monthly
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY period DESC, site_code, cost_center_code, category
      LIMIT 2000
    `;
    const rows = db.prepare(sql).all(...args);
    return { ok: true, rows };
  });

  app.get("/budgets-vs-actual", async (req, reply) => {
    const period = String(req.query?.period || "").trim();
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period (YYYY-MM) required" });
    const dimension = String(req.query?.dimension || "cost_center_code").trim().toLowerCase();
    const allowedDims = ["site_code", "cost_center_code", "equipment_type", "category"];
    if (!allowedDims.includes(dimension) && dimension !== "combined") {
      return reply.code(400).send({ error: `dimension must be one of: ${allowedDims.join(", ")}, combined` });
    }
    const actuals = buildMonthlyActuals(period);
    const budgets = db.prepare(`
      SELECT period, site_code, cost_center_code, equipment_type, category, budget_amount, currency
      FROM finance_budgets_monthly WHERE period = ?
    `).all(period);

    const keyFor = (r) => dimension === "combined"
      ? `${r.site_code || ""}|${r.cost_center_code || ""}|${r.equipment_type || ""}|${r.category || ""}`
      : String(r[dimension] || "");

    const combined = new Map();
    for (const b of budgets) {
      const k = keyFor(b);
      const prev = combined.get(k) || { budget: 0, actual: 0, lines: 0 };
      prev.budget += Number(b.budget_amount || 0);
      prev.lines += 1;
      combined.set(k, prev);
    }
    for (const a of actuals) {
      const k = keyFor(a);
      const prev = combined.get(k) || { budget: 0, actual: 0, lines: 0 };
      prev.actual += Number(a.actual_amount || 0);
      prev.lines += 1;
      combined.set(k, prev);
    }
    const rows = Array.from(combined.entries()).map(([k, v]) => {
      const variance = Number(v.actual) - Number(v.budget);
      const pct = v.budget > 0 ? (variance / v.budget) * 100 : null;
      return {
        dimension_key: k,
        budget: Number(v.budget.toFixed(2)),
        actual: Number(v.actual.toFixed(2)),
        variance: Number(variance.toFixed(2)),
        variance_pct: pct == null ? null : Number(pct.toFixed(2)),
      };
    }).sort((a, b) => Math.abs(b.variance) - Math.abs(a.variance));

    const total = rows.reduce((acc, r) => {
      acc.budget += r.budget;
      acc.actual += r.actual;
      return acc;
    }, { budget: 0, actual: 0 });
    total.variance = Number((total.actual - total.budget).toFixed(2));
    total.budget = Number(total.budget.toFixed(2));
    total.actual = Number(total.actual.toFixed(2));

    return { ok: true, period, dimension, rows, total };
  });

  /* ============================================================
     FORECAST (hybrid: run-rate baseline + driver-based uplift)
  ============================================================ */

  function weightedBaselineByKey(periods) {
    const weights = [0.2, 0.3, 0.5];
    const byKey = new Map();
    periods.forEach((period, idx) => {
      const actuals = buildMonthlyActuals(period);
      for (const a of actuals) {
        const key = `${a.site_code || ""}|${a.cost_center_code || ""}|${a.equipment_type || ""}|${a.category}`;
        const prev = byKey.get(key) || {
          site_code: a.site_code,
          cost_center_code: a.cost_center_code,
          equipment_type: a.equipment_type,
          category: a.category,
          weighted_sum: 0,
          inputs: []
        };
        prev.weighted_sum += Number(a.actual_amount || 0) * weights[idx];
        prev.inputs.push({ period, amount: a.actual_amount, weight: weights[idx] });
        byKey.set(key, prev);
      }
    });
    return byKey;
  }

  function computeDrivers(nextPeriod) {
    const start = monthStart(nextPeriod);
    const end = monthEnd(nextPeriod);
    const openBreakdowns = db.prepare(`
      SELECT COUNT(*) AS c FROM breakdowns WHERE LOWER(COALESCE(status,'')) IN ('open','reported','in_progress')
    `).get();
    const recentDown = db.prepare(`
      SELECT COALESCE(SUM(hours_down), 0) AS h
      FROM breakdown_downtime_logs
      WHERE log_date >= DATE('now','-30 day')
    `).get();
    let plannedPm = 0;
    try {
      const row = db.prepare(`
        SELECT COUNT(*) AS c FROM maintenance_plans WHERE DATE(COALESCE(next_due, scheduled_date)) BETWEEN DATE(?) AND DATE(?)
      `).get(start, end);
      plannedPm = Number(row?.c || 0);
    } catch { plannedPm = 0; }
    return {
      open_breakdowns: Number(openBreakdowns?.c || 0),
      recent_downtime_hours: Number(recentDown?.h || 0),
      planned_pm_count: plannedPm,
    };
  }

  function upliftFactor(category, drivers) {
    // lightweight heuristic multipliers; scoped to categories most sensitive to drivers
    const base = 1.0;
    const pmBoost = drivers.planned_pm_count > 0 ? Math.min(0.15, drivers.planned_pm_count * 0.02) : 0;
    const breakdownBoost = drivers.open_breakdowns > 0 ? Math.min(0.2, drivers.open_breakdowns * 0.03) : 0;
    const downtimeBoost = drivers.recent_downtime_hours > 50
      ? Math.min(0.25, (drivers.recent_downtime_hours - 50) / 400)
      : 0;
    switch (category) {
      case "parts": return base + pmBoost + breakdownBoost * 0.5;
      case "labor": return base + pmBoost + breakdownBoost * 0.7;
      case "downtime": return base + breakdownBoost + downtimeBoost;
      case "fuel": return base + pmBoost * 0.3;
      case "lube": return base + pmBoost * 0.5;
      default: return base;
    }
  }

  app.post("/forecast/rebuild", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager"])) return;
    const start = String(req.body?.start_period || "").trim();
    const monthsAhead = Math.max(1, Math.min(6, Number(req.body?.months || 3)));
    if (!/^\d{4}-\d{2}$/.test(start)) return reply.code(400).send({ error: "start_period (YYYY-MM) required" });

    const batchId = `FCAST-${start}-${Date.now()}`;
    const historyPeriods = [addMonths(start, -3), addMonths(start, -2), addMonths(start, -1)];
    const baseline = weightedBaselineByKey(historyPeriods);

    const insForecast = db.prepare(`
      INSERT INTO finance_forecasts_monthly
        (batch_id, period, site_code, cost_center_code, equipment_type, category,
         baseline_amount, uplift_amount, forecast_amount, currency, assumptions_json)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    let saved = 0;
    const tx = db.transaction(() => {
      for (let i = 0; i < monthsAhead; i += 1) {
        const period = addMonths(start, i);
        const drivers = computeDrivers(period);
        for (const [, row] of baseline) {
          const baseAmt = Number(row.weighted_sum || 0);
          if (baseAmt <= 0) continue;
          const factor = upliftFactor(row.category, drivers);
          const uplift = baseAmt * (factor - 1.0);
          const forecast = baseAmt + uplift;
          const assumptions = {
            baseline_periods: historyPeriods,
            weights: [0.2, 0.3, 0.5],
            drivers,
            uplift_factor: Number(factor.toFixed(4)),
          };
          insForecast.run(
            batchId, period, row.site_code || "", row.cost_center_code || "", row.equipment_type || "",
            row.category, Number(baseAmt.toFixed(2)), Number(uplift.toFixed(2)), Number(forecast.toFixed(2)),
            "USD", JSON.stringify(assumptions)
          );
          saved += 1;
        }
      }
    });
    tx();
    writeAudit(db, req, {
      module: "finance", action: "forecast.rebuild", entity_type: "finance_forecasts_monthly",
      entity_id: batchId, payload: { batchId, start_period: start, months: monthsAhead, saved }
    });
    return { ok: true, batch_id: batchId, start_period: start, months: monthsAhead, saved };
  });

  app.get("/forecast", async (req) => {
    const batchId = req.query?.batch_id ? String(req.query.batch_id).trim() : null;
    const period = req.query?.period ? String(req.query.period).trim() : null;
    const where = [];
    const args = [];
    if (batchId) { where.push(`batch_id = ?`); args.push(batchId); }
    if (period) { where.push(`period = ?`); args.push(period); }
    if (!batchId) {
      const latest = db.prepare(`SELECT batch_id FROM finance_forecasts_monthly ORDER BY id DESC LIMIT 1`).get();
      if (latest?.batch_id) {
        where.push(`batch_id = ?`);
        args.push(latest.batch_id);
      }
    }
    const sql = `
      SELECT id, batch_id, period, site_code, cost_center_code, equipment_type, category,
             baseline_amount, uplift_amount, forecast_amount, currency, assumptions_json, created_at
      FROM finance_forecasts_monthly
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY period ASC, site_code, cost_center_code, category
      LIMIT 5000
    `;
    const rows = db.prepare(sql).all(...args);
    const totals = new Map();
    for (const r of rows) {
      const p = r.period;
      const prev = totals.get(p) || { period: p, baseline: 0, uplift: 0, forecast: 0 };
      prev.baseline += Number(r.baseline_amount || 0);
      prev.uplift += Number(r.uplift_amount || 0);
      prev.forecast += Number(r.forecast_amount || 0);
      totals.set(p, prev);
    }
    const byPeriod = Array.from(totals.values()).map((t) => ({
      period: t.period,
      baseline: Number(t.baseline.toFixed(2)),
      uplift: Number(t.uplift.toFixed(2)),
      forecast: Number(t.forecast.toFixed(2)),
    })).sort((a, b) => String(a.period).localeCompare(String(b.period)));
    return { ok: true, rows, totals_by_period: byPeriod };
  });

  /* ============================================================
     KPI DEFINITIONS (single source of truth registry)
  ============================================================ */

  app.get("/kpis/definitions", async () => {
    return { ok: true, kpis: KPI_DEFINITIONS };
  });

  app.get("/kpis/definitions/:code", async (req, reply) => {
    const code = String(req.params?.code || "").trim().toLowerCase();
    const def = KPI_DEFINITIONS.find((d) => d.code === code);
    if (!def) return reply.code(404).send({ error: "kpi not found" });
    return { ok: true, kpi: def };
  });

  /* ============================================================
     SSOT REPORT SET (single finance KPI + cost summary)
  ============================================================ */

  function computeKpiSnapshotForPeriod(period) {
    const start = monthStart(period);
    const end = monthEnd(period);
    const scheduled = 10;
    const dhRows = db.prepare(`
      SELECT dh.log_date,
             COUNT(DISTINCT dh.asset_id) AS used_assets,
             COALESCE(SUM(dh.hours_run), 0) AS run_hours
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.log_date BETWEEN DATE(?) AND DATE(?)
        AND dh.hours_run > 0
        AND a.is_standby = 0
      GROUP BY dh.log_date
    `).all(start, end);
    const run_hours = dhRows.reduce((s, r) => s + Number(r.run_hours || 0), 0);
    const available_hours = dhRows.reduce((s, r) => s + Number(r.used_assets) * scheduled, 0);
    const downtimeRow = db.prepare(`
      SELECT COALESCE(SUM(hours_down), 0) AS h
      FROM breakdown_downtime_logs WHERE log_date BETWEEN DATE(?) AND DATE(?)
    `).get(start, end);
    const downtime_hours = Number(downtimeRow?.h || 0);
    const breakdownCountRow = db.prepare(`
      SELECT COUNT(*) AS c FROM breakdowns
      WHERE DATE(COALESCE(reported_at, created_at, updated_at)) BETWEEN DATE(?) AND DATE(?)
    `).get(start, end);
    const breakdowns = Number(breakdownCountRow?.c || 0);
    const availability = available_hours > 0 ? ((available_hours - downtime_hours) / available_hours) * 100 : null;
    const utilization = available_hours > 0 ? (run_hours / available_hours) * 100 : null;
    const mtbf = breakdowns > 0 ? run_hours / breakdowns : null;
    const mttr = breakdowns > 0 ? downtime_hours / breakdowns : null;
    const actuals = buildMonthlyActuals(period);
    const totalCost = actuals.reduce((s, a) => s + Number(a.actual_amount || 0), 0);
    const cost_per_asset_hour = run_hours > 0 ? totalCost / run_hours : null;
    return {
      period,
      available_hours: Number(available_hours.toFixed(2)),
      run_hours: Number(run_hours.toFixed(2)),
      downtime_hours: Number(downtime_hours.toFixed(2)),
      breakdowns,
      availability: availability == null ? null : Number(availability.toFixed(2)),
      utilization: utilization == null ? null : Number(utilization.toFixed(2)),
      mtbf: mtbf == null ? null : Number(mtbf.toFixed(2)),
      mttr: mttr == null ? null : Number(mttr.toFixed(2)),
      total_cost: Number(totalCost.toFixed(2)),
      cost_per_asset_hour: cost_per_asset_hour == null ? null : Number(cost_per_asset_hour.toFixed(4)),
    };
  }

  app.get("/reports/ssot", async (req, reply) => {
    const period = String(req.query?.period || "").trim();
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period (YYYY-MM) required" });
    const snapshot = computeKpiSnapshotForPeriod(period);
    const actualsAgg = buildMonthlyActuals(period).map((r) => ({
      ...r,
      actual_amount: Number(Number(r.actual_amount || 0).toFixed(2)),
    }));
    const budgets = db.prepare(`
      SELECT category, SUM(budget_amount) AS budget
      FROM finance_budgets_monthly WHERE period = ? GROUP BY category
    `).all(period);
    const forecasts = db.prepare(`
      SELECT category, SUM(forecast_amount) AS forecast
      FROM finance_forecasts_monthly WHERE period = ? GROUP BY category
    `).all(period);
    const lock = ensureLockRow(period);
    return {
      ok: true,
      period,
      kpi: snapshot,
      actuals: actualsAgg,
      budgets,
      forecasts,
      lock,
      kpi_definitions: KPI_DEFINITIONS,
    };
  });

  app.get("/reports/ssot/export.csv", async (req, reply) => {
    const period = String(req.query?.period || "").trim();
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period (YYYY-MM) required" });
    const snapshot = computeKpiSnapshotForPeriod(period);
    const actualsAgg = buildMonthlyActuals(period);
    const esc = (v) => {
      const s = String(v ?? "");
      if (/[",\n]/.test(s)) return `"${s.replace(/"/g, "\"\"")}"`;
      return s;
    };
    const lines = [];
    lines.push(`# IRONLOG SSOT Report ${period}`);
    lines.push(`metric,value`);
    for (const [k, v] of Object.entries(snapshot)) lines.push(`${esc(k)},${esc(v)}`);
    lines.push("");
    lines.push("site_code,cost_center_code,equipment_type,category,actual_amount");
    for (const r of actualsAgg) {
      lines.push([
        r.site_code || "", r.cost_center_code || "", r.equipment_type || "", r.category,
        Number(r.actual_amount || 0).toFixed(2)
      ].map(esc).join(","));
    }
    reply.header("Content-Type", "text/csv; charset=utf-8");
    reply.header("Content-Disposition", `attachment; filename="ssot-${period}.csv"`);
    return reply.send(lines.join("\n"));
  });

  app.get("/reports/ssot/export.xlsx", async (req, reply) => {
    const period = String(req.query?.period || "").trim();
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period (YYYY-MM) required" });
    const snapshot = computeKpiSnapshotForPeriod(period);
    const actualsAgg = buildMonthlyActuals(period);
    const budgets = db.prepare(`
      SELECT category, SUM(budget_amount) AS budget
      FROM finance_budgets_monthly WHERE period = ? GROUP BY category
    `).all(period);
    const forecasts = db.prepare(`
      SELECT category, SUM(forecast_amount) AS forecast
      FROM finance_forecasts_monthly WHERE period = ? GROUP BY category
    `).all(period);
    const ExcelJS = (await import("exceljs")).default;
    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const wsKpi = wb.addWorksheet("KPIs");
    wsKpi.columns = [
      { header: "Metric", key: "metric", width: 28 },
      { header: "Value", key: "value", width: 20 },
    ];
    for (const [k, v] of Object.entries(snapshot)) wsKpi.addRow({ metric: k, value: v });
    wsKpi.getRow(1).font = { bold: true };

    const wsActuals = wb.addWorksheet("Actuals");
    wsActuals.columns = [
      { header: "Site", key: "site_code", width: 14 },
      { header: "Cost Center", key: "cost_center_code", width: 18 },
      { header: "Equipment Type", key: "equipment_type", width: 18 },
      { header: "Category", key: "category", width: 14 },
      { header: "Actual", key: "actual_amount", width: 14 },
    ];
    wsActuals.addRows(actualsAgg.map((r) => ({
      site_code: r.site_code || "",
      cost_center_code: r.cost_center_code || "",
      equipment_type: r.equipment_type || "",
      category: r.category,
      actual_amount: Number(Number(r.actual_amount || 0).toFixed(2)),
    })));
    wsActuals.getRow(1).font = { bold: true };
    wsActuals.getColumn("actual_amount").numFmt = "#,##0.00";

    const wsBudget = wb.addWorksheet("Budget");
    wsBudget.columns = [
      { header: "Category", key: "category", width: 16 },
      { header: "Budget", key: "budget", width: 16 },
    ];
    wsBudget.addRows(budgets.map((r) => ({
      category: r.category,
      budget: Number(Number(r.budget || 0).toFixed(2)),
    })));
    wsBudget.getRow(1).font = { bold: true };
    wsBudget.getColumn("budget").numFmt = "#,##0.00";

    const wsForecast = wb.addWorksheet("Forecast");
    wsForecast.columns = [
      { header: "Category", key: "category", width: 16 },
      { header: "Forecast", key: "forecast", width: 16 },
    ];
    wsForecast.addRows(forecasts.map((r) => ({
      category: r.category,
      forecast: Number(Number(r.forecast || 0).toFixed(2)),
    })));
    wsForecast.getRow(1).font = { bold: true };
    wsForecast.getColumn("forecast").numFmt = "#,##0.00";

    const wsDef = wb.addWorksheet("KPI Definitions");
    wsDef.columns = [
      { header: "Code", key: "code", width: 24 },
      { header: "Label", key: "label", width: 30 },
      { header: "Unit", key: "unit", width: 20 },
      { header: "Formula", key: "formula", width: 60 },
      { header: "Source Tables", key: "source_tables", width: 40 },
      { header: "Exclusions", key: "exclusions", width: 40 },
    ];
    wsDef.addRows(KPI_DEFINITIONS.map((d) => ({
      code: d.code,
      label: d.label,
      unit: d.unit,
      formula: d.formula,
      source_tables: (d.source_tables || []).join(", "),
      exclusions: (d.exclusions || []).join(", "),
    })));
    wsDef.getRow(1).font = { bold: true };

    const buffer = await wb.xlsx.writeBuffer();
    reply.header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    reply.header("Content-Disposition", `attachment; filename="ssot-${period}.xlsx"`);
    return reply.send(Buffer.from(buffer));
  });
}
