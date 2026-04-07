import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function isShift(v) {
  const s = String(v || "").trim().toLowerCase();
  return s === "day" || s === "night";
}

function toNumberOrNull(v) {
  if (v == null || String(v).trim() === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function getRole(req) {
  return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
}
function getSiteCode(req) {
  return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
}

function requireRoles(req, reply, roles) {
  const role = getRole(req);
  if (!roles.includes(role)) {
    reply.code(403).send({ error: `forbidden: role '${role}' cannot perform this action` });
    return false;
  }
  return true;
}

export default async function operationsRoutes(app) {
  ensureAuditTable(db);
  db.prepare(`
    CREATE TABLE IF NOT EXISTS operations_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      op_date TEXT NOT NULL DEFAULT (date('now')),
      tonnes_moved REAL,
      product_type TEXT,
      product_produced REAL,
      trucks_loaded INTEGER,
      loads_count INTEGER,
      crusher_feed_tonnes REAL,
      crusher_output_tonnes REAL,
      crusher_hours REAL,
      crusher_downtime_hours REAL,
      weighbridge_amount REAL,
      trucks_delivered INTEGER,
      product_delivered REAL,
      client_delivered_to TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  const opCols = db.prepare(`PRAGMA table_info(operations_logs)`).all();
  if (!opCols.some((c) => String(c.name || "") === "site_code")) {
    db.prepare(`ALTER TABLE operations_logs ADD COLUMN site_code TEXT`).run();
  }
  if (!opCols.some((c) => String(c.name || "") === "loads_count")) {
    db.prepare(`ALTER TABLE operations_logs ADD COLUMN loads_count INTEGER`).run();
  }
  if (!opCols.some((c) => String(c.name || "") === "crusher_feed_tonnes")) {
    db.prepare(`ALTER TABLE operations_logs ADD COLUMN crusher_feed_tonnes REAL`).run();
  }
  if (!opCols.some((c) => String(c.name || "") === "crusher_output_tonnes")) {
    db.prepare(`ALTER TABLE operations_logs ADD COLUMN crusher_output_tonnes REAL`).run();
  }
  if (!opCols.some((c) => String(c.name || "") === "crusher_hours")) {
    db.prepare(`ALTER TABLE operations_logs ADD COLUMN crusher_hours REAL`).run();
  }
  if (!opCols.some((c) => String(c.name || "") === "crusher_downtime_hours")) {
    db.prepare(`ALTER TABLE operations_logs ADD COLUMN crusher_downtime_hours REAL`).run();
  }
  db.prepare(`
    CREATE TABLE IF NOT EXISTS operations_daily_closing (
      op_date TEXT PRIMARY KEY,
      shift_name TEXT,
      supervisor_name TEXT,
      variance_note TEXT,
      checklist_weighbridge_reconciled INTEGER NOT NULL DEFAULT 0,
      checklist_trucks_reconciled INTEGER NOT NULL DEFAULT 0,
      checklist_client_confirmed INTEGER NOT NULL DEFAULT 0,
      status TEXT NOT NULL DEFAULT 'open',
      closed_at TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS site_zones (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL,
      name TEXT NOT NULL,
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(site_code, name)
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS operations_targets (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL,
      target_date TEXT NOT NULL,
      material_type TEXT NOT NULL,
      target_tonnage REAL NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(site_code, target_date, material_type)
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS operations_daily (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL,
      op_date TEXT NOT NULL,
      shift TEXT NOT NULL,
      material_type TEXT NOT NULL,
      zone_id INTEGER,
      planned_tonnage REAL,
      actual_tonnage REAL,
      loads_count INTEGER,
      avg_cycle_time REAL,
      operator_name TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (zone_id) REFERENCES site_zones(id) ON DELETE SET NULL
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS operations_equipment_usage (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL,
      usage_date TEXT NOT NULL,
      operations_daily_id INTEGER NOT NULL,
      asset_id INTEGER NOT NULL,
      role TEXT NOT NULL,
      hours_used REAL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (operations_daily_id) REFERENCES operations_daily(id) ON DELETE CASCADE,
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS operations_delays (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL,
      delay_date TEXT NOT NULL,
      delay_type TEXT NOT NULL,
      hours_lost REAL NOT NULL DEFAULT 0,
      impact_tonnage REAL,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  // Site Zones
  app.get("/site/zones", async (req) => {
    const site_code = getSiteCode(req);
    const rows = db.prepare(`
      SELECT id, site_code, name, active, created_at
      FROM site_zones
      WHERE site_code = ?
      ORDER BY active DESC, name ASC
    `).all(site_code);
    return { ok: true, rows };
  });

  app.post("/site/zones", async (req, reply) => {
    const site_code = getSiteCode(req);
    const name = String(req.body?.name || "").trim();
    if (!name) return reply.code(400).send({ error: "name is required" });
    const active = req.body?.active === false ? 0 : 1;
    db.prepare(`
      INSERT INTO site_zones (site_code, name, active)
      VALUES (?, ?, ?)
      ON CONFLICT(site_code, name) DO UPDATE SET active = excluded.active
    `).run(site_code, name, active);
    return { ok: true };
  });

  // Daily Production Capture
  app.post("/site/daily", async (req, reply) => {
    const site_code = getSiteCode(req);
    const op_date = String(req.body?.op_date || "").trim();
    const shift = String(req.body?.shift || "").trim().toLowerCase();
    const material_type = String(req.body?.material_type || "").trim();
    if (!isDate(op_date)) return reply.code(400).send({ error: "op_date must be YYYY-MM-DD" });
    if (!isShift(shift)) return reply.code(400).send({ error: "shift must be day or night" });
    if (!material_type) return reply.code(400).send({ error: "material_type is required" });

    const zone_id = toNumberOrNull(req.body?.zone_id);
    const planned_tonnage = toNumberOrNull(req.body?.planned_tonnage);
    const actual_tonnage = toNumberOrNull(req.body?.actual_tonnage);
    const loads_count = toNumberOrNull(req.body?.loads_count);
    const avg_cycle_time = toNumberOrNull(req.body?.avg_cycle_time);
    const operator_name = String(req.body?.operator_name || "").trim() || null;
    const notes = String(req.body?.notes || "").trim() || null;

    const ins = db.prepare(`
      INSERT INTO operations_daily (
        site_code, op_date, shift, material_type, zone_id, planned_tonnage, actual_tonnage,
        loads_count, avg_cycle_time, operator_name, notes
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(
      site_code,
      op_date,
      shift,
      material_type,
      zone_id == null ? null : Math.round(zone_id),
      planned_tonnage,
      actual_tonnage,
      loads_count == null ? null : Math.round(loads_count),
      avg_cycle_time,
      operator_name,
      notes
    );
    return { ok: true, id: Number(ins.lastInsertRowid) };
  });

  app.get("/site/daily", async (req, reply) => {
    const site_code = getSiteCode(req);
    const from = String(req.query?.from || "").trim();
    const to = String(req.query?.to || "").trim();
    if (from && !isDate(from)) return reply.code(400).send({ error: "from must be YYYY-MM-DD" });
    if (to && !isDate(to)) return reply.code(400).send({ error: "to must be YYYY-MM-DD" });

    const where = ["d.site_code = ?"];
    const params = [site_code];
    if (from) {
      where.push("d.op_date >= ?");
      params.push(from);
    }
    if (to) {
      where.push("d.op_date <= ?");
      params.push(to);
    }

    const rows = db.prepare(`
      SELECT
        d.id, d.site_code, d.op_date, d.shift, d.material_type, d.zone_id,
        z.name AS zone_name,
        d.planned_tonnage, d.actual_tonnage, d.loads_count, d.avg_cycle_time,
        d.operator_name, d.notes, d.created_at
      FROM operations_daily d
      LEFT JOIN site_zones z ON z.id = d.zone_id
      WHERE ${where.join(" AND ")}
      ORDER BY d.op_date DESC, d.shift DESC, d.id DESC
      LIMIT 1000
    `).all(...params);
    return { ok: true, rows };
  });

  // Equipment contribution to production
  app.post("/site/daily/:id/equipment", async (req, reply) => {
    const site_code = getSiteCode(req);
    const operations_daily_id = Number(req.params?.id || 0);
    const asset_id = Number(req.body?.asset_id || 0);
    const role = String(req.body?.role || "").trim();
    const hours_used = toNumberOrNull(req.body?.hours_used);
    if (!operations_daily_id) return reply.code(400).send({ error: "invalid daily id" });
    if (!asset_id) return reply.code(400).send({ error: "asset_id is required" });
    if (!role) return reply.code(400).send({ error: "role is required" });

    const daily = db.prepare(`
      SELECT id, op_date FROM operations_daily WHERE id = ? AND site_code = ?
    `).get(operations_daily_id, site_code);
    if (!daily) return reply.code(404).send({ error: "daily entry not found" });

    const ins = db.prepare(`
      INSERT INTO operations_equipment_usage (
        site_code, usage_date, operations_daily_id, asset_id, role, hours_used
      ) VALUES (?, ?, ?, ?, ?, ?)
    `).run(site_code, String(daily.op_date), operations_daily_id, asset_id, role, hours_used);
    return { ok: true, id: Number(ins.lastInsertRowid) };
  });

  app.get("/site/daily/:id/equipment", async (req, reply) => {
    const site_code = getSiteCode(req);
    const operations_daily_id = Number(req.params?.id || 0);
    if (!operations_daily_id) return reply.code(400).send({ error: "invalid daily id" });
    const rows = db.prepare(`
      SELECT
        e.id, e.operations_daily_id, e.asset_id, a.asset_code, a.asset_name,
        e.role, e.hours_used, e.created_at
      FROM operations_equipment_usage e
      LEFT JOIN assets a ON a.id = e.asset_id
      WHERE e.site_code = ? AND e.operations_daily_id = ?
      ORDER BY e.id DESC
      LIMIT 300
    `).all(site_code, operations_daily_id);
    return { ok: true, rows };
  });

  // Targets
  app.post("/site/targets", async (req, reply) => {
    const site_code = getSiteCode(req);
    const target_date = String(req.body?.target_date || "").trim();
    const material_type = String(req.body?.material_type || "").trim();
    const target_tonnage = toNumberOrNull(req.body?.target_tonnage);
    if (!isDate(target_date)) return reply.code(400).send({ error: "target_date must be YYYY-MM-DD" });
    if (!material_type) return reply.code(400).send({ error: "material_type is required" });
    if (target_tonnage == null) return reply.code(400).send({ error: "target_tonnage is required" });

    db.prepare(`
      INSERT INTO operations_targets (site_code, target_date, material_type, target_tonnage)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(site_code, target_date, material_type) DO UPDATE SET
        target_tonnage = excluded.target_tonnage,
        updated_at = datetime('now')
    `).run(site_code, target_date, material_type, target_tonnage);
    return { ok: true };
  });

  app.get("/site/targets", async (req, reply) => {
    const site_code = getSiteCode(req);
    const from = String(req.query?.from || "").trim();
    const to = String(req.query?.to || "").trim();
    if (from && !isDate(from)) return reply.code(400).send({ error: "from must be YYYY-MM-DD" });
    if (to && !isDate(to)) return reply.code(400).send({ error: "to must be YYYY-MM-DD" });

    const where = ["site_code = ?"];
    const params = [site_code];
    if (from) {
      where.push("target_date >= ?");
      params.push(from);
    }
    if (to) {
      where.push("target_date <= ?");
      params.push(to);
    }

    const rows = db.prepare(`
      SELECT id, site_code, target_date, material_type, target_tonnage, created_at, updated_at
      FROM operations_targets
      WHERE ${where.join(" AND ")}
      ORDER BY target_date DESC, material_type ASC, id DESC
      LIMIT 1000
    `).all(...params);
    return { ok: true, rows };
  });

  // Operational delays (non-mechanical)
  app.post("/site/delays", async (req, reply) => {
    const site_code = getSiteCode(req);
    const delay_date = String(req.body?.delay_date || "").trim();
    const delay_type = String(req.body?.delay_type || "").trim();
    const hours_lost = toNumberOrNull(req.body?.hours_lost);
    const impact_tonnage = toNumberOrNull(req.body?.impact_tonnage);
    const notes = String(req.body?.notes || "").trim() || null;
    if (!isDate(delay_date)) return reply.code(400).send({ error: "delay_date must be YYYY-MM-DD" });
    if (!delay_type) return reply.code(400).send({ error: "delay_type is required" });
    if (hours_lost == null) return reply.code(400).send({ error: "hours_lost is required" });

    const ins = db.prepare(`
      INSERT INTO operations_delays (site_code, delay_date, delay_type, hours_lost, impact_tonnage, notes)
      VALUES (?, ?, ?, ?, ?, ?)
    `).run(site_code, delay_date, delay_type, hours_lost, impact_tonnage, notes);
    return { ok: true, id: Number(ins.lastInsertRowid) };
  });

  app.get("/site/delays", async (req, reply) => {
    const site_code = getSiteCode(req);
    const from = String(req.query?.from || "").trim();
    const to = String(req.query?.to || "").trim();
    if (from && !isDate(from)) return reply.code(400).send({ error: "from must be YYYY-MM-DD" });
    if (to && !isDate(to)) return reply.code(400).send({ error: "to must be YYYY-MM-DD" });

    const where = ["site_code = ?"];
    const params = [site_code];
    if (from) {
      where.push("delay_date >= ?");
      params.push(from);
    }
    if (to) {
      where.push("delay_date <= ?");
      params.push(to);
    }

    const rows = db.prepare(`
      SELECT id, site_code, delay_date, delay_type, hours_lost, impact_tonnage, notes, created_at
      FROM operations_delays
      WHERE ${where.join(" AND ")}
      ORDER BY delay_date DESC, id DESC
      LIMIT 1000
    `).all(...params);
    return { ok: true, rows };
  });

  // Dashboard summary
  app.get("/site/dashboard", async (req, reply) => {
    const site_code = getSiteCode(req);
    const date = String(req.query?.date || "").trim() || new Date().toISOString().slice(0, 10);
    if (!isDate(date)) return reply.code(400).send({ error: "date must be YYYY-MM-DD" });

    const today = db.prepare(`
      SELECT
        COALESCE(SUM(actual_tonnage), 0) AS actual_tonnage,
        COALESCE(SUM(planned_tonnage), 0) AS planned_tonnage,
        COALESCE(SUM(loads_count), 0) AS loads_count,
        COALESCE(COUNT(DISTINCT zone_id), 0) AS active_zones
      FROM operations_daily
      WHERE site_code = ? AND op_date = ?
    `).get(site_code, date);

    const target = db.prepare(`
      SELECT COALESCE(SUM(target_tonnage), 0) AS target_tonnage
      FROM operations_targets
      WHERE site_code = ? AND target_date = ?
    `).get(site_code, date);

    const weekRows = db.prepare(`
      SELECT op_date, COALESCE(SUM(actual_tonnage), 0) AS total_actual
      FROM operations_daily
      WHERE site_code = ? AND op_date BETWEEN date(?, '-6 day') AND ?
      GROUP BY op_date
      ORDER BY op_date ASC
    `).all(site_code, date, date);

    let best_day = null;
    let worst_day = null;
    let week_total = 0;
    for (const r of weekRows) {
      const v = Number(r.total_actual || 0);
      week_total += v;
      if (!best_day || v > Number(best_day.total_actual || 0)) best_day = r;
      if (!worst_day || v < Number(worst_day.total_actual || 0)) worst_day = r;
    }

    const losses = db.prepare(`
      SELECT
        COALESCE((
          SELECT SUM(hours_down)
          FROM breakdown_downtime_logs
          WHERE log_date BETWEEN date(?, '-6 day') AND ?
        ), 0) AS breakdown_hours,
        COALESCE((
          SELECT SUM(hours_lost)
          FROM operations_delays
          WHERE site_code = ? AND delay_date BETWEEN date(?, '-6 day') AND ?
        ), 0) AS operational_delay_hours
    `).get(date, date, site_code, date, date);

    const actual = Number(today?.actual_tonnage || 0);
    const targetTons = Number(target?.target_tonnage || 0);
    const achieved_pct = targetTons > 0 ? (actual / targetTons) * 100 : 0;

    return {
      ok: true,
      date,
      today: {
        total_tons_produced: actual,
        planned_tonnage: Number(today?.planned_tonnage || 0),
        target_tonnage: targetTons,
        achieved_pct,
        loads_moved: Number(today?.loads_count || 0),
        active_zones: Number(today?.active_zones || 0),
      },
      week: {
        total_production: week_total,
        best_day: best_day ? { date: best_day.op_date, tons: Number(best_day.total_actual || 0) } : null,
        worst_day: worst_day ? { date: worst_day.op_date, tons: Number(worst_day.total_actual || 0) } : null,
      },
      losses: {
        breakdown_hours: Number(losses?.breakdown_hours || 0),
        operational_delay_hours: Number(losses?.operational_delay_hours || 0),
      },
    };
  });

  app.post("/", async (req, reply) => {
    const site_code = getSiteCode(req);
    const op_date = String(req.body?.op_date || "").trim() || new Date().toISOString().slice(0, 10);
    if (!isDate(op_date)) return reply.code(400).send({ error: "op_date must be YYYY-MM-DD" });

    const tonnes_moved = toNumberOrNull(req.body?.tonnes_moved);
    const product_type = String(req.body?.product_type || "").trim() || null;
    const product_produced = toNumberOrNull(req.body?.product_produced);
    const trucks_loaded = toNumberOrNull(req.body?.trucks_loaded);
    const loads_count = toNumberOrNull(req.body?.loads_count);
    const crusher_feed_tonnes = toNumberOrNull(req.body?.crusher_feed_tonnes);
    const crusher_output_tonnes = toNumberOrNull(req.body?.crusher_output_tonnes);
    const crusher_hours = toNumberOrNull(req.body?.crusher_hours);
    const crusher_downtime_hours = toNumberOrNull(req.body?.crusher_downtime_hours);
    const weighbridge_amount = toNumberOrNull(req.body?.weighbridge_amount);
    const trucks_delivered = toNumberOrNull(req.body?.trucks_delivered);
    const product_delivered = toNumberOrNull(req.body?.product_delivered);
    const client_delivered_to = String(req.body?.client_delivered_to || "").trim() || null;
    const notes = String(req.body?.notes || "").trim() || null;

    const ins = db.prepare(`
      INSERT INTO operations_logs (
        op_date, tonnes_moved, product_type, product_produced, trucks_loaded, loads_count,
        crusher_feed_tonnes, crusher_output_tonnes, crusher_hours, crusher_downtime_hours, weighbridge_amount,
        trucks_delivered, product_delivered, client_delivered_to, notes, site_code
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(
      op_date,
      tonnes_moved,
      product_type,
      product_produced,
      trucks_loaded == null ? null : Math.round(trucks_loaded),
      loads_count == null ? null : Math.round(loads_count),
      crusher_feed_tonnes,
      crusher_output_tonnes,
      crusher_hours,
      crusher_downtime_hours,
      weighbridge_amount,
      trucks_delivered == null ? null : Math.round(trucks_delivered),
      product_delivered,
      client_delivered_to,
      notes,
      site_code
    );

    return { ok: true, id: Number(ins.lastInsertRowid) };
  });

  app.get("/", async (req, reply) => {
    const site_code = getSiteCode(req);
    const from = String(req.query?.from || "").trim();
    const to = String(req.query?.to || "").trim();
    if (from && !isDate(from)) return reply.code(400).send({ error: "from must be YYYY-MM-DD" });
    if (to && !isDate(to)) return reply.code(400).send({ error: "to must be YYYY-MM-DD" });

    const where = [];
    const params = [];
    if (from) {
      where.push("op_date >= ?");
      params.push(from);
    }
    if (to) {
      where.push("op_date <= ?");
      params.push(to);
    }
    where.push("COALESCE(site_code,'main') = ?");
    params.push(site_code);

    const rows = db.prepare(`
      SELECT
        id, site_code, op_date, tonnes_moved, product_type, product_produced, trucks_loaded, loads_count,
        crusher_feed_tonnes, crusher_output_tonnes, crusher_hours, crusher_downtime_hours, weighbridge_amount,
        trucks_delivered, product_delivered, client_delivered_to, notes, created_at
      FROM operations_logs
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY op_date DESC, id DESC
      LIMIT 500
    `).all(...params);

    return { ok: true, rows };
  });

  // GET /api/operations/closing/:date
  app.get("/closing/:date", async (req, reply) => {
    const op_date = String(req.params?.date || "").trim();
    if (!isDate(op_date)) return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
    const row = db.prepare(`
      SELECT
        op_date, shift_name, supervisor_name, variance_note,
        checklist_weighbridge_reconciled, checklist_trucks_reconciled, checklist_client_confirmed,
        status, closed_at, updated_at
      FROM operations_daily_closing
      WHERE op_date = ?
    `).get(op_date);
    return { ok: true, row: row || null };
  });

  // POST /api/operations/closing
  app.post("/closing", async (req, reply) => {
    const op_date = String(req.body?.op_date || "").trim();
    if (!isDate(op_date)) return reply.code(400).send({ error: "op_date must be YYYY-MM-DD" });
    const shift_name = String(req.body?.shift_name || "").trim() || null;
    const supervisor_name = String(req.body?.supervisor_name || "").trim() || null;
    const variance_note = String(req.body?.variance_note || "").trim() || null;
    const c1 = req.body?.checklist_weighbridge_reconciled ? 1 : 0;
    const c2 = req.body?.checklist_trucks_reconciled ? 1 : 0;
    const c3 = req.body?.checklist_client_confirmed ? 1 : 0;
    const close_day = req.body?.close_day ? 1 : 0;
    const reopen_day = req.body?.reopen_day ? 1 : 0;
    const reopen_reason = String(req.body?.reopen_reason || "").trim() || null;
    if (reopen_day && !requireRoles(req, reply, ["admin", "supervisor"])) return;
    if (reopen_day && !reopen_reason) {
      return reply.code(400).send({ error: "reopen_reason is required when reopening day" });
    }
    const status = close_day ? "closed" : "open";
    const closed_at = close_day && !reopen_day ? new Date().toISOString() : null;

    db.prepare(`
      INSERT INTO operations_daily_closing (
        op_date, shift_name, supervisor_name, variance_note,
        checklist_weighbridge_reconciled, checklist_trucks_reconciled, checklist_client_confirmed,
        status, closed_at, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      ON CONFLICT(op_date) DO UPDATE SET
        shift_name = excluded.shift_name,
        supervisor_name = excluded.supervisor_name,
        variance_note = excluded.variance_note,
        checklist_weighbridge_reconciled = excluded.checklist_weighbridge_reconciled,
        checklist_trucks_reconciled = excluded.checklist_trucks_reconciled,
        checklist_client_confirmed = excluded.checklist_client_confirmed,
        status = excluded.status,
        closed_at = excluded.closed_at,
        updated_at = datetime('now')
    `).run(
      op_date,
      shift_name,
      supervisor_name,
      variance_note,
      c1,
      c2,
      c3,
      status,
      closed_at
    );

    const row = db.prepare(`
      SELECT
        op_date, shift_name, supervisor_name, variance_note,
        checklist_weighbridge_reconciled, checklist_trucks_reconciled, checklist_client_confirmed,
        status, closed_at, updated_at
      FROM operations_daily_closing
      WHERE op_date = ?
    `).get(op_date);
    if (reopen_day) {
      writeAudit(db, req, {
        module: "operations",
        action: "reopen_day",
        entity_type: "operations_daily_closing",
        entity_id: op_date,
        payload: { reason: reopen_reason, status: row?.status || "open" },
      });
    } else if (close_day) {
      writeAudit(db, req, {
        module: "operations",
        action: "close_day",
        entity_type: "operations_daily_closing",
        entity_id: op_date,
        payload: { supervisor_name, status: row?.status || "closed" },
      });
    } else {
      writeAudit(db, req, {
        module: "operations",
        action: "save_closing_draft",
        entity_type: "operations_daily_closing",
        entity_id: op_date,
        payload: { status: row?.status || "open" },
      });
    }
    return { ok: true, row };
  });

  // GET /api/operations/closing/:date/history
  app.get("/closing/:date/history", async (req, reply) => {
    const op_date = String(req.params?.date || "").trim();
    if (!isDate(op_date)) return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
    const rows = db.prepare(`
      SELECT
        id, module, action, entity_type, entity_id, username, role, payload_json, created_at
      FROM audit_logs
      WHERE module = 'operations'
        AND entity_type = 'operations_daily_closing'
        AND entity_id = ?
      ORDER BY id DESC
      LIMIT 200
    `).all(op_date).map((r) => {
      let payload = null;
      try {
        payload = r.payload_json ? JSON.parse(r.payload_json) : null;
      } catch {
        payload = r.payload_json || null;
      }
      return {
        id: Number(r.id || 0),
        action: String(r.action || ""),
        username: String(r.username || ""),
        role: String(r.role || ""),
        created_at: String(r.created_at || ""),
        payload,
      };
    });
    return { ok: true, rows };
  });
}

