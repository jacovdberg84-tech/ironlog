import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
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

  app.post("/", async (req, reply) => {
    const site_code = getSiteCode(req);
    const op_date = String(req.body?.op_date || "").trim() || new Date().toISOString().slice(0, 10);
    if (!isDate(op_date)) return reply.code(400).send({ error: "op_date must be YYYY-MM-DD" });

    const tonnes_moved = toNumberOrNull(req.body?.tonnes_moved);
    const product_type = String(req.body?.product_type || "").trim() || null;
    const product_produced = toNumberOrNull(req.body?.product_produced);
    const trucks_loaded = toNumberOrNull(req.body?.trucks_loaded);
    const weighbridge_amount = toNumberOrNull(req.body?.weighbridge_amount);
    const trucks_delivered = toNumberOrNull(req.body?.trucks_delivered);
    const product_delivered = toNumberOrNull(req.body?.product_delivered);
    const client_delivered_to = String(req.body?.client_delivered_to || "").trim() || null;
    const notes = String(req.body?.notes || "").trim() || null;

    const ins = db.prepare(`
      INSERT INTO operations_logs (
        op_date, tonnes_moved, product_type, product_produced, trucks_loaded, weighbridge_amount,
        trucks_delivered, product_delivered, client_delivered_to, notes, site_code
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(
      op_date,
      tonnes_moved,
      product_type,
      product_produced,
      trucks_loaded == null ? null : Math.round(trucks_loaded),
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
        id, site_code, op_date, tonnes_moved, product_type, product_produced, trucks_loaded, weighbridge_amount,
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

