import { db } from "../db/client.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function toNum(v) {
  if (v == null || String(v).trim() === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
function getSiteCode(req) {
  return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
}

const VALID_STATUSES = ["queued", "loading", "in_transit", "delivered", "returned"];
const VALID_EXCEPTION_TYPES = ["shortage", "over_delivery", "delay", "rejection", "damage", "other"];
const VALID_EXCEPTION_STATUSES = ["open", "resolved", "waived"];

export default async function dispatchRoutes(app) {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS dispatch_trips (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      op_date TEXT NOT NULL,
      trip_no TEXT,
      truck_reg TEXT NOT NULL,
      driver_name TEXT,
      product_type TEXT,
      client_name TEXT,
      target_tonnes REAL,
      actual_tonnes REAL,
      status TEXT NOT NULL DEFAULT 'queued',
      queued_at TEXT,
      loading_at TEXT,
      in_transit_at TEXT,
      delivered_at TEXT,
      returned_at TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  const cols = db.prepare(`PRAGMA table_info(dispatch_trips)`).all();
  const hasCol = (c) => cols.some((r) => String(r.name || "") === c);
  if (!hasCol("pod_ref")) db.prepare(`ALTER TABLE dispatch_trips ADD COLUMN pod_ref TEXT`).run();
  if (!hasCol("pod_link")) db.prepare(`ALTER TABLE dispatch_trips ADD COLUMN pod_link TEXT`).run();
  if (!hasCol("pod_captured_at")) db.prepare(`ALTER TABLE dispatch_trips ADD COLUMN pod_captured_at TEXT`).run();
  if (!hasCol("site_code")) db.prepare(`ALTER TABLE dispatch_trips ADD COLUMN site_code TEXT`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS dispatch_exceptions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      trip_id INTEGER NOT NULL,
      exception_type TEXT NOT NULL,
      severity TEXT NOT NULL DEFAULT 'medium',
      status TEXT NOT NULL DEFAULT 'open',
      owner_name TEXT,
      note TEXT,
      resolution_note TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      resolved_at TEXT,
      FOREIGN KEY (trip_id) REFERENCES dispatch_trips(id) ON DELETE CASCADE
    )
  `).run();

  app.post("/trips", async (req, reply) => {
    const site_code = getSiteCode(req);
    const op_date = String(req.body?.op_date || "").trim() || new Date().toISOString().slice(0, 10);
    if (!isDate(op_date)) return reply.code(400).send({ error: "op_date must be YYYY-MM-DD" });
    const truck_reg = String(req.body?.truck_reg || "").trim();
    if (!truck_reg) return reply.code(400).send({ error: "truck_reg is required" });
    const product_type = String(req.body?.product_type || "").trim() || null;
    const client_name = String(req.body?.client_name || "").trim() || null;
    const trip_no = String(req.body?.trip_no || "").trim() || null;
    const driver_name = String(req.body?.driver_name || "").trim() || null;
    const target_tonnes = toNum(req.body?.target_tonnes);
    const actual_tonnes = toNum(req.body?.actual_tonnes);
    const notes = String(req.body?.notes || "").trim() || null;

    const ins = db.prepare(`
      INSERT INTO dispatch_trips (
        op_date, trip_no, truck_reg, driver_name, product_type, client_name, target_tonnes, actual_tonnes,
        status, queued_at, notes, site_code, created_at, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'queued', datetime('now'), ?, ?, datetime('now'), datetime('now'))
    `).run(op_date, trip_no, truck_reg, driver_name, product_type, client_name, target_tonnes, actual_tonnes, notes, site_code);

    return { ok: true, id: Number(ins.lastInsertRowid) };
  });

  app.get("/trips", async (req, reply) => {
    const site_code = getSiteCode(req);
    const from = String(req.query?.from || "").trim();
    const to = String(req.query?.to || "").trim();
    const status = String(req.query?.status || "").trim().toLowerCase();
    if (from && !isDate(from)) return reply.code(400).send({ error: "from must be YYYY-MM-DD" });
    if (to && !isDate(to)) return reply.code(400).send({ error: "to must be YYYY-MM-DD" });
    if (status && !VALID_STATUSES.includes(status)) return reply.code(400).send({ error: "invalid status filter" });

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
    if (status) {
      where.push("status = ?");
      params.push(status);
    }
    where.push("COALESCE(site_code,'main') = ?");
    params.push(site_code);

    const rows = db.prepare(`
      SELECT
        id, op_date, trip_no, truck_reg, driver_name, product_type, client_name,
        target_tonnes, actual_tonnes, status, queued_at, loading_at, in_transit_at, delivered_at, returned_at, notes,
        site_code,
        pod_ref, pod_link, pod_captured_at,
        created_at, updated_at
      FROM dispatch_trips
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY op_date DESC, id DESC
      LIMIT 1000
    `).all(...params);
    return { ok: true, rows };
  });

  app.post("/trips/:id/status", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "valid trip id required" });
    const next = String(req.body?.status || "").trim().toLowerCase();
    if (!VALID_STATUSES.includes(next)) return reply.code(400).send({ error: "invalid status" });
    const row = db.prepare(`SELECT id, pod_ref FROM dispatch_trips WHERE id = ?`).get(id);
    if (!row) return reply.code(404).send({ error: "trip not found" });
    const requirePodForDelivered = String(req.body?.require_pod_for_delivered ?? "1") !== "0";
    if (next === "delivered" && requirePodForDelivered && !String(row.pod_ref || "").trim()) {
      return reply.code(400).send({ error: "POD ref required before marking delivered" });
    }

    const setCols = ["status = ?", "updated_at = datetime('now')"];
    const params = [next];
    if (next === "loading") setCols.push("loading_at = COALESCE(loading_at, datetime('now'))");
    if (next === "in_transit") setCols.push("in_transit_at = COALESCE(in_transit_at, datetime('now'))");
    if (next === "delivered") setCols.push("delivered_at = COALESCE(delivered_at, datetime('now'))");
    if (next === "returned") setCols.push("returned_at = COALESCE(returned_at, datetime('now'))");
    params.push(id);

    db.prepare(`UPDATE dispatch_trips SET ${setCols.join(", ")} WHERE id = ?`).run(...params);
    const updated = db.prepare(`SELECT * FROM dispatch_trips WHERE id = ?`).get(id);
    return { ok: true, row: updated };
  });

  app.post("/trips/:id/pod", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "valid trip id required" });
    const pod_ref = String(req.body?.pod_ref || "").trim();
    if (!pod_ref) return reply.code(400).send({ error: "pod_ref is required" });
    const pod_link = String(req.body?.pod_link || "").trim() || null;
    const row = db.prepare(`SELECT id FROM dispatch_trips WHERE id = ?`).get(id);
    if (!row) return reply.code(404).send({ error: "trip not found" });
    db.prepare(`
      UPDATE dispatch_trips
      SET pod_ref = ?, pod_link = ?, pod_captured_at = datetime('now'), updated_at = datetime('now')
      WHERE id = ?
    `).run(pod_ref, pod_link, id);
    const updated = db.prepare(`SELECT * FROM dispatch_trips WHERE id = ?`).get(id);
    return { ok: true, row: updated };
  });

  app.post("/exceptions", async (req, reply) => {
    const trip_id = Number(req.body?.trip_id || 0);
    if (!Number.isFinite(trip_id) || trip_id <= 0) return reply.code(400).send({ error: "trip_id is required" });
    const exception_type = String(req.body?.exception_type || "").trim().toLowerCase();
    if (!VALID_EXCEPTION_TYPES.includes(exception_type)) {
      return reply.code(400).send({ error: "invalid exception_type" });
    }
    const severity = String(req.body?.severity || "medium").trim().toLowerCase();
    const owner_name = String(req.body?.owner_name || "").trim() || null;
    const note = String(req.body?.note || "").trim() || null;
    const trip = db.prepare(`SELECT id FROM dispatch_trips WHERE id = ?`).get(trip_id);
    if (!trip) return reply.code(404).send({ error: "trip not found" });
    const ins = db.prepare(`
      INSERT INTO dispatch_exceptions (trip_id, exception_type, severity, status, owner_name, note)
      VALUES (?, ?, ?, 'open', ?, ?)
    `).run(trip_id, exception_type, severity, owner_name, note);
    return { ok: true, id: Number(ins.lastInsertRowid) };
  });

  app.get("/exceptions", async (req) => {
    const site_code = getSiteCode(req);
    const from = String(req.query?.from || "").trim();
    const to = String(req.query?.to || "").trim();
    const status = String(req.query?.status || "").trim().toLowerCase();
    const where = [];
    const params = [];
    if (from) {
      where.push("t.op_date >= ?");
      params.push(from);
    }
    if (to) {
      where.push("t.op_date <= ?");
      params.push(to);
    }
    if (status && VALID_EXCEPTION_STATUSES.includes(status)) {
      where.push("e.status = ?");
      params.push(status);
    }
    where.push("COALESCE(t.site_code,'main') = ?");
    params.push(site_code);
    const rows = db.prepare(`
      SELECT
        e.id, e.trip_id, e.exception_type, e.severity, e.status, e.owner_name, e.note, e.resolution_note, e.created_at, e.resolved_at,
        t.op_date, t.trip_no, t.truck_reg, t.client_name, t.product_type
      FROM dispatch_exceptions e
      JOIN dispatch_trips t ON t.id = e.trip_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY e.id DESC
      LIMIT 500
    `).all(...params);
    return { ok: true, rows };
  });

  app.post("/exceptions/:id/resolve", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "valid exception id required" });
    const status = String(req.body?.status || "resolved").trim().toLowerCase();
    if (!["resolved", "waived"].includes(status)) return reply.code(400).send({ error: "status must be resolved or waived" });
    const resolution_note = String(req.body?.resolution_note || "").trim() || null;
    const row = db.prepare(`SELECT id FROM dispatch_exceptions WHERE id = ?`).get(id);
    if (!row) return reply.code(404).send({ error: "exception not found" });
    db.prepare(`
      UPDATE dispatch_exceptions
      SET status = ?, resolution_note = ?, resolved_at = datetime('now')
      WHERE id = ?
    `).run(status, resolution_note, id);
    const updated = db.prepare(`SELECT * FROM dispatch_exceptions WHERE id = ?`).get(id);
    return { ok: true, row: updated };
  });

  app.get("/kpi", async (req, reply) => {
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
    const whereSql = where.length ? `WHERE ${where.join(" AND ")}` : "";

    const total = db.prepare(`SELECT COUNT(*) AS n FROM dispatch_trips ${whereSql}`).get(...params);
    const byStatus = db.prepare(`
      SELECT status, COUNT(*) AS n
      FROM dispatch_trips
      ${whereSql}
      GROUP BY status
    `).all(...params);
    const deliveredTonnes = db.prepare(`
      SELECT IFNULL(SUM(actual_tonnes),0) AS t
      FROM dispatch_trips
      ${whereSql} ${whereSql ? "AND" : "WHERE"} status = 'delivered'
    `).get(...params);
    const avgTurn = db.prepare(`
      SELECT AVG((julianday(delivered_at) - julianday(queued_at)) * 24.0) AS hours
      FROM dispatch_trips
      ${whereSql} ${whereSql ? "AND" : "WHERE"} delivered_at IS NOT NULL AND queued_at IS NOT NULL
    `).get(...params);
    const deliveredTrips = db.prepare(`
      SELECT COUNT(*) AS n
      FROM dispatch_trips
      ${whereSql} ${whereSql ? "AND" : "WHERE"} status = 'delivered'
    `).get(...params);
    const deliveredWithPod = db.prepare(`
      SELECT COUNT(*) AS n
      FROM dispatch_trips
      ${whereSql} ${whereSql ? "AND" : "WHERE"} status = 'delivered' AND TRIM(COALESCE(pod_ref, '')) <> ''
    `).get(...params);
    const exceptionsOpen = db.prepare(`
      SELECT COUNT(*) AS n
      FROM dispatch_exceptions e
      JOIN dispatch_trips t ON t.id = e.trip_id
      ${where.length ? `WHERE ${where.join(" AND ")} AND e.status = 'open'` : "WHERE e.status = 'open'"}
    `).get(...params);

    const map = Object.fromEntries(VALID_STATUSES.map((s) => [s, 0]));
    byStatus.forEach((r) => {
      map[String(r.status || "")] = Number(r.n || 0);
    });

    return {
      ok: true,
      total_trips: Number(total?.n || 0),
      by_status: map,
      delivered_tonnes: Number(deliveredTonnes?.t || 0),
      avg_turnaround_hours: Number(avgTurn?.hours || 0),
      delivered_with_pod_pct: Number(deliveredTrips?.n || 0) > 0
        ? Number(((Number(deliveredWithPod?.n || 0) / Number(deliveredTrips?.n || 0)) * 100).toFixed(2))
        : null,
      exceptions_open: Number(exceptionsOpen?.n || 0),
    };
  });
}

