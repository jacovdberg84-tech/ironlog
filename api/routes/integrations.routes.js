// IRONLOG/api/routes/integrations.routes.js
// Integration Hub v1: formal API/ETL layer with ERP + Payroll connectors,
// sync jobs, retry queue, dead-letter, and failure monitoring.

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
  return Array.from(new Set([...many, ...one]));
}
function hasAnyRole(req, allowed) {
  return getRoles(req).some((r) => allowed.includes(r));
}
function requireRoles(req, reply, allowed) {
  if (!hasAnyRole(req, allowed)) {
    reply.code(403).send({ error: `role '${getRole(req)}' not allowed` });
    return false;
  }
  return true;
}
function normalizeCode(s) { return String(s || "").trim().toUpperCase().replace(/\s+/g, "-"); }

const CONNECTOR_REGISTRY = {
  erp_journal_export: {
    kind: "erp",
    direction: "outbound",
    label: "ERP Journal Export",
    payload_hint: { run_id: "finance_posting_runs.id" },
  },
  payroll_labor_sync: {
    kind: "payroll",
    direction: "inbound",
    label: "Payroll Labor Sync",
    payload_hint: { period: "YYYY-MM" },
  },
};

const JOB_STATUS = {
  QUEUED: "queued",
  RUNNING: "running",
  RETRY_WAIT: "retry_wait",
  SUCCEEDED: "succeeded",
  FAILED_PERMANENT: "failed_permanent",
  CANCELLED: "cancelled",
};

const DEFAULT_MAX_ATTEMPTS = 5;
const BACKOFF_BASE_MS = 30_000;

function nextBackoffMs(attempt) {
  const base = BACKOFF_BASE_MS;
  const factor = Math.pow(2, Math.max(0, attempt - 1));
  const jitter = Math.floor(Math.random() * 1000);
  return Math.min(60 * 60 * 1000, base * factor + jitter);
}

export default async function integrationsRoutes(app) {
  ensureAuditTable(db);

  db.prepare(`
    CREATE TABLE IF NOT EXISTS integration_connections (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      connection_code TEXT NOT NULL UNIQUE,
      connector_key TEXT NOT NULL,
      label TEXT NOT NULL,
      config_json TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS integration_jobs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      connection_code TEXT NOT NULL,
      connector_key TEXT NOT NULL,
      payload_json TEXT,
      status TEXT NOT NULL DEFAULT 'queued',
      attempts INTEGER NOT NULL DEFAULT 0,
      max_attempts INTEGER NOT NULL DEFAULT 5,
      idempotency_key TEXT,
      external_ref TEXT,
      error_message TEXT,
      error_code TEXT,
      scheduled_at TEXT NOT NULL DEFAULT (datetime('now')),
      started_at TEXT,
      finished_at TEXT,
      next_attempt_at TEXT,
      dead_letter_at TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_int_jobs_status ON integration_jobs(status)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_int_jobs_conn ON integration_jobs(connector_key, status)`).run();
  db.prepare(`CREATE UNIQUE INDEX IF NOT EXISTS idx_int_jobs_idem ON integration_jobs(connector_key, idempotency_key) WHERE idempotency_key IS NOT NULL`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS integration_job_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      job_id INTEGER NOT NULL,
      event_type TEXT NOT NULL,
      message TEXT,
      meta_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (job_id) REFERENCES integration_jobs(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_int_events_job ON integration_job_events(job_id)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS integration_dead_letter (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      job_id INTEGER NOT NULL,
      connector_key TEXT NOT NULL,
      reason TEXT,
      payload_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      acknowledged_at TEXT,
      acknowledged_by TEXT
    )
  `).run();

  function logEvent(jobId, eventType, message, meta) {
    db.prepare(`
      INSERT INTO integration_job_events (job_id, event_type, message, meta_json)
      VALUES (?, ?, ?, ?)
    `).run(jobId, String(eventType || "event"), message ? String(message) : null, meta ? JSON.stringify(meta) : null);
  }

  /* --------- CONNECTOR REGISTRY + CONNECTIONS --------- */

  app.get("/connectors", async () => {
    return { ok: true, connectors: Object.entries(CONNECTOR_REGISTRY).map(([k, v]) => ({ key: k, ...v })) };
  });

  app.post("/connections/upsert", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "executive"])) return;
    const connection_code = normalizeCode(req.body?.connection_code);
    const connector_key = String(req.body?.connector_key || "").trim();
    const label = String(req.body?.label || "").trim();
    if (!connection_code || !connector_key || !label) return reply.code(400).send({ error: "connection_code, connector_key, label required" });
    if (!CONNECTOR_REGISTRY[connector_key]) return reply.code(400).send({ error: `unknown connector_key '${connector_key}'` });
    const config = req.body?.config != null ? JSON.stringify(req.body.config) : null;
    db.prepare(`
      INSERT INTO integration_connections (connection_code, connector_key, label, config_json, active, created_by, updated_at)
      VALUES (?, ?, ?, ?, 1, ?, datetime('now'))
      ON CONFLICT(connection_code) DO UPDATE SET
        connector_key = excluded.connector_key,
        label = excluded.label,
        config_json = excluded.config_json,
        active = 1,
        updated_at = datetime('now')
    `).run(connection_code, connector_key, label, config, getUser(req));
    writeAudit(db, req, { module: "integrations", action: "connections.upsert", entity_type: "integration_connections", entity_id: connection_code, payload: { connection_code, connector_key } });
    return { ok: true, connection_code };
  });

  app.get("/connections", async () => {
    const rows = db.prepare(`
      SELECT id, connection_code, connector_key, label, config_json, active, created_at, updated_at
      FROM integration_connections
      ORDER BY label ASC
      LIMIT 500
    `).all().map((r) => ({
      ...r,
      config: (() => { try { return JSON.parse(r.config_json || "{}"); } catch { return {}; } })()
    }));
    return { ok: true, rows };
  });

  /* --------- ENQUEUE + RUN JOBS --------- */

  app.post("/jobs/enqueue", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "executive", "finance", "procurement"])) return;
    const connection_code = normalizeCode(req.body?.connection_code);
    const connector_key = String(req.body?.connector_key || "").trim();
    if (!connection_code || !connector_key) return reply.code(400).send({ error: "connection_code and connector_key required" });
    const conn = db.prepare(`SELECT connection_code, active FROM integration_connections WHERE connection_code = ?`).get(connection_code);
    if (!conn) return reply.code(404).send({ error: "connection not found" });
    if (!conn.active) return reply.code(409).send({ error: "connection inactive" });
    const idempotency_key = req.body?.idempotency_key ? String(req.body.idempotency_key).trim() : null;
    if (idempotency_key) {
      const existing = db.prepare(`
        SELECT id, status FROM integration_jobs WHERE connector_key = ? AND idempotency_key = ?
      `).get(connector_key, idempotency_key);
      if (existing) return { ok: true, duplicate: true, id: existing.id, status: existing.status };
    }
    const payload_json = req.body?.payload != null ? JSON.stringify(req.body.payload) : null;
    const max_attempts = Math.max(1, Math.min(10, Number(req.body?.max_attempts || DEFAULT_MAX_ATTEMPTS)));
    const r = db.prepare(`
      INSERT INTO integration_jobs
        (connection_code, connector_key, payload_json, status, max_attempts, idempotency_key, created_by)
      VALUES (?, ?, ?, 'queued', ?, ?, ?)
    `).run(connection_code, connector_key, payload_json, max_attempts, idempotency_key, getUser(req));
    const jobId = Number(r.lastInsertRowid);
    logEvent(jobId, "enqueue", "job queued", { connection_code, connector_key });
    writeAudit(db, req, { module: "integrations", action: "jobs.enqueue", entity_type: "integration_jobs", entity_id: jobId, payload: { connection_code, connector_key } });
    return { ok: true, id: jobId, status: "queued" };
  });

  function runConnector(connectorKey, payload, connection) {
    // Phase 1 connectors use export-style handoff: mark succeeded after validating payload.
    if (connectorKey === "erp_journal_export") {
      const runId = Number(payload?.run_id || 0);
      if (!runId) throw new Error("run_id required for erp_journal_export");
      const run = db.prepare(`SELECT id, run_number, status FROM finance_posting_runs WHERE id = ?`).get(runId);
      if (!run) throw new Error(`finance_posting_runs id ${runId} not found`);
      const expected = ["exported", "posted"];
      if (!expected.includes(String(run.status).toLowerCase())) {
        throw new Error(`run ${runId} status=${run.status}; expected one of ${expected.join(",")}`);
      }
      return { ok: true, exported_run_number: run.run_number, run_id: runId };
    }
    if (connectorKey === "payroll_labor_sync") {
      const period = String(payload?.period || "").trim();
      if (!/^\d{4}-\d{2}$/.test(period)) throw new Error("period (YYYY-MM) required for payroll_labor_sync");
      const row = db.prepare(`
        SELECT COALESCE(SUM(COALESCE(labor_hours, 0)), 0) AS hours
        FROM work_orders
        WHERE status IN ('completed','approved','closed')
          AND DATE(COALESCE(completed_at, closed_at)) BETWEEN DATE(? || '-01') AND DATE(? || '-01','+1 month','-1 day')
      `).get(period, period);
      return { ok: true, period, labor_hours: Number(row?.hours || 0) };
    }
    throw new Error(`no runner for connector ${connectorKey}`);
  }

  function claimJob(jobId) {
    const job = db.prepare(`
      SELECT * FROM integration_jobs WHERE id = ? AND status IN ('queued', 'retry_wait')
    `).get(jobId);
    if (!job) return null;
    db.prepare(`
      UPDATE integration_jobs
      SET status = 'running', started_at = datetime('now'), attempts = attempts + 1, updated_at = datetime('now')
      WHERE id = ?
    `).run(jobId);
    return { ...job, attempts: job.attempts + 1 };
  }

  app.post("/jobs/:id/run", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "executive", "finance", "procurement"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const job = claimJob(id);
    if (!job) return reply.code(409).send({ error: "job not in queued/retry state" });
    logEvent(id, "start", `attempt ${job.attempts}`, { attempts: job.attempts });
    let payload = {};
    try { payload = job.payload_json ? JSON.parse(job.payload_json) : {}; } catch {}
    const conn = db.prepare(`SELECT * FROM integration_connections WHERE connection_code = ?`).get(job.connection_code);
    try {
      const result = runConnector(job.connector_key, payload, conn);
      db.prepare(`
        UPDATE integration_jobs
        SET status = 'succeeded', finished_at = datetime('now'), error_message = NULL, error_code = NULL,
            external_ref = ?, updated_at = datetime('now')
        WHERE id = ?
      `).run(result?.exported_run_number || result?.period || null, id);
      logEvent(id, "success", "completed", result);
      writeAudit(db, req, { module: "integrations", action: "jobs.run.success", entity_type: "integration_jobs", entity_id: id, payload: { connector_key: job.connector_key, result } });
      return { ok: true, id, status: "succeeded", result };
    } catch (e) {
      const msg = e?.message || "error";
      const canRetry = job.attempts < job.max_attempts;
      if (canRetry) {
        const next = new Date(Date.now() + nextBackoffMs(job.attempts)).toISOString();
        db.prepare(`
          UPDATE integration_jobs
          SET status = 'retry_wait', error_message = ?, error_code = 'runtime',
              next_attempt_at = ?, updated_at = datetime('now'), finished_at = NULL
          WHERE id = ?
        `).run(msg, next, id);
        logEvent(id, "retry", msg, { next_attempt_at: next, attempts: job.attempts });
      } else {
        db.prepare(`
          UPDATE integration_jobs
          SET status = 'failed_permanent', error_message = ?, error_code = 'runtime',
              dead_letter_at = datetime('now'), updated_at = datetime('now'), finished_at = datetime('now')
          WHERE id = ?
        `).run(msg, id);
        db.prepare(`
          INSERT INTO integration_dead_letter (job_id, connector_key, reason, payload_json)
          VALUES (?, ?, ?, ?)
        `).run(id, job.connector_key, msg, job.payload_json);
        logEvent(id, "dead_letter", msg, {});
      }
      writeAudit(db, req, { module: "integrations", action: "jobs.run.error", entity_type: "integration_jobs", entity_id: id, payload: { connector_key: job.connector_key, error: msg, attempt: job.attempts } });
      return reply.code(500).send({ ok: false, id, error: msg, retriable: canRetry });
    }
  });

  app.post("/jobs/:id/cancel", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    const reason = req.body?.reason ? String(req.body.reason).trim() : "cancelled";
    db.prepare(`
      UPDATE integration_jobs
      SET status = 'cancelled', finished_at = datetime('now'), updated_at = datetime('now'), error_message = ?
      WHERE id = ? AND status IN ('queued', 'retry_wait')
    `).run(reason, id);
    logEvent(id, "cancel", reason, {});
    return { ok: true, id, status: "cancelled" };
  });

  app.post("/jobs/:id/retry-now", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "finance", "procurement"])) return;
    const id = Number(req.params?.id || 0);
    const job = db.prepare(`SELECT id, status, attempts, max_attempts FROM integration_jobs WHERE id = ?`).get(id);
    if (!job) return reply.code(404).send({ error: "job not found" });
    if (!["retry_wait", "failed_permanent"].includes(String(job.status))) {
      return reply.code(409).send({ error: `cannot retry status=${job.status}` });
    }
    db.prepare(`
      UPDATE integration_jobs
      SET status = 'queued', next_attempt_at = NULL, dead_letter_at = NULL,
          max_attempts = CASE WHEN ? THEN attempts + 3 ELSE max_attempts END,
          updated_at = datetime('now')
      WHERE id = ?
    `).run(String(job.status) === "failed_permanent" ? 1 : 0, id);
    logEvent(id, "manual_retry", "user retry", { by: getUser(req) });
    return { ok: true, id, status: "queued" };
  });

  app.get("/jobs", async (req) => {
    const status = req.query?.status ? String(req.query.status).trim().toLowerCase() : null;
    const connector = req.query?.connector ? String(req.query.connector).trim() : null;
    const where = [];
    const args = [];
    if (status) { where.push("LOWER(status) = ?"); args.push(status); }
    if (connector) { where.push("connector_key = ?"); args.push(connector); }
    const sql = `
      SELECT id, connection_code, connector_key, status, attempts, max_attempts, idempotency_key,
             external_ref, error_message, scheduled_at, started_at, finished_at, next_attempt_at,
             dead_letter_at, created_by, created_at, updated_at
      FROM integration_jobs
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY id DESC LIMIT 500
    `;
    const rows = db.prepare(sql).all(...args);
    return { ok: true, rows };
  });

  app.get("/jobs/:id", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const job = db.prepare(`SELECT * FROM integration_jobs WHERE id = ?`).get(id);
    if (!job) return reply.code(404).send({ error: "job not found" });
    const events = db.prepare(`
      SELECT id, event_type, message, meta_json, created_at
      FROM integration_job_events WHERE job_id = ? ORDER BY id ASC
    `).all(id).map((e) => ({
      ...e,
      meta: (() => { try { return JSON.parse(e.meta_json || "{}"); } catch { return {}; } })()
    }));
    return { ok: true, job, events };
  });

  app.get("/dead-letter", async () => {
    const rows = db.prepare(`
      SELECT id, job_id, connector_key, reason, created_at, acknowledged_at, acknowledged_by
      FROM integration_dead_letter
      ORDER BY id DESC LIMIT 500
    `).all();
    return { ok: true, rows };
  });

  app.post("/dead-letter/:id/acknowledge", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    db.prepare(`
      UPDATE integration_dead_letter
      SET acknowledged_at = datetime('now'), acknowledged_by = ?
      WHERE id = ? AND acknowledged_at IS NULL
    `).run(getUser(req), id);
    return { ok: true, id };
  });

  /* --------- MONITORING --------- */

  app.get("/monitoring/summary", async () => {
    const byStatus = db.prepare(`
      SELECT status, COUNT(*) AS c FROM integration_jobs GROUP BY status
    `).all();
    const byConnector = db.prepare(`
      SELECT connector_key,
             SUM(CASE WHEN status = 'succeeded' THEN 1 ELSE 0 END) AS succeeded,
             SUM(CASE WHEN status = 'failed_permanent' THEN 1 ELSE 0 END) AS failed,
             SUM(CASE WHEN status = 'retry_wait' THEN 1 ELSE 0 END) AS retry_wait,
             SUM(CASE WHEN status = 'queued' THEN 1 ELSE 0 END) AS queued,
             COUNT(*) AS total
      FROM integration_jobs GROUP BY connector_key
    `).all();
    const oldestQueued = db.prepare(`
      SELECT id, connector_key, scheduled_at FROM integration_jobs
      WHERE status = 'queued' ORDER BY scheduled_at ASC LIMIT 1
    `).get();
    const topErrors = db.prepare(`
      SELECT error_message, COUNT(*) AS c FROM integration_jobs
      WHERE error_message IS NOT NULL AND status IN ('retry_wait','failed_permanent')
      GROUP BY error_message ORDER BY c DESC LIMIT 10
    `).all();
    const deadLetterOpen = db.prepare(`
      SELECT COUNT(*) AS c FROM integration_dead_letter WHERE acknowledged_at IS NULL
    `).get();
    return {
      ok: true,
      by_status: byStatus,
      by_connector: byConnector,
      oldest_queued: oldestQueued || null,
      top_errors: topErrors,
      dead_letter_open: Number(deadLetterOpen?.c || 0),
    };
  });
}
