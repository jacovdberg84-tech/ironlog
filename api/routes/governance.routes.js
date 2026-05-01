// IRONLOG/api/routes/governance.routes.js
// Governance v2: segregation of duties policy registry + evaluator + violation log.

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

const DEFAULT_POLICIES = [
  {
    policy_code: "sod_po_create_approve",
    label: "PO creator cannot also approve",
    mode: "block",
    restricted_action: "procurement.po.approve",
    precursor_action: "procurement.po.create",
    window_minutes: 43200,
  },
  {
    policy_code: "sod_po_approve_receive",
    label: "PO approver cannot also receive",
    mode: "warn",
    restricted_action: "procurement.po.receive",
    precursor_action: "procurement.po.approve",
    window_minutes: 43200,
  },
  {
    policy_code: "sod_invoice_capture_post",
    label: "Invoice capturer cannot also post journals",
    mode: "warn",
    restricted_action: "finance.journals.mark_posted",
    precursor_action: "procurement.invoice.capture",
    window_minutes: 43200,
  },
  {
    policy_code: "sod_period_close_reopen",
    label: "Period closer cannot unilaterally reopen",
    mode: "block",
    restricted_action: "finance.period.reopen",
    precursor_action: "finance.period.close",
    window_minutes: 129600,
  },
  {
    policy_code: "sod_run_post_reverse",
    label: "Run poster cannot reverse own run",
    mode: "block",
    restricted_action: "finance.journals.reverse",
    precursor_action: "finance.journals.mark_posted",
    window_minutes: 43200,
  },
];

export default async function governanceRoutes(app) {
  ensureAuditTable(db);

  db.prepare(`
    CREATE TABLE IF NOT EXISTS sod_policies (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      policy_code TEXT NOT NULL UNIQUE,
      label TEXT NOT NULL,
      mode TEXT NOT NULL DEFAULT 'block',
      restricted_action TEXT NOT NULL,
      precursor_action TEXT NOT NULL,
      window_minutes INTEGER NOT NULL DEFAULT 43200,
      scope_json TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS sod_violations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      policy_code TEXT NOT NULL,
      username TEXT NOT NULL,
      restricted_action TEXT NOT NULL,
      precursor_action TEXT NOT NULL,
      entity_type TEXT,
      entity_id TEXT,
      mode TEXT NOT NULL,
      blocked INTEGER NOT NULL DEFAULT 0,
      detail_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      acknowledged_at TEXT,
      acknowledged_by TEXT,
      resolution_notes TEXT
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_sodv_policy ON sod_violations(policy_code)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_sodv_user ON sod_violations(username)`).run();

  // Seed defaults (idempotent) if table is empty.
  const count = db.prepare(`SELECT COUNT(*) AS c FROM sod_policies`).get();
  if (!Number(count?.c || 0)) {
    const ins = db.prepare(`
      INSERT INTO sod_policies (policy_code, label, mode, restricted_action, precursor_action, window_minutes, active, created_by)
      VALUES (?, ?, ?, ?, ?, ?, 1, 'system')
    `);
    for (const p of DEFAULT_POLICIES) ins.run(p.policy_code, p.label, p.mode, p.restricted_action, p.precursor_action, p.window_minutes);
  }

  app.get("/policies", async () => {
    const rows = db.prepare(`
      SELECT id, policy_code, label, mode, restricted_action, precursor_action, window_minutes,
             scope_json, active, created_at, updated_at
      FROM sod_policies ORDER BY policy_code ASC
    `).all();
    return { ok: true, rows };
  });

  app.post("/policies/upsert", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "executive"])) return;
    const policy_code = String(req.body?.policy_code || "").trim().toLowerCase();
    const label = String(req.body?.label || "").trim();
    const mode = String(req.body?.mode || "block").trim().toLowerCase();
    const restricted_action = String(req.body?.restricted_action || "").trim();
    const precursor_action = String(req.body?.precursor_action || "").trim();
    const window_minutes = Math.max(1, Number(req.body?.window_minutes || 43200));
    if (!policy_code || !label || !restricted_action || !precursor_action) return reply.code(400).send({ error: "policy_code, label, restricted_action, precursor_action required" });
    if (!["block", "warn"].includes(mode)) return reply.code(400).send({ error: "mode must be block or warn" });
    db.prepare(`
      INSERT INTO sod_policies (policy_code, label, mode, restricted_action, precursor_action, window_minutes, active, created_by, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, 1, ?, datetime('now'))
      ON CONFLICT(policy_code) DO UPDATE SET
        label = excluded.label,
        mode = excluded.mode,
        restricted_action = excluded.restricted_action,
        precursor_action = excluded.precursor_action,
        window_minutes = excluded.window_minutes,
        active = 1,
        updated_at = datetime('now')
    `).run(policy_code, label, mode, restricted_action, precursor_action, window_minutes, getUser(req));
    writeAudit(db, req, { module: "governance", action: "policies.upsert", entity_type: "sod_policies", entity_id: policy_code, payload: { policy_code, mode } });
    return { ok: true, policy_code };
  });

  function evaluateForUser(action, username, entity) {
    const policies = db.prepare(`
      SELECT policy_code, label, mode, restricted_action, precursor_action, window_minutes
      FROM sod_policies WHERE active = 1 AND LOWER(restricted_action) = LOWER(?)
    `).all(String(action));
    const results = [];
    for (const p of policies) {
      const row = db.prepare(`
        SELECT id, action, entity_type, entity_id, created_at
        FROM audit_logs
        WHERE LOWER(username) = LOWER(?)
          AND LOWER(action) = LOWER(?)
          AND DATE(created_at) >= DATE('now', ?)
        ORDER BY id DESC LIMIT 1
      `).get(username, p.precursor_action, `-${Math.ceil(p.window_minutes / 60 / 24)} day`);
      if (row) {
        results.push({
          violated: true,
          policy: p,
          evidence: row,
        });
      }
    }
    return results;
  }

  function recordViolation(policy, username, restricted_action, entity, blocked) {
    const r = db.prepare(`
      INSERT INTO sod_violations (policy_code, username, restricted_action, precursor_action, entity_type, entity_id, mode, blocked, detail_json)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(
      policy.policy_code, username, restricted_action, policy.precursor_action,
      entity?.entity_type || null, entity?.entity_id != null ? String(entity.entity_id) : null,
      policy.mode, blocked ? 1 : 0, JSON.stringify(entity || {})
    );
    return Number(r.lastInsertRowid);
  }

  app.post("/evaluate", async (req, reply) => {
    const action = String(req.body?.action || "").trim();
    const username = String(req.body?.username || getUser(req)).trim();
    const entity = req.body?.entity || null;
    if (!action) return reply.code(400).send({ error: "action required" });
    const results = evaluateForUser(action, username, entity);
    const violations = [];
    for (const res of results) {
      const blocked = res.policy.mode === "block";
      const id = recordViolation(res.policy, username, action, entity, blocked);
      violations.push({ id, policy_code: res.policy.policy_code, mode: res.policy.mode, blocked, evidence: res.evidence });
    }
    const blockedAny = violations.some((v) => v.blocked);
    return { ok: true, blocked: blockedAny, violations };
  });

  app.get("/violations", async (req) => {
    const ack = String(req.query?.ack || "").trim().toLowerCase();
    let where = "";
    if (ack === "open") where = "WHERE acknowledged_at IS NULL";
    else if (ack === "ack") where = "WHERE acknowledged_at IS NOT NULL";
    const rows = db.prepare(`
      SELECT id, policy_code, username, restricted_action, precursor_action, entity_type, entity_id,
             mode, blocked, detail_json, created_at, acknowledged_at, acknowledged_by, resolution_notes
      FROM sod_violations ${where}
      ORDER BY id DESC LIMIT 500
    `).all();
    return { ok: true, rows };
  });

  app.post("/violations/:id/acknowledge", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "executive"])) return;
    const id = Number(req.params?.id || 0);
    const notes = req.body?.notes ? String(req.body.notes).trim() : null;
    db.prepare(`
      UPDATE sod_violations
      SET acknowledged_at = datetime('now'), acknowledged_by = ?, resolution_notes = ?
      WHERE id = ? AND acknowledged_at IS NULL
    `).run(getUser(req), notes, id);
    writeAudit(db, req, { module: "governance", action: "violations.acknowledge", entity_type: "sod_violations", entity_id: id, payload: { id, notes } });
    return { ok: true, id };
  });
}
