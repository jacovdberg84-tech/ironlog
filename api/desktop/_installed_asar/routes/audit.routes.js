import { db } from "../db/client.js";
import { ensureAuditTable } from "../utils/audit.js";

export default async function auditRoutes(app) {
  ensureAuditTable(db);

  function getRole(req) {
    return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
  }

  function requireAuditRead(req, reply) {
    const role = getRole(req);
    if (!["admin", "supervisor"].includes(role)) {
      reply.code(403).send({ error: `role '${role || "unknown"}' not allowed` });
      return false;
    }
    return true;
  }

  // GET /api/audit?module=&action=&entity_type=&username=&limit=
  app.get("/", async (req, reply) => {
    if (!requireAuditRead(req, reply)) return;

    const module = String(req.query?.module || "").trim();
    const action = String(req.query?.action || "").trim();
    const entity_type = String(req.query?.entity_type || "").trim();
    const username = String(req.query?.username || "").trim();
    const limitInput = Number(req.query?.limit ?? 200);
    const limit = Number.isFinite(limitInput) ? Math.max(1, Math.min(1000, Math.trunc(limitInput))) : 200;

    const where = [];
    const params = [];
    if (module) {
      where.push("module = ?");
      params.push(module);
    }
    if (action) {
      where.push("action = ?");
      params.push(action);
    }
    if (entity_type) {
      where.push("entity_type = ?");
      params.push(entity_type);
    }
    if (username) {
      where.push("username = ?");
      params.push(username);
    }

    const rows = db.prepare(`
      SELECT
        id, module, action, entity_type, entity_id, username, role, payload_json, created_at
      FROM audit_logs
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY id DESC
      LIMIT ${limit}
    `).all(...params).map((r) => {
      let payload = null;
      try {
        payload = r.payload_json ? JSON.parse(r.payload_json) : null;
      } catch {
        payload = r.payload_json || null;
      }
      return {
        id: Number(r.id),
        module: r.module,
        action: r.action,
        entity_type: r.entity_type,
        entity_id: r.entity_id,
        username: r.username,
        role: r.role,
        payload,
        created_at: r.created_at,
      };
    });

    return reply.send({ ok: true, rows });
  });
}
