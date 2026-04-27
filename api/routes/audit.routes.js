import { db } from "../db/client.js";
import { ensureAuditTable } from "../utils/audit.js";

export default async function auditRoutes(app) {
  ensureAuditTable(db);

  function getRoles(req) {
    const fromMany = String(req.headers["x-user-roles"] || "")
      .split(",")
      .map((x) => String(x || "").trim().toLowerCase())
      .filter(Boolean);
    const fromSingle = String(req.headers["x-user-role"] || "")
      .split(",")
      .map((x) => String(x || "").trim().toLowerCase())
      .filter(Boolean);
    const merged = Array.from(new Set([...fromMany, ...fromSingle]));
    return merged.length ? merged : ["admin"];
  }
  function getRole(req) {
    return getRoles(req)[0] || "admin";
  }
  function canCrossSite(req) {
    const roles = getRoles(req);
    return roles.includes("admin") || roles.includes("supervisor");
  }

  function requireAuditRead(req, reply) {
    const roles = getRoles(req);
    if (!roles.includes("admin") && !roles.includes("supervisor")) {
      reply.code(403).send({ error: `role '${getRole(req) || "unknown"}' not allowed` });
      return false;
    }
    return true;
  }

  function parseJsonMaybe(raw) {
    if (!raw) return null;
    try { return JSON.parse(String(raw)); } catch { return raw; }
  }
  function normalizeAuditRow(r) {
    const payload = parseJsonMaybe(r.payload_json);
    const before = parseJsonMaybe(r.before_json);
    const after = parseJsonMaybe(r.after_json);
    return {
      id: Number(r.id),
      module: r.module,
      action: r.action,
      entity_type: r.entity_type,
      entity_id: r.entity_id,
      username: r.username,
      role: r.role,
      site_code: r.site_code,
      source_app: r.source_app || "web",
      source_channel: r.source_channel || "api",
      request_id: r.request_id || null,
      ip_address: r.ip_address || null,
      user_agent: r.user_agent || null,
      payload,
      before,
      after,
      created_at: r.created_at,
    };
  }

  function queryTimeline(req) {
    const module = String(req.query?.module || "").trim();
    const action = String(req.query?.action || "").trim();
    const entity_type = String(req.query?.entity_type || "").trim();
    const entity_id = String(req.query?.entity_id || "").trim();
    const username = String(req.query?.username || "").trim();
    const source_app = String(req.query?.source_app || "").trim().toLowerCase();
    const site_code = String(req.query?.site_code || "").trim().toLowerCase();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const cursor = Number(req.query?.cursor || 0);
    const limitInput = Number(req.query?.limit ?? 100);
    const limit = Number.isFinite(limitInput) ? Math.max(1, Math.min(500, Math.trunc(limitInput))) : 100;

    const where = [];
    const params = [];
    if (cursor > 0) {
      where.push("id < ?");
      params.push(cursor);
    }
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
    if (entity_id) {
      where.push("entity_id = ?");
      params.push(entity_id);
    }
    if (username) {
      where.push("username = ?");
      params.push(username);
    }
    if (source_app) {
      where.push("LOWER(COALESCE(source_app, '')) = ?");
      params.push(source_app);
    }
    if (start) {
      where.push("datetime(created_at) >= datetime(?)");
      params.push(start);
    }
    if (end) {
      where.push("datetime(created_at) <= datetime(?)");
      params.push(end);
    }
    if (site_code && canCrossSite(req)) {
      where.push("LOWER(COALESCE(site_code, '')) = ?");
      params.push(site_code);
    } else {
      const reqSite = String(req.headers["x-site-code"] || "").trim().toLowerCase();
      if (reqSite) {
        where.push("(LOWER(COALESCE(site_code, '')) = ? OR site_code IS NULL)");
        params.push(reqSite);
      }
    }

    const rows = db.prepare(`
      SELECT
        id, module, action, entity_type, entity_id, username, role, site_code, source_app, source_channel,
        request_id, ip_address, user_agent, before_json, after_json, payload_json, created_at
      FROM audit_logs
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY id DESC
      LIMIT ${limit + 1}
    `).all(...params).map(normalizeAuditRow);

    const has_more = rows.length > limit;
    const sliced = has_more ? rows.slice(0, limit) : rows;
    const next_cursor = has_more ? Number(sliced[sliced.length - 1]?.id || 0) : null;
    return { rows: sliced, pagination: { limit, has_more, next_cursor } };
  }

  // GET /api/audit/timeline
  app.get("/timeline", async (req, reply) => {
    if (!requireAuditRead(req, reply)) return;
    const out = queryTimeline(req);
    return reply.send({ ok: true, ...out });
  });

  app.get("/timeline/:id", async (req, reply) => {
    if (!requireAuditRead(req, reply)) return;
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
    const r = db.prepare(`
      SELECT
        id, module, action, entity_type, entity_id, username, role, site_code, source_app, source_channel,
        request_id, ip_address, user_agent, before_json, after_json, payload_json, created_at
      FROM audit_logs
      WHERE id = ?
      LIMIT 1
    `).get(id);
    if (!r) return reply.code(404).send({ ok: false, error: "not found" });
    return reply.send({ ok: true, row: normalizeAuditRow(r) });
  });

  // Backward-compatible endpoint.
  app.get("/", async (req, reply) => {
    if (!requireAuditRead(req, reply)) return;
    const out = queryTimeline(req);
    return reply.send({ ok: true, rows: out.rows });
  });
}
