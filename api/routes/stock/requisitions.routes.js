// IRONLOG/api/routes/stock/requisitions.routes.js — printable stores requisitions.
// Registered by routes/stock.routes.js under /api/stock.
import { db } from "../../db/client.js";
import { writeAudit } from "../../utils/audit.js";
import {
  buildRequisitionPdf,
  ensureRequisitionSchema,
  getRequisition,
  listRequisitions,
  requisitionFor,
} from "../../utils/storesRequisition.js";

const STORES_ROLES = ["admin", "supervisor", "stores", "storeman"];

export default function registerRequisitionRoutes(app, ctx) {
  const { requireRoles, getSiteCode } = ctx;
  const user = (req) => String(req.headers["x-user-name"] || "").trim() || null;

  function forwardHeaders(req) {
    const out = { "content-type": "application/json" };
    for (const h of ["authorization", "x-user-name", "x-user-role", "x-user-roles", "x-site-code", "x-user-permissions"]) {
      if (req.headers[h]) out[h] = req.headers[h];
    }
    return out;
  }

  // Recent requisitions (for reprinting).
  app.get("/requisitions", async (req, reply) => {
    if (!requireRoles(req, reply, STORES_ROLES)) return;
    ensureRequisitionSchema(db);
    const days = Math.min(365, Math.max(1, Number(req.query?.days || 30)));
    return { ok: true, rows: listRequisitions(db, { site: getSiteCode(req), days }) };
  });

  // Requisition for request lines already in the stores queue: { request_ids } or { work_order_id }.
  app.post("/requisitions", async (req, reply) => {
    if (!requireRoles(req, reply, STORES_ROLES)) return;
    ensureRequisitionSchema(db);
    let ids = Array.isArray(req.body?.request_ids) ? req.body.request_ids : [];
    const woId = Number(req.body?.work_order_id || 0);
    if (!ids.length && woId) {
      ids = db.prepare(`
        SELECT id FROM maintenance_parts_requests
        WHERE work_order_id = ? AND LOWER(COALESCE(status, 'requested')) IN ('requested', 'ordered', 'received')
        ORDER BY id
      `).all(woId).map((r) => r.id);
      if (!ids.length) return reply.code(404).send({ error: `No open part requests on work order #${woId}` });
    }
    try {
      const r = requisitionFor(db, ids, { site: getSiteCode(req), createdBy: user(req) });
      if (r.created) writeAudit(db, req, { module: "stock", action: "requisition_create", entity_type: "stores_requisition", entity_id: r.id, payload: { request_ids: ids } });
      return { ok: true, id: r.id, created: r.created, requisition: getRequisition(db, r.id) };
    } catch (err) {
      return reply.code(err.status || 400).send({ error: err.message });
    }
  });

  // Someone asks at the counter: record the request lines, then the requisition.
  // { requested_by, asset_code?, work_order_id?, notes?, lines: [{ part_code?, part_name?, qty, urgency? }] }
  app.post("/requisitions/walk-up", async (req, reply) => {
    if (!requireRoles(req, reply, STORES_ROLES)) return;
    ensureRequisitionSchema(db);
    const requestedBy = String(req.body?.requested_by || "").trim().slice(0, 120);
    const lines = (Array.isArray(req.body?.lines) ? req.body.lines : [])
      .map((l) => ({ part_code: String(l?.part_code || "").trim(), part_name: String(l?.part_name || "").trim(), qty: Number(l?.qty || 0), urgency: String(l?.urgency || "normal") }))
      .filter((l) => (l.part_code || l.part_name) && l.qty > 0)
      .map((l) => {
        // A stock code typed into the description box is still that stock item.
        if (l.part_code || !l.part_name) return l;
        const hit = db.prepare(`SELECT part_code FROM parts WHERE UPPER(TRIM(part_code)) = UPPER(TRIM(?))`).get(l.part_name);
        return hit ? { ...l, part_code: hit.part_code, part_name: "" } : l;
      });
    if (!requestedBy) return reply.code(400).send({ error: "Who is asking for the parts?" });
    if (!lines.length) return reply.code(400).send({ error: "Add at least one part with a quantity" });
    const notes = String(req.body?.notes || "").trim().slice(0, 300) || null;
    let assetId = null;
    let woId = Number(req.body?.work_order_id || 0) || null;
    if (woId) {
      const wo = db.prepare(`SELECT id, asset_id FROM work_orders WHERE id = ?`).get(woId);
      if (!wo) return reply.code(404).send({ error: `Work order #${woId} not found` });
      assetId = wo.asset_id;
    }
    const code = String(req.body?.asset_code || "").trim();
    if (code) {
      const a = db.prepare(`SELECT id FROM assets WHERE UPPER(asset_code) = UPPER(?)`).get(code);
      if (!a) return reply.code(404).send({ error: `Machine ${code} not found` });
      if (assetId && assetId !== a.id) return reply.code(409).send({ error: "That machine does not match the work order" });
      assetId = a.id;
    }
    // Same request lines as the workshop's, so they show in the stores queue.
    const ids = [];
    for (const l of lines) {
      const res = await app.inject({
        method: "POST",
        url: "/api/maintenance/parts-requests",
        headers: forwardHeaders(req),
        payload: {
          ...l,
          work_order_id: woId,
          asset_id: assetId,
          notes: [`Asked at stores by ${requestedBy}`, notes].filter(Boolean).join(" — "),
        },
      });
      const body = res.json();
      if (res.statusCode >= 400 || !body?.id) return reply.code(res.statusCode >= 400 ? res.statusCode : 400).send({ error: body?.error || "Could not record the request" });
      ids.push(body.id);
    }
    const r = requisitionFor(db, ids, { site: getSiteCode(req), createdBy: user(req), requestedBy, walkUp: true, notes });
    writeAudit(db, req, { module: "stock", action: "requisition_walk_up", entity_type: "stores_requisition", entity_id: r.id, payload: { requested_by: requestedBy, lines } });
    return { ok: true, id: r.id, requisition: getRequisition(db, r.id) };
  });

  app.get("/requisitions/:id.pdf", async (req, reply) => {
    if (!requireRoles(req, reply, STORES_ROLES)) return;
    ensureRequisitionSchema(db);
    const r = getRequisition(db, req.params.id);
    if (!r) return reply.code(404).send({ error: "Requisition not found" });
    const pdf = await buildRequisitionPdf(db, r);
    db.prepare(`UPDATE stores_requisitions SET printed_count = printed_count + 1, last_printed_at = datetime('now') WHERE id = ?`).run(r.id);
    reply.header("Content-Type", "application/pdf");
    reply.header("Content-Disposition", `inline; filename="${r.number}.pdf"`);
    return reply.send(pdf);
  });
}
