// IRONLOG/api/routes/stock/part-orders.routes.js — Stores part orders and store QR profile.
// Registered by routes/stock.routes.js; shared helpers arrive through ctx.
import { db } from "../../db/client.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerPartOrdersRoutes(app, ctx) {
  const {
    PART_ORDER_STATUSES,
    PART_ORDER_WRITE_ROLES,
    buildStoreQrProfile,
    getPartByCode,
    getSiteCode,
    getStoredStoreQrProfile,
    isYmd,
    mapPartOrderRow,
    receivePartOrderToStock,
    requireRoles,
    summarizePartOrders,
    syncLinkedPartOrder,
    upsertStoreQrProfile,
    validatePartOrderLinks,
  } = ctx;

  // GET /api/stock/part-orders?start=&end=&status=
  app.get("/part-orders", async (req, reply) => {
    const site_code = getSiteCode(req);
    const start = String(req.query?.start || req.query?.date_from || "").trim();
    const end = String(req.query?.end || req.query?.date_to || "").trim();
    const status = String(req.query?.status || "").trim().toLowerCase();

    const where = ["LOWER(TRIM(COALESCE(o.site_code, 'main'))) = ?"];
    const params = [site_code];
    if (isYmd(start)) {
      where.push("o.order_date >= ?");
      params.push(start);
    }
    if (isYmd(end)) {
      where.push("o.order_date <= ?");
      params.push(end);
    }
    if (status && PART_ORDER_STATUSES.has(status)) {
      where.push("LOWER(COALESCE(o.status, 'on_order')) = ?");
      params.push(status);
    } else {
      where.push("LOWER(COALESCE(o.status, 'on_order')) <> 'cancelled'");
    }

    const rows = db.prepare(`
      SELECT
        o.id,
        o.site_code,
        o.part_id,
        o.part_code,
        o.part_name,
        o.qty,
        o.unit_cost,
        o.currency,
        o.supplier_name,
        o.po_number,
        o.requisition_number,
        o.invoice_number,
        o.current_location,
        o.asset_id,
        a.asset_code,
        a.asset_name,
        o.work_order_id,
        o.breakdown_id,
        o.offsite_repair_id,
        o.responsible_person,
        o.order_date,
        o.expected_arrival_date,
        o.arrived_date,
        o.status,
        o.notes,
        o.stock_movement_id,
        o.created_by,
        o.created_at,
        o.updated_at,
        p.part_name AS catalog_part_name
      FROM stores_part_orders o
      LEFT JOIN parts p ON p.id = o.part_id
      LEFT JOIN assets a ON a.id = o.asset_id
      WHERE ${where.join(" AND ")}
      ORDER BY o.order_date DESC, o.id DESC
      LIMIT 2000
    `).all(...params).map(mapPartOrderRow);

    return reply.send({
      ok: true,
      site_code,
      start: start || null,
      end: end || null,
      rows,
      summary: summarizePartOrders(rows),
    });
  });

  // POST /api/stock/part-orders
  app.post("/part-orders", async (req, reply) => {
    if (!requireRoles(req, reply, PART_ORDER_WRITE_ROLES)) return;
    const body = req.body || {};
    const site_code = getSiteCode(req);
    const created_by = String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
    const part_code = String(body.part_code || "").trim();
    let part_name = String(body.part_name || "").trim();
    const qty = Number(body.qty ?? 1);
    let unit_cost = Math.max(0, Number(body.unit_cost ?? 0));
    const currency = String(body.currency || "USD").trim().toUpperCase() || "USD";
    const supplier_name = String(body.supplier_name || "").trim() || null;
    const po_number = String(body.po_number || "").trim() || null;
    const requisition_number = String(body.requisition_number || "").trim() || null;
    const invoice_number = String(body.invoice_number || "").trim() || null;
    const current_location = String(body.current_location || "").trim() || null;
    const asset_code = String(body.asset_code || "").trim();
    const linkedAsset = asset_code
      ? db.prepare(`SELECT id, asset_code FROM assets WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?)) LIMIT 1`).get(asset_code)
      : null;
    if (asset_code && !linkedAsset) return reply.code(404).send({ error: "linked asset not found" });
    let asset_id = linkedAsset ? Number(linkedAsset.id) : null;
    const work_order_id = Number(body.work_order_id || 0) || null;
    const breakdown_id = Number(body.breakdown_id || 0) || null;
    const offsite_repair_id = Number(body.offsite_repair_id || 0) || null;
    const responsible_person = String(body.responsible_person || "").trim() || null;
    const linkCheck = validatePartOrderLinks({ asset_id, work_order_id, breakdown_id, offsite_repair_id });
    if (linkCheck.error) return reply.code(linkCheck.status).send({ error: linkCheck.error });
    asset_id = linkCheck.asset_id;
    const order_date = String(body.order_date || "").trim();
    const expected_arrival_date = String(body.expected_arrival_date || "").trim() || null;
    const notes = String(body.notes || "").trim() || null;
    let status = String(body.status || "on_order").trim().toLowerCase();
    if (!PART_ORDER_STATUSES.has(status) || status === "cancelled") status = "on_order";
    if (!isYmd(order_date)) return reply.code(400).send({ error: "order_date must be YYYY-MM-DD" });
    if (!Number.isFinite(qty) || qty <= 0) return reply.code(400).send({ error: "qty must be > 0" });
    if (!part_code && !part_name) return reply.code(400).send({ error: "part_code or part_name is required" });

    let part_id = null;
    if (part_code) {
      const part = getPartByCode.get(part_code);
      if (part) {
        part_id = Number(part.id);
        if (!part_name) part_name = String(part.part_name || part_code).trim();
        if (!unit_cost && Number(part.unit_cost || 0) > 0) unit_cost = Number(part.unit_cost);
      }
    }
    if (!part_name) part_name = part_code;

    const arrived_date = status === "arrived" ? (isYmd(body.arrived_date) ? body.arrived_date : order_date) : null;
    const now = new Date().toISOString();

    const info = db.prepare(`
      INSERT INTO stores_part_orders (
        site_code, part_id, part_code, part_name, qty, unit_cost, currency,
        supplier_name, po_number, requisition_number, invoice_number, current_location,
        asset_id, work_order_id, breakdown_id, offsite_repair_id, responsible_person,
        order_date, expected_arrival_date, arrived_date,
        status, notes, created_by, created_at, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(
      site_code,
      part_id,
      part_code || null,
      part_name,
      qty,
      unit_cost,
      currency,
      supplier_name,
      po_number,
      requisition_number,
      invoice_number,
      current_location,
      asset_id,
      work_order_id,
      breakdown_id,
      offsite_repair_id,
      responsible_person,
      order_date,
      expected_arrival_date,
      arrived_date,
      status,
      notes,
      created_by,
      now,
      now,
    );

    writeAudit(db, req, {
      module: "stock",
      action: "part_order.create",
      entity_type: "stores_part_order",
      entity_id: String(info.lastInsertRowid),
      after: { part_code, part_name, qty, unit_cost, status, order_date },
    });

    const newId = Number(info.lastInsertRowid || 0);
    let stock_receipt = null;
    if (status === "arrived" && newId) {
      const row = db.prepare(`SELECT * FROM stores_part_orders WHERE id = ?`).get(newId);
      stock_receipt = receivePartOrderToStock(row, req);
    }
    syncLinkedPartOrder(newId);

    return reply.send({ ok: true, id: newId, stock_receipt });
  });

  // PATCH /api/stock/part-orders/:id
  app.patch("/part-orders/:id", async (req, reply) => {
    if (!requireRoles(req, reply, PART_ORDER_WRITE_ROLES)) return;
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ error: "invalid id" });
    const site_code = getSiteCode(req);
    const body = req.body || {};
    const existing = db.prepare(`
      SELECT * FROM stores_part_orders
      WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
    `).get(id, site_code);
    if (!existing) return reply.code(404).send({ error: "part order not found" });

    const prevStatus = String(existing.status || "on_order").toLowerCase();

    const status = body.status != null ? String(body.status).trim().toLowerCase() : prevStatus;
    if (!PART_ORDER_STATUSES.has(status)) return reply.code(400).send({ error: "invalid status" });

    const qty = body.qty != null ? Number(body.qty) : Number(existing.qty || 0);
    const unit_cost = body.unit_cost != null ? Math.max(0, Number(body.unit_cost)) : Number(existing.unit_cost || 0);
    if (!Number.isFinite(qty) || qty <= 0) return reply.code(400).send({ error: "qty must be > 0" });

    const order_date = body.order_date != null ? String(body.order_date).trim() : String(existing.order_date || "");
    if (!isYmd(order_date)) return reply.code(400).send({ error: "order_date must be YYYY-MM-DD" });

    let arrived_date = body.arrived_date != null ? String(body.arrived_date).trim() || null : existing.arrived_date;
    if (status === "arrived" && !isYmd(arrived_date || "")) {
      arrived_date = new Date().toISOString().slice(0, 10);
    }
    if (status !== "arrived") arrived_date = body.arrived_date === null ? null : arrived_date;

    let asset_id = existing.asset_id ?? null;
    if (body.asset_code != null) {
      const assetCode = String(body.asset_code || "").trim();
      const linked = assetCode
        ? db.prepare(`SELECT id FROM assets WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?)) LIMIT 1`).get(assetCode)
        : null;
      if (assetCode && !linked) return reply.code(404).send({ error: "linked asset not found" });
      asset_id = linked ? Number(linked.id) : null;
    }
    const work_order_id = body.work_order_id !== undefined ? (Number(body.work_order_id || 0) || null) : existing.work_order_id;
    const breakdown_id = body.breakdown_id !== undefined ? (Number(body.breakdown_id || 0) || null) : existing.breakdown_id;
    const offsite_repair_id = body.offsite_repair_id !== undefined ? (Number(body.offsite_repair_id || 0) || null) : existing.offsite_repair_id;
    const linkCheck = validatePartOrderLinks({ asset_id, work_order_id, breakdown_id, offsite_repair_id });
    if (linkCheck.error) return reply.code(linkCheck.status).send({ error: linkCheck.error });
    asset_id = linkCheck.asset_id;

    const now = new Date().toISOString();
    db.prepare(`
      UPDATE stores_part_orders
      SET
        part_name = ?,
        qty = ?,
        unit_cost = ?,
        currency = ?,
        supplier_name = ?,
        po_number = ?,
        requisition_number = ?,
        invoice_number = ?,
        current_location = ?,
        asset_id = ?,
        work_order_id = ?,
        breakdown_id = ?,
        offsite_repair_id = ?,
        responsible_person = ?,
        order_date = ?,
        expected_arrival_date = ?,
        arrived_date = ?,
        status = ?,
        notes = ?,
        updated_at = ?
      WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
    `).run(
      String(body.part_name || existing.part_name || "").trim() || existing.part_name,
      qty,
      unit_cost,
      String(body.currency || existing.currency || "USD").trim().toUpperCase() || "USD",
      body.supplier_name != null ? String(body.supplier_name).trim() || null : existing.supplier_name,
      body.po_number != null ? String(body.po_number).trim() || null : existing.po_number,
      body.requisition_number != null
        ? String(body.requisition_number).trim() || null
        : existing.requisition_number,
      body.invoice_number != null
        ? String(body.invoice_number).trim() || null
        : existing.invoice_number,
      body.current_location != null
        ? String(body.current_location).trim() || null
        : existing.current_location,
      asset_id,
      work_order_id,
      breakdown_id,
      offsite_repair_id,
      body.responsible_person !== undefined ? String(body.responsible_person || "").trim() || null : existing.responsible_person,
      order_date,
      body.expected_arrival_date != null ? String(body.expected_arrival_date).trim() || null : existing.expected_arrival_date,
      arrived_date,
      status,
      body.notes != null ? String(body.notes).trim() || null : existing.notes,
      now,
      id,
      site_code,
    );

    writeAudit(db, req, {
      module: "stock",
      action: "part_order.update",
      entity_type: "stores_part_order",
      entity_id: String(id),
      before: existing,
      after: { status, qty, unit_cost, order_date, arrived_date },
    });

    let stock_receipt = null;
    if (status === "arrived" && prevStatus !== "arrived") {
      const updated = db.prepare(`SELECT * FROM stores_part_orders WHERE id = ?`).get(id);
      stock_receipt = receivePartOrderToStock(updated, req);
    }
    syncLinkedPartOrder(id);

    return reply.send({ ok: true, id, stock_receipt });
  });

  // DELETE /api/stock/part-orders/:id  (soft cancel)
  app.delete("/part-orders/:id", async (req, reply) => {
    if (!requireRoles(req, reply, PART_ORDER_WRITE_ROLES)) return;
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ error: "invalid id" });
    const site_code = getSiteCode(req);
    const existing = db.prepare(`
      SELECT id FROM stores_part_orders
      WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
    `).get(id, site_code);
    if (!existing) return reply.code(404).send({ error: "part order not found" });
    db.prepare(`
      UPDATE stores_part_orders
      SET status = 'cancelled', updated_at = datetime('now')
      WHERE id = ?
    `).run(id);
    return reply.send({ ok: true, id, status: "cancelled" });
  });

  // POST /api/stock/part-orders/:id/receive — push an arrived line into store inventory
  app.post("/part-orders/:id/receive", async (req, reply) => {
    if (!requireRoles(req, reply, PART_ORDER_WRITE_ROLES)) return;
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ error: "invalid id" });
    const site_code = getSiteCode(req);
    const row = db.prepare(`
      SELECT * FROM stores_part_orders
      WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
    `).get(id, site_code);
    if (!row) return reply.code(404).send({ error: "part order not found" });

    const stock_receipt = receivePartOrderToStock(row, req);
    if (!stock_receipt.received && !stock_receipt.already && stock_receipt.error) {
      return reply.code(400).send({ ok: false, error: stock_receipt.error, stock_receipt });
    }
    return reply.send({ ok: true, id, stock_receipt });
  });

  // GET /api/stock/store-qr-profile
  app.get("/store-qr-profile", async (req, reply) => {
    const site_code = String(req.query?.site || getSiteCode(req)).trim().toLowerCase() || "main";
    const stored = getStoredStoreQrProfile.get(site_code);
    let storedPayload = null;
    if (stored?.qr_payload) {
      try {
        storedPayload = JSON.parse(String(stored.qr_payload || "{}"));
      } catch {
        storedPayload = null;
      }
    }
    const live = buildStoreQrProfile(site_code, req);
    return reply.send({
      ok: true,
      site_code,
      stored: storedPayload
        ? { qr_payload: storedPayload, qr_text: stored.qr_text, generated_at: stored.generated_at }
        : null,
      live_preview: live.profile,
      live_qr_text: live.qrText,
    });
  });

  // POST /api/stock/store-qr-profile/refresh
  app.post("/store-qr-profile/refresh", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores", "storeman"])) return;
    const site_code = String(req.body?.site || req.query?.site || getSiteCode(req)).trim().toLowerCase() || "main";
    const built = buildStoreQrProfile(site_code, req);
    upsertStoreQrProfile.run(site_code, JSON.stringify(built.profile), built.qrText);
    return reply.send({
      ok: true,
      site_code,
      qr_payload: built.profile,
      qr_text: built.qrText,
    });
  });
}
