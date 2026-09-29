// IRONLOG/api/routes/procurement/purchase-orders.routes.js — Purchase orders, receipts, invoices, three-way match and exceptions.
// Registered by routes/procurement.routes.js; shared helpers arrive through ctx.
import { db } from "../../db/client.js";
import { getSiteCode, getUser } from "../../utils/request.js";

export default function registerPurchaseOrdersRoutes(app, ctx) {
  const {
    bySiteReq,
    derivePoStatus,
    getLocationByCode,
    nextPONumber,
    nextReceiptNumber,
    requireRoles,
  } = ctx;

  app.post("/requisitions/:id/create-po", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const siteCode = getSiteCode(req);
    const reqn = db.prepare(`
      SELECT *
      FROM procurement_requisitions
      WHERE id = ? AND ${bySiteReq}
    `).get(id, siteCode);
    if (!reqn) return reply.code(404).send({ error: "requisition not found" });
    if (!["approved", "approved_all", "po_ready", "received", "partially_received"].includes(String(reqn.status || "").toLowerCase())) {
      return reply.code(409).send({ error: "requisition must be approved before PO creation" });
    }
    const existing = db.prepare(`SELECT id, po_number, status FROM procurement_purchase_orders WHERE requisition_id = ? ORDER BY id DESC LIMIT 1`).get(id);
    if (existing && String(existing.status || "").toLowerCase() !== "cancelled") {
      return reply.send({ ok: true, duplicate: true, po_id: Number(existing.id), po_number: String(existing.po_number || "") });
    }
    const lines = db.prepare(`
      SELECT *
      FROM procurement_requisition_lines
      WHERE requisition_id = ?
      ORDER BY line_no ASC
    `).all(id);
    if (!lines.length) return reply.code(409).send({ error: "cannot create PO without requisition lines" });
    const supplier = reqn.supplier_name
      ? db.prepare(`SELECT id, supplier_code, supplier_name, currency FROM suppliers WHERE LOWER(supplier_name) = LOWER(?) OR UPPER(supplier_code) = UPPER(?) LIMIT 1`).get(String(reqn.supplier_name), String(reqn.supplier_name))
      : null;
    const po_number = nextPONumber();
    const currency = supplier?.currency ? String(supplier.currency).toUpperCase() : "USD";
    const tx = db.transaction(() => {
      const head = db.prepare(`
        INSERT INTO procurement_purchase_orders (
          po_number, requisition_id, supplier_id, site_code, currency, status, notes, created_by, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, 'draft', ?, ?, datetime('now'), datetime('now'))
      `).run(po_number, id, supplier ? Number(supplier.id) : null, siteCode, currency, reqn.notes || null, getUser(req));
      const po_id = Number(head.lastInsertRowid);
      const insLine = db.prepare(`
        INSERT INTO procurement_purchase_order_lines (
          po_id, line_no, requisition_line_id, part_id, description, quantity_ordered, unit_price, line_total, needed_by_date, cost_center_code, labor_tag
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      let subtotal = 0;
      for (const l of lines) {
        const qty = Number(l.quantity || 0);
        const unit = Number(l.net_price ?? l.gross_price ?? 0);
        const line_total = Number((qty * unit).toFixed(2));
        subtotal += line_total;
        insLine.run(
          po_id,
          Number(l.line_no || 0),
          Number(l.id || 0),
          l.part_id ? Number(l.part_id) : null,
          l.description || null,
          qty,
          unit,
          line_total,
          l.needed_by_date || null,
          req.body?.cost_center_code ? String(req.body.cost_center_code).trim() : null,
          req.body?.labor_tag ? String(req.body.labor_tag).trim() : null
        );
      }
      db.prepare(`UPDATE procurement_purchase_orders SET subtotal = ?, updated_at = datetime('now') WHERE id = ?`).run(Number(subtotal.toFixed(2)), po_id);
      db.prepare(`
        UPDATE procurement_requisitions
        SET status = 'po_ready', po_number = ?, updated_at = datetime('now')
        WHERE id = ?
      `).run(po_number, id);
      return po_id;
    });
    const po_id = tx();
    return { ok: true, po_id, po_number };
  });

  app.get("/purchase-orders", async (req, reply) => {
    const status = String(req.query?.status || "").trim().toLowerCase();
    const siteCode = getSiteCode(req);
    const where = [`LOWER(TRIM(COALESCE(po.site_code, 'main'))) = ?`];
    const params = [siteCode];
    if (status) {
      where.push("LOWER(po.status) = ?");
      params.push(status);
    }
    const rows = db.prepare(`
      SELECT
        po.id,
        po.po_number,
        po.requisition_id,
        po.currency,
        po.status,
        po.subtotal,
        po.created_at,
        po.updated_at,
        s.supplier_code,
        s.supplier_name
      FROM procurement_purchase_orders po
      LEFT JOIN suppliers s ON s.id = po.supplier_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY po.id DESC
      LIMIT 300
    `).all(...params).map((r) => ({ ...r, subtotal: Number(r.subtotal || 0) }));
    return reply.send({ ok: true, rows });
  });

  app.get("/purchase-orders/:id/detail", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const siteCode = getSiteCode(req);
    const po = db.prepare(`
      SELECT po.*, s.supplier_code, s.supplier_name
      FROM procurement_purchase_orders po
      LEFT JOIN suppliers s ON s.id = po.supplier_id
      WHERE po.id = ? AND LOWER(TRIM(COALESCE(po.site_code, 'main'))) = ?
    `).get(id, siteCode);
    if (!po) return reply.code(404).send({ error: "PO not found" });
    const lines = db.prepare(`
      SELECT pol.*, p.part_code, p.part_name
      FROM procurement_purchase_order_lines pol
      LEFT JOIN parts p ON p.id = pol.part_id
      WHERE pol.po_id = ?
      ORDER BY pol.line_no ASC
    `).all(id);
    const receipts = db.prepare(`
      SELECT id, receipt_number, receipt_date, status, location_code, received_by, created_at
      FROM procurement_goods_receipts
      WHERE po_id = ?
      ORDER BY id DESC
    `).all(id);
    return reply.send({ ok: true, po, lines, receipts });
  });

  app.post("/purchase-orders/:id/approve", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const po = db.prepare(`SELECT id, status FROM procurement_purchase_orders WHERE id = ?`).get(id);
    if (!po) return reply.code(404).send({ error: "PO not found" });
    if (!["draft", "approved", "sent", "partially_received", "received"].includes(String(po.status || "").toLowerCase())) {
      return reply.code(409).send({ error: `cannot approve from status ${po.status}` });
    }
    db.prepare(`
      UPDATE procurement_purchase_orders
      SET status = 'approved', approved_at = datetime('now'), updated_at = datetime('now')
      WHERE id = ?
    `).run(id);
    return { ok: true, id, status: "approved" };
  });

  app.post("/purchase-orders/:id/send", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const po = db.prepare(`SELECT id, status FROM procurement_purchase_orders WHERE id = ?`).get(id);
    if (!po) return reply.code(404).send({ error: "PO not found" });
    if (!["approved", "sent", "partially_received", "received"].includes(String(po.status || "").toLowerCase())) {
      return reply.code(409).send({ error: `cannot send from status ${po.status}` });
    }
    db.prepare(`
      UPDATE procurement_purchase_orders
      SET status = CASE WHEN status = 'approved' THEN 'sent' ELSE status END,
          sent_at = COALESCE(sent_at, datetime('now')),
          updated_at = datetime('now')
      WHERE id = ?
    `).run(id);
    return { ok: true, id };
  });

  app.post("/purchase-orders/:id/receive", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores", "procurement"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const po = db.prepare(`
      SELECT id, po_number, requisition_id, status
      FROM procurement_purchase_orders
      WHERE id = ?
    `).get(id);
    if (!po) return reply.code(404).send({ error: "PO not found" });
    if (!["approved", "sent", "partially_received", "received"].includes(String(po.status || "").toLowerCase())) {
      return reply.code(409).send({ error: `cannot receive from status ${po.status}` });
    }
    const lines = Array.isArray(req.body?.lines) ? req.body.lines : [];
    if (!lines.length) return reply.code(400).send({ error: "lines array is required" });
    const location_code = String(req.body?.location_code || "MAIN").trim().toUpperCase();
    const location = getLocationByCode.get(location_code);
    if (!location) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    const receipt_date = req.body?.receipt_date ? String(req.body.receipt_date).trim() : new Date().toISOString().slice(0, 10);
    const receipt_number = String(req.body?.receipt_number || "").trim() || nextReceiptNumber();

    const tx = db.transaction(() => {
      const insReceipt = db.prepare(`
        INSERT INTO procurement_goods_receipts (
          po_id, receipt_number, receipt_date, received_by, location_code, status, notes
        ) VALUES (?, ?, ?, ?, ?, 'posted', ?)
      `).run(id, receipt_number, receipt_date, getUser(req), location_code, req.body?.notes ? String(req.body.notes).trim() : null);
      const receipt_id = Number(insReceipt.lastInsertRowid);
      const getPoLine = db.prepare(`
        SELECT id, line_no, part_id, quantity_ordered, quantity_received, unit_price
        FROM procurement_purchase_order_lines
        WHERE id = ? AND po_id = ?
      `);
      const insReceiptLine = db.prepare(`
        INSERT INTO procurement_goods_receipt_lines (
          receipt_id, po_line_id, part_id, quantity_received, unit_price, line_total, cost_center_code
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
      `);
      const updPoLine = db.prepare(`
        UPDATE procurement_purchase_order_lines
        SET quantity_received = quantity_received + ?
        WHERE id = ?
      `);
      const insMove = db.prepare(`
        INSERT INTO stock_movements (
          part_id, quantity, movement_type, reference, location_id
        ) VALUES (?, ?, 'in', ?, ?)
      `);

      for (const row of lines) {
        const po_line_id = Number(row?.po_line_id || 0);
        const qty = Number(row?.quantity_received || 0);
        if (!Number.isFinite(po_line_id) || po_line_id <= 0 || !Number.isFinite(qty) || qty <= 0) {
          throw new Error("each line requires po_line_id and quantity_received > 0");
        }
        const poLine = getPoLine.get(po_line_id, id);
        if (!poLine) throw new Error(`po_line_id ${po_line_id} not found for PO`);
        const outstanding = Number(poLine.quantity_ordered || 0) - Number(poLine.quantity_received || 0);
        if (qty > outstanding + 1e-9) {
          throw new Error(`line ${poLine.line_no}: received qty exceeds outstanding (${outstanding})`);
        }
        const unit_price = row?.unit_price != null && row.unit_price !== "" ? Number(row.unit_price) : Number(poLine.unit_price || 0);
        const line_total = Number((qty * unit_price).toFixed(2));
        insReceiptLine.run(
          receipt_id,
          po_line_id,
          poLine.part_id ? Number(poLine.part_id) : null,
          qty,
          unit_price,
          line_total,
          row?.cost_center_code ? String(row.cost_center_code).trim() : null
        );
        updPoLine.run(qty, po_line_id);
        if (poLine.part_id) {
          insMove.run(Number(poLine.part_id), qty, `po:${id}:receipt:${receipt_id}:line:${po_line_id}`, Number(location.id));
        }
      }

      const poStatus = derivePoStatus(id);
      db.prepare(`UPDATE procurement_purchase_orders SET status = ?, updated_at = datetime('now') WHERE id = ?`).run(poStatus, id);
      if (Number(po.requisition_id || 0) > 0) {
        const reqStatus = poStatus === "received" ? "received" : "partially_received";
        db.prepare(`UPDATE procurement_requisitions SET status = ?, updated_at = datetime('now') WHERE id = ?`).run(reqStatus, Number(po.requisition_id));
      }
      return { receipt_id, receipt_number, po_status: poStatus };
    });
    const result = tx();
    return { ok: true, ...result };
  });

  app.post("/purchase-orders/:id/invoices", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const po = db.prepare(`SELECT id, supplier_id, currency FROM procurement_purchase_orders WHERE id = ?`).get(id);
    if (!po) return reply.code(404).send({ error: "PO not found" });
    const invoice_number = String(req.body?.invoice_number || "").trim();
    const invoice_date = String(req.body?.invoice_date || "").trim() || new Date().toISOString().slice(0, 10);
    if (!invoice_number) return reply.code(400).send({ error: "invoice_number is required" });
    const lines = Array.isArray(req.body?.lines) ? req.body.lines : [];
    if (!lines.length) return reply.code(400).send({ error: "lines array is required" });
    const tx = db.transaction(() => {
      const head = db.prepare(`
        INSERT INTO procurement_invoices (
          po_id, invoice_number, supplier_id, invoice_date, status, currency, subtotal, tax, total, captured_by, notes
        ) VALUES (?, ?, ?, ?, 'captured', ?, 0, ?, 0, ?, ?)
      `).run(
        id,
        invoice_number,
        po.supplier_id ? Number(po.supplier_id) : null,
        invoice_date,
        String(req.body?.currency || po.currency || "USD").trim().toUpperCase(),
        Number(req.body?.tax || 0),
        getUser(req),
        req.body?.notes ? String(req.body.notes).trim() : null
      );
      const invoice_id = Number(head.lastInsertRowid);
      const insLine = db.prepare(`
        INSERT INTO procurement_invoice_lines (
          invoice_id, po_line_id, part_id, description, quantity_invoiced, unit_price, line_total, cost_center_code
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
      `);
      let subtotal = 0;
      for (const l of lines) {
        const po_line_id = Number(l?.po_line_id || 0);
        const qty = Number(l?.quantity_invoiced || 0);
        const unit_price = Number(l?.unit_price || 0);
        if (!Number.isFinite(po_line_id) || po_line_id <= 0 || !Number.isFinite(qty) || qty <= 0 || !Number.isFinite(unit_price) || unit_price < 0) {
          throw new Error("invoice lines require po_line_id, quantity_invoiced > 0, unit_price >= 0");
        }
        const poLine = db.prepare(`SELECT id, part_id, description FROM procurement_purchase_order_lines WHERE id = ? AND po_id = ?`).get(po_line_id, id);
        if (!poLine) throw new Error(`po_line_id ${po_line_id} not found for PO`);
        const line_total = Number((qty * unit_price).toFixed(2));
        subtotal += line_total;
        insLine.run(
          invoice_id,
          po_line_id,
          poLine.part_id ? Number(poLine.part_id) : null,
          l?.description ? String(l.description).trim() : (poLine.description || null),
          qty,
          unit_price,
          line_total,
          l?.cost_center_code ? String(l.cost_center_code).trim() : null
        );
      }
      const tax = Number(req.body?.tax || 0);
      const total = Number((subtotal + tax).toFixed(2));
      db.prepare(`UPDATE procurement_invoices SET subtotal = ?, total = ? WHERE id = ?`).run(Number(subtotal.toFixed(2)), total, invoice_id);
      return { invoice_id, invoice_number, subtotal: Number(subtotal.toFixed(2)), tax, total };
    });
    const result = tx();
    return { ok: true, ...result };
  });

  app.post("/purchase-orders/:id/three-way-match", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const po = db.prepare(`SELECT id FROM procurement_purchase_orders WHERE id = ?`).get(id);
    if (!po) return reply.code(404).send({ error: "PO not found" });
    const qtyTol = Math.max(0, Number(req.body?.quantity_tolerance || 0));
    const priceTolPct = Math.max(0, Number(req.body?.price_tolerance_pct || 0));
    const totalTol = Math.max(0, Number(req.body?.total_tolerance || 0));

    db.prepare(`DELETE FROM procurement_match_exceptions WHERE po_id = ? AND status = 'open'`).run(id);

    const poLines = db.prepare(`
      SELECT id, line_no, quantity_ordered, quantity_received, unit_price, line_total
      FROM procurement_purchase_order_lines
      WHERE po_id = ?
      ORDER BY line_no ASC
    `).all(id);
    const invoiceLines = db.prepare(`
      SELECT il.id, il.invoice_id, il.po_line_id, il.quantity_invoiced, il.unit_price, il.line_total
      FROM procurement_invoice_lines il
      JOIN procurement_invoices i ON i.id = il.invoice_id
      WHERE i.po_id = ?
    `).all(id);
    const receiptLines = db.prepare(`
      SELECT rl.id, rl.receipt_id, rl.po_line_id, rl.quantity_received, rl.unit_price, rl.line_total
      FROM procurement_goods_receipt_lines rl
      JOIN procurement_goods_receipts r ON r.id = rl.receipt_id
      WHERE r.po_id = ?
    `).all(id);
    const invByPoLine = invoiceLines.reduce((m, r) => {
      const k = Number(r.po_line_id || 0);
      const cur = m.get(k) || { qty: 0, total: 0, last_unit: 0, lines: [] };
      cur.qty += Number(r.quantity_invoiced || 0);
      cur.total += Number(r.line_total || 0);
      cur.last_unit = Number(r.unit_price || 0);
      cur.lines.push(r);
      m.set(k, cur);
      return m;
    }, new Map());
    const recByPoLine = receiptLines.reduce((m, r) => {
      const k = Number(r.po_line_id || 0);
      const cur = m.get(k) || { qty: 0, total: 0, lines: [] };
      cur.qty += Number(r.quantity_received || 0);
      cur.total += Number(r.line_total || 0);
      cur.lines.push(r);
      m.set(k, cur);
      return m;
    }, new Map());

    const insEx = db.prepare(`
      INSERT INTO procurement_match_exceptions (
        po_id, po_line_id, invoice_id, invoice_line_id, receipt_id, receipt_line_id,
        exception_type, severity, status, details_json, assigned_to
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'open', ?, ?)
    `);

    let exception_count = 0;
    for (const line of poLines) {
      const po_line_id = Number(line.id || 0);
      const poQty = Number(line.quantity_ordered || 0);
      const poUnit = Number(line.unit_price || 0);
      const poTotal = Number(line.line_total || 0);
      const inv = invByPoLine.get(po_line_id) || { qty: 0, total: 0, last_unit: 0, lines: [] };
      const rec = recByPoLine.get(po_line_id) || { qty: 0, total: 0, lines: [] };

      const qtyVsRec = Math.abs(poQty - rec.qty);
      const qtyVsInv = Math.abs(rec.qty - inv.qty);
      const priceVarPct = poUnit > 0 ? (Math.abs(inv.last_unit - poUnit) / poUnit) * 100 : 0;
      const totalVar = Math.abs(rec.total - inv.total);

      if (qtyVsRec > qtyTol) {
        insEx.run(id, po_line_id, null, null, rec.lines[0]?.receipt_id || null, rec.lines[0]?.id || null, "PO_vs_Receipt_qty", "high", JSON.stringify({ po_qty: poQty, receipt_qty: rec.qty, tolerance: qtyTol }), null);
        exception_count += 1;
      }
      if (qtyVsInv > qtyTol) {
        insEx.run(id, po_line_id, inv.lines[0]?.invoice_id || null, inv.lines[0]?.id || null, rec.lines[0]?.receipt_id || null, rec.lines[0]?.id || null, "Receipt_vs_Invoice_qty", "high", JSON.stringify({ receipt_qty: rec.qty, invoice_qty: inv.qty, tolerance: qtyTol }), null);
        exception_count += 1;
      }
      if (priceVarPct > priceTolPct) {
        insEx.run(id, po_line_id, inv.lines[0]?.invoice_id || null, inv.lines[0]?.id || null, null, null, "PO_vs_Invoice_unit_price", "warn", JSON.stringify({ po_unit_price: poUnit, invoice_unit_price: inv.last_unit, variance_pct: Number(priceVarPct.toFixed(2)), tolerance_pct: priceTolPct }), null);
        exception_count += 1;
      }
      if (totalVar > totalTol) {
        insEx.run(id, po_line_id, inv.lines[0]?.invoice_id || null, inv.lines[0]?.id || null, rec.lines[0]?.receipt_id || null, rec.lines[0]?.id || null, "Receipt_vs_Invoice_total", "warn", JSON.stringify({ receipt_total: Number(rec.total.toFixed(2)), invoice_total: Number(inv.total.toFixed(2)), variance_total: Number(totalVar.toFixed(2)), tolerance_total: totalTol, po_total: poTotal }), null);
        exception_count += 1;
      }
    }
    return { ok: true, po_id: id, exception_count };
  });

  app.get("/exceptions", async (req, reply) => {
    const status = String(req.query?.status || "open").trim().toLowerCase();
    const rows = db.prepare(`
      SELECT
        e.*,
        po.po_number
      FROM procurement_match_exceptions e
      LEFT JOIN procurement_purchase_orders po ON po.id = e.po_id
      WHERE LOWER(e.status) = ?
      ORDER BY e.id DESC
      LIMIT 500
    `).all(status).map((r) => ({
      ...r,
      details: (() => {
        try { return JSON.parse(String(r.details_json || "{}")); } catch { return {}; }
      })(),
    }));
    return reply.send({ ok: true, rows });
  });

  app.post("/exceptions/:id/resolve", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const ex = db.prepare(`SELECT id, status FROM procurement_match_exceptions WHERE id = ?`).get(id);
    if (!ex) return reply.code(404).send({ error: "exception not found" });
    if (String(ex.status || "").toLowerCase() === "resolved") return { ok: true, duplicate: true, id };
    db.prepare(`
      UPDATE procurement_match_exceptions
      SET status = 'resolved', resolved_by = ?, resolved_at = datetime('now')
      WHERE id = ?
    `).run(getUser(req), id);
    return { ok: true, id, status: "resolved" };
  });
}
