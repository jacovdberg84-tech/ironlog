// IRONLOG/api/routes/procurement/suppliers.routes.js — Suppliers and supplier catalogue.
// Registered by routes/procurement.routes.js; shared helpers arrive through ctx.
import { db } from "../../db/client.js";

export default function registerSuppliersRoutes(app, ctx) {
  const { getPartByCode, requireRoles } = ctx;

  app.post("/suppliers", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores", "procurement"])) return;
    const supplier_code = String(req.body?.supplier_code || "").trim().toUpperCase();
    const supplier_name = String(req.body?.supplier_name || "").trim();
    if (!supplier_code || !supplier_name) return reply.code(400).send({ error: "supplier_code and supplier_name are required" });
    const lead_time_days = Math.max(0, Number(req.body?.lead_time_days || 7));
    const currency = String(req.body?.currency || "USD").trim().toUpperCase() || "USD";
    db.prepare(`
      INSERT INTO suppliers (supplier_code, supplier_name, lead_time_days, currency, active, updated_at)
      VALUES (?, ?, ?, ?, 1, datetime('now'))
      ON CONFLICT(supplier_code) DO UPDATE SET
        supplier_name = excluded.supplier_name,
        lead_time_days = excluded.lead_time_days,
        currency = excluded.currency,
        active = 1,
        updated_at = datetime('now')
    `).run(supplier_code, supplier_name, lead_time_days, currency);
    return { ok: true, supplier_code, supplier_name, lead_time_days, currency };
  });

  app.get("/suppliers", async () => {
    const rows = db.prepare(`
      SELECT id, supplier_code, supplier_name, active, lead_time_days, currency, created_at, updated_at
      FROM suppliers
      ORDER BY supplier_name ASC
      LIMIT 300
    `).all();
    return { ok: true, rows };
  });

  app.post("/supplier-catalog", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores", "procurement"])) return;
    const supplier_code = String(req.body?.supplier_code || "").trim().toUpperCase();
    const part_code = String(req.body?.part_code || "").trim();
    if (!supplier_code || !part_code) return reply.code(400).send({ error: "supplier_code and part_code are required" });
    const supplier = db.prepare(`SELECT id FROM suppliers WHERE supplier_code = ?`).get(supplier_code);
    if (!supplier) return reply.code(404).send({ error: "supplier not found" });
    const part = getPartByCode.get(part_code);
    if (!part) return reply.code(404).send({ error: "part not found" });
    const supplier_part_code = req.body?.supplier_part_code != null ? String(req.body.supplier_part_code).trim() : null;
    const lead_time_days = req.body?.lead_time_days != null ? Math.max(0, Number(req.body.lead_time_days || 0)) : null;
    const last_price = req.body?.last_price != null ? Number(req.body.last_price) : null;
    const currency = req.body?.currency != null ? String(req.body.currency).trim().toUpperCase() : null;
    const effective_date = req.body?.effective_date != null ? String(req.body.effective_date).trim() : null;
    db.prepare(`
      INSERT INTO supplier_part_catalog (
        supplier_id, part_id, supplier_part_code, lead_time_days, last_price, currency, effective_date, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))
      ON CONFLICT(supplier_id, part_id) DO UPDATE SET
        supplier_part_code = excluded.supplier_part_code,
        lead_time_days = excluded.lead_time_days,
        last_price = excluded.last_price,
        currency = excluded.currency,
        effective_date = excluded.effective_date,
        updated_at = datetime('now')
    `).run(Number(supplier.id), Number(part.id), supplier_part_code, lead_time_days, last_price, currency, effective_date);
    return { ok: true, supplier_code, part_code };
  });

  app.get("/supplier-catalog", async (req, reply) => {
    const part_code = String(req.query?.part_code || "").trim();
    const supplier_code = String(req.query?.supplier_code || "").trim().toUpperCase();
    const where = [];
    const params = [];
    if (part_code) {
      where.push("p.part_code = ?");
      params.push(part_code);
    }
    if (supplier_code) {
      where.push("s.supplier_code = ?");
      params.push(supplier_code);
    }
    const rows = db.prepare(`
      SELECT
        c.id,
        s.supplier_code,
        s.supplier_name,
        p.part_code,
        p.part_name,
        c.supplier_part_code,
        c.lead_time_days,
        c.last_price,
        c.currency,
        c.effective_date,
        c.updated_at
      FROM supplier_part_catalog c
      JOIN suppliers s ON s.id = c.supplier_id
      JOIN parts p ON p.id = c.part_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY s.supplier_name ASC, p.part_code ASC
      LIMIT 500
    `).all(...params);
    return reply.send({ ok: true, rows });
  });
}
