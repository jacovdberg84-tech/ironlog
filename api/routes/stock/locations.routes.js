// IRONLOG/api/routes/stock/locations.routes.js — Locations, bins, stock depth, min/max levels and replenishment.
// Registered by routes/stock.routes.js; shared helpers arrive through ctx.
import { db } from "../../db/client.js";
import { writeAudit } from "../../utils/audit.js";
import { STOCK_CATEGORIES, STOCK_CATEGORY_KEYS, classifyStockItem, ensureStockCategorySchema, normalizeStockCategory, stockCategoryLabel } from "../../utils/stockCategory.js";

export default function registerLocationsRoutes(app, ctx) {
  const { getBinByCodeAtLocation, getLocationByCode, getOnHand, getPartByCode, requireRoles } = ctx;

  // Stock locations
  // GET /api/stock/locations?active=1
  app.get("/locations", async (req, reply) => {
    const onlyActive = String(req.query?.active || "1").trim() !== "0";
    const rows = db.prepare(`
      SELECT id, location_code, location_name, active, created_at
      FROM stock_locations
      WHERE (? = 0 OR active = 1)
      ORDER BY location_code ASC
      LIMIT 300
    `).all(onlyActive ? 1 : 0).map((r) => ({
      ...r,
      active: Number(r.active || 0),
    }));
    return reply.send({ ok: true, rows });
  });

  // POST /api/stock/locations
  // Body: { location_code, location_name?, active? }
  app.post("/locations", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const location_code = String(req.body?.location_code || "").trim().toUpperCase();
    const location_name = String(req.body?.location_name || "").trim() || null;
    const active = req.body?.active === 0 || req.body?.active === false ? 0 : 1;
    if (!location_code) return reply.code(400).send({ error: "location_code is required" });

    const existing = getLocationByCode.get(location_code);
    if (existing) {
      db.prepare(`
        UPDATE stock_locations
        SET location_name = COALESCE(?, location_name),
            active = ?
        WHERE id = ?
      `).run(location_name, active, Number(existing.id));
      return reply.send({ ok: true, id: Number(existing.id), location_code, updated: true });
    }

    const ins = db.prepare(`
      INSERT INTO stock_locations (location_code, location_name, active)
      VALUES (?, ?, ?)
    `).run(location_code, location_name, active);

    return reply.send({ ok: true, id: Number(ins.lastInsertRowid), location_code, created: true });
  });

  // Stock bins (location-scoped)
  // GET /api/stock/bins?location_code=&active=1
  app.get("/bins", async (req, reply) => {
    const location_code = String(req.query?.location_code || "").trim().toUpperCase();
    const onlyActive = String(req.query?.active || "1").trim() !== "0";
    const where = [];
    const params = [];
    if (location_code) {
      const loc = getLocationByCode.get(location_code);
      if (!loc) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
      where.push("b.location_id = ?");
      params.push(Number(loc.id));
    }
    if (onlyActive) where.push("b.active = 1");
    const rows = db.prepare(`
      SELECT
        b.id, b.location_id, b.bin_code, b.bin_name, b.active, b.created_at,
        l.location_code, l.location_name
      FROM stock_bins b
      JOIN stock_locations l ON l.id = b.location_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY l.location_code ASC, b.bin_code ASC
      LIMIT 500
    `).all(...params).map((r) => ({
      ...r,
      active: Number(r.active || 0),
    }));
    return reply.send({ ok: true, rows });
  });

  // POST /api/stock/bins
  // Body: { location_code, bin_code, bin_name?, active? }
  app.post("/bins", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const location_code = String(req.body?.location_code || "").trim().toUpperCase();
    const bin_code = String(req.body?.bin_code || "").trim().toUpperCase();
    const bin_name = String(req.body?.bin_name || "").trim() || null;
    const active = req.body?.active === 0 || req.body?.active === false ? 0 : 1;
    if (!location_code || !bin_code) return reply.code(400).send({ error: "location_code and bin_code are required" });
    const location = getLocationByCode.get(location_code);
    if (!location) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    const existing = getBinByCodeAtLocation.get(Number(location.id), bin_code);
    if (existing) {
      db.prepare(`
        UPDATE stock_bins
        SET bin_name = COALESCE(?, bin_name),
            active = ?
        WHERE id = ?
      `).run(bin_name, active, Number(existing.id));
      return reply.send({ ok: true, updated: true, id: Number(existing.id), location_code, bin_code });
    }
    const ins = db.prepare(`
      INSERT INTO stock_bins (location_id, bin_code, bin_name, active)
      VALUES (?, ?, ?, ?)
    `).run(Number(location.id), bin_code, bin_name, active);
    return reply.send({ ok: true, created: true, id: Number(ins.lastInsertRowid), location_code, bin_code });
  });

  // Inventory depth by part/location/bin
  // GET /api/stock/depth?part_code=&location_code=&bin_code=
  app.get("/depth", async (req, reply) => {
    const part_code = String(req.query?.part_code || "").trim();
    const location_code = String(req.query?.location_code || "").trim().toUpperCase();
    const bin_code = String(req.query?.bin_code || "").trim().toUpperCase();
    const where = [];
    const params = [];
    if (part_code) {
      where.push("p.part_code = ?");
      params.push(part_code);
    }
    if (location_code) {
      const loc = getLocationByCode.get(location_code);
      if (!loc) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
      where.push("COALESCE(sm.location_id, ?) = ?");
      params.push(Number(loc.id), Number(loc.id));
      if (bin_code) {
        const bin = getBinByCodeAtLocation.get(Number(loc.id), bin_code);
        if (!bin) return reply.code(404).send({ error: `bin_code not found at location ${location_code}: ${bin_code}` });
        where.push("COALESCE(sm.bin_id, ?) = ?");
        params.push(Number(bin.id), Number(bin.id));
      }
    }
    const rows = db.prepare(`
      SELECT
        p.id AS part_id,
        p.part_code,
        p.part_name,
        COALESCE(sm.location_id, l.id) AS location_id,
        COALESCE(l.location_code, 'UNSPECIFIED') AS location_code,
        COALESCE(l.location_name, 'Unspecified') AS location_name,
        sm.bin_id,
        COALESCE(b.bin_code, 'UNSPECIFIED') AS bin_code,
        COALESCE(b.bin_name, 'Unspecified') AS bin_name,
        COALESCE(SUM(sm.quantity), 0) AS on_hand,
        COALESCE((
          SELECT SUM(sr.quantity_reserved - sr.quantity_issued)
          FROM stock_reservations sr
          WHERE sr.part_id = p.id
            AND sr.status = 'active'
        ), 0) AS reserved
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      LEFT JOIN stock_locations l ON l.id = sm.location_id
      LEFT JOIN stock_bins b ON b.id = sm.bin_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      GROUP BY p.id, COALESCE(sm.location_id, l.id), sm.bin_id
      ORDER BY p.part_code ASC, location_code ASC, bin_code ASC
      LIMIT 2000
    `).all(...params).map((r) => {
      const on_hand = Number(r.on_hand || 0);
      const reserved = Number(r.reserved || 0);
      const on_order = Number(
        db.prepare(`
          SELECT COALESCE(SUM(CASE WHEN qty_requested > qty_received THEN (qty_requested - qty_received) ELSE 0 END), 0) AS qty
          FROM procurement_requisitions
          WHERE part_id = ?
            AND LOWER(status) IN ('approved', 'approved_all', 'po_ready', 'partially_received')
        `).get(Number(r.part_id || 0))?.qty || 0
      );
      return {
        ...r,
        on_hand: Number(on_hand.toFixed(2)),
        reserved: Number(reserved.toFixed(2)),
        on_order: Number(on_order.toFixed(2)),
        available: Number((on_hand - reserved).toFixed(2)),
      };
    });
    return reply.send({ ok: true, rows });
  });

  // Min-max policy upsert
  // POST /api/stock/min-max
  app.post("/min-max", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const part_code = String(req.body?.part_code || "").trim();
    const location_code = String(req.body?.location_code || "").trim().toUpperCase();
    const bin_code = String(req.body?.bin_code || "").trim().toUpperCase();
    if (!part_code || !location_code) return reply.code(400).send({ error: "part_code and location_code are required" });
    const part = getPartByCode.get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const location = getLocationByCode.get(location_code);
    if (!location) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    const bin = bin_code ? getBinByCodeAtLocation.get(Number(location.id), bin_code) : null;
    if (bin_code && !bin) return reply.code(404).send({ error: `bin_code not found at location ${location_code}: ${bin_code}` });
    const min_qty = Math.max(0, Number(req.body?.min_qty || 0));
    const max_qty = Math.max(min_qty, Number(req.body?.max_qty || min_qty));
    const reorder_qty = req.body?.reorder_qty != null ? Math.max(0, Number(req.body.reorder_qty || 0)) : null;
    const target_days = req.body?.target_days != null ? Math.max(0, Number(req.body.target_days || 0)) : null;
    db.prepare(`
      INSERT INTO stock_min_max (
        part_id, location_id, bin_id, min_qty, max_qty, reorder_qty, target_days, updated_by, updated_at
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      ON CONFLICT(part_id, location_id, bin_id) DO UPDATE SET
        min_qty = excluded.min_qty,
        max_qty = excluded.max_qty,
        reorder_qty = excluded.reorder_qty,
        target_days = excluded.target_days,
        updated_by = excluded.updated_by,
        updated_at = datetime('now')
    `).run(
      Number(part.id),
      Number(location.id),
      bin ? Number(bin.id) : null,
      min_qty,
      max_qty,
      reorder_qty,
      target_days,
      String(req.headers["x-user-name"] || "session-user")
    );
    return reply.send({ ok: true, part_code, location_code, bin_code: bin ? bin.bin_code : null, min_qty, max_qty, reorder_qty, target_days });
  });

  // GET /api/stock/min-max?part_code=&location_code=&bin_code=
  app.get("/min-max", async (req, reply) => {
    const part_code = String(req.query?.part_code || "").trim();
    const location_code = String(req.query?.location_code || "").trim().toUpperCase();
    const bin_code = String(req.query?.bin_code || "").trim().toUpperCase();
    const where = [];
    const params = [];
    if (part_code) {
      where.push("p.part_code = ?");
      params.push(part_code);
    }
    if (location_code) {
      where.push("l.location_code = ?");
      params.push(location_code);
    }
    if (bin_code) {
      where.push("UPPER(COALESCE(b.bin_code, '')) = ?");
      params.push(bin_code);
    }
    const rows = db.prepare(`
      SELECT
        mm.id,
        p.part_code,
        p.part_name,
        l.location_code,
        l.location_name,
        b.bin_code,
        b.bin_name,
        mm.min_qty,
        mm.max_qty,
        mm.reorder_qty,
        mm.target_days,
        mm.updated_by,
        mm.updated_at
      FROM stock_min_max mm
      JOIN parts p ON p.id = mm.part_id
      JOIN stock_locations l ON l.id = mm.location_id
      LEFT JOIN stock_bins b ON b.id = mm.bin_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY p.part_code ASC, l.location_code ASC, COALESCE(b.bin_code, '') ASC
      LIMIT 1000
    `).all(...params).map((r) => ({
      ...r,
      min_qty: Number(r.min_qty || 0),
      max_qty: Number(r.max_qty || 0),
      reorder_qty: r.reorder_qty == null ? null : Number(r.reorder_qty || 0),
      target_days: r.target_days == null ? null : Number(r.target_days || 0),
    }));
    return reply.send({ ok: true, rows });
  });

  // Replenishment suggestions from min-max and current on-hand
  // GET /api/stock/replenishment-suggestions?location_code=&bin_code=
  app.get("/replenishment-suggestions", async (req, reply) => {
    const location_code = String(req.query?.location_code || "").trim().toUpperCase();
    const bin_code = String(req.query?.bin_code || "").trim().toUpperCase();
    const where = [];
    const params = [];
    if (location_code) {
      where.push("l.location_code = ?");
      params.push(location_code);
    }
    if (bin_code) {
      where.push("UPPER(COALESCE(b.bin_code, '')) = ?");
      params.push(bin_code);
    }
    const policyRows = db.prepare(`
      SELECT
        mm.id,
        mm.part_id,
        p.part_code,
        p.part_name,
        l.id AS location_id,
        l.location_code,
        b.id AS bin_id,
        b.bin_code,
        mm.min_qty,
        mm.max_qty,
        mm.reorder_qty
      FROM stock_min_max mm
      JOIN parts p ON p.id = mm.part_id
      JOIN stock_locations l ON l.id = mm.location_id
      LEFT JOIN stock_bins b ON b.id = mm.bin_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY p.part_code ASC
      LIMIT 1500
    `).all(...params);
    const rows = policyRows.map((r) => {
      const on_hand = Number(
        db.prepare(`
          SELECT COALESCE(SUM(quantity), 0) AS q
          FROM stock_movements
          WHERE part_id = ?
            AND COALESCE(location_id, ?) = ?
            AND COALESCE(bin_id, 0) = COALESCE(?, 0)
        `).get(Number(r.part_id), Number(r.location_id), Number(r.location_id), r.bin_id == null ? null : Number(r.bin_id))?.q || 0
      );
      const need = Math.max(0, Number(r.min_qty || 0) - on_hand);
      const target = Number(r.reorder_qty || 0) > 0 ? Number(r.reorder_qty || 0) : Math.max(0, Number(r.max_qty || 0) - on_hand);
      return {
        part_code: r.part_code,
        part_name: r.part_name,
        location_code: r.location_code,
        bin_code: r.bin_code || null,
        min_qty: Number(r.min_qty || 0),
        max_qty: Number(r.max_qty || 0),
        on_hand: Number(on_hand.toFixed(2)),
        shortage_qty: Number(need.toFixed(2)),
        suggested_order_qty: Number(target.toFixed(2)),
        needs_replenishment: need > 0,
      };
    }).filter((r) => r.needs_replenishment);
    return reply.send({ ok: true, rows });
  });

  // Set minimum stock for a specific part
  // POST /api/stock/part-minimum
  // Body: { part_code, min_stock }
  // GET /api/stock/categories — the stock categories stores can assign.
  app.get("/categories", async () => ({ ok: true, categories: STOCK_CATEGORIES }));

  // POST /api/stock/part-category { part_code, category } — set an item's stock category by hand.
  // category "auto" hands the item back to the automatic rules.
  app.post("/part-category", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const part_code = String(req.body?.part_code || "").trim();
    const raw = String(req.body?.category || "").trim().toLowerCase();
    if (!part_code) return reply.code(400).send({ error: "part_code is required" });
    const category = raw === "auto" ? "auto" : normalizeStockCategory(raw);
    if (!category) {
      return reply.code(400).send({ error: `category must be one of: ${STOCK_CATEGORY_KEYS.join(", ")} or auto` });
    }
    ensureStockCategorySchema(db);
    const part = db.prepare(`SELECT id, part_code, part_name, stock_category, stock_category_source FROM parts WHERE part_code = ?`).get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const next = category === "auto" ? classifyStockItem(part) : category;
    const source = category === "auto" ? "auto" : "manual";
    db.prepare(`UPDATE parts SET stock_category = ?, stock_category_source = ? WHERE id = ?`).run(next, source, part.id);
    writeAudit(db, req, {
      module: "stock",
      action: "set_part_category",
      entity_type: "part",
      entity_id: part_code,
      payload: { before: part.stock_category || null, after: next, source },
    });
    return reply.send({ ok: true, part_code, stock_category: next, stock_category_label: stockCategoryLabel(next), stock_category_source: source });
  });

  app.post("/part-minimum", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const part_code = String(req.body?.part_code || "").trim();
    const min_stock = Number(req.body?.min_stock ?? NaN);
    if (!part_code) return reply.code(400).send({ error: "part_code is required" });
    if (!Number.isFinite(min_stock) || min_stock < 0) {
      return reply.code(400).send({ error: "min_stock must be a valid number >= 0" });
    }
    const part = getPartByCode.get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const before = Number(part.min_stock || 0);
    db.prepare(`UPDATE parts SET min_stock = ? WHERE id = ?`).run(Number(min_stock.toFixed(2)), Number(part.id));
    const on_hand = Number(getOnHand.get(part.id)?.on_hand || 0);

    writeAudit(db, req, {
      module: "stock",
      action: "set_part_minimum",
      entity_type: "part",
      entity_id: part_code,
      payload: { min_stock_before: before, min_stock_after: Number(min_stock.toFixed(2)), on_hand },
    });

    return reply.send({
      ok: true,
      part_code,
      min_stock_before: before,
      min_stock_after: Number(min_stock.toFixed(2)),
      on_hand,
      below_min: on_hand < Number(min_stock.toFixed(2)),
    });
  });
}
