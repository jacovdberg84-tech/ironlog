// IRONLOG/api/routes/stock/lube.routes.js — Lube stock on hand, month stock, minimums, receipts and issues.
// Registered by routes/stock.routes.js; shared helpers arrive through ctx.
import { db } from "../../db/client.js";
import { fetchLubeMonthStockSnapshot } from "../../utils/lubeMonthStock.js";
import { resolveLogCostCenterCode } from "../../utils/costAllocation.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerLubeRoutes(app, ctx) {
  const {
    getAssetByCode,
    getBinByCodeAtLocation,
    getLocationByCode,
    getOnHand,
    getPartByCode,
    hasColumn,
    normalizeOilTypeInput,
    requireRoles,
  } = ctx;

  // Lube stock lookup
  // GET /api/stock/lube-onhand?q=&location_code=
  app.get("/lube-onhand", async (req, reply) => {
    const q = String(req.query?.q || "").trim();
    const location_code = String(req.query?.location_code || "").trim().toUpperCase();
    const location = location_code ? getLocationByCode.get(location_code) : null;
    if (location_code && !location) {
      return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    }
    const where = [
      "(" +
        "LOWER(IFNULL(p.part_code, '')) LIKE '%oil%' OR " +
        "LOWER(IFNULL(p.part_name, '')) LIKE '%oil%' OR " +
        "LOWER(IFNULL(p.part_code, '')) LIKE '%lube%' OR " +
        "LOWER(IFNULL(p.part_name, '')) LIKE '%lube%' OR " +
        "LOWER(IFNULL(p.part_code, '')) LIKE '%grease%' OR " +
        "LOWER(IFNULL(p.part_name, '')) LIKE '%grease%'" +
      ")",
    ];
    const params = [];

    if (q) {
      where.push("(p.part_code LIKE ? OR p.part_name LIKE ?)");
      params.push(`%${q}%`, `%${q}%`);
    }

    // Location filter:
    // - When a location is provided, include that location AND legacy rows with NULL location_id
    //   (older data before location tracking).
    const joinMovements = location
      ? "LEFT JOIN stock_movements sm ON sm.part_id = p.id AND (sm.location_id = ? OR sm.location_id IS NULL)"
      : "LEFT JOIN stock_movements sm ON sm.part_id = p.id";
    if (location) params.unshift(Number(location.id));

    const rows = db.prepare(`
      SELECT
        p.id,
        p.part_code,
        p.part_name,
        p.min_stock,
        IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      ${joinMovements}
      WHERE ${where.join(" AND ")}
      GROUP BY p.id
      ORDER BY on_hand ASC, p.part_code ASC
      LIMIT 40
    `).all(...params).map((r) => ({
      id: Number(r.id),
      part_code: r.part_code,
      part_name: r.part_name,
      min_stock: Number(r.min_stock || 0),
      on_hand: Number(r.on_hand || 0),
      below_min: Number(r.on_hand || 0) < Number(r.min_stock || 0),
    }));

    const exact = q
      ? rows.find((r) => String(r.part_code || "").toLowerCase() === q.toLowerCase()) || null
      : null;

    return reply.send({ ok: true, q, location_code: location ? location.location_code : null, exact, rows });
  });

  // GET /api/stock/lube-month-stock?month=YYYY-MM&location_code=
  // Store (parts) opening/closing quantities for lube/oil/grease SKUs from cumulative stock_movements.
  app.get("/lube-month-stock", async (req, reply) => {
    const month = String(req.query?.month || "").trim();
    const location_code = String(req.query?.location_code || "").trim().toUpperCase();
    if (!/^\d{4}-\d{2}$/.test(month)) {
      return reply.code(400).send({ error: "month must be YYYY-MM" });
    }
    const location = location_code ? getLocationByCode.get(location_code) : null;
    if (location_code && !location) {
      return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    }
    try {
      const snap = fetchLubeMonthStockSnapshot(db, {
        month,
        location_id: location ? Number(location.id) : null,
      });
      return reply.send({
        ok: true,
        ...snap,
        location_code: location ? location.location_code : null,
      });
    } catch (e) {
      return reply.code(400).send({ error: String(e.message || e) });
    }
  });

  // Set minimum stock for lube/oil items
  // POST /api/stock/lube-minimums
  // Body: { min_stock?: number } default 210
  app.post("/lube-minimums", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const min_stock_input = Number(req.body?.min_stock ?? 210);
    if (!Number.isFinite(min_stock_input) || min_stock_input < 0) {
      return reply.code(400).send({ error: "min_stock must be a valid number >= 0" });
    }
    const min_stock = Number(min_stock_input.toFixed(2));

    const lubeParts = db.prepare(`
      SELECT id, part_code, part_name, min_stock
      FROM parts
      WHERE (
        LOWER(IFNULL(part_code, '')) LIKE '%oil%' OR
        LOWER(IFNULL(part_name, '')) LIKE '%oil%' OR
        LOWER(IFNULL(part_code, '')) LIKE '%lube%' OR
        LOWER(IFNULL(part_name, '')) LIKE '%lube%' OR
        LOWER(IFNULL(part_code, '')) LIKE '%grease%' OR
        LOWER(IFNULL(part_name, '')) LIKE '%grease%'
      )
      AND LOWER(IFNULL(part_name, '')) NOT LIKE '%filter%'
      ORDER BY part_code ASC
      LIMIT 500
    `).all();

    const upd = db.prepare(`UPDATE parts SET min_stock = ? WHERE id = ?`);
    const tx = db.transaction(() => {
      for (const p of lubeParts) {
        upd.run(min_stock, Number(p.id));
      }
    });
    tx();

    const rows = lubeParts.map((p) => ({
      part_code: p.part_code,
      part_name: p.part_name,
      min_stock_before: Number(p.min_stock || 0),
      min_stock_after: min_stock,
    }));

    writeAudit(db, req, {
      module: "stock",
      action: "set_lube_minimums",
      entity_type: "parts",
      payload: {
        min_stock,
        updated_count: rows.length,
      },
    });

    return reply.send({
      ok: true,
      min_stock,
      updated_count: rows.length,
      rows,
    });
  });

  // Manual lube log entry
  // POST /api/stock/lube-log
  // Body: { asset_code, log_date, oil_type?, quantity }
  app.post("/lube-log", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "artisan", "operator"])) return;
    const body = req.body || {};
    const asset_code = String(body.asset_code || "").trim();
    const log_date =
      body.log_date != null && String(body.log_date).trim() !== ""
        ? String(body.log_date).trim()
        : new Date().toISOString().slice(0, 10);
    const oil_type = normalizeOilTypeInput(body.oil_type, null);
    const quantity = Number(body.quantity ?? 0);

    if (!asset_code) return reply.code(400).send({ error: "asset_code is required" });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(log_date)) {
      return reply.code(400).send({ error: "log_date must be YYYY-MM-DD" });
    }
    if (!Number.isFinite(quantity) || quantity <= 0) {
      return reply.code(400).send({ error: "quantity must be > 0" });
    }

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });
    const cost_center_code = resolveLogCostCenterCode(db, asset.id, body.cost_center_code);

    const ins = db.prepare(`
      INSERT INTO oil_logs (asset_id, log_date, oil_type, quantity, cost_center_code)
      VALUES (?, ?, ?, ?, ?)
    `).run(asset.id, log_date, oil_type, quantity, cost_center_code);

    writeAudit(db, req, {
      module: "lube",
      action: "manual_log",
      entity_type: "asset",
      entity_id: asset_code,
      payload: { log_date, oil_type, quantity, cost_center_code },
    });

    return reply.send({
      ok: true,
      id: Number(ins.lastInsertRowid),
      asset_code,
      log_date,
      oil_type,
      quantity,
      cost_center_code,
    });
  });

  // Lube issue from stock (supports stock-number flow)
  // POST /api/stock/lube-issue
  // Body: { part_code, quantity, asset_code?, log_date?, oil_type?, notes?, location_code? }
  app.post("/lube-issue", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores", "artisan", "operator"])) return;
    const body = req.body || {};
    const part_code = String(body.part_code || "").trim();
    const asset_code = String(body.asset_code || "").trim();
    const quantity = Number(body.quantity ?? 0);
    const log_date =
      body.log_date != null && String(body.log_date).trim() !== ""
        ? String(body.log_date).trim()
        : new Date().toISOString().slice(0, 10);
    const oil_type = normalizeOilTypeInput(body.oil_type, null);
    const notes =
      body.notes != null && String(body.notes).trim() !== ""
        ? String(body.notes).trim()
        : null;
    const location_code = String(body.location_code || "").trim().toUpperCase();
    const bin_code = String(body.bin_code || "").trim().toUpperCase();

    if (!part_code) return reply.code(400).send({ error: "part_code is required" });
    if (!Number.isFinite(quantity) || quantity <= 0) {
      return reply.code(400).send({ error: "quantity must be > 0" });
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(log_date)) {
      return reply.code(400).send({ error: "log_date must be YYYY-MM-DD" });
    }

    const part = getPartByCode.get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const location = location_code ? getLocationByCode.get(location_code) : null;
    if (location_code && !location) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    const bin = (location && bin_code) ? getBinByCodeAtLocation.get(Number(location.id), bin_code) : null;
    if (bin_code && !location_code) return reply.code(400).send({ error: "location_code is required when bin_code is provided" });
    if (location && bin_code && !bin) return reply.code(404).send({ error: `bin_code not found at location ${location_code}: ${bin_code}` });

    let asset = null;
    if (asset_code) {
      asset = getAssetByCode.get(asset_code);
      if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });
    }

    const onHand = Number(getOnHand.get(part.id)?.on_hand || 0);
    if (onHand < quantity) {
      return reply.code(409).send({
        error: "insufficient stock",
        part_code,
        on_hand: onHand,
        requested: quantity,
      });
    }

    const tx = db.transaction(() => {
      const reference = asset ? `lube_issue:asset:${asset.id}` : `lube_issue:stock`;
      const mv = db.prepare(`
        INSERT INTO stock_movements (part_id, quantity, movement_type, reference, location_id, bin_id)
        VALUES (?, ?, 'out', ?, ?, ?)
      `).run(part.id, -Math.abs(quantity), reference, location ? Number(location.id) : null, bin ? Number(bin.id) : null);

      let lube_log_id = null;
      if (asset) {
        const cost_center_code = resolveLogCostCenterCode(db, asset.id, body.cost_center_code);
        const oilUnitCost = Number(part.unit_cost || 0) > 0 ? Number(part.unit_cost) : null;
        const oilLogColumns = ["asset_id", "log_date", "oil_type", "quantity", "cost_center_code"];
        const oilLogValues = [
          asset.id,
          log_date,
          normalizeOilTypeInput(oil_type, part.part_code || null),
          quantity,
          cost_center_code,
        ];
        if (hasColumn("oil_logs", "unit_cost")) {
          oilLogColumns.push("unit_cost");
          oilLogValues.push(oilUnitCost);
        }
        if (hasColumn("oil_logs", "part_id")) {
          oilLogColumns.push("part_id");
          oilLogValues.push(part.id);
        }
        const lg = db.prepare(`
          INSERT INTO oil_logs (${oilLogColumns.join(", ")})
          VALUES (${oilLogColumns.map(() => "?").join(", ")})
        `).run(...oilLogValues);
        lube_log_id = Number(lg.lastInsertRowid);
      }

      return {
        movement_id: Number(mv.lastInsertRowid),
        lube_log_id,
      };
    });

    const result = tx();
    const on_hand_after = Number(getOnHand.get(part.id)?.on_hand || 0);

    writeAudit(db, req, {
      module: "lube",
      action: "stock_issue",
      entity_type: "part",
      entity_id: part_code,
      payload: {
        asset_code: asset ? asset.asset_code : null,
        quantity,
        log_date,
        oil_type,
        notes,
        location_code: location ? location.location_code : null,
        on_hand_before: onHand,
        on_hand_after,
      },
    });

    return reply.send({
      ok: true,
      part_code,
      part_name: part.part_name,
      quantity,
      movement_id: result.movement_id,
      lube_log_id: result.lube_log_id,
      linked_asset_code: asset ? asset.asset_code : null,
      linked_asset_name: asset ? asset.asset_name : null,
      location_code: location ? location.location_code : null,
      bin_code: bin ? bin.bin_code : null,
      on_hand_before: onHand,
      on_hand_after,
      message: asset
        ? "Lube issued from stock and linked to asset log"
        : "Lube issued from stock (not linked to asset log)",
    });
  });
}
