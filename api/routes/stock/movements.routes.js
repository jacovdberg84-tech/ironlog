// IRONLOG/api/routes/stock/movements.routes.js — Stock movements, store allocations.
// Registered by routes/stock.routes.js; shared helpers arrive through ctx.
import { db } from "../../db/client.js";
import { normalizeMdmCode, validateAgainstMdmPolicy, validatePartGovernanceOptional } from "../../utils/masterdataGovernance.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerMovementsRoutes(app, ctx) {
  const {
    getAssetByCode,
    getBinByCodeAtLocation,
    getFxRate,
    getLocationByCode,
    getOnHand,
    getPartByCode,
    getRole,
    getSiteCode,
    getWoById,
    insertAlloc,
    insertGenericMove,
    insertMove,
    insertPart,
    requireRoles,
    unitCostToUsd,
    updatePartUnitCostUsd,
  } = ctx;

  // Manual stock movement entry
  // POST /api/stock/movement
  // Body: { part_code, quantity, movement_type: in|out|adjust, reference?, location_code?,
  //         part_name?, create_if_missing?, unit_cost?, cost_currency? (USD|ZAR|MZN) }
  app.post("/movement", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores", "storeman"])) return;
    const site = getSiteCode(req);
    const body = req.body || {};
    const part_code = String(body.part_code || "").trim();
    const movement_type = String(body.movement_type || "").trim().toLowerCase();
    const qtyIn = Number(body.quantity ?? 0);
    const reference = String(body.reference || "manual_entry").trim() || "manual_entry";
    const location_code = String(body.location_code || "").trim().toUpperCase();
    const bin_code = String(body.bin_code || "").trim().toUpperCase();
    const part_name = body.part_name != null ? String(body.part_name || "").trim() : "";
    const create_if_missing = body.create_if_missing === true || body.create_if_missing === 1;

    const department_code =
      body.department_code != null && String(body.department_code).trim() !== ""
        ? normalizeMdmCode(body.department_code)
        : null;
    const default_supplier_code =
      body.default_supplier_code != null && String(body.default_supplier_code).trim() !== ""
        ? normalizeMdmCode(body.default_supplier_code)
        : null;
    const data_owner_username =
      body.data_owner_username != null && String(body.data_owner_username).trim() !== ""
        ? String(body.data_owner_username).trim()
        : null;

    const rawCost = body.unit_cost != null && body.unit_cost !== "" ? Number(body.unit_cost) : null;
    const cost_currency = String(body.cost_currency || "USD").trim().toUpperCase();

    if (!part_code) return reply.code(400).send({ error: "part_code is required" });
    if (!["in", "out", "adjust"].includes(movement_type)) {
      return reply.code(400).send({ error: "movement_type must be one of: in, out, adjust" });
    }
    if (!Number.isFinite(qtyIn) || qtyIn === 0) {
      return reply.code(400).send({ error: "quantity must be a non-zero number" });
    }

    let unit_cost_usd = null;
    let cost_input = null;
    let cost_curr_out = null;
    if (movement_type === "in" && rawCost != null && Number.isFinite(rawCost) && rawCost > 0) {
      if (!["USD", "ZAR", "MZN"].includes(cost_currency)) {
        return reply.code(400).send({ error: "cost_currency must be USD, ZAR, or MZN" });
      }
      cost_input = rawCost;
      cost_curr_out = cost_currency;
      unit_cost_usd = unitCostToUsd(rawCost, cost_currency);
      if (unit_cost_usd == null || !Number.isFinite(unit_cost_usd)) {
        return reply.code(400).send({ error: "could not convert unit cost to USD — check FX rates" });
      }
    }

    let part = getPartByCode.get(part_code);
    if (!part) {
      // Allow receiving stock IN for brand-new stock numbers
      if (movement_type === "in" && create_if_missing) {
        if (!part_name) return reply.code(400).send({ error: "part_name is required when creating a new stock item" });
        const intakeObj = { department_code, default_supplier_code, data_owner_username };
        const polP = validateAgainstMdmPolicy(site, "part_stock_intake", intakeObj);
        if (!polP.ok) {
          return reply.code(400).send({ error: `missing required fields: ${polP.missing.join(", ")}` });
        }
        const pv = validatePartGovernanceOptional(site, intakeObj);
        if (!pv.ok) return reply.code(400).send({ error: pv.error });
        try {
          const uc = unit_cost_usd != null ? Number(unit_cost_usd.toFixed(6)) : 0;
          insertPart.run(part_code, part_name, uc, department_code, default_supplier_code, data_owner_username);
        } catch (e) {
          // In case of race, re-read
        }
        part = getPartByCode.get(part_code);
      }
    }
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const location = location_code ? getLocationByCode.get(location_code) : null;
    if (location_code && !location) {
      return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    }
    const bin = (location && bin_code) ? getBinByCodeAtLocation.get(Number(location.id), bin_code) : null;
    if (bin_code && !location_code) return reply.code(400).send({ error: "location_code is required when bin_code is provided" });
    if (location && bin_code && !bin) return reply.code(404).send({ error: `bin_code not found at location ${location_code}: ${bin_code}` });
    const cost_center_code =
      body.cost_center_code != null && String(body.cost_center_code).trim() !== ""
        ? String(body.cost_center_code).trim()
        : null;

    let qty = qtyIn;
    if (movement_type === "in") qty = Math.abs(qtyIn);
    if (movement_type === "out") qty = -Math.abs(qtyIn);
    if (movement_type === "adjust") qty = qtyIn;

    if (movement_type === "adjust") {
      const reqRole = getRole(req);
      const reqUser = String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
      const approvalPayload = JSON.stringify({
        part_code,
        quantity: qty,
        reference,
        location_code: location ? location.location_code : null,
      });
      const ins = db.prepare(`
        INSERT INTO approval_requests (
          module, action, entity_type, entity_id, status, payload_json, requested_by, requested_role
        )
        VALUES ('stock', 'adjust_movement', 'part', ?, 'pending', ?, ?, ?)
      `).run(part_code, approvalPayload, reqUser, reqRole);
      const request_id = Number(ins.lastInsertRowid);

      writeAudit(db, req, {
        module: "stock",
        action: "adjust_request",
        entity_type: "part",
        entity_id: part_code,
        payload: { request_id, quantity: qty, reference },
      });

      return reply.send({
        ok: true,
        pending_approval: true,
        request_id,
        part_code,
        movement_type,
        quantity: qty,
        reference,
        location_code: location ? location.location_code : null,
        message: "Stock adjustment submitted for approval",
      });
    }

    const onHand = Number(getOnHand.get(part.id)?.on_hand || 0);
    if (movement_type === "out" && onHand < Math.abs(qty)) {
      return reply.code(409).send({
        error: "insufficient stock",
        part_code,
        on_hand: onHand,
        requested: Math.abs(qty),
      });
    }

    const tx = db.transaction(() => {
      const ins = insertGenericMove.run(
        part.id,
        qty,
        movement_type,
        reference,
        location ? Number(location.id) : null,
        bin ? Number(bin.id) : null,
        cost_center_code,
        unit_cost_usd != null ? Number(unit_cost_usd.toFixed(6)) : null,
        cost_curr_out,
        cost_input,
      );
      const mid = Number(ins.lastInsertRowid);
      if (movement_type === "in" && unit_cost_usd != null) {
        updatePartUnitCostUsd.run(Number(unit_cost_usd.toFixed(6)), part.id);
      }
      return mid;
    });

    const movement_id = tx();
    const on_hand_after = Number(getOnHand.get(part.id)?.on_hand || 0);
    const line_value_usd =
      movement_type === "in" && unit_cost_usd != null
        ? Number((Math.abs(qty) * unit_cost_usd).toFixed(2))
        : null;

    writeAudit(db, req, {
      module: "stock",
      action: "manual_movement",
      entity_type: "part",
      entity_id: part_code,
      payload: {
        movement_type,
        quantity: qty,
        reference,
        on_hand_before: onHand,
        on_hand_after,
        unit_cost_usd,
        cost_currency: cost_curr_out,
        cost_input,
        line_value_usd,
      },
    });

    return reply.send({
      ok: true,
      movement_id,
      part_code,
      movement_type,
      quantity: qty,
      on_hand_before: onHand,
      on_hand_after,
      reference,
      location_code: location ? location.location_code : null,
      bin_code: bin ? bin.bin_code : null,
      cost_center_code,
      unit_cost_usd: unit_cost_usd != null ? Number(unit_cost_usd.toFixed(6)) : null,
      cost_currency: cost_curr_out,
      cost_input,
      line_value_usd,
      fx_rates: {
        zar_per_usd: getFxRate("zar_per_usd", 18.5),
        mzn_per_usd: getFxRate("mzn_per_usd", 64),
      },
    });
  });

  // Allocate stores/parts to an asset or work order
  // POST /api/stock/allocate
  // Body: { part_code, quantity, asset_code?, work_order_id?, allocation_date?, issued_by?, notes?, location_code? }
  app.post("/allocate", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const body = req.body || {};
    const part_code = String(body.part_code || "").trim();
    const quantity = Number(body.quantity ?? 0);
    const asset_code = String(body.asset_code || "").trim();
    const work_order_id = body.work_order_id != null ? Number(body.work_order_id) : null;
    const allocation_date =
      body.allocation_date != null && String(body.allocation_date).trim() !== ""
        ? String(body.allocation_date).trim()
        : new Date().toISOString().slice(0, 10);
    const issued_by =
      body.issued_by != null && String(body.issued_by).trim() !== ""
        ? String(body.issued_by).trim()
        : null;
    const notes =
      body.notes != null && String(body.notes).trim() !== ""
        ? String(body.notes).trim()
        : null;
    const location_code = String(body.location_code || "").trim().toUpperCase();
    const bin_code = String(body.bin_code || "").trim().toUpperCase();
    const cost_center_code =
      body.cost_center_code != null && String(body.cost_center_code).trim() !== ""
        ? String(body.cost_center_code).trim()
        : null;

    if (!part_code || !Number.isFinite(quantity) || quantity <= 0) {
      return reply.code(400).send({ error: "part_code and quantity (>0) are required" });
    }
    if (!asset_code && (!Number.isFinite(work_order_id) || work_order_id <= 0)) {
      return reply.code(400).send({ error: "Provide asset_code or work_order_id" });
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(allocation_date)) {
      return reply.code(400).send({ error: "allocation_date must be YYYY-MM-DD" });
    }

    const part = getPartByCode.get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const location = location_code ? getLocationByCode.get(location_code) : null;
    if (location_code && !location) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    const bin = (location && bin_code) ? getBinByCodeAtLocation.get(Number(location.id), bin_code) : null;
    if (bin_code && !location_code) return reply.code(400).send({ error: "location_code is required when bin_code is provided" });
    if (location && bin_code && !bin) return reply.code(404).send({ error: `bin_code not found at location ${location_code}: ${bin_code}` });

    let wo = null;
    if (Number.isFinite(work_order_id) && work_order_id > 0) {
      wo = getWoById.get(work_order_id);
      if (!wo) return reply.code(404).send({ error: `work_order not found: ${work_order_id}` });
    }

    let asset = null;
    if (asset_code) {
      asset = getAssetByCode.get(asset_code);
      if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });
    }

    const resolvedAssetId = wo ? Number(wo.asset_id) : Number(asset.id);
    if (wo && asset && Number(asset.id) !== Number(wo.asset_id)) {
      return reply.code(409).send({ error: "asset_code does not match work_order asset" });
    }

    const onHand = Number(getOnHand.get(part.id)?.on_hand || 0);
    if (onHand < quantity) {
      return reply.code(409).send({
        error: "insufficient stock",
        part_code,
        on_hand: onHand,
        requested: quantity
      });
    }

    const tx = db.transaction(() => {
      const reference = wo ? `work_order:${wo.id}` : `asset:${resolvedAssetId}:stores`;
      insertMove.run(
        part.id,
        -Math.abs(quantity),
        reference,
        location ? Number(location.id) : null,
        bin ? Number(bin.id) : null,
        cost_center_code
      );
      const ins = insertAlloc.run(
        resolvedAssetId,
        wo ? Number(wo.id) : null,
        part.id,
        quantity,
        allocation_date,
        issued_by,
        notes,
        location ? Number(location.id) : null,
        bin ? Number(bin.id) : null,
        cost_center_code
      );
      return Number(ins.lastInsertRowid);
    });

    const allocation_id = tx();

    writeAudit(db, req, {
      module: "stock",
      action: "allocate",
      entity_type: "work_order",
      entity_id: wo ? Number(wo.id) : resolvedAssetId,
      payload: {
        part_code,
        quantity,
        asset_id: resolvedAssetId,
        work_order_id: wo ? Number(wo.id) : null,
        location_code: location ? location.location_code : null,
        bin_code: bin ? bin.bin_code : null,
        cost_center_code,
      },
    });

    return reply.send({
      ok: true,
      allocation_id,
      asset_id: resolvedAssetId,
      work_order_id: wo ? Number(wo.id) : null,
      part_code,
      unit_cost_usd: Number(part.unit_cost || 0),
      line_value_usd: Number((Number(part.unit_cost || 0) * Number(quantity || 0)).toFixed(2)),
      issued: quantity,
      location_code: location ? location.location_code : null,
      bin_code: bin ? bin.bin_code : null,
      cost_center_code,
      on_hand_before: onHand,
      on_hand_after: onHand - quantity
    });
  });

  // List allocations
  // GET /api/stock/allocations?asset_code=&part_code=&start=&end=
  app.get("/allocations", async (req, reply) => {
    const asset_code = String(req.query?.asset_code || "").trim();
    const part_code = String(req.query?.part_code || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();

    const where = [];
    const params = [];
    if (asset_code) {
      where.push("a.asset_code = ?");
      params.push(asset_code);
    }
    if (part_code) {
      where.push("p.part_code = ?");
      params.push(part_code);
    }
    if (start && /^\d{4}-\d{2}-\d{2}$/.test(start)) {
      where.push("sa.allocation_date >= ?");
      params.push(start);
    }
    if (end && /^\d{4}-\d{2}-\d{2}$/.test(end)) {
      where.push("sa.allocation_date <= ?");
      params.push(end);
    }

    const rows = db.prepare(`
      SELECT
        sa.id,
        sa.allocation_date,
        a.asset_code,
        a.asset_name,
        sa.work_order_id,
        p.part_code,
        p.part_name,
        p.unit_cost,
        l.location_code,
        l.location_name,
        b.bin_code,
        b.bin_name,
        sa.cost_center_code,
        sa.quantity,
        sa.issued_by,
        sa.notes,
        sa.created_at
      FROM store_allocations sa
      JOIN assets a ON a.id = sa.asset_id
      JOIN parts p ON p.id = sa.part_id
      LEFT JOIN stock_locations l ON l.id = sa.location_id
      LEFT JOIN stock_bins b ON b.id = sa.bin_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY sa.id DESC
      LIMIT 500
    `).all(...params);

    return reply.send({ ok: true, rows });
  });
}
