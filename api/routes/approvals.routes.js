import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";

function hasColumn(table, col) {
  const rows = db.prepare(`PRAGMA table_info(${table})`).all();
  return rows.some((r) => String(r.name) === col);
}

function ensureColumn(table, colName, colDef) {
  if (!hasColumn(table, colName)) {
    db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
  }
}

function getRole(req) {
  return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
}

function getUser(req) {
  return String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
}

function requireRoles(req, reply, roles) {
  const role = getRole(req);
  if (!roles.includes(role)) {
    reply.code(403).send({ error: `role '${role || "unknown"}' not allowed` });
    return false;
  }
  return true;
}

function parsePayload(payloadJson) {
  try {
    return payloadJson ? JSON.parse(payloadJson) : null;
  } catch {
    return null;
  }
}

function getAssetCurrentHours(assetId) {
  const fromAssetHours = db.prepare(`
    SELECT total_hours
    FROM asset_hours
    WHERE asset_id = ?
  `).get(assetId);
  const assetHours = fromAssetHours?.total_hours == null ? null : Number(fromAssetHours.total_hours);

  const fromDailyClosing = db.prepare(`
    SELECT closing_hours
    FROM daily_hours
    WHERE asset_id = ?
      AND closing_hours IS NOT NULL
    ORDER BY work_date DESC, id DESC
    LIMIT 1
  `).get(assetId);
  const dailyClosing = fromDailyClosing?.closing_hours == null ? null : Number(fromDailyClosing.closing_hours);

  if (assetHours != null && Number.isFinite(assetHours) && dailyClosing != null && Number.isFinite(dailyClosing)) {
    if (Math.abs(assetHours - dailyClosing) > 5000) return dailyClosing;
    return dailyClosing >= assetHours ? dailyClosing : assetHours;
  }
  if (dailyClosing != null && Number.isFinite(dailyClosing)) return dailyClosing;
  if (assetHours != null && Number.isFinite(assetHours)) return assetHours;

  const fromDailyHours = db.prepare(`
    SELECT COALESCE(SUM(hours_run), 0) AS total_hours
    FROM daily_hours
    WHERE asset_id = ?
      AND is_used = 1
      AND hours_run > 0
  `).get(assetId);
  return Number(fromDailyHours?.total_hours || 0);
}

function closeWorkOrderWithPayload(approvalId, payload, req) {
  const woId = Number(payload?.work_order_id || 0);
  if (!Number.isFinite(woId) || woId <= 0) {
    throw new Error(`approval ${approvalId}: invalid work_order_id`);
  }

  const wo = db.prepare(`
    SELECT id, status, source, reference_id, asset_id
    FROM work_orders
    WHERE id = ?
  `).get(woId);
  if (!wo) throw new Error(`approval ${approvalId}: work order not found`);
  if (String(wo.status || "").toLowerCase() === "closed") return { woId, already_closed: true };

  const completion_notes =
    payload?.completion_notes != null && String(payload.completion_notes).trim() !== ""
      ? String(payload.completion_notes).trim()
      : null;
  const artisan_name =
    payload?.artisan_name != null && String(payload.artisan_name).trim() !== ""
      ? String(payload.artisan_name).trim()
      : null;
  const supervisor_name =
    payload?.supervisor_name != null && String(payload.supervisor_name).trim() !== ""
      ? String(payload.supervisor_name).trim()
      : getUser(req);

  const isServiceWO = String(wo.source || "").toLowerCase() === "service";
  if (isServiceWO && !artisan_name) throw new Error("artisan_name is required for service work order");
  if (isServiceWO && !completion_notes) throw new Error("completion_notes is required for service work order");

  const closeWorkOrder = db.prepare(`
    UPDATE work_orders
    SET
      status='closed',
      completed_at = datetime('now'),
      closed_at = datetime('now'),
      completion_notes = COALESCE(?, completion_notes),
      artisan_name = COALESCE(?, artisan_name),
      artisan_signed_at = CASE WHEN ? IS NOT NULL THEN datetime('now') ELSE artisan_signed_at END,
      supervisor_name = COALESCE(?, supervisor_name),
      supervisor_signed_at = CASE WHEN ? IS NOT NULL THEN datetime('now') ELSE supervisor_signed_at END
    WHERE id = ?
  `);

  const tx = db.transaction(() => {
    closeWorkOrder.run(
      completion_notes,
      artisan_name,
      artisan_name,
      supervisor_name,
      supervisor_name,
      woId
    );

    const isService = String(wo.source || "").toLowerCase() === "service";
    if (isService) {
      const planId = Number(wo.reference_id || 0);
      if (planId > 0) {
        const current = getAssetCurrentHours(Number(wo.asset_id || 0));
        db.prepare(`UPDATE maintenance_plans SET last_service_hours = ? WHERE id = ?`).run(current, planId);
      }
    }
  });
  tx();

  writeAudit(db, req, {
    module: "workorders",
    action: "close_approved",
    entity_type: "work_order",
    entity_id: woId,
    payload: { approval_id: approvalId },
  });
  return { woId, already_closed: false };
}

function applyStockAdjustWithPayload(approvalId, payload, req) {
  const partCode = String(payload?.part_code || "").trim();
  const quantity = Number(payload?.quantity ?? 0);
  const reference = String(payload?.reference || `approval_adjust:${approvalId}`).trim() || `approval_adjust:${approvalId}`;
  if (!partCode) throw new Error(`approval ${approvalId}: part_code is required`);
  if (!Number.isFinite(quantity) || quantity === 0) throw new Error(`approval ${approvalId}: quantity must be non-zero`);

  const part = db.prepare(`SELECT id FROM parts WHERE part_code = ?`).get(partCode);
  if (!part) throw new Error(`approval ${approvalId}: part_code not found`);

  const onHand = Number(
    db.prepare(`SELECT IFNULL(SUM(quantity), 0) AS on_hand FROM stock_movements WHERE part_id = ?`).get(part.id)?.on_hand || 0
  );
  if (quantity < 0 && onHand < Math.abs(quantity)) throw new Error(`approval ${approvalId}: insufficient stock`);

  db.prepare(`
    INSERT INTO stock_movements (part_id, quantity, movement_type, reference)
    VALUES (?, ?, 'adjust', ?)
  `).run(part.id, quantity, reference);

  const onHandAfter = Number(
    db.prepare(`SELECT IFNULL(SUM(quantity), 0) AS on_hand FROM stock_movements WHERE part_id = ?`).get(part.id)?.on_hand || 0
  );
  writeAudit(db, req, {
    module: "stock",
    action: "adjust_approved",
    entity_type: "part",
    entity_id: partCode,
    payload: { approval_id: approvalId, quantity, on_hand_before: onHand, on_hand_after: onHandAfter },
  });
  return { part_code: partCode, quantity, on_hand_before: onHand, on_hand_after: onHandAfter };
}

function approveProcurementRequisitionWithPayload(approvalId, payload, req) {
  const requisitionId = Number(payload?.requisition_id || 0);
  if (!Number.isFinite(requisitionId) || requisitionId <= 0) {
    throw new Error(`approval ${approvalId}: invalid requisition_id`);
  }
  const row = db.prepare(`
    SELECT id, status
    FROM procurement_requisitions
    WHERE id = ?
  `).get(requisitionId);
  if (!row) throw new Error(`approval ${approvalId}: requisition not found`);

  db.prepare(`
    UPDATE procurement_requisitions
    SET status = 'approved', updated_at = datetime('now')
    WHERE id = ?
  `).run(requisitionId);

  writeAudit(db, req, {
    module: "procurement",
    action: "requisition_approved",
    entity_type: "requisition",
    entity_id: requisitionId,
    payload: { approval_id: approvalId },
  });
  return { requisition_id: requisitionId, status: "approved" };
}

function receiveProcurementRequisitionWithPayload(approvalId, payload, req) {
  const requisitionId = Number(payload?.requisition_id || 0);
  const partCode = String(payload?.part_code || "").trim();
  const qtyReceive = Number(payload?.qty_receive ?? 0);
  const reference = String(payload?.reference || `requisition:${requisitionId}`).trim() || `requisition:${requisitionId}`;
  if (!Number.isFinite(requisitionId) || requisitionId <= 0) {
    throw new Error(`approval ${approvalId}: invalid requisition_id`);
  }
  if (!partCode) throw new Error(`approval ${approvalId}: invalid part_code`);
  if (!Number.isFinite(qtyReceive) || qtyReceive <= 0) throw new Error(`approval ${approvalId}: invalid qty_receive`);

  const part = db.prepare(`SELECT id FROM parts WHERE part_code = ?`).get(partCode);
  if (!part) throw new Error(`approval ${approvalId}: part not found`);

  const reqn = db.prepare(`
    SELECT id, qty_requested, qty_received, status
    FROM procurement_requisitions
    WHERE id = ?
  `).get(requisitionId);
  if (!reqn) throw new Error(`approval ${approvalId}: requisition not found`);
  const outstanding = Number(reqn.qty_requested || 0) - Number(reqn.qty_received || 0);
  if (qtyReceive > outstanding) throw new Error(`approval ${approvalId}: qty_receive exceeds outstanding (${outstanding})`);

  const tx = db.transaction(() => {
    db.prepare(`
      INSERT INTO stock_movements (part_id, quantity, movement_type, reference)
      VALUES (?, ?, 'in', ?)
    `).run(part.id, Math.abs(qtyReceive), reference);

    db.prepare(`
      UPDATE procurement_requisitions
      SET
        qty_received = qty_received + ?,
        status = CASE
          WHEN (qty_received + ?) >= qty_requested THEN 'received'
          ELSE 'approved'
        END,
        updated_at = datetime('now')
      WHERE id = ?
    `).run(qtyReceive, qtyReceive, requisitionId);
  });
  tx();

  const updated = db.prepare(`
    SELECT id, qty_requested, qty_received, status
    FROM procurement_requisitions
    WHERE id = ?
  `).get(requisitionId);

  writeAudit(db, req, {
    module: "procurement",
    action: "requisition_receive_approved",
    entity_type: "requisition",
    entity_id: requisitionId,
    payload: { approval_id: approvalId, qty_receive: qtyReceive, part_code: partCode },
  });

  return {
    requisition_id: requisitionId,
    qty_requested: Number(updated?.qty_requested || 0),
    qty_received: Number(updated?.qty_received || 0),
    status: String(updated?.status || ""),
    part_code: partCode,
    qty_receive: qtyReceive,
  };
}

export default async function approvalsRoutes(app) {
  ensureAuditTable(db);
  db.prepare(`
    CREATE TABLE IF NOT EXISTS approval_requests (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      module TEXT NOT NULL,
      action TEXT NOT NULL,
      entity_type TEXT,
      entity_id TEXT,
      status TEXT NOT NULL DEFAULT 'pending',
      payload_json TEXT,
      requested_by TEXT,
      requested_role TEXT,
      approved_by TEXT,
      approved_role TEXT,
      approved_at TEXT,
      rejected_by TEXT,
      rejected_role TEXT,
      rejected_at TEXT,
      decision_note TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  ensureColumn("approval_requests", "decision_note", "decision_note TEXT");

  app.post("/", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores", "artisan"])) return;
    const body = req.body || {};
    const module = String(body.module || "").trim();
    const action = String(body.action || "").trim();
    const entity_type = body.entity_type != null ? String(body.entity_type).trim() : null;
    const entity_id = body.entity_id != null ? String(body.entity_id).trim() : null;
    if (!module || !action) return reply.code(400).send({ error: "module and action are required" });
    let payload_json = null;
    if (body.payload != null) {
      try {
        payload_json = JSON.stringify(body.payload);
      } catch {
        return reply.code(400).send({ error: "payload must be valid JSON" });
      }
    }
    const ins = db.prepare(`
      INSERT INTO approval_requests (module, action, entity_type, entity_id, status, payload_json, requested_by, requested_role)
      VALUES (?, ?, ?, ?, 'pending', ?, ?, ?)
    `).run(module, action, entity_type, entity_id, payload_json, getUser(req), getRole(req));
    const id = Number(ins.lastInsertRowid);
    writeAudit(db, req, {
      module: "approvals",
      action: "request_create",
      entity_type: "approval_request",
      entity_id: id,
      payload: { module, action, entity_type, entity_id },
    });
    return { ok: true, id, status: "pending" };
  });

  app.get("/", async (req) => {
    const status = String(req.query?.status || "").trim().toLowerCase();
    const module = String(req.query?.module || "").trim().toLowerCase();
    const action = String(req.query?.action || "").trim().toLowerCase();
    const where = [];
    const params = [];
    if (status) { where.push("LOWER(status) = ?"); params.push(status); }
    if (module) { where.push("LOWER(module) = ?"); params.push(module); }
    if (action) { where.push("LOWER(action) = ?"); params.push(action); }
    const rows = db.prepare(`
      SELECT *
      FROM approval_requests
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY id DESC
      LIMIT 300
    `).all(...params).map((r) => ({ ...r, id: Number(r.id), payload: parsePayload(r.payload_json) }));
    return { ok: true, rows };
  });

  app.post("/:id/approve", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const note = String(req.body?.note || "").trim() || null;
    const row = db.prepare(`SELECT * FROM approval_requests WHERE id = ?`).get(id);
    if (!row) return reply.code(404).send({ error: "approval request not found" });
    if (String(row.status || "").toLowerCase() !== "pending") {
      return reply.code(409).send({ error: `approval request already ${row.status}` });
    }

    const payload = parsePayload(row.payload_json) || {};
    let execution = { ok: true };
    if (row.module === "stock" && row.action === "adjust_movement") {
      execution = applyStockAdjustWithPayload(id, payload, req);
    } else if (row.module === "workorders" && row.action === "close_work_order") {
      execution = closeWorkOrderWithPayload(id, payload, req);
    } else if (row.module === "procurement" && row.action === "approve_requisition") {
      execution = approveProcurementRequisitionWithPayload(id, payload, req);
    } else if (row.module === "procurement" && row.action === "receive_requisition") {
      execution = receiveProcurementRequisitionWithPayload(id, payload, req);
    }

    db.prepare(`
      UPDATE approval_requests
      SET status = 'approved', approved_by = ?, approved_role = ?, approved_at = datetime('now'), decision_note = ?
      WHERE id = ?
    `).run(getUser(req), getRole(req), note, id);

    writeAudit(db, req, {
      module: "approvals",
      action: "request_approve",
      entity_type: "approval_request",
      entity_id: id,
      payload: { module: row.module, action: row.action, note },
    });
    return { ok: true, id, status: "approved", execution };
  });

  app.post("/:id/reject", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const note = String(req.body?.note || "").trim() || null;
    const row = db.prepare(`SELECT id, status, module, action FROM approval_requests WHERE id = ?`).get(id);
    if (!row) return reply.code(404).send({ error: "approval request not found" });
    if (String(row.status || "").toLowerCase() !== "pending") {
      return reply.code(409).send({ error: `approval request already ${row.status}` });
    }
    db.prepare(`
      UPDATE approval_requests
      SET status = 'rejected', rejected_by = ?, rejected_role = ?, rejected_at = datetime('now'), decision_note = ?
      WHERE id = ?
    `).run(getUser(req), getRole(req), note, id);
    writeAudit(db, req, {
      module: "approvals",
      action: "request_reject",
      entity_type: "approval_request",
      entity_id: id,
      payload: { module: row.module, action: row.action, note },
    });
    return { ok: true, id, status: "rejected" };
  });
}
