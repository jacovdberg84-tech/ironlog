// IRONLOG/api/routes/workorders/closeout.routes.js — Parts issue, close requests, reopening, deletion requests and closing.
// Registered by routes/workorders.routes.js; shared helpers arrive through ctx.
import { applyIssuedQuantityToReservation, releaseWorkOrderReservations } from "../../utils/serviceTemplates.js";
import { db } from "../../db/client.js";
import { snapLastServiceHours } from "../../utils/serviceSchedule.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerCloseoutRoutes(app, ctx) {
  const { getAssetCurrentHours, getRole, requirePermission, requireRoles } = ctx;

    // Issue parts to a work order (creates stock movement OUT)
  // Body: { part_code, quantity }
  app.post("/:id/issue", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });

    const body = req.body || {};
    const part_code = String(body.part_code || "").trim();
    const quantity = Number(body.quantity ?? 0);

    if (!part_code || !Number.isFinite(quantity) || quantity <= 0) {
      return reply.code(400).send({ error: "part_code and quantity (>0) required" });
    }

    const wo = db.prepare(`SELECT id, status FROM work_orders WHERE id = ?`).get(id);
    if (!wo) return reply.code(404).send({ error: "work order not found" });
    if (wo.status === "closed") return reply.code(409).send({ error: "work order is closed" });

    const part = db.prepare(`SELECT id FROM parts WHERE part_code = ?`).get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });

    // Stock on hand = sum(movements)
    const onHandRow = db.prepare(`
      SELECT IFNULL(SUM(quantity), 0) AS on_hand
      FROM stock_movements
      WHERE part_id = ?
    `).get(part.id);

    const on_hand = Number(onHandRow.on_hand || 0);
    const otherReservations = Number(db.prepare(`
      SELECT COALESCE(SUM(quantity_reserved - quantity_issued), 0) AS reserved
      FROM stock_reservations
      WHERE part_id = ? AND work_order_id <> ? AND status = 'active'
    `).get(part.id, id)?.reserved || 0);
    const available_to_issue = Math.max(0, on_hand - otherReservations);
    if (available_to_issue < quantity) {
      return reply.code(409).send({
        error: "insufficient stock",
        part_code,
        on_hand,
        reserved_for_other_work: otherReservations,
        available_to_issue,
        requested: quantity
      });
    }

    // Insert movement (negative quantity = out)
    db.prepare(`
      INSERT INTO stock_movements (part_id, quantity, movement_type, reference)
      VALUES (?, ?, 'out', ?)
    `).run(part.id, -Math.abs(Math.trunc(quantity)), `work_order:${id}`);
    applyIssuedQuantityToReservation(db, id, part.id, quantity);

    writeAudit(db, req, {
      module: "workorders",
      action: "issue_part",
      entity_type: "work_order",
      entity_id: id,
      payload: { part_code, quantity },
    });

    return reply.send({ ok: true, part_code, issued: quantity, on_hand_before: on_hand, on_hand_after: on_hand - quantity });
  });

  // Request close approval for a work order
  app.post("/:id/request-close", async (req, reply) => {
    if (!requirePermission(req, reply, "workorders.close.request")) return;
    if (!requireRoles(req, reply, ["admin", "supervisor", "artisan"])) return;
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });

    const wo = db.prepare(`
      SELECT id, status, source
      FROM work_orders
      WHERE id = ?
    `).get(id);
    if (!wo) return reply.code(404).send({ error: "work order not found" });

    const status = String(wo.status || "").toLowerCase();
    if (!["completed", "approved"].includes(status)) {
      return reply.code(409).send({ error: "work order must be completed or approved before close approval request" });
    }

    const body = req.body || {};
    const completion_notes =
      body.completion_notes != null && String(body.completion_notes).trim() !== ""
        ? String(body.completion_notes).trim()
        : null;
    const artisan_name =
      body.artisan_name != null && String(body.artisan_name).trim() !== ""
        ? String(body.artisan_name).trim()
        : null;
    const supervisor_name =
      body.supervisor_name != null && String(body.supervisor_name).trim() !== ""
        ? String(body.supervisor_name).trim()
        : null;

    const duplicatePending = db.prepare(`
      SELECT id
      FROM approval_requests
      WHERE module = 'workorders'
        AND action = 'close_work_order'
        AND entity_type = 'work_order'
        AND entity_id = ?
        AND status = 'pending'
      ORDER BY id DESC
      LIMIT 1
    `).get(String(id));
    if (duplicatePending) {
      return reply.send({ ok: true, pending_approval: true, request_id: Number(duplicatePending.id), duplicate: true });
    }

    const payload_json = JSON.stringify({
      work_order_id: id,
      completion_notes,
      artisan_name,
      supervisor_name,
    });
    const requestedBy = String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
    const requestedRole = getRole(req);

    const ins = db.prepare(`
      INSERT INTO approval_requests (
        module, action, entity_type, entity_id, status, payload_json, requested_by, requested_role
      )
      VALUES ('workorders', 'close_work_order', 'work_order', ?, 'pending', ?, ?, ?)
    `).run(String(id), payload_json, requestedBy, requestedRole);
    const request_id = Number(ins.lastInsertRowid);

    writeAudit(db, req, {
      module: "workorders",
      action: "close_request",
      entity_type: "work_order",
      entity_id: id,
      payload: { request_id, source: wo.source },
    });

    return reply.send({ ok: true, pending_approval: true, request_id });
  });

  app.post("/:id/reopen", async (req, reply) => {
    if (!requirePermission(req, reply, "workorders.reopen")) return;
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params.id);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const note = String(req.body?.note || "").trim() || "Reopened by supervisor flow";
    const wo = db.prepare(`SELECT id, status FROM work_orders WHERE id = ?`).get(id);
    if (!wo) return reply.code(404).send({ error: "work order not found" });
    const status = String(wo.status || "").toLowerCase();
    if (!["closed", "approved", "completed"].includes(status)) {
      return reply.code(409).send({ error: "only closed/approved/completed work orders can be reopened" });
    }
    db.prepare(`UPDATE work_orders SET status = 'in_progress', closed_at = NULL WHERE id = ?`).run(id);
    writeAudit(db, req, {
      module: "workorders",
      action: "reopen",
      entity_type: "work_order",
      entity_id: id,
      payload: { from_status: status, to_status: "in_progress", note },
    });
    return reply.send({ ok: true, id, status: "in_progress" });
  });

  app.post("/:id/delete-request", async (req, reply) => {
    if (!requirePermission(req, reply, "workorders.delete.request")) return;
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params.id);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const wo = db.prepare(`SELECT id FROM work_orders WHERE id = ?`).get(id);
    if (!wo) return reply.code(404).send({ error: "work order not found" });
    const reason = String(req.body?.reason || "").trim();
    if (!reason) return reply.code(400).send({ error: "reason is required" });
    const requestedBy = String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
    const requestedRole = getRole(req);
    const payload_json = JSON.stringify({ work_order_id: id, reason, requested_by: requestedBy });
    const ins = db.prepare(`
      INSERT INTO approval_requests (
        module, action, entity_type, entity_id, status, payload_json, requested_by, requested_role
      ) VALUES ('workorders', 'delete_work_order', 'work_order', ?, 'pending', ?, ?, ?)
    `).run(String(id), payload_json, requestedBy, requestedRole);
    const request_id = Number(ins.lastInsertRowid);
    writeAudit(db, req, {
      module: "workorders",
      action: "delete_request",
      entity_type: "work_order",
      entity_id: id,
      payload: { request_id, reason },
    });
    return reply.send({ ok: true, pending_approval: true, request_id });
  });

  // Close a work order
  app.post("/:id/close", async (req, reply) => {
    if (!requirePermission(req, reply, "workorders.close.approve")) return;
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });

    const wo = db.prepare(`
      SELECT id, status, source, reference_id, asset_id
      FROM work_orders
      WHERE id = ?
    `).get(id);
    if (!wo) return reply.code(404).send({ error: "work order not found" });
    if (wo.status === "closed") return reply.code(409).send({ error: "work order already closed" });

    const body = req.body || {};
    const completion_notes =
      body.completion_notes != null && String(body.completion_notes).trim() !== ""
        ? String(body.completion_notes).trim()
        : null;
    const artisan_name =
      body.artisan_name != null && String(body.artisan_name).trim() !== ""
        ? String(body.artisan_name).trim()
        : null;
    const supervisor_name =
      body.supervisor_name != null && String(body.supervisor_name).trim() !== ""
        ? String(body.supervisor_name).trim()
        : null;

    const isServiceWO = String(wo.source || "").toLowerCase() === "service";
    if (isServiceWO && !artisan_name) {
      return reply.code(400).send({
        error: "artisan_name is required when closing a service work order"
      });
    }
    if (isServiceWO && !completion_notes) {
      return reply.code(400).send({
        error: "completion_notes is required when closing a service work order"
      });
    }

    const closeWorkOrder = db.prepare(`
      UPDATE work_orders
      SET
        status='closed',
        completed_at = datetime('now'),
        closed_at = datetime('now'),
        completion_notes = COALESCE(?, completion_notes),
        artisan_name = COALESCE(?, artisan_name),
        artisan_signed_at = CASE
          WHEN ? IS NOT NULL THEN datetime('now')
          ELSE artisan_signed_at
        END,
        supervisor_name = COALESCE(?, supervisor_name),
        supervisor_signed_at = CASE
          WHEN ? IS NOT NULL THEN datetime('now')
          ELSE supervisor_signed_at
        END
      WHERE id = ?
    `);

    const updatePlanLastServiceHours = db.prepare(`
      UPDATE maintenance_plans
      SET last_service_hours = ?
      WHERE id = ?
    `);
    const updateAllAssetPlanLastServiceHours = db.prepare(`
      UPDATE maintenance_plans
      SET last_service_hours = ?
      WHERE asset_id = ?
        AND active = 1
    `);
    const closeLinkedBreakdownWhenWorkIsFinished = db.prepare(`
      UPDATE breakdowns
      SET
        status = 'CLOSED',
        end_at = COALESCE(NULLIF(TRIM(end_at), ''), datetime('now'))
      WHERE id = ?
        AND NOT EXISTS (
          SELECT 1
          FROM work_orders linked
          WHERE linked.source = 'breakdown'
            AND linked.reference_id = breakdowns.id
            AND linked.id <> ?
            AND REPLACE(TRIM(LOWER(COALESCE(linked.status, ''))), ' ', '_')
              NOT IN ('completed', 'approved', 'closed')
        )
    `);

    const tx = db.transaction(() => {
      closeWorkOrder.run(
        completion_notes,
        artisan_name,
        artisan_name,
        supervisor_name,
        supervisor_name,
        id
      );
      releaseWorkOrderReservations(db, id);

      if (String(wo.source || "").trim().toLowerCase() === "breakdown" && Number(wo.reference_id || 0) > 0) {
        closeLinkedBreakdownWhenWorkIsFinished.run(Number(wo.reference_id), id);
      }

      let rolled_plan_id = null;
      let rolled_last_service_hours = null;

      const planId = Number(wo.reference_id || 0);

      if (isServiceWO && planId > 0) {
        const currentHours = getAssetCurrentHours(Number(wo.asset_id || 0));
        const planRow = db.prepare(`
          SELECT mp.interval_hours, a.asset_code
          FROM maintenance_plans mp
          JOIN assets a ON a.id = mp.asset_id
          WHERE mp.id = ?
        `).get(planId);
        const safeHours = snapLastServiceHours(
          Number.isFinite(currentHours) ? currentHours : 0,
          Number(planRow?.interval_hours || 0),
          planRow?.asset_code,
        );

        updatePlanLastServiceHours.run(safeHours, planId);
        updateAllAssetPlanLastServiceHours.run(safeHours, Number(wo.asset_id || 0));
        rolled_plan_id = planId;
        rolled_last_service_hours = safeHours;
      }

      return { rolled_plan_id, rolled_last_service_hours };
    });

    const result = tx();

    writeAudit(db, req, {
      module: "workorders",
      action: "close",
      entity_type: "work_order",
      entity_id: id,
      payload: {
        completion_notes,
        artisan_name,
        supervisor_name,
        rolled_plan_id: result.rolled_plan_id,
        rolled_last_service_hours: result.rolled_last_service_hours,
      },
    });

    return reply.send({
      ok: true,
      rolled_plan_id: result.rolled_plan_id,
      rolled_last_service_hours: result.rolled_last_service_hours,
      completion_notes,
      artisan_name,
      supervisor_name
    });
  });
}
