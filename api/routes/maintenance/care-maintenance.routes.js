// IRONLOG/api/routes/maintenance/care-maintenance.routes.js — care and maintenance plans.
// Registered by routes/maintenance.routes.js under /api/maintenance.
//   GET    /cm/plans                         plans for this site
//   POST   /cm/plans                         { name, start_date, end_date, labour_rate, notes }
//   GET    /cm/plans/:id                     plan, machines, estimate vs used, totals
//   PATCH  /cm/plans/:id                     plan fields, or { status: "closed" }
//   GET    /cm/plans/:id/export.csv          one row per machine, for Excel
//   GET    /cm/plans/:id/plan.pdf            the plan to print: summary, then work and parts per machine
//   POST   /cm/plans/:id/machines            { asset_ids: [...] }
//   PATCH  /cm/machines/:rowId               { scope, planned_hours, status, priority, notes }
//   DELETE /cm/machines/:rowId
//   PUT    /cm/machines/:rowId/parts         { lines: [{ part_code, description, qty, unit_cost }] }
//   POST   /cm/machines/:rowId/work-order    open the machine's work order
//   POST   /cm/machines/:rowId/request-parts raise Parts to Order requests for what stores lack
import { db } from "../../db/client.js";
import { writeAudit } from "../../utils/audit.js";
import { buildPdfBuffer } from "../../utils/pdfGenerator.js";
import { drawCarePlanPdf } from "../../utils/careMaintenancePdf.js";
import {
  addMachines,
  createMachineWorkOrder,
  createPlan,
  getMachineRow,
  getPlan,
  listPlans,
  planCsv,
  planDetail,
  removeMachine,
  requestMissingParts,
  setParts,
  updateMachine,
  updatePlan,
} from "../../utils/careMaintenance.js";

// Plant admin, foreman and managers plan and change; stores can look.
export const CM_EDIT_ROLES = ["admin", "supervisor", "workshop_admin", "plant_manager", "site_manager", "workshop_manager"];
export const CM_VIEW_ROLES = [...CM_EDIT_ROLES, "storeman", "stores"];

export default function registerCareMaintenanceRoutes(app, ctx) {
  const { requireMaintenanceRoles } = ctx;
  const site = (req) => String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
  const actor = (req) => String(req.headers?.["x-user-name"] || "").trim() || "admin";
  const audit = (req, action, entityType, id, payload) => {
    try {
      writeAudit(db, req, { module: "care_maintenance", action, entity_type: entityType, entity_id: String(id), payload });
    } catch {}
  };
  const machineOr404 = (req, reply) => {
    const row = getMachineRow(db, req.params.rowId, site(req));
    if (!row) reply.code(404).send({ ok: false, error: "Machine not found on this plan." });
    return row;
  };

  app.get("/cm/plans", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, CM_VIEW_ROLES)) return;
    return { ok: true, plans: listPlans(db, site(req)) };
  });

  app.post("/cm/plans", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, CM_EDIT_ROLES)) return;
    const r = createPlan(db, site(req), req.body || {}, actor(req));
    if (!r.ok) return reply.code(400).send(r);
    audit(req, "plan_create", "cm_plan", r.id, { name: req.body?.name });
    return r;
  });

  app.get("/cm/plans/:id", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, CM_VIEW_ROLES)) return;
    const detail = planDetail(db, req.params.id, site(req));
    if (!detail) return reply.code(404).send({ ok: false, error: "Plan not found." });
    return { ok: true, ...detail };
  });

  app.patch("/cm/plans/:id", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, CM_EDIT_ROLES)) return;
    if (!getPlan(db, req.params.id, site(req))) return reply.code(404).send({ ok: false, error: "Plan not found." });
    const r = updatePlan(db, req.params.id, req.body || {});
    if (!r.ok) return reply.code(400).send(r);
    audit(req, "plan_update", "cm_plan", req.params.id, req.body || {});
    return r;
  });

  app.get("/cm/plans/:id/export.csv", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, CM_VIEW_ROLES)) return;
    const detail = planDetail(db, req.params.id, site(req));
    if (!detail) return reply.code(404).send({ ok: false, error: "Plan not found." });
    const name = detail.plan.name.replace(/[^A-Za-z0-9-]+/g, "_").replace(/^_+|_+$/g, "") || "plan";
    reply.header("Content-Type", "text/csv; charset=utf-8");
    reply.header("Content-Disposition", `attachment; filename="care-maintenance-${name}.csv"`);
    return `﻿${planCsv(detail)}`;
  });

  app.get("/cm/plans/:id/plan.pdf", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, CM_VIEW_ROLES)) return;
    const detail = planDetail(db, req.params.id, site(req));
    if (!detail) return reply.code(404).send({ ok: false, error: "Plan not found." });
    const name = detail.plan.name.replace(/[^A-Za-z0-9-]+/g, "_").replace(/^_+|_+$/g, "") || "plan";
    const pdf = await buildPdfBuffer((doc) => drawCarePlanPdf(doc, detail), {
      title: "Care & maintenance plan",
      subtitle: detail.plan.name,
    });
    reply.header("Content-Type", "application/pdf");
    reply.header("Content-Disposition", `inline; filename="care-maintenance-${name}.pdf"`);
    return reply.send(pdf);
  });

  app.post("/cm/plans/:id/machines", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, CM_EDIT_ROLES)) return;
    const plan = getPlan(db, req.params.id, site(req));
    if (!plan) return reply.code(404).send({ ok: false, error: "Plan not found." });
    const r = addMachines(db, plan.id, req.body?.asset_ids, actor(req));
    audit(req, "machines_add", "cm_plan", plan.id, { added: r.added });
    return r;
  });

  app.patch("/cm/machines/:rowId", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, CM_EDIT_ROLES)) return;
    const row = machineOr404(req, reply);
    if (!row) return;
    const r = updateMachine(db, row.id, req.body || {}, actor(req));
    if (!r.ok) return reply.code(400).send(r);
    if (req.body?.status !== undefined && req.body.status !== row.status) {
      audit(req, "machine_status", "cm_plan_machine", row.id, { asset_code: row.asset_code, from: row.status, to: req.body.status });
    }
    return r;
  });

  app.delete("/cm/machines/:rowId", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, CM_EDIT_ROLES)) return;
    const row = machineOr404(req, reply);
    if (!row) return;
    audit(req, "machine_remove", "cm_plan_machine", row.id, { asset_code: row.asset_code, plan_id: row.plan_id });
    return removeMachine(db, row.id);
  });

  app.put("/cm/machines/:rowId/parts", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, CM_EDIT_ROLES)) return;
    const row = machineOr404(req, reply);
    if (!row) return;
    const r = setParts(db, row.id, req.body?.lines);
    if (!r.ok) return reply.code(400).send(r);
    return r;
  });

  app.post("/cm/machines/:rowId/work-order", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, CM_EDIT_ROLES)) return;
    const r = createMachineWorkOrder(db, req.params.rowId, site(req), actor(req));
    if (!r.ok) return reply.code(r.status || 400).send(r);
    audit(req, "work_order_create", "work_order", r.work_order_id, { cm_machine_row: Number(req.params.rowId) });
    return r;
  });

  app.post("/cm/machines/:rowId/request-parts", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, CM_EDIT_ROLES)) return;
    const r = requestMissingParts(db, req.params.rowId, site(req), actor(req));
    if (!r.ok) return reply.code(r.status || 400).send(r);
    if (r.created.length) audit(req, "parts_request", "cm_plan_machine", req.params.rowId, { requests: r.created.map((c) => c.request_id) });
    return r;
  });
}
