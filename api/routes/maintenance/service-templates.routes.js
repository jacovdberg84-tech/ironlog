// IRONLOG/api/routes/maintenance/service-templates.routes.js — Service templates, service planner and service estimates.
// Registered by routes/maintenance.routes.js; shared helpers arrive through ctx.
import { buildDueListFromPlans, meterUnitForAsset } from "../../utils/serviceSchedule.js";
import { buildServiceEstimatePreview } from "../../utils/serviceTemplates.js";
import { applyServiceTemplateProposals, buildServiceTemplateProposals } from "../../utils/serviceTemplateBuilder.js";
import { db } from "../../db/client.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerServiceTemplatesRoutes(app, ctx) {
  const {
    getAssetCurrentHours,
    listMaintenancePlans,
    loadTemplateDetail,
    maintenanceActor,
    requireMaintenancePermission,
    templateSummaryRows,
    writeTemplateItems,
  } = ctx;

  // =====================================================
  // SERVICE TEMPLATE MANAGEMENT
  // =====================================================
  app.get("/service-templates", async (req, reply) => {
    if (!requireMaintenancePermission(req, reply, "maintenance.templates.read")) return;
    return reply.send({ ok: true, templates: templateSummaryRows({ activeOnly: String(req.query?.active || "") === "1" }) });
  });

  // Proposed templates per machine and service interval, built from past
  // service work orders (or a service kit matched by stock code). Read-only.
  app.get("/service-templates/suggestions", async (req, reply) => {
    if (!requireMaintenancePermission(req, reply, "maintenance.templates.manage")) return;
    return reply.send({ ok: true, proposals: buildServiceTemplateProposals(db) });
  });

  // POST { keys: ["<asset_id>:<interval>", ...] } — create the chosen proposals.
  app.post("/service-templates/suggestions/apply", async (req, reply) => {
    if (!requireMaintenancePermission(req, reply, "maintenance.templates.manage")) return;
    const keys = Array.isArray(req.body?.keys) ? req.body.keys.map(String) : [];
    if (!keys.length) return reply.code(400).send({ ok: false, error: "Choose at least one template to create" });
    const created = applyServiceTemplateProposals(db, buildServiceTemplateProposals(db), keys);
    writeAudit(db, req, {
      module: "maintenance", action: "service_template.build_from_history", entity_type: "service_template",
      entity_id: created.map((c) => c.id).join(","), after: created,
    });
    return reply.code(201).send({ ok: true, created });
  });

  app.get("/service-templates/:id", async (req, reply) => {
    if (!requireMaintenancePermission(req, reply, "maintenance.templates.read")) return;
    const template = loadTemplateDetail(req.params?.id);
    if (!template) return reply.code(404).send({ ok: false, error: "Service template not found" });
    return reply.send({ ok: true, template });
  });

  app.post("/service-templates", async (req, reply) => {
    if (!requireMaintenancePermission(req, reply, "maintenance.templates.manage")) return;
    try {
      const body = req.body || {};
      const template_key = String(body.template_key || "").trim().toUpperCase();
      const name = String(body.name || "").trim();
      const service_interval_hours = Number(body.service_interval_hours || body.interval_hours || 0);
      const meter_unit = String(body.meter_unit || "hours").trim().toLowerCase() === "km" ? "km" : "hours";
      if (!template_key || !name || !Number.isFinite(service_interval_hours) || service_interval_hours <= 0) {
        return reply.code(400).send({ ok: false, error: "template_key, name and a positive service interval are required" });
      }
      const existing = db.prepare(`SELECT id FROM service_templates WHERE template_key = ? AND revision_number = 1`).get(template_key);
      if (existing) return reply.code(409).send({ ok: false, error: `Template key ${template_key} already exists` });
      const defaultAssignment = {
        asset_id: Number(body.asset_id || 0) || null,
        manufacturer: String(body.manufacturer || "").trim() || null,
        model: String(body.model || "").trim() || null,
        asset_category: String(body.asset_category || body.category || "").trim() || null,
      };
      if (defaultAssignment.asset_id && !db.prepare(`SELECT id FROM assets WHERE id = ?`).get(defaultAssignment.asset_id)) {
        return reply.code(404).send({ ok: false, error: "Assignment asset not found" });
      }

      const create = db.transaction(() => {
        const insert = db.prepare(`
          INSERT INTO service_templates (
            template_key, name, description, manufacturer, model, asset_category,
            service_interval_hours, meter_unit, estimated_duration_hours,
            default_labour_hours, default_labour_rate, active, revision_number
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 1)
        `).run(
          template_key, name, String(body.description || "").trim() || null,
          String(body.manufacturer || "").trim() || null, String(body.model || "").trim() || null,
          String(body.asset_category || body.category || "").trim() || null,
          service_interval_hours, meter_unit, Math.max(0, Number(body.estimated_duration_hours || 0)),
          Math.max(0, Number(body.default_labour_hours || 0)), Math.max(0, Number(body.default_labour_rate || 0)),
        );
        const id = Number(insert.lastInsertRowid);
        writeTemplateItems(id, body.items);
        if (defaultAssignment.asset_id || (defaultAssignment.manufacturer && defaultAssignment.model) || defaultAssignment.asset_category) {
          db.prepare(`
            INSERT INTO asset_service_template_assignments (
              asset_id, manufacturer, model, asset_category, service_template_id, priority, active
            ) VALUES (?, ?, ?, ?, ?, 100, 1)
          `).run(
            defaultAssignment.asset_id, defaultAssignment.manufacturer, defaultAssignment.model,
            defaultAssignment.asset_category, id,
          );
        }
        return id;
      });
      const id = create();
      const template = loadTemplateDetail(id);
      writeAudit(db, req, {
        module: "maintenance", action: "service_template.create", entity_type: "service_template", entity_id: String(id), after: template,
      });
      return reply.code(201).send({ ok: true, id, template });
    } catch (err) {
      return reply.code(400).send({ ok: false, error: err.message || "Unable to create service template" });
    }
  });

  // A revision is a new immutable definition. Existing estimates and work
  // orders remain linked to the exact version they were created from.
  app.post("/service-templates/:id/revisions", async (req, reply) => {
    if (!requireMaintenancePermission(req, reply, "maintenance.templates.manage")) return;
    try {
      const previous = loadTemplateDetail(req.params?.id);
      if (!previous) return reply.code(404).send({ ok: false, error: "Service template not found" });
      const body = req.body || {};
      const revision = Number(previous.revision_number || 1) + 1;
      const items = Object.prototype.hasOwnProperty.call(body, "items") ? body.items : previous.items;
      const create = db.transaction(() => {
        const insert = db.prepare(`
          INSERT INTO service_templates (
            template_key, name, description, manufacturer, model, asset_category,
            service_interval_hours, meter_unit, estimated_duration_hours,
            default_labour_hours, default_labour_rate, active, revision_number, supersedes_template_id
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
        `).run(
          previous.template_key,
          String(body.name ?? previous.name).trim(),
          String(body.description ?? previous.description ?? "").trim() || null,
          String(body.manufacturer ?? previous.manufacturer ?? "").trim() || null,
          String(body.model ?? previous.model ?? "").trim() || null,
          String(body.asset_category ?? previous.asset_category ?? "").trim() || null,
          Math.max(0.001, Number(body.service_interval_hours ?? previous.service_interval_hours)),
          String(body.meter_unit ?? previous.meter_unit) === "km" ? "km" : "hours",
          Math.max(0, Number(body.estimated_duration_hours ?? previous.estimated_duration_hours)),
          Math.max(0, Number(body.default_labour_hours ?? previous.default_labour_hours)),
          Math.max(0, Number(body.default_labour_rate ?? previous.default_labour_rate)),
          revision, Number(previous.id),
        );
        const id = Number(insert.lastInsertRowid);
        writeTemplateItems(id, items);
        // Keep the same matching scope for the newly approved definition.
        // The previous rows stay intact for audit/history, while only this
        // active revision participates in future planner matching.
        db.prepare(`
          INSERT INTO asset_service_template_assignments (
            asset_id, manufacturer, model, asset_category, service_template_id, priority, active
          )
          SELECT asset_id, manufacturer, model, asset_category, ?, priority, active
          FROM asset_service_template_assignments
          WHERE service_template_id = ?
        `).run(id, Number(previous.id));
        db.prepare(`UPDATE service_templates SET active = 0, updated_at = datetime('now') WHERE id = ?`).run(Number(previous.id));
        return id;
      });
      const id = create();
      const template = loadTemplateDetail(id);
      writeAudit(db, req, {
        module: "maintenance", action: "service_template.revise", entity_type: "service_template", entity_id: String(id),
        before: previous, after: template,
      });
      return reply.code(201).send({ ok: true, id, template });
    } catch (err) {
      return reply.code(400).send({ ok: false, error: err.message || "Unable to revise service template" });
    }
  });

  app.post("/service-templates/:id/assignments", async (req, reply) => {
    if (!requireMaintenancePermission(req, reply, "maintenance.templates.manage")) return;
    const templateId = Number(req.params?.id || 0);
    const template = db.prepare(`SELECT id FROM service_templates WHERE id = ?`).get(templateId);
    if (!template) return reply.code(404).send({ ok: false, error: "Service template not found" });
    const body = req.body || {};
    const assetId = Number(body.asset_id || 0) || null;
    const manufacturer = String(body.manufacturer || "").trim() || null;
    const model = String(body.model || "").trim() || null;
    const assetCategory = String(body.asset_category || body.category || "").trim() || null;
    if (!assetId && !(manufacturer && model) && !assetCategory) {
      return reply.code(400).send({ ok: false, error: "Assign by exact asset, make and model, or category" });
    }
    if (assetId && !db.prepare(`SELECT id FROM assets WHERE id = ?`).get(assetId)) {
      return reply.code(404).send({ ok: false, error: "Asset not found" });
    }
    const result = db.prepare(`
      INSERT INTO asset_service_template_assignments (
        asset_id, manufacturer, model, asset_category, service_template_id, priority, active
      ) VALUES (?, ?, ?, ?, ?, ?, 1)
    `).run(assetId, manufacturer, model, assetCategory, templateId, Number(body.priority || 100));
    writeAudit(db, req, {
      module: "maintenance", action: "service_template.assign", entity_type: "service_template_assignment",
      entity_id: String(Number(result.lastInsertRowid)), after: { template_id: templateId, asset_id: assetId, manufacturer, model, asset_category: assetCategory },
    });
    return reply.code(201).send({ ok: true, id: Number(result.lastInsertRowid) });
  });

  // =====================================================
  // COSTED SERVICE PLANNER
  // =====================================================
  app.get("/service-planner", async (req, reply) => {
    if (!requireMaintenancePermission(req, reply, "maintenance.templates.read")) return;
    try {
      const nearDueHours = Math.max(1, Number(req.query?.near_due_hours || 50));
      const includeAll = String(req.query?.include_all || "") === "1";
      const dueRows = buildDueListFromPlans(listMaintenancePlans(nearDueHours), getAssetCurrentHours, nearDueHours)
        .filter((row) => includeAll || String(row.status || "OK") !== "OK");
      const latestEstimate = db.prepare(`
        SELECT se.* FROM service_estimates se
        WHERE se.maintenance_plan_id = ?
        ORDER BY se.id DESC LIMIT 1
      `);
      const rows = dueRows.map((row) => {
        const preview = buildServiceEstimatePreview(db, {
          assetId: Number(row.asset_id), planId: Number(row.plan_id), meterReading: Number(row.current_hours || 0),
          intervalHours: Number(row.next_service_interval || row.interval_hours || 0), meterUnit: row.meter_unit || meterUnitForAsset(row.asset_code),
        });
        const latest = latestEstimate.get(Number(row.plan_id)) || null;
        return {
          ...row,
          template_status: preview.resolution.status,
          template_name: preview.estimate?.template_name || null,
          template_revision: preview.estimate?.template_revision || null,
          pricing_complete: preview.estimate?.pricing_complete ?? false,
          stock_available: preview.estimate?.stock_available ?? false,
          estimate_total: preview.estimate?.estimated_total_cost ?? null,
          warnings: preview.warnings,
          latest_estimate: latest,
          borris_advisory: {
            enabled: true,
            action: preview.resolution.status === "ambiguous" ? "review_assignment" : preview.resolution.status === "missing" ? "create_template" : "review_cost_and_stock",
            confidence: preview.resolution.status === "matched" ? "high" : "needs_review",
            evidence: ["live meter", "maintenance plan", "template assignment", "stock pricing"],
            can_mutate: false,
          },
        };
      });
      return reply.send({ ok: true, near_due_hours: nearDueHours, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || "Unable to load costed planner" });
    }
  });

  app.post("/service-estimates", async (req, reply) => {
    if (!requireMaintenancePermission(req, reply, "maintenance.estimates.create")) return;
    try {
      const body = req.body || {};
      const planId = Number(body.plan_id || 0);
      const plan = db.prepare(`
        SELECT mp.*, a.asset_code FROM maintenance_plans mp
        JOIN assets a ON a.id = mp.asset_id WHERE mp.id = ? AND mp.active = 1
      `).get(planId);
      if (!plan) return reply.code(404).send({ ok: false, error: "Active maintenance plan not found" });
      const meterReading = getAssetCurrentHours(Number(plan.asset_id));
      const hasHoursOverride = body.labour_hours != null && String(body.labour_hours).trim() !== "";
      const hasRateOverride = body.labour_rate != null && String(body.labour_rate).trim() !== "";
      const preview = buildServiceEstimatePreview(db, {
        assetId: Number(plan.asset_id), planId, meterReading, intervalHours: Number(plan.interval_hours),
        meterUnit: meterUnitForAsset(plan.asset_code),
        labourHours: hasHoursOverride ? body.labour_hours : null,
        labourRate: hasRateOverride ? body.labour_rate : null,
      });
      if (preview.resolution.status !== "matched") {
        return reply.code(409).send({ ok: false, error: preview.warnings[0], resolution: preview.resolution });
      }
      const changingLabour = preview.estimate.labour_hours !== preview.estimate.labour_hours_original
        || preview.estimate.labour_rate !== preview.estimate.labour_rate_original;
      const overrideReason = String(body.override_reason || "").trim() || null;
      if (changingLabour && !overrideReason) {
        return reply.code(400).send({ ok: false, error: "A reason is required when labour hours or rate is overridden" });
      }
      const persist = db.transaction(() => {
        const e = preview.estimate;
        const inserted = db.prepare(`
          INSERT INTO service_estimates (
            asset_id, service_template_id, maintenance_plan_id, meter_reading, template_name, template_revision,
            status, pricing_complete, estimated_parts_cost, estimated_oil_cost, estimated_consumables_cost,
            estimated_labour_cost, estimated_total_cost, labour_hours_original, labour_rate_original,
            labour_hours_override, labour_rate_override, override_reason, price_date, created_by
          ) VALUES (?, ?, ?, ?, ?, ?, 'draft', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, date('now'), ?)
        `).run(
          e.asset_id, e.service_template_id, planId, e.meter_reading, e.template_name, e.template_revision,
          e.pricing_complete ? 1 : 0, e.estimated_parts_cost, e.estimated_oil_cost, e.estimated_consumables_cost,
          e.estimated_labour_cost, e.estimated_total_cost, e.labour_hours_original, e.labour_rate_original,
          changingLabour ? e.labour_hours : null, changingLabour ? e.labour_rate : null, overrideReason, maintenanceActor(req),
        );
        const estimateId = Number(inserted.lastInsertRowid);
        const insertLine = db.prepare(`
          INSERT INTO service_estimate_items (
            service_estimate_id, stock_item_id, stock_part_code, description, item_type, quantity_required,
            unit_of_measure, unit_cost, estimated_line_total, stock_quantity_available, shortage_quantity,
            source_price_date, price_status, required, allow_substitute
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `);
        preview.items.forEach((item) => insertLine.run(
          estimateId, item.stock_item_id, item.stock_part_code, item.description, item.item_type, item.quantity_required,
          item.unit_of_measure, item.unit_cost, item.estimated_line_total, item.stock_quantity_available, item.shortage_quantity,
          item.source_price_date, item.price_status, item.required ? 1 : 0, item.allow_substitute ? 1 : 0,
        ));
        return estimateId;
      });
      const id = persist();
      writeAudit(db, req, {
        module: "maintenance", action: "service_estimate.create", entity_type: "service_estimate", entity_id: String(id),
        payload: { plan_id: planId, template_id: preview.estimate.service_template_id, pricing_complete: preview.estimate.pricing_complete },
      });
      return reply.code(201).send({ ok: true, id, status: "draft", preview });
    } catch (err) {
      req.log.error(err);
      return reply.code(400).send({ ok: false, error: err.message || "Unable to create service estimate" });
    }
  });

  app.get("/service-estimates/:id", async (req, reply) => {
    if (!requireMaintenancePermission(req, reply, "maintenance.templates.read")) return;
    const estimate = db.prepare(`SELECT * FROM service_estimates WHERE id = ?`).get(Number(req.params?.id || 0));
    if (!estimate) return reply.code(404).send({ ok: false, error: "Service estimate not found" });
    const items = db.prepare(`SELECT * FROM service_estimate_items WHERE service_estimate_id = ? ORDER BY id`).all(Number(estimate.id));
    return reply.send({ ok: true, estimate, items });
  });

  app.post("/service-estimates/:id/approve", async (req, reply) => {
    if (!requireMaintenancePermission(req, reply, "maintenance.estimates.approve")) return;
    const id = Number(req.params?.id || 0);
    const estimate = db.prepare(`SELECT * FROM service_estimates WHERE id = ?`).get(id);
    if (!estimate) return reply.code(404).send({ ok: false, error: "Service estimate not found" });
    if (String(estimate.status) !== "draft") return reply.code(409).send({ ok: false, error: "Only draft estimates can be approved" });
    if (Number(estimate.pricing_complete) !== 1) {
      return reply.code(409).send({ ok: false, error: "Price required before a service estimate can be approved" });
    }
    db.prepare(`UPDATE service_estimates SET status = 'approved', approved_by = ?, approved_at = datetime('now'), updated_at = datetime('now') WHERE id = ?`)
      .run(maintenanceActor(req), id);
    writeAudit(db, req, { module: "maintenance", action: "service_estimate.approve", entity_type: "service_estimate", entity_id: String(id) });
    return reply.send({ ok: true, id, status: "approved" });
  });

  app.post("/service-estimates/:id/convert", async (req, reply) => {
    if (!requireMaintenancePermission(req, reply, "maintenance.estimates.convert")) return;
    try {
      const id = Number(req.params?.id || 0);
      const reserveStock = req.body?.reserve_stock === true;
      if (reserveStock && !requireMaintenancePermission(req, reply, "maintenance.stock.reserve")) return;
      const estimate = db.prepare(`SELECT * FROM service_estimates WHERE id = ?`).get(id);
      if (!estimate) return reply.code(404).send({ ok: false, error: "Service estimate not found" });
      if (String(estimate.status) !== "approved") return reply.code(409).send({ ok: false, error: "Approved service estimate required before creating work order" });
      const items = db.prepare(`SELECT * FROM service_estimate_items WHERE service_estimate_id = ? ORDER BY id`).all(id);
      if (reserveStock && items.some((item) => Number(item.required) === 1 && Number(item.shortage_quantity || 0) > 0)) {
        return reply.code(409).send({ ok: false, error: "Cannot reserve stock while required material is short" });
      }
      const convert = db.transaction(() => {
        const existing = db.prepare(`
          SELECT id FROM work_orders
          WHERE source = 'service' AND reference_id = ?
            AND LOWER(COALESCE(status, 'open')) NOT IN ('closed', 'completed', 'approved', 'cancelled')
          ORDER BY id DESC LIMIT 1
        `).get(Number(estimate.maintenance_plan_id));
        const workOrderId = existing ? Number(existing.id) : Number(db.prepare(`
          INSERT INTO work_orders (asset_id, source, reference_id, status)
          VALUES (?, 'service', ?, 'open')
        `).run(Number(estimate.asset_id), Number(estimate.maintenance_plan_id)).lastInsertRowid);
        db.prepare(`
          UPDATE work_orders SET service_estimate_id = ?, service_template_id = ?, service_template_revision = ?,
            planned_labor_hours = ?, planned_labor_rate = ? WHERE id = ?
        `).run(id, estimate.service_template_id, estimate.template_revision,
          estimate.labour_hours_override ?? estimate.labour_hours_original,
          estimate.labour_rate_override ?? estimate.labour_rate_original, workOrderId);
        db.prepare(`DELETE FROM work_order_planned_materials WHERE work_order_id = ?`).run(workOrderId);
        const materialInsert = db.prepare(`
          INSERT INTO work_order_planned_materials (
            work_order_id, service_estimate_item_id, part_id, part_code, description, item_type,
            quantity_planned, unit_of_measure, unit_cost_snapshot, planned_line_total, source_price_date
          ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        `);
        items.forEach((item) => materialInsert.run(
          workOrderId, item.id, item.stock_item_id, item.stock_part_code, item.description, item.item_type,
          item.quantity_required, item.unit_of_measure, item.unit_cost, item.estimated_line_total, item.source_price_date,
        ));
        if (reserveStock) {
          const reserve = db.prepare(`
            INSERT INTO stock_reservations (work_order_id, part_id, quantity_reserved, quantity_issued, status, reserved_by, notes)
            VALUES (?, ?, ?, 0, 'active', ?, 'Reserved from approved service estimate')
            ON CONFLICT(work_order_id, part_id) DO UPDATE SET
              quantity_reserved = excluded.quantity_reserved, quantity_issued = 0, status = 'active',
              reserved_by = excluded.reserved_by, reserved_at = datetime('now'), released_at = NULL
          `);
          items.filter((item) => Number(item.required) === 1 && Number(item.stock_item_id || 0) > 0 && Number(item.quantity_required || 0) > 0)
            .forEach((item) => reserve.run(workOrderId, item.stock_item_id, item.quantity_required, maintenanceActor(req)));
        }
        db.prepare(`UPDATE service_estimates SET status = 'converted', work_order_id = ?, updated_at = datetime('now') WHERE id = ?`).run(workOrderId, id);
        return workOrderId;
      });
      const workOrderId = convert();
      writeAudit(db, req, {
        module: "maintenance", action: "service_estimate.convert", entity_type: "service_estimate", entity_id: String(id),
        payload: { work_order_id: workOrderId, stock_reserved: reserveStock },
      });
      return reply.send({ ok: true, id, work_order_id: workOrderId, stock_reserved: reserveStock });
    } catch (err) {
      req.log.error(err);
      return reply.code(400).send({ ok: false, error: err.message || "Unable to create planned work order" });
    }
  });
}
