// IRONLOG/api/routes/workorders/board.routes.js — Work order list and detail, repair requests, status, progress, costs and QR profiles.
// Registered by routes/workorders.routes.js; shared helpers arrive through ctx.
import { db } from "../../db/client.js";
import { notifyWorkOrderAssigned } from "../../utils/pushNotify.js";
import { stockCategorySql } from "../../utils/stockCategory.js";
import { writeAudit } from "../../utils/audit.js";
import { closeBreakdownIfWorkFinished } from "../../utils/workOrderSync.js";

export default function registerBoardRoutes(app, ctx) {
  const partsRequestsTableExists = () => Boolean(
    db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'maintenance_parts_requests'`).get()
  );

  const {
    buildWorkOrderQrProfile,
    canRoleTransition,
    enrichWorkOrderCosts,
    firstExistingColumn,
    getRole,
    getSiteCode,
    getStoredWoQr,
    hasColumn,
    hasTable,
    readLaborRateDefault,
    requireRoles,
    resolveAssignedUsername,
    technicianMatchesUser,
    upsertWoQr,
  } = ctx;

  // List work orders (filter by status / date range / equipment type optional)
  app.get("/", async (req) => {
    const status = (req.query?.status ? String(req.query.status) : "").trim();
    const fromDate = String(req.query?.from_date || "").trim().slice(0, 10);
    const toDate = String(req.query?.to_date || "").trim().slice(0, 10);
    const category = String(req.query?.category || req.query?.equipment_type || "").trim();
    // active=1: every job not yet closed, however old (the board view).
    const activeOnly = String(req.query?.active || "").trim() === "1";
    const siteCode = getSiteCode(req);

    const openedExpr = `
      CASE
        WHEN w.source = 'breakdown' THEN COALESCE(NULLIF(TRIM(b.start_at), ''), NULLIF(TRIM(b.breakdown_date), ''), w.opened_at)
        ELSE w.opened_at
      END
    `;

    let sql = `
      SELECT
        w.id,
        w.source,
        w.reference_id,
        w.status,
        w.assigned_artisan_name,
        w.assigned_at,
        w.started_at,
        w.repair_progress,
        w.repair_progress_at,
        w.labor_hours,
        w.labor_rate_per_hour,
        w.oil_cost,
        w.completed_at,
        w.artisan_name,
        w.supervisor_name,
        ${openedExpr} AS opened_at,
        w.closed_at,
        w.due_date,
        w.priority,
        w.job_description,
        b.description AS breakdown_description,
        b.component AS breakdown_component,
        b.critical AS breakdown_critical,
        b.parts_status,
        b.ets_repair_date,
        mp.service_name,
        a.asset_code,
        a.asset_name,
        a.category
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
      LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id AND w.source = 'service'
      WHERE LOWER(TRIM(COALESCE(w.site_code, 'main'))) = ?
    `;
    const params = [siteCode];

    if (status) {
      sql += ` AND w.status = ?`;
      params.push(status);
    } else if (activeOnly) {
      sql += ` AND REPLACE(TRIM(LOWER(COALESCE(w.status, 'open'))), ' ', '_') <> 'closed'`;
    }
    if (category) {
      sql += ` AND LOWER(TRIM(COALESCE(a.category, ''))) = LOWER(?)`;
      params.push(category);
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(fromDate)) {
      sql += ` AND date(${openedExpr}) >= date(?)`;
      params.push(fromDate);
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(toDate)) {
      sql += ` AND date(${openedExpr}) <= date(?)`;
      params.push(toDate);
    }

    const limit = fromDate || toDate || category || activeOnly ? 2000 : 200;
    sql += ` ORDER BY w.id DESC LIMIT ${limit}`;

    let rows = db.prepare(sql).all(...params);

    // Parts the job is still waiting on (requested or ordered, not yet received).
    if (rows.length && partsRequestsTableExists()) {
      const waiting = db.prepare(`
        SELECT work_order_id, COUNT(*) AS n, MIN(part_name) AS first_part
        FROM maintenance_parts_requests
        WHERE work_order_id IS NOT NULL
          AND LOWER(COALESCE(status, 'requested')) IN ('requested', 'ordered')
        GROUP BY work_order_id
      `).all();
      const byWo = new Map(waiting.map((w) => [Number(w.work_order_id), w]));
      rows = rows.map((r) => {
        const w = byWo.get(Number(r.id));
        return w ? { ...r, open_parts_requests: Number(w.n), first_waiting_part: w.first_part } : r;
      });
    }

    const role = getRole(req);
    const userName = String(req.headers["x-user-name"] || "").trim().toLowerCase();
    if (role === "artisan" && userName) {
      rows = rows.filter((r) =>
        technicianMatchesUser(r.assigned_artisan_name, userName)
      );
    }

    return rows;
  });

  app.post("/repair", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;

    const asset_code = String(req.body?.asset_code || "").trim();
    const description = String(req.body?.description || req.body?.issue || "").trim();
    const component = String(req.body?.component || "").trim() || null;
    const inspection_id = Number(req.body?.inspection_id || 0);
    const assignedRaw = String(req.body?.assigned_artisan_name || "").trim();
    const site_code = getSiteCode(req);

    if (!asset_code) return reply.code(400).send({ error: "asset_code is required" });
    if (!description) return reply.code(400).send({ error: "description is required" });

    const asset = db.prepare(`
      SELECT id, asset_code, asset_name
      FROM assets
      WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?))
      LIMIT 1
    `).get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    if (inspection_id > 0 && hasTable("manager_inspections")) {
      const insp = db.prepare(`SELECT id, work_order_id FROM manager_inspections WHERE id = ?`).get(inspection_id);
      if (!insp) return reply.code(404).send({ error: "Inspection not found" });
      if (insp.work_order_id) {
        return reply.send({
          ok: true,
          work_order_id: Number(insp.work_order_id),
          already_exists: true,
          status: db.prepare(`SELECT status FROM work_orders WHERE id = ?`).get(insp.work_order_id)?.status || null,
        });
      }
    }

    const source = inspection_id > 0 ? "inspection" : "manual";
    const reference_id = inspection_id > 0 ? inspection_id : null;
    const assigned_artisan_name = assignedRaw ? resolveAssignedUsername(assignedRaw) : null;
    const initialStatus = assigned_artisan_name ? "assigned" : "open";
    const assignedBy = assigned_artisan_name
      ? String(req.headers["x-user-name"] || "supervisor").trim() || "supervisor"
      : null;

    const noteLines = [
      inspection_id > 0 ? `Inspection repair #${inspection_id}` : "Repair work order (no breakdown logged)",
      `Asset: ${String(asset.asset_code || "")} — ${String(asset.asset_name || "")}`,
      component ? `Component: ${component}` : null,
      "",
      description,
    ].filter((x) => x !== null);
    const completion_notes = noteLines.join("\n").trim();

    const ins = db.prepare(`
      INSERT INTO work_orders (asset_id, source, reference_id, status, site_code)
      VALUES (?, ?, ?, ?, ?)
    `).run(asset.id, source, reference_id, initialStatus, site_code);
    const work_order_id = Number(ins.lastInsertRowid);

    if (hasColumn("work_orders", "job_description") && completion_notes) {
      db.prepare(`UPDATE work_orders SET job_description = ? WHERE id = ?`).run(completion_notes, work_order_id);
    }
    if (assigned_artisan_name && hasColumn("work_orders", "assigned_artisan_name")) {
      db.prepare(`
        UPDATE work_orders
        SET assigned_artisan_name = ?, assigned_at = datetime('now'), assigned_by = ?
        WHERE id = ?
      `).run(assigned_artisan_name, assignedBy, work_order_id);
    }

    if (inspection_id > 0 && hasTable("manager_inspections") && hasColumn("manager_inspections", "work_order_id")) {
      db.prepare(`UPDATE manager_inspections SET work_order_id = ? WHERE id = ?`).run(work_order_id, inspection_id);
    }

    writeAudit(db, req, {
      module: "workorders",
      action: "repair.create",
      entity_type: "work_order",
      entity_id: work_order_id,
      payload: { asset_code, source, inspection_id: inspection_id || null, assigned_artisan_name },
    });

    if (assigned_artisan_name) {
      notifyWorkOrderAssigned({
        workOrderId: work_order_id,
        assignedUsername: assigned_artisan_name,
        assetCode: asset.asset_code,
        source,
      }).catch((err) => console.error("[push] repair wo assign:", err?.message || err));
    }

    return reply.code(201).send({
      ok: true,
      work_order_id,
      status: initialStatus,
      source,
      assigned_artisan_name,
      breakdown_logged: false,
    });
  });

  app.get("/inspection-quality", async () => {
    const hasMi = hasTable("manager_inspections");
    if (!hasMi) {
      return {
        ok: true,
        score: { completeness: 0, photo_evidence: 0, comment_quality: 0, repeat_issue_rate: 0, overall: 0 },
        sample_size: 0,
      };
    }
    const notesCol = firstExistingColumn("manager_inspections", ["notes", "note", "remarks", "description"]);
    const checklistCol = hasColumn("manager_inspections", "checklist_json") ? "checklist_json" : null;
    const createdCol = firstExistingColumn("manager_inspections", ["created_at", "inspection_date", "date"]);
    const rows = db.prepare(`
      SELECT
        id,
        ${notesCol ? `${notesCol} AS notes` : "NULL AS notes"},
        ${checklistCol ? `${checklistCol} AS checklist_json` : "NULL AS checklist_json"},
        ${createdCol ? `${createdCol} AS created_at` : "NULL AS created_at"}
      FROM manager_inspections
      ORDER BY id DESC
      LIMIT 250
    `).all();
    const sample = Array.isArray(rows) ? rows : [];
    const total = sample.length || 1;
    let completeCount = 0;
    let withPhoto = 0;
    let goodComments = 0;
    const issueCounter = new Map();
    for (const row of sample) {
      const notes = String(row?.notes || "").trim();
      const checklistRaw = String(row?.checklist_json || "").trim();
      let items = [];
      if (checklistRaw) {
        try { items = JSON.parse(checklistRaw); } catch { items = []; }
      }
      const arr = Array.isArray(items) ? items : [];
      const allAnswered = arr.length > 0 ? arr.every((it) => String(it?.value || it?.answer || "").trim() !== "") : false;
      const hasComment = notes.length >= 12;
      if (allAnswered && hasComment) completeCount += 1;
      if (hasComment) goodComments += 1;
      const hasInlinePhoto = arr.some((it) => {
        const v = String(it?.photo_url || it?.image || "").trim();
        return v.length > 0;
      });
      if (hasInlinePhoto) withPhoto += 1;
      for (const it of arr) {
        const label = String(it?.label || it?.item || "").trim().toLowerCase();
        const val = String(it?.value || it?.answer || "").trim().toLowerCase();
        if (!label) continue;
        if (val === "fail" || val === "no" || val === "not_ok") {
          issueCounter.set(label, Number(issueCounter.get(label) || 0) + 1);
        }
      }
    }
    const repeated = [...issueCounter.values()].filter((c) => c > 1).length;
    const repeatIssueRate = Number(((repeated / Math.max(1, issueCounter.size)) * 100).toFixed(2));
    const completeness = Number(((completeCount / total) * 100).toFixed(2));
    const photoEvidence = Number(((withPhoto / total) * 100).toFixed(2));
    const commentQuality = Number(((goodComments / total) * 100).toFixed(2));
    const overall = Number((completeness * 0.35 + photoEvidence * 0.25 + commentQuality * 0.25 + (100 - repeatIssueRate) * 0.15).toFixed(2));
    return {
      ok: true,
      sample_size: sample.length,
      score: {
        completeness,
        photo_evidence: photoEvidence,
        comment_quality: commentQuality,
        repeat_issue_rate: repeatIssueRate,
        overall,
      },
    };
  });

  // Work order status transitions
  // Body: { status }
  app.post("/:id/status", async (req, reply) => {
    const role = getRole(req);
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });

    const nextStatus = String(req.body?.status || "").trim().toLowerCase();
    const allowedStatuses = ["open", "assigned", "in_progress", "completed", "approved", "closed"];
    if (!allowedStatuses.includes(nextStatus)) {
      return reply.code(400).send({
        error: `status must be one of: ${allowedStatuses.join(", ")}`
      });
    }

    const wo = db.prepare(`
      SELECT id, status, assigned_artisan_name, artisan_name, source, reference_id
      FROM work_orders
      WHERE id = ?
    `).get(id);
    if (!wo) return reply.code(404).send({ error: "work order not found" });

    const currentStatus = String(wo.status || "").toLowerCase();
    const transitions = {
      open: ["assigned", "in_progress", "closed"],
      assigned: ["in_progress", "open", "closed"],
      in_progress: ["completed", "assigned", "closed"],
      completed: ["approved", "in_progress", "closed"],
      approved: ["closed", "completed"],
      closed: [],
    };

    if (currentStatus === nextStatus) {
      return reply.send({ ok: true, id, status: currentStatus, unchanged: true });
    }

    const canMove = (transitions[currentStatus] || []).includes(nextStatus);
    if (!canMove) {
      return reply.code(409).send({
        error: `invalid transition from ${currentStatus} to ${nextStatus}`
      });
    }

    if (!canRoleTransition(role, currentStatus, nextStatus)) {
      return reply.code(403).send({
        error: `role '${role}' cannot transition ${currentStatus} -> ${nextStatus}`
      });
    }

    const userName = String(req.headers["x-user-name"] || "").trim();
    const assignedName = String(wo.assigned_artisan_name || "").trim();
    if (role === "artisan") {
      if (!assignedName) {
        return reply.code(409).send({ error: "work order is not assigned to a technician yet" });
      }
      if (!technicianMatchesUser(assignedName, userName)) {
        return reply.code(403).send({ error: "this work order is assigned to another technician" });
      }
    }

    const body = req.body || {};
    const completion_notes =
      body.completion_notes != null && String(body.completion_notes).trim() !== ""
        ? String(body.completion_notes).trim()
        : null;
    const artisan_name =
      body.artisan_name != null && String(body.artisan_name).trim() !== ""
        ? String(body.artisan_name).trim()
        : assignedName || userName || null;
    const supervisor_name =
      body.supervisor_name != null && String(body.supervisor_name).trim() !== ""
        ? String(body.supervisor_name).trim()
        : userName || null;

    if (nextStatus === "in_progress" && !["assigned", "open"].includes(currentStatus)) {
      return reply.code(409).send({ error: "work order must be assigned before starting" });
    }
    if (role === "artisan" && nextStatus === "in_progress" && currentStatus !== "assigned") {
      return reply.code(409).send({ error: "technician can only start an assigned work order" });
    }
    if (nextStatus === "completed") {
      if (!completion_notes) {
        return reply.code(400).send({ error: "completion_notes is required when marking a job complete" });
      }
    }
    if (nextStatus === "approved" && !["admin", "supervisor"].includes(role)) {
      return reply.code(403).send({ error: "only a supervisor can approve completed work" });
    }

    if (nextStatus === "in_progress") {
      db.prepare(`
        UPDATE work_orders
        SET
          status = ?,
          started_at = COALESCE(started_at, datetime('now')),
          artisan_name = COALESCE(?, artisan_name, assigned_artisan_name)
        WHERE id = ?
      `).run(nextStatus, artisan_name, id);
    } else if (nextStatus === "completed") {
      const labor_hours = req.body?.labor_hours != null ? Math.max(0, Number(req.body.labor_hours)) : null;
      const labor_rate_per_hour = req.body?.labor_rate_per_hour != null
        ? Math.max(0, Number(req.body.labor_rate_per_hour))
        : null;
      const oil_cost = req.body?.oil_cost != null ? Math.max(0, Number(req.body.oil_cost)) : null;
      db.prepare(`
        UPDATE work_orders
        SET
          status = ?,
          completed_at = datetime('now'),
          completion_notes = COALESCE(?, completion_notes),
          artisan_name = COALESCE(?, artisan_name, assigned_artisan_name),
          artisan_signed_at = datetime('now'),
          labor_hours = COALESCE(?, labor_hours),
          labor_rate_per_hour = COALESCE(?, labor_rate_per_hour),
          oil_cost = COALESCE(?, oil_cost)
        WHERE id = ?
      `).run(nextStatus, completion_notes, artisan_name, labor_hours, labor_rate_per_hour, oil_cost, id);
    } else if (nextStatus === "approved") {
      db.prepare(`
        UPDATE work_orders
        SET
          status = ?,
          supervisor_name = COALESCE(?, supervisor_name),
          supervisor_signed_at = datetime('now')
        WHERE id = ?
      `).run(nextStatus, supervisor_name, id);
    } else if (nextStatus === "closed") {
      db.prepare(`
        UPDATE work_orders
        SET status = ?, closed_at = COALESCE(closed_at, datetime('now'))
        WHERE id = ?
      `).run(nextStatus, id);
    } else {
      db.prepare(`
        UPDATE work_orders
        SET status = ?
        WHERE id = ?
      `).run(nextStatus, id);
    }
    // Finished breakdown work returns the machine to service.
    const breakdownClosed = String(wo.source || "").toLowerCase() === "breakdown" && ["completed", "approved", "closed"].includes(nextStatus)
      ? closeBreakdownIfWorkFinished(db, wo.reference_id)
      : false;

    writeAudit(db, req, {
      module: "workorders",
      action: "status_change",
      entity_type: "work_order",
      entity_id: id,
      payload: {
        from: currentStatus,
        to: nextStatus,
        assigned_artisan_name: assignedName || null,
        completion_notes: completion_notes || null,
      },
    });

    return reply.send({ ok: true, id, from: currentStatus, status: nextStatus, breakdown_closed: breakdownClosed });
  });

  // POST /api/workorders/:id/progress  { repair_progress } — shown on daily PDF for active jobs
  app.post("/:id/progress", async (req, reply) => {
    const role = getRole(req);
    const id = Number(req.params.id);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });

    const wo = db.prepare(`
      SELECT id, status, assigned_artisan_name
      FROM work_orders
      WHERE id = ?
    `).get(id);
    if (!wo) return reply.code(404).send({ error: "work order not found" });

    const status = String(wo.status || "").toLowerCase();
    if (!["open", "assigned", "in_progress"].includes(status)) {
      return reply.code(409).send({ error: "repair progress can only be updated on open, assigned, or in-progress work orders" });
    }

    const userName = String(req.headers["x-user-name"] || "").trim();
    const isSupervisor = ["admin", "supervisor"].includes(role);
    const isAssignedTech = role === "artisan" && technicianMatchesUser(wo.assigned_artisan_name, userName);
    if (!isSupervisor && !isAssignedTech) {
      return reply.code(403).send({ error: "only a supervisor or assigned technician can update repair progress" });
    }

    const repair_progress =
      req.body?.repair_progress != null ? String(req.body.repair_progress).trim() : "";
    db.prepare(`
      UPDATE work_orders
      SET
        repair_progress = ?,
        repair_progress_at = datetime('now')
      WHERE id = ?
    `).run(repair_progress || null, id);

    const row = db.prepare(`
      SELECT repair_progress, repair_progress_at FROM work_orders WHERE id = ?
    `).get(id);

    writeAudit(db, req, {
      module: "workorders",
      action: "repair_progress_update",
      entity_type: "work_order",
      entity_id: id,
      payload: { repair_progress: repair_progress || null },
    });

    return reply.send({
      ok: true,
      id,
      repair_progress: row?.repair_progress || null,
      repair_progress_at: row?.repair_progress_at || null,
    });
  });

  // POST /api/workorders/:id/costs — repair hours, labor rate, manual oil cost, technician
  app.post("/:id/costs", async (req, reply) => {
    const role = getRole(req);
    const id = Number(req.params.id);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });

    const wo = db.prepare(`
      SELECT id, status, assigned_artisan_name
      FROM work_orders
      WHERE id = ?
    `).get(id);
    if (!wo) return reply.code(404).send({ error: "work order not found" });

    const status = String(wo.status || "").toLowerCase();
    if (status === "closed") return reply.code(409).send({ error: "cannot edit costs on a closed work order" });

    const userName = String(req.headers["x-user-name"] || "").trim();
    const isSupervisor = ["admin", "supervisor"].includes(role);
    const isAssignedTech = role === "artisan" && technicianMatchesUser(wo.assigned_artisan_name, userName);
    if (!isSupervisor && !isAssignedTech) {
      return reply.code(403).send({ error: "only a supervisor or assigned technician can update repair costs" });
    }

    const body = req.body || {};
    const labor_hours = body.labor_hours != null ? Math.max(0, Number(body.labor_hours)) : null;
    const labor_rate_per_hour = body.labor_rate_per_hour != null
      ? Math.max(0, Number(body.labor_rate_per_hour))
      : null;
    const oil_cost = body.oil_cost != null ? Math.max(0, Number(body.oil_cost)) : null;
    let assigned_artisan_name = null;
    if (body.assigned_artisan_name != null) {
      if (!isSupervisor) return reply.code(403).send({ error: "only a supervisor can reassign the technician" });
      assigned_artisan_name = resolveAssignedUsername(String(body.assigned_artisan_name || "").trim());
    }

    db.prepare(`
      UPDATE work_orders
      SET
        labor_hours = COALESCE(?, labor_hours),
        labor_rate_per_hour = COALESCE(?, labor_rate_per_hour),
        oil_cost = COALESCE(?, oil_cost),
        assigned_artisan_name = COALESCE(?, assigned_artisan_name)
      WHERE id = ?
    `).run(
      labor_hours,
      labor_rate_per_hour,
      oil_cost,
      assigned_artisan_name,
      id,
    );

    writeAudit(db, req, {
      module: "workorders",
      action: "costs_update",
      entity_type: "work_order",
      entity_id: id,
      payload: {
        labor_hours,
        labor_rate_per_hour,
        oil_cost,
        assigned_artisan_name,
      },
    });

    return reply.send({ ok: true, id, default_labor_rate: readLaborRateDefault() });
  });

  // Work order detail (includes linked breakdown if source=breakdown)
  app.get("/:id", async (req, reply) => {
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });
    const siteCode = getSiteCode(req);

    const wo = db.prepare(`
      SELECT
        w.*,
        a.asset_code,
        a.asset_name,
        a.category
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE w.id = ?
        AND LOWER(TRIM(COALESCE(w.site_code, 'main'))) = ?
    `).get(id, siteCode);

    if (!wo) return reply.code(404).send({ error: "work order not found" });

    let breakdown = null;
    if (wo.source === "breakdown" && wo.reference_id) {
      breakdown = db.prepare(`
        SELECT id, breakdown_date, start_at, description, downtime_total_hours, critical, created_at
        FROM breakdowns
        WHERE id = ?
      `).get(wo.reference_id);
      if (breakdown) breakdown.critical = Boolean(breakdown.critical);
      // Show effective opened date from breakdown date/start instead of WO creation date.
      if (breakdown) {
        wo.opened_at = String(breakdown.start_at || breakdown.breakdown_date || wo.opened_at || "").trim() || wo.opened_at;
      }
    }

    // Parts issued to this WO (from stock_movements reference=work_order:<id>)
    const stockMovementCols = db.prepare(`
      PRAGMA table_info(stock_movements)
    `).all();
    const hasCreatedAt = stockMovementCols.some((c) => String(c.name) === "created_at");
    const movementDateExpr = hasCreatedAt ? "sm.created_at" : "sm.movement_date";

    const movements = db.prepare(`
      SELECT
        sm.id,
        ${movementDateExpr} AS movement_date,
        sm.quantity,
        sm.movement_type,
        sm.reference,
        p.part_code,
        p.part_name,
        p.consumable_kind,
        ${stockCategorySql("p")} AS stock_category,
        p.unit_cost
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      WHERE sm.reference = ?
      ORDER BY sm.id ASC
    `).all(`work_order:${id}`);

    const work_order = enrichWorkOrderCosts(wo, movements);
    const planned_materials = db.prepare(`
      SELECT
        pm.*,
        COALESCE((
          SELECT ABS(SUM(sm.quantity))
          FROM stock_movements sm
          WHERE sm.reference = ?
            AND sm.part_id = pm.part_id
            AND sm.quantity < 0
        ), 0) AS quantity_issued,
        COALESCE((
          SELECT sr.quantity_reserved - sr.quantity_issued
          FROM stock_reservations sr
          WHERE sr.work_order_id = pm.work_order_id
            AND sr.part_id = pm.part_id
            AND sr.status = 'active'
        ), 0) AS quantity_reserved_remaining
      FROM work_order_planned_materials pm
      WHERE pm.work_order_id = ?
      ORDER BY pm.id ASC
    `).all(`work_order:${id}`, id);

    const parts_requests = partsRequestsTableExists()
      ? db.prepare(`
          SELECT id, part_code, part_name, qty, urgency, status, requested_by, notes, status_notes, created_at, updated_at
          FROM maintenance_parts_requests
          WHERE work_order_id = ?
          ORDER BY CASE LOWER(COALESCE(status, 'requested')) WHEN 'requested' THEN 0 WHEN 'ordered' THEN 1 ELSE 2 END, id DESC
        `).all(id)
      : [];

    return {
      work_order,
      breakdown,
      parts_issued: movements,
      parts_requests,
      planned_materials,
      default_labor_rate: readLaborRateDefault(),
    };
  });

  app.get("/:id/qr-profile", async (req, reply) => {
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });

    const siteCode = getSiteCode(req);
    const wo = db.prepare(`
      SELECT
        w.id, w.asset_id, w.source, w.status, w.opened_at, w.closed_at,
        a.asset_code, a.asset_name, a.category
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE w.id = ?
        AND LOWER(TRIM(COALESCE(w.site_code, 'main'))) = ?
    `).get(id, siteCode);
    if (!wo) return reply.code(404).send({ error: "work order not found" });

    const built = buildWorkOrderQrProfile(wo, req);
    const stored = getStoredWoQr.get(id);
    let storedPayload = null;
    if (stored?.qr_payload) {
      try {
        storedPayload = JSON.parse(String(stored.qr_payload || "{}"));
      } catch {
        storedPayload = null;
      }
    }

    return {
      ok: true,
      work_order_id: id,
      stored: stored
        ? { qr_payload: storedPayload, qr_text: stored.qr_text, generated_at: stored.generated_at }
        : null,
      live_preview: built.profile,
      live_qr_text: built.qrText,
    };
  });

  app.post("/:id/qr-profile/refresh", async (req, reply) => {
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });
    const siteCode = getSiteCode(req);

    const wo = db.prepare(`
      SELECT
        w.id, w.asset_id, w.source, w.status, w.opened_at, w.closed_at,
        a.asset_code, a.asset_name, a.category
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE w.id = ?
        AND LOWER(TRIM(COALESCE(w.site_code, 'main'))) = ?
    `).get(id, siteCode);
    if (!wo) return reply.code(404).send({ error: "work order not found" });

    const built = buildWorkOrderQrProfile(wo, req);
    upsertWoQr.run(id, JSON.stringify(built.profile), built.qrText);

    return {
      ok: true,
      work_order_id: id,
      qr_payload: built.profile,
      qr_text: built.qrText,
    };
  });
}
