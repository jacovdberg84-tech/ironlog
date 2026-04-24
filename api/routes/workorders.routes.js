// IRONLOG/api/routes/workorders.routes.js
import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";

export default async function workOrderRoutes(app) {
  ensureAuditTable(db);
  db.prepare(`
    CREATE TABLE IF NOT EXISTS work_order_qr_profiles (
      work_order_id INTEGER PRIMARY KEY,
      qr_payload TEXT NOT NULL,
      qr_text TEXT NOT NULL,
      generated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (work_order_id) REFERENCES work_orders(id) ON DELETE CASCADE
    )
  `).run();

  function hasTable(tableName) {
    const row = db.prepare(`
      SELECT name
      FROM sqlite_master
      WHERE type = 'table' AND name = ?
    `).get(tableName);
    return Boolean(row);
  }
  function hasColumn(tableName, columnName) {
    if (!hasTable(tableName)) return false;
    const cols = db.prepare(`PRAGMA table_info(${tableName})`).all();
    return cols.some((c) => String(c.name || "") === String(columnName));
  }
  function firstExistingColumn(tableName, candidates) {
    for (const c of candidates) {
      if (hasColumn(tableName, c)) return c;
    }
    return null;
  }
  function resolveWebOrigin(req) {
    const envBase = String(process.env.IRONLOG_PUBLIC_BASE_URL || "").trim().replace(/\/+$/, "");
    if (envBase) return envBase;
    const protoHeader = String(req.headers["x-forwarded-proto"] || "").split(",")[0].trim();
    const hostHeader = String(req.headers["x-forwarded-host"] || req.headers.host || "").split(",")[0].trim();
    const proto = protoHeader || "http";
    if (hostHeader) return `${proto}://${hostHeader}`;
    return "";
  }

  function getRole(req) {
    return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
  }

  function requireRoles(req, reply, roles) {
    const role = getRole(req);
    if (!roles.includes(role)) {
      reply.code(403).send({ error: `role '${role || "unknown"}' not allowed` });
      return false;
    }
    return true;
  }

  function canRoleTransition(role, currentStatus, nextStatus) {
    const r = String(role || "").toLowerCase();
    if (r === "admin" || r === "supervisor") return true;
    if (r === "artisan") {
      const allowed = {
        assigned: ["in_progress"],
        in_progress: ["completed", "assigned"],
        completed: ["in_progress"],
      };
      return (allowed[currentStatus] || []).includes(nextStatus);
    }
    return false;
  }

  function hasColumn(table, col) {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name) === col);
  }

  function ensureColumn(table, colName, colDef) {
    if (!hasColumn(table, colName)) {
      db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
    }
  }

  // Backward-compatible schema upgrades for WO completion/sign-off.
  ensureColumn("work_orders", "completion_notes", "completion_notes TEXT");
  ensureColumn("work_orders", "artisan_name", "artisan_name TEXT");
  ensureColumn("work_orders", "artisan_signed_at", "artisan_signed_at TEXT");
  ensureColumn("work_orders", "supervisor_name", "supervisor_name TEXT");
  ensureColumn("work_orders", "supervisor_signed_at", "supervisor_signed_at TEXT");
  ensureColumn("work_orders", "completed_at", "completed_at TEXT");
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

  function getAssetCurrentHours(assetId) {
    const fromAssetHours = db.prepare(`
      SELECT total_hours
      FROM asset_hours
      WHERE asset_id = ?
    `).get(assetId);

    if (fromAssetHours && fromAssetHours.total_hours != null) {
      return Number(fromAssetHours.total_hours || 0);
    }

    const fromDailyHours = db.prepare(`
      SELECT COALESCE(SUM(hours_run), 0) AS total_hours
      FROM daily_hours
      WHERE asset_id = ?
        AND is_used = 1
        AND hours_run > 0
    `).get(assetId);

    return Number(fromDailyHours?.total_hours || 0);
  }

  const getStoredWoQr = db.prepare(`
    SELECT qr_payload, qr_text, generated_at
    FROM work_order_qr_profiles
    WHERE work_order_id = ?
  `);
  const upsertWoQr = db.prepare(`
    INSERT INTO work_order_qr_profiles (work_order_id, qr_payload, qr_text, generated_at)
    VALUES (?, ?, ?, datetime('now'))
    ON CONFLICT(work_order_id) DO UPDATE SET
      qr_payload = excluded.qr_payload,
      qr_text = excluded.qr_text,
      generated_at = datetime('now')
  `);

  function inferMakeModel(assetName, assetCode) {
    const name = String(assetName || "").trim();
    const code = String(assetCode || "").trim();
    const tokens = name.split(/\s+/).filter(Boolean);
    const make = tokens[0] ? String(tokens[0]).toUpperCase() : null;
    let model = null;
    if (tokens.length >= 2) {
      const second = String(tokens[1] || "");
      if (/[0-9]/.test(second) || second.length <= 12) model = second.toUpperCase();
    }
    if (!model && code) {
      const codeToken = code.split(/[-_\s]/).find((t) => /[0-9]/.test(t));
      if (codeToken) model = codeToken.toUpperCase();
    }
    return { make: make || null, model: model || null };
  }

  function buildWorkOrderQrProfile(wo, req) {
    const makeCol = firstExistingColumn("assets", ["make", "asset_make", "manufacturer", "brand"]);
    const modelCol = firstExistingColumn("assets", ["model", "asset_model"]);
    let assetMake = null;
    let assetModel = null;
    if (makeCol || modelCol) {
      const fields = [makeCol ? `${makeCol} AS make` : "NULL AS make", modelCol ? `${modelCol} AS model` : "NULL AS model"].join(", ");
      const row = db.prepare(`SELECT ${fields} FROM assets WHERE id = ?`).get(wo.asset_id);
      assetMake = row?.make != null ? String(row.make).trim() || null : null;
      assetModel = row?.model != null ? String(row.model).trim() || null : null;
    }
    if (!assetMake || !assetModel) {
      const inferred = inferMakeModel(wo.asset_name, wo.asset_code);
      if (!assetMake) assetMake = inferred.make;
      if (!assetModel) assetModel = inferred.model;
    }

    const origin = resolveWebOrigin(req);
    const scanUrl = origin
      ? `${origin}/web/workorder-qr.html?wo_id=${encodeURIComponent(String(wo.id))}`
      : `/web/workorder-qr.html?wo_id=${encodeURIComponent(String(wo.id))}`;

    const profile = {
      generated_at: new Date().toISOString(),
      work_order: {
        id: Number(wo.id),
        source: String(wo.source || ""),
        status: String(wo.status || ""),
        opened_at: wo.opened_at || null,
        closed_at: wo.closed_at || null,
      },
      asset: {
        asset_code: wo.asset_code || null,
        asset_name: wo.asset_name || null,
        category: wo.category || null,
        make: assetMake,
        model: assetModel,
      },
      scan_url: scanUrl,
    };

    const qrText = [
      `IRONLOG WO #${wo.id}`,
      `Scan URL: ${scanUrl}`,
      `Asset: ${wo.asset_code || "-"}`,
      `Status: ${String(wo.status || "").toUpperCase()}`,
      `Source: ${String(wo.source || "")}`,
    ].join("\n");

    return { profile, qrText };
  }

  // List work orders (filter by status optional)
  app.get("/", async (req) => {
    const status = (req.query?.status ? String(req.query.status) : "").trim();

    const baseSql = `
      SELECT
        w.id,
        w.source,
        w.reference_id,
        w.status,
        CASE
          WHEN w.source = 'breakdown' THEN COALESCE(NULLIF(TRIM(b.start_at), ''), NULLIF(TRIM(b.breakdown_date), ''), w.opened_at)
          ELSE w.opened_at
        END AS opened_at,
        w.closed_at,
        a.asset_code,
        a.asset_name,
        a.category
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
    `;

    const rows = status
      ? db.prepare(baseSql + ` WHERE w.status = ? ORDER BY w.id DESC LIMIT 200`).all(status)
      : db.prepare(baseSql + ` ORDER BY w.id DESC LIMIT 200`).all();

    return rows;
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
      SELECT id, status
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

    db.prepare(`
      UPDATE work_orders
      SET status = ?
      WHERE id = ?
    `).run(nextStatus, id);

    writeAudit(db, req, {
      module: "workorders",
      action: "status_change",
      entity_type: "work_order",
      entity_id: id,
      payload: { from: currentStatus, to: nextStatus },
    });

    return reply.send({ ok: true, id, from: currentStatus, status: nextStatus });
  });

  // Work order detail (includes linked breakdown if source=breakdown)
  app.get("/:id", async (req, reply) => {
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });

    const wo = db.prepare(`
      SELECT
        w.*,
        a.asset_code,
        a.asset_name,
        a.category
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE w.id = ?
    `).get(id);

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
        p.part_name
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      WHERE sm.reference = ?
      ORDER BY sm.id ASC
    `).all(`work_order:${id}`);

    return { work_order: wo, breakdown, parts_issued: movements };
  });

  app.get("/:id/qr-profile", async (req, reply) => {
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });

    const wo = db.prepare(`
      SELECT
        w.id, w.asset_id, w.source, w.status, w.opened_at, w.closed_at,
        a.asset_code, a.asset_name, a.category
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE w.id = ?
    `).get(id);
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

    const wo = db.prepare(`
      SELECT
        w.id, w.asset_id, w.source, w.status, w.opened_at, w.closed_at,
        a.asset_code, a.asset_name, a.category
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE w.id = ?
    `).get(id);
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
    if (on_hand < quantity) {
      return reply.code(409).send({
        error: "insufficient stock",
        part_code,
        on_hand,
        requested: quantity
      });
    }

    // Insert movement (negative quantity = out)
    db.prepare(`
      INSERT INTO stock_movements (part_id, quantity, movement_type, reference)
      VALUES (?, ?, 'out', ?)
    `).run(part.id, -Math.abs(Math.trunc(quantity)), `work_order:${id}`);

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

  // Close a work order
  app.post("/:id/close", async (req, reply) => {
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

    const tx = db.transaction(() => {
      closeWorkOrder.run(
        completion_notes,
        artisan_name,
        artisan_name,
        supervisor_name,
        supervisor_name,
        id
      );

      let rolled_plan_id = null;
      let rolled_last_service_hours = null;

      const planId = Number(wo.reference_id || 0);

      if (isServiceWO && planId > 0) {
        const currentHours = getAssetCurrentHours(Number(wo.asset_id || 0));
        const safeHours = Number.isFinite(currentHours) ? Number(currentHours.toFixed(2)) : 0;

        updatePlanLastServiceHours.run(safeHours, planId);
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