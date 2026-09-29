// IRONLOG/api/routes/workorders/scheduling.routes.js — Technicians, scheduling board, rules, assignment and escalations.
// Registered by routes/workorders.routes.js; shared helpers arrive through ctx.
import crypto from "crypto";
import { db } from "../../db/client.js";
import { notifyWorkOrderAssigned } from "../../utils/pushNotify.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerSchedulingRoutes(app, ctx) {
  const {
    getSiteCode,
    hasColumn,
    hasTable,
    inferPriorityForRow,
    normalizePriority,
    normalizeShift,
    requireRoles,
    resolveAssignedUsername,
    suggestRuleForWorkOrder,
  } = ctx;

  app.get("/technicians", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "artisan", "stores", "operator"])) return;
    const byUsername = new Map();
    const legacyNames = new Set();

    const addTech = (username, label) => {
      const u = String(username || "").trim();
      const lbl = String(label || u).trim();
      if (!u) return;
      if (!byUsername.has(u)) byUsername.set(u, lbl || u);
      legacyNames.add(lbl || u);
    };

    const rules = db.prepare(`
      SELECT DISTINCT artisan_name
      FROM wo_assignment_rules
      WHERE active = 1 AND TRIM(COALESCE(artisan_name, '')) <> ''
      ORDER BY artisan_name ASC
    `).all();
    for (const r of rules) {
      const n = String(r.artisan_name || "").trim();
      if (!n) continue;
      legacyNames.add(n);
      if (hasTable("users")) {
        const row = db.prepare(`
          SELECT username, full_name FROM users
          WHERE LOWER(TRIM(username)) = LOWER(?)
             OR LOWER(TRIM(COALESCE(full_name, ''))) = LOWER(?)
          LIMIT 1
        `).get(n, n);
        if (row) addTech(row.username, row.full_name || row.username);
        else addTech(n, n);
      } else {
        addTech(n, n);
      }
    }

    if (hasTable("users")) {
      const users = db.prepare(`
        SELECT username, full_name, role, roles_json
        FROM users
        WHERE COALESCE(active, 1) = 1
          AND (
            LOWER(TRIM(COALESCE(role, ''))) = 'artisan'
            OR LOWER(COALESCE(roles_json, '')) LIKE '%artisan%'
          )
      `).all();
      for (const u of users) {
        addTech(u.username, u.full_name || u.username);
      }
    }

    const assigned = db.prepare(`
      SELECT DISTINCT assigned_artisan_name AS name
      FROM work_orders
      WHERE TRIM(COALESCE(assigned_artisan_name, '')) <> ''
      ORDER BY assigned_artisan_name ASC
      LIMIT 200
    `).all();
    for (const r of assigned) {
      const n = String(r.name || "").trim();
      if (!n) continue;
      legacyNames.add(n);
      if (hasTable("users")) {
        const row = db.prepare(`
          SELECT username, full_name FROM users
          WHERE LOWER(TRIM(username)) = LOWER(?)
             OR LOWER(TRIM(COALESCE(full_name, ''))) = LOWER(?)
          LIMIT 1
        `).get(n, n);
        if (row) addTech(row.username, row.full_name || row.username);
        else addTech(n, n);
      } else {
        addTech(n, n);
      }
    }

    const technician_users = [...byUsername.entries()]
      .map(([username, label]) => ({ username, label }))
      .sort((a, b) => String(a.label).localeCompare(String(b.label)));

    return {
      ok: true,
      technician_users,
      technicians: [...legacyNames].filter(Boolean).sort((a, b) => a.localeCompare(b)),
    };
  });

  app.post("/technicians", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    if (!hasTable("users")) return reply.code(503).send({ error: "users table not available" });

    const username = String(req.body?.username || "").trim();
    const full_name = String(req.body?.full_name || "").trim() || username;
    const department = String(req.body?.department || "Workshop").trim() || "Workshop";
    const pin = String(req.body?.pin || "").replace(/\D/g, "");

    if (!username) return reply.code(400).send({ error: "username is required" });
    if (pin.length < 4 || pin.length > 6) {
      return reply.code(400).send({ error: "pin must be 4–6 digits" });
    }

    if (!hasColumn("users", "pin_hash")) {
      db.prepare(`ALTER TABLE users ADD COLUMN pin_hash TEXT`).run();
    }

    const pinHash = (() => {
      const salt = crypto.randomBytes(16);
      const hash = crypto.scryptSync(pin, salt, 64);
      return `scrypt$${salt.toString("base64")}$${hash.toString("base64")}`;
    })();

    const existing = db.prepare(`SELECT username FROM users WHERE lower(username) = lower(?)`).get(username);
    db.prepare(`
      INSERT INTO users (username, full_name, role, active, department, roles_json)
      VALUES (?, ?, 'artisan', 1, ?, ?)
      ON CONFLICT(username) DO UPDATE SET
        full_name = COALESCE(excluded.full_name, users.full_name),
        role = 'artisan',
        active = 1,
        department = COALESCE(excluded.department, users.department),
        roles_json = excluded.roles_json
    `).run(username, full_name, department, JSON.stringify(["artisan"]));
    db.prepare(`UPDATE users SET pin_hash = ? WHERE lower(username) = lower(?)`).run(pinHash, username);

    const ruleExists = db
      .prepare(`
        SELECT id FROM wo_assignment_rules
        WHERE lower(trim(artisan_name)) = lower(trim(?)) AND active = 1
        LIMIT 1
      `)
      .get(username);
    if (!ruleExists) {
      db.prepare(`
        INSERT INTO wo_assignment_rules (artisan_name, skill, location_code, shift, max_open_wos, active)
        VALUES (?, NULL, NULL, NULL, 8, 1)
      `).run(username);
    }

    writeAudit(db, req, {
      module: "workorders",
      action: existing ? "technician.update" : "technician.create",
      entity_type: "user",
      entity_id: username,
      payload: { full_name, department },
    });

    return { ok: true, username, full_name, created: !existing };
  });

  app.get("/schedule/board", async (req) => {
    const artisan = String(req.query?.artisan || "").trim().toLowerCase();
    const shift = normalizeShift(req.query?.shift || "");
    const priority = normalizePriority(req.query?.priority || "P3", "");
    const dueDate = String(req.query?.due_date || "").trim();

    const rows = db.prepare(`
      SELECT
        w.id, w.status, w.opened_at, w.closed_at, w.assigned_artisan_name, w.shift, w.priority, w.due_date, w.required_skill, w.location_code,
        a.asset_code, a.asset_name
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE w.status IN ('open','assigned','in_progress','completed','approved')
        AND LOWER(TRIM(COALESCE(w.site_code, 'main'))) = ?
      ORDER BY w.id DESC
      LIMIT 300
    `).all(getSiteCode(req)).map((r) => ({ ...r, priority: inferPriorityForRow(r) }));

    const filtered = rows.filter((r) => {
      if (artisan && String(r.assigned_artisan_name || "").toLowerCase() !== artisan) return false;
      if (shift && String(r.shift || "").toLowerCase() !== shift) return false;
      if (priority && String(r.priority || "").toUpperCase() !== priority) return false;
      if (dueDate && String(r.due_date || "") !== dueDate) return false;
      return true;
    });
    return { ok: true, rows: filtered };
  });

  app.get("/schedule/rules", async () => {
    const rows = db.prepare(`
      SELECT id, artisan_name, skill, location_code, shift, max_open_wos, active, created_at
      FROM wo_assignment_rules
      ORDER BY active DESC, artisan_name ASC, id DESC
      LIMIT 400
    `).all();
    return { ok: true, rows };
  });

  app.post("/schedule/rules", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const artisan_name = String(req.body?.artisan_name || "").trim();
    if (!artisan_name) return reply.code(400).send({ error: "artisan_name is required" });
    const skill = req.body?.skill != null ? String(req.body.skill).trim() || null : null;
    const location_code = req.body?.location_code != null ? String(req.body.location_code).trim() || null : null;
    const shift = normalizeShift(req.body?.shift);
    const max_open_wos = Math.max(1, Number(req.body?.max_open_wos || 8));
    const active = Number(req.body?.active ?? 1) ? 1 : 0;
    const id = Number(req.body?.id || 0);
    if (id > 0) {
      db.prepare(`
        UPDATE wo_assignment_rules
        SET artisan_name = ?, skill = ?, location_code = ?, shift = ?, max_open_wos = ?, active = ?
        WHERE id = ?
      `).run(artisan_name, skill, location_code, shift, max_open_wos, active, id);
      return { ok: true, id, updated: true };
    }
    const ins = db.prepare(`
      INSERT INTO wo_assignment_rules (artisan_name, skill, location_code, shift, max_open_wos, active)
      VALUES (?, ?, ?, ?, ?, ?)
    `).run(artisan_name, skill, location_code, shift, max_open_wos, active);
    return { ok: true, id: Number(ins.lastInsertRowid), created: true };
  });

  app.post("/schedule/rules/:id/delete", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    db.prepare(`DELETE FROM wo_assignment_rules WHERE id = ?`).run(id);
    return { ok: true, id };
  });

  app.get("/schedule/escalation-config", async () => {
    const row = db.prepare(`
      SELECT overdue_hours, level1_role, level2_role, level3_role, updated_at
      FROM work_order_escalation_config
      WHERE id = 1
    `).get() || { overdue_hours: 8, level1_role: "supervisor", level2_role: "manager", level3_role: "admin" };
    return { ok: true, config: row };
  });

  app.post("/schedule/escalation-config", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const overdue_hours = Math.max(1, Number(req.body?.overdue_hours || 8));
    const level1_role = String(req.body?.level1_role || "supervisor").trim().toLowerCase() || "supervisor";
    const level2_role = String(req.body?.level2_role || "manager").trim().toLowerCase() || "manager";
    const level3_role = String(req.body?.level3_role || "admin").trim().toLowerCase() || "admin";
    db.prepare(`
      UPDATE work_order_escalation_config
      SET overdue_hours = ?, level1_role = ?, level2_role = ?, level3_role = ?, updated_at = datetime('now')
      WHERE id = 1
    `).run(overdue_hours, level1_role, level2_role, level3_role);
    return { ok: true, config: { overdue_hours, level1_role, level2_role, level3_role } };
  });

  app.get("/schedule/escalations", async () => {
    const rows = db.prepare(`
      SELECT
        e.id, e.work_order_id, e.threshold_hours, e.chain_level, e.status, e.created_at, e.detail_json,
        w.assigned_artisan_name, w.priority, w.status AS work_order_status,
        a.asset_code
      FROM work_order_escalations e
      LEFT JOIN work_orders w ON w.id = e.work_order_id
      LEFT JOIN assets a ON a.id = w.asset_id
      ORDER BY e.id DESC
      LIMIT 200
    `).all().map((r) => {
      let detail = null;
      try { detail = r.detail_json ? JSON.parse(String(r.detail_json)) : null; } catch { detail = null; }
      return { ...r, detail };
    });
    return { ok: true, rows };
  });

  app.post("/schedule/escalations/:id/ack", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "manager"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid escalation id" });
    const row = db.prepare(`SELECT id, status FROM work_order_escalations WHERE id = ?`).get(id);
    if (!row) return reply.code(404).send({ error: "escalation not found" });
    if (String(row.status || "").toLowerCase() !== "open") return reply.code(409).send({ error: "escalation is not open" });
    db.prepare(`UPDATE work_order_escalations SET status = 'acknowledged' WHERE id = ?`).run(id);
    writeAudit(db, req, {
      module: "workorders",
      action: "escalation_ack",
      entity_type: "work_order_escalation",
      entity_id: id,
      payload: {},
    });
    return { ok: true, id, status: "acknowledged" };
  });

  app.post("/schedule/escalations/:id/next", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "manager"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid escalation id" });
    const row = db.prepare(`
      SELECT id, work_order_id, threshold_hours, chain_level, status
      FROM work_order_escalations
      WHERE id = ?
    `).get(id);
    if (!row) return reply.code(404).send({ error: "escalation not found" });
    const currentLevel = Math.max(1, Number(row.chain_level || 1));
    if (currentLevel >= 3) return reply.code(409).send({ error: "already at highest escalation level" });

    const cfg = db.prepare(`
      SELECT level1_role, level2_role, level3_role
      FROM work_order_escalation_config WHERE id = 1
    `).get() || { level1_role: "supervisor", level2_role: "manager", level3_role: "admin" };
    const nextLevel = currentLevel + 1;
    const roleByLevel = {
      1: String(cfg.level1_role || "supervisor"),
      2: String(cfg.level2_role || "manager"),
      3: String(cfg.level3_role || "admin"),
    };
    const nextRole = roleByLevel[nextLevel] || "admin";

    const det = db.prepare(`
      SELECT detail_json FROM work_order_escalations WHERE id = ?
    `).get(id);
    let detail = {};
    try { detail = det?.detail_json ? JSON.parse(String(det.detail_json)) : {}; } catch { detail = {}; }
    detail.chain_level = nextLevel;
    detail.escalation_role = nextRole;

    const ins = db.prepare(`
      INSERT INTO work_order_escalations (work_order_id, threshold_hours, chain_level, status, detail_json)
      VALUES (?, ?, ?, 'open', ?)
    `).run(Number(row.work_order_id), Number(row.threshold_hours), Number(nextLevel), JSON.stringify(detail));
    const newEscId = Number(ins.lastInsertRowid || 0);
    db.prepare(`
      INSERT INTO work_order_escalation_notifications (escalation_id, role_target, message)
      VALUES (?, ?, ?)
    `).run(newEscId, nextRole, `WO #${Number(row.work_order_id)} escalated to level ${nextLevel}`);
    db.prepare(`UPDATE work_order_escalations SET status = 'escalated' WHERE id = ?`).run(id);

    writeAudit(db, req, {
      module: "workorders",
      action: "escalation_next_level",
      entity_type: "work_order_escalation",
      entity_id: id,
      payload: { new_escalation_id: newEscId, next_level: nextLevel, next_role: nextRole },
    });
    return { ok: true, id, new_escalation_id: newEscId, next_level: nextLevel, next_role: nextRole };
  });

  app.post("/:id/schedule", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params.id);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const wo = db.prepare(`SELECT id FROM work_orders WHERE id = ?`).get(id);
    if (!wo) return reply.code(404).send({ error: "work order not found" });
    const assigned = req.body?.assigned_artisan_name == null ? null : String(req.body.assigned_artisan_name).trim() || null;
    const shift = normalizeShift(req.body?.shift);
    const priority = normalizePriority(req.body?.priority || "P3");
    const dueDate = req.body?.due_date ? String(req.body.due_date).slice(0, 10) : null;
    const requiredSkill = req.body?.required_skill != null ? String(req.body.required_skill).trim() || null : null;
    const locationCode = req.body?.location_code != null ? String(req.body.location_code).trim() || null : null;

    db.prepare(`
      UPDATE work_orders
      SET
        assigned_artisan_name = COALESCE(?, assigned_artisan_name),
        shift = COALESCE(?, shift),
        priority = COALESCE(?, priority),
        due_date = COALESCE(?, due_date),
        required_skill = COALESCE(?, required_skill),
        location_code = COALESCE(?, location_code),
        status = CASE
          WHEN ? IS NOT NULL AND status = 'open' THEN 'assigned'
          WHEN ? IS NULL AND status = 'assigned' THEN 'open'
          ELSE status
        END
      WHERE id = ?
    `).run(assigned, shift, priority, dueDate, requiredSkill, locationCode, assigned, assigned, id);

    writeAudit(db, req, {
      module: "workorders",
      action: "schedule_update",
      entity_type: "work_order",
      entity_id: id,
      payload: { assigned_artisan_name: assigned, shift, priority, due_date: dueDate, required_skill: requiredSkill, location_code: locationCode },
    });
    return { ok: true, id };
  });

  app.post("/:id/assign", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params.id);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });

    const assigned_artisan_name = resolveAssignedUsername(
      String(req.body?.assigned_artisan_name || "").trim()
    );
    if (!assigned_artisan_name) {
      return reply.code(400).send({ error: "assigned_artisan_name is required" });
    }

    const wo = db.prepare(`SELECT id, status, assigned_artisan_name FROM work_orders WHERE id = ?`).get(id);
    if (!wo) return reply.code(404).send({ error: "work order not found" });

    const status = String(wo.status || "").toLowerCase();
    if (!["open", "assigned"].includes(status)) {
      return reply.code(409).send({ error: "work order can only be assigned while open or already assigned" });
    }

    const assignedBy = String(req.headers["x-user-name"] || "supervisor").trim() || "supervisor";
    db.prepare(`
      UPDATE work_orders
      SET
        assigned_artisan_name = ?,
        assigned_at = datetime('now'),
        assigned_by = ?,
        status = 'assigned'
      WHERE id = ?
    `).run(assigned_artisan_name, assignedBy, id);

    writeAudit(db, req, {
      module: "workorders",
      action: "assign",
      entity_type: "work_order",
      entity_id: id,
      payload: { assigned_artisan_name, assigned_by: assignedBy, from_status: status },
    });

    const woDetail = db.prepare(`
      SELECT wo.source, a.asset_code
      FROM work_orders wo
      JOIN assets a ON a.id = wo.asset_id
      WHERE wo.id = ?
    `).get(id);
    notifyWorkOrderAssigned({
      workOrderId: id,
      assignedUsername: assigned_artisan_name,
      assetCode: woDetail?.asset_code,
      source: woDetail?.source,
    }).catch((err) => console.error("[push] wo assign:", err?.message || err));

    return reply.send({
      ok: true,
      id,
      status: "assigned",
      assigned_artisan_name,
      assigned_by: assignedBy,
    });
  });

  app.post("/schedule/auto-assign", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const rules = db.prepare(`
      SELECT artisan_name, skill, location_code, shift, max_open_wos, active
      FROM wo_assignment_rules
      WHERE active = 1
      ORDER BY id ASC
    `).all();
    if (!rules.length) return { ok: true, assigned_count: 0, note: "No active assignment rules configured" };

    const workloads = new Map();
    const loadRows = db.prepare(`
      SELECT assigned_artisan_name, COUNT(*) AS c
      FROM work_orders
      WHERE status IN ('open','assigned','in_progress')
        AND assigned_artisan_name IS NOT NULL
        AND TRIM(assigned_artisan_name) <> ''
      GROUP BY assigned_artisan_name
    `).all();
    for (const r of loadRows) workloads.set(String(r.assigned_artisan_name), Number(r.c || 0));

    const candidates = db.prepare(`
      SELECT id, required_skill, location_code, shift, assigned_artisan_name
      FROM work_orders
      WHERE status IN ('open','assigned')
      ORDER BY id ASC
      LIMIT 300
    `).all();

    let assignedCount = 0;
    const upd = db.prepare(`
      UPDATE work_orders
      SET assigned_artisan_name = ?, status = CASE WHEN status = 'open' THEN 'assigned' ELSE status END
      WHERE id = ?
    `);
    for (const wo of candidates) {
      if (String(wo.assigned_artisan_name || "").trim()) continue;
      const choice = suggestRuleForWorkOrder(wo, rules, workloads);
      if (!choice) continue;
      upd.run(String(choice.artisan_name), Number(wo.id));
      workloads.set(String(choice.artisan_name), Number(workloads.get(String(choice.artisan_name)) || 0) + 1);
      assignedCount += 1;
    }
    writeAudit(db, req, {
      module: "workorders",
      action: "auto_assign",
      entity_type: "work_order",
      entity_id: null,
      payload: { assigned_count: assignedCount },
    });
    return { ok: true, assigned_count: assignedCount };
  });

  app.post("/schedule/escalations/check", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const cfg = db.prepare(`
      SELECT overdue_hours, level1_role, level2_role, level3_role
      FROM work_order_escalation_config WHERE id = 1
    `).get() || { overdue_hours: 8, level1_role: "supervisor", level2_role: "manager", level3_role: "admin" };
    const threshold = Math.max(1, Number(req.body?.overdue_hours || cfg.overdue_hours || 8));
    const levelRole = {
      1: String(cfg.level1_role || "supervisor"),
      2: String(cfg.level2_role || "manager"),
      3: String(cfg.level3_role || "admin"),
    };
    const nowMs = Date.now();
    const rows = db.prepare(`
      SELECT id, status, opened_at, due_date, assigned_artisan_name, priority
      FROM work_orders
      WHERE status IN ('open','assigned','in_progress','completed','approved')
      ORDER BY id DESC
      LIMIT 500
    `).all();
    let escalated = 0;
    let notifications = 0;
    for (const r of rows) {
      const opened = Date.parse(String(r.opened_at || ""));
      const age = Number.isFinite(opened) ? Math.max(0, Math.floor((nowMs - opened) / 3600000)) : 0;
      const dueAt = Date.parse(String(r.due_date || ""));
      const dueHours = Number.isFinite(dueAt) ? Math.max(0, Math.floor((nowMs - dueAt) / 3600000)) : 0;
      const overdue = Math.max(age, dueHours);
      if (overdue <= threshold) continue;
      const level = overdue > threshold * 3 ? 3 : overdue > threshold * 2 ? 2 : 1;
      const dup = db.prepare(`
        SELECT id FROM work_order_escalations
        WHERE work_order_id = ? AND status = 'open' AND threshold_hours = ? AND chain_level = ?
        ORDER BY id DESC LIMIT 1
      `).get(Number(r.id), Number(threshold), Number(level));
      if (dup) continue;
      const ins = db.prepare(`
        INSERT INTO work_order_escalations (work_order_id, threshold_hours, chain_level, status, detail_json)
        VALUES (?, ?, 1, 'open', ?)
      `).run(Number(r.id), Number(threshold), JSON.stringify({
        work_order_id: Number(r.id),
        status: String(r.status || ""),
        assigned_artisan_name: String(r.assigned_artisan_name || ""),
        priority: normalizePriority(r.priority || "P3"),
        overdue_hours: overdue,
        chain_level: level,
        escalation_role: levelRole[level] || "supervisor",
      }));
      const escalationId = Number(ins.lastInsertRowid || 0);
      if (escalationId) {
        db.prepare(`
          UPDATE work_order_escalations
          SET chain_level = ?
          WHERE id = ?
        `).run(level, escalationId);
        db.prepare(`
          INSERT INTO work_order_escalation_notifications (escalation_id, role_target, message)
          VALUES (?, ?, ?)
        `).run(escalationId, String(levelRole[level] || "supervisor"), `WO #${Number(r.id)} overdue by ${overdue}h`);
        notifications += 1;
      }
      db.prepare(`UPDATE work_orders SET escalated_at = datetime('now') WHERE id = ?`).run(Number(r.id));
      escalated += 1;
    }
    writeAudit(db, req, {
      module: "workorders",
      action: "escalation_scan",
      entity_type: "work_order",
      entity_id: null,
      payload: { threshold_hours: threshold, escalated_count: escalated },
    });
    return { ok: true, threshold_hours: threshold, escalated_count: escalated, notification_count: notifications };
  });
}
