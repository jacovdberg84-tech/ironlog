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
  function getSiteCode(req) {
    return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
  }

  function requireRoles(req, reply, roles) {
    const role = getRole(req);
    if (!roles.includes(role)) {
      reply.code(403).send({ error: `role '${role || "unknown"}' not allowed` });
      return false;
    }
    return true;
  }

  const ROLE_PERMISSION_FALLBACK = {
    admin: ["*"],
    supervisor: ["workorders.close.request", "workorders.reopen", "workorders.delete.request", "workorders.close.approve"],
    plant_manager: ["workorders.close.approve", "workorders.reopen"],
    site_manager: ["workorders.close.approve", "workorders.reopen"],
    quality_manager: ["workorders.close.approve"],
    hr_manager: ["workorders.close.approve"],
    artisan: ["workorders.close.request"],
  };

  function getPermissions(req) {
    const fromHeader = String(req.headers["x-user-permissions"] || "")
      .split(",")
      .map((x) => String(x || "").trim())
      .filter(Boolean);
    if (fromHeader.length) return Array.from(new Set(fromHeader));
    return ROLE_PERMISSION_FALLBACK[getRole(req)] || [];
  }

  function requirePermission(req, reply, permissionKey) {
    const perms = getPermissions(req);
    if (perms.includes("*") || perms.includes(permissionKey)) return true;
    reply.code(403).send({ error: `permission '${permissionKey}' required` });
    return false;
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
  ensureColumn("work_orders", "assigned_artisan_name", "assigned_artisan_name TEXT");
  ensureColumn("work_orders", "shift", "shift TEXT");
  ensureColumn("work_orders", "priority", "priority TEXT");
  ensureColumn("work_orders", "due_date", "due_date TEXT");
  ensureColumn("work_orders", "required_skill", "required_skill TEXT");
  ensureColumn("work_orders", "location_code", "location_code TEXT");
  ensureColumn("work_orders", "escalated_at", "escalated_at TEXT");
  ensureColumn("work_orders", "site_code", "site_code TEXT DEFAULT 'main'");
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
  db.prepare(`
    CREATE TABLE IF NOT EXISTS wo_assignment_rules (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      artisan_name TEXT NOT NULL,
      skill TEXT,
      location_code TEXT,
      shift TEXT,
      max_open_wos INTEGER NOT NULL DEFAULT 8,
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS work_order_escalations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      work_order_id INTEGER NOT NULL,
      threshold_hours INTEGER NOT NULL,
      chain_level INTEGER NOT NULL DEFAULT 1,
      status TEXT NOT NULL DEFAULT 'open',
      detail_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS work_order_escalation_config (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      overdue_hours INTEGER NOT NULL DEFAULT 8,
      level1_role TEXT NOT NULL DEFAULT 'supervisor',
      level2_role TEXT NOT NULL DEFAULT 'manager',
      level3_role TEXT NOT NULL DEFAULT 'admin',
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS work_order_escalation_notifications (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      escalation_id INTEGER NOT NULL,
      role_target TEXT NOT NULL,
      message TEXT NOT NULL,
      sent_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`
    INSERT INTO work_order_escalation_config (id, overdue_hours, level1_role, level2_role, level3_role)
    SELECT 1, 8, 'supervisor', 'manager', 'admin'
    WHERE NOT EXISTS (SELECT 1 FROM work_order_escalation_config WHERE id = 1)
  `).run();

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
        AND DATE(work_date) IS NOT NULL
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

  function normalizeShift(value) {
    const v = String(value || "").trim().toLowerCase();
    if (v === "day" || v === "night") return v;
    return null;
  }

  function normalizePriority(value, fallback = "P3") {
    const v = String(value || fallback).trim().toUpperCase();
    if (["P1", "P2", "P3"].includes(v)) return v;
    return fallback;
  }

  function inferPriorityForRow(row) {
    if (String(row?.priority || "").trim()) return normalizePriority(row.priority);
    const s = String(row?.status || "").toLowerCase();
    const opened = Date.parse(String(row?.opened_at || ""));
    const ageHours = Number.isFinite(opened) ? Math.max(0, Math.floor((Date.now() - opened) / 3600000)) : 0;
    if ((s === "open" || s === "assigned") && ageHours > 72) return "P1";
    if ((s === "open" || s === "assigned") && ageHours > 48) return "P2";
    return "P3";
  }

  function suggestRuleForWorkOrder(wo, rules, workloads) {
    const filtered = (Array.isArray(rules) ? rules : []).filter((r) => {
      if (!Number(r.active || 0)) return false;
      const skillOk = !String(r.skill || "").trim() || String(r.skill).trim().toLowerCase() === String(wo.required_skill || "").trim().toLowerCase();
      const locOk = !String(r.location_code || "").trim() || String(r.location_code).trim().toLowerCase() === String(wo.location_code || "").trim().toLowerCase();
      const shiftOk = !String(r.shift || "").trim() || String(r.shift).trim().toLowerCase() === String(wo.shift || "").trim().toLowerCase();
      if (!skillOk || !locOk || !shiftOk) return false;
      const load = Number(workloads.get(String(r.artisan_name)) || 0);
      const maxOpen = Math.max(1, Number(r.max_open_wos || 8));
      return load < maxOpen;
    });
    filtered.sort((a, b) => {
      const loadA = Number(workloads.get(String(a.artisan_name)) || 0);
      const loadB = Number(workloads.get(String(b.artisan_name)) || 0);
      if (loadA !== loadB) return loadA - loadB;
      return String(a.artisan_name).localeCompare(String(b.artisan_name));
    });
    return filtered[0] || null;
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
    const siteCode = getSiteCode(req);

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
      WHERE LOWER(TRIM(COALESCE(w.site_code, 'main'))) = ?
    `;

    const rows = status
      ? db.prepare(baseSql + ` AND w.status = ? ORDER BY w.id DESC LIMIT 200`).all(siteCode, status)
      : db.prepare(baseSql + ` ORDER BY w.id DESC LIMIT 200`).all(siteCode);

    return rows;
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