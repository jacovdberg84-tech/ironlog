// IRONLOG/api/routes/workorders.routes.js
import { db } from "../db/client.js";
import { ensureAuditTable } from "../utils/audit.js";
import { ensureServiceTemplateSchema } from "../utils/serviceTemplates.js";
import { ensureStockCategorySchema, normalizeStockCategory } from "../utils/stockCategory.js";
import { holdsAnyRole } from "../utils/request.js";
import registerBoardRoutes from "./workorders/board.routes.js";
import registerSchedulingRoutes from "./workorders/scheduling.routes.js";
import registerCloseoutRoutes from "./workorders/closeout.routes.js";

export default async function workOrderRoutes(app) {
  ensureStockCategorySchema(db);
  ensureAuditTable(db);
  ensureServiceTemplateSchema(db);
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

  /** Match assigned technician to logged-in user (username or legacy display name). */
  function technicianIdentityKeys(nameOrUsername) {
    const raw = String(nameOrUsername || "").trim().toLowerCase();
    const keys = new Set();
    if (!raw) return keys;
    keys.add(raw);
    if (!hasTable("users")) return keys;
    const byUser = db.prepare(`
      SELECT username, full_name FROM users
      WHERE LOWER(TRIM(username)) = ?
      LIMIT 1
    `).get(raw);
    if (byUser) {
      keys.add(String(byUser.username || "").trim().toLowerCase());
      const fn = String(byUser.full_name || "").trim().toLowerCase();
      if (fn) keys.add(fn);
    }
    const byName = db.prepare(`
      SELECT username, full_name FROM users
      WHERE LOWER(TRIM(COALESCE(full_name, ''))) = ?
      LIMIT 1
    `).get(raw);
    if (byName) {
      keys.add(String(byName.username || "").trim().toLowerCase());
      keys.add(raw);
    }
    return keys;
  }

  function technicianMatchesUser(assignedName, userName) {
    const assignedKeys = technicianIdentityKeys(assignedName);
    const userKeys = technicianIdentityKeys(userName);
    if (!assignedKeys.size || !userKeys.size) return false;
    for (const k of userKeys) {
      if (assignedKeys.has(k)) return true;
    }
    return false;
  }

  /** Prefer login username when supervisor picks a display name. */
  function resolveAssignedUsername(nameOrUsername) {
    const raw = String(nameOrUsername || "").trim();
    if (!raw) return raw;
    if (!hasTable("users")) return raw;
    const byUser = db.prepare(`
      SELECT username FROM users WHERE LOWER(TRIM(username)) = LOWER(?) LIMIT 1
    `).get(raw);
    if (byUser) return String(byUser.username).trim();
    const byName = db.prepare(`
      SELECT username FROM users WHERE LOWER(TRIM(COALESCE(full_name, ''))) = LOWER(?) LIMIT 1
    `).get(raw);
    if (byName) return String(byName.username).trim();
    return raw;
  }

  function getRole(req) {
    return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
  }
  function getSiteCode(req) {
    return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
  }

  function requireRoles(req, reply, roles) {
    const role = getRole(req);
    if (!holdsAnyRole(req, roles)) {
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
  ensureColumn("work_orders", "job_description", "job_description TEXT");
  ensureColumn("work_orders", "artisan_name", "artisan_name TEXT");
  ensureColumn("work_orders", "artisan_signed_at", "artisan_signed_at TEXT");
  ensureColumn("work_orders", "supervisor_name", "supervisor_name TEXT");
  ensureColumn("work_orders", "supervisor_signed_at", "supervisor_signed_at TEXT");
  ensureColumn("work_orders", "completed_at", "completed_at TEXT");
  ensureColumn("work_orders", "assigned_artisan_name", "assigned_artisan_name TEXT");
  ensureColumn("work_orders", "assigned_at", "assigned_at TEXT");
  ensureColumn("work_orders", "assigned_by", "assigned_by TEXT");
  ensureColumn("work_orders", "started_at", "started_at TEXT");
  ensureColumn("work_orders", "shift", "shift TEXT");
  ensureColumn("work_orders", "priority", "priority TEXT");
  ensureColumn("work_orders", "due_date", "due_date TEXT");
  ensureColumn("work_orders", "required_skill", "required_skill TEXT");
  ensureColumn("work_orders", "location_code", "location_code TEXT");
  ensureColumn("work_orders", "escalated_at", "escalated_at TEXT");
  ensureColumn("work_orders", "site_code", "site_code TEXT DEFAULT 'main'");
  ensureColumn("work_orders", "repair_progress", "repair_progress TEXT");
  ensureColumn("work_orders", "repair_progress_at", "repair_progress_at TEXT");
  ensureColumn("work_orders", "labor_hours", "labor_hours REAL DEFAULT 0");
  ensureColumn("work_orders", "labor_rate_per_hour", "labor_rate_per_hour REAL");
  ensureColumn("work_orders", "oil_cost", "oil_cost REAL DEFAULT 0");
  function readLaborRateDefault() {
    try {
      const row = db.prepare(`SELECT value FROM cost_settings WHERE key = 'labor_cost_per_hour_default' LIMIT 1`).get();
      const v = Number(row?.value);
      return Number.isFinite(v) && v > 0 ? v : 35;
    } catch {
      return 35;
    }
  }

  function isOilPartRow(stockCategory) {
    return normalizeStockCategory(stockCategory) === "oil";
  }

  function sumIssuedOilCost(movements) {
    return (Array.isArray(movements) ? movements : []).reduce((sum, m) => {
      if (String(m.movement_type || "").toLowerCase() !== "out") return sum;
      if (!isOilPartRow(m.stock_category)) return sum;
      const qty = Math.abs(Number(m.quantity || 0));
      const unit = Number(m.unit_cost || 0);
      return sum + (qty * (Number.isFinite(unit) ? unit : 0));
    }, 0);
  }

  function enrichWorkOrderCosts(wo, movements = []) {
    const laborHours = Number(wo?.labor_hours || 0);
    const laborRate = Number.isFinite(Number(wo?.labor_rate_per_hour))
      ? Number(wo.labor_rate_per_hour)
      : readLaborRateDefault();
    const manualOil = Number(wo?.oil_cost || 0);
    const issuedOil = sumIssuedOilCost(movements);
    return {
      ...wo,
      labor_rate_per_hour: laborRate,
      labor_cost: Number((laborHours * laborRate).toFixed(2)),
      issued_oil_cost: Number(issuedOil.toFixed(2)),
      total_oil_cost: Number((manualOil + issuedOil).toFixed(2)),
    };
  }
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

  // Route groups live in routes/workorders/. They receive the shared helpers above through ctx.
  const ctx = {
    buildWorkOrderQrProfile,
    canRoleTransition,
    enrichWorkOrderCosts,
    firstExistingColumn,
    getAssetCurrentHours,
    getRole,
    getSiteCode,
    getStoredWoQr,
    hasColumn,
    hasTable,
    inferPriorityForRow,
    normalizePriority,
    normalizeShift,
    readLaborRateDefault,
    requirePermission,
    requireRoles,
    resolveAssignedUsername,
    suggestRuleForWorkOrder,
    technicianMatchesUser,
    upsertWoQr,
  };
  registerBoardRoutes(app, ctx);
  registerSchedulingRoutes(app, ctx);
  registerCloseoutRoutes(app, ctx);
}
