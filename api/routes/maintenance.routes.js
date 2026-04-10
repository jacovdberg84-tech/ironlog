// IRONLOG/api/routes/maintenance.routes.js
import { db } from "../db/client.js";
import multipart from "@fastify/multipart";
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import { buildPdfBuffer, sectionTitle, table } from "../utils/pdfGenerator.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function getAssetCurrentHoursInfo(assetId) {
  const fromAssetHours = db.prepare(`
    SELECT total_hours
    FROM asset_hours
    WHERE asset_id = ?
  `).get(assetId);

  const assetHours = fromAssetHours?.total_hours == null ? null : Number(fromAssetHours.total_hours);

  // Prefer latest hourmeter closing reading from Daily Input when it exists.
  // This guards against old/test values in asset_hours (e.g. accidental CSV import).
  const latestMeter = db.prepare(`
    SELECT closing_hours AS latest_closing
    FROM daily_hours
    WHERE asset_id = ?
      AND closing_hours IS NOT NULL
    ORDER BY work_date DESC, id DESC
    LIMIT 1
  `).get(assetId);
  const latestClosing = latestMeter?.latest_closing == null ? null : Number(latestMeter.latest_closing);

  // If asset_hours is present but wildly out of range vs max closing, trust max closing.
  // Heuristic: >5000 hour difference is almost certainly wrong for a live hourmeter.
  if (assetHours != null && latestClosing != null) {
    if (Math.abs(assetHours - latestClosing) > 5000) {
      return { hours: latestClosing, source: "daily_closing" };
    }
    // otherwise take the higher of the two (prevents lagging asset_hours)
    if (latestClosing >= assetHours) return { hours: latestClosing, source: "daily_closing" };
    return { hours: assetHours, source: "asset_hours" };
  }

  if (latestClosing != null) return { hours: latestClosing, source: "daily_closing" };
  if (assetHours != null) return { hours: assetHours, source: "asset_hours" };

  const fromDailyHours = db.prepare(`
    SELECT COALESCE(SUM(hours_run), 0) AS total_hours
    FROM daily_hours
    WHERE asset_id = ?
      AND is_used = 1
      AND hours_run > 0
  `).get(assetId);

  return { hours: Number(fromDailyHours?.total_hours || 0), source: "daily_sum" };
}

function getAssetCurrentHours(assetId) {
  return Number(getAssetCurrentHoursInfo(assetId).hours || 0);
}

function classifyDueStatus(remainingHours, nearDueHours = 50) {
  const remaining = Number(remainingHours || 0);
  const threshold = Math.max(1, Number(nearDueHours || 50));
  if (remaining <= 0) return "OVERDUE";
  if (remaining <= threshold) return "ALMOST DUE";
  return "OK";
}

/** SQL fragment: stock_movements row is an outbound issue (consumption). */
function sqlStockMovementOutbound(alias = "sm") {
  const s = alias;
  return `(LOWER(COALESCE(${s}.movement_type, '')) = 'out' OR COALESCE(${s}.quantity, 0) < 0)`;
}

/**
 * SQL boolean expression (SQLite): parts row is bucketed as oil/lubricant, not hard parts.
 * Uses parts.consumable_kind when set; otherwise part name/code heuristics.
 */
function sqlOilPartPredicate(alias = "p") {
  const p = alias;
  return `(
  LOWER(TRIM(COALESCE(${p}.consumable_kind, ''))) IN ('oil', 'lube', 'lubricant', 'hydraulic', 'hydraulic_oil', 'coolant', 'grease', 'hyd fluid', 'hydraulic fluid')
  OR (
    TRIM(COALESCE(${p}.consumable_kind, '')) = ''
    AND (
      INSTR(LOWER(' ' || REPLACE(REPLACE(COALESCE(${p}.part_name, ''), '-', ' '), '_', ' ') || ' '), ' oil ') > 0
      OR INSTR(LOWER(' ' || REPLACE(REPLACE(COALESCE(${p}.part_name, ''), '-', ' '), '_', ' ') || ' '), ' lube ') > 0
      OR LOWER(TRIM(COALESCE(${p}.part_name, ''))) LIKE 'lubricant%'
      OR LOWER(TRIM(COALESCE(${p}.part_code, ''))) LIKE 'oil%'
      OR LOWER(TRIM(COALESCE(${p}.part_code, ''))) LIKE 'lube%'
    )
  )
)`;
}

export default async function maintenanceRoutes(app) {
  const dataRoot = process.env.IRONLOG_DATA_DIR || process.cwd();
  await app.register(multipart, {
    limits: { fileSize: 20 * 1024 * 1024 },
  });

  db.prepare(`
    CREATE TABLE IF NOT EXISTS manager_inspections (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      inspection_date TEXT NOT NULL,
      inspector_name TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS manager_inspection_photos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      inspection_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      file_path TEXT NOT NULL,
      caption TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (inspection_id) REFERENCES manager_inspections(id) ON DELETE CASCADE
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS manager_damage_reports (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      report_date TEXT NOT NULL,
      inspector_name TEXT,
      hour_meter REAL,
      damage_location TEXT,
      severity TEXT,
      damage_description TEXT,
      immediate_action TEXT,
      out_of_service INTEGER NOT NULL DEFAULT 0,
      damage_time TEXT,
      responsible_person TEXT,
      pending_investigation INTEGER NOT NULL DEFAULT 0,
      hse_report_available INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS maintenance_service_history (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      plan_id INTEGER,
      service_name TEXT NOT NULL,
      service_date TEXT NOT NULL,
      service_hours REAL,
      notes TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT,
      FOREIGN KEY (plan_id) REFERENCES maintenance_plans(id) ON DELETE SET NULL
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS maintenance_histogram_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT DEFAULT 'main',
      event_date TEXT NOT NULL,
      location TEXT,
      part_code TEXT,
      part_name TEXT,
      approval_status TEXT,
      approved_by TEXT,
      notes TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS manager_damage_report_photos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      damage_report_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      file_path TEXT,
      image_data TEXT,
      caption TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (damage_report_id) REFERENCES manager_damage_reports(id) ON DELETE CASCADE
    )
  `).run();

  function hasColumn(table, col) {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name || "") === String(col));
  }
  function ensureColumn(table, colDef, colName) {
    if (!hasColumn(table, colName)) db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
  }
  function pickExistingColumn(table, candidates, fallback) {
    for (const c of candidates) {
      if (hasColumn(table, c)) return c;
    }
    return fallback;
  }
  ensureColumn("manager_inspections", "uuid TEXT", "uuid");
  ensureColumn("manager_inspections", "site_code TEXT DEFAULT 'main'", "site_code");
  ensureColumn("manager_inspections", "updated_at TEXT", "updated_at");
  ensureColumn("manager_inspection_photos", "uuid TEXT", "uuid");
  ensureColumn("manager_inspection_photos", "site_code TEXT DEFAULT 'main'", "site_code");
  ensureColumn("manager_inspection_photos", "updated_at TEXT", "updated_at");
  ensureColumn("manager_inspection_photos", "file_path TEXT", "file_path");
  ensureColumn("manager_inspection_photos", "caption TEXT", "caption");
  ensureColumn("manager_inspection_photos", "created_at TEXT", "created_at");
  ensureColumn("manager_damage_reports", "uuid TEXT", "uuid");
  ensureColumn("manager_damage_reports", "site_code TEXT DEFAULT 'main'", "site_code");
  ensureColumn("manager_damage_reports", "updated_at TEXT", "updated_at");
  ensureColumn("manager_damage_reports", "inspector_name TEXT", "inspector_name");
  ensureColumn("manager_damage_reports", "hour_meter REAL", "hour_meter");
  ensureColumn("manager_damage_reports", "damage_location TEXT", "damage_location");
  ensureColumn("manager_damage_reports", "severity TEXT", "severity");
  ensureColumn("manager_damage_reports", "damage_description TEXT", "damage_description");
  ensureColumn("manager_damage_reports", "immediate_action TEXT", "immediate_action");
  ensureColumn("manager_damage_reports", "out_of_service INTEGER NOT NULL DEFAULT 0", "out_of_service");
  ensureColumn("manager_damage_reports", "damage_time TEXT", "damage_time");
  ensureColumn("manager_damage_reports", "responsible_person TEXT", "responsible_person");
  ensureColumn("manager_damage_reports", "pending_investigation INTEGER NOT NULL DEFAULT 0", "pending_investigation");
  ensureColumn("manager_damage_reports", "hse_report_available INTEGER NOT NULL DEFAULT 0", "hse_report_available");
  ensureColumn("manager_damage_report_photos", "uuid TEXT", "uuid");
  ensureColumn("manager_damage_report_photos", "site_code TEXT DEFAULT 'main'", "site_code");
  ensureColumn("manager_damage_report_photos", "updated_at TEXT", "updated_at");
  ensureColumn("manager_damage_report_photos", "file_path TEXT", "file_path");
  ensureColumn("manager_damage_report_photos", "image_data TEXT", "image_data");
  ensureColumn("manager_damage_report_photos", "caption TEXT", "caption");
  ensureColumn("manager_damage_report_photos", "created_at TEXT", "created_at");
  ensureColumn("maintenance_histogram_events", "site_code TEXT DEFAULT 'main'", "site_code");
  ensureColumn("maintenance_histogram_events", "event_date TEXT", "event_date");
  ensureColumn("maintenance_histogram_events", "location TEXT", "location");
  ensureColumn("maintenance_histogram_events", "part_code TEXT", "part_code");
  ensureColumn("maintenance_histogram_events", "part_name TEXT", "part_name");
  ensureColumn("maintenance_histogram_events", "approval_status TEXT", "approval_status");
  ensureColumn("maintenance_histogram_events", "approved_by TEXT", "approved_by");
  ensureColumn("maintenance_histogram_events", "notes TEXT", "notes");
  ensureColumn("maintenance_histogram_events", "created_by TEXT", "created_by");
  ensureColumn("maintenance_histogram_events", "created_at TEXT", "created_at");
  ensureColumn("maintenance_histogram_events", "updated_at TEXT", "updated_at");
  try {
    const pt = db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name='parts' LIMIT 1`).get();
    if (pt) ensureColumn("parts", "consumable_kind TEXT", "consumable_kind");
  } catch {}

  // Backward compatibility for legacy schema where link column was manager_inspection_id.
  // Keep both readable by normalizing to inspection_id for all new queries/inserts.
  if (!hasColumn("manager_inspection_photos", "inspection_id")) {
    ensureColumn("manager_inspection_photos", "inspection_id INTEGER", "inspection_id");
    if (hasColumn("manager_inspection_photos", "manager_inspection_id")) {
      db.prepare(`
        UPDATE manager_inspection_photos
        SET inspection_id = manager_inspection_id
        WHERE inspection_id IS NULL
          AND manager_inspection_id IS NOT NULL
      `).run();
    }
  }
  const photoInspectionCol = hasColumn("manager_inspection_photos", "inspection_id")
    ? "inspection_id"
    : hasColumn("manager_inspection_photos", "manager_inspection_id")
      ? "manager_inspection_id"
      : "inspection_id";
  const photoPathCol = pickExistingColumn(
    "manager_inspection_photos",
    ["file_path", "photo_path", "path", "image_path", "url"],
    "file_path"
  );
  const photoCaptionCol = pickExistingColumn(
    "manager_inspection_photos",
    ["caption", "note", "notes", "description"],
    "caption"
  );
  const photoCreatedCol = pickExistingColumn(
    "manager_inspection_photos",
    ["created_at", "uploaded_at", "created_on"],
    "created_at"
  );
  const dmgPhotoReportCol = pickExistingColumn(
    "manager_damage_report_photos",
    ["damage_report_id", "manager_damage_report_id", "report_id"],
    "damage_report_id"
  );
  const dmgPhotoPathCol = pickExistingColumn(
    "manager_damage_report_photos",
    ["file_path", "photo_path", "path", "image_path", "url", "image_data"],
    "file_path"
  );
  const dmgPhotoCaptionCol = pickExistingColumn(
    "manager_damage_report_photos",
    ["caption", "note", "notes", "description"],
    "caption"
  );
  const dmgPhotoCreatedCol = pickExistingColumn(
    "manager_damage_report_photos",
    ["created_at", "uploaded_at", "created_on"],
    "created_at"
  );

  // Backfill normalized file_path from common legacy column names.
  if (photoPathCol !== "file_path" && hasColumn("manager_inspection_photos", "file_path")) {
    db.prepare(`
      UPDATE manager_inspection_photos
      SET file_path = ${photoPathCol}
      WHERE (file_path IS NULL OR TRIM(file_path) = '')
        AND ${photoPathCol} IS NOT NULL
    `).run();
  }

  const inspectionsDir = path.join(dataRoot, "uploads", "manager-inspections");
  fs.mkdirSync(inspectionsDir, { recursive: true });
  const damageReportsDir = path.join(dataRoot, "uploads", "manager-damage-reports");
  fs.mkdirSync(damageReportsDir, { recursive: true });

   // =====================================================
  // MAINTENANCE PLANS - LIST
  // GET /api/maintenance/plans
  // =====================================================
  app.get("/plans", async (req, reply) => {
    try {
      const rows = db.prepare(`
        SELECT
          mp.id,
          mp.asset_id,
          mp.service_name,
          mp.interval_hours,
          mp.last_service_hours,
          mp.active,
          a.asset_code,
          a.asset_name,
          a.category
        FROM maintenance_plans mp
        JOIN assets a ON a.id = mp.asset_id
        WHERE a.archived = 0
        ORDER BY a.asset_code ASC, mp.service_name ASC
      `).all();

      const plans = rows.map((r) => {
        const current_hours = getAssetCurrentHours(r.asset_id);
        return {
          ...r,
          current_hours: Number(current_hours.toFixed(2)),
          next_due_hours: Number((Number(r.last_service_hours || 0) + Number(r.interval_hours || 0)).toFixed(2)),
          remaining_hours: Number((Number(r.last_service_hours || 0) + Number(r.interval_hours || 0) - current_hours).toFixed(2)),
        };
      });

      return reply.send({
        ok: true,
        plans
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });
  // =====================================================
  // MAINTENANCE PLANS - CREATE
  // POST /api/maintenance/plans
  // =====================================================
  app.post("/plans", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const service_name = String(req.body?.service_name || "").trim();
      const interval_hours = Number(req.body?.interval_hours || 0);
      const last_service_hours = Number(req.body?.last_service_hours || 0);
      const active = Number(req.body?.active ?? 1) ? 1 : 0;

      if (!asset_id || !service_name || interval_hours <= 0) {
        return reply.code(400).send({
          ok: false,
          error: "asset_id, service_name and interval_hours are required"
        });
      }

      const asset = db.prepare(`
        SELECT id
        FROM assets
        WHERE id = ?
      `).get(asset_id);

      if (!asset) {
        return reply.code(404).send({
          ok: false,
          error: "Asset not found"
        });
      }

      const result = db.prepare(`
        INSERT INTO maintenance_plans (
          asset_id,
          service_name,
          interval_hours,
          last_service_hours,
          active
        )
        VALUES (?, ?, ?, ?, ?)
      `).run(
        asset_id,
        service_name,
        interval_hours,
        last_service_hours,
        active
      );

      return reply.send({
        ok: true,
        id: Number(result.lastInsertRowid)
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - UPDATE
  // PUT /api/maintenance/plans/:id
  // =====================================================
  app.put("/plans/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) {
        return reply.code(400).send({
          ok: false,
          error: "Invalid plan id"
        });
      }

      const existing = db.prepare(`
        SELECT *
        FROM maintenance_plans
        WHERE id = ?
      `).get(id);

      if (!existing) {
        return reply.code(404).send({
          ok: false,
          error: "Maintenance plan not found"
        });
      }

      const asset_id =
        req.body?.asset_id != null ? Number(req.body.asset_id) : Number(existing.asset_id);

      const service_name =
        req.body?.service_name != null
          ? String(req.body.service_name).trim()
          : String(existing.service_name || "").trim();

      const interval_hours =
        req.body?.interval_hours != null
          ? Number(req.body.interval_hours)
          : Number(existing.interval_hours || 0);

      const last_service_hours =
        req.body?.last_service_hours != null
          ? Number(req.body.last_service_hours)
          : Number(existing.last_service_hours || 0);

      const active =
        req.body?.active != null
          ? (Number(req.body.active) ? 1 : 0)
          : Number(existing.active || 0);

      if (!asset_id || !service_name || interval_hours <= 0) {
        return reply.code(400).send({
          ok: false,
          error: "asset_id, service_name and interval_hours are required"
        });
      }

      const asset = db.prepare(`
        SELECT id
        FROM assets
        WHERE id = ?
      `).get(asset_id);

      if (!asset) {
        return reply.code(404).send({
          ok: false,
          error: "Asset not found"
        });
      }

      db.prepare(`
        UPDATE maintenance_plans
        SET
          asset_id = ?,
          service_name = ?,
          interval_hours = ?,
          last_service_hours = ?,
          active = ?
        WHERE id = ?
      `).run(
        asset_id,
        service_name,
        interval_hours,
        last_service_hours,
        active,
        id
      );

      return reply.send({ ok: true });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - DELETE
  // DELETE /api/maintenance/plans/:id
  // =====================================================
  app.delete("/plans/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) {
        return reply.code(400).send({ ok: false, error: "Invalid plan id" });
      }

      const existing = db.prepare(`
        SELECT id
        FROM maintenance_plans
        WHERE id = ?
      `).get(id);

      if (!existing) {
        return reply.code(404).send({ ok: false, error: "Maintenance plan not found" });
      }

      db.prepare(`
        DELETE FROM maintenance_plans WHERE id = ?
      `).run(id);

      return reply.send({ ok: true });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - TOGGLE ACTIVE
  // PATCH /api/maintenance/plans/:id/toggle
  // =====================================================
  app.patch("/plans/:id/toggle", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) {
        return reply.code(400).send({ ok: false, error: "Invalid plan id" });
      }

      const plan = db.prepare(`
        SELECT id, active
        FROM maintenance_plans
        WHERE id = ?
      `).get(id);

      if (!plan) {
        return reply.code(404).send({ ok: false, error: "Maintenance plan not found" });
      }

      const newActive = plan.active ? 0 : 1;
      db.prepare(`
        UPDATE maintenance_plans
        SET active = ?
        WHERE id = ?
      `).run(newActive, id);

      return reply.send({ ok: true, id, active: newActive });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

    // =====================================================
  // GET LIVE HOURS FOR ONE ASSET
  // GET /api/maintenance/asset/:id/live-hours
  // =====================================================
  app.get("/asset/:id/live-hours", async (req, reply) => {
    try {
      const assetId = Number(req.params?.id || 0);
      if (!assetId) {
        return reply.code(400).send({
          ok: false,
          error: "Invalid asset id"
        });
      }

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name
        FROM assets
        WHERE id = ?
      `).get(assetId);

      if (!asset) {
        return reply.code(404).send({
          ok: false,
          error: "Asset not found"
        });
      }

      const currentInfo = getAssetCurrentHoursInfo(assetId);
      const current_hours = Number(currentInfo.hours || 0);

      return reply.send({
        ok: true,
        asset_id: asset.id,
        asset_code: asset.asset_code,
        asset_name: asset.asset_name,
        current_hours: Number(current_hours.toFixed(1)),
        current_hours_source: currentInfo.source
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // LIST MAINTENANCE DUE
  // GET /api/maintenance/due?date=2026-02-27&near_due_hours=50
  // =====================================================
  app.get("/due", async (req, reply) => {
    try {
      const date = String(req.query?.date || "").trim();
      const nearDueHours = Math.max(1, Number(req.query?.near_due_hours || 50));
      if (date && !isDate(date)) {
        return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
      }

      const rows = date
        ? db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              mp.active,
              a.asset_code,
              a.asset_name,
              a.category,
              IFNULL((
                SELECT SUM(dh.hours_run)
                FROM daily_hours dh
                WHERE dh.asset_id = a.id
                  AND dh.is_used = 1
                  AND dh.hours_run > 0
                  AND dh.work_date <= ?
              ), 0) AS current_hours
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.is_standby = 0
              AND a.archived = 0
            ORDER BY a.asset_code, mp.id
          `).all(date)
        : db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              mp.active,
              a.asset_code,
              a.asset_name,
              a.category,
              IFNULL((
                SELECT SUM(dh.hours_run)
                FROM daily_hours dh
                WHERE dh.asset_id = a.id
                  AND dh.is_used = 1
                  AND dh.hours_run > 0
              ), 0) AS current_hours
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.is_standby = 0
              AND a.archived = 0
            ORDER BY a.asset_code, mp.id
          `).all();

      const due = rows.map((r) => {
  const current = getAssetCurrentHours(r.asset_id);
  const next_due = Number(r.last_service_hours || 0) + Number(r.interval_hours || 0);
  const remaining = next_due - current;
  const status = classifyDueStatus(remaining, nearDueHours);

  return {
    plan_id: r.plan_id,
    asset_id: r.asset_id,
    asset_code: r.asset_code,
    asset_name: r.asset_name,
    category: r.category,
    service_name: r.service_name,
    interval_hours: Number(r.interval_hours || 0),
    last_service_hours: Number(r.last_service_hours || 0),
    current_hours: Number(current.toFixed(2)),
    next_due_hours: Number(next_due.toFixed(2)),
    remaining_hours: Number(remaining.toFixed(2)),
    is_overdue: remaining <= 0,
    status
  };
});

      return reply.send({
        ok: true,
        as_of: date || null,
        near_due_hours: nearDueHours,
        due
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // MAINTENANCE HISTORY (single-line per asset+service)
  // GET /api/maintenance/history?as_of=YYYY-MM-DD&days=14
  // =====================================================
  app.get("/history", async (req, reply) => {
    try {
      const as_of = String(req.query?.as_of || "").trim();
      const days = Math.max(3, Math.min(120, Number(req.query?.days || 14)));
      if (as_of && !isDate(as_of)) return reply.code(400).send({ error: "as_of must be YYYY-MM-DD" });

      const endDate = as_of || new Date().toISOString().slice(0, 10);
      const startD = new Date(endDate + "T00:00:00");
      startD.setDate(startD.getDate() - (days - 1));
      const startDate = startD.toISOString().slice(0, 10);

      const plans = db.prepare(`
        SELECT
          mp.id AS plan_id,
          mp.asset_id,
          mp.service_name,
          mp.interval_hours,
          mp.last_service_hours,
          mp.active,
          a.asset_code,
          a.asset_name,
          a.category
        FROM maintenance_plans mp
        JOIN assets a ON a.id = mp.asset_id
        WHERE mp.active = 1
          AND a.active = 1
          AND a.is_standby = 0
          AND a.archived = 0
        ORDER BY a.asset_code ASC, mp.service_name ASC
      `).all();

      const getLastServiced = db.prepare(`
        SELECT
          DATE(COALESCE(w.closed_at, w.completed_at)) AS last_serviced_date,
          COALESCE(w.closed_at, w.completed_at) AS last_serviced_at
        FROM work_orders w
        WHERE w.source = 'service'
          AND w.reference_id = ?
          AND w.status IN ('completed', 'approved', 'closed')
        ORDER BY COALESCE(w.closed_at, w.completed_at) DESC
        LIMIT 1
      `);
      const getLastBackfillServiced = db.prepare(`
        SELECT
          DATE(h.service_date) AS last_serviced_date,
          h.service_date AS last_serviced_at
        FROM maintenance_service_history h
        WHERE h.asset_id = ?
          AND UPPER(TRIM(h.service_name)) = UPPER(TRIM(?))
        ORDER BY h.service_date DESC, h.id DESC
        LIMIT 1
      `);

      const getAvgDaily = db.prepare(`
        SELECT
          COALESCE(SUM(hours_run), 0) AS total_run,
          COUNT(DISTINCT work_date) AS day_count
        FROM daily_hours
        WHERE asset_id = ?
          AND is_used = 1
          AND hours_run > 0
          AND work_date BETWEEN ? AND ?
      `);

      const addDays = (dateStr, add) => {
        const d = new Date(dateStr + "T00:00:00");
        d.setDate(d.getDate() + Math.round(add));
        return d.toISOString().slice(0, 10);
      };

      const rows = plans.map((p) => {
        const currentInfo = getAssetCurrentHoursInfo(Number(p.asset_id || 0));
        const current = Number(currentInfo.hours || 0);
        const next_due = Number(p.last_service_hours || 0) + Number(p.interval_hours || 0);
        const remaining = next_due - current;

        const lastWo = getLastServiced.get(Number(p.plan_id || 0));
        const lastBackfill = getLastBackfillServiced.get(Number(p.asset_id || 0), String(p.service_name || ""));
        const lastWoAt = String(lastWo?.last_serviced_at || "");
        const lastBackfillAt = String(lastBackfill?.last_serviced_at || "");
        const last = lastBackfillAt && (!lastWoAt || lastBackfillAt > lastWoAt) ? lastBackfill : lastWo;
        const avgRow = getAvgDaily.get(Number(p.asset_id || 0), startDate, endDate);
        const totalRun = Number(avgRow?.total_run || 0);
        const dayCount = Number(avgRow?.day_count || 0);
        const avgDaily = dayCount > 0 ? totalRun / dayCount : 0;

        const estDays = avgDaily > 0 ? Math.max(0, remaining / avgDaily) : null;
        const estDate = estDays == null ? null : addDays(endDate, estDays);

        return {
          plan_id: Number(p.plan_id || 0),
          asset_id: Number(p.asset_id || 0),
          asset_code: p.asset_code,
          asset_name: p.asset_name,
          service_name: p.service_name,
          last_serviced_date: last?.last_serviced_date || null,
          current_hours: Number(current.toFixed(2)),
          current_hours_source: currentInfo.source,
          remaining_hours: Number(remaining.toFixed(2)),
          avg_daily_hours: Number(avgDaily.toFixed(2)),
          estimated_service_date: estDate,
        };
      });

      return reply.send({ ok: true, as_of: endDate, range: { start: startDate, end: endDate }, rows });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // =====================================================
  // BACKFILL (ANCIENT) SERVICE HISTORY
  // POST /api/maintenance/history/backfill
  // Body: { asset_id, service_name, service_date, service_hours?, notes?, update_plan_last_hours?, plan_id? }
  // =====================================================
  app.post("/history/backfill", async (req, reply) => {
    try {
      const body = req.body || {};
      const assetId = Number(body.asset_id || 0);
      const serviceName = String(body.service_name || "").trim();
      const serviceDate = String(body.service_date || "").trim();
      const serviceHoursIn = body.service_hours;
      const notes = String(body.notes || "").trim() || null;
      const updatePlanLastHours = Number(body.update_plan_last_hours || 0) === 1;
      const planIdIn = Number(body.plan_id || 0);

      if (!assetId) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!serviceName) return reply.code(400).send({ ok: false, error: "service_name is required" });
      if (!isDate(serviceDate)) return reply.code(400).send({ ok: false, error: "service_date must be YYYY-MM-DD" });

      const serviceHours = serviceHoursIn == null || String(serviceHoursIn).trim() === ""
        ? null
        : Number(serviceHoursIn);
      if (serviceHours != null && (!Number.isFinite(serviceHours) || serviceHours < 0)) {
        return reply.code(400).send({ ok: false, error: "service_hours must be a valid number >= 0" });
      }

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name
        FROM assets
        WHERE id = ?
        LIMIT 1
      `).get(assetId);
      if (!asset) return reply.code(404).send({ ok: false, error: "asset not found" });

      let planId = planIdIn > 0 ? planIdIn : null;
      if (!planId) {
        const matchedPlan = db.prepare(`
          SELECT id
          FROM maintenance_plans
          WHERE asset_id = ?
            AND UPPER(TRIM(service_name)) = UPPER(TRIM(?))
          ORDER BY active DESC, id DESC
          LIMIT 1
        `).get(assetId, serviceName);
        if (matchedPlan?.id) planId = Number(matchedPlan.id);
      }

      const insert = db.prepare(`
        INSERT INTO maintenance_service_history (
          asset_id, plan_id, service_name, service_date, service_hours, notes, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
      `);
      const updatePlan = db.prepare(`
        UPDATE maintenance_plans
        SET last_service_hours = ?
        WHERE id = ?
      `);

      const tx = db.transaction(() => {
        const r = insert.run(
          assetId,
          planId,
          serviceName,
          serviceDate,
          serviceHours,
          notes,
          String(req.headers?.["x-user-name"] || "system")
        );
        if (updatePlanLastHours && planId && serviceHours != null) {
          updatePlan.run(Number(serviceHours), Number(planId));
        }
        return Number(r.lastInsertRowid || 0);
      });

      const id = tx();
      return reply.send({
        ok: true,
        id,
        asset_id: assetId,
        asset_code: asset.asset_code,
        service_name: serviceName,
        service_date: serviceDate,
        service_hours: serviceHours,
        plan_id: planId || null,
        plan_last_hours_updated: Boolean(updatePlanLastHours && planId && serviceHours != null),
      });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // GET /api/maintenance/history/backfill?asset_id=&limit=20
  app.get("/history/backfill", async (req, reply) => {
    try {
      const assetId = Number(req.query?.asset_id || 0);
      const limit = Math.max(1, Math.min(200, Number(req.query?.limit || 20)));
      const where = assetId > 0 ? "WHERE h.asset_id = ?" : "";
      const rows = db.prepare(`
        SELECT
          h.id,
          h.asset_id,
          h.plan_id,
          h.service_name,
          h.service_date,
          h.service_hours,
          h.notes,
          h.created_by,
          h.created_at,
          a.asset_code,
          a.asset_name
        FROM maintenance_service_history h
        JOIN assets a ON a.id = h.asset_id
        ${where}
        ORDER BY h.service_date DESC, h.id DESC
        LIMIT ?
      `).all(...(assetId > 0 ? [assetId, limit] : [limit]));
      return reply.send({ ok: true, rows });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // PUT /api/maintenance/history/backfill/:id
  app.put("/history/backfill/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
      const body = req.body || {};
      const serviceName = String(body.service_name || "").trim();
      const serviceDate = String(body.service_date || "").trim();
      const notes = String(body.notes || "").trim() || null;
      const serviceHoursIn = body.service_hours;
      const serviceHours = serviceHoursIn == null || String(serviceHoursIn).trim() === ""
        ? null
        : Number(serviceHoursIn);
      if (!serviceName) return reply.code(400).send({ ok: false, error: "service_name is required" });
      if (!isDate(serviceDate)) return reply.code(400).send({ ok: false, error: "service_date must be YYYY-MM-DD" });
      if (serviceHours != null && (!Number.isFinite(serviceHours) || serviceHours < 0)) {
        return reply.code(400).send({ ok: false, error: "service_hours must be a valid number >= 0" });
      }
      const cur = db.prepare(`SELECT id FROM maintenance_service_history WHERE id = ?`).get(id);
      if (!cur) return reply.code(404).send({ ok: false, error: "backfill entry not found" });
      db.prepare(`
        UPDATE maintenance_service_history
        SET service_name = ?, service_date = ?, service_hours = ?, notes = ?
        WHERE id = ?
      `).run(serviceName, serviceDate, serviceHours, notes, id);
      return reply.send({ ok: true, id });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // DELETE /api/maintenance/history/backfill/:id
  app.delete("/history/backfill/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
      const cur = db.prepare(`SELECT id FROM maintenance_service_history WHERE id = ?`).get(id);
      if (!cur) return reply.code(404).send({ ok: false, error: "backfill entry not found" });
      db.prepare(`DELETE FROM maintenance_service_history WHERE id = ?`).run(id);
      return reply.send({ ok: true, id });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // =====================================================
  // AUTO-GENERATE SERVICE WORK ORDERS
  // POST /api/maintenance/generate?date=2026-02-27
  // =====================================================
  app.post("/generate", async (req, reply) => {
    try {
      const date = String(req.query?.date || "").trim();
      if (date && !isDate(date)) {
        return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
      }

      const plans = date
        ? db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              IFNULL((
                SELECT SUM(dh.hours_run)
                FROM daily_hours dh
                WHERE dh.asset_id = mp.asset_id
                  AND dh.is_used = 1
                  AND dh.hours_run > 0
                  AND dh.work_date <= ?
              ), 0) AS current_hours
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.is_standby = 0
              AND a.archived = 0
          `).all(date)
        : db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              IFNULL((
                SELECT SUM(dh.hours_run)
                FROM daily_hours dh
                WHERE dh.asset_id = mp.asset_id
                  AND dh.is_used = 1
                  AND dh.hours_run > 0
              ), 0) AS current_hours
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.is_standby = 0
              AND a.archived = 0
          `).all();

      const hasOpenServiceWO = db.prepare(`
        SELECT 1
        FROM work_orders
        WHERE source = 'service'
          AND reference_id = ?
          AND status != 'closed'
        LIMIT 1
      `);

      const insertWO = db.prepare(`
        INSERT INTO work_orders (asset_id, source, reference_id, status)
        VALUES (?, 'service', ?, 'open')
      `);

      const tx = db.transaction(() => {
        const created = [];

        for (const p of plans) {
          const current = getAssetCurrentHours(p.asset_id);
          const next_due = Number(p.last_service_hours || 0) + Number(p.interval_hours || 0);
          const overdue = (next_due - current) <= 0;

          if (!overdue) continue;
          if (hasOpenServiceWO.get(p.plan_id)) continue;

          const wo = insertWO.run(p.asset_id, p.plan_id);
          created.push({
            work_order_id: Number(wo.lastInsertRowid),
            plan_id: p.plan_id,
            asset_id: p.asset_id,
            service_name: p.service_name,
            current_hours: Number(current.toFixed(2)),
            next_due_hours: Number(next_due.toFixed(2))
          });
        }

        return created;
      });

      const created = tx();

      return reply.send({
        ok: true,
        created_count: created.length,
        created
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // UPCOMING SERVICES PDF
  // GET /api/maintenance/due-upcoming.pdf?date=YYYY-MM-DD&near_due_hours=50&download=1
  // =====================================================
  app.get("/due-upcoming.pdf", async (req, reply) => {
    try {
      const date = String(req.query?.date || "").trim();
      if (date && !isDate(date)) {
        return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
      }
      const nearDueHours = Math.max(1, Number(req.query?.near_due_hours || 50));

      const rows = date
        ? db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              a.asset_code,
              a.asset_name
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.is_standby = 0
              AND a.archived = 0
            ORDER BY a.asset_code ASC, mp.service_name ASC
          `).all()
        : db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              a.asset_code,
              a.asset_name
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.is_standby = 0
              AND a.archived = 0
            ORDER BY a.asset_code ASC, mp.service_name ASC
          `).all();

      const dueRows = rows
        .map((r) => {
          const current = getAssetCurrentHours(r.asset_id);
          const nextDue = Number(r.last_service_hours || 0) + Number(r.interval_hours || 0);
          const remaining = nextDue - current;
          const status = classifyDueStatus(remaining, nearDueHours);
          return {
            asset_code: r.asset_code,
            asset_name: r.asset_name,
            service_name: r.service_name,
            current_hours: Number(current.toFixed(2)),
            next_due_hours: Number(nextDue.toFixed(2)),
            remaining_hours: Number(remaining.toFixed(2)),
            status,
          };
        })
        .filter((r) => r.status === "OVERDUE" || r.status === "ALMOST DUE")
        .sort((a, b) => Number(a.remaining_hours || 0) - Number(b.remaining_hours || 0));

      const asOfLabel = date || new Date().toISOString().slice(0, 10);
      const pdf = await buildPdfBuffer(
        (doc) => {
          sectionTitle(doc, "Upcoming Services");
          doc
            .font("Helvetica")
            .fontSize(10)
            .text(`As of: ${asOfLabel} | Threshold: <= ${nearDueHours.toFixed(0)}h flagged as ALMOST DUE`);
          doc.moveDown(0.4);

          table(
            doc,
            [
              { key: "asset_code", label: "Asset", width: 0.12 },
              { key: "asset_name", label: "Name", width: 0.2 },
              { key: "service_name", label: "Service", width: 0.18 },
              { key: "current_hours", label: "Current", width: 0.12, align: "right" },
              { key: "next_due_hours", label: "Next Due", width: 0.12, align: "right" },
              { key: "remaining_hours", label: "Remaining", width: 0.12, align: "right" },
              { key: "status", label: "Status", width: 0.14 },
            ],
            dueRows.length
              ? dueRows.map((r) => ({
                  ...r,
                  current_hours: Number(r.current_hours || 0).toFixed(1),
                  next_due_hours: Number(r.next_due_hours || 0).toFixed(1),
                  remaining_hours: Number(r.remaining_hours || 0).toFixed(1),
                }))
              : [
                  {
                    asset_code: "-",
                    asset_name: "No upcoming services within threshold",
                    service_name: "-",
                    current_hours: "-",
                    next_due_hours: "-",
                    remaining_hours: "-",
                    status: "-",
                  },
                ]
          );
        },
        {
          title: "IRONLOG",
          subtitle: "Maintenance Upcoming Services",
          rightText: `As of: ${asOfLabel}`,
          layout: "landscape",
        }
      );

      const isDownload = String(req.query?.download || "").trim() === "1";
      reply.header("Content-Type", "application/pdf");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="AML_Upcoming_Services_${asOfLabel}.pdf"`
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  async function buildWeeklyForumSummary(query = {}) {
      const nearDueHours = Math.max(1, Number(query?.near_due_hours || 50));
      const startIn = String(query?.start || "").trim();
      const endIn = String(query?.end || "").trim();

      const now = new Date();
      const day = now.getDay(); // 0=Sun ... 6=Sat
      const mondayOffset = day === 0 ? -6 : 1 - day;
      const monday = new Date(now);
      monday.setHours(0, 0, 0, 0);
      monday.setDate(monday.getDate() + mondayOffset);
      const friday = new Date(monday);
      friday.setDate(friday.getDate() + 4);
      const ymd = (d) => d.toISOString().slice(0, 10);
      const start = startIn && isDate(startIn) ? startIn : ymd(monday);
      const end = endIn && isDate(endIn) ? endIn : ymd(friday);
      if (!isDate(start) || !isDate(end)) {
        const e = new Error("start and end must be YYYY-MM-DD");
        e.statusCode = 400;
        throw e;
      }
      if (start > end) {
        const e = new Error("start must be <= end");
        e.statusCode = 400;
        throw e;
      }

      const hasTable = (name) =>
        Boolean(
          db.prepare(`
            SELECT 1
            FROM sqlite_master
            WHERE type = 'table' AND name = ?
            LIMIT 1
          `).get(name)
        );
      const hasColumn = (table, col) => {
        if (!hasTable(table)) return false;
        const rows = db.prepare(`PRAGMA table_info(${table})`).all();
        return rows.some((r) => String(r.name || "") === col);
      };

      const plans = hasTable("maintenance_plans")
        ? db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              a.asset_code,
              a.asset_name
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.archived = 0
              AND a.is_standby = 0
            ORDER BY a.asset_code ASC, mp.service_name ASC
          `).all()
        : [];

      const openWOs =
        hasTable("work_orders")
          ? Number(
              db.prepare(`
                SELECT COUNT(*) AS c
                FROM work_orders
                WHERE LOWER(COALESCE(status, 'open')) NOT IN ('closed', 'completed', 'approved')
              `).get()?.c || 0
            )
          : 0;

      const closedStatuses = "'closed','completed','approved'";
      const hasWOCompletedAt = hasColumn("work_orders", "completed_at");
      const woCloseExpr = hasWOCompletedAt ? "COALESCE(w.completed_at, w.closed_at)" : "w.closed_at";
      const smOutSql = sqlStockMovementOutbound("sm");
      const oilPartSql = sqlOilPartPredicate("p");

      const partsCost =
        hasTable("stock_movements") && hasTable("work_orders")
          ? Number(
              hasTable("parts")
                ? db.prepare(`
                    SELECT COALESCE(SUM(ABS(COALESCE(sm.quantity, 0)) * COALESCE(sm.unit_cost_usd, sm.cost_input, 0)), 0) AS v
                    FROM stock_movements sm
                    JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
                    JOIN parts p ON p.id = sm.part_id
                    WHERE ${woCloseExpr} IS NOT NULL
                      AND DATE(${woCloseExpr}) BETWEEN ? AND ?
                      AND (${smOutSql})
                      AND NOT (${oilPartSql})
                  `).get(start, end)?.v || 0
                : db.prepare(`
                    SELECT COALESCE(SUM(ABS(COALESCE(sm.quantity, 0)) * COALESCE(sm.unit_cost_usd, sm.cost_input, 0)), 0) AS v
                    FROM stock_movements sm
                    JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
                    WHERE ${woCloseExpr} IS NOT NULL
                      AND DATE(${woCloseExpr}) BETWEEN ? AND ?
                      AND (${smOutSql})
                  `).get(start, end)?.v || 0
            )
          : 0;

      const oilCostFromLogs = hasTable("oil_logs")
        ? Number(
            db.prepare(`
              SELECT COALESCE(SUM(COALESCE(quantity, 0) * COALESCE(unit_cost, 0)), 0) AS v
              FROM oil_logs
              WHERE log_date BETWEEN ? AND ?
            `).get(start, end)?.v || 0
          )
        : 0;

      const oilCostFromWoStock =
        hasTable("stock_movements") && hasTable("work_orders") && hasTable("parts")
          ? Number(
              db.prepare(`
                SELECT COALESCE(SUM(ABS(COALESCE(sm.quantity, 0)) * COALESCE(sm.unit_cost_usd, sm.cost_input, 0)), 0) AS v
                FROM stock_movements sm
                JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
                JOIN parts p ON p.id = sm.part_id
                WHERE ${woCloseExpr} IS NOT NULL
                  AND DATE(${woCloseExpr}) BETWEEN ? AND ?
                  AND (${smOutSql})
                  AND (${oilPartSql})
              `).get(start, end)?.v || 0
            )
          : 0;

      const oilCost = oilCostFromLogs + oilCostFromWoStock;

      const laborCost = hasTable("work_orders")
        ? Number(
            db.prepare(`
              SELECT COALESCE(SUM(COALESCE(labor_hours, 0) * COALESCE(labor_rate_per_hour, 0)), 0) AS v
              FROM work_orders w
              WHERE ${woCloseExpr} IS NOT NULL
                AND DATE(${woCloseExpr}) BETWEEN ? AND ?
            `).get(start, end)?.v || 0
          )
        : 0;

      const getAssetCurrentHoursSafe = (assetId) => {
        try {
          return getAssetCurrentHours(assetId);
        } catch {
          return 0;
        }
      };
      const forecastInputs = hasTable("weekly_forum_service_inputs")
        ? db.prepare(`
            SELECT plan_id, oil_part_code, oil_qty, parts_part_code, parts_qty, items_json, notes
            FROM weekly_forum_service_inputs
          `).all()
        : [];
      const inputByPlan = new Map((forecastInputs || []).map((r) => [Number(r.plan_id || 0), r]));
      const getPartPricing = (partCodeIn) => {
        const partCode = String(partCodeIn || "").trim();
        if (!partCode || !hasTable("parts") || !hasTable("stock_movements")) {
          return { unit_cost: 0, on_hand: 0, part_name: null };
        }
        const part = db.prepare(`
          SELECT id, part_name
          FROM parts
          WHERE UPPER(TRIM(part_code)) = UPPER(TRIM(?))
          LIMIT 1
        `).get(partCode);
        if (!part?.id) return { unit_cost: 0, on_hand: 0, part_name: null };
        const costRow = db.prepare(`
          SELECT COALESCE(unit_cost_usd, cost_input, 0) AS unit_cost
          FROM stock_movements
          WHERE part_id = ?
            AND COALESCE(unit_cost_usd, cost_input, 0) > 0
          ORDER BY id DESC
          LIMIT 1
        `).get(part.id);
        const onHandRow = db.prepare(`
          SELECT COALESCE(SUM(quantity), 0) AS on_hand
          FROM stock_movements
          WHERE part_id = ?
        `).get(part.id);
        return {
          unit_cost: Number(costRow?.unit_cost || 0),
          on_hand: Number(onHandRow?.on_hand || 0),
          part_name: String(part.part_name || ""),
        };
      };

      const forecastRows = plans
        .map((p) => {
          const current = Number(getAssetCurrentHoursSafe(p.asset_id) || 0);
          const nextDue = Number(p.last_service_hours || 0) + Number(p.interval_hours || 0);
          const remaining = nextDue - current;
          const status = classifyDueStatus(remaining, nearDueHours);

          const hist =
            hasTable("work_orders") && hasTable("stock_movements")
              ? hasTable("parts")
                ? db.prepare(`
                    SELECT
                      COUNT(DISTINCT w.id) AS service_events,
                      COALESCE(SUM(CASE WHEN sm.id IS NOT NULL AND (${smOutSql}) AND NOT (${oilPartSql})
                        THEN ABS(COALESCE(sm.quantity, 0)) ELSE 0 END), 0) AS parts_qty_total,
                      COALESCE(SUM(CASE WHEN sm.id IS NOT NULL AND (${smOutSql}) AND NOT (${oilPartSql})
                        THEN ABS(COALESCE(sm.quantity, 0)) * COALESCE(sm.unit_cost_usd, sm.cost_input, 0) ELSE 0 END), 0) AS parts_cost_total,
                      COALESCE(SUM(CASE WHEN sm.id IS NOT NULL AND (${smOutSql}) AND (${oilPartSql})
                        THEN ABS(COALESCE(sm.quantity, 0)) ELSE 0 END), 0) AS oil_qty_sm_total,
                      COALESCE(SUM(CASE WHEN sm.id IS NOT NULL AND (${smOutSql}) AND (${oilPartSql})
                        THEN ABS(COALESCE(sm.quantity, 0)) * COALESCE(sm.unit_cost_usd, sm.cost_input, 0) ELSE 0 END), 0) AS oil_cost_sm_total
                    FROM work_orders w
                    LEFT JOIN stock_movements sm ON sm.reference = ('work_order:' || w.id)
                    LEFT JOIN parts p ON p.id = sm.part_id
                    WHERE LOWER(COALESCE(w.source, '')) = 'service'
                      AND COALESCE(w.reference_id, 0) = ?
                      AND LOWER(COALESCE(w.status, '')) IN (${closedStatuses})
                  `).get(Number(p.plan_id || 0))
                : db.prepare(`
                    SELECT
                      COUNT(DISTINCT w.id) AS service_events,
                      COALESCE(SUM(CASE WHEN sm.id IS NOT NULL AND (${smOutSql})
                        THEN ABS(COALESCE(sm.quantity, 0)) ELSE 0 END), 0) AS parts_qty_total,
                      COALESCE(SUM(CASE WHEN sm.id IS NOT NULL AND (${smOutSql})
                        THEN ABS(COALESCE(sm.quantity, 0)) * COALESCE(sm.unit_cost_usd, sm.cost_input, 0) ELSE 0 END), 0) AS parts_cost_total,
                      0 AS oil_qty_sm_total,
                      0 AS oil_cost_sm_total
                    FROM work_orders w
                    LEFT JOIN stock_movements sm ON sm.reference = ('work_order:' || w.id)
                    WHERE LOWER(COALESCE(w.source, '')) = 'service'
                      AND COALESCE(w.reference_id, 0) = ?
                      AND LOWER(COALESCE(w.status, '')) IN (${closedStatuses})
                  `).get(Number(p.plan_id || 0))
              : null;

          const serviceEvents = Number(hist?.service_events || 0);
          const avgPartsQty = serviceEvents > 0 ? Number(hist.parts_qty_total || 0) / serviceEvents : 0;
          const avgPartsCost = serviceEvents > 0 ? Number(hist.parts_cost_total || 0) / serviceEvents : 0;

          const oilAvg = hasTable("oil_logs") && hasTable("work_orders")
            ? db.prepare(`
                SELECT
                  COALESCE(SUM(ol.quantity), 0) AS oil_qty_total,
                  COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, 0)), 0) AS oil_cost_total
                FROM oil_logs ol
                WHERE ol.asset_id = ?
                  AND ol.log_date IN (
                    SELECT DATE(${woCloseExpr})
                    FROM work_orders w
                    WHERE LOWER(COALESCE(w.source, '')) = 'service'
                      AND COALESCE(w.reference_id, 0) = ?
                      AND ${woCloseExpr} IS NOT NULL
                      AND LOWER(COALESCE(w.status, '')) IN (${closedStatuses})
                  )
              `).get(Number(p.asset_id || 0), Number(p.plan_id || 0))
            : null;
          const oilQtyLogs = Number(oilAvg?.oil_qty_total || 0);
          const oilCostLogsPlan = Number(oilAvg?.oil_cost_total || 0);
          const oilQtySm = Number(hist?.oil_qty_sm_total || 0);
          const oilCostSm = Number(hist?.oil_cost_sm_total || 0);
          const avgOilQty =
            serviceEvents > 0 ? (oilQtyLogs + oilQtySm) / serviceEvents : 0;
          const avgOilCost =
            serviceEvents > 0 ? (oilCostLogsPlan + oilCostSm) / serviceEvents : 0;

          const serviceKitCost = avgPartsCost + avgOilCost;
          const manual = inputByPlan.get(Number(p.plan_id || 0)) || null;
          let manualItems = [];
          try {
            const parsed = JSON.parse(String(manual?.items_json || "[]"));
            if (Array.isArray(parsed)) manualItems = parsed;
          } catch {}
          if (!manualItems.length) {
            manualItems = [
              { type: "oil", part_code: String(manual?.oil_part_code || "").trim(), qty: Number(manual?.oil_qty || 0) },
              { type: "part", part_code: String(manual?.parts_part_code || "").trim(), qty: Number(manual?.parts_qty || 0) },
            ].filter((x) => x.part_code && Number(x.qty || 0) > 0);
          }
          const pricedItems = manualItems.map((it) => {
            const type = String(it?.type || "part").toLowerCase() === "oil" ? "oil" : "part";
            const part_code = String(it?.part_code || "").trim();
            const qty = Math.max(0, Number(it?.qty || 0));
            const pricing = part_code ? getPartPricing(part_code) : { unit_cost: 0, on_hand: 0, part_name: null };
            return {
              type,
              part_code,
              part_name: pricing.part_name || null,
              qty: Number(qty.toFixed(2)),
              unit_cost: Number(Number(pricing.unit_cost || 0).toFixed(4)),
              on_hand: Number(Number(pricing.on_hand || 0).toFixed(2)),
              line_cost: Number((qty * Number(pricing.unit_cost || 0)).toFixed(2)),
            };
          }).filter((x) => x.part_code && x.qty > 0);
          const manualOilCost = pricedItems.filter((x) => x.type === "oil").reduce((s, x) => s + Number(x.line_cost || 0), 0);
          const manualPartsCost = pricedItems.filter((x) => x.type !== "oil").reduce((s, x) => s + Number(x.line_cost || 0), 0);
          const hasManualOverride = pricedItems.length > 0;
          return {
            plan_id: Number(p.plan_id || 0),
            asset_id: Number(p.asset_id || 0),
            asset_code: p.asset_code,
            asset_name: p.asset_name,
            service_name: p.service_name,
            current_hours: Number(current.toFixed(2)),
            next_due_hours: Number(nextDue.toFixed(2)),
            remaining_hours: Number(remaining.toFixed(2)),
            status,
            forecast: {
              service_events: serviceEvents,
              avg_oil_qty: Number(avgOilQty.toFixed(2)),
              avg_oil_cost: Number(avgOilCost.toFixed(2)),
              avg_parts_qty: Number(avgPartsQty.toFixed(2)),
              avg_parts_cost: Number(avgPartsCost.toFixed(2)),
              est_service_kit_cost: Number((hasManualOverride ? (manualOilCost + manualPartsCost) : serviceKitCost).toFixed(2)),
              cost_source: hasManualOverride ? "manual_store_pricing" : "historical_average",
              manual: {
                oil_cost_total: Number(manualOilCost.toFixed(2)),
                parts_cost_total: Number(manualPartsCost.toFixed(2)),
                items: pricedItems,
                notes: String(manual?.notes || ""),
              },
            },
          };
        })
        .filter((r) => r.status === "OVERDUE" || r.status === "ALMOST DUE")
        .sort((a, b) => Number(a.remaining_hours || 0) - Number(b.remaining_hours || 0))
        .slice(0, 40);

      const totalForecastCost = forecastRows.reduce(
        (s, r) => s + Number(r.forecast?.est_service_kit_cost || 0),
        0
      );

      return {
        ok: true,
        range: { start, end },
        kpis: {
          open_work_orders: openWOs,
          upcoming_services_flagged: forecastRows.length,
        },
        costs: {
          stores_oil_cost: Number(oilCost.toFixed(2)),
          stores_oil_from_logs: Number(oilCostFromLogs.toFixed(2)),
          stores_oil_from_work_orders: Number(oilCostFromWoStock.toFixed(2)),
          stores_parts_cost: Number(partsCost.toFixed(2)),
          maintenance_labor_cost: Number(laborCost.toFixed(2)),
          weekly_total_cost: Number((oilCost + partsCost + laborCost).toFixed(2)),
          upcoming_service_forecast_cost: Number(totalForecastCost.toFixed(2)),
        },
        upcoming_services: forecastRows,
      };
  }

  // =====================================================
  // WEEKLY FORUM SUMMARY (cross-functional alignment)
  // GET /api/maintenance/weekly-forum/summary?start=YYYY-MM-DD&end=YYYY-MM-DD&near_due_hours=50
  // =====================================================
  app.get("/weekly-forum/summary", async (req, reply) => {
    try {
      const data = await buildWeeklyForumSummary(req.query || {});
      return reply.send(data);
    } catch (err) {
      req.log.error(err);
      return reply.code(Number(err?.statusCode || 500)).send({ ok: false, error: err.message || String(err) });
    }
  });

  // =====================================================
  // WEEKLY FORUM PDF
  // GET /api/maintenance/weekly-forum.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&near_due_hours=50&download=1
  // =====================================================
  app.get("/weekly-forum.pdf", async (req, reply) => {
    try {
      const data = await buildWeeklyForumSummary(req.query || {});
      const start = String(data?.range?.start || "");
      const end = String(data?.range?.end || "");
      const isDownload = String(req.query?.download || "").trim() === "1";

      const pdf = await buildPdfBuffer(
        (doc) => {
          sectionTitle(doc, "Weekly Forum Summary");
          table(
            doc,
            [
              { key: "metric", label: "Metric", width: 0.62 },
              { key: "value", label: "Value", width: 0.38, align: "right" },
            ],
            [
              { metric: "Range", value: `${start} to ${end}` },
              { metric: "Open Work Orders", value: Number(data?.kpis?.open_work_orders || 0) },
              { metric: "Upcoming Services Flagged", value: Number(data?.kpis?.upcoming_services_flagged || 0) },
              {
                metric: "Stores parts (excl. oil/lube SKUs)",
                value: Number(data?.costs?.stores_parts_cost || 0).toFixed(2),
              },
              {
                metric: "Oil cost — lube log entries",
                value: Number(data?.costs?.stores_oil_from_logs || 0).toFixed(2),
              },
              {
                metric: "Oil cost — WO stock (oil/lube lines)",
                value: Number(data?.costs?.stores_oil_from_work_orders || 0).toFixed(2),
              },
              {
                metric: "Stores oil total",
                value: Number(data?.costs?.stores_oil_cost || 0).toFixed(2),
              },
              { metric: "Maintenance Labor Cost", value: Number(data?.costs?.maintenance_labor_cost || 0).toFixed(2) },
              { metric: "Weekly Total Cost", value: Number(data?.costs?.weekly_total_cost || 0).toFixed(2) },
              { metric: "Upcoming Service Forecast Cost", value: Number(data?.costs?.upcoming_service_forecast_cost || 0).toFixed(2) },
            ]
          );

          sectionTitle(doc, "Upcoming Services Forecast");
          const rows = Array.isArray(data?.upcoming_services) ? data.upcoming_services : [];
          table(
            doc,
            [
              { key: "machine", label: "Machine", width: 0.18 },
              { key: "service", label: "Service", width: 0.15 },
              { key: "current", label: "Current", width: 0.07, align: "right" },
              { key: "next", label: "Next Due", width: 0.07, align: "right" },
              { key: "remain", label: "Remaining", width: 0.07, align: "right" },
              { key: "status", label: "Status", width: 0.09 },
              { key: "oil", label: "Avg Oil Qty", width: 0.07, align: "right" },
              { key: "oil_cost", label: "Avg Oil $", width: 0.08, align: "right" },
              { key: "parts", label: "Avg Parts Qty", width: 0.07, align: "right" },
              { key: "parts_cost", label: "Avg Parts $", width: 0.08, align: "right" },
              { key: "kit", label: "Est Kit Cost", width: 0.07, align: "right" },
            ],
            rows.length
              ? rows.map((r) => ({
                  machine: `${String(r.asset_code || "-")} - ${String(r.asset_name || "-")}`,
                  service: String(r.service_name || "-"),
                  current: Number(r.current_hours || 0).toFixed(1),
                  next: Number(r.next_due_hours || 0).toFixed(1),
                  remain: Number(r.remaining_hours || 0).toFixed(1),
                  status: String(r.status || "-"),
                  oil: Number(r?.forecast?.avg_oil_qty || 0).toFixed(1),
                  oil_cost: Number(r?.forecast?.avg_oil_cost || 0).toFixed(2),
                  parts: Number(r?.forecast?.avg_parts_qty || 0).toFixed(1),
                  parts_cost: Number(r?.forecast?.avg_parts_cost || 0).toFixed(2),
                  kit: Number(r?.forecast?.est_service_kit_cost || 0).toFixed(2),
                }))
              : [{
                  machine: "-",
                  service: "No upcoming services within threshold",
                  current: "-",
                  next: "-",
                  remain: "-",
                  status: "-",
                  oil: "-",
                  oil_cost: "-",
                  parts: "-",
                  parts_cost: "-",
                  kit: "-",
                }]
          );
        },
        {
          title: "IRONLOG",
          subtitle: "Weekly Forum",
          rightText: `${start} to ${end}`,
          showPageNumbers: true,
          layout: "landscape",
        }
      );

      reply.header("Content-Type", "application/pdf");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="AML_Weekly_Forum_${end}.pdf"`
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(Number(err?.statusCode || 500)).send({ ok: false, error: err.message || String(err) });
    }
  });

  // =====================================================
  // WEEKLY FORUM ACTION TRACKER
  // =====================================================
  db.prepare(`
    CREATE TABLE IF NOT EXISTS weekly_forum_actions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      action_date TEXT NOT NULL,
      department TEXT NOT NULL,
      action_item TEXT NOT NULL,
      owner_name TEXT NOT NULL,
      due_date TEXT,
      status TEXT NOT NULL DEFAULT 'open',
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS weekly_forum_service_inputs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      plan_id INTEGER NOT NULL UNIQUE,
      oil_part_code TEXT,
      oil_qty REAL NOT NULL DEFAULT 0,
      parts_part_code TEXT,
      parts_qty REAL NOT NULL DEFAULT 0,
      items_json TEXT NOT NULL DEFAULT '[]',
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  try {
    const wfInputCols = db.prepare(`PRAGMA table_info(weekly_forum_service_inputs)`).all();
    const wfInputHasItems = wfInputCols.some((c) => String(c?.name || "") === "items_json");
    if (!wfInputHasItems) {
      db.prepare(`ALTER TABLE weekly_forum_service_inputs ADD COLUMN items_json TEXT NOT NULL DEFAULT '[]'`).run();
    }
  } catch {}

  app.get("/weekly-forum/actions", async (req, reply) => {
    try {
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const status = String(req.query?.status || "").trim().toLowerCase();
      if (start && !isDate(start)) return reply.code(400).send({ ok: false, error: "start must be YYYY-MM-DD" });
      if (end && !isDate(end)) return reply.code(400).send({ ok: false, error: "end must be YYYY-MM-DD" });

      const where = [];
      const params = [];
      if (start) {
        where.push("action_date >= ?");
        params.push(start);
      }
      if (end) {
        where.push("action_date <= ?");
        params.push(end);
      }
      if (status) {
        where.push("LOWER(COALESCE(status,'open')) = ?");
        params.push(status);
      }

      const rows = db.prepare(`
        SELECT
          id, action_date, department, action_item, owner_name, due_date,
          status, notes, created_at, updated_at
        FROM weekly_forum_actions
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY
          CASE LOWER(COALESCE(status, 'open'))
            WHEN 'open' THEN 0
            WHEN 'in_progress' THEN 1
            WHEN 'blocked' THEN 2
            ELSE 3
          END ASC,
          COALESCE(due_date, action_date) ASC,
          id DESC
      `).all(...params);
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
  app.get("/weekly-forum/parts", async (req, reply) => {
    try {
      const hasPartsTable = Boolean(
        db.prepare(`
          SELECT 1
          FROM sqlite_master
          WHERE type = 'table' AND name = 'parts'
          LIMIT 1
        `).get()
      );
      const rows = hasPartsTable
        ? db.prepare(`
            SELECT
              p.part_code,
              p.part_name,
              COALESCE(SUM(sm.quantity), 0) AS on_hand,
              COALESCE((
                SELECT COALESCE(sm2.unit_cost_usd, sm2.cost_input, 0)
                FROM stock_movements sm2
                WHERE sm2.part_id = p.id
                  AND COALESCE(sm2.unit_cost_usd, sm2.cost_input, 0) > 0
                ORDER BY sm2.id DESC
                LIMIT 1
              ), 0) AS latest_unit_cost
            FROM parts p
            LEFT JOIN stock_movements sm ON sm.part_id = p.id
            GROUP BY p.id
            ORDER BY p.part_code ASC
            LIMIT 1500
          `).all()
        : [];
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
  app.get("/weekly-forum/forecast-inputs", async (req, reply) => {
    try {
      const rows = db.prepare(`
        SELECT id, plan_id, oil_part_code, oil_qty, parts_part_code, parts_qty, items_json, notes, updated_at
        FROM weekly_forum_service_inputs
        ORDER BY plan_id ASC
      `).all();
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
  app.post("/weekly-forum/forecast-inputs", async (req, reply) => {
    try {
      const plan_id = Number(req.body?.plan_id || 0);
      const items = Array.isArray(req.body?.items) ? req.body.items : [];
      const oil_part_code = String(req.body?.oil_part_code || "").trim() || null;
      const oil_qty = Math.max(0, Number(req.body?.oil_qty || 0));
      const parts_part_code = String(req.body?.parts_part_code || "").trim() || null;
      const parts_qty = Math.max(0, Number(req.body?.parts_qty || 0));
      const notes = String(req.body?.notes || "").trim() || null;
      const normalizedItems = items
        .map((it) => ({
          type: String(it?.type || "part").toLowerCase() === "oil" ? "oil" : "part",
          part_code: String(it?.part_code || "").trim(),
          qty: Math.max(0, Number(it?.qty || 0)),
        }))
        .filter((it) => it.part_code && it.qty > 0);
      const items_json = JSON.stringify(normalizedItems);
      if (!plan_id) return reply.code(400).send({ ok: false, error: "plan_id is required" });
      db.prepare(`
        INSERT INTO weekly_forum_service_inputs (
          plan_id, oil_part_code, oil_qty, parts_part_code, parts_qty, items_json, notes, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))
        ON CONFLICT(plan_id) DO UPDATE SET
          oil_part_code = excluded.oil_part_code,
          oil_qty = excluded.oil_qty,
          parts_part_code = excluded.parts_part_code,
          parts_qty = excluded.parts_qty,
          items_json = excluded.items_json,
          notes = excluded.notes,
          updated_at = datetime('now')
      `).run(plan_id, oil_part_code, oil_qty, parts_part_code, parts_qty, items_json, notes);
      return reply.send({ ok: true, plan_id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/weekly-forum/actions", async (req, reply) => {
    try {
      const action_date = String(req.body?.action_date || "").trim() || new Date().toISOString().slice(0, 10);
      const department = String(req.body?.department || "").trim();
      const action_item = String(req.body?.action_item || "").trim();
      const owner_name = String(req.body?.owner_name || "").trim();
      const due_date = String(req.body?.due_date || "").trim() || null;
      const status = String(req.body?.status || "open").trim().toLowerCase() || "open";
      const notes = String(req.body?.notes || "").trim() || null;

      if (!isDate(action_date)) return reply.code(400).send({ ok: false, error: "action_date must be YYYY-MM-DD" });
      if (!department) return reply.code(400).send({ ok: false, error: "department is required" });
      if (!action_item) return reply.code(400).send({ ok: false, error: "action_item is required" });
      if (!owner_name) return reply.code(400).send({ ok: false, error: "owner_name is required" });
      if (due_date && !isDate(due_date)) return reply.code(400).send({ ok: false, error: "due_date must be YYYY-MM-DD" });

      const allowedStatuses = ["open", "in_progress", "blocked", "done"];
      if (!allowedStatuses.includes(status)) {
        return reply.code(400).send({ ok: false, error: "status must be open|in_progress|blocked|done" });
      }

      const ins = db.prepare(`
        INSERT INTO weekly_forum_actions (
          action_date, department, action_item, owner_name, due_date, status, notes, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(action_date, department, action_item, owner_name, due_date, status, notes);
      return reply.send({ ok: true, id: Number(ins.lastInsertRowid) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/weekly-forum/actions/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
      const existing = db.prepare(`SELECT id FROM weekly_forum_actions WHERE id = ?`).get(id);
      if (!existing) return reply.code(404).send({ ok: false, error: "action not found" });

      const department = req.body?.department != null ? String(req.body.department).trim() : undefined;
      const action_item = req.body?.action_item != null ? String(req.body.action_item).trim() : undefined;
      const owner_name = req.body?.owner_name != null ? String(req.body.owner_name).trim() : undefined;
      const due_date = req.body?.due_date != null ? (String(req.body.due_date).trim() || null) : undefined;
      const status = req.body?.status != null ? String(req.body.status).trim().toLowerCase() : undefined;
      const notes = req.body?.notes != null ? (String(req.body.notes).trim() || null) : undefined;

      if (due_date !== undefined && due_date && !isDate(due_date)) {
        return reply.code(400).send({ ok: false, error: "due_date must be YYYY-MM-DD" });
      }
      if (status !== undefined) {
        const allowedStatuses = ["open", "in_progress", "blocked", "done"];
        if (!allowedStatuses.includes(status)) {
          return reply.code(400).send({ ok: false, error: "status must be open|in_progress|blocked|done" });
        }
      }

      db.prepare(`
        UPDATE weekly_forum_actions
        SET
          department = COALESCE(?, department),
          action_item = COALESCE(?, action_item),
          owner_name = COALESCE(?, owner_name),
          due_date = COALESCE(?, due_date),
          status = COALESCE(?, status),
          notes = COALESCE(?, notes),
          updated_at = datetime('now')
        WHERE id = ?
      `).run(department ?? null, action_item ?? null, owner_name ?? null, due_date ?? null, status ?? null, notes ?? null, id);

      return reply.send({ ok: true, id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/histogram/events", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const location = String(req.query?.location || "").trim();
      const approval = String(req.query?.approval || "").trim();
      const part = String(req.query?.part || "").trim();
      const limitRaw = Number(req.query?.limit || 300);
      const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(2000, Math.trunc(limitRaw))) : 300;

      const where = ["COALESCE(site_code, 'main') = ?"];
      const params = [site_code];
      if (isDate(start)) {
        where.push("event_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("event_date <= ?");
        params.push(end);
      }
      if (location) {
        where.push("LOWER(COALESCE(location, '')) LIKE ?");
        params.push(`%${location.toLowerCase()}%`);
      }
      if (approval) {
        where.push("LOWER(COALESCE(approval_status, '')) LIKE ?");
        params.push(`%${approval.toLowerCase()}%`);
      }
      if (part) {
        where.push("(LOWER(COALESCE(part_code, '')) LIKE ? OR LOWER(COALESCE(part_name, '')) LIKE ?)");
        params.push(`%${part.toLowerCase()}%`, `%${part.toLowerCase()}%`);
      }

      const sql = `
        SELECT id, event_date, location, part_code, part_name, approval_status, approved_by, notes, created_by, created_at, updated_at
        FROM maintenance_histogram_events
        WHERE ${where.join(" AND ")}
        ORDER BY event_date DESC, id DESC
        LIMIT ${limit}
      `;
      const rows = db.prepare(sql).all(...params);
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/histogram/events", async (req, reply) => {
    try {
      const body = req.body || {};
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const created_by = String(req.headers?.["x-user-name"] || "system").trim() || "system";
      const event_date = String(body.event_date || "").trim();
      if (!isDate(event_date)) {
        return reply.code(400).send({ ok: false, error: "event_date must be YYYY-MM-DD" });
      }
      const location = String(body.location || "").trim();
      const part_code = String(body.part_code || "").trim();
      const part_name = String(body.part_name || "").trim();
      const approval_status = String(body.approval_status || "").trim();
      const approved_by = String(body.approved_by || "").trim();
      const notes = String(body.notes || "").trim();
      const now = new Date().toISOString();
      const info = db.prepare(`
        INSERT INTO maintenance_histogram_events (
          site_code, event_date, location, part_code, part_name, approval_status, approved_by, notes, created_by, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).run(site_code, event_date, location, part_code, part_name, approval_status, approved_by, notes, created_by, now, now);
      return reply.send({ ok: true, id: Number(info.lastInsertRowid || 0) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/histogram/events/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
      const body = req.body || {};
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const existing = db.prepare(`SELECT id FROM maintenance_histogram_events WHERE id = ? AND COALESCE(site_code, 'main') = ?`).get(id, site_code);
      if (!existing) return reply.code(404).send({ ok: false, error: "Event not found" });

      const event_date = String(body.event_date || "").trim();
      if (event_date && !isDate(event_date)) {
        return reply.code(400).send({ ok: false, error: "event_date must be YYYY-MM-DD" });
      }
      const location = String(body.location || "").trim();
      const part_code = String(body.part_code || "").trim();
      const part_name = String(body.part_name || "").trim();
      const approval_status = String(body.approval_status || "").trim();
      const approved_by = String(body.approved_by || "").trim();
      const notes = String(body.notes || "").trim();
      const now = new Date().toISOString();

      db.prepare(`
        UPDATE maintenance_histogram_events
        SET
          event_date = COALESCE(NULLIF(?, ''), event_date),
          location = ?,
          part_code = ?,
          part_name = ?,
          approval_status = ?,
          approved_by = ?,
          notes = ?,
          updated_at = ?
        WHERE id = ?
          AND COALESCE(site_code, 'main') = ?
      `).run(event_date, location, part_code, part_name, approval_status, approved_by, notes, now, id, site_code);
      return reply.send({ ok: true, id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.delete("/histogram/events/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const info = db.prepare(`
        DELETE FROM maintenance_histogram_events
        WHERE id = ?
          AND COALESCE(site_code, 'main') = ?
      `).run(id, site_code);
      if (!Number(info.changes || 0)) return reply.code(404).send({ ok: false, error: "Event not found" });
      return reply.send({ ok: true, id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/histogram/events.pdf", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const location = String(req.query?.location || "").trim();
      const approval = String(req.query?.approval || "").trim();
      const part = String(req.query?.part || "").trim();
      const download = String(req.query?.download || "").trim() === "1";

      const where = ["COALESCE(site_code, 'main') = ?"];
      const params = [site_code];
      if (isDate(start)) {
        where.push("event_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("event_date <= ?");
        params.push(end);
      }
      if (location) {
        where.push("LOWER(COALESCE(location, '')) LIKE ?");
        params.push(`%${location.toLowerCase()}%`);
      }
      if (approval) {
        where.push("LOWER(COALESCE(approval_status, '')) LIKE ?");
        params.push(`%${approval.toLowerCase()}%`);
      }
      if (part) {
        where.push("(LOWER(COALESCE(part_code, '')) LIKE ? OR LOWER(COALESCE(part_name, '')) LIKE ?)");
        params.push(`%${part.toLowerCase()}%`, `%${part.toLowerCase()}%`);
      }

      const rows = db.prepare(`
        SELECT event_date, location, part_code, part_name, approval_status, approved_by, notes, created_by
        FROM maintenance_histogram_events
        WHERE ${where.join(" AND ")}
        ORDER BY event_date DESC, id DESC
        LIMIT 2000
      `).all(...params);

      const periodLabel = `${isDate(start) ? start : "-"} to ${isDate(end) ? end : "-"}`;
      const pdf = await buildPdfBuffer((doc) => {
        sectionTitle(doc, "Maintenance Histogram Events");
        doc
          .font("Helvetica")
          .fontSize(10)
          .text(`Site: ${site_code} | Period: ${periodLabel} | Total events: ${rows.length}`);
        doc.moveDown(0.4);
        table(
          doc,
          [
            { key: "event_date", label: "Date", width: 0.1 },
            { key: "location", label: "Location", width: 0.16 },
            { key: "part_code", label: "Part Code", width: 0.12 },
            { key: "part_name", label: "Part Name", width: 0.14 },
            { key: "approval_status", label: "Approval", width: 0.1 },
            { key: "approved_by", label: "Approved By", width: 0.12 },
            { key: "notes", label: "Notes", width: 0.16 },
            { key: "created_by", label: "Captured By", width: 0.1 },
          ],
          rows.length
            ? rows.map((r) => ({
                event_date: String(r.event_date || "-"),
                location: String(r.location || "-"),
                part_code: String(r.part_code || "-"),
                part_name: String(r.part_name || "-"),
                approval_status: String(r.approval_status || "-"),
                approved_by: String(r.approved_by || "-"),
                notes: String(r.notes || "-"),
                created_by: String(r.created_by || "-"),
              }))
            : [{ event_date: "-", location: "No events found", part_code: "-", part_name: "-", approval_status: "-", approved_by: "-", notes: "-", created_by: "-" }]
        );
      });

      const dateTag = new Date().toISOString().slice(0, 10);
      reply
        .header("Content-Type", "application/pdf")
        .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename="maintenance-histogram-${dateTag}.pdf"`)
        .send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // =====================================================
  // MANAGER INSPECTIONS
  // =====================================================
  app.get("/inspections", async (req, reply) => {
    try {
      const assetId = Number(req.query?.asset_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const params = [];
      const where = [];
      if (assetId > 0) {
        where.push("mi.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("mi.inspection_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("mi.inspection_date <= ?");
        params.push(end);
      }

      const rows = db.prepare(`
        SELECT
          mi.id,
          mi.asset_id,
          mi.inspection_date,
          mi.inspector_name,
          mi.notes,
          mi.created_at,
          a.asset_code,
          a.asset_name
        FROM manager_inspections mi
        JOIN assets a ON a.id = mi.asset_id
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY mi.inspection_date DESC, mi.id DESC
      `).all(...params);

      const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
      let photosByInspection = new Map();
      if (ids.length) {
        const marks = ids.map(() => "?").join(",");
        const photos = db.prepare(`
          SELECT
            id,
            ${photoInspectionCol} AS inspection_id,
            ${photoPathCol} AS file_path,
            ${photoCaptionCol} AS caption,
            ${photoCreatedCol} AS created_at
          FROM manager_inspection_photos
          WHERE ${photoInspectionCol} IN (${marks})
          ORDER BY id ASC
        `).all(...ids);
        photosByInspection = photos.reduce((m, p) => {
          const k = Number(p.inspection_id);
          if (!m.has(k)) m.set(k, []);
          m.get(k).push(p);
          return m;
        }, new Map());
      }

      return reply.send({
        ok: true,
        rows: rows.map((r) => ({
          ...r,
          photos: photosByInspection.get(Number(r.id)) || [],
        })),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/inspections", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const inspection_date = String(req.body?.inspection_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(inspection_date)) return reply.code(400).send({ ok: false, error: "inspection_date must be YYYY-MM-DD" });

      const asset = db.prepare(`SELECT id FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const ins = db.prepare(`
        INSERT INTO manager_inspections (asset_id, uuid, site_code, inspection_date, inspector_name, notes, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(asset_id, crypto.randomUUID(), site_code, inspection_date, inspector_name, notes);

      return reply.send({ ok: true, id: Number(ins.lastInsertRowid) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/inspections/:id/photo", async (req, reply) => {
    try {
      const inspectionId = Number(req.params?.id || 0);
      if (!inspectionId) return reply.code(400).send({ ok: false, error: "Invalid inspection id" });

      const inspection = db.prepare(`SELECT id FROM manager_inspections WHERE id = ?`).get(inspectionId);
      if (!inspection) return reply.code(404).send({ ok: false, error: "Inspection not found" });

      const part = await req.file();
      if (!part) return reply.code(400).send({ ok: false, error: "Upload file field named 'file'" });

      const extRaw = path.extname(part.filename || "").toLowerCase();
      const ext = [".jpg", ".jpeg", ".png", ".webp"].includes(extRaw) ? extRaw : ".jpg";
      const safe = `mi_${inspectionId}_${Date.now()}_${Math.floor(Math.random() * 100000)}${ext}`;
      const absPath = path.join(inspectionsDir, safe);
      await fs.promises.writeFile(absPath, await part.toBuffer());

      const caption = String(req.query?.caption || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const relPath = path.join("uploads", "manager-inspections", safe).replace(/\\/g, "/");

      // Legacy compatibility: some DBs require manager_inspection_id NOT NULL,
      // others use inspection_id. If both exist, write both.
      const hasInspectionId = hasColumn("manager_inspection_photos", "inspection_id");
      const hasManagerInspectionId = hasColumn("manager_inspection_photos", "manager_inspection_id");
      const linkCols = [];
      const linkVals = [];
      if (hasInspectionId) {
        linkCols.push("inspection_id");
        linkVals.push(inspectionId);
      }
      if (hasManagerInspectionId) {
        linkCols.push("manager_inspection_id");
        linkVals.push(inspectionId);
      }
      if (!linkCols.length) {
        linkCols.push(photoInspectionCol);
        linkVals.push(inspectionId);
      }

      const hasImageData = hasColumn("manager_inspection_photos", "image_data");
      const insertCols = [...linkCols, "uuid", "site_code", "file_path", ...(hasImageData ? ["image_data"] : []), "caption", "updated_at"];
      const placeholders = [...insertCols.map((c) => (c === "updated_at" ? "datetime('now')" : "?"))].join(", ");
      const ins = db.prepare(`
        INSERT INTO manager_inspection_photos (${insertCols.join(", ")})
        VALUES (${placeholders})
      `).run(...linkVals, crypto.randomUUID(), site_code, relPath, ...(hasImageData ? [relPath] : []), caption);

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        inspection_id: inspectionId,
        file_path: `/${relPath}`,
        caption,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // =====================================================
  // LDV VEHICLE CHECKS (photos + click-to-pin damage markers)
  // =====================================================
  db.prepare(`
    CREATE TABLE IF NOT EXISTS vehicle_ldv_checks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      check_date TEXT NOT NULL,
      vehicle_registration TEXT,
      odometer_km REAL,
      inspector_name TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS vehicle_ldv_check_photos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      check_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      file_path TEXT NOT NULL,
      caption TEXT,
      markers_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (check_id) REFERENCES vehicle_ldv_checks(id) ON DELETE CASCADE
    )
  `).run();

  db.prepare(`CREATE INDEX IF NOT EXISTS idx_vehicle_ldv_checks_asset ON vehicle_ldv_checks(asset_id)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_vehicle_ldv_checks_date ON vehicle_ldv_checks(check_date)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_vehicle_ldv_photos_check ON vehicle_ldv_check_photos(check_id)`).run();

  const vehicleLdvcDir = path.join(dataRoot, "uploads", "vehicle-ldv-checks");
  fs.mkdirSync(vehicleLdvcDir, { recursive: true });

  app.get("/vehicle-ldv-checks", async (req, reply) => {
    try {
      const assetId = Number(req.query?.asset_id || 0);
      const checkIdFilter = Number(req.query?.check_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const params = [];
      const where = [];
      if (checkIdFilter > 0) {
        where.push("v.id = ?");
        params.push(checkIdFilter);
      }
      if (assetId > 0) {
        where.push("v.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("v.check_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("v.check_date <= ?");
        params.push(end);
      }

      const rows = db.prepare(`
        SELECT
          v.id,
          v.asset_id,
          v.check_date,
          v.vehicle_registration,
          v.odometer_km,
          v.inspector_name,
          v.notes,
          v.created_at,
          a.asset_code,
          a.asset_name
        FROM vehicle_ldv_checks v
        JOIN assets a ON a.id = v.asset_id
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY v.check_date DESC, v.id DESC
      `).all(...params);

      const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
      let photosByCheck = new Map();
      if (ids.length) {
        const marks = ids.map(() => "?").join(",");
        const photos = db.prepare(`
          SELECT id, check_id, file_path, caption, markers_json, created_at
          FROM vehicle_ldv_check_photos
          WHERE check_id IN (${marks})
          ORDER BY id ASC
        `).all(...ids);
        photosByCheck = photos.reduce((m, p) => {
          const k = Number(p.check_id);
          if (!m.has(k)) m.set(k, []);
          let markers = [];
          try {
            markers = p.markers_json ? JSON.parse(p.markers_json) : [];
          } catch {
            markers = [];
          }
          m.get(k).push({
            ...p,
            markers: Array.isArray(markers) ? markers : [],
          });
          return m;
        }, new Map());
      }

      return reply.send({
        ok: true,
        rows: rows.map((r) => ({
          ...r,
          photos: photosByCheck.get(Number(r.id)) || [],
        })),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/vehicle-ldv-checks", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const check_date = String(req.body?.check_date || "").trim() || new Date().toISOString().slice(0, 10);
      const vehicle_registration = String(req.body?.vehicle_registration || "").trim() || null;
      const odometer_km = req.body?.odometer_km != null && req.body?.odometer_km !== "" ? Number(req.body.odometer_km) : null;
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "check_date must be YYYY-MM-DD" });

      const asset = db.prepare(`SELECT id FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const ins = db.prepare(`
        INSERT INTO vehicle_ldv_checks (
          asset_id, uuid, site_code, check_date, vehicle_registration, odometer_km, inspector_name, notes, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(asset_id, crypto.randomUUID(), site_code, check_date, vehicle_registration, odometer_km, inspector_name, notes);

      return reply.send({ ok: true, id: Number(ins.lastInsertRowid) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/vehicle-ldv-checks/:id/photo", async (req, reply) => {
    try {
      const checkId = Number(req.params?.id || 0);
      if (!checkId) return reply.code(400).send({ ok: false, error: "Invalid check id" });

      const row = db.prepare(`SELECT id FROM vehicle_ldv_checks WHERE id = ?`).get(checkId);
      if (!row) return reply.code(404).send({ ok: false, error: "Vehicle check not found" });

      const part = await req.file();
      if (!part) return reply.code(400).send({ ok: false, error: "Upload file field named 'file'" });

      const extRaw = path.extname(part.filename || "").toLowerCase();
      const ext = [".jpg", ".jpeg", ".png", ".webp"].includes(extRaw) ? extRaw : ".jpg";
      const safe = `ldv_${checkId}_${Date.now()}_${Math.floor(Math.random() * 100000)}${ext}`;
      const absPath = path.join(vehicleLdvcDir, safe);
      await fs.promises.writeFile(absPath, await part.toBuffer());

      const caption = String(req.query?.caption || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const relPath = path.join("uploads", "vehicle-ldv-checks", safe).replace(/\\/g, "/");

      const ins = db.prepare(`
        INSERT INTO vehicle_ldv_check_photos (check_id, uuid, site_code, file_path, caption, markers_json, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(checkId, crypto.randomUUID(), site_code, relPath, caption, null);

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        check_id: checkId,
        file_path: `/${relPath}`,
        caption,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.patch("/vehicle-ldv-checks/photos/:photoId", async (req, reply) => {
    try {
      const photoId = Number(req.params?.photoId || 0);
      if (!photoId) return reply.code(400).send({ ok: false, error: "Invalid photo id" });

      const photo = db.prepare(`SELECT id, check_id FROM vehicle_ldv_check_photos WHERE id = ?`).get(photoId);
      if (!photo) return reply.code(404).send({ ok: false, error: "Photo not found" });

      const body = req.body || {};
      const markers = body.markers;
      const caption = body.caption != null ? String(body.caption).trim() || null : undefined;

      if (markers != null) {
        if (!Array.isArray(markers)) return reply.code(400).send({ ok: false, error: "markers must be an array" });
        if (markers.length > 80) return reply.code(400).send({ ok: false, error: "too many markers (max 80)" });
        for (const m of markers) {
          const x = Number(m?.x);
          const y = Number(m?.y);
          if (!Number.isFinite(x) || !Number.isFinite(y) || x < 0 || x > 1 || y < 0 || y > 1) {
            return reply.code(400).send({ ok: false, error: "each marker needs x,y in [0,1] (fraction of image)" });
          }
        }
        const json = JSON.stringify(
          markers.map((m) => ({
            x: Number(m.x),
            y: Number(m.y),
            label: m.label != null ? String(m.label).slice(0, 120) : "",
            note: m.note != null ? String(m.note).slice(0, 500) : "",
          }))
        );
        db.prepare(`
          UPDATE vehicle_ldv_check_photos
          SET markers_json = ?, updated_at = datetime('now')
          WHERE id = ?
        `).run(json, photoId);
      }
      if (caption !== undefined) {
        db.prepare(`
          UPDATE vehicle_ldv_check_photos
          SET caption = ?, updated_at = datetime('now')
          WHERE id = ?
        `).run(caption, photoId);
      }

      const updated = db.prepare(`SELECT id, check_id, file_path, caption, markers_json FROM vehicle_ldv_check_photos WHERE id = ?`).get(photoId);
      let markersOut = [];
      try {
        markersOut = updated.markers_json ? JSON.parse(updated.markers_json) : [];
      } catch {
        markersOut = [];
      }

      return reply.send({
        ok: true,
        photo: {
          id: Number(updated.id),
          check_id: Number(updated.check_id),
          file_path: `/${String(updated.file_path || "").replace(/\\/g, "/")}`,
          caption: updated.caption,
          markers: markersOut,
        },
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.get("/damage-reports", async (req, reply) => {
    try {
      const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
      const assetId = Number(req.query?.asset_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const responsiblePerson = String(req.query?.responsible_person || "").trim();
      const pendingInvestigationRaw = String(req.query?.pending_investigation || "").trim();
      const hseReportAvailableRaw = String(req.query?.hse_report_available || "").trim();
      const params = [];
      const where = [];
      if (assetId > 0) {
        where.push("dr.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("dr.report_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("dr.report_date <= ?");
        params.push(end);
      }
      if (responsiblePerson) {
        where.push("UPPER(COALESCE(dr.responsible_person, '')) LIKE UPPER(?)");
        params.push(`%${responsiblePerson}%`);
      }
      if (pendingInvestigationRaw === "0" || pendingInvestigationRaw === "1") {
        where.push("COALESCE(dr.pending_investigation, 0) = ?");
        params.push(Number(pendingInvestigationRaw));
      }
      if (hseReportAvailableRaw === "0" || hseReportAvailableRaw === "1") {
        where.push("COALESCE(dr.hse_report_available, 0) = ?");
        params.push(Number(hseReportAvailableRaw));
      }

      const rows = db.prepare(`
        SELECT
          dr.id,
          dr.asset_id,
          dr.report_date,
          dr.${drInspectorCol} AS inspector_name,
          dr.hour_meter,
          dr.damage_location,
          dr.severity,
          dr.damage_description,
          dr.immediate_action,
          dr.out_of_service,
          dr.damage_time,
          dr.responsible_person,
          dr.pending_investigation,
          dr.hse_report_available,
          dr.created_at,
          a.asset_code,
          a.asset_name
        FROM manager_damage_reports dr
        JOIN assets a ON a.id = dr.asset_id
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY dr.report_date DESC, dr.id DESC
      `).all(...params);

      const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
      let photosByReport = new Map();
      if (ids.length) {
        const marks = ids.map(() => "?").join(",");
        const photos = db.prepare(`
          SELECT
            id,
            ${dmgPhotoReportCol} AS damage_report_id,
            ${dmgPhotoPathCol} AS file_path,
            ${dmgPhotoCaptionCol} AS caption,
            ${dmgPhotoCreatedCol} AS created_at
          FROM manager_damage_report_photos
          WHERE ${dmgPhotoReportCol} IN (${marks})
          ORDER BY id ASC
        `).all(...ids);
        photosByReport = photos.reduce((m, p) => {
          const k = Number(p.damage_report_id);
          if (!m.has(k)) m.set(k, []);
          m.get(k).push(p);
          return m;
        }, new Map());
      }

      return reply.send({
        ok: true,
        rows: rows.map((r) => ({
          ...r,
          photos: photosByReport.get(Number(r.id)) || [],
        })),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/damage-reports", async (req, reply) => {
    try {
      const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
      const asset_id = Number(req.body?.asset_id || 0);
      const report_date = String(req.body?.report_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const hour_meter_raw = req.body?.hour_meter;
      const hour_meter =
        hour_meter_raw == null || String(hour_meter_raw).trim() === ""
          ? null
          : Number(hour_meter_raw);
      const damage_location = String(req.body?.damage_location || "").trim() || null;
      const severity = String(req.body?.severity || "").trim() || null;
      const damage_description = String(req.body?.damage_description || "").trim() || null;
      const immediate_action = String(req.body?.immediate_action || "").trim() || null;
      const out_of_service = Number(req.body?.out_of_service || 0) ? 1 : 0;
      const damage_time = String(req.body?.damage_time || "").trim() || null;
      const responsible_person = String(req.body?.responsible_person || "").trim() || null;
      const pending_investigation = Number(req.body?.pending_investigation || 0) ? 1 : 0;
      const hse_report_available = Number(req.body?.hse_report_available || 0) ? 1 : 0;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(report_date)) return reply.code(400).send({ ok: false, error: "report_date must be YYYY-MM-DD" });
      if (hour_meter != null && !Number.isFinite(hour_meter)) {
        return reply.code(400).send({ ok: false, error: "hour_meter must be numeric" });
      }
      if (!damage_location) return reply.code(400).send({ ok: false, error: "damage_location is required" });
      if (!severity) return reply.code(400).send({ ok: false, error: "severity is required" });
      if (!damage_description) return reply.code(400).send({ ok: false, error: "damage_description is required" });
      if (!immediate_action) return reply.code(400).send({ ok: false, error: "immediate_action is required" });
      if (damage_time && !/^\d{2}:\d{2}$/.test(damage_time)) {
        return reply.code(400).send({ ok: false, error: "damage_time must be HH:MM" });
      }

      const asset = db.prepare(`SELECT id FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const ins = db.prepare(`
        INSERT INTO manager_damage_reports (
          asset_id, uuid, site_code, report_date, ${drInspectorCol}, hour_meter,
          damage_location, severity, damage_description, immediate_action, out_of_service,
          damage_time, responsible_person, pending_investigation, hse_report_available, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(
        asset_id,
        crypto.randomUUID(),
        site_code,
        report_date,
        inspector_name,
        hour_meter,
        damage_location,
        severity,
        damage_description,
        immediate_action,
        out_of_service,
        damage_time,
        responsible_person,
        pending_investigation,
        hse_report_available
      );

      return reply.send({ ok: true, id: Number(ins.lastInsertRowid) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/damage-reports/:id/photo", async (req, reply) => {
    try {
      const reportId = Number(req.params?.id || 0);
      if (!reportId) return reply.code(400).send({ ok: false, error: "Invalid damage report id" });

      const report = db.prepare(`SELECT id FROM manager_damage_reports WHERE id = ?`).get(reportId);
      if (!report) return reply.code(404).send({ ok: false, error: "Damage report not found" });

      const part = await req.file();
      if (!part) return reply.code(400).send({ ok: false, error: "Upload file field named 'file'" });

      const extRaw = path.extname(part.filename || "").toLowerCase();
      const ext = [".jpg", ".jpeg", ".png", ".webp"].includes(extRaw) ? extRaw : ".jpg";
      const safe = `mdr_${reportId}_${Date.now()}_${Math.floor(Math.random() * 100000)}${ext}`;
      const absPath = path.join(damageReportsDir, safe);
      await fs.promises.writeFile(absPath, await part.toBuffer());

      const caption = String(req.query?.caption || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const relPath = path.join("uploads", "manager-damage-reports", safe).replace(/\\/g, "/");

      const hasReportId = hasColumn("manager_damage_report_photos", "damage_report_id");
      const hasLegacyReportId = hasColumn("manager_damage_report_photos", "manager_damage_report_id");
      const linkCols = [];
      const linkVals = [];
      if (hasReportId) {
        linkCols.push("damage_report_id");
        linkVals.push(reportId);
      }
      if (hasLegacyReportId) {
        linkCols.push("manager_damage_report_id");
        linkVals.push(reportId);
      }
      if (!linkCols.length) {
        linkCols.push(dmgPhotoReportCol);
        linkVals.push(reportId);
      }

      const hasImageData = hasColumn("manager_damage_report_photos", "image_data");
      const insertCols = [...linkCols, "uuid", "site_code", "file_path", ...(hasImageData ? ["image_data"] : []), "caption", "updated_at"];
      const placeholders = [...insertCols.map((c) => (c === "updated_at" ? "datetime('now')" : "?"))].join(", ");
      const ins = db.prepare(`
        INSERT INTO manager_damage_report_photos (${insertCols.join(", ")})
        VALUES (${placeholders})
      `).run(...linkVals, crypto.randomUUID(), site_code, relPath, ...(hasImageData ? [relPath] : []), caption);

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        damage_report_id: reportId,
        file_path: `/${relPath}`,
        caption,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });
}