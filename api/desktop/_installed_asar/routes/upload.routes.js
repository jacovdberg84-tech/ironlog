// IRONLOG/api/routes/upload.routes.js
import multipart from "@fastify/multipart";
import { db } from "../db/client.js";
import {
  parseCsvToObjects,
  requireHeaders,
  asFloat,
  asInt,
  asBool01,
  asDateYYYYMMDD
} from "../utils/csvImporter.js";

export default async function uploadRoutes(app) {
  // enable file uploads
  await app.register(multipart, {
    limits: { fileSize: 20 * 1024 * 1024 } // 20MB
  });

  // Helper: get asset id by code
  const getAssetIdByCode = db.prepare(`SELECT id FROM assets WHERE asset_code = ?`);
  const getPartIdByCode = db.prepare(`SELECT id FROM parts WHERE part_code = ?`);
  const getWorkOrderById = db.prepare(`SELECT id, asset_id FROM work_orders WHERE id = ?`);

  db.prepare(`
    CREATE TABLE IF NOT EXISTS store_allocations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      work_order_id INTEGER,
      part_id INTEGER NOT NULL,
      quantity REAL NOT NULL,
      allocation_date TEXT NOT NULL,
      issued_by TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT,
      FOREIGN KEY (work_order_id) REFERENCES work_orders(id) ON DELETE SET NULL,
      FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE RESTRICT
    )
  `).run();

  function hasColumn(table, col) {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name) === col);
  }

  if (!hasColumn("assets", "baseline_fuel_l_per_hour")) {
    db.prepare(`ALTER TABLE assets ADD COLUMN baseline_fuel_l_per_hour REAL DEFAULT 5.0`).run();
  }
  if (!hasColumn("fuel_logs", "hours_run")) {
    db.prepare(`ALTER TABLE fuel_logs ADD COLUMN hours_run REAL`).run();
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

  function requireUploadWrite(req, reply) {
    return requireRoles(req, reply, ["admin", "supervisor", "stores"]);
  }

  // -------------------------
  // POST /api/upload/assets
  // Columns: asset_code, asset_name, category, is_standby, active
  // -------------------------
  app.post("/assets", async (req, reply) => {
    if (!requireUploadWrite(req, reply)) return;
    const file = await req.file();
    if (!file) return reply.code(400).send({ error: "Upload a CSV file field named 'file'." });

    const buf = await file.toBuffer();
    const rows = parseCsvToObjects(buf);

    requireHeaders(rows, ["asset_code", "asset_name", "category"]);

    const insert = db.prepare(`
      INSERT INTO assets (asset_code, asset_name, category, is_standby, active)
      VALUES (?, ?, ?, ?, ?)
      ON CONFLICT(asset_code) DO UPDATE SET
        asset_name = excluded.asset_name,
        category = excluded.category,
        is_standby = excluded.is_standby,
        active = excluded.active
    `);

    const tx = db.transaction(() => {
      for (const r of rows) {
        const code = String(r.asset_code).trim();
        const name = String(r.asset_name).trim();
        const cat = String(r.category).trim();
        const standby = asBool01(r.is_standby, 0);
        const active = asBool01(r.active, 1);

        if (!code || !name || !cat) continue;
        insert.run(code, name, cat, standby, active);
      }
    });

    tx();

    return reply.send({ ok: true, imported: rows.length });
  });

    // -------------------------
  // POST /api/upload/hours
  //
  // Supports BOTH formats:
  //
  // FORMAT A (simple / legacy):
  // asset_code, work_date, hours_run, is_used, operator, notes
  //
  // FORMAT B (full daily-input style):
  // asset_code, work_date, scheduled_hours, opening_hours, closing_hours,
  // (optional hours_run), is_used, operator, notes
  //
  // Rules:
  // - If hours_run missing but opening+closing provided => hours_run = closing - opening
  // - If is_used=0 (standby) => force scheduled_hours=0 and hours_run=0
  // - hours_run must be 0..24
  // -------------------------
  app.post("/hours", async (req, reply) => {
    if (!requireUploadWrite(req, reply)) return;
    const file = await req.file();
    if (!file) return reply.code(400).send({ error: "Upload a CSV file field named 'file'." });

    const buf = await file.toBuffer();
    const rows = parseCsvToObjects(buf);

    // Only require identity columns
    requireHeaders(rows, ["asset_code", "work_date"]);

    const upsert = db.prepare(`
      INSERT INTO daily_hours (
        asset_id, work_date,
        scheduled_hours, opening_hours, closing_hours,
        hours_run, is_used, operator, notes
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(asset_id, work_date) DO UPDATE SET
        scheduled_hours = excluded.scheduled_hours,
        opening_hours = excluded.opening_hours,
        closing_hours = excluded.closing_hours,
        hours_run = excluded.hours_run,
        is_used = excluded.is_used,
        operator = excluded.operator,
        notes = excluded.notes
    `);

    const ensureAssetHours = db.prepare(`
      INSERT INTO asset_hours (asset_id, total_hours, last_updated)
      SELECT ?, 0, datetime('now')
      WHERE NOT EXISTS (
        SELECT 1 FROM asset_hours WHERE asset_id = ?
      )
    `);

    const getLatestClosing = db.prepare(`
      SELECT closing_hours
      FROM daily_hours
      WHERE asset_id = ?
        AND closing_hours IS NOT NULL
      ORDER BY work_date DESC, id DESC
      LIMIT 1
    `);

    const getTotalRun = db.prepare(`
      SELECT COALESCE(SUM(hours_run), 0) AS total_hours
      FROM daily_hours
      WHERE asset_id = ?
        AND is_used = 1
        AND hours_run > 0
    `);

    const getExistingTotal = db.prepare(`
      SELECT total_hours
      FROM asset_hours
      WHERE asset_id = ?
    `);

    const updateAssetHours = db.prepare(`
      UPDATE asset_hours
      SET total_hours = ?, last_updated = datetime('now')
      WHERE asset_id = ?
    `);

    const tx = db.transaction(() => {
      const touchedAssetIds = new Set();

      for (const r of rows) {
        const assetCode = String(r.asset_code).trim();
        const asset = getAssetIdByCode.get(assetCode);
        if (!asset) continue;

        const date = asDateYYYYMMDD(r.work_date);

        const used = asBool01(r.is_used, 1);

        const operator =
          r.operator != null && String(r.operator).trim() !== ""
            ? String(r.operator).trim()
            : null;

        const notes =
          r.notes != null && String(r.notes).trim() !== ""
            ? String(r.notes).trim()
            : null;

        // Optional full-mode fields
        const scheduled =
          r.scheduled_hours != null && String(r.scheduled_hours).trim() !== ""
            ? asFloat(r.scheduled_hours, 0)
            : null;

        const opening =
          r.opening_hours != null && String(r.opening_hours).trim() !== ""
            ? asFloat(r.opening_hours, null)
            : null;

        const closing =
          r.closing_hours != null && String(r.closing_hours).trim() !== ""
            ? asFloat(r.closing_hours, null)
            : null;

        // hours_run: provided OR computed from hourmeter delta
        let hrs =
          r.hours_run != null && String(r.hours_run).trim() !== ""
            ? asFloat(r.hours_run, 0)
            : (opening != null && closing != null ? (closing - opening) : 0);

        if (!Number.isFinite(hrs)) hrs = 0;

        // Standby honesty
        let schedFinal = scheduled;
        if (!used) {
          hrs = 0;
          schedFinal = 0;
        }

        // Range sanity
        if (hrs < 0 || hrs > 24) continue;

        upsert.run(
          asset.id,
          date,
          schedFinal,
          opening,
          closing,
          hrs,
          used,
          operator,
          notes
        );

        touchedAssetIds.add(Number(asset.id));
      }

      for (const assetId of touchedAssetIds) {
        ensureAssetHours.run(assetId, assetId);

        const latestClosingRow = getLatestClosing.get(assetId);
        const totalRunRow = getTotalRun.get(assetId);
        const existingRow = getExistingTotal.get(assetId);

        const latestClosing =
          latestClosingRow && latestClosingRow.closing_hours != null
            ? Number(latestClosingRow.closing_hours)
            : null;

        const derivedTotal =
          latestClosing != null && Number.isFinite(latestClosing)
            ? latestClosing
            : Number(totalRunRow?.total_hours || 0);

        const existingTotal = Number(existingRow?.total_hours || 0);
        const nextTotal = Math.max(existingTotal, Number.isFinite(derivedTotal) ? derivedTotal : 0);

        updateAssetHours.run(nextTotal, assetId);
      }

      return touchedAssetIds.size;
    });

    const syncedAssets = tx();

    return reply.send({ ok: true, imported: rows.length, synced_asset_hours: syncedAssets });
  });
  // -------------------------
  // POST /api/upload/fuel
  // Columns: asset_code, log_date, liters, source, hours_run(optional)
  // -------------------------
  app.post("/fuel", async (req, reply) => {
    if (!requireUploadWrite(req, reply)) return;
    const file = await req.file();
    if (!file) return reply.code(400).send({ error: "Upload a CSV file field named 'file'." });

    const buf = await file.toBuffer();
    const rows = parseCsvToObjects(buf);

    requireHeaders(rows, ["asset_code", "log_date", "liters"]);

    const insert = db.prepare(`
      INSERT INTO fuel_logs (asset_id, log_date, liters, source, hours_run)
      VALUES (?, ?, ?, ?, ?)
    `);

    const tx = db.transaction(() => {
      for (const r of rows) {
        const assetCode = String(r.asset_code).trim();
        const asset = getAssetIdByCode.get(assetCode);
        if (!asset) continue;

        const date = asDateYYYYMMDD(r.log_date);
        const liters = asFloat(r.liters, 0);
        const source = r.source != null && String(r.source).trim() !== "" ? String(r.source).trim() : null;
        const hoursRunRaw = r.hours_run != null && String(r.hours_run).trim() !== ""
          ? asFloat(r.hours_run, 0)
          : null;
        const hoursRun = hoursRunRaw != null && hoursRunRaw >= 0 ? hoursRunRaw : null;

        if (liters <= 0) continue;
        insert.run(asset.id, date, liters, source, hoursRun);
      }
    });

    tx();

    return reply.send({ ok: true, imported: rows.length });
  });

  // -------------------------
  // POST /api/upload/oil
  // Columns: asset_code, log_date, oil_type, quantity
  // -------------------------
  app.post("/oil", async (req, reply) => {
    if (!requireUploadWrite(req, reply)) return;
    const file = await req.file();
    if (!file) return reply.code(400).send({ error: "Upload a CSV file field named 'file'." });

    const buf = await file.toBuffer();
    const rows = parseCsvToObjects(buf);

    requireHeaders(rows, ["asset_code", "log_date", "quantity"]);

    const insert = db.prepare(`
      INSERT INTO oil_logs (asset_id, log_date, oil_type, quantity)
      VALUES (?, ?, ?, ?)
    `);

    const tx = db.transaction(() => {
      for (const r of rows) {
        const assetCode = String(r.asset_code).trim();
        const asset = getAssetIdByCode.get(assetCode);
        if (!asset) continue;

        const date = asDateYYYYMMDD(r.log_date);
        const qty = asFloat(r.quantity, 0);
        const oilType = r.oil_type != null && String(r.oil_type).trim() !== "" ? String(r.oil_type).trim() : null;

        if (qty <= 0) continue;
        insert.run(asset.id, date, oilType, qty);
      }
    });

    tx();

    return reply.send({ ok: true, imported: rows.length });
  });

  // -------------------------
  // POST /api/upload/parts
  // Columns: part_code, part_name, critical, min_stock
  // -------------------------
  app.post("/parts", async (req, reply) => {
    if (!requireUploadWrite(req, reply)) return;
    const file = await req.file();
    if (!file) return reply.code(400).send({ error: "Upload a CSV file field named 'file'." });

    const buf = await file.toBuffer();
    const rows = parseCsvToObjects(buf);

    requireHeaders(rows, ["part_code", "part_name"]);

    const upsert = db.prepare(`
      INSERT INTO parts (part_code, part_name, critical, min_stock)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(part_code) DO UPDATE SET
        part_name = excluded.part_name,
        critical = excluded.critical,
        min_stock = excluded.min_stock
    `);

    const tx = db.transaction(() => {
      for (const r of rows) {
        const code = String(r.part_code).trim();
        const name = String(r.part_name).trim();
        const critical = asBool01(r.critical, 0);
        const min = asInt(r.min_stock, 0);

        if (!code || !name) continue;
        upsert.run(code, name, critical, min);
      }
    });

    tx();

    return reply.send({ ok: true, imported: rows.length });
  });

  // -------------------------
  // POST /api/upload/opening_stock
  // Columns: part_code, quantity
  // Creates stock_movements as opening_balance
  // -------------------------
  app.post("/opening_stock", async (req, reply) => {
    if (!requireUploadWrite(req, reply)) return;
    const file = await req.file();
    if (!file) return reply.code(400).send({ error: "Upload a CSV file field named 'file'." });

    const buf = await file.toBuffer();
    const rows = parseCsvToObjects(buf);

    requireHeaders(rows, ["part_code", "quantity"]);

    const getPartId = db.prepare(`SELECT id FROM parts WHERE part_code = ?`);
    const insertMove = db.prepare(`
      INSERT INTO stock_movements (part_id, quantity, movement_type, reference)
      VALUES (?, ?, 'in', 'opening_balance')
    `);

    const tx = db.transaction(() => {
      for (const r of rows) {
        const code = String(r.part_code).trim();
        const part = getPartId.get(code);
        if (!part) continue;

        const qty = asInt(r.quantity, 0);
        if (qty <= 0) continue;

        insertMove.run(part.id, qty);
      }
    });

    tx();

    return reply.send({ ok: true, imported: rows.length });
  });

  // -------------------------
  // POST /api/upload/store_allocations
  // Columns:
  // part_code, quantity, allocation_date, asset_code, work_order_id, issued_by, notes
  // Rules:
  // - asset_code or work_order_id required per row
  // - if work_order_id provided, asset_code optional; if both provided they must match
  // - quantity must be > 0 and available in stock
  // -------------------------
  app.post("/store_allocations", async (req, reply) => {
    if (!requireUploadWrite(req, reply)) return;
    const file = await req.file();
    if (!file) return reply.code(400).send({ error: "Upload a CSV file field named 'file'." });

    const buf = await file.toBuffer();
    const rows = parseCsvToObjects(buf);
    requireHeaders(rows, ["part_code", "quantity"]);

    const getOnHand = db.prepare(`
      SELECT IFNULL(SUM(quantity), 0) AS on_hand
      FROM stock_movements
      WHERE part_id = ?
    `);
    const insertMove = db.prepare(`
      INSERT INTO stock_movements (part_id, quantity, movement_type, reference)
      VALUES (?, ?, 'out', ?)
    `);
    const insertAlloc = db.prepare(`
      INSERT INTO store_allocations (
        asset_id, work_order_id, part_id, quantity, allocation_date, issued_by, notes
      )
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `);

    const tx = db.transaction(() => {
      let imported = 0;
      let skipped = 0;

      for (const r of rows) {
        const partCode = String(r.part_code || "").trim();
        const quantity = asFloat(r.quantity, 0);
        const allocationDate = asDateYYYYMMDD(
          r.allocation_date != null && String(r.allocation_date).trim() !== ""
            ? r.allocation_date
            : new Date().toISOString().slice(0, 10)
        );
        const assetCode = String(r.asset_code || "").trim();
        const workOrderId =
          r.work_order_id != null && String(r.work_order_id).trim() !== ""
            ? asInt(r.work_order_id, 0)
            : 0;
        const issuedBy =
          r.issued_by != null && String(r.issued_by).trim() !== ""
            ? String(r.issued_by).trim()
            : null;
        const notes =
          r.notes != null && String(r.notes).trim() !== ""
            ? String(r.notes).trim()
            : null;

        if (!partCode || !Number.isFinite(quantity) || quantity <= 0) {
          skipped++;
          continue;
        }

        const part = getPartIdByCode.get(partCode);
        if (!part) {
          skipped++;
          continue;
        }

        let wo = null;
        if (workOrderId > 0) {
          wo = getWorkOrderById.get(workOrderId);
          if (!wo) {
            skipped++;
            continue;
          }
        }

        let assetId = 0;
        if (assetCode) {
          const asset = getAssetIdByCode.get(assetCode);
          if (!asset) {
            skipped++;
            continue;
          }
          assetId = Number(asset.id);
        }
        if (!assetId && wo) assetId = Number(wo.asset_id);
        if (!assetId) {
          skipped++;
          continue;
        }
        if (wo && assetCode && Number(wo.asset_id) !== assetId) {
          skipped++;
          continue;
        }

        const onHand = Number(getOnHand.get(part.id)?.on_hand || 0);
        if (onHand < quantity) {
          skipped++;
          continue;
        }

        const reference = wo ? `work_order:${wo.id}` : `asset:${assetId}:stores`;
        insertMove.run(part.id, -Math.abs(quantity), reference);
        insertAlloc.run(
          assetId,
          wo ? Number(wo.id) : null,
          Number(part.id),
          quantity,
          allocationDate,
          issuedBy,
          notes
        );
        imported++;
      }

      return { imported, skipped };
    });

    const result = tx();
    return reply.send({ ok: true, ...result, total_rows: rows.length });
  });

  // -------------------------
  // POST /api/upload/fuel_baseline
  // Columns: asset_code, baseline_fuel_l_per_hour
  // -------------------------
  app.post("/fuel_baseline", async (req, reply) => {
    if (!requireUploadWrite(req, reply)) return;
    const file = await req.file();
    if (!file) return reply.code(400).send({ error: "Upload a CSV file field named 'file'." });

    const buf = await file.toBuffer();
    const rows = parseCsvToObjects(buf);
    requireHeaders(rows, ["asset_code", "baseline_fuel_l_per_hour"]);

    const updateBaseline = db.prepare(`
      UPDATE assets
      SET baseline_fuel_l_per_hour = ?
      WHERE id = ?
    `);

    const tx = db.transaction(() => {
      let imported = 0;
      let skipped = 0;
      for (const r of rows) {
        const code = String(r.asset_code || "").trim();
        const baseline = asFloat(r.baseline_fuel_l_per_hour, 0);
        if (!code || !Number.isFinite(baseline) || baseline <= 0) {
          skipped++;
          continue;
        }

        const asset = getAssetIdByCode.get(code);
        if (!asset) {
          skipped++;
          continue;
        }

        updateBaseline.run(baseline, asset.id);
        imported++;
      }
      return { imported, skipped };
    });

    const result = tx();
    return reply.send({ ok: true, ...result, total_rows: rows.length });
  });

  // Simple info endpoint (what files/columns to use)
  app.get("/", async () => {
    return {
      ok: true,
      endpoints: [
        "POST /api/upload/assets",
        "POST /api/upload/hours",
        "POST /api/upload/fuel",
        "POST /api/upload/oil",
        "POST /api/upload/parts",
        "POST /api/upload/opening_stock",
        "POST /api/upload/store_allocations",
        "POST /api/upload/fuel_baseline"
      ],
      note: "Upload as multipart/form-data field name 'file'. Dates must be YYYY-MM-DD."
    };
  });
}