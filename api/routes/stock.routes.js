// IRONLOG/api/routes/stock.routes.js
import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";
import {
  ensureMasterDataSchema,
  normalizeMdmCode,
  validateAgainstMdmPolicy,
  validatePartGovernanceOptional,
} from "../utils/masterdataGovernance.js";

export default async function stockRoutes(app) {
  ensureAuditTable(db);
  ensureMasterDataSchema();

  function hasColumn(table, col) {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name) === col);
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

  function getSiteCode(req) {
    return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
  }

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

  db.prepare(`
    CREATE INDEX IF NOT EXISTS idx_store_allocations_asset_date
    ON store_allocations(asset_id, allocation_date)
  `).run();

  db.prepare(`
    CREATE INDEX IF NOT EXISTS idx_store_allocations_part
    ON store_allocations(part_id)
  `).run();

  db.prepare(`
    CREATE INDEX IF NOT EXISTS idx_store_allocations_wo
    ON store_allocations(work_order_id)
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS stock_locations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      location_code TEXT NOT NULL UNIQUE,
      location_name TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  const existingLocations = db.prepare(`SELECT COUNT(*) AS c FROM stock_locations`).get();
  if (Number(existingLocations?.c || 0) === 0) {
    db.prepare(`
      INSERT INTO stock_locations (location_code, location_name, active)
      VALUES
        ('MAIN', 'Main Store', 1),
        ('LUBE', 'Lube Store', 1),
        ('WORKSHOP', 'Workshop Store', 1)
    `).run();
  }

  if (!hasColumn("stock_movements", "location_id")) {
    db.prepare(`ALTER TABLE stock_movements ADD COLUMN location_id INTEGER`).run();
  }
  if (!hasColumn("stock_movements", "bin_id")) {
    db.prepare(`ALTER TABLE stock_movements ADD COLUMN bin_id INTEGER`).run();
  }
  if (!hasColumn("stock_movements", "cost_center_code")) {
    db.prepare(`ALTER TABLE stock_movements ADD COLUMN cost_center_code TEXT`).run();
  }
  if (!hasColumn("store_allocations", "location_id")) {
    db.prepare(`ALTER TABLE store_allocations ADD COLUMN location_id INTEGER`).run();
  }
  if (!hasColumn("store_allocations", "bin_id")) {
    db.prepare(`ALTER TABLE store_allocations ADD COLUMN bin_id INTEGER`).run();
  }
  if (!hasColumn("store_allocations", "cost_center_code")) {
    db.prepare(`ALTER TABLE store_allocations ADD COLUMN cost_center_code TEXT`).run();
  }
  if (!hasColumn("parts", "unit_cost")) {
    db.prepare(`ALTER TABLE parts ADD COLUMN unit_cost REAL DEFAULT 0`).run();
  }
  if (!hasColumn("stock_movements", "unit_cost_usd")) {
    db.prepare(`ALTER TABLE stock_movements ADD COLUMN unit_cost_usd REAL`).run();
  }
  if (!hasColumn("stock_movements", "cost_currency")) {
    db.prepare(`ALTER TABLE stock_movements ADD COLUMN cost_currency TEXT`).run();
  }
  if (!hasColumn("stock_movements", "cost_input")) {
    db.prepare(`ALTER TABLE stock_movements ADD COLUMN cost_input REAL`).run();
  }

  db.prepare(`
    CREATE TABLE IF NOT EXISTS cost_settings (
      key TEXT PRIMARY KEY,
      value REAL NOT NULL DEFAULT 0,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  const upsertFxDefault = db.prepare(`
    INSERT INTO cost_settings (key, value, updated_at) VALUES (?, ?, datetime('now'))
    ON CONFLICT(key) DO NOTHING
  `);
  upsertFxDefault.run("zar_per_usd", 18.5);
  upsertFxDefault.run("mzn_per_usd", 64);

  db.prepare(`
    CREATE TABLE IF NOT EXISTS stock_bins (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      location_id INTEGER NOT NULL,
      bin_code TEXT NOT NULL,
      bin_name TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(location_id, bin_code),
      FOREIGN KEY (location_id) REFERENCES stock_locations(id) ON DELETE CASCADE
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS stock_min_max (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      part_id INTEGER NOT NULL,
      location_id INTEGER NOT NULL,
      bin_id INTEGER,
      min_qty REAL NOT NULL DEFAULT 0,
      max_qty REAL NOT NULL DEFAULT 0,
      reorder_qty REAL,
      target_days INTEGER,
      updated_by TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(part_id, location_id, bin_id),
      FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE CASCADE,
      FOREIGN KEY (location_id) REFERENCES stock_locations(id) ON DELETE CASCADE,
      FOREIGN KEY (bin_id) REFERENCES stock_bins(id) ON DELETE SET NULL
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS stock_cycle_sessions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      location_id INTEGER,
      bin_id INTEGER,
      status TEXT NOT NULL DEFAULT 'draft',
      planned_date TEXT,
      counted_by TEXT,
      submitted_at TEXT,
      approved_by TEXT,
      approved_at TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (location_id) REFERENCES stock_locations(id) ON DELETE SET NULL,
      FOREIGN KEY (bin_id) REFERENCES stock_bins(id) ON DELETE SET NULL
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS stock_cycle_lines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      session_id INTEGER NOT NULL,
      part_id INTEGER NOT NULL,
      system_qty REAL NOT NULL DEFAULT 0,
      counted_qty REAL NOT NULL DEFAULT 0,
      variance_qty REAL NOT NULL DEFAULT 0,
      status TEXT NOT NULL DEFAULT 'draft',
      reason TEXT,
      approval_request_id INTEGER,
      approved_at TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(session_id, part_id),
      FOREIGN KEY (session_id) REFERENCES stock_cycle_sessions(id) ON DELETE CASCADE,
      FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE CASCADE
    )
  `).run();

  function getFxRate(key, fallback) {
    const row = db.prepare(`SELECT value FROM cost_settings WHERE key = ?`).get(key);
    const v = Number(row?.value);
    return Number.isFinite(v) && v > 0 ? v : fallback;
  }

  function normalizeOilTypeInput(raw, fallback = null) {
    const v = String(raw ?? "").trim();
    if (!v) return fallback;
    const blocked = new Set(["admin", "supervisor", "manager", "stores", "artisan", "operator"]);
    if (blocked.has(v.toLowerCase())) return fallback;
    return v;
  }

  /** Local currency amount per one unit of stock → USD per unit (rate = local units per 1 USD). */
  function unitCostToUsd(amount, currency) {
    const c = String(currency || "USD").toUpperCase();
    const n = Number(amount);
    if (!Number.isFinite(n) || n < 0) return null;
    if (n === 0) return 0;
    if (c === "USD") return n;
    if (c === "ZAR") {
      const zarPerUsd = getFxRate("zar_per_usd", 18.5);
      return n / zarPerUsd;
    }
    if (c === "MZN") {
      const mznPerUsd = getFxRate("mzn_per_usd", 64);
      return n / mznPerUsd;
    }
    return null;
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

  const getAssetByCode = db.prepare(`
    SELECT id, asset_code, asset_name
    FROM assets
    WHERE asset_code = ?
  `);
  const getPartByCode = db.prepare(`
    SELECT id, part_code, part_name, unit_cost
    FROM parts
    WHERE part_code = ?
  `);

  const insertPart = db.prepare(`
    INSERT INTO parts (
      part_code, part_name, critical, min_stock, unit_cost,
      department_code, default_supplier_code, data_owner_username
    )
    VALUES (?, ?, 0, 0, COALESCE(?, 0), ?, ?, ?)
  `);
  const getWoById = db.prepare(`
    SELECT id, asset_id, status
    FROM work_orders
    WHERE id = ?
  `);
  const getOnHand = db.prepare(`
    SELECT IFNULL(SUM(quantity), 0) AS on_hand
    FROM stock_movements
    WHERE part_id = ?
  `);
  const insertMove = db.prepare(`
    INSERT INTO stock_movements (part_id, quantity, movement_type, reference, location_id, bin_id, cost_center_code)
    VALUES (?, ?, 'out', ?, ?, ?, ?)
  `);
  const insertGenericMove = db.prepare(`
    INSERT INTO stock_movements (
      part_id, quantity, movement_type, reference, location_id,
      bin_id, cost_center_code, unit_cost_usd, cost_currency, cost_input
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const updatePartUnitCostUsd = db.prepare(`
    UPDATE parts SET unit_cost = ? WHERE id = ?
  `);
  const insertAlloc = db.prepare(`
    INSERT INTO store_allocations (
      asset_id, work_order_id, part_id, quantity, allocation_date, issued_by, notes, location_id, bin_id, cost_center_code
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  const getLocationByCode = db.prepare(`
    SELECT id, location_code, location_name, active
    FROM stock_locations
    WHERE location_code = ?
  `);
  const getBinByCodeAtLocation = db.prepare(`
    SELECT id, location_id, bin_code, bin_name, active
    FROM stock_bins
    WHERE location_id = ? AND UPPER(TRIM(bin_code)) = UPPER(TRIM(?))
    LIMIT 1
  `);

  // Stock on hand summary
  app.get("/onhand", async () => {
    const govCols = [];
    if (hasColumn("parts", "department_code")) govCols.push("p.department_code");
    if (hasColumn("parts", "default_supplier_code")) govCols.push("p.default_supplier_code");
    if (hasColumn("parts", "data_owner_username")) govCols.push("p.data_owner_username");
    const govSql = govCols.length ? `, ${govCols.join(", ")}` : "";
    const rows = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        p.critical,
        p.min_stock,
        p.unit_cost
        ${govSql},
        IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      GROUP BY p.id
      ORDER BY p.critical DESC, p.part_code ASC
    `).all();

    return rows.map(r => ({
      ...r,
      critical: Boolean(r.critical),
      on_hand: Number(r.on_hand),
      unit_cost: Number(r.unit_cost || 0),
      stock_value: Number((Number(r.on_hand || 0) * Number(r.unit_cost || 0)).toFixed(2)),
      below_min: Number(r.on_hand) < Number(r.min_stock)
    }));
  });

  // Stock locations
  // GET /api/stock/locations?active=1
  app.get("/locations", async (req, reply) => {
    const onlyActive = String(req.query?.active || "1").trim() !== "0";
    const rows = db.prepare(`
      SELECT id, location_code, location_name, active, created_at
      FROM stock_locations
      WHERE (? = 0 OR active = 1)
      ORDER BY location_code ASC
      LIMIT 300
    `).all(onlyActive ? 1 : 0).map((r) => ({
      ...r,
      active: Number(r.active || 0),
    }));
    return reply.send({ ok: true, rows });
  });

  // POST /api/stock/locations
  // Body: { location_code, location_name?, active? }
  app.post("/locations", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const location_code = String(req.body?.location_code || "").trim().toUpperCase();
    const location_name = String(req.body?.location_name || "").trim() || null;
    const active = req.body?.active === 0 || req.body?.active === false ? 0 : 1;
    if (!location_code) return reply.code(400).send({ error: "location_code is required" });

    const existing = getLocationByCode.get(location_code);
    if (existing) {
      db.prepare(`
        UPDATE stock_locations
        SET location_name = COALESCE(?, location_name),
            active = ?
        WHERE id = ?
      `).run(location_name, active, Number(existing.id));
      return reply.send({ ok: true, id: Number(existing.id), location_code, updated: true });
    }

    const ins = db.prepare(`
      INSERT INTO stock_locations (location_code, location_name, active)
      VALUES (?, ?, ?)
    `).run(location_code, location_name, active);

    return reply.send({ ok: true, id: Number(ins.lastInsertRowid), location_code, created: true });
  });

  // Stock bins (location-scoped)
  // GET /api/stock/bins?location_code=&active=1
  app.get("/bins", async (req, reply) => {
    const location_code = String(req.query?.location_code || "").trim().toUpperCase();
    const onlyActive = String(req.query?.active || "1").trim() !== "0";
    const where = [];
    const params = [];
    if (location_code) {
      const loc = getLocationByCode.get(location_code);
      if (!loc) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
      where.push("b.location_id = ?");
      params.push(Number(loc.id));
    }
    if (onlyActive) where.push("b.active = 1");
    const rows = db.prepare(`
      SELECT
        b.id, b.location_id, b.bin_code, b.bin_name, b.active, b.created_at,
        l.location_code, l.location_name
      FROM stock_bins b
      JOIN stock_locations l ON l.id = b.location_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY l.location_code ASC, b.bin_code ASC
      LIMIT 500
    `).all(...params).map((r) => ({
      ...r,
      active: Number(r.active || 0),
    }));
    return reply.send({ ok: true, rows });
  });

  // POST /api/stock/bins
  // Body: { location_code, bin_code, bin_name?, active? }
  app.post("/bins", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const location_code = String(req.body?.location_code || "").trim().toUpperCase();
    const bin_code = String(req.body?.bin_code || "").trim().toUpperCase();
    const bin_name = String(req.body?.bin_name || "").trim() || null;
    const active = req.body?.active === 0 || req.body?.active === false ? 0 : 1;
    if (!location_code || !bin_code) return reply.code(400).send({ error: "location_code and bin_code are required" });
    const location = getLocationByCode.get(location_code);
    if (!location) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    const existing = getBinByCodeAtLocation.get(Number(location.id), bin_code);
    if (existing) {
      db.prepare(`
        UPDATE stock_bins
        SET bin_name = COALESCE(?, bin_name),
            active = ?
        WHERE id = ?
      `).run(bin_name, active, Number(existing.id));
      return reply.send({ ok: true, updated: true, id: Number(existing.id), location_code, bin_code });
    }
    const ins = db.prepare(`
      INSERT INTO stock_bins (location_id, bin_code, bin_name, active)
      VALUES (?, ?, ?, ?)
    `).run(Number(location.id), bin_code, bin_name, active);
    return reply.send({ ok: true, created: true, id: Number(ins.lastInsertRowid), location_code, bin_code });
  });

  // Inventory depth by part/location/bin
  // GET /api/stock/depth?part_code=&location_code=&bin_code=
  app.get("/depth", async (req, reply) => {
    const part_code = String(req.query?.part_code || "").trim();
    const location_code = String(req.query?.location_code || "").trim().toUpperCase();
    const bin_code = String(req.query?.bin_code || "").trim().toUpperCase();
    const where = [];
    const params = [];
    if (part_code) {
      where.push("p.part_code = ?");
      params.push(part_code);
    }
    if (location_code) {
      const loc = getLocationByCode.get(location_code);
      if (!loc) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
      where.push("COALESCE(sm.location_id, ?) = ?");
      params.push(Number(loc.id), Number(loc.id));
      if (bin_code) {
        const bin = getBinByCodeAtLocation.get(Number(loc.id), bin_code);
        if (!bin) return reply.code(404).send({ error: `bin_code not found at location ${location_code}: ${bin_code}` });
        where.push("COALESCE(sm.bin_id, ?) = ?");
        params.push(Number(bin.id), Number(bin.id));
      }
    }
    const rows = db.prepare(`
      SELECT
        p.id AS part_id,
        p.part_code,
        p.part_name,
        COALESCE(sm.location_id, l.id) AS location_id,
        COALESCE(l.location_code, 'UNSPECIFIED') AS location_code,
        COALESCE(l.location_name, 'Unspecified') AS location_name,
        sm.bin_id,
        COALESCE(b.bin_code, 'UNSPECIFIED') AS bin_code,
        COALESCE(b.bin_name, 'Unspecified') AS bin_name,
        COALESCE(SUM(sm.quantity), 0) AS on_hand,
        0 AS reserved
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      LEFT JOIN stock_locations l ON l.id = sm.location_id
      LEFT JOIN stock_bins b ON b.id = sm.bin_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      GROUP BY p.id, location_id, sm.bin_id
      ORDER BY p.part_code ASC, location_code ASC, bin_code ASC
      LIMIT 2000
    `).all(...params).map((r) => {
      const on_hand = Number(r.on_hand || 0);
      const reserved = Number(r.reserved || 0);
      const on_order = Number(
        db.prepare(`
          SELECT COALESCE(SUM(CASE WHEN qty_requested > qty_received THEN (qty_requested - qty_received) ELSE 0 END), 0) AS qty
          FROM procurement_requisitions
          WHERE part_id = ?
            AND LOWER(status) IN ('approved', 'approved_all', 'po_ready', 'partially_received')
        `).get(Number(r.part_id || 0))?.qty || 0
      );
      return {
        ...r,
        on_hand: Number(on_hand.toFixed(2)),
        reserved: Number(reserved.toFixed(2)),
        on_order: Number(on_order.toFixed(2)),
        available: Number((on_hand - reserved).toFixed(2)),
      };
    });
    return reply.send({ ok: true, rows });
  });

  // Min-max policy upsert
  // POST /api/stock/min-max
  app.post("/min-max", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const part_code = String(req.body?.part_code || "").trim();
    const location_code = String(req.body?.location_code || "").trim().toUpperCase();
    const bin_code = String(req.body?.bin_code || "").trim().toUpperCase();
    if (!part_code || !location_code) return reply.code(400).send({ error: "part_code and location_code are required" });
    const part = getPartByCode.get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const location = getLocationByCode.get(location_code);
    if (!location) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    const bin = bin_code ? getBinByCodeAtLocation.get(Number(location.id), bin_code) : null;
    if (bin_code && !bin) return reply.code(404).send({ error: `bin_code not found at location ${location_code}: ${bin_code}` });
    const min_qty = Math.max(0, Number(req.body?.min_qty || 0));
    const max_qty = Math.max(min_qty, Number(req.body?.max_qty || min_qty));
    const reorder_qty = req.body?.reorder_qty != null ? Math.max(0, Number(req.body.reorder_qty || 0)) : null;
    const target_days = req.body?.target_days != null ? Math.max(0, Number(req.body.target_days || 0)) : null;
    db.prepare(`
      INSERT INTO stock_min_max (
        part_id, location_id, bin_id, min_qty, max_qty, reorder_qty, target_days, updated_by, updated_at
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      ON CONFLICT(part_id, location_id, bin_id) DO UPDATE SET
        min_qty = excluded.min_qty,
        max_qty = excluded.max_qty,
        reorder_qty = excluded.reorder_qty,
        target_days = excluded.target_days,
        updated_by = excluded.updated_by,
        updated_at = datetime('now')
    `).run(
      Number(part.id),
      Number(location.id),
      bin ? Number(bin.id) : null,
      min_qty,
      max_qty,
      reorder_qty,
      target_days,
      String(req.headers["x-user-name"] || "session-user")
    );
    return reply.send({ ok: true, part_code, location_code, bin_code: bin ? bin.bin_code : null, min_qty, max_qty, reorder_qty, target_days });
  });

  // GET /api/stock/min-max?part_code=&location_code=&bin_code=
  app.get("/min-max", async (req, reply) => {
    const part_code = String(req.query?.part_code || "").trim();
    const location_code = String(req.query?.location_code || "").trim().toUpperCase();
    const bin_code = String(req.query?.bin_code || "").trim().toUpperCase();
    const where = [];
    const params = [];
    if (part_code) {
      where.push("p.part_code = ?");
      params.push(part_code);
    }
    if (location_code) {
      where.push("l.location_code = ?");
      params.push(location_code);
    }
    if (bin_code) {
      where.push("UPPER(COALESCE(b.bin_code, '')) = ?");
      params.push(bin_code);
    }
    const rows = db.prepare(`
      SELECT
        mm.id,
        p.part_code,
        p.part_name,
        l.location_code,
        l.location_name,
        b.bin_code,
        b.bin_name,
        mm.min_qty,
        mm.max_qty,
        mm.reorder_qty,
        mm.target_days,
        mm.updated_by,
        mm.updated_at
      FROM stock_min_max mm
      JOIN parts p ON p.id = mm.part_id
      JOIN stock_locations l ON l.id = mm.location_id
      LEFT JOIN stock_bins b ON b.id = mm.bin_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY p.part_code ASC, l.location_code ASC, COALESCE(b.bin_code, '') ASC
      LIMIT 1000
    `).all(...params).map((r) => ({
      ...r,
      min_qty: Number(r.min_qty || 0),
      max_qty: Number(r.max_qty || 0),
      reorder_qty: r.reorder_qty == null ? null : Number(r.reorder_qty || 0),
      target_days: r.target_days == null ? null : Number(r.target_days || 0),
    }));
    return reply.send({ ok: true, rows });
  });

  // Replenishment suggestions from min-max and current on-hand
  // GET /api/stock/replenishment-suggestions?location_code=&bin_code=
  app.get("/replenishment-suggestions", async (req, reply) => {
    const location_code = String(req.query?.location_code || "").trim().toUpperCase();
    const bin_code = String(req.query?.bin_code || "").trim().toUpperCase();
    const where = [];
    const params = [];
    if (location_code) {
      where.push("l.location_code = ?");
      params.push(location_code);
    }
    if (bin_code) {
      where.push("UPPER(COALESCE(b.bin_code, '')) = ?");
      params.push(bin_code);
    }
    const policyRows = db.prepare(`
      SELECT
        mm.id,
        mm.part_id,
        p.part_code,
        p.part_name,
        l.id AS location_id,
        l.location_code,
        b.id AS bin_id,
        b.bin_code,
        mm.min_qty,
        mm.max_qty,
        mm.reorder_qty
      FROM stock_min_max mm
      JOIN parts p ON p.id = mm.part_id
      JOIN stock_locations l ON l.id = mm.location_id
      LEFT JOIN stock_bins b ON b.id = mm.bin_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY p.part_code ASC
      LIMIT 1500
    `).all(...params);
    const rows = policyRows.map((r) => {
      const on_hand = Number(
        db.prepare(`
          SELECT COALESCE(SUM(quantity), 0) AS q
          FROM stock_movements
          WHERE part_id = ?
            AND COALESCE(location_id, ?) = ?
            AND COALESCE(bin_id, 0) = COALESCE(?, 0)
        `).get(Number(r.part_id), Number(r.location_id), Number(r.location_id), r.bin_id == null ? null : Number(r.bin_id))?.q || 0
      );
      const need = Math.max(0, Number(r.min_qty || 0) - on_hand);
      const target = Number(r.reorder_qty || 0) > 0 ? Number(r.reorder_qty || 0) : Math.max(0, Number(r.max_qty || 0) - on_hand);
      return {
        part_code: r.part_code,
        part_name: r.part_name,
        location_code: r.location_code,
        bin_code: r.bin_code || null,
        min_qty: Number(r.min_qty || 0),
        max_qty: Number(r.max_qty || 0),
        on_hand: Number(on_hand.toFixed(2)),
        shortage_qty: Number(need.toFixed(2)),
        suggested_order_qty: Number(target.toFixed(2)),
        needs_replenishment: need > 0,
      };
    }).filter((r) => r.needs_replenishment);
    return reply.send({ ok: true, rows });
  });

  // Cycle count sessions
  // POST /api/stock/cycle-sessions
  app.post("/cycle-sessions", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const location_code = String(req.body?.location_code || "").trim().toUpperCase();
    const bin_code = String(req.body?.bin_code || "").trim().toUpperCase();
    const planned_date = String(req.body?.planned_date || "").trim() || new Date().toISOString().slice(0, 10);
    const notes = String(req.body?.notes || "").trim() || null;
    const location = location_code ? getLocationByCode.get(location_code) : null;
    if (location_code && !location) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    const bin = (location && bin_code) ? getBinByCodeAtLocation.get(Number(location.id), bin_code) : null;
    if (bin_code && !location_code) return reply.code(400).send({ error: "location_code is required when bin_code is provided" });
    if (location && bin_code && !bin) return reply.code(404).send({ error: `bin_code not found at location ${location_code}: ${bin_code}` });
    const ins = db.prepare(`
      INSERT INTO stock_cycle_sessions (
        location_id, bin_id, status, planned_date, counted_by, notes
      ) VALUES (?, ?, 'draft', ?, ?, ?)
    `).run(
      location ? Number(location.id) : null,
      bin ? Number(bin.id) : null,
      planned_date,
      String(req.headers["x-user-name"] || "session-user"),
      notes
    );
    return reply.send({ ok: true, session_id: Number(ins.lastInsertRowid), location_code: location ? location.location_code : null, bin_code: bin ? bin.bin_code : null });
  });

  // GET /api/stock/cycle-sessions?status=
  app.get("/cycle-sessions", async (req, reply) => {
    const status = String(req.query?.status || "").trim().toLowerCase();
    const where = [];
    const params = [];
    if (status) {
      where.push("LOWER(s.status) = ?");
      params.push(status);
    }
    const rows = db.prepare(`
      SELECT
        s.id, s.status, s.planned_date, s.counted_by, s.submitted_at, s.approved_by, s.approved_at, s.notes, s.created_at,
        l.location_code, b.bin_code,
        (SELECT COUNT(*) FROM stock_cycle_lines cl WHERE cl.session_id = s.id) AS line_count,
        (SELECT COALESCE(SUM(ABS(cl.variance_qty)), 0) FROM stock_cycle_lines cl WHERE cl.session_id = s.id) AS variance_abs
      FROM stock_cycle_sessions s
      LEFT JOIN stock_locations l ON l.id = s.location_id
      LEFT JOIN stock_bins b ON b.id = s.bin_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY s.id DESC
      LIMIT 300
    `).all(...params).map((r) => ({
      ...r,
      line_count: Number(r.line_count || 0),
      variance_abs: Number(Number(r.variance_abs || 0).toFixed(2)),
    }));
    return reply.send({ ok: true, rows });
  });

  // POST /api/stock/cycle-sessions/:id/lines/upsert
  // Body: { lines: [{ part_code, counted_qty, reason? }] }
  app.post("/cycle-sessions/:id/lines/upsert", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const session = db.prepare(`SELECT id, status, location_id, bin_id FROM stock_cycle_sessions WHERE id = ?`).get(id);
    if (!session) return reply.code(404).send({ error: "cycle session not found" });
    if (!["draft", "counting"].includes(String(session.status || "").toLowerCase())) {
      return reply.code(409).send({ error: `cannot edit lines when status is ${session.status}` });
    }
    const lines = Array.isArray(req.body?.lines) ? req.body.lines : [];
    if (!lines.length) return reply.code(400).send({ error: "lines array is required" });
    const upsert = db.prepare(`
      INSERT INTO stock_cycle_lines (
        session_id, part_id, system_qty, counted_qty, variance_qty, status, reason
      ) VALUES (?, ?, ?, ?, ?, 'draft', ?)
      ON CONFLICT(session_id, part_id) DO UPDATE SET
        system_qty = excluded.system_qty,
        counted_qty = excluded.counted_qty,
        variance_qty = excluded.variance_qty,
        reason = excluded.reason
    `);
    const tx = db.transaction(() => {
      for (const line of lines) {
        const part_code = String(line?.part_code || "").trim();
        const counted_qty = Number(line?.counted_qty ?? NaN);
        if (!part_code || !Number.isFinite(counted_qty) || counted_qty < 0) continue;
        const part = getPartByCode.get(part_code);
        if (!part) continue;
        const system_qty = Number(
          db.prepare(`
            SELECT COALESCE(SUM(quantity), 0) AS q
            FROM stock_movements
            WHERE part_id = ?
              AND COALESCE(location_id, 0) = COALESCE(?, 0)
              AND COALESCE(bin_id, 0) = COALESCE(?, 0)
          `).get(Number(part.id), session.location_id == null ? null : Number(session.location_id), session.bin_id == null ? null : Number(session.bin_id))?.q || 0
        );
        const variance = Number((counted_qty - system_qty).toFixed(2));
        upsert.run(
          id,
          Number(part.id),
          Number(system_qty.toFixed(2)),
          Number(counted_qty.toFixed(2)),
          variance,
          line?.reason ? String(line.reason).trim() : null
        );
      }
      db.prepare(`UPDATE stock_cycle_sessions SET status = 'counting' WHERE id = ? AND status = 'draft'`).run(id);
    });
    tx();
    return reply.send({ ok: true, session_id: id });
  });

  app.post("/cycle-sessions/:id/submit", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const session = db.prepare(`SELECT id, status FROM stock_cycle_sessions WHERE id = ?`).get(id);
    if (!session) return reply.code(404).send({ error: "cycle session not found" });
    if (!["draft", "counting"].includes(String(session.status || "").toLowerCase())) {
      return reply.code(409).send({ error: `cannot submit when status is ${session.status}` });
    }
    db.prepare(`
      UPDATE stock_cycle_sessions
      SET status = 'submitted', submitted_at = datetime('now')
      WHERE id = ?
    `).run(id);
    return reply.send({ ok: true, session_id: id, status: "submitted" });
  });

  app.post("/cycle-sessions/:id/approve", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const session = db.prepare(`SELECT id, status, location_id, bin_id FROM stock_cycle_sessions WHERE id = ?`).get(id);
    if (!session) return reply.code(404).send({ error: "cycle session not found" });
    if (!["submitted", "approved"].includes(String(session.status || "").toLowerCase())) {
      return reply.code(409).send({ error: `cannot approve when status is ${session.status}` });
    }
    const lines = db.prepare(`
      SELECT id, part_id, variance_qty
      FROM stock_cycle_lines
      WHERE session_id = ?
    `).all(id);
    const tx = db.transaction(() => {
      for (const line of lines) {
        const variance = Number(line.variance_qty || 0);
        if (!Number.isFinite(variance) || variance === 0) continue;
        db.prepare(`
          INSERT INTO stock_movements (part_id, quantity, movement_type, reference, location_id, bin_id)
          VALUES (?, ?, 'adjust', ?, ?, ?)
        `).run(
          Number(line.part_id),
          variance,
          `cycle_session:${id}:line:${Number(line.id)}`,
          session.location_id == null ? null : Number(session.location_id),
          session.bin_id == null ? null : Number(session.bin_id)
        );
        db.prepare(`
          UPDATE stock_cycle_lines
          SET status = 'approved', approved_at = datetime('now')
          WHERE id = ?
        `).run(Number(line.id));
      }
      db.prepare(`
        UPDATE stock_cycle_sessions
        SET status = 'approved', approved_by = ?, approved_at = datetime('now')
        WHERE id = ?
      `).run(String(req.headers["x-user-name"] || "session-user"), id);
    });
    tx();
    return reply.send({ ok: true, session_id: id, status: "approved" });
  });

  // Stock monitor summary + recent movements
  // GET /api/stock/monitor?part_code=
  app.get("/monitor", async (req, reply) => {
    const part_code = String(req.query?.part_code || "").trim();

    const where = [];
    const params = [];
    if (part_code) {
      where.push("p.part_code LIKE ?");
      params.push(`%${part_code}%`);
    }

    const rows = db.prepare(`
      SELECT
        p.id,
        p.part_code,
        p.part_name,
        p.critical,
        p.min_stock,
        p.unit_cost,
        IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      GROUP BY p.id
      ORDER BY p.critical DESC, on_hand ASC, p.part_code ASC
      LIMIT 300
    `).all(...params).map((r) => ({
      ...r,
      critical: Boolean(r.critical),
      on_hand: Number(r.on_hand || 0),
      unit_cost: Number(r.unit_cost || 0),
      stock_value: Number((Number(r.on_hand || 0) * Number(r.unit_cost || 0)).toFixed(2)),
      below_min: Number(r.on_hand || 0) < Number(r.min_stock || 0),
    }));

    const summary = {
      total_parts: rows.length,
      below_min: rows.filter((r) => r.below_min).length,
      critical_below_min: rows.filter((r) => r.below_min && r.critical).length,
      total_on_hand: Number(rows.reduce((acc, r) => acc + Number(r.on_hand || 0), 0).toFixed(2)),
      total_stock_value: Number(rows.reduce((acc, r) => acc + Number(r.stock_value || 0), 0).toFixed(2)),
    };

    const movementDateExpr = hasColumn("stock_movements", "created_at")
      ? "sm.created_at"
      : "sm.movement_date";

    const recent = db.prepare(`
      SELECT
        sm.id,
        ${movementDateExpr} AS created_at,
        sm.movement_type,
        sm.quantity,
        sm.reference,
        l.location_code,
        l.location_name,
        p.part_code,
        p.part_name
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      LEFT JOIN stock_locations l ON l.id = sm.location_id
      ORDER BY sm.id DESC
      LIMIT 30
    `).all().map((r) => ({
      ...r,
      quantity: Number(r.quantity || 0),
    }));

    return reply.send({ ok: true, summary, rows, recent });
  });

  // Inventory control summary
  // GET /api/stock/control-summary?part_code=
  app.get("/control-summary", async (req, reply) => {
    const part_code = String(req.query?.part_code || "").trim();

    const totalPartsRow = db.prepare(`SELECT COUNT(*) AS c FROM parts`).get();
    const belowMinRow = db.prepare(`
      SELECT COUNT(*) AS c
      FROM (
        SELECT p.id, IFNULL(SUM(sm.quantity), 0) AS on_hand, p.min_stock
        FROM parts p
        LEFT JOIN stock_movements sm ON sm.part_id = p.id
        GROUP BY p.id
        HAVING on_hand < IFNULL(p.min_stock, 0)
      )
    `).get();
    const lubeBelowRows = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        p.min_stock,
        IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      WHERE (
        LOWER(IFNULL(p.part_code, '')) LIKE '%oil%' OR
        LOWER(IFNULL(p.part_name, '')) LIKE '%oil%' OR
        LOWER(IFNULL(p.part_code, '')) LIKE '%lube%' OR
        LOWER(IFNULL(p.part_name, '')) LIKE '%lube%' OR
        LOWER(IFNULL(p.part_code, '')) LIKE '%grease%' OR
        LOWER(IFNULL(p.part_name, '')) LIKE '%grease%'
      )
      GROUP BY p.id
      HAVING on_hand < IFNULL(p.min_stock, 0)
      ORDER BY (IFNULL(p.min_stock, 0) - on_hand) DESC, p.part_code ASC
      LIMIT 10
    `).all().map((r) => ({
      ...r,
      on_hand: Number(r.on_hand || 0),
      min_stock: Number(r.min_stock || 0),
      shortage: Number((Number(r.min_stock || 0) - Number(r.on_hand || 0)).toFixed(2)),
    }));

    let part = null;
    let part_summary = null;
    if (part_code) {
      part = getPartByCode.get(part_code);
      if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
      const on_hand = Number(getOnHand.get(part.id)?.on_hand || 0);
      const min_stock = Number(part.min_stock || 0);

      const movementDateExpr = hasColumn("stock_movements", "created_at")
        ? "datetime(created_at)"
        : "datetime(movement_date)";
      const movementSummary = db.prepare(`
        SELECT
          IFNULL(SUM(CASE WHEN quantity > 0 THEN quantity ELSE 0 END), 0) AS qty_in_30d,
          IFNULL(SUM(CASE WHEN quantity < 0 THEN ABS(quantity) ELSE 0 END), 0) AS qty_out_30d,
          IFNULL(SUM(quantity), 0) AS net_30d,
          COUNT(*) AS movement_count_30d
        FROM stock_movements
        WHERE part_id = ?
          AND ${movementDateExpr} >= datetime('now', '-30 days')
      `).get(part.id);
      const movementCount7d = db.prepare(`
        SELECT COUNT(*) AS c
        FROM stock_movements
        WHERE part_id = ?
          AND ${movementDateExpr} >= datetime('now', '-7 days')
      `).get(part.id);

      part_summary = {
        part_code: part.part_code,
        part_name: part.part_name,
        on_hand,
        min_stock,
        below_min: on_hand < min_stock,
        qty_in_30d: Number(movementSummary?.qty_in_30d || 0),
        qty_out_30d: Number(movementSummary?.qty_out_30d || 0),
        net_30d: Number(movementSummary?.net_30d || 0),
        movement_count_30d: Number(movementSummary?.movement_count_30d || 0),
        movement_count_7d: Number(movementCount7d?.c || 0),
      };
    }

    return reply.send({
      ok: true,
      summary: {
        total_parts: Number(totalPartsRow?.c || 0),
        below_min_total: Number(belowMinRow?.c || 0),
        lube_below_min_count: lubeBelowRows.length,
      },
      part: part_summary,
      low_lube_rows: lubeBelowRows,
    });
  });

  // Set minimum stock for a specific part
  // POST /api/stock/part-minimum
  // Body: { part_code, min_stock }
  app.post("/part-minimum", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const part_code = String(req.body?.part_code || "").trim();
    const min_stock = Number(req.body?.min_stock ?? NaN);
    if (!part_code) return reply.code(400).send({ error: "part_code is required" });
    if (!Number.isFinite(min_stock) || min_stock < 0) {
      return reply.code(400).send({ error: "min_stock must be a valid number >= 0" });
    }
    const part = getPartByCode.get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const before = Number(part.min_stock || 0);
    db.prepare(`UPDATE parts SET min_stock = ? WHERE id = ?`).run(Number(min_stock.toFixed(2)), Number(part.id));
    const on_hand = Number(getOnHand.get(part.id)?.on_hand || 0);

    writeAudit(db, req, {
      module: "stock",
      action: "set_part_minimum",
      entity_type: "part",
      entity_id: part_code,
      payload: { min_stock_before: before, min_stock_after: Number(min_stock.toFixed(2)), on_hand },
    });

    return reply.send({
      ok: true,
      part_code,
      min_stock_before: before,
      min_stock_after: Number(min_stock.toFixed(2)),
      on_hand,
      below_min: on_hand < Number(min_stock.toFixed(2)),
    });
  });

  // Submit cycle count as stock adjustment approval request
  // POST /api/stock/cycle-count
  // Body: { part_code, counted_qty, reason? }
  app.post("/cycle-count", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const part_code = String(req.body?.part_code || "").trim();
    const counted_qty = Number(req.body?.counted_qty ?? NaN);
    const reason = String(req.body?.reason || "").trim() || "cycle_count";

    if (!part_code) return reply.code(400).send({ error: "part_code is required" });
    if (!Number.isFinite(counted_qty) || counted_qty < 0) {
      return reply.code(400).send({ error: "counted_qty must be a valid number >= 0" });
    }

    const part = getPartByCode.get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const on_hand = Number(getOnHand.get(part.id)?.on_hand || 0);
    const delta = Number((counted_qty - on_hand).toFixed(2));
    if (delta === 0) {
      return reply.send({
        ok: true,
        no_change: true,
        message: "Counted quantity matches current on-hand. No adjustment required.",
        part_code,
        on_hand,
        counted_qty: Number(counted_qty.toFixed(2)),
      });
    }

    const reqRole = getRole(req);
    const reqUser = String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
    const reference = `cycle_count:${reason}`;
    const approvalPayload = JSON.stringify({
      part_code,
      quantity: delta,
      reference,
    });

    const ins = db.prepare(`
      INSERT INTO approval_requests (
        module, action, entity_type, entity_id, status, payload_json, requested_by, requested_role
      )
      VALUES ('stock', 'adjust_movement', 'part', ?, 'pending', ?, ?, ?)
    `).run(part_code, approvalPayload, reqUser, reqRole);
    const request_id = Number(ins.lastInsertRowid);

    writeAudit(db, req, {
      module: "stock",
      action: "cycle_count_request",
      entity_type: "part",
      entity_id: part_code,
      payload: {
        request_id,
        on_hand_before: on_hand,
        counted_qty: Number(counted_qty.toFixed(2)),
        adjustment_qty: delta,
        reason,
      },
    });

    return reply.send({
      ok: true,
      pending_approval: true,
      request_id,
      part_code,
      on_hand_before: on_hand,
      counted_qty: Number(counted_qty.toFixed(2)),
      adjustment_qty: delta,
      reference,
      message: "Cycle count submitted for approval",
    });
  });

  // Lube stock lookup
  // GET /api/stock/lube-onhand?q=&location_code=
  app.get("/lube-onhand", async (req, reply) => {
    const q = String(req.query?.q || "").trim();
    const location_code = String(req.query?.location_code || "").trim().toUpperCase();
    const location = location_code ? getLocationByCode.get(location_code) : null;
    if (location_code && !location) {
      return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    }
    const where = [
      "(" +
        "LOWER(IFNULL(p.part_code, '')) LIKE '%oil%' OR " +
        "LOWER(IFNULL(p.part_name, '')) LIKE '%oil%' OR " +
        "LOWER(IFNULL(p.part_code, '')) LIKE '%lube%' OR " +
        "LOWER(IFNULL(p.part_name, '')) LIKE '%lube%' OR " +
        "LOWER(IFNULL(p.part_code, '')) LIKE '%grease%' OR " +
        "LOWER(IFNULL(p.part_name, '')) LIKE '%grease%'" +
      ")",
    ];
    const params = [];

    if (q) {
      where.push("(p.part_code LIKE ? OR p.part_name LIKE ?)");
      params.push(`%${q}%`, `%${q}%`);
    }

    // Location filter:
    // - When a location is provided, include that location AND legacy rows with NULL location_id
    //   (older data before location tracking).
    const joinMovements = location
      ? "LEFT JOIN stock_movements sm ON sm.part_id = p.id AND (sm.location_id = ? OR sm.location_id IS NULL)"
      : "LEFT JOIN stock_movements sm ON sm.part_id = p.id";
    if (location) params.unshift(Number(location.id));

    const rows = db.prepare(`
      SELECT
        p.id,
        p.part_code,
        p.part_name,
        p.min_stock,
        IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      ${joinMovements}
      WHERE ${where.join(" AND ")}
      GROUP BY p.id
      ORDER BY on_hand ASC, p.part_code ASC
      LIMIT 40
    `).all(...params).map((r) => ({
      id: Number(r.id),
      part_code: r.part_code,
      part_name: r.part_name,
      min_stock: Number(r.min_stock || 0),
      on_hand: Number(r.on_hand || 0),
      below_min: Number(r.on_hand || 0) < Number(r.min_stock || 0),
    }));

    const exact = q
      ? rows.find((r) => String(r.part_code || "").toLowerCase() === q.toLowerCase()) || null
      : null;

    return reply.send({ ok: true, q, location_code: location ? location.location_code : null, exact, rows });
  });

  // Set minimum stock for lube/oil items
  // POST /api/stock/lube-minimums
  // Body: { min_stock?: number } default 210
  app.post("/lube-minimums", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const min_stock_input = Number(req.body?.min_stock ?? 210);
    if (!Number.isFinite(min_stock_input) || min_stock_input < 0) {
      return reply.code(400).send({ error: "min_stock must be a valid number >= 0" });
    }
    const min_stock = Number(min_stock_input.toFixed(2));

    const lubeParts = db.prepare(`
      SELECT id, part_code, part_name, min_stock
      FROM parts
      WHERE (
        LOWER(IFNULL(part_code, '')) LIKE '%oil%' OR
        LOWER(IFNULL(part_name, '')) LIKE '%oil%' OR
        LOWER(IFNULL(part_code, '')) LIKE '%lube%' OR
        LOWER(IFNULL(part_name, '')) LIKE '%lube%' OR
        LOWER(IFNULL(part_code, '')) LIKE '%grease%' OR
        LOWER(IFNULL(part_name, '')) LIKE '%grease%'
      )
      AND LOWER(IFNULL(part_name, '')) NOT LIKE '%filter%'
      ORDER BY part_code ASC
      LIMIT 500
    `).all();

    const upd = db.prepare(`UPDATE parts SET min_stock = ? WHERE id = ?`);
    const tx = db.transaction(() => {
      for (const p of lubeParts) {
        upd.run(min_stock, Number(p.id));
      }
    });
    tx();

    const rows = lubeParts.map((p) => ({
      part_code: p.part_code,
      part_name: p.part_name,
      min_stock_before: Number(p.min_stock || 0),
      min_stock_after: min_stock,
    }));

    writeAudit(db, req, {
      module: "stock",
      action: "set_lube_minimums",
      entity_type: "parts",
      payload: {
        min_stock,
        updated_count: rows.length,
      },
    });

    return reply.send({
      ok: true,
      min_stock,
      updated_count: rows.length,
      rows,
    });
  });

  // GET /api/stock/fx-settings — ZAR/MZN per USD (for converting receipts to USD)
  app.get("/fx-settings", async () => {
    return {
      ok: true,
      zar_per_usd: getFxRate("zar_per_usd", 18.5),
      mzn_per_usd: getFxRate("mzn_per_usd", 64),
      note:
        "Receipt unit cost in ZAR or MZN is converted to USD as: USD = local_amount / rate (rate = local currency units per 1 USD).",
    };
  });

  // POST /api/stock/fx-settings — update rates (admin/supervisor/stores)
  app.post("/fx-settings", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const body = req.body || {};
    const zar = body.zar_per_usd != null ? Number(body.zar_per_usd) : null;
    const mzn = body.mzn_per_usd != null ? Number(body.mzn_per_usd) : null;
    const upsert = db.prepare(`
      INSERT INTO cost_settings (key, value, updated_at) VALUES (?, ?, datetime('now'))
      ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = datetime('now')
    `);
    if (zar != null) {
      if (!Number.isFinite(zar) || zar <= 0) return reply.code(400).send({ error: "zar_per_usd must be > 0" });
      upsert.run("zar_per_usd", zar);
    }
    if (mzn != null) {
      if (!Number.isFinite(mzn) || mzn <= 0) return reply.code(400).send({ error: "mzn_per_usd must be > 0" });
      upsert.run("mzn_per_usd", mzn);
    }
    return reply.send({
      ok: true,
      zar_per_usd: getFxRate("zar_per_usd", 18.5),
      mzn_per_usd: getFxRate("mzn_per_usd", 64),
    });
  });

  // Manual stock movement entry
  // POST /api/stock/movement
  // Body: { part_code, quantity, movement_type: in|out|adjust, reference?, location_code?,
  //         part_name?, create_if_missing?, unit_cost?, cost_currency? (USD|ZAR|MZN) }
  app.post("/movement", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const site = getSiteCode(req);
    const body = req.body || {};
    const part_code = String(body.part_code || "").trim();
    const movement_type = String(body.movement_type || "").trim().toLowerCase();
    const qtyIn = Number(body.quantity ?? 0);
    const reference = String(body.reference || "manual_entry").trim() || "manual_entry";
    const location_code = String(body.location_code || "").trim().toUpperCase();
    const bin_code = String(body.bin_code || "").trim().toUpperCase();
    const part_name = body.part_name != null ? String(body.part_name || "").trim() : "";
    const create_if_missing = body.create_if_missing === true || body.create_if_missing === 1;

    const department_code =
      body.department_code != null && String(body.department_code).trim() !== ""
        ? normalizeMdmCode(body.department_code)
        : null;
    const default_supplier_code =
      body.default_supplier_code != null && String(body.default_supplier_code).trim() !== ""
        ? normalizeMdmCode(body.default_supplier_code)
        : null;
    const data_owner_username =
      body.data_owner_username != null && String(body.data_owner_username).trim() !== ""
        ? String(body.data_owner_username).trim()
        : null;

    const rawCost = body.unit_cost != null && body.unit_cost !== "" ? Number(body.unit_cost) : null;
    const cost_currency = String(body.cost_currency || "USD").trim().toUpperCase();

    if (!part_code) return reply.code(400).send({ error: "part_code is required" });
    if (!["in", "out", "adjust"].includes(movement_type)) {
      return reply.code(400).send({ error: "movement_type must be one of: in, out, adjust" });
    }
    if (!Number.isFinite(qtyIn) || qtyIn === 0) {
      return reply.code(400).send({ error: "quantity must be a non-zero number" });
    }

    let unit_cost_usd = null;
    let cost_input = null;
    let cost_curr_out = null;
    if (movement_type === "in" && rawCost != null && Number.isFinite(rawCost) && rawCost > 0) {
      if (!["USD", "ZAR", "MZN"].includes(cost_currency)) {
        return reply.code(400).send({ error: "cost_currency must be USD, ZAR, or MZN" });
      }
      cost_input = rawCost;
      cost_curr_out = cost_currency;
      unit_cost_usd = unitCostToUsd(rawCost, cost_currency);
      if (unit_cost_usd == null || !Number.isFinite(unit_cost_usd)) {
        return reply.code(400).send({ error: "could not convert unit cost to USD — check FX rates" });
      }
    }

    let part = getPartByCode.get(part_code);
    if (!part) {
      // Allow receiving stock IN for brand-new stock numbers
      if (movement_type === "in" && create_if_missing) {
        if (!part_name) return reply.code(400).send({ error: "part_name is required when creating a new stock item" });
        const intakeObj = { department_code, default_supplier_code, data_owner_username };
        const polP = validateAgainstMdmPolicy(site, "part_stock_intake", intakeObj);
        if (!polP.ok) {
          return reply.code(400).send({ error: `missing required fields: ${polP.missing.join(", ")}` });
        }
        const pv = validatePartGovernanceOptional(site, intakeObj);
        if (!pv.ok) return reply.code(400).send({ error: pv.error });
        try {
          const uc = unit_cost_usd != null ? Number(unit_cost_usd.toFixed(6)) : 0;
          insertPart.run(part_code, part_name, uc, department_code, default_supplier_code, data_owner_username);
        } catch (e) {
          // In case of race, re-read
        }
        part = getPartByCode.get(part_code);
      }
    }
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const location = location_code ? getLocationByCode.get(location_code) : null;
    if (location_code && !location) {
      return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    }
    const bin = (location && bin_code) ? getBinByCodeAtLocation.get(Number(location.id), bin_code) : null;
    if (bin_code && !location_code) return reply.code(400).send({ error: "location_code is required when bin_code is provided" });
    if (location && bin_code && !bin) return reply.code(404).send({ error: `bin_code not found at location ${location_code}: ${bin_code}` });
    const cost_center_code =
      body.cost_center_code != null && String(body.cost_center_code).trim() !== ""
        ? String(body.cost_center_code).trim()
        : null;

    let qty = qtyIn;
    if (movement_type === "in") qty = Math.abs(qtyIn);
    if (movement_type === "out") qty = -Math.abs(qtyIn);
    if (movement_type === "adjust") qty = qtyIn;

    if (movement_type === "adjust") {
      const reqRole = getRole(req);
      const reqUser = String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
      const approvalPayload = JSON.stringify({
        part_code,
        quantity: qty,
        reference,
        location_code: location ? location.location_code : null,
      });
      const ins = db.prepare(`
        INSERT INTO approval_requests (
          module, action, entity_type, entity_id, status, payload_json, requested_by, requested_role
        )
        VALUES ('stock', 'adjust_movement', 'part', ?, 'pending', ?, ?, ?)
      `).run(part_code, approvalPayload, reqUser, reqRole);
      const request_id = Number(ins.lastInsertRowid);

      writeAudit(db, req, {
        module: "stock",
        action: "adjust_request",
        entity_type: "part",
        entity_id: part_code,
        payload: { request_id, quantity: qty, reference },
      });

      return reply.send({
        ok: true,
        pending_approval: true,
        request_id,
        part_code,
        movement_type,
        quantity: qty,
        reference,
        location_code: location ? location.location_code : null,
        message: "Stock adjustment submitted for approval",
      });
    }

    const onHand = Number(getOnHand.get(part.id)?.on_hand || 0);
    if (movement_type === "out" && onHand < Math.abs(qty)) {
      return reply.code(409).send({
        error: "insufficient stock",
        part_code,
        on_hand: onHand,
        requested: Math.abs(qty),
      });
    }

    const tx = db.transaction(() => {
      const ins = insertGenericMove.run(
        part.id,
        qty,
        movement_type,
        reference,
        location ? Number(location.id) : null,
        bin ? Number(bin.id) : null,
        cost_center_code,
        unit_cost_usd != null ? Number(unit_cost_usd.toFixed(6)) : null,
        cost_curr_out,
        cost_input,
      );
      const mid = Number(ins.lastInsertRowid);
      if (movement_type === "in" && unit_cost_usd != null) {
        updatePartUnitCostUsd.run(Number(unit_cost_usd.toFixed(6)), part.id);
      }
      return mid;
    });

    const movement_id = tx();
    const on_hand_after = Number(getOnHand.get(part.id)?.on_hand || 0);
    const line_value_usd =
      movement_type === "in" && unit_cost_usd != null
        ? Number((Math.abs(qty) * unit_cost_usd).toFixed(2))
        : null;

    writeAudit(db, req, {
      module: "stock",
      action: "manual_movement",
      entity_type: "part",
      entity_id: part_code,
      payload: {
        movement_type,
        quantity: qty,
        reference,
        on_hand_before: onHand,
        on_hand_after,
        unit_cost_usd,
        cost_currency: cost_curr_out,
        cost_input,
        line_value_usd,
      },
    });

    return reply.send({
      ok: true,
      movement_id,
      part_code,
      movement_type,
      quantity: qty,
      on_hand_before: onHand,
      on_hand_after,
      reference,
      location_code: location ? location.location_code : null,
      bin_code: bin ? bin.bin_code : null,
      cost_center_code,
      unit_cost_usd: unit_cost_usd != null ? Number(unit_cost_usd.toFixed(6)) : null,
      cost_currency: cost_curr_out,
      cost_input,
      line_value_usd,
      fx_rates: {
        zar_per_usd: getFxRate("zar_per_usd", 18.5),
        mzn_per_usd: getFxRate("mzn_per_usd", 64),
      },
    });
  });

  // Manual lube log entry
  // POST /api/stock/lube-log
  // Body: { asset_code, log_date, oil_type?, quantity }
  app.post("/lube-log", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "artisan", "operator"])) return;
    const body = req.body || {};
    const asset_code = String(body.asset_code || "").trim();
    const log_date =
      body.log_date != null && String(body.log_date).trim() !== ""
        ? String(body.log_date).trim()
        : new Date().toISOString().slice(0, 10);
    const oil_type = normalizeOilTypeInput(body.oil_type, null);
    const quantity = Number(body.quantity ?? 0);

    if (!asset_code) return reply.code(400).send({ error: "asset_code is required" });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(log_date)) {
      return reply.code(400).send({ error: "log_date must be YYYY-MM-DD" });
    }
    if (!Number.isFinite(quantity) || quantity <= 0) {
      return reply.code(400).send({ error: "quantity must be > 0" });
    }

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });

    const ins = db.prepare(`
      INSERT INTO oil_logs (asset_id, log_date, oil_type, quantity)
      VALUES (?, ?, ?, ?)
    `).run(asset.id, log_date, oil_type, quantity);

    writeAudit(db, req, {
      module: "lube",
      action: "manual_log",
      entity_type: "asset",
      entity_id: asset_code,
      payload: { log_date, oil_type, quantity },
    });

    return reply.send({
      ok: true,
      id: Number(ins.lastInsertRowid),
      asset_code,
      log_date,
      oil_type,
      quantity,
    });
  });

  // Lube issue from stock (supports stock-number flow)
  // POST /api/stock/lube-issue
  // Body: { part_code, quantity, asset_code?, log_date?, oil_type?, notes?, location_code? }
  app.post("/lube-issue", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores", "artisan", "operator"])) return;
    const body = req.body || {};
    const part_code = String(body.part_code || "").trim();
    const asset_code = String(body.asset_code || "").trim();
    const quantity = Number(body.quantity ?? 0);
    const log_date =
      body.log_date != null && String(body.log_date).trim() !== ""
        ? String(body.log_date).trim()
        : new Date().toISOString().slice(0, 10);
    const oil_type = normalizeOilTypeInput(body.oil_type, null);
    const notes =
      body.notes != null && String(body.notes).trim() !== ""
        ? String(body.notes).trim()
        : null;
    const location_code = String(body.location_code || "").trim().toUpperCase();
    const bin_code = String(body.bin_code || "").trim().toUpperCase();

    if (!part_code) return reply.code(400).send({ error: "part_code is required" });
    if (!Number.isFinite(quantity) || quantity <= 0) {
      return reply.code(400).send({ error: "quantity must be > 0" });
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(log_date)) {
      return reply.code(400).send({ error: "log_date must be YYYY-MM-DD" });
    }

    const part = getPartByCode.get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const location = location_code ? getLocationByCode.get(location_code) : null;
    if (location_code && !location) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    const bin = (location && bin_code) ? getBinByCodeAtLocation.get(Number(location.id), bin_code) : null;
    if (bin_code && !location_code) return reply.code(400).send({ error: "location_code is required when bin_code is provided" });
    if (location && bin_code && !bin) return reply.code(404).send({ error: `bin_code not found at location ${location_code}: ${bin_code}` });

    let asset = null;
    if (asset_code) {
      asset = getAssetByCode.get(asset_code);
      if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });
    }

    const onHand = Number(getOnHand.get(part.id)?.on_hand || 0);
    if (onHand < quantity) {
      return reply.code(409).send({
        error: "insufficient stock",
        part_code,
        on_hand: onHand,
        requested: quantity,
      });
    }

    const tx = db.transaction(() => {
      const reference = asset ? `lube_issue:asset:${asset.id}` : `lube_issue:stock`;
      const mv = db.prepare(`
        INSERT INTO stock_movements (part_id, quantity, movement_type, reference, location_id, bin_id)
        VALUES (?, ?, 'out', ?, ?, ?)
      `).run(part.id, -Math.abs(quantity), reference, location ? Number(location.id) : null, bin ? Number(bin.id) : null);

      let lube_log_id = null;
      if (asset) {
        const lg = db.prepare(`
          INSERT INTO oil_logs (asset_id, log_date, oil_type, quantity)
          VALUES (?, ?, ?, ?)
        `).run(asset.id, log_date, normalizeOilTypeInput(oil_type, part.part_code || null), quantity);
        lube_log_id = Number(lg.lastInsertRowid);
      }

      return {
        movement_id: Number(mv.lastInsertRowid),
        lube_log_id,
      };
    });

    const result = tx();
    const on_hand_after = Number(getOnHand.get(part.id)?.on_hand || 0);

    writeAudit(db, req, {
      module: "lube",
      action: "stock_issue",
      entity_type: "part",
      entity_id: part_code,
      payload: {
        asset_code: asset ? asset.asset_code : null,
        quantity,
        log_date,
        oil_type,
        notes,
        location_code: location ? location.location_code : null,
        on_hand_before: onHand,
        on_hand_after,
      },
    });

    return reply.send({
      ok: true,
      part_code,
      part_name: part.part_name,
      quantity,
      movement_id: result.movement_id,
      lube_log_id: result.lube_log_id,
      linked_asset_code: asset ? asset.asset_code : null,
      linked_asset_name: asset ? asset.asset_name : null,
      location_code: location ? location.location_code : null,
      bin_code: bin ? bin.bin_code : null,
      on_hand_before: onHand,
      on_hand_after,
      message: asset
        ? "Lube issued from stock and linked to asset log"
        : "Lube issued from stock (not linked to asset log)",
    });
  });

  // Allocate stores/parts to an asset or work order
  // POST /api/stock/allocate
  // Body: { part_code, quantity, asset_code?, work_order_id?, allocation_date?, issued_by?, notes?, location_code? }
  app.post("/allocate", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const body = req.body || {};
    const part_code = String(body.part_code || "").trim();
    const quantity = Number(body.quantity ?? 0);
    const asset_code = String(body.asset_code || "").trim();
    const work_order_id = body.work_order_id != null ? Number(body.work_order_id) : null;
    const allocation_date =
      body.allocation_date != null && String(body.allocation_date).trim() !== ""
        ? String(body.allocation_date).trim()
        : new Date().toISOString().slice(0, 10);
    const issued_by =
      body.issued_by != null && String(body.issued_by).trim() !== ""
        ? String(body.issued_by).trim()
        : null;
    const notes =
      body.notes != null && String(body.notes).trim() !== ""
        ? String(body.notes).trim()
        : null;
    const location_code = String(body.location_code || "").trim().toUpperCase();
    const bin_code = String(body.bin_code || "").trim().toUpperCase();
    const cost_center_code =
      body.cost_center_code != null && String(body.cost_center_code).trim() !== ""
        ? String(body.cost_center_code).trim()
        : null;

    if (!part_code || !Number.isFinite(quantity) || quantity <= 0) {
      return reply.code(400).send({ error: "part_code and quantity (>0) are required" });
    }
    if (!asset_code && (!Number.isFinite(work_order_id) || work_order_id <= 0)) {
      return reply.code(400).send({ error: "Provide asset_code or work_order_id" });
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(allocation_date)) {
      return reply.code(400).send({ error: "allocation_date must be YYYY-MM-DD" });
    }

    const part = getPartByCode.get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const location = location_code ? getLocationByCode.get(location_code) : null;
    if (location_code && !location) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    const bin = (location && bin_code) ? getBinByCodeAtLocation.get(Number(location.id), bin_code) : null;
    if (bin_code && !location_code) return reply.code(400).send({ error: "location_code is required when bin_code is provided" });
    if (location && bin_code && !bin) return reply.code(404).send({ error: `bin_code not found at location ${location_code}: ${bin_code}` });

    let wo = null;
    if (Number.isFinite(work_order_id) && work_order_id > 0) {
      wo = getWoById.get(work_order_id);
      if (!wo) return reply.code(404).send({ error: `work_order not found: ${work_order_id}` });
    }

    let asset = null;
    if (asset_code) {
      asset = getAssetByCode.get(asset_code);
      if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });
    }

    const resolvedAssetId = wo ? Number(wo.asset_id) : Number(asset.id);
    if (wo && asset && Number(asset.id) !== Number(wo.asset_id)) {
      return reply.code(409).send({ error: "asset_code does not match work_order asset" });
    }

    const onHand = Number(getOnHand.get(part.id)?.on_hand || 0);
    if (onHand < quantity) {
      return reply.code(409).send({
        error: "insufficient stock",
        part_code,
        on_hand: onHand,
        requested: quantity
      });
    }

    const tx = db.transaction(() => {
      const reference = wo ? `work_order:${wo.id}` : `asset:${resolvedAssetId}:stores`;
      insertMove.run(
        part.id,
        -Math.abs(quantity),
        reference,
        location ? Number(location.id) : null,
        bin ? Number(bin.id) : null,
        cost_center_code
      );
      const ins = insertAlloc.run(
        resolvedAssetId,
        wo ? Number(wo.id) : null,
        part.id,
        quantity,
        allocation_date,
        issued_by,
        notes,
        location ? Number(location.id) : null,
        bin ? Number(bin.id) : null,
        cost_center_code
      );
      return Number(ins.lastInsertRowid);
    });

    const allocation_id = tx();

    writeAudit(db, req, {
      module: "stock",
      action: "allocate",
      entity_type: "work_order",
      entity_id: wo ? Number(wo.id) : resolvedAssetId,
      payload: {
        part_code,
        quantity,
        asset_id: resolvedAssetId,
        work_order_id: wo ? Number(wo.id) : null,
        location_code: location ? location.location_code : null,
        bin_code: bin ? bin.bin_code : null,
        cost_center_code,
      },
    });

    return reply.send({
      ok: true,
      allocation_id,
      asset_id: resolvedAssetId,
      work_order_id: wo ? Number(wo.id) : null,
      part_code,
      unit_cost_usd: Number(part.unit_cost || 0),
      line_value_usd: Number((Number(part.unit_cost || 0) * Number(quantity || 0)).toFixed(2)),
      issued: quantity,
      location_code: location ? location.location_code : null,
      bin_code: bin ? bin.bin_code : null,
      cost_center_code,
      on_hand_before: onHand,
      on_hand_after: onHand - quantity
    });
  });

  // List allocations
  // GET /api/stock/allocations?asset_code=&part_code=&start=&end=
  app.get("/allocations", async (req, reply) => {
    const asset_code = String(req.query?.asset_code || "").trim();
    const part_code = String(req.query?.part_code || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();

    const where = [];
    const params = [];
    if (asset_code) {
      where.push("a.asset_code = ?");
      params.push(asset_code);
    }
    if (part_code) {
      where.push("p.part_code = ?");
      params.push(part_code);
    }
    if (start && /^\d{4}-\d{2}-\d{2}$/.test(start)) {
      where.push("sa.allocation_date >= ?");
      params.push(start);
    }
    if (end && /^\d{4}-\d{2}-\d{2}$/.test(end)) {
      where.push("sa.allocation_date <= ?");
      params.push(end);
    }

    const rows = db.prepare(`
      SELECT
        sa.id,
        sa.allocation_date,
        a.asset_code,
        a.asset_name,
        sa.work_order_id,
        p.part_code,
        p.part_name,
        p.unit_cost,
        l.location_code,
        l.location_name,
        b.bin_code,
        b.bin_name,
        sa.cost_center_code,
        sa.quantity,
        sa.issued_by,
        sa.notes,
        sa.created_at
      FROM store_allocations sa
      JOIN assets a ON a.id = sa.asset_id
      JOIN parts p ON p.id = sa.part_id
      LEFT JOIN stock_locations l ON l.id = sa.location_id
      LEFT JOIN stock_bins b ON b.id = sa.bin_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY sa.id DESC
      LIMIT 500
    `).all(...params);

    return reply.send({ ok: true, rows });
  });
}