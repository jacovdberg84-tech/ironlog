// IRONLOG/api/routes/stock.routes.js
import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";
import { ensureMasterDataSchema } from "../utils/masterdataGovernance.js";
import { ensureCostAllocationSchema } from "../utils/costAllocation.js";
import registerInventoryRoutes from "./stock/inventory.routes.js";
import registerLocationsRoutes from "./stock/locations.routes.js";
import registerCycleCountsRoutes from "./stock/cycle-counts.routes.js";
import registerLubeRoutes from "./stock/lube.routes.js";
import registerMovementsRoutes from "./stock/movements.routes.js";
import registerPartOrdersRoutes from "./stock/part-orders.routes.js";
import registerRequisitionRoutes from "./stock/requisitions.routes.js";
import registerDeliveryRoutes from "./stock/deliveries.routes.js";
import { autoCategorizePart, ensureStockCategorySchema } from "../utils/stockCategory.js";
import { holdsAnyRole } from "../utils/request.js";

export default async function stockRoutes(app) {
  ensureAuditTable(db);
  ensureStockCategorySchema(db);
  ensureMasterDataSchema();
  ensureCostAllocationSchema(db);

  function hasColumn(table, col) {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name) === col);
  }

  function hasTable(table) {
    return Boolean(db.prepare("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?").get(table));
  }

  function getRole(req) {
    return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
  }

  function requireRoles(req, reply, roles) {
    const role = getRole(req);
    if (!holdsAnyRole(req, roles)) {
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
  if (!hasColumn("oil_logs", "part_id")) {
    db.prepare(`ALTER TABLE oil_logs ADD COLUMN part_id INTEGER`).run();
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

  const PART_ORDER_STATUSES = new Set(["on_order", "in_transit", "arrived", "cancelled"]);
  const PART_ORDER_WRITE_ROLES = ["admin", "supervisor", "stores", "storeman", "procurement", "plant_manager", "site_manager"];

  db.prepare(`
    CREATE TABLE IF NOT EXISTS stores_part_orders (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL DEFAULT 'main',
      part_id INTEGER,
      part_code TEXT,
      part_name TEXT NOT NULL,
      qty REAL NOT NULL DEFAULT 1,
      unit_cost REAL NOT NULL DEFAULT 0,
      currency TEXT NOT NULL DEFAULT 'USD',
      supplier_name TEXT,
      po_number TEXT,
      order_date TEXT NOT NULL,
      expected_arrival_date TEXT,
      arrived_date TEXT,
      status TEXT NOT NULL DEFAULT 'on_order',
      notes TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE SET NULL
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_stores_part_orders_site_date ON stores_part_orders(site_code, order_date)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_stores_part_orders_status ON stores_part_orders(status)`).run();
  if (!hasColumn("stores_part_orders", "stock_movement_id")) {
    db.prepare(`ALTER TABLE stores_part_orders ADD COLUMN stock_movement_id INTEGER`).run();
  }
  if (!hasColumn("stores_part_orders", "requisition_number")) {
    db.prepare(`ALTER TABLE stores_part_orders ADD COLUMN requisition_number TEXT`).run();
  }
  if (!hasColumn("stores_part_orders", "invoice_number")) {
    db.prepare(`ALTER TABLE stores_part_orders ADD COLUMN invoice_number TEXT`).run();
  }
  if (!hasColumn("stores_part_orders", "current_location")) {
    db.prepare(`ALTER TABLE stores_part_orders ADD COLUMN current_location TEXT`).run();
  }
  for (const [name, ddl] of [
    ["asset_id", "asset_id INTEGER"],
    ["work_order_id", "work_order_id INTEGER"],
    ["breakdown_id", "breakdown_id INTEGER"],
    ["offsite_repair_id", "offsite_repair_id INTEGER"],
    ["responsible_person", "responsible_person TEXT"],
  ]) {
    if (!hasColumn("stores_part_orders", name)) db.prepare(`ALTER TABLE stores_part_orders ADD COLUMN ${ddl}`).run();
  }
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_stores_part_orders_asset ON stores_part_orders(asset_id)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_stores_part_orders_work_order ON stores_part_orders(work_order_id)`).run();

  function partOrderReceiveRef(orderId) {
    return `part_order:${Number(orderId)}`;
  }

  function receivePartOrderToStock(orderRow, req) {
    const orderId = Number(orderRow?.id || 0);
    if (!orderId) return { received: false, error: "invalid order id" };

    const existingMovementId = Number(orderRow?.stock_movement_id || 0);
    if (existingMovementId > 0) {
      return { received: false, already: true, movement_id: existingMovementId };
    }

    const ref = partOrderReceiveRef(orderId);
    const dup = db.prepare(`SELECT id FROM stock_movements WHERE reference = ? LIMIT 1`).get(ref);
    if (dup?.id) {
      const mid = Number(dup.id);
      db.prepare(`UPDATE stores_part_orders SET stock_movement_id = ? WHERE id = ?`).run(mid, orderId);
      return { received: false, already: true, movement_id: mid };
    }

    const part_code = String(orderRow?.part_code || "").trim();
    if (!part_code) {
      return {
        received: false,
        error: "Add a part code on this purchase line to receive it into store inventory.",
      };
    }

    let part = getPartByCode.get(part_code);
    if (!part) {
      const part_name = String(orderRow?.part_name || part_code).trim() || part_code;
      const uc = Math.max(0, Number(orderRow?.unit_cost || 0));
      try {
        autoCategorizePart(db, insertPart.run(part_code, part_name, uc, null, null, null).lastInsertRowid);
      } catch {
        // race — re-read
      }
      part = getPartByCode.get(part_code);
    }
    if (!part) return { received: false, error: `Could not create or find part ${part_code}` };

    const qty = Math.abs(Number(orderRow?.qty || 0));
    if (!Number.isFinite(qty) || qty <= 0) return { received: false, error: "invalid quantity on order line" };

    const unit_cost = Math.max(0, Number(orderRow?.unit_cost || 0));
    const currency = String(orderRow?.currency || "USD").trim().toUpperCase() || "USD";
    let unit_cost_usd = null;
    let cost_input = null;
    let cost_curr_out = null;
    if (unit_cost > 0) {
      cost_input = unit_cost;
      cost_curr_out = currency;
      unit_cost_usd = unitCostToUsd(unit_cost, currency);
      if (unit_cost_usd == null || !Number.isFinite(unit_cost_usd)) {
        unit_cost_usd = currency === "USD" ? unit_cost : null;
      }
    }

    const on_hand_before = Number(getOnHand.get(part.id)?.on_hand || 0);

    const receiveTx = db.transaction(() => {
      const ins = insertGenericMove.run(
        part.id,
        qty,
        "in",
        ref,
        null,
        null,
        null,
        unit_cost_usd != null ? Number(unit_cost_usd.toFixed(6)) : null,
        cost_curr_out,
        cost_input,
      );
      const movement_id = Number(ins.lastInsertRowid);
      if (unit_cost_usd != null) {
        updatePartUnitCostUsd.run(Number(unit_cost_usd.toFixed(6)), part.id);
      }
      db.prepare(`
        UPDATE stores_part_orders
        SET stock_movement_id = ?, part_id = ?, updated_at = datetime('now')
        WHERE id = ?
      `).run(movement_id, Number(part.id), orderId);
      return movement_id;
    });

    const movement_id = receiveTx();
    const on_hand_after = Number(getOnHand.get(part.id)?.on_hand || 0);

    writeAudit(db, req, {
      module: "stock",
      action: "part_order.receive",
      entity_type: "stores_part_order",
      entity_id: String(orderId),
      payload: {
        part_code,
        qty,
        movement_id,
        on_hand_before,
        on_hand_after,
        reference: ref,
      },
    });

    return {
      received: true,
      movement_id,
      part_code,
      qty,
      on_hand_before,
      on_hand_after,
    };
  }

  function isYmd(s) {
    return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
  }

  function partOrderLineTotal(row) {
    return Number((Number(row?.qty || 0) * Number(row?.unit_cost || 0)).toFixed(2));
  }

  function summarizePartOrders(rows) {
    const summary = {
      on_order: { count: 0, qty: 0, value: 0 },
      in_transit: { count: 0, qty: 0, value: 0 },
      arrived: { count: 0, qty: 0, value: 0 },
      cancelled: { count: 0, qty: 0, value: 0 },
      total_forecast: 0,
      total_arrived: 0,
      total_pending: 0,
    };
    for (const row of rows) {
      const status = String(row.status || "on_order").toLowerCase();
      const bucket = summary[status] || null;
      const qty = Number(row.qty || 0);
      const value = partOrderLineTotal(row);
      if (bucket) {
        bucket.count += 1;
        bucket.qty += qty;
        bucket.value += value;
      }
      if (status === "arrived") summary.total_arrived += value;
      if (status === "on_order" || status === "in_transit") summary.total_pending += value;
      if (status !== "cancelled") summary.total_forecast += value;
    }
    for (const key of Object.keys(summary)) {
      if (typeof summary[key] === "object" && summary[key] != null) {
        summary[key].qty = Number(summary[key].qty.toFixed(2));
        summary[key].value = Number(summary[key].value.toFixed(2));
      }
    }
    summary.total_forecast = Number(summary.total_forecast.toFixed(2));
    summary.total_arrived = Number(summary.total_arrived.toFixed(2));
    summary.total_pending = Number(summary.total_pending.toFixed(2));
    return summary;
  }

  function mapPartOrderRow(row) {
    const line_total = partOrderLineTotal(row);
    const stock_movement_id = row.stock_movement_id != null ? Number(row.stock_movement_id) : null;
    return {
      ...row,
      qty: Number(row.qty || 0),
      unit_cost: Number(row.unit_cost || 0),
      line_total,
      stock_movement_id,
      in_store_inventory: Boolean(stock_movement_id),
    };
  }

  function syncLinkedPartOrder(orderId) {
    const row = db.prepare(`SELECT * FROM stores_part_orders WHERE id = ?`).get(Number(orderId || 0));
    if (!row) return;
    let breakdownId = Number(row.breakdown_id || 0) || null;
    if (!breakdownId && row.work_order_id) {
      const wo = db.prepare(`SELECT source, reference_id FROM work_orders WHERE id = ?`).get(Number(row.work_order_id));
      if (String(wo?.source || "").toLowerCase() === "breakdown") breakdownId = Number(wo?.reference_id || 0) || null;
    }
    if (!breakdownId && row.offsite_repair_id) {
      const offsite = db.prepare(`SELECT breakdown_id FROM breakdown_offsite_repairs WHERE id = ?`).get(Number(row.offsite_repair_id));
      breakdownId = Number(offsite?.breakdown_id || 0) || null;
    }
    if (!breakdownId) return;
    const label = { on_order: "Ordered", in_transit: "In transit", arrived: "Received", cancelled: "Cancelled" }[String(row.status || "on_order").toLowerCase()] || "Ordered";
    db.prepare(`
      UPDATE breakdowns
      SET parts_ordered_date = COALESCE(parts_ordered_date, ?),
          parts_status = ?,
          parts_received_date = CASE WHEN ? = 'arrived' THEN COALESCE(?, date('now')) ELSE parts_received_date END,
          ets_repair_date = COALESCE(?, ets_repair_date)
      WHERE id = ?
    `).run(row.order_date, label, String(row.status || ""), row.arrived_date, row.expected_arrival_date, breakdownId);
  }

  function validatePartOrderLinks({ asset_id, work_order_id, breakdown_id, offsite_repair_id }) {
    const linked = [
      ["work order", work_order_id, "work_orders"],
      ["breakdown", breakdown_id, "breakdowns"],
      ["offsite repair", offsite_repair_id, "breakdown_offsite_repairs"],
    ];
    const linkedAssetIds = [];
    for (const [label, id, table] of linked) {
      if (!id) continue;
      const row = db.prepare(`SELECT asset_id FROM ${table} WHERE id = ?`).get(Number(id));
      if (!row) return { error: `${label} not found`, status: 404 };
      if (Number(row.asset_id || 0) > 0) linkedAssetIds.push(Number(row.asset_id));
    }
    const distinct = [...new Set(linkedAssetIds)];
    if (distinct.length > 1 || (asset_id && distinct[0] && Number(asset_id) !== distinct[0])) {
      return { error: "asset and linked records do not belong to the same machine", status: 400 };
    }
    return { asset_id: Number(asset_id || distinct[0] || 0) || null };
  }

  // ============================================================
  // STORE QR — single field terminal for storeman (scan → store-mobile.html)
  // ============================================================
  db.prepare(`
    CREATE TABLE IF NOT EXISTS stores_qr_profiles (
      site_code TEXT PRIMARY KEY,
      qr_payload TEXT NOT NULL,
      qr_text TEXT NOT NULL,
      generated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  const getStoredStoreQrProfile = db.prepare(`
    SELECT qr_payload, qr_text, generated_at
    FROM stores_qr_profiles
    WHERE LOWER(TRIM(site_code)) = ?
  `);
  const upsertStoreQrProfile = db.prepare(`
    INSERT INTO stores_qr_profiles (site_code, qr_payload, qr_text, generated_at)
    VALUES (?, ?, ?, datetime('now'))
    ON CONFLICT(site_code) DO UPDATE SET
      qr_payload = excluded.qr_payload,
      qr_text = excluded.qr_text,
      generated_at = datetime('now')
  `);

  function resolveWebOrigin(req) {
    const protoHeader = String(req.headers["x-forwarded-proto"] || "").split(",")[0].trim();
    const hostHeader = String(req.headers["x-forwarded-host"] || req.headers.host || "").split(",")[0].trim();
    const proto = protoHeader || "http";
    if (hostHeader) return `${proto}://${hostHeader}`;
    return "";
  }

  function buildStoreQrProfile(siteCode, req) {
    const site_code = String(siteCode || "main").trim().toLowerCase() || "main";
    const origin = resolveWebOrigin(req);
    const targetPath = `/web/store-mobile.html?site=${encodeURIComponent(site_code)}`;
    const scan_url = origin ? `${origin}${targetPath}` : targetPath;

    const inv = db.prepare(`
      SELECT
        COUNT(DISTINCT p.id) AS part_count,
        COALESCE(SUM(CASE WHEN oh.on_hand < p.min_stock THEN 1 ELSE 0 END), 0) AS below_min
      FROM parts p
      LEFT JOIN (
        SELECT part_id, IFNULL(SUM(quantity), 0) AS on_hand
        FROM stock_movements
        GROUP BY part_id
      ) oh ON oh.part_id = p.id
    `).get();

    const pending = db.prepare(`
      SELECT COUNT(*) AS c
      FROM stores_part_orders
      WHERE LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
        AND LOWER(COALESCE(status, 'on_order')) IN ('on_order', 'in_transit')
    `).get(site_code);

    const profile = {
      purpose: "stores_field_terminal",
      generated_at: new Date().toISOString(),
      site_code,
      scan_url,
      inventory: {
        part_count: Number(inv?.part_count || 0),
        below_min: Number(inv?.below_min || 0),
      },
      pending_arrivals: Number(pending?.c || 0),
    };

    const qrText = [
      `IRONLOG STORES — ${site_code.toUpperCase()}`,
      `Scan for store field terminal (inventory, receive, labels)`,
      `Scan URL: ${scan_url}`,
      `Parts in catalogue: ${profile.inventory.part_count}`,
      `Below minimum: ${profile.inventory.below_min}`,
      `Pending arrivals: ${profile.pending_arrivals}`,
    ].join("\n");

    return { profile, qrText };
  }

  // Route groups live in routes/stock/. They receive the shared helpers above through ctx.
  const ctx = {
    PART_ORDER_STATUSES,
    PART_ORDER_WRITE_ROLES,
    buildStoreQrProfile,
    getAssetByCode,
    getBinByCodeAtLocation,
    getFxRate,
    getLocationByCode,
    getOnHand,
    getPartByCode,
    getRole,
    getSiteCode,
    getStoredStoreQrProfile,
    getWoById,
    hasColumn,
    hasTable,
    insertAlloc,
    insertGenericMove,
    insertMove,
    insertPart,
    isYmd,
    mapPartOrderRow,
    normalizeOilTypeInput,
    receivePartOrderToStock,
    requireRoles,
    summarizePartOrders,
    syncLinkedPartOrder,
    unitCostToUsd,
    updatePartUnitCostUsd,
    upsertStoreQrProfile,
    validatePartOrderLinks,
  };
  registerInventoryRoutes(app, ctx);
  registerLocationsRoutes(app, ctx);
  registerCycleCountsRoutes(app, ctx);
  registerLubeRoutes(app, ctx);
  registerMovementsRoutes(app, ctx);
  registerPartOrdersRoutes(app, ctx);
  registerRequisitionRoutes(app, ctx);
  registerDeliveryRoutes(app, ctx);
}
