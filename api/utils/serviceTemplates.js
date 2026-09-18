/**
 * Reusable, costed maintenance-service templates.  Templates describe planned
 * work; stock_movements remain the sole source of truth for actual issues.
 */

export const SERVICE_TEMPLATE_ITEM_TYPES = new Set([
  "service_kit", "filter", "part", "consumable", "oil", "grease", "coolant", "other",
]);

const PRICE_REQUIRED = "price_required";

function hasColumn(db, table, column) {
  return db.prepare(`PRAGMA table_info(${table})`).all()
    .some((row) => String(row.name || "") === String(column || ""));
}

function assetTextExpression(db, candidates) {
  const available = candidates.filter((column) => hasColumn(db, "assets", column));
  if (!available.length) return "NULL";
  const values = available.map((column) => `NULLIF(TRIM(${column}), '')`);
  return values.length === 1 ? values[0] : `COALESCE(${values.join(", ")})`;
}

function text(value) {
  return String(value ?? "").trim();
}

function norm(value) {
  return text(value).toLowerCase();
}

function money(value) {
  return Number(Number(value || 0).toFixed(2));
}

function number(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function ensureServiceTemplateSchema(db) {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS service_templates (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      template_key TEXT NOT NULL,
      name TEXT NOT NULL,
      description TEXT,
      manufacturer TEXT,
      model TEXT,
      asset_category TEXT,
      service_interval_hours REAL NOT NULL,
      meter_unit TEXT NOT NULL DEFAULT 'hours',
      estimated_duration_hours REAL NOT NULL DEFAULT 0,
      default_labour_hours REAL NOT NULL DEFAULT 0,
      default_labour_rate REAL NOT NULL DEFAULT 0,
      active INTEGER NOT NULL DEFAULT 1,
      revision_number INTEGER NOT NULL DEFAULT 1,
      supersedes_template_id INTEGER,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(template_key, revision_number),
      FOREIGN KEY(supersedes_template_id) REFERENCES service_templates(id) ON DELETE SET NULL
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_service_templates_active_interval
    ON service_templates(active, meter_unit, service_interval_hours)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS service_template_items (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      service_template_id INTEGER NOT NULL,
      stock_item_id INTEGER,
      item_type TEXT NOT NULL DEFAULT 'part',
      description TEXT NOT NULL,
      quantity_required REAL NOT NULL DEFAULT 0,
      unit_of_measure TEXT NOT NULL DEFAULT 'ea',
      required INTEGER NOT NULL DEFAULT 1,
      allow_substitute INTEGER NOT NULL DEFAULT 0,
      notes TEXT,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY(service_template_id) REFERENCES service_templates(id) ON DELETE CASCADE,
      FOREIGN KEY(stock_item_id) REFERENCES parts(id) ON DELETE SET NULL
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_service_template_items_template
    ON service_template_items(service_template_id, sort_order, id)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS asset_service_template_assignments (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER,
      manufacturer TEXT,
      model TEXT,
      asset_category TEXT,
      service_template_id INTEGER NOT NULL,
      priority INTEGER NOT NULL DEFAULT 100,
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY(asset_id) REFERENCES assets(id) ON DELETE CASCADE,
      FOREIGN KEY(service_template_id) REFERENCES service_templates(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_template_assignments_match
    ON asset_service_template_assignments(active, asset_id, manufacturer, model, asset_category, priority)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS service_estimates (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      service_template_id INTEGER,
      maintenance_plan_id INTEGER NOT NULL,
      work_order_id INTEGER,
      meter_reading REAL NOT NULL DEFAULT 0,
      template_name TEXT,
      template_revision INTEGER,
      status TEXT NOT NULL DEFAULT 'draft',
      pricing_complete INTEGER NOT NULL DEFAULT 0,
      estimated_parts_cost REAL NOT NULL DEFAULT 0,
      estimated_oil_cost REAL NOT NULL DEFAULT 0,
      estimated_consumables_cost REAL NOT NULL DEFAULT 0,
      estimated_labour_cost REAL NOT NULL DEFAULT 0,
      estimated_total_cost REAL NOT NULL DEFAULT 0,
      labour_hours_original REAL NOT NULL DEFAULT 0,
      labour_rate_original REAL NOT NULL DEFAULT 0,
      labour_hours_override REAL,
      labour_rate_override REAL,
      override_reason TEXT,
      price_date TEXT,
      approved_by TEXT,
      approved_at TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      created_by TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY(asset_id) REFERENCES assets(id) ON DELETE RESTRICT,
      FOREIGN KEY(service_template_id) REFERENCES service_templates(id) ON DELETE SET NULL,
      FOREIGN KEY(maintenance_plan_id) REFERENCES maintenance_plans(id) ON DELETE RESTRICT,
      FOREIGN KEY(work_order_id) REFERENCES work_orders(id) ON DELETE SET NULL
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_service_estimates_planner
    ON service_estimates(asset_id, maintenance_plan_id, status, created_at DESC)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS service_estimate_items (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      service_estimate_id INTEGER NOT NULL,
      stock_item_id INTEGER,
      stock_part_code TEXT,
      description TEXT NOT NULL,
      item_type TEXT NOT NULL DEFAULT 'part',
      quantity_required REAL NOT NULL DEFAULT 0,
      unit_of_measure TEXT NOT NULL DEFAULT 'ea',
      unit_cost REAL,
      estimated_line_total REAL,
      stock_quantity_available REAL,
      shortage_quantity REAL,
      source_price_date TEXT,
      price_status TEXT NOT NULL DEFAULT 'priced',
      required INTEGER NOT NULL DEFAULT 1,
      allow_substitute INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY(service_estimate_id) REFERENCES service_estimates(id) ON DELETE CASCADE,
      FOREIGN KEY(stock_item_id) REFERENCES parts(id) ON DELETE SET NULL
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_service_estimate_items_estimate
    ON service_estimate_items(service_estimate_id, id)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS work_order_planned_materials (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      work_order_id INTEGER NOT NULL,
      service_estimate_item_id INTEGER,
      part_id INTEGER,
      part_code TEXT,
      description TEXT NOT NULL,
      item_type TEXT NOT NULL DEFAULT 'part',
      quantity_planned REAL NOT NULL DEFAULT 0,
      unit_of_measure TEXT NOT NULL DEFAULT 'ea',
      unit_cost_snapshot REAL,
      planned_line_total REAL,
      source_price_date TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY(work_order_id) REFERENCES work_orders(id) ON DELETE CASCADE,
      FOREIGN KEY(service_estimate_item_id) REFERENCES service_estimate_items(id) ON DELETE SET NULL,
      FOREIGN KEY(part_id) REFERENCES parts(id) ON DELETE SET NULL
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_wo_planned_materials_wo
    ON work_order_planned_materials(work_order_id, id)`).run();

  // Reservations deliberately do not create stock movements.  They reduce the
  // available-to-promise figure only while active and unissued.
  db.prepare(`
    CREATE TABLE IF NOT EXISTS stock_reservations (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      work_order_id INTEGER NOT NULL,
      part_id INTEGER NOT NULL,
      quantity_reserved REAL NOT NULL,
      quantity_issued REAL NOT NULL DEFAULT 0,
      status TEXT NOT NULL DEFAULT 'active',
      reserved_by TEXT,
      reserved_at TEXT NOT NULL DEFAULT (datetime('now')),
      released_at TEXT,
      notes TEXT,
      UNIQUE(work_order_id, part_id),
      FOREIGN KEY(work_order_id) REFERENCES work_orders(id) ON DELETE CASCADE,
      FOREIGN KEY(part_id) REFERENCES parts(id) ON DELETE RESTRICT
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_stock_reservations_active
    ON stock_reservations(part_id, status)`).run();

  const woColumns = [
    ["service_estimate_id", "service_estimate_id INTEGER"],
    ["service_template_id", "service_template_id INTEGER"],
    ["service_template_revision", "service_template_revision INTEGER"],
    ["planned_labor_hours", "planned_labor_hours REAL"],
    ["planned_labor_rate", "planned_labor_rate REAL"],
  ];
  for (const [column, definition] of woColumns) {
    if (!hasColumn(db, "work_orders", column)) {
      db.prepare(`ALTER TABLE work_orders ADD COLUMN ${definition}`).run();
    }
  }
}

function getAssetMatchData(db, assetId) {
  const makeExpression = assetTextExpression(db, ["make", "asset_make", "manufacturer", "brand"]);
  const modelExpression = assetTextExpression(db, ["model", "asset_model"]);
  return db.prepare(`
    SELECT id, asset_code, asset_name, category,
      ${makeExpression} AS manufacturer,
      ${modelExpression} AS model
    FROM assets WHERE id = ?
  `).get(assetId) || null;
}

/** Deterministic assignment resolver.  Ambiguity is always returned, never guessed. */
export function resolveServiceTemplate(db, { assetId, intervalHours, meterUnit = "hours" }) {
  const asset = getAssetMatchData(db, Number(assetId || 0));
  if (!asset) return { status: "missing", reason: "asset_not_found", asset: null, template: null };
  const interval = number(intervalHours);
  const unit = norm(meterUnit) === "km" ? "km" : "hours";
  const candidates = db.prepare(`
    SELECT a.id AS assignment_id, a.asset_id, a.manufacturer AS assignment_manufacturer,
      a.model AS assignment_model, a.asset_category AS assignment_category, a.priority,
      t.id, t.template_key, t.name, t.description, t.manufacturer, t.model, t.asset_category,
      t.service_interval_hours, t.meter_unit, t.estimated_duration_hours,
      t.default_labour_hours, t.default_labour_rate, t.revision_number
    FROM asset_service_template_assignments a
    JOIN service_templates t ON t.id = a.service_template_id
    WHERE a.active = 1 AND t.active = 1
      AND ABS(COALESCE(t.service_interval_hours, 0) - ?) < 0.0001
      AND LOWER(COALESCE(t.meter_unit, 'hours')) = ?
  `).all(interval, unit).map((row) => {
    let tier = 0;
    if (Number(row.asset_id || 0) === Number(asset.id)) tier = 3;
    else if (!row.asset_id && norm(row.assignment_manufacturer) && norm(row.assignment_model)
      && norm(row.assignment_manufacturer) === norm(asset.manufacturer)
      && norm(row.assignment_model) === norm(asset.model)) tier = 2;
    else if (!row.asset_id && !norm(row.assignment_manufacturer) && !norm(row.assignment_model)
      && norm(row.assignment_category) && norm(row.assignment_category) === norm(asset.category)) tier = 1;
    return { ...row, tier };
  }).filter((row) => row.tier > 0);

  if (!candidates.length) return { status: "missing", reason: "no_assignment", asset, template: null };
  const bestTier = Math.max(...candidates.map((row) => row.tier));
  const tierRows = candidates.filter((row) => row.tier === bestTier);
  const bestPriority = Math.max(...tierRows.map((row) => number(row.priority)));
  const best = tierRows.filter((row) => number(row.priority) === bestPriority);
  if (best.length !== 1) {
    return {
      status: "ambiguous",
      reason: "equally_valid_assignments",
      asset,
      candidates: best.map((row) => ({ id: row.id, name: row.name, revision_number: row.revision_number })),
      template: null,
    };
  }
  return { status: "matched", reason: null, asset, template: best[0] };
}

function latestPrice(db, partId) {
  if (!partId) return { unit_cost: null, source_price_date: null, source: PRICE_REQUIRED };
  const partCostExpression = hasColumn(db, "parts", "unit_cost") ? "unit_cost" : "NULL";
  const part = db.prepare(`SELECT id, part_code, part_name, ${partCostExpression} AS unit_cost FROM parts WHERE id = ?`).get(partId);
  if (!part) return { unit_cost: null, source_price_date: null, source: PRICE_REQUIRED };
  const hasReceiptPrice = hasColumn(db, "stock_movements", "unit_cost_usd")
    || hasColumn(db, "stock_movements", "cost_input");
  const priceExpression = hasColumn(db, "stock_movements", "unit_cost_usd")
    ? (hasColumn(db, "stock_movements", "cost_input")
      ? "COALESCE(unit_cost_usd, cost_input)"
      : "unit_cost_usd")
    : "cost_input";
  const dateExpression = hasColumn(db, "stock_movements", "created_at")
    ? "DATE(created_at)"
    : "NULL";
  const movement = hasReceiptPrice
    ? db.prepare(`
        SELECT ${priceExpression} AS unit_cost, ${dateExpression} AS price_date
        FROM stock_movements
        WHERE part_id = ? AND quantity > 0
          AND COALESCE(${priceExpression}, 0) > 0
        ORDER BY id DESC LIMIT 1
      `).get(partId)
    : null;
  if (number(movement?.unit_cost) > 0) {
    return { unit_cost: number(movement.unit_cost), source_price_date: movement.price_date || null, source: "latest_stock_receipt", part };
  }
  if (number(part.unit_cost) > 0) {
    return { unit_cost: number(part.unit_cost), source_price_date: null, source: "current_store_cost", part };
  }
  return { unit_cost: null, source_price_date: null, source: PRICE_REQUIRED, part };
}

function stockOnHand(db, partId) {
  if (!partId) return 0;
  const onHand = number(db.prepare(`SELECT COALESCE(SUM(quantity), 0) AS on_hand FROM stock_movements WHERE part_id = ?`).get(partId)?.on_hand);
  const reserved = db.prepare(`SELECT COALESCE(SUM(quantity_reserved - quantity_issued), 0) AS reserved
    FROM stock_reservations WHERE part_id = ? AND status = 'active'`).get(partId)?.reserved;
  return Math.max(0, onHand - number(reserved));
}

export function buildServiceEstimatePreview(db, { assetId, planId, meterReading, intervalHours, meterUnit = "hours", labourHours, labourRate }) {
  const resolution = resolveServiceTemplate(db, { assetId, intervalHours, meterUnit });
  if (resolution.status !== "matched") {
    return { resolution, estimate: null, items: [], warnings: [resolution.status === "ambiguous" ? "Template assignment needs review" : "No service template assigned"] };
  }
  const template = resolution.template;
  const originalHours = Math.max(0, number(template.default_labour_hours));
  const originalRate = Math.max(0, number(template.default_labour_rate));
  const useHours = labourHours == null ? originalHours : Math.max(0, number(labourHours));
  const useRate = labourRate == null ? originalRate : Math.max(0, number(labourRate));
  const rows = db.prepare(`
    SELECT id, stock_item_id, item_type, description, quantity_required, unit_of_measure,
      required, allow_substitute, notes
    FROM service_template_items
    WHERE service_template_id = ? ORDER BY sort_order, id
  `).all(template.id);
  const items = rows.map((row) => {
    const price = latestPrice(db, Number(row.stock_item_id || 0));
    const requiredQty = Math.max(0, number(row.quantity_required));
    const available = stockOnHand(db, Number(row.stock_item_id || 0));
    const priceRequired = requiredQty > 0 && price.unit_cost == null;
    return {
      template_item_id: Number(row.id),
      stock_item_id: Number(row.stock_item_id || 0) || null,
      stock_part_code: price.part?.part_code || null,
      description: text(row.description) || text(price.part?.part_name) || "Unnamed material",
      item_type: SERVICE_TEMPLATE_ITEM_TYPES.has(norm(row.item_type)) ? norm(row.item_type) : "other",
      quantity_required: requiredQty,
      unit_of_measure: text(row.unit_of_measure) || "ea",
      required: Number(row.required) === 1,
      allow_substitute: Number(row.allow_substitute) === 1,
      unit_cost: price.unit_cost == null ? null : money(price.unit_cost),
      estimated_line_total: price.unit_cost == null ? null : money(requiredQty * price.unit_cost),
      stock_quantity_available: money(available),
      shortage_quantity: money(Math.max(0, requiredQty - available)),
      source_price_date: price.source_price_date,
      price_status: price.source,
    };
  });
  const categoryTotal = (types) => money(items.filter((item) => types.has(item.item_type))
    .reduce((sum, item) => sum + number(item.estimated_line_total), 0));
  const partsCost = categoryTotal(new Set(["service_kit", "filter", "part"]));
  const oilCost = categoryTotal(new Set(["oil", "grease", "coolant"]));
  const consumablesCost = categoryTotal(new Set(["consumable", "other"]));
  const labourCost = money(useHours * useRate);
  const missingPrices = items.filter((item) => item.required && item.quantity_required > 0 && item.unit_cost == null);
  const shortages = items.filter((item) => item.required && item.shortage_quantity > 0);
  return {
    resolution,
    items,
    warnings: [
      ...missingPrices.map((item) => `${item.description}: price required`),
      ...shortages.map((item) => `${item.description}: shortage ${item.shortage_quantity}`),
    ],
    estimate: {
      asset_id: Number(assetId),
      maintenance_plan_id: Number(planId),
      service_template_id: Number(template.id),
      meter_reading: number(meterReading),
      template_name: template.name,
      template_revision: Number(template.revision_number || 1),
      labour_hours_original: originalHours,
      labour_rate_original: originalRate,
      labour_hours: useHours,
      labour_rate: useRate,
      estimated_parts_cost: partsCost,
      estimated_oil_cost: oilCost,
      estimated_consumables_cost: consumablesCost,
      estimated_labour_cost: labourCost,
      estimated_total_cost: money(partsCost + oilCost + consumablesCost + labourCost),
      pricing_complete: missingPrices.length === 0,
      stock_available: shortages.length === 0,
    },
  };
}

export function releaseWorkOrderReservations(db, workOrderId) {
  db.prepare(`UPDATE stock_reservations
    SET status = 'released', released_at = datetime('now')
    WHERE work_order_id = ? AND status = 'active'`).run(workOrderId);
}

export function applyIssuedQuantityToReservation(db, workOrderId, partId, quantity) {
  const row = db.prepare(`SELECT id, quantity_reserved, quantity_issued
    FROM stock_reservations WHERE work_order_id = ? AND part_id = ? AND status = 'active'`).get(workOrderId, partId);
  if (!row) return;
  const nextIssued = Math.min(number(row.quantity_reserved), number(row.quantity_issued) + Math.max(0, number(quantity)));
  const status = nextIssued >= number(row.quantity_reserved) ? "consumed" : "active";
  db.prepare(`UPDATE stock_reservations SET quantity_issued = ?, status = ?, released_at = CASE WHEN ? = 'consumed' THEN datetime('now') ELSE released_at END WHERE id = ?`)
    .run(nextIssued, status, status, row.id);
}
