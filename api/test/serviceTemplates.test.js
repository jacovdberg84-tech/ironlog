import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import {
  buildServiceEstimatePreview,
  ensureServiceTemplateSchema,
  resolveServiceTemplate,
} from "../utils/serviceTemplates.js";

function createDb() {
  const db = new Database(":memory:");
  db.pragma("foreign_keys = ON");
  db.exec(`
    CREATE TABLE assets (
      id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT, category TEXT,
      make TEXT, model TEXT
    );
    CREATE TABLE parts (
      id INTEGER PRIMARY KEY, part_code TEXT, part_name TEXT, unit_cost REAL
    );
    CREATE TABLE stock_movements (
      id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL, unit_cost_usd REAL,
      cost_input REAL, created_at TEXT
    );
    CREATE TABLE maintenance_plans (id INTEGER PRIMARY KEY, asset_id INTEGER);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER, status TEXT);
  `);
  ensureServiceTemplateSchema(db);
  return db;
}

function addTemplate(db, key, interval, { itemPartId = null, category = "Excavator" } = {}) {
  const templateId = Number(db.prepare(`
    INSERT INTO service_templates (
      template_key, name, asset_category, service_interval_hours, meter_unit,
      default_labour_hours, default_labour_rate
    ) VALUES (?, ?, ?, ?, 'hours', 2, 80)
  `).run(key, `${key} service`, category, interval).lastInsertRowid);
  db.prepare(`
    INSERT INTO service_template_items (
      service_template_id, stock_item_id, item_type, description, quantity_required, unit_of_measure
    ) VALUES (?, ?, 'filter', 'Engine filter', 2, 'ea')
  `).run(templateId, itemPartId);
  return templateId;
}

test("an exact asset assignment wins over a category template", () => {
  const db = createDb();
  db.prepare(`INSERT INTO assets VALUES (1, 'A300AM', 'Bell B30D', 'Excavator', 'Bell', 'B30D')`).run();
  const categoryTemplate = addTemplate(db, "EXC-500", 500);
  const exactTemplate = addTemplate(db, "A300-500", 500);
  db.prepare(`INSERT INTO asset_service_template_assignments (asset_category, service_template_id) VALUES ('Excavator', ?)`).run(categoryTemplate);
  db.prepare(`INSERT INTO asset_service_template_assignments (asset_id, service_template_id) VALUES (1, ?)`).run(exactTemplate);

  const result = resolveServiceTemplate(db, { assetId: 1, intervalHours: 500 });
  assert.equal(result.status, "matched");
  assert.equal(result.template.id, exactTemplate);
  db.close();
});

test("equally valid assignments remain ambiguous rather than being guessed", () => {
  const db = createDb();
  db.prepare(`INSERT INTO assets VALUES (1, 'A300AM', 'Bell B30D', 'Excavator', 'Bell', 'B30D')`).run();
  const one = addTemplate(db, "ONE-500", 500);
  const two = addTemplate(db, "TWO-500", 500);
  db.prepare(`INSERT INTO asset_service_template_assignments (asset_category, service_template_id) VALUES ('Excavator', ?)`).run(one);
  db.prepare(`INSERT INTO asset_service_template_assignments (asset_category, service_template_id) VALUES ('Excavator', ?)`).run(two);

  const result = resolveServiceTemplate(db, { assetId: 1, intervalHours: 500 });
  assert.equal(result.status, "ambiguous");
  assert.equal(result.candidates.length, 2);
  db.close();
});

test("a missing material price is never silently treated as zero", () => {
  const db = createDb();
  db.prepare(`INSERT INTO assets VALUES (1, 'A300AM', 'Bell B30D', 'Excavator', 'Bell', 'B30D')`).run();
  db.prepare(`INSERT INTO maintenance_plans VALUES (7, 1)`).run();
  db.prepare(`INSERT INTO parts VALUES (10, 'FLT-01', 'Engine filter', 0)`).run();
  const templateId = addTemplate(db, "A300-500", 500, { itemPartId: 10 });
  db.prepare(`INSERT INTO asset_service_template_assignments (asset_id, service_template_id) VALUES (1, ?)`).run(templateId);

  const preview = buildServiceEstimatePreview(db, {
    assetId: 1, planId: 7, meterReading: 1000, intervalHours: 500,
  });
  assert.equal(preview.estimate.pricing_complete, false);
  assert.equal(preview.items[0].unit_cost, null);
  assert.match(preview.warnings.join(" "), /price required/i);
  assert.equal(preview.estimate.estimated_total_cost, 160);
  db.close();
});
