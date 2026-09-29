import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import { applyServiceTemplateProposals, buildServiceTemplateProposals, matchKit, standardLabourHours } from "../utils/serviceTemplateBuilder.js";
import { buildServiceEstimatePreview, resolveServiceTemplate } from "../utils/serviceTemplates.js";

function seed() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT, category TEXT, active INTEGER DEFAULT 1, archived INTEGER DEFAULT 0);
    CREATE TABLE maintenance_plans (id INTEGER PRIMARY KEY, asset_id INTEGER, service_name TEXT, interval_hours REAL, last_service_hours REAL, active INTEGER DEFAULT 1);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER, status TEXT);
    CREATE TABLE parts (id INTEGER PRIMARY KEY, part_code TEXT UNIQUE, part_name TEXT, unit_cost REAL DEFAULT 0, min_stock REAL DEFAULT 0, critical INTEGER DEFAULT 0);
    CREATE TABLE stock_movements (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL, movement_type TEXT, reference TEXT, created_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE cost_settings (key TEXT PRIMARY KEY, value REAL);
    INSERT INTO cost_settings VALUES ('labor_cost_per_hour_default', 80);
    INSERT INTO assets (id, asset_code, asset_name, category) VALUES (1, 'E500AM', 'CAT 350', 'Excavator'), (2, 'A300AM', 'Bell B30D', 'ADT'), (3, 'X9', 'Mystery', 'Other');
    INSERT INTO maintenance_plans (id, asset_id, service_name, interval_hours, last_service_hours) VALUES
      (10, 1, '500', 500, 0), (11, 1, '1000 hour service', 1000, 0), (20, 2, '500', 500, 0), (30, 3, '250', 250, 0);
    INSERT INTO parts (id, part_code, part_name, unit_cost) VALUES
      (1, 'CAT350-500HR', 'CAT 350 500hr service kit', 828),
      (2, 'MLFPT2301', 'Fuchs Titan UTTO TO-4 SAE 50 210lt', 5.63),
      (3, 'MLFPT2701', 'Fuchs Titan Cargo MC 10W40 210lt', 6.13),
      (4, 'HOSE-1', 'Hydraulic hose', 40),
      (5, 'TEICH-A300AM-500HRS', 'Bell B30D 500hr service kit', 610),
      (6, 'TEICH-A300AM-1000HRS', 'Bell B30D 1000hr service kit', 900);
    -- Three past 500 h services on E500AM. The hose was a one-off extra.
    INSERT INTO work_orders (id, asset_id, source, reference_id, status) VALUES (281, 1, 'service', 10, 'closed'), (282, 1, 'service', 10, 'closed'), (283, 1, 'service', 10, 'closed');
    INSERT INTO stock_movements (part_id, quantity, movement_type, reference) VALUES
      (1, -1, 'out', 'work_order:281'), (2, -36, 'out', 'work_order:281'), (3, -32, 'out', 'work_order:281'),
      (1, -1, 'out', 'work_order:282'), (2, -34, 'out', 'work_order:282'), (3, -30, 'out', 'work_order:282'), (4, -1, 'out', 'work_order:282'),
      (1, -1, 'out', 'work_order:283'), (2, -40, 'out', 'work_order:283'), (3, -2, 'out', 'work_order:283'), (3, 1, 'return', 'work_order:283');
  `);
  return db;
}

test("standard labour grows with the service interval", () => {
  assert.equal(standardLabourHours(250), 2);
  assert.equal(standardLabourHours(500), 4);
  assert.equal(standardLabourHours(1000), 6);
  assert.equal(standardLabourHours(2000), 8);
  assert.equal(standardLabourHours(10000, "km"), 2);
});

test("kits are matched by machine code and exact interval", () => {
  const kits = [{ part_code: "TEICH-A300AM-500HRS", part_name: "" }, { part_code: "TEICH-A300AM-1000HRS", part_name: "" }, { part_code: "TEICH-E500AM 500HR", part_name: "" }];
  assert.equal(matchKit(kits, "A300AM", 500).part_code, "TEICH-A300AM-500HRS");
  assert.equal(matchKit(kits, "A300AM", 1000).part_code, "TEICH-A300AM-1000HRS");
  assert.equal(matchKit(kits, "E500AM", 500).part_code, "TEICH-E500AM 500HR");
  assert.equal(matchKit(kits, "E500AM", 1000), null);
});

test("proposals come from service history, then from coded kits", () => {
  const db = seed();
  const proposals = buildServiceTemplateProposals(db);
  const byKey = new Map(proposals.map((p) => [`${p.asset_code}:${p.interval}`, p]));

  const e500 = byKey.get("E500AM:500");
  assert.equal(e500.source, "history");
  assert.equal(e500.history_services, 3);
  assert.deepEqual(e500.items.map((i) => [i.part_code, i.item_type, i.quantity_required, i.unit_of_measure]), [
    ["CAT350-500HR", "service_kit", 1, "ea"],
    ["MLFPT2301", "oil", 36, "L"],
    ["MLFPT2701", "oil", 30, "L"],
  ], "median quantities; the one-off hose is left out; returns are netted off");
  assert.equal(e500.labour_hours, 4);
  assert.equal(e500.estimate.labour_cost, 320);
  assert.equal(e500.estimate.total, Math.round((828 + 36 * 5.63 + 30 * 6.13 + 320) * 100) / 100);

  const a300 = byKey.get("A300AM:500");
  assert.equal(a300.source, "kit_code");
  assert.deepEqual(a300.items.map((i) => i.part_code), ["TEICH-A300AM-500HRS"]);

  assert.equal(byKey.get("E500AM:1000").source, "none");
  assert.equal(byKey.get("X9:250").items.length, 0);
});

test("applied templates are assigned to the machine, priced, and replace earlier ones", () => {
  const db = seed();
  let proposals = buildServiceTemplateProposals(db);
  const created = applyServiceTemplateProposals(db, proposals, ["1:500", "2:500", "1:1000"]);
  assert.equal(created.length, 2, "proposals without items are skipped");

  const match = resolveServiceTemplate(db, { assetId: 1, intervalHours: 500 });
  assert.equal(match.status, "matched");
  const preview = buildServiceEstimatePreview(db, { assetId: 1, planId: 10, meterReading: 0, intervalHours: 500 });
  assert.equal(preview.estimate.estimated_labour_cost, 320);
  assert.equal(preview.estimate.estimated_total_cost, Math.round((828 + 36 * 5.63 + 30 * 6.13 + 320) * 100) / 100);

  // Rebuilding creates revision 2 and retires revision 1.
  proposals = buildServiceTemplateProposals(db);
  assert.equal(proposals.find((p) => p.key === "1:500").existing.asset_specific, true);
  const again = applyServiceTemplateProposals(db, proposals, ["1:500"]);
  assert.equal(again[0].revision, 2);
  const active = db.prepare(`SELECT COUNT(*) AS n FROM service_templates WHERE template_key = 'AUTO-E500AM-500' AND active = 1`).get().n;
  assert.equal(active, 1);
  assert.equal(resolveServiceTemplate(db, { assetId: 1, intervalHours: 500 }).template.id, again[0].id);
});

test("upcoming service costs use the machine's template instead of needing manual input", async () => {
  const { buildUpcomingServiceCostForecasts } = await import("../routes/maintenance.routes.js");
  const db = seed();
  applyServiceTemplateProposals(db, buildServiceTemplateProposals(db), ["2:500"]);
  const ctx = {
    hasTable: (name) => Boolean(db.prepare("SELECT name FROM sqlite_master WHERE type='table' AND name=?").get(name)),
    hasColumn: () => false,
    closedStatuses: "'closed'",
    woCloseExpr: "NULL",
    smOutSql: "1=1",
    oilPartSql: "0=1",
    smCostWithParts: "0",
    smCostNoParts: "0",
    lubeCostDefault: 4,
  };
  const [row] = buildUpcomingServiceCostForecasts(db, [{
    plan_id: 20, asset_id: 2, asset_code: "A300AM", asset_name: "Bell B30D", service_name: "500", last_service_hours: 0, interval_hours: 500,
  }], { nearDueHours: 50, horizonHours: 1000, getAssetHours: () => 480, ctx });
  assert.equal(row.forecast.cost_source, "service_template");
  assert.equal(row.needs_manual_input, false);
  assert.equal(row.forecast.est_service_kit_cost, 610);
  assert.equal(row.forecast.est_labor_cost, 320);
  assert.equal(row.forecast.est_total_cost, 930);
  assert.equal(row.forecast.template.name, "A300AM 500 h service");
});

test("model kits cover machines without their own coded kit", async () => {
  const { matchModelKit } = await import("../utils/serviceTemplateBuilder.js");
  const kits = [
    { part_code: "CAT350-500HR", part_name: "CAT 350 500hr service kit" },
    { part_code: "CAT-950-500HR", part_name: "500hr service kit" },
    { part_code: "TEICH-E501AM 1000HR", part_name: "Cat 350 1000hr service kit" },
  ];
  assert.equal(matchModelKit(kits, "E502AM", "CAT 350", 500).part_code, "CAT350-500HR");
  assert.equal(matchModelKit(kits, "F500AM", "CAT 950GC", 500).part_code, "CAT-950-500HR");
  assert.equal(matchModelKit(kits, "E502AM", "CAT 350", 1000), null, "another machine's coded kit is not borrowed");
  assert.equal(matchModelKit(kits, "E501AM", "CAT 350", 1000).part_code, "TEICH-E501AM 1000HR");
});

test("kit-only proposals borrow oils from a same-model machine's history", () => {
  const db = seed();
  db.exec(`INSERT INTO assets (id, asset_code, asset_name, category) VALUES (4, 'E501AM', 'CAT 350', 'Excavator');
    INSERT INTO maintenance_plans (id, asset_id, service_name, interval_hours, last_service_hours) VALUES (40, 4, '500', 500, 0);
    INSERT INTO parts (id, part_code, part_name, unit_cost) VALUES (7, 'TEICH-E501AM-500HRS', 'Cat 350 500hr service kit', 700);`);
  const p = buildServiceTemplateProposals(db).find((x) => x.asset_code === "E501AM" && x.interval === 500);
  assert.equal(p.source, "kit_code");
  assert.equal(p.oils_from, "E500AM");
  assert.deepEqual(p.items.map((i) => [i.part_code, i.quantity_required]), [["TEICH-E501AM-500HRS", 1], ["MLFPT2301", 36], ["MLFPT2701", 30]]);
  assert.equal(p.estimate.total, Math.round((700 + 36 * 5.63 + 30 * 6.13 + 320) * 100) / 100);
});
