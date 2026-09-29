import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import {
  autoCategorizePart,
  classifyStockItem,
  ensureStockCategorySchema,
  normalizeStockCategory,
  oilPartSql,
} from "../utils/stockCategory.js";

const expect = (category, items) => {
  for (const [part_code, part_name] of items) {
    assert.equal(classifyStockItem({ part_code, part_name }), category, `${part_code} ${part_name}`);
  }
};

// Item names taken from a real IRONLOG monthly stock report.
test("oils are recognised by product, not by the letters 'oil' in a name", () => {
  expect("oil", [
    ["MLFF90906", "Fuchs Renolin HO46"],
    ["MLFPT2701", "Fuchs Titan Cargo MC 10W40 210lt"],
    ["MLFGT4601", "Fuchs Titan Supergear 80w90 210lt"],
    ["MLFU00150", "Fuchs Super Lupex M2 EP 180kg"],
    ["X-15W40", "Engine oil 15W40"],
    ["GR-1", "Lithium grease EP2"],
  ]);
  expect("part", [
    ["224363", "Coil directional vale B30"],
    ["BN003468", "SEAL OIL DOUBLE LIP BELL B30E"],
    ["OF-1", "Oil filter"],
    ["4SH-20", "11/4\" multi spiral hose"],
    ["87611-20-20", "11/4 SAE flange tail"],
  ]);
});

test("G.E.T, tyres and major components are separated from parts", () => {
  expect("get", [
    ["1U3302PTA", "Tip Cat J300"],
    ["8E6259", "Retainer Cat J250 J300"],
    ["XR50-82", "Ripper tip"],
    ["5D9559-BOR", "Grader Blade 7FT Curved Boron"],
  ]);
  expect("tyre", [
    ["BOTO 23.5R25", "BOTO GCB5 23.5R25 Tyre"],
    ["PTT/6PO2854W1", "Pneu Powertrac Wildranger 265/65R17 A/T 112S"],
  ]);
  expect("component", [
    ["3306B 10Z50252", "Cat 3306B engine"],
    ["DC221587", "Turbo OM906 with wastgate"],
    ["150044", "FEEDER RAM (TEREX MDS M515) CYLINDER"],
    ["1W5039", "Hub (CR05AM)"],
    ["MA940 500 04 03", "Radiator"],
    ["DC221758", "COMPRESSOR A/CON LESS COUPLING BELL"],
  ]);
  expect("part", [
    ["140438", "KIT SEAL STEERING CYL - BELL B30D"],
    ["202621", "PLT SEAL B30D 16T HUB"],
    ["CAT-950-500HR", "500hr service kit"],
    ["150137", "Bearing assy B30"],
  ]);
});

test("category names from CSV or forms are normalised", () => {
  assert.equal(normalizeStockCategory("Oils & Lubricants"), "oil");
  assert.equal(normalizeStockCategory("G.E.T"), "get");
  assert.equal(normalizeStockCategory("Tyres"), "tyre");
  assert.equal(normalizeStockCategory("components"), "component");
  assert.equal(normalizeStockCategory("widgets"), null);
});

test("automatic categories are filled and refreshed, hand-picked ones are kept", () => {
  const db = new Database(":memory:");
  db.exec(`CREATE TABLE parts (id INTEGER PRIMARY KEY, part_code TEXT, part_name TEXT)`);
  db.exec(`INSERT INTO parts (part_code, part_name) VALUES ('MLFPT6001', 'Fuchs Truck Plus 15W40'), ('224363', 'Coil directional vale B30'), ('X1', 'Mystery item')`);
  ensureStockCategorySchema(db);
  const cat = (code) => db.prepare(`SELECT stock_category, stock_category_source FROM parts WHERE part_code = ?`).get(code);
  assert.deepEqual(cat("MLFPT6001"), { stock_category: "oil", stock_category_source: "auto" });
  assert.equal(cat("224363").stock_category, "part");

  db.prepare(`UPDATE parts SET stock_category = 'component', stock_category_source = 'manual' WHERE part_code = 'X1'`).run();
  db.prepare(`UPDATE parts SET part_name = 'Tip Cat J300' WHERE part_code = '224363'`).run();
  ensureStockCategorySchema(db, { refresh: true });
  assert.equal(cat("224363").stock_category, "get", "automatic rows follow a renamed item");
  assert.deepEqual(cat("X1"), { stock_category: "component", stock_category_source: "manual" });

  const id = db.prepare(`INSERT INTO parts (part_code, part_name) VALUES ('T1', '23.5R25 tyre')`).run().lastInsertRowid;
  autoCategorizePart(db, id);
  assert.equal(cat("T1").stock_category, "tyre");

  const oils = db.prepare(`SELECT part_code FROM parts p WHERE ${oilPartSql("p")} ORDER BY part_code`).all().map((r) => r.part_code);
  assert.deepEqual(oils, ["MLFPT6001"]);
});
