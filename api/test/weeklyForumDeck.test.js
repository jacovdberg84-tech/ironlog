import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";
import JSZip from "jszip";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-forum-deck-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
// Report schedulers start timers on import; keep them from holding the test open.
const realSetInterval = global.setInterval;
global.setInterval = (...args) => { const t = realSetInterval(...args); t.unref?.(); return t; };

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { isProductionAsset, issuedToEquipmentSql } = await import("../utils/weeklyForumFindings.js");
const app = Fastify({ logger: false });
for (const [file, prefix] of [
  ["../routes/hours.routes.js", "/api/hours"],
  ["../routes/dashboard.routes.js", "/api/dashboard"],
  ["../routes/maintenance.routes.js", "/api/maintenance"],
  ["../routes/workorders.routes.js", "/api/workorders"],
  ["../routes/breakdowns.routes.js", "/api/breakdowns"],
  ["../routes/breakdownOps.routes.js", "/api/breakdown-ops"],
  ["../routes/stock.routes.js", "/api/stock"],
  ["../routes/reports.routes.js", "/api/reports"],
]) {
  const { default: routes } = await import(file);
  await app.register(routes, { prefix });
}
await app.ready();
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

const cols = (t) => new Set(db.prepare(`PRAGMA table_info(${t})`).all().map((c) => c.name));
const has = (t, c) => cols(t).has(c);

test("production equipment excludes generators, welders, compressors and LDVs", () => {
  assert.equal(isProductionAsset({ asset_code: "A300AM", asset_name: "Bell B30D", category: "ADT" }), true);
  assert.equal(isProductionAsset({ asset_code: "EVR01AM", asset_name: "Everdigm drill rig", category: "Drill Rig" }), true);
  assert.equal(isProductionAsset({ asset_code: "WM01Z", asset_name: "Lincoln 305D Ranger", category: "Welding Machine" }), false);
  assert.equal(isProductionAsset({ asset_code: "CPR01AM", asset_name: "Ingersol Rand", category: "Compressor" }), false);
  assert.equal(isProductionAsset({ asset_code: "GS03AM", asset_name: "Perkins 100kVA", category: "Generator" }), false);
  assert.equal(isProductionAsset({ asset_code: "V07AM", asset_name: "Toyota Hilux", category: "LDV" }), false);
  assert.match(issuedToEquipmentSql("sm"), /work_order:%/);
});

test("weekly forum deck: content asked for by the forum", async (t) => {
  t.after(() => app.close());
  const wkStart = "2026-10-01";
  const wkEnd = "2026-10-08";
  db.exec(`
    INSERT INTO assets (id, asset_code, asset_name, category) VALUES
      (1, 'A300AM', 'Bell B30D', 'ADT'),
      (2, 'G01AM', 'CAT 140H', 'Grader'),
      (3, 'WM01Z', 'Lincoln 305D Ranger', 'Welding Machine'),
      (4, 'CPR01AM', 'Ingersol Rand', 'Compressor'),
      (5, 'V07AM', 'Toyota Hilux', 'LDV'),
      (6, 'EVR01AM', 'Everdigm drill rig', 'Drill Rig');
    INSERT INTO parts (id, part_code, part_name, unit_cost) VALUES (1, 'OIL-15W40', 'Rimula 15W40 20L', 100), (2, 'FLT-01', 'Oil filter', 20);
  `);
  const dh = db.prepare(`INSERT INTO daily_hours (asset_id, work_date, scheduled_hours, hours_run, is_used) VALUES (?, ?, ?, ?, 1)`);
  for (const [id, run] of [[1, 10], [2, 1], [3, 0.2], [4, 0.5], [6, 8]]) dh.run(id, "2026-10-02", 11, run);
  // Stock: a real issue to a work order, a lube issue, and a correction that removed double-booked oil.
  db.exec(`
    INSERT INTO stock_movements (part_id, quantity, movement_type, reference, created_at) VALUES
      (1, 50, 'in', 'INV-1', '2026-09-20 08:00:00'),
      (2, 20, 'in', 'INV-1', '2026-09-20 08:00:00'),
      (2, -2, 'out', 'work_order:10', '2026-10-03 09:00:00'),
      (1, -1, 'out', 'lube_issue:asset:1', '2026-10-03 10:00:00'),
      (1, -20, 'out', 'Remove oil captured twice', '2026-10-04 11:00:00'),
      (2, -1, 'out', 'work_order:10', '2026-09-10 09:00:00');
  `);
  db.exec(`
    INSERT INTO work_orders (id, asset_id, source, status, opened_at) VALUES
      (10, 1, 'manual', 'open', '2026-10-02'),
      (11, 5, 'breakdown', 'open', '2026-10-02'),
      (12, 2, 'breakdown', 'in_progress', '2026-10-03');
  `);
  // EVR01AM is in South Africa for repairs.
  const off = await app.inject({
    method: "POST",
    url: "/api/breakdown-ops/offsite-repairs",
    headers: { "x-user-name": "admin", "x-user-role": "admin", "x-user-roles": "admin", "x-site-code": "main" },
    payload: { asset_code: "EVR01AM", sent_date: "2026-09-15", vendor: "Everdigm SA", current_location: "Johannesburg, South Africa", repair_reason: "Rotation motor rebuild", repair_status: "in_repair", expected_return_date: "2026-10-30" },
  });
  if (off.statusCode >= 300) {
    // Fall back to a direct row if the route needs fields this test does not know about.
    db.prepare(`INSERT INTO breakdown_offsite_repairs (site_code, asset_id, repair_status, sent_date, expected_return_date, vendor, current_location, repair_reason) VALUES ('main', 6, 'in_repair', '2026-09-15', '2026-10-30', 'Everdigm SA', 'Johannesburg, South Africa', 'Rotation motor rebuild')`).run();
  }
  // Parts ordered in September and October.
  db.exec(`
    INSERT INTO stores_part_orders (site_code, part_name, qty, unit_cost, currency, order_date, status) VALUES
      ('main', 'Track chain', 2, 1500, 'USD', '2026-09-12', 'on_order'),
      ('main', 'Seal kit', 1, 18500, 'ZAR', '2026-09-20', 'arrived'),
      ('main', 'Cancelled line', 1, 9999, 'USD', '2026-09-21', 'cancelled'),
      ('main', 'Filters', 10, 25, 'USD', '2026-10-05', 'on_order');
  `);
  if (has("work_orders", "labor_hours")) db.exec(`UPDATE work_orders SET labor_hours = 0`);

  const gen = await app.inject({
    method: "POST",
    url: "/api/reports/maintenance-master/generate",
    headers: { "x-user-name": "admin", "x-user-role": "admin", "x-user-roles": "admin", "x-site-code": "main" },
    payload: { period_type: "weekly", start: wkStart, end: wkEnd, site_code: "main" },
  });
  assert.equal(gen.statusCode, 200, gen.body);
  if (process.env.FORUM_DECK_OUT) (await import("node:fs")).copyFileSync(gen.json().file_path, process.env.FORUM_DECK_OUT);
  const zip = await JSZip.loadAsync(readFileSync(gen.json().file_path));
  const names = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]));
  assert.equal(names.length, 11, "11 slides");
  const slides = await Promise.all(names.map(async (n) => [...(await zip.file(n).async("string")).matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join(" | ")));
  const all = slides.join("\n");

  // Slide 1: no reporting scope, no IronLog footer; black + yellow theme.
  assert.doesNotMatch(slides[0], /Reporting scope/);
  assert.doesNotMatch(all, /IRONLOG • Weekly Forum|IRONLOG • MAIN SITE/);
  assert.match(slides[0], /1 \/ 11/);
  assert.match(await zip.file(names[0]).async("string"), /FFCD11/);
  // Slide 2: no explanatory comment, no Action / owner column.
  assert.doesNotMatch(slides[1], /Live period figures|Action \/ owner/);
  assert.match(slides[1], /Weekly finding/);
  // Slides 4 and 5: production equipment only, no "Explain low use".
  for (const i of [3, 4]) {
    assert.doesNotMatch(slides[i], /WM01Z|CPR01AM|V07AM/, `slide ${i + 1} production only`);
  }
  assert.match(slides[4], /G01AM/);
  assert.doesNotMatch(slides[4], /Explain low use/);
  // Slide 6: no Return / next step.
  assert.doesNotMatch(slides[5], /Return \/ next step/);
  // Slide 7: production equipment only, no Owner / due.
  assert.doesNotMatch(slides[6], /Owner \/ due|V07AM/);
  assert.match(slides[6], /WO-12 • G01AM/);
  // Slide 8: equipment off site, EVR01AM in South Africa.
  assert.match(slides[7], /Equipment off site for repairs/);
  assert.match(slides[7], /EVR01AM/);
  assert.match(slides[7], /South Africa/);
  // Slide 9: parts issued counts issues only (2 × $20 + 1 × $100), not the oil correction.
  assert.match(slides[8], /Parts issued \| \$0 \| \$140/);
  // Slide 10: heading.
  assert.match(slides[9], /Upcoming maintenance/);
  assert.doesNotMatch(slides[9], /Next week maintenance plan/);
  // Slide 11: September costing next to October to date; parts ordered converts ZAR and skips cancelled.
  assert.match(slides[10], /September 2026 costing/);
  assert.match(slides[10], /October to date/);
  assert.match(slides[10], /Parts ordered \| \$4,000 \| \$250/);
  assert.match(slides[10], /Parts issued \| \$20 \| \$140/);
});
