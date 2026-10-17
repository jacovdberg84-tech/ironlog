import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import Fastify from "fastify";
import { componentReliability, groupReliabilityRows } from "../utils/reliabilityMetrics.js";

// Asset structure part 4: reliability by machine, model, area, class and failed component,
// with availability and maintenance cost per operating hour.
test("machines add up per group: MTBF, LTTR, availability and cost per hour", () => {
  const rows = [
    { make: "CAT", model: "777", area: "Pit", failure_count: 2, operating_hours: 400, downtime_hours: 20, scheduled_hours: 500, available_hours: 480, maintenance_cost: 4000 },
    { make: "CAT", model: "777", area: "Pit", failure_count: 0, operating_hours: 300, downtime_hours: 0, scheduled_hours: 500, available_hours: 500, maintenance_cost: 1000 },
    { make: "", model: "", area: null, failure_count: 1, operating_hours: 100, downtime_hours: 10, scheduled_hours: 200, available_hours: 190, maintenance_cost: null },
  ];
  const byModel = groupReliabilityRows(rows, (r) => [r.make, r.model].filter(Boolean).join(" "));
  assert.deepEqual(byModel[0], {
    key: "CAT 777", asset_count: 2, failure_count: 2, operating_hours: 700, downtime_hours: 20,
    mtbf_hours: 350, lttr_hours: 10, availability_pct: 98, maintenance_cost: 5000, cost_per_hour: 7.14,
  });
  assert.equal(byModel[1].key, "Not set", "machines without a model are grouped, not dropped");
  assert.equal(byModel[1].cost_per_hour, null, "no cost known, no cost per hour");

  const comp = componentReliability([{ fc_component: "HYD", downtime_hours: 6 }, { fc_component: "HYD", downtime_hours: 4 }, { fc_component: null, downtime_hours: 3 }], 800, { HYD: "Hydraulics" });
  assert.deepEqual(comp.map((c) => [c.label, c.failure_count, c.mtbf_hours, c.lttr_hours]), [["Hydraulics", 2, 400, 5], ["No failure code yet", 1, 800, 3]]);
});

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-relgroups-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";

test("the reliability report keeps to the site and groups by model, area and failed component", async (t) => {
  await import("../db/migrate.js");
  (await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
  const { db } = await import("../db/client.js");
  const { ironlogAuthHook } = await import("../auth/hook.js");
  const realSetInterval = globalThis.setInterval;
  globalThis.setInterval = (...args) => realSetInterval(...args).unref();
  const app = Fastify({ logger: false });
  app.addHook("onRequest", ironlogAuthHook);
  app.decorateRequest("ironlogAllowedSites", null);
  await app.register((await import("../routes/auth.routes.js")).default, { prefix: "/api/auth" });
  await app.register((await import("../routes/assets.routes.js")).default, { prefix: "/api/assets" });
  await app.register((await import("../routes/breakdowns.routes.js")).default, { prefix: "/api/breakdowns" });
  await app.register((await import("../routes/dashboard.routes.js")).default, { prefix: "/api/dashboard" });
  await app.register((await import("../routes/maintenance.routes.js")).default, { prefix: "/api/maintenance" });
  await app.ready();
  globalThis.setInterval = realSetInterval;
  t.after(async () => { await app.close(); try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

  // Added on the live database by the work order routes, which this test does not load.
  if (!db.prepare(`PRAGMA table_info(work_orders)`).all().some((c) => c.name === "completed_at")) db.prepare(`ALTER TABLE work_orders ADD COLUMN completed_at TEXT`).run();
  const asset = (code, site, make, model, area) => {
    const id = db.prepare(`INSERT INTO assets (asset_code, asset_name, category, active, site_code, make, model, area) VALUES (?, ?, 'Truck', 1, ?, ?, ?, ?)`)
      .run(code, code, site, make, model, area).lastInsertRowid;
    for (let d = 1; d <= 10; d++) {
      // On the live database the machine's site is copied onto its daily hours (utils/sites.js).
      db.prepare(`INSERT INTO daily_hours (asset_id, work_date, scheduled_hours, opening_hours, closing_hours, hours_run, is_used, site_code) VALUES (?, ?, 10, ?, ?, 8, 1, ?)`)
        .run(id, `2026-09-${String(d).padStart(2, "0")}`, d * 8, d * 8 + 8, site);
    }
    return id;
  };
  const t1 = asset("T01", "north", "CAT", "777", "Pit");
  asset("T02", "north", "CAT", "777", "Pit");
  const s1 = asset("S01", "south", "CAT", "777", "Pit");
  const bd = (assetId, comp) => db.prepare(`INSERT INTO breakdowns (asset_id, breakdown_date, status, description, downtime_total_hours, fc_component) VALUES (?, '2026-09-05', 'CLOSED', 'Hose', 6, ?)`).run(assetId, comp);
  bd(t1, "HYD");
  bd(s1, "ENG");
  db.prepare(`INSERT INTO work_orders (asset_id, source, status, labor_hours, labor_rate_per_hour, opened_at, completed_at, closed_at) VALUES (?, 'service', 'completed', 4, 50, '2026-09-03', '2026-09-03', '2026-09-03')`).run(t1);

  db.prepare(`INSERT INTO users (username, role, roles_json, active) VALUES ('boss', 'admin', '["admin"]', 1)`).run();
  const token = crypto.randomBytes(16).toString("hex");
  db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, (SELECT id FROM users WHERE username = 'boss'), datetime('now', '+1 day'))`).run(token);
  const res = await app.inject({ method: "GET", url: "/api/maintenance/reliability?start=2026-09-01&end=2026-09-10", headers: { authorization: `Bearer ${token}`, "x-site-code": "north" } });
  assert.equal(res.statusCode, 200, res.body);
  const data = res.json();
  assert.deepEqual(data.by_asset.map((r) => r.asset_code).sort(), ["T01", "T02"], "the south machine is not in the north report");
  assert.equal(data.summary.failure_count, 1);
  const t01 = data.by_asset.find((r) => r.asset_code === "T01");
  assert.deepEqual([t01.area, t01.make, t01.model, t01.maintenance_cost], ["Pit", "CAT", "777", 200], "4 h × $50 service labour");
  assert.ok(t01.cost_per_hour > 0);
  assert.equal(data.groups.model[0].key, "CAT 777");
  assert.equal(data.groups.model[0].asset_count, 2);
  assert.equal(data.groups.area[0].key, "Pit");
  assert.deepEqual(data.groups.component.map((c) => [c.code, c.label, c.failure_count]), [["HYD", "Hydraulics", 1]]);
  assert.ok(data.summary.availability_pct > 0 && data.summary.availability_pct <= 100, `availability ${data.summary.availability_pct}`);
});
