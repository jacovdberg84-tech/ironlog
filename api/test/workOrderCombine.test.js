import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Database from "better-sqlite3";
import Fastify from "fastify";
import { combinePrestartWorkOrders, previewCombine } from "../utils/workOrderCombine.js";
import { syncPrestartFaultWorkOrder } from "../utils/prestartFaults.js";

function db() {
  const d = new Database(":memory:");
  d.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER, status TEXT,
      site_code TEXT, job_description TEXT, repair_progress TEXT, assigned_artisan_name TEXT, labor_hours REAL,
      labor_rate_per_hour REAL, oil_cost REAL, opened_at TEXT, closed_at TEXT, completion_notes TEXT);
    CREATE TABLE maintenance_parts_requests (id INTEGER PRIMARY KEY, work_order_id INTEGER, part_name TEXT);
    CREATE TABLE work_order_technicians (id INTEGER PRIMARY KEY, work_order_id INTEGER, username TEXT, UNIQUE (work_order_id, username));
    CREATE TABLE stock_movements (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL, reference TEXT);
    CREATE TABLE breakdowns (id INTEGER PRIMARY KEY, primary_work_order_id INTEGER);
    INSERT INTO assets VALUES (1, 'E502AM', 'Excavator 502'), (2, 'E501AM', 'Excavator 501');
    INSERT INTO work_orders (id, asset_id, source, reference_id, status, job_description, repair_progress, assigned_artisan_name, labor_hours, labor_rate_per_hour, oil_cost, opened_at) VALUES
      (10, 1, 'prestart', 100, 'open', 'Pre-start fault reported by Ana on 2026-10-01:\n- Engine oil: low', NULL, NULL, NULL, NULL, NULL, '2026-10-01'),
      (11, 1, 'prestart', 101, 'in_progress', 'Pre-start fault reported by Ana on 2026-10-02:\n- Engine oil: low', 'Topped up, checking for leaks', 'jose', 2, 40, NULL, '2026-10-02'),
      (12, 1, 'prestart', 102, 'open', 'Pre-start faults reported by Joao on 2026-10-03:\n- Engine oil: low\n- Horn', NULL, NULL, 1.5, NULL, 12, '2026-10-03'),
      (13, 1, 'prestart', 103, 'completed', 'Old job', NULL, 'jose', NULL, NULL, NULL, '2026-09-20'),
      (14, 1, 'breakdown', 5, 'open', 'Boom cylinder', NULL, NULL, NULL, NULL, NULL, '2026-10-02'),
      (20, 2, 'prestart', 200, 'open', 'Other machine', NULL, NULL, NULL, NULL, NULL, '2026-10-02');
    INSERT INTO maintenance_parts_requests (work_order_id, part_name) VALUES (10, 'Oil filter'), (12, 'Horn relay'), (20, 'Not this machine');
    INSERT INTO work_order_technicians (work_order_id, username) VALUES (11, 'jose'), (12, 'jose'), (12, 'maria');
    INSERT INTO stock_movements (part_id, quantity, reference) VALUES (1, -5, 'work_order:10'), (2, -1, 'work_order:12'), (3, -2, 'work_order:120');
    INSERT INTO breakdowns VALUES (1, 12);
  `);
  return d;
}

test("preview keeps the job furthest along and lists only this machine's open pre-start jobs", () => {
  const p = previewCombine(db(), 10);
  assert.equal(p.ok, true);
  assert.equal(p.asset_code, "E502AM");
  assert.equal(p.keep.id, 11, "in progress with a technician beats older open jobs");
  assert.deepEqual(p.combine.map((w) => w.id), [10, 12]);
  assert.equal(p.combine[1].parts_requests, 1);
  assert.match(previewCombine(db(), 14).error, /Only pre-start/);
  assert.match(previewCombine(db(), 20).error, /only one open pre-start/);
});

test("combining moves faults, parts, stock, time and helpers onto one job and closes the rest", () => {
  const d = db();
  const r = combinePrestartWorkOrders(d, 12, { actor: "jaco" });
  assert.deepEqual([r.keep_id, r.combined_ids], [11, [10, 12]]);

  const keep = d.prepare("SELECT * FROM work_orders WHERE id = 11").get();
  assert.equal(keep.status, "in_progress");
  assert.equal(keep.labor_hours, 3.5, "2 h + 1.5 h");
  assert.equal(keep.oil_cost, 12);
  assert.equal(keep.labor_rate_per_hour, 40);
  // One clean list: each fault once, with how often and when it was reported.
  assert.equal(keep.job_description, [
    "Combined pre-start faults from 3 work orders (#10, #11, #12) reported 2026-10-01 to 2026-10-03:",
    "- Engine oil: low (reported 3 times: 1 Oct, 2 Oct, 3 Oct)",
    "- Horn (3 Oct)",
  ].join("\n"));

  assert.deepEqual(d.prepare("SELECT work_order_id, part_name FROM maintenance_parts_requests ORDER BY id").all().map((x) => [x.work_order_id, x.part_name]),
    [[11, "Oil filter"], [11, "Horn relay"], [20, "Not this machine"]]);
  assert.deepEqual(d.prepare("SELECT reference FROM stock_movements ORDER BY id").all().map((x) => x.reference),
    ["work_order:11", "work_order:11", "work_order:120"], "WO 120 is not WO 12");
  assert.deepEqual(d.prepare("SELECT work_order_id, username FROM work_order_technicians ORDER BY id").all().map((x) => [x.work_order_id, x.username]),
    [[11, "jose"], [12, "jose"], [11, "maria"]], "jose is already on WO 11, so his duplicate row stays behind");
  assert.equal(d.prepare("SELECT primary_work_order_id FROM breakdowns").get().primary_work_order_id, 11);

  for (const id of [10, 12]) {
    const w = d.prepare("SELECT status, merged_into_id, completion_notes, closed_at FROM work_orders WHERE id = ?").get(id);
    assert.equal(w.status, "closed");
    assert.equal(w.merged_into_id, 11);
    assert.equal(w.completion_notes, "Combined into WO #11 by jaco");
    assert.ok(w.closed_at);
  }
  // Untouched: the finished pre-start job, the breakdown job, the other machine.
  assert.deepEqual(d.prepare("SELECT id, status FROM work_orders WHERE id IN (13, 14, 20) ORDER BY id").all().map((w) => w.status), ["completed", "open", "open"]);

  // A resubmit of an old check, or tomorrow's fault, now lands on the kept job.
  const oil = [{ key: "oil", label: "Engine oil", comment: "low" }];
  assert.equal(syncPrestartFaultWorkOrder(d, { assetId: 1, checkId: 102, checkDate: "2026-10-03", faults: oil }).work_order_id, 11);
  assert.equal(syncPrestartFaultWorkOrder(d, { assetId: 1, checkId: 104, checkDate: "2026-10-04", faults: oil }).work_order_id, 11);
  assert.equal(d.prepare("SELECT COUNT(*) n FROM work_orders WHERE asset_id = 1 AND source = 'prestart' AND status NOT IN ('closed','completed')").get().n, 1);
});

test("only workshop leads can combine, and only after confirming", async () => {
  const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-wo-combine-"));
  process.env.DB_PATH = path.join(tempDir, "ironlog.db");
  process.env.IRONLOG_DATA_DIR = tempDir;
  await import("../db/migrate.js");
  const { db: live } = await import("../db/client.js");
  const { default: workOrderRoutes } = await import("../routes/workorders.routes.js");
  const app = Fastify({ logger: false });
  await app.register(workOrderRoutes, { prefix: "/api/workorders" });
  await app.ready();
  try {
    live.exec(`
      INSERT INTO assets (id, asset_code, asset_name, active) VALUES (1, 'E502AM', 'Excavator 502', 1);
      INSERT INTO work_orders (id, asset_id, source, reference_id, status) VALUES (1, 1, 'prestart', 1, 'open'), (2, 1, 'prestart', 2, 'open');
    `);
    const as = (role) => ({ "x-user-name": "jaco", "x-user-role": role, "x-user-roles": role });
    assert.equal((await app.inject({ method: "GET", url: "/api/workorders/1/combine-preview", headers: as("artisan") })).statusCode, 403);
    assert.equal((await app.inject({ method: "POST", url: "/api/workorders/1/combine", headers: as("artisan"), payload: { confirm: true } })).statusCode, 403);
    const preview = await app.inject({ method: "GET", url: "/api/workorders/1/combine-preview", headers: as("plant_manager") });
    assert.equal(preview.statusCode, 200);
    assert.deepEqual([preview.json().keep.id, preview.json().combine.map((w) => w.id)], [1, [2]]);
    assert.equal((await app.inject({ method: "POST", url: "/api/workorders/1/combine", headers: as("admin"), payload: {} })).statusCode, 400);
    const done = await app.inject({ method: "POST", url: "/api/workorders/1/combine", headers: as("workshop_admin"), payload: { confirm: true } });
    assert.equal(done.statusCode, 200);
    assert.equal(live.prepare("SELECT merged_into_id FROM work_orders WHERE id = 2").get().merged_into_id, 1);
    const list = await app.inject({ method: "GET", url: "/api/workorders?status=", headers: as("admin") });
    assert.equal(list.statusCode, 200);
  } finally {
    await app.close();
    try { live.close(); } catch {}
    rmSync(tempDir, { recursive: true, force: true });
  }
});
