import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-cm-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
const Fastify = (await import("fastify")).default;
const { db } = await import("../db/client.js");
const cm = await import("../utils/careMaintenance.js");
const { ensureTechSchema } = await import("../utils/techActivity.js");
const registerCareMaintenanceRoutes = (await import("../routes/maintenance/care-maintenance.routes.js")).default;
process.on("exit", () => { try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

db.exec(`
  CREATE TABLE IF NOT EXISTS assets (id INTEGER PRIMARY KEY, asset_code TEXT UNIQUE, asset_name TEXT, category TEXT, active INTEGER DEFAULT 1);
  CREATE TABLE IF NOT EXISTS parts (id INTEGER PRIMARY KEY, part_code TEXT UNIQUE, part_name TEXT, unit_cost REAL);
  CREATE TABLE IF NOT EXISTS stock_movements (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL, movement_type TEXT, reference TEXT, unit_cost_usd REAL, created_at TEXT DEFAULT (datetime('now')));
  CREATE TABLE IF NOT EXISTS work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER, status TEXT DEFAULT 'open',
    opened_at TEXT DEFAULT (datetime('now')), closed_at TEXT, site_code TEXT, job_description TEXT, assigned_artisan_name TEXT, labor_hours REAL, labor_rate_per_hour REAL);
  CREATE TABLE IF NOT EXISTS maintenance_parts_requests (id INTEGER PRIMARY KEY, site_code TEXT, asset_id INTEGER, asset_code TEXT, part_code TEXT, part_name TEXT,
    qty REAL, urgency TEXT, notes TEXT, work_order_id INTEGER, status TEXT DEFAULT 'requested', requested_by TEXT);
  CREATE TABLE IF NOT EXISTS mechanic_labor_entries (id INTEGER PRIMARY KEY, work_date TEXT, technician_name TEXT, hours REAL, asset_code TEXT, job_card_no TEXT, created_by TEXT);
  CREATE TABLE IF NOT EXISTS cost_settings (key TEXT PRIMARY KEY, value TEXT);
  INSERT INTO cost_settings VALUES ('labor_cost_per_hour_default', '40');
  INSERT INTO assets (id, asset_code, asset_name, category) VALUES (1, 'E502AM', 'Excavator 50t', 'Excavator'), (2, 'TD01AM', 'Dozer D8', 'Dozer'), (3, 'LDV07', 'Hilux', 'LDV');
  INSERT INTO parts (id, part_code, part_name, unit_cost) VALUES (10, 'FLT-001', 'Engine oil filter', 25), (11, 'SEAL-KIT', 'Boom cylinder seal kit', 300);
  INSERT INTO stock_movements (part_id, quantity, movement_type, reference) VALUES (10, 6, 'in', 'delivery'), (11, 1, 'in', 'delivery');
`);
ensureTechSchema(db);

test("a plan holds each machine's work, parts and hours, and prices them", () => {
  assert.equal(cm.createPlan(db, "main", { name: "  " }).ok, false, "a plan needs a name");
  const { id } = cm.createPlan(db, "main", { name: "C&M Nov 2026", start_date: "2026-11-01", end_date: "bad" }, "jaco");
  assert.equal(cm.getPlan(db, id, "main").end_date, null, "bad dates are dropped");
  assert.equal(cm.getPlan(db, id, "other"), null, "plans belong to one site");

  const added = cm.addMachines(db, id, [1, 2, 2, 99]);
  assert.deepEqual([added.added, added.skipped], [2, 1], "repeats and unknown machines are skipped");
  assert.equal(cm.addMachines(db, id, [1]).added, 0, "a machine is on a plan once");

  const rows = db.prepare(`SELECT id, asset_id FROM cm_plan_machines WHERE plan_id = ? ORDER BY asset_id`).all(id);
  const e502 = rows[0].id;
  assert.equal(cm.updateMachine(db, e502, { status: "broken" }).ok, false);
  cm.updateMachine(db, e502, { scope: "Reseal boom cylinder, full service", planned_hours: 20, priority: "high" }, "jaco");
  assert.equal(cm.setParts(db, e502, [{ description: "x", qty: 0 }]).ok, false, "quantity must be more than 0");
  cm.setParts(db, e502, [
    { part_code: "flt-001", qty: 4 },                                   // stores part: name and price filled in
    { part_code: "SEAL-KIT", qty: 2, unit_cost: 280 },                 // own price wins
    { description: "Boom pin bush (supplier quote)", qty: 2, unit_cost: 150 }, // not a stores part
  ]);

  const d = cm.planDetail(db, id, "main");
  const m = d.machines.find((x) => x.asset_code === "E502AM");
  assert.equal(d.machines[0].asset_code, "E502AM", "high priority first");
  assert.equal(d.plan.labour_rate_used, 40, "the workshop's labour rate when the plan has none");
  assert.deepEqual(m.parts.map((p) => [p.description, p.line_cost, p.state]), [
    ["Engine oil filter", 100, "in_stock"],
    ["Boom cylinder seal kit", 560, "short"],
    ["Boom pin bush (supplier quote)", 300, "short"],
  ]);
  assert.deepEqual(m.planned, { hours: 20, parts_cost: 960, labour_cost: 800, total: 1760 });
  assert.equal(m.parts_short, 2);
  assert.equal(d.totals.planned_total, 1760);
  assert.equal(d.totals.by_status.not_started, 2);
});

test("the machine's work order carries the plan, and time and parts used come back to it", () => {
  const plan = cm.listPlans(db, "main")[0];
  const row = db.prepare(`SELECT m.id FROM cm_plan_machines m WHERE m.plan_id = ? AND m.asset_id = 1`).get(plan.id);

  // Parts requests first, for what stores can't cover.
  const req = cm.requestMissingParts(db, row.id, "main", "jaco");
  assert.deepEqual(req.created.map((c) => [c.description, c.qty]), [["Boom cylinder seal kit", 1], ["Boom pin bush (supplier quote)", 2]],
    "only what is missing is ordered");
  assert.equal(cm.requestMissingParts(db, row.id, "main", "jaco").created.length, 0, "never ordered twice");

  const r = cm.createMachineWorkOrder(db, row.id, "main", "jaco");
  assert.equal(r.ok, true);
  const wo = db.prepare(`SELECT * FROM work_orders WHERE id = ?`).get(r.work_order_id);
  assert.equal(wo.source, "care_maintenance");
  assert.equal(wo.planned_labor_hours, 20);
  assert.match(wo.job_description, /Care & maintenance: C&M Nov 2026\nReseal boom cylinder/);
  assert.equal(db.prepare(`SELECT COUNT(*) AS n FROM work_order_planned_materials WHERE work_order_id = ?`).get(wo.id).n, 2, "stores parts go onto the job");
  assert.equal(db.prepare(`SELECT COUNT(*) AS n FROM maintenance_parts_requests WHERE work_order_id = ?`).get(wo.id).n, 2, "requests follow the job");
  assert.equal(cm.createMachineWorkOrder(db, row.id, "main").status, 409, "one open job per machine");

  // Work happens: 3 h in the portal (its timesheet copy is not counted twice), 1.5 h typed on the timesheet, filters issued.
  const ev = db.prepare(`INSERT INTO tech_activity_events (work_order_id, username, action, at) VALUES (?, 'ana', ?, ?)`);
  ev.run(wo.id, "start", "2026-11-02T06:00:00.000Z");
  ev.run(wo.id, "complete", "2026-11-02T09:00:00.000Z");
  db.prepare(`INSERT INTO mechanic_labor_entries (work_date, technician_name, hours, asset_code, job_card_no, created_by) VALUES
    ('2026-11-02', 'Ana', 3, 'E502AM', ?, 'portal:ana'), ('2026-11-03', 'Jose', 1.5, 'E502AM', ?, 'jaco')`).run(String(wo.id), String(wo.id));
  db.prepare(`INSERT INTO stock_movements (part_id, quantity, movement_type, reference, unit_cost_usd) VALUES (10, -4, 'out', ?, 26)`).run(`work_order:${wo.id}`);
  cm.updateMachine(db, row.id, { status: "in_progress" });

  const m = cm.planDetail(db, plan.id, "main").machines.find((x) => x.asset_code === "E502AM");
  assert.equal(m.work_order.id, wo.id);
  assert.deepEqual(m.actual, { hours: 4.5, parts_cost: 104, labour_cost: 180, total: 284 });
  assert.equal(m.hours_used_pct, 23);
  assert.deepEqual(m.parts.map((p) => p.state), ["issued", "requested", "requested"]);
  assert.equal(m.status_label, "In progress");

  const csv = cm.planCsv(cm.planDetail(db, plan.id, "main"));
  assert.match(csv, /^Machine,Description,Priority,Status/);
  assert.match(csv, /E502AM,Excavator 50t,high,In progress,WO #\d+ \(open\),"Reseal boom cylinder, full service",960,20,800,1760,4.5,104,180,284,2,/);
  assert.match(csv, /TOTAL,2 machines/);
});

test("only plant admin, foremen and managers can change a plan; stores can look", async (t) => {
  const app = Fastify();
  const requireMaintenanceRoles = (req, reply, allowed) => {
    const roles = String(req.headers["x-user-roles"] || "").split(",");
    if (!roles.some((r) => allowed.includes(r))) { reply.code(403).send({ ok: false }); return false; }
    return true;
  };
  await app.register(async (a) => registerCareMaintenanceRoutes(a, { requireMaintenanceRoles }), { prefix: "/api/maintenance" });
  t.after(() => app.close());
  const as = (role, site = "main") => ({ "x-user-roles": role, "x-user-name": role, "x-site-code": site });

  assert.equal((await app.inject({ method: "POST", url: "/api/maintenance/cm/plans", headers: as("artisan"), payload: { name: "x" } })).statusCode, 403);
  assert.equal((await app.inject({ method: "POST", url: "/api/maintenance/cm/plans", headers: as("storeman"), payload: { name: "x" } })).statusCode, 403);
  assert.equal((await app.inject({ method: "GET", url: "/api/maintenance/cm/plans", headers: as("storeman") })).statusCode, 200);

  const created = await app.inject({ method: "POST", url: "/api/maintenance/cm/plans", headers: as("workshop_admin"), payload: { name: "C&M Dec", labour_rate: 50 } });
  const id = created.json().id;
  await app.inject({ method: "POST", url: `/api/maintenance/cm/plans/${id}/machines`, headers: as("plant_manager"), payload: { asset_ids: [3] } });
  const detail = (await app.inject({ method: "GET", url: `/api/maintenance/cm/plans/${id}`, headers: as("plant_manager") })).json();
  const rowId = detail.machines[0].id;
  assert.equal(detail.plan.labour_rate_used, 50);

  assert.equal((await app.inject({ method: "PATCH", url: `/api/maintenance/cm/machines/${rowId}`, headers: as("storeman"), payload: { status: "done" } })).statusCode, 403);
  assert.equal((await app.inject({ method: "PATCH", url: `/api/maintenance/cm/machines/${rowId}`, headers: as("plant_manager", "other"), payload: { status: "done" } })).statusCode, 404,
    "another site's plan is out of reach");
  assert.equal((await app.inject({ method: "PATCH", url: `/api/maintenance/cm/machines/${rowId}`, headers: as("plant_manager"), payload: { status: "parked" } })).statusCode, 200);
  const csv = await app.inject({ method: "GET", url: `/api/maintenance/cm/plans/${id}/export.csv`, headers: as("storeman") });
  assert.match(csv.headers["content-type"], /text\/csv/);
  assert.match(csv.body, /LDV07,Hilux,normal,Parked \(preserved\)/);
  assert.equal((await app.inject({ method: "DELETE", url: `/api/maintenance/cm/machines/${rowId}`, headers: as("admin") })).statusCode, 200);
  assert.equal(cm.planDetail(db, id, "main").machines.length, 0);
});

test("the job board and portal name a C&M job by its work, not its parts list", async () => {
  const { jobLine } = await import("../utils/techPortal.js");
  assert.equal(jobLine({ source: "care_maintenance", job_description: "Care & maintenance: C&M Nov\nReseal boom cylinder\n\nParts planned:\n- 4 x FLT-001 Oil filter" }), "C&M: Reseal boom cylinder");
  assert.equal(jobLine({ source: "care_maintenance", job_description: "Care & maintenance: C&M Nov\n\nParts planned:\n- 1 x Seal kit" }), "Care & maintenance");
});
