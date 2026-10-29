import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";
import Database from "better-sqlite3";
import {
  OPEN_WORK_QUEUE_ROLES,
  buildOpenWorkQueue,
  compactQueueForBorris,
  describeOpenWorkQueue,
  isWorkQueueQuestion,
} from "../utils/openWorkQueue.js";
import { readWebAppSource } from "./webAppSource.js";

function seed() {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT, category TEXT);
    CREATE TABLE breakdowns (id INTEGER PRIMARY KEY, asset_id INTEGER, status TEXT, component TEXT, description TEXT,
      parts_status TEXT, critical INTEGER, breakdown_date TEXT, start_at TEXT, site_code TEXT);
    CREATE TABLE maintenance_plans (id INTEGER PRIMARY KEY, asset_id INTEGER, service_name TEXT);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER, status TEXT,
      opened_at TEXT, site_code TEXT, assigned_artisan_name TEXT, artisan_name TEXT, priority TEXT, due_date TEXT,
      job_description TEXT, repair_progress TEXT, merged_into_id INTEGER);
    CREATE TABLE parts (id INTEGER PRIMARY KEY, part_code TEXT, part_name TEXT, min_stock REAL);
    CREATE TABLE stock_movements (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL);
    CREATE TABLE maintenance_parts_requests (id INTEGER PRIMARY KEY, site_code TEXT, part_code TEXT, part_name TEXT,
      qty REAL, urgency TEXT, work_order_id INTEGER, status TEXT, created_at TEXT);
    CREATE TABLE stores_part_orders (id INTEGER PRIMARY KEY, site_code TEXT, part_id INTEGER, part_code TEXT, qty REAL,
      supplier_outstanding_qty REAL, expected_arrival_date TEXT, status TEXT);

    INSERT INTO assets VALUES (1, 'EX01', 'Excavator', 'Excavator'), (2, 'DT02', 'Dump truck', 'ADT'),
      (3, 'LD03', 'Loader', 'Loader'), (4, 'GR04', 'Grader', 'Grader');
    INSERT INTO breakdowns VALUES
      (1, 1, 'OPEN', 'Hydraulics', 'Boom cylinder leaking', 'Ordered', 1, '2026-10-20', NULL, 'main');
    INSERT INTO maintenance_plans VALUES (7, 3, '500h service');
    INSERT INTO parts VALUES (1, 'SEAL-1', 'Seal kit', 4), (2, 'PUMP-9', 'Hydraulic pump', 1), (3, 'FLT-2', 'Oil filter', 2);
    INSERT INTO stock_movements VALUES (1, 1, 3), (2, 2, 0), (3, 3, 10);
    INSERT INTO work_orders VALUES
      (10, 1, 'breakdown', 1, 'in_progress', '2026-10-21 08:00:00', 'main', 'sipho', NULL, NULL, NULL, NULL, 'Cylinder stripped', NULL),
      (11, 2, 'manual', NULL, 'open', '2026-10-25 09:00:00', 'main', NULL, NULL, NULL, NULL, 'Replace hydraulic pump\nsecond line', NULL, NULL),
      (12, 3, 'service', 7, 'assigned', '2026-10-28 07:00:00', 'main', 'johan', NULL, NULL, NULL, NULL, NULL, NULL),
      (13, 4, 'manual', NULL, 'open', '2026-09-01 07:00:00', 'main', NULL, NULL, NULL, NULL, 'Cab door hinge', NULL, NULL),
      (14, 4, 'manual', NULL, 'closed', '2026-09-01 07:00:00', 'main', NULL, NULL, NULL, NULL, 'Done', NULL, NULL),
      (15, 4, 'manual', NULL, 'Completed', '2026-09-01 07:00:00', 'main', NULL, NULL, NULL, NULL, 'Done', NULL, NULL),
      (16, 4, 'manual', NULL, 'open', '2026-09-01 07:00:00', 'main', NULL, NULL, NULL, NULL, 'Merged away', NULL, 13),
      (17, 4, 'manual', NULL, 'open', '2026-09-01 07:00:00', 'north', NULL, NULL, NULL, NULL, 'Other site', NULL, NULL);
    INSERT INTO maintenance_parts_requests VALUES
      (1, 'main', 'SEAL-1', 'Seal kit', 2, 'critical', 10, 'requested', '2026-10-21'),
      (2, 'main', 'PUMP-9', 'Hydraulic pump', 1, 'urgent', 11, 'requested', '2026-10-25'),
      (3, 'main', 'flt-2', 'Oil filter', 2, 'normal', 12, 'requested', '2026-10-28'),
      (4, 'main', 'PUMP-9', 'Hydraulic pump', 1, 'normal', 12, 'received', '2026-10-28'),
      (5, 'main', 'SEAL-1', 'Seal kit', 1, 'normal', 13, 'cancelled', '2026-10-01');
    INSERT INTO stores_part_orders VALUES
      (1, 'north', 2, 'PUMP-9', 1, NULL, '2026-11-02', 'on_order'),
      (2, 'main', 1, 'SEAL-1', 4, 3, '2026-11-05', 'in_transit'),
      (3, 'main', 1, 'SEAL-1', 4, NULL, '2026-10-01', 'arrived');
  `);
  return db;
}

test("lists every open work order on the site with machine, job, age, status and technician", () => {
  const q = buildOpenWorkQueue(seed(), { site: "main", today: "2026-10-29" });
  assert.equal(q.ok, true);
  assert.equal(q.as_of, "2026-10-29");
  assert.deepEqual(q.work_orders.map((w) => w.id), [10, 11, 13, 12], "down first, then parts blockers, then oldest");
  const bd = q.work_orders[0];
  assert.equal(bd.asset_code, "EX01");
  assert.equal(bd.job, "Hydraulics — Boom cylinder leaking");
  assert.equal(bd.age_days, 9, "breakdown jobs count from the breakdown date");
  assert.equal(bd.status, "in_progress");
  assert.equal(bd.technician, "sipho");
  assert.equal(bd.machine_down, true);
  assert.equal(bd.critical, true);
  assert.equal(bd.progress, "Cylinder stripped");
  const wo11 = q.work_orders.find((w) => w.id === 11);
  assert.equal(wo11.job, "Replace hydraulic pump", "only the first line of the job description");
  assert.equal(wo11.technician, null);
  assert.equal(q.work_orders.find((w) => w.id === 12).job, "500h service");
  assert.deepEqual(q.summary, {
    open: 4, machines_down: 1, unassigned: 2, parts_blocked: 1, parts_on_order: 0, parts_ready: 2, older_than_14_days: 1,
  });
});

test("each requested part shows stock on hand, below-minimum and on-order status", () => {
  const q = buildOpenWorkQueue(seed(), { site: "main", today: "2026-10-29" });
  const byId = new Map(q.work_orders.map((w) => [w.id, w]));

  const [seal] = byId.get(10).parts;
  assert.equal(seal.on_hand, 3);
  assert.equal(seal.in_stock, true, "3 on hand covers 2 requested");
  assert.equal(seal.below_min, true, "3 is under the minimum of 4");
  assert.equal(seal.on_order, true, "an in-transit Stores order on this site counts");
  assert.equal(seal.on_order_qty, 3, "outstanding supplier quantity is used when known");
  assert.equal(seal.expected_arrival, "2026-11-05");
  assert.equal(byId.get(10).parts_state, "ready");

  const [pump] = byId.get(11).parts;
  assert.equal(pump.on_hand, 0);
  assert.equal(pump.below_min, true);
  assert.equal(pump.on_order, false, "an order on another site does not count");
  assert.equal(byId.get(11).parts_state, "blocked");

  assert.equal(byId.get(12).parts.length, 1, "received requests are no longer waiting");
  assert.equal(byId.get(12).parts[0].on_hand, 10, "part codes match regardless of case");
  assert.equal(byId.get(12).parts[0].below_min, false);
  assert.equal(byId.get(13).parts_state, "none", "cancelled requests are ignored");
});

test("a short part already ordered on the request counts as on order", () => {
  const db = seed();
  db.prepare(`UPDATE maintenance_parts_requests SET status = 'ordered' WHERE id = 2`).run();
  const wo = buildOpenWorkQueue(db, { site: "main", today: "2026-10-29" }).work_orders.find((w) => w.id === 11);
  assert.equal(wo.parts[0].on_order, true);
  assert.equal(wo.parts_state, "on_order");
});

test("closed, completed and merged work orders and other sites are left out", () => {
  const ids = buildOpenWorkQueue(seed(), { site: "main", today: "2026-10-29" }).work_orders.map((w) => w.id);
  for (const gone of [14, 15, 16, 17]) assert.ok(!ids.includes(gone), `WO ${gone} should not be listed`);
  assert.deepEqual(buildOpenWorkQueue(seed(), { site: "north", today: "2026-10-29" }).work_orders.map((w) => w.id), [17]);
});

test("building the queue never changes a record", () => {
  const db = seed();
  const before = db.prepare(`SELECT
    (SELECT group_concat(id || status, ',') FROM work_orders) AS wo,
    (SELECT SUM(quantity) FROM stock_movements) AS stock,
    (SELECT group_concat(id || status, ',') FROM maintenance_parts_requests) AS reqs,
    (SELECT group_concat(id || status, ',') FROM stores_part_orders) AS orders`).get();
  const writes = [];
  const readOnly = new Proxy(db, {
    get(target, prop) {
      if (prop === "prepare") {
        return (sql) => {
          if (!/^\s*(SELECT|PRAGMA)\b/i.test(sql)) writes.push(sql);
          return target.prepare(sql);
        };
      }
      if (prop === "exec") return (sql) => { writes.push(sql); };
      const v = target[prop];
      return typeof v === "function" ? v.bind(target) : v;
    },
  });
  buildOpenWorkQueue(readOnly, { site: "main", today: "2026-10-29" });
  assert.deepEqual(writes, [], "only SELECT and PRAGMA statements are used");
  const after = db.prepare(`SELECT
    (SELECT group_concat(id || status, ',') FROM work_orders) AS wo,
    (SELECT SUM(quantity) FROM stock_movements) AS stock,
    (SELECT group_concat(id || status, ',') FROM maintenance_parts_requests) AS reqs,
    (SELECT group_concat(id || status, ',') FROM stores_part_orders) AS orders`).get();
  assert.deepEqual(after, before);
});

test("works on a database without parts requests or Stores orders yet", () => {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT, category TEXT);
    CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER, status TEXT, opened_at TEXT);
    INSERT INTO assets VALUES (1, 'EX01', 'Excavator', 'Excavator');
    INSERT INTO work_orders VALUES (1, 1, 'manual', NULL, 'open', '2026-10-29 06:00:00');
  `);
  const q = buildOpenWorkQueue(db, { site: "main", today: "2026-10-29" });
  assert.equal(q.work_orders.length, 1);
  assert.equal(q.work_orders[0].age_days, 0);
  assert.equal(q.work_orders[0].parts_state, "none");
  assert.deepEqual(buildOpenWorkQueue(new Database(":memory:"), { today: "2026-10-29" }).work_orders, []);
});

test("Borris gets a compact copy and a plain-language fallback", () => {
  const q = buildOpenWorkQueue(seed(), { site: "main", today: "2026-10-29" });
  const compact = compactQueueForBorris(q, 2);
  assert.equal(compact.shown, 2);
  assert.equal(compact.work_orders[0].wo, 10);
  assert.deepEqual(Object.keys(compact.work_orders[0].parts[0]), ["part", "qty", "on_hand", "below_min", "on_order", "eta"]);
  const text = describeOpenWorkQueue(q);
  assert.match(text, /Open work orders: 4/);
  assert.match(text, /1\. WO #10 EX01 .*machine down, critical/);
  assert.match(text, /WO #11 DT02 .*no technician assigned; parts short and not ordered: PUMP-9/);
  assert.match(text, /read-only/i);
  assert.equal(describeOpenWorkQueue({ work_orders: [] }), "There are no open work orders on this site right now.");
});

test("queue questions are recognised", () => {
  for (const q of [
    "Prioritise the open work orders",
    "Give me a practical priority assessment of all open work orders, including parts blockers.",
    "what are the parts blockers?",
    "show the work order queue",
    "which outstanding jobs should we do first",
  ]) assert.equal(isWorkQueueQuestion(q), true, q);
  for (const q of ["Torque for CAT 320 boom bolts", "How many oil filters on hand?", "Overdue services this week"]) {
    assert.equal(isWorkQueueQuestion(q), false, q);
  }
});

test("only leads may see the whole queue; the routes check this on the server", () => {
  assert.deepEqual(OPEN_WORK_QUEUE_ROLES, ["admin", "supervisor", "workshop_admin", "plant_manager"]);
  for (const role of ["artisan", "operator", "storeman", "stores"]) assert.ok(!OPEN_WORK_QUEUE_ROLES.includes(role));
  const weekly = fs.readFileSync(new URL("../routes/maintenance/weekly.routes.js", import.meta.url), "utf8");
  const route = weekly.slice(weekly.indexOf('app.get("/open-work-queue"'));
  assert.match(route.split("\n").slice(0, 3).join("\n"), /requireAnyRole\(req, reply, OPEN_WORK_QUEUE_ROLES\)/);
  const ironmind = fs.readFileSync(new URL("../routes/ironmind.routes.js", import.meta.url), "utf8");
  const ask = ironmind.slice(ironmind.indexOf('app.post("/ask"'), ironmind.indexOf("isManualQuestion(question)) {"));
  assert.match(ask, /holdsAnyRole\(req, OPEN_WORK_QUEUE_ROLES\)/);
  assert.match(ask, /code\(403\)/);
});

test("the Borris screen shows the queue with escaped text and restores the button on failure", async () => {
  const source = readWebAppSource();
  const start = source.indexOf("function borrisCommandMetric(");
  const helper = source.slice(start, source.indexOf("const BORRIS_DRAFT_COPY"));
  const els = {};
  const el = (id) => (els[id] ||= { id, textContent: "", innerHTML: "", disabled: false });
  const escapeHtml = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  let fetched = null;
  const context = vm.createContext({
    qs: el, API: "", escapeHtml, setStatus: () => {},
    fetchJson: async (url) => { fetched = url; return buildOpenWorkQueue(seed(), { site: "main", today: "2026-10-29" }); },
  });
  vm.runInContext(helper, context);
  await context.loadBorrisWorkQueue();
  assert.equal(fetched, "/api/maintenance/open-work-queue");
  const table = el("borrisQueueTable").innerHTML;
  assert.match(table, /<th>Fault \/ job<\/th>/);
  assert.match(table, /#10/);
  assert.match(table, /Below minimum/);
  assert.match(table, /On order · ETA 2026-11-05/);
  assert.match(table, /Parts short, not ordered/);
  assert.match(el("borrisQueueMetrics").innerHTML, /Parts blocked/);
  assert.equal(el("borrisQueueLoadBtn").disabled, false);

  context.fetchJson = async () => ({ ok: true, as_of: "2026-10-29", summary: {}, work_orders: [
    { id: 1, asset_code: "<b>X</b>", job: "<script>", parts: [], parts_state: "none" },
  ] });
  await context.loadBorrisWorkQueue();
  assert.ok(!el("borrisQueueTable").innerHTML.includes("<script>"));
  assert.ok(el("borrisQueueTable").innerHTML.includes("&lt;b&gt;X&lt;/b&gt;"));

  context.fetchJson = async () => { throw new Error("role 'artisan' not allowed"); };
  await context.loadBorrisWorkQueue();
  assert.match(el("borrisQueueMeta").textContent, /not allowed/);
  assert.equal(el("borrisQueueLoadBtn").disabled, false);
});
