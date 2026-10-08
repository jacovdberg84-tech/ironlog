import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-june-calendar-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
const Fastify = (await import("fastify")).default;
const juneRoutes = (await import("../routes/june.routes.js")).default;
const { db } = await import("../db/client.js");

const admin = { "x-user-role": "admin", "x-user-roles": "admin", "x-user-name": "jaco", "x-site-code": "main" };

test("June prepares an internal calendar change and only creates it after browser confirmation", async (t) => {
  t.after(async () => {
    db.close();
    rmSync(tempDir, { recursive: true, force: true });
  });
  const app = Fastify();
  await app.register(juneRoutes, { prefix: "/api/june" });
  t.after(() => app.close());
  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, asset_name TEXT, category TEXT);
    CREATE TABLE asset_hours (asset_id INTEGER PRIMARY KEY, total_hours REAL);
    CREATE TABLE daily_hours (id INTEGER PRIMARY KEY, asset_id INTEGER, work_date TEXT, closing_hours REAL, hours_run REAL, is_used INTEGER);
    CREATE TABLE maintenance_plans (id INTEGER PRIMARY KEY, asset_id INTEGER, service_name TEXT, interval_hours REAL, last_service_hours REAL, active INTEGER);
    CREATE TABLE parts (id INTEGER PRIMARY KEY, part_code TEXT, part_name TEXT, min_stock REAL, critical INTEGER, unit_cost REAL);
    CREATE TABLE stock_movements (id INTEGER PRIMARY KEY, part_id INTEGER, quantity REAL);
    CREATE TABLE stores_part_orders (
      id INTEGER PRIMARY KEY,
      site_code TEXT,
      part_code TEXT,
      part_name TEXT,
      qty REAL,
      unit_cost REAL,
      currency TEXT,
      status TEXT,
      order_date TEXT,
      expected_arrival_date TEXT,
      current_location TEXT,
      supplier_name TEXT,
      po_number TEXT,
      requisition_number TEXT
    );
  `);
  db.prepare("INSERT INTO assets (id, asset_code, asset_name, category) VALUES (1, 'GS04AM', 'John Deere 30KVA', 'Generator')").run();
  db.prepare("INSERT INTO asset_hours (asset_id, total_hours) VALUES (1, 35498.7)").run();
  db.prepare("INSERT INTO daily_hours (id, asset_id, work_date, closing_hours, hours_run, is_used) VALUES (1, 1, '2026-10-07', 2887, 9, 1)").run();
  db.prepare("INSERT INTO maintenance_plans (id, asset_id, service_name, interval_hours, last_service_hours, active) VALUES (1, 1, '500 hour service', 500, 2500, 1)").run();
  db.prepare("INSERT INTO parts (id, part_code, part_name, min_stock, critical, unit_cost) VALUES (1, 'KIT500', '500 hour service kit', 2, 1, 844.86)").run();
  db.prepare("INSERT INTO stock_movements (id, part_id, quantity) VALUES (1, 1, 1)").run();
  db.prepare(`INSERT INTO stores_part_orders
    (id, site_code, part_code, part_name, qty, unit_cost, currency, status, order_date, expected_arrival_date, current_location, supplier_name, po_number, requisition_number)
    VALUES (1, 'main', 'KIT500', '500 hour service kit', 3, 844.86, 'USD', 'warehouse_ready', '2026-10-01', '2026-10-15', 'Durban warehouse', 'AML Supply', 'PO-77', 'REQ-44')`).run();

  const draft = await app.inject({
    method: "POST",
    url: "/api/june/gateway/execute",
    headers: admin,
    payload: {
      name: "june_prepare_internal_calendar_change",
      arguments: {
        action: "create",
        title: "Weekly maintenance forum",
        event_date: "2026-10-12",
        start_time: "08:00",
        end_time: "09:00",
        category: "meeting",
      },
    },
  });
  assert.equal(draft.statusCode, 200);
  const result = draft.json().result;
  assert.equal(result.state, "pending_confirmation");
  assert.equal(result.event.title, "Weekly maintenance forum");
  assert.match(result.approval.token, /^[0-9a-f-]{36}$/i);

  const before = await app.inject({ method: "GET", url: "/api/june/calendar/internal/events?days=30", headers: admin });
  assert.equal(before.statusCode, 200);
  assert.equal(before.json().calendar.events.length, 0);

  const otherAdmin = await app.inject({
    method: "POST",
    url: `/api/june/calendar/internal/approvals/${result.approval.token}/confirm`,
    headers: { ...admin, "x-user-name": "another-admin" },
    payload: {},
  });
  assert.equal(otherAdmin.statusCode, 404);

  const confirmed = await app.inject({
    method: "POST",
    url: `/api/june/calendar/internal/approvals/${result.approval.token}/confirm`,
    headers: admin,
    payload: {},
  });
  assert.equal(confirmed.statusCode, 200);
  assert.equal(confirmed.json().event.title, "Weekly maintenance forum");

  const after = await app.inject({ method: "GET", url: "/api/june/calendar/internal/events?days=30", headers: admin });
  assert.equal(after.statusCode, 200);
  assert.equal(after.json().calendar.events.length, 1);
  assert.equal(after.json().calendar.events[0].title, "Weekly maintenance forum");

  const engineering = await app.inject({
    method: "POST",
    url: "/api/june/gateway/execute",
    headers: admin,
    payload: { name: "june_borris_engineering_brief", arguments: { asset_code: "GS04AM", as_of: "2026-10-08" } },
  });
  assert.equal(engineering.statusCode, 200);
  assert.equal(engineering.json().result.asset.current_meter, 2887);
  assert.equal(engineering.json().result.asset.meter_source, "daily_closing");
  assert.equal(engineering.json().result.service_plans[0].remaining_hours, 113);

  const stores = await app.inject({
    method: "POST",
    url: "/api/june/gateway/execute",
    headers: admin,
    payload: { name: "june_get_stores_brief", arguments: { scope: "part_lookup", query: "KIT500" } },
  });
  assert.equal(stores.statusCode, 200);
  assert.equal(stores.json().result.matching_parts[0].on_hand, 1);
  assert.equal(stores.json().result.matching_parts[0].shortage, 1);
  assert.equal(stores.json().result.parts_on_order[0].status, "warehouse_ready");
  assert.equal(stores.json().result.parts_on_order[0].supplier_name, "AML Supply");
});
