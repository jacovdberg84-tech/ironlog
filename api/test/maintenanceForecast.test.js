import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import {
  buildUpcomingServiceCostForecasts,
  serviceCostHistoryKey,
} from "../routes/maintenance.routes.js";

function forecastContext() {
  return {
    hasTable: (name) => new Set(["maintenance_plans", "work_orders", "stock_movements"]).has(name),
    hasColumn: (table, column) => table === "work_orders" && ["labor_hours", "labor_rate_per_hour"].includes(column),
    closedStatuses: "'closed','completed','approved'",
    woCloseExpr: "w.closed_at",
    smOutSql: "1=1",
    oilPartSql: "0=1",
    smCostWithParts: "COALESCE(sm.total_cost, 0)",
    smCostNoParts: "COALESCE(sm.total_cost, 0)",
    lubeCostDefault: 4,
  };
}

test("service-history key uses the named service tier before a legacy interval", () => {
  assert.equal(serviceCostHistoryKey({ service_name: "1000 hour service", interval_hours: 500 }), "interval:1000");
  assert.equal(serviceCostHistoryKey({ service_name: "", interval_hours: 500 }), "interval:500");
});

test("replacement plan reuses closed same-asset service history without mixing tiers", () => {
  const db = new Database(":memory:");
  db.exec(`
    CREATE TABLE maintenance_plans (id INTEGER, asset_id INTEGER, service_name TEXT, interval_hours REAL);
    CREATE TABLE work_orders (
      id INTEGER, source TEXT, reference_id INTEGER, status TEXT,
      labor_hours REAL, labor_rate_per_hour REAL, closed_at TEXT
    );
    CREATE TABLE stock_movements (id INTEGER, reference TEXT, quantity REAL, total_cost REAL);
  `);
  db.prepare("INSERT INTO maintenance_plans VALUES (?, ?, ?, ?)").run(1, 7, "1000 hour service", 500);
  db.prepare("INSERT INTO maintenance_plans VALUES (?, ?, ?, ?)").run(2, 7, "1000 hour service", 1000);
  db.prepare("INSERT INTO maintenance_plans VALUES (?, ?, ?, ?)").run(3, 7, "500 hour service", 500);
  db.prepare("INSERT INTO work_orders VALUES (?, ?, ?, ?, ?, ?, ?)").run(11, "service", 1, "closed", 10, 8, "2026-09-01");
  db.prepare("INSERT INTO stock_movements VALUES (?, ?, ?, ?)").run(1, "work_order:11", -2, 90);
  db.prepare("INSERT INTO work_orders VALUES (?, ?, ?, ?, ?, ?, ?)").run(12, "service", 3, "closed", 1, 8, "2026-09-02");
  db.prepare("INSERT INTO stock_movements VALUES (?, ?, ?, ?)").run(2, "work_order:12", -1, 50);

  const [row] = buildUpcomingServiceCostForecasts(db, [{
    plan_id: 2,
    asset_id: 7,
    asset_code: "A302AM",
    asset_name: "Bell B30D",
    service_name: "1000 hour service",
    last_service_hours: 19000,
    interval_hours: 1000,
  }], {
    nearDueHours: 50,
    horizonHours: 100,
    getAssetHours: () => 19960,
    ctx: forecastContext(),
  });

  assert.equal(row.forecast.service_events, 1);
  assert.equal(row.forecast.avg_parts_cost, 90);
  assert.equal(row.forecast.avg_labor_cost, 80);
  assert.equal(row.forecast.est_total_cost, 170);
  assert.equal(row.forecast.cost_source, "historical_asset_service_average");
  db.close();
});
