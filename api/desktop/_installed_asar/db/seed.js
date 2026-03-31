// IRONLOG/api/db/seed.js (SQLite)
import { db } from "./client.js";

function seed() {
  console.log("🌱 Seeding IRONLOG SQLite database...");

  const tx = db.transaction(() => {
    // Clean (dev-safe)
    db.exec(`
      DELETE FROM kpi_snapshots;
      DELETE FROM parts_orders;
      DELETE FROM stock_movements;
      DELETE FROM parts;
      DELETE FROM work_orders;
      DELETE FROM breakdowns;
      DELETE FROM maintenance_plans;
      DELETE FROM oil_logs;
      DELETE FROM fuel_logs;
      DELETE FROM daily_hours;
      DELETE FROM assets;
    `);

    // Assets
    const insertAsset = db.prepare(`
      INSERT INTO assets (asset_code, asset_name, category, is_standby, active)
      VALUES (?, ?, ?, ?, ?)
    `);

    const assets = [
   
    ];

    for (const a of assets) insertAsset.run(...a);

    const getAssetId = db.prepare(`SELECT id FROM assets WHERE asset_code = ?`);

    const excId = getAssetId.get("EXC-01").id;
    const dtId = getAssetId.get("DT-01").id;
    const genId = getAssetId.get("GEN-01").id;

    // Today date as YYYY-MM-DD
    const dateStr = new Date().toISOString().slice(0, 10);

    // Daily hours
    const insertHours = db.prepare(`
      INSERT INTO daily_hours (asset_id, work_date, hours_run, is_used, operator, notes)
      VALUES (?, ?, ?, ?, ?, ?)
    `);

    insertHours.run(excId, dateStr, 9.5, 1, "Operator A", "Normal shift");
    insertHours.run(dtId, dateStr, 7.0, 1, "Operator B", "Road haul");
    insertHours.run(genId, dateStr, 0.0, 0, null, "Standby unit");

    // Fuel
    const insertFuel = db.prepare(`
      INSERT INTO fuel_logs (asset_id, log_date, liters, source)
      VALUES (?, ?, ?, ?)
    `);

    insertFuel.run(excId, dateStr, 180.0, "bowser");
    insertFuel.run(dtId, dateStr, 220.0, "bowser");

    // Oil
    db.prepare(`
      INSERT INTO oil_logs (asset_id, log_date, oil_type, quantity)
      VALUES (?, ?, ?, ?)
    `).run(excId, dateStr, "Engine Oil", 5.0);

    // Maintenance plan
    db.prepare(`
      INSERT INTO maintenance_plans (asset_id, service_name, interval_hours, last_service_hours, active)
      VALUES (?, ?, ?, ?, 1)
    `).run(excId, "250h Service", 250, 120);

    // Breakdown
    const breakdownRes = db.prepare(`
      INSERT INTO breakdowns (asset_id, breakdown_date, description, downtime_hours, critical)
      VALUES (?, ?, ?, ?, ?)
    `).run(excId, dateStr, "Hydraulic leak", 3.5, 1);

    const breakdownId = breakdownRes.lastInsertRowid;

    // Work order linked to breakdown
    const woRes = db.prepare(`
      INSERT INTO work_orders (asset_id, source, reference_id, status)
      VALUES (?, 'breakdown', ?, 'open')
    `).run(excId, Number(breakdownId));

    const workOrderId = woRes.lastInsertRowid;

    // Parts
    const insertPart = db.prepare(`
      INSERT INTO parts (part_code, part_name, critical, min_stock)
      VALUES (?, ?, ?, ?)
    `);

    const parts = [
      ["FLT-001", "Oil Filter", 1, 5],
      ["HOS-010", "Hydraulic Hose 1/2\"", 1, 2],
      ["GRS-020", "Grease Cartridge", 0, 10]
    ];

    for (const p of parts) insertPart.run(...p);

    const getPartId = db.prepare(`SELECT id FROM parts WHERE part_code = ?`);
    const fltId = getPartId.get("FLT-001").id;
    const hosId = getPartId.get("HOS-010").id;
    const grsId = getPartId.get("GRS-020").id;

    // Opening stock movements (in)
    const insertMove = db.prepare(`
      INSERT INTO stock_movements (part_id, quantity, movement_type, reference)
      VALUES (?, ?, ?, ?)
    `);

    insertMove.run(fltId, 10, "in", "opening_balance");
    insertMove.run(hosId, 3, "in", "opening_balance");
    insertMove.run(grsId, 30, "in", "opening_balance");

    // Issue a part to the work order (out)
    insertMove.run(hosId, -1, "out", `work_order:${workOrderId}`);

    // Parts on order
    db.prepare(`
      INSERT INTO parts_orders (part_id, quantity, expected_date, status)
      VALUES (?, ?, ?, 'ordered')
    `).run(hosId, 5, dateStr);
  });

  tx();

  console.log("✅ Seeding complete.");
  console.log("DB file:", process.env.DB_PATH || "./db/ironlog.db");
}

try {
  seed();
} catch (err) {
  console.error("❌ Seed failed:", err.message);
  process.exitCode = 1;
}