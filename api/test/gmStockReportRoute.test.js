import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";
import ExcelJS from "exceljs";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-gm-stock-report-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { default: stockRoutes } = await import("../routes/stock.routes.js");
const app = Fastify({ logger: false });
await app.register(stockRoutes, { prefix: "/api/stock" });
await app.ready();

const locationId = db.prepare("SELECT id FROM stock_locations WHERE location_code = 'MAIN'").get().id;
const partId = Number(db.prepare(`
  INSERT INTO parts (part_code, part_name, critical, min_stock, unit_cost, department_code)
  VALUES ('KIT-500', '500 hour service kit', 1, 13, 125.50, 'WORKSHOP')
`).run().lastInsertRowid);
const insertMovement = db.prepare(`
  INSERT INTO stock_movements (
    part_id, quantity, movement_type, reference, created_at, location_id, unit_cost_usd
  ) VALUES (?, ?, ?, ?, ?, ?, ?)
`);
insertMovement.run(partId, 10, "receipt", "OPENING-001", "2026-08-31 08:00:00", locationId, 120);
insertMovement.run(partId, -3, "issue_to_work_order", "WO-100", "2026-09-10 10:00:00", locationId, 125.5);
insertMovement.run(partId, 5, "receipt", "PO-200", "2026-09-15 12:00:00", locationId, 125.5);

test("GM stock route creates a live workbook with period balances from the Stores ledger", async (t) => {
  t.after(async () => {
    await app.close();
    db.close();
    rmSync(tempDir, { recursive: true, force: true });
  });

  const reply = await app.inject({
    method: "GET",
    url: "/api/stock/gm-stock-report.xlsx?period=monthly&report_date=2026-09-20",
  });

  assert.equal(reply.statusCode, 200, reply.body);
  assert.match(reply.headers["content-type"], /spreadsheetml/);
  assert.match(reply.headers["content-disposition"], /IRONLOG_Monthly_Stock_Report_2026-09-01_to_2026-09-30\.xlsx/);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(reply.rawPayload);
  const register = workbook.getWorksheet("Workshop Spares");
  assert.equal(register.getCell("A5").value, "KIT-500");
  assert.equal(register.getCell("J5").value, 10);
  assert.equal(register.getCell("K5").value, 5);
  assert.equal(register.getCell("L5").value, 3);
  assert.equal(register.getCell("Q5").value, 12);
  assert.equal(register.getCell("U5").value, "REORDER");
  assert.equal(workbook.getWorksheet("Stock Movements").rowCount, 6);
});
