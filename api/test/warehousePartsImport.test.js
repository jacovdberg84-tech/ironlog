import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";
import ExcelJS from "exceljs";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-warehouse-import-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { default: stockRoutes } = await import("../routes/stock.routes.js");
const { parseWarehousePartsWorkbook } = await import("../utils/warehousePartsImport.js");

const app = Fastify({ logger: false });
await app.register(stockRoutes, { prefix: "/api/stock" });
await app.ready();

process.on("exit", () => {
  try { db.close(); } catch { /* already closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

const H = { "x-user-name": "bj.vandenberg", "x-user-role": "admin", "x-user-roles": "admin", "x-site-code": "main" };
const HEADERS = [
  "SITE", "WHS/ STATUS", "PRIORITY CODE", "SUPPLIER", "DESCRIPTION", "PART NO",
  "QTY RECVD", "PRICE", "COST", "WEIGHT / EACH", "WEIGHT TOTAL", "MANUAL REQ NOS / FLEET",
  "SAGE REQ NUMBER", "ORDER NO", "SALES ORDER", "ORDER QTY", "OUTSTANDING", "UOM", "ORDER DATE",
  "DATE WAITING / RECEIVED", "DAYS WAITING", "PRODUCT CODE", "COUNTRY OF ORIGIN", "COMMENTS", "COMMERCIAL INVOICE PACKING SLIP",
];

async function warehouseWorkbook({ received = 2 } = {}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("DBN & BKS AML LIST");
  ws.addRow(HEADERS);
  ws.addRow([
    "AML", "DBN", 1, "BARLOWORLD", "BUSHING", "387-5938", received, 9598.92, 19197.84,
    4.05, 8.1, "002017 / F500AM", "VS120-00338", "VS120P148656", "VS120S000289", 2,
    Math.max(0, 2 - received), "UN", new Date(Date.UTC(2026, 6, 7)), new Date(Date.UTC(2026, 6, 13)),
    6, "120-00110", "ZA", "Awaiting site shipment", "CI-123",
  ]);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

function multipartBody(filename, buffer) {
  const boundary = "----ironlogWarehouseImport";
  return {
    boundary,
    body: Buffer.concat([
      Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="status"\r\n\r\nwarehouse_ready\r\n`),
      Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="currency"\r\n\r\nZAR\r\n`),
      Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="${filename}"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`),
      buffer,
      Buffer.from(`\r\n--${boundary}--\r\n`),
    ]),
  };
}

test("weekly warehouse import maps supplier fields, updates matching lines and preserves completed status", async (t) => {
  t.after(() => app.close());
  db.exec("INSERT INTO assets (id, asset_code, asset_name, category, active) VALUES (1, 'F500AM', 'CAT 950GC', 'Loader', 1)");

  const firstFile = await warehouseWorkbook({ received: 2 });
  const parsed = await parseWarehousePartsWorkbook(firstFile, "DURBAN & BOKSBURG AML SITE UPDATE.xlsx");
  assert.equal(parsed.errors.length, 0);
  assert.equal(parsed.rows.length, 1);
  assert.equal(parsed.rows[0].source_reference, "warehouse:VS120P148656:387-5938");
  assert.equal(parsed.rows[0].order_date, "2026-07-07");
  assert.equal(parsed.rows[0].warehouse_date, "2026-07-13");

  let upload = multipartBody("DURBAN & BOKSBURG AML SITE UPDATE.xlsx", firstFile);
  let reply = await app.inject({
    method: "POST",
    url: "/api/stock/part-orders/warehouse-import",
    headers: { ...H, "content-type": `multipart/form-data; boundary=${upload.boundary}` },
    payload: upload.body,
  });
  assert.equal(reply.statusCode, 200, reply.body);
  assert.equal(reply.json().created, 1);
  let row = db.prepare(`SELECT * FROM stores_part_orders WHERE source_reference = ?`).get("warehouse:VS120P148656:387-5938");
  assert.equal(row.status, "warehouse_ready");
  assert.equal(row.currency, "ZAR");
  assert.equal(row.warehouse_code, "DBN");
  assert.equal(row.supplier_qty_received, 2);
  assert.equal(row.supplier_outstanding_qty, 0);
  assert.equal(row.asset_id, 1);

  upload = multipartBody("DURBAN & BOKSBURG AML SITE UPDATE.xlsx", await warehouseWorkbook({ received: 1 }));
  reply = await app.inject({
    method: "POST",
    url: "/api/stock/part-orders/warehouse-import",
    headers: { ...H, "content-type": `multipart/form-data; boundary=${upload.boundary}` },
    payload: upload.body,
  });
  assert.equal(reply.statusCode, 200, reply.body);
  assert.equal(reply.json().created, 0);
  assert.equal(reply.json().updated, 1);
  row = db.prepare(`SELECT * FROM stores_part_orders WHERE source_reference = ?`).get("warehouse:VS120P148656:387-5938");
  assert.equal(row.supplier_qty_received, 1);
  assert.equal(row.supplier_outstanding_qty, 1);

  db.prepare("UPDATE stores_part_orders SET status = 'arrived' WHERE id = ?").run(row.id);
  upload = multipartBody("DURBAN & BOKSBURG AML SITE UPDATE.xlsx", await warehouseWorkbook({ received: 2 }));
  reply = await app.inject({
    method: "POST",
    url: "/api/stock/part-orders/warehouse-import",
    headers: { ...H, "content-type": `multipart/form-data; boundary=${upload.boundary}` },
    payload: upload.body,
  });
  assert.equal(reply.statusCode, 200, reply.body);
  assert.equal(reply.json().completion_preserved, 1);
  assert.equal(db.prepare("SELECT status FROM stores_part_orders WHERE id = ?").get(row.id).status, "arrived");
});
