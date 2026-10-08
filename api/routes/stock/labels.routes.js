// IRONLOG/api/routes/stock/labels.routes.js — IronLog's own part labels.
// Registered by routes/stock.routes.js; shared helpers arrive through ctx.
//
// Parts without a maker's barcode get a printed label: the part code as a
// Code 128 barcode (any scanner) plus a QR (phones). The print list is kept
// on the server so the stores terminal can add to it and the office prints it.

import { db } from "../../db/client.js";
import { stockInfo } from "../../utils/techPortal.js";
import { ensurePartBarcodeSchema } from "../../utils/partBarcodes.js";
import { STORES_ROLES } from "./terminal.routes.js";

const MAX_COPIES = 100;

export function ensureLabelQueueSchema(database) {
  database.exec(`
    CREATE TABLE IF NOT EXISTS part_label_queue (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      part_id INTEGER NOT NULL UNIQUE,
      copies INTEGER NOT NULL DEFAULT 1,
      added_by TEXT,
      added_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
  `);
}

/** Parts in the shape the label printer wants: code, name, bin. */
function labelRows(database, parts, copiesOf = () => 1) {
  const stock = stockInfo(database, parts.map((p) => p.id));
  return parts.map((p) => ({
    part_code: p.part_code,
    part_name: p.part_name || "",
    bin: stock.get(p.id)?.bin || null,
    on_hand: stock.get(p.id)?.on_hand ?? 0,
    copies: copiesOf(p),
  }));
}

export default function registerLabelRoutes(app, ctx) {
  const { requireRoles } = ctx;
  ensureLabelQueueSchema(db);
  ensurePartBarcodeSchema(db);
  const userOf = (req) => String(req.headers["x-user-name"] || "").trim() || null;

  // GET /api/stock/labels/queue — the print list.
  app.get("/labels/queue", async (req, reply) => {
    if (!requireRoles(req, reply, STORES_ROLES)) return;
    const rows = db.prepare(`
      SELECT q.id AS queue_id, q.copies, q.added_by, q.added_at, p.id, p.part_code, p.part_name
      FROM part_label_queue q JOIN parts p ON p.id = q.part_id
      ORDER BY q.added_at, q.id
    `).all();
    const labels = labelRows(db, rows, (p) => p.copies);
    return { ok: true, rows: labels.map((l, i) => ({ ...l, queue_id: rows[i].queue_id, added_by: rows[i].added_by })) };
  });

  // POST /api/stock/labels/queue { part_code, copies } or { parts: [{ part_code, copies }] }
  // Adds to the print list; the same part again adds copies.
  app.post("/labels/queue", async (req, reply) => {
    if (!requireRoles(req, reply, STORES_ROLES)) return;
    const body = req.body || {};
    const wanted = Array.isArray(body.parts) ? body.parts : [{ part_code: body.part_code, copies: body.copies }];
    const lines = [];
    for (const w of wanted.slice(0, 500)) {
      const part = db.prepare(`SELECT id, part_code FROM parts WHERE UPPER(part_code) = UPPER(?)`).get(String(w?.part_code || "").trim());
      if (!part) return reply.code(404).send({ ok: false, error: `Part not found: ${w?.part_code || ""}` });
      const copies = Math.round(Number(w?.copies ?? 1));
      if (!Number.isFinite(copies) || copies < 1 || copies > MAX_COPIES) return reply.code(400).send({ ok: false, error: `Copies must be 1 to ${MAX_COPIES}` });
      lines.push({ part, copies });
    }
    const add = db.prepare(`
      INSERT INTO part_label_queue (part_id, copies, added_by) VALUES (?, ?, ?)
      ON CONFLICT(part_id) DO UPDATE SET copies = MIN(part_label_queue.copies + excluded.copies, ${MAX_COPIES}), added_by = excluded.added_by
    `);
    db.transaction(() => { for (const l of lines) add.run(l.part.id, l.copies, userOf(req)); })();
    const n = db.prepare(`SELECT COUNT(*) AS n FROM part_label_queue`).get().n;
    return { ok: true, added: lines.length, part_code: lines[0]?.part.part_code || null, in_list: n };
  });

  // PUT /api/stock/labels/queue/:id { copies }
  app.put("/labels/queue/:id", async (req, reply) => {
    if (!requireRoles(req, reply, STORES_ROLES)) return;
    const copies = Math.round(Number(req.body?.copies));
    if (!Number.isFinite(copies) || copies < 1 || copies > MAX_COPIES) return reply.code(400).send({ ok: false, error: `Copies must be 1 to ${MAX_COPIES}` });
    const r = db.prepare(`UPDATE part_label_queue SET copies = ? WHERE id = ?`).run(copies, Number(req.params.id));
    if (!r.changes) return reply.code(404).send({ ok: false, error: "Not on the print list" });
    return { ok: true };
  });

  // DELETE /api/stock/labels/queue/:id — one line; DELETE /api/stock/labels/queue — the whole list (after printing).
  app.delete("/labels/queue/:id", async (req, reply) => {
    if (!requireRoles(req, reply, STORES_ROLES)) return;
    db.prepare(`DELETE FROM part_label_queue WHERE id = ?`).run(Number(req.params.id));
    return { ok: true };
  });
  app.delete("/labels/queue", async (req, reply) => {
    if (!requireRoles(req, reply, STORES_ROLES)) return;
    const n = db.prepare(`DELETE FROM part_label_queue`).run().changes;
    return { ok: true, cleared: n };
  });

  // GET /api/stock/labels/parts?set=no_barcode|received&days=7 — parts to label in one go.
  app.get("/labels/parts", async (req, reply) => {
    if (!requireRoles(req, reply, STORES_ROLES)) return;
    const set = String(req.query?.set || "");
    let parts;
    if (set === "no_barcode") {
      // In stock, and no maker's barcode linked: these need IronLog's label.
      parts = db.prepare(`
        SELECT p.id, p.part_code, p.part_name
        FROM parts p
        WHERE NOT EXISTS (SELECT 1 FROM part_barcodes b WHERE b.part_id = p.id)
          AND (SELECT IFNULL(SUM(m.quantity), 0) FROM stock_movements m WHERE m.part_id = p.id) > 0
        ORDER BY p.part_code LIMIT 500
      `).all();
    } else if (set === "received") {
      const days = Math.min(Math.max(Math.round(Number(req.query?.days || 7)), 1), 90);
      parts = db.prepare(`
        SELECT DISTINCT p.id, p.part_code, p.part_name
        FROM stock_movements m JOIN parts p ON p.id = m.part_id
        WHERE m.movement_type = 'in' AND m.quantity > 0 AND datetime(m.created_at) >= datetime('now', ?)
        ORDER BY p.part_code LIMIT 500
      `).all(`-${days} days`);
    } else {
      return reply.code(400).send({ ok: false, error: "Choose which parts (set=no_barcode or set=received)" });
    }
    return { ok: true, rows: labelRows(db, parts) };
  });
}
