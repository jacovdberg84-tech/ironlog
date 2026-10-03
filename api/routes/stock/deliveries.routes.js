// IRONLOG/api/routes/stock/deliveries.routes.js — receive a whole delivery at once.
// Registered by routes/stock.routes.js; shared helpers arrive through ctx.
//
// A delivery is one supplier invoice / GRN with many lines. Supplier, reference,
// date, currency and store are entered once; each line is a part (existing, or
// a new code with its description), quantity, unit cost and optional bin. All
// lines are saved together or none are. Before saving, the delivery is checked
// against recent receipts (same invoice number, or the same part and quantity
// in the last 7 days); the storeman sees the matches and confirms or stops.

import { db } from "../../db/client.js";
import { validateAgainstMdmPolicy, validatePartGovernanceOptional } from "../../utils/masterdataGovernance.js";
import { writeAudit } from "../../utils/audit.js";
import { autoCategorizePart } from "../../utils/stockCategory.js";
import { stockInfo } from "../../utils/techPortal.js";

const DUPLICATE_DAYS = 7;

export function ensureDeliverySchema(database = db) {
  database.exec(`
    CREATE TABLE IF NOT EXISTS stock_deliveries (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL DEFAULT 'main',
      supplier TEXT,
      reference TEXT,
      received_date TEXT,
      currency TEXT,
      location_code TEXT,
      notes TEXT,
      line_count INTEGER NOT NULL DEFAULT 0,
      value_usd REAL,
      received_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
  `);
  const cols = database.prepare(`PRAGMA table_info(stock_movements)`).all().map((c) => c.name);
  if (!cols.includes("delivery_id")) database.prepare(`ALTER TABLE stock_movements ADD COLUMN delivery_id INTEGER`).run();
}

export default function registerDeliveryRoutes(app, ctx) {
  const {
    getBinByCodeAtLocation,
    getLocationByCode,
    getOnHand,
    getPartByCode,
    getSiteCode,
    insertGenericMove,
    insertPart,
    requireRoles,
    unitCostToUsd,
    updatePartUnitCostUsd,
  } = ctx;
  ensureDeliverySchema();
  const WRITE_ROLES = ["admin", "supervisor", "stores", "storeman"];

  // Part search for delivery lines: code or name, with on hand, bin and last cost.
  // GET /api/stock/parts/search?q=
  app.get("/parts/search", async (req) => {
    const q = String(req.query?.q || "").trim();
    if (q.length < 2) return { ok: true, rows: [] };
    const like = `%${q.replace(/[%_]/g, (m) => `\\${m}`)}%`;
    const rows = db.prepare(`
      SELECT id, part_code, part_name, unit_cost, COALESCE(min_stock, 0) AS min_stock
      FROM parts
      WHERE part_code LIKE ? ESCAPE '\\' OR part_name LIKE ? ESCAPE '\\'
      ORDER BY CASE WHEN UPPER(part_code) = UPPER(?) THEN 0 WHEN UPPER(part_code) LIKE UPPER(?) THEN 1 ELSE 2 END, part_code
      LIMIT 12
    `).all(like, like, q, `${q}%`);
    const stock = stockInfo(db, rows.map((r) => r.id));
    return {
      ok: true,
      rows: rows.map((r) => ({
        part_code: r.part_code,
        part_name: r.part_name,
        on_hand: stock.get(r.id)?.on_hand ?? 0,
        bin: stock.get(r.id)?.bin || null,
        min_stock: Number(r.min_stock || 0),
        unit_cost_usd: r.unit_cost == null ? null : Number(r.unit_cost),
      })),
    };
  });

  // Recent deliveries for the Receive tab.
  // GET /api/stock/deliveries?days=30
  app.get("/deliveries", async (req) => {
    const days = Math.min(365, Math.max(1, Number(req.query?.days || 30)));
    const rows = db.prepare(`
      SELECT * FROM stock_deliveries
      WHERE datetime(created_at) >= datetime('now', ?)
      ORDER BY id DESC LIMIT 100
    `).all(`-${days} days`);
    const lines = db.prepare(`
      SELECT sm.id AS movement_id, p.part_code, p.part_name, sm.quantity, sm.unit_cost_usd, sm.cost_input, sm.cost_currency
      FROM stock_movements sm JOIN parts p ON p.id = sm.part_id
      WHERE sm.delivery_id = ? ORDER BY sm.id
    `);
    return { ok: true, rows: rows.map((d) => ({ ...d, lines: lines.all(d.id) })) };
  });

  // POST /api/stock/deliveries
  // Body: { supplier?, reference, received_date?, currency?, location_code?, notes?,
  //         lines: [{ part_code, part_name?, quantity, unit_cost?, bin_code? }], confirm_duplicates? }
  app.post("/deliveries", async (req, reply) => {
    if (!requireRoles(req, reply, WRITE_ROLES)) return;
    const site = getSiteCode(req);
    const body = req.body || {};
    const user = String(req.headers["x-user-name"] || "").trim() || null;
    const supplier = String(body.supplier || "").trim().slice(0, 120) || null;
    const reference = String(body.reference || "").trim().slice(0, 80);
    const receivedDate = /^\d{4}-\d{2}-\d{2}$/.test(String(body.received_date || "")) ? body.received_date : new Date().toISOString().slice(0, 10);
    const currency = String(body.currency || "USD").trim().toUpperCase();
    const locationCode = String(body.location_code || "MAIN").trim().toUpperCase();
    const notes = String(body.notes || "").trim().slice(0, 500) || null;
    const rawLines = Array.isArray(body.lines) ? body.lines : [];

    if (!reference) return reply.code(400).send({ ok: false, error: "Enter the invoice or GRN number" });
    if (!["USD", "ZAR", "MZN"].includes(currency)) return reply.code(400).send({ ok: false, error: "Currency must be USD, ZAR or MZN" });
    const location = getLocationByCode.get(locationCode);
    if (!location) return reply.code(404).send({ ok: false, error: `Store not found: ${locationCode}` });

    // Check every line before saving anything.
    const problems = [];
    const lines = [];
    rawLines.forEach((l, i) => {
      const code = String(l?.part_code || "").trim().toUpperCase();
      const name = String(l?.part_name || "").trim();
      const qty = Number(l?.quantity);
      const costRaw = l?.unit_cost == null || l?.unit_cost === "" ? null : Number(l.unit_cost);
      const binCode = String(l?.bin_code || "").trim().toUpperCase();
      if (!code && !qty) return; // empty row
      const n = i + 1;
      if (!code) return problems.push(`Line ${n}: choose a part`);
      if (!Number.isFinite(qty) || qty <= 0) return problems.push(`Line ${n} (${code}): quantity must be more than 0`);
      if (costRaw != null && (!Number.isFinite(costRaw) || costRaw < 0)) return problems.push(`Line ${n} (${code}): unit cost must be a number`);
      const part = getPartByCode.get(code);
      if (!part && !name) return problems.push(`Line ${n}: ${code} is a new part, add its description`);
      const bin = binCode ? getBinByCodeAtLocation.get(Number(location.id), binCode) : null;
      if (binCode && !bin) return problems.push(`Line ${n} (${code}): bin ${binCode} not found in ${locationCode}`);
      let costUsd = null;
      if (costRaw != null && costRaw > 0) {
        costUsd = unitCostToUsd(costRaw, currency);
        if (costUsd == null || !Number.isFinite(costUsd)) return problems.push(`Line ${n}: could not convert ${currency} to USD, check the exchange rates`);
      }
      lines.push({ n, code, name, part, qty, costRaw: costRaw && costRaw > 0 ? costRaw : null, costUsd, bin });
    });
    if (!lines.length && !problems.length) problems.push("Add at least one line");
    const codes = lines.map((l) => l.code);
    const twice = codes.filter((c, i) => codes.indexOf(c) !== i);
    if (twice.length) problems.push(`${[...new Set(twice)].join(", ")} is on more than one line; put the total on one line`);
    if (problems.length) return reply.code(400).send({ ok: false, error: problems[0], problems });

    // New parts must meet the master data rules.
    for (const l of lines.filter((x) => !x.part)) {
      const pol = validateAgainstMdmPolicy(site, "part_stock_intake", {});
      if (!pol.ok) return reply.code(400).send({ ok: false, error: `${l.code}: missing required fields: ${pol.missing.join(", ")}` });
      const pv = validatePartGovernanceOptional(site, {});
      if (!pv.ok) return reply.code(400).send({ ok: false, error: `${l.code}: ${pv.error}` });
    }

    // Possible duplicates: the same invoice already received, or the same part and quantity lately.
    const duplicates = [];
    const sameRef = db.prepare(`
      SELECT id, supplier, received_date, line_count, created_at FROM stock_deliveries
      WHERE UPPER(TRIM(reference)) = UPPER(?) ORDER BY id DESC LIMIT 1
    `).get(reference);
    if (sameRef) duplicates.push({ kind: "invoice", text: `Invoice ${reference} was already received on ${sameRef.received_date || String(sameRef.created_at).slice(0, 10)} (${sameRef.line_count} line${sameRef.line_count === 1 ? "" : "s"}${sameRef.supplier ? `, ${sameRef.supplier}` : ""})` });
    const sameRefMove = !sameRef && db.prepare(`
      SELECT created_at FROM stock_movements WHERE movement_type = 'in' AND UPPER(TRIM(reference)) = UPPER(?) ORDER BY id DESC LIMIT 1
    `).get(reference);
    if (sameRefMove) duplicates.push({ kind: "invoice", text: `Reference ${reference} was already used on a receipt on ${String(sameRefMove.created_at).slice(0, 10)}` });
    const recent = db.prepare(`
      SELECT created_at, reference FROM stock_movements
      WHERE part_id = ? AND movement_type = 'in' AND ABS(quantity - ?) < 0.0001
        AND datetime(created_at) >= datetime('now', ?)
      ORDER BY id DESC LIMIT 1
    `);
    for (const l of lines.filter((x) => x.part)) {
      const r = recent.get(l.part.id, l.qty, `-${DUPLICATE_DAYS} days`);
      if (r) duplicates.push({ kind: "line", part_code: l.code, text: `${l.code}: ${l.qty} already received on ${String(r.created_at).slice(0, 10)}${r.reference ? ` (${r.reference})` : ""}` });
    }
    if (duplicates.length && body.confirm_duplicates !== true) {
      return reply.code(409).send({ ok: false, needs_confirmation: true, error: "This delivery may already be in stock", duplicates });
    }

    const tx = db.transaction(() => {
      const deliveryId = Number(db.prepare(`
        INSERT INTO stock_deliveries (site_code, supplier, reference, received_date, currency, location_code, notes, received_by)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
      `).run(site, supplier, reference, receivedDate, currency, location.location_code, notes, user).lastInsertRowid);
      const out = [];
      let value = 0;
      for (const l of lines) {
        let part = l.part;
        if (!part) {
          const pid = insertPart.run(l.code, l.name, l.costUsd != null ? Number(l.costUsd.toFixed(6)) : 0, null, null, null).lastInsertRowid;
          autoCategorizePart(db, pid);
          part = getPartByCode.get(l.code);
        }
        const before = Number(getOnHand.get(part.id)?.on_hand || 0);
        const mid = Number(insertGenericMove.run(
          part.id, l.qty, "in", reference, Number(location.id), l.bin ? Number(l.bin.id) : null, null,
          l.costUsd != null ? Number(l.costUsd.toFixed(6)) : null, l.costRaw != null ? currency : null, l.costRaw,
        ).lastInsertRowid);
        db.prepare(`UPDATE stock_movements SET delivery_id = ? WHERE id = ?`).run(deliveryId, mid);
        if (l.costUsd != null) updatePartUnitCostUsd.run(Number(l.costUsd.toFixed(6)), part.id);
        const lineValue = l.costUsd != null ? l.qty * l.costUsd : 0;
        value += lineValue;
        out.push({
          movement_id: mid,
          part_code: l.code,
          part_name: part.part_name,
          new_part: !l.part,
          quantity: l.qty,
          on_hand_before: before,
          on_hand_after: before + l.qty,
          unit_cost_usd: l.costUsd != null ? Number(l.costUsd.toFixed(4)) : null,
          line_value_usd: Number(lineValue.toFixed(2)),
          bin_code: l.bin ? l.bin.bin_code : null,
        });
      }
      db.prepare(`UPDATE stock_deliveries SET line_count = ?, value_usd = ? WHERE id = ?`).run(out.length, Number(value.toFixed(2)), deliveryId);
      return { deliveryId, out, value };
    });
    const { deliveryId, out, value } = tx();

    writeAudit(db, req, {
      module: "stock",
      action: "delivery_received",
      entity_type: "stock_delivery",
      entity_id: String(deliveryId),
      payload: { supplier, reference, received_date: receivedDate, lines: out.length, value_usd: Number(value.toFixed(2)), duplicates_confirmed: duplicates.length ? duplicates.map((d) => d.text) : undefined },
    });
    return {
      ok: true,
      delivery_id: deliveryId,
      reference,
      supplier,
      received_date: receivedDate,
      location_code: location.location_code,
      line_count: out.length,
      value_usd: Number(value.toFixed(2)),
      lines: out,
    };
  });
}
