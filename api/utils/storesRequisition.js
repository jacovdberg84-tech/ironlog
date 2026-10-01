// IRONLOG/api/utils/storesRequisition.js — printable stores requisitions.
//
// A requisition groups parts requests (maintenance_parts_requests, the same
// lines stores already see in "Workshop waiting on parts") onto one signed
// paper slip: who asked, for which machine / work order, what, how many, where
// it is in the stores. Requests from the portal or a work order are printed
// per work order; people who ask at the counter are captured with the walk-up
// form, which records the same request lines first. Reprinting keeps the
// number. Labels are English + Portuguese.

import { buildPdfBuffer, ensurePageSpace, table } from "./pdfGenerator.js";
import { stockInfo } from "./techPortal.js";

function hasColumn(db, table, col) {
  return db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
}

export function ensureRequisitionSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS stores_requisitions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL DEFAULT 'main',
      work_order_id INTEGER,
      asset_id INTEGER,
      requested_by TEXT,
      walk_up INTEGER NOT NULL DEFAULT 0,
      notes TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      printed_count INTEGER NOT NULL DEFAULT 0,
      last_printed_at TEXT
    );
  `);
  const hasRequests = db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'maintenance_parts_requests'`).get();
  if (hasRequests && !hasColumn(db, "maintenance_parts_requests", "requisition_id")) {
    db.prepare(`ALTER TABLE maintenance_parts_requests ADD COLUMN requisition_id INTEGER`).run();
  }
}

export const requisitionNumber = (id) => `SR-${String(id).padStart(6, "0")}`;

/**
 * Puts the given request lines on a requisition. Lines already on one keep it;
 * when every line is already on the same requisition that one is returned
 * (reprint). Otherwise the free lines get a new requisition.
 */
export function requisitionFor(db, requestIds, { site = "main", createdBy = null, requestedBy = null, walkUp = false, notes = null } = {}) {
  const ids = [...new Set(requestIds.map(Number).filter((n) => n > 0))];
  if (!ids.length) throw Object.assign(new Error("Choose at least one request line"), { status: 400 });
  const rows = db.prepare(`
    SELECT id, work_order_id, asset_id, requested_by, requisition_id FROM maintenance_parts_requests
    WHERE id IN (${ids.map(() => "?").join(", ")})
  `).all(...ids);
  if (rows.length !== ids.length) throw Object.assign(new Error("Some request lines were not found"), { status: 404 });
  const free = rows.filter((r) => !r.requisition_id);
  if (!free.length) return { id: rows[0].requisition_id, created: false };
  const wo = [...new Set(free.map((r) => r.work_order_id).filter(Boolean))];
  const assets = [...new Set(free.map((r) => r.asset_id).filter(Boolean))];
  const who = requestedBy || [...new Set(free.map((r) => r.requested_by).filter(Boolean))].join(", ") || null;
  const ins = db.prepare(`
    INSERT INTO stores_requisitions (site_code, work_order_id, asset_id, requested_by, walk_up, notes, created_by)
    VALUES (?, ?, ?, ?, ?, ?, ?)
  `).run(site, wo.length === 1 ? wo[0] : null, assets.length === 1 ? assets[0] : null, who, walkUp ? 1 : 0, notes, createdBy);
  const id = Number(ins.lastInsertRowid);
  const upd = db.prepare(`UPDATE maintenance_parts_requests SET requisition_id = ? WHERE id = ?`);
  for (const r of free) upd.run(id, r.id);
  return { id, created: true };
}

/** One requisition with its lines, machine, work order and stores stock. */
export function getRequisition(db, id) {
  const r = db.prepare(`SELECT * FROM stores_requisitions WHERE id = ?`).get(Number(id));
  if (!r) return null;
  const lines = db.prepare(`
    SELECT pr.id, pr.part_code, pr.part_name, pr.qty, pr.urgency, pr.status, pr.notes, pr.requested_by, pr.created_at,
      pr.work_order_id, pr.asset_code, p.id AS part_id, COALESCE(p.part_name, pr.part_name) AS description
    FROM maintenance_parts_requests pr
    LEFT JOIN parts p ON LOWER(p.part_code) = LOWER(pr.part_code)
    WHERE pr.requisition_id = ? ORDER BY pr.id
  `).all(r.id);
  const stock = stockInfo(db, lines.map((l) => l.part_id).filter(Boolean));
  const asset = r.asset_id ? db.prepare(`SELECT asset_code, asset_name FROM assets WHERE id = ?`).get(r.asset_id) : null;
  const wo = r.work_order_id
    ? db.prepare(`SELECT * FROM work_orders WHERE id = ?`).get(r.work_order_id)
    : null;
  return {
    ...r,
    number: requisitionNumber(r.id),
    asset,
    work_order: wo,
    lines: lines.map((l) => ({
      ...l,
      qty: Number(l.qty),
      on_hand: l.part_id ? stock.get(l.part_id)?.on_hand ?? 0 : null,
      bin: l.part_id ? stock.get(l.part_id)?.bin || null : null,
    })),
  };
}

export function listRequisitions(db, { site = null, days = 30 } = {}) {
  return db.prepare(`
    SELECT r.id, r.created_at, r.requested_by, r.work_order_id, r.walk_up, r.printed_count, a.asset_code,
      (SELECT COUNT(*) FROM maintenance_parts_requests pr WHERE pr.requisition_id = r.id) AS line_count
    FROM stores_requisitions r LEFT JOIN assets a ON a.id = r.asset_id
    WHERE datetime(r.created_at) >= datetime('now', ?) ${site ? "AND LOWER(r.site_code) = LOWER(?)" : ""}
    ORDER BY r.id DESC LIMIT 200
  `).all(...(site ? [`-${Number(days)} days`, site] : [`-${Number(days)} days`])).map((r) => ({ ...r, number: requisitionNumber(r.id) }));
}

const L = (en, pt) => `${en} / ${pt}`;

/** Printable requisition (A4 portrait) with sign-off boxes. */
export function buildRequisitionPdf(db, req) {
  return buildPdfBuffer((doc) => {
    const left = doc.page.margins.left;
    const width = doc.page.width - left - doc.page.margins.right;
    doc.y = Math.max(doc.y, 110);

    doc.font("Helvetica-Bold").fontSize(16).fillColor("#0f172a").text(L("Stores Requisition", "Requisição de Armazém"), left, doc.y, { width });
    doc.font("Helvetica-Bold").fontSize(13).fillColor("#0b3a7e").text(req.number, left, doc.y + 2, { width });
    doc.moveDown(0.6);

    const machine = req.asset ? `${req.asset.asset_code}${req.asset.asset_name ? ` — ${req.asset.asset_name}` : ""}` : (req.lines[0]?.asset_code || "—");
    const facts = [
      [L("Date", "Data"), String(req.created_at || "").slice(0, 16)],
      [L("Requested by", "Pedido por"), req.requested_by || "—"],
      [L("Machine", "Máquina"), machine],
      [L("Work order", "Ordem de trabalho"), req.work_order ? `#${req.work_order.id}` : "—"],
    ];
    if (req.walk_up) facts.push([L("Raised at", "Feito em"), L("Stores counter", "Balcão do armazém")]);
    if (req.notes) facts.push([L("Notes", "Notas"), req.notes]);
    doc.fontSize(10);
    for (const [k, v] of facts) {
      const y = doc.y;
      doc.font("Helvetica-Bold").fillColor("#334155").text(`${k}:`, left, y, { width: 170 });
      doc.font("Helvetica").fillColor("#0f172a").text(String(v), left + 175, y, { width: width - 175 });
      doc.y = Math.max(doc.y, y + 14);
    }
    doc.moveDown(0.8);

    table(doc, [
      { key: "n", label: "#", width: 0.05, align: "right" },
      { key: "code", label: "Code/Cód.", width: 0.14 },
      { key: "desc", label: "Description/Descrição", width: 0.32 },
      { key: "qty", label: "Qty/Qtd", width: 0.1, align: "right" },
      { key: "stock", label: "Stock", width: 0.09, align: "right" },
      { key: "bin", label: "Bin/Prat.", width: 0.13 },
      { key: "issued", label: "Issued/Entregue", width: 0.17, align: "right" },
    ], req.lines.map((l, i) => ({
      n: i + 1,
      code: l.part_code || "—",
      desc: `${l.description || l.part_name || ""}${l.urgency && l.urgency !== "normal" ? ` (${l.urgency})` : ""}`,
      qty: l.qty,
      stock: l.on_hand == null ? "?" : l.on_hand,
      bin: l.bin || "",
      issued: "",
    })), { compact: false, fontSize: 9, headerFontSize: 8, rowPadY: 7 });

    // Sign-off boxes.
    doc.moveDown(1.2);
    ensurePageSpace(doc, 150);
    const boxes = [
      L("Requested by", "Pedido por"),
      L("Authorised by", "Autorizado por"),
      L("Issued by (stores)", "Entregue por (armazém)"),
      L("Received by", "Recebido por"),
    ];
    const gap = 10;
    const bw = (width - gap) / 2;
    const bh = 58;
    const top = doc.y;
    boxes.forEach((label, i) => {
      const x = left + (i % 2) * (bw + gap);
      const y = top + Math.floor(i / 2) * (bh + gap);
      doc.save().lineWidth(0.8).strokeColor("#94a3b8").rect(x, y, bw, bh).stroke().restore();
      doc.font("Helvetica-Bold").fontSize(8.5).fillColor("#334155").text(label, x + 6, y + 5, { width: bw - 12 });
      doc.font("Helvetica").fontSize(8).fillColor("#64748b")
        .text(L("Name", "Nome"), x + 6, y + 22)
        .text(L("Signature", "Assinatura"), x + 6, y + 36)
        .text(L("Date", "Data"), x + bw * 0.62, y + 36);
    });
    doc.y = top + 2 * (bh + gap) + 6;
    doc.font("Helvetica").fontSize(8).fillColor("#64748b").text(
      L("Stores issue against this requisition and record the issue on the work order in the system.", "O armazém entrega contra esta requisição e regista a saída na ordem de trabalho no sistema."),
      left, doc.y, { width },
    );
  }, {
    db,
    layout: "portrait",
    title: `Stores Requisition ${req.number}`,
    subtitle: L("Stores Requisition", "Requisição de Armazém"),
    rightText: req.number,
  });
}
