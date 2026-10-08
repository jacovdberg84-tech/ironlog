// IRONLOG/api/utils/partBarcodes.js — suppliers' barcodes linked to stock items.
//
// Boxes arrive with the maker's barcode (EAN, UPC, Code 128 …). The stores
// links each one to an IronLog part once; after that a scan of the box finds
// the part. One part can have several barcodes (different brands or suppliers).

const ready = new WeakSet();

export function ensurePartBarcodeSchema(db) {
  if (ready.has(db)) return;
  db.exec(`
    CREATE TABLE IF NOT EXISTS part_barcodes (
      barcode TEXT PRIMARY KEY,
      part_id INTEGER NOT NULL,
      linked_by TEXT,
      linked_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_part_barcodes_part ON part_barcodes(part_id);
  `);
  ready.add(db);
}

/** Barcodes are stored trimmed and upper-case; scanners differ in case for Code 39/128. */
export function normalizeBarcode(raw) {
  return String(raw || "").trim().replace(/\s+/g, "").toUpperCase().slice(0, 120);
}

export function partByBarcode(db, raw) {
  ensurePartBarcodeSchema(db);
  const code = normalizeBarcode(raw);
  if (!code) return null;
  return db.prepare(`
    SELECT p.id, p.part_code, p.part_name, p.min_stock
    FROM part_barcodes b JOIN parts p ON p.id = b.part_id
    WHERE b.barcode = ?
  `).get(code) || null;
}

export function barcodesForPart(db, partId) {
  ensurePartBarcodeSchema(db);
  return db.prepare(`SELECT barcode, linked_by, linked_at FROM part_barcodes WHERE part_id = ? ORDER BY linked_at`).all(partId);
}

/**
 * Link a barcode to a part. Throws { status: 409, linked_to } when it already
 * belongs to another part and `replace` is not set.
 */
export function linkBarcode(db, { barcode, part, user = null, replace = false }) {
  ensurePartBarcodeSchema(db);
  const code = normalizeBarcode(barcode);
  if (code.length < 3) throw Object.assign(new Error("Scan the barcode on the box"), { status: 400 });
  const clash = db.prepare(`SELECT 1 FROM parts WHERE UPPER(part_code) = ? AND id <> ?`).get(code, part.id)
    || db.prepare(`SELECT 1 FROM assets WHERE UPPER(asset_code) = ?`).get(code);
  if (clash) throw Object.assign(new Error(`${code} is already a part or machine code in IronLog`), { status: 400 });
  const current = db.prepare(`
    SELECT b.part_id, p.part_code, p.part_name FROM part_barcodes b JOIN parts p ON p.id = b.part_id WHERE b.barcode = ?
  `).get(code);
  if (current && Number(current.part_id) === Number(part.id)) return { barcode: code, already: true };
  if (current && !replace) {
    throw Object.assign(new Error(`This barcode is linked to ${current.part_code}`), {
      status: 409,
      linked_to: { part_code: current.part_code, part_name: current.part_name },
    });
  }
  db.prepare(`
    INSERT INTO part_barcodes (barcode, part_id, linked_by, linked_at) VALUES (?, ?, ?, datetime('now'))
    ON CONFLICT(barcode) DO UPDATE SET part_id = excluded.part_id, linked_by = excluded.linked_by, linked_at = excluded.linked_at
  `).run(code, part.id, user);
  return { barcode: code, moved_from: current ? current.part_code : null };
}

export function unlinkBarcode(db, raw) {
  ensurePartBarcodeSchema(db);
  return db.prepare(`DELETE FROM part_barcodes WHERE barcode = ?`).run(normalizeBarcode(raw)).changes > 0;
}
