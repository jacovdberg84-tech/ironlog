// IRONLOG/api/utils/assetDetails.js — what a machine is: make, model, serial number, year, purchase and warranty.
//
// These feed reliability reports by model (asset structure part 4) and the
// repair-or-replace answers later. All are optional; a blank clears the value.
// Money is US dollars; dates are YYYY-MM-DD.

export const ASSET_DETAIL_FIELDS = {
  make: { label: "Make", type: "text", max: 60 },
  model: { label: "Model", type: "text", max: 60 },
  serial_number: { label: "Serial number", type: "text", max: 60 },
  year_built: { label: "Year built", type: "year" },
  purchase_date: { label: "Purchase date", type: "date" },
  purchase_price_usd: { label: "Purchase price (USD)", type: "money" },
  warranty_end: { label: "Warranty ends", type: "date" },
};

const SQL_TYPE = { text: "TEXT", year: "INTEGER", date: "TEXT", money: "REAL" };

/** The detail columns on assets (the live database gets them on first start). */
export function ensureAssetDetailColumns(db) {
  let have;
  try {
    have = new Set(db.prepare(`PRAGMA table_info(assets)`).all().map((c) => c.name));
  } catch {
    return;
  }
  if (!have.size) return;
  for (const [col, f] of Object.entries(ASSET_DETAIL_FIELDS)) {
    if (!have.has(col)) db.prepare(`ALTER TABLE assets ADD COLUMN ${col} ${SQL_TYPE[f.type]}`).run();
  }
}

const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(s) && !Number.isNaN(Date.parse(`${s}T00:00:00Z`))
  && new Date(`${s}T00:00:00Z`).toISOString().slice(0, 10) === s;

/**
 * The detail fields present in body, checked: { ok, values: { column: value|null } } or { ok:false, errors }.
 * Only fields sent are returned; "" or null clears a field.
 */
export function readAssetDetails(body = {}, { thisYear = new Date().getUTCFullYear() } = {}) {
  const values = {};
  const errors = [];
  for (const [col, f] of Object.entries(ASSET_DETAIL_FIELDS)) {
    if (!Object.prototype.hasOwnProperty.call(body, col)) continue;
    const raw = body[col] == null ? "" : String(body[col]).trim();
    if (raw === "") {
      values[col] = null;
      continue;
    }
    if (f.type === "text") {
      if (raw.length > f.max) errors.push(`${f.label} can be at most ${f.max} characters.`);
      else values[col] = raw;
    } else if (f.type === "year") {
      const y = Number(raw);
      if (!Number.isInteger(y) || y < 1950 || y > thisYear + 1) errors.push(`${f.label} must be a year between 1950 and ${thisYear + 1}.`);
      else values[col] = y;
    } else if (f.type === "date") {
      if (!isDate(raw)) errors.push(`${f.label} must be a date like 2026-10-16.`);
      else values[col] = raw;
    } else if (f.type === "money") {
      const n = Number(raw.replace(/[$,\s]/g, ""));
      if (!Number.isFinite(n) || n < 0) errors.push(`${f.label} must be an amount of 0 or more.`);
      else values[col] = Math.round(n * 100) / 100;
    }
  }
  if (values.purchase_date && values.warranty_end && values.warranty_end < values.purchase_date) {
    errors.push("Warranty ends before the purchase date.");
  }
  return errors.length ? { ok: false, errors } : { ok: true, values };
}

/** The detail fields of an asset row, for API replies (missing columns read as null). */
export function assetDetailsOf(row = {}) {
  const out = {};
  for (const col of Object.keys(ASSET_DETAIL_FIELDS)) out[col] = row[col] ?? null;
  return out;
}

/** "In warranty, 45 days left" / "Warranty ended 2026-01-31" / null when no date. */
export function warrantyStatus(warrantyEnd, today = new Date().toISOString().slice(0, 10)) {
  if (!warrantyEnd || !isDate(String(warrantyEnd))) return null;
  const days = Math.round((Date.parse(warrantyEnd) - Date.parse(today)) / 864e5);
  return days >= 0
    ? { in_warranty: true, days_left: days, text: `In warranty, ${days} day${days === 1 ? "" : "s"} left` }
    : { in_warranty: false, days_left: days, text: `Warranty ended ${warrantyEnd}` };
}

/** "CAT 140K · S/N CAT0140KJ2501 · 2019" for headers (empty when nothing is known). */
export function assetDetailLine(row = {}) {
  return [
    [row.make, row.model].filter(Boolean).join(" "),
    row.serial_number ? `S/N ${row.serial_number}` : "",
    row.year_built ? String(row.year_built) : "",
  ].filter(Boolean).join(" · ");
}
