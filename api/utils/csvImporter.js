// IRONLOG/api/utils/csvImporter.js
import { parse } from "csv-parse/sync";

// Detect delimiter from the header line (Excel in many locales uses ';')
function detectDelimiter(text) {
  // Remove BOM for detection safety
  const clean = text.replace(/^\uFEFF/, "");
  const firstLine = clean.split(/\r?\n/).find((l) => l.trim() !== "") || "";

  const commaCount = (firstLine.match(/,/g) || []).length;
  const semiCount = (firstLine.match(/;/g) || []).length;

  // If semicolons are more common, treat as ';' CSV
  return semiCount > commaCount ? ";" : ",";
}

// Accepts a Buffer (file bytes) and returns array of row objects using headers
export function parseCsvToObjects(buffer) {
  const text = buffer.toString("utf8");
  const delimiter = detectDelimiter(text);

  const records = parse(text, {
    columns: true,          // first row is header
    skip_empty_lines: true,
    bom: true,              // handles Excel BOM
    trim: true,
    delimiter              // <-- KEY FIX: supports ';' and ','
  });

  // Normalize header keys (lowercase)
  return records.map((row) => {
    const out = {};
    for (const [k, v] of Object.entries(row)) {
      out[String(k).trim().toLowerCase()] = typeof v === "string" ? v.trim() : v;
    }
    return out;
  });
}

export function requireHeaders(rows, headers = []) {
  if (!rows.length) throw new Error("CSV has no data rows.");

  const keys = new Set(Object.keys(rows[0] || {}));
  const missing = headers.filter((h) => !keys.has(h));
  if (missing.length) {
    throw new Error(`Missing required columns: ${missing.join(", ")}`);
  }
}

export function asInt(val, fallback = 0) {
  const n = Number.parseInt(String(val ?? "").trim(), 10);
  return Number.isFinite(n) ? n : fallback;
}

export function asFloat(val, fallback = 0) {
  const n = Number.parseFloat(String(val ?? "").trim());
  return Number.isFinite(n) ? n : fallback;
}

export function asBool01(val, fallback = 0) {
  const s = String(val ?? "").trim().toLowerCase();
  if (["1", "true", "yes", "y"].includes(s)) return 1;
  if (["0", "false", "no", "n"].includes(s)) return 0;
  return fallback;
}

export function asDateYYYYMMDD(val) {
  const s = String(val ?? "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    throw new Error(`Invalid date (use YYYY-MM-DD): "${s}"`);
  }
  return s;
}