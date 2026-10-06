// IRONLOG/api/utils/lubeModel.js — the monthly lube costing model (LUBE TEMPLATE .xlsm).
//
// The site submits a macro workbook each month. Its input cells are:
//   1. Control_Sheet       C7 year, C8 period, D16:D65 issue-book ref, F16:F65 day of month
//   INPUT 1 … INPUT 50     one issue day each, rows 6–55: D plant, E oil type, F qty, G reading, H reason
//   3.deliveries to site   one row per invoice line (kept from month to month): E date, G supplier,
//                          K description, L oil type, N litres per item, O items, Q invoice, R price, U ROE
//   4.list_of_site_lube_suppliers  A supplier names (feeds the supplier drop-down)
//   5.previous month closing stock B site description per standard type, N description, P qty
// IronLog fills those cells from its lube issues and receipts; everything else
// in the workbook (formulas, macros, pivots, protection) is left as it is.

import fs from "node:fs";
import path from "node:path";
import { getDataRoot } from "./storagePaths.js";
import { fetchLubeUsageLines } from "./lubeUsageLines.js";
import { ensureStockCategorySchema, oilPartSql } from "./stockCategory.js";
import { XlsxPatcher, excelDate } from "./xlsxPatch.js";

export const MODEL_OIL_TYPES = [
  "COOLANT", "ENGINE 10W40", "ENGINE 15W40", "FINAL DRIVES 20W40", "GEARBOX/F DRIVE 10W30",
  "GEARBOX/F DRIVE 80W90", "GEARBOX/F DRIVE 85W140", "GEARBOX/F DRIVE 85W90", "HYDRAULIC", "ATF",
  "TRANSMISSION 30W", "TRANSMISSION 50W", "TRANSMISSION PASSENGER 75W90", "GREASE", "SLEW GREASE",
  "WDB", "BRAKE FLUID",
];
const SERVICE_REASONS = new Set(["250", "500", "750", "1000"]);
const INPUT_SHEETS = 50;
const ROWS_PER_SHEET = 50; // rows 6–55
const CONTROL = "1. Control_Sheet";
const DELIVERIES = "3.deliveries to site";
const SUPPLIERS = "4.list_of_site_lube_suppliers";
const OPENING = "5.previous month closing stock";

/** The model's oil type for a stock item, from its name (null when unsure). */
export function suggestOilType(name, code = "") {
  const s = `${name || ""} ${code || ""}`.toUpperCase().replace(/[-_]/g, "").replace(/\s+/g, " ");
  const has = (re) => re.test(s);
  if (has(/COOLANT|ANTIFREEZE|ANTI FREEZE|FRICTOLIN|MAINTAIN ?LIFE/)) return "COOLANT";
  if (has(/BRAKE ?FLUID|DOT ?[345]/)) return "BRAKE FLUID";
  if (has(/WD ?40|WDB/)) return "WDB";
  if (has(/SLEW/)) return "SLEW GREASE";
  if (has(/GREASE|\bEP ?[0-3]\b|LMX|LITHIUM/)) return "GREASE";
  if (has(/\bATF\b|DEXRON|AUTOMATIC TRANS/)) return "ATF";
  if (has(/75W ?90/)) return "TRANSMISSION PASSENGER 75W90";
  if (has(/TO ?4.*(SAE ?)?50|SAE ?50|\b50W\b/)) return "TRANSMISSION 50W";
  if (has(/TO ?4.*(SAE ?)?30|SAE ?30|\b30W\b/)) return "TRANSMISSION 30W";
  if (has(/HYDRAUL|RENOLIN|\bHO ?(32|46|68)\b|\bHLP\b|\bAW ?(46|68)\b|TELLUS/)) return "HYDRAULIC";
  if (has(/85W ?140/)) return "GEARBOX/F DRIVE 85W140";
  if (has(/85W ?90/)) return "GEARBOX/F DRIVE 85W90";
  if (has(/80W ?90/)) return "GEARBOX/F DRIVE 80W90";
  if (has(/10W ?30/)) return "GEARBOX/F DRIVE 10W30";
  if (has(/20W ?40/)) return "FINAL DRIVES 20W40";
  if (has(/15W ?40/)) return "ENGINE 15W40";
  if (has(/10W ?40/)) return "ENGINE 10W40";
  return null;
}

/** Litres (or kg) in one stock unit, read from the name: "20L", "208 LT", "18KG". */
export function packSizeFromName(name) {
  const m = /(\d+(?:[.,]\d+)?)\s*(L|LT|LTR|LTRS|LITRE|LITRES|LITER|LITERS|KG|KGS)\b/i.exec(String(name || ""));
  if (!m) return null;
  const n = Number(m[1].replace(",", "."));
  return Number.isFinite(n) && n > 0 ? n : null;
}

export function ensureLubeModelSchema(db) {
  // The lube usage lines read oil_logs.unit_cost (also added by the reports routes).
  const oilCols = db.prepare(`PRAGMA table_info(oil_logs)`).all().map((c) => c.name);
  if (oilCols.length && !oilCols.includes("unit_cost")) db.prepare(`ALTER TABLE oil_logs ADD COLUMN unit_cost REAL`).run();
  db.exec(`
    CREATE TABLE IF NOT EXISTS lube_model_parts (
      part_id INTEGER PRIMARY KEY,
      oil_type TEXT,
      pack_size REAL,
      updated_by TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS lube_model_template (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      file_name TEXT,
      stored_path TEXT,
      size_bytes INTEGER,
      uploaded_by TEXT,
      uploaded_at TEXT
    );
  `);
}

export function templateDir() {
  const dir = path.join(getDataRoot(), "uploads", "lube-model");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

export function templateInfo(db) {
  ensureLubeModelSchema(db);
  const t = db.prepare(`SELECT * FROM lube_model_template WHERE id = 1`).get();
  return t && t.stored_path && fs.existsSync(t.stored_path) ? t : null;
}

/** Checks a workbook is the lube model before it is kept. */
export async function checkTemplate(buffer) {
  const p = await XlsxPatcher.open(buffer);
  const missing = [CONTROL, DELIVERIES, OPENING, "INPUT 1", `INPUT ${INPUT_SHEETS}`].filter((n) => !p.hasSheet(n));
  if (missing.length) throw Object.assign(new Error(`This is not the lube model: sheet${missing.length === 1 ? "" : "s"} missing: ${missing.join(", ")}`), { status: 400 });
  return true;
}

/** Every lube stock item with its model oil type and container size (saved or suggested). */
export function lubeParts(db) {
  ensureLubeModelSchema(db);
  ensureStockCategorySchema(db);
  return db.prepare(`
    SELECT p.id, p.part_code, p.part_name, lm.oil_type, lm.pack_size
    FROM parts p LEFT JOIN lube_model_parts lm ON lm.part_id = p.id
    WHERE ${oilPartSql("p")}
      OR p.id IN (SELECT part_id FROM oil_logs WHERE part_id IS NOT NULL)
      OR lm.part_id IS NOT NULL
    ORDER BY p.part_code
  `).all().map((p) => {
    const suggestedType = suggestOilType(p.part_name, p.part_code);
    const suggestedPack = packSizeFromName(p.part_name) ?? packSizeFromName(p.part_code);
    return {
      id: p.id,
      part_code: p.part_code,
      part_name: p.part_name,
      oil_type: p.oil_type || suggestedType,
      oil_type_saved: Boolean(p.oil_type),
      pack_size: p.pack_size ?? suggestedPack ?? 1,
      pack_size_saved: p.pack_size != null,
      pack_size_found: p.pack_size != null || suggestedPack != null,
    };
  });
}

export function savePartMapping(db, partCode, { oil_type, pack_size }, user) {
  ensureLubeModelSchema(db);
  const part = db.prepare(`SELECT id FROM parts WHERE UPPER(part_code) = UPPER(?)`).get(String(partCode || ""));
  if (!part) throw Object.assign(new Error("Stock item not found"), { status: 404 });
  const type = oil_type == null || oil_type === "" ? null : String(oil_type).toUpperCase();
  if (type && !MODEL_OIL_TYPES.includes(type)) throw Object.assign(new Error(`Oil type must be one of the model's types`), { status: 400 });
  const pack = pack_size == null || pack_size === "" ? null : Number(pack_size);
  if (pack != null && (!Number.isFinite(pack) || pack <= 0)) throw Object.assign(new Error("Container size must be more than 0"), { status: 400 });
  db.prepare(`
    INSERT INTO lube_model_parts (part_id, oil_type, pack_size, updated_by, updated_at) VALUES (?, ?, ?, ?, datetime('now'))
    ON CONFLICT(part_id) DO UPDATE SET oil_type = excluded.oil_type, pack_size = excluded.pack_size,
      updated_by = excluded.updated_by, updated_at = excluded.updated_at
  `).run(part.id, type, pack, user || null);
}

const pad = (n) => String(n).padStart(2, "0");
function monthRange(year, month) {
  const last = new Date(Date.UTC(year, month, 0)).getUTCDate();
  return { start: `${year}-${pad(month)}-01`, end: `${year}-${pad(month)}-${pad(last)}` };
}

/** Service interval for the reason column: "500HR_SERVICE", "SERVICE" or "TOP-UP". */
function reasonFor(db, line) {
  const name = (row) => {
    const n = String(row?.service_name || "").trim().replace(/\s*h(rs?)?\s*$/i, "");
    return SERVICE_REASONS.has(n) ? `${n}HR_SERVICE` : "SERVICE";
  };
  if (line.work_order_id) {
    const wo = db.prepare(`
      SELECT w.source, mp.service_name FROM work_orders w
      LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id AND w.source = 'service'
      WHERE w.id = ?
    `).get(line.work_order_id);
    if (wo?.source === "service") return name(wo);
  }
  if (line.asset_id) {
    // A lube issue on the day of a service on that machine counts as the service.
    const wo = db.prepare(`
      SELECT mp.service_name FROM work_orders w
      LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id
      WHERE w.asset_id = ? AND w.source = 'service'
        AND (date(COALESCE(w.completed_at, w.closed_at, w.started_at, w.opened_at)) BETWEEN date(?, '-1 day') AND date(?, '+1 day'))
      ORDER BY w.id DESC LIMIT 1
    `).get(line.asset_id, line.usage_date, line.usage_date);
    if (wo) return name(wo);
  }
  return "TOP-UP";
}

/**
 * What goes into the model for a month: issue days, deliveries, opening stock,
 * and what could not be mapped.
 */
export function collectMonth(db, year, month) {
  const { start, end } = monthRange(year, month);
  const parts = lubeParts(db);
  const byCode = new Map(parts.map((p) => [String(p.part_code).toUpperCase(), p]));
  const byId = new Map(parts.map((p) => [p.id, p]));
  const unmapped = new Map();
  const note = (p, why) => unmapped.set(p.part_code, { part_code: p.part_code, part_name: p.part_name, why });

  // Issues
  const lines = fetchLubeUsageLines(db, { start, end }).filter((l) => l.asset_code && Number(l.quantity) > 0);
  const issues = [];
  for (const l of lines) {
    const p = byCode.get(String(l.part_code || "").toUpperCase());
    const type = p?.oil_type || suggestOilType(l.part_name, l.part_code);
    if (!type) {
      note(p || { part_code: l.part_code, part_name: l.part_name }, "no oil type");
      continue;
    }
    const pack = p?.pack_size ?? packSizeFromName(l.part_name) ?? 1;
    issues.push({
      date: l.usage_date,
      day: Number(String(l.usage_date).slice(8, 10)),
      plant: l.asset_code,
      oil_type: type,
      qty: Number((Number(l.quantity) * pack).toFixed(2)),
      reading: l.smr != null ? Number(l.smr) : null,
      reason: reasonFor(db, l),
      part_code: l.part_code,
    });
  }
  // One INPUT sheet per day (more when a day has over 50 lines).
  const days = [];
  for (const d of [...new Set(issues.map((i) => i.day))].sort((a, b) => a - b)) {
    const rows = issues.filter((i) => i.day === d);
    for (let k = 0; k < rows.length; k += ROWS_PER_SHEET) days.push({ day: d, rows: rows.slice(k, k + ROWS_PER_SHEET) });
  }

  // Deliveries (receipts of lube items this month)
  const hasDeliveries = db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'stock_deliveries'`).get();
  const hasDeliveryCol = db.prepare(`PRAGMA table_info(stock_movements)`).all().some((c) => c.name === "delivery_id");
  const rate = (key, fallback) => {
    try {
      const v = Number(db.prepare(`SELECT value FROM cost_settings WHERE key = ?`).get(key)?.value);
      return Number.isFinite(v) && v > 0 ? v : fallback;
    } catch {
      return fallback;
    }
  };
  const rates = { mzn_per_usd: rate("mzn_per_usd", 64), zar_per_usd: rate("zar_per_usd", 18.5) };
  const receipts = db.prepare(`
    SELECT sm.id, sm.part_id, sm.quantity, sm.reference, sm.unit_cost_usd, sm.cost_currency, sm.cost_input,
      date(sm.created_at) AS moved_on
      ${hasDeliveries && hasDeliveryCol ? ", d.supplier, d.received_date, d.reference AS invoice" : ", NULL AS supplier, NULL AS received_date, NULL AS invoice"}
    FROM stock_movements sm
    ${hasDeliveries && hasDeliveryCol ? "LEFT JOIN stock_deliveries d ON d.id = sm.delivery_id" : ""}
    WHERE sm.movement_type = 'in' AND sm.quantity > 0
      AND date(COALESCE(${hasDeliveries && hasDeliveryCol ? "d.received_date, " : ""}sm.created_at)) BETWEEN ? AND ?
    ORDER BY sm.id
  `).all(start, end);
  const deliveries = [];
  for (const r of receipts) {
    const p = byId.get(r.part_id);
    if (!p) continue; // not a lube item
    if (!p.oil_type) {
      note(p, "no oil type");
      continue;
    }
    const currency = String(r.cost_currency || "USD").toUpperCase();
    const roe = currency === "MZN" ? rates.mzn_per_usd : currency === "ZAR" ? rates.zar_per_usd : 1;
    const unitLocal = r.cost_input != null ? Number(r.cost_input) : r.unit_cost_usd != null ? Number(r.unit_cost_usd) : null;
    deliveries.push({
      date: r.received_date || r.moved_on,
      supplier: r.supplier || null,
      description: p.part_name || p.part_code,
      oil_type: p.oil_type,
      pack_size: p.pack_size,
      items: Number(r.quantity),
      invoice: r.invoice || (r.reference && !/^(opening|manual_entry)$/i.test(r.reference) ? r.reference : null),
      total_price: unitLocal != null ? Number((unitLocal * Number(r.quantity)).toFixed(2)) : null,
      roe: r.cost_input != null ? roe : 1,
    });
  }

  // Opening stock: on hand at the start of the month, in litres, per oil type.
  const opening = new Map();
  for (const r of db.prepare(`
    SELECT part_id, COALESCE(SUM(quantity), 0) AS q FROM stock_movements
    WHERE datetime(created_at) < datetime(?) GROUP BY part_id
  `).all(start)) {
    const p = byId.get(r.part_id);
    if (!p || !p.oil_type || !(Number(r.q) > 0)) continue;
    opening.set(p.oil_type, Number(((opening.get(p.oil_type) || 0) + Number(r.q) * p.pack_size).toFixed(2)));
  }

  return {
    year, month, start, end,
    days,
    issue_lines: issues.length,
    issue_litres: Number(issues.reduce((s, i) => s + i.qty, 0).toFixed(2)),
    deliveries,
    opening: [...opening.entries()].map(([oil_type, qty]) => ({ oil_type, qty })),
    unmapped: [...unmapped.values()],
    too_many_days: days.length > INPUT_SHEETS,
  };
}

/** Fills the stored template for the month; returns the .xlsm bytes. */
export async function fillTemplate(templateBuffer, data) {
  const x = await XlsxPatcher.open(templateBuffer);

  // Control sheet: year, period, issue days (and no issue-book references).
  const control = { C7: data.year, C8: data.month };
  for (let i = 0; i < INPUT_SHEETS; i += 1) {
    const row = 16 + i;
    const day = data.days[i];
    control[`D${row}`] = null;
    control[`F${row}`] = day ? day.day : null;
  }
  await x.setCells(CONTROL, control);

  // INPUT 1–50: cleared, then one day each.
  for (let s = 1; s <= INPUT_SHEETS; s += 1) {
    const name = `INPUT ${s}`;
    if (!x.hasSheet(name)) continue;
    const cells = {};
    const day = data.days[s - 1];
    for (let k = 0; k < ROWS_PER_SHEET; k += 1) {
      const row = 6 + k;
      const r = day?.rows[k];
      cells[`C${row}`] = null;
      cells[`D${row}`] = r ? r.plant : null;
      cells[`E${row}`] = r ? r.oil_type : null;
      cells[`F${row}`] = r ? r.qty : null;
      cells[`G${row}`] = r && r.reading != null ? r.reading : null;
      cells[`H${row}`] = r ? r.reason : null;
    }
    await x.setCells(name, cells);
  }

  // Deliveries: added under the lines already there; a line already on the sheet is not added twice.
  const existing = await x.readColumns(DELIVERIES, ["E", "K", "O", "Q"]);
  const seen = new Set();
  let lastUsed = 3;
  for (const [row, v] of existing) {
    if (row < 4 || v.E == null) continue;
    lastUsed = Math.max(lastUsed, row);
    seen.add(`${String(v.Q ?? "").trim().toUpperCase()}|${String(v.K ?? "").trim().toUpperCase()}|${Number(v.O)}`);
  }
  let next = lastUsed + 1;
  const dcells = {};
  let added = 0;
  for (const d of data.deliveries) {
    const key = `${String(d.invoice ?? "").trim().toUpperCase()}|${String(d.description).trim().toUpperCase()}|${Number(d.items)}`;
    if (seen.has(key) || next > 1001) continue;
    seen.add(key);
    Object.assign(dcells, {
      [`E${next}`]: excelDate(d.date),
      [`G${next}`]: d.supplier,
      [`K${next}`]: d.description,
      [`L${next}`]: d.oil_type,
      [`N${next}`]: d.pack_size,
      [`O${next}`]: d.items,
      [`Q${next}`]: d.invoice,
      [`R${next}`]: d.total_price,
      [`U${next}`]: d.roe,
    });
    next += 1;
    added += 1;
  }
  if (added) await x.setCells(DELIVERIES, dcells);

  // Suppliers named on deliveries join the supplier drop-down list.
  if (x.hasSheet(SUPPLIERS)) {
    const list = await x.readColumns(SUPPLIERS, ["A"]);
    const names = new Set([...list.values()].map((v) => String(v.A).trim().toUpperCase()));
    let row = Math.max(2, ...list.keys()) + 1;
    const scells = {};
    for (const s of [...new Set(data.deliveries.map((d) => d.supplier).filter(Boolean))]) {
      if (names.has(s.toUpperCase())) continue;
      scells[`A${row}`] = s;
      names.add(s.toUpperCase());
      row += 1;
    }
    if (Object.keys(scells).length) await x.setCells(SUPPLIERS, scells);
  }

  // Opening stock: IronLog's closing stock per oil type, linked to the standard mapping rows.
  // With no stock history in IronLog the sheet is left as it is in the template.
  if (!data.opening.length) return { buffer: await x.save(), deliveries_added: added, opening_filled: false };
  const mapRows = await x.readColumns(OPENING, ["C"]);
  const ocells = {};
  for (let r = 4; r <= 34; r += 1) {
    ocells[`N${r}`] = null;
    ocells[`P${r}`] = null;
  }
  data.opening.forEach((o, i) => {
    if (i > 30) return;
    ocells[`N${4 + i}`] = o.oil_type;
    ocells[`P${4 + i}`] = o.qty;
  });
  for (const [row, v] of mapRows) {
    if (row < 4 || row > 40) continue;
    const t = String(v.C || "").trim().toUpperCase();
    if (MODEL_OIL_TYPES.includes(t)) ocells[`B${row}`] = t;
  }
  await x.setCells(OPENING, ocells);

  return { buffer: await x.save(), deliveries_added: added, opening_filled: true };
}
