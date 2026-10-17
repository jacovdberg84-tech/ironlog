// IRONLOG/api/utils/goLiveImport.js — the guided go-live import: one Excel workbook for a site's starting data.
//
// A new customer (or a new site) fills in one workbook: Machines, Parts, Stock on hand,
// Suppliers and Components. The workbook can be downloaded empty or filled with what
// IronLog already holds, so it can be corrected and uploaded again. Upload first runs a
// check that changes nothing and lists every problem by sheet, row and column, with how
// many records are new, changed or the same. Import runs the same check and, only when
// it is clean, applies everything in one transaction.
//
// Rules: codes are the key (asset_code, part_code, supplier_code; component + position
// per machine), so uploading the same file twice changes nothing. A blank cell clears
// an optional detail, except meter hours and stock, where blank means "leave as is".
// Stock on hand sets the balance: IronLog posts the difference as a go-live movement.
// Machines go to the site in the Site column (or the site chosen on screen) and only to
// sites the person may use. Money is US dollars; dates are YYYY-MM-DD.
import { ASSET_DETAIL_FIELDS, ensureAssetDetailColumns, readAssetDetails } from "./assetDetails.js";
import { addAssetComponent, ensureAssetComponentTables, readComponentFields, updateAssetComponent } from "./assetComponents.js";
import { listFailureCodes } from "./failureCodes.js";
import { ensureStockCategorySchema, normalizeStockCategory, STOCK_CATEGORIES } from "./stockCategory.js";
import { knownSites } from "./sites.js";

const YES_NO = ["yes", "no"];

// Each sheet: its columns (key, header, required, help). Headers are matched without regard to case.
export const GO_LIVE_SHEETS = [
  {
    key: "machines", name: "Machines", help: "One row per machine. The machine code is the key: a code IronLog already has is updated.",
    columns: [
      { key: "asset_code", header: "Machine code", required: true, help: "e.g. A302AM. Must be unique." },
      { key: "asset_name", header: "Machine name", required: true, help: "e.g. Bell B30D ADT" },
      { key: "category", header: "Category", required: true, help: "e.g. ADT, Excavator, Wheel loader, LDV" },
      { key: "site_code", header: "Site", help: "Site code, e.g. main. Blank = the site chosen when uploading." },
      { key: "is_standby", header: "Standby", help: "yes or no (blank = no)", list: YES_NO },
      { key: "area", header: "Area", help: "e.g. Pit, Plant, Workshop" },
      { key: "make", header: "Make", help: "e.g. Bell" },
      { key: "model", header: "Model", help: "e.g. B30D" },
      { key: "serial_number", header: "Serial number", help: "Serial or VIN" },
      { key: "year_built", header: "Year built", help: "e.g. 2019" },
      { key: "purchase_date", header: "Purchase date", help: "YYYY-MM-DD" },
      { key: "purchase_price_usd", header: "Purchase price (USD)", help: "Amount in US dollars" },
      { key: "warranty_end", header: "Warranty ends", help: "YYYY-MM-DD" },
      { key: "cost_center_code", header: "Cost centre", help: "Finance cost centre, e.g. PIT" },
      { key: "baseline_fuel_l_per_hour", header: "Fuel L/hr", help: "Expected fuel use in litres per hour" },
      { key: "meter_hours", header: "Meter hours now", help: "Hour meter reading today. Blank = leave as is." },
    ],
  },
  {
    key: "parts", name: "Parts", help: "One row per store item. The part code is the key.",
    columns: [
      { key: "part_code", header: "Part code", required: true, help: "e.g. 210105" },
      { key: "part_name", header: "Description", required: true, help: "e.g. LEVEL INDICATOR" },
      { key: "category", header: "Category", help: `${STOCK_CATEGORIES.map((c) => c.label).join(", ")}. Blank = IronLog decides.`, list: STOCK_CATEGORIES.map((c) => c.label) },
      { key: "unit_cost", header: "Unit cost (USD)", help: "Cost of one unit in US dollars" },
      { key: "min_stock", header: "Minimum stock", help: "Reorder when stock falls below this" },
      { key: "critical", header: "Critical", help: "yes or no (blank = no)", list: YES_NO },
    ],
  },
  {
    key: "stock", name: "Stock on hand", help: "The quantity on the shelf at go-live. IronLog posts the difference from what it has now.",
    columns: [
      { key: "part_code", header: "Part code", required: true, help: "A part on the Parts sheet or already in IronLog" },
      { key: "on_hand", header: "Quantity on hand", required: true, help: "Counted quantity, 0 or more" },
    ],
  },
  {
    key: "suppliers", name: "Suppliers", help: "One row per supplier. The supplier code is the key (use your ERP code, e.g. X3 000028).",
    columns: [
      { key: "supplier_code", header: "Supplier code", required: true, help: "e.g. 000028" },
      { key: "supplier_name", header: "Supplier name", required: true, help: "e.g. BELL EQUIPMENT SALES SA LTD -P" },
      { key: "currency", header: "Currency", help: "USD, ZAR, MZN… (blank = USD)" },
      { key: "lead_time_days", header: "Lead time (days)", help: "Usual days from order to delivery (blank = 7)" },
    ],
  },
  {
    key: "components", name: "Components", help: "Major components fitted to machines. Machine + component + position is the key.",
    columns: [
      { key: "asset_code", header: "Machine code", required: true, help: "A machine on the Machines sheet or already in IronLog" },
      { key: "component", header: "Component", required: true, help: "Component code from the failure code list, e.g. ENG, TRANS, HYD" },
      { key: "position", header: "Position", help: "e.g. Left front (blank if only one)" },
      { key: "serial_number", header: "Serial number", help: "" },
      { key: "make", header: "Make", help: "" },
      { key: "model", header: "Model", help: "" },
      { key: "fitted_date", header: "Fitted date", help: "YYYY-MM-DD" },
      { key: "fitted_hours", header: "Machine hours when fitted", help: "Hour meter when it was fitted" },
      { key: "expected_life_hours", header: "Expected life (hours)", help: "e.g. 12000" },
    ],
  },
];

export function ensureGoLiveTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS go_live_imports (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      file_name TEXT,
      site_code TEXT,
      summary_json TEXT NOT NULL,
      imported_at TEXT NOT NULL DEFAULT (datetime('now')),
      imported_by TEXT
    );
  `);
  ensureAssetDetailColumns(db);
  const cols = new Set(db.prepare(`PRAGMA table_info(assets)`).all().map((c) => c.name));
  if (cols.size) {
    if (!cols.has("cost_center_code")) db.prepare(`ALTER TABLE assets ADD COLUMN cost_center_code TEXT`).run();
    if (!cols.has("baseline_fuel_l_per_hour")) db.prepare(`ALTER TABLE assets ADD COLUMN baseline_fuel_l_per_hour REAL DEFAULT 5.0`).run();
  }
  ensureAssetComponentTables(db);
}

const hasTable = (db, t) => Boolean(db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?`).get(t));
const hasColumn = (db, t, c) => db.prepare(`PRAGMA table_info(${t})`).all().some((r) => r.name === c);

/** An Excel cell's value as text: dates as YYYY-MM-DD, formulas as their result. */
export function cellText(v) {
  if (v == null) return "";
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? "" : v.toISOString().slice(0, 10);
  if (typeof v === "object") {
    if (Array.isArray(v.richText)) return v.richText.map((r) => r.text).join("").trim();
    if ("result" in v) return cellText(v.result);
    if ("text" in v) return cellText(v.text);
    return "";
  }
  return String(v).trim();
}

/** "1 234,5" / "1,234.5" / 1234.5 → 1234.5; "" → null; nonsense → NaN. */
export function readNumber(raw) {
  let s = String(raw ?? "").trim().replace(/[\s$]/g, "");
  if (!s) return null;
  if (s.includes(",") && !s.includes(".")) s = s.replace(",", ".");
  else s = s.replace(/,/g, "");
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}

/** YYYY-MM-DD, or DD/MM/YYYY turned into it; anything else is returned as given (and fails the check). */
function readDate(raw) {
  const s = String(raw ?? "").trim();
  const m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$/);
  return m ? `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}` : s;
}

function readYesNo(raw) {
  const s = String(raw ?? "").trim().toLowerCase();
  if (!s) return { ok: true, value: 0 };
  if (["yes", "y", "true", "1", "sim", "s"].includes(s)) return { ok: true, value: 1 };
  if (["no", "n", "false", "0", "não", "nao"].includes(s)) return { ok: true, value: 0 };
  return { ok: false };
}

/** The workbook's rows per sheet: { sheets: { machines: [{ row, values, has }] … }, missing: [sheet names not found] }. */
export async function readGoLiveWorkbook(ExcelJS, buffer) {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buffer);
  } catch {
    return { ok: false, error: "This is not an Excel workbook (.xlsx). Download the go-live template and fill that in." };
  }
  const sheets = {};
  const unknownColumns = [];
  for (const def of GO_LIVE_SHEETS) {
    const ws = wb.worksheets.find((w) => w.name.trim().toLowerCase() === def.name.toLowerCase());
    sheets[def.key] = [];
    if (!ws) continue;
    const head = ws.getRow(1);
    const colOf = {};
    head.eachCell((cell, col) => {
      const h = cellText(cell.value).replace(/\s*\*$/, "").toLowerCase();
      const c = def.columns.find((x) => x.header.toLowerCase() === h || x.key === h);
      if (c) colOf[c.key] = col;
      else if (h) unknownColumns.push(`${def.name}: "${cellText(cell.value)}"`);
    });
    ws.eachRow({ includeEmpty: false }, (row, n) => {
      if (n === 1) return;
      const values = {};
      for (const c of def.columns) values[c.key] = colOf[c.key] ? cellText(row.getCell(colOf[c.key]).value) : undefined;
      if (Object.values(values).every((v) => v === undefined || v === "")) return;
      sheets[def.key].push({ row: n, values, has: (k) => colOf[k] != null });
    });
  }
  return { ok: true, sheets, unknownColumns };
}

/**
 * Checks the rows against IronLog and works out what would change. Changes nothing.
 * ctx: { site: the upload's site, allowedSites: null = all, today }
 * Returns { ok, errors: [{ sheet, row, column, message }], warnings: [...], summary: { sheet: { new, changed, same } }, actions }.
 */
export function planGoLiveImport(db, parsed, ctx) {
  ensureGoLiveTables(db);
  const errors = [];
  const warnings = [];
  const err = (sheet, row, column, message) => errors.push({ sheet, row, column, message });
  const warn = (sheet, row, column, message) => warnings.push({ sheet, row, column, message });
  const summary = Object.fromEntries(GO_LIVE_SHEETS.map((s) => [s.key, { new: 0, changed: 0, same: 0 }]));
  const actions = { machines: [], parts: [], stock: [], suppliers: [], components: [] };
  const sites = knownSites(db);
  const allowed = ctx.allowedSites;
  const header = (sheet, key) => GO_LIVE_SHEETS.find((s) => s.key === sheet).columns.find((c) => c.key === key).header;
  const sheetName = (key) => GO_LIVE_SHEETS.find((s) => s.key === key).name;
  const same = (a, b) => {
    const na = a == null || a === "" ? null : a;
    const nb = b == null || b === "" ? null : b;
    if (na == null || nb == null) return na === nb;
    return typeof na === "number" || typeof nb === "number" ? Math.abs(Number(na) - Number(nb)) < 1e-9 : String(na) === String(nb);
  };
  const count = (sheet, existing, changes) => {
    if (!existing) summary[sheet].new += 1;
    else if (Object.keys(changes).length) summary[sheet].changed += 1;
    else summary[sheet].same += 1;
  };
  const required = (sheet, r) => {
    let ok = true;
    for (const c of GO_LIVE_SHEETS.find((s) => s.key === sheet).columns) {
      if (c.required && !r.values[c.key]) { err(sheetName(sheet), r.row, c.header, `${c.header} is needed.`); ok = false; }
    }
    return ok;
  };
  const dupes = (sheet, keyOf) => {
    const seen = new Map();
    for (const r of parsed.sheets[sheet] || []) {
      const k = keyOf(r);
      if (!k) continue;
      if (seen.has(k)) err(sheetName(sheet), r.row, "", `Same ${k.split("\u0000")[0]} as row ${seen.get(k)}; keep one row.`);
      else seen.set(k, r.row);
    }
  };

  // Machines.
  const assetCols = new Set(db.prepare(`PRAGMA table_info(assets)`).all().map((c) => c.name));
  const getAsset = db.prepare(`SELECT * FROM assets WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?))`);
  const sheetMachines = new Map();
  dupes("machines", (r) => String(r.values.asset_code || "").toUpperCase());
  for (const r of parsed.sheets.machines || []) {
    const before = errors.length;
    if (!required("machines", r)) continue;
    const S = sheetName("machines");
    const v = r.values;
    const code = v.asset_code;
    if (code.length > 40 || /[\s;,"']/.test(code)) { err(S, r.row, header("machines", "asset_code"), "A machine code has no spaces, commas or quotes (up to 40 characters)."); continue; }
    const existing = getAsset.get(code);
    const site = String(v.site_code || "").trim().toLowerCase() || (existing ? String(existing.site_code || "").toLowerCase() : ctx.site);
    if (!sites.includes(site)) { err(S, r.row, header("machines", "site_code"), `${site} is not a site. Sites: ${sites.join(", ")}.`); continue; }
    if (allowed && !allowed.includes(site)) { err(S, r.row, header("machines", "site_code"), `You may not add machines to site ${site}.`); continue; }
    if (existing && String(existing.site_code || "").toLowerCase() !== site) {
      err(S, r.row, header("machines", "site_code"), `${existing.asset_code} is already on site ${existing.site_code}. Move machines between sites on the Machines screen.`);
      continue;
    }
    if (existing && allowed && !allowed.includes(String(existing.site_code || "").toLowerCase())) { err(S, r.row, header("machines", "asset_code"), `${code} belongs to a site you may not change.`); continue; }
    const values = { asset_name: v.asset_name.slice(0, 120), category: v.category.slice(0, 60), site_code: site };
    if (r.has("is_standby")) {
      const yn = readYesNo(v.is_standby);
      if (!yn.ok) err(S, r.row, header("machines", "is_standby"), "Write yes or no.");
      else values.is_standby = yn.value;
    }
    const detailBody = Object.fromEntries(Object.keys(ASSET_DETAIL_FIELDS).filter((k) => r.has(k)).map((k) => [k, ASSET_DETAIL_FIELDS[k].type === "date" ? readDate(v[k]) : v[k]]));
    const details = readAssetDetails(detailBody, { thisYear: Number(ctx.today.slice(0, 4)) });
    if (!details.ok) details.errors.forEach((m) => err(S, r.row, "", m));
    else Object.assign(values, details.values);
    if (r.has("cost_center_code")) {
      const cc = v.cost_center_code;
      if (cc.length > 40) err(S, r.row, header("machines", "cost_center_code"), "A cost centre is at most 40 characters.");
      else values.cost_center_code = cc || null;
    }
    if (r.has("baseline_fuel_l_per_hour") && v.baseline_fuel_l_per_hour !== "") {
      const n = readNumber(v.baseline_fuel_l_per_hour);
      if (!(n >= 0 && n <= 500)) err(S, r.row, header("machines", "baseline_fuel_l_per_hour"), "Litres per hour must be a number from 0 to 500.");
      else values.baseline_fuel_l_per_hour = n;
    }
    let meter = null;
    if (r.has("meter_hours") && v.meter_hours !== "") {
      const n = readNumber(v.meter_hours);
      if (!(n >= 0 && n <= 1e6)) err(S, r.row, header("machines", "meter_hours"), "Meter hours must be a number of hours (0 or more).");
      else {
        meter = Math.round(n * 10) / 10;
        if (existing) {
          const last = db.prepare(`SELECT closing_hours FROM daily_hours WHERE asset_id = ? AND closing_hours IS NOT NULL ORDER BY work_date DESC, id DESC LIMIT 1`).get(existing.id);
          if (last && meter < Number(last.closing_hours)) {
            err(S, r.row, header("machines", "meter_hours"), `Lower than the last daily reading (${Number(last.closing_hours)} h). Correct the daily input instead.`);
            meter = null;
          }
        }
      }
    }
    const changes = {};
    for (const [k, val] of Object.entries(values)) {
      if (!assetCols.has(k)) continue;
      if (!existing || !same(existing[k], val)) changes[k] = val;
    }
    if (existing && meter != null) {
      const cur = db.prepare(`SELECT total_hours FROM asset_hours WHERE asset_id = ?`).get(existing.id);
      if (!cur || !same(Number(cur.total_hours), meter)) changes.meter_hours = meter;
    }
    sheetMachines.set(code.toUpperCase(), site);
    if (errors.length > before) continue;
    count("machines", existing, changes);
    actions.machines.push({ code, id: existing?.id || null, values: Object.fromEntries(Object.entries(values).filter(([k]) => assetCols.has(k))), meter, changed: !existing || Object.keys(changes).length > 0 });
  }

  // Parts.
  ensureStockCategorySchema(db);
  const partCols = new Set(db.prepare(`PRAGMA table_info(parts)`).all().map((c) => c.name));
  const getPart = db.prepare(`SELECT * FROM parts WHERE part_code = ?`);
  const sheetParts = new Map();
  dupes("parts", (r) => r.values.part_code);
  for (const r of parsed.sheets.parts || []) {
    const before = errors.length;
    if (!required("parts", r)) continue;
    const S = sheetName("parts");
    const v = r.values;
    if (v.part_code.length > 60) { err(S, r.row, header("parts", "part_code"), "A part code is at most 60 characters."); continue; }
    const existing = getPart.get(v.part_code);
    const values = { part_name: v.part_name.slice(0, 200) };
    if (r.has("category") && v.category) {
      const label = STOCK_CATEGORIES.find((c) => c.label.toLowerCase() === v.category.toLowerCase());
      const cat = label ? label.key : normalizeStockCategory(v.category);
      if (!cat) err(S, r.row, header("parts", "category"), `Choose one of: ${STOCK_CATEGORIES.map((c) => c.label).join(", ")}.`);
      else values.stock_category = cat;
    }
    for (const [k, label] of [["unit_cost", "Unit cost"], ["min_stock", "Minimum stock"]]) {
      if (!r.has(k) || v[k] === "") continue;
      const n = readNumber(v[k]);
      if (!(n >= 0)) err(S, r.row, header("parts", k), `${label} must be a number, 0 or more.`);
      else values[k] = k === "unit_cost" ? Math.round(n * 10000) / 10000 : n;
    }
    if (r.has("critical")) {
      const yn = readYesNo(v.critical);
      if (!yn.ok) err(S, r.row, header("parts", "critical"), "Write yes or no.");
      else values.critical = yn.value;
    }
    const changes = {};
    for (const [k, val] of Object.entries(values)) if (partCols.has(k) && (!existing || !same(existing[k], val))) changes[k] = val;
    sheetParts.set(v.part_code, values.unit_cost ?? existing?.unit_cost ?? 0);
    if (errors.length > before) continue;
    count("parts", existing, changes);
    actions.parts.push({ code: v.part_code, id: existing?.id || null, values, changed: !existing || Object.keys(changes).length > 0 });
  }

  // Stock on hand.
  const onHandOf = db.prepare(`SELECT COALESCE(SUM(quantity), 0) AS q, COUNT(*) AS moves FROM stock_movements WHERE part_id = ?`);
  dupes("stock", (r) => r.values.part_code);
  for (const r of parsed.sheets.stock || []) {
    if (!required("stock", r)) continue;
    const S = sheetName("stock");
    const part = getPart.get(r.values.part_code);
    if (!part && !sheetParts.has(r.values.part_code)) { err(S, r.row, header("stock", "part_code"), `${r.values.part_code} is not on the Parts sheet or in IronLog.`); continue; }
    const target = readNumber(r.values.on_hand);
    if (!(target >= 0)) { err(S, r.row, header("stock", "on_hand"), "The quantity must be a number, 0 or more."); continue; }
    const cur = part ? onHandOf.get(part.id) : { q: 0, moves: 0 };
    const diff = Math.round((target - Number(cur.q)) * 10000) / 10000;
    if (!part) summary.stock.new += 1;
    else if (diff) summary.stock.changed += 1;
    else summary.stock.same += 1;
    if (diff && part && Number(cur.moves) > 0) warn(S, r.row, header("stock", "on_hand"), `${part.part_code}: IronLog has ${Number(cur.q)} on hand; this posts ${diff > 0 ? "+" : ""}${diff} to make it ${target}.`);
    actions.stock.push({ code: r.values.part_code, diff, target });
  }

  // Suppliers.
  const supplierTable = hasTable(db, "suppliers");
  if ((parsed.sheets.suppliers || []).length && !supplierTable) err(sheetName("suppliers"), 1, "", "The supplier list is not set up on this server yet. Open Procurement once, then try again.");
  const getSupplier = supplierTable ? db.prepare(`SELECT * FROM suppliers WHERE supplier_code = ?`) : null;
  dupes("suppliers", (r) => r.values.supplier_code);
  for (const r of supplierTable ? parsed.sheets.suppliers || [] : []) {
    const before = errors.length;
    if (!required("suppliers", r)) continue;
    const S = sheetName("suppliers");
    const v = r.values;
    const existing = getSupplier.get(v.supplier_code);
    const values = { supplier_name: v.supplier_name.slice(0, 160) };
    if (r.has("currency")) {
      const c = (v.currency || "USD").toUpperCase();
      if (!/^[A-Z]{3}$/.test(c)) err(S, r.row, header("suppliers", "currency"), "Currency is a 3-letter code like USD, ZAR or MZN.");
      else values.currency = c;
    }
    if (r.has("lead_time_days")) {
      const n = v.lead_time_days === "" ? 7 : readNumber(v.lead_time_days);
      if (!(Number.isInteger(n) && n >= 0 && n <= 365)) err(S, r.row, header("suppliers", "lead_time_days"), "Lead time is a whole number of days (0 to 365).");
      else values.lead_time_days = n;
    }
    const changes = {};
    for (const [k, val] of Object.entries(values)) if (!existing || !same(existing[k], val)) changes[k] = val;
    if (errors.length > before) continue;
    count("suppliers", existing, changes);
    actions.suppliers.push({ code: v.supplier_code, id: existing?.id || null, values, changed: !existing || Object.keys(changes).length > 0 });
  }

  // Components.
  const componentCodes = listFailureCodes(db, { all: true }).component.map((c) => c.code);
  dupes("components", (r) => `${String(r.values.asset_code || "").toUpperCase()}/${String(r.values.component || "").toUpperCase()}\u0000${String(r.values.position || "").toLowerCase()}`);
  for (const r of parsed.sheets.components || []) {
    if (!required("components", r)) continue;
    const S = sheetName("components");
    const v = r.values;
    const machine = getAsset.get(v.asset_code);
    const machineSite = machine ? String(machine.site_code || "").toLowerCase() : sheetMachines.get(v.asset_code.toUpperCase());
    if (!machineSite) { err(S, r.row, header("components", "asset_code"), `${v.asset_code} is not on the Machines sheet or in IronLog.`); continue; }
    if (allowed && !allowed.includes(machineSite)) { err(S, r.row, header("components", "asset_code"), `${v.asset_code} belongs to a site you may not change.`); continue; }
    const body = Object.fromEntries(Object.keys(v).filter((k) => k !== "asset_code" && r.has(k)).map((k) => [k, k === "fitted_date" ? readDate(v[k]) : v[k]]));
    body.component = v.component;
    const read = readComponentFields(body, { componentCodes, requireComponent: true });
    if (!read.ok) { read.errors.forEach((m) => err(S, r.row, "", m)); continue; }
    const position = read.values.position || null;
    const existing = machine
      ? db.prepare(`SELECT * FROM asset_components WHERE asset_id = ? AND component = ? AND COALESCE(LOWER(position), '') = ? AND status != 'removed' ORDER BY id DESC LIMIT 1`)
        .get(machine.id, read.values.component, String(position || "").toLowerCase())
      : null;
    const changes = {};
    for (const [k, val] of Object.entries(read.values)) if (!existing || !same(existing[k], val)) changes[k] = val;
    count("components", existing, changes);
    actions.components.push({ asset_code: v.asset_code, id: existing?.id || null, values: read.values, changed: !existing || Object.keys(changes).length > 0 });
  }

  for (const u of parsed.unknownColumns || []) warn("", 1, "", `Column not used: ${u}`);
  const rows = GO_LIVE_SHEETS.reduce((t, s) => t + (parsed.sheets[s.key] || []).length, 0);
  if (!rows) err("", 0, "", "The workbook has no rows to import. Fill in the go-live template and upload it again.");
  return { ok: errors.length === 0, errors, warnings, summary, actions, rows };
}

/** Applies a clean plan in one transaction. Returns the summary. */
export function applyGoLiveImport(db, plan, { by = "" } = {}) {
  if (!plan.ok) throw new Error("The import has problems; check the file first.");
  const assetCols = new Set(db.prepare(`PRAGMA table_info(assets)`).all().map((c) => c.name));
  const movementCost = hasColumn(db, "stock_movements", "unit_cost_usd");
  db.transaction(() => {
    for (const m of plan.actions.machines) {
      let id = m.id;
      const cols = Object.keys(m.values).filter((c) => assetCols.has(c));
      if (!id) {
        id = Number(db.prepare(`INSERT INTO assets (asset_code, ${cols.join(", ")}) VALUES (?, ${cols.map(() => "?").join(", ")})`).run(m.code, ...cols.map((c) => m.values[c])).lastInsertRowid);
      } else if (m.changed) {
        db.prepare(`UPDATE assets SET ${cols.map((c) => `${c} = ?`).join(", ")} WHERE id = ?`).run(...cols.map((c) => m.values[c]), id);
      }
      if (m.meter != null) {
        db.prepare(`INSERT INTO asset_hours (asset_id, total_hours, last_updated) VALUES (?, ?, datetime('now'))
          ON CONFLICT(asset_id) DO UPDATE SET total_hours = excluded.total_hours, last_updated = excluded.last_updated`).run(id, m.meter);
      }
    }
    const setCategory = db.prepare(`UPDATE parts SET stock_category = ?, stock_category_source = 'manual' WHERE part_code = ?`);
    const partCols = new Set(db.prepare(`PRAGMA table_info(parts)`).all().map((c) => c.name));
    for (const p of plan.actions.parts) {
      const { stock_category: cat, ...vals } = p.values;
      const cols = Object.keys(vals).filter((c) => partCols.has(c));
      if (!p.id) db.prepare(`INSERT INTO parts (part_code, ${cols.join(", ")}) VALUES (?, ${cols.map(() => "?").join(", ")})`).run(p.code, ...cols.map((c) => vals[c]));
      else if (p.changed) db.prepare(`UPDATE parts SET ${cols.map((c) => `${c} = ?`).join(", ")} WHERE id = ?`).run(...cols.map((c) => vals[c]), p.id);
      if (cat) setCategory.run(cat, p.code);
    }
    for (const s of plan.actions.stock) {
      if (!s.diff) continue;
      const part = db.prepare(`SELECT id, unit_cost FROM parts WHERE part_code = ?`).get(s.code);
      if (movementCost) {
        db.prepare(`INSERT INTO stock_movements (part_id, quantity, movement_type, reference, unit_cost_usd) VALUES (?, ?, ?, 'go_live_opening', ?)`)
          .run(part.id, s.diff, s.diff > 0 ? "in" : "adjust", Number(part.unit_cost || 0));
      } else {
        db.prepare(`INSERT INTO stock_movements (part_id, quantity, movement_type, reference) VALUES (?, ?, ?, 'go_live_opening')`).run(part.id, s.diff, s.diff > 0 ? "in" : "adjust");
      }
    }
    for (const s of plan.actions.suppliers) {
      const cols = Object.keys(s.values);
      if (!s.id) db.prepare(`INSERT INTO suppliers (supplier_code, ${cols.join(", ")}) VALUES (?, ${cols.map(() => "?").join(", ")})`).run(s.code, ...cols.map((c) => s.values[c]));
      else if (s.changed) db.prepare(`UPDATE suppliers SET ${cols.map((c) => `${c} = ?`).join(", ")}, updated_at = datetime('now') WHERE id = ?`).run(...cols.map((c) => s.values[c]), s.id);
    }
    for (const c of plan.actions.components) {
      if (c.id) { if (c.changed) updateAssetComponent(db, c.id, c.values, by); continue; }
      const machine = db.prepare(`SELECT id FROM assets WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?))`).get(c.asset_code);
      const meter = db.prepare(`SELECT total_hours FROM asset_hours WHERE asset_id = ?`).get(machine.id)?.total_hours;
      addAssetComponent(db, machine.id, c.values, meter, by);
    }
  })();
  if (plan.actions.parts.length) ensureStockCategorySchema(db, { refresh: true });
  return plan.summary;
}

/** The go-live workbook: a "Read me" sheet and one sheet per kind of record, optionally filled with current data. */
export function goLiveTemplate(ExcelJS, { rows = {}, site = "main", sites = [], componentCodes = [] } = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "IRONLOG";
  const readme = wb.addWorksheet("Read me");
  readme.columns = [{ width: 24 }, { width: 34 }, { width: 70 }];
  const lines = [
    ["IronLog go-live workbook"],
    [],
    ["How it works", "", "Fill in the sheets you need (you can leave a sheet empty). Upload the file under Uploads → Go-live import, press Check, fix anything listed, then Import."],
    ["", "", "Codes are the key: a code IronLog already has is updated, so you can upload the same file again after fixing it."],
    ["", "", "A blank cell clears that detail, except meter hours and stock on hand (blank = leave as is). Columns marked * must be filled in."],
    ["", "", "Dates: YYYY-MM-DD (DD/MM/YYYY also works). Money: US dollars. Yes/no columns: yes or no."],
    ["Sites", "", `Machines go to the site in the Site column, or to the site chosen when uploading (${site}). Sites: ${sites.join(", ") || site}.`],
    ["Components", "", `Component codes: ${componentCodes.join(", ")}.`],
    [],
    ["Sheet", "Column", "What to put in it"],
  ];
  for (const def of GO_LIVE_SHEETS) {
    lines.push([def.name, "", def.help]);
    for (const c of def.columns) lines.push(["", `${c.header}${c.required ? " *" : ""}`, c.help]);
  }
  lines.forEach((l) => readme.addRow(l));
  readme.getRow(1).font = { bold: true, size: 14 };
  readme.getRow(10).font = { bold: true };
  readme.getColumn(3).alignment = { wrapText: true, vertical: "top" };
  for (const def of GO_LIVE_SHEETS) {
    const ws = wb.addWorksheet(def.name, { views: [{ state: "frozen", ySplit: 1 }] });
    ws.columns = def.columns.map((c) => ({ header: `${c.header}${c.required ? " *" : ""}`, key: c.key, width: Math.max(14, c.header.length + 4) }));
    ws.getRow(1).font = { bold: true };
    def.columns.forEach((c, i) => { ws.getRow(1).getCell(i + 1).note = c.help; });
    for (const r of rows[def.key] || []) ws.addRow(r);
    def.columns.forEach((c, i) => {
      if (!c.list) return;
      // One range rule, so the sheet has no empty rows and new rows start right under the heading.
      const letter = ws.getColumn(i + 1).letter;
      const last = Math.max(1000, (rows[def.key] || []).length + 500);
      ws.dataValidations.add(`${letter}2:${letter}${last}`, { type: "list", allowBlank: true, formulae: [`"${c.list.join(",")}"`] });
    });
  }
  return wb;
}

/** Current data for a filled-in template (machines of the given sites only). */
export function goLiveCurrentRows(db, sitesShown) {
  ensureGoLiveTables(db);
  const placeholders = sitesShown.map(() => "?").join(", ");
  const machines = sitesShown.length ? db.prepare(`
    SELECT a.*, h.total_hours AS meter_hours FROM assets a LEFT JOIN asset_hours h ON h.asset_id = a.id
    WHERE COALESCE(a.archived, 0) = 0 AND LOWER(TRIM(COALESCE(a.site_code, 'main'))) IN (${placeholders}) ORDER BY a.asset_code
  `).all(...sitesShown) : [];
  const yn = (v) => (Number(v) ? "yes" : "no");
  const catLabel = Object.fromEntries(STOCK_CATEGORIES.map((c) => [c.key, c.label]));
  const parts = db.prepare(`SELECT p.*, (SELECT COALESCE(SUM(quantity), 0) FROM stock_movements m WHERE m.part_id = p.id) AS on_hand FROM parts p ORDER BY part_code`).all();
  const ids = new Set(machines.map((m) => m.id));
  const comps = db.prepare(`SELECT c.*, a.asset_code FROM asset_components c JOIN assets a ON a.id = c.asset_id WHERE c.status != 'removed' ORDER BY a.asset_code, c.component`).all().filter((c) => ids.has(c.asset_id));
  return {
    machines: machines.map((m) => ({ ...m, is_standby: yn(m.is_standby) })),
    parts: parts.map((p) => ({ ...p, category: p.stock_category_source === "manual" ? catLabel[p.stock_category] || "" : "", critical: yn(p.critical) })),
    stock: parts.filter((p) => Number(p.on_hand)).map((p) => ({ part_code: p.part_code, on_hand: Number(p.on_hand) })),
    suppliers: hasTable(db, "suppliers") ? db.prepare(`SELECT * FROM suppliers ORDER BY supplier_code`).all() : [],
    components: comps,
  };
}

/** Go-live checklist for a site: what is in place and what still needs doing. */
export function goLiveStatus(db, site) {
  ensureGoLiveTables(db);
  const n = (sql, ...a) => { try { return Number(db.prepare(sql).get(...a)?.c || 0); } catch { return 0; } };
  const s = String(site || "main").toLowerCase();
  const machines = n(`SELECT COUNT(*) AS c FROM assets WHERE COALESCE(archived, 0) = 0 AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?`, s);
  const withMeter = n(`SELECT COUNT(*) AS c FROM assets a WHERE COALESCE(a.archived, 0) = 0 AND LOWER(TRIM(COALESCE(a.site_code, 'main'))) = ?
    AND (EXISTS (SELECT 1 FROM asset_hours h WHERE h.asset_id = a.id AND h.total_hours > 0) OR EXISTS (SELECT 1 FROM daily_hours d WHERE d.asset_id = a.id AND d.closing_hours IS NOT NULL))`, s);
  const withDetails = n(`SELECT COUNT(*) AS c FROM assets WHERE COALESCE(archived, 0) = 0 AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ? AND COALESCE(make, '') != '' AND COALESCE(model, '') != ''`, s);
  const parts = n(`SELECT COUNT(*) AS c FROM parts`);
  const partsCosted = n(`SELECT COUNT(*) AS c FROM parts WHERE COALESCE(unit_cost, 0) > 0`);
  const stocked = n(`SELECT COUNT(*) AS c FROM (SELECT part_id FROM stock_movements GROUP BY part_id HAVING SUM(quantity) > 0)`);
  const suppliers = n(`SELECT COUNT(*) AS c FROM suppliers WHERE COALESCE(active, 1) = 1`);
  const components = n(`SELECT COUNT(*) AS c FROM asset_components c JOIN assets a ON a.id = c.asset_id WHERE c.status != 'removed' AND LOWER(TRIM(COALESCE(a.site_code, 'main'))) = ?`, s);
  const people = n(`SELECT COUNT(*) AS c FROM users WHERE COALESCE(active, 1) = 1`);
  const hoursWeek = n(`SELECT COUNT(DISTINCT work_date) AS c FROM daily_hours d JOIN assets a ON a.id = d.asset_id WHERE LOWER(TRIM(COALESCE(a.site_code, 'main'))) = ? AND d.work_date >= date('now', '-7 days')`, s);
  const step = (key, label, done, detail) => ({ key, label, done, detail });
  const many = (k, one, more = `${one}s`) => `${k} ${k === 1 ? one : more}`;
  return {
    site: s,
    steps: [
      step("machines", "Machines loaded", machines > 0, many(machines, "machine")),
      step("details", "Make and model filled in", machines > 0 && withDetails === machines, `${withDetails} of ${machines}`),
      step("meters", "Meter hours set", machines > 0 && withMeter === machines, `${withMeter} of ${machines}`),
      step("parts", "Store items loaded", parts > 0, `${many(parts, "item")}, ${partsCosted} with a unit cost`),
      step("stock", "Opening stock counted", stocked > 0, `${many(stocked, "item")} in stock`),
      step("suppliers", "Suppliers loaded", suppliers > 0, many(suppliers, "supplier")),
      step("components", "Major components listed", components > 0, many(components, "component")),
      step("people", "People added (Admin → Users)", people > 1, many(people, "person", "people")),
      step("hours", "Daily hours being captured", hoursWeek >= 3, `${hoursWeek} of the last 7 days`),
    ],
  };
}
