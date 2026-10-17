// IRONLOG/api/utils/x3SiteRequest.js — IronLog requisitions as Sage X3 Site Requests (import file).
//
// AML's Sage X3 has a custom Site Request screen (window OZS1, blocks ZSREQ0/ZSREQ2/ZSREQ1).
// The field codes below come from an X3 screen export of Site Request VS120-00353.
// IronLog writes one "E" line per requisition (header) followed by one "L" line per part,
// fields in the order of X3_HEADER_FIELDS / X3_LINE_FIELDS, every value in double quotes
// and separated by ";" — the same style as X3's own export. The X3 administrator creates
// an import template with these fields in this order (see docs/x3-site-request-import.md);
// X3 gives each request its Site Request number, which is then typed back into IronLog.
//
// X3 codes IronLog does not hold (sites, request type, signature channel, buyer, customer,
// project, the generic product code) come from settings. Suppliers use the IronLog supplier
// code when the supplier is in the supplier list, otherwise the name → code list in settings.

export const X3_HEADER_FIELDS = [
  { code: "ZSOURCESA", label: "Source from TSA?", from: "Setting (X3 menu number, 2 = Yes)" },
  { code: "ZDELFCY", label: "Delivery Site", from: "Setting", required: true },
  { code: "ZFCY", label: "Purchase Site", from: "Setting", required: true },
  { code: "ZREF", label: "Manual Req no.", from: "IronLog requisition number (prefix + number)" },
  { code: "ZBILLCUR", label: "Bill Currency", from: "Setting" },
  { code: "ZCONTRACT", label: "Contract No", from: "Empty" },
  { code: "ZDLVPIO", label: "Delivery priority", from: "Setting" },
  { code: "ZREQTYP", label: "Request type", from: "Setting", required: true },
  { code: "ZSIGNCHNL", label: "Signature Channel", from: "Setting", required: true },
  { code: "ZREQUESTBY", label: "Requested By", from: "Requester's X3 code (user list), else setting", required: true },
  { code: "ZBUYER", label: "Buyer", from: "Setting", required: true },
  { code: "ZDATEREQ", label: "Date Required", from: "Needed-by date, else today" },
  { code: "ZHBPSNUM", label: "Supplier", from: "Supplier code" },
  { code: "ZSUPPQUOTE", label: "Supplier Quote No.", from: "Requisition PO / quote number" },
  { code: "ZPREFSUPP", label: "Preferred Supplier", from: "Empty" },
  { code: "ZBPCORD", label: "Customer", from: "Setting" },
  { code: "ZBPCINV", label: "Bill-to customer", from: "Setting" },
  { code: "ZBPAADD", label: "Delivery address", from: "Setting" },
  { code: "ZCUR", label: "Currency", from: "Setting" },
  { code: "ZAX1", label: "Equipment", from: "Machine on the first part line" },
  { code: "ZVIN", label: "VIN No", from: "Machine serial number (machine details)" },
  { code: "ZODOUNIT", label: "ODO Unit of Measure", from: "Setting (empty = no reading)" },
  { code: "ZODOREADING", label: "ODO Reading", from: "Machine's last meter reading when an ODO unit is set" },
  { code: "ZAX2", label: "Job Card", from: "Empty (X3 dimension)" },
  { code: "ZAX3", label: "Project", from: "Setting" },
  { code: "ZSITEJOB", label: "Site Job Card", from: "Job card on the first part line" },
];

export const X3_LINE_FIELDS = [
  { code: "ZITMREF", label: "Product", from: "Setting (X3 product for parts)", required: true },
  { code: "ZOEM", label: "Part No", from: "IronLog part code" },
  { code: "ZPRINTDESC", label: "Description", from: "Line description" },
  { code: "ZUOM", label: "Unit", from: "Line unit, else setting" },
  { code: "ZQTY", label: "Quantity", from: "Line quantity" },
  { code: "ZEQUIPMENT", label: "Equipment No", from: "Line machine" },
  { code: "ZODOUNIT", label: "ODO Unit of Measure", from: "As header" },
  { code: "ZODOREADING", label: "ODO Reading", from: "As header" },
  { code: "ZJOBCARD", label: "Job Card", from: "Empty (X3 dimension)" },
  { code: "ZPROJECT", label: "Project", from: "Setting" },
  { code: "ZBPSNUM", label: "Recommended supplier", from: "Supplier code" },
  { code: "ZGROPRI", label: "Gross price", from: "Line gross price" },
  { code: "ZLCUR", label: "Currency", from: "Line currency, else header currency" },
  { code: "ZDISTYP", label: "Discount type", from: "Line discount type" },
  { code: "ZDISCRGVAL", label: "Discount", from: "Line discount" },
  { code: "ZNETPRI", label: "Net price", from: "Line net price, else gross price" },
  { code: "ZDATEREQ", label: "Date Required", from: "Line needed-by date, else header date" },
];

export const X3_DEFAULTS = {
  separator: ";",
  date_format: "DD/MM/YYYY",
  header_marker: "E",
  line_marker: "L",
  reference_prefix: "IL",
  source_tsa: "2",
  delivery_site: "",
  purchase_site: "",
  bill_currency: "USD",
  delivery_priority: "1",
  request_type: "",
  signature_channel: "",
  requested_by: "",
  buyer: "",
  customer: "",
  bill_to_customer: "",
  delivery_address: "",
  currency: "USD",
  project: "",
  item_code: "",
  unit: "UN",
  odo_unit: "",
  user_codes: "",
  supplier_codes: "",
};

const TEXT_KEYS = ["header_marker", "line_marker", "reference_prefix", "source_tsa", "delivery_site", "purchase_site", "bill_currency", "delivery_priority",
  "request_type", "signature_channel", "requested_by", "buyer", "customer", "bill_to_customer", "delivery_address", "currency", "project", "item_code", "unit", "odo_unit"];
const CODE_RE = /^[A-Za-z0-9][A-Za-z0-9 ._/*&()-]{0,39}$/;

export function ensureX3Tables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS x3_site_request_settings (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      settings_json TEXT NOT NULL,
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_by TEXT
    );
  `);
  const cols = db.prepare(`PRAGMA table_info(procurement_requisitions)`).all().map((c) => c.name);
  if (!cols.length) return;
  for (const [c, type] of [["x3_site_request_no", "TEXT"], ["x3_exported_at", "TEXT"], ["x3_exported_by", "TEXT"]]) {
    if (!cols.includes(c)) db.prepare(`ALTER TABLE procurement_requisitions ADD COLUMN ${c} ${type}`).run();
  }
}

export function readX3Settings(db) {
  ensureX3Tables(db);
  const row = db.prepare(`SELECT settings_json, updated_at, updated_by FROM x3_site_request_settings WHERE id = 1`).get();
  let saved = {};
  try { saved = JSON.parse(row?.settings_json || "{}") || {}; } catch { saved = {}; }
  return { ...X3_DEFAULTS, ...saved, updated_at: row?.updated_at || null, updated_by: row?.updated_by || null };
}

/** "name = code" lines → Map of lower-case name → code. Bad lines are reported. */
export function parseCodeList(text) {
  const map = new Map();
  const bad = [];
  String(text || "").split(/\r?\n/).forEach((raw, i) => {
    const line = raw.trim();
    if (!line) return;
    const at = line.lastIndexOf("=");
    const name = at > 0 ? line.slice(0, at).trim() : "";
    const code = at > 0 ? line.slice(at + 1).trim() : "";
    if (!name || !CODE_RE.test(code)) bad.push(i + 1);
    else map.set(name.toLowerCase(), code);
  });
  return { map, bad };
}

/** Checks and saves the settings: { ok, settings } or { ok:false, errors }. */
export function saveX3Settings(db, body = {}, by = "") {
  const cur = readX3Settings(db);
  const next = Object.fromEntries(Object.keys(X3_DEFAULTS).map((k) => [k, cur[k]]));
  const errors = [];
  if (body.separator !== undefined) {
    if (![";", ",", "|", "\t"].includes(body.separator)) errors.push("The separator must be ; , | or a tab.");
    else next.separator = body.separator;
  }
  if (body.date_format !== undefined) {
    if (!["DD/MM/YYYY", "DDMMYY", "YYYYMMDD", "YYYY-MM-DD"].includes(body.date_format)) errors.push("Choose a date format from the list.");
    else next.date_format = body.date_format;
  }
  for (const k of TEXT_KEYS) {
    if (body[k] === undefined) continue;
    const v = String(body[k] ?? "").trim();
    if (v && !CODE_RE.test(v)) errors.push(`"${v}" is not a valid X3 code (letters, numbers and - . / _ * up to 40).`);
    else next[k] = v;
  }
  for (const k of ["user_codes", "supplier_codes"]) {
    if (body[k] === undefined) continue;
    const text = String(body[k] ?? "");
    const { bad } = parseCodeList(text);
    if (bad.length) errors.push(`${k === "user_codes" ? "People" : "Suppliers"} list: write each line as name = code (check line ${bad.join(", ")}).`);
    else next[k] = text.trim();
  }
  if (!next.header_marker || !next.line_marker || next.header_marker === next.line_marker) errors.push("The header and line markers must be filled in and different.");
  if (errors.length) return { ok: false, errors };
  ensureX3Tables(db);
  db.prepare(`
    INSERT INTO x3_site_request_settings (id, settings_json, updated_at, updated_by) VALUES (1, ?, datetime('now'), ?)
    ON CONFLICT(id) DO UPDATE SET settings_json = excluded.settings_json, updated_at = excluded.updated_at, updated_by = excluded.updated_by
  `).run(JSON.stringify(next), by || null);
  return { ok: true, settings: readX3Settings(db) };
}

export function formatX3Date(ymd, fmt) {
  const [y, m, d] = String(ymd || "").slice(0, 10).split("-");
  if (!y || !m || !d) return "";
  if (fmt === "DDMMYY") return `${d}${m}${y.slice(2)}`;
  if (fmt === "YYYYMMDD") return `${y}${m}${d}`;
  if (fmt === "YYYY-MM-DD") return `${y}-${m}-${d}`;
  return `${d}/${m}/${y}`;
}

const num = (v) => {
  if (v == null || v === "") return "";
  const n = Number(v);
  return Number.isFinite(n) ? String(Math.round(n * 10000) / 10000) : "";
};
const validDate = (v) => (/^\d{4}-\d{2}-\d{2}/.test(String(v || "")) ? String(v).slice(0, 10) : "");

/**
 * Builds the X3 records for requisitions.
 * reqs: [{ requisition, lines, asset: { serial_number, meter } | null, supplier_code }]
 * Returns { ok, errors, warnings, records: [{ id, ref, header: {code: value}, lines: [{code: value}] }] }.
 */
export function buildX3SiteRequests(reqs, settings, { today }) {
  const s = { ...X3_DEFAULTS, ...settings };
  const errors = [];
  const warnings = [];
  const missing = [...X3_HEADER_FIELDS, ...X3_LINE_FIELDS].filter((f) => f.required && f.from.startsWith("Setting"));
  const settingFor = { ZDELFCY: "delivery_site", ZFCY: "purchase_site", ZREQTYP: "request_type", ZSIGNCHNL: "signature_channel", ZBUYER: "buyer", ZITMREF: "item_code" };
  const unset = missing.filter((f) => !s[settingFor[f.code]]).map((f) => f.label);
  if (unset.length) errors.push(`Fill in the X3 settings first: ${unset.join(", ")}.`);
  const users = parseCodeList(s.user_codes).map;
  const suppliers = parseCodeList(s.supplier_codes).map;

  const records = (reqs || []).map(({ requisition: r, lines, asset, supplier_code }) => {
    const ref = `${s.reference_prefix || ""}${String(r.id).padStart(6, "0")}`;
    const where = `Requisition ${r.id}`;
    if (String(r.status || "").toLowerCase() === "draft") errors.push(`${where} is still a draft. Finalize it first.`);
    if (!lines.length) errors.push(`${where} has no part lines.`);
    const requestedBy = users.get(String(r.requester || "").trim().toLowerCase()) || s.requested_by;
    if (!requestedBy) errors.push(`${where}: no X3 code for ${r.requester || "the requester"}. Add them to the people list or set "Requested By".`);
    const supplier = supplier_code || suppliers.get(String(r.supplier_name || "").trim().toLowerCase()) || "";
    if (r.supplier_name && !supplier) warnings.push(`${where}: no X3 code for supplier "${r.supplier_name}", so the supplier is left empty. Add it to the suppliers list.`);
    const first = lines[0] || {};
    const equipment = String(first.equipment_no || "").trim();
    if (!equipment) warnings.push(`${where}: no machine on the first part line, so Equipment is left empty.`);
    const reading = s.odo_unit && asset?.meter != null ? num(asset.meter) : "0";
    const headerDate = validDate(r.needed_by_date) || lines.map((l) => validDate(l.needed_by_date)).filter(Boolean).sort()[0] || today;
    const header = {
      ZSOURCESA: s.source_tsa, ZDELFCY: s.delivery_site, ZFCY: s.purchase_site, ZREF: ref, ZBILLCUR: s.bill_currency, ZCONTRACT: "",
      ZDLVPIO: s.delivery_priority, ZREQTYP: s.request_type, ZSIGNCHNL: s.signature_channel, ZREQUESTBY: requestedBy, ZBUYER: s.buyer,
      ZDATEREQ: formatX3Date(headerDate, s.date_format), ZHBPSNUM: supplier, ZSUPPQUOTE: String(r.po_number || "").trim(), ZPREFSUPP: "",
      ZBPCORD: s.customer, ZBPCINV: s.bill_to_customer, ZBPAADD: s.delivery_address, ZCUR: s.currency, ZAX1: equipment,
      ZVIN: equipment ? String(asset?.serial_number || "").trim() : "", ZODOUNIT: s.odo_unit, ZODOREADING: reading, ZAX2: "", ZAX3: s.project,
      ZSITEJOB: String(first.job_card || "").trim(),
    };
    const out = lines.map((l) => {
      if (!(Number(l.quantity) > 0)) errors.push(`${where} line ${l.line_no}: the quantity must be more than 0.`);
      return {
        ZITMREF: s.item_code, ZOEM: String(l.product_code || l.part_code || "").trim(), ZPRINTDESC: String(l.description || l.part_name || "").trim().slice(0, 60),
        ZUOM: String(l.unit || "").trim() || s.unit, ZQTY: num(l.quantity), ZEQUIPMENT: String(l.equipment_no || equipment).trim(),
        ZODOUNIT: s.odo_unit, ZODOREADING: reading, ZJOBCARD: "", ZPROJECT: s.project, ZBPSNUM: supplier,
        ZGROPRI: num(l.gross_price), ZLCUR: String(l.currency || "").trim() || s.currency, ZDISTYP: String(l.discount_type || "").trim(),
        ZDISCRGVAL: num(l.discount), ZNETPRI: num(l.net_price ?? l.gross_price), ZDATEREQ: formatX3Date(validDate(l.needed_by_date) || headerDate, s.date_format),
      };
    });
    if (r.x3_site_request_no) warnings.push(`${where} is already in X3 as ${r.x3_site_request_no}. Importing it again makes a second Site Request.`);
    else if (r.x3_exported_at) warnings.push(`${where} was already downloaded for X3 on ${String(r.x3_exported_at).slice(0, 10)}.`);
    return { id: r.id, ref, header, lines: out };
  });
  if (!records.length) errors.push("Choose at least one requisition.");
  return { ok: errors.length === 0, errors: [...new Set(errors)], warnings, records };
}

/** The import file: an E line per request then its L lines, quoted like X3's own export. */
export function x3SiteRequestFile(records, settings) {
  const s = { ...X3_DEFAULTS, ...settings };
  const q = (v) => `"${String(v ?? "").replace(/[\r\n]+/g, " ").replace(/"/g, "\"\"")}"`;
  const row = (marker, fields, values) => [marker, ...fields.map((f) => values[f.code])].map(q).join(s.separator);
  const out = [];
  for (const r of records) {
    out.push(row(s.header_marker, X3_HEADER_FIELDS, r.header));
    for (const l of r.lines) out.push(row(s.line_marker, X3_LINE_FIELDS, l));
  }
  return out.join("\r\n") + "\r\n";
}
