// IRONLOG/api/utils/sageJournal.js — turn a summarised journal run into a batch Sage can import.
//
// A journal run (finance_posting_runs) holds a debit and a credit line for every
// cost IronLog knows about, posted to IronLog's own account names (5200-PARTS-EXPENSE…).
// Finance maps each IronLog account to their Sage account once, chooses which
// categories go to Sage (downtime is a notional cost, so it is off by default) and
// how the cost centre is sent. The export then nets the run per Sage account (and
// cost centre, or machine), dates it at the run's end date and refuses to produce
// a batch that does not balance.
//
// Two layouts:
//   sage300 — an Excel file with Journal_Headers and Journal_Details sheets, the
//             layout Sage 300 GL "Import Journal Entries" reads. The cost centre is
//             normally part of the Sage 300 account (e.g. 5200-PLANT), so "segment".
//   csv     — one journal line per row: Date, Reference, Account, Cost centre,
//             Description, Debit, Credit. For Sage Evolution's journal batch import
//             and other systems; the columns can be matched in their import screen.

// IronLog's posting accounts (see routes/procurement/journals.routes.js) and what they hold.
export const IRONLOG_ACCOUNTS = {
  "5100-FUEL-EXPENSE": { label: "Fuel used by machines", category: "fuel" },
  "1410-FUEL-INVENTORY": { label: "Fuel stock", category: "fuel" },
  "5150-LUBE-EXPENSE": { label: "Oil and lube used", category: "lube" },
  "1420-LUBE-INVENTORY": { label: "Oil and lube stock", category: "lube" },
  "5200-PARTS-EXPENSE": { label: "Parts used on machines", category: "parts" },
  "1400-INVENTORY": { label: "Parts stock", category: "parts" },
  "5300-LABOR-EXPENSE": { label: "Workshop labour", category: "labor" },
  "2500-LABOR-CLEARING": { label: "Labour clearing", category: "labor" },
  "5400-DOWNTIME-EXPENSE": { label: "Downtime cost (notional)", category: "downtime" },
  "2600-DOWNTIME-CLEARING": { label: "Downtime clearing", category: "downtime" },
  "2100-GRNI": { label: "Goods received not invoiced", category: "procurement" },
  "2200-AP": { label: "Supplier invoices (accounts payable)", category: "procurement" },
};

export const SAGE_CATEGORIES = {
  parts: "Parts",
  labor: "Labour",
  fuel: "Fuel",
  lube: "Oil and lube",
  downtime: "Downtime (notional)",
  procurement_grn: "Goods received",
  procurement_ap: "Supplier invoices",
};

export const SAGE_DEFAULTS = {
  layout: "sage300",
  accounts: Object.fromEntries(Object.keys(IRONLOG_ACCOUNTS).map((a) => [a, ""])),
  categories: { parts: true, labor: true, fuel: true, lube: true, downtime: false, procurement_grn: true, procurement_ap: true },
  cost_centre: "segment",
  separator: "-",
  default_cost_centre: "",
  detail: "account",
  date_format: "YYYY-MM-DD",
  reference_prefix: "IL",
};

const LAYOUTS = ["sage300", "csv"];
const CC_MODES = ["segment", "column", "none"];
const SEPARATORS = ["-", "/", ".", "_", ""];
const DETAILS = ["account", "machine"];
const DATE_FORMATS = ["YYYY-MM-DD", "DD/MM/YYYY", "YYYYMMDD"];
const CODE_RE = /^[A-Za-z0-9][A-Za-z0-9 ._/-]{0,44}$/;

export function ensureSageTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS sage_export_settings (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      settings_json TEXT NOT NULL,
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_by TEXT
    );
  `);
}

function merge(saved = {}) {
  return {
    ...SAGE_DEFAULTS,
    ...saved,
    accounts: { ...SAGE_DEFAULTS.accounts, ...(saved.accounts || {}) },
    categories: { ...SAGE_DEFAULTS.categories, ...(saved.categories || {}) },
  };
}

export function readSageSettings(db) {
  ensureSageTables(db);
  const row = db.prepare(`SELECT settings_json, updated_at, updated_by FROM sage_export_settings WHERE id = 1`).get();
  let saved = {};
  try { saved = JSON.parse(row?.settings_json || "{}") || {}; } catch { saved = {}; }
  return { ...merge(saved), updated_at: row?.updated_at || null, updated_by: row?.updated_by || null };
}

/** Checks and saves the settings: { ok, settings } or { ok:false, errors }. */
export function saveSageSettings(db, body = {}, by = "") {
  const cur = readSageSettings(db);
  const errors = [];
  const pick = (key, allowed, label) => {
    if (body[key] === undefined) return cur[key];
    const v = String(body[key]);
    if (!allowed.includes(v)) errors.push(`${label} must be one of: ${allowed.map((a) => a || "(none)").join(", ")}.`);
    return v;
  };
  const next = {
    layout: pick("layout", LAYOUTS, "Layout"),
    cost_centre: pick("cost_centre", CC_MODES, "Cost centre"),
    separator: pick("separator", SEPARATORS, "Separator"),
    detail: pick("detail", DETAILS, "Detail"),
    date_format: pick("date_format", DATE_FORMATS, "Date format"),
    accounts: { ...cur.accounts },
    categories: { ...cur.categories },
    default_cost_centre: cur.default_cost_centre,
    reference_prefix: cur.reference_prefix,
  };
  if (body.accounts && typeof body.accounts === "object") {
    for (const [k, raw] of Object.entries(body.accounts)) {
      if (!IRONLOG_ACCOUNTS[k]) { errors.push(`Unknown IronLog account: ${k}.`); continue; }
      const v = String(raw ?? "").trim();
      if (v && !CODE_RE.test(v)) errors.push(`The Sage account for "${IRONLOG_ACCOUNTS[k].label}" can only use letters, numbers, spaces and - . / _ (up to 45).`);
      next.accounts[k] = v;
    }
  }
  if (body.categories && typeof body.categories === "object") {
    for (const [k, v] of Object.entries(body.categories)) {
      if (!SAGE_CATEGORIES[k]) { errors.push(`Unknown category: ${k}.`); continue; }
      next.categories[k] = v === true || v === 1 || v === "1" || v === "true";
    }
  }
  if (body.default_cost_centre !== undefined) {
    const v = String(body.default_cost_centre ?? "").trim();
    if (v && !CODE_RE.test(v)) errors.push("The default cost centre can only use letters, numbers, spaces and - . / _ (up to 45).");
    next.default_cost_centre = v;
  }
  if (body.reference_prefix !== undefined) {
    const v = String(body.reference_prefix ?? "").trim();
    if (!/^[A-Za-z0-9_-]{0,12}$/.test(v)) errors.push("The reference prefix can be up to 12 letters, numbers, - or _.");
    next.reference_prefix = v;
  }
  if (!Object.values(next.categories).some(Boolean)) errors.push("Send at least one category to Sage.");
  if (errors.length) return { ok: false, errors };
  ensureSageTables(db);
  db.prepare(`
    INSERT INTO sage_export_settings (id, settings_json, updated_at, updated_by) VALUES (1, ?, datetime('now'), ?)
    ON CONFLICT(id) DO UPDATE SET settings_json = excluded.settings_json, updated_at = excluded.updated_at, updated_by = excluded.updated_by
  `).run(JSON.stringify(next), by || null);
  return { ok: true, settings: readSageSettings(db) };
}

export function formatSageDate(ymd, fmt) {
  const [y, m, d] = String(ymd || "").slice(0, 10).split("-");
  if (fmt === "DD/MM/YYYY") return `${d}/${m}/${y}`;
  if (fmt === "YYYYMMDD") return `${y}${m}${d}`;
  return `${y}-${m}-${d}`;
}

const cents = (n) => Math.sign(Number(n || 0)) * Math.round(Math.abs(Number(n || 0)) * 100);
const clip = (s, n) => (String(s).length > n ? String(s).slice(0, n) : String(s));

/**
 * The Sage batch for a run: { ok, errors, missing_accounts, reference, date, entries, total_debit, total_credit, rounding }.
 * lines are finance_posting_run_lines rows. Entries net debit and credit per Sage account
 * and cost centre (and machine when detail is "machine"); zero entries are dropped.
 */
export function buildSageJournal(run, lines, settings) {
  const s = merge(settings);
  const errors = [];
  const reference = `${s.reference_prefix || ""}${run.run_number}`.slice(0, 60);
  const date = String(run.end_date || "").slice(0, 10);
  if (String(run.status || "").toLowerCase() === "reversed") errors.push("This run was reversed; build a new run instead.");
  const used = (lines || []).filter((l) => s.categories[String(l.category || "").toLowerCase()]);
  const missing = [...new Set(used.map((l) => l.account_code).filter((a) => !String(s.accounts[a] || "").trim()))];
  if (missing.length) {
    errors.push(`Set the Sage account for: ${missing.map((a) => (IRONLOG_ACCOUNTS[a] ? `${IRONLOG_ACCOUNTS[a].label} (${a})` : a)).join(", ")}.`);
  }
  if (!used.length) errors.push("Nothing in this run goes to Sage with the categories chosen in the Sage settings.");
  if (errors.length) return { ok: false, errors, missing_accounts: missing, reference, date, entries: [] };

  const groups = new Map();
  for (const l of used) {
    const cc = s.cost_centre === "none" ? "" : String(l.cost_center_code || s.default_cost_centre || "").trim();
    const base = String(s.accounts[l.account_code]).trim();
    const account = s.cost_centre === "segment" && cc ? `${base}${s.separator}${cc}` : base;
    const machine = s.detail === "machine" ? String(l.asset_code || "").trim() : "";
    const key = [account, s.cost_centre === "column" ? cc : "", machine].join("\u0000");
    const g = groups.get(key) || { account, cost_centre: s.cost_centre === "column" ? cc : "", machine, ironlog: new Set(), amount: 0 };
    g.amount += Number(l.debit || 0) - Number(l.credit || 0);
    g.ironlog.add(l.account_code);
    groups.set(key, g);
  }
  const period = run.start_date && run.end_date ? ` ${run.start_date} to ${run.end_date}` : "";
  let entries = [...groups.values()]
    .map((g) => {
      const labels = [...g.ironlog].map((a) => IRONLOG_ACCOUNTS[a]?.label || a).join(" / ");
      return { account: g.account, cost_centre: g.cost_centre, machine: g.machine, c: cents(g.amount),
        description: clip(`${labels}${g.machine ? ` ${g.machine}` : ""}${period}`, 60) };
    })
    .filter((e) => e.c !== 0);
  // Rounding each entry to the cent can leave the batch a few cents out; that goes on the largest entry.
  const diff = entries.reduce((t, e) => t + e.c, 0);
  let rounding = 0;
  if (diff !== 0) {
    if (Math.abs(diff) > entries.length) {
      return { ok: false, errors: [`The run does not balance (${(diff / 100).toFixed(2)} USD out). Rebuild it before sending to Sage.`], missing_accounts: [], reference, date, entries: [] };
    }
    const big = entries.reduce((a, e) => (Math.abs(e.c) > Math.abs(a.c) ? e : a), entries[0]);
    big.c -= diff;
    rounding = -diff / 100;
    entries = entries.filter((e) => e.c !== 0);
  }
  entries.sort((a, b) => (b.c > 0) - (a.c > 0) || a.account.localeCompare(b.account) || a.cost_centre.localeCompare(b.cost_centre) || a.machine.localeCompare(b.machine));
  entries = entries.map(({ c, ...e }) => ({ ...e, amount: c / 100, debit: c > 0 ? c / 100 : 0, credit: c < 0 ? -c / 100 : 0 }));
  const total = (k) => Math.round(entries.reduce((t, e) => t + e[k] * 100, 0)) / 100;
  return { ok: true, errors: [], missing_accounts: [], reference, date, entries, total_debit: total("debit"), total_credit: total("credit"), rounding,
    description: clip(`IronLog ${run.run_number}${period}`, 60), layout: s.layout, date_format: s.date_format, cost_centre: s.cost_centre };
}

export function sageCsv(journal) {
  const esc = (v) => {
    const t = String(v ?? "");
    return /[",\n]/.test(t) ? `"${t.replace(/"/g, "\"\"")}"` : t;
  };
  const withCc = journal.cost_centre === "column";
  const header = ["Date", "Reference", "Account", ...(withCc ? ["Cost centre"] : []), "Description", "Debit", "Credit"];
  const date = formatSageDate(journal.date, journal.date_format);
  return [header.join(",")].concat(journal.entries.map((e) => [
    date, journal.reference, e.account, ...(withCc ? [e.cost_centre] : []), e.description,
    e.debit ? e.debit.toFixed(2) : "", e.credit ? e.credit.toFixed(2) : "",
  ].map(esc).join(","))).join("\r\n") + "\r\n";
}

/** The Sage 300 GL import workbook: one batch, one journal entry, one detail line per entry. */
export function sage300Workbook(ExcelJS, journal) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "IRONLOG";
  const [y, m, d] = journal.date.split("-").map(Number);
  const when = new Date(Date.UTC(y, m - 1, d));
  const head = wb.addWorksheet("Journal_Headers");
  head.columns = ["BATCHID", "BTCHENTRY", "SRCELEDGER", "SRCETYPE", "DATEENTRY", "JRNLDESC"].map((h) => ({ header: h, key: h, width: h === "JRNLDESC" ? 44 : 14 }));
  head.addRow({ BATCHID: "000001", BTCHENTRY: "00001", SRCELEDGER: "GL", SRCETYPE: "JE", DATEENTRY: when, JRNLDESC: journal.description });
  head.getColumn("DATEENTRY").numFmt = "yyyy-mm-dd";
  const det = wb.addWorksheet("Journal_Details");
  det.columns = ["BATCHNBR", "JOURNALID", "TRANSNBR", "ACCTID", "TRANSAMT", "TRANSDESC", "TRANSREF", "TRANSDATE", "SRCELDGR", "SRCETYPE"]
    .map((h) => ({ header: h, key: h, width: h === "TRANSDESC" ? 50 : h === "ACCTID" ? 22 : 14 }));
  journal.entries.forEach((e, i) => det.addRow({
    BATCHNBR: "000001", JOURNALID: "00001", TRANSNBR: String((i + 1) * 20).padStart(10, "0"), ACCTID: e.account,
    TRANSAMT: e.amount, TRANSDESC: e.description, TRANSREF: journal.reference, TRANSDATE: when, SRCELDGR: "GL", SRCETYPE: "JE",
  }));
  det.getColumn("TRANSAMT").numFmt = "#,##0.00";
  det.getColumn("TRANSDATE").numFmt = "yyyy-mm-dd";
  for (const ws of [head, det]) ws.getRow(1).font = { bold: true };
  return wb;
}
