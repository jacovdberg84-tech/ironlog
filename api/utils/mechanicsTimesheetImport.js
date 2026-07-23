// Parse mechanics timesheet XLSX / CSV uploads (same layout as timesheet export).
import { parse } from "csv-parse/sync";
import ExcelJS from "exceljs";

const MONTH_NAMES = [
  "january", "february", "march", "april", "may", "june",
  "july", "august", "september", "october", "november", "december",
];

const HEADER_ALIASES = {
  date: "date",
  work_date: "date",
  "plant no": "plant_no",
  plant_no: "plant_no",
  plant: "plant_no",
  asset_code: "plant_no",
  "work hours": "hours",
  work_hours: "hours",
  hours: "hours",
  category: "category",
  "description of work carried out": "reason",
  description: "reason",
  reason: "reason",
  "time started": "time_started",
  time_started: "time_started",
  "time finished": "time_finished",
  time_finished: "time_finished",
  technician: "technician_name",
  technician_name: "technician_name",
  "job card no": "job_card_no",
  job_card_no: "job_card_no",
  smr: "smr",
};

function normHeader(h) {
  return String(h || "")
    .replace(/^\uFEFF/, "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function mapHeaderKey(h) {
  return HEADER_ALIASES[normHeader(h)] || null;
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function toYmd(y, m, d) {
  if (!Number.isFinite(y) || !Number.isFinite(m) || !Number.isFinite(d)) return null;
  if (y < 1900 || y > 2100 || m < 1 || m > 12 || d < 1 || d > 31) return null;
  const out = `${y}-${pad2(m)}-${pad2(d)}`;
  const check = new Date(`${out}T12:00:00`);
  if (Number.isNaN(check.getTime())) return null;
  if (check.getFullYear() !== y || check.getMonth() + 1 !== m || check.getDate() !== d) return null;
  return out;
}

/** Excel serial date (days since 1899-12-30) → YYYY-MM-DD */
function excelSerialToYmd(n) {
  if (!Number.isFinite(n) || n < 20000 || n > 80000) return null;
  const epoch = Date.UTC(1899, 11, 30);
  const ms = epoch + Math.round(n) * 86400000;
  const d = new Date(ms);
  return toYmd(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
}

export function parseTimesheetDate(value) {
  if (value == null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return toYmd(value.getFullYear(), value.getMonth() + 1, value.getDate());
  }
  if (typeof value === "number") {
    return excelSerialToYmd(value);
  }
  const s = String(value).trim();
  if (!s) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  const display = s.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (display) {
    const d = Number(display[1]);
    const monthIdx = MONTH_NAMES.indexOf(display[2].toLowerCase());
    const y = Number(display[3]);
    if (monthIdx >= 0) return toYmd(y, monthIdx + 1, d);
  }

  const slash = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (slash) {
    const a = Number(slash[1]);
    const b = Number(slash[2]);
    const y = Number(slash[3]);
    // Prefer DMY (common for this site); if first > 12 treat as DMY, if second > 12 treat as MDY
    if (a > 12) return toYmd(y, b, a);
    if (b > 12) return toYmd(y, a, b);
    return toYmd(y, b, a);
  }

  const asNum = Number(s);
  if (Number.isFinite(asNum) && String(s).match(/^\d+(\.\d+)?$/)) {
    return excelSerialToYmd(asNum);
  }
  return null;
}

function cellText(value) {
  if (value == null) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object" && value.text != null) return String(value.text).trim();
  if (typeof value === "object" && value.result != null) return String(value.result).trim();
  return String(value).trim();
}

function normalizeRowObject(raw) {
  const out = {};
  for (const [k, v] of Object.entries(raw || {})) {
    const key = mapHeaderKey(k);
    if (!key) continue;
    out[key] = v;
  }
  return out;
}

function parseTimeValue(v) {
  const s = cellText(v);
  if (!s) return null;
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(s)) {
    const parts = s.split(":");
    return `${pad2(Number(parts[0]))}:${pad2(Number(parts[1]))}`;
  }
  // Excel fraction of day
  const n = Number(v);
  if (Number.isFinite(n) && n >= 0 && n < 1.5) {
    const totalMins = Math.round((n % 1) * 24 * 60);
    const hh = Math.floor(totalMins / 60) % 24;
    const mm = totalMins % 60;
    return `${pad2(hh)}:${pad2(mm)}`;
  }
  return s.slice(0, 16);
}

/**
 * Convert a normalized object into a labor entry, or return { error }.
 */
export function normalizeTimesheetImportRow(raw, lineNo) {
  const row = normalizeRowObject(raw);
  const work_date = parseTimesheetDate(row.date);
  const technician_name = cellText(row.technician_name);
  const asset_code = cellText(row.plant_no).toUpperCase();
  const reason = cellText(row.reason);
  const hours = Number(row.hours);
  const category = cellText(row.category) || null;
  const time_started = parseTimeValue(row.time_started);
  const time_finished = parseTimeValue(row.time_finished);
  const job_card_no = cellText(row.job_card_no) || null;
  const smrRaw = row.smr;
  const smr = smrRaw === "" || smrRaw == null ? null : Number(smrRaw);

  if (!work_date && !technician_name && !asset_code && !reason && !(Number.isFinite(hours) && hours > 0)) {
    return { skip: true };
  }

  const errors = [];
  if (!work_date) errors.push(`Row ${lineNo}: invalid or missing Date`);
  if (!technician_name) errors.push(`Row ${lineNo}: Technician required`);
  if (!asset_code) errors.push(`Row ${lineNo}: Plant no required`);
  if (!reason) errors.push(`Row ${lineNo}: Description Of Work Carried Out required`);
  if (!Number.isFinite(hours) || hours <= 0) errors.push(`Row ${lineNo}: Work Hours must be > 0`);

  if (errors.length) return { error: errors.join("; ") };

  return {
    entry: {
      work_date,
      technician_name,
      asset_code,
      reason,
      hours: Number(hours.toFixed(2)),
      category,
      time_started,
      time_finished,
      job_card_no,
      smr: Number.isFinite(smr) ? smr : null,
    },
  };
}

function detectCsvDelimiter(text) {
  const clean = text.replace(/^\uFEFF/, "");
  const firstLine = clean.split(/\r?\n/).find((l) => l.trim() !== "") || "";
  const commaCount = (firstLine.match(/,/g) || []).length;
  const semiCount = (firstLine.match(/;/g) || []).length;
  return semiCount > commaCount ? ";" : ",";
}

export function parseTimesheetCsvBuffer(buffer) {
  const text = buffer.toString("utf8");
  const records = parse(text, {
    columns: true,
    skip_empty_lines: true,
    bom: true,
    trim: true,
    delimiter: detectCsvDelimiter(text),
    relax_column_count: true,
  });
  return records.map((r, idx) => ({ raw: r, lineNo: idx + 2 }));
}

function worksheetToObjects(ws) {
  const rows = [];
  let headers = null;
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    const values = row.values || [];
    // ExcelJS row.values is 1-indexed
    const cells = [];
    for (let i = 1; i < values.length; i += 1) {
      cells.push(values[i]);
    }
    if (!headers) {
      headers = cells.map((c) => cellText(c));
      const mapped = headers.map(mapHeaderKey).filter(Boolean);
      if (!mapped.includes("date") || !mapped.includes("plant_no")) {
        // Not a timesheet header row — keep scanning
        headers = null;
      }
      return;
    }
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = cells[i] ?? "";
    });
    rows.push({ raw: obj, lineNo: rowNumber });
  });
  return rows;
}

/**
 * Read XLSX buffer. Prefer "Mechanics log"; otherwise merge all month/data sheets (skip Info).
 */
export async function parseTimesheetXlsxBuffer(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const preferred = wb.getWorksheet("Mechanics log");
  if (preferred) return worksheetToObjects(preferred);

  const skip = new Set(["info"]);
  const all = [];
  wb.eachSheet((ws) => {
    const name = String(ws.name || "").trim().toLowerCase();
    if (skip.has(name)) return;
    all.push(...worksheetToObjects(ws));
  });
  return all;
}

/**
 * Parse uploaded timesheet file into validated entries.
 * @returns {{ entries: object[], errors: string[], skipped: number }}
 */
export async function parseMechanicsTimesheetUpload(buffer, filename = "") {
  const name = String(filename || "").toLowerCase();
  const isCsv = name.endsWith(".csv") || name.endsWith(".txt");
  const isXlsx = name.endsWith(".xlsx") || name.endsWith(".xlsm");

  let rows;
  if (isCsv) {
    rows = parseTimesheetCsvBuffer(buffer);
  } else if (isXlsx || !isCsv) {
    try {
      rows = await parseTimesheetXlsxBuffer(buffer);
    } catch (e) {
      if (isXlsx) throw e;
      // Fallback: try CSV if extension unclear
      rows = parseTimesheetCsvBuffer(buffer);
    }
  } else {
    throw new Error("Upload an .xlsx or .csv file matching the timesheet layout");
  }

  const entries = [];
  const errors = [];
  let skipped = 0;
  for (const { raw, lineNo } of rows) {
    const result = normalizeTimesheetImportRow(raw, lineNo);
    if (result.skip) {
      skipped += 1;
      continue;
    }
    if (result.error) {
      errors.push(result.error);
      continue;
    }
    entries.push(result.entry);
  }
  return { entries, errors, skipped };
}
