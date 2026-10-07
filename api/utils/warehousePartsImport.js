// IRONLOG/api/utils/warehousePartsImport.js
// Reads the weekly Durban / Boksburg supplier warehouse update without trusting
// its formatting. The exported rows are intentionally plain data so the route
// can preview, match and upsert them safely.
import ExcelJS from "exceljs";

const HEADER_ALIASES = {
  site: ["site"],
  warehouse_code: ["whs status", "warehouse status", "warehouse"],
  priority_code: ["priority code", "priority"],
  supplier_name: ["supplier"],
  part_name: ["description", "part description"],
  part_code: ["part no", "part number", "part no."],
  qty_received: ["qty recvd", "qty received", "quantity received"],
  unit_cost: ["price", "unit price"],
  line_cost: ["cost", "line cost", "total cost"],
  fleet_reference: ["manual req nos fleet", "manual req no fleet", "manual req fleet", "fleet"],
  requisition_number: ["sage req number", "sage requisition number", "requisition number", "req number"],
  order_number: ["order no", "order number", "po number"],
  sales_order: ["sales order"],
  qty_ordered: ["order qty", "qty ordered", "order quantity"],
  outstanding_qty: ["outstanding", "qty outstanding"],
  uom: ["uom", "unit of measure"],
  order_date: ["order date"],
  warehouse_date: ["date waiting received", "date received", "warehouse date"],
  waiting_days: ["days waiting", "waiting days"],
  product_code: ["product code"],
  country_of_origin: ["country of origin", "origin"],
  comments: ["comments", "comment", "notes"],
  invoice_number: ["commercial invoice packing slip", "packing slip", "invoice number", "invoice"],
};

function normaliseHeader(value) {
  return textValue(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

export function textValue(value) {
  if (value == null) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object") {
    if (Array.isArray(value.richText)) return value.richText.map((part) => String(part?.text || "")).join("").trim();
    if (value.result != null) return textValue(value.result);
    if (value.text != null) return String(value.text).trim();
    if (value.hyperlink != null) return String(value.text || value.hyperlink).trim();
    return "";
  }
  return String(value).trim();
}

function numericValue(value) {
  if (value && typeof value === "object" && value.result != null) return numericValue(value.result);
  const raw = textValue(value).replace(/\s/g, "").replace(/,/g, "");
  if (!raw) return null;
  const n = Number(raw);
  return Number.isFinite(n) ? n : null;
}

function ymdValue(value) {
  if (value instanceof Date && Number.isFinite(value.getTime())) return value.toISOString().slice(0, 10);
  if (value && typeof value === "object" && value.result != null) return ymdValue(value.result);
  const raw = textValue(value);
  if (!raw) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const d = new Date(raw);
  return Number.isFinite(d.getTime()) ? d.toISOString().slice(0, 10) : null;
}

function findHeaderMap(values) {
  const cells = values.map(normaliseHeader);
  const result = {};
  for (const [key, aliases] of Object.entries(HEADER_ALIASES)) {
    const index = cells.findIndex((header) => aliases.includes(header));
    if (index >= 0) result[key] = index;
  }
  return result;
}

function headerScore(map) {
  return ["part_code", "part_name", "order_number", "order_date", "qty_ordered", "supplier_name"]
    .reduce((count, key) => count + (map[key] != null ? 1 : 0), 0);
}

function sourceKey(row) {
  const reference = row.order_number || row.sales_order || row.requisition_number || row.fleet_reference || row.order_date;
  const part = row.part_code || row.part_name;
  return `warehouse:${String(reference || "unreferenced").trim().toUpperCase()}:${String(part || "unknown").trim().toUpperCase()}`;
}

function cleanText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function cellAt(values, index) {
  return index == null ? null : values[index];
}

function mapRow(values, map, rowNumber) {
  const getText = (key) => cleanText(textValue(cellAt(values, map[key])));
  const getNumber = (key) => numericValue(cellAt(values, map[key]));
  const qty_received = getNumber("qty_received") ?? 0;
  const qty_ordered = getNumber("qty_ordered") ?? qty_received;
  const outstanding_qty = getNumber("outstanding_qty") ?? Math.max(0, Number(qty_ordered || 0) - Number(qty_received || 0));
  const part_code = getText("part_code");
  const part_name = getText("part_name");
  const order_date = ymdValue(cellAt(values, map.order_date));
  const warehouse_date = ymdValue(cellAt(values, map.warehouse_date));
  const row = {
    source_row: rowNumber,
    site: getText("site"),
    warehouse_code: getText("warehouse_code"),
    priority_code: getText("priority_code"),
    supplier_name: getText("supplier_name"),
    part_name,
    part_code,
    qty_received: Number(qty_received || 0),
    qty_ordered: Number(qty_ordered || 0),
    outstanding_qty: Number(Math.max(0, outstanding_qty || 0)),
    unit_cost: Number(getNumber("unit_cost") || 0),
    line_cost: Number(getNumber("line_cost") || 0),
    fleet_reference: getText("fleet_reference"),
    requisition_number: getText("requisition_number"),
    order_number: getText("order_number"),
    sales_order: getText("sales_order"),
    uom: getText("uom"),
    order_date,
    warehouse_date,
    waiting_days: getNumber("waiting_days"),
    product_code: getText("product_code"),
    country_of_origin: getText("country_of_origin"),
    comments: getText("comments"),
    invoice_number: getText("invoice_number"),
  };
  row.source_reference = sourceKey(row);
  return row;
}

export async function parseWarehousePartsWorkbook(buffer, filename = "warehouse-parts.xlsx") {
  if (!/\.xlsx$/i.test(String(filename || ""))) {
    return { rows: [], warnings: [], errors: ["Upload the supplier update as an .xlsx workbook."] };
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  let selected = null;
  for (const worksheet of workbook.worksheets) {
    const maximum = Math.min(Math.max(worksheet.rowCount, 1), 12);
    for (let number = 1; number <= maximum; number += 1) {
      const values = worksheet.getRow(number).values.slice(1);
      const map = findHeaderMap(values);
      const score = headerScore(map);
      if (!selected || score > selected.score) selected = { worksheet, rowNumber: number, map, score };
    }
  }

  if (!selected || selected.score < 3 || selected.map.part_code == null || selected.map.part_name == null) {
    return {
      rows: [],
      warnings: [],
      errors: ["Could not find the supplier header row. Expected at least Part No, Description, Order No and Order Date."],
    };
  }

  const rows = [];
  const warnings = [];
  const seen = new Set();
  for (let rowNumber = selected.rowNumber + 1; rowNumber <= selected.worksheet.rowCount; rowNumber += 1) {
    const values = selected.worksheet.getRow(rowNumber).values.slice(1);
    const row = mapRow(values, selected.map, rowNumber);
    if (!row.part_code && !row.part_name) continue;
    if (!row.order_date) {
      warnings.push(`Row ${rowNumber}: skipped because Order Date is blank or invalid.`);
      continue;
    }
    if (!(row.qty_ordered > 0 || row.qty_received > 0)) {
      warnings.push(`Row ${rowNumber}: skipped because quantity is zero.`);
      continue;
    }
    if (seen.has(row.source_reference)) {
      warnings.push(`Row ${rowNumber}: duplicate supplier line (${row.source_reference}) skipped.`);
      continue;
    }
    seen.add(row.source_reference);
    rows.push(row);
  }

  return {
    rows,
    warnings,
    errors: [],
    sheet_name: selected.worksheet.name,
    header_row: selected.rowNumber,
    date_range: rows.length
      ? {
          start: rows.map((row) => row.order_date).sort()[0],
          end: rows.map((row) => row.order_date).sort().at(-1),
        }
      : null,
  };
}
