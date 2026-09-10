// Builds a prefilled copy of the protected AML weekly check-sheet template.
//
// This deliberately updates only sheet1's unlocked input cells inside the XLSX
// package. Repacking the complete workbook through a normal Excel writer would
// risk dropping the template's external links, pivot cache and protection.
import JSZip from "jszip";

export const AML_WEEKLY_SHEET_NAME = "WEEKLY CHECK SHEET V27";
export const AML_WEEKLY_SHEET_PATH = "xl/worksheets/sheet1.xml";
export const AML_WEEKLY_TEMPLATE_ROWS = { first: 14, last: 250 };

const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function decodeXml(value) {
  return String(value ?? "")
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function xmlText(fragment) {
  return decodeXml(
    String(fragment ?? "")
      .replace(/<br\s*\/>/gi, "\n")
      .replace(/<[^>]+>/g, "")
  );
}

function sharedStringsFromXml(xml) {
  return Array.from(String(xml || "").matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/gi)).map((match) => xmlText(match[1]));
}

function cellPattern(ref) {
  return new RegExp(
    `<c\\b[^>]*?\\br="${escapeRegExp(ref)}"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`,
    "i"
  );
}

function cellValue(xml, ref, sharedStrings = []) {
  const match = String(xml || "").match(cellPattern(ref));
  if (!match) return "";
  const cell = match[0];
  if (/\/>$/.test(cell)) return "";
  const inline = cell.match(/<is\b[^>]*>([\s\S]*?)<\/is>/i);
  if (inline) return xmlText(inline[1]);
  const raw = cell.match(/<v\b[^>]*>([\s\S]*?)<\/v>/i)?.[1] ?? "";
  if (/\bt="s"/i.test(cell)) return sharedStrings[Number(raw)] ?? "";
  return decodeXml(raw);
}

function cleanCellStartTag(cellXml) {
  const open = String(cellXml).match(/^<c\b[^>]*>/i)?.[0];
  if (!open) throw new Error("AML template input cell could not be read");
  return open
    .replace(/\/>$/, ">")
    .replace(/\s+t="[^"]*"/i, "");
}

function replaceTemplateCell(xml, ref, value, kind = "string") {
  const pattern = cellPattern(ref);
  if (!pattern.test(xml)) {
    throw new Error(`AML template is missing expected input cell ${ref}`);
  }
  return xml.replace(pattern, (cell) => {
    const open = cleanCellStartTag(cell);
    if (value == null || value === "") return `${open}</c>`;
    if (kind === "number") {
      const numeric = Number(value);
      if (!Number.isFinite(numeric)) return `${open}</c>`;
      return `${open}<v>${numeric}</v></c>`;
    }
    return `${open.replace(/>$/, ' t="inlineStr">')}<is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
  });
}

function excelDateSerial(ymd) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(ymd || ""))) return null;
  const time = Date.parse(`${ymd}T00:00:00Z`);
  return Number.isFinite(time) ? Math.round((time - EXCEL_EPOCH_UTC) / 86400000) : null;
}

function markWorkbookForRecalculation(xml) {
  const source = String(xml || "");
  if (/<calcPr\b/i.test(source)) {
    return source.replace(/<calcPr\b([^>]*)\/?>(?:<\/calcPr>)?/i, (tag, attrs) => {
      const withoutFlags = String(attrs || "")
        .replace(/\s+fullCalcOnLoad="[^"]*"/i, "")
        .replace(/\s+forceFullCalc="[^"]*"/i, "")
        .replace(/\/\s*$/, "");
      return `<calcPr${withoutFlags} fullCalcOnLoad="1" forceFullCalc="1"/>`;
    });
  }
  return source.replace(/<\/workbook>/i, '<calcPr fullCalcOnLoad="1" forceFullCalc="1"/></workbook>');
}

function cleanText(value, maxLength = 280) {
  const text = String(value ?? "").replace(/\s+/g, " ").trim();
  return text.length > maxLength ? `${text.slice(0, Math.max(1, maxLength - 1))}…` : text;
}

/** AML's working-status list is deliberately used verbatim so its formulas continue to classify rows. */
export function amlWorkingStatus({ breakdown = null, offsite = null, active = true } = {}) {
  if (breakdown) {
    const critical = Number(breakdown.critical || 0) === 1;
    if (offsite) return critical ? "Off Hire Major Breakdown on site" : "Off Hire Minor Breakdown on site";
    return critical ? "On Hire Breakdown Major Repairs " : "On Hire Breakdown Minor Repairs ";
  }
  return offsite || !active ? "Off Hire Available at site " : "On Hire Working";
}

export function amlBreakdownCategory(value) {
  const text = cleanText(value, 120).toUpperCase();
  if (!text) return "";
  if (/ELECTR|WIRING|ALTERNATOR|STARTER/.test(text)) return "ELECTRICAL";
  if (/ENGINE|MOTOR/.test(text)) return "ENGINE";
  if (/HYDRAUL/.test(text)) return "HYDRAULICS";
  if (/TRANSM|GEARBOX|TORQUE/.test(text)) return "TRANSMISSION";
  if (/TYRE|TIRE|WHEEL/.test(text)) return "TYRES";
  if (/BRAKE/.test(text)) return "BRAKES";
  if (/UNDERCARRIAGE|TRACK|ROLLER|SPROCKET/.test(text)) return "UNDERCARRIAGE";
  return text;
}

export function amlAdditionalComments({ breakdown = null, offsite = null } = {}) {
  const lines = [];
  if (breakdown?.work_order_id) {
    lines.push(`WO #${breakdown.work_order_id}${breakdown.work_order_status ? ` — ${breakdown.work_order_status}` : ""}`);
  }
  if (breakdown?.repair_progress) lines.push(`Repair progress: ${cleanText(breakdown.repair_progress, 100)}`);
  if (offsite) {
    const vendor = cleanText(offsite.vendor, 80);
    const sent = String(offsite.sent_date || "").trim();
    const detail = [
      "Offsite repair",
      vendor ? `vendor ${vendor}` : "",
      sent ? `sent ${sent}` : "",
      cleanText(offsite.repair_status, 60),
    ].filter(Boolean).join(" — ");
    lines.push(detail);
    if (offsite.notes) lines.push(cleanText(offsite.notes, 120));
  }
  return cleanText(lines.join(". "), 280);
}

/**
 * Creates a new workbook buffer from the creator-supplied AML template.
 * Only the unlocked weekly-entry fields are changed; all other XLSX parts
 * remain in the package, including the template's protected layout.
 */
export async function buildAmlWeeklyCheckSheet(templateBuffer, { weekEnding, records = [] } = {}) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(weekEnding || ""))) {
    throw new Error("week_ending must be YYYY-MM-DD");
  }
  const zip = await JSZip.loadAsync(templateBuffer);
  const sheetFile = zip.file(AML_WEEKLY_SHEET_PATH);
  const stringsFile = zip.file("xl/sharedStrings.xml");
  const workbookFile = zip.file("xl/workbook.xml");
  if (!sheetFile || !stringsFile || !workbookFile) {
    throw new Error("AML template is incomplete. Please upload the original AML V27 workbook again.");
  }

  let sheetXml = await sheetFile.async("string");
  const sharedStrings = sharedStringsFromXml(await stringsFile.async("string"));
  const rowByAssetCode = new Map();
  for (let row = AML_WEEKLY_TEMPLATE_ROWS.first; row <= AML_WEEKLY_TEMPLATE_ROWS.last; row += 1) {
    const code = cleanText(cellValue(sheetXml, `C${row}`, sharedStrings), 80).toUpperCase();
    if (code && !rowByAssetCode.has(code)) rowByAssetCode.set(code, row);
  }

  const unmatchedAssetCodes = [];
  let filledRows = 0;
  for (const record of records) {
    const assetCode = cleanText(record?.assetCode, 80).toUpperCase();
    if (!assetCode) continue;
    const row = rowByAssetCode.get(assetCode);
    if (!row) {
      unmatchedAssetCodes.push(assetCode);
      continue;
    }

    const breakdown = record.breakdown || null;
    const offsite = record.offsite || null;
    sheetXml = replaceTemplateCell(sheetXml, `D${row}`, "YES");
    if (Number.isFinite(Number(record.meterHours)) && Number(record.meterHours) > 0) {
      sheetXml = replaceTemplateCell(sheetXml, `J${row}`, Number(record.meterHours), "number");
    }
    sheetXml = replaceTemplateCell(
      sheetXml,
      `O${row}`,
      amlWorkingStatus({ breakdown, offsite, active: Number(record.active ?? 1) === 1 })
    );
    sheetXml = replaceTemplateCell(
      sheetXml,
      `R${row}`,
      breakdown ? amlBreakdownCategory(breakdown.component || breakdown.description) : ""
    );
    sheetXml = replaceTemplateCell(sheetXml, `S${row}`, breakdown ? cleanText(breakdown.description, 240) : "");
    const expectedDate = breakdown?.ets_repair_date || offsite?.expected_return_date || "";
    sheetXml = replaceTemplateCell(sheetXml, `T${row}`, excelDateSerial(expectedDate), "number");
    sheetXml = replaceTemplateCell(
      sheetXml,
      `U${row}`,
      amlAdditionalComments({ breakdown, offsite })
    );
    filledRows += 1;
  }

  sheetXml = replaceTemplateCell(sheetXml, "D7", excelDateSerial(weekEnding), "number");
  zip.file(AML_WEEKLY_SHEET_PATH, sheetXml);
  zip.file("xl/workbook.xml", markWorkbookForRecalculation(await workbookFile.async("string")));

  const buffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
  return {
    buffer,
    filledRows,
    templateRows: rowByAssetCode.size,
    unmatchedAssetCodes: Array.from(new Set(unmatchedAssetCodes)).sort(),
  };
}
