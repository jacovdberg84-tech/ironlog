// Adds known AML plant rows while retaining the creator-protected workbook
// package. This intentionally edits only the required worksheet XML parts;
// a normal spreadsheet writer could discard pivots, external links or locks.
import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

const TEMPLATE_PATH = path.resolve("templates", "AML Weekly Check Sheet V27.xlsx");
const WEEKLY_SHEET_PATH = "xl/worksheets/sheet1.xml";
const EAM_BLUEPRINT_PATH = "xl/worksheets/sheet5.xml";

const WEEKLY_ROWS = [
  { row: 94, assetCode: "G02AM" },
  { row: 95, assetCode: "W200AM" },
  { row: 96, assetCode: "W201AM" },
];

// Include the three entries that were previously added to the weekly sheet as
// well, so all six rows have valid VLOOKUP data in the same template release.
const BLUEPRINT_ENTRIES = [
  { assetCode: "D600AM", category: "DOZER", description: "CAT D6R DOZER" },
  { assetCode: "E503AM", category: "EXCAVATOR", description: "CAT 349 EXCAVATOR" },
  { assetCode: "E504AM", category: "EXCAVATOR", description: "CAT 349 EXCAVATOR" },
  { assetCode: "G02AM", category: "GRADER", description: "CAT 140K MOTOR GRADER" },
  { assetCode: "W200AM", category: "WATER TRUCK", description: "BELL B25D 23000L WATER TRUCK" },
  { assetCode: "W201AM", category: "WATER TRUCK", description: "BELL B25D 23000L WATER TRUCK" },
];

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function cellPattern(ref) {
  return new RegExp(`<c\\b[^>]*?\\br="${escapeRegExp(ref)}"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/c>)`, "i");
}

function stringCell(ref, value, existingCell = "") {
  const originalOpen = String(existingCell).match(/^<c\\b[^>]*>/i)?.[0];
  const open = originalOpen
    ? originalOpen.replace(/\s+t="[^"]*"/i, "").replace(/\s*\/>$/, ">")
    : `<c r="${ref}">`;
  return `${open.replace(/>$/, ' t="inlineStr">')}<is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
}

function replaceCell(xml, ref, value) {
  const pattern = cellPattern(ref);
  if (!pattern.test(xml)) throw new Error(`Expected existing AML weekly cell ${ref}`);
  return xml.replace(pattern, (existingCell) => stringCell(ref, value, existingCell));
}

function valueFromCell(cellXml, sharedStrings) {
  const source = String(cellXml || "");
  const inline = source.match(/<is\b[^>]*>([\s\S]*?)<\/is>/i);
  if (inline) return inline[1].replace(/<[^>]+>/g, "").trim();
  const raw = source.match(/<v\b[^>]*>([\s\S]*?)<\/v>/i)?.[1] || "";
  return /\bt="s"/i.test(source) ? String(sharedStrings[Number(raw)] || "").trim() : raw.trim();
}

function sharedStringsFromXml(xml) {
  return Array.from(String(xml || "").matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/gi)).map((match) =>
    match[1]
      .replace(/<[^>]+>/g, "")
      .replace(/&apos;/g, "'")
      .replace(/&quot;/g, '"')
      .replace(/&gt;/g, ">")
      .replace(/&lt;/g, "<")
      .replace(/&amp;/g, "&")
  );
}

function blueprintRowByAssetCode(xml, sharedStrings) {
  const rows = new Map();
  for (const match of String(xml).matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/gi)) {
    const row = Number(match[1]);
    const rowXml = match[0];
    const codeCell = rowXml.match(cellPattern(`A${row}`))?.[0] || "";
    const code = valueFromCell(codeCell, sharedStrings).toUpperCase();
    if (code) rows.set(code, row);
  }
  return rows;
}

function maxWorksheetRow(xml) {
  return Math.max(...Array.from(String(xml).matchAll(/<row\b[^>]*\br="(\d+)"/gi)).map((match) => Number(match[1])), 0);
}

function appendBlueprintRow(xml, row, entry) {
  const cells = [
    stringCell(`A${row}`, entry.assetCode),
    stringCell(`C${row}`, entry.category),
    stringCell(`D${row}`, entry.description),
  ].join("");
  const newRow = `<row r="${row}">${cells}</row>`;
  if (!/<\/sheetData>/i.test(xml)) throw new Error("AML EAM BLUE PRINT sheet is missing sheet data");
  return xml.replace(/<\/sheetData>/i, `${newRow}</sheetData>`);
}

function extendWorksheetDimension(xml, lastRow) {
  return String(xml).replace(
    /<dimension\s+ref="([A-Z]+)\d+:([A-Z]+)\d+"\s*\/>/i,
    (_, firstColumn, lastColumn) => `<dimension ref="${firstColumn}1:${lastColumn}${lastRow}"/>`
  );
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

const input = await fs.readFile(TEMPLATE_PATH);
const zip = await JSZip.loadAsync(input);
const weeklyFile = zip.file(WEEKLY_SHEET_PATH);
const blueprintFile = zip.file(EAM_BLUEPRINT_PATH);
const sharedStringsFile = zip.file("xl/sharedStrings.xml");
const workbookFile = zip.file("xl/workbook.xml");
if (!weeklyFile || !blueprintFile || !sharedStringsFile || !workbookFile) {
  throw new Error("The AML template does not contain the required worksheets");
}

let weeklyXml = await weeklyFile.async("string");
for (const entry of WEEKLY_ROWS) weeklyXml = replaceCell(weeklyXml, `C${entry.row}`, entry.assetCode);

const sharedStrings = sharedStringsFromXml(await sharedStringsFile.async("string"));
let blueprintXml = await blueprintFile.async("string");
const existingRows = blueprintRowByAssetCode(blueprintXml, sharedStrings);
let nextRow = maxWorksheetRow(blueprintXml) + 1;
for (const entry of BLUEPRINT_ENTRIES) {
  if (existingRows.has(entry.assetCode)) continue;
  blueprintXml = appendBlueprintRow(blueprintXml, nextRow, entry);
  existingRows.set(entry.assetCode, nextRow);
  nextRow += 1;
}
blueprintXml = extendWorksheetDimension(blueprintXml, maxWorksheetRow(blueprintXml));

zip.file(WEEKLY_SHEET_PATH, weeklyXml);
zip.file(EAM_BLUEPRINT_PATH, blueprintXml);
zip.file("xl/workbook.xml", markWorkbookForRecalculation(await workbookFile.async("string")));
await fs.writeFile(
  TEMPLATE_PATH,
  await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } })
);

console.log(JSON.stringify({
  weeklyRows: WEEKLY_ROWS,
  blueprintRows: BLUEPRINT_ENTRIES.map((entry) => ({ assetCode: entry.assetCode, row: existingRows.get(entry.assetCode) })),
}, null, 2));
