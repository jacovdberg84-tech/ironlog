import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import JSZip from "jszip";
import {
  AML_WEEKLY_SHEET_PATH,
  amlBreakdownCategory,
  amlWorkingStatus,
  buildAmlWeeklyCheckSheet,
} from "../utils/amlWeeklyCheckSheet.js";

const templatePath = path.resolve("templates", "AML Weekly Check Sheet V27.xlsx");

function xmlText(fragment) {
  return String(fragment || "")
    .replace(/<[^>]+>/g, "")
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
}

function sharedStrings(xml) {
  return Array.from(String(xml || "").matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/gi)).map((m) => xmlText(m[1]));
}

function cell(xml, ref) {
  const match = String(xml).match(new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)`, "i"));
  return match?.[0] || "";
}

function cellValue(xml, ref, strings) {
  const source = cell(xml, ref);
  const inline = source.match(/<is\b[^>]*>([\s\S]*?)<\/is>/i);
  if (inline) return xmlText(inline[1]);
  const value = source.match(/<v\b[^>]*>([\s\S]*?)<\/v>/i)?.[1] || "";
  return /\bt="s"/i.test(source) ? strings[Number(value)] || "" : xmlText(value);
}

test("AML weekly check sheet keeps protected template features while filling permitted cells", async () => {
  const template = await fs.readFile(templatePath);
  const originalZip = await JSZip.loadAsync(template);
  const sourceSheet = await originalZip.file(AML_WEEKLY_SHEET_PATH).async("string");
  const strings = sharedStrings(await originalZip.file("xl/sharedStrings.xml").async("string"));
  const assetCode = cellValue(sourceSheet, "C14", strings);
  const secondAssetCode = cellValue(sourceSheet, "C15", strings);
  assert.ok(assetCode, "template needs a plant number in its first entry row");
  assert.ok(secondAssetCode, "template needs a plant number in its second entry row");

  const result = await buildAmlWeeklyCheckSheet(template, {
    weekEnding: "2026-09-11",
    records: [{
      assetCode,
      meterHours: 23456.7,
      active: 1,
      breakdown: {
        critical: 1,
        component: "Engine",
        description: "Cylinder head repair",
        ets_repair_date: "2026-09-14",
        work_order_id: 123,
        work_order_status: "in_progress",
      },
    }, {
      assetCode: secondAssetCode,
      meterHours: 12345.6,
      active: 1,
      isOperational: true,
      breakdown: {
        critical: 1,
        component: "Engine",
        description: "Legacy work order still awaiting close-out",
        ets_repair_date: "2026-09-14",
      },
    }],
  });

  assert.equal(result.filledRows, 2);
  const outputZip = await JSZip.loadAsync(result.buffer);
  const outputSheet = await outputZip.file(AML_WEEKLY_SHEET_PATH).async("string");
  assert.match(cell(outputSheet, "D7"), /<v>\d+<\/v>/);
  assert.match(cell(outputSheet, "J14"), /<v>23456\.7<\/v>/);
  assert.equal(cellValue(outputSheet, "O14", []), "On Hire Breakdown Major Repairs ");
  assert.equal(cellValue(outputSheet, "R14", []), "ENGINE");
  assert.equal(cellValue(outputSheet, "S14", []), "Cylinder head repair");
  assert.equal(cellValue(outputSheet, "O15", []), "On Hire Working");
  assert.equal(cellValue(outputSheet, "R15", []), "");
  assert.equal(cellValue(outputSheet, "S15", []), "");
  assert.match(cell(outputSheet, "J15"), /<v>12345\.6<\/v>/);
  assert.match(outputSheet, /<sheetProtection\b/i);
  assert.ok(outputZip.file("xl/externalLinks/externalLink1.xml"), "external link should remain in the template copy");
  assert.ok(outputZip.file("xl/pivotCache/pivotCacheDefinition1.xml"), "pivot cache should remain in the template copy");
});

test("AML template includes the new grader and water-truck rows with working lookup data", async () => {
  const template = await fs.readFile(templatePath);
  const inputZip = await JSZip.loadAsync(template);
  const inputSheet = await inputZip.file(AML_WEEKLY_SHEET_PATH).async("string");
  const blueprintSheet = await inputZip.file("xl/worksheets/sheet5.xml").async("string");
  const strings = sharedStrings(await inputZip.file("xl/sharedStrings.xml").async("string"));
  const expectedRows = [
    { row: 94, code: "G02AM", category: "GRADER", description: "CAT 140K MOTOR GRADER" },
    { row: 95, code: "W200AM", category: "WATER TRUCK", description: "BELL B25D 23000L WATER TRUCK" },
    { row: 96, code: "W201AM", category: "WATER TRUCK", description: "BELL B25D 23000L WATER TRUCK" },
  ];

  for (const entry of expectedRows) {
    assert.equal(cellValue(inputSheet, `C${entry.row}`, strings), entry.code);
    const sourceRow = Array.from({ length: 1502 }, (_, index) => index + 1).find(
      (row) => cellValue(blueprintSheet, `A${row}`, strings) === entry.code
    );
    assert.ok(sourceRow, `EAM BLUE PRINT needs ${entry.code}`);
    assert.equal(cellValue(blueprintSheet, `C${sourceRow}`, strings), entry.category);
    assert.equal(cellValue(blueprintSheet, `D${sourceRow}`, strings), entry.description);
  }

  const result = await buildAmlWeeklyCheckSheet(template, {
    weekEnding: "2026-09-11",
    records: expectedRows.map((entry, index) => ({
      assetCode: entry.code,
      meterHours: 7000 + index,
      active: 1,
      isOperational: true,
    })),
  });
  assert.equal(result.filledRows, expectedRows.length);
  assert.deepEqual(result.unmatchedAssetCodes, []);

  const outputZip = await JSZip.loadAsync(result.buffer);
  const outputSheet = await outputZip.file(AML_WEEKLY_SHEET_PATH).async("string");
  for (const entry of expectedRows) {
    assert.equal(cellValue(outputSheet, `D${entry.row}`, []), "YES");
    assert.equal(cellValue(outputSheet, `O${entry.row}`, []), "On Hire Working");
  }
  assert.match(outputSheet, /<sheetProtection\b/i);
  assert.ok(outputZip.file("xl/externalLinks/externalLink1.xml"));
  assert.ok(outputZip.file("xl/pivotCache/pivotCacheDefinition1.xml"));
  assert.match(blueprintSheet, /<dimension ref="A1:DS1502"\/>/i);
});

test("AML status and category mapping match the protected template choices", () => {
  assert.equal(amlWorkingStatus({}), "On Hire Working");
  assert.equal(amlWorkingStatus({ offsite: { id: 1 } }), "Off Hire Available at site ");
  assert.equal(
    amlWorkingStatus({ breakdown: { critical: 1 } }),
    "On Hire Breakdown Major Repairs "
  );
  assert.equal(amlBreakdownCategory("hydraulic hose leak"), "HYDRAULICS");
  assert.equal(amlBreakdownCategory("starter motor fault"), "ELECTRICAL");
});
