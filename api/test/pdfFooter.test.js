import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-pdf-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
const { db } = await import("../db/client.js");
const { buildPdfBuffer, nowStamp } = await import("../utils/pdfGenerator.js");
process.on("exit", () => { try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

test("report footers show the time on the site's clock", () => {
  const at = new Date("2026-10-16T22:30:00Z");
  assert.equal(nowStamp("Africa/Maputo", at), "2026-10-17 00:30", "Maputo is two hours ahead of UTC");
  assert.equal(nowStamp("UTC", at), "2026-10-16 22:30");
  assert.match(nowStamp("Not/AZone", at), /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/, "a bad time zone still prints a time");
});

test("every page is numbered 'Page X of Y'", async () => {
  const PDFDocument = (await import("pdfkit")).default;
  const written = [];
  const text = PDFDocument.prototype.text;
  PDFDocument.prototype.text = function (s, ...rest) { written.push(String(s)); return text.call(this, s, ...rest); };
  try {
    const pdf = await buildPdfBuffer((doc) => { doc.text("one"); doc.addPage(); doc.text("two"); doc.addPage(); doc.text("three"); },
      { title: "Test", timezone: "Africa/Maputo" });
    assert.equal(pdf.subarray(0, 4).toString(), "%PDF");
  } finally {
    PDFDocument.prototype.text = text;
  }
  assert.deepEqual(written.filter((s) => s.startsWith("Page ")), ["Page 1 of 3", "Page 2 of 3", "Page 3 of 3"]);
  assert.ok(written.some((s) => /^Generated: .* \(Africa\/Maputo\)$/.test(s)));
});
