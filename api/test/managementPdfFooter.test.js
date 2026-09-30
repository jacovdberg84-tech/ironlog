import test from "node:test";
import assert from "node:assert/strict";
import zlib from "node:zlib";
import { buildPdfBuffer, pdfDocumentInfo } from "../utils/pdfGenerator.js";

/** Visible text of a pdfkit PDF (hex-encoded Tj strings in deflated content streams). */
function pdfText(buf) {
  const s = buf.toString("latin1");
  let out = "";
  const re = /stream\r?\n/g;
  let m;
  while ((m = re.exec(s))) {
    const start = m.index + m[0].length;
    const end = s.indexOf("endstream", start);
    try {
      const d = zlib.inflateSync(buf.subarray(start, end)).toString("latin1");
      for (const h of d.matchAll(/<([0-9a-fA-F]{2,})>/g)) out += Buffer.from(h[1], "hex").toString("latin1");
    } catch {}
  }
  return out;
}

test("management report footer shows the operating day without naming the source system", async () => {
  const pdf = await buildPdfBuffer((doc) => { doc.text("Body"); }, {
    title: "Daily",
    managementTitle: "AML / DAILY OPERATIONS",
    subtitle: "Operating date 29 September 2026",
    sourceText: "Operating day 2026-09-29",
    headerStyle: "management",
    showPageNumbers: true,
    layout: "landscape",
  });
  const text = pdfText(pdf);
  assert.match(text, /Operating day 2026-09-29/);
  assert.doesNotMatch(text, /Source:/);
  assert.doesNotMatch(text, /Ironlog/i);
});

test("file properties carry the company name, not the software", async () => {
  const opts = { title: "IRONLOG", managementTitle: "AML / DAILY OPERATIONS", subtitle: "Operating date 29 September 2026" };
  assert.deepEqual(pdfDocumentInfo(opts, { companyName: "Aggregates Mozambique Limitada", siteName: "AML" }), {
    Title: "AML / DAILY OPERATIONS — Operating date 29 September 2026",
    Author: "Aggregates Mozambique Limitada",
    Creator: "Aggregates Mozambique Limitada",
    Producer: "Aggregates Mozambique Limitada",
  });
  assert.equal(pdfDocumentInfo({ title: "IRONLOG" }, { siteName: "AML" }).Author, "AML");
  assert.deepEqual(pdfDocumentInfo({ title: "IRONLOG" }), { Title: "Report", Creator: "Maintenance reporting", Producer: "Maintenance reporting" });

  const pdf = await buildPdfBuffer((doc) => { doc.text("Body"); }, { ...opts, companyName: "Aggregates Mozambique Limitada", headerStyle: "management" });
  const raw = pdf.toString("latin1");
  assert.doesNotMatch(raw, /PDFKit/, "the library name no longer shows in the file properties");
  assert.match(raw, /\/Author/);
});
