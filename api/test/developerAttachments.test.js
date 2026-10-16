import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import {
  ATTACHMENT_LIMITS,
  attachmentsMarkdown,
  readAttachment,
  readAttachments,
} from "../utils/developerAttachments.js";
import { buildIssue } from "../utils/developerClaude.js";

const b64 = (text) => Buffer.from(text, "utf8").toString("base64");

async function wordFile(paragraphs) {
  const zip = new JSZip();
  const body = paragraphs.map((p) => `<w:p><w:r><w:t>${p}</w:t></w:r></w:p>`).join("");
  zip.file("word/document.xml", `<?xml version="1.0"?><w:document><w:body>${body}</w:body></w:document>`);
  return (await zip.generateAsync({ type: "nodebuffer" })).toString("base64");
}

async function excelFile() {
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet("Fuel");
  sheet.addRow(["Asset", "Litres"]);
  sheet.addRow(["EX-01", 420]);
  return Buffer.from(await book.xlsx.writeBuffer()).toString("base64");
}

test("text, Word and Excel documents are read as text", async () => {
  const txt = await readAttachment({ name: "notes.md", data_base64: b64("# Change\nAdd a column.\r\n") });
  assert.deepEqual(txt, { name: "notes.md", text: "# Change\nAdd a column.", truncated: false });

  const docx = await readAttachment({ name: "Spec.docx", data_base64: await wordFile(["Daily Log", "Add fuel cost &amp; litres"]) });
  assert.equal(docx.text, "Daily Log\nAdd fuel cost & litres");

  const xlsx = await readAttachment({ name: "fuel.xlsx", data_base64: await excelFile() });
  assert.match(xlsx.text, /Sheet: Fuel/);
  assert.match(xlsx.text, /EX-01\t420/);
});

test("documents Claude cannot read are refused with a plain message", async () => {
  assert.match((await readAttachment({ name: "scan.pdf", data_base64: b64("%PDF") })).error, /cannot be sent to Claude/);
  assert.match((await readAttachment({ name: "empty.txt", data_base64: "" })).error, /empty/);
  assert.match((await readAttachment({ name: "broken.docx", data_base64: b64("not a zip") })).error, /could not be opened/);
  const big = Buffer.alloc(ATTACHMENT_LIMITS.fileBytes + 1, 65).toString("base64");
  assert.match((await readAttachment({ name: "big.txt", data_base64: big })).error, /5 MB/);
});

test("file names are cleaned and long documents are cut short", async () => {
  const doc = await readAttachment({ name: "../../etc/<x>.txt", data_base64: b64("a".repeat(ATTACHMENT_LIMITS.fileChars + 50)) });
  assert.equal(doc.name, "_x_.txt");
  assert.equal(doc.truncated, true);
  assert.equal(doc.text.length, ATTACHMENT_LIMITS.fileChars);
});

test("a request takes at most five documents and stays within the issue size", async () => {
  assert.deepEqual(await readAttachments(undefined), { attachments: [] });
  const six = Array.from({ length: 6 }, (_, i) => ({ name: `${i}.txt`, data_base64: b64("hello") }));
  assert.match((await readAttachments(six)).error, /at most 5/);
  const long = Array.from({ length: 3 }, (_, i) => ({ name: `${i}.txt`, data_base64: b64("x".repeat(ATTACHMENT_LIMITS.fileChars)) }));
  assert.match((await readAttachments(long)).error, /too long/);

  const full = Array.from({ length: 2 }, (_, i) => ({ name: `${i}.txt`, data_base64: b64("y".repeat(ATTACHMENT_LIMITS.fileChars)) }));
  const { attachments } = await readAttachments(full);
  const issue = buildIssue({ title: "t", request_text: "z".repeat(8000), priority: "medium", model: "opus", created_by: "jaco", attachments });
  assert.ok(issue.body.length < 65536, "GitHub refuses issue bodies over 65536 characters");
});

test("documents go into the issue in a block their own backticks cannot break", () => {
  const md = attachmentsMarkdown([{ name: "code.md", text: "before\n```\ninside\n```\nafter", truncated: false }]);
  assert.match(md, /## Attached documents/);
  assert.match(md, /<summary>code\.md<\/summary>/);
  assert.match(md, /\n````\nbefore/);
  assert.match(md, /after\n````\n/);
  assert.equal(attachmentsMarkdown([]), "");
});
