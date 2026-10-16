// IRONLOG/api/utils/developerAttachments.js — Developer tab: documents sent with a request.
//
// Claude's GitHub job only sees the issue text and the code on main, so an
// uploaded document is read here and its text goes into the issue. Text and
// code files are used as they are; Word (.docx) and Excel (.xlsx) files are
// turned into plain text. PDFs and pictures cannot be read this way yet.
import ExcelJS from "exceljs";
import JSZip from "jszip";

export const ATTACHMENT_LIMITS = {
  files: 5,
  fileBytes: 5 * 1024 * 1024,
  fileChars: 20000,
  totalChars: 40000,
};

const TEXT_TYPES = new Set([
  "txt", "md", "csv", "tsv", "json", "log", "xml", "html", "htm", "css",
  "js", "mjs", "cjs", "ts", "sql", "yml", "yaml", "ini",
]);
export const ACCEPTED_TYPES = [...TEXT_TYPES, "docx", "xlsx"];

function extensionOf(name) {
  const lower = String(name || "").toLowerCase();
  const dot = lower.lastIndexOf(".");
  return dot > 0 ? lower.slice(dot + 1) : "";
}

export function cleanFileName(name) {
  const base = String(name || "").split(/[\\/]/).pop();
  return base.replace(/[^\w .()+-]/g, "_").replace(/\s+/g, " ").trim().slice(0, 100) || "document";
}

function decodeXml(text) {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, "&");
}

async function docxText(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")?.async("string");
  if (xml == null) throw new Error("not a Word document");
  return decodeXml(xml
    .replace(/<w:tab\/>/g, "\t")
    .replace(/<w:br\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n")
    .replace(/<[^>]+>/g, ""))
    .replace(/\n{3,}/g, "\n\n");
}

async function xlsxText(buffer, maxChars) {
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(buffer);
  const lines = [];
  let size = 0;
  for (const sheet of book.worksheets) {
    lines.push(`--- Sheet: ${sheet.name} ---`);
    sheet.eachRow({ includeEmpty: false }, (row) => {
      if (size > maxChars) return;
      const cells = [];
      row.eachCell({ includeEmpty: true }, (cell) => cells.push(String(cell.text ?? "").replace(/[\t\n]+/g, " ")));
      const line = cells.join("\t");
      size += line.length + 1;
      lines.push(line);
    });
  }
  return lines.join("\n");
}

/** Read one uploaded document: { name, data_base64 } → { name, text, truncated } or { error }. */
export async function readAttachment(input = {}, maxChars = ATTACHMENT_LIMITS.fileChars) {
  const name = cleanFileName(input.name);
  const type = extensionOf(name);
  if (!ACCEPTED_TYPES.includes(type)) {
    return { error: `${name}: this kind of file cannot be sent to Claude. Use a text, Word (.docx) or Excel (.xlsx) file. For a PDF or picture, copy the text into the request.` };
  }
  const buffer = Buffer.from(String(input.data_base64 || ""), "base64");
  if (!buffer.length) return { error: `${name} is empty.` };
  if (buffer.length > ATTACHMENT_LIMITS.fileBytes) return { error: `${name} is bigger than 5 MB.` };
  let text;
  try {
    if (type === "docx") text = await docxText(buffer);
    else if (type === "xlsx") text = await xlsxText(buffer, maxChars);
    else text = buffer.toString("utf8");
  } catch {
    return { error: `${name} could not be opened. Check it is a real ${type === "docx" ? "Word" : "Excel"} file.` };
  }
  text = text.replace(/\r\n?/g, "\n").replace(/\u0000/g, "").trim();
  if (!text) return { error: `${name} has no text in it that Claude can read.` };
  const truncated = text.length > maxChars;
  return { name, text: truncated ? text.slice(0, maxChars) : text, truncated };
}

/** Read all the documents for one request, keeping within the size GitHub allows for an issue. */
export async function readAttachments(list) {
  if (list == null) return { attachments: [] };
  if (!Array.isArray(list)) return { error: "The documents were not sent correctly. Try adding them again." };
  if (list.length > ATTACHMENT_LIMITS.files) return { error: `Send at most ${ATTACHMENT_LIMITS.files} documents with one request.` };
  const attachments = [];
  let left = ATTACHMENT_LIMITS.totalChars;
  for (const item of list) {
    if (left <= 0) return { error: "The documents are too long for one request. Send fewer or shorter documents." };
    const doc = await readAttachment(item, Math.min(ATTACHMENT_LIMITS.fileChars, left));
    if (doc.error) return { error: doc.error };
    left -= doc.text.length;
    attachments.push(doc);
  }
  return { attachments };
}

/** The issue section holding the documents. The fence is longer than any run of backticks in the text. */
export function attachmentsMarkdown(attachments = []) {
  if (!attachments.length) return "";
  const parts = ["## Attached documents", "Read these with the request. They were uploaded from the IronLog Developer tab."];
  for (const doc of attachments) {
    const longest = Math.max(2, ...(doc.text.match(/`+/g) || []).map((m) => m.length));
    const fence = "`".repeat(longest + 1);
    parts.push(
      "",
      `<details><summary>${doc.name.replace(/[<>&]/g, "_")}${doc.truncated ? " (cut short, too long to send in full)" : ""}</summary>`,
      "",
      fence,
      doc.text,
      fence,
      "",
      "</details>",
    );
  }
  return parts.join("\n");
}
