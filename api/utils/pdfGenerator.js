// IRONLOG/api/utils/pdfGenerator.js
import PDFDocument from "pdfkit";
import fs from "node:fs";
import path from "node:path";

const DEFAULT_MARGINS = { top: 56, bottom: 52, left: 70, right: 70 };
const BRAND = {
  primary: "#0b3a7e",
  primarySoft: "#dbeafe",
  text: "#0f172a",
  muted: "#334155",
  line: "#93c5fd",
  zebra: "#f8fafc",
};

function contentWidth(doc) {
  return doc.page.width - doc.page.margins.left - doc.page.margins.right;
}

function nowStamp() {
  // stable, readable timestamp in report footer
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(
    d.getMinutes()
  )}`;
}

/**
 * Ensures there's space for "neededHeight" on the page. If not, adds a new page.
 */
export function ensurePageSpace(doc, neededHeight = 60) {
  const bottomY = doc.page.height - doc.page.margins.bottom;
  if (doc.y + neededHeight > bottomY) doc.addPage();
}

/**
 * Draws consistent header and footer on the current page.
 * Call this ONCE per page, typically from doc.on("pageAdded", ...) and at start.
 */
export function drawHeaderFooter(doc, opts = {}) {
  const {
    title = "IRONLOG",
    subtitle = "Daily Report",
    rightText = "",
    showPageNumbers = true,
  } = opts;
  const displayTitle = String(title || "").trim().toUpperCase() === "IRONLOG" ? "AML" : title;

  // Important: header/footer drawing must NOT move the main content cursor.
  // PDFKit's text() updates doc.y, so we snapshot and restore it.
  const savedY = doc.y;

  const left = doc.page.margins.left;
  const right = doc.page.width - doc.page.margins.right;
  const top = doc.page.margins.top;

  // HEADER
  doc.save();

  // Header background bar (subtle)
  const barH = 32;
  doc
    .rect(left, top - 38, right - left, barH)
    .fillOpacity(1)
    .fill(BRAND.primarySoft);

  doc.fillOpacity(1);

  // Title left
  doc
    .fillColor(BRAND.primary)
    .font("Helvetica-Bold")
    .fontSize(14)
    .text(displayTitle, left + 10, top - 32, { width: (right - left) * 0.55 });

  doc
    .fillColor(BRAND.muted)
    .font("Helvetica")
    .fontSize(10)
    .text(subtitle, left + 10, top - 16, { width: (right - left) * 0.55 });

  // Right text (e.g. Date: YYYY-MM-DD)
  if (rightText) {
    doc
      .fillColor("#333333")
      .fillColor(BRAND.muted)
      .font("Helvetica")
      .fontSize(10)
      .text(rightText, left, top - 24, { width: right - left - 10, align: "right" });
  }

  // Divider line
  doc
    .moveTo(left, top - 6)
    .lineTo(right, top - 6)
    .lineWidth(1)
    .strokeOpacity(1)
    .stroke(BRAND.line);

  doc.restore();

  // FOOTER
  doc.save();
  // Keep footer INSIDE the content box. Writing below the bottom margin can
  // trigger PDFKit to auto-add blank pages.
  const footerY = doc.page.height - doc.page.margins.bottom - 16;

  doc
    .moveTo(left, footerY - 8)
    .lineTo(right, footerY - 8)
    .lineWidth(1)
    .strokeOpacity(0.15)
    .stroke("#000000");

  doc.fillOpacity(1).fillColor(BRAND.muted).font("Helvetica").fontSize(8);

  doc.text(`Generated: ${nowStamp()}`, left, footerY, {
    width: (right - left) * 0.6,
    align: "left",
    lineBreak: false,
  });

  if (showPageNumbers) {
    const pageNo = doc.page?.number || 1; // pdfkit page number exists on recent versions
    doc.text(`Page ${pageNo}`, left, footerY, { width: right - left, align: "right", lineBreak: false });
  }

  doc.restore();

  // Restore cursor and keep it below header
  doc.y = Math.max(savedY, doc.page.margins.top + 6);
}

/**
 * PDF builder wrapper (buffer output)
 */
export function buildPdfBuffer(buildFn, opts = {}) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: "A4",
      layout: opts.layout === "portrait" ? "portrait" : "landscape",
      margins: DEFAULT_MARGINS,
      bufferPages: true,
    });

    const chunks = [];
    doc.on("data", (c) => chunks.push(c));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);

    // Header/footer options (drawn after content to avoid PDFKit page churn)
    const headerOpts = {
      title: opts.title || "IRONLOG",
      subtitle: opts.subtitle || "Daily Report",
      rightText: opts.rightText || "",
      showPageNumbers: opts.showPageNumbers !== false,
    };

    buildFn(doc);

    if (!opts.disableHeaderFooter) {
      try {
        const range = doc.bufferedPageRange(); // { start, count }
        for (let i = range.start; i < range.start + range.count; i++) {
          doc.switchToPage(i);
          drawHeaderFooter(doc, headerOpts);
        }
      } catch {
        // if bufferPages/switchToPage not supported, skip header/footer gracefully
      }
    }

    doc.end();
  });
}

/**
 * Logo (optional)
 */
export function tryDrawLogo(doc, logoPath) {
  if (!logoPath) return;
  try {
    let resolved = logoPath;
    if (!fs.existsSync(resolved)) {
      const fallback = path.resolve(process.cwd(), "../branding/logo.png");
      if (fs.existsSync(fallback)) resolved = fallback;
    }
    if (fs.existsSync(resolved)) {
      const x = doc.page.margins.left + 8;
      const y = doc.page.margins.top - 34;
      doc.image(resolved, x, y, { width: 90 });
    }
  } catch {
    // ignore logo failures
  }
}

/**
 * Section title
 */
export function sectionTitle(doc, text) {
  ensurePageSpace(doc, 40);

  doc.moveDown(0.4);
  doc.font("Helvetica-Bold").fontSize(12).fillColor(BRAND.primary);
  doc.text(text, doc.page.margins.left, doc.y, { width: contentWidth(doc) });

  doc.moveDown(0.2);
  const y = doc.y;
  doc
    .moveTo(doc.page.margins.left, y)
    .lineTo(doc.page.width - doc.page.margins.right, y)
    .lineWidth(1)
    .strokeOpacity(1)
    .stroke(BRAND.line);

  doc.moveDown(0.6);
  doc.font("Helvetica").fillColor(BRAND.text);
}

/**
 * Key/value row (single line)
 */
export function kvLine(doc, k, v) {
  const left = doc.page.margins.left;
  const w = contentWidth(doc);

  doc.font("Helvetica").fontSize(10).fillColor(BRAND.text);
  doc.text(`${k}: `, left, doc.y, { continued: true, width: w });
  doc.font("Helvetica").text(String(v ?? ""), { width: w });
}

/**
 * KPI block: neat 2-column layout for KPIs
 * items = [{k,v}, ...]
 */
export function kvGrid(doc, items, cols = 2) {
  ensurePageSpace(doc, 70);

  const left = doc.page.margins.left;
  const w = contentWidth(doc);
  const colW = w / cols;
  const rowH = 14;

  doc.font("Helvetica").fontSize(10).fillColor(BRAND.text);

  let x = left;
  let y = doc.y;

  items.forEach((it, idx) => {
    const cx = left + (idx % cols) * colW;
    const cy = y + Math.floor(idx / cols) * rowH;

    doc.font("Helvetica-Bold").text(`${it.k}: `, cx, cy, { continued: true, width: colW });
    doc.font("Helvetica").text(String(it.v ?? ""), { width: colW });
  });

  const rows = Math.ceil(items.length / cols);
  doc.y = y + rows * rowH + 6;
}

/**
 * Proper table with column widths, header shading, and auto page breaks.
 *
 * columns = [
 *   { key: "asset", label: "Asset", width: 0.18, align: "left" },
 *   { key: "hours", label: "Hours", width: 0.12, align: "right" },
 * ]
 *
 * rows = [{asset:"A300AM", hours: 9.5}, ...]
 */
export function table(doc, columns, rows, opts = {}) {
  const left = doc.page.margins.left;
  const right = doc.page.width - doc.page.margins.right;
  const w = right - left;

  const fontSize = opts.fontSize ?? 9;
  const headerFontSize = opts.headerFontSize ?? 9;
  const rowPadY = opts.rowPadY ?? 4;
  const headerPadY = opts.headerPadY ?? 5;

  // compute absolute column widths
  const colAbs = columns.map((c) => ({
    ...c,
    absW: Math.floor((c.width ?? 0) * w),
    align: c.align || "left",
  }));

  // last column: take remainder so we perfectly fill width
  const used = colAbs.slice(0, -1).reduce((s, c) => s + c.absW, 0);
  colAbs[colAbs.length - 1].absW = Math.max(60, w - used);

  const headerH = headerFontSize + headerPadY * 2;
  const rowH = fontSize + rowPadY * 2;

  // header draw helper
  const drawHeader = () => {
    ensurePageSpace(doc, headerH + 10);

    const y = doc.y;

    // shaded header
    doc.save();
    doc
      .rect(left, y, w, headerH)
      .fillOpacity(1)
      .fill(BRAND.primary);
    doc.restore();

    doc.font("Helvetica-Bold").fontSize(headerFontSize).fillColor("#ffffff");

    let x = left;
    for (const c of colAbs) {
      doc.text(c.label, x + 4, y + headerPadY, { width: c.absW - 8, align: c.align });
      x += c.absW;
    }

    // bottom line
    doc
      .moveTo(left, y + headerH)
      .lineTo(right, y + headerH)
      .lineWidth(1)
      .strokeOpacity(1)
      .stroke(BRAND.line);

    doc.y = y + headerH + 2;
    doc.font("Helvetica").fontSize(fontSize).fillColor(BRAND.text);
  };

  drawHeader();

  const getRowHeight = (row) => {
    let maxTextH = fontSize;
    for (const c of colAbs) {
      const val = row?.[c.key];
      const txt = String(val ?? "");
      const h = doc.heightOfString(txt, { width: c.absW - 8, align: c.align });
      if (h > maxTextH) maxTextH = h;
    }
    return Math.max(rowH, Math.ceil(maxTextH + rowPadY * 2));
  };

  // rows
  for (let i = 0; i < rows.length; i++) {
    const rowDrawH = getRowHeight(rows[i]);
    ensurePageSpace(doc, rowDrawH + 12);

    // if we're near bottom, new page and redraw header
    const bottomY = doc.page.height - doc.page.margins.bottom;
    if (doc.y + rowDrawH > bottomY) {
      doc.addPage();
      drawHeader();
    }

    const y = doc.y;

    // zebra stripe
    if (i % 2 === 1) {
      doc.save();
      doc
        .rect(left, y, w, rowDrawH)
        .fillOpacity(1)
        .fill(BRAND.zebra);
      doc.restore();
    }

    let x = left;
    for (const c of colAbs) {
      const val = rows[i]?.[c.key];
      doc.text(String(val ?? ""), x + 4, y + rowPadY, { width: c.absW - 8, align: c.align });
      x += c.absW;
    }

    // row divider
    doc
      .moveTo(left, y + rowDrawH)
      .lineTo(right, y + rowDrawH)
      .lineWidth(1)
      .strokeOpacity(1)
      .stroke("#e2e8f0");

    doc.y = y + rowDrawH;
  }

  doc.moveDown(0.8);
}

/**
 * Compatibility helper (kept for older report code)
 * Headers/rows still supported but now respect margins and wrap.
 */
export function smallTable(doc, headers, rows) {
  const left = doc.page.margins.left;
  const w = contentWidth(doc);

  doc.font("Helvetica").fontSize(9).fillColor(BRAND.text);

  doc.text(headers.join(" | "), left, doc.y, { width: w });

  doc.moveDown(0.2);
  const y = doc.y;
  doc
    .moveTo(left, y)
    .lineTo(left + w, y)
    .lineWidth(1)
    .strokeOpacity(0.18)
    .stroke("#000000");

  doc.moveDown(0.4);

  for (const r of rows) {
    const line = r.map((x) => String(x ?? "")).join(" | ");
    doc.text(line, left, doc.y, { width: w });
  }

  doc.moveDown(0.8);
}