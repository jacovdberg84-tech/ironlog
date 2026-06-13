import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import sharp from "sharp";
import { ensurePageSpace } from "./pdfGenerator.js";

const PDF_NATIVE_EXTS = new Set([".jpg", ".jpeg", ".png"]);

/** Normalize camera uploads to JPEG so PDFKit can embed them reliably. */
export async function normalizeUploadedPhoto(buffer) {
  return sharp(buffer, { failOn: "none" })
    .rotate()
    .jpeg({ quality: 85, mozjpeg: true })
    .toBuffer();
}

/**
 * Resolve an on-disk image to a PDFKit-compatible file path.
 * Converts WebP/HEIC and mislabeled formats via sharp when needed.
 */
export async function resolvePdfImage(absPath) {
  const abs = String(absPath || "").trim();
  if (!abs || !fs.existsSync(abs)) return { path: null, temp: false };

  const ext = path.extname(abs).toLowerCase();
  if (PDF_NATIVE_EXTS.has(ext)) {
    try {
      await sharp(abs).metadata();
      return { path: abs, temp: false };
    } catch {
      /* re-encode below */
    }
  }

  const tmp = path.join(
    os.tmpdir(),
    `ironlog-pdf-${Date.now()}-${Math.random().toString(36).slice(2, 10)}.jpg`
  );
  await sharp(abs, { failOn: "none" }).rotate().jpeg({ quality: 85 }).toFile(tmp);
  return { path: tmp, temp: true };
}

export function drawPhotoInPdf(doc, absPath, opts = {}) {
  const maxW = Number(opts.maxWidth || 420);
  const maxH = Number(opts.maxHeight || 190);
  const missingLabel = String(opts.missingLabel || "-");

  ensurePageSpace(doc, maxH + 40);
  const left = doc.page.margins.left;
  const y = doc.y;

  if (!absPath || !fs.existsSync(absPath)) {
    doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${missingLabel}`);
    doc.moveDown(0.5);
    return false;
  }

  try {
    const dims = doc.image(absPath, left, y, { fit: [maxW, maxH], align: "left", valign: "top" });
    doc.y = y + (dims?.height || maxH) + 8;
    return true;
  } catch {
    doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo file exists but could not be rendered.");
    doc.moveDown(0.5);
    return false;
  }
}

export function cleanupTempPdfImages(paths) {
  for (const p of paths || []) {
    if (!p) continue;
    try {
      fs.unlinkSync(p);
    } catch {
      /* ignore */
    }
  }
}
