import fs from "node:fs";
import path from "node:path";
import { getDataRoot, normalizeStorageRel, resolveStorageAbs } from "./storagePaths.js";

const PDF_LOGO_SETTING_KEY = "pdf_company_logo_path";
export const PDF_REPORT_LOGO_REL_DIR = "uploads/report-branding";

export function ensureReportSettingsSchema(db) {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS report_settings (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
}

export function getReportSetting(db, key, fallback = "") {
  ensureReportSettingsSchema(db);
  const row = db.prepare(`SELECT value FROM report_settings WHERE key = ?`).get(String(key || ""));
  const val = row?.value;
  return val == null ? String(fallback || "") : String(val).trim();
}

export function setReportSetting(db, key, value) {
  ensureReportSettingsSchema(db);
  db.prepare(`
    INSERT INTO report_settings (key, value, updated_at)
    VALUES (?, ?, datetime('now'))
    ON CONFLICT(key) DO UPDATE SET
      value = excluded.value,
      updated_at = datetime('now')
  `).run(String(key || ""), String(value ?? "").trim());
}

function lookupCompanyName(db, company_code) {
  if (!company_code) return "";
  try {
    const row = db.prepare(`
      SELECT company_name
      FROM company_profiles
      WHERE company_code = ? AND COALESCE(active, 1) = 1
      LIMIT 1
    `).get(company_code);
    return row?.company_name ? String(row.company_name).trim() : "";
  } catch {
    return "";
  }
}

function lookupSiteName(db, site_code) {
  if (!site_code) return "";
  try {
    const row = db.prepare(`
      SELECT site_name
      FROM site_profiles
      WHERE site_code = ? AND COALESCE(active, 1) = 1
      LIMIT 1
    `).get(site_code);
    return row?.site_name ? String(row.site_name).trim() : "";
  } catch {
    return "";
  }
}

export function getPdfBrandingLogoDir(dataRoot = getDataRoot()) {
  return path.join(dataRoot, PDF_REPORT_LOGO_REL_DIR);
}

function legacyPdfLogoCandidates(dataRoot = getDataRoot()) {
  return [
    path.join(process.cwd(), "branding", "logo.png"),
    path.join(dataRoot, "branding", "logo.png"),
    path.join(process.cwd(), "..", "branding", "logo.png"),
  ];
}

/** Absolute path to the logo used on PDF headers, or null. */
export function resolvePdfCompanyLogoAbs(db, dataRoot = getDataRoot()) {
  const rel = getReportSetting(db, PDF_LOGO_SETTING_KEY);
  if (rel) {
    const abs = resolveStorageAbs(rel, dataRoot);
    if (abs && fs.existsSync(abs)) return abs;
  }
  for (const candidate of legacyPdfLogoCandidates(dataRoot)) {
    if (fs.existsSync(candidate)) return candidate;
  }
  return null;
}

export function setPdfCompanyLogoPath(db, relPath) {
  setReportSetting(db, PDF_LOGO_SETTING_KEY, String(relPath || "").trim());
}

export function clearPdfCompanyLogo(db, dataRoot = getDataRoot()) {
  const rel = getReportSetting(db, PDF_LOGO_SETTING_KEY);
  if (rel) {
    const abs = resolveStorageAbs(rel, dataRoot);
    if (abs && fs.existsSync(abs)) {
      try {
        fs.unlinkSync(abs);
      } catch {
        // ignore
      }
    }
  }
  setReportSetting(db, PDF_LOGO_SETTING_KEY, "");
}

export function getPdfReportBranding(db) {
  ensureReportSettingsSchema(db);
  const company_code = getReportSetting(db, "pdf_company_code");
  const site_code = getReportSetting(db, "pdf_site_code");
  let company_name = getReportSetting(db, "pdf_company_name");
  let site_name = getReportSetting(db, "pdf_site_name");
  if (!company_name && company_code) {
    company_name = lookupCompanyName(db, company_code) || company_code;
  }
  if (!site_name && site_code) {
    site_name = lookupSiteName(db, site_code) || site_code;
  }
  const company_logo_path = getReportSetting(db, PDF_LOGO_SETTING_KEY);
  const company_logo_abs = resolvePdfCompanyLogoAbs(db);
  const company_logo_custom = Boolean(company_logo_path);
  const company_logo_available = Boolean(company_logo_abs);
  const company_logo_url = company_logo_available ? "/api/reports/pdf-settings/logo" : "";
  return {
    company_code,
    company_name,
    site_code,
    site_name,
    company_logo_path,
    company_logo_custom,
    company_logo_available,
    company_logo_url,
  };
}

/** @deprecated use getPdfReportBranding */
export function getPdfReportSite(db) {
  const b = getPdfReportBranding(db);
  return { site_code: b.site_code, site_name: b.site_name };
}

export function savePdfReportBranding(db, payload = {}) {
  if (payload.company_code !== undefined) {
    setReportSetting(db, "pdf_company_code", String(payload.company_code || "").trim());
  }
  if (payload.company_name !== undefined) {
    setReportSetting(db, "pdf_company_name", String(payload.company_name || "").trim());
  }
  if (payload.site_code !== undefined) {
    setReportSetting(db, "pdf_site_code", String(payload.site_code || "").trim());
  }
  if (payload.site_name !== undefined) {
    setReportSetting(db, "pdf_site_name", String(payload.site_name || "").trim());
  }
  return getPdfReportBranding(db);
}

/** @deprecated use savePdfReportBranding */
export function savePdfReportSite(db, { site_code, site_name }) {
  return savePdfReportBranding(db, { site_code, site_name });
}
