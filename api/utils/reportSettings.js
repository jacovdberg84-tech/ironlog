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
  return { company_code, company_name, site_code, site_name };
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
