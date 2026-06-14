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

export function getPdfReportSite(db) {
  ensureReportSettingsSchema(db);
  const site_code = getReportSetting(db, "pdf_site_code");
  let site_name = getReportSetting(db, "pdf_site_name");
  if (!site_name && site_code) {
    try {
      const row = db.prepare(`
        SELECT site_name
        FROM site_profiles
        WHERE site_code = ? AND COALESCE(active, 1) = 1
        LIMIT 1
      `).get(site_code);
      if (row?.site_name) site_name = String(row.site_name).trim();
    } catch {
      // site_profiles may not exist on older DBs
    }
    if (!site_name) site_name = site_code;
  }
  return { site_code, site_name };
}

export function savePdfReportSite(db, { site_code, site_name }) {
  const code = String(site_code || "").trim();
  const name = String(site_name || "").trim();
  setReportSetting(db, "pdf_site_code", code);
  setReportSetting(db, "pdf_site_name", name);
  return getPdfReportSite(db);
}
