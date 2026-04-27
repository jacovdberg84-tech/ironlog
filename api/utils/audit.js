export function ensureAuditTable(db) {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS audit_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      module TEXT NOT NULL,
      action TEXT NOT NULL,
      entity_type TEXT,
      entity_id TEXT,
      username TEXT,
      role TEXT,
      site_code TEXT,
      source_app TEXT,
      source_channel TEXT,
      request_id TEXT,
      ip_address TEXT,
      user_agent TEXT,
      before_json TEXT,
      after_json TEXT,
      payload_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  const cols = new Set(db.prepare(`PRAGMA table_info(audit_logs)`).all().map((r) => String(r.name || "")));
  const ensureCol = (name, def) => {
    if (!cols.has(name)) db.prepare(`ALTER TABLE audit_logs ADD COLUMN ${def}`).run();
  };
  ensureCol("site_code", "site_code TEXT");
  ensureCol("source_app", "source_app TEXT");
  ensureCol("source_channel", "source_channel TEXT");
  ensureCol("request_id", "request_id TEXT");
  ensureCol("ip_address", "ip_address TEXT");
  ensureCol("user_agent", "user_agent TEXT");
  ensureCol("before_json", "before_json TEXT");
  ensureCol("after_json", "after_json TEXT");

  db.prepare(`
    CREATE INDEX IF NOT EXISTS idx_audit_logs_created
    ON audit_logs(created_at DESC)
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_audit_logs_actor ON audit_logs(username, created_at DESC)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_audit_logs_site ON audit_logs(site_code, created_at DESC)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_audit_logs_entity ON audit_logs(entity_type, entity_id, created_at DESC)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_audit_logs_action ON audit_logs(action, created_at DESC)`).run();
}

export function writeAudit(db, req, entry) {
  const module = String(entry?.module || "").trim() || "system";
  const action = String(entry?.action || "").trim() || "unknown";
  const entity_type = entry?.entity_type != null ? String(entry.entity_type) : null;
  const entity_id = entry?.entity_id != null ? String(entry.entity_id) : null;
  const username = String(req?.headers?.["x-user-name"] || "").trim() || "session-user";
  const role = String(req?.headers?.["x-user-role"] || "").trim().toLowerCase() || "unknown";
  const site_code = String(req?.headers?.["x-site-code"] || "").trim().toLowerCase() || null;
  const source_app = String(entry?.source_app || req?.headers?.["x-source-app"] || req?.headers?.["x-client-app"] || "web").trim().toLowerCase();
  const source_channel = String(entry?.source_channel || req?.headers?.["x-source-channel"] || "api").trim().toLowerCase();
  const request_id = String(req?.id || req?.headers?.["x-request-id"] || "").trim() || null;
  const ip_address = String(req?.ip || req?.socket?.remoteAddress || "").trim() || null;
  const user_agent = String(req?.headers?.["user-agent"] || "").trim() || null;

  let payload_json = null;
  if (entry?.payload != null) {
    try {
      payload_json = JSON.stringify(entry.payload);
    } catch {
      payload_json = String(entry.payload);
    }
  }
  let before_json = null;
  if (entry?.before != null) {
    try { before_json = JSON.stringify(entry.before); } catch { before_json = String(entry.before); }
  }
  let after_json = null;
  if (entry?.after != null) {
    try { after_json = JSON.stringify(entry.after); } catch { after_json = String(entry.after); }
  }

  db.prepare(`
    INSERT INTO audit_logs (
      module, action, entity_type, entity_id, username, role, site_code, source_app, source_channel,
      request_id, ip_address, user_agent, before_json, after_json, payload_json
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(
    module, action, entity_type, entity_id, username, role, site_code, source_app, source_channel,
    request_id, ip_address, user_agent, before_json, after_json, payload_json
  );
}
