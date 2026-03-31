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
      payload_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE INDEX IF NOT EXISTS idx_audit_logs_created
    ON audit_logs(created_at DESC)
  `).run();
}

export function writeAudit(db, req, entry) {
  const module = String(entry?.module || "").trim() || "system";
  const action = String(entry?.action || "").trim() || "unknown";
  const entity_type = entry?.entity_type != null ? String(entry.entity_type) : null;
  const entity_id = entry?.entity_id != null ? String(entry.entity_id) : null;
  const username = String(req?.headers?.["x-user-name"] || "").trim() || "session-user";
  const role = String(req?.headers?.["x-user-role"] || "").trim().toLowerCase() || "unknown";

  let payload_json = null;
  if (entry?.payload != null) {
    try {
      payload_json = JSON.stringify(entry.payload);
    } catch {
      payload_json = String(entry.payload);
    }
  }

  db.prepare(`
    INSERT INTO audit_logs (
      module, action, entity_type, entity_id, username, role, payload_json
    )
    VALUES (?, ?, ?, ?, ?, ?, ?)
  `).run(module, action, entity_type, entity_id, username, role, payload_json);
}
