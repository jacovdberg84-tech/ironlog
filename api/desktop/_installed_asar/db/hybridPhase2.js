// IRONLOG/api/db/hybridPhase2.js
import { db } from "./client.js";

function hasTable(table) {
  const row = db
    .prepare(`
      SELECT name
      FROM sqlite_master
      WHERE type = 'table' AND name = ?
    `)
    .get(table);
  return Boolean(row);
}

function hasColumn(table, col) {
  if (!hasTable(table)) return false;
  const rows = db.prepare(`PRAGMA table_info(${table})`).all();
  return rows.some((r) => String(r.name || "") === String(col));
}

function ensureSyncOutbox() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS sync_outbox (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      table_name TEXT NOT NULL,
      row_id INTEGER,
      row_uuid TEXT,
      op TEXT NOT NULL,                    -- upsert | delete
      site_code TEXT NOT NULL DEFAULT 'main',
      payload_json TEXT,
      changed_at TEXT NOT NULL DEFAULT (datetime('now')),
      attempts INTEGER NOT NULL DEFAULT 0,
      last_attempt_at TEXT,
      synced_at TEXT,
      error_text TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_sync_outbox_unsynced ON sync_outbox(synced_at, id);
    CREATE INDEX IF NOT EXISTS idx_sync_outbox_table_uuid ON sync_outbox(table_name, row_uuid);
  `);
}

function ensureSyncCheckpoints() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS sync_checkpoints (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      peer_name TEXT NOT NULL,
      table_name TEXT NOT NULL,
      last_outbox_id INTEGER NOT NULL DEFAULT 0,
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(peer_name, table_name)
    );
  `);
}

function ensureSyncState() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS sync_state (
      key TEXT PRIMARY KEY,
      value_json TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
  `);
  db.prepare(`
    INSERT INTO sync_state (key, value_json, updated_at)
    VALUES ('schema_version', '{"value":1}', datetime('now'))
    ON CONFLICT(key) DO NOTHING
  `).run();
}

function ensureSyncAppliedEvents() {
  db.exec(`
    CREATE TABLE IF NOT EXISTS sync_applied_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      peer_name TEXT NOT NULL,
      event_key TEXT NOT NULL,
      event_id INTEGER,
      table_name TEXT,
      row_uuid TEXT,
      op TEXT,
      applied_at TEXT NOT NULL DEFAULT (datetime('now')),
      result_json TEXT,
      UNIQUE(peer_name, event_key)
    );
    CREATE INDEX IF NOT EXISTS idx_sync_applied_peer_event_id ON sync_applied_events(peer_name, event_id);
  `);
}

function ensureOutboxTriggersForTable(table) {
  if (!hasTable(table)) return;
  if (!hasColumn(table, "id")) return;
  if (!hasColumn(table, "uuid")) return;
  if (!hasColumn(table, "site_code")) return;
  if (!hasColumn(table, "updated_at")) return;

  const qTable = table.replace(/"/g, "\"\"");
  const trigIns = `trg_${table}_sync_ai`;
  const trigUpd = `trg_${table}_sync_au`;
  const trigDel = `trg_${table}_sync_ad`;

  db.exec(`
    CREATE TRIGGER IF NOT EXISTS ${trigIns}
    AFTER INSERT ON ${qTable}
    FOR EACH ROW
    BEGIN
      UPDATE ${qTable}
      SET
        uuid = COALESCE(NULLIF(uuid, ''), lower(hex(randomblob(16)))),
        site_code = COALESCE(NULLIF(site_code, ''), 'main'),
        updated_at = COALESCE(NULLIF(updated_at, ''), datetime('now'))
      WHERE id = NEW.id;

      INSERT INTO sync_outbox (table_name, row_id, row_uuid, op, site_code, payload_json, changed_at)
      SELECT '${qTable}', id, uuid, 'upsert', COALESCE(site_code, 'main'), NULL, datetime('now')
      FROM ${qTable}
      WHERE id = NEW.id;
    END;
  `);

  db.exec(`
    CREATE TRIGGER IF NOT EXISTS ${trigUpd}
    AFTER UPDATE ON ${qTable}
    FOR EACH ROW
    BEGIN
      UPDATE ${qTable}
      SET
        uuid = COALESCE(NULLIF(uuid, ''), lower(hex(randomblob(16)))),
        site_code = COALESCE(NULLIF(site_code, ''), 'main'),
        updated_at = COALESCE(NULLIF(updated_at, ''), datetime('now'))
      WHERE id = NEW.id;

      INSERT INTO sync_outbox (table_name, row_id, row_uuid, op, site_code, payload_json, changed_at)
      SELECT '${qTable}', id, uuid, 'upsert', COALESCE(site_code, 'main'), NULL, datetime('now')
      FROM ${qTable}
      WHERE id = NEW.id;
    END;
  `);

  db.exec(`
    CREATE TRIGGER IF NOT EXISTS ${trigDel}
    AFTER DELETE ON ${qTable}
    FOR EACH ROW
    BEGIN
      INSERT INTO sync_outbox (table_name, row_id, row_uuid, op, site_code, payload_json, changed_at)
      VALUES ('${qTable}', OLD.id, COALESCE(NULLIF(OLD.uuid, ''), lower(hex(randomblob(16)))), 'delete', COALESCE(NULLIF(OLD.site_code, ''), 'main'), NULL, datetime('now'));
    END;
  `);
}

export function runHybridPhase2Migration() {
  ensureSyncOutbox();
  ensureSyncCheckpoints();
  ensureSyncState();
  ensureSyncAppliedEvents();

  const tracked = [
    "daily_hours",
    "fuel_logs",
    "oil_logs",
    "breakdowns",
    "breakdown_downtime_logs",
    "work_orders",
    "manager_inspections",
    "manager_inspection_photos",
  ];
  for (const t of tracked) ensureOutboxTriggersForTable(t);
}

