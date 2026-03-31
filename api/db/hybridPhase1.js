// IRONLOG/api/db/hybridPhase1.js
import crypto from "node:crypto";
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

function ensureColumn(table, colName, colDef) {
  if (!hasTable(table)) return;
  if (!hasColumn(table, colName)) {
    db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
  }
}

function ensureUuidIndex(table) {
  if (!hasTable(table) || !hasColumn(table, "uuid")) return;
  db.exec(`CREATE UNIQUE INDEX IF NOT EXISTS idx_${table}_uuid ON ${table}(uuid)`);
}

function ensureUpdatedAtTrigger(table) {
  if (!hasTable(table) || !hasColumn(table, "updated_at")) return;
  const trig = `trg_${table}_updated_at`;
  db.exec(`
    CREATE TRIGGER IF NOT EXISTS ${trig}
    AFTER UPDATE ON ${table}
    FOR EACH ROW
    BEGIN
      UPDATE ${table}
      SET updated_at = datetime('now')
      WHERE id = NEW.id;
    END;
  `);
}

function backfillPhase1(table) {
  if (!hasTable(table)) return;
  if (hasColumn(table, "site_code")) {
    db.prepare(`UPDATE ${table} SET site_code = 'main' WHERE site_code IS NULL OR TRIM(site_code) = ''`).run();
  }
  if (hasColumn(table, "updated_at")) {
    const createdCol = hasColumn(table, "created_at");
    if (createdCol) {
      db.prepare(`UPDATE ${table} SET updated_at = created_at WHERE updated_at IS NULL OR TRIM(updated_at) = ''`).run();
    }
    db.prepare(`UPDATE ${table} SET updated_at = datetime('now') WHERE updated_at IS NULL OR TRIM(updated_at) = ''`).run();
  }
  if (hasColumn(table, "uuid")) {
    const rows = db.prepare(`SELECT id FROM ${table} WHERE uuid IS NULL OR TRIM(uuid) = ''`).all();
    const upd = db.prepare(`UPDATE ${table} SET uuid = ? WHERE id = ?`);
    const tx = db.transaction((items) => {
      for (const r of items) {
        upd.run(crypto.randomUUID(), r.id);
      }
    });
    tx(rows);
  }
}

export function runHybridPhase1Migration() {
  const tables = [
    "daily_hours",
    "fuel_logs",
    "oil_logs",
    "breakdowns",
    "breakdown_downtime_logs",
    "work_orders",
    "manager_inspections",
    "manager_inspection_photos",
  ];

  for (const t of tables) {
    if (!hasTable(t)) continue;
    ensureColumn(t, "uuid", "uuid TEXT");
    ensureColumn(t, "site_code", "site_code TEXT DEFAULT 'main'");
    ensureColumn(t, "updated_at", "updated_at TEXT");
    ensureUuidIndex(t);
    ensureUpdatedAtTrigger(t);
    backfillPhase1(t);
  }
}

