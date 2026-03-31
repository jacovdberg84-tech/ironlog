// C:\IRONLOG\api\db\upgrade.js
import { db } from "./client.js";

function hasColumn(table, col) {
  const rows = db.prepare(`PRAGMA table_info(${table})`).all();
  return rows.some(r => r.name === col);
}

function addColumn(table, colDef) {
  db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
}
// Add to your upgrade.js steps
db.exec(`
  ALTER TABLE daily_hours ADD COLUMN scheduled_hours REAL;
`);
db.exec(`
  ALTER TABLE daily_hours ADD COLUMN opening_hours REAL;
`);
db.exec(`
  ALTER TABLE daily_hours ADD COLUMN closing_hours REAL;
`);

function ensureColumn(table, colName, colDef) {
  if (!hasColumn(table, colName)) {
    console.log(`+ Adding column ${table}.${colName}`);
    addColumn(table, colDef);
  } else {
    console.log(`= Column exists ${table}.${colName}`);
  }
}

function main() {
  console.log("IRONLOG DB UPGRADE starting...");

  // daily_hours new columns
  ensureColumn("daily_hours", "scheduled_hours", "scheduled_hours REAL NOT NULL DEFAULT 0");
  ensureColumn("daily_hours", "opening_hours", "opening_hours REAL");
  ensureColumn("daily_hours", "closing_hours", "closing_hours REAL");
  ensureColumn("daily_hours", "hours_run", "hours_run REAL NOT NULL DEFAULT 0");

  console.log("IRONLOG DB UPGRADE done ✅");
}

main();