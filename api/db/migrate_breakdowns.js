// One-off migration for breakdown incident model
import { db } from "./client.js";

console.log("Running breakdown migration...");

db.exec(`
PRAGMA foreign_keys = ON;

ALTER TABLE breakdowns ADD COLUMN status TEXT NOT NULL DEFAULT 'OPEN';
ALTER TABLE breakdowns ADD COLUMN start_at TEXT;
ALTER TABLE breakdowns ADD COLUMN end_at TEXT;
ALTER TABLE breakdowns ADD COLUMN component TEXT;
ALTER TABLE breakdowns ADD COLUMN downtime_total_hours REAL NOT NULL DEFAULT 0;
ALTER TABLE breakdowns ADD COLUMN primary_work_order_id INTEGER;

UPDATE breakdowns
SET downtime_total_hours = COALESCE(downtime_hours, 0)
WHERE COALESCE(downtime_total_hours, 0) = 0
  AND COALESCE(downtime_hours, 0) > 0;

CREATE TABLE IF NOT EXISTS breakdown_downtime_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  breakdown_id INTEGER NOT NULL,
  log_date TEXT NOT NULL,
  hours_down REAL NOT NULL DEFAULT 0,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (breakdown_id) REFERENCES breakdowns(id) ON DELETE CASCADE,
  UNIQUE (breakdown_id, log_date)
);

CREATE TABLE IF NOT EXISTS breakdown_components (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  breakdown_id INTEGER NOT NULL,
  component TEXT NOT NULL,
  symptom TEXT,
  work_order_id INTEGER,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (breakdown_id) REFERENCES breakdowns(id) ON DELETE CASCADE,
  FOREIGN KEY (work_order_id) REFERENCES work_orders(id) ON DELETE SET NULL
);

CREATE TRIGGER IF NOT EXISTS trg_bd_logs_updated_at
AFTER UPDATE ON breakdown_downtime_logs
FOR EACH ROW
BEGIN
  UPDATE breakdown_downtime_logs
  SET updated_at = datetime('now')
  WHERE id = NEW.id;
END;

CREATE TRIGGER IF NOT EXISTS trg_bd_logs_recalc_total_ai
AFTER INSERT ON breakdown_downtime_logs
FOR EACH ROW
BEGIN
  UPDATE breakdowns
  SET downtime_total_hours = (
    SELECT COALESCE(SUM(hours_down), 0)
    FROM breakdown_downtime_logs
    WHERE breakdown_id = NEW.breakdown_id
  )
  WHERE id = NEW.breakdown_id;
END;

CREATE TRIGGER IF NOT EXISTS trg_bd_logs_recalc_total_au
AFTER UPDATE ON breakdown_downtime_logs
FOR EACH ROW
BEGIN
  UPDATE breakdowns
  SET downtime_total_hours = (
    SELECT COALESCE(SUM(hours_down), 0)
    FROM breakdown_downtime_logs
    WHERE breakdown_id = NEW.breakdown_id
  )
  WHERE id = NEW.breakdown_id;
END;

CREATE TRIGGER IF NOT EXISTS trg_bd_logs_recalc_total_ad
AFTER DELETE ON breakdown_downtime_logs
FOR EACH ROW
BEGIN
  UPDATE breakdowns
  SET downtime_total_hours = (
    SELECT COALESCE(SUM(hours_down), 0)
    FROM breakdown_downtime_logs
    WHERE breakdown_id = OLD.breakdown_id
  )
  WHERE id = OLD.breakdown_id;
END;
`);

console.log("Breakdown migration complete ✅");
process.exit(0);