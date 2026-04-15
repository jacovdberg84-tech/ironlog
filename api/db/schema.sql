PRAGMA foreign_keys = ON;

/* =========================
   ASSETS
========================= */
CREATE TABLE IF NOT EXISTS assets (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_code TEXT NOT NULL UNIQUE,
  asset_name TEXT NOT NULL,
  category TEXT,
  active INTEGER NOT NULL DEFAULT 1,
  is_standby INTEGER NOT NULL DEFAULT 0,

  archived INTEGER NOT NULL DEFAULT 0,
  archive_reason TEXT,
  archived_at TEXT,

  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

/* =========================
   DAILY HOURS (CORE LOGIC)
========================= */
CREATE TABLE IF NOT EXISTS daily_hours (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id INTEGER NOT NULL,
  work_date TEXT NOT NULL,                -- YYYY-MM-DD

  scheduled_hours REAL NOT NULL DEFAULT 0,
  opening_hours REAL,
  closing_hours REAL,

  hours_run REAL NOT NULL DEFAULT 0,
  is_used INTEGER NOT NULL DEFAULT 1,     -- 1=production, 0=standby

  operator TEXT,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),

  FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT,
  UNIQUE (asset_id, work_date)
);

/* =========================
   FUEL LOGS
========================= */
CREATE TABLE IF NOT EXISTS fuel_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id INTEGER NOT NULL,
  log_date TEXT NOT NULL,
  liters REAL NOT NULL,
  source TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
);

/* =========================
   OIL LOGS
========================= */
CREATE TABLE IF NOT EXISTS oil_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id INTEGER NOT NULL,
  log_date TEXT NOT NULL,
  oil_type TEXT,
  quantity REAL NOT NULL,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
);

/* =========================
   WORK ORDERS
   NOTE: reference_id stays generic by design (breakdown.id or maintenance_plan.id)
========================= */
CREATE TABLE IF NOT EXISTS work_orders (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id INTEGER NOT NULL,
  source TEXT NOT NULL,            -- breakdown | service | manual
  reference_id INTEGER,            -- breakdown.id or maintenance_plan.id
  status TEXT NOT NULL DEFAULT 'open',
  opened_at TEXT NOT NULL DEFAULT (datetime('now')),
  closed_at TEXT,
  FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
);

/* =========================
   BREAKDOWNS (INCIDENT MODEL)
   One breakdown = one incident (OPEN/CLOSED)
   Downtime is logged in breakdown_downtime_logs and summed into downtime_total_hours.
========================= */
CREATE TABLE IF NOT EXISTS breakdowns (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id INTEGER NOT NULL,

  -- Original date kept for compatibility (used as "reported date")
  breakdown_date TEXT NOT NULL,            -- YYYY-MM-DD

  -- Incident lifecycle
  status TEXT NOT NULL DEFAULT 'OPEN',     -- OPEN | CLOSED
  start_at TEXT,                          -- ISO timestamp (optional)
  end_at TEXT,                            -- ISO timestamp (optional)

  -- Description / classification
  description TEXT NOT NULL,
  component TEXT,                         -- optional high-level component: engine/hydraulics/electrical/etc
  critical INTEGER NOT NULL DEFAULT 0,

  -- Total downtime for this incident (auto-updated by triggers from downtime logs)
  downtime_total_hours REAL NOT NULL DEFAULT 0,

  -- Primary WO link (ONE per incident by default)
  primary_work_order_id INTEGER,

  created_at TEXT NOT NULL DEFAULT (datetime('now')),

  FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT,
  FOREIGN KEY (primary_work_order_id) REFERENCES work_orders(id) ON DELETE SET NULL
);

/* =========================
   BREAKDOWN DOWNTIME LOGS (THE FIX)
   Log downtime per day (or per shift if you prefer one entry per date).
   UNIQUE (breakdown_id, log_date) prevents duplicates; you edit instead.
========================= */
CREATE TABLE IF NOT EXISTS breakdown_downtime_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  breakdown_id INTEGER NOT NULL,
  log_date TEXT NOT NULL,                  -- YYYY-MM-DD
  hours_down REAL NOT NULL DEFAULT 0,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),

  FOREIGN KEY (breakdown_id) REFERENCES breakdowns(id) ON DELETE CASCADE,
  UNIQUE (breakdown_id, log_date)
);

/* =========================
   OPTIONAL: BREAKDOWN COMPONENT LINES
   Lets you open another WO ONLY when it’s a different component/fault stream.
========================= */
CREATE TABLE IF NOT EXISTS breakdown_components (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  breakdown_id INTEGER NOT NULL,
  component TEXT NOT NULL,                 -- engine/hydraulics/electrical/etc
  symptom TEXT,
  work_order_id INTEGER,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),

  FOREIGN KEY (breakdown_id) REFERENCES breakdowns(id) ON DELETE CASCADE,
  FOREIGN KEY (work_order_id) REFERENCES work_orders(id) ON DELETE SET NULL
);

/* =========================
   MAINTENANCE PLANS
========================= */
CREATE TABLE IF NOT EXISTS maintenance_plans (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id INTEGER NOT NULL,
  service_name TEXT NOT NULL,
  interval_hours REAL NOT NULL,
  last_service_hours REAL NOT NULL DEFAULT 0,
  active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
);

/* =========================
   PARTS MASTER
========================= */
CREATE TABLE IF NOT EXISTS parts (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  part_code TEXT NOT NULL UNIQUE,
  part_name TEXT NOT NULL,
  min_stock REAL NOT NULL DEFAULT 0,
  critical INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

/* =========================
   STOCK MOVEMENTS
========================= */
CREATE TABLE IF NOT EXISTS stock_movements (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  part_id INTEGER NOT NULL,
  quantity REAL NOT NULL,          -- +in / -out
  movement_type TEXT NOT NULL,     -- in | out | adjust
  reference TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE RESTRICT
);

/* =========================
   PARTS ORDERS
========================= */
CREATE TABLE IF NOT EXISTS parts_orders (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  part_id INTEGER NOT NULL,
  quantity REAL NOT NULL,
  expected_date TEXT,
  status TEXT NOT NULL DEFAULT 'ordered',   -- ordered | received
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE RESTRICT
);

/* =========================
   TRIGGERS: Keep downtime_total_hours accurate
========================= */

-- Update updated_at on downtime logs
CREATE TRIGGER IF NOT EXISTS trg_bd_logs_updated_at
AFTER UPDATE ON breakdown_downtime_logs
FOR EACH ROW
BEGIN
  UPDATE breakdown_downtime_logs
  SET updated_at = datetime('now')
  WHERE id = NEW.id;
END;

-- Recalc total after insert
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

-- Recalc total after update
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

-- Recalc total after delete
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

/* =========================
   INDEXES (PERFORMANCE)
========================= */
CREATE INDEX IF NOT EXISTS idx_daily_hours_date ON daily_hours(work_date);
CREATE INDEX IF NOT EXISTS idx_daily_hours_asset ON daily_hours(asset_id);

CREATE INDEX IF NOT EXISTS idx_breakdowns_date ON breakdowns(breakdown_date);
CREATE INDEX IF NOT EXISTS idx_breakdowns_asset ON breakdowns(asset_id);
CREATE INDEX IF NOT EXISTS idx_breakdowns_status ON breakdowns(status);
CREATE INDEX IF NOT EXISTS idx_breakdowns_primary_wo ON breakdowns(primary_work_order_id);

CREATE INDEX IF NOT EXISTS idx_bd_logs_breakdown_id ON breakdown_downtime_logs(breakdown_id);
CREATE INDEX IF NOT EXISTS idx_bd_logs_log_date ON breakdown_downtime_logs(log_date);

CREATE INDEX IF NOT EXISTS idx_bd_components_breakdown_id ON breakdown_components(breakdown_id);
CREATE INDEX IF NOT EXISTS idx_bd_components_work_order_id ON breakdown_components(work_order_id);

CREATE INDEX IF NOT EXISTS idx_work_orders_status ON work_orders(status);
CREATE INDEX IF NOT EXISTS idx_stock_movements_part ON stock_movements(part_id);

/* =========================
   GET (Ground Engaging Tools) CHANGE SLIPS
========================= */
CREATE TABLE IF NOT EXISTS get_change_slips (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id INTEGER NOT NULL,
  slip_date TEXT NOT NULL,                 -- YYYY-MM-DD
  location TEXT,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS get_change_items (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  slip_id INTEGER NOT NULL,
  position TEXT,                           -- e.g. LH corner, bucket tooth #3
  part_code TEXT,
  part_name TEXT,
  qty REAL NOT NULL DEFAULT 1,
  reason TEXT,                             -- worn/broken/etc
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (slip_id) REFERENCES get_change_slips(id) ON DELETE CASCADE
);

/* =========================
   COMPONENT CHANGE SLIPS
========================= */
CREATE TABLE IF NOT EXISTS component_change_slips (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  asset_id INTEGER NOT NULL,
  slip_date TEXT NOT NULL,                 -- YYYY-MM-DD
  component TEXT NOT NULL,                 -- e.g. transmission, engine, diff
  serial_out TEXT,
  serial_in TEXT,
  hours_at_change REAL,                    -- machine hours at change
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
);

/* =========================
   LINK SLIPS TO WORK ORDERS (optional but powerful)
   (one slip can link to one WO; or multiple slips to one WO)
========================= */
CREATE TABLE IF NOT EXISTS work_order_links (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  work_order_id INTEGER NOT NULL,
  link_type TEXT NOT NULL,                 -- 'get_slip' | 'component_slip' | 'breakdown'
  link_id INTEGER NOT NULL,                -- id in the linked table
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  UNIQUE(work_order_id, link_type, link_id),
  FOREIGN KEY (work_order_id) REFERENCES work_orders(id) ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_get_slips_asset_date ON get_change_slips(asset_id, slip_date);
CREATE INDEX IF NOT EXISTS idx_component_slips_asset_date ON component_change_slips(asset_id, slip_date);
CREATE INDEX IF NOT EXISTS idx_wo_links_wo ON work_order_links(work_order_id);

ALTER TABLE breakdowns ADD COLUMN get_used INTEGER NOT NULL DEFAULT 0;
ALTER TABLE breakdowns ADD COLUMN get_hours_fitted REAL;
ALTER TABLE breakdowns ADD COLUMN get_hours_changed REAL;
ALTER TABLE assets ADD COLUMN baseline_fuel_l_per_hour REAL DEFAULT 5.0;

CREATE TABLE IF NOT EXISTS asset_hours (
  asset_id INTEGER PRIMARY KEY,
  total_hours REAL NOT NULL DEFAULT 0,
  last_updated TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (asset_id) REFERENCES assets(id)
);

INSERT INTO asset_hours (asset_id, total_hours, last_updated)
SELECT id, 0, datetime('now')
FROM assets
WHERE id NOT IN (SELECT asset_id FROM asset_hours);

CREATE TABLE IF NOT EXISTS ops_slip_reports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  site_code TEXT NOT NULL DEFAULT 'main',
  slip_type TEXT NOT NULL,
  asset_id INTEGER NOT NULL,
  report_date TEXT NOT NULL,
  payload_json TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  created_by TEXT,
  FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
);
CREATE INDEX IF NOT EXISTS idx_ops_slip_site_date ON ops_slip_reports(site_code, report_date);
CREATE INDEX IF NOT EXISTS idx_ops_slip_type ON ops_slip_reports(slip_type);

/* =========================
   TASKS
   Simple task management for team collaboration
========================= */
CREATE TABLE IF NOT EXISTS tasks (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  title TEXT NOT NULL,
  description TEXT,
  status TEXT NOT NULL DEFAULT 'open',     -- open | in_progress | done
  priority TEXT NOT NULL DEFAULT 'medium', -- high | medium | low
  project TEXT,                             -- Project/category name
  assigned_to TEXT,                         -- Username
  due_date TEXT,                           -- YYYY-MM-DD
  site_code TEXT NOT NULL DEFAULT 'main',
  created_by TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  completed_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_tasks_status ON tasks(status);
CREATE INDEX IF NOT EXISTS idx_tasks_priority ON tasks(priority);
CREATE INDEX IF NOT EXISTS idx_tasks_site ON tasks(site_code);
CREATE INDEX IF NOT EXISTS idx_tasks_assigned ON tasks(assigned_to);

/* =========================
   TASK COMMENTS
========================= */
CREATE TABLE IF NOT EXISTS task_comments (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  task_id INTEGER NOT NULL,
  comment TEXT NOT NULL,
  author TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (task_id) REFERENCES tasks(id) ON DELETE CASCADE
);
CREATE INDEX IF NOT EXISTS idx_task_comments_task ON task_comments(task_id);

/* =========================
   PROJECTS
========================= */
CREATE TABLE IF NOT EXISTS projects (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE,
  description TEXT,
  color TEXT DEFAULT '#3b82f6',
  site_code TEXT NOT NULL DEFAULT 'main',
  created_by TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_projects_site ON projects(site_code);