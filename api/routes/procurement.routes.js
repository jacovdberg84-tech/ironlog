import { db } from "../db/client.js";
import { ensureAuditTable } from "../utils/audit.js";
import { getRole, holdsAnyRole } from "../utils/request.js";
import registerRequisitionsRoutes from "./procurement/requisitions.routes.js";
import registerSuppliersRoutes from "./procurement/suppliers.routes.js";
import registerPurchaseOrdersRoutes from "./procurement/purchase-orders.routes.js";
import registerJournalsRoutes from "./procurement/journals.routes.js";

function getDepartment(req) {
  return String(req.headers["x-user-department"] || "").trim().toLowerCase() || null;
}
function getRoles(req) {
  const fromMany = String(req.headers["x-user-roles"] || "")
    .split(",")
    .map((x) => String(x || "").trim().toLowerCase())
    .filter(Boolean);
  const fromSingle = String(req.headers["x-user-role"] || "")
    .split(",")
    .map((x) => String(x || "").trim().toLowerCase())
    .filter(Boolean);
  const merged = Array.from(new Set([...fromMany, ...fromSingle]));
  return merged.length ? merged : ["operator"];
}
function isManagerScopeRole(req) {
  const roles = getRoles(req);
  return roles.some((r) => ["admin", "supervisor", "plant_manager", "site_manager", "executive"].includes(r));
}

function requireRoles(req, reply, roles) {
  const role = getRole(req);
  if (!holdsAnyRole(req, roles)) {
    reply.code(403).send({ error: `role '${role || "unknown"}' not allowed` });
    return false;
  }
  return true;
}

const REQUISITION_APPROVER_ROLES = [
  "admin",
  "supervisor",
  "plant_manager",
  "site_manager",
  "quality_manager",
  "hr_manager",
];
const STORE_EXECUTION_ROLES = ["admin", "supervisor", "storeman", "stores"];
const ROLE_PERMISSION_FALLBACK = {
  admin: ["*"],
  supervisor: ["procurement.requisition.create", "procurement.requisition.request_approval", "procurement.requisition.approve", "procurement.requisition.receive"],
  plant_manager: ["procurement.requisition.approve"],
  site_manager: ["procurement.requisition.approve"],
  quality_manager: ["procurement.requisition.approve"],
  hr_manager: ["procurement.requisition.approve"],
  procurement: ["procurement.requisition.create", "procurement.requisition.request_approval", "procurement.requisition.receive"],
  storeman: ["procurement.requisition.create", "procurement.requisition.request_approval", "procurement.requisition.receive"],
  stores: ["procurement.requisition.create", "procurement.requisition.request_approval", "procurement.requisition.receive"],
};

function getPermissions(req) {
  const fromHeader = String(req.headers["x-user-permissions"] || "")
    .split(",")
    .map((x) => String(x || "").trim())
    .filter(Boolean);
  if (fromHeader.length) return Array.from(new Set(fromHeader));
  const role = getRole(req);
  return ROLE_PERMISSION_FALLBACK[role] || [];
}

function requirePermission(req, reply, permissionKey) {
  const perms = getPermissions(req);
  if (perms.includes("*") || perms.includes(permissionKey)) return true;
  reply.code(403).send({ error: `permission '${permissionKey}' required` });
  return false;
}

export default async function procurementRoutes(app) {
  ensureAuditTable(db);

  db.prepare(`
    CREATE TABLE IF NOT EXISTS procurement_requisitions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      part_id INTEGER NOT NULL,
      qty_requested REAL NOT NULL,
      qty_received REAL NOT NULL DEFAULT 0,
      needed_by_date TEXT,
      supplier_name TEXT,
      po_number TEXT,
      bill_to TEXT,
      request_type TEXT,
      site_request_no TEXT,
      requester TEXT,
      notes TEXT,
      status TEXT NOT NULL DEFAULT 'draft',
      finalized_at TEXT,
      posted_at TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE RESTRICT
    )
  `).run();
  const cols = db.prepare(`PRAGMA table_info(procurement_requisitions)`).all();
  const hasCol = (c) => cols.some((r) => String(r.name) === c);
  if (!hasCol("supplier_name")) {
    db.prepare(`ALTER TABLE procurement_requisitions ADD COLUMN supplier_name TEXT`).run();
  }
  if (!hasCol("po_number")) {
    db.prepare(`ALTER TABLE procurement_requisitions ADD COLUMN po_number TEXT`).run();
  }
  if (!hasCol("bill_to")) {
    db.prepare(`ALTER TABLE procurement_requisitions ADD COLUMN bill_to TEXT`).run();
  }
  if (!hasCol("request_type")) {
    db.prepare(`ALTER TABLE procurement_requisitions ADD COLUMN request_type TEXT`).run();
  }
  if (!hasCol("site_request_no")) {
    db.prepare(`ALTER TABLE procurement_requisitions ADD COLUMN site_request_no TEXT`).run();
  }
  if (!hasCol("finalized_at")) {
    db.prepare(`ALTER TABLE procurement_requisitions ADD COLUMN finalized_at TEXT`).run();
  }
  if (!hasCol("posted_at")) {
    db.prepare(`ALTER TABLE procurement_requisitions ADD COLUMN posted_at TEXT`).run();
  }
  if (!hasCol("estimated_value")) {
    db.prepare(`ALTER TABLE procurement_requisitions ADD COLUMN estimated_value REAL`).run();
  }
  if (!hasCol("site_code")) {
    db.prepare(`ALTER TABLE procurement_requisitions ADD COLUMN site_code TEXT DEFAULT 'main'`).run();
  }
  if (!hasCol("department")) {
    db.prepare(`ALTER TABLE procurement_requisitions ADD COLUMN department TEXT`).run();
  }

  db.prepare(`
    CREATE TABLE IF NOT EXISTS procurement_requisition_lines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      requisition_id INTEGER NOT NULL,
      line_no INTEGER NOT NULL,
      product_code TEXT,
      part_id INTEGER,
      description TEXT,
      unit TEXT,
      quantity REAL NOT NULL DEFAULT 0,
      equipment_no TEXT,
      job_card TEXT,
      currency TEXT,
      gross_price REAL,
      discount_type TEXT,
      discount REAL,
      net_price REAL,
      line_value REAL,
      needed_by_date TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (requisition_id) REFERENCES procurement_requisitions(id) ON DELETE CASCADE,
      FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE SET NULL
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS procurement_requisition_attachments (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      requisition_id INTEGER NOT NULL,
      file_name TEXT NOT NULL,
      file_url TEXT,
      note TEXT,
      added_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (requisition_id) REFERENCES procurement_requisitions(id) ON DELETE CASCADE
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS procurement_requisition_approvers (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      requisition_id INTEGER NOT NULL,
      seq INTEGER NOT NULL,
      approver_name TEXT NOT NULL,
      approver_email TEXT,
      status TEXT NOT NULL DEFAULT 'pending',
      approved_at TEXT,
      comment TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (requisition_id) REFERENCES procurement_requisitions(id) ON DELETE CASCADE
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS approval_requests (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      module TEXT NOT NULL,
      action TEXT NOT NULL,
      entity_type TEXT,
      entity_id TEXT,
      status TEXT NOT NULL DEFAULT 'pending',
      payload_json TEXT,
      requested_by TEXT,
      requested_role TEXT,
      approved_by TEXT,
      approved_role TEXT,
      approved_at TEXT,
      rejected_by TEXT,
      rejected_role TEXT,
      rejected_at TEXT,
      decision_note TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS suppliers (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      supplier_code TEXT NOT NULL UNIQUE,
      supplier_name TEXT NOT NULL,
      active INTEGER NOT NULL DEFAULT 1,
      lead_time_days INTEGER NOT NULL DEFAULT 7,
      currency TEXT NOT NULL DEFAULT 'USD',
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS supplier_part_catalog (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      supplier_id INTEGER NOT NULL,
      part_id INTEGER NOT NULL,
      supplier_part_code TEXT,
      lead_time_days INTEGER,
      last_price REAL,
      currency TEXT,
      effective_date TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE (supplier_id, part_id),
      FOREIGN KEY (supplier_id) REFERENCES suppliers(id) ON DELETE CASCADE,
      FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE CASCADE
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS procurement_purchase_orders (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      po_number TEXT NOT NULL UNIQUE,
      requisition_id INTEGER,
      supplier_id INTEGER,
      site_code TEXT DEFAULT 'main',
      currency TEXT DEFAULT 'USD',
      status TEXT NOT NULL DEFAULT 'draft',
      approved_at TEXT,
      sent_at TEXT,
      notes TEXT,
      subtotal REAL NOT NULL DEFAULT 0,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (requisition_id) REFERENCES procurement_requisitions(id) ON DELETE SET NULL,
      FOREIGN KEY (supplier_id) REFERENCES suppliers(id) ON DELETE SET NULL
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS procurement_purchase_order_lines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      po_id INTEGER NOT NULL,
      line_no INTEGER NOT NULL,
      requisition_line_id INTEGER,
      part_id INTEGER,
      description TEXT,
      quantity_ordered REAL NOT NULL DEFAULT 0,
      quantity_received REAL NOT NULL DEFAULT 0,
      unit_price REAL NOT NULL DEFAULT 0,
      line_total REAL NOT NULL DEFAULT 0,
      needed_by_date TEXT,
      cost_center_code TEXT,
      labor_tag TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE (po_id, line_no),
      FOREIGN KEY (po_id) REFERENCES procurement_purchase_orders(id) ON DELETE CASCADE,
      FOREIGN KEY (requisition_line_id) REFERENCES procurement_requisition_lines(id) ON DELETE SET NULL,
      FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE SET NULL
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS procurement_goods_receipts (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      po_id INTEGER NOT NULL,
      receipt_number TEXT NOT NULL UNIQUE,
      receipt_date TEXT NOT NULL,
      received_by TEXT,
      location_code TEXT,
      status TEXT NOT NULL DEFAULT 'posted',
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (po_id) REFERENCES procurement_purchase_orders(id) ON DELETE CASCADE
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS procurement_goods_receipt_lines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      receipt_id INTEGER NOT NULL,
      po_line_id INTEGER NOT NULL,
      part_id INTEGER,
      quantity_received REAL NOT NULL DEFAULT 0,
      unit_price REAL NOT NULL DEFAULT 0,
      line_total REAL NOT NULL DEFAULT 0,
      cost_center_code TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (receipt_id) REFERENCES procurement_goods_receipts(id) ON DELETE CASCADE,
      FOREIGN KEY (po_line_id) REFERENCES procurement_purchase_order_lines(id) ON DELETE CASCADE,
      FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE SET NULL
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS procurement_invoices (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      po_id INTEGER NOT NULL,
      invoice_number TEXT NOT NULL UNIQUE,
      supplier_id INTEGER,
      invoice_date TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT 'captured',
      currency TEXT DEFAULT 'USD',
      subtotal REAL NOT NULL DEFAULT 0,
      tax REAL NOT NULL DEFAULT 0,
      total REAL NOT NULL DEFAULT 0,
      captured_by TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (po_id) REFERENCES procurement_purchase_orders(id) ON DELETE CASCADE,
      FOREIGN KEY (supplier_id) REFERENCES suppliers(id) ON DELETE SET NULL
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS procurement_invoice_lines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      invoice_id INTEGER NOT NULL,
      po_line_id INTEGER,
      part_id INTEGER,
      description TEXT,
      quantity_invoiced REAL NOT NULL DEFAULT 0,
      unit_price REAL NOT NULL DEFAULT 0,
      line_total REAL NOT NULL DEFAULT 0,
      cost_center_code TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (invoice_id) REFERENCES procurement_invoices(id) ON DELETE CASCADE,
      FOREIGN KEY (po_line_id) REFERENCES procurement_purchase_order_lines(id) ON DELETE SET NULL,
      FOREIGN KEY (part_id) REFERENCES parts(id) ON DELETE SET NULL
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS procurement_match_exceptions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      po_id INTEGER NOT NULL,
      po_line_id INTEGER,
      invoice_id INTEGER,
      invoice_line_id INTEGER,
      receipt_id INTEGER,
      receipt_line_id INTEGER,
      exception_type TEXT NOT NULL,
      severity TEXT NOT NULL DEFAULT 'warn',
      status TEXT NOT NULL DEFAULT 'open',
      details_json TEXT,
      assigned_to TEXT,
      resolved_by TEXT,
      resolved_at TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS finance_journal_staging (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      batch_id TEXT NOT NULL,
      tx_date TEXT NOT NULL,
      source_module TEXT NOT NULL,
      source_type TEXT NOT NULL,
      source_id TEXT NOT NULL,
      account_code TEXT NOT NULL,
      cost_center_code TEXT,
      description TEXT,
      debit REAL NOT NULL DEFAULT 0,
      credit REAL NOT NULL DEFAULT 0,
      currency TEXT DEFAULT 'USD',
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS finance_posting_runs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      run_number TEXT NOT NULL UNIQUE,
      period TEXT NOT NULL,
      start_date TEXT NOT NULL,
      end_date TEXT NOT NULL,
      run_type TEXT NOT NULL DEFAULT 'summary',
      status TEXT NOT NULL DEFAULT 'draft',
      currency TEXT DEFAULT 'USD',
      total_debit REAL NOT NULL DEFAULT 0,
      total_credit REAL NOT NULL DEFAULT 0,
      line_count INTEGER NOT NULL DEFAULT 0,
      notes TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      exported_at TEXT,
      exported_by TEXT,
      posted_at TEXT,
      posted_by TEXT,
      posted_reference TEXT,
      reversed_at TEXT,
      reversed_by TEXT,
      reversed_reason TEXT
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS finance_posting_run_lines (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      run_id INTEGER NOT NULL,
      category TEXT NOT NULL,
      tx_date TEXT NOT NULL,
      source_module TEXT NOT NULL,
      source_type TEXT NOT NULL,
      source_ref TEXT,
      account_code TEXT NOT NULL,
      cost_center_code TEXT,
      site_code TEXT,
      equipment_type TEXT,
      asset_code TEXT,
      description TEXT,
      debit REAL NOT NULL DEFAULT 0,
      credit REAL NOT NULL DEFAULT 0,
      currency TEXT DEFAULT 'USD',
      qty REAL,
      unit_cost REAL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (run_id) REFERENCES finance_posting_runs(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_fprl_run ON finance_posting_run_lines(run_id)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_fprl_cat ON finance_posting_run_lines(category)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_fprl_cc ON finance_posting_run_lines(cost_center_code)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS finance_period_locks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      period TEXT NOT NULL UNIQUE,
      status TEXT NOT NULL DEFAULT 'open',
      locked_by TEXT,
      locked_at TEXT,
      reopened_by TEXT,
      reopened_at TEXT,
      reopen_reason TEXT,
      closed_by TEXT,
      closed_at TEXT,
      notes TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  const bySiteReq = `
    LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
  `;
  const getPartByCode = db.prepare(`SELECT id, part_code, part_name FROM parts WHERE part_code = ?`);
  const getLocationByCode = db.prepare(`SELECT id, location_code FROM stock_locations WHERE location_code = ?`);

  function nextPONumber() {
    const row = db.prepare(`SELECT IFNULL(MAX(id), 0) + 1 AS n FROM procurement_purchase_orders`).get();
    const n = Number(row?.n || 1);
    const year = new Date().getFullYear();
    return `PO-${year}-${String(n).padStart(5, "0")}`;
  }
  function nextReceiptNumber() {
    const row = db.prepare(`SELECT IFNULL(MAX(id), 0) + 1 AS n FROM procurement_goods_receipts`).get();
    const n = Number(row?.n || 1);
    const year = new Date().getFullYear();
    return `GRN-${year}-${String(n).padStart(5, "0")}`;
  }

  function derivePoStatus(poId) {
    const lines = db.prepare(`
      SELECT quantity_ordered, quantity_received
      FROM procurement_purchase_order_lines
      WHERE po_id = ?
    `).all(poId);
    if (!lines.length) return "draft";
    const totalOrdered = lines.reduce((s, l) => s + Number(l.quantity_ordered || 0), 0);
    const totalReceived = lines.reduce((s, l) => s + Number(l.quantity_received || 0), 0);
    if (totalReceived <= 0) return "approved";
    if (totalReceived + 1e-9 >= totalOrdered) return "received";
    return "partially_received";
  }

  /* ================================================================
     FINANCE POSTING RUNS (summarized journals)
     Categories: parts, labor, downtime, fuel, lube, procurement_grn, procurement_ap
     ---------------------------------------------------------------
     POST /journals/summarize             -> build draft run from source data
     GET  /journals/runs                  -> list runs
     GET  /journals/runs/:id              -> run header + summary
     GET  /journals/runs/:id/lines        -> run lines
     POST /journals/runs/:id/mark-exported
     POST /journals/runs/:id/mark-posted
     POST /journals/runs/:id/reverse
     GET  /journals/runs/:id/export.csv
     GET  /journals/runs/:id/export.xlsx
  ================================================================ */

  function nextRunNumber(period) {
    const row = db.prepare(`SELECT IFNULL(MAX(id), 0) + 1 AS n FROM finance_posting_runs`).get();
    const n = Number(row?.n || 1);
    return `FIN-${String(period || "").replace(/-/g, "")}-${String(n).padStart(5, "0")}`;
  }

  function tableHasColumn(table, col) {
    try {
      const rows = db.prepare(`PRAGMA table_info(${table})`).all();
      return rows.some((r) => String(r.name) === col);
    } catch { return false; }
  }

  function readCostSetting(key, fallback) {
    try {
      const row = db.prepare(`SELECT value FROM cost_settings WHERE key = ? LIMIT 1`).get(key);
      const v = Number(row?.value);
      return Number.isFinite(v) ? v : fallback;
    } catch { return fallback; }
  }

  function periodForDate(d) {
    const s = String(d || "").trim();
    if (/^\d{4}-\d{2}/.test(s)) return s.slice(0, 7);
    return new Date().toISOString().slice(0, 7);
  }

  function isPeriodLocked(period) {
    const row = db.prepare(`SELECT status FROM finance_period_locks WHERE period = ?`).get(String(period || ""));
    const s = String(row?.status || "open").toLowerCase();
    return s === "locked" || s === "closed";
  }

  function recalcRunTotals(runId) {
    const row = db.prepare(`
      SELECT COUNT(*) AS lines,
             COALESCE(SUM(debit), 0) AS d,
             COALESCE(SUM(credit), 0) AS c
      FROM finance_posting_run_lines WHERE run_id = ?
    `).get(runId);
    db.prepare(`
      UPDATE finance_posting_runs
      SET line_count = ?, total_debit = ?, total_credit = ?
      WHERE id = ?
    `).run(Number(row?.lines || 0), Number(row?.d || 0), Number(row?.c || 0), runId);
  }

  // Route groups live in routes/procurement/. They receive the shared helpers above through ctx.
  const ctx = {
    REQUISITION_APPROVER_ROLES,
    STORE_EXECUTION_ROLES,
    bySiteReq,
    derivePoStatus,
    getDepartment,
    getLocationByCode,
    getPartByCode,
    isManagerScopeRole,
    isPeriodLocked,
    nextPONumber,
    nextReceiptNumber,
    nextRunNumber,
    periodForDate,
    readCostSetting,
    recalcRunTotals,
    requirePermission,
    requireRoles,
    tableHasColumn,
  };
  registerRequisitionsRoutes(app, ctx);
  registerSuppliersRoutes(app, ctx);
  registerPurchaseOrdersRoutes(app, ctx);
  registerJournalsRoutes(app, ctx);
}
