import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";

function getRole(req) {
  return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
}

function getUser(req) {
  return String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
}
function getSiteCode(req) {
  return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
}
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
  if (!roles.includes(role)) {
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

  app.post("/requisitions", async (req, reply) => {
    if (!requirePermission(req, reply, "procurement.requisition.create")) return;
    if (!requireRoles(req, reply, STORE_EXECUTION_ROLES)) return;
    const part_code = String(req.body?.part_code || "").trim();
    const qty_requested = Number(req.body?.qty_requested ?? 0);
    const needed_by_date =
      req.body?.needed_by_date != null && String(req.body.needed_by_date).trim() !== ""
        ? String(req.body.needed_by_date).trim()
        : null;
    const supplier_name =
      req.body?.supplier_name != null && String(req.body.supplier_name).trim() !== ""
        ? String(req.body.supplier_name).trim()
        : null;
    const po_number =
      req.body?.po_number != null && String(req.body.po_number).trim() !== ""
        ? String(req.body.po_number).trim()
        : null;
    const bill_to =
      req.body?.bill_to != null && String(req.body.bill_to).trim() !== ""
        ? String(req.body.bill_to).trim()
        : "workshop";
    const request_type =
      req.body?.request_type != null && String(req.body.request_type).trim() !== ""
        ? String(req.body.request_type).trim()
        : "site";
    const notes =
      req.body?.notes != null && String(req.body.notes).trim() !== ""
        ? String(req.body.notes).trim()
        : null;
    const estimated_value =
      req.body?.estimated_value != null && String(req.body.estimated_value).trim() !== ""
        ? Number(req.body.estimated_value)
        : null;
    if (!part_code) return reply.code(400).send({ error: "part_code is required" });
    if (!Number.isFinite(qty_requested) || qty_requested <= 0) {
      return reply.code(400).send({ error: "qty_requested must be > 0" });
    }
    if (estimated_value != null && (!Number.isFinite(estimated_value) || estimated_value < 0)) {
      return reply.code(400).send({ error: "estimated_value must be a valid number >= 0" });
    }

    const part = db.prepare(`SELECT id, part_code, part_name FROM parts WHERE part_code = ?`).get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });

    const requester = getUser(req);
    const siteCode = getSiteCode(req);
    const department = getDepartment(req);
    const ins = db.prepare(`
      INSERT INTO procurement_requisitions (
        part_id, qty_requested, qty_received, needed_by_date, supplier_name, po_number, bill_to, request_type, requester, notes, estimated_value, site_code, department, status, created_at, updated_at
      ) VALUES (?, ?, 0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'draft', datetime('now'), datetime('now'))
    `).run(part.id, qty_requested, needed_by_date, supplier_name, po_number, bill_to, request_type, requester, notes, estimated_value, siteCode, department);
    const id = Number(ins.lastInsertRowid);

    db.prepare(`
      INSERT INTO procurement_requisition_lines (
        requisition_id, line_no, product_code, part_id, description, quantity
      ) VALUES (?, 1, ?, ?, ?, ?)
    `).run(id, part.part_code, part.id, part.part_name, qty_requested);

    writeAudit(db, req, {
      module: "procurement",
      action: "requisition_create",
      entity_type: "requisition",
      entity_id: id,
      payload: { part_code, qty_requested, needed_by_date, supplier_name, po_number, bill_to, request_type, estimated_value },
    });

    return { ok: true, id, part_code: part.part_code, part_name: part.part_name, qty_requested, supplier_name, po_number, bill_to, request_type, estimated_value };
  });

  app.get("/requisitions", async (req) => {
    const status = String(req.query?.status || "").trim().toLowerCase();
    const siteCode = getSiteCode(req);
    const reqDepartment = getDepartment(req);
    const where = [];
    const params = [];
    where.push("LOWER(TRIM(COALESCE(pr.site_code, 'main'))) = ?");
    params.push(siteCode);
    if (!isManagerScopeRole(req) && reqDepartment) {
      where.push("LOWER(TRIM(COALESCE(pr.department, ''))) = ?");
      params.push(reqDepartment);
    }
    if (status) {
      where.push("LOWER(pr.status) = ?");
      params.push(status);
    }
    const rows = db.prepare(`
      SELECT
        pr.id,
        p.part_code,
        p.part_name,
        pr.qty_requested,
        pr.qty_received,
        pr.needed_by_date,
        pr.supplier_name,
        pr.po_number,
        pr.bill_to,
        pr.request_type,
        pr.site_request_no,
        pr.requester,
        pr.notes,
        pr.estimated_value,
        pr.status,
        pr.finalized_at,
        pr.posted_at,
        pr.created_at,
        pr.updated_at,
        (
          SELECT ar.id
          FROM approval_requests ar
          WHERE ar.module = 'procurement'
            AND ar.entity_type = 'requisition'
            AND ar.entity_id = CAST(pr.id AS TEXT)
          ORDER BY ar.id DESC
          LIMIT 1
        ) AS latest_approval_id,
        (
          SELECT ar.status
          FROM approval_requests ar
          WHERE ar.module = 'procurement'
            AND ar.entity_type = 'requisition'
            AND ar.entity_id = CAST(pr.id AS TEXT)
          ORDER BY ar.id DESC
          LIMIT 1
        ) AS latest_approval_status
      FROM procurement_requisitions pr
      JOIN parts p ON p.id = pr.part_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY pr.id DESC
      LIMIT 300
    `).all(...params).map((r) => ({
      ...r,
      qty_requested: Number(r.qty_requested || 0),
      qty_received: Number(r.qty_received || 0),
      estimated_value: r.estimated_value == null ? null : Number(r.estimated_value || 0),
      qty_outstanding: Number((Number(r.qty_requested || 0) - Number(r.qty_received || 0)).toFixed(2)),
    }));
    return { ok: true, rows };
  });

  app.get("/requisitions/:id/detail", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const siteCode = getSiteCode(req);
    const reqDepartment = getDepartment(req);
    const enforceDept = !isManagerScopeRole(req) && reqDepartment;

    const header = db.prepare(`
      SELECT pr.*, p.part_code, p.part_name
      FROM procurement_requisitions pr
      LEFT JOIN parts p ON p.id = pr.part_id
      WHERE pr.id = ?
        AND LOWER(TRIM(COALESCE(pr.site_code, 'main'))) = ?
        ${enforceDept ? "AND LOWER(TRIM(COALESCE(pr.department, ''))) = ?" : ""}
    `).get(...(enforceDept ? [id, siteCode, reqDepartment] : [id, siteCode]));
    if (!header) return reply.code(404).send({ error: "requisition not found" });

    const lines = db.prepare(`
      SELECT l.*, p.part_code, p.part_name
      FROM procurement_requisition_lines l
      LEFT JOIN parts p ON p.id = l.part_id
      WHERE l.requisition_id = ?
      ORDER BY l.line_no ASC, l.id ASC
    `).all(id);
    const attachments = db.prepare(`
      SELECT id, file_name, file_url, note, added_by, created_at
      FROM procurement_requisition_attachments
      WHERE requisition_id = ?
      ORDER BY id DESC
    `).all(id);
    const approvers = db.prepare(`
      SELECT id, seq, approver_name, approver_email, status, approved_at, comment, created_at
      FROM procurement_requisition_approvers
      WHERE requisition_id = ?
      ORDER BY seq ASC, id ASC
    `).all(id);
    return { ok: true, requisition: header, lines, attachments, approvers };
  });

  app.post("/requisitions/:id/lines", async (req, reply) => {
    if (!requireRoles(req, reply, STORE_EXECUTION_ROLES)) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const reqn = db.prepare(`SELECT id, status FROM procurement_requisitions WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?`).get(id, getSiteCode(req));
    if (!reqn) return reply.code(404).send({ error: "requisition not found" });
    if (["finalized", "posted", "approval_in_progress", "approved_all", "po_ready", "received"].includes(String(reqn.status || ""))) {
      return reply.code(409).send({ error: `cannot add line when status is ${reqn.status}` });
    }

    const product_code = String(req.body?.product_code || req.body?.part_code || "").trim();
    const quantity = Number(req.body?.quantity ?? 0);
    const description =
      req.body?.description != null && String(req.body.description).trim() !== ""
        ? String(req.body.description).trim()
        : null;
    if (!product_code) return reply.code(400).send({ error: "product_code is required" });
    if (!Number.isFinite(quantity) || quantity <= 0) return reply.code(400).send({ error: "quantity must be > 0" });

    const part = db.prepare(`SELECT id, part_code, part_name FROM parts WHERE part_code = ?`).get(product_code);
    const next = db.prepare(`SELECT COALESCE(MAX(line_no), 0) + 1 AS n FROM procurement_requisition_lines WHERE requisition_id = ?`).get(id);
    const line_no = Number(next?.n || 1);
    const ins = db.prepare(`
      INSERT INTO procurement_requisition_lines (
        requisition_id, line_no, product_code, part_id, description, unit, quantity, equipment_no, job_card, currency, gross_price, discount_type, discount, net_price, line_value, needed_by_date
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(
      id,
      line_no,
      product_code,
      part?.id || null,
      description || part?.part_name || null,
      req.body?.unit != null ? String(req.body.unit).trim() : null,
      quantity,
      req.body?.equipment_no != null ? String(req.body.equipment_no).trim() : null,
      req.body?.job_card != null ? String(req.body.job_card).trim() : null,
      req.body?.currency != null ? String(req.body.currency).trim() : null,
      req.body?.gross_price != null && req.body.gross_price !== "" ? Number(req.body.gross_price) : null,
      req.body?.discount_type != null ? String(req.body.discount_type).trim() : null,
      req.body?.discount != null && req.body.discount !== "" ? Number(req.body.discount) : null,
      req.body?.net_price != null && req.body.net_price !== "" ? Number(req.body.net_price) : null,
      req.body?.line_value != null && req.body.line_value !== "" ? Number(req.body.line_value) : null,
      req.body?.needed_by_date != null ? String(req.body.needed_by_date).trim() : null
    );
    return { ok: true, line_id: Number(ins.lastInsertRowid), line_no };
  });

  app.post("/requisitions/:id/attachments", async (req, reply) => {
    if (!requireRoles(req, reply, STORE_EXECUTION_ROLES)) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const reqn = db.prepare(`SELECT id FROM procurement_requisitions WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?`).get(id, getSiteCode(req));
    if (!reqn) return reply.code(404).send({ error: "requisition not found" });
    const file_name = String(req.body?.file_name || "").trim();
    if (!file_name) return reply.code(400).send({ error: "file_name is required" });
    const ins = db.prepare(`
      INSERT INTO procurement_requisition_attachments (requisition_id, file_name, file_url, note, added_by)
      VALUES (?, ?, ?, ?, ?)
    `).run(
      id,
      file_name,
      req.body?.file_url != null ? String(req.body.file_url).trim() : null,
      req.body?.note != null ? String(req.body.note).trim() : null,
      getUser(req)
    );
    return { ok: true, attachment_id: Number(ins.lastInsertRowid) };
  });

  app.post("/requisitions/:id/finalize", async (req, reply) => {
    if (!requireRoles(req, reply, STORE_EXECUTION_ROLES)) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const reqn = db.prepare(`SELECT id, status, site_request_no FROM procurement_requisitions WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?`).get(id, getSiteCode(req));
    if (!reqn) return reply.code(404).send({ error: "requisition not found" });
    const cnt = db.prepare(`SELECT COUNT(*) AS c FROM procurement_requisition_lines WHERE requisition_id = ?`).get(id);
    if (Number(cnt?.c || 0) <= 0) return reply.code(409).send({ error: "cannot finalize without lines" });

    let site_request_no = String(reqn.site_request_no || "").trim();
    if (!site_request_no) {
      const now = new Date();
      const y = now.getFullYear();
      const m = String(now.getMonth() + 1).padStart(2, "0");
      site_request_no = `SR-${y}${m}-${String(id).padStart(5, "0")}`;
    }

    db.prepare(`
      UPDATE procurement_requisitions
      SET status = 'finalized', site_request_no = ?, finalized_at = datetime('now'), updated_at = datetime('now')
      WHERE id = ?
    `).run(site_request_no, id);
    return { ok: true, id, status: "finalized", site_request_no };
  });

  app.post("/requisitions/:id/post", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const reqn = db.prepare(`SELECT id, status FROM procurement_requisitions WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?`).get(id, getSiteCode(req));
    if (!reqn) return reply.code(404).send({ error: "requisition not found" });
    if (!["finalized", "posted"].includes(String(reqn.status || ""))) {
      return reply.code(409).send({ error: "only finalized requisitions can be posted" });
    }
    db.prepare(`
      UPDATE procurement_requisitions
      SET status = 'posted', posted_at = datetime('now'), updated_at = datetime('now')
      WHERE id = ?
    `).run(id);
    return { ok: true, id, status: "posted" };
  });

  app.post("/requisitions/:id/approvers", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const reqn = db.prepare(`SELECT id FROM procurement_requisitions WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?`).get(id, getSiteCode(req));
    if (!reqn) return reply.code(404).send({ error: "requisition not found" });
    const list = Array.isArray(req.body?.approvers) ? req.body.approvers : [];
    if (!list.length) return reply.code(400).send({ error: "approvers array is required" });

    const tx = db.transaction(() => {
      db.prepare(`DELETE FROM procurement_requisition_approvers WHERE requisition_id = ?`).run(id);
      const ins = db.prepare(`
        INSERT INTO procurement_requisition_approvers (requisition_id, seq, approver_name, approver_email, status)
        VALUES (?, ?, ?, ?, 'pending')
      `);
      list.forEach((a, idx) => {
        const name = String(a?.name || "").trim();
        const email = String(a?.email || "").trim();
        if (!name) return;
        ins.run(id, idx + 1, name, email || null);
      });
    });
    tx();
    return { ok: true };
  });

  app.post("/requisitions/:id/send-approval", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const reqn = db.prepare(`SELECT id, status FROM procurement_requisitions WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?`).get(id, getSiteCode(req));
    if (!reqn) return reply.code(404).send({ error: "requisition not found" });
    if (!["posted", "approval_in_progress"].includes(String(reqn.status || ""))) {
      return reply.code(409).send({ error: "requisition must be posted before approval routing" });
    }
    const approvers = db.prepare(`
      SELECT id, seq, approver_name, approver_email, status
      FROM procurement_requisition_approvers
      WHERE requisition_id = ?
      ORDER BY seq ASC
    `).all(id);
    if (!approvers.length) return reply.code(409).send({ error: "no approvers configured" });
    db.prepare(`
      UPDATE procurement_requisitions
      SET status = 'approval_in_progress', updated_at = datetime('now')
      WHERE id = ?
    `).run(id);
    writeAudit(db, req, {
      module: "procurement",
      action: "send_approval",
      entity_type: "requisition",
      entity_id: id,
      payload: { approvers: approvers.map((a) => ({ seq: a.seq, name: a.approver_name, email: a.approver_email })) },
    });
    return { ok: true, status: "approval_in_progress", next_approver: approvers.find((a) => String(a.status) === "pending") || null };
  });

  app.post("/requisitions/:id/approve", async (req, reply) => {
    if (!requirePermission(req, reply, "procurement.requisition.approve")) return;
    if (!requireRoles(req, reply, REQUISITION_APPROVER_ROLES)) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const approver_name = String(req.body?.approver_name || getUser(req)).trim();
    if (!approver_name) return reply.code(400).send({ error: "approver_name is required" });
    const comment =
      req.body?.comment != null && String(req.body.comment).trim() !== ""
        ? String(req.body.comment).trim()
        : null;

    const reqn = db.prepare(`SELECT id, status FROM procurement_requisitions WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?`).get(id, getSiteCode(req));
    if (!reqn) return reply.code(404).send({ error: "requisition not found" });
    if (!["approval_in_progress"].includes(String(reqn.status || ""))) {
      return reply.code(409).send({ error: "requisition is not in approval flow" });
    }

    const next = db.prepare(`
      SELECT id, seq, approver_name, status
      FROM procurement_requisition_approvers
      WHERE requisition_id = ?
        AND status = 'pending'
      ORDER BY seq ASC
      LIMIT 1
    `).get(id);
    if (!next) {
      db.prepare(`UPDATE procurement_requisitions SET status = 'approved_all', updated_at = datetime('now') WHERE id = ?`).run(id);
      return { ok: true, status: "approved_all", done: true };
    }
    if (String(next.approver_name || "").toLowerCase() !== approver_name.toLowerCase()) {
      return reply.code(409).send({ error: `next approver is '${next.approver_name}', not '${approver_name}'` });
    }

    db.prepare(`
      UPDATE procurement_requisition_approvers
      SET status = 'approved', approved_at = datetime('now'), comment = ?
      WHERE id = ?
    `).run(comment, Number(next.id));

    const remaining = db.prepare(`
      SELECT COUNT(*) AS c
      FROM procurement_requisition_approvers
      WHERE requisition_id = ? AND status = 'pending'
    `).get(id);

    if (Number(remaining?.c || 0) <= 0) {
      db.prepare(`
        UPDATE procurement_requisitions
        SET status = 'approved_all', updated_at = datetime('now')
        WHERE id = ?
      `).run(id);
      return { ok: true, status: "approved_all", done: true };
    }
    return { ok: true, status: "approval_in_progress", done: false };
  });

  app.post("/requisitions/:id/request-approval", async (req, reply) => {
    if (!requirePermission(req, reply, "procurement.requisition.request_approval")) return;
    if (!requireRoles(req, reply, STORE_EXECUTION_ROLES)) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });

    const row = db.prepare(`
      SELECT pr.id, pr.status, p.part_code, p.part_name, pr.qty_requested, pr.qty_received
      FROM procurement_requisitions pr
      JOIN parts p ON p.id = pr.part_id
      WHERE pr.id = ?
        AND LOWER(TRIM(COALESCE(pr.site_code, 'main'))) = ?
    `).get(id, getSiteCode(req));
    if (!row) return reply.code(404).send({ error: "requisition not found" });

    const pending = db.prepare(`
      SELECT id FROM approval_requests
      WHERE module = 'procurement'
        AND action = 'approve_requisition'
        AND entity_type = 'requisition'
        AND entity_id = ?
        AND status = 'pending'
      LIMIT 1
    `).get(String(id));
    if (pending) return { ok: true, pending_approval: true, request_id: Number(pending.id), duplicate: true };

    const payload_json = JSON.stringify({
      requisition_id: id,
      part_code: row.part_code,
      qty_requested: Number(row.qty_requested || 0),
    });
    const ins = db.prepare(`
      INSERT INTO approval_requests (module, action, entity_type, entity_id, status, payload_json, requested_by, requested_role)
      VALUES ('procurement', 'approve_requisition', 'requisition', ?, 'pending', ?, ?, ?)
    `).run(String(id), payload_json, getUser(req), getRole(req));

    db.prepare(`UPDATE procurement_requisitions SET status = 'pending_approval', updated_at = datetime('now') WHERE id = ?`).run(id);
    return { ok: true, pending_approval: true, request_id: Number(ins.lastInsertRowid) };
  });

  app.post("/requisitions/:id/request-receive", async (req, reply) => {
    if (!requirePermission(req, reply, "procurement.requisition.receive")) return;
    if (!requireRoles(req, reply, STORE_EXECUTION_ROLES)) return;
    const id = Number(req.params?.id || 0);
    const qty_receive = Number(req.body?.qty_receive ?? 0);
    const reference = String(req.body?.reference || `requisition:${id}`).trim() || `requisition:${id}`;
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    if (!Number.isFinite(qty_receive) || qty_receive <= 0) return reply.code(400).send({ error: "qty_receive must be > 0" });

    const row = db.prepare(`
      SELECT pr.id, pr.status, p.part_code, p.part_name, pr.qty_requested, pr.qty_received
      FROM procurement_requisitions pr
      JOIN parts p ON p.id = pr.part_id
      WHERE pr.id = ?
        AND LOWER(TRIM(COALESCE(pr.site_code, 'main'))) = ?
    `).get(id, getSiteCode(req));
    if (!row) return reply.code(404).send({ error: "requisition not found" });

    const outstanding = Number(row.qty_requested || 0) - Number(row.qty_received || 0);
    if (outstanding <= 0) return reply.code(409).send({ error: "requisition already fully received" });
    if (qty_receive > outstanding) return reply.code(409).send({ error: `qty_receive exceeds outstanding (${outstanding})` });

    const payload_json = JSON.stringify({
      requisition_id: id,
      part_code: row.part_code,
      qty_receive,
      reference,
    });
    const ins = db.prepare(`
      INSERT INTO approval_requests (module, action, entity_type, entity_id, status, payload_json, requested_by, requested_role)
      VALUES ('procurement', 'receive_requisition', 'requisition', ?, 'pending', ?, ?, ?)
    `).run(String(id), payload_json, getUser(req), getRole(req));
    return { ok: true, pending_approval: true, request_id: Number(ins.lastInsertRowid) };
  });

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

  app.post("/suppliers", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores", "procurement"])) return;
    const supplier_code = String(req.body?.supplier_code || "").trim().toUpperCase();
    const supplier_name = String(req.body?.supplier_name || "").trim();
    if (!supplier_code || !supplier_name) return reply.code(400).send({ error: "supplier_code and supplier_name are required" });
    const lead_time_days = Math.max(0, Number(req.body?.lead_time_days || 7));
    const currency = String(req.body?.currency || "USD").trim().toUpperCase() || "USD";
    db.prepare(`
      INSERT INTO suppliers (supplier_code, supplier_name, lead_time_days, currency, active, updated_at)
      VALUES (?, ?, ?, ?, 1, datetime('now'))
      ON CONFLICT(supplier_code) DO UPDATE SET
        supplier_name = excluded.supplier_name,
        lead_time_days = excluded.lead_time_days,
        currency = excluded.currency,
        active = 1,
        updated_at = datetime('now')
    `).run(supplier_code, supplier_name, lead_time_days, currency);
    return { ok: true, supplier_code, supplier_name, lead_time_days, currency };
  });

  app.get("/suppliers", async () => {
    const rows = db.prepare(`
      SELECT id, supplier_code, supplier_name, active, lead_time_days, currency, created_at, updated_at
      FROM suppliers
      ORDER BY supplier_name ASC
      LIMIT 300
    `).all();
    return { ok: true, rows };
  });

  app.post("/supplier-catalog", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores", "procurement"])) return;
    const supplier_code = String(req.body?.supplier_code || "").trim().toUpperCase();
    const part_code = String(req.body?.part_code || "").trim();
    if (!supplier_code || !part_code) return reply.code(400).send({ error: "supplier_code and part_code are required" });
    const supplier = db.prepare(`SELECT id FROM suppliers WHERE supplier_code = ?`).get(supplier_code);
    if (!supplier) return reply.code(404).send({ error: "supplier not found" });
    const part = getPartByCode.get(part_code);
    if (!part) return reply.code(404).send({ error: "part not found" });
    const supplier_part_code = req.body?.supplier_part_code != null ? String(req.body.supplier_part_code).trim() : null;
    const lead_time_days = req.body?.lead_time_days != null ? Math.max(0, Number(req.body.lead_time_days || 0)) : null;
    const last_price = req.body?.last_price != null ? Number(req.body.last_price) : null;
    const currency = req.body?.currency != null ? String(req.body.currency).trim().toUpperCase() : null;
    const effective_date = req.body?.effective_date != null ? String(req.body.effective_date).trim() : null;
    db.prepare(`
      INSERT INTO supplier_part_catalog (
        supplier_id, part_id, supplier_part_code, lead_time_days, last_price, currency, effective_date, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))
      ON CONFLICT(supplier_id, part_id) DO UPDATE SET
        supplier_part_code = excluded.supplier_part_code,
        lead_time_days = excluded.lead_time_days,
        last_price = excluded.last_price,
        currency = excluded.currency,
        effective_date = excluded.effective_date,
        updated_at = datetime('now')
    `).run(Number(supplier.id), Number(part.id), supplier_part_code, lead_time_days, last_price, currency, effective_date);
    return { ok: true, supplier_code, part_code };
  });

  app.get("/supplier-catalog", async (req, reply) => {
    const part_code = String(req.query?.part_code || "").trim();
    const supplier_code = String(req.query?.supplier_code || "").trim().toUpperCase();
    const where = [];
    const params = [];
    if (part_code) {
      where.push("p.part_code = ?");
      params.push(part_code);
    }
    if (supplier_code) {
      where.push("s.supplier_code = ?");
      params.push(supplier_code);
    }
    const rows = db.prepare(`
      SELECT
        c.id,
        s.supplier_code,
        s.supplier_name,
        p.part_code,
        p.part_name,
        c.supplier_part_code,
        c.lead_time_days,
        c.last_price,
        c.currency,
        c.effective_date,
        c.updated_at
      FROM supplier_part_catalog c
      JOIN suppliers s ON s.id = c.supplier_id
      JOIN parts p ON p.id = c.part_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY s.supplier_name ASC, p.part_code ASC
      LIMIT 500
    `).all(...params);
    return reply.send({ ok: true, rows });
  });

  app.post("/requisitions/:id/create-po", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const siteCode = getSiteCode(req);
    const reqn = db.prepare(`
      SELECT *
      FROM procurement_requisitions
      WHERE id = ? AND ${bySiteReq}
    `).get(id, siteCode);
    if (!reqn) return reply.code(404).send({ error: "requisition not found" });
    if (!["approved", "approved_all", "po_ready", "received", "partially_received"].includes(String(reqn.status || "").toLowerCase())) {
      return reply.code(409).send({ error: "requisition must be approved before PO creation" });
    }
    const existing = db.prepare(`SELECT id, po_number, status FROM procurement_purchase_orders WHERE requisition_id = ? ORDER BY id DESC LIMIT 1`).get(id);
    if (existing && String(existing.status || "").toLowerCase() !== "cancelled") {
      return reply.send({ ok: true, duplicate: true, po_id: Number(existing.id), po_number: String(existing.po_number || "") });
    }
    const lines = db.prepare(`
      SELECT *
      FROM procurement_requisition_lines
      WHERE requisition_id = ?
      ORDER BY line_no ASC
    `).all(id);
    if (!lines.length) return reply.code(409).send({ error: "cannot create PO without requisition lines" });
    const supplier = reqn.supplier_name
      ? db.prepare(`SELECT id, supplier_code, supplier_name, currency FROM suppliers WHERE LOWER(supplier_name) = LOWER(?) OR UPPER(supplier_code) = UPPER(?) LIMIT 1`).get(String(reqn.supplier_name), String(reqn.supplier_name))
      : null;
    const po_number = nextPONumber();
    const currency = supplier?.currency ? String(supplier.currency).toUpperCase() : "USD";
    const tx = db.transaction(() => {
      const head = db.prepare(`
        INSERT INTO procurement_purchase_orders (
          po_number, requisition_id, supplier_id, site_code, currency, status, notes, created_by, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, 'draft', ?, ?, datetime('now'), datetime('now'))
      `).run(po_number, id, supplier ? Number(supplier.id) : null, siteCode, currency, reqn.notes || null, getUser(req));
      const po_id = Number(head.lastInsertRowid);
      const insLine = db.prepare(`
        INSERT INTO procurement_purchase_order_lines (
          po_id, line_no, requisition_line_id, part_id, description, quantity_ordered, unit_price, line_total, needed_by_date, cost_center_code, labor_tag
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      let subtotal = 0;
      for (const l of lines) {
        const qty = Number(l.quantity || 0);
        const unit = Number(l.net_price ?? l.gross_price ?? 0);
        const line_total = Number((qty * unit).toFixed(2));
        subtotal += line_total;
        insLine.run(
          po_id,
          Number(l.line_no || 0),
          Number(l.id || 0),
          l.part_id ? Number(l.part_id) : null,
          l.description || null,
          qty,
          unit,
          line_total,
          l.needed_by_date || null,
          req.body?.cost_center_code ? String(req.body.cost_center_code).trim() : null,
          req.body?.labor_tag ? String(req.body.labor_tag).trim() : null
        );
      }
      db.prepare(`UPDATE procurement_purchase_orders SET subtotal = ?, updated_at = datetime('now') WHERE id = ?`).run(Number(subtotal.toFixed(2)), po_id);
      db.prepare(`
        UPDATE procurement_requisitions
        SET status = 'po_ready', po_number = ?, updated_at = datetime('now')
        WHERE id = ?
      `).run(po_number, id);
      return po_id;
    });
    const po_id = tx();
    return { ok: true, po_id, po_number };
  });

  app.get("/purchase-orders", async (req, reply) => {
    const status = String(req.query?.status || "").trim().toLowerCase();
    const siteCode = getSiteCode(req);
    const where = [`LOWER(TRIM(COALESCE(po.site_code, 'main'))) = ?`];
    const params = [siteCode];
    if (status) {
      where.push("LOWER(po.status) = ?");
      params.push(status);
    }
    const rows = db.prepare(`
      SELECT
        po.id,
        po.po_number,
        po.requisition_id,
        po.currency,
        po.status,
        po.subtotal,
        po.created_at,
        po.updated_at,
        s.supplier_code,
        s.supplier_name
      FROM procurement_purchase_orders po
      LEFT JOIN suppliers s ON s.id = po.supplier_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY po.id DESC
      LIMIT 300
    `).all(...params).map((r) => ({ ...r, subtotal: Number(r.subtotal || 0) }));
    return reply.send({ ok: true, rows });
  });

  app.get("/purchase-orders/:id/detail", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const siteCode = getSiteCode(req);
    const po = db.prepare(`
      SELECT po.*, s.supplier_code, s.supplier_name
      FROM procurement_purchase_orders po
      LEFT JOIN suppliers s ON s.id = po.supplier_id
      WHERE po.id = ? AND LOWER(TRIM(COALESCE(po.site_code, 'main'))) = ?
    `).get(id, siteCode);
    if (!po) return reply.code(404).send({ error: "PO not found" });
    const lines = db.prepare(`
      SELECT pol.*, p.part_code, p.part_name
      FROM procurement_purchase_order_lines pol
      LEFT JOIN parts p ON p.id = pol.part_id
      WHERE pol.po_id = ?
      ORDER BY pol.line_no ASC
    `).all(id);
    const receipts = db.prepare(`
      SELECT id, receipt_number, receipt_date, status, location_code, received_by, created_at
      FROM procurement_goods_receipts
      WHERE po_id = ?
      ORDER BY id DESC
    `).all(id);
    return reply.send({ ok: true, po, lines, receipts });
  });

  app.post("/purchase-orders/:id/approve", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const po = db.prepare(`SELECT id, status FROM procurement_purchase_orders WHERE id = ?`).get(id);
    if (!po) return reply.code(404).send({ error: "PO not found" });
    if (!["draft", "approved", "sent", "partially_received", "received"].includes(String(po.status || "").toLowerCase())) {
      return reply.code(409).send({ error: `cannot approve from status ${po.status}` });
    }
    db.prepare(`
      UPDATE procurement_purchase_orders
      SET status = 'approved', approved_at = datetime('now'), updated_at = datetime('now')
      WHERE id = ?
    `).run(id);
    return { ok: true, id, status: "approved" };
  });

  app.post("/purchase-orders/:id/send", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const po = db.prepare(`SELECT id, status FROM procurement_purchase_orders WHERE id = ?`).get(id);
    if (!po) return reply.code(404).send({ error: "PO not found" });
    if (!["approved", "sent", "partially_received", "received"].includes(String(po.status || "").toLowerCase())) {
      return reply.code(409).send({ error: `cannot send from status ${po.status}` });
    }
    db.prepare(`
      UPDATE procurement_purchase_orders
      SET status = CASE WHEN status = 'approved' THEN 'sent' ELSE status END,
          sent_at = COALESCE(sent_at, datetime('now')),
          updated_at = datetime('now')
      WHERE id = ?
    `).run(id);
    return { ok: true, id };
  });

  app.post("/purchase-orders/:id/receive", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores", "procurement"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const po = db.prepare(`
      SELECT id, po_number, requisition_id, status
      FROM procurement_purchase_orders
      WHERE id = ?
    `).get(id);
    if (!po) return reply.code(404).send({ error: "PO not found" });
    if (!["approved", "sent", "partially_received", "received"].includes(String(po.status || "").toLowerCase())) {
      return reply.code(409).send({ error: `cannot receive from status ${po.status}` });
    }
    const lines = Array.isArray(req.body?.lines) ? req.body.lines : [];
    if (!lines.length) return reply.code(400).send({ error: "lines array is required" });
    const location_code = String(req.body?.location_code || "MAIN").trim().toUpperCase();
    const location = getLocationByCode.get(location_code);
    if (!location) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    const receipt_date = req.body?.receipt_date ? String(req.body.receipt_date).trim() : new Date().toISOString().slice(0, 10);
    const receipt_number = String(req.body?.receipt_number || "").trim() || nextReceiptNumber();

    const tx = db.transaction(() => {
      const insReceipt = db.prepare(`
        INSERT INTO procurement_goods_receipts (
          po_id, receipt_number, receipt_date, received_by, location_code, status, notes
        ) VALUES (?, ?, ?, ?, ?, 'posted', ?)
      `).run(id, receipt_number, receipt_date, getUser(req), location_code, req.body?.notes ? String(req.body.notes).trim() : null);
      const receipt_id = Number(insReceipt.lastInsertRowid);
      const getPoLine = db.prepare(`
        SELECT id, line_no, part_id, quantity_ordered, quantity_received, unit_price
        FROM procurement_purchase_order_lines
        WHERE id = ? AND po_id = ?
      `);
      const insReceiptLine = db.prepare(`
        INSERT INTO procurement_goods_receipt_lines (
          receipt_id, po_line_id, part_id, quantity_received, unit_price, line_total, cost_center_code
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
      `);
      const updPoLine = db.prepare(`
        UPDATE procurement_purchase_order_lines
        SET quantity_received = quantity_received + ?
        WHERE id = ?
      `);
      const insMove = db.prepare(`
        INSERT INTO stock_movements (
          part_id, quantity, movement_type, reference, location_id
        ) VALUES (?, ?, 'in', ?, ?)
      `);

      for (const row of lines) {
        const po_line_id = Number(row?.po_line_id || 0);
        const qty = Number(row?.quantity_received || 0);
        if (!Number.isFinite(po_line_id) || po_line_id <= 0 || !Number.isFinite(qty) || qty <= 0) {
          throw new Error("each line requires po_line_id and quantity_received > 0");
        }
        const poLine = getPoLine.get(po_line_id, id);
        if (!poLine) throw new Error(`po_line_id ${po_line_id} not found for PO`);
        const outstanding = Number(poLine.quantity_ordered || 0) - Number(poLine.quantity_received || 0);
        if (qty > outstanding + 1e-9) {
          throw new Error(`line ${poLine.line_no}: received qty exceeds outstanding (${outstanding})`);
        }
        const unit_price = row?.unit_price != null && row.unit_price !== "" ? Number(row.unit_price) : Number(poLine.unit_price || 0);
        const line_total = Number((qty * unit_price).toFixed(2));
        insReceiptLine.run(
          receipt_id,
          po_line_id,
          poLine.part_id ? Number(poLine.part_id) : null,
          qty,
          unit_price,
          line_total,
          row?.cost_center_code ? String(row.cost_center_code).trim() : null
        );
        updPoLine.run(qty, po_line_id);
        if (poLine.part_id) {
          insMove.run(Number(poLine.part_id), qty, `po:${id}:receipt:${receipt_id}:line:${po_line_id}`, Number(location.id));
        }
      }

      const poStatus = derivePoStatus(id);
      db.prepare(`UPDATE procurement_purchase_orders SET status = ?, updated_at = datetime('now') WHERE id = ?`).run(poStatus, id);
      if (Number(po.requisition_id || 0) > 0) {
        const reqStatus = poStatus === "received" ? "received" : "partially_received";
        db.prepare(`UPDATE procurement_requisitions SET status = ?, updated_at = datetime('now') WHERE id = ?`).run(reqStatus, Number(po.requisition_id));
      }
      return { receipt_id, receipt_number, po_status: poStatus };
    });
    const result = tx();
    return { ok: true, ...result };
  });

  app.post("/purchase-orders/:id/invoices", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const po = db.prepare(`SELECT id, supplier_id, currency FROM procurement_purchase_orders WHERE id = ?`).get(id);
    if (!po) return reply.code(404).send({ error: "PO not found" });
    const invoice_number = String(req.body?.invoice_number || "").trim();
    const invoice_date = String(req.body?.invoice_date || "").trim() || new Date().toISOString().slice(0, 10);
    if (!invoice_number) return reply.code(400).send({ error: "invoice_number is required" });
    const lines = Array.isArray(req.body?.lines) ? req.body.lines : [];
    if (!lines.length) return reply.code(400).send({ error: "lines array is required" });
    const tx = db.transaction(() => {
      const head = db.prepare(`
        INSERT INTO procurement_invoices (
          po_id, invoice_number, supplier_id, invoice_date, status, currency, subtotal, tax, total, captured_by, notes
        ) VALUES (?, ?, ?, ?, 'captured', ?, 0, ?, 0, ?, ?)
      `).run(
        id,
        invoice_number,
        po.supplier_id ? Number(po.supplier_id) : null,
        invoice_date,
        String(req.body?.currency || po.currency || "USD").trim().toUpperCase(),
        Number(req.body?.tax || 0),
        getUser(req),
        req.body?.notes ? String(req.body.notes).trim() : null
      );
      const invoice_id = Number(head.lastInsertRowid);
      const insLine = db.prepare(`
        INSERT INTO procurement_invoice_lines (
          invoice_id, po_line_id, part_id, description, quantity_invoiced, unit_price, line_total, cost_center_code
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
      `);
      let subtotal = 0;
      for (const l of lines) {
        const po_line_id = Number(l?.po_line_id || 0);
        const qty = Number(l?.quantity_invoiced || 0);
        const unit_price = Number(l?.unit_price || 0);
        if (!Number.isFinite(po_line_id) || po_line_id <= 0 || !Number.isFinite(qty) || qty <= 0 || !Number.isFinite(unit_price) || unit_price < 0) {
          throw new Error("invoice lines require po_line_id, quantity_invoiced > 0, unit_price >= 0");
        }
        const poLine = db.prepare(`SELECT id, part_id, description FROM procurement_purchase_order_lines WHERE id = ? AND po_id = ?`).get(po_line_id, id);
        if (!poLine) throw new Error(`po_line_id ${po_line_id} not found for PO`);
        const line_total = Number((qty * unit_price).toFixed(2));
        subtotal += line_total;
        insLine.run(
          invoice_id,
          po_line_id,
          poLine.part_id ? Number(poLine.part_id) : null,
          l?.description ? String(l.description).trim() : (poLine.description || null),
          qty,
          unit_price,
          line_total,
          l?.cost_center_code ? String(l.cost_center_code).trim() : null
        );
      }
      const tax = Number(req.body?.tax || 0);
      const total = Number((subtotal + tax).toFixed(2));
      db.prepare(`UPDATE procurement_invoices SET subtotal = ?, total = ? WHERE id = ?`).run(Number(subtotal.toFixed(2)), total, invoice_id);
      return { invoice_id, invoice_number, subtotal: Number(subtotal.toFixed(2)), tax, total };
    });
    const result = tx();
    return { ok: true, ...result };
  });

  app.post("/purchase-orders/:id/three-way-match", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const po = db.prepare(`SELECT id FROM procurement_purchase_orders WHERE id = ?`).get(id);
    if (!po) return reply.code(404).send({ error: "PO not found" });
    const qtyTol = Math.max(0, Number(req.body?.quantity_tolerance || 0));
    const priceTolPct = Math.max(0, Number(req.body?.price_tolerance_pct || 0));
    const totalTol = Math.max(0, Number(req.body?.total_tolerance || 0));

    db.prepare(`DELETE FROM procurement_match_exceptions WHERE po_id = ? AND status = 'open'`).run(id);

    const poLines = db.prepare(`
      SELECT id, line_no, quantity_ordered, quantity_received, unit_price, line_total
      FROM procurement_purchase_order_lines
      WHERE po_id = ?
      ORDER BY line_no ASC
    `).all(id);
    const invoiceLines = db.prepare(`
      SELECT il.id, il.invoice_id, il.po_line_id, il.quantity_invoiced, il.unit_price, il.line_total
      FROM procurement_invoice_lines il
      JOIN procurement_invoices i ON i.id = il.invoice_id
      WHERE i.po_id = ?
    `).all(id);
    const receiptLines = db.prepare(`
      SELECT rl.id, rl.receipt_id, rl.po_line_id, rl.quantity_received, rl.unit_price, rl.line_total
      FROM procurement_goods_receipt_lines rl
      JOIN procurement_goods_receipts r ON r.id = rl.receipt_id
      WHERE r.po_id = ?
    `).all(id);
    const invByPoLine = invoiceLines.reduce((m, r) => {
      const k = Number(r.po_line_id || 0);
      const cur = m.get(k) || { qty: 0, total: 0, last_unit: 0, lines: [] };
      cur.qty += Number(r.quantity_invoiced || 0);
      cur.total += Number(r.line_total || 0);
      cur.last_unit = Number(r.unit_price || 0);
      cur.lines.push(r);
      m.set(k, cur);
      return m;
    }, new Map());
    const recByPoLine = receiptLines.reduce((m, r) => {
      const k = Number(r.po_line_id || 0);
      const cur = m.get(k) || { qty: 0, total: 0, lines: [] };
      cur.qty += Number(r.quantity_received || 0);
      cur.total += Number(r.line_total || 0);
      cur.lines.push(r);
      m.set(k, cur);
      return m;
    }, new Map());

    const insEx = db.prepare(`
      INSERT INTO procurement_match_exceptions (
        po_id, po_line_id, invoice_id, invoice_line_id, receipt_id, receipt_line_id,
        exception_type, severity, status, details_json, assigned_to
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'open', ?, ?)
    `);

    let exception_count = 0;
    for (const line of poLines) {
      const po_line_id = Number(line.id || 0);
      const poQty = Number(line.quantity_ordered || 0);
      const poUnit = Number(line.unit_price || 0);
      const poTotal = Number(line.line_total || 0);
      const inv = invByPoLine.get(po_line_id) || { qty: 0, total: 0, last_unit: 0, lines: [] };
      const rec = recByPoLine.get(po_line_id) || { qty: 0, total: 0, lines: [] };

      const qtyVsRec = Math.abs(poQty - rec.qty);
      const qtyVsInv = Math.abs(rec.qty - inv.qty);
      const priceVarPct = poUnit > 0 ? (Math.abs(inv.last_unit - poUnit) / poUnit) * 100 : 0;
      const totalVar = Math.abs(rec.total - inv.total);

      if (qtyVsRec > qtyTol) {
        insEx.run(id, po_line_id, null, null, rec.lines[0]?.receipt_id || null, rec.lines[0]?.id || null, "PO_vs_Receipt_qty", "high", JSON.stringify({ po_qty: poQty, receipt_qty: rec.qty, tolerance: qtyTol }), null);
        exception_count += 1;
      }
      if (qtyVsInv > qtyTol) {
        insEx.run(id, po_line_id, inv.lines[0]?.invoice_id || null, inv.lines[0]?.id || null, rec.lines[0]?.receipt_id || null, rec.lines[0]?.id || null, "Receipt_vs_Invoice_qty", "high", JSON.stringify({ receipt_qty: rec.qty, invoice_qty: inv.qty, tolerance: qtyTol }), null);
        exception_count += 1;
      }
      if (priceVarPct > priceTolPct) {
        insEx.run(id, po_line_id, inv.lines[0]?.invoice_id || null, inv.lines[0]?.id || null, null, null, "PO_vs_Invoice_unit_price", "warn", JSON.stringify({ po_unit_price: poUnit, invoice_unit_price: inv.last_unit, variance_pct: Number(priceVarPct.toFixed(2)), tolerance_pct: priceTolPct }), null);
        exception_count += 1;
      }
      if (totalVar > totalTol) {
        insEx.run(id, po_line_id, inv.lines[0]?.invoice_id || null, inv.lines[0]?.id || null, rec.lines[0]?.receipt_id || null, rec.lines[0]?.id || null, "Receipt_vs_Invoice_total", "warn", JSON.stringify({ receipt_total: Number(rec.total.toFixed(2)), invoice_total: Number(inv.total.toFixed(2)), variance_total: Number(totalVar.toFixed(2)), tolerance_total: totalTol, po_total: poTotal }), null);
        exception_count += 1;
      }
    }
    return { ok: true, po_id: id, exception_count };
  });

  app.get("/exceptions", async (req, reply) => {
    const status = String(req.query?.status || "open").trim().toLowerCase();
    const rows = db.prepare(`
      SELECT
        e.*,
        po.po_number
      FROM procurement_match_exceptions e
      LEFT JOIN procurement_purchase_orders po ON po.id = e.po_id
      WHERE LOWER(e.status) = ?
      ORDER BY e.id DESC
      LIMIT 500
    `).all(status).map((r) => ({
      ...r,
      details: (() => {
        try { return JSON.parse(String(r.details_json || "{}")); } catch { return {}; }
      })(),
    }));
    return reply.send({ ok: true, rows });
  });

  app.post("/exceptions/:id/resolve", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const ex = db.prepare(`SELECT id, status FROM procurement_match_exceptions WHERE id = ?`).get(id);
    if (!ex) return reply.code(404).send({ error: "exception not found" });
    if (String(ex.status || "").toLowerCase() === "resolved") return { ok: true, duplicate: true, id };
    db.prepare(`
      UPDATE procurement_match_exceptions
      SET status = 'resolved', resolved_by = ?, resolved_at = datetime('now')
      WHERE id = ?
    `).run(getUser(req), id);
    return { ok: true, id, status: "resolved" };
  });

  app.post("/journals/build", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement", "stores"])) return;
    const start = String(req.body?.start || "").trim();
    const end = String(req.body?.end || "").trim();
    if (!start || !end) return reply.code(400).send({ error: "start and end are required" });
    const batch_id = String(req.body?.batch_id || "").trim() || `JRN-${new Date().toISOString().slice(0, 10)}-${Date.now()}`;
    const defaultCostCenter = req.body?.default_cost_center_code ? String(req.body.default_cost_center_code).trim() : null;
    const tx = db.transaction(() => {
      db.prepare(`DELETE FROM finance_journal_staging WHERE batch_id = ?`).run(batch_id);
      const ins = db.prepare(`
        INSERT INTO finance_journal_staging (
          batch_id, tx_date, source_module, source_type, source_id,
          account_code, cost_center_code, description, debit, credit, currency
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      const receiptRows = db.prepare(`
        SELECT
          r.receipt_date AS tx_date,
          rl.id AS receipt_line_id,
          rl.line_total,
          COALESCE(NULLIF(TRIM(rl.cost_center_code), ''), NULLIF(TRIM(pol.cost_center_code), ''), ?) AS cost_center_code,
          po.currency
        FROM procurement_goods_receipt_lines rl
        JOIN procurement_goods_receipts r ON r.id = rl.receipt_id
        JOIN procurement_purchase_orders po ON po.id = r.po_id
        LEFT JOIN procurement_purchase_order_lines pol ON pol.id = rl.po_line_id
        WHERE DATE(r.receipt_date) BETWEEN DATE(?) AND DATE(?)
      `).all(defaultCostCenter, start, end);
      for (const row of receiptRows) {
        const amount = Number(row.line_total || 0);
        if (amount <= 0) continue;
        const tx_date = String(row.tx_date || start);
        const cc = row.cost_center_code ? String(row.cost_center_code) : null;
        const cur = String(row.currency || "USD");
        ins.run(batch_id, tx_date, "procurement", "goods_receipt", String(row.receipt_line_id), "1400-INVENTORY", cc, `GRN line ${row.receipt_line_id}`, amount, 0, cur);
        ins.run(batch_id, tx_date, "procurement", "goods_receipt", String(row.receipt_line_id), "2100-GRNI", cc, `GRN accrual ${row.receipt_line_id}`, 0, amount, cur);
      }

      const invoiceRows = db.prepare(`
        SELECT
          i.invoice_date AS tx_date,
          il.id AS invoice_line_id,
          il.line_total,
          COALESCE(NULLIF(TRIM(il.cost_center_code), ''), NULLIF(TRIM(pol.cost_center_code), ''), ?) AS cost_center_code,
          i.currency
        FROM procurement_invoice_lines il
        JOIN procurement_invoices i ON i.id = il.invoice_id
        LEFT JOIN procurement_purchase_order_lines pol ON pol.id = il.po_line_id
        WHERE DATE(i.invoice_date) BETWEEN DATE(?) AND DATE(?)
      `).all(defaultCostCenter, start, end);
      for (const row of invoiceRows) {
        const amount = Number(row.line_total || 0);
        if (amount <= 0) continue;
        const tx_date = String(row.tx_date || start);
        const cc = row.cost_center_code ? String(row.cost_center_code) : null;
        const cur = String(row.currency || "USD");
        ins.run(batch_id, tx_date, "procurement", "invoice", String(row.invoice_line_id), "2100-GRNI", cc, `Invoice clear GRNI ${row.invoice_line_id}`, amount, 0, cur);
        ins.run(batch_id, tx_date, "procurement", "invoice", String(row.invoice_line_id), "2200-AP", cc, `AP recognition ${row.invoice_line_id}`, 0, amount, cur);
      }
    });
    tx();
    const summary = db.prepare(`
      SELECT
        COUNT(*) AS lines,
        COALESCE(SUM(debit), 0) AS debit_total,
        COALESCE(SUM(credit), 0) AS credit_total
      FROM finance_journal_staging
      WHERE batch_id = ?
    `).get(batch_id);
    return {
      ok: true,
      batch_id,
      lines: Number(summary?.lines || 0),
      debit_total: Number(Number(summary?.debit_total || 0).toFixed(2)),
      credit_total: Number(Number(summary?.credit_total || 0).toFixed(2)),
      balanced: Number(Number(summary?.debit_total || 0).toFixed(2)) === Number(Number(summary?.credit_total || 0).toFixed(2)),
    };
  });

  app.get("/journals/export.csv", async (req, reply) => {
    const batch_id = String(req.query?.batch_id || "").trim();
    if (!batch_id) return reply.code(400).send({ error: "batch_id is required" });
    const rows = db.prepare(`
      SELECT
        batch_id, tx_date, source_module, source_type, source_id,
        account_code, cost_center_code, description, debit, credit, currency
      FROM finance_journal_staging
      WHERE batch_id = ?
      ORDER BY id ASC
    `).all(batch_id);
    if (!rows.length) return reply.code(404).send({ error: "no journal lines for batch_id" });
    const esc = (v) => {
      const s = String(v ?? "");
      if (/[",\n]/.test(s)) return `"${s.replace(/"/g, "\"\"")}"`;
      return s;
    };
    const header = ["batch_id", "tx_date", "source_module", "source_type", "source_id", "account_code", "cost_center_code", "description", "debit", "credit", "currency"];
    const csv = [header.join(",")]
      .concat(rows.map((r) => [
        r.batch_id,
        r.tx_date,
        r.source_module,
        r.source_type,
        r.source_id,
        r.account_code,
        r.cost_center_code || "",
        r.description || "",
        Number(r.debit || 0).toFixed(2),
        Number(r.credit || 0).toFixed(2),
        r.currency || "USD",
      ].map(esc).join(",")))
      .join("\n");
    reply.header("Content-Type", "text/csv; charset=utf-8");
    reply.header("Content-Disposition", `attachment; filename="journal-${batch_id}.csv"`);
    return reply.send(csv);
  });

  app.get("/journals/export.xlsx", async (req, reply) => {
    const batch_id = String(req.query?.batch_id || "").trim();
    if (!batch_id) return reply.code(400).send({ error: "batch_id is required" });
    const rows = db.prepare(`
      SELECT
        batch_id, tx_date, source_module, source_type, source_id,
        account_code, cost_center_code, description, debit, credit, currency
      FROM finance_journal_staging
      WHERE batch_id = ?
      ORDER BY id ASC
    `).all(batch_id);
    if (!rows.length) return reply.code(404).send({ error: "no journal lines for batch_id" });
    const ExcelJS = (await import("exceljs")).default;
    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();
    const ws = wb.addWorksheet("Journals");
    ws.columns = [
      { header: "Batch ID", key: "batch_id", width: 20 },
      { header: "Date", key: "tx_date", width: 14 },
      { header: "Source Module", key: "source_module", width: 18 },
      { header: "Source Type", key: "source_type", width: 18 },
      { header: "Source ID", key: "source_id", width: 14 },
      { header: "Account", key: "account_code", width: 18 },
      { header: "Cost Center", key: "cost_center_code", width: 16 },
      { header: "Description", key: "description", width: 40 },
      { header: "Debit", key: "debit", width: 14 },
      { header: "Credit", key: "credit", width: 14 },
      { header: "Currency", key: "currency", width: 10 },
    ];
    ws.addRows(rows.map((r) => ({
      ...r,
      debit: Number(r.debit || 0),
      credit: Number(r.credit || 0),
    })));
    ws.getRow(1).font = { bold: true };
    ws.getColumn("debit").numFmt = "#,##0.00";
    ws.getColumn("credit").numFmt = "#,##0.00";
    const buffer = await wb.xlsx.writeBuffer();
    reply.header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    reply.header("Content-Disposition", `attachment; filename="journal-${batch_id}.xlsx"`);
    return reply.send(Buffer.from(buffer));
  });

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

  app.post("/journals/summarize", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement", "stores"])) return;
    const start = String(req.body?.start || "").trim();
    const end = String(req.body?.end || "").trim();
    if (!start || !end) return reply.code(400).send({ error: "start and end are required (YYYY-MM-DD)" });

    const categoriesInput = Array.isArray(req.body?.categories) ? req.body.categories : null;
    const categorySet = new Set(
      (categoriesInput || ["parts", "labor", "downtime", "fuel", "lube", "procurement_grn", "procurement_ap"])
        .map((x) => String(x || "").trim().toLowerCase())
        .filter(Boolean)
    );

    const defaultCostCenter = req.body?.default_cost_center_code ? String(req.body.default_cost_center_code).trim() : null;
    const currency = String(req.body?.currency || "USD").trim().toUpperCase() || "USD";
    const notes = req.body?.notes ? String(req.body.notes).trim() : null;
    const period = periodForDate(start);

    if (isPeriodLocked(period)) {
      return reply.code(409).send({ error: `period ${period} is locked` });
    }

    const fuelDefault = readCostSetting("fuel_cost_per_liter_default", 1.5);
    const lubeDefault = readCostSetting("lube_cost_per_qty_default", 4.0);
    const laborDefault = readCostSetting("labor_cost_per_hour_default", 35.0);
    const downtimeDefault = readCostSetting("downtime_cost_per_hour_default", 120.0);

    const assetsHasCC = tableHasColumn("assets", "cost_center_code");
    const assetsHasSite = tableHasColumn("assets", "site_code");

    const run_number = nextRunNumber(period);
    const user = getUser(req);

    const insRun = db.prepare(`
      INSERT INTO finance_posting_runs
        (run_number, period, start_date, end_date, run_type, status, currency, notes, created_by)
      VALUES (?, ?, ?, ?, 'summary', 'draft', ?, ?, ?)
    `);
    const insLine = db.prepare(`
      INSERT INTO finance_posting_run_lines
        (run_id, category, tx_date, source_module, source_type, source_ref,
         account_code, cost_center_code, site_code, equipment_type, asset_code,
         description, debit, credit, currency, qty, unit_cost)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    const tx = db.transaction(() => {
      const r = insRun.run(run_number, period, start, end, currency, notes, user);
      const runId = Number(r.lastInsertRowid);

      if (categorySet.has("parts")) {
        const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
        const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
        const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
        const smHasCC = smCols.some((c) => String(c.name) === "cost_center_code");
        const rows = db.prepare(`
          SELECT
            ${smDateExpr} AS tx_date,
            COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
            COALESCE(a.category, '') AS equipment_type,
            ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
            ${smHasCC
              ? `COALESCE(NULLIF(TRIM(sm.cost_center_code), ''),
                         ${assetsHasCC ? "NULLIF(TRIM(a.cost_center_code), '')," : ""}
                         ?)`
              : `COALESCE(${assetsHasCC ? "NULLIF(TRIM(a.cost_center_code), ''), " : ""}?)`} AS cost_center_code,
            SUM(ABS(sm.quantity)) AS qty,
            SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)) AS amount
          FROM stock_movements sm
          JOIN parts p ON p.id = sm.part_id
          LEFT JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
          LEFT JOIN assets a ON a.id = w.asset_id
          WHERE sm.movement_type = 'out'
            AND ${smDateExpr} BETWEEN DATE(?) AND DATE(?)
          GROUP BY tx_date, asset_code, cost_center_code, equipment_type, site_code
        `).all(defaultCostCenter, start, end);
        for (const row of rows) {
          const amount = Number(row.amount || 0);
          if (amount <= 0) continue;
          const cc = row.cost_center_code || null;
          const site = row.site_code || null;
          const eq = row.equipment_type || null;
          const ref = `parts:${row.tx_date}:${row.asset_code}`;
          insLine.run(runId, "parts", row.tx_date, "maintenance", "parts_issue", ref,
            "5200-PARTS-EXPENSE", cc, site, eq, row.asset_code,
            `Parts issued ${row.asset_code} on ${row.tx_date}`, amount, 0, currency, Number(row.qty || 0), null);
          insLine.run(runId, "parts", row.tx_date, "maintenance", "parts_issue", ref,
            "1400-INVENTORY", cc, site, eq, row.asset_code,
            `Parts inventory reduction ${row.asset_code}`, 0, amount, currency, Number(row.qty || 0), null);
        }
      }

      if (categorySet.has("labor")) {
        const woCols = db.prepare(`PRAGMA table_info(work_orders)`).all();
        const woHasCompleted = woCols.some((c) => String(c.name) === "completed_at");
        const woHasClosed = woCols.some((c) => String(c.name) === "closed_at");
        const woHasLaborHours = woCols.some((c) => String(c.name) === "labor_hours");
        const woHasLaborRate = woCols.some((c) => String(c.name) === "labor_rate_per_hour");
        const woHasCC = woCols.some((c) => String(c.name) === "cost_center_code");
        if (woHasLaborHours) {
          const dateExpr = woHasCompleted && woHasClosed
            ? "DATE(COALESCE(w.completed_at, w.closed_at))"
            : (woHasCompleted ? "DATE(w.completed_at)" : "DATE(w.closed_at)");
          const rows = db.prepare(`
            SELECT
              ${dateExpr} AS tx_date,
              COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
              COALESCE(a.category, '') AS equipment_type,
              ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
              ${woHasCC
                ? `COALESCE(NULLIF(TRIM(w.cost_center_code), ''),
                           ${assetsHasCC ? "NULLIF(TRIM(a.cost_center_code), ''), " : ""} ?)`
                : `COALESCE(${assetsHasCC ? "NULLIF(TRIM(a.cost_center_code), ''), " : ""} ?)`} AS cost_center_code,
              SUM(COALESCE(w.labor_hours, 0)) AS hours,
              SUM(COALESCE(w.labor_hours, 0) * COALESCE(${woHasLaborRate ? "w.labor_rate_per_hour" : "NULL"}, ?)) AS amount
            FROM work_orders w
            LEFT JOIN assets a ON a.id = w.asset_id
            WHERE w.status IN ('completed','approved','closed')
              AND ${dateExpr} BETWEEN DATE(?) AND DATE(?)
            GROUP BY tx_date, asset_code, cost_center_code, equipment_type, site_code
          `).all(defaultCostCenter, laborDefault, start, end);
          for (const row of rows) {
            const amount = Number(row.amount || 0);
            if (amount <= 0) continue;
            const cc = row.cost_center_code || null;
            const site = row.site_code || null;
            const eq = row.equipment_type || null;
            const ref = `labor:${row.tx_date}:${row.asset_code}`;
            insLine.run(runId, "labor", row.tx_date, "maintenance", "labor", ref,
              "5300-LABOR-EXPENSE", cc, site, eq, row.asset_code,
              `Labor ${row.asset_code} ${Number(row.hours).toFixed(2)}h`, amount, 0, currency, Number(row.hours || 0), null);
            insLine.run(runId, "labor", row.tx_date, "maintenance", "labor", ref,
              "2500-LABOR-CLEARING", cc, site, eq, row.asset_code,
              `Labor clearing ${row.asset_code}`, 0, amount, currency, Number(row.hours || 0), null);
          }
        }
      }

      if (categorySet.has("downtime")) {
        const rows = db.prepare(`
          SELECT
            l.log_date AS tx_date,
            COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
            COALESCE(a.category, '') AS equipment_type,
            ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
            ${assetsHasCC ? "COALESCE(NULLIF(TRIM(a.cost_center_code), ''), ?)" : "?"} AS cost_center_code,
            SUM(COALESCE(l.hours_down, 0)) AS hours,
            SUM(COALESCE(l.hours_down, 0) * COALESCE(a.downtime_cost_per_hour, ?)) AS amount
          FROM breakdown_downtime_logs l
          JOIN breakdowns b ON b.id = l.breakdown_id
          JOIN assets a ON a.id = b.asset_id
          WHERE l.log_date BETWEEN DATE(?) AND DATE(?)
          GROUP BY tx_date, asset_code, cost_center_code, equipment_type, site_code
        `).all(defaultCostCenter, downtimeDefault, start, end);
        for (const row of rows) {
          const amount = Number(row.amount || 0);
          if (amount <= 0) continue;
          const cc = row.cost_center_code || null;
          const site = row.site_code || null;
          const eq = row.equipment_type || null;
          const ref = `downtime:${row.tx_date}:${row.asset_code}`;
          insLine.run(runId, "downtime", row.tx_date, "maintenance", "downtime", ref,
            "5400-DOWNTIME-EXPENSE", cc, site, eq, row.asset_code,
            `Downtime ${row.asset_code} ${Number(row.hours).toFixed(2)}h`, amount, 0, currency, Number(row.hours || 0), null);
          insLine.run(runId, "downtime", row.tx_date, "maintenance", "downtime", ref,
            "2600-DOWNTIME-CLEARING", cc, site, eq, row.asset_code,
            `Downtime clearing ${row.asset_code}`, 0, amount, currency, Number(row.hours || 0), null);
        }
      }

      if (categorySet.has("fuel")) {
        const flCols = db.prepare(`PRAGMA table_info(fuel_logs)`).all();
        const flHasUnit = flCols.some((c) => String(c.name) === "unit_cost_per_liter");
        const rows = db.prepare(`
          SELECT
            fl.log_date AS tx_date,
            COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
            COALESCE(a.category, '') AS equipment_type,
            ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
            ${assetsHasCC ? "COALESCE(NULLIF(TRIM(a.cost_center_code), ''), ?)" : "?"} AS cost_center_code,
            SUM(COALESCE(fl.liters, 0)) AS qty,
            SUM(COALESCE(fl.liters, 0) * COALESCE(${flHasUnit ? "fl.unit_cost_per_liter" : "NULL"}, a.fuel_cost_per_liter, ?)) AS amount
          FROM fuel_logs fl
          JOIN assets a ON a.id = fl.asset_id
          WHERE fl.log_date BETWEEN DATE(?) AND DATE(?)
          GROUP BY tx_date, asset_code, cost_center_code, equipment_type, site_code
        `).all(defaultCostCenter, fuelDefault, start, end);
        for (const row of rows) {
          const amount = Number(row.amount || 0);
          if (amount <= 0) continue;
          const cc = row.cost_center_code || null;
          const site = row.site_code || null;
          const eq = row.equipment_type || null;
          const ref = `fuel:${row.tx_date}:${row.asset_code}`;
          insLine.run(runId, "fuel", row.tx_date, "operations", "fuel", ref,
            "5100-FUEL-EXPENSE", cc, site, eq, row.asset_code,
            `Fuel ${row.asset_code} ${Number(row.qty).toFixed(2)}L`, amount, 0, currency, Number(row.qty || 0), null);
          insLine.run(runId, "fuel", row.tx_date, "operations", "fuel", ref,
            "1410-FUEL-INVENTORY", cc, site, eq, row.asset_code,
            `Fuel inventory reduction ${row.asset_code}`, 0, amount, currency, Number(row.qty || 0), null);
        }
      }

      if (categorySet.has("lube")) {
        const olCols = db.prepare(`PRAGMA table_info(oil_logs)`).all();
        const olHasUnit = olCols.some((c) => String(c.name) === "unit_cost");
        const rows = db.prepare(`
          SELECT
            ol.log_date AS tx_date,
            COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
            COALESCE(a.category, '') AS equipment_type,
            ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
            ${assetsHasCC ? "COALESCE(NULLIF(TRIM(a.cost_center_code), ''), ?)" : "?"} AS cost_center_code,
            SUM(COALESCE(ol.quantity, 0)) AS qty,
            SUM(COALESCE(ol.quantity, 0) * COALESCE(${olHasUnit ? "ol.unit_cost" : "NULL"}, ?)) AS amount
          FROM oil_logs ol
          JOIN assets a ON a.id = ol.asset_id
          WHERE ol.log_date BETWEEN DATE(?) AND DATE(?)
          GROUP BY tx_date, asset_code, cost_center_code, equipment_type, site_code
        `).all(defaultCostCenter, lubeDefault, start, end);
        for (const row of rows) {
          const amount = Number(row.amount || 0);
          if (amount <= 0) continue;
          const cc = row.cost_center_code || null;
          const site = row.site_code || null;
          const eq = row.equipment_type || null;
          const ref = `lube:${row.tx_date}:${row.asset_code}`;
          insLine.run(runId, "lube", row.tx_date, "operations", "lube", ref,
            "5150-LUBE-EXPENSE", cc, site, eq, row.asset_code,
            `Lube ${row.asset_code} ${Number(row.qty).toFixed(2)}`, amount, 0, currency, Number(row.qty || 0), null);
          insLine.run(runId, "lube", row.tx_date, "operations", "lube", ref,
            "1420-LUBE-INVENTORY", cc, site, eq, row.asset_code,
            `Lube inventory reduction ${row.asset_code}`, 0, amount, currency, Number(row.qty || 0), null);
        }
      }

      if (categorySet.has("procurement_grn")) {
        const rows = db.prepare(`
          SELECT
            r.receipt_date AS tx_date,
            rl.id AS ref_id,
            rl.line_total AS amount,
            COALESCE(NULLIF(TRIM(rl.cost_center_code), ''), NULLIF(TRIM(pol.cost_center_code), ''), ?) AS cost_center_code,
            po.currency
          FROM procurement_goods_receipt_lines rl
          JOIN procurement_goods_receipts r ON r.id = rl.receipt_id
          JOIN procurement_purchase_orders po ON po.id = r.po_id
          LEFT JOIN procurement_purchase_order_lines pol ON pol.id = rl.po_line_id
          WHERE DATE(r.receipt_date) BETWEEN DATE(?) AND DATE(?)
        `).all(defaultCostCenter, start, end);
        for (const row of rows) {
          const amount = Number(row.amount || 0);
          if (amount <= 0) continue;
          const cur = String(row.currency || currency);
          insLine.run(runId, "procurement_grn", row.tx_date, "procurement", "goods_receipt", `grn:${row.ref_id}`,
            "1400-INVENTORY", row.cost_center_code || null, null, null, null,
            `GRN line ${row.ref_id}`, amount, 0, cur, null, null);
          insLine.run(runId, "procurement_grn", row.tx_date, "procurement", "goods_receipt", `grn:${row.ref_id}`,
            "2100-GRNI", row.cost_center_code || null, null, null, null,
            `GRN accrual ${row.ref_id}`, 0, amount, cur, null, null);
        }
      }

      if (categorySet.has("procurement_ap")) {
        const rows = db.prepare(`
          SELECT
            i.invoice_date AS tx_date,
            il.id AS ref_id,
            il.line_total AS amount,
            COALESCE(NULLIF(TRIM(il.cost_center_code), ''), NULLIF(TRIM(pol.cost_center_code), ''), ?) AS cost_center_code,
            i.currency
          FROM procurement_invoice_lines il
          JOIN procurement_invoices i ON i.id = il.invoice_id
          LEFT JOIN procurement_purchase_order_lines pol ON pol.id = il.po_line_id
          WHERE DATE(i.invoice_date) BETWEEN DATE(?) AND DATE(?)
        `).all(defaultCostCenter, start, end);
        for (const row of rows) {
          const amount = Number(row.amount || 0);
          if (amount <= 0) continue;
          const cur = String(row.currency || currency);
          insLine.run(runId, "procurement_ap", row.tx_date, "procurement", "invoice", `inv:${row.ref_id}`,
            "2100-GRNI", row.cost_center_code || null, null, null, null,
            `Invoice clear GRNI ${row.ref_id}`, amount, 0, cur, null, null);
          insLine.run(runId, "procurement_ap", row.tx_date, "procurement", "invoice", `inv:${row.ref_id}`,
            "2200-AP", row.cost_center_code || null, null, null, null,
            `AP recognition ${row.ref_id}`, 0, amount, cur, null, null);
        }
      }

      recalcRunTotals(runId);
      try {
        writeAudit(db, req, {
          module: "finance",
          action: "journals.summarize",
          entity_type: "finance_posting_runs",
          entity_id: runId,
          payload: { run_number, period, start, end, categories: Array.from(categorySet) }
        });
      } catch {}
      return runId;
    });

    let runId;
    try { runId = tx(); }
    catch (e) { return reply.code(500).send({ error: `summarize failed: ${e.message}` }); }

    const header = db.prepare(`SELECT * FROM finance_posting_runs WHERE id = ?`).get(runId);
    return reply.send({
      ok: true,
      run: header,
      balanced: Number(Number(header?.total_debit || 0).toFixed(2)) === Number(Number(header?.total_credit || 0).toFixed(2))
    });
  });

  app.get("/journals/runs", async (req) => {
    const status = req.query?.status ? String(req.query.status).trim().toLowerCase() : null;
    const period = req.query?.period ? String(req.query.period).trim() : null;
    const where = [];
    const args = [];
    if (status) { where.push(`LOWER(status) = ?`); args.push(status); }
    if (period) { where.push(`period = ?`); args.push(period); }
    const sql = `
      SELECT id, run_number, period, start_date, end_date, run_type, status,
             currency, total_debit, total_credit, line_count, created_by,
             created_at, exported_at, posted_at, reversed_at
      FROM finance_posting_runs
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY id DESC LIMIT 500
    `;
    const rows = db.prepare(sql).all(...args);
    return { ok: true, rows };
  });

  app.get("/journals/runs/:id", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const run = db.prepare(`SELECT * FROM finance_posting_runs WHERE id = ?`).get(id);
    if (!run) return reply.code(404).send({ error: "run not found" });
    const byCategory = db.prepare(`
      SELECT category,
             COUNT(*) AS lines,
             COALESCE(SUM(debit), 0) AS debit_total,
             COALESCE(SUM(credit), 0) AS credit_total
      FROM finance_posting_run_lines
      WHERE run_id = ?
      GROUP BY category
      ORDER BY category
    `).all(id);
    return { ok: true, run, by_category: byCategory };
  });

  app.get("/journals/runs/:id/lines", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const category = req.query?.category ? String(req.query.category).trim().toLowerCase() : null;
    const where = [`run_id = ?`];
    const args = [id];
    if (category) { where.push(`LOWER(category) = ?`); args.push(category); }
    const rows = db.prepare(`
      SELECT id, category, tx_date, source_module, source_type, source_ref,
             account_code, cost_center_code, site_code, equipment_type, asset_code,
             description, debit, credit, currency, qty, unit_cost
      FROM finance_posting_run_lines
      WHERE ${where.join(" AND ")}
      ORDER BY tx_date ASC, id ASC
      LIMIT 5000
    `).all(...args);
    return { ok: true, rows };
  });

  app.post("/journals/runs/:id/mark-exported", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const run = db.prepare(`SELECT id, status FROM finance_posting_runs WHERE id = ?`).get(id);
    if (!run) return reply.code(404).send({ error: "run not found" });
    if (String(run.status).toLowerCase() === "posted") return reply.code(409).send({ error: "run already posted" });
    db.prepare(`
      UPDATE finance_posting_runs
      SET status = 'exported', exported_at = datetime('now'), exported_by = ?
      WHERE id = ?
    `).run(getUser(req), id);
    try {
      writeAudit(db, req, { module: "finance", action: "journals.mark_exported", entity_type: "finance_posting_runs", entity_id: id, payload: {} });
    } catch {}
    return { ok: true, id, status: "exported" };
  });

  app.post("/journals/runs/:id/mark-posted", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const run = db.prepare(`SELECT id, status, period FROM finance_posting_runs WHERE id = ?`).get(id);
    if (!run) return reply.code(404).send({ error: "run not found" });
    if (String(run.status).toLowerCase() === "posted") return { ok: true, id, status: "posted", duplicate: true };
    if (String(run.status).toLowerCase() === "reversed") return reply.code(409).send({ error: "run reversed" });
    if (isPeriodLocked(run.period)) return reply.code(409).send({ error: `period ${run.period} is locked` });
    const ref = req.body?.posted_reference ? String(req.body.posted_reference).trim() : null;
    db.prepare(`
      UPDATE finance_posting_runs
      SET status = 'posted', posted_at = datetime('now'), posted_by = ?, posted_reference = ?
      WHERE id = ?
    `).run(getUser(req), ref, id);
    try {
      writeAudit(db, req, { module: "finance", action: "journals.mark_posted", entity_type: "finance_posting_runs", entity_id: id, payload: { posted_reference: ref } });
    } catch {}
    return { ok: true, id, status: "posted", posted_reference: ref };
  });

  app.post("/journals/runs/:id/reverse", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const reason = req.body?.reason ? String(req.body.reason).trim() : "";
    if (!reason) return reply.code(400).send({ error: "reason is required" });
    const run = db.prepare(`SELECT id, status, period FROM finance_posting_runs WHERE id = ?`).get(id);
    if (!run) return reply.code(404).send({ error: "run not found" });
    if (String(run.status).toLowerCase() === "reversed") return { ok: true, id, status: "reversed", duplicate: true };
    if (isPeriodLocked(run.period)) return reply.code(409).send({ error: `period ${run.period} is locked` });
    db.prepare(`
      UPDATE finance_posting_runs
      SET status = 'reversed', reversed_at = datetime('now'), reversed_by = ?, reversed_reason = ?
      WHERE id = ?
    `).run(getUser(req), reason, id);
    try {
      writeAudit(db, req, { module: "finance", action: "journals.reverse", entity_type: "finance_posting_runs", entity_id: id, payload: { reason } });
    } catch {}
    return { ok: true, id, status: "reversed" };
  });

  app.get("/journals/runs/:id/export.csv", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const run = db.prepare(`SELECT * FROM finance_posting_runs WHERE id = ?`).get(id);
    if (!run) return reply.code(404).send({ error: "run not found" });
    const rows = db.prepare(`
      SELECT
        ? AS run_number, category, tx_date, source_module, source_type, source_ref,
        account_code, cost_center_code, site_code, equipment_type, asset_code,
        description, debit, credit, currency, qty, unit_cost
      FROM finance_posting_run_lines
      WHERE run_id = ?
      ORDER BY tx_date ASC, id ASC
    `).all(run.run_number, id);
    if (!rows.length) return reply.code(404).send({ error: "no lines" });
    const esc = (v) => {
      const s = String(v ?? "");
      if (/[",\n]/.test(s)) return `"${s.replace(/"/g, "\"\"")}"`;
      return s;
    };
    const header = ["run_number", "category", "tx_date", "source_module", "source_type", "source_ref",
      "account_code", "cost_center_code", "site_code", "equipment_type", "asset_code",
      "description", "debit", "credit", "currency", "qty", "unit_cost"];
    const csv = [header.join(",")]
      .concat(rows.map((r) => [
        r.run_number, r.category, r.tx_date, r.source_module, r.source_type, r.source_ref || "",
        r.account_code, r.cost_center_code || "", r.site_code || "", r.equipment_type || "", r.asset_code || "",
        r.description || "",
        Number(r.debit || 0).toFixed(2), Number(r.credit || 0).toFixed(2),
        r.currency || "USD",
        r.qty == null ? "" : Number(r.qty).toFixed(4),
        r.unit_cost == null ? "" : Number(r.unit_cost).toFixed(4),
      ].map(esc).join(",")))
      .join("\n");
    reply.header("Content-Type", "text/csv; charset=utf-8");
    reply.header("Content-Disposition", `attachment; filename="${run.run_number}.csv"`);
    return reply.send(csv);
  });

  app.get("/journals/runs/:id/export.xlsx", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const run = db.prepare(`SELECT * FROM finance_posting_runs WHERE id = ?`).get(id);
    if (!run) return reply.code(404).send({ error: "run not found" });
    const rows = db.prepare(`
      SELECT category, tx_date, source_module, source_type, source_ref,
             account_code, cost_center_code, site_code, equipment_type, asset_code,
             description, debit, credit, currency, qty, unit_cost
      FROM finance_posting_run_lines
      WHERE run_id = ?
      ORDER BY tx_date ASC, id ASC
    `).all(id);
    if (!rows.length) return reply.code(404).send({ error: "no lines" });
    const ExcelJS = (await import("exceljs")).default;
    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();
    const ws = wb.addWorksheet("PostingRun");
    ws.columns = [
      { header: "Run Number", key: "run_number", width: 22 },
      { header: "Category", key: "category", width: 16 },
      { header: "Date", key: "tx_date", width: 14 },
      { header: "Source Module", key: "source_module", width: 18 },
      { header: "Source Type", key: "source_type", width: 18 },
      { header: "Source Ref", key: "source_ref", width: 28 },
      { header: "Account", key: "account_code", width: 20 },
      { header: "Cost Center", key: "cost_center_code", width: 16 },
      { header: "Site", key: "site_code", width: 14 },
      { header: "Equipment Type", key: "equipment_type", width: 18 },
      { header: "Asset", key: "asset_code", width: 16 },
      { header: "Description", key: "description", width: 40 },
      { header: "Debit", key: "debit", width: 14 },
      { header: "Credit", key: "credit", width: 14 },
      { header: "Currency", key: "currency", width: 10 },
      { header: "Qty", key: "qty", width: 12 },
      { header: "Unit Cost", key: "unit_cost", width: 14 },
    ];
    ws.addRows(rows.map((r) => ({
      ...r,
      run_number: run.run_number,
      debit: Number(r.debit || 0),
      credit: Number(r.credit || 0),
      qty: r.qty == null ? null : Number(r.qty),
      unit_cost: r.unit_cost == null ? null : Number(r.unit_cost),
    })));
    ws.getRow(1).font = { bold: true };
    ws.getColumn("debit").numFmt = "#,##0.00";
    ws.getColumn("credit").numFmt = "#,##0.00";
    ws.getColumn("qty").numFmt = "#,##0.0000";
    ws.getColumn("unit_cost").numFmt = "#,##0.0000";
    const buffer = await wb.xlsx.writeBuffer();
    reply.header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    reply.header("Content-Disposition", `attachment; filename="${run.run_number}.xlsx"`);
    return reply.send(Buffer.from(buffer));
  });
}
