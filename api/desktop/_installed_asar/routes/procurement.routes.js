import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";

function getRole(req) {
  return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
}

function getUser(req) {
  return String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
}

function requireRoles(req, reply, roles) {
  const role = getRole(req);
  if (!roles.includes(role)) {
    reply.code(403).send({ error: `role '${role || "unknown"}' not allowed` });
    return false;
  }
  return true;
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
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
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
    const ins = db.prepare(`
      INSERT INTO procurement_requisitions (
        part_id, qty_requested, qty_received, needed_by_date, supplier_name, po_number, bill_to, request_type, requester, notes, estimated_value, status, created_at, updated_at
      ) VALUES (?, ?, 0, ?, ?, ?, ?, ?, ?, ?, ?, 'draft', datetime('now'), datetime('now'))
    `).run(part.id, qty_requested, needed_by_date, supplier_name, po_number, bill_to, request_type, requester, notes, estimated_value);
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
    const where = [];
    const params = [];
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

    const header = db.prepare(`
      SELECT pr.*, p.part_code, p.part_name
      FROM procurement_requisitions pr
      LEFT JOIN parts p ON p.id = pr.part_id
      WHERE pr.id = ?
    `).get(id);
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
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const reqn = db.prepare(`SELECT id, status FROM procurement_requisitions WHERE id = ?`).get(id);
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
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const reqn = db.prepare(`SELECT id FROM procurement_requisitions WHERE id = ?`).get(id);
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
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const reqn = db.prepare(`SELECT id, status, site_request_no FROM procurement_requisitions WHERE id = ?`).get(id);
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
    const reqn = db.prepare(`SELECT id, status FROM procurement_requisitions WHERE id = ?`).get(id);
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
    const reqn = db.prepare(`SELECT id FROM procurement_requisitions WHERE id = ?`).get(id);
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
    const reqn = db.prepare(`SELECT id, status FROM procurement_requisitions WHERE id = ?`).get(id);
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
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const approver_name = String(req.body?.approver_name || getUser(req)).trim();
    if (!approver_name) return reply.code(400).send({ error: "approver_name is required" });
    const comment =
      req.body?.comment != null && String(req.body.comment).trim() !== ""
        ? String(req.body.comment).trim()
        : null;

    const reqn = db.prepare(`SELECT id, status FROM procurement_requisitions WHERE id = ?`).get(id);
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
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });

    const row = db.prepare(`
      SELECT pr.id, pr.status, p.part_code, p.part_name, pr.qty_requested, pr.qty_received
      FROM procurement_requisitions pr
      JOIN parts p ON p.id = pr.part_id
      WHERE pr.id = ?
    `).get(id);
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
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
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
    `).get(id);
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
}
