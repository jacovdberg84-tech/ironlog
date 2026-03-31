import multipart from "@fastify/multipart";
import fs from "node:fs";
import path from "node:path";
import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";

const DEFAULT_DEPARTMENTS = [
  "Operations",
  "Maintenance",
  "Stores",
  "Safety",
  "HR",
  "Finance",
  "Engineering",
  "Quality",
];
const LEGAL_STATUSES = ["draft", "pending_approval", "approved", "rejected", "superseded"];

function safeFilePart(v) {
  return String(v || "")
    .replace(/[^a-zA-Z0-9._-]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 80);
}

function hasColumn(tableName, columnName) {
  const rows = db.prepare(`PRAGMA table_info(${tableName})`).all();
  return rows.some((r) => String(r.name || "").toLowerCase() === String(columnName || "").toLowerCase());
}

function ensureColumn(tableName, columnSqlDef, columnName) {
  if (!hasColumn(tableName, columnName)) {
    db.prepare(`ALTER TABLE ${tableName} ADD COLUMN ${columnSqlDef}`).run();
  }
}

export default async function legalRoutes(app) {
  const dataRoot = process.env.IRONLOG_DATA_DIR || process.cwd();
  ensureAuditTable(db);

  await app.register(multipart, {
    limits: { fileSize: 30 * 1024 * 1024 }, // 30MB
  });

  db.prepare(`
    CREATE TABLE IF NOT EXISTS legal_documents (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      department TEXT NOT NULL,
      title TEXT NOT NULL,
      doc_type TEXT,
      version TEXT,
      owner TEXT,
      effective_date TEXT,
      expiry_date TEXT,
      file_name TEXT NOT NULL,
      file_path TEXT NOT NULL,
      mime_type TEXT,
      size_bytes INTEGER,
      uploaded_by TEXT,
      status TEXT NOT NULL DEFAULT 'draft',
      approved_by TEXT,
      approved_at TEXT,
      approval_note TEXT,
      supersedes_document_id INTEGER,
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  ensureColumn("legal_documents", "status TEXT NOT NULL DEFAULT 'draft'", "status");
  ensureColumn("legal_documents", "approved_by TEXT", "approved_by");
  ensureColumn("legal_documents", "approved_at TEXT", "approved_at");
  ensureColumn("legal_documents", "approval_note TEXT", "approval_note");
  ensureColumn("legal_documents", "supersedes_document_id INTEGER", "supersedes_document_id");

  db.prepare(`
    CREATE TABLE IF NOT EXISTS legal_document_actions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      document_id INTEGER NOT NULL,
      action TEXT NOT NULL,
      from_status TEXT,
      to_status TEXT,
      note TEXT,
      username TEXT,
      role TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (document_id) REFERENCES legal_documents(id) ON DELETE CASCADE
    )
  `).run();

  const legalDir = path.join(dataRoot, "uploads", "legal");
  if (!fs.existsSync(legalDir)) {
    fs.mkdirSync(legalDir, { recursive: true });
  }

  function getRole(req) {
    return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
  }

  function getUser(req) {
    return String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
  }

  function requireLegalWrite(req, reply) {
    const role = getRole(req);
    if (!["admin", "supervisor"].includes(role)) {
      reply.code(403).send({ error: `role '${role || "unknown"}' not allowed` });
      return false;
    }
    return true;
  }

  // GET /api/legal/departments
  app.get("/departments", async () => {
    const rows = db.prepare(`
      SELECT DISTINCT department
      FROM legal_documents
      WHERE department IS NOT NULL AND TRIM(department) <> ''
      ORDER BY department ASC
    `).all();
    const extra = rows.map((r) => String(r.department || "").trim()).filter(Boolean);
    const departments = Array.from(new Set([...DEFAULT_DEPARTMENTS, ...extra]));
    return { ok: true, departments };
  });

  // GET /api/legal?department=&q=&include_inactive=0
  app.get("/", async (req) => {
    const department = String(req.query?.department || "").trim();
    const q = String(req.query?.q || "").trim().toLowerCase();
    const includeInactive = String(req.query?.include_inactive || "0") === "1";
    const status = String(req.query?.status || "").trim().toLowerCase();

    const where = [];
    const params = [];
    if (!includeInactive) where.push("ld.active = 1");
    if (department) {
      where.push("ld.department = ?");
      params.push(department);
    }
    if (status && LEGAL_STATUSES.includes(status)) {
      where.push("ld.status = ?");
      params.push(status);
    }
    if (q) {
      where.push("(LOWER(ld.title) LIKE ? OR LOWER(ld.doc_type) LIKE ? OR LOWER(ld.owner) LIKE ? OR LOWER(ld.file_name) LIKE ?)");
      params.push(`%${q}%`, `%${q}%`, `%${q}%`, `%${q}%`);
    }

    const rows = db.prepare(`
      SELECT
        ld.id,
        ld.department,
        ld.title,
        ld.doc_type,
        ld.version,
        ld.owner,
        ld.effective_date,
        ld.expiry_date,
        ld.file_name,
        ld.mime_type,
        ld.size_bytes,
        ld.uploaded_by,
        ld.status,
        ld.approved_by,
        ld.approved_at,
        ld.approval_note,
        ld.supersedes_document_id,
        ld.active,
        ld.created_at
      FROM legal_documents ld
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY ld.id DESC
      LIMIT 500
    `).all(...params).map((r) => ({
      ...r,
      id: Number(r.id),
      active: Number(r.active),
      size_bytes: Number(r.size_bytes || 0),
    }));

    return { ok: true, rows };
  });

  // GET /api/legal/expiry?days=90&include_archived=0&department=&status=approved
  app.get("/expiry", async (req) => {
    const daysRaw = Number(req.query?.days ?? 90);
    const days = Number.isFinite(daysRaw) ? Math.min(3650, Math.max(1, Math.trunc(daysRaw))) : 90;
    const includeArchived = String(req.query?.include_archived || "0") === "1";
    const department = String(req.query?.department || "").trim();
    const status = String(req.query?.status || "approved").trim().toLowerCase();

    const where = ["ld.expiry_date IS NOT NULL", "TRIM(ld.expiry_date) <> ''"];
    const params = [];
    if (!includeArchived) where.push("ld.active = 1");
    if (department) {
      where.push("ld.department = ?");
      params.push(department);
    }
    if (status && status !== "all" && LEGAL_STATUSES.includes(status)) {
      where.push("ld.status = ?");
      params.push(status);
    }

    const rows = db.prepare(`
      SELECT
        ld.id,
        ld.department,
        ld.title,
        ld.doc_type,
        ld.version,
        ld.owner,
        ld.status,
        ld.active,
        ld.expiry_date,
        CAST(julianday(ld.expiry_date) - julianday(DATE('now')) AS INTEGER) AS days_to_expiry
      FROM legal_documents ld
      WHERE ${where.join(" AND ")}
      ORDER BY ld.expiry_date ASC, ld.id DESC
      LIMIT 1000
    `).all(...params).map((r) => ({
      ...r,
      id: Number(r.id),
      active: Number(r.active),
      days_to_expiry: Number(r.days_to_expiry),
    }));

    const expired = rows.filter((r) => Number(r.days_to_expiry) < 0);
    const d30 = rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= 30);
    const d60 = rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= 60);
    const d90 = rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= 90);
    const inWindow = rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= days);

    return {
      ok: true,
      days,
      summary: {
        total_with_expiry: rows.length,
        expired: expired.length,
        due_30: d30.length,
        due_60: d60.length,
        due_90: d90.length,
        due_in_window: inWindow.length,
      },
      rows,
      due: inWindow,
      expired,
    };
  });

  // POST /api/legal/upload (multipart)
  // fields: file + department + title + doc_type? + version? + owner? + effective_date? + expiry_date?
  app.post("/upload", async (req, reply) => {
    if (!requireLegalWrite(req, reply)) return;

    const file = await req.file();
    if (!file) return reply.code(400).send({ error: "Upload a file field named 'file'." });

    const fields = file.fields || {};
    const department = String(fields.department?.value || "").trim();
    const title = String(fields.title?.value || "").trim();
    const doc_type = String(fields.doc_type?.value || "").trim() || null;
    const version = String(fields.version?.value || "").trim() || null;
    const owner = String(fields.owner?.value || "").trim() || null;
    const effective_date = String(fields.effective_date?.value || "").trim() || null;
    const expiry_date = String(fields.expiry_date?.value || "").trim() || null;

    if (!department) return reply.code(400).send({ error: "department is required" });
    if (!title) return reply.code(400).send({ error: "title is required" });

    const original = String(file.filename || "document.bin");
    const ext = path.extname(original || "").slice(0, 12);
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const fileName = `${stamp}_${safeFilePart(department)}_${safeFilePart(title)}${ext || ""}`;
    const absPath = path.join(legalDir, fileName);

    const buf = await file.toBuffer();
    fs.writeFileSync(absPath, buf);

    const relPath = path.join("uploads", "legal", fileName).replace(/\\/g, "/");
    const uploaded_by = getUser(req);
    const ins = db.prepare(`
      INSERT INTO legal_documents (
        department, title, doc_type, version, owner, effective_date, expiry_date,
        file_name, file_path, mime_type, size_bytes, uploaded_by, status, active
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'draft', 1)
    `).run(
      department,
      title,
      doc_type,
      version,
      owner,
      effective_date,
      expiry_date,
      original,
      relPath,
      String(file.mimetype || "").trim() || null,
      Number(buf.length || 0),
      uploaded_by
    );

    const id = Number(ins.lastInsertRowid);
    db.prepare(`
      INSERT INTO legal_document_actions (
        document_id, action, from_status, to_status, note, username, role
      ) VALUES (?, 'upload', NULL, 'draft', NULL, ?, ?)
    `).run(id, uploaded_by, getRole(req));

    writeAudit(db, req, {
      module: "legal",
      action: "upload",
      entity_type: "legal_document",
      entity_id: id,
      payload: { department, title, file_name: original },
    });

    return reply.send({ ok: true, id });
  });

  // GET /api/legal/:id/download
  app.get("/:id/download", async (req, reply) => {
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });

    const row = db.prepare(`
      SELECT id, file_name, file_path, mime_type
      FROM legal_documents
      WHERE id = ?
    `).get(id);
    if (!row) return reply.code(404).send({ error: "document not found" });

    const absPath = path.join(dataRoot, String(row.file_path || ""));
    if (!fs.existsSync(absPath)) return reply.code(404).send({ error: "file not found on disk" });

    const content = fs.readFileSync(absPath);
    reply
      .header("Content-Type", row.mime_type || "application/octet-stream")
      .header("Content-Disposition", `attachment; filename="${row.file_name || "document.bin"}"`)
      .send(content);
  });

  // GET /api/legal/:id/actions
  app.get("/:id/actions", async (req, reply) => {
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });

    const row = db.prepare(`SELECT id FROM legal_documents WHERE id = ?`).get(id);
    if (!row) return reply.code(404).send({ error: "document not found" });

    const actions = db.prepare(`
      SELECT id, action, from_status, to_status, note, username, role, created_at
      FROM legal_document_actions
      WHERE document_id = ?
      ORDER BY id DESC
      LIMIT 200
    `).all(id).map((a) => ({ ...a, id: Number(a.id) }));

    return { ok: true, actions };
  });

  // POST /api/legal/:id/status
  // Body: { status, note?, supersedes_document_id? }
  app.post("/:id/status", async (req, reply) => {
    if (!requireLegalWrite(req, reply)) return;
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });

    const next = String(req.body?.status || "").trim().toLowerCase();
    const note = String(req.body?.note || "").trim() || null;
    const supersedes_document_id =
      req.body?.supersedes_document_id != null ? Number(req.body.supersedes_document_id) : null;
    if (!LEGAL_STATUSES.includes(next)) {
      return reply.code(400).send({ error: `status must be one of: ${LEGAL_STATUSES.join(", ")}` });
    }

    const doc = db.prepare(`
      SELECT id, status, active
      FROM legal_documents
      WHERE id = ?
    `).get(id);
    if (!doc) return reply.code(404).send({ error: "document not found" });
    if (Number(doc.active) !== 1) return reply.code(409).send({ error: "cannot change status of archived document" });

    const current = String(doc.status || "draft").toLowerCase();
    const transitions = {
      draft: ["pending_approval", "superseded"],
      rejected: ["pending_approval", "superseded"],
      pending_approval: ["approved", "rejected", "superseded"],
      approved: ["superseded"],
      superseded: [],
    };
    if (current === next) {
      return { ok: true, id, status: current, unchanged: true };
    }
    if (!(transitions[current] || []).includes(next)) {
      return reply.code(409).send({ error: `invalid transition ${current} -> ${next}` });
    }

    const username = getUser(req);
    const role = getRole(req);
    const approved_by = next === "approved" ? username : null;
    const approved_at = next === "approved" ? "datetime('now')" : "NULL";

    db.prepare(`
      UPDATE legal_documents
      SET
        status = ?,
        approved_by = ?,
        approved_at = ${approved_at},
        approval_note = ?,
        supersedes_document_id = CASE WHEN ? IS NOT NULL THEN ? ELSE supersedes_document_id END
      WHERE id = ?
    `).run(next, approved_by, note, supersedes_document_id, supersedes_document_id, id);

    db.prepare(`
      INSERT INTO legal_document_actions (
        document_id, action, from_status, to_status, note, username, role
      ) VALUES (?, 'status_change', ?, ?, ?, ?, ?)
    `).run(id, current, next, note, username, role);

    writeAudit(db, req, {
      module: "legal",
      action: "status_change",
      entity_type: "legal_document",
      entity_id: id,
      payload: { from: current, to: next, note, supersedes_document_id },
    });

    return { ok: true, id, from: current, status: next };
  });

  // POST /api/legal/:id/archive  { active: 0|1 }
  app.post("/:id/archive", async (req, reply) => {
    if (!requireLegalWrite(req, reply)) return;
    const id = Number(req.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: "invalid id" });
    const active = req.body?.active === 0 || req.body?.active === false ? 0 : 1;

    const row = db.prepare(`SELECT id FROM legal_documents WHERE id = ?`).get(id);
    if (!row) return reply.code(404).send({ error: "document not found" });

    db.prepare(`
      UPDATE legal_documents
      SET active = ?
      WHERE id = ?
    `).run(active, id);

    const current = db.prepare(`SELECT status FROM legal_documents WHERE id = ?`).get(id);
    db.prepare(`
      INSERT INTO legal_document_actions (
        document_id, action, from_status, to_status, note, username, role
      ) VALUES (?, ?, ?, ?, NULL, ?, ?)
    `).run(id, active ? "reactivate" : "archive", current?.status || null, current?.status || null, getUser(req), getRole(req));

    writeAudit(db, req, {
      module: "legal",
      action: active ? "reactivate" : "archive",
      entity_type: "legal_document",
      entity_id: id,
      payload: { active },
    });

    return { ok: true, id, active };
  });
}
