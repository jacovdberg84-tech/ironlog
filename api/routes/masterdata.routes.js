// IRONLOG/api/routes/masterdata.routes.js — Master data governance (departments, cost centers, suppliers, governance patches)
import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";
import {
  applyMasterDataApproval,
  ensureMasterDataSchema,
  normalizeMdmCode,
} from "../utils/masterdataGovernance.js";

function getSiteCode(req) {
  return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
}

function getUser(req) {
  return String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
}

function getRole(req) {
  return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
}

function requireRoles(req, reply, roles) {
  const role = getRole(req);
  if (!roles.includes(role)) {
    reply.code(403).send({ error: `role '${role || "unknown"}' not allowed` });
    return false;
  }
  return true;
}

/** Admin / supervisor: full MDM CRUD. Managers: read + approve via /api/approvals. */
const MDM_WRITE_ROLES = ["admin", "supervisor"];
const MDM_READ_ROLES = [
  "admin",
  "supervisor",
  "plant_manager",
  "site_manager",
  "executive",
  "finance",
  "procurement",
  "storeman",
  "stores",
  "quality_manager",
  "hr_manager",
];

/** Roles that may submit a change request for approval */
const MDM_REQUEST_ROLES = [
  "procurement",
  "storeman",
  "stores",
  "plant_manager",
  "site_manager",
  "quality_manager",
  "hr_manager",
  "supervisor",
  "artisan",
  "operator",
];

export default async function masterdataRoutes(app) {
  ensureAuditTable(db);
  ensureMasterDataSchema();

  function insertApprovalRequest(payload, req) {
    const payload_json = JSON.stringify(payload);
    const ins = db
      .prepare(
        `
      INSERT INTO approval_requests (module, action, entity_type, entity_id, status, payload_json, requested_by, requested_role)
      VALUES ('masterdata', 'apply_change', 'masterdata', ?, 'pending', ?, ?, ?)
    `
      )
      .run(
        String(payload.master || "change"),
        payload_json,
        getUser(req),
        getRole(req)
      );
    return Number(ins.lastInsertRowid);
  }

  // ---------- Summary ----------
  app.get("/summary", async (req, reply) => {
    if (!requireRoles(req, reply, MDM_READ_ROLES)) return;
    const site = getSiteCode(req);
    const dep = db.prepare(`SELECT COUNT(*) AS c FROM mdm_departments WHERE site_code = ? AND active = 1`).get(site);
    const cc = db.prepare(`SELECT COUNT(*) AS c FROM mdm_cost_centers WHERE site_code = ? AND active = 1`).get(site);
    const sup = db.prepare(`SELECT COUNT(*) AS c FROM mdm_suppliers WHERE site_code = ? AND active = 1`).get(site);
    const pending = db
      .prepare(
        `
      SELECT COUNT(*) AS c FROM approval_requests
      WHERE module = 'masterdata' AND LOWER(status) = 'pending'
    `
      )
      .get();
    return {
      ok: true,
      site_code: site,
      counts: {
        departments: Number(dep?.c || 0),
        cost_centers: Number(cc?.c || 0),
        suppliers: Number(sup?.c || 0),
        pending_masterdata_approvals: Number(pending?.c || 0),
      },
    };
  });

  // ---------- Departments ----------
  app.get("/departments", async (req, reply) => {
    if (!requireRoles(req, reply, MDM_READ_ROLES)) return;
    const site = getSiteCode(req);
    const rows = db
      .prepare(
        `
      SELECT id, site_code, code, name, owner_username, active, created_at, updated_at
      FROM mdm_departments
      WHERE site_code = ?
      ORDER BY code ASC
    `
      )
      .all(site);
    return { ok: true, rows };
  });

  app.post("/departments", async (req, reply) => {
    if (!requireRoles(req, reply, MDM_WRITE_ROLES)) return;
    const site = getSiteCode(req);
    const body = req.body || {};
    try {
      const out = applyMasterDataApproval({
        site_code: site,
        change_type: "create",
        master: "department",
        record: body,
      });
      writeAudit(db, req, {
        module: "masterdata",
        action: "department_create",
        entity_type: "mdm_department",
        entity_id: out.code,
        payload: body,
      });
      return { ok: true, ...out };
    } catch (e) {
      return reply.code(400).send({ error: e.message || String(e) });
    }
  });

  app.patch("/departments/:code", async (req, reply) => {
    if (!requireRoles(req, reply, MDM_WRITE_ROLES)) return;
    const site = getSiteCode(req);
    const code = normalizeMdmCode(req.params.code);
    try {
      const out = applyMasterDataApproval({
        site_code: site,
        change_type: "update",
        master: "department",
        code,
        patch: req.body || {},
      });
      writeAudit(db, req, {
        module: "masterdata",
        action: "department_update",
        entity_type: "mdm_department",
        entity_id: code,
        payload: req.body || {},
      });
      return { ok: true, ...out };
    } catch (e) {
      return reply.code(400).send({ error: e.message || String(e) });
    }
  });

  // ---------- Cost centers ----------
  app.get("/cost-centers", async (req, reply) => {
    if (!requireRoles(req, reply, MDM_READ_ROLES)) return;
    const site = getSiteCode(req);
    const rows = db
      .prepare(
        `
      SELECT id, site_code, code, name, department_code, owner_username, active, created_at, updated_at
      FROM mdm_cost_centers
      WHERE site_code = ?
      ORDER BY code ASC
    `
      )
      .all(site);
    return { ok: true, rows };
  });

  app.post("/cost-centers", async (req, reply) => {
    if (!requireRoles(req, reply, MDM_WRITE_ROLES)) return;
    const site = getSiteCode(req);
    try {
      const out = applyMasterDataApproval({
        site_code: site,
        change_type: "create",
        master: "cost_center",
        record: req.body || {},
      });
      writeAudit(db, req, {
        module: "masterdata",
        action: "cost_center_create",
        entity_type: "mdm_cost_center",
        entity_id: out.code,
        payload: req.body || {},
      });
      return { ok: true, ...out };
    } catch (e) {
      return reply.code(400).send({ error: e.message || String(e) });
    }
  });

  app.patch("/cost-centers/:code", async (req, reply) => {
    if (!requireRoles(req, reply, MDM_WRITE_ROLES)) return;
    const site = getSiteCode(req);
    const code = normalizeMdmCode(req.params.code);
    try {
      const out = applyMasterDataApproval({
        site_code: site,
        change_type: "update",
        master: "cost_center",
        code,
        patch: req.body || {},
      });
      writeAudit(db, req, {
        module: "masterdata",
        action: "cost_center_update",
        entity_type: "mdm_cost_center",
        entity_id: code,
        payload: req.body || {},
      });
      return { ok: true, ...out };
    } catch (e) {
      return reply.code(400).send({ error: e.message || String(e) });
    }
  });

  // ---------- Suppliers ----------
  app.get("/suppliers", async (req, reply) => {
    if (!requireRoles(req, reply, MDM_READ_ROLES)) return;
    const site = getSiteCode(req);
    const rows = db
      .prepare(
        `
      SELECT id, site_code, supplier_code, name, contact_email, owner_username, active, created_at, updated_at
      FROM mdm_suppliers
      WHERE site_code = ?
      ORDER BY supplier_code ASC
    `
      )
      .all(site);
    return { ok: true, rows };
  });

  app.post("/suppliers", async (req, reply) => {
    if (!requireRoles(req, reply, MDM_WRITE_ROLES)) return;
    const site = getSiteCode(req);
    try {
      const out = applyMasterDataApproval({
        site_code: site,
        change_type: "create",
        master: "supplier",
        record: req.body || {},
      });
      writeAudit(db, req, {
        module: "masterdata",
        action: "supplier_create",
        entity_type: "mdm_supplier",
        entity_id: out.supplier_code,
        payload: req.body || {},
      });
      return { ok: true, ...out };
    } catch (e) {
      return reply.code(400).send({ error: e.message || String(e) });
    }
  });

  app.patch("/suppliers/:code", async (req, reply) => {
    if (!requireRoles(req, reply, MDM_WRITE_ROLES)) return;
    const site = getSiteCode(req);
    const supplier_code = normalizeMdmCode(req.params.code);
    try {
      const out = applyMasterDataApproval({
        site_code: site,
        change_type: "update",
        master: "supplier",
        supplier_code,
        patch: req.body || {},
      });
      writeAudit(db, req, {
        module: "masterdata",
        action: "supplier_update",
        entity_type: "mdm_supplier",
        entity_id: supplier_code,
        payload: req.body || {},
      });
      return { ok: true, ...out };
    } catch (e) {
      return reply.code(400).send({ error: e.message || String(e) });
    }
  });

  /**
   * POST /change-requests
   * Body: { site_code?: (defaults header), change_type, master, record?, code?, supplier_code?, asset_code?, part_code?, patch? }
   * Creates approval_requests for supervisor/admin (or manager) to approve via /api/approvals/:id/approve
   */
  app.post("/change-requests", async (req, reply) => {
    if (!requireRoles(req, reply, MDM_REQUEST_ROLES)) return;
    const site = String(req.body?.site_code || getSiteCode(req))
      .trim()
      .toLowerCase() || "main";
    const change_type = String(req.body?.change_type || "").trim().toLowerCase();
    const master = String(req.body?.master || "").trim().toLowerCase();
    if (!["create", "update"].includes(change_type)) {
      return reply.code(400).send({ error: "change_type must be create or update" });
    }
    if (!master) return reply.code(400).send({ error: "master is required" });

    const payload = {
      site_code: site,
      change_type,
      master,
      record: req.body?.record,
      code: req.body?.code != null ? normalizeMdmCode(req.body.code) : undefined,
      supplier_code:
        req.body?.supplier_code != null ? normalizeMdmCode(req.body.supplier_code) : undefined,
      asset_code: req.body?.asset_code != null ? String(req.body.asset_code).trim() : undefined,
      part_code: req.body?.part_code != null ? String(req.body.part_code).trim() : undefined,
      patch: req.body?.patch || {},
    };

    // Admins can apply immediately (optional shortcut)
    if (MDM_WRITE_ROLES.includes(getRole(req))) {
      try {
        const execution = applyMasterDataApproval(payload);
        writeAudit(db, req, {
          module: "masterdata",
          action: "change_apply_direct",
          entity_type: "masterdata",
          entity_id: String(master),
          payload,
        });
        return { ok: true, applied: true, execution };
      } catch (e) {
        return reply.code(400).send({ error: e.message || String(e) });
      }
    }

    const id = insertApprovalRequest(payload, req);
    writeAudit(db, req, {
      module: "masterdata",
      action: "change_request",
      entity_type: "approval_request",
      entity_id: id,
      payload: { master, change_type },
    });
    return { ok: true, pending_approval: true, request_id: id };
  });
}
