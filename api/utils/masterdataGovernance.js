// IRONLOG/api/utils/masterdataGovernance.js — schema + validation + apply approved MDM changes
import { db } from "../db/client.js";

function hasColumn(table, col) {
  try {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name) === col);
  } catch {
    return false;
  }
}

function ensureColumn(table, colDef) {
  const [colName] = String(colDef).trim().split(/\s+/);
  if (!hasColumn(table, colName)) {
    db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
  }
}

export function normalizeMdmCode(raw) {
  return String(raw || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "_");
}

export function ensureMasterDataSchema() {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS mdm_departments (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL DEFAULT 'main',
      code TEXT NOT NULL,
      name TEXT NOT NULL,
      owner_username TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(site_code, code)
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS mdm_cost_centers (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL DEFAULT 'main',
      code TEXT NOT NULL,
      name TEXT NOT NULL,
      department_code TEXT,
      owner_username TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(site_code, code)
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS mdm_suppliers (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL DEFAULT 'main',
      supplier_code TEXT NOT NULL,
      name TEXT NOT NULL,
      contact_email TEXT,
      owner_username TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(site_code, supplier_code)
    )
  `).run();

  db.prepare(`CREATE INDEX IF NOT EXISTS idx_mdm_dept_site ON mdm_departments(site_code, active)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_mdm_cc_site ON mdm_cost_centers(site_code, active)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_mdm_sup_site ON mdm_suppliers(site_code, active)`).run();

  ensureColumn("assets", "department_code TEXT");
  ensureColumn("assets", "cost_center_code TEXT");
  ensureColumn("assets", "data_owner_username TEXT");
  ensureColumn("parts", "department_code TEXT");
  ensureColumn("parts", "default_supplier_code TEXT");
  ensureColumn("parts", "data_owner_username TEXT");
}

export function departmentExists(siteCode, code) {
  const c = normalizeMdmCode(code);
  if (!c) return false;
  const row = db
    .prepare(
      `SELECT 1 AS ok FROM mdm_departments WHERE site_code = ? AND code = ? AND active = 1`
    )
    .get(String(siteCode || "main").toLowerCase(), c);
  return Boolean(row);
}

export function costCenterExists(siteCode, code) {
  const c = normalizeMdmCode(code);
  if (!c) return false;
  const row = db
    .prepare(`SELECT 1 AS ok FROM mdm_cost_centers WHERE site_code = ? AND code = ? AND active = 1`)
    .get(String(siteCode || "main").toLowerCase(), c);
  return Boolean(row);
}

export function supplierExists(siteCode, supplierCode) {
  const c = normalizeMdmCode(supplierCode);
  if (!c) return false;
  const row = db
    .prepare(`SELECT 1 AS ok FROM mdm_suppliers WHERE site_code = ? AND supplier_code = ? AND active = 1`)
    .get(String(siteCode || "main").toLowerCase(), c);
  return Boolean(row);
}

/**
 * Apply an approved masterdata change (invoked from approvals route).
 * Payload shape:
 * { site_code, change_type: 'create'|'update', master: string, record?: object, id?: number, patch?: object, asset_code?: string, part_code?: string }
 */
export function applyMasterDataApproval(payload) {
  const site = String(payload?.site_code || "main").trim().toLowerCase() || "main";
  const changeType = String(payload?.change_type || "").trim().toLowerCase();
  const master = String(payload?.master || "").trim().toLowerCase();

  if (changeType === "create") {
    const rec = payload.record || {};
    if (master === "department") {
      const code = normalizeMdmCode(rec.code);
      const name = String(rec.name || "").trim();
      const owner_username = rec.owner_username != null ? String(rec.owner_username).trim() || null : null;
      const active = rec.active === 0 || rec.active === false ? 0 : 1;
      if (!code || !name) throw new Error("department code and name are required");
      try {
        db.prepare(
          `
          INSERT INTO mdm_departments (site_code, code, name, owner_username, active, updated_at)
          VALUES (?, ?, ?, ?, ?, datetime('now'))
        `
        ).run(site, code, name, owner_username, active);
      } catch (e) {
        if (String(e.message || e).includes("UNIQUE")) throw new Error("duplicate department code for this site");
        throw e;
      }
      return { master: "department", code, site };
    }
    if (master === "cost_center") {
      const code = normalizeMdmCode(rec.code);
      const name = String(rec.name || "").trim();
      const department_code = rec.department_code != null ? normalizeMdmCode(rec.department_code) : null;
      const owner_username = rec.owner_username != null ? String(rec.owner_username).trim() || null : null;
      const active = rec.active === 0 || rec.active === false ? 0 : 1;
      if (!code || !name) throw new Error("cost center code and name are required");
      if (department_code && !departmentExists(site, department_code)) {
        throw new Error(`department '${department_code}' does not exist for this site`);
      }
      try {
        db.prepare(
          `
          INSERT INTO mdm_cost_centers (site_code, code, name, department_code, owner_username, active, updated_at)
          VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
        `
        ).run(site, code, name, department_code, owner_username, active);
      } catch (e) {
        if (String(e.message || e).includes("UNIQUE")) throw new Error("duplicate cost center code for this site");
        throw e;
      }
      return { master: "cost_center", code, site };
    }
    if (master === "supplier") {
      const supplier_code = normalizeMdmCode(rec.supplier_code || rec.code);
      const name = String(rec.name || "").trim();
      const contact_email = rec.contact_email != null ? String(rec.contact_email).trim() || null : null;
      const owner_username = rec.owner_username != null ? String(rec.owner_username).trim() || null : null;
      const active = rec.active === 0 || rec.active === false ? 0 : 1;
      if (!supplier_code || !name) throw new Error("supplier code and name are required");
      try {
        db.prepare(
          `
          INSERT INTO mdm_suppliers (site_code, supplier_code, name, contact_email, owner_username, active, updated_at)
          VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
        `
        ).run(site, supplier_code, name, contact_email, owner_username, active);
      } catch (e) {
        if (String(e.message || e).includes("UNIQUE")) throw new Error("duplicate supplier code for this site");
        throw e;
      }
      return { master: "supplier", supplier_code, site };
    }
    throw new Error(`unsupported create master: ${master}`);
  }

  if (changeType === "update") {
    const patch = payload.patch || {};
    if (master === "department") {
      const code = normalizeMdmCode(payload.code || patch.code);
      if (!code) throw new Error("department code is required");
      const row = db
        .prepare(`SELECT id FROM mdm_departments WHERE site_code = ? AND code = ?`)
        .get(site, code);
      if (!row) throw new Error("department not found");
      const sets = [];
      const vals = [];
      if (patch.name !== undefined) {
        sets.push("name = ?");
        vals.push(String(patch.name || "").trim());
      }
      if (patch.owner_username !== undefined) {
        sets.push("owner_username = ?");
        vals.push(patch.owner_username != null ? String(patch.owner_username).trim() : null);
      }
      if (patch.active !== undefined) {
        sets.push("active = ?");
        vals.push(patch.active === 0 || patch.active === false ? 0 : 1);
      }
      if (!sets.length) return { master: "department", code, site, updated: false };
      sets.push("updated_at = datetime('now')");
      vals.push(site, code);
      db.prepare(`UPDATE mdm_departments SET ${sets.join(", ")} WHERE site_code = ? AND code = ?`).run(...vals);
      return { master: "department", code, site, updated: true };
    }
    if (master === "cost_center") {
      const code = normalizeMdmCode(payload.code || patch.code);
      if (!code) throw new Error("cost center code is required");
      const row = db
        .prepare(`SELECT id FROM mdm_cost_centers WHERE site_code = ? AND code = ?`)
        .get(site, code);
      if (!row) throw new Error("cost center not found");
      const sets = [];
      const vals = [];
      if (patch.name !== undefined) {
        sets.push("name = ?");
        vals.push(String(patch.name || "").trim());
      }
      if (patch.department_code !== undefined) {
        const department_code =
          patch.department_code != null ? normalizeMdmCode(patch.department_code) : null;
        if (department_code && !departmentExists(site, department_code)) {
          throw new Error(`department '${department_code}' does not exist for this site`);
        }
        sets.push("department_code = ?");
        vals.push(department_code);
      }
      if (patch.owner_username !== undefined) {
        sets.push("owner_username = ?");
        vals.push(patch.owner_username != null ? String(patch.owner_username).trim() : null);
      }
      if (patch.active !== undefined) {
        sets.push("active = ?");
        vals.push(patch.active === 0 || patch.active === false ? 0 : 1);
      }
      if (!sets.length) return { master: "cost_center", code, site, updated: false };
      sets.push("updated_at = datetime('now')");
      vals.push(site, code);
      db.prepare(`UPDATE mdm_cost_centers SET ${sets.join(", ")} WHERE site_code = ? AND code = ?`).run(...vals);
      return { master: "cost_center", code, site, updated: true };
    }
    if (master === "supplier") {
      const supplier_code = normalizeMdmCode(payload.supplier_code || patch.supplier_code);
      if (!supplier_code) throw new Error("supplier code is required");
      const row = db
        .prepare(`SELECT id FROM mdm_suppliers WHERE site_code = ? AND supplier_code = ?`)
        .get(site, supplier_code);
      if (!row) throw new Error("supplier not found");
      const sets = [];
      const vals = [];
      if (patch.name !== undefined) {
        sets.push("name = ?");
        vals.push(String(patch.name || "").trim());
      }
      if (patch.contact_email !== undefined) {
        sets.push("contact_email = ?");
        vals.push(patch.contact_email != null ? String(patch.contact_email).trim() : null);
      }
      if (patch.owner_username !== undefined) {
        sets.push("owner_username = ?");
        vals.push(patch.owner_username != null ? String(patch.owner_username).trim() : null);
      }
      if (patch.active !== undefined) {
        sets.push("active = ?");
        vals.push(patch.active === 0 || patch.active === false ? 0 : 1);
      }
      if (!sets.length) return { master: "supplier", supplier_code, site, updated: false };
      sets.push("updated_at = datetime('now')");
      vals.push(site, supplier_code);
      db.prepare(`UPDATE mdm_suppliers SET ${sets.join(", ")} WHERE site_code = ? AND supplier_code = ?`).run(...vals);
      return { master: "supplier", supplier_code, site, updated: true };
    }
    if (master === "asset_governance") {
      const asset_code = String(payload.asset_code || patch.asset_code || "").trim();
      if (!asset_code) throw new Error("asset_code is required");
      const sets = [];
      const vals = [];
      if (patch.department_code !== undefined) {
        const dept = patch.department_code != null ? normalizeMdmCode(patch.department_code) : null;
        if (dept && !departmentExists(site, dept)) throw new Error(`department '${dept}' not found`);
        sets.push("department_code = ?");
        vals.push(dept);
      }
      if (patch.cost_center_code !== undefined) {
        const cc = patch.cost_center_code != null ? normalizeMdmCode(patch.cost_center_code) : null;
        if (cc && !costCenterExists(site, cc)) throw new Error(`cost center '${cc}' not found`);
        sets.push("cost_center_code = ?");
        vals.push(cc);
      }
      if (patch.data_owner_username !== undefined) {
        sets.push("data_owner_username = ?");
        vals.push(patch.data_owner_username != null ? String(patch.data_owner_username).trim() : null);
      }
      if (!sets.length) return { master: "asset_governance", asset_code, site, updated: false };
      const a = db.prepare(`SELECT id FROM assets WHERE asset_code = ?`).get(asset_code);
      if (!a) throw new Error("asset not found");
      vals.push(asset_code);
      db.prepare(`UPDATE assets SET ${sets.join(", ")} WHERE asset_code = ?`).run(...vals);
      return { master: "asset_governance", asset_code, site, updated: true };
    }
    if (master === "part_governance") {
      const part_code = String(payload.part_code || patch.part_code || "").trim();
      if (!part_code) throw new Error("part_code is required");
      const sets = [];
      const vals = [];
      if (patch.department_code !== undefined) {
        const dept = patch.department_code != null ? normalizeMdmCode(patch.department_code) : null;
        if (dept && !departmentExists(site, dept)) throw new Error(`department '${dept}' not found`);
        sets.push("department_code = ?");
        vals.push(dept);
      }
      if (patch.default_supplier_code !== undefined) {
        const sup = patch.default_supplier_code != null ? normalizeMdmCode(patch.default_supplier_code) : null;
        if (sup && !supplierExists(site, sup)) throw new Error(`supplier '${sup}' not found`);
        sets.push("default_supplier_code = ?");
        vals.push(sup);
      }
      if (patch.data_owner_username !== undefined) {
        sets.push("data_owner_username = ?");
        vals.push(patch.data_owner_username != null ? String(patch.data_owner_username).trim() : null);
      }
      if (!sets.length) return { master: "part_governance", part_code, site, updated: false };
      const p = db.prepare(`SELECT id FROM parts WHERE part_code = ?`).get(part_code);
      if (!p) throw new Error("part not found");
      vals.push(part_code);
      db.prepare(`UPDATE parts SET ${sets.join(", ")} WHERE part_code = ?`).run(...vals);
      return { master: "part_governance", part_code, site, updated: true };
    }
    throw new Error(`unsupported update master: ${master}`);
  }

  throw new Error(`unsupported change_type: ${changeType}`);
}

export function validateAssetGovernanceOptional(siteCode, body) {
  const site = String(siteCode || "main").toLowerCase();
  if (body.department_code != null && String(body.department_code).trim()) {
    const d = normalizeMdmCode(body.department_code);
    if (!departmentExists(site, d)) {
      return { ok: false, error: `department '${d}' is not in the master list for this site` };
    }
  }
  if (body.cost_center_code != null && String(body.cost_center_code).trim()) {
    const c = normalizeMdmCode(body.cost_center_code);
    if (!costCenterExists(site, c)) {
      return { ok: false, error: `cost center '${c}' is not in the master list for this site` };
    }
  }
  return { ok: true };
}
