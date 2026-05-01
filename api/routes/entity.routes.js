// IRONLOG/api/routes/entity.routes.js
// Multi-entity foundation: company -> site, currency readiness, tax profiles.

import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";

function getRole(req) {
  return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
}
function getUser(req) {
  return String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
}
function getRoles(req) {
  const many = String(req.headers["x-user-roles"] || "")
    .split(",").map((x) => String(x || "").trim().toLowerCase()).filter(Boolean);
  const one = String(req.headers["x-user-role"] || "")
    .split(",").map((x) => String(x || "").trim().toLowerCase()).filter(Boolean);
  return Array.from(new Set([...many, ...one]));
}
function hasAnyRole(req, allowed) {
  const roles = getRoles(req);
  return roles.some((r) => allowed.includes(r));
}
function requireRoles(req, reply, allowed) {
  if (!hasAnyRole(req, allowed)) {
    reply.code(403).send({ error: `role '${getRole(req)}' not allowed` });
    return false;
  }
  return true;
}

function normalizeCode(s) {
  return String(s || "").trim().toUpperCase().replace(/\s+/g, "-");
}
function normalizeCcy(s) {
  return String(s || "").trim().toUpperCase();
}

export default async function entityRoutes(app) {
  ensureAuditTable(db);

  db.prepare(`
    CREATE TABLE IF NOT EXISTS company_profiles (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      company_code TEXT NOT NULL UNIQUE,
      company_name TEXT NOT NULL,
      registration_no TEXT,
      base_currency TEXT NOT NULL DEFAULT 'USD',
      reporting_currency TEXT,
      fiscal_year_start TEXT,
      tax_region TEXT,
      timezone TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      notes TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS site_profiles (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL UNIQUE,
      company_code TEXT NOT NULL,
      site_name TEXT NOT NULL,
      region TEXT,
      timezone TEXT,
      local_currency TEXT,
      tax_region TEXT,
      address TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      notes TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_site_profiles_company ON site_profiles(company_code)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS currency_rates (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      from_currency TEXT NOT NULL,
      to_currency TEXT NOT NULL,
      rate REAL NOT NULL,
      effective_date TEXT NOT NULL,
      source TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE (from_currency, to_currency, effective_date)
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_ccyr_lookup ON currency_rates(from_currency, to_currency, effective_date)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS tax_profiles (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      tax_code TEXT NOT NULL UNIQUE,
      label TEXT NOT NULL,
      rate_pct REAL NOT NULL DEFAULT 0,
      region TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      applies_to TEXT,
      notes TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  /* --------- COMPANY --------- */
  app.post("/companies", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "executive"])) return;
    const company_code = normalizeCode(req.body?.company_code);
    const company_name = String(req.body?.company_name || "").trim();
    if (!company_code || !company_name) return reply.code(400).send({ error: "company_code and company_name required" });
    const base_currency = normalizeCcy(req.body?.base_currency || "USD") || "USD";
    const reporting_currency = req.body?.reporting_currency ? normalizeCcy(req.body.reporting_currency) : null;
    const registration_no = req.body?.registration_no ? String(req.body.registration_no).trim() : null;
    const fiscal_year_start = req.body?.fiscal_year_start ? String(req.body.fiscal_year_start).trim() : null;
    const tax_region = req.body?.tax_region ? String(req.body.tax_region).trim() : null;
    const timezone = req.body?.timezone ? String(req.body.timezone).trim() : null;
    const notes = req.body?.notes ? String(req.body.notes).trim() : null;
    db.prepare(`
      INSERT INTO company_profiles
        (company_code, company_name, registration_no, base_currency, reporting_currency,
         fiscal_year_start, tax_region, timezone, active, notes, created_by, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?, datetime('now'))
      ON CONFLICT(company_code) DO UPDATE SET
        company_name = excluded.company_name,
        registration_no = excluded.registration_no,
        base_currency = excluded.base_currency,
        reporting_currency = excluded.reporting_currency,
        fiscal_year_start = excluded.fiscal_year_start,
        tax_region = excluded.tax_region,
        timezone = excluded.timezone,
        notes = excluded.notes,
        active = 1,
        updated_at = datetime('now')
    `).run(company_code, company_name, registration_no, base_currency, reporting_currency,
      fiscal_year_start, tax_region, timezone, notes, getUser(req));
    writeAudit(db, req, {
      module: "entity", action: "company.upsert", entity_type: "company_profiles",
      entity_id: company_code, payload: { company_code, company_name, base_currency, reporting_currency }
    });
    return { ok: true, company_code };
  });

  app.get("/companies", async () => {
    const rows = db.prepare(`
      SELECT id, company_code, company_name, registration_no, base_currency, reporting_currency,
             fiscal_year_start, tax_region, timezone, active, notes, created_at, updated_at
      FROM company_profiles
      ORDER BY company_name ASC
      LIMIT 500
    `).all();
    return { ok: true, rows };
  });

  app.post("/companies/:code/deactivate", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin"])) return;
    const code = normalizeCode(req.params?.code);
    db.prepare(`UPDATE company_profiles SET active = 0, updated_at = datetime('now') WHERE company_code = ?`).run(code);
    writeAudit(db, req, { module: "entity", action: "company.deactivate", entity_type: "company_profiles", entity_id: code, payload: { code } });
    return { ok: true, company_code: code, active: 0 };
  });

  /* --------- SITE --------- */
  app.post("/sites", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "executive", "plant_manager", "site_manager"])) return;
    const site_code = normalizeCode(req.body?.site_code);
    const company_code = normalizeCode(req.body?.company_code);
    const site_name = String(req.body?.site_name || "").trim();
    if (!site_code || !company_code || !site_name) return reply.code(400).send({ error: "site_code, company_code, site_name required" });
    const company = db.prepare(`SELECT company_code FROM company_profiles WHERE company_code = ?`).get(company_code);
    if (!company) return reply.code(404).send({ error: "company not found" });
    const region = req.body?.region ? String(req.body.region).trim() : null;
    const timezone = req.body?.timezone ? String(req.body.timezone).trim() : null;
    const local_currency = req.body?.local_currency ? normalizeCcy(req.body.local_currency) : null;
    const tax_region = req.body?.tax_region ? String(req.body.tax_region).trim() : null;
    const address = req.body?.address ? String(req.body.address).trim() : null;
    const notes = req.body?.notes ? String(req.body.notes).trim() : null;
    db.prepare(`
      INSERT INTO site_profiles
        (site_code, company_code, site_name, region, timezone, local_currency, tax_region, address, active, notes, created_by, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?, datetime('now'))
      ON CONFLICT(site_code) DO UPDATE SET
        company_code = excluded.company_code,
        site_name = excluded.site_name,
        region = excluded.region,
        timezone = excluded.timezone,
        local_currency = excluded.local_currency,
        tax_region = excluded.tax_region,
        address = excluded.address,
        notes = excluded.notes,
        active = 1,
        updated_at = datetime('now')
    `).run(site_code, company_code, site_name, region, timezone, local_currency, tax_region, address, notes, getUser(req));
    writeAudit(db, req, {
      module: "entity", action: "site.upsert", entity_type: "site_profiles",
      entity_id: site_code, payload: { site_code, company_code, site_name, local_currency }
    });
    return { ok: true, site_code, company_code };
  });

  app.get("/sites", async (req) => {
    const company = req.query?.company_code ? normalizeCode(req.query.company_code) : null;
    const args = [];
    let where = "";
    if (company) { where = "WHERE s.company_code = ?"; args.push(company); }
    const rows = db.prepare(`
      SELECT s.id, s.site_code, s.company_code, s.site_name, s.region, s.timezone,
             s.local_currency, s.tax_region, s.address, s.active, s.notes, s.created_at, s.updated_at,
             c.company_name
      FROM site_profiles s
      LEFT JOIN company_profiles c ON c.company_code = s.company_code
      ${where}
      ORDER BY s.site_name ASC
      LIMIT 1000
    `).all(...args);
    return { ok: true, rows };
  });

  app.post("/sites/:code/deactivate", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin"])) return;
    const code = normalizeCode(req.params?.code);
    db.prepare(`UPDATE site_profiles SET active = 0, updated_at = datetime('now') WHERE site_code = ?`).run(code);
    writeAudit(db, req, { module: "entity", action: "site.deactivate", entity_type: "site_profiles", entity_id: code, payload: { code } });
    return { ok: true, site_code: code, active: 0 };
  });

  /* --------- CURRENCY RATES --------- */
  app.post("/currency/rates/upsert", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "executive", "finance"])) return;
    const rows = Array.isArray(req.body?.rows) ? req.body.rows : [];
    if (!rows.length) return reply.code(400).send({ error: "rows required" });
    const ins = db.prepare(`
      INSERT INTO currency_rates (from_currency, to_currency, rate, effective_date, source, notes)
      VALUES (?, ?, ?, ?, ?, ?)
      ON CONFLICT(from_currency, to_currency, effective_date) DO UPDATE SET
        rate = excluded.rate,
        source = excluded.source,
        notes = excluded.notes
    `);
    let saved = 0;
    const tx = db.transaction(() => {
      for (const r of rows) {
        const from = normalizeCcy(r?.from_currency);
        const to = normalizeCcy(r?.to_currency);
        const rate = Number(r?.rate);
        const effective = String(r?.effective_date || "").trim();
        if (!from || !to || !Number.isFinite(rate) || rate <= 0 || !effective) continue;
        ins.run(from, to, rate, effective, r?.source ? String(r.source).trim() : null, r?.notes ? String(r.notes).trim() : null);
        saved += 1;
      }
    });
    tx();
    writeAudit(db, req, { module: "entity", action: "currency.rates.upsert", entity_type: "currency_rates", entity_id: String(saved), payload: { saved } });
    return { ok: true, saved };
  });

  app.get("/currency/rates", async (req) => {
    const from = req.query?.from ? normalizeCcy(req.query.from) : null;
    const to = req.query?.to ? normalizeCcy(req.query.to) : null;
    const where = [];
    const args = [];
    if (from) { where.push("from_currency = ?"); args.push(from); }
    if (to) { where.push("to_currency = ?"); args.push(to); }
    const sql = `
      SELECT id, from_currency, to_currency, rate, effective_date, source, notes, created_at
      FROM currency_rates
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY effective_date DESC, id DESC
      LIMIT 1000
    `;
    const rows = db.prepare(sql).all(...args);
    return { ok: true, rows };
  });

  app.get("/currency/convert", async (req, reply) => {
    const from = normalizeCcy(req.query?.from);
    const to = normalizeCcy(req.query?.to);
    const amount = Number(req.query?.amount);
    const asOf = req.query?.as_of ? String(req.query.as_of).trim() : null;
    if (!from || !to || !Number.isFinite(amount)) return reply.code(400).send({ error: "from, to, amount required" });
    if (from === to) return { ok: true, from, to, amount, converted: amount, rate: 1, effective_date: asOf || null };
    const row = asOf
      ? db.prepare(`
          SELECT rate, effective_date FROM currency_rates
          WHERE from_currency = ? AND to_currency = ? AND DATE(effective_date) <= DATE(?)
          ORDER BY DATE(effective_date) DESC LIMIT 1
        `).get(from, to, asOf)
      : db.prepare(`
          SELECT rate, effective_date FROM currency_rates
          WHERE from_currency = ? AND to_currency = ?
          ORDER BY DATE(effective_date) DESC LIMIT 1
        `).get(from, to);
    if (!row) return reply.code(404).send({ error: `no rate for ${from}->${to}` });
    const rate = Number(row.rate);
    return {
      ok: true, from, to, amount,
      rate,
      effective_date: row.effective_date,
      converted: Number((amount * rate).toFixed(4))
    };
  });

  /* --------- TAX PROFILES --------- */
  app.post("/tax/profiles/upsert", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "executive", "finance"])) return;
    const tax_code = normalizeCode(req.body?.tax_code);
    const label = String(req.body?.label || "").trim();
    if (!tax_code || !label) return reply.code(400).send({ error: "tax_code and label required" });
    const rate_pct = Number(req.body?.rate_pct || 0);
    const region = req.body?.region ? String(req.body.region).trim() : null;
    const applies_to = req.body?.applies_to ? String(req.body.applies_to).trim() : null;
    const notes = req.body?.notes ? String(req.body.notes).trim() : null;
    db.prepare(`
      INSERT INTO tax_profiles (tax_code, label, rate_pct, region, active, applies_to, notes, created_by, updated_at)
      VALUES (?, ?, ?, ?, 1, ?, ?, ?, datetime('now'))
      ON CONFLICT(tax_code) DO UPDATE SET
        label = excluded.label,
        rate_pct = excluded.rate_pct,
        region = excluded.region,
        applies_to = excluded.applies_to,
        notes = excluded.notes,
        active = 1,
        updated_at = datetime('now')
    `).run(tax_code, label, rate_pct, region, applies_to, notes, getUser(req));
    writeAudit(db, req, {
      module: "entity", action: "tax.upsert", entity_type: "tax_profiles",
      entity_id: tax_code, payload: { tax_code, rate_pct, region }
    });
    return { ok: true, tax_code };
  });

  app.get("/tax/profiles", async () => {
    const rows = db.prepare(`
      SELECT id, tax_code, label, rate_pct, region, active, applies_to, notes, created_at, updated_at
      FROM tax_profiles
      ORDER BY tax_code ASC LIMIT 500
    `).all();
    return { ok: true, rows };
  });

  /* --------- ENTITY TREE (hierarchical view) --------- */
  app.get("/tree", async () => {
    const companies = db.prepare(`
      SELECT company_code, company_name, base_currency, reporting_currency, active
      FROM company_profiles ORDER BY company_name ASC
    `).all();
    const sites = db.prepare(`
      SELECT company_code, site_code, site_name, local_currency, region, active
      FROM site_profiles ORDER BY site_name ASC
    `).all();
    const byCompany = new Map();
    for (const c of companies) byCompany.set(c.company_code, { ...c, sites: [] });
    for (const s of sites) {
      const parent = byCompany.get(s.company_code);
      if (parent) parent.sites.push(s);
    }
    return { ok: true, tree: Array.from(byCompany.values()) };
  });
}
