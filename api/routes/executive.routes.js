// IRONLOG/api/routes/executive.routes.js
// Evidence Packs + Executive Command Center + Board Monthly Pack.

import { db } from "../db/client.js";
import crypto from "node:crypto";
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
  return getRoles(req).some((r) => allowed.includes(r));
}
function requireRoles(req, reply, allowed) {
  if (!hasAnyRole(req, allowed)) {
    reply.code(403).send({ error: `role '${getRole(req)}' not allowed` });
    return false;
  }
  return true;
}
function monthStart(period) {
  const [y, m] = String(period).split("-");
  return `${y}-${String(m).padStart(2, "0")}-01`;
}
function monthEnd(period) {
  const [y, m] = String(period).split("-").map((x) => Number(x));
  const end = new Date(Date.UTC(y, m, 0));
  return end.toISOString().slice(0, 10);
}
function tableExists(name) {
  try {
    const row = db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name = ?`).get(name);
    return !!row;
  } catch { return false; }
}
function safeAll(sql, args = []) {
  try { return db.prepare(sql).all(...args); } catch { return []; }
}
function safeGet(sql, args = []) {
  try { return db.prepare(sql).get(...args) || null; } catch { return null; }
}
function hashContent(obj) {
  const s = typeof obj === "string" ? obj : JSON.stringify(obj);
  return crypto.createHash("sha256").update(s).digest("hex");
}

export default async function executiveRoutes(app) {
  ensureAuditTable(db);

  db.prepare(`
    CREATE TABLE IF NOT EXISTS compliance_evidence_packs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      pack_code TEXT NOT NULL UNIQUE,
      pack_type TEXT NOT NULL,
      period TEXT NOT NULL,
      site_code TEXT,
      company_code TEXT,
      summary_json TEXT,
      sections_json TEXT,
      integrity_hash TEXT NOT NULL,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_cep_period ON compliance_evidence_packs(period)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_cep_type ON compliance_evidence_packs(pack_type)`).run();

  /* ============================================================
     EVIDENCE PACK BUILDERS
  ============================================================ */

  function buildAuditPack(period) {
    const start = monthStart(period);
    const end = monthEnd(period);
    const logs = tableExists("audit_logs") ? safeAll(`
      SELECT module, action, entity_type, entity_id, username, role, site_code, source_app, source_channel, created_at
      FROM audit_logs
      WHERE DATE(created_at) BETWEEN DATE(?) AND DATE(?)
      ORDER BY id ASC
    `, [start, end]) : [];
    const violations = tableExists("sod_violations") ? safeAll(`
      SELECT id, policy_code, username, restricted_action, precursor_action, mode, blocked, created_at, acknowledged_at
      FROM sod_violations
      WHERE DATE(created_at) BETWEEN DATE(?) AND DATE(?)
      ORDER BY id ASC
    `, [start, end]) : [];
    const periodLock = tableExists("finance_period_locks") ? safeGet(`SELECT * FROM finance_period_locks WHERE period = ?`, [period]) : null;
    const summary = {
      audit_count: logs.length,
      violations_count: violations.length,
      period_locked: periodLock ? String(periodLock.status || "open") : "open",
    };
    return { summary, sections: { audit_logs: logs, sod_violations: violations, period_lock: periodLock } };
  }

  function buildProcurementPack(period) {
    const start = monthStart(period);
    const end = monthEnd(period);
    const pos = tableExists("procurement_purchase_orders") ? safeAll(`
      SELECT po.id, po.po_number, po.status, po.order_date, po.supplier_id, po.total_amount, po.currency, po.site_code
      FROM procurement_purchase_orders po
      WHERE DATE(COALESCE(po.order_date, po.created_at)) BETWEEN DATE(?) AND DATE(?)
      ORDER BY po.id ASC
    `, [start, end]) : [];
    const receipts = tableExists("procurement_goods_receipts") ? safeAll(`
      SELECT id, receipt_number, receipt_date, po_id, status, received_by, location_code
      FROM procurement_goods_receipts
      WHERE DATE(receipt_date) BETWEEN DATE(?) AND DATE(?)
      ORDER BY id ASC
    `, [start, end]) : [];
    const invoices = tableExists("procurement_invoices") ? safeAll(`
      SELECT id, invoice_number, invoice_date, po_id, status, subtotal, tax, total, currency
      FROM procurement_invoices
      WHERE DATE(invoice_date) BETWEEN DATE(?) AND DATE(?)
      ORDER BY id ASC
    `, [start, end]) : [];
    const exceptions = tableExists("procurement_match_exceptions") ? safeAll(`
      SELECT id, po_id, invoice_id, receipt_id, exception_type, severity, status, created_at, resolved_at
      FROM procurement_match_exceptions
      WHERE DATE(created_at) BETWEEN DATE(?) AND DATE(?)
      ORDER BY id ASC
    `, [start, end]) : [];
    const runs = tableExists("finance_posting_runs") ? safeAll(`
      SELECT id, run_number, period, status, total_debit, total_credit, posted_at, posted_reference
      FROM finance_posting_runs WHERE period = ? ORDER BY id ASC
    `, [period]) : [];
    const summary = {
      po_count: pos.length,
      grn_count: receipts.length,
      invoice_count: invoices.length,
      exceptions_count: exceptions.length,
      posting_runs: runs.length,
    };
    return { summary, sections: { purchase_orders: pos, goods_receipts: receipts, invoices, exceptions, posting_runs: runs } };
  }

  function buildHsePack(period) {
    const start = monthStart(period);
    const end = monthEnd(period);
    const breakdowns = tableExists("breakdowns") ? safeAll(`
      SELECT b.id, b.asset_id, b.reported_at, b.status, b.severity, b.description
      FROM breakdowns b
      WHERE DATE(COALESCE(b.reported_at, b.created_at, b.updated_at)) BETWEEN DATE(?) AND DATE(?)
      ORDER BY b.id ASC
    `, [start, end]) : [];
    const dtLogs = tableExists("breakdown_downtime_logs") ? safeAll(`
      SELECT id, breakdown_id, log_date, hours_down, notes
      FROM breakdown_downtime_logs
      WHERE log_date BETWEEN DATE(?) AND DATE(?)
      ORDER BY id ASC
    `, [start, end]) : [];
    const totalDown = dtLogs.reduce((s, r) => s + Number(r.hours_down || 0), 0);
    const severe = breakdowns.filter((b) => String(b.severity || "").toLowerCase() === "severe").length;
    const summary = {
      breakdowns_count: breakdowns.length,
      severe_breakdowns: severe,
      total_downtime_hours: Number(totalDown.toFixed(2)),
    };
    return { summary, sections: { breakdowns, downtime_logs: dtLogs } };
  }

  function buildEvidencePack(pack_type, period, user, site_code, company_code) {
    let data = { summary: {}, sections: {} };
    if (pack_type === "audit") data = buildAuditPack(period);
    else if (pack_type === "procurement") data = buildProcurementPack(period);
    else if (pack_type === "hse") data = buildHsePack(period);
    else throw new Error(`unknown pack_type '${pack_type}'`);
    const pack_code = `EVID-${pack_type.toUpperCase()}-${period}-${Date.now()}`;
    const integrity_hash = hashContent({ pack_type, period, sections: data.sections });
    const r = db.prepare(`
      INSERT INTO compliance_evidence_packs
        (pack_code, pack_type, period, site_code, company_code, summary_json, sections_json, integrity_hash, created_by)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(pack_code, pack_type, period, site_code || null, company_code || null,
      JSON.stringify(data.summary), JSON.stringify(data.sections), integrity_hash, user);
    return { id: Number(r.lastInsertRowid), pack_code, integrity_hash, summary: data.summary };
  }

  app.post("/evidence-packs/build", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "executive", "finance"])) return;
    const pack_type = String(req.body?.pack_type || "").trim().toLowerCase();
    const period = String(req.body?.period || "").trim();
    if (!["audit", "procurement", "hse"].includes(pack_type)) return reply.code(400).send({ error: "pack_type must be audit|procurement|hse" });
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period (YYYY-MM) required" });
    const site_code = req.body?.site_code ? String(req.body.site_code).trim() : null;
    const company_code = req.body?.company_code ? String(req.body.company_code).trim() : null;
    try {
      const res = buildEvidencePack(pack_type, period, getUser(req), site_code, company_code);
      writeAudit(db, req, { module: "executive", action: "evidence.build", entity_type: "compliance_evidence_packs", entity_id: res.pack_code, payload: { pack_type, period } });
      return { ok: true, ...res };
    } catch (e) { return reply.code(500).send({ error: e.message }); }
  });

  app.get("/evidence-packs", async (req) => {
    const period = req.query?.period ? String(req.query.period).trim() : null;
    const pack_type = req.query?.pack_type ? String(req.query.pack_type).trim() : null;
    const where = [];
    const args = [];
    if (period) { where.push("period = ?"); args.push(period); }
    if (pack_type) { where.push("pack_type = ?"); args.push(pack_type); }
    const sql = `
      SELECT id, pack_code, pack_type, period, site_code, company_code, summary_json, integrity_hash, created_by, created_at
      FROM compliance_evidence_packs
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY id DESC LIMIT 500
    `;
    const rows = db.prepare(sql).all(...args).map((r) => ({
      ...r,
      summary: (() => { try { return JSON.parse(r.summary_json || "{}"); } catch { return {}; } })()
    }));
    return { ok: true, rows };
  });

  app.get("/evidence-packs/:id", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const row = db.prepare(`SELECT * FROM compliance_evidence_packs WHERE id = ?`).get(id);
    if (!row) return reply.code(404).send({ error: "pack not found" });
    const verify = hashContent({ pack_type: row.pack_type, period: row.period, sections: JSON.parse(row.sections_json || "{}") });
    const integrity_ok = verify === row.integrity_hash;
    return {
      ok: true,
      pack: {
        ...row,
        summary: (() => { try { return JSON.parse(row.summary_json || "{}"); } catch { return {}; } })(),
        sections: (() => { try { return JSON.parse(row.sections_json || "{}"); } catch { return {}; } })(),
        integrity_ok,
      }
    };
  });

  /* ============================================================
     EXECUTIVE COMMAND CENTER + BOARD MONTHLY PACK
  ============================================================ */

  function financialSummary(period) {
    const runs = tableExists("finance_posting_runs") ? safeAll(`
      SELECT status, COUNT(*) AS c, COALESCE(SUM(total_debit), 0) AS debit, COALESCE(SUM(total_credit), 0) AS credit
      FROM finance_posting_runs WHERE period = ? GROUP BY status
    `, [period]) : [];
    const budgets = tableExists("finance_budgets_monthly") ? safeAll(`
      SELECT category, SUM(budget_amount) AS budget FROM finance_budgets_monthly WHERE period = ? GROUP BY category
    `, [period]) : [];
    const forecasts = tableExists("finance_forecasts_monthly") ? safeAll(`
      SELECT category, SUM(forecast_amount) AS forecast FROM finance_forecasts_monthly WHERE period = ? GROUP BY category
    `, [period]) : [];
    return { posting_runs: runs, budgets, forecasts };
  }

  function operationalSummary(period) {
    const start = monthStart(period);
    const end = monthEnd(period);
    const dh = safeAll(`
      SELECT COALESCE(SUM(hours_run), 0) AS run_hours, COUNT(DISTINCT asset_id) AS used_assets_total
      FROM daily_hours WHERE log_date BETWEEN DATE(?) AND DATE(?) AND hours_run > 0
    `, [start, end]);
    const down = safeGet(`
      SELECT COALESCE(SUM(hours_down), 0) AS h FROM breakdown_downtime_logs WHERE log_date BETWEEN DATE(?) AND DATE(?)
    `, [start, end]);
    const breakdownCount = safeGet(`
      SELECT COUNT(*) AS c FROM breakdowns WHERE DATE(COALESCE(reported_at, created_at, updated_at)) BETWEEN DATE(?) AND DATE(?)
    `, [start, end]);
    return {
      run_hours: Number(dh[0]?.run_hours || 0),
      used_assets_total: Number(dh[0]?.used_assets_total || 0),
      downtime_hours: Number(down?.h || 0),
      breakdowns: Number(breakdownCount?.c || 0),
    };
  }

  function integrationHealth() {
    if (!tableExists("integration_jobs")) return { by_status: [], dead_letter_open: 0, oldest_queued: null };
    const byStatus = safeAll(`SELECT status, COUNT(*) AS c FROM integration_jobs GROUP BY status`);
    const deadLetterOpen = tableExists("integration_dead_letter")
      ? Number(safeGet(`SELECT COUNT(*) AS c FROM integration_dead_letter WHERE acknowledged_at IS NULL`)?.c || 0)
      : 0;
    const oldestQueued = safeGet(`SELECT id, connector_key, scheduled_at FROM integration_jobs WHERE status = 'queued' ORDER BY scheduled_at ASC LIMIT 1`);
    return { by_status: byStatus, dead_letter_open: deadLetterOpen, oldest_queued: oldestQueued };
  }

  function governanceHealth(period) {
    if (!tableExists("sod_violations")) return { open: 0, acknowledged: 0, by_policy: [] };
    const start = monthStart(period);
    const end = monthEnd(period);
    const byPolicy = safeAll(`
      SELECT policy_code, COUNT(*) AS c FROM sod_violations
      WHERE DATE(created_at) BETWEEN DATE(?) AND DATE(?)
      GROUP BY policy_code
    `, [start, end]);
    const open = Number(safeGet(`SELECT COUNT(*) AS c FROM sod_violations WHERE acknowledged_at IS NULL`)?.c || 0);
    const ack = Number(safeGet(`SELECT COUNT(*) AS c FROM sod_violations WHERE acknowledged_at IS NOT NULL`)?.c || 0);
    return { open, acknowledged: ack, by_policy: byPolicy };
  }

  app.get("/command-center", async (req, reply) => {
    const period = String(req.query?.period || "").trim();
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period (YYYY-MM) required" });
    return {
      ok: true,
      period,
      financial: financialSummary(period),
      operational: operationalSummary(period),
      integrations: integrationHealth(),
      governance: governanceHealth(period),
    };
  });

  function buildNarrative(period) {
    const fin = financialSummary(period);
    const ops = operationalSummary(period);
    const gov = governanceHealth(period);
    const integ = integrationHealth();
    const totalBudget = fin.budgets.reduce((s, r) => s + Number(r.budget || 0), 0);
    const totalForecast = fin.forecasts.reduce((s, r) => s + Number(r.forecast || 0), 0);
    const availability = ops.run_hours + ops.downtime_hours > 0
      ? ((ops.run_hours) / (ops.run_hours + ops.downtime_hours)) * 100
      : null;
    const bullets = [];
    bullets.push(`Period ${period}: tracked ${ops.breakdowns} breakdowns, ${ops.downtime_hours.toFixed(1)}h downtime against ${ops.run_hours.toFixed(1)}h run hours.`);
    if (availability != null) bullets.push(`Fleet availability proxy: ${availability.toFixed(1)}%.`);
    if (totalBudget > 0) bullets.push(`Monthly budget set at ${totalBudget.toFixed(2)}; rolling forecast ${totalForecast.toFixed(2)}.`);
    if (gov.open > 0) bullets.push(`Governance: ${gov.open} open SoD violations requiring review.`);
    else bullets.push(`Governance: no open SoD violations outstanding.`);
    if (integ.dead_letter_open > 0) bullets.push(`Integrations: ${integ.dead_letter_open} dead-letter items awaiting triage.`);
    else bullets.push(`Integrations: queue healthy, no dead-letter backlog.`);
    return { bullets };
  }

  app.get("/board-pack", async (req, reply) => {
    const period = String(req.query?.period || "").trim();
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period (YYYY-MM) required" });
    const financial = financialSummary(period);
    const operational = operationalSummary(period);
    const integrations = integrationHealth();
    const governance = governanceHealth(period);
    const narrative = buildNarrative(period);
    return { ok: true, period, narrative, financial, operational, integrations, governance };
  });

  app.get("/board-pack/export.xlsx", async (req, reply) => {
    const period = String(req.query?.period || "").trim();
    if (!/^\d{4}-\d{2}$/.test(period)) return reply.code(400).send({ error: "period (YYYY-MM) required" });
    const financial = financialSummary(period);
    const operational = operationalSummary(period);
    const integrations = integrationHealth();
    const governance = governanceHealth(period);
    const narrative = buildNarrative(period);
    const ExcelJS = (await import("exceljs")).default;
    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const wsN = wb.addWorksheet("Narrative");
    wsN.columns = [{ header: "Highlights", key: "line", width: 120 }];
    for (const b of narrative.bullets) wsN.addRow({ line: b });
    wsN.getRow(1).font = { bold: true };

    const wsFin = wb.addWorksheet("Financial");
    wsFin.columns = [
      { header: "Section", key: "sec", width: 20 },
      { header: "Key", key: "k", width: 30 },
      { header: "Value", key: "v", width: 30 },
    ];
    financial.posting_runs.forEach((r) => wsFin.addRow({ sec: "posting_run_status", k: r.status, v: `count=${r.c} debit=${Number(r.debit || 0).toFixed(2)} credit=${Number(r.credit || 0).toFixed(2)}` }));
    financial.budgets.forEach((r) => wsFin.addRow({ sec: "budget", k: r.category, v: Number(r.budget || 0).toFixed(2) }));
    financial.forecasts.forEach((r) => wsFin.addRow({ sec: "forecast", k: r.category, v: Number(r.forecast || 0).toFixed(2) }));
    wsFin.getRow(1).font = { bold: true };

    const wsOps = wb.addWorksheet("Operational");
    wsOps.columns = [{ header: "Metric", key: "k", width: 24 }, { header: "Value", key: "v", width: 24 }];
    Object.entries(operational).forEach(([k, v]) => wsOps.addRow({ k, v }));
    wsOps.getRow(1).font = { bold: true };

    const wsInt = wb.addWorksheet("Integrations");
    wsInt.columns = [{ header: "Status", key: "s", width: 20 }, { header: "Count", key: "c", width: 12 }];
    integrations.by_status.forEach((r) => wsInt.addRow({ s: r.status, c: Number(r.c || 0) }));
    wsInt.addRow({ s: "dead_letter_open", c: integrations.dead_letter_open });
    wsInt.getRow(1).font = { bold: true };

    const wsGov = wb.addWorksheet("Governance");
    wsGov.columns = [{ header: "Policy", key: "p", width: 30 }, { header: "Violations", key: "c", width: 14 }];
    governance.by_policy.forEach((r) => wsGov.addRow({ p: r.policy_code, c: Number(r.c || 0) }));
    wsGov.addRow({ p: "(open total)", c: governance.open });
    wsGov.addRow({ p: "(acknowledged total)", c: governance.acknowledged });
    wsGov.getRow(1).font = { bold: true };

    const buffer = await wb.xlsx.writeBuffer();
    reply.header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    reply.header("Content-Disposition", `attachment; filename="board-pack-${period}.xlsx"`);
    return reply.send(Buffer.from(buffer));
  });
}
