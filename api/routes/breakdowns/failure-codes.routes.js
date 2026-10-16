// IRONLOG/api/routes/breakdowns/failure-codes.routes.js — failure codes on breakdowns (utils/failureCodes.js).
// Registered by routes/breakdowns.routes.js under /api/breakdowns.
//   GET  /failure-codes                 the lists (?all=1 includes entries switched off)
//   PUT  /failure-codes/:kind           { rows: [{ code, label_en, label_pt, active }] } — admins and foremen
//   GET  /failure-codes/pending         this site's breakdowns of the last 90 days whose codes are not confirmed
//   GET  /failure-codes/summary         ?start&end — coded breakdowns counted by component, problem, cause
//   PUT  /:id/failure-codes             { component, problem, cause, action } — foremen confirm
import { db } from "../../db/client.js";
import { writeAudit } from "../../utils/audit.js";
import { getUser, requireAnyRole } from "../../utils/request.js";
import { requestSite, siteSql } from "../../utils/sites.js";
import {
  FAILURE_KINDS, breakdownFailureCodes, componentFromText, ensureFailureCodeTables, listFailureCodes,
  readFailureCodes, saveFailureCodeList, setBreakdownFailureCodes,
} from "../../utils/failureCodes.js";

// Foremen and managers confirm codes; admins and foremen (plant admin, plant manager) keep the lists.
export const FC_CONFIRM_ROLES = ["admin", "supervisor", "workshop_admin", "plant_manager", "site_manager"];
export const FC_LIST_ROLES = ["admin", "supervisor", "workshop_admin", "plant_manager"];

const isYmd = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ""));

export default function registerFailureCodeRoutes(app) {
  ensureFailureCodeTables(db);
  const audit = (req, action, id, payload) => {
    try {
      writeAudit(db, req, { module: "breakdowns", action, entity_type: "breakdown", entity_id: String(id), payload });
    } catch {}
  };

  // The lists are the company's, the same on every site.
  app.get("/failure-codes", async (req) => ({ ok: true, kinds: FAILURE_KINDS, lists: listFailureCodes(db, { all: String(req.query?.all || "") === "1" }) }));

  app.put("/failure-codes/:kind", async (req, reply) => {
    if (!requireAnyRole(req, reply, FC_LIST_ROLES)) return;
    const kind = String(req.params.kind || "");
    const before = listFailureCodes(db, { all: true })[kind] || null;
    const r = saveFailureCodeList(db, kind, req.body?.rows, getUser(req));
    if (!r.ok) return reply.code(400).send({ ok: false, error: r.errors.join(" "), errors: r.errors });
    try {
      writeAudit(db, req, { module: "breakdowns", action: "failure_codes.list", entity_type: "failure_codes", entity_id: kind, before, after: r.list });
    } catch {}
    return r;
  });

  app.get("/failure-codes/pending", async (req) => {
    const rows = db.prepare(`
      SELECT b.id, b.breakdown_date, b.status, b.description, b.component, b.downtime_total_hours,
             b.fc_component, b.fc_problem, b.fc_cause, b.fc_action, b.fc_suggested_by, b.fc_suggested_at,
             a.asset_code, a.asset_name, a.category
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      WHERE ${siteSql("a.site_code")} = ?
        AND b.fc_confirmed_at IS NULL
        AND b.breakdown_date >= date('now', '-90 days')
      ORDER BY CASE WHEN UPPER(COALESCE(b.status, '')) = 'CLOSED' THEN 0 ELSE 1 END, b.breakdown_date DESC, b.id DESC
      LIMIT 200
    `).all(requestSite(req));
    return {
      ok: true,
      rows: rows.map((r) => ({
        ...r,
        downtime_total_hours: Number(r.downtime_total_hours || 0),
        // What the "what failed" text already says, when the technician has not picked a component.
        guess_component: r.fc_component ? null : componentFromText(db, r.component),
      })),
    };
  });

  app.get("/failure-codes/summary", async (req) => {
    const start = isYmd(req.query?.start) ? req.query.start : null;
    const end = isYmd(req.query?.end) ? req.query.end : null;
    const where = [`${siteSql("a.site_code")} = ?`];
    const params = [requestSite(req)];
    if (start) { where.push("b.breakdown_date >= ?"); params.push(start); }
    if (end) { where.push("b.breakdown_date <= ?"); params.push(end); }
    const base = `FROM breakdowns b JOIN assets a ON a.id = b.asset_id WHERE ${where.join(" AND ")}`;
    const totals = db.prepare(`
      SELECT COUNT(*) AS breakdowns,
             SUM(CASE WHEN b.fc_component IS NOT NULL THEN 1 ELSE 0 END) AS coded,
             SUM(CASE WHEN b.fc_confirmed_at IS NOT NULL THEN 1 ELSE 0 END) AS confirmed
      ${base}
    `).get(...params);
    const lists = listFailureCodes(db, { all: true });
    const label = (kind, code) => lists[kind].find((x) => x.code === code)?.label_en || code;
    const by = (kind) => db.prepare(`
      SELECT b.fc_${kind} AS code, COUNT(*) AS breakdowns, ROUND(SUM(COALESCE(b.downtime_total_hours, 0)), 1) AS downtime_hours
      ${base} AND b.fc_${kind} IS NOT NULL
      GROUP BY b.fc_${kind}
      ORDER BY breakdowns DESC, downtime_hours DESC
    `).all(...params).map((r) => ({ ...r, label: label(kind, r.code), downtime_hours: Number(r.downtime_hours || 0) }));
    return {
      ok: true,
      start, end,
      totals: { breakdowns: totals.breakdowns || 0, coded: totals.coded || 0, confirmed: totals.confirmed || 0 },
      by_component: by("component"),
      by_problem: by("problem"),
      by_cause: by("cause"),
    };
  });

  app.put("/:id/failure-codes", async (req, reply) => {
    // site: the breakdown number in the address is checked by the site guard (utils/siteGuard.js).
    if (!requireAnyRole(req, reply, FC_CONFIRM_ROLES)) return;
    const id = Number(req.params.id);
    const b = db.prepare(`SELECT id, fc_component, fc_problem, fc_cause, fc_action FROM breakdowns WHERE id = ?`).get(id);
    if (!b) return reply.code(404).send({ ok: false, error: "Breakdown not found" });
    const r = readFailureCodes(db, req.body || {}, b);
    if (!r.ok) return reply.code(400).send({ ok: false, error: r.errors.join(" "), errors: r.errors });
    const before = breakdownFailureCodes(db, id);
    const codes = setBreakdownFailureCodes(db, id, r.values, getUser(req));
    audit(req, codes.confirmed_at ? "failure_codes.confirm" : "failure_codes.save", id, { before, after: codes });
    return { ok: true, breakdown_id: id, codes };
  });
}
