import { db } from "../db/client.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function toDateStart(daysBack = 30) {
  const d = new Date();
  d.setDate(d.getDate() - Math.max(0, Number(daysBack || 0)));
  return d.toISOString().slice(0, 10);
}
function getSiteCode(req) {
  return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
}

export default async function qualityRoutes(app) {
  app.get("/", async (req, reply) => {
    const site_code = getSiteCode(req);
    const from = String(req.query?.from || "").trim() || toDateStart(30);
    const to = String(req.query?.to || "").trim() || new Date().toISOString().slice(0, 10);
    if (!isDate(from) || !isDate(to)) {
      return reply.code(400).send({ error: "from and to must be YYYY-MM-DD" });
    }

    const dailyRows = db.prepare(`
      SELECT
        dh.work_date,
        a.asset_code,
        a.asset_name,
        COALESCE(dh.input_unit, 'hours') AS input_unit,
        COALESCE(dh.is_used, 0) AS is_used,
        COALESCE(dh.scheduled_hours, 0) AS scheduled_hours,
        COALESCE(dh.hours_run, 0) AS hours_run
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date BETWEEN ? AND ?
        AND a.active = 1
    `).all(from, to);
    const dailyIssues = [];
    for (const r of dailyRows) {
      const isUsed = Number(r.is_used || 0) === 1;
      const unit = String(r.input_unit || "hours").toLowerCase();
      const sched = Number(r.scheduled_hours || 0);
      const run = Number(r.hours_run || 0);
      if (!isUsed && run > 0) {
        dailyIssues.push({ type: "daily_standby_with_run", severity: "high", date: r.work_date, asset_code: r.asset_code, details: `Standby row has run ${run}` });
      }
      if (isUsed && sched <= 0) {
        dailyIssues.push({ type: "daily_production_no_schedule", severity: "high", date: r.work_date, asset_code: r.asset_code, details: "Production row with scheduled 0" });
      }
      if (isUsed && run <= 0) {
        dailyIssues.push({ type: "daily_production_no_run", severity: "medium", date: r.work_date, asset_code: r.asset_code, details: "Production row with run 0" });
      }
      if (unit === "hours" && run > 24) {
        dailyIssues.push({ type: "daily_hours_over_24", severity: "high", date: r.work_date, asset_code: r.asset_code, details: `Hours unit run >24 (${run})` });
      }
    }

    const dispatchPodGaps = db.prepare(`
      SELECT id AS trip_id, op_date, truck_reg, client_name, status, pod_ref
      FROM dispatch_trips
      WHERE op_date BETWEEN ? AND ?
        AND COALESCE(site_code, 'main') = ?
        AND status = 'delivered'
        AND TRIM(COALESCE(pod_ref, '')) = ''
      ORDER BY op_date DESC, id DESC
      LIMIT 500
    `).all(from, to, site_code).map((r) => ({
      type: "dispatch_delivered_no_pod",
      severity: "high",
      date: r.op_date,
      entity_id: r.trip_id,
      details: `Trip ${r.trip_id} delivered without POD`,
      truck_reg: r.truck_reg,
      client_name: r.client_name,
    }));

    const exceptionsOpen = db.prepare(`
      SELECT e.id, t.op_date, e.trip_id, e.exception_type, e.severity, e.owner_name, e.created_at
      FROM dispatch_exceptions e
      JOIN dispatch_trips t ON t.id = e.trip_id
      WHERE t.op_date BETWEEN ? AND ?
        AND COALESCE(t.site_code, 'main') = ?
        AND e.status = 'open'
      ORDER BY e.id DESC
      LIMIT 500
    `).all(from, to, site_code).map((r) => ({
      type: "dispatch_exception_open",
      severity: String(r.severity || "medium"),
      date: r.op_date,
      entity_id: r.id,
      details: `Exception ${r.exception_type} still open (trip ${r.trip_id})`,
      owner_name: r.owner_name || "",
      created_at: r.created_at,
    }));

    const approvalsPending = db.prepare(`
      SELECT id, module, action, requester, created_at
      FROM approval_requests
      WHERE status = 'pending'
      ORDER BY id DESC
      LIMIT 500
    `).all().map((r) => ({
      type: "approval_pending",
      severity: "medium",
      date: String(r.created_at || "").slice(0, 10),
      entity_id: r.id,
      details: `${r.module}.${r.action} pending`,
      requester: r.requester || "",
      created_at: r.created_at,
    }));

    const rows = [
      ...dailyIssues,
      ...dispatchPodGaps,
      ...exceptionsOpen,
      ...approvalsPending,
    ].sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")));

    const summary = {
      total: rows.length,
      daily_issues: dailyIssues.length,
      dispatch_pod_gaps: dispatchPodGaps.length,
      exceptions_open: exceptionsOpen.length,
      approvals_pending: approvalsPending.length,
      high: rows.filter((r) => String(r.severity).toLowerCase() === "high").length,
      medium: rows.filter((r) => String(r.severity).toLowerCase() === "medium").length,
      low: rows.filter((r) => String(r.severity).toLowerCase() === "low").length,
    };

    return { ok: true, from, to, summary, rows: rows.slice(0, 1000) };
  });
}

