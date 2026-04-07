import { generateIronmindReport, getIronmindHistory, getLatestIronmindReport } from "../utils/ironmind.js";
import { db } from "../db/client.js";

function toBool(v) {
  const s = String(v || "").trim().toLowerCase();
  return s === "1" || s === "true" || s === "yes" || s === "y";
}

export default async function ironmindRoutes(app) {
  function isDate(v) {
    return /^\d{4}-\d{2}-\d{2}$/.test(String(v || "").trim());
  }
  function todayYmd() {
    return new Date().toISOString().slice(0, 10);
  }
  function parseQuestionDates(question, fallbackDate) {
    const q = String(question || "");
    const matches = q.match(/\b\d{4}-\d{2}-\d{2}\b/g) || [];
    if (matches.length >= 2) return { start: matches[0], end: matches[1] };
    if (matches.length === 1) return { start: matches[0], end: matches[0] };
    return { start: fallbackDate, end: fallbackDate };
  }
  function parseAssetCode(question) {
    const q = String(question || "").toUpperCase();
    const m = q.match(/\b[A-Z0-9]{2,}[A-Z][0-9A-Z-]*\b/g) || [];
    const deny = new Set(["PLEASE", "DOWNTIME", "SELECTED", "TIME", "FROM", "TO", "AND", "THE", "FOR"]);
    return (m.find((x) => !deny.has(x)) || "").trim();
  }

  app.get("/history", async (req, reply) => {
    try {
      const reportType = String(req.query?.report_type || "daily_admin").trim() || "daily_admin";
      const days = Math.max(1, Math.min(60, Number(req.query?.days || 7)));
      const reports = getIronmindHistory({ reportType, limit: days });
      return reply.send({ ok: true, reports });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.get("/latest", async (req, reply) => {
    try {
      const reportType = String(req.query?.report_type || "daily_admin").trim() || "daily_admin";
      const row = getLatestIronmindReport(reportType);
      if (!row) {
        return reply.send({ ok: true, report: null });
      }
      return reply.send({ ok: true, report: row });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/run", async (req, reply) => {
    try {
      const body = req.body || {};
      const reportDate = String(body.report_date || "").trim() || undefined;
      const reportType = String(body.report_type || "daily_admin").trim() || "daily_admin";
      const force = toBool(body.force);
      const contextNotes = String(body.context_notes || "").trim();
      const detailMode = toBool(body.detail_mode);

      const report = await generateIronmindReport({
        reportDate,
        reportType,
        force,
        contextNotes: contextNotes || undefined,
        detailMode,
      });
      return reply.send({ ok: true, report });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/ask", async (req, reply) => {
    try {
      const body = req.body || {};
      const question = String(body.question || "").trim();
      if (!question) return reply.code(400).send({ ok: false, error: "question is required" });
      const fallbackDate = isDate(body.date) ? String(body.date) : todayYmd();
      const parsed = parseQuestionDates(question, fallbackDate);
      const start = isDate(body.start) ? String(body.start) : parsed.start;
      const end = isDate(body.end) ? String(body.end) : parsed.end;
      const assetCode = String(body.asset_code || parseAssetCode(question)).trim().toUpperCase();
      if (!assetCode) {
        return reply.send({
          ok: true,
          short_answer: "Please include an asset code, for example: G01AM from 2026-02-14 to 2026-04-07.",
        });
      }
      const row = db.prepare(`
        SELECT
          a.asset_code,
          COUNT(DISTINCT l.log_date) AS logged_days,
          COALESCE(SUM(l.hours_down), 0) AS downtime_hours,
          COALESCE(MIN(l.log_date), '') AS first_log_date,
          COALESCE(MAX(l.log_date), '') AS last_log_date
        FROM assets a
        LEFT JOIN breakdowns b ON b.asset_id = a.id
        LEFT JOIN breakdown_downtime_logs l
          ON l.breakdown_id = b.id
         AND l.log_date BETWEEN ? AND ?
        WHERE UPPER(a.asset_code) = UPPER(?)
      `).get(start, end, assetCode);
      if (!row?.asset_code) {
        return reply.send({ ok: true, short_answer: `No asset found for code ${assetCode}.` });
      }
      const breakdowns = db.prepare(`
        SELECT COUNT(*) AS c
        FROM breakdowns b
        JOIN assets a ON a.id = b.asset_id
        WHERE UPPER(a.asset_code) = UPPER(?)
          AND b.breakdown_date BETWEEN ? AND ?
      `).get(assetCode, start, end);
      const openBreakdown = db.prepare(`
        SELECT b.id, b.breakdown_date
        FROM breakdowns b
        JOIN assets a ON a.id = b.asset_id
        WHERE UPPER(a.asset_code) = UPPER(?)
          AND b.status = 'OPEN'
        ORDER BY b.id DESC
        LIMIT 1
      `).get(assetCode);

      const hours = Number(row.downtime_hours || 0);
      const short = `${assetCode}: ${hours.toFixed(1)}h downtime from ${start} to ${end} across ${Number(row.logged_days || 0)} logged day(s).`;
      return reply.send({
        ok: true,
        short_answer: short,
        details: {
          asset_code: assetCode,
          start,
          end,
          downtime_hours: hours,
          logged_days: Number(row.logged_days || 0),
          first_log_date: row.first_log_date || null,
          last_log_date: row.last_log_date || null,
          breakdowns_in_range: Number(breakdowns?.c || 0),
          has_open_breakdown: Boolean(openBreakdown),
          open_breakdown_date: openBreakdown?.breakdown_date || null,
        },
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
}
