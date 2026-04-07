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
    const ym = q.match(/\b(\d{4})-(\d{2})\b/);
    if (ym) {
      const y = Number(ym[1]);
      const m = Number(ym[2]);
      if (y >= 2000 && m >= 1 && m <= 12) {
        const start = `${String(y).padStart(4, "0")}-${String(m).padStart(2, "0")}-01`;
        const endDate = new Date(Date.UTC(y, m, 0));
        const end = `${endDate.getUTCFullYear()}-${String(endDate.getUTCMonth() + 1).padStart(2, "0")}-${String(endDate.getUTCDate()).padStart(2, "0")}`;
        return { start, end };
      }
    }
    const monthMap = {
      january: 1, jan: 1,
      february: 2, feb: 2,
      march: 3, mar: 3,
      april: 4, apr: 4,
      may: 5,
      june: 6, jun: 6,
      july: 7, jul: 7,
      august: 8, aug: 8,
      september: 9, sep: 9, sept: 9,
      october: 10, oct: 10,
      november: 11, nov: 11,
      december: 12, dec: 12,
    };
    const lower = q.toLowerCase();
    const monthKey = Object.keys(monthMap).find((k) => new RegExp(`\\b${k}\\b`, "i").test(lower));
    if (monthKey) {
      const fallbackYear = Number(String(fallbackDate).slice(0, 4)) || new Date().getUTCFullYear();
      const yearMatch = lower.match(/\b(20\d{2})\b/);
      const year = yearMatch ? Number(yearMatch[1]) : fallbackYear;
      const m = monthMap[monthKey];
      const start = `${String(year).padStart(4, "0")}-${String(m).padStart(2, "0")}-01`;
      const endDate = new Date(Date.UTC(year, m, 0));
      const end = `${endDate.getUTCFullYear()}-${String(endDate.getUTCMonth() + 1).padStart(2, "0")}-${String(endDate.getUTCDate()).padStart(2, "0")}`;
      return { start, end };
    }
    return { start: fallbackDate, end: fallbackDate };
  }
  function parseAssetCode(question) {
    const q = String(question || "").toUpperCase();
    const m = q.match(/\b[A-Z0-9]{2,}[A-Z][0-9A-Z-]*\b/g) || [];
    const deny = new Set(["PLEASE", "DOWNTIME", "SELECTED", "TIME", "FROM", "TO", "AND", "THE", "FOR", "FUEL", "USAGE", "RECURRING", "FAILURES", "PM", "OVERDUE", "RISK"]);
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
      const qLower = question.toLowerCase();
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

      const asksFuel = qLower.includes("fuel") || qLower.includes("l/hr") || qLower.includes("km/l");
      const asksRecurring = qLower.includes("recurring") || qLower.includes("repeat") || qLower.includes("failures");
      const asksPm = qLower.includes("pm") || qLower.includes("overdue") || qLower.includes("maintenance");

      if (asksFuel) {
        const hasBaseline = db.prepare(`PRAGMA table_info(assets)`).all().some((c) => String(c.name) === "baseline_fuel_l_per_hour");
        const fuelRow = db.prepare(`
          SELECT
            a.asset_code,
            COALESCE(SUM(fl.liters), 0) AS liters,
            COALESCE(SUM(fl.hours_run), 0) AS run_hours,
            COALESCE(a.baseline_fuel_l_per_hour, 0) AS baseline_lph
          FROM assets a
          LEFT JOIN fuel_logs fl ON fl.asset_id = a.id AND fl.log_date BETWEEN ? AND ?
          WHERE UPPER(a.asset_code) = UPPER(?)
          GROUP BY a.id
        `).get(start, end, assetCode);
        if (!fuelRow) {
          return reply.send({ ok: true, short_answer: `No fuel records found for ${assetCode} in ${start} to ${end}.` });
        }
        const liters = Number(fuelRow.liters || 0);
        const runHours = Number(fuelRow.run_hours || 0);
        const actualLph = runHours > 0 ? liters / runHours : 0;
        const baseline = Number(fuelRow.baseline_lph || 0);
        const overPct = baseline > 0 ? ((actualLph / baseline) - 1) * 100 : null;
        const short = baseline > 0
          ? `${assetCode}: fuel ${actualLph.toFixed(2)} L/h vs baseline ${baseline.toFixed(2)} L/h (${overPct >= 0 ? "+" : ""}${overPct.toFixed(1)}%) from ${start} to ${end}.`
          : `${assetCode}: fuel ${actualLph.toFixed(2)} L/h from ${start} to ${end}. Baseline not configured.`;
        return reply.send({
          ok: true,
          short_answer: short,
          details: {
            asset_code: assetCode,
            start,
            end,
            liters,
            run_hours: runHours,
            actual_lph: Number(actualLph.toFixed(3)),
            baseline_lph: baseline,
            over_pct: overPct == null ? null : Number(overPct.toFixed(2)),
            baseline_configured: hasBaseline && baseline > 0,
          },
        });
      }

      if (asksRecurring) {
        const recurring = db.prepare(`
          SELECT
            COUNT(DISTINCT l.breakdown_id) AS incidents,
            COALESCE(SUM(l.hours_down), 0) AS downtime_hours,
            COALESCE(MIN(l.log_date), '') AS first_log_date,
            COALESCE(MAX(l.log_date), '') AS last_log_date
          FROM breakdown_downtime_logs l
          JOIN breakdowns b ON b.id = l.breakdown_id
          JOIN assets a ON a.id = b.asset_id
          WHERE UPPER(a.asset_code) = UPPER(?)
            AND l.log_date BETWEEN ? AND ?
        `).get(assetCode, start, end);
        const incidents = Number(recurring?.incidents || 0);
        const dt = Number(recurring?.downtime_hours || 0);
        const short = `${assetCode}: ${incidents} recurring failure incident(s), ${dt.toFixed(1)}h downtime between ${start} and ${end}.`;
        return reply.send({
          ok: true,
          short_answer: short,
          details: {
            asset_code: assetCode,
            start,
            end,
            incidents,
            downtime_hours: dt,
            first_log_date: recurring?.first_log_date || null,
            last_log_date: recurring?.last_log_date || null,
          },
        });
      }

      if (asksPm) {
        const pm = db.prepare(`
          SELECT
            mp.service_name,
            (COALESCE((
              SELECT SUM(dh.hours_run)
              FROM daily_hours dh
              JOIN assets a2 ON a2.id = dh.asset_id
              WHERE dh.asset_id = mp.asset_id
                AND dh.is_used = 1
                AND dh.hours_run > 0
                AND dh.work_date <= ?
            ), 0) - (mp.last_service_hours + mp.interval_hours)) AS overdue_hours
          FROM maintenance_plans mp
          JOIN assets a ON a.id = mp.asset_id
          WHERE UPPER(a.asset_code) = UPPER(?)
            AND mp.active = 1
          ORDER BY overdue_hours DESC
          LIMIT 1
        `).get(end, assetCode);
        const overdue = Number(pm?.overdue_hours || 0);
        const riskBand = overdue >= 200 ? "high" : overdue >= 50 ? "medium" : overdue > 0 ? "low" : "none";
        const short = overdue > 0
          ? `${assetCode}: PM overdue by ${overdue.toFixed(1)}h (${riskBand} risk) as of ${end}${pm?.service_name ? ` on ${pm.service_name}` : ""}.`
          : `${assetCode}: no active PM overdue as of ${end}.`;
        return reply.send({
          ok: true,
          short_answer: short,
          details: {
            asset_code: assetCode,
            as_of: end,
            service_name: pm?.service_name || null,
            overdue_hours: overdue,
            risk_band: riskBand,
          },
        });
      }

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

  app.get("/risk-board", async (req, reply) => {
    try {
      const asOf = isDate(req.query?.date) ? String(req.query.date) : todayYmd();
      const limit = Math.max(1, Math.min(20, Number(req.query?.limit || 8)));
      const rows = db.prepare(`
        SELECT r.asset_code, r.risk_score, r.confidence, r.reasons_json, r.features_json, r.report_date
        FROM ironmind_asset_risk_snapshots r
        JOIN (
          SELECT asset_code, MAX(report_date) AS latest_date
          FROM ironmind_asset_risk_snapshots
          WHERE report_date <= ?
          GROUP BY asset_code
        ) x ON x.asset_code = r.asset_code AND x.latest_date = r.report_date
        ORDER BY r.risk_score DESC, r.confidence DESC, r.asset_code ASC
        LIMIT ?
      `).all(asOf, limit);

      const items = (rows || []).map((r) => {
        let reasons = [];
        try { reasons = JSON.parse(String(r.reasons_json || "[]")); } catch {}
        let features = {};
        try { features = JSON.parse(String(r.features_json || "{}")); } catch {}
        return {
          asset_code: String(r.asset_code || ""),
          report_date: String(r.report_date || ""),
          risk_score: Number(r.risk_score || 0),
          confidence: Number(r.confidence || 0),
          reasons: Array.isArray(reasons) ? reasons : [],
          features,
        };
      });

      return reply.send({ ok: true, as_of: asOf, items });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
}
