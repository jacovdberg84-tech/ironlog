import { generateIronmindReport, getIronmindHistory, getLatestIronmindReport } from "../utils/ironmind.js";

function toBool(v) {
  const s = String(v || "").trim().toLowerCase();
  return s === "1" || s === "true" || s === "yes" || s === "y";
}

export default async function ironmindRoutes(app) {
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
}
