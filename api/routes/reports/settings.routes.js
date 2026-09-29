// IRONLOG/api/routes/reports/settings.routes.js — Custom report builder, SMTP and PDF settings, report subscriptions.
// Registered by routes/reports.routes.js; shared helpers arrive through ctx.
import ExcelJS from "exceljs";
import fs from "node:fs";
import path from "node:path";
import sharp from "sharp";
import { buildSmtpTransport, formatSmtpError, getSmtpSettingsRow, saveSmtpSettings, smtpPublicPayload } from "../../utils/mail.js";
import { clearPdfCompanyLogo, getPdfReportBranding, resolvePdfCompanyLogoAbs, savePdfReportBranding, setPdfCompanyLogoPath } from "../../utils/reportSettings.js";
import { db } from "../../db/client.js";

export default function registerSettingsRoutes(app, ctx) {
  const {
    allowedChannels,
    allowedFrequencies,
    allowedReportTypes,
    dataRoot,
    datasetWithAvailableColumns,
    deliverSubscription,
    nextRunForSchedule,
    parseRecipients,
    parseTimeHhMm,
    pdfBrandingLogoDir,
    reportDatasets,
    requireAdmin,
    runCustomBuilderQuery,
  } = ctx;

  app.get("/custom-builder/meta", async (_req, reply) => {
    const datasets = Object.keys(reportDatasets)
      .map((k) => datasetWithAvailableColumns(k))
      .filter(Boolean)
      .map((d) => ({
        key: d.key,
        table: d.table,
        columns: d.availableColumns.map((c) => c.id),
      }));
    return reply.send({ ok: true, datasets });
  });

  app.get("/custom-builder/templates", async (_req, reply) => {
    const rows = db.prepare(`
      SELECT id, name, dataset, columns_json, filters_json, created_by, created_at, updated_at
      FROM report_templates
      ORDER BY updated_at DESC, id DESC
      LIMIT 200
    `).all();
    const templates = rows.map((r) => {
      let columns = [];
      let filters = {};
      try { columns = JSON.parse(String(r.columns_json || "[]")); } catch {}
      try { filters = JSON.parse(String(r.filters_json || "{}")); } catch {}
      return {
        id: Number(r.id || 0),
        name: String(r.name || ""),
        dataset: String(r.dataset || ""),
        columns: Array.isArray(columns) ? columns : [],
        filters: filters && typeof filters === "object" ? filters : {},
        created_by: String(r.created_by || ""),
        created_at: r.created_at,
        updated_at: r.updated_at,
      };
    });
    return reply.send({ ok: true, templates });
  });

  app.post("/custom-builder/templates", async (req, reply) => {
    const body = req.body || {};
    const name = String(body.name || "").trim();
    const dataset = String(body.dataset || "").trim();
    const ds = datasetWithAvailableColumns(dataset);
    if (!name) return reply.code(400).send({ ok: false, error: "Template name is required" });
    if (!ds) return reply.code(400).send({ ok: false, error: "Invalid dataset" });
    const validCols = new Set(ds.availableColumns.map((c) => c.id));
    const columns = Array.from(new Set((Array.isArray(body.columns) ? body.columns : []).map((c) => String(c || "").trim())))
      .filter((c) => validCols.has(c))
      .slice(0, 25);
    if (!columns.length) return reply.code(400).send({ ok: false, error: "Select at least one valid column" });
    const filters = body.filters && typeof body.filters === "object" ? body.filters : {};
    const who = String(req.headers["x-user-name"] || "system");
    const id = Number(body.id || 0);
    if (id > 0) {
      const existing = db.prepare("SELECT id FROM report_templates WHERE id = ? LIMIT 1").get(id);
      if (!existing) return reply.code(404).send({ ok: false, error: "Template not found" });
      db.prepare(`
        UPDATE report_templates
        SET name = ?, dataset = ?, columns_json = ?, filters_json = ?, created_by = ?, updated_at = datetime('now')
        WHERE id = ?
      `).run(name, ds.key, JSON.stringify(columns), JSON.stringify(filters), who, id);
      return reply.send({ ok: true, id });
    }
    const out = db.prepare(`
      INSERT INTO report_templates (name, dataset, columns_json, filters_json, created_by, updated_at)
      VALUES (?, ?, ?, ?, ?, datetime('now'))
    `).run(name, ds.key, JSON.stringify(columns), JSON.stringify(filters), who);
    return reply.send({ ok: true, id: Number(out.lastInsertRowid || 0) });
  });

  app.delete("/custom-builder/templates/:id", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ ok: false, error: "Invalid template id" });
    const out = db.prepare("DELETE FROM report_templates WHERE id = ?").run(id);
    if (!Number(out.changes || 0)) return reply.code(404).send({ ok: false, error: "Template not found" });
    return reply.send({ ok: true });
  });

  app.post("/custom-builder/preview", async (req, reply) => {
    try {
      const result = runCustomBuilderQuery(req.body || {});
      return reply.send({ ok: true, ...result });
    } catch (err) {
      return reply.code(400).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/custom-builder/export.xlsx", async (req, reply) => {
    try {
      const result = runCustomBuilderQuery(req.body || {});
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Custom Report");
      const headers = result.columns.map((c) => String(c || ""));
      ws.addRow(headers);
      for (const r of result.rows) {
        ws.addRow(headers.map((h) => r?.[h] ?? ""));
      }
      ws.views = [{ state: "frozen", ySplit: 1 }];
      ws.columns.forEach((col) => {
        let width = 12;
        col.eachCell({ includeEmpty: true }, (cell) => {
          width = Math.max(width, Math.min(48, String(cell.value ?? "").length + 2));
        });
        col.width = width;
      });
      const safeDataset = String(result.dataset || "report").replace(/[^a-z0-9_-]/gi, "_");
      const stamp = new Date().toISOString().slice(0, 10);
      const buf = await wb.xlsx.writeBuffer();
      return reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Content-Disposition", `attachment; filename="IRONLOG_Custom_${safeDataset}_${stamp}.xlsx"`)
        .send(Buffer.from(buf));
    } catch (err) {
      return reply.code(400).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/smtp-settings", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    return reply.send({ ok: true, settings: smtpPublicPayload(getSmtpSettingsRow()) });
  });

  app.post("/smtp-settings", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const body = req.body || {};
    const who = String(req.headers["x-user-name"] || "system");
    const out = saveSmtpSettings({
      host: body.host,
      port: body.port,
      secure: body.secure,
      username: body.username,
      password: String(body.password || ""),
      from_email: body.from_email,
      from_name: body.from_name,
      updated_by: who,
    });
    if (!out.ok) return reply.code(400).send({ ok: false, error: out.error });
    return reply.send({ ok: true, settings: out.settings });
  });

  app.post("/smtp-settings/test", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const toRaw = String(req.body?.to || "").trim();
    if (!toRaw) return reply.code(400).send({ ok: false, error: "Recipient email is required" });
    const recipients = parseRecipients(toRaw);
    if (!recipients.length) return reply.code(400).send({ ok: false, error: "Invalid recipient list" });
    try {
      const smtp = buildSmtpTransport();
      if (smtp.error) return reply.code(400).send({ ok: false, error: smtp.error });
      await smtp.transporter.verify();
      await smtp.transporter.sendMail({
        from: smtp.from,
        to: recipients.join(", "),
        subject: "IRONLOG SMTP test email",
        text: `SMTP test successful at ${new Date().toISOString()}`,
      });
      return reply.send({ ok: true, message: "Test email sent" });
    } catch (err) {
      return reply.code(500).send({ ok: false, error: formatSmtpError(err) });
    }
  });

  app.get("/pdf-settings", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const branding = getPdfReportBranding(db);
    let sites = [];
    let companies = [];
    try {
      sites = db.prepare(`
        SELECT site_code, site_name, company_code
        FROM site_profiles
        WHERE COALESCE(active, 1) = 1
        ORDER BY site_name ASC
        LIMIT 500
      `).all();
    } catch {
      sites = [];
    }
    try {
      companies = db.prepare(`
        SELECT company_code, company_name
        FROM company_profiles
        WHERE COALESCE(active, 1) = 1
        ORDER BY company_name ASC
        LIMIT 200
      `).all();
    } catch {
      companies = [];
    }
    return reply.send({ ok: true, ...branding, sites, companies });
  });

  app.get("/pdf-settings/logo", async (req, reply) => {
    const abs = resolvePdfCompanyLogoAbs(db, dataRoot);
    if (!abs || !fs.existsSync(abs)) {
      return reply.code(404).send({ ok: false, error: "No PDF logo configured" });
    }
    const ext = path.extname(abs).toLowerCase();
    const contentType = ext === ".png" ? "image/png" : ext === ".webp" ? "image/webp" : "image/jpeg";
    reply.header("Content-Type", contentType);
    reply.header("Cache-Control", "no-store");
    return reply.send(fs.readFileSync(abs));
  });

  app.post("/pdf-settings/logo", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    try {
      const part = await req.file();
      if (!part) return reply.code(400).send({ ok: false, error: "Upload an image in form field 'file'" });

      const mime = String(part.mimetype || "").toLowerCase();
      if (!mime.startsWith("image/")) {
        return reply.code(400).send({ ok: false, error: "File must be an image (PNG, JPEG, or WebP)" });
      }

      const rawBuffer = await part.toBuffer();
      if (!rawBuffer?.length) return reply.code(400).send({ ok: false, error: "Empty file" });

      const meta = await sharp(rawBuffer, { failOn: "none" }).metadata();
      const hasAlpha = Boolean(meta.hasAlpha);
      const pipeline = sharp(rawBuffer, { failOn: "none" })
        .rotate()
        .resize({ width: 480, height: 240, fit: "inside", withoutEnlargement: true });

      let outBuffer;
      let ext;
      if (hasAlpha) {
        outBuffer = await pipeline.png({ compressionLevel: 9 }).toBuffer();
        ext = ".png";
      } else {
        outBuffer = await pipeline.jpeg({ quality: 90, mozjpeg: true }).toBuffer();
        ext = ".jpg";
      }

      clearPdfCompanyLogo(db, dataRoot);
      const fileName = `company-logo${ext}`;
      const absPath = path.join(pdfBrandingLogoDir, fileName);
      await fs.promises.writeFile(absPath, outBuffer);

      const relPath = path.join("uploads", "report-branding", fileName).replace(/\\/g, "/");
      setPdfCompanyLogoPath(db, relPath);
      const branding = getPdfReportBranding(db);
      return reply.send({ ok: true, message: "PDF logo uploaded.", ...branding });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.delete("/pdf-settings/logo", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    clearPdfCompanyLogo(db, dataRoot);
    const branding = getPdfReportBranding(db);
    return reply.send({ ok: true, message: "PDF logo removed.", ...branding });
  });

  app.post("/pdf-settings", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const company_code = String(req.body?.company_code || "").trim();
    const company_name = String(req.body?.company_name || "").trim();
    const site_code = String(req.body?.site_code || "").trim();
    const site_name = String(req.body?.site_name || "").trim();
    if (!company_name && !company_code && !site_name && !site_code) {
      return reply.code(400).send({ ok: false, error: "Enter a company name and/or site name" });
    }
    const saved = savePdfReportBranding(db, { company_code, company_name, site_code, site_name });
    return reply.send({ ok: true, ...saved });
  });

  app.get("/subscriptions", async (_req, reply) => {
    const rows = db.prepare(`
      SELECT
        id, name, report_type, channel, recipients, schedule_frequency, send_time,
        day_of_week, day_of_month, active, filters_json, last_sent_at, next_run_at,
        created_by, created_at, updated_at
      FROM report_subscriptions
      ORDER BY active DESC, next_run_at ASC, id DESC
      LIMIT 300
    `).all();
    const subscriptions = rows.map((r) => {
      let filters = {};
      try { filters = JSON.parse(String(r.filters_json || "{}")); } catch {}
      return {
        id: Number(r.id || 0),
        name: String(r.name || ""),
        report_type: String(r.report_type || ""),
        channel: String(r.channel || ""),
        recipients: parseRecipients(r.recipients || ""),
        schedule_frequency: String(r.schedule_frequency || "weekly"),
        send_time: parseTimeHhMm(r.send_time),
        day_of_week: r.day_of_week == null ? null : Number(r.day_of_week),
        day_of_month: r.day_of_month == null ? null : Number(r.day_of_month),
        active: Number(r.active || 0) === 1 ? 1 : 0,
        filters: filters && typeof filters === "object" ? filters : {},
        last_sent_at: r.last_sent_at,
        next_run_at: r.next_run_at,
        created_by: String(r.created_by || ""),
        created_at: r.created_at,
        updated_at: r.updated_at,
      };
    });
    return reply.send({ ok: true, subscriptions });
  });

  app.post("/subscriptions", async (req, reply) => {
    try {
      const body = req.body || {};
      const id = Number(body.id || 0);
      const name = String(body.name || "").trim();
      const reportType = String(body.report_type || "").trim().toLowerCase();
      const channel = String(body.channel || "").trim().toLowerCase();
      const freq = String(body.schedule_frequency || "weekly").trim().toLowerCase();
      const sendTime = parseTimeHhMm(body.send_time);
      const dayOfWeek = body.day_of_week == null ? null : Math.max(0, Math.min(6, Number(body.day_of_week)));
      const dayOfMonth = body.day_of_month == null ? null : Math.max(1, Math.min(28, Number(body.day_of_month)));
      const active = Number(body.active ?? 1) === 1 ? 1 : 0;
      const recipients = parseRecipients(body.recipients || "");
      const filters = body.filters && typeof body.filters === "object" ? body.filters : {};
      if (!name) return reply.code(400).send({ ok: false, error: "Name is required" });
      if (!allowedReportTypes.has(reportType)) return reply.code(400).send({ ok: false, error: "Invalid report_type" });
      if (!allowedChannels.has(channel)) return reply.code(400).send({ ok: false, error: "Invalid channel" });
      if (!allowedFrequencies.has(freq)) return reply.code(400).send({ ok: false, error: "Invalid schedule_frequency" });
      if (!recipients.length) return reply.code(400).send({ ok: false, error: "At least one recipient is required" });
      const who = String(req.headers["x-user-name"] || "system");
      const nextRun = active ? nextRunForSchedule({
        schedule_frequency: freq,
        send_time: sendTime,
        day_of_week: dayOfWeek,
        day_of_month: dayOfMonth,
      }, new Date()) : null;
      if (id > 0) {
        const ex = db.prepare(`SELECT id FROM report_subscriptions WHERE id = ? LIMIT 1`).get(id);
        if (!ex) return reply.code(404).send({ ok: false, error: "Subscription not found" });
        db.prepare(`
          UPDATE report_subscriptions
          SET
            name = ?, report_type = ?, channel = ?, recipients = ?, schedule_frequency = ?, send_time = ?,
            day_of_week = ?, day_of_month = ?, active = ?, filters_json = ?, next_run_at = ?, updated_at = datetime('now')
          WHERE id = ?
        `).run(
          name, reportType, channel, recipients.join(","), freq, sendTime, dayOfWeek, dayOfMonth, active,
          JSON.stringify(filters), nextRun, id
        );
        return reply.send({ ok: true, id });
      }
      const out = db.prepare(`
        INSERT INTO report_subscriptions (
          name, report_type, channel, recipients, schedule_frequency, send_time, day_of_week, day_of_month,
          active, filters_json, next_run_at, created_by, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(
        name, reportType, channel, recipients.join(","), freq, sendTime, dayOfWeek, dayOfMonth,
        active, JSON.stringify(filters), nextRun, who
      );
      return reply.send({ ok: true, id: Number(out.lastInsertRowid || 0) });
    } catch (err) {
      req.log?.error?.(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.delete("/subscriptions/:id", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
    const out = db.prepare(`DELETE FROM report_subscriptions WHERE id = ?`).run(id);
    if (!Number(out.changes || 0)) return reply.code(404).send({ ok: false, error: "Subscription not found" });
    return reply.send({ ok: true });
  });

  app.post("/subscriptions/:id/send-now", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
    const row = db.prepare(`SELECT * FROM report_subscriptions WHERE id = ? LIMIT 1`).get(id);
    if (!row) return reply.code(404).send({ ok: false, error: "Subscription not found" });
    try {
      const out = await deliverSubscription(row, true);
      return reply.send({ ok: true, ...out });
    } catch (err) {
      db.prepare(`
        INSERT INTO report_delivery_logs (subscription_id, report_type, channel, recipients, status, detail, created_at)
        VALUES (?, ?, ?, ?, 'failed', ?, datetime('now'))
      `).run(id, row.report_type, row.channel, row.recipients, String(err.message || err));
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/subscriptions/logs", async (req, reply) => {
    const limit = Math.max(1, Math.min(200, Number(req.query?.limit || 50)));
    const rows = db.prepare(`
      SELECT id, subscription_id, report_type, channel, recipients, status, detail, created_at
      FROM report_delivery_logs
      ORDER BY created_at DESC, id DESC
      LIMIT ?
    `).all(limit);
    return reply.send({ ok: true, rows });
  });
}
