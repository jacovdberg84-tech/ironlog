// IRONLOG/api/routes/stock/lube-model.routes.js — the monthly lube costing model.
// Registered by routes/stock.routes.js; shared helpers arrive through ctx.
//
// The site's LUBE TEMPLATE workbook is uploaded once (again when a new version
// comes). For a month, IronLog fills its input cells from lube issues and
// receipts and hands the workbook back; see utils/lubeModel.js.

import fs from "node:fs";
import path from "node:path";
import multipart from "@fastify/multipart";
import { db } from "../../db/client.js";
import { writeAudit } from "../../utils/audit.js";
import {
  MODEL_OIL_TYPES,
  checkTemplate,
  collectMonth,
  ensureLubeModelSchema,
  fillTemplate,
  lubeParts,
  savePartMapping,
  templateDir,
  templateInfo,
} from "../../utils/lubeModel.js";

const ROLES = ["admin", "supervisor", "stores", "storeman", "workshop_admin", "plant_manager", "site_manager"];

function monthParam(req) {
  const m = /^(\d{4})-(\d{2})$/.exec(String(req.query?.month || "").trim());
  if (!m || Number(m[2]) < 1 || Number(m[2]) > 12) return null;
  return { year: Number(m[1]), month: Number(m[2]) };
}

export default function registerLubeModelRoutes(app, ctx) {
  const { requireRoles } = ctx;
  ensureLubeModelSchema(db);

  // GET /api/stock/lube-model — template on file, lube items with their model oil type and size.
  app.get("/lube-model", async (req, reply) => {
    if (!requireRoles(req, reply, ROLES)) return;
    const t = templateInfo(db);
    return {
      ok: true,
      template: t ? { file_name: t.file_name, size_bytes: t.size_bytes, uploaded_by: t.uploaded_by, uploaded_at: t.uploaded_at } : null,
      oil_types: MODEL_OIL_TYPES,
      parts: lubeParts(db),
    };
  });

  // PUT /api/stock/lube-model/parts/:code { oil_type, pack_size }
  app.put("/lube-model/parts/:code", async (req, reply) => {
    if (!requireRoles(req, reply, ROLES)) return;
    try {
      savePartMapping(db, req.params.code, req.body || {}, String(req.headers["x-user-name"] || "").trim() || null);
    } catch (err) {
      return reply.code(err.status || 400).send({ ok: false, error: err.message });
    }
    return { ok: true, part: lubeParts(db).find((p) => String(p.part_code).toUpperCase() === String(req.params.code).toUpperCase()) || null };
  });

  // GET /api/stock/lube-model/preview?month=YYYY-MM — what will go into the model.
  app.get("/lube-model/preview", async (req, reply) => {
    if (!requireRoles(req, reply, ROLES)) return;
    const m = monthParam(req);
    if (!m) return reply.code(400).send({ ok: false, error: "Choose a month (YYYY-MM)" });
    const data = collectMonth(db, m.year, m.month);
    const byType = new Map();
    for (const d of data.days) for (const r of d.rows) byType.set(r.oil_type, Number(((byType.get(r.oil_type) || 0) + r.qty).toFixed(2)));
    return {
      ok: true,
      month: `${m.year}-${String(m.month).padStart(2, "0")}`,
      template: Boolean(templateInfo(db)),
      issue_days: data.days.length,
      issue_lines: data.issue_lines,
      issue_litres: data.issue_litres,
      issues_by_type: [...byType.entries()].map(([oil_type, qty]) => ({ oil_type, qty })).sort((a, b) => b.qty - a.qty),
      deliveries: data.deliveries,
      opening: data.opening,
      unmapped: data.unmapped,
      too_many_days: data.too_many_days,
      days: data.days.map((d) => ({ day: d.day, lines: d.rows.length })),
    };
  });

  // GET /api/stock/lube-model/download?month=YYYY-MM — the filled workbook.
  app.get("/lube-model/download", async (req, reply) => {
    if (!requireRoles(req, reply, ROLES)) return;
    const m = monthParam(req);
    if (!m) return reply.code(400).send({ ok: false, error: "Choose a month (YYYY-MM)" });
    const t = templateInfo(db);
    if (!t) return reply.code(409).send({ ok: false, error: "Upload the lube model template first" });
    const data = collectMonth(db, m.year, m.month);
    if (data.too_many_days) return reply.code(409).send({ ok: false, error: "More issue days than the model's 50 input sheets" });
    const { buffer, deliveries_added } = await fillTemplate(fs.readFileSync(t.stored_path), data);
    writeAudit(db, req, {
      module: "stock",
      action: "lube_model_download",
      entity_type: "lube_model",
      entity_id: `${m.year}-${m.month}`,
      payload: { issue_days: data.days.length, issue_lines: data.issue_lines, deliveries_added, unmapped: data.unmapped.length },
    });
    const ext = path.extname(t.file_name || "").toLowerCase() === ".xlsx" ? "xlsx" : "xlsm";
    const monthName = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"][m.month - 1];
    return reply
      .header("Content-Type", ext === "xlsm" ? "application/vnd.ms-excel.sheet.macroEnabled.12" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="LUBE_MODEL_${monthName}${String(m.year).slice(2)}_IRONLOG.${ext}"`)
      .send(buffer);
  });

  // POST /api/stock/lube-model/template (multipart "file") — keep the site's current template.
  app.register(async (sub) => {
    await sub.register(multipart, { limits: { fileSize: 40 * 1024 * 1024, files: 1 } });
    sub.post("/lube-model/template", async (req, reply) => {
      if (!requireRoles(req, reply, ROLES)) return;
      const part = await req.file();
      if (!part) return reply.code(400).send({ ok: false, error: "Attach the lube model workbook" });
      const name = String(part.filename || "lube-model.xlsm");
      if (!/\.(xlsm|xlsx)$/i.test(name)) return reply.code(400).send({ ok: false, error: "The lube model must be an .xlsm or .xlsx workbook" });
      const buf = await part.toBuffer();
      try {
        await checkTemplate(buf);
      } catch (err) {
        return reply.code(err.status || 400).send({ ok: false, error: err.status ? err.message : "Could not read that workbook" });
      }
      const stored = path.join(templateDir(), `lube-model-template${path.extname(name).toLowerCase()}`);
      fs.writeFileSync(stored, buf);
      const user = String(req.headers["x-user-name"] || "").trim() || null;
      db.prepare(`
        INSERT INTO lube_model_template (id, file_name, stored_path, size_bytes, uploaded_by, uploaded_at)
        VALUES (1, ?, ?, ?, ?, datetime('now'))
        ON CONFLICT(id) DO UPDATE SET file_name = excluded.file_name, stored_path = excluded.stored_path,
          size_bytes = excluded.size_bytes, uploaded_by = excluded.uploaded_by, uploaded_at = excluded.uploaded_at
      `).run(name, stored, buf.length, user);
      writeAudit(db, req, { module: "stock", action: "lube_model_template", entity_type: "lube_model", entity_id: "template", payload: { file_name: name, size_bytes: buf.length } });
      return { ok: true, template: { file_name: name, size_bytes: buf.length, uploaded_by: user } };
    });
  });
}
