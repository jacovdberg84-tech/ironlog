// IRONLOG/api/routes/procurement/x3-site-requests.routes.js — requisitions downloaded as Sage X3 Site Requests.
// Registered by routes/procurement.routes.js; shared helpers arrive through ctx. See utils/x3SiteRequest.js.
import { db } from "../../db/client.js";
import { getSiteCode, getUser } from "../../utils/request.js";
import { writeAudit } from "../../utils/audit.js";
import { buildX3SiteRequests, ensureX3Tables, readX3Settings, saveX3Settings, X3_HEADER_FIELDS, X3_LINE_FIELDS, x3SiteRequestFile } from "../../utils/x3SiteRequest.js";

const X3_SETTINGS_ROLES = ["admin", "supervisor", "procurement"];
const X3_EXPORT_ROLES = ["admin", "supervisor", "procurement", "stores", "storeman"];

export default function registerX3SiteRequestRoutes(app, ctx) {
  const { requireRoles } = ctx;
  ensureX3Tables(db);

  const todayMaputo = () => new Intl.DateTimeFormat("en-CA", { timeZone: "Africa/Maputo" }).format(new Date());

  // Requisitions of the request's site with their lines, machine details and supplier code.
  const loadForX3 = (ids, siteCode) => ids.map((id) => {
    const requisition = db.prepare(`
      SELECT * FROM procurement_requisitions WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
    `).get(id, siteCode);
    if (!requisition) return null;
    const lines = db.prepare(`
      SELECT l.*, p.part_code, p.part_name
      FROM procurement_requisition_lines l LEFT JOIN parts p ON p.id = l.part_id
      WHERE l.requisition_id = ? ORDER BY l.line_no ASC, l.id ASC
    `).all(id);
    const equipment = String(lines[0]?.equipment_no || "").trim();
    let asset = null;
    if (equipment) {
      const cols = db.prepare(`PRAGMA table_info(assets)`).all().map((c) => c.name);
      const a = db.prepare(`SELECT id, ${cols.includes("serial_number") ? "serial_number" : "NULL AS serial_number"} FROM assets WHERE asset_code = ?`).get(equipment);
      if (a) {
        const m = db.prepare(`SELECT closing_hours FROM daily_hours WHERE asset_id = ? AND closing_hours IS NOT NULL ORDER BY work_date DESC LIMIT 1`).get(a.id);
        asset = { serial_number: a.serial_number, meter: m?.closing_hours ?? null };
      }
    }
    const sup = requisition.supplier_name
      ? db.prepare(`SELECT supplier_code FROM suppliers WHERE LOWER(TRIM(supplier_name)) = LOWER(TRIM(?)) LIMIT 1`).get(requisition.supplier_name)
      : null;
    return { requisition, lines, asset, supplier_code: sup?.supplier_code || "" };
  });

  const readIds = (body) => [...new Set((Array.isArray(body?.ids) ? body.ids : []).map(Number).filter((n) => Number.isInteger(n) && n > 0))].slice(0, 200);

  app.get("/x3/site-requests/settings", async () => {
    // site: company-wide X3 settings.
    return { ok: true, settings: readX3Settings(db), header_fields: X3_HEADER_FIELDS, line_fields: X3_LINE_FIELDS };
  });

  app.put("/x3/site-requests/settings", async (req, reply) => {
    // site: company-wide X3 settings (admin, supervisor, procurement).
    if (!requireRoles(req, reply, X3_SETTINGS_ROLES)) return;
    const r = saveX3Settings(db, req.body || {}, getUser(req));
    if (!r.ok) return reply.code(400).send({ error: r.errors.join(" "), errors: r.errors });
    try { writeAudit(db, req, { module: "procurement", action: "x3.settings", entity_type: "x3_site_request_settings", entity_id: 1, payload: {} }); } catch {}
    return { ok: true, settings: r.settings };
  });

  app.post("/x3/site-requests/check", async (req, reply) => {
    const ids = readIds(req.body);
    const found = loadForX3(ids, getSiteCode(req));
    const notFound = ids.filter((_, i) => !found[i]);
    if (notFound.length) return reply.code(404).send({ error: `Requisition not found on this site: ${notFound.join(", ")}` });
    const built = buildX3SiteRequests(found, readX3Settings(db), { today: todayMaputo() });
    return { ok: built.ok, errors: built.errors, warnings: built.warnings, records: built.records };
  });

  app.post("/x3/site-requests/export", async (req, reply) => {
    if (!requireRoles(req, reply, X3_EXPORT_ROLES)) return;
    const ids = readIds(req.body);
    const siteCode = getSiteCode(req);
    const found = loadForX3(ids, siteCode);
    const notFound = ids.filter((_, i) => !found[i]);
    if (notFound.length) return reply.code(404).send({ error: `Requisition not found on this site: ${notFound.join(", ")}` });
    const settings = readX3Settings(db);
    const built = buildX3SiteRequests(found, settings, { today: todayMaputo() });
    if (!built.ok) return reply.code(409).send({ error: built.errors.join(" "), errors: built.errors });
    const user = getUser(req);
    const mark = db.prepare(`UPDATE procurement_requisitions SET x3_exported_at = datetime('now'), x3_exported_by = ? WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?`);
    db.transaction(() => { for (const id of ids) mark.run(user, id, siteCode); })();
    try { writeAudit(db, req, { module: "procurement", action: "x3.export", entity_type: "requisition", entity_id: ids.join(","), payload: { ids, refs: built.records.map((r) => r.ref) } }); } catch {}
    const stamp = todayMaputo().replace(/-/g, "");
    reply.header("Content-Type", "text/csv; charset=utf-8");
    reply.header("Content-Disposition", `attachment; filename="X3_SiteRequests_${stamp}_${ids.length === 1 ? `REQ${ids[0]}` : `${ids.length}reqs`}.csv"`);
    return reply.send(x3SiteRequestFile(built.records, settings));
  });

  app.put("/requisitions/:id/x3-number", async (req, reply) => {
    if (!requireRoles(req, reply, X3_EXPORT_ROLES)) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isInteger(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const no = String(req.body?.x3_site_request_no ?? "").trim();
    if (no && !/^[A-Za-z0-9][A-Za-z0-9._/-]{0,29}$/.test(no)) return reply.code(400).send({ error: "Type the X3 Site Request number as X3 shows it, e.g. VS120-00353." });
    const r = db.prepare(`UPDATE procurement_requisitions SET x3_site_request_no = ?, updated_at = datetime('now') WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?`)
      .run(no || null, id, getSiteCode(req));
    if (!r.changes) return reply.code(404).send({ error: "requisition not found" });
    try { writeAudit(db, req, { module: "procurement", action: "x3.number", entity_type: "requisition", entity_id: id, payload: { x3_site_request_no: no || null } }); } catch {}
    return { ok: true, id, x3_site_request_no: no || null };
  });
}
