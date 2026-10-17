// IRONLOG/api/routes/upload/go-live.routes.js — the guided go-live import (one Excel workbook). See utils/goLiveImport.js.
// Registered by routes/upload.routes.js, after the file upload support.
import { db } from "../../db/client.js";
import { writeAudit } from "../../utils/audit.js";
import { listFailureCodes } from "../../utils/failureCodes.js";
import {
  applyGoLiveImport, ensureGoLiveTables, goLiveCurrentRows, goLiveStatus, goLiveTemplate, planGoLiveImport, readGoLiveWorkbook,
} from "../../utils/goLiveImport.js";
import { getUser, holdsAnyRole } from "../../utils/request.js";
import { knownSites, requestSite } from "../../utils/sites.js";

const GO_LIVE_ROLES = ["admin", "supervisor", "workshop_admin"];

export default function registerGoLiveRoutes(app) {
  ensureGoLiveTables(db);

  const allowed = (req, reply) => {
    if (holdsAnyRole(req, GO_LIVE_ROLES)) return true;
    reply.code(403).send({ error: "The go-live import is for administrators and plant administrators." });
    return false;
  };
  const todayMaputo = () => new Intl.DateTimeFormat("en-CA", { timeZone: "Africa/Maputo" }).format(new Date());
  const loadExcel = async () => (await import("exceljs")).default;

  // Reads the uploaded workbook and checks it against IronLog for the request's site.
  const checkUpload = async (req, reply) => {
    const file = await req.file();
    if (!file) { reply.code(400).send({ error: "Choose the filled-in go-live workbook (.xlsx)." }); return null; }
    const parsed = await readGoLiveWorkbook(await loadExcel(), await file.toBuffer());
    if (!parsed.ok) { reply.code(400).send({ error: parsed.error }); return null; }
    const plan = planGoLiveImport(db, parsed, { site: requestSite(req), allowedSites: req.ironlogAllowedSites || null, today: todayMaputo() });
    return { plan, fileName: String(file.filename || "").slice(0, 120) };
  };
  const reportOf = ({ ok, errors, warnings, summary, rows }) => ({ ok, errors, warnings, summary, rows });

  app.get("/go-live/template.xlsx", async (req, reply) => {
    if (!allowed(req, reply)) return;
    const site = requestSite(req);
    const sites = knownSites(db);
    const mine = req.ironlogAllowedSites ? sites.filter((s) => req.ironlogAllowedSites.includes(s)) : sites;
    const filled = String(req.query?.filled || "") === "1";
    const wb = goLiveTemplate(await loadExcel(), {
      rows: filled ? goLiveCurrentRows(db, [site].filter((s) => mine.includes(s))) : {},
      site,
      sites: mine,
      componentCodes: listFailureCodes(db, { all: true }).component.map((c) => c.code),
    });
    reply.header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    reply.header("Content-Disposition", `attachment; filename="IronLog_GoLive_${site}${filled ? "_current" : ""}.xlsx"`);
    return reply.send(Buffer.from(await wb.xlsx.writeBuffer()));
  });

  app.post("/go-live/check", async (req, reply) => {
    // site: machines go to the request's site (or the Site column, checked against the person's sites).
    if (!allowed(req, reply)) return;
    const r = await checkUpload(req, reply);
    if (!r) return;
    return reportOf(r.plan);
  });

  app.post("/go-live/import", async (req, reply) => {
    // site: as /go-live/check; the import is refused unless the check is clean.
    if (!allowed(req, reply)) return;
    const r = await checkUpload(req, reply);
    if (!r) return;
    if (!r.plan.ok) return reply.code(409).send({ error: "The file still has problems. Check it, fix them and try again.", ...reportOf(r.plan) });
    const by = getUser(req);
    applyGoLiveImport(db, r.plan, { by });
    const id = db.prepare(`INSERT INTO go_live_imports (file_name, site_code, summary_json, imported_by) VALUES (?, ?, ?, ?)`)
      .run(r.fileName, requestSite(req), JSON.stringify(r.plan.summary), by || null).lastInsertRowid;
    try { writeAudit(db, req, { module: "go_live", action: "go_live.import", entity_type: "go_live_imports", entity_id: id, payload: { file: r.fileName, summary: r.plan.summary } }); } catch {}
    return { ...reportOf(r.plan), imported: true, id: Number(id) };
  });

  app.get("/go-live/status", async (req, reply) => {
    // site: the checklist is for the request's site.
    if (!allowed(req, reply)) return;
    const history = db.prepare(`SELECT id, file_name, site_code, summary_json, imported_at, imported_by FROM go_live_imports ORDER BY id DESC LIMIT 10`).all()
      .map((h) => ({ ...h, summary: JSON.parse(h.summary_json || "{}"), summary_json: undefined }));
    return { ok: true, ...goLiveStatus(db, requestSite(req)), history };
  });
}
