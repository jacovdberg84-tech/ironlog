// IRONLOG/api/routes/maintenance/inspections.routes.js — Manager, tyre, undercarriage and artisan inspections.
// Registered by routes/maintenance.routes.js; shared helpers arrive through ctx.
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { UNDERCARRIAGE_CHECKLIST_ITEMS, UNDERCARRIAGE_TRACK_SAG_POINTS, UNDERCARRIAGE_WEAR_BANDS, buildUndercarriageComponentSchema, normalizeUndercarriageChecklist, normalizeUndercarriageTrackSag, summarizeUndercarriageInspection } from "../../utils/undercarriageTemplate.js";
import { buildPdfBuffer, pdfBodyTop } from "../../utils/pdfGenerator.js";
import { db } from "../../db/client.js";
import { getPdfReportBranding } from "../../utils/reportSettings.js";
import { isDate } from "../../utils/request.js";

export default function registerInspectionsRoutes(app, ctx) {
  const {
    buildInspectionWorkOrderNotes,
    buildTyreLifecycleSummary,
    buildTyreMonthlySurveyData,
    buildTyreSurveyXlsxBuffer,
    buildUndercarriageMeasurementsForSave,
    buildUndercarriageXlsxBuffer,
    drawTyreSurveyPdfMachine,
    drawUndercarriageInspectionPdf,
    enrichTyreLifecycleRow,
    ensureColumn,
    ensureTyreLifecycleSchema,
    getAssetHoursInfoAsOf,
    getTyreThresholds,
    getUndercarriageWearProfileRow,
    hasColumn,
    importUndercarriageWearProfileFromLatest,
    inspectionsDir,
    isMonth,
    monthBoundsYmd,
    normalizeArtisanInspectionChecklist,
    normalizeInspectionChecklist,
    normalizeInspectionParts,
    normalizeTyreRows,
    parseUndercarriageInspectionRow,
    photoCaptionCol,
    photoCreatedCol,
    photoInspectionCol,
    photoPathCol,
    resolveStorageAbs,
    saveUndercarriageWearProfile,
  } = ctx;

  // =====================================================
  // MANAGER INSPECTIONS
  // =====================================================
  app.get("/inspections", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const params = [];
      const where = ["LOWER(TRIM(COALESCE(mi.site_code, 'main'))) = ?"];
      params.push(site_code);
      if (assetId > 0) {
        where.push("mi.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("mi.inspection_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("mi.inspection_date <= ?");
        params.push(end);
      }

      const rows = db.prepare(`
        SELECT
          mi.id,
          mi.asset_id,
          mi.inspection_date,
          mi.inspector_name,
          mi.notes,
          mi.machine_hours,
          mi.live_hours_snapshot,
          mi.live_hours_source,
          mi.checklist_json,
          mi.required_parts_json,
          mi.defect_severity,
          mi.defect_component,
          mi.defect_risk,
          mi.recommended_action,
          mi.inspection_type,
          mi.evidence_required,
          mi.evidence_photo_count,
          mi.work_order_id,
          mi.created_at,
          a.asset_code,
          a.asset_name
        FROM manager_inspections mi
        JOIN assets a ON a.id = mi.asset_id
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY mi.inspection_date DESC, mi.id DESC
      `).all(...params);

      const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
      let photosByInspection = new Map();
      if (ids.length) {
        const marks = ids.map(() => "?").join(",");
        const photos = db.prepare(`
          SELECT
            id,
            ${photoInspectionCol} AS inspection_id,
            ${photoPathCol} AS file_path,
            ${photoCaptionCol} AS caption,
            ${photoCreatedCol} AS created_at
          FROM manager_inspection_photos
          WHERE ${photoInspectionCol} IN (${marks})
          ORDER BY id ASC
        `).all(...ids);
        photosByInspection = photos.reduce((m, p) => {
          const k = Number(p.inspection_id);
          if (!m.has(k)) m.set(k, []);
          m.get(k).push(p);
          return m;
        }, new Map());
      }

      return reply.send({
        ok: true,
        rows: rows.map((r) => {
          const toChecklistLabel = (key) =>
            String(key || "")
              .replaceAll("_", " ")
              .replace(/\b\w/g, (m) => m.toUpperCase())
              .trim();
          let checklist = [];
          let required_parts = [];
          try {
            let parsed = JSON.parse(String(r.checklist_json || "null"));
            // mobile ingest bundle shape: { checklist, checklist_details, ... }.
            if (parsed && !Array.isArray(parsed) && typeof parsed === "object" && parsed.checklist && typeof parsed.checklist === "object") {
              parsed = parsed.checklist;
            }
            if (Array.isArray(parsed)) {
              checklist = parsed;
            } else if (parsed && typeof parsed === "object") {
              let details = null;
              if (parsed.checklist_details && typeof parsed.checklist_details === "object") {
                details = parsed.checklist_details;
              } else {
                try {
                  const d = JSON.parse(String(r.checklist_detail_json || "null"));
                  if (d && typeof d === "object") details = d;
                } catch {}
              }
              checklist = Object.entries(parsed).map(([key, status]) => {
                const st = String(status || "").trim().toLowerCase();
                const ok = st === "ok" ? true : (st === "attention" || st === "unsafe" || st === "fail" || st === "failed") ? false : null;
                const note = String(details?.[key]?.comment || details?.[key]?.note || details?.[key]?.notes || "").trim() || null;
                return { key, label: toChecklistLabel(key), ok, note };
              });
            }
          } catch {}
          try {
            const pj = JSON.parse(String(r.required_parts_json || "[]"));
            if (Array.isArray(pj)) required_parts = pj;
          } catch {}
          return {
            ...r,
            checklist,
            required_parts,
            photos: photosByInspection.get(Number(r.id)) || [],
          };
        }),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/inspections", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const inspection_date = String(req.body?.inspection_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(inspection_date)) return reply.code(400).send({ ok: false, error: "inspection_date must be YYYY-MM-DD" });

      const asset = db.prepare(`SELECT id, asset_code, asset_name FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const liveInfo = getAssetHoursInfoAsOf(asset_id, inspection_date);
      const liveSnap = Number(liveInfo.hours || 0);
      const liveSource = String(liveInfo.source || "");

      let machine_hours = null;
      const mhRaw = req.body?.machine_hours;
      if (mhRaw != null && mhRaw !== "") {
        const n = Number(mhRaw);
        if (Number.isFinite(n) && n >= 0) machine_hours = n;
      }
      if (machine_hours == null) machine_hours = liveSnap;

      const checklist = normalizeInspectionChecklist(req.body?.checklist);
      const checklist_json = JSON.stringify(checklist);
      const required_parts = normalizeInspectionParts(req.body?.required_parts);
      const required_parts_json = JSON.stringify(required_parts);
      const inspection_type = String(req.body?.inspection_type || "machine_general").trim().toLowerCase() || "machine_general";
      const defect_severity = String(req.body?.defect_severity || "").trim().toLowerCase() || null;
      const defect_component = String(req.body?.defect_component || "").trim() || null;
      const defect_risk = String(req.body?.defect_risk || "").trim() || null;
      const recommended_action = String(req.body?.recommended_action || "").trim() || null;
      const evidence_required = Number(req.body?.evidence_required ?? 1) ? 1 : 0;
      const evidence_photo_count = Math.max(0, Number(req.body?.evidence_photo_count || 0));

      const anyChecklistFail = checklist.some((c) => c.ok === false);
      const enforceEvidenceRules = Number(req.body?.enforce_evidence_rules || 0) === 1;
      if (enforceEvidenceRules && evidence_required && anyChecklistFail) {
        const missingFailComment = checklist.find((c) => c.ok === false && !String(c.note || "").trim());
        if (missingFailComment) {
          return reply.code(400).send({ ok: false, error: `Comment required for failed checklist item: ${missingFailComment.label}` });
        }
        if (evidence_photo_count < 1) {
          return reply.code(400).send({ ok: false, error: "At least one evidence photo is required for failed checklist items" });
        }
      }
      const hasParts = required_parts.length > 0;
      const createExplicit = Boolean(req.body?.create_work_order);
      const createOnIssues = req.body?.create_work_order_on_issues !== false;
      const shouldCreateWo =
        createExplicit || (createOnIssues && (anyChecklistFail || hasParts));

      const ins = db.prepare(`
        INSERT INTO manager_inspections (
          asset_id, uuid, site_code, inspection_date, inspector_name, notes,
          machine_hours, live_hours_snapshot, live_hours_source,
          checklist_json, required_parts_json,
          inspection_type,
          defect_severity, defect_component, defect_risk, recommended_action,
          evidence_required, evidence_photo_count,
          updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(
        asset_id,
        crypto.randomUUID(),
        site_code,
        inspection_date,
        inspector_name,
        notes,
        machine_hours,
        liveSnap,
        liveSource,
        checklist_json,
        required_parts_json,
        inspection_type,
        defect_severity,
        defect_component,
        defect_risk,
        recommended_action,
        evidence_required,
        evidence_photo_count
      );

      const inspectionId = Number(ins.lastInsertRowid);
      let work_order_id = null;

      if (shouldCreateWo) {
        ensureColumn("work_orders", "job_description TEXT", "job_description");
        ensureColumn("work_orders", "completion_notes TEXT", "completion_notes");
        const woNotes = buildInspectionWorkOrderNotes({
          inspectionId,
          inspection_date,
          asset_code: String(asset.asset_code || ""),
          asset_name: String(asset.asset_name || ""),
          checklist,
          required_parts,
          notes,
        });
        const wo = db.prepare(`
          INSERT INTO work_orders (asset_id, source, reference_id, status)
          VALUES (?, 'inspection', ?, 'open')
        `).run(asset_id, inspectionId);
        work_order_id = Number(wo.lastInsertRowid);
        if (work_order_id > 0 && woNotes) {
          if (hasColumn("work_orders", "job_description")) {
            db.prepare(`UPDATE work_orders SET job_description = ? WHERE id = ?`).run(woNotes, work_order_id);
          } else if (hasColumn("work_orders", "completion_notes")) {
            db.prepare(`UPDATE work_orders SET completion_notes = ? WHERE id = ?`).run(woNotes, work_order_id);
          }
        }
        if (work_order_id > 0) {
          db.prepare(`UPDATE manager_inspections SET work_order_id = ? WHERE id = ?`).run(work_order_id, inspectionId);
        }
      }

      return reply.send({ ok: true, id: inspectionId, work_order_id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/inspections/:id/photo", async (req, reply) => {
    try {
      const inspectionId = Number(req.params?.id || 0);
      if (!inspectionId) return reply.code(400).send({ ok: false, error: "Invalid inspection id" });

      const inspection = db.prepare(`SELECT id FROM manager_inspections WHERE id = ?`).get(inspectionId);
      if (!inspection) return reply.code(404).send({ ok: false, error: "Inspection not found" });

      const part = await req.file();
      if (!part) return reply.code(400).send({ ok: false, error: "Upload file field named 'file'" });

      const extRaw = path.extname(part.filename || "").toLowerCase();
      const ext = [".jpg", ".jpeg", ".png", ".webp"].includes(extRaw) ? extRaw : ".jpg";
      const safe = `mi_${inspectionId}_${Date.now()}_${Math.floor(Math.random() * 100000)}${ext}`;
      const absPath = path.join(inspectionsDir, safe);
      await fs.promises.writeFile(absPath, await part.toBuffer());

      const caption = String(req.query?.caption || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const relPath = path.join("uploads", "manager-inspections", safe).replace(/\\/g, "/");

      // Legacy compatibility: some DBs require manager_inspection_id NOT NULL,
      // others use inspection_id. If both exist, write both.
      const hasInspectionId = hasColumn("manager_inspection_photos", "inspection_id");
      const hasManagerInspectionId = hasColumn("manager_inspection_photos", "manager_inspection_id");
      const linkCols = [];
      const linkVals = [];
      if (hasInspectionId) {
        linkCols.push("inspection_id");
        linkVals.push(inspectionId);
      }
      if (hasManagerInspectionId) {
        linkCols.push("manager_inspection_id");
        linkVals.push(inspectionId);
      }
      if (!linkCols.length) {
        linkCols.push(photoInspectionCol);
        linkVals.push(inspectionId);
      }

      const hasImageData = hasColumn("manager_inspection_photos", "image_data");
      const insertCols = [...linkCols, "uuid", "site_code", "file_path", ...(hasImageData ? ["image_data"] : []), "caption", "updated_at"];
      const placeholders = [...insertCols.map((c) => (c === "updated_at" ? "datetime('now')" : "?"))].join(", ");
      const ins = db.prepare(`
        INSERT INTO manager_inspection_photos (${insertCols.join(", ")})
        VALUES (${placeholders})
      `).run(...linkVals, crypto.randomUUID(), site_code, relPath, ...(hasImageData ? [relPath] : []), caption);

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        inspection_id: inspectionId,
        file_path: `/${relPath}`,
        caption,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.delete("/inspections/:id", async (req, reply) => {
    try {
      const inspectionId = Number(req.params?.id || 0);
      if (!inspectionId) return reply.code(400).send({ ok: false, error: "Invalid inspection id" });
      const current = db.prepare(`
        SELECT id, work_order_id
        FROM manager_inspections
        WHERE id = ?
        LIMIT 1
      `).get(inspectionId);
      if (!current) return reply.code(404).send({ ok: false, error: "Inspection not found" });

      const photos = db.prepare(`
        SELECT ${photoPathCol} AS file_path
        FROM manager_inspection_photos
        WHERE ${photoInspectionCol} = ?
      `).all(inspectionId);
      for (const p of photos) {
        const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
        if (!rel) continue;
        const abs = resolveStorageAbs(rel);
        if (!abs || !fs.existsSync(abs)) continue;
        try { fs.unlinkSync(abs); } catch {}
      }

      db.prepare(`DELETE FROM manager_inspection_photos WHERE ${photoInspectionCol} = ?`).run(inspectionId);
      db.prepare(`DELETE FROM manager_inspections WHERE id = ?`).run(inspectionId);
      return reply.send({ ok: true, id: inspectionId, work_order_id: current.work_order_id || null });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/tyre-inspections/lifecycle", async (req, reply) => {
    try {
      ensureTyreLifecycleSchema();
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      if (!assetId) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      const summary = buildTyreLifecycleSummary(assetId, site_code);
      if (!summary) return reply.code(404).send({ ok: false, error: "Asset not found" });
      return reply.send({ ok: true, ...summary });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/tyre-inspections", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const params = [site_code];
      const where = ["ti.site_code = ?"];
      if (assetId > 0) {
        where.push("ti.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("ti.inspection_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("ti.inspection_date <= ?");
        params.push(end);
      }
      const rows = db.prepare(`
        SELECT
          ti.id,
          ti.asset_id,
          ti.inspection_date,
          ti.inspector_name,
          ti.running_hours,
          ti.total_tyre_cost,
          ti.cost_per_running_hour,
          ti.tyres_json,
          ti.notes,
          ti.created_at,
          a.asset_code,
          a.asset_name
        FROM tyre_inspections ti
        JOIN assets a ON a.id = ti.asset_id
        WHERE ${where.join(" AND ")}
        ORDER BY ti.inspection_date DESC, ti.id DESC
      `).all(...params);
      return reply.send({
        ok: true,
        rows: rows.map((r) => {
          let tyres = [];
          try {
            const parsed = JSON.parse(String(r.tyres_json || "[]"));
            tyres = normalizeTyreRows(parsed);
          } catch {}
          return { ...r, tyres };
        }),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/tyre-inspections", async (req, reply) => {
    try {
      ensureTyreLifecycleSchema();
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const asset_id = Number(req.body?.asset_id || 0);
      const inspection_date = String(req.body?.inspection_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(inspection_date)) return reply.code(400).send({ ok: false, error: "inspection_date must be YYYY-MM-DD" });
      const asset = db.prepare(`SELECT id FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const tyresIn = normalizeTyreRows(req.body?.tyres);
      const runningRaw = req.body?.running_hours;
      const running_hours = runningRaw == null || runningRaw === ""
        ? 0
        : Number(runningRaw);
      if (!Number.isFinite(running_hours) || running_hours < 0) {
        return reply.code(400).send({ ok: false, error: "running_hours must be a positive number" });
      }

      const thresholds = getTyreThresholds();
      const tyres = tyresIn.map((row) => enrichTyreLifecycleRow(row, {
        asset_id,
        site_code,
        inspection_date,
        running_hours,
        thresholds,
        persist: true,
      }));

      const total_tyre_cost = Number(
        tyres.reduce((sum, t) => sum + Number(t.tyre_cost || 0), 0).toFixed(2),
      );
      const withCost = tyres.filter((t) => t.cost_per_hour != null);
      const cost_per_running_hour = withCost.length
        ? Number(withCost.reduce((sum, t) => sum + Number(t.cost_per_hour || 0), 0).toFixed(4))
        : 0;
      const alerts = tyres
        .filter((t) => t.lifecycle_status === "warn" || t.lifecycle_status === "replace")
        .map((t) => ({
          position_key: t.position_key,
          position_label: t.position_label,
          lifecycle_status: t.lifecycle_status,
          tread_alert: t.tread_alert,
        }));

      const ins = db.prepare(`
        INSERT INTO tyre_inspections (
          asset_id, uuid, site_code, inspection_date, inspector_name,
          running_hours, total_tyre_cost, cost_per_running_hour, tyres_json, notes, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(
        asset_id,
        crypto.randomUUID(),
        site_code,
        inspection_date,
        inspector_name,
        running_hours,
        total_tyre_cost,
        cost_per_running_hour,
        JSON.stringify(tyres),
        notes,
      );

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        total_tyre_cost,
        cost_per_running_hour,
        fleet_cost_per_hour: cost_per_running_hour,
        tyres,
        alerts,
        thresholds,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/tyre-inspections/survey.xlsx", async (req, reply) => {
    try {
      const month = String(req.query?.month || "").trim() || new Date().toISOString().slice(0, 7);
      if (!isMonth(month)) return reply.code(400).send({ ok: false, error: "month must be YYYY-MM" });
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const data = buildTyreMonthlySurveyData(month, site_code);
      const buf = await buildTyreSurveyXlsxBuffer(data);
      reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Cache-Control", "no-store")
        .header("Content-Disposition", `attachment; filename="IRONLOG_Tyre_Survey_${month}.xlsx"`)
        .send(Buffer.from(buf));
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/tyre-inspections/survey.pdf", async (req, reply) => {
    try {
      const month = String(req.query?.month || "").trim() || new Date().toISOString().slice(0, 7);
      if (!isMonth(month)) return reply.code(400).send({ ok: false, error: "month must be YYYY-MM" });
      const isDownload = String(req.query?.download || "").trim() === "1";
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const data = buildTyreMonthlySurveyData(month, site_code);
      const branding = data.branding || getPdfReportBranding(db);
      const monthLabel = String(data.month || "");

      const pdf = await buildPdfBuffer(
        (doc) => {
          doc.y = pdfBodyTop(doc, { siteName: branding.site_name });
          doc.font("Helvetica").fontSize(10).fillColor("#334155");
          doc.text(
            `Project: ${data.project || "-"}   Country: ${data.country || "-"}   Machines: ${data.machines?.length || 0}`,
          );
          doc.moveDown(0.5);
          if (!data.machines?.length) {
            doc.text("No tyre inspections found for this month.");
            return;
          }
          for (const machine of data.machines) {
            drawTyreSurveyPdfMachine(doc, machine, branding.site_name);
          }
        },
        {
          title: "IRONLOG",
          subtitle: "Monthly Tyre Survey Report",
          rightText: monthLabel,
          showPageNumbers: true,
          layout: "landscape",
        },
      );

      reply.header("Content-Type", "application/pdf");
      reply.header("Cache-Control", "no-store");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="IRONLOG_Tyre_Survey_${monthLabel}.pdf"`,
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/undercarriage-inspections/template", async (_req, reply) => {
    return reply.send({
      ok: true,
      components: buildUndercarriageComponentSchema(),
      checklist_items: UNDERCARRIAGE_CHECKLIST_ITEMS,
      track_sag_points: UNDERCARRIAGE_TRACK_SAG_POINTS,
      wear_bands: UNDERCARRIAGE_WEAR_BANDS,
    });
  });

  app.get("/undercarriage-inspections/wear-profile", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      if (!assetId) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      const asset = db.prepare(`SELECT id, asset_code, asset_name, category FROM assets WHERE id = ?`).get(assetId);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });
      const profile = getUndercarriageWearProfileRow(assetId, site_code);
      return reply.send({
        ok: true,
        asset,
        profile,
        has_profile: Boolean(profile?.configured_count),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/undercarriage-inspections/wear-profile", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const asset_id = Number(req.body?.asset_id || 0);
      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      const asset = db.prepare(`SELECT id, asset_code FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });
      const updated_by = String(req.headers?.["x-user-name"] || req.body?.updated_by || "").trim() || null;
      const profile = saveUndercarriageWearProfile({
        asset_id,
        site_code,
        limits: req.body?.limits,
        source: String(req.body?.source || "manual").trim(),
        notes: String(req.body?.notes || "").trim() || null,
        updated_by,
      });
      return reply.send({
        ok: true,
        asset_code: asset.asset_code,
        profile,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/undercarriage-inspections/wear-profile/import-latest", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const asset_id = Number(req.body?.asset_id || req.query?.asset_id || 0);
      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      const asset = db.prepare(`SELECT id, asset_code FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });
      const updated_by = String(req.headers?.["x-user-name"] || "").trim() || null;
      const profile = importUndercarriageWearProfileFromLatest(asset_id, site_code, updated_by);
      if (!profile) {
        return reply.code(404).send({ ok: false, error: "No previous inspection found to import limits from" });
      }
      return reply.send({
        ok: true,
        asset_code: asset.asset_code,
        profile,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/undercarriage-inspections/latest", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      if (!assetId) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      const row = db.prepare(`
        SELECT ui.*, a.asset_code, a.asset_name, a.category
        FROM undercarriage_inspections ui
        JOIN assets a ON a.id = ui.asset_id
        WHERE ui.asset_id = ?
          AND LOWER(TRIM(COALESCE(ui.site_code, 'main'))) = LOWER(TRIM(?))
        ORDER BY ui.inspection_date DESC, ui.id DESC
        LIMIT 1
      `).get(assetId, site_code);
      if (!row) return reply.send({ ok: true, row: null });
      return reply.send({ ok: true, row: parseUndercarriageInspectionRow(row) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/undercarriage-inspections", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const params = [site_code];
      const where = ["LOWER(TRIM(COALESCE(ui.site_code, 'main'))) = LOWER(TRIM(?))"];
      if (assetId > 0) {
        where.push("ui.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("ui.inspection_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("ui.inspection_date <= ?");
        params.push(end);
      }
      const rows = db.prepare(`
        SELECT ui.*, a.asset_code, a.asset_name, a.category
        FROM undercarriage_inspections ui
        JOIN assets a ON a.id = ui.asset_id
        WHERE ${where.join(" AND ")}
        ORDER BY ui.inspection_date DESC, ui.id DESC
        LIMIT 200
      `).all(...params);
      return reply.send({
        ok: true,
        rows: rows.map((r) => parseUndercarriageInspectionRow(r)),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/undercarriage-inspections", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const asset_id = Number(req.body?.asset_id || 0);
      const inspection_date = String(req.body?.inspection_date || "").trim() || new Date().toISOString().slice(0, 10);
      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(inspection_date)) return reply.code(400).send({ ok: false, error: "inspection_date must be YYYY-MM-DD" });
      const asset = db.prepare(`SELECT id, asset_code, asset_name, category FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const smuRaw = req.body?.smu ?? req.body?.running_hours;
      const smu = smuRaw == null || smuRaw === "" ? null : Number(smuRaw);
      if (smu != null && (!Number.isFinite(smu) || smu < 0)) {
        return reply.code(400).send({ ok: false, error: "smu must be a positive number" });
      }

      const measurements = buildUndercarriageMeasurementsForSave(req.body?.measurements, {
        asset_id,
        site_code,
        inspection_date,
        smu,
      });
      const track_sag = normalizeUndercarriageTrackSag(req.body?.track_sag || {});
      const checklist = normalizeUndercarriageChecklist(req.body?.checklist || {});
      const summary = summarizeUndercarriageInspection(measurements);
      const branding = getPdfReportBranding(db);

      const ins = db.prepare(`
        INSERT INTO undercarriage_inspections (
          asset_id, uuid, site_code, inspection_date, inspector_name, smu,
          job_no, site_name, planner, serial_no, unit_assembly, model, yard_no,
          work_order_no, component_group, group_id, component_serial_no, part_no,
          cost_center, measurements_json, track_sag_json, checklist_json, summary_json,
          notes, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(
        asset_id,
        crypto.randomUUID(),
        site_code,
        inspection_date,
        String(req.body?.inspector_name || "").trim() || null,
        smu,
        String(req.body?.job_no || "").trim() || null,
        String(req.body?.site_name || branding.site_name || "").trim() || null,
        String(req.body?.planner || "").trim() || null,
        String(req.body?.serial_no || "").trim() || null,
        String(req.body?.unit_assembly || "").trim() || null,
        String(req.body?.model || req.body?.machine_model || asset.category || "").trim() || null,
        String(req.body?.yard_no || "").trim() || null,
        String(req.body?.work_order_no || "").trim() || null,
        String(req.body?.component_group || "").trim() || null,
        String(req.body?.group_id || "").trim() || null,
        String(req.body?.component_serial_no || "").trim() || null,
        String(req.body?.part_no || "").trim() || null,
        String(req.body?.cost_center || "").trim() || null,
        JSON.stringify(measurements),
        JSON.stringify(track_sag),
        JSON.stringify(checklist),
        JSON.stringify(summary),
        String(req.body?.notes || checklist.comments || "").trim() || null,
      );

      const updateWearProfile = req.body?.update_wear_profile !== false;
      if (updateWearProfile) {
        const limits = measurements.map((m) => ({
          key: m.key,
          base: m.base,
          wear_limit: m.wear_limit,
        }));
        saveUndercarriageWearProfile({
          asset_id,
          site_code,
          limits,
          source: "inspection_save",
          updated_by: String(req.body?.inspector_name || req.headers?.["x-user-name"] || "").trim() || null,
        });
      }

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        asset_code: asset.asset_code,
        summary,
        measurements,
        track_sag,
        checklist,
        pdf_url: `/api/maintenance/undercarriage-inspections/${Number(ins.lastInsertRowid)}.pdf`,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/undercarriage-inspections/:id.pdf", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
      const isDownload = String(req.query?.download || "").trim() === "1";
      const row = db.prepare(`
        SELECT ui.*, a.asset_code, a.asset_name, a.category
        FROM undercarriage_inspections ui
        JOIN assets a ON a.id = ui.asset_id
        WHERE ui.id = ?
        LIMIT 1
      `).get(id);
      if (!row) return reply.code(404).send({ ok: false, error: "Inspection not found" });
      const insp = parseUndercarriageInspectionRow(row);
      const branding = getPdfReportBranding(db);
      const pdf = await buildPdfBuffer(
        (doc) => drawUndercarriageInspectionPdf(doc, insp, branding),
        {
          title: "IRONLOG",
          subtitle: "Undercarriage Report",
          rightText: insp.asset_code || "",
          showPageNumbers: true,
          layout: "landscape",
        },
      );
      reply.header("Content-Type", "application/pdf");
      reply.header("Cache-Control", "no-store");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="IRONLOG_Undercarriage_${insp.asset_code || "report"}_${insp.inspection_date || id}.pdf"`,
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/undercarriage-inspections/report.xlsx", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const inspectionId = Number(req.query?.inspection_id || req.query?.id || 0);
      const assetId = Number(req.query?.asset_id || 0);
      const month = String(req.query?.month || "").trim();
      let rows = [];

      if (inspectionId > 0) {
        const one = db.prepare(`
          SELECT ui.*, a.asset_code, a.asset_name, a.category
          FROM undercarriage_inspections ui
          JOIN assets a ON a.id = ui.asset_id
          WHERE ui.id = ?
          LIMIT 1
        `).get(inspectionId);
        if (!one) return reply.code(404).send({ ok: false, error: "Inspection not found" });
        rows = [parseUndercarriageInspectionRow(one)];
      } else {
        const params = [site_code];
        const where = ["LOWER(TRIM(COALESCE(ui.site_code, 'main'))) = LOWER(TRIM(?))"];
        if (assetId > 0) {
          where.push("ui.asset_id = ?");
          params.push(assetId);
        }
        if (isMonth(month)) {
          const bounds = monthBoundsYmd(month);
          if (bounds) {
            where.push("ui.inspection_date >= ?");
            where.push("ui.inspection_date <= ?");
            params.push(bounds.start, bounds.end);
          }
        }
        const raw = db.prepare(`
          SELECT ui.*, a.asset_code, a.asset_name, a.category
          FROM undercarriage_inspections ui
          JOIN assets a ON a.id = ui.asset_id
          WHERE ${where.join(" AND ")}
          ORDER BY a.asset_code ASC, ui.inspection_date DESC, ui.id DESC
        `).all(...params);
        rows = raw.map((r) => parseUndercarriageInspectionRow(r));
      }

      if (!rows.length) return reply.code(404).send({ ok: false, error: "No inspections found for export" });
      const buf = await buildUndercarriageXlsxBuffer(rows);
      const suffix = inspectionId > 0
        ? `_${rows[0]?.asset_code || "report"}_${rows[0]?.inspection_date || inspectionId}`
        : (isMonth(month) ? `_${month}` : "");
      reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Cache-Control", "no-store")
        .header("Content-Disposition", `attachment; filename="IRONLOG_Undercarriage${suffix}.xlsx"`)
        .send(Buffer.from(buf));
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // =====================================================
  // ARTISAN INSPECTIONS (daily general checklist)
  // =====================================================
  app.get("/artisan-inspections", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const params = [];
      const where = ["LOWER(TRIM(COALESCE(ai.site_code, 'main'))) = ?"];
      params.push(site_code);
      if (assetId > 0) {
        where.push("ai.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("ai.inspection_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("ai.inspection_date <= ?");
        params.push(end);
      }

      const rows = db.prepare(`
        SELECT
          ai.id,
          ai.asset_id,
          ai.inspection_date,
          ai.inspector_name,
          ai.form_number,
          ai.shift,
          ai.notes,
          ai.machine_hours,
          ai.live_hours_snapshot,
          ai.live_hours_source,
          ai.checklist_json,
          ai.created_at,
          a.asset_code,
          a.asset_name
        FROM artisan_inspections ai
        JOIN assets a ON a.id = ai.asset_id
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY ai.inspection_date DESC, ai.id DESC
      `).all(...params);

      return reply.send({
        ok: true,
        rows: rows.map((r) => {
          let checklist = [];
          try {
            const cj = JSON.parse(String(r.checklist_json || "[]"));
            if (Array.isArray(cj)) checklist = cj;
          } catch {}
          return { ...r, checklist };
        }),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/artisan-inspections", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const inspection_date = String(req.body?.inspection_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const form_number = String(req.body?.form_number || "").trim() || null;
      const shift = String(req.body?.shift || "").trim().toLowerCase() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(inspection_date)) return reply.code(400).send({ ok: false, error: "inspection_date must be YYYY-MM-DD" });
      if (shift && !["day", "night"].includes(shift)) {
        return reply.code(400).send({ ok: false, error: "shift must be day or night" });
      }

      const asset = db.prepare(`SELECT id FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const liveInfo = getAssetHoursInfoAsOf(asset_id, inspection_date);
      const liveSnap = Number(liveInfo.hours || 0);
      const liveSource = String(liveInfo.source || "");

      let machine_hours = null;
      const mhRaw = req.body?.machine_hours;
      if (mhRaw != null && mhRaw !== "") {
        const n = Number(mhRaw);
        if (Number.isFinite(n) && n >= 0) machine_hours = n;
      }
      if (machine_hours == null) machine_hours = liveSnap;

      const checklist = normalizeArtisanInspectionChecklist(req.body?.checklist);
      const checklist_json = JSON.stringify(checklist);

      const ins = db.prepare(`
        INSERT INTO artisan_inspections (
          asset_id, uuid, site_code, inspection_date, inspector_name, form_number, shift, notes,
          machine_hours, live_hours_snapshot, live_hours_source, checklist_json, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(
        asset_id,
        crypto.randomUUID(),
        site_code,
        inspection_date,
        inspector_name,
        form_number,
        shift,
        notes,
        machine_hours,
        liveSnap,
        liveSource,
        checklist_json
      );

      return reply.send({ ok: true, id: Number(ins.lastInsertRowid) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });
}
