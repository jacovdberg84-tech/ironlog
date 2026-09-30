// IRONLOG/api/routes/reports/inspections.routes.js — Inspection, damage report and legal compliance reports.
// Registered by routes/reports.routes.js; shared helpers arrive through ctx.
import ExcelJS from "exceljs";
import fs from "node:fs";
import path from "node:path";
import { buildPdfBuffer, ensurePageSpace, kvGrid, sectionTitle, table, tryDrawLogo } from "../../utils/pdfGenerator.js";
import { cleanupTempPdfImages, drawPhotoInPdf } from "../../utils/imagePdf.js";
import { db } from "../../db/client.js";
import { isDate } from "../../utils/request.js";
import { verifyRecordToken } from "../../utils/signedLinks.js";

export default function registerInspectionsRoutes(app, ctx) {
  const {
    buildVehicleLdvChecklistPdfRows,
    compactCell,
    fmtNum,
    hasColumn,
    machinePrestartProfileFromCheckMode,
    makeArtisanFormNumber,
    pickExistingColumn,
    resolveCheckPhotoForPdf,
    resolveStorageAbs,
    todayYmd,
  } = ctx;

  // GET /api/reports/vehicle-ldv-check/:id.pdf?download=1
  async function sendVehicleCheckPdf(req, reply) {
    const id = Number(req.params?.id || 0);
    const download = String(req.query?.download || "").trim() === "1";
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).send({ error: "valid vehicle check id required" });
    }

    const check = db.prepare(`
      SELECT
        v.id,
        v.check_date,
        v.vehicle_registration,
        v.odometer_km,
        v.smu_hours,
        v.inspector_name,
        v.notes,
        v.check_mode,
        v.checklist_json,
        v.created_at,
        a.asset_code,
        a.asset_name,
        a.category
      FROM vehicle_ldv_checks v
      JOIN assets a ON a.id = v.asset_id
      WHERE v.id = ?
    `).get(id);
    if (!check) return reply.code(404).send({ error: "vehicle check not found" });

    const isMachinePrestart = Boolean(machinePrestartProfileFromCheckMode(check.check_mode));

    const photos = db.prepare(`
      SELECT id, file_path, caption, markers_json, created_at
      FROM vehicle_ldv_check_photos
      WHERE check_id = ?
      ORDER BY id ASC
    `).all(id).map((p) => {
      let markers = [];
      try {
        markers = p.markers_json ? JSON.parse(p.markers_json) : [];
      } catch {
        markers = [];
      }
      return { ...p, markers: Array.isArray(markers) ? markers : [] };
    });
    let checklistRows = [];
    try {
      checklistRows = buildVehicleLdvChecklistPdfRows(check.check_mode, check.checklist_json);
    } catch {
      checklistRows = [];
    }

    const tempPdfImages = [];
    const photosForPdf = [];
    for (const p of photos) {
      const resolved = await resolveCheckPhotoForPdf(p);
      if (resolved.tempPath) tempPdfImages.push(resolved.tempPath);
      photosForPdf.push(resolved);
    }

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, isMachinePrestart ? "Machine pre-start" : "LDV Vehicle Check");
        kvGrid(doc, [
          { k: "Check #", v: check.id },
          { k: "Date", v: check.check_date || "" },
          { k: "Vehicle", v: `${check.asset_code || ""} - ${check.asset_name || ""}`.trim() },
          { k: "Category", v: check.category || "" },
          { k: "Registration", v: check.vehicle_registration || "-" },
          { k: "Mode", v: check.check_mode || "ldv_general" },
          { k: "Odometer (km)", v: check.odometer_km == null ? "-" : fmtNum(check.odometer_km, 0) },
          { k: "SMU (hours)", v: check.smu_hours == null ? "-" : fmtNum(check.smu_hours, 1) },
          { k: "Inspector", v: check.inspector_name || "-" },
          { k: "Created At", v: check.created_at || "" },
        ], 2);

        if (checklistRows.length) {
          sectionTitle(doc, "Pre-Start Checklist");
          table(
            doc,
            [
              { key: "item", label: "Item", width: 0.72 },
              { key: "status", label: "Result", width: 0.28, align: "center" },
            ],
            checklistRows
          );
        }

        sectionTitle(doc, "Notes");
        doc
          .font("Helvetica")
          .fontSize(10)
          .fillColor("#111111")
          .text(compactCell(check.notes || "-", 2000), {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });

        sectionTitle(doc, "Photos and Pinned Damages");
        if (!photosForPdf.length) {
          doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No photos attached.");
          return;
        }

        for (const p of photosForPdf) {
          ensurePageSpace(doc, 300);
          doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
          doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`, {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });
          doc.moveDown(0.2);
          drawPhotoInPdf(doc, p.pdfPath, { missingLabel: p.rel || p.file_path || "-", maxWidth: 420, maxHeight: 190 });

          const markers = Array.isArray(p.markers) ? p.markers : [];
          if (!markers.length) {
            doc.font("Helvetica").fontSize(9).fillColor("#555555").text("No pinned damages on this photo.");
            doc.moveDown(0.4);
            continue;
          }

          doc.font("Helvetica-Bold").fontSize(9).fillColor("#111111").text("Pinned damages:");
          doc.moveDown(0.15);
          markers.forEach((m, idx) => {
            const label = compactCell(m?.label || "Damage", 80);
            const note = compactCell(m?.note || "", 160);
            doc
              .font("Helvetica")
              .fontSize(9)
              .fillColor("#111111")
              .text(
                `${idx + 1}. ${label}${note ? ` — ${note}` : ""} (x:${fmtNum((Number(m?.x) || 0) * 100, 1)}%, y:${fmtNum((Number(m?.y) || 0) * 100, 1)}%)`,
                {
                  width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
                }
              );
          });
          doc.moveDown(0.5);
        }
      },
      {
        title: "IRONLOG",
        subtitle: isMachinePrestart ? "Machine pre-start report" : "LDV Vehicle Check Report",
        rightText: `Check #${check.id}`,
        showPageNumbers: true,
      }
    );

    cleanupTempPdfImages(tempPdfImages);

    const pdfBase = isMachinePrestart ? "AML_Machine_Prestart" : "AML_LDV_Check";
    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="${pdfBase}_${check.id}.pdf"`
      )
      .send(pdf);
  }

  app.get("/vehicle-ldv-check/:id.pdf", sendVehicleCheckPdf);

  // GET /api/reports/prestart-check/:id.pdf?t=<signed token>
  // No login: operators open the check they just submitted from the QR page.
  // Only a valid, unexpired signature for this check id opens it.
  app.get("/prestart-check/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!verifyRecordToken(db, "prestart-pdf", id, req.query?.t)) {
      return reply.code(403).send({ error: "This PDF link has expired or is not valid. Open the pre-start again to get a new link." });
    }
    return sendVehicleCheckPdf(req, reply);
  });

  // GET /api/reports/vehicle-ldv-checks.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&asset_id=123&with_photos=1&download=1
  app.get("/vehicle-ldv-checks.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const assetId = Number(req.query?.asset_id || 0);
    const withPhotos = String(req.query?.with_photos || "").trim() === "1";
    const download = String(req.query?.download || "").trim() === "1";

    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const where = ["v.check_date >= ?", "v.check_date <= ?"];
    const params = [start, end];
    if (assetId > 0) {
      where.push("v.asset_id = ?");
      params.push(assetId);
    }

    const rows = db.prepare(`
      SELECT
        v.id,
        v.asset_id,
        v.check_date,
        v.vehicle_registration,
        v.odometer_km,
        v.inspector_name,
        v.notes,
        a.asset_code,
        a.asset_name
      FROM vehicle_ldv_checks v
      JOIN assets a ON a.id = v.asset_id
      WHERE ${where.join(" AND ")}
      ORDER BY v.check_date DESC, v.id DESC
      LIMIT 1000
    `).all(...params);

    const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
    const photosByCheck = new Map();
    if (ids.length) {
      const marks = ids.map(() => "?").join(",");
      const photos = db.prepare(`
        SELECT check_id, id, file_path, caption, markers_json, created_at
        FROM vehicle_ldv_check_photos
        WHERE check_id IN (${marks})
        ORDER BY check_id ASC, id ASC
      `).all(...ids);
      for (const p of photos) {
        const key = Number(p.check_id);
        if (!photosByCheck.has(key)) photosByCheck.set(key, []);
        let markers = [];
        try {
          markers = p.markers_json ? JSON.parse(p.markers_json) : [];
        } catch {
          markers = [];
        }
        photosByCheck.get(key).push({
          ...p,
          markers: Array.isArray(markers) ? markers : [],
        });
      }
    }

    const summary = {
      count: rows.length,
      vehicles: new Set(rows.map((r) => Number(r.asset_id))).size,
      inspectors: new Set(rows.map((r) => String(r.inspector_name || "").trim()).filter(Boolean)).size,
      photos: rows.reduce((acc, r) => acc + (photosByCheck.get(Number(r.id)) || []).length, 0),
      pins: rows.reduce(
        (acc, r) => acc + (photosByCheck.get(Number(r.id)) || []).reduce((n, p) => n + ((p.markers || []).length || 0), 0),
        0
      ),
    };

    const tempPdfImages = [];
    const photosByCheckPdf = new Map();
    if (withPhotos) {
      for (const [checkId, photos] of photosByCheck.entries()) {
        const resolvedPhotos = [];
        for (const p of photos) {
          const resolved = await resolveCheckPhotoForPdf(p);
          if (resolved.tempPath) tempPdfImages.push(resolved.tempPath);
          resolvedPhotos.push(resolved);
        }
        photosByCheckPdf.set(checkId, resolvedPhotos);
      }
    }

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "LDV Vehicle Checks Summary");
        kvGrid(doc, [
          { k: "From", v: start },
          { k: "To", v: end },
          { k: "Checks", v: summary.count },
          { k: "Vehicles", v: summary.vehicles },
          { k: "Inspectors", v: summary.inspectors },
          { k: "Photos", v: summary.photos },
          { k: "Damage pins", v: summary.pins },
        ], 2);

        sectionTitle(doc, "Checks");
        table(
          doc,
          ["ID", "Date", "Vehicle", "Reg", "Inspector", "Photos", "Pins", "Notes"],
          rows.length
            ? rows.map((r) => {
                const photos = photosByCheck.get(Number(r.id)) || [];
                const pinCount = photos.reduce((n, p) => n + ((p.markers || []).length || 0), 0);
                return {
                  ID: r.id,
                  Date: r.check_date || "",
                  Vehicle: `${r.asset_code || ""} ${r.asset_name ? `- ${compactCell(r.asset_name, 30)}` : ""}`.trim(),
                  Reg: compactCell(r.vehicle_registration || "-", 16),
                  Inspector: compactCell(r.inspector_name || "-", 20),
                  Photos: photos.length,
                  Pins: pinCount,
                  Notes: compactCell(r.notes || "-", 40),
                };
              })
            : [{ ID: "-", Date: "-", Vehicle: "No checks found in selected period", Reg: "-", Inspector: "-", Photos: "-", Pins: "-", Notes: "-" }]
        );

        if (!withPhotos) return;

        for (const r of rows) {
          const photos = photosByCheckPdf.get(Number(r.id)) || [];
          if (!photos.length) continue;
          ensurePageSpace(doc, 40);
          sectionTitle(doc, `Check #${r.id} — ${r.asset_code || ""} (${r.check_date || ""})`);
          for (const p of photos) {
            ensurePageSpace(doc, 240);
            doc.font("Helvetica-Bold").fontSize(9).fillColor("#111111")
              .text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""} • pins: ${(p.markers || []).length}`);
            doc.moveDown(0.15);
            drawPhotoInPdf(doc, p.pdfPath, { missingLabel: p.rel || p.file_path || "-", maxWidth: 360, maxHeight: 150 });
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: "LDV Vehicle Checks",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    cleanupTempPdfImages(tempPdfImages);

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_LDV_Checks_${start}_to_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/artisan-inspection-form.pdf?asset_id=123&date=YYYY-MM-DD&shift=day|night&inspector_name=...&form_number=...&download=1
  app.get("/artisan-inspection-form.pdf", async (req, reply) => {
    const download = String(req.query?.download || "").trim() === "1";
    const assetId = Number(req.query?.asset_id || 0);
    const date = String(req.query?.date || "").trim();
    const shift = String(req.query?.shift || "").trim().toLowerCase();
    const inspectorName = String(req.query?.inspector_name || "").trim();
    const formNumber = String(req.query?.form_number || "").trim() || makeArtisanFormNumber();

    let asset = null;
    if (assetId > 0) {
      asset = db.prepare(`SELECT id, asset_code, asset_name, category FROM assets WHERE id = ?`).get(assetId);
    }

    const safeDate = isDate(date) ? date : "";
    const safeShift = ["day", "night"].includes(shift) ? shift.toUpperCase() : "";
    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const checklistRows = [
      "Pre-start visual condition (machine / plant)",
      "Guards, covers, and safety devices",
      "Hydraulic hoses, leaks, and fittings",
      "Electrical panels / cabling / lights",
      "Lubrication points / levels",
      "Brakes / steering / controls response",
      "Alarms, horn, and warning systems",
      "Housekeeping around machine / plant",
    ];

    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);
        sectionTitle(doc, "Daily Artisan Inspection (Blank Form)");
        kvGrid(doc, [
          { k: "Form No.", v: formNumber },
          { k: "Date", v: safeDate || "________________" },
          { k: "Shift", v: safeShift || "________________" },
          { k: "Artisan", v: inspectorName || "________________" },
          { k: "Asset Code", v: asset?.asset_code || "________________" },
          { k: "Asset Name", v: asset?.asset_name || "________________" },
          { k: "Category", v: asset?.category || "________________" },
          { k: "Machine Hours", v: "________________" },
        ], 2);

        sectionTitle(doc, "General checklist (tick one)");
        doc.font("Helvetica").fontSize(10).fillColor("#111111");
        const leftX = doc.page.margins.left;
        const contentW = doc.page.width - doc.page.margins.left - doc.page.margins.right;
        const colOkX = leftX + contentW * 0.58;
        const colFailX = leftX + contentW * 0.68;
        const colNaX = leftX + contentW * 0.80;
        const noteLabelX = leftX + contentW * 0.06;
        const noteLineX = leftX + contentW * 0.16;
        const noteLineW = contentW * 0.80;
        doc.font("Helvetica-Bold").fontSize(9);
        const headY = doc.y;
        doc.text("OK", colOkX, headY);
        doc.text("FAIL", colFailX, headY);
        doc.text("N/A", colNaX, headY);
        doc.moveDown(0.6);
        doc.font("Helvetica").fontSize(10);
        checklistRows.forEach((label, idx) => {
          ensurePageSpace(doc, 42);
          const rowY = doc.y;
          doc.text(`${idx + 1}. ${label}`, leftX, rowY, { width: contentW * 0.55 });
          doc.text("[ ]", colOkX, rowY);
          doc.text("[ ]", colFailX, rowY);
          doc.text("[ ]", colNaX, rowY);
          const noteY = rowY + 13;
          doc.text("Note:", noteLabelX, noteY);
          doc.moveTo(noteLineX, noteY + 10).lineTo(noteLineX + noteLineW, noteY + 10).strokeColor("#777777").lineWidth(0.6).stroke();
          doc.y = noteY + 14;
          doc.moveDown(0.2);
        });

        sectionTitle(doc, "Notes");
        for (let i = 0; i < 5; i += 1) {
          doc.text("________________________________________________________________________________________", {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });
          doc.moveDown(0.2);
        }

        doc.moveDown(0.6);
        doc.text("Artisan signature: ____________________________    Time returned: ____________________", {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });
        doc.moveDown(0.2);
        doc.text("Supervisor received by: _______________________    Date: _____________________________", {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });
      },
      {
        title: "IRONLOG",
        subtitle: "Artisan Inspection Blank Form",
        rightText: `Form ${formNumber}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Artisan_Inspection_Form_${formNumber}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/artisan-inspection/:id.pdf?download=1
  app.get("/artisan-inspection/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    const download = String(req.query?.download || "").trim() === "1";
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).send({ error: "valid inspection id required" });
    }

    const aiInspectorCol = pickExistingColumn("artisan_inspections", ["inspector_name", "inspector"], "inspector_name");
    const hasAiMachineHours = hasColumn("artisan_inspections", "machine_hours");
    const hasAiLiveSnap = hasColumn("artisan_inspections", "live_hours_snapshot");
    const hasAiShift = hasColumn("artisan_inspections", "shift");
    const hasAiChecklist = hasColumn("artisan_inspections", "checklist_json");
    const hasAiLiveSource = hasColumn("artisan_inspections", "live_hours_source");
    const legacyMeterSql = `COALESCE((
          SELECT MAX(dh.closing_hours)
          FROM daily_hours dh
          WHERE dh.asset_id = ai.asset_id
            AND dh.closing_hours IS NOT NULL
            AND dh.work_date <= ai.inspection_date
        ), 0)`;
    const machineHoursSelect = hasAiMachineHours
      ? `COALESCE(ai.machine_hours, ${legacyMeterSql})`
      : legacyMeterSql;
    const liveSnapSelect = hasAiLiveSnap
      ? `COALESCE(ai.live_hours_snapshot, ${legacyMeterSql})`
      : legacyMeterSql;
    const liveSourceSelect = hasAiLiveSource ? "ai.live_hours_source" : "''";

    const inspection = db.prepare(`
      SELECT
        ai.id,
        ai.inspection_date,
        ai.${aiInspectorCol} AS inspector_name,
        ${hasColumn("artisan_inspections", "form_number") ? "ai.form_number" : "''"} AS form_number,
        ai.notes,
        ${hasAiShift ? "ai.shift" : "''"} AS shift,
        ${machineHoursSelect} AS machine_hours,
        ${liveSnapSelect} AS live_hours_snapshot,
        ${liveSourceSelect} AS live_hours_source,
        ${hasAiChecklist ? "ai.checklist_json" : `''`} AS checklist_json,
        a.asset_code,
        a.asset_name,
        a.category
      FROM artisan_inspections ai
      JOIN assets a ON a.id = ai.asset_id
      WHERE ai.id = ?
    `).get(id);
    if (!inspection) return reply.code(404).send({ error: "artisan inspection not found" });

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Daily Artisan Inspection");
        kvGrid(doc, [
          { k: "Inspection #", v: inspection.id },
          { k: "Form No.", v: inspection.form_number || "—" },
          { k: "Date", v: inspection.inspection_date || "" },
          { k: "Shift", v: inspection.shift ? String(inspection.shift).toUpperCase() : "—" },
          { k: "Inspector", v: inspection.inspector_name || "-" },
          { k: "Asset Code", v: inspection.asset_code || "" },
          { k: "Asset Name", v: inspection.asset_name || "" },
          { k: "Recorded machine hours", v: Number(inspection.machine_hours || 0).toFixed(1) },
          {
            k: "Live hours (snapshot)",
            v: `${Number(inspection.live_hours_snapshot ?? inspection.machine_hours ?? 0).toFixed(1)}${
              inspection.live_hours_source ? ` (${inspection.live_hours_source})` : ""
            }`,
          },
          { k: "Category", v: inspection.category || "" },
        ], 2);

        let checklist = [];
        try {
          const cj = JSON.parse(String(inspection.checklist_json || "[]"));
          if (Array.isArray(cj)) checklist = cj;
        } catch {}
        if (checklist.length) {
          sectionTitle(doc, "General checklist");
          doc.font("Helvetica").fontSize(10).fillColor("#111111");
          for (const c of checklist) {
            const st = c.ok === true ? "OK" : c.ok === false ? "FAIL" : "N/A";
            doc.text(`• ${String(c.label || c.key || "")}: ${st}${c.note ? ` — ${c.note}` : ""}`, {
              width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
            });
            doc.moveDown(0.15);
          }
        }

        sectionTitle(doc, "Notes");
        doc
          .font("Helvetica")
          .fontSize(10)
          .fillColor("#111111")
          .text(compactCell(inspection.notes || "-", 2000), {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });
      },
      {
        title: "IRONLOG",
        subtitle: "Artisan Inspection Report",
        rightText: `Inspection #${inspection.id}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Artisan_Inspection_${inspection.id}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/manager-inspection/:id.pdf?download=1
  app.get("/manager-inspection/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    const download = String(req.query?.download || "").trim() === "1";
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).send({ error: "valid inspection id required" });
    }

    const miInspectorCol = pickExistingColumn("manager_inspections", ["inspector_name", "inspector"], "inspector_name");
    const miNotesCol = pickExistingColumn("manager_inspections", ["notes", "note", "remarks", "description"], "notes");
    const hasMiChecklistDetail = hasColumn("manager_inspections", "checklist_detail_json");
    const photoInspectionCol = pickExistingColumn("manager_inspection_photos", ["inspection_id", "manager_inspection_id"], "inspection_id");
    const photoPathCol = pickExistingColumn("manager_inspection_photos", ["file_path", "photo_path", "path", "image_path", "url"], "file_path");
    const photoCaptionCol = pickExistingColumn("manager_inspection_photos", ["caption", "note", "notes", "description"], "caption");
    const photoCreatedCol = pickExistingColumn("manager_inspection_photos", ["created_at", "uploaded_at", "created_on"], "created_at");

    const hasMiMachineHours = hasColumn("manager_inspections", "machine_hours");
    const hasMiLiveSnap = hasColumn("manager_inspections", "live_hours_snapshot");
    const hasMiChecklist = hasColumn("manager_inspections", "checklist_json");
    const hasMiParts = hasColumn("manager_inspections", "required_parts_json");
    const hasMiWo = hasColumn("manager_inspections", "work_order_id");
    const legacyMeterSql = `COALESCE((
          SELECT MAX(dh.closing_hours)
          FROM daily_hours dh
          WHERE dh.asset_id = mi.asset_id
            AND dh.closing_hours IS NOT NULL
            AND dh.work_date <= mi.inspection_date
        ), 0)`;
    const machineHoursSelect = hasMiMachineHours
      ? `COALESCE(mi.machine_hours, ${legacyMeterSql})`
      : legacyMeterSql;
    const liveSnapSelect = hasMiLiveSnap
      ? `COALESCE(mi.live_hours_snapshot, ${legacyMeterSql})`
      : legacyMeterSql;
    const liveSrcSelect = hasColumn("manager_inspections", "live_hours_source")
      ? `mi.live_hours_source`
      : `''`;

    const inspection = db.prepare(`
      SELECT
        mi.id,
        mi.inspection_date,
        mi.${miInspectorCol} AS inspector_name,
        mi.${miNotesCol} AS notes,
        ${hasMiChecklistDetail ? "mi.checklist_detail_json" : `''`} AS checklist_detail_json,
        mi.created_at,
        ${machineHoursSelect} AS machine_hours,
        ${liveSnapSelect} AS live_hours_snapshot,
        ${liveSrcSelect} AS live_hours_source,
        ${hasMiChecklist ? "mi.checklist_json" : `''`} AS checklist_json,
        ${hasMiParts ? "mi.required_parts_json" : `''`} AS required_parts_json,
        ${hasMiWo ? "mi.work_order_id" : `NULL`} AS work_order_id,
        a.asset_code,
        a.asset_name,
        a.category
      FROM manager_inspections mi
      JOIN assets a ON a.id = mi.asset_id
      WHERE mi.id = ?
    `).get(id);
    if (!inspection) return reply.code(404).send({ error: "manager inspection not found" });

    const photos = db.prepare(`
      SELECT id, ${photoPathCol} AS file_path, ${photoCaptionCol} AS caption, ${photoCreatedCol} AS created_at
      FROM manager_inspection_photos
      WHERE ${photoInspectionCol} = ?
      ORDER BY id ASC
    `).all(id);

    const toChecklistLabel = (key) =>
      String(key || "")
        .replaceAll("_", " ")
        .replace(/\b\w/g, (m) => m.toUpperCase())
        .trim();
    const parseChecklistOkStatus = (status) => {
      if (status === true) return true;
      if (status === false) return false;
      if (typeof status === "number") {
        if (status === 1) return true;
        if (status === 0) return false;
      }
      const st = String(status || "").trim().toLowerCase();
      if (!st) return null;
      if (["ok", "pass", "passed", "true", "yes", "good"].includes(st)) return true;
      if (["attention", "unsafe", "fail", "failed", "false", "no", "bad", "fault"].includes(st)) return false;
      return null;
    };
    const parseManagerChecklist = (rawChecklist, rawDetails) => {
      let checklistParsed = null;
      let detailsParsed = null;
      try {
        checklistParsed = JSON.parse(String(rawChecklist || "null"));
      } catch {}
      try {
        detailsParsed = JSON.parse(String(rawDetails || "null"));
      } catch {}

      // Legacy/mobile ingest bundle shape: { checklist, checklist_details, ... }.
      if (checklistParsed && !Array.isArray(checklistParsed) && typeof checklistParsed === "object") {
        if (checklistParsed.checklist_details && (detailsParsed == null || typeof detailsParsed !== "object")) {
          detailsParsed = checklistParsed.checklist_details;
        }
        if (checklistParsed.checklist && typeof checklistParsed.checklist === "object") {
          checklistParsed = checklistParsed.checklist;
        }
      }

      if (Array.isArray(checklistParsed)) {
        return checklistParsed.map((c) => ({
          key: String(c?.key || "").trim(),
          label: String(c?.label || c?.key || "").trim(),
          ok: c?.ok === true ? true : c?.ok === false ? false : null,
          note: String(c?.note || "").trim() || null,
        }));
      }

      if (checklistParsed && typeof checklistParsed === "object") {
        return Object.entries(checklistParsed).map(([key, status]) => {
          const st = String(status || "").trim().toLowerCase();
          const ok = st === "ok" ? true : (st === "attention" || st === "unsafe" || st === "fail" || st === "failed") ? false : null;
          const detail = detailsParsed && typeof detailsParsed === "object" ? detailsParsed[key] : null;
          const note = String(detail?.comment || detail?.note || detail?.notes || "").trim() || null;
          return { key, label: toChecklistLabel(key), ok, note };
        });
      }
      return [];
    };

    const parseComponentNotes = (raw) => {
      const text = String(raw || "").trim();
      if (!text) return [];
      return text
        .split(/\r?\n|;/)
        .map((s) => String(s || "").trim())
        .filter(Boolean)
        .map((line) => {
          const m = line.match(/^([^:|-]+)\s*[:|-]\s*(.+)$/);
          if (m) {
            return { component: String(m[1] || "").trim(), note: String(m[2] || "").trim() };
          }
          return { component: "General", note: line };
        });
    };
    const extractChecklistDetailNotes = (raw) => {
      const text = String(raw || "").trim();
      if (!text) return [];
      try {
        const parsed = JSON.parse(text);
        const out = [];
        const walk = (node, label = "") => {
          if (node == null) return;
          if (Array.isArray(node)) {
            for (const item of node) walk(item, label);
            return;
          }
          if (typeof node === "object") {
            const comp = String(node.component || node.label || node.key || label || "General").trim();
            const note = String(node.note || node.notes || node.description || node.comment || "").trim();
            if (note) out.push({ component: comp, note });
            for (const [k, v] of Object.entries(node)) {
              if (["component", "label", "key", "note", "notes", "description", "comment"].includes(k)) continue;
              walk(v, comp || k);
            }
          }
        };
        walk(parsed);
        return out;
      } catch {
        return [];
      }
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Manager Inspection");
        kvGrid(doc, [
          { k: "Inspection #", v: inspection.id },
          { k: "Date", v: inspection.inspection_date || "" },
          { k: "Inspector", v: inspection.inspector_name || "-" },
          { k: "Asset Code", v: inspection.asset_code || "" },
          { k: "Asset Name", v: inspection.asset_name || "" },
          { k: "Recorded machine hours", v: Number(inspection.machine_hours || 0).toFixed(1) },
          {
            k: "Live hours (snapshot)",
            v: `${Number(inspection.live_hours_snapshot ?? inspection.machine_hours ?? 0).toFixed(1)}${
              inspection.live_hours_source ? ` (${inspection.live_hours_source})` : ""
            }`,
          },
          { k: "Work order", v: inspection.work_order_id ? `#${inspection.work_order_id}` : "—" },
          { k: "Category", v: inspection.category || "" },
        ], 2);

        const checklist = parseManagerChecklist(inspection.checklist_json, inspection.checklist_detail_json);
        if (checklist.length) {
          sectionTitle(doc, "Checklist");
          table(
            doc,
            [
              { key: "component", label: "Component", width: 0.4 },
              { key: "status", label: "Status", width: 0.15, align: "center" },
              { key: "note", label: "Note", width: 0.45 },
            ],
            checklist.map((c) => ({
              component: compactCell(String(c.label || c.key || "-"), 80),
              status: c.ok === true ? "OK" : c.ok === false ? "FAIL" : "N/A",
              note: compactCell(String(c.note || "-"), 150),
            })),
            { fontSize: 9, compact: true }
          );
        }

        let reqParts = [];
        try {
          const pj = JSON.parse(String(inspection.required_parts_json || "[]"));
          if (Array.isArray(pj)) reqParts = pj;
        } catch {}
        if (reqParts.length) {
          sectionTitle(doc, "Required parts");
          doc.font("Helvetica").fontSize(10).fillColor("#111111");
          for (const p of reqParts) {
            doc.text(
              `• ${String(p.part_code || "")} × ${Number(p.qty || 0)}${p.note ? ` — ${p.note}` : ""}`,
              { width: doc.page.width - doc.page.margins.left - doc.page.margins.right }
            );
            doc.moveDown(0.15);
          }
        }

        const rawNotes = String(inspection.notes || "").trim();
        let noteRows = parseComponentNotes(rawNotes);
        if (!noteRows.length) {
          noteRows = extractChecklistDetailNotes(inspection.checklist_detail_json);
        }
        sectionTitle(doc, "Notes");
        if (rawNotes) {
          doc
            .font("Helvetica")
            .fontSize(10)
            .fillColor("#111111")
            .text(rawNotes, {
              width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
            });
          doc.moveDown(0.2);
        }
        if (noteRows.length) {
          table(
            doc,
            [
              { key: "component", label: "Component", width: 0.3 },
              { key: "notes", label: "Notes", width: 0.7 },
            ],
            noteRows.map((n) => ({
              component: compactCell(n.component || "General", 80),
              notes: compactCell(n.note || "-", 180),
            })),
            { fontSize: 9, compact: true }
          );
        } else {
          doc
            .font("Helvetica")
            .fontSize(10)
            .fillColor("#111111")
            .text("-", {
              width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
            });
        }

        sectionTitle(doc, "Photos");
        if (!photos.length) {
          doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No photos attached.");
          return;
        }

        for (const p of photos) {
          const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
          const abs = resolveStorageAbs(rel);
          ensurePageSpace(doc, 230);
          doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
          doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`, {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });
          doc.moveDown(0.2);
          if (abs && fs.existsSync(abs)) {
            try {
              doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 180], align: "left", valign: "top" });
              doc.y += 186;
            } catch {
              doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo file exists but could not be rendered.");
              doc.moveDown(0.5);
            }
          } else {
            doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
            doc.moveDown(0.5);
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: "Manager Inspection Report",
        rightText: `Inspection #${inspection.id}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Manager_Inspection_${inspection.id}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/manager-inspections.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&asset_id=123&with_photos=1&download=1
  app.get("/manager-inspections.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const assetId = Number(req.query?.asset_id || 0);
    const withPhotos = String(req.query?.with_photos || "").trim() === "1";
    const download = String(req.query?.download || "").trim() === "1";

    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const miInspectorCol = pickExistingColumn("manager_inspections", ["inspector_name", "inspector"], "inspector_name");
    const miNotesCol = pickExistingColumn("manager_inspections", ["notes", "note", "remarks", "description"], "notes");
    const hasMiChecklistDetail = hasColumn("manager_inspections", "checklist_detail_json");
    const photoInspectionCol = pickExistingColumn("manager_inspection_photos", ["inspection_id", "manager_inspection_id"], "inspection_id");
    const photoPathCol = pickExistingColumn("manager_inspection_photos", ["file_path", "photo_path", "path", "image_path", "url"], "file_path");
    const photoCaptionCol = pickExistingColumn("manager_inspection_photos", ["caption", "note", "notes", "description"], "caption");
    const photoCreatedCol = pickExistingColumn("manager_inspection_photos", ["created_at", "uploaded_at", "created_on"], "created_at");

    const where = ["mi.inspection_date >= ?", "mi.inspection_date <= ?"];
    const params = [start, end];
    if (assetId > 0) {
      where.push("mi.asset_id = ?");
      params.push(assetId);
    }

    const hasMiMh = hasColumn("manager_inspections", "machine_hours");
    const hasMiWo = hasColumn("manager_inspections", "work_order_id");
    const hasMiChecklist = hasColumn("manager_inspections", "checklist_json");
    const toChecklistLabel = (key) =>
      String(key || "")
        .replaceAll("_", " ")
        .replace(/\b\w/g, (m) => m.toUpperCase())
        .trim();
    const parseManagerChecklist = (rawChecklist, rawDetails) => {
      let checklistParsed = null;
      let detailsParsed = null;
      try {
        checklistParsed = JSON.parse(String(rawChecklist || "null"));
      } catch {}
      try {
        detailsParsed = JSON.parse(String(rawDetails || "null"));
      } catch {}
      if (checklistParsed && !Array.isArray(checklistParsed) && typeof checklistParsed === "object") {
        if (checklistParsed.checklist_details && (detailsParsed == null || typeof detailsParsed !== "object")) {
          detailsParsed = checklistParsed.checklist_details;
        }
        if (checklistParsed.checklist && typeof checklistParsed.checklist === "object") {
          checklistParsed = checklistParsed.checklist;
        }
      }
      if (Array.isArray(checklistParsed)) {
        return checklistParsed.map((c) => ({
          key: String(c?.key || "").trim(),
          label: String(c?.label || c?.key || "").trim(),
          ok: parseChecklistOkStatus(c?.ok),
          note: String(c?.note || c?.comment || c?.notes || "").trim() || null,
        }));
      }
      if (checklistParsed && typeof checklistParsed === "object") {
        return Object.entries(checklistParsed).map(([key, status]) => {
          const statusObj = status && typeof status === "object" && !Array.isArray(status) ? status : null;
          const ok = parseChecklistOkStatus(statusObj ? (statusObj.ok ?? statusObj.status ?? statusObj.value) : status);
          const detail = detailsParsed && typeof detailsParsed === "object" ? detailsParsed[key] : null;
          const note = String(
            statusObj?.note ||
            statusObj?.comment ||
            statusObj?.notes ||
            detail?.comment ||
            detail?.note ||
            detail?.notes ||
            ""
          ).trim() || null;
          return { key, label: toChecklistLabel(key), ok, note };
        });
      }
      return [];
    };
    const rows = db.prepare(`
      SELECT
        mi.id,
        mi.asset_id,
        mi.inspection_date,
        mi.${miInspectorCol} AS inspector_name,
        mi.${miNotesCol} AS notes,
        ${hasMiChecklistDetail ? "mi.checklist_detail_json" : `''`} AS checklist_detail_json,
        ${hasMiMh ? "mi.machine_hours" : "NULL"} AS machine_hours,
        ${hasMiChecklist ? "mi.checklist_json" : `''`} AS checklist_json,
        ${hasMiWo ? "mi.work_order_id" : "NULL"} AS work_order_id,
        a.asset_code,
        a.asset_name
      FROM manager_inspections mi
      JOIN assets a ON a.id = mi.asset_id
      WHERE ${where.join(" AND ")}
      ORDER BY mi.inspection_date DESC, mi.id DESC
      LIMIT 1000
    `).all(...params);

    const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
    const photosByInspection = new Map();
    if (ids.length) {
      const marks = ids.map(() => "?").join(",");
      const photos = db.prepare(`
        SELECT ${photoInspectionCol} AS inspection_id, id, ${photoPathCol} AS file_path, ${photoCaptionCol} AS caption, ${photoCreatedCol} AS created_at
        FROM manager_inspection_photos
        WHERE ${photoInspectionCol} IN (${marks})
        ORDER BY ${photoInspectionCol} ASC, id ASC
      `).all(...ids);
      for (const p of photos) {
        const key = Number(p.inspection_id);
        if (!photosByInspection.has(key)) photosByInspection.set(key, []);
        photosByInspection.get(key).push(p);
      }
    }

    const summary = {
      count: rows.length,
      assets: new Set(rows.map((r) => Number(r.asset_id))).size,
      inspectors: new Set(rows.map((r) => String(r.inspector_name || "").trim()).filter(Boolean)).size,
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Manager Inspections Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Asset Filter", v: assetId > 0 ? String(assetId) : "All assets" },
          { k: "Inspections", v: String(summary.count) },
          { k: "Assets Covered", v: String(summary.assets) },
          { k: "Inspectors", v: String(summary.inspectors) },
        ], 2);

        sectionTitle(doc, "Inspection Entries");
        table(
          doc,
          [
            { key: "id", label: "ID", width: 0.07, align: "right" },
            { key: "date", label: "Date", width: 0.11 },
            { key: "asset", label: "Asset", width: 0.14 },
            { key: "name", label: "Asset Name", width: 0.18 },
            { key: "hrs", label: "Hrs", width: 0.07, align: "right" },
            { key: "wo", label: "WO", width: 0.07, align: "right" },
            { key: "inspector", label: "Inspector", width: 0.12 },
            { key: "notes", label: "Notes", width: 0.24 },
          ],
          rows.length
            ? rows.map((r) => {
                const checklist = parseManagerChecklist(r.checklist_json, r.checklist_detail_json);
                const failedChecklist = checklist.filter((c) => c && c.ok === false);
                const checklistFindings = failedChecklist
                  .map((c) => {
                    const label = String(c.label || c.key || "Item").trim();
                    const note = String(c.note || "").trim();
                    return note ? `${label}: ${note}` : label;
                  })
                  .filter(Boolean)
                  .join(" | ");
                const summaryNotes = [String(r.notes || "").trim(), checklistFindings]
                  .filter(Boolean)
                  .join(" | ");
                return {
                  id: String(r.id),
                  date: r.inspection_date || "",
                  asset: r.asset_code || "",
                  name: r.asset_name || "",
                  hrs:
                    r.machine_hours != null && Number.isFinite(Number(r.machine_hours))
                      ? Number(r.machine_hours).toFixed(1)
                      : "—",
                  wo: r.work_order_id ? String(r.work_order_id) : "—",
                  inspector: r.inspector_name || "-",
                  notes: compactCell(summaryNotes || "-", 100),
                };
              })
            : [{
                id: "-",
                date: "-",
                asset: "-",
                name: "No inspections found in selected period",
                hrs: "-",
                wo: "-",
                inspector: "-",
                notes: "-",
              }]
        );

        if (withPhotos) {
          sectionTitle(doc, "Inspection Photos");
          if (!rows.length) {
            doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No inspections in selected period.");
          } else {
            for (const r of rows) {
              const photos = photosByInspection.get(Number(r.id)) || [];
              const checklist = parseManagerChecklist(r.checklist_json, r.checklist_detail_json);
              const failedChecklist = checklist.filter((c) => c && c.ok === false);
              const checklistNotes = failedChecklist
                .map((c) => {
                  const label = String(c.label || c.key || "Item").trim();
                  const note = String(c.note || "").trim();
                  return note ? `${label}: ${note}` : label;
                })
                .filter(Boolean);
              const extractChecklistDetailText = (raw) => {
                const text = String(raw || "").trim();
                if (!text) return "";
                try {
                  const parsed = JSON.parse(text);
                  const snippets = [];
                  const walk = (node) => {
                    if (node == null) return;
                    if (Array.isArray(node)) {
                      for (const item of node) walk(item);
                      return;
                    }
                    if (typeof node === "object") {
                      const label = String(node.component || node.label || node.key || "").trim();
                      const note = String(node.note || node.notes || node.description || node.comment || "").trim();
                      if (note) snippets.push(label ? `${label}: ${note}` : note);
                      for (const [k, v] of Object.entries(node)) {
                        if (["component", "label", "key", "note", "notes", "description", "comment"].includes(k)) continue;
                        walk(v);
                      }
                    }
                  };
                  walk(parsed);
                  return snippets.join(" | ");
                } catch {
                  return "";
                }
              };
              const description = String(r.notes || "").trim() || extractChecklistDetailText(r.checklist_detail_json);
              ensurePageSpace(doc, 60);
              doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
              doc.text(
                `Inspection #${r.id} | ${r.inspection_date} | ${r.asset_code}${r.asset_name ? ` - ${r.asset_name}` : ""}`,
                { width: doc.page.width - doc.page.margins.left - doc.page.margins.right }
              );
              doc.moveDown(0.2);
              doc.font("Helvetica").fontSize(9).fillColor("#111111");
              doc.text(`Description: ${compactCell(description || "-", 700)}`, {
                width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
              });
              if (failedChecklist.length) {
                doc.moveDown(0.1);
                doc.text(`Checklist failures: ${failedChecklist.map((c) => String(c.label || c.key || "Item")).join("; ")}`, {
                  width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
                });
              }
              if (checklistNotes.length) {
                doc.moveDown(0.1);
                doc.text(`Failure notes: ${compactCell(checklistNotes.join(" | "), 700)}`, {
                  width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
                });
              }
              doc.moveDown(0.15);
              if (!photos.length) {
                doc.font("Helvetica").fontSize(9).fillColor("#666666").text("No photos attached.");
                doc.moveDown(0.3);
                continue;
              }
              for (const p of photos) {
                const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
                const abs = resolveStorageAbs(rel);
                ensurePageSpace(doc, 220);
                doc.font("Helvetica").fontSize(9).fillColor("#111111");
                doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`);
                doc.moveDown(0.15);
                if (abs && fs.existsSync(abs)) {
                  try {
                    doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 170], align: "left", valign: "top" });
                    doc.y += 176;
                  } catch {
                    doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo exists but could not be rendered.");
                    doc.moveDown(0.3);
                  }
                } else {
                  doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
                  doc.moveDown(0.3);
                }
              }
              doc.moveDown(0.2);
            }
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: withPhotos ? "Manager Inspections Report (With Photos)" : "Manager Inspections Report",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Manager_Inspections_${withPhotos ? "WithPhotos_" : ""}${start}_to_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/damage-report/:id.pdf?download=1
  app.get("/damage-report/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    const download = String(req.query?.download || "").trim() === "1";
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).send({ error: "valid damage report id required" });
    }

    const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
    const drPhotoReportCol = pickExistingColumn("manager_damage_report_photos", ["damage_report_id", "manager_damage_report_id", "report_id"], "damage_report_id");
    const drPhotoPathCol = pickExistingColumn("manager_damage_report_photos", ["file_path", "photo_path", "path", "image_path", "url", "image_data"], "file_path");
    const drPhotoCaptionCol = pickExistingColumn("manager_damage_report_photos", ["caption", "note", "notes", "description"], "caption");
    const drPhotoCreatedCol = pickExistingColumn("manager_damage_report_photos", ["created_at", "uploaded_at", "created_on"], "created_at");

    const report = db.prepare(`
      SELECT
        dr.id,
        dr.report_date,
        dr.${drInspectorCol} AS inspector_name,
        dr.hour_meter,
        dr.damage_location,
        dr.severity,
        dr.damage_description,
        dr.immediate_action,
        dr.out_of_service,
        dr.damage_time,
        dr.responsible_person,
        dr.pending_investigation,
        dr.hse_report_available,
        dr.created_at,
        a.asset_code,
        a.asset_name,
        a.category
      FROM manager_damage_reports dr
      JOIN assets a ON a.id = dr.asset_id
      WHERE dr.id = ?
    `).get(id);
    if (!report) return reply.code(404).send({ error: "damage report not found" });

    const photos = db.prepare(`
      SELECT id, ${drPhotoPathCol} AS file_path, ${drPhotoCaptionCol} AS caption, ${drPhotoCreatedCol} AS created_at
      FROM manager_damage_report_photos
      WHERE ${drPhotoReportCol} = ?
      ORDER BY id ASC
    `).all(id);

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Damage Report");
        kvGrid(doc, [
          { k: "Report #", v: report.id },
          { k: "Date", v: report.report_date || "" },
          { k: "Inspector", v: report.inspector_name || "-" },
          { k: "Asset Code", v: report.asset_code || "" },
          { k: "Asset Name", v: report.asset_name || "" },
          { k: "Hours", v: report.hour_meter == null ? "-" : Number(report.hour_meter || 0).toFixed(1) },
          { k: "Damage Time", v: report.damage_time || "-" },
          { k: "Location", v: report.damage_location || "-" },
          { k: "Responsible Person", v: report.responsible_person || "-" },
          { k: "Severity", v: String(report.severity || "-").toUpperCase() },
          { k: "Out of Service", v: Number(report.out_of_service || 0) ? "YES" : "NO" },
          { k: "Pending Investigation", v: Number(report.pending_investigation || 0) ? "YES" : "NO" },
          { k: "HSE Report Available", v: Number(report.hse_report_available || 0) ? "YES" : "NO" },
          { k: "Category", v: report.category || "" },
          { k: "Created At", v: report.created_at || "" },
        ], 2);

        sectionTitle(doc, "Damage Description");
        doc.font("Helvetica").fontSize(10).fillColor("#111111").text(compactCell(report.damage_description || "-", 2000), {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });

        sectionTitle(doc, "Immediate Action");
        doc.font("Helvetica").fontSize(10).fillColor("#111111").text(compactCell(report.immediate_action || "-", 2000), {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });

        sectionTitle(doc, "Photos");
        if (!photos.length) {
          doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No photos attached.");
          return;
        }
        for (const p of photos) {
          const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
          const abs = resolveStorageAbs(rel);
          ensurePageSpace(doc, 230);
          doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
          doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`, {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });
          doc.moveDown(0.2);
          if (abs && fs.existsSync(abs)) {
            try {
              doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 180], align: "left", valign: "top" });
              doc.y += 186;
            } catch {
              doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo file exists but could not be rendered.");
              doc.moveDown(0.5);
            }
          } else {
            doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
            doc.moveDown(0.5);
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: "Damage Report",
        rightText: `Report #${report.id}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename="AML_Damage_Report_${report.id}.pdf"`)
      .send(pdf);
  });

  // GET /api/reports/damage-reports.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&asset_id=123&with_photos=1&download=1
  app.get("/damage-reports.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const assetId = Number(req.query?.asset_id || 0);
    const withPhotos = String(req.query?.with_photos || "").trim() === "1";
    const download = String(req.query?.download || "").trim() === "1";

    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
    const drPhotoReportCol = pickExistingColumn("manager_damage_report_photos", ["damage_report_id", "manager_damage_report_id", "report_id"], "damage_report_id");
    const drPhotoPathCol = pickExistingColumn("manager_damage_report_photos", ["file_path", "photo_path", "path", "image_path", "url", "image_data"], "file_path");
    const drPhotoCaptionCol = pickExistingColumn("manager_damage_report_photos", ["caption", "note", "notes", "description"], "caption");
    const drPhotoCreatedCol = pickExistingColumn("manager_damage_report_photos", ["created_at", "uploaded_at", "created_on"], "created_at");

    const where = ["dr.report_date >= ?", "dr.report_date <= ?"];
    const params = [start, end];
    if (assetId > 0) {
      where.push("dr.asset_id = ?");
      params.push(assetId);
    }

    const rows = db.prepare(`
      SELECT
        dr.id,
        dr.asset_id,
        dr.report_date,
        dr.${drInspectorCol} AS inspector_name,
        dr.hour_meter,
        dr.damage_location,
        dr.severity,
        dr.damage_description,
        dr.immediate_action,
        dr.out_of_service,
        dr.damage_time,
        dr.responsible_person,
        dr.pending_investigation,
        dr.hse_report_available,
        a.asset_code,
        a.asset_name
      FROM manager_damage_reports dr
      JOIN assets a ON a.id = dr.asset_id
      WHERE ${where.join(" AND ")}
      ORDER BY dr.report_date DESC, dr.id DESC
      LIMIT 1000
    `).all(...params);

    const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
    const photosByReport = new Map();
    if (ids.length) {
      const marks = ids.map(() => "?").join(",");
      const photos = db.prepare(`
        SELECT ${drPhotoReportCol} AS damage_report_id, id, ${drPhotoPathCol} AS file_path, ${drPhotoCaptionCol} AS caption, ${drPhotoCreatedCol} AS created_at
        FROM manager_damage_report_photos
        WHERE ${drPhotoReportCol} IN (${marks})
        ORDER BY ${drPhotoReportCol} ASC, id ASC
      `).all(...ids);
      for (const p of photos) {
        const key = Number(p.damage_report_id);
        if (!photosByReport.has(key)) photosByReport.set(key, []);
        photosByReport.get(key).push(p);
      }
    }

    const summary = {
      count: rows.length,
      assets: new Set(rows.map((r) => Number(r.asset_id))).size,
      out_of_service: rows.filter((r) => Number(r.out_of_service || 0) === 1).length,
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Damage Reports Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Asset Filter", v: assetId > 0 ? String(assetId) : "All assets" },
          { k: "Reports", v: String(summary.count) },
          { k: "Assets Covered", v: String(summary.assets) },
          { k: "Out of Service", v: String(summary.out_of_service) },
        ], 2);

        sectionTitle(doc, "Damage Entries");
        table(
          doc,
          [
            { key: "id", label: "ID", width: 0.08, align: "right" },
            { key: "date", label: "Date", width: 0.12 },
            { key: "asset", label: "Asset", width: 0.13 },
            { key: "inspector", label: "Inspector", width: 0.11 },
            { key: "time", label: "Time", width: 0.06, align: "center" },
            { key: "location", label: "Location", width: 0.1 },
            { key: "resp", label: "Responsible", width: 0.1 },
            { key: "severity", label: "Severity", width: 0.08 },
            { key: "hours", label: "Hours", width: 0.08, align: "right" },
            { key: "out", label: "OOS", width: 0.06, align: "center" },
            { key: "inv", label: "Inv", width: 0.05, align: "center" },
            { key: "hse", label: "HSE", width: 0.05, align: "center" },
            { key: "desc", label: "Description", width: 0.08 },
          ],
          rows.length
            ? rows.map((r) => ({
                id: String(r.id),
                date: r.report_date || "",
                asset: r.asset_code || "",
                inspector: r.inspector_name || "-",
                time: r.damage_time || "-",
                location: compactCell(r.damage_location || "", 40),
                resp: compactCell(r.responsible_person || "", 20),
                severity: String(r.severity || "-").toUpperCase(),
                hours: r.hour_meter == null ? "-" : Number(r.hour_meter || 0).toFixed(1),
                out: Number(r.out_of_service || 0) ? "YES" : "NO",
                inv: Number(r.pending_investigation || 0) ? "Y" : "N",
                hse: Number(r.hse_report_available || 0) ? "Y" : "N",
                desc: compactCell(r.damage_description || "", 60),
              }))
            : [{
                id: "-",
                date: "-",
                asset: "-",
                inspector: "-",
                time: "-",
                location: "-",
                resp: "-",
                severity: "-",
                hours: "-",
                out: "-",
                inv: "-",
                hse: "-",
                desc: "No damage reports found in selected period",
              }]
        );

        if (withPhotos) {
          sectionTitle(doc, "Damage Photos");
          if (!rows.length) {
            doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No damage reports in selected period.");
          } else {
            for (const r of rows) {
              const photos = photosByReport.get(Number(r.id)) || [];
              ensurePageSpace(doc, 60);
              doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
              doc.text(
                `Report #${r.id} | ${r.report_date} | ${r.asset_code}${r.asset_name ? ` - ${r.asset_name}` : ""} | Severity: ${String(r.severity || "-").toUpperCase()}`,
                { width: doc.page.width - doc.page.margins.left - doc.page.margins.right }
              );
              doc.moveDown(0.2);
              if (!photos.length) {
                doc.font("Helvetica").fontSize(9).fillColor("#666666").text("No photos attached.");
                doc.moveDown(0.3);
                continue;
              }
              for (const p of photos) {
                const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
                const abs = resolveStorageAbs(rel);
                ensurePageSpace(doc, 220);
                doc.font("Helvetica").fontSize(9).fillColor("#111111");
                doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`);
                doc.moveDown(0.15);
                if (abs && fs.existsSync(abs)) {
                  try {
                    doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 170], align: "left", valign: "top" });
                    doc.y += 176;
                  } catch {
                    doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo exists but could not be rendered.");
                    doc.moveDown(0.3);
                  }
                } else {
                  doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
                  doc.moveDown(0.3);
                }
              }
              doc.moveDown(0.2);
            }
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: withPhotos ? "Damage Reports (With Photos)" : "Damage Reports",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename="AML_Damage_Reports_${withPhotos ? "WithPhotos_" : ""}${start}_to_${end}.pdf"`)
      .send(pdf);
  });

  // GET /api/reports/damage-reports.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD&asset_id=123
  app.get("/damage-reports.xlsx", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const assetId = Number(req.query?.asset_id || 0);
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
    const drPhotoReportCol = pickExistingColumn("manager_damage_report_photos", ["damage_report_id", "manager_damage_report_id", "report_id"], "damage_report_id");

    const where = ["dr.report_date >= ?", "dr.report_date <= ?"];
    const params = [start, end];
    if (assetId > 0) {
      where.push("dr.asset_id = ?");
      params.push(assetId);
    }

    const rows = db.prepare(`
      SELECT
        dr.id,
        dr.asset_id,
        dr.report_date,
        dr.${drInspectorCol} AS inspector_name,
        dr.hour_meter,
        dr.damage_location,
        dr.severity,
        dr.damage_description,
        dr.immediate_action,
        dr.out_of_service,
        dr.damage_time,
        dr.responsible_person,
        dr.pending_investigation,
        dr.hse_report_available,
        a.asset_code,
        a.asset_name
      FROM manager_damage_reports dr
      JOIN assets a ON a.id = dr.asset_id
      WHERE ${where.join(" AND ")}
      ORDER BY dr.report_date DESC, dr.id DESC
      LIMIT 5000
    `).all(...params);

    const photoCounts = new Map();
    if (rows.length) {
      const ids = rows.map((r) => Number(r.id || 0)).filter((n) => n > 0);
      if (ids.length) {
        const marks = ids.map(() => "?").join(",");
        const grouped = db.prepare(`
          SELECT ${drPhotoReportCol} AS damage_report_id, COUNT(*) AS photo_count
          FROM manager_damage_report_photos
          WHERE ${drPhotoReportCol} IN (${marks})
          GROUP BY ${drPhotoReportCol}
        `).all(...ids);
        grouped.forEach((g) => photoCounts.set(Number(g.damage_report_id || 0), Number(g.photo_count || 0)));
      }
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();
    const ws = wb.addWorksheet("Damage Reports");
    ws.columns = [
      { header: "Report ID", key: "id", width: 12 },
      { header: "Report Date", key: "report_date", width: 14 },
      { header: "Damage Time", key: "damage_time", width: 12 },
      { header: "Asset Code", key: "asset_code", width: 14 },
      { header: "Asset Name", key: "asset_name", width: 28 },
      { header: "Inspector", key: "inspector_name", width: 20 },
      { header: "Hour Meter", key: "hour_meter", width: 12 },
      { header: "Severity", key: "severity", width: 12 },
      { header: "Damage Location", key: "damage_location", width: 24 },
      { header: "Responsible Person", key: "responsible_person", width: 24 },
      { header: "Damage Description", key: "damage_description", width: 44 },
      { header: "Immediate Action", key: "immediate_action", width: 36 },
      { header: "Out Of Service", key: "out_of_service", width: 14 },
      { header: "Pending Investigation", key: "pending_investigation", width: 18 },
      { header: "HSE Report Available", key: "hse_report_available", width: 18 },
      { header: "Photo Count", key: "photo_count", width: 12 },
    ];
    ws.getRow(1).font = { bold: true };
    rows.forEach((r) => {
      ws.addRow({
        id: Number(r.id || 0),
        report_date: r.report_date || "",
        damage_time: r.damage_time || "",
        asset_code: r.asset_code || "",
        asset_name: r.asset_name || "",
        inspector_name: r.inspector_name || "",
        hour_meter: r.hour_meter == null ? "" : Number(r.hour_meter || 0),
        severity: String(r.severity || "").toUpperCase(),
        damage_location: r.damage_location || "",
        responsible_person: r.responsible_person || "",
        damage_description: r.damage_description || "",
        immediate_action: r.immediate_action || "",
        out_of_service: Number(r.out_of_service || 0) ? "YES" : "NO",
        pending_investigation: Number(r.pending_investigation || 0) ? "YES" : "NO",
        hse_report_available: Number(r.hse_report_available || 0) ? "YES" : "NO",
        photo_count: photoCounts.get(Number(r.id || 0)) || 0,
      });
    });
    ws.views = [{ state: "frozen", ySplit: 1 }];

    const summary = wb.addWorksheet("Summary");
    summary.columns = [
      { header: "Metric", key: "metric", width: 26 },
      { header: "Value", key: "value", width: 24 },
    ];
    summary.getRow(1).font = { bold: true };
    summary.addRows([
      { metric: "Start Date", value: start },
      { metric: "End Date", value: end },
      { metric: "Asset Filter", value: assetId > 0 ? String(assetId) : "All assets" },
      { metric: "Total Reports", value: rows.length },
      { metric: "Assets Covered", value: new Set(rows.map((r) => Number(r.asset_id || 0))).size },
      { metric: "Out Of Service", value: rows.filter((r) => Number(r.out_of_service || 0) === 1).length },
      { metric: "Pending Investigation", value: rows.filter((r) => Number(r.pending_investigation || 0) === 1).length },
      { metric: "HSE Report Available", value: rows.filter((r) => Number(r.hse_report_available || 0) === 1).length },
    ]);

    const buffer = await wb.xlsx.writeBuffer();
    return reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="AML_Damage_Reports_${start}_to_${end}.xlsx"`)
      .send(Buffer.from(buffer));
  });

  // =========================
  // LEGAL COMPLIANCE PDF
  // =========================
  // GET /api/reports/legal-compliance.pdf?days=90&department=&status=approved&download=1
  app.get("/legal-compliance.pdf", async (req, reply) => {
    const daysRaw = Number(req.query?.days ?? 90);
    const days = Number.isFinite(daysRaw) ? Math.min(3650, Math.max(1, Math.trunc(daysRaw))) : 90;
    const department = String(req.query?.department || "").trim();
    const status = String(req.query?.status || "approved").trim().toLowerCase();
    const download = String(req.query?.download || "").trim() === "1";

    const where = ["ld.expiry_date IS NOT NULL", "TRIM(ld.expiry_date) <> ''"];
    const params = [];
    if (department) {
      where.push("ld.department = ?");
      params.push(department);
    }
    if (status && status !== "all") {
      where.push("ld.status = ?");
      params.push(status);
    }

    const rows = db.prepare(`
      SELECT
        ld.id,
        ld.department,
        ld.title,
        ld.doc_type,
        ld.version,
        ld.owner,
        ld.status,
        ld.active,
        ld.expiry_date,
        CAST(julianday(ld.expiry_date) - julianday(DATE('now')) AS INTEGER) AS days_to_expiry
      FROM legal_documents ld
      WHERE ${where.join(" AND ")}
      ORDER BY ld.expiry_date ASC, ld.id DESC
      LIMIT 1000
    `).all(...params).map((r) => ({
      ...r,
      id: Number(r.id),
      active: Number(r.active),
      days_to_expiry: Number(r.days_to_expiry),
    }));

    const dueRows = rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= days);
    const expiredRows = rows.filter((r) => Number(r.days_to_expiry) < 0);
    const summary = {
      total_with_expiry: rows.length,
      expired: expiredRows.length,
      due_30: rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= 30).length,
      due_60: rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= 60).length,
      due_90: rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= 90).length,
      due_window: dueRows.length,
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Legal Compliance Summary");
        kvGrid(doc, [
          { k: "Window (days)", v: String(days) },
          { k: "Department", v: department || "All" },
          { k: "Status", v: status || "approved" },
          { k: "Total with Expiry", v: fmtNum(summary.total_with_expiry, 0) },
          { k: "Expired", v: fmtNum(summary.expired, 0) },
          { k: "Due in 30 days", v: fmtNum(summary.due_30, 0) },
          { k: "Due in 60 days", v: fmtNum(summary.due_60, 0) },
          { k: "Due in 90 days", v: fmtNum(summary.due_90, 0) },
          { k: `Due in ${days} days`, v: fmtNum(summary.due_window, 0) },
        ], 2);

        sectionTitle(doc, "Expired Documents");
        table(
          doc,
          [
            { key: "id", label: "ID", width: 0.08, align: "right" },
            { key: "department", label: "Department", width: 0.16 },
            { key: "title", label: "Title", width: 0.30 },
            { key: "status", label: "Status", width: 0.12 },
            { key: "expiry", label: "Expiry", width: 0.14 },
            { key: "days", label: "Days", width: 0.10, align: "right" },
            { key: "owner", label: "Owner", width: 0.10 },
          ],
          expiredRows.length
            ? expiredRows.map((r) => ({
                id: String(r.id),
                department: r.department || "-",
                title: compactCell(r.title || "-", 100),
                status: r.status || "-",
                expiry: r.expiry_date || "-",
                days: fmtNum(r.days_to_expiry, 0),
                owner: compactCell(r.owner || "-", 40),
              }))
            : [{ id: "-", department: "-", title: "No expired documents", status: "-", expiry: "-", days: "-", owner: "-" }]
        );

        sectionTitle(doc, `Due In ${days} Days`);
        table(
          doc,
          [
            { key: "id", label: "ID", width: 0.08, align: "right" },
            { key: "department", label: "Department", width: 0.16 },
            { key: "title", label: "Title", width: 0.30 },
            { key: "status", label: "Status", width: 0.12 },
            { key: "expiry", label: "Expiry", width: 0.14 },
            { key: "days", label: "Days", width: 0.10, align: "right" },
            { key: "owner", label: "Owner", width: 0.10 },
          ],
          dueRows.length
            ? dueRows.map((r) => ({
                id: String(r.id),
                department: r.department || "-",
                title: compactCell(r.title || "-", 100),
                status: r.status || "-",
                expiry: r.expiry_date || "-",
                days: fmtNum(r.days_to_expiry, 0),
                owner: compactCell(r.owner || "-", 40),
              }))
            : [{ id: "-", department: "-", title: "No due documents in selected window", status: "-", expiry: "-", days: "-", owner: "-" }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Legal Compliance Report",
        rightText: `${department || "All"} | ${days}d`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Legal_Compliance_${todayYmd()}.pdf"`
      )
      .send(pdf);
  });
}
