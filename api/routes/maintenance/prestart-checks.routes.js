// IRONLOG/api/routes/maintenance/prestart-checks.routes.js — LDV and machine pre-start checks, checklist hub and damage reports.
// Registered by routes/maintenance.routes.js; shared helpers arrive through ctx.
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { checklistToJsonObject, getMachinePrestartTemplate, listMachinePrestartProfiles, machinePrestartCheckMode, normalizeMachinePrestartChecklist, resolveMachinePrestartProfile } from "../../utils/machinePrestartTemplates.js";
import { db } from "../../db/client.js";
import { isDate } from "../../utils/request.js";
import { faultMessage, notesWithFaults, prestartFaultList, syncPrestartFaultWorkOrder, unansweredChecks } from "../../utils/prestartFaults.js";
import { listDailyPrestarts, prestartDeductionForProductionFleet } from "../../utils/prestartDaily.js";
import { normalizeUploadedPhoto } from "../../utils/imagePdf.js";

export default function registerPrestartChecksRoutes(app, ctx) {
  const {
    applyLdvKmToAllChecksForDate,
    checklistRowStatus,
    damageReportsDir,
    dmgPhotoCaptionCol,
    dmgPhotoCreatedCol,
    dmgPhotoPathCol,
    dmgPhotoReportCol,
    getLatestLdvOdometerKm,
    getLatestMachineSmuHours,
    getLdvPrestartCheckRow,
    getMachinePrestartCheckRow,
    getPriorLdvPrestartOdometerKm,
    getSanitizedLdvBaselineKm,
    hasColumn,
    isLdvOdometerOutlier,
    isLdvPrestartAssetCode,
    normalizeLdvPrestartChecklist,
    pickExistingColumn,
    purgeLdvPoisonedKmBaselines,
    requireMaintenanceRoles,
    resolveLdvCorrectionOpeningKm,
    syncLdvPrestartToDailyHours,
    syncMachinePrestartToDailyHours,
    upsertMachineDailyHoursCorrection,
    vehicleLdvcDir,
  } = ctx;

  app.get("/vehicle-ldv-checks", async (req, reply) => {
    try {
      const assetId = Number(req.query?.asset_id || 0);
      const checkIdFilter = Number(req.query?.check_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const checkModeFilter = String(req.query?.check_mode || "").trim().toLowerCase();
      const assetCodeFilter = String(req.query?.asset_code || "").trim().toUpperCase();
      const params = [];
      const where = [];
      if (checkIdFilter > 0) {
        where.push("v.id = ?");
        params.push(checkIdFilter);
      }
      if (assetId > 0) {
        where.push("v.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("v.check_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("v.check_date <= ?");
        params.push(end);
      }
      if (assetCodeFilter) {
        where.push("UPPER(a.asset_code) = ?");
        params.push(assetCodeFilter);
      }
      if (checkModeFilter === "ldv") {
        where.push("COALESCE(v.check_mode, 'ldv_general') = 'prestart'");
      } else if (checkModeFilter === "machine") {
        where.push("COALESCE(v.check_mode, '') LIKE 'machine_prestart_%'");
      } else if (checkModeFilter === "vehicle") {
        where.push("COALESCE(v.check_mode, 'ldv_general') = 'ldv_general'");
      }

      const rows = db.prepare(`
        SELECT
          v.id,
          v.asset_id,
          v.check_date,
          v.vehicle_registration,
          v.odometer_km,
          v.smu_hours,
          v.inspector_name,
          v.notes,
          v.check_mode,
          v.created_at,
          a.asset_code,
          a.asset_name
        FROM vehicle_ldv_checks v
        JOIN assets a ON a.id = v.asset_id
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY v.check_date DESC, v.id DESC
        LIMIT 500
      `).all(...params);

      const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
      let photosByCheck = new Map();
      if (ids.length) {
        const marks = ids.map(() => "?").join(",");
        const photos = db.prepare(`
          SELECT id, check_id, file_path, caption, markers_json, created_at
          FROM vehicle_ldv_check_photos
          WHERE check_id IN (${marks})
          ORDER BY id ASC
        `).all(...ids);
        photosByCheck = photos.reduce((m, p) => {
          const k = Number(p.check_id);
          if (!m.has(k)) m.set(k, []);
          let markers = [];
          try {
            markers = p.markers_json ? JSON.parse(p.markers_json) : [];
          } catch {
            markers = [];
          }
          m.get(k).push({
            ...p,
            markers: Array.isArray(markers) ? markers : [],
          });
          return m;
        }, new Map());
      }

      return reply.send({
        ok: true,
        rows: rows.map((r) => {
          const photos = photosByCheck.get(Number(r.id)) || [];
          return {
            ...r,
            photos,
            photo_count: photos.length,
          };
        }),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // GET /api/maintenance/checklist-hub?date=YYYY-MM-DD
  app.get("/checklist-hub", async (req, reply) => {
    try {
      const check_date = String(req.query?.date || "").trim() || new Date().toISOString().slice(0, 10);
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "date must be YYYY-MM-DD" });

      const assets = db.prepare(`
        SELECT id, asset_code, asset_name, category
        FROM assets
        WHERE archived = 0
        ORDER BY asset_code ASC
      `).all();

      const ldvAssets = [];
      const machineGroupMap = new Map();
      for (const profile of listMachinePrestartProfiles()) {
        machineGroupMap.set(profile.id, {
          profile_id: profile.id,
          title: profile.title,
          assets: [],
        });
      }

      for (const a of assets) {
        const assetId = Number(a.id);
        const code = String(a.asset_code || "");
        if (isLdvPrestartAssetCode(code)) {
          const row = getLdvPrestartCheckRow(assetId, check_date);
          const st = row
            ? { ...checklistRowStatus(row.checklist_json, normalizeLdvPrestartChecklist), check_id: Number(row.id) }
            : { status: "pending", check_id: null };
          ldvAssets.push({
            asset_id: assetId,
            asset_code: code,
            asset_name: String(a.asset_name || ""),
            category: String(a.category || ""),
            kind: "ldv",
            ...st,
          });
          continue;
        }
        const profileId = resolveMachinePrestartProfile(a.category, a.asset_name, code);
        if (!profileId) continue;
        const mode = machinePrestartCheckMode(profileId);
        if (!mode) continue;
        const row = getMachinePrestartCheckRow(assetId, check_date, mode);
        const st = row
          ? {
              ...checklistRowStatus(row.checklist_json, (parsed) =>
                normalizeMachinePrestartChecklist(profileId, parsed)
              ),
              check_id: Number(row.id),
            }
          : { status: "pending", check_id: null };
        const group = machineGroupMap.get(profileId);
        if (group) {
          group.assets.push({
            asset_id: assetId,
            asset_code: code,
            asset_name: String(a.asset_name || ""),
            category: String(a.category || ""),
            kind: "machine",
            profile_id: profileId,
            ...st,
          });
        }
      }

      const machine_groups = [...machineGroupMap.values()].filter((g) => g.assets.length > 0);
      const summary = {
        ldv_total: ldvAssets.length,
        ldv_compliant: ldvAssets.filter((x) => x.status === "compliant").length,
        machine_total: machine_groups.reduce((s, g) => s + g.assets.length, 0),
        machine_compliant: machine_groups.reduce(
          (s, g) => s + g.assets.filter((x) => x.status === "compliant").length,
          0
        ),
      };

      const commentRows = db.prepare(`
        SELECT v.id AS check_id, v.check_date, v.inspector_name, v.notes, v.check_mode, v.updated_at,
               a.asset_code, a.asset_name
        FROM vehicle_ldv_checks v
        JOIN assets a ON a.id = v.asset_id
        WHERE v.check_date = ?
          AND TRIM(COALESCE(v.notes, '')) != ''
        ORDER BY datetime(COALESCE(v.updated_at, v.check_date)) DESC, v.id DESC
      `).all(check_date);
      const comments = commentRows.map((r) => {
        const mode = String(r.check_mode || "").toLowerCase();
        const kind = mode.includes("machine_prestart") ? "machine" : "ldv";
        return {
          kind,
          check_id: Number(r.check_id),
          asset_code: String(r.asset_code || ""),
          asset_name: String(r.asset_name || ""),
          inspector_name: String(r.inspector_name || "").trim() || null,
          notes: String(r.notes || "").trim(),
          updated_at: r.updated_at || null,
        };
      });

      return reply.send({
        ok: true,
        check_date,
        summary,
        ldv: ldvAssets,
        machine_groups,
        comments,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.get("/vehicle-ldv-checks/prestart-context", async (req, reply) => {
    try {
      const asset_code = String(req.query?.asset_code || "").trim().toUpperCase();
      const check_date = String(req.query?.check_date || "").trim() || new Date().toISOString().slice(0, 10);
      if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "check_date must be YYYY-MM-DD" });

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name, category
        FROM assets
        WHERE UPPER(asset_code) = UPPER(?)
        LIMIT 1
      `).get(asset_code);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const existing = getLdvPrestartCheckRow(Number(asset.id), check_date);

      const raw_previous_odometer_km = getLatestLdvOdometerKm(Number(asset.id), check_date, {
        excludeCheckId: existing?.id ? Number(existing.id) : 0,
      });
      const previous_odometer_km = getSanitizedLdvBaselineKm(Number(asset.id), check_date, {
        excludeCheckId: existing?.id ? Number(existing.id) : 0,
      });
      const baseline_poisoned =
        raw_previous_odometer_km != null &&
        previous_odometer_km != null &&
        Math.abs(Number(raw_previous_odometer_km) - Number(previous_odometer_km)) > 1;

      let checklist = normalizeLdvPrestartChecklist({});
      if (existing?.checklist_json) {
        try {
          const parsed = JSON.parse(String(existing.checklist_json || "{}"));
          if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
            checklist = normalizeLdvPrestartChecklist(parsed);
          }
        } catch {}
      }

      return reply.send({
        ok: true,
        asset: {
          id: Number(asset.id),
          asset_code: String(asset.asset_code || ""),
          asset_name: String(asset.asset_name || ""),
          category: String(asset.category || ""),
        },
        check_date,
        previous_odometer_km,
        raw_previous_odometer_km,
        baseline_poisoned,
        previous_odometer_source:
          previous_odometer_km != null ? "sanitized_prior_prestart_or_daily" : null,
        existing_prestart: existing
          ? {
              id: Number(existing.id),
              odometer_km: existing.odometer_km == null ? null : Number(existing.odometer_km),
              inspector_name: existing.inspector_name || "",
              notes: existing.notes || "",
              checklist,
            }
          : null,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/vehicle-ldv-checks/prestart", async (req, reply) => {
    try {
      const asset_code = String(req.body?.asset_code || "").trim().toUpperCase();
      const check_date = String(req.body?.check_date || "").trim() || new Date().toISOString().slice(0, 10);
      const odometer_km_raw = req.body?.odometer_km;
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      const checklistObj = req.body?.checklist || {};
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "check_date must be YYYY-MM-DD" });
      if (odometer_km_raw == null || String(odometer_km_raw).trim() === "") {
        return reply.code(400).send({ ok: false, error: "odometer_km is required" });
      }
      const odometer_km = Number(odometer_km_raw);
      if (!Number.isFinite(odometer_km) || odometer_km < 0) {
        return reply.code(400).send({ ok: false, error: "odometer_km must be a valid number >= 0" });
      }

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name
        FROM assets
        WHERE UPPER(asset_code) = UPPER(?)
        LIMIT 1
      `).get(asset_code);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const existing = getLdvPrestartCheckRow(Number(asset.id), check_date);

      const previousOdometer =
        getSanitizedLdvBaselineKm(Number(asset.id), check_date, {
          excludeCheckId: existing?.id ? Number(existing.id) : 0,
        }) ??
        getLatestLdvOdometerKm(Number(asset.id), check_date, {
          excludeCheckId: existing?.id ? Number(existing.id) : 0,
        });
      const kmReviewNeeded =
        (previousOdometer != null && odometer_km < previousOdometer) ||
        isLdvOdometerOutlier(odometer_km, previousOdometer);

      const checklist = normalizeLdvPrestartChecklist(checklistObj);
      const unanswered = unansweredChecks(checklist, checklistObj);
      if (unanswered.length) {
        return reply.code(400).send({
          ok: false,
          error: `Mark every check OK or Fault before submitting (${unanswered.join(", ")}).`,
        });
      }
      const faults = prestartFaultList(checklist, req.body?.faults);

      const checklistJson = JSON.stringify(
        checklist.reduce((acc, c) => {
          acc[c.key] = Boolean(c.ok);
          return acc;
        }, {})
      );
      const reviewNote = kmReviewNeeded ? "KM flagged for supervisor review" : null;
      const mergedNotes = [notesWithFaults(notes, faults), reviewNote].filter(Boolean).join(" | ") || null;

      let checkId = 0;
      if (existing?.id) {
        checkId = Number(existing.id);
        applyLdvKmToAllChecksForDate(
          Number(asset.id),
          check_date,
          odometer_km,
          inspector_name,
          mergedNotes,
          checklistJson
        );
      } else {
        const ins = db.prepare(`
          INSERT INTO vehicle_ldv_checks (
            asset_id, uuid, site_code, check_date, vehicle_registration, odometer_km, inspector_name, notes, check_mode, checklist_json, updated_at
          )
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'prestart', ?, datetime('now'))
        `).run(
          Number(asset.id),
          crypto.randomUUID(),
          site_code,
          check_date,
          String(asset.asset_code || ""),
          odometer_km,
          inspector_name,
          mergedNotes,
          checklistJson
        );
        checkId = Number(ins.lastInsertRowid);
      }

      const dailySync = syncLdvPrestartToDailyHours(
        Number(asset.id),
        check_date,
        odometer_km,
        inspector_name,
        previousOdometer,
        { unusual_km: kmReviewNeeded }
      );

      const faultWo = syncPrestartFaultWorkOrder(db, {
        assetId: Number(asset.id), checkId, siteCode: site_code, checkDate: check_date, operator: inspector_name, faults,
      });

      return reply.send({
        ok: true,
        id: checkId,
        faults: faults.length,
        fault_work_order_id: faultWo && !faultWo.closed ? faultWo.work_order_id : null,
        asset_code: String(asset.asset_code || ""),
        check_date,
        odometer_km: Number(odometer_km.toFixed(1)),
        previous_odometer_km: previousOdometer == null ? null : Number(previousOdometer.toFixed(1)),
        km_review_needed: kmReviewNeeded,
        daily_input_sync: dailySync || { synced: false },
        message: faultMessage(faults, faultWo) || (kmReviewNeeded
          ? "Pre-start saved. KM looks unusual — daily input not updated until a supervisor reviews."
          : "Pre-start captured. KM reading saved to IRONLOG."),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // POST /api/maintenance/vehicle-ldv-checks/prestart-correction — fix wrong KM on LDV + daily hours (supervisor)
  app.post("/vehicle-ldv-checks/prestart-correction", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;

      const asset_code = String(req.body?.asset_code || "").trim().toUpperCase();
      const work_date = String(req.body?.work_date || req.body?.check_date || "").trim()
        || new Date().toISOString().slice(0, 10);
      const closing_km = Number(req.body?.closing_km ?? req.body?.correct_odometer_km ?? req.body?.odometer_km);
      const opening_km_raw = req.body?.opening_km;
      const inspector_name = String(req.body?.inspector_name || "Supervisor correction").trim() || "Supervisor correction";
      const correction_note = String(req.body?.notes || "Supervisor KM correction").trim() || "Supervisor KM correction";
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      if (!isLdvPrestartAssetCode(asset_code)) {
        return reply.code(400).send({ ok: false, error: "Asset is not an LDV pre-start code (V01–V15)" });
      }
      if (!isDate(work_date)) return reply.code(400).send({ ok: false, error: "work_date must be YYYY-MM-DD" });
      if (!Number.isFinite(closing_km) || closing_km < 0) {
        return reply.code(400).send({ ok: false, error: "closing_km must be a valid number >= 0" });
      }

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name FROM assets WHERE UPPER(asset_code) = UPPER(?) LIMIT 1
      `).get(asset_code);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const assetId = Number(asset.id);
      const existing = getLdvPrestartCheckRow(assetId, work_date);

      let checkId = Number(existing?.id || 0);
      const opening_km = resolveLdvCorrectionOpeningKm(
        assetId,
        work_date,
        closing_km,
        checkId,
        opening_km_raw
      );
      if (!Number.isFinite(opening_km) || opening_km < 0) {
        return reply.code(400).send({ ok: false, error: "opening_km could not be resolved — pass opening_km explicitly" });
      }
      if (closing_km < opening_km) {
        return reply.code(400).send({
          ok: false,
          error: `Closing KM (${closing_km}) cannot be less than opening KM (${opening_km}). Enter opening KM manually if needed.`,
          opening_km,
        });
      }

      const previousOdometer = getLatestLdvOdometerKm(assetId, work_date, {
        excludeCheckId: checkId,
      });

      const checklistJsonOut = (() => {
        if (existing?.checklist_json && String(existing.checklist_json).trim()) {
          return String(existing.checklist_json);
        }
        return JSON.stringify(
          normalizeLdvPrestartChecklist({}).reduce((acc, c) => {
            acc[c.key] = true;
            return acc;
          }, {})
        );
      })();

      const purged = purgeLdvPoisonedKmBaselines(assetId, work_date, closing_km);

      const hasAnyCheckRow = Boolean(
        db.prepare(`SELECT id FROM vehicle_ldv_checks WHERE asset_id = ? AND check_date = ? LIMIT 1`).get(
          assetId,
          work_date
        )
      );

      let canonical;
      if (hasAnyCheckRow) {
        canonical = applyLdvKmToAllChecksForDate(
          assetId,
          work_date,
          closing_km,
          inspector_name,
          correction_note,
          checklistJsonOut
        );
      } else {
        const ins = db.prepare(`
          INSERT INTO vehicle_ldv_checks (
            asset_id, uuid, site_code, check_date, vehicle_registration, odometer_km,
            inspector_name, notes, check_mode, checklist_json, updated_at
          )
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'prestart', ?, datetime('now'))
        `).run(
          assetId,
          crypto.randomUUID(),
          site_code,
          work_date,
          String(asset.asset_code || ""),
          closing_km,
          inspector_name,
          correction_note,
          checklistJsonOut
        );
        canonical = getLdvPrestartCheckRow(assetId, work_date) || { id: Number(ins.lastInsertRowid) };
      }
      checkId = Number(canonical?.id || checkId || 0);

      const run_km = Math.max(0, closing_km - opening_km);
      const hasInputUnit = hasColumn("daily_hours", "input_unit");
      const dailyNote = `Supervisor KM correction | ${correction_note}`;
      if (hasInputUnit) {
        db.prepare(`
          INSERT INTO daily_hours (
            asset_id, work_date, scheduled_hours, opening_hours, closing_hours,
            hours_run, input_unit, is_used, operator, notes
          )
          VALUES (?, ?, 0, ?, ?, ?, 'km', 1, ?, ?)
          ON CONFLICT(asset_id, work_date) DO UPDATE SET
            opening_hours = excluded.opening_hours,
            closing_hours = excluded.closing_hours,
            hours_run = excluded.hours_run,
            input_unit = 'km',
            is_used = 1,
            operator = excluded.operator,
            notes = excluded.notes
        `).run(assetId, work_date, opening_km, closing_km, run_km, inspector_name, dailyNote);
      } else {
        db.prepare(`
          INSERT INTO daily_hours (
            asset_id, work_date, scheduled_hours, opening_hours, closing_hours,
            hours_run, is_used, operator, notes
          )
          VALUES (?, ?, 0, ?, ?, ?, 1, ?, ?)
          ON CONFLICT(asset_id, work_date) DO UPDATE SET
            opening_hours = excluded.opening_hours,
            closing_hours = excluded.closing_hours,
            hours_run = excluded.hours_run,
            is_used = 1,
            operator = excluded.operator,
            notes = excluded.notes
        `).run(assetId, work_date, opening_km, closing_km, run_km, inspector_name, dailyNote);
      }

      if (hasColumn("daily_hours", "input_unit")) {
        db.prepare(`
          INSERT INTO asset_input_units (asset_id, input_unit, updated_at)
          VALUES (?, 'km', datetime('now'))
          ON CONFLICT(asset_id) DO UPDATE SET input_unit = 'km', updated_at = datetime('now')
        `).run(assetId);
      }

      const priorBaseline =
        getSanitizedLdvBaselineKm(assetId, work_date, { excludeCheckId: checkId }) ??
        getPriorLdvPrestartOdometerKm(assetId, work_date) ??
        (previousOdometer != null ? previousOdometer : null);

      return reply.send({
        ok: true,
        asset_code,
        work_date,
        check_id: checkId,
        opening_km: Number(opening_km.toFixed(1)),
        closing_km: Number(closing_km.toFixed(1)),
        run_km: Number(run_km.toFixed(1)),
        previous_odometer_km: priorBaseline == null ? null : Number(priorBaseline.toFixed(1)),
        purged_baselines: purged,
        message: `KM corrected for ${asset_code} on ${work_date}.`,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // Machine-group prestart (excavator, dozer, etc.) — same storage as LDV checks, different check_mode; no daily_hours sync.

  // GET /api/maintenance/prestart/daily-summary?date=YYYY-MM-DD
  app.get("/prestart/daily-summary", async (req, reply) => {
    try {
      const date = String(req.query?.date || "").trim() || new Date().toISOString().slice(0, 10);
      if (!isDate(date)) return reply.code(400).send({ ok: false, error: "date must be YYYY-MM-DD" });
      const summary = listDailyPrestarts(db, date);
      const production = prestartDeductionForProductionFleet(db, date);
      return reply.send({
        ok: true,
        date,
        ...summary,
        production_deduction: production,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.get("/machine-prestart/context", async (req, reply) => {
    try {
      const asset_code = String(req.query?.asset_code || "").trim().toUpperCase();
      const check_date = String(req.query?.check_date || "").trim() || new Date().toISOString().slice(0, 10);
      if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "check_date must be YYYY-MM-DD" });

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name, category
        FROM assets
        WHERE UPPER(asset_code) = UPPER(?)
        LIMIT 1
      `).get(asset_code);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const profileId = resolveMachinePrestartProfile(asset.category, asset.asset_name, asset.asset_code);
      if (!profileId) {
        return reply.code(404).send({
          ok: false,
          error: "No machine pre-start template for this asset. Set category or name (e.g. Excavator, Haul truck).",
        });
      }
      const template = getMachinePrestartTemplate(profileId);
      const mode = machinePrestartCheckMode(profileId);
      if (!template || !mode) return reply.code(500).send({ ok: false, error: "Template resolution failed" });

      const existing = getMachinePrestartCheckRow(Number(asset.id), check_date, mode);

      const previous_smu_hours = getLatestMachineSmuHours(Number(asset.id), check_date, {
        excludeCheckId: existing?.id ? Number(existing.id) : 0,
      });

      let checklist = normalizeMachinePrestartChecklist(profileId, {});
      if (existing?.checklist_json) {
        try {
          const parsed = JSON.parse(String(existing.checklist_json || "{}"));
          if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
            checklist = normalizeMachinePrestartChecklist(profileId, parsed);
          }
        } catch {}
      }

      return reply.send({
        ok: true,
        profile_id: profileId,
        check_mode: mode,
        template,
        asset: {
          id: Number(asset.id),
          asset_code: String(asset.asset_code || ""),
          asset_name: String(asset.asset_name || ""),
          category: String(asset.category || ""),
        },
        check_date,
        previous_smu_hours,
        previous_smu_source:
          previous_smu_hours != null ? "daily_hours_or_prior_prestart" : null,
        existing_check: existing
          ? {
              id: Number(existing.id),
              smu_hours: existing.smu_hours == null ? null : Number(existing.smu_hours),
              inspector_name: existing.inspector_name || "",
              notes: existing.notes || "",
              checklist,
            }
          : null,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/machine-prestart", async (req, reply) => {
    try {
      const asset_code = String(req.body?.asset_code || "").trim().toUpperCase();
      const check_date = String(req.body?.check_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      const checklistObj = req.body?.checklist || {};
      const smu_raw = req.body?.smu_hours;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "check_date must be YYYY-MM-DD" });

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name, category
        FROM assets
        WHERE UPPER(asset_code) = UPPER(?)
        LIMIT 1
      `).get(asset_code);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const profileId = resolveMachinePrestartProfile(asset.category, asset.asset_name, asset.asset_code);
      if (!profileId) {
        return reply.code(400).send({
          ok: false,
          error: "No machine pre-start template for this asset category/name.",
        });
      }
      const mode = machinePrestartCheckMode(profileId);
      if (!mode) return reply.code(500).send({ ok: false, error: "Template resolution failed" });

      let smu_hours = null;
      if (smu_raw != null && String(smu_raw).trim() !== "") {
        const smu = Number(smu_raw);
        if (!Number.isFinite(smu) || smu < 0) {
          return reply.code(400).send({ ok: false, error: "smu_hours must be a valid number >= 0 when provided." });
        }
        smu_hours = Number(smu.toFixed(1));
      }

      const checklist = normalizeMachinePrestartChecklist(profileId, checklistObj);
      const unanswered = unansweredChecks(checklist, checklistObj);
      if (unanswered.length) {
        return reply.code(400).send({
          ok: false,
          error: `Mark every check OK or Fault before submitting (${unanswered.join(", ")}).`,
        });
      }
      const faults = prestartFaultList(checklist, req.body?.faults);
      const savedNotes = notesWithFaults(notes, faults);

      const checklistJson = JSON.stringify(checklistToJsonObject(checklist));

      const existing = getMachinePrestartCheckRow(Number(asset.id), check_date, mode);

      let checkId = 0;
      if (existing?.id) {
        checkId = Number(existing.id);
        db.prepare(`
          UPDATE vehicle_ldv_checks
          SET
            inspector_name = ?,
            notes = ?,
            checklist_json = ?,
            smu_hours = ?,
            odometer_km = NULL,
            check_mode = ?,
            updated_at = datetime('now')
          WHERE id = ?
        `).run(inspector_name, savedNotes, checklistJson, smu_hours, mode, checkId);
      } else {
        const ins = db.prepare(`
          INSERT INTO vehicle_ldv_checks (
            asset_id, uuid, site_code, check_date, vehicle_registration, odometer_km,
            inspector_name, notes, check_mode, checklist_json, smu_hours, updated_at
          )
          VALUES (?, ?, ?, ?, ?, NULL, ?, ?, ?, ?, ?, datetime('now'))
        `).run(
          Number(asset.id),
          crypto.randomUUID(),
          site_code,
          check_date,
          String(asset.asset_code || ""),
          inspector_name,
          savedNotes,
          mode,
          checklistJson,
          smu_hours
        );
        checkId = Number(ins.lastInsertRowid);
      }

      let dailySync = { synced: false };
      if (smu_hours != null) {
        const previousSmu = getLatestMachineSmuHours(Number(asset.id), check_date, {
          excludeCheckId: checkId,
        });
        dailySync = syncMachinePrestartToDailyHours(
          Number(asset.id),
          check_date,
          smu_hours,
          inspector_name,
          previousSmu
        );
      }

      const faultWo = syncPrestartFaultWorkOrder(db, {
        assetId: Number(asset.id), checkId, siteCode: site_code, checkDate: check_date, operator: inspector_name, faults,
      });

      return reply.send({
        ok: true,
        id: checkId,
        faults: faults.length,
        fault_work_order_id: faultWo && !faultWo.closed ? faultWo.work_order_id : null,
        asset_code: String(asset.asset_code || ""),
        check_date,
        profile_id: profileId,
        check_mode: mode,
        smu_hours,
        daily_input_sync: dailySync,
        message: faultMessage(faults, faultWo) || "Machine pre-start saved to IRONLOG.",
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // POST /api/maintenance/machine-prestart/hours-correction — fix wrong SMU + daily hours (supervisor)
  app.post("/machine-prestart/hours-correction", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;

      const asset_code = String(req.body?.asset_code || "").trim().toUpperCase();
      const work_date = String(req.body?.work_date || req.body?.check_date || "").trim()
        || new Date().toISOString().slice(0, 10);
      const closing_hours = Number(
        req.body?.closing_hours ?? req.body?.correct_smu_hours ?? req.body?.smu_hours
      );
      const opening_hours_raw = req.body?.opening_hours;
      const inspector_name = String(req.body?.inspector_name || "Supervisor correction").trim() || "Supervisor correction";
      const correction_note = String(req.body?.notes || "Supervisor hours correction").trim() || "Supervisor hours correction";
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      if (!isDate(work_date)) return reply.code(400).send({ ok: false, error: "work_date must be YYYY-MM-DD" });
      if (!Number.isFinite(closing_hours) || closing_hours < 0) {
        return reply.code(400).send({ ok: false, error: "closing_hours must be a valid number >= 0" });
      }

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name, category
        FROM assets
        WHERE UPPER(asset_code) = UPPER(?)
        LIMIT 1
      `).get(asset_code);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const profileId = resolveMachinePrestartProfile(asset.category, asset.asset_name, asset.asset_code);
      if (!profileId) {
        return reply.code(400).send({ ok: false, error: "No machine pre-start template for this asset." });
      }
      const mode = machinePrestartCheckMode(profileId);
      if (!mode) return reply.code(500).send({ ok: false, error: "Template resolution failed" });

      const assetId = Number(asset.id);
      const existing = getMachinePrestartCheckRow(assetId, work_date, mode);

      let checkId = Number(existing?.id || 0);
      const previousSmu = getLatestMachineSmuHours(assetId, work_date, {
        excludeCheckId: checkId,
      });
      const opening_hours =
        opening_hours_raw != null && String(opening_hours_raw).trim() !== ""
          ? Number(opening_hours_raw)
          : previousSmu;
      if (!Number.isFinite(opening_hours) || opening_hours < 0) {
        return reply.code(400).send({
          ok: false,
          error: "opening_hours could not be resolved — pass opening_hours explicitly",
        });
      }
      if (closing_hours < opening_hours) {
        return reply.code(400).send({
          ok: false,
          error: `Closing hours (${closing_hours}) cannot be less than opening hours (${opening_hours}).`,
          previous_smu_hours: previousSmu,
          opening_hours,
        });
      }

      const checklistJsonOut = (() => {
        if (existing?.checklist_json && String(existing.checklist_json).trim()) {
          return String(existing.checklist_json);
        }
        const checklist = normalizeMachinePrestartChecklist(
          profileId,
          Object.fromEntries(
            normalizeMachinePrestartChecklist(profileId, {}).map((c) => [c.key, true])
          )
        );
        return JSON.stringify(checklistToJsonObject(checklist));
      })();

      if (checkId > 0) {
        db.prepare(`
          UPDATE vehicle_ldv_checks
          SET smu_hours = ?, inspector_name = ?, notes = ?, checklist_json = COALESCE(NULLIF(checklist_json, ''), ?),
              check_mode = ?, updated_at = datetime('now')
          WHERE id = ?
        `).run(closing_hours, inspector_name, correction_note, checklistJsonOut, mode, checkId);
      } else {
        const ins = db.prepare(`
          INSERT INTO vehicle_ldv_checks (
            asset_id, uuid, site_code, check_date, vehicle_registration, odometer_km,
            inspector_name, notes, check_mode, checklist_json, smu_hours, updated_at
          )
          VALUES (?, ?, ?, ?, ?, NULL, ?, ?, ?, ?, ?, datetime('now'))
        `).run(
          assetId,
          crypto.randomUUID(),
          site_code,
          work_date,
          String(asset.asset_code || ""),
          inspector_name,
          correction_note,
          mode,
          checklistJsonOut,
          closing_hours
        );
        checkId = Number(ins.lastInsertRowid);
      }

      const run_hours = upsertMachineDailyHoursCorrection(
        assetId,
        work_date,
        opening_hours,
        closing_hours,
        inspector_name,
        correction_note
      );

      return reply.send({
        ok: true,
        asset_code,
        work_date,
        check_id: checkId,
        profile_id: profileId,
        opening_hours: Number(opening_hours.toFixed(1)),
        closing_hours: Number(closing_hours.toFixed(1)),
        run_hours: Number(run_hours.toFixed(1)),
        previous_smu_hours: previousSmu == null ? null : Number(previousSmu.toFixed(1)),
        message: `Hours corrected for ${asset_code} on ${work_date}.`,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/vehicle-ldv-checks", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const check_date = String(req.body?.check_date || "").trim() || new Date().toISOString().slice(0, 10);
      const vehicle_registration = String(req.body?.vehicle_registration || "").trim() || null;
      const odometer_km = req.body?.odometer_km != null && req.body?.odometer_km !== "" ? Number(req.body.odometer_km) : null;
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "check_date must be YYYY-MM-DD" });

      const asset = db.prepare(`SELECT id FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const ins = db.prepare(`
        INSERT INTO vehicle_ldv_checks (
          asset_id, uuid, site_code, check_date, vehicle_registration, odometer_km, inspector_name, notes, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(asset_id, crypto.randomUUID(), site_code, check_date, vehicle_registration, odometer_km, inspector_name, notes);

      return reply.send({ ok: true, id: Number(ins.lastInsertRowid) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/vehicle-ldv-checks/:id/photo", async (req, reply) => {
    try {
      const checkId = Number(req.params?.id || 0);
      if (!checkId) return reply.code(400).send({ ok: false, error: "Invalid check id" });

      const row = db.prepare(`SELECT id FROM vehicle_ldv_checks WHERE id = ?`).get(checkId);
      if (!row) return reply.code(404).send({ ok: false, error: "Vehicle check not found" });

      const part = await req.file();
      if (!part) return reply.code(400).send({ ok: false, error: "Upload file field named 'file'" });

      const rawBuffer = await part.toBuffer();
      let fileBuffer;
      try {
        fileBuffer = await normalizeUploadedPhoto(rawBuffer);
      } catch {
        return reply.code(400).send({ ok: false, error: "Could not process image file" });
      }

      const safe = `ldv_${checkId}_${Date.now()}_${Math.floor(Math.random() * 100000)}.jpg`;
      const absPath = path.join(vehicleLdvcDir, safe);
      await fs.promises.writeFile(absPath, fileBuffer);

      const caption = String(req.query?.caption || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const relPath = path.join("uploads", "vehicle-ldv-checks", safe).replace(/\\/g, "/");

      const ins = db.prepare(`
        INSERT INTO vehicle_ldv_check_photos (check_id, uuid, site_code, file_path, caption, markers_json, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(checkId, crypto.randomUUID(), site_code, relPath, caption, null);

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        check_id: checkId,
        file_path: `/${relPath}`,
        caption,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.patch("/vehicle-ldv-checks/photos/:photoId", async (req, reply) => {
    try {
      const photoId = Number(req.params?.photoId || 0);
      if (!photoId) return reply.code(400).send({ ok: false, error: "Invalid photo id" });

      const photo = db.prepare(`SELECT id, check_id FROM vehicle_ldv_check_photos WHERE id = ?`).get(photoId);
      if (!photo) return reply.code(404).send({ ok: false, error: "Photo not found" });

      const body = req.body || {};
      const markers = body.markers;
      const caption = body.caption != null ? String(body.caption).trim() || null : undefined;

      if (markers != null) {
        if (!Array.isArray(markers)) return reply.code(400).send({ ok: false, error: "markers must be an array" });
        if (markers.length > 80) return reply.code(400).send({ ok: false, error: "too many markers (max 80)" });
        for (const m of markers) {
          const x = Number(m?.x);
          const y = Number(m?.y);
          if (!Number.isFinite(x) || !Number.isFinite(y) || x < 0 || x > 1 || y < 0 || y > 1) {
            return reply.code(400).send({ ok: false, error: "each marker needs x,y in [0,1] (fraction of image)" });
          }
        }
        const json = JSON.stringify(
          markers.map((m) => ({
            x: Number(m.x),
            y: Number(m.y),
            label: m.label != null ? String(m.label).slice(0, 120) : "",
            note: m.note != null ? String(m.note).slice(0, 500) : "",
          }))
        );
        db.prepare(`
          UPDATE vehicle_ldv_check_photos
          SET markers_json = ?, updated_at = datetime('now')
          WHERE id = ?
        `).run(json, photoId);
      }
      if (caption !== undefined) {
        db.prepare(`
          UPDATE vehicle_ldv_check_photos
          SET caption = ?, updated_at = datetime('now')
          WHERE id = ?
        `).run(caption, photoId);
      }

      const updated = db.prepare(`SELECT id, check_id, file_path, caption, markers_json FROM vehicle_ldv_check_photos WHERE id = ?`).get(photoId);
      let markersOut = [];
      try {
        markersOut = updated.markers_json ? JSON.parse(updated.markers_json) : [];
      } catch {
        markersOut = [];
      }

      return reply.send({
        ok: true,
        photo: {
          id: Number(updated.id),
          check_id: Number(updated.check_id),
          file_path: `/${String(updated.file_path || "").replace(/\\/g, "/")}`,
          caption: updated.caption,
          markers: markersOut,
        },
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.get("/damage-reports", async (req, reply) => {
    try {
      const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
      const assetId = Number(req.query?.asset_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const responsiblePerson = String(req.query?.responsible_person || "").trim();
      const pendingInvestigationRaw = String(req.query?.pending_investigation || "").trim();
      const hseReportAvailableRaw = String(req.query?.hse_report_available || "").trim();
      const params = [];
      const where = [];
      if (assetId > 0) {
        where.push("dr.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("dr.report_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("dr.report_date <= ?");
        params.push(end);
      }
      if (responsiblePerson) {
        where.push("UPPER(COALESCE(dr.responsible_person, '')) LIKE UPPER(?)");
        params.push(`%${responsiblePerson}%`);
      }
      if (pendingInvestigationRaw === "0" || pendingInvestigationRaw === "1") {
        where.push("COALESCE(dr.pending_investigation, 0) = ?");
        params.push(Number(pendingInvestigationRaw));
      }
      if (hseReportAvailableRaw === "0" || hseReportAvailableRaw === "1") {
        where.push("COALESCE(dr.hse_report_available, 0) = ?");
        params.push(Number(hseReportAvailableRaw));
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
          dr.created_at,
          a.asset_code,
          a.asset_name
        FROM manager_damage_reports dr
        JOIN assets a ON a.id = dr.asset_id
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY dr.report_date DESC, dr.id DESC
      `).all(...params);

      const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
      let photosByReport = new Map();
      if (ids.length) {
        const marks = ids.map(() => "?").join(",");
        const photos = db.prepare(`
          SELECT
            id,
            ${dmgPhotoReportCol} AS damage_report_id,
            ${dmgPhotoPathCol} AS file_path,
            ${dmgPhotoCaptionCol} AS caption,
            ${dmgPhotoCreatedCol} AS created_at
          FROM manager_damage_report_photos
          WHERE ${dmgPhotoReportCol} IN (${marks})
          ORDER BY id ASC
        `).all(...ids);
        photosByReport = photos.reduce((m, p) => {
          const k = Number(p.damage_report_id);
          if (!m.has(k)) m.set(k, []);
          m.get(k).push(p);
          return m;
        }, new Map());
      }

      return reply.send({
        ok: true,
        rows: rows.map((r) => ({
          ...r,
          photos: photosByReport.get(Number(r.id)) || [],
        })),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/damage-reports", async (req, reply) => {
    try {
      const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
      const asset_id = Number(req.body?.asset_id || 0);
      const report_date = String(req.body?.report_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const hour_meter_raw = req.body?.hour_meter;
      const hour_meter =
        hour_meter_raw == null || String(hour_meter_raw).trim() === ""
          ? null
          : Number(hour_meter_raw);
      const damage_location = String(req.body?.damage_location || "").trim() || null;
      const severity = String(req.body?.severity || "").trim() || null;
      const damage_description = String(req.body?.damage_description || "").trim() || null;
      const immediate_action = String(req.body?.immediate_action || "").trim() || null;
      const out_of_service = Number(req.body?.out_of_service || 0) ? 1 : 0;
      const damage_time = String(req.body?.damage_time || "").trim() || null;
      const responsible_person = String(req.body?.responsible_person || "").trim() || null;
      const pending_investigation = Number(req.body?.pending_investigation || 0) ? 1 : 0;
      const hse_report_available = Number(req.body?.hse_report_available || 0) ? 1 : 0;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(report_date)) return reply.code(400).send({ ok: false, error: "report_date must be YYYY-MM-DD" });
      if (hour_meter != null && !Number.isFinite(hour_meter)) {
        return reply.code(400).send({ ok: false, error: "hour_meter must be numeric" });
      }
      if (!damage_location) return reply.code(400).send({ ok: false, error: "damage_location is required" });
      if (!severity) return reply.code(400).send({ ok: false, error: "severity is required" });
      if (!damage_description) return reply.code(400).send({ ok: false, error: "damage_description is required" });
      if (!immediate_action) return reply.code(400).send({ ok: false, error: "immediate_action is required" });
      if (damage_time && !/^\d{2}:\d{2}$/.test(damage_time)) {
        return reply.code(400).send({ ok: false, error: "damage_time must be HH:MM" });
      }

      const asset = db.prepare(`SELECT id FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const ins = db.prepare(`
        INSERT INTO manager_damage_reports (
          asset_id, uuid, site_code, report_date, ${drInspectorCol}, hour_meter,
          damage_location, severity, damage_description, immediate_action, out_of_service,
          damage_time, responsible_person, pending_investigation, hse_report_available, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(
        asset_id,
        crypto.randomUUID(),
        site_code,
        report_date,
        inspector_name,
        hour_meter,
        damage_location,
        severity,
        damage_description,
        immediate_action,
        out_of_service,
        damage_time,
        responsible_person,
        pending_investigation,
        hse_report_available
      );

      return reply.send({ ok: true, id: Number(ins.lastInsertRowid) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/damage-reports/:id/photo", async (req, reply) => {
    try {
      const reportId = Number(req.params?.id || 0);
      if (!reportId) return reply.code(400).send({ ok: false, error: "Invalid damage report id" });

      const report = db.prepare(`SELECT id FROM manager_damage_reports WHERE id = ?`).get(reportId);
      if (!report) return reply.code(404).send({ ok: false, error: "Damage report not found" });

      const part = await req.file();
      if (!part) return reply.code(400).send({ ok: false, error: "Upload file field named 'file'" });

      const extRaw = path.extname(part.filename || "").toLowerCase();
      const ext = [".jpg", ".jpeg", ".png", ".webp"].includes(extRaw) ? extRaw : ".jpg";
      const safe = `mdr_${reportId}_${Date.now()}_${Math.floor(Math.random() * 100000)}${ext}`;
      const absPath = path.join(damageReportsDir, safe);
      await fs.promises.writeFile(absPath, await part.toBuffer());

      const caption = String(req.query?.caption || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const relPath = path.join("uploads", "manager-damage-reports", safe).replace(/\\/g, "/");

      const hasReportId = hasColumn("manager_damage_report_photos", "damage_report_id");
      const hasLegacyReportId = hasColumn("manager_damage_report_photos", "manager_damage_report_id");
      const linkCols = [];
      const linkVals = [];
      if (hasReportId) {
        linkCols.push("damage_report_id");
        linkVals.push(reportId);
      }
      if (hasLegacyReportId) {
        linkCols.push("manager_damage_report_id");
        linkVals.push(reportId);
      }
      if (!linkCols.length) {
        linkCols.push(dmgPhotoReportCol);
        linkVals.push(reportId);
      }

      const hasImageData = hasColumn("manager_damage_report_photos", "image_data");
      const insertCols = [...linkCols, "uuid", "site_code", "file_path", ...(hasImageData ? ["image_data"] : []), "caption", "updated_at"];
      const placeholders = [...insertCols.map((c) => (c === "updated_at" ? "datetime('now')" : "?"))].join(", ");
      const ins = db.prepare(`
        INSERT INTO manager_damage_report_photos (${insertCols.join(", ")})
        VALUES (${placeholders})
      `).run(...linkVals, crypto.randomUUID(), site_code, relPath, ...(hasImageData ? [relPath] : []), caption);

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        damage_report_id: reportId,
        file_path: `/${relPath}`,
        caption,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });
}
