// IRONLOG/api/routes/dashboard/fuel.routes.js — Fuel log, FAMS sync, baselines, comparisons and shift scenarios.
// Registered by routes/dashboard.routes.js; shared helpers arrive through ctx.
import ExcelJS from "exceljs";
import { aggregateFuelBenchmarkByCategory } from "../../utils/fuelBenchmarkAggregate.js";
import { db } from "../../db/client.js";
import { ensureFamsFuelSchema, getFamsSyncStatus, listFamsUnmatched, previewFamsLegacyDuplicates, removeFamsLegacyDuplicates, syncFamsFuel } from "../../utils/famsFuel.js";
import { famsSelectedDateRange } from "../../utils/famsFuelRange.js";
import { fuelBenchmarkAssetsInRangeSql, sqlFuelMetricModeExpr } from "../../utils/fuelMetricMode.js";
import { summarizeFuelBenchmarkRows } from "../../utils/fuelRunFromLogs.js";
import { assetConsumption, consumptionNote, ensureFuelBenchmarkSchema, sortBenchmarkRows } from "../../utils/fuelConsumption.js";
import { resolveLogCostCenterCode } from "../../utils/costAllocation.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerFuelRoutes(app, ctx) {
  const {
    addShiftScenarioWorkbook,
    buildFuelDailyUsageSeries,
    buildShiftScenarioData,
    fuelPreviousPeriodRange,
    requireRoles,
    scenarioHoursFromRequest,
    siteCodeFromReq,
    todayYYYYMMDD,
  } = ctx;
  ensureFuelBenchmarkSchema(db);

  // GET /api/dashboard/fuel/fams/status — FAMS auto-sync status (no secrets)
  app.get("/fuel/fams/status", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores", "operator", "artisan"])) return;
    try {
      ensureFamsFuelSchema();
      return reply.send(getFamsSyncStatus());
    } catch (err) {
      return reply.code(500).send({ ok: false, error: err?.message || String(err) });
    }
  });

  // GET /api/dashboard/fuel/fams/unmatched — unmatched FAMS registrations
  app.get("/fuel/fams/unmatched", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    try {
      ensureFamsFuelSchema();
      const rows = listFamsUnmatched({ limit: Number(req.query?.limit || 200) });
      return reply.send({ ok: true, rows });
    } catch (err) {
      return reply.code(500).send({ ok: false, error: err?.message || String(err) });
    }
  });

  // POST /api/dashboard/fuel/fams/sync — manual "Sync FAMS Now" (same logic as hourly job)
  app.post("/fuel/fams/sync", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    try {
      const force = Boolean(req.body?.force);
      const startDate = String(req.body?.start_date || "").trim();
      const endDate = String(req.body?.end_date || "").trim();
      let range = null;
      if (startDate || endDate) {
        if (!startDate || !endDate) {
          return reply.code(400).send({ ok: false, error: "Select both a From date and a To date" });
        }
        try {
          range = famsSelectedDateRange({ startDate, endDate });
        } catch (err) {
          return reply.code(400).send({ ok: false, error: err?.message || String(err) });
        }
      }
      const result = await syncFamsFuel({ log: req.log || console, force, range });
      if (!result?.ok && result?.error) {
        return reply.code(502).send(result);
      }
      return reply.send({ ...result, status: getFamsSyncStatus() });
    } catch (err) {
      // Never crash the API on FAMS failure
      return reply.code(500).send({ ok: false, error: err?.message || String(err) });
    }
  });

  // POST /api/dashboard/fuel/fams/duplicates/preview — read-only exact legacy/FAMS duplicate check
  app.post("/fuel/fams/duplicates/preview", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const startDate = String(req.body?.start_date || "").trim();
    const endDate = String(req.body?.end_date || "").trim();
    try {
      const range = famsSelectedDateRange({ startDate, endDate });
      return reply.send({ ok: true, range: { start: range.startYmd, end: range.endYmd }, ...previewFamsLegacyDuplicates({
        startDate: range.startYmd,
        endDate: range.endYmd,
      }) });
    } catch (err) {
      return reply.code(400).send({ ok: false, error: err?.message || String(err) });
    }
  });

  // POST /api/dashboard/fuel/fams/duplicates/remove — only removes exact, unambiguous legacy duplicates
  app.post("/fuel/fams/duplicates/remove", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const startDate = String(req.body?.start_date || "").trim();
    const endDate = String(req.body?.end_date || "").trim();
    try {
      const range = famsSelectedDateRange({ startDate, endDate });
      const result = removeFamsLegacyDuplicates({ startDate: range.startYmd, endDate: range.endYmd });
      return reply.send({ ok: true, range: { start: range.startYmd, end: range.endYmd }, ...result });
    } catch (err) {
      return reply.code(400).send({ ok: false, error: err?.message || String(err) });
    }
  });

  // POST /api/dashboard/fuel/log
  // Body: {
  //   asset_code, log_date?, liters,
  //   hours_run?, meter_run_value?, meter_unit? ('hours'|'km'),
  //   source?, force_duplicate?
  // }
  app.post("/fuel/log", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "operator", "artisan"])) return;
    const body = req.body || {};
    const asset_code = String(body.asset_code || "").trim();
    const log_date =
      body.log_date != null && String(body.log_date).trim() !== ""
        ? String(body.log_date).trim()
        : todayYYYYMMDD();
    const liters = Number(body.liters ?? 0);
    const hours_run =
      body.hours_run != null && String(body.hours_run).trim() !== ""
        ? Number(body.hours_run)
        : null;
    const meter_run_value =
      body.meter_run_value != null && String(body.meter_run_value).trim() !== ""
        ? Number(body.meter_run_value)
        : null;
    const meter_unit_raw = String(body.meter_unit || "").trim().toLowerCase();
    const meter_unit = meter_unit_raw === "km" ? "km" : meter_unit_raw === "hours" ? "hours" : null;
    const source =
      body.source != null && String(body.source).trim() !== ""
        ? String(body.source).trim()
        : null;
    const forceDuplicate = Boolean(body.force_duplicate);

    if (!asset_code) return reply.code(400).send({ error: "asset_code is required" });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(log_date)) {
      return reply.code(400).send({ error: "log_date must be YYYY-MM-DD" });
    }
    if (!Number.isFinite(liters) || liters <= 0) {
      return reply.code(400).send({ error: "liters must be > 0" });
    }
    if (hours_run != null && (!Number.isFinite(hours_run) || hours_run < 0)) {
      return reply.code(400).send({ error: "hours_run must be >= 0" });
    }
    if (meter_run_value != null && (!Number.isFinite(meter_run_value) || meter_run_value < 0)) {
      return reply.code(400).send({ error: "meter_run_value must be >= 0" });
    }

    const asset = db.prepare(`SELECT id, cost_center_code FROM assets WHERE asset_code = ?`).get(asset_code);
    if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });
    const cost_center_code = resolveLogCostCenterCode(db, asset.id, body.cost_center_code);

    // Guard against accidental double-capture: exact same entry within 60 seconds.
    if (!forceDuplicate) {
      const recentDup = db.prepare(`
        SELECT id, created_at
        FROM fuel_logs
        WHERE asset_id = ?
          AND log_date = ?
          AND ABS(COALESCE(liters, 0) - ?) < 0.000001
          AND (
            (? IS NULL AND hours_run IS NULL)
            OR ABS(COALESCE(hours_run, 0) - COALESCE(?, 0)) < 0.000001
          )
          AND (
            (? IS NULL AND meter_run_value IS NULL)
            OR ABS(COALESCE(meter_run_value, 0) - COALESCE(?, 0)) < 0.000001
          )
          AND COALESCE(LOWER(meter_unit), '') = COALESCE(LOWER(?), '')
          AND COALESCE(source, '') = COALESCE(?, '')
          AND datetime(created_at) >= datetime('now', '-60 seconds')
        ORDER BY id DESC
        LIMIT 1
      `).get(asset.id, log_date, liters, hours_run, hours_run, meter_run_value, meter_run_value, meter_unit, source);

      if (recentDup) {
        return reply.code(409).send({
          error: "possible_duplicate_recent",
          message: "Possible duplicate: same fuel input was saved recently. Confirm to save again.",
          duplicate_id: Number(recentDup.id),
          duplicate_created_at: recentDup.created_at,
        });
      }
    }

    const hours_final = hours_run != null ? hours_run : (meter_unit === "hours" ? meter_run_value : null);
    const ins = db.prepare(`
      INSERT INTO fuel_logs (asset_id, log_date, liters, source, hours_run, meter_run_value, meter_unit, cost_center_code)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `).run(asset.id, log_date, liters, source, hours_final, meter_run_value, meter_unit, cost_center_code);

    writeAudit(db, req, {
      module: "fuel",
      action: "manual_log",
      entity_type: "asset",
      entity_id: asset_code,
      payload: { log_date, liters, hours_run: hours_final, meter_run_value, meter_unit, source, cost_center_code },
    });

    return reply.send({
      ok: true,
      id: Number(ins.lastInsertRowid),
      asset_code,
      cost_center_code,
      log_date,
      liters,
      hours_run: hours_final,
      meter_run_value,
      meter_unit,
      source,
    });
  });

  // POST /api/dashboard/fuel/repair-meter-chain
  // Body: { asset_code?: string }
  // Repairs day opening meter to previous day's closing meter when mismatch detected.
  app.post("/fuel/repair-meter-chain", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const asset_code = String(req.body?.asset_code || "").trim();

    let assetFilterSql = "";
    const params = [];
    if (asset_code) {
      const asset = db.prepare(`SELECT id, asset_code FROM assets WHERE asset_code = ?`).get(asset_code);
      if (!asset) return reply.code(404).send({ error: `asset not found: ${asset_code}` });
      assetFilterSql = "AND d2.asset_id = ?";
      params.push(asset.id);
    }

    const candidates = db.prepare(`
      SELECT
        d2.id,
        d2.asset_id,
        d2.work_date,
        d2.opening_hours AS old_opening_hours,
        d2.closing_hours AS closing_hours,
        d2.hours_run AS old_hours_run,
        (
          SELECT d1.closing_hours
          FROM daily_hours d1
          WHERE d1.asset_id = d2.asset_id
            AND d1.work_date < d2.work_date
            AND d1.closing_hours IS NOT NULL
          ORDER BY d1.work_date DESC, d1.id DESC
          LIMIT 1
        ) AS expected_opening_hours
      FROM daily_hours d2
      WHERE 1 = 1
        ${assetFilterSql}
    `).all(...params).filter((r) => {
      if (r.expected_opening_hours == null) return false;
      if (r.old_opening_hours == null) return true;
      return Math.abs(Number(r.old_opening_hours) - Number(r.expected_opening_hours)) > 0.0001;
    });

    const updateRow = db.prepare(`
      UPDATE daily_hours
      SET
        opening_hours = ?,
        hours_run = CASE
          WHEN closing_hours IS NOT NULL AND closing_hours >= ? THEN (closing_hours - ?)
          ELSE hours_run
        END
      WHERE id = ?
    `);

    const tx = db.transaction(() => {
      for (const r of candidates) {
        const nextOpen = Number(r.expected_opening_hours);
        updateRow.run(nextOpen, nextOpen, nextOpen, r.id);
      }
    });
    tx();

    writeAudit(db, req, {
      module: "fuel",
      action: "repair_meter_chain",
      entity_type: "asset",
      entity_id: asset_code || "all",
      payload: { repaired_rows: candidates.length },
    });

    return reply.send({
      ok: true,
      asset_code: asset_code || null,
      repaired_rows: candidates.length,
      sample: candidates.slice(0, 20).map((r) => ({
        id: Number(r.id),
        asset_id: Number(r.asset_id),
        work_date: r.work_date,
        old_opening_hours: r.old_opening_hours == null ? null : Number(r.old_opening_hours),
        expected_opening_hours: Number(r.expected_opening_hours),
      })),
    });
  });

  // POST /api/dashboard/fuel/clear-from-date/preview
  // Body: { from_date: 'YYYY-MM-DD', asset_code?: string, clear_daily_hours?: boolean }
  app.post("/fuel/clear-from-date/preview", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const fromDate = String(req.body?.from_date || "").trim();
    const assetCode = String(req.body?.asset_code || "").trim();
    const clearDailyHours = Boolean(req.body?.clear_daily_hours);

    if (!/^\d{4}-\d{2}-\d{2}$/.test(fromDate)) {
      return reply.code(400).send({ error: "from_date must be YYYY-MM-DD" });
    }

    let assetId = null;
    if (assetCode) {
      const asset = db.prepare(`SELECT id FROM assets WHERE asset_code = ?`).get(assetCode);
      if (!asset) return reply.code(404).send({ error: `asset not found: ${assetCode}` });
      assetId = Number(asset.id);
    }

    const whereSql = assetId != null ? "WHERE log_date >= ? AND asset_id = ?" : "WHERE log_date >= ?";
    const whereParams = assetId != null ? [fromDate, assetId] : [fromDate];
    const dayRows = db.prepare(`
      SELECT DISTINCT asset_id, log_date
      FROM fuel_logs
      ${whereSql}
    `).all(...whereParams);
    const logsToDelete = Number(
      db.prepare(`SELECT COUNT(*) AS n FROM fuel_logs ${whereSql}`).get(...whereParams)?.n || 0
    );

    return reply.send({
      ok: true,
      from_date: fromDate,
      asset_code: assetCode || null,
      clear_daily_hours: clearDailyHours,
      deleted_logs: logsToDelete,
      affected_days: dayRows.length,
    });
  });

  // POST /api/dashboard/fuel/clear-from-date
  // Body: { from_date: 'YYYY-MM-DD', asset_code?: string, clear_daily_hours?: boolean }
  app.post("/fuel/clear-from-date", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const fromDate = String(req.body?.from_date || "").trim();
    const assetCode = String(req.body?.asset_code || "").trim();
    const clearDailyHours = Boolean(req.body?.clear_daily_hours);

    if (!/^\d{4}-\d{2}-\d{2}$/.test(fromDate)) {
      return reply.code(400).send({ error: "from_date must be YYYY-MM-DD" });
    }

    let assetId = null;
    if (assetCode) {
      const asset = db.prepare(`SELECT id, asset_code FROM assets WHERE asset_code = ?`).get(assetCode);
      if (!asset) return reply.code(404).send({ error: `asset not found: ${assetCode}` });
      assetId = Number(asset.id);
    }

    const whereSql = assetId != null ? "WHERE log_date >= ? AND asset_id = ?" : "WHERE log_date >= ?";
    const whereParams = assetId != null ? [fromDate, assetId] : [fromDate];

    const tx = db.transaction(() => {
      const dayRows = db.prepare(`
        SELECT DISTINCT asset_id, log_date
        FROM fuel_logs
        ${whereSql}
      `).all(...whereParams);

      const logsToDelete = Number(
        db.prepare(`SELECT COUNT(*) AS n FROM fuel_logs ${whereSql}`).get(...whereParams)?.n || 0
      );

      const deleted = db.prepare(`DELETE FROM fuel_logs ${whereSql}`).run(...whereParams);

      let clearedDailyHoursRows = 0;
      if (clearDailyHours && dayRows.length > 0) {
        const clearDaily = db.prepare(`
          UPDATE daily_hours
          SET opening_hours = NULL,
              closing_hours = NULL,
              hours_run = NULL
          WHERE asset_id = ?
            AND work_date = ?
        `);
        for (const d of dayRows) {
          const res = clearDaily.run(Number(d.asset_id), String(d.log_date));
          clearedDailyHoursRows += Number(res.changes || 0);
        }
      }

      return {
        deleted_logs: Number(deleted.changes || logsToDelete || 0),
        affected_days: dayRows.length,
        cleared_daily_hours_rows: clearedDailyHoursRows,
      };
    });

    const summary = tx();

    writeAudit(db, req, {
      module: "fuel",
      action: "clear_from_date",
      entity_type: "asset",
      entity_id: assetCode || "all",
      payload: {
        from_date: fromDate,
        clear_daily_hours: clearDailyHours,
        deleted_logs: summary.deleted_logs,
        affected_days: summary.affected_days,
        cleared_daily_hours_rows: summary.cleared_daily_hours_rows,
      },
    });

    return reply.send({
      ok: true,
      from_date: fromDate,
      asset_code: assetCode || null,
      clear_daily_hours: clearDailyHours,
      deleted_logs: Number(summary.deleted_logs || 0),
      affected_days: Number(summary.affected_days || 0),
      cleared_daily_hours_rows: Number(summary.cleared_daily_hours_rows || 0),
    });
  });

  // POST /api/dashboard/fuel/machine-hours
  // Body: { fuel_log_id: number, opening_meter: number, closing_meter: number }
  app.post("/fuel/machine-hours", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "operator", "artisan"])) return;
    const fuelLogId = Number(req.body?.fuel_log_id || 0);
    const openingMeter = Number(req.body?.opening_meter);
    const closingMeter = Number(req.body?.closing_meter);

    if (!Number.isInteger(fuelLogId) || fuelLogId <= 0) {
      return reply.code(400).send({ error: "fuel_log_id must be a valid integer" });
    }
    if (!Number.isFinite(openingMeter) || openingMeter < 0) {
      return reply.code(400).send({ error: "opening_meter must be >= 0" });
    }
    if (!Number.isFinite(closingMeter) || closingMeter < 0) {
      return reply.code(400).send({ error: "closing_meter must be >= 0" });
    }
    if (closingMeter < openingMeter) {
      return reply.code(400).send({ error: "closing_meter must be >= opening_meter" });
    }

    const fuelLog = db.prepare(`
      SELECT fl.id, fl.asset_id, fl.log_date, a.asset_code
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.id = ?
      LIMIT 1
    `).get(fuelLogId);
    if (!fuelLog) return reply.code(404).send({ error: "fuel log not found" });

    const runDelta = Number((closingMeter - openingMeter).toFixed(3));
    const current = db.prepare(`
      SELECT COALESCE(LOWER(meter_unit), '') AS meter_unit
      FROM fuel_logs
      WHERE id = ?
      LIMIT 1
    `).get(fuelLogId);
    const unit = String(current?.meter_unit || "").toLowerCase();
    db.prepare(`
      UPDATE fuel_logs
      SET
        open_meter_value = ?,
        close_meter_value = ?,
        meter_run_value = ?,
        hours_run = CASE
          WHEN ? = 'km' THEN hours_run
          ELSE ?
        END
      WHERE id = ?
    `).run(
      Number(openingMeter.toFixed(3)),
      Number(closingMeter.toFixed(3)),
      Number(closingMeter.toFixed(3)),
      unit,
      runDelta,
      fuelLogId
    );

    writeAudit(db, req, {
      module: "fuel",
      action: "edit_machine_hours",
      entity_type: "fuel_log",
      entity_id: String(fuelLogId),
      payload: {
        asset_code: fuelLog.asset_code,
        log_date: fuelLog.log_date,
        opening_meter: Number(openingMeter.toFixed(3)),
        closing_meter: Number(closingMeter.toFixed(3)),
        hours_run: runDelta,
      },
    });

    return reply.send({
      ok: true,
      fuel_log_id: Number(fuelLogId),
      asset_code: fuelLog.asset_code,
      log_date: fuelLog.log_date,
      opening_meter: Number(openingMeter.toFixed(3)),
      closing_meter: Number(closingMeter.toFixed(3)),
      hours_run: runDelta,
    });
  });

  // GET /api/dashboard/fuel/baseline?asset_code=A300AM
  app.get("/fuel/baseline", async (req, reply) => {
    const asset_code = String(req.query?.asset_code || "").trim();

    if (asset_code) {
      const row = db.prepare(`
        SELECT
          a.id,
          a.asset_code,
          a.asset_name,
          COALESCE(a.baseline_fuel_l_per_hour, 5.0) AS baseline_fuel_l_per_hour,
          COALESCE(a.baseline_fuel_km_per_l, 2.0) AS baseline_fuel_km_per_l,
          CASE
            WHEN UPPER(COALESCE(a.asset_code, '')) GLOB 'V[0-9][0-9]AM' THEN 'km'
            ELSE COALESCE(NULLIF(TRIM(a.utilization_mode), ''), CASE
            WHEN LOWER(COALESCE(a.category, '')) LIKE '%truck%'
              OR LOWER(COALESCE(a.category, '')) LIKE '%vehicle%'
              OR LOWER(COALESCE(a.category, '')) LIKE '%ldv%'
              OR LOWER(COALESCE(a.category, '')) LIKE '%pickup%'
              OR LOWER(COALESCE(a.category, '')) LIKE '%bakkie%'
              OR LOWER(COALESCE(a.asset_code, '')) LIKE 'ldv%'
              OR UPPER(COALESCE(a.asset_code, '')) GLOB 'V[0-9][0-9]AM'
              OR LOWER(COALESCE(a.asset_name, '')) LIKE '%ldv%'
              THEN 'km'
            ELSE 'hours'
          END)
          END AS metric_mode
        FROM assets a
        WHERE a.asset_code = ?
      `).get(asset_code);
      if (!row) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });
      return reply.send({ ok: true, asset: row });
    }

    const rows = db.prepare(`
      SELECT
        a.id,
        a.asset_code,
        a.asset_name,
        COALESCE(a.baseline_fuel_l_per_hour, 5.0) AS baseline_fuel_l_per_hour,
        COALESCE(a.baseline_fuel_km_per_l, 2.0) AS baseline_fuel_km_per_l,
        CASE
          WHEN UPPER(COALESCE(a.asset_code, '')) GLOB 'V[0-9][0-9]AM' THEN 'km'
          ELSE COALESCE(NULLIF(TRIM(a.utilization_mode), ''), CASE
          WHEN LOWER(COALESCE(a.category, '')) LIKE '%truck%'
            OR LOWER(COALESCE(a.category, '')) LIKE '%vehicle%'
            OR LOWER(COALESCE(a.category, '')) LIKE '%ldv%'
            OR LOWER(COALESCE(a.category, '')) LIKE '%pickup%'
            OR LOWER(COALESCE(a.category, '')) LIKE '%bakkie%'
            OR LOWER(COALESCE(a.asset_code, '')) LIKE 'ldv%'
            OR UPPER(COALESCE(a.asset_code, '')) GLOB 'V[0-9][0-9]AM'
            OR LOWER(COALESCE(a.asset_name, '')) LIKE '%ldv%'
            THEN 'km'
          ELSE 'hours'
        END)
        END AS metric_mode
      FROM assets a
      WHERE a.active = 1
      ORDER BY a.asset_code ASC
      LIMIT 500
    `).all();

    return reply.send({ ok: true, rows });
  });

  // POST /api/dashboard/fuel/baseline
  // Body: { asset_code, metric_mode?, baseline_fuel_l_per_hour?, baseline_fuel_km_per_l? }
  app.post("/fuel/baseline", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const body = req.body || {};
    const asset_code = String(body.asset_code || "").trim();
    const mode = String(body.metric_mode || "").trim().toLowerCase();
    const baselineLph = body.baseline_fuel_l_per_hour != null ? Number(body.baseline_fuel_l_per_hour) : null;
    const baselineKmpl = body.baseline_fuel_km_per_l != null ? Number(body.baseline_fuel_km_per_l) : null;

    if (!asset_code) return reply.code(400).send({ error: "asset_code is required" });
    const asset = db.prepare(`
      SELECT
        id, asset_code, asset_name, category,
        CASE
          WHEN UPPER(COALESCE(asset_code, '')) GLOB 'V[0-9][0-9]AM' THEN 'km'
          ELSE COALESCE(NULLIF(TRIM(utilization_mode), ''), CASE
            WHEN LOWER(COALESCE(category, '')) LIKE '%truck%'
              OR LOWER(COALESCE(category, '')) LIKE '%vehicle%'
              OR LOWER(COALESCE(category, '')) LIKE '%ldv%'
              OR LOWER(COALESCE(category, '')) LIKE '%pickup%'
              OR LOWER(COALESCE(category, '')) LIKE '%bakkie%'
              OR LOWER(COALESCE(asset_code, '')) LIKE 'ldv%'
              OR UPPER(COALESCE(asset_code, '')) GLOB 'V[0-9][0-9]AM'
              OR LOWER(COALESCE(asset_name, '')) LIKE '%ldv%'
              THEN 'km'
            ELSE 'hours'
          END)
        END AS metric_mode
      FROM assets
      WHERE asset_code = ?
    `).get(asset_code);
    if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });

    const assetMode = mode === "km" || mode === "hours" ? mode : String(asset.metric_mode || "hours").toLowerCase();
    if (assetMode === "km") {
      const v = baselineKmpl != null ? baselineKmpl : baselineLph;
      if (!Number.isFinite(v) || v <= 0) {
        return reply.code(400).send({ error: "baseline_fuel_km_per_l must be > 0 for km mode" });
      }
      db.prepare(`UPDATE assets SET baseline_fuel_km_per_l = ?, fuel_benchmark_set_at = datetime('now') WHERE id = ?`).run(v, asset.id);
      writeAudit(db, req, {
        module: "fuel",
        action: "baseline_update",
        entity_type: "asset",
        entity_id: asset_code,
        payload: { metric_mode: "km", baseline_fuel_km_per_l: v },
      });
      return reply.send({
        ok: true,
        asset_code: asset.asset_code,
        asset_name: asset.asset_name,
        metric_mode: "km",
        baseline_fuel_km_per_l: Number(v.toFixed(3)),
      });
    }
    const v = baselineLph != null ? baselineLph : baselineKmpl;
    if (!Number.isFinite(v) || v <= 0) {
      return reply.code(400).send({ error: "baseline_fuel_l_per_hour must be > 0 for hours mode" });
    }
    db.prepare(`UPDATE assets SET baseline_fuel_l_per_hour = ?, fuel_benchmark_set_at = datetime('now') WHERE id = ?`).run(v, asset.id);

    writeAudit(db, req, {
      module: "fuel",
      action: "baseline_update",
      entity_type: "asset",
      entity_id: asset_code,
      payload: { metric_mode: "hours", baseline_fuel_l_per_hour: v },
    });

    return reply.send({
      ok: true,
      asset_code: asset.asset_code,
      asset_name: asset.asset_name,
      metric_mode: "hours",
      baseline_fuel_l_per_hour: Number(v.toFixed(3)),
    });
  });

  // GET /api/dashboard/fuel/period-compare?start=&end=&mode=&asset_code=
  // Daily fuel liters for selected period vs equal-length prior period (aligned by day index).
  app.get("/fuel/period-compare", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const modeFilter = String(req.query?.mode || "").trim().toLowerCase();
    const assetFilter = String(req.query?.asset_code || "").trim();

    if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }
    if (start > end) {
      return reply.code(400).send({ error: "start must be <= end" });
    }

    const prev = fuelPreviousPeriodRange(start, end);
    const current = buildFuelDailyUsageSeries(start, end, modeFilter, assetFilter);
    const previous = buildFuelDailyUsageSeries(prev.start, prev.end, modeFilter, assetFilter);
    const deltaLiters = Number((current.total_liters - previous.total_liters).toFixed(2));
    const deltaPct = previous.total_liters > 0
      ? Number(((deltaLiters / previous.total_liters) * 100).toFixed(1))
      : null;

    return reply.send({
      ok: true,
      mode: modeFilter || null,
      asset_code: assetFilter || null,
      current,
      previous,
      delta: {
        liters: deltaLiters,
        pct: deltaPct,
      },
    });
  });

  // GET /api/dashboard/shift-scenario?start=&end=&base_hours=11&scenario_hours=8&asset_code=
  app.get("/shift-scenario", async (req, reply) => {
    try {
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const assetCode = String(req.query?.asset_code || "").trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end) || start > end) {
        return reply.code(400).send({ error: "Provide valid start/end dates" });
      }
      const baseHours = scenarioHoursFromRequest(req.query?.base_hours, 11);
      const scenarioHours = scenarioHoursFromRequest(req.query?.scenario_hours, 8);
      const data = buildShiftScenarioData(start, end, baseHours, scenarioHours, assetCode, siteCodeFromReq(req));
      return reply.send({
        ok: true,
        start,
        end,
        asset_code: assetCode || null,
        assumptions: {
          fuel: "Fuel scales with logged L/hr.",
          nonfuel: "Recorded lube, parts, labour and downtime cost is held constant per active equipment shift.",
        },
        ...data,
      });
    } catch (error) {
      req.log.error(error);
      return reply.code(500).send({ error: error.message || String(error) });
    }
  });

  // GET /api/dashboard/shift-scenario.xlsx?start=&end=&base_hours=11&scenario_hours=8&asset_code=
  app.get("/shift-scenario.xlsx", async (req, reply) => {
    try {
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const assetCode = String(req.query?.asset_code || "").trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end) || start > end) {
        return reply.code(400).send({ error: "Provide valid start/end dates" });
      }
      const baseHours = scenarioHoursFromRequest(req.query?.base_hours, 11);
      const scenarioHours = scenarioHoursFromRequest(req.query?.scenario_hours, 8);
      const data = buildShiftScenarioData(start, end, baseHours, scenarioHours, assetCode, siteCodeFromReq(req));
      const workbook = new ExcelJS.Workbook();
      addShiftScenarioWorkbook(workbook, data, { start, end, assetCode });
      const buffer = await workbook.xlsx.writeBuffer();
      return reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Content-Disposition", `attachment; filename="IRONLOG_Shift_Scenario_${start}_to_${end}.xlsx"`)
        .send(buffer);
    } catch (error) {
      req.log.error(error);
      return reply.code(500).send({ error: error.message || String(error) });
    }
  });

  // GET /api/dashboard/fuel?start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15
  app.get("/fuel", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const toleranceInput = Number(req.query?.tolerance ?? 0.15);
    const tolerance = Number.isFinite(toleranceInput) ? Math.max(0, toleranceInput) : 0.15;
    const modeFilter = String(req.query?.mode || "").trim().toLowerCase(); // 'km' | 'hours' | ''
    const assetFilter = String(req.query?.asset_code || "").trim().toLowerCase();

    if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }

    const fuelByAsset = db.prepare(fuelBenchmarkAssetsInRangeSql()).all(start, end);
    const rows = sortBenchmarkRows(fuelByAsset
      .map((r) => {
        const row = assetConsumption(db, r, start, end, tolerance);
        return { ...row, note: consumptionNote(row) };
      })
      .filter((r) => r.fuel_liters > 0)
      .filter((r) => (assetFilter ? String(r.asset_code || "").trim().toLowerCase() === assetFilter : true))
      .filter((r) => (modeFilter === "km" ? r.metric_mode === "km" : modeFilter === "hours" ? r.metric_mode === "hours" : true)));

    const summary = summarizeFuelBenchmarkRows(rows);
    // Keep legacy key used by UI
    summary.excessive_count = summary.excessive;

    const category_rows = assetFilter
      ? []
      : aggregateFuelBenchmarkByCategory(rows, tolerance);

    return reply.send({
      ok: true,
      start,
      end,
      tolerance,
      mode: modeFilter || null,
      summary: {
        ...summary,
        categories: category_rows.length,
      },
      rows,
      category_rows,
    });
  });

  // GET /api/dashboard/fuel/duplicates?start=YYYY-MM-DD&end=YYYY-MM-DD&mode=km|hours
  app.get("/fuel/duplicates", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const modeFilter = String(req.query?.mode || "").trim().toLowerCase();
    const assetFilter = String(req.query?.asset_code || "").trim().toLowerCase();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }

    const rows = db.prepare(`
      WITH fuel_rows AS (
        SELECT
          fl.id,
          fl.asset_id,
          a.asset_code,
          a.asset_name,
          fl.log_date,
          COALESCE(fl.liters, 0) AS liters,
          COALESCE(LOWER(fl.meter_unit), '') AS meter_unit,
          COALESCE(fl.meter_run_value, -1) AS meter_run_value,
          COALESCE(fl.hours_run, -1) AS hours_run,
          COALESCE(TRIM(fl.source), '') AS source,
          COALESCE(NULLIF(TRIM(a.utilization_mode), ''), CASE
            WHEN LOWER(COALESCE(a.category, '')) LIKE '%truck%'
              OR LOWER(COALESCE(a.category, '')) LIKE '%vehicle%'
              OR LOWER(COALESCE(a.category, '')) LIKE '%ldv%'
              OR LOWER(COALESCE(a.category, '')) LIKE '%pickup%'
              OR LOWER(COALESCE(a.category, '')) LIKE '%bakkie%'
              OR LOWER(COALESCE(a.asset_code, '')) LIKE 'ldv%'
              OR UPPER(COALESCE(a.asset_code, '')) GLOB 'V[0-9][0-9]AM'
              OR LOWER(COALESCE(a.asset_name, '')) LIKE '%ldv%'
              THEN 'km'
            ELSE 'hours'
          END) AS metric_mode
        FROM fuel_logs fl
        JOIN assets a ON a.id = fl.asset_id
        WHERE fl.log_date BETWEEN ? AND ?
      ),
      tagged AS (
        SELECT
          fr.*,
          COUNT(*) OVER (
            PARTITION BY
              fr.asset_id,
              fr.log_date,
              ROUND(fr.liters, 3),
              fr.meter_unit,
              ROUND(fr.meter_run_value, 3),
              ROUND(fr.hours_run, 3),
              fr.source
          ) AS duplicate_count,
          ROW_NUMBER() OVER (
            PARTITION BY
              fr.asset_id,
              fr.log_date,
              ROUND(fr.liters, 3),
              fr.meter_unit,
              ROUND(fr.meter_run_value, 3),
              ROUND(fr.hours_run, 3),
              fr.source
            ORDER BY fr.id
          ) AS duplicate_rank
        FROM fuel_rows fr
      )
      SELECT
        id,
        asset_id,
        asset_code,
        asset_name,
        log_date,
        liters,
        meter_unit,
        meter_run_value,
        hours_run,
        source,
        CASE WHEN LOWER(COALESCE(metric_mode, 'hours')) = 'km' THEN 'km' ELSE 'hours' END AS metric_mode,
        duplicate_count,
        duplicate_rank
      FROM tagged
      WHERE duplicate_count > 1
        AND (
          ? = ''
          OR (? = 'km' AND LOWER(COALESCE(metric_mode, 'hours')) = 'km')
          OR (? = 'hours' AND LOWER(COALESCE(metric_mode, 'hours')) <> 'km')
        )
        AND (? = '' OR LOWER(COALESCE(asset_code, '')) = ?)
      ORDER BY log_date DESC, asset_code ASC, id ASC
    `).all(start, end, modeFilter, modeFilter, modeFilter, assetFilter, assetFilter);

    const summary = {
      duplicate_rows: rows.length,
      duplicate_groups: new Set(
        rows.map((r) => [
          r.asset_id,
          r.log_date,
          Number(r.liters || 0).toFixed(3),
          String(r.meter_unit || ""),
          Number(r.meter_run_value || 0).toFixed(3),
          Number(r.hours_run || 0).toFixed(3),
          String(r.source || ""),
        ].join("|"))
      ).size,
      fuel_liters: Number(rows.reduce((a, r) => a + Number(r.liters || 0), 0).toFixed(2)),
    };

    return reply.send({ ok: true, start, end, mode: modeFilter || null, summary, rows });
  });

  // GET /api/dashboard/fuel/daily?asset_code=A300AM&start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15
  // Returns one row per fuel fill entry; L/hr for hour assets, km/L for LDV assets.
  app.get("/fuel/daily", async (req, reply) => {
    const assetCode = String(req.query?.asset_code || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const toleranceInput = Number(req.query?.tolerance ?? 0.15);
    const tolerance = Number.isFinite(toleranceInput) ? Math.max(0, toleranceInput) : 0.15;

    if (!assetCode) return reply.code(400).send({ error: "asset_code is required" });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }

    const asset = db.prepare(`
      SELECT a.id AS asset_id, a.id, a.asset_code, a.asset_name, a.category, a.archived, a.hire_billing_mode,
        COALESCE(a.baseline_fuel_l_per_hour, 5.0) AS oem_lph,
        COALESCE(a.baseline_fuel_km_per_l, 2.0) AS oem_kmpl,
        a.fuel_benchmark_set_at,
        ${sqlFuelMetricModeExpr("a")} AS metric_mode
      FROM assets a
      WHERE a.asset_code = ?
    `).get(assetCode);
    if (!asset) return reply.code(404).send({ error: `asset not found: ${assetCode}` });

    // Same fill-to-fill calculation as the benchmark, one row per fill.
    const calc = assetConsumption(db, asset, start, end, tolerance, { withTrace: true });
    const mode = calc.metric_mode;
    let lastGood = null;
    const rows = calc.trace.map((t) => {
      const ok = t.status === "ok";
      const run = ok ? Number(t.run || 0) : 0;
      const lph = ok && mode === "hours" && run > 0 ? t.interval_liters / run : null;
      const kmpl = ok && mode === "km" && run > 0 && t.interval_liters > 0 ? run / t.interval_liters : null;
      const isExcessive = mode === "km"
        ? kmpl != null && calc.low_threshold_km_per_l != null && kmpl < calc.low_threshold_km_per_l
        : lph != null && calc.excessive_threshold_lph != null && lph > calc.excessive_threshold_lph;
      const open = ok && t.reading != null ? t.reading - run : lastGood;
      if (ok && t.reading != null) lastGood = t.reading;
      return {
        id: Number(t.id),
        log_date: t.date,
        metric_mode: mode,
        fuel_liters: Number(Number(t.liters || 0).toFixed(2)),
        interval_liters: t.interval_liters == null ? null : Number(Number(t.interval_liters).toFixed(2)),
        run_value: Number(run.toFixed(2)),
        run_unit: mode === "km" ? "km" : "hours",
        hours_run: mode === "hours" ? Number(run.toFixed(2)) : 0,
        km_run: mode === "km" ? Number(run.toFixed(2)) : 0,
        meter_value: t.reading,
        open_meter_value: open == null ? null : Number(Number(open).toFixed(2)),
        close_meter_value: t.reading,
        meter_unit_display: mode === "km" ? "km" : "hours",
        invalid_delta: /reading ignored|meter went back|counted from this reading/.test(t.status),
        status: ok ? "" : t.status,
        actual_lph: lph == null ? null : Number(lph.toFixed(3)),
        oem_lph: calc.oem_lph,
        excessive_threshold_lph: calc.excessive_threshold_lph,
        actual_km_per_l: kmpl == null ? null : Number(kmpl.toFixed(3)),
        oem_km_per_l: calc.oem_km_per_l,
        low_threshold_km_per_l: calc.low_threshold_km_per_l,
        is_excessive: isExcessive,
        source: t.source || null,
      };
    });

    const summary = {
      days: rows.length,
      fuel_liters: calc.fuel_liters,
      matched_liters: calc.matched_liters,
      coverage_pct: calc.coverage_pct,
      hours_run: calc.hours_run,
      km_run: calc.km_run,
      excessive_days: rows.filter((r) => r.is_excessive).length,
      rejected_readings: calc.suspect_readings,
      run_source: calc.run_source,
      note: consumptionNote(calc),
      oem_lph: calc.oem_lph,
      oem_km_per_l: calc.oem_km_per_l,
      metric_mode: mode,
      avg_lph: calc.actual_lph,
      avg_km_per_l: calc.actual_km_per_l,
    };

    return reply.send({
      ok: true,
      asset_code: asset.asset_code,
      asset_name: asset.asset_name,
      start,
      end,
      tolerance,
      summary,
      rows,
    });
  });

  // DELETE /api/dashboard/fuel/log/:id
  app.delete("/fuel/log/:id", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isInteger(id) || id <= 0) {
      return reply.code(400).send({ error: "valid fuel log id is required" });
    }

    const row = db.prepare(`
      SELECT fl.id, fl.asset_id, fl.log_date, fl.liters, fl.hours_run, fl.source, a.asset_code
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.id = ?
    `).get(id);
    if (!row) return reply.code(404).send({ error: "fuel log not found" });

    db.prepare(`DELETE FROM fuel_logs WHERE id = ?`).run(id);

    writeAudit(db, req, {
      module: "fuel",
      action: "delete_log",
      entity_type: "fuel_log",
      entity_id: String(id),
      payload: {
        asset_code: row.asset_code,
        log_date: row.log_date,
        liters: row.liters,
        hours_run: row.hours_run,
        source: row.source,
      },
    });

    return reply.send({ ok: true, deleted_id: id, asset_code: row.asset_code });
  });
}
