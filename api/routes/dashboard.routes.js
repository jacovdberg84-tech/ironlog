// IRONLOG/api/routes/dashboard.routes.js
import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";

function todayYYYYMMDD() {
  return new Date().toISOString().slice(0, 10);
}

function monthRangeFromYYYYMM(monthStr) {
  const [y, m] = String(monthStr).split("-").map((n) => Number(n));
  const start = new Date(Date.UTC(y, m - 1, 1));
  const end = new Date(Date.UTC(y, m, 0));
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

function monthIdFromDateStr(dateStr) {
  return String(dateStr || "").slice(0, 7);
}

/** First day of month for a YYYY-MM-DD date string. */
function monthStartIso(dateStr) {
  const s = String(dateStr || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return `${s.slice(0, 7)}-01`;
}

function isToyotaHiluxAsset(asset) {
  const code = String(asset?.asset_code || "").toLowerCase();
  const name = String(asset?.asset_name || "").toLowerCase();
  return name.includes("toyota") && name.includes("hilux") || code.includes("hilux");
}

function eachDateInclusiveYMD(startStr, endStr, fn) {
  const start = new Date(`${startStr}T12:00:00`);
  const end = new Date(`${endStr}T12:00:00`);
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    fn(d.toISOString().slice(0, 10));
  }
}

export default async function dashboardRoutes(app) {
  ensureAuditTable(db);

  function hasColumn(table, col) {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name) === col);
  }
  function getBreakdownDowntimeColumn() {
    const rows = db.prepare(`PRAGMA table_info(breakdowns)`).all();
    const names = new Set(rows.map((r) => String(r.name)));
    if (names.has("downtime_total_hours")) return "downtime_total_hours";
    if (names.has("downtime_hours")) return "downtime_hours";
    return "downtime_hours";
  }

  function ensureColumn(table, colName, colDef) {
    if (!hasColumn(table, colName)) {
      db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
    }
  }

  function getRole(req) {
    return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
  }

  function requireRoles(req, reply, roles) {
    const role = getRole(req);
    if (!roles.includes(role)) {
      reply.code(403).send({ error: `role '${role || "unknown"}' not allowed` });
      return false;
    }
    return true;
  }

  db.prepare(`
    CREATE TABLE IF NOT EXISTS lube_type_mappings (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      oil_key TEXT NOT NULL UNIQUE,
      part_code TEXT NOT NULL,
      updated_by TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  function slaPriority(status, ageHours) {
    const s = String(status || "").toLowerCase();
    const a = Number(ageHours || 0);
    if (s === "completed" && a > 48) return "P1";
    if (s === "in_progress" && a > 72) return "P1";
    if ((s === "open" || s === "assigned") && a > 72) return "P1";

    if (s === "completed" && a > 24) return "P2";
    if (s === "in_progress" && a > 48) return "P2";
    if ((s === "open" || s === "assigned") && a > 48) return "P2";

    return "P3";
  }

  // OEM baseline fuel benchmark (L/hr) per asset.
  ensureColumn("assets", "baseline_fuel_l_per_hour", "baseline_fuel_l_per_hour REAL DEFAULT 5.0");
  ensureColumn("assets", "baseline_fuel_km_per_l", "baseline_fuel_km_per_l REAL DEFAULT 2.0");
  ensureColumn("assets", "fuel_cost_per_liter", "fuel_cost_per_liter REAL");
  ensureColumn("assets", "downtime_cost_per_hour", "downtime_cost_per_hour REAL");
  ensureColumn("daily_hours", "input_unit", "input_unit TEXT DEFAULT 'hours'");
  ensureColumn("assets", "utilization_mode", "utilization_mode TEXT DEFAULT 'hours'");
  ensureColumn("assets", "km_per_hour_factor", "km_per_hour_factor REAL DEFAULT 10.0");
  ensureColumn("parts", "unit_cost", "unit_cost REAL DEFAULT 0");
  ensureColumn("oil_logs", "unit_cost", "unit_cost REAL");
  ensureColumn("fuel_logs", "unit_cost_per_liter", "unit_cost_per_liter REAL");
  ensureColumn("fuel_logs", "hours_run", "hours_run REAL");
  ensureColumn("fuel_logs", "meter_run_value", "meter_run_value REAL");
  ensureColumn("fuel_logs", "meter_unit", "meter_unit TEXT");
  ensureColumn("fuel_logs", "open_meter_value", "open_meter_value REAL");
  ensureColumn("fuel_logs", "close_meter_value", "close_meter_value REAL");
  ensureColumn("work_orders", "labor_hours", "labor_hours REAL DEFAULT 0");
  ensureColumn("work_orders", "labor_rate_per_hour", "labor_rate_per_hour REAL");

  db.prepare(`
    CREATE TABLE IF NOT EXISTS cost_settings (
      key TEXT PRIMARY KEY,
      value REAL NOT NULL DEFAULT 0,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  const upsertCostSetting = db.prepare(`
    INSERT INTO cost_settings (key, value, updated_at)
    VALUES (?, ?, datetime('now'))
    ON CONFLICT(key) DO NOTHING
  `);
  upsertCostSetting.run("fuel_cost_per_liter_default", 1.5);
  upsertCostSetting.run("lube_cost_per_qty_default", 4.0);
  upsertCostSetting.run("labor_cost_per_hour_default", 35.0);
  upsertCostSetting.run("downtime_cost_per_hour_default", 120.0);

  // -----------------------------
  // Prepared statements (reuse)
  // -----------------------------

  // Per-asset production rows for the day (standby excluded)
  const getDayAssetHours = db.prepare(`
    SELECT
      dh.asset_id,
      a.asset_code,
      a.asset_name,
      a.category,
      COALESCE(NULLIF(TRIM(dh.input_unit), ''), '') AS input_unit,
      CASE
        WHEN (
          (INSTR(LOWER(COALESCE(a.asset_name, '')), 'toyota') > 0 AND INSTR(LOWER(COALESCE(a.asset_name, '')), 'hilux') > 0)
          OR INSTR(LOWER(COALESCE(a.asset_code, '')), 'hilux') > 0
        ) THEN 'km'
        ELSE 'hours'
      END AS utilization_mode,
      COALESCE(NULLIF(a.km_per_hour_factor, 0), 10.0) AS km_per_hour_factor,
      COALESCE(dh.scheduled_hours, 0) AS scheduled_hours,
      COALESCE(dh.hours_run, 0) AS run_hours
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date = ?
      AND dh.is_used = 1
      AND a.active = 1
      AND a.is_standby = 0
  `);
  const getActiveFleetAssets = db.prepare(`
    SELECT id AS asset_id, asset_code, asset_name, category, utilization_mode, km_per_hour_factor
    FROM assets
    WHERE active = 1
      AND is_standby = 0
  `);

  // Per-asset downtime for the day from downtime logs (standby excluded)
  const getDayAssetDowntime = db.prepare(`
    SELECT
      b.asset_id,
      COALESCE(SUM(l.hours_down), 0) AS downtime_hours
    FROM breakdown_downtime_logs l
    JOIN breakdowns b ON b.id = l.breakdown_id
    JOIN assets a ON a.id = b.asset_id
    WHERE l.log_date = ?
      AND a.active = 1
      AND a.is_standby = 0
    GROUP BY b.asset_id
  `);
  const getOpenBreakdownAssetIdsByDay = db.prepare(`
    SELECT DISTINCT b.asset_id
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    WHERE b.status = 'OPEN'
      AND b.breakdown_date <= ?
      AND a.active = 1
      AND a.is_standby = 0
  `);
  const breakdownDowntimeCol = getBreakdownDowntimeColumn();
  const getDayAssetDowntimeFallback = db.prepare(`
    SELECT
      b.asset_id,
      COALESCE(SUM(COALESCE(b.${breakdownDowntimeCol}, 0)), 0) AS downtime_hours
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    WHERE b.breakdown_date = ?
      AND a.active = 1
      AND a.is_standby = 0
    GROUP BY b.asset_id
  `);

  // Downtime reasons summary (from downtime log notes)
  const getDowntimeReasons = db.prepare(`
    SELECT
      CASE
        WHEN l.notes IS NULL OR TRIM(l.notes) = '' THEN 'Unspecified'
        WHEN INSTR(l.notes, '—') > 0 THEN TRIM(SUBSTR(l.notes, INSTR(l.notes, '—') + 1))
        ELSE 'Unspecified'
      END AS reason,
      COALESCE(SUM(l.hours_down), 0) AS hours_down,
      COUNT(DISTINCT l.breakdown_id) AS incidents
    FROM breakdown_downtime_logs l
    WHERE l.log_date = ?
    GROUP BY reason
    ORDER BY hours_down DESC, incidents DESC
    LIMIT 8
  `);

  /**
   * Fleet KPI for one calendar day (same rules as dashboard split KPI).
   * When includePerAsset=true, builds per_asset_kpi for that day (debug / detail).
   */
  function computeFleetKpiForDay(dayStr, scheduledFallback, opts = {}) {
    const includePerAsset = Boolean(opts.includePerAsset);
    const assetRows = getDayAssetHours.all(dayStr);
    const logDowntimeRows = getDayAssetDowntime.all(dayStr);
    const downtimeRows = logDowntimeRows.length
      ? logDowntimeRows
      : getDayAssetDowntimeFallback.all(dayStr);
    const downtimeByAsset = new Map(
      downtimeRows.map((r) => [Number(r.asset_id || 0), Number(r.downtime_hours || 0)])
    );
    const openBreakdownAssets = new Set(
      getOpenBreakdownAssetIdsByDay.all(dayStr).map((r) => Number(r.asset_id || 0))
    );
    const assetIdsInHours = new Set(assetRows.map((r) => Number(r.asset_id || 0)));

    let scheduled_hours = 0;
    let run_hours = 0;
    let downtime_hours = 0;
    let utilization_base_hours = 0;
    const per_asset_kpi = includePerAsset ? [] : null;
    const contributingAssetIds = new Set();

    assetRows.forEach((r) => {
      const assetId = Number(r.asset_id || 0);
      const rowScheduled = Number(r.scheduled_hours);
      const scheduled = Math.max(
        0,
        Number.isFinite(rowScheduled) && rowScheduled > 0
          ? rowScheduled
          : Number(scheduledFallback || 0)
      );
      const runRaw = Math.max(0, Number(r.run_hours || 0));
      const mode = isToyotaHiluxAsset(r) ? "km" : "hours";
      const kmPerHour = Math.max(0.1, Number(r.km_per_hour_factor || 10));
      const run = mode === "km" ? (runRaw / kmPerHour) : runRaw;
      const loggedDownRaw = Math.max(0, Number(downtimeByAsset.get(assetId) || 0));
      const loggedDown = loggedDownRaw > 0
        ? loggedDownRaw
        : (openBreakdownAssets.has(assetId) ? scheduled : 0);
      const cappedDown = Math.min(loggedDown, scheduled);
      const contributes_to_kpi = true;
      const runEff = Math.min(run, scheduled);
      scheduled_hours += scheduled;
      run_hours += runEff;
      downtime_hours += cappedDown;
      utilization_base_hours += scheduled;
      const s = Number(scheduled.toFixed(2));
      const runN = Number(runEff.toFixed(2));
      const downN = Number(cappedDown.toFixed(2));
      if (s > 0 || runN > 0 || downN > 0) contributingAssetIds.add(assetId);
      if (includePerAsset) {
        per_asset_kpi.push({
          asset_id: assetId,
          asset_code: String(r.asset_code || ""),
          asset_name: String(r.asset_name || ""),
          category: String(r.category || ""),
          utilization_mode: mode,
          contributes_to_kpi,
          km_per_hour_factor: Number(kmPerHour.toFixed(2)),
          scheduled_hours: s,
          meter_run_value: Number(runRaw.toFixed(2)),
          run_hours: runN,
          downtime_hours: downN,
          available_hours: Number(Math.max(0, scheduled - cappedDown).toFixed(2)),
        });
      }
    });

    const activeFleetIds = getActiveFleetAssets.all().map((r) => Number(r.asset_id || 0));
    const missingFleetIds = activeFleetIds.filter((id) => id > 0 && !assetIdsInHours.has(id));
    const downtimeOnlyIds = downtimeRows
      .map((r) => Number(r.asset_id || 0))
      .filter((id) => id > 0 && !assetIdsInHours.has(id));
    const includeIds = Array.from(new Set([...downtimeOnlyIds, ...missingFleetIds]));
    if (includeIds.length) {
      const uniq = includeIds.slice(0, 500);
      const placeholders = uniq.map(() => "?").join(",");
      const extraAssets = db.prepare(`
        SELECT id AS asset_id, asset_code, asset_name, category, utilization_mode, km_per_hour_factor
        FROM assets
        WHERE id IN (${placeholders})
      `).all(...uniq);

      for (const a of extraAssets) {
        const assetId = Number(a.asset_id || 0);
        const scheduled = Math.max(0, Number(scheduledFallback || 0));
        const loggedDownRaw = Math.max(0, Number(downtimeByAsset.get(assetId) || 0));
        const loggedDown = loggedDownRaw > 0
          ? loggedDownRaw
          : (openBreakdownAssets.has(assetId) ? scheduled : 0);
        const cappedDown = Math.min(loggedDown, scheduled);
        const cat = String(a.category || "");
        const mode = isToyotaHiluxAsset(a) ? "km" : "hours";
        const kmPerHour = Math.max(0.1, Number(a.km_per_hour_factor || 10));
        const runRaw = 0;
        const run = mode === "km" ? (runRaw / kmPerHour) : runRaw;
        const runEff = Math.min(run, scheduled);

        scheduled_hours += scheduled;
        run_hours += runEff;
        downtime_hours += cappedDown;

        const s = Number(scheduled.toFixed(2));
        const runN = Number(runEff.toFixed(2));
        const downN = Number(cappedDown.toFixed(2));
        if (s > 0 || runN > 0 || downN > 0) contributingAssetIds.add(assetId);
        if (includePerAsset) {
          per_asset_kpi.push({
            asset_id: assetId,
            asset_code: String(a.asset_code || ""),
            asset_name: String(a.asset_name || ""),
            category: cat,
            utilization_mode: mode,
            contributes_to_kpi: true,
            km_per_hour_factor: Number(kmPerHour.toFixed(2)),
            scheduled_hours: s,
            meter_run_value: Number(runRaw.toFixed(2)),
            run_hours: runN,
            downtime_hours: downN,
            available_hours: Number(Math.max(0, scheduled - cappedDown).toFixed(2)),
          });
        }
      }
    }

    return {
      scheduled_hours,
      run_hours,
      downtime_hours,
      utilization_base_hours,
      contributingAssetIds,
      per_asset_kpi: includePerAsset ? per_asset_kpi : [],
    };
  }

  // GET /api/dashboard/asset-kpi/weekly?start=YYYY-MM-DD&end=YYYY-MM-DD&scheduled=10
  // Rolls up dashboard KPI rules (availability / utilization) across a date range per asset and by equipment category.
  app.get("/asset-kpi/weekly", async (req, reply) => {
    const startIn = String(req.query?.start || "").trim();
    const endIn = String(req.query?.end || "").trim();
    const scheduledRaw = Number(req.query?.scheduled ?? 10);
    const scheduledFallback = Number.isFinite(scheduledRaw) && scheduledRaw > 0 ? scheduledRaw : 10;

    const now = new Date();
    const dow = now.getDay();
    const mondayOffset = dow === 0 ? -6 : 1 - dow;
    const monday = new Date(now);
    monday.setHours(0, 0, 0, 0);
    monday.setDate(monday.getDate() + mondayOffset);
    const friday = new Date(monday);
    friday.setDate(friday.getDate() + 4);
    const ymd = (d) => d.toISOString().slice(0, 10);
    const start = /^\d{4}-\d{2}-\d{2}$/.test(startIn) ? startIn : ymd(monday);
    const end = /^\d{4}-\d{2}-\d{2}$/.test(endIn) ? endIn : ymd(friday);
    if (start > end) {
      return reply.code(400).send({ ok: false, error: "start must be <= end" });
    }

    const assetMap = new Map();
    let daysInRange = 0;
    eachDateInclusiveYMD(start, end, (dayStr) => {
      daysInRange += 1;
      const k = computeFleetKpiForDay(dayStr, scheduledFallback, { includePerAsset: true });
      for (const row of k.per_asset_kpi || []) {
        const id = Number(row.asset_id || 0);
        if (!id) continue;
        if (!assetMap.has(id)) {
          assetMap.set(id, {
            asset_id: id,
            asset_code: String(row.asset_code || ""),
            asset_name: String(row.asset_name || ""),
            category: String(row.category || ""),
            utilization_mode: String(row.utilization_mode || "hours"),
            scheduled_hours: 0,
            run_hours: 0,
            downtime_hours: 0,
            available_hours: 0,
            days_with_data: 0,
          });
        }
        const a = assetMap.get(id);
        const s = Number(row.scheduled_hours || 0);
        const r = Number(row.run_hours || 0);
        const d = Number(row.downtime_hours || 0);
        const v = Number(row.available_hours || 0);
        a.scheduled_hours += s;
        a.run_hours += r;
        a.downtime_hours += d;
        a.available_hours += v;
        if (s > 0 || r > 0 || d > 0) a.days_with_data += 1;
      }
    });

    const pct = (num, den) =>
      den > 0 && Number.isFinite(num) ? Number(((num / den) * 100).toFixed(1)) : null;

    const by_asset = Array.from(assetMap.values()).map((a) => {
      const sched = a.scheduled_hours;
      const avail = a.available_hours;
      const run = a.run_hours;
      return {
        asset_id: a.asset_id,
        asset_code: a.asset_code,
        asset_name: a.asset_name,
        category: a.category,
        utilization_mode: a.utilization_mode,
        days_with_data: a.days_with_data,
        days_in_range: daysInRange,
        scheduled_hours: Number(sched.toFixed(2)),
        run_hours: Number(run.toFixed(2)),
        downtime_hours: Number(a.downtime_hours.toFixed(2)),
        available_hours: Number(avail.toFixed(2)),
        availability_pct: pct(avail, sched),
        utilization_pct: pct(run, avail),
      };
    });

    by_asset.sort((x, y) => {
      if (x.utilization_pct == null && y.utilization_pct == null) {
        return String(x.asset_code || "").localeCompare(String(y.asset_code || ""));
      }
      if (x.utilization_pct == null) return 1;
      if (y.utilization_pct == null) return -1;
      return y.utilization_pct - x.utilization_pct;
    });

    const catMap = new Map();
    for (const a of by_asset) {
      const catKey = String(a.category || "").trim() || "Uncategorized";
      if (!catMap.has(catKey)) {
        catMap.set(catKey, {
          category: catKey,
          scheduled_hours: 0,
          run_hours: 0,
          downtime_hours: 0,
          available_hours: 0,
          asset_ids: new Set(),
        });
      }
      const c = catMap.get(catKey);
      c.scheduled_hours += a.scheduled_hours;
      c.run_hours += a.run_hours;
      c.downtime_hours += a.downtime_hours;
      c.available_hours += a.available_hours;
      c.asset_ids.add(a.asset_id);
    }

    const by_category = Array.from(catMap.values()).map((c) => {
      const sched = c.scheduled_hours;
      const avail = c.available_hours;
      const run = c.run_hours;
      return {
        category: c.category,
        asset_count: c.asset_ids.size,
        scheduled_hours: Number(sched.toFixed(2)),
        run_hours: Number(run.toFixed(2)),
        downtime_hours: Number(c.downtime_hours.toFixed(2)),
        available_hours: Number(avail.toFixed(2)),
        availability_pct: pct(avail, sched),
        utilization_pct: pct(run, avail),
      };
    });

    by_category.sort((x, y) => {
      if (x.utilization_pct == null && y.utilization_pct == null) {
        return String(x.category || "").localeCompare(String(y.category || ""));
      }
      if (x.utilization_pct == null) return 1;
      if (y.utilization_pct == null) return -1;
      return y.utilization_pct - x.utilization_pct;
    });

    const fleet_sched = by_asset.reduce((s, r) => s + r.scheduled_hours, 0);
    const fleet_avail = by_asset.reduce((s, r) => s + r.available_hours, 0);
    const fleet_run = by_asset.reduce((s, r) => s + r.run_hours, 0);
    const fleet_down = by_asset.reduce((s, r) => s + r.downtime_hours, 0);

    return reply.send({
      ok: true,
      range: { start, end },
      scheduled_fallback: scheduledFallback,
      days_in_range: daysInRange,
      definitions: {
        availability_pct: "(sum of available hours) / (sum of scheduled hours) × 100; available = scheduled − downtime (capped per day).",
        utilization_pct: "(sum of effective run hours) / (sum of available hours) × 100 (matches main dashboard KPI).",
      },
      fleet: {
        scheduled_hours: Number(fleet_sched.toFixed(2)),
        available_hours: Number(fleet_avail.toFixed(2)),
        run_hours: Number(fleet_run.toFixed(2)),
        downtime_hours: Number(fleet_down.toFixed(2)),
        availability_pct: pct(fleet_avail, fleet_sched),
        utilization_pct: pct(fleet_run, fleet_avail),
      },
      by_category,
      by_asset,
    });
  });

  // GET /api/dashboard?date=YYYY-MM-DD&scheduled=10
  app.get("/", async (req, reply) => {
    const date = String(req.query?.date || todayYYYYMMDD()).trim();
    const scheduledRaw = Number(req.query?.scheduled ?? 10);
    const scheduledFallback = Number.isFinite(scheduledRaw) && scheduledRaw > 0
      ? scheduledRaw
      : 10;

    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
    }

    // =========================
    // KPI (Split) — gauges: month-to-date through selected date; per-asset table: selected day
    // Availability = (Scheduled - Downtime) / Scheduled
    // Utilization  = Run / (Scheduled - Downtime)
    // =========================

    const dayK = computeFleetKpiForDay(date, scheduledFallback, { includePerAsset: true });
    const per_asset_kpi = dayK.per_asset_kpi;
    const run_hours = dayK.run_hours;

    const mtdStart = monthStartIso(date);
    let mtd_scheduled = 0;
    let mtd_run = 0;
    let mtd_downtime = 0;
    let mtd_utilization_base = 0;
    let mtd_day_count = 0;
    const mtdAssetIds = new Set();
    eachDateInclusiveYMD(mtdStart, date, (dayStr) => {
      mtd_day_count += 1;
      const dr = computeFleetKpiForDay(dayStr, scheduledFallback, { includePerAsset: false });
      mtd_scheduled += dr.scheduled_hours;
      mtd_run += dr.run_hours;
      mtd_downtime += dr.downtime_hours;
      mtd_utilization_base += Number(dr.utilization_base_hours || 0);
      dr.contributingAssetIds.forEach((id) => mtdAssetIds.add(id));
    });

    // Safety fallback: if MTD scheduled remained zero, derive from active fleet.
    if (mtd_scheduled <= 0 && scheduledFallback > 0 && mtd_day_count > 0) {
      const activeFleetCountRow = db.prepare(`
        SELECT COUNT(*) AS c
        FROM assets
        WHERE active = 1
          AND is_standby = 0
      `).get();
      const activeFleetCount = Number(activeFleetCountRow?.c || 0);
      if (activeFleetCount > 0) {
        mtd_scheduled = activeFleetCount * scheduledFallback * mtd_day_count;
      }
    }

    const available_hours = Math.max(0, mtd_scheduled - mtd_downtime);
    const availability =
      mtd_scheduled > 0 ? (available_hours / mtd_scheduled) * 100 : null;
    const utilization =
      mtd_utilization_base > 0 ? (mtd_run / mtd_utilization_base) * 100 : null;
    const used_assets = mtdAssetIds.size;

    const scheduled_hours = mtd_scheduled;
    const downtime_hours = mtd_downtime;
    // run_hours = selected day only (for cost-per-run-hour); gauges use mtd_* above

    // Alerts summaries
    const lowStockCount = db.prepare(`
      SELECT COUNT(*) AS c
      FROM (
        SELECT p.id, IFNULL(SUM(sm.quantity),0) AS on_hand, p.min_stock
        FROM parts p
        LEFT JOIN stock_movements sm ON sm.part_id = p.id
        GROUP BY p.id
        HAVING on_hand < p.min_stock
      )
    `).get();

    const overdueMaintCount = db.prepare(`
      SELECT COUNT(*) AS c
      FROM (
        SELECT
          mp.id,
          (IFNULL((
            SELECT SUM(dh.hours_run)
            FROM daily_hours dh
            WHERE dh.asset_id = mp.asset_id
              AND dh.is_used = 1
              AND dh.hours_run > 0
              AND dh.work_date <= ?
          ),0) - (mp.last_service_hours + mp.interval_hours)) AS diff
        FROM maintenance_plans mp
        JOIN assets a ON a.id = mp.asset_id
        WHERE mp.active = 1 AND a.active = 1 AND a.is_standby = 0
      )
      WHERE diff >= 0
    `).get(date);

    const hasWOCompletedAt = hasColumn("work_orders", "completed_at");
    const hasBreakdownStatus = hasColumn("breakdowns", "status");
    const woCompletedFilter = hasWOCompletedAt
      ? "AND (w.completed_at IS NULL OR TRIM(COALESCE(w.completed_at, '')) = '')"
      : "";
    const breakdownOpenFilter = hasBreakdownStatus
      ? `AND (
          w.source <> 'breakdown'
          OR TRIM(LOWER(COALESCE(b.status, ''))) IN ('open', 'in_progress')
        )`
      : "";

    const openWOCount = db.prepare(`
      SELECT COUNT(*) AS c
      FROM work_orders w
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
      WHERE REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
        AND (w.closed_at IS NULL OR TRIM(COALESCE(w.closed_at, '')) = '')
        ${woCompletedFilter}
        ${breakdownOpenFilter}
    `).get();

    // Major downtime list (use downtime logs aggregated per breakdown for this day)
    const majorDowntime = db.prepare(`
      SELECT
        a.asset_code,
        b.description,
        SUM(l.hours_down) AS downtime_hours,
        b.critical
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date = ?
      GROUP BY l.breakdown_id
      ORDER BY downtime_hours DESC
      LIMIT 5
    `).all(date).map(r => ({
      ...r,
      downtime_hours: Number(r.downtime_hours || 0),
      critical: Boolean(r.critical)
    }));

    const downtimeReasons = getDowntimeReasons.all(date).map(r => ({
      reason: r.reason,
      hours_down: Number(r.hours_down || 0),
      incidents: Number(r.incidents || 0)
    }));

    const criticalLowStock = db.prepare(`
      SELECT p.part_code, p.part_name, p.min_stock, IFNULL(SUM(sm.quantity),0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      WHERE p.critical = 1
      GROUP BY p.id
      HAVING on_hand < p.min_stock
      ORDER BY on_hand ASC
      LIMIT 8
    `).all().map(r => ({ ...r, on_hand: Number(r.on_hand) }));

    const openWOs = db.prepare(`
      SELECT
        w.id,
        a.asset_code,
        w.source,
        w.status,
        CASE
          WHEN w.source = 'breakdown' THEN COALESCE(NULLIF(TRIM(b.start_at), ''), NULLIF(TRIM(b.breakdown_date), ''), w.opened_at)
          ELSE w.opened_at
        END AS opened_at
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
      WHERE REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
        AND (w.closed_at IS NULL OR TRIM(COALESCE(w.closed_at, '')) = '')
        ${woCompletedFilter}
        ${breakdownOpenFilter}
      ORDER BY w.id DESC
      LIMIT 8
    `).all();

    const openWOSla = db.prepare(`
      SELECT
        w.id,
        a.asset_code,
        w.source,
        w.status,
        CASE
          WHEN w.source = 'breakdown' THEN COALESCE(NULLIF(TRIM(b.start_at), ''), NULLIF(TRIM(b.breakdown_date), ''), w.opened_at)
          ELSE w.opened_at
        END AS opened_at,
        CAST((
          julianday('now') - julianday(
            COALESCE(
              CASE
                WHEN w.source = 'breakdown' THEN COALESCE(NULLIF(TRIM(b.start_at), ''), NULLIF(TRIM(b.breakdown_date), ''), w.opened_at)
                ELSE w.opened_at
              END,
              datetime('now')
            )
          )
        ) * 24 AS INTEGER) AS age_hours
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
      WHERE REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
        AND (w.closed_at IS NULL OR TRIM(COALESCE(w.closed_at, '')) = '')
        ${woCompletedFilter}
        ${breakdownOpenFilter}
      ORDER BY age_hours DESC, w.id DESC
      LIMIT 200
    `).all().map((r) => ({
      id: Number(r.id),
      asset_code: r.asset_code,
      source: r.source,
      status: r.status,
      opened_at: r.opened_at,
      age_hours: Number(r.age_hours || 0),
    }));

    const sla_summary = {
      open_gt_24h: openWOSla.filter((r) => r.age_hours > 24).length,
      in_progress_gt_48h: openWOSla.filter((r) => String(r.status) === "in_progress" && r.age_hours > 48).length,
      completed_gt_12h: 0,
    };
    const sla_breaches = openWOSla
      .filter((r) => {
        const s = String(r.status || "").toLowerCase();
        if (s === "in_progress") return r.age_hours > 48;
        return r.age_hours > 24;
      })
      .map((r) => ({
        ...r,
        priority: slaPriority(r.status, r.age_hours),
      }))
      .slice(0, 8);

    const lubeDaily = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ol.quantity), 0) AS qty,
        COALESCE(SUM(
          ol.quantity * COALESCE(
            ol.unit_cost,
            (SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1),
            4.0
          )
        ), 0) AS lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date = ?
      GROUP BY a.id
      ORDER BY qty DESC, a.asset_code ASC
      LIMIT 8
    `).all(date).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      qty: Number(r.qty || 0),
      lube_cost: Number(r.lube_cost || 0),
    }));

    const lubeDailyByType = db.prepare(`
      SELECT
        a.asset_code,
        CASE
          WHEN LOWER(TRIM(COALESCE(ol.oil_type, ''))) IN ('admin','supervisor','manager','stores','artisan','operator') THEN 'UNSPECIFIED'
          ELSE COALESCE(NULLIF(TRIM(ol.oil_type), ''), 'UNSPECIFIED')
        END AS oil_type,
        COALESCE(SUM(ol.quantity), 0) AS qty,
        COALESCE(SUM(
          ol.quantity * COALESCE(
            ol.unit_cost,
            (SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1),
            4.0
          )
        ), 0) AS lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date = ?
      GROUP BY a.asset_code, oil_type
      ORDER BY a.asset_code ASC, qty DESC, oil_type ASC
      LIMIT 400
    `).all(date);
    const byTypeMap = new Map();
    for (const r of lubeDailyByType) {
      const code = String(r.asset_code || "");
      if (!byTypeMap.has(code)) byTypeMap.set(code, []);
      byTypeMap.get(code).push({
        oil_type: String(r.oil_type || "UNSPECIFIED"),
        qty: Number(r.qty || 0),
        lube_cost: Number(r.lube_cost || 0),
      });
    }
    for (const row of lubeDaily) {
      row.by_oil_type = byTypeMap.get(String(row.asset_code || "")) || [];
    }

    const lubeTotalRow = db.prepare(`
      SELECT
        COALESCE(SUM(quantity), 0) AS qty_total,
        COALESCE(SUM(
          quantity * COALESCE(
            unit_cost,
            (SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1),
            4.0
          )
        ), 0) AS total_lube_cost
      FROM oil_logs
      WHERE log_date = ?
    `).get(date);

    const settingsRows = db.prepare(`
      SELECT key, value
      FROM cost_settings
      WHERE key IN (
        'fuel_cost_per_liter_default',
        'lube_cost_per_qty_default',
        'labor_cost_per_hour_default',
        'downtime_cost_per_hour_default'
      )
    `).all();
    const settings = {
      fuel_cost_per_liter_default: 1.5,
      lube_cost_per_qty_default: 4.0,
      labor_cost_per_hour_default: 35.0,
      downtime_cost_per_hour_default: 120.0,
    };
    for (const r of settingsRows) {
      const k = String(r.key || "").trim();
      const v = Number(r.value);
      if (k && Number.isFinite(v)) settings[k] = v;
    }

    const fuelCostRow = db.prepare(`
      SELECT COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS fuel_cost
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date = ?
    `).get(settings.fuel_cost_per_liter_default, date);

    const lubeCostRow = db.prepare(`
      SELECT COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS lube_cost
      FROM oil_logs ol
      WHERE ol.log_date = ?
    `).get(settings.lube_cost_per_qty_default, date);

    const woCostRow = db.prepare(`
      SELECT
        COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS labor_hours,
        COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      WHERE DATE(COALESCE(w.completed_at, w.closed_at)) = ?
        AND w.status IN ('completed', 'approved', 'closed')
    `).get(settings.labor_cost_per_hour_default, date);

    const downtimeCostRow = db.prepare(`
      SELECT COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS downtime_cost
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date = ?
    `).get(settings.downtime_cost_per_hour_default, date);

    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
    const partsCostRows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS parts_cost
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      LEFT JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
      LEFT JOIN assets a ON a.id = w.asset_id
      WHERE sm.movement_type = 'out'
        AND ${smDateExpr} = ?
      GROUP BY a.id
    `).all(date).map((r) => ({
      asset_code: r.asset_code || "UNLINKED",
      asset_name: r.asset_name || "Unlinked to WO",
      parts_cost: Number(r.parts_cost || 0),
    }));
    const parts_cost = Number(partsCostRows.reduce((acc, r) => acc + Number(r.parts_cost || 0), 0).toFixed(2));

    const assetFuelCost = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS fuel_cost
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date = ?
      GROUP BY a.id
    `).all(settings.fuel_cost_per_liter_default, date);

    const assetLubeCost = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date = ?
      GROUP BY a.id
    `).all(settings.lube_cost_per_qty_default, date);

    const assetLaborCost = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE DATE(COALESCE(w.completed_at, w.closed_at)) = ?
        AND w.status IN ('completed', 'approved', 'closed')
      GROUP BY a.id
    `).all(settings.labor_cost_per_hour_default, date);

    const assetDowntimeCost = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS downtime_cost
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date = ?
      GROUP BY a.id
    `).all(settings.downtime_cost_per_hour_default, date);

    const byAsset = new Map();
    const putCost = (rows, keyName) => {
      for (const r of rows) {
        const code = String(r.asset_code || "UNLINKED");
        if (!byAsset.has(code)) {
          byAsset.set(code, {
            asset_code: code,
            asset_name: r.asset_name || "Unlinked",
            fuel_cost: 0,
            lube_cost: 0,
            parts_cost: 0,
            labor_cost: 0,
            downtime_cost: 0,
            total_cost: 0,
          });
        }
        const row = byAsset.get(code);
        row[keyName] += Number(r[keyName] || 0);
      }
    };
    putCost(assetFuelCost, "fuel_cost");
    putCost(assetLubeCost, "lube_cost");
    putCost(partsCostRows, "parts_cost");
    putCost(assetLaborCost, "labor_cost");
    putCost(assetDowntimeCost, "downtime_cost");
    const top_asset_costs = Array.from(byAsset.values())
      .map((r) => ({
        ...r,
        fuel_cost: Number(r.fuel_cost.toFixed(2)),
        lube_cost: Number(r.lube_cost.toFixed(2)),
        parts_cost: Number(r.parts_cost.toFixed(2)),
        labor_cost: Number(r.labor_cost.toFixed(2)),
        downtime_cost: Number(r.downtime_cost.toFixed(2)),
        total_cost: Number((r.fuel_cost + r.lube_cost + r.parts_cost + r.labor_cost + r.downtime_cost).toFixed(2)),
      }))
      .filter((r) => r.total_cost > 0)
      .sort((a, b) => b.total_cost - a.total_cost)
      .slice(0, 8);

    const fuel_cost = Number(fuelCostRow?.fuel_cost || 0);
    const lube_cost = Number(lubeCostRow?.lube_cost || 0);
    const labor_cost = Number(woCostRow?.labor_cost || 0);
    const labor_hours = Number(woCostRow?.labor_hours || 0);
    const downtime_cost = Number(downtimeCostRow?.downtime_cost || 0);
    const total_cost = Number((fuel_cost + lube_cost + parts_cost + labor_cost + downtime_cost).toFixed(2));
    const cost_per_run_hour = run_hours > 0 ? Number((total_cost / run_hours).toFixed(2)) : null;

    return {
      ok: true,
      date,

      // Keep this for UI display; KPI uses per-row scheduled from daily_hours
      scheduled_hours_per_asset: scheduledFallback,

      kpi: {
        used_assets,
        scheduled_hours,
        available_hours,
        utilization_base_hours: Number(mtd_utilization_base.toFixed(2)),
        run_hours: mtd_run,
        downtime_hours,
        availability: availability == null ? null : Number(availability.toFixed(2)),
        utilization: utilization == null ? null : Number(utilization.toFixed(2)),
        basis: "mtd",
        mtd_start: mtdStart,
        mtd_end: date,
      },

      alerts: {
        low_stock: Number(lowStockCount.c || 0),
        overdue_maintenance: Number(overdueMaintCount.c || 0),
        open_work_orders: Number(openWOCount.c || 0)
      },

      major_downtime: majorDowntime,
      downtime_reasons: downtimeReasons,
      per_asset_kpi,

      critical_low_stock: criticalLowStock,
      open_work_orders: openWOs,
      workorder_sla: {
        summary: sla_summary,
        breaches: sla_breaches,
      },
      lube_usage: {
        date,
        qty_total: Number(lubeTotalRow?.qty_total || 0),
        total_lube_cost: Number(lubeTotalRow?.total_lube_cost || 0),
        rows: lubeDaily,
      },
      cost_engine: {
        date,
        settings,
        fuel_cost: Number(fuel_cost.toFixed(2)),
        lube_cost: Number(lube_cost.toFixed(2)),
        parts_cost,
        labor_cost: Number(labor_cost.toFixed(2)),
        labor_hours: Number(labor_hours.toFixed(2)),
        downtime_cost: Number(downtime_cost.toFixed(2)),
        total_cost,
        cost_per_run_hour,
        top_asset_costs,
      },
    };
  });

  // GET /api/dashboard/lube?start=YYYY-MM-DD&end=YYYY-MM-DD
  app.get("/lube", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }

    const defaultLubeCost = Number(
      db.prepare(`SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1`).get()?.value
    );
    const lubeUnitFallback = Number.isFinite(defaultLubeCost) && defaultLubeCost > 0 ? defaultLubeCost : 4.0;

    const rows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS total_lube_cost,
        COUNT(*) AS entries
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.id
      ORDER BY qty_total DESC, a.asset_code ASC
      LIMIT 200
    `).all(lubeUnitFallback, start, end).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      qty_total: Number(r.qty_total || 0),
      total_lube_cost: Number(r.total_lube_cost || 0),
      entries: Number(r.entries || 0),
    }));

    const byTypeRows = db.prepare(`
      SELECT
        a.asset_code,
        CASE
          WHEN LOWER(TRIM(COALESCE(ol.oil_type, ''))) IN ('admin','supervisor','manager','stores','artisan','operator') THEN 'UNSPECIFIED'
          ELSE COALESCE(NULLIF(TRIM(ol.oil_type), ''), 'UNSPECIFIED')
        END AS oil_type,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS total_lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.asset_code, oil_type
      ORDER BY a.asset_code ASC, qty_total DESC, oil_type ASC
      LIMIT 1200
    `).all(lubeUnitFallback, start, end);
    const byTypeLookup = new Map();
    for (const r of byTypeRows) {
      const code = String(r.asset_code || "");
      if (!byTypeLookup.has(code)) byTypeLookup.set(code, []);
      byTypeLookup.get(code).push({
        oil_type: String(r.oil_type || "UNSPECIFIED"),
        qty_total: Number(r.qty_total || 0),
        total_lube_cost: Number(r.total_lube_cost || 0),
      });
    }
    for (const row of rows) {
      row.by_oil_type = byTypeLookup.get(String(row.asset_code || "")) || [];
    }

    const summary = db.prepare(`
      SELECT
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS total_lube_cost,
        COUNT(*) AS entries,
        COUNT(DISTINCT ol.asset_id) AS assets
      FROM oil_logs ol
      WHERE ol.log_date BETWEEN ? AND ?
    `).get(lubeUnitFallback, start, end);

    return {
      ok: true,
      start,
      end,
      summary: {
        qty_total: Number(summary?.qty_total || 0),
        total_lube_cost: Number(summary?.total_lube_cost || 0),
        entries: Number(summary?.entries || 0),
        assets: Number(summary?.assets || 0),
      },
      rows,
    };
  });

  // GET /api/dashboard/lube/analytics?months=6
  // Usage by oil type/stock code + monthly trend + low-stock forecast
  app.get("/lube/analytics", async (req, reply) => {
    const monthsRaw = Number(req.query?.months ?? 6);
    const months = Number.isFinite(monthsRaw) ? Math.max(1, Math.min(24, Math.trunc(monthsRaw))) : 6;

    const endDate = todayYYYYMMDD();
    const endObj = new Date(`${endDate}T00:00:00`);
    const startObj = new Date(endObj);
    startObj.setMonth(startObj.getMonth() - (months - 1));
    startObj.setDate(1);
    const startDate = startObj.toISOString().slice(0, 10);

    const daySpan = Math.max(1, Math.round((Date.parse(`${endDate}T00:00:00`) - Date.parse(`${startDate}T00:00:00`)) / 86400000) + 1);

    const byType = db.prepare(`
      SELECT
        CASE
          WHEN LOWER(TRIM(COALESCE(ol.oil_type, ''))) IN ('admin','supervisor','manager','stores','artisan','operator') THEN 'UNSPECIFIED'
          ELSE COALESCE(NULLIF(TRIM(ol.oil_type), ''), 'UNSPECIFIED')
        END AS oil_key,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COUNT(*) AS entries
      FROM oil_logs ol
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY oil_key
      ORDER BY qty_total DESC
      LIMIT 200
    `).all(startDate, endDate).map((r) => ({
      oil_key: r.oil_key,
      qty_total: Number(r.qty_total || 0),
      entries: Number(r.entries || 0),
    }));

    const trend = db.prepare(`
      SELECT
        SUBSTR(ol.log_date, 1, 7) AS month,
        CASE
          WHEN LOWER(TRIM(COALESCE(ol.oil_type, ''))) IN ('admin','supervisor','manager','stores','artisan','operator') THEN 'UNSPECIFIED'
          ELSE COALESCE(NULLIF(TRIM(ol.oil_type), ''), 'UNSPECIFIED')
        END AS oil_key,
        COALESCE(SUM(ol.quantity), 0) AS qty
      FROM oil_logs ol
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY month, oil_key
      ORDER BY month ASC, qty DESC
      LIMIT 800
    `).all(startDate, endDate).map((r) => ({
      month: r.month,
      oil_key: r.oil_key,
      qty: Number(r.qty || 0),
    }));

    const mapRows = db.prepare(`
      SELECT oil_key, part_code
      FROM lube_type_mappings
    `).all();
    const mapping = new Map(
      mapRows.map((m) => [String(m.oil_key || "").trim().toLowerCase(), String(m.part_code || "").trim()])
    );

    const getPart = db.prepare(`
      SELECT part_code, part_name, min_stock, IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      WHERE p.part_code = ?
      GROUP BY p.id
      LIMIT 1
    `);

    const forecast = byType.map((r) => {
      const mappedCode = mapping.get(String(r.oil_key || "").trim().toLowerCase()) || String(r.oil_key || "");
      const p = getPart.get(mappedCode);
      const avg_daily_use = Number((r.qty_total / daySpan).toFixed(3));
      const on_hand = p ? Number(p.on_hand || 0) : null;
      const min_stock = p ? Number(p.min_stock || 0) : null;
      const days_to_min =
        p && avg_daily_use > 0
          ? Number(((on_hand - min_stock) / avg_daily_use).toFixed(1))
          : null;
      return {
        oil_key: r.oil_key,
        qty_total: Number(r.qty_total.toFixed(2)),
        entries: r.entries,
        avg_daily_use,
        mapped_part_code: mappedCode || null,
        part_code: p?.part_code || null,
        part_name: p?.part_name || null,
        on_hand,
        min_stock,
        days_to_min,
        low_risk: p ? (on_hand <= min_stock || (days_to_min != null && days_to_min <= 30)) : false,
      };
    }).sort((a, b) => {
      const ar = Number(Boolean(a.low_risk));
      const br = Number(Boolean(b.low_risk));
      if (ar !== br) return br - ar;
      const ad = a.days_to_min == null ? 999999 : a.days_to_min;
      const bd = b.days_to_min == null ? 999999 : b.days_to_min;
      return ad - bd;
    });

    return reply.send({
      ok: true,
      start: startDate,
      end: endDate,
      months,
      summary: {
        oils: byType.length,
        qty_total: Number(byType.reduce((acc, x) => acc + Number(x.qty_total || 0), 0).toFixed(2)),
        low_risk_count: forecast.filter((x) => x.low_risk).length,
      },
      by_type: byType,
      trend,
      forecast,
    });
  });

  // GET /api/dashboard/lube/mappings
  app.get("/lube/mappings", async () => {
    const rows = db.prepare(`
      SELECT oil_key, part_code, updated_by, updated_at
      FROM lube_type_mappings
      ORDER BY oil_key ASC
      LIMIT 400
    `).all();
    return { ok: true, rows };
  });

  // POST /api/dashboard/lube/mappings
  // Body: { oil_key, part_code }
  app.post("/lube/mappings", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const oil_key = String(req.body?.oil_key || "").trim();
    const part_code = String(req.body?.part_code || "").trim();
    if (!oil_key) return reply.code(400).send({ error: "oil_key is required" });
    if (!part_code) return reply.code(400).send({ error: "part_code is required" });

    const part = db.prepare(`
      SELECT id, part_code, part_name
      FROM parts
      WHERE part_code = ?
    `).get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });

    const user = String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
    db.prepare(`
      INSERT INTO lube_type_mappings (oil_key, part_code, updated_by, updated_at)
      VALUES (?, ?, ?, datetime('now'))
      ON CONFLICT(oil_key) DO UPDATE SET
        part_code = excluded.part_code,
        updated_by = excluded.updated_by,
        updated_at = datetime('now')
    `).run(oil_key, part_code, user);

    writeAudit(db, req, {
      module: "lube",
      action: "mapping_upsert",
      entity_type: "lube_mapping",
      entity_id: oil_key,
      payload: { oil_key, part_code },
    });

    return { ok: true, oil_key, part_code, part_name: part.part_name };
  });

  // GET /api/dashboard/cost/settings
  app.get("/cost/settings", async () => {
    const rows = db.prepare(`
      SELECT key, value
      FROM cost_settings
      ORDER BY key ASC
    `).all();
    const settings = {};
    for (const r of rows) settings[String(r.key)] = Number(r.value || 0);
    return { ok: true, settings };
  });

  // POST /api/dashboard/cost/settings
  // Body: { fuel_cost_per_liter_default?, lube_cost_per_qty_default?, labor_cost_per_hour_default?, downtime_cost_per_hour_default? }
  app.post("/cost/settings", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const body = req.body || {};
    const allowed = [
      "fuel_cost_per_liter_default",
      "lube_cost_per_qty_default",
      "labor_cost_per_hour_default",
      "downtime_cost_per_hour_default",
    ];
    const updates = [];
    for (const k of allowed) {
      if (body[k] == null || body[k] === "") continue;
      const v = Number(body[k]);
      if (!Number.isFinite(v) || v < 0) {
        return reply.code(400).send({ error: `${k} must be a valid number >= 0` });
      }
      updates.push({ key: k, value: Number(v.toFixed(4)) });
    }
    if (!updates.length) return reply.code(400).send({ error: "provide at least one setting value" });

    const upsert = db.prepare(`
      INSERT INTO cost_settings (key, value, updated_at)
      VALUES (?, ?, datetime('now'))
      ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = datetime('now')
    `);
    const tx = db.transaction((rowsToSave) => {
      for (const r of rowsToSave) upsert.run(r.key, r.value);
    });
    tx(updates);

    writeAudit(db, req, {
      module: "cost",
      action: "settings_update",
      entity_type: "cost_settings",
      payload: updates,
    });

    return { ok: true, updates };
  });

  // POST /api/dashboard/cost/asset-rates
  // Body: { asset_code, fuel_cost_per_liter?, downtime_cost_per_hour?, utilization_mode?, km_per_hour_factor? }
  app.post("/cost/asset-rates", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const body = req.body || {};
    const asset_code = String(body.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ error: "asset_code is required" });

    const asset = db.prepare(`
      SELECT id, asset_code, asset_name
      FROM assets
      WHERE asset_code = ?
    `).get(asset_code);
    if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });

    const fuelCostRaw = body.fuel_cost_per_liter;
    const downCostRaw = body.downtime_cost_per_hour;
    const utilizationModeRaw = body.utilization_mode;
    const kmPerHourRaw = body.km_per_hour_factor;
    const updates = [];
    const params = [];
    if (fuelCostRaw != null && String(fuelCostRaw).trim() !== "") {
      const v = Number(fuelCostRaw);
      if (!Number.isFinite(v) || v < 0) return reply.code(400).send({ error: "fuel_cost_per_liter must be >= 0" });
      updates.push("fuel_cost_per_liter = ?");
      params.push(Number(v.toFixed(4)));
    }
    if (downCostRaw != null && String(downCostRaw).trim() !== "") {
      const v = Number(downCostRaw);
      if (!Number.isFinite(v) || v < 0) return reply.code(400).send({ error: "downtime_cost_per_hour must be >= 0" });
      updates.push("downtime_cost_per_hour = ?");
      params.push(Number(v.toFixed(4)));
    }
    if (utilizationModeRaw != null && String(utilizationModeRaw).trim() !== "") {
      const mode = String(utilizationModeRaw).trim().toLowerCase();
      if (!["hours", "km"].includes(mode)) {
        return reply.code(400).send({ error: "utilization_mode must be 'hours' or 'km'" });
      }
      updates.push("utilization_mode = ?");
      params.push(mode);
    }
    if (kmPerHourRaw != null && String(kmPerHourRaw).trim() !== "") {
      const v = Number(kmPerHourRaw);
      if (!Number.isFinite(v) || v <= 0) return reply.code(400).send({ error: "km_per_hour_factor must be > 0" });
      updates.push("km_per_hour_factor = ?");
      params.push(Number(v.toFixed(4)));
    }
    if (!updates.length) return reply.code(400).send({ error: "provide at least one asset rate value" });

    db.prepare(`
      UPDATE assets
      SET ${updates.join(", ")}
      WHERE id = ?
    `).run(...params, asset.id);

    const row = db.prepare(`
      SELECT asset_code, asset_name, fuel_cost_per_liter, downtime_cost_per_hour, utilization_mode, km_per_hour_factor
      FROM assets
      WHERE id = ?
    `).get(asset.id);

    writeAudit(db, req, {
      module: "cost",
      action: "asset_rate_update",
      entity_type: "asset",
      entity_id: asset_code,
      payload: row,
    });

    return { ok: true, asset: row };
  });

  // POST /api/dashboard/cost/part-cost
  // Body: { part_code, unit_cost }
  app.post("/cost/part-cost", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const body = req.body || {};
    const part_code = String(body.part_code || "").trim();
    const unit_cost = Number(body.unit_cost);
    if (!part_code) return reply.code(400).send({ error: "part_code is required" });
    if (!Number.isFinite(unit_cost) || unit_cost < 0) return reply.code(400).send({ error: "unit_cost must be >= 0" });

    const part = db.prepare(`
      SELECT id, part_code, part_name
      FROM parts
      WHERE part_code = ?
    `).get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });

    db.prepare(`
      UPDATE parts
      SET unit_cost = ?
      WHERE id = ?
    `).run(Number(unit_cost.toFixed(4)), part.id);

    writeAudit(db, req, {
      module: "cost",
      action: "part_cost_update",
      entity_type: "part",
      entity_id: part_code,
      payload: { unit_cost: Number(unit_cost.toFixed(4)) },
    });

    return { ok: true, part_code: part.part_code, part_name: part.part_name, unit_cost: Number(unit_cost.toFixed(4)) };
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

    const asset = db.prepare(`SELECT id FROM assets WHERE asset_code = ?`).get(asset_code);
    if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });

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
      INSERT INTO fuel_logs (asset_id, log_date, liters, source, hours_run, meter_run_value, meter_unit)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `).run(asset.id, log_date, liters, source, hours_final, meter_run_value, meter_unit);

    writeAudit(db, req, {
      module: "fuel",
      action: "manual_log",
      entity_type: "asset",
      entity_id: asset_code,
      payload: { log_date, liters, hours_run: hours_final, meter_run_value, meter_unit, source },
    });

    return reply.send({
      ok: true,
      id: Number(ins.lastInsertRowid),
      asset_code,
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
      db.prepare(`UPDATE assets SET baseline_fuel_km_per_l = ? WHERE id = ?`).run(v, asset.id);
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
    db.prepare(`UPDATE assets SET baseline_fuel_l_per_hour = ? WHERE id = ?`).run(v, asset.id);

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

    const fuelByAsset = db.prepare(`
      SELECT
        a.id AS asset_id,
        a.asset_code,
        a.asset_name,
        a.category,
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
        END AS metric_mode,
        COALESCE(NULLIF(a.km_per_hour_factor, 0), 10.0) AS km_per_hour_factor,
        COALESCE(a.baseline_fuel_l_per_hour, 5.0) AS oem_lph,
        COALESCE(a.baseline_fuel_km_per_l, 2.0) AS oem_kmpl,
        COALESCE(SUM(fl.liters), 0) AS fuel_liters,
        COUNT(fl.id) AS fill_count
      FROM assets a
      LEFT JOIN fuel_logs fl
        ON fl.asset_id = a.id
       AND fl.log_date BETWEEN ? AND ?
      WHERE a.active = 1
        AND UPPER(COALESCE(a.asset_code, '')) NOT GLOB 'V[0-9][0-9]AM'
      GROUP BY a.id
      ORDER BY a.asset_code ASC
    `).all(start, end);

    const getFuelLogsInRange = db.prepare(`
      SELECT
        id,
        log_date,
        COALESCE(LOWER(meter_unit), '') AS meter_unit,
        COALESCE(meter_run_value, 0) AS meter_run_value,
        COALESCE(hours_run, 0) AS hours_run,
        open_meter_value,
        close_meter_value
      FROM fuel_logs
      WHERE asset_id = ?
        AND log_date BETWEEN ? AND ?
      ORDER BY log_date ASC, id ASC
    `);
    const getFuelLogBeforeRange = db.prepare(`
      SELECT
        id,
        log_date,
        COALESCE(LOWER(meter_unit), '') AS meter_unit,
        COALESCE(meter_run_value, 0) AS meter_run_value,
        COALESCE(hours_run, 0) AS hours_run,
        open_meter_value,
        close_meter_value
      FROM fuel_logs
      WHERE asset_id = ?
        AND log_date < ?
        AND (
          COALESCE(meter_run_value, 0) > 0
          OR COALESCE(hours_run, 0) > 0
        )
      ORDER BY log_date DESC, id DESC
      LIMIT 1
    `);

    function getRunFromFuel(assetId, startDate, endDate) {
      const rows = getFuelLogsInRange.all(assetId, startDate, endDate);
      if (!rows.length) return { km_run: 0, hours_run: 0 };

      const prev = getFuelLogBeforeRange.get(assetId, startDate);
      let prevKmMeter = null;
      let prevHoursMeter = null;
      if (prev) {
        const prevUnit = String(prev.meter_unit || "").toLowerCase();
        const prevMeter = Number(prev.meter_run_value || 0);
        if (prevUnit === "km" && prevMeter > 0) prevKmMeter = prevMeter;
        if (prevUnit === "hours" && prevMeter > 0) prevHoursMeter = prevMeter;
      }

      let km_run = 0;
      let hours_run = 0;

      for (const row of rows) {
        const unit = String(row.meter_unit || "").toLowerCase();
        const meter = Number(row.meter_run_value || 0);
        const legacyHours = Number(row.hours_run || 0);
        const openMeter = row.open_meter_value == null ? null : Number(row.open_meter_value);
        const closeMeter = row.close_meter_value == null ? null : Number(row.close_meter_value);

        if (openMeter != null && closeMeter != null && closeMeter > openMeter) {
          const delta = closeMeter - openMeter;
          if (unit === "km") km_run += delta;
          else hours_run += delta;
          continue;
        }

        if (unit === "km" && meter > 0) {
          if (prevKmMeter != null) {
            const delta = meter - prevKmMeter;
            if (Number.isFinite(delta) && delta > 0) km_run += delta;
          }
          prevKmMeter = meter;
          continue;
        }

        if (unit === "hours" && meter > 0) {
          if (prevHoursMeter != null) {
            const delta = meter - prevHoursMeter;
            if (Number.isFinite(delta) && delta > 0) hours_run += delta;
          }
          prevHoursMeter = meter;
          continue;
        }

        // Legacy rows where hours_run already stores run-between-fills.
        if (legacyHours > 0) hours_run += legacyHours;
      }

      return { km_run, hours_run };
    }

    const rows = fuelByAsset.map((r) => {
      const mode = String(r.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
      const fuelRun = getRunFromFuel(r.asset_id, start, end) || {};
      const fuelKm = Number(fuelRun.km_run || 0);
      const fuelHours = Number(fuelRun.hours_run || 0);
      const km = fuelKm > 0 ? fuelKm : 0;
      const hours = fuelHours > 0 ? fuelHours : 0;
      const fuel = Number(r.fuel_liters || 0);
      const oem = Number(r.oem_lph || 5);
      const oemK = Number(r.oem_kmpl || 2);
      const fillCount = Number(r.fill_count || 0);
      const lph = hours > 0 ? fuel / hours : null;
      const kmpl = fuel > 0 && km > 0 ? km / fuel : null;
      const excessiveThreshold = oem * (1 + tolerance);
      const lowThresholdKmpl = oemK * Math.max(0, 1 - tolerance);
      const hasEnoughSamples = fillCount >= 2;
      const is_excessive = hasEnoughSamples && (mode === "km"
        ? (kmpl != null && kmpl < lowThresholdKmpl)
        : (lph != null && lph > excessiveThreshold));
      return {
        asset_id: Number(r.asset_id),
        asset_code: r.asset_code,
        asset_name: r.asset_name,
        metric_mode: mode,
        fuel_liters: Number(fuel.toFixed(2)),
        km_run: Number(km.toFixed(2)),
        hours_run: Number(hours.toFixed(2)),
        actual_lph: lph == null ? null : Number(lph.toFixed(3)),
        oem_lph: Number(oem.toFixed(3)),
        excessive_threshold_lph: Number(excessiveThreshold.toFixed(3)),
        variance_lph: lph == null ? null : Number((lph - oem).toFixed(3)),
        actual_km_per_l: kmpl == null ? null : Number(kmpl.toFixed(3)),
        oem_km_per_l: Number(oemK.toFixed(3)),
        low_threshold_km_per_l: Number(lowThresholdKmpl.toFixed(3)),
        variance_km_per_l: kmpl == null ? null : Number((kmpl - oemK).toFixed(3)),
        fill_count: fillCount,
        has_enough_samples: hasEnoughSamples,
        is_excessive,
      };
    }).filter((r) => r.fuel_liters > 0)
      // Temporary business rule: exclude LDV/km-mode assets from benchmark list/flags.
      .filter((r) => r.metric_mode !== "km")
      .filter((r) => (assetFilter ? String(r.asset_code || "").trim().toLowerCase() === assetFilter : true))
      .filter((r) => (modeFilter === "km" ? r.metric_mode === "km" : modeFilter === "hours" ? r.metric_mode === "hours" : true))
      .sort((a, b) => {
        const ex = Number(Boolean(b.is_excessive)) - Number(Boolean(a.is_excessive));
        if (ex !== 0) return ex;
        const av = a.metric_mode === "km" ? Number(a.variance_km_per_l || -999) : Number(a.variance_lph || -999);
        const bv = b.metric_mode === "km" ? Number(b.variance_km_per_l || -999) : Number(b.variance_lph || -999);
        return bv - av;
      });

    const summary = rows.reduce(
      (acc, r) => {
        acc.assets += 1;
        acc.fuel_liters += Number(r.fuel_liters || 0);
        acc.hours_run += Number(r.hours_run || 0);
        acc.km_run += Number(r.km_run || 0);
        if (r.metric_mode === "km") {
          acc.km_assets += 1;
          if (Number(r.km_run || 0) > 0) acc.km_fuel += Number(r.fuel_liters || 0);
        } else {
          acc.hours_assets += 1;
          if (Number(r.hours_run || 0) > 0) acc.hours_fuel += Number(r.fuel_liters || 0);
        }
        if (r.is_excessive) acc.excessive_count += 1;
        return acc;
      },
      { assets: 0, fuel_liters: 0, hours_run: 0, km_run: 0, excessive_count: 0, km_assets: 0, hours_assets: 0, km_fuel: 0, hours_fuel: 0 }
    );

    summary.fuel_liters = Number(summary.fuel_liters.toFixed(2));
    summary.hours_run = Number(summary.hours_run.toFixed(2));
    summary.km_run = Number(summary.km_run.toFixed(2));
    summary.avg_lph = summary.hours_run > 0
      ? Number((summary.hours_fuel / summary.hours_run).toFixed(3))
      : null;
    summary.avg_km_per_l = summary.km_fuel > 0
      ? Number((summary.km_run / summary.km_fuel).toFixed(3))
      : null;

    return reply.send({
      ok: true,
      start,
      end,
      tolerance,
      mode: modeFilter || null,
      summary,
      rows,
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
      SELECT
        id, asset_code, asset_name, category,
        COALESCE(baseline_fuel_l_per_hour, 5.0) AS oem_lph,
        COALESCE(baseline_fuel_km_per_l, 2.0) AS oem_kmpl,
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
        END AS metric_mode,
        COALESCE(NULLIF(km_per_hour_factor, 0), 10.0) AS km_per_hour_factor
      FROM assets
      WHERE asset_code = ?
    `).get(assetCode);
    if (!asset) return reply.code(404).send({ error: `asset not found: ${assetCode}` });

    const fuelRows = db.prepare(`
      SELECT
        fl.id,
        fl.log_date,
        COALESCE(fl.liters, 0) AS fuel_liters,
        COALESCE(CASE WHEN fl.hours_run > 0 THEN fl.hours_run ELSE 0 END, 0) AS meter_hours,
        COALESCE(fl.meter_run_value, 0) AS meter_run_value,
        COALESCE(LOWER(fl.meter_unit), '') AS meter_unit,
        fl.open_meter_value,
        fl.close_meter_value,
        fl.source
      FROM fuel_logs fl
      WHERE fl.asset_id = ?
        AND fl.log_date BETWEEN ? AND ?
      ORDER BY fl.log_date ASC, fl.id ASC
    `).all(asset.id, start, end);

    // First row in selected window should still calculate from last fill before start.
    const previousFill = db.prepare(`
      SELECT
        fl.id,
        fl.log_date,
        fl.hours_run,
        fl.meter_run_value,
        COALESCE(LOWER(fl.meter_unit), '') AS meter_unit
      FROM fuel_logs fl
      WHERE fl.asset_id = ?
        AND fl.log_date < ?
        AND (
          fl.hours_run > 0
          OR fl.meter_run_value > 0
        )
      ORDER BY fl.log_date DESC, fl.id DESC
      LIMIT 1
    `).get(asset.id, start);

    const mode = String(asset.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
    const kmPerHour = Math.max(0.1, Number(asset.km_per_hour_factor || 10));
    const oem = Number(asset.oem_lph || 5);
    const oemK = Number(asset.oem_kmpl || 2);
    const threshold = oem * (1 + tolerance);
    const lowThresholdKmpl = oemK * Math.max(0, 1 - tolerance);

    function toModeMeter(row) {
      const unit = String(row?.meter_unit || "").toLowerCase();
      const v = Number(row?.meter_run_value || 0);
      if (unit === "km" && v > 0) return mode === "km" ? v : (v / kmPerHour);
      if (unit === "hours" && v > 0) return mode === "km" ? (v * kmPerHour) : v;
      const h = Number(row?.meter_hours || row?.hours_run || 0);
      if (h > 0) return mode === "km" ? (h * kmPerHour) : h;
      return 0;
    }

    let prevMeter = previousFill ? toModeMeter(previousFill) : 0;
    if (!(prevMeter > 0)) prevMeter = null;

    const rows = fuelRows.map((d) => {
      const meter = toModeMeter(d);
      const rowOpen = d.open_meter_value == null ? null : Number(d.open_meter_value);
      const rowClose = d.close_meter_value == null ? null : Number(d.close_meter_value);
      const hasRowMeters = rowOpen != null && rowClose != null && rowOpen > 0 && rowClose > 0;

      let openMeter = prevMeter;
      let closeMeter = meter > 0 ? meter : null;
      let runBetween = null;
      let invalidDelta = false;

      // Prefer explicit per-fill opening/closing readings from fuel logs.
      if (hasRowMeters) {
        openMeter = rowOpen;
        closeMeter = rowClose;
        const delta = rowClose - rowOpen;
        if (Number.isFinite(delta) && delta > 0) runBetween = delta;
        else if (Number.isFinite(delta) && delta <= 0) invalidDelta = true;
      } else if (openMeter != null && meter > 0) {
        const delta = meter - openMeter;
        if (Number.isFinite(delta) && delta > 0) runBetween = delta;
        else if (Number.isFinite(delta) && delta <= 0) invalidDelta = true;
      }
      const fuel = Number(d.fuel_liters || 0);
      const lph = (!invalidDelta && mode === "hours" && runBetween != null && runBetween > 0) ? (fuel / runBetween) : null;
      const kmpl = (!invalidDelta && mode === "km" && fuel > 0 && runBetween != null && runBetween > 0) ? (runBetween / fuel) : null;
      const isExcessive = mode === "km" ? (kmpl != null && kmpl < lowThresholdKmpl) : (lph != null && lph > threshold);
      if (closeMeter != null && closeMeter > 0) prevMeter = closeMeter;
      return {
        id: Number(d.id),
        log_date: d.log_date,
        metric_mode: mode,
        fuel_liters: Number(fuel.toFixed(2)),
        run_value: runBetween == null ? 0 : Number(runBetween.toFixed(2)),
        run_unit: mode === "km" ? "km" : "hours",
        hours_run: mode === "hours" ? (runBetween == null ? 0 : Number(runBetween.toFixed(2))) : 0,
        km_run: mode === "km" ? (runBetween == null ? 0 : Number(runBetween.toFixed(2))) : 0,
        meter_value: closeMeter != null && closeMeter > 0 ? Number(closeMeter.toFixed(2)) : null,
        open_meter_value: openMeter != null && openMeter > 0 ? Number(openMeter.toFixed(2)) : null,
        close_meter_value: closeMeter != null && closeMeter > 0 ? Number(closeMeter.toFixed(2)) : null,
        meter_unit_display: mode === "km" ? "km" : "hours",
        invalid_delta: invalidDelta,
        actual_lph: lph == null ? null : Number(lph.toFixed(3)),
        oem_lph: Number(oem.toFixed(3)),
        excessive_threshold_lph: Number(threshold.toFixed(3)),
        actual_km_per_l: kmpl == null ? null : Number(kmpl.toFixed(3)),
        oem_km_per_l: Number(oemK.toFixed(3)),
        low_threshold_km_per_l: Number(lowThresholdKmpl.toFixed(3)),
        is_excessive: isExcessive,
        source: d.source || null,
      };
    });

    const summary = rows.reduce(
      (acc, r) => {
        acc.days += 1;
        acc.fuel_liters += Number(r.fuel_liters || 0);
        acc.hours_run += Number(r.hours_run || 0);
        acc.km_run += Number(r.km_run || 0);
        if (r.is_excessive) acc.excessive_days += 1;
        return acc;
      },
      { days: 0, fuel_liters: 0, hours_run: 0, km_run: 0, excessive_days: 0 }
    );
    summary.fuel_liters = Number(summary.fuel_liters.toFixed(2));
    summary.hours_run = Number(summary.hours_run.toFixed(2));
    summary.km_run = Number(summary.km_run.toFixed(2));
    summary.metric_mode = mode;
    summary.avg_lph = summary.hours_run > 0
      ? Number((summary.fuel_liters / summary.hours_run).toFixed(3))
      : null;
    summary.avg_km_per_l = summary.fuel_liters > 0 && summary.km_run > 0
      ? Number((summary.km_run / summary.fuel_liters).toFixed(3))
      : null;

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

  // GET /api/dashboard/cost/trend?months=12&end_month=YYYY-MM
  app.get("/cost/trend", async (req, reply) => {
    const monthsRaw = Number(req.query?.months ?? 12);
    const months = Number.isFinite(monthsRaw) ? Math.max(1, Math.min(36, Math.trunc(monthsRaw))) : 12;
    const endMonthInput = String(req.query?.end_month || "").trim();
    const nowMonth = todayYYYYMMDD().slice(0, 7);
    const endMonth = /^\d{4}-\d{2}$/.test(endMonthInput) ? endMonthInput : nowMonth;
    if (!/^\d{4}-\d{2}$/.test(endMonth)) {
      return reply.code(400).send({ error: "end_month must be YYYY-MM" });
    }

    const settingsRows = db.prepare(`
      SELECT key, value
      FROM cost_settings
      WHERE key IN (
        'fuel_cost_per_liter_default',
        'lube_cost_per_qty_default',
        'labor_cost_per_hour_default',
        'downtime_cost_per_hour_default'
      )
    `).all();
    const settings = {
      fuel_cost_per_liter_default: 1.5,
      lube_cost_per_qty_default: 4.0,
      labor_cost_per_hour_default: 35.0,
      downtime_cost_per_hour_default: 120.0,
    };
    for (const r of settingsRows) {
      const k = String(r.key || "").trim();
      const v = Number(r.value);
      if (k && Number.isFinite(v)) settings[k] = v;
    }

    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";

    const getMonthCost = (monthId) => {
      const r = monthRangeFromYYYYMM(monthId);

      const fuel = db.prepare(`
        SELECT COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS v
        FROM fuel_logs fl
        JOIN assets a ON a.id = fl.asset_id
        WHERE fl.log_date BETWEEN ? AND ?
      `).get(settings.fuel_cost_per_liter_default, r.start, r.end);

      const lube = db.prepare(`
        SELECT COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS v
        FROM oil_logs ol
        WHERE ol.log_date BETWEEN ? AND ?
      `).get(settings.lube_cost_per_qty_default, r.start, r.end);

      const parts = db.prepare(`
        SELECT COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS v
        FROM stock_movements sm
        JOIN parts p ON p.id = sm.part_id
        WHERE sm.movement_type = 'out'
          AND ${smDateExpr} BETWEEN ? AND ?
      `).get(r.start, r.end);

      const labor = db.prepare(`
        SELECT
          COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS hrs,
          COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS v
        FROM work_orders w
        WHERE DATE(COALESCE(w.completed_at, w.closed_at)) BETWEEN ? AND ?
          AND w.status IN ('completed', 'approved', 'closed')
      `).get(settings.labor_cost_per_hour_default, r.start, r.end);

      const down = db.prepare(`
        SELECT COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS v
        FROM breakdown_downtime_logs l
        JOIN breakdowns b ON b.id = l.breakdown_id
        JOIN assets a ON a.id = b.asset_id
        WHERE l.log_date BETWEEN ? AND ?
      `).get(settings.downtime_cost_per_hour_default, r.start, r.end);

      const run = db.prepare(`
        SELECT COALESCE(SUM(hours_run), 0) AS h
        FROM daily_hours
        WHERE work_date BETWEEN ? AND ?
          AND is_used = 1
          AND hours_run > 0
      `).get(r.start, r.end);

      const fuel_cost = Number(fuel?.v || 0);
      const lube_cost = Number(lube?.v || 0);
      const parts_cost = Number(parts?.v || 0);
      const labor_cost = Number(labor?.v || 0);
      const labor_hours = Number(labor?.hrs || 0);
      const downtime_cost = Number(down?.v || 0);
      const run_hours = Number(run?.h || 0);
      const total_cost = Number((fuel_cost + lube_cost + parts_cost + labor_cost + downtime_cost).toFixed(2));
      return {
        month: monthId,
        fuel_cost: Number(fuel_cost.toFixed(2)),
        lube_cost: Number(lube_cost.toFixed(2)),
        parts_cost: Number(parts_cost.toFixed(2)),
        labor_cost: Number(labor_cost.toFixed(2)),
        labor_hours: Number(labor_hours.toFixed(2)),
        downtime_cost: Number(downtime_cost.toFixed(2)),
        run_hours: Number(run_hours.toFixed(2)),
        total_cost,
        cost_per_run_hour: run_hours > 0 ? Number((total_cost / run_hours).toFixed(2)) : null,
      };
    };

    const endDate = new Date(`${endMonth}-01T00:00:00Z`);
    const rows = [];
    for (let i = months - 1; i >= 0; i--) {
      const d = new Date(endDate);
      d.setUTCMonth(d.getUTCMonth() - i);
      const m = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
      rows.push(getMonthCost(m));
    }

    const prev = rows.length >= 2 ? rows[rows.length - 2] : null;
    const cur = rows.length ? rows[rows.length - 1] : null;
    const mom_variance = prev && cur ? Number((cur.total_cost - prev.total_cost).toFixed(2)) : null;
    const mom_variance_pct = prev && cur && Number(prev.total_cost) > 0
      ? Number((((cur.total_cost - prev.total_cost) / prev.total_cost) * 100).toFixed(2))
      : null;

    return {
      ok: true,
      months,
      end_month: endMonth,
      rows,
      latest: cur,
      mom: {
        previous_month: prev?.month || null,
        current_month: cur?.month || null,
        variance: mom_variance,
        variance_pct: mom_variance_pct,
      },
    };
  });

  // GET /api/dashboard/reliability?start=YYYY-MM-DD&end=YYYY-MM-DD
  // MTBF = operating hours / failure count
  // LTTR = downtime hours / failure count
  app.get("/reliability", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }

    const failuresRow = db.prepare(`
      SELECT COUNT(*) AS n
      FROM breakdowns
      WHERE breakdown_date BETWEEN ? AND ?
    `).get(start, end);
    const failure_count = Number(failuresRow?.n || 0);

    const runRow = db.prepare(`
      SELECT COALESCE(SUM(hours_run), 0) AS run_hours
      FROM daily_hours
      WHERE work_date BETWEEN ? AND ?
        AND is_used = 1
        AND hours_run > 0
    `).get(start, end);
    const operating_hours = Number(runRow?.run_hours || 0);

    let downtime_hours = 0;
    if (hasColumn("breakdowns", "downtime_total_hours")) {
      const dtRow = db.prepare(`
        SELECT COALESCE(SUM(downtime_total_hours), 0) AS dt
        FROM breakdowns
        WHERE breakdown_date BETWEEN ? AND ?
      `).get(start, end);
      downtime_hours = Number(dtRow?.dt || 0);
    } else if (hasColumn("breakdowns", "downtime_hours")) {
      const dtRow = db.prepare(`
        SELECT COALESCE(SUM(downtime_hours), 0) AS dt
        FROM breakdowns
        WHERE breakdown_date BETWEEN ? AND ?
      `).get(start, end);
      downtime_hours = Number(dtRow?.dt || 0);
    }

    const mtbf_hours = failure_count > 0 ? operating_hours / failure_count : null;
    const lttr_hours = failure_count > 0 ? downtime_hours / failure_count : null;

    return {
      ok: true,
      start,
      end,
      failure_count,
      operating_hours: Number(operating_hours.toFixed(2)),
      downtime_hours: Number(downtime_hours.toFixed(2)),
      mtbf_hours: mtbf_hours == null ? null : Number(mtbf_hours.toFixed(2)),
      lttr_hours: lttr_hours == null ? null : Number(lttr_hours.toFixed(2)),
    };
  });

  // GET /api/dashboard/reliability/trend?weeks=12&end=YYYY-MM-DD
  app.get("/reliability/trend", async (req, reply) => {
    const weeks = Math.max(1, Math.min(52, Number(req.query?.weeks || 12)));
    const endStr = String(req.query?.end || todayYYYYMMDD()).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(endStr)) {
      return reply.code(400).send({ error: "end must be YYYY-MM-DD" });
    }
    const endDate = new Date(`${endStr}T00:00:00`);
    const points = [];

    for (let i = weeks - 1; i >= 0; i--) {
      const wEnd = new Date(endDate);
      wEnd.setDate(endDate.getDate() - (i * 7));
      const wStart = new Date(wEnd);
      wStart.setDate(wEnd.getDate() - 6);
      const fmt = (d) => d.toISOString().slice(0, 10);
      const start = fmt(wStart);
      const end = fmt(wEnd);

      const failuresRow = db.prepare(`
        SELECT COUNT(*) AS n
        FROM breakdowns
        WHERE breakdown_date BETWEEN ? AND ?
      `).get(start, end);
      const failure_count = Number(failuresRow?.n || 0);

      const runRow = db.prepare(`
        SELECT COALESCE(SUM(hours_run), 0) AS run_hours
        FROM daily_hours
        WHERE work_date BETWEEN ? AND ?
          AND is_used = 1
          AND hours_run > 0
      `).get(start, end);
      const operating_hours = Number(runRow?.run_hours || 0);

      let downtime_hours = 0;
      if (hasColumn("breakdowns", "downtime_total_hours")) {
        const dtRow = db.prepare(`
          SELECT COALESCE(SUM(downtime_total_hours), 0) AS dt
          FROM breakdowns
          WHERE breakdown_date BETWEEN ? AND ?
        `).get(start, end);
        downtime_hours = Number(dtRow?.dt || 0);
      } else if (hasColumn("breakdowns", "downtime_hours")) {
        const dtRow = db.prepare(`
          SELECT COALESCE(SUM(downtime_hours), 0) AS dt
          FROM breakdowns
          WHERE breakdown_date BETWEEN ? AND ?
        `).get(start, end);
        downtime_hours = Number(dtRow?.dt || 0);
      }

      const mtbf_hours = failure_count > 0 ? operating_hours / failure_count : null;
      const lttr_hours = failure_count > 0 ? downtime_hours / failure_count : null;

      points.push({
        start,
        end,
        label: end.slice(5),
        failure_count,
        mtbf_hours: mtbf_hours == null ? null : Number(mtbf_hours.toFixed(2)),
        lttr_hours: lttr_hours == null ? null : Number(lttr_hours.toFixed(2)),
      });
    }

    return { ok: true, weeks, end: endStr, points };
  });

  // POST /api/dashboard/workorders/:id/nudge
  app.post("/workorders/:id/nudge", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "artisan"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "valid work order id required" });

    const wo = db.prepare(`
      SELECT id, status, opened_at
      FROM work_orders
      WHERE id = ?
    `).get(id);
    if (!wo) return reply.code(404).send({ error: "work order not found" });

    const note = String(req.body?.note || "").trim() || null;
    writeAudit(db, req, {
      module: "workorders",
      action: "nudge_supervisor",
      entity_type: "work_order",
      entity_id: id,
      payload: {
        status: wo.status,
        opened_at: wo.opened_at,
        note,
      },
    });

    return { ok: true, id, nudged: true };
  });
}