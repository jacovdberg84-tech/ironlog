// IRONLOG/api/routes/assets.routes.js
import { db } from "../db/client.js";
import { getAssetCurrentHoursInfo } from "../utils/assetMeterHours.js";
import { ensureMasterDataSchema } from "../utils/masterdataGovernance.js";
import {
  resolveMachinePrestartProfile,
  getMachinePrestartTemplate,
  machinePrestartCheckMode,
} from "../utils/machinePrestartTemplates.js";
import { ensureCostAllocationSchema } from "../utils/costAllocation.js";
import { ensurePlantHireSchema } from "../utils/plantHire.js";
import { resolveNextServiceForAssetPlans, classifyServiceDue } from "../utils/serviceSchedule.js";
import registerRegisterRoutes from "./assets/register.routes.js";
import registerQrProfilesRoutes from "./assets/qr-profiles.routes.js";
import registerHistoryRoutes from "./assets/history.routes.js";

function isLdvPrestartQrAsset(assetCode) {
  return /^V(0[1-9]|1[0-5])AM$/i.test(String(assetCode || "").trim());
}

// These records were imported with the ownership label in the equipment-class
// field. Keep the correction deliberately narrow so a later user edit is never
// overwritten on a subsequent API restart.
const LEGACY_ASSET_CATEGORY_CORRECTIONS = [
  { assetCode: "G02AM", category: "Grader" },
  { assetCode: "W200AM", category: "Water Truck" },
  { assetCode: "W201AM", category: "Water Truck" },
  { assetCode: "E503AM", category: "50T Excavator" },
  { assetCode: "E504AM", category: "50T Excavator" },
];

function applyLegacyAssetCategoryCorrections() {
  const updateCategory = db.prepare(`
    UPDATE assets
    SET category = ?
    WHERE UPPER(asset_code) = ?
      AND LOWER(TRIM(COALESCE(category, ''))) = 'internal (amlph)'
  `);
  for (const correction of LEGACY_ASSET_CATEGORY_CORRECTIONS) {
    updateCategory.run(correction.category, correction.assetCode);
  }
}

export default async function assetRoutes(app) {
  ensureMasterDataSchema();
  ensureCostAllocationSchema(db);
  ensurePlantHireSchema(db);
  applyLegacyAssetCategoryCorrections();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS asset_qr_profiles (
      asset_id INTEGER PRIMARY KEY,
      qr_payload TEXT NOT NULL,
      qr_text TEXT NOT NULL,
      generated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS asset_undercarriage_qr_profiles (
      asset_id INTEGER PRIMARY KEY,
      qr_payload TEXT NOT NULL,
      qr_text TEXT NOT NULL,
      generated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS asset_tyre_qr_profiles (
      asset_id INTEGER PRIMARY KEY,
      qr_payload TEXT NOT NULL,
      qr_text TEXT NOT NULL,
      generated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();

  /* =========================
     PREPARED STATEMENTS
  ========================= */

  const getAssetByCode = db.prepare(`
    SELECT
      id, asset_code, asset_name, category,
      active, is_standby,
      archived, archive_reason, archived_at,
      created_at
    FROM assets
    WHERE asset_code = ?
  `);

  const getAssetById = db.prepare(`
    SELECT
      id, asset_code, asset_name, category,
      active, is_standby,
      archived, archive_reason, archived_at,
      created_at
    FROM assets
    WHERE id = ?
  `);
  const getStoredQrProfile = db.prepare(`
    SELECT qr_payload, qr_text, generated_at
    FROM asset_qr_profiles
    WHERE asset_id = ?
  `);
  const upsertQrProfile = db.prepare(`
    INSERT INTO asset_qr_profiles (asset_id, qr_payload, qr_text, generated_at)
    VALUES (?, ?, ?, datetime('now'))
    ON CONFLICT(asset_id) DO UPDATE SET
      qr_payload = excluded.qr_payload,
      qr_text = excluded.qr_text,
      generated_at = datetime('now')
  `);
  const getStoredUndercarriageQrProfile = db.prepare(`
    SELECT qr_payload, qr_text, generated_at
    FROM asset_undercarriage_qr_profiles
    WHERE asset_id = ?
  `);
  const upsertUndercarriageQrProfile = db.prepare(`
    INSERT INTO asset_undercarriage_qr_profiles (asset_id, qr_payload, qr_text, generated_at)
    VALUES (?, ?, ?, datetime('now'))
    ON CONFLICT(asset_id) DO UPDATE SET
      qr_payload = excluded.qr_payload,
      qr_text = excluded.qr_text,
      generated_at = datetime('now')
  `);
  const getStoredTyreQrProfile = db.prepare(`
    SELECT qr_payload, qr_text, generated_at
    FROM asset_tyre_qr_profiles
    WHERE asset_id = ?
  `);
  const upsertTyreQrProfile = db.prepare(`
    INSERT INTO asset_tyre_qr_profiles (asset_id, qr_payload, qr_text, generated_at)
    VALUES (?, ?, ?, datetime('now'))
    ON CONFLICT(asset_id) DO UPDATE SET
      qr_payload = excluded.qr_payload,
      qr_text = excluded.qr_text,
      generated_at = datetime('now')
  `);

  function hasTable(tableName) {
    const row = db.prepare(`
      SELECT name
      FROM sqlite_master
      WHERE type = 'table' AND name = ?
    `).get(tableName);
    return Boolean(row);
  }
  function hasColumn(tableName, columnName) {
    if (!hasTable(tableName)) return false;
    const cols = db.prepare(`PRAGMA table_info(${tableName})`).all();
    return cols.some((c) => String(c.name || "") === String(columnName));
  }
  function firstExistingColumn(tableName, candidates) {
    for (const c of candidates) {
      if (hasColumn(tableName, c)) return c;
    }
    return null;
  }

  const insertAsset = db.prepare(`
    INSERT INTO assets (
      asset_code, asset_name, category,
      active, is_standby,
      archived, archive_reason, archived_at,
      department_code, cost_center_code, site_code, data_owner_username
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const updateActive = db.prepare(`UPDATE assets SET active = ? WHERE asset_code = ?`);
  const updateStandby = db.prepare(`UPDATE assets SET is_standby = ? WHERE asset_code = ?`);

  function siteCodeFromReq(req) {
    return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
  }

  function buildMachineStatus(assetId) {
    const latestBreakdown = db.prepare(`
      SELECT
        breakdown_date,
        status,
        end_at
      FROM breakdowns
      WHERE asset_id = ? 
      ORDER BY breakdown_date DESC, id DESC
      LIMIT 1
    `).get(assetId);
    const latestDaily = db.prepare(`
      SELECT is_used, work_date
      FROM daily_hours
      WHERE asset_id = ?
      ORDER BY work_date DESC, id DESC
      LIMIT 1
    `).get(assetId);

    const isBreakdownOpen = (() => {
      if (!latestBreakdown) return false;
      const st = String(latestBreakdown.status || "").trim().toUpperCase();
      if (st === "OPEN") return true;
      if (st && st !== "OPEN") return false;
      return !latestBreakdown.end_at;
    })();

    if (isBreakdownOpen) {
      if (!latestDaily?.work_date) return "DOWN";
      const bdDate = String(latestBreakdown.breakdown_date || "");
      const dhDate = String(latestDaily.work_date || "");
      if (bdDate && dhDate && bdDate >= dhDate) return "DOWN";
    }

    if (!latestDaily) return "UNKNOWN";
    return Number(latestDaily.is_used) === 1 ? "PRODUCTION" : "STANDBY";
  }

  function buildInspectionSummary(assetId) {
    const candidates = [];
    if (hasTable("manager_inspections")) {
      const dateCol = firstExistingColumn("manager_inspections", ["inspection_date", "check_date", "created_at"]);
      if (dateCol) {
      candidates.push(
        db.prepare(`
          SELECT DATE(MAX(${dateCol})) AS latest_date
          FROM manager_inspections
          WHERE asset_id = ?
        `).get(assetId)?.latest_date || null
      );
      }
    }
    if (hasTable("tyre_inspections")) {
      const dateCol = firstExistingColumn("tyre_inspections", ["inspection_date", "check_date", "created_at", "date"]);
      if (dateCol) {
      candidates.push(
        db.prepare(`
          SELECT DATE(MAX(${dateCol})) AS latest_date
          FROM tyre_inspections
          WHERE asset_id = ?
        `).get(assetId)?.latest_date || null
      );
      }
    }
    if (hasTable("tire_inspections")) {
      const dateCol = firstExistingColumn("tire_inspections", ["inspection_date", "check_date", "created_at", "date"]);
      if (dateCol) {
      candidates.push(
        db.prepare(`
          SELECT DATE(MAX(${dateCol})) AS latest_date
          FROM tire_inspections
          WHERE asset_id = ?
        `).get(assetId)?.latest_date || null
      );
      }
    }
    const sorted = candidates.filter(Boolean).sort().reverse();
    return sorted[0] || null;
  }

  function resolveWebOrigin(req) {
    const envBase = String(process.env.IRONLOG_PUBLIC_BASE_URL || "").trim().replace(/\/+$/, "");
    if (envBase) return envBase;
    const protoHeader = String(req.headers["x-forwarded-proto"] || "").split(",")[0].trim();
    const hostHeader = String(req.headers["x-forwarded-host"] || req.headers.host || "").split(",")[0].trim();
    const proto = protoHeader || "http";
    if (hostHeader) return `${proto}://${hostHeader}`;
    return "";
  }

  function inferMakeModelFromName(name, code) {
    const raw = String(name || "").trim();
    if (!raw) return { make: null, model: null };
    const parts = raw.split(/\s+/).filter(Boolean);
    if (!parts.length) return { make: null, model: null };
    const make = parts[0] ? String(parts[0]).toUpperCase() : null;
    let model = null;
    if (parts.length >= 2) {
      const second = String(parts[1] || "");
      if (/[0-9]/.test(second) || second.length <= 12) {
        model = second.toUpperCase();
      }
    }
    if (!model) {
      const codeToken = String(code || "").split(/[-_\s]/).find((t) => /[0-9]/.test(t));
      if (codeToken) model = codeToken.toUpperCase();
    }
    return { make: make || null, model: model || null };
  }

  function buildQrProfile(asset, req) {
    const makeCol = firstExistingColumn("assets", ["make", "asset_make", "manufacturer", "brand"]);
    const modelCol = firstExistingColumn("assets", ["model", "asset_model"]);
    let assetMake = null;
    let assetModel = null;
    if (makeCol || modelCol) {
      const fields = [makeCol ? `${makeCol} AS make` : "NULL AS make", modelCol ? `${modelCol} AS model` : "NULL AS model"].join(", ");
      const row = db.prepare(`SELECT ${fields} FROM assets WHERE id = ?`).get(asset.id);
      assetMake = row?.make != null ? String(row.make).trim() || null : null;
      assetModel = row?.model != null ? String(row.model).trim() || null : null;
    }
    if (!assetMake || !assetModel) {
      const inferred = inferMakeModelFromName(asset.asset_name, asset.asset_code);
      if (!assetMake) assetMake = inferred.make;
      if (!assetModel) assetModel = inferred.model;
    }

    const meter = getAssetCurrentHoursInfo(asset.id);
    const currentHours = Number(Number(meter.hours || 0).toFixed(1));

    const planRows = db.prepare(`
      SELECT id, service_name, interval_hours, last_service_hours, active
      FROM maintenance_plans
      WHERE asset_id = ?
        AND active = 1
      ORDER BY interval_hours ASC, id ASC
    `).all(asset.id);
    const nextResolved = planRows.length
      ? resolveNextServiceForAssetPlans(planRows, currentHours, asset.asset_code)
      : null;
    const nextDueMeta = nextResolved
      ? classifyServiceDue(
          nextResolved.remaining_hours,
          asset.asset_code,
          nextResolved.interval_hours,
          50,
        )
      : null;
    const nextService = nextResolved
      ? {
          service_name: nextResolved.service_name,
          next_due_hours: nextResolved.next_due_hours,
          remaining_hours: nextResolved.remaining_hours,
          is_overdue: nextDueMeta?.is_overdue ?? nextResolved.remaining_hours <= 0,
          is_almost_due: nextDueMeta?.is_almost_due ?? false,
          status: nextDueMeta?.status ?? "OK",
          meter_unit: nextDueMeta?.meter_unit ?? "hours",
          near_due_threshold: nextDueMeta?.near_due_threshold ?? 50,
          schedule_mode: nextResolved.schedule_mode,
        }
      : null;

    const machineProfileId = resolveMachinePrestartProfile(asset.category, asset.asset_name, asset.asset_code);
    const machineTemplate = machineProfileId ? getMachinePrestartTemplate(machineProfileId) : null;
    const machine_prestart =
      machineProfileId && machineTemplate
        ? {
            profile_id: machineProfileId,
            check_mode: machinePrestartCheckMode(machineProfileId),
            template_title: String(machineTemplate.title || "Machine pre-start"),
          }
        : null;

    const fuel30d = db.prepare(`
      SELECT
        COALESCE(SUM(liters), 0) AS liters_30d,
        MAX(log_date) AS latest_log_date
      FROM fuel_logs
      WHERE asset_id = ?
        AND log_date >= DATE('now', '-30 day')
    `).get(asset.id);
    const latestFuel = db.prepare(`
      SELECT log_date, liters
      FROM fuel_logs
      WHERE asset_id = ?
      ORDER BY log_date DESC, id DESC
      LIMIT 1
    `).get(asset.id);

    const profile = {
      generated_at: new Date().toISOString(),
      asset: {
        asset_code: asset.asset_code,
        asset_name: asset.asset_name || null,
        category: asset.category || null,
        make: assetMake,
        model: assetModel,
      },
      scan_url: (() => {
        const origin = resolveWebOrigin(req);
        let targetPath = `/web/asset-qr.html?asset_code=${encodeURIComponent(asset.asset_code)}`;
        if (isLdvPrestartQrAsset(asset.asset_code)) {
          targetPath = `/web/ldv-prestart.html?asset_code=${encodeURIComponent(asset.asset_code)}`;
        } else if (machine_prestart?.profile_id) {
          targetPath = `/web/machine-prestart.html?asset_code=${encodeURIComponent(asset.asset_code)}`;
        }
        if (!origin) return targetPath;
        return `${origin}${targetPath}`;
      })(),
      status: buildMachineStatus(asset.id),
      meter: {
        current_hours: currentHours,
        source: meter.source || "unknown",
      },
      next_service_due: nextService
        ? {
            service_name: nextService.service_name,
            next_due_hours: nextService.next_due_hours,
            remaining_hours: nextService.remaining_hours,
            is_overdue: nextService.remaining_hours <= 0,
          }
        : null,
      fuel: {
        liters_last_30_days: Number(Number(fuel30d?.liters_30d || 0).toFixed(1)),
        last_fill_date: latestFuel?.log_date || null,
        last_fill_liters: latestFuel?.liters != null ? Number(Number(latestFuel.liters).toFixed(1)) : null,
      },
      inspections: {
        last_inspection_date: buildInspectionSummary(asset.id),
      },
      machine_prestart,
    };

    const nextDueText = nextService
      ? `${nextService.service_name} at ${nextService.next_due_hours}h (${nextService.remaining_hours}h remaining)`
      : "No active maintenance plan";
    const fuelText = `${profile.fuel.liters_last_30_days}L in last 30 days`;
    const inspectText = profile.inspections.last_inspection_date || "No inspection date";
    const qrText = [
      `IRONLOG ${asset.asset_code}`,
      `Scan URL: ${profile.scan_url}`,
      `Status: ${profile.status}`,
      `Current meter: ${currentHours}h`,
      `Next service: ${nextDueText}`,
      `Fuel: ${fuelText}`,
      `Last inspection: ${inspectText}`,
    ].join("\n");

    return { profile, qrText };
  }

  function buildUndercarriageQrProfile(asset, req) {
    const meter = getAssetCurrentHoursInfo(asset.id);
    const currentHours = Number(Number(meter.hours || 0).toFixed(1));
    const origin = resolveWebOrigin(req);
    const targetPath = `/web/undercarriage-mobile.html?asset_code=${encodeURIComponent(asset.asset_code)}`;
    const scan_url = origin ? `${origin}${targetPath}` : targetPath;

    let lastInspectionDate = null;
    let worstWearPct = null;
    if (hasTable("undercarriage_inspections")) {
      const last = db.prepare(`
        SELECT inspection_date, summary_json
        FROM undercarriage_inspections
        WHERE asset_id = ?
        ORDER BY inspection_date DESC, id DESC
        LIMIT 1
      `).get(asset.id);
      if (last) {
        lastInspectionDate = String(last.inspection_date || "").trim() || null;
        try {
          const summary = JSON.parse(String(last.summary_json || "{}"));
          worstWearPct = summary?.worst_wear_pct ?? null;
        } catch {}
      }
    }

    const profile = {
      purpose: "undercarriage_inspection",
      generated_at: new Date().toISOString(),
      asset: {
        asset_code: asset.asset_code,
        asset_name: asset.asset_name || null,
        category: asset.category || null,
      },
      scan_url,
      meter: {
        current_hours: currentHours,
        source: meter.source || "unknown",
      },
      undercarriage: {
        last_inspection_date: lastInspectionDate,
        worst_wear_pct: worstWearPct != null ? Number(Number(worstWearPct).toFixed(1)) : null,
      },
    };

    const wearText = worstWearPct != null ? `${Number(worstWearPct).toFixed(1)}% max wear` : "No prior inspection";
    const qrText = [
      `IRONLOG UNDERCARRIAGE — ${asset.asset_code}`,
      `Scan to open inspection for this machine only`,
      `Scan URL: ${scan_url}`,
      `SMU: ${currentHours}h`,
      `Last UC inspection: ${lastInspectionDate || "None"} (${wearText})`,
    ].join("\n");

    return { profile, qrText };
  }

  function buildTyreQrProfile(asset, req) {
    const meter = getAssetCurrentHoursInfo(asset.id);
    const currentHours = Number(Number(meter.hours || 0).toFixed(1));
    const origin = resolveWebOrigin(req);
    const targetPath = `/web/tyre-inspection-mobile.html?asset_code=${encodeURIComponent(asset.asset_code)}`;
    const scan_url = origin ? `${origin}${targetPath}` : targetPath;

    let lastInspectionDate = null;
    let lastRunningHours = null;
    if (hasTable("tyre_inspections")) {
      const last = db.prepare(`
        SELECT inspection_date, running_hours
        FROM tyre_inspections
        WHERE asset_id = ?
        ORDER BY inspection_date DESC, id DESC
        LIMIT 1
      `).get(asset.id);
      if (last) {
        lastInspectionDate = String(last.inspection_date || "").trim() || null;
        lastRunningHours = last.running_hours != null ? Number(last.running_hours) : null;
      }
    }

    let assetMake = null;
    let assetModel = null;
    const makeCol = firstExistingColumn("assets", ["make", "asset_make", "manufacturer", "brand"]);
    const modelCol = firstExistingColumn("assets", ["model", "asset_model"]);
    if (makeCol || modelCol) {
      const row = db.prepare(`SELECT ${makeCol ? `${makeCol} AS make` : "NULL AS make"}, ${modelCol ? `${modelCol} AS model` : "NULL AS model"} FROM assets WHERE id = ?`).get(asset.id);
      assetMake = row?.make || null;
      assetModel = row?.model || null;
    }
    if (!assetMake || !assetModel) {
      const inferred = inferMakeModelFromName(asset.asset_name, asset.asset_code);
      if (!assetMake) assetMake = inferred.make;
      if (!assetModel) assetModel = inferred.model;
    }

    const profile = {
      purpose: "tyre_inspection",
      generated_at: new Date().toISOString(),
      asset: {
        id: asset.id,
        asset_code: asset.asset_code,
        asset_name: asset.asset_name || null,
        category: asset.category || null,
        make: assetMake,
        model: assetModel,
      },
      scan_url,
      meter: {
        current_hours: currentHours,
        source: meter.source || "unknown",
      },
      tyres: {
        last_inspection_date: lastInspectionDate,
        last_running_hours: lastRunningHours,
      },
    };

    const lastText = lastInspectionDate
      ? `${lastInspectionDate} @ ${lastRunningHours != null ? `${Number(lastRunningHours).toFixed(1)}h` : "—"}`
      : "No prior inspection";
    const qrText = [
      `IRONLOG TYRE INSPECTION — ${asset.asset_code}`,
      `Scan to open tyre inspection for this machine only`,
      `Scan URL: ${scan_url}`,
      `Current meter: ${currentHours}h`,
      `Last tyre inspection: ${lastText}`,
    ].join("\n");

    return { profile, qrText };
  }

  // Route groups live in routes/assets/. They receive the shared helpers above through ctx.
  const ctx = {
    buildMachineStatus,
    buildQrProfile,
    buildTyreQrProfile,
    buildUndercarriageQrProfile,
    getAssetByCode,
    getAssetById,
    getStoredQrProfile,
    getStoredTyreQrProfile,
    getStoredUndercarriageQrProfile,
    hasTable,
    insertAsset,
    siteCodeFromReq,
    upsertQrProfile,
    upsertTyreQrProfile,
    upsertUndercarriageQrProfile,
  };
  registerRegisterRoutes(app, ctx);
  registerQrProfilesRoutes(app, ctx);
  registerHistoryRoutes(app, ctx);
}
