/**
 * FAMS fuel sync — fuel transactions only.
 * Does NOT write daily_hours or asset master hours (QR pre-start remains authoritative).
 */
import { db } from "../db/client.js";

let lastSyncMemory = {
  status: "idle",
  last_attempt_at: null,
  last_success_at: null,
  last_error: null,
  last_result: null,
};

let syncInFlight = null;

function env(name, fallback = "") {
  const v = process.env[name];
  if (v === undefined || v === null || String(v).trim() === "") return fallback;
  return String(v).trim();
}

function hasColumn(table, col) {
  try {
    return db.prepare(`PRAGMA table_info(${table})`).all().some((r) => String(r.name) === col);
  } catch {
    return false;
  }
}

function ensureColumn(table, colName, colDef) {
  if (!hasColumn(table, colName)) {
    db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colName} ${colDef}`).run();
  }
}

export function ensureFamsFuelSchema() {
  ensureColumn("fuel_logs", "hours_run", "hours_run REAL");
  ensureColumn("fuel_logs", "meter_run_value", "meter_run_value REAL");
  ensureColumn("fuel_logs", "meter_unit", "meter_unit TEXT");
  ensureColumn("fuel_logs", "open_meter_value", "open_meter_value REAL");
  ensureColumn("fuel_logs", "close_meter_value", "close_meter_value REAL");
  ensureColumn("fuel_logs", "unit_cost_per_liter", "unit_cost_per_liter REAL");
  ensureColumn("fuel_logs", "cost_center_code", "cost_center_code TEXT");
  ensureColumn("fuel_logs", "fams_id", "fams_id INTEGER");
  ensureColumn("fuel_logs", "fams_tr_id", "fams_tr_id TEXT");
  ensureColumn("fuel_logs", "fams_equipment_id", "fams_equipment_id INTEGER");
  ensureColumn("fuel_logs", "fams_registration", "fams_registration TEXT");

  try {
    db.prepare(
      `CREATE UNIQUE INDEX IF NOT EXISTS idx_fuel_logs_fams_id ON fuel_logs(fams_id) WHERE fams_id IS NOT NULL`
    ).run();
  } catch (_) {
    // Older SQLite without partial indexes — fall back to plain unique when possible.
    try {
      db.prepare(`CREATE UNIQUE INDEX IF NOT EXISTS idx_fuel_logs_fams_id_all ON fuel_logs(fams_id)`).run();
    } catch (_) {}
  }

  db.exec(`
    CREATE TABLE IF NOT EXISTS fams_unmatched_fuel (
      fams_id INTEGER PRIMARY KEY,
      fams_tr_id TEXT,
      fams_equipment_id INTEGER,
      registration TEXT,
      asset_name TEXT,
      log_date TEXT,
      liters REAL,
      product TEXT,
      store TEXT,
      cost_centre TEXT,
      opening_reading REAL,
      closing_reading REAL,
      total_reading REAL,
      measurement_name TEXT,
      litre_per_hour REAL,
      km_per_litre REAL,
      fuel_price REAL,
      payload_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      resolved_at TEXT
    );

    CREATE TABLE IF NOT EXISTS fams_sync_state (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      enabled INTEGER NOT NULL DEFAULT 1,
      last_attempt_at TEXT,
      last_success_at TEXT,
      last_error TEXT,
      last_received INTEGER,
      last_imported INTEGER,
      last_skipped INTEGER,
      last_unmatched INTEGER,
      last_range_start TEXT,
      last_range_end TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
  `);

  db.prepare(`
    INSERT INTO fams_sync_state (id, enabled)
    SELECT 1, 1
    WHERE NOT EXISTS (SELECT 1 FROM fams_sync_state WHERE id = 1)
  `).run();
}

export function getFamsConfig() {
  const enabledRaw = env("FAMS_ENABLED", "false").toLowerCase();
  const enabled = ["1", "true", "yes", "on"].includes(enabledRaw);
  return {
    enabled,
    baseUrl: env("FAMS_BASE_URL", "https://famsoldapi.azurewebsites.net/Handler1.ashx"),
    client: env("FAMS_CLIENT", "FAMS2018"),
    auth: env("FAMS_AUTH", ""),
    siteId: env("FAMS_SITE_ID", "305"),
    pollMs: Math.max(60_000, Number(env("FAMS_POLL_MS", String(60 * 60 * 1000))) || 60 * 60 * 1000),
  };
}

function ymdSlash(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}/${m}/${day}`;
}

function ymdDash(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

/** First day of current month → today (local calendar). */
export function famsCurrentMonthRange(now = new Date()) {
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return {
    startSlash: ymdSlash(start),
    endSlash: ymdSlash(end),
    startYmd: ymdDash(start),
    endYmd: ymdDash(end),
  };
}

function toDateOnly(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
  if (m) return `${m[1]}-${String(m[2]).padStart(2, "0")}-${String(m[3]).padStart(2, "0")}`;
  return "";
}

function parseNumberLoose(raw, fallback = null) {
  if (raw == null || raw === "") return fallback;
  const n = Number.parseFloat(String(raw).trim().replace(/\s/g, "").replace(/,/g, ""));
  return Number.isFinite(n) ? n : fallback;
}

function normalizeMeterUnit(raw) {
  const s = String(raw || "").trim().toLowerCase();
  if (!s) return null;
  if (s.includes("km")) return "km";
  if (s.includes("hour") || s.includes("hr") || s === "h") return "hours";
  return null;
}

function findAssetByRegistration(registration) {
  const code = String(registration || "").trim();
  if (!code) return null;
  return (
    db
      .prepare(
        `
      SELECT id, asset_code
      FROM assets
      WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?))
      LIMIT 1
    `
      )
      .get(code) || null
  );
}

function famsIdExists(famsId) {
  const id = Number(famsId);
  if (!Number.isFinite(id) || id <= 0) return false;
  const inFuel = db.prepare(`SELECT 1 AS ok FROM fuel_logs WHERE fams_id = ? LIMIT 1`).get(id);
  if (inFuel) return true;
  const inUnmatched = db.prepare(`SELECT 1 AS ok FROM fams_unmatched_fuel WHERE fams_id = ? LIMIT 1`).get(id);
  return Boolean(inUnmatched);
}

/**
 * Earlier FAMS CSV uploads did not retain the source transaction ID. Link a
 * legacy row only when its asset, day, litres and normalized source map to one
 * FAMS transaction exactly; otherwise leave it alone rather than guessing.
 */
function normalizedFamsSource(value) {
  return String(value || "")
    .split("|")
    .map((part) => part.trim())
    .filter((part) => part && !/^(none|n\/?a|not available|-)$/i.test(part))
    .join(" | ")
    .toLowerCase();
}

function linkExactLegacyFamsFuelRow({
  assetId,
  logDate,
  liters,
  source,
  famsId,
  famsTrId,
  famsEquipmentId,
  famsRegistration,
  opening,
  closing,
  fuelPrice,
  costCentre,
}) {
  const sourceKey = normalizedFamsSource(source);
  if (!sourceKey) return false;
  const candidates = db.prepare(`
    SELECT id, source
    FROM fuel_logs
    WHERE asset_id = ?
      AND log_date = ?
      AND fams_id IS NULL
      AND ABS(COALESCE(liters, 0) - ?) < 0.000001
    ORDER BY id ASC
    LIMIT 10
  `).all(assetId, logDate, liters)
    .filter((candidate) => normalizedFamsSource(candidate.source) === sourceKey);
  if (candidates.length !== 1) return false;

  const updated = db.prepare(`
    UPDATE fuel_logs
    SET
      fams_id = ?,
      fams_tr_id = ?,
      fams_equipment_id = ?,
      fams_registration = ?,
      open_meter_value = COALESCE(?, open_meter_value),
      close_meter_value = COALESCE(?, close_meter_value),
      unit_cost_per_liter = COALESCE(?, unit_cost_per_liter),
      cost_center_code = COALESCE(?, cost_center_code)
    WHERE id = ?
      AND fams_id IS NULL
  `).run(
    famsId,
    famsTrId,
    famsEquipmentId,
    famsRegistration,
    opening,
    closing,
    fuelPrice,
    costCentre,
    candidates[0].id,
  );
  return Number(updated.changes || 0) === 1;
}

function confirmedFamsLegacyDuplicateRows({ startDate, endDate, limit = 10000 } = {}) {
  const maxRows = Math.max(1, Math.min(10000, Number(limit) || 10000));
  const rows = db.prepare(`
    SELECT
      f.id,
      f.asset_id,
      f.log_date,
      f.liters,
      f.source,
      f.meter_unit,
      f.meter_run_value,
      f.fams_id,
      a.asset_code,
      a.asset_name
    FROM fuel_logs f
    JOIN assets a ON a.id = f.asset_id
    WHERE f.log_date >= ?
      AND f.log_date <= ?
      AND COALESCE(f.liters, 0) > 0
    ORDER BY f.log_date ASC, a.asset_code ASC, f.id ASC
  `).all(startDate, endDate);
  const groups = new Map();
  for (const row of rows) {
    const sourceKey = normalizedFamsSource(row.source);
    if (!sourceKey) continue;
    const litresKey = Number(row.liters || 0).toFixed(4);
    const key = `${row.asset_id}|${row.log_date}|${litresKey}|${sourceKey}`;
    const group = groups.get(key) || { fams: [], legacy: [] };
    if (row.fams_id != null) group.fams.push(row);
    else group.legacy.push(row);
    groups.set(key, group);
  }
  const confirmed = [];
  for (const group of groups.values()) {
    if (group.fams.length !== 1 || group.legacy.length !== 1) continue;
    const fams = group.fams[0];
    const legacy = group.legacy[0];
    confirmed.push({
      fams_log_id: fams.id,
      legacy_log_id: legacy.id,
      log_date: fams.log_date,
      liters: fams.liters,
      source: fams.source,
      meter_unit: fams.meter_unit,
      meter_run_value: fams.meter_run_value,
      legacy_meter_run_value: legacy.meter_run_value,
      fams_id: fams.fams_id,
      asset_code: fams.asset_code,
      asset_name: fams.asset_name,
    });
  }
  return confirmed.slice(0, maxRows);
}

/** Preview only unique, unambiguous FAMS-versus-legacy duplicate pairs. */
export function previewFamsLegacyDuplicates({ startDate, endDate, limit = 100 } = {}) {
  ensureFamsFuelSchema();
  const rows = confirmedFamsLegacyDuplicateRows({ startDate, endDate, limit });
  return { found: rows.length, rows };
}

/** Keep the FAMS-tagged row and remove its confirmed legacy duplicate. */
export function removeFamsLegacyDuplicates({ startDate, endDate } = {}) {
  ensureFamsFuelSchema();
  const rows = confirmedFamsLegacyDuplicateRows({ startDate, endDate, limit: 10000 });
  const remove = db.prepare(`DELETE FROM fuel_logs WHERE id = ? AND fams_id IS NULL`);
  const tx = db.transaction((matches) => {
    let removed = 0;
    for (const match of matches) removed += Number(remove.run(match.legacy_log_id).changes || 0);
    return removed;
  });
  return { found: rows.length, removed: tx(rows), rows: rows.slice(0, 100) };
}

function persistSyncState(patch) {
  ensureFamsFuelSchema();
  const cur = db.prepare(`SELECT * FROM fams_sync_state WHERE id = 1`).get() || {};
  db.prepare(
    `
    UPDATE fams_sync_state SET
      last_attempt_at = ?,
      last_success_at = ?,
      last_error = ?,
      last_received = ?,
      last_imported = ?,
      last_skipped = ?,
      last_unmatched = ?,
      last_range_start = ?,
      last_range_end = ?,
      updated_at = datetime('now')
    WHERE id = 1
  `
  ).run(
    patch.last_attempt_at ?? cur.last_attempt_at ?? null,
    patch.last_success_at ?? cur.last_success_at ?? null,
    patch.last_error !== undefined ? patch.last_error : cur.last_error ?? null,
    patch.last_received ?? cur.last_received ?? null,
    patch.last_imported ?? cur.last_imported ?? null,
    patch.last_skipped ?? cur.last_skipped ?? null,
    patch.last_unmatched ?? cur.last_unmatched ?? null,
    patch.last_range_start ?? cur.last_range_start ?? null,
    patch.last_range_end ?? cur.last_range_end ?? null
  );
}

export function getFamsSyncStatus() {
  ensureFamsFuelSchema();
  const cfg = getFamsConfig();
  const row = db.prepare(`SELECT * FROM fams_sync_state WHERE id = 1`).get() || {};
  const unmatchedOpen = Number(
    db.prepare(`SELECT COUNT(*) AS c FROM fams_unmatched_fuel WHERE resolved_at IS NULL`).get()?.c || 0
  );
  let status = "disabled";
  if (cfg.enabled) {
    if (lastSyncMemory.status === "running") status = "syncing";
    else if (row.last_error) status = "error";
    else if (row.last_success_at) status = "connected";
    else status = "idle";
  }
  return {
    ok: true,
    enabled: cfg.enabled,
    status,
    configured: Boolean(cfg.auth),
    site_id: cfg.siteId,
    base_url: cfg.baseUrl,
    // Never expose FAMS_AUTH
    poll_ms: cfg.pollMs,
    last_attempt_at: row.last_attempt_at || lastSyncMemory.last_attempt_at,
    last_success_at: row.last_success_at || lastSyncMemory.last_success_at,
    last_error: row.last_error || lastSyncMemory.last_error,
    last_received: row.last_received,
    last_imported: row.last_imported,
    last_skipped: row.last_skipped,
    last_unmatched: row.last_unmatched,
    last_range_start: row.last_range_start,
    last_range_end: row.last_range_end,
    unmatched_open: unmatchedOpen,
    note: "FAMS supplies fuel only. Equipment hours remain from QR pre-start / daily hours.",
  };
}

export async function fetchFamsReportingLogbook({ startSlash, endSlash, log = console } = {}) {
  const cfg = getFamsConfig();
  if (!cfg.auth) throw new Error("FAMS_AUTH is not configured");
  const url = new URL(cfg.baseUrl);
  url.searchParams.set("Client", cfg.client);
  url.searchParams.set("Auth", cfg.auth);
  url.searchParams.set("action", "ArrayDataSetToJSON");
  url.searchParams.set("queryType", "SqlSp");
  url.searchParams.set("sp", "get_ReportinglogbookRev2");
  url.searchParams.set("paramlist", `${cfg.siteId},${startSlash},${endSlash}`);

  const res = await fetch(url.toString(), {
    method: "GET",
    headers: { Accept: "application/json" },
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`FAMS HTTP ${res.status}${txt ? `: ${txt.slice(0, 200)}` : ""}`);
  }
  const data = await res.json();
  if (!Array.isArray(data)) {
    // Some handlers wrap in { Table: [...] } or similar
    if (Array.isArray(data?.Table)) return data.Table;
    if (Array.isArray(data?.data)) return data.data;
    if (Array.isArray(data?.rows)) return data.rows;
    throw new Error("FAMS response was not a JSON array");
  }
  log.info?.(`[FAMS] ${data.length} records received`);
  return data;
}

function insertMatchedFuelRow(row, asset) {
  const famsId = Number(row.ID);
  const registration = String(row.Registration || "").trim();
  const logDate = toDateOnly(row.DateOnly || row.Date || row.ClosingReadingDate);
  const liters = parseNumberLoose(row.Volume, 0);
  if (!logDate || !(liters > 0) || !Number.isFinite(famsId)) return false;

  const opening = parseNumberLoose(row.OpeningReading, null);
  const closing = parseNumberLoose(row.ClosingReading, null);
  const totalReading = parseNumberLoose(row.TotalReading, null);
  const meterUnit = normalizeMeterUnit(row.Measurement_Name);
  const fuelPrice = parseNumberLoose(row.FuelPrice, null);
  const store = String(row.Store || "").trim();
  const operator = String(row.Operator || "").trim();
  const driver = String(row.Driver || "").trim();
  const source =
    [store, operator, driver].filter(Boolean).join(" | ") || "FAMS";
  const costCentre = String(row.Costcentre || row.CostcentreDesc || "").trim() || null;
  const famsTrId = String(row.TrId || "").trim() || null;
  const famsEquipmentId = row.Equipment_ID != null ? Number(row.Equipment_ID) : null;

  if (linkExactLegacyFamsFuelRow({
    assetId: asset.id,
    logDate,
    liters,
    source,
    famsId,
    famsTrId,
    famsEquipmentId,
    famsRegistration: registration || null,
    opening,
    closing,
    fuelPrice: fuelPrice != null && fuelPrice > 0 ? fuelPrice : null,
    costCentre,
  })) {
    return "linked_legacy";
  }

  // Store FAMS meter fields on fuel_logs for consumption analysis only.
  // Never writes daily_hours / asset master hours.
  db.prepare(
    `
    INSERT INTO fuel_logs (
      asset_id, log_date, liters, source,
      hours_run, meter_unit, meter_run_value,
      open_meter_value, close_meter_value,
      unit_cost_per_liter, cost_center_code,
      fams_id, fams_tr_id, fams_equipment_id, fams_registration
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `
  ).run(
    asset.id,
    logDate,
    liters,
    source,
    totalReading,
    meterUnit,
    totalReading,
    opening,
    closing,
    fuelPrice != null && fuelPrice > 0 ? fuelPrice : null,
    costCentre,
    famsId,
    famsTrId,
    famsEquipmentId,
    registration || null
  );
  return "inserted";
}

function insertUnmatchedFuelRow(row) {
  const famsId = Number(row.ID);
  const registration = String(row.Registration || "").trim();
  const logDate = toDateOnly(row.DateOnly || row.Date || row.ClosingReadingDate);
  const liters = parseNumberLoose(row.Volume, 0);
  if (!Number.isFinite(famsId) || famsId <= 0) return false;

  db.prepare(
    `
    INSERT OR IGNORE INTO fams_unmatched_fuel (
      fams_id, fams_tr_id, fams_equipment_id, registration, asset_name,
      log_date, liters, product, store, cost_centre,
      opening_reading, closing_reading, total_reading, measurement_name,
      litre_per_hour, km_per_litre, fuel_price, payload_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `
  ).run(
    famsId,
    String(row.TrId || "").trim() || null,
    row.Equipment_ID != null ? Number(row.Equipment_ID) : null,
    registration || null,
    String(row.Name || "").trim() || null,
    logDate || null,
    liters > 0 ? liters : null,
    String(row.Product || "").trim() || null,
    String(row.Store || "").trim() || null,
    String(row.Costcentre || "").trim() || null,
    parseNumberLoose(row.OpeningReading, null),
    parseNumberLoose(row.ClosingReading, null),
    parseNumberLoose(row.TotalReading, null),
    String(row.Measurement_Name || "").trim() || null,
    parseNumberLoose(row.LitrePerHour, null),
    parseNumberLoose(row.KMPerLitre, null),
    parseNumberLoose(row.FuelPrice, null),
    JSON.stringify({
      ID: row.ID,
      TrId: row.TrId,
      Equipment_ID: row.Equipment_ID,
      Registration: row.Registration,
      Name: row.Name,
      Date: row.Date,
      DateOnly: row.DateOnly,
      Volume: row.Volume,
      Store_ID: row.Store_ID,
      Store: row.Store,
      Costcentre: row.Costcentre,
      CostcentreDesc: row.CostcentreDesc,
      Totalizer: row.Totalizer,
    })
  );
  return true;
}

/**
 * Reusable FAMS sync: current month → today.
 * Safe to call from scheduler or manual button.
 */
export async function syncFamsFuel({ log = console, force = false, range = null } = {}) {
  ensureFamsFuelSchema();
  const cfg = getFamsConfig();
  if (!cfg.enabled && !force) {
    log.info?.("[FAMS] Sync skipped (FAMS_ENABLED is off)");
    return { ok: true, skipped: true, reason: "disabled", ...getFamsSyncStatus() };
  }
  if (!cfg.auth) {
    const err = "FAMS_AUTH is not configured";
    lastSyncMemory = { ...lastSyncMemory, status: "error", last_error: err, last_attempt_at: new Date().toISOString() };
    persistSyncState({ last_attempt_at: new Date().toISOString(), last_error: err });
    log.warn?.(`[FAMS] Sync failed: ${err}`);
    return { ok: false, error: err };
  }

  if (syncInFlight) {
    log.info?.("[FAMS] Sync already in progress — reusing in-flight promise");
    return syncInFlight;
  }

  syncInFlight = (async () => {
    const attemptAt = new Date().toISOString();
    lastSyncMemory = { ...lastSyncMemory, status: "running", last_attempt_at: attemptAt, last_error: null };
    persistSyncState({ last_attempt_at: attemptAt, last_error: null });
    log.info?.("[FAMS] Sync started");

    const syncRange = range || famsCurrentMonthRange();
    try {
      const rows = await fetchFamsReportingLogbook({
        startSlash: syncRange.startSlash,
        endSlash: syncRange.endSlash,
        log,
      });

      let imported = 0;
      let skipped = 0;
      let unmatched = 0;
      let invalid = 0;
      let linkedLegacy = 0;

      const tx = db.transaction((records) => {
        for (const row of records) {
          const famsId = Number(row?.ID);
          if (!Number.isFinite(famsId) || famsId <= 0) {
            invalid += 1;
            continue;
          }
          if (famsIdExists(famsId)) {
            skipped += 1;
            continue;
          }
          const registration = String(row.Registration || "").trim();
          const liters = parseNumberLoose(row.Volume, 0);
          if (!(liters > 0)) {
            invalid += 1;
            continue;
          }
          const asset = findAssetByRegistration(registration);
          if (asset) {
            const insertResult = insertMatchedFuelRow(row, asset);
            if (insertResult) {
              if (insertResult === "linked_legacy") linkedLegacy += 1;
              else imported += 1;
            } else invalid += 1;
          } else {
            if (insertUnmatchedFuelRow(row)) unmatched += 1;
            else invalid += 1;
          }
        }
      });
      tx(rows);

      const successAt = new Date().toISOString();
      const result = {
        ok: true,
        range: { start: syncRange.startYmd, end: syncRange.endYmd },
        received: rows.length,
        imported,
        linked_legacy: linkedLegacy,
        skipped,
        unmatched,
        invalid,
      };
      lastSyncMemory = {
        status: "connected",
        last_attempt_at: attemptAt,
        last_success_at: successAt,
        last_error: null,
        last_result: result,
      };
      persistSyncState({
        last_attempt_at: attemptAt,
        last_success_at: successAt,
        last_error: null,
        last_received: rows.length,
        last_imported: imported,
        last_skipped: skipped,
        last_unmatched: unmatched,
        last_range_start: syncRange.startYmd,
        last_range_end: syncRange.endYmd,
      });
      log.info?.(`[FAMS] ${imported} new transactions imported`);
      log.info?.(`[FAMS] ${skipped} existing transactions skipped`);
      log.info?.(`[FAMS] ${unmatched} unmatched assets`);
      log.info?.("[FAMS] Sync completed");
      return result;
    } catch (err) {
      const msg = err?.message || String(err);
      lastSyncMemory = {
        ...lastSyncMemory,
        status: "error",
        last_attempt_at: attemptAt,
        last_error: msg,
      };
      persistSyncState({
        last_attempt_at: attemptAt,
        last_error: msg,
        last_range_start: syncRange.startYmd,
        last_range_end: syncRange.endYmd,
      });
      log.warn?.(`[FAMS] Sync failed: ${msg}`);
      return { ok: false, error: msg, range: { start: syncRange.startYmd, end: syncRange.endYmd } };
    } finally {
      syncInFlight = null;
    }
  })();

  return syncInFlight;
}

export function listFamsUnmatched({ limit = 200 } = {}) {
  ensureFamsFuelSchema();
  const lim = Math.max(1, Math.min(1000, Number(limit) || 200));
  return db
    .prepare(
      `
    SELECT fams_id, registration, asset_name, log_date, liters, store, cost_centre, created_at
    FROM fams_unmatched_fuel
    WHERE resolved_at IS NULL
    ORDER BY log_date DESC, fams_id DESC
    LIMIT ?
  `
    )
    .all(lim);
}

export async function runFamsAutoScheduler(log = console) {
  ensureFamsFuelSchema();
  const cfg = getFamsConfig();
  if (!cfg.enabled) {
    log.info?.("[FAMS] auto sync disabled (set FAMS_ENABLED=true)");
    return null;
  }
  if (!cfg.auth) {
    log.warn?.("[FAMS] auto sync enabled but FAMS_AUTH is missing");
    return null;
  }

  // Startup sync shortly after boot (non-fatal).
  setTimeout(() => {
    syncFamsFuel({ log }).catch((err) => {
      log.warn?.(`[FAMS] startup sync failed: ${err?.message || err}`);
    });
  }, 8_000);

  const timer = setInterval(() => {
    syncFamsFuel({ log }).catch((err) => {
      log.warn?.(`[FAMS] hourly sync failed: ${err?.message || err}`);
    });
  }, cfg.pollMs);

  log.info?.(`[FAMS] auto sync every ${Math.round(cfg.pollMs / 1000)}s (month-to-date → today)`);
  return timer;
}
