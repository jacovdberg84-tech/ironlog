// IRONLOG telematics — device registry, live snapshots, fault events, daily_hours sync.
// FSC650 / FSC150 units post normalized payloads to /api/telematics/ingest.

import { db } from "../db/client.js";

const PILOT_ASSETS = [
  { asset_code: "A300AM", unit_model: "FSC650" },
  { asset_code: "F500AM", unit_model: "FSC150" },
];

function hasColumn(table, col) {
  const rows = db.prepare(`PRAGMA table_info(${table})`).all();
  return rows.some((r) => String(r.name || "") === col);
}

export function ensureTelematicsTables() {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS telematics_devices (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL UNIQUE,
      device_serial TEXT NOT NULL UNIQUE,
      unit_model TEXT NOT NULL,
      external_id TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`
    CREATE INDEX IF NOT EXISTS idx_telematics_devices_serial ON telematics_devices(device_serial)
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS telematics_snapshots (
      asset_id INTEGER PRIMARY KEY,
      device_serial TEXT NOT NULL,
      recorded_at TEXT NOT NULL,
      engine_hours REAL,
      run_hours REAL,
      idle_hours REAL,
      run_seconds_today REAL,
      idle_seconds_today REAL,
      ignition_on INTEGER,
      latitude REAL,
      longitude REAL,
      speed_kmh REAL,
      active_fault_count INTEGER NOT NULL DEFAULT 0,
      payload_json TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS telematics_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      device_serial TEXT NOT NULL,
      event_time TEXT NOT NULL,
      event_type TEXT NOT NULL DEFAULT 'fault',
      fault_code TEXT,
      description TEXT,
      severity TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      cleared_at TEXT,
      payload_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`
    CREATE INDEX IF NOT EXISTS idx_telematics_events_asset_time
    ON telematics_events(asset_id, event_time DESC)
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS telematics_ingest_log (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      device_serial TEXT,
      asset_code TEXT,
      status TEXT NOT NULL,
      detail TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  if (!hasColumn("daily_hours", "meter_source")) {
    db.prepare(`ALTER TABLE daily_hours ADD COLUMN meter_source TEXT DEFAULT 'manual'`).run();
  }
}

export function ensurePilotDevices() {
  ensureTelematicsTables();
  const upsert = db.prepare(`
    INSERT INTO telematics_devices (asset_id, device_serial, unit_model, external_id, active, updated_at)
    VALUES (?, ?, ?, ?, 1, datetime('now'))
    ON CONFLICT(asset_id) DO UPDATE SET
      unit_model = excluded.unit_model,
      active = 1,
      updated_at = datetime('now')
  `);
  for (const pilot of PILOT_ASSETS) {
    const asset = db.prepare(`SELECT id FROM assets WHERE asset_code = ?`).get(pilot.asset_code);
    if (!asset) continue;
    const serial = `FSC-${pilot.asset_code}`;
    upsert.run(asset.id, serial, pilot.unit_model, serial);
  }
}

export function getDeviceBySerial(deviceSerial) {
  return db.prepare(`
    SELECT d.*, a.asset_code, a.asset_name, a.active AS asset_active
    FROM telematics_devices d
    JOIN assets a ON a.id = d.asset_id
    WHERE d.device_serial = ? AND d.active = 1
  `).get(String(deviceSerial || "").trim());
}

export function getDeviceByAssetCode(assetCode) {
  return db.prepare(`
    SELECT d.*, a.asset_code, a.asset_name, a.active AS asset_active
    FROM telematics_devices d
    JOIN assets a ON a.id = d.asset_id
    WHERE LOWER(a.asset_code) = LOWER(?) AND d.active = 1
  `).get(String(assetCode || "").trim());
}

export function isTelematicsAsset(assetId) {
  const row = db.prepare(`
    SELECT 1 AS ok FROM telematics_devices WHERE asset_id = ? AND active = 1 LIMIT 1
  `).get(assetId);
  return Boolean(row?.ok);
}

/** Lock manual hourmeters only when a live FSC feed exists (not merely registered). */
export function isTelematicsMeterLocked(assetId) {
  const row = db.prepare(`
    SELECT 1 AS ok
    FROM telematics_snapshots s
    JOIN telematics_devices d ON d.asset_id = s.asset_id AND d.active = 1
    WHERE s.asset_id = ?
      AND s.engine_hours IS NOT NULL
      AND datetime(s.updated_at) >= datetime('now', '-24 hours')
    LIMIT 1
  `).get(assetId);
  return Boolean(row?.ok);
}

function prevDate(dateStr) {
  const d = new Date(`${dateStr}T12:00:00`);
  d.setDate(d.getDate() - 1);
  return d.toISOString().slice(0, 10);
}

function numOrNull(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

export function syncTelematicsDailyHours(assetId, workDate) {
  const snap = db.prepare(`SELECT * FROM telematics_snapshots WHERE asset_id = ?`).get(assetId);
  if (!snap) return { ok: false, reason: "no_snapshot" };

  const closing = numOrNull(snap.engine_hours);
  if (closing == null) return { ok: false, reason: "no_engine_hours" };

  const yDate = prevDate(workDate);
  const yRow = db.prepare(`
    SELECT closing_hours, scheduled_hours, is_used
    FROM daily_hours
    WHERE asset_id = ? AND work_date = ?
  `).get(assetId, yDate);

  let opening = numOrNull(yRow?.closing_hours);
  let hoursRun = null;
  const runSecToday = numOrNull(snap.run_seconds_today);
  if (runSecToday != null && runSecToday >= 0) {
    hoursRun = Number((runSecToday / 3600).toFixed(2));
  }
  if (opening == null && hoursRun != null) {
    opening = Number((closing - hoursRun).toFixed(2));
  }
  if (opening == null) opening = closing;
  if (hoursRun == null) hoursRun = Number(Math.max(0, closing - opening).toFixed(2));

  const scheduled = numOrNull(yRow?.scheduled_hours) ?? 0;
  const isUsed = yRow?.is_used != null ? Number(yRow.is_used) : 1;
  const notes = `Telematics sync ${snap.recorded_at || ""}`.trim();

  db.prepare(`
    INSERT INTO daily_hours (
      asset_id, work_date,
      scheduled_hours, opening_hours, closing_hours,
      hours_run, input_unit, is_used, operator, notes, meter_source
    )
    VALUES (?, ?, ?, ?, ?, ?, 'hours', ?, NULL, ?, 'telematics')
    ON CONFLICT(asset_id, work_date) DO UPDATE SET
      opening_hours = excluded.opening_hours,
      closing_hours = excluded.closing_hours,
      hours_run = excluded.hours_run,
      meter_source = 'telematics',
      notes = CASE
        WHEN daily_hours.meter_source = 'telematics' THEN excluded.notes
        ELSE COALESCE(daily_hours.notes, '') || ' | Overwritten by telematics'
      END
  `).run(assetId, workDate, scheduled, opening, closing, hoursRun, isUsed, notes);

  return {
    ok: true,
    asset_id: assetId,
    work_date: workDate,
    opening_hours: opening,
    closing_hours: closing,
    hours_run: hoursRun,
    meter_source: "telematics",
  };
}

function upsertFaultEvents(assetId, deviceSerial, eventTime, faults) {
  if (!Array.isArray(faults) || !faults.length) return 0;
  const insert = db.prepare(`
    INSERT INTO telematics_events (
      asset_id, device_serial, event_time, event_type,
      fault_code, description, severity, active, payload_json
    )
    VALUES (?, ?, ?, 'fault', ?, ?, ?, ?, ?)
  `);
  let count = 0;
  for (const f of faults) {
    const code = String(f?.code || f?.fault_code || "").trim();
    if (!code) continue;
    const active = f?.active === false || f?.cleared === true ? 0 : 1;
    insert.run(
      assetId,
      deviceSerial,
      eventTime,
      code,
      String(f?.description || f?.message || "").trim() || null,
      String(f?.severity || "warning").trim().toLowerCase(),
      active,
      JSON.stringify(f)
    );
    count += 1;
  }
  return count;
}

export function ingestTelematicsPayload(payload) {
  ensureTelematicsTables();
  const body = payload && typeof payload === "object" ? payload : {};
  const deviceSerial = String(body.device_id || body.device_serial || "").trim();
  const assetCode = String(body.asset_code || "").trim();
  let device = deviceSerial ? getDeviceBySerial(deviceSerial) : null;
  if (!device && assetCode) device = getDeviceByAssetCode(assetCode);
  if (!device) {
    db.prepare(`
      INSERT INTO telematics_ingest_log (device_serial, asset_code, status, detail)
      VALUES (?, ?, 'rejected', ?)
    `).run(deviceSerial || null, assetCode || null, "Unknown device or asset — register pilot device first");
    return { ok: false, error: "Unknown telematics device or asset" };
  }

  const recordedAt = String(body.recorded_at || body.timestamp || new Date().toISOString()).trim();
  const engineHours = numOrNull(body.engine_hours ?? body.meter_hours ?? body.hours);
  const runHours = numOrNull(body.run_hours);
  const idleHours = numOrNull(body.idle_hours);
  const runSecondsToday = numOrNull(body.run_seconds_today ?? body.run_time_today_sec);
  const idleSecondsToday = numOrNull(body.idle_seconds_today ?? body.idle_time_today_sec);
  const ignitionOn = body.ignition_on === true || body.ignition === true || body.ignition === 1 ? 1 : 0;
  const latitude = numOrNull(body.latitude ?? body.lat);
  const longitude = numOrNull(body.longitude ?? body.lng ?? body.lon);
  const speedKmh = numOrNull(body.speed_kmh ?? body.speed);
  const faults = Array.isArray(body.fault_codes) ? body.fault_codes
    : Array.isArray(body.faults) ? body.faults
      : Array.isArray(body.dtc) ? body.dtc
        : [];
  const activeFaultCount = faults.filter((f) => f?.active !== false && f?.cleared !== true).length;

  db.prepare(`
    INSERT INTO telematics_snapshots (
      asset_id, device_serial, recorded_at,
      engine_hours, run_hours, idle_hours,
      run_seconds_today, idle_seconds_today,
      ignition_on, latitude, longitude, speed_kmh,
      active_fault_count, payload_json, updated_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
    ON CONFLICT(asset_id) DO UPDATE SET
      device_serial = excluded.device_serial,
      recorded_at = excluded.recorded_at,
      engine_hours = excluded.engine_hours,
      run_hours = excluded.run_hours,
      idle_hours = excluded.idle_hours,
      run_seconds_today = excluded.run_seconds_today,
      idle_seconds_today = excluded.idle_seconds_today,
      ignition_on = excluded.ignition_on,
      latitude = excluded.latitude,
      longitude = excluded.longitude,
      speed_kmh = excluded.speed_kmh,
      active_fault_count = excluded.active_fault_count,
      payload_json = excluded.payload_json,
      updated_at = datetime('now')
  `).run(
    device.asset_id,
    device.device_serial,
    recordedAt,
    engineHours,
    runHours,
    idleHours,
    runSecondsToday,
    idleSecondsToday,
    ignitionOn,
    latitude,
    longitude,
    speedKmh,
    activeFaultCount,
    JSON.stringify(body)
  );

  const faultCount = upsertFaultEvents(device.asset_id, device.device_serial, recordedAt, faults);
  const workDate = recordedAt.slice(0, 10);
  const daily = syncTelematicsDailyHours(device.asset_id, workDate);

  db.prepare(`
    INSERT INTO telematics_ingest_log (device_serial, asset_code, status, detail)
    VALUES (?, ?, 'ok', ?)
  `).run(device.device_serial, device.asset_code, `snapshot updated; faults=${faultCount}; daily=${daily.ok ? "synced" : daily.reason}`);

  return {
    ok: true,
    asset_code: device.asset_code,
    device_serial: device.device_serial,
    recorded_at: recordedAt,
    faults_recorded: faultCount,
    daily_hours: daily,
  };
}

export function listFleetSnapshots() {
  ensureTelematicsTables();
  return db.prepare(`
    SELECT
      a.asset_code,
      a.asset_name,
      d.unit_model,
      d.device_serial,
      s.recorded_at,
      s.engine_hours,
      s.run_hours,
      s.idle_hours,
      s.run_seconds_today,
      s.idle_seconds_today,
      s.ignition_on,
      s.latitude,
      s.longitude,
      s.speed_kmh,
      s.active_fault_count,
      s.updated_at,
      CASE
        WHEN s.updated_at IS NULL THEN 'offline'
        WHEN datetime(s.updated_at) < datetime('now', '-15 minutes') THEN 'stale'
        ELSE 'live'
      END AS link_status
    FROM telematics_devices d
    JOIN assets a ON a.id = d.asset_id
    LEFT JOIN telematics_snapshots s ON s.asset_id = d.asset_id
    WHERE d.active = 1 AND a.active = 1
    ORDER BY a.asset_code ASC
  `).all();
}

export function listRecentFaults(limit = 50) {
  ensureTelematicsTables();
  return db.prepare(`
    SELECT
      e.id,
      e.event_time,
      e.fault_code,
      e.description,
      e.severity,
      e.active,
      a.asset_code,
      a.asset_name,
      d.unit_model
    FROM telematics_events e
    JOIN assets a ON a.id = e.asset_id
    LEFT JOIN telematics_devices d ON d.asset_id = e.asset_id
    WHERE e.event_type = 'fault'
    ORDER BY e.event_time DESC, e.id DESC
    LIMIT ?
  `).all(Math.max(1, Math.min(500, Number(limit) || 50)));
}
