// Cartrack Fleet API (Mozambique) — live fleet status, events, morning speeding reports.

import crypto from "node:crypto";
import nodemailer from "nodemailer";
import { db } from "../db/client.js";

const DEFAULT_BASE_URL = "https://fleetapi-mz.cartrack.com/rest";

const SPEED_EVENT_RE = /speed|overspeed|speeding|excess/i;

function configSecret() {
  const raw = String(process.env.IRONLOG_SMTP_SECRET || process.env.IRONLOG_AUTH_SECRET || "").trim();
  return raw || "IRONLOG_SMTP_DEFAULT_SECRET_CHANGE_ME";
}

function encryptSecret(plain) {
  const iv = crypto.randomBytes(12);
  const key = crypto.createHash("sha256").update(configSecret()).digest();
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  const encrypted = Buffer.concat([cipher.update(String(plain || ""), "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return `${iv.toString("base64")}.${tag.toString("base64")}.${encrypted.toString("base64")}`;
}

function decryptSecret(cipherText) {
  const raw = String(cipherText || "").trim();
  if (!raw) return "";
  const [ivB64, tagB64, encB64] = raw.split(".");
  if (!ivB64 || !tagB64 || !encB64) return null;
  try {
    const iv = Buffer.from(ivB64, "base64");
    const tag = Buffer.from(tagB64, "base64");
    const enc = Buffer.from(encB64, "base64");
    const key = crypto.createHash("sha256").update(configSecret()).digest();
    const decipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
    decipher.setAuthTag(tag);
    return Buffer.concat([decipher.update(enc), decipher.final()]).toString("utf8");
  } catch {
    return null;
  }
}

export function ensureCartrackTables() {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS cartrack_settings (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      base_url TEXT,
      username TEXT,
      password_enc TEXT,
      morning_recipients TEXT,
      morning_enabled INTEGER NOT NULL DEFAULT 1,
      morning_hour INTEGER NOT NULL DEFAULT 6,
      morning_minute INTEGER NOT NULL DEFAULT 0,
      updated_by TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`INSERT INTO cartrack_settings (id) SELECT 1 WHERE NOT EXISTS (SELECT 1 FROM cartrack_settings WHERE id = 1)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS cartrack_vehicle_snapshots (
      registration TEXT PRIMARY KEY,
      asset_code TEXT,
      vehicle_id TEXT,
      vehicle_name TEXT,
      ignition_on INTEGER,
      speed_kmh REAL,
      latitude REAL,
      longitude REAL,
      odometer_km REAL,
      last_event_at TEXT,
      status_json TEXT,
      synced_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS cartrack_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      event_id TEXT,
      registration TEXT NOT NULL,
      asset_code TEXT,
      event_type TEXT,
      event_type_label TEXT,
      event_time TEXT NOT NULL,
      speed_kmh REAL,
      speed_limit_kmh REAL,
      latitude REAL,
      longitude REAL,
      driver_name TEXT,
      description TEXT,
      is_speeding INTEGER NOT NULL DEFAULT 0,
      payload_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(event_id)
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_cartrack_events_time ON cartrack_events(event_time DESC)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_cartrack_events_reg ON cartrack_events(registration, event_time DESC)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS cartrack_morning_reports (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      report_date TEXT NOT NULL,
      summary_json TEXT NOT NULL,
      emailed_at TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(report_date)
    )
  `).run();
}

function settingsRow() {
  ensureCartrackTables();
  return db.prepare(`SELECT * FROM cartrack_settings WHERE id = 1`).get();
}

export function getCartrackPublicSettings() {
  const row = settingsRow() || {};
  const envUser = String(process.env.CARTRACK_USERNAME || "").trim();
  const envBase = String(process.env.CARTRACK_BASE_URL || "").trim();
  const hasEnv = Boolean(envUser && String(process.env.CARTRACK_PASSWORD || "").trim());
  const hasDb = Boolean(String(row.username || "").trim() && String(row.password_enc || "").trim());
  return {
    configured: hasEnv || hasDb,
    source: hasEnv ? "env" : hasDb ? "database" : "none",
    base_url: envBase || String(row.base_url || "").trim() || DEFAULT_BASE_URL,
    username: hasEnv ? envUser : String(row.username || "").trim(),
    has_password: hasEnv || Boolean(String(row.password_enc || "").trim()),
    morning_enabled: Number(row.morning_enabled ?? 1) === 1,
    morning_time: `${String(row.morning_hour ?? 6).padStart(2, "0")}:${String(row.morning_minute ?? 0).padStart(2, "0")}`,
    morning_recipients: String(row.morning_recipients || process.env.CARTRACK_MORNING_RECIPIENTS || "").trim(),
    updated_at: row.updated_at || null,
    updated_by: row.updated_by || null,
  };
}

export function saveCartrackSettings({ base_url, username, password, morning_recipients, morning_enabled, morning_hour, morning_minute, updated_by }) {
  ensureCartrackTables();
  const existing = settingsRow() || {};
  const nextPassEnc = (() => {
    if (password != null && String(password).trim() !== "") return encryptSecret(String(password).trim());
    return String(existing.password_enc || "");
  })();
  db.prepare(`
    UPDATE cartrack_settings
    SET base_url = ?, username = ?, password_enc = ?,
        morning_recipients = ?, morning_enabled = ?, morning_hour = ?, morning_minute = ?,
        updated_by = ?, updated_at = datetime('now')
    WHERE id = 1
  `).run(
    String(base_url || DEFAULT_BASE_URL).trim() || DEFAULT_BASE_URL,
    String(username || "").trim(),
    nextPassEnc,
    String(morning_recipients || "").trim(),
    morning_enabled === false || String(morning_enabled) === "0" ? 0 : 1,
    Math.max(0, Math.min(23, Number(morning_hour ?? existing.morning_hour ?? 6))),
    Math.max(0, Math.min(59, Number(morning_minute ?? existing.morning_minute ?? 0))),
    String(updated_by || "admin").trim()
  );
  return getCartrackPublicSettings();
}

function getCartrackCredentials() {
  const envUser = String(process.env.CARTRACK_USERNAME || "").trim();
  const envPass = String(process.env.CARTRACK_PASSWORD || "").trim();
  if (envUser && envPass) {
    return {
      baseUrl: String(process.env.CARTRACK_BASE_URL || DEFAULT_BASE_URL).trim().replace(/\/$/, ""),
      username: envUser,
      password: envPass,
    };
  }
  const row = settingsRow();
  if (!row?.username || !row.password_enc) return null;
  const password = decryptSecret(row.password_enc);
  if (!password) return null;
  return {
    baseUrl: String(row.base_url || DEFAULT_BASE_URL).trim().replace(/\/$/, ""),
    username: String(row.username).trim(),
    password,
  };
}

function normalizeRegistration(reg) {
  return String(reg || "").trim().toUpperCase().replace(/\s+/g, "");
}

function resolveAssetCode(registration, vehicleName = "") {
  const reg = normalizeRegistration(registration);
  if (!reg) return null;
  const byCode = db.prepare(`
    SELECT asset_code FROM assets
    WHERE UPPER(REPLACE(asset_code, ' ', '')) = ?
      AND archived = 0
    LIMIT 1
  `).get(reg);
  if (byCode?.asset_code) return String(byCode.asset_code);
  const hay = `${registration} ${vehicleName}`.trim();
  if (!hay) return null;
  const assets = db.prepare(`SELECT asset_code, asset_name FROM assets WHERE archived = 0`).all();
  const low = hay.toLowerCase();
  for (const a of assets) {
    const code = String(a.asset_code || "");
    const name = String(a.asset_name || "");
    if (code && low.includes(code.toLowerCase())) return code;
    if (name && low.includes(name.toLowerCase())) return code;
  }
  return reg;
}

function unwrapList(payload) {
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload?.data)) return payload.data;
  if (Array.isArray(payload?.vehicles)) return payload.vehicles;
  if (Array.isArray(payload?.events)) return payload.events;
  if (Array.isArray(payload?.results)) return payload.results;
  return [];
}

function pick(obj, keys) {
  for (const k of keys) {
    if (obj?.[k] != null && obj[k] !== "") return obj[k];
  }
  return null;
}

function extractCartrackLocation(row) {
  const loc = row?.location && typeof row.location === "object" ? row.location : null;
  const lat = Number(
    pick(row, ["latitude", "lat", "gps_latitude"])
    ?? pick(loc, ["latitude", "lat", "gps_latitude"])
  );
  const lng = Number(
    pick(row, ["longitude", "lng", "lon", "gps_longitude"])
    ?? pick(loc, ["longitude", "lng", "lon", "gps_longitude"])
  );
  return {
    latitude: Number.isFinite(lat) ? lat : null,
    longitude: Number.isFinite(lng) ? lng : null,
    position_description: String(pick(loc, ["position_description", "description"]) || ""),
    location_updated: String(pick(loc, ["updated"]) || ""),
  };
}

function pickIgnitionOn(row) {
  const v = pick(row, ["ignition", "ignition_on", "ignition_status"]);
  if (v === true || v === 1 || v === "1") return 1;
  const s = String(v ?? "").toLowerCase();
  if (s === "on" || s === "true") return 1;
  return 0;
}

function normalizeCartrackOdometerKm(row) {
  const raw = Number(pick(row, ["odometer", "odometer_km", "mileage"]));
  if (!Number.isFinite(raw) || raw <= 0) return null;
  // Cartrack MZ returns odometer in metres — convert to km.
  return raw >= 10000 ? raw / 1000 : raw;
}

export async function cartrackApiGet(path, query = {}) {
  const creds = getCartrackCredentials();
  if (!creds) {
    const err = new Error("Cartrack is not configured. Add API credentials in User Admin → Cartrack.");
    err.code = "NOT_CONFIGURED";
    throw err;
  }
  const url = new URL(`${creds.baseUrl}/${String(path || "").replace(/^\//, "")}`);
  for (const [k, v] of Object.entries(query)) {
    if (v == null || v === "") continue;
    url.searchParams.set(k, String(v));
  }
  const auth = Buffer.from(`${creds.username}:${creds.password}`).toString("base64");
  const res = await fetch(url.toString(), {
    method: "GET",
    headers: {
      Authorization: `Basic ${auth}`,
      Accept: "application/json",
    },
    signal: AbortSignal.timeout(45000),
  });
  const text = await res.text();
  let body = null;
  try {
    body = text ? JSON.parse(text) : null;
  } catch {
    body = { raw: text };
  }
  if (!res.ok) {
    const errObj = body?.error && typeof body.error === "object" ? body.error : null;
    let msg = String(errObj?.message || body?.message || body?.raw || res.statusText || `HTTP ${res.status}`);
    if (errObj?.data && typeof errObj.data === "object") {
      const parts = Object.entries(errObj.data).flatMap(([k, v]) =>
        (Array.isArray(v) ? v : [v]).map((x) => `${k}: ${x}`)
      );
      if (parts.length) msg = `${msg} (${parts.join("; ")})`;
    }
    const err = new Error(`Cartrack API error (${res.status}): ${msg}`);
    err.status = res.status;
    throw err;
  }
  return body;
}

async function fetchAllPages(path, query = {}, maxPages = 20) {
  const all = [];
  for (let page = 1; page <= maxPages; page += 1) {
    const body = await cartrackApiGet(path, { ...query, page, per_page: 100 });
    const chunk = unwrapList(body);
    if (!chunk.length) break;
    all.push(...chunk);
    const lastPage = Number(body?.meta?.last_page || body?.pagination?.last_page || body?.last_page || 0);
    if (lastPage && page >= lastPage) break;
    if (chunk.length < 100) break;
  }
  return all;
}

function normalizeStatusRow(row) {
  const registration = normalizeRegistration(
    pick(row, ["registration", "vehicle_registration", "reg", "license_plate", "plate_number"])
  );
  const vehicleName = String(pick(row, ["vehicle_name", "name", "description"]) || "");
  const assetCode = resolveAssetCode(registration, vehicleName);
  const gps = extractCartrackLocation(row);
  const displayName = vehicleName || gps.position_description || "";
  return {
    registration,
    asset_code: assetCode,
    vehicle_id: String(pick(row, ["vehicle_id", "id"]) || ""),
    vehicle_name: displayName,
    ignition_on: pickIgnitionOn(row),
    speed_kmh: Number(pick(row, ["speed", "speed_kmh", "gps_speed", "road_speed"]) || 0),
    latitude: gps.latitude,
    longitude: gps.longitude,
    odometer_km: normalizeCartrackOdometerKm(row),
    last_event_at: String(
      pick(row, ["event_ts", "last_update", "updated_at", "event_time", "timestamp"])
      || gps.location_updated
      || ""
    ),
    status_json: JSON.stringify(row),
  };
}

function normalizeEventRow(row) {
  const registration = normalizeRegistration(
    pick(row, ["registration", "vehicle_registration", "reg", "license_plate"])
  );
  const eventType = String(pick(row, ["event_type", "type", "event_type_code", "code"]) || "").trim();
  const eventLabel = String(pick(row, ["event_type_description", "event_description", "description", "name"]) || eventType);
  const eventTime = String(pick(row, ["event_time", "timestamp", "created_at", "date_time"]) || "");
  const speed = Number(pick(row, ["speed", "speed_kmh", "actual_speed"]) || 0);
  const limit = Number(pick(row, ["speed_limit", "speed_limit_kmh", "limit"]) || 0);
  const isSpeeding = SPEED_EVENT_RE.test(`${eventType} ${eventLabel}`) || (limit > 0 && speed > limit);
  const gps = extractCartrackLocation(row);
  const eventId = String(
    pick(row, ["event_id", "id", "uuid"]) || `${registration}|${eventTime}|${eventType}|${speed}`
  );
  return {
    event_id: eventId,
    registration,
    asset_code: resolveAssetCode(registration, String(pick(row, ["vehicle_name", "name"]) || "")),
    event_type: eventType,
    event_type_label: eventLabel,
    event_time: eventTime,
    speed_kmh: Number.isFinite(speed) ? speed : null,
    speed_limit_kmh: Number.isFinite(limit) && limit > 0 ? limit : null,
    latitude: gps.latitude,
    longitude: gps.longitude,
    driver_name: String(pick(row, ["driver_name", "driver"]) || ""),
    description: eventLabel,
    is_speeding: isSpeeding ? 1 : 0,
    payload_json: JSON.stringify(row),
  };
}

export async function syncCartrackFleetStatus() {
  const rows = await fetchAllPages("vehicles/status");
  const upsert = db.prepare(`
    INSERT INTO cartrack_vehicle_snapshots (
      registration, asset_code, vehicle_id, vehicle_name, ignition_on, speed_kmh,
      latitude, longitude, odometer_km, last_event_at, status_json, synced_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
    ON CONFLICT(registration) DO UPDATE SET
      asset_code = excluded.asset_code,
      vehicle_id = excluded.vehicle_id,
      vehicle_name = excluded.vehicle_name,
      ignition_on = excluded.ignition_on,
      speed_kmh = excluded.speed_kmh,
      latitude = excluded.latitude,
      longitude = excluded.longitude,
      odometer_km = excluded.odometer_km,
      last_event_at = excluded.last_event_at,
      status_json = excluded.status_json,
      synced_at = datetime('now')
  `);
  let count = 0;
  for (const raw of rows) {
    const n = normalizeStatusRow(raw);
    if (!n.registration) continue;
    upsert.run(
      n.registration, n.asset_code, n.vehicle_id, n.vehicle_name, n.ignition_on, n.speed_kmh,
      n.latitude, n.longitude, n.odometer_km, n.last_event_at, n.status_json
    );
    count += 1;
  }
  return { synced: count };
}

export function enrichCartrackLiveRow(row, speedRegs = new Set()) {
  const lat = Number(row.latitude);
  const lng = Number(row.longitude);
  const hasGps = Number.isFinite(lat) && Number.isFinite(lng) && (lat !== 0 || lng !== 0);
  const code = row.asset_code || row.registration;
  let raw = {};
  try {
    raw = JSON.parse(row.status_json || "{}");
  } catch {
    raw = {};
  }
  const roadLimit = Number(raw.road_speed);
  const speed = Number(row.speed_kmh ?? raw.speed ?? 0);
  const overLimit = Number.isFinite(roadLimit) && roadLimit > 0 && speed > roadLimit;
  const fuelPct = raw.fuel?.precentage_left ?? raw.fuel?.percentage_left;
  return {
    ...row,
    has_gps: hasGps,
    is_speeding: speedRegs.has(row.registration) || speedRegs.has(code) || overLimit,
    is_idling: raw.idling === true,
    road_speed_limit: Number.isFinite(roadLimit) && roadLimit > 0 ? roadLimit : null,
    rpm: Number(raw.rpm) || null,
    fuel_pct: fuelPct != null && Number.isFinite(Number(fuelPct)) ? Number(fuelPct) : null,
  };
}

function cartrackTimestampRange(startDate, endDate) {
  const start = String(startDate || "").slice(0, 10);
  const end = String(endDate || start).slice(0, 10);
  return {
    start_timestamp: `${start} 00:00:00`,
    end_timestamp: `${end} 23:59:59`,
  };
}

function parseCartrackYmd(dateStr) {
  const d = String(dateStr || "").slice(0, 10);
  const [y, m, day] = d.split("-").map(Number);
  return new Date(y, m - 1, day);
}

function formatCartrackYmd(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function addCartrackDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

async function fetchCartrackEventsForDay(dayYmd) {
  const nextYmd = formatCartrackYmd(addCartrackDays(parseCartrackYmd(dayYmd), 1));
  const params = {
    start_timestamp: `${dayYmd} 00:00:00`,
    end_timestamp: `${nextYmd} 00:00:00`,
  };
  try {
    return await fetchAllPages("vehicles/events", params);
  } catch (err) {
    const match = String(err?.message || "").match(/before or equal to (\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})/i);
    if (!match) throw err;
    return fetchAllPages("vehicles/events", { ...params, end_timestamp: match[1] });
  }
}

async function fetchCartrackEvents(startDate, endDate) {
  const start = parseCartrackYmd(startDate);
  const end = parseCartrackYmd(endDate);
  const all = [];
  for (let d = start; d <= end; d = addCartrackDays(d, 1)) {
    const rows = await fetchCartrackEventsForDay(formatCartrackYmd(d));
    all.push(...rows);
  }
  return all;
}

export async function syncCartrackEvents({ startDate, endDate }) {
  const start = String(startDate || "").slice(0, 10);
  const end = String(endDate || start).slice(0, 10);
  const rows = await fetchCartrackEvents(start, end);
  const ins = db.prepare(`
    INSERT INTO cartrack_events (
      event_id, registration, asset_code, event_type, event_type_label, event_time,
      speed_kmh, speed_limit_kmh, latitude, longitude, driver_name, description,
      is_speeding, payload_json
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(event_id) DO UPDATE SET
      asset_code = excluded.asset_code,
      event_type_label = excluded.event_type_label,
      speed_kmh = excluded.speed_kmh,
      speed_limit_kmh = excluded.speed_limit_kmh,
      is_speeding = excluded.is_speeding,
      payload_json = excluded.payload_json
  `);
  let count = 0;
  let speeding = 0;
  for (const raw of rows) {
    const n = normalizeEventRow(raw);
    if (!n.registration || !n.event_time) continue;
    ins.run(
      n.event_id, n.registration, n.asset_code, n.event_type, n.event_type_label, n.event_time,
      n.speed_kmh, n.speed_limit_kmh, n.latitude, n.longitude, n.driver_name, n.description,
      n.is_speeding, n.payload_json
    );
    count += 1;
    if (n.is_speeding) speeding += 1;
  }
  return { synced: count, speeding };
}

export function listCartrackFleetFromDb() {
  return db.prepare(`
    SELECT registration, asset_code, vehicle_id, vehicle_name, ignition_on, speed_kmh,
           latitude, longitude, odometer_km, last_event_at, synced_at
    FROM cartrack_vehicle_snapshots
    ORDER BY asset_code ASC, registration ASC
  `).all();
}

export function listCartrackEventsFromDb({ startDate, endDate, speedingOnly = false, limit = 200 } = {}) {
  const start = String(startDate || "").slice(0, 10);
  const end = String(endDate || start).slice(0, 10);
  const lim = Math.max(1, Math.min(500, Number(limit) || 200));
  const where = ["date(event_time) >= date(?)", "date(event_time) <= date(?)"];
  const params = [start, end];
  if (speedingOnly) where.push("is_speeding = 1");
  return db.prepare(`
    SELECT id, event_id, registration, asset_code, event_type, event_type_label, event_time,
           speed_kmh, speed_limit_kmh, driver_name, description, is_speeding
    FROM cartrack_events
    WHERE ${where.join(" AND ")}
    ORDER BY event_time DESC
    LIMIT ?
  `).all(...params, lim);
}

export function buildMorningSpeedingReport(reportDate) {
  const date = String(reportDate || "").slice(0, 10);
  const events = listCartrackEventsFromDb({ startDate: date, endDate: date, speedingOnly: true, limit: 500 });
  const byVehicle = new Map();
  for (const e of events) {
    const key = e.asset_code || e.registration || "Unknown";
    if (!byVehicle.has(key)) {
      byVehicle.set(key, { asset_code: key, registration: e.registration, count: 0, max_speed: 0, events: [] });
    }
    const row = byVehicle.get(key);
    row.count += 1;
    row.max_speed = Math.max(row.max_speed, Number(e.speed_kmh || 0));
    if (row.events.length < 5) row.events.push(e);
  }
  const vehicles = Array.from(byVehicle.values()).sort((a, b) => b.count - a.count);
  const summary = {
    report_date: date,
    total_speeding_events: events.length,
    vehicles_with_speeding: vehicles.length,
    vehicles,
    events,
    generated_at: new Date().toISOString(),
  };
  db.prepare(`
    INSERT INTO cartrack_morning_reports (report_date, summary_json, created_at)
    VALUES (?, ?, datetime('now'))
    ON CONFLICT(report_date) DO UPDATE SET summary_json = excluded.summary_json
  `).run(date, JSON.stringify(summary));
  return summary;
}

export function formatMorningReportText(summary) {
  const lines = [];
  lines.push(`IRONLOG — Cartrack speeding report`);
  lines.push(`Date: ${summary.report_date}`);
  lines.push(`Total speeding events: ${summary.total_speeding_events}`);
  lines.push(`Vehicles involved: ${summary.vehicles_with_speeding}`);
  lines.push("");
  if (!summary.vehicles?.length) {
    lines.push("No speeding events recorded for this date.");
    return lines.join("\n");
  }
  lines.push("By vehicle:");
  for (const v of summary.vehicles) {
    lines.push(`  • ${v.asset_code} (${v.registration || "—"}) — ${v.count} event(s), max ${Number(v.max_speed || 0).toFixed(0)} km/h`);
  }
  lines.push("");
  lines.push("Recent events:");
  for (const e of (summary.events || []).slice(0, 25)) {
    const spd = e.speed_kmh != null ? `${Number(e.speed_kmh).toFixed(0)} km/h` : "—";
    const lim = e.speed_limit_kmh != null ? ` (limit ${Number(e.speed_limit_kmh).toFixed(0)})` : "";
    lines.push(`  ${String(e.event_time || "").slice(0, 16)}  ${e.asset_code || e.registration}  ${spd}${lim}  ${e.event_type_label || e.event_type || ""}`);
  }
  return lines.join("\n");
}

function buildSmtpTransport() {
  const row = db.prepare(`SELECT * FROM smtp_settings WHERE id = 1`).get();
  if (!row?.host || !row?.username || !row?.password_enc || !row?.from_email) {
    return { error: "SMTP is not configured" };
  }
  const password = decryptSecret(row.password_enc);
  if (!password) return { error: "SMTP password could not be decrypted" };
  const fromEmail = String(row.from_email).trim();
  const transporter = nodemailer.createTransport({
    host: String(row.host).trim(),
    port: Math.max(1, Number(row.port || 587)),
    secure: Number(row.secure || 0) === 1,
    requireTLS: Number(row.secure || 0) !== 1 && Number(row.port || 587) === 587,
    auth: { user: String(row.username).trim(), pass: password },
    tls: { minVersion: "TLSv1.2" },
  });
  const from = row.from_name
    ? `"${String(row.from_name).replace(/"/g, "")}" <${fromEmail}>`
    : fromEmail;
  return { transporter, from };
}

export async function sendCartrackMorningEmail(summary, recipients) {
  const to = (Array.isArray(recipients) ? recipients : String(recipients || "").split(","))
    .map((x) => String(x).trim())
    .filter(Boolean);
  if (!to.length) return { ok: false, error: "no_recipients" };
  const smtp = buildSmtpTransport();
  if (smtp.error) return { ok: false, error: smtp.error };
  const text = formatMorningReportText(summary);
  await smtp.transporter.sendMail({
    from: smtp.from,
    to: to.join(", "),
    subject: `IRONLOG Cartrack speeding report — ${summary.report_date} (${summary.total_speeding_events} events)`,
    text,
  });
  db.prepare(`UPDATE cartrack_morning_reports SET emailed_at = datetime('now') WHERE report_date = ?`).run(summary.report_date);
  return { ok: true, recipients: to };
}

export async function runCartrackMorningJob({ reportDate, sendEmail = true, log = console } = {}) {
  const date = String(reportDate || "").slice(0, 10);
  const creds = getCartrackCredentials();
  if (!creds) {
    log.warn?.("[cartrack] morning job skipped — not configured");
    return { ok: false, error: "not_configured" };
  }
  await syncCartrackFleetStatus();
  await syncCartrackEvents({ startDate: date, endDate: date });
  const summary = buildMorningSpeedingReport(date);
  let emailed = false;
  if (sendEmail) {
    const row = settingsRow() || {};
    const recipients = String(row.morning_recipients || process.env.CARTRACK_MORNING_RECIPIENTS || "")
      .split(",")
      .map((x) => x.trim())
      .filter(Boolean);
    if (recipients.length) {
      const mail = await sendCartrackMorningEmail(summary, recipients);
      emailed = Boolean(mail.ok);
      if (!mail.ok) log.warn?.(`[cartrack] email not sent: ${mail.error}`);
    }
  }
  return { ok: true, report_date: date, summary, emailed };
}

export async function runCartrackAutoScheduler(log = console) {
  ensureCartrackTables();
  const enabled = String(process.env.CARTRACK_MORNING_ENABLED || "1").trim().toLowerCase();
  if (["0", "false", "no", "off"].includes(enabled)) {
    log.info?.("[cartrack] morning scheduler disabled");
    return null;
  }
  const row = settingsRow() || {};
  if (Number(row.morning_enabled ?? 1) !== 1) {
    log.info?.("[cartrack] morning scheduler disabled in settings");
    return null;
  }
  const runHour = Math.max(0, Math.min(23, Number(process.env.CARTRACK_MORNING_HOUR ?? row.morning_hour ?? 6)));
  const runMinute = Math.max(0, Math.min(59, Number(process.env.CARTRACK_MORNING_MINUTE ?? row.morning_minute ?? 0)));
  let lastMinuteKey = "";

  const maybeRun = async () => {
    const now = new Date();
    const minuteKey = `${now.getFullYear()}-${now.getMonth() + 1}-${now.getDate()} ${now.getHours()}:${now.getMinutes()}`;
    if (now.getHours() !== runHour || now.getMinutes() !== runMinute) return;
    if (lastMinuteKey === minuteKey) return;
    lastMinuteKey = minuteKey;
    const yesterday = new Date(now);
    yesterday.setDate(yesterday.getDate() - 1);
    const reportDate = yesterday.toISOString().slice(0, 10);
    try {
      const result = await runCartrackMorningJob({ reportDate, sendEmail: true, log });
      log.info?.(`[cartrack] morning report for ${reportDate}: ${result.summary?.total_speeding_events ?? 0} speeding event(s)`);
    } catch (err) {
      log.error?.(`[cartrack] morning job failed: ${err?.message || err}`);
    }
  };

  await maybeRun();
  const timer = setInterval(maybeRun, 30 * 1000);
  log.info?.(`[cartrack] morning scheduler at ${String(runHour).padStart(2, "0")}:${String(runMinute).padStart(2, "0")} (yesterday's speeding)`);
  return timer;
}
