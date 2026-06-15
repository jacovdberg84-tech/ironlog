// Unitech Moz (GpsGate) KML fleet feed — Afungi site vehicles.

import crypto from "node:crypto";
import { db } from "../db/client.js";
import { getCartrackSpeedAlertKmh } from "./cartrack.js";
import {
  normalizeGpsRegistration,
  resolveGpsAssetCode,
} from "./gpsVehicleLinks.js";

const GPSGATE_USERS_PATH = "https://unitechmoz.gpsgate.com/comGpsGate/rpc/KmlFeed/v.1/UsersInGroups";

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

function normalizeRegistration(reg) {
  return normalizeGpsRegistration(reg);
}

function registrationFromName(name) {
  const raw = String(name || "").trim();
  if (!raw) return "";
  const beforeDash = raw.split(" - ")[0].trim();
  const compact = normalizeRegistration(beforeDash);
  if (compact.length >= 5) return compact;
  return normalizeRegistration(raw.replace(/[^A-Za-z0-9]/g, ""));
}

function resolveAssetCode(registration, vehicleName = "") {
  return resolveGpsAssetCode(registration, vehicleName);
}

function decodeXmlEntities(s) {
  return String(s || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function pickDescField(html, label) {
  const re = new RegExp(`<b>${label}:<\\/b>\\s*([^<]+)`, "i");
  const m = String(html || "").match(re);
  return m ? m[1].trim() : null;
}

function parsePositionAgeDays(raw) {
  const text = String(raw || "").trim();
  if (!text) return null;
  const days = Number(text.split(/\s+/)[0]);
  return Number.isFinite(days) ? days : null;
}

function parseGpsGateTimestamp(raw) {
  const text = String(raw || "").trim();
  if (!text) return null;
  const ms = Date.parse(text.replace(/\s+/g, " "));
  if (!Number.isFinite(ms)) return null;
  return new Date(ms).toISOString().replace("T", " ").slice(0, 19);
}

function parseSpeedKmh(raw) {
  const m = String(raw || "").match(/([\d.]+)/);
  if (!m) return 0;
  const n = Number(m[1]);
  return Number.isFinite(n) ? n : 0;
}

export function ensureUnitechTables() {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS unitech_settings (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      kml_feed_url_enc TEXT,
      feed_label TEXT,
      enabled INTEGER NOT NULL DEFAULT 1,
      updated_by TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`INSERT INTO unitech_settings (id) SELECT 1 WHERE NOT EXISTS (SELECT 1 FROM unitech_settings WHERE id = 1)`).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS unitech_vehicle_snapshots (
      registration TEXT PRIMARY KEY,
      asset_code TEXT,
      vehicle_id TEXT,
      vehicle_name TEXT,
      speed_kmh REAL,
      bearing REAL,
      latitude REAL,
      longitude REAL,
      position_age_days REAL,
      last_event_at TEXT,
      status_json TEXT,
      synced_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
}

function settingsRow() {
  ensureUnitechTables();
  return db.prepare(`SELECT * FROM unitech_settings WHERE id = 1`).get();
}

export function getUnitechPublicSettings() {
  const row = settingsRow() || {};
  const envUrl = String(process.env.UNITECH_KML_FEED_URL || "").trim();
  const hasEnv = Boolean(envUrl);
  const hasDb = Boolean(String(row.kml_feed_url_enc || "").trim());
  return {
    configured: hasEnv || hasDb,
    source: hasEnv ? "env" : hasDb ? "database" : "none",
    feed_label: String(row.feed_label || process.env.UNITECH_FEED_LABEL || "Afungi (Unitech)").trim(),
    enabled: Number(row.enabled ?? 1) === 1,
    has_feed_url: hasEnv || hasDb,
    updated_at: row.updated_at || null,
    updated_by: row.updated_by || null,
  };
}

export function saveUnitechSettings({ kml_feed_url, feed_label, enabled, updated_by }) {
  ensureUnitechTables();
  const existing = settingsRow() || {};
  const nextUrlEnc = (() => {
    if (kml_feed_url != null && String(kml_feed_url).trim() !== "") {
      return encryptSecret(String(kml_feed_url).trim());
    }
    return String(existing.kml_feed_url_enc || "");
  })();
  db.prepare(`
    UPDATE unitech_settings
    SET kml_feed_url_enc = ?, feed_label = ?, enabled = ?, updated_by = ?, updated_at = datetime('now')
    WHERE id = 1
  `).run(
    nextUrlEnc,
    String(feed_label || existing.feed_label || "Afungi (Unitech)").trim(),
    enabled === false || String(enabled) === "0" ? 0 : 1,
    String(updated_by || "admin").trim()
  );
  return getUnitechPublicSettings();
}

function getUnitechFeedUrl() {
  const envUrl = String(process.env.UNITECH_KML_FEED_URL || "").trim();
  if (envUrl) return envUrl;
  const row = settingsRow();
  if (!row?.kml_feed_url_enc) return null;
  const url = decryptSecret(row.kml_feed_url_enc);
  return url ? String(url).trim() : null;
}

function extractSessionQuery(feedXml) {
  const decoded = decodeXmlEntities(feedXml);
  const httpQuery = (decoded.match(/<httpQuery>([\s\S]*?)<\/httpQuery>/i) || [])[1] || "";
  const sessionMatch = httpQuery.match(/sessionId=([^&]+)/i);
  if (!sessionMatch) return null;
  const groupMatch = httpQuery.match(/groupIds=([^&]+)/i);
  const appMatch = httpQuery.match(/appId=([^&]+)/i);
  return {
    sessionId: decodeXmlEntities(sessionMatch[1]),
    groupIds: groupMatch ? decodeXmlEntities(groupMatch[1]) : "378",
    appId: appMatch ? decodeXmlEntities(appMatch[1]) : "35",
  };
}

async function fetchKmlText(url, timeoutMs = 30000) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      signal: ctrl.signal,
      headers: { Accept: "application/vnd.google-earth.kml+xml, application/xml, text/xml, */*" },
    });
    const text = await res.text();
    if (!res.ok) {
      const err = new Error(`Unitech KML HTTP ${res.status}: ${text.slice(0, 200)}`);
      err.code = "HTTP_ERROR";
      throw err;
    }
    return text;
  } finally {
    clearTimeout(timer);
  }
}

export function parseUnitechKmlPlacemarks(kmlXml) {
  const placemarks = [];
  const re = /<Placemark>([\s\S]*?)<\/Placemark>/gi;
  let match;
  while ((match = re.exec(String(kmlXml || "")))) {
    const block = match[1];
    const name = decodeXmlEntities((block.match(/<name>([\s\S]*?)<\/name>/i) || [])[1] || "").trim();
    const descHtml = decodeXmlEntities((block.match(/<description>([\s\S]*?)<\/description>/i) || [])[1] || "");
    const coordsRaw = String((block.match(/<coordinates>([\s\S]*?)<\/coordinates>/i) || [])[1] || "").trim();
    const parts = coordsRaw.split(",").map((x) => Number(x.trim()));
    const lng = parts[0];
    const lat = parts[1];
    if (!name || !Number.isFinite(lat) || !Number.isFinite(lng)) continue;
    const username = pickDescField(descHtml, "Username") || registrationFromName(name);
    const registration = normalizeRegistration(username || registrationFromName(name));
    const positionAgeRaw = pickDescField(descHtml, "Position age");
    const timestampRaw = pickDescField(descHtml, "TimeStamp");
    const speedRaw = pickDescField(descHtml, "Speed");
    const bearingRaw = pickDescField(descHtml, "Bearing");
    placemarks.push({
      name,
      registration,
      vehicle_id: String(username || registration),
      vehicle_name: name,
      speed_kmh: parseSpeedKmh(speedRaw),
      bearing: Number(bearingRaw) || null,
      latitude: lat,
      longitude: lng,
      position_age_days: parsePositionAgeDays(positionAgeRaw),
      last_event_at: parseGpsGateTimestamp(timestampRaw),
      status_json: JSON.stringify({
        name,
        username,
        position_age: positionAgeRaw,
        timestamp: timestampRaw,
        speed: speedRaw,
        bearing: bearingRaw,
      }),
    });
  }
  return placemarks;
}

export async function fetchUnitechKmlDocument() {
  const feedUrl = getUnitechFeedUrl();
  if (!feedUrl) {
    const err = new Error("Unitech KML feed URL not configured");
    err.code = "NOT_CONFIGURED";
    throw err;
  }
  const settings = getUnitechPublicSettings();
  if (!settings.enabled) {
    const err = new Error("Unitech feed is disabled");
    err.code = "DISABLED";
    throw err;
  }
  const feedXml = await fetchKmlText(feedUrl);
  const session = extractSessionQuery(feedXml);
  if (!session?.sessionId) {
    const err = new Error("Unitech KML feed did not return a session");
    err.code = "INVALID_FEED";
    throw err;
  }
  const q = new URLSearchParams({
    groupIds: session.groupIds,
    sessionId: session.sessionId,
    appId: session.appId,
  });
  return fetchKmlText(`${GPSGATE_USERS_PATH}?${q.toString()}`);
}

export async function syncUnitechFleetStatus() {
  const kmlXml = await fetchUnitechKmlDocument();
  const rows = parseUnitechKmlPlacemarks(kmlXml);
  const upsert = db.prepare(`
    INSERT INTO unitech_vehicle_snapshots (
      registration, asset_code, vehicle_id, vehicle_name, speed_kmh, bearing,
      latitude, longitude, position_age_days, last_event_at, status_json, synced_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
    ON CONFLICT(registration) DO UPDATE SET
      asset_code = excluded.asset_code,
      vehicle_id = excluded.vehicle_id,
      vehicle_name = excluded.vehicle_name,
      speed_kmh = excluded.speed_kmh,
      bearing = excluded.bearing,
      latitude = excluded.latitude,
      longitude = excluded.longitude,
      position_age_days = excluded.position_age_days,
      last_event_at = excluded.last_event_at,
      status_json = excluded.status_json,
      synced_at = datetime('now')
  `);
  let count = 0;
  for (const raw of rows) {
    if (!raw.registration) continue;
    const assetCode = resolveAssetCode(raw.registration, raw.vehicle_name);
    upsert.run(
      raw.registration,
      assetCode,
      raw.vehicle_id,
      raw.vehicle_name,
      raw.speed_kmh,
      raw.bearing,
      raw.latitude,
      raw.longitude,
      raw.position_age_days,
      raw.last_event_at,
      raw.status_json
    );
    count += 1;
  }
  return { synced: count, feed_label: getUnitechPublicSettings().feed_label };
}

export function listUnitechFleetFromDb() {
  ensureUnitechTables();
  return db.prepare(`
    SELECT registration, asset_code, vehicle_id, vehicle_name, speed_kmh, bearing,
           latitude, longitude, position_age_days, last_event_at, synced_at, status_json
    FROM unitech_vehicle_snapshots
    ORDER BY asset_code ASC, registration ASC
  `).all();
}

export function enrichUnitechLiveRow(row, speedRegs = new Set()) {
  const lat = Number(row.latitude);
  const lng = Number(row.longitude);
  const hasGps = Number.isFinite(lat) && Number.isFinite(lng) && (lat !== 0 || lng !== 0);
  const code = row.asset_code || row.registration;
  const alertKmh = getCartrackSpeedAlertKmh();
  const speed = Number(row.speed_kmh || 0);
  const positionAgeDays = Number(row.position_age_days);
  const positionStale = Number.isFinite(positionAgeDays) && positionAgeDays >= 1;
  const overThreshold = Number.isFinite(speed) && speed >= alertKmh;
  return {
    ...row,
    gps_source: "unitech",
    gps_provider: getUnitechPublicSettings().feed_label || "Unitech",
    ignition_on: null,
    odometer_km: null,
    has_gps: hasGps,
    is_speeding: speedRegs.has(row.registration) || speedRegs.has(code) || overThreshold,
    speed_alert_kmh: alertKmh,
    is_idling: false,
    road_speed_limit: null,
    position_stale: positionStale,
    position_age_days: Number.isFinite(positionAgeDays) ? positionAgeDays : null,
    bearing: row.bearing != null ? Number(row.bearing) : null,
    fuel_pct: null,
    tracker_battery_pct: null,
    supply_voltage_v: null,
    ev_battery_pct: null,
    charging_status: null,
  };
}

export async function testUnitechConnection() {
  const kmlXml = await fetchUnitechKmlDocument();
  const rows = parseUnitechKmlPlacemarks(kmlXml);
  const fresh = rows.filter((r) => !Number.isFinite(r.position_age_days) || r.position_age_days < 1).length;
  return {
    vehicle_count: rows.length,
    fresh_positions: fresh,
    message: rows.length
      ? `Unitech connected — ${rows.length} vehicle(s) in KML feed (${fresh} with position under 1 day old).`
      : "Unitech connected but no vehicles found in KML feed.",
  };
}

export function runUnitechAutoScheduler(log = console) {
  ensureUnitechTables();
  const settings = getUnitechPublicSettings();
  if (!settings.configured || !settings.enabled) {
    log.info?.("[unitech] scheduler skipped — not configured or disabled");
    return null;
  }
  const pollMs = Math.max(60_000, Number(process.env.UNITECH_POLL_MS || 3 * 60 * 1000));
  setInterval(() => {
    syncUnitechFleetStatus().catch((err) => {
      log.warn?.(`[unitech] KML poll failed: ${err?.message || err}`);
    });
  }, pollMs);
  log.info?.(`[unitech] KML poll every ${Math.round(pollMs / 1000)}s (${settings.feed_label})`);
  return pollMs;
}
