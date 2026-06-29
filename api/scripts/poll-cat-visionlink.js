#!/usr/bin/env node
/**
 * CAT VisionLink -> IRONLOG telematics bridge (poll once).
 *
 * This script is intentionally endpoint-agnostic:
 * - You provide a full CAT data URL in VISIONLINK_LASTREPORTED_URL
 * - We normalize common payload shapes and post to /api/telematics/ingest
 *
 * Run:
 *   node scripts/poll-cat-visionlink.js
 */

import path from "path";
import { fileURLToPath } from "url";
import dotenv from "dotenv";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
dotenv.config({ path: path.join(__dirname, "../.env") });
dotenv.config({ path: path.join(process.cwd(), ".env") });

function mustEnv(name) {
  const v = String(process.env[name] || "").trim();
  if (!v) throw new Error(`Missing required env var: ${name}`);
  return v;
}

function numOrNull(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function pick(obj, keys) {
  if (!obj || typeof obj !== "object") return undefined;
  for (const key of keys) {
    if (obj[key] !== undefined && obj[key] !== null) return obj[key];
  }
  return undefined;
}

function toArrayPayload(payload) {
  if (Array.isArray(payload)) return payload;
  if (!payload || typeof payload !== "object") return [];
  const candidates = [
    payload.data,
    payload.items,
    payload.results,
    payload.assets,
    payload.equipment,
    payload.lastReported,
  ];
  for (const c of candidates) {
    if (Array.isArray(c)) return c;
  }
  return [];
}

function parseAssetMapFromEnv() {
  const raw = String(process.env.VISIONLINK_ASSET_MAP_JSON || "").trim();
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (err) {
    throw new Error(`VISIONLINK_ASSET_MAP_JSON is not valid JSON: ${err.message}`);
  }
}

function normalizeVisionLinkRecord(rec, mapBySerial, mapByAssetCode) {
  const deviceSerialRaw = pick(rec, [
    "deviceSerialNumber",
    "device_serial",
    "deviceSerial",
    "deviceId",
    "device_id",
    "serialNumber",
    "serial",
    "mainDeviceSerial",
    "main_device_serial",
  ]);
  const equipmentCodeRaw = pick(rec, [
    "asset_code",
    "assetCode",
    "equipmentId",
    "equipment_id",
    "assetId",
    "asset_id",
    "machineId",
    "machine_id",
  ]);

  const deviceSerial = String(deviceSerialRaw || "").trim();
  const sourceAssetCode = String(equipmentCodeRaw || "").trim().toUpperCase();
  const mappedAssetCode = (
    mapBySerial[deviceSerial]
    || mapBySerial[sourceAssetCode]
    || mapByAssetCode[sourceAssetCode]
    || sourceAssetCode
  );
  const assetCode = String(mappedAssetCode || "").trim().toUpperCase();

  const engineHours = numOrNull(pick(rec, [
    "engineHours",
    "engine_hours",
    "serviceHours",
    "service_hours",
    "meterHours",
    "meter_hours",
    "hours",
  ]));

  const runSecondsToday = numOrNull(pick(rec, [
    "runSecondsToday",
    "run_seconds_today",
    "runTimeTodaySeconds",
    "run_time_today_sec",
  ]));

  const idleSecondsToday = numOrNull(pick(rec, [
    "idleSecondsToday",
    "idle_seconds_today",
    "idleTimeTodaySeconds",
    "idle_time_today_sec",
  ]));

  const timestamp = String(pick(rec, [
    "timestamp",
    "recordedAt",
    "recorded_at",
    "lastReported",
    "last_reported",
    "eventTime",
    "event_time",
  ]) || new Date().toISOString()).trim();

  const lat = numOrNull(pick(rec, ["latitude", "lat"]));
  const lon = numOrNull(pick(rec, ["longitude", "lon", "lng"]));
  const speedKmh = numOrNull(pick(rec, ["speedKmh", "speed_kmh", "speed"]));
  const ignition = pick(rec, ["ignition", "ignitionOn", "ignition_on"]);

  if (!assetCode && !deviceSerial) return null;
  if (engineHours == null) return null;

  return {
    asset_code: assetCode || undefined,
    device_serial: deviceSerial || undefined,
    recorded_at: timestamp,
    engine_hours: engineHours,
    run_seconds_today: runSecondsToday,
    idle_seconds_today: idleSecondsToday,
    latitude: lat,
    longitude: lon,
    speed_kmh: speedKmh,
    ignition_on: ignition === true || ignition === 1 || String(ignition).toLowerCase() === "on",
    source: "cat_visionlink",
    payload: rec,
  };
}

async function fetchAccessToken() {
  const tokenUrl = mustEnv("VISIONLINK_TOKEN_URL");
  const clientId = mustEnv("VISIONLINK_CLIENT_ID");
  const clientSecret = mustEnv("VISIONLINK_CLIENT_SECRET");
  const scope = String(process.env.VISIONLINK_SCOPE || "").trim();

  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
  });
  if (scope) body.set("scope", scope);

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Token request failed (${res.status}): ${txt.slice(0, 500)}`);
  }
  const json = await res.json();
  const token = String(json?.access_token || "").trim();
  if (!token) throw new Error("Token response missing access_token");
  return token;
}

async function fetchVisionLinkData(accessToken) {
  const url = mustEnv("VISIONLINK_LASTREPORTED_URL");
  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
    },
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`VisionLink data request failed (${res.status}): ${txt.slice(0, 500)}`);
  }
  return res.json();
}

async function postToIronlog(records) {
  const apiBase = String(process.env.IRONLOG_API_BASE || "http://127.0.0.1:3001").trim().replace(/\/+$/, "");
  const ingestUrl = `${apiBase}/api/telematics/ingest`;
  const apiKey = String(process.env.TELEMATICS_API_KEY || process.env.FSC_TELEMATICS_API_KEY || "").trim();

  let ok = 0;
  let failed = 0;
  for (const row of records) {
    const payload = {
      asset_code: row.asset_code,
      device_serial: row.device_serial,
      recorded_at: row.recorded_at,
      engine_hours: row.engine_hours,
      run_seconds_today: row.run_seconds_today,
      idle_seconds_today: row.idle_seconds_today,
      latitude: row.latitude,
      longitude: row.longitude,
      speed_kmh: row.speed_kmh,
      ignition_on: row.ignition_on,
      source: row.source,
      upstream_payload: row.payload,
    };

    const res = await fetch(ingestUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...(apiKey ? { "x-api-key": apiKey } : {}),
      },
      body: JSON.stringify(payload),
    });
    if (res.ok) ok += 1;
    else {
      failed += 1;
      const txt = await res.text();
      console.warn(`[visionlink] ingest failed for ${row.asset_code || row.device_serial}: ${res.status} ${txt}`);
    }
  }
  return { ok, failed };
}

async function main() {
  const map = parseAssetMapFromEnv();
  const mapBySerial = {};
  const mapByAssetCode = {};
  for (const [k, v] of Object.entries(map)) {
    const key = String(k || "").trim().toUpperCase();
    const value = String(v || "").trim().toUpperCase();
    if (!key || !value) continue;
    mapBySerial[key] = value;
    mapByAssetCode[key] = value;
  }

  const token = await fetchAccessToken();
  const raw = await fetchVisionLinkData(token);
  const rows = toArrayPayload(raw)
    .map((r) => normalizeVisionLinkRecord(r, mapBySerial, mapByAssetCode))
    .filter(Boolean);

  if (!rows.length) {
    console.log("[visionlink] no usable records found in payload");
    return;
  }

  const result = await postToIronlog(rows);
  console.log(`[visionlink] processed=${rows.length} ingested=${result.ok} failed=${result.failed}`);
}

main().catch((err) => {
  console.error("[visionlink] bridge failed:", err?.message || err);
  process.exitCode = 1;
});
