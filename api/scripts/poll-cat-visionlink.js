#!/usr/bin/env node
/**
 * CAT ISO 15143 (AEMP 2.0) -> IRONLOG telematics bridge.
 *
 * Spec: api/scripts/cat-visionlink-openapi.yaml
 * Base: https://api.cat.com/telematics/iso15143
 *
 * Run:
 *   node scripts/poll-cat-visionlink.js
 *   node scripts/poll-cat-visionlink.js --dry-run
 */

import path from "path";
import { randomUUID } from "crypto";
import { fileURLToPath } from "url";
import dotenv from "dotenv";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
dotenv.config({ path: path.join(__dirname, "../.env") });
dotenv.config({ path: path.join(process.cwd(), ".env") });

const ISO_ACCEPT = "application/iso15143-snapshot+json";

function env(name, fallback = "") {
  const v = process.env[name];
  if (v === undefined || v === null || String(v).trim() === "") return fallback;
  return String(v).trim();
}

function mustEnv(name) {
  const v = env(name);
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

function parseJsonMap(raw, label) {
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
      throw new Error("expected JSON object");
    }
    return parsed;
  } catch (err) {
    throw new Error(`${label} is not valid JSON: ${err.message}`);
  }
}

function buildLookup(map) {
  const out = {};
  for (const [k, v] of Object.entries(map)) {
    const key = String(k || "").trim().toUpperCase();
    const value = String(v || "").trim().toUpperCase();
    if (!key || !value) continue;
    out[key] = value;
  }
  return out;
}

function getConfig() {
  const mode = env("VISIONLINK_API_MODE", env("VISIONLINK_EQUIPMENT_SERIAL") ? "equipment" : "fleet").toLowerCase();
  const authMethod = env("VISIONLINK_AUTH_METHOD", "catFedLogin").toLowerCase();
  const baseUrl = env("VISIONLINK_BASE_URL", "https://api.cat.com/telematics/iso15143").replace(/\/+$/, "");

  let tokenUrl = env("VISIONLINK_TOKEN_URL");
  if (!tokenUrl) {
    if (authMethod === "entraid" || authMethod === "entra") {
      tokenUrl = "https://login.microsoftonline.com/ceb177bf-013b-49ab-8a9c-4abce32afc1e/oauth2/v2.0/token";
    } else {
      tokenUrl = "https://fedlogin.cat.com/as/token.oauth2?pfidpadapterid=OAuthAdapterCCDS";
    }
  }

  return {
    mode,
    authMethod,
    baseUrl,
    tokenUrl,
    clientId: env("VISIONLINK_CLIENT_ID"),
    clientSecret: env("VISIONLINK_CLIENT_SECRET"),
    scope: env("VISIONLINK_SCOPE"),
    customDataUrl: env("VISIONLINK_LASTREPORTED_URL"),
    fleetStartPage: Math.max(1, Number(env("VISIONLINK_FLEET_PAGE", "1")) || 1),
    make: env("VISIONLINK_EQUIPMENT_MAKE", "CAT"),
    model: env("VISIONLINK_EQUIPMENT_MODEL", "null"),
    serial: env("VISIONLINK_EQUIPMENT_SERIAL", ""),
    assetMap: buildLookup(parseJsonMap(env("VISIONLINK_ASSET_MAP_JSON"), "VISIONLINK_ASSET_MAP_JSON")),
    deviceMap: buildLookup(parseJsonMap(env("VISIONLINK_DEVICE_MAP_JSON"), "VISIONLINK_DEVICE_MAP_JSON")),
    apiBase: env("IRONLOG_API_BASE", "http://127.0.0.1:3001").replace(/\/+$/, ""),
    apiKey: env("TELEMATICS_API_KEY", env("FSC_TELEMATICS_API_KEY")),
    dryRun: process.argv.includes("--dry-run"),
  };
}

function plannedUrls(config) {
  if (config.customDataUrl) return [config.customDataUrl];
  if (config.mode === "equipment") {
    if (!config.serial) throw new Error("VISIONLINK_EQUIPMENT_SERIAL is required for equipment mode");
    const model = encodeURIComponent(config.model || "null");
    const make = encodeURIComponent(config.make || "CAT");
    const serial = encodeURIComponent(config.serial);
    return [`${config.baseUrl}/fleet/equipment/makeModelSerial/${make}/${model}/${serial}`];
  }
  return [`${config.baseUrl}/fleet/${config.fleetStartPage}`];
}

function extractEquipmentList(payload) {
  if (!payload || typeof payload !== "object") return { equipment: [], snapshotTime: null, links: [] };
  const equipment = Array.isArray(payload.Equipment) ? payload.Equipment : [];
  const snapshotTime = payload.SnapshotTime || payload.snapshotTime || null;
  const links = Array.isArray(payload.Links) ? payload.Links : [];
  return { equipment, snapshotTime, links };
}

function mapSeverity(sev) {
  const s = String(sev || "").trim().toLowerCase();
  if (s.includes("high") || s.includes("severe") || s.includes("critical")) return "critical";
  if (s.includes("medium") || s.includes("moderate")) return "warning";
  return "warning";
}

function normalizeIsoEquipment(eq, config, snapshotTime) {
  const header = eq?.EquipmentHeader || {};
  const catSerial = String(header.SerialNumber || "").trim().toUpperCase();
  const equipmentId = String(header.EquipmentID || "").trim().toUpperCase();
  const make = String(header.OEMName || config.make || "CAT").trim();
  const model = String(header.Model || config.model || "").trim();

  const assetCode = (
    config.assetMap[catSerial]
    || config.assetMap[equipmentId]
    || config.assetMap[config.serial.toUpperCase()]
    || catSerial
    || equipmentId
  );

  const deviceSerial = (
    config.deviceMap[catSerial]
    || config.deviceMap[equipmentId]
    || config.deviceMap[config.serial.toUpperCase()]
    || catSerial
  );

  const op = eq?.CumulativeOperatingHours || {};
  const idle = eq?.CumulativeIdleHours || {};
  const loc = eq?.Location || {};
  const engine = eq?.EngineStatus || {};

  const engineHours = numOrNull(op.Hour);
  const recordedAt = String(
    op.datetime
    || loc.datetime
    || engine.datetime
    || snapshotTime
    || new Date().toISOString()
  ).trim();

  if (!assetCode && !deviceSerial) return null;
  if (engineHours == null) return null;

  const faults = Array.isArray(eq?.FaultCode)
    ? eq.FaultCode.map((f) => ({
        code: String(f?.CodeIdentifier || "").trim(),
        description: String(f?.CodeDescription || "").trim() || null,
        severity: mapSeverity(f?.CodeSeverity),
        active: true,
      })).filter((f) => f.code)
    : [];

  return {
    asset_code: assetCode || undefined,
    device_serial: deviceSerial || undefined,
    recorded_at: recordedAt,
    engine_hours: engineHours,
    idle_hours: numOrNull(idle.Hour),
    latitude: numOrNull(loc.Latitude),
    longitude: numOrNull(loc.Longitude),
    ignition_on: engine.Running === true,
    faults,
    source: "cat_iso15143",
    payload: {
      make,
      model,
      cat_serial: catSerial,
      equipment_id: equipmentId,
      snapshot_time: snapshotTime,
      equipment: eq,
    },
  };
}

function normalizeLegacyRecord(rec, config) {
  const deviceSerial = String(pick(rec, [
    "deviceSerialNumber", "device_serial", "deviceSerial", "deviceId", "serialNumber", "serial",
  ]) || "").trim().toUpperCase();
  const sourceAssetCode = String(pick(rec, [
    "asset_code", "assetCode", "equipmentId", "equipment_id",
  ]) || "").trim().toUpperCase();
  const assetCode = config.assetMap[deviceSerial] || config.assetMap[sourceAssetCode] || sourceAssetCode;
  const engineHours = numOrNull(pick(rec, [
    "engineHours", "engine_hours", "serviceHours", "meter_hours", "hours", "Hour",
  ]));
  if (!assetCode && !deviceSerial) return null;
  if (engineHours == null) return null;
  return {
    asset_code: assetCode || undefined,
    device_serial: deviceSerial || undefined,
    recorded_at: String(pick(rec, ["timestamp", "recorded_at", "datetime"]) || new Date().toISOString()),
    engine_hours: engineHours,
    source: "cat_custom",
    payload: rec,
  };
}

async function fetchAccessToken(config) {
  if (!config.clientId || !config.clientSecret) {
    throw new Error("VISIONLINK_CLIENT_ID and VISIONLINK_CLIENT_SECRET are required once CAT issues credentials");
  }

  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: config.clientId,
    client_secret: config.clientSecret,
  });
  if (config.scope) body.set("scope", config.scope);

  const res = await fetch(config.tokenUrl, {
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

async function fetchCatJson(url, token) {
  const res = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: ISO_ACCEPT,
      "X-Cat-API-Tracking-Id": randomUUID(),
    },
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`CAT API request failed (${res.status}) ${url}: ${txt.slice(0, 500)}`);
  }
  return res.json();
}

function nextFleetPage(links, currentPage) {
  const next = (links || []).find((l) => String(l?.Rel || l?.rel || "").toLowerCase() === "next");
  if (!next) return null;
  const href = String(next.Href || next.href || "").trim();
  if (href) return href;
  return null;
}

async function fetchIsoPayload(config, token) {
  if (config.customDataUrl) {
    const res = await fetch(config.customDataUrl, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: `${ISO_ACCEPT}, application/json`,
        "X-Cat-API-Tracking-Id": randomUUID(),
      },
    });
    if (!res.ok) {
      const txt = await res.text();
      throw new Error(`Custom data request failed (${res.status}): ${txt.slice(0, 500)}`);
    }
    return res.json();
  }

  if (config.mode === "equipment") {
    const [url] = plannedUrls(config);
    return fetchCatJson(url, token);
  }

  const allEquipment = [];
  let pageUrl = `${config.baseUrl}/fleet/${config.fleetStartPage}`;
  let snapshotTime = null;
  let guard = 0;
  while (pageUrl && guard < 100) {
    guard += 1;
    const payload = await fetchCatJson(pageUrl, token);
    const parsed = extractEquipmentList(payload);
    allEquipment.push(...parsed.equipment);
    snapshotTime = parsed.snapshotTime || snapshotTime;
    const nextUrl = nextFleetPage(parsed.links, guard);
    pageUrl = nextUrl;
  }

  return {
    Equipment: allEquipment,
    SnapshotTime: snapshotTime,
  };
}

function recordsFromPayload(payload, config) {
  const { equipment, snapshotTime } = extractEquipmentList(payload);
  if (equipment.length) {
    return equipment
      .map((eq) => normalizeIsoEquipment(eq, config, snapshotTime))
      .filter(Boolean);
  }

  const legacyRows = Array.isArray(payload) ? payload : [];
  return legacyRows
    .map((r) => normalizeLegacyRecord(r, config))
    .filter(Boolean);
}

async function postToIronlog(records, config) {
  const ingestUrl = `${config.apiBase}/api/telematics/ingest`;
  let ok = 0;
  let failed = 0;

  for (const row of records) {
    const payload = {
      asset_code: row.asset_code,
      device_serial: row.device_serial,
      recorded_at: row.recorded_at,
      engine_hours: row.engine_hours,
      idle_hours: row.idle_hours,
      latitude: row.latitude,
      longitude: row.longitude,
      ignition_on: row.ignition_on,
      faults: row.faults,
      source: row.source,
      upstream_payload: row.payload,
    };

    if (config.dryRun) {
      console.log("[visionlink] dry-run ingest", JSON.stringify(payload, null, 2));
      ok += 1;
      continue;
    }

    const res = await fetch(ingestUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...(config.apiKey ? { "x-api-key": config.apiKey } : {}),
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
  const config = getConfig();
  const urls = plannedUrls(config);

  console.log("[visionlink] mode:", config.mode);
  console.log("[visionlink] token:", config.tokenUrl);
  console.log("[visionlink] data:", urls.join(" | "));

  if (config.dryRun && (!config.clientId || !config.clientSecret)) {
    console.log("[visionlink] dry-run only (no CAT credentials yet) — configure api/.env when CAT replies");
    console.log("[visionlink] expected env:");
    console.log("  VISIONLINK_CLIENT_ID=...");
    console.log("  VISIONLINK_CLIENT_SECRET=...");
    console.log("  VISIONLINK_EQUIPMENT_SERIAL=JZ400729");
    console.log("  VISIONLINK_ASSET_MAP_JSON={\"JZ400729\":\"<IRONLOG_ASSET_CODE>\"}");
    console.log("  VISIONLINK_DEVICE_MAP_JSON={\"JZ400729\":\"23032800H0ZR0054\"}");
    return;
  }

  const token = await fetchAccessToken(config);
  const payload = await fetchIsoPayload(config, token);
  const records = recordsFromPayload(payload, config);

  if (!records.length) {
    console.log("[visionlink] no usable equipment records found");
    return;
  }

  const result = await postToIronlog(records, config);
  console.log(`[visionlink] processed=${records.length} ingested=${result.ok} failed=${result.failed}`);
}

main().catch((err) => {
  console.error("[visionlink] bridge failed:", err?.message || err);
  process.exitCode = 1;
});
