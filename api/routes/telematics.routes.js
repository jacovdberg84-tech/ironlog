// IRONLOG/api/routes/telematics.routes.js
// FSC telematics ingest + live fleet API (pilot: A300AM, F500AM).

import {
  ensurePilotDevices,
  ensureTelematicsTables,
  getDeviceByAssetCode,
  ingestTelematicsPayload,
  listFleetSnapshots,
  listRecentFaults,
  getActiveFaultSummary,
  listTelematicsDevices,
  upsertTelematicsDevice,
  deactivateTelematicsDevice,
  syncTelematicsDailyHours,
} from "../utils/telematics.js";
import { db } from "../db/client.js";

function getRoles(req) {
  const many = String(req.headers["x-user-roles"] || "")
    .split(",").map((x) => String(x || "").trim().toLowerCase()).filter(Boolean);
  const one = String(req.headers["x-user-role"] || "")
    .split(",").map((x) => String(x || "").trim().toLowerCase()).filter(Boolean);
  return Array.from(new Set([...many, ...one]));
}

function hasAnyRole(req, allowed) {
  return getRoles(req).some((r) => allowed.includes(r));
}

function requireRoles(req, reply, allowed) {
  if (!hasAnyRole(req, allowed)) {
    reply.code(403).send({ error: "not allowed" });
    return false;
  }
  return true;
}

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function checkIngestAuth(req, reply) {
  const expected = String(process.env.FSC_TELEMATICS_API_KEY || process.env.TELEMATICS_API_KEY || "").trim();
  if (!expected) return true;
  const headerKey = String(req.headers["x-api-key"] || req.headers["x-fsc-api-key"] || "").trim();
  const queryKey = String(req.query?.api_key || "").trim();
  if (headerKey === expected || queryKey === expected) return true;
  reply.code(401).send({ error: "invalid telematics API key" });
  return false;
}

export default async function telematicsRoutes(app) {
  ensureTelematicsTables();
  if (String(process.env.TELEMATICS_AUTO_PILOT || "").trim() === "1") {
    ensurePilotDevices();
  }

  // POST /api/telematics/ingest — FSC webhook / poll push (API key when configured)
  app.post("/ingest", async (req, reply) => {
    if (!checkIngestAuth(req, reply)) return;
    const out = ingestTelematicsPayload(req.body);
    if (!out.ok) return reply.code(404).send(out);
    return reply.send(out);
  });

  // GET /api/telematics/fleet — live fleet cards for dashboard
  app.get("/fleet", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager", "operator", "artisan", "stores"])) return;
    const fleet = listFleetSnapshots();
    const faults = listRecentFaults(30);
    const active = getActiveFaultSummary();
    return reply.send({
      ok: true,
      fleet,
      recent_faults: faults,
      active_fault_count: active.fault_count,
      units_with_faults: active.units_with_faults,
    });
  });

  // GET /api/telematics/faults/active — lightweight poll for global fault banner
  app.get("/faults/active", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager", "operator", "artisan", "stores"])) return;
    const summary = getActiveFaultSummary();
    return reply.send({ ok: true, ...summary });
  });

  // GET /api/telematics/assets/:asset_code
  app.get("/assets/:asset_code", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager", "operator", "artisan", "stores"])) return;
    const assetCode = String(req.params.asset_code || "").trim();
    const device = getDeviceByAssetCode(assetCode);
    if (!device) return reply.code(404).send({ error: "No telematics device for asset" });
    const snapshot = db.prepare(`SELECT * FROM telematics_snapshots WHERE asset_id = ?`).get(device.asset_id);
    const faults = db.prepare(`
      SELECT id, event_time, fault_code, description, severity, active
      FROM telematics_events
      WHERE asset_id = ? AND event_type = 'fault'
      ORDER BY event_time DESC, id DESC
      LIMIT 50
    `).all(device.asset_id);
    return reply.send({ ok: true, device, snapshot: snapshot || null, faults });
  });

  // POST /api/telematics/sync-daily?work_date=YYYY-MM-DD — reconcile all telematics assets
  app.post("/sync-daily", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;
    const workDate = String(req.query?.work_date || req.body?.work_date || new Date().toISOString().slice(0, 10)).trim();
    if (!isDate(workDate)) return reply.code(400).send({ error: "work_date must be YYYY-MM-DD" });
    const devices = db.prepare(`SELECT asset_id, device_serial FROM telematics_devices WHERE active = 1`).all();
    const results = devices.map((d) => ({
      device_serial: d.device_serial,
      ...syncTelematicsDailyHours(d.asset_id, workDate),
    }));
    return reply.send({ ok: true, work_date: workDate, results });
  });

  // GET /api/telematics/devices
  app.get("/devices", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;
    const includeInactive = String(req.query?.all || "").trim() === "1";
    const devices = listTelematicsDevices({ includeInactive });
    return reply.send({ ok: true, devices });
  });

  // POST /api/telematics/devices — register, add, or replace FSC unit on an asset
  app.post("/devices", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const out = upsertTelematicsDevice({
      assetCode: req.body?.asset_code,
      deviceSerial: req.body?.device_serial,
      unitModel: req.body?.unit_model,
      externalId: req.body?.external_id,
      replaceFaulty: req.body?.replace_faulty === true
        || String(req.body?.replace_faulty || "").trim() === "1",
    });
    if (!out.ok) {
      const code = String(out.error || "").includes("not found") ? 404 : 400;
      return reply.code(code).send(out);
    }
    return reply.send(out);
  });

  // POST /api/telematics/devices/:id/deactivate — retire a unit mapping
  app.post("/devices/:id/deactivate", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const out = deactivateTelematicsDevice(req.params?.id);
    if (!out.ok) return reply.code(out.error === "Device not found" ? 404 : 400).send(out);
    return reply.send(out);
  });
}
