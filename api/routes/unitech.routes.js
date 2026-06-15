// Unitech Moz (GpsGate KML) — Afungi site fleet feed.

import {
  ensureUnitechTables,
  getUnitechPublicSettings,
  saveUnitechSettings,
  syncUnitechFleetStatus,
  listUnitechFleetFromDb,
  enrichUnitechLiveRow,
  testUnitechConnection,
} from "../utils/unitech.js";

function getRoles(req) {
  const many = String(req.headers["x-user-roles"] || "")
    .split(",").map((x) => String(x || "").trim().toLowerCase()).filter(Boolean);
  const one = String(req.headers["x-user-role"] || "")
    .split(",").map((x) => String(x || "").trim().toLowerCase()).filter(Boolean);
  return Array.from(new Set([...many, ...one]));
}

function requireRoles(req, reply, allowed) {
  if (!getRoles(req).some((r) => allowed.includes(r))) {
    reply.code(403).send({ ok: false, error: "not allowed" });
    return false;
  }
  return true;
}

export default async function unitechRoutes(app) {
  ensureUnitechTables();

  app.get("/settings", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;
    return reply.send({ ok: true, settings: getUnitechPublicSettings() });
  });

  app.put("/settings", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const updated_by = String(req.headers["x-user-name"] || "admin").trim();
    const settings = saveUnitechSettings({
      kml_feed_url: req.body?.kml_feed_url,
      feed_label: req.body?.feed_label,
      enabled: req.body?.enabled,
      updated_by,
    });
    return reply.send({ ok: true, settings });
  });

  app.post("/test-connection", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    try {
      const result = await testUnitechConnection();
      return reply.send({ ok: true, ...result });
    } catch (err) {
      return reply.code(err.code === "NOT_CONFIGURED" ? 400 : 502).send({
        ok: false,
        error: String(err.message || err),
      });
    }
  });

  app.post("/sync", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;
    try {
      const status = await syncUnitechFleetStatus();
      return reply.send({ ok: true, status });
    } catch (err) {
      return reply.code(502).send({ ok: false, error: String(err.message || err) });
    }
  });

  app.get("/fleet", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager", "operator", "artisan", "stores"])) return;
    const settings = getUnitechPublicSettings();
    const fleet = listUnitechFleetFromDb().map((v) => enrichUnitechLiveRow(v));
    const fresh = fleet.filter((v) => v.has_gps && !v.position_stale).length;
    return reply.send({
      ok: true,
      configured: settings.configured,
      enabled: settings.enabled,
      feed_label: settings.feed_label,
      summary: {
        total_vehicles: fleet.length,
        with_gps: fleet.filter((v) => v.has_gps).length,
        fresh_positions: fresh,
        last_sync: fleet[0]?.synced_at || null,
      },
      fleet,
    });
  });
}
