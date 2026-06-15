// Cartrack Fleet API routes — dashboard fleet, events, morning speeding reports.

import { db } from "../db/client.js";
import {
  ensureCartrackTables,
  getCartrackPublicSettings,
  saveCartrackSettings,
  syncCartrackFleetStatus,
  syncCartrackEvents,
  listCartrackFleetFromDb,
  listCartrackEventsFromDb,
  buildMorningSpeedingReport,
  formatMorningReportText,
  runCartrackMorningJob,
  sendCartrackMorningEmail,
  cartrackApiGet,
  enrichCartrackLiveRow,
} from "../utils/cartrack.js";
import { buildPdfBuffer, sectionTitle, table } from "../utils/pdfGenerator.js";
import {
  getUnitechPublicSettings,
  syncUnitechFleetStatus,
  listUnitechFleetFromDb,
  enrichUnitechLiveRow,
} from "../utils/unitech.js";

function mergeGpsFleet({ cartrackRows, unitechRows, speedRegs }) {
  const cartrackFleet = cartrackRows.map((v) => enrichCartrackLiveRow(v, speedRegs));
  const unitechFleet = unitechRows.map((v) => enrichUnitechLiveRow(v, speedRegs));
  return [...cartrackFleet, ...unitechFleet];
}

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

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

async function emailMorningReport(summary, recipients) {
  return sendCartrackMorningEmail(summary, recipients);
}

export default async function cartrackRoutes(app) {
  ensureCartrackTables();

  app.get("/settings", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;
    return reply.send({ ok: true, settings: getCartrackPublicSettings() });
  });

  app.put("/settings", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const updated_by = String(req.headers["x-user-name"] || "admin").trim();
    const settings = saveCartrackSettings({
      base_url: req.body?.base_url,
      username: req.body?.username,
      password: req.body?.password,
      morning_recipients: req.body?.morning_recipients,
      morning_enabled: req.body?.morning_enabled,
      morning_hour: req.body?.morning_hour,
      morning_minute: req.body?.morning_minute,
      speed_alert_kmh: req.body?.speed_alert_kmh,
      updated_by,
    });
    return reply.send({ ok: true, settings });
  });

  app.post("/test-connection", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    try {
      const body = await cartrackApiGet("vehicles/status", { per_page: 10 });
      const vehicles = Array.isArray(body?.data) ? body.data : Array.isArray(body) ? body : [];
      const count = vehicles.length;
      const message = count
        ? `Cartrack connected — ${count} vehicle(s) visible to API user.`
        : "Cartrack connected (auth OK) but no vehicles returned. Ask Cartrack to assign your fleet to this API user.";
      return reply.send({ ok: true, message, vehicle_count: count });
    } catch (err) {
      return reply.code(err.code === "NOT_CONFIGURED" ? 400 : 502).send({
        ok: false,
        error: String(err.message || err),
      });
    }
  });

  app.get("/live", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager", "operator", "artisan", "stores"])) return;
    const refresh = String(req.query?.refresh || "1").trim() !== "0";
    let syncError = null;
    let unitechSyncError = null;
    if (refresh) {
      try {
        await syncCartrackFleetStatus();
      } catch (err) {
        syncError = String(err.message || err);
      }
      const unitechSettings = getUnitechPublicSettings();
      if (unitechSettings.configured && unitechSettings.enabled) {
        try {
          await syncUnitechFleetStatus();
        } catch (err) {
          unitechSyncError = String(err.message || err);
        }
      }
    }
    const settings = getCartrackPublicSettings();
    const unitechSettings = getUnitechPublicSettings();
    const fleet = listCartrackFleetFromDb();
    const unitechFleet = listUnitechFleetFromDb();
    const today = new Date().toISOString().slice(0, 10);
    const speedingToday = listCartrackEventsFromDb({
      startDate: today,
      endDate: today,
      speedingOnly: true,
      limit: 200,
    });
    const speedRegs = new Set(speedingToday.map((e) => e.registration || e.asset_code));
    const positioned = mergeGpsFleet({ cartrackRows: fleet, unitechRows: unitechFleet, speedRegs });
    const live = positioned.filter((v) => Number(v.ignition_on) === 1).length;
    const syncTimes = [
      fleet[0]?.synced_at,
      unitechFleet[0]?.synced_at,
    ].filter(Boolean).sort().reverse();
    return reply.send({
      ok: true,
      configured: settings.configured || unitechSettings.configured,
      base_url: settings.base_url,
      refreshed: refresh && !syncError && !unitechSyncError,
      sync_error: syncError,
      unitech_sync_error: unitechSyncError,
      unitech: unitechSettings,
      summary: {
        total_vehicles: positioned.length,
        cartrack_vehicles: fleet.length,
        unitech_vehicles: unitechFleet.length,
        with_gps: positioned.filter((v) => v.has_gps).length,
        ignition_on: live,
        speeding_today: speedingToday.length,
        last_sync: syncTimes[0] || null,
      },
      fleet: positioned,
      speeding_today: speedingToday,
    });
  });

  app.get("/fleet", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager", "operator", "artisan", "stores"])) return;
    const settings = getCartrackPublicSettings();
    const unitechSettings = getUnitechPublicSettings();
    const fleet = listCartrackFleetFromDb();
    const unitechFleet = listUnitechFleetFromDb();
    const today = new Date().toISOString().slice(0, 10);
    const speedingToday = listCartrackEventsFromDb({
      startDate: today,
      endDate: today,
      speedingOnly: true,
      limit: 50,
    });
    const speedRegs = new Set(speedingToday.map((e) => e.registration || e.asset_code));
    const enriched = mergeGpsFleet({ cartrackRows: fleet, unitechRows: unitechFleet, speedRegs });
    const live = enriched.filter((v) => Number(v.ignition_on) === 1).length;
    const syncTimes = [
      fleet[0]?.synced_at,
      unitechFleet[0]?.synced_at,
    ].filter(Boolean).sort().reverse();
    return reply.send({
      ok: true,
      configured: settings.configured || unitechSettings.configured,
      base_url: settings.base_url,
      unitech: unitechSettings,
      summary: {
        total_vehicles: enriched.length,
        cartrack_vehicles: fleet.length,
        unitech_vehicles: unitechFleet.length,
        ignition_on: live,
        speeding_today: speedingToday.length,
        last_sync: syncTimes[0] || null,
      },
      fleet: enriched,
      speeding_today: speedingToday,
    });
  });

  app.post("/sync", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;
    try {
      const status = await syncCartrackFleetStatus();
      let unitech = null;
      const unitechSettings = getUnitechPublicSettings();
      if (unitechSettings.configured && unitechSettings.enabled) {
        try {
          unitech = await syncUnitechFleetStatus();
        } catch (err) {
          unitech = { error: String(err.message || err) };
        }
      }
      const start = String(req.body?.start_date || req.query?.start_date || new Date().toISOString().slice(0, 10));
      const end = String(req.body?.end_date || req.query?.end_date || start);
      const events = await syncCartrackEvents({ startDate: start, endDate: end });
      return reply.send({ ok: true, status, unitech, events });
    } catch (err) {
      return reply.code(502).send({ ok: false, error: String(err.message || err) });
    }
  });

  app.get("/events", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager", "operator", "artisan", "stores"])) return;
    const start = String(req.query?.start || req.query?.start_date || new Date().toISOString().slice(0, 10));
    const end = String(req.query?.end || req.query?.end_date || start);
    const speedingOnly = String(req.query?.speeding_only || "").trim() === "1";
    const rows = listCartrackEventsFromDb({ startDate: start, endDate: end, speedingOnly, limit: 300 });
    return reply.send({ ok: true, start_date: start, end_date: end, rows });
  });

  app.get("/morning-report", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;
    const date = String(req.query?.date || "").trim();
    if (!isDate(date)) {
      return reply.code(400).send({ ok: false, error: "date must be YYYY-MM-DD" });
    }
    const cached = db.prepare(`SELECT summary_json FROM cartrack_morning_reports WHERE report_date = ?`).get(date);
    if (cached?.summary_json) {
      try {
        return reply.send({ ok: true, summary: JSON.parse(String(cached.summary_json)) });
      } catch {}
    }
    const summary = buildMorningSpeedingReport(date);
    return reply.send({ ok: true, summary });
  });

  app.get("/morning-report.pdf", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;
    const date = String(req.query?.date || "").trim();
    if (!isDate(date)) return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
    try {
      const summary = buildMorningSpeedingReport(date);
      const pdf = await buildPdfBuffer((doc) => {
        sectionTitle(doc, `Cartrack speeding report — ${date}`);
        doc.fontSize(10).fillColor("#334155");
        doc.text(`Alert threshold: ${summary.speed_alert_kmh ?? 100} km/h (IRONLOG live GPS — Cartrack API events are typically 160+ km/h only)`);
        doc.text(`Total events: ${summary.total_speeding_events} · Vehicles: ${summary.vehicles_with_speeding}`);
        doc.moveDown(0.5);
        const columns = [
          { key: "time", label: "Time", width: 72 },
          { key: "vehicle", label: "Vehicle", width: 56 },
          { key: "speed", label: "Speed", width: 36 },
          { key: "limit", label: "Limit", width: 36 },
          { key: "driver", label: "Driver", width: 56 },
          { key: "type", label: "Type", width: 72 },
        ];
        const rows = (summary.events || []).slice(0, 80).map((e) => ({
          time: String(e.event_time || "").slice(0, 16),
          vehicle: e.asset_code || e.registration || "",
          speed: e.speed_kmh != null ? `${Number(e.speed_kmh).toFixed(0)}` : "—",
          limit: e.speed_limit_kmh != null ? `${Number(e.speed_limit_kmh).toFixed(0)}` : "—",
          driver: e.driver_name || "",
          type: e.event_type_label || e.event_type || "",
        }));
        if (!rows.length) {
          doc.fontSize(10).fillColor("#64748b");
          doc.text("No speeding events recorded for this date.");
          return;
        }
        table(doc, columns, rows, { fontSize: 8, compact: true });
      }, { title: "Cartrack Speeding Report", layout: "landscape" });
      reply.header("Content-Type", "application/pdf");
      reply.header("Content-Disposition", `inline; filename="cartrack_speeding_${date}.pdf"`);
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: String(err.message || err) });
    }
  });

  app.post("/morning-report/send", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;
    const date = String(req.body?.date || req.query?.date || "").trim();
    if (!isDate(date)) return reply.code(400).send({ ok: false, error: "date must be YYYY-MM-DD" });
    const recipients = String(req.body?.recipients || "")
      .split(",")
      .map((x) => x.trim())
      .filter(Boolean);
    const settings = getCartrackPublicSettings();
    const to = recipients.length
      ? recipients
      : String(settings.morning_recipients || "").split(",").map((x) => x.trim()).filter(Boolean);
    if (!to.length) {
      return reply.code(400).send({ ok: false, error: "No email recipients. Set morning recipients in Cartrack settings." });
    }
    try {
      await syncCartrackFleetStatus();
      await syncCartrackEvents({ startDate: date, endDate: date });
      const summary = buildMorningSpeedingReport(date);
      const mail = await emailMorningReport(summary, to);
      if (!mail.ok) return reply.code(400).send(mail);
      return reply.send({ ok: true, summary, emailed_to: to });
    } catch (err) {
      return reply.code(502).send({ ok: false, error: String(err.message || err) });
    }
  });

  app.post("/morning-report/run", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const date = String(req.body?.date || "").trim();
    const reportDate = isDate(date)
      ? date
      : (() => {
          const d = new Date();
          d.setDate(d.getDate() - 1);
          return d.toISOString().slice(0, 10);
        })();
    try {
      const result = await runCartrackMorningJob({ reportDate, sendEmail: false });
      const settings = getCartrackPublicSettings();
      const to = String(settings.morning_recipients || "").split(",").map((x) => x.trim()).filter(Boolean);
      if (to.length && req.body?.send_email !== false) {
        await emailMorningReport(result.summary, to);
      }
      return reply.send({ ok: true, ...result });
    } catch (err) {
      return reply.code(502).send({ ok: false, error: String(err.message || err) });
    }
  });
}
