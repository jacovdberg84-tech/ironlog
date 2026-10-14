import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-fleet-sec-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
const Fastify = (await import("fastify")).default;
const { db } = await import("../db/client.js");
const { ensureCartrackTables } = await import("../utils/cartrack.js");
const { ensureUnitechTables } = await import("../utils/unitech.js");
const { buildFleetSecurityReport, trackerTimeMs } = await import("../utils/fleetSecurityReport.js");
const cartrackRoutes = (await import("../routes/cartrack.routes.js")).default;
process.on("exit", () => { try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

db.exec(`
  CREATE TABLE IF NOT EXISTS assets (id INTEGER PRIMARY KEY, asset_code TEXT UNIQUE, asset_name TEXT, category TEXT, active INTEGER DEFAULT 1, archived INTEGER DEFAULT 0);
  INSERT INTO assets (asset_code, asset_name, category) VALUES
    ('V05AM', 'Toyota Hilux 2.8', 'LDV'), ('V06AM', 'Toyota Land Cruiser', 'LDV'), ('B01AM', 'Staff bus', 'Bus'),
    ('V09AM', 'Hilux (spare)', 'LDV'), ('E501AM', 'Excavator', 'Excavator');
`);
ensureCartrackTables();
ensureUnitechTables();
// "Now" is 2026-10-14 10:00 site time.
const now = new Date("2026-10-14T08:00:00Z");
const snap = db.prepare(`INSERT INTO cartrack_vehicle_snapshots (registration, asset_code, vehicle_name, last_event_at, synced_at, tracker_battery_pct, supply_voltage_v) VALUES (?, ?, ?, ?, ?, ?, ?)`);
snap.run("AJM753MP", "V05AM", "Hilux", "2026-10-14 09:30:00", "2026-10-14 07:31:00", 90, 12.6);       // fine
snap.run("AKB100MP", "V06AM", "Cruiser", "2026-10-10 18:00:00", "2026-10-14 07:31:00", 10, 0);        // quiet 4 days, no power
snap.run("XYZ999", "XYZ999", "Unknown bakkie", "2026-10-13 20:00:00", "2026-10-14 07:31:00", null, 11.1); // not linked to a fleet number
db.prepare(`INSERT INTO unitech_vehicle_snapshots (registration, asset_code, vehicle_name, position_age_days, last_event_at) VALUES (?, ?, ?, ?, ?)`)
  .run("AJM753MP-U", "V05AM", "Hilux site", 0.1, "2026-10-14 09:50:00");
db.prepare(`INSERT INTO unitech_vehicle_snapshots (registration, asset_code, vehicle_name, position_age_days, last_event_at) VALUES (?, ?, ?, ?, ?)`)
  .run("BUS01", "B01AM", "Bus", 2.5, "2026-10-11 22:00:00");

const ev = db.prepare(`INSERT INTO cartrack_events (event_id, registration, asset_code, event_type, event_type_label, event_time, speed_kmh, speed_limit_kmh, latitude, longitude, driver_name, is_speeding)
  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);
let n = 0;
const add = (reg, code, type, label, time, speed = 0, limit = null, speeding = 0) => ev.run(`e${++n}`, reg, code, type, label, time, speed, limit, -10.77, 40.47, "Jose", speeding);
add("AKB100MP", "V06AM", "tamper", "Tracker tamper alert", "2026-10-10 17:55:00");
add("AKB100MP", "V06AM", "power", "Main power disconnected", "2026-10-10 17:58:00");
add("AJM753MP", "V05AM", "ign", "Ignition on", "2026-10-12 22:15:00");
add("AJM753MP", "V05AM", "pos", "Position", "2026-10-12 23:40:00", 64);
add("AJM753MP", "V05AM", "pos", "Position", "2026-10-12 14:00:00", 70);          // daytime: not night
add("AJM753MP", "V05AM", "harsh", "Harsh braking", "2026-10-12 03:00:00", 0);     // night, but not moving
add("AJM753MP", "V05AM", "ironlog_speed_alert", "Speed ≥ 100 km/h", "2026-10-12 15:00:00", 121, 100, 1);
add("AJM753MP-U", "V05AM", "ironlog_speed_alert", "Speed ≥ 61 km/h", "2026-10-13 09:00:00", 72, 61, 1);
add("AJM753MP", "V05AM", "tamper", "Tracker tamper alert", "2026-09-30 10:00:00");  // before the period

test("tracking unit per vehicle, and whether each is reporting", () => {
  const r = buildFleetSecurityReport(db, { from: "2026-10-08", to: "2026-10-14", now });
  const by = Object.fromEntries(r.units.map((u) => [u.asset_code, u]));
  assert.equal(by.V05AM.tracking_unit, "Cartrack + Unitech");
  assert.equal(by.V05AM.registration, "AJM753MP, AJM753MP-U");
  assert.equal(by.V05AM.status_label, "Reporting");
  assert.equal(by.V05AM.description, "Toyota Hilux 2.8");
  assert.equal(by.V06AM.tracking_unit, "Cartrack");
  assert.equal(by.V06AM.status_label, "Not reporting for 3 days");
  assert.equal(by.B01AM.tracking_unit, "Unitech");
  assert.equal(by.B01AM.status_label, "No report for 60 h");
  assert.equal(by.XYZ999.linked, false, "unlinked units still show, under their registration");
  assert.equal(by.XYZ999.description, "Unknown bakkie (not linked to a fleet number)");
  assert.deepEqual(r.untracked.map((a) => a.asset_code), ["V09AM"], "vehicles with no tracker; plant is not listed");
  assert.deepEqual([r.summary.cartrack, r.summary.unitech, r.summary.both], [2, 1, 1]);
});

test("security issues: alerts, night movement, speeding, quiet and unpowered trackers", () => {
  const r = buildFleetSecurityReport(db, { from: "2026-10-08", to: "2026-10-14", now, notes: "  Stone thrown at V05AM near Mute, 12 Oct.  " });
  assert.deepEqual(r.issues.alerts.map((a) => [a.vehicle, a.event]), [["V06AM", "Tracker tamper alert"], ["V06AM", "Main power disconnected"]],
    "only alerts in the period");
  assert.deepEqual(r.issues.night, [{ date: "2026-10-12", vehicle: "V05AM", first: "22:15", last: "23:40", events: 2, max_speed: 64 }]);
  assert.deepEqual(r.issues.speeding.map((s) => [s.vehicle, s.tracker, s.events, s.max_speed]), [["V05AM", "Cartrack + Unitech", 2, 121]]);
  assert.deepEqual(r.issues.offline.map((o) => [o.vehicle, o.tracking_unit]), [["B01AM", "Unitech"], ["V06AM", "Cartrack"]]);
  assert.deepEqual(r.issues.power.map((p) => [p.vehicle, p.issue.split(" (")[0]]), [
    ["V06AM", "No power to tracker"], ["V06AM", "Tracker backup battery low"], ["XYZ999", "Low vehicle battery at tracker"],
  ]);
  assert.equal(r.notes, "Stone thrown at V05AM near Mute, 12 Oct.");
  assert.ok(r.event_types.find((t) => t.label === "Harsh braking" && !t.security), "every event type is listed in the appendix");

  const early = buildFleetSecurityReport(db, { from: "2026-10-08", to: "2026-10-14", now, nightStart: 23, nightEnd: 4 });
  assert.equal(early.issues.night[0].first, "23:40", "night hours can be changed");
});

test("PDF text drops characters the built-in fonts can't print", async () => {
  const { pdfSafe } = await import("../utils/fleetSecurityReport.js");
  assert.equal(pdfSafe("Speed ≥ 100 km/h — João’s “bakkie” 🚗"), 'Speed >= 100 km/h - João\'s "bakkie" ');
});

test("tracker times without a zone are site time", () => {
  assert.equal(trackerTimeMs("2026-10-14 10:00:00"), Date.parse("2026-10-14T08:00:00Z"));
  assert.equal(trackerTimeMs("2026-10-14 08:00:00", { utc: true }), Date.parse("2026-10-14T08:00:00Z"));
  assert.equal(trackerTimeMs("2026-10-14T10:00:00+02:00"), Date.parse("2026-10-14T08:00:00Z"));
});

test("the PDF is for managers and plant admin only", async (t) => {
  const app = Fastify();
  await app.register(cartrackRoutes, { prefix: "/api/cartrack" });
  t.after(() => app.close());
  const get = (role, qs) => app.inject({ method: "GET", url: `/api/cartrack/security-report.pdf?${qs}`, headers: { "x-user-roles": role } });
  assert.equal((await get("artisan", "from=2026-10-08&to=2026-10-14")).statusCode, 403);
  assert.equal((await get("admin", "from=2026-10-14&to=2026-10-08")).statusCode, 400);
  const pdf = await get("workshop_admin", "from=2026-10-08&to=2026-10-14&notes=Road%20block%20at%20Palma");
  assert.equal(pdf.statusCode, 200);
  assert.equal(pdf.headers["content-type"], "application/pdf");
  assert.equal(pdf.rawPayload.subarray(0, 5).toString(), "%PDF-");
  const json = (await get("admin", "from=2026-10-08&to=2026-10-14&format=json&night_start=21")).json();
  assert.equal(json.report.night_hours, "21:00 to 05:00");
});
