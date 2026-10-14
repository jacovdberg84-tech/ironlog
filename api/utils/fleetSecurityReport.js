// IRONLOG/api/utils/fleetSecurityReport.js — fleet tracking security report.
//
// For site security: which tracking unit each vehicle carries (Cartrack,
// Unitech or both), whether it is still reporting, and the security issues
// the trackers recorded over a period:
//   - tracker alerts from Cartrack (tamper, panic, power cut, towing, no-go zone…)
//   - vehicles moving at night (outside the allowed hours)
//   - speeding (Cartrack and Unitech)
//   - trackers that stopped reporting, and tracker power problems
// Everything is read from the tables the Fleet Track sync already fills;
// nothing is stored here.
import { enrichCartrackLiveRow } from "./cartrack.js";

// Cartrack event names that point to a security problem rather than driving.
export const SECURITY_EVENT_RE = /tamper|panic|sos|duress|power|disconnect|unplug|tow|jam|unauthori[sz]ed|no.?go|geo.?fence|zone|impact|accident|crash|theft|stolen|immobili[sz]|alarm|door|hijack|gps (?:lost|antenna)|antenna/i;
// Road and site vehicles, as opposed to plant (excavators, dozers…), for the "no tracker" list.
export const VEHICLE_CATEGORY_RE = /\b(ldv|light|vehicle|bakkie|pick.?up|bus|truck|lorry|van|car|ambulance|tanker|tipper|lowbed|horse)\b/i;

const SITE_OFFSET = "+02:00"; // Africa/Maputo, no daylight saving

function hasTable(db, name) {
  return Boolean(db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?`).get(name));
}

/** A time from the trackers as milliseconds; times without a zone are site time. */
export function trackerTimeMs(value, { utc = false } = {}) {
  const s = String(value || "").trim();
  if (!s) return NaN;
  if (/[zZ]|[+-]\d{2}:?\d{2}$/.test(s)) return Date.parse(s);
  const iso = s.replace(" ", "T");
  return Date.parse(`${iso.length === 16 ? `${iso}:00` : iso}${utc ? "Z" : SITE_OFFSET}`);
}

/** Site-time "YYYY-MM-DD HH:MM" for a millisecond time. */
function siteStamp(ms) {
  if (!Number.isFinite(ms)) return "—";
  return new Date(ms + 2 * 3600000).toISOString().replace("T", " ").slice(0, 16);
}

function reportingState(hoursSince) {
  if (!Number.isFinite(hoursSince)) return { status: "unknown", label: "No report on record" };
  if (hoursSince < 24) return { status: "ok", label: "Reporting" };
  if (hoursSince < 72) return { status: "late", label: `No report for ${Math.round(hoursSince)} h` };
  return { status: "offline", label: `Not reporting for ${Math.floor(hoursSince / 24)} days` };
}

/** One row per vehicle with the tracking unit(s) it carries. */
export function trackedUnits(db, { now = new Date() } = {}) {
  const assets = hasTable(db, "assets")
    ? new Map(db.prepare(`SELECT asset_code, asset_name, category FROM assets`).all().map((a) => [String(a.asset_code).toUpperCase(), a]))
    : new Map();
  const byVehicle = new Map();
  const add = (key, unit) => {
    const e = byVehicle.get(key) || { units: [] };
    e.units.push(unit);
    byVehicle.set(key, e);
  };
  if (hasTable(db, "cartrack_vehicle_snapshots")) {
    for (const raw of db.prepare(`SELECT * FROM cartrack_vehicle_snapshots`).all()) {
      const r = enrichCartrackLiveRow(raw);
      const lastMs = Number.isFinite(trackerTimeMs(r.last_event_at)) ? trackerTimeMs(r.last_event_at) : trackerTimeMs(r.synced_at, { utc: true });
      add(String(r.asset_code || r.registration).toUpperCase(), {
        provider: "Cartrack",
        registration: r.registration,
        vehicle_name: r.vehicle_name,
        last_ms: lastMs,
        hours_since: (now.getTime() - lastMs) / 3600000,
        tracker_battery_pct: r.tracker_battery_pct,
        supply_voltage_v: r.supply_voltage_v,
      });
    }
  }
  if (hasTable(db, "unitech_vehicle_snapshots")) {
    for (const r of db.prepare(`SELECT * FROM unitech_vehicle_snapshots`).all()) {
      const age = Number(r.position_age_days);
      const lastMs = Number.isFinite(trackerTimeMs(r.last_event_at)) ? trackerTimeMs(r.last_event_at) : NaN;
      add(String(r.asset_code || r.registration).toUpperCase(), {
        provider: "Unitech",
        registration: r.registration,
        vehicle_name: r.vehicle_name,
        last_ms: lastMs,
        // Unitech gives the age of the last position directly.
        hours_since: r.position_age_days != null && Number.isFinite(age) ? age * 24 : (now.getTime() - lastMs) / 3600000,
        tracker_battery_pct: null,
        supply_voltage_v: null,
      });
    }
  }
  return [...byVehicle.entries()].map(([code, e]) => {
    const asset = assets.get(code);
    // The vehicle is reporting when any of its units is.
    const freshest = e.units.reduce((a, u) => (Number.isFinite(u.hours_since) && (!Number.isFinite(a) || u.hours_since < a) ? u.hours_since : a), NaN);
    const state = reportingState(freshest);
    const lastMs = Math.max(...e.units.map((u) => (Number.isFinite(u.last_ms) ? u.last_ms : -Infinity)));
    return {
      asset_code: asset ? asset.asset_code : code,
      description: asset
        ? asset.asset_name || ""
        : `${e.units.find((u) => u.vehicle_name)?.vehicle_name || "Unknown"} (not linked to a fleet number)`,
      category: asset?.category || "",
      linked: Boolean(asset),
      registration: [...new Set(e.units.map((u) => u.registration))].join(", "),
      tracking_unit: [...new Set(e.units.map((u) => u.provider))].sort().join(" + "),
      last_report: Number.isFinite(lastMs) ? siteStamp(lastMs) : "—",
      hours_since: Number.isFinite(freshest) ? Number(freshest.toFixed(1)) : null,
      status: state.status,
      status_label: state.label,
      units: e.units,
    };
  }).sort((a, b) => a.asset_code.localeCompare(b.asset_code));
}

/** Active vehicles (by type) that no tracker reports for. */
export function untrackedVehicles(db, tracked) {
  if (!hasTable(db, "assets")) return [];
  const have = new Set(tracked.map((t) => String(t.asset_code).toUpperCase()));
  const archived = db.prepare(`PRAGMA table_info(assets)`).all().some((c) => c.name === "archived") ? "AND COALESCE(archived, 0) = 0" : "";
  return db.prepare(`SELECT asset_code, asset_name, category FROM assets WHERE COALESCE(active, 1) = 1 ${archived} ORDER BY asset_code`).all()
    .filter((a) => VEHICLE_CATEGORY_RE.test(String(a.category || "")) && !have.has(String(a.asset_code).toUpperCase()));
}

function eventsInRange(db, from, to) {
  if (!hasTable(db, "cartrack_events")) return [];
  return db.prepare(`
    SELECT registration, asset_code, event_type, event_type_label, event_time, speed_kmh, speed_limit_kmh,
      latitude, longitude, driver_name, description, is_speeding
    FROM cartrack_events
    WHERE date(event_time) >= date(?) AND date(event_time) <= date(?)
    ORDER BY event_time
  `).all(from, to);
}

const hourOf = (eventTime) => Number(String(eventTime || "").slice(11, 13));
const inNight = (h, start, end) => (start > end ? h >= start || h < end : h >= start && h < end);
const MOVING_RE = /ignition on|trip start|start of trip|moving|driving|departure/i;

/**
 * The whole report for a period. `nightStart`/`nightEnd` are the hours (site
 * time) when vehicles should not move; default 19:00 to 05:00.
 */
export function buildFleetSecurityReport(db, { from, to, now = new Date(), nightStart = 19, nightEnd = 5, notes = "" } = {}) {
  const units = trackedUnits(db, { now });
  const nameOf = new Map();
  for (const u of units) for (const x of u.units) nameOf.set(String(x.registration).toUpperCase(), u.asset_code);
  const vehicleOf = (e) => nameOf.get(String(e.registration || "").toUpperCase()) || e.asset_code || e.registration;
  const unitechRegs = new Set(units.flatMap((u) => u.units.filter((x) => x.provider === "Unitech").map((x) => String(x.registration).toUpperCase())));
  const events = eventsInRange(db, from, to);

  const alerts = events
    .filter((e) => e.event_type !== "ironlog_speed_alert" && SECURITY_EVENT_RE.test(`${e.event_type || ""} ${e.event_type_label || ""}`))
    .map((e) => ({
      time: String(e.event_time).slice(0, 16),
      vehicle: vehicleOf(e),
      event: e.event_type_label || e.event_type,
      driver: e.driver_name || "",
      location: Number.isFinite(Number(e.latitude)) && Number(e.latitude) !== 0 ? `${Number(e.latitude).toFixed(5)}, ${Number(e.longitude).toFixed(5)}` : "",
    }));

  // Night movement, one line per vehicle per day.
  const nightMap = new Map();
  for (const e of events) {
    const h = hourOf(e.event_time);
    if (!Number.isFinite(h) || !inNight(h, nightStart, nightEnd)) continue;
    if (!(Number(e.speed_kmh) > 5) && !MOVING_RE.test(`${e.event_type || ""} ${e.event_type_label || ""}`)) continue;
    const vehicle = vehicleOf(e);
    const key = `${vehicle}|${String(e.event_time).slice(0, 10)}`;
    const n = nightMap.get(key) || { date: String(e.event_time).slice(0, 10), vehicle, first: e.event_time, last: e.event_time, events: 0, max_speed: 0 };
    n.events += 1;
    n.last = e.event_time;
    n.max_speed = Math.max(n.max_speed, Number(e.speed_kmh) || 0);
    nightMap.set(key, n);
  }
  const night = [...nightMap.values()].map((n) => ({ ...n, first: String(n.first).slice(11, 16), last: String(n.last).slice(11, 16) }));

  // Speeding per vehicle.
  const speedMap = new Map();
  for (const e of events.filter((x) => Number(x.is_speeding) === 1)) {
    const vehicle = vehicleOf(e);
    const s = speedMap.get(vehicle) || { vehicle, trackers: new Set(), events: 0, max_speed: 0, limit: e.speed_limit_kmh, last: e.event_time };
    s.trackers.add(unitechRegs.has(String(e.registration || "").toUpperCase()) ? "Unitech" : "Cartrack");
    s.events += 1;
    if ((Number(e.speed_kmh) || 0) >= s.max_speed) { s.max_speed = Number(e.speed_kmh) || 0; s.limit = e.speed_limit_kmh; }
    s.last = e.event_time;
    speedMap.set(vehicle, s);
  }
  const speeding = [...speedMap.values()].sort((a, b) => b.events - a.events)
    .map(({ trackers, ...s }) => ({ ...s, tracker: [...trackers].sort().join(" + "), last: String(s.last).slice(0, 16) }));

  // Per unit: a vehicle with two units still counts when one of them goes quiet.
  const offline = units.flatMap((u) => u.units.map((x) => ({ u, x, state: reportingState(x.hours_since) })))
    .filter(({ state }) => state.status !== "ok")
    .map(({ u, x, state }) => ({ vehicle: u.asset_code, tracking_unit: x.provider, last_report: Number.isFinite(x.last_ms) ? siteStamp(x.last_ms) : "—", issue: state.label }));
  const power = [];
  for (const u of units) {
    for (const x of u.units) {
      const v = Number(x.supply_voltage_v);
      const b = Number(x.tracker_battery_pct);
      if (x.supply_voltage_v != null && Number.isFinite(v) && v < 1) power.push({ vehicle: u.asset_code, issue: "No power to tracker (0 V): unit unplugged or main battery off" });
      else if (x.supply_voltage_v != null && Number.isFinite(v) && v < 11.5) power.push({ vehicle: u.asset_code, issue: `Low vehicle battery at tracker (${v.toFixed(1)} V)` });
      if (x.tracker_battery_pct != null && Number.isFinite(b) && b < 20) power.push({ vehicle: u.asset_code, issue: `Tracker backup battery low (${Math.round(b)}%)` });
    }
  }

  // Every event type seen, so nothing is hidden by the keyword list above.
  const typeMap = new Map();
  for (const e of events) {
    const label = e.event_type_label || e.event_type || "Unknown";
    typeMap.set(label, (typeMap.get(label) || 0) + 1);
  }
  const eventTypes = [...typeMap.entries()].map(([label, count]) => ({ label, count, security: SECURITY_EVENT_RE.test(label) })).sort((a, b) => b.count - a.count);

  const untracked = untrackedVehicles(db, units);
  const count = (s) => units.filter((u) => u.tracking_unit === s).length;
  return {
    from,
    to,
    generated_at: siteStamp(now.getTime()),
    night_hours: `${String(nightStart).padStart(2, "0")}:00 to ${String(nightEnd).padStart(2, "0")}:00`,
    notes: String(notes || "").trim().slice(0, 3000),
    units,
    untracked,
    issues: { alerts, night, speeding, offline, power },
    event_types: eventTypes,
    summary: {
      vehicles_tracked: units.length,
      cartrack: count("Cartrack"),
      unitech: count("Unitech"),
      both: count("Cartrack + Unitech"),
      reporting: units.filter((u) => u.status === "ok").length,
      not_reporting: offline.length,
      untracked: untracked.length,
      tracker_alerts: alerts.length,
      night_movements: night.length,
      speeding_events: speeding.reduce((s, x) => s + x.events, 0),
      speeding_vehicles: speeding.length,
      power_issues: power.length,
      events_checked: events.length,
    },
  };
}

/** Text the PDF's built-in fonts can print (they have no ≥, ≤ or smart quotes). */
export function pdfSafe(value) {
  return String(value ?? "")
    .replace(/≥/g, ">=").replace(/≤/g, "<=").replace(/[–—]/g, "-").replace(/[‘’]/g, "'").replace(/[“”]/g, '"').replace(/→/g, "->")
    .replace(/[^\n\x20-\x7E\u00A0-\u00FF]/g, "");
}

const cleanRows = (rows) => rows.map((r) => Object.fromEntries(Object.entries(r).map(([k, v]) => [k, typeof v === "string" ? pdfSafe(v) : v])));

/** Draws the report into a PDF document (pdfGenerator helpers passed in). */
export function writeFleetSecurityPdf(doc, report, { sectionTitle, table: rawTable, kvGrid }) {
  const table = (d, cols, rows, opts) => rawTable(d, cols, cleanRows(rows), opts);
  const s = report.summary;
  const para = (text, color = "#334155") => {
    doc.font("Helvetica").fontSize(9.5).fillColor(color).text(text, doc.page.margins.left, doc.y, { width: doc.page.width - doc.page.margins.left - doc.page.margins.right });
    doc.moveDown(0.4);
  };
  const none = (text) => para(text, "#64748b");

  sectionTitle(doc, `Fleet tracking security report: ${report.from} to ${report.to}`);
  kvGrid(doc, [
    { k: "Vehicles with a tracker", v: `${s.vehicles_tracked} (Cartrack ${s.cartrack}, Unitech ${s.unitech}, both ${s.both})` },
    { k: "Reporting now", v: `${s.reporting} of ${s.vehicles_tracked}` },
    { k: "Trackers not reporting", v: s.not_reporting },
    { k: "Vehicles with no tracker", v: s.untracked },
    { k: "Tracker security alerts", v: s.tracker_alerts },
    { k: `Night movement (${report.night_hours})`, v: s.night_movements },
    { k: "Speeding", v: `${s.speeding_events} events, ${s.speeding_vehicles} vehicles` },
    { k: "Tracker power problems", v: s.power_issues },
  ], 2);
  para(`Generated ${report.generated_at} (site time). Positions and alerts come from Cartrack and Unitech as synced into IronLog; times are as the trackers report them.`, "#64748b");

  if (report.notes) {
    sectionTitle(doc, "Other incidents and notes");
    para(pdfSafe(report.notes));
  }

  sectionTitle(doc, "1. Tracking unit per vehicle");
  if (report.units.length) {
    table(doc, [
      { key: "asset_code", label: "Fleet no.", width: 0.11 },
      { key: "description", label: "Description", width: 0.22 },
      { key: "registration", label: "Registration", width: 0.15 },
      { key: "tracking_unit", label: "Tracking unit", width: 0.15 },
      { key: "last_report", label: "Last report", width: 0.15 },
      { key: "status_label", label: "Status", width: 0.22 },
    ], report.units, { fontSize: 8 });
  } else none("No tracked vehicles synced yet. Check the Cartrack and Unitech settings under User Admin.");

  if (report.untracked.length) {
    sectionTitle(doc, "Vehicles with no tracker found");
    para("Active vehicles (by type) that neither Cartrack nor Unitech reports for. If a vehicle does have a unit, link its registration under User Admin → Cartrack → vehicle links.", "#64748b");
    table(doc, [
      { key: "asset_code", label: "Fleet no.", width: 0.2 },
      { key: "asset_name", label: "Description", width: 0.5 },
      { key: "category", label: "Type", width: 0.3 },
    ], report.untracked, { fontSize: 8 });
  }

  const issues = report.issues;
  sectionTitle(doc, "2. Tracker security alerts");
  if (issues.alerts.length) {
    table(doc, [
      { key: "time", label: "Time", width: 0.15 },
      { key: "vehicle", label: "Vehicle", width: 0.12 },
      { key: "event", label: "Alert", width: 0.35 },
      { key: "driver", label: "Driver", width: 0.15 },
      { key: "location", label: "Location (lat, long)", width: 0.23 },
    ], issues.alerts, { fontSize: 8 });
  } else none("No tamper, panic, power, towing, zone or similar alerts recorded in this period.");

  sectionTitle(doc, `3. Vehicles moving at night (${report.night_hours})`);
  if (issues.night.length) {
    table(doc, [
      { key: "date", label: "Date", width: 0.16 },
      { key: "vehicle", label: "Vehicle", width: 0.16 },
      { key: "first", label: "First seen", width: 0.16 },
      { key: "last", label: "Last seen", width: 0.16 },
      { key: "events", label: "Events", width: 0.16 },
      { key: "max_speed", label: "Top speed (km/h)", width: 0.2 },
    ], issues.night, { fontSize: 8 });
  } else none("No vehicle movement recorded at night in this period.");

  sectionTitle(doc, "4. Speeding");
  if (issues.speeding.length) {
    table(doc, [
      { key: "vehicle", label: "Vehicle", width: 0.16 },
      { key: "tracker", label: "Tracker", width: 0.14 },
      { key: "events", label: "Events", width: 0.12 },
      { key: "max_speed", label: "Top speed (km/h)", width: 0.18 },
      { key: "limit", label: "Limit (km/h)", width: 0.15 },
      { key: "last", label: "Last time", width: 0.25 },
    ], issues.speeding, { fontSize: 8 });
  } else none("No speeding recorded in this period.");

  sectionTitle(doc, "5. Trackers not reporting and power problems");
  if (issues.offline.length || issues.power.length) {
    const rows = [
      ...issues.offline.map((o) => ({ vehicle: o.vehicle, unit: o.tracking_unit, detail: `${o.issue} (last report ${o.last_report})` })),
      ...issues.power.map((p) => ({ vehicle: p.vehicle, unit: "Cartrack", detail: p.issue })),
    ];
    para("A tracker that stops reporting or loses power can mean a fault, a flat battery, or a unit that was tampered with. Each one should be checked.", "#64748b");
    table(doc, [
      { key: "vehicle", label: "Vehicle", width: 0.16 },
      { key: "unit", label: "Tracking unit", width: 0.16 },
      { key: "detail", label: "Issue", width: 0.68 },
    ], rows, { fontSize: 8 });
  } else none("All trackers are reporting and powered.");

  if (report.event_types.length) {
    sectionTitle(doc, "Appendix: all tracker events in this period");
    table(doc, [
      { key: "label", label: "Event type", width: 0.6 },
      { key: "count", label: "Count", width: 0.2 },
      { key: "flag", label: "Security", width: 0.2 },
    ], report.event_types.map((t) => ({ ...t, flag: t.security ? "Yes" : "" })), { fontSize: 8 });
  }
}
