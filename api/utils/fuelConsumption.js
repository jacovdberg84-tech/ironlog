// IRONLOG/api/utils/fuelConsumption.js — fuel consumption per machine (L/hr, km/L).
//
// Fill-to-fill: the litres put in at a fill are what the machine used since the
// previous fill, so each fill's litres are matched with the meter movement
// since the last good meter reading. A reading that cannot be right (meter went
// backwards, or moved more than the machine can run in the days between fills)
// is rejected: its litres are carried to the next good reading instead of
// inflating or deflating the period. Fills without a reading are carried the
// same way. When too little of the fuel can be matched to meter readings, the
// daily hours captured on IronLog are used instead (hours machines only).
//
// A machine's OEM benchmark counts as "not set" while it still holds the column
// default (5 L/hr, 2 km/L) and nobody has saved it; such machines are shown but
// never flagged EXCESSIVE.

import { normalizeEquipmentCategory } from "./fuelBenchmarkAggregate.js";
import { isOperationalHireAsset } from "./hiredEquipment.js";

export const MAX_RUN_PER_DAY = { hours: 24, km: 1500 };
export const MIN_COVERAGE = 0.5;
const DEFAULT_LPH = 5;
const DEFAULT_KMPL = 2;

const dayNo = (d) => Math.floor(Date.parse(`${String(d).slice(0, 10)}T00:00:00Z`) / 86400000);

function hasColumn(db, table, col) {
  return db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
}

export function ensureFuelBenchmarkSchema(db) {
  if (!hasColumn(db, "assets", "fuel_benchmark_set_at")) {
    db.prepare(`ALTER TABLE assets ADD COLUMN fuel_benchmark_set_at TEXT`).run();
  }
}

/** The cumulative meter reading on a fill in the machine's unit, or null. */
export function readingOf(row, mode) {
  const unit = String(row.meter_unit || "").trim().toLowerCase() || mode;
  if (unit !== mode) return null;
  const close = Number(row.close_meter_value);
  if (Number.isFinite(close) && close > 0) return close;
  const meter = Number(row.meter_run_value);
  return Number.isFinite(meter) && meter > 0 ? meter : null;
}

/**
 * fills: fills in the period, oldest first ({ id, log_date, liters, meter_unit,
 * meter_run_value, close_meter_value, hours_run }). prev: the last fill before
 * the period that has a reading in this unit (or null).
 */
export function fillToFill(fills, prev, mode) {
  const unitMode = mode === "km" ? "km" : "hours";
  const perDay = MAX_RUN_PER_DAY[unitMode];
  const list = (fills || []).map((f) => ({ ...f, reading: readingOf(f, unitMode), liters: Number(f.liters || 0) }));
  const fits = (fromReading, fromDate, toReading, toDate) => {
    const delta = toReading - fromReading;
    const allowed = perDay * (Math.max(0, dayNo(toDate) - dayNo(fromDate)) + 1);
    return delta > 0 && delta <= allowed;
  };

  let ref = prev ? readingOf(prev, unitMode) : null;
  let refDate = prev?.log_date || null;
  let pending = 0;
  let run = 0;
  let matched = 0;
  let unmatched = 0;
  let matchedFills = 0;
  let otherUnit = 0;
  const suspect = [];
  // One entry per fill: what happened to its reading and litres.
  const trace = list.map((f) => ({ id: f.id ?? null, date: f.log_date, liters: f.liters, reading: f.reading, run: null, interval_liters: null, status: "" }));
  let waiting = []; // trace entries whose litres are carried to the next good reading
  const carry = (i, status) => { trace[i].status = status; waiting.push(i); };
  // Carried litres are settled: "… carried to next reading" becomes where they went.
  const settle = (where) => {
    for (const w of waiting) trace[w].status = trace[w].status.replace(/carried to next reading$/, where);
    waiting = [];
  };

  for (let i = 0; i < list.length; i += 1) {
    const f = list[i];
    if (f.reading == null) {
      // Old rows with only the hours since the last fill.
      const legacy = unitMode === "hours" && !f.meter_unit && Number(f.hours_run) > 0 && !(Number(f.meter_run_value) > 0) ? Number(f.hours_run) : 0;
      if (legacy > 0 && legacy <= perDay * 31) {
        run += legacy;
        matched += pending + f.liters;
        matchedFills += 1;
        Object.assign(trace[i], { run: legacy, interval_liters: pending + f.liters, status: "ok" });
        settle(`carried to the ${f.log_date} fill`);
        pending = 0;
        continue;
      }
      const other = String(f.meter_unit || "").trim() && String(f.meter_unit).trim().toLowerCase() !== unitMode && Number(f.meter_run_value) > 0;
      if (other) otherUnit += 1;
      pending += f.liters;
      carry(i, other ? `reading in ${unitMode === "km" ? "hours" : "km"}: carried to next reading` : "no reading: carried to next reading");
      continue;
    }
    if (ref == null) {
      // First reading we know of: the fuel before it covers running we cannot see.
      unmatched += pending + f.liters;
      settle("not counted (no earlier reading)");
      trace[i].status = "first reading: not counted";
      pending = 0;
      ref = f.reading;
      refDate = f.log_date;
      continue;
    }
    if (f.reading === ref) {
      pending += f.liters; // topped up without running (or the meter was not read)
      carry(i, "same reading as last fill: carried to next reading");
      continue;
    }
    if (fits(ref, refDate, f.reading, f.log_date)) {
      run += f.reading - ref;
      matched += pending + f.liters;
      matchedFills += 1;
      Object.assign(trace[i], { run: f.reading - ref, interval_liters: pending + f.liters, status: "ok" });
      settle(`carried to the ${f.log_date} fill`);
      pending = 0;
      ref = f.reading;
      refDate = f.log_date;
      continue;
    }
    // Not possible from the last good reading. If the next reading follows on
    // from this one (and not from the old one), the meter was changed or the
    // old reading was wrong: start again from here. Otherwise this one is wrong.
    const next = list.slice(i + 1).find((x) => x.reading != null && x.reading !== f.reading);
    const restart = next && fits(f.reading, f.log_date, next.reading, next.log_date) && !fits(ref, refDate, next.reading, next.log_date);
    suspect.push({
      id: f.id ?? null,
      date: f.log_date,
      reading: f.reading,
      previous: ref,
      previous_date: refDate,
      reason: f.reading < ref ? "meter went back" : "more than the machine can run in that time",
      action: restart ? "counted from here" : "ignored",
    });
    const why = f.reading < ref ? "meter went back" : `jump of ${Number((f.reading - ref).toFixed(1))} ${unitMode === "km" ? "km" : "h"}`;
    if (restart) {
      unmatched += pending + f.liters;
      settle("not counted (meter restarted)");
      trace[i].status = `${why}: counted from this reading on`;
      pending = 0;
      ref = f.reading;
      refDate = f.log_date;
    } else {
      pending += f.liters;
      carry(i, `${why}: reading ignored, litres carried to next reading`);
    }
  }
  unmatched += pending;
  settle("not counted (no later reading in period)");
  const total = list.reduce((s, f) => s + f.liters, 0);
  return {
    run: Number(run.toFixed(2)),
    matched_liters: Number(matched.toFixed(2)),
    unmatched_liters: Number(unmatched.toFixed(2)),
    total_liters: Number(total.toFixed(2)),
    coverage: total > 0 ? matched / total : 0,
    matched_fills: matchedFills,
    other_unit_fills: otherUnit,
    suspect,
    trace,
  };
}

export function benchmarkIsSet(value, defaultValue, setAt) {
  const v = Number(value);
  if (!Number.isFinite(v) || v <= 0) return false;
  return Math.abs(v - defaultValue) > 1e-9 || Boolean(setAt);
}

/**
 * Consumption for one machine over a period. asset: a row of
 * fuelBenchmarkAssetsInRangeSql(). Returns the row every fuel screen and
 * report uses.
 */
export function assetConsumption(db, asset, start, end, tolerance = 0.15, { withTrace = false } = {}) {
  const mode = String(asset.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
  const fills = db.prepare(`
    SELECT id, log_date, COALESCE(liters, 0) AS liters, LOWER(COALESCE(meter_unit, '')) AS meter_unit,
      meter_run_value, close_meter_value, COALESCE(hours_run, 0) AS hours_run, source
    FROM fuel_logs WHERE asset_id = ? AND log_date BETWEEN ? AND ?
    ORDER BY log_date, id
  `).all(asset.asset_id, start, end);
  const before = db.prepare(`
    SELECT id, log_date, LOWER(COALESCE(meter_unit, '')) AS meter_unit, meter_run_value, close_meter_value
    FROM fuel_logs WHERE asset_id = ? AND log_date < ?
      AND (COALESCE(meter_run_value, 0) > 0 OR COALESCE(close_meter_value, 0) > 0)
    ORDER BY log_date DESC, id DESC LIMIT 10
  `).all(asset.asset_id, start);
  const prev = before.find((r) => readingOf(r, mode) != null) || null;
  const f2f = fillToFill(fills, prev, mode);

  const skipDaily = Number(asset.archived || 0) === 1 && isOperationalHireAsset(asset);
  const dailyHours = mode === "hours" && !skipDaily
    ? Number(db.prepare(`
        SELECT COALESCE(SUM(hours_run), 0) AS v FROM daily_hours
        WHERE asset_id = ? AND work_date BETWEEN ? AND ? AND COALESCE(is_used, 1) = 1
          AND LOWER(COALESCE(NULLIF(TRIM(input_unit), ''), 'hours')) <> 'km' AND COALESCE(hours_run, 0) > 0
      `).get(asset.asset_id, start, end)?.v || 0)
    : 0;

  const fuel = Number(asset.fuel_liters || f2f.total_liters || 0);
  let source = "none";
  let run = 0;
  let basisLiters = 0;
  if (f2f.run > 0 && f2f.coverage >= MIN_COVERAGE) {
    source = "fill_meter";
    run = f2f.run;
    basisLiters = f2f.matched_liters;
  } else if (mode === "hours" && dailyHours > 0) {
    source = "daily_hours";
    run = dailyHours;
    basisLiters = fuel;
  }

  const lphSet = benchmarkIsSet(asset.oem_lph, DEFAULT_LPH, asset.fuel_benchmark_set_at);
  const kmplSet = benchmarkIsSet(asset.oem_kmpl, DEFAULT_KMPL, asset.fuel_benchmark_set_at);
  const oem = lphSet ? Number(asset.oem_lph) : null;
  const oemK = kmplSet ? Number(asset.oem_kmpl) : null;
  const hours = mode === "hours" ? run : 0;
  const km = mode === "km" ? run : 0;
  const lph = hours > 0 && basisLiters > 0 ? basisLiters / hours : null;
  const kmpl = km > 0 && basisLiters > 0 ? km / basisLiters : null;
  const threshold = oem != null ? oem * (1 + tolerance) : null;
  const lowKmpl = oemK != null ? oemK * Math.max(0, 1 - tolerance) : null;
  const fillCount = Number(asset.fill_count || fills.length || 0);
  const enough = fillCount >= 2;
  const isExcessive = enough && (mode === "km"
    ? kmpl != null && lowKmpl != null && kmpl < lowKmpl
    : lph != null && threshold != null && lph > threshold);
  const r3 = (n) => (n == null ? null : Number(Number(n).toFixed(3)));

  return {
    asset_id: Number(asset.asset_id),
    asset_code: asset.asset_code,
    asset_name: asset.asset_name,
    category: normalizeEquipmentCategory(asset.category),
    metric_mode: mode,
    is_hired: isOperationalHireAsset(asset) || Boolean(String(asset.hire_billing_mode || "").trim()),
    archived: Number(asset.archived || 0),
    run_source: source,
    fuel_liters: Number(fuel.toFixed(2)),
    basis_liters: Number(basisLiters.toFixed(2)),
    matched_liters: f2f.matched_liters,
    coverage_pct: fuel > 0 ? Number(((f2f.matched_liters / fuel) * 100).toFixed(0)) : 0,
    daily_hours: Number(dailyHours.toFixed(2)),
    km_run: Number(km.toFixed(2)),
    hours_run: Number(hours.toFixed(2)),
    actual_lph: r3(lph),
    oem_lph: r3(oem),
    oem_set: mode === "km" ? kmplSet : lphSet,
    excessive_threshold_lph: r3(threshold),
    threshold_lph: r3(threshold),
    variance_lph: lph != null && oem != null ? r3(lph - oem) : null,
    actual_km_per_l: r3(kmpl),
    oem_km_per_l: r3(oemK),
    low_threshold_km_per_l: r3(lowKmpl),
    threshold_km_per_l: r3(lowKmpl),
    variance_km_per_l: kmpl != null && oemK != null ? r3(kmpl - oemK) : null,
    fill_count: fillCount,
    has_enough_samples: enough,
    is_excessive: Boolean(isExcessive),
    flag: isExcessive ? "EXCESSIVE" : "OK",
    suspect_readings: f2f.suspect.length,
    suspect: f2f.suspect.slice(0, 20),
    other_unit_fills: f2f.other_unit_fills,
    ...(withTrace ? { trace: f2f.trace.map((t, i) => ({ ...t, source: fills[i]?.source || "" })) } : {}),
  };
}

/** Sort: EXCESSIVE first, then by how far over the benchmark. */
export function sortBenchmarkRows(rows) {
  const over = (r) => (r.metric_mode === "km" ? -(r.variance_km_per_l ?? 999) : r.variance_lph ?? -999);
  return rows.sort((a, b) => (Number(b.is_excessive) - Number(a.is_excessive)) || (over(b) - over(a)));
}

/** Short plain-language note for a row (reports and the dashboard). */
export function consumptionNote(r) {
  const bits = [];
  if (r.run_source === "daily_hours") bits.push("from daily hours (fuel meter readings incomplete)");
  else if (r.run_source === "fill_meter" && r.coverage_pct < 90) bits.push(`${r.coverage_pct}% of fuel matched to meter readings`);
  else if (r.run_source === "none") bits.push(r.metric_mode === "km" ? "no usable odometer readings" : "no usable hours");
  if (r.suspect_readings) bits.push(`${r.suspect_readings} meter reading${r.suspect_readings === 1 ? "" : "s"} rejected`);
  if (r.other_unit_fills) bits.push(`${r.other_unit_fills} fill${r.other_unit_fills === 1 ? "" : "s"} recorded in ${r.metric_mode === "km" ? "hours" : "km"}`);
  if (!r.oem_set) bits.push("benchmark not set");
  return bits.join("; ");
}
