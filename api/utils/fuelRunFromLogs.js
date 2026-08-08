/**
 * Derive km_run and hours_run from ordered fuel log rows.
 * FAMS exports often omit Measurement; meter_unit is then blank while KMHour
 * still holds the cumulative reading — infer km vs hours from the asset mode.
 *
 * open_meter_value of 0 is treated as unset (common import default). Using
 * open=0 with close=full cumulative meter was inflating hours and collapsing
 * fleet Avg L/hr (e.g. ~0.16 L/hr).
 */
export function getRunFromFuelRows(logs, prevRow, assetMetricMode) {
  const mode = String(assetMetricMode || "hours").toLowerCase() === "km" ? "km" : "hours";
  if (!logs || !logs.length) return { km_run: 0, hours_run: 0 };

  let prevKmMeter = null;
  let prevHoursMeter = null;
  if (prevRow) {
    let prevUnit = String(prevRow.meter_unit || "").toLowerCase();
    const prevClose = Number(prevRow.close_meter_value || 0);
    const prevMeter = prevClose > 0 ? prevClose : Number(prevRow.meter_run_value || 0);
    if (!prevUnit && prevMeter > 0) prevUnit = mode;
    if (prevUnit === "km" && prevMeter > 0) prevKmMeter = prevMeter;
    if (prevUnit === "hours" && prevMeter > 0) prevHoursMeter = prevMeter;
  }

  let km_run = 0;
  let hours_run = 0;

  for (const row of logs) {
    let unit = String(row.meter_unit || "").toLowerCase();
    const meter = Number(row.meter_run_value || 0);
    const legacyHours = Number(row.hours_run || 0);
    const openRaw = row.open_meter_value == null ? null : Number(row.open_meter_value);
    const closeRaw = row.close_meter_value == null ? null : Number(row.close_meter_value);
    // 0 means unset for open; require a real open reading before trusting open/close delta
    const openMeter = openRaw != null && Number.isFinite(openRaw) && openRaw > 0 ? openRaw : null;
    const closeMeter = closeRaw != null && Number.isFinite(closeRaw) && closeRaw > 0 ? closeRaw : null;

    if (!unit && meter > 0) unit = mode;

    if (openMeter != null && closeMeter != null && closeMeter > openMeter) {
      const delta = closeMeter - openMeter;
      if (unit === "km") km_run += delta;
      else hours_run += delta;
      continue;
    }

    if (unit === "km" && meter > 0) {
      if (prevKmMeter != null) {
        const delta = meter - prevKmMeter;
        if (Number.isFinite(delta) && delta > 0) km_run += delta;
      }
      prevKmMeter = meter;
      continue;
    }

    if (unit === "hours" && meter > 0) {
      if (prevHoursMeter != null) {
        const delta = meter - prevHoursMeter;
        if (Number.isFinite(delta) && delta > 0) hours_run += delta;
      }
      prevHoursMeter = meter;
      continue;
    }

    if (legacyHours > 0) hours_run += legacyHours;
  }

  return { km_run, hours_run };
}

/** Inclusive calendar days between YYYY-MM-DD dates. */
export function daysInclusive(start, end) {
  const a = Date.parse(`${String(start).slice(0, 10)}T00:00:00Z`);
  const b = Date.parse(`${String(end).slice(0, 10)}T00:00:00Z`);
  if (!Number.isFinite(a) || !Number.isFinite(b) || b < a) return 0;
  return Math.floor((b - a) / 86400000) + 1;
}

/**
 * Choose hours for L/hr benchmarking. Prefer fuel-derived hours when plausible;
 * otherwise fall back to daily_hours (avoids cumulative-meter blowups).
 */
export function resolveBenchmarkHoursRun(fuelHours, dailyHours, periodStart, periodEnd) {
  const fuel = Number(fuelHours || 0);
  const daily = Number(dailyHours || 0);
  if (fuel <= 0) return daily > 0 ? daily : 0;
  if (daily <= 0) return fuel;
  const days = daysInclusive(periodStart, periodEnd);
  const maxPlausible = Math.max(24, days * 24);
  if (fuel > maxPlausible) return daily;
  return fuel;
}

/**
 * Fleet summary averages: Avg L/hr must use hours-mode fuel/hours only;
 * Avg km/L must use km-mode fuel/km only.
 */
export function summarizeFuelBenchmarkRows(rows) {
  const summary = (Array.isArray(rows) ? rows : []).reduce(
    (acc, r) => {
      const mode = String(r.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
      const fuel = Number(r.fuel_liters || 0);
      const hours = Number(r.hours_run || 0);
      const km = Number(r.km_run || 0);
      acc.assets += 1;
      acc.fuel_liters += fuel;
      if (mode === "km") {
        acc.km_assets += 1;
        acc.km_run += km;
        if (km > 0) acc.km_fuel += fuel;
      } else {
        acc.hours_assets += 1;
        acc.hours_run += hours;
        if (hours > 0) acc.hours_fuel += fuel;
      }
      if (r.is_excessive || r.flag === "EXCESSIVE") acc.excessive += 1;
      return acc;
    },
    {
      assets: 0,
      fuel_liters: 0,
      hours_run: 0,
      km_run: 0,
      excessive: 0,
      hours_fuel: 0,
      km_fuel: 0,
      hours_assets: 0,
      km_assets: 0,
    }
  );
  summary.fuel_liters = Number(summary.fuel_liters.toFixed(2));
  summary.hours_run = Number(summary.hours_run.toFixed(2));
  summary.km_run = Number(summary.km_run.toFixed(2));
  summary.avg_lph = summary.hours_run > 0
    ? Number((summary.hours_fuel / summary.hours_run).toFixed(3))
    : null;
  summary.avg_km_per_l = summary.km_fuel > 0
    ? Number((summary.km_run / summary.km_fuel).toFixed(3))
    : null;
  return summary;
}
