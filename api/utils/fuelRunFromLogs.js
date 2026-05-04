/**
 * Derive km_run and hours_run from ordered fuel log rows.
 * FAMS exports often omit Measurement; meter_unit is then blank while KMHour
 * still holds the cumulative reading — infer km vs hours from the asset mode.
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
    const openMeter = row.open_meter_value == null ? null : Number(row.open_meter_value);
    const closeMeter = row.close_meter_value == null ? null : Number(row.close_meter_value);

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
