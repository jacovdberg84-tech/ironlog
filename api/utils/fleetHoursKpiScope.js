/**
 * Daily production rows that contribute to hour-based fleet KPIs (planned / run / availability).
 *
 * Excludes:
 * - Daily standby (`daily_hours.is_used = 0`)
 * - Master standby assets (`assets.is_standby`)
 * - KM-based equipment: `daily_hours.input_unit` = km, `assets.utilization_mode` = km, or Hilux heuristic (matches dashboard JS).
 */
export function andDailyHoursFleetHoursOnly(dh = "dh", a = "a") {
  return `
    AND ${dh}.is_used = 1
    AND ${a}.active = 1
    AND ${a}.is_standby = 0
    AND NOT (
      LOWER(TRIM(COALESCE(${dh}.input_unit, ''))) = 'km'
      OR LOWER(TRIM(COALESCE(${a}.utilization_mode, ''))) = 'km'
      OR (
        (INSTR(LOWER(COALESCE(${a}.asset_name, '')), 'toyota') > 0 AND INSTR(LOWER(COALESCE(${a}.asset_name, '')), 'hilux') > 0)
        OR INSTR(LOWER(COALESCE(${a}.asset_code, '')), 'hilux') > 0
      )
    )
  `;
}

/** Same asset filter without daily_hours (e.g. breakdowns joined only to assets). */
export function andAssetFleetHoursOnly(a = "a") {
  return `
    AND ${a}.active = 1
    AND ${a}.is_standby = 0
    AND NOT (
      LOWER(TRIM(COALESCE(${a}.utilization_mode, ''))) = 'km'
      OR (
        (INSTR(LOWER(COALESCE(${a}.asset_name, '')), 'toyota') > 0 AND INSTR(LOWER(COALESCE(${a}.asset_name, '')), 'hilux') > 0)
        OR INSTR(LOWER(COALESCE(${a}.asset_code, '')), 'hilux') > 0
      )
    )
  `;
}
