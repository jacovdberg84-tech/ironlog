/**
 * Resolve whether an asset should use km/L vs L/hr fuel benchmarking.
 * Keep SQL (sqlFuelMetricModeExpr) and JS (inferFuelMetricMode) aligned.
 */
import { BMH_HIRE_ASSET_CODES, HIRE_KM_CODE_PREFIXES } from "./hiredEquipment.js";

const HOUR_HIRE_CODES = new Set(BMH_HIRE_ASSET_CODES);

export function inferFuelMetricMode({ category, asset_code, asset_name, utilization_mode } = {}) {
  const mode = String(utilization_mode || "").trim().toLowerCase();
  if (mode === "km" || mode === "hours") return mode;

  const cat = String(category || "").toLowerCase();
  const code = String(asset_code || "").toUpperCase();
  const name = String(asset_name || "").toLowerCase();

  if (HOUR_HIRE_CODES.has(code)) return "hours";
  if (HIRE_KM_CODE_PREFIXES.some((p) => code.startsWith(p))) return "km";
  if (/^V\d{2}AM$/.test(code) || /^T\d{2}AM$/.test(code)) return "km";

  const kmHints = ["truck", "vehicle", "ldv", "pickup", "bakkie", "tipper", "dump", "haul", "spinner"];
  if (kmHints.some((k) => cat.includes(k) || name.includes(k))) return "km";
  if (name.includes("toyota") && name.includes("hilux")) return "km";
  if (code.includes("HILUX")) return "km";

  return "hours";
}

/** SQL CASE expression for fuel benchmark metric mode. */
export function sqlFuelMetricModeExpr(a = "a") {
  const hourCodesSql = BMH_HIRE_ASSET_CODES.map((c) => `'${c}'`).join(", ");
  return `
    CASE
      WHEN UPPER(COALESCE(${a}.asset_code, '')) IN (${hourCodesSql}) THEN 'hours'
      WHEN UPPER(COALESCE(${a}.asset_code, '')) LIKE 'BMP%' THEN 'km'
      WHEN UPPER(COALESCE(${a}.asset_code, '')) GLOB 'V[0-9][0-9]AM' THEN 'km'
      WHEN UPPER(COALESCE(${a}.asset_code, '')) GLOB 'T[0-9][0-9]AM' THEN 'km'
      WHEN UPPER(COALESCE(${a}.asset_code, '')) LIKE 'PTT%' THEN 'km'
      WHEN UPPER(COALESCE(${a}.asset_code, '')) LIKE 'LDV%' THEN 'km'
      ELSE COALESCE(NULLIF(TRIM(${a}.utilization_mode), ''), CASE
        WHEN LOWER(COALESCE(${a}.category, '')) LIKE '%truck%'
          OR LOWER(COALESCE(${a}.category, '')) LIKE '%vehicle%'
          OR LOWER(COALESCE(${a}.category, '')) LIKE '%ldv%'
          OR LOWER(COALESCE(${a}.category, '')) LIKE '%pickup%'
          OR LOWER(COALESCE(${a}.category, '')) LIKE '%bakkie%'
          OR LOWER(COALESCE(${a}.category, '')) LIKE '%tipper%'
          OR LOWER(COALESCE(${a}.category, '')) LIKE '%dump%'
          OR LOWER(COALESCE(${a}.category, '')) LIKE '%haul%'
          OR LOWER(COALESCE(${a}.category, '')) LIKE '%spinner%'
          OR LOWER(COALESCE(${a}.asset_code, '')) LIKE 'ldv%'
          OR UPPER(COALESCE(${a}.asset_code, '')) GLOB 'V[0-9][0-9]AM'
          OR UPPER(COALESCE(${a}.asset_code, '')) GLOB 'T[0-9][0-9]AM'
          OR UPPER(COALESCE(${a}.asset_code, '')) LIKE 'PTT%'
          OR UPPER(COALESCE(${a}.asset_code, '')) LIKE 'BMP%'
          OR LOWER(COALESCE(${a}.asset_name, '')) LIKE '%ldv%'
          OR LOWER(COALESCE(${a}.asset_name, '')) LIKE '%truck%'
          OR LOWER(COALESCE(${a}.asset_name, '')) LIKE '%tipper%'
          OR LOWER(COALESCE(${a}.asset_name, '')) LIKE '%dump%'
          OR LOWER(COALESCE(${a}.asset_name, '')) LIKE '%haul%'
          OR LOWER(COALESCE(${a}.asset_name, '')) LIKE '%spinner%'
          OR (
            (INSTR(LOWER(COALESCE(${a}.asset_name, '')), 'toyota') > 0 AND INSTR(LOWER(COALESCE(${a}.asset_name, '')), 'hilux') > 0)
            OR INSTR(LOWER(COALESCE(${a}.asset_code, '')), 'hilux') > 0
          )
          THEN 'km'
        ELSE 'hours'
      END)
    END
  `;
}

import { sqlIncludeArchivedHireAssets } from "./hiredEquipment.js";

/**
 * Assets with fuel logs in a date range (includes contractor / hired units with fuel,
 * not only active fleet cards).
 */
export function fuelBenchmarkAssetsInRangeSql() {
  const metric = sqlFuelMetricModeExpr("a");
  return `
    SELECT
      a.id AS asset_id,
      a.asset_code,
      a.asset_name,
      a.category,
      COALESCE(a.archived, 0) AS archived,
      COALESCE(a.active, 1) AS active,
      NULLIF(TRIM(COALESCE(a.hire_billing_mode, '')), '') AS hire_billing_mode,
      ${metric} AS metric_mode,
      COALESCE(NULLIF(a.km_per_hour_factor, 0), 10.0) AS km_per_hour_factor,
      COALESCE(a.baseline_fuel_l_per_hour, 5.0) AS oem_lph,
      COALESCE(a.baseline_fuel_km_per_l, 2.0) AS oem_kmpl,
      fl.fuel_liters AS fuel_liters,
      fl.fill_count AS fill_count
    FROM assets a
    INNER JOIN (
      SELECT asset_id, SUM(liters) AS fuel_liters, COUNT(id) AS fill_count
      FROM fuel_logs
      WHERE log_date BETWEEN ? AND ?
      GROUP BY asset_id
    ) fl ON fl.asset_id = a.id
    WHERE ${sqlIncludeArchivedHireAssets("a")}
    ORDER BY a.asset_code ASC
  `;
}
