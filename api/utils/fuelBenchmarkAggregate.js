/** Normalize assets.category for fuel benchmark rollups. */
export function normalizeEquipmentCategory(raw) {
  const s = String(raw || "").trim();
  return s || "Uncategorized";
}

/**
 * Roll up per-asset fuel benchmark rows into equipment category totals for a date range.
 * @param {Array<object>} assetRows
 * @param {number} tolerance
 */
export function aggregateFuelBenchmarkByCategory(assetRows, tolerance = 0.15) {
  const tol = Number.isFinite(Number(tolerance)) ? Math.max(0, Number(tolerance)) : 0.15;
  const groups = new Map();

  for (const r of assetRows || []) {
    const category = normalizeEquipmentCategory(r.category);
    const mode = String(r.metric_mode || "hours").toLowerCase() === "km" ? "km" : "hours";
    const key = `${category}\0${mode}`;
    const g = groups.get(key) || {
      category,
      metric_mode: mode,
      asset_count: 0,
      fuel_liters: 0,
      hours_run: 0,
      km_run: 0,
      fill_count: 0,
      excessive_asset_count: 0,
      oem_lph_weighted: 0,
      oem_km_per_l_weighted: 0,
      hours_weight: 0,
      fuel_weight_km: 0,
      asset_codes: [],
    };

    g.asset_count += 1;
    g.fuel_liters += Number(r.fuel_liters || 0);
    g.hours_run += Number(r.hours_run || 0);
    g.km_run += Number(r.km_run || 0);
    g.fill_count += Number(r.fill_count || 0);
    if (r.is_excessive || r.flag === "EXCESSIVE") g.excessive_asset_count += 1;

    if (mode === "hours") {
      const hrs = Number(r.hours_run || 0);
      if (hrs > 0) {
        g.hours_weight += hrs;
        g.oem_lph_weighted += hrs * Number(r.oem_lph || 5);
      }
    } else {
      const fuel = Number(r.fuel_liters || 0);
      if (fuel > 0) {
        g.fuel_weight_km += fuel;
        g.oem_km_per_l_weighted += fuel * Number(r.oem_km_per_l || 2);
      }
    }
    if (r.asset_code) g.asset_codes.push(String(r.asset_code));
    groups.set(key, g);
  }

  return Array.from(groups.values())
    .map((g) => {
      const fuel = Number(g.fuel_liters.toFixed(2));
      const hours = Number(g.hours_run.toFixed(2));
      const km = Number(g.km_run.toFixed(2));
      const mode = g.metric_mode;
      let actual_lph = null;
      let oem_lph = null;
      let actual_km_per_l = null;
      let oem_km_per_l = null;
      let is_excessive = false;

      if (mode === "km") {
        actual_km_per_l = fuel > 0 && km > 0 ? Number((km / fuel).toFixed(3)) : null;
        oem_km_per_l = g.fuel_weight_km > 0
          ? Number((g.oem_km_per_l_weighted / g.fuel_weight_km).toFixed(3))
          : null;
        const lowThreshold = oem_km_per_l != null ? oem_km_per_l * Math.max(0, 1 - tol) : null;
        is_excessive = g.fill_count >= 2
          && actual_km_per_l != null
          && lowThreshold != null
          && actual_km_per_l < lowThreshold;
      } else {
        actual_lph = hours > 0 ? Number((fuel / hours).toFixed(3)) : null;
        oem_lph = g.hours_weight > 0
          ? Number((g.oem_lph_weighted / g.hours_weight).toFixed(3))
          : null;
        const excessiveThreshold = oem_lph != null ? oem_lph * (1 + tol) : null;
        is_excessive = g.fill_count >= 2
          && actual_lph != null
          && excessiveThreshold != null
          && actual_lph > excessiveThreshold;
      }

      return {
        category: g.category,
        metric_mode: mode,
        asset_count: g.asset_count,
        asset_codes: g.asset_codes.sort((a, b) => a.localeCompare(b)),
        fuel_liters: fuel,
        hours_run: hours,
        km_run: km,
        fill_count: g.fill_count,
        actual_lph,
        oem_lph,
        actual_km_per_l,
        oem_km_per_l,
        excessive_asset_count: g.excessive_asset_count,
        is_excessive,
        flag: is_excessive ? "EXCESSIVE" : "OK",
        variance_lph: actual_lph != null && oem_lph != null ? Number((actual_lph - oem_lph).toFixed(3)) : null,
        variance_km_per_l: actual_km_per_l != null && oem_km_per_l != null
          ? Number((actual_km_per_l - oem_km_per_l).toFixed(3))
          : null,
      };
    })
    .sort((a, b) => {
      const ex = Number(Boolean(b.is_excessive)) - Number(Boolean(a.is_excessive));
      if (ex !== 0) return ex;
      return Number(b.fuel_liters || 0) - Number(a.fuel_liters || 0);
    });
}
