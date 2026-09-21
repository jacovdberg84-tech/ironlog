// The dashboard owns the single source of truth for daily fleet-KPI math.
// Maintenance registers that calculation here once at startup so reliability
// can use the identical per-asset run and downtime totals without duplicating
// or drifting from the Asset KPI rules.
let assetKpiRangeBuilder = null;

export function registerAssetKpiRangeBuilder(builder) {
  assetKpiRangeBuilder = typeof builder === "function" ? builder : null;
}

export function getAssetKpiRangeBuilder() {
  return assetKpiRangeBuilder;
}
