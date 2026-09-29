// IRONLOG/api/utils/serviceCostSource.js — plain-language names for service cost sources.
const LABELS = {
  service_template: "service template",
  manual_all_in_estimate: "all-in quote",
  manual_parts_and_labor: "manual parts and labour",
  manual_store_pricing: "manual parts at store prices",
  manual_labor: "manual labour",
  historical_average: "past service average",
  historical_asset_service_average: "past service average",
};

/** Where an upcoming service's cost came from, for reports and screens. */
export function serviceCostSourceLabel(row) {
  if (row?.needs_manual_input) return "needs pricing";
  return LABELS[String(row?.forecast?.cost_source || "")] || "needs pricing";
}
