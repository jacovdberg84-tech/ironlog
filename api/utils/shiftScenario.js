function finite(value, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function rounded(value, places = 2) {
  const factor = 10 ** places;
  return Math.round((finite(value) + Number.EPSILON) * factor) / factor;
}

function scenarioRow(source, baseHours, scenarioHours) {
  const activeDays = finite(source.active_days);
  const runHours = finite(source.run_hours);
  const fuelLiters = finite(source.fuel_liters);
  const fuelCost = finite(source.fuel_cost);
  const nonFuelCost = finite(source.nonfuel_cost);

  if (String(source.metric_mode || "hours").toLowerCase() === "km") {
    return { excluded: true, reason: "Distance-based asset (km)" };
  }
  if (runHours <= 0) return { excluded: true, reason: "No recorded operating hours" };
  if (activeDays <= 0) return { excluded: true, reason: "No active operating days" };
  if (fuelLiters <= 0) return { excluded: true, reason: "No fuel logged in selected period" };

  const litersPerHour = fuelLiters / runHours;
  const fuelCostPerHour = fuelCost / runHours;
  const nonFuelCostPerShift = nonFuelCost / activeDays;
  const baseFuelLitersPerShift = litersPerHour * baseHours;
  const scenarioFuelLitersPerShift = litersPerHour * scenarioHours;
  const baseFuelCostPerShift = fuelCostPerHour * baseHours;
  const scenarioFuelCostPerShift = fuelCostPerHour * scenarioHours;
  const baseTotalCostPerShift = baseFuelCostPerShift + nonFuelCostPerShift;
  const scenarioTotalCostPerShift = scenarioFuelCostPerShift + nonFuelCostPerShift;

  return {
    excluded: false,
    row: {
      asset_id: finite(source.asset_id),
      asset_code: String(source.asset_code || "-"),
      asset_name: String(source.asset_name || ""),
      category: String(source.category || ""),
      active_days: activeDays,
      logged_run_hours: runHours,
      fuel_liters: fuelLiters,
      fuel_cost: fuelCost,
      nonfuel_cost: nonFuelCost,
      actual_liters_per_hour: litersPerHour,
      fuel_cost_per_hour: fuelCostPerHour,
      nonfuel_cost_per_shift: nonFuelCostPerShift,
      base_fuel_liters_per_shift: baseFuelLitersPerShift,
      scenario_fuel_liters_per_shift: scenarioFuelLitersPerShift,
      fuel_liters_saved_per_shift: baseFuelLitersPerShift - scenarioFuelLitersPerShift,
      base_fuel_cost_per_shift: baseFuelCostPerShift,
      scenario_fuel_cost_per_shift: scenarioFuelCostPerShift,
      base_total_cost_per_shift: baseTotalCostPerShift,
      scenario_total_cost_per_shift: scenarioTotalCostPerShift,
      cost_saved_per_shift: baseTotalCostPerShift - scenarioTotalCostPerShift,
      base_cost_per_operating_hour: baseTotalCostPerShift / baseHours,
      scenario_cost_per_operating_hour: scenarioTotalCostPerShift / scenarioHours,
      cost_per_operating_hour_change: (scenarioTotalCostPerShift / scenarioHours) - (baseTotalCostPerShift / baseHours),
      period_fuel_liters_saved: (baseFuelLitersPerShift - scenarioFuelLitersPerShift) * activeDays,
      period_cost_saved: (baseTotalCostPerShift - scenarioTotalCostPerShift) * activeDays,
    },
  };
}

/**
 * Models an alternative operating-shift length using recorded fuel and cost data.
 * Fuel scales with operating hours. Recorded non-fuel cost is held constant per
 * active shift, so managers can see the cost-per-hour impact of a shorter shift.
 */
export function buildShiftScenario(rows, { baseHours = 11, scenarioHours = 8 } = {}) {
  const base = Math.max(0.25, finite(baseHours, 11));
  const scenario = Math.max(0.25, finite(scenarioHours, 8));
  const included = [];
  const excluded = [];

  for (const source of Array.isArray(rows) ? rows : []) {
    const result = scenarioRow(source || {}, base, scenario);
    if (result.excluded) {
      excluded.push({
        asset_code: String(source?.asset_code || "-"),
        asset_name: String(source?.asset_name || ""),
        reason: result.reason,
      });
    } else {
      included.push(result.row);
    }
  }

  included.sort((a, b) => b.period_cost_saved - a.period_cost_saved || a.asset_code.localeCompare(b.asset_code));

  const fleetRaw = included.reduce((acc, row) => {
    acc.active_shifts += row.active_days;
    acc.base_operating_hours += row.active_days * base;
    acc.scenario_operating_hours += row.active_days * scenario;
    acc.base_fuel_liters += row.base_fuel_liters_per_shift * row.active_days;
    acc.scenario_fuel_liters += row.scenario_fuel_liters_per_shift * row.active_days;
    acc.base_fuel_cost += row.base_fuel_cost_per_shift * row.active_days;
    acc.scenario_fuel_cost += row.scenario_fuel_cost_per_shift * row.active_days;
    acc.nonfuel_cost += row.nonfuel_cost;
    return acc;
  }, {
    active_shifts: 0,
    base_operating_hours: 0,
    scenario_operating_hours: 0,
    base_fuel_liters: 0,
    scenario_fuel_liters: 0,
    base_fuel_cost: 0,
    scenario_fuel_cost: 0,
    nonfuel_cost: 0,
  });
  const fleetTotalBase = fleetRaw.base_fuel_cost + fleetRaw.nonfuel_cost;
  const fleetTotalScenario = fleetRaw.scenario_fuel_cost + fleetRaw.nonfuel_cost;
  const fleet = {
    ...fleetRaw,
    fuel_liters_saved: fleetRaw.base_fuel_liters - fleetRaw.scenario_fuel_liters,
    fuel_cost_saved: fleetRaw.base_fuel_cost - fleetRaw.scenario_fuel_cost,
    base_total_cost: fleetTotalBase,
    scenario_total_cost: fleetTotalScenario,
    total_cost_saved: fleetTotalBase - fleetTotalScenario,
    base_cost_per_operating_hour: fleetRaw.base_operating_hours > 0 ? fleetTotalBase / fleetRaw.base_operating_hours : 0,
    scenario_cost_per_operating_hour: fleetRaw.scenario_operating_hours > 0 ? fleetTotalScenario / fleetRaw.scenario_operating_hours : 0,
  };
  fleet.cost_per_operating_hour_change = fleet.scenario_cost_per_operating_hour - fleet.base_cost_per_operating_hour;

  const present = (row) => {
    const copy = {};
    for (const [key, value] of Object.entries(row)) {
      copy[key] = typeof value === "number" ? rounded(value, key.includes("liters_per_hour") ? 3 : 2) : value;
    }
    return copy;
  };

  return {
    base_hours: rounded(base, 2),
    scenario_hours: rounded(scenario, 2),
    rows: included.map(present),
    excluded,
    fleet: present(fleet),
  };
}
