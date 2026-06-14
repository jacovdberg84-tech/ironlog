export const UNDERCARRIAGE_CARRIER_ROLLER_IDS = ["X", "Y", "Z", "T"];
export const UNDERCARRIAGE_TRACK_ROLLER_COUNT = 12;

export const UNDERCARRIAGE_WEAR_BANDS = [
  { key: "good", label: "Good", min: 0, max: 75, color: "#22c55e", argb: "FF22C55E" },
  { key: "warn", label: "Monitor", min: 76, max: 100, color: "#eab308", argb: "FFEAB308" },
  { key: "critical", label: "Plan replace", min: 101, max: 120, color: "#f97316", argb: "FFF97316" },
  { key: "replace", label: "Replace now", min: 121, max: Infinity, color: "#ef4444", argb: "FFEF4444" },
];

export const UNDERCARRIAGE_CHECKLIST_ITEMS = [
  { key: "oil_leaks", label: "Oil leaks" },
  { key: "loose_bolts", label: "Loose bolts" },
  { key: "cracked_broken", label: "Cracked / broken" },
  { key: "packing_mud", label: "Packing / mud" },
];

export const UNDERCARRIAGE_TRACK_SAG_POINTS = ["A", "B", "C", "D"];

export function buildUndercarriageComponentSchema() {
  const rows = [];
  const addPair = (group, label, wearDirection) => {
    for (const side of ["LH", "RH"]) {
      rows.push({
        key: `${group}_${side.toLowerCase()}`,
        group,
        label,
        side,
        wear_direction: wearDirection,
      });
    }
  };

  addPair("bushings", "Bushings", "down");
  addPair("links", "Links", "up");
  addPair("pins", "Pins", "down");
  addPair("track_shoe", "Track Shoe", "down");

  for (const id of UNDERCARRIAGE_CARRIER_ROLLER_IDS) {
    for (const side of ["LH", "RH"]) {
      rows.push({
        key: `carrier_${id.toLowerCase()}_${side.toLowerCase()}`,
        group: "carrier_rollers",
        label: `Carrier Roller ${id}`,
        side,
        roller_id: id,
        wear_direction: "down",
      });
    }
  }

  for (let n = 1; n <= UNDERCARRIAGE_TRACK_ROLLER_COUNT; n += 1) {
    for (const side of ["LH", "RH"]) {
      rows.push({
        key: `track_roller_${n}_${side.toLowerCase()}`,
        group: "track_rollers",
        label: `Track Roller ${n}`,
        side,
        roller_num: n,
        wear_direction: "down",
      });
    }
  }

  addPair("grouser_height", "Grouser Height", "down");
  return rows;
}

export function undercarriageOptionalNumber(raw) {
  if (raw === "" || raw == null) return null;
  const n = Number(raw);
  return Number.isFinite(n) ? n : null;
}

export function calcUndercarriageWearPct(measurement, base, wearLimit, wearDirection = "down") {
  const m = undercarriageOptionalNumber(measurement);
  const b = undercarriageOptionalNumber(base);
  const w = undercarriageOptionalNumber(wearLimit);
  if (m == null || b == null || w == null) return null;
  let pct;
  if (String(wearDirection || "down").toLowerCase() === "up") {
    const denom = w - b;
    if (!Number.isFinite(denom) || denom <= 0) return null;
    pct = ((m - b) / denom) * 100;
  } else {
    const denom = b - w;
    if (!Number.isFinite(denom) || denom <= 0) return null;
    pct = ((b - m) / denom) * 100;
  }
  return Number(Math.max(0, pct).toFixed(2));
}

export function undercarriageWearBand(pct) {
  const n = undercarriageOptionalNumber(pct);
  if (n == null) return null;
  for (const band of UNDERCARRIAGE_WEAR_BANDS) {
    if (n >= band.min && n <= band.max) return band;
  }
  return UNDERCARRIAGE_WEAR_BANDS[UNDERCARRIAGE_WEAR_BANDS.length - 1];
}

export function calcUndercarriageLifeExpectancy({
  currentHours,
  currentPct,
  previousHours,
  previousPct,
  targetPct = 100,
}) {
  const hrs = undercarriageOptionalNumber(currentHours);
  const pct = undercarriageOptionalNumber(currentPct);
  const prevHrs = undercarriageOptionalNumber(previousHours);
  const prevPct = undercarriageOptionalNumber(previousPct);
  const target = undercarriageOptionalNumber(targetPct) ?? 100;
  if (hrs == null || pct == null || prevHrs == null || prevPct == null) return null;
  const deltaHours = hrs - prevHrs;
  const deltaPct = pct - prevPct;
  if (deltaHours <= 0 || deltaPct <= 0) return null;
  const hrsPerPct = deltaHours / deltaPct;
  const remaining = Math.max(0, target - pct);
  if (remaining <= 0) return 0;
  return Number((remaining * hrsPerPct).toFixed(0));
}

export function enrichUndercarriageMeasurement(row, {
  currentHours = null,
  previousRow = null,
} = {}) {
  const wear_direction = row?.wear_direction || "down";
  const pct = calcUndercarriageWearPct(row?.measurement, row?.base, row?.wear_limit, wear_direction);
  const band = undercarriageWearBand(pct);
  const prevPct = previousRow ? calcUndercarriageWearPct(
    previousRow.measurement,
    previousRow.base ?? row?.base,
    previousRow.wear_limit ?? row?.wear_limit,
    wear_direction,
  ) : null;
  const prevHours = undercarriageOptionalNumber(previousRow?.inspection_hours);
  const hrs = undercarriageOptionalNumber(currentHours);
  let wear_usage_pct = null;
  let wear_rate_pct_per_hour = null;
  if (pct != null && prevPct != null) wear_usage_pct = Number((pct - prevPct).toFixed(2));
  if (wear_usage_pct != null && hrs != null && prevHours != null) {
    const deltaHours = hrs - prevHours;
    if (deltaHours > 0) wear_rate_pct_per_hour = Number((wear_usage_pct / deltaHours).toFixed(4));
  }
  const life_expectancy_hours = calcUndercarriageLifeExpectancy({
    currentHours: hrs,
    currentPct: pct,
    previousHours: prevHours,
    previousPct: prevPct,
  });

  return {
    ...row,
    wear_pct: pct,
    wear_band: band?.key || null,
    wear_band_label: band?.label || null,
    wear_usage_pct,
    wear_rate_pct_per_hour,
    life_expectancy_hours,
  };
}

export function normalizeUndercarriageMeasurements(raw, schemaRows = null) {
  const schema = Array.isArray(schemaRows) && schemaRows.length
    ? schemaRows
    : buildUndercarriageComponentSchema();
  const schemaByKey = new Map(schema.map((s) => [String(s.key || "").toLowerCase(), s]));
  const input = Array.isArray(raw) ? raw : [];
  const inputByKey = new Map(input.map((r) => [String(r?.key || r?.component_key || "").toLowerCase(), r]));

  return schema.map((schemaRow) => {
    const key = String(schemaRow.key || "").toLowerCase();
    const src = inputByKey.get(key) || {};
    return {
      key,
      group: schemaRow.group,
      label: schemaRow.label,
      side: schemaRow.side,
      roller_id: schemaRow.roller_id || null,
      roller_num: schemaRow.roller_num || null,
      wear_direction: schemaRow.wear_direction || "down",
      measurement: undercarriageOptionalNumber(src.measurement),
      base: undercarriageOptionalNumber(src.base),
      wear_limit: undercarriageOptionalNumber(src.wear_limit),
    };
  });
}

export function normalizeUndercarriageTrackSag(raw) {
  const out = {};
  for (const pt of UNDERCARRIAGE_TRACK_SAG_POINTS) {
    out[pt] = undercarriageOptionalNumber(raw?.[pt]);
  }
  return out;
}

export function normalizeUndercarriageChecklist(raw) {
  const out = {
    general_condition: String(raw?.general_condition || "").trim().toLowerCase() || null,
    comments: String(raw?.comments || "").trim() || null,
    items: {},
  };
  for (const item of UNDERCARRIAGE_CHECKLIST_ITEMS) {
    const sideRaw = raw?.items?.[item.key] || raw?.[item.key] || {};
    out.items[item.key] = {
      lh: sideRaw.lh === true || String(sideRaw.lh || "").toLowerCase() === "yes",
      rh: sideRaw.rh === true || String(sideRaw.rh || "").toLowerCase() === "yes",
    };
  }
  return out;
}

export function summarizeUndercarriageInspection(measurements) {
  const rows = Array.isArray(measurements) ? measurements : [];
  const withPct = rows.filter((r) => r.wear_pct != null);
  const worst = withPct.reduce((best, row) => {
    if (!best || Number(row.wear_pct) > Number(best.wear_pct)) return row;
    return best;
  }, null);
  const counts = { good: 0, warn: 0, critical: 0, replace: 0, unknown: 0 };
  for (const row of rows) {
    const band = row.wear_band || "unknown";
    counts[band] = (counts[band] || 0) + 1;
  }
  return {
    measured_count: withPct.length,
    component_count: rows.length,
    worst_wear_pct: worst?.wear_pct ?? null,
    worst_component: worst ? `${worst.label} (${worst.side})` : null,
    wear_band_counts: counts,
  };
}
