/**
 * Machine-group daily prestart templates (heavy plant).
 * Matched from asset category + name; stored on vehicle_ldv_checks.check_mode as machine_prestart_<profileId>.
 */

const TEMPLATES = {
  excavator: {
    id: "excavator",
    title: "Excavator pre-start",
    sections: [
      {
        id: "fluids",
        title: "Fluid levels",
        items: [
          { key: "engine_oil_ok", label: "Engine oil level OK" },
          { key: "coolant_ok", label: "Coolant level OK" },
          { key: "hydraulic_oil_ok", label: "Hydraulic oil level OK" },
          { key: "swing_gear_oil_ok", label: "Swing / slew gear oil OK (if applicable)" },
          { key: "final_drive_ok", label: "Final drives / track gear oil OK" },
        ],
      },
      {
        id: "undercarriage",
        title: "Undercarriage",
        items: [
          { key: "track_tension_ok", label: "Track tension OK" },
          { key: "rollers_idlers_ok", label: "Rollers & idlers OK (no seized / flat spots)" },
          { key: "sprocket_ok", label: "Sprocket teeth OK" },
          { key: "cuts_leaks_ok", label: "No abnormal cuts, cracks, or leaks on tracks" },
        ],
      },
      {
        id: "safety",
        title: "Safety & cab",
        items: [
          { key: "fire_extinguisher_ok", label: "Fire extinguisher present & charged" },
          { key: "seat_belt_ok", label: "Seat belt OK" },
          { key: "mirrors_camera_ok", label: "Mirrors / cameras clean & working" },
          { key: "horn_estop_ok", label: "Horn & emergency stop OK" },
          { key: "windows_guard_ok", label: "Windows / guards intact" },
        ],
      },
    ],
  },
  dozer: {
    id: "dozer",
    title: "Dozer pre-start",
    sections: [
      {
        id: "fluids",
        title: "Fluid levels",
        items: [
          { key: "engine_oil_ok", label: "Engine oil level OK" },
          { key: "coolant_ok", label: "Coolant level OK" },
          { key: "hydraulic_oil_ok", label: "Hydraulic / transmission oil OK" },
          { key: "final_drive_ok", label: "Final drives OK" },
        ],
      },
      {
        id: "undercarriage",
        title: "Undercarriage",
        items: [
          { key: "track_tension_ok", label: "Track tension OK" },
          { key: "blade_pins_ok", label: "Blade / ripper pins & hydraulics OK" },
          { key: "rollers_sprocket_ok", label: "Rollers & sprocket OK" },
        ],
      },
      {
        id: "safety",
        title: "Safety",
        items: [
          { key: "fire_extinguisher_ok", label: "Fire extinguisher OK" },
          { key: "seat_belt_ok", label: "Seat belt OK" },
          { key: "horn_estop_ok", label: "Horn & E-stop OK" },
        ],
      },
    ],
  },
  wheel_loader: {
    id: "wheel_loader",
    title: "Wheel loader pre-start",
    sections: [
      {
        id: "fluids",
        title: "Fluid levels",
        items: [
          { key: "engine_oil_ok", label: "Engine oil OK" },
          { key: "coolant_ok", label: "Coolant OK" },
          { key: "hydraulic_oil_ok", label: "Hydraulic oil OK" },
          { key: "transmission_ok", label: "Transmission / axle oils OK" },
          { key: "brake_fluid_ok", label: "Brake fluid OK" },
        ],
      },
      {
        id: "machine",
        title: "Machine & tyres",
        items: [
          { key: "tyre_condition_ok", label: "Tyres — pressure & damage OK" },
          { key: "articulation_ok", label: "Centre articulation / pins OK" },
          { key: "bucket_linkage_ok", label: "Bucket & linkage OK" },
        ],
      },
      {
        id: "safety",
        title: "Safety",
        items: [
          { key: "lights_beacon_ok", label: "Lights & beacon OK" },
          { key: "horn_reversing_ok", label: "Horn & reversing alarm OK" },
          { key: "seat_belt_ok", label: "Seat belt OK" },
        ],
      },
    ],
  },
  haul_truck: {
    id: "haul_truck",
    title: "Haul truck pre-start",
    sections: [
      {
        id: "fluids",
        title: "Fluid levels",
        items: [
          { key: "engine_oil_ok", label: "Engine oil OK" },
          { key: "coolant_ok", label: "Coolant OK" },
          { key: "hydraulic_oil_ok", label: "Hydraulic oil OK" },
          { key: "brake_fluid_ok", label: "Brake fluid OK" },
        ],
      },
      {
        id: "machine",
        title: "Body & running gear",
        items: [
          { key: "tyre_condition_ok", label: "Tyres OK" },
          { key: "body_lifts_ok", label: "Body / hoist & tail door OK" },
          { key: "steering_ok", label: "Steering free play OK" },
        ],
      },
      {
        id: "safety",
        title: "Safety",
        items: [
          { key: "lights_beacon_ok", label: "Lights & beacon OK" },
          { key: "seat_belt_ok", label: "Seat belt OK" },
          { key: "horn_reversing_ok", label: "Horn & reversing alarm OK" },
        ],
      },
    ],
  },
  grader: {
    id: "grader",
    title: "Grader pre-start",
    sections: [
      {
        id: "fluids",
        title: "Fluid levels",
        items: [
          { key: "engine_oil_ok", label: "Engine oil OK" },
          { key: "coolant_ok", label: "Coolant OK" },
          { key: "hydraulic_oil_ok", label: "Hydraulic oil OK" },
          { key: "circle_drive_ok", label: "Circle drive oil OK (if applicable)" },
        ],
      },
      {
        id: "machine",
        title: "Moldboard & tyres",
        items: [
          { key: "moldboard_ok", label: "Moldboard & linkage OK" },
          { key: "tyre_condition_ok", label: "Tyres OK" },
        ],
      },
      {
        id: "safety",
        title: "Safety",
        items: [
          { key: "lights_beacon_ok", label: "Lights & beacon OK" },
          { key: "horn_ok", label: "Horn OK" },
          { key: "seat_belt_ok", label: "Seat belt OK" },
        ],
      },
    ],
  },
  mobile_crane: {
    id: "mobile_crane",
    title: "Mobile crane pre-start",
    sections: [
      {
        id: "fluids",
        title: "Fluid levels",
        items: [
          { key: "engine_oil_ok", label: "Engine oil OK" },
          { key: "coolant_ok", label: "Coolant OK" },
          { key: "hydraulic_oil_ok", label: "Hydraulic oil OK" },
        ],
      },
      {
        id: "crane",
        title: "Crane & carrier",
        items: [
          { key: "outriggers_ok", label: "Outriggers / pads OK" },
          { key: "wire_rope_ok", label: "Wire rope / hook block OK" },
          { key: "tyre_condition_ok", label: "Carrier tyres OK" },
        ],
      },
      {
        id: "safety",
        title: "Safety",
        items: [
          { key: "load_charts_ok", label: "Load chart / duty indicator accessible" },
          { key: "horn_estop_ok", label: "Horn & E-stop OK" },
          { key: "seat_belt_ok", label: "Seat belt OK" },
        ],
      },
    ],
  },
  crusher: {
    id: "crusher",
    title: "Mobile crusher pre-start",
    sections: [
      {
        id: "fluids",
        title: "Fluid levels",
        items: [
          { key: "engine_oil_ok", label: "Engine oil level OK" },
          { key: "coolant_ok", label: "Coolant level OK" },
          { key: "hydraulic_oil_ok", label: "Hydraulic oil level OK" },
          { key: "grease_points_ok", label: "Auto / manual grease system OK" },
        ],
      },
      {
        id: "crushing",
        title: "Crushing plant",
        items: [
          { key: "feed_hopper_ok", label: "Feed hopper, grizzly & feeder OK" },
          { key: "crusher_liner_ok", label: "Jaw / cone liner & chamber OK (no loose parts)" },
          { key: "discharge_conveyor_ok", label: "Discharge conveyor, belt & scrapers OK" },
          { key: "dust_suppression_ok", label: "Dust suppression / water spray OK" },
          { key: "leaks_damage_ok", label: "No abnormal leaks, cracks, or damage" },
        ],
      },
      {
        id: "running_gear",
        title: "Tracks / undercarriage",
        items: [
          { key: "track_tension_ok", label: "Track tension OK" },
          { key: "rollers_idlers_ok", label: "Rollers & idlers OK" },
          { key: "sprocket_ok", label: "Sprocket teeth OK" },
        ],
      },
      {
        id: "safety",
        title: "Safety",
        items: [
          { key: "guards_ok", label: "Guards & covers in place" },
          { key: "horn_estop_ok", label: "Horn & emergency stop OK" },
          { key: "fire_extinguisher_ok", label: "Fire extinguisher present & charged" },
          { key: "seat_belt_ok", label: "Seat belt OK" },
          { key: "lights_beacon_ok", label: "Lights & beacon OK" },
        ],
      },
    ],
  },
  mobile_screen: {
    id: "mobile_screen",
    title: "Mobile screen pre-start",
    sections: [
      {
        id: "fluids",
        title: "Fluid levels",
        items: [
          { key: "engine_oil_ok", label: "Engine oil level OK" },
          { key: "coolant_ok", label: "Coolant level OK" },
          { key: "hydraulic_oil_ok", label: "Hydraulic oil level OK" },
        ],
      },
      {
        id: "screening",
        title: "Screening plant",
        items: [
          { key: "screen_media_ok", label: "Screen panels / mesh & tension OK" },
          { key: "feed_conveyor_ok", label: "Feed conveyor & chutes OK" },
          { key: "discharge_conveyors_ok", label: "Discharge conveyors & belts OK" },
          { key: "springs_mounts_ok", label: "Springs, mounts & vibration unit OK" },
          { key: "leaks_damage_ok", label: "No abnormal leaks, cracks, or damage" },
        ],
      },
      {
        id: "running_gear",
        title: "Tracks / undercarriage",
        items: [
          { key: "track_tension_ok", label: "Track tension OK" },
          { key: "rollers_idlers_ok", label: "Rollers & idlers OK" },
          { key: "sprocket_ok", label: "Sprocket teeth OK" },
        ],
      },
      {
        id: "safety",
        title: "Safety",
        items: [
          { key: "guards_ok", label: "Guards & covers in place" },
          { key: "horn_estop_ok", label: "Horn & emergency stop OK" },
          { key: "fire_extinguisher_ok", label: "Fire extinguisher present & charged" },
          { key: "seat_belt_ok", label: "Seat belt OK" },
          { key: "lights_beacon_ok", label: "Lights & beacon OK" },
        ],
      },
    ],
  },
};

export function getMachinePrestartTemplate(profileId) {
  const id = String(profileId || "").trim().toLowerCase();
  return TEMPLATES[id] || null;
}

export function listMachinePrestartProfiles() {
  return Object.values(TEMPLATES).map((t) => ({
    id: t.id,
    title: t.title,
  }));
}

/** Site mobile crushers (CR01AM–CR05AM) and screens (e.g. FIN694). */
const CRUSHER_ASSET_CODES = new Set(["CR01AM", "CR02AM", "CR03AM", "CR04AM", "CR05AM"]);
const MOBILE_SCREEN_ASSET_CODES = new Set(["FIN694"]);

function resolveProfileByAssetCode(assetCode) {
  const code = String(assetCode || "").trim().toUpperCase();
  if (!code) return null;
  if (CRUSHER_ASSET_CODES.has(code) || /^CR0[1-9]AM$/.test(code)) return "crusher";
  if (MOBILE_SCREEN_ASSET_CODES.has(code)) return "mobile_screen";
  return null;
}

/**
 * Resolve which machine prestart profile applies from asset code, category, and name.
 */
export function resolveMachinePrestartProfile(category, assetName = "", assetCode = "") {
  const byCode = resolveProfileByAssetCode(assetCode);
  if (byCode) return byCode;

  const hay = `${String(category || "")} ${String(assetName || "")}`.toLowerCase();
  if (!hay.trim()) return null;
  if (/(mobile\s*)?crusher|crushing\s*plant|jaw\s*crusher|cone\s*crusher|\bcrusher\b/.test(hay)) return "crusher";
  if (/(mobile\s*)?(screen|screener|finisher|screening\s*plant)/.test(hay)) return "mobile_screen";
  if (/(excavator|digger|shovel)/.test(hay)) return "excavator";
  if (/(dozer|bulldozer)/.test(hay)) return "dozer";
  if (/(wheel\s*loader|front\s*end\s*loader|\bloader\b)/.test(hay) && !/excavator/.test(hay)) return "wheel_loader";
  if (/(motor\s*grader|\bgrader\b)/.test(hay)) return "grader";
  if (/(haul\s*truck|rigid\s*dump|articulated\s*dump|\badt\b|dump\s*truck)/.test(hay)) return "haul_truck";
  if (/(mobile\s*crane|pick\s*and\s*carry|\bcrane\b)/.test(hay)) return "mobile_crane";
  return null;
}

export function machinePrestartCheckMode(profileId) {
  const id = String(profileId || "").trim().toLowerCase();
  if (!getMachinePrestartTemplate(id)) return null;
  return `machine_prestart_${id}`;
}

function flattenTemplateItems(template) {
  if (!template?.sections) return [];
  const out = [];
  for (const sec of template.sections) {
    for (const it of sec.items || []) {
      out.push({ sectionId: sec.id, key: it.key, label: it.label });
    }
  }
  return out;
}

export function normalizeMachinePrestartChecklist(profileId, input) {
  const template = getMachinePrestartTemplate(profileId);
  if (!template) return [];
  const src = input && typeof input === "object" && !Array.isArray(input) ? input : {};
  return flattenTemplateItems(template).map((row) => ({
    key: row.key,
    label: row.label,
    section_id: row.sectionId,
    ok: src[row.key] === true,
  }));
}

export function checklistToJsonObject(checklistArray) {
  const obj = {};
  for (const row of Array.isArray(checklistArray) ? checklistArray : []) {
    const k = String(row?.key || "").trim();
    if (!k) continue;
    obj[k] = Boolean(row?.ok);
  }
  return obj;
}
