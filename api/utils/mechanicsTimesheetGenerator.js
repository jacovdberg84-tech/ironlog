// Generate mechanics labour timesheet rows (startup inspections + real work + varied fill).

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

const STARTUP_HOURS = 0.5;
const DEFAULT_MIN_TECH_HOURS = 6;
const DEFAULT_SYNTHETIC_THROUGH = "2026-03-31";

const EXCLUDED_ASSET_CODES = new Set([
  "E017", "E018", "E025", "CRT01AM", "FIN694",
]);

const TECH_BY_GROUP = {
  dumptruck: ["Sergio", "Arnold"],
  loader: ["Ronnie"],
  excavator: ["Charles"],
  crusher_screen: ["Moses"],
};

const JOB_POOLS = {
  Service: [
    "250hr PM service",
    "500hr service",
    "1000hr service",
    "2000hr major service",
    "Engine oil and filter change",
    "Hydraulic oil and filter change",
    "Transmission service",
    "Diff oil change",
    "Final drive oil change",
    "Swing bearing grease service",
    "Fuel system filter replacement",
    "Coolant flush and refill",
    "Air filter replacement",
    "Cab air filter replacement",
    "Track roller inspection service",
  ],
  Maintenance: [
    "Grease all grease points",
    "Tyre pressure check and adjust",
    "Track tension check and adjust",
    "Battery terminals clean and tighten",
    "Fan belt inspection and adjust",
    "Radiator external clean",
    "Work lights check and repair",
    "Mirror adjustment and mount check",
    "Fire extinguisher inspection",
    "Seat belt inspection",
    "Reverse alarm test and repair",
    "Wheel nut torque check",
    "Pivot pin wear inspection",
    "Undercarriage visual inspection",
    "Hydraulic cylinder rod inspection",
  ],
  Breakdown: [
    "Hydraulic hose replacement",
    "Front spindle bearing failure",
    "Remove and inspect diff",
    "Brake system repair",
    "Cooling system leak repair",
    "Electrical fault finding",
    "Starter motor replacement",
    "Alternator replacement",
    "Turbo oil feed line repair",
    "Steering cylinder seal replacement",
    "Bucket cylinder repack",
    "Transmission oil leak",
    "Drive shaft UJ replacement",
    "Park brake adjustment",
    "Wiring harness repair",
    "Aircon compressor fault",
    "Fuel pump replacement",
    "Radiator hose replacement",
  ],
};

const GENERIC_WORK_NOTES = new Set([
  "repaired",
  "work was done",
  "problem repaired",
  "done",
  "completed",
  "fixed",
  "mechanical",
  "mechanical breakdown",
]);

function hasTable(db, table) {
  try {
    const r = db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name=? LIMIT 1`).get(String(table));
    return Boolean(r?.name);
  } catch {
    return false;
  }
}

function hasColumn(db, table, col) {
  try {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name) === col);
  } catch {
    return false;
  }
}

function isExcludedAsset(code) {
  return EXCLUDED_ASSET_CODES.has(String(code || "").trim().toUpperCase());
}

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function fmtNum(v, digits = 2) {
  const n = Number(v);
  if (!Number.isFinite(n)) return null;
  return Number(n.toFixed(digits));
}

function formatDisplayDate(ymd) {
  const parts = String(ymd || "").split("-").map(Number);
  if (parts.length !== 3 || !parts.every((n) => Number.isFinite(n))) return String(ymd || "");
  const [y, m, d] = parts;
  const month = MONTH_NAMES[m - 1] || "";
  return `${String(d).padStart(2, "0")} ${month} ${y}`;
}

function hashSeed(...parts) {
  const s = parts.join("|");
  let h = 0;
  for (let i = 0; i < s.length; i += 1) h = (h * 31 + s.charCodeAt(i)) % 2147483647;
  return h;
}

function shuffled(arr, seed) {
  const copy = [...arr];
  for (let i = copy.length - 1; i > 0; i -= 1) {
    const j = hashSeed(seed, i, "shuffle") % (i + 1);
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

function isWorkingDay(ymd) {
  const d = new Date(`${ymd}T12:00:00`);
  if (Number.isNaN(d.getTime())) return false;
  return d.getDay() !== 0;
}

function enumerateDates(from, to) {
  const out = [];
  if (!isDate(from) || !isDate(to) || from > to) return out;
  const cur = new Date(`${from}T12:00:00`);
  const end = new Date(`${to}T12:00:00`);
  while (cur <= end) {
    const y = cur.getFullYear();
    const m = String(cur.getMonth() + 1).padStart(2, "0");
    const d = String(cur.getDate()).padStart(2, "0");
    const ymd = `${y}-${m}-${d}`;
    if (isWorkingDay(ymd)) out.push(ymd);
    cur.setDate(cur.getDate() + 1);
  }
  return out;
}

function classifyAsset(asset) {
  const code = String(asset.asset_code || "").trim().toUpperCase();
  if (isExcludedAsset(code)) return null;
  const cat = String(asset.category || "").toLowerCase();
  const name = String(asset.asset_name || "").toLowerCase();
  const hay = `${cat} ${name}`;

  if (/adt|dump|haul\s*truck|articulated/.test(hay) || /^A\d+AM$/.test(code)) return "dumptruck";
  if (/excavator|digger|shovel/.test(hay) || /^E\d+AM$/.test(code)) return "excavator";
  if (/^F\d+AM$/.test(code) || (/\bloader\b/.test(hay) && !/forklift|tlb|backhoe|block/.test(hay))) {
    return "loader";
  }
  if (/crusher|screen|trommel/.test(hay) || /^CR\d+AM$/.test(code)) return "crusher_screen";
  return null;
}

function assignDumptruckTech(assetCode, dumptruckCodes) {
  const idx = dumptruckCodes.indexOf(String(assetCode || "").toUpperCase());
  return idx % 2 === 0 ? "Sergio" : "Arnold";
}

function techForAsset(group, assetCode, dumptruckCodes) {
  const techs = TECH_BY_GROUP[group] || [];
  if (!techs.length) return "Workshop";
  if (group === "dumptruck") return assignDumptruckTech(assetCode, dumptruckCodes);
  return techs[0];
}

function timeToMinutes(t) {
  const [hh, mm] = String(t || "0:0").split(":").map((x) => Number(x) || 0);
  return hh * 60 + mm;
}

function minutesToTime(totalMinutes) {
  const m = Math.max(0, Math.round(totalMinutes));
  const hh = Math.floor(m / 60);
  const mm = m % 60;
  return `${hh}:${String(mm).padStart(2, "0")}`;
}

function addMinutesToTime(timeStr, minutes) {
  const [hh, mm] = String(timeStr || "0:0").split(":").map((x) => Number(x) || 0);
  return minutesToTime(hh * 60 + mm + minutes);
}

function techDayStartMinutes(date, tech) {
  return 6 * 60 + (hashSeed(date, tech, "daystart") % 4) * 15;
}

function cleanBreakdownDesc(desc, component) {
  let s = String(desc || "")
    .replace(/^DOWN\s*[—\-]\s*/i, "")
    .replace(/^DOWN\s+/i, "")
    .trim();
  if (GENERIC_WORK_NOTES.has(s.toLowerCase()) && component) {
    s = `${String(component).trim()} failure`;
  }
  if (!s || GENERIC_WORK_NOTES.has(s.toLowerCase())) return "Mechanical breakdown";
  return s;
}

function isGenericNote(text) {
  const t = String(text || "").trim().toLowerCase();
  return !t || GENERIC_WORK_NOTES.has(t);
}

function inferCategory(reason, explicit) {
  if (explicit) return String(explicit);
  const r = String(reason || "").toLowerCase();
  if (r.includes("startup")) return "Startup";
  if (/\b\d+hr\b|\bpm service\b|\bservice\b/.test(r) && !r.includes("breakdown")) return "Service";
  if (/grease|inspect|check|adjust|torque|clean/.test(r)) return "Maintenance";
  return "Breakdown";
}

function woSourceCategory(source) {
  const s = String(source || "").toLowerCase();
  if (s === "service") return "Service";
  if (s === "manual" || s === "maintenance") return "Maintenance";
  return "Breakdown";
}

function describeWorkOrder(row) {
  if (!isGenericNote(row.completion_notes)) return String(row.completion_notes).trim();
  if (!isGenericNote(row.job_description)) return String(row.job_description).trim();
  if (row.service_name) return `${String(row.service_name).trim()}hr service`;
  return null;
}

class UsageTracker {
  constructor(maxRecent = 10) {
    this.maxRecent = maxRecent;
    this.techRecent = new Map();
    this.techCounts = new Map();
  }

  record(tech, desc) {
    const key = String(desc || "").trim();
    if (!key) return;
    const recent = this.techRecent.get(tech) || [];
    recent.push(key);
    while (recent.length > this.maxRecent) recent.shift();
    this.techRecent.set(tech, recent);
    const counts = this.techCounts.get(tech) || new Map();
    counts.set(key, (counts.get(key) || 0) + 1);
    this.techCounts.set(tech, counts);
  }

  pick(tech, date, category, pool, salt) {
    const list = pool.filter(Boolean);
    if (!list.length) return "General workshop duties";
    const recent = new Set(this.techRecent.get(tech) || []);
    let candidates = list.filter((d) => !recent.has(d));
    if (!candidates.length) candidates = list;

    const counts = this.techCounts.get(tech) || new Map();
    const minCount = Math.min(...candidates.map((d) => counts.get(d) || 0));
    const tier = candidates.filter((d) => (counts.get(d) || 0) === minCount);
    const desc = tier[hashSeed(date, tech, salt, category, "pick") % tier.length];
    this.record(tech, desc);
    return desc;
  }
}

function resolveSmr(db, assetId, usageDate, cache) {
  const aid = Number(assetId);
  const d = String(usageDate || "").trim();
  if (!Number.isFinite(aid) || aid <= 0 || !d) return null;
  const key = `${aid}|${d}`;
  if (cache.has(key)) return cache.get(key);

  let smr = null;
  if (hasTable(db, "daily_hours") && hasColumn(db, "daily_hours", "closing_hours")) {
    const exact = db.prepare(`
      SELECT closing_hours FROM daily_hours
      WHERE asset_id = ? AND work_date = ? AND closing_hours IS NOT NULL
      LIMIT 1
    `).get(aid, d);
    if (exact?.closing_hours != null) {
      smr = fmtNum(exact.closing_hours, 1);
    } else {
      const near = db.prepare(`
        SELECT closing_hours FROM daily_hours
        WHERE asset_id = ? AND work_date <= ? AND closing_hours IS NOT NULL
        ORDER BY work_date DESC LIMIT 1
      `).get(aid, d);
      if (near?.closing_hours != null) smr = fmtNum(near.closing_hours, 1);
    }
  }
  if (smr == null && hasTable(db, "asset_hours")) {
    const ah = db.prepare(`SELECT total_hours FROM asset_hours WHERE asset_id = ?`).get(aid);
    if (ah?.total_hours != null) smr = fmtNum(ah.total_hours, 1);
  }
  cache.set(key, smr);
  return smr;
}

function loadStartupFleet(db) {
  if (!hasTable(db, "assets")) return { groups: {}, assetByCode: new Map() };
  const assets = db.prepare(`
    SELECT id, asset_code, asset_name, category
    FROM assets
    WHERE COALESCE(active, 1) = 1
    ORDER BY asset_code ASC
  `).all();

  const groups = {
    dumptruck: [],
    loader: [],
    excavator: [],
    crusher_screen: [],
  };
  const assetByCode = new Map();

  for (const a of assets) {
    const code = String(a.asset_code || "").trim().toUpperCase();
    if (isExcludedAsset(code)) continue;
    const group = classifyAsset({ ...a, asset_code: code });
    if (!group) continue;
    assetByCode.set(code, { ...a, asset_code: code });
    groups[group].push({ ...a, asset_code: code });
  }

  return { groups, assetByCode };
}

function loadServiceCatalog(db) {
  if (!hasTable(db, "maintenance_plans") || !hasTable(db, "assets")) return [];
  const rows = db.prepare(`
    SELECT mp.service_name, mp.interval_hours, a.id AS asset_id, a.asset_code
    FROM maintenance_plans mp
    JOIN assets a ON a.id = mp.asset_id
    WHERE COALESCE(mp.active, 1) = 1
    ORDER BY a.asset_code ASC, mp.interval_hours ASC
  `).all();
  return rows
    .map((r) => ({
      asset_id: Number(r.asset_id),
      asset_code: String(r.asset_code || "").toUpperCase(),
      service_name: String(r.service_name || "").trim(),
      interval_hours: Number(r.interval_hours || 0),
    }))
    .filter((r) => r.asset_code && !isExcludedAsset(r.asset_code) && r.service_name);
}

function loadBreakdownsByDate(db, from, to) {
  if (!hasTable(db, "breakdowns") || !hasTable(db, "assets")) return new Map();
  const hasWo = hasTable(db, "work_orders");
  const hasDowntimeLogs = hasTable(db, "breakdown_downtime_logs");
  const hasComponent = hasColumn(db, "breakdowns", "component");

  const rows = db.prepare(`
    SELECT
      b.id,
      b.breakdown_date,
      b.description,
      ${hasComponent ? "b.component" : "NULL AS component"},
      b.asset_id,
      a.asset_code,
      a.asset_name,
      a.category,
      ${hasWo ? "w.labor_hours" : "NULL AS labor_hours"},
      ${hasWo && hasColumn(db, "work_orders", "completion_notes") ? "w.completion_notes" : "NULL AS completion_notes"}
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    ${hasWo ? "LEFT JOIN work_orders w ON w.id = b.primary_work_order_id" : ""}
    WHERE b.breakdown_date >= ? AND b.breakdown_date <= ?
    ORDER BY b.breakdown_date ASC, b.id ASC
  `).all(from, to);

  const downtimeByBdDate = new Map();
  if (hasDowntimeLogs) {
    const logs = db.prepare(`
      SELECT breakdown_id, log_date, hours_down
      FROM breakdown_downtime_logs
      WHERE log_date >= ? AND log_date <= ?
    `).all(from, to);
    for (const l of logs) {
      const key = `${Number(l.breakdown_id)}|${String(l.log_date)}`;
      downtimeByBdDate.set(key, Number(l.hours_down || 0));
    }
  }

  const byDate = new Map();
  for (const r of rows) {
    const code = String(r.asset_code || "").toUpperCase();
    if (isExcludedAsset(code)) continue;
    const date = String(r.breakdown_date || "");
    if (!date) continue;

    const woNote = describeWorkOrder(r);
    const description = woNote && !isGenericNote(woNote)
      ? woNote
      : cleanBreakdownDesc(r.description, r.component);

    const bdKey = `${Number(r.id)}|${date}`;
    const downtime = downtimeByBdDate.get(bdKey) || 0;
    let hours = Number(r.labor_hours || 0);
    if (!Number.isFinite(hours) || hours <= 0) {
      hours = downtime > 0 ? Math.min(8, Math.max(2, downtime * 0.6)) : 3 + (hashSeed(date, code, description) % 3);
    }
    hours = Math.max(1, Math.min(8, Number(hours.toFixed(1))));

    if (!byDate.has(date)) byDate.set(date, []);
    byDate.get(date).push({
      asset_id: Number(r.asset_id),
      asset_code: code,
      description,
      hours,
      category: "Breakdown",
    });
  }
  return byDate;
}

function loadCompletedWorkByDate(db, from, to) {
  if (!hasTable(db, "work_orders") || !hasTable(db, "assets")) return new Map();
  const hasClosed = hasColumn(db, "work_orders", "closed_at");
  if (!hasClosed) return new Map();

  const hasPlans = hasTable(db, "maintenance_plans");
  const rows = db.prepare(`
    SELECT
      date(w.closed_at) AS work_date,
      w.source,
      ${hasColumn(db, "work_orders", "labor_hours") ? "w.labor_hours" : "0 AS labor_hours"},
      ${hasColumn(db, "work_orders", "completion_notes") ? "w.completion_notes" : "'' AS completion_notes"},
      ${hasColumn(db, "work_orders", "job_description") ? "w.job_description" : "'' AS job_description"},
      a.id AS asset_id,
      a.asset_code,
      ${hasPlans ? "mp.service_name" : "NULL AS service_name"}
    FROM work_orders w
    JOIN assets a ON a.id = w.asset_id
    ${hasPlans ? "LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id AND LOWER(w.source) = 'service'" : ""}
    WHERE w.closed_at IS NOT NULL
      AND date(w.closed_at) >= ?
      AND date(w.closed_at) <= ?
      AND LOWER(COALESCE(w.source, '')) IN ('service', 'manual', 'maintenance')
    ORDER BY w.closed_at ASC
  `).all(from, to);

  const byDate = new Map();
  for (const r of rows) {
    const code = String(r.asset_code || "").toUpperCase();
    if (isExcludedAsset(code)) continue;
    const date = String(r.work_date || "");
    const description = describeWorkOrder(r);
    if (!date || !description) continue;

    let hours = Number(r.labor_hours || 0);
    if (!Number.isFinite(hours) || hours <= 0) {
      hours = woSourceCategory(r.source) === "Service" ? 4 + (hashSeed(date, code, description) % 3) : 2 + (hashSeed(date, code) % 3);
    }
    hours = Math.max(0.5, Math.min(8, Number(hours.toFixed(1))));

    if (!byDate.has(date)) byDate.set(date, []);
    byDate.get(date).push({
      asset_id: Number(r.asset_id),
      asset_code: code,
      description,
      hours,
      category: woSourceCategory(r.source),
    });
  }
  return byDate;
}

function loadSavedEntriesByDate(db, from, to, siteCode) {
  const out = new Map();
  if (!hasTable(db, "mechanic_labor_entries")) return out;

  const cols = ["id", "work_date", "technician_name", "hours", "asset_code", "reason", "labor_rate_per_hour"];
  for (const c of ["category", "time_started", "time_finished", "job_card_no", "smr"]) {
    if (hasColumn(db, "mechanic_labor_entries", c)) cols.push(c);
  }

  const rows = db.prepare(`
    SELECT ${cols.join(", ")}
    FROM mechanic_labor_entries
    WHERE work_date >= ? AND work_date <= ?
      AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
    ORDER BY work_date ASC, id ASC
  `).all(from, to, siteCode);

  for (const r of rows) {
    const code = String(r.asset_code || "").toUpperCase();
    if (isExcludedAsset(code)) continue;
    const date = String(r.work_date || "");
    if (!out.has(date)) out.set(date, []);
    out.get(date).push(r);
  }
  return out;
}

function makeRow({
  work_date,
  asset_code,
  asset_id,
  hours,
  category,
  reason,
  technician_name,
  time_started,
  time_finished,
  job_card_no,
  smr,
  source,
}) {
  return {
    work_date,
    date_display: formatDisplayDate(work_date),
    asset_code: String(asset_code || "").toUpperCase(),
    asset_id: asset_id != null ? Number(asset_id) : null,
    hours: fmtNum(hours, 2) ?? 0,
    category: String(category || ""),
    reason: String(reason || ""),
    technician_name: String(technician_name || ""),
    time_started: String(time_started || ""),
    time_finished: String(time_finished || ""),
    job_card_no: String(job_card_no || ""),
    smr: smr != null ? fmtNum(smr, 1) : null,
    source: String(source || "generated"),
  };
}

function scheduleTechRows(rows, work_date) {
  const byTech = new Map();
  for (const r of rows) {
    const tech = r.technician_name || "Workshop";
    if (!byTech.has(tech)) byTech.set(tech, []);
    byTech.get(tech).push(r);
  }

  const scheduled = [];
  for (const [tech, techRows] of byTech) {
    const startups = techRows.filter((r) => r.category === "Startup");
    const others = techRows.filter((r) => r.category !== "Startup");
    const startupOrder = shuffled(startups, `${work_date}|${tech}|startup`);
    const otherOrder = shuffled(others, `${work_date}|${tech}|other`);

    let cursor = techDayStartMinutes(work_date, tech);
    for (const r of [...startupOrder, ...otherOrder]) {
      const mins = Math.round(Number(r.hours || 0) * 60);
      const start = minutesToTime(cursor);
      const end = addMinutesToTime(start, mins);
      scheduled.push({
        ...r,
        technician_name: tech,
        time_started: start,
        time_finished: end,
      });
      cursor += mins;
    }
  }

  scheduled.sort((a, b) => {
    const d = String(a.work_date).localeCompare(String(b.work_date));
    if (d !== 0) return d;
    const t = String(a.technician_name).localeCompare(String(b.technician_name));
    if (t !== 0) return t;
    return timeToMinutes(a.time_started) - timeToMinutes(b.time_started);
  });
  return scheduled;
}

function pickGroupAsset(groupAssets, date, tech, salt) {
  if (!groupAssets?.length) return { asset_code: "ST01AM", asset_id: null };
  const order = shuffled(groupAssets, `${date}|${tech}|asset|${salt}`);
  const a = order[0];
  return { asset_code: a.asset_code, asset_id: a.id };
}

function pickServiceFromCatalog(catalog, groupAssets, date, tech, salt) {
  const codes = new Set(groupAssets.map((a) => a.asset_code));
  const eligible = catalog.filter((c) => codes.has(c.asset_code));
  if (!eligible.length) return null;
  const item = eligible[hashSeed(date, tech, salt, "svc") % eligible.length];
  return {
    asset_code: item.asset_code,
    asset_id: item.asset_id,
    description: `${item.service_name}hr service`,
    category: "Service",
  };
}

function planFillJobs(remaining, date, tech) {
  const jobs = [];
  let rem = remaining;
  let salt = 0;
  while (rem >= 0.5 && salt < 8) {
    const roll = hashSeed(date, tech, salt, "chunk") % 100;
    let category = "Service";
    if (roll >= 40 && roll < 72) category = "Maintenance";
    else if (roll >= 72) category = "Breakdown";

    const maxChunk = rem <= 2
      ? rem
      : Math.min(rem, 1.5 + (hashSeed(date, tech, salt, "hrs") % 5) * 0.5);
    const hrs = Math.max(0.5, Math.round(Math.min(maxChunk, rem) * 2) / 2);
    jobs.push({ hours: hrs, category, salt });
    rem = Number((rem - hrs).toFixed(2));
    salt += 1;
  }
  return jobs;
}

function generateDayRows({
  work_date,
  fleet,
  breakdowns,
  completedWork,
  serviceCatalog,
  dumptruckCodes,
  minTechHours,
  smrCache,
  usageTracker,
  db,
}) {
  const raw = [];
  const { groups, assetByCode } = fleet;
  const dayUsedKeys = new Set();

  for (const [group, assets] of Object.entries(groups)) {
    const order = shuffled(assets, `${work_date}|${group}|fleet`);
    for (const a of order) {
      const tech = techForAsset(group, a.asset_code, dumptruckCodes);
      raw.push(makeRow({
        work_date,
        asset_code: a.asset_code,
        asset_id: a.id,
        hours: STARTUP_HOURS,
        category: "Startup",
        reason: "Startup inspection",
        technician_name: tech,
        source: "startup",
      }));
    }
  }

  const addWorkRow = (item, tech, source) => {
    const key = `${item.asset_code}|${item.category}|${item.description}`;
    if (dayUsedKeys.has(key)) return;
    dayUsedKeys.add(key);
    usageTracker.record(tech, item.description);
    raw.push(makeRow({
      work_date,
      asset_code: item.asset_code,
      asset_id: item.asset_id,
      hours: item.hours,
      category: item.category,
      reason: item.description,
      technician_name: tech,
      source,
    }));
  };

  for (const item of completedWork || []) {
    const group = classifyAsset({ asset_code: item.asset_code });
    const tech = group ? techForAsset(group, item.asset_code, dumptruckCodes) : "Joaquim";
    addWorkRow(item, tech, "work_order");
  }

  for (const bd of breakdowns || []) {
    const group = classifyAsset(bd);
    const tech = group ? techForAsset(group, bd.asset_code, dumptruckCodes) : "Joaquim";
    addWorkRow(bd, tech, "breakdown");
  }

  const hoursByTech = new Map();
  for (const r of raw) {
    const t = r.technician_name;
    hoursByTech.set(t, (hoursByTech.get(t) || 0) + Number(r.hours || 0));
  }

  const allTechs = [
    ...TECH_BY_GROUP.dumptruck,
    ...TECH_BY_GROUP.loader,
    ...TECH_BY_GROUP.excavator,
    ...TECH_BY_GROUP.crusher_screen,
  ];

  for (const tech of allTechs) {
    let remaining = minTechHours - (hoursByTech.get(tech) || 0);
    if (remaining < 0.5) continue;

    const groupKey = Object.entries(TECH_BY_GROUP).find(([, list]) => list.includes(tech))?.[0];
    const groupAssets = groupKey ? groups[groupKey] : groups.loader;
    const fillJobs = planFillJobs(remaining, work_date, tech);

    for (const job of fillJobs) {
      let asset = pickGroupAsset(groupAssets, work_date, tech, job.salt);
      let description = null;
      let category = job.category;

      if (job.category === "Service" && hashSeed(work_date, tech, job.salt, "svcMix") % 100 < 45) {
        const fromCatalog = pickServiceFromCatalog(serviceCatalog, groupAssets, work_date, tech, job.salt);
        if (fromCatalog) {
          const recent = new Set(usageTracker.techRecent.get(tech) || []);
          if (!recent.has(fromCatalog.description)) {
            asset = { asset_code: fromCatalog.asset_code, asset_id: fromCatalog.asset_id };
            description = fromCatalog.description;
            category = "Service";
          }
        }
      }

      if (!description) {
        description = usageTracker.pick(
          tech,
          work_date,
          category,
          JOB_POOLS[category] || JOB_POOLS.Maintenance,
          job.salt,
        );
      }

      const key = `${asset.asset_code}|${category}|${description}`;
      if (dayUsedKeys.has(key)) {
        description = usageTracker.pick(tech, work_date, category, JOB_POOLS[category] || JOB_POOLS.Maintenance, `${job.salt}|alt`);
      }
      const finalKey = `${asset.asset_code}|${category}|${description}`;
      if (dayUsedKeys.has(finalKey)) continue;
      dayUsedKeys.add(finalKey);

      raw.push(makeRow({
        work_date,
        asset_code: asset.asset_code,
        asset_id: asset.asset_id,
        hours: job.hours,
        category,
        reason: description,
        technician_name: tech,
        source: "synthetic_fill",
      }));
    }
  }

  for (const tech of allTechs) {
    const total = raw
      .filter((r) => r.technician_name === tech)
      .reduce((s, r) => s + Number(r.hours || 0), 0);
    const short = Number((minTechHours - total).toFixed(2));
    if (short < 0.5) continue;

    const groupKey = Object.entries(TECH_BY_GROUP).find(([, list]) => list.includes(tech))?.[0];
    const groupAssets = groupKey ? groups[groupKey] : groups.loader;
    const asset = pickGroupAsset(groupAssets, work_date, tech, "topup");
    const category = hashSeed(work_date, tech, "topup") % 2 === 0 ? "Maintenance" : "Service";
    const description = usageTracker.pick(
      tech,
      work_date,
      category,
      JOB_POOLS[category],
      "topup",
    );
    raw.push(makeRow({
      work_date,
      asset_code: asset.asset_code,
      asset_id: asset.asset_id,
      hours: Math.max(0.5, Math.round(short * 2) / 2),
      category,
      reason: description,
      technician_name: tech,
      source: "synthetic_topup",
    }));
  }

  for (const r of raw) {
    const asset = assetByCode.get(r.asset_code);
    const aid = r.asset_id || asset?.id;
    r.smr = resolveSmr(db, aid, work_date, smrCache);
  }

  return scheduleTechRows(raw, work_date);
}

function mapSavedToRows(savedRows, db, smrCache) {
  return savedRows.map((r) => {
    const asset = db.prepare(`
      SELECT id FROM assets WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?)) LIMIT 1
    `).get(String(r.asset_code || ""));
    const smr = r.smr != null && Number.isFinite(Number(r.smr))
      ? fmtNum(r.smr, 1)
      : resolveSmr(db, asset?.id, r.work_date, smrCache);
    return makeRow({
      work_date: r.work_date,
      asset_code: r.asset_code,
      asset_id: asset?.id,
      hours: r.hours,
      category: r.category || inferCategory(r.reason),
      reason: r.reason,
      technician_name: r.technician_name,
      time_started: r.time_started || "",
      time_finished: r.time_finished || "",
      job_card_no: r.job_card_no || "",
      smr,
      source: "saved",
    });
  });
}

/**
 * @param {import('better-sqlite3').Database} db
 * @param {{
 *   from: string,
 *   to: string,
 *   siteCode?: string,
 * }} opts
 */
export function generateMechanicsTimesheet(db, opts = {}) {
  const from = String(opts.from || "").trim();
  const to = String(opts.to || "").trim();
  const siteCode = String(opts.siteCode || "main").trim().toLowerCase() || "main";

  if (!isDate(from) || !isDate(to)) {
    throw new Error("from and to (YYYY-MM-DD) are required");
  }
  if (from > to) throw new Error("from must be on or before to");

  const savedByDate = loadSavedEntriesByDate(db, from, to, siteCode);
  const smrCache = new Map();

  const allRows = [];
  const meta = {
    from,
    to,
    site_code: siteCode,
    excluded_assets: [...EXCLUDED_ASSET_CODES],
    days_from_saved: 0,
    days_synthetic: 0,
    row_count: 0,
  };

  for (const work_date of enumerateDates(from, to)) {
    const saved = savedByDate.get(work_date) || [];
    if (!saved.length) continue;
    allRows.push(...mapSavedToRows(saved, db, smrCache));
    meta.days_from_saved += 1;
  }

  meta.row_count = allRows.length;
  return { rows: allRows, meta };
}

/** Map rows to costing / Excel column names. */
export function mechanicsTimesheetToExportRows(rows) {
  return rows.map((r) => ({
    Date: r.date_display || formatDisplayDate(r.work_date),
    "Plant no": r.asset_code,
    "Work Hours": r.hours,
    Category: r.category,
    "Description Of Work Carried Out": r.reason,
    "Time Started": r.time_started,
    "Time finished": r.time_finished,
    Technician: r.technician_name,
    "Job Card No": r.job_card_no,
    SMR: r.smr,
    work_date: r.work_date,
    source: r.source,
  }));
}
