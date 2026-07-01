// Generate mechanics labour timesheet rows (startup inspections + breakdowns + fill).

const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

const STARTUP_HOURS = 0.5;
const DEFAULT_MIN_TECH_HOURS = 6;
const DEFAULT_SYNTHETIC_THROUGH = "2026-03-31";

const TECH_BY_GROUP = {
  dumptruck: ["Sergio", "Arnold"],
  loader: ["Ronnie"],
  excavator: ["Charles"],
  crusher_screen: ["Moses"],
};

const FILLER_DESCRIPTIONS = [
  "Hydraulic hose replacement",
  "Bearing inspection and grease",
  "Brake system adjustment",
  "500hr service",
  "Remove and inspect diff",
  "Track tension adjustment",
  "Cooling system repair",
  "Electrical fault finding",
  "Welding on chassis crack",
  "Replace worn pins and bushes",
];

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
  const cat = String(asset.category || "").toLowerCase();
  const name = String(asset.asset_name || "").toLowerCase();
  const hay = `${cat} ${name}`;

  if (/adt|dump|haul\s*truck|articulated/.test(hay) || /^A\d+AM$/.test(code)) return "dumptruck";
  if (/excavator|digger|shovel/.test(hay) || /^E\d+AM$/.test(code)) return "excavator";
  if (/^F\d+AM$/.test(code) || (/\bloader\b/.test(hay) && !/forklift|tlb|backhoe|block/.test(hay))) {
    return "loader";
  }
  if (/crusher|screen|trommel/.test(hay) || /^CR\d+AM$/.test(code) || code === "FIN694") {
    return "crusher_screen";
  }
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

function cleanBreakdownDesc(desc) {
  const s = String(desc || "")
    .replace(/^DOWN\s*[—\-]\s*/i, "")
    .replace(/^DOWN\s+/i, "")
    .trim();
  return s || "Mechanical breakdown";
}

function inferCategory(reason) {
  const r = String(reason || "").toLowerCase();
  if (r.includes("startup")) return "Startup";
  if (r.includes("inspection") && r.includes("startup")) return "Startup";
  return "Breakdown";
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
    const group = classifyAsset(a);
    if (!group) continue;
    const code = String(a.asset_code || "").trim().toUpperCase();
    assetByCode.set(code, a);
    groups[group].push({ ...a, asset_code: code });
  }

  return { groups, assetByCode };
}

function loadBreakdownsByDate(db, from, to) {
  if (!hasTable(db, "breakdowns") || !hasTable(db, "assets")) return new Map();
  const hasWo = hasTable(db, "work_orders");
  const hasDowntimeLogs = hasTable(db, "breakdown_downtime_logs");

  const rows = db.prepare(`
    SELECT
      b.id,
      b.breakdown_date,
      b.description,
      b.asset_id,
      a.asset_code,
      a.asset_name,
      a.category,
      ${hasWo ? "w.labor_hours" : "NULL AS labor_hours"}
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
    const date = String(r.breakdown_date || "");
    if (!date) continue;
    const bdKey = `${Number(r.id)}|${date}`;
    const downtime = downtimeByBdDate.get(bdKey) || 0;
    let hours = Number(r.labor_hours || 0);
    if (!Number.isFinite(hours) || hours <= 0) {
      hours = downtime > 0 ? Math.min(8, Math.max(2, downtime * 0.6)) : 4;
    }
    hours = Math.max(1, Math.min(8, Number(hours.toFixed(1))));

    if (!byDate.has(date)) byDate.set(date, []);
    byDate.get(date).push({
      asset_id: Number(r.asset_id),
      asset_code: String(r.asset_code || "").toUpperCase(),
      asset_name: String(r.asset_name || ""),
      category: String(r.category || ""),
      description: cleanBreakdownDesc(r.description),
      hours,
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

function scheduleTechRows(rows) {
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
    startups.sort((a, b) => String(a.asset_code).localeCompare(String(b.asset_code)));
    others.sort((a, b) => Number(b.hours) - Number(a.hours));

    let cursor = 6 * 60;
    for (const r of [...startups, ...others]) {
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

function pickFillerDescription(date, tech, idx) {
  const h = hashSeed(date, tech, idx, "filler");
  return FILLER_DESCRIPTIONS[h % FILLER_DESCRIPTIONS.length];
}

function pickFillerAsset(groupAssets, date, tech, idx) {
  if (!groupAssets?.length) return { asset_code: "ST01AM", asset_id: null };
  const h = hashSeed(date, tech, idx, "asset");
  const a = groupAssets[h % groupAssets.length];
  return { asset_code: a.asset_code, asset_id: a.id };
}

function generateDayRows({
  work_date,
  fleet,
  breakdowns,
  dumptruckCodes,
  minTechHours,
  smrCache,
  db,
}) {
  const raw = [];
  const { groups, assetByCode } = fleet;

  for (const [group, assets] of Object.entries(groups)) {
    for (const a of assets) {
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

  const usedBreakdownKeys = new Set();
  for (const bd of breakdowns || []) {
    const group = classifyAsset(bd);
    const tech = group ? techForAsset(group, bd.asset_code, dumptruckCodes) : "Joaquim";
    const key = `${bd.asset_code}|${bd.description}`;
    if (usedBreakdownKeys.has(key)) continue;
    usedBreakdownKeys.add(key);
    raw.push(makeRow({
      work_date,
      asset_code: bd.asset_code,
      asset_id: bd.asset_id,
      hours: bd.hours,
      category: "Breakdown",
      reason: bd.description,
      technician_name: tech,
      source: "breakdown",
    }));
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
    const groupKey = Object.entries(TECH_BY_GROUP).find(([, list]) => list.includes(tech))?.[0];
    const groupAssets = groupKey ? groups[groupKey] : groups.loader;
    if (remaining >= 0.5) {
      const hrs = Math.max(0.5, Math.round(remaining * 2) / 2);
      const pick = pickFillerAsset(groupAssets, work_date, tech, 0);
      raw.push(makeRow({
        work_date,
        asset_code: pick.asset_code,
        asset_id: pick.asset_id,
        hours: hrs,
        category: "Breakdown",
        reason: pickFillerDescription(work_date, tech, 0),
        technician_name: tech,
        source: "synthetic_fill",
      }));
    }
  }

  for (const r of raw) {
    const asset = assetByCode.get(r.asset_code);
    const aid = r.asset_id || asset?.id;
    r.smr = resolveSmr(db, aid, work_date, smrCache);
  }

  return scheduleTechRows(raw);
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
 *   syntheticThrough?: string,
 *   siteCode?: string,
 *   minTechHours?: number,
 * }} opts
 */
export function generateMechanicsTimesheet(db, opts = {}) {
  const from = String(opts.from || "").trim();
  const to = String(opts.to || "").trim();
  const syntheticThrough = String(opts.syntheticThrough || DEFAULT_SYNTHETIC_THROUGH).trim();
  const siteCode = String(opts.siteCode || "main").trim().toLowerCase() || "main";
  const minTechHours = Number.isFinite(Number(opts.minTechHours)) && Number(opts.minTechHours) > 0
    ? Number(opts.minTechHours)
    : DEFAULT_MIN_TECH_HOURS;

  if (!isDate(from) || !isDate(to)) {
    throw new Error("from and to (YYYY-MM-DD) are required");
  }
  if (from > to) throw new Error("from must be on or before to");

  const fleet = loadStartupFleet(db);
  const dumptruckCodes = fleet.groups.dumptruck.map((a) => a.asset_code);
  const breakdownsByDate = loadBreakdownsByDate(db, from, to);
  const savedByDate = loadSavedEntriesByDate(db, from, to, siteCode);
  const smrCache = new Map();

  const allRows = [];
  const meta = {
    from,
    to,
    synthetic_through: syntheticThrough,
    site_code: siteCode,
    min_tech_hours: minTechHours,
    days_generated: 0,
    days_from_saved: 0,
    days_synthetic: 0,
    startup_fleet: {
      dumptrucks: fleet.groups.dumptruck.length,
      loaders: fleet.groups.loader.length,
      excavators: fleet.groups.excavator.length,
      crushers_screens: fleet.groups.crusher_screen.length,
    },
  };

  for (const work_date of enumerateDates(from, to)) {
    const useSynthetic = work_date <= syntheticThrough;
    const saved = savedByDate.get(work_date) || [];

    if (!useSynthetic && saved.length) {
      allRows.push(...mapSavedToRows(saved, db, smrCache));
      meta.days_from_saved += 1;
      continue;
    }

    const dayRows = generateDayRows({
      work_date,
      fleet,
      breakdowns: breakdownsByDate.get(work_date) || [],
      dumptruckCodes,
      minTechHours,
      smrCache,
      db,
    });
    allRows.push(...dayRows);
    meta.days_synthetic += 1;
    meta.days_generated += 1;
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
