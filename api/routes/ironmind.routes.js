import { generateIronmindReport, getIronmindHistory, getLatestIronmindReport, getIronmindSettings, setIronmindSettings } from "../utils/ironmind.js";
import { db } from "../db/client.js";
import { buildPdfBuffer, sectionTitle, table } from "../utils/pdfGenerator.js";

function toBool(v) {
  const s = String(v || "").trim().toLowerCase();
  return s === "1" || s === "true" || s === "yes" || s === "y";
}

const RSG_SERVICE_PROFILES = [
  {
    key: "cat-350-6000",
    make: "CAT",
    modelContains: ["350"],
    service_hours: 6000,
    title: "CAT 350 - 6000hr service",
    tasks: [
      "Drain and replace engine oil, final drives, and swing motor oils",
      "Inspect hydraulic tank condition and top-up/replace per OEM procedure",
      "Replace engine oil, fuel, hydraulic return, pilot, and air filters",
      "Inspect contamination indicators and perform leak checks under load",
      "Record all meter readings and sign off service completion",
    ],
    oils: [
      { name: "Hydraulic oil", qty: 210, unit: "L", part_hint: "hydraulic oil" },
      { name: "Engine oil", qty: 40, unit: "L", part_hint: "engine oil" },
      { name: "SAE50 final drives and swing motors", qty: 80, unit: "L", part_hint: "sae50" },
    ],
    filters: [
      { name: "Engine oil filter", qty: 1, unit: "ea", part_hint: "oil filter" },
      { name: "Fuel filter primary", qty: 1, unit: "ea", part_hint: "fuel filter" },
      { name: "Fuel filter secondary", qty: 1, unit: "ea", part_hint: "fuel filter" },
      { name: "Hydraulic return filter", qty: 1, unit: "ea", part_hint: "hydraulic filter" },
      { name: "Pilot filter", qty: 1, unit: "ea", part_hint: "pilot filter" },
      { name: "Air filter outer", qty: 1, unit: "ea", part_hint: "air filter" },
      { name: "Air filter inner", qty: 1, unit: "ea", part_hint: "air filter" },
    ],
  },
];

export default async function ironmindRoutes(app) {
  function isDate(v) {
    return /^\d{4}-\d{2}-\d{2}$/.test(String(v || "").trim());
  }
  function todayYmd() {
    return new Date().toISOString().slice(0, 10);
  }
  function parseQuestionDates(question, fallbackDate) {
    const q = String(question || "");
    const matches = q.match(/\b\d{4}-\d{2}-\d{2}\b/g) || [];
    if (matches.length >= 2) return { start: matches[0], end: matches[1] };
    if (matches.length === 1) return { start: matches[0], end: matches[0] };
    const ym = q.match(/\b(\d{4})-(\d{2})\b/);
    if (ym) {
      const y = Number(ym[1]);
      const m = Number(ym[2]);
      if (y >= 2000 && m >= 1 && m <= 12) {
        const start = `${String(y).padStart(4, "0")}-${String(m).padStart(2, "0")}-01`;
        const endDate = new Date(Date.UTC(y, m, 0));
        const end = `${endDate.getUTCFullYear()}-${String(endDate.getUTCMonth() + 1).padStart(2, "0")}-${String(endDate.getUTCDate()).padStart(2, "0")}`;
        return { start, end };
      }
    }
    const monthMap = {
      january: 1, jan: 1,
      february: 2, feb: 2,
      march: 3, mar: 3,
      april: 4, apr: 4,
      may: 5,
      june: 6, jun: 6,
      july: 7, jul: 7,
      august: 8, aug: 8,
      september: 9, sep: 9, sept: 9,
      october: 10, oct: 10,
      november: 11, nov: 11,
      december: 12, dec: 12,
    };
    const lower = q.toLowerCase();
    const monthKey = Object.keys(monthMap).find((k) => new RegExp(`\\b${k}\\b`, "i").test(lower));
    if (monthKey) {
      const fallbackYear = Number(String(fallbackDate).slice(0, 4)) || new Date().getUTCFullYear();
      const yearMatch = lower.match(/\b(20\d{2})\b/);
      const year = yearMatch ? Number(yearMatch[1]) : fallbackYear;
      const m = monthMap[monthKey];
      const start = `${String(year).padStart(4, "0")}-${String(m).padStart(2, "0")}-01`;
      const endDate = new Date(Date.UTC(year, m, 0));
      const end = `${endDate.getUTCFullYear()}-${String(endDate.getUTCMonth() + 1).padStart(2, "0")}-${String(endDate.getUTCDate()).padStart(2, "0")}`;
      return { start, end };
    }
    return { start: fallbackDate, end: fallbackDate };
  }
  function parseAssetCode(question) {
    const q = String(question || "").toUpperCase();
    const m = q.match(/\b[A-Z0-9]{2,}[A-Z][0-9A-Z-]*\b/g) || [];
    const deny = new Set(["PLEASE", "DOWNTIME", "SELECTED", "TIME", "FROM", "TO", "AND", "THE", "FOR", "FUEL", "USAGE", "RECURRING", "FAILURES", "PM", "OVERDUE", "RISK"]);
    return (m.find((x) => !deny.has(x)) || "").trim();
  }
  function hasColumn(table, col) {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name) === col);
  }
  db.prepare(`
    CREATE TABLE IF NOT EXISTS rsg_service_profiles (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      profile_key TEXT NOT NULL UNIQUE,
      make TEXT,
      model_match TEXT,
      service_hours INTEGER NOT NULL,
      title TEXT NOT NULL,
      tasks_json TEXT NOT NULL DEFAULT '[]',
      checks_json TEXT NOT NULL DEFAULT '[]',
      post_service_checks_json TEXT NOT NULL DEFAULT '[]',
      safety_json TEXT NOT NULL DEFAULT '[]',
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS rsg_service_profile_items (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      profile_id INTEGER NOT NULL,
      item_type TEXT NOT NULL, -- oil | filter
      item_name TEXT NOT NULL,
      qty REAL NOT NULL,
      unit TEXT NOT NULL DEFAULT 'L',
      part_hint TEXT,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY(profile_id) REFERENCES rsg_service_profiles(id) ON DELETE CASCADE
    )
  `).run();
  function getAiConfig() {
    const openaiKey = process.env.OPENAI_API_KEY;
    const openaiModel = process.env.OPENAI_MODEL || "gpt-4o-mini";
    if (openaiKey) return { provider: "openai", apiKey: openaiKey, model: openaiModel };
    return { provider: null };
  }
  function normalizeText(v) {
    return String(v || "").trim().toUpperCase();
  }
  function parseJsonArray(v) {
    try {
      const parsed = JSON.parse(String(v || "[]"));
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }
  function parseModelMatchTokens(v) {
    return String(v || "")
      .split(",")
      .map((x) => normalizeText(x))
      .filter(Boolean);
  }
  function getDbRsgProfiles() {
    const profiles = db.prepare(`
      SELECT
        id, profile_key, make, model_match, service_hours, title,
        tasks_json, checks_json, post_service_checks_json, safety_json, active
      FROM rsg_service_profiles
      WHERE active = 1
      ORDER BY service_hours ASC, id ASC
    `).all();
    if (!profiles.length) return [];
    const items = db.prepare(`
      SELECT
        profile_id, item_type, item_name, qty, unit, part_hint, sort_order
      FROM rsg_service_profile_items
      ORDER BY profile_id ASC, sort_order ASC, id ASC
    `).all();
    const byProfile = new Map();
    for (const it of items) {
      const pid = Number(it.profile_id || 0);
      if (!byProfile.has(pid)) byProfile.set(pid, []);
      byProfile.get(pid).push(it);
    }
    return profiles.map((p) => {
      const pItems = byProfile.get(Number(p.id || 0)) || [];
      return {
        key: String(p.profile_key || "").trim(),
        make: String(p.make || "").trim(),
        modelContains: parseModelMatchTokens(p.model_match),
        service_hours: Number(p.service_hours || 0),
        title: String(p.title || "").trim(),
        tasks: parseJsonArray(p.tasks_json).map((x) => String(x || "")).filter(Boolean),
        checks: parseJsonArray(p.checks_json).map((x) => String(x || "")).filter(Boolean),
        post_service_checks: parseJsonArray(p.post_service_checks_json).map((x) => String(x || "")).filter(Boolean),
        safety: parseJsonArray(p.safety_json).map((x) => String(x || "")).filter(Boolean),
        oils: pItems
          .filter((x) => String(x.item_type || "") === "oil")
          .map((x) => ({
            name: String(x.item_name || ""),
            qty: Number(x.qty || 0),
            unit: String(x.unit || "L"),
            part_hint: String(x.part_hint || x.item_name || ""),
          }))
          .filter((x) => x.name && Number.isFinite(x.qty) && x.qty > 0),
        filters: pItems
          .filter((x) => String(x.item_type || "") === "filter")
          .map((x) => ({
            name: String(x.item_name || ""),
            qty: Number(x.qty || 0),
            unit: String(x.unit || "ea"),
            part_hint: String(x.part_hint || x.item_name || ""),
          }))
          .filter((x) => x.name && Number.isFinite(x.qty) && x.qty > 0),
      };
    }).filter((p) => p.key && p.service_hours > 0 && p.title);
  }
  function pickRsgServiceProfile({ assetCode, equipmentName, serviceHours }) {
    const dbProfiles = getDbRsgProfiles();
    const sourceProfiles = dbProfiles.length ? dbProfiles : RSG_SERVICE_PROFILES;
    const code = normalizeText(assetCode);
    const eq = normalizeText(equipmentName);
    const hours = Number(serviceHours || 0);
    return sourceProfiles.find((p) => {
      if (Number(p.service_hours || 0) !== hours) return false;
      const makeOk = normalizeText(p.make) ? (code.includes(normalizeText(p.make)) || eq.includes(normalizeText(p.make))) : true;
      if (!makeOk) return false;
      const modelTokens = Array.isArray(p.modelContains) ? p.modelContains.map(normalizeText).filter(Boolean) : [];
      if (!modelTokens.length) return true;
      return modelTokens.some((t) => code.includes(t) || eq.includes(t));
    }) || null;
  }

  // RSG profile maintenance (document-driven master data)
  app.get("/rsg/profiles", async (_req, reply) => {
    const rows = getDbRsgProfiles();
    return reply.send({ ok: true, rows });
  });
  function upsertRsgProfile(b = {}) {
    const profile_key = String(b.profile_key || "").trim().toLowerCase();
    const make = String(b.make || "").trim();
    const model_match = String(b.model_match || "").trim();
    const service_hours = Math.max(1, Number(b.service_hours || 0));
    const title = String(b.title || "").trim();
    const tasks = Array.isArray(b.tasks) ? b.tasks.map((x) => String(x || "").trim()).filter(Boolean) : [];
    const checks = Array.isArray(b.checks) ? b.checks.map((x) => String(x || "").trim()).filter(Boolean) : [];
    const post_service_checks = Array.isArray(b.post_service_checks) ? b.post_service_checks.map((x) => String(x || "").trim()).filter(Boolean) : [];
    const safety = Array.isArray(b.safety) ? b.safety.map((x) => String(x || "").trim()).filter(Boolean) : [];
    const oils = Array.isArray(b.oils) ? b.oils : [];
    const filters = Array.isArray(b.filters) ? b.filters : [];
    if (!profile_key || !service_hours || !title) {
      const e = new Error("profile_key, service_hours, and title are required");
      e.statusCode = 400;
      throw e;
    }
    const tx = db.transaction(() => {
      const ins = db.prepare(`
        INSERT INTO rsg_service_profiles (
          profile_key, make, model_match, service_hours, title,
          tasks_json, checks_json, post_service_checks_json, safety_json, active, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, datetime('now'))
        ON CONFLICT(profile_key) DO UPDATE SET
          make = excluded.make,
          model_match = excluded.model_match,
          service_hours = excluded.service_hours,
          title = excluded.title,
          tasks_json = excluded.tasks_json,
          checks_json = excluded.checks_json,
          post_service_checks_json = excluded.post_service_checks_json,
          safety_json = excluded.safety_json,
          active = 1,
          updated_at = datetime('now')
      `).run(
        profile_key,
        make || null,
        model_match || null,
        service_hours,
        title,
        JSON.stringify(tasks),
        JSON.stringify(checks),
        JSON.stringify(post_service_checks),
        JSON.stringify(safety)
      );
      const pid =
        Number(ins.lastInsertRowid || 0) ||
        Number(db.prepare(`SELECT id FROM rsg_service_profiles WHERE profile_key = ?`).get(profile_key)?.id || 0);
      db.prepare(`DELETE FROM rsg_service_profile_items WHERE profile_id = ?`).run(pid);
      const insertItem = db.prepare(`
        INSERT INTO rsg_service_profile_items (
          profile_id, item_type, item_name, qty, unit, part_hint, sort_order, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `);
      let sort = 1;
      for (const o of oils) {
        const name = String(o?.name || "").trim();
        const qty = Number(o?.qty || 0);
        if (!name || !(qty > 0)) continue;
        insertItem.run(pid, "oil", name, qty, String(o?.unit || "L"), String(o?.part_hint || name), sort++);
      }
      for (const f of filters) {
        const name = String(f?.name || "").trim();
        const qty = Number(f?.qty || 0);
        if (!name || !(qty > 0)) continue;
        insertItem.run(pid, "filter", name, qty, String(f?.unit || "ea"), String(f?.part_hint || name), sort++);
      }
      return { id: pid, profile_key };
    });
    return tx();
  }
  app.post("/rsg/profiles", async (req, reply) => {
    try {
      const result = upsertRsgProfile(req.body || {});
      return reply.send({ ok: true, id: result.id, profile_key: result.profile_key });
    } catch (err) {
      req.log.error(err);
      return reply.code(Number(err?.statusCode || 500)).send({ ok: false, error: err.message || String(err) });
    }
  });
  app.post("/rsg/profiles/import", async (req, reply) => {
    try {
      const rows = Array.isArray(req.body?.profiles) ? req.body.profiles : [];
      if (!rows.length) return reply.code(400).send({ ok: false, error: "profiles[] is required" });
      let upserted = 0;
      for (const p of rows) {
        upsertRsgProfile(p || {});
        upserted += 1;
      }
      return reply.send({ ok: true, requested: rows.length, upserted });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
  function getPartOnHandById(partId) {
    const id = Number(partId || 0);
    if (!id) return 0;
    const row = db.prepare(`
      SELECT COALESCE(SUM(quantity), 0) AS on_hand
      FROM stock_movements
      WHERE part_id = ?
    `).get(id);
    return Number(row?.on_hand || 0);
  }
  function findPartForHint(hint) {
    const h = String(hint || "").trim();
    if (!h) return null;
    const row = db.prepare(`
      SELECT id, part_code, part_name
      FROM parts
      WHERE LOWER(COALESCE(part_name, '')) LIKE LOWER(?)
         OR LOWER(COALESCE(part_code, '')) LIKE LOWER(?)
      ORDER BY
        CASE
          WHEN LOWER(COALESCE(part_name, '')) = LOWER(?) THEN 0
          WHEN LOWER(COALESCE(part_code, '')) = LOWER(?) THEN 1
          ELSE 2
        END,
        id ASC
      LIMIT 1
    `).get(`%${h}%`, `%${h}%`, h, h);
    return row || null;
  }
  function buildRsgReadiness({ oils = [], filters = [] }) {
    const rows = [];
    for (const o of oils) {
      const required = Number(o?.qty || 0);
      if (!(required > 0)) continue;
      const part = findPartForHint(o?.part_hint || o?.name || "");
      const onHand = part ? getPartOnHandById(part.id) : 0;
      rows.push({
        category: "oil",
        name: String(o?.name || ""),
        unit: String(o?.unit || "L"),
        required_qty: Number(required.toFixed(2)),
        on_hand_qty: Number(onHand.toFixed(2)),
        sufficient: onHand >= required,
        shortage_qty: Number(Math.max(0, required - onHand).toFixed(2)),
        part_code: part?.part_code || null,
        part_name: part?.part_name || null,
      });
    }
    for (const f of filters) {
      const required = Number(f?.qty || 0);
      if (!(required > 0)) continue;
      const part = findPartForHint(f?.part_hint || f?.name || "");
      const onHand = part ? getPartOnHandById(part.id) : 0;
      rows.push({
        category: "filter",
        name: String(f?.name || ""),
        unit: String(f?.unit || "ea"),
        required_qty: Number(required.toFixed(2)),
        on_hand_qty: Number(onHand.toFixed(2)),
        sufficient: onHand >= required,
        shortage_qty: Number(Math.max(0, required - onHand).toFixed(2)),
        part_code: part?.part_code || null,
        part_name: part?.part_name || null,
      });
    }
    const insufficient = rows.filter((r) => !r.sufficient);
    return {
      items: rows,
      all_sufficient: insufficient.length === 0,
      insufficient_count: insufficient.length,
    };
  }
  function getAssetOilProfile(assetId, limit = 6) {
    const id = Number(assetId || 0);
    if (!id) return [];
    const hasOilLogs = db.prepare(`
      SELECT 1 AS ok
      FROM sqlite_master
      WHERE type='table' AND name='oil_logs'
      LIMIT 1
    `).get();
    if (!hasOilLogs) return [];

    const hasMappings = db.prepare(`
      SELECT 1 AS ok
      FROM sqlite_master
      WHERE type='table' AND name='lube_type_mappings'
      LIMIT 1
    `).get();
    const hasParts = db.prepare(`
      SELECT 1 AS ok
      FROM sqlite_master
      WHERE type='table' AND name='parts'
      LIMIT 1
    `).get();

    const rows = db.prepare(`
      SELECT
        COALESCE(NULLIF(TRIM(o.oil_type), ''), 'UNSPECIFIED') AS oil_key,
        COUNT(*) AS fills,
        COALESCE(SUM(o.quantity), 0) AS qty_total,
        MAX(o.log_date) AS last_used
      FROM oil_logs o
      WHERE o.asset_id = ?
      GROUP BY COALESCE(NULLIF(TRIM(o.oil_type), ''), 'UNSPECIFIED')
      ORDER BY fills DESC, qty_total DESC, last_used DESC
      LIMIT ?
    `).all(id, Number(limit));

    if (!rows.length) return [];

    let mappingByOil = new Map();
    if (hasMappings && hasParts) {
      const mapped = db.prepare(`
        SELECT
          LOWER(TRIM(m.oil_key)) AS oil_key_lc,
          TRIM(m.part_code) AS part_code,
          TRIM(p.part_name) AS part_name
        FROM lube_type_mappings m
        LEFT JOIN parts p ON UPPER(TRIM(p.part_code)) = UPPER(TRIM(m.part_code))
      `).all();
      mappingByOil = new Map(
        mapped
          .filter((r) => String(r.oil_key_lc || "").trim() !== "")
          .map((r) => [String(r.oil_key_lc).trim(), { part_code: r.part_code || "", part_name: r.part_name || "" }])
      );
    }

    return rows.map((r) => {
      const oilKey = String(r.oil_key || "").trim();
      const mapped = mappingByOil.get(oilKey.toLowerCase()) || null;
      const label = mapped?.part_name
        ? `${oilKey} (${mapped.part_name}${mapped.part_code ? ` - ${mapped.part_code}` : ""})`
        : oilKey;
      return {
        name: label,
        qty: Number(Number(r.qty_total || 0).toFixed(2)),
        unit: "L",
      };
    }).filter((o) => o.name && Number.isFinite(o.qty) && o.qty > 0);
  }

  async function tryGenerateRsgPlanWithAi({ equipmentLabel, serviceHours, preferredOils = [] }) {
    const cfg = getAiConfig();
    if (!cfg.provider) return null;
    const system = [
      "You are a heavy-equipment maintenance planner.",
      "Return strict JSON only.",
      "Schema: {service_title:string,tasks:[string],oils:[{name:string,qty:number,unit:string}],checks:[string],post_service_checks:[string],safety:[string]}",
      "Use practical values. If exact OEM value is unknown, provide conservative estimate and mention 'verify with OEM manual' in checks.",
    ].join(" ");
    const oilHint = preferredOils.length
      ? `Use these site oils and quantities as the default unless clearly unsafe: ${preferredOils.map((o) => `${o.name} ${o.qty}${o.unit || "L"}`).join("; ")}.`
      : "If exact oil grades are uncertain, keep conservative values and tell user to verify with OEM manual.";
    const user = `Generate a ${serviceHours} hour Recommended Service Guide for ${equipmentLabel}. Include key tasks, oil/lube quantities, checks before release, and safety steps. ${oilHint}`;
    const res = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${cfg.apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: cfg.model,
        temperature: 0.2,
        messages: [
          { role: "system", content: system },
          { role: "user", content: user },
        ],
      }),
    });
    if (!res.ok) return null;
    const data = await res.json();
    const content = String(data?.choices?.[0]?.message?.content || "").trim();
    const start = content.indexOf("{");
    const end = content.lastIndexOf("}");
    if (start < 0 || end < start) return null;
    try {
      const parsed = JSON.parse(content.slice(start, end + 1));
      return parsed;
    } catch {
      return null;
    }
  }
  function buildFallbackRsgPlan({ equipmentLabel, serviceHours }) {
    return {
      service_title: `${equipmentLabel} - ${serviceHours}hr service`,
      tasks: [
        "Drain and replace engine oil and filters",
        "Replace hydraulic return/pilot filters and inspect suction strainers",
        "Replace fuel filters and water separator element",
        "Inspect air intake system and replace air filters if restricted",
        "Inspect undercarriage, pins/bushings, and grease all lubrication points",
        "Inspect cooling pack; clean cores and verify fan operation",
      ],
      oils: [
        { name: "Engine oil", qty: 32, unit: "L" },
        { name: "Final drive oil (each side)", qty: 8, unit: "L" },
        { name: "Swing drive oil", qty: 6, unit: "L" },
        { name: "Grease", qty: 3, unit: "kg" },
      ],
      checks: [
        "Verify exact capacities against OEM manual before fill",
        "Inspect for leaks at all changed filters and drain points",
        "Record hourmeter and service completion in maintenance history",
      ],
      post_service_checks: [
        "Warm-up run for 15-20 minutes and re-check fluid levels",
        "Check fault codes and confirm no active alarms",
        "Function-test boom/arm/bucket/swing/travel and verify operating pressures",
      ],
      safety: [
        "LOTO and isolate machine before service",
        "Use spill kits and approved waste-oil disposal process",
        "Use calibrated torque specs for critical fasteners",
      ],
      filters: [],
    };
  }
  function normalizeRsgPlan(raw, equipmentLabel, serviceHours) {
    const fallback = buildFallbackRsgPlan({ equipmentLabel, serviceHours });
    const plan = raw && typeof raw === "object" ? raw : {};
    const asList = (v, fb) => Array.isArray(v) ? v.map((x) => String(x || "").trim()).filter(Boolean) : fb;
    const oils = Array.isArray(plan.oils)
      ? plan.oils
          .map((o) => ({
            name: String(o?.name || "").trim(),
            qty: Number(o?.qty),
            unit: String(o?.unit || "").trim() || "L",
          }))
          .filter((o) => o.name && Number.isFinite(o.qty) && o.qty > 0)
      : fallback.oils;
    const filters = Array.isArray(plan.filters)
      ? plan.filters
          .map((f) => ({
            name: String(f?.name || "").trim(),
            qty: Number(f?.qty),
            unit: String(f?.unit || "").trim() || "ea",
          }))
          .filter((f) => f.name && Number.isFinite(f.qty) && f.qty > 0)
      : fallback.filters;
    return {
      service_title: String(plan.service_title || fallback.service_title),
      tasks: asList(plan.tasks, fallback.tasks),
      oils,
      checks: asList(plan.checks, fallback.checks),
      post_service_checks: asList(plan.post_service_checks, fallback.post_service_checks),
      safety: asList(plan.safety, fallback.safety),
      filters,
    };
  }
  async function buildRsgPlan({ assetId, assetCode, equipmentName, serviceHours }) {
    const label = [assetCode, equipmentName].filter(Boolean).join(" - ") || "Equipment";
    const preferredOils = getAssetOilProfile(assetId, 6);
    const profile = pickRsgServiceProfile({ assetCode, equipmentName, serviceHours });
    const aiPlan = await tryGenerateRsgPlanWithAi({ equipmentLabel: label, serviceHours, preferredOils });
    const plan = normalizeRsgPlan(aiPlan, label, serviceHours);

    if (profile) {
      plan.service_title = profile.title || plan.service_title;
      if (Array.isArray(profile.tasks) && profile.tasks.length) plan.tasks = profile.tasks.map((t) => String(t));
      if (Array.isArray(profile.oils) && profile.oils.length) {
        plan.oils = profile.oils.map((o) => ({
          name: String(o.name || ""),
          qty: Number(o.qty || 0),
          unit: String(o.unit || "L"),
          part_hint: String(o.part_hint || o.name || ""),
        })).filter((o) => o.name && Number.isFinite(o.qty) && o.qty > 0);
      }
      if (Array.isArray(profile.filters) && profile.filters.length) {
        plan.filters = profile.filters.map((f) => ({
          name: String(f.name || ""),
          qty: Number(f.qty || 0),
          unit: String(f.unit || "ea"),
          part_hint: String(f.part_hint || f.name || ""),
        })).filter((f) => f.name && Number.isFinite(f.qty) && f.qty > 0);
      }
      plan.checks = [
        ...plan.checks,
        "Profile-based service pack applied (site configured). Verify all capacities with OEM manual before execution.",
      ];
    }

    // Prefer real site-recorded oils/quantities when available for this asset.
    if (!profile && preferredOils.length) {
      plan.oils = preferredOils;
      plan.checks = [
        ...plan.checks,
        "Oil types and quantities are sourced from site oil history for this asset; verify against OEM service manual before execution.",
      ];
    }
    const readiness = buildRsgReadiness({ oils: plan.oils || [], filters: plan.filters || [] });
    return { plan, readiness, profile_key: profile?.key || null };
  }

  app.get("/history", async (req, reply) => {
    try {
      const reportType = String(req.query?.report_type || "daily_admin").trim() || "daily_admin";
      const days = Math.max(1, Math.min(60, Number(req.query?.days || 7)));
      const reports = getIronmindHistory({ reportType, limit: days });
      return reply.send({ ok: true, reports });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.get("/latest", async (req, reply) => {
    try {
      const reportType = String(req.query?.report_type || "daily_admin").trim() || "daily_admin";
      const row = getLatestIronmindReport(reportType);
      if (!row) {
        return reply.send({ ok: true, report: null });
      }
      return reply.send({ ok: true, report: row });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.get("/settings", async (req, reply) => {
    try {
      return reply.send({ ok: true, settings: getIronmindSettings() });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.put("/settings", async (req, reply) => {
    try {
      const body = req.body || {};
      const settings = setIronmindSettings(body);
      return reply.send({ ok: true, settings });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/run", async (req, reply) => {
    try {
      const body = req.body || {};
      const reportDate = String(body.report_date || "").trim() || undefined;
      const reportType = String(body.report_type || "daily_admin").trim() || "daily_admin";
      const force = toBool(body.force);
      const contextNotes = String(body.context_notes || "").trim();
      const detailMode = toBool(body.detail_mode);

      const report = await generateIronmindReport({
        reportDate,
        reportType,
        force,
        contextNotes: contextNotes || undefined,
        detailMode,
      });
      return reply.send({ ok: true, report });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/ask", async (req, reply) => {
    try {
      const body = req.body || {};
      const question = String(body.question || "").trim();
      if (!question) return reply.code(400).send({ ok: false, error: "question is required" });
      const fallbackDate = isDate(body.date) ? String(body.date) : todayYmd();
      const parsed = parseQuestionDates(question, fallbackDate);
      const start = isDate(body.start) ? String(body.start) : parsed.start;
      const end = isDate(body.end) ? String(body.end) : parsed.end;
      const assetCode = String(body.asset_code || parseAssetCode(question)).trim().toUpperCase();
      const qLower = question.toLowerCase();
      if (!assetCode) {
        return reply.send({
          ok: true,
          short_answer: "Please include an asset code, for example: G01AM from 2026-02-14 to 2026-04-07.",
        });
      }
      const row = db.prepare(`
        SELECT
          a.asset_code,
          COUNT(DISTINCT l.log_date) AS logged_days,
          COALESCE(SUM(l.hours_down), 0) AS downtime_hours,
          COALESCE(MIN(l.log_date), '') AS first_log_date,
          COALESCE(MAX(l.log_date), '') AS last_log_date
        FROM assets a
        LEFT JOIN breakdowns b ON b.asset_id = a.id
        LEFT JOIN breakdown_downtime_logs l
          ON l.breakdown_id = b.id
         AND l.log_date BETWEEN ? AND ?
        WHERE UPPER(a.asset_code) = UPPER(?)
      `).get(start, end, assetCode);
      if (!row?.asset_code) {
        return reply.send({ ok: true, short_answer: `No asset found for code ${assetCode}.` });
      }
      const breakdowns = db.prepare(`
        SELECT COUNT(*) AS c
        FROM breakdowns b
        JOIN assets a ON a.id = b.asset_id
        WHERE UPPER(a.asset_code) = UPPER(?)
          AND b.breakdown_date BETWEEN ? AND ?
      `).get(assetCode, start, end);
      const openBreakdown = db.prepare(`
        SELECT b.id, b.breakdown_date
        FROM breakdowns b
        JOIN assets a ON a.id = b.asset_id
        WHERE UPPER(a.asset_code) = UPPER(?)
          AND b.status = 'OPEN'
        ORDER BY b.id DESC
        LIMIT 1
      `).get(assetCode);

      const asksFuel = qLower.includes("fuel") || qLower.includes("l/hr") || qLower.includes("km/l");
      const asksRecurring = qLower.includes("recurring") || qLower.includes("repeat") || qLower.includes("failures");
      const asksPm = qLower.includes("pm") || qLower.includes("overdue") || qLower.includes("maintenance");

      if (asksFuel) {
        const hasBaseline = db.prepare(`PRAGMA table_info(assets)`).all().some((c) => String(c.name) === "baseline_fuel_l_per_hour");
        const fuelRow = db.prepare(`
          SELECT
            a.asset_code,
            COALESCE(SUM(fl.liters), 0) AS liters,
            COALESCE(SUM(fl.hours_run), 0) AS run_hours,
            COALESCE(a.baseline_fuel_l_per_hour, 0) AS baseline_lph
          FROM assets a
          LEFT JOIN fuel_logs fl ON fl.asset_id = a.id AND fl.log_date BETWEEN ? AND ?
          WHERE UPPER(a.asset_code) = UPPER(?)
          GROUP BY a.id
        `).get(start, end, assetCode);
        if (!fuelRow) {
          return reply.send({ ok: true, short_answer: `No fuel records found for ${assetCode} in ${start} to ${end}.` });
        }
        const liters = Number(fuelRow.liters || 0);
        const runHours = Number(fuelRow.run_hours || 0);
        const actualLph = runHours > 0 ? liters / runHours : 0;
        const baseline = Number(fuelRow.baseline_lph || 0);
        const overPct = baseline > 0 ? ((actualLph / baseline) - 1) * 100 : null;
        const short = baseline > 0
          ? `${assetCode}: fuel ${actualLph.toFixed(2)} L/h vs baseline ${baseline.toFixed(2)} L/h (${overPct >= 0 ? "+" : ""}${overPct.toFixed(1)}%) from ${start} to ${end}.`
          : `${assetCode}: fuel ${actualLph.toFixed(2)} L/h from ${start} to ${end}. Baseline not configured.`;
        return reply.send({
          ok: true,
          short_answer: short,
          details: {
            asset_code: assetCode,
            start,
            end,
            liters,
            run_hours: runHours,
            actual_lph: Number(actualLph.toFixed(3)),
            baseline_lph: baseline,
            over_pct: overPct == null ? null : Number(overPct.toFixed(2)),
            baseline_configured: hasBaseline && baseline > 0,
          },
        });
      }

      if (asksRecurring) {
        const recurring = db.prepare(`
          SELECT
            COUNT(DISTINCT l.breakdown_id) AS incidents,
            COALESCE(SUM(l.hours_down), 0) AS downtime_hours,
            COALESCE(MIN(l.log_date), '') AS first_log_date,
            COALESCE(MAX(l.log_date), '') AS last_log_date
          FROM breakdown_downtime_logs l
          JOIN breakdowns b ON b.id = l.breakdown_id
          JOIN assets a ON a.id = b.asset_id
          WHERE UPPER(a.asset_code) = UPPER(?)
            AND l.log_date BETWEEN ? AND ?
        `).get(assetCode, start, end);
        const incidents = Number(recurring?.incidents || 0);
        const dt = Number(recurring?.downtime_hours || 0);
        const short = `${assetCode}: ${incidents} recurring failure incident(s), ${dt.toFixed(1)}h downtime between ${start} and ${end}.`;
        return reply.send({
          ok: true,
          short_answer: short,
          details: {
            asset_code: assetCode,
            start,
            end,
            incidents,
            downtime_hours: dt,
            first_log_date: recurring?.first_log_date || null,
            last_log_date: recurring?.last_log_date || null,
          },
        });
      }

      if (asksPm) {
        const pm = db.prepare(`
          SELECT
            mp.service_name,
            (COALESCE((
              SELECT SUM(dh.hours_run)
              FROM daily_hours dh
              JOIN assets a2 ON a2.id = dh.asset_id
              WHERE dh.asset_id = mp.asset_id
                AND dh.is_used = 1
                AND dh.hours_run > 0
                AND dh.work_date <= ?
            ), 0) - (mp.last_service_hours + mp.interval_hours)) AS overdue_hours
          FROM maintenance_plans mp
          JOIN assets a ON a.id = mp.asset_id
          WHERE UPPER(a.asset_code) = UPPER(?)
            AND mp.active = 1
          ORDER BY overdue_hours DESC
          LIMIT 1
        `).get(end, assetCode);
        const overdue = Number(pm?.overdue_hours || 0);
        const riskBand = overdue >= 200 ? "high" : overdue >= 50 ? "medium" : overdue > 0 ? "low" : "none";
        const short = overdue > 0
          ? `${assetCode}: PM overdue by ${overdue.toFixed(1)}h (${riskBand} risk) as of ${end}${pm?.service_name ? ` on ${pm.service_name}` : ""}.`
          : `${assetCode}: no active PM overdue as of ${end}.`;
        return reply.send({
          ok: true,
          short_answer: short,
          details: {
            asset_code: assetCode,
            as_of: end,
            service_name: pm?.service_name || null,
            overdue_hours: overdue,
            risk_band: riskBand,
          },
        });
      }

      const hours = Number(row.downtime_hours || 0);
      const short = `${assetCode}: ${hours.toFixed(1)}h downtime from ${start} to ${end} across ${Number(row.logged_days || 0)} logged day(s).`;
      return reply.send({
        ok: true,
        short_answer: short,
        details: {
          asset_code: assetCode,
          start,
          end,
          downtime_hours: hours,
          logged_days: Number(row.logged_days || 0),
          first_log_date: row.first_log_date || null,
          last_log_date: row.last_log_date || null,
          breakdowns_in_range: Number(breakdowns?.c || 0),
          has_open_breakdown: Boolean(openBreakdown),
          open_breakdown_date: openBreakdown?.breakdown_date || null,
        },
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // POST /api/ironmind/rsg/plan
  // Body: { asset_code?, equipment_name?, service_hours? }
  app.post("/rsg/plan", async (req, reply) => {
    try {
      const body = req.body || {};
      const serviceHours = Math.max(250, Number(body.service_hours || 2000));
      let assetCode = String(body.asset_code || "").trim().toUpperCase();
      let equipmentName = String(body.equipment_name || "").trim();
      if (assetCode) {
        const a = db.prepare(`SELECT asset_code, asset_name FROM assets WHERE UPPER(asset_code)=UPPER(?) LIMIT 1`).get(assetCode);
        if (a) {
          assetCode = String(a.asset_code || assetCode).toUpperCase();
          if (!equipmentName) equipmentName = String(a.asset_name || "");
        }
      }
      const assetRow = assetCode
        ? db.prepare(`SELECT id, asset_code, asset_name FROM assets WHERE UPPER(asset_code)=UPPER(?) LIMIT 1`).get(assetCode)
        : null;
      const rsg = await buildRsgPlan({
        assetId: Number(assetRow?.id || 0),
        assetCode,
        equipmentName: equipmentName || String(assetRow?.asset_name || ""),
        serviceHours,
      });
      return reply.send({
        ok: true,
        asset_code: assetCode || null,
        equipment_name: equipmentName || null,
        service_hours: serviceHours,
        plan: rsg.plan,
        readiness: rsg.readiness,
        profile_key: rsg.profile_key,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // POST /api/ironmind/rsg/create-wo
  // Body: { asset_code, service_hours?, equipment_name? }
  app.post("/rsg/create-wo", async (req, reply) => {
    try {
      const body = req.body || {};
      const assetCode = String(body.asset_code || "").trim().toUpperCase();
      const serviceHours = Math.max(250, Number(body.service_hours || 2000));
      const equipmentNameIn = String(body.equipment_name || "").trim();
      if (!assetCode) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      const asset = db.prepare(`SELECT id, asset_code, asset_name FROM assets WHERE UPPER(asset_code)=UPPER(?) LIMIT 1`).get(assetCode);
      if (!asset) return reply.code(404).send({ ok: false, error: `asset not found: ${assetCode}` });
      const equipmentName = equipmentNameIn || String(asset.asset_name || "");
      const rsg = await buildRsgPlan({ assetId: asset.id, assetCode: asset.asset_code, equipmentName, serviceHours });
      const plan = rsg.plan;

      const wo = db.prepare(`
        INSERT INTO work_orders (asset_id, source, reference_id, status)
        VALUES (?, ?, NULL, 'open')
      `).run(asset.id, `rsg_${serviceHours}h_service`);
      const woId = Number(wo.lastInsertRowid || 0);

      if (woId > 0 && hasColumn("work_orders", "completion_notes")) {
        const notes = [
          `RSG Service Guide: ${plan.service_title}`,
          "",
          "Tasks:",
          ...plan.tasks.map((t) => `- ${t}`),
          "",
          "Oil / Lubricants:",
          ...plan.oils.map((o) => `- ${o.name}: ${o.qty} ${o.unit}`),
          "",
          "Checks:",
          ...plan.checks.map((c) => `- ${c}`),
          "",
          "Post-service checks:",
          ...plan.post_service_checks.map((c) => `- ${c}`),
          "",
          "Safety:",
          ...plan.safety.map((s) => `- ${s}`),
        ].join("\n");
        db.prepare(`UPDATE work_orders SET completion_notes = ? WHERE id = ?`).run(notes, woId);
      }

      return reply.send({
        ok: true,
        work_order_id: woId,
        asset_code: String(asset.asset_code || assetCode),
        service_hours: serviceHours,
        source: `rsg_${serviceHours}h_service`,
        plan,
        readiness: rsg.readiness,
        profile_key: rsg.profile_key,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // GET /api/ironmind/rsg/preview.pdf?asset_code=A300AM&service_hours=2000&download=1
  app.get("/rsg/preview.pdf", async (req, reply) => {
    try {
      const assetCodeIn = String(req.query?.asset_code || "").trim().toUpperCase();
      const serviceHours = Math.max(250, Number(req.query?.service_hours || 2000));
      const download = String(req.query?.download || "").trim() === "1";
      if (!assetCodeIn) return reply.code(400).send({ ok: false, error: "asset_code is required" });

      const asset = db.prepare(`SELECT id, asset_code, asset_name FROM assets WHERE UPPER(asset_code)=UPPER(?) LIMIT 1`).get(assetCodeIn);
      if (!asset) return reply.code(404).send({ ok: false, error: `asset not found: ${assetCodeIn}` });

      const rsg = await buildRsgPlan({
        assetId: asset.id,
        assetCode: String(asset.asset_code || assetCodeIn),
        equipmentName: String(asset.asset_name || ""),
        serviceHours,
      });
      const plan = rsg.plan;

      const pdf = await buildPdfBuffer((doc) => {
        sectionTitle(doc, "Service Guide");
        table(
          doc,
          [
            { key: "k", label: "Field", width: 0.32 },
            { key: "v", label: "Value", width: 0.68 },
          ],
          [
            { k: "Asset", v: String(asset.asset_code || "-") },
            { k: "Equipment", v: String(asset.asset_name || "-") },
            { k: "Service Interval", v: `${serviceHours} hours` },
            { k: "Guide Title", v: String(plan.service_title || "-") },
          ]
        );

        sectionTitle(doc, "Service Tasks");
        table(
          doc,
          [
            { key: "idx", label: "#", width: 0.08, align: "right" },
            { key: "task", label: "Task", width: 0.92 },
          ],
          (plan.tasks || []).map((t, i) => ({ idx: String(i + 1), task: String(t || "") }))
        );

        sectionTitle(doc, "Oil / Lubricants");
        table(
          doc,
          [
            { key: "name", label: "Oil/Lube", width: 0.56 },
            { key: "qty", label: "Qty", width: 0.20, align: "right" },
            { key: "unit", label: "Unit", width: 0.24, align: "center" },
          ],
          (plan.oils || []).map((o) => ({
            name: String(o?.name || ""),
            qty: Number(o?.qty || 0).toFixed(2),
            unit: String(o?.unit || "L"),
          }))
        );

        sectionTitle(doc, "Filters / Service Kit");
        table(
          doc,
          [
            { key: "name", label: "Filter / Kit Item", width: 0.62 },
            { key: "qty", label: "Qty", width: 0.2, align: "right" },
            { key: "unit", label: "Unit", width: 0.18, align: "center" },
          ],
          (plan.filters || []).length
            ? (plan.filters || []).map((f) => ({
                name: String(f?.name || ""),
                qty: Number(f?.qty || 0).toFixed(2),
                unit: String(f?.unit || "ea"),
              }))
            : [{ name: "No explicit filter kit defined in profile", qty: "-", unit: "-" }]
        );

        sectionTitle(doc, "Stores Readiness Check");
        table(
          doc,
          [
            { key: "category", label: "Type", width: 0.12 },
            { key: "name", label: "Requirement", width: 0.34 },
            { key: "required", label: "Required", width: 0.12, align: "right" },
            { key: "on_hand", label: "On Hand", width: 0.12, align: "right" },
            { key: "shortage", label: "Shortage", width: 0.12, align: "right" },
            { key: "part", label: "Stores Link", width: 0.1 },
            { key: "status", label: "Status", width: 0.08, align: "center" },
          ],
          (rsg.readiness?.items || []).length
            ? rsg.readiness.items.map((it) => ({
                category: String(it.category || "").toUpperCase(),
                name: String(it.name || ""),
                required: `${Number(it.required_qty || 0).toFixed(2)} ${String(it.unit || "")}`.trim(),
                on_hand: `${Number(it.on_hand_qty || 0).toFixed(2)} ${String(it.unit || "")}`.trim(),
                shortage: `${Number(it.shortage_qty || 0).toFixed(2)} ${String(it.unit || "")}`.trim(),
                part: it.part_code ? `${String(it.part_code)}${it.part_name ? ` (${String(it.part_name)})` : ""}` : "Not linked",
                status: it.sufficient ? "OK" : "SHORT",
              }))
            : [{ category: "-", name: "No readiness items", required: "-", on_hand: "-", shortage: "-", part: "-", status: "-" }]
        );

        sectionTitle(doc, "Checks");
        table(
          doc,
          [
            { key: "idx", label: "#", width: 0.08, align: "right" },
            { key: "txt", label: "Checks / Verification", width: 0.92 },
          ],
          (plan.checks || []).map((t, i) => ({ idx: String(i + 1), txt: String(t || "") }))
        );

        sectionTitle(doc, "Post-Service Checks");
        table(
          doc,
          [
            { key: "idx", label: "#", width: 0.08, align: "right" },
            { key: "txt", label: "Post-Service", width: 0.92 },
          ],
          (plan.post_service_checks || []).map((t, i) => ({ idx: String(i + 1), txt: String(t || "") }))
        );
      }, {
        title: "IRONLOG",
        subtitle: "Recommended Service Guide (RSG)",
        rightText: `${asset.asset_code} • ${serviceHours}h`,
        layout: "landscape",
      });

      reply
        .header("Content-Type", "application/pdf")
        .header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate")
        .header("Pragma", "no-cache")
        .header("Expires", "0")
        .header(
          "Content-Disposition",
          `${download ? "attachment" : "inline"}; filename="RSG_${asset.asset_code}_${serviceHours}h.pdf"`
        )
        .send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/risk-board", async (req, reply) => {
    try {
      const asOf = isDate(req.query?.date) ? String(req.query.date) : todayYmd();
      const limit = Math.max(1, Math.min(20, Number(req.query?.limit || 8)));
      const rows = db.prepare(`
        SELECT r.asset_code, r.risk_score, r.confidence, r.reasons_json, r.features_json, r.report_date
        FROM ironmind_asset_risk_snapshots r
        JOIN (
          SELECT asset_code, MAX(report_date) AS latest_date
          FROM ironmind_asset_risk_snapshots
          WHERE report_date <= ?
          GROUP BY asset_code
        ) x ON x.asset_code = r.asset_code AND x.latest_date = r.report_date
        ORDER BY r.risk_score DESC, r.confidence DESC, r.asset_code ASC
        LIMIT ?
      `).all(asOf, limit);

      const items = (rows || []).map((r) => {
        let reasons = [];
        try { reasons = JSON.parse(String(r.reasons_json || "[]")); } catch {}
        let features = {};
        try { features = JSON.parse(String(r.features_json || "{}")); } catch {}
        return {
          asset_code: String(r.asset_code || ""),
          report_date: String(r.report_date || ""),
          risk_score: Number(r.risk_score || 0),
          confidence: Number(r.confidence || 0),
          reasons: Array.isArray(reasons) ? reasons : [],
          features,
        };
      });

      return reply.send({ ok: true, as_of: asOf, items });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
}
