import { db } from "../db/client.js";

function nowIso() {
  return new Date().toISOString();
}

function formatLocalYmd(date) {
  const d = new Date(date);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function ymdOffsetFromNow(daysBack = 1) {
  const d = new Date();
  d.setDate(d.getDate() - Number(daysBack || 0));
  return formatLocalYmd(d);
}

export function ensureIronmindTable() {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS ironmind_reports (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      report_date TEXT NOT NULL,
      report_type TEXT NOT NULL DEFAULT 'daily_admin',
      summary TEXT NOT NULL,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `).run();

  db.prepare(`
    CREATE UNIQUE INDEX IF NOT EXISTS idx_ironmind_reports_unique_date_type
    ON ironmind_reports(report_date, report_type)
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS ironmind_asset_risk_snapshots (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      report_date TEXT NOT NULL,
      asset_code TEXT NOT NULL,
      risk_score REAL NOT NULL DEFAULT 0,
      confidence REAL NOT NULL DEFAULT 0,
      reasons_json TEXT NOT NULL DEFAULT '[]',
      features_json TEXT NOT NULL DEFAULT '{}',
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `).run();
  db.prepare(`
    CREATE UNIQUE INDEX IF NOT EXISTS idx_ironmind_asset_risk_unique
    ON ironmind_asset_risk_snapshots(report_date, asset_code)
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS ironmind_settings (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL,
      updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `).run();
}

function getAiConfig() {
  const openaiKey = process.env.OPENAI_API_KEY;
  const openaiModel = process.env.OPENAI_MODEL || "gpt-4o-mini";

  const azureKey = process.env.AZURE_OPENAI_API_KEY;
  const azureEndpoint = process.env.AZURE_OPENAI_ENDPOINT;
  const azureDeployment = process.env.AZURE_OPENAI_DEPLOYMENT;
  const azureApiVersion = process.env.AZURE_OPENAI_API_VERSION || "2024-02-15-preview";

  const foundryKey =
    process.env.FOUNDRY_API_KEY ||
    process.env.AZURE_FOUNDRY_API_KEY ||
    process.env.AZURE_OPENAI_API_KEY ||
    "";
  const foundryEndpoint =
    process.env.FOUNDRY_ENDPOINT ||
    process.env.AZURE_FOUNDRY_ENDPOINT ||
    process.env.AZURE_EXISTING_AIPROJECT_ENDPOINT ||
    "";
  const foundryModel =
    process.env.FOUNDRY_MODEL ||
    process.env.AZURE_FOUNDRY_MODEL ||
    process.env.OPENAI_MODEL ||
    "gpt-4o-mini";

  if (openaiKey) return { provider: "openai", apiKey: openaiKey, model: openaiModel };
  if (azureKey && azureEndpoint && azureDeployment) {
    return {
      provider: "azure_openai",
      apiKey: azureKey,
      endpoint: azureEndpoint,
      deployment: azureDeployment,
      apiVersion: azureApiVersion,
    };
  }
  if (foundryKey && foundryEndpoint) {
    return {
      provider: "foundry",
      apiKey: foundryKey,
      endpoint: foundryEndpoint,
      model: foundryModel,
    };
  }

  return { provider: null };
}

function normalizeFoundryChatEndpoint(endpoint) {
  const base = String(endpoint || "").trim().replace(/\/+$/, "");
  if (!base) return "";
  if (/\/openai\/v1$/i.test(base)) return `${base}/chat/completions`;
  if (/\/api\/projects\//i.test(base)) return `${base}/openai/v1/chat/completions`;
  return `${base}/openai/v1/chat/completions`;
}

function safeList(arr, fallback) {
  return Array.isArray(arr) && arr.length ? arr : fallback;
}

function hasColumn(table, col) {
  const rows = db.prepare(`PRAGMA table_info(${table})`).all();
  return rows.some((r) => String(r.name) === col);
}
function hasTable(name) {
  const row = db
    .prepare("SELECT name FROM sqlite_master WHERE type='table' AND name = ?")
    .get(String(name || ""));
  return Boolean(row);
}

function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

function maxTrustedDailyRunHours() {
  const v = Number(getIronmindSettingValue("max_daily_run_hours", process.env.IRONMIND_MAX_DAILY_RUN_HOURS || 24));
  return Number.isFinite(v) && v > 0 ? v : 24;
}

function getIronmindSettingValue(key, fallback) {
  try {
    if (!hasTable("ironmind_settings")) return fallback;
    const row = db.prepare(`SELECT value FROM ironmind_settings WHERE key = ? LIMIT 1`).get(String(key || ""));
    if (!row || row.value == null) return fallback;
    return row.value;
  } catch {
    return fallback;
  }
}

function getCategoryTrustedDailyRunHours(category) {
  const base = maxTrustedDailyRunHours();
  const key = String(category || "").trim().toLowerCase();
  if (!key) return base;
  if (key.includes("ldv") || key.includes("pickup") || key.includes("light vehicle")) {
    const configured = getIronmindSettingValue("max_daily_run_hours_ldv", process.env.IRONMIND_MAX_DAILY_RUN_HOURS_LDV || Math.min(base, 16));
    return Number.isFinite(Number(configured)) && Number(configured) > 0 ? Number(configured) : Math.min(base, 16);
  }
  if (key.includes("truck") || key.includes("tipper")) {
    const configured = getIronmindSettingValue("max_daily_run_hours_truck", process.env.IRONMIND_MAX_DAILY_RUN_HOURS_TRUCK || Math.min(base, 18));
    return Number.isFinite(Number(configured)) && Number(configured) > 0 ? Number(configured) : Math.min(base, 18);
  }
  if (key.includes("excavator") || key.includes("loader") || key.includes("dozer")) {
    const configured = getIronmindSettingValue("max_daily_run_hours_heavy", process.env.IRONMIND_MAX_DAILY_RUN_HOURS_HEAVY || base);
    return Number.isFinite(Number(configured)) && Number(configured) > 0 ? Number(configured) : base;
  }
  return base;
}

export function getIronmindSettings() {
  ensureIronmindTable();
  const num = (v, d) => {
    const n = Number(v);
    return Number.isFinite(n) && n > 0 ? n : d;
  };
  const baseDefault = num(process.env.IRONMIND_MAX_DAILY_RUN_HOURS || 24, 24);
  const defaults = {
    max_daily_run_hours: baseDefault,
    max_daily_run_hours_ldv: num(process.env.IRONMIND_MAX_DAILY_RUN_HOURS_LDV || Math.min(baseDefault, 16), Math.min(baseDefault, 16)),
    max_daily_run_hours_truck: num(process.env.IRONMIND_MAX_DAILY_RUN_HOURS_TRUCK || Math.min(baseDefault, 18), Math.min(baseDefault, 18)),
    max_daily_run_hours_heavy: num(process.env.IRONMIND_MAX_DAILY_RUN_HOURS_HEAVY || baseDefault, baseDefault),
  };
  const rows = db.prepare(`SELECT key, value FROM ironmind_settings`).all();
  const out = { ...defaults };
  for (const r of rows) {
    const k = String(r.key || "").trim();
    if (!Object.prototype.hasOwnProperty.call(out, k)) continue;
    out[k] = num(r.value, out[k]);
  }
  return out;
}

export function setIronmindSettings(patch = {}) {
  ensureIronmindTable();
  const allowed = ["max_daily_run_hours", "max_daily_run_hours_ldv", "max_daily_run_hours_truck", "max_daily_run_hours_heavy"];
  const upsert = db.prepare(`
    INSERT INTO ironmind_settings (key, value, updated_at)
    VALUES (?, ?, datetime('now'))
    ON CONFLICT(key) DO UPDATE SET
      value = excluded.value,
      updated_at = excluded.updated_at
  `);
  const tx = db.transaction(() => {
    for (const k of allowed) {
      if (!Object.prototype.hasOwnProperty.call(patch, k)) continue;
      const n = Number(patch[k]);
      if (!Number.isFinite(n) || n <= 0) continue;
      upsert.run(k, String(n));
    }
  });
  tx();
  return getIronmindSettings();
}

function computeAssetRiskSignals(reportDate) {
  const assets = db.prepare(`
    SELECT id, asset_code, asset_name, category
    FROM assets
    WHERE active = 1 AND IFNULL(is_standby, 0) = 0
  `).all();
  const byAssetId = new Map();
  assets.forEach((a) => {
    byAssetId.set(Number(a.id), {
      asset_id: Number(a.id),
      asset_code: String(a.asset_code || ""),
      asset_name: String(a.asset_name || ""),
      category: String(a.category || ""),
      incidents_30d: 0,
      downtime_30d: 0,
      overdue_hours: 0,
      max_open_wo_age_hours: 0,
      fuel_over_pct: 0,
    });
  });

  const failures = db.prepare(`
    SELECT b.asset_id, COUNT(DISTINCT l.breakdown_id) AS incidents, COALESCE(SUM(l.hours_down), 0) AS downtime
    FROM breakdown_downtime_logs l
    JOIN breakdowns b ON b.id = l.breakdown_id
    JOIN assets a ON a.id = b.asset_id
    WHERE l.log_date BETWEEN DATE(?, '-29 day') AND DATE(?)
      AND a.active = 1
      AND IFNULL(a.is_standby, 0) = 0
    GROUP BY b.asset_id
  `).all(reportDate, reportDate);
  failures.forEach((r) => {
    const row = byAssetId.get(Number(r.asset_id));
    if (!row) return;
    row.incidents_30d = Number(r.incidents || 0);
    row.downtime_30d = Number(r.downtime || 0);
  });

  const sumTrustedRun = db.prepare(`
    SELECT COALESCE(SUM(dh.hours_run), 0) AS run_total
    FROM daily_hours dh
    WHERE dh.asset_id = ?
      AND dh.is_used = 1
      AND dh.hours_run > 0
      AND dh.hours_run <= ?
      AND dh.work_date <= ?
  `);
  const activePlans = db.prepare(`
    SELECT
      mp.asset_id,
      mp.interval_hours,
      mp.last_service_hours,
      a.category
    FROM maintenance_plans mp
    JOIN assets a ON a.id = mp.asset_id
    WHERE mp.active = 1
      AND a.active = 1
      AND IFNULL(a.is_standby, 0) = 0
  `).all();
  for (const p of activePlans) {
    const aid = Number(p.asset_id || 0);
    if (!aid) continue;
    const cap = getCategoryTrustedDailyRunHours(p.category);
    const runTotal = Number(sumTrustedRun.get(aid, cap, reportDate)?.run_total || 0);
    const overdue = runTotal - (Number(p.last_service_hours || 0) + Number(p.interval_hours || 0));
    if (overdue < 0) continue;
    const row = byAssetId.get(aid);
    if (!row) continue;
    row.overdue_hours = Math.max(Number(row.overdue_hours || 0), Number(overdue || 0));
  }

  const woAges = db.prepare(`
    SELECT w.asset_id, MAX(CAST((julianday('now') - julianday(COALESCE(w.opened_at, datetime('now')))) * 24 AS INTEGER)) AS max_age
    FROM work_orders w
    JOIN assets a ON a.id = w.asset_id
    WHERE w.status IN ('open', 'assigned', 'in_progress', 'completed')
      AND a.active = 1
      AND IFNULL(a.is_standby, 0) = 0
    GROUP BY w.asset_id
  `).all();
  woAges.forEach((r) => {
    const row = byAssetId.get(Number(r.asset_id));
    if (!row) return;
    row.max_open_wo_age_hours = Number(r.max_age || 0);
  });

  if (hasColumn("assets", "baseline_fuel_l_per_hour")) {
    const fuelRows = db.prepare(`
      SELECT a.id AS asset_id,
        CASE WHEN COALESCE(a.baseline_fuel_l_per_hour, 0) > 0
          THEN ((COALESCE(SUM(fl.liters), 0) / SUM(fl.hours_run)) / a.baseline_fuel_l_per_hour - 1.0) * 100.0
          ELSE 0
        END AS over_pct
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date BETWEEN DATE(?, '-13 day') AND DATE(?)
        AND COALESCE(fl.hours_run, 0) > 0
        AND a.active = 1
        AND IFNULL(a.is_standby, 0) = 0
      GROUP BY a.id
      HAVING COALESCE(a.baseline_fuel_l_per_hour, 0) > 0
    `).all(reportDate, reportDate);
    fuelRows.forEach((r) => {
      const row = byAssetId.get(Number(r.asset_id));
      if (!row) return;
      row.fuel_over_pct = Number(r.over_pct || 0);
    });
  }

  const scored = Array.from(byAssetId.values()).map((r) => {
    const failurePts = clamp(r.incidents_30d * 12, 0, 35);
    const downtimePts = clamp(r.downtime_30d * 1.2, 0, 20);
    const overduePts = clamp(r.overdue_hours / 20, 0, 20);
    const woAgePts = clamp(r.max_open_wo_age_hours / 6, 0, 15);
    const fuelPts = clamp(Math.max(0, r.fuel_over_pct - 10) * 0.8, 0, 20);
    const riskScore = Number(clamp(failurePts + downtimePts + overduePts + woAgePts + fuelPts, 0, 100).toFixed(2));
    const signalCount = [r.incidents_30d > 0, r.downtime_30d > 0, r.overdue_hours > 0, r.max_open_wo_age_hours > 0, r.fuel_over_pct > 0]
      .filter(Boolean).length;
    const confidence = Number(clamp(35 + signalCount * 12 + Math.min(20, r.incidents_30d * 4), 20, 95).toFixed(0));
    const reasons = [];
    if (r.incidents_30d >= 2) reasons.push(`${r.incidents_30d} breakdown incidents in 30d`);
    if (r.downtime_30d > 8) reasons.push(`${r.downtime_30d.toFixed(1)}h downtime in 30d`);
    if (r.overdue_hours > 0) reasons.push(`${r.overdue_hours.toFixed(1)}h PM overdue`);
    if (r.max_open_wo_age_hours > 24) reasons.push(`open WO age ${r.max_open_wo_age_hours}h`);
    if (r.fuel_over_pct > 20) reasons.push(`fuel rate +${r.fuel_over_pct.toFixed(1)}% vs baseline`);
    return { ...r, risk_score: riskScore, confidence, reasons };
  }).sort((a, b) => b.risk_score - a.risk_score);

  return scored;
}

function toIronmindFormat({ repairsNeeded, operationalRisks, suggestions, dataGaps }) {
  const block = (title, lines) => {
    const normalized = safeList(lines, ["Insufficient data"]).map((s) => `- ${String(s || "").trim() || "Insufficient data"}`);
    return `${title}\n${normalized.join("\n")}`;
  };

  return [
    "IRONMIND DAILY INSIGHT",
    "",
    block("Repairs Needed", repairsNeeded),
    "",
    block("Operational Risks", operationalRisks),
    "",
    block("Suggestions", suggestions),
    "",
    block("Data Gaps", dataGaps),
  ].join("\n");
}

function appendDataAnomaliesSection(baseSummary, lines) {
  const block = (title, items) => {
    const normalized = (items && items.length
      ? items
      : ["No suspicious daily_input rows in the last 90 days."]
    ).map((s) => `- ${String(s || "").trim()}`);
    return `${title}\n${normalized.join("\n")}`;
  };
  return `${String(baseSummary || "").trim()}\n\n${block("Data Anomalies", lines)}`;
}

function buildDailyInputAnomalyLines(rows, maxTrustedHours) {
  const maxH = Number(maxTrustedHours || 24);
  if (!Array.isArray(rows) || !rows.length) {
    return [
      `No rows flagged in the last 90 days (hours_run > ${maxH}h or closing_hours < opening_hours). PM overdue math ignores those run rows.`,
    ];
  }
  return rows.map((r) => {
    const code = String(r.asset_code || "?").trim();
    const d = String(r.work_date || "?").trim();
    const reason = String(r.reason || "flagged").trim();
    const hr = r.hours_run != null ? `${Number(r.hours_run).toFixed(1)}h run` : "run n/a";
    const op = r.opening_hours != null ? Number(r.opening_hours).toFixed(1) : "-";
    const cl = r.closing_hours != null ? Number(r.closing_hours).toFixed(1) : "-";
    return `${code} on ${d}: ${reason} (${hr}, open ${op} / close ${cl}) — correct Daily Input or repair meter chain.`;
  });
}

function buildFallbackInsight(data) {
  const repairsNeeded = [];
  const operationalRisks = [];
  const suggestions = [];
  const dataGaps = [];

  for (const r of data.overdueMaintenance.slice(0, 4)) {
    repairsNeeded.push(`${r.asset_code || "Unknown asset"}: PM overdue by ${Number(r.overdue_hours || 0).toFixed(1)} hours (${r.service_name || "service"}).`);
  }

  for (const w of data.openWorkOrders.slice(0, 3)) {
    if (Number(w.age_hours || 0) >= 24) {
      repairsNeeded.push(`${w.asset_code || "Unknown asset"}: Work order #${w.id} still ${w.status || "open"} at ${w.age_hours}h age.`);
    }
  }

  if (Number(data.kpi.availability_pct) < 70) {
    operationalRisks.push(`Fleet availability for ${data.report_date} is ${data.kpi.availability_pct.toFixed(1)}%, below 70%.`);
  }
  if (Number(data.kpi.utilization_pct) < 60) {
    operationalRisks.push(`Fleet utilization for ${data.report_date} is ${data.kpi.utilization_pct.toFixed(1)}%, below 60%.`);
  }

  for (const s of data.lowStockCritical.slice(0, 3)) {
    operationalRisks.push(`Critical stock risk: ${s.part_code || "part"} on hand ${Number(s.on_hand || 0)} below min ${Number(s.min_stock || 0)}.`);
  }
  for (const r of (data.recurringFailures30d || []).slice(0, 3)) {
    operationalRisks.push(`${r.asset_code || "Unknown asset"} recurring breakdown pattern: ${Number(r.incidents || 0)} incidents in 30d, ${Number(r.downtime_hours || 0).toFixed(1)}h downtime.`);
  }
  for (const f of (data.fuelAnomalies14d || []).slice(0, 2)) {
    operationalRisks.push(`${f.asset_code || "Unknown asset"} high fuel trend: ${f.actual_lph.toFixed(2)} L/h vs baseline ${f.baseline_lph.toFixed(2)} L/h (${f.over_pct.toFixed(1)}% over).`);
  }

  suggestions.push("Close or update aged work orders before next shift handover.");
  suggestions.push("Prioritize overdue preventive maintenance on highest-utilization assets.");
  suggestions.push("Reconcile critical spares below minimum and confirm delivery ETA.");
  suggestions.push("Review recurring-failure assets for root cause elimination and planned intervention before next failure window.");
  suggestions.push("Validate abnormal fuel-use assets (idling, route/load, injector/air leaks) and schedule targeted inspections.");
  for (const a of (data.topRiskAssets || []).slice(0, 3)) {
    suggestions.push(`${a.asset_code || "Asset"} risk ${Number(a.risk_score || 0).toFixed(0)}/100 (confidence ${Number(a.confidence || 0).toFixed(0)}%): ${safeList(a.reasons, ["insufficient signals"]).join("; ")}.`);
  }

  if (Number(data.dataCoverage.active_assets || 0) > 0 && Number(data.dataCoverage.assets_with_daily_entry || 0) === 0) {
    dataGaps.push("Insufficient data: no daily hours captured for active assets.");
  }
  if (Number(data.dataCoverage.breakdown_logs || 0) === 0) {
    dataGaps.push("Insufficient data: no breakdown downtime logs recorded.");
  }
  if (Number(data.dataCoverage.operations_daily_rows || 0) === 0) {
    dataGaps.push("Insufficient data: no site operations daily rows recorded.");
  }

  return {
    repairsNeeded: safeList(repairsNeeded, ["Insufficient data"]),
    operationalRisks: safeList(operationalRisks, ["Insufficient data"]),
    suggestions: safeList(suggestions, ["Insufficient data"]),
    dataGaps: safeList(dataGaps, ["Insufficient data"]),
  };
}

function parseJsonObject(text) {
  const src = String(text || "").trim();
  if (!src) return null;
  try {
    return JSON.parse(src);
  } catch (_) {
    const start = src.indexOf("{");
    const end = src.lastIndexOf("}");
    if (start >= 0 && end > start) {
      try {
        return JSON.parse(src.slice(start, end + 1));
      } catch (_) {
        return null;
      }
    }
    return null;
  }
}

async function callIronmindAi(structuredData, opts = {}) {
  const cfg = getAiConfig();
  if (!cfg.provider) return null;
  const contextNotes = String(opts.contextNotes || "").trim();
  const detailMode = Boolean(opts.detailMode);

  const systemPrompt = [
    "You are IRONMIND, the operational intelligence layer for IRONLOG.",
    "",
    "Rules:",
    "- Use only the provided data.",
    "- Do not guess.",
    "- If data is missing or unclear, state 'Insufficient data'.",
    "- Be concise, factual, and operational.",
    "- No motivational language.",
    "- No chit-chat.",
    "- Respond as JSON only with keys: repairs_needed, operational_risks, suggestions, data_gaps.",
    "- Structured JSON may include daily_input_anomalies: treat those as data quality issues, not equipment failure.",
    "- Prefer workshop-actionable bullets in this format: 'ASSET: issue | Owner: role | Parts: item/code or none | ETA: value'.",
    detailMode
      ? "- Each key should contain 4-8 concise but specific bullets with asset codes, magnitudes, and action intent where possible."
      : "- Each key must be an array of short strings.",
    "- Predictive statements must be evidence-based from trend signals only; never present certainty.",
  ].join("\n");
  const contextBlock = contextNotes ? `\n\nAdditional operator context:\n${contextNotes}` : "";
  const userPrompt = `Structured plant data for report date ${structuredData.report_date}:\n${JSON.stringify(structuredData, null, 2)}${contextBlock}`;

  try {
    if (cfg.provider === "openai") {
      const res = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${cfg.apiKey}`,
        },
        body: JSON.stringify({
          model: cfg.model,
          temperature: Number(process.env.IRONMIND_TEMPERATURE ?? 0.1),
          max_tokens: Number(process.env.IRONMIND_MAX_TOKENS ?? 700),
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: userPrompt },
          ],
        }),
      });
      const data = await res.json();
      const text = data?.choices?.[0]?.message?.content;
      return parseJsonObject(text);
    }

    if (cfg.provider === "azure_openai") {
      const url = `${cfg.endpoint.replace(/\/$/, "")}/openai/deployments/${encodeURIComponent(
        cfg.deployment
      )}/chat/completions?api-version=${encodeURIComponent(cfg.apiVersion)}`;
      const res = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": cfg.apiKey,
        },
        body: JSON.stringify({
          temperature: Number(process.env.IRONMIND_TEMPERATURE ?? 0.1),
          max_tokens: Number(process.env.IRONMIND_MAX_TOKENS ?? 700),
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: userPrompt },
          ],
        }),
      });
      const data = await res.json();
      const text = data?.choices?.[0]?.message?.content;
      return parseJsonObject(text);
    }

    if (cfg.provider === "foundry") {
      const url = normalizeFoundryChatEndpoint(cfg.endpoint);
      const res = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": cfg.apiKey,
          Authorization: `Bearer ${cfg.apiKey}`,
        },
        body: JSON.stringify({
          model: cfg.model,
          temperature: Number(process.env.IRONMIND_TEMPERATURE ?? 0.1),
          max_tokens: Number(process.env.IRONMIND_MAX_TOKENS ?? 700),
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: userPrompt },
          ],
        }),
      });
      const data = await res.json();
      const text = data?.choices?.[0]?.message?.content || data?.output_text;
      return parseJsonObject(text);
    }
  } catch (err) {
    console.error("[ironmind] ai call failed:", err?.message || err);
    return null;
  }

  return null;
}

function buildStructuredData(reportDate) {
  const maxRunHours = maxTrustedDailyRunHours();
  const activeAssets = db.prepare(`
    SELECT COUNT(*) AS c
    FROM assets
    WHERE active = 1 AND IFNULL(is_standby, 0) = 0
  `).get();

  const dailyCoverage = db.prepare(`
    SELECT COUNT(DISTINCT asset_id) AS c
    FROM daily_hours
    WHERE work_date = ? AND is_used = 1
  `).get(reportDate);

  const kpiRow = db.prepare(`
    SELECT
      COALESCE(SUM(scheduled_hours), 0) AS scheduled_hours,
      COALESCE(SUM(hours_run), 0) AS run_hours,
      COUNT(DISTINCT asset_id) AS used_assets
    FROM daily_hours
    WHERE work_date = ? AND is_used = 1
  `).get(reportDate);

  const downtimeRow = db.prepare(`
    SELECT COALESCE(SUM(l.hours_down), 0) AS downtime_hours, COUNT(*) AS log_count
    FROM breakdown_downtime_logs l
    WHERE l.log_date = ?
  `).get(reportDate);

  const planRows = db.prepare(`
    SELECT
      mp.asset_id,
      mp.service_name,
      mp.interval_hours,
      mp.last_service_hours,
      a.asset_code,
      a.category
    FROM maintenance_plans mp
    JOIN assets a ON a.id = mp.asset_id
    WHERE mp.active = 1
      AND a.active = 1
      AND IFNULL(a.is_standby, 0) = 0
  `).all();
  const sumTrustedRun = db.prepare(`
    SELECT COALESCE(SUM(dh.hours_run), 0) AS run_total
    FROM daily_hours dh
    WHERE dh.asset_id = ?
      AND dh.is_used = 1
      AND dh.hours_run > 0
      AND dh.hours_run <= ?
      AND dh.work_date <= ?
  `);
  const overdueMaintenance = planRows
    .map((p) => {
      const cap = getCategoryTrustedDailyRunHours(p.category);
      const runTotal = Number(sumTrustedRun.get(Number(p.asset_id || 0), cap, reportDate)?.run_total || 0);
      const overdue = runTotal - (Number(p.last_service_hours || 0) + Number(p.interval_hours || 0));
      return {
        asset_code: p.asset_code,
        service_name: p.service_name,
        overdue_hours: Number(overdue || 0),
      };
    })
    .filter((r) => Number(r.overdue_hours || 0) >= 0)
    .sort((a, b) => Number(b.overdue_hours || 0) - Number(a.overdue_hours || 0))
    .slice(0, 8);

  const openWorkOrders = db.prepare(`
    SELECT
      w.id,
      a.asset_code,
      w.status,
      CAST((julianday('now') - julianday(COALESCE(w.opened_at, datetime('now')))) * 24 AS INTEGER) AS age_hours
    FROM work_orders w
    JOIN assets a ON a.id = w.asset_id
    WHERE w.status IN ('open', 'assigned', 'in_progress', 'completed')
    ORDER BY age_hours DESC
    LIMIT 8
  `).all().map((r) => ({
    id: Number(r.id),
    asset_code: r.asset_code,
    status: r.status,
    age_hours: Number(r.age_hours || 0),
  }));

  const lowStockCritical = db.prepare(`
    SELECT p.part_code, p.min_stock, IFNULL(SUM(sm.quantity), 0) AS on_hand
    FROM parts p
    LEFT JOIN stock_movements sm ON sm.part_id = p.id
    WHERE p.critical = 1
    GROUP BY p.id
    HAVING on_hand < p.min_stock
    ORDER BY on_hand ASC
    LIMIT 8
  `).all().map((r) => ({
    part_code: r.part_code,
    min_stock: Number(r.min_stock || 0),
    on_hand: Number(r.on_hand || 0),
  }));

  const downtimeTop = db.prepare(`
    SELECT
      a.asset_code,
      COALESCE(SUM(l.hours_down), 0) AS downtime_hours,
      COUNT(DISTINCT l.breakdown_id) AS incidents
    FROM breakdown_downtime_logs l
    JOIN breakdowns b ON b.id = l.breakdown_id
    JOIN assets a ON a.id = b.asset_id
    WHERE l.log_date = ?
    GROUP BY a.id
    ORDER BY downtime_hours DESC
    LIMIT 6
  `).all(reportDate).map((r) => ({
    asset_code: r.asset_code,
    downtime_hours: Number(r.downtime_hours || 0),
    incidents: Number(r.incidents || 0),
  }));

  const siteOpsDateCol = hasColumn("operations_daily", "operation_date")
    ? "operation_date"
    : (hasColumn("operations_daily", "op_date") ? "op_date" : null);
  const siteOpsCount = (hasTable("operations_daily") && siteOpsDateCol)
    ? db.prepare(`
      SELECT COUNT(*) AS c
      FROM operations_daily
      WHERE ${siteOpsDateCol} = ?
    `).get(reportDate)
    : { c: 0 };
  const hasFuelBaseline = hasColumn("assets", "baseline_fuel_l_per_hour");
  const recurringFailures30d = db.prepare(`
    SELECT
      a.asset_code,
      COUNT(DISTINCT l.breakdown_id) AS incidents,
      COALESCE(SUM(l.hours_down), 0) AS downtime_hours
    FROM breakdown_downtime_logs l
    JOIN breakdowns b ON b.id = l.breakdown_id
    JOIN assets a ON a.id = b.asset_id
    WHERE l.log_date BETWEEN DATE(?, '-29 day') AND DATE(?)
      AND a.active = 1
      AND IFNULL(a.is_standby, 0) = 0
    GROUP BY a.id
    HAVING incidents >= 2
    ORDER BY incidents DESC, downtime_hours DESC
    LIMIT 8
  `).all(reportDate, reportDate).map((r) => ({
    asset_code: r.asset_code,
    incidents: Number(r.incidents || 0),
    downtime_hours: Number(r.downtime_hours || 0),
  }));
  const fuelAnomalies14d = hasFuelBaseline
    ? db.prepare(`
      SELECT asset_code, baseline_lph, actual_lph, over_pct
      FROM (
        SELECT
          a.asset_code AS asset_code,
          COALESCE(a.baseline_fuel_l_per_hour, 0) AS baseline_lph,
          CASE WHEN COALESCE(SUM(COALESCE(fl.hours_run, 0)), 0) > 0
            THEN COALESCE(SUM(COALESCE(fl.liters, 0)), 0) / SUM(COALESCE(fl.hours_run, 0))
            ELSE 0
          END AS actual_lph,
          CASE WHEN COALESCE(a.baseline_fuel_l_per_hour, 0) > 0
            THEN ((CASE WHEN COALESCE(SUM(COALESCE(fl.hours_run, 0)), 0) > 0
              THEN COALESCE(SUM(COALESCE(fl.liters, 0)), 0) / SUM(COALESCE(fl.hours_run, 0))
              ELSE 0
            END) / a.baseline_fuel_l_per_hour - 1.0) * 100.0
            ELSE 0
          END AS over_pct
        FROM fuel_logs fl
        JOIN assets a ON a.id = fl.asset_id
        WHERE fl.log_date BETWEEN DATE(?, '-13 day') AND DATE(?)
          AND COALESCE(fl.hours_run, 0) > 0
          AND a.active = 1
          AND IFNULL(a.is_standby, 0) = 0
        GROUP BY a.id
      )
      WHERE baseline_lph > 0
        AND actual_lph > baseline_lph * 1.2
      ORDER BY over_pct DESC
      LIMIT 8
    `).all(reportDate, reportDate).map((r) => ({
      asset_code: r.asset_code,
      baseline_lph: Number(r.baseline_lph || 0),
      actual_lph: Number(r.actual_lph || 0),
      over_pct: Number(r.over_pct || 0),
    }))
    : [];

  const scheduled = Number(kpiRow?.scheduled_hours || 0);
  const run = Number(kpiRow?.run_hours || 0);
  const downtime = Number(downtimeRow?.downtime_hours || 0);
  const available = Math.max(0, scheduled - downtime);
  const availabilityPct = scheduled > 0 ? (available / scheduled) * 100 : 0;
  const utilizationPct = scheduled > 0 ? (run / scheduled) * 100 : 0;
  const assetRiskSignals = computeAssetRiskSignals(reportDate);
  const topRiskAssets = assetRiskSignals
    .filter((r) => Number(r.risk_score || 0) >= 35)
    .slice(0, 8)
    .map((r) => ({
      asset_code: r.asset_code,
      risk_score: Number(r.risk_score || 0),
      confidence: Number(r.confidence || 0),
      reasons: r.reasons || [],
      incidents_30d: Number(r.incidents_30d || 0),
      downtime_30d: Number(r.downtime_30d || 0),
      overdue_hours: Number(r.overdue_hours || 0),
      fuel_over_pct: Number(r.fuel_over_pct || 0),
    }));

  const recentRows = db.prepare(`
    SELECT
      a.asset_code AS asset_code,
      a.category AS category,
      dh.work_date AS work_date,
      dh.hours_run AS hours_run,
      dh.opening_hours AS opening_hours,
      dh.closing_hours AS closing_hours
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.is_used = 1
      AND dh.work_date BETWEEN date(?, '-89 day') AND ?
      AND a.active = 1
      AND IFNULL(a.is_standby, 0) = 0
    ORDER BY dh.work_date DESC, dh.hours_run DESC
    LIMIT 500
  `).all(reportDate, reportDate);
  const dailyInputAnomalies = recentRows.map((r) => {
    const cap = getCategoryTrustedDailyRunHours(r.category);
    let reason = "";
    if (r.hours_run != null && Number(r.hours_run) > cap) reason = `hours_run exceeds trusted max (${cap}h for category)`;
    else if (r.opening_hours != null && r.closing_hours != null && Number(r.closing_hours) < Number(r.opening_hours)) {
      reason = "invalid meter (close < open)";
    }
    if (!reason) return null;
    return {
      asset_code: r.asset_code,
      work_date: r.work_date,
      hours_run: r.hours_run == null ? null : Number(r.hours_run),
      opening_hours: r.opening_hours == null ? null : Number(r.opening_hours),
      closing_hours: r.closing_hours == null ? null : Number(r.closing_hours),
      reason,
    };
  }).filter(Boolean).slice(0, 25).map((r) => ({
    asset_code: r.asset_code,
    work_date: r.work_date,
    hours_run: r.hours_run == null ? null : Number(r.hours_run),
    opening_hours: r.opening_hours == null ? null : Number(r.opening_hours),
    closing_hours: r.closing_hours == null ? null : Number(r.closing_hours),
    reason: r.reason,
  }));

  return {
    report_date: reportDate,
    kpi: {
      used_assets: Number(kpiRow?.used_assets || 0),
      scheduled_hours: Number(scheduled.toFixed(2)),
      run_hours: Number(run.toFixed(2)),
      downtime_hours: Number(downtime.toFixed(2)),
      availability_pct: Number(availabilityPct.toFixed(2)),
      utilization_pct: Number(utilizationPct.toFixed(2)),
    },
    overdueMaintenance,
    openWorkOrders,
    lowStockCritical,
    downtimeTop,
    recurringFailures30d,
    fuelAnomalies14d,
    topRiskAssets,
    daily_input_anomalies: dailyInputAnomalies,
    trusted_daily_run_hours_max: maxRunHours,
    dataCoverage: {
      active_assets: Number(activeAssets?.c || 0),
      assets_with_daily_entry: Number(dailyCoverage?.c || 0),
      breakdown_logs: Number(downtimeRow?.log_count || 0),
      operations_daily_rows: Number(siteOpsCount?.c || 0),
    },
  };
}

export async function generateIronmindReport({
  reportDate,
  reportType = "daily_admin",
  force = false,
  contextNotes = "",
  detailMode = false,
}) {
  ensureIronmindTable();

  const targetDate = String(reportDate || ymdOffsetFromNow(0)).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(targetDate)) {
    throw new Error("reportDate must be YYYY-MM-DD");
  }

  const existing = db.prepare(`
    SELECT id, report_date, report_type, summary, created_at
    FROM ironmind_reports
    WHERE report_date = ? AND report_type = ?
    LIMIT 1
  `).get(targetDate, reportType);

  if (existing && !force) {
    return {
      id: Number(existing.id),
      report_date: existing.report_date,
      report_type: existing.report_type,
      summary: existing.summary,
      created_at: existing.created_at,
      created: false,
    };
  }

  const structuredData = buildStructuredData(targetDate);
  for (const r of (structuredData.topRiskAssets || [])) {
    db.prepare(`
      INSERT INTO ironmind_asset_risk_snapshots
        (report_date, asset_code, risk_score, confidence, reasons_json, features_json, created_at)
      VALUES (?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(report_date, asset_code)
      DO UPDATE SET
        risk_score = excluded.risk_score,
        confidence = excluded.confidence,
        reasons_json = excluded.reasons_json,
        features_json = excluded.features_json,
        created_at = excluded.created_at
    `).run(
      targetDate,
      String(r.asset_code || ""),
      Number(r.risk_score || 0),
      Number(r.confidence || 0),
      JSON.stringify(r.reasons || []),
      JSON.stringify({
        incidents_30d: Number(r.incidents_30d || 0),
        downtime_30d: Number(r.downtime_30d || 0),
        overdue_hours: Number(r.overdue_hours || 0),
        fuel_over_pct: Number(r.fuel_over_pct || 0),
      }),
      nowIso()
    );
  }
  const aiJson = await callIronmindAi(structuredData, { contextNotes, detailMode });

  const parsed = aiJson && typeof aiJson === "object"
    ? {
        repairsNeeded: Array.isArray(aiJson.repairs_needed) ? aiJson.repairs_needed : [],
        operationalRisks: Array.isArray(aiJson.operational_risks) ? aiJson.operational_risks : [],
        suggestions: Array.isArray(aiJson.suggestions) ? aiJson.suggestions : [],
        dataGaps: Array.isArray(aiJson.data_gaps) ? aiJson.data_gaps : [],
      }
    : buildFallbackInsight(structuredData);

  const anomalyLines = buildDailyInputAnomalyLines(
    structuredData.daily_input_anomalies,
    structuredData.trusted_daily_run_hours_max
  );
  const summary = appendDataAnomaliesSection(toIronmindFormat(parsed), anomalyLines);

  db.prepare(`
    INSERT INTO ironmind_reports (report_date, report_type, summary, created_at)
    VALUES (?, ?, ?, ?)
    ON CONFLICT(report_date, report_type)
    DO UPDATE SET
      summary = excluded.summary,
      created_at = excluded.created_at
  `).run(targetDate, reportType, summary, nowIso());

  const row = db.prepare(`
    SELECT id, report_date, report_type, summary, created_at
    FROM ironmind_reports
    WHERE report_date = ? AND report_type = ?
    LIMIT 1
  `).get(targetDate, reportType);

  return {
    id: Number(row.id),
    report_date: row.report_date,
    report_type: row.report_type,
    summary: row.summary,
    created_at: row.created_at,
    created: true,
  };
}

export function getLatestIronmindReport(reportType = "daily_admin") {
  ensureIronmindTable();
  const row = db.prepare(`
    SELECT id, report_date, report_type, summary, created_at
    FROM ironmind_reports
    WHERE report_type = ?
    ORDER BY report_date DESC, id DESC
    LIMIT 1
  `).get(reportType);

  if (!row) return null;
  return {
    id: Number(row.id),
    report_date: row.report_date,
    report_type: row.report_type,
    summary: row.summary,
    created_at: row.created_at,
  };
}

export function getIronmindHistory({ reportType = "daily_admin", limit = 7 } = {}) {
  ensureIronmindTable();
  const cappedLimit = Math.max(1, Math.min(60, Number(limit || 7)));
  const rows = db.prepare(`
    SELECT id, report_date, report_type, summary, created_at
    FROM ironmind_reports
    WHERE report_type = ?
    ORDER BY report_date DESC, id DESC
    LIMIT ?
  `).all(reportType, cappedLimit);

  return (rows || []).map((row) => ({
    id: Number(row.id),
    report_date: row.report_date,
    report_type: row.report_type,
    summary: row.summary,
    created_at: row.created_at,
  }));
}

export async function runIronmindAutoScheduler(log = console) {
  ensureIronmindTable();

  const enabled = String(process.env.IRONMIND_AUTO_ENABLED || "1").trim().toLowerCase();
  if (["0", "false", "no", "off"].includes(enabled)) {
    log.info?.("[ironmind] auto scheduler disabled");
    return null;
  }

  const runHour = Math.max(0, Math.min(23, Number(process.env.IRONMIND_RUN_HOUR ?? 6)));
  const runMinute = Math.max(0, Math.min(59, Number(process.env.IRONMIND_RUN_MINUTE ?? 0)));
  const dayOffset = Math.max(0, Number(process.env.IRONMIND_TARGET_OFFSET_DAYS ?? 1));
  const reportType = String(process.env.IRONMIND_REPORT_TYPE || "daily_admin").trim() || "daily_admin";
  const autoModeRaw = String(process.env.IRONMIND_AUTO_MODE || "daily").trim().toLowerCase();
  const autoMode = autoModeRaw === "interval" ? "interval" : "daily";
  const intervalMinutes = Math.max(5, Math.min(1440, Number(process.env.IRONMIND_RUN_INTERVAL_MINUTES ?? 180)));
  const forceRefresh = ["1", "true", "yes", "on"].includes(
    String(process.env.IRONMIND_AUTO_FORCE_REFRESH || "0").trim().toLowerCase()
  );

  let lastMinuteKey = "";
  let lastIntervalRunAtMs = 0;

  const maybeRun = async (reason = "tick") => {
    const now = new Date();
    const nowMs = Date.now();
    const minuteKey = `${now.getFullYear()}-${now.getMonth() + 1}-${now.getDate()} ${now.getHours()}:${now.getMinutes()}`;
    const shouldRunByClock = now.getHours() === runHour && now.getMinutes() === runMinute;
    const shouldRunByInterval =
      autoMode === "interval" &&
      (reason === "startup" || (nowMs - lastIntervalRunAtMs) >= (intervalMinutes * 60 * 1000));

    if (autoMode === "daily") {
      if (!shouldRunByClock && reason !== "startup") return;
      if (shouldRunByClock && lastMinuteKey === minuteKey) return;
      if (shouldRunByClock) lastMinuteKey = minuteKey;
    } else if (!shouldRunByInterval) {
      return;
    }

    const targetDate = ymdOffsetFromNow(dayOffset);

    try {
      const result = await generateIronmindReport({
        reportDate: targetDate,
        reportType,
        force: forceRefresh,
      });
      if (autoMode === "interval") lastIntervalRunAtMs = nowMs;
      log.info?.(`[ironmind] ${reason} report ready for ${result.report_date} (id=${result.id}, created=${result.created})`);
    } catch (err) {
      log.error?.(`[ironmind] auto run failed: ${err?.message || err}`);
    }
  };

  await maybeRun("startup");
  const timer = setInterval(() => {
    maybeRun("tick");
  }, 30 * 1000);

  if (autoMode === "interval") {
    log.info?.(
      `[ironmind] auto scheduler enabled in interval mode (${intervalMinutes} min, force_refresh=${forceRefresh ? "on" : "off"}), target offset ${dayOffset} day(s)`
    );
  } else {
    log.info?.(
      `[ironmind] auto scheduler enabled at ${String(runHour).padStart(2, "0")}:${String(runMinute).padStart(2, "0")}, target offset ${dayOffset} day(s)`
    );
  }
  return timer;
}
