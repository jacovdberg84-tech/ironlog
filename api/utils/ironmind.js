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

  suggestions.push("Close or update aged work orders before next shift handover.");
  suggestions.push("Prioritize overdue preventive maintenance on highest-utilization assets.");
  suggestions.push("Reconcile critical spares below minimum and confirm delivery ETA.");

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

async function callIronmindAi(structuredData) {
  const cfg = getAiConfig();
  if (!cfg.provider) return null;

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
    "- Each key must be an array of short strings.",
  ].join("\n");

  const userPrompt = `Structured plant data for report date ${structuredData.report_date}:\n${JSON.stringify(structuredData, null, 2)}`;

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

  const overdueMaintenance = db.prepare(`
    SELECT
      a.asset_code,
      mp.service_name,
      (COALESCE((
        SELECT SUM(dh.hours_run)
        FROM daily_hours dh
        WHERE dh.asset_id = mp.asset_id
          AND dh.is_used = 1
          AND dh.hours_run > 0
          AND dh.work_date <= ?
      ), 0) - (mp.last_service_hours + mp.interval_hours)) AS overdue_hours
    FROM maintenance_plans mp
    JOIN assets a ON a.id = mp.asset_id
    WHERE mp.active = 1 AND a.active = 1 AND IFNULL(a.is_standby, 0) = 0
    HAVING overdue_hours >= 0
    ORDER BY overdue_hours DESC
    LIMIT 8
  `).all(reportDate).map((r) => ({
    asset_code: r.asset_code,
    service_name: r.service_name,
    overdue_hours: Number(r.overdue_hours || 0),
  }));

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

  const siteOpsCount = db.prepare(`
    SELECT COUNT(*) AS c
    FROM operations_daily
    WHERE operation_date = ?
  `).get(reportDate);

  const scheduled = Number(kpiRow?.scheduled_hours || 0);
  const run = Number(kpiRow?.run_hours || 0);
  const downtime = Number(downtimeRow?.downtime_hours || 0);
  const available = Math.max(0, scheduled - downtime);
  const availabilityPct = scheduled > 0 ? (available / scheduled) * 100 : 0;
  const utilizationPct = available > 0 ? (run / available) * 100 : 0;

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
    dataCoverage: {
      active_assets: Number(activeAssets?.c || 0),
      assets_with_daily_entry: Number(dailyCoverage?.c || 0),
      breakdown_logs: Number(downtimeRow?.log_count || 0),
      operations_daily_rows: Number(siteOpsCount?.c || 0),
    },
  };
}

export async function generateIronmindReport({ reportDate, reportType = "daily_admin", force = false }) {
  ensureIronmindTable();

  const targetDate = String(reportDate || ymdOffsetFromNow(1)).trim();
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
  const aiJson = await callIronmindAi(structuredData);

  const parsed = aiJson && typeof aiJson === "object"
    ? {
        repairsNeeded: Array.isArray(aiJson.repairs_needed) ? aiJson.repairs_needed : [],
        operationalRisks: Array.isArray(aiJson.operational_risks) ? aiJson.operational_risks : [],
        suggestions: Array.isArray(aiJson.suggestions) ? aiJson.suggestions : [],
        dataGaps: Array.isArray(aiJson.data_gaps) ? aiJson.data_gaps : [],
      }
    : buildFallbackInsight(structuredData);

  const summary = toIronmindFormat(parsed);

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

  let lastMinuteKey = "";

  const maybeRun = async (reason = "tick") => {
    const now = new Date();
    const minuteKey = `${now.getFullYear()}-${now.getMonth() + 1}-${now.getDate()} ${now.getHours()}:${now.getMinutes()}`;
    const shouldRunByClock = now.getHours() === runHour && now.getMinutes() === runMinute;

    if (!shouldRunByClock && reason !== "startup") return;
    if (shouldRunByClock && lastMinuteKey === minuteKey) return;
    if (shouldRunByClock) lastMinuteKey = minuteKey;

    const targetDate = ymdOffsetFromNow(dayOffset);

    try {
      const result = await generateIronmindReport({
        reportDate: targetDate,
        reportType,
        force: false,
      });
      log.info?.(`[ironmind] ${reason} report ready for ${result.report_date} (id=${result.id}, created=${result.created})`);
    } catch (err) {
      log.error?.(`[ironmind] auto run failed: ${err?.message || err}`);
    }
  };

  await maybeRun("startup");
  const timer = setInterval(() => {
    maybeRun("tick");
  }, 30 * 1000);

  log.info?.(`[ironmind] auto scheduler enabled at ${String(runHour).padStart(2, "0")}:${String(runMinute).padStart(2, "0")}, target offset ${dayOffset} day(s)`);
  return timer;
}
