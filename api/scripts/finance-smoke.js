#!/usr/bin/env node

/**
 * End-to-end smoke test for the Finance Integration flow:
 *   summarize -> list -> detail -> mark exported -> mark posted
 *   period checklist/load -> update item -> close (force) -> reopen
 *   budget upsert -> budget vs actual -> forecast rebuild -> forecast load
 *   SSOT report
 *
 * Usage (PowerShell):
 *   $env:API_BASE="http://localhost:3001"; node api/scripts/finance-smoke.js
 */

const API_BASE = process.env.API_BASE || "http://localhost:3001";
const SITE_CODE = process.env.SITE_CODE || "main";

function headers() {
  return {
    "Content-Type": "application/json",
    "x-user-role": "admin",
    "x-user-name": "finance-smoke",
    "x-user-roles": "admin,supervisor,plant_manager,procurement,stores",
    "x-site-code": SITE_CODE,
  };
}

async function call(method, path, body) {
  const res = await fetch(`${API_BASE}${path}`, {
    method,
    headers: headers(),
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let json = {};
  try { json = text ? JSON.parse(text) : {}; } catch { json = { raw: text }; }
  if (!res.ok) {
    const msg = json?.error || `HTTP ${res.status}`;
    throw new Error(`${method} ${path} failed: ${msg}`);
  }
  return json;
}

function log(label, payload) {
  const short = typeof payload === "object" ? JSON.stringify(payload).slice(0, 200) : String(payload);
  console.log(`\n[OK] ${label}\n     ${short}${short.length >= 200 ? "..." : ""}`);
}

function today() { return new Date().toISOString().slice(0, 10); }
function minus(days) {
  const d = new Date(); d.setUTCDate(d.getUTCDate() - days);
  return d.toISOString().slice(0, 10);
}
function currentPeriod() { return today().slice(0, 7); }
function addPeriod(period, delta) {
  const [y, m] = period.split("-").map((x) => Number(x));
  const d = new Date(Date.UTC(y, m - 1 + delta, 1));
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
}

async function run() {
  const period = currentPeriod();
  const startDate = minus(30);
  const endDate = today();

  console.log(`Finance smoke: API=${API_BASE} | site=${SITE_CODE} | period=${period} | range ${startDate}..${endDate}`);

  const buildRun = await call("POST", "/api/procurement/journals/summarize", {
    start: startDate,
    end: endDate,
    categories: ["parts", "labor", "downtime", "fuel", "lube"],
    default_cost_center_code: "SMOKE-CC",
    currency: "USD",
    notes: "smoke test",
  });
  log("summarize built", { run_number: buildRun?.run?.run_number, lines: buildRun?.run?.line_count, debit: buildRun?.run?.total_debit, credit: buildRun?.run?.total_credit, balanced: buildRun?.balanced });
  const runId = Number(buildRun?.run?.id || 0);
  if (!runId) throw new Error("no run id returned");

  const list = await call("GET", `/api/procurement/journals/runs`);
  log("runs list", { rows: list?.rows?.length });

  const detail = await call("GET", `/api/procurement/journals/runs/${runId}`);
  log("run detail", { by_category: detail?.by_category?.length });

  await call("POST", `/api/procurement/journals/runs/${runId}/mark-exported`, {});
  log("marked exported", { runId });

  const posted = await call("POST", `/api/procurement/journals/runs/${runId}/mark-posted`, { posted_reference: `SMOKE-REF-${runId}` });
  log("marked posted", posted);

  const checklist = await call("GET", `/api/finance/periods/${period}/checklist`);
  log("checklist loaded", { items: checklist?.items?.length, lock: checklist?.lock?.status });

  if (checklist?.items?.length) {
    const firstCode = checklist.items[0].code;
    await call("POST", `/api/finance/periods/${period}/checklist/${firstCode}`, { status: "done", notes: "smoke" });
    log("checklist item done", { firstCode });
  }

  try {
    const close = await call("POST", `/api/finance/periods/${period}/close`, { force: true });
    log("period closed (force)", close);
  } catch (e) {
    console.log(`[WARN] close failed: ${e.message}`);
  }

  try {
    const reopen = await call("POST", `/api/finance/periods/${period}/reopen`, { reason: "smoke reopen" });
    log("period reopened", reopen);
  } catch (e) {
    console.log(`[WARN] reopen failed: ${e.message}`);
  }

  const saveBudgets = await call("POST", `/api/finance/budgets/upsert`, {
    rows: [
      { period, site_code: "", cost_center_code: "SMOKE-CC", equipment_type: "", category: "parts", budget_amount: 1234.56 },
      { period, site_code: "", cost_center_code: "SMOKE-CC", equipment_type: "", category: "labor", budget_amount: 987.65 },
    ],
  });
  log("budgets saved", saveBudgets);

  const bva = await call("GET", `/api/finance/budgets-vs-actual?period=${encodeURIComponent(period)}&dimension=cost_center_code`);
  log("budget vs actual", bva?.total);

  const forecastStart = addPeriod(period, 1);
  const fcast = await call("POST", `/api/finance/forecast/rebuild`, { start_period: forecastStart, months: 3 });
  log("forecast rebuilt", fcast);

  const fcastLoaded = await call("GET", `/api/finance/forecast?batch_id=${encodeURIComponent(fcast.batch_id)}`);
  log("forecast load", { rows: fcastLoaded?.rows?.length, totals: fcastLoaded?.totals_by_period?.length });

  const ssot = await call("GET", `/api/finance/reports/ssot?period=${encodeURIComponent(period)}`);
  log("ssot report", {
    availability: ssot?.kpi?.availability,
    utilization: ssot?.kpi?.utilization,
    mtbf: ssot?.kpi?.mtbf,
    mttr: ssot?.kpi?.mttr,
    cost_per_asset_hour: ssot?.kpi?.cost_per_asset_hour,
    actuals: ssot?.actuals?.length,
    budgets: ssot?.budgets?.length,
  });

  const defs = await call("GET", `/api/finance/kpis/definitions`);
  log("kpi definitions", { count: defs?.kpis?.length });

  console.log("\nFinance smoke OK");
}

run().catch((e) => {
  console.error(`\n[FAIL] ${e.message}`);
  process.exit(1);
});
