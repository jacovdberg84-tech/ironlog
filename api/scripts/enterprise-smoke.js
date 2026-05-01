#!/usr/bin/env node

/**
 * Enterprise rollout smoke test:
 *   entity (company/site/currency/tax)
 *   -> integration hub (connection, enqueue, run, monitor, retry, dead-letter)
 *   -> governance policies + evaluate
 *   -> evidence packs (audit/procurement/hse)
 *   -> executive command-center + board pack
 */

const API_BASE = process.env.API_BASE || "http://localhost:3001";
const SITE_CODE = process.env.SITE_CODE || "main";

function headers() {
  return {
    "Content-Type": "application/json",
    "x-user-role": "admin",
    "x-user-name": "enterprise-smoke",
    "x-user-roles": "admin,supervisor,executive,finance,procurement",
    "x-site-code": SITE_CODE,
  };
}

async function call(method, path, body) {
  const res = await fetch(`${API_BASE}${path}`, {
    method, headers: headers(),
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let json = {};
  try { json = text ? JSON.parse(text) : {}; } catch { json = { raw: text }; }
  if (!res.ok) throw new Error(`${method} ${path} -> HTTP ${res.status}: ${json?.error || text}`);
  return json;
}

function log(label, payload) {
  const short = typeof payload === "object" ? JSON.stringify(payload).slice(0, 200) : String(payload);
  console.log(`\n[OK] ${label}\n     ${short}${short.length >= 200 ? "..." : ""}`);
}

function currentPeriod() { return new Date().toISOString().slice(0, 7); }

async function run() {
  const period = currentPeriod();
  console.log(`Enterprise smoke API=${API_BASE} period=${period}`);

  await call("POST", "/api/entity/companies", { company_code: "CO-SMOKE", company_name: "Smoke Co", base_currency: "USD", reporting_currency: "USD" });
  await call("POST", "/api/entity/sites", { site_code: "SITE-SMOKE", company_code: "CO-SMOKE", site_name: "Smoke Site", local_currency: "USD" });
  log("company + site", { company: "CO-SMOKE", site: "SITE-SMOKE" });

  const tree = await call("GET", "/api/entity/tree");
  log("entity tree", { companies: tree.tree?.length });

  await call("POST", "/api/entity/currency/rates/upsert", { rows: [{ from_currency: "EUR", to_currency: "USD", rate: 1.07, effective_date: `${period}-01`, source: "smoke" }] });
  log("currency rate upserted", { pair: "EUR->USD" });

  const conv = await call("GET", `/api/entity/currency/convert?from=EUR&to=USD&amount=100`);
  log("currency convert", conv);

  await call("POST", "/api/entity/tax/profiles/upsert", { tax_code: "VAT-SMOKE", label: "Smoke VAT", rate_pct: 15, region: "ZA" });
  log("tax profile upserted", { code: "VAT-SMOKE" });

  await call("POST", "/api/integrations/connections/upsert", { connection_code: "ERP-SMOKE", connector_key: "erp_journal_export", label: "Smoke ERP" });
  log("integration connection", { code: "ERP-SMOKE" });

  const connList = await call("GET", "/api/integrations/connections");
  log("integration connections", { rows: connList.rows?.length });

  const enqFail = await call("POST", "/api/integrations/jobs/enqueue", {
    connection_code: "ERP-SMOKE",
    connector_key: "erp_journal_export",
    payload: { run_id: 999999 },
    idempotency_key: `smoke-erp-${Date.now()}`,
    max_attempts: 2,
  });
  log("enqueued fail job", enqFail);
  const runFail = await fetch(`${API_BASE}/api/integrations/jobs/${enqFail.id}/run`, { method: "POST", headers: headers() });
  const runFailJson = await runFail.json().catch(() => ({}));
  log("ran fail job", { status: runFail.status, body: runFailJson });

  const monitor = await call("GET", "/api/integrations/monitoring/summary");
  log("integration monitor", monitor);

  const dl = await call("GET", "/api/integrations/dead-letter");
  log("dead letter", { rows: dl.rows?.length });

  const policies = await call("GET", "/api/governance/policies");
  log("governance policies", { rows: policies.rows?.length });
  const evalRes = await call("POST", "/api/governance/evaluate", { action: "finance.journals.reverse", username: "finance-smoke" });
  log("governance evaluate", evalRes);
  const violations = await call("GET", "/api/governance/violations");
  log("governance violations", { rows: violations.rows?.length });

  const audit = await call("POST", "/api/executive/evidence-packs/build", { pack_type: "audit", period });
  log("evidence audit pack", audit);
  const proc = await call("POST", "/api/executive/evidence-packs/build", { pack_type: "procurement", period });
  log("evidence procurement pack", proc);
  const hse = await call("POST", "/api/executive/evidence-packs/build", { pack_type: "hse", period });
  log("evidence hse pack", hse);
  const detail = await call("GET", `/api/executive/evidence-packs/${audit.id}`);
  log("pack detail integrity_ok", { ok: detail.pack?.integrity_ok });

  const cmd = await call("GET", `/api/executive/command-center?period=${encodeURIComponent(period)}`);
  log("command center", { ops: cmd.operational, gov: cmd.governance, integ: cmd.integrations });

  const board = await call("GET", `/api/executive/board-pack?period=${encodeURIComponent(period)}`);
  log("board pack", { bullets: board.narrative?.bullets?.length });

  console.log("\nEnterprise smoke OK");
}

run().catch((e) => {
  console.error(`\n[FAIL] ${e.message}`);
  process.exit(1);
});
