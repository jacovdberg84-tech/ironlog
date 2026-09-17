import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";

const source = readFileSync(new URL("../../web/maintenance.js", import.meta.url), "utf8");
const functions = source.slice(
  source.indexOf("async function downloadAssetKpiExport("),
  source.indexOf("async function exportExecutivePackFromAssetKpi(")
);

function harness(status = 200) {
  const calls = { alerts: [], downloads: [], messages: [], requests: [], timers: [] };
  const message = {
    set className(value) { calls.messages.push({ type: "class", value }); },
    set textContent(value) { calls.messages.push({ type: "text", value }); },
  };
  const context = vm.createContext({
    API: "/api",
    URLSearchParams,
    akpLastMeta: { start: "2026-08-01", end: "2026-08-31", sched: 10, asset_codes: ["A300AM"] },
    akpSelectedAssetCodes: () => [],
    authHeaders: () => ({ Authorization: "Bearer kpi-session" }),
    fetch: async (url, options) => {
      calls.requests.push({ url, options });
      return {
        ok: status === 200,
        status,
        text: async () => JSON.stringify({ error: "Login required" }),
        blob: async () => ({ type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      };
    },
    URL: { createObjectURL: () => "blob:asset-kpi", revokeObjectURL() {} },
    document: {
      getElementById: (id) => (id === "akpMsg" ? message : null),
      body: { appendChild() {} },
      createElement: () => ({ click() { calls.downloads.push(this.download); }, remove() {} }),
    },
    setTimeout: (fn) => calls.timers.push(fn),
    alert: (message) => calls.alerts.push(message),
  });
  vm.runInContext(functions, context);
  return { calls, context };
}

test("Asset KPI Excel export uses the current login token and saves a named download", async () => {
  const { calls, context } = harness();
  await context.exportAssetKpiToExcel();
  assert.equal(calls.requests[0].options.headers.Authorization, "Bearer kpi-session");
  assert.match(calls.requests[0].url, /^\/api\/dashboard\/asset-kpi\.xlsx\?/);
  assert.match(calls.requests[0].url, /asset_codes=A300AM/);
  assert.deepEqual(calls.downloads, ["IRONLOG_Asset_KPI_2026-08-01_to_2026-08-31.xlsx"]);
});

test("Asset KPI Excel export reports an expired login instead of opening an unauthenticated tab", async () => {
  const { calls, context } = harness(401);
  await context.exportAssetKpiToExcel();
  assert.equal(calls.requests[0].options.headers.Authorization, "Bearer kpi-session");
  assert.deepEqual(calls.downloads, []);
  assert.match(calls.alerts[0], /Login required/);
});
