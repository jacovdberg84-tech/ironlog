import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";

const source = readFileSync(new URL("../../web/maintenance.js", import.meta.url), "utf8");
const functions = source.slice(
  source.indexOf("async function downloadProtectedXlsxFile("),
  source.indexOf("function mpWeekRangeLabel(")
);

function harness(status = 200) {
  const calls = { alerts: [], downloads: [], requests: [], timers: [] };
  const context = vm.createContext({
    API: "/api",
    getDueThresholdHours: () => 50,
    authHeaders: () => ({ Authorization: "Bearer maintenance-session" }),
    fetch: async (url, options) => {
      calls.requests.push({ url, options });
      return {
        ok: status === 200,
        status,
        text: async () => JSON.stringify({ error: "Login required" }),
        blob: async () => ({ type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      };
    },
    URL: { createObjectURL: () => "blob:maintenance-plans", revokeObjectURL() {} },
    document: {
      body: { appendChild() {} },
      createElement: () => ({ click() { calls.downloads.push(this.download); }, remove() {} }),
    },
    Date,
    setTimeout: (fn) => calls.timers.push(fn),
    alert: (message) => calls.alerts.push(message),
  });
  vm.runInContext(functions, context);
  return { calls, context };
}

test("Maintenance-plan Excel export uses the current login token and selected threshold", async () => {
  const { calls, context } = harness();
  await context.downloadMaintenancePlansXlsx();

  assert.equal(calls.requests[0].options.headers.Authorization, "Bearer maintenance-session");
  assert.equal(calls.requests[0].options.cache, "no-store");
  assert.equal(calls.requests[0].url, "/api/maintenance/plans.xlsx?near_due_hours=50");
  assert.match(calls.downloads[0], /^IRONLOG_Maintenance_Plans_\d{4}-\d{2}-\d{2}\.xlsx$/);
});

test("Maintenance-plan Excel export explains an expired login", async () => {
  const { calls, context } = harness(401);
  await context.downloadMaintenancePlansXlsx();

  assert.deepEqual(calls.downloads, []);
  assert.match(calls.alerts[0], /session has expired/i);
});
