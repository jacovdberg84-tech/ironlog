import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";

const source = readFileSync(new URL("../../web/maintenance.js", import.meta.url), "utf8");
const functions = source.slice(
  source.indexOf("function exportReliabilityToExcel()"),
  source.indexOf("async function loadAssetKpiWeekly()")
);

test("executive reliability export carries the selected period and current login", async () => {
  const calls = { requests: [], downloads: [] };
  const context = vm.createContext({
    API: "/api",
    URLSearchParams,
    relLastMeta: { start: "2026-08-01", end: "2026-08-31", scheduled: 11, category: "Production fleet", asset_ids: "12,23" },
    downloadProtectedXlsxFile: async (url, filename) => {
      calls.requests.push(url);
      calls.downloads.push(filename);
    },
    alert: (message) => { throw new Error(message); },
  });
  vm.runInContext(functions, context);
  await context.exportReliabilityExecutiveToExcel();
  assert.match(calls.requests[0], /^\/api\/maintenance\/reliability-executive\.xlsx\?/);
  assert.match(calls.requests[0], /start=2026-08-01/);
  assert.match(calls.requests[0], /scheduled=11/);
  assert.match(calls.requests[0], /asset_ids=12%2C23/);
  assert.deepEqual(calls.downloads, ["IRONLOG_MTBF_LTTR_Executive_2026-08-01_to_2026-08-31.xlsx"]);
});
