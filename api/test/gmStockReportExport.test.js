import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";

const appSource = readFileSync(new URL("../../web/app.js", import.meta.url), "utf8");
const functions = appSource.slice(
  appSource.indexOf("function ensureGmStockReportDate()"),
  appSource.indexOf("\nlet storesPartOrdersCache")
);

function harness() {
  const elements = {
    gmStockReportPeriod: { value: "monthly" },
    gmStockReportDate: { value: "" },
    gmStockReportDateLabel: { textContent: "" },
    gmStockReportHelp: { textContent: "" },
  };
  const calls = { downloads: [], statuses: [], alerts: [] };
  const context = vm.createContext({
    API: "/api",
    URLSearchParams,
    qs: (id) => elements[id] || null,
    todayLocalYmd: () => "2026-09-28",
    setStatus: (message) => calls.statuses.push(message),
    downloadAuthedFile: async (url, filename) => {
      calls.downloads.push({ url, filename });
      return true;
    },
    alert: (message) => calls.alerts.push(message),
  });
  vm.runInContext(functions, context);
  return { elements, calls, context };
}

test("GM stock Excel uses the authenticated download helper and selected monthly period", async () => {
  const { elements, calls, context } = harness();
  await context.downloadGmStockReportXlsx();

  assert.equal(elements.gmStockReportDate.value, "2026-09-28");
  assert.equal(calls.downloads[0].url, "/api/api/stock/gm-stock-report.xlsx?period=monthly&report_date=2026-09-28");
  assert.equal(calls.downloads[0].filename, "IRONLOG_Monthly_Stock_Report_2026-09-28.xlsx");
  assert.deepEqual(calls.statuses, ["Preparing GM stock Excel...", "GM stock Excel downloaded."]);
  assert.deepEqual(calls.alerts, []);
});

test("GM stock controls describe the weekly seven-day report", () => {
  const { elements, context } = harness();
  elements.gmStockReportPeriod.value = "weekly";
  context.updateGmStockReportHelp();

  assert.equal(elements.gmStockReportDateLabel.textContent, "Week ending");
  assert.equal(elements.gmStockReportHelp.textContent, "Weekly uses the seven days ending on the selected date.");
});

test("Stores page exposes the GM stock report download control", () => {
  const indexSource = readFileSync(new URL("../../web/index.html", import.meta.url), "utf8");
  const routeSource = readFileSync(new URL("../routes/stock.routes.js", import.meta.url), "utf8");
  assert.match(indexSource, /id="downloadGmStockReportXlsx"/);
  assert.match(routeSource, /app\.get\("\/gm-stock-report\.xlsx"/);
});
