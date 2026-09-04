import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";

const source = readFileSync(new URL("../../web/maintenance.js", import.meta.url), "utf8");
const helper = source.slice(
  source.indexOf("async function openProtectedPdf("),
  source.indexOf("function renderMaintenancePlanningKpis(")
);

function harness(status = 200) {
  const calls = { alerts: [], requests: [], downloads: [], previews: [], timers: [] };
  const preview = {
    closed: false,
    opener: {},
    location: { replace: (url) => calls.previews.push(url) },
    close() { this.closed = true; },
  };
  const context = vm.createContext({
    authHeaders: () => ({ Authorization: "Bearer artisan-session" }),
    fetch: async (url, options) => {
      calls.requests.push({ url, options });
      return {
        ok: status === 200,
        status,
        json: async () => ({ error: "Login required" }),
        blob: async () => ({ type: "application/pdf" }),
      };
    },
    window: { open: () => preview },
    document: {
      body: { appendChild() {} },
      createElement: () => ({ click() { calls.downloads.push(this.download); }, remove() {} }),
    },
    URL: { createObjectURL: () => "blob:artisan-pdf", revokeObjectURL() {} },
    setTimeout: (fn) => calls.timers.push(fn),
    alert: (message) => calls.alerts.push(message),
  });
  vm.runInContext(helper, context);
  return { context, calls, preview };
}

test("Artisan inspection PDF preview uses the current login token", async () => {
  const { context, calls } = harness();
  await context.openProtectedPdf("/api/reports/artisan-inspection-form.pdf?date=2026-09-04");
  assert.equal(calls.requests[0].options.headers.Authorization, "Bearer artisan-session");
  assert.deepEqual(calls.previews, ["blob:artisan-pdf"]);
  assert.equal(calls.timers.length, 1);
});

test("Artisan inspection PDF download uses the supplied filename", async () => {
  const { context, calls } = harness();
  await context.openProtectedPdf("/api/reports/artisan-inspection/42.pdf", {
    download: true,
    filename: "IRONLOG_Artisan_Inspection_42.pdf",
  });
  assert.deepEqual(calls.downloads, ["IRONLOG_Artisan_Inspection_42.pdf"]);
});

test("expired login closes the Artisan PDF preview", async () => {
  const { context, calls, preview } = harness(401);
  await context.openProtectedPdf("/api/reports/artisan-inspection-form.pdf");
  assert.equal(preview.closed, true);
  assert.equal(calls.previews.length, 0);
  assert.match(calls.alerts[0], /session has expired/);
});
