import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const source = readFileSync(new URL('../../web/workorders.js', import.meta.url), 'utf8');
const functions = source.slice(source.indexOf('async function fetchWorkOrderPdf('), source.indexOf('async function downloadPlantLaborOilReport('));

function harness(status = 200) {
  const calls = { alerts: [], requests: [], downloads: [], previews: [], timers: [] };
  const preview = { closed: false, location: { replace: url => calls.previews.push(url) }, close() { this.closed = true; } };
  const context = vm.createContext({
    API: '/api',
    authHeaders: () => ({ Authorization: 'Bearer test-session' }),
    fetch: async (url, options) => {
      calls.requests.push({ url, options });
      return { ok: status === 200, status, json: async () => ({ error: 'Login required' }), blob: async () => ({ type: 'application/pdf' }) };
    },
    window: { open: () => preview },
    document: { body: { appendChild() {} }, createElement: () => ({ click() { calls.downloads.push(this.download); }, remove() {} }) },
    URL: { createObjectURL: () => 'blob:test-pdf', revokeObjectURL() {} },
    setTimeout: fn => calls.timers.push(fn),
    alert: message => calls.alerts.push(message),
  });
  vm.runInContext(functions, context);
  return { context, calls, preview };
}

test('work order PDF preview fetches with the current login token', async () => {
  const { context, calls } = harness();
  await context.openWorkOrderPdf(245);
  assert.equal(calls.requests[0].options.headers.Authorization, 'Bearer test-session');
  assert.equal(calls.requests[0].url, '/api/reports/workorder/245.pdf');
  assert.deepEqual(calls.previews, ['blob:test-pdf']);
  assert.equal(calls.timers.length, 1);
});

test('work order PDF download authenticates and saves a PDF filename', async () => {
  const { context, calls } = harness();
  await context.downloadWorkOrderPdf(245);
  assert.equal(calls.requests[0].options.headers.Authorization, 'Bearer test-session');
  assert.deepEqual(calls.downloads, ['IRONLOG_WorkOrder_245.pdf']);
});

test('expired login closes the preview without displaying an error response as a PDF', async () => {
  const { context, calls, preview } = harness(401);
  await context.openWorkOrderPdf(245);
  assert.equal(preview.closed, true);
  assert.equal(calls.previews.length, 0);
  assert.match(calls.alerts[0], /session has expired/);
});
