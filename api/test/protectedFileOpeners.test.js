import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const root = new URL('../..', import.meta.url);
const readWeb = (name) => readFileSync(new URL(`web/${name}`, root), 'utf8');

test('operational slip PDF is fetched with the active admin session', async () => {
  const source = readWeb('breakdown-ops.js');
  const start = source.indexOf('async function openProtectedPdf(');
  const end = source.indexOf('async function ensureOpenBreakdownOps(');
  assert.ok(start >= 0 && end > start, 'operational slip PDF helpers are present');

  const calls = { requests: [], previews: [], alerts: [] };
  const preview = {
    closed: false,
    opener: null,
    location: { replace: (url) => calls.previews.push(url) },
    close() { this.closed = true; },
  };
  const context = vm.createContext({
    API: '/api',
    authHeaders: () => ({ Authorization: 'Bearer admin-session' }),
    fetch: async (url, options) => {
      calls.requests.push({ url, options });
      return { ok: true, status: 200, blob: async () => ({ type: 'application/pdf' }) };
    },
    window: { open: () => preview },
    URL: { createObjectURL: () => 'blob:operational-slip', revokeObjectURL() {} },
    setTimeout() {},
    alert: (message) => calls.alerts.push(message),
  });
  vm.runInContext(source.slice(start, end), context);

  await context.openBoSlipPdf(73);
  assert.equal(calls.requests[0].url, '/api/breakdown-ops/slips/73/pdf');
  assert.equal(calls.requests[0].options.headers.Authorization, 'Bearer admin-session');
  assert.deepEqual(calls.previews, ['blob:operational-slip']);
  assert.deepEqual(calls.alerts, []);
});

test('admin report buttons no longer navigate directly to protected API files', () => {
  const app = readWeb('app.js');
  const maintenance = readWeb('maintenance.js');
  const breakdownOps = readWeb('breakdown-ops.js');

  assert.match(app, /function openAuthedReport\(/);
  assert.match(app, /openAuthedReport\(`\$\{API\}\/api\/breakdown-ops\/slips\/\$\{n\}\/pdf`/);
  assert.match(maintenance, /function openProtectedPdf\(/);
  assert.match(breakdownOps, /async function openProtectedPdf\(/);

  for (const source of [app, maintenance, breakdownOps]) {
    assert.doesNotMatch(
      source,
      /window\.open\(`\$\{API\}\/[^`]*(?:\.pdf|\.xlsx|\.csv|\/download)[^`]*`/,
      'protected file routes must be fetched with auth headers before opening',
    );
  }
});
