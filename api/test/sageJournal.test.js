import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import Database from "better-sqlite3";
import Fastify from "fastify";
import ExcelJS from "exceljs";
import * as sage from "../utils/sageJournal.js";

// Connecting to other systems, part B: journal batches for Sage.
const pair = (category, debitAcc, creditAcc, amount, cc, asset) => [
  { category, account_code: debitAcc, cost_center_code: cc, asset_code: asset, debit: amount, credit: 0 },
  { category, account_code: creditAcc, cost_center_code: cc, asset_code: asset, debit: 0, credit: amount },
];
const LINES = [
  ...pair("parts", "5200-PARTS-EXPENSE", "1400-INVENTORY", 100.004, "PIT", "T01"),
  ...pair("parts", "5200-PARTS-EXPENSE", "1400-INVENTORY", 50.004, "PIT", "T02"),
  ...pair("parts", "5200-PARTS-EXPENSE", "1400-INVENTORY", 20, null, "UNLINKED"),
  ...pair("downtime", "5400-DOWNTIME-EXPENSE", "2600-DOWNTIME-CLEARING", 900, "PIT", "T01"),
  ...pair("procurement_grn", "1400-INVENTORY", "2100-GRNI", 300, null, null),
  ...pair("procurement_ap", "2100-GRNI", "2200-AP", 300, null, null),
];
const RUN = { id: 1, run_number: "JR-2026-09-001", start_date: "2026-09-01", end_date: "2026-09-30", status: "draft" };
const ACCOUNTS = { "5200-PARTS-EXPENSE": "5200", "1400-INVENTORY": "1400", "2100-GRNI": "2100", "2200-AP": "2200" };

test("a run becomes a balanced Sage batch: mapped, netted, downtime left out, rounding on the largest line", () => {
  const missing = sage.buildSageJournal(RUN, LINES, { accounts: { "5200-PARTS-EXPENSE": "5200" } });
  assert.equal(missing.ok, false);
  assert.deepEqual(missing.missing_accounts.sort(), ["1400-INVENTORY", "2100-GRNI", "2200-AP"], "downtime is off by default, so its accounts are not asked for");
  assert.match(missing.errors[0], /Set the Sage account for: Parts stock \(1400-INVENTORY\)/);

  const j = sage.buildSageJournal(RUN, LINES, { accounts: ACCOUNTS, default_cost_centre: "WS" });
  assert.equal(j.ok, true, j.errors.join());
  assert.equal(j.reference, "ILJR-2026-09-001");
  assert.equal(j.date, "2026-09-30", "dated at the run's end");
  assert.deepEqual(j.entries.map((e) => [e.account, e.debit, e.credit]), [
    ["1400-PIT", 0, 150.01],
    ["1400-WS", 280, 0],
    ["2200-WS", 0, 300],
    ["5200-PIT", 150.01, 0],
    ["5200-WS", 20, 0],
  ].sort((a, b) => (b[1] > 0) - (a[1] > 0) || a[0].localeCompare(b[0])));
  assert.equal(j.entries.some((e) => e.account.startsWith("2100")), false, "GRNI received and invoiced nets to nothing");
  assert.equal(j.total_debit, j.total_credit);
  assert.equal(j.total_debit, 450.01);
  assert.equal(j.rounding, 0);

  const byMachine = sage.buildSageJournal(RUN, LINES, { accounts: ACCOUNTS, cost_centre: "column", detail: "machine", categories: { procurement_grn: false, procurement_ap: false } });
  assert.deepEqual(byMachine.entries.filter((e) => e.account === "5200").map((e) => [e.cost_centre, e.machine, e.debit]),
    [["", "UNLINKED", 20], ["PIT", "T01", 100], ["PIT", "T02", 50]]);
  assert.equal(byMachine.total_debit, byMachine.total_credit, "rounded per machine and still balanced");
  assert.match(byMachine.entries[0].description, /Parts used on machines (T01|T02|UNLINKED) 2026-09-01 to 2026-09-30/);

  const third = (cc) => ({ category: "parts", account_code: "5200-PARTS-EXPENSE", cost_center_code: cc, debit: 0.006, credit: 0 });
  const cents = sage.buildSageJournal(RUN, [third("A"), third("B"), third("C"), { category: "parts", account_code: "1400-INVENTORY", cost_center_code: "D", debit: 0, credit: 0.018 }],
    { accounts: ACCOUNTS, cost_centre: "column", categories: { parts: true } });
  assert.equal(cents.ok, true);
  assert.deepEqual([cents.total_debit, cents.total_credit, cents.rounding], [0.03, 0.03, -0.01], "a cent of rounding goes on the largest line, not left out of balance");

  const off = sage.buildSageJournal(RUN, [{ category: "parts", account_code: "5200-PARTS-EXPENSE", debit: 10, credit: 0 }], { accounts: ACCOUNTS });
  assert.match(off.errors[0], /does not balance \(10\.00 USD out\)/);
  assert.match(sage.buildSageJournal({ ...RUN, status: "reversed" }, LINES, { accounts: ACCOUNTS }).errors[0], /reversed/);
});

test("the CSV and the Sage 300 workbook carry the batch", async () => {
  const j = sage.buildSageJournal(RUN, LINES, { accounts: ACCOUNTS, cost_centre: "column", date_format: "DD/MM/YYYY", reference_prefix: "" });
  const csv = sage.sageCsv(j).trim().split("\r\n");
  assert.equal(csv[0], "Date,Reference,Account,Cost centre,Description,Debit,Credit");
  assert.match(csv[1], /^30\/09\/2026,JR-2026-09-001,/);
  assert.equal(sage.formatSageDate("2026-09-30", "YYYYMMDD"), "20260930");

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await sage.sage300Workbook(ExcelJS, sage.buildSageJournal(RUN, LINES, { accounts: ACCOUNTS })).xlsx.writeBuffer());
  const head = wb.getWorksheet("Journal_Headers");
  assert.deepEqual(head.getRow(1).values.slice(1), ["BATCHID", "BTCHENTRY", "SRCELEDGER", "SRCETYPE", "DATEENTRY", "JRNLDESC"]);
  assert.equal(head.getRow(2).getCell(6).value, "IronLog JR-2026-09-001 2026-09-01 to 2026-09-30");
  const det = wb.getWorksheet("Journal_Details");
  const rows = det.getSheetValues().slice(2).map((r) => r.slice(1));
  assert.deepEqual(rows.map((r) => r[2]), rows.map((_, i) => String((i + 1) * 20).padStart(10, "0")));
  assert.equal(Math.round(rows.reduce((t, r) => t + r[4] * 100, 0)), 0, "debits positive, credits negative, summing to zero");
});

test("settings are checked before they are saved", () => {
  const d = new Database(":memory:");
  assert.equal(sage.readSageSettings(d).categories.downtime, false);
  const bad = sage.saveSageSettings(d, { layout: "quickbooks", accounts: { "9999-X": "1", "5200-PARTS-EXPENSE": "52;00" }, reference_prefix: "much too long prefix", categories: Object.fromEntries(Object.keys(sage.SAGE_CATEGORIES).map((k) => [k, false])) });
  assert.equal(bad.ok, false);
  assert.equal(bad.errors.length, 5);
  const ok = sage.saveSageSettings(d, { layout: "csv", accounts: { "5200-PARTS-EXPENSE": "5200/000" }, categories: { downtime: true } }, "fin");
  assert.equal(ok.ok, true);
  assert.deepEqual([ok.settings.layout, ok.settings.accounts["5200-PARTS-EXPENSE"], ok.settings.accounts["1400-INVENTORY"], ok.settings.categories.downtime, ok.settings.categories.parts, ok.settings.updated_by],
    ["csv", "5200/000", "", true, true, "fin"]);
  assert.equal(sage.saveSageSettings(d, { cost_centre: "none" }).settings.accounts["5200-PARTS-EXPENSE"], "5200/000", "a partial save keeps the rest");
});

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-sage-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;
process.env.IRONLOG_AUTH_REQUIRED = "true";

test("finance maps the accounts and downloads the batch; others cannot change the settings", async (t) => {
  await import("../db/migrate.js");
  (await import("../db/hybridPhase1.js")).runHybridPhase1Migration();
  const { db } = await import("../db/client.js");
  const { ironlogAuthHook } = await import("../auth/hook.js");
  const realSetInterval = globalThis.setInterval;
  globalThis.setInterval = (...args) => realSetInterval(...args).unref();
  const app = Fastify({ logger: false });
  app.addHook("onRequest", ironlogAuthHook);
  app.decorateRequest("ironlogAllowedSites", null);
  await app.register((await import("../routes/auth.routes.js")).default, { prefix: "/api/auth" });
  // The stores routes make the stock tables the procurement routes build on.
  await app.register((await import("../routes/stock.routes.js")).default, { prefix: "/api/stock" });
  await app.register((await import("../routes/procurement.routes.js")).default, { prefix: "/api/procurement" });
  await app.ready();
  globalThis.setInterval = realSetInterval;
  t.after(async () => { await app.close(); try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

  const session = (username, role) => {
    db.prepare(`INSERT INTO users (username, role, roles_json, active) VALUES (?, ?, ?, 1)`).run(username, role, JSON.stringify([role]));
    const token = crypto.randomBytes(16).toString("hex");
    db.prepare(`INSERT INTO auth_sessions (token, user_id, expires_at) VALUES (?, (SELECT id FROM users WHERE username = ?), datetime('now', '+1 day'))`).run(token, username);
    return { authorization: `Bearer ${token}` };
  };
  const fin = session("fin", "finance");
  const tech = session("tech", "artisan");

  const runId = db.prepare(`INSERT INTO finance_posting_runs (run_number, period, start_date, end_date, run_type, status, currency) VALUES ('JR-2026-09-001', '2026-09', '2026-09-01', '2026-09-30', 'summary', 'draft', 'USD')`).run().lastInsertRowid;
  const ins = db.prepare(`INSERT INTO finance_posting_run_lines (run_id, category, tx_date, source_module, source_type, account_code, cost_center_code, asset_code, description, debit, credit, currency) VALUES (?, ?, '2026-09-10', 'maintenance', 'x', ?, ?, ?, 'x', ?, ?, 'USD')`);
  for (const l of LINES) ins.run(runId, l.category, l.account_code, l.cost_center_code, l.asset_code, l.debit, l.credit);

  const preview = await app.inject({ method: "GET", url: `/api/procurement/journals/runs/${runId}/sage-preview`, headers: fin });
  assert.equal(preview.statusCode, 200, preview.body);
  assert.equal(preview.json().ok, false);
  assert.equal((await app.inject({ method: "GET", url: `/api/procurement/journals/runs/${runId}/export-sage`, headers: fin })).statusCode, 409, "not until the accounts are mapped");

  assert.equal((await app.inject({ method: "PUT", url: "/api/procurement/journals/sage-settings", headers: tech, payload: { accounts: ACCOUNTS } })).statusCode, 403);
  const bad = await app.inject({ method: "PUT", url: "/api/procurement/journals/sage-settings", headers: fin, payload: { layout: "nope" } });
  assert.equal(bad.statusCode, 400);
  const saved = await app.inject({ method: "PUT", url: "/api/procurement/journals/sage-settings", headers: fin, payload: { accounts: ACCOUNTS, layout: "csv", cost_centre: "column" } });
  assert.equal(saved.statusCode, 200, saved.body);

  assert.equal((await app.inject({ method: "GET", url: `/api/procurement/journals/runs/${runId}/export-sage`, headers: tech })).statusCode, 403);
  const csv = await app.inject({ method: "GET", url: `/api/procurement/journals/runs/${runId}/export-sage`, headers: fin });
  assert.equal(csv.statusCode, 200, csv.body);
  assert.match(csv.headers["content-disposition"], /Sage_JR-2026-09-001\.csv/);
  assert.match(csv.body, /^Date,Reference,Account,Cost centre,Description,Debit,Credit\r\n2026-09-30,ILJR-2026-09-001,/);

  await app.inject({ method: "PUT", url: "/api/procurement/journals/sage-settings", headers: fin, payload: { layout: "sage300" } });
  const xlsx = await app.inject({ method: "GET", url: `/api/procurement/journals/runs/${runId}/export-sage`, headers: fin });
  assert.equal(xlsx.statusCode, 200);
  assert.match(xlsx.headers["content-type"], /spreadsheetml/);
  const audit = db.prepare(`SELECT action FROM audit_logs WHERE module = 'finance' ORDER BY id`).all().map((r) => r.action);
  assert.deepEqual(audit, ["journals.sage_settings", "journals.export_sage", "journals.sage_settings", "journals.export_sage"]);
});
