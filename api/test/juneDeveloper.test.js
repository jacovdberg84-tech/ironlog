import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-june-dev-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
delete process.env.IRONLOG_GITHUB_TOKEN;
const { db } = await import("../db/client.js");
const jd = await import("../utils/juneDeveloper.js");
const { ensureDeveloperTables } = await import("../utils/developerClaude.js");
const { buildDailyBriefing, briefingSite } = await import("../utils/juneDailyBriefing.js");
process.on("exit", () => { try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

const ctx = { siteCode: "main", user: "jaco" };
ensureDeveloperTables(db);
const add = db.prepare(`
  INSERT INTO developer_requests (site_code, title, request_text, priority, model, created_by, issue_number, status, status_json)
  VALUES (?, ?, 'x', 'medium', ?, 'jaco', ?, ?, ?)
`);
add.run("main", "Tidy the Developer setup box", "sonnet", 61, "deployed", JSON.stringify({ claude_note: "Done." }));
add.run("main", "Fuel cost column on the Daily Log", "opus", 66, "in_review", JSON.stringify({
  claude_note: "**Claude finished**\n### What changed\n- [x] Added the column\n- [ ] Check the PDF\nSee [the PR](https://github.com/x/pull/7) and `daily-input.js`.",
  changes: { files: [{ file: "web/app/daily-input.js" }, { file: "api/test/a.test.js" }] },
  pull_request: { number: 7, merged_at: null },
}));
add.run("main", "Bigger font on the stores terminal", "sonnet", 67, "working", JSON.stringify({ claude_note: "Claude Code is working…" }));
add.run("other", "Another site's request", "opus", 70, "in_review", null);

test("June lists what Claude is working on, for this site only", async () => {
  const s = await jd.developerStatusForJune(ctx, { dbConn: db });
  assert.equal(s.github_connected, false);
  assert.deepEqual(s.counts, { waiting_on_you: 1, in_progress: 1, live: 1 });
  assert.equal(s.requests.length, 3, "the other site's request is not included");
  const fuel = s.requests.find((r) => r.issue_number === 66);
  assert.equal(fuel.status_text, "Pull request ready for review");
  assert.equal(fuel.needs_you, true);
  assert.equal(fuel.files_changed, 2);
  assert.equal(fuel.model, "Claude Opus");
  // Spoken note: no markdown symbols, links reduced to their words.
  assert.equal(fuel.claude_note, "Claude finished What changed Done: Added the column To do: Check the PDF See the PR and daily-input.js.");
});

test("June finds a request by issue number or by words from its title", () => {
  assert.equal(jd.findDeveloperRequest(ctx, "66", { dbConn: db }).issue_number, 66);
  assert.equal(jd.findDeveloperRequest(ctx, "#67", { dbConn: db }).issue_number, 67);
  assert.equal(jd.findDeveloperRequest(ctx, "the fuel cost one", { dbConn: db }).issue_number, 66);
  assert.equal(jd.findDeveloperRequest(ctx, "stores terminal font", { dbConn: db }).issue_number, 67);
  assert.equal(jd.findDeveloperRequest(ctx, "", { dbConn: db }).issue_number, 67, "empty means the latest unfinished one");
  assert.equal(jd.findDeveloperRequest(ctx, "70", { dbConn: db }), null, "other sites are out of reach");
  assert.equal(jd.findDeveloperRequest(ctx, "zebra crossing", { dbConn: db }), null);
});

test("a new request is only prepared, with sensible defaults, never sent", () => {
  const r = jd.prepareDeveloperRequest({ title: "  Add Portuguese labels to the stores terminal  ", request_text: "Translate the Issue and Receive buttons into Portuguese for the storemen.", model: "sonnet" });
  assert.equal(r.ready, true);
  assert.equal(r.kind, "developer_request");
  assert.deepEqual(r.draft, { title: "Add Portuguese labels to the stores terminal", request_text: "Translate the Issue and Receive buttons into Portuguese for the storemen.", priority: "medium", model: "sonnet" });
  assert.match(r.next_step, /Nothing has been sent/);
  assert.equal(jd.prepareDeveloperRequest({ title: "x", request_text: "too short" }).ready, false);
  assert.equal(jd.prepareDeveloperRequest({ title: "Fix", request_text: "Make the dashboard load faster please.", model: "gpt" }).draft.model, "opus", "unknown models fall back to Opus");
});

test("a reply is prepared for the right request, and refused for finished ones", () => {
  const r = jd.prepareDeveloperReply(ctx, { request: "fuel cost", message: "Also show the column on the daily PDF." }, { dbConn: db });
  assert.equal(r.ready, true);
  assert.deepEqual(r.draft, { request_id: 2, title: "Fuel cost column on the Daily Log", issue_number: 66, message: "Also show the column on the daily PDF." });
  const done = jd.prepareDeveloperReply(ctx, { request: "61", message: "One more thing." }, { dbConn: db });
  assert.equal(done.ready, false);
  assert.match(done.error, /already finished/);
  assert.equal(jd.prepareDeveloperReply(ctx, { request: "66", message: "x" }, { dbConn: db }).ready, false);
  assert.match(jd.prepareDeveloperReply(ctx, { request: "nothing like it", message: "Hello Claude" }, { dbConn: db }).error, /could not find/);
});

test("June can open a request on screen", () => {
  const r = jd.openDeveloperRequest(ctx, { request: "stores" }, { dbConn: db });
  assert.equal(r.found, true);
  assert.equal(r.request.request_id, 3);
});

test("the daily briefing mentions changes waiting for review", async () => {
  const dev = jd.developerBriefing(ctx, { dbConn: db });
  assert.deepEqual(dev, {
    waiting_on_you: [{ title: "Fuel cost column on the Daily Log", status: "Pull request ready for review" }],
    claude_working_on: ["Bigger font on the stores terminal"],
  });
  const b = await buildDailyBriefing({
    site: briefingSite({}),
    fetchImpl: async () => ({ ok: false, status: 500 }),
    ironlogBrief: () => ({ metrics: {}, open_breakdowns: [], pm_due: [] }),
    calendarToday: () => ({}),
    developer: () => dev,
  });
  assert.equal(b.developer.waiting_on_you.length, 1);
  assert.ok(b.order.some((line) => /Developer Centre/.test(line)));
  // Nothing to report: no developer section at all.
  const quiet = await buildDailyBriefing({ site: briefingSite({}), fetchImpl: async () => ({ ok: false, status: 500 }), ironlogBrief: () => ({}), calendarToday: () => ({}), developer: () => null });
  assert.equal(quiet.developer, undefined);
  assert.ok(!quiet.order.some((line) => /Developer Centre/.test(line)));
});
