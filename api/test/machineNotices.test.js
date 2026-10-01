import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-notices-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { default: maintenanceRoutes } = await import("../routes/maintenance.routes.js");
const { default: workOrderRoutes } = await import("../routes/workorders.routes.js");
const app = Fastify({ logger: false });
await app.register(workOrderRoutes, { prefix: "/api/workorders" });
await app.register(maintenanceRoutes, { prefix: "/api/maintenance" });
await app.ready();
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

db.exec(`INSERT INTO assets (id, asset_code, asset_name, category, active) VALUES (1, 'A301AM', 'CAT 950 Loader', 'loader', 1)`);

const as = (role, user = "admin1") => ({ "x-user-name": user, "x-user-role": role, "x-user-roles": role, "x-site-code": "main" });
const call = async (method, url, payload, headers = {}) => {
  const res = await app.inject({ method, url, headers, payload });
  return { code: res.statusCode, body: res.json() };
};

async function prestart(date, { fault = false, ack = null, name = "Carlos" } = {}) {
  const ctx = await call("GET", `/api/maintenance/machine-prestart/context?asset_code=A301AM&check_date=${date}`);
  assert.equal(ctx.code, 200, JSON.stringify(ctx.body));
  const keys = ctx.body.template.sections.flatMap((s) => s.items.map((i) => i.key));
  const checklist = Object.fromEntries(keys.map((k, i) => [k, !(fault && i === 0)]));
  const res = await call("POST", "/api/maintenance/machine-prestart", {
    asset_code: "A301AM", check_date: date, inspector_name: name, smu_hours: 1000, checklist,
    faults: fault ? { [keys[0]]: "leaking" } : {}, ...(ack ? { notice_ack: ack } : {}),
  });
  assert.equal(res.code, 200, JSON.stringify(res.body));
  return { ctx: ctx.body, res: res.body };
}

test("fault puts up a notice, next operator ticks it, admin edits, work order done ends it", async (t) => {
  t.after(() => app.close());

  const first = await prestart("2026-09-29");
  assert.equal(first.ctx.notice, null, "no notice before any fault");

  const faulty = await prestart("2026-09-30", { fault: true });
  const woId = faulty.res.fault_work_order_id;
  assert.ok(woId, "fault opened a work order");

  const next = await prestart("2026-10-01", { name: "Joao" });
  const notice = next.ctx.notice;
  assert.ok(notice, "next operator sees the notice");
  assert.equal(notice.work_order_id, woId);
  assert.match(notice.message_en, /A301AM.*Report the machine to the workshop/);
  assert.match(notice.message_pt, /Leve a máquina à oficina/);
  assert.doesNotMatch(notice.message_pt, /\(\s*\)|\(,/, "fault names are listed");

  // The tick is saved with the check.
  await call("POST", "/api/maintenance/machine-prestart", {
    asset_code: "A301AM", check_date: "2026-10-01", inspector_name: "Joao", smu_hours: 1000,
    checklist: Object.fromEntries(next.ctx.template.sections.flatMap((s) => s.items.map((i) => [i.key, true]))),
    notice_ack: notice.id,
  });
  assert.equal((await call("GET", "/api/maintenance/machine-notices?asset_code=A301AM", null, as("artisan"))).code, 403, "supervisors only");
  const view = await call("GET", `/api/maintenance/machine-notices?work_order_id=${woId}`, null, as("workshop_admin"));
  assert.equal(view.code, 200);
  assert.equal(view.body.notice.ack_count, 1);
  assert.equal(view.body.acks[0].operator, "Joao");
  assert.equal(view.body.presets.length, 4);

  // Admin chooses a preset: Portuguese comes with it.
  const put = await call("PUT", "/api/maintenance/machine-notices", { work_order_id: woId, preset: "do_not_operate" }, as("workshop_admin"));
  assert.equal(put.code, 200, JSON.stringify(put.body));
  assert.equal(put.body.notice.source, "admin");
  assert.match(put.body.notice.message_pt, /Não opere/);
  assert.equal(put.body.notice.ack_count, 0, "a new message needs a new tick");
  assert.equal((await call("PUT", "/api/maintenance/machine-notices", { work_order_id: woId, message_en: " " }, as("workshop_admin"))).code, 400);

  // Another fault does not overwrite what the admin wrote.
  await prestart("2026-10-02", { fault: true });
  const ctx2 = await call("GET", "/api/maintenance/machine-prestart/context?asset_code=A301AM&check_date=2026-10-03");
  assert.match(ctx2.body.notice.message_en, /Do not operate/);

  // Clearing, then an auto notice tied to a work order ends when it is completed.
  assert.equal((await call("DELETE", "/api/maintenance/machine-notices?asset_code=A301AM", null, as("plant_manager"))).code, 200);
  assert.equal((await call("GET", "/api/maintenance/machine-prestart/context?asset_code=A301AM&check_date=2026-10-03")).body.notice, null);
  const f2 = await prestart("2026-10-03", { fault: true });
  assert.ok((await call("GET", "/api/maintenance/machine-prestart/context?asset_code=A301AM&check_date=2026-10-04")).body.notice);
  db.prepare("UPDATE work_orders SET status = 'completed' WHERE id = ?").run(f2.res.fault_work_order_id);
  assert.equal((await call("GET", "/api/maintenance/machine-prestart/context?asset_code=A301AM&check_date=2026-10-04")).body.notice, null, "repair done, notice gone");
});

test("Portuguese notice keeps the English name of a check it has no translation for", async () => {
  const { autoNoticeText } = await import("../utils/machineNotices.js");
  const t = autoNoticeText("LDV01", [{ label: "Some new check" }]);
  assert.match(t.pt, /\(Some new check\)/);
});
