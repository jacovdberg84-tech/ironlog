import test from "node:test";
import assert from "node:assert/strict";
import { existsSync, mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Fastify from "fastify";
import sharp from "sharp";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-artisan-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_DATA_DIR = tempDir;

await import("../db/migrate.js");
const { db } = await import("../db/client.js");
const { default: workOrderRoutes } = await import("../routes/workorders.routes.js");
const { default: maintenanceRoutes } = await import("../routes/maintenance.routes.js");
const { default: reportsRoutes } = await import("../routes/reports.routes.js");
const { default: techRoutes } = await import("../routes/tech.routes.js");
const { ARTISAN_SECTIONS } = await import("../utils/artisanInspection.js");
// The reports routes start hourly schedulers; they must not keep the test process alive.
const realSetInterval = globalThis.setInterval;
globalThis.setInterval = (...args) => realSetInterval(...args).unref();
const app = Fastify({ logger: false });
await app.register(workOrderRoutes, { prefix: "/api/workorders" });
await app.register(maintenanceRoutes, { prefix: "/api/maintenance" });
await app.register(reportsRoutes, { prefix: "/api/reports" });
await app.register(techRoutes, { prefix: "/api/tech" });
await app.ready();
globalThis.setInterval = realSetInterval;
process.on("exit", () => {
  try { db.close(); } catch { /* closed */ }
  rmSync(tempDir, { recursive: true, force: true });
});

db.exec(`INSERT INTO assets (id, asset_code, asset_name, category, active) VALUES (1, 'A301AM', 'CAT 740 ADT', 'adt', 1)`);

const as = (role = "artisan", user = "jose") => ({ "x-user-name": user, "x-user-role": role, "x-user-roles": role, "x-site-code": "main" });
const call = async (method, url, payload, headers = as()) => {
  const res = await app.inject({ method, url, headers, payload });
  return { code: res.statusCode, body: res.headers["content-type"]?.includes("pdf") ? res.rawPayload : res.json() };
};
const keys = ARTISAN_SECTIONS.flatMap((s) => s.items.map((i) => i.key));
const allOk = () => Object.fromEntries(keys.map((k) => [k, "ok"]));

async function photo(query) {
  const jpg = await sharp({ create: { width: 40, height: 30, channels: 3, background: "#c00" } }).jpeg().toBuffer();
  const boundary = "----ironlogtest";
  const body = Buffer.concat([
    Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="p.jpg"\r\nContent-Type: image/jpeg\r\n\r\n`),
    jpg,
    Buffer.from(`\r\n--${boundary}--\r\n`),
  ]);
  const res = await app.inject({
    method: "POST", url: `/api/tech/inspections/photos?${new URLSearchParams(query)}`,
    headers: { ...as(), "content-type": `multipart/form-data; boundary=${boundary}` }, payload: body,
  });
  return { code: res.statusCode, body: res.json() };
}

test("technician inspection: checks, details, photos, fault work order, notice, PDF", async (t) => {
  t.after(() => app.close());

  const tpl = (await call("GET", "/api/tech/inspection-template")).body;
  assert.ok(tpl.sections.length >= 6);
  assert.ok(tpl.sections.every((s) => s.pt && s.items.every((i) => i.pt)), "every check has Portuguese");
  assert.equal((await call("GET", "/api/tech/inspection-template", null, as("operator"))).code, 403, "workshop staff only");

  // Everything has to be answered, faults need a comment, a result is required.
  const base = { asset_code: "A301AM", inspection_type: "weekly", shift: "day", machine_hours: 5120, location: "Pit 2", overall_result: "fit" };
  const partial = await call("POST", "/api/tech/inspections", { ...base, answers: { engine_oil: "ok" } });
  assert.equal(partial.code, 400);
  assert.match(partial.body.error, /Answer every check/);
  const noComment = await call("POST", "/api/tech/inspections", { ...base, answers: { ...allOk(), hydraulics: "fault" } });
  assert.match(noComment.body.error, /Say what is wrong/);
  assert.match((await call("POST", "/api/tech/inspections", { ...base, overall_result: "", answers: allOk() })).body.error, /overall result/);

  // A clean inspection: no work order.
  const clean = await call("POST", "/api/tech/inspections", { ...base, answers: { ...allOk(), fire: "na" }, client_event_id: "c1", inspection_client_id: "i1" });
  assert.equal(clean.code, 200, JSON.stringify(clean.body));
  assert.equal(clean.body.work_order_id, null);
  assert.equal(clean.body.fault_list, undefined);

  // Two faults, machine not fit: one work order, high priority, notice says do not operate.
  const faulty = {
    ...base, overall_result: "not_fit", notes: "Parked at workshop",
    answers: { ...allOk(), hydraulics: "fault", tyres_tracks: "fault" },
    comments: { hydraulics: "Boom hose leaking", tyres_tracks: "Left front cut", engine_oil: "Topped up 2 L" },
    client_event_id: "c2", inspection_client_id: "i2",
  };
  const r = await call("POST", "/api/tech/inspections", faulty);
  assert.equal(r.code, 200, JSON.stringify(r.body));
  assert.equal(r.body.faults, 2);
  const woId = r.body.work_order_id;
  const wo = db.prepare("SELECT * FROM work_orders WHERE id = ?").get(woId);
  assert.equal(wo.source, "artisan_inspection");
  assert.equal(wo.priority, "high");
  assert.match(wo.job_description, /NOT fit for work/);
  assert.match(wo.job_description, /- Hydraulic hoses, leaks, and fittings: Boom hose leaking/);
  const again = await call("POST", "/api/tech/inspections", faulty);
  assert.equal(again.body.duplicate, true, "sent twice from a phone with poor signal: saved once");
  const resent = await call("POST", "/api/tech/inspections", { ...faulty, client_event_id: "c2-retry" });
  assert.equal(resent.body.id, r.body.id, "the same form re-sent with a new event id is still saved once");
  assert.equal(resent.body.work_order_id, woId);
  assert.equal(db.prepare("SELECT COUNT(*) n FROM work_orders WHERE source = 'artisan_inspection'").get().n, 1);

  const ctx = await call("GET", "/api/maintenance/machine-prestart/context?asset_code=A301AM");
  assert.match(ctx.body.notice.message_en, /found on A301AM at a workshop inspection.*Do not operate/);
  assert.match(ctx.body.notice.message_pt, /Mangueiras hidráulicas/);

  // Photos, including one for the fault; they show on the work order too.
  const p1 = await photo({ inspection_client_id: "i2", item_key: "hydraulics", client_event_id: "p1" });
  assert.equal(p1.code, 200, JSON.stringify(p1.body));
  assert.ok(existsSync(path.join(tempDir, p1.body.url)));
  assert.equal((await photo({ inspection_client_id: "i2", item_key: "hydraulics", client_event_id: "p1" })).body.duplicate, true);
  await photo({ inspection_client_id: "i2", caption: "Overall view", client_event_id: "p2" });
  assert.equal((await photo({ inspection_client_id: "nope" })).code, 404);
  const woPhotos = db.prepare("SELECT caption FROM work_order_photos WHERE work_order_id = ? ORDER BY id").all(woId).map((x) => x.caption);
  assert.deepEqual(woPhotos, ["Inspection: Hydraulic hoses, leaks, and fittings", "Inspection: Overall view"]);

  const list = (await call("GET", "/api/tech/inspections?asset_code=A301AM")).body.rows;
  assert.equal(list.length, 2);
  assert.deepEqual([list[0].overall_result, list[0].photo_count, list[0].faults.length], ["not_fit", 2, 2]);
  assert.equal((await call("GET", "/api/tech/inspections")).body.rows.length, 2, "my inspections");

  const maint = (await call("GET", "/api/maintenance/artisan-inspections", null, as("admin", "admin"))).body.rows;
  assert.equal(maint[0].work_order_id, woId);
  assert.equal(maint[0].location, "Pit 2");

  for (const url of [`/api/tech/inspections/${r.body.id}/pdf`, `/api/reports/artisan-inspection/${r.body.id}.pdf`]) {
    const pdf = await call("GET", url, null, as("admin", "admin"));
    assert.equal(pdf.code, 200, url);
    assert.equal(pdf.body.subarray(0, 4).toString(), "%PDF");
  }
});
