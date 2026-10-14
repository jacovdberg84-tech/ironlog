import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import {
  faultMessage,
  notesWithFaults,
  prestartFaultList,
  syncPrestartFaultWorkOrder,
  unansweredChecks,
} from "../utils/prestartFaults.js";

const checklist = [
  { key: "brakes_ok", label: "Brakes", ok: false },
  { key: "lights_ok", label: "Lights", ok: true },
  { key: "tyres_ok", label: "Tyres", ok: false },
];

function db() {
  const d = new Database(":memory:");
  d.exec(`CREATE TABLE work_orders (id INTEGER PRIMARY KEY, asset_id INTEGER, source TEXT, reference_id INTEGER,
    status TEXT, site_code TEXT, job_description TEXT, closed_at TEXT, completion_notes TEXT)`);
  return d;
}

test("unanswered checks are the ones without a true/false answer", () => {
  assert.deepEqual(unansweredChecks(checklist, { brakes_ok: false, lights_ok: true }), ["Tyres"]);
  assert.deepEqual(unansweredChecks(checklist, { brakes_ok: false, lights_ok: true, tyres_ok: "yes" }), ["Tyres"]);
  assert.deepEqual(unansweredChecks(checklist, null), ["Brakes", "Lights", "Tyres"]);
});

test("faults carry the operator's comment", () => {
  assert.deepEqual(prestartFaultList(checklist, { brakes_ok: " soft pedal " }), [
    { key: "brakes_ok", label: "Brakes", comment: "soft pedal" },
    { key: "tyres_ok", label: "Tyres", comment: "" },
  ]);
});

test("notes get one Faults line, rebuilt each time", () => {
  const faults = prestartFaultList(checklist, { brakes_ok: "soft pedal" });
  assert.equal(notesWithFaults("Cab dusty", faults), "Cab dusty | Faults: Brakes (soft pedal); Tyres");
  assert.equal(notesWithFaults("Cab dusty | Faults: Old | KM flagged for supervisor review", faults), "Cab dusty | Faults: Brakes (soft pedal); Tyres");
  assert.equal(notesWithFaults("", []), null);
});

test("faults open one work order per check; resubmits update it", () => {
  const d = db();
  const faults = prestartFaultList(checklist, { brakes_ok: "soft pedal" });
  const first = syncPrestartFaultWorkOrder(d, { assetId: 4, checkId: 9, checkDate: "2026-09-29", operator: "Sipho", faults });
  assert.deepEqual(first, { work_order_id: 1, created: true, closed: false });
  const wo = d.prepare("SELECT * FROM work_orders WHERE id = 1").get();
  assert.equal(wo.source, "prestart");
  assert.equal(wo.status, "open");
  assert.equal(wo.job_description, "Pre-start faults reported by Sipho on 2026-09-29:\n- Brakes: soft pedal\n- Tyres");
  const again = syncPrestartFaultWorkOrder(d, { assetId: 4, checkId: 9, checkDate: "2026-09-29", operator: "Sipho", faults: faults.slice(0, 1) });
  assert.deepEqual(again, { work_order_id: 1, created: false, closed: false });
  assert.equal(d.prepare("SELECT COUNT(*) n FROM work_orders").get().n, 1);
  assert.match(faultMessage(faults, first), /work order #1/);
});

test("resubmitting with no faults closes an untouched work order only", () => {
  const d = db();
  const faults = prestartFaultList(checklist, {});
  syncPrestartFaultWorkOrder(d, { assetId: 4, checkId: 9, checkDate: "2026-09-29", operator: "Sipho", faults });
  assert.deepEqual(syncPrestartFaultWorkOrder(d, { assetId: 4, checkId: 9, checkDate: "2026-09-29", faults: [] }), { work_order_id: 1, created: false, closed: true });
  assert.equal(d.prepare("SELECT status FROM work_orders WHERE id = 1").get().status, "closed");

  syncPrestartFaultWorkOrder(d, { assetId: 4, checkId: 10, checkDate: "2026-09-30", faults });
  d.prepare("UPDATE work_orders SET status = 'in_progress' WHERE id = 2").run();
  assert.equal(syncPrestartFaultWorkOrder(d, { assetId: 4, checkId: 10, checkDate: "2026-09-30", faults: [] }), null);
  assert.equal(d.prepare("SELECT status FROM work_orders WHERE id = 2").get().status, "in_progress");
  assert.equal(syncPrestartFaultWorkOrder(d, { assetId: 4, checkId: 11, checkDate: "2026-10-01", faults: [] }), null);
});

test("a second day's faults join the machine's open pre-start job instead of a new one", () => {
  const d = db();
  const brakes = prestartFaultList(checklist, { brakes_ok: "soft pedal" });
  const day1 = syncPrestartFaultWorkOrder(d, { assetId: 4, checkId: 20, checkDate: "2026-10-01", operator: "Sipho", faults: brakes });
  const day2 = syncPrestartFaultWorkOrder(d, { assetId: 4, checkId: 21, checkDate: "2026-10-02", operator: "Joao", faults: brakes });
  assert.deepEqual(day2, { work_order_id: day1.work_order_id, created: false, closed: false, added: true });
  assert.equal(d.prepare("SELECT COUNT(*) n FROM work_orders").get().n, 1);
  const desc = d.prepare("SELECT job_description FROM work_orders WHERE id = ?").get(day1.work_order_id).job_description;
  assert.match(desc, /^Pre-start faults reported by Sipho on 2026-10-01:/);
  assert.match(desc, /Also reported by Joao on 2026-10-02:\n- Brakes: soft pedal \(again\)\n- Tyres \(again\)/);
  assert.match(faultMessage(brakes, day2), /work order #1 \(added to the open pre-start job\)/);

  // Day 2's check resubmitted clean: the shared job stays open, with a note.
  assert.equal(syncPrestartFaultWorkOrder(d, { assetId: 4, checkId: 21, checkDate: "2026-10-02", operator: "Joao", faults: [] }), null);
  const wo = d.prepare("SELECT status, job_description FROM work_orders WHERE id = 1").get();
  assert.equal(wo.status, "open");
  assert.match(wo.job_description, /Re-checked by Joao on 2026-10-02: no faults on this pre-start\./);

  // Another machine still gets its own job; a finished job means the next fault opens a new one.
  assert.equal(syncPrestartFaultWorkOrder(d, { assetId: 5, checkId: 22, checkDate: "2026-10-02", faults: brakes }).created, true);
  d.prepare("UPDATE work_orders SET status = 'completed' WHERE id = 1").run();
  assert.equal(syncPrestartFaultWorkOrder(d, { assetId: 4, checkId: 23, checkDate: "2026-10-03", faults: brakes }).created, true);
});
