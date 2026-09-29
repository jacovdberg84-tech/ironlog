import test from "node:test";
import assert from "node:assert/strict";
import {
  breakdownNextStep,
  downtimeCell,
  draftWeeklyFindings,
  isWorkOrderDone,
  workOrderBlockerCell,
  workOrderOwnerCell,
  workOrderProgressCell,
} from "../utils/weeklyForumFindings.js";

test("signed-off (approved) work orders are finished work", () => {
  for (const s of ["completed", "approved", "closed", "Done"]) assert.equal(isWorkOrderDone(s), true, s);
  for (const s of ["open", "assigned", "in_progress", "In Progress"]) assert.equal(isWorkOrderDone(s), false, s);
  assert.equal(workOrderProgressCell({ status: "approved" }), "Signed off");
  assert.equal(workOrderProgressCell({ status: "in_progress", repair_progress: "Pump fitted, testing" }), "In progress — Pump fitted, testing");
});

test("owner and due date come from the work order", () => {
  assert.equal(workOrderOwnerCell({ status: "assigned", assigned_artisan_name: "Sipho", due_date: "2026-10-02" }), "Sipho • due 2026-10-02");
  assert.equal(workOrderOwnerCell({ status: "open" }), "No artisan assigned • no due date");
  assert.equal(workOrderOwnerCell({ status: "open", ets_repair_date: "2026-09-30" }), "No artisan assigned • due 2026-09-30");
  assert.equal(workOrderOwnerCell({ status: "approved", completed_at: "2026-09-24 10:00", assigned_artisan_name: "Sipho" }), "Done 2026-09-24 • Sipho");
});

test("parts blockers come from open parts requests, then the breakdown", () => {
  assert.equal(workOrderBlockerCell({ status: "in_progress", open_parts_requests: 2, first_waiting_part: "Hydraulic pump" }), "Waiting: Hydraulic pump +1 more");
  assert.equal(workOrderBlockerCell({ status: "open", parts_status: "Ordered" }), "Parts: Ordered");
  assert.equal(workOrderBlockerCell({ status: "open" }), "No parts outstanding");
  assert.equal(workOrderBlockerCell({ status: "closed", open_parts_requests: 1 }), "—");
});

test("downtime explains zero hours logged over several days", () => {
  assert.equal(downtimeCell({ downtime_hours: 0, day_count: 7 }), "0 h logged • down 7 days");
  assert.equal(downtimeCell({ downtime_hours: 11, day_count: 1 }), "11 h");
  assert.equal(downtimeCell({ downtime_hours: 5.5, day_count: 2 }), "5.5 h • 2 days");
  assert.equal(breakdownNextStep({ status: "CLOSED", end_at: "2026-09-25 14:00:00" }), "Returned 2026-09-25");
  assert.equal(breakdownNextStep({ status: "OPEN", parts_status: "In transit", ets_repair_date: "2026-09-25" }), "Parts: In transit • Return 2026-09-25");
  assert.equal(breakdownNextStep({ status: "OPEN" }), "Set return target");
});

test("drafted findings summarise the week from the data", () => {
  const f = draftWeeklyFindings({
    selectedKpis: { downtime: 36.5 },
    previousKpis: { downtime: 11.6 },
    breakdownRows: [
      { asset_code: "A301AM", downtime_hours: 11, component: "Steering coupling", status: "OPEN" },
      { asset_code: "G01AM", downtime_hours: 5.5, status: "CLOSED" },
    ],
    workOrderRows: [
      { status: "approved" }, { status: "closed" },
      { status: "open" }, { status: "in_progress", assigned_artisan_name: "Sipho", parts_status: "Ordered" },
    ],
    selectedCosts: { total: 2507, partsIssued: 1227, internalLabor: 1280 },
    previousCosts: { total: 1000 },
  });
  assert.equal(f.Downtime.finding, "36.5 h mechanical downtime, up from 11.6 h last week. Largest: A301AM 11 h (Steering coupling).");
  assert.equal(f.Downtime.action, "Return A301AM to service");
  assert.equal(f.Repairs.finding, "2 jobs finished, 2 still open, 1 waiting on parts.");
  assert.equal(f.Repairs.action, "Assign artisans to 1 open job");
  assert.equal(f.Costs.finding, "$2,507 this week (parts $1,227, labour $1,280) vs $1,000 last week (+151%).");
  assert.equal(f.Costs.action, "Review the cost increase");
});
