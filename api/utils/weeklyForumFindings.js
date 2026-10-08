// IRONLOG/api/utils/weeklyForumFindings.js — wording for the weekly forum deck.
// Work orders move open → assigned → in_progress → completed → approved (signed
// off) → closed. Completed, approved and closed jobs are finished work.

const DONE = new Set(["completed", "approved", "closed", "done"]);
const STATUS_LABELS = {
  open: "Not started",
  assigned: "Assigned",
  in_progress: "In progress",
  completed: "Completed",
  approved: "Signed off",
  closed: "Closed",
  done: "Done",
};

const key = (status) => String(status || "open").trim().toLowerCase().replace(/\s+/g, "_");

export const DONE_WORK_ORDER_STATUSES = [...DONE];

export function isWorkOrderDone(status) {
  return DONE.has(key(status));
}

export function workOrderStatusLabel(status) {
  return STATUS_LABELS[key(status)] || String(status || "Open").replace(/_/g, " ");
}

const fmtHours = (h) => {
  const n = Number(h) || 0;
  return `${Number.isInteger(n) ? n : Number(n.toFixed(1))} h`;
};

/**
 * Hours-down cell. Downtime is logged against scheduled hours, so a machine that
 * was down on days it was not scheduled logs 0 h; say so instead of a bare "0 h".
 */
export function downtimeCell({ downtime_hours, day_count }) {
  const days = Number(day_count) || 0;
  const hours = Number(downtime_hours) || 0;
  const dayText = `${days} day${days === 1 ? "" : "s"}`;
  if (hours <= 0) return days ? `0 h logged • down ${dayText}` : "0 h logged";
  return days > 1 ? `${fmtHours(hours)} • ${dayText}` : fmtHours(hours);
}

/** Return / next-step cell for a breakdown. */
export function breakdownNextStep(b) {
  const closed = String(b.status || "").trim().toUpperCase() === "CLOSED";
  if (closed) return b.end_at ? `Returned ${String(b.end_at).slice(0, 10)}` : "Returned to service";
  const parts = [b.parts_status ? `Parts: ${b.parts_status}` : "", b.ets_repair_date ? `Return ${String(b.ets_repair_date).slice(0, 10)}` : ""].filter(Boolean);
  return parts.length ? parts.join(" • ") : "Set return target";
}

/** Owner / due cell for a work order. */
export function workOrderOwnerCell(w) {
  if (isWorkOrderDone(w.status)) {
    const doneAt = w.completed_at || w.closed_at;
    return doneAt ? `Done ${String(doneAt).slice(0, 10)}${w.assigned_artisan_name ? ` • ${w.assigned_artisan_name}` : ""}` : "Done";
  }
  const owner = String(w.assigned_artisan_name || "").trim() || "No artisan assigned";
  const due = w.due_date || w.ets_repair_date;
  return due ? `${owner} • due ${String(due).slice(0, 10)}` : `${owner} • no due date`;
}

/** Progress cell: status plus the latest repair progress note. */
export function workOrderProgressCell(w) {
  const label = workOrderStatusLabel(w.status);
  const note = String(w.repair_progress || "").trim();
  return note && !isWorkOrderDone(w.status) ? `${label} — ${note}` : label;
}

/** Parts blocker cell. */
export function workOrderBlockerCell(w) {
  if (isWorkOrderDone(w.status)) return "—";
  if (Number(w.open_parts_requests) > 0) {
    const more = Number(w.open_parts_requests) > 1 ? ` +${Number(w.open_parts_requests) - 1} more` : "";
    return `Waiting: ${w.first_waiting_part || "parts"}${more}`;
  }
  if (w.parts_status) return `Parts: ${w.parts_status}`;
  return "No parts outstanding";
}

const pctChange = (current, previous) => {
  const c = Number(current) || 0;
  const p = Number(previous) || 0;
  if (!p) return null;
  return ((c - p) / p) * 100;
};

/**
 * Drafted "What changed this week" findings from the week's data, used where
 * nobody has entered a weekly review input for that area.
 */
export function draftWeeklyFindings({ selectedKpis = {}, previousKpis = {}, breakdownRows = [], workOrderRows = [], selectedCosts = {}, previousCosts = {}, money = (v) => `$${Math.round(Number(v) || 0).toLocaleString("en-US")}` }) {
  const dt = Number(selectedKpis.downtime) || 0;
  const prevDt = Number(previousKpis.downtime) || 0;
  const worst = [...breakdownRows].sort((a, b) => (Number(b.downtime_hours) || 0) - (Number(a.downtime_hours) || 0))[0];
  const openDown = breakdownRows.filter((b) => String(b.status || "").toUpperCase() !== "CLOSED");
  const direction = dt > prevDt ? "up from" : dt < prevDt ? "down from" : "same as";
  const downtime = {
    finding: `${fmtHours(dt)} mechanical downtime, ${direction} ${fmtHours(prevDt)} last week.`
      + (worst && Number(worst.downtime_hours) > 0 ? ` Largest: ${worst.asset_code} ${fmtHours(worst.downtime_hours)}${worst.component || worst.description ? ` (${String(worst.component || worst.description).slice(0, 40)})` : ""}.` : ""),
    action: openDown.length
      ? `Return ${openDown.slice(0, 3).map((b) => b.asset_code).join(", ")} to service`
      : "No machines still down",
  };

  const done = workOrderRows.filter((w) => isWorkOrderDone(w.status));
  const open = workOrderRows.filter((w) => !isWorkOrderDone(w.status));
  const unassigned = open.filter((w) => !String(w.assigned_artisan_name || "").trim());
  const waiting = open.filter((w) => Number(w.open_parts_requests) > 0 || String(w.parts_status || "").trim());
  const repairs = {
    finding: `${done.length} job${done.length === 1 ? "" : "s"} finished, ${open.length} still open`
      + (waiting.length ? `, ${waiting.length} waiting on parts` : "") + ".",
    action: unassigned.length
      ? `Assign artisans to ${unassigned.length} open job${unassigned.length === 1 ? "" : "s"}`
      : open.length ? "Keep due dates current on open jobs" : "No open jobs",
  };

  const total = Number(selectedCosts.total) || 0;
  const prevTotal = Number(previousCosts.total) || 0;
  const change = pctChange(total, prevTotal);
  const costs = {
    finding: `${money(total)} this week (parts ${money(selectedCosts.partsIssued)}, labour ${money(selectedCosts.internalLabor)})`
      + (prevTotal ? ` vs ${money(prevTotal)} last week${change == null ? "" : ` (${change >= 0 ? "+" : ""}${change.toFixed(0)}%)`}.` : "; nothing recorded last week."),
    action: change != null && change > 25 ? "Review the cost increase" : "No cost action flagged",
  };
  return { Downtime: downtime, Repairs: repairs, Costs: costs };
}

// Production equipment for the forum's equipment KPIs. Support equipment —
// generators, welding machines, compressors and light vehicles (LDVs) — is
// left out of the rankings, low-use list and repair status.
const SUPPORT_EQUIPMENT = /generat|gen\s?set|weld|compress|\bldv\b|light (delivery )?vehicle|hilux|bakkie/i;
const SUPPORT_CODE = /^(GS|GEN|WM|CPR|COMP)\d|^V\d{2}AM$|^LDV/i;

export function isProductionAsset({ category, asset_name, asset_code } = {}) {
  if (SUPPORT_CODE.test(String(asset_code || "").trim())) return false;
  return !SUPPORT_EQUIPMENT.test(`${category || ""} ${asset_name || ""}`);
}

/**
 * Stock-out movements that are parts issued to a machine or job. Corrections
 * booked as stock-out (for example removing oil captured twice) are not issues
 * and must not be charged as maintenance cost.
 */
export function issuedToEquipmentSql(sm = "sm") {
  return `(
    ${sm}.reference LIKE 'work_order:%'
    OR ${sm}.reference LIKE 'asset:%'
    OR ${sm}.reference LIKE 'lube_issue:%'
  )`;
}
