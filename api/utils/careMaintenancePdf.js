// IRONLOG/api/utils/careMaintenancePdf.js — a care and maintenance plan as a PDF for meetings and the workshop wall.
//
// Page 1 is the plan at a glance: dates, labour rate, estimate against what has been used,
// and one line per machine. After that, every machine gets its work to do and its parts
// list with where each part stands (issued, in stores, ordered, not in stores).
import { ensurePageSpace, kvGrid, sectionTitle, table } from "./pdfGenerator.js";

export const PART_STATE_LABELS = {
  issued: "Issued",
  in_stock: "In stores",
  requested: "Requested",
  ordered: "Ordered",
  received: "Received",
  short: "Not in stores",
};

const money = (v) => `$${Number(v || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const hours = (v) => `${Number(v || 0).toLocaleString("en-US", { maximumFractionDigits: 1 })} h`;
const cap = (s) => String(s || "").charAt(0).toUpperCase() + String(s || "").slice(1);

/** Draws the plan into a PDFKit document (used inside buildPdfBuffer). */
export function drawCarePlanPdf(doc, detail) {
  const { plan, machines, totals } = detail;
  const dates = [plan.start_date, plan.end_date].filter(Boolean).join(" to ") || "No dates set";
  const statusLine = Object.entries(detail.statuses || {})
    .filter(([k]) => totals.by_status?.[k])
    .map(([k, label]) => `${totals.by_status[k]} ${label.toLowerCase()}`)
    .join(", ") || "No machines yet";

  sectionTitle(doc, `Plan summary: ${plan.name}`);
  kvGrid(doc, [
    { k: "Period", v: dates },
    { k: "Plan status", v: plan.status === "closed" ? "Closed" : "Active" },
    { k: "Labour rate", v: `${money(plan.labour_rate_used)} per hour` },
    { k: "Machines", v: `${totals.machines} (${statusLine})` },
    { k: "Estimate", v: `${money(totals.planned_total)} (parts ${money(totals.planned_parts)}, labour ${money(totals.planned_labour)}, ${hours(totals.planned_hours)})` },
    { k: "Used so far", v: `${money(totals.actual_total)} (parts ${money(totals.actual_parts)}, labour ${money(totals.actual_labour)}, ${hours(totals.actual_hours)})` },
    { k: "Parts still needed", v: totals.parts_short ? `${totals.parts_short} part line${totals.parts_short === 1 ? "" : "s"}` : "None" },
  ], 2);
  if (plan.notes) {
    doc.font("Helvetica").fontSize(9).fillColor("#334155").text(`Notes: ${plan.notes}`, doc.page.margins.left, doc.y, { width: doc.page.width - doc.page.margins.left - doc.page.margins.right });
  }

  sectionTitle(doc, "Machines");
  if (!machines.length) {
    doc.font("Helvetica").fontSize(10).text("No machines on this plan yet.");
  } else {
    table(doc, [
      { key: "machine", label: "Machine", width: 0.09 },
      { key: "name", label: "Description", width: 0.14 },
      { key: "priority", label: "Priority", width: 0.07 },
      { key: "status", label: "Status", width: 0.12 },
      { key: "wo", label: "Work order", width: 0.1 },
      { key: "hours", label: "Hours plan / used", width: 0.13, align: "right" },
      { key: "planned", label: "Estimate", width: 0.1, align: "right" },
      { key: "used", label: "Used", width: 0.1, align: "right" },
      { key: "pct", label: "Used %", width: 0.07, align: "right" },
      { key: "short", label: "To order", width: 0.08, align: "right" },
    ], [
      ...machines.map((m) => ({
        machine: m.asset_code || "",
        name: m.asset_name || m.category || "",
        priority: cap(m.priority),
        status: m.status_label,
        wo: m.work_order ? `#${m.work_order.id} ${m.work_order.status || ""}`.trim() : "Not opened",
        hours: `${hours(m.planned.hours)} / ${hours(m.actual.hours)}`,
        planned: money(m.planned.total),
        used: money(m.actual.total),
        pct: m.hours_used_pct == null ? "" : `${m.hours_used_pct}%`,
        short: m.parts_short || "",
      })),
      {
        machine: "TOTAL", name: `${totals.machines} machines`, priority: "", status: "", wo: "",
        hours: `${hours(totals.planned_hours)} / ${hours(totals.actual_hours)}`,
        planned: money(totals.planned_total), used: money(totals.actual_total),
        pct: totals.planned_hours > 0 ? `${Math.round((totals.actual_hours / totals.planned_hours) * 100)}%` : "",
        short: totals.parts_short || "",
      },
    ]);
  }

  if (!machines.length) return;
  doc.addPage();
  sectionTitle(doc, "Work and parts per machine");
  const left = () => doc.page.margins.left;
  const width = () => doc.page.width - doc.page.margins.left - doc.page.margins.right;
  for (const m of machines) {
    ensurePageSpace(doc, 90);
    doc.moveDown(0.4);
    doc.font("Helvetica-Bold").fontSize(11).fillColor("#0f172a")
      .text(`${m.asset_code}${m.asset_name ? `  ${m.asset_name}` : ""}`, left(), doc.y, { width: width() });
    const meta = [`${cap(m.priority)} priority`, m.status_label, m.work_order ? `WO #${m.work_order.id} (${m.work_order.status})` : "No work order yet"];
    if (m.work_order?.technician) meta.push(`Technician: ${m.work_order.technician}`);
    meta.push(`Estimate ${money(m.planned.total)}, used ${money(m.actual.total)}`);
    doc.font("Helvetica").fontSize(9).fillColor("#475569").text(meta.join("  ·  "), left(), doc.y, { width: width() });
    doc.fillColor("#0f172a").fontSize(10);
    doc.moveDown(0.2);
    doc.font("Helvetica-Bold").text("Work to do: ", left(), doc.y, { continued: true, width: width() })
      .font("Helvetica").text(m.scope || "Not written yet.");
    if (m.notes) doc.font("Helvetica-Bold").text("Notes: ", left(), doc.y, { continued: true, width: width() }).font("Helvetica").text(m.notes);
    doc.moveDown(0.3);
    if (m.parts.length) {
      table(doc, [
        { key: "code", label: "Part code", width: 0.14 },
        { key: "desc", label: "Description", width: 0.36 },
        { key: "qty", label: "Qty", width: 0.07, align: "right" },
        { key: "unit", label: "Unit cost", width: 0.11, align: "right" },
        { key: "line", label: "Line cost", width: 0.11, align: "right" },
        { key: "issued", label: "Issued", width: 0.07, align: "right" },
        { key: "state", label: "Where it stands", width: 0.14 },
      ], m.parts.map((p) => ({
        code: p.part_code || "",
        desc: p.description || "",
        qty: p.qty,
        unit: money(p.unit_cost),
        line: money(p.line_cost),
        issued: p.issued || "",
        state: PART_STATE_LABELS[p.state] || p.state,
      })), { compact: true });
    } else {
      doc.font("Helvetica").fontSize(9).fillColor("#475569").text("No parts planned.", left(), doc.y, { width: width() });
      doc.fillColor("#0f172a");
    }
    doc.moveDown(0.6);
  }
}
