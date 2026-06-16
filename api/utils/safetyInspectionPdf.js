import {
  buildPdfBuffer,
  sectionTitle,
  kvGrid,
  table,
  ensurePageSpace,
} from "./pdfGenerator.js";

function contentWidth(doc) {
  return doc.page.width - doc.page.margins.left - doc.page.margins.right;
}

export function parseSafetyChecklistJson(raw) {
  try {
    const parsed = JSON.parse(String(raw || "[]"));
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function checklistOkLabel(ok) {
  if (ok === true) return "OK";
  if (ok === false) return "FAIL";
  return "N/A";
}

export function enrichSafetyInspectionRow(row) {
  const checklist = parseSafetyChecklistJson(row?.checklist_json);
  const failures = checklist.filter((x) => x && x.ok === false);
  const st = String(row?.status || "pending").toLowerCase();
  const is_flagged = st === "fail" || failures.length > 0;
  const statusLabel = is_flagged
    ? "FLAGGED"
    : st === "pass"
      ? "PASS"
      : st === "attention"
        ? "ATTENTION"
        : st.toUpperCase();
  return { ...row, checklist, failures, is_flagged, statusLabel };
}

export function drawSafetyChecklistTable(doc, checklist) {
  const rows = (checklist || []).map((c) => ({
    item: String(c.label || c.key || "—"),
    status: checklistOkLabel(c.ok),
    note: String(c.note || "").trim() || "—",
  }));
  if (!rows.length) {
    doc.font("Helvetica").fontSize(9).fillColor("#64748b");
    doc.text("No checklist items recorded.", doc.page.margins.left, doc.y, { width: contentWidth(doc) });
    doc.moveDown(0.4);
    return;
  }
  table(
    doc,
    [
      { key: "item", label: "Checklist item", width: 0.44 },
      { key: "status", label: "Result", width: 0.12, align: "center" },
      { key: "note", label: "Notes / findings", width: 0.44 },
    ],
    rows,
    { fontSize: 9, compact: true }
  );
}

export function drawSafetyInspectionBlock(doc, insp) {
  const r = enrichSafetyInspectionRow(insp);
  const nameBits = [String(r.item_name || "").trim(), String(r.location || "").trim()].filter(Boolean);

  ensurePageSpace(doc, 140);
  sectionTitle(doc, `${r.item_code || "Equipment"} — ${r.inspection_date || "—"}`);

  kvGrid(
    doc,
    [
      { k: "Equipment code", v: r.item_code || "—" },
      { k: "Name / location", v: nameBits.length ? nameBits.join(" · ") : "—" },
      { k: "Equipment type", v: String(r.template_title || r.template_key || "—") },
      { k: "Inspection date", v: String(r.inspection_date || "—") },
      { k: "Inspector", v: String(r.inspector_name || "—") },
      { k: "Overall status", v: r.statusLabel },
    ],
    2
  );

  if (r.failures.length) {
    doc.moveDown(0.25);
    doc.font("Helvetica-Bold").fontSize(9).fillColor("#b91c1c");
    doc.text("Flagged findings", doc.page.margins.left, doc.y, { width: contentWidth(doc) });
    doc.moveDown(0.15);
    doc.font("Helvetica").fontSize(9).fillColor("#334155");
    for (const f of r.failures) {
      const label = String(f.label || f.key || "Item").trim();
      const note = String(f.note || "").trim();
      const line = note ? `• ${label}: ${note}` : `• ${label}`;
      doc.text(line, doc.page.margins.left + 6, doc.y, { width: contentWidth(doc) - 6 });
      doc.moveDown(0.12);
    }
    doc.moveDown(0.2);
  }

  const notes = String(r.notes || "").trim();
  if (notes) {
    doc.font("Helvetica-Bold").fontSize(9).fillColor("#334155");
    doc.text("General notes", doc.page.margins.left, doc.y, { width: contentWidth(doc) });
    doc.moveDown(0.1);
    doc.font("Helvetica").fontSize(9).fillColor("#334155");
    doc.text(notes, doc.page.margins.left, doc.y, { width: contentWidth(doc) });
    doc.moveDown(0.35);
  }

  doc.font("Helvetica-Bold").fontSize(10).fillColor("#0b3a7e");
  doc.text("Checklist", doc.page.margins.left, doc.y, { width: contentWidth(doc) });
  doc.moveDown(0.2);
  drawSafetyChecklistTable(doc, r.checklist);
  doc.moveDown(0.5);
}

export async function buildSafetyInspectionReportPdf({
  inspections,
  start,
  end,
  selectedCount = 0,
}) {
  const enriched = (inspections || []).map(enrichSafetyInspectionRow);
  const flagged = enriched.filter((r) => r.is_flagged);
  const passCount = enriched.filter((r) => String(r.status || "").toLowerCase() === "pass").length;
  const attentionCount = enriched.filter((r) => String(r.status || "").toLowerCase() === "attention").length;
  const failCount = enriched.filter((r) => String(r.status || "").toLowerCase() === "fail").length;
  const itemCount = new Set(enriched.map((r) => String(r.item_code || ""))).size;

  return buildPdfBuffer(
    (doc) => {
      sectionTitle(doc, "Safety equipment inspection report");
      kvGrid(
        doc,
        [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Equipment selected", v: selectedCount ? String(selectedCount) : "All" },
          { k: "Inspection rows", v: String(enriched.length) },
          { k: "Equipment covered", v: String(itemCount) },
          { k: "Pass", v: String(passCount) },
          { k: "Attention", v: String(attentionCount) },
          { k: "Fail (flagged)", v: String(failCount) },
          { k: "Total flagged", v: String(flagged.length) },
        ],
        2
      );

      sectionTitle(doc, "Summary");
      table(
        doc,
        [
          { key: "date", label: "Date", width: 0.12 },
          { key: "equipment", label: "Equipment", width: 0.16 },
          { key: "name", label: "Name", width: 0.28 },
          { key: "status", label: "Status", width: 0.12, align: "center" },
          { key: "inspector", label: "Inspector", width: 0.14 },
          { key: "failed", label: "Failed items", width: 0.18, align: "center" },
        ],
        enriched.length
          ? enriched.map((r) => ({
              date: String(r.inspection_date || ""),
              equipment: String(r.item_code || "-"),
              name: String(r.item_name || "-"),
              status: r.statusLabel,
              inspector: String(r.inspector_name || "-"),
              failed: String(r.failures.length || 0),
            }))
          : [{
              date: "-",
              equipment: "-",
              name: "No inspections found for selected filter",
              status: "-",
              inspector: "-",
              failed: "-",
            }],
        { fontSize: 8, compact: true }
      );

      if (enriched.length) {
        sectionTitle(doc, "Inspection details");
        for (const insp of enriched) {
          drawSafetyInspectionBlock(doc, insp);
        }
      }
    },
    {
      title: "IRONLOG",
      subtitle: "Safety Equipment Inspection Report",
      rightText: `${start} to ${end}`,
      showPageNumbers: true,
    }
  );
}

export async function buildSingleSafetyItemInspectionPdf(inspection, { noInspectionMessage = null } = {}) {
  const r = enrichSafetyInspectionRow(inspection);
  return buildPdfBuffer(
    (doc) => {
      sectionTitle(doc, String(r.template_title || "Safety equipment inspection"));
      if (noInspectionMessage) {
        doc.font("Helvetica").fontSize(10).fillColor("#b45309");
        doc.text(noInspectionMessage, doc.page.margins.left, doc.y, { width: contentWidth(doc) });
        doc.moveDown(0.4);
      }
      drawSafetyInspectionBlock(doc, inspection);
    },
    {
      title: "IRONLOG",
      subtitle: String(r.template_title || "Safety Inspection"),
      rightText: `${r.item_code || "Equipment"} · ${r.inspection_date || "—"}`,
      showPageNumbers: true,
    }
  );
}
