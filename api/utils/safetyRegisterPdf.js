import path from "node:path";
import {
  buildPdfBuffer,
  sectionTitle,
  kvGrid,
  table,
  ensurePageSpace,
  tryDrawLogo,
} from "./pdfGenerator.js";

function contentWidth(doc) {
  return doc.page.width - doc.page.margins.left - doc.page.margins.right;
}

function checklistCell(row) {
  if (row?.ok === true) return "OK";
  if (row?.ok === false) return "FAIL";
  return "";
}

function buildColumns(templateItems) {
  const chkCount = Math.max(1, templateItems.length);
  const fixed = { code: 0.1, name: 0.18, location: 0.14, inspector: 0.1, notes: 0.12 };
  const fixedSum = Object.values(fixed).reduce((a, b) => a + b, 0);
  const chkW = Math.max(0.05, (1 - fixedSum) / chkCount);

  const columns = [
    { key: "code", label: "Code", width: fixed.code },
    { key: "name", label: "Name", width: fixed.name },
    { key: "location", label: "Location", width: fixed.location },
    ...templateItems.map((it, i) => ({
      key: `chk_${i}`,
      label: String(i + 1),
      width: chkW,
      align: "center",
    })),
    { key: "inspector", label: "Inspector", width: fixed.inspector },
    { key: "notes", label: "Notes", width: fixed.notes },
  ];

  const used = columns.slice(0, -1).reduce((s, c) => s + c.width, 0);
  columns[columns.length - 1].width = Math.max(0.08, 1 - used);
  return columns;
}

function buildRows(templateItems, items, blank) {
  return items.map(({ item, checklist, inspection }) => {
    const row = {
      code: String(item.item_code || ""),
      name: String(item.item_name || ""),
      location: String(item.location || ""),
      inspector: blank ? "" : String(inspection?.inspector_name || ""),
      notes: blank ? "" : String(inspection?.notes || ""),
    };
    templateItems.forEach((it, i) => {
      if (blank) {
        row[`chk_${i}`] = "";
        return;
      }
      const hit = (checklist || []).find((r) => r.key === it.key);
      row[`chk_${i}`] = checklistCell(hit);
    });
    return row;
  });
}

function drawChecklistLegend(doc, templateItems) {
  if (!templateItems.length) return;
  ensurePageSpace(doc, 40);
  doc.font("Helvetica-Bold").fontSize(9).fillColor("#334155");
  doc.text("Checklist items:", doc.page.margins.left, doc.y, { width: contentWidth(doc) });
  doc.moveDown(0.15);
  doc.font("Helvetica").fontSize(8).fillColor("#0f172a");
  templateItems.forEach((it, idx) => {
    doc.text(`${idx + 1}. ${String(it.label || it.key || "")}`, doc.page.margins.left + 6, doc.y, {
      width: contentWidth(doc) - 6,
    });
  });
  doc.moveDown(0.35);
}

/**
 * @param {{ groups: Array<{ template: object, items: Array<object> }>, checkDate: string, blank?: boolean }} opts
 */
export async function buildSafetyRegisterPdf({ groups, checkDate, blank = false }) {
  const logoPath = path.join(process.cwd(), "branding", "logo.png");
  const safeGroups = Array.isArray(groups) ? groups : [];
  const totalItems = safeGroups.reduce((n, g) => n + (g.items?.length || 0), 0);

  return buildPdfBuffer(
    (doc) => {
      tryDrawLogo(doc, logoPath);

      sectionTitle(doc, "Safety Equipment Register");
      kvGrid(
        doc,
        [
          { k: "Inspection date", v: checkDate },
          { k: "Equipment types", v: safeGroups.length },
          { k: "Total items", v: totalItems },
          { k: "Sheet type", v: blank ? "Blank (manual walk-around)" : "Register (with captured results)" },
        ],
        2
      );

      if (!safeGroups.length) {
        doc.font("Helvetica").fontSize(10).fillColor("#555555");
        doc.text("No safety equipment registered for the selected type.", doc.page.margins.left, doc.y, {
          width: contentWidth(doc),
        });
        return;
      }

      for (const group of safeGroups) {
        const template = group.template || {};
        const templateItems = Array.isArray(template.items) ? template.items : [];
        const items = Array.isArray(group.items) ? group.items : [];
        if (!items.length) continue;

        ensurePageSpace(doc, 140);
        sectionTitle(doc, String(template.title || template.template_key || "Safety equipment"));
        drawChecklistLegend(doc, templateItems);

        const columns = buildColumns(templateItems);
        const rows = buildRows(templateItems, items, blank);
        table(doc, columns, rows, { fontSize: 8, headerFontSize: 8, compact: true });
      }
    },
    {
      title: "IRONLOG",
      subtitle: "Safety Equipment Register",
      rightText: `Date: ${checkDate}`,
      showPageNumbers: true,
      layout: "landscape",
    }
  );
}
