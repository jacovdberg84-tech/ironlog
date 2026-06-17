import {
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

function money(n) {
  return `$${Number(n || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function pct(n) {
  if (n == null || !Number.isFinite(Number(n))) return "—";
  return `${Number(n).toFixed(1)}%`;
}

function titleCaseCat(cat) {
  const s = String(cat || "").replace(/_/g, " ");
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : "Other";
}

function docTable(headers, rows) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: headers.map(
          (h) =>
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: h, bold: true })] })],
            }),
        ),
      }),
      ...rows.map(
        (cells) =>
          new TableRow({
            children: cells.map((c) => new TableCell({ children: [new Paragraph(String(c ?? ""))] })),
          }),
      ),
    ],
  });
}

function heading(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({ heading: level, children: [new TextRun(text)] });
}

function para(text, bold = false) {
  return new Paragraph({ children: [new TextRun({ text, bold })] });
}

function spacer() {
  return new Paragraph({ children: [new TextRun("")] });
}

/**
 * @param {object} data
 * @param {string} data.periodLabel
 * @param {string} data.prevPeriodLabel
 * @param {string} data.siteCode
 * @param {object} data.currentBva - budgets-vs-actual total + rows by category
 * @param {object} data.prevBva
 * @param {object} data.currentActuals - actuals-by-category
 * @param {object} data.prevActuals
 * @param {object} data.upcoming
 * @param {Array} data.plantHireLines
 * @param {number} data.plantHireBudget
 */
export async function buildBudgetMeetingDocxBuffer(data) {
  const cats = ["parts", "labor", "fuel", "lube", "downtime", "plant_hire"];
  const curActMap = new Map((data.currentActuals?.rows || []).map((r) => [r.category, r.amount]));
  const prevActMap = new Map((data.prevActuals?.rows || []).map((r) => [r.category, r.amount]));
  const curBudMap = new Map(
    (data.currentBva?.rows || []).map((r) => [String(r.dimension_key || "").split("|")[0] || r.dimension_key, r.budget]),
  );
  const curBvaMap = new Map(
    (data.currentBva?.rows || []).map((r) => [String(r.dimension_key || ""), r]),
  );

  const momRows = cats.map((cat) => {
    const prev = Number(prevActMap.get(cat) || 0);
    const cur = Number(curActMap.get(cat) || 0);
    const chg = cur - prev;
    const chgPct = prev > 0 ? (chg / prev) * 100 : null;
    return [titleCaseCat(cat), money(prev), money(cur), money(chg), pct(chgPct)];
  });
  const momTotalPrev = Number(data.prevActuals?.total || 0);
  const momTotalCur = Number(data.currentActuals?.total || 0);
  momRows.push([
    "TOTAL",
    money(momTotalPrev),
    money(momTotalCur),
    money(momTotalCur - momTotalPrev),
    pct(momTotalPrev > 0 ? ((momTotalCur - momTotalPrev) / momTotalPrev) * 100 : null),
  ]);

  const bvaRows = cats.map((cat) => {
    const row = curBvaMap.get(cat) || {};
    const budget = Number(row.budget ?? curBudMap.get(cat) ?? 0);
    const actual = Number(curActMap.get(cat) || row.actual || 0);
    const variance = actual - budget;
    const variancePct = budget > 0 ? (variance / budget) * 100 : null;
    return [titleCaseCat(cat), money(budget), money(actual), money(variance), pct(variancePct)];
  });
  const t = data.currentBva?.total || {};
  bvaRows.push([
    "TOTAL",
    money(t.budget),
    money(t.actual),
    money(t.variance),
    pct(t.budget > 0 ? (Number(t.variance || 0) / Number(t.budget || 1)) * 100 : null),
  ]);

  const plantHireRows = (data.plantHireLines || []).map((r) => [
    r.asset_code,
    r.asset_name || "",
    r.billing_mode === "hourly" ? "Hourly" : "Fixed / month",
    r.billing_mode === "hourly" ? Number(r.hours_run || 0).toFixed(1) : "—",
    r.billing_mode === "hourly" ? money(r.rate) + "/hr" : money(r.rate) + "/mo",
    money(r.amount),
  ]);

  const up = data.upcoming || {};
  const children = [
    heading("Budget Meeting — Cost Summary"),
    para(`Reporting month: ${data.periodLabel}  |  Previous month: ${data.prevPeriodLabel}  |  Site: ${data.siteCode}`),
    spacer(),
    heading("Executive Summary", HeadingLevel.HEADING_2),
    para(
      `This month actual ${money(t.actual)} vs budget ${money(t.budget)} (variance ${money(t.variance)}). ` +
        `Plant hire budget: ${money(data.plantHireBudget)} | Plant hire actual (computed): ${money(curActMap.get("plant_hire") || 0)}.`,
    ),
    para(
      `Upcoming: parts on order ${money(up.parts_on_order)} | in transit ${money(up.parts_in_transit)} | ` +
        `arrived ${money(up.parts_arrived)} | maintenance forecast ${money(up.maintenance_total)}.`,
      true,
    ),
    spacer(),
    heading("Last Month vs This Month (Actuals)", HeadingLevel.HEADING_2),
    docTable(["Category", data.prevPeriodLabel, data.periodLabel, "Change $", "Change %"], momRows),
    spacer(),
    heading("This Month — Budget vs Actual", HeadingLevel.HEADING_2),
    docTable(["Category", "Budget", "Actual", "Variance $", "Variance %"], bvaRows),
    spacer(),
    heading("Upcoming Costs", HeadingLevel.HEADING_2),
    docTable(
      ["Item", "Amount"],
      [
        ["Parts on order", money(up.parts_on_order)],
        ["Parts in transit", money(up.parts_in_transit)],
        ["Parts arrived (period)", money(up.parts_arrived)],
        ["Upcoming maintenance (forecast)", money(up.maintenance_total)],
        ["TOTAL upcoming (parts + maintenance)", money(up.upcoming_total)],
      ],
    ),
    spacer(),
    heading("Plant Hire — Rates & This Month Cost", HeadingLevel.HEADING_2),
    para(
      plantHireRows.length
        ? "Hourly hire uses production hours logged in IRONLOG. Fixed monthly charges apply once per asset per month."
        : "No plant hire costs calculated yet. Set hire rates under Assets → Plant Hire & Budget.",
    ),
  ];

  if (plantHireRows.length) {
    children.push(
      docTable(
        ["Asset", "Name", "Billing", "Hours", "Rate", "Month cost"],
        plantHireRows,
      ),
      para(`Plant hire subtotal: ${money((data.plantHireLines || []).reduce((s, r) => s + Number(r.amount || 0), 0))}`, true),
    );
  }

  children.push(
    spacer(),
    para(`Generated by IRONLOG on ${new Date().toISOString().slice(0, 16).replace("T", " ")} UTC`, false),
  );

  const doc = new Document({ sections: [{ children }] });
  return Packer.toBuffer(doc);
}
