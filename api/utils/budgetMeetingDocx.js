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

const COST_CATS = ["parts", "labor", "fuel", "lube", "downtime"];

function sumCats(map, cats) {
  return cats.reduce((s, c) => s + Number(map.get(c) || 0), 0);
}

/**
 * @param {object} data
 * @param {number} data.operatingBudget - single monthly expense budget (all categories)
 * @param {number} data.plantHireBudget - monthly plant hire income target
 */
export async function buildBudgetMeetingDocxBuffer(data) {
  const curActMap = new Map((data.currentActuals?.rows || []).map((r) => [r.category, r.amount]));
  const prevActMap = new Map((data.prevActuals?.rows || []).map((r) => [r.category, r.amount]));

  const operatingActual = sumCats(curActMap, COST_CATS);
  const operatingBudget = Number(data.operatingBudget || 0);
  const operatingVariance = operatingActual - operatingBudget;

  const plantIncome = Number(curActMap.get("plant_hire") || 0);
  const plantBudget = Number(data.plantHireBudget || 0);
  const plantVariance = plantIncome - plantBudget;

  const momRows = COST_CATS.map((cat) => {
    const prev = Number(prevActMap.get(cat) || 0);
    const cur = Number(curActMap.get(cat) || 0);
    const chg = cur - prev;
    const chgPct = prev > 0 ? (chg / prev) * 100 : null;
    return [titleCaseCat(cat), money(prev), money(cur), money(chg), pct(chgPct)];
  });
  const momPrevTotal = sumCats(prevActMap, COST_CATS);
  const momCurTotal = sumCats(curActMap, COST_CATS);
  momRows.push([
    "TOTAL (operating costs)",
    money(momPrevTotal),
    money(momCurTotal),
    money(momCurTotal - momPrevTotal),
    pct(momPrevTotal > 0 ? ((momCurTotal - momPrevTotal) / momPrevTotal) * 100 : null),
  ]);

  const actualBreakdownRows = COST_CATS.map((cat) => {
    const actual = Number(curActMap.get(cat) || 0);
    const share = operatingActual > 0 ? (actual / operatingActual) * 100 : null;
    return [titleCaseCat(cat), money(actual), pct(share)];
  });
  actualBreakdownRows.push(["TOTAL (operating costs)", money(operatingActual), operatingActual > 0 ? "100.0%" : "—"]);

  const plantHireRows = (data.plantHireLines || []).map((r) => [
    r.asset_code,
    r.asset_name || "",
    r.billing_mode === "hourly" ? "Hourly" : "Fixed / month",
    r.billing_mode === "hourly" ? Number(r.hours_run || 0).toFixed(1) : "—",
    r.billing_mode === "hourly" ? `${money(r.rate)}/hr` : `${money(r.rate)}/mo`,
    money(r.amount),
  ]);

  const laborDetail = Array.isArray(data.mechanicLaborDetail) ? data.mechanicLaborDetail : [];
  const laborRows = laborDetail.map((r) => [
    r.technician_name || "Unassigned",
    Number(r.hours || 0).toFixed(1),
    `${money(r.rate)}/hr`,
    money(r.labor_cost),
  ]);
  const laborTotalCost = Number(
    data.mechanicLaborTotalCost != null
      ? data.mechanicLaborTotalCost
      : laborDetail.reduce((s, r) => s + Number(r.labor_cost || 0), 0),
  );
  const laborTotalHours = Number(
    data.mechanicLaborTotalHours != null
      ? data.mechanicLaborTotalHours
      : laborDetail.reduce((s, r) => s + Number(r.hours || 0), 0),
  );

  const up = data.upcoming || {};
  const children = [
    heading("Budget Meeting — Cost & Income Summary"),
    para(`Reporting month: ${data.periodLabel}  |  Previous month: ${data.prevPeriodLabel}  |  Site: ${data.siteCode}`),
    spacer(),
    heading("Executive Summary", HeadingLevel.HEADING_2),
    para(
      `Operating costs: ${money(operatingActual)} actual vs ${money(operatingBudget)} monthly budget ` +
        `(variance ${money(operatingVariance)}).`,
    ),
    para(
      `Mechanics labor (Mechanics Cost log): ${money(laborTotalCost)} ` +
        `(${laborTotalHours.toFixed(1)} hours across ${laborDetail.length} technician${laborDetail.length === 1 ? "" : "s"}).`,
      true,
    ),
    para(
      `Plant hire income: ${money(plantIncome)} actual vs ${money(plantBudget)} income target ` +
        `(variance ${money(plantVariance)}). Income from contractor plant charges is separate from operating costs.`,
      true,
    ),
    para(
      `Upcoming: parts on order ${money(up.parts_on_order)} | in transit ${money(up.parts_in_transit)} | ` +
        `arrived ${money(up.parts_arrived)} | maintenance forecast ${money(up.maintenance_total)}.`,
    ),
    spacer(),
    heading("Monthly Budget vs Actual (Operating Costs)", HeadingLevel.HEADING_2),
    para("One total monthly budget for all operating expenses — breakdown by category is actual spend only."),
    docTable(
      ["", "Budget", "Actual", "Variance $", "Variance %"],
      [
        [
          "TOTAL operating costs",
          money(operatingBudget),
          money(operatingActual),
          money(operatingVariance),
          pct(operatingBudget > 0 ? (operatingVariance / operatingBudget) * 100 : null),
        ],
      ],
    ),
    spacer(),
    heading("This Month — Actual Spend by Category", HeadingLevel.HEADING_2),
    docTable(["Category", "Actual", "% of total"], actualBreakdownRows),
    spacer(),
    heading("Last Month vs This Month (Operating Costs)", HeadingLevel.HEADING_2),
    docTable(["Category", data.prevPeriodLabel, data.periodLabel, "Change $", "Change %"], momRows),
    spacer(),
    heading("Mechanics Labor Cost", HeadingLevel.HEADING_2),
    para(
      laborRows.length
        ? "From Mechanics Cost daily labor entries for the selected month — mechanic, effective rate, and period total."
        : "No mechanics labor logged for this period under Maintenance → Mechanics Cost.",
    ),
  ];

  if (laborRows.length) {
    children.push(
      docTable(["Mechanic", "Hours", "Rate", "Labor $"], laborRows),
      para(
        `Total labor: ${money(laborTotalCost)} (${laborTotalHours.toFixed(1)} hours)` +
          (data.mechanicLaborDefaultRate != null
            ? ` | Default rate when blank: ${money(data.mechanicLaborDefaultRate)}/hr`
            : ""),
        true,
      ),
      spacer(),
    );
  } else {
    children.push(spacer());
  }

  children.push(
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
    heading("Plant Hire Income — Rates & This Month", HeadingLevel.HEADING_2),
    docTable(
      ["", "Income target", "Actual income", "Variance $", "Variance %"],
      [
        [
          "Plant hire (month)",
          money(plantBudget),
          money(plantIncome),
          money(plantVariance),
          pct(plantBudget > 0 ? (plantVariance / plantBudget) * 100 : null),
        ],
      ],
    ),
    para(
      plantHireRows.length
        ? "Detail below — hourly hire uses production hours logged in IRONLOG."
        : "Set contractor hire rates under Assets → Plant Hire & Budget to calculate income.",
    ),
  );

  if (plantHireRows.length) {
    children.push(
      docTable(["Asset", "Name", "Billing", "Hours", "Rate", "Income"], plantHireRows),
      para(`Plant hire income subtotal: ${money(plantIncome)}`, true),
    );
  }

  children.push(
    spacer(),
    para(`Generated by IRONLOG on ${new Date().toISOString().slice(0, 16).replace("T", " ")} UTC`, false),
  );

  const doc = new Document({ sections: [{ children }] });
  return Packer.toBuffer(doc);
}
