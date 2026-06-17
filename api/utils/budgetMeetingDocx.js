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

function buildBvaForCats(curActMap, curBvaMap, curBudMap, cats) {
  const rows = cats.map((cat) => {
    const row = curBvaMap.get(cat) || {};
    const budget = Number(row.budget ?? curBudMap.get(cat) ?? 0);
    const actual = Number(curActMap.get(cat) || row.actual || 0);
    const variance = actual - budget;
    const variancePct = budget > 0 ? (variance / budget) * 100 : null;
    return [titleCaseCat(cat), money(budget), money(actual), money(variance), pct(variancePct)];
  });
  const budget = sumCats(
    new Map(cats.map((c) => [c, Number((curBvaMap.get(c) || {}).budget ?? curBudMap.get(c) ?? 0)])),
    cats,
  );
  const actual = sumCats(curActMap, cats);
  const variance = actual - budget;
  rows.push([
    "TOTAL (operating costs)",
    money(budget),
    money(actual),
    money(variance),
    pct(budget > 0 ? (variance / budget) * 100 : null),
  ]);
  return { rows, budget, actual, variance };
}

/**
 * @param {object} data
 */
export async function buildBudgetMeetingDocxBuffer(data) {
  const curActMap = new Map((data.currentActuals?.rows || []).map((r) => [r.category, r.amount]));
  const prevActMap = new Map((data.prevActuals?.rows || []).map((r) => [r.category, r.amount]));
  const curBudMap = new Map(
    (data.currentBva?.rows || []).map((r) => [String(r.dimension_key || "").split("|")[0] || r.dimension_key, r.budget]),
  );
  const curBvaMap = new Map(
    (data.currentBva?.rows || []).map((r) => [String(r.dimension_key || ""), r]),
  );

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

  const bva = buildBvaForCats(curActMap, curBvaMap, curBudMap, COST_CATS);

  const plantHireRows = (data.plantHireLines || []).map((r) => [
    r.asset_code,
    r.asset_name || "",
    r.billing_mode === "hourly" ? "Hourly" : "Fixed / month",
    r.billing_mode === "hourly" ? Number(r.hours_run || 0).toFixed(1) : "—",
    r.billing_mode === "hourly" ? `${money(r.rate)}/hr` : `${money(r.rate)}/mo`,
    money(r.amount),
  ]);

  const downtimeRows = (data.downtimeDetail || []).map((r) => [
    r.asset_code,
    r.asset_name || "",
    String(r.down_days || 0),
    Number(r.downtime_hours || 0).toFixed(1),
    `${money(r.downtime_rate)}/hr`,
    money(r.downtime_cost),
  ]);

  const up = data.upcoming || {};
  const children = [
    heading("Budget Meeting — Cost & Income Summary"),
    para(`Reporting month: ${data.periodLabel}  |  Previous month: ${data.prevPeriodLabel}  |  Site: ${data.siteCode}`),
    spacer(),
    heading("Executive Summary", HeadingLevel.HEADING_2),
    para(
      `Operating costs this month: ${money(bva.actual)} actual vs ${money(bva.budget)} budget ` +
        `(variance ${money(bva.variance)}).`,
    ),
    para(
      `Plant hire income: ${money(plantIncome)} actual vs ${money(plantBudget)} budget ` +
        `(variance ${money(plantVariance)}). Plant charges are contractor income, not an operating cost.`,
      true,
    ),
    para(
      `Upcoming: parts on order ${money(up.parts_on_order)} | in transit ${money(up.parts_in_transit)} | ` +
        `arrived ${money(up.parts_arrived)} | maintenance forecast ${money(up.maintenance_total)}.`,
    ),
    spacer(),
    heading("Last Month vs This Month (Operating Costs)", HeadingLevel.HEADING_2),
    para("Excludes plant hire income. Downtime uses scheduled daily hours on each down day × machine downtime rate."),
    docTable(["Category", data.prevPeriodLabel, data.periodLabel, "Change $", "Change %"], momRows),
    spacer(),
    heading("This Month — Budget vs Actual (Operating Costs)", HeadingLevel.HEADING_2),
    docTable(["Category", "Budget", "Actual", "Variance $", "Variance %"], bva.rows),
    spacer(),
    heading("Downtime Cost Detail", HeadingLevel.HEADING_2),
    para(
      downtimeRows.length
        ? "Each down day uses the scheduled hours from Daily Input for that date (e.g. 9 h × 10 days = 90 h). " +
          "Down days include breakdown downtime logs and days covered by an open breakdown or assigned work order."
        : "No scheduled-hours downtime cost calculated for this period.",
    ),
  ];

  if (downtimeRows.length) {
    children.push(
      docTable(
        ["Asset", "Name", "Down days", "Downtime hrs", "Rate", "Cost"],
        downtimeRows,
      ),
      para(
        `Downtime subtotal: ${money((data.downtimeDetail || []).reduce((s, r) => s + Number(r.downtime_cost || 0), 0))} ` +
          `(${Number(data.downtimeTotalHours || 0).toFixed(1)} hours)`,
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
    para(
      plantHireRows.length
        ? "Contractor plant charges are income to your operation. Hourly hire uses production hours logged in IRONLOG."
        : "No plant hire income calculated yet. Set hire rates under Assets → Plant Hire & Budget.",
    ),
  );

  if (plantHireRows.length) {
    children.push(
      docTable(
        ["Asset", "Name", "Billing", "Hours", "Rate", "Income"],
        plantHireRows,
      ),
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
