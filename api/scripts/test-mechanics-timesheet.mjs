import { db } from "../db/client.js";
import { generateMechanicsTimesheet, mechanicsTimesheetToExportRows } from "../utils/mechanicsTimesheetGenerator.js";

const { rows, meta } = generateMechanicsTimesheet(db, {
  from: "2026-01-02",
  to: "2026-01-15",
  syntheticThrough: "2026-03-31",
});
const exportRows = mechanicsTimesheetToExportRows(rows);
console.log("meta", meta);

const cats = {};
const descCounts = {};
for (const r of exportRows) {
  cats[r.Category] = (cats[r.Category] || 0) + 1;
  const k = `${r.Technician}|${r["Description Of Work Carried Out"]}`;
  descCounts[k] = (descCounts[k] || 0) + 1;
}
console.log("categories", cats);
const repeats = Object.entries(descCounts).filter(([, c]) => c > 3);
console.log("repeats (>3 per tech)", repeats.slice(0, 8));

const jan2 = exportRows.filter((r) => r.work_date === "2026-01-02");
console.log("jan2 tech hours", jan2.reduce((acc, r) => {
  acc[r.Technician] = (acc[r.Technician] || 0) + Number(r["Work Hours"] || 0);
  return acc;
}, {}));
console.log("jan2 sample", jan2.slice(0, 6).map((r) => ({
  cat: r.Category, desc: r["Description Of Work Carried Out"], tech: r.Technician,
})));
