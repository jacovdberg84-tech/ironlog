import { db } from "../db/client.js";
import { generateMechanicsTimesheet, mechanicsTimesheetToExportRows } from "../utils/mechanicsTimesheetGenerator.js";

const { rows, meta } = generateMechanicsTimesheet(db, {
  from: "2026-01-02",
  to: "2026-01-05",
  syntheticThrough: "2026-03-31",
});
const exportRows = mechanicsTimesheetToExportRows(rows);
console.log("meta", meta);
console.log("sample", exportRows.slice(0, 8));
console.log("jan2 tech hours", exportRows
  .filter((r) => r.work_date === "2026-01-02")
  .reduce((acc, r) => {
    acc[r.Technician] = (acc[r.Technician] || 0) + Number(r["Work Hours"] || 0);
    return acc;
  }, {}));
