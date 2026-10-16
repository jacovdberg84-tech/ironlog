// deploy/site-counts.cjs — read-only: how records are spread across site codes in one IronLog database.
// Usage: node deploy/site-counts.cjs <path to ironlog.db>   (run by the server-check workflow)
const path = require("path");
const { createRequire } = require("module");
const Database = createRequire(path.join(__dirname, "../api/package.json"))("better-sqlite3");

const d = new Database(process.argv[2], { readonly: true, fileMustExist: true });
const tables = ["assets", "work_orders", "breakdowns", "daily_hours", "fuel_logs", "oil_logs", "stock_deliveries",
  "procurement_purchase_orders", "tech_shifts", "manager_inspections", "vehicle_ldv_checks"];
for (const t of tables) {
  try {
    const rows = d.prepare(`SELECT COALESCE(NULLIF(TRIM(site_code), ''), '(blank)') AS s, COUNT(*) AS n FROM ${t} GROUP BY 1 ORDER BY 2 DESC`).all();
    console.log(t.padEnd(30), rows.map((r) => `${r.s}=${r.n}`).join("  ") || "(empty)");
  } catch (e) {
    console.log(t.padEnd(30), `- ${e.message}`);
  }
}
try {
  const u = d.prepare(`SELECT COUNT(*) AS n, SUM(CASE WHEN allowed_locations IS NOT NULL AND allowed_locations NOT IN ('', '[]', 'null') THEN 1 ELSE 0 END) AS limited FROM users WHERE COALESCE(active, 1) = 1`).get();
  console.log("users".padEnd(30), `active=${u.n}  limited to some sites=${u.limited || 0}`);
  for (const r of d.prepare(`SELECT allowed_locations AS a, COUNT(*) AS n FROM users WHERE COALESCE(active, 1) = 1 AND allowed_locations IS NOT NULL AND allowed_locations NOT IN ('', '[]', 'null') GROUP BY 1`).all()) {
    console.log("".padEnd(30), `${r.n} user(s) limited to ${r.a}`);
  }
} catch (e) {
  console.log("users".padEnd(30), `- ${e.message}`);
}
try {
  console.log("site_profiles".padEnd(30), d.prepare(`SELECT GROUP_CONCAT(site_code, ', ') AS s FROM site_profiles`).get().s || "(none)");
} catch {
  console.log("site_profiles".padEnd(30), "(none)");
}
