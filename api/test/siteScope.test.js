import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Site isolation guard. A route handler that reads or changes a table holding site
// records must say which site it means (site_code, getSiteCode, x-site-code…). The
// handlers that did not when this guard was added are listed in siteScopeBaseline.json;
// that list may only shrink. Fixing one? Remove its line. Adding a handler? Scope it.
const here = path.dirname(fileURLToPath(import.meta.url));
const routesDir = path.join(here, "..", "routes");
const baselineFile = path.join(here, "siteScopeBaseline.json");

export const SITE_TABLES = ["assets", "daily_hours", "fuel_logs", "oil_logs", "work_orders", "breakdowns", "breakdown_downtime_logs",
  "ops_slip_reports", "breakdown_offsite_repairs", "tasks", "projects", "stores_part_orders", "stock_deliveries", "manager_inspections",
  "artisan_inspections", "manager_damage_reports", "maintenance_parts_requests", "mechanic_labor_entries", "tyre_inspections",
  "undercarriage_inspections", "vehicle_ldv_checks", "procurement_requisitions", "procurement_purchase_orders", "tech_shifts",
  "tech_other_time", "operations_logs", "dispatch_trips"];

function routeFiles(dir) {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => {
    const p = path.join(dir, e.name);
    return e.isDirectory() ? routeFiles(p) : p.endsWith(".js") ? [p] : [];
  });
}

export function unscopedHandlers() {
  const found = [];
  for (const file of routeFiles(routesDir)) {
    const src = fs.readFileSync(file, "utf8");
    const rel = path.relative(routesDir, file).split(path.sep).join("/");
    const heads = [...src.matchAll(/app\.(get|post|put|patch|delete)\(\s*["'`]([^"'`]+)["'`]/g)];
    heads.forEach((m, k) => {
      const body = src.slice(m.index, k + 1 < heads.length ? heads[k + 1].index : src.length);
      const touches = SITE_TABLES.some((t) => new RegExp(`\\b(FROM|JOIN|INTO|UPDATE)\\s+${t}\\b`, "i").test(body));
      // Scoped: names a site, uses a helper from utils/sites.js, or carries a "// site:" note saying why
      // it needs no filter (a person's own records, or a record number the guard in utils/siteGuard.js checks).
      const scoped = /site_code|getSiteCode|x-site-code|siteCode|siteOf|\bsite\b|andSite|andAssetSite|siteSql|requestSite|RequestSite|\/\/ site:/i.test(body);
      if (touches && !scoped) found.push(`${m[1].toUpperCase()} ${rel} ${m[2]}`);
    });
  }
  return [...new Set(found)].sort();
}

test("no new route reads or changes site records without choosing a site", () => {
  const baseline = JSON.parse(fs.readFileSync(baselineFile, "utf8"));
  const now = unscopedHandlers();
  const added = now.filter((h) => !baseline.includes(h));
  assert.deepEqual(added, [], `These handlers touch site records without a site filter. Filter by the request's site (getSiteCode(req)):\n${added.join("\n")}`);
  const fixed = baseline.filter((h) => !now.includes(h));
  assert.deepEqual(fixed, [], `Now site-scoped (well done): remove these lines from test/siteScopeBaseline.json:\n${fixed.join("\n")}`);
});
