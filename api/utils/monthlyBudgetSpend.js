// IRONLOG/api/utils/monthlyBudgetSpend.js
// Source model for the monthly departmental costing export. Operational source
// entries remain untouched; the model adds transparent allocation columns so
// Plant and Workshop can reconcile without double-counting internal charges.
import { getOperatingBudgetAmount } from "./monthlyBudget.js";
import { ensureStockCategorySchema, oilPartSql } from "./stockCategory.js";
import { fetchLubeUsageLines } from "./lubeUsageLines.js";
import { buildPlantHireLines } from "./plantHire.js";
import {
  PLANT_DEPARTMENT,
  WORKSHOP_DEPARTMENT,
  listDepartmentalCostingInputs,
} from "./departmentalCostingInputs.js";

function hasTable(db, name) {
  try {
    return Boolean(db.prepare("SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ? LIMIT 1").get(name));
  } catch {
    return false;
  }
}

function hasColumn(db, table, column) {
  if (!hasTable(db, table)) return false;
  try {
    return db.prepare(`PRAGMA table_info(${table})`).all().some((row) => String(row.name) === column);
  } catch {
    return false;
  }
}

function num(value, digits = 2) {
  const parsed = Number(value || 0);
  return Number.isFinite(parsed) ? Number(parsed.toFixed(digits)) : 0;
}

function monthBounds(month) {
  const [year, monthNumber] = String(month || "").split("-").map(Number);
  if (!Number.isInteger(year) || !Number.isInteger(monthNumber) || monthNumber < 1 || monthNumber > 12) return null;
  const start = `${year}-${String(monthNumber).padStart(2, "0")}-01`;
  const end = new Date(Date.UTC(year, monthNumber, 0)).toISOString().slice(0, 10);
  return { start, end };
}

function normalizedSiteCode(value) {
  return String(value || "main").trim().toLowerCase() || "main";
}

function readDefaultRate(db) {
  if (!hasTable(db, "cost_settings")) return 80;
  try {
    const row = db.prepare("SELECT value FROM cost_settings WHERE key = 'labor_cost_per_hour_default' LIMIT 1").get();
    const rate = Number(row?.value);
    return Number.isFinite(rate) && rate > 0 ? rate : 80;
  } catch {
    return 80;
  }
}

function readDefaultLubeRate(db) {
  if (!hasTable(db, "cost_settings")) return 4;
  try {
    const row = db.prepare("SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1").get();
    const rate = Number(row?.value);
    return Number.isFinite(rate) && rate > 0 ? rate : 4;
  } catch {
    return 4;
  }
}

function stockMovementDateExpr(db, alias = "sm") {
  if (hasColumn(db, "stock_movements", "created_at")) return `DATE(${alias}.created_at)`;
  if (hasColumn(db, "stock_movements", "movement_date")) return `DATE(${alias}.movement_date)`;
  return "DATE('1970-01-01')";
}

function stockMovementCostExpr(db, movement = "sm", part = "p") {
  const quantity = `ABS(COALESCE(${movement}.quantity, 0))`;
  const choices = [];
  if (hasColumn(db, "stock_movements", "total_cost")) choices.push(`NULLIF(ABS(COALESCE(${movement}.total_cost, 0)), 0)`);
  if (hasColumn(db, "stock_movements", "unit_cost")) choices.push(`NULLIF(${quantity} * COALESCE(${movement}.unit_cost, 0), 0)`);
  if (hasColumn(db, "parts", "unit_cost")) choices.push(`NULLIF(${quantity} * COALESCE(${part}.unit_cost, 0), 0)`);
  if (hasColumn(db, "stock_movements", "unit_cost_usd")) choices.push(`NULLIF(${quantity} * COALESCE(${movement}.unit_cost_usd, 0), 0)`);
  if (hasColumn(db, "stock_movements", "cost_input")) choices.push(`NULLIF(${quantity} * COALESCE(${movement}.cost_input, 0), 0)`);
  return `COALESCE(${choices.join(", ") || "0"}, 0)`;
}

function assetBelongsToSite(asset, siteCode, siteAware) {
  if (!siteAware) return true;
  return normalizedSiteCode(asset?.site_code) === siteCode;
}

function isWorkshopAsset(asset) {
  return String(asset?.cost_center_code || "").trim().toUpperCase() === WORKSHOP_DEPARTMENT;
}

function sourceRow(values = {}) {
  return {
    date: "", source_id: "", source_type: "", asset_code: "UNLINKED", asset_name: "", description: "",
    quantity: 0, unit_rate: 0, source_amount: 0, plant_revenue: 0, plant_labour_charge: 0,
    plant_lube_charge: 0, plant_parts: 0, plant_other_operating: 0, plant_depreciation: 0,
    workshop_labour_cost: 0, workshop_lube_cost: 0, workshop_duties: 0,
    workshop_other_operating: 0, workshop_recovery: 0, ...values,
  };
}

function safePlantHireLines(db, period, siteCode, assetByCode, assetsHaveSite) {
  try {
    return buildPlantHireLines(db, period).filter((line) => {
      const asset = assetByCode.get(String(line.asset_code || "").toUpperCase());
      return assetBelongsToSite(asset || line, siteCode, assetsHaveSite) && !isWorkshopAsset(asset || line);
    });
  } catch {
    // Older databases may not yet include the full plant-hire asset schema.
    return [];
  }
}

/** Builds a traceable Plant/Workshop monthly costing model. */
export function buildMonthlyBudgetSpend(db, { month, siteCode = "main" } = {}) {
  const bounds = monthBounds(month);
  if (!bounds) throw new Error("month (YYYY-MM) required");

  const site_code = normalizedSiteCode(siteCode);
  const defaultLaborRate = readDefaultRate(db);
  const defaultLubeRate = readDefaultLubeRate(db);
  const assetsHaveSite = hasColumn(db, "assets", "site_code");
  const assetsHaveCostCenter = hasColumn(db, "assets", "cost_center_code");
  const assetRows = hasTable(db, "assets")
    ? db.prepare(`
        SELECT id, asset_code, asset_name,
          ${assetsHaveSite ? "COALESCE(NULLIF(TRIM(site_code), ''), 'main')" : "'main'"} AS site_code,
          ${assetsHaveCostCenter ? "COALESCE(NULLIF(TRIM(cost_center_code), ''), '')" : "''"} AS cost_center_code
        FROM assets
      `).all()
    : [];
  const assetById = new Map(assetRows.map((row) => [Number(row.id), row]));
  const assetByCode = new Map(assetRows.map((row) => [String(row.asset_code || "").trim().toUpperCase(), row]));

  let budget = { budget_amount: 0, source: null, resolved_site_code: null };
  try { budget = getOperatingBudgetAmount(db, month, site_code); } catch {}

  const categories = [
    { key: "mechanics", label: "Mechanics cost", source: "Saved mechanic labour entries", amount: 0, records: 0 },
    { key: "lubes", label: "Lubes used", source: "Lube usage logs and direct stores issues", amount: 0, records: 0 },
    { key: "maintenance", label: "Maintenance service work", source: "Completed service work-order labour not linked to a mechanic job card", amount: 0, records: 0 },
    { key: "parts", label: "Parts used", source: "Outbound non-lubricant stores issues", amount: 0, records: 0 },
    { key: "other", label: "Plant other operating expense", source: "Manual departmental entries", amount: 0, records: 0 },
  ];
  const category = new Map(categories.map((row) => [row.key, row]));
  const details = [];
  const addDetail = (row) => details.push({
    date: "", source_id: "", source_type: "", asset_code: "UNLINKED", asset_name: "", reference: "",
    category: "", description: "", quantity: 0, unit_rate: 0, amount: 0, allocation_role: "", ...row,
  });

  if (hasTable(db, "mechanic_labor_entries")) {
    const hasSite = hasColumn(db, "mechanic_labor_entries", "site_code");
    const hasRate = hasColumn(db, "mechanic_labor_entries", "labor_rate_per_hour");
    const hasCategory = hasColumn(db, "mechanic_labor_entries", "category");
    const hasJobCard = hasColumn(db, "mechanic_labor_entries", "job_card_no");
    const rows = db.prepare(`
      SELECT id, work_date, technician_name, asset_code, COALESCE(reason, '') AS reason,
        ${hasCategory ? "COALESCE(category, '')" : "''"} AS work_category,
        ${hasJobCard ? "COALESCE(job_card_no, '')" : "''"} AS job_card_no,
        COALESCE(hours, 0) AS hours,
        ${hasRate ? "COALESCE(NULLIF(labor_rate_per_hour, 0), ?)" : "?"} AS unit_rate
      FROM mechanic_labor_entries
      WHERE work_date BETWEEN ? AND ? ${hasSite ? "AND LOWER(COALESCE(NULLIF(TRIM(site_code), ''), 'main')) = ?" : ""}
      ORDER BY work_date ASC, id ASC
    `).all(defaultLaborRate, bounds.start, bounds.end, ...(hasSite ? [site_code] : []));
    for (const row of rows) {
      const amount = num(Number(row.hours || 0) * Number(row.unit_rate || defaultLaborRate));
      category.get("mechanics").amount += amount;
      category.get("mechanics").records += 1;
      const asset = assetByCode.get(String(row.asset_code || "").trim().toUpperCase());
      addDetail({
        date: String(row.work_date || ""), source_id: `mechanic_labor:${row.id}`, source_type: "Mechanic labour",
        asset_code: String(row.asset_code || "UNLINKED"), asset_name: String(asset?.asset_name || ""),
        reference: row.job_card_no ? `Job card ${row.job_card_no}` : "Mechanic entry",
        description: [row.technician_name, row.work_category, row.reason].filter(Boolean).join(" — ") || "Mechanic labour",
        category: "Mechanics cost", quantity: num(row.hours || 0), unit_rate: num(row.unit_rate || defaultLaborRate), amount, allocation_role: "labour",
      });
    }
  }

  if (hasTable(db, "oil_logs") && hasTable(db, "assets")) {
    const lubeLines = fetchLubeUsageLines(db, { start: bounds.start, end: bounds.end, lubeUnitFallback: defaultLubeRate })
      .filter((line) => assetBelongsToSite(assetById.get(Number(line.asset_id)), site_code, assetsHaveSite));
    for (const line of lubeLines) {
      const amount = num(line.line_cost || 0);
      category.get("lubes").amount += amount;
      category.get("lubes").records += 1;
      addDetail({
        date: String(line.usage_date || ""),
        source_id: `${line.source || "lube"}:${line.id || `${line.asset_id || "asset"}:${line.usage_date}`}`,
        source_type: "Lube issue", asset_code: String(line.asset_code || "UNLINKED"), asset_name: String(line.asset_name || ""),
        reference: line.work_order_id ? `WO #${line.work_order_id}` : String(line.source || "lube usage"),
        description: [line.part_code, line.part_name].filter(Boolean).join(" — ") || String(line.lube_type || "Lube"),
        category: "Lubes used", quantity: num(line.quantity || 0, 3), unit_rate: num(line.unit_cost || 0, 4), amount, allocation_role: "lube",
      });
    }
  }

  if (hasTable(db, "stock_movements") && hasTable(db, "parts")) {
    ensureStockCategorySchema(db);
    const dateExpr = stockMovementDateExpr(db);
    const costExpr = stockMovementCostExpr(db);
    const hasWorkOrders = hasTable(db, "work_orders");
    const assetSiteExpr = assetsHaveSite ? "COALESCE(aw.site_code, aa.site_code, 'main')" : "'main'";
    const joinWorkOrders = hasWorkOrders
      ? "LEFT JOIN work_orders w ON sm.reference = ('work_order:' || w.id) LEFT JOIN assets aw ON aw.id = w.asset_id"
      : "LEFT JOIN assets aw ON 0";
    const partsRows = db.prepare(`
      SELECT sm.id, ${dateExpr} AS issue_date, COALESCE(aw.asset_code, aa.asset_code, 'UNLINKED') AS asset_code,
        COALESCE(aw.asset_name, aa.asset_name, '') AS asset_name, ${assetSiteExpr} AS asset_site_code,
        p.part_code, p.part_name, sm.reference, ABS(COALESCE(sm.quantity, 0)) AS quantity,
        CASE WHEN ABS(COALESCE(sm.quantity, 0)) > 0 THEN (${costExpr}) / ABS(COALESCE(sm.quantity, 0)) ELSE 0 END AS unit_rate,
        (${costExpr}) AS amount
      FROM stock_movements sm JOIN parts p ON p.id = sm.part_id ${joinWorkOrders}
      LEFT JOIN assets aa ON sm.reference LIKE ('asset:' || aa.id || ':stores')
      WHERE (LOWER(COALESCE(sm.movement_type, '')) = 'out' OR COALESCE(sm.quantity, 0) < 0)
        AND ${dateExpr} BETWEEN ? AND ? AND NOT (${oilPartSql("p")})
      ORDER BY ${dateExpr} ASC, sm.id ASC
    `).all(bounds.start, bounds.end)
      .filter((row) => !assetsHaveSite || !row.asset_code || row.asset_code === "UNLINKED" || normalizedSiteCode(row.asset_site_code) === site_code);
    for (const row of partsRows) {
      const amount = num(row.amount || 0);
      category.get("parts").amount += amount;
      category.get("parts").records += 1;
      addDetail({
        date: String(row.issue_date || ""), source_id: `stock_movement:${row.id}`, source_type: "Stores part issue",
        asset_code: String(row.asset_code || "UNLINKED"), asset_name: String(row.asset_name || ""),
        reference: String(row.reference || "Stores issue"), description: [row.part_code, row.part_name].filter(Boolean).join(" — ") || "Stores issue",
        category: "Parts used", quantity: num(row.quantity || 0, 3), unit_rate: num(row.unit_rate || 0, 4), amount, allocation_role: "part",
      });
    }
  }

  if (hasTable(db, "work_orders")) {
    const hasCompleted = hasColumn(db, "work_orders", "completed_at");
    const hasLaborHours = hasColumn(db, "work_orders", "labor_hours");
    const hasLaborRate = hasColumn(db, "work_orders", "labor_rate_per_hour");
    const hasJobDescription = hasColumn(db, "work_orders", "job_description");
    const hasSite = hasColumn(db, "work_orders", "site_code");
    const hasMechanics = hasTable(db, "mechanic_labor_entries");
    const mechanicsHasJobCard = hasMechanics && hasColumn(db, "mechanic_labor_entries", "job_card_no");
    const mechanicsHasSite = hasMechanics && hasColumn(db, "mechanic_labor_entries", "site_code");
    const closedAt = hasCompleted ? "COALESCE(NULLIF(w.completed_at, ''), NULLIF(w.closed_at, ''), w.opened_at)" : "COALESCE(NULLIF(w.closed_at, ''), w.opened_at)";
    const jobCardExclusion = mechanicsHasJobCard
      ? `AND NOT EXISTS (SELECT 1 FROM mechanic_labor_entries ml WHERE TRIM(COALESCE(ml.job_card_no, '')) = CAST(w.id AS TEXT) ${mechanicsHasSite ? "AND LOWER(COALESCE(NULLIF(TRIM(ml.site_code), ''), 'main')) = ?" : ""})`
      : "";
    const rows = db.prepare(`
      SELECT w.id, DATE(${closedAt}) AS closed_date, a.asset_code, a.asset_name,
        COALESCE(${hasJobDescription ? "w.job_description" : "NULL"}, '') AS job_description,
        ${hasLaborHours ? "COALESCE(w.labor_hours, 0)" : "0"} AS hours,
        ${hasLaborRate ? "COALESCE(NULLIF(w.labor_rate_per_hour, 0), ?)" : "?"} AS unit_rate
      FROM work_orders w LEFT JOIN assets a ON a.id = w.asset_id
      WHERE LOWER(COALESCE(w.source, '')) = 'service' AND LOWER(COALESCE(w.status, '')) IN ('completed', 'approved', 'closed')
        AND DATE(${closedAt}) BETWEEN ? AND ? ${hasSite ? "AND LOWER(COALESCE(NULLIF(TRIM(w.site_code), ''), 'main')) = ?" : ""} ${jobCardExclusion}
      ORDER BY DATE(${closedAt}) ASC, w.id ASC
    `).all(defaultLaborRate, bounds.start, bounds.end, ...(hasSite ? [site_code] : []), ...(mechanicsHasJobCard && mechanicsHasSite ? [site_code] : []));
    for (const row of rows) {
      const amount = num(Number(row.hours || 0) * Number(row.unit_rate || defaultLaborRate));
      category.get("maintenance").amount += amount;
      category.get("maintenance").records += 1;
      addDetail({
        date: String(row.closed_date || ""), source_id: `work_order:${row.id}:labour`, source_type: "Service work-order labour",
        asset_code: String(row.asset_code || "UNLINKED"), asset_name: String(row.asset_name || ""), reference: `WO #${row.id}`,
        description: String(row.job_description || "Completed service work order"), quantity: num(row.hours || 0),
        category: "Maintenance service work", unit_rate: num(row.unit_rate || defaultLaborRate), amount, allocation_role: "labour",
      });
    }
  }

  const manualInputs = listDepartmentalCostingInputs(db, { period: month, siteCode: site_code });
  for (const input of manualInputs) {
    const inputType = input.input_type;
    const amount = num(input.amount || 0);
    const asset = assetByCode.get(String(input.asset_code || "").trim().toUpperCase());
    const allocationRole = String(input.entry_type || "").trim();
    if (allocationRole === "plant_other_operating") {
      category.get("other").amount += amount;
      category.get("other").records += 1;
    }
    addDetail({
      // The entry is assigned to the selected accounting month; created_at stays
      // in the input register as the audit timestamp rather than changing report period.
      date: bounds.end, source_id: `departmental_input:${input.id}`,
      source_type: inputType?.label || "Manual departmental entry",
      asset_code: String(input.asset_code || "UNALLOCATED").trim().toUpperCase() || "UNALLOCATED", asset_name: String(asset?.asset_name || ""),
      reference: String(input.reference || `Manual entry #${input.id}`), description: String(input.description || inputType?.label || "Manual departmental entry"),
      category: String(inputType?.label || "Manual departmental entry"), quantity: 1, unit_rate: amount, amount, allocation_role: allocationRole,
    });
  }

  const source_lines = [];
  const addSource = (row) => source_lines.push(sourceRow({
    ...row, asset_code: String(row.asset_code || "UNLINKED").trim().toUpperCase() || "UNLINKED", source_amount: num(row.source_amount || 0),
  }));
  for (const detail of details) {
    const assetCode = String(detail.asset_code || "UNLINKED").trim().toUpperCase() || "UNLINKED";
    const asset = assetByCode.get(assetCode);
    const amount = num(detail.amount || 0);
    const plantAsset = Boolean(asset) && !isWorkshopAsset(asset);
    const base = { date: detail.date, source_id: detail.source_id, source_type: detail.source_type, asset_code: assetCode,
      asset_name: detail.asset_name || asset?.asset_name || "", description: detail.description,
      quantity: detail.quantity, unit_rate: detail.unit_rate, source_amount: amount };
    switch (detail.allocation_role) {
      case "labour": addSource({ ...base, workshop_labour_cost: amount, plant_labour_charge: plantAsset ? amount : 0, workshop_recovery: plantAsset ? amount : 0 }); break;
      case "lube": addSource({ ...base, workshop_lube_cost: amount, plant_lube_charge: plantAsset ? amount : 0, workshop_recovery: plantAsset ? amount : 0 }); break;
      case "part": addSource({ ...base, plant_parts: !asset || plantAsset ? amount : 0, workshop_other_operating: asset && isWorkshopAsset(asset) ? amount : 0 }); break;
      case "plant_revenue": addSource({ ...base, plant_revenue: amount }); break;
      case "plant_other_operating": addSource({ ...base, plant_other_operating: amount }); break;
      case "plant_depreciation": addSource({ ...base, plant_depreciation: amount }); break;
      case "workshop_duties": addSource({ ...base, workshop_duties: amount }); break;
      case "workshop_other_operating": addSource({ ...base, workshop_other_operating: amount }); break;
      default: addSource(base); break;
    }
  }

  for (const line of safePlantHireLines(db, month, site_code, assetByCode, assetsHaveSite)) {
    addSource({
      date: bounds.end, source_id: `plant_hire:${String(line.asset_code || "").toUpperCase()}:${month}`,
      source_type: "External plant hire billing", asset_code: String(line.asset_code || "UNALLOCATED"), asset_name: String(line.asset_name || ""),
      description: `${line.billing_mode === "fixed_monthly" ? "Monthly" : "Hourly"} plant hire billing`,
      quantity: line.billing_mode === "fixed_monthly" ? 1 : num(line.hours_run || line.km_run || 0, 3), unit_rate: num(line.rate || 0, 4),
      source_amount: num(line.amount || 0), plant_revenue: num(line.amount || 0),
    });
  }
  source_lines.sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.source_id).localeCompare(String(b.source_id)));

  const emptyPlantTotals = () => ({ revenue: 0, labour: 0, lube: 0, parts: 0, other_operating: 0, depreciation: 0 });
  const addTotals = (target, row) => {
    target.revenue += Number(row.plant_revenue || 0); target.labour += Number(row.plant_labour_charge || 0);
    target.lube += Number(row.plant_lube_charge || 0); target.parts += Number(row.plant_parts || 0);
    target.other_operating += Number(row.plant_other_operating || 0); target.depreciation += Number(row.plant_depreciation || 0);
    return target;
  };
  const plantTotals = source_lines.reduce(addTotals, emptyPlantTotals());
  plantTotals.operating_cost = num(plantTotals.labour + plantTotals.lube + plantTotals.parts + plantTotals.other_operating);
  plantTotals.total_cost = num(plantTotals.operating_cost + plantTotals.depreciation);
  plantTotals.net_contribution = num(plantTotals.revenue - plantTotals.total_cost);
  plantTotals.profit_margin_pct = plantTotals.revenue > 0 ? num((plantTotals.net_contribution / plantTotals.revenue) * 100, 1) : null;

  const workshop = source_lines.reduce((totals, row) => {
    totals.labour_cost += Number(row.workshop_labour_cost || 0); totals.lube_cost += Number(row.workshop_lube_cost || 0);
    totals.duties += Number(row.workshop_duties || 0); totals.other_operating += Number(row.workshop_other_operating || 0);
    totals.recovery += Number(row.workshop_recovery || 0); return totals;
  }, { labour_cost: 0, lube_cost: 0, duties: 0, other_operating: 0, recovery: 0 });
  workshop.total_operating_cost = num(workshop.labour_cost + workshop.lube_cost + workshop.duties + workshop.other_operating);
  workshop.net_cost_recovery_balance = num(workshop.recovery - workshop.total_operating_cost);

  const plantAssetMap = new Map();
  for (const row of source_lines) {
    const hasPlantValue = [row.plant_revenue, row.plant_labour_charge, row.plant_lube_charge, row.plant_parts, row.plant_other_operating, row.plant_depreciation]
      .some((value) => Math.abs(Number(value || 0)) > 0.0001);
    if (!hasPlantValue) continue;
    const code = String(row.asset_code || "UNALLOCATED").trim().toUpperCase() || "UNALLOCATED";
    const existing = plantAssetMap.get(code) || { asset_code: code, asset_name: row.asset_name || assetByCode.get(code)?.asset_name || "", ...emptyPlantTotals() };
    addTotals(existing, row); plantAssetMap.set(code, existing);
  }
  const plant_assets = Array.from(plantAssetMap.values()).map((row) => {
    const operating_cost = num(row.labour + row.lube + row.parts + row.other_operating);
    const total_cost = num(operating_cost + row.depreciation);
    const net_contribution = num(row.revenue - total_cost);
    return {
      ...Object.fromEntries(Object.entries(row).map(([key, value]) => [key, typeof value === "number" ? num(value) : value])),
      operating_cost, total_cost, net_contribution, profit_margin_pct: row.revenue > 0 ? num((net_contribution / row.revenue) * 100, 1) : null,
    };
  }).sort((a, b) => String(a.asset_code).localeCompare(String(b.asset_code)));

  const reconciliation = source_lines
    .filter((row) => Math.abs(Number(row.workshop_recovery || 0)) > 0.0001 || Math.abs(Number(row.plant_labour_charge || 0) + Number(row.plant_lube_charge || 0)) > 0.0001)
    .map((row) => {
      const plant_expense = num(Number(row.plant_labour_charge || 0) + Number(row.plant_lube_charge || 0));
      const recovery = num(row.workshop_recovery || 0); const difference = num(recovery - plant_expense);
      const assetCode = String(row.asset_code || "UNLINKED").trim().toUpperCase() || "UNLINKED";
      return { source_id: row.source_id, source_type: row.source_type, asset_code: assetCode, description: row.description,
        workshop_recovery: recovery, plant_expense, difference, status: assetCode === "UNLINKED" ? "Missing plant asset" : Math.abs(difference) < 0.01 ? "Matched" : "Unmatched" };
    });

  const total_spend = num(plantTotals.operating_cost);
  const budget_amount = num(budget?.budget_amount || 0); const budget_is_set = budget_amount > 0;
  const variance = budget_is_set ? num(total_spend - budget_amount) : null;
  const budget_used_pct = budget_is_set ? num((total_spend / budget_amount) * 100, 1) : null;

  return {
    month: String(month), start: bounds.start, end: bounds.end, site_code, currency: "USD",
    budget: { amount: budget_amount, is_set: budget_is_set, source: budget?.source || null, resolved_site_code: budget?.resolved_site_code || null },
    total_spend, variance, budget_used_pct,
    categories: categories.map((row) => ({ ...row, amount: num(row.amount) })),
    details: details.map((row) => ({ ...row, amount: num(row.amount), quantity: num(row.quantity, 3), unit_rate: num(row.unit_rate, 4) }))
      .sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.source_id).localeCompare(String(b.source_id))),
    departmental: {
      plant_department: PLANT_DEPARTMENT, workshop_department: WORKSHOP_DEPARTMENT, plant: plantTotals,
      workshop: Object.fromEntries(Object.entries(workshop).map(([key, value]) => [key, num(value)])),
      consolidated_result: num(plantTotals.net_contribution + workshop.net_cost_recovery_balance),
      plant_assets, reconciliation, manual_inputs: manualInputs, source_lines,
    },
  };
}
