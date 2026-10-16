// IRONLOG/api/utils/monthlyBudgetSpend.js
// Financially transparent monthly budget-versus-spend source data.  This is
// deliberately separate from the broader fleet-cost report: it uses the
// labour, lube and stores activity that management reviews each month.
import { getOperatingBudgetAmount } from "./monthlyBudget.js";
import { ensureStockCategorySchema, oilPartSql } from "./stockCategory.js";
import { fetchLubeUsageLines } from "./lubeUsageLines.js";

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
  if (!Number.isInteger(year) || !Number.isInteger(monthNumber) || monthNumber < 1 || monthNumber > 12) {
    return null;
  }
  const start = `${year}-${String(monthNumber).padStart(2, "0")}-01`;
  const end = new Date(Date.UTC(year, monthNumber, 0)).toISOString().slice(0, 10);
  return { start, end };
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
  return String(asset?.site_code || "main").trim().toLowerCase() === siteCode;
}

/**
 * Build a management-ready monthly report source.  Categories are deliberately
 * mutually understandable: service work-order labour is excluded when a
 * mechanic entry carries the matching job-card number, preventing the most
 * common double capture between the two workflows.
 */
export function buildMonthlyBudgetSpend(db, { month, siteCode = "main" } = {}) {
  const bounds = monthBounds(month);
  if (!bounds) throw new Error("month (YYYY-MM) required");

  const site_code = String(siteCode || "main").trim().toLowerCase() || "main";
  const defaultLaborRate = readDefaultRate(db);
  const defaultLubeRate = readDefaultLubeRate(db);
  const assetsHaveSite = hasColumn(db, "assets", "site_code");
  const assetRows = hasTable(db, "assets")
    ? db.prepare(`SELECT id, asset_code, asset_name, ${assetsHaveSite ? "COALESCE(NULLIF(TRIM(site_code), ''), 'main')" : "'main'"} AS site_code FROM assets`).all()
    : [];
  const assetById = new Map(assetRows.map((row) => [Number(row.id), row]));

  let budget = { budget_amount: 0, source: null, resolved_site_code: null };
  try {
    budget = getOperatingBudgetAmount(db, month, site_code);
  } catch {
    // The finance module creates this table during startup. A report remains
    // useful on a legacy database even if its first budget has not been saved.
  }

  const categories = [
    {
      key: "mechanics",
      label: "Mechanics cost",
      source: "Saved mechanic labour entries",
      amount: 0,
      records: 0,
    },
    {
      key: "lubes",
      label: "Lubes used",
      source: "Lube usage logs and direct stores issues",
      amount: 0,
      records: 0,
    },
    {
      key: "maintenance",
      label: "Maintenance service work",
      source: "Completed service work-order labour not linked to a mechanic job card",
      amount: 0,
      records: 0,
    },
    {
      key: "parts",
      label: "Parts used",
      source: "Outbound non-lubricant stores issues",
      amount: 0,
      records: 0,
    },
  ];
  const category = new Map(categories.map((row) => [row.key, row]));
  const details = [];

  if (hasTable(db, "mechanic_labor_entries")) {
    const hasSite = hasColumn(db, "mechanic_labor_entries", "site_code");
    const hasRate = hasColumn(db, "mechanic_labor_entries", "labor_rate_per_hour");
    const hasCategory = hasColumn(db, "mechanic_labor_entries", "category");
    const hasJobCard = hasColumn(db, "mechanic_labor_entries", "job_card_no");
    const rows = db.prepare(`
      SELECT
        work_date,
        technician_name,
        asset_code,
        COALESCE(reason, '') AS reason,
        ${hasCategory ? "COALESCE(category, '')" : "''"} AS work_category,
        ${hasJobCard ? "COALESCE(job_card_no, '')" : "''"} AS job_card_no,
        COALESCE(hours, 0) AS hours,
        ${hasRate ? "COALESCE(NULLIF(labor_rate_per_hour, 0), ?)" : "?"} AS unit_rate
      FROM mechanic_labor_entries
      WHERE work_date BETWEEN ? AND ?
        ${hasSite ? "AND LOWER(COALESCE(NULLIF(TRIM(site_code), ''), 'main')) = ?" : ""}
      ORDER BY work_date ASC, id ASC
    `).all(defaultLaborRate, bounds.start, bounds.end, ...(hasSite ? [site_code] : []));

    for (const row of rows) {
      const amount = num(Number(row.hours || 0) * Number(row.unit_rate || defaultLaborRate));
      category.get("mechanics").amount += amount;
      category.get("mechanics").records += 1;
      details.push({
        date: String(row.work_date || ""),
        category: "Mechanics cost",
        asset_code: String(row.asset_code || "UNLINKED"),
        reference: row.job_card_no ? `Job card ${row.job_card_no}` : "Mechanic entry",
        description: [row.work_category, row.reason].filter(Boolean).join(" — ") || "Mechanic labour",
        quantity: num(row.hours || 0),
        unit_rate: num(row.unit_rate || defaultLaborRate),
        amount,
      });
    }
  }

  if (hasTable(db, "oil_logs") && hasTable(db, "assets")) {
    const lubeLines = fetchLubeUsageLines(db, {
      start: bounds.start,
      end: bounds.end,
      lubeUnitFallback: defaultLubeRate,
    }).filter((line) => assetBelongsToSite(assetById.get(Number(line.asset_id)), site_code, assetsHaveSite));

    for (const line of lubeLines) {
      const amount = num(line.line_cost || 0);
      category.get("lubes").amount += amount;
      category.get("lubes").records += 1;
      details.push({
        date: String(line.usage_date || ""),
        category: "Lubes used",
        asset_code: String(line.asset_code || "UNLINKED"),
        reference: line.work_order_id ? `WO #${line.work_order_id}` : String(line.source || "lube usage"),
        description: [line.part_code, line.part_name].filter(Boolean).join(" — ") || String(line.lube_type || "Lube"),
        quantity: num(line.quantity || 0, 3),
        unit_rate: num(line.unit_cost || 0, 4),
        amount,
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
      SELECT
        ${dateExpr} AS issue_date,
        COALESCE(aw.asset_code, aa.asset_code, 'UNLINKED') AS asset_code,
        ${assetSiteExpr} AS asset_site_code,
        p.part_code,
        p.part_name,
        COALESCE(p.stock_category, 'part') AS stock_category,
        sm.reference,
        ABS(COALESCE(sm.quantity, 0)) AS quantity,
        CASE
          WHEN ABS(COALESCE(sm.quantity, 0)) > 0 THEN (${costExpr}) / ABS(COALESCE(sm.quantity, 0))
          ELSE 0
        END AS unit_rate,
        (${costExpr}) AS amount
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      ${joinWorkOrders}
      LEFT JOIN assets aa ON sm.reference LIKE ('asset:' || aa.id || ':stores')
      WHERE (LOWER(COALESCE(sm.movement_type, '')) = 'out' OR COALESCE(sm.quantity, 0) < 0)
        AND ${dateExpr} BETWEEN ? AND ?
        AND NOT (${oilPartSql("p")})
      ORDER BY ${dateExpr} ASC, sm.id ASC
    `).all(bounds.start, bounds.end)
      .filter((row) => !assetsHaveSite || !row.asset_code || row.asset_code === "UNLINKED" || String(row.asset_site_code || "main").trim().toLowerCase() === site_code);

    for (const row of partsRows) {
      const amount = num(row.amount || 0);
      category.get("parts").amount += amount;
      category.get("parts").records += 1;
      details.push({
        date: String(row.issue_date || ""),
        category: "Parts used",
        asset_code: String(row.asset_code || "UNLINKED"),
        reference: String(row.reference || "Stores issue"),
        description: [row.part_code, row.part_name].filter(Boolean).join(" — ") || "Stores issue",
        quantity: num(row.quantity || 0, 3),
        unit_rate: num(row.unit_rate || 0, 4),
        amount,
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
      ? `AND NOT EXISTS (
          SELECT 1
          FROM mechanic_labor_entries ml
          WHERE TRIM(COALESCE(ml.job_card_no, '')) = CAST(w.id AS TEXT)
            ${mechanicsHasSite ? "AND LOWER(COALESCE(NULLIF(TRIM(ml.site_code), ''), 'main')) = ?" : ""}
        )`
      : "";
    const rows = db.prepare(`
      SELECT
        w.id,
        DATE(${closedAt}) AS closed_date,
        a.asset_code,
        COALESCE(${hasJobDescription ? "w.job_description" : "NULL"}, '') AS job_description,
        COALESCE(w.source, '') AS source,
        ${hasLaborHours ? "COALESCE(w.labor_hours, 0)" : "0"} AS hours,
        ${hasLaborRate ? "COALESCE(NULLIF(w.labor_rate_per_hour, 0), ?)" : "?"} AS unit_rate
      FROM work_orders w
      LEFT JOIN assets a ON a.id = w.asset_id
      WHERE LOWER(COALESCE(w.source, '')) = 'service'
        AND LOWER(COALESCE(w.status, '')) IN ('completed', 'approved', 'closed')
        AND DATE(${closedAt}) BETWEEN ? AND ?
        ${hasSite ? "AND LOWER(COALESCE(NULLIF(TRIM(w.site_code), ''), 'main')) = ?" : ""}
        ${jobCardExclusion}
      ORDER BY DATE(${closedAt}) ASC, w.id ASC
    `).all(
      defaultLaborRate,
      bounds.start,
      bounds.end,
      ...(hasSite ? [site_code] : []),
      ...(mechanicsHasJobCard && mechanicsHasSite ? [site_code] : []),
    );

    for (const row of rows) {
      const amount = num(Number(row.hours || 0) * Number(row.unit_rate || defaultLaborRate));
      category.get("maintenance").amount += amount;
      category.get("maintenance").records += 1;
      details.push({
        date: String(row.closed_date || ""),
        category: "Maintenance service work",
        asset_code: String(row.asset_code || "UNLINKED"),
        reference: `WO #${row.id}`,
        description: String(row.job_description || "Completed service work order"),
        quantity: num(row.hours || 0),
        unit_rate: num(row.unit_rate || defaultLaborRate),
        amount,
      });
    }
  }

  const spend = categories.map((row) => ({ ...row, amount: num(row.amount) }));
  const total_spend = num(spend.reduce((sum, row) => sum + Number(row.amount || 0), 0));
  const budget_amount = num(budget?.budget_amount || 0);
  const budget_is_set = budget_amount > 0;
  const variance = budget_is_set ? num(total_spend - budget_amount) : null;
  const budget_used_pct = budget_is_set ? num((total_spend / budget_amount) * 100, 1) : null;

  return {
    month: String(month),
    start: bounds.start,
    end: bounds.end,
    site_code,
    currency: "USD",
    budget: {
      amount: budget_amount,
      is_set: budget_is_set,
      source: budget?.source || null,
      resolved_site_code: budget?.resolved_site_code || null,
    },
    total_spend,
    variance,
    budget_used_pct,
    categories: spend,
    details: details
      .map((row) => ({ ...row, amount: num(row.amount), quantity: num(row.quantity, 3), unit_rate: num(row.unit_rate, 4) }))
      .sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.category).localeCompare(String(b.category))),
  };
}
