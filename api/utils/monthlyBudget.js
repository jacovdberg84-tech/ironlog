/** Monthly budget rows in finance_budgets_monthly (single total or plant hire income). */

export function resolveSiteCandidates(siteHint) {
  const s = String(siteHint || "").trim().toLowerCase();
  const out = [];
  if (s) out.push(s);
  for (const x of ["main", "", "default"]) {
    if (!out.includes(x)) out.push(x);
  }
  return out;
}

export function getMonthlyBudgetRow(db, period, category, siteHint) {
  const sites = resolveSiteCandidates(siteHint);
  for (const site of sites) {
    const row = db.prepare(`
      SELECT budget_amount, currency, notes, updated_at, site_code
      FROM finance_budgets_monthly
      WHERE period = ?
        AND category = ?
        AND COALESCE(site_code, '') = ?
        AND COALESCE(cost_center_code, '') = ''
        AND COALESCE(equipment_type, '') = ''
      LIMIT 1
    `).get(period, category, site);
    if (row && Number(row.budget_amount || 0) > 0) {
      return {
        budget_amount: Number(row.budget_amount),
        currency: row.currency || "USD",
        notes: row.notes || null,
        updated_at: row.updated_at || null,
        site_code: row.site_code || site || null,
      };
    }
  }
  const row = db.prepare(`
    SELECT budget_amount, currency, notes, updated_at, site_code
    FROM finance_budgets_monthly
    WHERE period = ?
      AND category = ?
      AND COALESCE(cost_center_code, '') = ''
      AND COALESCE(equipment_type, '') = ''
    ORDER BY updated_at DESC
    LIMIT 1
  `).get(period, category);
  if (!row) {
    return { budget_amount: 0, currency: "USD", notes: null, updated_at: null, site_code: null };
  }
  return {
    budget_amount: Number(row.budget_amount || 0),
    currency: row.currency || "USD",
    notes: row.notes || null,
    updated_at: row.updated_at || null,
    site_code: row.site_code || null,
  };
}

export function upsertMonthlyBudget(db, { period, site_code = "", category, budget_amount, notes = null, created_by = null }) {
  db.prepare(`
    INSERT INTO finance_budgets_monthly
      (period, site_code, cost_center_code, equipment_type, category, budget_amount, currency, notes, created_by)
    VALUES (?, ?, '', '', ?, ?, 'USD', ?, ?)
    ON CONFLICT(period, site_code, cost_center_code, equipment_type, category) DO UPDATE SET
      budget_amount = excluded.budget_amount,
      notes = excluded.notes,
      updated_at = datetime('now')
  `).run(period, site_code, category, budget_amount, notes, created_by);
}

const EXPENSE_CATEGORIES = ["parts", "labor", "fuel", "lube", "downtime"];

/** Single monthly operating expense budget (category `operating`), with legacy per-category sum fallback. */
export function getOperatingBudgetAmount(db, period, siteHint) {
  const direct = getMonthlyBudgetRow(db, period, "operating", siteHint);
  if (Number(direct.budget_amount || 0) > 0) {
    return { budget_amount: Number(direct.budget_amount), source: "operating", resolved_site_code: direct.site_code };
  }
  const sites = resolveSiteCandidates(siteHint);
  for (const site of sites) {
    const rows = db.prepare(`
      SELECT category, budget_amount
      FROM finance_budgets_monthly
      WHERE period = ?
        AND COALESCE(site_code, '') = ?
        AND COALESCE(cost_center_code, '') = ''
        AND COALESCE(equipment_type, '') = ''
        AND category IN (${EXPENSE_CATEGORIES.map(() => "?").join(",")})
    `).all(period, site, ...EXPENSE_CATEGORIES);
    const sum = rows.reduce((s, r) => s + Number(r.budget_amount || 0), 0);
    if (sum > 0) {
      return { budget_amount: Number(sum.toFixed(2)), source: "legacy_categories", resolved_site_code: site || null };
    }
  }
  return { budget_amount: 0, source: null, resolved_site_code: null };
}
