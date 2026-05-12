/**
 * Lube-store month opening/closing balances from stock_movements (cumulative qty).
 * Opening = on-hand at end of previous calendar month; closing = end of selected month.
 */

function stockMovementDateExpr(db) {
  const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
  const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
  return hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
}

const LUBE_PART_FILTER = `(
  LOWER(IFNULL(p.part_code, '')) LIKE '%oil%' OR
  LOWER(IFNULL(p.part_name, '')) LIKE '%oil%' OR
  LOWER(IFNULL(p.part_code, '')) LIKE '%lube%' OR
  LOWER(IFNULL(p.part_name, '')) LIKE '%lube%' OR
  LOWER(IFNULL(p.part_code, '')) LIKE '%grease%' OR
  LOWER(IFNULL(p.part_name, '')) LIKE '%grease%'
)`;

/** @param {*} db SQLite database (better-sqlite3) */
export function fetchLubeMonthStockSnapshot(db, opts) {
  const month = String(opts?.month || "").trim();
  if (!/^\d{4}-\d{2}$/.test(month)) {
    throw new Error("month must be YYYY-MM");
  }
  const monthStart = `${month}-01`;
  const openingRow = db.prepare(`SELECT date(?, '-1 day') AS d`).get(monthStart);
  const closingRow = db.prepare(`SELECT date(?, '+1 month', '-1 day') AS d`).get(monthStart);
  const opening_as_of = String(openingRow?.d || "");
  const closing_as_of = String(closingRow?.d || "");
  const smDateExpr = stockMovementDateExpr(db);
  const location_id = opts.location_id != null ? Number(opts.location_id) : null;
  const hasLoc = Number.isFinite(location_id) && location_id > 0;

  const locJoinSql = hasLoc
    ? `AND (sm.location_id = ? OR sm.location_id IS NULL)`
    : "";

  const sql = `
    SELECT
      p.part_code,
      p.part_name,
      p.min_stock,
      COALESCE(SUM(CASE WHEN ${smDateExpr} <= ? THEN sm.quantity ELSE 0 END), 0) AS opening_qty,
      COALESCE(SUM(sm.quantity), 0) AS closing_qty
    FROM parts p
    LEFT JOIN stock_movements sm ON sm.part_id = p.id
      AND ${smDateExpr} <= ?
      ${locJoinSql}
    WHERE ${LUBE_PART_FILTER}
    GROUP BY p.id
    ORDER BY p.part_code ASC
    LIMIT 800
  `;

  const params = hasLoc
    ? [closing_as_of, location_id, opening_as_of]
    : [closing_as_of, opening_as_of];

  const rows = db.prepare(sql).all(...params).map((r) => {
    const opening_qty = Number(r.opening_qty || 0);
    const closing_qty = Number(r.closing_qty || 0);
    return {
      part_code: r.part_code,
      part_name: r.part_name,
      min_stock: Number(r.min_stock || 0),
      opening_qty,
      closing_qty,
      net_month_movement: Number((closing_qty - opening_qty).toFixed(4)),
    };
  });

  return {
    month,
    month_start: monthStart,
    month_end: closing_as_of,
    opening_as_of,
    closing_as_of,
    rows,
  };
}
