// IRONLOG/api/routes/stock/inventory.routes.js — Stock on hand, GM stock report, stock monitor, control summary, movement report and FX settings.
// Registered by routes/stock.routes.js; shared helpers arrive through ctx.
import { buildWorkshopInventoryReportWorkbook, resolveWorkshopInventoryReportPeriod } from "../../utils/workshopInventoryReport.js";
import { db } from "../../db/client.js";
import { oilPartSql, stockCategoryLabel, stockCategorySql } from "../../utils/stockCategory.js";

export default function registerInventoryRoutes(app, ctx) {
  const { getFxRate, getOnHand, getPartByCode, getSiteCode, hasColumn, hasTable, requireRoles } = ctx;

  // Stock on hand summary
  app.get("/onhand", async () => {
    const govCols = [];
    if (hasColumn("parts", "department_code")) govCols.push("p.department_code");
    if (hasColumn("parts", "default_supplier_code")) govCols.push("p.default_supplier_code");
    if (hasColumn("parts", "data_owner_username")) govCols.push("p.data_owner_username");
    const govSql = govCols.length ? `, ${govCols.join(", ")}` : "";
    const rows = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        p.critical,
        p.min_stock,
        p.unit_cost,
        ${stockCategorySql("p")} AS stock_category,
        COALESCE(p.stock_category_source, 'auto') AS stock_category_source
        ${govSql},
        IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      GROUP BY p.id
      ORDER BY p.critical DESC, p.part_code ASC
    `).all();

    return rows.map(r => ({
      ...r,
      critical: Boolean(r.critical),
      stock_category_label: stockCategoryLabel(r.stock_category),
      on_hand: Number(r.on_hand),
      unit_cost: Number(r.unit_cost || 0),
      stock_value: Number((Number(r.on_hand || 0) * Number(r.unit_cost || 0)).toFixed(2)),
      below_min: Number(r.on_hand) < Number(r.min_stock)
    }));
  });

  // GM-ready weekly/monthly inventory register. It uses the same stock movement
  // ledger as Stores, with period opening and closing balances for every item.
  app.get("/gm-stock-report.xlsx", async (req, reply) => {
    try {
      const period = resolveWorkshopInventoryReportPeriod(req.query?.period, req.query?.report_date);
      const dateColumn = hasColumn("stock_movements", "created_at") ? "sm.created_at" : "sm.movement_date";
      const dateExpr = `DATE(${dateColumn})`;
      const hasMovementCost = hasColumn("stock_movements", "unit_cost_usd");
      const hasLocation = hasColumn("stock_movements", "location_id");
      const hasBin = hasColumn("stock_movements", "bin_id");
      const hasSupplierMaster = hasTable("mdm_suppliers");
      const supplierJoin = hasSupplierMaster
        ? "LEFT JOIN mdm_suppliers sup ON sup.site_code = ? AND sup.supplier_code = p.default_supplier_code"
        : "";
      const itemRows = db.prepare(`
        WITH movement_summary AS (
          SELECT
            sm.part_id,
            COALESCE(SUM(CASE WHEN ${dateExpr} < DATE(?) THEN sm.quantity ELSE 0 END), 0) AS opening_qty,
            COALESCE(SUM(CASE
              WHEN ${dateExpr} BETWEEN DATE(?) AND DATE(?)
                AND sm.quantity > 0
                AND LOWER(COALESCE(sm.movement_type, '')) NOT LIKE '%return%'
                AND LOWER(COALESCE(sm.movement_type, '')) NOT LIKE '%adjust%'
                AND LOWER(COALESCE(sm.movement_type, '')) NOT LIKE '%transfer%'
              THEN sm.quantity ELSE 0 END), 0) AS receipts_qty,
            COALESCE(SUM(CASE
              WHEN ${dateExpr} BETWEEN DATE(?) AND DATE(?)
                AND sm.quantity < 0
                AND LOWER(COALESCE(sm.movement_type, '')) NOT LIKE '%adjust%'
                AND LOWER(COALESCE(sm.movement_type, '')) NOT LIKE '%transfer%'
              THEN ABS(sm.quantity) ELSE 0 END), 0) AS issues_qty,
            COALESCE(SUM(CASE
              WHEN ${dateExpr} BETWEEN DATE(?) AND DATE(?)
                AND sm.quantity > 0
                AND LOWER(COALESCE(sm.movement_type, '')) LIKE '%return%'
              THEN sm.quantity ELSE 0 END), 0) AS returns_qty,
            COALESCE(SUM(CASE
              WHEN ${dateExpr} BETWEEN DATE(?) AND DATE(?)
                AND sm.quantity > 0
                AND LOWER(COALESCE(sm.movement_type, '')) LIKE '%transfer%'
              THEN sm.quantity ELSE 0 END), 0) AS transfers_in_qty,
            COALESCE(SUM(CASE
              WHEN ${dateExpr} BETWEEN DATE(?) AND DATE(?)
                AND sm.quantity < 0
                AND LOWER(COALESCE(sm.movement_type, '')) LIKE '%transfer%'
              THEN ABS(sm.quantity) ELSE 0 END), 0) AS transfers_out_qty,
            COALESCE(SUM(CASE
              WHEN ${dateExpr} BETWEEN DATE(?) AND DATE(?)
                AND LOWER(COALESCE(sm.movement_type, '')) LIKE '%adjust%'
              THEN sm.quantity ELSE 0 END), 0) AS adjustments_qty,
            COALESCE(SUM(CASE WHEN ${dateExpr} <= DATE(?) THEN sm.quantity ELSE 0 END), 0) AS closing_qty,
            MAX(CASE WHEN ${dateExpr} <= DATE(?) THEN ${dateColumn} END) AS last_movement_at,
            GROUP_CONCAT(DISTINCT ${hasLocation ? "COALESCE(loc.location_code, 'UNSPECIFIED')" : "'UNSPECIFIED'"}) AS location
          FROM stock_movements sm
          ${hasLocation ? "LEFT JOIN stock_locations loc ON loc.id = sm.location_id" : ""}
          GROUP BY sm.part_id
        ), minmax AS (
          SELECT
            part_id,
            MAX(min_qty) AS min_qty,
            MAX(max_qty) AS max_qty,
            MAX(COALESCE(reorder_qty, 0)) AS reorder_point
          FROM stock_min_max
          GROUP BY part_id
        )
        SELECT
          p.part_code,
          p.part_name,
          p.critical,
          ${stockCategorySql("p")} AS stock_category,
          COALESCE(NULLIF(TRIM(${hasSupplierMaster ? "sup.name" : "p.default_supplier_code"}), ''), '') AS supplier,
          COALESCE(mm.min_qty, p.min_stock, 0) AS min_qty,
          COALESCE(mm.max_qty, 0) AS max_qty,
          COALESCE(NULLIF(mm.reorder_point, 0), COALESCE(mm.min_qty, p.min_stock, 0)) AS reorder_point,
          COALESCE(ms.opening_qty, 0) AS opening_qty,
          COALESCE(ms.receipts_qty, 0) AS receipts_qty,
          COALESCE(ms.issues_qty, 0) AS issues_qty,
          COALESCE(ms.returns_qty, 0) AS returns_qty,
          COALESCE(ms.transfers_in_qty, 0) AS transfers_in_qty,
          COALESCE(ms.transfers_out_qty, 0) AS transfers_out_qty,
          COALESCE(ms.adjustments_qty, 0) AS adjustments_qty,
          COALESCE(ms.closing_qty, 0) AS closing_qty,
          COALESCE(p.unit_cost, 0) AS unit_cost,
          COALESCE(ms.last_movement_at, '') AS last_movement_at,
          COALESCE(ms.location, 'Unspecified') AS location,
          CASE WHEN ${oilPartSql("p")} THEN 1 ELSE 0 END AS is_lube
        FROM parts p
        LEFT JOIN movement_summary ms ON ms.part_id = p.id
        LEFT JOIN minmax mm ON mm.part_id = p.id
        ${supplierJoin}
        ORDER BY p.critical DESC, p.part_code ASC
      `).all(
        period.start_date, // opening balance
        period.start_date, period.end_date, // receipts
        period.start_date, period.end_date, // issues
        period.start_date, period.end_date, // returns
        period.start_date, period.end_date, // transfers in
        period.start_date, period.end_date, // transfers out
        period.start_date, period.end_date, // adjustments
        period.end_date, // closing balance
        period.end_date, // last movement date
        ...(hasSupplierMaster ? [getSiteCode(req)] : []),
      ).map((row) => ({
        ...row,
        critical: Boolean(row.critical),
        is_lube: Boolean(row.is_lube),
      }));

      const movementRows = db.prepare(`
        SELECT
          ${dateColumn} AS movement_at,
          sm.movement_type,
          sm.quantity,
          sm.reference,
          p.part_code,
          p.part_name,
          ${hasLocation ? "COALESCE(loc.location_code, 'Unspecified')" : "'Unspecified'"} AS location_code,
          ${hasBin ? "COALESCE(bin.bin_code, '')" : "''"} AS bin_code,
          COALESCE(NULLIF(${hasMovementCost ? "sm.unit_cost_usd" : "0"}, 0), p.unit_cost, 0) AS unit_cost
        FROM stock_movements sm
        JOIN parts p ON p.id = sm.part_id
        ${hasLocation ? "LEFT JOIN stock_locations loc ON loc.id = sm.location_id" : ""}
        ${hasBin ? "LEFT JOIN stock_bins bin ON bin.id = sm.bin_id" : ""}
        WHERE ${dateExpr} BETWEEN DATE(?) AND DATE(?)
        ORDER BY ${dateColumn} DESC, sm.id DESC
        LIMIT 10000
      `).all(period.start_date, period.end_date).map((row) => ({
        ...row,
        transaction_type: String(row.movement_type || "").replace(/[_-]+/g, " ").replace(/\b\w/g, (letter) => letter.toUpperCase()) || "Movement",
      }));

      const buffer = await buildWorkshopInventoryReportWorkbook(
        { items: itemRows, movements: movementRows },
        { reportType: period.report_type, startDate: period.start_date, endDate: period.end_date },
      );
      const label = period.report_type === "weekly" ? "Weekly" : "Monthly";
      return reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Content-Disposition", `attachment; filename="IRONLOG_${label}_Stock_Report_${period.start_date}_to_${period.end_date}.xlsx"`)
        .send(Buffer.from(buffer));
    } catch (err) {
      req.log.error(err);
      return reply.code(400).send({ ok: false, error: err.message || String(err) });
    }
  });

  // Stock monitor summary + recent movements
  // GET /api/stock/monitor?part_code=
  app.get("/monitor", async (req, reply) => {
    const part_code = String(req.query?.part_code || "").trim();

    const where = [];
    const params = [];
    if (part_code) {
      where.push("p.part_code LIKE ?");
      params.push(`%${part_code}%`);
    }

    const rows = db.prepare(`
      SELECT
        p.id,
        p.part_code,
        p.part_name,
        p.critical,
        p.min_stock,
        p.unit_cost,
        ${stockCategorySql("p")} AS stock_category,
        COALESCE(p.stock_category_source, 'auto') AS stock_category_source,
        IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      GROUP BY p.id
      ORDER BY p.critical DESC, on_hand ASC, p.part_code ASC
      LIMIT 300
    `).all(...params).map((r) => ({
      ...r,
      critical: Boolean(r.critical),
      stock_category_label: stockCategoryLabel(r.stock_category),
      on_hand: Number(r.on_hand || 0),
      unit_cost: Number(r.unit_cost || 0),
      stock_value: Number((Number(r.on_hand || 0) * Number(r.unit_cost || 0)).toFixed(2)),
      below_min: Number(r.on_hand || 0) < Number(r.min_stock || 0),
    }));

    const summary = {
      total_parts: rows.length,
      below_min: rows.filter((r) => r.below_min).length,
      critical_below_min: rows.filter((r) => r.below_min && r.critical).length,
      total_on_hand: Number(rows.reduce((acc, r) => acc + Number(r.on_hand || 0), 0).toFixed(2)),
      total_stock_value: Number(rows.reduce((acc, r) => acc + Number(r.stock_value || 0), 0).toFixed(2)),
    };

    const movementDateExpr = hasColumn("stock_movements", "created_at")
      ? "sm.created_at"
      : "sm.movement_date";

    const recent = db.prepare(`
      SELECT
        sm.id,
        ${movementDateExpr} AS created_at,
        sm.movement_type,
        sm.quantity,
        sm.reference,
        l.location_code,
        l.location_name,
        p.part_code,
        p.part_name
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      LEFT JOIN stock_locations l ON l.id = sm.location_id
      ORDER BY sm.id DESC
      LIMIT 30
    `).all().map((r) => ({
      ...r,
      quantity: Number(r.quantity || 0),
    }));

    return reply.send({ ok: true, summary, rows, recent });
  });

  // Inventory control summary
  // GET /api/stock/control-summary?part_code=
  app.get("/control-summary", async (req, reply) => {
    const part_code = String(req.query?.part_code || "").trim();

    const totalPartsRow = db.prepare(`SELECT COUNT(*) AS c FROM parts`).get();
    const belowMinRow = db.prepare(`
      SELECT COUNT(*) AS c
      FROM (
        SELECT p.id, IFNULL(SUM(sm.quantity), 0) AS on_hand, p.min_stock
        FROM parts p
        LEFT JOIN stock_movements sm ON sm.part_id = p.id
        GROUP BY p.id
        HAVING on_hand < IFNULL(p.min_stock, 0)
      )
    `).get();
    const lubeBelowRows = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        p.min_stock,
        IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      WHERE ${oilPartSql("p")}
      GROUP BY p.id
      HAVING on_hand < IFNULL(p.min_stock, 0)
      ORDER BY (IFNULL(p.min_stock, 0) - on_hand) DESC, p.part_code ASC
      LIMIT 10
    `).all().map((r) => ({
      ...r,
      on_hand: Number(r.on_hand || 0),
      min_stock: Number(r.min_stock || 0),
      shortage: Number((Number(r.min_stock || 0) - Number(r.on_hand || 0)).toFixed(2)),
    }));

    let part = null;
    let part_summary = null;
    if (part_code) {
      part = getPartByCode.get(part_code);
      if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
      const on_hand = Number(getOnHand.get(part.id)?.on_hand || 0);
      const min_stock = Number(part.min_stock || 0);

      const movementDateExpr = hasColumn("stock_movements", "created_at")
        ? "datetime(created_at)"
        : "datetime(movement_date)";
      const movementSummary = db.prepare(`
        SELECT
          IFNULL(SUM(CASE WHEN quantity > 0 THEN quantity ELSE 0 END), 0) AS qty_in_30d,
          IFNULL(SUM(CASE WHEN quantity < 0 THEN ABS(quantity) ELSE 0 END), 0) AS qty_out_30d,
          IFNULL(SUM(quantity), 0) AS net_30d,
          COUNT(*) AS movement_count_30d
        FROM stock_movements
        WHERE part_id = ?
          AND ${movementDateExpr} >= datetime('now', '-30 days')
      `).get(part.id);
      const movementCount7d = db.prepare(`
        SELECT COUNT(*) AS c
        FROM stock_movements
        WHERE part_id = ?
          AND ${movementDateExpr} >= datetime('now', '-7 days')
      `).get(part.id);

      part_summary = {
        part_code: part.part_code,
        part_name: part.part_name,
        on_hand,
        min_stock,
        below_min: on_hand < min_stock,
        qty_in_30d: Number(movementSummary?.qty_in_30d || 0),
        qty_out_30d: Number(movementSummary?.qty_out_30d || 0),
        net_30d: Number(movementSummary?.net_30d || 0),
        movement_count_30d: Number(movementSummary?.movement_count_30d || 0),
        movement_count_7d: Number(movementCount7d?.c || 0),
      };
    }

    return reply.send({
      ok: true,
      summary: {
        total_parts: Number(totalPartsRow?.c || 0),
        below_min_total: Number(belowMinRow?.c || 0),
        lube_below_min_count: lubeBelowRows.length,
      },
      part: part_summary,
      low_lube_rows: lubeBelowRows,
    });
  });

  // Stock movements for a date range (ledger-style report)
  // GET /api/stock/movements-report?date_from=YYYY-MM-DD&date_to=YYYY-MM-DD&part_code=
  app.get("/movements-report", async (req, reply) => {
    function parseDateOnly(s) {
      const t = String(s || "").trim();
      return /^\d{4}-\d{2}-\d{2}$/.test(t) ? t : null;
    }

    const date_from = parseDateOnly(req.query?.date_from);
    const date_to = parseDateOnly(req.query?.date_to);
    if (!date_from || !date_to) {
      return reply.code(400).send({ error: "date_from and date_to are required (YYYY-MM-DD)" });
    }
    if (date_from > date_to) {
      return reply.code(400).send({ error: "date_from must be on or before date_to" });
    }

    function normalizeStockReportPartFilter(raw) {
      let s = String(raw || "").trim().replace(/\s+/g, " ");
      if (!s) return "";
      // Matches datalist labels "PARTCODE - Description" when users pick or paste the full line.
      const dashIdx = s.indexOf(" - ");
      if (dashIdx > 0) s = s.slice(0, dashIdx).trim();
      return s;
    }

    const part_filter = normalizeStockReportPartFilter(req.query?.part_code);
    const smDateCol = hasColumn("stock_movements", "created_at") ? "sm.created_at" : "sm.movement_date";
    const smDateTimeExpr = `datetime(${smDateCol})`;

    const startDt = `${date_from} 00:00:00`;
    const endDt = `${date_to} 23:59:59`;

    const where = [`${smDateTimeExpr} >= datetime(?)`, `${smDateTimeExpr} <= datetime(?)`];
    const params = [startDt, endDt];

    if (part_filter) {
      const exactPart = db.prepare(`
        SELECT id FROM parts WHERE UPPER(TRIM(part_code)) = UPPER(TRIM(?))
      `).get(part_filter);
      if (exactPart) {
        where.push("sm.part_id = ?");
        params.push(Number(exactPart.id));
      } else {
        // Substring match without LIKE wildcards (% and _ in codes match literally).
        where.push("instr(LOWER(IFNULL(p.part_code,'')), LOWER(?)) > 0");
        params.push(part_filter);
      }
    }

    const whereSql = where.join(" AND ");

    const summaryRow = db.prepare(`
      SELECT
        COUNT(*) AS movement_count,
        IFNULL(SUM(CASE WHEN sm.quantity > 0 THEN sm.quantity ELSE 0 END), 0) AS qty_in,
        IFNULL(SUM(CASE WHEN sm.quantity < 0 THEN ABS(sm.quantity) ELSE 0 END), 0) AS qty_out,
        IFNULL(SUM(sm.quantity), 0) AS net_qty
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      WHERE ${whereSql}
    `).get(...params);

    const totalMatchRow = db.prepare(`
      SELECT COUNT(*) AS c
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      WHERE ${whereSql}
    `).get(...params);

    const maxRows = 5000;
    const rows = db.prepare(`
      SELECT
        sm.id,
        ${smDateCol} AS movement_at,
        sm.movement_type,
        sm.quantity,
        sm.reference,
        l.location_code,
        l.location_name,
        ${hasColumn("stock_movements", "bin_id") ? "b.bin_code" : "NULL"} AS bin_code,
        p.part_code,
        p.part_name
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      LEFT JOIN stock_locations l ON l.id = sm.location_id
      ${hasColumn("stock_movements", "bin_id") ? "LEFT JOIN stock_bins b ON b.id = sm.bin_id" : ""}
      WHERE ${whereSql}
      ORDER BY ${smDateTimeExpr} DESC, sm.id DESC
      LIMIT ${maxRows}
    `).all(...params).map((r) => ({
      ...r,
      id: Number(r.id || 0),
      quantity: Number(r.quantity || 0),
    }));

    const totalMatching = Number(totalMatchRow?.c || 0);

    return reply.send({
      ok: true,
      date_from,
      date_to,
      part_filter: part_filter || null,
      summary: {
        movement_count: Number(summaryRow?.movement_count || 0),
        qty_in: Number(summaryRow?.qty_in || 0),
        qty_out: Number(summaryRow?.qty_out || 0),
        net_qty: Number(summaryRow?.net_qty || 0),
      },
      truncated: totalMatching > rows.length,
      row_limit: maxRows,
      total_matching: totalMatching,
      rows,
    });
  });

  // GET /api/stock/fx-settings — ZAR/MZN per USD (for converting receipts to USD)
  app.get("/fx-settings", async () => {
    return {
      ok: true,
      zar_per_usd: getFxRate("zar_per_usd", 18.5),
      mzn_per_usd: getFxRate("mzn_per_usd", 64),
      note:
        "Receipt unit cost in ZAR or MZN is converted to USD as: USD = local_amount / rate (rate = local currency units per 1 USD).",
    };
  });

  // POST /api/stock/fx-settings — update rates (admin/supervisor/stores)
  app.post("/fx-settings", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const body = req.body || {};
    const zar = body.zar_per_usd != null ? Number(body.zar_per_usd) : null;
    const mzn = body.mzn_per_usd != null ? Number(body.mzn_per_usd) : null;
    const upsert = db.prepare(`
      INSERT INTO cost_settings (key, value, updated_at) VALUES (?, ?, datetime('now'))
      ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = datetime('now')
    `);
    if (zar != null) {
      if (!Number.isFinite(zar) || zar <= 0) return reply.code(400).send({ error: "zar_per_usd must be > 0" });
      upsert.run("zar_per_usd", zar);
    }
    if (mzn != null) {
      if (!Number.isFinite(mzn) || mzn <= 0) return reply.code(400).send({ error: "mzn_per_usd must be > 0" });
      upsert.run("mzn_per_usd", mzn);
    }
    return reply.send({
      ok: true,
      zar_per_usd: getFxRate("zar_per_usd", 18.5),
      mzn_per_usd: getFxRate("mzn_per_usd", 64),
    });
  });
}
