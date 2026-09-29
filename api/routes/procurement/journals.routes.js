// IRONLOG/api/routes/procurement/journals.routes.js — Procurement journals: build, summarise, export, post and reverse.
// Registered by routes/procurement.routes.js; shared helpers arrive through ctx.
import { db } from "../../db/client.js";
import { getUser } from "../../utils/request.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerJournalsRoutes(app, ctx) {
  const {
    isPeriodLocked,
    nextRunNumber,
    periodForDate,
    readCostSetting,
    recalcRunTotals,
    requireRoles,
    tableHasColumn,
  } = ctx;

  app.post("/journals/build", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement", "stores"])) return;
    const start = String(req.body?.start || "").trim();
    const end = String(req.body?.end || "").trim();
    if (!start || !end) return reply.code(400).send({ error: "start and end are required" });
    const batch_id = String(req.body?.batch_id || "").trim() || `JRN-${new Date().toISOString().slice(0, 10)}-${Date.now()}`;
    const defaultCostCenter = req.body?.default_cost_center_code ? String(req.body.default_cost_center_code).trim() : null;
    const tx = db.transaction(() => {
      db.prepare(`DELETE FROM finance_journal_staging WHERE batch_id = ?`).run(batch_id);
      const ins = db.prepare(`
        INSERT INTO finance_journal_staging (
          batch_id, tx_date, source_module, source_type, source_id,
          account_code, cost_center_code, description, debit, credit, currency
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      const receiptRows = db.prepare(`
        SELECT
          r.receipt_date AS tx_date,
          rl.id AS receipt_line_id,
          rl.line_total,
          COALESCE(NULLIF(TRIM(rl.cost_center_code), ''), NULLIF(TRIM(pol.cost_center_code), ''), ?) AS cost_center_code,
          po.currency
        FROM procurement_goods_receipt_lines rl
        JOIN procurement_goods_receipts r ON r.id = rl.receipt_id
        JOIN procurement_purchase_orders po ON po.id = r.po_id
        LEFT JOIN procurement_purchase_order_lines pol ON pol.id = rl.po_line_id
        WHERE DATE(r.receipt_date) BETWEEN DATE(?) AND DATE(?)
      `).all(defaultCostCenter, start, end);
      for (const row of receiptRows) {
        const amount = Number(row.line_total || 0);
        if (amount <= 0) continue;
        const tx_date = String(row.tx_date || start);
        const cc = row.cost_center_code ? String(row.cost_center_code) : null;
        const cur = String(row.currency || "USD");
        ins.run(batch_id, tx_date, "procurement", "goods_receipt", String(row.receipt_line_id), "1400-INVENTORY", cc, `GRN line ${row.receipt_line_id}`, amount, 0, cur);
        ins.run(batch_id, tx_date, "procurement", "goods_receipt", String(row.receipt_line_id), "2100-GRNI", cc, `GRN accrual ${row.receipt_line_id}`, 0, amount, cur);
      }

      const invoiceRows = db.prepare(`
        SELECT
          i.invoice_date AS tx_date,
          il.id AS invoice_line_id,
          il.line_total,
          COALESCE(NULLIF(TRIM(il.cost_center_code), ''), NULLIF(TRIM(pol.cost_center_code), ''), ?) AS cost_center_code,
          i.currency
        FROM procurement_invoice_lines il
        JOIN procurement_invoices i ON i.id = il.invoice_id
        LEFT JOIN procurement_purchase_order_lines pol ON pol.id = il.po_line_id
        WHERE DATE(i.invoice_date) BETWEEN DATE(?) AND DATE(?)
      `).all(defaultCostCenter, start, end);
      for (const row of invoiceRows) {
        const amount = Number(row.line_total || 0);
        if (amount <= 0) continue;
        const tx_date = String(row.tx_date || start);
        const cc = row.cost_center_code ? String(row.cost_center_code) : null;
        const cur = String(row.currency || "USD");
        ins.run(batch_id, tx_date, "procurement", "invoice", String(row.invoice_line_id), "2100-GRNI", cc, `Invoice clear GRNI ${row.invoice_line_id}`, amount, 0, cur);
        ins.run(batch_id, tx_date, "procurement", "invoice", String(row.invoice_line_id), "2200-AP", cc, `AP recognition ${row.invoice_line_id}`, 0, amount, cur);
      }
    });
    tx();
    const summary = db.prepare(`
      SELECT
        COUNT(*) AS lines,
        COALESCE(SUM(debit), 0) AS debit_total,
        COALESCE(SUM(credit), 0) AS credit_total
      FROM finance_journal_staging
      WHERE batch_id = ?
    `).get(batch_id);
    return {
      ok: true,
      batch_id,
      lines: Number(summary?.lines || 0),
      debit_total: Number(Number(summary?.debit_total || 0).toFixed(2)),
      credit_total: Number(Number(summary?.credit_total || 0).toFixed(2)),
      balanced: Number(Number(summary?.debit_total || 0).toFixed(2)) === Number(Number(summary?.credit_total || 0).toFixed(2)),
    };
  });

  app.get("/journals/export.csv", async (req, reply) => {
    const batch_id = String(req.query?.batch_id || "").trim();
    if (!batch_id) return reply.code(400).send({ error: "batch_id is required" });
    const rows = db.prepare(`
      SELECT
        batch_id, tx_date, source_module, source_type, source_id,
        account_code, cost_center_code, description, debit, credit, currency
      FROM finance_journal_staging
      WHERE batch_id = ?
      ORDER BY id ASC
    `).all(batch_id);
    if (!rows.length) return reply.code(404).send({ error: "no journal lines for batch_id" });
    const esc = (v) => {
      const s = String(v ?? "");
      if (/[",\n]/.test(s)) return `"${s.replace(/"/g, "\"\"")}"`;
      return s;
    };
    const header = ["batch_id", "tx_date", "source_module", "source_type", "source_id", "account_code", "cost_center_code", "description", "debit", "credit", "currency"];
    const csv = [header.join(",")]
      .concat(rows.map((r) => [
        r.batch_id,
        r.tx_date,
        r.source_module,
        r.source_type,
        r.source_id,
        r.account_code,
        r.cost_center_code || "",
        r.description || "",
        Number(r.debit || 0).toFixed(2),
        Number(r.credit || 0).toFixed(2),
        r.currency || "USD",
      ].map(esc).join(",")))
      .join("\n");
    reply.header("Content-Type", "text/csv; charset=utf-8");
    reply.header("Content-Disposition", `attachment; filename="journal-${batch_id}.csv"`);
    return reply.send(csv);
  });

  app.get("/journals/export.xlsx", async (req, reply) => {
    const batch_id = String(req.query?.batch_id || "").trim();
    if (!batch_id) return reply.code(400).send({ error: "batch_id is required" });
    const rows = db.prepare(`
      SELECT
        batch_id, tx_date, source_module, source_type, source_id,
        account_code, cost_center_code, description, debit, credit, currency
      FROM finance_journal_staging
      WHERE batch_id = ?
      ORDER BY id ASC
    `).all(batch_id);
    if (!rows.length) return reply.code(404).send({ error: "no journal lines for batch_id" });
    const ExcelJS = (await import("exceljs")).default;
    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();
    const ws = wb.addWorksheet("Journals");
    ws.columns = [
      { header: "Batch ID", key: "batch_id", width: 20 },
      { header: "Date", key: "tx_date", width: 14 },
      { header: "Source Module", key: "source_module", width: 18 },
      { header: "Source Type", key: "source_type", width: 18 },
      { header: "Source ID", key: "source_id", width: 14 },
      { header: "Account", key: "account_code", width: 18 },
      { header: "Cost Center", key: "cost_center_code", width: 16 },
      { header: "Description", key: "description", width: 40 },
      { header: "Debit", key: "debit", width: 14 },
      { header: "Credit", key: "credit", width: 14 },
      { header: "Currency", key: "currency", width: 10 },
    ];
    ws.addRows(rows.map((r) => ({
      ...r,
      debit: Number(r.debit || 0),
      credit: Number(r.credit || 0),
    })));
    ws.getRow(1).font = { bold: true };
    ws.getColumn("debit").numFmt = "#,##0.00";
    ws.getColumn("credit").numFmt = "#,##0.00";
    const buffer = await wb.xlsx.writeBuffer();
    reply.header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    reply.header("Content-Disposition", `attachment; filename="journal-${batch_id}.xlsx"`);
    return reply.send(Buffer.from(buffer));
  });

  app.post("/journals/summarize", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement", "stores", "finance", "plant_manager"])) return;
    const start = String(req.body?.start || "").trim();
    const end = String(req.body?.end || "").trim();
    if (!start || !end) return reply.code(400).send({ error: "start and end are required (YYYY-MM-DD)" });

    const categoriesInput = Array.isArray(req.body?.categories) ? req.body.categories : null;
    const categorySet = new Set(
      (categoriesInput || ["parts", "labor", "downtime", "fuel", "lube", "procurement_grn", "procurement_ap"])
        .map((x) => String(x || "").trim().toLowerCase())
        .filter(Boolean)
    );

    const defaultCostCenter = req.body?.default_cost_center_code ? String(req.body.default_cost_center_code).trim() : null;
    const currency = String(req.body?.currency || "USD").trim().toUpperCase() || "USD";
    const notes = req.body?.notes ? String(req.body.notes).trim() : null;
    const period = periodForDate(start);

    if (isPeriodLocked(period)) {
      return reply.code(409).send({ error: `period ${period} is locked` });
    }

    const fuelDefault = readCostSetting("fuel_cost_per_liter_default", 1.5);
    const lubeDefault = readCostSetting("lube_cost_per_qty_default", 4.0);
    const laborDefault = readCostSetting("labor_cost_per_hour_default", 35.0);
    const downtimeDefault = readCostSetting("downtime_cost_per_hour_default", 120.0);

    const assetsHasCC = tableHasColumn("assets", "cost_center_code");
    const assetsHasSite = tableHasColumn("assets", "site_code");

    const run_number = nextRunNumber(period);
    const user = getUser(req);

    const insRun = db.prepare(`
      INSERT INTO finance_posting_runs
        (run_number, period, start_date, end_date, run_type, status, currency, notes, created_by)
      VALUES (?, ?, ?, ?, 'summary', 'draft', ?, ?, ?)
    `);
    const insLine = db.prepare(`
      INSERT INTO finance_posting_run_lines
        (run_id, category, tx_date, source_module, source_type, source_ref,
         account_code, cost_center_code, site_code, equipment_type, asset_code,
         description, debit, credit, currency, qty, unit_cost)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    const tx = db.transaction(() => {
      const r = insRun.run(run_number, period, start, end, currency, notes, user);
      const runId = Number(r.lastInsertRowid);

      if (categorySet.has("parts")) {
        const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
        const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
        const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
        const smHasCC = smCols.some((c) => String(c.name) === "cost_center_code");
        const rows = db.prepare(`
          SELECT
            ${smDateExpr} AS tx_date,
            COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
            COALESCE(a.category, '') AS equipment_type,
            ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
            ${smHasCC
              ? `COALESCE(NULLIF(TRIM(sm.cost_center_code), ''),
                         ${assetsHasCC ? "NULLIF(TRIM(a.cost_center_code), '')," : ""}
                         ?)`
              : `COALESCE(${assetsHasCC ? "NULLIF(TRIM(a.cost_center_code), ''), " : ""}?)`} AS cost_center_code,
            SUM(ABS(sm.quantity)) AS qty,
            SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)) AS amount
          FROM stock_movements sm
          JOIN parts p ON p.id = sm.part_id
          LEFT JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
          LEFT JOIN assets a ON a.id = w.asset_id
          WHERE sm.movement_type = 'out'
            AND ${smDateExpr} BETWEEN DATE(?) AND DATE(?)
          GROUP BY tx_date, asset_code, cost_center_code, equipment_type, site_code
        `).all(defaultCostCenter, start, end);
        for (const row of rows) {
          const amount = Number(row.amount || 0);
          if (amount <= 0) continue;
          const cc = row.cost_center_code || null;
          const site = row.site_code || null;
          const eq = row.equipment_type || null;
          const ref = `parts:${row.tx_date}:${row.asset_code}`;
          insLine.run(runId, "parts", row.tx_date, "maintenance", "parts_issue", ref,
            "5200-PARTS-EXPENSE", cc, site, eq, row.asset_code,
            `Parts issued ${row.asset_code} on ${row.tx_date}`, amount, 0, currency, Number(row.qty || 0), null);
          insLine.run(runId, "parts", row.tx_date, "maintenance", "parts_issue", ref,
            "1400-INVENTORY", cc, site, eq, row.asset_code,
            `Parts inventory reduction ${row.asset_code}`, 0, amount, currency, Number(row.qty || 0), null);
        }
      }

      if (categorySet.has("labor")) {
        const woCols = db.prepare(`PRAGMA table_info(work_orders)`).all();
        const woHasCompleted = woCols.some((c) => String(c.name) === "completed_at");
        const woHasClosed = woCols.some((c) => String(c.name) === "closed_at");
        const woHasLaborHours = woCols.some((c) => String(c.name) === "labor_hours");
        const woHasLaborRate = woCols.some((c) => String(c.name) === "labor_rate_per_hour");
        const woHasCC = woCols.some((c) => String(c.name) === "cost_center_code");
        if (woHasLaborHours) {
          const dateExpr = woHasCompleted && woHasClosed
            ? "DATE(COALESCE(w.completed_at, w.closed_at))"
            : (woHasCompleted ? "DATE(w.completed_at)" : "DATE(w.closed_at)");
          const rows = db.prepare(`
            SELECT
              ${dateExpr} AS tx_date,
              COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
              COALESCE(a.category, '') AS equipment_type,
              ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
              ${woHasCC
                ? `COALESCE(NULLIF(TRIM(w.cost_center_code), ''),
                           ${assetsHasCC ? "NULLIF(TRIM(a.cost_center_code), ''), " : ""} ?)`
                : `COALESCE(${assetsHasCC ? "NULLIF(TRIM(a.cost_center_code), ''), " : ""} ?)`} AS cost_center_code,
              SUM(COALESCE(w.labor_hours, 0)) AS hours,
              SUM(COALESCE(w.labor_hours, 0) * COALESCE(${woHasLaborRate ? "w.labor_rate_per_hour" : "NULL"}, ?)) AS amount
            FROM work_orders w
            LEFT JOIN assets a ON a.id = w.asset_id
            WHERE w.status IN ('completed','approved','closed')
              AND ${dateExpr} BETWEEN DATE(?) AND DATE(?)
            GROUP BY tx_date, asset_code, cost_center_code, equipment_type, site_code
          `).all(defaultCostCenter, laborDefault, start, end);
          for (const row of rows) {
            const amount = Number(row.amount || 0);
            if (amount <= 0) continue;
            const cc = row.cost_center_code || null;
            const site = row.site_code || null;
            const eq = row.equipment_type || null;
            const ref = `labor:${row.tx_date}:${row.asset_code}`;
            insLine.run(runId, "labor", row.tx_date, "maintenance", "labor", ref,
              "5300-LABOR-EXPENSE", cc, site, eq, row.asset_code,
              `Labor ${row.asset_code} ${Number(row.hours).toFixed(2)}h`, amount, 0, currency, Number(row.hours || 0), null);
            insLine.run(runId, "labor", row.tx_date, "maintenance", "labor", ref,
              "2500-LABOR-CLEARING", cc, site, eq, row.asset_code,
              `Labor clearing ${row.asset_code}`, 0, amount, currency, Number(row.hours || 0), null);
          }
        }
      }

      if (categorySet.has("downtime")) {
        const rows = db.prepare(`
          SELECT
            l.log_date AS tx_date,
            COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
            COALESCE(a.category, '') AS equipment_type,
            ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
            ${assetsHasCC ? "COALESCE(NULLIF(TRIM(a.cost_center_code), ''), ?)" : "?"} AS cost_center_code,
            SUM(COALESCE(l.hours_down, 0)) AS hours,
            SUM(COALESCE(l.hours_down, 0) * COALESCE(a.downtime_cost_per_hour, ?)) AS amount
          FROM breakdown_downtime_logs l
          JOIN breakdowns b ON b.id = l.breakdown_id
          JOIN assets a ON a.id = b.asset_id
          WHERE l.log_date BETWEEN DATE(?) AND DATE(?)
          GROUP BY tx_date, asset_code, cost_center_code, equipment_type, site_code
        `).all(defaultCostCenter, downtimeDefault, start, end);
        for (const row of rows) {
          const amount = Number(row.amount || 0);
          if (amount <= 0) continue;
          const cc = row.cost_center_code || null;
          const site = row.site_code || null;
          const eq = row.equipment_type || null;
          const ref = `downtime:${row.tx_date}:${row.asset_code}`;
          insLine.run(runId, "downtime", row.tx_date, "maintenance", "downtime", ref,
            "5400-DOWNTIME-EXPENSE", cc, site, eq, row.asset_code,
            `Downtime ${row.asset_code} ${Number(row.hours).toFixed(2)}h`, amount, 0, currency, Number(row.hours || 0), null);
          insLine.run(runId, "downtime", row.tx_date, "maintenance", "downtime", ref,
            "2600-DOWNTIME-CLEARING", cc, site, eq, row.asset_code,
            `Downtime clearing ${row.asset_code}`, 0, amount, currency, Number(row.hours || 0), null);
        }
      }

      if (categorySet.has("fuel")) {
        const flCols = db.prepare(`PRAGMA table_info(fuel_logs)`).all();
        const flHasUnit = flCols.some((c) => String(c.name) === "unit_cost_per_liter");
        const rows = db.prepare(`
          SELECT
            fl.log_date AS tx_date,
            COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
            COALESCE(a.category, '') AS equipment_type,
            ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
            ${assetsHasCC ? "COALESCE(NULLIF(TRIM(a.cost_center_code), ''), ?)" : "?"} AS cost_center_code,
            SUM(COALESCE(fl.liters, 0)) AS qty,
            SUM(COALESCE(fl.liters, 0) * COALESCE(${flHasUnit ? "fl.unit_cost_per_liter" : "NULL"}, a.fuel_cost_per_liter, ?)) AS amount
          FROM fuel_logs fl
          JOIN assets a ON a.id = fl.asset_id
          WHERE fl.log_date BETWEEN DATE(?) AND DATE(?)
          GROUP BY tx_date, asset_code, cost_center_code, equipment_type, site_code
        `).all(defaultCostCenter, fuelDefault, start, end);
        for (const row of rows) {
          const amount = Number(row.amount || 0);
          if (amount <= 0) continue;
          const cc = row.cost_center_code || null;
          const site = row.site_code || null;
          const eq = row.equipment_type || null;
          const ref = `fuel:${row.tx_date}:${row.asset_code}`;
          insLine.run(runId, "fuel", row.tx_date, "operations", "fuel", ref,
            "5100-FUEL-EXPENSE", cc, site, eq, row.asset_code,
            `Fuel ${row.asset_code} ${Number(row.qty).toFixed(2)}L`, amount, 0, currency, Number(row.qty || 0), null);
          insLine.run(runId, "fuel", row.tx_date, "operations", "fuel", ref,
            "1410-FUEL-INVENTORY", cc, site, eq, row.asset_code,
            `Fuel inventory reduction ${row.asset_code}`, 0, amount, currency, Number(row.qty || 0), null);
        }
      }

      if (categorySet.has("lube")) {
        const olCols = db.prepare(`PRAGMA table_info(oil_logs)`).all();
        const olHasUnit = olCols.some((c) => String(c.name) === "unit_cost");
        const rows = db.prepare(`
          SELECT
            ol.log_date AS tx_date,
            COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
            COALESCE(a.category, '') AS equipment_type,
            ${assetsHasSite ? "COALESCE(a.site_code, '')" : "''"} AS site_code,
            ${assetsHasCC ? "COALESCE(NULLIF(TRIM(a.cost_center_code), ''), ?)" : "?"} AS cost_center_code,
            SUM(COALESCE(ol.quantity, 0)) AS qty,
            SUM(COALESCE(ol.quantity, 0) * COALESCE(${olHasUnit ? "ol.unit_cost" : "NULL"}, ?)) AS amount
          FROM oil_logs ol
          JOIN assets a ON a.id = ol.asset_id
          WHERE ol.log_date BETWEEN DATE(?) AND DATE(?)
          GROUP BY tx_date, asset_code, cost_center_code, equipment_type, site_code
        `).all(defaultCostCenter, lubeDefault, start, end);
        for (const row of rows) {
          const amount = Number(row.amount || 0);
          if (amount <= 0) continue;
          const cc = row.cost_center_code || null;
          const site = row.site_code || null;
          const eq = row.equipment_type || null;
          const ref = `lube:${row.tx_date}:${row.asset_code}`;
          insLine.run(runId, "lube", row.tx_date, "operations", "lube", ref,
            "5150-LUBE-EXPENSE", cc, site, eq, row.asset_code,
            `Lube ${row.asset_code} ${Number(row.qty).toFixed(2)}`, amount, 0, currency, Number(row.qty || 0), null);
          insLine.run(runId, "lube", row.tx_date, "operations", "lube", ref,
            "1420-LUBE-INVENTORY", cc, site, eq, row.asset_code,
            `Lube inventory reduction ${row.asset_code}`, 0, amount, currency, Number(row.qty || 0), null);
        }
      }

      if (categorySet.has("procurement_grn")) {
        const rows = db.prepare(`
          SELECT
            r.receipt_date AS tx_date,
            rl.id AS ref_id,
            rl.line_total AS amount,
            COALESCE(NULLIF(TRIM(rl.cost_center_code), ''), NULLIF(TRIM(pol.cost_center_code), ''), ?) AS cost_center_code,
            po.currency
          FROM procurement_goods_receipt_lines rl
          JOIN procurement_goods_receipts r ON r.id = rl.receipt_id
          JOIN procurement_purchase_orders po ON po.id = r.po_id
          LEFT JOIN procurement_purchase_order_lines pol ON pol.id = rl.po_line_id
          WHERE DATE(r.receipt_date) BETWEEN DATE(?) AND DATE(?)
        `).all(defaultCostCenter, start, end);
        for (const row of rows) {
          const amount = Number(row.amount || 0);
          if (amount <= 0) continue;
          const cur = String(row.currency || currency);
          insLine.run(runId, "procurement_grn", row.tx_date, "procurement", "goods_receipt", `grn:${row.ref_id}`,
            "1400-INVENTORY", row.cost_center_code || null, null, null, null,
            `GRN line ${row.ref_id}`, amount, 0, cur, null, null);
          insLine.run(runId, "procurement_grn", row.tx_date, "procurement", "goods_receipt", `grn:${row.ref_id}`,
            "2100-GRNI", row.cost_center_code || null, null, null, null,
            `GRN accrual ${row.ref_id}`, 0, amount, cur, null, null);
        }
      }

      if (categorySet.has("procurement_ap")) {
        const rows = db.prepare(`
          SELECT
            i.invoice_date AS tx_date,
            il.id AS ref_id,
            il.line_total AS amount,
            COALESCE(NULLIF(TRIM(il.cost_center_code), ''), NULLIF(TRIM(pol.cost_center_code), ''), ?) AS cost_center_code,
            i.currency
          FROM procurement_invoice_lines il
          JOIN procurement_invoices i ON i.id = il.invoice_id
          LEFT JOIN procurement_purchase_order_lines pol ON pol.id = il.po_line_id
          WHERE DATE(i.invoice_date) BETWEEN DATE(?) AND DATE(?)
        `).all(defaultCostCenter, start, end);
        for (const row of rows) {
          const amount = Number(row.amount || 0);
          if (amount <= 0) continue;
          const cur = String(row.currency || currency);
          insLine.run(runId, "procurement_ap", row.tx_date, "procurement", "invoice", `inv:${row.ref_id}`,
            "2100-GRNI", row.cost_center_code || null, null, null, null,
            `Invoice clear GRNI ${row.ref_id}`, amount, 0, cur, null, null);
          insLine.run(runId, "procurement_ap", row.tx_date, "procurement", "invoice", `inv:${row.ref_id}`,
            "2200-AP", row.cost_center_code || null, null, null, null,
            `AP recognition ${row.ref_id}`, 0, amount, cur, null, null);
        }
      }

      recalcRunTotals(runId);
      try {
        writeAudit(db, req, {
          module: "finance",
          action: "journals.summarize",
          entity_type: "finance_posting_runs",
          entity_id: runId,
          payload: { run_number, period, start, end, categories: Array.from(categorySet) }
        });
      } catch {}
      return runId;
    });

    let runId;
    try { runId = tx(); }
    catch (e) { return reply.code(500).send({ error: `summarize failed: ${e.message}` }); }

    const header = db.prepare(`SELECT * FROM finance_posting_runs WHERE id = ?`).get(runId);
    return reply.send({
      ok: true,
      run: header,
      balanced: Number(Number(header?.total_debit || 0).toFixed(2)) === Number(Number(header?.total_credit || 0).toFixed(2))
    });
  });

  app.get("/journals/runs", async (req) => {
    const status = req.query?.status ? String(req.query.status).trim().toLowerCase() : null;
    const period = req.query?.period ? String(req.query.period).trim() : null;
    const where = [];
    const args = [];
    if (status) { where.push(`LOWER(status) = ?`); args.push(status); }
    if (period) { where.push(`period = ?`); args.push(period); }
    const sql = `
      SELECT id, run_number, period, start_date, end_date, run_type, status,
             currency, total_debit, total_credit, line_count, created_by,
             created_at, exported_at, posted_at, reversed_at
      FROM finance_posting_runs
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY id DESC LIMIT 500
    `;
    const rows = db.prepare(sql).all(...args);
    return { ok: true, rows };
  });

  app.get("/journals/runs/:id", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const run = db.prepare(`SELECT * FROM finance_posting_runs WHERE id = ?`).get(id);
    if (!run) return reply.code(404).send({ error: "run not found" });
    const byCategory = db.prepare(`
      SELECT category,
             COUNT(*) AS lines,
             COALESCE(SUM(debit), 0) AS debit_total,
             COALESCE(SUM(credit), 0) AS credit_total
      FROM finance_posting_run_lines
      WHERE run_id = ?
      GROUP BY category
      ORDER BY category
    `).all(id);
    const bySite = db.prepare(`
      SELECT COALESCE(NULLIF(TRIM(site_code), ''), '(unassigned)') AS site_code,
             COUNT(*) AS lines,
             COALESCE(SUM(debit), 0) AS debit_total,
             COALESCE(SUM(credit), 0) AS credit_total
      FROM finance_posting_run_lines
      WHERE run_id = ?
      GROUP BY COALESCE(NULLIF(TRIM(site_code), ''), '(unassigned)')
      ORDER BY debit_total DESC
    `).all(id);
    return { ok: true, run, by_category: byCategory, by_site: bySite };
  });

  app.get("/journals/runs/:id/lines", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const category = req.query?.category ? String(req.query.category).trim().toLowerCase() : null;
    const where = [`run_id = ?`];
    const args = [id];
    if (category) { where.push(`LOWER(category) = ?`); args.push(category); }
    const rows = db.prepare(`
      SELECT id, category, tx_date, source_module, source_type, source_ref,
             account_code, cost_center_code, site_code, equipment_type, asset_code,
             description, debit, credit, currency, qty, unit_cost
      FROM finance_posting_run_lines
      WHERE ${where.join(" AND ")}
      ORDER BY tx_date ASC, id ASC
      LIMIT 5000
    `).all(...args);
    return { ok: true, rows };
  });

  app.post("/journals/runs/:id/mark-exported", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "procurement", "finance", "plant_manager"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const run = db.prepare(`SELECT id, status FROM finance_posting_runs WHERE id = ?`).get(id);
    if (!run) return reply.code(404).send({ error: "run not found" });
    if (String(run.status).toLowerCase() === "posted") return reply.code(409).send({ error: "run already posted" });
    db.prepare(`
      UPDATE finance_posting_runs
      SET status = 'exported', exported_at = datetime('now'), exported_by = ?
      WHERE id = ?
    `).run(getUser(req), id);
    try {
      writeAudit(db, req, { module: "finance", action: "journals.mark_exported", entity_type: "finance_posting_runs", entity_id: id, payload: {} });
    } catch {}
    return { ok: true, id, status: "exported" };
  });

  app.post("/journals/runs/:id/mark-posted", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const run = db.prepare(`SELECT id, status, period FROM finance_posting_runs WHERE id = ?`).get(id);
    if (!run) return reply.code(404).send({ error: "run not found" });
    if (String(run.status).toLowerCase() === "posted") return { ok: true, id, status: "posted", duplicate: true };
    if (String(run.status).toLowerCase() === "reversed") return reply.code(409).send({ error: "run reversed" });
    if (isPeriodLocked(run.period)) return reply.code(409).send({ error: `period ${run.period} is locked` });
    const ref = req.body?.posted_reference ? String(req.body.posted_reference).trim() : null;
    db.prepare(`
      UPDATE finance_posting_runs
      SET status = 'posted', posted_at = datetime('now'), posted_by = ?, posted_reference = ?
      WHERE id = ?
    `).run(getUser(req), ref, id);
    try {
      writeAudit(db, req, { module: "finance", action: "journals.mark_posted", entity_type: "finance_posting_runs", entity_id: id, payload: { posted_reference: ref } });
    } catch {}
    return { ok: true, id, status: "posted", posted_reference: ref };
  });

  app.post("/journals/runs/:id/reverse", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const reason = req.body?.reason ? String(req.body.reason).trim() : "";
    if (!reason) return reply.code(400).send({ error: "reason is required" });
    const run = db.prepare(`SELECT id, status, period FROM finance_posting_runs WHERE id = ?`).get(id);
    if (!run) return reply.code(404).send({ error: "run not found" });
    if (String(run.status).toLowerCase() === "reversed") return { ok: true, id, status: "reversed", duplicate: true };
    if (isPeriodLocked(run.period)) return reply.code(409).send({ error: `period ${run.period} is locked` });
    db.prepare(`
      UPDATE finance_posting_runs
      SET status = 'reversed', reversed_at = datetime('now'), reversed_by = ?, reversed_reason = ?
      WHERE id = ?
    `).run(getUser(req), reason, id);
    try {
      writeAudit(db, req, { module: "finance", action: "journals.reverse", entity_type: "finance_posting_runs", entity_id: id, payload: { reason } });
    } catch {}
    return { ok: true, id, status: "reversed" };
  });

  app.get("/journals/runs/:id/export.csv", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const run = db.prepare(`SELECT * FROM finance_posting_runs WHERE id = ?`).get(id);
    if (!run) return reply.code(404).send({ error: "run not found" });
    const rows = db.prepare(`
      SELECT
        ? AS run_number, category, tx_date, source_module, source_type, source_ref,
        account_code, cost_center_code, site_code, equipment_type, asset_code,
        description, debit, credit, currency, qty, unit_cost
      FROM finance_posting_run_lines
      WHERE run_id = ?
      ORDER BY tx_date ASC, id ASC
    `).all(run.run_number, id);
    if (!rows.length) return reply.code(404).send({ error: "no lines" });
    const esc = (v) => {
      const s = String(v ?? "");
      if (/[",\n]/.test(s)) return `"${s.replace(/"/g, "\"\"")}"`;
      return s;
    };
    const header = ["run_number", "category", "tx_date", "source_module", "source_type", "source_ref",
      "account_code", "cost_center_code", "site_code", "equipment_type", "asset_code",
      "description", "debit", "credit", "currency", "qty", "unit_cost"];
    const csv = [header.join(",")]
      .concat(rows.map((r) => [
        r.run_number, r.category, r.tx_date, r.source_module, r.source_type, r.source_ref || "",
        r.account_code, r.cost_center_code || "", r.site_code || "", r.equipment_type || "", r.asset_code || "",
        r.description || "",
        Number(r.debit || 0).toFixed(2), Number(r.credit || 0).toFixed(2),
        r.currency || "USD",
        r.qty == null ? "" : Number(r.qty).toFixed(4),
        r.unit_cost == null ? "" : Number(r.unit_cost).toFixed(4),
      ].map(esc).join(",")))
      .join("\n");
    reply.header("Content-Type", "text/csv; charset=utf-8");
    reply.header("Content-Disposition", `attachment; filename="${run.run_number}.csv"`);
    return reply.send(csv);
  });

  app.get("/journals/runs/:id/export.xlsx", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const run = db.prepare(`SELECT * FROM finance_posting_runs WHERE id = ?`).get(id);
    if (!run) return reply.code(404).send({ error: "run not found" });
    const rows = db.prepare(`
      SELECT category, tx_date, source_module, source_type, source_ref,
             account_code, cost_center_code, site_code, equipment_type, asset_code,
             description, debit, credit, currency, qty, unit_cost
      FROM finance_posting_run_lines
      WHERE run_id = ?
      ORDER BY tx_date ASC, id ASC
    `).all(id);
    if (!rows.length) return reply.code(404).send({ error: "no lines" });
    const ExcelJS = (await import("exceljs")).default;
    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();
    const ws = wb.addWorksheet("PostingRun");
    ws.columns = [
      { header: "Run Number", key: "run_number", width: 22 },
      { header: "Category", key: "category", width: 16 },
      { header: "Date", key: "tx_date", width: 14 },
      { header: "Source Module", key: "source_module", width: 18 },
      { header: "Source Type", key: "source_type", width: 18 },
      { header: "Source Ref", key: "source_ref", width: 28 },
      { header: "Account", key: "account_code", width: 20 },
      { header: "Cost Center", key: "cost_center_code", width: 16 },
      { header: "Site", key: "site_code", width: 14 },
      { header: "Equipment Type", key: "equipment_type", width: 18 },
      { header: "Asset", key: "asset_code", width: 16 },
      { header: "Description", key: "description", width: 40 },
      { header: "Debit", key: "debit", width: 14 },
      { header: "Credit", key: "credit", width: 14 },
      { header: "Currency", key: "currency", width: 10 },
      { header: "Qty", key: "qty", width: 12 },
      { header: "Unit Cost", key: "unit_cost", width: 14 },
    ];
    ws.addRows(rows.map((r) => ({
      ...r,
      run_number: run.run_number,
      debit: Number(r.debit || 0),
      credit: Number(r.credit || 0),
      qty: r.qty == null ? null : Number(r.qty),
      unit_cost: r.unit_cost == null ? null : Number(r.unit_cost),
    })));
    ws.getRow(1).font = { bold: true };
    ws.getColumn("debit").numFmt = "#,##0.00";
    ws.getColumn("credit").numFmt = "#,##0.00";
    ws.getColumn("qty").numFmt = "#,##0.0000";
    ws.getColumn("unit_cost").numFmt = "#,##0.0000";
    const buffer = await wb.xlsx.writeBuffer();
    reply.header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    reply.header("Content-Disposition", `attachment; filename="${run.run_number}.xlsx"`);
    return reply.send(Buffer.from(buffer));
  });
}
