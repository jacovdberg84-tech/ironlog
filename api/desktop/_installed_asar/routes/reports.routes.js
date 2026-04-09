// IRONLOG/api/routes/reports.routes.js
import path from "node:path";
import fs from "node:fs";
import ExcelJS from "exceljs";
import { db } from "../db/client.js";
import {
  buildPdfBuffer,
  tryDrawLogo,
  sectionTitle,
  kvGrid,
  table,
  ensurePageSpace,
} from "../utils/pdfGenerator.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function yn(v) {
  return v ? "YES" : "NO";
}

function fmtNum(v, dp = 1) {
  if (v == null || v === "") return "";
  const n = Number(v);
  if (!Number.isFinite(n)) return String(v);
  return dp === 0 ? String(Math.round(n)) : n.toFixed(dp).replace(/\.0$/, "");
}

function compactCell(v, max = 140) {
  const s = String(v ?? "").replace(/\s+/g, " ").trim();
  if (!s) return "";
  return s.length > max ? `${s.slice(0, Math.max(1, max - 1))}...` : s;
}

function isMonth(m) {
  return /^\d{4}-\d{2}$/.test(String(m || "").trim());
}

function monthRange(monthStr) {
  const [y, m] = String(monthStr).split("-").map((n) => Number(n));
  const start = new Date(Date.UTC(y, m - 1, 1));
  const end = new Date(Date.UTC(y, m, 0));
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

function prevMonth(monthStr) {
  const [y, m] = String(monthStr).split("-").map((n) => Number(n));
  const d = new Date(Date.UTC(y, m - 1, 1));
  d.setUTCMonth(d.getUTCMonth() - 1);
  const yy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  return `${yy}-${mm}`;
}

function kpiDaily(date, scheduled) {
  const usedRow = db.prepare(`
    SELECT COUNT(DISTINCT dh.asset_id) AS used_assets
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date = ?
      AND dh.is_used = 1
      AND dh.hours_run > 0
      AND a.active = 1
      AND a.is_standby = 0
  `).get(date);

  const used_assets = Number(usedRow.used_assets || 0);
  const available_hours = used_assets * scheduled;

  const runRow = db.prepare(`
    SELECT IFNULL(SUM(dh.hours_run), 0) AS run_hours
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date = ?
      AND dh.is_used = 1
      AND dh.hours_run > 0
      AND a.active = 1
      AND a.is_standby = 0
  `).get(date);

  const run_hours = Number(runRow.run_hours || 0);

  const dtRow = db.prepare(`
    SELECT IFNULL(SUM(l.hours_down), 0) AS downtime_hours
    FROM breakdown_downtime_logs l
    JOIN breakdowns b ON b.id = l.breakdown_id
    JOIN assets a ON a.id = b.asset_id
    WHERE l.log_date = ?
      AND a.active = 1
      AND a.is_standby = 0
      AND b.asset_id IN (
        SELECT DISTINCT dh.asset_id
        FROM daily_hours dh
        WHERE dh.work_date = ?
          AND dh.is_used = 1
          AND dh.hours_run > 0
      )
  `).get(date, date);

  const downtime_hours = Number(dtRow.downtime_hours || 0);

  const availability = available_hours > 0 ? ((available_hours - downtime_hours) / available_hours) * 100 : null;
  const utilization = available_hours > 0 ? (run_hours / available_hours) * 100 : null;

  return {
    used_assets,
    available_hours,
    run_hours,
    downtime_hours,
    availability: availability == null ? null : Number(availability.toFixed(2)),
    utilization: utilization == null ? null : Number(utilization.toFixed(2)),
  };
}

function kpiRange(start, end, scheduled) {
  const daily = db.prepare(`
    SELECT
      dh.work_date,
      COUNT(DISTINCT dh.asset_id) AS used_assets,
      IFNULL(SUM(dh.hours_run), 0) AS run_hours
    FROM daily_hours dh
    JOIN assets a ON a.id = dh.asset_id
    WHERE dh.work_date BETWEEN ? AND ?
      AND dh.is_used = 1
      AND dh.hours_run > 0
      AND a.active = 1
      AND a.is_standby = 0
    GROUP BY dh.work_date
    ORDER BY dh.work_date
  `).all(start, end);

  const available_hours = daily.reduce((acc, d) => acc + (Number(d.used_assets) * scheduled), 0);
  const run_hours = daily.reduce((acc, d) => acc + Number(d.run_hours || 0), 0);

  const dtRow = db.prepare(`
    SELECT IFNULL(SUM(downtime_hours), 0) AS downtime_hours
    FROM breakdowns
    WHERE breakdown_date BETWEEN ? AND ?
  `).get(start, end);

  const downtime_hours = Number(dtRow.downtime_hours || 0);

  const availability = available_hours > 0 ? ((available_hours - downtime_hours) / available_hours) * 100 : null;
  const utilization = available_hours > 0 ? (run_hours / available_hours) * 100 : null;

  return {
    available_hours,
    run_hours,
    downtime_hours,
    availability: availability == null ? null : Number(availability.toFixed(2)),
    utilization: utilization == null ? null : Number(utilization.toFixed(2)),
    daily: daily.map(d => ({
      date: d.work_date,
      used_assets: Number(d.used_assets),
      available_hours: Number(d.used_assets) * scheduled,
      run_hours: Number(d.run_hours || 0),
    })),
  };
}

// ---- Excel helpers ----
function addTableSheet(workbook, name, columns, rows) {
  const ws = workbook.addWorksheet(name);
  ws.columns = columns.map(c => ({ header: c.header, key: c.key, width: c.width ?? 18 }));
  ws.getRow(1).font = { bold: true };

  for (const r of rows) ws.addRow(r);

  ws.views = [{ state: "frozen", ySplit: 1 }];

  // simple borders
  const lastRow = ws.rowCount;
  const lastCol = ws.columnCount;
  for (let i = 1; i <= lastRow; i++) {
    for (let j = 1; j <= lastCol; j++) {
      const cell = ws.getCell(i, j);
      cell.border = {
        top: { style: "thin", color: { argb: "FF2A2A2A" } },
        left: { style: "thin", color: { argb: "FF2A2A2A" } },
        bottom: { style: "thin", color: { argb: "FF2A2A2A" } },
        right: { style: "thin", color: { argb: "FF2A2A2A" } },
      };
    }
  }

  return ws;
}

export default async function reportsRoutes(app) {
  const dataRoot = process.env.IRONLOG_DATA_DIR || process.cwd();
  function hasColumn(table, col) {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name || "") === String(col));
  }
  function pickExistingColumn(table, candidates, fallback) {
    for (const c of candidates) {
      if (hasColumn(table, c)) return c;
    }
    return fallback;
  }
  function resolveStorageAbs(relPath) {
    const rel = String(relPath || "").replace(/\\/g, "/").replace(/^\/+/, "");
    if (!rel) return "";
    const fromData = path.join(dataRoot, rel);
    if (fs.existsSync(fromData)) return fromData;
    return path.join(process.cwd(), rel);
  }

  function ensureColumn(table, colName, colDef) {
    if (!hasColumn(table, colName)) {
      db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
    }
  }

  ensureColumn("assets", "fuel_cost_per_liter", "fuel_cost_per_liter REAL");
  ensureColumn("assets", "downtime_cost_per_hour", "downtime_cost_per_hour REAL");
  ensureColumn("parts", "unit_cost", "unit_cost REAL DEFAULT 0");
  ensureColumn("oil_logs", "unit_cost", "unit_cost REAL");
  ensureColumn("fuel_logs", "unit_cost_per_liter", "unit_cost_per_liter REAL");
  ensureColumn("fuel_logs", "hours_run", "hours_run REAL");
  ensureColumn("work_orders", "labor_hours", "labor_hours REAL DEFAULT 0");
  ensureColumn("work_orders", "labor_rate_per_hour", "labor_rate_per_hour REAL");

  db.prepare(`
    CREATE TABLE IF NOT EXISTS cost_settings (
      key TEXT PRIMARY KEY,
      value REAL NOT NULL DEFAULT 0,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  const upsertCostSetting = db.prepare(`
    INSERT INTO cost_settings (key, value, updated_at)
    VALUES (?, ?, datetime('now'))
    ON CONFLICT(key) DO NOTHING
  `);
  upsertCostSetting.run("fuel_cost_per_liter_default", 1.5);
  upsertCostSetting.run("lube_cost_per_qty_default", 4.0);
  upsertCostSetting.run("labor_cost_per_hour_default", 35.0);
  upsertCostSetting.run("downtime_cost_per_hour_default", 120.0);
  db.prepare(`
    CREATE TABLE IF NOT EXISTS operations_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      op_date TEXT NOT NULL DEFAULT (date('now')),
      tonnes_moved REAL,
      product_type TEXT,
      product_produced REAL,
      trucks_loaded INTEGER,
      weighbridge_amount REAL,
      trucks_delivered INTEGER,
      product_delivered REAL,
      client_delivered_to TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  function costDefaults() {
    const rows = db.prepare(`
      SELECT key, value
      FROM cost_settings
      WHERE key IN (
        'fuel_cost_per_liter_default',
        'lube_cost_per_qty_default',
        'labor_cost_per_hour_default',
        'downtime_cost_per_hour_default'
      )
    `).all();
    const d = {
      fuel_cost_per_liter_default: 1.5,
      lube_cost_per_qty_default: 4.0,
      labor_cost_per_hour_default: 35.0,
      downtime_cost_per_hour_default: 120.0,
    };
    for (const r of rows) {
      const k = String(r.key || "").trim();
      const v = Number(r.value);
      if (k && Number.isFinite(v)) d[k] = v;
    }
    return d;
  }

  // =========================
  // WORK ORDER PDF (manual form + sign-off)
  // =========================
  app.get("/workorder/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    const download = String(req.query?.download || "").trim() === "1";
    const minimal = String(req.query?.minimal || "").trim() === "1";
    const nohf = String(req.query?.nohf || "").trim() === "1";
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).send({ error: "valid work order id required" });
    }

    const wo = db.prepare(`
      SELECT
        w.id,
        w.asset_id,
        w.source,
        w.reference_id,
        w.status,
        w.opened_at,
        w.closed_at,
        a.asset_code,
        a.asset_name,
        a.category
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE w.id = ?
    `).get(id);

    if (!wo) return reply.code(404).send({ error: "work order not found" });

    let breakdown = null;
    if (String(wo.source || "").toLowerCase() === "breakdown" && wo.reference_id) {
      breakdown = db.prepare(`
        SELECT id, breakdown_date, description, component, critical, downtime_total_hours
        FROM breakdowns
        WHERE id = ?
      `).get(wo.reference_id);
    }

    let servicePlan = null;
    if (String(wo.source || "").toLowerCase() === "service" && wo.reference_id) {
      servicePlan = db.prepare(`
        SELECT id, service_name, interval_hours, last_service_hours, active
        FROM maintenance_plans
        WHERE id = ?
      `).get(wo.reference_id);
    }

    const stockMovementCols = db.prepare(`
      PRAGMA table_info(stock_movements)
    `).all();
    const hasCreatedAt = stockMovementCols.some((c) => String(c.name) === "created_at");
    const movementDateExpr = hasCreatedAt ? "sm.created_at" : "sm.movement_date";

    const issuedParts = db.prepare(`
      SELECT
        sm.id,
        ${movementDateExpr} AS movement_date,
        sm.quantity,
        p.part_code,
        p.part_name
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      WHERE sm.reference = ?
      ORDER BY sm.id ASC
    `).all(`work_order:${id}`);

    const logoPath = path.join(process.cwd(), "branding", "logo.png");

    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Work Order");
        kvGrid(doc, [
          { k: "WO #", v: wo.id },
          { k: "Status", v: String(wo.status || "").toUpperCase() },
          { k: "Source", v: String(wo.source || "").toUpperCase() },
          { k: "Reference ID", v: wo.reference_id ?? "-" },
          { k: "Asset Code", v: wo.asset_code || "" },
          { k: "Asset Name", v: wo.asset_name || "" },
          { k: "Category", v: wo.category || "" },
          { k: "Opened At", v: wo.opened_at || "" },
          { k: "Closed At", v: wo.closed_at || "-" },
        ], 2);

        if (minimal) return;

        if (breakdown) {
          sectionTitle(doc, "Linked Breakdown");
          kvGrid(doc, [
            { k: "Breakdown #", v: breakdown.id },
            { k: "Date", v: breakdown.breakdown_date || "" },
            { k: "Component", v: breakdown.component || "General" },
            { k: "Critical", v: yn(breakdown.critical) },
            { k: "Downtime (hrs)", v: fmtNum(breakdown.downtime_total_hours, 1) },
            { k: "Description", v: breakdown.description || "" },
          ], 2);
        }

        if (servicePlan) {
          sectionTitle(doc, "Linked Service Plan");
          kvGrid(doc, [
            { k: "Plan #", v: servicePlan.id },
            { k: "Service", v: servicePlan.service_name || "" },
            { k: "Interval (hrs)", v: fmtNum(servicePlan.interval_hours, 1) },
            { k: "Last Service (hrs)", v: fmtNum(servicePlan.last_service_hours, 1) },
            { k: "Active", v: yn(servicePlan.active) },
          ], 2);
        }

        sectionTitle(doc, "Issued Parts");
        table(
          doc,
          [
            { key: "date", label: "Date", width: 0.24 },
            { key: "part_code", label: "Part Code", width: 0.16 },
            { key: "part_name", label: "Part Name", width: 0.42 },
            { key: "qty", label: "Qty", width: 0.18, align: "right" },
          ],
          issuedParts.length
            ? issuedParts.map((p) => ({
                date: p.movement_date || "",
                part_code: p.part_code || "",
                part_name: p.part_name || "",
                qty: fmtNum(Math.abs(Number(p.quantity || 0)), 0),
              }))
            : [{ date: "-", part_code: "-", part_name: "No parts issued", qty: "-" }]
        );

        sectionTitle(doc, "Parts & Lubes Required (Manual)");
        doc
          .font("Helvetica")
          .fontSize(9)
          .fillColor("#333333")
          .text(
            "Use this section to request parts/lubes from Stores for the service/repair, even if nothing has been issued yet.",
            { width: 520 }
          );
        doc.moveDown(0.4);

        // Draw a fixed-height manual grid (avoids table auto page-break quirks)
        const left2 = doc.page.margins.left;
        const right2 = doc.page.width - doc.page.margins.right;
        const w2 = right2 - left2;
        const headerH2 = 18;
        const rowH2 = 18;
        const rows2 = 8;
        const gridH2 = headerH2 + rows2 * rowH2;

        ensurePageSpace(doc, gridH2 + 14);

        const yTop = doc.y;
        const cols = [
          { label: "Type", w: 0.12 },
          { label: "Code", w: 0.18 },
          { label: "Description", w: 0.34 },
          { label: "Qty", w: 0.10, align: "right" },
          { label: "Issued By", w: 0.14 },
          { label: "Date", w: 0.12 },
        ];
        const abs = cols.map((c) => Math.floor(c.w * w2));
        const used2 = abs.slice(0, -1).reduce((a, b) => a + b, 0);
        abs[abs.length - 1] = Math.max(60, w2 - used2);

        // Header background
        doc.save();
        doc.rect(left2, yTop, w2, headerH2).fillOpacity(0.06).fill("#000000");
        doc.restore();

        // Outer border
        doc.save();
        doc.rect(left2, yTop, w2, gridH2).lineWidth(1).strokeOpacity(0.18).stroke("#000000");
        doc.restore();

        // Vertical lines + header labels
        let x2 = left2;
        doc.font("Helvetica-Bold").fontSize(9).fillColor("#111111");
        for (let i = 0; i < cols.length; i++) {
          const cw = abs[i];
          const padX = 4;
          const align = cols[i].align || "left";
          doc.text(cols[i].label, x2 + padX, yTop + 5, { width: cw - padX * 2, align });
          if (i > 0) {
            doc
              .moveTo(x2, yTop)
              .lineTo(x2, yTop + gridH2)
              .lineWidth(1)
              .strokeOpacity(0.12)
              .stroke("#000000");
          }
          x2 += cw;
        }

        // Horizontal row lines
        for (let r = 0; r <= rows2; r++) {
          const yy = yTop + headerH2 + r * rowH2;
          doc
            .moveTo(left2, yy)
            .lineTo(right2, yy)
            .lineWidth(1)
            .strokeOpacity(0.08)
            .stroke("#000000");
        }

        doc.y = yTop + gridH2 + 10;

        sectionTitle(doc, "Recorded Completion");
        kvGrid(doc, [
          { k: "Completed At", v: wo.completed_at || wo.closed_at || "-" },
          { k: "Artisan", v: wo.artisan_name || "-" },
          { k: "Artisan Signed At", v: wo.artisan_signed_at || "-" },
          { k: "Supervisor", v: wo.supervisor_name || "-" },
          { k: "Supervisor Signed At", v: wo.supervisor_signed_at || "-" },
          {
            k: "Completion Notes",
            v: wo.completion_notes
              ? compactCell(String(wo.completion_notes || "").replace(/\s+/g, " ").trim(), 220)
              : "-",
          },
        ], 2);

        sectionTitle(doc, "Manual Work Execution (Artisan)");
        doc.font("Helvetica").fontSize(10);
        doc.text("Job Description / Findings:", { width: 460 });
        doc.moveDown(0.2);
        for (let i = 0; i < 4; i++) {
          const y = doc.y + 12;
          doc.moveTo(doc.page.margins.left, y).lineTo(doc.page.width - doc.page.margins.right, y).strokeOpacity(0.2).stroke("#000000");
          doc.moveDown(0.8);
        }

        doc.moveDown(0.4);
        doc.text("Actions Performed:", { width: 460 });
        doc.moveDown(0.2);
        for (let i = 0; i < 4; i++) {
          const y = doc.y + 12;
          doc.moveTo(doc.page.margins.left, y).lineTo(doc.page.width - doc.page.margins.right, y).strokeOpacity(0.2).stroke("#000000");
          doc.moveDown(0.8);
        }

        sectionTitle(doc, "Sign-Off");
        const left = doc.page.margins.left;
        const right = doc.page.width - doc.page.margins.right;
        const mid = left + (right - left) / 2;
        const y0 = doc.y + 6;

        doc.font("Helvetica-Bold").fontSize(10);
        doc.text("Artisan Signature", left, y0, { width: (right - left) / 2 - 20 });
        doc.text("Supervisor Signature", mid + 20, y0, { width: (right - left) / 2 - 20 });

        const lineY = y0 + 30;
        doc.moveTo(left, lineY).lineTo(mid - 20, lineY).strokeOpacity(0.35).stroke("#000000");
        doc.moveTo(mid + 20, lineY).lineTo(right, lineY).strokeOpacity(0.35).stroke("#000000");

        doc.font("Helvetica").fontSize(9);
        doc.text("Name:", left, lineY + 8);
        doc.text("Date:", left + 150, lineY + 8);
        doc.text("Name:", mid + 20, lineY + 8);
        doc.text("Date:", mid + 170, lineY + 8);
      },
      {
        title: "IRONLOG",
        subtitle: "Work Order Job Card",
        rightText: `WO #${wo.id}`,
        showPageNumbers: true,
        disableHeaderFooter: nohf,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="IRONLOG_WO_${wo.id}.pdf"`
      )
      .send(pdf);
  });

  // =========================
  // ASSET HISTORY PDF
  // =========================
  // GET /api/reports/asset-history/:asset_code.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&download=1
  app.get("/asset-history/:asset_code.pdf", async (req, reply) => {
    const asset_code = String(req.params?.asset_code || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const download = String(req.query?.download || "").trim() === "1";

    if (!asset_code) return reply.code(400).send({ error: "asset_code is required" });

    const asset = db.prepare(`
      SELECT id, asset_code, asset_name, category
      FROM assets
      WHERE asset_code = ?
    `).get(asset_code);
    if (!asset) return reply.code(404).send({ error: "asset not found" });

    const startOk = start && isDate(start);
    const endOk = end && isDate(end);
    const periodText = `${startOk ? start : "beginning"} to ${endOk ? end : "today"}`;

    const dateFilter = (col) => {
      const clauses = [];
      const params = [];
      if (startOk) { clauses.push(`${col} >= ?`); params.push(start); }
      if (endOk) { clauses.push(`${col} <= ?`); params.push(end); }
      return { sql: clauses.length ? ` AND ${clauses.join(" AND ")}` : "", params };
    };

    const hasTable = (name) => {
      const row = db.prepare(`
        SELECT name
        FROM sqlite_master
        WHERE type = 'table' AND name = ?
      `).get(name);
      return Boolean(row);
    };

    const bdF = dateFilter("b.breakdown_date");
    const breakdowns = db.prepare(`
      SELECT
        b.breakdown_date AS date,
        b.description,
        b.component,
        b.critical,
        COALESCE(b.downtime_total_hours, 0) AS downtime_hours
      FROM breakdowns b
      WHERE b.asset_id = ? ${bdF.sql}
      ORDER BY b.breakdown_date DESC, b.id DESC
      LIMIT 300
    `).all(asset.id, ...bdF.params).map((r) => ({
      date: r.date,
      description: r.description || "",
      component: r.component || "General",
      critical: Boolean(r.critical),
      downtime_hours: Number(r.downtime_hours || 0),
    }));

    const woF = dateFilter("DATE(w.opened_at)");
    const workOrders = db.prepare(`
      SELECT
        w.id,
        DATE(w.opened_at) AS date,
        w.source,
        w.status,
        w.opened_at,
        w.closed_at
      FROM work_orders w
      WHERE w.asset_id = ?
        AND w.closed_at IS NULL
        AND REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
        AND (w.completed_at IS NULL OR TRIM(COALESCE(w.completed_at, '')) = '')
        ${woF.sql}
      ORDER BY w.id DESC
      LIMIT 300
    `).all(asset.id, ...woF.params);

    const getF = dateFilter("g.slip_date");
    const getSlips = hasTable("get_change_slips")
      ? db.prepare(`
          SELECT
            g.id,
            g.slip_date AS date,
            g.location,
            g.notes
          FROM get_change_slips g
          WHERE g.asset_id = ? ${getF.sql}
          ORDER BY g.slip_date DESC, g.id DESC
          LIMIT 300
        `).all(asset.id, ...getF.params)
      : [];

    const compF = dateFilter("c.slip_date");
    const componentSlips = hasTable("component_change_slips")
      ? db.prepare(`
          SELECT
            c.id,
            c.slip_date AS date,
            c.component,
            c.serial_out,
            c.serial_in,
            c.hours_at_change
          FROM component_change_slips c
          WHERE c.asset_id = ? ${compF.sql}
          ORDER BY c.slip_date DESC, c.id DESC
          LIMIT 300
        `).all(asset.id, ...compF.params)
      : [];

    const breakdownsPdf = breakdowns.slice(0, 60);
    const workOrdersPdf = workOrders.slice(0, 60);
    const getSlipsPdf = getSlips.slice(0, 60);
    const componentSlipsPdf = componentSlips.slice(0, 60);

    const oilF = dateFilter("o.log_date");
    const oil = db.prepare(`
      SELECT IFNULL(SUM(o.quantity), 0) AS oil_qty
      FROM oil_logs o
      WHERE o.asset_id = ? ${oilF.sql}
    `).get(asset.id, ...oilF.params);

    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateCol = hasCreatedAt ? "DATE(sm.created_at)" : "DATE('now')";
    const smF = dateFilter(smDateCol);
    const partsUsed = db.prepare(`
      SELECT IFNULL(SUM(ABS(sm.quantity)), 0) AS qty
      FROM stock_movements sm
      JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
      WHERE w.asset_id = ?
        AND sm.movement_type = 'out'
        ${smF.sql}
    `).get(asset.id, ...smF.params);

    const defaults = costDefaults();
    const fuelF = dateFilter("fl.log_date");
    const fuelCost = db.prepare(`
      SELECT COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS value
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.asset_id = ? ${fuelF.sql}
    `).get(defaults.fuel_cost_per_liter_default, asset.id, ...fuelF.params);

    const oilCost = db.prepare(`
      SELECT COALESCE(SUM(o.quantity * COALESCE(o.unit_cost, ?)), 0) AS value
      FROM oil_logs o
      WHERE o.asset_id = ? ${oilF.sql}
    `).get(defaults.lube_cost_per_qty_default, asset.id, ...oilF.params);

    const partsCost = db.prepare(`
      SELECT COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS value
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
      WHERE w.asset_id = ?
        AND sm.movement_type = 'out'
        ${smF.sql}
    `).get(asset.id, ...smF.params);

    const woCost = db.prepare(`
      SELECT
        COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS labor_hours,
        COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      WHERE w.asset_id = ?
        AND DATE(COALESCE(w.completed_at, w.closed_at, w.opened_at))
            ${startOk ? ">= ?" : ">= DATE('1900-01-01')"}
        AND DATE(COALESCE(w.completed_at, w.closed_at, w.opened_at))
            ${endOk ? "<= ?" : "<= DATE('now')"}
    `).get(
      defaults.labor_cost_per_hour_default,
      asset.id,
      ...(startOk ? [start] : []),
      ...(endOk ? [end] : [])
    );

    const downtimeCost = db.prepare(`
      SELECT COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS value
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE b.asset_id = ?
        AND l.log_date ${startOk ? ">= ?" : ">= DATE('1900-01-01')"}
        AND l.log_date ${endOk ? "<= ?" : "<= DATE('now')"}
    `).get(
      defaults.downtime_cost_per_hour_default,
      asset.id,
      ...(startOk ? [start] : []),
      ...(endOk ? [end] : [])
    );

    const totalCost = Number(
      (
        Number(fuelCost?.value || 0) +
        Number(oilCost?.value || 0) +
        Number(partsCost?.value || 0) +
        Number(woCost?.labor_cost || 0) +
        Number(downtimeCost?.value || 0)
      ).toFixed(2)
    );

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Asset");
        kvGrid(doc, [
          { k: "Asset Code", v: asset.asset_code || "" },
          { k: "Asset Name", v: asset.asset_name || "" },
          { k: "Category", v: asset.category || "" },
          { k: "Period", v: periodText },
        ], 2);

        sectionTitle(doc, "Period Summary");
        kvGrid(doc, [
          { k: "Breakdowns", v: breakdowns.length },
          { k: "Work Orders", v: workOrders.length },
          { k: "GET Slips", v: getSlips.length },
          { k: "Component Slips", v: componentSlips.length },
          { k: "Total Downtime (hrs)", v: fmtNum(breakdowns.reduce((a, r) => a + Number(r.downtime_hours || 0), 0), 1) },
          { k: "Parts Used (qty)", v: fmtNum(partsUsed?.qty || 0, 0) },
          { k: "Oil Used (qty)", v: fmtNum(oil?.oil_qty || 0, 1) },
        ], 2);

        sectionTitle(doc, "Cost Summary");
        kvGrid(doc, [
          { k: "Fuel Cost", v: fmtNum(fuelCost?.value || 0, 2) },
          { k: "Oil/Lube Cost", v: fmtNum(oilCost?.value || 0, 2) },
          { k: "Parts Cost", v: fmtNum(partsCost?.value || 0, 2) },
          { k: "Labor Cost", v: fmtNum(woCost?.labor_cost || 0, 2) },
          { k: "Labor Hours", v: fmtNum(woCost?.labor_hours || 0, 1) },
          { k: "Downtime Cost", v: fmtNum(downtimeCost?.value || 0, 2) },
          { k: "Total Maintenance Cost", v: fmtNum(totalCost, 2) },
        ], 2);

        sectionTitle(doc, "Breakdowns");
        table(
          doc,
          [
            { key: "date", label: "Date", width: 0.16 },
            { key: "component", label: "Component", width: 0.18 },
            { key: "downtime", label: "Downtime", width: 0.12, align: "right" },
            { key: "critical", label: "Critical", width: 0.10, align: "center" },
            { key: "description", label: "Description", width: 0.44 },
          ],
          breakdownsPdf.length
            ? breakdownsPdf.map((r) => ({
                date: r.date,
                component: r.component,
                downtime: fmtNum(r.downtime_hours, 1),
                critical: r.critical ? "YES" : "NO",
                description: compactCell(r.description, 180),
              }))
            : [{ date: "-", component: "-", downtime: "-", critical: "-", description: "No breakdowns in range" }]
        );

        sectionTitle(doc, "Work Orders");
        table(
          doc,
          [
            { key: "id", label: "WO#", width: 0.12, align: "right" },
            { key: "date", label: "Date", width: 0.16 },
            { key: "source", label: "Source", width: 0.16 },
            { key: "status", label: "Status", width: 0.16 },
            { key: "opened", label: "Opened", width: 0.40 },
          ],
          workOrdersPdf.length
            ? workOrdersPdf.map((r) => ({
                id: String(r.id),
                date: r.date || "",
                source: String(r.source || ""),
                status: String(r.status || ""),
                opened: r.opened_at || "",
              }))
            : [{ id: "-", date: "-", source: "-", status: "-", opened: "-" }]
        );

        sectionTitle(doc, "GET Change Slips");
        table(
          doc,
          [
            { key: "id", label: "Slip#", width: 0.14, align: "right" },
            { key: "date", label: "Date", width: 0.18 },
            { key: "location", label: "Location", width: 0.28 },
            { key: "notes", label: "Notes", width: 0.40 },
          ],
          getSlipsPdf.length
            ? getSlipsPdf.map((r) => ({
                id: String(r.id),
                date: r.date || "",
                location: r.location || "-",
                notes: compactCell(r.notes, 140),
              }))
            : [{ id: "-", date: "-", location: "-", notes: "No GET slips in range" }]
        );

        sectionTitle(doc, "Component Change Slips");
        table(
          doc,
          [
            { key: "id", label: "Slip#", width: 0.12, align: "right" },
            { key: "date", label: "Date", width: 0.16 },
            { key: "component", label: "Component", width: 0.24 },
            { key: "serial", label: "Serial Out -> In", width: 0.34 },
            { key: "hours", label: "Hours", width: 0.14, align: "right" },
          ],
          componentSlipsPdf.length
            ? componentSlipsPdf.map((r) => ({
                id: String(r.id),
                date: r.date || "",
                component: r.component || "",
                serial: `${r.serial_out || "-"} -> ${r.serial_in || "-"}`,
                hours: r.hours_at_change == null ? "-" : fmtNum(r.hours_at_change, 1),
              }))
            : [{ id: "-", date: "-", component: "-", serial: "-", hours: "-" }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Asset History Report",
        rightText: `${asset.asset_code} | ${periodText}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="IRONLOG_AssetHistory_${asset.asset_code}.pdf"`
      )
      .send(pdf);
  });

  // =========================
  // LUBE USAGE PDF
  // =========================
  // GET /api/reports/lube.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&download=1
  app.get("/lube.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const download = String(req.query?.download || "").trim() === "1";
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const rows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COUNT(*) AS entries
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.id
      ORDER BY qty_total DESC, a.asset_code ASC
      LIMIT 400
    `).all(start, end).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      qty_total: Number(r.qty_total || 0),
      entries: Number(r.entries || 0),
    }));

    const summary = db.prepare(`
      SELECT
        COALESCE(SUM(quantity), 0) AS qty_total,
        COUNT(*) AS entries,
        COUNT(DISTINCT asset_id) AS assets
      FROM oil_logs
      WHERE log_date BETWEEN ? AND ?
    `).get(start, end);

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Lube Usage Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Total Qty", v: fmtNum(summary?.qty_total || 0, 1) },
          { k: "Total Entries", v: fmtNum(summary?.entries || 0, 0) },
          { k: "Assets Logged", v: fmtNum(summary?.assets || 0, 0) },
        ], 2);

        sectionTitle(doc, "Lube Usage by Machine");
        table(
          doc,
          [
            { key: "asset_code", label: "Asset Code", width: 0.18 },
            { key: "asset_name", label: "Asset Name", width: 0.46 },
            { key: "entries", label: "Entries", width: 0.14, align: "right" },
            { key: "qty_total", label: "Qty Total", width: 0.22, align: "right" },
          ],
          rows.length
            ? rows.map((r) => ({
                asset_code: r.asset_code,
                asset_name: r.asset_name || "",
                entries: fmtNum(r.entries, 0),
                qty_total: fmtNum(r.qty_total, 1),
              }))
            : [{ asset_code: "-", asset_name: "No lube usage in period", entries: "-", qty_total: "-" }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Lube Usage Report",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="IRONLOG_Lube_${start}_to_${end}.pdf"`
      )
      .send(pdf);
  });

  // =========================
  // FUEL BENCHMARK PDF
  // =========================
  // GET /api/reports/fuel-benchmark.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15&download=1
  app.get("/fuel-benchmark.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const toleranceInput = Number(req.query?.tolerance ?? 0.15);
    const tolerance = Number.isFinite(toleranceInput) ? Math.max(0, toleranceInput) : 0.15;
    const download = String(req.query?.download || "").trim() === "1";
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const fuelByAsset = db.prepare(`
      SELECT
        a.id AS asset_id,
        a.asset_code,
        a.asset_name,
        COALESCE(a.baseline_fuel_l_per_hour, 5.0) AS oem_lph,
        COALESCE(SUM(fl.liters), 0) AS fuel_liters
      FROM assets a
      LEFT JOIN fuel_logs fl
        ON fl.asset_id = a.id
       AND fl.log_date BETWEEN ? AND ?
      WHERE a.active = 1
      GROUP BY a.id
      ORDER BY a.asset_code ASC
    `).all(start, end);

    const getHoursDaily = db.prepare(`
      SELECT COALESCE(SUM(hours_run), 0) AS hours_run
      FROM daily_hours
      WHERE asset_id = ?
        AND work_date BETWEEN ? AND ?
        AND is_used = 1
        AND hours_run > 0
    `);
    const getHoursFuel = db.prepare(`
      SELECT COALESCE(SUM(hours_run), 0) AS hours_run
      FROM fuel_logs
      WHERE asset_id = ?
        AND log_date BETWEEN ? AND ?
        AND hours_run > 0
    `);

    const rows = fuelByAsset.map((r) => {
      const dailyHours = Number(getHoursDaily.get(r.asset_id, start, end)?.hours_run || 0);
      const fuelHours = Number(getHoursFuel.get(r.asset_id, start, end)?.hours_run || 0);
      const hours = dailyHours > 0 ? dailyHours : fuelHours;
      const fuel = Number(r.fuel_liters || 0);
      const oem = Number(r.oem_lph || 5);
      const lph = hours > 0 ? fuel / hours : null;
      const excessiveThreshold = oem * (1 + tolerance);
      const is_excessive = lph != null && lph > excessiveThreshold;
      return {
        asset_code: r.asset_code,
        asset_name: r.asset_name,
        fuel_liters: Number(fuel.toFixed(2)),
        hours_run: Number(hours.toFixed(2)),
        actual_lph: lph == null ? null : Number(lph.toFixed(3)),
        oem_lph: Number(oem.toFixed(3)),
        threshold_lph: Number(excessiveThreshold.toFixed(3)),
        variance_lph: lph == null ? null : Number((lph - oem).toFixed(3)),
        flag: is_excessive ? "EXCESSIVE" : "OK",
      };
    }).filter((r) => r.fuel_liters > 0 || r.hours_run > 0)
      .sort((a, b) => {
        const ex = (b.flag === "EXCESSIVE" ? 1 : 0) - (a.flag === "EXCESSIVE" ? 1 : 0);
        if (ex !== 0) return ex;
        return Number(b.variance_lph || -999) - Number(a.variance_lph || -999);
      });

    const summary = rows.reduce(
      (acc, r) => {
        acc.assets += 1;
        acc.fuel_liters += Number(r.fuel_liters || 0);
        acc.hours_run += Number(r.hours_run || 0);
        if (r.flag === "EXCESSIVE") acc.excessive += 1;
        return acc;
      },
      { assets: 0, fuel_liters: 0, hours_run: 0, excessive: 0 }
    );
    summary.fuel_liters = Number(summary.fuel_liters.toFixed(2));
    summary.hours_run = Number(summary.hours_run.toFixed(2));
    summary.avg_lph = summary.hours_run > 0
      ? Number((summary.fuel_liters / summary.hours_run).toFixed(3))
      : null;

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Fuel Benchmark Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Tolerance", v: `${fmtNum(tolerance * 100, 1)}% above OEM` },
          { k: "Assets", v: fmtNum(summary.assets, 0) },
          { k: "Excessive", v: fmtNum(summary.excessive, 0) },
          { k: "Fuel Total (L)", v: fmtNum(summary.fuel_liters, 2) },
          { k: "Hours Run", v: fmtNum(summary.hours_run, 2) },
          { k: "Avg L/hr", v: summary.avg_lph == null ? "-" : fmtNum(summary.avg_lph, 3) },
        ], 2);

        sectionTitle(doc, "Fuel Benchmark by Machine");
        table(
          doc,
          [
            { key: "asset_code", label: "Asset", width: 0.12 },
            { key: "asset_name", label: "Name", width: 0.24 },
            { key: "fuel_liters", label: "Fuel (L)", width: 0.12, align: "right" },
            { key: "hours_run", label: "Hours", width: 0.10, align: "right" },
            { key: "actual_lph", label: "Actual L/hr", width: 0.14, align: "right" },
            { key: "oem_lph", label: "OEM L/hr", width: 0.10, align: "right" },
            { key: "variance_lph", label: "Variance", width: 0.10, align: "right" },
            { key: "flag", label: "Flag", width: 0.08, align: "center" },
          ],
          rows.length
            ? rows.map((r) => ({
                asset_code: r.asset_code,
                asset_name: r.asset_name || "",
                fuel_liters: fmtNum(r.fuel_liters, 2),
                hours_run: fmtNum(r.hours_run, 2),
                actual_lph: r.actual_lph == null ? "-" : fmtNum(r.actual_lph, 3),
                oem_lph: fmtNum(r.oem_lph, 3),
                variance_lph: r.variance_lph == null ? "-" : fmtNum(r.variance_lph, 3),
                flag: r.flag,
              }))
            : [{
                asset_code: "-",
                asset_name: "No fuel benchmark data for period",
                fuel_liters: "-",
                hours_run: "-",
                actual_lph: "-",
                oem_lph: "-",
                variance_lph: "-",
                flag: "-",
              }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Fuel Benchmark Report",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="IRONLOG_Fuel_Benchmark_${start}_to_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/fuel-machine-history.pdf?asset_code=A300AM&start=YYYY-MM-DD&end=YYYY-MM-DD&tolerance=0.15&download=1
  app.get("/fuel-machine-history.pdf", async (req, reply) => {
    const assetCode = String(req.query?.asset_code || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const toleranceInput = Number(req.query?.tolerance ?? 0.15);
    const tolerance = Number.isFinite(toleranceInput) ? Math.max(0, toleranceInput) : 0.15;
    const download = String(req.query?.download || "").trim() === "1";

    if (!assetCode) return reply.code(400).send({ error: "asset_code is required" });
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const asset = db.prepare(`
      SELECT id, asset_code, asset_name, COALESCE(baseline_fuel_l_per_hour, 5.0) AS oem_lph
      FROM assets
      WHERE asset_code = ?
    `).get(assetCode);
    if (!asset) return reply.code(404).send({ error: `asset not found: ${assetCode}` });

    const fuelRows = db.prepare(`
      SELECT
        fl.id,
        fl.log_date,
        COALESCE(fl.liters, 0) AS fuel_liters,
        COALESCE(CASE WHEN fl.hours_run > 0 THEN fl.hours_run ELSE 0 END, 0) AS meter_hours,
        fl.source
      FROM fuel_logs fl
      WHERE fl.asset_id = ?
        AND fl.log_date BETWEEN ? AND ?
      ORDER BY fl.log_date ASC, fl.id ASC
    `).all(asset.id, start, end);

    const previousFill = db.prepare(`
      SELECT fl.hours_run
      FROM fuel_logs fl
      WHERE fl.asset_id = ?
        AND fl.log_date < ?
        AND fl.hours_run > 0
      ORDER BY fl.log_date DESC, fl.id DESC
      LIMIT 1
    `).get(asset.id, start);

    const oem = Number(asset.oem_lph || 5);
    const threshold = oem * (1 + tolerance);
    let prevMeter = previousFill && Number(previousFill.hours_run) > 0
      ? Number(previousFill.hours_run)
      : null;

    const rows = fuelRows.map((d) => {
      const meter = Number(d.meter_hours || 0);
      let hoursBetween = null;
      if (prevMeter != null && meter > 0) {
        const delta = meter - prevMeter;
        if (Number.isFinite(delta) && delta > 0) hoursBetween = delta;
      }
      const fuel = Number(d.fuel_liters || 0);
      const lph = hoursBetween != null && hoursBetween > 0 ? fuel / hoursBetween : null;
      const flag = lph != null && lph > threshold ? "EXCESSIVE" : "OK";
      if (meter > 0) prevMeter = meter;
      return {
        log_date: d.log_date,
        fuel_liters: Number(fuel.toFixed(2)),
        hours_between: hoursBetween == null ? 0 : Number(hoursBetween.toFixed(2)),
        actual_lph: lph == null ? null : Number(lph.toFixed(3)),
        oem_lph: Number(oem.toFixed(3)),
        flag,
        source: d.source || "",
      };
    });

    const summary = rows.reduce(
      (acc, r) => {
        acc.fill_days += 1;
        acc.fuel_liters += Number(r.fuel_liters || 0);
        acc.hours_between += Number(r.hours_between || 0);
        if (r.flag === "EXCESSIVE") acc.excessive_days += 1;
        return acc;
      },
      { fill_days: 0, fuel_liters: 0, hours_between: 0, excessive_days: 0 }
    );
    summary.fuel_liters = Number(summary.fuel_liters.toFixed(2));
    summary.hours_between = Number(summary.hours_between.toFixed(2));
    summary.avg_lph = summary.hours_between > 0
      ? Number((summary.fuel_liters / summary.hours_between).toFixed(3))
      : null;

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Machine Fuel Fill History");
        kvGrid(doc, [
          { k: "Asset", v: `${asset.asset_code} ${asset.asset_name ? `- ${asset.asset_name}` : ""}` },
          { k: "Period", v: `${start} to ${end}` },
          { k: "Tolerance", v: `${fmtNum(tolerance * 100, 1)}% above OEM` },
          { k: "OEM L/hr", v: fmtNum(oem, 3) },
          { k: "Fill Days", v: fmtNum(summary.fill_days, 0) },
          { k: "Excessive Days", v: fmtNum(summary.excessive_days, 0) },
          { k: "Fuel Total (L)", v: fmtNum(summary.fuel_liters, 2) },
          { k: "Hours Between Fills", v: fmtNum(summary.hours_between, 2) },
          { k: "Avg L/hr", v: summary.avg_lph == null ? "-" : fmtNum(summary.avg_lph, 3) },
        ], 2);

        sectionTitle(doc, "Fill Entries");
        table(
          doc,
          [
            { key: "log_date", label: "Date", width: 0.16 },
            { key: "fuel_liters", label: "Fuel (L)", width: 0.14, align: "right" },
            { key: "hours_between", label: "Hours Between", width: 0.16, align: "right" },
            { key: "actual_lph", label: "L/hr", width: 0.12, align: "right" },
            { key: "oem_lph", label: "OEM L/hr", width: 0.12, align: "right" },
            { key: "flag", label: "Status", width: 0.10, align: "center" },
            { key: "source", label: "Source", width: 0.20 },
          ],
          rows.length
            ? rows.map((r) => ({
                log_date: r.log_date,
                fuel_liters: fmtNum(r.fuel_liters, 2),
                hours_between: fmtNum(r.hours_between, 2),
                actual_lph: r.actual_lph == null ? "-" : fmtNum(r.actual_lph, 3),
                oem_lph: fmtNum(r.oem_lph, 3),
                flag: r.flag,
                source: r.source || "",
              }))
            : [{
                log_date: "-",
                fuel_liters: "-",
                hours_between: "-",
                actual_lph: "-",
                oem_lph: "-",
                flag: "-",
                source: "No fill data in selected period",
              }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Fuel Machine History",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="IRONLOG_Fuel_History_${asset.asset_code}_${start}_to_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/manager-inspection/:id.pdf?download=1
  app.get("/manager-inspection/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    const download = String(req.query?.download || "").trim() === "1";
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).send({ error: "valid inspection id required" });
    }

    const miInspectorCol = pickExistingColumn("manager_inspections", ["inspector_name", "inspector"], "inspector_name");
    const photoInspectionCol = pickExistingColumn("manager_inspection_photos", ["inspection_id", "manager_inspection_id"], "inspection_id");
    const photoPathCol = pickExistingColumn("manager_inspection_photos", ["file_path", "photo_path", "path", "image_path", "url"], "file_path");
    const photoCaptionCol = pickExistingColumn("manager_inspection_photos", ["caption", "note", "notes", "description"], "caption");
    const photoCreatedCol = pickExistingColumn("manager_inspection_photos", ["created_at", "uploaded_at", "created_on"], "created_at");

    const inspection = db.prepare(`
      SELECT
        mi.id,
        mi.inspection_date,
        mi.${miInspectorCol} AS inspector_name,
        mi.notes,
        mi.created_at,
        COALESCE((
          SELECT MAX(dh.closing_hours)
          FROM daily_hours dh
          WHERE dh.asset_id = mi.asset_id
            AND dh.closing_hours IS NOT NULL
            AND dh.work_date <= mi.inspection_date
        ), 0) AS machine_hours,
        a.asset_code,
        a.asset_name,
        a.category
      FROM manager_inspections mi
      JOIN assets a ON a.id = mi.asset_id
      WHERE mi.id = ?
    `).get(id);
    if (!inspection) return reply.code(404).send({ error: "manager inspection not found" });

    const photos = db.prepare(`
      SELECT id, ${photoPathCol} AS file_path, ${photoCaptionCol} AS caption, ${photoCreatedCol} AS created_at
      FROM manager_inspection_photos
      WHERE ${photoInspectionCol} = ?
      ORDER BY id ASC
    `).all(id);

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Manager Inspection");
        kvGrid(doc, [
          { k: "Inspection #", v: inspection.id },
          { k: "Date", v: inspection.inspection_date || "" },
          { k: "Inspector", v: inspection.inspector_name || "-" },
          { k: "Asset Code", v: inspection.asset_code || "" },
          { k: "Asset Name", v: inspection.asset_name || "" },
          { k: "Machine Hours", v: Number(inspection.machine_hours || 0).toFixed(1) },
          { k: "Category", v: inspection.category || "" },
          { k: "Created At", v: inspection.created_at || "" },
        ], 2);

        sectionTitle(doc, "Notes");
        doc
          .font("Helvetica")
          .fontSize(10)
          .fillColor("#111111")
          .text(compactCell(inspection.notes || "-", 2000), {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });

        sectionTitle(doc, "Photos");
        if (!photos.length) {
          doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No photos attached.");
          return;
        }

        for (const p of photos) {
          const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
          const abs = resolveStorageAbs(rel);
          ensurePageSpace(doc, 230);
          doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
          doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`, {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });
          doc.moveDown(0.2);
          if (abs && fs.existsSync(abs)) {
            try {
              doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 180], align: "left", valign: "top" });
              doc.y += 186;
            } catch {
              doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo file exists but could not be rendered.");
              doc.moveDown(0.5);
            }
          } else {
            doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
            doc.moveDown(0.5);
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: "Manager Inspection Report",
        rightText: `Inspection #${inspection.id}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Manager_Inspection_${inspection.id}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/manager-inspections.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&asset_id=123&with_photos=1&download=1
  app.get("/manager-inspections.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const assetId = Number(req.query?.asset_id || 0);
    const withPhotos = String(req.query?.with_photos || "").trim() === "1";
    const download = String(req.query?.download || "").trim() === "1";

    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const miInspectorCol = pickExistingColumn("manager_inspections", ["inspector_name", "inspector"], "inspector_name");
    const photoInspectionCol = pickExistingColumn("manager_inspection_photos", ["inspection_id", "manager_inspection_id"], "inspection_id");
    const photoPathCol = pickExistingColumn("manager_inspection_photos", ["file_path", "photo_path", "path", "image_path", "url"], "file_path");
    const photoCaptionCol = pickExistingColumn("manager_inspection_photos", ["caption", "note", "notes", "description"], "caption");
    const photoCreatedCol = pickExistingColumn("manager_inspection_photos", ["created_at", "uploaded_at", "created_on"], "created_at");

    const where = ["mi.inspection_date >= ?", "mi.inspection_date <= ?"];
    const params = [start, end];
    if (assetId > 0) {
      where.push("mi.asset_id = ?");
      params.push(assetId);
    }

    const rows = db.prepare(`
      SELECT
        mi.id,
        mi.asset_id,
        mi.inspection_date,
        mi.${miInspectorCol} AS inspector_name,
        mi.notes,
        a.asset_code,
        a.asset_name
      FROM manager_inspections mi
      JOIN assets a ON a.id = mi.asset_id
      WHERE ${where.join(" AND ")}
      ORDER BY mi.inspection_date DESC, mi.id DESC
      LIMIT 1000
    `).all(...params);

    const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
    const photosByInspection = new Map();
    if (ids.length) {
      const marks = ids.map(() => "?").join(",");
      const photos = db.prepare(`
        SELECT ${photoInspectionCol} AS inspection_id, id, ${photoPathCol} AS file_path, ${photoCaptionCol} AS caption, ${photoCreatedCol} AS created_at
        FROM manager_inspection_photos
        WHERE ${photoInspectionCol} IN (${marks})
        ORDER BY ${photoInspectionCol} ASC, id ASC
      `).all(...ids);
      for (const p of photos) {
        const key = Number(p.inspection_id);
        if (!photosByInspection.has(key)) photosByInspection.set(key, []);
        photosByInspection.get(key).push(p);
      }
    }

    const summary = {
      count: rows.length,
      assets: new Set(rows.map((r) => Number(r.asset_id))).size,
      inspectors: new Set(rows.map((r) => String(r.inspector_name || "").trim()).filter(Boolean)).size,
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Manager Inspections Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Asset Filter", v: assetId > 0 ? String(assetId) : "All assets" },
          { k: "Inspections", v: String(summary.count) },
          { k: "Assets Covered", v: String(summary.assets) },
          { k: "Inspectors", v: String(summary.inspectors) },
        ], 2);

        sectionTitle(doc, "Inspection Entries");
        table(
          doc,
          [
            { key: "id", label: "ID", width: 0.08, align: "right" },
            { key: "date", label: "Date", width: 0.13 },
            { key: "asset", label: "Asset", width: 0.18 },
            { key: "name", label: "Asset Name", width: 0.22 },
            { key: "inspector", label: "Inspector", width: 0.14 },
            { key: "notes", label: "Notes", width: 0.25 },
          ],
          rows.length
            ? rows.map((r) => ({
                id: String(r.id),
                date: r.inspection_date || "",
                asset: r.asset_code || "",
                name: r.asset_name || "",
                inspector: r.inspector_name || "-",
                notes: compactCell(r.notes || "", 120),
              }))
            : [{
                id: "-",
                date: "-",
                asset: "-",
                name: "No inspections found in selected period",
                inspector: "-",
                notes: "-",
              }]
        );

        if (withPhotos) {
          sectionTitle(doc, "Inspection Photos");
          if (!rows.length) {
            doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No inspections in selected period.");
          } else {
            for (const r of rows) {
              const photos = photosByInspection.get(Number(r.id)) || [];
              ensurePageSpace(doc, 60);
              doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
              doc.text(
                `Inspection #${r.id} | ${r.inspection_date} | ${r.asset_code}${r.asset_name ? ` - ${r.asset_name}` : ""}`,
                { width: doc.page.width - doc.page.margins.left - doc.page.margins.right }
              );
              doc.moveDown(0.2);
              if (!photos.length) {
                doc.font("Helvetica").fontSize(9).fillColor("#666666").text("No photos attached.");
                doc.moveDown(0.3);
                continue;
              }
              for (const p of photos) {
                const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
                const abs = resolveStorageAbs(rel);
                ensurePageSpace(doc, 220);
                doc.font("Helvetica").fontSize(9).fillColor("#111111");
                doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`);
                doc.moveDown(0.15);
                if (abs && fs.existsSync(abs)) {
                  try {
                    doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 170], align: "left", valign: "top" });
                    doc.y += 176;
                  } catch {
                    doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo exists but could not be rendered.");
                    doc.moveDown(0.3);
                  }
                } else {
                  doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
                  doc.moveDown(0.3);
                }
              }
              doc.moveDown(0.2);
            }
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: withPhotos ? "Manager Inspections Report (With Photos)" : "Manager Inspections Report",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="AML_Manager_Inspections_${withPhotos ? "WithPhotos_" : ""}${start}_to_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/damage-report/:id.pdf?download=1
  app.get("/damage-report/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    const download = String(req.query?.download || "").trim() === "1";
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).send({ error: "valid damage report id required" });
    }

    const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
    const drPhotoReportCol = pickExistingColumn("manager_damage_report_photos", ["damage_report_id", "manager_damage_report_id", "report_id"], "damage_report_id");
    const drPhotoPathCol = pickExistingColumn("manager_damage_report_photos", ["file_path", "photo_path", "path", "image_path", "url", "image_data"], "file_path");
    const drPhotoCaptionCol = pickExistingColumn("manager_damage_report_photos", ["caption", "note", "notes", "description"], "caption");
    const drPhotoCreatedCol = pickExistingColumn("manager_damage_report_photos", ["created_at", "uploaded_at", "created_on"], "created_at");

    const report = db.prepare(`
      SELECT
        dr.id,
        dr.report_date,
        dr.${drInspectorCol} AS inspector_name,
        dr.hour_meter,
        dr.damage_location,
        dr.severity,
        dr.damage_description,
        dr.immediate_action,
        dr.out_of_service,
        dr.damage_time,
        dr.responsible_person,
        dr.pending_investigation,
        dr.hse_report_available,
        dr.created_at,
        a.asset_code,
        a.asset_name,
        a.category
      FROM manager_damage_reports dr
      JOIN assets a ON a.id = dr.asset_id
      WHERE dr.id = ?
    `).get(id);
    if (!report) return reply.code(404).send({ error: "damage report not found" });

    const photos = db.prepare(`
      SELECT id, ${drPhotoPathCol} AS file_path, ${drPhotoCaptionCol} AS caption, ${drPhotoCreatedCol} AS created_at
      FROM manager_damage_report_photos
      WHERE ${drPhotoReportCol} = ?
      ORDER BY id ASC
    `).all(id);

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Damage Report");
        kvGrid(doc, [
          { k: "Report #", v: report.id },
          { k: "Date", v: report.report_date || "" },
          { k: "Inspector", v: report.inspector_name || "-" },
          { k: "Asset Code", v: report.asset_code || "" },
          { k: "Asset Name", v: report.asset_name || "" },
          { k: "Hours", v: report.hour_meter == null ? "-" : Number(report.hour_meter || 0).toFixed(1) },
          { k: "Damage Time", v: report.damage_time || "-" },
          { k: "Location", v: report.damage_location || "-" },
          { k: "Responsible Person", v: report.responsible_person || "-" },
          { k: "Severity", v: String(report.severity || "-").toUpperCase() },
          { k: "Out of Service", v: Number(report.out_of_service || 0) ? "YES" : "NO" },
          { k: "Pending Investigation", v: Number(report.pending_investigation || 0) ? "YES" : "NO" },
          { k: "HSE Report Available", v: Number(report.hse_report_available || 0) ? "YES" : "NO" },
          { k: "Category", v: report.category || "" },
          { k: "Created At", v: report.created_at || "" },
        ], 2);

        sectionTitle(doc, "Damage Description");
        doc.font("Helvetica").fontSize(10).fillColor("#111111").text(compactCell(report.damage_description || "-", 2000), {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });

        sectionTitle(doc, "Immediate Action");
        doc.font("Helvetica").fontSize(10).fillColor("#111111").text(compactCell(report.immediate_action || "-", 2000), {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });

        sectionTitle(doc, "Photos");
        if (!photos.length) {
          doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No photos attached.");
          return;
        }
        for (const p of photos) {
          const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
          const abs = resolveStorageAbs(rel);
          ensurePageSpace(doc, 230);
          doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
          doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`, {
            width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
          });
          doc.moveDown(0.2);
          if (abs && fs.existsSync(abs)) {
            try {
              doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 180], align: "left", valign: "top" });
              doc.y += 186;
            } catch {
              doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo file exists but could not be rendered.");
              doc.moveDown(0.5);
            }
          } else {
            doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
            doc.moveDown(0.5);
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: "Damage Report",
        rightText: `Report #${report.id}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename="AML_Damage_Report_${report.id}.pdf"`)
      .send(pdf);
  });

  // GET /api/reports/damage-reports.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&asset_id=123&with_photos=1&download=1
  app.get("/damage-reports.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const assetId = Number(req.query?.asset_id || 0);
    const withPhotos = String(req.query?.with_photos || "").trim() === "1";
    const download = String(req.query?.download || "").trim() === "1";

    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });
    }

    const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
    const drPhotoReportCol = pickExistingColumn("manager_damage_report_photos", ["damage_report_id", "manager_damage_report_id", "report_id"], "damage_report_id");
    const drPhotoPathCol = pickExistingColumn("manager_damage_report_photos", ["file_path", "photo_path", "path", "image_path", "url", "image_data"], "file_path");
    const drPhotoCaptionCol = pickExistingColumn("manager_damage_report_photos", ["caption", "note", "notes", "description"], "caption");
    const drPhotoCreatedCol = pickExistingColumn("manager_damage_report_photos", ["created_at", "uploaded_at", "created_on"], "created_at");

    const where = ["dr.report_date >= ?", "dr.report_date <= ?"];
    const params = [start, end];
    if (assetId > 0) {
      where.push("dr.asset_id = ?");
      params.push(assetId);
    }

    const rows = db.prepare(`
      SELECT
        dr.id,
        dr.asset_id,
        dr.report_date,
        dr.${drInspectorCol} AS inspector_name,
        dr.hour_meter,
        dr.damage_location,
        dr.severity,
        dr.damage_description,
        dr.immediate_action,
        dr.out_of_service,
        dr.damage_time,
        dr.responsible_person,
        dr.pending_investigation,
        dr.hse_report_available,
        a.asset_code,
        a.asset_name
      FROM manager_damage_reports dr
      JOIN assets a ON a.id = dr.asset_id
      WHERE ${where.join(" AND ")}
      ORDER BY dr.report_date DESC, dr.id DESC
      LIMIT 1000
    `).all(...params);

    const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
    const photosByReport = new Map();
    if (ids.length) {
      const marks = ids.map(() => "?").join(",");
      const photos = db.prepare(`
        SELECT ${drPhotoReportCol} AS damage_report_id, id, ${drPhotoPathCol} AS file_path, ${drPhotoCaptionCol} AS caption, ${drPhotoCreatedCol} AS created_at
        FROM manager_damage_report_photos
        WHERE ${drPhotoReportCol} IN (${marks})
        ORDER BY ${drPhotoReportCol} ASC, id ASC
      `).all(...ids);
      for (const p of photos) {
        const key = Number(p.damage_report_id);
        if (!photosByReport.has(key)) photosByReport.set(key, []);
        photosByReport.get(key).push(p);
      }
    }

    const summary = {
      count: rows.length,
      assets: new Set(rows.map((r) => Number(r.asset_id))).size,
      out_of_service: rows.filter((r) => Number(r.out_of_service || 0) === 1).length,
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Damage Reports Summary");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Asset Filter", v: assetId > 0 ? String(assetId) : "All assets" },
          { k: "Reports", v: String(summary.count) },
          { k: "Assets Covered", v: String(summary.assets) },
          { k: "Out of Service", v: String(summary.out_of_service) },
        ], 2);

        sectionTitle(doc, "Damage Entries");
        table(
          doc,
          [
            { key: "id", label: "ID", width: 0.08, align: "right" },
            { key: "date", label: "Date", width: 0.12 },
            { key: "asset", label: "Asset", width: 0.13 },
            { key: "inspector", label: "Inspector", width: 0.11 },
            { key: "time", label: "Time", width: 0.06, align: "center" },
            { key: "location", label: "Location", width: 0.1 },
            { key: "resp", label: "Responsible", width: 0.1 },
            { key: "severity", label: "Severity", width: 0.08 },
            { key: "hours", label: "Hours", width: 0.08, align: "right" },
            { key: "out", label: "OOS", width: 0.06, align: "center" },
            { key: "inv", label: "Inv", width: 0.05, align: "center" },
            { key: "hse", label: "HSE", width: 0.05, align: "center" },
            { key: "desc", label: "Description", width: 0.08 },
          ],
          rows.length
            ? rows.map((r) => ({
                id: String(r.id),
                date: r.report_date || "",
                asset: r.asset_code || "",
                inspector: r.inspector_name || "-",
                time: r.damage_time || "-",
                location: compactCell(r.damage_location || "", 40),
                resp: compactCell(r.responsible_person || "", 20),
                severity: String(r.severity || "-").toUpperCase(),
                hours: r.hour_meter == null ? "-" : Number(r.hour_meter || 0).toFixed(1),
                out: Number(r.out_of_service || 0) ? "YES" : "NO",
                inv: Number(r.pending_investigation || 0) ? "Y" : "N",
                hse: Number(r.hse_report_available || 0) ? "Y" : "N",
                desc: compactCell(r.damage_description || "", 60),
              }))
            : [{
                id: "-",
                date: "-",
                asset: "-",
                inspector: "-",
                time: "-",
                location: "-",
                resp: "-",
                severity: "-",
                hours: "-",
                out: "-",
                inv: "-",
                hse: "-",
                desc: "No damage reports found in selected period",
              }]
        );

        if (withPhotos) {
          sectionTitle(doc, "Damage Photos");
          if (!rows.length) {
            doc.font("Helvetica").fontSize(10).fillColor("#555555").text("No damage reports in selected period.");
          } else {
            for (const r of rows) {
              const photos = photosByReport.get(Number(r.id)) || [];
              ensurePageSpace(doc, 60);
              doc.font("Helvetica-Bold").fontSize(10).fillColor("#111111");
              doc.text(
                `Report #${r.id} | ${r.report_date} | ${r.asset_code}${r.asset_name ? ` - ${r.asset_name}` : ""} | Severity: ${String(r.severity || "-").toUpperCase()}`,
                { width: doc.page.width - doc.page.margins.left - doc.page.margins.right }
              );
              doc.moveDown(0.2);
              if (!photos.length) {
                doc.font("Helvetica").fontSize(9).fillColor("#666666").text("No photos attached.");
                doc.moveDown(0.3);
                continue;
              }
              for (const p of photos) {
                const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
                const abs = resolveStorageAbs(rel);
                ensurePageSpace(doc, 220);
                doc.font("Helvetica").fontSize(9).fillColor("#111111");
                doc.text(`Photo #${p.id}${p.caption ? ` - ${p.caption}` : ""}`);
                doc.moveDown(0.15);
                if (abs && fs.existsSync(abs)) {
                  try {
                    doc.image(abs, doc.page.margins.left, doc.y, { fit: [420, 170], align: "left", valign: "top" });
                    doc.y += 176;
                  } catch {
                    doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo exists but could not be rendered.");
                    doc.moveDown(0.3);
                  }
                } else {
                  doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text(`Photo missing: ${rel || "-"}`);
                  doc.moveDown(0.3);
                }
              }
              doc.moveDown(0.2);
            }
          }
        }
      },
      {
        title: "IRONLOG",
        subtitle: withPhotos ? "Damage Reports (With Photos)" : "Damage Reports",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename="AML_Damage_Reports_${withPhotos ? "WithPhotos_" : ""}${start}_to_${end}.pdf"`)
      .send(pdf);
  });

  // =========================
  // STOCK MONITOR PDF
  // =========================
  // GET /api/reports/stock-monitor.pdf?part_code=FLT&download=1
  app.get("/stock-monitor.pdf", async (req, reply) => {
    const part_code = String(req.query?.part_code || "").trim();
    const download = String(req.query?.download || "").trim() === "1";

    const where = [];
    const params = [];
    if (part_code) {
      where.push("p.part_code LIKE ?");
      params.push(`%${part_code}%`);
    }

    const rows = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        p.critical,
        p.min_stock,
        IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      GROUP BY p.id
      ORDER BY p.critical DESC, on_hand ASC, p.part_code ASC
      LIMIT 500
    `).all(...params).map((r) => ({
      part_code: r.part_code,
      part_name: r.part_name,
      critical: Boolean(r.critical),
      min_stock: Number(r.min_stock || 0),
      on_hand: Number(r.on_hand || 0),
      below_min: Number(r.on_hand || 0) < Number(r.min_stock || 0),
    }));

    const summary = {
      total_parts: rows.length,
      below_min: rows.filter((r) => r.below_min).length,
      critical_below_min: rows.filter((r) => r.below_min && r.critical).length,
      total_on_hand: Number(rows.reduce((acc, r) => acc + Number(r.on_hand || 0), 0).toFixed(2)),
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Stock Monitor Summary");
        kvGrid(doc, [
          { k: "Filter", v: part_code || "All parts" },
          { k: "Total Parts", v: fmtNum(summary.total_parts, 0) },
          { k: "Below Min", v: fmtNum(summary.below_min, 0) },
          { k: "Critical Below Min", v: fmtNum(summary.critical_below_min, 0) },
          { k: "Total On Hand", v: fmtNum(summary.total_on_hand, 1) },
        ], 2);

        sectionTitle(doc, "Stock Levels");
        table(
          doc,
          [
            { key: "part_code", label: "Part Code", width: 0.16 },
            { key: "part_name", label: "Part Name", width: 0.44 },
            { key: "on_hand", label: "On Hand", width: 0.12, align: "right" },
            { key: "min_stock", label: "Min", width: 0.10, align: "right" },
            { key: "critical", label: "Critical", width: 0.08, align: "center" },
            { key: "below", label: "Below Min", width: 0.10, align: "center" },
          ],
          rows.length
            ? rows.map((r) => ({
                part_code: r.part_code,
                part_name: r.part_name || "",
                on_hand: fmtNum(r.on_hand, 1),
                min_stock: fmtNum(r.min_stock, 1),
                critical: r.critical ? "Y" : "N",
                below: r.below_min ? "YES" : "NO",
              }))
            : [{
                part_code: "-",
                part_name: "No parts for selected filter",
                on_hand: "-",
                min_stock: "-",
                critical: "-",
                below: "-",
              }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Stock Monitor Report",
        rightText: part_code ? `Filter: ${part_code}` : "All parts",
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="IRONLOG_Stock_Monitor${part_code ? `_${part_code}` : ""}.pdf"`
      )
      .send(pdf);
  });

  // =========================
  // LEGAL COMPLIANCE PDF
  // =========================
  // GET /api/reports/legal-compliance.pdf?days=90&department=&status=approved&download=1
  app.get("/legal-compliance.pdf", async (req, reply) => {
    const daysRaw = Number(req.query?.days ?? 90);
    const days = Number.isFinite(daysRaw) ? Math.min(3650, Math.max(1, Math.trunc(daysRaw))) : 90;
    const department = String(req.query?.department || "").trim();
    const status = String(req.query?.status || "approved").trim().toLowerCase();
    const download = String(req.query?.download || "").trim() === "1";

    const where = ["ld.expiry_date IS NOT NULL", "TRIM(ld.expiry_date) <> ''"];
    const params = [];
    if (department) {
      where.push("ld.department = ?");
      params.push(department);
    }
    if (status && status !== "all") {
      where.push("ld.status = ?");
      params.push(status);
    }

    const rows = db.prepare(`
      SELECT
        ld.id,
        ld.department,
        ld.title,
        ld.doc_type,
        ld.version,
        ld.owner,
        ld.status,
        ld.active,
        ld.expiry_date,
        CAST(julianday(ld.expiry_date) - julianday(DATE('now')) AS INTEGER) AS days_to_expiry
      FROM legal_documents ld
      WHERE ${where.join(" AND ")}
      ORDER BY ld.expiry_date ASC, ld.id DESC
      LIMIT 1000
    `).all(...params).map((r) => ({
      ...r,
      id: Number(r.id),
      active: Number(r.active),
      days_to_expiry: Number(r.days_to_expiry),
    }));

    const dueRows = rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= days);
    const expiredRows = rows.filter((r) => Number(r.days_to_expiry) < 0);
    const summary = {
      total_with_expiry: rows.length,
      expired: expiredRows.length,
      due_30: rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= 30).length,
      due_60: rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= 60).length,
      due_90: rows.filter((r) => Number(r.days_to_expiry) >= 0 && Number(r.days_to_expiry) <= 90).length,
      due_window: dueRows.length,
    };

    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "Legal Compliance Summary");
        kvGrid(doc, [
          { k: "Window (days)", v: String(days) },
          { k: "Department", v: department || "All" },
          { k: "Status", v: status || "approved" },
          { k: "Total with Expiry", v: fmtNum(summary.total_with_expiry, 0) },
          { k: "Expired", v: fmtNum(summary.expired, 0) },
          { k: "Due in 30 days", v: fmtNum(summary.due_30, 0) },
          { k: "Due in 60 days", v: fmtNum(summary.due_60, 0) },
          { k: "Due in 90 days", v: fmtNum(summary.due_90, 0) },
          { k: `Due in ${days} days`, v: fmtNum(summary.due_window, 0) },
        ], 2);

        sectionTitle(doc, "Expired Documents");
        table(
          doc,
          [
            { key: "id", label: "ID", width: 0.08, align: "right" },
            { key: "department", label: "Department", width: 0.16 },
            { key: "title", label: "Title", width: 0.30 },
            { key: "status", label: "Status", width: 0.12 },
            { key: "expiry", label: "Expiry", width: 0.14 },
            { key: "days", label: "Days", width: 0.10, align: "right" },
            { key: "owner", label: "Owner", width: 0.10 },
          ],
          expiredRows.length
            ? expiredRows.map((r) => ({
                id: String(r.id),
                department: r.department || "-",
                title: compactCell(r.title || "-", 100),
                status: r.status || "-",
                expiry: r.expiry_date || "-",
                days: fmtNum(r.days_to_expiry, 0),
                owner: compactCell(r.owner || "-", 40),
              }))
            : [{ id: "-", department: "-", title: "No expired documents", status: "-", expiry: "-", days: "-", owner: "-" }]
        );

        sectionTitle(doc, `Due In ${days} Days`);
        table(
          doc,
          [
            { key: "id", label: "ID", width: 0.08, align: "right" },
            { key: "department", label: "Department", width: 0.16 },
            { key: "title", label: "Title", width: 0.30 },
            { key: "status", label: "Status", width: 0.12 },
            { key: "expiry", label: "Expiry", width: 0.14 },
            { key: "days", label: "Days", width: 0.10, align: "right" },
            { key: "owner", label: "Owner", width: 0.10 },
          ],
          dueRows.length
            ? dueRows.map((r) => ({
                id: String(r.id),
                department: r.department || "-",
                title: compactCell(r.title || "-", 100),
                status: r.status || "-",
                expiry: r.expiry_date || "-",
                days: fmtNum(r.days_to_expiry, 0),
                owner: compactCell(r.owner || "-", 40),
              }))
            : [{ id: "-", department: "-", title: "No due documents in selected window", status: "-", expiry: "-", days: "-", owner: "-" }]
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Legal Compliance Report",
        rightText: `${department || "All"} | ${days}d`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="IRONLOG_Legal_Compliance_${days}d.pdf"`
      )
      .send(pdf);
  });

  // =========================
  // DAILY XLSX
  // =========================
  app.get("/daily.xlsx", async (req, reply) => {
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");

    const date = String(req.query?.date || "").trim();
    const scheduled = Number(req.query?.scheduled ?? 10);

    if (!isDate(date)) return reply.code(400).send({ error: "date (YYYY-MM-DD) required" });

    const hours = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        a.category,
        dh.is_used,
        dh.scheduled_hours,
        dh.opening_hours,
        dh.closing_hours,
        dh.hours_run,
        dh.operator,
        dh.notes
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date = ?
      ORDER BY a.asset_code
    `).all(date);

    const fuel = db.prepare(`
      SELECT a.asset_code, a.asset_name, fl.liters, fl.source
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date = ?
      ORDER BY a.asset_code
    `).all(date);

    const oil = db.prepare(`
      SELECT a.asset_code, a.asset_name, ol.oil_type, ol.quantity
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date = ?
      ORDER BY a.asset_code
    `).all(date);

    const breakdowns = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        b.description,
        b.downtime_hours,
        b.critical
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      WHERE b.breakdown_date = ?
      ORDER BY b.downtime_hours DESC
    `).all(date).map(r => ({ ...r, critical: Boolean(r.critical) }));

    const upcoming = db.prepare(`
      SELECT
        mp.id AS plan_id,
        a.asset_code,
        a.asset_name,
        mp.service_name,
        mp.interval_hours,
        mp.last_service_hours,
        COALESCE(
          (SELECT dh.closing_hours
           FROM daily_hours dh
           WHERE dh.asset_id = a.id
             AND dh.work_date <= ?
             AND dh.closing_hours IS NOT NULL
           ORDER BY dh.work_date DESC
           LIMIT 1),
          (SELECT IFNULL(SUM(dh2.hours_run), 0)
           FROM daily_hours dh2
           WHERE dh2.asset_id = a.id
             AND dh2.work_date <= ?
             AND dh2.is_used = 1
             AND dh2.hours_run > 0)
        ) AS current_hours
      FROM maintenance_plans mp
      JOIN assets a ON a.id = mp.asset_id
      WHERE mp.active = 1
        AND a.active = 1
        AND a.is_standby = 0
      ORDER BY a.asset_code
    `).all(date, date).map(r => {
      const current = Number(r.current_hours || 0);
      const next_due = Number(r.last_service_hours || 0) + Number(r.interval_hours || 0);
      const hours_left = next_due - current;
      return {
        ...r,
        current_hours: current,
        next_due,
        hours_left,
        status: hours_left <= 0 ? "OVERDUE" : (hours_left <= 50 ? "DUE SOON" : "OK"),
      };
    }).sort((a, b) => a.hours_left - b.hours_left);

    const kpi = kpiDaily(date, scheduled);
    const defaults = costDefaults();

    const fuel_total = fuel.reduce((a, r) => a + Number(r.liters || 0), 0);
    const oil_total = oil.reduce((a, r) => a + Number(r.quantity || 0), 0);
    const breakdown_total = breakdowns.reduce((a, r) => a + Number(r.downtime_hours || 0), 0);

    const fuelCostRows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS fuel_cost
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date = ?
      GROUP BY a.id
      ORDER BY a.asset_code
    `).all(defaults.fuel_cost_per_liter_default, date).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      fuel_cost: Number(r.fuel_cost || 0),
    }));

    const lubeCostRows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date = ?
      GROUP BY a.id
      ORDER BY a.asset_code
    `).all(defaults.lube_cost_per_qty_default, date).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      lube_cost: Number(r.lube_cost || 0),
    }));

    const laborRows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS labor_hours,
        COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE DATE(COALESCE(w.completed_at, w.closed_at)) = ?
        AND w.status IN ('completed', 'approved', 'closed')
      GROUP BY a.id
      ORDER BY a.asset_code
    `).all(defaults.labor_cost_per_hour_default, date).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      labor_hours: Number(r.labor_hours || 0),
      labor_cost: Number(r.labor_cost || 0),
    }));

    const downtimeRows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(l.hours_down), 0) AS downtime_hours,
        COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS downtime_cost
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date = ?
      GROUP BY a.id
      ORDER BY a.asset_code
    `).all(defaults.downtime_cost_per_hour_default, date).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      downtime_hours: Number(r.downtime_hours || 0),
      downtime_cost: Number(r.downtime_cost || 0),
    }));

    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
    const partsRows = db.prepare(`
      SELECT
        COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
        COALESCE(a.asset_name, 'Unlinked') AS asset_name,
        COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS parts_cost
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      LEFT JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
      LEFT JOIN assets a ON a.id = w.asset_id
      WHERE sm.movement_type = 'out'
        AND ${smDateExpr} = ?
      GROUP BY a.id
      ORDER BY asset_code
    `).all(date).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      parts_cost: Number(r.parts_cost || 0),
    }));

    const fuelCostTotal = Number(fuelCostRows.reduce((a, r) => a + Number(r.fuel_cost || 0), 0).toFixed(2));
    const lubeCostTotal = Number(lubeCostRows.reduce((a, r) => a + Number(r.lube_cost || 0), 0).toFixed(2));
    const laborCostTotal = Number(laborRows.reduce((a, r) => a + Number(r.labor_cost || 0), 0).toFixed(2));
    const laborHoursTotal = Number(laborRows.reduce((a, r) => a + Number(r.labor_hours || 0), 0).toFixed(2));
    const downtimeCostTotal = Number(downtimeRows.reduce((a, r) => a + Number(r.downtime_cost || 0), 0).toFixed(2));
    const partsCostTotal = Number(partsRows.reduce((a, r) => a + Number(r.parts_cost || 0), 0).toFixed(2));
    const totalCost = Number((fuelCostTotal + lubeCostTotal + laborCostTotal + downtimeCostTotal + partsCostTotal).toFixed(2));
    const costPerRunHour = Number(kpi.run_hours || 0) > 0 ? Number((totalCost / Number(kpi.run_hours || 1)).toFixed(2)) : null;

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const wsSummary = wb.addWorksheet("Summary");
    wsSummary.columns = [
      { header: "Key", key: "k", width: 28 },
      { header: "Value", key: "v", width: 22 },
    ];
    wsSummary.getRow(1).font = { bold: true };
    wsSummary.addRow({ k: "Date", v: date });
    wsSummary.addRow({ k: "Scheduled hours / asset", v: scheduled });
    wsSummary.addRow({ k: "Used assets", v: kpi.used_assets });
    wsSummary.addRow({ k: "Available hours", v: kpi.available_hours });
    wsSummary.addRow({ k: "Run hours", v: kpi.run_hours });
    wsSummary.addRow({ k: "Downtime hours", v: kpi.downtime_hours });
    wsSummary.addRow({ k: "Availability %", v: kpi.availability ?? "N/A" });
    wsSummary.addRow({ k: "Utilization %", v: kpi.utilization ?? "N/A" });
    wsSummary.addRow({ k: "Fuel total (L)", v: Number(fuel_total.toFixed(2)) });
    wsSummary.addRow({ k: "Oil total (Qty)", v: Number(oil_total.toFixed(2)) });
    wsSummary.addRow({ k: "Downtime total (hrs)", v: Number(breakdown_total.toFixed(2)) });
    wsSummary.views = [{ state: "frozen", ySplit: 1 }];
    wsSummary.getCell("A1").font = { bold: true };

    addTableSheet(
      wb,
      "Hours",
      [
        { header: "Asset Code", key: "asset_code", width: 14 },
        { header: "Asset Name", key: "asset_name", width: 26 },
        { header: "Category", key: "category", width: 14 },
        { header: "Production", key: "is_used", width: 12 },
        { header: "Scheduled", key: "scheduled_hours", width: 12 },
        { header: "Opening", key: "opening_hours", width: 12 },
        { header: "Closing", key: "closing_hours", width: 12 },
        { header: "Run Hours", key: "hours_run", width: 12 },
        { header: "Operator", key: "operator", width: 16 },
        { header: "Notes", key: "notes", width: 30 },
      ],
      hours.map(r => ({ ...r, is_used: r.is_used ? "Y" : "N" }))
    );

    addTableSheet(
      wb,
      "Breakdowns",
      [
        { header: "Asset Code", key: "asset_code", width: 14 },
        { header: "Asset Name", key: "asset_name", width: 26 },
        { header: "Downtime (hrs)", key: "downtime_hours", width: 14 },
        { header: "Critical", key: "critical", width: 10 },
        { header: "Description", key: "description", width: 40 },
      ],
      breakdowns.map(r => ({ ...r, critical: r.critical ? "YES" : "NO" }))
    );

    addTableSheet(
      wb,
      "Fuel",
      [
        { header: "Asset Code", key: "asset_code", width: 14 },
        { header: "Asset Name", key: "asset_name", width: 26 },
        { header: "Liters", key: "liters", width: 12 },
        { header: "Source", key: "source", width: 18 },
      ],
      fuel
    );

    addTableSheet(
      wb,
      "Oil",
      [
        { header: "Asset Code", key: "asset_code", width: 14 },
        { header: "Asset Name", key: "asset_name", width: 26 },
        { header: "Oil Type", key: "oil_type", width: 16 },
        { header: "Quantity", key: "quantity", width: 12 },
      ],
      oil
    );

    addTableSheet(
      wb,
      "Upcoming Services",
      [
        { header: "Asset Code", key: "asset_code", width: 14 },
        { header: "Asset Name", key: "asset_name", width: 26 },
        { header: "Service", key: "service_name", width: 22 },
        { header: "Interval (hrs)", key: "interval_hours", width: 14 },
        { header: "Last Service (hrs)", key: "last_service_hours", width: 16 },
        { header: "Current Hours", key: "current_hours", width: 14 },
        { header: "Next Due @", key: "next_due", width: 12 },
        { header: "Hours Left", key: "hours_left", width: 12 },
        { header: "Status", key: "status", width: 12 },
      ],
      upcoming
    );

    const costByAsset = new Map();
    const mergeCostRows = (rows, key) => {
      for (const r of rows) {
        const code = String(r.asset_code || "UNLINKED");
        if (!costByAsset.has(code)) {
          costByAsset.set(code, {
            asset_code: code,
            asset_name: r.asset_name || "Unlinked",
            fuel_cost: 0,
            lube_cost: 0,
            parts_cost: 0,
            labor_hours: 0,
            labor_cost: 0,
            downtime_hours: 0,
            downtime_cost: 0,
            total_cost: 0,
          });
        }
        const row = costByAsset.get(code);
        row[key] += Number(r[key] || 0);
      }
    };
    mergeCostRows(fuelCostRows, "fuel_cost");
    mergeCostRows(lubeCostRows, "lube_cost");
    mergeCostRows(partsRows, "parts_cost");
    mergeCostRows(laborRows, "labor_hours");
    mergeCostRows(laborRows, "labor_cost");
    mergeCostRows(downtimeRows, "downtime_hours");
    mergeCostRows(downtimeRows, "downtime_cost");

    const costRows = Array.from(costByAsset.values())
      .map((r) => ({
        ...r,
        fuel_cost: Number(r.fuel_cost.toFixed(2)),
        lube_cost: Number(r.lube_cost.toFixed(2)),
        parts_cost: Number(r.parts_cost.toFixed(2)),
        labor_hours: Number(r.labor_hours.toFixed(2)),
        labor_cost: Number(r.labor_cost.toFixed(2)),
        downtime_hours: Number(r.downtime_hours.toFixed(2)),
        downtime_cost: Number(r.downtime_cost.toFixed(2)),
        total_cost: Number((r.fuel_cost + r.lube_cost + r.parts_cost + r.labor_cost + r.downtime_cost).toFixed(2)),
      }))
      .filter((r) => r.total_cost > 0)
      .sort((a, b) => b.total_cost - a.total_cost);

    const buffer = await wb.xlsx.writeBuffer();

    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Daily_${date}.xlsx"`)
      .send(Buffer.from(buffer));
  });

  // =========================
  // MONTHLY COST XLSX
  // =========================
  // GET /api/reports/cost-monthly.xlsx?month=YYYY-MM
  app.get("/cost-monthly.xlsx", async (req, reply) => {
    const month = String(req.query?.month || "").trim();
    if (!isMonth(month)) {
      return reply.code(400).send({ error: "month (YYYY-MM) required" });
    }

    const current = monthRange(month);
    const previousMonth = prevMonth(month);
    const previous = monthRange(previousMonth);
    const defaults = costDefaults();

    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";

    const buildAssetCosts = (start, end) => {
      const fuelRows = db.prepare(`
        SELECT a.asset_code, a.asset_name, a.category,
          COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS fuel_cost
        FROM fuel_logs fl
        JOIN assets a ON a.id = fl.asset_id
        WHERE fl.log_date BETWEEN ? AND ?
        GROUP BY a.id
      `).all(defaults.fuel_cost_per_liter_default, start, end);

      const lubeRows = db.prepare(`
        SELECT a.asset_code, a.asset_name, a.category,
          COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS lube_cost
        FROM oil_logs ol
        JOIN assets a ON a.id = ol.asset_id
        WHERE ol.log_date BETWEEN ? AND ?
        GROUP BY a.id
      `).all(defaults.lube_cost_per_qty_default, start, end);

      const partsRows = db.prepare(`
        SELECT
          COALESCE(a.asset_code, 'UNLINKED') AS asset_code,
          COALESCE(a.asset_name, 'Unlinked') AS asset_name,
          COALESCE(a.category, 'Unassigned') AS category,
          COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS parts_cost
        FROM stock_movements sm
        JOIN parts p ON p.id = sm.part_id
        LEFT JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
        LEFT JOIN assets a ON a.id = w.asset_id
        WHERE sm.movement_type = 'out'
          AND ${smDateExpr} BETWEEN ? AND ?
        GROUP BY a.id
      `).all(start, end);

      const laborRows = db.prepare(`
        SELECT a.asset_code, a.asset_name, a.category,
          COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS labor_hours,
          COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
        FROM work_orders w
        JOIN assets a ON a.id = w.asset_id
        WHERE DATE(COALESCE(w.completed_at, w.closed_at)) BETWEEN ? AND ?
          AND w.status IN ('completed', 'approved', 'closed')
        GROUP BY a.id
      `).all(defaults.labor_cost_per_hour_default, start, end);

      const downtimeRows = db.prepare(`
        SELECT a.asset_code, a.asset_name, a.category,
          COALESCE(SUM(l.hours_down), 0) AS downtime_hours,
          COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS downtime_cost
        FROM breakdown_downtime_logs l
        JOIN breakdowns b ON b.id = l.breakdown_id
        JOIN assets a ON a.id = b.asset_id
        WHERE l.log_date BETWEEN ? AND ?
        GROUP BY a.id
      `).all(defaults.downtime_cost_per_hour_default, start, end);

      const map = new Map();
      const ensure = (r) => {
        const code = String(r.asset_code || "UNLINKED");
        if (!map.has(code)) {
          map.set(code, {
            asset_code: code,
            asset_name: r.asset_name || "Unlinked",
            category: r.category || "Unassigned",
            fuel_cost: 0,
            lube_cost: 0,
            parts_cost: 0,
            labor_hours: 0,
            labor_cost: 0,
            downtime_hours: 0,
            downtime_cost: 0,
            total_cost: 0,
          });
        }
        return map.get(code);
      };
      for (const r of fuelRows) ensure(r).fuel_cost += Number(r.fuel_cost || 0);
      for (const r of lubeRows) ensure(r).lube_cost += Number(r.lube_cost || 0);
      for (const r of partsRows) ensure(r).parts_cost += Number(r.parts_cost || 0);
      for (const r of laborRows) {
        const row = ensure(r);
        row.labor_hours += Number(r.labor_hours || 0);
        row.labor_cost += Number(r.labor_cost || 0);
      }
      for (const r of downtimeRows) {
        const row = ensure(r);
        row.downtime_hours += Number(r.downtime_hours || 0);
        row.downtime_cost += Number(r.downtime_cost || 0);
      }

      return Array.from(map.values())
        .map((r) => {
          const total = Number(r.fuel_cost || 0) + Number(r.lube_cost || 0) + Number(r.parts_cost || 0) + Number(r.labor_cost || 0) + Number(r.downtime_cost || 0);
          return {
            ...r,
            fuel_cost: Number(r.fuel_cost.toFixed(2)),
            lube_cost: Number(r.lube_cost.toFixed(2)),
            parts_cost: Number(r.parts_cost.toFixed(2)),
            labor_hours: Number(r.labor_hours.toFixed(2)),
            labor_cost: Number(r.labor_cost.toFixed(2)),
            downtime_hours: Number(r.downtime_hours.toFixed(2)),
            downtime_cost: Number(r.downtime_cost.toFixed(2)),
            total_cost: Number(total.toFixed(2)),
          };
        })
        .filter((r) => r.total_cost > 0);
    };

    const currentAssetCosts = buildAssetCosts(current.start, current.end);
    const prevAssetCosts = buildAssetCosts(previous.start, previous.end);
    const prevByAsset = new Map(prevAssetCosts.map((r) => [r.asset_code, Number(r.total_cost || 0)]));

    const assetsWithVariance = currentAssetCosts
      .map((r) => {
        const prevTotal = Number(prevByAsset.get(r.asset_code) || 0);
        const variance = Number((r.total_cost - prevTotal).toFixed(2));
        const variance_pct = prevTotal > 0 ? Number((((r.total_cost - prevTotal) / prevTotal) * 100).toFixed(2)) : null;
        return {
          ...r,
          prev_total_cost: Number(prevTotal.toFixed(2)),
          variance,
          variance_pct,
        };
      })
      .sort((a, b) => b.total_cost - a.total_cost);

    const rollupByCategory = (rows) => {
      const m = new Map();
      for (const r of rows) {
        const key = String(r.category || "Unassigned");
        if (!m.has(key)) {
          m.set(key, { category: key, fuel_cost: 0, lube_cost: 0, parts_cost: 0, labor_cost: 0, downtime_cost: 0, total_cost: 0 });
        }
        const row = m.get(key);
        row.fuel_cost += Number(r.fuel_cost || 0);
        row.lube_cost += Number(r.lube_cost || 0);
        row.parts_cost += Number(r.parts_cost || 0);
        row.labor_cost += Number(r.labor_cost || 0);
        row.downtime_cost += Number(r.downtime_cost || 0);
        row.total_cost += Number(r.total_cost || 0);
      }
      return Array.from(m.values()).map((r) => ({
        ...r,
        fuel_cost: Number(r.fuel_cost.toFixed(2)),
        lube_cost: Number(r.lube_cost.toFixed(2)),
        parts_cost: Number(r.parts_cost.toFixed(2)),
        labor_cost: Number(r.labor_cost.toFixed(2)),
        downtime_cost: Number(r.downtime_cost.toFixed(2)),
        total_cost: Number(r.total_cost.toFixed(2)),
      }));
    };

    const currentCat = rollupByCategory(assetsWithVariance);
    const prevCat = rollupByCategory(prevAssetCosts);
    const prevByCat = new Map(prevCat.map((r) => [r.category, Number(r.total_cost || 0)]));
    const categoryWithVariance = currentCat
      .map((r) => {
        const prevTotal = Number(prevByCat.get(r.category) || 0);
        const variance = Number((r.total_cost - prevTotal).toFixed(2));
        const variance_pct = prevTotal > 0 ? Number((((r.total_cost - prevTotal) / prevTotal) * 100).toFixed(2)) : null;
        return {
          ...r,
          prev_total_cost: Number(prevTotal.toFixed(2)),
          variance,
          variance_pct,
        };
      })
      .sort((a, b) => b.total_cost - a.total_cost);

    const totals = assetsWithVariance.reduce((acc, r) => {
      acc.fuel += Number(r.fuel_cost || 0);
      acc.lube += Number(r.lube_cost || 0);
      acc.parts += Number(r.parts_cost || 0);
      acc.labor += Number(r.labor_cost || 0);
      acc.downtime += Number(r.downtime_cost || 0);
      acc.total += Number(r.total_cost || 0);
      return acc;
    }, { fuel: 0, lube: 0, parts: 0, labor: 0, downtime: 0, total: 0 });
    const prevTotal = Number(prevAssetCosts.reduce((acc, r) => acc + Number(r.total_cost || 0), 0).toFixed(2));
    const varianceTotal = Number((Number(totals.total.toFixed(2)) - prevTotal).toFixed(2));
    const variancePct = prevTotal > 0 ? Number(((varianceTotal / prevTotal) * 100).toFixed(2)) : null;

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const wsSummary = wb.addWorksheet("Summary");
    wsSummary.columns = [
      { header: "Key", key: "k", width: 32 },
      { header: "Value", key: "v", width: 24 },
    ];
    wsSummary.getRow(1).font = { bold: true };
    wsSummary.addRow({ k: "Month", v: month });
    wsSummary.addRow({ k: "Period", v: `${current.start} to ${current.end}` });
    wsSummary.addRow({ k: "Previous Month", v: `${previousMonth} (${previous.start} to ${previous.end})` });
    wsSummary.addRow({ k: "Assets with Cost Activity", v: assetsWithVariance.length });
    wsSummary.addRow({ k: "Fuel Cost", v: Number(totals.fuel.toFixed(2)) });
    wsSummary.addRow({ k: "Lube Cost", v: Number(totals.lube.toFixed(2)) });
    wsSummary.addRow({ k: "Parts Cost", v: Number(totals.parts.toFixed(2)) });
    wsSummary.addRow({ k: "Labor Cost", v: Number(totals.labor.toFixed(2)) });
    wsSummary.addRow({ k: "Downtime Cost", v: Number(totals.downtime.toFixed(2)) });
    wsSummary.addRow({ k: "Total Cost", v: Number(totals.total.toFixed(2)) });
    wsSummary.addRow({ k: "Previous Total Cost", v: prevTotal });
    wsSummary.addRow({ k: "Variance", v: varianceTotal });
    wsSummary.addRow({ k: "Variance %", v: variancePct == null ? "N/A" : variancePct });
    wsSummary.views = [{ state: "frozen", ySplit: 1 }];

    addTableSheet(
      wb,
      "Asset Costs",
      [
        { header: "Asset Code", key: "asset_code", width: 14 },
        { header: "Asset Name", key: "asset_name", width: 24 },
        { header: "Category", key: "category", width: 16 },
        { header: "Fuel", key: "fuel_cost", width: 12 },
        { header: "Lube", key: "lube_cost", width: 12 },
        { header: "Parts", key: "parts_cost", width: 12 },
        { header: "Labor Hrs", key: "labor_hours", width: 11 },
        { header: "Labor", key: "labor_cost", width: 12 },
        { header: "Downtime Hrs", key: "downtime_hours", width: 13 },
        { header: "Downtime", key: "downtime_cost", width: 12 },
        { header: "Total", key: "total_cost", width: 12 },
        { header: `Prev (${previousMonth})`, key: "prev_total_cost", width: 14 },
        { header: "Variance", key: "variance", width: 12 },
        { header: "Variance %", key: "variance_pct", width: 12 },
      ],
      assetsWithVariance.length
        ? assetsWithVariance
        : [{
            asset_code: "-",
            asset_name: "No cost activity in month",
            category: "-",
            fuel_cost: 0, lube_cost: 0, parts_cost: 0, labor_hours: 0, labor_cost: 0, downtime_hours: 0, downtime_cost: 0, total_cost: 0,
            prev_total_cost: 0, variance: 0, variance_pct: null,
          }]
    );

    addTableSheet(
      wb,
      "Category Costs",
      [
        { header: "Category", key: "category", width: 20 },
        { header: "Fuel", key: "fuel_cost", width: 12 },
        { header: "Lube", key: "lube_cost", width: 12 },
        { header: "Parts", key: "parts_cost", width: 12 },
        { header: "Labor", key: "labor_cost", width: 12 },
        { header: "Downtime", key: "downtime_cost", width: 12 },
        { header: "Total", key: "total_cost", width: 12 },
        { header: `Prev (${previousMonth})`, key: "prev_total_cost", width: 14 },
        { header: "Variance", key: "variance", width: 12 },
        { header: "Variance %", key: "variance_pct", width: 12 },
      ],
      categoryWithVariance.length
        ? categoryWithVariance
        : [{
            category: "Unassigned", fuel_cost: 0, lube_cost: 0, parts_cost: 0, labor_cost: 0, downtime_cost: 0,
            total_cost: 0, prev_total_cost: 0, variance: 0, variance_pct: null,
          }]
    );

    const buffer = await wb.xlsx.writeBuffer();
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Cost_Monthly_${month}.xlsx"`)
      .send(Buffer.from(buffer));
  });

  // =========================
  // DAILY PDF
  // =========================
  app.get("/daily.pdf", async (req, reply) => {
    const reportRevision = "daily-pdf-no-cost-r2026-04-04b";
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    reply.header("X-IRONLOG-Report-Revision", reportRevision);

    const date = String(req.query?.date || "").trim();
    const scheduled = Number(req.query?.scheduled ?? 10);
    if (!isDate(date)) return reply.code(400).send({ error: "date (YYYY-MM-DD) required" });

    const logoPath = path.join(process.cwd(), "branding", "logo.png");

    const hours = db.prepare(`
      SELECT a.asset_code, a.asset_name, dh.hours_run, dh.is_used, dh.operator, dh.notes
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date = ?
      ORDER BY a.asset_code
    `).all(date);

    const fuel = db.prepare(`
      SELECT a.asset_code, fl.liters, fl.source
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date = ?
      ORDER BY a.asset_code
    `).all(date);

    const oil = db.prepare(`
      SELECT a.asset_code, ol.oil_type, ol.quantity
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date = ?
      ORDER BY a.asset_code
    `).all(date);

    const breakdowns = db.prepare(`
      SELECT a.asset_code, b.description, b.downtime_hours, b.critical
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      WHERE b.breakdown_date = ?
      ORDER BY b.downtime_hours DESC
    `).all(date).map(r => ({ ...r, critical: Boolean(r.critical) }));

    const openWOs = db.prepare(`
      SELECT w.id, a.asset_code, w.source, w.status, w.opened_at
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE w.closed_at IS NULL
        AND REPLACE(TRIM(LOWER(COALESCE(w.status, ''))), ' ', '_') IN ('open', 'assigned', 'in_progress')
        AND (w.completed_at IS NULL OR TRIM(COALESCE(w.completed_at, '')) = '')
      ORDER BY w.id DESC
      LIMIT 30
    `).all();

    const stockCritical = db.prepare(`
      SELECT p.part_code, p.part_name, p.min_stock, IFNULL(SUM(sm.quantity),0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      WHERE p.critical = 1
      GROUP BY p.id
      ORDER BY on_hand ASC
      LIMIT 20
    `).all().map(r => ({
      ...r,
      on_hand: Number(r.on_hand),
      below_min: Number(r.on_hand) < Number(r.min_stock),
    }));
    const hoursPdf = hours.slice(0, 40);
    const fuelPdf = fuel.slice(0, 40);
    const oilPdf = oil.slice(0, 40);
    const breakdownsPdf = breakdowns.slice(0, 40);
    const openWOsPdf = openWOs.slice(0, 40);
    const stockCriticalPdf = stockCritical.slice(0, 40);

    const kpi = kpiDaily(date, scheduled);
    const defaults = costDefaults();

    const fuelCostRow = db.prepare(`
      SELECT COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS value
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date = ?
    `).get(defaults.fuel_cost_per_liter_default, date);

    const lubeCostRow = db.prepare(`
      SELECT COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS value
      FROM oil_logs ol
      WHERE ol.log_date = ?
    `).get(defaults.lube_cost_per_qty_default, date);

    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
    const partsCostRow = db.prepare(`
      SELECT COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS value
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      WHERE sm.movement_type = 'out'
        AND ${smDateExpr} = ?
    `).get(date);

    const laborRow = db.prepare(`
      SELECT
        COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS labor_hours,
        COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      WHERE DATE(COALESCE(w.completed_at, w.closed_at)) = ?
        AND w.status IN ('completed', 'approved', 'closed')
    `).get(defaults.labor_cost_per_hour_default, date);

    const downtimeCostRow = db.prepare(`
      SELECT COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS value
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date = ?
    `).get(defaults.downtime_cost_per_hour_default, date);

    const totalCost = Number(
      (
        Number(fuelCostRow?.value || 0) +
        Number(lubeCostRow?.value || 0) +
        Number(partsCostRow?.value || 0) +
        Number(laborRow?.labor_cost || 0) +
        Number(downtimeCostRow?.value || 0)
      ).toFixed(2)
    );
    const costPerRunHour = Number(kpi.run_hours || 0) > 0 ? totalCost / Number(kpi.run_hours || 1) : null;

    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "KPIs");
        kvGrid(doc, [
          { k: "Date", v: date },
          { k: "Scheduled hours / asset", v: fmtNum(scheduled, 0) },
          { k: "Used assets", v: fmtNum(kpi.used_assets, 0) },
          { k: "Available hours", v: fmtNum(kpi.available_hours, 0) },
          { k: "Run hours", v: fmtNum(kpi.run_hours, 1) },
          { k: "Downtime hours", v: fmtNum(kpi.downtime_hours, 1) },
          { k: "Availability %", v: kpi.availability == null ? "N/A" : `${fmtNum(kpi.availability, 2)}%` },
          { k: "Utilization %", v: kpi.utilization == null ? "N/A" : `${fmtNum(kpi.utilization, 2)}%` },
        ], 2);

        sectionTitle(doc, "Hours Logged");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.16 },
            { key: "name", label: "Name", width: 0.24 },
            { key: "hours", label: "Run Hrs", width: 0.10, align: "right" },
            { key: "used", label: "Used", width: 0.08, align: "center" },
            { key: "operator", label: "Operator", width: 0.18 },
            { key: "notes", label: "Notes", width: 0.24 },
          ],
          hoursPdf.map(r => ({
            asset: r.asset_code,
            name: r.asset_name ?? "",
            hours: fmtNum(r.hours_run, 1),
            used: r.is_used ? "Y" : "N",
            operator: compactCell(r.operator ?? "", 40),
            notes: compactCell(r.notes ?? "", 90),
          }))
        );

        sectionTitle(doc, "Fuel");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.20 },
            { key: "liters", label: "Liters", width: 0.14, align: "right" },
            { key: "source", label: "Source", width: 0.66 },
          ],
          fuelPdf.map(r => ({
            asset: r.asset_code,
            liters: fmtNum(r.liters, 1),
            source: compactCell(r.source ?? "", 90),
          }))
        );

        sectionTitle(doc, "Oil");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.20 },
            { key: "type", label: "Type", width: 0.50 },
            { key: "qty", label: "Qty", width: 0.30, align: "right" },
          ],
          oilPdf.map(r => ({
            asset: r.asset_code,
            type: compactCell(r.oil_type ?? "", 80),
            qty: fmtNum(r.quantity, 1),
          }))
        );

        sectionTitle(doc, "Breakdowns & Downtime");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.16 },
            { key: "hrs", label: "Downtime (hrs)", width: 0.14, align: "right" },
            { key: "crit", label: "Critical", width: 0.12, align: "center" },
            { key: "desc", label: "Description", width: 0.58 },
          ],
          breakdownsPdf.map(r => ({
            asset: r.asset_code,
            hrs: fmtNum(r.downtime_hours, 1),
            crit: r.critical ? "YES" : "NO",
            desc: compactCell(r.description ?? "", 140),
          }))
        );

        sectionTitle(doc, "Open Work Orders");
        table(
          doc,
          [
            { key: "wo", label: "WO#", width: 0.12, align: "right" },
            { key: "asset", label: "Asset", width: 0.14 },
            { key: "source", label: "Source", width: 0.20 },
            { key: "status", label: "Status", width: 0.14 },
            { key: "opened", label: "Opened", width: 0.40 },
          ],
          openWOsPdf.map(r => ({
            wo: String(r.id),
            asset: r.asset_code,
            source: compactCell(r.source ?? "", 24),
            status: compactCell(r.status ?? "", 24),
            opened: r.opened_at ?? "",
          }))
        );

        sectionTitle(doc, "Critical Stock (Top 20)");
        table(
          doc,
          [
            { key: "part", label: "Part", width: 0.18 },
            { key: "name", label: "Name", width: 0.46 },
            { key: "on", label: "On hand", width: 0.12, align: "right" },
            { key: "min", label: "Min", width: 0.10, align: "right" },
            { key: "below", label: "Below min", width: 0.14, align: "center" },
          ],
          stockCriticalPdf.map(r => ({
            part: r.part_code,
            name: r.part_name ?? "",
            on: fmtNum(r.on_hand, 0),
            min: fmtNum(r.min_stock, 0),
            below: r.below_min ? "YES" : "NO",
          }))
        );
      },
      {
        title: "IRONLOG",
        subtitle: `Daily Operations Report (${reportRevision})`,
        rightText: `Date: ${date}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header("Content-Disposition", `inline; filename="IRONLOG_Daily_${date}.pdf"`)
      .send(pdf);
  });

  // =========================
  // WEEKLY PDF
  // =========================
  app.get("/weekly.pdf", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const scheduled = Number(req.query?.scheduled ?? 10);

    if (!isDate(start) || !isDate(end)) return reply.code(400).send({ error: "start and end (YYYY-MM-DD) required" });

    const logoPath = path.join(process.cwd(), "branding", "logo.png");

    const kpi = kpiRange(start, end, scheduled);
    const defaults = costDefaults();

    const majorDowntime = db.prepare(`
      SELECT a.asset_code, b.breakdown_date, b.downtime_hours, b.critical, b.description
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      WHERE b.breakdown_date BETWEEN ? AND ?
      ORDER BY b.downtime_hours DESC
      LIMIT 25
    `).all(start, end).map(r => ({ ...r, critical: Boolean(r.critical) }));

    const overdue = db.prepare(`
      SELECT
        mp.id AS plan_id,
        a.asset_code,
        a.asset_name,
        mp.service_name,
        mp.interval_hours,
        mp.last_service_hours,
        IFNULL((
          SELECT SUM(dh.hours_run)
          FROM daily_hours dh
          WHERE dh.asset_id = a.id
            AND dh.is_used = 1
            AND dh.hours_run > 0
            AND dh.work_date <= ?
        ), 0) AS current_hours
      FROM maintenance_plans mp
      JOIN assets a ON a.id = mp.asset_id
      WHERE mp.active = 1
        AND a.active = 1
        AND a.is_standby = 0
    `).all(end).map(r => {
      const current = Number(r.current_hours || 0);
      const next_due = Number(r.last_service_hours || 0) + Number(r.interval_hours || 0);
      const remaining = next_due - current;
      return { ...r, current_hours: current, next_due, remaining, is_overdue: remaining <= 0 };
    }).filter(x => x.is_overdue).sort((a, b) => a.remaining - b.remaining).slice(0, 30);

    const lowStock = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        p.critical,
        p.min_stock,
        IFNULL(SUM(sm.quantity),0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      GROUP BY p.id
      HAVING on_hand < p.min_stock
      ORDER BY p.critical DESC, on_hand ASC
      LIMIT 40
    `).all().map(r => ({ ...r, critical: Boolean(r.critical), on_hand: Number(r.on_hand) }));

    const onOrderCritical = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        po.quantity,
        po.expected_date,
        po.status
      FROM parts_orders po
      JOIN parts p ON p.id = po.part_id
      WHERE p.critical = 1
        AND po.status != 'received'
      ORDER BY po.expected_date ASC
      LIMIT 40
    `).all();
    const dailyPdf = (kpi.daily || []).slice(0, 40);
    const majorDowntimePdf = majorDowntime.slice(0, 40);
    const overduePdf = overdue.slice(0, 40);
    const lowStockPdf = lowStock.slice(0, 40);
    const onOrderCriticalPdf = onOrderCritical.slice(0, 40);

    const fuelCostRow = db.prepare(`
      SELECT COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS value
      FROM fuel_logs fl
      JOIN assets a ON a.id = fl.asset_id
      WHERE fl.log_date BETWEEN ? AND ?
    `).get(defaults.fuel_cost_per_liter_default, start, end);

    const lubeCostRow = db.prepare(`
      SELECT COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS value
      FROM oil_logs ol
      WHERE ol.log_date BETWEEN ? AND ?
    `).get(defaults.lube_cost_per_qty_default, start, end);

    const smCols = db.prepare(`PRAGMA table_info(stock_movements)`).all();
    const hasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
    const smDateExpr = hasCreatedAt ? "DATE(sm.created_at)" : "DATE(sm.movement_date)";
    const partsCostRow = db.prepare(`
      SELECT COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS value
      FROM stock_movements sm
      JOIN parts p ON p.id = sm.part_id
      WHERE sm.movement_type = 'out'
        AND ${smDateExpr} BETWEEN ? AND ?
    `).get(start, end);

    const laborRow = db.prepare(`
      SELECT
        COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS labor_hours,
        COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS labor_cost
      FROM work_orders w
      WHERE DATE(COALESCE(w.completed_at, w.closed_at)) BETWEEN ? AND ?
        AND w.status IN ('completed', 'approved', 'closed')
    `).get(defaults.labor_cost_per_hour_default, start, end);

    const downtimeCostRow = db.prepare(`
      SELECT COALESCE(SUM(l.hours_down * COALESCE(a.downtime_cost_per_hour, ?)), 0) AS value
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE l.log_date BETWEEN ? AND ?
    `).get(defaults.downtime_cost_per_hour_default, start, end);

    const totalCost = Number(
      (
        Number(fuelCostRow?.value || 0) +
        Number(lubeCostRow?.value || 0) +
        Number(partsCostRow?.value || 0) +
        Number(laborRow?.labor_cost || 0) +
        Number(downtimeCostRow?.value || 0)
      ).toFixed(2)
    );
    const costPerRunHour = Number(kpi.run_hours || 0) > 0 ? totalCost / Number(kpi.run_hours || 1) : null;

    const pdf = await buildPdfBuffer(
      (doc) => {
        tryDrawLogo(doc, logoPath);

        sectionTitle(doc, "KPIs (Period)");
        kvGrid(doc, [
          { k: "Period", v: `${start} to ${end}` },
          { k: "Scheduled hours / asset", v: fmtNum(scheduled, 0) },
          { k: "Available hours", v: fmtNum(kpi.available_hours, 0) },
          { k: "Run hours", v: fmtNum(kpi.run_hours, 1) },
          { k: "Downtime hours", v: fmtNum(kpi.downtime_hours, 1) },
          { k: "Availability %", v: kpi.availability == null ? "N/A" : `${fmtNum(kpi.availability, 2)}%` },
          { k: "Utilization %", v: kpi.utilization == null ? "N/A" : `${fmtNum(kpi.utilization, 2)}%` },
        ], 2);

        sectionTitle(doc, "Cost Engine (Period)");
        kvGrid(doc, [
          { k: "Fuel Cost", v: fmtNum(fuelCostRow?.value || 0, 2) },
          { k: "Oil/Lube Cost", v: fmtNum(lubeCostRow?.value || 0, 2) },
          { k: "Parts Cost", v: fmtNum(partsCostRow?.value || 0, 2) },
          { k: "Labor Cost", v: fmtNum(laborRow?.labor_cost || 0, 2) },
          { k: "Labor Hours", v: fmtNum(laborRow?.labor_hours || 0, 1) },
          { k: "Downtime Cost", v: fmtNum(downtimeCostRow?.value || 0, 2) },
          { k: "Total Cost", v: fmtNum(totalCost, 2) },
          { k: "Cost / Run Hour", v: costPerRunHour == null ? "N/A" : fmtNum(costPerRunHour, 2) },
        ], 2);

        sectionTitle(doc, "Daily Summary");
        table(
          doc,
          [
            { key: "date", label: "Date", width: 0.22 },
            { key: "used", label: "Used assets", width: 0.18, align: "right" },
            { key: "avail", label: "Avail hrs", width: 0.30, align: "right" },
            { key: "run", label: "Run hrs", width: 0.30, align: "right" },
          ],
          dailyPdf.map(d => ({
            date: d.date,
            used: fmtNum(d.used_assets, 0),
            avail: fmtNum(d.available_hours, 0),
            run: fmtNum(d.run_hours, 1),
          }))
        );

        sectionTitle(doc, "Major Downtime (Top 25)");
        table(
          doc,
          [
            { key: "date", label: "Date", width: 0.16 },
            { key: "asset", label: "Asset", width: 0.14 },
            { key: "hrs", label: "Hrs", width: 0.10, align: "right" },
            { key: "crit", label: "Crit", width: 0.10, align: "center" },
            { key: "desc", label: "Description", width: 0.50 },
          ],
          majorDowntimePdf.map(r => ({
            date: r.breakdown_date,
            asset: r.asset_code,
            hrs: fmtNum(r.downtime_hours, 1),
            crit: r.critical ? "YES" : "NO",
            desc: compactCell(r.description ?? "", 120),
          }))
        );

        sectionTitle(doc, "Overdue Maintenance (Top 30)");
        table(
          doc,
          [
            { key: "asset", label: "Asset", width: 0.14 },
            { key: "service", label: "Service", width: 0.40 },
            { key: "current", label: "Current", width: 0.15, align: "right" },
            { key: "next", label: "Next due", width: 0.15, align: "right" },
            { key: "over", label: "Overdue by", width: 0.16, align: "right" },
          ],
          overduePdf.map(r => ({
            asset: r.asset_code,
            service: compactCell(r.service_name ?? "", 90),
            current: fmtNum(r.current_hours, 1),
            next: fmtNum(r.next_due, 1),
            over: fmtNum(Math.abs(r.remaining), 1),
          }))
        );

        sectionTitle(doc, "Low Stock (Below Min)");
        table(
          doc,
          [
            { key: "part", label: "Part", width: 0.18 },
            { key: "name", label: "Name", width: 0.52 },
            { key: "on", label: "On hand", width: 0.14, align: "right" },
            { key: "min", label: "Min", width: 0.10, align: "right" },
            { key: "crit", label: "Critical", width: 0.06, align: "center" },
          ],
          lowStockPdf.map(r => ({
            part: r.part_code,
            name: compactCell(r.part_name ?? "", 90),
            on: fmtNum(r.on_hand, 0),
            min: fmtNum(r.min_stock, 0),
            crit: r.critical ? "Y" : "N",
          }))
        );

        sectionTitle(doc, "Critical Parts On Order");
        table(
          doc,
          [
            { key: "part", label: "Part", width: 0.18 },
            { key: "name", label: "Name", width: 0.46 },
            { key: "qty", label: "Qty", width: 0.10, align: "right" },
            { key: "exp", label: "Expected", width: 0.14 },
            { key: "status", label: "Status", width: 0.12 },
          ],
          onOrderCriticalPdf.map(r => ({
            part: r.part_code,
            name: compactCell(r.part_name ?? "", 80),
            qty: fmtNum(r.quantity, 0),
            exp: r.expected_date ?? "",
            status: r.status ?? "",
          }))
        );
      },
      {
        title: "IRONLOG",
        subtitle: "Weekly Operations Report",
        rightText: `Period: ${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header("Content-Disposition", `inline; filename="IRONLOG_Weekly_${start}_to_${end}.pdf"`)
      .send(pdf);
  });

  // GET /api/reports/operations.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&download=1
  app.get("/operations.pdf", async (req, reply) => {
    const reportRevision = "ops-pdf-r2026-04-04b";
    reply.header("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    reply.header("Pragma", "no-cache");
    reply.header("Expires", "0");
    reply.header("X-IRONLOG-Report-Revision", reportRevision);
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const download = String(req.query?.download || "").trim() === "1";
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }

    const rows = db.prepare(`
      SELECT
        op_date, tonnes_moved, product_type, product_produced, trucks_loaded, weighbridge_amount,
        trucks_delivered, product_delivered, client_delivered_to, notes
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      ORDER BY op_date ASC, id ASC
      LIMIT 1000
    `).all(start, end);

    const totals = rows.reduce((acc, r) => {
      acc.tonnes += Number(r.tonnes_moved || 0);
      acc.produced += Number(r.product_produced || 0);
      acc.loaded += Number(r.trucks_loaded || 0);
      acc.delivered += Number(r.trucks_delivered || 0);
      acc.weighbridge += Number(r.weighbridge_amount || 0);
      acc.productDelivered += Number(r.product_delivered || 0);
      return acc;
    }, { tonnes: 0, produced: 0, loaded: 0, delivered: 0, weighbridge: 0, productDelivered: 0 });

    const byProduct = db.prepare(`
      SELECT
        COALESCE(NULLIF(TRIM(product_type), ''), 'Unspecified') AS product_type,
        COUNT(*) AS entries,
        IFNULL(SUM(tonnes_moved), 0) AS tonnes_moved,
        IFNULL(SUM(product_produced), 0) AS product_produced,
        IFNULL(SUM(product_delivered), 0) AS product_delivered
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      GROUP BY COALESCE(NULLIF(TRIM(product_type), ''), 'Unspecified')
      ORDER BY tonnes_moved DESC
      LIMIT 100
    `).all(start, end);

    const byClient = db.prepare(`
      SELECT
        COALESCE(NULLIF(TRIM(client_delivered_to), ''), 'Unspecified') AS client_name,
        COUNT(*) AS entries,
        IFNULL(SUM(trucks_delivered), 0) AS trucks_delivered,
        IFNULL(SUM(product_delivered), 0) AS product_delivered
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      GROUP BY COALESCE(NULLIF(TRIM(client_delivered_to), ''), 'Unspecified')
      ORDER BY product_delivered DESC
      LIMIT 100
    `).all(start, end);

    const pdf = await buildPdfBuffer(
      (doc) => {
        const logoPath = path.join(process.cwd(), "branding", "logo.png");
        tryDrawLogo(doc, logoPath);
        sectionTitle(doc, "Operations Summary");
        kvGrid(doc, [
          { label: "Period", value: `${start} to ${end}` },
          { label: "Entries", value: String(rows.length) },
          { label: "Tonnes moved", value: fmtNum(totals.tonnes, 2) },
          { label: "Produced", value: fmtNum(totals.produced, 2) },
          { label: "Delivered", value: fmtNum(totals.productDelivered, 2) },
          { label: "Trucks loaded", value: fmtNum(totals.loaded, 0) },
          { label: "Trucks delivered", value: fmtNum(totals.delivered, 0) },
          { label: "Weighbridge", value: fmtNum(totals.weighbridge, 2) },
        ], 2);

        sectionTitle(doc, "By Product Type");
        table(
          doc,
          [
            { key: "product", label: "Product", width: 0.36 },
            { key: "entries", label: "Entries", width: 0.12, align: "right" },
            { key: "tonnes", label: "Tonnes", width: 0.16, align: "right" },
            { key: "produced", label: "Produced", width: 0.18, align: "right" },
            { key: "delivered", label: "Delivered", width: 0.18, align: "right" },
          ],
          byProduct.map((r) => ({
            product: compactCell(r.product_type, 70),
            entries: fmtNum(r.entries, 0),
            tonnes: fmtNum(r.tonnes_moved, 2),
            produced: fmtNum(r.product_produced, 2),
            delivered: fmtNum(r.product_delivered, 2),
          }))
        );

        sectionTitle(doc, "By Client");
        table(
          doc,
          [
            { key: "client", label: "Client", width: 0.46 },
            { key: "entries", label: "Entries", width: 0.12, align: "right" },
            { key: "trucks", label: "Trucks", width: 0.16, align: "right" },
            { key: "delivered", label: "Delivered", width: 0.26, align: "right" },
          ],
          byClient.map((r) => ({
            client: compactCell(r.client_name, 80),
            entries: fmtNum(r.entries, 0),
            trucks: fmtNum(r.trucks_delivered, 0),
            delivered: fmtNum(r.product_delivered, 2),
          }))
        );

        sectionTitle(doc, "Entry Detail");
        table(
          doc,
          [
            { key: "date", label: "Date", width: 0.12 },
            { key: "product", label: "Product", width: 0.16 },
            { key: "tonnes", label: "Tonnes", width: 0.10, align: "right" },
            { key: "prod", label: "Produced", width: 0.10, align: "right" },
            { key: "loaded", label: "Loaded", width: 0.08, align: "right" },
            { key: "wb", label: "Weighbridge", width: 0.12, align: "right" },
            { key: "delTrk", label: "Deliv Trucks", width: 0.10, align: "right" },
            { key: "delProd", label: "Delivered", width: 0.10, align: "right" },
            { key: "client", label: "Client", width: 0.12 },
          ],
          rows.slice(0, 250).map((r) => ({
            date: r.op_date || "",
            product: compactCell(r.product_type || "", 24),
            tonnes: fmtNum(r.tonnes_moved, 2),
            prod: fmtNum(r.product_produced, 2),
            loaded: fmtNum(r.trucks_loaded, 0),
            wb: fmtNum(r.weighbridge_amount, 2),
            delTrk: fmtNum(r.trucks_delivered, 0),
            delProd: fmtNum(r.product_delivered, 2),
            client: compactCell(r.client_delivered_to || "", 22),
          }))
        );
      },
      {
        title: "IRONLOG",
        subtitle: `Operations Report (${reportRevision})`,
        rightText: `Period: ${start} to ${end}`,
        showPageNumbers: true,
      }
    );

    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="IRONLOG_Operations_${start}_to_${end}.pdf"`
      )
      .send(pdf);
  });

  // GET /api/reports/operations.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD
  app.get("/operations.xlsx", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }

    const rows = db.prepare(`
      SELECT
        op_date, tonnes_moved, product_type, product_produced, trucks_loaded, weighbridge_amount,
        trucks_delivered, product_delivered, client_delivered_to, notes
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      ORDER BY op_date ASC, id ASC
      LIMIT 5000
    `).all(start, end);

    const byDate = db.prepare(`
      SELECT
        op_date,
        COUNT(*) AS entries,
        IFNULL(SUM(tonnes_moved), 0) AS tonnes_moved,
        IFNULL(SUM(product_produced), 0) AS product_produced,
        IFNULL(SUM(trucks_loaded), 0) AS trucks_loaded,
        IFNULL(SUM(weighbridge_amount), 0) AS weighbridge_amount,
        IFNULL(SUM(trucks_delivered), 0) AS trucks_delivered,
        IFNULL(SUM(product_delivered), 0) AS product_delivered
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      GROUP BY op_date
      ORDER BY op_date ASC
    `).all(start, end);

    const byProduct = db.prepare(`
      SELECT
        COALESCE(NULLIF(TRIM(product_type), ''), 'Unspecified') AS product_type,
        COUNT(*) AS entries,
        IFNULL(SUM(tonnes_moved), 0) AS tonnes_moved,
        IFNULL(SUM(product_produced), 0) AS product_produced,
        IFNULL(SUM(product_delivered), 0) AS product_delivered
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      GROUP BY COALESCE(NULLIF(TRIM(product_type), ''), 'Unspecified')
      ORDER BY tonnes_moved DESC
    `).all(start, end);

    const byClient = db.prepare(`
      SELECT
        COALESCE(NULLIF(TRIM(client_delivered_to), ''), 'Unspecified') AS client_name,
        COUNT(*) AS entries,
        IFNULL(SUM(trucks_delivered), 0) AS trucks_delivered,
        IFNULL(SUM(product_delivered), 0) AS product_delivered
      FROM operations_logs
      WHERE op_date BETWEEN ? AND ?
      GROUP BY COALESCE(NULLIF(TRIM(client_delivered_to), ''), 'Unspecified')
      ORDER BY product_delivered DESC
    `).all(start, end);

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    addTableSheet(
      wb,
      "Operations Detail",
      [
        { header: "Date", key: "op_date", width: 14 },
        { header: "Product Type", key: "product_type", width: 24 },
        { header: "Tonnes Moved", key: "tonnes_moved", width: 16 },
        { header: "Product Produced", key: "product_produced", width: 18 },
        { header: "Trucks Loaded", key: "trucks_loaded", width: 14 },
        { header: "Weighbridge Amount", key: "weighbridge_amount", width: 18 },
        { header: "Trucks Delivered", key: "trucks_delivered", width: 16 },
        { header: "Product Delivered", key: "product_delivered", width: 16 },
        { header: "Client Delivered To", key: "client_delivered_to", width: 24 },
        { header: "Notes", key: "notes", width: 36 },
      ],
      rows.map((r) => ({
        op_date: r.op_date || "",
        product_type: r.product_type || "",
        tonnes_moved: Number(r.tonnes_moved || 0),
        product_produced: Number(r.product_produced || 0),
        trucks_loaded: Number(r.trucks_loaded || 0),
        weighbridge_amount: Number(r.weighbridge_amount || 0),
        trucks_delivered: Number(r.trucks_delivered || 0),
        product_delivered: Number(r.product_delivered || 0),
        client_delivered_to: r.client_delivered_to || "",
        notes: r.notes || "",
      }))
    );

    addTableSheet(
      wb,
      "By Date",
      [
        { header: "Date", key: "op_date", width: 14 },
        { header: "Entries", key: "entries", width: 12 },
        { header: "Tonnes Moved", key: "tonnes_moved", width: 16 },
        { header: "Produced", key: "product_produced", width: 14 },
        { header: "Trucks Loaded", key: "trucks_loaded", width: 14 },
        { header: "Weighbridge", key: "weighbridge_amount", width: 14 },
        { header: "Trucks Delivered", key: "trucks_delivered", width: 16 },
        { header: "Delivered", key: "product_delivered", width: 14 },
      ],
      byDate.map((r) => ({
        op_date: r.op_date,
        entries: Number(r.entries || 0),
        tonnes_moved: Number(r.tonnes_moved || 0),
        product_produced: Number(r.product_produced || 0),
        trucks_loaded: Number(r.trucks_loaded || 0),
        weighbridge_amount: Number(r.weighbridge_amount || 0),
        trucks_delivered: Number(r.trucks_delivered || 0),
        product_delivered: Number(r.product_delivered || 0),
      }))
    );

    addTableSheet(
      wb,
      "By Product",
      [
        { header: "Product Type", key: "product_type", width: 26 },
        { header: "Entries", key: "entries", width: 12 },
        { header: "Tonnes Moved", key: "tonnes_moved", width: 16 },
        { header: "Produced", key: "product_produced", width: 14 },
        { header: "Delivered", key: "product_delivered", width: 14 },
      ],
      byProduct.map((r) => ({
        product_type: r.product_type,
        entries: Number(r.entries || 0),
        tonnes_moved: Number(r.tonnes_moved || 0),
        product_produced: Number(r.product_produced || 0),
        product_delivered: Number(r.product_delivered || 0),
      }))
    );

    addTableSheet(
      wb,
      "By Client",
      [
        { header: "Client", key: "client_name", width: 28 },
        { header: "Entries", key: "entries", width: 12 },
        { header: "Trucks Delivered", key: "trucks_delivered", width: 16 },
        { header: "Product Delivered", key: "product_delivered", width: 16 },
      ],
      byClient.map((r) => ({
        client_name: r.client_name,
        entries: Number(r.entries || 0),
        trucks_delivered: Number(r.trucks_delivered || 0),
        product_delivered: Number(r.product_delivered || 0),
      }))
    );

    const buffer = await wb.xlsx.writeBuffer();
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Operations_${start}_to_${end}.xlsx"`)
      .send(Buffer.from(buffer));
  });
}