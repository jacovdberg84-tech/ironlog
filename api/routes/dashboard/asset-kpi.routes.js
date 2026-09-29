// IRONLOG/api/routes/dashboard/asset-kpi.routes.js — Asset KPI weekly data and exports, LDV pre-start compliance.
// Registered by routes/dashboard.routes.js; shared helpers arrive through ctx.
import ExcelJS from "exceljs";
import { AlignmentType, BorderStyle, Document, HeadingLevel, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from "docx";
import { db } from "../../db/client.js";

export default function registerAssetKpiRoutes(app, ctx) {
  const { buildAssetKpiRange, hasColumn, siteCodeFromReq, todayYYYYMMDD } = ctx;

  // GET /api/dashboard/asset-kpi/weekly?start=YYYY-MM-DD&end=YYYY-MM-DD&scheduled=10
  // Rolls up dashboard KPI rules (availability / utilization) across a date range per asset and by equipment category.
  app.get("/asset-kpi/weekly", async (req, reply) => {
    const startIn = String(req.query?.start || "").trim();
    const endIn = String(req.query?.end || "").trim();
    const scheduledRaw = Number(req.query?.scheduled ?? 10);
    const scheduledFallback = Number.isFinite(scheduledRaw) && scheduledRaw > 0 ? scheduledRaw : 10;

    const now = new Date();
    const dow = now.getDay();
    const mondayOffset = dow === 0 ? -6 : 1 - dow;
    const monday = new Date(now);
    monday.setHours(0, 0, 0, 0);
    monday.setDate(monday.getDate() + mondayOffset);
    const friday = new Date(monday);
    friday.setDate(friday.getDate() + 4);
    const ymd = (d) => d.toISOString().slice(0, 10);
    const start = /^\d{4}-\d{2}-\d{2}$/.test(startIn) ? startIn : ymd(monday);
    const end = /^\d{4}-\d{2}-\d{2}$/.test(endIn) ? endIn : ymd(friday);
    if (start > end) {
      return reply.code(400).send({ ok: false, error: "start must be <= end" });
    }

    const siteCode = siteCodeFromReq(req);
    const assetCodes = String(req.query?.asset_codes || "")
      .split(",")
      .map((c) => c.trim())
      .filter(Boolean);
    return reply.send(buildAssetKpiRange(start, end, scheduledFallback, siteCode, assetCodes));
  });

  // GET /api/dashboard/ldv-prestart/compliance?date=YYYY-MM-DD
  app.get("/ldv-prestart/compliance", async (req, reply) => {
    try {
      const date = String(req.query?.date || "").trim() || todayYYYYMMDD();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
        return reply.code(400).send({ ok: false, error: "date must be YYYY-MM-DD" });
      }
      const codes = Array.from({ length: 15 }, (_, i) => `V${String(i + 1).padStart(2, "0")}AM`);
      const marks = codes.map(() => "?").join(",");
      const assets = db.prepare(`
        SELECT id AS asset_id, asset_code, asset_name
        FROM assets
        WHERE UPPER(asset_code) IN (${marks})
          AND COALESCE(active, 1) = 1
      `).all(...codes);
      const hasVehicleChecks = Boolean(
        db.prepare(`SELECT 1 AS ok FROM sqlite_master WHERE type='table' AND name='vehicle_ldv_checks' LIMIT 1`).get()
      );
      if (!hasVehicleChecks) {
        return reply.send({
          ok: true,
          date,
          summary: { total: assets.length, compliant: 0, missing: assets.length, pct: 0 },
          rows: assets.map((a) => ({
            asset_id: Number(a.asset_id || 0),
            asset_code: String(a.asset_code || ""),
            asset_name: String(a.asset_name || ""),
            status: "missing",
            reason: "No pre-start table",
          })),
        });
      }
      const hasCheckMode = hasColumn("vehicle_ldv_checks", "check_mode");
      const hasChecklistJson = hasColumn("vehicle_ldv_checks", "checklist_json");
      const checkRows = db.prepare(`
        SELECT
          v.id,
          v.asset_id,
          v.odometer_km,
          ${hasChecklistJson ? "v.checklist_json" : "NULL AS checklist_json"}
        FROM vehicle_ldv_checks v
        JOIN (
          SELECT asset_id, MAX(id) AS max_id
          FROM vehicle_ldv_checks
          WHERE check_date = ?
            ${hasCheckMode ? "AND COALESCE(check_mode, 'ldv_general') = 'prestart'" : ""}
          GROUP BY asset_id
        ) m ON m.asset_id = v.asset_id AND m.max_id = v.id
      `).all(date);
      const checkMap = new Map(checkRows.map((r) => [Number(r.asset_id || 0), r]));
      const rows = assets.map((a) => {
        const aid = Number(a.asset_id || 0);
        const c = checkMap.get(aid);
        if (!c) {
          return {
            asset_id: aid,
            asset_code: String(a.asset_code || ""),
            asset_name: String(a.asset_name || ""),
            status: "missing",
            reason: "No pre-start submitted",
          };
        }
        const odometerOk = c.odometer_km != null && Number.isFinite(Number(c.odometer_km));
        let checklistOk = true;
        if (hasChecklistJson) {
          try {
            const raw = c.checklist_json ? JSON.parse(String(c.checklist_json || "{}")) : {};
            const requiredKeys = ["brakes_ok", "lights_ok", "tyres_ok", "oil_coolant_ok", "leaks_damage_ok", "safety_items_ok"];
            checklistOk = requiredKeys.every((k) => raw && typeof raw === "object" && raw[k] === true);
          } catch {
            checklistOk = false;
          }
        }
        const compliant = odometerOk && checklistOk;
        return {
          asset_id: aid,
          asset_code: String(a.asset_code || ""),
          asset_name: String(a.asset_name || ""),
          status: compliant ? "compliant" : "attention",
          reason: compliant ? "Pre-start complete" : (!odometerOk ? "Missing odometer" : "Checklist incomplete"),
        };
      }).sort((x, y) => String(x.asset_code || "").localeCompare(String(y.asset_code || "")));
      const compliant = rows.filter((r) => r.status === "compliant").length;
      const total = rows.length;
      const missing = Math.max(0, total - compliant);
      const pct = total > 0 ? Number(((compliant / total) * 100).toFixed(1)) : 0;
      return reply.send({
        ok: true,
        date,
        summary: { total, compliant, missing, pct },
        rows,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/asset-kpi.xlsx", async (req, reply) => {
    const startIn = String(req.query?.start || "").trim();
    const endIn = String(req.query?.end || "").trim();
    const scheduledRaw = Number(req.query?.scheduled ?? 10);
    const scheduledFallback = Number.isFinite(scheduledRaw) && scheduledRaw > 0 ? scheduledRaw : 10;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(startIn) || !/^\d{4}-\d{2}-\d{2}$/.test(endIn) || startIn > endIn) {
      return reply.code(400).send({ ok: false, error: "Provide valid start/end dates" });
    }
    const siteCode = siteCodeFromReq(req);
    const assetCodes = String(req.query?.asset_codes || "")
      .split(",")
      .map((c) => c.trim())
      .filter(Boolean);
    const data = buildAssetKpiRange(startIn, endIn, scheduledFallback, siteCode, assetCodes);
    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const summary = wb.addWorksheet("Summary");
    summary.addRow(["Asset KPI Report"]);
    summary.addRow(["Period", `${startIn} to ${endIn}`]);
    summary.addRow(["Scheduled fallback", scheduledFallback]);
    summary.addRow(["Equipment", assetCodes.length ? assetCodes.join(", ") : "All assets"]);
    summary.addRow([]);
    summary.addRow(["Metric", "Value"]);
    summary.addRow(["Scheduled Hours", data.fleet.scheduled_hours]);
    summary.addRow(["Available Hours", data.fleet.available_hours]);
    summary.addRow(["Run Hours", data.fleet.run_hours]);
    summary.addRow(["Downtime Hours", data.fleet.downtime_hours]);
    summary.addRow(["Availability %", data.fleet.availability_pct]);
    summary.addRow(["Utilization %", data.fleet.utilization_pct]);

    const cat = wb.addWorksheet("By Category");
    cat.addRow(["Category", "Assets", "Scheduled h", "Available h", "Run h", "Downtime h", "Availability %", "Utilization %"]);
    data.by_category.forEach((r) => cat.addRow([r.category, r.asset_count, r.scheduled_hours, r.available_hours, r.run_hours, r.downtime_hours, r.availability_pct, r.utilization_pct]));

    const assets = wb.addWorksheet("By Asset");
    assets.addRow([
      "Asset",
      "Name",
      "Category",
      "Department",
      "Cost center",
      "Data owner",
      "Mode",
      "Days with data",
      "Scheduled h",
      "Available h",
      "Run h",
      "Downtime h",
      "Availability %",
      "Utilization %",
    ]);
    data.by_asset.forEach((r) =>
      assets.addRow([
        r.asset_code,
        r.asset_name,
        r.category,
        r.department_code ?? "",
        r.cost_center_code ?? "",
        r.data_owner_username ?? "",
        r.utilization_mode,
        r.days_with_data,
        r.scheduled_hours,
        r.available_hours,
        r.run_hours,
        r.downtime_hours,
        r.availability_pct,
        r.utilization_pct,
      ])
    );

    const daily = wb.addWorksheet("Daily Trend");
    daily.addRow(["Date", "Scheduled h", "Available h", "Run h", "Downtime h", "Availability %", "Utilization %"]);
    data.daily_series.forEach((r) => {
      const availPct = r.scheduled_hours > 0 ? Number(((r.available_hours / r.scheduled_hours) * 100).toFixed(1)) : null;
      const utilPct = r.scheduled_hours > 0 ? Number(((r.run_hours / r.scheduled_hours) * 100).toFixed(1)) : null;
      daily.addRow([r.date, r.scheduled_hours, r.available_hours, r.run_hours, r.downtime_hours, availPct, utilPct]);
    });

    const assetDaily = wb.addWorksheet("Asset Daily Trend");
    assetDaily.addRow(["Asset", "Date", "Scheduled h", "Available h", "Run h", "Downtime h", "Availability %", "Utilization %"]);
    data.by_asset.forEach((asset) => {
      (asset.daily_points || []).forEach((p) => {
        const availPct = p.scheduled_hours > 0 ? Number(((p.available_hours / p.scheduled_hours) * 100).toFixed(1)) : null;
        const utilPct = p.scheduled_hours > 0 ? Number(((p.run_hours / p.scheduled_hours) * 100).toFixed(1)) : null;
        assetDaily.addRow([asset.asset_code, p.date, p.scheduled_hours, p.available_hours, p.run_hours, p.downtime_hours, availPct, utilPct]);
      });
    });

    // Month-by-month KPI from January of the end date's year through the end month.
    const monthRangesForYearUpTo = (endDate) => {
      const year = Number(endDate.slice(0, 4));
      const endMonth = Number(endDate.slice(5, 7));
      const ranges = [];
      for (let m = 1; m <= endMonth; m += 1) {
        const mm = String(m).padStart(2, "0");
        const first = `${year}-${mm}-01`;
        const lastDay = new Date(Date.UTC(year, m, 0)).getUTCDate();
        let last = `${year}-${mm}-${String(lastDay).padStart(2, "0")}`;
        if (last > endDate) last = endDate;
        ranges.push({ month: `${year}-${mm}`, start: first, end: last });
      }
      return ranges;
    };
    const monthLabels = ["", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    const monthNameOf = (ym) => `${monthLabels[Number(ym.slice(5, 7))] || ym} ${ym.slice(0, 4)}`;

    const monthly = wb.addWorksheet("Monthly KPI");
    monthly.addRow(["Month", "Scheduled h", "Available h", "Run h", "Downtime h", "Availability %", "Utilization %"]);
    const monthlyByAsset = wb.addWorksheet("Monthly by Asset");
    monthlyByAsset.addRow(["Month", "Asset", "Name", "Category", "Mode", "Scheduled h", "Available h", "Run h", "Downtime h", "Availability %", "Utilization %"]);

    monthRangesForYearUpTo(endIn).forEach((mr) => {
      const md = buildAssetKpiRange(mr.start, mr.end, scheduledFallback, siteCode, assetCodes);
      const label = monthNameOf(mr.month);
      monthly.addRow([
        label,
        md.fleet.scheduled_hours,
        md.fleet.available_hours,
        md.fleet.run_hours,
        md.fleet.downtime_hours,
        md.fleet.availability_pct,
        md.fleet.utilization_pct,
      ]);
      md.by_asset.forEach((a) => {
        monthlyByAsset.addRow([
          label,
          a.asset_code,
          a.asset_name,
          a.category,
          a.utilization_mode,
          a.scheduled_hours,
          a.available_hours,
          a.run_hours,
          a.downtime_hours,
          a.availability_pct,
          a.utilization_pct,
        ]);
      });
    });

    [summary, cat, assets, daily, assetDaily, monthly, monthlyByAsset].forEach((ws) => {
      ws.columns?.forEach((col) => { col.width = Math.min(24, Math.max(12, (col.header ? String(col.header).length : 12) + 2)); });
    });

    const buffer = await wb.xlsx.writeBuffer();
    return reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Asset_KPI_${startIn}_to_${endIn}.xlsx"`)
      .send(buffer);
  });

  // GET /api/dashboard/asset-kpi.docx?start=&end=&scheduled=&avail_target=&util_target=&view=category|asset
  app.get("/asset-kpi.docx", async (req, reply) => {
    const startIn = String(req.query?.start || "").trim();
    const endIn   = String(req.query?.end   || "").trim();
    const scheduledRaw = Number(req.query?.scheduled ?? 10);
    const scheduledFallback = Number.isFinite(scheduledRaw) && scheduledRaw > 0 ? scheduledRaw : 10;
    const availTarget = req.query?.avail_target != null ? Number(req.query.avail_target) : 85;
    const utilTarget  = req.query?.util_target  != null ? Number(req.query.util_target)  : 75;
    const viewMode    = String(req.query?.view || "category");
    const categoryFilter = String(req.query?.category || "").trim();

    if (!/^\d{4}-\d{2}-\d{2}$/.test(startIn) || !/^\d{4}-\d{2}-\d{2}$/.test(endIn) || startIn > endIn) {
      return reply.code(400).send({ ok: false, error: "Provide valid start/end dates" });
    }
    const siteCode = siteCodeFromReq(req);
    const assetCodes = String(req.query?.asset_codes || "").split(",").map(c => c.trim()).filter(Boolean);
    const kpi = buildAssetKpiRange(startIn, endIn, scheduledFallback, siteCode, assetCodes);

    // Fetch breakdowns in range for downtime reasons
    const breakdownRows = db.prepare(`
      SELECT b.breakdown_date, b.description, b.component, b.downtime_total_hours, b.status,
             a.asset_code, a.asset_name, a.category
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      WHERE b.breakdown_date >= ? AND b.breakdown_date <= ?
        AND (b.downtime_total_hours > 0 OR b.status = 'OPEN')
      ORDER BY b.downtime_total_hours DESC NULLS LAST, b.breakdown_date DESC
      LIMIT 300
    `).all(startIn, endIn);

    const fmt1 = v => v == null ? "—" : Number(v).toFixed(1);
    const fmtPct = v => v == null ? "—" : `${Number(v).toFixed(1)}%`;
    const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: "D1D5DB" };
    const cellBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };

    const headerCell = (text, bgColor = "2563EB") => new TableCell({
      children: [new Paragraph({ children: [new TextRun({ text: String(text), bold: true, color: "FFFFFF", size: 20 })], alignment: AlignmentType.CENTER })],
      shading: { fill: bgColor },
      borders: cellBorders,
    });
    const dataCell = (text, right = false, bold = false, color = "111827") => new TableCell({
      children: [new Paragraph({ children: [new TextRun({ text: String(text ?? "—"), bold, color, size: 18 })], alignment: right ? AlignmentType.RIGHT : AlignmentType.LEFT })],
      borders: cellBorders,
    });
    const pctCell = (val, target) => {
      const pct = val == null ? null : Number(val);
      const isBad = pct != null && target != null && pct < target;
      return new TableCell({
        children: [new Paragraph({ children: [new TextRun({ text: fmtPct(val), bold: isBad, color: isBad ? "DC2626" : "111827", size: 18 })], alignment: AlignmentType.RIGHT })],
        borders: cellBorders,
        shading: isBad ? { fill: "FEF2F2" } : undefined,
      });
    };

    // Build rows for chart table
    const allAssets = Array.isArray(kpi.by_asset) ? kpi.by_asset : [];
    const filtered = categoryFilter
      ? allAssets.filter(a => String(a.category || "").trim() === categoryFilter)
      : allAssets;
    const chartRows = viewMode === "asset"
      ? filtered.map(a => ({ label: `${a.asset_code} ${a.asset_name}`.trim(), availability_pct: a.availability_pct, utilization_pct: a.utilization_pct, downtime_hours: a.downtime_hours }))
      : (() => {
          const cats = categoryFilter
            ? (() => {
                const catMap = new Map();
                for (const a of filtered) {
                  const k = String(a.category || "Uncategorized").trim();
                  if (!catMap.has(k)) catMap.set(k, { label: k, sched: 0, avail: 0, run: 0, down: 0 });
                  const c = catMap.get(k);
                  c.sched += a.scheduled_hours; c.avail += a.available_hours; c.run += a.run_hours; c.down += a.downtime_hours;
                }
                return [...catMap.values()].map(c => ({ label: c.label, availability_pct: c.sched > 0 ? Number(((c.avail / c.sched) * 100).toFixed(1)) : null, utilization_pct: c.sched > 0 ? Number(((c.run / c.sched) * 100).toFixed(1)) : null, downtime_hours: c.down }));
              })()
            : (Array.isArray(kpi.by_category) ? kpi.by_category : []).map(c => ({ label: c.category, availability_pct: c.availability_pct, utilization_pct: c.utilization_pct, downtime_hours: c.downtime_hours }));
          return cats;
        })();

    // Group breakdowns by asset/category for reasons section
    const reasonsByCategory = new Map();
    for (const b of breakdownRows) {
      const cat = viewMode === "asset"
        ? `${b.asset_code} — ${b.asset_name}`
        : String(b.category || "Uncategorized").trim();
      if (categoryFilter && viewMode !== "asset" && cat !== categoryFilter) continue;
      if (!reasonsByCategory.has(cat)) reasonsByCategory.set(cat, []);
      reasonsByCategory.get(cat).push(b);
    }

    const children = [];

    // Title
    children.push(new Paragraph({ text: "Asset KPI Report", heading: HeadingLevel.HEADING_1 }));
    children.push(new Paragraph({ children: [new TextRun({ text: `Period: ${startIn} to ${endIn}`, size: 20, color: "64748B" })], spacing: { after: 120 } }));
    children.push(new Paragraph({ children: [new TextRun({ text: `Targets — Availability: ${availTarget}%  |  Utilization: ${utilTarget}%`, size: 20, color: "64748B" })], spacing: { after: 240 } }));

    // Fleet summary
    const fleet = kpi.fleet || {};
    children.push(new Paragraph({ text: "Fleet Summary", heading: HeadingLevel.HEADING_2 }));
    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [headerCell("Scheduled h"), headerCell("Available h"), headerCell("Run h"), headerCell("Downtime h"), headerCell("Availability %"), headerCell("Utilization %")] }),
        new TableRow({ children: [dataCell(fmt1(fleet.scheduled_hours), true), dataCell(fmt1(fleet.available_hours), true), dataCell(fmt1(fleet.run_hours), true), dataCell(fmt1(fleet.downtime_hours), true), pctCell(fleet.availability_pct, availTarget), pctCell(fleet.utilization_pct, utilTarget)] }),
      ],
    }));
    children.push(new Paragraph({ text: "", spacing: { after: 200 } }));

    // KPI by category / asset table
    children.push(new Paragraph({ text: viewMode === "asset" ? "KPI per Asset" : "KPI by Equipment Type", heading: HeadingLevel.HEADING_2 }));
    const kpiTableRows = [
      new TableRow({ children: [headerCell("Type / Asset"), headerCell("Downtime h"), headerCell("Availability %"), headerCell("Utilization %")] }),
      ...chartRows.map(r => new TableRow({ children: [
        dataCell(r.label),
        dataCell(fmt1(r.downtime_hours), true),
        pctCell(r.availability_pct, availTarget),
        pctCell(r.utilization_pct, utilTarget),
      ]})),
    ];
    children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: kpiTableRows }));
    children.push(new Paragraph({ text: "", spacing: { after: 200 } }));

    // Downtime reasons
    children.push(new Paragraph({ text: "Downtime Reasons by Equipment Type", heading: HeadingLevel.HEADING_2 }));
    children.push(new Paragraph({ children: [new TextRun({ text: "Breakdowns with recorded downtime in this period, grouped by equipment type.", size: 18, color: "64748B" })], spacing: { after: 160 } }));

    if (reasonsByCategory.size === 0) {
      children.push(new Paragraph({ children: [new TextRun({ text: "No breakdowns with downtime recorded in this period.", color: "64748B", size: 18 })] }));
    } else {
      for (const [cat, bdRows] of [...reasonsByCategory.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
        children.push(new Paragraph({ text: cat, heading: HeadingLevel.HEADING_3, spacing: { before: 160 } }));
        const bdTableRows = [
          new TableRow({ children: [headerCell("Date", "374151"), headerCell("Asset", "374151"), headerCell("Component", "374151"), headerCell("Description", "374151"), headerCell("Downtime h", "374151"), headerCell("Status", "374151")] }),
          ...bdRows.slice(0, 20).map(b => new TableRow({ children: [
            dataCell(b.breakdown_date || "—"),
            dataCell(`${b.asset_code} ${b.asset_name}`),
            dataCell(b.component || "—"),
            dataCell(b.description || "—"),
            dataCell(fmt1(b.downtime_total_hours), true),
            dataCell(b.status || "—"),
          ]})),
        ];
        children.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: bdTableRows }));
      }
    }

    const doc = new Document({ sections: [{ children }] });
    const buffer = await Packer.toBuffer(doc);
    return reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
      .header("Content-Disposition", `attachment; filename="IRONLOG_Asset_KPI_${startIn}_to_${endIn}.docx"`)
      .send(buffer);
  });
}
