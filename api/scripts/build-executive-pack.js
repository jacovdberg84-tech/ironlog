import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";

function parseArgs(argv) {
  const out = {
    start: "",
    end: "",
    baseUrl: process.env.IRONLOG_API_BASE_URL || "http://127.0.0.1:3001",
    siteCode: process.env.IRONLOG_SITE_CODE || "main",
    scheduledHours: Number(process.env.IRONLOG_SCHEDULED_HOURS || 10),
    nearDueHours: Number(process.env.IRONLOG_NEAR_DUE_HOURS || 50),
    outPath: "",
    timeoutMs: Number(process.env.IRONLOG_PACK_TIMEOUT_MS || 20000),
  };

  for (let i = 2; i < argv.length; i += 1) {
    const a = String(argv[i] || "");
    const next = String(argv[i + 1] || "");
    if (a === "--start" && next) out.start = next;
    if (a === "--end" && next) out.end = next;
    if (a === "--base-url" && next) out.baseUrl = next;
    if (a === "--site-code" && next) out.siteCode = next.toLowerCase();
    if (a === "--scheduled-hours" && next) out.scheduledHours = Number(next);
    if (a === "--near-due-hours" && next) out.nearDueHours = Number(next);
    if (a === "--out" && next) out.outPath = next;
    if (a === "--timeout-ms" && next) out.timeoutMs = Number(next);
  }
  return out;
}

function isYmd(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function safeNum(v, fallback = 0) {
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

function asArray(v) {
  return Array.isArray(v) ? v : [];
}

async function fetchJson(baseUrl, routePath, siteCode, timeoutMs) {
  const url = `${String(baseUrl).replace(/\/+$/, "")}${routePath}`;
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), Math.max(1000, timeoutMs));
  try {
    const res = await fetch(url, {
      headers: { "x-site-code": siteCode },
      signal: ctrl.signal,
    });
    const text = await res.text();
    let json = {};
    try {
      json = text ? JSON.parse(text) : {};
    } catch {
      json = {};
    }
    if (!res.ok) {
      throw new Error(`${routePath} -> HTTP ${res.status} ${json?.error || res.statusText}`);
    }
    return json;
  } finally {
    clearTimeout(timer);
  }
}

async function fetchSafe(baseUrl, routePath, siteCode, timeoutMs, fallback = {}) {
  try {
    return await fetchJson(baseUrl, routePath, siteCode, timeoutMs);
  } catch (err) {
    console.warn(`[warn] ${err?.message || err}`);
    return fallback;
  }
}

function autosizeColumns(ws, maxWidth = 45) {
  ws.columns.forEach((col) => {
    let width = 10;
    col.eachCell({ includeEmpty: true }, (cell) => {
      const len = String(cell.value ?? "").length;
      width = Math.max(width, Math.min(maxWidth, len + 2));
    });
    col.width = width;
  });
}

function writeRows(ws, headers, rows) {
  ws.addRow(headers);
  ws.getRow(1).font = { bold: true };
  for (const row of rows) ws.addRow(headers.map((h) => row[h]));
  ws.views = [{ state: "frozen", ySplit: 1 }];
  ws.autoFilter = { from: "A1", to: `${String.fromCharCode(64 + headers.length)}1` };
  autosizeColumns(ws);
}

function addNamedRange(workbook, name, ws, startRow, endRow, endCol) {
  if (endRow < startRow) return;
  const ref = `'${ws.name}'!$A$${startRow}:$${endCol}$${endRow}`;
  workbook.definedNames.add(name, ref);
}

function ymdNow() {
  return new Date().toISOString().slice(0, 10);
}

function monthStart(ymd) {
  return `${String(ymd).slice(0, 7)}-01`;
}

async function main() {
  const opts = parseArgs(process.argv);
  const end = isYmd(opts.end) ? opts.end : ymdNow();
  const start = isYmd(opts.start) ? opts.start : monthStart(end);
  if (!isYmd(start) || !isYmd(end)) throw new Error("Use --start and --end as YYYY-MM-DD");
  if (start > end) throw new Error("start must be <= end");

  const outPath = opts.outPath
    ? path.resolve(opts.outPath)
    : path.resolve(process.cwd(), "..", "exports", "ExecutivePackData.xlsx");
  fs.mkdirSync(path.dirname(outPath), { recursive: true });

  const qCommon = `start=${encodeURIComponent(start)}&end=${encodeURIComponent(end)}`;
  const kpi = await fetchSafe(opts.baseUrl, `/api/dashboard/asset-kpi/weekly?${qCommon}&scheduled=${encodeURIComponent(String(opts.scheduledHours || 10))}`, opts.siteCode, opts.timeoutMs, {});
  const fuel = await fetchSafe(opts.baseUrl, `/api/dashboard/fuel?${qCommon}&tolerance=0.15`, opts.siteCode, opts.timeoutMs, {});
  const fuelDup = await fetchSafe(opts.baseUrl, `/api/dashboard/fuel/duplicates?${qCommon}`, opts.siteCode, opts.timeoutMs, {});
  const wfSummary = await fetchSafe(opts.baseUrl, `/api/maintenance/weekly-forum/summary?${qCommon}&near_due_hours=${encodeURIComponent(String(opts.nearDueHours || 50))}`, opts.siteCode, opts.timeoutMs, {});
  const wfActions = await fetchSafe(opts.baseUrl, `/api/maintenance/weekly-forum/actions?${qCommon}`, opts.siteCode, opts.timeoutMs, {});
  const lube = await fetchSafe(opts.baseUrl, `/api/dashboard/lube/analytics?${qCommon}`, opts.siteCode, opts.timeoutMs, {});
  const mgrIns = await fetchSafe(opts.baseUrl, `/api/maintenance/inspections?${qCommon}`, opts.siteCode, opts.timeoutMs, {});
  const artIns = await fetchSafe(opts.baseUrl, `/api/maintenance/artisan-inspections?${qCommon}`, opts.siteCode, opts.timeoutMs, {});
  const stock = await fetchSafe(opts.baseUrl, "/api/stock/monitor", opts.siteCode, opts.timeoutMs, {});

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "IRONLOG";
  workbook.created = new Date();

  // 00_Control
  {
    const ws = workbook.addWorksheet("00_Control");
    const rows = [
      { key: "start_date", value: start },
      { key: "end_date", value: end },
      { key: "site_code", value: opts.siteCode },
      { key: "generated_at", value: new Date().toISOString() },
      { key: "base_url", value: opts.baseUrl },
    ];
    writeRows(ws, ["key", "value"], rows);
    workbook.definedNames.add("ctl_start_date", `'00_Control'!$B$2`);
    workbook.definedNames.add("ctl_end_date", `'00_Control'!$B$3`);
    workbook.definedNames.add("ctl_site_code", `'00_Control'!$B$4`);
    workbook.definedNames.add("ctl_generated_at", `'00_Control'!$B$5`);
  }

  // 01_HSE_Summary (inspection-driven proxy)
  {
    const ws = workbook.addWorksheet("01_HSE_Summary");
    const mgrRows = asArray(mgrIns.rows);
    const artRows = asArray(artIns.rows);
    const totalInspections = mgrRows.length + artRows.length;
    const mgrFindings = mgrRows.reduce((acc, r) => acc + asArray(r.checklist).filter((c) => c?.ok === false).length, 0);
    const artFindings = artRows.reduce((acc, r) => acc + asArray(r.checklist).filter((c) => c?.ok === false).length, 0);
    const rows = [{
      period_start: start,
      period_end: end,
      manager_inspections: mgrRows.length,
      artisan_inspections: artRows.length,
      inspections_total: totalInspections,
      findings_open_proxy: mgrFindings + artFindings,
      notes: "Use dedicated HSE table if available in your deployment.",
    }];
    const headers = Object.keys(rows[0]);
    writeRows(ws, headers, rows);
    addNamedRange(workbook, "rng_hse_summary", ws, 1, 2, "G");
  }

  // 03_Plant_Performance
  {
    const ws = workbook.addWorksheet("03_Plant_Performance");
    const rows = asArray(kpi.by_asset).map((r) => ({
      asset_code: r.asset_code || "",
      asset_name: r.asset_name || "",
      category: r.category || "",
      scheduled_hours: safeNum(r.scheduled_hours),
      run_hours: safeNum(r.run_hours),
      downtime_hours: safeNum(r.downtime_hours),
      available_hours: safeNum(r.available_hours),
      availability_pct: safeNum(r.availability_pct, null),
      utilization_pct: safeNum(r.utilization_pct, null),
      days_with_data: safeNum(r.days_with_data),
      days_in_range: safeNum(r.days_in_range),
    }));
    const headers = rows.length ? Object.keys(rows[0]) : ["asset_code", "asset_name", "category", "scheduled_hours", "run_hours", "downtime_hours", "available_hours", "availability_pct", "utilization_pct", "days_with_data", "days_in_range"];
    writeRows(ws, headers, rows);
    addNamedRange(workbook, "rng_plant_perf_table", ws, 1, rows.length + 1, "K");
  }

  // 04_Maintenance_Cost_Per_Machine
  {
    const ws = workbook.addWorksheet("04_Maint_Cost_Machine");
    const rows = asArray(wfSummary.upcoming_services).map((r) => ({
      asset_code: r.asset_code || "",
      asset_name: r.asset_name || "",
      service_name: r.service_name || "",
      remaining_hours: safeNum(r.remaining_hours),
      avg_oil_cost: safeNum(r?.forecast?.avg_oil_cost),
      avg_parts_cost: safeNum(r?.forecast?.avg_parts_cost),
      est_service_kit_cost: safeNum(r?.forecast?.est_service_kit_cost),
    }));
    const headers = rows.length ? Object.keys(rows[0]) : ["asset_code", "asset_name", "service_name", "remaining_hours", "avg_oil_cost", "avg_parts_cost", "est_service_kit_cost"];
    writeRows(ws, headers, rows);
    addNamedRange(workbook, "rng_breakdown_cost_machine", ws, 1, rows.length + 1, "G");
  }

  // 05_Parts_Tracking
  {
    const ws = workbook.addWorksheet("05_Parts_Tracking");
    const rows = asArray(stock.rows).map((r) => ({
      part_code: r.part_code || "",
      part_name: r.part_name || "",
      category: r.category || "",
      on_hand: safeNum(r.on_hand),
      min_stock: safeNum(r.min_stock),
      below_min_flag: safeNum(r.is_below_min),
      critical_flag: safeNum(r.is_critical),
      unit_cost: safeNum(r.unit_cost),
      stock_value: safeNum(r.stock_value),
    }));
    const headers = rows.length ? Object.keys(rows[0]) : ["part_code", "part_name", "category", "on_hand", "min_stock", "below_min_flag", "critical_flag", "unit_cost", "stock_value"];
    writeRows(ws, headers, rows);
    addNamedRange(workbook, "rng_parts_tracking_table", ws, 1, rows.length + 1, "I");
  }

  // 06_Production_Support
  {
    const ws = workbook.addWorksheet("06_Production_Support");
    const rows = asArray(wfActions.rows).map((r) => ({
      action_date: r.action_date || "",
      owner: r.owner || "",
      action_text: r.action_text || "",
      status: r.status || "",
      due_date: r.due_date || "",
      closed_date: r.closed_date || "",
      priority: r.priority || "",
    }));
    const headers = rows.length ? Object.keys(rows[0]) : ["action_date", "owner", "action_text", "status", "due_date", "closed_date", "priority"];
    writeRows(ws, headers, rows);
    addNamedRange(workbook, "rng_production_support_backlog", ws, 1, rows.length + 1, "G");
  }

  // 07_Lubrication_Cost_Per_Machine
  {
    const ws = workbook.addWorksheet("07_Lube_Cost_Machine");
    const rows = asArray(lube.rows).map((r) => ({
      asset_code: r.asset_code || "",
      asset_name: r.asset_name || "",
      lube_type: r.lube_type || "",
      qty_total: safeNum(r.qty_total),
      entries: safeNum(r.entries),
      total_lube_cost: safeNum(r.total_lube_cost),
    }));
    const headers = rows.length ? Object.keys(rows[0]) : ["asset_code", "asset_name", "lube_type", "qty_total", "entries", "total_lube_cost"];
    writeRows(ws, headers, rows);
    addNamedRange(workbook, "rng_lube_cost_machine", ws, 1, rows.length + 1, "F");
  }

  // 08_Inspections
  {
    const ws = workbook.addWorksheet("08_Inspections");
    const rows = [
      { inspection_type: "manager", completed_count: asArray(mgrIns.rows).length },
      { inspection_type: "artisan", completed_count: asArray(artIns.rows).length },
      { inspection_type: "total", completed_count: asArray(mgrIns.rows).length + asArray(artIns.rows).length },
    ];
    writeRows(ws, ["inspection_type", "completed_count"], rows);
    addNamedRange(workbook, "rng_inspections_completed", ws, 1, rows.length + 1, "B");
  }

  // 09_Fuel_Security_Anomalies
  {
    const ws = workbook.addWorksheet("09_Fuel_Security");
    const rows = asArray(fuel.rows).map((r) => ({
      asset_code: r.asset_code || "",
      asset_name: r.asset_name || "",
      metric_mode: r.metric_mode || "",
      fuel_liters: safeNum(r.fuel_liters),
      hours_run: safeNum(r.hours_run),
      km_run: safeNum(r.km_run),
      actual_lph: r.actual_lph == null ? null : safeNum(r.actual_lph),
      oem_lph: r.oem_lph == null ? null : safeNum(r.oem_lph),
      variance_lph: r.variance_lph == null ? null : safeNum(r.variance_lph),
      is_excessive: r.is_excessive ? 1 : 0,
      duplicate_rows: asArray(fuelDup.rows).filter((d) => String(d.asset_code || "") === String(r.asset_code || "")).length,
    }));
    const headers = rows.length ? Object.keys(rows[0]) : ["asset_code", "asset_name", "metric_mode", "fuel_liters", "hours_run", "km_run", "actual_lph", "oem_lph", "variance_lph", "is_excessive", "duplicate_rows"];
    writeRows(ws, headers, rows);
    addNamedRange(workbook, "rng_fuel_security_anomalies", ws, 1, rows.length + 1, "K");
  }

  // 10_Slide_Map
  {
    const ws = workbook.addWorksheet("10_Slide_Map");
    const rows = [
      { slide_no: 1, slide_title: "Safety (HSE)", named_range: "rng_hse_summary", owner: "HSE" },
      { slide_no: 2, slide_title: "Plant Performance", named_range: "rng_plant_perf_table", owner: "Maintenance" },
      { slide_no: 3, slide_title: "Breakdown & Maintenance (Cost/Machine)", named_range: "rng_breakdown_cost_machine", owner: "Maintenance" },
      { slide_no: 4, slide_title: "Parts Tracking", named_range: "rng_parts_tracking_table", owner: "Stores" },
      { slide_no: 5, slide_title: "Production Support", named_range: "rng_production_support_backlog", owner: "Operations" },
      { slide_no: 6, slide_title: "Lubrication (Cost/Machine)", named_range: "rng_lube_cost_machine", owner: "Maintenance" },
      { slide_no: 7, slide_title: "Inspections Done", named_range: "rng_inspections_completed", owner: "Maintenance" },
      { slide_no: 8, slide_title: "Security (Fuel Anomalies)", named_range: "rng_fuel_security_anomalies", owner: "Security" },
    ];
    writeRows(ws, ["slide_no", "slide_title", "named_range", "owner"], rows);
  }

  await workbook.xlsx.writeFile(outPath);
  console.log(`[ok] Executive pack written: ${outPath}`);
  console.log(`[ok] Range: ${start} to ${end} | site: ${opts.siteCode}`);
}

main().catch((err) => {
  console.error(`[error] ${err?.message || err}`);
  process.exit(1);
});
