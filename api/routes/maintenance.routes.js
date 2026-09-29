// IRONLOG/api/routes/maintenance.routes.js
import { db } from "../db/client.js";
import multipart from "@fastify/multipart";
import fs from "node:fs";
import path from "node:path";
import ExcelJS from "exceljs";
import { ensurePageSpace, pdfBodyBottom, pdfBodyTop, table } from "../utils/pdfGenerator.js";
import { getPdfReportBranding, getReportSetting } from "../utils/reportSettings.js";
import { ensureAuditTable } from "../utils/audit.js";
import {
  buildReliabilityIncidentsForAssets,
  computeMtbfLttr,
  round2,
} from "../utils/reliabilityMetrics.js";
import { getAssetKpiRangeBuilder } from "../utils/assetKpiRangeProvider.js";
import { resolveStorageAbs as resolveStorageAbsPath, getDataRoot } from "../utils/storagePaths.js";
import {
  enrichPlansWithNextService,
  groupActivePlansByAsset,
  snapLastServiceHours,
  planIntervalHours,
  classifyServiceDueStatus,
  meterUnitForAsset,
} from "../utils/serviceSchedule.js";
import {
  buildUndercarriageComponentSchema,
  enrichUndercarriageMeasurement,
  normalizeUndercarriageMeasurements,
  normalizeUndercarriageWearLimits,
  countConfiguredWearLimits,
  UNDERCARRIAGE_CHECKLIST_ITEMS,
  UNDERCARRIAGE_TRACK_SAG_POINTS,
  UNDERCARRIAGE_WEAR_BANDS,
} from "../utils/undercarriageTemplate.js";
import { SERVICE_TEMPLATE_ITEM_TYPES, buildServiceEstimatePreview, ensureServiceTemplateSchema } from "../utils/serviceTemplates.js";
import { isDate } from "../utils/request.js";
import registerServiceTemplatesRoutes from "./maintenance/service-templates.routes.js";
import registerPlansRoutes from "./maintenance/plans.routes.js";
import registerInsightsRoutes from "./maintenance/insights.routes.js";
import registerWeeklyRoutes from "./maintenance/weekly.routes.js";
import registerCostingRoutes from "./maintenance/costing.routes.js";
import registerInspectionsRoutes from "./maintenance/inspections.routes.js";
import registerPrestartChecksRoutes from "./maintenance/prestart-checks.routes.js";
import registerPartsRequestsRoutes from "./maintenance/parts-requests.routes.js";
import registerMechanicLaborRoutes from "./maintenance/mechanic-labor.routes.js";
import { ensureStockCategorySchema, oilPartSql as oilCategorySql } from "../utils/stockCategory.js";

function isMonth(s) {
  return /^\d{4}-\d{2}$/.test(String(s || "").trim());
}

function getMaintenanceRoles(req) {
  const many = String(req.headers["x-user-roles"] || "")
    .split(",")
    .map((x) => String(x || "").trim().toLowerCase())
    .filter(Boolean);
  const one = String(req.headers["x-user-role"] || "")
    .split(",")
    .map((x) => String(x || "").trim().toLowerCase())
    .filter(Boolean);
  return Array.from(new Set([...many, ...one]));
}

function requireMaintenanceRoles(req, reply, allowed) {
  const roles = getMaintenanceRoles(req);
  if (!roles.some((r) => allowed.includes(r))) {
    reply.code(403).send({ ok: false, error: "not allowed" });
    return false;
  }
  return true;
}

function addDaysYmd(ymd, days) {
  const d = new Date(`${String(ymd).trim()}T12:00:00`);
  d.setDate(d.getDate() + Number(days || 0));
  return d.toISOString().slice(0, 10);
}

function listDaysInclusiveYmd(start, end) {
  const out = [];
  let cur = String(start || "").trim();
  const endDay = String(end || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(cur) || !/^\d{4}-\d{2}-\d{2}$/.test(endDay)) return out;
  while (cur <= endDay) {
    out.push(cur);
    cur = addDaysYmd(cur, 1);
  }
  return out;
}

function mondayOfWeekYmd(ymd) {
  const d = new Date(`${String(ymd).trim()}T12:00:00`);
  const dow = d.getDay();
  const offset = dow === 0 ? -6 : 1 - dow;
  d.setDate(d.getDate() + offset);
  return d.toISOString().slice(0, 10);
}

function weeksOverlappingMonth(year, month) {
  const y = Number(year);
  const m = Number(month);
  if (!Number.isFinite(y) || !Number.isFinite(m) || m < 1 || m > 12) return [];
  const monthStart = `${y}-${String(m).padStart(2, "0")}-01`;
  const lastDay = new Date(y, m, 0).getDate();
  const monthEnd = `${y}-${String(m).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;
  const weeks = [];
  let ws = mondayOfWeekYmd(monthStart);
  for (let i = 0; i < 6; i++) {
    const we = addDaysYmd(ws, 6);
    if (we >= monthStart && ws <= monthEnd) weeks.push(ws);
    ws = addDaysYmd(ws, 7);
  }
  return weeks;
}

function weekdayShortYmd(ymd) {
  const s = String(ymd || "").trim();
  if (!isDate(s)) return "";
  return new Date(`${s}T12:00:00`).toLocaleDateString("en-US", { weekday: "short" });
}

function formatWiDayTitleYmd(ymd) {
  const s = String(ymd || "").trim();
  if (!isDate(s)) return s;
  const d = new Date(`${s}T12:00:00`);
  const day = d.toLocaleDateString("en-US", { weekday: "long" });
  return `${day} ${s}`;
}

function wiFormatMinutesPdf(mins) {
  const n = Math.max(0, Number(mins || 0));
  if (n < 60) return `${n} min`;
  const h = Math.floor(n / 60);
  const m = n % 60;
  return m ? `${h}h ${m}m` : `${h}h`;
}

function ensureWeeklyInspectionSchema() {
  const assetCols = db.prepare(`PRAGMA table_info(weekly_inspection_assets)`).all();
  if (!assetCols.some((c) => c.name === "est_minutes")) {
    db.prepare(`ALTER TABLE weekly_inspection_assets ADD COLUMN est_minutes INTEGER NOT NULL DEFAULT 30`).run();
  }
  const entryCols = db.prepare(`PRAGMA table_info(weekly_inspection_entries)`).all();
  if (!entryCols.some((c) => c.name === "released")) {
    db.prepare(`ALTER TABLE weekly_inspection_entries ADD COLUMN released INTEGER NOT NULL DEFAULT 0`).run();
  }
  db.prepare(`
    CREATE TABLE IF NOT EXISTS weekly_inspection_week_plan (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      week_start TEXT NOT NULL,
      asset_id INTEGER NOT NULL,
      planned_date TEXT NOT NULL,
      est_minutes INTEGER NOT NULL DEFAULT 30,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(week_start, asset_id),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS weekly_inspection_slots (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      planned_date TEXT NOT NULL,
      asset_id INTEGER NOT NULL,
      est_minutes INTEGER NOT NULL DEFAULT 30,
      status TEXT NOT NULL DEFAULT 'pending',
      inspector_name TEXT,
      notes TEXT,
      completed_at TEXT,
      released INTEGER NOT NULL DEFAULT 0,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(planned_date, asset_id),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();
  const slotCols = db.prepare(`PRAGMA table_info(weekly_inspection_slots)`).all();
  if (!slotCols.some((c) => c.name === "status")) {
    db.prepare(`ALTER TABLE weekly_inspection_slots ADD COLUMN status TEXT NOT NULL DEFAULT 'pending'`).run();
  }
  if (!slotCols.some((c) => c.name === "inspector_name")) {
    db.prepare(`ALTER TABLE weekly_inspection_slots ADD COLUMN inspector_name TEXT`).run();
  }
  if (!slotCols.some((c) => c.name === "notes")) {
    db.prepare(`ALTER TABLE weekly_inspection_slots ADD COLUMN notes TEXT`).run();
  }
  if (!slotCols.some((c) => c.name === "completed_at")) {
    db.prepare(`ALTER TABLE weekly_inspection_slots ADD COLUMN completed_at TEXT`).run();
  }
  if (!slotCols.some((c) => c.name === "released")) {
    db.prepare(`ALTER TABLE weekly_inspection_slots ADD COLUMN released INTEGER NOT NULL DEFAULT 0`).run();
  }
  migrateLegacyWeeklyInspectionPlansToSlots();
}

function migrateLegacyWeeklyInspectionPlansToSlots() {
  const slotCount = Number(db.prepare(`SELECT COUNT(*) AS c FROM weekly_inspection_slots`).get()?.c || 0);
  if (slotCount > 0) return;
  const legacy = db.prepare(`
    SELECT wp.asset_id, wp.planned_date, wp.est_minutes, wp.sort_order,
           e.status, e.inspector_name, e.notes, e.completed_at, COALESCE(e.released, 0) AS released
    FROM weekly_inspection_week_plan wp
    LEFT JOIN weekly_inspection_entries e
      ON e.asset_id = wp.asset_id AND e.week_start = wp.week_start
    WHERE TRIM(COALESCE(wp.planned_date, '')) != ''
  `).all();
  if (!legacy.length) return;
  const ins = db.prepare(`
    INSERT OR IGNORE INTO weekly_inspection_slots (
      planned_date, asset_id, est_minutes, status, inspector_name, notes, completed_at, released,
      sort_order, created_at, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
  `);
  for (const row of legacy) {
    const planned_date = String(row.planned_date || "").trim();
    if (!isDate(planned_date)) continue;
    const status = String(row.status || "pending").toLowerCase();
    ins.run(
      planned_date,
      Number(row.asset_id),
      Math.max(5, Number(row.est_minutes ?? 30) || 30),
      ["pending", "done", "skipped"].includes(status) ? status : "pending",
      row.inspector_name || null,
      row.notes || null,
      row.completed_at || null,
      status === "done" ? 1 : Number(row.released || 0),
      Number(row.sort_order || 0)
    );
  }
}

function normalizeEquipCategory(raw) {
  const s = String(raw || "").trim();
  return s || "Uncategorized";
}

function monthBoundsYmd(year, month) {
  const y = Number(year);
  const m = Number(month);
  const monthStart = `${y}-${String(m).padStart(2, "0")}-01`;
  const lastDay = new Date(y, m, 0).getDate();
  const monthEnd = `${y}-${String(m).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;
  return { monthStart, monthEnd, lastDay };
}

function buildMonthCalendarWeeks(year, month) {
  const { monthStart, monthEnd, lastDay } = monthBoundsYmd(year, month);
  const firstDow = new Date(`${monthStart}T12:00:00`).getDay();
  const startPad = (firstDow + 6) % 7;
  const cells = [];
  for (let i = 0; i < startPad; i++) cells.push({ date: null, in_month: false, day: null });
  for (let d = 1; d <= lastDay; d++) {
    const date = `${year}-${String(month).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    cells.push({ date, in_month: true, day: d });
  }
  while (cells.length % 7 !== 0) cells.push({ date: null, in_month: false, day: null });
  const weeks = [];
  for (let i = 0; i < cells.length; i += 7) weeks.push(cells.slice(i, i + 7));
  return { monthStart, monthEnd, weeks };
}

function loadWeeklyInspectionRosterAssets() {
  return db.prepare(`
    SELECT
      wia.id,
      wia.asset_id,
      wia.notes,
      wia.sort_order,
      wia.active,
      COALESCE(wia.est_minutes, 30) AS est_minutes,
      a.asset_code,
      a.asset_name,
      a.category
    FROM weekly_inspection_assets wia
    JOIN assets a ON a.id = wia.asset_id
    WHERE COALESCE(wia.active, 1) = 1
      AND COALESCE(a.active, 1) = 1
      AND COALESCE(a.archived, 0) = 0
    ORDER BY COALESCE(wia.sort_order, 0), a.asset_code ASC
  `).all();
}

function loadWeeklyInspectionSlotsBetween(monthStart, monthEnd) {
  return db.prepare(`
    SELECT
      s.id,
      s.planned_date,
      s.asset_id,
      s.est_minutes,
      s.status,
      s.inspector_name,
      s.notes,
      s.completed_at,
      COALESCE(s.released, 0) AS released,
      s.sort_order,
      a.asset_code,
      a.asset_name
    FROM weekly_inspection_slots s
    JOIN assets a ON a.id = s.asset_id
    WHERE s.planned_date >= ?
      AND s.planned_date <= ?
    ORDER BY s.planned_date ASC, COALESCE(s.sort_order, 0) ASC, a.asset_code ASC
  `).all(monthStart, monthEnd);
}

function upsertWeeklyInspectionRosterAsset(assetId, estMinutes, notes = "") {
  const est = Math.max(5, Number(estMinutes ?? 30) || 30);
  const existing = db.prepare(`SELECT id FROM weekly_inspection_assets WHERE asset_id = ?`).get(assetId);
  if (existing) {
    db.prepare(`
      UPDATE weekly_inspection_assets
      SET active = 1, est_minutes = ?, notes = COALESCE(NULLIF(?, ''), notes), updated_at = datetime('now')
      WHERE asset_id = ?
    `).run(est, notes, assetId);
    return Number(existing.id);
  }
  const maxSort = Number(db.prepare(`SELECT COALESCE(MAX(sort_order), 0) AS m FROM weekly_inspection_assets`).get()?.m || 0);
  const ins = db.prepare(`
    INSERT INTO weekly_inspection_assets (asset_id, notes, sort_order, est_minutes, active, created_at, updated_at)
    VALUES (?, ?, ?, ?, 1, datetime('now'), datetime('now'))
  `).run(assetId, notes, maxSort + 1, est);
  return Number(ins.lastInsertRowid);
}

function computeWeeklyInspectionComplianceFromSlots(assets, slots, year, month, todayYmd) {
  const today = String(todayYmd || new Date().toISOString().slice(0, 10));
  const { monthStart, monthEnd } = monthBoundsYmd(year, month);
  const weekStarts = weeksOverlappingMonth(year, month);
  const slotsByDate = {};

  let totalSlots = 0;
  let doneCount = 0;
  let pendingCount = 0;
  let skippedCount = 0;
  let notReleasedCount = 0;
  let estMinutesTotal = 0;
  let estMinutesReleased = 0;
  const byAsset = {};
  const dayAgendaMap = {};

  for (const s of slots || []) {
    const plannedDate = String(s.planned_date || "").trim();
    const assetId = Number(s.asset_id);
    if (!isDate(plannedDate) || !assetId) continue;
    const status = String(s.status || "pending").toLowerCase();
    const est = Math.max(5, Number(s.est_minutes ?? 30) || 30);
    if (!slotsByDate[plannedDate]) slotsByDate[plannedDate] = [];
    slotsByDate[plannedDate].push(s);

    if (!byAsset[assetId]) {
      byAsset[assetId] = {
        asset_id: assetId,
        asset_code: String(s.asset_code || ""),
        asset_name: String(s.asset_name || ""),
        total: 0,
        done: 0,
        pending: 0,
        skipped: 0,
        not_released: 0,
        score: 100,
      };
    }

    totalSlots += 1;
    estMinutesTotal += est;
    byAsset[assetId].total += 1;

    if (!dayAgendaMap[plannedDate]) {
      dayAgendaMap[plannedDate] = { date: plannedDate, est_minutes: 0, items: [] };
    }
    dayAgendaMap[plannedDate].est_minutes += est;
    dayAgendaMap[plannedDate].items.push({
      slot_id: Number(s.id),
      asset_id: assetId,
      asset_code: String(s.asset_code || ""),
      asset_name: String(s.asset_name || ""),
      est_minutes: est,
      status,
      released: status === "done",
    });

    if (status === "done") {
      doneCount += 1;
      estMinutesReleased += est;
      byAsset[assetId].done += 1;
    } else if (status === "skipped") {
      skippedCount += 1;
      byAsset[assetId].skipped += 1;
    } else {
      pendingCount += 1;
      byAsset[assetId].pending += 1;
      if (today > plannedDate) {
        notReleasedCount += 1;
        byAsset[assetId].not_released += 1;
      }
    }
  }

  const weekly_gaps = [];
  for (const asset of assets || []) {
    const assetId = Number(asset.asset_id);
    if (!assetId) continue;
    for (const weekStart of weekStarts) {
      const weekEnd = addDaysYmd(weekStart, 6);
      const touchesMonth = weekEnd >= monthStart && weekStart <= monthEnd;
      if (!touchesMonth) continue;
      const hasSlot = (slots || []).some((row) => {
        const d = String(row.planned_date || "");
        return Number(row.asset_id) === assetId && d >= weekStart && d <= weekEnd;
      });
      if (!hasSlot) {
        weekly_gaps.push({
          asset_id: assetId,
          asset_code: String(asset.asset_code || ""),
          asset_name: String(asset.asset_name || ""),
          week_start: weekStart,
          week_end: weekEnd,
        });
      }
    }
  }

  for (const row of Object.values(byAsset)) {
    const base = row.total ? (row.done / row.total) * 100 : 100;
    const penalty = row.not_released * 10;
    row.score = Math.max(0, Math.round(base - penalty));
  }

  const completionPct = totalSlots ? Math.round((doneCount / totalSlots) * 100) : 100;
  const penaltyPoints = notReleasedCount * 5;
  const complianceScore = Math.max(0, Math.round(completionPct - penaltyPoints));

  return {
    score: complianceScore,
    completion_pct: completionPct,
    total_slots: totalSlots,
    done_count: doneCount,
    pending_count: pendingCount,
    skipped_count: skippedCount,
    not_released_count: notReleasedCount,
    penalty_points: penaltyPoints,
    est_minutes_total: estMinutesTotal,
    est_minutes_released: estMinutesReleased,
    by_asset: Object.values(byAsset).sort((x, y) => x.score - y.score || String(x.asset_code).localeCompare(String(y.asset_code))),
    day_agenda: Object.values(dayAgendaMap).sort((x, y) => String(x.date).localeCompare(String(y.date))),
    weekly_gaps,
  };
}

function buildWeeklyInspectionCalendarData(query = {}) {
  ensureWeeklyInspectionSchema();
  const monthRaw = String(query.month || "").trim();
  let year;
  let month;
  if (isMonth(monthRaw)) {
    [year, month] = monthRaw.split("-").map(Number);
  } else {
    const anchor = isDate(String(query.week_start || "").trim())
      ? String(query.week_start).trim()
      : new Date().toISOString().slice(0, 10);
    const d = new Date(`${anchor}T12:00:00`);
    year = d.getFullYear();
    month = d.getMonth() + 1;
  }
  const { monthStart, monthEnd, weeks: calendarWeeks } = buildMonthCalendarWeeks(year, month);
  const assets = loadWeeklyInspectionRosterAssets();
  const slots = loadWeeklyInspectionSlotsBetween(monthStart, monthEnd);
  const slotsByDate = {};
  for (const s of slots) {
    const date = String(s.planned_date);
    if (!slotsByDate[date]) slotsByDate[date] = [];
    slotsByDate[date].push({
      id: Number(s.id),
      asset_id: Number(s.asset_id),
      asset_code: String(s.asset_code || ""),
      asset_name: String(s.asset_name || ""),
      est_minutes: Number(s.est_minutes || 30),
      status: String(s.status || "pending").toLowerCase(),
      inspector_name: String(s.inspector_name || ""),
      released: Number(s.released || 0),
    });
  }
  const today = new Date().toISOString().slice(0, 10);
  const calendar_weeks = calendarWeeks.map((row) => row.map((cell) => {
    if (!cell?.date) return { ...cell, is_today: false, slots: [] };
    return {
      ...cell,
      is_today: String(cell.date) === today,
      slots: slotsByDate[String(cell.date)] || [],
    };
  }));
  const compliance = computeWeeklyInspectionComplianceFromSlots(assets, slots, year, month, today);
  return {
    ok: true,
    month: `${year}-${String(month).padStart(2, "0")}`,
    year,
    month_num: month,
    month_start: monthStart,
    month_end: monthEnd,
    calendar_weeks,
    assets,
    slots,
    compliance,
  };
}

function addWeeklyInspectionSlot({ planned_date, asset_id, est_minutes }) {
  ensureWeeklyInspectionSchema();
  const date = String(planned_date || "").trim();
  const aid = Number(asset_id || 0);
  if (!isDate(date)) throw new Error("planned_date must be YYYY-MM-DD");
  if (!aid) throw new Error("asset_id is required");
  const roster = db.prepare(`
    SELECT COALESCE(est_minutes, 30) AS est_minutes
    FROM weekly_inspection_assets
    WHERE asset_id = ? AND COALESCE(active, 1) = 1
  `).get(aid);
  if (!roster) throw new Error("Add equipment to the workshop roster first");
  const est = Math.max(5, Number(est_minutes ?? roster.est_minutes ?? 30) || 30);
  const sortOrder = Number(db.prepare(`
    SELECT COALESCE(MAX(sort_order), 0) AS m
    FROM weekly_inspection_slots
    WHERE planned_date = ?
  `).get(date)?.m || 0) + 1;
  db.prepare(`
    INSERT INTO weekly_inspection_slots (
      planned_date, asset_id, est_minutes, sort_order, created_at, updated_at
    ) VALUES (?, ?, ?, ?, datetime('now'), datetime('now'))
    ON CONFLICT(planned_date, asset_id) DO UPDATE SET
      est_minutes = excluded.est_minutes,
      sort_order = excluded.sort_order,
      updated_at = datetime('now')
  `).run(date, aid, est, sortOrder);
  return db.prepare(`
    SELECT
      s.id, s.planned_date, s.asset_id, s.est_minutes, s.status, s.inspector_name,
      COALESCE(s.released, 0) AS released, a.asset_code, a.asset_name
    FROM weekly_inspection_slots s
    JOIN assets a ON a.id = s.asset_id
    WHERE s.planned_date = ? AND s.asset_id = ?
  `).get(date, aid);
}

function copyWeeklyInspectionDay(from_date, to_date) {
  ensureWeeklyInspectionSchema();
  const from = String(from_date || "").trim();
  const to = String(to_date || "").trim();
  if (!isDate(from)) throw new Error("from_date must be YYYY-MM-DD");
  if (!isDate(to)) throw new Error("to_date must be YYYY-MM-DD");
  if (from === to) throw new Error("Source and target dates must differ");

  const sourceSlots = db.prepare(`
    SELECT asset_id, est_minutes, sort_order
    FROM weekly_inspection_slots
    WHERE planned_date = ?
    ORDER BY COALESCE(sort_order, 0) ASC, asset_id ASC
  `).all(from);
  if (!sourceSlots.length) throw new Error("No equipment scheduled on the source day");

  const ins = db.prepare(`
    INSERT INTO weekly_inspection_slots (
      planned_date, asset_id, est_minutes, sort_order, status, released, created_at, updated_at
    ) VALUES (?, ?, ?, ?, 'pending', 0, datetime('now'), datetime('now'))
    ON CONFLICT(planned_date, asset_id) DO NOTHING
  `);

  let copied = 0;
  let skipped = 0;
  for (const row of sourceSlots) {
    const est = Math.max(5, Number(row.est_minutes ?? 30) || 30);
    const result = ins.run(to, Number(row.asset_id), est, Number(row.sort_order || 0));
    if (Number(result.changes || 0) > 0) copied += 1;
    else skipped += 1;
  }
  if (!copied && skipped) {
    throw new Error("All equipment from that day is already scheduled on the target date");
  }
  return { from_date: from, to_date: to, copied, skipped, total: sourceSlots.length };
}

function clearWeeklyInspectionRoster({ clear_slots = true } = {}) {
  ensureWeeklyInspectionSchema();
  const roster_cleared = Number(
    db.prepare(`SELECT COUNT(*) AS c FROM weekly_inspection_assets WHERE COALESCE(active, 1) = 1`).get()?.c || 0,
  );
  db.prepare(`DELETE FROM weekly_inspection_assets`).run();
  let slots_cleared = 0;
  if (clear_slots) {
    slots_cleared = Number(db.prepare(`SELECT COUNT(*) AS c FROM weekly_inspection_slots`).get()?.c || 0);
    db.prepare(`DELETE FROM weekly_inspection_slots`).run();
    // Prevent legacy week-plan rows from repopulating the calendar after a clear.
    if (db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name='weekly_inspection_week_plan'`).get()) {
      db.prepare(`DELETE FROM weekly_inspection_week_plan`).run();
    }
    if (db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name='weekly_inspection_entries'`).get()) {
      db.prepare(`DELETE FROM weekly_inspection_entries`).run();
    }
  }
  return { roster_cleared, slots_cleared };
}

function wiPdfSlotStatusTag(status) {
  const st = String(status || "pending").toLowerCase();
  if (st === "done") return { tag: "REL", color: "#15803d" };
  if (st === "skipped") return { tag: "SKIP", color: "#a16207" };
  return { tag: "PEN", color: "#475569" };
}

function wiPdfMonthTitle(ym) {
  const parts = String(ym || "").trim().split("-");
  const y = Number(parts[0]);
  const m = Number(parts[1]);
  if (!y || !m) return String(ym || "");
  return new Date(y, m - 1, 1).toLocaleDateString("en-US", { month: "long", year: "numeric" });
}

function wiPdfAddLandscapePage(doc) {
  doc.addPage({
    size: doc.page.size || "A4",
    layout: "landscape",
    margins: doc.page.margins,
  });
}

function wiEnsureBodySpace(doc, neededHeight, siteName) {
  if (doc.y + neededHeight > pdfBodyBottom(doc)) {
    wiPdfAddLandscapePage(doc);
    doc.y = pdfBodyTop(doc, { siteName });
  }
}

function drawWeeklyInspectionCalendarPdfGrid(doc, data, opts = {}) {
  const siteName = String(opts.siteName || "");
  const bodyTop = () => pdfBodyTop(doc, { siteName });
  const weeks = Array.isArray(data?.calendar_weeks) ? data.calendar_weeks : [];
  if (!weeks.length) {
    doc.font("Helvetica").fontSize(10).fillColor("#64748b").text("No calendar data for this month.");
    return;
  }
  const margin = doc.page.margins;
  const contentW = doc.page.width - margin.left - margin.right;
  const colW = contentW / 7;
  const lineH = 7;
  const dayNames = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
  const monthLabel = String(data?.month || "");

  const drawMonthCaption = (y) => {
    if (!monthLabel) return y;
    doc.font("Helvetica-Bold").fontSize(12).fillColor("#0f172a");
    doc.text(wiPdfMonthTitle(monthLabel), margin.left, y, { width: contentW, align: "center" });
    return y + 18;
  };

  const drawDayHeaders = (y) => {
    dayNames.forEach((label, i) => {
      doc.font("Helvetica-Bold").fontSize(8).fillColor("#64748b");
      doc.text(label, margin.left + i * colW + 3, y, { width: colW - 6, align: "center" });
    });
    return y + 14;
  };

  let y = drawMonthCaption(doc.y + 2);
  y = drawDayHeaders(y);

  for (const row of weeks) {
    let maxSlots = 0;
    for (let i = 0; i < 7; i += 1) {
      const cell = row?.[i] || {};
      if (cell?.in_month && cell?.date) {
        maxSlots = Math.max(maxSlots, Array.isArray(cell.slots) ? cell.slots.length : 0);
      }
    }
    const cellH = Math.max(38, 14 + maxSlots * lineH + 6);

    if (y + cellH > pdfBodyBottom(doc)) {
      wiPdfAddLandscapePage(doc);
      y = bodyTop() + 4;
      y = drawMonthCaption(y);
      y = drawDayHeaders(y);
    }

    for (let i = 0; i < 7; i += 1) {
      const cell = row?.[i] || {};
      const x = margin.left + i * colW;
      const inMonth = Boolean(cell?.in_month && cell?.date);
      if (!inMonth) {
        doc.save();
        doc.fillColor("#f8fafc").rect(x, y, colW - 2, cellH).fill();
        doc.strokeColor("#e2e8f0").lineWidth(0.75).rect(x, y, colW - 2, cellH).stroke();
        doc.restore();
        continue;
      }
      doc.save();
      doc.fillColor("#ffffff").rect(x, y, colW - 2, cellH).fill();
      doc.strokeColor("#cbd5e1").lineWidth(0.75).rect(x, y, colW - 2, cellH).stroke();
      doc.restore();
      doc.font("Helvetica-Bold").fontSize(9).fillColor("#0f172a");
      doc.text(String(cell.day || ""), x + 4, y + 4, { width: colW - 8 });
      const slots = Array.isArray(cell.slots) ? cell.slots : [];
      let lineY = y + 16;
      for (const slot of slots) {
        const meta = wiPdfSlotStatusTag(slot.status);
        doc.font("Helvetica").fontSize(6.5).fillColor(meta.color);
        doc.text(
          `${meta.tag} ${String(slot.asset_code || "-")} ${Number(slot.est_minutes || 30)}m`,
          x + 3,
          lineY,
          { width: colW - 8, lineBreak: false },
        );
        lineY += lineH;
      }
    }
    y += cellH + 3;
  }
  doc.y = Math.min(y + 8, pdfBodyBottom(doc) - 4);
}

function updateWeeklyInspectionSlotStatus({ slot_id, asset_id, planned_date, status, inspector_name }) {
  ensureWeeklyInspectionSchema();
  let row = null;
  if (slot_id) {
    row = db.prepare(`SELECT id, asset_id, planned_date, status FROM weekly_inspection_slots WHERE id = ?`).get(Number(slot_id));
  } else if (asset_id && isDate(String(planned_date || "").trim())) {
    row = db.prepare(`
      SELECT id, asset_id, planned_date, status
      FROM weekly_inspection_slots
      WHERE asset_id = ? AND planned_date = ?
    `).get(Number(asset_id), String(planned_date).trim());
  }
  if (!row) throw new Error("Inspection slot not found");
  const next = String(status || "pending").trim().toLowerCase();
  if (!["pending", "done", "skipped"].includes(next)) {
    throw new Error("status must be pending, done, or skipped");
  }
  const completed_at = next === "done" ? new Date().toISOString() : null;
  const released = next === "done" ? 1 : 0;
  db.prepare(`
    UPDATE weekly_inspection_slots
    SET status = ?, inspector_name = ?, completed_at = ?, released = ?, updated_at = datetime('now')
    WHERE id = ?
  `).run(next, inspector_name || null, completed_at, released, Number(row.id));
  return db.prepare(`
    SELECT
      id, planned_date, asset_id, est_minutes, status, inspector_name, notes, completed_at,
      COALESCE(released, 0) AS released
    FROM weekly_inspection_slots
    WHERE id = ?
  `).get(Number(row.id));
}

function getAssetCurrentHoursInfo(assetId) {
  const fromAssetHours = db.prepare(`
    SELECT total_hours
    FROM asset_hours
    WHERE asset_id = ?
  `).get(assetId);

  const assetHours = fromAssetHours?.total_hours == null ? null : Number(fromAssetHours.total_hours);

  // Prefer latest hourmeter closing reading from Daily Input when it exists.
  // This guards against old/test values in asset_hours (e.g. accidental CSV import).
  const latestMeter = db.prepare(`
    SELECT closing_hours AS latest_closing
    FROM daily_hours
    WHERE asset_id = ?
      AND closing_hours IS NOT NULL
      AND DATE(work_date) IS NOT NULL
    ORDER BY work_date DESC, id DESC
    LIMIT 1
  `).get(assetId);
  const latestClosing = latestMeter?.latest_closing == null ? null : Number(latestMeter.latest_closing);

  // If asset_hours is present but wildly out of range vs max closing, trust max closing.
  // Heuristic: >5000 hour difference is almost certainly wrong for a live hourmeter.
  if (assetHours != null && latestClosing != null) {
    if (Math.abs(assetHours - latestClosing) > 5000) {
      return { hours: latestClosing, source: "daily_closing" };
    }
    // otherwise take the higher of the two (prevents lagging asset_hours)
    if (latestClosing >= assetHours) return { hours: latestClosing, source: "daily_closing" };
    return { hours: assetHours, source: "asset_hours" };
  }

  if (latestClosing != null) return { hours: latestClosing, source: "daily_closing" };
  if (assetHours != null) return { hours: assetHours, source: "asset_hours" };

  const fromDailyHours = db.prepare(`
    SELECT COALESCE(SUM(hours_run), 0) AS total_hours
    FROM daily_hours
    WHERE asset_id = ?
      AND is_used = 1
      AND hours_run > 0
  `).get(assetId);

  return { hours: Number(fromDailyHours?.total_hours || 0), source: "daily_sum" };
}

function getAssetCurrentHours(assetId) {
  return Number(getAssetCurrentHoursInfo(assetId).hours || 0);
}

/** Meter / usage as of inspection date (daily rows with work_date <= as_of). Falls back to current fleet logic. */
function getAssetHoursInfoAsOf(assetId, asOfYmd) {
  const aid = Number(assetId || 0);
  if (!aid || !isDate(String(asOfYmd || "").trim())) return getAssetCurrentHoursInfo(aid);
  const asOf = String(asOfYmd).trim();

  const meter = db.prepare(`
    SELECT closing_hours AS latest_closing
    FROM daily_hours
    WHERE asset_id = ?
      AND closing_hours IS NOT NULL
      AND DATE(work_date) IS NOT NULL
      AND work_date <= ?
    ORDER BY work_date DESC, id DESC
    LIMIT 1
  `).get(aid, asOf);
  const closing = meter?.latest_closing == null ? null : Number(meter.latest_closing);
  if (closing != null && Number.isFinite(closing)) {
    return { hours: closing, source: "daily_closing" };
  }

  const fromDailyHours = db.prepare(`
    SELECT COALESCE(SUM(hours_run), 0) AS total_hours
    FROM daily_hours
    WHERE asset_id = ?
      AND is_used = 1
      AND hours_run > 0
      AND work_date <= ?
  `).get(aid, asOf);
  const th = Number(fromDailyHours?.total_hours || 0);
  if (th > 0) return { hours: th, source: "daily_sum" };

  return getAssetCurrentHoursInfo(aid);
}

function classifyDueStatus(remainingHours, nearDueHours = 50, assetCode = null, serviceInterval = 0) {
  if (assetCode) {
    return classifyServiceDueStatus(remainingHours, assetCode, serviceInterval, nearDueHours);
  }
  const remaining = Number(remainingHours || 0);
  const threshold = Math.max(1, Number(nearDueHours || 50));
  if (remaining <= 0) return "OVERDUE";
  if (remaining <= threshold) return "ALMOST DUE";
  return "OK";
}

/** SQL fragment: stock_movements row is an outbound issue (consumption). */
function sqlStockMovementOutbound(alias = "sm") {
  const s = alias;
  return `(LOWER(COALESCE(${s}.movement_type, '')) = 'out' OR COALESCE(${s}.quantity, 0) < 0)`;
}

/** SQL boolean expression (SQLite): parts row is an oil / lubricant store item, not a hard part. */
function sqlOilPartPredicate(alias = "p") {
  ensureStockCategorySchema(db);
  return oilCategorySql(alias);
}

/**
 * SQL expression for one stock_movements row's monetary cost (optionally joined to parts as `p`).
 * Aligns with asset maintenance history: line total_cost, sm.unit_cost×qty, parts.unit_cost×qty, then unit_cost_usd/cost_input×qty.
 */
/** SQL expression: sum of oil_log line costs with unit_cost fallback when stores issue omits cost. */
function sqlOilLogCostSumExpr(olAlias = "ol", unitCostFallbackSql = "?") {
  const ol = olAlias;
  return `COALESCE(SUM(COALESCE(${ol}.quantity, 0) * COALESCE(NULLIF(${ol}.unit_cost, 0), ${unitCostFallbackSql})), 0)`;
}

function sqlStockMovementDateExpr(hasColumnFn, smAlias = "sm") {
  const sm = smAlias;
  if (hasColumnFn("stock_movements", "created_at")) return `DATE(${sm}.created_at)`;
  if (hasColumnFn("stock_movements", "movement_date")) return `DATE(${sm}.movement_date)`;
  return `DATE(${sm}.created_at)`;
}

function readLubeCostPerQtyDefault(dbConn, hasTableFn) {
  let fallback = 4.0;
  if (hasTableFn("cost_settings")) {
    const row = dbConn.prepare(`
      SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1
    `).get();
    const n = Number(row?.value);
    if (Number.isFinite(n) && n > 0) fallback = n;
  }
  return fallback;
}

function sqlStockMovementLineCostExpr(smAlias = "sm", partsAlias = "p", joinPartsTable, flags) {
  const sm = smAlias;
  const p = partsAlias;
  if (flags.hasSmTotalCost) {
    return `ABS(COALESCE(${sm}.total_cost, 0))`;
  }
  if (flags.hasSmUnitCost) {
    return `ABS(COALESCE(${sm}.quantity, 0)) * COALESCE(${sm}.unit_cost, 0)`;
  }
  if (joinPartsTable && flags.hasPartsUnitCost) {
    return `ABS(COALESCE(${sm}.quantity, 0)) * COALESCE(${p}.unit_cost, 0)`;
  }
  if (flags.hasSmUnitCostUsd || flags.hasSmCostInput) {
    return `ABS(COALESCE(${sm}.quantity, 0)) * COALESCE(${sm}.unit_cost_usd, ${sm}.cost_input, 0)`;
  }
  return `0`;
}

function dbHasTable(dbConn, name) {
  return Boolean(
    dbConn.prepare(`
      SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ? LIMIT 1
    `).get(String(name || "")),
  );
}

function dbHasColumn(dbConn, table, col) {
  if (!dbHasTable(dbConn, table)) return false;
  return dbConn.prepare(`PRAGMA table_info(${table})`).all()
    .some((r) => String(r.name || "") === String(col || ""));
}

function buildStockCostSqlContext(dbConn) {
  const hasTable = (name) => dbHasTable(dbConn, name);
  const hasColumn = (table, col) => dbHasColumn(dbConn, table, col);
  const closedStatuses = "'closed','completed','approved'";
  const hasWOCompletedAt = hasColumn("work_orders", "completed_at");
  const woCloseExpr = hasWOCompletedAt ? "COALESCE(w.completed_at, w.closed_at)" : "w.closed_at";
  const smOutSql = sqlStockMovementOutbound("sm");
  const oilPartSql = sqlOilPartPredicate("p");
  const smLineFlags = {
    hasSmTotalCost: hasColumn("stock_movements", "total_cost"),
    hasSmUnitCost: hasColumn("stock_movements", "unit_cost"),
    hasPartsUnitCost: hasTable("parts") && hasColumn("parts", "unit_cost"),
    hasSmUnitCostUsd: hasColumn("stock_movements", "unit_cost_usd"),
    hasSmCostInput: hasColumn("stock_movements", "cost_input"),
  };
  return {
    hasTable,
    hasColumn,
    closedStatuses,
    woCloseExpr,
    smOutSql,
    oilPartSql,
    smCostWithParts: sqlStockMovementLineCostExpr("sm", "p", true, smLineFlags),
    smCostNoParts: sqlStockMovementLineCostExpr("sm", "p", false, smLineFlags),
  };
}

function getPartPricingForForecast(dbConn, partCodeIn, ctx) {
  const partCode = String(partCodeIn || "").trim();
  if (!partCode || !ctx.hasTable("parts") || !ctx.hasTable("stock_movements")) {
    return { unit_cost: 0, on_hand: 0, part_name: null };
  }
  const part = dbConn.prepare(`
    SELECT id, part_name, COALESCE(unit_cost, 0) AS part_list_unit_cost
    FROM parts
    WHERE UPPER(TRIM(part_code)) = UPPER(TRIM(?))
    LIMIT 1
  `).get(partCode);
  if (!part?.id) return { unit_cost: 0, on_hand: 0, part_name: null };
  const costRow = dbConn.prepare(`
    SELECT COALESCE(unit_cost_usd, cost_input, 0) AS unit_cost
    FROM stock_movements
    WHERE part_id = ?
      AND COALESCE(unit_cost_usd, cost_input, 0) > 0
    ORDER BY id DESC
    LIMIT 1
  `).get(part.id);
  const onHandRow = dbConn.prepare(`
    SELECT COALESCE(SUM(quantity), 0) AS on_hand
    FROM stock_movements
    WHERE part_id = ?
  `).get(part.id);
  const fromMove = Number(costRow?.unit_cost || 0);
  const fromList = Number(part?.part_list_unit_cost || 0);
  return {
    unit_cost: fromMove > 0 ? fromMove : fromList,
    on_hand: Number(onHandRow?.on_hand || 0),
    part_name: String(part.part_name || ""),
  };
}

/**
 * Cost history must survive a maintenance-plan replacement.  Work orders retain
 * the plan ID that existed when the service was closed, whereas the current
 * schedule may have a new ID for the same asset and service interval.  Prefer
 * the number named in the service label (old imports occasionally saved a
 * 1000-hour service with a 500-hour interval), then fall back to the interval.
 */
export function serviceCostHistoryKey(plan) {
  const serviceName = String(plan?.service_name || plan?.serviceName || "").trim();
  const namedInterval = serviceName.match(/(\d+(?:\.\d+)?)\s*(?:h(?:r|our)?s?|km)?/i);
  if (namedInterval) {
    const n = Number(namedInterval[1]);
    if (Number.isFinite(n) && n > 0) return `interval:${n}`;
  }
  const interval = Number(plan?.next_service_interval || plan?.interval_hours || plan?.intervalHours || 0);
  return Number.isFinite(interval) && interval > 0 ? `interval:${interval}` : "";
}

const MAINTENANCE_TEMPLATE_ROLE_PERMISSIONS = {
  admin: ["*"],
  workshop_admin: [
    "maintenance.templates.read",
    "maintenance.templates.manage",
    "maintenance.estimates.create",
    "maintenance.estimates.approve",
    "maintenance.estimates.convert",
    "maintenance.stock.reserve",
  ],
  plant_manager: [
    "maintenance.templates.read",
    "maintenance.estimates.create",
    "maintenance.estimates.approve",
  ],
  site_manager: [
    "maintenance.templates.read",
    "maintenance.estimates.create",
    "maintenance.estimates.approve",
  ],
  supervisor: [
    "maintenance.templates.read",
    "maintenance.estimates.create",
    "maintenance.estimates.approve",
    "maintenance.estimates.convert",
  ],
  plant_clerk: ["maintenance.templates.read", "maintenance.estimates.create"],
  artisan: ["maintenance.templates.read"],
};

function getMaintenancePermissions(req) {
  const fromHeader = String(req.headers["x-user-permissions"] || "")
    .split(",")
    .map((value) => String(value || "").trim())
    .filter(Boolean);
  if (fromHeader.length) return Array.from(new Set(fromHeader));
  return Array.from(new Set(getMaintenanceRoles(req)
    .flatMap((role) => MAINTENANCE_TEMPLATE_ROLE_PERMISSIONS[role] || [])));
}

function requireMaintenancePermission(req, reply, permission) {
  const permissions = getMaintenancePermissions(req);
  if (permissions.includes("*") || permissions.includes(permission)) return true;
  reply.code(403).send({ ok: false, error: `permission '${permission}' required` });
  return false;
}

function maintenanceActor(req) {
  return String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
}

/** Upcoming service kit + labor cost per maintenance plan (historical averages or saved manual inputs). */
export function buildUpcomingServiceCostForecasts(dbConn, plans, opts = {}) {
  const nearDueHours = Math.max(1, Number(opts.nearDueHours || 50));
  const maxRemainingHours = opts.maxRemainingHours != null
    ? Math.max(0, Number(opts.maxRemainingHours))
    : Math.max(nearDueHours, Number(opts.horizonHours || 100));
  const ctx = opts.ctx || buildStockCostSqlContext(dbConn);
  const { hasTable, hasColumn, closedStatuses, woCloseExpr, smOutSql, oilPartSql, smCostWithParts, smCostNoParts } = ctx;

  const forecastInputs = hasTable("weekly_forum_service_inputs")
    ? dbConn.prepare(`
        SELECT plan_id, oil_part_code, oil_qty, parts_part_code, parts_qty, items_json, notes,
          COALESCE(labor_total, 0) AS labor_total,
          COALESCE(all_in_total, 0) AS all_in_total
        FROM weekly_forum_service_inputs
      `).all()
    : [];
  const inputByPlan = new Map((forecastInputs || []).map((r) => [Number(r.plan_id || 0), r]));

  // Historical service work orders point at the plan that was active when the
  // work was completed. Build a same-asset / same-service index so a recreated
  // plan continues to use the actual cost history instead of appearing free.
  const historicalPlans = hasTable("maintenance_plans")
    ? dbConn.prepare(`SELECT id, asset_id, service_name, interval_hours FROM maintenance_plans`).all()
    : [];
  const historicalPlanIdsByService = new Map();
  for (const plan of historicalPlans) {
    const assetId = Number(plan.asset_id || 0);
    const key = serviceCostHistoryKey(plan);
    if (!assetId || !key) continue;
    const indexKey = `${assetId}:${key}`;
    if (!historicalPlanIdsByService.has(indexKey)) historicalPlanIdsByService.set(indexKey, []);
    historicalPlanIdsByService.get(indexKey).push(Number(plan.id || 0));
  }

  const getAssetHoursSafe = (assetId) => {
    try {
      if (typeof opts.getAssetHours === "function") {
        return Number(opts.getAssetHours(Number(assetId || 0)) || 0);
      }
      return Number(getAssetCurrentHoursInfo(Number(assetId || 0)).hours || 0);
    } catch {
      return 0;
    }
  };

  return (Array.isArray(plans) ? plans : [])
    .map((p) => {
      const planId = Number(p.plan_id || 0);
      const assetId = Number(p.asset_id || 0);
      const current = getAssetHoursSafe(assetId);
      const nextDue = Number(p.last_service_hours || 0) + Number(p.interval_hours || 0);
      const remaining = nextDue - current;
      const status = classifyDueStatus(remaining, nearDueHours);
      const historyKey = serviceCostHistoryKey(p);
      const historyPlanIds = Array.from(new Set([
        planId,
        ...(historicalPlanIdsByService.get(`${assetId}:${historyKey}`) || []),
      ].filter((id) => Number(id) > 0)));
      // A malformed legacy plan must return an unpriced row rather than break
      // the complete forecast query with an empty IN clause.
      const historyPlanBinds = historyPlanIds.length ? historyPlanIds : [-1];
      const historyPlanMarks = historyPlanBinds.map(() => "?").join(",");

      const hist =
        hasTable("work_orders") && hasTable("stock_movements")
          ? hasTable("parts")
            ? dbConn.prepare(`
                SELECT
                  COUNT(DISTINCT w.id) AS service_events,
                  COALESCE(SUM(CASE WHEN sm.id IS NOT NULL AND (${smOutSql}) AND NOT (${oilPartSql})
                    THEN ABS(COALESCE(sm.quantity, 0)) ELSE 0 END), 0) AS parts_qty_total,
                  COALESCE(SUM(CASE WHEN sm.id IS NOT NULL AND (${smOutSql}) AND NOT (${oilPartSql})
                    THEN (${smCostWithParts}) ELSE 0 END), 0) AS parts_cost_total,
                  COALESCE(SUM(CASE WHEN sm.id IS NOT NULL AND (${smOutSql}) AND (${oilPartSql})
                    THEN ABS(COALESCE(sm.quantity, 0)) ELSE 0 END), 0) AS oil_qty_sm_total,
                  COALESCE(SUM(CASE WHEN sm.id IS NOT NULL AND (${smOutSql}) AND (${oilPartSql})
                    THEN (${smCostWithParts}) ELSE 0 END), 0) AS oil_cost_sm_total
                FROM work_orders w
                LEFT JOIN stock_movements sm ON sm.reference = ('work_order:' || w.id)
                LEFT JOIN parts p ON p.id = sm.part_id
                WHERE LOWER(COALESCE(w.source, '')) = 'service'
                  AND COALESCE(w.reference_id, 0) IN (${historyPlanMarks})
                  AND LOWER(COALESCE(w.status, '')) IN (${closedStatuses})
              `).get(...historyPlanBinds)
            : dbConn.prepare(`
                SELECT
                  COUNT(DISTINCT w.id) AS service_events,
                  COALESCE(SUM(CASE WHEN sm.id IS NOT NULL AND (${smOutSql})
                    THEN ABS(COALESCE(sm.quantity, 0)) ELSE 0 END), 0) AS parts_qty_total,
                  COALESCE(SUM(CASE WHEN sm.id IS NOT NULL AND (${smOutSql})
                    THEN (${smCostNoParts}) ELSE 0 END), 0) AS parts_cost_total,
                  0 AS oil_qty_sm_total,
                  0 AS oil_cost_sm_total
                FROM work_orders w
                LEFT JOIN stock_movements sm ON sm.reference = ('work_order:' || w.id)
                WHERE LOWER(COALESCE(w.source, '')) = 'service'
                  AND COALESCE(w.reference_id, 0) IN (${historyPlanMarks})
                  AND LOWER(COALESCE(w.status, '')) IN (${closedStatuses})
              `).get(...historyPlanBinds)
          : null;

      const serviceEvents = Number(hist?.service_events || 0);
      const avgPartsQty = serviceEvents > 0 ? Number(hist.parts_qty_total || 0) / serviceEvents : 0;
      const avgPartsCost = serviceEvents > 0 ? Number(hist.parts_cost_total || 0) / serviceEvents : 0;

      const lubeCostDefault = Number(ctx.lubeCostDefault || 4);
      const oilAvg = hasTable("oil_logs") && hasTable("work_orders")
        ? dbConn.prepare(`
            SELECT
              COALESCE(SUM(ol.quantity), 0) AS oil_qty_total,
              ${sqlOilLogCostSumExpr("ol", "?")} AS oil_cost_total
            FROM oil_logs ol
            WHERE ol.asset_id = ?
              AND ol.log_date IN (
                SELECT DATE(${woCloseExpr})
                FROM work_orders w
                WHERE LOWER(COALESCE(w.source, '')) = 'service'
                  AND COALESCE(w.reference_id, 0) IN (${historyPlanMarks})
                  AND ${woCloseExpr} IS NOT NULL
                  AND LOWER(COALESCE(w.status, '')) IN (${closedStatuses})
              )
          `).get(lubeCostDefault, assetId, ...historyPlanBinds)
        : null;
      const oilQtyLogs = Number(oilAvg?.oil_qty_total || 0);
      const oilCostLogsPlan = Number(oilAvg?.oil_cost_total || 0);
      const oilQtySm = Number(hist?.oil_qty_sm_total || 0);
      const oilCostSm = Number(hist?.oil_cost_sm_total || 0);
      const avgOilQty = serviceEvents > 0 ? (oilQtyLogs + oilQtySm) / serviceEvents : 0;
      const avgOilCost = serviceEvents > 0 ? (oilCostLogsPlan + oilCostSm) / serviceEvents : 0;

      const laborHist = hasTable("work_orders") && hasColumn("work_orders", "labor_hours") && hasColumn("work_orders", "labor_rate_per_hour")
        ? dbConn.prepare(`
            SELECT
              COUNT(*) AS service_events,
              COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, 0)), 0) AS labor_cost_total
            FROM work_orders w
            WHERE LOWER(COALESCE(w.source, '')) = 'service'
              AND COALESCE(w.reference_id, 0) IN (${historyPlanMarks})
              AND LOWER(COALESCE(w.status, '')) IN (${closedStatuses})
          `).get(...historyPlanBinds)
        : null;
      const laborEvents = Number(laborHist?.service_events || 0);
      const avgLaborCost = laborEvents > 0 ? Number(laborHist.labor_cost_total || 0) / laborEvents : 0;

      const serviceKitCost = avgPartsCost + avgOilCost;
      const manual = inputByPlan.get(planId) || null;
      let manualItems = [];
      try {
        const parsed = JSON.parse(String(manual?.items_json || "[]"));
        if (Array.isArray(parsed)) manualItems = parsed;
      } catch {}
      if (!manualItems.length) {
        manualItems = [
          { type: "oil", part_code: String(manual?.oil_part_code || "").trim(), qty: Number(manual?.oil_qty || 0) },
          { type: "part", part_code: String(manual?.parts_part_code || "").trim(), qty: Number(manual?.parts_qty || 0) },
        ].filter((x) => x.part_code && Number(x.qty || 0) > 0);
      }
      const pricedItems = manualItems.map((it) => {
        const type = String(it?.type || "part").toLowerCase() === "oil" ? "oil" : "part";
        const part_code = String(it?.part_code || "").trim();
        const qty = Math.max(0, Number(it?.qty || 0));
        const pricing = part_code ? getPartPricingForForecast(dbConn, part_code, ctx) : { unit_cost: 0, on_hand: 0, part_name: null };
        return {
          type,
          part_code,
          part_name: pricing.part_name || null,
          qty: Number(qty.toFixed(2)),
          unit_cost: Number(Number(pricing.unit_cost || 0).toFixed(4)),
          on_hand: Number(Number(pricing.on_hand || 0).toFixed(2)),
          line_cost: Number((qty * Number(pricing.unit_cost || 0)).toFixed(2)),
        };
      }).filter((x) => x.part_code && x.qty > 0);
      const manualOilCost = pricedItems.filter((x) => x.type === "oil").reduce((s, x) => s + Number(x.line_cost || 0), 0);
      const manualPartsCost = pricedItems.filter((x) => x.type !== "oil").reduce((s, x) => s + Number(x.line_cost || 0), 0);
      const manualLaborTotal = Math.max(0, Number(manual?.labor_total || 0));
      const manualAllInTotal = Math.max(0, Number(manual?.all_in_total || 0));
      const hasManualParts = pricedItems.length > 0;
      const hasManualLabor = manualLaborTotal > 0;
      const hasManualAllIn = manualAllInTotal > 0;
      const hasManualOverride = hasManualParts || hasManualLabor;
      // A machine's service template prices the kit, oils and standard labour
      // from current store costs. It fills whatever the manual inputs leave out.
      const template = hasManualAllIn ? null : templateCostForPlan(dbConn, {
        assetId, planId, assetCode: p.asset_code, intervalHours: p.interval_hours, meterReading: nextDue,
      });
      const useTemplateParts = !hasManualParts && template && template.materials > 0;
      const useTemplateLabor = !hasManualLabor && template && template.labour > 0;
      // A quoted all-in estimate deliberately replaces the detailed split. It avoids
      // presenting an invented kit/labor breakdown when only a supplier budget is known.
      const estKitCost = Number((hasManualAllIn ? 0 : (hasManualParts ? (manualOilCost + manualPartsCost) : useTemplateParts ? template.materials : serviceKitCost)).toFixed(2));
      const estLaborCost = Number((hasManualAllIn ? 0 : (hasManualLabor ? manualLaborTotal : useTemplateLabor ? template.labour : avgLaborCost)).toFixed(2));
      const estTotalCost = Number((hasManualAllIn ? manualAllInTotal : (estKitCost + estLaborCost)).toFixed(2));
      const costSource = hasManualAllIn
        ? "manual_all_in_estimate"
        : hasManualOverride
        ? (hasManualParts && hasManualLabor ? "manual_parts_and_labor" : hasManualParts ? "manual_store_pricing" : "manual_labor")
        : (useTemplateParts || useTemplateLabor)
        ? "service_template"
        : (estTotalCost > 0 && (serviceEvents > 0 || laborEvents > 0)
          ? (historyPlanIds.length > 1 ? "historical_asset_service_average" : "historical_average")
          : "none");

      return {
        plan_id: planId,
        asset_id: assetId,
        asset_code: p.asset_code,
        asset_name: p.asset_name,
        service_name: p.service_name,
        current_hours: Number(current.toFixed(2)),
        next_due_hours: Number(nextDue.toFixed(2)),
        remaining_hours: Number(remaining.toFixed(2)),
        status,
        // A service without a defensible amount is not a zero-cost service.
        needs_manual_input: estTotalCost <= 0,
        forecast: {
          service_events: serviceEvents,
          avg_oil_qty: Number(avgOilQty.toFixed(2)),
          avg_oil_cost: Number(avgOilCost.toFixed(2)),
          avg_parts_qty: Number(avgPartsQty.toFixed(2)),
          avg_parts_cost: Number(avgPartsCost.toFixed(2)),
          avg_labor_cost: Number(avgLaborCost.toFixed(2)),
          est_service_kit_cost: estKitCost,
          est_labor_cost: estLaborCost,
          est_total_cost: estTotalCost,
          cost_source: costSource,
          template: template ? { id: template.id, name: template.name, pricing_complete: template.pricing_complete, labour_hours: template.labour_hours } : null,
          manual: {
            oil_cost_total: Number(manualOilCost.toFixed(2)),
            parts_cost_total: Number(manualPartsCost.toFixed(2)),
            labor_total: Number(manualLaborTotal.toFixed(2)),
            all_in_total: Number(manualAllInTotal.toFixed(2)),
            items: pricedItems,
            notes: String(manual?.notes || ""),
          },
        },
      };
    })
    .filter((r) => Number(r.remaining_hours || 0) <= maxRemainingHours)
    .sort((a, b) => Number(a.remaining_hours || 0) - Number(b.remaining_hours || 0));
}

/** Template-based cost for a plan: materials at store cost plus standard labour. Null when no template applies. */
function templateCostForPlan(dbConn, { assetId, planId, assetCode, intervalHours, meterReading }) {
  if (!dbHasTable(dbConn, "service_templates")) return null;
  try {
    const preview = buildServiceEstimatePreview(dbConn, {
      assetId, planId, meterReading, intervalHours, meterUnit: meterUnitForAsset(assetCode),
    });
    const e = preview?.estimate;
    if (!e) return null;
    return {
      id: e.service_template_id,
      name: e.template_name,
      materials: Number(e.estimated_parts_cost || 0) + Number(e.estimated_oil_cost || 0) + Number(e.estimated_consumables_cost || 0),
      labour: Number(e.estimated_labour_cost || 0),
      labour_hours: Number(e.labour_hours || 0),
      pricing_complete: Boolean(e.pricing_complete),
    };
  } catch {
    return null;
  }
}

function ensureBreakdownRepairLaborSchema(dbConn) {
  dbConn.prepare(`
    CREATE TABLE IF NOT EXISTS breakdown_repair_labor (
      breakdown_id INTEGER PRIMARY KEY,
      labor_hours REAL NOT NULL DEFAULT 0,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (breakdown_id) REFERENCES breakdowns(id) ON DELETE CASCADE
    )
  `).run();
}

/**
 * Breakdown incidents in [start,end]: machine downtime vs actual repair labor (not downtime × labor rate).
 */
function buildInsightsBreakdownLaborIncidents(dbConn, startDate, endDate, opts = {}) {
  ensureBreakdownRepairLaborSchema(dbConn);
  const scheduledFallback = Math.max(1, Number(opts.scheduledFallback || 10));
  const laborRate = Math.max(0, Number(opts.laborRate || 35));
  const empty = {
    incidents: [],
    by_asset: new Map(),
    totals: { downtime_hours: 0, repair_labor_hours: 0, repair_labor_cost: 0, needs_input_count: 0 },
  };
  if (!dbHasTable(dbConn, "breakdowns")) return empty;

  const incidentMap = new Map();
  const upsertIncident = (base, downtimeDelta) => {
    const bid = Number(base.breakdown_id || 0);
    if (!bid || !Number.isFinite(downtimeDelta) || downtimeDelta <= 0) return;
    const cur = incidentMap.get(bid) || {
      breakdown_id: bid,
      asset_id: Number(base.asset_id || 0),
      asset_code: String(base.asset_code || ""),
      asset_name: String(base.asset_name || ""),
      description: String(base.description || ""),
      status: String(base.status || ""),
      primary_work_order_id: Number(base.primary_work_order_id || 0) || null,
      downtime_hours: 0,
    };
    cur.downtime_hours = Number(cur.downtime_hours || 0) + downtimeDelta;
    incidentMap.set(bid, cur);
  };

  if (dbHasTable(dbConn, "breakdown_downtime_logs")) {
    for (const r of dbConn.prepare(`
      SELECT
        b.id AS breakdown_id,
        b.asset_id,
        a.asset_code,
        a.asset_name,
        b.description,
        b.status,
        b.primary_work_order_id,
        COALESCE(SUM(l.hours_down), 0) AS downtime_hours
      FROM breakdown_downtime_logs l
      JOIN breakdowns b ON b.id = l.breakdown_id
      JOIN assets a ON a.id = b.asset_id
      WHERE DATE(l.log_date) BETWEEN DATE(?) AND DATE(?)
      GROUP BY b.id
      HAVING downtime_hours > 0
    `).all(startDate, endDate)) {
      upsertIncident(r, Number(r.downtime_hours || 0));
    }
  }

  const todayYmd = new Date().toISOString().slice(0, 10);
  const imputeEnd = endDate < todayYmd ? endDate : todayYmd;
  const days = listDaysInclusiveYmd(startDate, imputeEnd);
  const getLogForBreakdownDay = dbHasTable(dbConn, "breakdown_downtime_logs")
    ? dbConn.prepare(`
        SELECT 1 FROM breakdown_downtime_logs
        WHERE breakdown_id = ? AND log_date = ? AND COALESCE(hours_down, 0) > 0
        LIMIT 1
      `)
    : null;
  const getScheduledForAssetDay = dbHasTable(dbConn, "daily_hours")
    ? dbConn.prepare(`
        SELECT scheduled_hours FROM daily_hours
        WHERE asset_id = ? AND work_date = ?
        LIMIT 1
      `)
    : null;
  const breakdownDateExpr = dbHasColumn(dbConn, "breakdowns", "breakdown_date")
    ? "b.breakdown_date"
    : "DATE(COALESCE(b.created_at, b.updated_at))";

  for (const br of dbConn.prepare(`
    SELECT
      b.id AS breakdown_id,
      b.asset_id,
      a.asset_code,
      a.asset_name,
      b.description,
      b.status,
      b.primary_work_order_id,
      DATE(COALESCE(${breakdownDateExpr}, b.created_at)) AS breakdown_day
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    LEFT JOIN work_orders wo ON wo.id = b.primary_work_order_id
    WHERE b.status = 'OPEN'
      AND (wo.id IS NULL OR LOWER(TRIM(COALESCE(wo.status, ''))) NOT IN ('completed', 'approved', 'closed'))
      AND DATE(COALESCE(${breakdownDateExpr}, b.created_at)) <= DATE(?)
  `).all(imputeEnd)) {
    const breakdownDay = String(br.breakdown_day || startDate);
    let imputed = 0;
    for (const day of days) {
      if (day < breakdownDay || day < startDate) continue;
      if (getLogForBreakdownDay?.get(Number(br.breakdown_id || 0), day)) continue;
      const sched = Number(getScheduledForAssetDay?.get(Number(br.asset_id || 0), day)?.scheduled_hours || 0);
      imputed += sched > 0 ? sched : scheduledFallback;
    }
    if (imputed > 0) upsertIncident(br, imputed);
  }

  const manualByBreakdown = new Map(
    dbConn.prepare(`SELECT breakdown_id, labor_hours, notes FROM breakdown_repair_labor`).all()
      .map((r) => [Number(r.breakdown_id || 0), r]),
  );
  const getWo = dbHasTable(dbConn, "work_orders")
    ? dbConn.prepare(`
        SELECT id, labor_hours, labor_rate_per_hour
        FROM work_orders
        WHERE id = ?
        LIMIT 1
      `)
    : null;

  const incidents = [];
  const byAsset = new Map();
  for (const inc of incidentMap.values()) {
    const bid = Number(inc.breakdown_id || 0);
    const downtimeHours = Number(Number(inc.downtime_hours || 0).toFixed(2));
    if (downtimeHours <= 0) continue;

    const manual = manualByBreakdown.get(bid);
    const wo = inc.primary_work_order_id && getWo ? getWo.get(inc.primary_work_order_id) : null;
    const manualLabor = Math.max(0, Number(manual?.labor_hours || 0));
    const woLabor = Math.max(0, Number(wo?.labor_hours || 0));
    let actualLabor = 0;
    let laborSource = "none";
    if (manualLabor > 0) {
      actualLabor = manualLabor;
      laborSource = "manual";
    } else if (woLabor > 0) {
      actualLabor = woLabor;
      laborSource = "work_order";
    }
    const rate = Number(wo?.labor_rate_per_hour) > 0 ? Number(wo.labor_rate_per_hour) : laborRate;
    const repairLaborCost = Number((actualLabor * rate).toFixed(2));
    const needsLaborInput = downtimeHours > 0 && actualLabor <= 0;

    const row = {
      breakdown_id: bid,
      asset_id: Number(inc.asset_id || 0),
      asset_code: inc.asset_code,
      asset_name: inc.asset_name,
      description: inc.description,
      status: inc.status,
      downtime_hours: downtimeHours,
      actual_labor_hours: Number(actualLabor.toFixed(2)),
      labor_rate: Number(rate.toFixed(2)),
      repair_labor_cost: repairLaborCost,
      labor_source: laborSource,
      needs_labor_input: needsLaborInput,
      labor_notes: String(manual?.notes || ""),
    };
    incidents.push(row);

    const aid = Number(inc.asset_id || 0);
    if (aid > 0) {
      const ar = byAsset.get(aid) || {
        asset_id: aid,
        downtime_hours: 0,
        repair_labor_hours: 0,
        repair_labor_cost: 0,
      };
      ar.downtime_hours += downtimeHours;
      ar.repair_labor_hours += actualLabor;
      ar.repair_labor_cost += repairLaborCost;
      byAsset.set(aid, ar);
    }
  }

  incidents.sort((a, b) => Number(b.downtime_hours || 0) - Number(a.downtime_hours || 0));

  const totals = incidents.reduce(
    (t, r) => ({
      downtime_hours: t.downtime_hours + Number(r.downtime_hours || 0),
      repair_labor_hours: t.repair_labor_hours + Number(r.actual_labor_hours || 0),
      repair_labor_cost: t.repair_labor_cost + Number(r.repair_labor_cost || 0),
      needs_input_count: t.needs_input_count + (r.needs_labor_input ? 1 : 0),
    }),
    { downtime_hours: 0, repair_labor_hours: 0, repair_labor_cost: 0, needs_input_count: 0 },
  );
  totals.downtime_hours = Number(totals.downtime_hours.toFixed(2));
  totals.repair_labor_hours = Number(totals.repair_labor_hours.toFixed(2));
  totals.repair_labor_cost = Number(totals.repair_labor_cost.toFixed(2));

  return { incidents, by_asset: byAsset, totals };
}

export default async function maintenanceRoutes(app) {
  ensureAuditTable(db);
  // These tables store reusable service definitions and cost snapshots. They
  // deliberately do not create stock movements; issuing remains a workshop/stores action.
  ensureServiceTemplateSchema(db);
  // A standard plant service schedule is one automatic 500h/1000h rotation.
  // Backfill the companion plan so existing assets no longer require two
  // schedules to be configured manually and generated work orders retain the
  // correct service type through their maintenance-plan reference.
  const standardPlanRows = db.prepare(`
    SELECT mp.*, a.asset_code
    FROM maintenance_plans mp
    JOIN assets a ON a.id = mp.asset_id
    WHERE mp.active = 1
      AND mp.interval_hours IN (500, 1000)
      AND a.asset_code NOT LIKE 'V__AM'
  `).all();
  const standardPlansByAsset = groupActivePlansByAsset(standardPlanRows);
  const insertStandardPlan = db.prepare(`
    INSERT INTO maintenance_plans (asset_id, service_name, interval_hours, last_service_hours, active)
    VALUES (?, ?, ?, ?, 1)
  `);
  const ensureStandardPlans = db.transaction(() => {
    for (const [assetId, plans] of standardPlansByAsset) {
      const intervals = new Set(plans.map(planIntervalHours));
      const seed = plans[0];
      for (const interval of [500, 1000]) {
        if (intervals.has(interval)) continue;
        insertStandardPlan.run(
          assetId,
          `${interval} hour service`,
          interval,
          snapLastServiceHours(Number(seed.last_service_hours || 0), interval, seed.asset_code),
        );
      }
    }
  });
  ensureStandardPlans();
  const dataRoot = getDataRoot();
  function resolveStorageAbs(relPath) {
    return resolveStorageAbsPath(relPath, dataRoot);
  }
  await app.register(multipart, {
    limits: { fileSize: 20 * 1024 * 1024 },
  });

  db.prepare(`
    CREATE TABLE IF NOT EXISTS manager_inspections (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      inspection_date TEXT NOT NULL,
      inspector_name TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS manager_inspection_photos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      inspection_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      file_path TEXT NOT NULL,
      caption TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (inspection_id) REFERENCES manager_inspections(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS artisan_inspections (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      inspection_date TEXT NOT NULL,
      inspector_name TEXT,
      shift TEXT,
      notes TEXT,
      machine_hours REAL,
      live_hours_snapshot REAL,
      live_hours_source TEXT,
      checklist_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS manager_damage_reports (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      report_date TEXT NOT NULL,
      inspector_name TEXT,
      hour_meter REAL,
      damage_location TEXT,
      severity TEXT,
      damage_description TEXT,
      immediate_action TEXT,
      out_of_service INTEGER NOT NULL DEFAULT 0,
      damage_time TEXT,
      responsible_person TEXT,
      pending_investigation INTEGER NOT NULL DEFAULT 0,
      hse_report_available INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS maintenance_service_history (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      plan_id INTEGER,
      service_name TEXT NOT NULL,
      service_date TEXT NOT NULL,
      service_hours REAL,
      notes TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT,
      FOREIGN KEY (plan_id) REFERENCES maintenance_plans(id) ON DELETE SET NULL
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS maintenance_histogram_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT DEFAULT 'main',
      event_date TEXT NOT NULL,
      asset_number TEXT,
      location TEXT,
      part_code TEXT,
      part_name TEXT,
      approval_status TEXT,
      approved_by TEXT,
      notes TEXT,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS maintenance_parts_requests (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT DEFAULT 'main',
      asset_id INTEGER,
      asset_code TEXT,
      part_code TEXT,
      part_name TEXT NOT NULL,
      qty REAL NOT NULL DEFAULT 1,
      urgency TEXT NOT NULL DEFAULT 'normal',
      notes TEXT,
      work_order_id INTEGER,
      status TEXT NOT NULL DEFAULT 'requested',
      requested_by TEXT NOT NULL,
      ordered_by TEXT,
      status_notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE SET NULL
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS mechanic_labor_entries (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      work_date TEXT NOT NULL,
      technician_name TEXT NOT NULL,
      hours REAL NOT NULL DEFAULT 0,
      asset_code TEXT NOT NULL,
      reason TEXT,
      labor_rate_per_hour REAL,
      site_code TEXT NOT NULL DEFAULT 'main',
      created_by TEXT,
      updated_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`
    CREATE INDEX IF NOT EXISTS idx_mechanic_labor_date
    ON mechanic_labor_entries(work_date, site_code)
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS manager_damage_report_photos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      damage_report_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      file_path TEXT,
      image_data TEXT,
      caption TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (damage_report_id) REFERENCES manager_damage_reports(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS weekly_inspection_assets (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL UNIQUE,
      active INTEGER NOT NULL DEFAULT 1,
      notes TEXT,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS weekly_inspection_entries (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      week_start TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT 'pending',
      inspector_name TEXT,
      notes TEXT,
      completed_at TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(asset_id, week_start),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS tyre_inspections (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      inspection_date TEXT NOT NULL,
      inspector_name TEXT,
      running_hours REAL,
      total_tyre_cost REAL NOT NULL DEFAULT 0,
      cost_per_running_hour REAL NOT NULL DEFAULT 0,
      tyres_json TEXT NOT NULL DEFAULT '[]',
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS undercarriage_inspections (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      inspection_date TEXT NOT NULL,
      inspector_name TEXT,
      smu REAL,
      job_no TEXT,
      site_name TEXT,
      planner TEXT,
      serial_no TEXT,
      unit_assembly TEXT,
      model TEXT,
      yard_no TEXT,
      work_order_no TEXT,
      component_group TEXT,
      group_id TEXT,
      component_serial_no TEXT,
      part_no TEXT,
      cost_center TEXT,
      measurements_json TEXT NOT NULL DEFAULT '[]',
      track_sag_json TEXT NOT NULL DEFAULT '{}',
      checklist_json TEXT NOT NULL DEFAULT '{}',
      summary_json TEXT NOT NULL DEFAULT '{}',
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS undercarriage_wear_profiles (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      site_code TEXT NOT NULL DEFAULT 'main',
      limits_json TEXT NOT NULL DEFAULT '[]',
      source TEXT,
      notes TEXT,
      updated_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(asset_id, site_code),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();

  function hasColumn(table, col) {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name || "") === String(col));
  }
  function hasTable(table) {
    const r = db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name = ? LIMIT 1`).get(String(table || ""));
    return Boolean(r?.name);
  }
  function ensureColumn(table, colDef, colName) {
    if (!hasColumn(table, colName)) db.prepare(`ALTER TABLE ${table} ADD COLUMN ${colDef}`).run();
  }
  function pickExistingColumn(table, candidates, fallback) {
    for (const c of candidates) {
      if (hasColumn(table, c)) return c;
    }
    return fallback;
  }
  ensureColumn("manager_inspections", "uuid TEXT", "uuid");
  ensureColumn("manager_inspections", "site_code TEXT DEFAULT 'main'", "site_code");
  ensureColumn("manager_inspections", "updated_at TEXT", "updated_at");
  ensureColumn("manager_inspections", "machine_hours REAL", "machine_hours");
  ensureColumn("manager_inspections", "live_hours_snapshot REAL", "live_hours_snapshot");
  ensureColumn("manager_inspections", "live_hours_source TEXT", "live_hours_source");
  ensureColumn("manager_inspections", "checklist_json TEXT", "checklist_json");
  ensureColumn("manager_inspections", "required_parts_json TEXT", "required_parts_json");
  ensureColumn("manager_inspections", "work_order_id INTEGER", "work_order_id");
  ensureColumn("manager_inspections", "defect_severity TEXT", "defect_severity");
  ensureColumn("manager_inspections", "defect_component TEXT", "defect_component");
  ensureColumn("manager_inspections", "defect_risk TEXT", "defect_risk");
  ensureColumn("manager_inspections", "recommended_action TEXT", "recommended_action");
  ensureColumn("manager_inspections", "inspection_type TEXT DEFAULT 'machine_general'", "inspection_type");
  ensureColumn("manager_inspections", "evidence_required INTEGER DEFAULT 1", "evidence_required");
  ensureColumn("manager_inspections", "evidence_photo_count INTEGER DEFAULT 0", "evidence_photo_count");
  ensureColumn("artisan_inspections", "uuid TEXT", "uuid");
  ensureColumn("artisan_inspections", "site_code TEXT DEFAULT 'main'", "site_code");
  ensureColumn("artisan_inspections", "updated_at TEXT", "updated_at");
  ensureColumn("artisan_inspections", "machine_hours REAL", "machine_hours");
  ensureColumn("artisan_inspections", "live_hours_snapshot REAL", "live_hours_snapshot");
  ensureColumn("artisan_inspections", "live_hours_source TEXT", "live_hours_source");
  ensureColumn("artisan_inspections", "checklist_json TEXT", "checklist_json");
  ensureColumn("artisan_inspections", "shift TEXT", "shift");
  ensureColumn("artisan_inspections", "form_number TEXT", "form_number");
  ensureColumn("manager_inspection_photos", "uuid TEXT", "uuid");
  ensureColumn("manager_inspection_photos", "site_code TEXT DEFAULT 'main'", "site_code");
  ensureColumn("manager_inspection_photos", "updated_at TEXT", "updated_at");
  ensureColumn("manager_inspection_photos", "file_path TEXT", "file_path");
  ensureColumn("manager_inspection_photos", "caption TEXT", "caption");
  ensureColumn("manager_inspection_photos", "created_at TEXT", "created_at");
  ensureColumn("manager_damage_reports", "uuid TEXT", "uuid");
  ensureColumn("manager_damage_reports", "site_code TEXT DEFAULT 'main'", "site_code");
  ensureColumn("manager_damage_reports", "updated_at TEXT", "updated_at");
  ensureColumn("manager_damage_reports", "inspector_name TEXT", "inspector_name");
  ensureColumn("manager_damage_reports", "hour_meter REAL", "hour_meter");
  ensureColumn("manager_damage_reports", "damage_location TEXT", "damage_location");
  ensureColumn("manager_damage_reports", "severity TEXT", "severity");
  ensureColumn("manager_damage_reports", "damage_description TEXT", "damage_description");
  ensureColumn("manager_damage_reports", "immediate_action TEXT", "immediate_action");
  ensureColumn("manager_damage_reports", "out_of_service INTEGER NOT NULL DEFAULT 0", "out_of_service");
  ensureColumn("manager_damage_reports", "damage_time TEXT", "damage_time");
  ensureColumn("manager_damage_reports", "responsible_person TEXT", "responsible_person");
  ensureColumn("manager_damage_reports", "pending_investigation INTEGER NOT NULL DEFAULT 0", "pending_investigation");
  ensureColumn("manager_damage_reports", "hse_report_available INTEGER NOT NULL DEFAULT 0", "hse_report_available");
  ensureColumn("manager_damage_report_photos", "uuid TEXT", "uuid");
  ensureColumn("manager_damage_report_photos", "site_code TEXT DEFAULT 'main'", "site_code");
  ensureColumn("manager_damage_report_photos", "updated_at TEXT", "updated_at");
  ensureColumn("manager_damage_report_photos", "file_path TEXT", "file_path");
  ensureColumn("manager_damage_report_photos", "image_data TEXT", "image_data");
  ensureColumn("manager_damage_report_photos", "caption TEXT", "caption");
  ensureColumn("manager_damage_report_photos", "created_at TEXT", "created_at");
  ensureColumn("tyre_inspections", "uuid TEXT", "uuid");
  ensureColumn("tyre_inspections", "site_code TEXT DEFAULT 'main'", "site_code");
  ensureColumn("tyre_inspections", "inspector_name TEXT", "inspector_name");
  ensureColumn("tyre_inspections", "running_hours REAL", "running_hours");
  ensureColumn("tyre_inspections", "total_tyre_cost REAL NOT NULL DEFAULT 0", "total_tyre_cost");
  ensureColumn("tyre_inspections", "cost_per_running_hour REAL NOT NULL DEFAULT 0", "cost_per_running_hour");
  ensureColumn("tyre_inspections", "tyres_json TEXT NOT NULL DEFAULT '[]'", "tyres_json");
  ensureColumn("tyre_inspections", "notes TEXT", "notes");
  ensureColumn("tyre_inspections", "updated_at TEXT", "updated_at");
  ensureColumn("maintenance_histogram_events", "site_code TEXT DEFAULT 'main'", "site_code");
  ensureColumn("maintenance_histogram_events", "event_date TEXT", "event_date");
  ensureColumn("maintenance_histogram_events", "asset_number TEXT", "asset_number");
  ensureColumn("maintenance_histogram_events", "location TEXT", "location");
  ensureColumn("maintenance_histogram_events", "part_code TEXT", "part_code");
  ensureColumn("maintenance_histogram_events", "part_name TEXT", "part_name");
  ensureColumn("maintenance_histogram_events", "approval_status TEXT", "approval_status");
  ensureColumn("maintenance_histogram_events", "approved_by TEXT", "approved_by");
  ensureColumn("maintenance_histogram_events", "notes TEXT", "notes");
  ensureColumn("maintenance_histogram_events", "created_by TEXT", "created_by");
  ensureColumn("maintenance_histogram_events", "created_at TEXT", "created_at");
  ensureColumn("maintenance_histogram_events", "updated_at TEXT", "updated_at");
  try {
    const pt = db.prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name='parts' LIMIT 1`).get();
    if (pt) ensureColumn("parts", "consumable_kind TEXT", "consumable_kind");
  } catch {}

  // Backward compatibility for legacy schema where link column was manager_inspection_id.
  // Keep both readable by normalizing to inspection_id for all new queries/inserts.
  if (!hasColumn("manager_inspection_photos", "inspection_id")) {
    ensureColumn("manager_inspection_photos", "inspection_id INTEGER", "inspection_id");
    if (hasColumn("manager_inspection_photos", "manager_inspection_id")) {
      db.prepare(`
        UPDATE manager_inspection_photos
        SET inspection_id = manager_inspection_id
        WHERE inspection_id IS NULL
          AND manager_inspection_id IS NOT NULL
      `).run();
    }
  }
  const photoInspectionCol = hasColumn("manager_inspection_photos", "inspection_id")
    ? "inspection_id"
    : hasColumn("manager_inspection_photos", "manager_inspection_id")
      ? "manager_inspection_id"
      : "inspection_id";
  const photoPathCol = pickExistingColumn(
    "manager_inspection_photos",
    ["file_path", "photo_path", "path", "image_path", "url"],
    "file_path"
  );
  const photoCaptionCol = pickExistingColumn(
    "manager_inspection_photos",
    ["caption", "note", "notes", "description"],
    "caption"
  );
  const photoCreatedCol = pickExistingColumn(
    "manager_inspection_photos",
    ["created_at", "uploaded_at", "created_on"],
    "created_at"
  );
  const dmgPhotoReportCol = pickExistingColumn(
    "manager_damage_report_photos",
    ["damage_report_id", "manager_damage_report_id", "report_id"],
    "damage_report_id"
  );
  const dmgPhotoPathCol = pickExistingColumn(
    "manager_damage_report_photos",
    ["file_path", "photo_path", "path", "image_path", "url", "image_data"],
    "file_path"
  );
  const dmgPhotoCaptionCol = pickExistingColumn(
    "manager_damage_report_photos",
    ["caption", "note", "notes", "description"],
    "caption"
  );
  const dmgPhotoCreatedCol = pickExistingColumn(
    "manager_damage_report_photos",
    ["created_at", "uploaded_at", "created_on"],
    "created_at"
  );

  // Backfill normalized file_path from common legacy column names.
  if (photoPathCol !== "file_path" && hasColumn("manager_inspection_photos", "file_path")) {
    db.prepare(`
      UPDATE manager_inspection_photos
      SET file_path = ${photoPathCol}
      WHERE (file_path IS NULL OR TRIM(file_path) = '')
        AND ${photoPathCol} IS NOT NULL
    `).run();
  }

  const inspectionsDir = path.join(dataRoot, "uploads", "manager-inspections");
  fs.mkdirSync(inspectionsDir, { recursive: true });
  const damageReportsDir = path.join(dataRoot, "uploads", "manager-damage-reports");
  fs.mkdirSync(damageReportsDir, { recursive: true });

  // =====================================================
  // MAINTENANCE PLANS - LIST
  // GET /api/maintenance/plans
  // =====================================================
  function displayPlanServiceName(plan) {
    const assetCode = String(plan?.asset_code || "").trim();
    const meterUnit = meterUnitForAsset(assetCode);
    const interval = planIntervalHours(plan);
    const supplied = String(plan?.service_name || "").trim();
    if (meterUnit === "km") return "10000 km service";
    if (!supplied || /^\d+(?:\.0+)?$/.test(supplied)) {
      return `${Number(interval || 0).toFixed(0)} hour service`;
    }
    return supplied;
  }

  function listMaintenancePlans(nearDueHours = 50) {
    const rows = db.prepare(`
      SELECT
        mp.id,
        mp.asset_id,
        mp.service_name,
        mp.interval_hours,
        mp.last_service_hours,
        mp.active,
        a.asset_code,
        a.asset_name,
        a.category
      FROM maintenance_plans mp
      JOIN assets a ON a.id = mp.asset_id
      WHERE a.archived = 0
      ORDER BY a.asset_code ASC, mp.service_name ASC
    `).all();

    return enrichPlansWithNextService(rows, getAssetCurrentHours, nearDueHours);
  }

  function cleanTemplateItemInput(raw) {
    const item = raw && typeof raw === "object" ? raw : {};
    const itemType = String(item.item_type || "part").trim().toLowerCase();
    const description = String(item.description || "").trim();
    const quantity = Number(item.quantity_required ?? item.quantity ?? 0);
    let stockItemId = Number(item.stock_item_id ?? item.part_id ?? 0) || null;
    const stockPartCode = String(item.stock_part_code ?? item.part_code ?? "").trim();
    if (!SERVICE_TEMPLATE_ITEM_TYPES.has(itemType)) {
      throw new Error(`Unsupported material type: ${itemType || "blank"}`);
    }
    if (!description) throw new Error("Each template material needs a description");
    if (!Number.isFinite(quantity) || quantity < 0) {
      throw new Error(`Invalid quantity for ${description}`);
    }
    if (!stockItemId && stockPartCode) {
      stockItemId = Number(db.prepare(`SELECT id FROM parts WHERE part_code = ?`).get(stockPartCode)?.id || 0) || null;
      if (!stockItemId) throw new Error(`Stock part code ${stockPartCode} was not found`);
    }
    if (stockItemId && !db.prepare(`SELECT id FROM parts WHERE id = ?`).get(stockItemId)) {
      throw new Error(`Stock item ${stockItemId} was not found`);
    }
    return {
      stock_item_id: stockItemId,
      item_type: itemType,
      description,
      quantity_required: Number(quantity.toFixed(3)),
      unit_of_measure: String(item.unit_of_measure || item.uom || "ea").trim() || "ea",
      required: item.required === false || Number(item.required) === 0 ? 0 : 1,
      allow_substitute: item.allow_substitute === true || Number(item.allow_substitute) === 1 ? 1 : 0,
      notes: String(item.notes || "").trim() || null,
    };
  }

  function writeTemplateItems(templateId, items) {
    const insert = db.prepare(`
      INSERT INTO service_template_items (
        service_template_id, stock_item_id, item_type, description, quantity_required,
        unit_of_measure, required, allow_substitute, notes, sort_order
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    (Array.isArray(items) ? items : []).map(cleanTemplateItemInput).forEach((item, index) => {
      insert.run(
        templateId, item.stock_item_id, item.item_type, item.description, item.quantity_required,
        item.unit_of_measure, item.required, item.allow_substitute, item.notes, index,
      );
    });
  }

  function templateSummaryRows({ activeOnly = false } = {}) {
    return db.prepare(`
      SELECT
        t.*,
        COUNT(DISTINCT i.id) AS item_count,
        COUNT(DISTINCT a.id) AS assignment_count
      FROM service_templates t
      LEFT JOIN service_template_items i ON i.service_template_id = t.id
      LEFT JOIN asset_service_template_assignments a ON a.service_template_id = t.id AND a.active = 1
      ${activeOnly ? "WHERE t.active = 1" : ""}
      GROUP BY t.id
      ORDER BY t.template_key ASC, t.revision_number DESC
    `).all().map((row) => ({
      ...row,
      active: Number(row.active) === 1,
      item_count: Number(row.item_count || 0),
      assignment_count: Number(row.assignment_count || 0),
    }));
  }

  function loadTemplateDetail(id) {
    const template = db.prepare(`SELECT * FROM service_templates WHERE id = ?`).get(Number(id));
    if (!template) return null;
    const items = db.prepare(`
      SELECT i.*, p.part_code, p.part_name
      FROM service_template_items i
      LEFT JOIN parts p ON p.id = i.stock_item_id
      WHERE i.service_template_id = ?
      ORDER BY i.sort_order ASC, i.id ASC
    `).all(Number(id));
    const assignments = db.prepare(`
      SELECT a.*, asset.asset_code, asset.asset_name
      FROM asset_service_template_assignments a
      LEFT JOIN assets asset ON asset.id = a.asset_id
      WHERE a.service_template_id = ?
      ORDER BY a.priority DESC, a.id ASC
    `).all(Number(id));
    return { ...template, active: Number(template.active) === 1, items, assignments };
  }

  // =====================================================
  // MTBF / LTTR (reliability) — selected dates & equipment
  // MTBF = operating hours ÷ failure count
  // LTTR = downtime hours ÷ failure count (same as dashboard reliability)
  // GET /api/maintenance/reliability?start=YYYY-MM-DD&end=YYYY-MM-DD&asset_ids=1,2&category=Excavator
  // =====================================================
  function parseReliabilityAssetIds(raw) {
    const out = new Set();
    const s = String(raw || "").trim();
    if (!s) return [];
    s.split(/[,;\s]+/).forEach((part) => {
      const n = Number(part);
      if (Number.isFinite(n) && n > 0) out.add(Math.floor(n));
    });
    return Array.from(out);
  }

  function breakdownDowntimeColumnName() {
    if (hasColumn("breakdowns", "downtime_total_hours")) return "downtime_total_hours";
    if (hasColumn("breakdowns", "downtime_hours")) return "downtime_hours";
    return null;
  }

  function buildMaintenanceReliabilityReport(start, end, opts = {}) {
    const categoryFilter = String(opts.category || "").trim();
    const scheduledFallback = Math.max(0.5, Number(opts.scheduled ?? 10) || 10);
    const siteCode = String(opts.site_code || "main").trim().toLowerCase() || "main";
    let assetIds = Array.isArray(opts.asset_ids) ? opts.asset_ids.map((x) => Number(x)).filter((n) => n > 0) : [];

    let assetRows = [];
    if (assetIds.length) {
      const marks = assetIds.map(() => "?").join(",");
      assetRows = db.prepare(`
        SELECT id AS asset_id, asset_code, asset_name, category
        FROM assets
        WHERE id IN (${marks})
          AND COALESCE(active, 1) = 1
          AND COALESCE(archived, 0) = 0
        ORDER BY asset_code ASC
      `).all(...assetIds);
    } else {
      const params = [];
      let where = `COALESCE(active, 1) = 1 AND COALESCE(archived, 0) = 0`;
      if (categoryFilter) {
        where += ` AND TRIM(COALESCE(category, '')) = ?`;
        params.push(categoryFilter);
      }
      assetRows = db.prepare(`
        SELECT id AS asset_id, asset_code, asset_name, category
        FROM assets
        WHERE ${where}
        ORDER BY asset_code ASC
      `).all(...params);
    }

    assetIds = assetRows.map((r) => Number(r.asset_id || 0)).filter((n) => n > 0);
    if (!assetIds.length) {
      return {
        start,
        end,
        category: categoryFilter || null,
        asset_filter_count: 0,
        summary: {
          failure_count: 0,
          operating_hours: 0,
          downtime_hours: 0,
          mtbf_hours: null,
          lttr_hours: null,
        },
        by_asset: [],
      };
    }

    const marks = assetIds.map(() => "?").join(",");

    const runByAsset = new Map(
      db.prepare(`
        SELECT asset_id, COALESCE(SUM(hours_run), 0) AS run_hours
        FROM daily_hours
        WHERE work_date BETWEEN ? AND ?
          AND is_used = 1
          AND hours_run > 0
          AND asset_id IN (${marks})
        GROUP BY asset_id
      `).all(start, end, ...assetIds).map((r) => [Number(r.asset_id || 0), Number(r.run_hours || 0)])
    );

    const { incidents, byAsset: reliabilityByAsset } = buildReliabilityIncidentsForAssets(db, {
      assetIds,
      start,
      end,
      hasTable,
      hasColumn,
    });

    // Use the same daily KPI engine as Asset KPI for production hours and
    // downtime. This includes explicit downtime, historical header fallback,
    // and only the same open-breakdown imputation used for availability.
    // If the dashboard provider is unavailable during an isolated route test,
    // retain the recorded-downtime calculation as a safe fallback.
    const buildAssetKpiRange = getAssetKpiRangeBuilder();
    const kpiRange = buildAssetKpiRange
      ? buildAssetKpiRange(
        start,
        end,
        scheduledFallback,
        siteCode,
        assetRows.map((a) => String(a.asset_code || "").trim()).filter(Boolean),
      )
      : null;
    const kpiByAsset = new Map(
      (Array.isArray(kpiRange?.by_asset) ? kpiRange.by_asset : [])
        .map((row) => [Number(row.asset_id || 0), row])
        .filter(([assetId]) => assetId > 0)
    );

    const by_asset = assetRows.map((a) => {
      const aid = Number(a.asset_id || 0);
      const kpi = kpiByAsset.get(aid) || null;
      const operating_hours = kpi ? Number(kpi.run_hours || 0) : Number(runByAsset.get(aid) || 0);
      const rel = reliabilityByAsset.get(aid) || { failure_count: 0, downtime_hours: 0 };
      const failure_count = Number(rel.failure_count || 0);
      const recorded_downtime_hours = Number(rel.downtime_hours || 0);
      const downtime_hours = kpi ? Number(kpi.downtime_hours || 0) : recorded_downtime_hours;
      const { mtbf_hours, lttr_hours } = computeMtbfLttr(operating_hours, failure_count, downtime_hours);
      return {
        asset_id: aid,
        asset_code: String(a.asset_code || ""),
        asset_name: String(a.asset_name || ""),
        category: String(a.category || ""),
        failure_count,
        operating_hours: round2(operating_hours),
        downtime_hours: round2(downtime_hours),
        recorded_downtime_hours: round2(recorded_downtime_hours),
        mtbf_hours,
        lttr_hours,
      };
    }).sort((x, y) => {
      const xm = x.mtbf_hours == null ? Infinity : Number(x.mtbf_hours);
      const ym = y.mtbf_hours == null ? Infinity : Number(y.mtbf_hours);
      return xm - ym;
    });

    const failure_count = by_asset.reduce((s, r) => s + Number(r.failure_count || 0), 0);
    const operating_hours = by_asset.reduce((s, r) => s + Number(r.operating_hours || 0), 0);
    const downtime_hours = by_asset.reduce((s, r) => s + Number(r.downtime_hours || 0), 0);
    const recorded_downtime_hours = by_asset.reduce((s, r) => s + Number(r.recorded_downtime_hours || 0), 0);
    const { mtbf_hours, lttr_hours } = computeMtbfLttr(operating_hours, failure_count, downtime_hours);

    const incidentsWithAsset = incidents.map((inc) => {
      const a = assetRows.find((r) => Number(r.asset_id) === Number(inc.asset_id));
      return {
        ...inc,
        asset_code: String(a?.asset_code || ""),
        asset_name: String(a?.asset_name || ""),
      };
    }).sort((x, y) => {
      const c = String(x.asset_code || "").localeCompare(String(y.asset_code || ""));
      if (c !== 0) return c;
      return String(x.breakdown_date || "").localeCompare(String(y.breakdown_date || ""));
    });

    return {
      start,
      end,
      category: categoryFilter || null,
      asset_filter_count: assetIds.length,
      scheduled_fallback: scheduledFallback,
      downtime_basis: kpiRange ? "asset_kpi_daily" : "recorded_breakdown_downtime",
      formulas: {
        mtbf: "operating_hours / failure_count",
        lttr: "downtime_hours / failure_count",
        failures: "distinct breakdown incidents with recorded downtime in period (daily logs, else breakdown header when reported in period)",
        operating_hours: "Asset KPI daily run hours for the selected equipment",
        downtime_hours: "Asset KPI daily downtime for the selected equipment (logged/header downtime plus the same open-breakdown imputation used by availability)",
      },
      summary: {
        failure_count,
        operating_hours: round2(operating_hours),
        downtime_hours: round2(downtime_hours),
        recorded_downtime_hours: round2(recorded_downtime_hours),
        mtbf_hours,
        lttr_hours,
      },
      by_asset,
      incidents: incidentsWithAsset,
    };
  }

  /**
   * Downtime for Maintenance Insights: sum daily breakdown_downtime_logs in range,
   * impute missing days for open/in-progress breakdowns, then legacy header fallback.
   */
  function buildMaintenanceInsightsDowntime(startDate, endDate, opts = {}) {
    const scheduledFallback = Math.max(1, Number(opts.scheduledFallback || 10));
    const empty = { by_component: [], by_team: [], downtime_daily: [], by_asset: new Map(), total_hours: 0 };
    if (!hasTable("breakdowns")) return empty;

    const canReadDowntimeLogs = hasTable("breakdown_downtime_logs");
    const woAssignedCol = hasColumn("work_orders", "assigned_artisan_name")
      ? "assigned_artisan_name"
      : (hasColumn("work_orders", "artisan_name") ? "artisan_name" : "");
    const breakdownDateExpr = hasColumn("breakdowns", "breakdown_date")
      ? "b.breakdown_date"
      : "DATE(COALESCE(b.created_at, b.updated_at))";
    const teamExpr = woAssignedCol
      ? `COALESCE(NULLIF(TRIM(wo.${woAssignedCol}), ''), 'Unassigned')`
      : `'Unassigned'`;

    const componentMap = new Map();
    const teamMap = new Map();
    const dailyMap = new Map();
    const assetMap = new Map();
    const loggedBreakdownDays = new Set();
    let totalHours = 0;

    const titleCase = (s) => String(s || "uncategorized").replace(/\b\w/g, (m) => m.toUpperCase());

    const addHours = (breakdownId, componentKey, team, day, hours, assetId = 0) => {
      const hrs = Number(hours || 0);
      const bid = Number(breakdownId || 0);
      if (!Number.isFinite(hrs) || hrs <= 0 || !bid) return;
      totalHours += hrs;

      const ck = String(componentKey || "uncategorized").trim().toLowerCase() || "uncategorized";
      const comp = componentMap.get(ck) || { component_key: ck, incidents: new Set(), hours: 0 };
      comp.incidents.add(bid);
      comp.hours += hrs;
      componentMap.set(ck, comp);

      const tk = String(team || "Unassigned").trim() || "Unassigned";
      const tm = teamMap.get(tk) || { team: tk, incidents: new Set(), hours: 0 };
      tm.incidents.add(bid);
      tm.hours += hrs;
      teamMap.set(tk, tm);

      if (day) dailyMap.set(day, Number(dailyMap.get(day) || 0) + hrs);

      const aid = Number(assetId || 0);
      if (aid > 0) {
        const ar = assetMap.get(aid) || { asset_id: aid, downtime_hours: 0 };
        ar.downtime_hours += hrs;
        assetMap.set(aid, ar);
      }
    };

    if (canReadDowntimeLogs) {
      const logRows = db.prepare(`
        SELECT
          b.id AS breakdown_id,
          b.asset_id,
          LOWER(TRIM(COALESCE(b.component, 'uncategorized'))) AS component_key,
          ${teamExpr} AS team,
          DATE(l.log_date) AS log_date,
          COALESCE(l.hours_down, 0) AS hours_down
        FROM breakdown_downtime_logs l
        JOIN breakdowns b ON b.id = l.breakdown_id
        LEFT JOIN work_orders wo ON wo.id = b.primary_work_order_id
        WHERE DATE(l.log_date) BETWEEN DATE(?) AND DATE(?)
          AND COALESCE(l.hours_down, 0) > 0
      `).all(startDate, endDate);
      for (const r of logRows) {
        loggedBreakdownDays.add(`${Number(r.breakdown_id || 0)}|${String(r.log_date || "")}`);
        addHours(
          r.breakdown_id,
          r.component_key,
          r.team,
          r.log_date,
          r.hours_down,
          r.asset_id,
        );
      }
    }

    const todayYmd = new Date().toISOString().slice(0, 10);
    const imputeEndDay = endDate < todayYmd ? endDate : todayYmd;
    const days = listDaysInclusiveYmd(startDate, imputeEndDay);
    const getScheduledForAssetDay = hasTable("daily_hours")
      ? db.prepare(`
          SELECT scheduled_hours, is_used, hours_run
          FROM daily_hours
          WHERE asset_id = ? AND work_date = ?
          LIMIT 1
        `)
      : null;
    const getLogForBreakdownDay = canReadDowntimeLogs
      ? db.prepare(`
          SELECT hours_down
          FROM breakdown_downtime_logs
          WHERE breakdown_id = ? AND log_date = ? AND COALESCE(hours_down, 0) > 0
          LIMIT 1
        `)
      : null;

    const openRows = db.prepare(`
      SELECT
        b.id AS breakdown_id,
        b.asset_id,
        LOWER(TRIM(COALESCE(b.component, 'uncategorized'))) AS component_key,
        DATE(COALESCE(${breakdownDateExpr}, b.created_at)) AS breakdown_day,
        ${teamExpr} AS team
      FROM breakdowns b
      LEFT JOIN work_orders wo ON wo.id = b.primary_work_order_id
      WHERE b.status = 'OPEN'
        AND (wo.id IS NULL OR LOWER(TRIM(COALESCE(wo.status, ''))) NOT IN ('completed', 'approved', 'closed'))
        AND DATE(COALESCE(${breakdownDateExpr}, b.created_at)) <= DATE(?)
    `).all(imputeEndDay);

    for (const br of openRows) {
      const bid = Number(br.breakdown_id || 0);
      const aid = Number(br.asset_id || 0);
      if (!bid || !aid) continue;
      const breakdownDay = String(br.breakdown_day || startDate);
      for (const day of days) {
        if (day < breakdownDay || day < startDate) continue;
        const logKey = `${bid}|${day}`;
        if (loggedBreakdownDays.has(logKey)) continue;
        if (getLogForBreakdownDay?.get(bid, day)) continue;

        const dh = getScheduledForAssetDay?.get(aid, day);
        const rowScheduled = Number(dh?.scheduled_hours);
        const isUsed = Number(dh?.is_used ?? 1);
        const runHours = Number(dh?.hours_run || 0);
        let impute = scheduledFallback;
        if (Number.isFinite(rowScheduled) && rowScheduled > 0) {
          impute = rowScheduled;
        } else if (isUsed !== 1 && runHours <= 0) {
          impute = scheduledFallback;
        }
        addHours(bid, br.component_key, br.team, day, impute, aid);
      }
    }

    if (totalHours <= 0) {
      const dtCol = breakdownDowntimeColumnName();
      if (dtCol) {
        const fallbackRows = db.prepare(`
          SELECT
            b.id AS breakdown_id,
            b.asset_id,
            LOWER(TRIM(COALESCE(b.component, 'uncategorized'))) AS component_key,
            ${teamExpr} AS team,
            DATE(COALESCE(${breakdownDateExpr}, b.created_at)) AS breakdown_day,
            COALESCE(b.${dtCol}, 0) AS downtime_hours
          FROM breakdowns b
          LEFT JOIN work_orders wo ON wo.id = b.primary_work_order_id
          WHERE DATE(COALESCE(${breakdownDateExpr}, b.created_at)) BETWEEN DATE(?) AND DATE(?)
            AND COALESCE(b.${dtCol}, 0) > 0
        `).all(startDate, endDate);
        for (const r of fallbackRows) {
          addHours(
            r.breakdown_id,
            r.component_key,
            r.team,
            r.breakdown_day,
            r.downtime_hours,
            r.asset_id,
          );
        }
      }
    }

    return {
      by_component: [...componentMap.values()]
        .map((r) => ({
          component: titleCase(r.component_key),
          incidents: r.incidents.size,
          downtime_hours: Number(r.hours.toFixed(2)),
        }))
        .sort((a, b) => Number(b.downtime_hours || 0) - Number(a.downtime_hours || 0))
        .slice(0, 20),
      by_team: [...teamMap.values()]
        .map((r) => ({
          team: r.team,
          incidents: r.incidents.size,
          downtime_hours: Number(r.hours.toFixed(2)),
        }))
        .sort((a, b) => Number(b.downtime_hours || 0) - Number(a.downtime_hours || 0))
        .slice(0, 20),
      downtime_daily: listDaysInclusiveYmd(startDate, endDate)
        .map((day) => ({
          day,
          downtime_hours: Number(Number(dailyMap.get(day) || 0).toFixed(2)),
        }))
        .filter((r) => r.downtime_hours > 0),
      by_asset: assetMap,
      total_hours: Number(totalHours.toFixed(2)),
    };
  }

  async function loadInsightsExportPayload(req) {
    const q = new URLSearchParams();
    const copy = (name, fallback = "") => {
      const v = String(req.query?.[name] ?? fallback).trim();
      if (v !== "") q.set(name, v);
    };
    copy("start");
    copy("end");
    copy("near_due_hours", "50");
    copy("predictive_horizon_hours", "100");
    copy("checklist_fail_threshold", "2");
    copy("fuel_variance_threshold", "15");

    const injected = await app.inject({
      method: "GET",
      url: `/api/maintenance/insights?${q.toString()}`,
      headers: {
        "x-user-name": String(req.headers?.["x-user-name"] || "system"),
        "x-user-role": String(req.headers?.["x-user-role"] || "admin"),
        "x-user-roles": String(req.headers?.["x-user-roles"] || "admin"),
        "x-site-code": String(req.headers?.["x-site-code"] || "main"),
      },
    });
    if (injected.statusCode >= 400) {
      let payload = {};
      try { payload = JSON.parse(String(injected.payload || "{}")); } catch {}
      const err = new Error(payload?.error || "Failed to build insights export");
      err.statusCode = injected.statusCode;
      throw err;
    }
    const data = JSON.parse(String(injected.payload || "{}"));
    if (!String(data?.range?.start || "").trim()) {
      const err = new Error("Maintenance insights returned an empty payload");
      err.statusCode = 500;
      throw err;
    }
    return data;
  }

  function addInsightsExportSummarySheet(wb, data) {
    const wsSummary = wb.addWorksheet("Summary");
    wsSummary.columns = [{ header: "Field", key: "field", width: 34 }, { header: "Value", key: "value", width: 30 }];
    wsSummary.addRows([
      { field: "Start", value: data?.range?.start || "" },
      { field: "End", value: data?.range?.end || "" },
      { field: "Near Due Hours", value: Number(data?.range?.near_due_hours || 0) },
      { field: "Predictive Horizon Hours", value: Number(data?.range?.predictive_horizon_hours || 0) },
      { field: "Upcoming services", value: Number(data?.parts_planning?.upcoming_service_count || 0) },
      { field: "Forecast total ($)", value: Number(data?.parts_planning?.total_upcoming_cost || 0) },
      { field: "Needs manual input", value: Number(data?.parts_planning?.needs_manual_input_count || 0) },
      { field: "Labor rate ($/hr)", value: Number(data?.range?.labor_cost_per_hour || 0) },
    ]);
    wsSummary.getRow(1).font = { bold: true };
  }

  function addInsightsPartsDemandSheets(wb, data) {
    const wsUpcomingCost = wb.addWorksheet("Upcoming Service Costs");
    wsUpcomingCost.columns = [
      { header: "Asset Code", key: "asset_code", width: 14 },
      { header: "Asset Name", key: "asset_name", width: 24 },
      { header: "Service", key: "service_name", width: 20 },
      { header: "Remaining Hrs", key: "remaining_hours", width: 14 },
      { header: "Status", key: "status", width: 12 },
      { header: "Kit Cost", key: "est_service_kit_cost", width: 12 },
      { header: "Labor Cost", key: "est_labor_cost", width: 12 },
      { header: "Total Cost", key: "est_total_cost", width: 12 },
      { header: "Cost Source", key: "cost_source", width: 18 },
      { header: "Needs Manual Input", key: "needs_manual_input", width: 16 },
    ];
    wsUpcomingCost.addRows(
      (Array.isArray(data?.parts_planning?.upcoming_cost_forecasts) ? data.parts_planning.upcoming_cost_forecasts : []).map((r) => ({
        asset_code: r.asset_code,
        asset_name: r.asset_name,
        service_name: r.service_name,
        remaining_hours: Number(r.remaining_hours || 0),
        status: r.status,
        est_service_kit_cost: Number(r?.forecast?.est_service_kit_cost || 0),
        est_labor_cost: Number(r?.forecast?.est_labor_cost || 0),
        est_total_cost: Number(r?.forecast?.est_total_cost || 0),
        cost_source: r?.forecast?.cost_source || "",
        needs_manual_input: r.needs_manual_input ? "yes" : "no",
      })),
    );

    const wsParts = wb.addWorksheet("Parts Demand");
    wsParts.columns = [
      { header: "Part", key: "part_name", width: 30 },
      { header: "Suggested Qty", key: "suggested_qty", width: 14 },
      { header: "Est Cost", key: "est_cost", width: 12 },
      { header: "On Hand", key: "on_hand", width: 12 },
      { header: "Gap Qty", key: "gap_qty", width: 12 },
      { header: "Linked Services", key: "linked_services", width: 36 },
    ];
    wsParts.addRows((Array.isArray(data?.parts_planning?.suggestions) ? data.parts_planning.suggestions : []).map((r) => ({
      ...r,
      linked_services: Array.isArray(r?.linked_services) ? r.linked_services.join(", ") : "",
    })));
    wsUpcomingCost.getRow(1).font = { bold: true };
    wsParts.getRow(1).font = { bold: true };
  }

  function addInsightsCostPerMachineSheet(wb, data) {
    const wsCost = wb.addWorksheet("Cost Per Machine");
    wsCost.columns = [
      { header: "Asset Code", key: "asset_code", width: 14 },
      { header: "Asset Name", key: "asset_name", width: 28 },
      { header: "Service Jobs", key: "service_jobs", width: 12 },
      { header: "Down Hrs", key: "downtime_hours", width: 12 },
      { header: "Repair Labor Hrs", key: "repair_labor_hours", width: 16 },
      { header: "WO Labor $", key: "wo_labor_cost", width: 12 },
      { header: "Repair Labor $", key: "repair_labor_cost", width: 14 },
      { header: "Total Labor $", key: "labor_cost", width: 12 },
      { header: "Parts Cost", key: "parts_cost", width: 12 },
      { header: "Lube Cost", key: "lube_cost", width: 12 },
      { header: "Outsourced Cost", key: "outsourced_cost", width: 14 },
      { header: "Total Cost", key: "total_cost", width: 12 },
    ];
    wsCost.addRows(Array.isArray(data?.maintenance_cost) ? data.maintenance_cost : []);
    wsCost.getRow(1).font = { bold: true };
    ["downtime_hours", "repair_labor_hours", "wo_labor_cost", "repair_labor_cost", "labor_cost", "parts_cost", "lube_cost", "outsourced_cost", "total_cost"].forEach((key) => {
      wsCost.getColumn(key).numFmt = "#,##0.00";
    });
  }

  async function buildWeeklyForumSummary(query = {}) {
      const nearDueHours = Math.max(1, Number(query?.near_due_hours || 50));
      const startIn = String(query?.start || "").trim();
      const endIn = String(query?.end || "").trim();

      const now = new Date();
      const day = now.getDay(); // 0=Sun ... 6=Sat
      const mondayOffset = day === 0 ? -6 : 1 - day;
      const monday = new Date(now);
      monday.setHours(0, 0, 0, 0);
      monday.setDate(monday.getDate() + mondayOffset);
      const friday = new Date(monday);
      friday.setDate(friday.getDate() + 4);
      const ymd = (d) => d.toISOString().slice(0, 10);
      const start = startIn && isDate(startIn) ? startIn : ymd(monday);
      const end = endIn && isDate(endIn) ? endIn : ymd(friday);
      if (!isDate(start) || !isDate(end)) {
        const e = new Error("start and end must be YYYY-MM-DD");
        e.statusCode = 400;
        throw e;
      }
      if (start > end) {
        const e = new Error("start must be <= end");
        e.statusCode = 400;
        throw e;
      }

      const hasTable = (name) =>
        Boolean(
          db.prepare(`
            SELECT 1
            FROM sqlite_master
            WHERE type = 'table' AND name = ?
            LIMIT 1
          `).get(name)
        );
      const hasColumn = (table, col) => {
        if (!hasTable(table)) return false;
        const rows = db.prepare(`PRAGMA table_info(${table})`).all();
        return rows.some((r) => String(r.name || "") === col);
      };

      const plans = hasTable("maintenance_plans")
        ? db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              a.asset_code,
              a.asset_name
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.archived = 0
              AND a.is_standby = 0
            ORDER BY a.asset_code ASC, mp.service_name ASC
          `).all()
        : [];

      const openWOs =
        hasTable("work_orders")
          ? Number(
              db.prepare(`
                SELECT COUNT(*) AS c
                FROM work_orders
                WHERE LOWER(COALESCE(status, 'open')) NOT IN ('closed', 'completed', 'approved')
              `).get()?.c || 0
            )
          : 0;

      const closedStatuses = "'closed','completed','approved'";
      const hasWOCompletedAt = hasColumn("work_orders", "completed_at");
      const woCloseExpr = hasWOCompletedAt ? "COALESCE(w.completed_at, w.closed_at)" : "w.closed_at";
      const smOutSql = sqlStockMovementOutbound("sm");
      const oilPartSql = sqlOilPartPredicate("p");
      const smLineFlags = {
        hasSmTotalCost: hasColumn("stock_movements", "total_cost"),
        hasSmUnitCost: hasColumn("stock_movements", "unit_cost"),
        hasPartsUnitCost: hasTable("parts") && hasColumn("parts", "unit_cost"),
        hasSmUnitCostUsd: hasColumn("stock_movements", "unit_cost_usd"),
        hasSmCostInput: hasColumn("stock_movements", "cost_input"),
      };
      const smCostWithParts = sqlStockMovementLineCostExpr("sm", "p", true, smLineFlags);
      const smCostNoParts = sqlStockMovementLineCostExpr("sm", "p", false, smLineFlags);
      const lubeCostDefault = readLubeCostPerQtyDefault(db, hasTable);
      const smDateExpr = sqlStockMovementDateExpr(hasColumn);

      const partsCost =
        hasTable("stock_movements") && hasTable("work_orders")
          ? Number(
              hasTable("parts")
                ? db.prepare(`
                    SELECT COALESCE(SUM((${smCostWithParts})), 0) AS v
                    FROM stock_movements sm
                    JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
                    JOIN parts p ON p.id = sm.part_id
                    WHERE ${woCloseExpr} IS NOT NULL
                      AND DATE(${woCloseExpr}) BETWEEN ? AND ?
                      AND (${smOutSql})
                      AND NOT (${oilPartSql})
                  `).get(start, end)?.v || 0
                : db.prepare(`
                    SELECT COALESCE(SUM((${smCostNoParts})), 0) AS v
                    FROM stock_movements sm
                    JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
                    WHERE ${woCloseExpr} IS NOT NULL
                      AND DATE(${woCloseExpr}) BETWEEN ? AND ?
                      AND (${smOutSql})
                  `).get(start, end)?.v || 0
            )
          : 0;

      const oilCostFromLogs = hasTable("oil_logs")
        ? Number(
            db.prepare(`
              SELECT ${sqlOilLogCostSumExpr("ol", "?")} AS v
              FROM oil_logs ol
              WHERE ol.log_date BETWEEN ? AND ?
            `).get(lubeCostDefault, start, end)?.v || 0
          )
        : 0;

      const oilCostFromDirectAssetStores =
        hasTable("stock_movements") && hasTable("parts")
          ? Number(
              db.prepare(`
                SELECT COALESCE(SUM((${smCostWithParts})), 0) AS v
                FROM stock_movements sm
                JOIN parts p ON p.id = sm.part_id
                WHERE sm.reference LIKE 'asset:%:stores'
                  AND ${smDateExpr} BETWEEN ? AND ?
                  AND (${smOutSql})
                  AND (${oilPartSql})
              `).get(start, end)?.v || 0
            )
          : 0;

      const oilCostFromWoStock =
        hasTable("stock_movements") && hasTable("work_orders") && hasTable("parts")
          ? Number(
              db.prepare(`
                SELECT COALESCE(SUM((${smCostWithParts})), 0) AS v
                FROM stock_movements sm
                JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
                JOIN parts p ON p.id = sm.part_id
                WHERE ${woCloseExpr} IS NOT NULL
                  AND DATE(${woCloseExpr}) BETWEEN ? AND ?
                  AND (${smOutSql})
                  AND (${oilPartSql})
              `).get(start, end)?.v || 0
            )
          : 0;

      const oilCost = oilCostFromLogs + oilCostFromWoStock + oilCostFromDirectAssetStores;

      const laborCost = hasTable("work_orders")
        ? Number(
            db.prepare(`
              SELECT COALESCE(SUM(COALESCE(labor_hours, 0) * COALESCE(labor_rate_per_hour, 0)), 0) AS v
              FROM work_orders w
              WHERE ${woCloseExpr} IS NOT NULL
                AND DATE(${woCloseExpr}) BETWEEN ? AND ?
            `).get(start, end)?.v || 0
          )
        : 0;

      /** Per-asset actual consumption costs for the selected date range (historical, not forecast). */
      const hasLaborHoursCol = hasColumn("work_orders", "labor_hours");
      const hasLaborRateCol = hasColumn("work_orders", "labor_rate_per_hour");
      const mergeActual = new Map();
      const putActual = (assetId, assetCode, assetName, patch) => {
        const aid = Number(assetId || 0);
        if (!aid) return;
        const cur = mergeActual.get(aid) || {
          asset_id: aid,
          asset_code: String(assetCode || ""),
          asset_name: String(assetName || ""),
          parts_cost: 0,
          lubes_logs_cost: 0,
          lubes_work_order_cost: 0,
          labor_cost: 0,
          closed_work_orders: 0,
        };
        if (assetCode) cur.asset_code = String(assetCode);
        if (assetName) cur.asset_name = String(assetName);
        Object.assign(cur, patch);
        mergeActual.set(aid, cur);
      };
      if (hasTable("work_orders") && hasTable("assets")) {
        if (hasTable("stock_movements")) {
          const partRows = hasTable("parts")
            ? db.prepare(`
                SELECT w.asset_id, a.asset_code, a.asset_name,
                  COALESCE(SUM((${smCostWithParts})), 0) AS v
                FROM stock_movements sm
                JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
                JOIN assets a ON a.id = w.asset_id
                JOIN parts p ON p.id = sm.part_id
                WHERE ${woCloseExpr} IS NOT NULL
                  AND DATE(${woCloseExpr}) BETWEEN ? AND ?
                  AND (${smOutSql})
                  AND NOT (${oilPartSql})
                GROUP BY w.asset_id
              `).all(start, end)
            : db.prepare(`
                SELECT w.asset_id, a.asset_code, a.asset_name,
                  COALESCE(SUM((${smCostNoParts})), 0) AS v
                FROM stock_movements sm
                JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
                JOIN assets a ON a.id = w.asset_id
                WHERE ${woCloseExpr} IS NOT NULL
                  AND DATE(${woCloseExpr}) BETWEEN ? AND ?
                  AND (${smOutSql})
                GROUP BY w.asset_id
              `).all(start, end);
          for (const r of partRows || []) putActual(r.asset_id, r.asset_code, r.asset_name, { parts_cost: Number(r.v || 0) });
        }
        if (hasTable("oil_logs")) {
          const oilLogRows = db.prepare(`
            SELECT ol.asset_id, a.asset_code, a.asset_name,
              ${sqlOilLogCostSumExpr("ol", "?")} AS v
            FROM oil_logs ol
            JOIN assets a ON a.id = ol.asset_id
            WHERE ol.log_date BETWEEN ? AND ?
            GROUP BY ol.asset_id
          `).all(lubeCostDefault, start, end);
          for (const r of oilLogRows || []) putActual(r.asset_id, r.asset_code, r.asset_name, { lubes_logs_cost: Number(r.v || 0) });
        }
        if (hasTable("stock_movements") && hasTable("parts")) {
          const directOilRows = db.prepare(`
            SELECT
              a.id AS asset_id,
              a.asset_code,
              a.asset_name,
              COALESCE(SUM((${smCostWithParts})), 0) AS v
            FROM stock_movements sm
            JOIN parts p ON p.id = sm.part_id
            JOIN assets a ON a.id = CAST(REPLACE(REPLACE(sm.reference, 'asset:', ''), ':stores', '') AS INTEGER)
            WHERE sm.reference LIKE 'asset:%:stores'
              AND ${smDateExpr} BETWEEN ? AND ?
              AND (${smOutSql})
              AND (${oilPartSql})
            GROUP BY a.id
          `).all(start, end);
          for (const r of directOilRows || []) {
            const aid = Number(r.asset_id || 0);
            const cur = mergeActual.get(aid);
            const prior = cur ? Number(cur.lubes_logs_cost || 0) : 0;
            putActual(r.asset_id, r.asset_code, r.asset_name, {
              lubes_logs_cost: prior + Number(r.v || 0),
            });
          }
        }
        if (hasTable("stock_movements") && hasTable("parts")) {
          const oilWoRows = db.prepare(`
            SELECT w.asset_id, a.asset_code, a.asset_name,
              COALESCE(SUM((${smCostWithParts})), 0) AS v
            FROM stock_movements sm
            JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
            JOIN assets a ON a.id = w.asset_id
            JOIN parts p ON p.id = sm.part_id
            WHERE ${woCloseExpr} IS NOT NULL
              AND DATE(${woCloseExpr}) BETWEEN ? AND ?
              AND (${smOutSql})
              AND (${oilPartSql})
            GROUP BY w.asset_id
          `).all(start, end);
          for (const r of oilWoRows || []) putActual(r.asset_id, r.asset_code, r.asset_name, { lubes_work_order_cost: Number(r.v || 0) });
        }
        if (hasLaborHoursCol && hasLaborRateCol) {
          const labRows = db.prepare(`
            SELECT w.asset_id, a.asset_code, a.asset_name,
              COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, 0)), 0) AS v
            FROM work_orders w
            JOIN assets a ON a.id = w.asset_id
            WHERE ${woCloseExpr} IS NOT NULL
              AND DATE(${woCloseExpr}) BETWEEN ? AND ?
            GROUP BY w.asset_id
          `).all(start, end);
          for (const r of labRows || []) putActual(r.asset_id, r.asset_code, r.asset_name, { labor_cost: Number(r.v || 0) });
        }
        const cntRows = db.prepare(`
          SELECT w.asset_id, a.asset_code, a.asset_name, COUNT(*) AS c
          FROM work_orders w
          JOIN assets a ON a.id = w.asset_id
          WHERE ${woCloseExpr} IS NOT NULL
            AND DATE(${woCloseExpr}) BETWEEN ? AND ?
            AND LOWER(COALESCE(w.status, '')) IN (${closedStatuses})
          GROUP BY w.asset_id
        `).all(start, end);
        for (const r of cntRows || []) putActual(r.asset_id, r.asset_code, r.asset_name, { closed_work_orders: Number(r.c || 0) });
      }
      const periodActualsByAsset = Array.from(mergeActual.values())
        .map((r) => {
          const lubes = Number(r.lubes_logs_cost || 0) + Number(r.lubes_work_order_cost || 0);
          const total = Number(r.parts_cost || 0) + lubes + Number(r.labor_cost || 0);
          return {
            asset_id: r.asset_id,
            asset_code: r.asset_code,
            asset_name: r.asset_name,
            parts_cost: Number(Number(r.parts_cost || 0).toFixed(2)),
            lubes_logs_cost: Number(Number(r.lubes_logs_cost || 0).toFixed(2)),
            lubes_work_order_cost: Number(Number(r.lubes_work_order_cost || 0).toFixed(2)),
            lubes_total_cost: Number(lubes.toFixed(2)),
            labor_cost: Number(Number(r.labor_cost || 0).toFixed(2)),
            period_total_cost: Number(total.toFixed(2)),
            closed_work_orders: Number(r.closed_work_orders || 0),
          };
        })
        .sort((a, b) => Number(b.period_total_cost || 0) - Number(a.period_total_cost || 0));

      const getAssetCurrentHoursSafe = (assetId) => {
        try {
          return getAssetCurrentHours(assetId);
        } catch {
          return 0;
        }
      };
      const forecastRows = buildUpcomingServiceCostForecasts(db, plans, {
        nearDueHours,
        maxRemainingHours: nearDueHours,
        ctx: {
          hasTable,
          hasColumn,
          closedStatuses,
          woCloseExpr,
          smOutSql,
          oilPartSql,
          smCostWithParts,
          smCostNoParts,
          lubeCostDefault,
        },
      }).filter((r) => r.status === "OVERDUE" || r.status === "ALMOST DUE")
        .slice(0, 40);

      const totalForecastCost = forecastRows.reduce(
        (s, r) => s + Number(r.forecast?.est_total_cost || r.forecast?.est_service_kit_cost || 0),
        0
      );

      return {
        ok: true,
        range: { start, end },
        kpis: {
          open_work_orders: openWOs,
          upcoming_services_flagged: forecastRows.length,
        },
        costs: {
          stores_oil_cost: Number(oilCost.toFixed(2)),
          stores_oil_from_logs: Number((oilCostFromLogs + oilCostFromDirectAssetStores).toFixed(2)),
          stores_oil_from_work_orders: Number(oilCostFromWoStock.toFixed(2)),
          stores_oil_from_direct_asset_issues: Number(oilCostFromDirectAssetStores.toFixed(2)),
          stores_parts_cost: Number(partsCost.toFixed(2)),
          maintenance_labor_cost: Number(laborCost.toFixed(2)),
          weekly_total_cost: Number((oilCost + partsCost + laborCost).toFixed(2)),
          upcoming_service_forecast_cost: Number(totalForecastCost.toFixed(2)),
        },
        /** Closed-WO-dated parts/lube stock + oil_logs + labor in [start,end], grouped by equipment. */
        period_actuals_by_asset: periodActualsByAsset,
        upcoming_services: forecastRows,
      };
  }

  // =====================================================
  // WEEKLY FORUM ACTION TRACKER
  // =====================================================
  db.prepare(`
    CREATE TABLE IF NOT EXISTS weekly_forum_actions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      action_date TEXT NOT NULL,
      department TEXT NOT NULL,
      action_item TEXT NOT NULL,
      owner_name TEXT NOT NULL,
      due_date TEXT,
      status TEXT NOT NULL DEFAULT 'open',
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS weekly_forum_review_notes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      period_start TEXT NOT NULL,
      period_end TEXT NOT NULL,
      area TEXT NOT NULL,
      weekly_finding TEXT NOT NULL,
      action_owner TEXT NOT NULL,
      due_date TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(period_start, period_end, area)
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS weekly_forum_service_inputs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      plan_id INTEGER NOT NULL UNIQUE,
      oil_part_code TEXT,
      oil_qty REAL NOT NULL DEFAULT 0,
      parts_part_code TEXT,
      parts_qty REAL NOT NULL DEFAULT 0,
      items_json TEXT NOT NULL DEFAULT '[]',
      notes TEXT,
      labor_total REAL NOT NULL DEFAULT 0,
      all_in_total REAL NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  try {
    const wfInputCols = db.prepare(`PRAGMA table_info(weekly_forum_service_inputs)`).all();
    const wfInputHasItems = wfInputCols.some((c) => String(c?.name || "") === "items_json");
    if (!wfInputHasItems) {
      db.prepare(`ALTER TABLE weekly_forum_service_inputs ADD COLUMN items_json TEXT NOT NULL DEFAULT '[]'`).run();
    }
    const wfInputHasLabor = wfInputCols.some((c) => String(c?.name || "") === "labor_total");
    if (!wfInputHasLabor) {
      db.prepare(`ALTER TABLE weekly_forum_service_inputs ADD COLUMN labor_total REAL NOT NULL DEFAULT 0`).run();
    }
    const wfInputHasAllInTotal = wfInputCols.some((c) => String(c?.name || "") === "all_in_total");
    if (!wfInputHasAllInTotal) {
      db.prepare(`ALTER TABLE weekly_forum_service_inputs ADD COLUMN all_in_total REAL NOT NULL DEFAULT 0`).run();
    }
  } catch {}

  function normalizeInspectionChecklist(raw) {
    if (!Array.isArray(raw)) return [];
    return raw
      .map((x) => ({
        key: String(x?.key || "").trim(),
        label: String(x?.label || "").trim(),
        ok: x?.ok === true ? true : x?.ok === false ? false : null,
        note: String(x?.note || "").trim() || null,
      }))
      .filter((x) => x.key && x.label);
  }

  function normalizeArtisanInspectionChecklist(raw) {
    if (!Array.isArray(raw)) return [];
    return raw
      .map((x) => ({
        key: String(x?.key || "").trim(),
        label: String(x?.label || "").trim(),
        ok: x?.ok === true ? true : x?.ok === false ? false : null,
        note: String(x?.note || "").trim() || null,
      }))
      .filter((x) => x.key && x.label);
  }

  function normalizeInspectionParts(raw) {
    if (!Array.isArray(raw)) return [];
    const out = [];
    for (const x of raw) {
      const part_code = String(x?.part_code || "").trim();
      const qty = Math.max(0, Number(x?.qty || 0));
      if (!part_code || !Number.isFinite(qty) || qty <= 0) continue;
      let part_id = x?.part_id != null ? Number(x.part_id) : null;
      if (!part_id || part_id <= 0) {
        const pr = db.prepare(`SELECT id FROM parts WHERE UPPER(TRIM(part_code)) = UPPER(TRIM(?)) LIMIT 1`).get(part_code);
        part_id = pr?.id != null ? Number(pr.id) : null;
      }
      out.push({
        part_id,
        part_code,
        qty,
        note: String(x?.note || "").trim() || null,
      });
    }
    return out;
  }

  function buildInspectionWorkOrderNotes({
    inspectionId,
    inspection_date,
    asset_code,
    asset_name,
    checklist,
    required_parts,
    notes,
  }) {
    const lines = [
      `Manager inspection #${inspectionId} (${inspection_date})`,
      `Asset: ${asset_code || ""} — ${asset_name || ""}`,
      "",
    ];
    if (checklist.length) {
      lines.push("Checklist:");
      for (const c of checklist) {
        const st = c.ok === true ? "OK" : c.ok === false ? "FAIL" : "N/A";
        lines.push(`- ${c.label}: ${st}${c.note ? ` (${c.note})` : ""}`);
      }
      lines.push("");
    }
    if (required_parts.length) {
      lines.push("Required parts:");
      for (const p of required_parts) {
        lines.push(`- ${p.part_code} × ${p.qty}${p.note ? ` — ${p.note}` : ""}`);
      }
      lines.push("");
    }
    if (notes) {
      lines.push("Inspector notes:");
      lines.push(notes);
    }
    return lines.join("\n").trim();
  }

  // =====================================================
  // TYRE INSPECTIONS + LIFECYCLE
  // =====================================================
  function ensureTyreLifecycleSchema() {
    db.prepare(`
      CREATE TABLE IF NOT EXISTS tyre_installs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        asset_id INTEGER NOT NULL,
        site_code TEXT NOT NULL DEFAULT 'main',
        position_key TEXT NOT NULL,
        position_label TEXT,
        serial_number TEXT,
        install_date TEXT NOT NULL,
        install_running_hours REAL NOT NULL DEFAULT 0,
        tyre_cost REAL NOT NULL DEFAULT 0,
        removed_date TEXT,
        removed_running_hours REAL,
        removed_reason TEXT,
        status TEXT NOT NULL DEFAULT 'active',
        notes TEXT,
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        updated_at TEXT NOT NULL DEFAULT (datetime('now')),
        FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
      )
    `).run();
    db.prepare(`
      CREATE INDEX IF NOT EXISTS idx_tyre_installs_asset_pos
      ON tyre_installs(asset_id, position_key, status)
    `).run();
  }

  function getTyreThresholds() {
    const warn = Number(getReportSetting(db, "tyre_warn_tread_mm", "8"));
    const min = Number(getReportSetting(db, "tyre_min_tread_mm", "3"));
    return {
      warn_tread_mm: Number.isFinite(warn) && warn > 0 ? warn : 8,
      min_tread_mm: Number.isFinite(min) && min > 0 ? min : 3,
    };
  }

  function lookupTyreRunningHoursNearDate(assetId, dateYmd) {
    if (!assetId || !isDate(String(dateYmd || ""))) return null;
    const row = db.prepare(`
      SELECT running_hours
      FROM tyre_inspections
      WHERE asset_id = ? AND inspection_date <= ?
      ORDER BY inspection_date DESC, id DESC
      LIMIT 1
    `).get(Number(assetId), String(dateYmd).trim());
    const hours = Number(row?.running_hours);
    return Number.isFinite(hours) && hours >= 0 ? hours : null;
  }

  const TYRE_SURVEY_POSITIONS = [
    { key: "front_left", code: "LF", order: 1 },
    { key: "front_right", code: "RF", order: 2 },
    { key: "rear_right_inner", code: "RM", order: 3 },
    { key: "rear_right_outer", code: "RR", order: 4 },
    { key: "rear_left_outer", code: "LR", order: 5 },
    { key: "rear_left_inner", code: "LM", order: 6 },
  ];

  function tyreSurveyPositionCode(positionKey) {
    const key = String(positionKey || "").trim().toLowerCase();
    const hit = TYRE_SURVEY_POSITIONS.find((p) => p.key === key);
    if (hit) return hit.code;
    const fromRow = String(key || "").trim().toUpperCase();
    return fromRow || "-";
  }

  function tyreEffectiveTreadDepth(tyreRow) {
    const outer = tyreRow?.rtd_outer ?? tyreRow?.tread_depth;
    const tread = outer == null ? null : Number(outer);
    return Number.isFinite(tread) ? tread : null;
  }

  function tyreOptionalNumber(raw) {
    if (raw === "" || raw == null) return null;
    const n = Number(raw);
    return Number.isFinite(n) ? n : null;
  }

  function formatTyreSurveyDescription(tyreRow) {
    const desc = String(tyreRow?.tyre_description || "").trim();
    if (desc) return desc;
    return [tyreRow?.tyre_make, tyreRow?.brand_number].map((x) => String(x || "").trim()).filter(Boolean).join(" ");
  }

  function monthBoundsYmd(month) {
    const m = String(month || "").trim();
    if (!isMonth(m)) return null;
    const [y, mo] = m.split("-").map(Number);
    const lastDay = new Date(y, mo, 0).getDate();
    return {
      start: `${m}-01`,
      end: `${m}-${String(lastDay).padStart(2, "0")}`,
    };
  }

  function formatSurveyAuditDate(ymd) {
    if (!isDate(String(ymd || ""))) return String(ymd || "");
    const d = new Date(`${String(ymd).trim()}T12:00:00`);
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return `${String(d.getDate()).padStart(2, "0")}-${months[d.getMonth()]}-${String(d.getFullYear()).slice(-2)}`;
  }

  function buildTyreSurveyPositionRow(surveyPos, tyreRow, installRow, runningHours) {
    const otd = tyreOptionalNumber(tyreRow?.original_tread_depth);
    const rtdOuter = tyreOptionalNumber(tyreRow?.rtd_outer ?? tyreRow?.tread_depth);
    const rtdInner = tyreOptionalNumber(tyreRow?.rtd_inner);
    const installHours = tyreOptionalNumber(installRow?.install_running_hours ?? tyreRow?.install_running_hours);
    const running = Number(runningHours || 0);
    const hoursOnTyre = installHours == null
      ? tyreOptionalNumber(tyreRow?.hours_on_tyre)
      : Math.max(0, running - installHours);
    const purchasePrice = Number(tyreRow?.tyre_cost || installRow?.tyre_cost || 0);
    let tdUsed = null;
    if (otd != null && rtdOuter != null && otd >= rtdOuter) tdUsed = Number((otd - rtdOuter).toFixed(2));
    const tdPctUsed = tdUsed != null && otd > 0 ? Number(((tdUsed / otd) * 100).toFixed(1)) : null;
    const rtdPctLeft = rtdOuter != null && otd > 0 ? Number(((rtdOuter / otd) * 100).toFixed(1)) : null;
    const hourPerMm = tdUsed > 0 && hoursOnTyre != null ? Number((hoursOnTyre / tdUsed).toFixed(1)) : null;
    let costPerHour = tyreOptionalNumber(tyreRow?.cost_per_hour);
    if (costPerHour == null && hoursOnTyre > 0 && purchasePrice > 0) {
      costPerHour = Number((purchasePrice / hoursOnTyre).toFixed(4));
    }
    return {
      position: surveyPos.code,
      serial_number: tyreRow?.serial_number || installRow?.serial_number || null,
      brand_number: tyreRow?.brand_number || null,
      tyre_make: tyreRow?.tyre_make || null,
      tyre_description: formatTyreSurveyDescription(tyreRow),
      pressure_recommended: tyreOptionalNumber(tyreRow?.pressure_recommended),
      pressure_cold: tyreOptionalNumber(tyreRow?.pressure),
      pressure_hot: tyreOptionalNumber(tyreRow?.pressure_hot),
      purchase_price: purchasePrice > 0 ? purchasePrice : null,
      hours_fitted: installHours,
      otd,
      rtd_outer: rtdOuter,
      rtd_inner: rtdInner,
      td_used: tdUsed,
      td_pct_used: tdPctUsed,
      rtd_pct_left: rtdPctLeft,
      hour_per_mm: hourPerMm,
      cost_per_hour: costPerHour,
    };
  }

  function buildTyreMonthlySurveyData(month, siteCode) {
    ensureTyreLifecycleSchema();
    const bounds = monthBoundsYmd(month);
    if (!bounds) throw new Error("month must be YYYY-MM");

    const site_code = String(siteCode || "main").trim().toLowerCase() || "main";
    const branding = getPdfReportBranding(db);
    const project = String(getReportSetting(db, "tyre_survey_project", branding.site_code || site_code)).trim().toUpperCase();
    const country = String(getReportSetting(db, "tyre_survey_country", "Mozambique")).trim() || "Mozambique";

    const rows = db.prepare(`
      SELECT
        ti.id,
        ti.asset_id,
        ti.inspection_date,
        ti.running_hours,
        ti.tyres_json,
        a.asset_code,
        a.asset_name,
        a.category
      FROM tyre_inspections ti
      JOIN assets a ON a.id = ti.asset_id
      WHERE LOWER(TRIM(COALESCE(ti.site_code, 'main'))) = LOWER(TRIM(?))
        AND ti.inspection_date >= ?
        AND ti.inspection_date <= ?
      ORDER BY a.asset_code ASC, ti.inspection_date DESC, ti.id DESC
    `).all(site_code, bounds.start, bounds.end);

    const latestByAsset = new Map();
    for (const row of rows) {
      if (!latestByAsset.has(row.asset_id)) latestByAsset.set(row.asset_id, row);
    }

    const installStmt = db.prepare(`
      SELECT position_key, serial_number, install_running_hours, tyre_cost
      FROM tyre_installs
      WHERE asset_id = ?
        AND LOWER(TRIM(COALESCE(site_code, 'main'))) = LOWER(TRIM(?))
        AND LOWER(TRIM(COALESCE(status, 'active'))) = 'active'
    `);

    const machines = [];
    for (const insp of latestByAsset.values()) {
      let tyresRaw = [];
      try {
        tyresRaw = normalizeTyreRows(JSON.parse(String(insp.tyres_json || "[]")));
      } catch {}
      const tyreByKey = new Map(tyresRaw.map((t) => [String(t.position_key || "").toLowerCase(), t]));
      const installByKey = new Map(
        installStmt.all(Number(insp.asset_id), site_code).map((i) => [String(i.position_key || "").toLowerCase(), i]),
      );
      const positions = TYRE_SURVEY_POSITIONS.map((surveyPos) => {
        const tyreRow = tyreByKey.get(surveyPos.key) || {};
        const installRow = installByKey.get(surveyPos.key) || null;
        return buildTyreSurveyPositionRow(surveyPos, tyreRow, installRow, insp.running_hours);
      });
      machines.push({
        audit_date: insp.inspection_date,
        audit_date_label: formatSurveyAuditDate(insp.inspection_date),
        project,
        country,
        machine_number: insp.asset_code,
        machine_name: insp.asset_name,
        insp_hrs: Number(Number(insp.running_hours || 0).toFixed(1)),
        machine_type: String(insp.category || "").trim(),
        positions,
      });
    }

    return {
      month,
      period: bounds,
      project,
      country,
      site_code,
      branding,
      machines,
    };
  }

  function evaluateTyreTreadStatus(treadDepth, thresholds) {
    const tread = treadDepth == null ? null : Number(treadDepth);
    if (!Number.isFinite(tread)) {
      return { lifecycle_status: "unknown", tread_alert: null };
    }
    if (tread <= thresholds.min_tread_mm) {
      return {
        lifecycle_status: "replace",
        tread_alert: `Tread ${tread.toFixed(1)} mm — at or below minimum (${thresholds.min_tread_mm} mm)`,
      };
    }
    if (tread <= thresholds.warn_tread_mm) {
      return {
        lifecycle_status: "warn",
        tread_alert: `Tread ${tread.toFixed(1)} mm — approaching end of life (warn ${thresholds.warn_tread_mm} mm)`,
      };
    }
    return { lifecycle_status: "ok", tread_alert: null };
  }

  function getActiveTyreInstall(assetId, positionKey, siteCode) {
    return db.prepare(`
      SELECT *
      FROM tyre_installs
      WHERE asset_id = ?
        AND LOWER(TRIM(position_key)) = LOWER(TRIM(?))
        AND LOWER(TRIM(COALESCE(site_code, 'main'))) = LOWER(TRIM(?))
        AND LOWER(TRIM(COALESCE(status, 'active'))) = 'active'
      ORDER BY id DESC
      LIMIT 1
    `).get(Number(assetId), String(positionKey || ""), String(siteCode || "main"));
  }

  function closeTyreInstall(installId, { removed_date, removed_running_hours, removed_reason }) {
    db.prepare(`
      UPDATE tyre_installs
      SET status = 'removed',
          removed_date = ?,
          removed_running_hours = ?,
          removed_reason = ?,
          updated_at = datetime('now')
      WHERE id = ?
    `).run(
      removed_date || null,
      removed_running_hours == null ? null : Number(removed_running_hours),
      removed_reason || null,
      Number(installId),
    );
  }

  function createTyreInstall({
    asset_id,
    site_code,
    position_key,
    position_label,
    serial_number,
    install_date,
    install_running_hours,
    tyre_cost,
    notes,
  }) {
    const ins = db.prepare(`
      INSERT INTO tyre_installs (
        asset_id, site_code, position_key, position_label, serial_number,
        install_date, install_running_hours, tyre_cost, status, notes, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'active', ?, datetime('now'))
    `).run(
      Number(asset_id),
      String(site_code || "main"),
      String(position_key || ""),
      String(position_label || position_key || ""),
      serial_number || null,
      install_date,
      Math.max(0, Number(install_running_hours || 0)),
      Math.max(0, Number(tyre_cost || 0)),
      notes || null,
    );
    return Number(ins.lastInsertRowid);
  }

  function tyreChangeDetected(activeInstall, tyreRow, inspection_date) {
    const serial = String(tyreRow?.serial_number || "").trim();
    const activeSerial = String(activeInstall?.serial_number || "").trim();
    const changedDate = isDate(String(tyreRow?.last_changed_date || "").trim())
      ? String(tyreRow.last_changed_date).trim()
      : null;
    const installDate = String(activeInstall?.install_date || "").trim();
    if (!activeInstall) return { changed: true, reason: "new_position" };
    if (serial && activeSerial && serial.toLowerCase() !== activeSerial.toLowerCase()) {
      return { changed: true, reason: "serial_change" };
    }
    if (changedDate && installDate && changedDate > installDate) {
      return { changed: true, reason: "change_date" };
    }
    if (changedDate && !installDate) {
      return { changed: true, reason: "change_date" };
    }
    const costNow = Number(tyreRow?.tyre_cost || 0);
    const costWas = Number(activeInstall?.tyre_cost || 0);
    if (changedDate && changedDate === inspection_date && costNow > 0 && costNow !== costWas) {
      return { changed: true, reason: "cost_and_date" };
    }
    return { changed: false, reason: null };
  }

  function enrichTyreLifecycleRow(tyreRow, {
    asset_id,
    site_code,
    inspection_date,
    running_hours,
    thresholds,
    persist = true,
  }) {
    const position_key = String(tyreRow?.position_key || "").trim().toLowerCase();
    const active = getActiveTyreInstall(asset_id, position_key, site_code);
    const change = tyreChangeDetected(active, tyreRow, inspection_date);
    let installId = Number(active?.id || 0);
    let install_date = String(active?.install_date || "");
    let install_running_hours = Number(active?.install_running_hours || 0);
    let tyre_cost = Number(active?.tyre_cost || 0);

    if (change.changed) {
      if (active) {
        closeTyreInstall(active.id, {
          removed_date: inspection_date,
          removed_running_hours: running_hours,
          removed_reason: change.reason,
        });
      }
      const changedDate = isDate(String(tyreRow?.last_changed_date || "").trim())
        ? String(tyreRow.last_changed_date).trim()
        : inspection_date;
      install_running_hours = lookupTyreRunningHoursNearDate(asset_id, changedDate);
      if (install_running_hours == null) install_running_hours = running_hours;
      install_date = changedDate;
      tyre_cost = Number(tyreRow?.tyre_cost || 0);
      if (persist) {
        installId = createTyreInstall({
          asset_id,
          site_code,
          position_key,
          position_label: tyreRow?.position_label,
          serial_number: tyreRow?.serial_number,
          install_date,
          install_running_hours,
          tyre_cost,
          notes: change.reason,
        });
      } else {
        installId = 0;
      }
    } else if (active) {
      installId = Number(active.id);
      install_date = String(active.install_date || "");
      install_running_hours = Number(active.install_running_hours || 0);
      tyre_cost = Number(tyreRow?.tyre_cost || 0) > 0
        ? Number(tyreRow.tyre_cost)
        : Number(active.tyre_cost || 0);
      if (persist && tyre_cost !== Number(active.tyre_cost || 0)) {
        db.prepare(`
          UPDATE tyre_installs
          SET tyre_cost = ?, serial_number = COALESCE(?, serial_number), updated_at = datetime('now')
          WHERE id = ?
        `).run(
          tyre_cost,
          String(tyreRow?.serial_number || "").trim() || null,
          installId,
        );
      }
    } else if (persist) {
      install_running_hours = running_hours;
      install_date = inspection_date;
      tyre_cost = Number(tyreRow?.tyre_cost || 0);
      installId = createTyreInstall({
        asset_id,
        site_code,
        position_key,
        position_label: tyreRow?.position_label,
        serial_number: tyreRow?.serial_number,
        install_date,
        install_running_hours,
        tyre_cost,
        notes: "initial_capture",
      });
    }

    const hours_on_tyre = Math.max(0, Number(running_hours || 0) - Number(install_running_hours || 0));
    const cost_per_hour = hours_on_tyre > 0 && tyre_cost > 0
      ? Number((tyre_cost / hours_on_tyre).toFixed(4))
      : null;
    const treadEval = evaluateTyreTreadStatus(tyreEffectiveTreadDepth(tyreRow), thresholds);

    return {
      ...tyreRow,
      install_id: installId || null,
      install_date: install_date || null,
      install_running_hours: Number(install_running_hours.toFixed(1)),
      hours_on_tyre: Number(hours_on_tyre.toFixed(1)),
      cost_per_hour,
      lifecycle_status: treadEval.lifecycle_status,
      tread_alert: treadEval.tread_alert,
      change_detected: change.changed,
      change_reason: change.reason,
    };
  }

  function buildTyreLifecycleSummary(assetId, siteCode) {
    ensureTyreLifecycleSchema();
    const thresholds = getTyreThresholds();
    const asset = db.prepare(`
      SELECT id, asset_code, asset_name
      FROM assets
      WHERE id = ?
      LIMIT 1
    `).get(Number(assetId));
    if (!asset) return null;

    const latest = db.prepare(`
      SELECT id, inspection_date, running_hours, tyres_json
      FROM tyre_inspections
      WHERE asset_id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = LOWER(TRIM(?))
      ORDER BY inspection_date DESC, id DESC
      LIMIT 1
    `).get(Number(assetId), String(siteCode || "main"));

    const latestRunning = Number(latest?.running_hours || 0);
    const latestMap = new Map();
    if (latest) {
      try {
        for (const row of normalizeTyreRows(JSON.parse(String(latest.tyres_json || "[]")))) {
          latestMap.set(String(row.position_key || "").toLowerCase(), row);
        }
      } catch {}
    }

    const activeInstalls = db.prepare(`
      SELECT *
      FROM tyre_installs
      WHERE asset_id = ?
        AND LOWER(TRIM(COALESCE(site_code, 'main'))) = LOWER(TRIM(?))
        AND LOWER(TRIM(COALESCE(status, 'active'))) = 'active'
      ORDER BY position_key ASC, id DESC
    `).all(Number(assetId), String(siteCode || "main"));

    const positions = activeInstalls.map((inst) => {
      const key = String(inst.position_key || "").toLowerCase();
      const latestRow = latestMap.get(key) || {};
      const hours_on_tyre = Math.max(0, latestRunning - Number(inst.install_running_hours || 0));
      const tyre_cost = Number(inst.tyre_cost || 0);
      const cost_per_hour = hours_on_tyre > 0 && tyre_cost > 0
        ? Number((tyre_cost / hours_on_tyre).toFixed(4))
        : null;
      const treadEval = evaluateTyreTreadStatus(tyreEffectiveTreadDepth(latestRow), thresholds);
      return {
        install_id: Number(inst.id),
        position_key: inst.position_key,
        position_label: inst.position_label || inst.position_key,
        serial_number: inst.serial_number || latestRow.serial_number || null,
        install_date: inst.install_date,
        install_running_hours: Number(Number(inst.install_running_hours || 0).toFixed(1)),
        tyre_cost,
        hours_on_tyre: Number(hours_on_tyre.toFixed(1)),
        cost_per_hour,
        pressure: latestRow.pressure ?? null,
        tread_depth: latestRow.tread_depth ?? latestRow.rtd_outer ?? null,
        rtd_outer: latestRow.rtd_outer ?? latestRow.tread_depth ?? null,
        rtd_inner: latestRow.rtd_inner ?? null,
        tyre_make: latestRow.tyre_make ?? null,
        lifecycle_status: treadEval.lifecycle_status,
        tread_alert: treadEval.tread_alert,
        last_inspection_date: latest?.inspection_date || null,
      };
    });

    const removed = db.prepare(`
      SELECT
        id, position_key, position_label, serial_number,
        install_date, install_running_hours, removed_date, removed_running_hours,
        tyre_cost, removed_reason, notes
      FROM tyre_installs
      WHERE asset_id = ?
        AND LOWER(TRIM(COALESCE(site_code, 'main'))) = LOWER(TRIM(?))
        AND LOWER(TRIM(COALESCE(status, 'active'))) = 'removed'
      ORDER BY COALESCE(removed_date, install_date) DESC, id DESC
      LIMIT 100
    `).all(Number(assetId), String(siteCode || "main")).map((row) => {
      const hours = Math.max(
        0,
        Number(row.removed_running_hours ?? 0) - Number(row.install_running_hours ?? 0),
      );
      const cost = Number(row.tyre_cost || 0);
      return {
        ...row,
        hours_on_tyre: Number(hours.toFixed(1)),
        cost_per_hour: hours > 0 && cost > 0 ? Number((cost / hours).toFixed(4)) : null,
      };
    });

    const withCost = positions.filter((p) => p.cost_per_hour != null);
    const fleet_cost_per_hour = withCost.length
      ? Number(withCost.reduce((sum, p) => sum + Number(p.cost_per_hour || 0), 0).toFixed(4))
      : null;
    const alerts = positions.filter((p) => p.lifecycle_status === "warn" || p.lifecycle_status === "replace");

    return {
      asset,
      thresholds,
      latest_inspection: latest
        ? {
            id: Number(latest.id),
            inspection_date: latest.inspection_date,
            running_hours: Number(Number(latest.running_hours || 0).toFixed(1)),
          }
        : null,
      positions,
      change_history: removed,
      summary: {
        active_tyres: positions.length,
        warn_count: positions.filter((p) => p.lifecycle_status === "warn").length,
        replace_count: positions.filter((p) => p.lifecycle_status === "replace").length,
        fleet_cost_per_hour,
        avg_cost_per_hour: withCost.length
          ? Number((withCost.reduce((s, p) => s + Number(p.cost_per_hour || 0), 0) / withCost.length).toFixed(4))
          : null,
      },
      alerts,
    };
  }

  function normalizeTyreRows(raw) {
    if (!Array.isArray(raw)) return [];
    return raw
      .map((x) => {
        const position_key = String(x?.position_key || "").trim().toLowerCase();
        const position_label = String(x?.position_label || "").trim();
        const pressure = tyreOptionalNumber(x?.pressure);
        const pressure_recommended = tyreOptionalNumber(x?.pressure_recommended);
        const pressure_hot = tyreOptionalNumber(x?.pressure_hot);
        const rtd_outer = tyreOptionalNumber(x?.rtd_outer ?? x?.tread_depth);
        const rtd_inner = tyreOptionalNumber(x?.rtd_inner);
        const original_tread_depth = tyreOptionalNumber(x?.original_tread_depth);
        const tread_depth = rtd_outer;
        const tyre_cost = tyreOptionalNumber(x?.tyre_cost) ?? 0;
        const row = {
          position_key,
          position_label: position_label || position_key,
          survey_code: String(x?.survey_code || tyreSurveyPositionCode(position_key)).trim().toUpperCase(),
          tyre_make: String(x?.tyre_make || "").trim() || null,
          brand_number: String(x?.brand_number || "").trim() || null,
          tyre_description: String(x?.tyre_description || "").trim() || null,
          pressure,
          pressure_recommended,
          pressure_hot,
          original_tread_depth,
          rtd_outer,
          rtd_inner,
          tread_depth,
          serial_number: String(x?.serial_number || "").trim() || null,
          last_changed_date: isDate(String(x?.last_changed_date || "").trim()) ? String(x.last_changed_date).trim() : null,
          tyre_cost: Number.isFinite(tyre_cost) && tyre_cost > 0 ? tyre_cost : 0,
        };
        if (x?.install_id != null) row.install_id = Number(x.install_id);
        if (x?.install_date) row.install_date = String(x.install_date);
        if (x?.install_running_hours != null) row.install_running_hours = Number(x.install_running_hours);
        if (x?.hours_on_tyre != null) row.hours_on_tyre = Number(x.hours_on_tyre);
        if (x?.cost_per_hour != null) row.cost_per_hour = Number(x.cost_per_hour);
        if (x?.lifecycle_status) row.lifecycle_status = String(x.lifecycle_status);
        if (x?.tread_alert) row.tread_alert = String(x.tread_alert);
        if (x?.change_detected != null) row.change_detected = Boolean(x.change_detected);
        if (x?.change_reason) row.change_reason = String(x.change_reason);
        return row;
      })
      .filter((x) => x.position_key);
  }

  const TYRE_SURVEY_TABLE_HEADERS = [
    "Position",
    "Serial Number",
    "Brand Number",
    "Tyre Description",
    "Pressure Recom",
    "Pressure Cold",
    "Pressure Hot",
    "Purchase Price",
    "Hours tyre was fitted",
    "OTD (mm)",
    "RTD Outer (mm)",
    "RTD Inner (mm)",
    "TD Used (mm)",
    "TD % Used",
    "RTD % Left",
    "Hour per (mm)",
    "Cost ($) per Hour",
  ];

  function surveyCellValue(v) {
    if (v == null || v === "") return "";
    return v;
  }

  function appendTyreSurveyMachineBlock(ws, machine, startRow) {
    const yellowFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
    const blueFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFADD8E6" } };
    const headerFont = { bold: true };
    let r = startRow;

    const metaRow1 = ws.getRow(r++);
    metaRow1.getCell(1).value = "Date of Audit";
    metaRow1.getCell(1).font = headerFont;
    metaRow1.getCell(2).value = machine.audit_date_label || machine.audit_date || "";
    metaRow1.getCell(3).value = "Project";
    metaRow1.getCell(3).font = headerFont;
    metaRow1.getCell(4).value = machine.project || "";
    metaRow1.getCell(5).value = "Country";
    metaRow1.getCell(5).font = headerFont;
    metaRow1.getCell(6).value = machine.country || "";

    const metaRow2 = ws.getRow(r++);
    metaRow2.getCell(1).value = "Machine Number";
    metaRow2.getCell(1).font = headerFont;
    metaRow2.getCell(2).value = machine.machine_number || "";
    metaRow2.getCell(3).value = "Insp Hrs";
    metaRow2.getCell(3).font = headerFont;
    metaRow2.getCell(4).value = machine.insp_hrs ?? "";
    metaRow2.getCell(5).value = "Machine Type";
    metaRow2.getCell(5).font = headerFont;
    metaRow2.getCell(6).value = machine.machine_type || "";

    const yellowRow = ws.getRow(r++);
    for (let c = 1; c <= TYRE_SURVEY_TABLE_HEADERS.length; c += 1) {
      yellowRow.getCell(c).fill = yellowFill;
    }

    const tableHeaderRow = ws.getRow(r++);
    TYRE_SURVEY_TABLE_HEADERS.forEach((label, idx) => {
      const cell = tableHeaderRow.getCell(idx + 1);
      cell.value = label;
      cell.font = headerFont;
      cell.fill = blueFill;
    });

    for (const pos of machine.positions || []) {
      const dataRow = ws.getRow(r++);
      const values = [
        pos.position,
        pos.serial_number,
        pos.brand_number,
        pos.tyre_description,
        pos.pressure_recommended,
        pos.pressure_cold,
        pos.pressure_hot,
        pos.purchase_price,
        pos.hours_fitted,
        pos.otd,
        pos.rtd_outer,
        pos.rtd_inner,
        pos.td_used,
        pos.td_pct_used,
        pos.rtd_pct_left,
        pos.hour_per_mm,
        pos.cost_per_hour,
      ];
      values.forEach((val, idx) => {
        dataRow.getCell(idx + 1).value = surveyCellValue(val);
      });
    }

    return r + 1;
  }

  async function buildTyreSurveyXlsxBuffer(data) {
    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();
    const ws = wb.addWorksheet("Monthly Survey");
    ws.columns = TYRE_SURVEY_TABLE_HEADERS.map((h) => ({ width: Math.max(12, Math.min(28, h.length + 2)) }));

    ws.getCell(1, 1).value = "Monthly Tyre Survey Report";
    ws.getCell(1, 1).font = { bold: true, size: 14 };
    ws.getCell(2, 1).value = `Period: ${data.month}`;
    ws.getCell(3, 1).value = `Project: ${data.project || ""}   Country: ${data.country || ""}`;

    let row = 5;
    if (!data.machines?.length) {
      ws.getCell(row, 1).value = "No tyre inspections found for this month.";
    } else {
      for (const machine of data.machines) {
        row = appendTyreSurveyMachineBlock(ws, machine, row);
      }
    }

    return wb.xlsx.writeBuffer();
  }

  function drawTyreSurveyPdfMachine(doc, machine, siteName) {
    const bottomY = pdfBodyBottom(doc);
    const needed = 140 + (machine.positions?.length || 0) * 14;
    if (doc.y + needed > bottomY) {
      doc.addPage({ layout: "landscape" });
      doc.y = pdfBodyTop(doc, { siteName });
    }

    doc.font("Helvetica-Bold").fontSize(10).fillColor("#0f172a");
    doc.text(
      `Date of Audit: ${machine.audit_date_label || machine.audit_date || "-"}   Project: ${machine.project || "-"}   Country: ${machine.country || "-"}`,
    );
    doc.moveDown(0.2);
    doc.text(
      `Machine: ${machine.machine_number || "-"}   Insp Hrs: ${machine.insp_hrs ?? "-"}   Type: ${machine.machine_type || "-"}`,
    );
    doc.moveDown(0.35);

    const cols = [
      { key: "position", label: "Pos", width: 28 },
      { key: "serial_number", label: "Serial", width: 52 },
      { key: "brand_number", label: "Brand", width: 40 },
      { key: "tyre_description", label: "Description", width: 72 },
      { key: "pressure_cold", label: "Cold", width: 32 },
      { key: "rtd_outer", label: "RTD out", width: 36 },
      { key: "rtd_inner", label: "RTD in", width: 36 },
      { key: "td_used", label: "TD used", width: 36 },
      { key: "td_pct_used", label: "TD %", width: 32 },
      { key: "cost_per_hour", label: "$/hr", width: 36 },
    ];
    const rows = (machine.positions || []).map((p) => ({
      position: p.position || "-",
      serial_number: p.serial_number || "-",
      brand_number: p.brand_number || "-",
      tyre_description: p.tyre_description || p.tyre_make || "-",
      pressure_cold: p.pressure_cold == null ? "-" : String(p.pressure_cold),
      rtd_outer: p.rtd_outer == null ? "-" : Number(p.rtd_outer).toFixed(1),
      rtd_inner: p.rtd_inner == null ? "-" : Number(p.rtd_inner).toFixed(1),
      td_used: p.td_used == null ? "-" : Number(p.td_used).toFixed(1),
      td_pct_used: p.td_pct_used == null ? "-" : `${Number(p.td_pct_used).toFixed(1)}%`,
      cost_per_hour: p.cost_per_hour == null ? "-" : Number(p.cost_per_hour).toFixed(4),
    }));
    table(doc, cols, rows, { fontSize: 7 });
    doc.moveDown(0.6);
  }

  // =====================================================
  // UNDERCARRIAGE INSPECTIONS
  // =====================================================
  function parseUndercarriageInspectionRow(row) {
    if (!row) return null;
    let measurements = [];
    let track_sag = {};
    let checklist = {};
    let summary = {};
    try { measurements = JSON.parse(String(row.measurements_json || "[]")); } catch {}
    try { track_sag = JSON.parse(String(row.track_sag_json || "{}")); } catch {}
    try { checklist = JSON.parse(String(row.checklist_json || "{}")); } catch {}
    try { summary = JSON.parse(String(row.summary_json || "{}")); } catch {}
    return {
      ...row,
      smu: row.smu == null ? null : Number(row.smu),
      measurements,
      track_sag,
      checklist,
      summary,
    };
  }

  function getPreviousUndercarriageInspection(assetId, siteCode, beforeDate, excludeId = 0) {
    const params = [Number(assetId), String(siteCode || "main"), String(beforeDate || "")];
    let excludeSql = "";
    if (Number(excludeId) > 0) {
      excludeSql = " AND id <> ?";
      params.push(Number(excludeId));
    }
    return db.prepare(`
      SELECT *
      FROM undercarriage_inspections
      WHERE asset_id = ?
        AND LOWER(TRIM(COALESCE(site_code, 'main'))) = LOWER(TRIM(?))
        AND inspection_date <= ?
        ${excludeSql}
      ORDER BY inspection_date DESC, id DESC
      LIMIT 1
    `).get(...params);
  }

  function buildUndercarriageMeasurementsForSave(rawMeasurements, {
    asset_id,
    site_code,
    inspection_date,
    smu,
    excludeId = 0,
  }) {
    const schema = buildUndercarriageComponentSchema();
    const profile = getUndercarriageWearProfileRow(asset_id, site_code);
    let normalized = normalizeUndercarriageMeasurements(rawMeasurements, schema);
    if (profile?.limits?.length) {
      const limitsByKey = new Map(
        profile.limits.map((r) => [String(r.key || "").toLowerCase(), r]),
      );
      normalized = normalized.map((row) => {
        const lim = limitsByKey.get(String(row.key || "").toLowerCase());
        if (!lim) return row;
        return {
          ...row,
          base: row.base ?? lim.base,
          wear_limit: row.wear_limit ?? lim.wear_limit,
        };
      });
    }
    const prev = getPreviousUndercarriageInspection(asset_id, site_code, inspection_date, excludeId);
    const prevMap = new Map();
    if (prev) {
      const prevRows = normalizeUndercarriageMeasurements(JSON.parse(String(prev.measurements_json || "[]")), schema);
      for (const p of prevRows) {
        prevMap.set(String(p.key || "").toLowerCase(), { ...p, inspection_hours: prev.smu });
      }
    }
    return normalized.map((row) => enrichUndercarriageMeasurement(row, {
      currentHours: smu,
      previousRow: prevMap.get(String(row.key || "").toLowerCase()) || null,
    }));
  }

  function parseUndercarriageWearProfileRow(row) {
    if (!row) return null;
    let limits = [];
    try { limits = JSON.parse(String(row.limits_json || "[]")); } catch {}
    const schema = buildUndercarriageComponentSchema();
    return {
      ...row,
      limits: normalizeUndercarriageWearLimits(limits, schema),
      configured_count: countConfiguredWearLimits(normalizeUndercarriageWearLimits(limits, schema)),
    };
  }

  function getUndercarriageWearProfileRow(assetId, siteCode) {
    const row = db.prepare(`
      SELECT *
      FROM undercarriage_wear_profiles
      WHERE asset_id = ?
        AND LOWER(TRIM(COALESCE(site_code, 'main'))) = LOWER(TRIM(?))
      LIMIT 1
    `).get(Number(assetId), String(siteCode || "main"));
    return parseUndercarriageWearProfileRow(row);
  }

  function saveUndercarriageWearProfile({
    asset_id,
    site_code,
    limits,
    source = "manual",
    notes = null,
    updated_by = null,
  }) {
    const schema = buildUndercarriageComponentSchema();
    const normalized = normalizeUndercarriageWearLimits(limits, schema);
    db.prepare(`
      INSERT INTO undercarriage_wear_profiles (
        asset_id, site_code, limits_json, source, notes, updated_by, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
      ON CONFLICT(asset_id, site_code) DO UPDATE SET
        limits_json = excluded.limits_json,
        source = excluded.source,
        notes = COALESCE(excluded.notes, undercarriage_wear_profiles.notes),
        updated_by = excluded.updated_by,
        updated_at = datetime('now')
    `).run(
      Number(asset_id),
      String(site_code || "main"),
      JSON.stringify(normalized),
      String(source || "manual"),
      notes,
      updated_by,
    );
    return getUndercarriageWearProfileRow(asset_id, site_code);
  }

  function importUndercarriageWearProfileFromLatest(assetId, siteCode, updatedBy = null) {
    const latest = db.prepare(`
      SELECT measurements_json
      FROM undercarriage_inspections
      WHERE asset_id = ?
        AND LOWER(TRIM(COALESCE(site_code, 'main'))) = LOWER(TRIM(?))
      ORDER BY inspection_date DESC, id DESC
      LIMIT 1
    `).get(Number(assetId), String(siteCode || "main"));
    if (!latest) return null;
    let measurements = [];
    try { measurements = JSON.parse(String(latest.measurements_json || "[]")); } catch {}
    const limits = measurements.map((m) => ({
      key: m.key,
      base: m.base,
      wear_limit: m.wear_limit,
    }));
    return saveUndercarriageWearProfile({
      asset_id: assetId,
      site_code: siteCode,
      limits,
      source: "import_latest_inspection",
      updated_by: updatedBy,
    });
  }

  function undercarriageWearArgb(bandKey) {
    const hit = UNDERCARRIAGE_WEAR_BANDS.find((b) => b.key === bandKey);
    return hit?.argb || null;
  }

  function formatUndercarriageAuditDate(ymd) {
    if (!isDate(String(ymd || ""))) return String(ymd || "");
    const d = new Date(`${String(ymd).trim()}T12:00:00`);
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return `${String(d.getDate()).padStart(2, "0")}-${months[d.getMonth()]}-${String(d.getFullYear()).slice(-2)}`;
  }

  function drawUndercarriageInspectionPdf(doc, insp, branding) {
    const siteName = branding?.site_name || "";
    doc.y = pdfBodyTop(doc, { siteName });
    doc.font("Helvetica-Bold").fontSize(11).fillColor("#0f172a");
    doc.text("Undercarriage Report");
    doc.moveDown(0.25);
    doc.font("Helvetica").fontSize(9).fillColor("#334155");
    doc.text(
      `Date: ${formatUndercarriageAuditDate(insp.inspection_date)}   Machine: ${insp.asset_code || "-"}   SMU: ${insp.smu ?? "-"}   Site: ${insp.site_name || siteName || "-"}`,
    );
    doc.text(
      `Serial: ${insp.serial_no || "-"}   Model: ${insp.model || insp.category || "-"}   Inspector: ${insp.inspector_name || "-"}   Job: ${insp.job_no || "-"}`,
    );
    if (insp.summary?.worst_component) {
      doc.moveDown(0.15);
      doc.fillColor("#b45309");
      doc.text(`Highest wear: ${insp.summary.worst_component} — ${Number(insp.summary.worst_wear_pct || 0).toFixed(1)}%`);
      doc.fillColor("#334155");
    }
    doc.moveDown(0.35);

    const cols = [
      { key: "component", label: "Component", width: 0.18 },
      { key: "side", label: "Side", width: 0.05 },
      { key: "measurement", label: "Meas", width: 0.07 },
      { key: "base", label: "Base", width: 0.07 },
      { key: "wear_limit", label: "Limit", width: 0.07 },
      { key: "wear_pct", label: "% Wear", width: 0.08 },
      { key: "life_hrs", label: "Life hrs", width: 0.08 },
      { key: "usage_pct", label: "Usage %", width: 0.08 },
    ];
    const rows = (insp.measurements || []).filter((m) => m.measurement != null || m.base != null).map((m) => ({
      component: m.label || m.key,
      side: m.side || "-",
      measurement: m.measurement == null ? "-" : Number(m.measurement).toFixed(1),
      base: m.base == null ? "-" : Number(m.base).toFixed(1),
      wear_limit: m.wear_limit == null ? "-" : Number(m.wear_limit).toFixed(1),
      wear_pct: m.wear_pct == null ? "-" : `${Number(m.wear_pct).toFixed(1)}%`,
      life_hrs: m.life_expectancy_hours == null ? "-" : String(m.life_expectancy_hours),
      usage_pct: m.wear_usage_pct == null ? "-" : `${Number(m.wear_usage_pct).toFixed(1)}%`,
    }));
    table(doc, cols, rows, { fontSize: 7 });

    const sag = insp.track_sag || {};
    const sagVals = UNDERCARRIAGE_TRACK_SAG_POINTS.map((p) => `${p}:${sag[p] ?? "-"}mm`).join("  ");
    ensurePageSpace(doc, 48);
    doc.moveDown(0.35);
    doc.font("Helvetica-Bold").fontSize(9).text("Track sag (mm)");
    doc.font("Helvetica").fontSize(8).text(sagVals || "—");

    ensurePageSpace(doc, 60);
    doc.moveDown(0.35);
    doc.font("Helvetica-Bold").fontSize(9).text("Condition checklist");
    doc.font("Helvetica").fontSize(8);
    for (const item of UNDERCARRIAGE_CHECKLIST_ITEMS) {
      const st = insp.checklist?.items?.[item.key] || {};
      doc.text(`${item.label}: LH ${st.lh ? "Yes" : "No"}   RH ${st.rh ? "Yes" : "No"}`);
    }
    doc.text(`General condition: ${insp.checklist?.general_condition || "-"}`);
    if (insp.checklist?.comments || insp.notes) {
      doc.moveDown(0.2);
      doc.text(`Comments: ${insp.checklist?.comments || insp.notes}`);
    }

    doc.moveDown(0.35);
    doc.fontSize(7).fillColor("#64748b");
    doc.text("Wear bands: Green 0–75% | Yellow 76–100% | Orange 101–120% | Red >120%");
  }

  async function buildUndercarriageXlsxBuffer(inspRows) {
    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();
    const ws = wb.addWorksheet("Undercarriage");
    const headerFont = { bold: true };
    const blueFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFADD8E6" } };
    const yellowFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };

    ws.getCell(1, 1).value = "UNDERCARRIAGE REPORT";
    ws.getCell(1, 1).font = { bold: true, size: 14 };
    ws.getCell(2, 1).value = "Colour: Green 0-75% | Yellow 76-100% | Orange 101-120% | Red >120%";

    let rowNum = 4;
    for (const insp of inspRows) {
      const meta1 = ws.getRow(rowNum++);
      meta1.getCell(1).value = "Date of Audit";
      meta1.getCell(1).font = headerFont;
      meta1.getCell(2).value = formatUndercarriageAuditDate(insp.inspection_date);
      meta1.getCell(3).value = "Machine";
      meta1.getCell(3).font = headerFont;
      meta1.getCell(4).value = insp.asset_code || "";
      meta1.getCell(5).value = "SMU";
      meta1.getCell(5).font = headerFont;
      meta1.getCell(6).value = insp.smu ?? "";

      const meta2 = ws.getRow(rowNum++);
      meta2.getCell(1).value = "Serial No.";
      meta2.getCell(1).font = headerFont;
      meta2.getCell(2).value = insp.serial_no || "";
      meta2.getCell(3).value = "Model";
      meta2.getCell(3).font = headerFont;
      meta2.getCell(4).value = insp.model || insp.category || "";
      meta2.getCell(5).value = "Site";
      meta2.getCell(5).font = headerFont;
      meta2.getCell(6).value = insp.site_name || "";

      const yellow = ws.getRow(rowNum++);
      for (let c = 1; c <= 12; c += 1) yellow.getCell(c).fill = yellowFill;

      const headers = [
        "Component", "Side", "Measurement", "Base", "Wear Limit", "% Wear",
        "Wear usage %", "Life expectancy (hrs)", "Wear rate %/hr",
      ];
      const hdr = ws.getRow(rowNum++);
      headers.forEach((label, idx) => {
        const cell = hdr.getCell(idx + 1);
        cell.value = label;
        cell.font = headerFont;
        cell.fill = blueFill;
      });

      for (const m of insp.measurements || []) {
        if (m.measurement == null && m.base == null && m.wear_limit == null) continue;
        const dataRow = ws.getRow(rowNum++);
        dataRow.getCell(1).value = m.label || m.key;
        dataRow.getCell(2).value = m.side || "";
        dataRow.getCell(3).value = m.measurement ?? "";
        dataRow.getCell(4).value = m.base ?? "";
        dataRow.getCell(5).value = m.wear_limit ?? "";
        dataRow.getCell(6).value = m.wear_pct ?? "";
        dataRow.getCell(7).value = m.wear_usage_pct ?? "";
        dataRow.getCell(8).value = m.life_expectancy_hours ?? "";
        dataRow.getCell(9).value = m.wear_rate_pct_per_hour ?? "";
        const argb = undercarriageWearArgb(m.wear_band);
        if (argb) {
          dataRow.getCell(6).fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
        }
      }
      rowNum += 1;
    }

    ws.columns = [
      { width: 22 }, { width: 8 }, { width: 12 }, { width: 10 }, { width: 12 },
      { width: 10 }, { width: 12 }, { width: 18 }, { width: 14 },
    ];
    return wb.xlsx.writeBuffer();
  }

  // =====================================================
  // LDV VEHICLE CHECKS (photos + click-to-pin damage markers)
  // =====================================================
  db.prepare(`
    CREATE TABLE IF NOT EXISTS vehicle_ldv_checks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      check_date TEXT NOT NULL,
      vehicle_registration TEXT,
      odometer_km REAL,
      inspector_name TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS vehicle_ldv_check_photos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      check_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      file_path TEXT NOT NULL,
      caption TEXT,
      markers_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (check_id) REFERENCES vehicle_ldv_checks(id) ON DELETE CASCADE
    )
  `).run();

  db.prepare(`CREATE INDEX IF NOT EXISTS idx_vehicle_ldv_checks_asset ON vehicle_ldv_checks(asset_id)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_vehicle_ldv_checks_date ON vehicle_ldv_checks(check_date)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_vehicle_ldv_photos_check ON vehicle_ldv_check_photos(check_id)`).run();
  ensureColumn("vehicle_ldv_checks", "check_mode TEXT DEFAULT 'ldv_general'", "check_mode");
  ensureColumn("vehicle_ldv_checks", "checklist_json TEXT", "checklist_json");
  ensureColumn("vehicle_ldv_checks", "smu_hours REAL", "smu_hours");

  const vehicleLdvcDir = path.join(dataRoot, "uploads", "vehicle-ldv-checks");
  fs.mkdirSync(vehicleLdvcDir, { recursive: true });

  function normalizeLdvPrestartChecklist(input) {
    const defaults = [
      { key: "brakes_ok", label: "Brakes" },
      { key: "lights_ok", label: "Lights" },
      { key: "tyres_ok", label: "Tyres" },
      { key: "oil_coolant_ok", label: "Oil/Coolant" },
      { key: "leaks_damage_ok", label: "Leaks/Damage" },
      { key: "safety_items_ok", label: "Safety Items" },
    ];
    const src = input && typeof input === "object" ? input : {};
    return defaults.map((d) => ({
      key: d.key,
      label: d.label,
      ok: src[d.key] === true,
    }));
  }

  function isLdvOdometerOutlier(km, referenceKm = null) {
    const n = Number(km);
    if (!Number.isFinite(n) || n < 0) return true;
    if (n > 500000) return true;
    const ref = Number(referenceKm);
    if (Number.isFinite(ref) && ref >= 0 && n > ref * 1.25 + 500) return true;
    return false;
  }

  function getSanitizedLdvBaselineKm(assetId, checkDate, opts = {}) {
    if (!assetId || !isDate(checkDate)) return null;
    const excludeCheckId = Number(opts.excludeCheckId || 0) || 0;

    const priorPrestart = getPriorLdvPrestartOdometerKm(assetId, checkDate);
    if (priorPrestart != null && !isLdvOdometerOutlier(priorPrestart)) return priorPrestart;

    if (hasColumn("daily_hours", "input_unit")) {
      const dailyRows = db.prepare(`
        SELECT closing_hours
        FROM daily_hours
        WHERE asset_id = ?
          AND work_date < ?
          AND LOWER(COALESCE(NULLIF(TRIM(input_unit), ''), 'hours')) = 'km'
          AND closing_hours IS NOT NULL
        ORDER BY work_date DESC, id DESC
        LIMIT 6
      `).all(assetId, checkDate);
      for (const row of dailyRows) {
        const n = Number(row.closing_hours);
        if (Number.isFinite(n) && n >= 0 && !isLdvOdometerOutlier(n, priorPrestart)) return n;
      }
    }

    const priorChecks = db.prepare(`
      SELECT odometer_km
      FROM vehicle_ldv_checks
      WHERE asset_id = ?
        AND odometer_km IS NOT NULL
        AND check_date < ?
        AND ${ldvPrestartModeSql("check_mode")}
        ${excludeCheckId > 0 ? "AND id != ?" : ""}
      ORDER BY check_date DESC, id DESC
      LIMIT 12
    `).all(...(excludeCheckId > 0 ? [assetId, checkDate, excludeCheckId] : [assetId, checkDate]));

    const sane = priorChecks
      .map((r) => Number(r.odometer_km))
      .filter((v) => Number.isFinite(v) && v >= 0 && !isLdvOdometerOutlier(v, priorPrestart));
    if (sane.length) return Math.max(...sane);

    return priorPrestart != null && !isLdvOdometerOutlier(priorPrestart) ? priorPrestart : null;
  }

  function purgeLdvPoisonedKmBaselines(assetId, workDate, trustedClosingKm) {
    if (!assetId || !isDate(workDate) || !Number.isFinite(Number(trustedClosingKm))) {
      return { daily_rows: 0, check_rows: 0 };
    }
    const trusted = Number(trustedClosingKm);
    let dailyRows = 0;
    let checkRows = 0;

    if (hasColumn("daily_hours", "input_unit")) {
      const badDaily = db.prepare(`
        SELECT id, closing_hours, notes
        FROM daily_hours
        WHERE asset_id = ?
          AND LOWER(COALESCE(NULLIF(TRIM(input_unit), ''), 'hours')) = 'km'
          AND closing_hours IS NOT NULL
          AND (
            closing_hours > 500000
            OR closing_hours > ?
          )
      `).all(assetId, trusted * 1.25 + 500);
      const updDaily = db.prepare(`
        UPDATE daily_hours
        SET closing_hours = ?, hours_run = ?,
            notes = TRIM(COALESCE(notes, '') || ' | Supervisor baseline KM purge')
        WHERE id = ?
      `);
      for (const row of badDaily) {
        const full = db
          .prepare(`SELECT opening_hours, work_date FROM daily_hours WHERE id = ?`)
          .get(row.id);
        const open = Number(full?.opening_hours);
        let newClose;
        let newRun;
        if (Number.isFinite(open) && open >= 0 && !isLdvOdometerOutlier(open, trusted)) {
          newClose = open;
          newRun = 0;
        } else if (String(full?.work_date || "") === String(workDate)) {
          newClose = trusted;
          newRun = 0;
        } else {
          newClose = Number.isFinite(open) && open >= 0 ? open : trusted;
          newRun = 0;
        }
        updDaily.run(newClose, newRun, Number(row.id));
        dailyRows += 1;
      }
    }

    const badChecks = db.prepare(`
      SELECT id, odometer_km
      FROM vehicle_ldv_checks
      WHERE asset_id = ?
        AND odometer_km IS NOT NULL
        AND (
          odometer_km > 500000
          OR odometer_km > ?
        )
        AND NOT (check_date = ? AND ${ldvPrestartModeSql("check_mode")})
    `).all(assetId, trusted * 1.25 + 500, workDate);

    const nullCheck = db.prepare(`
      UPDATE vehicle_ldv_checks
      SET odometer_km = NULL, notes = COALESCE(notes, '') || ' | Supervisor baseline KM purge', updated_at = datetime('now')
      WHERE id = ?
    `);
    for (const row of badChecks) {
      nullCheck.run(Number(row.id));
      checkRows += 1;
    }

    return { daily_rows: dailyRows, check_rows: checkRows };
  }

  function applyLdvKmToAllChecksForDate(assetId, workDate, closingKm, inspectorName, notes, checklistJson) {
    const closing = Number(closingKm);
    const checklistJsonOut = String(checklistJson || "");
    db.prepare(`
      UPDATE vehicle_ldv_checks
      SET odometer_km = ?, inspector_name = ?, notes = ?, check_mode = 'prestart',
          checklist_json = COALESCE(NULLIF(checklist_json, ''), ?), updated_at = datetime('now')
      WHERE asset_id = ? AND check_date = ?
    `).run(closing, inspectorName, notes, checklistJsonOut, assetId, workDate);
    return getLdvPrestartCheckRow(assetId, workDate);
  }

  function getLatestLdvOdometerKm(assetId, checkDate, opts = {}) {
    if (!assetId || !isDate(checkDate)) return null;
    const excludeCheckId = Number(opts.excludeCheckId || 0) || 0;

    /** Prior daily input closing (km) — preferred baseline for LDV odometer. */
    let fromDaily = null;
    if (hasColumn("daily_hours", "input_unit")) {
      const dailyRow = db.prepare(`
        SELECT closing_hours
        FROM daily_hours
        WHERE asset_id = ?
          AND work_date < ?
          AND LOWER(COALESCE(NULLIF(TRIM(input_unit), ''), 'hours')) = 'km'
          AND closing_hours IS NOT NULL
        ORDER BY work_date DESC, id DESC
        LIMIT 1
      `).get(assetId, checkDate);
      if (dailyRow?.closing_hours != null) {
        const n = Number(dailyRow.closing_hours);
        if (Number.isFinite(n) && n >= 0) fromDaily = n;
      }
    }

    /** Pre-start checks strictly before this date (never same-day — avoids a bad resubmit poisoning "previous"). */
    const priorChecks = db.prepare(`
      SELECT id, odometer_km, check_date
      FROM vehicle_ldv_checks
      WHERE asset_id = ?
        AND odometer_km IS NOT NULL
        AND check_date < ?
      ORDER BY check_date DESC, id DESC
      LIMIT 8
    `).all(assetId, checkDate);

    const priorKm = priorChecks
      .map((r) => Number(r.odometer_km))
      .filter((v) => Number.isFinite(v) && v >= 0);

    if (fromDaily != null) {
      const trustedPrestart = priorKm.filter((v) => v >= fromDaily * 0.98);
      const pool = [fromDaily, ...trustedPrestart];
      return Math.max(...pool);
    }

    if (priorKm.length) return Math.max(...priorKm);

    /** Same-day earlier submission when editing an existing check (exclude current record). */
    if (excludeCheckId > 0) {
      const sameDay = db.prepare(`
        SELECT odometer_km
        FROM vehicle_ldv_checks
        WHERE asset_id = ?
          AND check_date = ?
          AND id != ?
          AND odometer_km IS NOT NULL
        ORDER BY id DESC
        LIMIT 1
      `).get(assetId, checkDate, excludeCheckId);
      if (sameDay?.odometer_km != null) {
        const n = Number(sameDay.odometer_km);
        if (Number.isFinite(n) && n >= 0) return n;
      }
    }

    return null;
  }

  function resolveLdvOpeningKm(assetId, checkDate, odometerKm, previousOdometerKm, existingOpening) {
    if (previousOdometerKm != null && Number.isFinite(Number(previousOdometerKm))) {
      return Number(previousOdometerKm);
    }
    if (existingOpening != null && Number.isFinite(Number(existingOpening))) {
      return Number(existingOpening);
    }
    return Number(odometerKm);
  }

  function isLdvPrestartAssetCode(assetCode) {
    return /^V(0[1-9]|1[0-5])AM$/i.test(String(assetCode || "").trim());
  }

  function ldvPrestartModeSql(column = "check_mode") {
    return `LOWER(TRIM(COALESCE(${column}, 'ldv_general'))) = 'prestart'`;
  }

  function getLdvPrestartCheckRow(assetId, checkDate) {
    if (!assetId || !isDate(checkDate)) return null;
    return db.prepare(`
      SELECT id, check_date, odometer_km, inspector_name, notes, checklist_json, check_mode, updated_at
      FROM vehicle_ldv_checks
      WHERE asset_id = ?
        AND check_date = ?
        AND ${ldvPrestartModeSql("check_mode")}
      ORDER BY datetime(updated_at) DESC, id DESC
      LIMIT 1
    `).get(assetId, checkDate);
  }

  function getPriorLdvPrestartOdometerKm(assetId, checkDate) {
    if (!assetId || !isDate(checkDate)) return null;
    const row = db.prepare(`
      SELECT odometer_km
      FROM vehicle_ldv_checks
      WHERE asset_id = ?
        AND check_date < ?
        AND odometer_km IS NOT NULL
        AND ${ldvPrestartModeSql("check_mode")}
      ORDER BY check_date DESC, id DESC
      LIMIT 1
    `).get(assetId, checkDate);
    if (row?.odometer_km == null) return null;
    const n = Number(row.odometer_km);
    return Number.isFinite(n) && n >= 0 ? n : null;
  }

  function resolveLdvCorrectionOpeningKm(assetId, workDate, closingKm, excludeCheckId, explicitOpening) {
    const closing = Number(closingKm);
    if (!Number.isFinite(closing) || closing < 0) return null;

    if (explicitOpening != null && String(explicitOpening).trim() !== "") {
      const n = Number(explicitOpening);
      if (Number.isFinite(n) && n >= 0 && n <= closing && !isLdvOdometerOutlier(n, closing)) {
        return n;
      }
    }

    const sanitized = getSanitizedLdvBaselineKm(assetId, workDate, {
      excludeCheckId: Number(excludeCheckId || 0) || 0,
    });
    if (sanitized != null && closing >= sanitized) return sanitized;

    const priorPrestart = getPriorLdvPrestartOdometerKm(assetId, workDate);
    if (priorPrestart != null && !isLdvOdometerOutlier(priorPrestart, closing) && closing >= priorPrestart) {
      return priorPrestart;
    }

    if (hasColumn("daily_hours", "input_unit")) {
      const dailyRows = db.prepare(`
        SELECT closing_hours
        FROM daily_hours
        WHERE asset_id = ?
          AND work_date < ?
          AND LOWER(COALESCE(NULLIF(TRIM(input_unit), ''), 'hours')) = 'km'
          AND closing_hours IS NOT NULL
        ORDER BY work_date DESC, id DESC
        LIMIT 6
      `).all(assetId, workDate);
      for (const row of dailyRows) {
        const n = Number(row.closing_hours);
        if (Number.isFinite(n) && n >= 0 && !isLdvOdometerOutlier(n, closing) && closing >= n) return n;
      }
    }

    return closing;
  }

  function getMachinePrestartCheckRow(assetId, checkDate, mode) {
    const primary = db.prepare(`
      SELECT id, check_date, odometer_km, smu_hours, inspector_name, notes, checklist_json, check_mode
      FROM vehicle_ldv_checks
      WHERE asset_id = ?
        AND check_date = ?
        AND check_mode = ?
      ORDER BY id DESC
      LIMIT 1
    `).get(assetId, checkDate, mode);
    if (primary) return primary;
    return db.prepare(`
      SELECT id, check_date, odometer_km, smu_hours, inspector_name, notes, checklist_json, check_mode
      FROM vehicle_ldv_checks
      WHERE asset_id = ?
        AND check_date = ?
        AND check_mode LIKE 'machine_prestart_%'
      ORDER BY id DESC
      LIMIT 1
    `).get(assetId, checkDate);
  }

  function checklistRowStatus(checklistJson, normalizeFn) {
    if (!checklistJson) return { status: "pending", check_id: null };
    try {
      const parsed = JSON.parse(String(checklistJson || "{}"));
      const rows = normalizeFn(parsed);
      if (!rows.length) return { status: "pending", check_id: null };
      const allOk = rows.every((r) => r.ok === true);
      return { status: allOk ? "compliant" : "pending" };
    } catch {
      return { status: "pending" };
    }
  }

  function syncLdvPrestartToDailyHours(assetId, checkDate, odometerKm, inspectorName, previousOdometerKm, opts = {}) {
    if (!assetId || !isDate(checkDate) || !Number.isFinite(Number(odometerKm))) return { synced: false };
    const odometer = Number(odometerKm);
    const previous =
      previousOdometerKm != null && Number.isFinite(Number(previousOdometerKm))
        ? Number(previousOdometerKm)
        : null;
    const unusual =
      Boolean(opts.unusual_km) ||
      (previous != null && odometer < previous) ||
      isLdvOdometerOutlier(odometer, previous);

    if (unusual) {
      return {
        synced: false,
        skipped: true,
        reason: "unusual_km",
        work_date: checkDate,
        message:
          "Pre-start KM saved. Daily input was not updated because the reading looks unusual — a supervisor can correct it.",
      };
    }

    const existing = db.prepare(`
      SELECT id, scheduled_hours, opening_hours, closing_hours, hours_run, is_used, operator, notes
      FROM daily_hours
      WHERE asset_id = ?
        AND work_date = ?
      LIMIT 1
    `).get(assetId, checkDate);

    const opening = resolveLdvOpeningKm(
      assetId,
      checkDate,
      odometer,
      previousOdometerKm,
      existing?.opening_hours
    );
    const runDelta = Math.max(0, odometer - opening);

    if (existing?.id) {
      const nextNotes = (() => {
        const base = String(existing.notes || "").trim();
        const tag = "LDV pre-start KM captured via QR";
        if (!base) return tag;
        return base.includes(tag) ? base : `${base} | ${tag}`;
      })();
      const hasInputUnit = hasColumn("daily_hours", "input_unit");
      if (hasInputUnit) {
        db.prepare(`
          UPDATE daily_hours
          SET
            opening_hours = ?,
            closing_hours = ?,
            hours_run = ?,
            input_unit = 'km',
            operator = COALESCE(NULLIF(operator, ''), ?),
            notes = ?
          WHERE id = ?
        `).run(opening, odometer, runDelta, inspectorName || null, nextNotes, Number(existing.id));
      } else {
        db.prepare(`
          UPDATE daily_hours
          SET
            opening_hours = ?,
            closing_hours = ?,
            hours_run = ?,
            operator = COALESCE(NULLIF(operator, ''), ?),
            notes = ?
          WHERE id = ?
        `).run(opening, odometer, runDelta, inspectorName || null, nextNotes, Number(existing.id));
      }
      return {
        synced: true,
        mode: "updated",
        work_date: checkDate,
        opening_km: Number(opening.toFixed(1)),
        closing_km: Number(odometer.toFixed(1)),
        run_km: Number(runDelta.toFixed(1)),
      };
    }

    const hasInputUnit = hasColumn("daily_hours", "input_unit");
    if (hasInputUnit) {
      db.prepare(`
        INSERT INTO daily_hours (
          asset_id, work_date, scheduled_hours, opening_hours, closing_hours,
          hours_run, input_unit, is_used, operator, notes
        )
        VALUES (?, ?, 0, ?, ?, ?, 'km', 1, ?, ?)
        ON CONFLICT(asset_id, work_date) DO UPDATE SET
          opening_hours = excluded.opening_hours,
          closing_hours = excluded.closing_hours,
          hours_run = excluded.hours_run,
          input_unit = 'km',
          operator = COALESCE(NULLIF(daily_hours.operator, ''), excluded.operator),
          notes = COALESCE(NULLIF(daily_hours.notes, ''), excluded.notes)
      `).run(assetId, checkDate, opening, odometer, runDelta, inspectorName || null, "LDV pre-start KM captured via QR");
    } else {
      db.prepare(`
        INSERT INTO daily_hours (
          asset_id, work_date, scheduled_hours, opening_hours, closing_hours,
          hours_run, is_used, operator, notes
        )
        VALUES (?, ?, 0, ?, ?, ?, 1, ?, ?)
        ON CONFLICT(asset_id, work_date) DO UPDATE SET
          opening_hours = excluded.opening_hours,
          closing_hours = excluded.closing_hours,
          hours_run = excluded.hours_run,
          operator = COALESCE(NULLIF(daily_hours.operator, ''), excluded.operator),
          notes = COALESCE(NULLIF(daily_hours.notes, ''), excluded.notes)
      `).run(assetId, checkDate, opening, odometer, runDelta, inspectorName || null, "LDV pre-start KM captured via QR");
    }
    return {
      synced: true,
      mode: "inserted",
      work_date: checkDate,
      opening_km: Number(opening.toFixed(1)),
      closing_km: Number(odometer.toFixed(1)),
      run_km: Number(runDelta.toFixed(1)),
    };
  }

  function getLatestMachineSmuHours(assetId, checkDate, opts = {}) {
    if (!assetId || !isDate(checkDate)) return null;
    const excludeCheckId = Number(opts.excludeCheckId || 0) || 0;

    let fromDaily = null;
    const dailySql = hasColumn("daily_hours", "input_unit")
      ? `
        SELECT closing_hours
        FROM daily_hours
        WHERE asset_id = ?
          AND work_date < ?
          AND closing_hours IS NOT NULL
          AND LOWER(COALESCE(NULLIF(TRIM(input_unit), ''), 'hours')) != 'km'
        ORDER BY work_date DESC, id DESC
        LIMIT 1
      `
      : `
        SELECT closing_hours
        FROM daily_hours
        WHERE asset_id = ?
          AND work_date < ?
          AND closing_hours IS NOT NULL
        ORDER BY work_date DESC, id DESC
        LIMIT 1
      `;
    const dailyRow = db.prepare(dailySql).get(assetId, checkDate);
    if (dailyRow?.closing_hours != null) {
      const n = Number(dailyRow.closing_hours);
      if (Number.isFinite(n) && n >= 0) fromDaily = n;
    }

    const priorChecks = db.prepare(`
      SELECT id, smu_hours, check_date
      FROM vehicle_ldv_checks
      WHERE asset_id = ?
        AND smu_hours IS NOT NULL
        AND check_date < ?
        AND check_mode LIKE 'machine_prestart_%'
      ORDER BY check_date DESC, id DESC
      LIMIT 8
    `).all(assetId, checkDate);

    const priorSmu = priorChecks
      .map((r) => Number(r.smu_hours))
      .filter((v) => Number.isFinite(v) && v >= 0);

    if (fromDaily != null) {
      const trustedPrestart = priorSmu.filter((v) => v >= fromDaily * 0.98);
      const pool = [fromDaily, ...trustedPrestart];
      return Math.max(...pool);
    }

    if (priorSmu.length) return Math.max(...priorSmu);

    if (excludeCheckId > 0) {
      const sameDay = db.prepare(`
        SELECT smu_hours
        FROM vehicle_ldv_checks
        WHERE asset_id = ?
          AND check_date = ?
          AND id != ?
          AND smu_hours IS NOT NULL
          AND check_mode LIKE 'machine_prestart_%'
        ORDER BY id DESC
        LIMIT 1
      `).get(assetId, checkDate, excludeCheckId);
      if (sameDay?.smu_hours != null) {
        const n = Number(sameDay.smu_hours);
        if (Number.isFinite(n) && n >= 0) return n;
      }
    }

    return null;
  }

  function syncMachinePrestartToDailyHours(assetId, checkDate, smuHours, inspectorName, previousSmuHours) {
    if (!assetId || !isDate(checkDate) || !Number.isFinite(Number(smuHours))) return { synced: false };
    const closing = Number(smuHours);
    const existing = db.prepare(`
      SELECT id, scheduled_hours, opening_hours, closing_hours, hours_run, is_used, operator, notes
      FROM daily_hours
      WHERE asset_id = ?
        AND work_date = ?
      LIMIT 1
    `).get(assetId, checkDate);

    const opening = resolveLdvOpeningKm(
      assetId,
      checkDate,
      closing,
      previousSmuHours,
      existing?.opening_hours
    );
    const runDelta = Math.max(0, closing - opening);

    const nextNotes = (() => {
      const base = String(existing?.notes || "").trim();
      const tag = "Machine pre-start SMU captured via checklist";
      if (!base) return tag;
      return base.includes(tag) ? base : `${base} | ${tag}`;
    })();

    const hasInputUnit = hasColumn("daily_hours", "input_unit");
    if (existing?.id) {
      if (hasInputUnit) {
        db.prepare(`
          UPDATE daily_hours
          SET
            opening_hours = ?,
            closing_hours = ?,
            hours_run = ?,
            input_unit = 'hours',
            operator = COALESCE(NULLIF(operator, ''), ?),
            notes = ?
          WHERE id = ?
        `).run(opening, closing, runDelta, inspectorName || null, nextNotes, Number(existing.id));
      } else {
        db.prepare(`
          UPDATE daily_hours
          SET
            opening_hours = ?,
            closing_hours = ?,
            hours_run = ?,
            operator = COALESCE(NULLIF(operator, ''), ?),
            notes = ?
          WHERE id = ?
        `).run(opening, closing, runDelta, inspectorName || null, nextNotes, Number(existing.id));
      }
    } else if (hasInputUnit) {
      db.prepare(`
        INSERT INTO daily_hours (
          asset_id, work_date, scheduled_hours, opening_hours, closing_hours,
          hours_run, input_unit, is_used, operator, notes
        )
        VALUES (?, ?, 0, ?, ?, ?, 'hours', 1, ?, ?)
        ON CONFLICT(asset_id, work_date) DO UPDATE SET
          opening_hours = excluded.opening_hours,
          closing_hours = excluded.closing_hours,
          hours_run = excluded.hours_run,
          input_unit = 'hours',
          operator = COALESCE(NULLIF(daily_hours.operator, ''), excluded.operator),
          notes = COALESCE(NULLIF(daily_hours.notes, ''), excluded.notes)
      `).run(assetId, checkDate, opening, closing, runDelta, inspectorName || null, nextNotes);
    } else {
      db.prepare(`
        INSERT INTO daily_hours (
          asset_id, work_date, scheduled_hours, opening_hours, closing_hours,
          hours_run, is_used, operator, notes
        )
        VALUES (?, ?, 0, ?, ?, ?, 1, ?, ?)
        ON CONFLICT(asset_id, work_date) DO UPDATE SET
          opening_hours = excluded.opening_hours,
          closing_hours = excluded.closing_hours,
          hours_run = excluded.hours_run,
          operator = COALESCE(NULLIF(daily_hours.operator, ''), excluded.operator),
          notes = COALESCE(NULLIF(daily_hours.notes, ''), excluded.notes)
      `).run(assetId, checkDate, opening, closing, runDelta, inspectorName || null, nextNotes);
    }

    if (hasInputUnit) {
      db.prepare(`
        INSERT INTO asset_input_units (asset_id, input_unit, updated_at)
        VALUES (?, 'hours', datetime('now'))
        ON CONFLICT(asset_id) DO UPDATE SET input_unit = 'hours', updated_at = datetime('now')
      `).run(assetId);
    }

    return {
      synced: true,
      unit: "hours",
      mode: existing?.id ? "updated" : "inserted",
      work_date: checkDate,
      opening_hours: Number(opening.toFixed(1)),
      closing_hours: Number(closing.toFixed(1)),
      run_hours: Number(runDelta.toFixed(1)),
    };
  }

  function upsertMachineDailyHoursCorrection(assetId, workDate, openingHours, closingHours, inspectorName, correctionNote) {
    const runHours = Math.max(0, closingHours - openingHours);
    const hasInputUnit = hasColumn("daily_hours", "input_unit");
    const dailyNote = `Supervisor hours correction | ${correctionNote}`;
    if (hasInputUnit) {
      db.prepare(`
        INSERT INTO daily_hours (
          asset_id, work_date, scheduled_hours, opening_hours, closing_hours,
          hours_run, input_unit, is_used, operator, notes
        )
        VALUES (?, ?, 0, ?, ?, ?, 'hours', 1, ?, ?)
        ON CONFLICT(asset_id, work_date) DO UPDATE SET
          opening_hours = excluded.opening_hours,
          closing_hours = excluded.closing_hours,
          hours_run = excluded.hours_run,
          input_unit = 'hours',
          is_used = 1,
          operator = excluded.operator,
          notes = excluded.notes
      `).run(assetId, workDate, openingHours, closingHours, runHours, inspectorName, dailyNote);
      db.prepare(`
        INSERT INTO asset_input_units (asset_id, input_unit, updated_at)
        VALUES (?, 'hours', datetime('now'))
        ON CONFLICT(asset_id) DO UPDATE SET input_unit = 'hours', updated_at = datetime('now')
      `).run(assetId);
    } else {
      db.prepare(`
        INSERT INTO daily_hours (
          asset_id, work_date, scheduled_hours, opening_hours, closing_hours,
          hours_run, is_used, operator, notes
        )
        VALUES (?, ?, 0, ?, ?, ?, 1, ?, ?)
        ON CONFLICT(asset_id, work_date) DO UPDATE SET
          opening_hours = excluded.opening_hours,
          closing_hours = excluded.closing_hours,
          hours_run = excluded.hours_run,
          is_used = 1,
          operator = excluded.operator,
          notes = excluded.notes
      `).run(assetId, workDate, openingHours, closingHours, runHours, inspectorName, dailyNote);
    }
    return runHours;
  }

  const PARTS_REQUEST_STATUSES = new Set(["requested", "ordered", "received", "cancelled"]);
  const PARTS_REQUEST_URGENCY = new Set(["normal", "urgent", "critical"]);
  const PARTS_REQUEST_MANAGERS = ["admin", "supervisor", "stores", "plant_manager", "site_manager", "workshop_manager"];

  const MECHANIC_LABOR_EDITORS = [
    "admin",
    "supervisor",
    "plant_manager",
    "site_manager",
    "workshop_manager",
    "artisan",
    "stores",
  ];
  const MECHANIC_LABOR_RATE_MANAGERS = [
    "admin",
    "supervisor",
    "plant_manager",
    "site_manager",
    "workshop_manager",
  ];
  const MECHANIC_MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
  ];

  function readMechanicLaborDefaultRate() {
    try {
      const row = db.prepare(`
        SELECT value FROM cost_settings WHERE key = 'labor_cost_per_hour_default' LIMIT 1
      `).get();
      const v = Number(row?.value);
      return Number.isFinite(v) && v > 0 ? v : 35;
    } catch {
      return 35;
    }
  }

  function enrichMechanicLaborRow(row, defaultRate) {
    const hours = Number(row?.hours || 0);
    const rate = Number.isFinite(Number(row?.labor_rate_per_hour)) && Number(row.labor_rate_per_hour) > 0
      ? Number(row.labor_rate_per_hour)
      : defaultRate;
    return {
      ...row,
      hours: Number(hours.toFixed(2)),
      labor_rate_per_hour: Number(rate.toFixed(2)),
      labor_cost: Number((hours * rate).toFixed(2)),
    };
  }

  function mechanicLaborEntryBody(body = {}) {
    const work_date = String(body.work_date || "").trim();
    const technician_name = String(body.technician_name || "").trim();
    const asset_code = String(body.asset_code || "").trim().toUpperCase();
    const reason = String(body.reason || "").trim();
    const hours = Math.max(0, Number(body.hours || 0));
    const labor_rate_per_hour = body.labor_rate_per_hour != null
      ? Math.max(0, Number(body.labor_rate_per_hour))
      : null;
    const category = String(body.category || "").trim() || null;
    const time_started = String(body.time_started || "").trim() || null;
    const time_finished = String(body.time_finished || "").trim() || null;
    const job_card_no = String(body.job_card_no || "").trim() || null;
    const smrValue = body.smr;
    const smr = smrValue === "" || smrValue == null ? null : Number(smrValue);
    return {
      work_date,
      technician_name,
      asset_code,
      reason,
      hours,
      labor_rate_per_hour,
      category,
      time_started,
      time_finished,
      job_card_no,
      smr: Number.isFinite(smr) ? Number(smr.toFixed(1)) : null,
    };
  }

  function resolveMechanicLaborSmr(work_date, asset_code, site_code) {
    if (!isDate(work_date) || !asset_code) return null;
    try {
      const asset = db.prepare(`
        SELECT id
        FROM assets
        WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?))
        LIMIT 1
      `).get(asset_code);
      if (!asset?.id) return null;
      const daily = db.prepare(`
        SELECT closing_hours, opening_hours
        FROM daily_hours
        WHERE asset_id = ?
          AND work_date = ?
          AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
        ORDER BY id DESC
        LIMIT 1
      `).get(asset.id, work_date, String(site_code || "main").trim().toLowerCase() || "main");
      const closing = Number(daily?.closing_hours);
      if (Number.isFinite(closing) && closing >= 0) return Number(closing.toFixed(1));
      const opening = Number(daily?.opening_hours);
      return Number.isFinite(opening) && opening >= 0 ? Number(opening.toFixed(1)) : null;
    } catch {
      return null;
    }
  }

  function enrichMechanicLaborSmr(row, site_code) {
    if (row?.smr != null && Number.isFinite(Number(row.smr))) return row;
    return {
      ...row,
      smr: resolveMechanicLaborSmr(row?.work_date, row?.asset_code, site_code),
    };
  }

  function ensureMechanicLaborExtendedColumns() {
    for (const [col, def] of [
      ["category", "category TEXT"],
      ["time_started", "time_started TEXT"],
      ["time_finished", "time_finished TEXT"],
      ["job_card_no", "job_card_no TEXT"],
      ["smr", "smr REAL"],
    ]) {
      try {
        const rows = db.prepare(`PRAGMA table_info(mechanic_labor_entries)`).all();
        if (rows.length && !rows.some((r) => String(r.name) === col)) {
          db.prepare(`ALTER TABLE mechanic_labor_entries ADD COLUMN ${def}`).run();
        }
      } catch {}
    }
  }

  const MECHANICS_TIMESHEET_COLS = [
    { header: "Date", key: "Date", width: 18 },
    { header: "Plant no", key: "Plant no", width: 14 },
    { header: "Work Hours", key: "Work Hours", width: 11 },
    { header: "Category", key: "Category", width: 12 },
    { header: "Description Of Work Carried Out", key: "Description Of Work Carried Out", width: 42 },
    { header: "Time Started", key: "Time Started", width: 12 },
    { header: "Time finished", key: "Time finished", width: 12 },
    { header: "Technician", key: "Technician", width: 14 },
    { header: "Job Card No", key: "Job Card No", width: 12 },
    { header: "SMR", key: "SMR", width: 10 },
  ];

  function parseMechanicsTimesheetRange(req) {
    const year = String(req.query?.year || "").trim();
    let from = String(req.query?.from || "").trim();
    let to = String(req.query?.to || "").trim();
    if (/^\d{4}$/.test(year)) {
      from = `${year}-01-01`;
      to = `${year}-12-31`;
    }
    return { from, to };
  }

  function buildMechanicsTimesheetWorkbook(exportRows, meta) {
    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const wsLog = wb.addWorksheet("Mechanics log", { views: [{ state: "frozen", ySplit: 1 }] });
    wsLog.columns = MECHANICS_TIMESHEET_COLS;
    for (const r of exportRows) {
      wsLog.addRow({
        Date: r.Date,
        "Plant no": r["Plant no"],
        "Work Hours": r["Work Hours"],
        Category: r.Category,
        "Description Of Work Carried Out": r["Description Of Work Carried Out"],
        "Time Started": r["Time Started"],
        "Time finished": r["Time finished"],
        Technician: r.Technician,
        "Job Card No": r["Job Card No"],
        SMR: r.SMR != null ? r.SMR : "",
      });
    }
    wsLog.getRow(1).font = { bold: true };

    const yearMatch = String(meta.from || "").match(/^(\d{4})/);
    const year = yearMatch ? yearMatch[1] : "report";
    for (let m = 1; m <= 12; m += 1) {
      const monthKey = `${year}-${String(m).padStart(2, "0")}`;
      const monthRows = exportRows.filter((r) => String(r.work_date || "").startsWith(monthKey));
      if (!monthRows.length) continue;
      const ws = wb.addWorksheet(MECHANIC_MONTH_NAMES[m - 1], { views: [{ state: "frozen", ySplit: 1 }] });
      ws.columns = MECHANICS_TIMESHEET_COLS;
      for (const r of monthRows) {
        ws.addRow({
          Date: r.Date,
          "Plant no": r["Plant no"],
          "Work Hours": r["Work Hours"],
          Category: r.Category,
          "Description Of Work Carried Out": r["Description Of Work Carried Out"],
          "Time Started": r["Time Started"],
          "Time finished": r["Time finished"],
          Technician: r.Technician,
          "Job Card No": r["Job Card No"],
          SMR: r.SMR != null ? r.SMR : "",
        });
      }
      ws.getRow(1).font = { bold: true };
    }

    const wsInfo = wb.addWorksheet("Info");
    wsInfo.columns = [
      { header: "Field", key: "field", width: 28 },
      { header: "Value", key: "value", width: 36 },
    ];
    wsInfo.addRows([
      { field: "Period", value: `${meta.from} to ${meta.to}` },
      { field: "Days with saved entries", value: meta.days_from_saved },
      { field: "Total rows", value: meta.row_count },
      { field: "Excluded plant", value: (meta.excluded_assets || []).join(", ") },
      { field: "Source", value: "Saved mechanic labor entries only (no synthetic fill)" },
      { field: "Upload columns", value: "Date, Plant no, Work Hours, Category, Description Of Work Carried Out, Time Started, Time finished, Technician, Job Card No, SMR" },
    ]);
    wsInfo.getRow(1).font = { bold: true };

    return { wb, year };
  }

  // Route groups live in routes/maintenance/. They receive the shared helpers above through ctx.
  const ctx = {
    MECHANIC_LABOR_EDITORS,
    MECHANIC_LABOR_RATE_MANAGERS,
    MECHANIC_MONTH_NAMES,
    PARTS_REQUEST_MANAGERS,
    PARTS_REQUEST_STATUSES,
    PARTS_REQUEST_URGENCY,
    addDaysYmd,
    addInsightsCostPerMachineSheet,
    addInsightsExportSummarySheet,
    addInsightsPartsDemandSheets,
    addWeeklyInspectionSlot,
    applyLdvKmToAllChecksForDate,
    buildInsightsBreakdownLaborIncidents,
    buildInspectionWorkOrderNotes,
    buildMaintenanceInsightsDowntime,
    buildMaintenanceReliabilityReport,
    buildMechanicsTimesheetWorkbook,
    buildTyreLifecycleSummary,
    buildTyreMonthlySurveyData,
    buildTyreSurveyXlsxBuffer,
    buildUndercarriageMeasurementsForSave,
    buildUndercarriageXlsxBuffer,
    buildUpcomingServiceCostForecasts,
    buildWeeklyForumSummary,
    buildWeeklyInspectionCalendarData,
    checklistRowStatus,
    clearWeeklyInspectionRoster,
    copyWeeklyInspectionDay,
    damageReportsDir,
    displayPlanServiceName,
    dmgPhotoCaptionCol,
    dmgPhotoCreatedCol,
    dmgPhotoPathCol,
    dmgPhotoReportCol,
    drawTyreSurveyPdfMachine,
    drawUndercarriageInspectionPdf,
    drawWeeklyInspectionCalendarPdfGrid,
    enrichMechanicLaborRow,
    enrichMechanicLaborSmr,
    enrichTyreLifecycleRow,
    ensureBreakdownRepairLaborSchema,
    ensureColumn,
    ensureMechanicLaborExtendedColumns,
    ensureTyreLifecycleSchema,
    ensureWeeklyInspectionSchema,
    getAssetCurrentHours,
    getAssetCurrentHoursInfo,
    getAssetHoursInfoAsOf,
    getLatestLdvOdometerKm,
    getLatestMachineSmuHours,
    getLdvPrestartCheckRow,
    getMachinePrestartCheckRow,
    getMaintenanceRoles,
    getPriorLdvPrestartOdometerKm,
    getSanitizedLdvBaselineKm,
    getTyreThresholds,
    getUndercarriageWearProfileRow,
    hasColumn,
    hasTable,
    importUndercarriageWearProfileFromLatest,
    inspectionsDir,
    isLdvOdometerOutlier,
    isLdvPrestartAssetCode,
    isMonth,
    listMaintenancePlans,
    loadInsightsExportPayload,
    loadTemplateDetail,
    maintenanceActor,
    mechanicLaborEntryBody,
    monthBoundsYmd,
    normalizeArtisanInspectionChecklist,
    normalizeEquipCategory,
    normalizeInspectionChecklist,
    normalizeInspectionParts,
    normalizeLdvPrestartChecklist,
    normalizeTyreRows,
    parseMechanicsTimesheetRange,
    parseReliabilityAssetIds,
    parseUndercarriageInspectionRow,
    photoCaptionCol,
    photoCreatedCol,
    photoInspectionCol,
    photoPathCol,
    pickExistingColumn,
    purgeLdvPoisonedKmBaselines,
    readMechanicLaborDefaultRate,
    requireMaintenancePermission,
    requireMaintenanceRoles,
    resolveLdvCorrectionOpeningKm,
    resolveMechanicLaborSmr,
    resolveStorageAbs,
    saveUndercarriageWearProfile,
    syncLdvPrestartToDailyHours,
    syncMachinePrestartToDailyHours,
    templateSummaryRows,
    updateWeeklyInspectionSlotStatus,
    upsertMachineDailyHoursCorrection,
    vehicleLdvcDir,
    wiEnsureBodySpace,
    wiFormatMinutesPdf,
    writeTemplateItems,
  };
  registerServiceTemplatesRoutes(app, ctx);
  registerPlansRoutes(app, ctx);
  registerInsightsRoutes(app, ctx);
  registerWeeklyRoutes(app, ctx);
  registerCostingRoutes(app, ctx);
  registerInspectionsRoutes(app, ctx);
  registerPrestartChecksRoutes(app, ctx);
  registerPartsRequestsRoutes(app, ctx);
  registerMechanicLaborRoutes(app, ctx);
}
