// IRONLOG/api/routes/maintenance.routes.js
import { db } from "../db/client.js";
import multipart from "@fastify/multipart";
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import ExcelJS from "exceljs";
import { buildPdfBuffer, ensurePageSpace, pdfBodyBottom, pdfBodyTop, sectionTitle, table } from "../utils/pdfGenerator.js";
import { getPdfReportBranding, getReportSetting } from "../utils/reportSettings.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";
import {
  resolveMachinePrestartProfile,
  getMachinePrestartTemplate,
  listMachinePrestartProfiles,
  normalizeMachinePrestartChecklist,
  checklistToJsonObject,
  machinePrestartCheckMode,
} from "../utils/machinePrestartTemplates.js";
import { listDailyPrestarts, prestartDeductionForProductionFleet } from "../utils/prestartDaily.js";
import {
  buildReliabilityIncidentsForAssets,
  computeMtbfLttr,
  round2,
} from "../utils/reliabilityMetrics.js";
import { normalizeUploadedPhoto } from "../utils/imagePdf.js";
import {
  generateMechanicsTimesheet,
  mechanicsTimesheetToExportRows,
} from "../utils/mechanicsTimesheetGenerator.js";
import { parseMechanicsTimesheetUpload } from "../utils/mechanicsTimesheetImport.js";
import { resolveStorageAbs as resolveStorageAbsPath, getDataRoot } from "../utils/storagePaths.js";
import {
  buildDueListFromPlans,
  enrichPlansWithNextService,
  hasRotatingSchedule,
  resolveNextServiceForAssetPlans,
  groupActivePlansByAsset,
  snapLastServiceHours,
  planIntervalHours,
  resolveLegacyPlanDue,
  classifyServiceDue,
  classifyServiceDueStatus,
  meterUnitForAsset,
} from "../utils/serviceSchedule.js";
import { enrichDueRowsWithEstimates } from "../utils/maintenanceEstimates.js";
import {
  buildUndercarriageComponentSchema,
  enrichUndercarriageMeasurement,
  normalizeUndercarriageChecklist,
  normalizeUndercarriageMeasurements,
  normalizeUndercarriageTrackSag,
  normalizeUndercarriageWearLimits,
  applyWearLimitsToMeasurements,
  countConfiguredWearLimits,
  summarizeUndercarriageInspection,
  undercarriageWearBand,
  UNDERCARRIAGE_CHECKLIST_ITEMS,
  UNDERCARRIAGE_TRACK_SAG_POINTS,
  UNDERCARRIAGE_WEAR_BANDS,
} from "../utils/undercarriageTemplate.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

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

/**
 * SQL boolean expression (SQLite): parts row is bucketed as oil/lubricant, not hard parts.
 * Uses parts.consumable_kind when set; otherwise part name/code heuristics.
 */
function sqlOilPartPredicate(alias = "p") {
  const p = alias;
  return `(
  LOWER(TRIM(COALESCE(${p}.consumable_kind, ''))) IN ('oil', 'lube', 'lubricant', 'hydraulic', 'hydraulic_oil', 'coolant', 'grease', 'hyd fluid', 'hydraulic fluid')
  OR (
    TRIM(COALESCE(${p}.consumable_kind, '')) = ''
    AND (
      INSTR(LOWER(' ' || REPLACE(REPLACE(COALESCE(${p}.part_name, ''), '-', ' '), '_', ' ') || ' '), ' oil ') > 0
      OR INSTR(LOWER(' ' || REPLACE(REPLACE(COALESCE(${p}.part_name, ''), '-', ' '), '_', ' ') || ' '), ' lube ') > 0
      OR LOWER(TRIM(COALESCE(${p}.part_name, ''))) LIKE 'lubricant%'
      OR LOWER(TRIM(COALESCE(${p}.part_code, ''))) LIKE 'oil%'
      OR LOWER(TRIM(COALESCE(${p}.part_code, ''))) LIKE 'lube%'
    )
  )
)`;
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

/** Upcoming service kit + labor cost per maintenance plan (historical averages or saved manual inputs). */
function buildUpcomingServiceCostForecasts(dbConn, plans, opts = {}) {
  const nearDueHours = Math.max(1, Number(opts.nearDueHours || 50));
  const maxRemainingHours = opts.maxRemainingHours != null
    ? Math.max(0, Number(opts.maxRemainingHours))
    : Math.max(nearDueHours, Number(opts.horizonHours || 100));
  const ctx = opts.ctx || buildStockCostSqlContext(dbConn);
  const { hasTable, hasColumn, closedStatuses, woCloseExpr, smOutSql, oilPartSql, smCostWithParts, smCostNoParts } = ctx;

  const forecastInputs = hasTable("weekly_forum_service_inputs")
    ? dbConn.prepare(`
        SELECT plan_id, oil_part_code, oil_qty, parts_part_code, parts_qty, items_json, notes,
          COALESCE(labor_total, 0) AS labor_total
        FROM weekly_forum_service_inputs
      `).all()
    : [];
  const inputByPlan = new Map((forecastInputs || []).map((r) => [Number(r.plan_id || 0), r]));

  const getAssetHoursSafe = (assetId) => {
    try {
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
                  AND COALESCE(w.reference_id, 0) = ?
                  AND LOWER(COALESCE(w.status, '')) IN (${closedStatuses})
              `).get(planId)
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
                  AND COALESCE(w.reference_id, 0) = ?
                  AND LOWER(COALESCE(w.status, '')) IN (${closedStatuses})
              `).get(planId)
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
                  AND COALESCE(w.reference_id, 0) = ?
                  AND ${woCloseExpr} IS NOT NULL
                  AND LOWER(COALESCE(w.status, '')) IN (${closedStatuses})
              )
          `).get(lubeCostDefault, assetId, planId)
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
              AND COALESCE(w.reference_id, 0) = ?
              AND LOWER(COALESCE(w.status, '')) IN (${closedStatuses})
          `).get(planId)
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
      const hasManualParts = pricedItems.length > 0;
      const hasManualLabor = manualLaborTotal > 0;
      const hasManualOverride = hasManualParts || hasManualLabor;
      const estKitCost = Number((hasManualParts ? (manualOilCost + manualPartsCost) : serviceKitCost).toFixed(2));
      const estLaborCost = Number((hasManualLabor ? manualLaborTotal : avgLaborCost).toFixed(2));
      const estTotalCost = Number((estKitCost + estLaborCost).toFixed(2));
      const costSource = hasManualOverride
        ? (hasManualParts && hasManualLabor ? "manual_parts_and_labor" : hasManualParts ? "manual_store_pricing" : "manual_labor")
        : (serviceEvents > 0 || laborEvents > 0 ? "historical_average" : "none");

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
        needs_manual_input: costSource === "none",
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
          manual: {
            oil_cost_total: Number(manualOilCost.toFixed(2)),
            parts_cost_total: Number(manualPartsCost.toFixed(2)),
            labor_total: Number(manualLaborTotal.toFixed(2)),
            items: pricedItems,
            notes: String(manual?.notes || ""),
          },
        },
      };
    })
    .filter((r) => Number(r.remaining_hours || 0) <= maxRemainingHours)
    .sort((a, b) => Number(a.remaining_hours || 0) - Number(b.remaining_hours || 0));
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
  app.get("/plans", async (req, reply) => {
    try {
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

      const plans = enrichPlansWithNextService(rows, getAssetCurrentHours, 50);

      return reply.send({
        ok: true,
        plans
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });
  // =====================================================
  // MAINTENANCE PLANS - CREATE
  // POST /api/maintenance/plans
  // =====================================================
  app.post("/plans", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const service_name = String(req.body?.service_name || "").trim();
      const interval_hours = Number(req.body?.interval_hours || 0);
      const active = Number(req.body?.active ?? 1) ? 1 : 0;

      if (!asset_id || !service_name || interval_hours <= 0) {
        return reply.code(400).send({
          ok: false,
          error: "asset_id, service_name and interval_hours are required"
        });
      }

      const asset = db.prepare(`
        SELECT id, asset_code
        FROM assets
        WHERE id = ?
      `).get(asset_id);

      if (!asset) {
        return reply.code(404).send({
          ok: false,
          error: "Asset not found"
        });
      }

      const last_service_hours = snapLastServiceHours(
        Number(req.body?.last_service_hours || 0),
        interval_hours,
        asset.asset_code,
      );

      const result = db.prepare(`
        INSERT INTO maintenance_plans (
          asset_id,
          service_name,
          interval_hours,
          last_service_hours,
          active
        )
        VALUES (?, ?, ?, ?, ?)
      `).run(
        asset_id,
        service_name,
        interval_hours,
        last_service_hours,
        active
      );
      writeAudit(db, req, {
        module: "maintenance",
        action: "plan.create",
        entity_type: "maintenance_plan",
        entity_id: String(Number(result.lastInsertRowid || 0)),
        after: { asset_id, service_name, interval_hours, last_service_hours, active },
      });

      return reply.send({
        ok: true,
        id: Number(result.lastInsertRowid)
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - UPDATE
  // PUT /api/maintenance/plans/:id
  // =====================================================
  app.put("/plans/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) {
        return reply.code(400).send({
          ok: false,
          error: "Invalid plan id"
        });
      }

      const existing = db.prepare(`
        SELECT *
        FROM maintenance_plans
        WHERE id = ?
      `).get(id);

      if (!existing) {
        return reply.code(404).send({
          ok: false,
          error: "Maintenance plan not found"
        });
      }

      const asset_id =
        req.body?.asset_id != null ? Number(req.body.asset_id) : Number(existing.asset_id);

      const service_name =
        req.body?.service_name != null
          ? String(req.body.service_name).trim()
          : String(existing.service_name || "").trim();

      const interval_hours =
        req.body?.interval_hours != null
          ? Number(req.body.interval_hours)
          : Number(existing.interval_hours || 0);

      const last_service_hours_raw =
        req.body?.last_service_hours != null
          ? Number(req.body.last_service_hours)
          : Number(existing.last_service_hours || 0);

      const active =
        req.body?.active != null
          ? (Number(req.body.active) ? 1 : 0)
          : Number(existing.active || 0);

      if (!asset_id || !service_name || interval_hours <= 0) {
        return reply.code(400).send({
          ok: false,
          error: "asset_id, service_name and interval_hours are required"
        });
      }

      const asset = db.prepare(`
        SELECT id, asset_code
        FROM assets
        WHERE id = ?
      `).get(asset_id);

      if (!asset) {
        return reply.code(404).send({
          ok: false,
          error: "Asset not found"
        });
      }

      const last_service_hours = snapLastServiceHours(
        last_service_hours_raw,
        interval_hours,
        asset.asset_code,
      );

      db.prepare(`
        UPDATE maintenance_plans
        SET
          asset_id = ?,
          service_name = ?,
          interval_hours = ?,
          last_service_hours = ?,
          active = ?
        WHERE id = ?
      `).run(
        asset_id,
        service_name,
        interval_hours,
        last_service_hours,
        active,
        id
      );
      writeAudit(db, req, {
        module: "maintenance",
        action: "plan.update",
        entity_type: "maintenance_plan",
        entity_id: String(id),
        before: existing,
        after: { asset_id, service_name, interval_hours, last_service_hours, active },
      });

      return reply.send({ ok: true });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - DELETE
  // DELETE /api/maintenance/plans/:id
  // =====================================================
  app.delete("/plans/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) {
        return reply.code(400).send({ ok: false, error: "Invalid plan id" });
      }

      const existing = db.prepare(`
        SELECT *
        FROM maintenance_plans
        WHERE id = ?
      `).get(id);

      if (!existing) {
        return reply.code(404).send({ ok: false, error: "Maintenance plan not found" });
      }

      db.prepare(`
        DELETE FROM maintenance_plans WHERE id = ?
      `).run(id);
      writeAudit(db, req, {
        module: "maintenance",
        action: "plan.delete",
        entity_type: "maintenance_plan",
        entity_id: String(id),
        before: existing,
      });

      return reply.send({ ok: true });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - TOGGLE ACTIVE
  // PATCH /api/maintenance/plans/:id/toggle
  // =====================================================
  app.patch("/plans/:id/toggle", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) {
        return reply.code(400).send({ ok: false, error: "Invalid plan id" });
      }

      const plan = db.prepare(`
        SELECT id, active
        FROM maintenance_plans
        WHERE id = ?
      `).get(id);

      if (!plan) {
        return reply.code(404).send({ ok: false, error: "Maintenance plan not found" });
      }

      const newActive = plan.active ? 0 : 1;
      db.prepare(`
        UPDATE maintenance_plans
        SET active = ?
        WHERE id = ?
      `).run(newActive, id);

      return reply.send({ ok: true, id, active: newActive });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // =====================================================
  // MAINTENANCE PLANS - REBASE LAST SERVICE TO LIVE HOURS
  // POST /api/maintenance/plans/:id/rebase-last-service
  // =====================================================
  app.post("/plans/:id/rebase-last-service", async (req, reply) => {
    try {
      const idFromPath = Number(req.params?.id || 0);
      const idFromBody = Number(req.body?.plan_id || 0);
      const id = idFromPath > 0 ? idFromPath : (idFromBody > 0 ? idFromBody : 0);

      let plan = null;
      if (id > 0) {
        plan = db.prepare(`
          SELECT mp.id, mp.asset_id, mp.service_name, a.asset_code, a.asset_name
          FROM maintenance_plans mp
          JOIN assets a ON a.id = mp.asset_id
          WHERE mp.id = ?
          LIMIT 1
        `).get(id);
      }
      if (!plan) {
        const assetId = Number(req.body?.asset_id || 0);
        const serviceName = String(req.body?.service_name || "").trim();
        if (assetId > 0 && serviceName) {
          plan = db.prepare(`
            SELECT mp.id, mp.asset_id, mp.service_name, a.asset_code, a.asset_name
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.asset_id = ?
              AND UPPER(TRIM(mp.service_name)) = UPPER(TRIM(?))
            ORDER BY mp.active DESC, mp.id DESC
            LIMIT 1
          `).get(assetId, serviceName);
        }
      }

      if (!plan) {
        return reply.code(400).send({ ok: false, error: "Invalid plan id or plan lookup context" });
      }

      const currentInfo = getAssetCurrentHoursInfo(Number(plan.asset_id || 0));
      const liveHours = Number(currentInfo.hours || 0);
      const planMeta = db.prepare(`
        SELECT interval_hours FROM maintenance_plans WHERE id = ?
      `).get(Number(plan.id));
      const safeHours = snapLastServiceHours(
        Number.isFinite(liveHours) ? liveHours : 0,
        Number(planMeta?.interval_hours || 0),
        plan.asset_code,
      );

      db.prepare(`
        UPDATE maintenance_plans
        SET last_service_hours = ?
        WHERE id = ?
      `).run(safeHours, Number(plan.id));

      db.prepare(`
        UPDATE maintenance_plans
        SET last_service_hours = ?
        WHERE asset_id = ?
          AND active = 1
      `).run(safeHours, Number(plan.asset_id || 0));

      return reply.send({
        ok: true,
        id: Number(plan.id),
        asset_id: Number(plan.asset_id || 0),
        asset_code: plan.asset_code,
        asset_name: plan.asset_name,
        service_name: plan.service_name,
        last_service_hours: safeHours,
        source: currentInfo.source
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

    // =====================================================
  // GET LIVE HOURS FOR ONE ASSET
  // GET /api/maintenance/asset/:id/live-hours?as_of=YYYY-MM-DD (optional; meter/usage up to that date)
  // =====================================================
  app.get("/asset/:id/live-hours", async (req, reply) => {
    try {
      const assetId = Number(req.params?.id || 0);
      if (!assetId) {
        return reply.code(400).send({
          ok: false,
          error: "Invalid asset id"
        });
      }

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name
        FROM assets
        WHERE id = ?
      `).get(assetId);

      if (!asset) {
        return reply.code(404).send({
          ok: false,
          error: "Asset not found"
        });
      }

      const asOf = String(req.query?.as_of || "").trim();
      const currentInfo = isDate(asOf)
        ? getAssetHoursInfoAsOf(assetId, asOf)
        : getAssetCurrentHoursInfo(assetId);
      const current_hours = Number(currentInfo.hours || 0);

      const assetPlans = db.prepare(`
        SELECT id, service_name, interval_hours, last_service_hours, active
        FROM maintenance_plans
        WHERE asset_id = ?
          AND active = 1
        ORDER BY interval_hours ASC, id ASC
      `).all(assetId);
      const next_service = assetPlans.length
        ? resolveNextServiceForAssetPlans(assetPlans, current_hours, asset.asset_code)
        : null;

      return reply.send({
        ok: true,
        asset_id: asset.id,
        asset_code: asset.asset_code,
        asset_name: asset.asset_name,
        as_of: isDate(asOf) ? asOf : null,
        current_hours: Number(current_hours.toFixed(1)),
        current_hours_source: currentInfo.source,
        next_service,
        rotating_schedule: hasRotatingSchedule(assetPlans),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // LIST MAINTENANCE DUE
  // GET /api/maintenance/due?date=2026-02-27&near_due_hours=50
  // =====================================================
  app.get("/due", async (req, reply) => {
    try {
      const date = String(req.query?.date || "").trim();
      const nearDueHours = Math.max(1, Number(req.query?.near_due_hours || 50));
      if (date && !isDate(date)) {
        return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
      }

      const rows = date
        ? db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              mp.active,
              a.asset_code,
              a.asset_name,
              a.category,
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
              AND a.archived = 0
            ORDER BY a.asset_code, mp.id
          `).all(date)
        : db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              mp.active,
              a.asset_code,
              a.asset_name,
              a.category,
              IFNULL((
                SELECT SUM(dh.hours_run)
                FROM daily_hours dh
                WHERE dh.asset_id = a.id
                  AND dh.is_used = 1
                  AND dh.hours_run > 0
              ), 0) AS current_hours
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.is_standby = 0
              AND a.archived = 0
            ORDER BY a.asset_code, mp.id
          `).all();

      const getHours = (assetId) => {
        if (date) return Number(rows.find((r) => Number(r.asset_id) === Number(assetId))?.current_hours ?? getAssetCurrentHours(assetId));
        return getAssetCurrentHours(assetId);
      };
      const due = enrichDueRowsWithEstimates(
        db,
        buildDueListFromPlans(rows, getHours, nearDueHours),
        { as_of: date || new Date().toISOString().slice(0, 10), history_days: 14 },
      );

      return reply.send({
        ok: true,
        as_of: date || null,
        near_due_hours: nearDueHours,
        due
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // MAINTENANCE HISTORY (single-line per asset+service)
  // GET /api/maintenance/history?as_of=YYYY-MM-DD&days=14
  // =====================================================
  app.get("/history", async (req, reply) => {
    try {
      const as_of = String(req.query?.as_of || "").trim();
      const days = Math.max(3, Math.min(120, Number(req.query?.days || 14)));
      if (as_of && !isDate(as_of)) return reply.code(400).send({ error: "as_of must be YYYY-MM-DD" });

      const endDate = as_of || new Date().toISOString().slice(0, 10);
      const startD = new Date(endDate + "T00:00:00");
      startD.setDate(startD.getDate() - (days - 1));
      const startDate = startD.toISOString().slice(0, 10);

      const plans = db.prepare(`
        SELECT
          mp.id AS plan_id,
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
        WHERE mp.active = 1
          AND a.active = 1
          AND a.is_standby = 0
          AND a.archived = 0
        ORDER BY a.asset_code ASC, mp.service_name ASC
      `).all();

      const getLastServiced = db.prepare(`
        SELECT
          DATE(COALESCE(w.closed_at, w.completed_at)) AS last_serviced_date,
          COALESCE(w.closed_at, w.completed_at) AS last_serviced_at
        FROM work_orders w
        WHERE w.source = 'service'
          AND w.reference_id = ?
          AND w.status IN ('completed', 'approved', 'closed')
        ORDER BY COALESCE(w.closed_at, w.completed_at) DESC
        LIMIT 1
      `);
      const getLastBackfillServiced = db.prepare(`
        SELECT
          h.id AS history_id,
          DATE(h.service_date) AS last_serviced_date,
          h.service_date AS last_serviced_at
        FROM maintenance_service_history h
        WHERE h.asset_id = ?
          AND UPPER(TRIM(h.service_name)) = UPPER(TRIM(?))
        ORDER BY h.service_date DESC, h.id DESC
        LIMIT 1
      `);

      const getAvgDaily = db.prepare(`
        SELECT
          COALESCE(SUM(hours_run), 0) AS total_run,
          COUNT(DISTINCT work_date) AS day_count
        FROM daily_hours
        WHERE asset_id = ?
          AND is_used = 1
          AND hours_run > 0
          AND work_date BETWEEN ? AND ?
      `);

      const addDays = (dateStr, add) => {
        const d = new Date(dateStr + "T00:00:00");
        d.setDate(d.getDate() + Math.round(add));
        return d.toISOString().slice(0, 10);
      };

      const getLastServicedAnyOnAsset = db.prepare(`
        SELECT
          DATE(COALESCE(w.closed_at, w.completed_at)) AS last_serviced_date,
          COALESCE(w.closed_at, w.completed_at) AS last_serviced_at
        FROM work_orders w
        WHERE w.source = 'service'
          AND w.asset_id = ?
          AND w.status IN ('completed', 'approved', 'closed')
        ORDER BY COALESCE(w.closed_at, w.completed_at) DESC
        LIMIT 1
      `);

      const rows = [];
      const byAsset = groupActivePlansByAsset(plans);
      for (const [, assetPlans] of byAsset) {
        const sample = assetPlans[0];
        const assetId = Number(sample.asset_id || 0);
        const currentInfo = getAssetCurrentHoursInfo(assetId);
        const current = Number(currentInfo.hours || 0);
        const rotating = hasRotatingSchedule(assetPlans);

        if (rotating) {
          const resolved = resolveNextServiceForAssetPlans(
            assetPlans,
            current,
            String(sample.asset_code || ""),
          );
          const lastWo = getLastServicedAnyOnAsset.get(assetId);
          const lastBackfill = db.prepare(`
            SELECT
              h.id AS history_id,
              DATE(h.service_date) AS last_serviced_date,
              h.service_date AS last_serviced_at
            FROM maintenance_service_history h
            WHERE h.asset_id = ?
            ORDER BY h.service_date DESC, h.id DESC
            LIMIT 1
          `).get(assetId);
          const lastWoAt = String(lastWo?.last_serviced_at || "");
          const lastBackfillAt = String(lastBackfill?.last_serviced_at || "");
          const useBackfill = Boolean(lastBackfillAt && (!lastWoAt || lastBackfillAt > lastWoAt));
          const last = useBackfill ? lastBackfill : lastWo;
          const avgRow = getAvgDaily.get(assetId, startDate, endDate);
          const totalRun = Number(avgRow?.total_run || 0);
          const dayCount = Number(avgRow?.day_count || 0);
          const avgDaily = dayCount > 0 ? totalRun / dayCount : 0;
          const remaining = Number(resolved?.remaining_hours || 0);
          const estDays = avgDaily > 0 ? Math.max(0, remaining / avgDaily) : null;
          const estDate = estDays == null ? null : addDays(endDate, estDays);

          const dueMeta = classifyServiceDue(
            remaining,
            String(sample.asset_code || ""),
            Number(resolved?.interval_hours || resolved?.next_service_interval || 0),
            50,
          );

          rows.push({
            plan_id: Number(resolved?.plan_id || 0),
            asset_id: assetId,
            asset_code: sample.asset_code,
            asset_name: sample.asset_name,
            service_name: resolved?.service_name || sample.service_name,
            last_serviced_date: last?.last_serviced_date || null,
            last_service_source: useBackfill ? "backfill" : (last?.last_serviced_at ? "work_order" : null),
            last_service_history_id: useBackfill ? Number(lastBackfill?.history_id || 0) : null,
            current_hours: Number(current.toFixed(2)),
            current_hours_source: currentInfo.source,
            remaining_hours: Number(remaining.toFixed(2)),
            avg_daily_hours: Number(avgDaily.toFixed(2)),
            estimated_service_date: estDate,
            schedule_mode: "rotating",
            next_due_hours: resolved?.next_due_hours ?? null,
            last_service_hours: resolved?.last_service_hours ?? null,
            meter_unit: dueMeta.meter_unit,
            near_due_threshold: dueMeta.near_due_threshold,
            status: dueMeta.status,
            is_almost_due: dueMeta.is_almost_due,
          });
          continue;
        }

        for (const p of assetPlans) {
          const resolved = resolveLegacyPlanDue(p, current, String(p.asset_code || ""));
          const remaining = resolved.remaining_hours;

          const lastWo = getLastServiced.get(Number(p.plan_id || 0));
          const lastBackfill = getLastBackfillServiced.get(assetId, String(p.service_name || ""));
          const lastWoAt = String(lastWo?.last_serviced_at || "");
          const lastBackfillAt = String(lastBackfill?.last_serviced_at || "");
          const useBackfill = Boolean(lastBackfillAt && (!lastWoAt || lastBackfillAt > lastWoAt));
          const last = useBackfill ? lastBackfill : lastWo;
          const avgRow = getAvgDaily.get(assetId, startDate, endDate);
          const totalRun = Number(avgRow?.total_run || 0);
          const dayCount = Number(avgRow?.day_count || 0);
          const avgDaily = dayCount > 0 ? totalRun / dayCount : 0;

          const estDays = avgDaily > 0 ? Math.max(0, remaining / avgDaily) : null;
          const estDate = estDays == null ? null : addDays(endDate, estDays);

          const dueMeta = classifyServiceDue(
            remaining,
            String(p.asset_code || ""),
            Number(resolved.interval_hours || 0),
            50,
          );

          rows.push({
            plan_id: Number(p.plan_id || 0),
            asset_id: assetId,
            asset_code: p.asset_code,
            asset_name: p.asset_name,
            service_name: p.service_name,
            last_serviced_date: last?.last_serviced_date || null,
            last_service_source: useBackfill ? "backfill" : (last?.last_serviced_at ? "work_order" : null),
            last_service_history_id: useBackfill ? Number(lastBackfill?.history_id || 0) : null,
            current_hours: Number(current.toFixed(2)),
            current_hours_source: currentInfo.source,
            remaining_hours: Number(remaining.toFixed(2)),
            avg_daily_hours: Number(avgDaily.toFixed(2)),
            estimated_service_date: estDate,
            schedule_mode: "grid",
            next_due_hours: resolved.next_due_hours,
            last_service_hours: resolved.last_service_hours,
            meter_unit: dueMeta.meter_unit,
            near_due_threshold: dueMeta.near_due_threshold,
            status: dueMeta.status,
            is_almost_due: dueMeta.is_almost_due,
          });
        }
      }

      return reply.send({ ok: true, as_of: endDate, range: { start: startDate, end: endDate }, rows });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

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

    const by_asset = assetRows.map((a) => {
      const aid = Number(a.asset_id || 0);
      const operating_hours = Number(runByAsset.get(aid) || 0);
      const rel = reliabilityByAsset.get(aid) || { failure_count: 0, downtime_hours: 0 };
      const failure_count = Number(rel.failure_count || 0);
      const downtime_hours = Number(rel.downtime_hours || 0);
      const { mtbf_hours, lttr_hours } = computeMtbfLttr(operating_hours, failure_count, downtime_hours);
      return {
        asset_id: aid,
        asset_code: String(a.asset_code || ""),
        asset_name: String(a.asset_name || ""),
        category: String(a.category || ""),
        failure_count,
        operating_hours: round2(operating_hours),
        downtime_hours: round2(downtime_hours),
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
      formulas: {
        mtbf: "operating_hours / failure_count",
        lttr: "downtime_hours / failure_count",
        failures: "distinct breakdown incidents with downtime > 0 in period (daily logs, else header when reported in period, else linked breakdown WO wall-clock)",
        operating_hours: "sum of daily_hours.hours_run (is_used=1) in period",
        downtime_hours: "per-incident downtime in period (same sources as failures)",
      },
      summary: {
        failure_count,
        operating_hours: round2(operating_hours),
        downtime_hours: round2(downtime_hours),
        mtbf_hours,
        lttr_hours,
      },
      by_asset,
      incidents: incidentsWithAsset,
    };
  }

  app.get("/reliability", async (req, reply) => {
    try {
      const endDate = String(req.query?.end || "").trim() || new Date().toISOString().slice(0, 10);
      const startDate = String(req.query?.start || "").trim() || (() => {
        const d = new Date(`${endDate}T00:00:00`);
        d.setDate(d.getDate() - 29);
        return d.toISOString().slice(0, 10);
      })();
      if (!isDate(startDate) || !isDate(endDate)) {
        return reply.code(400).send({ ok: false, error: "start and end must be YYYY-MM-DD" });
      }
      if (startDate > endDate) {
        return reply.code(400).send({ ok: false, error: "start must be <= end" });
      }
      const asset_ids = parseReliabilityAssetIds(req.query?.asset_ids);
      const category = String(req.query?.category || "").trim();
      const data = buildMaintenanceReliabilityReport(startDate, endDate, { asset_ids, category });
      return reply.send({ ok: true, ...data });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/reliability.xlsx", async (req, reply) => {
    try {
      const endDate = String(req.query?.end || "").trim() || new Date().toISOString().slice(0, 10);
      const startDate = String(req.query?.start || "").trim() || (() => {
        const d = new Date(`${endDate}T00:00:00`);
        d.setDate(d.getDate() - 29);
        return d.toISOString().slice(0, 10);
      })();
      if (!isDate(startDate) || !isDate(endDate) || startDate > endDate) {
        return reply.code(400).send({ ok: false, error: "Provide valid start/end dates" });
      }
      const asset_ids = parseReliabilityAssetIds(req.query?.asset_ids);
      const category = String(req.query?.category || "").trim();
      const data = buildMaintenanceReliabilityReport(startDate, endDate, { asset_ids, category });

      const wb = new ExcelJS.Workbook();
      wb.creator = "IRONLOG";
      wb.created = new Date();

      const summary = wb.addWorksheet("Summary");
      summary.addRow(["MTBF / LTTR Report"]);
      summary.addRow(["Period", `${startDate} to ${endDate}`]);
      summary.addRow(["Category filter", category || "All"]);
      summary.addRow(["Assets in scope", data.asset_filter_count]);
      summary.addRow([]);
      summary.addRow(["Metric", "Value"]);
      summary.addRow(["Failures", data.summary.failure_count]);
      summary.addRow(["Operating hours", data.summary.operating_hours]);
      summary.addRow(["Downtime hours", data.summary.downtime_hours]);
      summary.addRow(["MTBF (hours)", data.summary.mtbf_hours ?? ""]);
      summary.addRow(["LTTR (hours)", data.summary.lttr_hours ?? ""]);

      const assets = wb.addWorksheet("By Asset");
      assets.columns = [
        { header: "Asset", key: "asset_code", width: 14 },
        { header: "Name", key: "asset_name", width: 28 },
        { header: "Category", key: "category", width: 18 },
        { header: "Failures", key: "failure_count", width: 12 },
        { header: "Operating h", key: "operating_hours", width: 14 },
        { header: "Downtime h", key: "downtime_hours", width: 14 },
        { header: "MTBF h", key: "mtbf_hours", width: 12 },
        { header: "LTTR h", key: "lttr_hours", width: 12 },
      ];
      assets.addRows(data.by_asset || []);

      const incSheet = wb.addWorksheet("Incidents");
      incSheet.columns = [
        { header: "Asset", key: "asset_code", width: 14 },
        { header: "Breakdown #", key: "breakdown_id", width: 12 },
        { header: "Report date", key: "breakdown_date", width: 14 },
        { header: "WO #", key: "work_order_id", width: 10 },
        { header: "Downtime h (period)", key: "downtime_hours", width: 18 },
        { header: "Source", key: "downtime_source", width: 16 },
        { header: "Log h (period)", key: "log_downtime_in_period", width: 14 },
        { header: "Header h (total)", key: "header_downtime_hours", width: 14 },
        { header: "WO opened", key: "work_order_opened_at", width: 20 },
        { header: "WO closed", key: "work_order_closed_at", width: 20 },
        { header: "Description", key: "description", width: 40 },
      ];
      incSheet.addRows(data.incidents || []);

      const buf = await wb.xlsx.writeBuffer();
      reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Content-Disposition", `attachment; filename="IRONLOG_MTBF_LTTR_${startDate}_to_${endDate}.xlsx"`)
        .send(Buffer.from(buf));
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

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

  // =====================================================
  // MAINTENANCE INSIGHTS (High-Impact analytics starter)
  // GET /api/maintenance/insights?start=YYYY-MM-DD&end=YYYY-MM-DD&near_due_hours=50
  // =====================================================
  app.get("/insights", async (req, reply) => {
    try {
      const endDate = String(req.query?.end || "").trim() || new Date().toISOString().slice(0, 10);
      if (!isDate(endDate)) return reply.code(400).send({ ok: false, error: "end must be YYYY-MM-DD" });
      const startDate = String(req.query?.start || "").trim() || (() => {
        const d = new Date(`${endDate}T00:00:00`);
        d.setDate(d.getDate() - 29);
        return d.toISOString().slice(0, 10);
      })();
      if (!isDate(startDate)) return reply.code(400).send({ ok: false, error: "start must be YYYY-MM-DD" });
      const nearDueHours = Math.max(1, Number(req.query?.near_due_hours || 50));
      const predictiveHorizonHours = Math.max(nearDueHours, Number(req.query?.predictive_horizon_hours || 100));
      const checklistFailThreshold = Math.max(1, Number(req.query?.checklist_fail_threshold || 2));
      const fuelVarianceThreshold = Math.max(0, Number(req.query?.fuel_variance_threshold || 15));

      const plans = db.prepare(`
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
      `).all();
      const atRiskPlans = plans
        .map((p) => {
          const currentInfo = getAssetCurrentHoursInfo(Number(p.asset_id || 0));
          const current = Number(currentInfo.hours || 0);
          const nextDue = Number(p.last_service_hours || 0) + Number(p.interval_hours || 0);
          const remaining = Number((nextDue - current).toFixed(2));
          return {
            plan_id: Number(p.plan_id || 0),
            asset_id: Number(p.asset_id || 0),
            asset_code: p.asset_code,
            asset_name: p.asset_name,
            service_name: p.service_name,
            current_hours: Number(current.toFixed(2)),
            next_due_hours: Number(nextDue.toFixed(2)),
            remaining_hours: remaining,
            risk: remaining <= 0 ? "OVERDUE" : remaining <= nearDueHours ? "NEAR_DUE" : remaining <= predictiveHorizonHours ? "WATCH" : "OK",
          };
        })
        .filter((r) => r.risk !== "OK")
        .sort((a, b) => Number(a.remaining_hours || 0) - Number(b.remaining_hours || 0))
        .slice(0, 30);

      const hasChecklistDetailJson = hasColumn("manager_inspections", "checklist_detail_json");
      const managerInspections = db.prepare(`
        SELECT asset_id, checklist_json, ${hasChecklistDetailJson ? "checklist_detail_json" : "NULL AS checklist_detail_json"}
        FROM manager_inspections
        WHERE inspection_date BETWEEN ? AND ?
      `).all(startDate, endDate);
      const failByAsset = new Map();
      const parseChecklistFailCount = (rawChecklist) => {
        try {
          const parsed = JSON.parse(String(rawChecklist || "null"));
          if (!parsed) return 0;
          if (Array.isArray(parsed)) return parsed.filter((x) => x?.ok === false).length;
          if (typeof parsed === "object") {
            if (parsed.checklist && typeof parsed.checklist === "object") {
              return Object.values(parsed.checklist).filter((v) => {
                const s = String(v || "").trim().toLowerCase();
                return s === "attention" || s === "unsafe" || s === "fail" || s === "failed";
              }).length;
            }
            return Object.values(parsed).filter((v) => {
              const s = String(v || "").trim().toLowerCase();
              return s === "attention" || s === "unsafe" || s === "fail" || s === "failed";
            }).length;
          }
          return 0;
        } catch {
          return 0;
        }
      };
      for (const r of managerInspections) {
        const aid = Number(r.asset_id || 0);
        if (!aid) continue;
        const fails = parseChecklistFailCount(r.checklist_json);
        if (!fails) continue;
        failByAsset.set(aid, Number(failByAsset.get(aid) || 0) + fails);
      }
      const repeatedChecklistFailures = [...failByAsset.entries()]
        .filter(([, failCount]) => failCount >= checklistFailThreshold)
        .map(([asset_id, fail_count]) => {
          const a = db.prepare(`SELECT asset_code, asset_name FROM assets WHERE id = ? LIMIT 1`).get(asset_id) || {};
          return {
            asset_id: Number(asset_id),
            asset_code: String(a.asset_code || ""),
            asset_name: String(a.asset_name || ""),
            fail_count: Number(fail_count || 0),
          };
        })
        .sort((a, b) => Number(b.fail_count || 0) - Number(a.fail_count || 0))
        .slice(0, 20);

      const fuelAnomalies = hasColumn("assets", "baseline_fuel_l_per_hour")
        ? db.prepare(`
            SELECT
              a.id AS asset_id,
              a.asset_code,
              a.asset_name,
              COALESCE(a.baseline_fuel_l_per_hour, 0) AS baseline_lph,
              COALESCE(SUM(fl.liters), 0) AS fuel_liters,
              COALESCE(SUM(CASE WHEN COALESCE(fl.hours_run, 0) > 0 THEN fl.hours_run ELSE 0 END), 0) AS hours_run,
              COUNT(fl.id) AS fills
            FROM assets a
            JOIN fuel_logs fl ON fl.asset_id = a.id
            WHERE fl.log_date BETWEEN ? AND ?
              AND a.active = 1
            GROUP BY a.id
            HAVING fills >= 2 AND hours_run > 0 AND baseline_lph > 0
          `).all(startDate, endDate)
            .map((r) => {
              const actual = Number(r.fuel_liters || 0) / Number(r.hours_run || 1);
              const baseline = Number(r.baseline_lph || 0);
              const ratio = baseline > 0 ? actual / baseline : 0;
              return {
                asset_id: Number(r.asset_id || 0),
                asset_code: r.asset_code,
                asset_name: r.asset_name,
                baseline_lph: Number(baseline.toFixed(3)),
                actual_lph: Number(actual.toFixed(3)),
                variance_pct: Number(((ratio - 1) * 100).toFixed(1)),
              };
            })
            .filter((r) => Number(r.variance_pct || 0) >= fuelVarianceThreshold)
            .sort((a, b) => Number(b.variance_pct || 0) - Number(a.variance_pct || 0))
            .slice(0, 20)
        : [];

      const predictive = {
        at_risk_plans: atRiskPlans,
        repeated_checklist_failures: repeatedChecklistFailures,
        fuel_anomalies: fuelAnomalies,
      };

      const upcomingPlans = plans
        .map((p) => {
          const current = Number(getAssetCurrentHoursInfo(Number(p.asset_id || 0)).hours || 0);
          const nextDue = Number(p.last_service_hours || 0) + Number(p.interval_hours || 0);
          return {
            service_name: String(p.service_name || "").trim(),
            remaining_hours: Number((nextDue - current).toFixed(2)),
          };
        })
        .filter((r) => r.service_name && Number(r.remaining_hours || 0) <= predictiveHorizonHours);
      const upcomingByService = upcomingPlans.reduce((m, r) => {
        const key = String(r.service_name || "").trim().toLowerCase();
        if (!key) return m;
        const cur = m.get(key) || { service_name: r.service_name, due_count: 0 };
        cur.due_count += 1;
        m.set(key, cur);
        return m;
      }, new Map());

      const canReadMaintenanceParts = hasTable("maintenance_records") && hasTable("maintenance_parts");
      const canJoinPartsCatalog = canReadMaintenanceParts && hasTable("parts");
      const historicalParts = canReadMaintenanceParts
        ? db.prepare(`
            SELECT
              LOWER(TRIM(COALESCE(mr.service_type, ''))) AS service_key,
              COALESCE(mp.part_name, '') AS part_name,
              AVG(COALESCE(mp.quantity, 0)) AS avg_qty,
              AVG(COALESCE(mp.quantity, 0) * COALESCE(${canJoinPartsCatalog ? "p.unit_cost" : "0"}, 0)) AS avg_unit_cost
            FROM maintenance_records mr
            JOIN maintenance_parts mp ON mp.maintenance_record_id = mr.id
            ${canJoinPartsCatalog ? "LEFT JOIN parts p ON LOWER(TRIM(p.part_name)) = LOWER(TRIM(mp.part_name))" : ""}
            WHERE DATE(mr.maintenance_date) BETWEEN DATE(?) AND DATE(?)
              AND TRIM(COALESCE(mr.service_type, '')) <> ''
              AND TRIM(COALESCE(mp.part_name, '')) <> ''
            GROUP BY 1, 2
          `).all(startDate, endDate)
        : [];
      const partDemandMap = new Map();
      for (const r of historicalParts) {
        const serviceKey = String(r.service_key || "");
        const due = upcomingByService.get(serviceKey);
        if (!due) continue;
        const part = String(r.part_name || "").trim();
        if (!part) continue;
        const suggested = Number(r.avg_qty || 0) * Number(due.due_count || 0);
        if (!Number.isFinite(suggested) || suggested <= 0) continue;
        const avgUnitCost = Number(r.avg_unit_cost || 0);
        const estCost = avgUnitCost > 0 ? avgUnitCost * Number(due.due_count || 0) : 0;
        const cur = partDemandMap.get(part) || { part_name: part, suggested_qty: 0, est_cost: 0, linked_services: new Set() };
        cur.suggested_qty += suggested;
        cur.est_cost += estCost;
        cur.linked_services.add(due.service_name);
        partDemandMap.set(part, cur);
      }
      const getPartOnHand = db.prepare(`
        SELECT COALESCE(SUM(
          CASE
            WHEN LOWER(COALESCE(movement_type, '')) = 'out' THEN -ABS(COALESCE(quantity, 0))
            ELSE COALESCE(quantity, 0)
          END
        ), 0) AS on_hand
        FROM stock_movements
        WHERE LOWER(TRIM(COALESCE(reference, ''))) = LOWER(TRIM(?))
      `);
      const partsDemand = [...partDemandMap.values()]
        .map((x) => {
          const onHand = Number(getPartOnHand.get(x.part_name)?.on_hand || 0);
          const suggested = Number(x.suggested_qty || 0);
          return {
            part_name: x.part_name,
            suggested_qty: Number(suggested.toFixed(1)),
            est_cost: Number(Number(x.est_cost || 0).toFixed(2)),
            on_hand: Number(onHand.toFixed(1)),
            gap_qty: Number(Math.max(0, suggested - onHand).toFixed(1)),
            linked_services: [...x.linked_services].slice(0, 4),
          };
        })
        .sort((a, b) => Number(b.est_cost || 0) - Number(a.est_cost || 0) || Number(b.gap_qty || 0) - Number(a.gap_qty || 0))
        .slice(0, 30);

      const upcomingCostForecasts = buildUpcomingServiceCostForecasts(db, plans, {
        nearDueHours,
        horizonHours: predictiveHorizonHours,
      });
      const totalUpcomingCost = upcomingCostForecasts.reduce(
        (s, r) => s + Number(r.forecast?.est_total_cost || 0),
        0,
      );
      const needsManualCount = upcomingCostForecasts.filter((r) => r.needs_manual_input).length;

      const partsPlanning = {
        upcoming_service_count: upcomingPlans.length,
        total_upcoming_cost: Number(totalUpcomingCost.toFixed(2)),
        needs_manual_input_count: needsManualCount,
        suggestions: partsDemand,
        upcoming_cost_forecasts: upcomingCostForecasts.slice(0, 40),
      };

      const woHasOpenedAt = hasColumn("work_orders", "opened_at");
      const woHasUpdatedAt = hasColumn("work_orders", "updated_at");
      const woHasAssignedAt = hasColumn("work_orders", "assigned_at");
      const woHasCompletedAt = hasColumn("work_orders", "completed_at");
      const woHasClosedAt = hasColumn("work_orders", "closed_at");
      const woHasSupervisorDecisionAt = hasColumn("work_orders", "supervisor_decision_at");
      const woDateBasis = woHasOpenedAt
        ? "opened_at"
        : (woHasUpdatedAt ? "updated_at" : (woHasClosedAt ? "closed_at" : "datetime('now')"));
      const serviceWos = db.prepare(`
        SELECT
          id,
          ${woHasOpenedAt ? "opened_at" : "NULL AS opened_at"},
          ${woHasAssignedAt ? "assigned_at" : "NULL AS assigned_at"},
          ${woHasCompletedAt ? "completed_at" : "NULL AS completed_at"},
          ${woHasClosedAt ? "closed_at" : "NULL AS closed_at"},
          ${woHasSupervisorDecisionAt ? "supervisor_decision_at" : "NULL AS supervisor_decision_at"}
        FROM work_orders
        WHERE source = 'service'
          AND DATE(COALESCE(${woDateBasis}, '1970-01-01')) BETWEEN DATE(?) AND DATE(?)
      `).all(startDate, endDate);
      const hoursBetween = (a, b) => {
        if (!a || !b) return null;
        const start = new Date(String(a));
        const end = new Date(String(b));
        const ms = Number(end - start);
        if (!Number.isFinite(ms) || ms <= 0) return null;
        return ms / 36e5;
      };
      const collectAvg = (vals) => {
        const clean = vals.filter((n) => Number.isFinite(n) && n >= 0);
        if (!clean.length) return null;
        return Number((clean.reduce((a, b) => a + b, 0) / clean.length).toFixed(2));
      };
      const sla = {
        work_orders: serviceWos.length,
        avg_open_to_assign_hours: collectAvg(serviceWos.map((r) => hoursBetween(r.opened_at, r.assigned_at))),
        avg_open_to_complete_hours: collectAvg(serviceWos.map((r) => hoursBetween(r.opened_at, r.completed_at))),
        avg_complete_to_approve_hours: collectAvg(serviceWos.map((r) => hoursBetween(r.completed_at, r.supervisor_decision_at))),
        avg_open_to_close_hours: collectAvg(serviceWos.map((r) => hoursBetween(r.opened_at, r.closed_at))),
      };

      const costSettingsDefaults = {
        fuel_cost_per_liter_default: 1.5,
        lube_cost_per_qty_default: 4.0,
        labor_cost_per_hour_default: 35.0,
        downtime_cost_per_hour_default: 120.0,
      };
      const costSettings = { ...costSettingsDefaults };
      if (hasTable("cost_settings")) {
        const settingRows = db.prepare(`
          SELECT key, value
          FROM cost_settings
          WHERE key IN (
            'fuel_cost_per_liter_default',
            'lube_cost_per_qty_default',
            'labor_cost_per_hour_default',
            'downtime_cost_per_hour_default'
          )
        `).all();
        for (const row of settingRows) {
          const k = String(row.key || "").trim();
          const v = Number(row.value);
          if (k && Number.isFinite(v)) costSettings[k] = v;
        }
      }
      const laborRate = Number(costSettings.labor_cost_per_hour_default || 35);
      const lubeDefault = Number(costSettings.lube_cost_per_qty_default || 4.0);

      const insightsDowntime = buildMaintenanceInsightsDowntime(startDate, endDate, { scheduledFallback: 10 });
      const breakdownLabor = buildInsightsBreakdownLaborIncidents(db, startDate, endDate, {
        scheduledFallback: 10,
        laborRate,
      });
      const downtime = {
        by_component: insightsDowntime.by_component,
        by_team: insightsDowntime.by_team,
        total_hours: insightsDowntime.total_hours,
        labor: {
          incidents: breakdownLabor.incidents.slice(0, 50),
          totals: breakdownLabor.totals,
        },
      };

      const smCols = hasTable("stock_movements")
        ? db.prepare(`PRAGMA table_info(stock_movements)`).all()
        : [];
      const smHasCreatedAt = smCols.some((c) => String(c.name) === "created_at");
      const smHasMovementDate = smCols.some((c) => String(c.name) === "movement_date");
      const smDateExpr = smHasCreatedAt
        ? "DATE(sm.created_at)"
        : smHasMovementDate
        ? "DATE(sm.movement_date)"
        : "DATE('now')";

      const activeAssets = db.prepare(`
        SELECT id, asset_code, asset_name
        FROM assets
        WHERE COALESCE(active, 1) = 1
          AND COALESCE(archived, 0) = 0
      `).all();

      const costByAsset = new Map();
      const ensureCostRow = (assetId, assetCode, assetName) => {
        const aid = Number(assetId || 0);
        if (!aid) return null;
        if (!costByAsset.has(aid)) {
          costByAsset.set(aid, {
            asset_id: aid,
            asset_code: String(assetCode || ""),
            asset_name: String(assetName || ""),
            service_jobs: 0,
            downtime_hours: 0,
            repair_labor_hours: 0,
            wo_labor_hours: 0,
            repair_labor_cost: 0,
            wo_labor_cost: 0,
            labor_cost: 0,
            parts_cost: 0,
            lube_cost: 0,
            outsourced_cost: 0,
            total_cost: 0,
          });
        }
        return costByAsset.get(aid);
      };
      for (const a of activeAssets) {
        ensureCostRow(a.id, a.asset_code, a.asset_name);
      }

      if (hasTable("work_orders")) {
        const woRows = db.prepare(`
          SELECT
            a.id AS asset_id,
            a.asset_code,
            a.asset_name,
            COUNT(DISTINCT w.id) AS service_jobs,
            COALESCE(SUM(COALESCE(w.labor_hours, 0)), 0) AS wo_labor_hours,
            COALESCE(SUM(COALESCE(w.labor_hours, 0) * COALESCE(w.labor_rate_per_hour, ?)), 0) AS wo_labor_cost
          FROM work_orders w
          JOIN assets a ON a.id = w.asset_id
          WHERE LOWER(COALESCE(w.source, '')) = 'service'
            AND DATE(COALESCE(w.completed_at, w.closed_at, w.opened_at, w.updated_at)) BETWEEN DATE(?) AND DATE(?)
            AND (
              w.status IN ('completed', 'approved', 'closed')
              OR COALESCE(w.labor_hours, 0) > 0
            )
          GROUP BY a.id
        `).all(laborRate, startDate, endDate);
        for (const r of woRows) {
          const row = ensureCostRow(r.asset_id, r.asset_code, r.asset_name);
          if (!row) continue;
          row.service_jobs = Number(r.service_jobs || 0);
          row.wo_labor_hours = Number(r.wo_labor_hours || 0);
          row.wo_labor_cost = Number(r.wo_labor_cost || 0);
        }
      }

      for (const [assetId, downRow] of insightsDowntime.by_asset.entries()) {
        const asset =
          activeAssets.find((a) => Number(a.id || 0) === Number(assetId))
          || db.prepare(`SELECT id, asset_code, asset_name FROM assets WHERE id = ? LIMIT 1`).get(assetId)
          || {};
        const row = ensureCostRow(assetId, asset.asset_code, asset.asset_name);
        if (!row) continue;
        row.downtime_hours = Number(Number(downRow.downtime_hours || 0).toFixed(2));
      }

      for (const [assetId, labRow] of breakdownLabor.by_asset.entries()) {
        const asset =
          activeAssets.find((a) => Number(a.id || 0) === Number(assetId))
          || db.prepare(`SELECT id, asset_code, asset_name FROM assets WHERE id = ? LIMIT 1`).get(assetId)
          || {};
        const row = ensureCostRow(assetId, asset.asset_code, asset.asset_name);
        if (!row) continue;
        row.repair_labor_hours = Number(Number(labRow.repair_labor_hours || 0).toFixed(2));
        row.repair_labor_cost = Number(Number(labRow.repair_labor_cost || 0).toFixed(2));
      }

      if (hasTable("stock_movements") && hasTable("parts")) {
        const partRows = db.prepare(`
          SELECT
            a.id AS asset_id,
            a.asset_code,
            a.asset_name,
            COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0) AS parts_cost
          FROM stock_movements sm
          JOIN parts p ON p.id = sm.part_id
          LEFT JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
          LEFT JOIN assets a ON a.id = w.asset_id
          WHERE sm.movement_type = 'out'
            AND ${smDateExpr} BETWEEN DATE(?) AND DATE(?)
            AND a.id IS NOT NULL
          GROUP BY a.id
        `).all(startDate, endDate);
        for (const r of partRows) {
          const row = ensureCostRow(r.asset_id, r.asset_code, r.asset_name);
          if (!row) continue;
          row.parts_cost = Number(r.parts_cost || 0);
        }
      }

      const canReadMaintenanceLubes = hasTable("maintenance_records") && hasTable("maintenance_lubes");
      if (canReadMaintenanceParts) {
        const legacyPartRows = db.prepare(`
          SELECT
            mr.asset_id,
            COALESCE(SUM(COALESCE(mp.quantity, 0) * COALESCE(p.unit_cost, 0)), 0) AS parts_cost
          FROM maintenance_records mr
          JOIN maintenance_parts mp ON mp.maintenance_record_id = mr.id
          LEFT JOIN parts p ON LOWER(TRIM(p.part_name)) = LOWER(TRIM(mp.part_name))
          WHERE DATE(mr.maintenance_date) BETWEEN DATE(?) AND DATE(?)
          GROUP BY mr.asset_id
        `).all(startDate, endDate);
        for (const r of legacyPartRows) {
          const row = ensureCostRow(r.asset_id, null, null);
          if (!row) continue;
          row.parts_cost += Number(r.parts_cost || 0);
        }
      }

      if (hasTable("oil_logs")) {
        const hasOilUnit = hasColumn("oil_logs", "unit_cost");
        const lubeRows = db.prepare(`
          SELECT
            a.id AS asset_id,
            a.asset_code,
            a.asset_name,
            COALESCE(SUM(ol.quantity * COALESCE(${hasOilUnit ? "ol.unit_cost" : "NULL"}, ?)), 0) AS lube_cost
          FROM oil_logs ol
          JOIN assets a ON a.id = ol.asset_id
          WHERE ol.log_date BETWEEN DATE(?) AND DATE(?)
          GROUP BY a.id
        `).all(lubeDefault, startDate, endDate);
        for (const r of lubeRows) {
          const row = ensureCostRow(r.asset_id, r.asset_code, r.asset_name);
          if (!row) continue;
          row.lube_cost = Number(r.lube_cost || 0);
        }
      }

      if (canReadMaintenanceLubes) {
        const legacyLubeRows = db.prepare(`
          SELECT
            mr.asset_id,
            COALESCE(SUM(COALESCE(ml.quantity, 0) * COALESCE(p.unit_cost, 0)), 0) AS lube_cost
          FROM maintenance_records mr
          JOIN maintenance_lubes ml ON ml.maintenance_record_id = mr.id
          LEFT JOIN parts p ON LOWER(TRIM(p.part_name)) LIKE '%' || LOWER(TRIM(ml.lube_type)) || '%'
          WHERE DATE(mr.maintenance_date) BETWEEN DATE(?) AND DATE(?)
          GROUP BY mr.asset_id
        `).all(startDate, endDate);
        for (const r of legacyLubeRows) {
          const row = ensureCostRow(r.asset_id, null, null);
          if (!row) continue;
          row.lube_cost += Number(r.lube_cost || 0);
        }
      }

      const maintenanceCost = Array.from(costByAsset.values())
        .map((row) => {
          const woLabor = Number(row.wo_labor_cost || 0);
          const repairLabor = Number(row.repair_labor_cost || 0);
          const labor = woLabor + repairLabor;
          const parts = Number(row.parts_cost || 0);
          const lube = Number(row.lube_cost || 0);
          const outsourced = Number(row.outsourced_cost || 0);
          const total = labor + parts + lube + outsourced;
          return {
            ...row,
            labor_cost: Number(labor.toFixed(2)),
            wo_labor_cost: Number(woLabor.toFixed(2)),
            repair_labor_cost: Number(repairLabor.toFixed(2)),
            repair_labor_hours: Number(Number(row.repair_labor_hours || 0).toFixed(2)),
            parts_cost: Number(parts.toFixed(2)),
            lube_cost: Number(lube.toFixed(2)),
            outsourced_cost: Number(outsourced.toFixed(2)),
            total_cost: Number(total.toFixed(2)),
          };
        })
        .filter(
          (r) =>
            Number(r.total_cost || 0) > 0 ||
            Number(r.service_jobs || 0) > 0 ||
            Number(r.downtime_hours || 0) > 0
        )
        .sort((a, b) => Number(b.total_cost || 0) - Number(a.total_cost || 0));

      const downtimeTrend = insightsDowntime.downtime_daily;
      const woLaborTrend = hasTable("work_orders")
        ? db.prepare(`
            SELECT
              DATE(COALESCE(wo.completed_at, wo.closed_at, wo.opened_at, wo.updated_at)) AS day,
              COALESCE(SUM(COALESCE(wo.labor_hours, 0) * COALESCE(wo.labor_rate_per_hour, ?)), 0) AS labor_cost
            FROM work_orders wo
            WHERE DATE(COALESCE(wo.completed_at, wo.closed_at, wo.opened_at, wo.updated_at)) BETWEEN DATE(?) AND DATE(?)
            GROUP BY day
            ORDER BY day ASC
          `).all(laborRate, startDate, endDate)
        : [];
      const costTrend = (woLaborTrend || []).map((r) => ({
        day: String(r.day || ""),
        labor_cost: Number(Number(r.labor_cost || 0).toFixed(2)),
      }));

      return reply.send({
        ok: true,
        range: {
          start: startDate,
          end: endDate,
          near_due_hours: nearDueHours,
          predictive_horizon_hours: predictiveHorizonHours,
          checklist_fail_threshold: checklistFailThreshold,
          fuel_variance_threshold: fuelVarianceThreshold,
          labor_cost_per_hour: laborRate,
        },
        predictive,
        parts_planning: partsPlanning,
        downtime,
        sla,
        maintenance_cost: maintenanceCost,
        trends: {
          downtime_daily: downtimeTrend,
          labor_daily: costTrend,
        },
      });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // POST /api/maintenance/breakdowns/:id/repair-labor
  // Body: { labor_hours, notes? } — actual technician hours (not machine downtime hours)
  app.post("/breakdowns/:id/repair-labor", async (req, reply) => {
    try {
      ensureBreakdownRepairLaborSchema(db);
      const breakdownId = Number(req.params?.id || 0);
      const labor_hours = Math.max(0, Number(req.body?.labor_hours || 0));
      const notes = String(req.body?.notes || "").trim() || null;
      if (!breakdownId) return reply.code(400).send({ ok: false, error: "breakdown id is required" });

      const breakdown = db.prepare(`
        SELECT b.id, b.asset_id, a.asset_code, a.asset_name
        FROM breakdowns b
        JOIN assets a ON a.id = b.asset_id
        WHERE b.id = ?
        LIMIT 1
      `).get(breakdownId);
      if (!breakdown) return reply.code(404).send({ ok: false, error: "breakdown not found" });

      db.prepare(`
        INSERT INTO breakdown_repair_labor (breakdown_id, labor_hours, notes, updated_at)
        VALUES (?, ?, ?, datetime('now'))
        ON CONFLICT(breakdown_id) DO UPDATE SET
          labor_hours = excluded.labor_hours,
          notes = excluded.notes,
          updated_at = datetime('now')
      `).run(breakdownId, labor_hours, notes);

      return reply.send({
        ok: true,
        breakdown_id: breakdownId,
        asset_code: breakdown.asset_code,
        labor_hours: Number(labor_hours.toFixed(2)),
        notes,
      });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // =====================================================
  // GOVERNANCE SIGNALS (Data quality + anomaly detection)
  // GET /api/maintenance/governance/signals?start=YYYY-MM-DD&end=YYYY-MM-DD&meter_jump_threshold=500
  // =====================================================
  app.get("/governance/signals", async (req, reply) => {
    try {
      const endDate = String(req.query?.end || "").trim() || new Date().toISOString().slice(0, 10);
      if (!isDate(endDate)) return reply.code(400).send({ ok: false, error: "end must be YYYY-MM-DD" });
      const startDate = String(req.query?.start || "").trim() || (() => {
        const d = new Date(`${endDate}T00:00:00`);
        d.setDate(d.getDate() - 29);
        return d.toISOString().slice(0, 10);
      })();
      if (!isDate(startDate)) return reply.code(400).send({ ok: false, error: "start must be YYYY-MM-DD" });
      const meterJumpThreshold = Math.max(50, Number(req.query?.meter_jump_threshold || 500));

      const activeAssets = db.prepare(`
        SELECT id, asset_code, asset_name
        FROM assets
        WHERE COALESCE(active, 1) = 1
      `).all();

      const missingMeterReadings = activeAssets
        .map((a) => {
          const info = getAssetCurrentHoursInfo(Number(a.id || 0));
          return {
            asset_id: Number(a.id || 0),
            asset_code: String(a.asset_code || ""),
            asset_name: String(a.asset_name || ""),
            current_hours: Number(info.hours || 0),
            source: String(info.source || ""),
          };
        })
        .filter((r) => Number(r.current_hours || 0) <= 0)
        .slice(0, 50);

      const inconsistentStatuses = hasTable("work_orders")
        ? db.prepare(`
            SELECT
              w.id AS work_order_id,
              a.asset_code,
              a.asset_name,
              w.status,
              w.opened_at,
              w.completed_at,
              w.closed_at
            FROM work_orders w
            LEFT JOIN assets a ON a.id = w.asset_id
            WHERE
              (LOWER(COALESCE(w.status, '')) IN ('completed','approved','closed') AND w.completed_at IS NULL)
              OR (LOWER(COALESCE(w.status, '')) = 'closed' AND w.closed_at IS NULL)
            ORDER BY w.id DESC
            LIMIT 100
          `).all()
        : [];

      const stalePlans = hasTable("maintenance_plans")
        ? db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              a.asset_code,
              a.asset_name,
              mp.service_name,
              mp.last_service_hours,
              mp.interval_hours
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE COALESCE(mp.active, 1) = 1
          `).all()
            .map((p) => {
              const current = Number(getAssetCurrentHoursInfo(Number(p.asset_id || 0)).hours || 0);
              const nextDue = Number(p.last_service_hours || 0) + Number(p.interval_hours || 0);
              const remaining = Number((nextDue - current).toFixed(2));
              return { ...p, remaining_hours: remaining };
            })
            .filter((r) => Number(r.remaining_hours || 0) <= -500)
            .sort((a, b) => Number(a.remaining_hours || 0) - Number(b.remaining_hours || 0))
            .slice(0, 50)
        : [];

      const fuelDuplicates = hasTable("fuel_logs")
        ? db.prepare(`
            SELECT
              fl.asset_id,
              a.asset_code,
              a.asset_name,
              fl.log_date,
              ROUND(COALESCE(fl.liters, 0), 2) AS liters,
              COUNT(*) AS duplicate_count
            FROM fuel_logs fl
            LEFT JOIN assets a ON a.id = fl.asset_id
            WHERE fl.log_date BETWEEN ? AND ?
            GROUP BY fl.asset_id, fl.log_date, ROUND(COALESCE(fl.liters, 0), 2)
            HAVING COUNT(*) >= 2
            ORDER BY duplicate_count DESC, fl.log_date DESC
            LIMIT 100
          `).all(startDate, endDate)
        : [];

      const fuelSpikes = hasTable("fuel_logs")
        ? db.prepare(`
            SELECT
              fl.asset_id,
              a.asset_code,
              a.asset_name,
              fl.log_date,
              COALESCE(fl.liters, 0) AS liters,
              (
                SELECT AVG(COALESCE(f2.liters, 0))
                FROM fuel_logs f2
                WHERE f2.asset_id = fl.asset_id
                  AND f2.log_date BETWEEN ? AND ?
              ) AS avg_liters
            FROM fuel_logs fl
            LEFT JOIN assets a ON a.id = fl.asset_id
            WHERE fl.log_date BETWEEN ? AND ?
            ORDER BY fl.log_date DESC
          `).all(startDate, endDate, startDate, endDate)
            .map((r) => {
              const liters = Number(r.liters || 0);
              const avg = Number(r.avg_liters || 0);
              const ratio = avg > 0 ? liters / avg : 0;
              return {
                asset_id: Number(r.asset_id || 0),
                asset_code: String(r.asset_code || ""),
                asset_name: String(r.asset_name || ""),
                log_date: String(r.log_date || ""),
                liters: Number(liters.toFixed(2)),
                avg_liters: Number(avg.toFixed(2)),
                spike_ratio: Number(ratio.toFixed(2)),
              };
            })
            .filter((r) => Number(r.avg_liters || 0) > 0 && Number(r.spike_ratio || 0) >= 1.8)
            .sort((a, b) => Number(b.spike_ratio || 0) - Number(a.spike_ratio || 0))
            .slice(0, 50)
        : [];

      const meterJumps = hasTable("daily_inputs")
        ? db.prepare(`
            SELECT
              di.asset_id,
              a.asset_code,
              a.asset_name,
              di.input_date,
              COALESCE(di.hour_meter_closing, 0) AS hour_meter_closing
            FROM daily_inputs di
            LEFT JOIN assets a ON a.id = di.asset_id
            WHERE di.input_date BETWEEN ? AND ?
            ORDER BY di.asset_id, di.input_date
          `).all(startDate, endDate)
            .reduce((acc, r) => {
              const aid = Number(r.asset_id || 0);
              if (!aid) return acc;
              const prev = acc._lastByAsset.get(aid);
              const cur = Number(r.hour_meter_closing || 0);
              if (prev && Number.isFinite(prev.value) && Number.isFinite(cur)) {
                const jump = cur - prev.value;
                if (Math.abs(jump) >= meterJumpThreshold) {
                  acc.rows.push({
                    asset_id: aid,
                    asset_code: String(r.asset_code || ""),
                    asset_name: String(r.asset_name || ""),
                    input_date: String(r.input_date || ""),
                    previous_meter: Number(prev.value.toFixed(1)),
                    current_meter: Number(cur.toFixed(1)),
                    jump: Number(jump.toFixed(1)),
                  });
                }
              }
              acc._lastByAsset.set(aid, { value: cur, date: String(r.input_date || "") });
              return acc;
            }, { rows: [], _lastByAsset: new Map() }).rows.slice(0, 100)
        : [];

      return reply.send({
        ok: true,
        range: { start: startDate, end: endDate, meter_jump_threshold: meterJumpThreshold },
        quality: {
          missing_meter_readings: missingMeterReadings,
          inconsistent_statuses: inconsistentStatuses,
          stale_plans: stalePlans,
        },
        anomalies: {
          fuel_spikes: fuelSpikes,
          fuel_duplicates: fuelDuplicates,
          suspicious_meter_jumps: meterJumps,
        },
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

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

  // GET /api/maintenance/insights.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD&near_due_hours=50
  app.get("/insights.xlsx", async (req, reply) => {
    try {
      const data = await loadInsightsExportPayload(req);
      const wb = new ExcelJS.Workbook();
      wb.creator = "IRONLOG";
      wb.created = new Date();

      const wsSummary = wb.addWorksheet("Summary");
      wsSummary.columns = [{ header: "Field", key: "field", width: 34 }, { header: "Value", key: "value", width: 30 }];
      wsSummary.addRows([
        { field: "Start", value: data?.range?.start || "" },
        { field: "End", value: data?.range?.end || "" },
        { field: "Near Due Hours", value: Number(data?.range?.near_due_hours || 0) },
        { field: "Predictive Horizon Hours", value: Number(data?.range?.predictive_horizon_hours || 0) },
        { field: "Checklist Fail Threshold", value: Number(data?.range?.checklist_fail_threshold || 0) },
        { field: "Fuel Variance Threshold (%)", value: Number(data?.range?.fuel_variance_threshold || 0) },
      ]);

      const wsPredictive = wb.addWorksheet("Predictive Alerts");
      wsPredictive.columns = [
        { header: "Asset Code", key: "asset_code", width: 14 },
        { header: "Asset Name", key: "asset_name", width: 28 },
        { header: "Service", key: "service_name", width: 24 },
        { header: "Remaining Hrs", key: "remaining_hours", width: 14 },
        { header: "Risk", key: "risk", width: 12 },
      ];
      wsPredictive.addRows(Array.isArray(data?.predictive?.at_risk_plans) ? data.predictive.at_risk_plans : []);

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
        })),
      );

      const wsDowntimeComp = wb.addWorksheet("Downtime Components");
      wsDowntimeComp.columns = [
        { header: "Component", key: "component", width: 24 },
        { header: "Incidents", key: "incidents", width: 12 },
        { header: "Downtime Hrs", key: "downtime_hours", width: 14 },
      ];
      wsDowntimeComp.addRows(Array.isArray(data?.downtime?.by_component) ? data.downtime.by_component : []);

      const wsDowntimeTeam = wb.addWorksheet("Downtime Teams");
      wsDowntimeTeam.columns = [
        { header: "Team", key: "team", width: 24 },
        { header: "Incidents", key: "incidents", width: 12 },
        { header: "Downtime Hrs", key: "downtime_hours", width: 14 },
      ];
      wsDowntimeTeam.addRows(Array.isArray(data?.downtime?.by_team) ? data.downtime.by_team : []);

      const wsSla = wb.addWorksheet("SLA");
      wsSla.columns = [{ header: "Metric", key: "metric", width: 36 }, { header: "Hours", key: "hours", width: 14 }];
      wsSla.addRows([
        { metric: "Open -> Assign", hours: data?.sla?.avg_open_to_assign_hours ?? "" },
        { metric: "Open -> Complete", hours: data?.sla?.avg_open_to_complete_hours ?? "" },
        { metric: "Complete -> Approve", hours: data?.sla?.avg_complete_to_approve_hours ?? "" },
        { metric: "Open -> Close", hours: data?.sla?.avg_open_to_close_hours ?? "" },
        { metric: "Service Work Orders", hours: Number(data?.sla?.work_orders || 0) },
      ]);

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
        { header: "Total Cost", key: "total_cost", width: 12 },
      ];
      wsCost.addRows(Array.isArray(data?.maintenance_cost) ? data.maintenance_cost : []);

      const wsRepairLabor = wb.addWorksheet("Breakdown Repair Labor");
      wsRepairLabor.columns = [
        { header: "Asset Code", key: "asset_code", width: 14 },
        { header: "Asset Name", key: "asset_name", width: 24 },
        { header: "Breakdown ID", key: "breakdown_id", width: 12 },
        { header: "Status", key: "status", width: 12 },
        { header: "Down Hrs", key: "downtime_hours", width: 12 },
        { header: "Repair Labor Hrs", key: "actual_labor_hours", width: 16 },
        { header: "Labor Rate", key: "labor_rate", width: 12 },
        { header: "Repair Labor $", key: "repair_labor_cost", width: 14 },
        { header: "Source", key: "labor_source", width: 12 },
        { header: "Needs Input", key: "needs_labor_input", width: 12 },
        { header: "Notes", key: "labor_notes", width: 28 },
      ];
      wsRepairLabor.addRows(
        Array.isArray(data?.downtime?.labor?.incidents)
          ? data.downtime.labor.incidents.map((r) => ({
              ...r,
              needs_labor_input: r.needs_labor_input ? "yes" : "no",
            }))
          : [],
      );

      const wsTrends = wb.addWorksheet("Trends");
      wsTrends.columns = [
        { header: "Day", key: "day", width: 14 },
        { header: "Downtime Hrs", key: "downtime_hours", width: 14 },
        { header: "Labor Cost", key: "labor_cost", width: 14 },
      ];
      const downtimeDaily = new Map((Array.isArray(data?.trends?.downtime_daily) ? data.trends.downtime_daily : []).map((r) => [String(r.day || ""), Number(r.downtime_hours || 0)]));
      const laborDaily = new Map((Array.isArray(data?.trends?.labor_daily) ? data.trends.labor_daily : []).map((r) => [String(r.day || ""), Number(r.labor_cost || 0)]));
      const daySet = new Set([...downtimeDaily.keys(), ...laborDaily.keys()]);
      const days = [...daySet].filter(Boolean).sort();
      wsTrends.addRows(days.map((day) => ({
        day,
        downtime_hours: Number((downtimeDaily.get(day) || 0).toFixed(2)),
        labor_cost: Number((laborDaily.get(day) || 0).toFixed(2)),
      })));

      const buffer = await wb.xlsx.writeBuffer();
      return reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Content-Disposition", `attachment; filename="IRONLOG_Maintenance_Insights_${data?.range?.start || "start"}_to_${data?.range?.end || "end"}.xlsx"`)
        .send(buffer);
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

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

  // GET /api/maintenance/insights/parts-demand.xlsx
  app.get("/insights/parts-demand.xlsx", async (req, reply) => {
    try {
      const data = await loadInsightsExportPayload(req);
      const wb = new ExcelJS.Workbook();
      wb.creator = "IRONLOG";
      wb.created = new Date();
      addInsightsExportSummarySheet(wb, data);
      addInsightsPartsDemandSheets(wb, data);
      const buffer = await wb.xlsx.writeBuffer();
      return reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header(
          "Content-Disposition",
          `attachment; filename="IRONLOG_Parts_Demand_${data?.range?.start || "start"}_to_${data?.range?.end || "end"}.xlsx"`,
        )
        .send(buffer);
    } catch (e) {
      req.log.error(e);
      return reply.code(Number(e.statusCode || 500)).send({ ok: false, error: e.message || String(e) });
    }
  });

  // GET /api/maintenance/insights/cost-per-machine.xlsx
  app.get("/insights/cost-per-machine.xlsx", async (req, reply) => {
    try {
      const data = await loadInsightsExportPayload(req);
      const wb = new ExcelJS.Workbook();
      wb.creator = "IRONLOG";
      wb.created = new Date();
      addInsightsExportSummarySheet(wb, data);
      addInsightsCostPerMachineSheet(wb, data);
      const buffer = await wb.xlsx.writeBuffer();
      return reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header(
          "Content-Disposition",
          `attachment; filename="IRONLOG_Maintenance_Cost_Per_Machine_${data?.range?.start || "start"}_to_${data?.range?.end || "end"}.xlsx"`,
        )
        .send(buffer);
    } catch (e) {
      req.log.error(e);
      return reply.code(Number(e.statusCode || 500)).send({ ok: false, error: e.message || String(e) });
    }
  });

  // GET /api/maintenance/insights.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&near_due_hours=50&download=1
  app.get("/insights.pdf", async (req, reply) => {
    try {
      const download = String(req.query?.download || "").trim() === "1";
      const data = await loadInsightsExportPayload(req);
      const start = String(data?.range?.start || "");
      const end = String(data?.range?.end || "");

      const pdf = await buildPdfBuffer((doc) => {
        sectionTitle(doc, "Maintenance Insights Summary");
        table(
          doc,
          ["Metric", "Value"],
          [
            { Metric: "Period", Value: `${start} to ${end}` },
            { Metric: "Near Due Hours", Value: String(data?.range?.near_due_hours || "-") },
            { Metric: "Predictive Horizon Hours", Value: String(data?.range?.predictive_horizon_hours || "-") },
            { Metric: "Checklist Fail Threshold", Value: String(data?.range?.checklist_fail_threshold || "-") },
            { Metric: "Fuel Variance Threshold (%)", Value: String(data?.range?.fuel_variance_threshold || "-") },
          ],
          [0.45, 0.55]
        );

        sectionTitle(doc, "Predictive At-Risk Plans");
        const risk = Array.isArray(data?.predictive?.at_risk_plans) ? data.predictive.at_risk_plans : [];
        table(
          doc,
          ["Asset", "Service", "Remaining Hrs", "Risk"],
          risk.length
            ? risk.slice(0, 20).map((r) => ({
                Asset: `${String(r.asset_code || "-")} - ${String(r.asset_name || "-")}`,
                Service: String(r.service_name || "-"),
                "Remaining Hrs": Number(r.remaining_hours || 0).toFixed(1),
                Risk: String(r.risk || "-"),
              }))
            : [{ Asset: "-", Service: "No at-risk plans", "Remaining Hrs": "-", Risk: "-" }],
          [0.36, 0.28, 0.18, 0.18]
        );

        sectionTitle(doc, "SLA");
        table(
          doc,
          ["Metric", "Hours"],
          [
            { Metric: "Open -> Assign", Hours: data?.sla?.avg_open_to_assign_hours ?? "-" },
            { Metric: "Open -> Complete", Hours: data?.sla?.avg_open_to_complete_hours ?? "-" },
            { Metric: "Complete -> Approve", Hours: data?.sla?.avg_complete_to_approve_hours ?? "-" },
            { Metric: "Open -> Close", Hours: data?.sla?.avg_open_to_close_hours ?? "-" },
            { Metric: "Service Work Orders", Hours: Number(data?.sla?.work_orders || 0) },
          ],
          [0.65, 0.35]
        );

        sectionTitle(doc, "Breakdown Downtime vs Repair Labor");
        const laborIncidents = Array.isArray(data?.downtime?.labor?.incidents) ? data.downtime.labor.incidents : [];
        const laborTotals = data?.downtime?.labor?.totals || {};
        table(
          doc,
          ["Metric", "Value"],
          [
            { Metric: "Machine downtime hours", Value: Number(laborTotals.downtime_hours || 0).toFixed(2) },
            { Metric: "Actual repair labor hours", Value: Number(laborTotals.repair_labor_hours || 0).toFixed(2) },
            { Metric: "Repair labor cost", Value: Number(laborTotals.repair_labor_cost || 0).toFixed(2) },
            { Metric: "Incidents needing labor input", Value: String(laborTotals.needs_input_count || 0) },
          ],
          [0.55, 0.45]
        );
        table(
          doc,
          ["Asset", "Down Hrs", "Repair Hrs", "Repair $", "Source"],
          laborIncidents.length
            ? laborIncidents.slice(0, 20).map((r) => ({
                Asset: `${String(r.asset_code || "-")} - ${String(r.asset_name || "-")}`,
                "Down Hrs": Number(r.downtime_hours || 0).toFixed(2),
                "Repair Hrs": Number(r.actual_labor_hours || 0).toFixed(2),
                "Repair $": Number(r.repair_labor_cost || 0).toFixed(2),
                Source: String(r.labor_source || "-"),
              }))
            : [{ Asset: "No breakdown labor in period", "Down Hrs": "-", "Repair Hrs": "-", "Repair $": "-", Source: "-" }],
          [0.38, 0.14, 0.14, 0.14, 0.2]
        );

        sectionTitle(doc, "Top Maintenance Cost Per Machine");
        const costs = Array.isArray(data?.maintenance_cost) ? data.maintenance_cost : [];
        table(
          doc,
          ["Asset", "Jobs", "Down Hrs", "Repair Hrs", "Labor", "Parts", "Total"],
          costs.length
            ? costs.slice(0, 20).map((r) => ({
                Asset: `${String(r.asset_code || "-")} - ${String(r.asset_name || "-")}`,
                Jobs: Number(r.service_jobs || 0),
                "Down Hrs": Number(r.downtime_hours || 0).toFixed(2),
                "Repair Hrs": Number(r.repair_labor_hours || 0).toFixed(2),
                Labor: Number(r.labor_cost || 0).toFixed(2),
                Parts: Number(r.parts_cost || 0).toFixed(2),
                Total: Number(r.total_cost || 0).toFixed(2),
              }))
            : [{ Asset: "No costs in period", Jobs: "-", "Down Hrs": "-", "Repair Hrs": "-", Labor: "-", Parts: "-", Total: "-" }],
          [0.32, 0.08, 0.1, 0.1, 0.13, 0.13, 0.14]
        );
      }, {
        title: "IRONLOG",
        subtitle: "Maintenance Insights",
        rightText: `${start} to ${end}`,
        showPageNumbers: true,
      });

      return reply
        .header("Content-Type", "application/pdf")
        .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename="IRONLOG_Maintenance_Insights_${start}_to_${end}.pdf"`)
        .send(pdf);
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // =====================================================
  // BACKFILL (ANCIENT) SERVICE HISTORY
  // POST /api/maintenance/history/backfill
  // Body: { asset_id, service_name, service_date, service_hours?, notes?, update_plan_last_hours?, plan_id? }
  // =====================================================
  app.post("/history/backfill", async (req, reply) => {
    try {
      const body = req.body || {};
      const assetId = Number(body.asset_id || 0);
      const serviceName = String(body.service_name || "").trim();
      const serviceDate = String(body.service_date || "").trim();
      const serviceHoursIn = body.service_hours;
      const notes = String(body.notes || "").trim() || null;
      const updatePlanLastHours = Number(body.update_plan_last_hours || 0) === 1;
      const planIdIn = Number(body.plan_id || 0);

      if (!assetId) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!serviceName) return reply.code(400).send({ ok: false, error: "service_name is required" });
      if (!isDate(serviceDate)) return reply.code(400).send({ ok: false, error: "service_date must be YYYY-MM-DD" });

      const serviceHours = serviceHoursIn == null || String(serviceHoursIn).trim() === ""
        ? null
        : Number(serviceHoursIn);
      if (serviceHours != null && (!Number.isFinite(serviceHours) || serviceHours < 0)) {
        return reply.code(400).send({ ok: false, error: "service_hours must be a valid number >= 0" });
      }

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name
        FROM assets
        WHERE id = ?
        LIMIT 1
      `).get(assetId);
      if (!asset) return reply.code(404).send({ ok: false, error: "asset not found" });

      let planId = planIdIn > 0 ? planIdIn : null;
      let planInterval = 0;
      if (!planId) {
        const matchedPlan = db.prepare(`
          SELECT id, interval_hours
          FROM maintenance_plans
          WHERE asset_id = ?
            AND UPPER(TRIM(service_name)) = UPPER(TRIM(?))
          ORDER BY active DESC, id DESC
          LIMIT 1
        `).get(assetId, serviceName);
        if (matchedPlan?.id) {
          planId = Number(matchedPlan.id);
          planInterval = Number(matchedPlan.interval_hours || 0);
        }
      } else {
        const matchedPlan = db.prepare(`
          SELECT interval_hours FROM maintenance_plans WHERE id = ?
        `).get(planId);
        planInterval = Number(matchedPlan?.interval_hours || 0);
      }
      if (!planInterval) planInterval = planIntervalHours({ service_name: serviceName, interval_hours: 0 });

      const snappedServiceHours = serviceHours != null
        ? snapLastServiceHours(serviceHours, planInterval, asset.asset_code)
        : null;

      const insert = db.prepare(`
        INSERT INTO maintenance_service_history (
          asset_id, plan_id, service_name, service_date, service_hours, notes, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
      `);
      const updatePlan = db.prepare(`
        UPDATE maintenance_plans
        SET last_service_hours = ?
        WHERE id = ?
      `);

      const tx = db.transaction(() => {
        const r = insert.run(
          assetId,
          planId,
          serviceName,
          serviceDate,
          snappedServiceHours,
          notes,
          String(req.headers?.["x-user-name"] || "system")
        );
        if (updatePlanLastHours && planId && snappedServiceHours != null) {
          updatePlan.run(snappedServiceHours, Number(planId));
          db.prepare(`
            UPDATE maintenance_plans SET last_service_hours = ?
            WHERE asset_id = ? AND active = 1
          `).run(snappedServiceHours, assetId);
        }
        return Number(r.lastInsertRowid || 0);
      });

      const id = tx();
      writeAudit(db, req, {
        module: "maintenance",
        action: "service_history.create",
        entity_type: "maintenance_service_history",
        entity_id: String(id),
        after: {
          asset_id: assetId,
          plan_id: planId || null,
          service_name: serviceName,
          service_date: serviceDate,
          service_hours: snappedServiceHours,
        },
      });
      return reply.send({
        ok: true,
        id,
        asset_id: assetId,
        asset_code: asset.asset_code,
        service_name: serviceName,
        service_date: serviceDate,
        service_hours: snappedServiceHours,
        plan_id: planId || null,
        plan_last_hours_updated: Boolean(updatePlanLastHours && planId && snappedServiceHours != null),
      });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // GET /api/maintenance/history/backfill?asset_id=&asset_code=&limit=200&include_work_orders=1
  app.get("/history/backfill", async (req, reply) => {
    try {
      let assetId = Number(req.query?.asset_id || 0);
      const assetCode = String(req.query?.asset_code || "").trim();
      if (!assetId && assetCode) {
        const a = db.prepare(`
          SELECT id FROM assets WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?)) LIMIT 1
        `).get(assetCode);
        assetId = Number(a?.id || 0);
      }
      const limit = Math.max(1, Math.min(500, Number(req.query?.limit || 200)));
      const includeWorkOrders = String(req.query?.include_work_orders || "1").trim() !== "0";
      const where = assetId > 0 ? "WHERE h.asset_id = ?" : "";
      const backfillRows = db.prepare(`
        SELECT
          h.id,
          h.asset_id,
          h.plan_id,
          h.service_name,
          h.service_date,
          h.service_hours,
          h.notes,
          h.created_by,
          h.created_at,
          a.asset_code,
          a.asset_name,
          'backfill' AS record_source
        FROM maintenance_service_history h
        JOIN assets a ON a.id = h.asset_id
        ${where}
        ORDER BY h.service_date DESC, h.id DESC
        LIMIT ?
      `).all(...(assetId > 0 ? [assetId, limit] : [limit]));

      let woRows = [];
      if (includeWorkOrders && hasTable("work_orders")) {
        const woWhere = assetId > 0 ? "AND w.asset_id = ?" : "";
        woRows = db.prepare(`
          SELECT
            w.id,
            w.asset_id,
            w.reference_id AS plan_id,
            COALESCE(mp.service_name, 'Service') AS service_name,
            DATE(COALESCE(w.closed_at, w.completed_at)) AS service_date,
            NULL AS service_hours,
            w.completion_notes AS notes,
            w.artisan_name AS created_by,
            COALESCE(w.closed_at, w.completed_at) AS created_at,
            a.asset_code,
            a.asset_name,
            'work_order' AS record_source
          FROM work_orders w
          JOIN assets a ON a.id = w.asset_id
          LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id
          WHERE LOWER(COALESCE(w.source, '')) = 'service'
            AND LOWER(COALESCE(w.status, '')) IN ('completed', 'approved', 'closed')
            ${woWhere}
          ORDER BY COALESCE(w.closed_at, w.completed_at) DESC, w.id DESC
          LIMIT ?
        `).all(...(assetId > 0 ? [assetId, limit] : [limit]));
      }

      const merged = [...backfillRows, ...woRows]
        .sort((a, b) => {
          const da = String(a.service_date || a.created_at || "");
          const dbd = String(b.service_date || b.created_at || "");
          return dbd.localeCompare(da) || Number(b.id || 0) - Number(a.id || 0);
        })
        .slice(0, limit);

      return reply.send({ ok: true, rows: merged, asset_id: assetId || null });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // PUT /api/maintenance/history/backfill/:id
  app.put("/history/backfill/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
      const body = req.body || {};
      const serviceName = String(body.service_name || "").trim();
      const serviceDate = String(body.service_date || "").trim();
      const notes = String(body.notes || "").trim() || null;
      const serviceHoursIn = body.service_hours;
      const serviceHours = serviceHoursIn == null || String(serviceHoursIn).trim() === ""
        ? null
        : Number(serviceHoursIn);
      if (!serviceName) return reply.code(400).send({ ok: false, error: "service_name is required" });
      if (!isDate(serviceDate)) return reply.code(400).send({ ok: false, error: "service_date must be YYYY-MM-DD" });
      if (serviceHours != null && (!Number.isFinite(serviceHours) || serviceHours < 0)) {
        return reply.code(400).send({ ok: false, error: "service_hours must be a valid number >= 0" });
      }
      const cur = db.prepare(`SELECT * FROM maintenance_service_history WHERE id = ?`).get(id);
      if (!cur) return reply.code(404).send({ ok: false, error: "backfill entry not found" });
      db.prepare(`
        UPDATE maintenance_service_history
        SET service_name = ?, service_date = ?, service_hours = ?, notes = ?
        WHERE id = ?
      `).run(serviceName, serviceDate, serviceHours, notes, id);
      writeAudit(db, req, {
        module: "maintenance",
        action: "service_history.update",
        entity_type: "maintenance_service_history",
        entity_id: String(id),
        before: cur,
        after: { ...cur, service_name: serviceName, service_date: serviceDate, service_hours: serviceHours, notes },
      });
      return reply.send({ ok: true, id });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // DELETE /api/maintenance/history/backfill/:id
  app.delete("/history/backfill/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
      const cur = db.prepare(`SELECT * FROM maintenance_service_history WHERE id = ?`).get(id);
      if (!cur) return reply.code(404).send({ ok: false, error: "backfill entry not found" });
      db.prepare(`DELETE FROM maintenance_service_history WHERE id = ?`).run(id);
      writeAudit(db, req, {
        module: "maintenance",
        action: "service_history.delete",
        entity_type: "maintenance_service_history",
        entity_id: String(id),
        before: cur,
      });
      return reply.send({ ok: true, id });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // =====================================================
  // AUTO-GENERATE SERVICE WORK ORDERS
  // POST /api/maintenance/generate?date=2026-02-27
  // =====================================================
  app.post("/generate", async (req, reply) => {
    try {
      const date = String(req.query?.date || "").trim();
      if (date && !isDate(date)) {
        return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
      }
      const nearDueHours = Math.max(1, Number(req.body?.near_due_hours || req.query?.near_due_hours || 50));
      const planIdsRaw = Array.isArray(req.body?.plan_ids) ? req.body.plan_ids : [];
      const requestedPlanIds = [...new Set(planIdsRaw.map((v) => Number(v)).filter((v) => Number.isInteger(v) && v > 0))];

      const plans = date
        ? db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              a.asset_code,
              a.asset_name,
              a.category,
              IFNULL((
                SELECT SUM(dh.hours_run)
                FROM daily_hours dh
                WHERE dh.asset_id = mp.asset_id
                  AND dh.is_used = 1
                  AND dh.hours_run > 0
                  AND dh.work_date <= ?
              ), 0) AS current_hours
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.is_standby = 0
              AND a.archived = 0
          `).all(date)
        : db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              a.asset_code,
              a.asset_name,
              a.category,
              IFNULL((
                SELECT SUM(dh.hours_run)
                FROM daily_hours dh
                WHERE dh.asset_id = mp.asset_id
                  AND dh.is_used = 1
                  AND dh.hours_run > 0
              ), 0) AS current_hours
            FROM maintenance_plans mp
            JOIN assets a ON a.id = mp.asset_id
            WHERE mp.active = 1
              AND a.active = 1
              AND a.is_standby = 0
              AND a.archived = 0
          `).all();

      const hasOpenServiceWO = db.prepare(`
        SELECT 1
        FROM work_orders
        WHERE source = 'service'
          AND reference_id = ?
          AND status != 'closed'
        LIMIT 1
      `);

      const insertWO = db.prepare(`
        INSERT INTO work_orders (asset_id, source, reference_id, status)
        VALUES (?, 'service', ?, 'open')
      `);

      const currentByAsset = new Map();
      for (const plan of plans) {
        const assetId = Number(plan.asset_id || 0);
        if (!currentByAsset.has(assetId)) {
          currentByAsset.set(assetId, date ? Number(plan.current_hours || 0) : getAssetCurrentHours(assetId));
        }
      }
      const nextServices = buildDueListFromPlans(
        plans,
        (assetId) => Number(currentByAsset.get(Number(assetId)) || 0),
        nearDueHours,
      );

      const tx = db.transaction(() => {
        const created = [];

        for (const p of nextServices) {
          const current = Number(p.current_hours || 0);
          const next_due = Number(p.next_due_hours || 0);
          const remaining = Number(p.remaining_hours || 0);
          const isOverdue = remaining <= 0;
          const isRequestedPlan = requestedPlanIds.includes(Number(p.plan_id || 0));
          const shouldCreate = requestedPlanIds.length ? isRequestedPlan : isOverdue;
          if (!shouldCreate) continue;
          if (hasOpenServiceWO.get(p.plan_id)) continue;

          const wo = insertWO.run(p.asset_id, p.plan_id);
          created.push({
            work_order_id: Number(wo.lastInsertRowid),
            plan_id: p.plan_id,
            asset_id: p.asset_id,
            service_name: p.service_name,
            current_hours: Number(current.toFixed(2)),
            next_due_hours: Number(next_due.toFixed(2))
          });
        }

        return created;
      });

      const created = tx();

      return reply.send({
        ok: true,
        near_due_hours: nearDueHours,
        requested_plan_count: requestedPlanIds.length,
        created_count: created.length,
        created
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({
        ok: false,
        error: err.message
      });
    }
  });

  // =====================================================
  // UPCOMING SERVICES PDF
  // GET /api/maintenance/due-upcoming.pdf?date=YYYY-MM-DD&near_due_hours=50&download=1
  // =====================================================
  app.get("/due-upcoming.pdf", async (req, reply) => {
    try {
      const date = String(req.query?.date || "").trim();
      if (date && !isDate(date)) {
        return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
      }
      const nearDueHours = Math.max(1, Number(req.query?.near_due_hours || 50));
      const withinHoursRaw = req.query?.within_hours;
      const withinHours = withinHoursRaw != null && String(withinHoursRaw).trim() !== ""
        ? Math.max(0, Number(withinHoursRaw))
        : null;
      const planIdsRaw = req.query?.plan_ids;
      const requestedPlanIds = Array.isArray(planIdsRaw)
        ? [...new Set(planIdsRaw.map((v) => Number(v)).filter((v) => Number.isInteger(v) && v > 0))]
        : [...new Set(String(planIdsRaw || "").split(/[,\s]+/).map((v) => Number(v)).filter((v) => Number.isInteger(v) && v > 0))];
      const asOfLabel = date || new Date().toISOString().slice(0, 10);
      const historyDays = 14;

      const rows = date
        ? db.prepare(`
            SELECT
              mp.id AS plan_id,
              mp.asset_id,
              mp.service_name,
              mp.interval_hours,
              mp.last_service_hours,
              a.asset_code,
              a.asset_name,
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
              AND a.archived = 0
            ORDER BY a.asset_code ASC, mp.service_name ASC
          `).all(date)
        : db.prepare(`
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
              AND a.is_standby = 0
              AND a.archived = 0
            ORDER BY a.asset_code ASC, mp.service_name ASC
          `).all();

      const getHours = (assetId) => {
        if (date) {
          return Number(rows.find((r) => Number(r.asset_id) === Number(assetId))?.current_hours ?? getAssetCurrentHours(assetId));
        }
        return getAssetCurrentHours(assetId);
      };

      const dueRows = enrichDueRowsWithEstimates(
        db,
        buildDueListFromPlans(rows, getHours, nearDueHours),
        { as_of: asOfLabel, history_days: historyDays },
      )
        .map((r) => ({
          plan_id: Number(r.plan_id || 0),
          asset_code: r.asset_code,
          asset_name: r.asset_name,
          service_name: r.service_name,
          current_hours: Number(r.current_hours || 0),
          next_due_hours: Number(r.next_due_hours || 0),
          remaining_hours: Number(r.remaining_hours || 0),
          estimated_service_date: r.estimated_service_date,
          status: r.status,
        }))
        .filter((r) => {
          if (requestedPlanIds.length) return requestedPlanIds.includes(Number(r.plan_id || 0));
          if (withinHours != null && Number.isFinite(withinHours)) {
            return Number(r.remaining_hours || 0) <= withinHours;
          }
          return true;
        })
        .sort((a, b) => Number(a.remaining_hours || 0) - Number(b.remaining_hours || 0));

      const pdfScopeLabel = requestedPlanIds.length
        ? `Selected equipment only (${requestedPlanIds.length})`
        : withinHours != null && Number.isFinite(withinHours)
          ? `Remaining <= ${withinHours.toFixed(0)}h`
          : "All active service schedules";

      const pdf = await buildPdfBuffer(
        (doc) => {
          sectionTitle(doc, "Upcoming Services");
          doc
            .font("Helvetica")
            .fontSize(10)
            .text(
              `As of: ${asOfLabel} | ${pdfScopeLabel} | <= ${nearDueHours.toFixed(0)}h flagged ALMOST DUE | Est. date from ${historyDays}-day avg usage`,
            );
          doc.moveDown(0.4);

          table(
            doc,
            [
              { key: "asset_code", label: "Asset", width: 0.1 },
              { key: "asset_name", label: "Name", width: 0.16 },
              { key: "service_name", label: "Service", width: 0.15 },
              { key: "current_hours", label: "Current", width: 0.1, align: "right" },
              { key: "next_due_hours", label: "Next Due", width: 0.1, align: "right" },
              { key: "remaining_hours", label: "Remaining", width: 0.1, align: "right" },
              { key: "estimated_service_date", label: "Est. Date", width: 0.11 },
              { key: "status", label: "Status", width: 0.12 },
            ],
            dueRows.length
              ? dueRows.map((r) => ({
                  ...r,
                  current_hours: Number(r.current_hours || 0).toFixed(1),
                  next_due_hours: Number(r.next_due_hours || 0).toFixed(1),
                  remaining_hours: Number(r.remaining_hours || 0).toFixed(1),
                  estimated_service_date: r.estimated_service_date || "—",
                }))
              : [
                  {
                    asset_code: "-",
                    asset_name: requestedPlanIds.length
                      ? "No upcoming services for selected equipment"
                      : withinHours != null && Number.isFinite(withinHours)
                        ? `No services with remaining <= ${withinHours.toFixed(0)}h`
                        : "No upcoming services",
                    service_name: "-",
                    current_hours: "-",
                    next_due_hours: "-",
                    remaining_hours: "-",
                    estimated_service_date: "-",
                    status: "-",
                  },
                ]
          );
        },
        {
          title: "IRONLOG",
          subtitle: "Maintenance Upcoming Services",
          rightText: `As of: ${asOfLabel}`,
          layout: "landscape",
        }
      );

      const isDownload = String(req.query?.download || "").trim() === "1";
      reply.header("Cache-Control", "no-store, no-cache, must-revalidate");
      reply.header("Pragma", "no-cache");
      reply.header("Content-Type", "application/pdf");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="AML_Upcoming_Services_${asOfLabel}.pdf"`
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

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
  // WEEKLY FORUM SUMMARY (cross-functional alignment)
  // GET /api/maintenance/weekly-forum/summary?start=YYYY-MM-DD&end=YYYY-MM-DD&near_due_hours=50
  // =====================================================
  app.get("/weekly-forum/summary", async (req, reply) => {
    try {
      const data = await buildWeeklyForumSummary(req.query || {});
      return reply.send(data);
    } catch (err) {
      req.log.error(err);
      return reply.code(Number(err?.statusCode || 500)).send({ ok: false, error: err.message || String(err) });
    }
  });

  // =====================================================
  // WEEKLY FORUM PDF
  // GET /api/maintenance/weekly-forum.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&near_due_hours=50&download=1
  // =====================================================
  app.get("/weekly-forum.pdf", async (req, reply) => {
    try {
      const data = await buildWeeklyForumSummary(req.query || {});
      const start = String(data?.range?.start || "");
      const end = String(data?.range?.end || "");
      const isDownload = String(req.query?.download || "").trim() === "1";

      const pdf = await buildPdfBuffer(
        (doc) => {
          sectionTitle(doc, "Weekly Forum Summary");
          table(
            doc,
            [
              { key: "metric", label: "Metric", width: 0.62 },
              { key: "value", label: "Value", width: 0.38, align: "right" },
            ],
            [
              { metric: "Range", value: `${start} to ${end}` },
              { metric: "Open Work Orders", value: Number(data?.kpis?.open_work_orders || 0) },
              { metric: "Upcoming Services Flagged", value: Number(data?.kpis?.upcoming_services_flagged || 0) },
              {
                metric: "Stores parts (excl. oil/lube SKUs)",
                value: Number(data?.costs?.stores_parts_cost || 0).toFixed(2),
              },
              {
                metric: "Oil cost — lube log entries",
                value: Number(data?.costs?.stores_oil_from_logs || 0).toFixed(2),
              },
              {
                metric: "Oil cost — WO stock (oil/lube lines)",
                value: Number(data?.costs?.stores_oil_from_work_orders || 0).toFixed(2),
              },
              {
                metric: "Stores oil total",
                value: Number(data?.costs?.stores_oil_cost || 0).toFixed(2),
              },
              { metric: "Maintenance Labor Cost", value: Number(data?.costs?.maintenance_labor_cost || 0).toFixed(2) },
              { metric: "Weekly Total Cost", value: Number(data?.costs?.weekly_total_cost || 0).toFixed(2) },
              { metric: "Upcoming Service Forecast Cost", value: Number(data?.costs?.upcoming_service_forecast_cost || 0).toFixed(2) },
            ]
          );

          const actualRows = Array.isArray(data?.period_actuals_by_asset) ? data.period_actuals_by_asset : [];
          sectionTitle(doc, "Period actuals by equipment (historical — selected date range)");
          table(
            doc,
            [
              { key: "machine", label: "Machine", width: 0.34 },
              { key: "parts", label: "Parts $", width: 0.12, align: "right" },
              { key: "lubes", label: "Lubes $", width: 0.12, align: "right" },
              { key: "labor", label: "Labor $", width: 0.12, align: "right" },
              { key: "total", label: "Total $", width: 0.14, align: "right" },
              { key: "wos", label: "Closed WOs", width: 0.16, align: "right" },
            ],
            actualRows.length
              ? actualRows.map((r) => ({
                  machine: `${String(r.asset_code || "-")} - ${String(r.asset_name || "-")}`,
                  parts: Number(r.parts_cost || 0).toFixed(2),
                  lubes: Number(r.lubes_total_cost || 0).toFixed(2),
                  labor: Number(r.labor_cost || 0).toFixed(2),
                  total: Number(r.period_total_cost || 0).toFixed(2),
                  wos: String(r.closed_work_orders ?? 0),
                }))
              : [
                  {
                    machine: "No equipment consumption recorded in range",
                    parts: "-",
                    lubes: "-",
                    labor: "-",
                    total: "-",
                    wos: "-",
                  },
                ]
          );

          sectionTitle(doc, "Upcoming Services Forecast");
          const rows = Array.isArray(data?.upcoming_services) ? data.upcoming_services : [];
          table(
            doc,
            [
              { key: "machine", label: "Machine", width: 0.18 },
              { key: "service", label: "Service", width: 0.15 },
              { key: "current", label: "Current", width: 0.07, align: "right" },
              { key: "next", label: "Next Due", width: 0.07, align: "right" },
              { key: "remain", label: "Remaining", width: 0.07, align: "right" },
              { key: "status", label: "Status", width: 0.09 },
              { key: "oil", label: "Avg Oil Qty", width: 0.07, align: "right" },
              { key: "oil_cost", label: "Avg Oil $", width: 0.08, align: "right" },
              { key: "parts", label: "Avg Parts Qty", width: 0.07, align: "right" },
              { key: "parts_cost", label: "Avg Parts $", width: 0.08, align: "right" },
              { key: "kit", label: "Est Kit Cost", width: 0.07, align: "right" },
            ],
            rows.length
              ? rows.map((r) => ({
                  machine: `${String(r.asset_code || "-")} - ${String(r.asset_name || "-")}`,
                  service: String(r.service_name || "-"),
                  current: Number(r.current_hours || 0).toFixed(1),
                  next: Number(r.next_due_hours || 0).toFixed(1),
                  remain: Number(r.remaining_hours || 0).toFixed(1),
                  status: String(r.status || "-"),
                  oil: Number(r?.forecast?.avg_oil_qty || 0).toFixed(1),
                  oil_cost: Number(r?.forecast?.avg_oil_cost || 0).toFixed(2),
                  parts: Number(r?.forecast?.avg_parts_qty || 0).toFixed(1),
                  parts_cost: Number(r?.forecast?.avg_parts_cost || 0).toFixed(2),
                  kit: Number(r?.forecast?.est_service_kit_cost || 0).toFixed(2),
                }))
              : [{
                  machine: "-",
                  service: "No upcoming services within threshold",
                  current: "-",
                  next: "-",
                  remain: "-",
                  status: "-",
                  oil: "-",
                  oil_cost: "-",
                  parts: "-",
                  parts_cost: "-",
                  kit: "-",
                }]
          );
        },
        {
          title: "IRONLOG",
          subtitle: "Weekly Forum",
          rightText: `${start} to ${end}`,
          showPageNumbers: true,
          layout: "landscape",
        }
      );

      reply.header("Content-Type", "application/pdf");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="AML_Weekly_Forum_${end}.pdf"`
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(Number(err?.statusCode || 500)).send({ ok: false, error: err.message || String(err) });
    }
  });

  // =====================================================
  // WEEKLY INSPECTION CALENDAR
  // =====================================================
  app.get("/weekly-inspections/candidate-assets", async (req, reply) => {
    try {
      ensureWeeklyInspectionSchema();
      const days = Math.max(7, Math.min(120, Number(req.query?.days || 30) || 30));
      const since = addDaysYmd(new Date().toISOString().slice(0, 10), -days);
      const rows = db.prepare(`
        SELECT
          a.id AS asset_id,
          a.asset_code,
          a.asset_name,
          a.category,
          MAX(dh.work_date) AS last_used_date,
          ROUND(SUM(CASE WHEN dh.work_date >= ? THEN COALESCE(dh.hours_run, 0) ELSE 0 END), 1) AS recent_hours
        FROM assets a
        INNER JOIN daily_hours dh ON dh.asset_id = a.id
        WHERE COALESCE(a.active, 1) = 1
          AND COALESCE(a.archived, 0) = 0
          AND COALESCE(a.is_standby, 0) = 0
          AND dh.work_date >= ?
          AND COALESCE(dh.is_used, 1) = 1
          AND COALESCE(dh.hours_run, 0) > 0
        GROUP BY a.id, a.asset_code, a.asset_name, a.category
        ORDER BY a.category ASC, a.asset_code ASC
      `).all(since, since);
      const groups = {};
      for (const r of rows) {
        const cat = normalizeEquipCategory(r.category);
        if (!groups[cat]) groups[cat] = [];
        groups[cat].push(r);
      }
      return reply.send({
        ok: true,
        since,
        days,
        assets: rows,
        groups,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/weekly-inspections/calendar", async (req, reply) => {
    try {
      const month = String(req.query?.month || "").trim();
      const weekStart = String(req.query?.week_start || "").trim();
      if (month && !isMonth(month)) {
        return reply.code(400).send({ ok: false, error: "month must be YYYY-MM" });
      }
      if (weekStart && !isDate(weekStart)) {
        return reply.code(400).send({ ok: false, error: "week_start must be YYYY-MM-DD" });
      }
      return reply.send(buildWeeklyInspectionCalendarData(req.query || {}));
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/weekly-inspections/assets", async (_req, reply) => {
    try {
      const rows = db.prepare(`
        SELECT
          wia.id,
          wia.asset_id,
          wia.notes,
          wia.sort_order,
          wia.active,
          COALESCE(wia.est_minutes, 30) AS est_minutes,
          a.asset_code,
          a.asset_name
        FROM weekly_inspection_assets wia
        JOIN assets a ON a.id = wia.asset_id
        WHERE COALESCE(wia.active, 1) = 1
        ORDER BY COALESCE(wia.sort_order, 0), a.asset_code ASC
      `).all();
      return reply.send({ ok: true, assets: rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/weekly-inspections/assets", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const notes = String(req.body?.notes || "").trim();
      const est_minutes = Math.max(5, Number(req.body?.est_minutes ?? 30) || 30);
      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      const asset = db.prepare(`
        SELECT id, asset_code, asset_name
        FROM assets
        WHERE id = ?
          AND COALESCE(active, 1) = 1
          AND COALESCE(archived, 0) = 0
      `).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });
      const existing = db.prepare(`SELECT id FROM weekly_inspection_assets WHERE asset_id = ?`).get(asset_id);
      if (existing) {
        db.prepare(`
          UPDATE weekly_inspection_assets
          SET active = 1, notes = ?, est_minutes = ?, updated_at = datetime('now')
          WHERE asset_id = ?
        `).run(notes, est_minutes, asset_id);
        return reply.send({ ok: true, id: Number(existing.id), asset_id, reactivated: true });
      }
      const maxSort = Number(db.prepare(`SELECT COALESCE(MAX(sort_order), 0) AS m FROM weekly_inspection_assets`).get()?.m || 0);
      const ins = db.prepare(`
        INSERT INTO weekly_inspection_assets (asset_id, notes, sort_order, est_minutes, active, created_at, updated_at)
        VALUES (?, ?, ?, ?, 1, datetime('now'), datetime('now'))
      `).run(asset_id, notes, maxSort + 1, est_minutes);
      return reply.send({ ok: true, id: Number(ins.lastInsertRowid), asset_id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/weekly-inspections/assets/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
      const row = db.prepare(`
        SELECT id, asset_id, notes, COALESCE(est_minutes, 30) AS est_minutes
        FROM weekly_inspection_assets
        WHERE id = ? AND COALESCE(active, 1) = 1
      `).get(id);
      if (!row) return reply.code(404).send({ ok: false, error: "Schedule row not found" });
      const notes = req.body?.notes != null ? String(req.body.notes || "").trim() : String(row.notes || "").trim();
      const est_minutes = req.body?.est_minutes != null
        ? Math.max(5, Number(req.body.est_minutes) || 30)
        : Number(row.est_minutes || 30);
      db.prepare(`
        UPDATE weekly_inspection_assets
        SET notes = ?, est_minutes = ?, updated_at = datetime('now')
        WHERE id = ?
      `).run(notes, est_minutes, id);
      db.prepare(`
        UPDATE weekly_inspection_slots
        SET est_minutes = ?, updated_at = datetime('now')
        WHERE asset_id = ?
      `).run(est_minutes, Number(row.asset_id));
      return reply.send({ ok: true, id, asset_id: Number(row.asset_id), notes, est_minutes });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/weekly-inspections/slots", async (req, reply) => {
    try {
      const planned_date = String(req.body?.planned_date || "").trim();
      const asset_id = Number(req.body?.asset_id || 0);
      const est_minutes = req.body?.est_minutes != null
        ? Math.max(5, Number(req.body.est_minutes) || 30)
        : undefined;
      const slot = addWeeklyInspectionSlot({ planned_date, asset_id, est_minutes });
      return reply.send({ ok: true, slot });
    } catch (err) {
      req.log.error(err);
      return reply.code(400).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/weekly-inspections/slots/copy-day", async (req, reply) => {
    try {
      const from_date = String(req.body?.from_date || "").trim();
      const to_date = String(req.body?.to_date || "").trim();
      const result = copyWeeklyInspectionDay(from_date, to_date);
      return reply.send({ ok: true, ...result });
    } catch (err) {
      req.log.error(err);
      return reply.code(400).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.delete("/weekly-inspections/slots/:id", async (req, reply) => {
    try {
      ensureWeeklyInspectionSchema();
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid slot id" });
      const row = db.prepare(`SELECT id, planned_date, asset_id FROM weekly_inspection_slots WHERE id = ?`).get(id);
      if (!row) return reply.code(404).send({ ok: false, error: "Inspection slot not found" });
      db.prepare(`DELETE FROM weekly_inspection_slots WHERE id = ?`).run(id);
      return reply.send({ ok: true, id, planned_date: row.planned_date, asset_id: Number(row.asset_id) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.delete("/weekly-inspections/assets/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
      const row = db.prepare(`SELECT id, asset_id FROM weekly_inspection_assets WHERE id = ?`).get(id);
      if (!row) return reply.code(404).send({ ok: false, error: "Schedule row not found" });
      db.prepare(`UPDATE weekly_inspection_assets SET active = 0, updated_at = datetime('now') WHERE id = ?`).run(id);
      return reply.send({ ok: true, id, asset_id: Number(row.asset_id) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.delete("/weekly-inspections/roster", async (req, reply) => {
    try {
      const clearSlots = String(req.query?.clear_slots ?? "1").trim() !== "0";
      const result = clearWeeklyInspectionRoster({ clear_slots: clearSlots });
      return reply.send({ ok: true, ...result, clear_slots: clearSlots });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/weekly-inspections/entries", async (req, reply) => {
    try {
      const slot_id = Number(req.body?.slot_id || 0);
      const asset_id = Number(req.body?.asset_id || 0);
      const planned_date = String(req.body?.planned_date || "").trim();
      const status = String(req.body?.status || "pending").trim().toLowerCase();
      const inspector_name = String(req.body?.inspector_name || "").trim();
      const slot = updateWeeklyInspectionSlotStatus({
        slot_id: slot_id || undefined,
        asset_id: asset_id || undefined,
        planned_date,
        status,
        inspector_name,
      });
      return reply.send({ ok: true, slot });
    } catch (err) {
      req.log.error(err);
      return reply.code(400).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/weekly-inspections.pdf", async (req, reply) => {
    try {
      const data = buildWeeklyInspectionCalendarData(req.query || {});
      const isDownload = String(req.query?.download || "").trim() === "1";
      const compliance = data.compliance || {};
      const rosterAssets = Array.isArray(data.assets) ? data.assets : [];
      const weeklyGaps = Array.isArray(compliance.weekly_gaps) ? compliance.weekly_gaps : [];
      const monthLabel = String(data.month || "");
      const branding = getPdfReportBranding(db);
      const pdf = await buildPdfBuffer(
        (doc) => {
          doc.y = pdfBodyTop(doc, { siteName: branding.site_name });
          doc.font("Helvetica").fontSize(10).fillColor("#334155");
          doc.text(
            `Released: ${Number(compliance.done_count ?? 0)} / ${Number(compliance.total_slots ?? 0)} · Overdue: ${Number(compliance.not_released_count ?? 0)} · Planned: ${wiFormatMinutesPdf(compliance.est_minutes_total || 0)}`,
            doc.page.margins.left,
            doc.y,
            { width: doc.page.width - doc.page.margins.left - doc.page.margins.right },
          );
          doc.moveDown(0.55);
          drawWeeklyInspectionCalendarPdfGrid(doc, data, { siteName: branding.site_name });
          wiEnsureBodySpace(doc, 36, branding.site_name);
          if (rosterAssets.length || weeklyGaps.length) {
            if (rosterAssets.length) {
              sectionTitle(doc, "Workshop roster");
              doc.font("Helvetica").fontSize(8).fillColor("#334155");
              doc.text(
                rosterAssets
                  .map((a) => `${String(a.asset_code || "-")} (${Number(a.est_minutes || 30)} min default)`)
                  .join("  ·  "),
                { lineGap: 2 },
              );
            }
            if (weeklyGaps.length) {
              wiEnsureBodySpace(doc, 48, branding.site_name);
              doc.moveDown(0.35);
              sectionTitle(doc, "Missing weekly workshop visit");
              doc.fontSize(8).fillColor("#b45309");
              const gapLines = weeklyGaps.slice(0, 40).map((g) =>
                `${String(g.asset_code || "-")} — week ${String(g.week_start || "").slice(5)} to ${String(g.week_end || "").slice(5)}`,
              );
              doc.text(gapLines.join("\n"), { lineGap: 2 });
              if (weeklyGaps.length > 40) doc.text(`…and ${weeklyGaps.length - 40} more`, { lineGap: 2 });
            }
          }
          wiEnsureBodySpace(doc, 20, branding.site_name);
          doc.moveDown(0.35);
          doc.fontSize(8).fillColor("#64748b");
          doc.text("Legend: REL = released  |  SKIP = skipped  |  PEN = pending");
        },
        {
          title: "IRONLOG",
          subtitle: "Workshop Inspection Calendar",
          rightText: monthLabel,
          showPageNumbers: true,
          layout: "landscape",
        },
      );
      reply.header("Content-Type", "application/pdf");
      reply.header("Cache-Control", "no-store");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="IRONLOG_Workshop_Inspections_${monthLabel || "calendar"}.pdf"`
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

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
    CREATE TABLE IF NOT EXISTS weekly_forum_service_inputs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      plan_id INTEGER NOT NULL UNIQUE,
      oil_part_code TEXT,
      oil_qty REAL NOT NULL DEFAULT 0,
      parts_part_code TEXT,
      parts_qty REAL NOT NULL DEFAULT 0,
      items_json TEXT NOT NULL DEFAULT '[]',
      notes TEXT,
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
  } catch {}

  app.get("/weekly-forum/actions", async (req, reply) => {
    try {
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const status = String(req.query?.status || "").trim().toLowerCase();
      if (start && !isDate(start)) return reply.code(400).send({ ok: false, error: "start must be YYYY-MM-DD" });
      if (end && !isDate(end)) return reply.code(400).send({ ok: false, error: "end must be YYYY-MM-DD" });

      const where = [];
      const params = [];
      if (start) {
        where.push("action_date >= ?");
        params.push(start);
      }
      if (end) {
        where.push("action_date <= ?");
        params.push(end);
      }
      if (status) {
        where.push("LOWER(COALESCE(status,'open')) = ?");
        params.push(status);
      }

      const rows = db.prepare(`
        SELECT
          id, action_date, department, action_item, owner_name, due_date,
          status, notes, created_at, updated_at
        FROM weekly_forum_actions
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY
          CASE LOWER(COALESCE(status, 'open'))
            WHEN 'open' THEN 0
            WHEN 'in_progress' THEN 1
            WHEN 'blocked' THEN 2
            ELSE 3
          END ASC,
          COALESCE(due_date, action_date) ASC,
          id DESC
      `).all(...params);
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
  app.get("/weekly-forum/parts", async (req, reply) => {
    try {
      const hasPartsTable = Boolean(
        db.prepare(`
          SELECT 1
          FROM sqlite_master
          WHERE type = 'table' AND name = 'parts'
          LIMIT 1
        `).get()
      );
      const rows = hasPartsTable
        ? db.prepare(`
            SELECT
              p.id,
              p.part_code,
              p.part_name,
              COALESCE(SUM(sm.quantity), 0) AS on_hand,
              COALESCE((
                SELECT COALESCE(sm2.unit_cost_usd, sm2.cost_input, 0)
                FROM stock_movements sm2
                WHERE sm2.part_id = p.id
                  AND COALESCE(sm2.unit_cost_usd, sm2.cost_input, 0) > 0
                ORDER BY sm2.id DESC
                LIMIT 1
              ), 0) AS latest_unit_cost
            FROM parts p
            LEFT JOIN stock_movements sm ON sm.part_id = p.id
            GROUP BY p.id
            ORDER BY p.part_code ASC
            LIMIT 1500
          `).all()
        : [];
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
  app.get("/weekly-forum/forecast-inputs", async (req, reply) => {
    try {
      const rows = db.prepare(`
        SELECT id, plan_id, oil_part_code, oil_qty, parts_part_code, parts_qty, items_json, notes,
          COALESCE(labor_total, 0) AS labor_total, updated_at
        FROM weekly_forum_service_inputs
        ORDER BY plan_id ASC
      `).all();
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
  app.post("/weekly-forum/forecast-inputs", async (req, reply) => {
    try {
      const plan_id = Number(req.body?.plan_id || 0);
      const items = Array.isArray(req.body?.items) ? req.body.items : [];
      const oil_part_code = String(req.body?.oil_part_code || "").trim() || null;
      const oil_qty = Math.max(0, Number(req.body?.oil_qty || 0));
      const parts_part_code = String(req.body?.parts_part_code || "").trim() || null;
      const parts_qty = Math.max(0, Number(req.body?.parts_qty || 0));
      const notes = String(req.body?.notes || "").trim() || null;
      const labor_total = Math.max(0, Number(req.body?.labor_total || 0));
      const normalizedItems = items
        .map((it) => ({
          type: String(it?.type || "part").toLowerCase() === "oil" ? "oil" : "part",
          part_code: String(it?.part_code || "").trim(),
          qty: Math.max(0, Number(it?.qty || 0)),
        }))
        .filter((it) => it.part_code && it.qty > 0);
      const items_json = JSON.stringify(normalizedItems);
      if (!plan_id) return reply.code(400).send({ ok: false, error: "plan_id is required" });
      db.prepare(`
        INSERT INTO weekly_forum_service_inputs (
          plan_id, oil_part_code, oil_qty, parts_part_code, parts_qty, items_json, notes, labor_total, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
        ON CONFLICT(plan_id) DO UPDATE SET
          oil_part_code = excluded.oil_part_code,
          oil_qty = excluded.oil_qty,
          parts_part_code = excluded.parts_part_code,
          parts_qty = excluded.parts_qty,
          items_json = excluded.items_json,
          notes = excluded.notes,
          labor_total = excluded.labor_total,
          updated_at = datetime('now')
      `).run(plan_id, oil_part_code, oil_qty, parts_part_code, parts_qty, items_json, notes, labor_total);
      return reply.send({ ok: true, plan_id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/weekly-forum/actions", async (req, reply) => {
    try {
      const action_date = String(req.body?.action_date || "").trim() || new Date().toISOString().slice(0, 10);
      const department = String(req.body?.department || "").trim();
      const action_item = String(req.body?.action_item || "").trim();
      const owner_name = String(req.body?.owner_name || "").trim();
      const due_date = String(req.body?.due_date || "").trim() || null;
      const status = String(req.body?.status || "open").trim().toLowerCase() || "open";
      const notes = String(req.body?.notes || "").trim() || null;

      if (!isDate(action_date)) return reply.code(400).send({ ok: false, error: "action_date must be YYYY-MM-DD" });
      if (!department) return reply.code(400).send({ ok: false, error: "department is required" });
      if (!action_item) return reply.code(400).send({ ok: false, error: "action_item is required" });
      if (!owner_name) return reply.code(400).send({ ok: false, error: "owner_name is required" });
      if (due_date && !isDate(due_date)) return reply.code(400).send({ ok: false, error: "due_date must be YYYY-MM-DD" });

      const allowedStatuses = ["open", "in_progress", "blocked", "done"];
      if (!allowedStatuses.includes(status)) {
        return reply.code(400).send({ ok: false, error: "status must be open|in_progress|blocked|done" });
      }

      const ins = db.prepare(`
        INSERT INTO weekly_forum_actions (
          action_date, department, action_item, owner_name, due_date, status, notes, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(action_date, department, action_item, owner_name, due_date, status, notes);
      return reply.send({ ok: true, id: Number(ins.lastInsertRowid) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/weekly-forum/actions/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
      const existing = db.prepare(`SELECT id FROM weekly_forum_actions WHERE id = ?`).get(id);
      if (!existing) return reply.code(404).send({ ok: false, error: "action not found" });

      const department = req.body?.department != null ? String(req.body.department).trim() : undefined;
      const action_item = req.body?.action_item != null ? String(req.body.action_item).trim() : undefined;
      const owner_name = req.body?.owner_name != null ? String(req.body.owner_name).trim() : undefined;
      const due_date = req.body?.due_date != null ? (String(req.body.due_date).trim() || null) : undefined;
      const status = req.body?.status != null ? String(req.body.status).trim().toLowerCase() : undefined;
      const notes = req.body?.notes != null ? (String(req.body.notes).trim() || null) : undefined;

      if (due_date !== undefined && due_date && !isDate(due_date)) {
        return reply.code(400).send({ ok: false, error: "due_date must be YYYY-MM-DD" });
      }
      if (status !== undefined) {
        const allowedStatuses = ["open", "in_progress", "blocked", "done"];
        if (!allowedStatuses.includes(status)) {
          return reply.code(400).send({ ok: false, error: "status must be open|in_progress|blocked|done" });
        }
      }

      db.prepare(`
        UPDATE weekly_forum_actions
        SET
          department = COALESCE(?, department),
          action_item = COALESCE(?, action_item),
          owner_name = COALESCE(?, owner_name),
          due_date = COALESCE(?, due_date),
          status = COALESCE(?, status),
          notes = COALESCE(?, notes),
          updated_at = datetime('now')
        WHERE id = ?
      `).run(department ?? null, action_item ?? null, owner_name ?? null, due_date ?? null, status ?? null, notes ?? null, id);

      return reply.send({ ok: true, id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/histogram/events", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const location = String(req.query?.location || "").trim();
      const approval = String(req.query?.approval || "").trim();
      const part = String(req.query?.part || "").trim();
      const limitRaw = Number(req.query?.limit || 300);
      const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(2000, Math.trunc(limitRaw))) : 300;

      const where = ["COALESCE(site_code, 'main') = ?"];
      const params = [site_code];
      if (isDate(start)) {
        where.push("event_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("event_date <= ?");
        params.push(end);
      }
      if (location) {
        where.push("LOWER(COALESCE(location, '')) LIKE ?");
        params.push(`%${location.toLowerCase()}%`);
      }
      if (approval) {
        where.push("LOWER(COALESCE(approval_status, '')) LIKE ?");
        params.push(`%${approval.toLowerCase()}%`);
      }
      if (part) {
        where.push("(LOWER(COALESCE(part_code, '')) LIKE ? OR LOWER(COALESCE(part_name, '')) LIKE ?)");
        params.push(`%${part.toLowerCase()}%`, `%${part.toLowerCase()}%`);
      }

      const sql = `
        SELECT id, event_date, asset_number, location, part_code, part_name, approval_status, approved_by, notes, created_by, created_at, updated_at
        FROM maintenance_histogram_events
        WHERE ${where.join(" AND ")}
        ORDER BY event_date DESC, id DESC
        LIMIT ${limit}
      `;
      const rows = db.prepare(sql).all(...params);
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/histogram/events", async (req, reply) => {
    try {
      const body = req.body || {};
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const created_by = String(req.headers?.["x-user-name"] || "system").trim() || "system";
      const event_date = String(body.event_date || "").trim();
      if (!isDate(event_date)) {
        return reply.code(400).send({ ok: false, error: "event_date must be YYYY-MM-DD" });
      }
      const location = String(body.location || "").trim();
      const asset_number = String(body.asset_number || "").trim();
      const part_code = String(body.part_code || "").trim();
      const part_name = String(body.part_name || "").trim();
      const approval_status = String(body.approval_status || "").trim();
      const approved_by = String(body.approved_by || "").trim();
      const notes = String(body.notes || "").trim();
      const now = new Date().toISOString();
      const info = db.prepare(`
        INSERT INTO maintenance_histogram_events (
          site_code, event_date, asset_number, location, part_code, part_name, approval_status, approved_by, notes, created_by, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).run(site_code, event_date, asset_number, location, part_code, part_name, approval_status, approved_by, notes, created_by, now, now);
      return reply.send({ ok: true, id: Number(info.lastInsertRowid || 0) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/histogram/events/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
      const body = req.body || {};
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const existing = db.prepare(`SELECT id FROM maintenance_histogram_events WHERE id = ? AND COALESCE(site_code, 'main') = ?`).get(id, site_code);
      if (!existing) return reply.code(404).send({ ok: false, error: "Event not found" });

      const event_date = String(body.event_date || "").trim();
      if (event_date && !isDate(event_date)) {
        return reply.code(400).send({ ok: false, error: "event_date must be YYYY-MM-DD" });
      }
      const location = String(body.location || "").trim();
      const asset_number = String(body.asset_number || "").trim();
      const part_code = String(body.part_code || "").trim();
      const part_name = String(body.part_name || "").trim();
      const approval_status = String(body.approval_status || "").trim();
      const approved_by = String(body.approved_by || "").trim();
      const notes = String(body.notes || "").trim();
      const now = new Date().toISOString();

      db.prepare(`
        UPDATE maintenance_histogram_events
        SET
          event_date = COALESCE(NULLIF(?, ''), event_date),
          asset_number = ?,
          location = ?,
          part_code = ?,
          part_name = ?,
          approval_status = ?,
          approved_by = ?,
          notes = ?,
          updated_at = ?
        WHERE id = ?
          AND COALESCE(site_code, 'main') = ?
      `).run(event_date, asset_number, location, part_code, part_name, approval_status, approved_by, notes, now, id, site_code);
      return reply.send({ ok: true, id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.delete("/histogram/events/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const info = db.prepare(`
        DELETE FROM maintenance_histogram_events
        WHERE id = ?
          AND COALESCE(site_code, 'main') = ?
      `).run(id, site_code);
      if (!Number(info.changes || 0)) return reply.code(404).send({ ok: false, error: "Event not found" });
      return reply.send({ ok: true, id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/histogram/events.pdf", async (req, reply) => {
    try {
      const site_code = String(req.query?.site_code || req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const location = String(req.query?.location || "").trim();
      const approval = String(req.query?.approval || "").trim();
      const part = String(req.query?.part || "").trim();
      const include_all = String(req.query?.include_all || "").trim() === "1";
      const download = String(req.query?.download || "").trim() === "1";

      const where = ["COALESCE(site_code, 'main') = ?"];
      const params = [site_code];
      if (!include_all && isDate(start)) {
        where.push("event_date >= ?");
        params.push(start);
      }
      if (!include_all && isDate(end)) {
        where.push("event_date <= ?");
        params.push(end);
      }
      if (!include_all && location) {
        where.push("LOWER(COALESCE(location, '')) LIKE ?");
        params.push(`%${location.toLowerCase()}%`);
      }
      if (!include_all && approval) {
        where.push("LOWER(COALESCE(approval_status, '')) LIKE ?");
        params.push(`%${approval.toLowerCase()}%`);
      }
      if (!include_all && part) {
        where.push("(LOWER(COALESCE(part_code, '')) LIKE ? OR LOWER(COALESCE(part_name, '')) LIKE ?)");
        params.push(`%${part.toLowerCase()}%`, `%${part.toLowerCase()}%`);
      }

      const rows = db.prepare(`
        SELECT event_date, asset_number, location, part_code, part_name, approval_status, approved_by, notes, created_by
        FROM maintenance_histogram_events
        WHERE ${where.join(" AND ")}
        ORDER BY event_date DESC, id DESC
      `).all(...params);

      const periodLabel = include_all ? "ALL EVENTS" : `${isDate(start) ? start : "-"} to ${isDate(end) ? end : "-"}`;
      const pdf = await buildPdfBuffer((doc) => {
        sectionTitle(doc, "Maintenance Histogram Events");
        doc
          .font("Helvetica")
          .fontSize(10)
          .text(`Site: ${site_code} | Period: ${periodLabel} | Total events: ${rows.length}`);
        doc.moveDown(0.4);
        table(
          doc,
          [
            { key: "event_date", label: "Date", width: 0.1 },
            { key: "asset_number", label: "Asset No", width: 0.1 },
            { key: "location", label: "Location", width: 0.13 },
            { key: "part_code", label: "Part Code", width: 0.11 },
            { key: "part_name", label: "Part Name", width: 0.13 },
            { key: "approval_status", label: "Approval", width: 0.1 },
            { key: "approved_by", label: "Approved By", width: 0.12 },
            { key: "notes", label: "Notes", width: 0.13 },
            { key: "created_by", label: "Captured By", width: 0.08 },
          ],
          rows.length
            ? rows.map((r) => ({
                event_date: String(r.event_date || "-"),
                asset_number: String(r.asset_number || "-"),
                location: String(r.location || "-"),
                part_code: String(r.part_code || "-"),
                part_name: String(r.part_name || "-"),
                approval_status: String(r.approval_status || "-"),
                approved_by: String(r.approved_by || "-"),
                notes: String(r.notes || "-"),
                created_by: String(r.created_by || "-"),
              }))
            : [{ event_date: "-", asset_number: "-", location: "No events found", part_code: "-", part_name: "-", approval_status: "-", approved_by: "-", notes: "-", created_by: "-" }]
        );
      });

      const dateTag = new Date().toISOString().slice(0, 10);
      reply
        .header("Content-Type", "application/pdf")
        .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename="maintenance-histogram-${dateTag}.pdf"`)
        .send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // =====================================================
  // MANAGER INSPECTIONS
  // =====================================================
  app.get("/inspections", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const params = [];
      const where = ["LOWER(TRIM(COALESCE(mi.site_code, 'main'))) = ?"];
      params.push(site_code);
      if (assetId > 0) {
        where.push("mi.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("mi.inspection_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("mi.inspection_date <= ?");
        params.push(end);
      }

      const rows = db.prepare(`
        SELECT
          mi.id,
          mi.asset_id,
          mi.inspection_date,
          mi.inspector_name,
          mi.notes,
          mi.machine_hours,
          mi.live_hours_snapshot,
          mi.live_hours_source,
          mi.checklist_json,
          mi.required_parts_json,
          mi.defect_severity,
          mi.defect_component,
          mi.defect_risk,
          mi.recommended_action,
          mi.inspection_type,
          mi.evidence_required,
          mi.evidence_photo_count,
          mi.work_order_id,
          mi.created_at,
          a.asset_code,
          a.asset_name
        FROM manager_inspections mi
        JOIN assets a ON a.id = mi.asset_id
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY mi.inspection_date DESC, mi.id DESC
      `).all(...params);

      const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
      let photosByInspection = new Map();
      if (ids.length) {
        const marks = ids.map(() => "?").join(",");
        const photos = db.prepare(`
          SELECT
            id,
            ${photoInspectionCol} AS inspection_id,
            ${photoPathCol} AS file_path,
            ${photoCaptionCol} AS caption,
            ${photoCreatedCol} AS created_at
          FROM manager_inspection_photos
          WHERE ${photoInspectionCol} IN (${marks})
          ORDER BY id ASC
        `).all(...ids);
        photosByInspection = photos.reduce((m, p) => {
          const k = Number(p.inspection_id);
          if (!m.has(k)) m.set(k, []);
          m.get(k).push(p);
          return m;
        }, new Map());
      }

      return reply.send({
        ok: true,
        rows: rows.map((r) => {
          const toChecklistLabel = (key) =>
            String(key || "")
              .replaceAll("_", " ")
              .replace(/\b\w/g, (m) => m.toUpperCase())
              .trim();
          let checklist = [];
          let required_parts = [];
          try {
            let parsed = JSON.parse(String(r.checklist_json || "null"));
            // mobile ingest bundle shape: { checklist, checklist_details, ... }.
            if (parsed && !Array.isArray(parsed) && typeof parsed === "object" && parsed.checklist && typeof parsed.checklist === "object") {
              parsed = parsed.checklist;
            }
            if (Array.isArray(parsed)) {
              checklist = parsed;
            } else if (parsed && typeof parsed === "object") {
              let details = null;
              if (parsed.checklist_details && typeof parsed.checklist_details === "object") {
                details = parsed.checklist_details;
              } else {
                try {
                  const d = JSON.parse(String(r.checklist_detail_json || "null"));
                  if (d && typeof d === "object") details = d;
                } catch {}
              }
              checklist = Object.entries(parsed).map(([key, status]) => {
                const st = String(status || "").trim().toLowerCase();
                const ok = st === "ok" ? true : (st === "attention" || st === "unsafe" || st === "fail" || st === "failed") ? false : null;
                const note = String(details?.[key]?.comment || details?.[key]?.note || details?.[key]?.notes || "").trim() || null;
                return { key, label: toChecklistLabel(key), ok, note };
              });
            }
          } catch {}
          try {
            const pj = JSON.parse(String(r.required_parts_json || "[]"));
            if (Array.isArray(pj)) required_parts = pj;
          } catch {}
          return {
            ...r,
            checklist,
            required_parts,
            photos: photosByInspection.get(Number(r.id)) || [],
          };
        }),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

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

  app.post("/inspections", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const inspection_date = String(req.body?.inspection_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(inspection_date)) return reply.code(400).send({ ok: false, error: "inspection_date must be YYYY-MM-DD" });

      const asset = db.prepare(`SELECT id, asset_code, asset_name FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const liveInfo = getAssetHoursInfoAsOf(asset_id, inspection_date);
      const liveSnap = Number(liveInfo.hours || 0);
      const liveSource = String(liveInfo.source || "");

      let machine_hours = null;
      const mhRaw = req.body?.machine_hours;
      if (mhRaw != null && mhRaw !== "") {
        const n = Number(mhRaw);
        if (Number.isFinite(n) && n >= 0) machine_hours = n;
      }
      if (machine_hours == null) machine_hours = liveSnap;

      const checklist = normalizeInspectionChecklist(req.body?.checklist);
      const checklist_json = JSON.stringify(checklist);
      const required_parts = normalizeInspectionParts(req.body?.required_parts);
      const required_parts_json = JSON.stringify(required_parts);
      const inspection_type = String(req.body?.inspection_type || "machine_general").trim().toLowerCase() || "machine_general";
      const defect_severity = String(req.body?.defect_severity || "").trim().toLowerCase() || null;
      const defect_component = String(req.body?.defect_component || "").trim() || null;
      const defect_risk = String(req.body?.defect_risk || "").trim() || null;
      const recommended_action = String(req.body?.recommended_action || "").trim() || null;
      const evidence_required = Number(req.body?.evidence_required ?? 1) ? 1 : 0;
      const evidence_photo_count = Math.max(0, Number(req.body?.evidence_photo_count || 0));

      const anyChecklistFail = checklist.some((c) => c.ok === false);
      const enforceEvidenceRules = Number(req.body?.enforce_evidence_rules || 0) === 1;
      if (enforceEvidenceRules && evidence_required && anyChecklistFail) {
        const missingFailComment = checklist.find((c) => c.ok === false && !String(c.note || "").trim());
        if (missingFailComment) {
          return reply.code(400).send({ ok: false, error: `Comment required for failed checklist item: ${missingFailComment.label}` });
        }
        if (evidence_photo_count < 1) {
          return reply.code(400).send({ ok: false, error: "At least one evidence photo is required for failed checklist items" });
        }
      }
      const hasParts = required_parts.length > 0;
      const createExplicit = Boolean(req.body?.create_work_order);
      const createOnIssues = req.body?.create_work_order_on_issues !== false;
      const shouldCreateWo =
        createExplicit || (createOnIssues && (anyChecklistFail || hasParts));

      const ins = db.prepare(`
        INSERT INTO manager_inspections (
          asset_id, uuid, site_code, inspection_date, inspector_name, notes,
          machine_hours, live_hours_snapshot, live_hours_source,
          checklist_json, required_parts_json,
          inspection_type,
          defect_severity, defect_component, defect_risk, recommended_action,
          evidence_required, evidence_photo_count,
          updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(
        asset_id,
        crypto.randomUUID(),
        site_code,
        inspection_date,
        inspector_name,
        notes,
        machine_hours,
        liveSnap,
        liveSource,
        checklist_json,
        required_parts_json,
        inspection_type,
        defect_severity,
        defect_component,
        defect_risk,
        recommended_action,
        evidence_required,
        evidence_photo_count
      );

      const inspectionId = Number(ins.lastInsertRowid);
      let work_order_id = null;

      if (shouldCreateWo) {
        ensureColumn("work_orders", "job_description TEXT", "job_description");
        ensureColumn("work_orders", "completion_notes TEXT", "completion_notes");
        const woNotes = buildInspectionWorkOrderNotes({
          inspectionId,
          inspection_date,
          asset_code: String(asset.asset_code || ""),
          asset_name: String(asset.asset_name || ""),
          checklist,
          required_parts,
          notes,
        });
        const wo = db.prepare(`
          INSERT INTO work_orders (asset_id, source, reference_id, status)
          VALUES (?, 'inspection', ?, 'open')
        `).run(asset_id, inspectionId);
        work_order_id = Number(wo.lastInsertRowid);
        if (work_order_id > 0 && woNotes) {
          if (hasColumn("work_orders", "job_description")) {
            db.prepare(`UPDATE work_orders SET job_description = ? WHERE id = ?`).run(woNotes, work_order_id);
          } else if (hasColumn("work_orders", "completion_notes")) {
            db.prepare(`UPDATE work_orders SET completion_notes = ? WHERE id = ?`).run(woNotes, work_order_id);
          }
        }
        if (work_order_id > 0) {
          db.prepare(`UPDATE manager_inspections SET work_order_id = ? WHERE id = ?`).run(work_order_id, inspectionId);
        }
      }

      return reply.send({ ok: true, id: inspectionId, work_order_id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/inspections/:id/photo", async (req, reply) => {
    try {
      const inspectionId = Number(req.params?.id || 0);
      if (!inspectionId) return reply.code(400).send({ ok: false, error: "Invalid inspection id" });

      const inspection = db.prepare(`SELECT id FROM manager_inspections WHERE id = ?`).get(inspectionId);
      if (!inspection) return reply.code(404).send({ ok: false, error: "Inspection not found" });

      const part = await req.file();
      if (!part) return reply.code(400).send({ ok: false, error: "Upload file field named 'file'" });

      const extRaw = path.extname(part.filename || "").toLowerCase();
      const ext = [".jpg", ".jpeg", ".png", ".webp"].includes(extRaw) ? extRaw : ".jpg";
      const safe = `mi_${inspectionId}_${Date.now()}_${Math.floor(Math.random() * 100000)}${ext}`;
      const absPath = path.join(inspectionsDir, safe);
      await fs.promises.writeFile(absPath, await part.toBuffer());

      const caption = String(req.query?.caption || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const relPath = path.join("uploads", "manager-inspections", safe).replace(/\\/g, "/");

      // Legacy compatibility: some DBs require manager_inspection_id NOT NULL,
      // others use inspection_id. If both exist, write both.
      const hasInspectionId = hasColumn("manager_inspection_photos", "inspection_id");
      const hasManagerInspectionId = hasColumn("manager_inspection_photos", "manager_inspection_id");
      const linkCols = [];
      const linkVals = [];
      if (hasInspectionId) {
        linkCols.push("inspection_id");
        linkVals.push(inspectionId);
      }
      if (hasManagerInspectionId) {
        linkCols.push("manager_inspection_id");
        linkVals.push(inspectionId);
      }
      if (!linkCols.length) {
        linkCols.push(photoInspectionCol);
        linkVals.push(inspectionId);
      }

      const hasImageData = hasColumn("manager_inspection_photos", "image_data");
      const insertCols = [...linkCols, "uuid", "site_code", "file_path", ...(hasImageData ? ["image_data"] : []), "caption", "updated_at"];
      const placeholders = [...insertCols.map((c) => (c === "updated_at" ? "datetime('now')" : "?"))].join(", ");
      const ins = db.prepare(`
        INSERT INTO manager_inspection_photos (${insertCols.join(", ")})
        VALUES (${placeholders})
      `).run(...linkVals, crypto.randomUUID(), site_code, relPath, ...(hasImageData ? [relPath] : []), caption);

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        inspection_id: inspectionId,
        file_path: `/${relPath}`,
        caption,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.delete("/inspections/:id", async (req, reply) => {
    try {
      const inspectionId = Number(req.params?.id || 0);
      if (!inspectionId) return reply.code(400).send({ ok: false, error: "Invalid inspection id" });
      const current = db.prepare(`
        SELECT id, work_order_id
        FROM manager_inspections
        WHERE id = ?
        LIMIT 1
      `).get(inspectionId);
      if (!current) return reply.code(404).send({ ok: false, error: "Inspection not found" });

      const photos = db.prepare(`
        SELECT ${photoPathCol} AS file_path
        FROM manager_inspection_photos
        WHERE ${photoInspectionCol} = ?
      `).all(inspectionId);
      for (const p of photos) {
        const rel = String(p.file_path || "").replace(/\\/g, "/").replace(/^\/+/, "");
        if (!rel) continue;
        const abs = resolveStorageAbs(rel);
        if (!abs || !fs.existsSync(abs)) continue;
        try { fs.unlinkSync(abs); } catch {}
      }

      db.prepare(`DELETE FROM manager_inspection_photos WHERE ${photoInspectionCol} = ?`).run(inspectionId);
      db.prepare(`DELETE FROM manager_inspections WHERE id = ?`).run(inspectionId);
      return reply.send({ ok: true, id: inspectionId, work_order_id: current.work_order_id || null });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

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

  app.get("/tyre-inspections/lifecycle", async (req, reply) => {
    try {
      ensureTyreLifecycleSchema();
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      if (!assetId) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      const summary = buildTyreLifecycleSummary(assetId, site_code);
      if (!summary) return reply.code(404).send({ ok: false, error: "Asset not found" });
      return reply.send({ ok: true, ...summary });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/tyre-inspections", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const params = [site_code];
      const where = ["ti.site_code = ?"];
      if (assetId > 0) {
        where.push("ti.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("ti.inspection_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("ti.inspection_date <= ?");
        params.push(end);
      }
      const rows = db.prepare(`
        SELECT
          ti.id,
          ti.asset_id,
          ti.inspection_date,
          ti.inspector_name,
          ti.running_hours,
          ti.total_tyre_cost,
          ti.cost_per_running_hour,
          ti.tyres_json,
          ti.notes,
          ti.created_at,
          a.asset_code,
          a.asset_name
        FROM tyre_inspections ti
        JOIN assets a ON a.id = ti.asset_id
        WHERE ${where.join(" AND ")}
        ORDER BY ti.inspection_date DESC, ti.id DESC
      `).all(...params);
      return reply.send({
        ok: true,
        rows: rows.map((r) => {
          let tyres = [];
          try {
            const parsed = JSON.parse(String(r.tyres_json || "[]"));
            tyres = normalizeTyreRows(parsed);
          } catch {}
          return { ...r, tyres };
        }),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/tyre-inspections", async (req, reply) => {
    try {
      ensureTyreLifecycleSchema();
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const asset_id = Number(req.body?.asset_id || 0);
      const inspection_date = String(req.body?.inspection_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(inspection_date)) return reply.code(400).send({ ok: false, error: "inspection_date must be YYYY-MM-DD" });
      const asset = db.prepare(`SELECT id FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const tyresIn = normalizeTyreRows(req.body?.tyres);
      const runningRaw = req.body?.running_hours;
      const running_hours = runningRaw == null || runningRaw === ""
        ? 0
        : Number(runningRaw);
      if (!Number.isFinite(running_hours) || running_hours < 0) {
        return reply.code(400).send({ ok: false, error: "running_hours must be a positive number" });
      }

      const thresholds = getTyreThresholds();
      const tyres = tyresIn.map((row) => enrichTyreLifecycleRow(row, {
        asset_id,
        site_code,
        inspection_date,
        running_hours,
        thresholds,
        persist: true,
      }));

      const total_tyre_cost = Number(
        tyres.reduce((sum, t) => sum + Number(t.tyre_cost || 0), 0).toFixed(2),
      );
      const withCost = tyres.filter((t) => t.cost_per_hour != null);
      const cost_per_running_hour = withCost.length
        ? Number(withCost.reduce((sum, t) => sum + Number(t.cost_per_hour || 0), 0).toFixed(4))
        : 0;
      const alerts = tyres
        .filter((t) => t.lifecycle_status === "warn" || t.lifecycle_status === "replace")
        .map((t) => ({
          position_key: t.position_key,
          position_label: t.position_label,
          lifecycle_status: t.lifecycle_status,
          tread_alert: t.tread_alert,
        }));

      const ins = db.prepare(`
        INSERT INTO tyre_inspections (
          asset_id, uuid, site_code, inspection_date, inspector_name,
          running_hours, total_tyre_cost, cost_per_running_hour, tyres_json, notes, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(
        asset_id,
        crypto.randomUUID(),
        site_code,
        inspection_date,
        inspector_name,
        running_hours,
        total_tyre_cost,
        cost_per_running_hour,
        JSON.stringify(tyres),
        notes,
      );

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        total_tyre_cost,
        cost_per_running_hour,
        fleet_cost_per_hour: cost_per_running_hour,
        tyres,
        alerts,
        thresholds,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

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

  app.get("/tyre-inspections/survey.xlsx", async (req, reply) => {
    try {
      const month = String(req.query?.month || "").trim() || new Date().toISOString().slice(0, 7);
      if (!isMonth(month)) return reply.code(400).send({ ok: false, error: "month must be YYYY-MM" });
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const data = buildTyreMonthlySurveyData(month, site_code);
      const buf = await buildTyreSurveyXlsxBuffer(data);
      reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Cache-Control", "no-store")
        .header("Content-Disposition", `attachment; filename="IRONLOG_Tyre_Survey_${month}.xlsx"`)
        .send(Buffer.from(buf));
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/tyre-inspections/survey.pdf", async (req, reply) => {
    try {
      const month = String(req.query?.month || "").trim() || new Date().toISOString().slice(0, 7);
      if (!isMonth(month)) return reply.code(400).send({ ok: false, error: "month must be YYYY-MM" });
      const isDownload = String(req.query?.download || "").trim() === "1";
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const data = buildTyreMonthlySurveyData(month, site_code);
      const branding = data.branding || getPdfReportBranding(db);
      const monthLabel = String(data.month || "");

      const pdf = await buildPdfBuffer(
        (doc) => {
          doc.y = pdfBodyTop(doc, { siteName: branding.site_name });
          doc.font("Helvetica").fontSize(10).fillColor("#334155");
          doc.text(
            `Project: ${data.project || "-"}   Country: ${data.country || "-"}   Machines: ${data.machines?.length || 0}`,
          );
          doc.moveDown(0.5);
          if (!data.machines?.length) {
            doc.text("No tyre inspections found for this month.");
            return;
          }
          for (const machine of data.machines) {
            drawTyreSurveyPdfMachine(doc, machine, branding.site_name);
          }
        },
        {
          title: "IRONLOG",
          subtitle: "Monthly Tyre Survey Report",
          rightText: monthLabel,
          showPageNumbers: true,
          layout: "landscape",
        },
      );

      reply.header("Content-Type", "application/pdf");
      reply.header("Cache-Control", "no-store");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="IRONLOG_Tyre_Survey_${monthLabel}.pdf"`,
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

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

  app.get("/undercarriage-inspections/template", async (_req, reply) => {
    return reply.send({
      ok: true,
      components: buildUndercarriageComponentSchema(),
      checklist_items: UNDERCARRIAGE_CHECKLIST_ITEMS,
      track_sag_points: UNDERCARRIAGE_TRACK_SAG_POINTS,
      wear_bands: UNDERCARRIAGE_WEAR_BANDS,
    });
  });

  app.get("/undercarriage-inspections/wear-profile", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      if (!assetId) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      const asset = db.prepare(`SELECT id, asset_code, asset_name, category FROM assets WHERE id = ?`).get(assetId);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });
      const profile = getUndercarriageWearProfileRow(assetId, site_code);
      return reply.send({
        ok: true,
        asset,
        profile,
        has_profile: Boolean(profile?.configured_count),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/undercarriage-inspections/wear-profile", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const asset_id = Number(req.body?.asset_id || 0);
      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      const asset = db.prepare(`SELECT id, asset_code FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });
      const updated_by = String(req.headers?.["x-user-name"] || req.body?.updated_by || "").trim() || null;
      const profile = saveUndercarriageWearProfile({
        asset_id,
        site_code,
        limits: req.body?.limits,
        source: String(req.body?.source || "manual").trim(),
        notes: String(req.body?.notes || "").trim() || null,
        updated_by,
      });
      return reply.send({
        ok: true,
        asset_code: asset.asset_code,
        profile,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/undercarriage-inspections/wear-profile/import-latest", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const asset_id = Number(req.body?.asset_id || req.query?.asset_id || 0);
      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      const asset = db.prepare(`SELECT id, asset_code FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });
      const updated_by = String(req.headers?.["x-user-name"] || "").trim() || null;
      const profile = importUndercarriageWearProfileFromLatest(asset_id, site_code, updated_by);
      if (!profile) {
        return reply.code(404).send({ ok: false, error: "No previous inspection found to import limits from" });
      }
      return reply.send({
        ok: true,
        asset_code: asset.asset_code,
        profile,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/undercarriage-inspections/latest", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      if (!assetId) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      const row = db.prepare(`
        SELECT ui.*, a.asset_code, a.asset_name, a.category
        FROM undercarriage_inspections ui
        JOIN assets a ON a.id = ui.asset_id
        WHERE ui.asset_id = ?
          AND LOWER(TRIM(COALESCE(ui.site_code, 'main'))) = LOWER(TRIM(?))
        ORDER BY ui.inspection_date DESC, ui.id DESC
        LIMIT 1
      `).get(assetId, site_code);
      if (!row) return reply.send({ ok: true, row: null });
      return reply.send({ ok: true, row: parseUndercarriageInspectionRow(row) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/undercarriage-inspections", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const params = [site_code];
      const where = ["LOWER(TRIM(COALESCE(ui.site_code, 'main'))) = LOWER(TRIM(?))"];
      if (assetId > 0) {
        where.push("ui.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("ui.inspection_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("ui.inspection_date <= ?");
        params.push(end);
      }
      const rows = db.prepare(`
        SELECT ui.*, a.asset_code, a.asset_name, a.category
        FROM undercarriage_inspections ui
        JOIN assets a ON a.id = ui.asset_id
        WHERE ${where.join(" AND ")}
        ORDER BY ui.inspection_date DESC, ui.id DESC
        LIMIT 200
      `).all(...params);
      return reply.send({
        ok: true,
        rows: rows.map((r) => parseUndercarriageInspectionRow(r)),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/undercarriage-inspections", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const asset_id = Number(req.body?.asset_id || 0);
      const inspection_date = String(req.body?.inspection_date || "").trim() || new Date().toISOString().slice(0, 10);
      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(inspection_date)) return reply.code(400).send({ ok: false, error: "inspection_date must be YYYY-MM-DD" });
      const asset = db.prepare(`SELECT id, asset_code, asset_name, category FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const smuRaw = req.body?.smu ?? req.body?.running_hours;
      const smu = smuRaw == null || smuRaw === "" ? null : Number(smuRaw);
      if (smu != null && (!Number.isFinite(smu) || smu < 0)) {
        return reply.code(400).send({ ok: false, error: "smu must be a positive number" });
      }

      const measurements = buildUndercarriageMeasurementsForSave(req.body?.measurements, {
        asset_id,
        site_code,
        inspection_date,
        smu,
      });
      const track_sag = normalizeUndercarriageTrackSag(req.body?.track_sag || {});
      const checklist = normalizeUndercarriageChecklist(req.body?.checklist || {});
      const summary = summarizeUndercarriageInspection(measurements);
      const branding = getPdfReportBranding(db);

      const ins = db.prepare(`
        INSERT INTO undercarriage_inspections (
          asset_id, uuid, site_code, inspection_date, inspector_name, smu,
          job_no, site_name, planner, serial_no, unit_assembly, model, yard_no,
          work_order_no, component_group, group_id, component_serial_no, part_no,
          cost_center, measurements_json, track_sag_json, checklist_json, summary_json,
          notes, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(
        asset_id,
        crypto.randomUUID(),
        site_code,
        inspection_date,
        String(req.body?.inspector_name || "").trim() || null,
        smu,
        String(req.body?.job_no || "").trim() || null,
        String(req.body?.site_name || branding.site_name || "").trim() || null,
        String(req.body?.planner || "").trim() || null,
        String(req.body?.serial_no || "").trim() || null,
        String(req.body?.unit_assembly || "").trim() || null,
        String(req.body?.model || req.body?.machine_model || asset.category || "").trim() || null,
        String(req.body?.yard_no || "").trim() || null,
        String(req.body?.work_order_no || "").trim() || null,
        String(req.body?.component_group || "").trim() || null,
        String(req.body?.group_id || "").trim() || null,
        String(req.body?.component_serial_no || "").trim() || null,
        String(req.body?.part_no || "").trim() || null,
        String(req.body?.cost_center || "").trim() || null,
        JSON.stringify(measurements),
        JSON.stringify(track_sag),
        JSON.stringify(checklist),
        JSON.stringify(summary),
        String(req.body?.notes || checklist.comments || "").trim() || null,
      );

      const updateWearProfile = req.body?.update_wear_profile !== false;
      if (updateWearProfile) {
        const limits = measurements.map((m) => ({
          key: m.key,
          base: m.base,
          wear_limit: m.wear_limit,
        }));
        saveUndercarriageWearProfile({
          asset_id,
          site_code,
          limits,
          source: "inspection_save",
          updated_by: String(req.body?.inspector_name || req.headers?.["x-user-name"] || "").trim() || null,
        });
      }

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        asset_code: asset.asset_code,
        summary,
        measurements,
        track_sag,
        checklist,
        pdf_url: `/api/maintenance/undercarriage-inspections/${Number(ins.lastInsertRowid)}.pdf`,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

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

  app.get("/undercarriage-inspections/:id.pdf", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "invalid id" });
      const isDownload = String(req.query?.download || "").trim() === "1";
      const row = db.prepare(`
        SELECT ui.*, a.asset_code, a.asset_name, a.category
        FROM undercarriage_inspections ui
        JOIN assets a ON a.id = ui.asset_id
        WHERE ui.id = ?
        LIMIT 1
      `).get(id);
      if (!row) return reply.code(404).send({ ok: false, error: "Inspection not found" });
      const insp = parseUndercarriageInspectionRow(row);
      const branding = getPdfReportBranding(db);
      const pdf = await buildPdfBuffer(
        (doc) => drawUndercarriageInspectionPdf(doc, insp, branding),
        {
          title: "IRONLOG",
          subtitle: "Undercarriage Report",
          rightText: insp.asset_code || "",
          showPageNumbers: true,
          layout: "landscape",
        },
      );
      reply.header("Content-Type", "application/pdf");
      reply.header("Cache-Control", "no-store");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="IRONLOG_Undercarriage_${insp.asset_code || "report"}_${insp.inspection_date || id}.pdf"`,
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

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

  app.get("/undercarriage-inspections/report.xlsx", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const inspectionId = Number(req.query?.inspection_id || req.query?.id || 0);
      const assetId = Number(req.query?.asset_id || 0);
      const month = String(req.query?.month || "").trim();
      let rows = [];

      if (inspectionId > 0) {
        const one = db.prepare(`
          SELECT ui.*, a.asset_code, a.asset_name, a.category
          FROM undercarriage_inspections ui
          JOIN assets a ON a.id = ui.asset_id
          WHERE ui.id = ?
          LIMIT 1
        `).get(inspectionId);
        if (!one) return reply.code(404).send({ ok: false, error: "Inspection not found" });
        rows = [parseUndercarriageInspectionRow(one)];
      } else {
        const params = [site_code];
        const where = ["LOWER(TRIM(COALESCE(ui.site_code, 'main'))) = LOWER(TRIM(?))"];
        if (assetId > 0) {
          where.push("ui.asset_id = ?");
          params.push(assetId);
        }
        if (isMonth(month)) {
          const bounds = monthBoundsYmd(month);
          if (bounds) {
            where.push("ui.inspection_date >= ?");
            where.push("ui.inspection_date <= ?");
            params.push(bounds.start, bounds.end);
          }
        }
        const raw = db.prepare(`
          SELECT ui.*, a.asset_code, a.asset_name, a.category
          FROM undercarriage_inspections ui
          JOIN assets a ON a.id = ui.asset_id
          WHERE ${where.join(" AND ")}
          ORDER BY a.asset_code ASC, ui.inspection_date DESC, ui.id DESC
        `).all(...params);
        rows = raw.map((r) => parseUndercarriageInspectionRow(r));
      }

      if (!rows.length) return reply.code(404).send({ ok: false, error: "No inspections found for export" });
      const buf = await buildUndercarriageXlsxBuffer(rows);
      const suffix = inspectionId > 0
        ? `_${rows[0]?.asset_code || "report"}_${rows[0]?.inspection_date || inspectionId}`
        : (isMonth(month) ? `_${month}` : "");
      reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Cache-Control", "no-store")
        .header("Content-Disposition", `attachment; filename="IRONLOG_Undercarriage${suffix}.xlsx"`)
        .send(Buffer.from(buf));
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // =====================================================
  // ARTISAN INSPECTIONS (daily general checklist)
  // =====================================================
  app.get("/artisan-inspections", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const assetId = Number(req.query?.asset_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const params = [];
      const where = ["LOWER(TRIM(COALESCE(ai.site_code, 'main'))) = ?"];
      params.push(site_code);
      if (assetId > 0) {
        where.push("ai.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("ai.inspection_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("ai.inspection_date <= ?");
        params.push(end);
      }

      const rows = db.prepare(`
        SELECT
          ai.id,
          ai.asset_id,
          ai.inspection_date,
          ai.inspector_name,
          ai.form_number,
          ai.shift,
          ai.notes,
          ai.machine_hours,
          ai.live_hours_snapshot,
          ai.live_hours_source,
          ai.checklist_json,
          ai.created_at,
          a.asset_code,
          a.asset_name
        FROM artisan_inspections ai
        JOIN assets a ON a.id = ai.asset_id
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY ai.inspection_date DESC, ai.id DESC
      `).all(...params);

      return reply.send({
        ok: true,
        rows: rows.map((r) => {
          let checklist = [];
          try {
            const cj = JSON.parse(String(r.checklist_json || "[]"));
            if (Array.isArray(cj)) checklist = cj;
          } catch {}
          return { ...r, checklist };
        }),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/artisan-inspections", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const inspection_date = String(req.body?.inspection_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const form_number = String(req.body?.form_number || "").trim() || null;
      const shift = String(req.body?.shift || "").trim().toLowerCase() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(inspection_date)) return reply.code(400).send({ ok: false, error: "inspection_date must be YYYY-MM-DD" });
      if (shift && !["day", "night"].includes(shift)) {
        return reply.code(400).send({ ok: false, error: "shift must be day or night" });
      }

      const asset = db.prepare(`SELECT id FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const liveInfo = getAssetHoursInfoAsOf(asset_id, inspection_date);
      const liveSnap = Number(liveInfo.hours || 0);
      const liveSource = String(liveInfo.source || "");

      let machine_hours = null;
      const mhRaw = req.body?.machine_hours;
      if (mhRaw != null && mhRaw !== "") {
        const n = Number(mhRaw);
        if (Number.isFinite(n) && n >= 0) machine_hours = n;
      }
      if (machine_hours == null) machine_hours = liveSnap;

      const checklist = normalizeArtisanInspectionChecklist(req.body?.checklist);
      const checklist_json = JSON.stringify(checklist);

      const ins = db.prepare(`
        INSERT INTO artisan_inspections (
          asset_id, uuid, site_code, inspection_date, inspector_name, form_number, shift, notes,
          machine_hours, live_hours_snapshot, live_hours_source, checklist_json, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(
        asset_id,
        crypto.randomUUID(),
        site_code,
        inspection_date,
        inspector_name,
        form_number,
        shift,
        notes,
        machine_hours,
        liveSnap,
        liveSource,
        checklist_json
      );

      return reply.send({ ok: true, id: Number(ins.lastInsertRowid) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

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

  app.get("/vehicle-ldv-checks", async (req, reply) => {
    try {
      const assetId = Number(req.query?.asset_id || 0);
      const checkIdFilter = Number(req.query?.check_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const checkModeFilter = String(req.query?.check_mode || "").trim().toLowerCase();
      const assetCodeFilter = String(req.query?.asset_code || "").trim().toUpperCase();
      const params = [];
      const where = [];
      if (checkIdFilter > 0) {
        where.push("v.id = ?");
        params.push(checkIdFilter);
      }
      if (assetId > 0) {
        where.push("v.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("v.check_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("v.check_date <= ?");
        params.push(end);
      }
      if (assetCodeFilter) {
        where.push("UPPER(a.asset_code) = ?");
        params.push(assetCodeFilter);
      }
      if (checkModeFilter === "ldv") {
        where.push("COALESCE(v.check_mode, 'ldv_general') = 'prestart'");
      } else if (checkModeFilter === "machine") {
        where.push("COALESCE(v.check_mode, '') LIKE 'machine_prestart_%'");
      } else if (checkModeFilter === "vehicle") {
        where.push("COALESCE(v.check_mode, 'ldv_general') = 'ldv_general'");
      }

      const rows = db.prepare(`
        SELECT
          v.id,
          v.asset_id,
          v.check_date,
          v.vehicle_registration,
          v.odometer_km,
          v.smu_hours,
          v.inspector_name,
          v.notes,
          v.check_mode,
          v.created_at,
          a.asset_code,
          a.asset_name
        FROM vehicle_ldv_checks v
        JOIN assets a ON a.id = v.asset_id
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY v.check_date DESC, v.id DESC
        LIMIT 500
      `).all(...params);

      const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
      let photosByCheck = new Map();
      if (ids.length) {
        const marks = ids.map(() => "?").join(",");
        const photos = db.prepare(`
          SELECT id, check_id, file_path, caption, markers_json, created_at
          FROM vehicle_ldv_check_photos
          WHERE check_id IN (${marks})
          ORDER BY id ASC
        `).all(...ids);
        photosByCheck = photos.reduce((m, p) => {
          const k = Number(p.check_id);
          if (!m.has(k)) m.set(k, []);
          let markers = [];
          try {
            markers = p.markers_json ? JSON.parse(p.markers_json) : [];
          } catch {
            markers = [];
          }
          m.get(k).push({
            ...p,
            markers: Array.isArray(markers) ? markers : [],
          });
          return m;
        }, new Map());
      }

      return reply.send({
        ok: true,
        rows: rows.map((r) => {
          const photos = photosByCheck.get(Number(r.id)) || [];
          return {
            ...r,
            photos,
            photo_count: photos.length,
          };
        }),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // GET /api/maintenance/checklist-hub?date=YYYY-MM-DD
  app.get("/checklist-hub", async (req, reply) => {
    try {
      const check_date = String(req.query?.date || "").trim() || new Date().toISOString().slice(0, 10);
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "date must be YYYY-MM-DD" });

      const assets = db.prepare(`
        SELECT id, asset_code, asset_name, category
        FROM assets
        WHERE archived = 0
        ORDER BY asset_code ASC
      `).all();

      const ldvAssets = [];
      const machineGroupMap = new Map();
      for (const profile of listMachinePrestartProfiles()) {
        machineGroupMap.set(profile.id, {
          profile_id: profile.id,
          title: profile.title,
          assets: [],
        });
      }

      for (const a of assets) {
        const assetId = Number(a.id);
        const code = String(a.asset_code || "");
        if (isLdvPrestartAssetCode(code)) {
          const row = getLdvPrestartCheckRow(assetId, check_date);
          const st = row
            ? { ...checklistRowStatus(row.checklist_json, normalizeLdvPrestartChecklist), check_id: Number(row.id) }
            : { status: "pending", check_id: null };
          ldvAssets.push({
            asset_id: assetId,
            asset_code: code,
            asset_name: String(a.asset_name || ""),
            category: String(a.category || ""),
            kind: "ldv",
            ...st,
          });
          continue;
        }
        const profileId = resolveMachinePrestartProfile(a.category, a.asset_name, code);
        if (!profileId) continue;
        const mode = machinePrestartCheckMode(profileId);
        if (!mode) continue;
        const row = getMachinePrestartCheckRow(assetId, check_date, mode);
        const st = row
          ? {
              ...checklistRowStatus(row.checklist_json, (parsed) =>
                normalizeMachinePrestartChecklist(profileId, parsed)
              ),
              check_id: Number(row.id),
            }
          : { status: "pending", check_id: null };
        const group = machineGroupMap.get(profileId);
        if (group) {
          group.assets.push({
            asset_id: assetId,
            asset_code: code,
            asset_name: String(a.asset_name || ""),
            category: String(a.category || ""),
            kind: "machine",
            profile_id: profileId,
            ...st,
          });
        }
      }

      const machine_groups = [...machineGroupMap.values()].filter((g) => g.assets.length > 0);
      const summary = {
        ldv_total: ldvAssets.length,
        ldv_compliant: ldvAssets.filter((x) => x.status === "compliant").length,
        machine_total: machine_groups.reduce((s, g) => s + g.assets.length, 0),
        machine_compliant: machine_groups.reduce(
          (s, g) => s + g.assets.filter((x) => x.status === "compliant").length,
          0
        ),
      };

      const commentRows = db.prepare(`
        SELECT v.id AS check_id, v.check_date, v.inspector_name, v.notes, v.check_mode, v.updated_at,
               a.asset_code, a.asset_name
        FROM vehicle_ldv_checks v
        JOIN assets a ON a.id = v.asset_id
        WHERE v.check_date = ?
          AND TRIM(COALESCE(v.notes, '')) != ''
        ORDER BY datetime(COALESCE(v.updated_at, v.check_date)) DESC, v.id DESC
      `).all(check_date);
      const comments = commentRows.map((r) => {
        const mode = String(r.check_mode || "").toLowerCase();
        const kind = mode.includes("machine_prestart") ? "machine" : "ldv";
        return {
          kind,
          check_id: Number(r.check_id),
          asset_code: String(r.asset_code || ""),
          asset_name: String(r.asset_name || ""),
          inspector_name: String(r.inspector_name || "").trim() || null,
          notes: String(r.notes || "").trim(),
          updated_at: r.updated_at || null,
        };
      });

      return reply.send({
        ok: true,
        check_date,
        summary,
        ldv: ldvAssets,
        machine_groups,
        comments,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.get("/vehicle-ldv-checks/prestart-context", async (req, reply) => {
    try {
      const asset_code = String(req.query?.asset_code || "").trim().toUpperCase();
      const check_date = String(req.query?.check_date || "").trim() || new Date().toISOString().slice(0, 10);
      if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "check_date must be YYYY-MM-DD" });

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name, category
        FROM assets
        WHERE UPPER(asset_code) = UPPER(?)
        LIMIT 1
      `).get(asset_code);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const existing = getLdvPrestartCheckRow(Number(asset.id), check_date);

      const raw_previous_odometer_km = getLatestLdvOdometerKm(Number(asset.id), check_date, {
        excludeCheckId: existing?.id ? Number(existing.id) : 0,
      });
      const previous_odometer_km = getSanitizedLdvBaselineKm(Number(asset.id), check_date, {
        excludeCheckId: existing?.id ? Number(existing.id) : 0,
      });
      const baseline_poisoned =
        raw_previous_odometer_km != null &&
        previous_odometer_km != null &&
        Math.abs(Number(raw_previous_odometer_km) - Number(previous_odometer_km)) > 1;

      let checklist = normalizeLdvPrestartChecklist({});
      if (existing?.checklist_json) {
        try {
          const parsed = JSON.parse(String(existing.checklist_json || "{}"));
          if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
            checklist = normalizeLdvPrestartChecklist(parsed);
          }
        } catch {}
      }

      return reply.send({
        ok: true,
        asset: {
          id: Number(asset.id),
          asset_code: String(asset.asset_code || ""),
          asset_name: String(asset.asset_name || ""),
          category: String(asset.category || ""),
        },
        check_date,
        previous_odometer_km,
        raw_previous_odometer_km,
        baseline_poisoned,
        previous_odometer_source:
          previous_odometer_km != null ? "sanitized_prior_prestart_or_daily" : null,
        existing_prestart: existing
          ? {
              id: Number(existing.id),
              odometer_km: existing.odometer_km == null ? null : Number(existing.odometer_km),
              inspector_name: existing.inspector_name || "",
              notes: existing.notes || "",
              checklist,
            }
          : null,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/vehicle-ldv-checks/prestart", async (req, reply) => {
    try {
      const asset_code = String(req.body?.asset_code || "").trim().toUpperCase();
      const check_date = String(req.body?.check_date || "").trim() || new Date().toISOString().slice(0, 10);
      const odometer_km_raw = req.body?.odometer_km;
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      const checklistObj = req.body?.checklist || {};
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "check_date must be YYYY-MM-DD" });
      if (odometer_km_raw == null || String(odometer_km_raw).trim() === "") {
        return reply.code(400).send({ ok: false, error: "odometer_km is required" });
      }
      const odometer_km = Number(odometer_km_raw);
      if (!Number.isFinite(odometer_km) || odometer_km < 0) {
        return reply.code(400).send({ ok: false, error: "odometer_km must be a valid number >= 0" });
      }

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name
        FROM assets
        WHERE UPPER(asset_code) = UPPER(?)
        LIMIT 1
      `).get(asset_code);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const existing = getLdvPrestartCheckRow(Number(asset.id), check_date);

      const previousOdometer =
        getSanitizedLdvBaselineKm(Number(asset.id), check_date, {
          excludeCheckId: existing?.id ? Number(existing.id) : 0,
        }) ??
        getLatestLdvOdometerKm(Number(asset.id), check_date, {
          excludeCheckId: existing?.id ? Number(existing.id) : 0,
        });
      const kmReviewNeeded =
        (previousOdometer != null && odometer_km < previousOdometer) ||
        isLdvOdometerOutlier(odometer_km, previousOdometer);

      const checklist = normalizeLdvPrestartChecklist(checklistObj);
      const failed = checklist.filter((c) => !c.ok);
      if (failed.length) {
        return reply.code(400).send({
          ok: false,
          error: `Complete all pre-start checks before starting (${failed.map((c) => c.label).join(", ")}).`,
        });
      }

      const checklistJson = JSON.stringify(
        checklist.reduce((acc, c) => {
          acc[c.key] = Boolean(c.ok);
          return acc;
        }, {})
      );
      const reviewNote = kmReviewNeeded ? "KM flagged for supervisor review" : null;
      const mergedNotes = [notes, reviewNote].filter(Boolean).join(" | ") || null;

      let checkId = 0;
      if (existing?.id) {
        checkId = Number(existing.id);
        applyLdvKmToAllChecksForDate(
          Number(asset.id),
          check_date,
          odometer_km,
          inspector_name,
          mergedNotes,
          checklistJson
        );
      } else {
        const ins = db.prepare(`
          INSERT INTO vehicle_ldv_checks (
            asset_id, uuid, site_code, check_date, vehicle_registration, odometer_km, inspector_name, notes, check_mode, checklist_json, updated_at
          )
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'prestart', ?, datetime('now'))
        `).run(
          Number(asset.id),
          crypto.randomUUID(),
          site_code,
          check_date,
          String(asset.asset_code || ""),
          odometer_km,
          inspector_name,
          mergedNotes,
          checklistJson
        );
        checkId = Number(ins.lastInsertRowid);
      }

      const dailySync = syncLdvPrestartToDailyHours(
        Number(asset.id),
        check_date,
        odometer_km,
        inspector_name,
        previousOdometer,
        { unusual_km: kmReviewNeeded }
      );

      return reply.send({
        ok: true,
        id: checkId,
        asset_code: String(asset.asset_code || ""),
        check_date,
        odometer_km: Number(odometer_km.toFixed(1)),
        previous_odometer_km: previousOdometer == null ? null : Number(previousOdometer.toFixed(1)),
        km_review_needed: kmReviewNeeded,
        daily_input_sync: dailySync || { synced: false },
        message: kmReviewNeeded
          ? "Pre-start saved. KM looks unusual — daily input not updated until a supervisor reviews."
          : "Pre-start captured. KM reading saved to IRONLOG.",
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // POST /api/maintenance/vehicle-ldv-checks/prestart-correction — fix wrong KM on LDV + daily hours (supervisor)
  app.post("/vehicle-ldv-checks/prestart-correction", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;

      const asset_code = String(req.body?.asset_code || "").trim().toUpperCase();
      const work_date = String(req.body?.work_date || req.body?.check_date || "").trim()
        || new Date().toISOString().slice(0, 10);
      const closing_km = Number(req.body?.closing_km ?? req.body?.correct_odometer_km ?? req.body?.odometer_km);
      const opening_km_raw = req.body?.opening_km;
      const inspector_name = String(req.body?.inspector_name || "Supervisor correction").trim() || "Supervisor correction";
      const correction_note = String(req.body?.notes || "Supervisor KM correction").trim() || "Supervisor KM correction";
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      if (!isLdvPrestartAssetCode(asset_code)) {
        return reply.code(400).send({ ok: false, error: "Asset is not an LDV pre-start code (V01–V15)" });
      }
      if (!isDate(work_date)) return reply.code(400).send({ ok: false, error: "work_date must be YYYY-MM-DD" });
      if (!Number.isFinite(closing_km) || closing_km < 0) {
        return reply.code(400).send({ ok: false, error: "closing_km must be a valid number >= 0" });
      }

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name FROM assets WHERE UPPER(asset_code) = UPPER(?) LIMIT 1
      `).get(asset_code);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const assetId = Number(asset.id);
      const existing = getLdvPrestartCheckRow(assetId, work_date);

      let checkId = Number(existing?.id || 0);
      const opening_km = resolveLdvCorrectionOpeningKm(
        assetId,
        work_date,
        closing_km,
        checkId,
        opening_km_raw
      );
      if (!Number.isFinite(opening_km) || opening_km < 0) {
        return reply.code(400).send({ ok: false, error: "opening_km could not be resolved — pass opening_km explicitly" });
      }
      if (closing_km < opening_km) {
        return reply.code(400).send({
          ok: false,
          error: `Closing KM (${closing_km}) cannot be less than opening KM (${opening_km}). Enter opening KM manually if needed.`,
          opening_km,
        });
      }

      const previousOdometer = getLatestLdvOdometerKm(assetId, work_date, {
        excludeCheckId: checkId,
      });

      const checklistJsonOut = (() => {
        if (existing?.checklist_json && String(existing.checklist_json).trim()) {
          return String(existing.checklist_json);
        }
        return JSON.stringify(
          normalizeLdvPrestartChecklist({}).reduce((acc, c) => {
            acc[c.key] = true;
            return acc;
          }, {})
        );
      })();

      const purged = purgeLdvPoisonedKmBaselines(assetId, work_date, closing_km);

      const hasAnyCheckRow = Boolean(
        db.prepare(`SELECT id FROM vehicle_ldv_checks WHERE asset_id = ? AND check_date = ? LIMIT 1`).get(
          assetId,
          work_date
        )
      );

      let canonical;
      if (hasAnyCheckRow) {
        canonical = applyLdvKmToAllChecksForDate(
          assetId,
          work_date,
          closing_km,
          inspector_name,
          correction_note,
          checklistJsonOut
        );
      } else {
        const ins = db.prepare(`
          INSERT INTO vehicle_ldv_checks (
            asset_id, uuid, site_code, check_date, vehicle_registration, odometer_km,
            inspector_name, notes, check_mode, checklist_json, updated_at
          )
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'prestart', ?, datetime('now'))
        `).run(
          assetId,
          crypto.randomUUID(),
          site_code,
          work_date,
          String(asset.asset_code || ""),
          closing_km,
          inspector_name,
          correction_note,
          checklistJsonOut
        );
        canonical = getLdvPrestartCheckRow(assetId, work_date) || { id: Number(ins.lastInsertRowid) };
      }
      checkId = Number(canonical?.id || checkId || 0);

      const run_km = Math.max(0, closing_km - opening_km);
      const hasInputUnit = hasColumn("daily_hours", "input_unit");
      const dailyNote = `Supervisor KM correction | ${correction_note}`;
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
            is_used = 1,
            operator = excluded.operator,
            notes = excluded.notes
        `).run(assetId, work_date, opening_km, closing_km, run_km, inspector_name, dailyNote);
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
        `).run(assetId, work_date, opening_km, closing_km, run_km, inspector_name, dailyNote);
      }

      if (hasColumn("daily_hours", "input_unit")) {
        db.prepare(`
          INSERT INTO asset_input_units (asset_id, input_unit, updated_at)
          VALUES (?, 'km', datetime('now'))
          ON CONFLICT(asset_id) DO UPDATE SET input_unit = 'km', updated_at = datetime('now')
        `).run(assetId);
      }

      const priorBaseline =
        getSanitizedLdvBaselineKm(assetId, work_date, { excludeCheckId: checkId }) ??
        getPriorLdvPrestartOdometerKm(assetId, work_date) ??
        (previousOdometer != null ? previousOdometer : null);

      return reply.send({
        ok: true,
        asset_code,
        work_date,
        check_id: checkId,
        opening_km: Number(opening_km.toFixed(1)),
        closing_km: Number(closing_km.toFixed(1)),
        run_km: Number(run_km.toFixed(1)),
        previous_odometer_km: priorBaseline == null ? null : Number(priorBaseline.toFixed(1)),
        purged_baselines: purged,
        message: `KM corrected for ${asset_code} on ${work_date}.`,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // Machine-group prestart (excavator, dozer, etc.) — same storage as LDV checks, different check_mode; no daily_hours sync.

  // GET /api/maintenance/prestart/daily-summary?date=YYYY-MM-DD
  app.get("/prestart/daily-summary", async (req, reply) => {
    try {
      const date = String(req.query?.date || "").trim() || new Date().toISOString().slice(0, 10);
      if (!isDate(date)) return reply.code(400).send({ ok: false, error: "date must be YYYY-MM-DD" });
      const summary = listDailyPrestarts(db, date);
      const production = prestartDeductionForProductionFleet(db, date);
      return reply.send({
        ok: true,
        date,
        ...summary,
        production_deduction: production,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.get("/machine-prestart/context", async (req, reply) => {
    try {
      const asset_code = String(req.query?.asset_code || "").trim().toUpperCase();
      const check_date = String(req.query?.check_date || "").trim() || new Date().toISOString().slice(0, 10);
      if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "check_date must be YYYY-MM-DD" });

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name, category
        FROM assets
        WHERE UPPER(asset_code) = UPPER(?)
        LIMIT 1
      `).get(asset_code);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const profileId = resolveMachinePrestartProfile(asset.category, asset.asset_name, asset.asset_code);
      if (!profileId) {
        return reply.code(404).send({
          ok: false,
          error: "No machine pre-start template for this asset. Set category or name (e.g. Excavator, Haul truck).",
        });
      }
      const template = getMachinePrestartTemplate(profileId);
      const mode = machinePrestartCheckMode(profileId);
      if (!template || !mode) return reply.code(500).send({ ok: false, error: "Template resolution failed" });

      const existing = getMachinePrestartCheckRow(Number(asset.id), check_date, mode);

      const previous_smu_hours = getLatestMachineSmuHours(Number(asset.id), check_date, {
        excludeCheckId: existing?.id ? Number(existing.id) : 0,
      });

      let checklist = normalizeMachinePrestartChecklist(profileId, {});
      if (existing?.checklist_json) {
        try {
          const parsed = JSON.parse(String(existing.checklist_json || "{}"));
          if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
            checklist = normalizeMachinePrestartChecklist(profileId, parsed);
          }
        } catch {}
      }

      return reply.send({
        ok: true,
        profile_id: profileId,
        check_mode: mode,
        template,
        asset: {
          id: Number(asset.id),
          asset_code: String(asset.asset_code || ""),
          asset_name: String(asset.asset_name || ""),
          category: String(asset.category || ""),
        },
        check_date,
        previous_smu_hours,
        previous_smu_source:
          previous_smu_hours != null ? "daily_hours_or_prior_prestart" : null,
        existing_check: existing
          ? {
              id: Number(existing.id),
              smu_hours: existing.smu_hours == null ? null : Number(existing.smu_hours),
              inspector_name: existing.inspector_name || "",
              notes: existing.notes || "",
              checklist,
            }
          : null,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/machine-prestart", async (req, reply) => {
    try {
      const asset_code = String(req.body?.asset_code || "").trim().toUpperCase();
      const check_date = String(req.body?.check_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      const checklistObj = req.body?.checklist || {};
      const smu_raw = req.body?.smu_hours;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "check_date must be YYYY-MM-DD" });

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name, category
        FROM assets
        WHERE UPPER(asset_code) = UPPER(?)
        LIMIT 1
      `).get(asset_code);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const profileId = resolveMachinePrestartProfile(asset.category, asset.asset_name, asset.asset_code);
      if (!profileId) {
        return reply.code(400).send({
          ok: false,
          error: "No machine pre-start template for this asset category/name.",
        });
      }
      const mode = machinePrestartCheckMode(profileId);
      if (!mode) return reply.code(500).send({ ok: false, error: "Template resolution failed" });

      let smu_hours = null;
      if (smu_raw != null && String(smu_raw).trim() !== "") {
        const smu = Number(smu_raw);
        if (!Number.isFinite(smu) || smu < 0) {
          return reply.code(400).send({ ok: false, error: "smu_hours must be a valid number >= 0 when provided." });
        }
        smu_hours = Number(smu.toFixed(1));
      }

      const checklist = normalizeMachinePrestartChecklist(profileId, checklistObj);
      const failed = checklist.filter((c) => !c.ok);
      if (failed.length) {
        return reply.code(400).send({
          ok: false,
          error: `Complete all checks before submitting (${failed.map((c) => c.label).join(", ")}).`,
        });
      }

      const checklistJson = JSON.stringify(checklistToJsonObject(checklist));

      const existing = getMachinePrestartCheckRow(Number(asset.id), check_date, mode);

      let checkId = 0;
      if (existing?.id) {
        checkId = Number(existing.id);
        db.prepare(`
          UPDATE vehicle_ldv_checks
          SET
            inspector_name = ?,
            notes = ?,
            checklist_json = ?,
            smu_hours = ?,
            odometer_km = NULL,
            check_mode = ?,
            updated_at = datetime('now')
          WHERE id = ?
        `).run(inspector_name, notes, checklistJson, smu_hours, mode, checkId);
      } else {
        const ins = db.prepare(`
          INSERT INTO vehicle_ldv_checks (
            asset_id, uuid, site_code, check_date, vehicle_registration, odometer_km,
            inspector_name, notes, check_mode, checklist_json, smu_hours, updated_at
          )
          VALUES (?, ?, ?, ?, ?, NULL, ?, ?, ?, ?, ?, datetime('now'))
        `).run(
          Number(asset.id),
          crypto.randomUUID(),
          site_code,
          check_date,
          String(asset.asset_code || ""),
          inspector_name,
          notes,
          mode,
          checklistJson,
          smu_hours
        );
        checkId = Number(ins.lastInsertRowid);
      }

      let dailySync = { synced: false };
      if (smu_hours != null) {
        const previousSmu = getLatestMachineSmuHours(Number(asset.id), check_date, {
          excludeCheckId: checkId,
        });
        dailySync = syncMachinePrestartToDailyHours(
          Number(asset.id),
          check_date,
          smu_hours,
          inspector_name,
          previousSmu
        );
      }

      return reply.send({
        ok: true,
        id: checkId,
        asset_code: String(asset.asset_code || ""),
        check_date,
        profile_id: profileId,
        check_mode: mode,
        smu_hours,
        daily_input_sync: dailySync,
        message: "Machine pre-start saved to IRONLOG.",
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // POST /api/maintenance/machine-prestart/hours-correction — fix wrong SMU + daily hours (supervisor)
  app.post("/machine-prestart/hours-correction", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, ["admin", "supervisor", "plant_manager", "site_manager"])) return;

      const asset_code = String(req.body?.asset_code || "").trim().toUpperCase();
      const work_date = String(req.body?.work_date || req.body?.check_date || "").trim()
        || new Date().toISOString().slice(0, 10);
      const closing_hours = Number(
        req.body?.closing_hours ?? req.body?.correct_smu_hours ?? req.body?.smu_hours
      );
      const opening_hours_raw = req.body?.opening_hours;
      const inspector_name = String(req.body?.inspector_name || "Supervisor correction").trim() || "Supervisor correction";
      const correction_note = String(req.body?.notes || "Supervisor hours correction").trim() || "Supervisor hours correction";
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
      if (!isDate(work_date)) return reply.code(400).send({ ok: false, error: "work_date must be YYYY-MM-DD" });
      if (!Number.isFinite(closing_hours) || closing_hours < 0) {
        return reply.code(400).send({ ok: false, error: "closing_hours must be a valid number >= 0" });
      }

      const asset = db.prepare(`
        SELECT id, asset_code, asset_name, category
        FROM assets
        WHERE UPPER(asset_code) = UPPER(?)
        LIMIT 1
      `).get(asset_code);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const profileId = resolveMachinePrestartProfile(asset.category, asset.asset_name, asset.asset_code);
      if (!profileId) {
        return reply.code(400).send({ ok: false, error: "No machine pre-start template for this asset." });
      }
      const mode = machinePrestartCheckMode(profileId);
      if (!mode) return reply.code(500).send({ ok: false, error: "Template resolution failed" });

      const assetId = Number(asset.id);
      const existing = getMachinePrestartCheckRow(assetId, work_date, mode);

      let checkId = Number(existing?.id || 0);
      const previousSmu = getLatestMachineSmuHours(assetId, work_date, {
        excludeCheckId: checkId,
      });
      const opening_hours =
        opening_hours_raw != null && String(opening_hours_raw).trim() !== ""
          ? Number(opening_hours_raw)
          : previousSmu;
      if (!Number.isFinite(opening_hours) || opening_hours < 0) {
        return reply.code(400).send({
          ok: false,
          error: "opening_hours could not be resolved — pass opening_hours explicitly",
        });
      }
      if (closing_hours < opening_hours) {
        return reply.code(400).send({
          ok: false,
          error: `Closing hours (${closing_hours}) cannot be less than opening hours (${opening_hours}).`,
          previous_smu_hours: previousSmu,
          opening_hours,
        });
      }

      const checklistJsonOut = (() => {
        if (existing?.checklist_json && String(existing.checklist_json).trim()) {
          return String(existing.checklist_json);
        }
        const checklist = normalizeMachinePrestartChecklist(
          profileId,
          Object.fromEntries(
            normalizeMachinePrestartChecklist(profileId, {}).map((c) => [c.key, true])
          )
        );
        return JSON.stringify(checklistToJsonObject(checklist));
      })();

      if (checkId > 0) {
        db.prepare(`
          UPDATE vehicle_ldv_checks
          SET smu_hours = ?, inspector_name = ?, notes = ?, checklist_json = COALESCE(NULLIF(checklist_json, ''), ?),
              check_mode = ?, updated_at = datetime('now')
          WHERE id = ?
        `).run(closing_hours, inspector_name, correction_note, checklistJsonOut, mode, checkId);
      } else {
        const ins = db.prepare(`
          INSERT INTO vehicle_ldv_checks (
            asset_id, uuid, site_code, check_date, vehicle_registration, odometer_km,
            inspector_name, notes, check_mode, checklist_json, smu_hours, updated_at
          )
          VALUES (?, ?, ?, ?, ?, NULL, ?, ?, ?, ?, ?, datetime('now'))
        `).run(
          assetId,
          crypto.randomUUID(),
          site_code,
          work_date,
          String(asset.asset_code || ""),
          inspector_name,
          correction_note,
          mode,
          checklistJsonOut,
          closing_hours
        );
        checkId = Number(ins.lastInsertRowid);
      }

      const run_hours = upsertMachineDailyHoursCorrection(
        assetId,
        work_date,
        opening_hours,
        closing_hours,
        inspector_name,
        correction_note
      );

      return reply.send({
        ok: true,
        asset_code,
        work_date,
        check_id: checkId,
        profile_id: profileId,
        opening_hours: Number(opening_hours.toFixed(1)),
        closing_hours: Number(closing_hours.toFixed(1)),
        run_hours: Number(run_hours.toFixed(1)),
        previous_smu_hours: previousSmu == null ? null : Number(previousSmu.toFixed(1)),
        message: `Hours corrected for ${asset_code} on ${work_date}.`,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/vehicle-ldv-checks", async (req, reply) => {
    try {
      const asset_id = Number(req.body?.asset_id || 0);
      const check_date = String(req.body?.check_date || "").trim() || new Date().toISOString().slice(0, 10);
      const vehicle_registration = String(req.body?.vehicle_registration || "").trim() || null;
      const odometer_km = req.body?.odometer_km != null && req.body?.odometer_km !== "" ? Number(req.body.odometer_km) : null;
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "check_date must be YYYY-MM-DD" });

      const asset = db.prepare(`SELECT id FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const ins = db.prepare(`
        INSERT INTO vehicle_ldv_checks (
          asset_id, uuid, site_code, check_date, vehicle_registration, odometer_km, inspector_name, notes, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(asset_id, crypto.randomUUID(), site_code, check_date, vehicle_registration, odometer_km, inspector_name, notes);

      return reply.send({ ok: true, id: Number(ins.lastInsertRowid) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/vehicle-ldv-checks/:id/photo", async (req, reply) => {
    try {
      const checkId = Number(req.params?.id || 0);
      if (!checkId) return reply.code(400).send({ ok: false, error: "Invalid check id" });

      const row = db.prepare(`SELECT id FROM vehicle_ldv_checks WHERE id = ?`).get(checkId);
      if (!row) return reply.code(404).send({ ok: false, error: "Vehicle check not found" });

      const part = await req.file();
      if (!part) return reply.code(400).send({ ok: false, error: "Upload file field named 'file'" });

      const rawBuffer = await part.toBuffer();
      let fileBuffer;
      try {
        fileBuffer = await normalizeUploadedPhoto(rawBuffer);
      } catch {
        return reply.code(400).send({ ok: false, error: "Could not process image file" });
      }

      const safe = `ldv_${checkId}_${Date.now()}_${Math.floor(Math.random() * 100000)}.jpg`;
      const absPath = path.join(vehicleLdvcDir, safe);
      await fs.promises.writeFile(absPath, fileBuffer);

      const caption = String(req.query?.caption || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const relPath = path.join("uploads", "vehicle-ldv-checks", safe).replace(/\\/g, "/");

      const ins = db.prepare(`
        INSERT INTO vehicle_ldv_check_photos (check_id, uuid, site_code, file_path, caption, markers_json, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(checkId, crypto.randomUUID(), site_code, relPath, caption, null);

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        check_id: checkId,
        file_path: `/${relPath}`,
        caption,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.patch("/vehicle-ldv-checks/photos/:photoId", async (req, reply) => {
    try {
      const photoId = Number(req.params?.photoId || 0);
      if (!photoId) return reply.code(400).send({ ok: false, error: "Invalid photo id" });

      const photo = db.prepare(`SELECT id, check_id FROM vehicle_ldv_check_photos WHERE id = ?`).get(photoId);
      if (!photo) return reply.code(404).send({ ok: false, error: "Photo not found" });

      const body = req.body || {};
      const markers = body.markers;
      const caption = body.caption != null ? String(body.caption).trim() || null : undefined;

      if (markers != null) {
        if (!Array.isArray(markers)) return reply.code(400).send({ ok: false, error: "markers must be an array" });
        if (markers.length > 80) return reply.code(400).send({ ok: false, error: "too many markers (max 80)" });
        for (const m of markers) {
          const x = Number(m?.x);
          const y = Number(m?.y);
          if (!Number.isFinite(x) || !Number.isFinite(y) || x < 0 || x > 1 || y < 0 || y > 1) {
            return reply.code(400).send({ ok: false, error: "each marker needs x,y in [0,1] (fraction of image)" });
          }
        }
        const json = JSON.stringify(
          markers.map((m) => ({
            x: Number(m.x),
            y: Number(m.y),
            label: m.label != null ? String(m.label).slice(0, 120) : "",
            note: m.note != null ? String(m.note).slice(0, 500) : "",
          }))
        );
        db.prepare(`
          UPDATE vehicle_ldv_check_photos
          SET markers_json = ?, updated_at = datetime('now')
          WHERE id = ?
        `).run(json, photoId);
      }
      if (caption !== undefined) {
        db.prepare(`
          UPDATE vehicle_ldv_check_photos
          SET caption = ?, updated_at = datetime('now')
          WHERE id = ?
        `).run(caption, photoId);
      }

      const updated = db.prepare(`SELECT id, check_id, file_path, caption, markers_json FROM vehicle_ldv_check_photos WHERE id = ?`).get(photoId);
      let markersOut = [];
      try {
        markersOut = updated.markers_json ? JSON.parse(updated.markers_json) : [];
      } catch {
        markersOut = [];
      }

      return reply.send({
        ok: true,
        photo: {
          id: Number(updated.id),
          check_id: Number(updated.check_id),
          file_path: `/${String(updated.file_path || "").replace(/\\/g, "/")}`,
          caption: updated.caption,
          markers: markersOut,
        },
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.get("/damage-reports", async (req, reply) => {
    try {
      const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
      const assetId = Number(req.query?.asset_id || 0);
      const start = String(req.query?.start || "").trim();
      const end = String(req.query?.end || "").trim();
      const responsiblePerson = String(req.query?.responsible_person || "").trim();
      const pendingInvestigationRaw = String(req.query?.pending_investigation || "").trim();
      const hseReportAvailableRaw = String(req.query?.hse_report_available || "").trim();
      const params = [];
      const where = [];
      if (assetId > 0) {
        where.push("dr.asset_id = ?");
        params.push(assetId);
      }
      if (isDate(start)) {
        where.push("dr.report_date >= ?");
        params.push(start);
      }
      if (isDate(end)) {
        where.push("dr.report_date <= ?");
        params.push(end);
      }
      if (responsiblePerson) {
        where.push("UPPER(COALESCE(dr.responsible_person, '')) LIKE UPPER(?)");
        params.push(`%${responsiblePerson}%`);
      }
      if (pendingInvestigationRaw === "0" || pendingInvestigationRaw === "1") {
        where.push("COALESCE(dr.pending_investigation, 0) = ?");
        params.push(Number(pendingInvestigationRaw));
      }
      if (hseReportAvailableRaw === "0" || hseReportAvailableRaw === "1") {
        where.push("COALESCE(dr.hse_report_available, 0) = ?");
        params.push(Number(hseReportAvailableRaw));
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
          dr.created_at,
          a.asset_code,
          a.asset_name
        FROM manager_damage_reports dr
        JOIN assets a ON a.id = dr.asset_id
        ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
        ORDER BY dr.report_date DESC, dr.id DESC
      `).all(...params);

      const ids = rows.map((r) => Number(r.id)).filter((n) => n > 0);
      let photosByReport = new Map();
      if (ids.length) {
        const marks = ids.map(() => "?").join(",");
        const photos = db.prepare(`
          SELECT
            id,
            ${dmgPhotoReportCol} AS damage_report_id,
            ${dmgPhotoPathCol} AS file_path,
            ${dmgPhotoCaptionCol} AS caption,
            ${dmgPhotoCreatedCol} AS created_at
          FROM manager_damage_report_photos
          WHERE ${dmgPhotoReportCol} IN (${marks})
          ORDER BY id ASC
        `).all(...ids);
        photosByReport = photos.reduce((m, p) => {
          const k = Number(p.damage_report_id);
          if (!m.has(k)) m.set(k, []);
          m.get(k).push(p);
          return m;
        }, new Map());
      }

      return reply.send({
        ok: true,
        rows: rows.map((r) => ({
          ...r,
          photos: photosByReport.get(Number(r.id)) || [],
        })),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/damage-reports", async (req, reply) => {
    try {
      const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
      const asset_id = Number(req.body?.asset_id || 0);
      const report_date = String(req.body?.report_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const hour_meter_raw = req.body?.hour_meter;
      const hour_meter =
        hour_meter_raw == null || String(hour_meter_raw).trim() === ""
          ? null
          : Number(hour_meter_raw);
      const damage_location = String(req.body?.damage_location || "").trim() || null;
      const severity = String(req.body?.severity || "").trim() || null;
      const damage_description = String(req.body?.damage_description || "").trim() || null;
      const immediate_action = String(req.body?.immediate_action || "").trim() || null;
      const out_of_service = Number(req.body?.out_of_service || 0) ? 1 : 0;
      const damage_time = String(req.body?.damage_time || "").trim() || null;
      const responsible_person = String(req.body?.responsible_person || "").trim() || null;
      const pending_investigation = Number(req.body?.pending_investigation || 0) ? 1 : 0;
      const hse_report_available = Number(req.body?.hse_report_available || 0) ? 1 : 0;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";

      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(report_date)) return reply.code(400).send({ ok: false, error: "report_date must be YYYY-MM-DD" });
      if (hour_meter != null && !Number.isFinite(hour_meter)) {
        return reply.code(400).send({ ok: false, error: "hour_meter must be numeric" });
      }
      if (!damage_location) return reply.code(400).send({ ok: false, error: "damage_location is required" });
      if (!severity) return reply.code(400).send({ ok: false, error: "severity is required" });
      if (!damage_description) return reply.code(400).send({ ok: false, error: "damage_description is required" });
      if (!immediate_action) return reply.code(400).send({ ok: false, error: "immediate_action is required" });
      if (damage_time && !/^\d{2}:\d{2}$/.test(damage_time)) {
        return reply.code(400).send({ ok: false, error: "damage_time must be HH:MM" });
      }

      const asset = db.prepare(`SELECT id FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const ins = db.prepare(`
        INSERT INTO manager_damage_reports (
          asset_id, uuid, site_code, report_date, ${drInspectorCol}, hour_meter,
          damage_location, severity, damage_description, immediate_action, out_of_service,
          damage_time, responsible_person, pending_investigation, hse_report_available, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
      `).run(
        asset_id,
        crypto.randomUUID(),
        site_code,
        report_date,
        inspector_name,
        hour_meter,
        damage_location,
        severity,
        damage_description,
        immediate_action,
        out_of_service,
        damage_time,
        responsible_person,
        pending_investigation,
        hse_report_available
      );

      return reply.send({ ok: true, id: Number(ins.lastInsertRowid) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/damage-reports/:id/photo", async (req, reply) => {
    try {
      const reportId = Number(req.params?.id || 0);
      if (!reportId) return reply.code(400).send({ ok: false, error: "Invalid damage report id" });

      const report = db.prepare(`SELECT id FROM manager_damage_reports WHERE id = ?`).get(reportId);
      if (!report) return reply.code(404).send({ ok: false, error: "Damage report not found" });

      const part = await req.file();
      if (!part) return reply.code(400).send({ ok: false, error: "Upload file field named 'file'" });

      const extRaw = path.extname(part.filename || "").toLowerCase();
      const ext = [".jpg", ".jpeg", ".png", ".webp"].includes(extRaw) ? extRaw : ".jpg";
      const safe = `mdr_${reportId}_${Date.now()}_${Math.floor(Math.random() * 100000)}${ext}`;
      const absPath = path.join(damageReportsDir, safe);
      await fs.promises.writeFile(absPath, await part.toBuffer());

      const caption = String(req.query?.caption || "").trim() || null;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const relPath = path.join("uploads", "manager-damage-reports", safe).replace(/\\/g, "/");

      const hasReportId = hasColumn("manager_damage_report_photos", "damage_report_id");
      const hasLegacyReportId = hasColumn("manager_damage_report_photos", "manager_damage_report_id");
      const linkCols = [];
      const linkVals = [];
      if (hasReportId) {
        linkCols.push("damage_report_id");
        linkVals.push(reportId);
      }
      if (hasLegacyReportId) {
        linkCols.push("manager_damage_report_id");
        linkVals.push(reportId);
      }
      if (!linkCols.length) {
        linkCols.push(dmgPhotoReportCol);
        linkVals.push(reportId);
      }

      const hasImageData = hasColumn("manager_damage_report_photos", "image_data");
      const insertCols = [...linkCols, "uuid", "site_code", "file_path", ...(hasImageData ? ["image_data"] : []), "caption", "updated_at"];
      const placeholders = [...insertCols.map((c) => (c === "updated_at" ? "datetime('now')" : "?"))].join(", ");
      const ins = db.prepare(`
        INSERT INTO manager_damage_report_photos (${insertCols.join(", ")})
        VALUES (${placeholders})
      `).run(...linkVals, crypto.randomUUID(), site_code, relPath, ...(hasImageData ? [relPath] : []), caption);

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        damage_report_id: reportId,
        file_path: `/${relPath}`,
        caption,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  const PARTS_REQUEST_STATUSES = new Set(["requested", "ordered", "received", "cancelled"]);
  const PARTS_REQUEST_URGENCY = new Set(["normal", "urgent", "critical"]);
  const PARTS_REQUEST_MANAGERS = ["admin", "supervisor", "stores", "plant_manager", "site_manager", "workshop_manager"];

  // GET /api/maintenance/parts-requests/rfq.pdf?ids=1,2,3&supplier=...&reference=...
  app.get("/parts-requests/rfq.pdf", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const idsRaw = req.query?.ids;
      const requestIds = Array.isArray(idsRaw)
        ? [...new Set(idsRaw.map((v) => Number(v)).filter((v) => Number.isInteger(v) && v > 0))]
        : [...new Set(String(idsRaw || "").split(/[,\s]+/).map((v) => Number(v)).filter((v) => Number.isInteger(v) && v > 0))];
      if (!requestIds.length) {
        return reply.code(400).send({ ok: false, error: "Select at least one parts request (ids)" });
      }

      const supplier = String(req.query?.supplier || "").trim();
      if (!supplier) {
        return reply.code(400).send({ ok: false, error: "supplier is required" });
      }

      const reference = String(req.query?.reference || "").trim() || `RFQ-${new Date().toISOString().slice(0, 10)}`;
      const contact = String(req.query?.contact || "").trim();
      const email = String(req.query?.email || "").trim();
      const phone = String(req.query?.phone || "").trim();
      const required_by = String(req.query?.required_by || "").trim();
      const notes = String(req.query?.notes || "").trim();
      const requested_by = String(req.query?.requested_by || req.headers?.["x-user-name"] || "").trim() || "system";
      const asOfLabel = new Date().toISOString().slice(0, 10);
      const branding = getPdfReportBranding(db);

      const placeholders = requestIds.map(() => "?").join(", ");
      const rows = db.prepare(`
        SELECT
          pr.id,
          pr.asset_code,
          pr.part_code,
          pr.part_name,
          pr.qty,
          pr.urgency,
          pr.notes,
          pr.work_order_id,
          pr.status,
          pr.requested_by,
          pr.created_at,
          a.asset_name
        FROM maintenance_parts_requests pr
        LEFT JOIN assets a ON a.id = pr.asset_id
        WHERE pr.id IN (${placeholders})
          AND COALESCE(pr.site_code, 'main') = ?
        ORDER BY
          CASE LOWER(COALESCE(pr.urgency, 'normal'))
            WHEN 'critical' THEN 0
            WHEN 'urgent' THEN 1
            ELSE 2
          END ASC,
          pr.asset_code ASC,
          pr.part_code ASC,
          pr.id ASC
      `).all(...requestIds, site_code);

      if (!rows.length) {
        return reply.code(404).send({ ok: false, error: "No matching parts requests found" });
      }

      const urgencyLabel = (u) => {
        const x = String(u || "normal").toLowerCase();
        if (x === "critical") return "Critical";
        if (x === "urgent") return "Urgent";
        return "Normal";
      };

      const lineRows = rows.map((r, idx) => ({
        line_no: String(idx + 1),
        part_code: String(r.part_code || "—"),
        part_name: String(r.part_name || "—"),
        qty: Number(r.qty || 0).toFixed(1).replace(/\.0$/, ""),
        asset: [r.asset_code, r.asset_name].map((x) => String(x || "").trim()).filter(Boolean).join(" — ") || "—",
        work_order_id: r.work_order_id ? String(r.work_order_id) : "—",
        urgency: urgencyLabel(r.urgency),
        notes: String(r.notes || "").trim() || "—",
      }));

      const pdf = await buildPdfBuffer(
        (doc) => {
          sectionTitle(doc, "Request for Quote");
          doc.font("Helvetica").fontSize(10).fillColor("#0f172a");
          doc.text(`Reference: ${reference}`);
          doc.text(`Date: ${asOfLabel}`);
          if (branding.company_name) doc.text(`From: ${branding.company_name}${branding.site_name ? ` — ${branding.site_name}` : ""}`);
          doc.text(`Prepared by: ${requested_by}`);
          doc.moveDown(0.35);
          doc.font("Helvetica-Bold").text("To:");
          doc.font("Helvetica");
          doc.text(supplier);
          if (contact) doc.text(`Attention: ${contact}`);
          if (email) doc.text(`Email: ${email}`);
          if (phone) doc.text(`Phone: ${phone}`);
          if (required_by) doc.text(`Quote required by: ${required_by}`);
          doc.moveDown(0.5);
          doc.font("Helvetica").fontSize(10).text(
            "Please provide pricing, availability, and lead time for the parts listed below.",
          );
          doc.moveDown(0.4);

          table(
            doc,
            [
              { key: "line_no", label: "#", width: 0.04, align: "center" },
              { key: "part_code", label: "Part code", width: 0.12 },
              { key: "part_name", label: "Description", width: 0.24 },
              { key: "qty", label: "Qty", width: 0.06, align: "right" },
              { key: "asset", label: "Asset", width: 0.16 },
              { key: "work_order_id", label: "WO #", width: 0.07, align: "center" },
              { key: "urgency", label: "Urgency", width: 0.09 },
              { key: "notes", label: "Notes", width: 0.22 },
            ],
            lineRows,
          );

          if (notes) {
            doc.moveDown(0.6);
            sectionTitle(doc, "Additional notes / terms");
            doc.font("Helvetica").fontSize(10).text(notes, { width: doc.page.width - doc.page.margins.left - doc.page.margins.right });
          }

          doc.moveDown(1.2);
          doc.font("Helvetica").fontSize(10);
          doc.text("Authorized by: ________________________________     Date: ________________");
        },
        {
          title: branding.company_name || "IRONLOG",
          subtitle: "Parts Request for Quote",
          rightText: reference,
          layout: "landscape",
          db,
        },
      );

      const isDownload = String(req.query?.download || "").trim() === "1";
      const safeRef = reference.replace(/[^\w.-]+/g, "_");
      reply.header("Cache-Control", "no-store, no-cache, must-revalidate");
      reply.header("Pragma", "no-cache");
      reply.header("Content-Type", "application/pdf");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="${safeRef}.pdf"`,
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/parts-requests", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const status = String(req.query?.status || "").trim().toLowerCase();
      const mine = String(req.query?.mine || "").trim() === "1";
      const userName = String(req.headers?.["x-user-name"] || "").trim();
      const limitRaw = Number(req.query?.limit || 500);
      const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(2000, Math.trunc(limitRaw))) : 500;

      const where = ["COALESCE(pr.site_code, 'main') = ?"];
      const params = [site_code];
      if (status && PARTS_REQUEST_STATUSES.has(status)) {
        where.push("LOWER(COALESCE(pr.status, 'requested')) = ?");
        params.push(status);
      }
      if (mine && userName) {
        where.push("LOWER(COALESCE(pr.requested_by, '')) = ?");
        params.push(userName.toLowerCase());
      }

      const rows = db.prepare(`
        SELECT
          pr.id,
          pr.site_code,
          pr.asset_id,
          pr.asset_code,
          pr.part_code,
          pr.part_name,
          pr.qty,
          pr.urgency,
          pr.notes,
          pr.work_order_id,
          pr.status,
          pr.requested_by,
          pr.ordered_by,
          pr.status_notes,
          pr.created_at,
          pr.updated_at,
          a.asset_name
        FROM maintenance_parts_requests pr
        LEFT JOIN assets a ON a.id = pr.asset_id
        WHERE ${where.join(" AND ")}
        ORDER BY
          CASE LOWER(COALESCE(pr.status, 'requested'))
            WHEN 'requested' THEN 0
            WHEN 'ordered' THEN 1
            WHEN 'received' THEN 2
            ELSE 3
          END ASC,
          CASE LOWER(COALESCE(pr.urgency, 'normal'))
            WHEN 'critical' THEN 0
            WHEN 'urgent' THEN 1
            ELSE 2
          END ASC,
          pr.created_at DESC,
          pr.id DESC
        LIMIT ${limit}
      `).all(...params);
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/parts-requests", async (req, reply) => {
    try {
      const body = req.body || {};
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const requested_by = String(req.headers?.["x-user-name"] || body.requested_by || "system").trim() || "system";
      const part_code = String(body.part_code || "").trim();
      let part_name = String(body.part_name || "").trim();
      const qty = Number(body.qty ?? 1);
      const urgencyRaw = String(body.urgency || "normal").trim().toLowerCase();
      const urgency = PARTS_REQUEST_URGENCY.has(urgencyRaw) ? urgencyRaw : "normal";
      const notes = String(body.notes || "").trim();
      const work_order_id = Number(body.work_order_id || 0) || null;
      const asset_id = Number(body.asset_id || 0) || null;

      if (!part_code && !part_name) {
        return reply.code(400).send({ ok: false, error: "Part description or part code is required" });
      }
      if (!Number.isFinite(qty) || qty <= 0) {
        return reply.code(400).send({ ok: false, error: "qty must be greater than zero" });
      }

      if (part_code && !part_name) {
        const hit = db.prepare(`SELECT part_name FROM parts WHERE LOWER(part_code) = LOWER(?) LIMIT 1`).get(part_code);
        if (hit?.part_name) part_name = String(hit.part_name).trim();
      }
      if (!part_name) part_name = part_code;

      let asset_code = String(body.asset_code || "").trim();
      if (asset_id) {
        const asset = db.prepare(`SELECT id, asset_code FROM assets WHERE id = ? LIMIT 1`).get(asset_id);
        if (!asset) return reply.code(400).send({ ok: false, error: "Invalid asset" });
        asset_code = String(asset.asset_code || asset_code).trim();
      }

      const now = new Date().toISOString();
      const info = db.prepare(`
        INSERT INTO maintenance_parts_requests (
          site_code, asset_id, asset_code, part_code, part_name, qty, urgency, notes,
          work_order_id, status, requested_by, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'requested', ?, ?, ?)
      `).run(
        site_code,
        asset_id,
        asset_code || null,
        part_code || null,
        part_name,
        qty,
        urgency,
        notes || null,
        work_order_id,
        requested_by,
        now,
        now
      );

      writeAudit(db, req, {
        module: "maintenance",
        action: "parts_request.create",
        entity_type: "maintenance_parts_request",
        entity_id: String(info.lastInsertRowid),
        after: { part_code, part_name, qty, urgency, asset_code, work_order_id },
      });

      return reply.send({ ok: true, id: Number(info.lastInsertRowid || 0) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.patch("/parts-requests/:id/status", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });

      const body = req.body || {};
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const userName = String(req.headers?.["x-user-name"] || "").trim() || "system";
      const roles = getMaintenanceRoles(req);
      const isManager = roles.some((r) => PARTS_REQUEST_MANAGERS.includes(r));

      const status = String(body.status || "").trim().toLowerCase();
      if (!PARTS_REQUEST_STATUSES.has(status)) {
        return reply.code(400).send({ ok: false, error: "Invalid status" });
      }

      const existing = db.prepare(`
        SELECT id, status, requested_by
        FROM maintenance_parts_requests
        WHERE id = ? AND COALESCE(site_code, 'main') = ?
      `).get(id, site_code);
      if (!existing) return reply.code(404).send({ ok: false, error: "Request not found" });

      if (!isManager) {
        const own = String(existing.requested_by || "").trim().toLowerCase() === userName.toLowerCase();
        if (!own || status !== "cancelled" || String(existing.status || "").toLowerCase() !== "requested") {
          return reply.code(403).send({ ok: false, error: "not allowed" });
        }
      }

      const status_notes = String(body.status_notes || "").trim() || null;
      const ordered_by = ["ordered", "received"].includes(status) ? userName : null;
      const now = new Date().toISOString();

      db.prepare(`
        UPDATE maintenance_parts_requests
        SET status = ?, status_notes = ?, ordered_by = COALESCE(?, ordered_by), updated_at = ?
        WHERE id = ? AND COALESCE(site_code, 'main') = ?
      `).run(status, status_notes, ordered_by, now, id, site_code);

      writeAudit(db, req, {
        module: "maintenance",
        action: "parts_request.status",
        entity_type: "maintenance_parts_request",
        entity_id: String(id),
        after: { status, status_notes },
      });

      return reply.send({ ok: true, id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

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
    return { work_date, technician_name, asset_code, reason, hours, labor_rate_per_hour };
  }

  // GET /api/maintenance/mechanic-labor?date=YYYY-MM-DD
  app.get("/mechanic-labor", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, MECHANIC_LABOR_EDITORS)) return;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const date = String(req.query?.date || "").trim();
      if (!isDate(date)) {
        return reply.code(400).send({ ok: false, error: "date (YYYY-MM-DD) required" });
      }
      const defaultRate = readMechanicLaborDefaultRate();
      const rows = db.prepare(`
        SELECT id, work_date, technician_name, hours, asset_code, reason, labor_rate_per_hour, created_by, updated_at
        FROM mechanic_labor_entries
        WHERE work_date = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
        ORDER BY id ASC
      `).all(date, site_code).map((r) => enrichMechanicLaborRow(r, defaultRate));

      const totals = rows.reduce(
        (acc, r) => {
          acc.entries += 1;
          acc.hours += Number(r.hours || 0);
          acc.labor_cost += Number(r.labor_cost || 0);
          return acc;
        },
        { entries: 0, hours: 0, labor_cost: 0 },
      );
      totals.hours = Number(totals.hours.toFixed(2));
      totals.labor_cost = Number(totals.labor_cost.toFixed(2));

      return reply.send({ ok: true, date, default_labor_rate: defaultRate, rows, totals });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // PUT /api/maintenance/mechanic-labor/settings — default labor rate
  app.put("/mechanic-labor/settings", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, MECHANIC_LABOR_RATE_MANAGERS)) return;
      const rate = Math.max(0, Number(req.body?.labor_rate_per_hour || 0));
      if (!Number.isFinite(rate) || rate <= 0) {
        return reply.code(400).send({ ok: false, error: "labor_rate_per_hour must be > 0" });
      }
      db.prepare(`
        CREATE TABLE IF NOT EXISTS cost_settings (
          key TEXT PRIMARY KEY,
          value TEXT NOT NULL,
          updated_at TEXT NOT NULL DEFAULT (datetime('now'))
        )
      `).run();
      db.prepare(`
        INSERT INTO cost_settings (key, value, updated_at)
        VALUES ('labor_cost_per_hour_default', ?, datetime('now'))
        ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = datetime('now')
      `).run(String(rate));
      writeAudit(db, req, {
        module: "maintenance",
        action: "mechanic_labor.rate_default",
        entity_type: "cost_settings",
        entity_id: "labor_cost_per_hour_default",
        after: { labor_rate_per_hour: rate },
      });
      return reply.send({ ok: true, default_labor_rate: rate });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // POST /api/maintenance/mechanic-labor
  app.post("/mechanic-labor", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, MECHANIC_LABOR_EDITORS)) return;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const userName = String(req.headers?.["x-user-name"] || "").trim() || "system";
      const parsed = mechanicLaborEntryBody(req.body || {});
      if (!isDate(parsed.work_date)) {
        return reply.code(400).send({ ok: false, error: "work_date (YYYY-MM-DD) required" });
      }
      if (!parsed.technician_name) {
        return reply.code(400).send({ ok: false, error: "technician_name required" });
      }
      if (!parsed.asset_code) {
        return reply.code(400).send({ ok: false, error: "asset_code required" });
      }
      if (!parsed.reason) {
        return reply.code(400).send({ ok: false, error: "reason required" });
      }
      if (!Number.isFinite(parsed.hours) || parsed.hours <= 0) {
        return reply.code(400).send({ ok: false, error: "hours must be > 0" });
      }

      const info = db.prepare(`
        INSERT INTO mechanic_labor_entries (
          work_date, technician_name, hours, asset_code, reason,
          labor_rate_per_hour, site_code, created_by, updated_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).run(
        parsed.work_date,
        parsed.technician_name,
        parsed.hours,
        parsed.asset_code,
        parsed.reason,
        parsed.labor_rate_per_hour,
        site_code,
        userName,
        userName,
      );

      writeAudit(db, req, {
        module: "maintenance",
        action: "mechanic_labor.create",
        entity_type: "mechanic_labor_entry",
        entity_id: String(info.lastInsertRowid),
        after: parsed,
      });

      const defaultRate = readMechanicLaborDefaultRate();
      const row = enrichMechanicLaborRow(
        db.prepare(`SELECT * FROM mechanic_labor_entries WHERE id = ?`).get(info.lastInsertRowid),
        defaultRate,
      );
      return reply.send({ ok: true, row });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // POST /api/maintenance/mechanic-labor/batch — save many rows for one day
  app.post("/mechanic-labor/batch", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, MECHANIC_LABOR_EDITORS)) return;
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const userName = String(req.headers?.["x-user-name"] || "").trim() || "system";
      const body = req.body || {};
      const work_date = String(body.work_date || "").trim();
      const mode = String(body.mode || "append").trim().toLowerCase();
      const rawEntries = Array.isArray(body.entries) ? body.entries : [];

      if (!isDate(work_date)) {
        return reply.code(400).send({ ok: false, error: "work_date (YYYY-MM-DD) required" });
      }
      if (!rawEntries.length) {
        return reply.code(400).send({ ok: false, error: "entries array required" });
      }

      const validEntries = [];
      const errors = [];
      rawEntries.forEach((item, idx) => {
        const parsed = mechanicLaborEntryBody({ ...item, work_date });
        const line = idx + 1;
        if (!parsed.technician_name) errors.push(`Row ${line}: technician_name required`);
        if (!parsed.asset_code) errors.push(`Row ${line}: asset_code required`);
        if (!parsed.reason) errors.push(`Row ${line}: reason required`);
        if (!Number.isFinite(parsed.hours) || parsed.hours <= 0) errors.push(`Row ${line}: hours must be > 0`);
        if (
          !parsed.technician_name ||
          !parsed.asset_code ||
          !parsed.reason ||
          !Number.isFinite(parsed.hours) ||
          parsed.hours <= 0
        ) {
          return;
        }
        validEntries.push(parsed);
      });

      if (errors.length) {
        return reply.code(400).send({ ok: false, error: errors.slice(0, 8).join("; "), errors });
      }
      if (!validEntries.length) {
        return reply.code(400).send({ ok: false, error: "no valid entries to save" });
      }

      const insertStmt = db.prepare(`
        INSERT INTO mechanic_labor_entries (
          work_date, technician_name, hours, asset_code, reason,
          labor_rate_per_hour, site_code, created_by, updated_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);

      const tx = db.transaction(() => {
        let deleted = 0;
        if (mode === "replace") {
          const info = db.prepare(`
            DELETE FROM mechanic_labor_entries
            WHERE work_date = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
          `).run(work_date, site_code);
          deleted = Number(info.changes || 0);
        }
        const ids = [];
        for (const parsed of validEntries) {
          const info = insertStmt.run(
            work_date,
            parsed.technician_name,
            parsed.hours,
            parsed.asset_code,
            parsed.reason,
            parsed.labor_rate_per_hour,
            site_code,
            userName,
            userName,
          );
          ids.push(Number(info.lastInsertRowid || 0));
        }
        return { deleted, ids };
      });

      const result = tx();
      writeAudit(db, req, {
        module: "maintenance",
        action: "mechanic_labor.batch",
        entity_type: "mechanic_labor_entry",
        entity_id: work_date,
        payload: { mode, saved: validEntries.length, deleted: result.deleted },
      });

      const defaultRate = readMechanicLaborDefaultRate();
      const rows = result.ids
        .map((id) => db.prepare(`SELECT * FROM mechanic_labor_entries WHERE id = ?`).get(id))
        .filter(Boolean)
        .map((r) => enrichMechanicLaborRow(r, defaultRate));

      return reply.send({
        ok: true,
        work_date,
        mode,
        saved: validEntries.length,
        deleted: result.deleted,
        skipped: Math.max(0, rawEntries.length - validEntries.length),
        rows,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // PATCH /api/maintenance/mechanic-labor/:id
  app.patch("/mechanic-labor/:id", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, MECHANIC_LABOR_EDITORS)) return;
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const userName = String(req.headers?.["x-user-name"] || "").trim() || "system";
      const existing = db.prepare(`
        SELECT id FROM mechanic_labor_entries
        WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
      `).get(id, site_code);
      if (!existing) return reply.code(404).send({ ok: false, error: "Entry not found" });

      const parsed = mechanicLaborEntryBody(req.body || {});
      if (!isDate(parsed.work_date)) {
        return reply.code(400).send({ ok: false, error: "work_date (YYYY-MM-DD) required" });
      }
      if (!parsed.technician_name || !parsed.asset_code || !parsed.reason) {
        return reply.code(400).send({ ok: false, error: "technician_name, asset_code and reason required" });
      }
      if (!Number.isFinite(parsed.hours) || parsed.hours <= 0) {
        return reply.code(400).send({ ok: false, error: "hours must be > 0" });
      }

      db.prepare(`
        UPDATE mechanic_labor_entries
        SET
          work_date = ?,
          technician_name = ?,
          hours = ?,
          asset_code = ?,
          reason = ?,
          labor_rate_per_hour = ?,
          updated_by = ?,
          updated_at = datetime('now')
        WHERE id = ?
      `).run(
        parsed.work_date,
        parsed.technician_name,
        parsed.hours,
        parsed.asset_code,
        parsed.reason,
        parsed.labor_rate_per_hour,
        userName,
        id,
      );

      writeAudit(db, req, {
        module: "maintenance",
        action: "mechanic_labor.update",
        entity_type: "mechanic_labor_entry",
        entity_id: String(id),
        after: parsed,
      });

      const defaultRate = readMechanicLaborDefaultRate();
      const row = enrichMechanicLaborRow(
        db.prepare(`SELECT * FROM mechanic_labor_entries WHERE id = ?`).get(id),
        defaultRate,
      );
      return reply.send({ ok: true, row });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // DELETE /api/maintenance/mechanic-labor/:id
  app.delete("/mechanic-labor/:id", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, MECHANIC_LABOR_EDITORS)) return;
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const existing = db.prepare(`
        SELECT id FROM mechanic_labor_entries
        WHERE id = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
      `).get(id, site_code);
      if (!existing) return reply.code(404).send({ ok: false, error: "Entry not found" });
      db.prepare(`DELETE FROM mechanic_labor_entries WHERE id = ?`).run(id);
      writeAudit(db, req, {
        module: "maintenance",
        action: "mechanic_labor.delete",
        entity_type: "mechanic_labor_entry",
        entity_id: String(id),
      });
      return reply.send({ ok: true, id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // GET /api/maintenance/mechanic-labor.xlsx?year=2026
  app.get("/mechanic-labor.xlsx", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, MECHANIC_LABOR_EDITORS)) return;
      const year = String(req.query?.year || new Date().getFullYear()).trim();
      if (!/^\d{4}$/.test(year)) {
        return reply.code(400).send({ ok: false, error: "year (YYYY) required" });
      }
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const defaultRate = readMechanicLaborDefaultRate();
      const start = `${year}-01-01`;
      const end = `${year}-12-31`;

      const rawRows = db.prepare(`
        SELECT id, work_date, technician_name, hours, asset_code, reason, labor_rate_per_hour
        FROM mechanic_labor_entries
        WHERE work_date >= ? AND work_date <= ?
          AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
        ORDER BY work_date ASC, id ASC
      `).all(start, end, site_code);

      const rows = rawRows.map((r) => enrichMechanicLaborRow(r, defaultRate));

      const wb = new ExcelJS.Workbook();
      wb.creator = "IRONLOG";
      wb.created = new Date();

      const wsSummary = wb.addWorksheet("Summary");
      wsSummary.columns = [
        { header: "Field", key: "field", width: 28 },
        { header: "Value", key: "value", width: 24 },
      ];
      const yearHours = rows.reduce((s, r) => s + Number(r.hours || 0), 0);
      const yearCost = rows.reduce((s, r) => s + Number(r.labor_cost || 0), 0);
      wsSummary.addRows([
        { field: "Year", value: year },
        { field: "Site", value: site_code },
        { field: "Default labor rate ($/hr)", value: defaultRate },
        { field: "Total entries", value: rows.length },
        { field: "Total hours", value: Number(yearHours.toFixed(2)) },
        { field: "Total labor cost ($)", value: Number(yearCost.toFixed(2)) },
      ]);

      const monthCols = [
        { header: "Date", key: "work_date", width: 14 },
        { header: "Technician", key: "technician_name", width: 22 },
        { header: "Hours", key: "hours", width: 10 },
        { header: "Plant no", key: "asset_code", width: 14 },
        { header: "Reason", key: "reason", width: 36 },
        { header: "Rate ($/hr)", key: "labor_rate_per_hour", width: 12 },
        { header: "Labor cost ($)", key: "labor_cost", width: 14 },
      ];

      for (let m = 1; m <= 12; m += 1) {
        const monthKey = `${year}-${String(m).padStart(2, "0")}`;
        const sheetName = MECHANIC_MONTH_NAMES[m - 1];
        const ws = wb.addWorksheet(sheetName, { views: [{ state: "frozen", ySplit: 1 }] });
        ws.columns = monthCols;
        const monthRows = rows.filter((r) => String(r.work_date || "").startsWith(monthKey));
        ws.addRows(monthRows);
        if (!monthRows.length) {
          ws.addRow({
            work_date: "-",
            technician_name: "No entries",
            hours: 0,
            asset_code: "",
            reason: "",
            labor_rate_per_hour: defaultRate,
            labor_cost: 0,
          });
        }
        const monthHours = monthRows.reduce((s, r) => s + Number(r.hours || 0), 0);
        const monthCost = monthRows.reduce((s, r) => s + Number(r.labor_cost || 0), 0);
        ws.addRow({});
        ws.addRow({
          work_date: "TOTAL",
          technician_name: "",
          hours: Number(monthHours.toFixed(2)),
          asset_code: "",
          reason: "",
          labor_rate_per_hour: "",
          labor_cost: Number(monthCost.toFixed(2)),
        });
      }

      const byTech = new Map();
      const byPlant = new Map();
      for (const r of rows) {
        const tech = String(r.technician_name || "").trim() || "Unknown";
        const plant = String(r.asset_code || "").trim() || "Unknown";
        if (!byTech.has(tech)) byTech.set(tech, { technician_name: tech, hours: 0, labor_cost: 0, entries: 0 });
        if (!byPlant.has(plant)) byPlant.set(plant, { asset_code: plant, hours: 0, labor_cost: 0, entries: 0 });
        const t = byTech.get(tech);
        const p = byPlant.get(plant);
        t.entries += 1;
        t.hours += Number(r.hours || 0);
        t.labor_cost += Number(r.labor_cost || 0);
        p.entries += 1;
        p.hours += Number(r.hours || 0);
        p.labor_cost += Number(r.labor_cost || 0);
      }

      const wsTech = wb.addWorksheet("Technicians");
      wsTech.columns = [
        { header: "Technician", key: "technician_name", width: 24 },
        { header: "Entries", key: "entries", width: 10 },
        { header: "Total Hours", key: "hours", width: 14 },
        { header: "Labor cost ($)", key: "labor_cost", width: 14 },
      ];
      wsTech.addRows(
        [...byTech.values()]
          .map((r) => ({
            ...r,
            hours: Number(r.hours.toFixed(2)),
            labor_cost: Number(r.labor_cost.toFixed(2)),
          }))
          .sort((a, b) => b.hours - a.hours),
      );

      const wsPlant = wb.addWorksheet("Plant Summary");
      wsPlant.columns = [
        { header: "Plant no", key: "asset_code", width: 14 },
        { header: "Entries", key: "entries", width: 10 },
        { header: "Total Hours", key: "hours", width: 14 },
        { header: "Labor cost ($)", key: "labor_cost", width: 14 },
      ];
      wsPlant.addRows(
        [...byPlant.values()]
          .map((r) => ({
            ...r,
            hours: Number(r.hours.toFixed(2)),
            labor_cost: Number(r.labor_cost.toFixed(2)),
          }))
          .sort((a, b) => b.hours - a.hours),
      );

      const buffer = await wb.xlsx.writeBuffer();
      return reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Content-Disposition", `attachment; filename="IRONLOG_Mechanics_Cost_${year}.xlsx"`)
        .send(buffer);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

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

  // GET /api/maintenance/mechanic-labor/timesheet?from=&to=&year=
  app.get("/mechanic-labor/timesheet", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, MECHANIC_LABOR_EDITORS)) return;
      ensureMechanicLaborExtendedColumns();
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const { from, to } = parseMechanicsTimesheetRange(req);
      if (!isDate(from) || !isDate(to)) {
        return reply.code(400).send({ ok: false, error: "from and to (YYYY-MM-DD), or year=YYYY, required" });
      }
      if (from > to) {
        return reply.code(400).send({ ok: false, error: "from must be on or before to" });
      }

      const { rows, meta } = generateMechanicsTimesheet(db, {
        from,
        to,
        siteCode: site_code,
      });
      const exportRows = mechanicsTimesheetToExportRows(rows);
      return reply.send({ ok: true, meta, rows: exportRows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // GET /api/maintenance/mechanic-labor/timesheet.xlsx?year=2026
  app.get("/mechanic-labor/timesheet.xlsx", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, MECHANIC_LABOR_EDITORS)) return;
      ensureMechanicLaborExtendedColumns();
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const { from, to } = parseMechanicsTimesheetRange(req);
      if (!isDate(from) || !isDate(to)) {
        return reply.code(400).send({ ok: false, error: "from and to (YYYY-MM-DD), or year=YYYY, required" });
      }

      const { rows, meta } = generateMechanicsTimesheet(db, {
        from,
        to,
        siteCode: site_code,
      });
      const exportRows = mechanicsTimesheetToExportRows(rows);
      const { wb, year } = buildMechanicsTimesheetWorkbook(exportRows, meta);
      const buffer = await wb.xlsx.writeBuffer();
      return reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Content-Disposition", `attachment; filename="IRONLOG_Mechanics_Timesheet_${year}.xlsx"`)
        .send(buffer);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // POST /api/maintenance/mechanic-labor/timesheet/import — multipart file (xlsx/csv)
  // Query/fields: mode=append|replace_dates (default replace_dates)
  app.post("/mechanic-labor/timesheet/import", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, MECHANIC_LABOR_EDITORS)) return;
      ensureMechanicLaborExtendedColumns();
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const userName = String(req.headers?.["x-user-name"] || "").trim() || "system";

      const part = await req.file();
      if (!part) {
        return reply.code(400).send({ ok: false, error: "Upload file field named 'file' (.xlsx or .csv)" });
      }
      const buffer = await part.toBuffer();
      const filename = String(part.filename || "upload.xlsx");

      const modeRaw = String(
        req.query?.mode ||
        (typeof part.fields?.mode?.value === "string" ? part.fields.mode.value : "") ||
        "replace_dates",
      ).trim().toLowerCase();
      const mode = modeRaw === "append" ? "append" : "replace_dates";

      const parsed = await parseMechanicsTimesheetUpload(buffer, filename);
      if (parsed.errors.length && !parsed.entries.length) {
        return reply.code(400).send({
          ok: false,
          error: parsed.errors.slice(0, 10).join("; "),
          errors: parsed.errors.slice(0, 50),
        });
      }
      if (!parsed.entries.length) {
        return reply.code(400).send({ ok: false, error: "No valid timesheet rows found in file" });
      }

      const insertStmt = db.prepare(`
        INSERT INTO mechanic_labor_entries (
          work_date, technician_name, hours, asset_code, reason,
          labor_rate_per_hour, site_code, created_by, updated_by,
          category, time_started, time_finished, job_card_no, smr
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `);
      const deleteByDateStmt = db.prepare(`
        DELETE FROM mechanic_labor_entries
        WHERE work_date = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
      `);

      const dates = [...new Set(parsed.entries.map((e) => e.work_date))].sort();
      const tx = db.transaction(() => {
        let deleted = 0;
        if (mode === "replace_dates") {
          for (const d of dates) {
            deleted += Number(deleteByDateStmt.run(d, site_code).changes || 0);
          }
        }
        const ids = [];
        for (const e of parsed.entries) {
          const info = insertStmt.run(
            e.work_date,
            e.technician_name,
            e.hours,
            e.asset_code,
            e.reason,
            null,
            site_code,
            userName,
            userName,
            e.category,
            e.time_started,
            e.time_finished,
            e.job_card_no,
            e.smr,
          );
          ids.push(Number(info.lastInsertRowid || 0));
        }
        return { deleted, ids };
      });

      const result = tx();
      writeAudit(db, req, {
        module: "maintenance",
        action: "mechanic_labor.timesheet_import",
        entity_type: "mechanic_labor_entry",
        entity_id: dates[0] || null,
        payload: {
          mode,
          filename,
          imported: parsed.entries.length,
          deleted: result.deleted,
          dates: dates.length,
          row_errors: parsed.errors.length,
          skipped: parsed.skipped,
        },
      });

      return reply.send({
        ok: true,
        mode,
        filename,
        imported: parsed.entries.length,
        deleted: result.deleted,
        skipped: parsed.skipped,
        dates,
        date_count: dates.length,
        warnings: parsed.errors.slice(0, 20),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
}
