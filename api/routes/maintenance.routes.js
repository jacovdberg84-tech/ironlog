// IRONLOG/api/routes/maintenance.routes.js
import { db } from "../db/client.js";
import multipart from "@fastify/multipart";
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import ExcelJS from "exceljs";
import { buildPdfBuffer, ensurePageSpace, pdfBodyBottom, pdfBodyTop, sectionTitle, table } from "../utils/pdfGenerator.js";
import { getPdfReportBranding } from "../utils/reportSettings.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";
import {
  resolveMachinePrestartProfile,
  getMachinePrestartTemplate,
  listMachinePrestartProfiles,
  normalizeMachinePrestartChecklist,
  checklistToJsonObject,
  machinePrestartCheckMode,
} from "../utils/machinePrestartTemplates.js";
import { normalizeUploadedPhoto } from "../utils/imagePdf.js";
import { resolveStorageAbs as resolveStorageAbsPath, getDataRoot } from "../utils/storagePaths.js";

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

function classifyDueStatus(remainingHours, nearDueHours = 50) {
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

      const oilAvg = hasTable("oil_logs") && hasTable("work_orders")
        ? dbConn.prepare(`
            SELECT
              COALESCE(SUM(ol.quantity), 0) AS oil_qty_total,
              COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, 0)), 0) AS oil_cost_total
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
          `).get(assetId, planId)
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

      const plans = rows.map((r) => {
        const current_hours = getAssetCurrentHours(r.asset_id);
        return {
          ...r,
          current_hours: Number(current_hours.toFixed(2)),
          next_due_hours: Number((Number(r.last_service_hours || 0) + Number(r.interval_hours || 0)).toFixed(2)),
          remaining_hours: Number((Number(r.last_service_hours || 0) + Number(r.interval_hours || 0) - current_hours).toFixed(2)),
        };
      });

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
      const last_service_hours = Number(req.body?.last_service_hours || 0);
      const active = Number(req.body?.active ?? 1) ? 1 : 0;

      if (!asset_id || !service_name || interval_hours <= 0) {
        return reply.code(400).send({
          ok: false,
          error: "asset_id, service_name and interval_hours are required"
        });
      }

      const asset = db.prepare(`
        SELECT id
        FROM assets
        WHERE id = ?
      `).get(asset_id);

      if (!asset) {
        return reply.code(404).send({
          ok: false,
          error: "Asset not found"
        });
      }

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

      const last_service_hours =
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
        SELECT id
        FROM assets
        WHERE id = ?
      `).get(asset_id);

      if (!asset) {
        return reply.code(404).send({
          ok: false,
          error: "Asset not found"
        });
      }

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
      const safeHours = Number.isFinite(liveHours) ? Number(liveHours.toFixed(2)) : 0;

      db.prepare(`
        UPDATE maintenance_plans
        SET last_service_hours = ?
        WHERE id = ?
      `).run(safeHours, Number(plan.id));

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

      return reply.send({
        ok: true,
        asset_id: asset.id,
        asset_code: asset.asset_code,
        asset_name: asset.asset_name,
        as_of: isDate(asOf) ? asOf : null,
        current_hours: Number(current_hours.toFixed(1)),
        current_hours_source: currentInfo.source
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

      const due = rows.map((r) => {
  const current = getAssetCurrentHours(r.asset_id);
  const next_due = Number(r.last_service_hours || 0) + Number(r.interval_hours || 0);
  const remaining = next_due - current;
  const status = classifyDueStatus(remaining, nearDueHours);

  return {
    plan_id: r.plan_id,
    asset_id: r.asset_id,
    asset_code: r.asset_code,
    asset_name: r.asset_name,
    category: r.category,
    service_name: r.service_name,
    interval_hours: Number(r.interval_hours || 0),
    last_service_hours: Number(r.last_service_hours || 0),
    current_hours: Number(current.toFixed(2)),
    next_due_hours: Number(next_due.toFixed(2)),
    remaining_hours: Number(remaining.toFixed(2)),
    is_overdue: remaining <= 0,
    status
  };
});

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

      const rows = plans.map((p) => {
        const currentInfo = getAssetCurrentHoursInfo(Number(p.asset_id || 0));
        const current = Number(currentInfo.hours || 0);
        const next_due = Number(p.last_service_hours || 0) + Number(p.interval_hours || 0);
        const remaining = next_due - current;

        const lastWo = getLastServiced.get(Number(p.plan_id || 0));
        const lastBackfill = getLastBackfillServiced.get(Number(p.asset_id || 0), String(p.service_name || ""));
        const lastWoAt = String(lastWo?.last_serviced_at || "");
        const lastBackfillAt = String(lastBackfill?.last_serviced_at || "");
        const useBackfill = Boolean(lastBackfillAt && (!lastWoAt || lastBackfillAt > lastWoAt));
        const last = useBackfill ? lastBackfill : lastWo;
        const avgRow = getAvgDaily.get(Number(p.asset_id || 0), startDate, endDate);
        const totalRun = Number(avgRow?.total_run || 0);
        const dayCount = Number(avgRow?.day_count || 0);
        const avgDaily = dayCount > 0 ? totalRun / dayCount : 0;

        const estDays = avgDaily > 0 ? Math.max(0, remaining / avgDaily) : null;
        const estDate = estDays == null ? null : addDays(endDate, estDays);

        return {
          plan_id: Number(p.plan_id || 0),
          asset_id: Number(p.asset_id || 0),
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
        };
      });

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
    const canBreakdowns = hasTable("breakdowns");
    const canDtLogs = hasTable("breakdown_downtime_logs");
    const dtCol = breakdownDowntimeColumnName();
    const breakdownDateExpr = hasColumn("breakdowns", "breakdown_date")
      ? "b.breakdown_date"
      : "DATE(COALESCE(b.created_at, b.updated_at))";

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

    const failuresByAsset = new Map();
    if (canBreakdowns) {
      for (const r of db.prepare(`
        SELECT b.asset_id, COUNT(*) AS failure_count
        FROM breakdowns b
        WHERE ${breakdownDateExpr} BETWEEN ? AND ?
          AND b.asset_id IN (${marks})
        GROUP BY b.asset_id
      `).all(start, end, ...assetIds)) {
        failuresByAsset.set(Number(r.asset_id || 0), Number(r.failure_count || 0));
      }
    }

    const downtimeByAsset = new Map();
    if (canDtLogs) {
      for (const r of db.prepare(`
        SELECT b.asset_id, COALESCE(SUM(l.hours_down), 0) AS downtime_hours
        FROM breakdown_downtime_logs l
        JOIN breakdowns b ON b.id = l.breakdown_id
        WHERE l.log_date BETWEEN ? AND ?
          AND b.asset_id IN (${marks})
        GROUP BY b.asset_id
      `).all(start, end, ...assetIds)) {
        downtimeByAsset.set(Number(r.asset_id || 0), Number(r.downtime_hours || 0));
      }
    }
    if (canBreakdowns && dtCol) {
      for (const r of db.prepare(`
        SELECT b.asset_id, COALESCE(SUM(b.${dtCol}), 0) AS downtime_hours
        FROM breakdowns b
        WHERE ${breakdownDateExpr} BETWEEN ? AND ?
          AND b.asset_id IN (${marks})
        GROUP BY b.asset_id
      `).all(start, end, ...assetIds)) {
        const aid = Number(r.asset_id || 0);
        const base = Number(r.downtime_hours || 0);
        if (!downtimeByAsset.has(aid) || Number(downtimeByAsset.get(aid) || 0) <= 0) {
          downtimeByAsset.set(aid, base);
        }
      }
    }

    const by_asset = assetRows.map((a) => {
      const aid = Number(a.asset_id || 0);
      const operating_hours = Number(runByAsset.get(aid) || 0);
      const failure_count = Number(failuresByAsset.get(aid) || 0);
      const downtime_hours = Number(downtimeByAsset.get(aid) || 0);
      const mtbf_hours = failure_count > 0 ? operating_hours / failure_count : null;
      const lttr_hours = failure_count > 0 ? downtime_hours / failure_count : null;
      return {
        asset_id: aid,
        asset_code: String(a.asset_code || ""),
        asset_name: String(a.asset_name || ""),
        category: String(a.category || ""),
        failure_count,
        operating_hours: Number(operating_hours.toFixed(2)),
        downtime_hours: Number(downtime_hours.toFixed(2)),
        mtbf_hours: mtbf_hours == null ? null : Number(mtbf_hours.toFixed(2)),
        lttr_hours: lttr_hours == null ? null : Number(lttr_hours.toFixed(2)),
      };
    }).sort((x, y) => {
      const xm = x.mtbf_hours == null ? Infinity : Number(x.mtbf_hours);
      const ym = y.mtbf_hours == null ? Infinity : Number(y.mtbf_hours);
      return xm - ym;
    });

    const failure_count = by_asset.reduce((s, r) => s + Number(r.failure_count || 0), 0);
    const operating_hours = by_asset.reduce((s, r) => s + Number(r.operating_hours || 0), 0);
    const downtime_hours = by_asset.reduce((s, r) => s + Number(r.downtime_hours || 0), 0);
    const mtbf_hours = failure_count > 0 ? operating_hours / failure_count : null;
    const lttr_hours = failure_count > 0 ? downtime_hours / failure_count : null;

    return {
      start,
      end,
      category: categoryFilter || null,
      asset_filter_count: assetIds.length,
      formulas: {
        mtbf: "operating_hours / failure_count",
        lttr: "downtime_hours / failure_count",
        failures: "breakdowns with breakdown_date in selected period",
        operating_hours: "sum of daily_hours.hours_run (is_used=1) in period",
        downtime_hours: "sum of breakdown_downtime_logs in period, else breakdown header downtime",
      },
      summary: {
        failure_count,
        operating_hours: Number(operating_hours.toFixed(2)),
        downtime_hours: Number(downtime_hours.toFixed(2)),
        mtbf_hours: mtbf_hours == null ? null : Number(mtbf_hours.toFixed(2)),
        lttr_hours: lttr_hours == null ? null : Number(lttr_hours.toFixed(2)),
      },
      by_asset,
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
            GROUP BY service_key, part_name
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

  // GET /api/maintenance/insights.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD&near_due_hours=50
  app.get("/insights.xlsx", async (req, reply) => {
    try {
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
        return reply.code(injected.statusCode).send(payload?.error ? payload : { ok: false, error: "Failed to build insights export" });
      }

      const data = JSON.parse(String(injected.payload || "{}"));
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

  // GET /api/maintenance/insights.pdf?start=YYYY-MM-DD&end=YYYY-MM-DD&near_due_hours=50&download=1
  app.get("/insights.pdf", async (req, reply) => {
    try {
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
      const download = String(req.query?.download || "").trim() === "1";

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
        return reply.code(injected.statusCode).send(payload?.error ? payload : { ok: false, error: "Failed to build insights PDF" });
      }
      const data = JSON.parse(String(injected.payload || "{}"));
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
      if (!planId) {
        const matchedPlan = db.prepare(`
          SELECT id
          FROM maintenance_plans
          WHERE asset_id = ?
            AND UPPER(TRIM(service_name)) = UPPER(TRIM(?))
          ORDER BY active DESC, id DESC
          LIMIT 1
        `).get(assetId, serviceName);
        if (matchedPlan?.id) planId = Number(matchedPlan.id);
      }

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
          serviceHours,
          notes,
          String(req.headers?.["x-user-name"] || "system")
        );
        if (updatePlanLastHours && planId && serviceHours != null) {
          updatePlan.run(Number(serviceHours), Number(planId));
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
          service_hours: serviceHours,
        },
      });
      return reply.send({
        ok: true,
        id,
        asset_id: assetId,
        asset_code: asset.asset_code,
        service_name: serviceName,
        service_date: serviceDate,
        service_hours: serviceHours,
        plan_id: planId || null,
        plan_last_hours_updated: Boolean(updatePlanLastHours && planId && serviceHours != null),
      });
    } catch (e) {
      req.log.error(e);
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  // GET /api/maintenance/history/backfill?asset_id=&limit=20
  app.get("/history/backfill", async (req, reply) => {
    try {
      const assetId = Number(req.query?.asset_id || 0);
      const limit = Math.max(1, Math.min(200, Number(req.query?.limit || 20)));
      const where = assetId > 0 ? "WHERE h.asset_id = ?" : "";
      const rows = db.prepare(`
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
          a.asset_name
        FROM maintenance_service_history h
        JOIN assets a ON a.id = h.asset_id
        ${where}
        ORDER BY h.service_date DESC, h.id DESC
        LIMIT ?
      `).all(...(assetId > 0 ? [assetId, limit] : [limit]));
      return reply.send({ ok: true, rows });
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

      const tx = db.transaction(() => {
        const created = [];

        for (const p of plans) {
          const current = getAssetCurrentHours(p.asset_id);
          const next_due = Number(p.last_service_hours || 0) + Number(p.interval_hours || 0);
          const remaining = next_due - current;
          const isOverdue = remaining <= 0;
          const isNearDue = remaining <= nearDueHours;
          const isRequestedPlan = requestedPlanIds.includes(Number(p.plan_id || 0));
          const shouldCreate = requestedPlanIds.length ? (isRequestedPlan && isNearDue) : isOverdue;
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

      const rows = date
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
              AND a.is_standby = 0
              AND a.archived = 0
            ORDER BY a.asset_code ASC, mp.service_name ASC
          `).all()
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

      const dueRows = rows
        .map((r) => {
          const current = getAssetCurrentHours(r.asset_id);
          const nextDue = Number(r.last_service_hours || 0) + Number(r.interval_hours || 0);
          const remaining = nextDue - current;
          const status = classifyDueStatus(remaining, nearDueHours);
          return {
            asset_code: r.asset_code,
            asset_name: r.asset_name,
            service_name: r.service_name,
            current_hours: Number(current.toFixed(2)),
            next_due_hours: Number(nextDue.toFixed(2)),
            remaining_hours: Number(remaining.toFixed(2)),
            status,
          };
        })
        .filter((r) => r.status === "OVERDUE" || r.status === "ALMOST DUE")
        .sort((a, b) => Number(a.remaining_hours || 0) - Number(b.remaining_hours || 0));

      const asOfLabel = date || new Date().toISOString().slice(0, 10);
      const pdf = await buildPdfBuffer(
        (doc) => {
          sectionTitle(doc, "Upcoming Services");
          doc
            .font("Helvetica")
            .fontSize(10)
            .text(`As of: ${asOfLabel} | Threshold: <= ${nearDueHours.toFixed(0)}h flagged as ALMOST DUE`);
          doc.moveDown(0.4);

          table(
            doc,
            [
              { key: "asset_code", label: "Asset", width: 0.12 },
              { key: "asset_name", label: "Name", width: 0.2 },
              { key: "service_name", label: "Service", width: 0.18 },
              { key: "current_hours", label: "Current", width: 0.12, align: "right" },
              { key: "next_due_hours", label: "Next Due", width: 0.12, align: "right" },
              { key: "remaining_hours", label: "Remaining", width: 0.12, align: "right" },
              { key: "status", label: "Status", width: 0.14 },
            ],
            dueRows.length
              ? dueRows.map((r) => ({
                  ...r,
                  current_hours: Number(r.current_hours || 0).toFixed(1),
                  next_due_hours: Number(r.next_due_hours || 0).toFixed(1),
                  remaining_hours: Number(r.remaining_hours || 0).toFixed(1),
                }))
              : [
                  {
                    asset_code: "-",
                    asset_name: "No upcoming services within threshold",
                    service_name: "-",
                    current_hours: "-",
                    next_due_hours: "-",
                    remaining_hours: "-",
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
              SELECT COALESCE(SUM(COALESCE(quantity, 0) * COALESCE(unit_cost, 0)), 0) AS v
              FROM oil_logs
              WHERE log_date BETWEEN ? AND ?
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

      const oilCost = oilCostFromLogs + oilCostFromWoStock;

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
              COALESCE(SUM(COALESCE(ol.quantity, 0) * COALESCE(ol.unit_cost, 0)), 0) AS v
            FROM oil_logs ol
            JOIN assets a ON a.id = ol.asset_id
            WHERE ol.log_date BETWEEN ? AND ?
            GROUP BY ol.asset_id
          `).all(start, end);
          for (const r of oilLogRows || []) putActual(r.asset_id, r.asset_code, r.asset_name, { lubes_logs_cost: Number(r.v || 0) });
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
          stores_oil_from_logs: Number(oilCostFromLogs.toFixed(2)),
          stores_oil_from_work_orders: Number(oilCostFromWoStock.toFixed(2)),
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
        if (work_order_id > 0 && hasColumn("work_orders", "completion_notes") && woNotes) {
          db.prepare(`UPDATE work_orders SET completion_notes = ? WHERE id = ?`).run(woNotes, work_order_id);
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
  // TYRE INSPECTIONS
  // =====================================================
  function normalizeTyreRows(raw) {
    if (!Array.isArray(raw)) return [];
    return raw
      .map((x) => {
        const position_key = String(x?.position_key || "").trim().toLowerCase();
        const position_label = String(x?.position_label || "").trim();
        const pressureRaw = x?.pressure;
        const treadRaw = x?.tread_depth;
        const costRaw = x?.tyre_cost;
        const pressure = pressureRaw === "" || pressureRaw == null ? null : Number(pressureRaw);
        const tread_depth = treadRaw === "" || treadRaw == null ? null : Number(treadRaw);
        const tyre_cost = costRaw === "" || costRaw == null ? 0 : Number(costRaw);
        return {
          position_key,
          position_label: position_label || position_key,
          pressure: Number.isFinite(pressure) ? pressure : null,
          tread_depth: Number.isFinite(tread_depth) ? tread_depth : null,
          serial_number: String(x?.serial_number || "").trim() || null,
          last_changed_date: isDate(String(x?.last_changed_date || "").trim()) ? String(x.last_changed_date).trim() : null,
          tyre_cost: Number.isFinite(tyre_cost) && tyre_cost > 0 ? tyre_cost : 0,
        };
      })
      .filter((x) => x.position_key);
  }

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
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const asset_id = Number(req.body?.asset_id || 0);
      const inspection_date = String(req.body?.inspection_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim() || null;
      const notes = String(req.body?.notes || "").trim() || null;
      if (!asset_id) return reply.code(400).send({ ok: false, error: "asset_id is required" });
      if (!isDate(inspection_date)) return reply.code(400).send({ ok: false, error: "inspection_date must be YYYY-MM-DD" });
      const asset = db.prepare(`SELECT id FROM assets WHERE id = ?`).get(asset_id);
      if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

      const tyres = normalizeTyreRows(req.body?.tyres);
      const runningRaw = req.body?.running_hours;
      const running_hours = runningRaw == null || runningRaw === ""
        ? 0
        : Number(runningRaw);
      if (!Number.isFinite(running_hours) || running_hours < 0) {
        return reply.code(400).send({ ok: false, error: "running_hours must be a positive number" });
      }
      const total_tyre_cost = Number(
        tyres.reduce((sum, t) => sum + Number(t.tyre_cost || 0), 0).toFixed(2)
      );
      const cost_per_running_hour = Number(
        (total_tyre_cost / (running_hours > 0 ? running_hours : 1)).toFixed(4)
      );

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
        notes
      );

      return reply.send({
        ok: true,
        id: Number(ins.lastInsertRowid),
        total_tyre_cost,
        cost_per_running_hour,
      });
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

  function syncLdvPrestartToDailyHours(assetId, checkDate, odometerKm, inspectorName, previousOdometerKm) {
    if (!assetId || !isDate(checkDate) || !Number.isFinite(Number(odometerKm))) return { synced: false };
    const odometer = Number(odometerKm);
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

      const ldvCheckStmt = db.prepare(`
        SELECT id, checklist_json
        FROM vehicle_ldv_checks
        WHERE asset_id = ?
          AND check_date = ?
          AND COALESCE(check_mode, 'ldv_general') = 'prestart'
        ORDER BY id DESC
        LIMIT 1
      `);

      for (const a of assets) {
        const assetId = Number(a.id);
        const code = String(a.asset_code || "");
        if (isLdvPrestartAssetCode(code)) {
          const row = ldvCheckStmt.get(assetId, check_date);
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
        const row = db.prepare(`
          SELECT id, checklist_json
          FROM vehicle_ldv_checks
          WHERE asset_id = ?
            AND check_date = ?
            AND check_mode = ?
          ORDER BY id DESC
          LIMIT 1
        `).get(assetId, check_date, mode);
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

      return reply.send({
        ok: true,
        check_date,
        summary,
        ldv: ldvAssets,
        machine_groups,
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

      const existing = db.prepare(`
        SELECT id, check_date, odometer_km, inspector_name, notes, checklist_json
        FROM vehicle_ldv_checks
        WHERE asset_id = ?
          AND check_date = ?
          AND COALESCE(check_mode, 'ldv_general') = 'prestart'
        ORDER BY id DESC
        LIMIT 1
      `).get(Number(asset.id), check_date);

      const previous_odometer_km = getLatestLdvOdometerKm(Number(asset.id), check_date, {
        excludeCheckId: existing?.id ? Number(existing.id) : 0,
      });

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
        previous_odometer_source:
          previous_odometer_km != null ? "daily_hours_or_prior_prestart" : null,
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

      const existing = db.prepare(`
        SELECT id
        FROM vehicle_ldv_checks
        WHERE asset_id = ?
          AND check_date = ?
          AND COALESCE(check_mode, 'ldv_general') = 'prestart'
        ORDER BY id DESC
        LIMIT 1
      `).get(Number(asset.id), check_date);

      const previousOdometer = getLatestLdvOdometerKm(Number(asset.id), check_date, {
        excludeCheckId: existing?.id ? Number(existing.id) : 0,
      });
      if (previousOdometer != null && odometer_km < previousOdometer) {
        return reply.code(400).send({
          ok: false,
          error: `Odometer cannot move backwards (previous ${previousOdometer.toFixed(1)} km).`,
          previous_odometer_km: Number(previousOdometer.toFixed(1)),
        });
      }

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

      let checkId = 0;
      if (existing?.id) {
        checkId = Number(existing.id);
        db.prepare(`
          UPDATE vehicle_ldv_checks
          SET
            odometer_km = ?,
            inspector_name = ?,
            notes = ?,
            check_mode = 'prestart',
            checklist_json = ?,
            updated_at = datetime('now')
          WHERE id = ?
        `).run(odometer_km, inspector_name, notes, checklistJson, checkId);
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
          notes,
          checklistJson
        );
        checkId = Number(ins.lastInsertRowid);
      }

      const dailySync = syncLdvPrestartToDailyHours(
        Number(asset.id),
        check_date,
        odometer_km,
        inspector_name,
        previousOdometer
      );

      return reply.send({
        ok: true,
        id: checkId,
        asset_code: String(asset.asset_code || ""),
        check_date,
        odometer_km: Number(odometer_km.toFixed(1)),
        previous_odometer_km: previousOdometer == null ? null : Number(previousOdometer.toFixed(1)),
        daily_input_sync: dailySync || { synced: false },
        message: "Pre-start captured. KM reading saved to IRONLOG.",
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
      const existing = db.prepare(`
        SELECT id, checklist_json FROM vehicle_ldv_checks
        WHERE asset_id = ? AND check_date = ? AND COALESCE(check_mode, 'ldv_general') = 'prestart'
        ORDER BY id DESC LIMIT 1
      `).get(assetId, work_date);

      let checkId = Number(existing?.id || 0);
      const previousOdometer = getLatestLdvOdometerKm(assetId, work_date, {
        excludeCheckId: checkId,
      });
      const opening_km =
        opening_km_raw != null && String(opening_km_raw).trim() !== ""
          ? Number(opening_km_raw)
          : previousOdometer;
      if (!Number.isFinite(opening_km) || opening_km < 0) {
        return reply.code(400).send({ ok: false, error: "opening_km could not be resolved — pass opening_km explicitly" });
      }
      if (closing_km < opening_km) {
        return reply.code(400).send({
          ok: false,
          error: `Closing KM (${closing_km}) cannot be less than opening KM (${opening_km}).`,
          previous_odometer_km: previousOdometer,
          opening_km,
        });
      }

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

      if (checkId > 0) {
        db.prepare(`
          UPDATE vehicle_ldv_checks
          SET odometer_km = ?, inspector_name = ?, notes = ?, check_mode = 'prestart',
              checklist_json = COALESCE(NULLIF(checklist_json, ''), ?), updated_at = datetime('now')
          WHERE id = ?
        `).run(closing_km, inspector_name, correction_note, checklistJsonOut, checkId);
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
        checkId = Number(ins.lastInsertRowid);
      }

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

      return reply.send({
        ok: true,
        asset_code,
        work_date,
        check_id: checkId,
        opening_km: Number(opening_km.toFixed(1)),
        closing_km: Number(closing_km.toFixed(1)),
        run_km: Number(run_km.toFixed(1)),
        previous_odometer_km: previousOdometer == null ? null : Number(previousOdometer.toFixed(1)),
        message: `KM corrected for ${asset_code} on ${work_date}.`,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // Machine-group prestart (excavator, dozer, etc.) — same storage as LDV checks, different check_mode; no daily_hours sync.
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

      const existing = db.prepare(`
        SELECT id, check_date, odometer_km, smu_hours, inspector_name, notes, checklist_json
        FROM vehicle_ldv_checks
        WHERE asset_id = ?
          AND check_date = ?
          AND check_mode = ?
        ORDER BY id DESC
        LIMIT 1
      `).get(Number(asset.id), check_date, mode);

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

      const existing = db.prepare(`
        SELECT id
        FROM vehicle_ldv_checks
        WHERE asset_id = ?
          AND check_date = ?
          AND check_mode = ?
        ORDER BY id DESC
        LIMIT 1
      `).get(Number(asset.id), check_date, mode);

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

      return reply.send({
        ok: true,
        id: checkId,
        asset_code: String(asset.asset_code || ""),
        check_date,
        profile_id: profileId,
        check_mode: mode,
        smu_hours,
        message: "Machine pre-start saved to IRONLOG.",
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
}