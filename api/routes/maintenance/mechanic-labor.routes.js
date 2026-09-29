// IRONLOG/api/routes/maintenance/mechanic-labor.routes.js — Mechanic labour entries, settings and timesheets.
// Registered by routes/maintenance.routes.js; shared helpers arrive through ctx.
import ExcelJS from "exceljs";
import { createManagementSummary, styleManagementDetailSheet } from "../../utils/managementWorkbook.js";
import { db } from "../../db/client.js";
import { generateMechanicsTimesheet, mechanicsTimesheetToExportRows } from "../../utils/mechanicsTimesheetGenerator.js";
import { isDate } from "../../utils/request.js";
import { parseMechanicsTimesheetUpload } from "../../utils/mechanicsTimesheetImport.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerMechanicLaborRoutes(app, ctx) {
  const {
    MECHANIC_LABOR_EDITORS,
    MECHANIC_LABOR_RATE_MANAGERS,
    MECHANIC_MONTH_NAMES,
    buildMechanicsTimesheetWorkbook,
    enrichMechanicLaborRow,
    enrichMechanicLaborSmr,
    ensureMechanicLaborExtendedColumns,
    mechanicLaborEntryBody,
    parseMechanicsTimesheetRange,
    readMechanicLaborDefaultRate,
    requireMaintenanceRoles,
    resolveMechanicLaborSmr,
  } = ctx;

  // GET /api/maintenance/mechanic-labor?date=YYYY-MM-DD
  app.get("/mechanic-labor", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, MECHANIC_LABOR_EDITORS)) return;
      ensureMechanicLaborExtendedColumns();
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const date = String(req.query?.date || "").trim();
      if (!isDate(date)) {
        return reply.code(400).send({ ok: false, error: "date (YYYY-MM-DD) required" });
      }
      const defaultRate = readMechanicLaborDefaultRate();
      const rows = db.prepare(`
        SELECT
          id, work_date, technician_name, hours, asset_code, reason, labor_rate_per_hour,
          category, time_started, time_finished, job_card_no, smr, created_by, updated_at
        FROM mechanic_labor_entries
        WHERE work_date = ? AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
        ORDER BY id ASC
      `).all(date, site_code)
        .map((r) => enrichMechanicLaborSmr(r, site_code))
        .map((r) => enrichMechanicLaborRow(r, defaultRate));

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
      ensureMechanicLaborExtendedColumns();
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const userName = String(req.headers?.["x-user-name"] || "").trim() || "system";
      const parsed = mechanicLaborEntryBody(req.body || {});
      if (parsed.smr == null) parsed.smr = resolveMechanicLaborSmr(parsed.work_date, parsed.asset_code, site_code);
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
          labor_rate_per_hour, site_code, created_by, updated_by,
          category, time_started, time_finished, job_card_no, smr
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
        parsed.category,
        parsed.time_started,
        parsed.time_finished,
        parsed.job_card_no,
        parsed.smr,
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
      ensureMechanicLaborExtendedColumns();
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
        if (parsed.smr == null) parsed.smr = resolveMechanicLaborSmr(work_date, parsed.asset_code, site_code);
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
          labor_rate_per_hour, site_code, created_by, updated_by,
          category, time_started, time_finished, job_card_no, smr
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            parsed.category,
            parsed.time_started,
            parsed.time_finished,
            parsed.job_card_no,
            parsed.smr,
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
      ensureMechanicLaborExtendedColumns();
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
      if (parsed.smr == null) parsed.smr = resolveMechanicLaborSmr(parsed.work_date, parsed.asset_code, site_code);
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
          category = ?,
          time_started = ?,
          time_finished = ?,
          job_card_no = ?,
          smr = ?,
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
        parsed.category,
        parsed.time_started,
        parsed.time_finished,
        parsed.job_card_no,
        parsed.smr,
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
      ensureMechanicLaborExtendedColumns();
      const year = String(req.query?.year || new Date().getFullYear()).trim();
      if (!/^\d{4}$/.test(year)) {
        return reply.code(400).send({ ok: false, error: "year (YYYY) required" });
      }
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const defaultRate = readMechanicLaborDefaultRate();
      const start = `${year}-01-01`;
      const end = `${year}-12-31`;

      const rawRows = db.prepare(`
        SELECT
          id, work_date, technician_name, hours, asset_code, reason, labor_rate_per_hour,
          category, time_started, time_finished, job_card_no, smr
        FROM mechanic_labor_entries
        WHERE work_date >= ? AND work_date <= ?
          AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
        ORDER BY work_date ASC, id ASC
      `).all(start, end, site_code);

      const rows = rawRows
        .map((r) => enrichMechanicLaborSmr(r, site_code))
        .map((r) => enrichMechanicLaborRow(r, defaultRate));

      const wb = new ExcelJS.Workbook();
      wb.creator = "IRONLOG";
      wb.created = new Date();

      const yearHours = rows.reduce((s, r) => s + Number(r.hours || 0), 0);
      const yearCost = rows.reduce((s, r) => s + Number(r.labor_cost || 0), 0);
      const technicianCount = new Set(rows.map((row) => String(row.technician_name || "").trim()).filter(Boolean)).size;
      createManagementSummary(wb, {
        title: "IRONLOG Mechanics Cost Report",
        periodLabel: `Reporting year: ${year}`,
        cards: [
          { label: "LABOR HOURS", value: Number(yearHours.toFixed(2)), numFmt: "#,##0.00" },
          { label: "LABOR COST", value: Number(yearCost.toFixed(2)), numFmt: "$#,##0.00" },
          { label: "WORK ENTRIES", value: rows.length, numFmt: "#,##0" },
          { label: "TECHNICIANS", value: technicianCount, numFmt: "#,##0" },
        ],
        scopeLines: [
          `Site: ${site_code}. Default labor rate: $${Number(defaultRate || 0).toFixed(2)} per hour.`,
          "Monthly sheets retain the full job-card, time, technician, and SMR history.",
        ],
      });

      const monthCols = [
        { header: "Date", key: "work_date", width: 14 },
        { header: "Plant no", key: "asset_code", width: 14 },
        { header: "Work Hours", key: "hours", width: 12 },
        { header: "Category", key: "category", width: 18 },
        { header: "Description Of Work Carried Out", key: "reason", width: 42 },
        { header: "Time Started", key: "time_started", width: 14 },
        { header: "Time finished", key: "time_finished", width: 14 },
        { header: "Technician", key: "technician_name", width: 22 },
        { header: "Job Card No", key: "job_card_no", width: 16 },
        { header: "SMR", key: "smr", width: 12 },
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
            asset_code: "",
            hours: 0,
            category: "",
            reason: "",
            time_started: "",
            time_finished: "",
            technician_name: "No entries",
            job_card_no: "",
            smr: "",
            labor_rate_per_hour: defaultRate,
            labor_cost: 0,
          });
        }
        const monthHours = monthRows.reduce((s, r) => s + Number(r.hours || 0), 0);
        const monthCost = monthRows.reduce((s, r) => s + Number(r.labor_cost || 0), 0);
        ws.addRow({});
        ws.addRow({
          work_date: "TOTAL",
          asset_code: "",
          hours: Number(monthHours.toFixed(2)),
          category: "",
          reason: "",
          time_started: "",
          time_finished: "",
          technician_name: "",
          job_card_no: "",
          smr: "",
          labor_rate_per_hour: "",
          labor_cost: Number(monthCost.toFixed(2)),
        });
        styleManagementDetailSheet(ws, {
          title: `Mechanics cost – ${sheetName} ${year}`,
          subtitle: `Site: ${site_code} · Scheduled work window: 06:00 to 17:00`,
          frozenColumns: 2,
          numberFormats: {
            hours: "#,##0.00",
            smr: "#,##0.0",
            labor_rate_per_hour: "$#,##0.00",
            labor_cost: "$#,##0.00",
          },
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
      styleManagementDetailSheet(wsTech, {
        title: "Mechanics cost by technician",
        subtitle: `Reporting year: ${year} · Site: ${site_code}`,
        frozenColumns: 1,
        numberFormats: { entries: "#,##0", hours: "#,##0.00", labor_cost: "$#,##0.00" },
      });

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
      styleManagementDetailSheet(wsPlant, {
        title: "Mechanics cost by equipment",
        subtitle: `Reporting year: ${year} · Site: ${site_code}`,
        frozenColumns: 1,
        numberFormats: { entries: "#,##0", hours: "#,##0.00", labor_cost: "$#,##0.00" },
      });

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
