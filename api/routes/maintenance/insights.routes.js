// IRONLOG/api/routes/maintenance/insights.routes.js — Reliability, maintenance insights, governance signals and histogram events.
// Registered by routes/maintenance.routes.js; shared helpers arrive through ctx.
import ExcelJS from "exceljs";
import { buildPdfBuffer, sectionTitle, table } from "../../utils/pdfGenerator.js";
import { buildReliabilityExecutiveWorkbook } from "../../utils/reliabilityExecutiveWorkbook.js";
import { db } from "../../db/client.js";
import { isDate } from "../../utils/request.js";

export default function registerInsightsRoutes(app, ctx) {
  const {
    addInsightsCostPerMachineSheet,
    addInsightsExportSummarySheet,
    addInsightsPartsDemandSheets,
    buildInsightsBreakdownLaborIncidents,
    buildMaintenanceInsightsDowntime,
    buildMaintenanceReliabilityReport,
    buildUpcomingServiceCostForecasts,
    ensureBreakdownRepairLaborSchema,
    getAssetCurrentHoursInfo,
    hasColumn,
    hasTable,
    loadInsightsExportPayload,
    parseReliabilityAssetIds,
  } = ctx;

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
      const scheduled = Math.max(0.5, Number(req.query?.scheduled ?? 10) || 10);
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const data = buildMaintenanceReliabilityReport(startDate, endDate, { asset_ids, category, scheduled, site_code });
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
      const scheduled = Math.max(0.5, Number(req.query?.scheduled ?? 10) || 10);
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const data = buildMaintenanceReliabilityReport(startDate, endDate, { asset_ids, category, scheduled, site_code });

      const wb = new ExcelJS.Workbook();
      wb.creator = "IRONLOG";
      wb.created = new Date();

      const summary = wb.addWorksheet("Summary");
      summary.addRow(["MTBF / LTTR Report"]);
      summary.addRow(["Period", `${startDate} to ${endDate}`]);
      summary.addRow(["Category filter", category || "All"]);
      summary.addRow(["Assets in scope", data.asset_filter_count]);
      summary.addRow(["Asset KPI scheduled fallback", data.scheduled_fallback]);
      summary.addRow(["Downtime basis", data.downtime_basis === "asset_kpi_daily" ? "Asset KPI daily downtime" : "Recorded breakdown downtime"]);
      summary.addRow([]);
      summary.addRow(["Metric", "Value"]);
      summary.addRow(["Failures", data.summary.failure_count]);
      summary.addRow(["Operating hours", data.summary.operating_hours]);
      summary.addRow(["Downtime hours", data.summary.downtime_hours]);
      summary.addRow(["Recorded incident downtime (audit)", data.summary.recorded_downtime_hours]);
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
        { header: "Recorded incident h", key: "recorded_downtime_hours", width: 18 },
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

  app.get("/reliability-executive.xlsx", async (req, reply) => {
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
      const scheduled = Math.max(0.5, Number(req.query?.scheduled ?? 10) || 10);
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const data = buildMaintenanceReliabilityReport(startDate, endDate, { asset_ids, category, scheduled, site_code });
      const buf = await buildReliabilityExecutiveWorkbook(data, { startDate, endDate, category });
      reply
        .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .header("Content-Disposition", `attachment; filename="IRONLOG_MTBF_LTTR_Executive_${startDate}_to_${endDate}.xlsx"`)
        .send(Buffer.from(buf));
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

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
}
