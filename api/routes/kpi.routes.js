// IRONLOG/api/routes/kpi.routes.js
import { db } from "../db/client.js";
import { andDailyHoursFleetHoursOnly } from "../utils/fleetHoursKpiScope.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

export default async function kpiRoutes(app) {
  // Daily KPI for a date
  // GET /api/kpi/daily/2026-02-27?scheduled=10
  app.get("/daily/:date", async (req, reply) => {
    const date = String(req.params.date || "").trim();
    const scheduled = Number(req.query?.scheduled ?? 10); // default 10 hours per asset/day

    if (!isDate(date)) return reply.code(400).send({ error: "date must be YYYY-MM-DD" });
    if (!Number.isFinite(scheduled) || scheduled <= 0 || scheduled > 24) {
      return reply.code(400).send({ error: "scheduled must be 1..24" });
    }

    // Used assets ONLY (exclude standby + exclude is_used=0 + exclude hours_run=0)
    const usedAssets = db.prepare(`
      SELECT DISTINCT dh.asset_id
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date = ?
        AND dh.hours_run > 0
        ${andDailyHoursFleetHoursOnly("dh", "a")}
    `).all(date);

    const used_count = usedAssets.length;
    const available_hours = used_count * scheduled;

    const runRow = db.prepare(`
      SELECT IFNULL(SUM(dh.hours_run), 0) AS run_hours
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date = ?
        AND dh.hours_run > 0
        ${andDailyHoursFleetHoursOnly("dh", "a")}
    `).get(date);

    const run_hours = Number(runRow.run_hours || 0);

    // Downtime for the same day, only for used assets
    let downtime_hours = 0;
    if (used_count > 0) {
      const downtimeRow = db.prepare(`
        SELECT IFNULL(SUM(b.downtime_hours), 0) AS dt
        FROM breakdowns b
        WHERE b.breakdown_date = ?
          AND b.asset_id IN (
            SELECT DISTINCT dh.asset_id
            FROM daily_hours dh
            JOIN assets a ON a.id = dh.asset_id
            WHERE dh.work_date = ?
              AND dh.hours_run > 0
              ${andDailyHoursFleetHoursOnly("dh", "a")}
          )
      `).get(date, date);
      downtime_hours = Number(downtimeRow.dt || 0);
    }

    const availability = available_hours > 0
      ? ((available_hours - downtime_hours) / available_hours) * 100
      : null;

    const utilization = available_hours > 0
      ? (run_hours / available_hours) * 100
      : null;

    const major_downtime = db.prepare(`
      SELECT
        b.id,
        b.breakdown_date,
        b.description,
        b.downtime_hours,
        b.critical,
        a.asset_code,
        a.asset_name
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      WHERE b.breakdown_date = ?
      ORDER BY b.downtime_hours DESC, b.id DESC
      LIMIT 10
    `).all(date).map(r => ({ ...r, critical: Boolean(r.critical) }));

    return {
      ok: true,
      period: { type: "daily", date },
      scheduled_hours_per_asset: scheduled,
      used_assets_count: used_count,
      available_hours,
      run_hours,
      downtime_hours,
      availability: availability == null ? null : Number(availability.toFixed(2)),
      utilization: utilization == null ? null : Number(utilization.toFixed(2)),
      major_downtime
    };
  });

  // Range KPI (daily + weekly reports can use this)
  // GET /api/kpi/range?start=2026-02-24&end=2026-03-02&scheduled=10
  app.get("/range", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    const scheduled = Number(req.query?.scheduled ?? 10);

    if (!isDate(start) || !isDate(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }
    if (!Number.isFinite(scheduled) || scheduled <= 0 || scheduled > 24) {
      return reply.code(400).send({ error: "scheduled must be 1..24" });
    }

    // Used assets per day are counted day-by-day (correct for standby)
    const usedPerDay = db.prepare(`
      SELECT
        dh.work_date,
        COUNT(DISTINCT dh.asset_id) AS used_assets,
        IFNULL(SUM(dh.hours_run), 0) AS run_hours
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      WHERE dh.work_date BETWEEN ? AND ?
        AND dh.hours_run > 0
        ${andDailyHoursFleetHoursOnly("dh", "a")}
      GROUP BY dh.work_date
      ORDER BY dh.work_date
    `).all(start, end);

    const available_hours = usedPerDay.reduce((acc, d) => acc + (Number(d.used_assets) * scheduled), 0);
    const run_hours = usedPerDay.reduce((acc, d) => acc + Number(d.run_hours || 0), 0);

    const downtimeRow = db.prepare(`
      SELECT IFNULL(SUM(b.downtime_hours), 0) AS dt
      FROM breakdowns b
      WHERE b.breakdown_date BETWEEN ? AND ?
    `).get(start, end);

    const downtime_hours = Number(downtimeRow.dt || 0);

    const availability = available_hours > 0
      ? ((available_hours - downtime_hours) / available_hours) * 100
      : null;

    const utilization = available_hours > 0
      ? (run_hours / available_hours) * 100
      : null;

    const major_downtime = db.prepare(`
      SELECT
        b.id,
        b.breakdown_date,
        b.description,
        b.downtime_hours,
        b.critical,
        a.asset_code,
        a.asset_name
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      WHERE b.breakdown_date BETWEEN ? AND ?
      ORDER BY b.downtime_hours DESC, b.id DESC
      LIMIT 20
    `).all(start, end).map(r => ({ ...r, critical: Boolean(r.critical) }));

    return {
      ok: true,
      period: { type: "range", start, end },
      scheduled_hours_per_asset: scheduled,
      totals: {
        available_hours,
        run_hours,
        downtime_hours,
        availability: availability == null ? null : Number(availability.toFixed(2)),
        utilization: utilization == null ? null : Number(utilization.toFixed(2))
      },
      daily: usedPerDay.map(d => ({
        date: d.work_date,
        used_assets: Number(d.used_assets),
        available_hours: Number(d.used_assets) * scheduled,
        run_hours: Number(d.run_hours || 0)
      })),
      major_downtime
    };
  });
}