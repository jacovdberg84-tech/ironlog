// IRONLOG/api/routes/alerts.routes.js
import { db } from "../db/client.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

export default async function alertsRoutes(app) {
  // GET /api/alerts?asof=2026-02-27
  app.get("/", async (req, reply) => {
    const asof = String(req.query?.asof || "").trim();

    if (asof && !isDate(asof)) {
      return reply.code(400).send({ error: "asof must be YYYY-MM-DD" });
    }

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
      LIMIT 100
    `).all().map(r => ({ ...r, critical: Boolean(r.critical), on_hand: Number(r.on_hand) }));

    const criticalOnOrder = db.prepare(`
      SELECT
        p.part_code,
        p.part_name,
        po.quantity,
        po.expected_date,
        po.status
      FROM parts_orders po
      JOIN parts p ON p.id = po.part_id
      WHERE p.critical = 1 AND po.status != 'received'
      ORDER BY po.expected_date ASC
      LIMIT 100
    `).all();

    const openWorkOrders = db.prepare(`
      SELECT w.id, a.asset_code, w.source, w.status, w.opened_at
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      WHERE w.status != 'closed'
      ORDER BY w.id DESC
      LIMIT 100
    `).all();

    const overdueMaintenance = db.prepare(`
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
            ${asof ? "AND dh.work_date <= ?" : ""}
        ), 0) AS current_hours
      FROM maintenance_plans mp
      JOIN assets a ON a.id = mp.asset_id
      WHERE mp.active = 1
        AND a.active = 1
        AND a.is_standby = 0
    `);

    const plans = asof ? overdueMaintenance.all(asof) : overdueMaintenance.all();

    const overdue = plans.map(r => {
      const current = Number(r.current_hours || 0);
      const next_due = Number(r.last_service_hours || 0) + Number(r.interval_hours || 0);
      const remaining = next_due - current;
      return {
        ...r,
        current_hours: Number(current.toFixed(2)),
        next_due_hours: Number(next_due.toFixed(2)),
        overdue_by_hours: remaining < 0 ? Number(Math.abs(remaining).toFixed(2)) : 0,
        is_overdue: remaining <= 0
      };
    }).filter(x => x.is_overdue)
      .sort((a,b) => b.overdue_by_hours - a.overdue_by_hours)
      .slice(0, 100);

    return {
      ok: true,
      asof: asof || null,
      alerts: {
        low_stock: { count: lowStock.length, items: lowStock },
        critical_on_order: { count: criticalOnOrder.length, items: criticalOnOrder },
        open_work_orders: { count: openWorkOrders.length, items: openWorkOrders },
        overdue_maintenance: { count: overdue.length, items: overdue }
      }
    };
  });
}