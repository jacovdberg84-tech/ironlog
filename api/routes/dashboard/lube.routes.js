// IRONLOG/api/routes/dashboard/lube.routes.js — Lube usage, analytics and mappings.
// Registered by routes/dashboard.routes.js; shared helpers arrive through ctx.
import { db } from "../../db/client.js";
import { fetchLubeUsageLines } from "../../utils/lubeUsageLines.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerLubeRoutes(app, ctx) {
  const { requireRoles, todayYYYYMMDD } = ctx;

  // GET /api/dashboard/lube?start=YYYY-MM-DD&end=YYYY-MM-DD
  app.get("/lube", async (req, reply) => {
    try {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
      return reply.code(400).send({ error: "start and end must be YYYY-MM-DD" });
    }

    const defaultLubeCost = Number(
      db.prepare(`SELECT value FROM cost_settings WHERE key = 'lube_cost_per_qty_default' LIMIT 1`).get()?.value
    );
    const lubeUnitFallback = Number.isFinite(defaultLubeCost) && defaultLubeCost > 0 ? defaultLubeCost : 4.0;

    const rows = db.prepare(`
      SELECT
        a.asset_code,
        a.asset_name,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS total_lube_cost,
        COUNT(*) AS entries
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.id
      ORDER BY qty_total DESC, a.asset_code ASC
      LIMIT 200
    `).all(lubeUnitFallback, start, end).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      qty_total: Number(r.qty_total || 0),
      total_lube_cost: Number(r.total_lube_cost || 0),
      entries: Number(r.entries || 0),
    }));

    const byTypeRows = db.prepare(`
      SELECT
        a.asset_code,
        CASE
          WHEN LOWER(TRIM(COALESCE(ol.oil_type, ''))) IN ('admin','supervisor','manager','stores','artisan','operator') THEN 'UNSPECIFIED'
          ELSE COALESCE(NULLIF(TRIM(ol.oil_type), ''), 'UNSPECIFIED')
        END AS oil_type,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS total_lube_cost
      FROM oil_logs ol
      JOIN assets a ON a.id = ol.asset_id
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY a.asset_code, oil_type
      ORDER BY a.asset_code ASC, qty_total DESC, oil_type ASC
      LIMIT 1200
    `).all(lubeUnitFallback, start, end);
    const byTypeLookup = new Map();
    for (const r of byTypeRows) {
      const code = String(r.asset_code || "");
      if (!byTypeLookup.has(code)) byTypeLookup.set(code, []);
      byTypeLookup.get(code).push({
        oil_type: String(r.oil_type || "UNSPECIFIED"),
        qty_total: Number(r.qty_total || 0),
        total_lube_cost: Number(r.total_lube_cost || 0),
      });
    }
    for (const row of rows) {
      row.by_oil_type = byTypeLookup.get(String(row.asset_code || "")) || [];
    }

    const summary = db.prepare(`
      SELECT
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COALESCE(SUM(ol.quantity * COALESCE(ol.unit_cost, ?)), 0) AS total_lube_cost,
        COUNT(*) AS entries,
        COUNT(DISTINCT ol.asset_id) AS assets
      FROM oil_logs ol
      WHERE ol.log_date BETWEEN ? AND ?
    `).get(lubeUnitFallback, start, end);

    const lines = fetchLubeUsageLines(db, { start, end, lubeUnitFallback });
    const lineSummary = lines.reduce(
      (acc, r) => {
        acc.qty_total += Number(r.quantity || 0);
        acc.line_cost += Number(r.line_cost || 0);
        acc.entries += 1;
        if (r.asset_code) acc.assets.add(String(r.asset_code));
        return acc;
      },
      { qty_total: 0, line_cost: 0, entries: 0, assets: new Set() },
    );

    return {
      ok: true,
      start,
      end,
      summary: {
        qty_total: Number(summary?.qty_total || 0),
        total_lube_cost: Number(summary?.total_lube_cost || 0),
        entries: Number(summary?.entries || 0),
        assets: Number(summary?.assets || 0),
        line_entries: lineSummary.entries,
        line_qty_total: Number(lineSummary.qty_total.toFixed(3)),
        line_total_cost: Number(lineSummary.line_cost.toFixed(2)),
        line_assets: lineSummary.assets.size,
      },
      rows,
      lines,
    };
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ error: err.message || String(err) });
    }
  });

  // GET /api/dashboard/lube/analytics?months=6
  // Usage by oil type/stock code + monthly trend + low-stock forecast
  app.get("/lube/analytics", async (req, reply) => {
    const monthsRaw = Number(req.query?.months ?? 6);
    const months = Number.isFinite(monthsRaw) ? Math.max(1, Math.min(24, Math.trunc(monthsRaw))) : 6;

    const endDate = todayYYYYMMDD();
    const endObj = new Date(`${endDate}T00:00:00`);
    const startObj = new Date(endObj);
    startObj.setMonth(startObj.getMonth() - (months - 1));
    startObj.setDate(1);
    const startDate = startObj.toISOString().slice(0, 10);

    const daySpan = Math.max(1, Math.round((Date.parse(`${endDate}T00:00:00`) - Date.parse(`${startDate}T00:00:00`)) / 86400000) + 1);

    const byType = db.prepare(`
      SELECT
        CASE
          WHEN LOWER(TRIM(COALESCE(ol.oil_type, ''))) IN ('admin','supervisor','manager','stores','artisan','operator') THEN 'UNSPECIFIED'
          ELSE COALESCE(NULLIF(TRIM(ol.oil_type), ''), 'UNSPECIFIED')
        END AS oil_key,
        COALESCE(SUM(ol.quantity), 0) AS qty_total,
        COUNT(*) AS entries
      FROM oil_logs ol
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY oil_key
      ORDER BY qty_total DESC
      LIMIT 200
    `).all(startDate, endDate).map((r) => ({
      oil_key: r.oil_key,
      qty_total: Number(r.qty_total || 0),
      entries: Number(r.entries || 0),
    }));

    const trend = db.prepare(`
      SELECT
        SUBSTR(ol.log_date, 1, 7) AS month,
        CASE
          WHEN LOWER(TRIM(COALESCE(ol.oil_type, ''))) IN ('admin','supervisor','manager','stores','artisan','operator') THEN 'UNSPECIFIED'
          ELSE COALESCE(NULLIF(TRIM(ol.oil_type), ''), 'UNSPECIFIED')
        END AS oil_key,
        COALESCE(SUM(ol.quantity), 0) AS qty
      FROM oil_logs ol
      WHERE ol.log_date BETWEEN ? AND ?
      GROUP BY month, oil_key
      ORDER BY month ASC, qty DESC
      LIMIT 800
    `).all(startDate, endDate).map((r) => ({
      month: r.month,
      oil_key: r.oil_key,
      qty: Number(r.qty || 0),
    }));

    const mapRows = db.prepare(`
      SELECT oil_key, part_code
      FROM lube_type_mappings
    `).all();
    const mapping = new Map(
      mapRows.map((m) => [String(m.oil_key || "").trim().toLowerCase(), String(m.part_code || "").trim()])
    );

    const getPart = db.prepare(`
      SELECT part_code, part_name, min_stock, IFNULL(SUM(sm.quantity), 0) AS on_hand
      FROM parts p
      LEFT JOIN stock_movements sm ON sm.part_id = p.id
      WHERE p.part_code = ?
      GROUP BY p.id
      LIMIT 1
    `);

    const forecast = byType.map((r) => {
      const mappedCode = mapping.get(String(r.oil_key || "").trim().toLowerCase()) || String(r.oil_key || "");
      const p = getPart.get(mappedCode);
      const avg_daily_use = Number((r.qty_total / daySpan).toFixed(3));
      const on_hand = p ? Number(p.on_hand || 0) : null;
      const min_stock = p ? Number(p.min_stock || 0) : null;
      const days_to_min =
        p && avg_daily_use > 0
          ? Number(((on_hand - min_stock) / avg_daily_use).toFixed(1))
          : null;
      return {
        oil_key: r.oil_key,
        qty_total: Number(r.qty_total.toFixed(2)),
        entries: r.entries,
        avg_daily_use,
        mapped_part_code: mappedCode || null,
        part_code: p?.part_code || null,
        part_name: p?.part_name || null,
        on_hand,
        min_stock,
        days_to_min,
        low_risk: p ? (on_hand <= min_stock || (days_to_min != null && days_to_min <= 30)) : false,
      };
    }).sort((a, b) => {
      const ar = Number(Boolean(a.low_risk));
      const br = Number(Boolean(b.low_risk));
      if (ar !== br) return br - ar;
      const ad = a.days_to_min == null ? 999999 : a.days_to_min;
      const bd = b.days_to_min == null ? 999999 : b.days_to_min;
      return ad - bd;
    });

    return reply.send({
      ok: true,
      start: startDate,
      end: endDate,
      months,
      summary: {
        oils: byType.length,
        qty_total: Number(byType.reduce((acc, x) => acc + Number(x.qty_total || 0), 0).toFixed(2)),
        low_risk_count: forecast.filter((x) => x.low_risk).length,
      },
      by_type: byType,
      trend,
      forecast,
    });
  });

  // GET /api/dashboard/lube/mappings
  app.get("/lube/mappings", async () => {
    const rows = db.prepare(`
      SELECT oil_key, part_code, updated_by, updated_at
      FROM lube_type_mappings
      ORDER BY oil_key ASC
      LIMIT 400
    `).all();
    return { ok: true, rows };
  });

  // POST /api/dashboard/lube/mappings
  // Body: { oil_key, part_code }
  app.post("/lube/mappings", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const oil_key = String(req.body?.oil_key || "").trim();
    const part_code = String(req.body?.part_code || "").trim();
    if (!oil_key) return reply.code(400).send({ error: "oil_key is required" });
    if (!part_code) return reply.code(400).send({ error: "part_code is required" });

    const part = db.prepare(`
      SELECT id, part_code, part_name
      FROM parts
      WHERE part_code = ?
    `).get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });

    const user = String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
    db.prepare(`
      INSERT INTO lube_type_mappings (oil_key, part_code, updated_by, updated_at)
      VALUES (?, ?, ?, datetime('now'))
      ON CONFLICT(oil_key) DO UPDATE SET
        part_code = excluded.part_code,
        updated_by = excluded.updated_by,
        updated_at = datetime('now')
    `).run(oil_key, part_code, user);

    writeAudit(db, req, {
      module: "lube",
      action: "mapping_upsert",
      entity_type: "lube_mapping",
      entity_id: oil_key,
      payload: { oil_key, part_code },
    });

    return { ok: true, oil_key, part_code, part_name: part.part_name };
  });
}
