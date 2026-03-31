// GET /api/assets/:asset_code/history?start=YYYY-MM-DD&end=YYYY-MM-DD
app.get("/:asset_code/history", async (req, reply) => {
  const asset_code = String(req.params.asset_code || "").trim();
  const start = String(req.query?.start || "").trim();
  const end = String(req.query?.end || "").trim();

  const asset = db.prepare(`SELECT id, asset_code, asset_name, category FROM assets WHERE asset_code = ?`).get(asset_code);
  if (!asset) return reply.code(404).send({ error: "Asset not found" });

  const whereDate = (col) => {
    // If no range provided, just ignore
    const clauses = [];
    const params = [];
    if (start && /^\d{4}-\d{2}-\d{2}$/.test(start)) { clauses.push(`${col} >= ?`); params.push(start); }
    if (end && /^\d{4}-\d{2}-\d{2}$/.test(end)) { clauses.push(`${col} <= ?`); params.push(end); }
    return { sql: clauses.length ? " AND " + clauses.join(" AND ") : "", params };
  };

  // ---- Breakdowns (with downtime total for the day range)
  const bdDate = whereDate("b.breakdown_date");
  const breakdowns = db.prepare(`
    SELECT
      b.id,
      b.breakdown_date AS date,
      b.description AS title,
      b.critical,
      COALESCE(SUM(l.hours_down), 0) AS downtime_hours
    FROM breakdowns b
    LEFT JOIN breakdown_downtime_logs l ON l.breakdown_id = b.id
    WHERE b.asset_id = ? ${bdDate.sql}
    GROUP BY b.id
    ORDER BY b.breakdown_date DESC, b.id DESC
    LIMIT 200
  `).all(asset.id, ...bdDate.params).map(r => ({
    type: "breakdown",
    date: r.date,
    title: r.title,
    work_order_id: db.prepare(`
      SELECT id FROM work_orders
      WHERE source='breakdown' AND reference_id=?
      ORDER BY id DESC LIMIT 1
    `).get(r.id)?.id ?? null,
    details: {
      breakdown_id: r.id,
      critical: Boolean(r.critical),
      downtime_hours: Number(r.downtime_hours || 0),
    }
  }));

  // ---- Work orders (manual/service) for this asset
  const woDate = whereDate("DATE(w.opened_at)");
  const workOrders = db.prepare(`
    SELECT
      w.id,
      DATE(w.opened_at) AS date,
      w.source,
      w.status,
      w.opened_at,
      w.closed_at
    FROM work_orders w
    WHERE w.asset_id = ? ${woDate.sql}
    ORDER BY w.id DESC
    LIMIT 200
  `).all(asset.id, ...woDate.params).map(r => ({
    type: "work_order",
    date: r.date,
    title: `WO #${r.id} (${r.source})`,
    work_order_id: r.id,
    details: {
      status: r.status,
      opened_at: r.opened_at,
      closed_at: r.closed_at
    }
  }));

  // ---- GET slips
  const gsDate = whereDate("g.slip_date");
  const getSlips = db.prepare(`
    SELECT g.id, g.slip_date AS date, g.location, g.notes
    FROM get_change_slips g
    WHERE g.asset_id = ? ${gsDate.sql}
    ORDER BY g.slip_date DESC, g.id DESC
    LIMIT 200
  `).all(asset.id, ...gsDate.params).map(g => {
    const items = db.prepare(`
      SELECT position, part_code, part_name, qty, reason
      FROM get_change_items
      WHERE slip_id = ?
      ORDER BY id ASC
    `).all(g.id);
    const wo = db.prepare(`
      SELECT work_order_id FROM work_order_links
      WHERE link_type='get_slip' AND link_id=?
      ORDER BY id DESC LIMIT 1
    `).get(g.id);

    return {
      type: "get_slip",
      date: g.date,
      title: `GET Change Slip #${g.id}`,
      work_order_id: wo?.work_order_id ?? null,
      details: { slip_id: g.id, location: g.location, notes: g.notes, items }
    };
  });

  // ---- Component slips
  const csDate = whereDate("c.slip_date");
  const componentSlips = db.prepare(`
    SELECT
      c.id,
      c.slip_date AS date,
      c.component,
      c.serial_out,
      c.serial_in,
      c.hours_at_change,
      c.notes
    FROM component_change_slips c
    WHERE c.asset_id = ? ${csDate.sql}
    ORDER BY c.slip_date DESC, c.id DESC
    LIMIT 200
  `).all(asset.id, ...csDate.params).map(c => {
    const wo = db.prepare(`
      SELECT work_order_id FROM work_order_links
      WHERE link_type='component_slip' AND link_id=?
      ORDER BY id DESC LIMIT 1
    `).get(c.id);

    return {
      type: "component_slip",
      date: c.date,
      title: `Component Change: ${c.component} (#${c.id})`,
      work_order_id: wo?.work_order_id ?? null,
      details: {
        slip_id: c.id,
        component: c.component,
        serial_out: c.serial_out,
        serial_in: c.serial_in,
        hours_at_change: c.hours_at_change,
        notes: c.notes
      }
    };
  });

  // ---- Merge + sort newest first
  const all = [...breakdowns, ...workOrders, ...getSlips, ...componentSlips]
    .sort((a, b) => String(b.date).localeCompare(String(a.date)));

  return reply.send({
    ok: true,
    asset,
    history: all
  });
});