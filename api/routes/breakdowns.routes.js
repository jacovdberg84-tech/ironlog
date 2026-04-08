// IRONLOG/api/routes/breakdowns.routes.js
import { db } from "../db/client.js";

/** Parse YYYY-MM-DDTHH:mm (no timezone) as local wall time. */
function parseLocalDateTime(s) {
  const m = String(s || "")
    .trim()
    .match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  const h = Number(m[4]);
  const mi = Number(m[5]);
  const se = m[6] != null ? Number(m[6]) : 0;
  if ([y, mo, d, h, mi, se].some((n) => Number.isNaN(n))) return null;
  return new Date(y, mo - 1, d, h, mi, se, 0);
}

/** Split downtime across calendar days for breakdown_downtime_logs. */
function downtimeSegmentsFromRange(timeDownStr, timeUpStr) {
  const d0 = parseLocalDateTime(timeDownStr);
  const d1 = parseLocalDateTime(timeUpStr);
  if (!d0 || !d1 || d1 <= d0) return null;
  const out = [];
  let cur = new Date(d0);
  while (cur < d1) {
    const y = cur.getFullYear();
    const mo = String(cur.getMonth() + 1).padStart(2, "0");
    const day = String(cur.getDate()).padStart(2, "0");
    const logDate = `${y}-${mo}-${day}`;
    const nextMidnight = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 1, 0, 0, 0, 0);
    const segmentEnd = d1 < nextMidnight ? d1 : nextMidnight;
    const hours = (segmentEnd - cur) / 3600000;
    if (hours <= 0 || !Number.isFinite(hours)) {
      cur = new Date(nextMidnight);
      continue;
    }
    const capped = Math.min(24, hours);
    out.push({ log_date: logDate, hours_down: Number(capped.toFixed(4)) });
    cur = segmentEnd;
  }
  return out.length ? out : null;
}

function toSqliteWallDatetime(s) {
  const m = String(s || "")
    .trim()
    .match(/^(\d{4}-\d{2}-\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return null;
  const sec = m[4] != null ? String(m[4]).padStart(2, "0") : "00";
  return `${m[1]} ${m[2]}:${m[3]}:${sec}`;
}

export default async function breakdownRoutes(app) {
  /* =====================================================
     PREPARED STATEMENTS
  ===================================================== */
  const listOpenComponentWOs = db.prepare(`
  SELECT
    bc.id as component_line_id,
    bc.component,
    bc.symptom,
    wo.id as work_order_id,
    wo.status
  FROM breakdown_components bc
  JOIN work_orders wo ON wo.id = bc.work_order_id
  WHERE bc.breakdown_id = ?
    AND wo.status = 'open'
  ORDER BY wo.id DESC
`);

  const getAssetByCode = db.prepare(`
    SELECT id, asset_code, asset_name
    FROM assets
    WHERE asset_code = ?
  `);

  const getBreakdownById = db.prepare(`
    SELECT b.*, a.asset_code, a.asset_name, a.category
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    WHERE b.id = ?
  `);

  const getOpenBreakdownByAssetId = db.prepare(`
    SELECT b.id, b.primary_work_order_id, b.breakdown_date, b.start_at
    FROM breakdowns b
    WHERE b.asset_id = ?
      AND b.status = 'OPEN'
    ORDER BY b.id DESC
    LIMIT 1
  `);

  const listOpenBreakdownsAll = db.prepare(`
    SELECT
      b.id,
      b.asset_id,
      a.asset_code,
      b.breakdown_date,
      b.start_at,
      b.description,
      b.primary_work_order_id,
      wo.status AS primary_work_order_status
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    LEFT JOIN work_orders wo ON wo.id = b.primary_work_order_id
    WHERE b.status = 'OPEN'
      AND (wo.status IS NULL OR wo.status NOT IN ('completed','approved','closed'))
    ORDER BY b.id DESC
    LIMIT 500
  `);
  const listOpenBreakdownsAllForDate = db.prepare(`
    SELECT
      b.id,
      b.asset_id,
      a.asset_code,
      b.breakdown_date,
      b.start_at,
      b.description,
      b.primary_work_order_id,
      wo.status AS primary_work_order_status,
      COALESCE((
        SELECT SUM(l.hours_down)
        FROM breakdown_downtime_logs l
        WHERE l.breakdown_id = b.id
          AND l.log_date = ?
      ), 0) AS hours_down_for_date
    FROM breakdowns b
    JOIN assets a ON a.id = b.asset_id
    LEFT JOIN work_orders wo ON wo.id = b.primary_work_order_id
    WHERE b.status = 'OPEN'
      AND (wo.status IS NULL OR wo.status NOT IN ('completed','approved','closed'))
    ORDER BY b.id DESC
    LIMIT 500
  `);

  /* ------------------
     BREAKDOWN CORE
  ------------------ */

  // UPDATED: includes GET fields
  const insertBreakdown = db.prepare(`
    INSERT INTO breakdowns (
      asset_id,
      breakdown_date,
      status,
      start_at,
      description,
      component,
      critical,
      downtime_total_hours,
      primary_work_order_id,
      get_used,
      get_hours_fitted,
      get_hours_changed
    )
    VALUES (?, ?, 'OPEN', datetime('now'), ?, ?, ?, 0, NULL, ?, ?, ?)
  `);

  const insertWorkOrder = db.prepare(`
    INSERT INTO work_orders (asset_id, source, reference_id, status)
    VALUES (?, 'breakdown', ?, 'open')
  `);

  const linkPrimaryWO = db.prepare(`
    UPDATE breakdowns
    SET primary_work_order_id = ?
    WHERE id = ?
  `);
  const updateBreakdownStart = db.prepare(`
    UPDATE breakdowns
    SET
      breakdown_date = COALESCE(?, breakdown_date),
      start_at = COALESCE(?, start_at)
    WHERE id = ?
      AND status = 'OPEN'
  `);

  const insertBreakdownShortClosed = db.prepare(`
    INSERT INTO breakdowns (
      asset_id,
      breakdown_date,
      status,
      start_at,
      end_at,
      description,
      component,
      critical,
      downtime_total_hours,
      primary_work_order_id,
      get_used,
      get_hours_fitted,
      get_hours_changed
    )
    VALUES (?, ?, 'CLOSED', ?, ?, ?, ?, ?, 0, NULL, 0, NULL, NULL)
  `);

  const closeWorkOrderQuick = db.prepare(`
    UPDATE work_orders SET status = 'closed', closed_at = datetime('now') WHERE id = ?
  `);

  const insertOilLog = db.prepare(`
    INSERT INTO oil_logs (asset_id, log_date, oil_type, quantity)
    VALUES (?, ?, ?, ?)
  `);

  const closeBreakdown = db.prepare(`
    UPDATE breakdowns
    SET status = 'CLOSED',
        end_at = datetime('now')
    WHERE id = ?
  `);

  const reopenBreakdown = db.prepare(`
    UPDATE breakdowns
    SET status = 'OPEN',
        end_at = NULL
    WHERE id = ?
  `);

  /* ------------------
     DOWNTIME LOGGING
  ------------------ */

  const upsertDowntimeLog = db.prepare(`
    INSERT INTO breakdown_downtime_logs (breakdown_id, log_date, hours_down, notes)
    VALUES (?, ?, ?, ?)
    ON CONFLICT(breakdown_id, log_date)
    DO UPDATE SET
      hours_down = excluded.hours_down,
      notes = excluded.notes,
      updated_at = datetime('now')
  `);

  const listDowntimeLogs = db.prepare(`
    SELECT id, log_date, hours_down, notes, created_at, updated_at
    FROM breakdown_downtime_logs
    WHERE breakdown_id = ?
    ORDER BY log_date DESC
  `);

  /* ------------------
     COMPONENT LINES
  ------------------ */

  const insertComponentLine = db.prepare(`
    INSERT INTO breakdown_components (breakdown_id, component, symptom)
    VALUES (?, ?, ?)
  `);

  const listComponentLines = db.prepare(`
    SELECT id, component, symptom, work_order_id, created_at
    FROM breakdown_components
    WHERE breakdown_id = ?
    ORDER BY id DESC
  `);

  const getComponentLine = db.prepare(`
    SELECT *
    FROM breakdown_components
    WHERE id = ? AND breakdown_id = ?
  `);

  const updateComponentLine = db.prepare(`
    UPDATE breakdown_components
    SET component = COALESCE(?, component),
        symptom  = COALESCE(?, symptom)
    WHERE id = ? AND breakdown_id = ?
  `);

  const deleteComponentLine = db.prepare(`
    DELETE FROM breakdown_components
    WHERE id = ? AND breakdown_id = ? AND work_order_id IS NULL
  `);

  const linkComponentWO = db.prepare(`
    UPDATE breakdown_components
    SET work_order_id = ?
    WHERE id = ? AND breakdown_id = ? AND work_order_id IS NULL
  `);

  /* =====================================================
     HELPERS
  ===================================================== */

  function isDate(s) {
    return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
  }

  function num(n) {
    const x = Number(n);
    return Number.isFinite(x) ? x : NaN;
  }

  function toBoolInt(v) {
    // handles true/false, 1/0, "1"/"0"
    if (v === true) return 1;
    if (v === false) return 0;
    const n = Number(v);
    return n === 1 ? 1 : 0;
  }

  function optPositiveNumber(v) {
    if (v === null || v === undefined || v === "") return null;
    const n = Number(v);
    return Number.isFinite(n) ? n : NaN;
  }

  function validateGetFields(body, reply) {
    const get_used = toBoolInt(body?.get_used);
    const get_hours_fitted = optPositiveNumber(body?.get_hours_fitted);
    const get_hours_changed = optPositiveNumber(body?.get_hours_changed);

    if (get_used) {
      if (get_hours_fitted == null || Number.isNaN(get_hours_fitted) || get_hours_fitted <= 0) {
        reply.code(400).send({ error: "GET used requires get_hours_fitted (> 0)" });
        return null;
      }
      if (get_hours_changed == null || Number.isNaN(get_hours_changed) || get_hours_changed <= 0) {
        reply.code(400).send({ error: "GET used requires get_hours_changed (> 0)" });
        return null;
      }
    }

    return {
      get_used,
      get_hours_fitted: get_used ? get_hours_fitted : null,
      get_hours_changed: get_used ? get_hours_changed : null,
    };
  }

  /* =====================================================
     ROUTES
     Base assumed: /api/breakdowns
  ===================================================== */

  // ---------------------------
  // List breakdowns
  // ---------------------------
  app.get("/", async () => {
    const rows = db.prepare(`
      SELECT
        b.id,
        b.breakdown_date,
        b.status,
        b.start_at,
        b.end_at,
        b.description,
        b.component,
        b.downtime_total_hours,
        b.critical,
        b.primary_work_order_id,
        b.get_used,
        b.get_hours_fitted,
        b.get_hours_changed,
        b.created_at,
        a.asset_code,
        a.asset_name,
        a.category
      FROM breakdowns b
      JOIN assets a ON a.id = b.asset_id
      ORDER BY b.id DESC
      LIMIT 200
    `).all();

    return rows.map(r => ({
      ...r,
      critical: Boolean(r.critical),
      downtime_total_hours: Number(r.downtime_total_hours || 0),
      get_used: Boolean(r.get_used),
      get_hours_fitted: r.get_hours_fitted == null ? null : Number(r.get_hours_fitted),
      get_hours_changed: r.get_hours_changed == null ? null : Number(r.get_hours_changed),
    }));
  });

  // ---------------------------
  // Get one breakdown (with logs + components)
  // ---------------------------
  app.get("/:id", async (req, reply) => {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) return reply.code(400).send({ error: "Invalid breakdown id" });

    const b = getBreakdownById.get(id);
    if (!b) return reply.code(404).send({ error: "Breakdown not found" });

    return {
      ...b,
      critical: Boolean(b.critical),
      downtime_total_hours: Number(b.downtime_total_hours || 0),
      get_used: Boolean(b.get_used),
      get_hours_fitted: b.get_hours_fitted == null ? null : Number(b.get_hours_fitted),
      get_hours_changed: b.get_hours_changed == null ? null : Number(b.get_hours_changed),
      downtime_logs: listDowntimeLogs.all(id).map(l => ({ ...l, hours_down: Number(l.hours_down || 0) })),
      components: listComponentLines.all(id),
    };
  });

  // ---------------------------
  // Find open breakdown for asset_code (Daily Input helper)
  // GET /api/breakdowns/open?asset_code=A300AM
  // ---------------------------
  app.get("/open", async (req, reply) => {
    const asset_code = String(req.query.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ error: "asset_code is required" });

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const open = getOpenBreakdownByAssetId.get(asset.id);
    return { ok: true, breakdown: open || null };
  });

  // ---------------------------
  // List all open breakdowns (Daily Input helper)
  // GET /api/breakdowns/open-all
  // ---------------------------
  app.get("/open-all", async (req, reply) => {
    const date = String(req.query?.date || "").trim();
    const rows = isDate(date)
      ? listOpenBreakdownsAllForDate.all(date)
      : listOpenBreakdownsAll.all();
    return { ok: true, rows: rows.map((r) => ({ ...r })) };
  });

  // ---------------------------
  // Ensure open breakdown exists (Daily Input helper)
  // POST /api/breakdowns/ensure-open
  // { asset_code, breakdown_date, description?, component?, critical?, get_used?, get_hours_fitted?, get_hours_changed? }
  // ---------------------------
  app.post("/ensure-open", async (req, reply) => {
    const body = req.body || {};
    const asset_code = String(body.asset_code || "").trim();
    const breakdown_date = String(body.breakdown_date || "").trim(); // YYYY-MM-DD
    const start_date = String(body.start_date || "").trim(); // optional YYYY-MM-DD manual correction
    const description = String(body.description || "Down - Daily Input").trim();
    const component = body.component ? String(body.component).trim() : null;
    const critical = body.critical ? 1 : 0;

    if (!asset_code || !isDate(breakdown_date)) {
      return reply.code(400).send({ error: "asset_code and breakdown_date(YYYY-MM-DD) required" });
    }
    if (start_date && !isDate(start_date)) {
      return reply.code(400).send({ error: "start_date must be YYYY-MM-DD when provided" });
    }
    const effectiveStartDate = start_date || breakdown_date;

    // GET validation (optional input)
    const getPack = validateGetFields(body, reply);
    if (!getPack) return;

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const existing = getOpenBreakdownByAssetId.get(asset.id);
    if (existing) {
      if (start_date) {
        updateBreakdownStart.run(
          effectiveStartDate,
          `${effectiveStartDate} 00:00:00`,
          Number(existing.id)
        );
      }
      return reply.send({
        ok: true,
        breakdown_id: existing.id,
        primary_work_order_id: existing.primary_work_order_id,
        created: false,
      });
    }

    const tx = db.transaction(() => {
      const b = insertBreakdown.run(
        asset.id,
        effectiveStartDate,
        description,
        component,
        critical,
        getPack.get_used,
        getPack.get_hours_fitted,
        getPack.get_hours_changed
      );

      const breakdownId = Number(b.lastInsertRowid);

      const wo = insertWorkOrder.run(asset.id, breakdownId);
      const workOrderId = Number(wo.lastInsertRowid);

      linkPrimaryWO.run(workOrderId, breakdownId);
      if (start_date) {
        updateBreakdownStart.run(
          effectiveStartDate,
          `${effectiveStartDate} 00:00:00`,
          breakdownId
        );
      }

      return { breakdownId, workOrderId };
    });

    const r = tx();
    return reply.code(201).send({
      ok: true,
      breakdown_id: r.breakdownId,
      primary_work_order_id: r.workOrderId,
      created: true,
    });
  });

  // ---------------------------
  // Create breakdown (manual create)
  // ---------------------------
  app.post("/", async (req, reply) => {
    const body = req.body || {};
    const asset_code = String(body.asset_code || "").trim();
    const breakdown_date = String(body.breakdown_date || "").trim();
    const description = String(body.description || "").trim();
    const component = body.component ? String(body.component).trim() : null;
    const critical = body.critical ? 1 : 0;

    if (!asset_code || !isDate(breakdown_date) || !description) {
      return reply.code(400).send({ error: "asset_code, breakdown_date, description required" });
    }

    // GET validation (required if get_used true)
    const getPack = validateGetFields(body, reply);
    if (!getPack) return;

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const tx = db.transaction(() => {
      const b = insertBreakdown.run(
        asset.id,
        breakdown_date,
        description,
        component,
        critical,
        getPack.get_used,
        getPack.get_hours_fitted,
        getPack.get_hours_changed
      );

      const breakdownId = Number(b.lastInsertRowid);

      const wo = insertWorkOrder.run(asset.id, breakdownId);
      const workOrderId = Number(wo.lastInsertRowid);

      linkPrimaryWO.run(workOrderId, breakdownId);

      return { breakdownId, workOrderId };
    });

    const result = tx();

    return reply.code(201).send({
      ok: true,
      breakdown_id: result.breakdownId,
      primary_work_order_id: result.workOrderId,
      get_used: Boolean(getPack.get_used),
      get_hours_fitted: getPack.get_hours_fitted,
      get_hours_changed: getPack.get_hours_changed,
    });
  });

  // ---------------------------
  // Short breakdown (quick close): downtime logs → fleet hours, parts to WO, oils to oil_logs
  // POST /api/breakdowns/short-complete
  // Body: { asset_code, breakdown_date, description, component?, critical?,
  //         time_down?, time_up? (YYYY-MM-DDTHH:mm) OR hours_down (0..24),
  //         parts?: [{ part_code, quantity }], oils?: [{ oil_type, quantity }] }
  // ---------------------------
  app.post("/short-complete", async (req, reply) => {
    const body = req.body || {};
    const asset_code = String(body.asset_code || "").trim();
    const breakdown_date = String(body.breakdown_date || "").trim();
    const description = String(body.description || "").trim();
    const component = body.component ? String(body.component).trim() : null;
    const critical = body.critical ? 1 : 0;
    const time_down = body.time_down ? String(body.time_down).trim() : "";
    const time_up = body.time_up ? String(body.time_up).trim() : "";

    if (!asset_code || !isDate(breakdown_date) || !description) {
      return reply.code(400).send({ error: "asset_code, breakdown_date (YYYY-MM-DD), description required" });
    }

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    let segments;
    let startAt = null;
    let endAt = null;

    if (time_down && time_up) {
      segments = downtimeSegmentsFromRange(time_down, time_up);
      if (!segments || !segments.length) {
        return reply.code(400).send({ error: "time_down and time_up must be valid (time_up after time_down)" });
      }
      startAt = toSqliteWallDatetime(time_down);
      endAt = toSqliteWallDatetime(time_up);
    } else {
      const h = num(body.hours_down);
      if (Number.isNaN(h) || h <= 0 || h > 24) {
        return reply.code(400).send({ error: "Either time_down + time_up, or hours_down between 0 and 24, required" });
      }
      segments = [{ log_date: breakdown_date, hours_down: Number(h.toFixed(4)) }];
    }

    const parts = Array.isArray(body.parts) ? body.parts : [];
    const oils = Array.isArray(body.oils) ? body.oils : [];

    try {
      const result = db.transaction(() => {
        const ins = insertBreakdownShortClosed.run(
          asset.id,
          breakdown_date,
          startAt,
          endAt,
          description,
          component,
          critical,
        );
        const breakdownId = Number(ins.lastInsertRowid);

        const wo = insertWorkOrder.run(asset.id, breakdownId);
        const woId = Number(wo.lastInsertRowid);
        linkPrimaryWO.run(woId, breakdownId);

        for (const seg of segments) {
          upsertDowntimeLog.run(
            breakdownId,
            seg.log_date,
            seg.hours_down,
            `Short breakdown — ${description}`,
          );
        }

        closeWorkOrderQuick.run(woId);

        const issued = [];
        for (const p of parts) {
          const part_code = String(p.part_code || "").trim();
          const quantity = Number(p.quantity ?? 0);
          if (!part_code || !Number.isFinite(quantity) || quantity <= 0) continue;
          const partRow = db.prepare(`SELECT id FROM parts WHERE part_code = ?`).get(part_code);
          if (!partRow) throw new Error(`part_code not found: ${part_code}`);
          const onHandRow = db.prepare(`
            SELECT IFNULL(SUM(quantity), 0) AS on_hand FROM stock_movements WHERE part_id = ?
          `).get(partRow.id);
          const on_hand = Number(onHandRow.on_hand || 0);
          if (on_hand < quantity) {
            throw new Error(`insufficient stock for ${part_code}: have ${on_hand}, need ${quantity}`);
          }
          db.prepare(`
            INSERT INTO stock_movements (part_id, quantity, movement_type, reference)
            VALUES (?, ?, 'out', ?)
          `).run(partRow.id, -Math.abs(Math.trunc(quantity)), `work_order:${woId}`);
          issued.push({ part_code, quantity: Math.trunc(quantity) });
        }

        const oilsLogged = [];
        for (const o of oils) {
          const oil_type = String(o.oil_type || "").trim();
          const quantity = Number(o.quantity ?? 0);
          if (!oil_type || !Number.isFinite(quantity) || quantity <= 0) continue;
          insertOilLog.run(asset.id, breakdown_date, oil_type, quantity);
          oilsLogged.push({ oil_type, quantity: Number(quantity.toFixed(3)) });
        }

        const updated = getBreakdownById.get(breakdownId);
        return {
          breakdown_id: breakdownId,
          work_order_id: woId,
          downtime_total_hours: Number(updated?.downtime_total_hours || 0),
          parts_issued: issued,
          oils_logged: oilsLogged,
        };
      })();

      return reply.code(201).send({ ok: true, ...result });
    } catch (e) {
      req.log.error(e);
      return reply.code(400).send({ error: e.message || String(e) });
    }
  });

  // ---------------------------
  // Log downtime (no WO creation)
  // ---------------------------
  app.post("/:id/downtime", async (req, reply) => {
    const breakdown_id = Number(req.params.id);
    const body = req.body || {};

    const log_date = String(body.log_date || "").trim();
    const hours_down = num(body.hours_down);
    const notes = body.notes ? String(body.notes).trim() : null;

    if (!Number.isInteger(breakdown_id) || breakdown_id <= 0) {
      return reply.code(400).send({ error: "Invalid breakdown id" });
    }
    if (!isDate(log_date)) return reply.code(400).send({ error: "log_date(YYYY-MM-DD) required" });
    if (Number.isNaN(hours_down) || hours_down < 0 || hours_down > 24) {
      return reply.code(400).send({ error: "hours_down must be 0..24" });
    }

    const b = getBreakdownById.get(breakdown_id);
    if (!b) return reply.code(404).send({ error: "Breakdown not found" });

    upsertDowntimeLog.run(breakdown_id, log_date, hours_down, notes);

    const updated = getBreakdownById.get(breakdown_id);
    return reply.send({
      ok: true,
      breakdown_id,
      downtime_total_hours: Number(updated?.downtime_total_hours || 0),
    });
  });

  // ---------------------------
  // Close / Reopen
  // ---------------------------
  app.post("/:id/close", async (req, reply) => {
    const id = Number(req.params.id);
    const b = getBreakdownById.get(id);
    if (!b) return reply.code(404).send({ error: "Breakdown not found" });

    const openComponentWOs = listOpenComponentWOs.all(id);
    if (openComponentWOs.length) {
      return reply.code(409).send({
        error: "Cannot close breakdown: component work orders still open",
        open_component_work_orders: openComponentWOs,
      });
    }

    closeBreakdown.run(id);
    return reply.send({ ok: true, breakdown_id: id, status: "CLOSED" });
  });

  app.post("/:id/reopen", async (req, reply) => {
    const id = Number(req.params.id);
    const b = getBreakdownById.get(id);
    if (!b) return reply.code(404).send({ error: "Breakdown not found" });

    reopenBreakdown.run(id);
    return reply.send({ ok: true, breakdown_id: id, status: "OPEN" });
  });

  /* =====================================================
     COMPONENT ROUTES
  ===================================================== */

  // List components for breakdown
  app.get("/:id/components", async (req, reply) => {
    const breakdown_id = Number(req.params.id);
    const b = getBreakdownById.get(breakdown_id);
    if (!b) return reply.code(404).send({ error: "Breakdown not found" });
    return listComponentLines.all(breakdown_id);
  });

  // Add component line (does NOT create WO)
  app.post("/:id/components", async (req, reply) => {
    const breakdown_id = Number(req.params.id);
    const component = String(req.body?.component || "").trim();
    const symptom = req.body?.symptom ? String(req.body.symptom).trim() : null;

    if (!component) return reply.code(400).send({ error: "component required" });

    const b = getBreakdownById.get(breakdown_id);
    if (!b) return reply.code(404).send({ error: "Breakdown not found" });

    const info = insertComponentLine.run(breakdown_id, component, symptom);
    return reply.code(201).send({ ok: true, breakdown_id, component_id: Number(info.lastInsertRowid) });
  });

  // Edit component line
  app.patch("/:id/components/:componentId", async (req, reply) => {
    const breakdown_id = Number(req.params.id);
    const component_id = Number(req.params.componentId);

    const existing = getComponentLine.get(component_id, breakdown_id);
    if (!existing) return reply.code(404).send({ error: "Component line not found" });

    const c = req.body?.component !== undefined ? String(req.body.component || "").trim() : null;
    const s = req.body?.symptom !== undefined ? String(req.body.symptom || "").trim() : null;
    if (c !== null && c.length === 0) return reply.code(400).send({ error: "component cannot be empty" });

    updateComponentLine.run(c, s, component_id, breakdown_id);
    return reply.send({ ok: true });
  });

  // Delete component line (only if no WO linked)
  app.delete("/:id/components/:componentId", async (req, reply) => {
    const breakdown_id = Number(req.params.id);
    const component_id = Number(req.params.componentId);

    const existing = getComponentLine.get(component_id, breakdown_id);
    if (!existing) return reply.code(404).send({ error: "Component line not found" });

    if (existing.work_order_id) {
      return reply.code(409).send({ error: "Cannot delete component line with a linked work order" });
    }

    const info = deleteComponentLine.run(component_id, breakdown_id);
    if (info.changes === 0) return reply.code(409).send({ error: "Delete not allowed" });

    return reply.send({ ok: true });
  });

  // Create WO for a specific component line (ONLY ONCE)
  app.post("/:id/components/:componentId/create-wo", async (req, reply) => {
    const breakdown_id = Number(req.params.id);
    const component_id = Number(req.params.componentId);

    const b = getBreakdownById.get(breakdown_id);
    if (!b) return reply.code(404).send({ error: "Breakdown not found" });

    const line = getComponentLine.get(component_id, breakdown_id);
    if (!line) return reply.code(404).send({ error: "Component line not found" });

    if (line.work_order_id) {
      return reply.send({
        ok: true,
        breakdown_id,
        component_id,
        work_order_id: line.work_order_id,
        already_exists: true,
      });
    }

    const tx = db.transaction(() => {
      const wo = insertWorkOrder.run(b.asset_id, breakdown_id);
      const workOrderId = Number(wo.lastInsertRowid);

      const link = linkComponentWO.run(workOrderId, component_id, breakdown_id);
      if (link.changes === 0) {
        db.prepare("DELETE FROM work_orders WHERE id = ?").run(workOrderId);
        const fresh = getComponentLine.get(component_id, breakdown_id);
        return { workOrderId: fresh.work_order_id, already_exists: true };
      }

      return { workOrderId, already_exists: false };
    });

    const r = tx();
    return reply.code(201).send({
      ok: true,
      breakdown_id,
      component_id,
      work_order_id: r.workOrderId,
      already_exists: r.already_exists,
    });
  });
}