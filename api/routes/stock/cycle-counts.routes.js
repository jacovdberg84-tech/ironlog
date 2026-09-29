// IRONLOG/api/routes/stock/cycle-counts.routes.js — Cycle count sessions and quick counts.
// Registered by routes/stock.routes.js; shared helpers arrive through ctx.
import { db } from "../../db/client.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerCycleCountsRoutes(app, ctx) {
  const { getBinByCodeAtLocation, getLocationByCode, getOnHand, getPartByCode, getRole, requireRoles } = ctx;

  // Cycle count sessions
  // POST /api/stock/cycle-sessions
  app.post("/cycle-sessions", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const location_code = String(req.body?.location_code || "").trim().toUpperCase();
    const bin_code = String(req.body?.bin_code || "").trim().toUpperCase();
    const planned_date = String(req.body?.planned_date || "").trim() || new Date().toISOString().slice(0, 10);
    const notes = String(req.body?.notes || "").trim() || null;
    const location = location_code ? getLocationByCode.get(location_code) : null;
    if (location_code && !location) return reply.code(404).send({ error: `location_code not found: ${location_code}` });
    const bin = (location && bin_code) ? getBinByCodeAtLocation.get(Number(location.id), bin_code) : null;
    if (bin_code && !location_code) return reply.code(400).send({ error: "location_code is required when bin_code is provided" });
    if (location && bin_code && !bin) return reply.code(404).send({ error: `bin_code not found at location ${location_code}: ${bin_code}` });
    const ins = db.prepare(`
      INSERT INTO stock_cycle_sessions (
        location_id, bin_id, status, planned_date, counted_by, notes
      ) VALUES (?, ?, 'draft', ?, ?, ?)
    `).run(
      location ? Number(location.id) : null,
      bin ? Number(bin.id) : null,
      planned_date,
      String(req.headers["x-user-name"] || "session-user"),
      notes
    );
    return reply.send({ ok: true, session_id: Number(ins.lastInsertRowid), location_code: location ? location.location_code : null, bin_code: bin ? bin.bin_code : null });
  });

  // GET /api/stock/cycle-sessions?status=
  app.get("/cycle-sessions", async (req, reply) => {
    const status = String(req.query?.status || "").trim().toLowerCase();
    const where = [];
    const params = [];
    if (status) {
      where.push("LOWER(s.status) = ?");
      params.push(status);
    }
    const rows = db.prepare(`
      SELECT
        s.id, s.status, s.planned_date, s.counted_by, s.submitted_at, s.approved_by, s.approved_at, s.notes, s.created_at,
        l.location_code, b.bin_code,
        (SELECT COUNT(*) FROM stock_cycle_lines cl WHERE cl.session_id = s.id) AS line_count,
        (SELECT COALESCE(SUM(ABS(cl.variance_qty)), 0) FROM stock_cycle_lines cl WHERE cl.session_id = s.id) AS variance_abs
      FROM stock_cycle_sessions s
      LEFT JOIN stock_locations l ON l.id = s.location_id
      LEFT JOIN stock_bins b ON b.id = s.bin_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY s.id DESC
      LIMIT 300
    `).all(...params).map((r) => ({
      ...r,
      line_count: Number(r.line_count || 0),
      variance_abs: Number(Number(r.variance_abs || 0).toFixed(2)),
    }));
    return reply.send({ ok: true, rows });
  });

  // POST /api/stock/cycle-sessions/:id/lines/upsert
  // Body: { lines: [{ part_code, counted_qty, reason? }] }
  app.post("/cycle-sessions/:id/lines/upsert", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const session = db.prepare(`SELECT id, status, location_id, bin_id FROM stock_cycle_sessions WHERE id = ?`).get(id);
    if (!session) return reply.code(404).send({ error: "cycle session not found" });
    if (!["draft", "counting"].includes(String(session.status || "").toLowerCase())) {
      return reply.code(409).send({ error: `cannot edit lines when status is ${session.status}` });
    }
    const lines = Array.isArray(req.body?.lines) ? req.body.lines : [];
    if (!lines.length) return reply.code(400).send({ error: "lines array is required" });
    const upsert = db.prepare(`
      INSERT INTO stock_cycle_lines (
        session_id, part_id, system_qty, counted_qty, variance_qty, status, reason
      ) VALUES (?, ?, ?, ?, ?, 'draft', ?)
      ON CONFLICT(session_id, part_id) DO UPDATE SET
        system_qty = excluded.system_qty,
        counted_qty = excluded.counted_qty,
        variance_qty = excluded.variance_qty,
        reason = excluded.reason
    `);
    const tx = db.transaction(() => {
      for (const line of lines) {
        const part_code = String(line?.part_code || "").trim();
        const counted_qty = Number(line?.counted_qty ?? NaN);
        if (!part_code || !Number.isFinite(counted_qty) || counted_qty < 0) continue;
        const part = getPartByCode.get(part_code);
        if (!part) continue;
        const system_qty = Number(
          db.prepare(`
            SELECT COALESCE(SUM(quantity), 0) AS q
            FROM stock_movements
            WHERE part_id = ?
              AND COALESCE(location_id, 0) = COALESCE(?, 0)
              AND COALESCE(bin_id, 0) = COALESCE(?, 0)
          `).get(Number(part.id), session.location_id == null ? null : Number(session.location_id), session.bin_id == null ? null : Number(session.bin_id))?.q || 0
        );
        const variance = Number((counted_qty - system_qty).toFixed(2));
        upsert.run(
          id,
          Number(part.id),
          Number(system_qty.toFixed(2)),
          Number(counted_qty.toFixed(2)),
          variance,
          line?.reason ? String(line.reason).trim() : null
        );
      }
      db.prepare(`UPDATE stock_cycle_sessions SET status = 'counting' WHERE id = ? AND status = 'draft'`).run(id);
    });
    tx();
    return reply.send({ ok: true, session_id: id });
  });

  app.post("/cycle-sessions/:id/submit", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const session = db.prepare(`SELECT id, status FROM stock_cycle_sessions WHERE id = ?`).get(id);
    if (!session) return reply.code(404).send({ error: "cycle session not found" });
    if (!["draft", "counting"].includes(String(session.status || "").toLowerCase())) {
      return reply.code(409).send({ error: `cannot submit when status is ${session.status}` });
    }
    db.prepare(`
      UPDATE stock_cycle_sessions
      SET status = 'submitted', submitted_at = datetime('now')
      WHERE id = ?
    `).run(id);
    return reply.send({ ok: true, session_id: id, status: "submitted" });
  });

  app.post("/cycle-sessions/:id/approve", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor"])) return;
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ error: "invalid id" });
    const session = db.prepare(`SELECT id, status, location_id, bin_id FROM stock_cycle_sessions WHERE id = ?`).get(id);
    if (!session) return reply.code(404).send({ error: "cycle session not found" });
    if (!["submitted", "approved"].includes(String(session.status || "").toLowerCase())) {
      return reply.code(409).send({ error: `cannot approve when status is ${session.status}` });
    }
    const lines = db.prepare(`
      SELECT id, part_id, variance_qty
      FROM stock_cycle_lines
      WHERE session_id = ?
    `).all(id);
    const tx = db.transaction(() => {
      for (const line of lines) {
        const variance = Number(line.variance_qty || 0);
        if (!Number.isFinite(variance) || variance === 0) continue;
        db.prepare(`
          INSERT INTO stock_movements (part_id, quantity, movement_type, reference, location_id, bin_id)
          VALUES (?, ?, 'adjust', ?, ?, ?)
        `).run(
          Number(line.part_id),
          variance,
          `cycle_session:${id}:line:${Number(line.id)}`,
          session.location_id == null ? null : Number(session.location_id),
          session.bin_id == null ? null : Number(session.bin_id)
        );
        db.prepare(`
          UPDATE stock_cycle_lines
          SET status = 'approved', approved_at = datetime('now')
          WHERE id = ?
        `).run(Number(line.id));
      }
      db.prepare(`
        UPDATE stock_cycle_sessions
        SET status = 'approved', approved_by = ?, approved_at = datetime('now')
        WHERE id = ?
      `).run(String(req.headers["x-user-name"] || "session-user"), id);
    });
    tx();
    return reply.send({ ok: true, session_id: id, status: "approved" });
  });

  // Submit cycle count as stock adjustment approval request
  // POST /api/stock/cycle-count
  // Body: { part_code, counted_qty, reason? }
  app.post("/cycle-count", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const part_code = String(req.body?.part_code || "").trim();
    const counted_qty = Number(req.body?.counted_qty ?? NaN);
    const reason = String(req.body?.reason || "").trim() || "cycle_count";

    if (!part_code) return reply.code(400).send({ error: "part_code is required" });
    if (!Number.isFinite(counted_qty) || counted_qty < 0) {
      return reply.code(400).send({ error: "counted_qty must be a valid number >= 0" });
    }

    const part = getPartByCode.get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });
    const on_hand = Number(getOnHand.get(part.id)?.on_hand || 0);
    const delta = Number((counted_qty - on_hand).toFixed(2));
    if (delta === 0) {
      return reply.send({
        ok: true,
        no_change: true,
        message: "Counted quantity matches current on-hand. No adjustment required.",
        part_code,
        on_hand,
        counted_qty: Number(counted_qty.toFixed(2)),
      });
    }

    const reqRole = getRole(req);
    const reqUser = String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
    const reference = `cycle_count:${reason}`;
    const approvalPayload = JSON.stringify({
      part_code,
      quantity: delta,
      reference,
    });

    const ins = db.prepare(`
      INSERT INTO approval_requests (
        module, action, entity_type, entity_id, status, payload_json, requested_by, requested_role
      )
      VALUES ('stock', 'adjust_movement', 'part', ?, 'pending', ?, ?, ?)
    `).run(part_code, approvalPayload, reqUser, reqRole);
    const request_id = Number(ins.lastInsertRowid);

    writeAudit(db, req, {
      module: "stock",
      action: "cycle_count_request",
      entity_type: "part",
      entity_id: part_code,
      payload: {
        request_id,
        on_hand_before: on_hand,
        counted_qty: Number(counted_qty.toFixed(2)),
        adjustment_qty: delta,
        reason,
      },
    });

    return reply.send({
      ok: true,
      pending_approval: true,
      request_id,
      part_code,
      on_hand_before: on_hand,
      counted_qty: Number(counted_qty.toFixed(2)),
      adjustment_qty: delta,
      reference,
      message: "Cycle count submitted for approval",
    });
  });
}
