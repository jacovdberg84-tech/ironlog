// IRONLOG/api/routes/hours.routes.js
import { db } from "../db/client.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function prevDate(dateStr) {
  const d = new Date(dateStr + "T00:00:00");
  d.setDate(d.getDate() - 1);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function numOrNull(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

export default async function hoursRoutes(app) {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS asset_input_units (
      asset_id INTEGER PRIMARY KEY,
      input_unit TEXT NOT NULL DEFAULT 'hours',
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();

  const getAssetInputUnit = db.prepare(`
    SELECT input_unit
    FROM asset_input_units
    WHERE asset_id = ?
  `);

  const upsertAssetInputUnit = db.prepare(`
    INSERT INTO asset_input_units (asset_id, input_unit, updated_at)
    VALUES (?, ?, datetime('now'))
    ON CONFLICT(asset_id) DO UPDATE SET
      input_unit = excluded.input_unit,
      updated_at = datetime('now')
  `);

  function getCarryForwardRow(assetId, workDate) {
    // Prefer exact yesterday; if missing, fall back to latest prior day.
    const y = prevDate(workDate);
    const yRow = db.prepare(`
      SELECT closing_hours, scheduled_hours, work_date AS source_date, COALESCE(NULLIF(TRIM(input_unit), ''), 'hours') AS input_unit
      FROM daily_hours
      WHERE asset_id = ? AND work_date = ?
    `).get(assetId, y);
    if (yRow && (yRow.closing_hours != null || yRow.scheduled_hours != null)) return yRow;

    const prevRow = db.prepare(`
      SELECT closing_hours, scheduled_hours, work_date AS source_date, COALESCE(NULLIF(TRIM(input_unit), ''), 'hours') AS input_unit
      FROM daily_hours
      WHERE asset_id = ?
        AND work_date < ?
      ORDER BY work_date DESC
      LIMIT 1
    `).get(assetId, workDate);
    return prevRow || null;
  }

  function hasColumn(table, col) {
    const rows = db.prepare(`PRAGMA table_info(${table})`).all();
    return rows.some((r) => String(r.name || "") === String(col));
  }
  if (!hasColumn("daily_hours", "input_unit")) {
    db.prepare(`ALTER TABLE daily_hours ADD COLUMN input_unit TEXT DEFAULT 'hours'`).run();
  }

  // -------------------------
  // GET hours rows for a date
  // -------------------------
  app.get("/:date", async (req, reply) => {
    const date = String(req.params.date || "").trim();
    if (!isDate(date)) return reply.code(400).send({ error: "date must be YYYY-MM-DD" });

    const rows = db.prepare(`
      SELECT
        dh.id,
        dh.work_date,
        dh.scheduled_hours,
        dh.opening_hours,
        dh.closing_hours,
        dh.hours_run,
        COALESCE(ai.input_unit, COALESCE(NULLIF(TRIM(dh.input_unit), ''), 'hours')) AS input_unit,
        CASE WHEN ai.input_unit IS NOT NULL THEN 1 ELSE 0 END AS input_unit_locked,
        dh.is_used,
        dh.operator,
        dh.notes,
        a.asset_code,
        a.asset_name,
        a.category,
        a.is_standby
      FROM daily_hours dh
      JOIN assets a ON a.id = dh.asset_id
      LEFT JOIN asset_input_units ai ON ai.asset_id = dh.asset_id
      WHERE dh.work_date = ?
      ORDER BY a.asset_code
    `).all(date);

    return rows.map(r => ({
      ...r,
      is_used: Boolean(r.is_used),
      is_standby: Boolean(r.is_standby)
    }));
  });

  // ------------------------------------------------
  // GET suggested defaults (yesterday carry-over)
  // ------------------------------------------------
  // /api/hours/defaults?asset_code=EXC-01&work_date=2026-02-27
  app.get("/defaults", async (req, reply) => {
    const asset_code = String(req.query?.asset_code || "").trim();
    const work_date = String(req.query?.work_date || "").trim();

    if (!asset_code || !isDate(work_date)) {
      return reply.code(400).send({ error: "asset_code and work_date(YYYY-MM-DD) required" });
    }

    const asset = db.prepare(`SELECT id FROM assets WHERE asset_code = ?`).get(asset_code);
    if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });

    const yRow = getCarryForwardRow(asset.id, work_date);

    return {
      ok: true,
      asset_code,
      work_date,
      suggested_opening_hours: yRow?.closing_hours ?? null,
      suggested_scheduled_hours: yRow?.scheduled_hours ?? null,
      suggested_opening_from_date: yRow?.source_date ?? null,
      suggested_input_unit: (getAssetInputUnit.get(asset.id)?.input_unit || yRow?.input_unit || "hours"),
      input_unit_locked: Boolean(getAssetInputUnit.get(asset.id)?.input_unit)
    };
  });

  // -----------------------------------------
  // POST upsert daily hours (Plant logic)
  // -----------------------------------------
  // Body:
  // {
  //   asset_code, work_date,
  //   is_used (true/false),
  //   scheduled_hours,
  //   opening_hours, closing_hours,
  //   hours_run,
  //   operator, notes
  // }
  app.post("/", async (req, reply) => {
    const body = req.body || {};

    const asset_code = String(body.asset_code || "").trim();
    const work_date = String(body.work_date || "").trim();

    if (!asset_code || !isDate(work_date)) {
      return reply.code(400).send({ error: "asset_code and work_date(YYYY-MM-DD) required" });
    }

    const asset = db.prepare(`
      SELECT id, is_standby, active
      FROM assets
      WHERE asset_code = ?
    `).get(asset_code);

    if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });
    if (Number(asset.active) === 0) return reply.code(409).send({ error: "asset is inactive" });

    // If the asset itself is standby in master data, it cannot be used for production
    if (Number(asset.is_standby) === 1 && body.is_used !== false) {
      // you can still log it, but it must be standby/not used
      // we force is_used=0 here for safety
    }

    const is_used = (body.is_used === false || Number(asset.is_standby) === 1) ? 0 : 1;
    const input_unit_raw = String(body.input_unit || "hours").trim().toLowerCase();
    const requested_input_unit = input_unit_raw === "km" ? "km" : "hours";
    const lockedInputUnit = String(getAssetInputUnit.get(asset.id)?.input_unit || "").trim().toLowerCase();
    const input_unit = lockedInputUnit === "km" || lockedInputUnit === "hours"
      ? lockedInputUnit
      : requested_input_unit;

    // scheduled_hours variable per day
    let scheduled_hours = numOrNull(body.scheduled_hours);
    if (scheduled_hours != null && (scheduled_hours < 0 || scheduled_hours > 24)) {
      return reply.code(400).send({ error: "scheduled_hours must be 0..24" });
    }

    let opening_hours = numOrNull(body.opening_hours);
    let closing_hours = numOrNull(body.closing_hours);
    let hours_run = numOrNull(body.hours_run);

    // Validate numeric
    const badNonNeg = (n) => n != null && n < 0;
    if (badNonNeg(opening_hours) || badNonNeg(closing_hours)) {
      return reply.code(400).send({ error: "opening_hours/closing_hours must be >= 0" });
    }
    if (hours_run != null && hours_run < 0) {
      return reply.code(400).send({ error: "hours_run must be >= 0" });
    }
    if (input_unit !== "km" && hours_run != null && hours_run > 24) {
      return reply.code(400).send({ error: "hours_run must be 0..24 for hours input_unit" });
    }

    // Auto opening from carry-forward closing if missing
    if (opening_hours == null) {
      const yRow = getCarryForwardRow(asset.id, work_date);

      if (yRow?.closing_hours != null) opening_hours = Number(yRow.closing_hours);
      if (scheduled_hours == null && yRow?.scheduled_hours != null) scheduled_hours = Number(yRow.scheduled_hours);
    }

    if (scheduled_hours == null) scheduled_hours = 0;

    // Compute missing values
    // If opening+closing -> hours_run
    if (opening_hours != null && closing_hours != null) {
      hours_run = closing_hours - opening_hours;
    }
    // If opening+hours_run -> closing
    else if (opening_hours != null && hours_run != null) {
      closing_hours = opening_hours + hours_run;
    }
    // If closing+hours_run -> opening
    else if (closing_hours != null && hours_run != null) {
      opening_hours = closing_hours - hours_run;
    }

    if (hours_run == null) hours_run = 0;

    // Safety checks
    if (hours_run < 0) {
      return reply.code(400).send({
        error: "Hourmeter mismatch: closing is lower than opening. Check opening/closing values."
      });
    }
    if (input_unit !== "km" && hours_run > 24) {
      return reply.code(400).send({
        error: "hours_run exceeds 24. Check opening/closing values."
      });
    }

    // -------------------------
    // HARD RULES (your request)
    // -------------------------

    // Standby/not-used cannot generate hours
    if (is_used === 0 && hours_run > 0) {
      return reply.code(400).send({
        error: "Standby selected. A standby asset cannot generate hours. Set asset to Production or set hours to 0."
      });
    }

    // Production must generate hours
    if (is_used === 1 && hours_run === 0) {
      return reply.code(400).send({
        error: "Production selected but no hours recorded. Enter closing hourmeter (or hours run) OR set asset to Standby."
      });
    }

    // Production should have scheduled hours (variable)
    if (is_used === 1 && scheduled_hours === 0) {
      return reply.code(400).send({
        error: "Production selected but scheduled hours is 0. Enter scheduled hours OR set asset to Standby."
      });
    }

    const operator = body.operator != null && String(body.operator).trim() !== "" ? String(body.operator).trim() : null;
    const notes = body.notes != null && String(body.notes).trim() !== "" ? String(body.notes).trim() : null;

    // Upsert
    db.prepare(`
      INSERT INTO daily_hours (
        asset_id, work_date,
        scheduled_hours, opening_hours, closing_hours,
        hours_run, input_unit, is_used, operator, notes
      )
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(asset_id, work_date) DO UPDATE SET
        scheduled_hours = excluded.scheduled_hours,
        opening_hours = excluded.opening_hours,
        closing_hours = excluded.closing_hours,
        hours_run = excluded.hours_run,
        input_unit = excluded.input_unit,
        is_used = excluded.is_used,
        operator = excluded.operator,
        notes = excluded.notes
    `).run(
      asset.id,
      work_date,
      scheduled_hours,
      opening_hours,
      closing_hours,
      hours_run,
      input_unit,
      is_used,
      operator,
      notes
    );

    // Persist selected input unit per asset; later daily rows are locked to this choice.
    upsertAssetInputUnit.run(asset.id, input_unit);

    return reply.send({
      ok: true,
      asset_code,
      work_date,
      scheduled_hours,
      opening_hours,
      closing_hours,
      hours_run,
      input_unit,
      input_unit_locked: true,
      is_used: Boolean(is_used)
    });
  });
}