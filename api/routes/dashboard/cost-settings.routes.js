// IRONLOG/api/routes/dashboard/cost-settings.routes.js — Cost settings, asset rates and part costs.
// Registered by routes/dashboard.routes.js; shared helpers arrive through ctx.
import { db } from "../../db/client.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerCostSettingsRoutes(app, ctx) {
  const { requireRoles } = ctx;

  // GET /api/dashboard/cost/settings
  app.get("/cost/settings", async () => {
    const rows = db.prepare(`
      SELECT key, value
      FROM cost_settings
      ORDER BY key ASC
    `).all();
    const settings = {};
    for (const r of rows) settings[String(r.key)] = Number(r.value || 0);
    return { ok: true, settings };
  });

  // POST /api/dashboard/cost/settings
  // Body: { fuel_cost_per_liter_default?, lube_cost_per_qty_default?, labor_cost_per_hour_default?, downtime_cost_per_hour_default? }
  app.post("/cost/settings", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const body = req.body || {};
    const allowed = [
      "fuel_cost_per_liter_default",
      "lube_cost_per_qty_default",
      "labor_cost_per_hour_default",
      "downtime_cost_per_hour_default",
    ];
    const updates = [];
    for (const k of allowed) {
      if (body[k] == null || body[k] === "") continue;
      const v = Number(body[k]);
      if (!Number.isFinite(v) || v < 0) {
        return reply.code(400).send({ error: `${k} must be a valid number >= 0` });
      }
      updates.push({ key: k, value: Number(v.toFixed(4)) });
    }
    if (!updates.length) return reply.code(400).send({ error: "provide at least one setting value" });

    const upsert = db.prepare(`
      INSERT INTO cost_settings (key, value, updated_at)
      VALUES (?, ?, datetime('now'))
      ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = datetime('now')
    `);
    const tx = db.transaction((rowsToSave) => {
      for (const r of rowsToSave) upsert.run(r.key, r.value);
    });
    tx(updates);

    writeAudit(db, req, {
      module: "cost",
      action: "settings_update",
      entity_type: "cost_settings",
      payload: updates,
    });

    return { ok: true, updates };
  });

  // POST /api/dashboard/cost/asset-rates
  // Body: { asset_code, fuel_cost_per_liter?, downtime_cost_per_hour?, utilization_mode?, km_per_hour_factor? }
  app.post("/cost/asset-rates", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const body = req.body || {};
    const asset_code = String(body.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ error: "asset_code is required" });

    const asset = db.prepare(`
      SELECT id, asset_code, asset_name
      FROM assets
      WHERE asset_code = ?
    `).get(asset_code);
    if (!asset) return reply.code(404).send({ error: `asset_code not found: ${asset_code}` });

    const fuelCostRaw = body.fuel_cost_per_liter;
    const downCostRaw = body.downtime_cost_per_hour;
    const utilizationModeRaw = body.utilization_mode;
    const kmPerHourRaw = body.km_per_hour_factor;
    const updates = [];
    const params = [];
    if (fuelCostRaw != null && String(fuelCostRaw).trim() !== "") {
      const v = Number(fuelCostRaw);
      if (!Number.isFinite(v) || v < 0) return reply.code(400).send({ error: "fuel_cost_per_liter must be >= 0" });
      updates.push("fuel_cost_per_liter = ?");
      params.push(Number(v.toFixed(4)));
    }
    if (downCostRaw != null && String(downCostRaw).trim() !== "") {
      const v = Number(downCostRaw);
      if (!Number.isFinite(v) || v < 0) return reply.code(400).send({ error: "downtime_cost_per_hour must be >= 0" });
      updates.push("downtime_cost_per_hour = ?");
      params.push(Number(v.toFixed(4)));
    }
    if (utilizationModeRaw != null && String(utilizationModeRaw).trim() !== "") {
      const mode = String(utilizationModeRaw).trim().toLowerCase();
      if (!["hours", "km"].includes(mode)) {
        return reply.code(400).send({ error: "utilization_mode must be 'hours' or 'km'" });
      }
      updates.push("utilization_mode = ?");
      params.push(mode);
    }
    if (kmPerHourRaw != null && String(kmPerHourRaw).trim() !== "") {
      const v = Number(kmPerHourRaw);
      if (!Number.isFinite(v) || v <= 0) return reply.code(400).send({ error: "km_per_hour_factor must be > 0" });
      updates.push("km_per_hour_factor = ?");
      params.push(Number(v.toFixed(4)));
    }
    if (!updates.length) return reply.code(400).send({ error: "provide at least one asset rate value" });

    db.prepare(`
      UPDATE assets
      SET ${updates.join(", ")}
      WHERE id = ?
    `).run(...params, asset.id);

    const row = db.prepare(`
      SELECT asset_code, asset_name, fuel_cost_per_liter, downtime_cost_per_hour, utilization_mode, km_per_hour_factor
      FROM assets
      WHERE id = ?
    `).get(asset.id);

    writeAudit(db, req, {
      module: "cost",
      action: "asset_rate_update",
      entity_type: "asset",
      entity_id: asset_code,
      payload: row,
    });

    return { ok: true, asset: row };
  });

  // POST /api/dashboard/cost/part-cost
  // Body: { part_code, unit_cost }
  app.post("/cost/part-cost", async (req, reply) => {
    if (!requireRoles(req, reply, ["admin", "supervisor", "stores"])) return;
    const body = req.body || {};
    const part_code = String(body.part_code || "").trim();
    const unit_cost = Number(body.unit_cost);
    if (!part_code) return reply.code(400).send({ error: "part_code is required" });
    if (!Number.isFinite(unit_cost) || unit_cost < 0) return reply.code(400).send({ error: "unit_cost must be >= 0" });

    const part = db.prepare(`
      SELECT id, part_code, part_name
      FROM parts
      WHERE part_code = ?
    `).get(part_code);
    if (!part) return reply.code(404).send({ error: `part_code not found: ${part_code}` });

    db.prepare(`
      UPDATE parts
      SET unit_cost = ?
      WHERE id = ?
    `).run(Number(unit_cost.toFixed(4)), part.id);

    writeAudit(db, req, {
      module: "cost",
      action: "part_cost_update",
      entity_type: "part",
      entity_id: part_code,
      payload: { unit_cost: Number(unit_cost.toFixed(4)) },
    });

    return { ok: true, part_code: part.part_code, part_name: part.part_name, unit_cost: Number(unit_cost.toFixed(4)) };
  });
}
