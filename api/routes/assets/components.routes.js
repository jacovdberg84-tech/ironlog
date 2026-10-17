// IRONLOG/api/routes/assets/components.routes.js — a machine's major components (utils/assetComponents.js).
// Registered by routes/assets.routes.js under /api/assets.
//   GET   /component-life                       this site's fitted components at 80% of their life or more (?min_pct=)
//   GET   /:asset_code/components               fitted and removed components, hours since fitted, life used
//   POST  /:asset_code/components               fit a component { component, position?, make?, model?, serial_number?, fitted_date?, fitted_hours?, expected_life_hours?, notes? }
//   PATCH /:asset_code/components/:id           correct its details
//   POST  /:asset_code/components/:id/change    { removed_date, removed_hours?, reason, serial_number?, make?, model?, expected_life_hours?, notes? }
import { db } from "../../db/client.js";
import { getAssetCurrentHoursInfo } from "../../utils/assetMeterHours.js";
import { writeAudit } from "../../utils/audit.js";
import { getUser, requireAnyRole } from "../../utils/request.js";
import { andSite } from "../../utils/sites.js";
import { listFailureCodes } from "../../utils/failureCodes.js";
import {
  COMPONENT_WARN_PCT, addAssetComponent, componentLabels, componentView, ensureAssetComponentTables, listAssetComponents,
  readComponentFields, replaceAssetComponent, updateAssetComponent,
} from "../../utils/assetComponents.js";
import { ASSET_EDIT_ROLES } from "./register.routes.js";

const isYmd = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ""));

export default function registerComponentRoutes(app, ctx) {
  const { getAssetByCode } = ctx;
  ensureAssetComponentTables(db);
  const codes = () => listFailureCodes(db).component.map((c) => c.code);
  const meterOf = (assetId) => Number(getAssetCurrentHoursInfo(assetId).hours || 0);
  const audit = (req, action, asset, id, payload) => {
    try {
      writeAudit(db, req, { module: "assets", action, entity_type: "asset_component", entity_id: String(id), payload: { asset_code: asset.asset_code, ...payload } });
    } catch {}
  };
  const assetOr404 = (req, reply) => {
    const a = getAssetByCode.get(String(req.params.asset_code || "").trim());
    if (!a) reply.code(404).send({ ok: false, error: "Asset not found" });
    return a;
  };
  const componentOr404 = (req, reply, asset) => {
    const c = db.prepare(`SELECT * FROM asset_components WHERE id = ? AND asset_id = ?`).get(Number(req.params.id), asset.id);
    if (!c) reply.code(404).send({ ok: false, error: "Component not found on this machine." });
    return c;
  };

  app.get("/component-life", async (req) => {
    const min = Math.max(0, Math.min(1000, Number(req.query?.min_pct ?? COMPONENT_WARN_PCT) || 0));
    const rows = db.prepare(`
      SELECT c.*, a.asset_code, a.asset_name
      FROM asset_components c
      JOIN assets a ON a.id = c.asset_id
      WHERE c.status <> 'removed' AND c.expected_life_hours > 0 AND COALESCE(a.archived, 0) = 0${andSite("a", req)}
    `).all();
    const labels = componentLabels(db);
    const meters = new Map();
    const out = rows
      .map((r) => {
        if (!meters.has(r.asset_id)) meters.set(r.asset_id, meterOf(r.asset_id));
        return componentView(r, meters.get(r.asset_id), labels);
      })
      .filter((r) => r.life_used_pct != null && r.life_used_pct >= min)
      .sort((a, b) => b.life_used_pct - a.life_used_pct);
    return { ok: true, min_pct: min, rows: out };
  });

  app.get("/:asset_code/components", async (req, reply) => {
    // site: the machine code in the address is checked by the site guard (utils/siteGuard.js).
    const asset = assetOr404(req, reply);
    if (!asset) return;
    const meter = meterOf(asset.id);
    return { ok: true, asset_code: asset.asset_code, meter_hours: meter, ...listAssetComponents(db, asset.id, meter) };
  });

  app.post("/:asset_code/components", async (req, reply) => {
    // site: the machine code in the address is checked by the site guard (utils/siteGuard.js).
    if (!requireAnyRole(req, reply, ASSET_EDIT_ROLES)) return;
    const asset = assetOr404(req, reply);
    if (!asset) return;
    const r = readComponentFields(req.body || {}, { componentCodes: codes(), requireComponent: true });
    if (!r.ok) return reply.code(400).send({ ok: false, error: r.errors.join(" "), errors: r.errors });
    const id = addAssetComponent(db, asset.id, r.values, meterOf(asset.id), getUser(req));
    audit(req, "component.fit", asset, id, r.values);
    return { ok: true, id };
  });

  app.patch("/:asset_code/components/:id", async (req, reply) => {
    // site: the machine code in the address is checked by the site guard (utils/siteGuard.js).
    if (!requireAnyRole(req, reply, ASSET_EDIT_ROLES)) return;
    const asset = assetOr404(req, reply);
    if (!asset) return;
    const c = componentOr404(req, reply, asset);
    if (!c) return;
    const r = readComponentFields(req.body || {}, { componentCodes: [...codes(), c.component] });
    if (!r.ok) return reply.code(400).send({ ok: false, error: r.errors.join(" "), errors: r.errors });
    updateAssetComponent(db, c.id, r.values, getUser(req));
    audit(req, "component.update", asset, c.id, { before: c, changes: r.values });
    return { ok: true, id: c.id };
  });

  app.post("/:asset_code/components/:id/change", async (req, reply) => {
    // site: the machine code in the address is checked by the site guard (utils/siteGuard.js).
    if (!requireAnyRole(req, reply, ASSET_EDIT_ROLES)) return;
    const asset = assetOr404(req, reply);
    if (!asset) return;
    const c = componentOr404(req, reply, asset);
    if (!c) return;
    if (c.status === "removed") return reply.code(409).send({ ok: false, error: "This component was already changed." });
    const body = req.body || {};
    const removedDate = String(body.removed_date || "").trim() || new Date().toISOString().slice(0, 10);
    if (!isYmd(removedDate)) return reply.code(400).send({ ok: false, error: "Date changed must be a date like 2026-10-17." });
    if (c.fitted_date && removedDate < c.fitted_date) return reply.code(400).send({ ok: false, error: `Date changed is before it was fitted (${c.fitted_date}).` });
    const meter = meterOf(asset.id);
    const hoursRaw = String(body.removed_hours ?? "").trim();
    const removedHours = hoursRaw === "" ? meter : Number(hoursRaw);
    if (!Number.isFinite(removedHours) || removedHours < 0) return reply.code(400).send({ ok: false, error: "Machine hours at the change must be a number of hours." });
    if (c.fitted_hours != null && removedHours < Number(c.fitted_hours)) {
      return reply.code(400).send({ ok: false, error: `Machine hours at the change (${removedHours}) are below the hours when it was fitted (${c.fitted_hours}).` });
    }
    const reason = String(body.reason || "").trim().slice(0, 200);
    if (!reason) return reply.code(400).send({ ok: false, error: "Say why it was changed (e.g. failed, end of life, upgrade)." });
    const fitted = readComponentFields(
      Object.fromEntries(["serial_number", "make", "model", "expected_life_hours", "notes"].filter((k) => Object.prototype.hasOwnProperty.call(body, k)).map((k) => [k, body[k]])),
    );
    if (!fitted.ok) return reply.code(400).send({ ok: false, error: fitted.errors.join(" "), errors: fitted.errors });
    const newId = replaceAssetComponent(db, c, { removed_date: removedDate, removed_hours: removedHours, reason, fitted: fitted.values }, getUser(req));
    audit(req, "component.change", asset, c.id, { removed_date: removedDate, removed_hours: removedHours, reason, hours_run: c.fitted_hours != null ? removedHours - Number(c.fitted_hours) : null, new_id: newId, new_serial: fitted.values.serial_number || null });
    return { ok: true, removed_id: c.id, new_id: newId };
  });
}
