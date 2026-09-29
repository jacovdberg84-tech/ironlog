// IRONLOG/api/utils/serviceTemplateBuilder.js — proposes service templates from real history.
//
// For every machine and service interval in the maintenance plans, the builder
// looks at the service work orders already done for that plan and what stores
// issued to them. Items issued on at least half of those services become the
// template, at their typical (median) quantity. Machines with no history fall
// back to a service kit whose stock code names the machine and interval (for
// example TEICH-A300AM-500HRS). Labour is not captured, so the template uses a
// standard time for the interval at the default labour rate. Nothing is saved
// until the proposals are applied.
import { ensureStockCategorySchema, stockCategorySql } from "./stockCategory.js";
import { ensureServiceTemplateSchema, resolveServiceTemplate } from "./serviceTemplates.js";
import { meterUnitForAsset } from "./serviceSchedule.js";

/** Standard labour hours for a service interval (hours meters) or LDV service (km). */
export function standardLabourHours(interval, meterUnit = "hours") {
  const n = Number(interval) || 0;
  if (meterUnit === "km") return 2;
  if (n <= 250) return 2;
  if (n <= 500) return 4;
  if (n <= 1000) return 6;
  if (n <= 2000) return 8;
  return 12;
}

function median(values) {
  const v = [...values].sort((a, b) => a - b);
  if (!v.length) return 0;
  const mid = Math.floor(v.length / 2);
  return v.length % 2 ? v[mid] : (v[mid - 1] + v[mid]) / 2;
}

function itemTypeFor(part) {
  const name = String(part.part_name || "").toLowerCase();
  if (part.stock_category === "oil") return /grease/.test(name) ? "grease" : /coolant|anti-?freeze/.test(name) ? "coolant" : "oil";
  if (/\bkits?\b/.test(name) || /\d{2,5}\s?h(r|rs|ours?)?\b/.test(name)) return "service_kit";
  if (/\bfilters?\b/.test(name)) return "filter";
  return "part";
}

function roundQty(qty, type) {
  // Oils go by the litre; parts by whole units.
  if (["oil", "grease", "coolant"].includes(type)) return Math.max(1, Math.round(qty));
  return Math.max(1, Math.ceil(qty - 1e-9));
}

function labourRate(db) {
  try {
    const row = db.prepare(`SELECT value FROM cost_settings WHERE key = 'labor_cost_per_hour_default' LIMIT 1`).get();
    const v = Number(row?.value);
    return Number.isFinite(v) && v > 0 ? v : 0;
  } catch {
    return 0;
  }
}

const norm = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");

/** "500HR", "500 hrs", "500 hour" in a kit's code or name — but not 1500HR or 350500. */
function hasIntervalMark(kit, iv) {
  return new RegExp(`(^|[^0-9])${iv}\\s?H`, "i").test(`${kit.part_code} ${kit.part_name}`);
}

/** Service kit in stores whose code or name mentions this machine and interval, e.g. TEICH-A300AM-500HRS. */
export function matchKit(kits, assetCode, interval) {
  const code = norm(assetCode);
  const iv = String(Math.round(Number(interval) || 0));
  if (!code || iv === "0") return null;
  const hits = kits.filter((k) => norm(`${k.part_code} ${k.part_name}`).includes(code) && hasIntervalMark(k, iv));
  return hits.length === 1 ? hits[0] : null;
}

/**
 * Fallback: a model kit (e.g. "CAT 350 500hr service kit" for a CAT 350) when no kit
 * names the machine. Kits coded to a different machine (TEICH-E501AM ...) are skipped.
 */
export function matchModelKit(kits, assetCode, assetName, interval) {
  const iv = String(Math.round(Number(interval) || 0));
  if (iv === "0") return null;
  const words = String(assetName || "").toUpperCase().split(/[^A-Z0-9]+/).filter(Boolean);
  const brand = words.find((w) => /^[A-Z]{2,}$/.test(w));
  const models = [...new Set(words.filter((w) => /\d/.test(w)).flatMap((w) => [w, w.replace(/[A-Z]+$/, "")]).filter((w) => w.length >= 3))];
  if (!brand || !models.length) return null;
  const own = norm(assetCode);
  const hits = kits.filter((k) => {
    const text = norm(`${k.part_code} ${k.part_name}`);
    const otherMachine = (String(k.part_code).toUpperCase().match(/[A-Z]{1,3}\d{2,3}[A-Z]{2}\b/g) || []).some((m) => norm(m) !== own);
    return !otherMachine && text.includes(brand) && models.some((m) => text.includes(m)) && hasIntervalMark(k, iv);
  });
  return hits.length === 1 ? hits[0] : null;
}

/**
 * Proposals for every active machine + service interval.
 * Each proposal: key, asset, interval, source, history_services, items[], labour, estimate, existing template.
 */
export function buildServiceTemplateProposals(db) {
  ensureStockCategorySchema(db);
  ensureServiceTemplateSchema(db);
  const rate = labourRate(db);
  const plans = db.prepare(`
    SELECT mp.id AS plan_id, mp.asset_id, mp.service_name, mp.interval_hours,
      a.asset_code, a.asset_name, a.category
    FROM maintenance_plans mp
    JOIN assets a ON a.id = mp.asset_id
    WHERE mp.active = 1 AND a.active = 1 AND COALESCE(a.archived, 0) = 0
      AND COALESCE(mp.interval_hours, 0) > 0
    ORDER BY a.asset_code, mp.interval_hours
  `).all();

  const parts = new Map(db.prepare(`
    SELECT id, part_code, part_name, COALESCE(unit_cost, 0) AS unit_cost, ${stockCategorySql("parts")} AS stock_category
    FROM parts
  `).all().map((p) => [p.id, p]));
  const kits = [...parts.values()].filter((p) => itemTypeFor(p) === "service_kit");

  // Net quantity issued to each service work order, per part.
  const woIssues = db.prepare(`
    SELECT w.id AS work_order_id, sm.part_id, -SUM(sm.quantity) AS qty
    FROM work_orders w
    JOIN stock_movements sm ON sm.reference = ('work_order:' || w.id)
    WHERE LOWER(COALESCE(w.source, '')) = 'service' AND w.reference_id = ?
    GROUP BY w.id, sm.part_id
    HAVING qty > 0
  `);

  const groups = new Map();
  for (const p of plans) {
    const key = `${p.asset_id}:${Number(p.interval_hours)}`;
    if (!groups.has(key)) groups.set(key, { ...p, plan_ids: [] });
    groups.get(key).plan_ids.push(p.plan_id);
  }

  const proposals = [];
  for (const [key, g] of groups) {
    const meterUnit = meterUnitForAsset(g.asset_code);
    const perWo = new Map();
    for (const planId of g.plan_ids) {
      for (const row of woIssues.all(planId)) {
        if (!perWo.has(row.work_order_id)) perWo.set(row.work_order_id, new Map());
        perWo.get(row.work_order_id).set(row.part_id, Number(row.qty));
      }
    }
    const services = perWo.size;
    const items = [];
    let source = "none";
    if (services) {
      source = "history";
      const byPart = new Map();
      for (const issues of perWo.values()) for (const [partId, qty] of issues) {
        if (!byPart.has(partId)) byPart.set(partId, []);
        byPart.get(partId).push(qty);
      }
      for (const [partId, qtys] of byPart) {
        if (qtys.length < Math.ceil(services / 2)) continue; // one-off extras are not part of the standard service
        const part = parts.get(partId);
        if (!part) continue;
        const type = itemTypeFor(part);
        items.push({ part, type, qty: roundQty(median(qtys), type), seen: qtys.length });
      }
    }
    if (!items.length) {
      const kit = matchKit(kits, g.asset_code, g.interval_hours);
      const modelKit = kit ? null : matchModelKit(kits, g.asset_code, g.asset_name, g.interval_hours);
      if (kit || modelKit) {
        source = kit ? "kit_code" : "kit_model";
        items.push({ part: kit || modelKit, type: "service_kit", qty: 1, seen: 0 });
      }
    }
    const order = { service_kit: 0, filter: 1, part: 2, oil: 3, grease: 4, coolant: 5 };
    items.sort((a, b) => (order[a.type] ?? 9) - (order[b.type] ?? 9) || String(a.part.part_code).localeCompare(String(b.part.part_code)));

    const labourHours = standardLabourHours(g.interval_hours, meterUnit);
    const partsCost = items.reduce((s, it) => s + it.qty * Number(it.part.unit_cost || 0), 0);
    const missingPrices = items.filter((it) => !(Number(it.part.unit_cost) > 0)).map((it) => it.part.part_code);
    const existing = resolveServiceTemplate(db, { assetId: g.asset_id, intervalHours: g.interval_hours, meterUnit });
    const unit = meterUnit === "km" ? "km" : "h";
    proposals.push({
      key,
      asset_id: g.asset_id,
      asset_code: g.asset_code,
      asset_name: g.asset_name,
      category: g.category,
      service_name: g.service_name,
      interval: Number(g.interval_hours),
      meter_unit: meterUnit,
      template_key: `AUTO-${norm(g.asset_code)}-${Math.round(Number(g.interval_hours))}`,
      name: `${g.asset_code} ${Number(g.interval_hours).toLocaleString("en-US")} ${unit} service`,
      source,
      history_services: services,
      items: items.map((it) => ({
        stock_item_id: it.part.id,
        part_code: it.part.part_code,
        description: it.part.part_name || it.part.part_code,
        item_type: it.type,
        quantity_required: it.qty,
        unit_of_measure: ["oil", "coolant"].includes(it.type) ? "L" : it.type === "grease" ? "kg" : "ea",
        unit_cost: Number(it.part.unit_cost || 0),
        seen_on_services: it.seen,
      })),
      labour_hours: labourHours,
      labour_rate: rate,
      estimate: {
        parts_cost: Math.round(partsCost * 100) / 100,
        labour_cost: Math.round(labourHours * rate * 100) / 100,
        total: Math.round((partsCost + labourHours * rate) * 100) / 100,
        missing_prices: missingPrices,
      },
      existing: existing.status === "matched"
        ? { id: existing.template.id, name: existing.template.name, asset_specific: existing.template.tier === 3 }
        : existing.status === "ambiguous" ? { ambiguous: true } : null,
    });
  }
  // Same-model machines use the same oils. Where a proposal has no oils yet,
  // borrow them from a sibling of the same model and interval built from history.
  const oilTypes = new Set(["oil", "grease", "coolant"]);
  for (const p of proposals) {
    if (!p.items.length || p.items.some((it) => oilTypes.has(it.item_type))) continue;
    const donor = proposals.find((d) => d !== p && d.source === "history" && d.interval === p.interval
      && norm(d.asset_name) && norm(d.asset_name) === norm(p.asset_name) && d.items.some((it) => oilTypes.has(it.item_type)));
    if (!donor) continue;
    const oils = donor.items.filter((it) => oilTypes.has(it.item_type)).map((it) => ({ ...it, seen_on_services: 0, borrowed_from: donor.asset_code }));
    p.items.push(...oils);
    p.oils_from = donor.asset_code;
    const oilCost = oils.reduce((sum, it) => sum + it.quantity_required * Number(it.unit_cost || 0), 0);
    p.estimate.parts_cost = Math.round((p.estimate.parts_cost + oilCost) * 100) / 100;
    p.estimate.total = Math.round((p.estimate.parts_cost + p.estimate.labour_cost) * 100) / 100;
  }
  return proposals;
}

/**
 * Creates templates for the chosen proposals, assigned to their machine.
 * A machine's previous machine-specific template for the same interval is
 * retired (kept for history). Returns created template ids.
 */
export function applyServiceTemplateProposals(db, proposals, keys) {
  const wanted = new Set((Array.isArray(keys) ? keys : []).map(String));
  const chosen = proposals.filter((p) => wanted.has(p.key) && p.items.length);
  const insertTemplate = db.prepare(`
    INSERT INTO service_templates (
      template_key, name, description, manufacturer, model, asset_category,
      service_interval_hours, meter_unit, estimated_duration_hours,
      default_labour_hours, default_labour_rate, active, revision_number, supersedes_template_id
    ) VALUES (?, ?, ?, NULL, NULL, ?, ?, ?, ?, ?, ?, 1, ?, ?)
  `);
  const insertItem = db.prepare(`
    INSERT INTO service_template_items (
      service_template_id, stock_item_id, item_type, description, quantity_required,
      unit_of_measure, required, allow_substitute, notes, sort_order
    ) VALUES (?, ?, ?, ?, ?, ?, 1, 0, ?, ?)
  `);
  const latestRevision = db.prepare(`SELECT id, revision_number FROM service_templates WHERE template_key = ? ORDER BY revision_number DESC LIMIT 1`);
  const retireAssetAssignments = db.prepare(`
    UPDATE asset_service_template_assignments SET active = 0, updated_at = datetime('now')
    WHERE asset_id = ? AND active = 1 AND service_template_id IN (
      SELECT id FROM service_templates WHERE ABS(service_interval_hours - ?) < 0.0001 AND LOWER(meter_unit) = ?
    )
  `);
  const retireTemplate = db.prepare(`UPDATE service_templates SET active = 0, updated_at = datetime('now') WHERE id = ?`);
  const assign = db.prepare(`
    INSERT INTO asset_service_template_assignments (asset_id, service_template_id, priority, active)
    VALUES (?, ?, 100, 1)
  `);
  const created = [];
  db.transaction(() => {
    for (const p of chosen) {
      const prev = latestRevision.get(p.template_key);
      const revision = prev ? Number(prev.revision_number) + 1 : 1;
      const note = p.source === "history"
        ? `Built from ${p.history_services} past service${p.history_services === 1 ? "" : "s"}; standard labour ${p.labour_hours} h.`
        : `Service kit matched by ${p.source === "kit_model" ? "machine model" : "stock code"}; ${p.oils_from ? `oils as used on ${p.oils_from}` : "add oils"} and check quantities. Standard labour ${p.labour_hours} h.`;
      const id = Number(insertTemplate.run(
        p.template_key, p.name, note, p.category || null, p.interval, p.meter_unit,
        p.labour_hours, p.labour_hours, p.labour_rate, revision, prev ? prev.id : null,
      ).lastInsertRowid);
      p.items.forEach((it, index) => insertItem.run(
        id, it.stock_item_id, it.item_type, it.description, it.quantity_required, it.unit_of_measure,
        it.borrowed_from ? `Quantity as used on ${it.borrowed_from}` : it.seen_on_services ? `Issued on ${it.seen_on_services} of ${p.history_services} services` : null, index,
      ));
      retireAssetAssignments.run(p.asset_id, p.interval, p.meter_unit);
      if (prev) retireTemplate.run(prev.id);
      assign.run(p.asset_id, id);
      created.push({ key: p.key, id, template_key: p.template_key, revision });
    }
  })();
  return created;
}
