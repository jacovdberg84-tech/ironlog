/**
 * Default safety equipment checklist templates (seeded into DB on first run).
 * Site admins can override items via PUT /api/safety/templates/:key
 */

export const DEFAULT_SAFETY_TEMPLATES = {
  fire_extinguisher: {
    template_key: "fire_extinguisher",
    title: "Fire extinguisher inspection",
    items: [
      { key: "ext_present", label: "Extinguisher present and accessible" },
      { key: "ext_seal_pin", label: "Safety pin and tamper seal intact" },
      { key: "ext_pressure", label: "Pressure gauge in serviceable range" },
      { key: "ext_label", label: "Inspection tag and expiry date valid" },
      { key: "ext_mounting", label: "Mounting bracket and signage acceptable" },
    ],
  },
};

/** Slug for template_key from a display name or explicit key. */
export function normalizeTemplateKey(raw) {
  const s = String(raw || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 48);
  return s || "category";
}

export function isValidTemplateKey(key) {
  return /^[a-z][a-z0-9_]{1,47}$/.test(String(key || ""));
}

export function normalizeTemplateItems(items) {
  const out = [];
  const seen = new Set();
  for (const raw of Array.isArray(items) ? items : []) {
    const key = String(raw?.key || "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9_]/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_|_$/g, "");
    const label = String(raw?.label || "").trim();
    if (!key || !label || seen.has(key)) continue;
    seen.add(key);
    out.push({ key, label });
  }
  return out;
}

export function parseTemplateItemsJson(json) {
  try {
    const parsed = typeof json === "string" ? JSON.parse(json) : json;
    return normalizeTemplateItems(parsed);
  } catch {
    return [];
  }
}

export function buildChecklistFromTemplate(templateItems, submission) {
  const byKey = {};
  for (const row of Array.isArray(submission) ? submission : []) {
    const k = String(row?.key || "").trim();
    if (!k) continue;
    byKey[k] = row;
  }
  return normalizeTemplateItems(templateItems).map((it) => {
    const hit = byKey[it.key] || {};
    let ok = null;
    const raw = hit.ok;
    if (raw === true || raw === "ok" || raw === "true" || raw === 1) ok = true;
    else if (raw === false || raw === "fail" || raw === "false" || raw === 0) ok = false;
    const note = String(hit.note || hit.notes || "").trim() || null;
    return { key: it.key, label: it.label, ok, note };
  });
}

export function checklistStatus(checklist) {
  const rows = Array.isArray(checklist) ? checklist : [];
  if (!rows.length) return "pending";
  if (rows.some((r) => r.ok === false)) return "fail";
  if (rows.every((r) => r.ok === true)) return "pass";
  return "attention";
}
