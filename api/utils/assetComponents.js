// IRONLOG/api/utils/assetComponents.js — a machine's major components and how much of their life is used.
//
// Each machine can list its major components (engine, transmission, final drives,
// undercarriage…), named with the component list used for failure codes so
// failures and component life line up. A component carries its serial number,
// make and model, when it was fitted (date and machine hours) and, optionally, the
// hours it is expected to last. Hours since fitted = the machine's meter now minus
// the meter when fitted. Changing a component closes the old one (removed date,
// hours and reason) and fits the new one in its place, so the history stays.
import { listFailureCodes } from "./failureCodes.js";

export const COMPONENT_WARN_PCT = 80;

export function ensureAssetComponentTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS asset_components (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      component TEXT NOT NULL,
      position TEXT,
      make TEXT,
      model TEXT,
      serial_number TEXT,
      fitted_date TEXT,
      fitted_hours REAL,
      expected_life_hours REAL,
      notes TEXT,
      status TEXT NOT NULL DEFAULT 'fitted',
      removed_date TEXT,
      removed_hours REAL,
      removed_reason TEXT,
      replaced_by_id INTEGER,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      created_by TEXT,
      updated_at TEXT,
      updated_by TEXT,
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    );
    CREATE INDEX IF NOT EXISTS idx_asset_components_asset ON asset_components(asset_id, status);
  `);
}

const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(s) && new Date(`${s}T00:00:00Z`).toISOString().slice(0, 10) === s;
const TEXT = { position: 40, make: 60, model: 60, serial_number: 60, notes: 300 };

/** The fields sent, checked: { ok, values } or { ok:false, errors }. componentCodes = the list's codes. */
export function readComponentFields(body = {}, { componentCodes = [], requireComponent = false } = {}) {
  const values = {};
  const errors = [];
  const has = (k) => Object.prototype.hasOwnProperty.call(body, k);
  if (has("component") || requireComponent) {
    const c = String(body.component ?? "").trim().toUpperCase();
    if (!c) errors.push("Choose which component it is.");
    else if (componentCodes.length && !componentCodes.includes(c)) errors.push(`${c} is not on the component list.`);
    else values.component = c;
  }
  for (const [k, max] of Object.entries(TEXT)) {
    if (!has(k)) continue;
    const v = String(body[k] ?? "").trim();
    if (v.length > max) errors.push(`${k.replace("_", " ")} can be at most ${max} characters.`);
    else values[k] = v || null;
  }
  if (has("fitted_date")) {
    const v = String(body.fitted_date ?? "").trim();
    if (v && !isDate(v)) errors.push("Fitted date must be a date like 2026-10-17.");
    else values.fitted_date = v || null;
  }
  for (const k of ["fitted_hours", "expected_life_hours"]) {
    if (!has(k)) continue;
    const raw = String(body[k] ?? "").trim();
    if (raw === "") { values[k] = null; continue; }
    const n = Number(raw);
    if (!Number.isFinite(n) || n < 0 || n > 1e6) errors.push(`${k === "fitted_hours" ? "Machine hours when fitted" : "Expected life"} must be a number of hours (0 or more).`);
    else values[k] = Math.round(n * 10) / 10;
  }
  return errors.length ? { ok: false, errors } : { ok: true, values };
}

/** A component as shown: hours since fitted and life used against the machine's meter now. */
export function componentView(row, meterHours, labels = {}) {
  const meter = Number(meterHours);
  const fitted = row.fitted_hours == null ? null : Number(row.fitted_hours);
  const end = row.status === "removed" ? (row.removed_hours == null ? null : Number(row.removed_hours)) : (Number.isFinite(meter) ? meter : null);
  const hoursSince = fitted != null && end != null ? Math.max(0, Math.round((end - fitted) * 10) / 10) : null;
  const life = row.expected_life_hours == null ? null : Number(row.expected_life_hours);
  const pct = hoursSince != null && life > 0 ? Math.round((hoursSince / life) * 100) : null;
  return {
    ...row,
    label: labels[row.component] || row.component,
    hours_since_fitted: hoursSince,
    life_used_pct: pct,
    life_left_hours: hoursSince != null && life > 0 ? Math.round((life - hoursSince) * 10) / 10 : null,
    due: pct != null && pct >= COMPONENT_WARN_PCT,
  };
}

export function componentLabels(db) {
  return Object.fromEntries(listFailureCodes(db, { all: true }).component.map((c) => [c.code, c.label_en]));
}

/** { fitted: [...], removed: [...] } for one machine. */
export function listAssetComponents(db, assetId, meterHours) {
  ensureAssetComponentTables(db);
  const labels = componentLabels(db);
  const rows = db.prepare(`SELECT * FROM asset_components WHERE asset_id = ? ORDER BY status = 'removed', component, position, id`).all(assetId);
  const all = rows.map((r) => componentView(r, meterHours, labels));
  return {
    fitted: all.filter((r) => r.status !== "removed"),
    removed: all.filter((r) => r.status === "removed").sort((a, b) => String(b.removed_date || "").localeCompare(String(a.removed_date || ""))),
  };
}

export function addAssetComponent(db, assetId, values, meterHours, by = "") {
  ensureAssetComponentTables(db);
  const v = { ...values };
  if (v.fitted_hours == null && Number.isFinite(Number(meterHours)) && !v.fitted_date) v.fitted_hours = Number(meterHours);
  if (!v.fitted_date) v.fitted_date = new Date().toISOString().slice(0, 10);
  const cols = Object.keys(v);
  const id = db.prepare(`INSERT INTO asset_components (asset_id, ${cols.join(", ")}, created_by) VALUES (?, ${cols.map(() => "?").join(", ")}, ?)`)
    .run(assetId, ...cols.map((c) => v[c]), by || null).lastInsertRowid;
  return Number(id);
}

export function updateAssetComponent(db, id, values, by = "") {
  const cols = Object.keys(values);
  if (!cols.length) return;
  db.prepare(`UPDATE asset_components SET ${cols.map((c) => `${c} = ?`).join(", ")}, updated_at = datetime('now'), updated_by = ? WHERE id = ?`)
    .run(...cols.map((c) => values[c]), by || null, id);
}

/**
 * Change a component: the old one is removed (date, machine hours, reason) and the new one is fitted
 * in its place with the same component and position. Returns the new component's id.
 */
export function replaceAssetComponent(db, oldRow, { removed_date, removed_hours, reason, fitted = {} }, by = "") {
  ensureAssetComponentTables(db);
  let newId;
  db.transaction(() => {
    newId = addAssetComponent(db, oldRow.asset_id, {
      component: oldRow.component,
      position: oldRow.position,
      make: fitted.make !== undefined ? fitted.make : oldRow.make,
      model: fitted.model !== undefined ? fitted.model : oldRow.model,
      serial_number: fitted.serial_number ?? null,
      expected_life_hours: fitted.expected_life_hours !== undefined ? fitted.expected_life_hours : oldRow.expected_life_hours,
      notes: fitted.notes ?? null,
      fitted_date: removed_date,
      fitted_hours: removed_hours,
    }, removed_hours, by);
    db.prepare(`
      UPDATE asset_components
      SET status = 'removed', removed_date = ?, removed_hours = ?, removed_reason = ?, replaced_by_id = ?, updated_at = datetime('now'), updated_by = ?
      WHERE id = ?
    `).run(removed_date, removed_hours, reason || null, newId, by || null, oldRow.id);
  })();
  return newId;
}
