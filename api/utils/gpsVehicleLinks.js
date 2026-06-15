// GPS registration / IMEI → IRONLOG asset_code mapping (Cartrack + Unitech).

import { db } from "../db/client.js";

export function normalizeGpsRegistration(reg) {
  return String(reg || "").trim().toUpperCase().replace(/\s+/g, "");
}

export function ensureGpsVehicleLinkTables() {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS gps_vehicle_links (
      registration TEXT PRIMARY KEY,
      asset_code TEXT NOT NULL,
      gps_source TEXT,
      vehicle_name TEXT,
      notes TEXT,
      updated_by TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_gps_vehicle_links_asset ON gps_vehicle_links(asset_code)`).run();
}

export function resolveGpsAssetCode(registration, vehicleName = "") {
  ensureGpsVehicleLinkTables();
  const reg = normalizeGpsRegistration(registration);
  if (!reg) return null;

  const linked = db.prepare(`
    SELECT asset_code FROM gps_vehicle_links WHERE registration = ? LIMIT 1
  `).get(reg);
  if (linked?.asset_code) return String(linked.asset_code).trim().toUpperCase();

  const byCode = db.prepare(`
    SELECT asset_code FROM assets
    WHERE UPPER(REPLACE(asset_code, ' ', '')) = ?
      AND archived = 0
    LIMIT 1
  `).get(reg);
  if (byCode?.asset_code) return String(byCode.asset_code);

  const hay = `${registration} ${vehicleName}`.trim();
  if (hay) {
    const assets = db.prepare(`SELECT asset_code, asset_name FROM assets WHERE archived = 0`).all();
    const low = hay.toLowerCase();
    for (const a of assets) {
      const code = String(a.asset_code || "");
      const name = String(a.asset_name || "");
      if (code && low.includes(code.toLowerCase())) return code;
      if (name && low.includes(name.toLowerCase())) return code;
    }
  }

  return reg;
}

export function listGpsVehicleLinks() {
  ensureGpsVehicleLinkTables();
  return db.prepare(`
    SELECT registration, asset_code, gps_source, vehicle_name, notes, updated_by, updated_at
    FROM gps_vehicle_links
    ORDER BY asset_code ASC, registration ASC
  `).all();
}

export function upsertGpsVehicleLink({
  registration,
  asset_code,
  gps_source,
  vehicle_name,
  notes,
  updated_by,
}) {
  ensureGpsVehicleLinkTables();
  const reg = normalizeGpsRegistration(registration);
  const code = String(asset_code || "").trim().toUpperCase();
  if (!reg) throw new Error("registration is required");
  if (!code) throw new Error("asset_code is required");
  const asset = db.prepare(`SELECT id FROM assets WHERE UPPER(asset_code) = ? AND archived = 0 LIMIT 1`).get(code);
  if (!asset) throw new Error(`No active asset with code ${code}`);

  db.prepare(`
    INSERT INTO gps_vehicle_links (registration, asset_code, gps_source, vehicle_name, notes, updated_by, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
    ON CONFLICT(registration) DO UPDATE SET
      asset_code = excluded.asset_code,
      gps_source = excluded.gps_source,
      vehicle_name = excluded.vehicle_name,
      notes = excluded.notes,
      updated_by = excluded.updated_by,
      updated_at = datetime('now')
  `).run(
    reg,
    code,
    String(gps_source || "any").trim().toLowerCase() || "any",
    String(vehicle_name || "").trim() || null,
    String(notes || "").trim() || null,
    String(updated_by || "admin").trim()
  );
  return { registration: reg, asset_code: code };
}

export function deleteGpsVehicleLink(registration) {
  ensureGpsVehicleLinkTables();
  const reg = normalizeGpsRegistration(registration);
  const result = db.prepare(`DELETE FROM gps_vehicle_links WHERE registration = ?`).run(reg);
  return { deleted: result.changes > 0, registration: reg };
}

function assetCodeExists(code) {
  if (!code) return false;
  return Boolean(db.prepare(`
    SELECT id FROM assets WHERE UPPER(asset_code) = ? AND archived = 0 LIMIT 1
  `).get(String(code).trim().toUpperCase()));
}

function isLikelyUnmapped(assetCode, registration) {
  const reg = normalizeGpsRegistration(registration);
  const code = String(assetCode || "").trim().toUpperCase();
  if (!reg) return false;
  if (!code || code === reg) return true;
  return !assetCodeExists(code);
}

export function listGpsFleetMappingSuggestions() {
  ensureGpsVehicleLinkTables();
  const rows = [];
  const seen = new Set();

  const pushRow = (row) => {
    const reg = normalizeGpsRegistration(row.registration);
    if (!reg || seen.has(`${row.gps_source}:${reg}`)) return;
    if (!isLikelyUnmapped(row.asset_code, reg)) return;
    seen.add(`${row.gps_source}:${reg}`);
    rows.push({
      registration: reg,
      current_label: row.asset_code || reg,
      vehicle_name: row.vehicle_name || null,
      gps_source: row.gps_source || "unknown",
      linked_asset_code: null,
    });
  };

  try {
    for (const row of db.prepare(`
      SELECT registration, asset_code, vehicle_name, 'cartrack' AS gps_source
      FROM cartrack_vehicle_snapshots
    `).all()) {
      pushRow(row);
    }
  } catch {
    /* table may not exist yet */
  }

  try {
    for (const row of db.prepare(`
      SELECT registration, asset_code, vehicle_name, 'unitech' AS gps_source
      FROM unitech_vehicle_snapshots
    `).all()) {
      pushRow(row);
    }
  } catch {
    /* table may not exist yet */
  }

  const links = new Map(listGpsVehicleLinks().map((l) => [l.registration, l.asset_code]));
  return rows
    .map((r) => ({ ...r, linked_asset_code: links.get(r.registration) || null }))
    .filter((r) => !r.linked_asset_code)
    .sort((a, b) => String(a.registration).localeCompare(String(b.registration)));
}

export function applyGpsVehicleLinksToSnapshots() {
  ensureGpsVehicleLinkTables();
  let updated = 0;

  const applyTable = (table) => {
    const rows = db.prepare(`
      SELECT registration, vehicle_name FROM ${table}
    `).all();
    const upd = db.prepare(`UPDATE ${table} SET asset_code = ? WHERE registration = ?`);
    for (const row of rows) {
      const next = resolveGpsAssetCode(row.registration, row.vehicle_name);
      if (!next) continue;
      upd.run(next, row.registration);
      updated += 1;
    }
  };

  try {
    applyTable("cartrack_vehicle_snapshots");
  } catch {
    /* optional */
  }
  try {
    applyTable("unitech_vehicle_snapshots");
  } catch {
    /* optional */
  }

  return { updated };
}
