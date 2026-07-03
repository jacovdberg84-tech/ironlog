/**
 * Idempotent setup for BMH / Polar / BMP hired equipment.
 * Run: node api/scripts/setup-hired-equipment.mjs
 */
import { db } from "../db/client.js";
import { buildKnownHireAssetPatches, inferHireContractorLabel } from "../utils/hiredEquipment.js";
import { ensurePlantHireSchema } from "../utils/plantHire.js";

ensurePlantHireSchema(db);

const findAsset = db.prepare(`
  SELECT id, asset_code, asset_name, category, hire_billing_mode, utilization_mode, active, archived
  FROM assets
  WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?))
  LIMIT 1
`);

const updateAsset = db.prepare(`
  UPDATE assets
  SET
    category = COALESCE(?, category),
    hire_billing_mode = COALESCE(?, hire_billing_mode),
    utilization_mode = COALESCE(?, utilization_mode),
    active = COALESCE(?, active),
    archived = COALESCE(?, archived),
    archive_reason = CASE WHEN ? = 1 THEN COALESCE(archive_reason, 'Hired — FAMS fuel only (skip daily hours)') ELSE archive_reason END
  WHERE id = ?
`);

const patches = buildKnownHireAssetPatches();
let updated = 0;
let missing = [];

for (const patch of patches) {
  const row = findAsset.get(patch.asset_code);
  if (!row) {
    missing.push(patch.asset_code);
    continue;
  }
  const archive = patch.archived != null ? Number(patch.archived) : null;
  updateAsset.run(
    patch.category || null,
    patch.hire_billing_mode || null,
    patch.utilization_mode || null,
    patch.active != null ? Number(patch.active) : null,
    archive,
    archive,
    row.id,
  );
  console.log(
    `  ${patch.asset_code} → ${patch.category} | hire=${patch.hire_billing_mode} | mode=${patch.utilization_mode} | was archived=${row.archived}`,
  );
  updated += 1;
}

console.log(`\nUpdated ${updated} asset(s).`);
if (missing.length) {
  console.log(`Missing asset codes (create in Assets first): ${missing.join(", ")}`);
}

const register = db.prepare(`
  SELECT asset_code, category, hire_billing_mode, utilization_mode, active, archived
  FROM assets
  WHERE UPPER(asset_code) IN (${patches.map(() => "?").join(", ")})
  ORDER BY asset_code
`).all(...patches.map((p) => p.asset_code));

console.log("\nRegister snapshot:");
for (const r of register) {
  const contractor = inferHireContractorLabel(r.asset_code, r.category);
  console.log(
    `  ${r.asset_code} | ${contractor || "?"} | ${r.utilization_mode || "-"} | hire=${r.hire_billing_mode || "-"} | active=${r.active} archived=${r.archived}`,
  );
}
