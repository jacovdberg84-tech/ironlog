/** Shared hired / contractor equipment rules (BMH, Polar, BMP km units). */

/** BMH plant — hour-meter benchmarking & hourly hire. */
export const BMH_HIRE_ASSET_CODES = [
  "E017",
  "E018",
  "E025",
  "BR3",
  "BW10",
  "BW11",
];

/** Registration prefixes treated as km/L (FAMS KMHour = odometer). */
export const HIRE_KM_CODE_PREFIXES = ["BMP", "PTT", "LDV"];

/** Polar hired trucks PTT10–PTT26 (km). */
export function isPolarHireCode(assetCode) {
  const m = /^PTT(\d+)$/i.exec(String(assetCode || "").trim());
  if (!m) return null;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return false;
  return n >= 10 && n <= 26;
}

export function isBmhHireCode(assetCode) {
  return BMH_HIRE_ASSET_CODES.includes(String(assetCode || "").trim().toUpperCase());
}

export function isKmHireCode(assetCode) {
  const code = String(assetCode || "").trim().toUpperCase();
  if (!code) return false;
  if (HIRE_KM_CODE_PREFIXES.some((p) => code.startsWith(p))) return true;
  if (isPolarHireCode(code)) return true;
  return false;
}

export function isHireCategory(category) {
  const c = String(category || "").toLowerCase();
  return c.includes("contractor hire") || c.includes("contractor") || c.includes("hire");
}

export function isOperationalHireAsset(row) {
  if (!row) return false;
  const active = Number(row.active ?? 1) === 1;
  const hasHireMode = Boolean(String(row.hire_billing_mode || "").trim());
  const hireCat = isHireCategory(row.category);
  const knownCode = isBmhHireCode(row.asset_code) || isKmHireCode(row.asset_code);
  return active && (hasHireMode || hireCat || knownCode);
}

/** Archived for daily-hours skip, but still active for fuel / hire reporting. */
export function isArchivedHireOperational(row) {
  return Number(row?.archived || 0) === 1 && isOperationalHireAsset(row);
}

/**
 * SQL predicate: include normal fleet OR archived-but-active hire/contractor units.
 */
export function sqlIncludeArchivedHireAssets(a = "a") {
  return `(
    COALESCE(${a}.archived, 0) = 0
    OR (
      COALESCE(${a}.active, 1) = 1
      AND (
        NULLIF(TRIM(COALESCE(${a}.hire_billing_mode, '')), '') IS NOT NULL
        OR LOWER(COALESCE(${a}.category, '')) LIKE '%contractor%'
        OR LOWER(COALESCE(${a}.category, '')) LIKE '%hire%'
        OR UPPER(COALESCE(${a}.asset_code, '')) IN (${BMH_HIRE_ASSET_CODES.map((c) => `'${c}'`).join(", ")})
        OR UPPER(COALESCE(${a}.asset_code, '')) LIKE 'BMP%'
        OR UPPER(COALESCE(${a}.asset_code, '')) LIKE 'PTT%'
      )
    )
  )`;
}

export function inferHireContractorLabel(assetCode, category) {
  const code = String(assetCode || "").trim().toUpperCase();
  if (isBmhHireCode(code) || String(category || "").toUpperCase().includes("BMH")) return "BMH";
  if (isPolarHireCode(code) || String(category || "").toUpperCase().includes("POLAR")) return "Polar";
  if (code.startsWith("BMP")) return "BMP";
  return "";
}

/** Idempotent asset patches for known hired fleet (run via setup script). */
export function buildKnownHireAssetPatches() {
  const patches = [];
  for (const code of BMH_HIRE_ASSET_CODES) {
    patches.push({
      asset_code: code,
      category: "Contractor Hire (BMH)",
      hire_billing_mode: "hourly",
      utilization_mode: "hours",
      active: 1,
    });
  }
  for (let n = 10; n <= 26; n += 1) {
    const code = `PTT${n}`;
    patches.push({
      asset_code: code,
      category: "Contractor Hire (Polar)",
      hire_billing_mode: "hourly",
      utilization_mode: "km",
      active: 1,
    });
  }
  patches.push({
    asset_code: "BMP006",
    category: "Contractor Hire (BMP)",
    hire_billing_mode: "hourly",
    utilization_mode: "km",
    active: 1,
  });
  return patches;
}
