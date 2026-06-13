import fs from "node:fs";
import path from "node:path";

export function getDataRoot() {
  return process.env.IRONLOG_DATA_DIR || process.cwd();
}

/** Normalize DB or URL paths to a relative storage key (uploads/...). */
export function normalizeStorageRel(relPath) {
  let rel = String(relPath || "").trim().replace(/\\/g, "/");
  if (!rel) return "";
  rel = rel.replace(/^https?:\/\/[^/]+/i, "");
  rel = rel.replace(/^\/+/, "");
  if (rel.startsWith("vehicle-ldv-checks/")) rel = `uploads/${rel}`;
  if (rel.startsWith("manager-inspections/")) rel = `uploads/${rel}`;
  if (rel.startsWith("manager-damage-reports/")) rel = `uploads/${rel}`;
  return rel;
}

export function resolveStorageAbs(relPath, dataRoot = getDataRoot()) {
  const rel = normalizeStorageRel(relPath);
  if (!rel) return "";

  const bases = [
    dataRoot,
    process.cwd(),
    process.env.IRONLOG_APP_DIR,
    process.env.IRONLOG_APP_DIR ? path.join(process.env.IRONLOG_APP_DIR, "..") : null,
  ].filter(Boolean);

  const candidates = [];
  const add = (base, subRel) => {
    const abs = path.join(base, subRel);
    if (!candidates.includes(abs)) candidates.push(abs);
  };

  for (const base of bases) {
    add(base, rel);
  }

  const baseName = path.basename(rel);
  const subDirs = ["vehicle-ldv-checks", "manager-inspections", "manager-damage-reports"];
  if (baseName && baseName !== rel) {
    for (const base of bases) {
      for (const sub of subDirs) {
        if (rel.includes(sub)) add(base, path.join("uploads", sub, baseName));
      }
    }
  }

  for (const abs of candidates) {
    if (fs.existsSync(abs)) return abs;
  }
  return candidates[0] || "";
}
