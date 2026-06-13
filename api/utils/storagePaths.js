import fs from "node:fs";
import path from "node:path";

export function getDataRoot() {
  return process.env.IRONLOG_DATA_DIR || process.cwd();
}

export function resolveStorageAbs(relPath, dataRoot = getDataRoot()) {
  const rel = String(relPath || "").replace(/\\/g, "/").replace(/^\/+/, "");
  if (!rel) return "";
  const fromData = path.join(dataRoot, rel);
  if (fs.existsSync(fromData)) return fromData;
  return path.join(process.cwd(), rel);
}
