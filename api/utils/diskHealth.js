// IRONLOG/api/utils/diskHealth.js — how full the server's disk is, and what IronLog keeps on it.
//
// The live disk once filled to 100% with on-server database copies before
// anyone noticed. Administrators now see a banner at the top of every screen
// when space runs low (GET /api/admin/backups/disk), and the Off-site backups
// card shows a gauge and the biggest IronLog folders.
import fs from "node:fs";
import path from "node:path";

// Warn early enough to act (days, not minutes): below 10% or 5 GB free; urgent below 5% or 2 GB.
export const DISK_LIMITS = { warn_pct: 10, warn_gb: 5, critical_pct: 5, critical_gb: 2 };

const gb = (bytes) => Number((bytes / 1e9).toFixed(1));

/** Disk level from statfs numbers: { name, free_gb, total_gb, used_pct, level, message }. */
export function diskLevel({ bavail, blocks, bsize }, name = "Server disk") {
  const free = bavail * bsize;
  const total = blocks * bsize;
  const freePct = total > 0 ? (free / total) * 100 : 0;
  const out = { name, free_gb: gb(free), total_gb: gb(total), used_pct: total > 0 ? Math.round(100 - freePct) : null, level: "ok", message: "" };
  const backups = name !== "Server disk";
  if (free < DISK_LIMITS.critical_gb * 1e9 || freePct < DISK_LIMITS.critical_pct) {
    out.level = "critical";
    out.message = backups
      ? `${name} almost full: only ${out.free_gb} GB free (${out.used_pct}% used). New database copies will stop until space is freed; IronLog keeps working.`
      : `${name} almost full: only ${out.free_gb} GB free (${out.used_pct}% used). IronLog may stop saving work. Free up space now (Admin → Off-site backups shows what is using it).`;
  } else if (free < DISK_LIMITS.warn_gb * 1e9 || freePct < DISK_LIMITS.warn_pct) {
    out.level = "warn";
    out.message = `${name} is getting full: ${out.free_gb} GB free (${out.used_pct}% used). Plan a clean-up soon.`;
  }
  return out;
}

/** Disk holding `dir`, or null when it cannot be read. */
export function diskHealth(dir = ".", statfs = fs.statfsSync, name = "Server disk") {
  try {
    return diskLevel(statfs(path.resolve(dir)), name);
  } catch {
    return null;
  }
}

/** The server disk, plus the backup volume when database copies are kept on a separate disk. */
export function serverDisks({ dbPath = "", backupDir = "" } = {}, { statfs = fs.statfsSync, stat = fs.statSync } = {}) {
  const dbDir = path.dirname(dbPath || ".");
  const disks = [diskHealth(dbDir, statfs)].filter(Boolean);
  try {
    if (backupDir && stat(backupDir).dev !== stat(dbDir).dev) {
      const vol = diskHealth(backupDir, statfs, "Backup volume");
      if (vol) disks.push(vol);
    }
  } catch {}
  return disks;
}

const RANK = { ok: 0, warn: 1, critical: 2 };
/** The disk most in need of attention (the server disk when all are fine). */
export function worstDisk(disks = []) {
  return disks.reduce((worst, d) => (!worst || RANK[d.level] > RANK[worst.level] ? d : worst), null);
}

/** Total size and file count under a folder; stops counting after `maxFiles` (then `partial` is true). */
export function folderSize(dir, maxFiles = 50000) {
  let bytes = 0;
  let files = 0;
  let partial = false;
  const stack = [dir];
  while (stack.length) {
    const cur = stack.pop();
    let entries;
    try {
      entries = fs.readdirSync(cur, { withFileTypes: true });
    } catch {
      continue;
    }
    for (const e of entries) {
      const p = path.join(cur, e.name);
      if (e.isDirectory()) stack.push(p);
      else if (e.isFile()) {
        if (files >= maxFiles) { partial = true; break; }
        try { bytes += fs.statSync(p).size; files += 1; } catch {}
      }
    }
    if (partial) break;
  }
  return { bytes, files, partial };
}

let usageCache = { at: 0, key: "", rows: null };

/** What IronLog keeps on the disk, largest first. Cached for 10 minutes (folder walks are slow). */
export function ironlogUsage({ dbPath = "", backupDir = "", dataDir = "" } = {}, now = Date.now()) {
  const key = `${dbPath}|${backupDir}|${dataDir}`;
  if (usageCache.rows && usageCache.key === key && now - usageCache.at < 10 * 60 * 1000) return usageCache.rows;
  const rows = [];
  try { rows.push({ label: "Database", bytes: fs.statSync(dbPath).size, files: 1 }); } catch {}
  if (backupDir) rows.push({ label: "Database copies on this server", ...folderSize(backupDir) });
  if (dataDir) {
    for (const sub of ["uploads", "workshop-files", "reports-cache"]) {
      const dir = path.join(dataDir, sub);
      if (fs.existsSync(dir)) rows.push({ label: { uploads: "Photos and uploads", "workshop-files": "Workshop library files", "reports-cache": "Saved reports" }[sub], ...folderSize(dir) });
    }
  }
  rows.sort((a, b) => b.bytes - a.bytes);
  usageCache = { at: now, key, rows };
  return rows;
}
