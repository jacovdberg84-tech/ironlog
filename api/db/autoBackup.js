import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import { db, dbPathResolved } from "./client.js";

function nowStamp() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  const ss = String(d.getSeconds()).padStart(2, "0");
  return `${yyyy}${mm}${dd}_${hh}${mi}${ss}`;
}

async function statSignature(filePath) {
  try {
    const s = await fsp.stat(filePath);
    return `${s.mtimeMs}:${s.size}`;
  } catch {
    return "missing";
  }
}

async function buildDbChangeSignature() {
  const wal = `${dbPathResolved}-wal`;
  const shm = `${dbPathResolved}-shm`;
  const [baseSig, walSig, shmSig] = await Promise.all([
    statSignature(dbPathResolved),
    statSignature(wal),
    statSignature(shm),
  ]);
  return `${baseSig}|${walSig}|${shmSig}`;
}

async function pruneBackups(backupDir, keepCount) {
  const entries = await fsp.readdir(backupDir, { withFileTypes: true });
  const files = entries
    .filter((e) => e.isFile() && /^ironlog_\d{8}_\d{6}\.db$/.test(e.name))
    .map((e) => e.name)
    .sort()
    .reverse();

  const toDelete = files.slice(Math.max(keepCount, 1));
  await Promise.allSettled(toDelete.map((name) => fsp.unlink(path.join(backupDir, name))));
}

async function createBackup(backupDir, logger) {
  const stamp = nowStamp();
  const tmpPath = path.join(backupDir, `ironlog_${stamp}.tmp.db`);
  const finalPath = path.join(backupDir, `ironlog_${stamp}.db`);
  await db.backup(tmpPath);
  await fsp.rename(tmpPath, finalPath);
  logger.info({ file: finalPath }, "[db-backup] backup created");
}

export function startDbAutoBackup(logger = console) {
  const backupDir = process.env.IRONLOG_DB_BACKUP_DIR
    ? path.resolve(process.env.IRONLOG_DB_BACKUP_DIR)
    : path.join(path.dirname(dbPathResolved), "backups");
  const pollMs = Number(process.env.IRONLOG_DB_BACKUP_POLL_MS || 5000);
  const debounceMs = Number(process.env.IRONLOG_DB_BACKUP_DEBOUNCE_MS || 10000);
  const keepCount = Number(process.env.IRONLOG_DB_BACKUP_KEEP || 336);

  fs.mkdirSync(backupDir, { recursive: true });

  let lastSig = "";
  let lastBackupAt = 0;
  let busy = false;
  let stopped = false;

  const tick = async () => {
    if (stopped || busy) return;
    busy = true;
    try {
      const sig = await buildDbChangeSignature();
      const changed = sig !== lastSig;
      const now = Date.now();
      if (changed && now - lastBackupAt >= debounceMs) {
        await createBackup(backupDir, logger);
        await pruneBackups(backupDir, keepCount);
        lastBackupAt = now;
      }
      lastSig = sig;
    } catch (err) {
      logger.error({ err }, "[db-backup] backup tick failed");
    } finally {
      busy = false;
    }
  };

  tick();
  const timer = setInterval(tick, Math.max(1000, pollMs));

  logger.info(
    {
      dbPath: dbPathResolved,
      backupDir,
      pollMs,
      debounceMs,
      keepCount,
    },
    "[db-backup] auto-backup enabled",
  );

  return () => {
    stopped = true;
    clearInterval(timer);
  };
}
