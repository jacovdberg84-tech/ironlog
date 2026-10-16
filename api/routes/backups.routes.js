import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import { exec } from "node:child_process";
import { db, dbPathResolved } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";
import { backupBusy, backupStatus, getBackupSettings, runBackup, runRestoreTest, saveBackupSettings, startInBackground, webdavClient } from "../utils/autoBackup.js";
import { getCompanyProfile } from "../utils/companyProfile.js";
import { serverDisks, worstDisk } from "../utils/diskHealth.js";

function requestRoles(req) {
  const fromMany = String(req.headers["x-user-roles"] || "")
    .split(",")
    .map((x) => String(x || "").trim().toLowerCase())
    .filter(Boolean);
  const fromSingle = String(req.headers["x-user-role"] || "")
    .split(",")
    .map((x) => String(x || "").trim().toLowerCase())
    .filter(Boolean);
  const merged = Array.from(new Set([...fromMany, ...fromSingle]));
  return merged.length ? merged : ["admin"];
}

function requireAdmin(req, reply) {
  const roles = requestRoles(req);
  if (!roles.includes("admin") && !roles.includes("supervisor")) {
    reply.code(403).send({ ok: false, error: "admin or supervisor role required" });
    return false;
  }
  return true;
}

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

export function backupDirPath() {
  if (process.env.IRONLOG_DB_BACKUP_DIR) return path.resolve(process.env.IRONLOG_DB_BACKUP_DIR);
  return path.join(path.dirname(dbPathResolved), "backups");
}

function manualRestoreDirPath() {
  return path.join(path.dirname(dbPathResolved), "manual_restore");
}

async function listBackupFiles() {
  const dir = backupDirPath();
  await fsp.mkdir(dir, { recursive: true });
  const entries = await fsp.readdir(dir, { withFileTypes: true });
  const files = [];
  for (const e of entries) {
    if (!e.isFile()) continue;
    if (!/^ironlog_.*\.db$/i.test(e.name)) continue;
    const abs = path.join(dir, e.name);
    try {
      const st = await fsp.stat(abs);
      files.push({
        name: e.name,
        path: abs,
        bytes: Number(st.size || 0),
        modified_at: new Date(st.mtimeMs || Date.now()).toISOString(),
      });
    } catch {}
  }
  files.sort((a, b) => String(b.modified_at).localeCompare(String(a.modified_at)));
  return files;
}

function formatPowerShellRestoreCommand(filePath) {
  const dbFile = dbPathResolved.replace(/\\/g, "\\\\");
  const src = String(filePath || "").replace(/\\/g, "\\\\");
  return [
    `$db = "${dbFile}"`,
    `$src = "${src}"`,
    `Copy-Item -Path $src -Destination $db -Force`,
    `Remove-Item -Path "$($db)-wal","$($db)-shm" -ErrorAction SilentlyContinue`,
    `# Restart IRONLOG API/Desktop app after copy`,
  ].join("\n");
}

function executeRestoreEnabled() {
  return String(process.env.IRONLOG_ENABLE_RESTORE_EXECUTE || "0").trim() === "1";
}

function restartCommand() {
  return String(process.env.IRONLOG_RESTART_COMMAND || "").trim();
}

function requireAdminOnly(req, reply) {
  if (!requestRoles(req).includes("admin")) {
    reply.code(403).send({ ok: false, error: "Only an administrator can change backup settings." });
    return false;
  }
  return true;
}

const companyFolderName = () => {
  const p = getCompanyProfile(db);
  return p.short_name || p.company_name;
};

export default async function backupsRoutes(app) {
  ensureAuditTable(db);
  db.prepare(`
    CREATE TABLE IF NOT EXISTS backup_restore_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      backup_name TEXT NOT NULL,
      backup_path TEXT NOT NULL,
      action TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT 'queued',
      requested_by TEXT,
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  // Automatic and off-site backups (utils/autoBackup.js).
  app.get("/status", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    return { ok: true, running: backupBusy(), ...backupStatus(db, { dbPath: dbPathResolved, backupDir: backupDirPath() }) };
  });

  // Light check for the admin banner (no folder walk).
  app.get("/disk", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const disks = serverDisks({ dbPath: dbPathResolved, backupDir: backupDirPath() });
    return { ok: true, disk: worstDisk(disks), disks };
  });

  app.put("/settings", async (req, reply) => {
    if (!requireAdminOnly(req, reply)) return;
    const r = saveBackupSettings(db, req.body || {}, String(req.headers["x-user-name"] || ""));
    if (!r.ok) return reply.code(400).send(r);
    const { webdav_password, ...logged } = req.body || {};
    writeAudit(db, req, { module: "backups", action: "settings_update", entity_type: "backup_settings", entity_id: "1", payload: { ...logged, password_changed: Boolean(webdav_password) } });
    return r;
  });

  // Checks the storage login without saving it (or with the saved one when no password is given).
  app.post("/test-connection", async (req, reply) => {
    if (!requireAdminOnly(req, reply)) return;
    const b = req.body || {};
    try {
      let client;
      if (b.webdav_url && b.webdav_user && b.webdav_password) {
        client = webdavClient({ url: String(b.webdav_url).trim(), user: String(b.webdav_user).trim(), password: String(b.webdav_password) });
      } else {
        const s = getBackupSettings(db);
        if (!s.offsite_configured) return reply.code(400).send({ ok: false, error: "Enter the storage address, username and password first." });
        const { decryptJuneConnectorSecret } = await import("../utils/juneOutlook.js");
        const row = db.prepare(`SELECT webdav_password_enc FROM backup_settings WHERE id = 1`).get();
        client = webdavClient({ url: s.webdav_url, user: s.webdav_user, password: decryptJuneConnectorSecret(row.webdav_password_enc) || "" });
      }
      await client.test();
      return { ok: true, message: "Connected to the storage." };
    } catch (e) {
      return reply.code(424).send({ ok: false, error: `Could not connect: ${e.message || e}` });
    }
  });

  // Both start in the background and answer at once; the result shows under Recent runs.
  app.post("/run-now", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const by = String(req.headers["x-user-name"] || "");
    const started = startInBackground(async () => {
      const run = await runBackup(db, { kind: "manual", by, backupDir: backupDirPath(), companyName: companyFolderName() });
      writeAudit(db, req, { module: "backups", action: "run_now", entity_type: "backup_run", entity_id: String(run.id), payload: { status: run.status, offsite: run.offsite } });
    }, req.log);
    if (!started) return reply.code(409).send({ ok: false, error: "A backup or restore test is already running. Try again in a few minutes." });
    return reply.code(202).send({ ok: true, started: true });
  });

  app.post("/restore-test", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    const by = String(req.headers["x-user-name"] || "");
    const started = startInBackground(async () => {
      const run = await runRestoreTest(db, { by, companyName: companyFolderName() });
      writeAudit(db, req, { module: "backups", action: "restore_test", entity_type: "backup_run", entity_id: String(run.id), payload: { status: run.status } });
    }, req.log);
    if (!started) return reply.code(409).send({ ok: false, error: "A backup or restore test is already running. Try again in a few minutes." });
    return reply.code(202).send({ ok: true, started: true });
  });

  app.get("/list", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    try {
      const files = await listBackupFiles();
      return reply.send({
        ok: true,
        db_path: dbPathResolved,
        backup_dir: backupDirPath(),
        execute_restore_enabled: executeRestoreEnabled(),
        restart_command_set: Boolean(restartCommand()),
        count: files.length,
        files,
      });
    } catch (e) {
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/create", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    try {
      const dir = backupDirPath();
      await fsp.mkdir(dir, { recursive: true });
      const stamp = nowStamp();
      const tmpPath = path.join(dir, `ironlog_${stamp}.tmp.db`);
      const finalPath = path.join(dir, `ironlog_${stamp}.db`);
      await db.backup(tmpPath);
      await fsp.rename(tmpPath, finalPath);
      const st = await fsp.stat(finalPath);
      writeAudit(db, req, {
        module: "backups",
        action: "create",
        entity_type: "backup_file",
        entity_id: path.basename(finalPath),
        after: { bytes: Number(st.size || 0), path: finalPath },
      });
      return reply.send({
        ok: true,
        file: {
          name: path.basename(finalPath),
          path: finalPath,
          bytes: Number(st.size || 0),
          modified_at: new Date(st.mtimeMs || Date.now()).toISOString(),
        },
      });
    } catch (e) {
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/restore/preview", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    try {
      const backupName = String(req.body?.backup_name || "").trim();
      if (!backupName) return reply.code(400).send({ ok: false, error: "backup_name is required" });
      const backupPath = path.join(backupDirPath(), backupName);
      if (!fs.existsSync(backupPath)) return reply.code(404).send({ ok: false, error: "Backup file not found" });
      const st = await fsp.stat(backupPath);
      return reply.send({
        ok: true,
        preview: {
          backup_name: backupName,
          backup_path: backupPath,
          backup_bytes: Number(st.size || 0),
          backup_modified_at: new Date(st.mtimeMs || Date.now()).toISOString(),
          target_db_path: dbPathResolved,
          warning:
            "Restore apply is staged for safety. Copy action should be done while API/Desktop app is stopped, then restart IRONLOG.",
          confirm_text: "RESTORE",
        },
      });
    } catch (e) {
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/restore/apply", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    try {
      const backupName = String(req.body?.backup_name || "").trim();
      const confirmText = String(req.body?.confirm_text || "").trim().toUpperCase();
      const notes = String(req.body?.notes || "").trim();
      if (!backupName) return reply.code(400).send({ ok: false, error: "backup_name is required" });
      if (confirmText !== "RESTORE") {
        return reply.code(400).send({ ok: false, error: "confirm_text must be RESTORE" });
      }
      const backupPath = path.join(backupDirPath(), backupName);
      if (!fs.existsSync(backupPath)) return reply.code(404).send({ ok: false, error: "Backup file not found" });

      // Always capture a point-in-time backup before issuing restore steps.
      const dir = backupDirPath();
      await fsp.mkdir(dir, { recursive: true });
      const preStamp = nowStamp();
      const preTmp = path.join(dir, `ironlog_pre_restore_${preStamp}.tmp.db`);
      const preFinal = path.join(dir, `ironlog_pre_restore_${preStamp}.db`);
      await db.backup(preTmp);
      await fsp.rename(preTmp, preFinal);

      const restoreDir = manualRestoreDirPath();
      await fsp.mkdir(restoreDir, { recursive: true });
      const planPath = path.join(restoreDir, `restore_plan_${preStamp}.json`);
      const plan = {
        created_at: new Date().toISOString(),
        backup_name: backupName,
        backup_path: backupPath,
        target_db_path: dbPathResolved,
        pre_restore_backup: preFinal,
        notes: notes || null,
        apply_steps: [
          "Stop IRONLOG API/Desktop app",
          "Copy selected backup over active DB file",
          "Delete .db-wal and .db-shm sidecar files",
          "Start IRONLOG again and verify /health",
        ],
      };
      await fsp.writeFile(planPath, JSON.stringify(plan, null, 2), "utf8");

      const requestedBy = String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
      const inserted = db.prepare(`
        INSERT INTO backup_restore_events (backup_name, backup_path, action, status, requested_by, notes)
        VALUES (?, ?, 'restore', 'staged', ?, ?)
      `).run(backupName, backupPath, requestedBy, notes || null);

      writeAudit(db, req, {
        module: "backups",
        action: "restore_staged",
        entity_type: "backup_file",
        entity_id: backupName,
        after: { pre_restore_backup: preFinal, plan_path: planPath, notes: notes || null },
      });

      return reply.send({
        ok: true,
        event_id: Number(inserted.lastInsertRowid || 0),
        staged: {
          backup_name: backupName,
          backup_path: backupPath,
          pre_restore_backup: preFinal,
          plan_path: planPath,
        },
        next_steps: {
          message: "Run these steps while the app/API is stopped, then restart IRONLOG.",
          powershell: formatPowerShellRestoreCommand(backupPath),
        },
      });
    } catch (e) {
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/restore/execute", async (req, reply) => {
    if (!requireAdmin(req, reply)) return;
    if (!executeRestoreEnabled()) {
      return reply.code(400).send({
        ok: false,
        error: "execute restore is disabled (set IRONLOG_ENABLE_RESTORE_EXECUTE=1)",
      });
    }
    const restartCmd = restartCommand();
    if (!restartCmd) {
      return reply.code(400).send({
        ok: false,
        error: "IRONLOG_RESTART_COMMAND is not set",
      });
    }
    try {
      const backupName = String(req.body?.backup_name || "").trim();
      const confirmText = String(req.body?.confirm_text || "").trim().toUpperCase();
      const notes = String(req.body?.notes || "").trim();
      if (!backupName) return reply.code(400).send({ ok: false, error: "backup_name is required" });
      if (confirmText !== "RESTORE_NOW") {
        return reply.code(400).send({ ok: false, error: "confirm_text must be RESTORE_NOW" });
      }
      const backupPath = path.join(backupDirPath(), backupName);
      if (!fs.existsSync(backupPath)) return reply.code(404).send({ ok: false, error: "Backup file not found" });

      const dir = backupDirPath();
      await fsp.mkdir(dir, { recursive: true });
      const preStamp = nowStamp();
      const preTmp = path.join(dir, `ironlog_pre_execute_${preStamp}.tmp.db`);
      const preFinal = path.join(dir, `ironlog_pre_execute_${preStamp}.db`);
      await db.backup(preTmp);
      await fsp.rename(preTmp, preFinal);

      writeAudit(db, req, {
        module: "backups",
        action: "restore_execute_requested",
        entity_type: "backup_file",
        entity_id: backupName,
        after: { pre_restore_backup: preFinal, notes: notes || null },
      });

      try {
        db.pragma("wal_checkpoint(TRUNCATE)");
      } catch {}
      try {
        db.close();
      } catch {}

      await fsp.copyFile(backupPath, dbPathResolved);
      await Promise.allSettled([
        fsp.unlink(`${dbPathResolved}-wal`),
        fsp.unlink(`${dbPathResolved}-shm`),
      ]);

      exec(restartCmd, { windowsHide: true }, () => {});

      reply.send({
        ok: true,
        message: "Restore executed. Restart command has been triggered.",
        backup_name: backupName,
        pre_restore_backup: preFinal,
      });
      setTimeout(() => {
        process.exit(0);
      }, 250);
    } catch (e) {
      return reply.code(500).send({ ok: false, error: e.message || String(e) });
    }
  });
}
