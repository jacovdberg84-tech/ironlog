// IRONLOG/api/utils/autoBackup.js — automatic backups, off-site copies and restore tests.
//
// Every night (site time, from the company profile) IronLog:
//   1. copies the database with SQLite's online backup and checks the copy opens
//      and passes an integrity check;
//   2. compresses it and sends it off the server to the company's storage
//      (WebDAV: a Hetzner Storage Box, or any WebDAV server), together with any
//      new or changed uploaded files (photos, documents, logos);
//   3. keeps the last few copies on the server and 30 days off-site.
// On the 1st of each month it downloads the newest off-site copy and proves it
// opens: a restore test. Every run is recorded in backup_runs, failures alert
// the administrators, and Admin → Backups shows it all.
//
// The storage login is stored encrypted with the server-only connector key.
import fs from "node:fs";
import fsp from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import zlib from "node:zlib";
import { pipeline } from "node:stream/promises";
import { Readable } from "node:stream";
import Database from "better-sqlite3";
import { decryptJuneConnectorSecret, encryptJuneConnectorSecret } from "./juneOutlook.js";
import { getDataRoot } from "./storagePaths.js";

export const FILE_FOLDERS = ["uploads", "workshop-files", "branding"];
const AUTO_PREFIX = "ironlog_auto_";
const MAX_FILES_PER_RUN = 3000;
const REQUEST_TIMEOUT_MS = 10 * 60 * 1000;

export function ensureBackupTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS backup_settings (
      id INTEGER PRIMARY KEY CHECK (id = 1),
      enabled INTEGER NOT NULL DEFAULT 1,
      run_hour INTEGER NOT NULL DEFAULT 2,
      keep_local INTEGER NOT NULL DEFAULT 7,
      keep_remote_days INTEGER NOT NULL DEFAULT 30,
      include_files INTEGER NOT NULL DEFAULT 1,
      webdav_url TEXT,
      webdav_user TEXT,
      webdav_password_enc TEXT,
      remote_folder TEXT,
      updated_by TEXT,
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    INSERT OR IGNORE INTO backup_settings (id) VALUES (1);
    CREATE TABLE IF NOT EXISTS backup_runs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      kind TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT 'running',
      started_at TEXT NOT NULL DEFAULT (datetime('now')),
      finished_at TEXT,
      db_file TEXT,
      db_bytes INTEGER,
      check_result TEXT,
      offsite TEXT,
      files_uploaded INTEGER NOT NULL DEFAULT 0,
      files_pending INTEGER NOT NULL DEFAULT 0,
      error TEXT,
      triggered_by TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_backup_runs_kind ON backup_runs(kind, id);
    CREATE TABLE IF NOT EXISTS backup_remote_files (
      rel_path TEXT PRIMARY KEY,
      bytes INTEGER NOT NULL,
      mtime_ms INTEGER NOT NULL,
      uploaded_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
  `);
}

const clampInt = (v, lo, hi, d) => {
  const n = Number(v);
  return Number.isInteger(n) && n >= lo && n <= hi ? n : d;
};

/** Settings without the password; `has_password` says whether one is stored. */
export function getBackupSettings(db) {
  ensureBackupTables(db);
  const r = db.prepare(`SELECT * FROM backup_settings WHERE id = 1`).get();
  return {
    enabled: Number(r.enabled) === 1,
    run_hour: r.run_hour,
    keep_local: r.keep_local,
    keep_remote_days: r.keep_remote_days,
    include_files: Number(r.include_files) === 1,
    webdav_url: r.webdav_url || "",
    webdav_user: r.webdav_user || "",
    has_password: Boolean(r.webdav_password_enc),
    remote_folder: r.remote_folder || "",
    offsite_configured: Boolean(r.webdav_url && r.webdav_user && r.webdav_password_enc),
    updated_by: r.updated_by,
    updated_at: r.updated_at,
  };
}

export function saveBackupSettings(db, body = {}, actor = "") {
  ensureBackupTables(db);
  const cur = db.prepare(`SELECT * FROM backup_settings WHERE id = 1`).get();
  const url = body.webdav_url !== undefined ? String(body.webdav_url || "").trim().replace(/\/+$/, "") : cur.webdav_url || "";
  if (url && !/^https:\/\/[^\s/]+/i.test(url)) return { ok: false, error: "The storage address must start with https:// (for a Storage Box: https://uXXXXXX.your-storagebox.de)." };
  const folder = body.remote_folder !== undefined
    ? String(body.remote_folder || "").trim().replace(/^\/+|\/+$/g, "").replace(/[^A-Za-z0-9/_-]/g, "-").slice(0, 120)
    : cur.remote_folder || "";
  let passwordEnc = cur.webdav_password_enc;
  if (body.webdav_password !== undefined && String(body.webdav_password) !== "") {
    try {
      passwordEnc = encryptJuneConnectorSecret(String(body.webdav_password));
    } catch {
      return { ok: false, error: "The server's encryption key is missing, so the password cannot be stored safely." };
    }
  }
  if (body.clear_password) passwordEnc = null;
  db.prepare(`
    UPDATE backup_settings SET enabled = ?, run_hour = ?, keep_local = ?, keep_remote_days = ?, include_files = ?,
      webdav_url = ?, webdav_user = ?, webdav_password_enc = ?, remote_folder = ?, updated_by = ?, updated_at = datetime('now')
    WHERE id = 1
  `).run(
    body.enabled !== undefined ? (body.enabled === false || String(body.enabled) === "0" ? 0 : 1) : cur.enabled,
    clampInt(body.run_hour, 0, 23, cur.run_hour),
    clampInt(body.keep_local, 1, 60, cur.keep_local),
    clampInt(body.keep_remote_days, 7, 3650, cur.keep_remote_days),
    body.include_files !== undefined ? (body.include_files === false || String(body.include_files) === "0" ? 0 : 1) : cur.include_files,
    url || null,
    body.webdav_user !== undefined ? String(body.webdav_user || "").trim() || null : cur.webdav_user,
    passwordEnc,
    folder || null,
    actor || null,
  );
  return { ok: true, settings: getBackupSettings(db) };
}

// ------------------------------------------------------------------ WebDAV

/** A small WebDAV client over fetch: enough to create folders, put, list, get and delete. */
export function webdavClient({ url, user, password, fetchImpl = globalThis.fetch }) {
  const base = String(url).replace(/\/+$/, "");
  const auth = `Basic ${Buffer.from(`${user}:${password}`).toString("base64")}`;
  const href = (p) => `${base}/${String(p).split("/").filter(Boolean).map(encodeURIComponent).join("/")}`;
  const call = async (method, p, { body, headers = {}, ok = [200, 201, 204, 207] } = {}) => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
    try {
      const init = { method, headers: { Authorization: auth, ...headers }, body, signal: controller.signal };
      if (body && typeof body.pipe === "function") init.duplex = "half";
      const res = await fetchImpl(href(p), init);
      if (!ok.includes(res.status)) {
        const hint = res.status === 401 ? " (check the username and password)" : res.status === 404 ? " (folder not found)" : "";
        throw new Error(`Storage ${method} ${p || "/"} failed: HTTP ${res.status}${hint}`);
      }
      return res;
    } catch (e) {
      if (e?.name === "AbortError") throw new Error(`Storage ${method} ${p || "/"} timed out`);
      if (e instanceof TypeError) throw new Error(`Could not reach ${base}. Check the address, and that WebDAV is switched on for the storage.`);
      throw e;
    } finally {
      clearTimeout(timer);
    }
  };
  const mkdirs = async (dir) => {
    let at = "";
    for (const part of String(dir).split("/").filter(Boolean)) {
      at = `${at}/${part}`;
      // 405 = it already exists.
      await call("MKCOL", at, { ok: [200, 201, 204, 405] });
    }
  };
  const list = async (dir) => {
    const res = await call("PROPFIND", `${dir}/`, { headers: { Depth: "1", "Content-Type": "application/xml" }, ok: [207, 404] });
    if (res.status === 404) return [];
    const xml = await res.text();
    const names = [];
    for (const m of xml.matchAll(/<(?:[a-zA-Z]+:)?href>([^<]+)<\/(?:[a-zA-Z]+:)?href>/g)) {
      const name = decodeURIComponent(m[1].replace(/\/+$/, "").split("/").pop() || "");
      if (name) names.push(name);
    }
    const self = String(dir).split("/").filter(Boolean).pop();
    return names.filter((n) => n !== self);
  };
  return {
    test: async () => { await call("PROPFIND", "/", { headers: { Depth: "0" }, ok: [207] }); },
    mkdirs,
    list,
    put: (p, body, size) => call("PUT", p, { body, headers: size != null ? { "Content-Length": String(size) } : {} }),
    get: (p) => call("GET", p, { ok: [200] }),
    remove: (p) => call("DELETE", p, { ok: [200, 204, 404] }),
  };
}

function storageFor(db, fetchImpl) {
  const r = db.prepare(`SELECT webdav_url, webdav_user, webdav_password_enc FROM backup_settings WHERE id = 1`).get();
  if (!r?.webdav_url || !r.webdav_user || !r.webdav_password_enc) return null;
  const password = decryptJuneConnectorSecret(r.webdav_password_enc);
  if (!password) throw new Error("The stored storage password could not be read. Enter it again under Admin → Backups.");
  return webdavClient({ url: r.webdav_url, user: r.webdav_user, password, fetchImpl });
}

/** The company's folder on the storage, e.g. ironlog/aml. */
export function remoteRoot(settings, companyName = "") {
  if (settings.remote_folder) return settings.remote_folder;
  const slug = String(companyName || "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || os.hostname().toLowerCase();
  return `ironlog/${slug}`;
}

// ------------------------------------------------------------------ backup steps

function stamp(now = new Date()) {
  return now.toISOString().replace(/[-:]/g, "").replace("T", "_").slice(0, 15);
}

/** Opens a database copy read-only and checks it: integrity, tables, key row counts. */
export function checkDatabaseCopy(file) {
  const copy = new Database(file, { readonly: true, fileMustExist: true });
  try {
    const integrity = copy.prepare(`PRAGMA quick_check`).pluck().get();
    const tables = copy.prepare(`SELECT COUNT(*) FROM sqlite_master WHERE type = 'table'`).pluck().get();
    const count = (t) => {
      try { return copy.prepare(`SELECT COUNT(*) FROM ${t}`).pluck().get(); } catch { return null; }
    };
    return { ok: integrity === "ok", integrity, tables, assets: count("assets"), work_orders: count("work_orders"), users: count("users") };
  } finally {
    copy.close();
  }
}

async function gzipFile(src, dest) {
  await pipeline(fs.createReadStream(src), zlib.createGzip({ level: 6 }), fs.createWriteStream(dest));
  return (await fsp.stat(dest)).size;
}

async function pruneLocal(dir, keep) {
  const names = (await fsp.readdir(dir)).filter((n) => n.startsWith(AUTO_PREFIX) && n.endsWith(".db")).sort().reverse();
  for (const n of names.slice(keep)) await fsp.rm(path.join(dir, n), { force: true });
  return Math.max(0, names.length - keep);
}

/** Off-site copies older than `days`, never the newest 7. Names carry their date: ironlog_auto_YYYYMMDD_HHMMSS.db.gz */
export function remoteToPrune(names, days, now = new Date()) {
  const cutoff = stamp(new Date(now.getTime() - days * 86400000));
  const copies = names.filter((n) => n.startsWith(AUTO_PREFIX) && n.endsWith(".db.gz")).sort().reverse();
  return copies.slice(7).filter((n) => n.slice(AUTO_PREFIX.length, AUTO_PREFIX.length + 15) < cutoff);
}

/** Uploaded files that are new or changed since they were last sent off-site. */
export function filesToSync(db, dataRoot = getDataRoot()) {
  const known = new Map(db.prepare(`SELECT rel_path, bytes, mtime_ms FROM backup_remote_files`).all().map((r) => [r.rel_path, r]));
  const out = [];
  const walk = (rel) => {
    let entries;
    try { entries = fs.readdirSync(path.join(dataRoot, rel), { withFileTypes: true }); } catch { return; }
    for (const e of entries) {
      const childRel = `${rel}/${e.name}`;
      if (e.isDirectory()) walk(childRel);
      else if (e.isFile()) {
        const st = fs.statSync(path.join(dataRoot, childRel));
        const k = known.get(childRel);
        if (!k || k.bytes !== st.size || k.mtime_ms !== Math.round(st.mtimeMs)) out.push({ rel: childRel, bytes: st.size, mtime_ms: Math.round(st.mtimeMs) });
      }
    }
  };
  for (const f of FILE_FOLDERS) walk(f);
  return out.sort((a, b) => a.rel.localeCompare(b.rel));
}

function startRun(db, kind, by) {
  return Number(db.prepare(`INSERT INTO backup_runs (kind, triggered_by) VALUES (?, ?)`).run(kind, by || null).lastInsertRowid);
}

function finishRun(db, id, fields) {
  const cols = Object.keys(fields);
  db.prepare(`UPDATE backup_runs SET ${cols.map((c) => `${c} = ?`).join(", ")}, finished_at = datetime('now') WHERE id = ?`)
    .run(...cols.map((c) => (fields[c] != null && typeof fields[c] === "object" ? JSON.stringify(fields[c]) : fields[c])), id);
}

/**
 * One backup: local copy + check, then off-site copy and file sync when storage
 * is set up. Returns the run row. Never throws.
 */
export async function runBackup(db, { kind = "nightly", by = "system", backupDir, companyName = "", dataRoot = getDataRoot(), fetchImpl, now = new Date() } = {}) {
  ensureBackupTables(db);
  const settings = getBackupSettings(db);
  const id = startRun(db, kind, by);
  const result = { status: "ok", offsite: "not set up", files_uploaded: 0, files_pending: 0 };
  let gz = null;
  let sending = false;
  try {
    await fsp.mkdir(backupDir, { recursive: true });
    const name = `${AUTO_PREFIX}${stamp(now)}.db`;
    const tmp = path.join(backupDir, `${name}.tmp`);
    const file = path.join(backupDir, name);
    await db.backup(tmp);
    await fsp.rename(tmp, file);
    result.db_file = name;
    result.db_bytes = (await fsp.stat(file)).size;
    const check = checkDatabaseCopy(file);
    result.check_result = check;
    if (!check.ok) throw new Error(`The backup copy failed its integrity check (${check.integrity}).`);
    await pruneLocal(backupDir, settings.keep_local);

    const storage = storageFor(db, fetchImpl);
    if (storage) {
      sending = true;
      const root = remoteRoot(settings, companyName);
      gz = `${file}.gz`;
      const size = await gzipFile(file, gz);
      await storage.mkdirs(`${root}/db`);
      await storage.put(`${root}/db/${name}.gz`, fs.createReadStream(gz), size);
      for (const old of remoteToPrune(await storage.list(`${root}/db`), settings.keep_remote_days, now)) await storage.remove(`${root}/db/${old}`);
      result.offsite = "uploaded";

      if (settings.include_files) {
        const todo = filesToSync(db, dataRoot);
        const mark = db.prepare(`
          INSERT INTO backup_remote_files (rel_path, bytes, mtime_ms, uploaded_at) VALUES (?, ?, ?, datetime('now'))
          ON CONFLICT(rel_path) DO UPDATE SET bytes = excluded.bytes, mtime_ms = excluded.mtime_ms, uploaded_at = excluded.uploaded_at
        `);
        const made = new Set();
        for (const f of todo.slice(0, MAX_FILES_PER_RUN)) {
          const dir = `${root}/files/${path.posix.dirname(f.rel)}`;
          if (!made.has(dir)) { await storage.mkdirs(dir); made.add(dir); }
          await storage.put(`${root}/files/${f.rel}`, fs.createReadStream(path.join(dataRoot, f.rel)), f.bytes);
          mark.run(f.rel, f.bytes, f.mtime_ms);
          result.files_uploaded += 1;
        }
        // A big first sync carries on the next night.
        result.files_pending = Math.max(0, todo.length - MAX_FILES_PER_RUN);
      }
    }
  } catch (e) {
    result.status = "failed";
    result.error = String(e?.message || e).slice(0, 1000);
    if (sending && result.offsite !== "uploaded") result.offsite = "failed";
  } finally {
    if (gz) await fsp.rm(gz, { force: true }).catch(() => {});
  }
  finishRun(db, id, result);
  return db.prepare(`SELECT * FROM backup_runs WHERE id = ?`).get(id);
}

/** Downloads the newest off-site copy and proves it opens: the restore test. Never throws. */
export async function runRestoreTest(db, { by = "system", companyName = "", fetchImpl, workDir = os.tmpdir() } = {}) {
  ensureBackupTables(db);
  const settings = getBackupSettings(db);
  const id = startRun(db, "restore_test", by);
  const result = { status: "ok" };
  const tmp = path.join(workDir, `ironlog-restore-test-${process.pid}-${id}.db`);
  try {
    const storage = storageFor(db, fetchImpl);
    if (!storage) throw new Error("Off-site storage is not set up, so there is nothing to restore from.");
    const root = remoteRoot(settings, companyName);
    const copies = (await storage.list(`${root}/db`)).filter((n) => n.startsWith(AUTO_PREFIX) && n.endsWith(".db.gz")).sort();
    const newest = copies.at(-1);
    if (!newest) throw new Error("No off-site copies found yet.");
    const res = await storage.get(`${root}/db/${newest}`);
    await pipeline(Readable.fromWeb(res.body), zlib.createGunzip(), fs.createWriteStream(tmp));
    result.db_file = newest;
    result.db_bytes = (await fsp.stat(tmp)).size;
    const check = checkDatabaseCopy(tmp);
    result.check_result = check;
    if (!check.ok) throw new Error(`The off-site copy ${newest} failed its integrity check (${check.integrity}).`);
    result.offsite = "restored";
  } catch (e) {
    result.status = "failed";
    result.error = String(e?.message || e).slice(0, 1000);
  } finally {
    await fsp.rm(tmp, { force: true }).catch(() => {});
  }
  finishRun(db, id, result);
  return db.prepare(`SELECT * FROM backup_runs WHERE id = ?`).get(id);
}

// ------------------------------------------------------------------ status

function lastRun(db, kind, status = null) {
  return db.prepare(`SELECT * FROM backup_runs WHERE kind IN (${kind.map(() => "?").join(",")})${status ? " AND status = ?" : ""} ORDER BY id DESC LIMIT 1`)
    .get(...kind, ...(status ? [status] : [])) || null;
}

const hoursSince = (sqlTime, now) => (sqlTime ? (now.getTime() - Date.parse(`${String(sqlTime).replace(" ", "T")}Z`)) / 3600000 : Infinity);

/** What Admin → Backups shows, with plain warnings for anything that needs attention. */
export function backupStatus(db, { now = new Date(), dbPath = "", backupDir = "" } = {}) {
  ensureBackupTables(db);
  const settings = getBackupSettings(db);
  const lastBackup = lastRun(db, ["nightly", "manual"]);
  const lastGood = lastRun(db, ["nightly", "manual"], "ok");
  const lastOffsite = db.prepare(`SELECT * FROM backup_runs WHERE offsite = 'uploaded' ORDER BY id DESC LIMIT 1`).get() || null;
  const lastRestore = lastRun(db, ["restore_test"]);
  const warnings = [];
  if (!settings.enabled) warnings.push("Automatic backups are switched off.");
  if (hoursSince(lastGood?.finished_at, now) > 26) warnings.push(lastGood ? "No good backup in the last 26 hours." : "No backup has run yet.");
  if (lastBackup?.status === "failed") warnings.push(`The last backup failed: ${lastBackup.error}`);
  if (!settings.offsite_configured) warnings.push("Off-site storage is not set up: backups stay on this server only.");
  else if (hoursSince(lastOffsite?.finished_at, now) > 26) warnings.push("Nothing has been copied off-site in the last 26 hours.");
  if (settings.offsite_configured && (!lastRestore || hoursSince(lastRestore.finished_at, now) > 32 * 24)) warnings.push("No restore test in the last month.");
  else if (lastRestore?.status === "failed") warnings.push(`The last restore test failed: ${lastRestore.error}`);
  let disk = null;
  try {
    const st = fs.statfsSync(path.dirname(dbPath || "."));
    disk = { free_gb: Number(((st.bavail * st.bsize) / 1e9).toFixed(1)), total_gb: Number(((st.blocks * st.bsize) / 1e9).toFixed(1)) };
    if (disk.free_gb < 2) warnings.push(`Disk space is low: ${disk.free_gb} GB free.`);
  } catch {}
  let dbBytes = null;
  try { dbBytes = fs.statSync(dbPath).size; } catch {}
  const parse = (r) => (r ? { ...r, check_result: r.check_result ? JSON.parse(r.check_result) : null } : null);
  return {
    ok: warnings.length === 0,
    warnings,
    settings,
    last_backup: parse(lastBackup),
    last_good_backup: parse(lastGood),
    last_offsite: parse(lastOffsite),
    last_restore_test: parse(lastRestore),
    recent: db.prepare(`SELECT id, kind, status, started_at, finished_at, db_bytes, offsite, files_uploaded, files_pending, error FROM backup_runs ORDER BY id DESC LIMIT 15`).all(),
    server: {
      uptime_hours: Number((process.uptime() / 3600).toFixed(1)),
      node: process.version,
      db_bytes: dbBytes,
      disk,
      backup_dir: backupDir,
      files_waiting: settings.include_files ? filesToSync(db).length : 0,
    },
  };
}

// ------------------------------------------------------------------ schedule

/** Hour and day of month at site, for the nightly and monthly runs. */
export function siteClock(timezone, now = new Date()) {
  let parts;
  try {
    parts = new Intl.DateTimeFormat("en-CA", { timeZone: timezone, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", hourCycle: "h23" }).formatToParts(now);
  } catch {
    parts = new Intl.DateTimeFormat("en-CA", { timeZone: "UTC", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", hourCycle: "h23" }).formatToParts(now);
  }
  const get = (t) => parts.find((p) => p.type === t)?.value;
  return { date: `${get("year")}-${get("month")}-${get("day")}`, hour: Number(get("hour")), day: Number(get("day")) };
}

/** True when tonight's backup is due: the set hour has come and nothing ran for this site date yet. */
export function nightlyDue(db, { timezone, now = new Date() }) {
  const s = getBackupSettings(db);
  if (!s.enabled) return false;
  const clock = siteClock(timezone, now);
  if (clock.hour < s.run_hour) return false;
  const last = db.prepare(`SELECT started_at FROM backup_runs WHERE kind = 'nightly' ORDER BY id DESC LIMIT 1`).get();
  if (!last) return true;
  return siteClock(timezone, new Date(`${String(last.started_at).replace(" ", "T")}Z`)).date !== clock.date;
}

let schedulerBusy = false;

/** Whether a backup or restore test is running now (scheduler or button). */
export function backupBusy() {
  return schedulerBusy;
}

/**
 * Starts a backup or restore test without waiting for it (a first sync of many
 * photos can outlast a web request). Returns false when one is already running.
 */
export function startInBackground(job, log = console) {
  if (schedulerBusy) return false;
  schedulerBusy = true;
  Promise.resolve()
    .then(job)
    .catch((e) => log.warn?.(`[backup] ${e?.message || e}`))
    .finally(() => { schedulerBusy = false; });
  return true;
}

/** Checks every minute; runs the nightly backup, then on the 1st the restore test. Alerts admins on failure. */
export function startBackupScheduler({ db, backupDir, profile, notify, log = console, intervalMs = 60000 }) {
  ensureBackupTables(db);
  const tick = async () => {
    if (schedulerBusy) return;
    const p = profile();
    if (!nightlyDue(db, { timezone: p.timezone })) return;
    schedulerBusy = true;
    try {
      const run = await runBackup(db, { kind: "nightly", backupDir, companyName: p.short_name || p.company_name });
      if (run.status !== "ok") await notify?.(`Nightly backup failed: ${run.error}`).catch(() => {});
      if (siteClock(p.timezone).day === 1 && getBackupSettings(db).offsite_configured) {
        const test = await runRestoreTest(db, { companyName: p.short_name || p.company_name });
        if (test.status !== "ok") await notify?.(`Monthly restore test failed: ${test.error}`).catch(() => {});
      }
      log.info?.(`[backup] nightly ${run.status}${run.offsite ? `, off-site ${run.offsite}` : ""}`);
    } catch (e) {
      log.warn?.(`[backup] scheduler error: ${e?.message || e}`);
    } finally {
      schedulerBusy = false;
    }
  };
  const timer = setInterval(() => { tick(); }, intervalMs);
  timer.unref?.();
  return timer;
}
