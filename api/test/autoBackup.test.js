import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import zlib from "node:zlib";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-autobackup-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.JUNE_CONNECTOR_ENCRYPTION_SECRET = "test-secret-for-backup-passwords-0123456789";
const Fastify = (await import("fastify")).default;
const { db } = await import("../db/client.js");
const ab = await import("../utils/autoBackup.js");
process.on("exit", () => { try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

db.exec(`
  CREATE TABLE IF NOT EXISTS assets (id INTEGER PRIMARY KEY, asset_code TEXT);
  INSERT INTO assets (asset_code) VALUES ('E501AM'), ('D600AM');
`);
const backupDir = path.join(tempDir, "backups");
const dataRoot = path.join(tempDir, "data");
fs.mkdirSync(path.join(dataRoot, "uploads", "wo"), { recursive: true });
fs.writeFileSync(path.join(dataRoot, "uploads", "wo", "photo 1.jpg"), "jpeg-bytes");
fs.mkdirSync(path.join(dataRoot, "branding"), { recursive: true });
fs.writeFileSync(path.join(dataRoot, "branding", "logo.png"), "png");

/** An in-memory WebDAV server behind fetch. */
function fakeDav({ user = "u123", password = "pw" } = {}) {
  const files = new Map();
  const dirs = new Set([""]);
  const calls = [];
  const fetchImpl = async (url, init) => {
    const p = decodeURIComponent(new URL(url).pathname).replace(/\/+$/, "");
    calls.push(`${init.method} ${p}`);
    if (init.headers.Authorization !== `Basic ${Buffer.from(`${user}:${password}`).toString("base64")}`) return new Response("", { status: 401 });
    if (init.method === "MKCOL") { if (dirs.has(p)) return new Response("", { status: 405 }); dirs.add(p); return new Response("", { status: 201 }); }
    if (init.method === "PUT") {
      const chunks = [];
      for await (const c of init.body) chunks.push(Buffer.from(c));
      files.set(p, Buffer.concat(chunks));
      return new Response("", { status: 201 });
    }
    if (init.method === "GET") return files.has(p) ? new Response(files.get(p), { status: 200 }) : new Response("", { status: 404 });
    if (init.method === "DELETE") { files.delete(p); return new Response(null, { status: 204 }); }
    if (init.method === "PROPFIND") {
      if (!dirs.has(p)) return new Response("", { status: 404 });
      const kids = [...files.keys()].filter((f) => path.posix.dirname(f) === (p || "/"));
      const xml = `<?xml version="1.0"?><d:multistatus xmlns:d="DAV:"><d:response><d:href>${p}/</d:href></d:response>${kids.map((k) => `<d:response><d:href>${encodeURI(k)}</d:href></d:response>`).join("")}</d:multistatus>`;
      return new Response(xml, { status: 207 });
    }
    return new Response("", { status: 400 });
  };
  return { files, dirs, calls, fetchImpl };
}

test("without off-site storage, a checked copy is kept on the server and the status says so", async () => {
  const run = await ab.runBackup(db, { kind: "manual", backupDir, dataRoot, companyName: "AML" });
  assert.equal(run.status, "ok");
  assert.equal(run.offsite, "not set up");
  assert.match(run.db_file, /^ironlog_auto_\d{8}_\d{6}\.db$/);
  const check = JSON.parse(run.check_result);
  assert.equal(check.ok, true);
  assert.equal(check.assets, 2);
  const status = ab.backupStatus(db, { dbPath: process.env.DB_PATH, backupDir });
  assert.ok(status.warnings.some((w) => /Off-site storage is not set up/.test(w)));
  assert.ok(!status.warnings.some((w) => /No good backup/.test(w)));
});

test("storage settings: https only, password stored encrypted and never returned", () => {
  assert.equal(ab.saveBackupSettings(db, { webdav_url: "ftp://x" }).ok, false);
  const r = ab.saveBackupSettings(db, { webdav_url: "https://u123.your-storagebox.de/", webdav_user: "u123", webdav_password: "pw", run_hour: 3, keep_local: 99 }, "jaco");
  assert.equal(r.ok, true);
  assert.equal(r.settings.webdav_url, "https://u123.your-storagebox.de");
  assert.equal(r.settings.has_password, true);
  assert.equal(r.settings.run_hour, 3);
  assert.equal(r.settings.keep_local, 7, "out-of-range values keep the old setting");
  assert.equal("webdav_password" in r.settings, false);
  const raw = db.prepare(`SELECT webdav_password_enc FROM backup_settings`).get().webdav_password_enc;
  assert.ok(raw && !raw.includes("pw"), "not stored in plain text");
  assert.equal(ab.saveBackupSettings(db, { webdav_password: "" }).settings.has_password, true, "a blank password box keeps the saved one");
});

test("off-site: the database goes up compressed, uploaded files sync once, old copies are pruned", async () => {
  const dav = fakeDav();
  const now = new Date("2026-10-15T00:30:00Z");
  // Twelve old nightly copies already off-site: the oldest beyond 30 days should go, the newest 7 always stay.
  dav.dirs.add("/ironlog"); dav.dirs.add("/ironlog/aml"); dav.dirs.add("/ironlog/aml/db");
  for (let i = 1; i <= 12; i++) dav.files.set(`/ironlog/aml/db/ironlog_auto_2026${String(i < 10 ? 8 : 9).padStart(2, "0")}${String(i).padStart(2, "0")}_000000.db.gz`, Buffer.from("old"));
  const run = await ab.runBackup(db, { kind: "nightly", backupDir, dataRoot, companyName: "AML", fetchImpl: dav.fetchImpl, now });
  assert.equal(run.status, "ok", run.error);
  assert.equal(run.offsite, "uploaded");
  assert.equal(run.files_uploaded, 2);
  const gz = dav.files.get(`/ironlog/aml/db/${run.db_file}.gz`);
  assert.ok(gz, "database copy is off-site");
  assert.equal(zlib.gunzipSync(gz).subarray(0, 15).toString(), "SQLite format 3");
  assert.equal(dav.files.get("/ironlog/aml/files/uploads/wo/photo 1.jpg").toString(), "jpeg-bytes");
  const left = [...dav.files.keys()].filter((k) => k.includes("/db/ironlog_auto_2026"));
  assert.ok(!left.includes("/ironlog/aml/db/ironlog_auto_20260801_000000.db.gz"), "a copy older than 30 days is pruned");
  assert.ok(left.includes("/ironlog/aml/db/ironlog_auto_20260912_000000.db.gz"));

  const again = await ab.runBackup(db, { kind: "manual", backupDir, dataRoot, companyName: "AML", fetchImpl: dav.fetchImpl, now: new Date("2026-10-15T00:31:00Z") });
  assert.equal(again.files_uploaded, 0, "unchanged files are not sent again");
  fs.writeFileSync(path.join(dataRoot, "uploads", "wo", "photo 1.jpg"), "edited-jpeg-bytes");
  const third = await ab.runBackup(db, { kind: "manual", backupDir, dataRoot, companyName: "AML", fetchImpl: dav.fetchImpl, now: new Date("2026-10-15T00:32:00Z") });
  assert.equal(third.files_uploaded, 1, "a changed file is sent again");
});

test("a wrong password fails the run, keeps the local copy and shows in the status", async () => {
  const dav = fakeDav({ password: "different" });
  const run = await ab.runBackup(db, { kind: "manual", backupDir, dataRoot, companyName: "AML", fetchImpl: dav.fetchImpl });
  assert.equal(run.status, "failed");
  assert.equal(run.offsite, "failed");
  assert.match(run.error, /HTTP 401 \(check the username and password\)/);
  assert.ok(fs.existsSync(path.join(backupDir, run.db_file)), "the local copy is still made");
  assert.ok(ab.backupStatus(db, { dbPath: process.env.DB_PATH, backupDir }).warnings.some((w) => /last backup failed/.test(w)));
});

test("restore test: the newest off-site copy is downloaded and opened; a damaged one fails", async () => {
  const dav = fakeDav();
  const run = await ab.runBackup(db, { kind: "manual", backupDir, dataRoot, companyName: "AML", fetchImpl: dav.fetchImpl });
  const good = await ab.runRestoreTest(db, { companyName: "AML", fetchImpl: dav.fetchImpl, workDir: tempDir });
  assert.equal(good.status, "ok", good.error);
  assert.equal(good.db_file, `${run.db_file}.gz`);
  assert.equal(JSON.parse(good.check_result).assets, 2);
  dav.files.set("/ironlog/aml/db/ironlog_auto_29991231_000000.db.gz", zlib.gzipSync(Buffer.from("not a database")));
  const bad = await ab.runRestoreTest(db, { companyName: "AML", fetchImpl: dav.fetchImpl, workDir: tempDir });
  assert.equal(bad.status, "failed");
  assert.match(bad.error, /not a database|integrity/i);
});

test("nightly runs once per site day, after the set hour, in the company's time zone", () => {
  db.exec(`DELETE FROM backup_runs WHERE kind = 'nightly'`);
  ab.saveBackupSettings(db, { run_hour: 2 });
  // 23:30 UTC on the 14th is 01:30 on the 15th in Maputo: not yet 02:00.
  assert.equal(ab.nightlyDue(db, { timezone: "Africa/Maputo", now: new Date("2026-10-14T23:30:00Z") }), false);
  assert.equal(ab.nightlyDue(db, { timezone: "Africa/Maputo", now: new Date("2026-10-15T00:05:00Z") }), true);
  db.prepare(`INSERT INTO backup_runs (kind, status, started_at) VALUES ('nightly', 'ok', '2026-10-15 00:05:00')`).run();
  assert.equal(ab.nightlyDue(db, { timezone: "Africa/Maputo", now: new Date("2026-10-15T10:00:00Z") }), false, "already ran today");
  assert.equal(ab.nightlyDue(db, { timezone: "Africa/Maputo", now: new Date("2026-10-16T00:05:00Z") }), true);
  ab.saveBackupSettings(db, { enabled: false });
  assert.equal(ab.nightlyDue(db, { timezone: "Africa/Maputo", now: new Date("2026-10-17T00:05:00Z") }), false, "switched off");
  ab.saveBackupSettings(db, { enabled: true });
});

test("only administrators change backup settings; supervisors can see the status", async (t) => {
  const backupsRoutes = (await import("../routes/backups.routes.js")).default;
  const app = Fastify();
  await app.register(backupsRoutes, { prefix: "/api/admin/backups" });
  t.after(() => app.close());
  const as = (role) => ({ "x-user-roles": role, "x-user-name": role });
  assert.equal((await app.inject({ method: "GET", url: "/api/admin/backups/status", headers: as("artisan") })).statusCode, 403);
  const st = await app.inject({ method: "GET", url: "/api/admin/backups/status", headers: as("supervisor") });
  assert.equal(st.statusCode, 200);
  assert.equal(JSON.stringify(st.json()).includes("pw"), false, "the password never leaves the server");
  assert.equal((await app.inject({ method: "PUT", url: "/api/admin/backups/settings", headers: as("supervisor"), payload: { run_hour: 4 } })).statusCode, 403);
  const ok = await app.inject({ method: "PUT", url: "/api/admin/backups/settings", headers: as("admin"), payload: { run_hour: 4, webdav_password: "secret-new" } });
  assert.equal(ok.statusCode, 200);
  const audit = db.prepare(`SELECT payload_json FROM audit_logs WHERE module = 'backups' AND action = 'settings_update' ORDER BY id DESC LIMIT 1`).get();
  assert.equal(audit.payload_json.includes("secret-new"), false, "the password is not written to the audit log");
  assert.match(audit.payload_json, /"password_changed":true/);
});

test("an unreachable storage address gets a message that says what to check", async () => {
  const c = ab.webdavClient({ url: "https://u1.your-storagebox.de", user: "u", password: "p", fetchImpl: async () => { throw new TypeError("fetch failed"); } });
  await assert.rejects(c.test(), /Could not reach https:\/\/u1\.your-storagebox\.de\. Check the address, and that WebDAV is switched on/);
});

test("Back up now starts in the background and refuses a second run while one is going", async (t) => {
  const backupsRoutes = (await import("../routes/backups.routes.js")).default;
  const app = Fastify();
  await app.register(backupsRoutes, { prefix: "/api/admin/backups" });
  t.after(() => app.close());
  ab.saveBackupSettings(db, { clear_password: true });
  const h = { "x-user-roles": "admin", "x-user-name": "jaco" };
  const first = await app.inject({ method: "POST", url: "/api/admin/backups/run-now", headers: h });
  assert.equal(first.statusCode, 202);
  const second = await app.inject({ method: "POST", url: "/api/admin/backups/restore-test", headers: h });
  assert.equal(second.statusCode, 409);
  for (let i = 0; i < 50 && ab.backupBusy(); i++) await new Promise((r) => setTimeout(r, 50));
  assert.equal(ab.backupBusy(), false);
  const status = (await app.inject({ method: "GET", url: "/api/admin/backups/status", headers: h })).json();
  assert.equal(status.recent[0].kind, "manual");
  assert.equal(status.recent[0].status, "ok");
});
