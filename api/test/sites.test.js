import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import Database from "better-sqlite3";
import { checkAllowedLocations, knownSites, repairSiteDataOnce, unwrapSiteList } from "../utils/sites.js";

// AML's live data, October 2026: one site, machines with no site, and a garbled "workshop".
function amlLike() {
  const d = new Database(":memory:");
  d.exec(`CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT, site_code TEXT);
          CREATE TABLE work_orders (id INTEGER PRIMARY KEY, site_code TEXT);
          CREATE TABLE tech_shifts (id INTEGER PRIMARY KEY, site_code TEXT);
          CREATE TABLE users (id INTEGER PRIMARY KEY, username TEXT, allowed_locations TEXT);
          INSERT INTO assets (asset_code, site_code) VALUES ('E501AM', NULL), ('D600AM', ''), ('L1', 'main');
          INSERT INTO work_orders (site_code) VALUES ('main'), ('["workshop"]'), ('["workshop"]');
          INSERT INTO tech_shifts (site_code) VALUES ('["workshop"]'), ('main');
          INSERT INTO users (username, allowed_locations) VALUES
            ('foreman', '["main"]'), ('clerk', '["dashboard","daily input","breakdown ops","reports"]'),
            ('fuelman', '["site operations","fuel"]'), ('tech', '["[\\"workshop\\"]"]'), ('boss', NULL);`);
  return d;
}

test("site lists saved as text are opened up; screen names are not sites", () => {
  assert.deepEqual(unwrapSiteList('["[\\"workshop\\"]"]'), ["workshop"]);
  assert.deepEqual(unwrapSiteList(' Main , north-pit '), ["main", "north-pit"]);
  assert.deepEqual(unwrapSiteList(["main", "MAIN"]), ["main"]);
  const d = amlLike();
  assert.deepEqual(knownSites(d), ["main"], "garbled codes are not sites");
  assert.deepEqual(checkAllowedLocations(d, "main"), { ok: true, sites: ["main"] });
  assert.deepEqual(checkAllowedLocations(d, ""), { ok: true, sites: null }, "empty = all sites");
  const bad = checkAllowedLocations(d, "dashboard, fuel");
  assert.equal(bad.ok, false);
  assert.match(bad.error, /Not a site: dashboard, fuel\. Sites: main\. Leave empty for all sites/);
  d.exec(`CREATE TABLE site_profiles (site_code TEXT, active INTEGER DEFAULT 1); INSERT INTO site_profiles VALUES ('north-pit', 1);`);
  assert.deepEqual(checkAllowedLocations(d, '["north-pit","main"]').sites, ["north-pit", "main"]);
});

test("the one-time tidy-up copies the database, fixes records and users, audits each, and runs once", async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ironlog-sites-"));
  const file = path.join(dir, "ironlog.db");
  const src = amlLike();
  await src.backup(file);
  const d = new Database(file);
  const audited = [];
  const r = await repairSiteDataOnce(d, { backupDir: path.join(dir, "copies"), audit: (e) => audited.push(e), log: { info() {} } });
  assert.equal(r.records.length, 5, "3 garbled + 2 machines with no site");
  assert.deepEqual(d.prepare(`SELECT DISTINCT site_code FROM work_orders`).all().map((x) => x.site_code), ["main"]);
  assert.deepEqual(d.prepare(`SELECT DISTINCT site_code FROM tech_shifts`).all().map((x) => x.site_code), ["main"]);
  assert.deepEqual(d.prepare(`SELECT site_code FROM assets ORDER BY id`).all().map((x) => x.site_code), ["main", "main", "main"]);
  const users = Object.fromEntries(d.prepare(`SELECT username, allowed_locations FROM users`).all().map((u) => [u.username, u.allowed_locations]));
  assert.deepEqual(users, { foreman: '["main"]', clerk: null, fuelman: null, tech: null, boss: null }, "foreman keeps main; the 3 wrong settings are cleared");
  assert.deepEqual(r.users.map((u) => u.username), ["clerk", "fuelman", "tech"]);
  assert.equal(audited.length, 8);
  assert.deepEqual(audited.find((a) => a.entity_type === "work_orders").before, { site_code: '["workshop"]' });
  const copy = new Database(r.backup, { readonly: true });
  assert.equal(copy.prepare(`SELECT COUNT(*) AS n FROM work_orders WHERE site_code LIKE '[%'`).get().n, 2, "the copy holds the data as it was");
  assert.equal(await repairSiteDataOnce(d, { backupDir: path.join(dir, "copies") }), null, "runs once");
  fs.rmSync(dir, { recursive: true, force: true });
});

test("a clean multi-site company is left alone", async () => {
  const d = new Database(":memory:");
  d.exec(`CREATE TABLE assets (id INTEGER PRIMARY KEY, site_code TEXT); CREATE TABLE users (id INTEGER PRIMARY KEY, username TEXT, allowed_locations TEXT);
          INSERT INTO assets (site_code) VALUES ('north'), ('south'); INSERT INTO users (username, allowed_locations) VALUES ('n', '["north"]');`);
  const r = await repairSiteDataOnce(d, {});
  assert.deepEqual([r.records.length, r.users.length, r.backup], [0, 0, null]);
  assert.deepEqual(knownSites(d), ["main", "north", "south"]);
});
