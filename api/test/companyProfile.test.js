import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import Database from "better-sqlite3";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-company-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
const Fastify = (await import("fastify")).default;
const { db } = await import("../db/client.js");
const cp = await import("../utils/companyProfile.js");
const { briefingSite, securityCheck } = await import("../utils/juneDailyBriefing.js");
const companyRoutes = (await import("../routes/company.routes.js")).default;
process.on("exit", () => { try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

const freshDb = (withAssets) => {
  const d = new Database(":memory:");
  d.exec(`CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT)`);
  if (withAssets) d.exec(`INSERT INTO assets (asset_code) VALUES ('E501AM')`);
  return d;
};

test("an existing install keeps its AML wording; a new company starts neutral", () => {
  const aml = cp.getCompanyProfile(freshDb(true));
  assert.equal(aml.short_name, "AML");
  assert.equal(aml.header_prefix, "AML");
  assert.equal(aml.timezone, "Africa/Maputo");
  assert.equal(aml.location, "Palma, Cabo Delgado");
  assert.equal(aml.site_name, "", "the site name is shared with every PDF header, so it is not filled in");
  assert.equal(cp.reportTitle(aml, "Daily operations report"), "AML · IronLog — Daily operations report");
  assert.equal(cp.reportBanner(aml, "Daily operations"), "AML / DAILY OPERATIONS");

  const fresh = cp.getCompanyProfile(freshDb(false));
  assert.equal(fresh.short_name, "");
  assert.equal(fresh.location, "");
  assert.equal(fresh.timezone, "UTC");
  assert.equal(fresh.currency, "USD");
  assert.equal(cp.reportTitle(fresh, "Daily operations report"), "IronLog — Daily operations report");
  assert.equal(cp.reportBanner(fresh, "Daily operations"), "DAILY OPERATIONS");
});

test("the profile is filled only once, so an admin's changes stick", () => {
  const d = freshDb(true);
  cp.getCompanyProfile(d);
  assert.equal(cp.saveCompanyProfile(d, { short_name: "", company_name: "Mozambique Minerals Lda" }).ok, true);
  const p = cp.getCompanyProfile(d);
  assert.equal(p.short_name, "");
  assert.equal(p.header_prefix, "Mozambique Minerals Lda", "the company name stands in when there is no short name");
});

test("saving checks currency, time zone, coordinates and email, and saves nothing when one is wrong", () => {
  const d = freshDb(false);
  const bad = cp.saveCompanyProfile(d, { currency: "dollars", timezone: "Mars/Olympus", latitude: "95", email: "nope", short_name: "XYZ" });
  assert.equal(bad.ok, false);
  assert.equal(bad.errors.length, 4);
  assert.equal(cp.getCompanyProfile(d).short_name, "", "nothing saved");
  const good = cp.saveCompanyProfile(d, { currency: "zar", timezone: "Africa/Johannesburg", latitude: "-26.2", longitude: "28.04", email: "ops@example.com", short_name: "  NewCo  " });
  assert.equal(good.ok, true);
  assert.equal(good.profile.currency, "ZAR");
  assert.equal(good.profile.short_name, "NewCo");
  assert.equal(cp.saveCompanyProfile(d, { currency: "DOL" }).ok, false, "three letters is not enough: it must be a real currency");
  assert.equal(cp.saveCompanyProfile(d, { currency: "mzn" }).profile.currency, "MZN");
});

test("June's briefing site comes from the profile, and server settings still win", () => {
  const p = { location: "Tete", country: "Mozambique", latitude: "-16.15", longitude: "33.58", timezone: "Africa/Maputo", nearby: "Tete, Moatize" };
  const site = briefingSite({}, cp.briefingSiteFromProfile(p));
  assert.deepEqual([site.name, site.latitude, site.nearby], ["Tete", -16.15, ["Tete", "Moatize"]]);
  assert.equal(briefingSite({ JUNE_BRIEFING_LOCATION: "Pemba" }, cp.briefingSiteFromProfile(p)).name, "Pemba");
  const blank = briefingSite({}, cp.briefingSiteFromProfile({}));
  assert.equal(blank.latitude, null);
  assert.match(securityCheck(blank, "2026-10-15").search_for, /near the site in the last 7 days/);
});

test("anyone signed in can read the profile; only an administrator can change it", async (t) => {
  db.exec(`CREATE TABLE IF NOT EXISTS assets (id INTEGER PRIMARY KEY, asset_code TEXT)`);
  const app = Fastify();
  await app.register(companyRoutes, { prefix: "/api/company" });
  t.after(() => app.close());
  const as = (role) => ({ "x-user-roles": role, "x-user-role": role, "x-user-name": role });
  const got = await app.inject({ method: "GET", url: "/api/company/profile", headers: as("artisan") });
  assert.equal(got.statusCode, 200);
  assert.ok(got.json().fields.timezone);
  assert.equal((await app.inject({ method: "PUT", url: "/api/company/profile", headers: as("plant_manager"), payload: { short_name: "X" } })).statusCode, 403);
  const bad = await app.inject({ method: "PUT", url: "/api/company/profile", headers: as("admin"), payload: { currency: "12" } });
  assert.equal(bad.statusCode, 400);
  assert.match(bad.json().error, /Currency/);
  const ok = await app.inject({ method: "PUT", url: "/api/company/profile", headers: as("admin"), payload: { short_name: "PILOT", country: "Zambia" } });
  assert.equal(ok.json().profile.header_prefix, "PILOT");
  const audit = db.prepare(`SELECT payload_json FROM audit_logs WHERE module = 'company' ORDER BY id DESC LIMIT 1`).get();
  assert.match(audit.payload_json, /"short_name":\{"from":"","to":"PILOT"\}/);
});
