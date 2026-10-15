// IRONLOG/api/routes/licence.routes.js — the customer's licence.
//   GET  /api/licence/state       this site's mode, modules and warnings (any signed-in user)
//   POST /api/licence/install     { key }  administrators only
//   GET  /api/licence/usage       machines and users per site against the bands (administrators)
//   GET  /api/licence/usage.csv   the same for the vendor's invoice
// licenceHook (registered in server.js after sign-in) applies it to every request.
import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";
import { getCompanyProfile } from "../utils/companyProfile.js";
import { MODULES, installLicence, licenceDecision, licenceState, licenceUsage } from "../utils/licensing.js";
import { requireAnyRole } from "../utils/request.js";

const CACHE_MS = 30000;
let cache = new Map();

function siteToday() {
  const tz = getCompanyProfile(db).timezone || "UTC";
  try {
    return new Date().toLocaleDateString("en-CA", { timeZone: tz });
  } catch {
    return new Date().toISOString().slice(0, 10);
  }
}

const siteOf = (req) => String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";

/** The licence state for a site, cached for 30 seconds. */
export function currentLicenceState(siteCode) {
  const hit = cache.get(siteCode);
  if (hit && Date.now() - hit.at < CACHE_MS) return hit.state;
  const state = licenceState(db, { siteCode, today: siteToday() });
  cache.set(siteCode, { at: Date.now(), state });
  return state;
}

export function clearLicenceCache() {
  cache = new Map();
}

/** onRequest hook: refuses modules outside the licence, and saving when read-only. */
export async function licenceHook(req, reply) {
  const urlPath = req.url.split("?")[0];
  if (!urlPath.startsWith("/api/")) return;
  let state;
  try {
    state = currentLicenceState(siteOf(req));
  } catch {
    return; // A licence problem must never take the system down.
  }
  const d = licenceDecision(state, { method: req.method, urlPath });
  if (d) return reply.code(d.status).send({ ok: false, error: d.error, licence: true, module: d.module || null });
}

export default async function licenceRoutes(app) {
  ensureAuditTable(db);

  app.get("/state", async (req) => {
    const s = currentLicenceState(siteOf(req));
    return {
      ok: true,
      ...s,
      module_labels: Object.fromEntries(Object.entries(MODULES).map(([k, m]) => [k, m.label])),
      module_tabs: Object.fromEntries(Object.entries(MODULES).map(([k, m]) => [k, m.tabs])),
    };
  });

  app.post("/install", async (req, reply) => {
    if (!requireAnyRole(req, reply, ["admin"])) return;
    const r = installLicence(db, req.body?.key, String(req.headers["x-user-name"] || ""));
    if (!r.ok) return reply.code(400).send(r);
    clearLicenceCache();
    writeAudit(db, req, { module: "licence", action: "install", entity_type: "licence", entity_id: String(r.licence.licence_id || ""), payload: { licensee: r.licence.licensee, expires_at: r.licence.expires_at, sites: Object.keys(r.licence.sites || {}) } });
    return { ok: true, state: currentLicenceState(siteOf(req)) };
  });

  app.get("/usage", async (req, reply) => {
    if (!requireAnyRole(req, reply, ["admin"])) return;
    return { ok: true, ...licenceUsage(db, { today: siteToday() }) };
  });

  app.get("/usage.csv", async (req, reply) => {
    if (!requireAnyRole(req, reply, ["admin"])) return;
    const u = licenceUsage(db, { today: siteToday() });
    const p = getCompanyProfile(db);
    const q = (v) => (/[",\n]/.test(String(v ?? "")) ? `"${String(v).replace(/"/g, '""')}"` : String(v ?? ""));
    const lines = [
      ["Company", p.company_name || p.short_name || ""],
      ["Generated (UTC)", u.generated_at.slice(0, 16).replace("T", " ")],
      ["Active users", u.active_users],
      [],
      ["Site", "Package", "Active machines", "Machines licensed", "Over band", "Modules"],
      ...u.sites.map((s) => [s.site_code, s.package || "", s.active_assets, s.asset_limit ?? "", s.over_band ? "Yes" : "", (s.modules || []).map((m) => MODULES[m]?.label || m).join("; ")]),
    ];
    reply.header("Content-Type", "text/csv; charset=utf-8");
    reply.header("Content-Disposition", `attachment; filename="ironlog-usage-${siteToday()}.csv"`);
    return `﻿${lines.map((r) => r.map(q).join(",")).join("\r\n")}`;
  });
}
