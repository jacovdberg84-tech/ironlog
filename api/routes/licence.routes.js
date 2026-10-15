// IRONLOG/api/routes/licence.routes.js — the customer's licence.
//   GET  /api/licence/state       this site's mode, modules and warnings (any signed-in user)
//   POST /api/licence/install     { key }  administrators only
//   GET  /api/licence/usage       machines and users per site against the bands (administrators)
//   GET  /api/licence/usage.csv   the same for the vendor's invoice
// licenceHook (registered in server.js after sign-in) applies it to every request.
import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";
import { getCompanyProfile } from "../utils/companyProfile.js";
import { MODULES, installLicence, licenceDecision, licenceState, licenceUsage, vendorPublicKeys } from "../utils/licensing.js";
import {
  catalogue,
  createSigningKey,
  isIssuer,
  issueLicence,
  issuedKey,
  listIssued,
  loadSigningKey,
  localPublicKeys,
  publicKeyOf,
  saveBand,
  savePackage,
  savePrices,
  signingKeyPem,
} from "../utils/licenceIssuer.js";
import { getRoles, requireAnyRole } from "../utils/request.js";

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

// The vendor's keys, plus this server's own when it is the issuer.
const trustedKeys = () => [...vendorPublicKeys(), ...localPublicKeys()];

const siteOf = (req) => String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";

/** The licence state for a site, cached for 30 seconds. */
export function currentLicenceState(siteCode) {
  const hit = cache.get(siteCode);
  if (hit && Date.now() - hit.at < CACHE_MS) return hit.state;
  const state = licenceState(db, { siteCode, today: siteToday(), publicKeys: trustedKeys() });
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
    const r = installLicence(db, req.body?.key, String(req.headers["x-user-name"] || ""), { publicKeys: trustedKeys() });
    if (!r.ok) return reply.code(400).send(r);
    clearLicenceCache();
    writeAudit(db, req, { module: "licence", action: "install", entity_type: "licence", entity_id: String(r.licence.licence_id || ""), payload: { licensee: r.licence.licensee, expires_at: r.licence.expires_at, sites: Object.keys(r.licence.sites || {}) } });
    return { ok: true, state: currentLicenceState(siteOf(req)) };
  });

  app.get("/usage", async (req, reply) => {
    if (!requireAnyRole(req, reply, ["admin"])) return;
    return { ok: true, ...licenceUsage(db, { today: siteToday(), publicKeys: trustedKeys() }) };
  });

  app.get("/usage.csv", async (req, reply) => {
    if (!requireAnyRole(req, reply, ["admin"])) return;
    const u = licenceUsage(db, { today: siteToday(), publicKeys: trustedKeys() });
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

  // ---- The vendor's licence issuer: only the IronLog owner (isIssuer), signed in as an admin.
  const issuerOnly = (req, reply) => {
    if (isIssuer(req.headers["x-user-name"], getRoles(req))) return true;
    reply.code(403).send({ ok: false, error: "Only the IronLog owner can issue licences." });
    return false;
  };
  const issuerAudit = (req, action, payload) => {
    try { writeAudit(db, req, { module: "licence_issuer", action, entity_type: "licence", payload }); } catch {}
  };

  app.get("/issuer/status", async (req) => {
    const allowed = isIssuer(req.headers["x-user-name"], getRoles(req));
    if (!allowed) return { ok: true, allowed: false };
    const key = loadSigningKey();
    return { ok: true, allowed: true, has_signing_key: Boolean(key), public_key: key ? publicKeyOf(key) : null, built_in: key ? vendorPublicKeys().includes(publicKeyOf(key)) : false };
  });

  app.post("/issuer/setup", async (req, reply) => {
    if (!issuerOnly(req, reply)) return;
    const r = createSigningKey();
    if (!r.ok) return reply.code(409).send(r);
    issuerAudit(req, "signing_key_created", { public_key: r.public_key });
    clearLicenceCache();
    return r;
  });

  app.get("/issuer/signing-key.pem", async (req, reply) => {
    if (!issuerOnly(req, reply)) return;
    const pem = signingKeyPem();
    if (!pem) return reply.code(404).send({ ok: false, error: "No signing key on this server yet." });
    issuerAudit(req, "signing_key_downloaded", {});
    reply.header("Content-Type", "application/x-pem-file");
    reply.header("Content-Disposition", 'attachment; filename="ironlog-licence-signing-key.pem"');
    return pem;
  });

  app.get("/issuer/catalogue", async (req, reply) => {
    if (!issuerOnly(req, reply)) return;
    return { ok: true, ...catalogue(db), issued: listIssued(db) };
  });

  app.post("/issuer/packages", async (req, reply) => {
    if (!issuerOnly(req, reply)) return;
    const r = savePackage(db, req.body || {});
    if (!r.ok) return reply.code(400).send(r);
    issuerAudit(req, "package_save", { name: req.body?.name, modules: req.body?.modules });
    return r;
  });

  app.post("/issuer/bands", async (req, reply) => {
    if (!issuerOnly(req, reply)) return;
    const r = saveBand(db, req.body || {});
    if (!r.ok) return reply.code(400).send(r);
    issuerAudit(req, "band_save", { label: req.body?.label, max_assets: req.body?.max_assets });
    return r;
  });

  app.put("/issuer/prices", async (req, reply) => {
    if (!issuerOnly(req, reply)) return;
    const r = savePrices(db, req.body?.prices);
    if (!r.ok) return reply.code(400).send(r);
    issuerAudit(req, "prices_save", { count: (req.body?.prices || []).length });
    return r;
  });

  app.post("/issuer/issue", async (req, reply) => {
    if (!issuerOnly(req, reply)) return;
    const r = issueLicence(db, req.body || {}, { by: String(req.headers["x-user-name"] || ""), today: siteToday() });
    if (!r.ok) return reply.code(400).send(r);
    issuerAudit(req, "issue", { licence_id: r.licence_id, licensee: r.payload.licensee, expires_at: r.payload.expires_at, sites: Object.keys(r.payload.sites) });
    return r;
  });

  app.get("/issuer/issued/:id/key", async (req, reply) => {
    if (!issuerOnly(req, reply)) return;
    const k = issuedKey(db, req.params.id);
    if (!k) return reply.code(404).send({ ok: false, error: "Licence not found." });
    return { ok: true, ...k };
  });
}
