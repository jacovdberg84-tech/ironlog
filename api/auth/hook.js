// IRONLOG/api/auth/hook.js — optional Bearer session + IRONLOG_AUTH_REQUIRED
import { db } from "../db/client.js";
export { isAuthRequired } from "./config.js";
import { isAuthRequired, isPublicAuthRequest } from "./config.js";
import { defaultSite, knownSitesCached, unwrapSiteList } from "../utils/sites.js";
import { API_KEY_PREFIX, keyAllows, resolveApiKey, touchApiKey, underRateLimit } from "../utils/apiKeys.js";
import { touchSession } from "../utils/signInSecurity.js";

function resolveSession(token) {
  try {
    return db
      .prepare(`
        SELECT u.username, u.role, u.roles_json, u.active, u.allowed_locations, u.department
        FROM auth_sessions s
        JOIN users u ON u.id = s.user_id
        WHERE s.token = ?
          AND datetime(s.expires_at) > datetime('now')
      `)
      .get(token);
  } catch {
    // Fresh DBs may not have auth tables yet during early startup.
    return null;
  }
}

function getPermissionsForRoles(rolesIn) {
  const roles = Array.from(
    new Set(
      (Array.isArray(rolesIn) ? rolesIn : [])
        .map((r) => String(r || "").trim().toLowerCase())
        .filter(Boolean)
    )
  );
  if (!roles.length) return [];
  try {
    const marks = roles.map(() => "?").join(",");
    const rows = db.prepare(`
      SELECT DISTINCT permission_key
      FROM role_permissions
      WHERE role_key IN (${marks})
      ORDER BY permission_key ASC
    `).all(...roles);
    return rows.map((r) => String(r.permission_key || "").trim()).filter(Boolean);
  } catch {
    return [];
  }
}

function parseAllowedLocations(raw) {
  if (!raw) return null;
  // Lists saved as text ('["workshop"]') are opened up (utils/sites.js).
  const out = unwrapSiteList(raw);
  return out.length ? out : null;
}

const VALID_ROLES = [
  "admin",
  "executive",
  "finance",
  "storeman",
  "procurement",
  "hr_manager",
  "quality_manager",
  "site_manager",
  "plant_manager",
  "workshop_admin",
  "plant_clerk",
  "supervisor",
  "artisan",
  "operator",
  // Legacy role retained for backward compatibility.
  "stores",
];

const WORKSHOP_ADMIN_PREFIXES = [
  "/api/assets",
  "/api/maintenance",
  "/api/workorders",
  "/api/breakdowns",
  "/api/breakdown-ops",
  "/api/alerts",
  "/api/approvals",
  "/api/stock",
  "/api/procurement",
  "/api/reports",
  "/api/telematics",
  "/api/safety",
];

function expandRouteRoles(roles, url = "") {
  const out = new Set(roles);
  if (out.has("workshop_admin") && WORKSHOP_ADMIN_PREFIXES.some((prefix) => url.startsWith(prefix))) {
    out.add("supervisor");
  }
  if (out.has("plant_clerk")) out.add("operator");
  if (out.has("storeman")) out.add("stores");
  if (out.has("stores")) out.add("storeman");
  return [...out];
}

// Many routes check only x-user-role against the older role names (supervisor,
// operator, stores). Prefer those over the newer names they stand in for, so a
// storeman passes "stores" checks just like a workshop admin passes
// "supervisor" ones. x-user-roles still carries every role.
const NEWER_ROLE_NAMES = new Set(["workshop_admin", "plant_clerk", "storeman"]);
export function primaryRouteRole(routeRoles) {
  return routeRoles.find((role) => !NEWER_ROLE_NAMES.has(role)) || routeRoles[0];
}

function applyLegacyRouteRoles(req) {
  const roles = [
    ...String(req.headers["x-user-roles"] || "").split(","),
    ...String(req.headers["x-user-role"] || "").split(","),
  ]
    .map((role) => role.trim().toLowerCase())
    .filter((role) => VALID_ROLES.includes(role));
  if (!roles.length) return;
  const routeRoles = expandRouteRoles([...new Set(roles)], req.url.split("?")[0]);
  req.headers["x-user-role"] = primaryRouteRole(routeRoles);
  req.headers["x-user-roles"] = routeRoles.join(",");
}

function parseRoles(raw, fallbackRole = "operator") {
  let arr = [];
  if (raw != null && raw !== "") {
    try {
      const v = typeof raw === "string" ? JSON.parse(raw) : raw;
      if (Array.isArray(v)) arr = v.map((x) => String(x || "").trim().toLowerCase());
    } catch {}
  }
  if (!arr.length && fallbackRole) arr = [String(fallbackRole || "").trim().toLowerCase()];
  const normalized = arr.map((r) => (r === "stores" ? "storeman" : r));
  const out = Array.from(new Set(normalized.filter((r) => VALID_ROLES.includes(r))));
  return out.length ? out : ["operator"];
}

function apiKeyAccess(req, reply, presented, url) {
  const key = resolveApiKey(db, presented);
  if (!key) return reply.code(401).send({ error: "invalid, cancelled or expired API key" });
  if (!keyAllows(key, req.method, url)) {
    return reply.code(403).send({ error: `This API key may not use ${req.method} ${url}. Keys are read-only; see the API documentation for what each scope allows.` });
  }
  if (!underRateLimit(key.id)) return reply.code(429).send({ error: "Too many requests from this API key this minute. Try again in a minute." });
  let requestedSite = String(req.headers["x-site-code"] || "").trim().toLowerCase();
  if (requestedSite && !knownSitesCached(db).includes(requestedSite)) requestedSite = "";
  if (Array.isArray(key.sites) && key.sites.length) {
    const effective = requestedSite || key.sites[0];
    if (!key.sites.includes(effective)) return reply.code(403).send({ error: "site access denied", allowed_locations: key.sites });
    req.headers["x-site-code"] = effective;
    req.ironlogAllowedSites = key.sites;
  } else {
    req.headers["x-site-code"] = requestedSite || defaultSite();
  }
  // Read access as a supervisor would have, limited above to the key's listed addresses.
  req.headers["x-user-name"] = `api:${key.name}`;
  req.headers["x-user-role"] = "supervisor";
  req.headers["x-user-roles"] = "supervisor";
  delete req.headers["x-user-permissions"];
  delete req.headers["x-user-department"];
  req.ironlogApiKey = key;
  try { touchApiKey(db, key, url); } catch {}
}

export async function ironlogAuthHook(req, reply) {
  const url = req.url.split("?")[0];
  if (!url.startsWith("/api/")) return;

  if (isPublicAuthRequest(url, req.method)) return;
  if (url === "/api/notifications/config" && req.method === "GET") return;

  const auth = String(req.headers.authorization || "");
  const token = auth.startsWith("Bearer ") ? auth.slice(7).trim() : "";

  // Another system's API key (utils/apiKeys.js): read-only, its scopes and sites only.
  const presentedKey = token.startsWith(API_KEY_PREFIX) ? token : String(req.headers["x-api-key"] || "").trim();
  if (presentedKey) return apiKeyAccess(req, reply, presentedKey, url);

  if (token) {
    const row = resolveSession(token);
    if (row && Number(row.active) === 1) {
      touchSession(db, token);
      const roles = parseRoles(row.roles_json, row.role);
      const allowedLocations = parseAllowedLocations(row.allowed_locations);
      let requestedSite = String(req.headers["x-site-code"] || "").trim().toLowerCase();
      // A site the company doesn't have (an old value kept by the browser) means the default
      // site, so nobody is left looking at empty lists.
      if (requestedSite && !knownSitesCached(db).includes(requestedSite)) {
        requestedSite = "";
        req.headers["x-site-code"] = defaultSite();
      }
      // Main administrators have global site access even if a stale or narrow
      // location list remains on their account.
      if (!roles.includes("admin") && Array.isArray(allowedLocations) && allowedLocations.length) {
        const effective = requestedSite || allowedLocations[0];
        if (!allowedLocations.includes(effective)) {
          return reply.code(403).send({
            error: "site access denied",
            allowed_locations: allowedLocations,
          });
        }
        req.headers["x-site-code"] = effective;
        // The record guard (utils/siteGuard.js) keeps this person to these sites.
        req.ironlogAllowedSites = allowedLocations;
      }
      const routeRoles = expandRouteRoles(roles, url);
      req.headers["x-user-name"] = row.username;
      req.headers["x-user-role"] = primaryRouteRole(routeRoles) || roles[0];
      req.headers["x-user-roles"] = routeRoles.join(",");
      req.headers["x-user-department"] = String(row.department || "").trim().toLowerCase();
      const perms = getPermissionsForRoles(roles);
      if (perms.length) req.headers["x-user-permissions"] = perms.join(",");
      return;
    }
    if (isAuthRequired()) {
      return reply.code(401).send({ error: "invalid or expired session" });
    }
    applyLegacyRouteRoles(req);
    return;
  }

  if (isAuthRequired()) {
    return reply.code(401).send({ error: "login required" });
  }
  applyLegacyRouteRoles(req);
}
