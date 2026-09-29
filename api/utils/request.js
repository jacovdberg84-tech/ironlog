// IRONLOG/api/utils/request.js — shared request helpers for route modules.
// The auth hook (auth/hook.js) sets the x-user-* headers from the session, so
// these helpers only read what the hook already resolved.

export function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function splitRoleHeader(value) {
  return String(value || "")
    .split(",")
    .map((x) => String(x || "").trim().toLowerCase())
    .filter(Boolean);
}

/** All roles on the request (x-user-roles plus x-user-role), lowercased and de-duplicated. */
export function getRoles(req, { fallback = null } = {}) {
  const merged = Array.from(
    new Set([...splitRoleHeader(req.headers["x-user-roles"]), ...splitRoleHeader(req.headers["x-user-role"])])
  );
  if (!merged.length && fallback) return [fallback];
  return merged;
}

/** Primary role (x-user-role). Defaults to admin to match legacy local mode. */
export function getRole(req) {
  return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
}

export function getUser(req, fallback = "session-user") {
  return String(req.headers["x-user-name"] || fallback).trim() || fallback;
}

export function getSiteCode(req) {
  return String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
}

export function hasAnyRole(req, allowed, opts) {
  return getRoles(req, opts).some((r) => allowed.includes(r));
}

/** Sends 403 and returns false unless the request holds at least one allowed role. */
export function requireAnyRole(req, reply, allowed, opts) {
  if (!hasAnyRole(req, allowed, opts)) {
    reply.code(403).send({ error: `role '${getRole(req)}' not allowed` });
    return false;
  }
  return true;
}
