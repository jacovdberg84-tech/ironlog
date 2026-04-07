// IRONLOG/api/auth/hook.js — optional Bearer session + IRONLOG_AUTH_REQUIRED
import { db } from "../db/client.js";

export function isAuthRequired() {
  const v = process.env.IRONLOG_AUTH_REQUIRED;
  return v === "1" || String(v).toLowerCase() === "true";
}

function resolveSession(token) {
  try {
    return db
      .prepare(`
        SELECT u.username, u.role, u.roles_json, u.active, u.allowed_locations
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

function parseAllowedLocations(raw) {
  if (!raw) return null;
  try {
    const arr = typeof raw === "string" ? JSON.parse(raw) : raw;
    if (!Array.isArray(arr)) return null;
    const out = arr
      .map((x) => String(x || "").trim().toLowerCase())
      .filter(Boolean);
    return out.length ? Array.from(new Set(out)) : null;
  } catch {
    return null;
  }
}

const VALID_ROLES = ["admin", "supervisor", "stores", "artisan", "operator"];

function parseRoles(raw, fallbackRole = "operator") {
  let arr = [];
  if (raw != null && raw !== "") {
    try {
      const v = typeof raw === "string" ? JSON.parse(raw) : raw;
      if (Array.isArray(v)) arr = v.map((x) => String(x || "").trim().toLowerCase());
    } catch {}
  }
  if (!arr.length && fallbackRole) arr = [String(fallbackRole || "").trim().toLowerCase()];
  const out = Array.from(new Set(arr.filter((r) => VALID_ROLES.includes(r))));
  return out.length ? out : ["operator"];
}

export async function ironlogAuthHook(req, reply) {
  const url = req.url.split("?")[0];
  if (!url.startsWith("/api/")) return;

  if (url === "/api/auth/login" && req.method === "POST") return;
  if (url === "/api/auth/tabs" && req.method === "GET") return;

  const auth = String(req.headers.authorization || "");
  const token = auth.startsWith("Bearer ") ? auth.slice(7).trim() : "";

  if (token) {
    const row = resolveSession(token);
    if (row && Number(row.active) === 1) {
      const roles = parseRoles(row.roles_json, row.role);
      const allowedLocations = parseAllowedLocations(row.allowed_locations);
      const requestedSite = String(req.headers["x-site-code"] || "").trim().toLowerCase();
      if (Array.isArray(allowedLocations) && allowedLocations.length) {
        const effective = requestedSite || allowedLocations[0];
        if (!allowedLocations.includes(effective)) {
          return reply.code(403).send({
            error: "site access denied",
            allowed_locations: allowedLocations,
          });
        }
        req.headers["x-site-code"] = effective;
      }
      req.headers["x-user-name"] = row.username;
      req.headers["x-user-role"] = roles[0];
      req.headers["x-user-roles"] = roles.join(",");
      return;
    }
    if (isAuthRequired()) {
      return reply.code(401).send({ error: "invalid or expired session" });
    }
    return;
  }

  if (isAuthRequired()) {
    return reply.code(401).send({ error: "login required" });
  }
}
