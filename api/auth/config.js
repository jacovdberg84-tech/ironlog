export function isAuthRequired(env = process.env) {
  const configured = env.IRONLOG_AUTH_REQUIRED;
  if (configured != null && String(configured).trim() !== "") {
    return configured === "1" || String(configured).trim().toLowerCase() === "true";
  }

  if (env.IRONLOG_DESKTOP === "1") return false;
  const host = String(env.HOST || "0.0.0.0").trim().toLowerCase();
  return !new Set(["127.0.0.1", "localhost", "::1"]).has(host);
}

const PUBLIC_AUTH_REQUESTS = new Set([
  "POST /api/auth/login",
  "POST /api/auth/pin/login",
  "POST /api/auth/setup-password",
  "GET /api/auth/pin/roster",
  "GET /api/auth/config",
  "GET /api/auth/tabs",
]);

export function isPublicAuthRequest(url, method) {
  return PUBLIC_AUTH_REQUESTS.has(`${String(method || "GET").toUpperCase()} ${String(url || "")}`);
}
