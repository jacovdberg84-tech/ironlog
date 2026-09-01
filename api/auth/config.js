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
  const normalizedMethod = String(method || "GET").toUpperCase();
  const normalizedUrl = String(url || "");
  const requestKey = `${normalizedMethod} ${normalizedUrl}`;
  if (PUBLIC_AUTH_REQUESTS.has(requestKey)) return true;

  // Field operators open these forms by scanning an asset QR code. Keep only
  // capture endpoints public; maintenance administration stays protected.
  const publicPrestartRequests = new Set([
    "GET /api/maintenance/machine-prestart/context",
    "POST /api/maintenance/machine-prestart",
    "GET /api/maintenance/vehicle-ldv-checks/prestart-context",
    "POST /api/maintenance/vehicle-ldv-checks/prestart",
  ]);
  if (publicPrestartRequests.has(requestKey)) return true;
  if (
    normalizedMethod === "POST" &&
    /^\/api\/maintenance\/vehicle-ldv-checks\/\d+\/photo$/.test(normalizedUrl)
  ) return true;
  // Older printed asset QR labels land on the hub first; its read-only profile
  // is needed to route the operator to the correct LDV or machine checklist.
  if (
    normalizedMethod === "GET" &&
    /^\/api\/assets\/[^/]+\/qr-profile$/.test(normalizedUrl)
  ) return true;
  return false;
}
