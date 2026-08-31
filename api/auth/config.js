export function isAuthRequired(env = process.env) {
  const configured = env.IRONLOG_AUTH_REQUIRED;
  if (configured != null && String(configured).trim() !== "") {
    return configured === "1" || String(configured).trim().toLowerCase() === "true";
  }

  if (env.IRONLOG_DESKTOP === "1") return false;
  const host = String(env.HOST || "0.0.0.0").trim().toLowerCase();
  return !new Set(["127.0.0.1", "localhost", "::1"]).has(host);
}
