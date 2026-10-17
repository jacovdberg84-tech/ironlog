// IRONLOG/api/routes/apiKeys.routes.js — keys for other systems (utils/apiKeys.js).
//   GET  /api/api-keys                 administrators: the keys (never the secret)
//   POST /api/api-keys                 administrators: { name, scopes: [...], sites?: [...], expires_on? } → the key, shown once
//   POST /api/api-keys/:id/cancel      administrators
//   GET  /api/api-keys/catalogue       anyone signed in: what each scope may read (the documentation page)
//   GET  /api/api-keys/openapi.json    anyone signed in: the same as an OpenAPI 3 file for developers
// Licensed as the "api_access" module (utils/licensing.js).
import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";
import { getUser, requireAnyRole } from "../utils/request.js";
import { checkAllowedLocations } from "../utils/sites.js";
import { API_KEY_RATE_PER_MIN, apiCatalogue, createApiKey, ensureApiKeyTables, listApiKeys, openApiDocument, revokeApiKey } from "../utils/apiKeys.js";
import { publicBase } from "./microsoftSignin.routes.js";

export default async function apiKeyRoutes(app) {
  ensureApiKeyTables(db);
  ensureAuditTable(db);
  const audit = (req, action, id, payload) => {
    try {
      writeAudit(db, req, { module: "api_keys", action, entity_type: "api_key", entity_id: String(id), payload });
    } catch {}
  };

  app.get("/", async (req, reply) => {
    if (!requireAnyRole(req, reply, ["admin"])) return;
    return { ok: true, keys: listApiKeys(db), rate_per_minute: API_KEY_RATE_PER_MIN };
  });

  app.post("/", async (req, reply) => {
    if (!requireAnyRole(req, reply, ["admin"])) return;
    const r = createApiKey(db, req.body || {}, { by: getUser(req), checkSites: (list) => checkAllowedLocations(db, list) });
    if (!r.ok) return reply.code(400).send({ ok: false, error: r.errors.join(" "), errors: r.errors });
    audit(req, "api_key.create", r.row.id, { name: r.row.name, scopes: r.row.scopes, sites: r.row.sites, expires_on: r.row.expires_on, key_prefix: r.row.key_prefix });
    return { ok: true, key: r.key, row: r.row };
  });

  app.post("/:id/cancel", async (req, reply) => {
    if (!requireAnyRole(req, reply, ["admin"])) return;
    if (!revokeApiKey(db, req.params.id, getUser(req))) return reply.code(404).send({ ok: false, error: "Key not found or already cancelled." });
    audit(req, "api_key.cancel", req.params.id, {});
    return { ok: true };
  });

  app.get("/catalogue", async () => ({ ok: true, rate_per_minute: API_KEY_RATE_PER_MIN, scopes: apiCatalogue() }));

  app.get("/openapi.json", async (req) => openApiDocument(publicBase(req)));
}
