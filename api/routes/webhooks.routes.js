// IRONLOG/api/routes/webhooks.routes.js — IronLog tells other systems when something happens (utils/webhooks.js).
//   GET    /api/webhooks                               administrators: the webhooks (never the secret) and the events
//   POST   /api/webhooks                               administrators: { name, url, events: [...], sites?: [...] } → the secret, shown once
//   PATCH  /api/webhooks/:id                           administrators: name, url, events, sites, active
//   DELETE /api/webhooks/:id                           administrators
//   POST   /api/webhooks/:id/secret                    administrators: a new secret, shown once
//   POST   /api/webhooks/:id/test                      administrators: sends a "ping" now and says how it went
//   GET    /api/webhooks/:id/deliveries                administrators: recent messages and how they went
//   POST   /api/webhooks/:id/deliveries/:did/retry     administrators: send a failed message again
//   GET    /api/webhooks/catalogue                     anyone signed in: the events (documentation)
// A scan every WEBHOOK_SCAN_SECONDS finds new events and sends what is due. Licensed as "api_access".
import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";
import { getUser, requireAnyRole } from "../utils/request.js";
import { checkAllowedLocations, defaultSite } from "../utils/sites.js";
import {
  WEBHOOK_SCAN_SECONDS, createWebhook, deleteWebhook, ensureWebhookTables, listDeliveries, listWebhooks, queuePing,
  retryDelivery, rotateWebhookSecret, scanWebhookEvents, sendDueDeliveries, updateWebhook, webhookCatalogue,
} from "../utils/webhooks.js";
import { currentLicenceState } from "./licence.routes.js";

/** One pass: new events are queued and due messages sent. Never throws. */
export async function runWebhookPass(log) {
  try {
    const cache = new Map();
    const moduleOn = (site) => {
      const s = site || defaultSite();
      if (!cache.has(s)) {
        try { cache.set(s, (currentLicenceState(s).modules || []).includes("api_access")); } catch { cache.set(s, false); }
      }
      return cache.get(s);
    };
    scanWebhookEvents(db, { moduleOn });
    return await sendDueDeliveries(db);
  } catch (e) {
    log?.warn?.({ err: e }, "webhook pass failed");
    return null;
  }
}

export default async function webhookRoutes(app) {
  ensureWebhookTables(db);
  ensureAuditTable(db);
  const audit = (req, action, id, payload) => {
    try { writeAudit(db, req, { module: "webhooks", action, entity_type: "webhook", entity_id: String(id), payload }); } catch {}
  };
  const checkSites = (list) => checkAllowedLocations(db, list);
  const admin = (req, reply) => requireAnyRole(req, reply, ["admin"]);

  let busy = false;
  const timer = setInterval(async () => {
    if (busy) return;
    busy = true;
    try { await runWebhookPass(app.log); } finally { busy = false; }
  }, WEBHOOK_SCAN_SECONDS * 1000);
  timer.unref?.();
  app.addHook("onClose", async () => clearInterval(timer));

  app.get("/catalogue", async () => ({ ok: true, events: webhookCatalogue() }));

  app.get("/", async (req, reply) => {
    if (!admin(req, reply)) return;
    return { ok: true, webhooks: listWebhooks(db), events: webhookCatalogue() };
  });

  app.post("/", async (req, reply) => {
    if (!admin(req, reply)) return;
    const r = createWebhook(db, req.body || {}, { by: getUser(req), checkSites });
    if (!r.ok) return reply.code(400).send({ ok: false, error: r.errors.join(" "), errors: r.errors });
    audit(req, "webhook.create", r.hook.id, { name: r.hook.name, url: r.hook.url, events: r.hook.events, sites: r.hook.sites });
    return { ok: true, secret: r.secret, webhook: r.hook };
  });

  app.patch("/:id", async (req, reply) => {
    if (!admin(req, reply)) return;
    const r = updateWebhook(db, req.params.id, req.body || {}, { by: getUser(req), checkSites });
    if (!r) return reply.code(404).send({ ok: false, error: "Webhook not found." });
    if (!r.ok) return reply.code(400).send({ ok: false, error: r.errors.join(" "), errors: r.errors });
    audit(req, "webhook.update", req.params.id, { name: r.hook.name, url: r.hook.url, events: r.hook.events, sites: r.hook.sites, active: r.hook.active });
    return { ok: true, webhook: r.hook };
  });

  app.delete("/:id", async (req, reply) => {
    if (!admin(req, reply)) return;
    if (!deleteWebhook(db, req.params.id)) return reply.code(404).send({ ok: false, error: "Webhook not found." });
    audit(req, "webhook.delete", req.params.id, {});
    return { ok: true };
  });

  app.post("/:id/secret", async (req, reply) => {
    if (!admin(req, reply)) return;
    const secret = rotateWebhookSecret(db, req.params.id, getUser(req));
    if (!secret) return reply.code(404).send({ ok: false, error: "Webhook not found." });
    audit(req, "webhook.secret", req.params.id, { secret_prefix: secret.slice(0, 10) });
    return { ok: true, secret };
  });

  app.post("/:id/test", async (req, reply) => {
    if (!admin(req, reply)) return;
    if (!listWebhooks(db).some((h) => h.id === Number(req.params.id))) return reply.code(404).send({ ok: false, error: "Webhook not found." });
    const uid = queuePing(db, req.params.id);
    await sendDueDeliveries(db);
    const d = db.prepare(`SELECT status, last_status_code, last_error FROM webhook_deliveries WHERE uid = ?`).get(uid);
    audit(req, "webhook.test", req.params.id, { status: d?.status, code: d?.last_status_code });
    return { ok: d?.status === "delivered", status: d?.status, status_code: d?.last_status_code ?? null, error: d?.last_error || null };
  });

  app.get("/:id/deliveries", async (req, reply) => {
    if (!admin(req, reply)) return;
    return { ok: true, deliveries: listDeliveries(db, req.params.id, req.query?.limit) };
  });

  app.post("/:id/deliveries/:did/retry", async (req, reply) => {
    if (!admin(req, reply)) return;
    if (!retryDelivery(db, req.params.id, req.params.did)) return reply.code(404).send({ ok: false, error: "Message not found, or already delivered." });
    audit(req, "webhook.retry", req.params.id, { delivery: req.params.did });
    await sendDueDeliveries(db);
    return { ok: true };
  });
}
