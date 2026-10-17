// IRONLOG/api/utils/webhooks.js — IronLog tells other systems when something happens (webhooks).
//
// An administrator adds a webhook: a name, the https address of the other system, which
// events it wants and, optionally, which sites. IronLog then POSTs a small JSON message to
// that address for each event, signed with the webhook's secret (shown once, stored
// encrypted) so the other system can tell it really came from IronLog:
//
//   X-IronLog-Signature: t=<unix seconds>,v1=<hex HMAC-SHA256 of "<t>.<body>" with the secret>
//
// Events are found by a scan every WEBHOOK_SCAN_SECONDS rather than inside each screen's
// code, so a breakdown logged from the portal, the API or an import is caught the same way.
// The scan remembers what it has seen (webhook_seen); records that existed before anyone
// subscribed are never sent. A failed delivery is retried on RETRY_MINUTES, then marked failed.
// Addresses must be https and public; private and local addresses are refused (set
// IRONLOG_WEBHOOK_ALLOW_PRIVATE=1 for an on-site system or for testing). Redirects are not followed.
// Webhooks are part of the licensed "api_access" module.
import crypto from "node:crypto";
import dns from "node:dns/promises";
import net from "node:net";
import { decryptJuneConnectorSecret, encryptJuneConnectorSecret } from "./juneOutlook.js";

export const WEBHOOK_SCAN_SECONDS = 30;
export const RETRY_MINUTES = [1, 5, 30, 120, 360, 720];
const SEND_TIMEOUT_MS = 10000;

export const WEBHOOK_EVENTS = {
  "breakdown.created": { label: "Breakdown logged", text: "A machine breakdown is logged (any screen, the technician portal or the API)." },
  "breakdown.closed": { label: "Breakdown closed", text: "A breakdown is closed: the machine is back, with its downtime hours and failure codes." },
  "workorder.created": { label: "Work order opened", text: "A work order is opened." },
  "workorder.completed": { label: "Work order completed", text: "A work order is completed, approved or closed, with its labour hours." },
  "stock.below_minimum": { label: "Store item below minimum", text: "A store item falls below its minimum stock (sent again only after it has been restocked)." },
};

export function ensureWebhookTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS webhooks (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      url TEXT NOT NULL,
      events_json TEXT NOT NULL,
      sites_json TEXT,
      secret_enc TEXT NOT NULL,
      secret_prefix TEXT NOT NULL,
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      created_by TEXT,
      updated_at TEXT,
      updated_by TEXT,
      last_success_at TEXT,
      last_failure_at TEXT,
      failures_in_row INTEGER NOT NULL DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS webhook_deliveries (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      uid TEXT NOT NULL UNIQUE,
      webhook_id INTEGER NOT NULL,
      event TEXT NOT NULL,
      record_key TEXT,
      payload_json TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT 'pending',
      attempts INTEGER NOT NULL DEFAULT 0,
      next_attempt_at TEXT NOT NULL DEFAULT (datetime('now')),
      last_status_code INTEGER,
      last_error TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      delivered_at TEXT,
      FOREIGN KEY (webhook_id) REFERENCES webhooks(id) ON DELETE CASCADE
    );
    CREATE INDEX IF NOT EXISTS idx_webhook_deliveries_due ON webhook_deliveries(status, next_attempt_at);
    CREATE INDEX IF NOT EXISTS idx_webhook_deliveries_hook ON webhook_deliveries(webhook_id, id);
    CREATE TABLE IF NOT EXISTS webhook_seen (
      event TEXT NOT NULL,
      record_key TEXT NOT NULL,
      seen_at TEXT NOT NULL DEFAULT (datetime('now')),
      PRIMARY KEY (event, record_key)
    );
  `);
}

const parse = (raw) => {
  try { const v = JSON.parse(raw || "null"); return Array.isArray(v) ? v : null; } catch { return null; }
};

function publicHook(r) {
  return {
    id: r.id, name: r.name, url: r.url, events: parse(r.events_json) || [], sites: parse(r.sites_json), secret_prefix: r.secret_prefix,
    active: Boolean(r.active), created_at: r.created_at, created_by: r.created_by, last_success_at: r.last_success_at,
    last_failure_at: r.last_failure_at, failures_in_row: Number(r.failures_in_row || 0),
  };
}

/** Private, loopback, link-local and other non-public addresses. */
export function isPrivateAddress(ip) {
  const v = net.isIP(ip);
  if (v === 4) {
    const [a, b] = ip.split(".").map(Number);
    return a === 10 || a === 127 || a === 0 || (a === 169 && b === 254) || (a === 172 && b >= 16 && b <= 31) || (a === 192 && b === 168)
      || (a === 100 && b >= 64 && b <= 127) || a >= 224;
  }
  if (v === 6) {
    const s = ip.toLowerCase();
    if (s.startsWith("::ffff:")) return isPrivateAddress(s.slice(7));
    return s === "::1" || s === "::" || s.startsWith("fc") || s.startsWith("fd") || s.startsWith("fe8") || s.startsWith("fe9") || s.startsWith("fea") || s.startsWith("feb") || s.startsWith("ff");
  }
  return true;
}

const allowPrivate = (env = process.env) => String(env.IRONLOG_WEBHOOK_ALLOW_PRIVATE || "") === "1";

/** The address as typed, checked: { ok, url } or { ok:false, error }. DNS is checked when sending. */
export function checkWebhookUrl(raw, env = process.env) {
  let u;
  try { u = new URL(String(raw || "").trim()); } catch { return { ok: false, error: "Type the full address, starting with https://" }; }
  if (u.username || u.password) return { ok: false, error: "Leave the user name and password out of the address." };
  const priv = allowPrivate(env);
  if (u.protocol !== "https:" && !(priv && u.protocol === "http:")) return { ok: false, error: "The address must start with https://" };
  const host = u.hostname.replace(/^\[|\]$/g, "");
  if (!priv && (host === "localhost" || host.endsWith(".localhost") || host.endsWith(".local") || (net.isIP(host) && isPrivateAddress(host)))) {
    return { ok: false, error: "That is a private or local address. Webhooks go to systems on the internet." };
  }
  if (String(u).length > 500) return { ok: false, error: "The address is too long." };
  return { ok: true, url: u.toString() };
}

function readHookBody(body, { checkSites, requireAll }) {
  const errors = [];
  const out = {};
  if (requireAll || body.name !== undefined) {
    const name = String(body.name || "").trim();
    if (!name || name.length > 60) errors.push("Give the webhook a name of up to 60 characters, e.g. the system it goes to.");
    else out.name = name;
  }
  if (requireAll || body.url !== undefined) {
    const u = checkWebhookUrl(body.url);
    if (!u.ok) errors.push(u.error);
    else out.url = u.url;
  }
  if (requireAll || body.events !== undefined) {
    const events = [...new Set((Array.isArray(body.events) ? body.events : []).map((e) => String(e).trim()))];
    const bad = events.filter((e) => !WEBHOOK_EVENTS[e]);
    if (!events.length || bad.length) errors.push(`Choose which events to send${bad.length ? ` (unknown: ${bad.join(", ")})` : ""}.`);
    else out.events_json = JSON.stringify(events);
  }
  if (requireAll || body.sites !== undefined) {
    const list = (Array.isArray(body.sites) ? body.sites : String(body.sites || "").split(",")).map((s) => String(s).trim().toLowerCase()).filter(Boolean);
    if (!list.length) out.sites_json = null;
    else {
      const r = checkSites(list);
      if (!r.ok) errors.push(r.error);
      else out.sites_json = JSON.stringify(r.sites);
    }
  }
  return { errors, out };
}

export function listWebhooks(db) {
  ensureWebhookTables(db);
  return db.prepare(`SELECT * FROM webhooks ORDER BY active DESC, id DESC`).all().map(publicHook);
}

/** { ok, secret (shown once), hook } or { ok:false, errors }. */
export function createWebhook(db, body = {}, { by = "", checkSites = (l) => ({ ok: true, sites: l }) } = {}) {
  ensureWebhookTables(db);
  const { errors, out } = readHookBody(body, { checkSites, requireAll: true });
  if (errors.length) return { ok: false, errors };
  const secret = `whsec_${crypto.randomBytes(24).toString("base64url")}`;
  const id = db.prepare(`INSERT INTO webhooks (name, url, events_json, sites_json, secret_enc, secret_prefix, created_by) VALUES (?, ?, ?, ?, ?, ?, ?)`)
    .run(out.name, out.url, out.events_json, out.sites_json, encryptJuneConnectorSecret(secret), secret.slice(0, 10), by || null).lastInsertRowid;
  return { ok: true, secret, hook: publicHook(db.prepare(`SELECT * FROM webhooks WHERE id = ?`).get(id)) };
}

/** Changes name, address, events, sites or active; { ok, hook } or { ok:false, errors } / null when not found. */
export function updateWebhook(db, id, body = {}, { by = "", checkSites = (l) => ({ ok: true, sites: l }) } = {}) {
  ensureWebhookTables(db);
  if (!db.prepare(`SELECT id FROM webhooks WHERE id = ?`).get(Number(id))) return null;
  const { errors, out } = readHookBody(body, { checkSites, requireAll: false });
  if (body.active !== undefined) out.active = body.active ? 1 : 0;
  if (errors.length) return { ok: false, errors };
  const cols = Object.keys(out);
  if (cols.length) {
    db.prepare(`UPDATE webhooks SET ${cols.map((c) => `${c} = ?`).join(", ")}, updated_at = datetime('now'), updated_by = ? WHERE id = ?`).run(...cols.map((c) => out[c]), by || null, Number(id));
    if (out.active === 1) db.prepare(`UPDATE webhooks SET failures_in_row = 0 WHERE id = ?`).run(Number(id));
  }
  return { ok: true, hook: publicHook(db.prepare(`SELECT * FROM webhooks WHERE id = ?`).get(Number(id))) };
}

/** A new secret for a webhook (the old one stops working at once); returns it, shown once. */
export function rotateWebhookSecret(db, id, by = "") {
  ensureWebhookTables(db);
  const secret = `whsec_${crypto.randomBytes(24).toString("base64url")}`;
  const r = db.prepare(`UPDATE webhooks SET secret_enc = ?, secret_prefix = ?, updated_at = datetime('now'), updated_by = ? WHERE id = ?`)
    .run(encryptJuneConnectorSecret(secret), secret.slice(0, 10), by || null, Number(id));
  return r.changes ? secret : null;
}

export function deleteWebhook(db, id) {
  ensureWebhookTables(db);
  db.prepare(`DELETE FROM webhook_deliveries WHERE webhook_id = ?`).run(Number(id));
  return db.prepare(`DELETE FROM webhooks WHERE id = ?`).run(Number(id)).changes > 0;
}

export function listDeliveries(db, webhookId, limit = 25) {
  ensureWebhookTables(db);
  return db.prepare(`
    SELECT id, uid, event, record_key, status, attempts, next_attempt_at, last_status_code, last_error, created_at, delivered_at
    FROM webhook_deliveries WHERE webhook_id = ? ORDER BY id DESC LIMIT ?
  `).all(Number(webhookId), Math.min(100, Math.max(1, Number(limit) || 25)));
}

/** Sends a failed or waiting delivery again now. */
export function retryDelivery(db, webhookId, deliveryId) {
  return db.prepare(`UPDATE webhook_deliveries SET status = 'pending', next_attempt_at = datetime('now') WHERE id = ? AND webhook_id = ? AND status != 'delivered'`)
    .run(Number(deliveryId), Number(webhookId)).changes > 0;
}

/** The signature header value for a body. */
export function signWebhook(secret, body, t = Math.floor(Date.now() / 1000)) {
  return `t=${t},v1=${crypto.createHmac("sha256", secret).update(`${t}.${body}`).digest("hex")}`;
}

const hasTable = (db, t) => Boolean(db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?`).get(t));
const cols = (db, t) => new Set(db.prepare(`PRAGMA table_info(${t})`).all().map((c) => c.name));

// What each event looks for: rows not yet seen, as { key, site, data }.
function eventSources(db) {
  const a = hasTable(db, "assets") ? cols(db, "assets") : new Set();
  const b = hasTable(db, "breakdowns") ? cols(db, "breakdowns") : null;
  const w = hasTable(db, "work_orders") ? cols(db, "work_orders") : null;
  const assetSite = a.has("site_code") ? "a.site_code" : "NULL";
  const pick = (set, alias, list) => list.map((c) => (set.has(c) ? `${alias}.${c}` : `NULL AS ${c}`)).join(", ");
  const unseen = (event, alias) => `NOT EXISTS (SELECT 1 FROM webhook_seen s WHERE s.event = '${event}' AND s.record_key = CAST(${alias}.id AS TEXT))`;
  const src = {};
  if (b) {
    const bSite = b.has("site_code") ? `COALESCE(NULLIF(TRIM(b.site_code), ''), ${assetSite})` : assetSite;
    const fields = pick(b, "b", ["breakdown_date", "status", "start_at", "end_at", "description", "critical", "downtime_total_hours", "primary_work_order_id", "fc_component", "fc_problem", "fc_cause", "fc_action"]);
    const base = (event, where) => () => db.prepare(`
      SELECT b.id, ${fields}, a.asset_code, a.asset_name, a.category, ${bSite} AS site_code
      FROM breakdowns b LEFT JOIN assets a ON a.id = b.asset_id
      WHERE ${where} ${unseen(event, "b")} ORDER BY b.id LIMIT 500
    `).all().map((r) => ({
      key: String(r.id), site: r.site_code ? String(r.site_code).toLowerCase() : null,
      data: {
        breakdown_id: r.id, asset_code: r.asset_code, asset_name: r.asset_name, category: r.category, breakdown_date: r.breakdown_date,
        status: r.status, description: r.description, critical: Boolean(r.critical), start_at: r.start_at, end_at: r.end_at,
        downtime_hours: r.downtime_total_hours == null ? null : Number(r.downtime_total_hours), work_order_id: r.primary_work_order_id,
        failure_codes: { component: r.fc_component, problem: r.fc_problem, cause: r.fc_cause, action: r.fc_action },
      },
    }));
    src["breakdown.created"] = base("breakdown.created", "1 = 1 AND");
    src["breakdown.closed"] = base("breakdown.closed", "UPPER(COALESCE(b.status, '')) = 'CLOSED' AND");
  }
  if (w) {
    const wSite = w.has("site_code") ? `COALESCE(NULLIF(TRIM(wo.site_code), ''), ${assetSite})` : assetSite;
    const fields = pick(w, "wo", ["source", "reference_id", "status", "opened_at", "completed_at", "closed_at", "labor_hours", "assigned_to", "title", "description"]);
    const base = (event, where) => () => db.prepare(`
      SELECT wo.id, ${fields}, a.asset_code, a.asset_name, ${wSite} AS site_code
      FROM work_orders wo LEFT JOIN assets a ON a.id = wo.asset_id
      WHERE ${where} ${unseen(event, "wo")} ORDER BY wo.id LIMIT 500
    `).all().map((r) => ({
      key: String(r.id), site: r.site_code ? String(r.site_code).toLowerCase() : null,
      data: {
        work_order_id: r.id, asset_code: r.asset_code, asset_name: r.asset_name, source: r.source, reference_id: r.reference_id, status: r.status,
        title: r.title || r.description || null, assigned_to: r.assigned_to, opened_at: r.opened_at, completed_at: r.completed_at, closed_at: r.closed_at,
        labour_hours: r.labor_hours == null ? null : Number(r.labor_hours),
      },
    }));
    src["workorder.created"] = base("workorder.created", "1 = 1 AND");
    src["workorder.completed"] = base("workorder.completed", "LOWER(COALESCE(wo.status, '')) IN ('completed', 'approved', 'closed') AND");
  }
  if (hasTable(db, "parts") && hasTable(db, "stock_movements")) {
    const p = cols(db, "parts");
    src["stock.below_minimum"] = () => db.prepare(`
      SELECT p.id, p.part_code, p.part_name, p.min_stock, ${p.has("unit_cost") ? "p.unit_cost" : "NULL AS unit_cost"},
             (SELECT COALESCE(SUM(m.quantity), 0) FROM stock_movements m WHERE m.part_id = p.id) AS on_hand
      FROM parts p
      WHERE COALESCE(p.min_stock, 0) > 0 ${unseen("stock.below_minimum", "p").replace(/^/, "AND ")}
        AND (SELECT COALESCE(SUM(m.quantity), 0) FROM stock_movements m WHERE m.part_id = p.id) < p.min_stock
      ORDER BY p.id LIMIT 500
    `).all().map((r) => ({
      key: String(r.id), site: null,
      data: { part_code: r.part_code, part_name: r.part_name, on_hand: Number(r.on_hand), min_stock: Number(r.min_stock), unit_cost_usd: r.unit_cost == null ? null : Number(r.unit_cost) },
    }));
  }
  return src;
}

/**
 * One scan: finds new events, queues a delivery for each subscribed webhook (by event and site)
 * and remembers the records. Events nobody subscribes to are remembered without sending.
 * moduleOn(site) says whether the licence allows webhooks. Returns how many deliveries were queued.
 */
export function scanWebhookEvents(db, { moduleOn = () => true, now = new Date() } = {}) {
  ensureWebhookTables(db);
  const hooks = db.prepare(`SELECT * FROM webhooks WHERE active = 1`).all().map((h) => ({ ...h, events: parse(h.events_json) || [], sites: parse(h.sites_json) }));
  const seen = db.prepare(`INSERT OR IGNORE INTO webhook_seen (event, record_key) VALUES (?, ?)`);
  const queue = db.prepare(`INSERT INTO webhook_deliveries (uid, webhook_id, event, record_key, payload_json) VALUES (?, ?, ?, ?, ?)`);
  let queued = 0;
  // A part back at or above its minimum may warn again next time it drops.
  if (hasTable(db, "parts") && hasTable(db, "stock_movements")) {
    db.prepare(`
      DELETE FROM webhook_seen WHERE event = 'stock.below_minimum' AND record_key IN (
        SELECT CAST(p.id AS TEXT) FROM parts p
        WHERE (SELECT COALESCE(SUM(m.quantity), 0) FROM stock_movements m WHERE m.part_id = p.id) >= COALESCE(p.min_stock, 0))
    `).run();
  }
  for (const [event, find] of Object.entries(eventSources(db))) {
    const rows = find();
    if (!rows.length) continue;
    const subs = hooks.filter((h) => h.events.includes(event));
    db.transaction(() => {
      for (const r of rows) {
        seen.run(event, r.key);
        for (const h of subs) {
          if (h.sites && r.site && !h.sites.includes(r.site)) continue;
          if (!moduleOn(r.site)) continue;
          const uid = crypto.randomUUID();
          queue.run(uid, h.id, event, r.key, JSON.stringify({ id: uid, event, occurred_at: now.toISOString(), site_code: r.site, data: r.data }));
          queued += 1;
        }
      }
    })();
  }
  return queued;
}

/** Queues a "ping" to check the address and signature. Returns the delivery uid. */
export function queuePing(db, webhookId, now = new Date()) {
  const uid = crypto.randomUUID();
  db.prepare(`INSERT INTO webhook_deliveries (uid, webhook_id, event, record_key, payload_json) VALUES (?, ?, 'ping', NULL, ?)`)
    .run(uid, Number(webhookId), JSON.stringify({ id: uid, event: "ping", occurred_at: now.toISOString(), site_code: null, data: { message: "Test message from IronLog. If you can read this, the webhook works." } }));
  return uid;
}

async function addressAllowed(url, env) {
  if (allowPrivate(env)) return null;
  const host = new URL(url).hostname.replace(/^\[|\]$/g, "");
  let addrs;
  try {
    addrs = net.isIP(host) ? [{ address: host }] : await dns.lookup(host, { all: true });
  } catch {
    return "The address could not be found (DNS).";
  }
  if (!addrs.length || addrs.some((x) => isPrivateAddress(x.address))) return "The address points to a private or local network.";
  return null;
}

/** Sends the deliveries that are due (oldest first). Returns { sent, failed }. */
export async function sendDueDeliveries(db, { limit = 50, env = process.env, fetchImpl = globalThis.fetch } = {}) {
  ensureWebhookTables(db);
  const due = db.prepare(`
    SELECT d.*, w.url, w.secret_enc, w.active FROM webhook_deliveries d JOIN webhooks w ON w.id = d.webhook_id
    WHERE d.status = 'pending' AND d.next_attempt_at <= datetime('now') ORDER BY d.id LIMIT ?
  `).all(limit);
  let sent = 0;
  let failed = 0;
  for (const d of due) {
    if (!d.active && d.event !== "ping") continue;
    const secret = decryptJuneConnectorSecret(d.secret_enc) || "";
    const body = d.payload_json;
    let code = null;
    let error = await addressAllowed(d.url, env);
    if (!error) {
      try {
        const res = await fetchImpl(d.url, {
          method: "POST",
          redirect: "manual",
          signal: AbortSignal.timeout(SEND_TIMEOUT_MS),
          headers: {
            "Content-Type": "application/json",
            "User-Agent": "IronLog-Webhooks/1",
            "X-IronLog-Event": d.event,
            "X-IronLog-Delivery": d.uid,
            "X-IronLog-Signature": signWebhook(secret, body),
          },
          body,
        });
        code = res.status;
        if (res.status < 200 || res.status >= 300) error = `The system answered ${res.status}${res.status >= 300 && res.status < 400 ? " (a redirect; use the final address)" : ""}.`;
        try { await res.body?.cancel?.(); } catch { /* nothing to read */ }
      } catch (e) {
        error = e?.name === "TimeoutError" ? "No answer within 10 seconds." : `Could not connect: ${String(e?.cause?.code || e?.message || e).slice(0, 120)}`;
      }
    }
    const attempts = Number(d.attempts) + 1;
    if (!error) {
      db.prepare(`UPDATE webhook_deliveries SET status = 'delivered', attempts = ?, last_status_code = ?, last_error = NULL, delivered_at = datetime('now') WHERE id = ?`).run(attempts, code, d.id);
      db.prepare(`UPDATE webhooks SET last_success_at = datetime('now'), failures_in_row = 0 WHERE id = ?`).run(d.webhook_id);
      sent += 1;
    } else {
      const wait = RETRY_MINUTES[attempts - 1];
      const final = d.event === "ping" || wait == null;
      db.prepare(`UPDATE webhook_deliveries SET status = ?, attempts = ?, last_status_code = ?, last_error = ?, next_attempt_at = datetime('now', ?) WHERE id = ?`)
        .run(final ? "failed" : "pending", attempts, code, error, `+${wait || 0} minutes`, d.id);
      db.prepare(`UPDATE webhooks SET last_failure_at = datetime('now'), failures_in_row = failures_in_row + 1 WHERE id = ?`).run(d.webhook_id);
      failed += 1;
    }
  }
  // Keep 60 days of delivered messages.
  db.prepare(`DELETE FROM webhook_deliveries WHERE status = 'delivered' AND delivered_at < datetime('now', '-60 days')`).run();
  return { sent, failed };
}

/** The event list for the admin screen and the documentation. */
export function webhookCatalogue() {
  return Object.entries(WEBHOOK_EVENTS).map(([event, e]) => ({ event, label: e.label, description: e.text }));
}
