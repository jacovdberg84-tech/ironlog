// IRONLOG/api/routes/sync.routes.js
import { db } from "../db/client.js";
import crypto from "node:crypto";

const SYNC_TABLES = new Set([
  "daily_hours",
  "fuel_logs",
  "oil_logs",
  "breakdowns",
  "breakdown_downtime_logs",
  "work_orders",
  "manager_inspections",
  "manager_inspection_photos",
]);
const SYNC_SCHEMA_VERSION = 1;

function getRole(req) {
  return String(req.headers["x-user-role"] || "admin").trim().toLowerCase();
}

function getSiteCode(req) {
  return String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
}

function getPeerName(req) {
  return String(req.query?.peer || req.body?.peer || "").trim().toLowerCase();
}

function getSchemaVersion(req) {
  const raw = Number(req.query?.schema_version ?? req.body?.schema_version ?? SYNC_SCHEMA_VERSION);
  if (!Number.isFinite(raw)) return SYNC_SCHEMA_VERSION;
  return Math.trunc(raw);
}

function isDryRun(req) {
  const raw = req.query?.dry_run ?? req.body?.dry_run ?? 0;
  return String(raw).trim() === "1" || raw === true;
}

function makeEventKey(ev) {
  const src = JSON.stringify({
    id: Number(ev?.id || 0),
    table: String(ev?.table || ""),
    op: String(ev?.op || ""),
    row_uuid: String(ev?.row_uuid || ev?.payload?.uuid || ""),
    changed_at: String(ev?.changed_at || ""),
    payload: ev?.payload ?? null,
  });
  return crypto.createHash("sha256").update(src).digest("hex");
}

function requireSyncAdmin(req, reply) {
  const role = getRole(req);
  if (!["admin", "supervisor"].includes(role)) {
    reply.code(403).send({ error: `role '${role || "unknown"}' not allowed` });
    return false;
  }
  return true;
}

function loadRowPayload(tableName, rowId) {
  if (!SYNC_TABLES.has(tableName)) return null;
  const sql = `SELECT * FROM ${tableName} WHERE id = ? LIMIT 1`;
  const row = db.prepare(sql).get(rowId);
  return row || null;
}

function tableColumns(tableName) {
  if (!SYNC_TABLES.has(tableName)) return [];
  return db
    .prepare(`PRAGMA table_info(${tableName})`)
    .all()
    .map((c) => String(c.name || ""))
    .filter(Boolean);
}

function toMillis(ts) {
  if (!ts) return 0;
  const t = Date.parse(String(ts));
  return Number.isFinite(t) ? t : 0;
}

function applyUpsertEvent(ev) {
  const table = String(ev?.table || "").trim();
  if (!SYNC_TABLES.has(table)) return { status: "error", reason: "invalid_table" };
  const payload = ev?.payload && typeof ev.payload === "object" ? ev.payload : null;
  if (!payload) return { status: "error", reason: "payload_required" };
  const rowUuid = String(ev?.row_uuid || payload.uuid || "").trim();
  if (!rowUuid) return { status: "error", reason: "row_uuid_required" };

  const cols = tableColumns(table);
  if (!cols.length) return { status: "error", reason: "table_not_available" };
  if (!cols.includes("uuid")) return { status: "error", reason: "uuid_column_missing" };
  const hasUpdatedAt = cols.includes("updated_at");

  const existing = db.prepare(`SELECT id, uuid, updated_at FROM ${table} WHERE uuid = ? LIMIT 1`).get(rowUuid);
  const incomingUpdatedAt = hasUpdatedAt ? String(payload.updated_at || ev.changed_at || new Date().toISOString()) : null;

  if (existing && hasUpdatedAt) {
    const existingMs = toMillis(existing.updated_at);
    const incomingMs = toMillis(incomingUpdatedAt);
    if (incomingMs > 0 && existingMs > 0 && incomingMs <= existingMs) {
      return { status: "skipped", reason: "stale_event", row_uuid: rowUuid };
    }
  }

  const allowed = cols.filter((c) => c !== "id");
  const data = {};
  for (const c of allowed) {
    if (Object.prototype.hasOwnProperty.call(payload, c)) data[c] = payload[c];
  }
  data.uuid = rowUuid;
  if (cols.includes("site_code") && !String(data.site_code || "").trim()) data.site_code = "main";
  if (hasUpdatedAt) data.updated_at = incomingUpdatedAt;

  if (existing) {
    const setCols = Object.keys(data).filter((k) => k !== "uuid");
    if (!setCols.length) return { status: "applied", mode: "update", row_uuid: rowUuid, row_id: Number(existing.id) };
    const sql = `UPDATE ${table} SET ${setCols.map((c) => `${c} = ?`).join(", ")} WHERE uuid = ?`;
    const vals = setCols.map((c) => data[c]);
    db.prepare(sql).run(...vals, rowUuid);
    const after = db.prepare(`SELECT id FROM ${table} WHERE uuid = ? LIMIT 1`).get(rowUuid);
    return { status: "applied", mode: "update", row_uuid: rowUuid, row_id: Number(after?.id || existing.id || 0) };
  }

  const insertCols = Object.keys(data);
  const sql = `INSERT INTO ${table} (${insertCols.join(", ")}) VALUES (${insertCols.map(() => "?").join(", ")})`;
  const vals = insertCols.map((c) => data[c]);
  const res = db.prepare(sql).run(...vals);
  return { status: "applied", mode: "insert", row_uuid: rowUuid, row_id: Number(res.lastInsertRowid || 0) };
}

function applyDeleteEvent(ev) {
  const table = String(ev?.table || "").trim();
  if (!SYNC_TABLES.has(table)) return { status: "error", reason: "invalid_table" };
  const rowUuid = String(ev?.row_uuid || "").trim();
  if (!rowUuid) return { status: "error", reason: "row_uuid_required" };
  const res = db.prepare(`DELETE FROM ${table} WHERE uuid = ?`).run(rowUuid);
  return { status: "applied", mode: "delete", row_uuid: rowUuid, deleted: Number(res.changes || 0) };
}

export default async function syncRoutes(app) {
  // GET /api/sync/outbox?limit=500
  app.get("/outbox", async (req, reply) => {
    if (!requireSyncAdmin(req, reply)) return;
    const limitRaw = Number(req.query?.limit ?? 500);
    const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(5000, Math.trunc(limitRaw))) : 500;
    const rows = db
      .prepare(`
        SELECT
          id, table_name, row_id, row_uuid, op, site_code,
          payload_json, changed_at, attempts, last_attempt_at, synced_at, error_text
        FROM sync_outbox
        WHERE synced_at IS NULL
        ORDER BY id ASC
        LIMIT ?
      `)
      .all(limit);
    return { ok: true, count: rows.length, rows };
  });

  // POST /api/sync/outbox/ack { ids: [1,2,3] }
  app.post("/outbox/ack", async (req, reply) => {
    if (!requireSyncAdmin(req, reply)) return;
    const idsRaw = Array.isArray(req.body?.ids) ? req.body.ids : [];
    const ids = idsRaw.map((n) => Number(n)).filter((n) => Number.isInteger(n) && n > 0);
    if (!ids.length) return reply.code(400).send({ error: "ids[] required" });
    const marks = ids.map(() => "?").join(",");
    const res = db
      .prepare(`UPDATE sync_outbox SET synced_at = datetime('now'), error_text = NULL WHERE id IN (${marks})`)
      .run(...ids);
    return { ok: true, acknowledged: Number(res.changes || 0) };
  });

  // POST /api/sync/outbox/fail { ids: [1,2], error: "..." }
  app.post("/outbox/fail", async (req, reply) => {
    if (!requireSyncAdmin(req, reply)) return;
    const idsRaw = Array.isArray(req.body?.ids) ? req.body.ids : [];
    const ids = idsRaw.map((n) => Number(n)).filter((n) => Number.isInteger(n) && n > 0);
    const errText = String(req.body?.error || "sync_failed").slice(0, 1000);
    if (!ids.length) return reply.code(400).send({ error: "ids[] required" });
    const marks = ids.map(() => "?").join(",");
    const res = db
      .prepare(`
        UPDATE sync_outbox
        SET
          attempts = COALESCE(attempts, 0) + 1,
          last_attempt_at = datetime('now'),
          error_text = ?
        WHERE id IN (${marks})
      `)
      .run(errText, ...ids);
    return { ok: true, marked_failed: Number(res.changes || 0) };
  });

  // GET /api/sync/stats
  app.get("/stats", async (req, reply) => {
    if (!requireSyncAdmin(req, reply)) return;
    const unsynced = db.prepare(`SELECT COUNT(*) AS c FROM sync_outbox WHERE synced_at IS NULL`).get();
    const failed = db
      .prepare(`SELECT COUNT(*) AS c FROM sync_outbox WHERE synced_at IS NULL AND COALESCE(error_text, '') <> ''`)
      .get();
    const latest = db.prepare(`SELECT MAX(id) AS max_id FROM sync_outbox`).get();
    return {
      ok: true,
      schema_version: SYNC_SCHEMA_VERSION,
      unsynced: Number(unsynced?.c || 0),
      failed: Number(failed?.c || 0),
      latest_outbox_id: Number(latest?.max_id || 0),
    };
  });

  // GET /api/sync/state?peer=laptop-a
  app.get("/state", async (req, reply) => {
    if (!requireSyncAdmin(req, reply)) return;
    const schemaVersion = getSchemaVersion(req);
    if (schemaVersion !== SYNC_SCHEMA_VERSION) {
      return reply.code(400).send({ ok: false, error: `unsupported schema_version ${schemaVersion}` });
    }
    const peer = getPeerName(req);
    const stateRows = db
      .prepare(`SELECT key, value_json, updated_at FROM sync_state ORDER BY key ASC`)
      .all();
    const checkpoints = peer
      ? db
          .prepare(
            `SELECT peer_name, table_name, last_outbox_id, updated_at
             FROM sync_checkpoints
             WHERE peer_name = ?
             ORDER BY table_name ASC`
          )
          .all(peer)
      : db
          .prepare(
            `SELECT peer_name, table_name, last_outbox_id, updated_at
             FROM sync_checkpoints
             ORDER BY peer_name ASC, table_name ASC
             LIMIT 500`
          )
          .all();
    const applied = peer
      ? db.prepare(`SELECT COUNT(*) AS c FROM sync_applied_events WHERE peer_name = ?`).get(peer)
      : db.prepare(`SELECT COUNT(*) AS c FROM sync_applied_events`).get();

    return {
      ok: true,
      schema_version: SYNC_SCHEMA_VERSION,
      peer: peer || null,
      sync_state: stateRows,
      checkpoints,
      applied_events: Number(applied?.c || 0),
    };
  });

  // GET /api/sync/pull?since_id=0&limit=500&peer=laptop-a
  app.get("/pull", async (req, reply) => {
    if (!requireSyncAdmin(req, reply)) return;
    const schemaVersion = getSchemaVersion(req);
    if (schemaVersion !== SYNC_SCHEMA_VERSION) {
      return reply.code(400).send({ ok: false, error: `unsupported schema_version ${schemaVersion}` });
    }
    const dryRun = isDryRun(req);
    const limitRaw = Number(req.query?.limit ?? 500);
    const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(5000, Math.trunc(limitRaw))) : 500;
    const sinceRaw = Number(req.query?.since_id ?? 0);
    let sinceId = Number.isFinite(sinceRaw) ? Math.max(0, Math.trunc(sinceRaw)) : 0;
    const peer = getPeerName(req);
    const siteCode = getSiteCode(req);

    if (peer) {
      const cp = db
        .prepare(
          `SELECT MAX(last_outbox_id) AS last_id
           FROM sync_checkpoints
           WHERE peer_name = ?`
        )
        .get(peer);
      const cpSince = Number(cp?.last_id || 0);
      if (cpSince > sinceId) sinceId = cpSince;
    }

    const rows = db
      .prepare(
        `SELECT id, table_name, row_id, row_uuid, op, site_code, changed_at
         FROM sync_outbox
         WHERE id > ?
           AND (site_code = ? OR site_code = 'main')
         ORDER BY id ASC
         LIMIT ?`
      )
      .all(sinceId, siteCode, limit);

    const events = rows.map((r) => {
      const payload = r.op === "delete" ? null : loadRowPayload(r.table_name, r.row_id);
      return {
        id: Number(r.id),
        table: r.table_name,
        op: r.op,
        row_id: r.row_id == null ? null : Number(r.row_id),
        row_uuid: r.row_uuid || null,
        site_code: r.site_code || "main",
        changed_at: r.changed_at || null,
        payload,
      };
    });

    const lastId = events.length ? events[events.length - 1].id : sinceId;
    return {
      ok: true,
      schema_version: SYNC_SCHEMA_VERSION,
      dry_run: dryRun,
      since_id: sinceId,
      last_id: lastId,
      count: events.length,
      events,
    };
  });

  // POST /api/sync/checkpoint { peer: "laptop-a", last_outbox_id: 1234, table_name?: "daily_hours" }
  app.post("/checkpoint", async (req, reply) => {
    if (!requireSyncAdmin(req, reply)) return;
    const schemaVersion = getSchemaVersion(req);
    if (schemaVersion !== SYNC_SCHEMA_VERSION) {
      return reply.code(400).send({ ok: false, error: `unsupported schema_version ${schemaVersion}` });
    }
    const peer = getPeerName(req);
    const lastOutboxId = Number(req.body?.last_outbox_id);
    const tableNameRaw = String(req.body?.table_name || "").trim();
    const tableName = tableNameRaw && SYNC_TABLES.has(tableNameRaw) ? tableNameRaw : "__all__";

    if (!peer) return reply.code(400).send({ error: "peer required" });
    if (!Number.isInteger(lastOutboxId) || lastOutboxId < 0) {
      return reply.code(400).send({ error: "last_outbox_id must be >= 0 integer" });
    }

    db.prepare(
      `INSERT INTO sync_checkpoints (peer_name, table_name, last_outbox_id, updated_at)
       VALUES (?, ?, ?, datetime('now'))
       ON CONFLICT(peer_name, table_name)
       DO UPDATE SET
         last_outbox_id = excluded.last_outbox_id,
         updated_at = datetime('now')`
    ).run(peer, tableName, lastOutboxId);

    return { ok: true, schema_version: SYNC_SCHEMA_VERSION, peer, table_name: tableName, last_outbox_id: lastOutboxId };
  });

  // POST /api/sync/apply { peer, events: [{ id?, table, op, row_uuid, changed_at, payload }] }
  app.post("/apply", async (req, reply) => {
    if (!requireSyncAdmin(req, reply)) return;
    const schemaVersion = getSchemaVersion(req);
    if (schemaVersion !== SYNC_SCHEMA_VERSION) {
      return reply.code(400).send({ ok: false, error: `unsupported schema_version ${schemaVersion}` });
    }
    const dryRun = isDryRun(req);
    const peer = getPeerName(req);
    const eventsIn = Array.isArray(req.body?.events) ? req.body.events : [];
    if (!peer) return reply.code(400).send({ error: "peer required" });
    if (!eventsIn.length) return reply.code(400).send({ error: "events[] required" });

    const alreadyAppliedStmt = db.prepare(
      `SELECT id FROM sync_applied_events WHERE peer_name = ? AND event_key = ? LIMIT 1`
    );
    const markAppliedStmt = db.prepare(
      `INSERT INTO sync_applied_events (peer_name, event_key, event_id, table_name, row_uuid, op, result_json, applied_at)
       VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))
       ON CONFLICT(peer_name, event_key) DO NOTHING`
    );

    const applyTx = db.transaction((items) =>
      items.map((ev) => {
        const op = String(ev?.op || "").trim().toLowerCase();
        const eventKey = makeEventKey(ev);
        const rowUuid = String(ev?.row_uuid || ev?.payload?.uuid || "").trim() || null;
        const eventId = Number.isInteger(Number(ev?.id)) ? Number(ev.id) : null;

        const prior = alreadyAppliedStmt.get(peer, eventKey);
        if (prior) {
          return { id: ev?.id ?? null, status: "skipped", reason: "duplicate_event", event_key: eventKey };
        }

        let result;
        try {
          if (op === "upsert") result = { id: ev?.id ?? null, ...applyUpsertEvent(ev) };
          else if (op === "delete") result = { id: ev?.id ?? null, ...applyDeleteEvent(ev) };
          else result = { id: ev?.id ?? null, status: "error", reason: "invalid_op" };
        } catch (err) {
          result = { id: ev?.id ?? null, status: "error", reason: String(err?.message || "apply_failed").slice(0, 200) };
        }

        if (result.status !== "error") {
          markAppliedStmt.run(
            peer,
            eventKey,
            eventId,
            String(ev?.table || "").trim() || null,
            rowUuid,
            op || null,
            JSON.stringify(result)
          );
        }
        return { ...result, event_key: eventKey };
      })
    );

    const results = dryRun
      ? eventsIn.map((ev) => {
          const op = String(ev?.op || "").trim().toLowerCase();
          const eventKey = makeEventKey(ev);
          const prior = alreadyAppliedStmt.get(peer, eventKey);
          if (prior) return { id: ev?.id ?? null, status: "skipped", reason: "duplicate_event", event_key: eventKey };
          if (!["upsert", "delete"].includes(op)) {
            return { id: ev?.id ?? null, status: "error", reason: "invalid_op", event_key: eventKey };
          }
          return { id: ev?.id ?? null, status: "would_apply", op, event_key: eventKey };
        })
      : applyTx(eventsIn);
    const applied = results.filter((r) => r.status === "applied").length;
    const skipped = results.filter((r) => r.status === "skipped").length;
    const failed = results.filter((r) => r.status === "error").length;
    const wouldApply = results.filter((r) => r.status === "would_apply").length;

    // auto-advance peer checkpoint to max incoming id (if provided)
    const maxIncomingId = eventsIn.reduce((m, e) => {
      const id = Number(e?.id);
      return Number.isInteger(id) && id > m ? id : m;
    }, 0);
    if (!dryRun && maxIncomingId > 0) {
      db.prepare(
        `INSERT INTO sync_checkpoints (peer_name, table_name, last_outbox_id, updated_at)
         VALUES (?, ?, ?, datetime('now'))
         ON CONFLICT(peer_name, table_name)
         DO UPDATE SET
           last_outbox_id = excluded.last_outbox_id,
           updated_at = datetime('now')`
      ).run(peer, "__inbound__", maxIncomingId);
    }

    return {
      ok: true,
      schema_version: SYNC_SCHEMA_VERSION,
      dry_run: dryRun,
      peer,
      received: eventsIn.length,
      applied,
      would_apply: wouldApply,
      skipped,
      failed,
      results,
    };
  });
}

