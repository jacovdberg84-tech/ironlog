import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import { db } from "../db/client.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function hasColumn(table, col) {
  const rows = db.prepare(`PRAGMA table_info(${table})`).all();
  return rows.some((r) => String(r.name) === col);
}

function pickExistingColumn(table, candidates, fallback) {
  for (const c of candidates) {
    if (hasColumn(table, c)) return c;
  }
  return fallback;
}

/** Avoid storing multi‑MB base64 blobs in inspectpro_ingest_events.payload_json. */
function compactPayloadJson(body) {
  try {
    const o = body && typeof body === "object" ? { ...body } : {};
    if (Array.isArray(o.photos)) {
      o.photos = `[${o.photos.length} photo(s); base64 omitted]`;
    }
    return JSON.stringify(o);
  } catch {
    return "{}";
  }
}

function parseDataUrl(dataUrl) {
  const s = String(dataUrl || "").trim();
  const m = /^data:([^;]+);base64,(.+)$/is.exec(s);
  if (!m) return null;
  try {
    const buf = Buffer.from(String(m[2]).replace(/\s/g, ""), "base64");
    if (!buf.length) return null;
    const mime = String(m[1] || "image/jpeg").split(";")[0].trim();
    return { mime, buffer: buf };
  } catch {
    return null;
  }
}

function extFromMime(mime) {
  const m = String(mime || "").toLowerCase();
  if (m.includes("png")) return ".png";
  if (m.includes("webp")) return ".webp";
  if (m.includes("gif")) return ".gif";
  return ".jpg";
}

function requireInspectproKey(req, reply) {
  const required = String(process.env.INSPECTPRO_INGEST_KEY || "").trim();
  if (!required) {
    reply.code(503).send({ ok: false, error: "INSPECTPRO_INGEST_KEY is not configured on server" });
    return false;
  }
  const provided = String(req.headers["x-inspectpro-key"] || "").trim();
  if (!provided || provided !== required) {
    reply.code(401).send({ ok: false, error: "invalid inspectpro key" });
    return false;
  }
  return true;
}

function classifyIngestError(err) {
  const msg = String(err?.message || err || "").trim();
  const lower = msg.toLowerCase();
  if (!msg) return { code: "INGEST_ERROR", retryable: false, httpStatus: 400 };
  if (lower.includes("database is locked") || lower.includes("database busy") || lower.includes("sql_busy")) {
    return { code: "DB_BUSY", retryable: true, httpStatus: 503, retryAfterSeconds: 30 };
  }
  if (lower.includes("asset not found")) {
    return { code: "ASSET_NOT_FOUND", retryable: false, httpStatus: 404 };
  }
  if (
    lower.includes("required") ||
    lower.includes("must be") ||
    lower.includes("invalid") ||
    lower.includes("uuid is required") ||
    lower.includes("event_type must be")
  ) {
    return { code: "VALIDATION_ERROR", retryable: false, httpStatus: 422 };
  }
  return { code: "INGEST_ERROR", retryable: false, httpStatus: 400 };
}

export default async function inspectproRoutes(app) {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS inspectpro_ingest_events (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      event_uuid TEXT NOT NULL UNIQUE,
      event_type TEXT NOT NULL,
      source TEXT NOT NULL DEFAULT 'inspectpro-manager',
      status TEXT NOT NULL DEFAULT 'ok',
      asset_code TEXT,
      target_id INTEGER,
      error_message TEXT,
      payload_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  const upsertEvent = db.prepare(`
    INSERT INTO inspectpro_ingest_events (
      event_uuid, event_type, source, status, asset_code, target_id, error_message, payload_json, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
    ON CONFLICT(event_uuid) DO UPDATE SET
      event_type = excluded.event_type,
      source = excluded.source,
      status = excluded.status,
      asset_code = excluded.asset_code,
      target_id = excluded.target_id,
      error_message = excluded.error_message,
      payload_json = excluded.payload_json,
      updated_at = excluded.updated_at
  `);

  const dataRoot = process.env.IRONLOG_DATA_DIR || process.cwd();
  const inspectionsDir = path.join(dataRoot, "uploads", "manager-inspections");
  const damageReportsDir = path.join(dataRoot, "uploads", "manager-damage-reports");
  fs.mkdirSync(inspectionsDir, { recursive: true });
  fs.mkdirSync(damageReportsDir, { recursive: true });

  /** Store checklist snapshot, hour meter, and optional data-URL photos for a manager inspection. */
  function persistInspectionIngestAttachments(body, inspectionId, siteCode) {
    if (!inspectionId) return;
    const sc = String(siteCode || "main").trim().toLowerCase() || "main";

    const bundle = {
      checklist: body?.checklist ?? null,
      checklist_details: body?.checklist_details ?? null,
      manager_signature: body?.manager_signature ?? null,
      manager_name: body?.manager_name ?? null,
      hour_meter_reading: body?.hour_meter_reading ?? null,
      source: body?.source ?? null,
    };
    const checklistJson = JSON.stringify(bundle);

    let machineHours = null;
    const hmRaw = body?.hour_meter_reading ?? body?.machine_hours;
    if (hmRaw != null && String(hmRaw).trim() !== "") {
      const n = Number(hmRaw);
      if (Number.isFinite(n) && n >= 0) machineHours = n;
    }

    if (hasColumn("manager_inspections", "checklist_json")) {
      if (machineHours != null && hasColumn("manager_inspections", "machine_hours")) {
        db.prepare(`
          UPDATE manager_inspections
          SET checklist_json = ?, machine_hours = ?, updated_at = datetime('now')
          WHERE id = ?
        `).run(checklistJson, machineHours, inspectionId);
      } else {
        db.prepare(`
          UPDATE manager_inspections SET checklist_json = ?, updated_at = datetime('now') WHERE id = ?
        `).run(checklistJson, inspectionId);
      }
    } else if (machineHours != null && hasColumn("manager_inspections", "machine_hours")) {
      db.prepare(`UPDATE manager_inspections SET machine_hours = ?, updated_at = datetime('now') WHERE id = ?`).run(
        machineHours,
        inspectionId
      );
    }

    if (!Array.isArray(body?.photos)) return;

    const idCol = pickExistingColumn("manager_inspection_photos", ["inspection_id", "manager_inspection_id"], "inspection_id");
    db.prepare(`DELETE FROM manager_inspection_photos WHERE ${idCol} = ?`).run(inspectionId);

    const hasInspectionId = hasColumn("manager_inspection_photos", "inspection_id");
    const hasManagerInspectionId = hasColumn("manager_inspection_photos", "manager_inspection_id");
    const hasImageData = hasColumn("manager_inspection_photos", "image_data");

    body.photos.forEach((dataUrl, i) => {
      const parsed = parseDataUrl(dataUrl);
      if (!parsed) return;
      const ext = extFromMime(parsed.mime);
      const safe = `mi_${inspectionId}_${Date.now()}_${i}_${Math.floor(Math.random() * 100000)}${ext}`;
      fs.writeFileSync(path.join(inspectionsDir, safe), parsed.buffer);
      const relPath = path.join("uploads", "manager-inspections", safe).replace(/\\/g, "/");
      const linkCols = [];
      const linkVals = [];
      if (hasInspectionId) {
        linkCols.push("inspection_id");
        linkVals.push(inspectionId);
      }
      if (hasManagerInspectionId) {
        linkCols.push("manager_inspection_id");
        linkVals.push(inspectionId);
      }
      const insertCols = [...linkCols, "uuid", "site_code", "file_path", ...(hasImageData ? ["image_data"] : []), "caption", "updated_at"];
      const placeholders = insertCols.map((c) => (c === "updated_at" ? "datetime('now')" : "?")).join(", ");
      db.prepare(`
        INSERT INTO manager_inspection_photos (${insertCols.join(", ")})
        VALUES (${placeholders})
      `).run(...linkVals, crypto.randomUUID(), sc, relPath, ...(hasImageData ? [relPath] : []), null);
    });
  }

  function persistDamageIngestAttachments(body, reportId, siteCode) {
    if (!reportId || !Array.isArray(body?.photos)) return;
    const sc = String(siteCode || "main").trim().toLowerCase() || "main";

    const idCol = pickExistingColumn(
      "manager_damage_report_photos",
      ["damage_report_id", "manager_damage_report_id"],
      "damage_report_id"
    );
    db.prepare(`DELETE FROM manager_damage_report_photos WHERE ${idCol} = ?`).run(reportId);

    const hasReportId = hasColumn("manager_damage_report_photos", "damage_report_id");
    const hasLegacyReportId = hasColumn("manager_damage_report_photos", "manager_damage_report_id");
    const hasImageData = hasColumn("manager_damage_report_photos", "image_data");

    body.photos.forEach((dataUrl, i) => {
      const parsed = parseDataUrl(dataUrl);
      if (!parsed) return;
      const ext = extFromMime(parsed.mime);
      const safe = `mdr_${reportId}_${Date.now()}_${i}_${Math.floor(Math.random() * 100000)}${ext}`;
      fs.writeFileSync(path.join(damageReportsDir, safe), parsed.buffer);
      const relPath = path.join("uploads", "manager-damage-reports", safe).replace(/\\/g, "/");
      const linkCols = [];
      const linkVals = [];
      if (hasReportId) {
        linkCols.push("damage_report_id");
        linkVals.push(reportId);
      }
      if (hasLegacyReportId) {
        linkCols.push("manager_damage_report_id");
        linkVals.push(reportId);
      }
      if (!linkCols.length) {
        linkCols.push(idCol);
        linkVals.push(reportId);
      }
      const insertCols = [...linkCols, "uuid", "site_code", "file_path", ...(hasImageData ? ["image_data"] : []), "caption", "updated_at"];
      const placeholders = insertCols.map((c) => (c === "updated_at" ? "datetime('now')" : "?")).join(", ");
      db.prepare(`
        INSERT INTO manager_damage_report_photos (${insertCols.join(", ")})
        VALUES (${placeholders})
      `).run(...linkVals, crypto.randomUUID(), sc, relPath, ...(hasImageData ? [relPath] : []), null);
    });
  }

  const resolveAsset = (body) => {
    const assetIdIn = Number(body?.asset_id || 0);
    if (assetIdIn > 0) {
      return db.prepare(`SELECT id, asset_code FROM assets WHERE id = ? LIMIT 1`).get(assetIdIn);
    }
    const assetCode = String(body?.asset_code || "").trim().toUpperCase();
    if (!assetCode) return null;
    return db.prepare(`SELECT id, asset_code FROM assets WHERE UPPER(asset_code)=UPPER(?) LIMIT 1`).get(assetCode);
  };

  const ingestEvent = (body) => {
    const eventType = String(body.event_type || "").trim().toLowerCase();
    const eventUuid = String(body.uuid || body.event_uuid || "").trim();
    const source = String(body.source || "inspectpro-manager").trim() || "inspectpro-manager";
    if (!eventUuid) throw new Error("uuid is required");
    if (!["inspection", "damage_report"].includes(eventType)) {
      throw new Error("event_type must be inspection or damage_report");
    }

    let status = "ok";
    let targetId = null;
    let assetCode = null;
    let errorMessage = null;
    let errorCode = null;
    let retryable = false;
    let httpStatus = 400;
    let retryAfterSeconds = null;
    try {
      const tx = db.transaction(() => {
        const asset = resolveAsset(body);
        if (!asset) throw new Error("asset not found (asset_id or asset_code required)");
        assetCode = String(asset.asset_code || "");

        if (eventType === "inspection") {
          const inspectionDate = String(body.inspection_date || "").trim();
          if (!isDate(inspectionDate)) throw new Error("inspection_date must be YYYY-MM-DD");
          const inspector = String(body.inspector_name || "").trim() || null;
          const notes = String(body.notes || "").trim() || null;
          const siteCode = String(body.site_code || "main").trim().toLowerCase() || "main";
          const r = db.prepare(`
            INSERT INTO manager_inspections (asset_id, uuid, site_code, inspection_date, inspector_name, notes, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
            ON CONFLICT(uuid) DO UPDATE SET
              asset_id = excluded.asset_id,
              site_code = excluded.site_code,
              inspection_date = excluded.inspection_date,
              inspector_name = excluded.inspector_name,
              notes = excluded.notes,
              updated_at = excluded.updated_at
          `).run(asset.id, eventUuid, siteCode, inspectionDate, inspector, notes);
          if (Number(r.lastInsertRowid || 0) > 0) {
            targetId = Number(r.lastInsertRowid);
          } else {
            const row = db.prepare(`SELECT id FROM manager_inspections WHERE uuid = ? LIMIT 1`).get(eventUuid);
            targetId = Number(row?.id || 0) || null;
          }
          if (targetId) persistInspectionIngestAttachments(body, targetId, siteCode);
        } else {
          const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
          const reportDate = String(body.report_date || "").trim();
          if (!isDate(reportDate)) throw new Error("report_date must be YYYY-MM-DD");
          const severity = String(body.severity || "").trim() || null;
          const damageLocation = String(body.damage_location || "").trim() || null;
          const damageDescription = String(body.damage_description || "").trim() || null;
          const immediateAction = String(body.immediate_action || "").trim() || null;
          const damageTime = String(body.damage_time || "").trim() || null;
          const responsiblePerson = String(body.responsible_person || "").trim() || null;
          const inspector = String(body.inspector_name || "").trim() || null;
          const hourMeterRaw = body.hour_meter;
          const hourMeter = hourMeterRaw == null || String(hourMeterRaw).trim() === "" ? null : Number(hourMeterRaw);
          const outOfService = Number(body.out_of_service || 0) ? 1 : 0;
          const pendingInvestigation = Number(body.pending_investigation || 0) ? 1 : 0;
          const hseReportAvailable = Number(body.hse_report_available || 0) ? 1 : 0;
          const siteCode = String(body.site_code || "main").trim().toLowerCase() || "main";
          if (!severity || !damageLocation || !damageDescription || !immediateAction) {
            throw new Error("severity, damage_location, damage_description, immediate_action are required");
          }
          if (hourMeter != null && !Number.isFinite(hourMeter)) throw new Error("hour_meter must be numeric");
          if (damageTime && !/^\d{2}:\d{2}$/.test(damageTime)) throw new Error("damage_time must be HH:MM");
          const r = db.prepare(`
            INSERT INTO manager_damage_reports (
              asset_id, uuid, site_code, report_date, ${drInspectorCol}, hour_meter,
              damage_location, severity, damage_description, immediate_action, out_of_service,
              damage_time, responsible_person, pending_investigation, hse_report_available, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
            ON CONFLICT(uuid) DO UPDATE SET
              asset_id = excluded.asset_id,
              site_code = excluded.site_code,
              report_date = excluded.report_date,
              ${drInspectorCol} = excluded.${drInspectorCol},
              hour_meter = excluded.hour_meter,
              damage_location = excluded.damage_location,
              severity = excluded.severity,
              damage_description = excluded.damage_description,
              immediate_action = excluded.immediate_action,
              out_of_service = excluded.out_of_service,
              damage_time = excluded.damage_time,
              responsible_person = excluded.responsible_person,
              pending_investigation = excluded.pending_investigation,
              hse_report_available = excluded.hse_report_available,
              updated_at = excluded.updated_at
          `).run(
            asset.id, eventUuid, siteCode, reportDate, inspector, hourMeter,
            damageLocation, severity, damageDescription, immediateAction, outOfService,
            damageTime, responsiblePerson, pendingInvestigation, hseReportAvailable
          );
          if (Number(r.lastInsertRowid || 0) > 0) {
            targetId = Number(r.lastInsertRowid);
          } else {
            const row = db.prepare(`SELECT id FROM manager_damage_reports WHERE uuid = ? LIMIT 1`).get(eventUuid);
            targetId = Number(row?.id || 0) || null;
          }
          if (targetId) persistDamageIngestAttachments(body, targetId, siteCode);
        }
      });
      tx();
    } catch (err) {
      status = "error";
      errorMessage = String(err?.message || err);
      const c = classifyIngestError(err);
      errorCode = c.code;
      retryable = Boolean(c.retryable);
      httpStatus = Number(c.httpStatus || 400);
      retryAfterSeconds = Number(c.retryAfterSeconds || 0) || null;
    }

    upsertEvent.run(
      eventUuid,
      eventType,
      source,
      status,
      assetCode,
      targetId,
      errorMessage,
      compactPayloadJson(body || {})
    );
    return { status, errorMessage, errorCode, retryable, httpStatus, retryAfterSeconds, eventUuid, eventType, targetId, assetCode };
  };

  app.post("/events", async (req, reply) => {
    if (!requireInspectproKey(req, reply)) return;
    const body = req.body || {};
    const { status, errorMessage, errorCode, retryable, httpStatus, retryAfterSeconds, eventUuid, eventType, targetId, assetCode } = ingestEvent(body);
    if (status !== "ok") {
      if (retryAfterSeconds) reply.header("Retry-After", String(retryAfterSeconds));
      return reply.code(httpStatus || 400).send({
        ok: false,
        status,
        error: errorMessage,
        code: errorCode || "INGEST_ERROR",
        retryable: Boolean(retryable),
        retry_after_seconds: retryAfterSeconds,
        uuid: eventUuid,
        event_type: eventType,
      });
    }
    return reply.send({ ok: true, status, uuid: eventUuid, event_type: eventType, target_id: targetId, asset_code: assetCode });
  });

  // Compatibility aliases for integrators expecting direct paths.
  async function replyInspectionPayload(req, reply) {
    if (!requireInspectproKey(req, reply)) return;
    const body = { ...(req.body || {}), event_type: "inspection" };
    const { status, errorMessage, errorCode, retryable, httpStatus, retryAfterSeconds, eventUuid, eventType, targetId, assetCode } = ingestEvent(body);
    if (status !== "ok") {
      if (retryAfterSeconds) reply.header("Retry-After", String(retryAfterSeconds));
      return reply.code(httpStatus || 400).send({
        ok: false,
        status,
        error: errorMessage,
        code: errorCode || "INGEST_ERROR",
        retryable: Boolean(retryable),
        retry_after_seconds: retryAfterSeconds,
        uuid: eventUuid,
        event_type: eventType,
      });
    }
    return reply.send({ ok: true, status, uuid: eventUuid, event_type: eventType, target_id: targetId, asset_code: assetCode });
  }

  app.post("/inspection", replyInspectionPayload);
  // Plural path used by Manager app docs / older clients.
  app.post("/inspections", replyInspectionPayload);

  app.post("/damage-report", async (req, reply) => {
    if (!requireInspectproKey(req, reply)) return;
    const body = { ...(req.body || {}), event_type: "damage_report" };
    const { status, errorMessage, errorCode, retryable, httpStatus, retryAfterSeconds, eventUuid, eventType, targetId, assetCode } = ingestEvent(body);
    if (status !== "ok") {
      if (retryAfterSeconds) reply.header("Retry-After", String(retryAfterSeconds));
      return reply.code(httpStatus || 400).send({
        ok: false,
        status,
        error: errorMessage,
        code: errorCode || "INGEST_ERROR",
        retryable: Boolean(retryable),
        retry_after_seconds: retryAfterSeconds,
        uuid: eventUuid,
        event_type: eventType,
      });
    }
    return reply.send({ ok: true, status, uuid: eventUuid, event_type: eventType, target_id: targetId, asset_code: assetCode });
  });

  app.get("/status", async (req, reply) => {
    try {
      const limit = Math.max(1, Math.min(200, Number(req.query?.limit || 20)));
      const rows = db.prepare(`
        SELECT id, event_uuid, event_type, source, status, asset_code, target_id, error_message, created_at, updated_at
        FROM inspectpro_ingest_events
        ORDER BY id DESC
        LIMIT ?
      `).all(limit);
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // Pull endpoints for InspectPro "domain sync" flows.
  app.get("/pull/assets", async (req, reply) => {
    if (!requireInspectproKey(req, reply)) return;
    try {
      const rows = db.prepare(`
        SELECT id, asset_code, asset_name, category, active, archived
        FROM assets
        WHERE active = 1 AND IFNULL(archived, 0) = 0
        ORDER BY asset_code ASC
      `).all();
      return reply.send({
        ok: true,
        assets: rows.map((r) => ({
          asset_id: Number(r.id),
          asset_code: String(r.asset_code || ""),
          asset_name: String(r.asset_name || ""),
          category: String(r.category || ""),
          active: Number(r.active || 0) === 1,
        })),
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/pull/recent", async (req, reply) => {
    if (!requireInspectproKey(req, reply)) return;
    try {
      const days = Math.max(1, Math.min(90, Number(req.query?.days || 14)));
      const inspections = db.prepare(`
        SELECT mi.id, mi.uuid, mi.inspection_date, mi.inspector_name, mi.notes, mi.updated_at, a.asset_code
        FROM manager_inspections mi
        JOIN assets a ON a.id = mi.asset_id
        WHERE mi.inspection_date >= date('now', ?)
        ORDER BY mi.inspection_date DESC, mi.id DESC
        LIMIT 500
      `).all(`-${days} day`);
      const drInspectorCol = pickExistingColumn("manager_damage_reports", ["inspector_name", "inspector", "manager_name"], "inspector_name");
      const damages = db.prepare(`
        SELECT dr.id, dr.uuid, dr.report_date, dr.${drInspectorCol} AS inspector_name,
               dr.hour_meter, dr.damage_location, dr.severity, dr.damage_description,
               dr.immediate_action, dr.out_of_service, dr.updated_at, a.asset_code
        FROM manager_damage_reports dr
        JOIN assets a ON a.id = dr.asset_id
        WHERE dr.report_date >= date('now', ?)
        ORDER BY dr.report_date DESC, dr.id DESC
        LIMIT 500
      `).all(`-${days} day`);
      return reply.send({
        ok: true,
        inspections,
        damage_reports: damages,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
}

