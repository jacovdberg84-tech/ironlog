// IRONLOG/api/routes/assets.routes.js
import { db } from "../db/client.js";
import { getAssetCurrentHoursInfo } from "../utils/assetMeterHours.js";
import {
  ensureMasterDataSchema,
  validateAgainstMdmPolicy,
  validateAssetGovernanceOptional,
} from "../utils/masterdataGovernance.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

export default async function assetRoutes(app) {
  ensureMasterDataSchema();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS asset_qr_profiles (
      asset_id INTEGER PRIMARY KEY,
      qr_payload TEXT NOT NULL,
      qr_text TEXT NOT NULL,
      generated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
    )
  `).run();

  /* =========================
     PREPARED STATEMENTS
  ========================= */

  const getAssetByCode = db.prepare(`
    SELECT
      id, asset_code, asset_name, category,
      active, is_standby,
      archived, archive_reason, archived_at,
      created_at
    FROM assets
    WHERE asset_code = ?
  `);

  const getAssetById = db.prepare(`
    SELECT
      id, asset_code, asset_name, category,
      active, is_standby,
      archived, archive_reason, archived_at,
      created_at
    FROM assets
    WHERE id = ?
  `);
  const getStoredQrProfile = db.prepare(`
    SELECT qr_payload, qr_text, generated_at
    FROM asset_qr_profiles
    WHERE asset_id = ?
  `);
  const upsertQrProfile = db.prepare(`
    INSERT INTO asset_qr_profiles (asset_id, qr_payload, qr_text, generated_at)
    VALUES (?, ?, ?, datetime('now'))
    ON CONFLICT(asset_id) DO UPDATE SET
      qr_payload = excluded.qr_payload,
      qr_text = excluded.qr_text,
      generated_at = datetime('now')
  `);

  function hasTable(tableName) {
    const row = db.prepare(`
      SELECT name
      FROM sqlite_master
      WHERE type = 'table' AND name = ?
    `).get(tableName);
    return Boolean(row);
  }
  function hasColumn(tableName, columnName) {
    if (!hasTable(tableName)) return false;
    const cols = db.prepare(`PRAGMA table_info(${tableName})`).all();
    return cols.some((c) => String(c.name || "") === String(columnName));
  }
  function firstExistingColumn(tableName, candidates) {
    for (const c of candidates) {
      if (hasColumn(tableName, c)) return c;
    }
    return null;
  }

  const insertAsset = db.prepare(`
    INSERT INTO assets (
      asset_code, asset_name, category,
      active, is_standby,
      archived, archive_reason, archived_at,
      department_code, cost_center_code, data_owner_username
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  const updateActive = db.prepare(`UPDATE assets SET active = ? WHERE asset_code = ?`);
  const updateStandby = db.prepare(`UPDATE assets SET is_standby = ? WHERE asset_code = ?`);

  /* =========================
     ROUTES
  ========================= */

  // GET /api/assets?include_archived=1
  app.get("/", async (req) => {
    const includeArchived = String(req.query?.include_archived || "0") === "1";

    const rows = db.prepare(`
      SELECT
        id, asset_code, asset_name, category,
        active, is_standby,
        archived, archive_reason, archived_at,
        department_code, cost_center_code, data_owner_username,
        created_at
      FROM assets
      WHERE (? = 1 OR archived = 0)
      ORDER BY asset_code ASC
    `).all(includeArchived ? 1 : 0);

    return rows.map((r) => ({
      ...r,
      active: Number(r.active),
      is_standby: Number(r.is_standby),
      archived: Number(r.archived),
    }));
  });

  function siteCodeFromReq(req) {
    return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
  }

  // POST /api/assets
  app.post("/", async (req, reply) => {
    const body = req.body || {};
    const asset_code = String(body.asset_code || "").trim();
    const asset_name = String(body.asset_name || "").trim();
    const category = String(body.category || "").trim() || null;

    const active = body.active === 0 || body.active === false ? 0 : 1;
    const is_standby = body.is_standby ? 1 : 0;

    // archived fields (optional)
    const archived = body.archived ? 1 : 0;
    const archive_reason = archived ? (String(body.archive_reason || "").trim() || null) : null;
    const archived_at = archived ? (String(body.archived_at || "").trim() || null) : null;

    const department_code =
      body.department_code != null && String(body.department_code).trim() !== ""
        ? String(body.department_code).trim().toUpperCase()
        : null;
    const cost_center_code =
      body.cost_center_code != null && String(body.cost_center_code).trim() !== ""
        ? String(body.cost_center_code).trim().toUpperCase()
        : null;
    const data_owner_username =
      body.data_owner_username != null && String(body.data_owner_username).trim() !== ""
        ? String(body.data_owner_username).trim()
        : null;

    if (!asset_code || !asset_name) {
      return reply.code(400).send({ error: "asset_code and asset_name are required" });
    }

    const policyBody = {
      department_code,
      cost_center_code,
      data_owner_username,
    };
    const pol = validateAgainstMdmPolicy(siteCodeFromReq(req), "asset", policyBody);
    if (!pol.ok) {
      return reply.code(400).send({ error: `missing required fields: ${pol.missing.join(", ")}` });
    }

    const gov = validateAssetGovernanceOptional(siteCodeFromReq(req), {
      department_code,
      cost_center_code,
    });
    if (!gov.ok) return reply.code(400).send({ error: gov.error });

    try {
      const r = insertAsset.run(
        asset_code,
        asset_name,
        category,
        active,
        is_standby,
        archived,
        archive_reason,
        archived_at,
        department_code,
        cost_center_code,
        data_owner_username
      );
      return reply.code(201).send({ ok: true, id: Number(r.lastInsertRowid) });
    } catch (e) {
      return reply.code(400).send({ error: e.message || String(e) });
    }
  });

  // POST /api/assets/:asset_code/archive  { archived: true/false, reason? }
  app.post("/:asset_code/archive", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    const body = req.body || {};

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const archived = body.archived ? 1 : 0;
    const reason = archived ? (String(body.reason || "").trim() || null) : null;

    if (archived) {
      db.prepare(`
        UPDATE assets
        SET archived = 1, archive_reason = ?, archived_at = datetime('now')
        WHERE asset_code = ?
      `).run(reason, asset_code);
    } else {
      db.prepare(`
        UPDATE assets
        SET archived = 0, archive_reason = NULL, archived_at = NULL
        WHERE asset_code = ?
      `).run(asset_code);
    }

    return { ok: true };
  });

  // PATCH /api/assets/:asset_code  { active?, is_standby? }
  app.patch("/:asset_code", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    const body = req.body || {};

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const tx = db.transaction(() => {
      if (body.active !== undefined) updateActive.run(body.active ? 1 : 0, asset_code);
      if (body.is_standby !== undefined) updateStandby.run(body.is_standby ? 1 : 0, asset_code);
    });

    tx();

    const updated = getAssetByCode.get(asset_code);
    return {
      ok: true,
      asset: {
        ...updated,
        active: Number(updated.active),
        is_standby: Number(updated.is_standby),
        archived: Number(updated.archived),
      },
    };
  });

  // ------------------------------------------------------------
  // Hour meter snapshot (e.g. InspectPro Manager)
  // GET /api/assets/:asset_id_or_asset_code/hours
  // Resolves numeric segment as asset id first, then asset_code.
  // ------------------------------------------------------------
  app.get("/:identifier/hours", async (req, reply) => {
    const identifier = String(req.params.identifier || "").trim();
    if (!identifier) return reply.code(404).send({ error: "Asset not found" });

    let asset = null;
    if (/^\d+$/.test(identifier)) {
      asset = getAssetById.get(Number(identifier));
    }
    if (!asset) {
      asset = getAssetByCode.get(identifier);
    }
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const meter = getAssetCurrentHoursInfo(asset.id);
    return {
      ok: true,
      asset_id: asset.id,
      asset_code: asset.asset_code,
      asset_name: asset.asset_name,
      category: asset.category ?? null,
      current_hours: meter.hours,
      hour_meter_source: meter.source,
      latest_daily_work_date: meter.latest_work_date ?? null,
    };
  });

  function buildMachineStatus(assetId) {
    const latestBreakdown = db.prepare(`
      SELECT
        breakdown_date,
        status,
        end_at
      FROM breakdowns
      WHERE asset_id = ? 
      ORDER BY breakdown_date DESC, id DESC
      LIMIT 1
    `).get(assetId);
    const latestDaily = db.prepare(`
      SELECT is_used, work_date
      FROM daily_hours
      WHERE asset_id = ?
      ORDER BY work_date DESC, id DESC
      LIMIT 1
    `).get(assetId);

    const isBreakdownOpen = (() => {
      if (!latestBreakdown) return false;
      const st = String(latestBreakdown.status || "").trim().toUpperCase();
      if (st === "OPEN") return true;
      if (st && st !== "OPEN") return false;
      return !latestBreakdown.end_at;
    })();

    if (isBreakdownOpen) {
      if (!latestDaily?.work_date) return "DOWN";
      const bdDate = String(latestBreakdown.breakdown_date || "");
      const dhDate = String(latestDaily.work_date || "");
      if (bdDate && dhDate && bdDate >= dhDate) return "DOWN";
    }

    if (!latestDaily) return "UNKNOWN";
    return Number(latestDaily.is_used) === 1 ? "PRODUCTION" : "STANDBY";
  }

  function buildInspectionSummary(assetId) {
    const candidates = [];
    if (hasTable("manager_inspections")) {
      const dateCol = firstExistingColumn("manager_inspections", ["inspection_date", "check_date", "created_at"]);
      if (dateCol) {
      candidates.push(
        db.prepare(`
          SELECT DATE(MAX(${dateCol})) AS latest_date
          FROM manager_inspections
          WHERE asset_id = ?
        `).get(assetId)?.latest_date || null
      );
      }
    }
    if (hasTable("tyre_inspections")) {
      const dateCol = firstExistingColumn("tyre_inspections", ["inspection_date", "check_date", "created_at", "date"]);
      if (dateCol) {
      candidates.push(
        db.prepare(`
          SELECT DATE(MAX(${dateCol})) AS latest_date
          FROM tyre_inspections
          WHERE asset_id = ?
        `).get(assetId)?.latest_date || null
      );
      }
    }
    if (hasTable("tire_inspections")) {
      const dateCol = firstExistingColumn("tire_inspections", ["inspection_date", "check_date", "created_at", "date"]);
      if (dateCol) {
      candidates.push(
        db.prepare(`
          SELECT DATE(MAX(${dateCol})) AS latest_date
          FROM tire_inspections
          WHERE asset_id = ?
        `).get(assetId)?.latest_date || null
      );
      }
    }
    const sorted = candidates.filter(Boolean).sort().reverse();
    return sorted[0] || null;
  }

  function resolveWebOrigin(req) {
    const envBase = String(process.env.IRONLOG_PUBLIC_BASE_URL || "").trim().replace(/\/+$/, "");
    if (envBase) return envBase;
    const protoHeader = String(req.headers["x-forwarded-proto"] || "").split(",")[0].trim();
    const hostHeader = String(req.headers["x-forwarded-host"] || req.headers.host || "").split(",")[0].trim();
    const proto = protoHeader || "http";
    if (hostHeader) return `${proto}://${hostHeader}`;
    return "";
  }

  function buildQrProfile(asset, req) {
    function inferMakeModelFromName(name, code) {
      const raw = String(name || "").trim();
      if (!raw) return { make: null, model: null };
      const parts = raw.split(/\s+/).filter(Boolean);
      if (!parts.length) return { make: null, model: null };

      // Typical names like "CAT 336D Excavator" => make CAT, model 336D
      const make = parts[0] ? String(parts[0]).toUpperCase() : null;
      let model = null;
      if (parts.length >= 2) {
        const second = String(parts[1] || "");
        if (/[0-9]/.test(second) || second.length <= 12) {
          model = second.toUpperCase();
        }
      }
      // Fallback: derive model-ish token from asset code if name is sparse.
      if (!model) {
        const codeToken = String(code || "").split(/[-_\s]/).find((t) => /[0-9]/.test(t));
        if (codeToken) model = codeToken.toUpperCase();
      }
      return {
        make: make || null,
        model: model || null,
      };
    }

    const makeCol = firstExistingColumn("assets", ["make", "asset_make", "manufacturer", "brand"]);
    const modelCol = firstExistingColumn("assets", ["model", "asset_model"]);
    let assetMake = null;
    let assetModel = null;
    if (makeCol || modelCol) {
      const fields = [makeCol ? `${makeCol} AS make` : "NULL AS make", modelCol ? `${modelCol} AS model` : "NULL AS model"].join(", ");
      const row = db.prepare(`SELECT ${fields} FROM assets WHERE id = ?`).get(asset.id);
      assetMake = row?.make != null ? String(row.make).trim() || null : null;
      assetModel = row?.model != null ? String(row.model).trim() || null : null;
    }
    if (!assetMake || !assetModel) {
      const inferred = inferMakeModelFromName(asset.asset_name, asset.asset_code);
      if (!assetMake) assetMake = inferred.make;
      if (!assetModel) assetModel = inferred.model;
    }

    const meter = getAssetCurrentHoursInfo(asset.id);
    const currentHours = Number(Number(meter.hours || 0).toFixed(1));

    const planRows = db.prepare(`
      SELECT service_name, interval_hours, last_service_hours
      FROM maintenance_plans
      WHERE asset_id = ?
        AND active = 1
      ORDER BY id ASC
    `).all(asset.id);
    const dueRows = planRows.map((p) => {
      const nextDueHours = Number(p.last_service_hours || 0) + Number(p.interval_hours || 0);
      const remaining = nextDueHours - currentHours;
      return {
        service_name: String(p.service_name || "Service"),
        next_due_hours: Number(nextDueHours.toFixed(1)),
        remaining_hours: Number(remaining.toFixed(1)),
      };
    }).sort((a, b) => a.remaining_hours - b.remaining_hours);
    const nextService = dueRows[0] || null;

    const fuel30d = db.prepare(`
      SELECT
        COALESCE(SUM(liters), 0) AS liters_30d,
        MAX(log_date) AS latest_log_date
      FROM fuel_logs
      WHERE asset_id = ?
        AND log_date >= DATE('now', '-30 day')
    `).get(asset.id);
    const latestFuel = db.prepare(`
      SELECT log_date, liters
      FROM fuel_logs
      WHERE asset_id = ?
      ORDER BY log_date DESC, id DESC
      LIMIT 1
    `).get(asset.id);

    const profile = {
      generated_at: new Date().toISOString(),
      asset: {
        asset_code: asset.asset_code,
        asset_name: asset.asset_name || null,
        category: asset.category || null,
        make: assetMake,
        model: assetModel,
      },
      scan_url: (() => {
        const origin = resolveWebOrigin(req);
        if (!origin) return `/web/asset-qr.html?asset_code=${encodeURIComponent(asset.asset_code)}`;
        return `${origin}/web/asset-qr.html?asset_code=${encodeURIComponent(asset.asset_code)}`;
      })(),
      status: buildMachineStatus(asset.id),
      meter: {
        current_hours: currentHours,
        source: meter.source || "unknown",
      },
      next_service_due: nextService
        ? {
            service_name: nextService.service_name,
            next_due_hours: nextService.next_due_hours,
            remaining_hours: nextService.remaining_hours,
            is_overdue: nextService.remaining_hours <= 0,
          }
        : null,
      fuel: {
        liters_last_30_days: Number(Number(fuel30d?.liters_30d || 0).toFixed(1)),
        last_fill_date: latestFuel?.log_date || null,
        last_fill_liters: latestFuel?.liters != null ? Number(Number(latestFuel.liters).toFixed(1)) : null,
      },
      inspections: {
        last_inspection_date: buildInspectionSummary(asset.id),
      },
    };

    const nextDueText = nextService
      ? `${nextService.service_name} at ${nextService.next_due_hours}h (${nextService.remaining_hours}h remaining)`
      : "No active maintenance plan";
    const fuelText = `${profile.fuel.liters_last_30_days}L in last 30 days`;
    const inspectText = profile.inspections.last_inspection_date || "No inspection date";
    const qrText = [
      `IRONLOG ${asset.asset_code}`,
      `Scan URL: ${profile.scan_url}`,
      `Status: ${profile.status}`,
      `Current meter: ${currentHours}h`,
      `Next service: ${nextDueText}`,
      `Fuel: ${fuelText}`,
      `Last inspection: ${inspectText}`,
    ].join("\n");

    return { profile, qrText };
  }

  app.get("/:asset_code/qr-profile", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ error: "Asset code is required" });
    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const stored = getStoredQrProfile.get(asset.id);
    const live = buildQrProfile(asset, req);
    let storedPayload = null;
    if (stored?.qr_payload) {
      try {
        storedPayload = JSON.parse(String(stored.qr_payload || "{}"));
      } catch {
        storedPayload = null;
      }
    }

    return {
      ok: true,
      asset_code: asset.asset_code,
      stored: stored
        ? {
            qr_payload: storedPayload,
            qr_text: stored.qr_text,
            generated_at: stored.generated_at,
          }
        : null,
      live_preview: live.profile,
      live_qr_text: live.qrText,
    };
  });

  app.post("/:asset_code/qr-profile/refresh", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ error: "Asset code is required" });
    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const built = buildQrProfile(asset, req);
    upsertQrProfile.run(asset.id, JSON.stringify(built.profile), built.qrText);

    return {
      ok: true,
      asset_code: asset.asset_code,
      qr_payload: built.profile,
      qr_text: built.qrText,
    };
  });

  // ============================================================
  // ASSET HISTORY (Timeline)
  // GET /api/assets/:asset_code/history?start=YYYY-MM-DD&end=YYYY-MM-DD
  // ============================================================
  app.get("/:asset_code/history", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const siteCode = String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";

    const startOk = start && isDate(start);
    const endOk = end && isDate(end);

    function dateFilter(col) {
      const clauses = [];
      const params = [];
      if (startOk) { clauses.push(`${col} >= ?`); params.push(start); }
      if (endOk) { clauses.push(`${col} <= ?`); params.push(end); }
      return { sql: clauses.length ? " AND " + clauses.join(" AND ") : "", params };
    }

    function hasTable(tableName) {
      const row = db.prepare(`
        SELECT name
        FROM sqlite_master
        WHERE type = 'table' AND name = ?
      `).get(tableName);
      return Boolean(row);
    }

    function hasColumn(tableName, columnName) {
      if (!hasTable(tableName)) return false;
      const cols = db.prepare(`PRAGMA table_info(${tableName})`).all();
      return cols.some((c) => String(c.name) === columnName);
    }

    function firstExistingTable(candidates) {
      for (const name of candidates) {
        if (hasTable(name)) return name;
      }
      return null;
    }

    function firstExistingColumn(tableName, candidates) {
      for (const name of candidates) {
        if (hasColumn(tableName, name)) return name;
      }
      return null;
    }

    // ---- BREAKDOWNS (include downtime total + downtime log lines)
    const bdF = dateFilter("b.breakdown_date");
    let breakdowns = [];
    let downtimeLogsByBreakdown = new Map();

    try {
      breakdowns = db.prepare(`
        SELECT
          b.id,
          b.breakdown_date AS date,
          b.description AS title,
          b.critical,
          COALESCE(SUM(l.hours_down), 0) AS downtime_hours
        FROM breakdowns b
        LEFT JOIN breakdown_downtime_logs l ON l.breakdown_id = b.id
        WHERE b.asset_id = ? ${bdF.sql}
        GROUP BY b.id
        ORDER BY b.breakdown_date DESC, b.id DESC
        LIMIT 200
      `).all(asset.id, ...bdF.params);

      const ids = breakdowns.map((b) => b.id);
      if (ids.length) {
        const placeholders = ids.map(() => "?").join(",");
        const logs = db.prepare(`
          SELECT breakdown_id, log_date, hours_down, notes, created_at
          FROM breakdown_downtime_logs
          WHERE breakdown_id IN (${placeholders})
          ORDER BY log_date DESC, id DESC
        `).all(...ids);

        downtimeLogsByBreakdown = new Map();
        for (const l of logs) {
          const arr = downtimeLogsByBreakdown.get(l.breakdown_id) || [];
          arr.push({
            log_date: l.log_date,
            hours_down: Number(l.hours_down || 0),
            notes: l.notes || "",
            created_at: l.created_at,
          });
          downtimeLogsByBreakdown.set(l.breakdown_id, arr);
        }
      }
    } catch {
      breakdowns = db.prepare(`
        SELECT
          b.id,
          b.breakdown_date AS date,
          b.description AS title,
          b.critical,
          b.downtime_hours AS downtime_hours
        FROM breakdowns b
        WHERE b.asset_id = ? ${bdF.sql}
        ORDER BY b.breakdown_date DESC, b.id DESC
        LIMIT 200
      `).all(asset.id, ...bdF.params);

      downtimeLogsByBreakdown = new Map();
    }

    const woForBreakdown = db.prepare(`
      SELECT id
      FROM work_orders
      WHERE source = 'breakdown' AND reference_id = ?
      ORDER BY id DESC
      LIMIT 1
    `);

    const breakdownEvents = breakdowns.map((b) => ({
      type: "breakdown",
      date: b.date,
      title: b.title,
      work_order_id: woForBreakdown.get(b.id)?.id ?? null,
      details: {
        breakdown_id: b.id,
        critical: Boolean(b.critical),
        downtime_hours: Number(b.downtime_hours || 0),
        downtime_logs: downtimeLogsByBreakdown.get(b.id) || [],
        photos: [],
        photo: null,
      },
    }));

    if (hasTable("breakdown_photos") && breakdownEvents.length) {
      const ids = breakdownEvents.map((e) => Number(e.details.breakdown_id)).filter((n) => Number.isFinite(n));
      if (ids.length) {
        const placeholders = ids.map(() => "?").join(",");
        const rows = db.prepare(`
          SELECT breakdown_id, file_path, photo_stage, created_at
          FROM breakdown_photos
          WHERE breakdown_id IN (${placeholders})
          ORDER BY id DESC
        `).all(...ids);

        const byBreakdown = new Map();
        for (const r of rows) {
          const arr = byBreakdown.get(r.breakdown_id) || [];
          arr.push({
            file_path: r.file_path,
            photo_stage: r.photo_stage || null,
            created_at: r.created_at || null,
          });
          byBreakdown.set(r.breakdown_id, arr);
        }

        for (const ev of breakdownEvents) {
          const photos = byBreakdown.get(Number(ev.details.breakdown_id)) || [];
          ev.details.photos = photos;
          ev.details.photo = photos[0]?.file_path || null;
        }
      }
    }

    // ---- WORK ORDERS
    const woF = dateFilter("DATE(w.opened_at)");
    const workOrders = db.prepare(`
      SELECT
        w.id,
        DATE(w.opened_at) AS date,
        w.source,
        w.status,
        w.opened_at,
        w.closed_at
      FROM work_orders w
      WHERE w.asset_id = ? ${woF.sql}
      ORDER BY w.id DESC
      LIMIT 200
    `).all(asset.id, ...woF.params).map((w) => ({
      type: "work_order",
      date: w.date,
      title: `WO #${w.id} (${w.source})`,
      work_order_id: w.id,
      details: {
        source: w.source,
        status: w.status,
        opened_at: w.opened_at,
        closed_at: w.closed_at,
      },
    }));

    if (hasTable("work_order_photos") && workOrders.length) {
      const ids = workOrders.map((w) => Number(w.work_order_id)).filter((n) => Number.isFinite(n));
      if (ids.length) {
        const placeholders = ids.map(() => "?").join(",");
        const rows = db.prepare(`
          SELECT work_order_id, file_path, photo_stage, created_at
          FROM work_order_photos
          WHERE work_order_id IN (${placeholders})
          ORDER BY id DESC
        `).all(...ids);

        const byWo = new Map();
        for (const r of rows) {
          const arr = byWo.get(r.work_order_id) || [];
          arr.push({
            file_path: r.file_path,
            photo_stage: r.photo_stage || null,
            created_at: r.created_at || null,
          });
          byWo.set(r.work_order_id, arr);
        }

        for (const ev of workOrders) {
          const photos = byWo.get(Number(ev.work_order_id)) || [];
          ev.details.photos = photos;
          ev.details.photo = photos[0]?.file_path || null;
        }
      }
    }

    // ---- GET CHANGE SLIPS (optional table)
    let getSlipEvents = [];
    if (hasTable("get_change_slips")) {
      const gsF = dateFilter("g.slip_date");
      getSlipEvents = db.prepare(`
        SELECT g.id, g.slip_date AS date, g.location, g.notes
        FROM get_change_slips g
        WHERE g.asset_id = ? ${gsF.sql}
        ORDER BY g.slip_date DESC, g.id DESC
        LIMIT 200
      `).all(asset.id, ...gsF.params).map((g) => {
        const items = hasTable("get_change_items")
          ? db.prepare(`
              SELECT position, part_code, part_name, qty, reason
              FROM get_change_items
              WHERE slip_id = ?
              ORDER BY id ASC
            `).all(g.id)
          : [];

        const wo = hasTable("work_order_links")
          ? db.prepare(`
              SELECT work_order_id
              FROM work_order_links
              WHERE link_type = 'get_slip' AND link_id = ?
              ORDER BY id DESC
              LIMIT 1
            `).get(g.id)
          : null;

        return {
          type: "get_slip",
          date: g.date,
          title: `GET Change Slip #${g.id}`,
          work_order_id: wo?.work_order_id ?? null,
          details: {
            slip_id: g.id,
            location: g.location,
            notes: g.notes,
            items,
          },
        };
      });
    }

    // ---- COMPONENT CHANGE SLIPS (optional table)
    let componentSlipEvents = [];
    if (hasTable("component_change_slips")) {
      const csF = dateFilter("c.slip_date");
      componentSlipEvents = db.prepare(`
        SELECT
          c.id,
          c.slip_date AS date,
          c.component,
          c.serial_out,
          c.serial_in,
          c.hours_at_change,
          c.notes
        FROM component_change_slips c
        WHERE c.asset_id = ? ${csF.sql}
        ORDER BY c.slip_date DESC, c.id DESC
        LIMIT 200
      `).all(asset.id, ...csF.params).map((c) => {
        const wo = hasTable("work_order_links")
          ? db.prepare(`
              SELECT work_order_id
              FROM work_order_links
              WHERE link_type = 'component_slip' AND link_id = ?
              ORDER BY id DESC
              LIMIT 1
            `).get(c.id)
          : null;

        return {
          type: "component_slip",
          date: c.date,
          title: `Component Change: ${c.component} (#${c.id})`,
          work_order_id: wo?.work_order_id ?? null,
          details: {
            slip_id: c.id,
            component: c.component,
            serial_out: c.serial_out,
            serial_in: c.serial_in,
            hours_at_change: c.hours_at_change,
            notes: c.notes,
          },
        };
      });
    }

    // ---- BREAKDOWN OPS SLIPS (hose / GET / component / tyre — stored in ops_slip_reports)
    const summarizeOpsSlipForHistory = (slip_type, payload_json) => {
      let p = {};
      try {
        p = JSON.parse(String(payload_json || "{}"));
      } catch {
        return { parse_error: true };
      }
      const picturesAttached = Array.isArray(p.pictures) ? p.pictures.length : 0;
      const base = { pictures_attached: picturesAttached };
      if (slip_type === "hose_failure") {
        return {
          ...base,
          date_fitted: p.date_fitted || null,
          reason: p.reason_fitted ? String(p.reason_fitted).slice(0, 240) : null,
          preventable: Boolean(p.preventable),
          hose_part_code: p.hose_part_code || null,
          oil_loss_part_code: p.oil_loss_part_code || null,
          hose_qty: p.hose_qty,
          oil_loss_qty: p.oil_loss_qty,
          slip_total_usd: p.slip_total_usd,
        };
      }
      if (slip_type === "get_change") {
        return {
          ...base,
          part_code: p.part_code || null,
          part_qty: p.part_qty,
          supplier: p.supplier || null,
          date_changed: p.date_changed || null,
          description_part_code: p.description_part_code || null,
          hours_fitted: p.hours_fitted,
        };
      }
      if (slip_type === "component_change") {
        return {
          ...base,
          date_changed: p.date_changed || null,
          component_type: p.component_type || null,
          part_code: p.part_code || null,
          reason: p.reason ? String(p.reason).slice(0, 200) : null,
          hours_in_service: p.hours_in_service,
          line_total_usd: p.cost,
        };
      }
      if (slip_type === "tyre_change") {
        const tyres = Array.isArray(p.tyres) ? p.tyres : [];
        return {
          ...base,
          tyre_lines: tyres.length,
          positions: tyres.map((t) => t.position).filter(Boolean).slice(0, 10),
        };
      }
      return base;
    };

    let opsSlipEvents = [];
    if (hasTable("ops_slip_reports")) {
      const opsF = dateFilter("r.report_date");
      const slipTitle = {
        hose_failure: "Hose failure slip",
        get_change: "G.E.T. change slip",
        component_change: "Component change slip",
        tyre_change: "Tyre change slip",
      };
      const rows = db
        .prepare(
          `
        SELECT r.id, r.slip_type, r.report_date AS date, r.created_at, r.created_by, r.payload_json
        FROM ops_slip_reports r
        WHERE r.asset_id = ? AND r.site_code = ? ${opsF.sql}
        ORDER BY r.report_date DESC, r.id DESC
        LIMIT 200
      `
        )
        .all(asset.id, siteCode, ...opsF.params);

      opsSlipEvents = rows.map((r) => ({
        type: "ops_slip",
        date: r.date,
        title: `${slipTitle[r.slip_type] || "Ops slip"} #${r.id}`,
        work_order_id: null,
        details: {
          slip_id: Number(r.id),
          slip_type: r.slip_type,
          created_at: r.created_at,
          created_by: r.created_by || null,
          summary: summarizeOpsSlipForHistory(r.slip_type, r.payload_json),
        },
      }));
    }

    // ---- DAMAGE REPORTS (schema-flexible / optional tables)
    let damageEvents = [];
    const damageTable = firstExistingTable([
      "equipment_damage_reports",
      "damage_reports",
      "asset_damage_reports",
    ]);

    if (damageTable) {
      const dDateCol = firstExistingColumn(damageTable, ["report_date", "damage_date", "date", "created_at"]);
      const dTitleCol = firstExistingColumn(damageTable, ["title", "damage_type", "component", "area", "category"]);
      const dDescCol = firstExistingColumn(damageTable, ["description", "notes", "details", "report"]);
      const dPhotoCol = firstExistingColumn(damageTable, ["photo_path", "photo_url", "image_path", "image_url", "image"]);

      const dDateExpr = dDateCol ? `DATE(d.${dDateCol})` : "DATE('now')";
      const dF = dateFilter(dDateExpr);

      const dTitleExpr = dTitleCol ? `COALESCE(d.${dTitleCol}, 'Damage Report')` : "'Damage Report'";
      const dDescExpr = dDescCol ? `d.${dDescCol}` : "NULL";
      const dPhotoExpr = dPhotoCol ? `d.${dPhotoCol}` : "NULL";

      damageEvents = db.prepare(`
        SELECT
          d.id,
          ${dDateExpr} AS date,
          ${dTitleExpr} AS title,
          ${dDescExpr} AS notes,
          ${dPhotoExpr} AS photo
        FROM ${damageTable} d
        WHERE d.asset_id = ? ${dF.sql}
        ORDER BY date DESC, d.id DESC
        LIMIT 200
      `).all(asset.id, ...dF.params).map((d) => ({
        type: "damage_report",
        date: d.date,
        title: `Damage: ${d.title || `Report #${d.id}`}`,
        work_order_id: null,
        details: {
          damage_id: d.id,
          notes: d.notes || null,
          photo: d.photo || null,
        },
      }));
    }

    // ---- TYRE EVENTS (change slips + inspections, schema-flexible)
    let tyreChangeEvents = [];
    let tyreInspectionEvents = [];

    const tyreChangeTable = firstExistingTable(["tyre_change_slips", "tire_change_slips"]);
    if (tyreChangeTable) {
      const tcDateCol = firstExistingColumn(tyreChangeTable, ["slip_date", "change_date", "date", "created_at"]);
      const tcPositionCol = firstExistingColumn(tyreChangeTable, ["position", "wheel_position", "tyre_position"]);
      const tcOutCol = firstExistingColumn(tyreChangeTable, ["serial_out", "tyre_out", "tire_out"]);
      const tcInCol = firstExistingColumn(tyreChangeTable, ["serial_in", "tyre_in", "tire_in"]);
      const tcHoursCol = firstExistingColumn(tyreChangeTable, ["hours_at_change", "machine_hours", "hour_meter"]);
      const tcNotesCol = firstExistingColumn(tyreChangeTable, ["notes", "reason", "description"]);
      const tcPhotoCol = firstExistingColumn(tyreChangeTable, ["photo_path", "photo_url", "image_path", "image_url", "image"]);

      const tcDateExpr = tcDateCol ? `DATE(t.${tcDateCol})` : "DATE('now')";
      const tcF = dateFilter(tcDateExpr);
      const tcPositionExpr = tcPositionCol ? `t.${tcPositionCol}` : "NULL";
      const tcOutExpr = tcOutCol ? `t.${tcOutCol}` : "NULL";
      const tcInExpr = tcInCol ? `t.${tcInCol}` : "NULL";
      const tcHoursExpr = tcHoursCol ? `t.${tcHoursCol}` : "NULL";
      const tcNotesExpr = tcNotesCol ? `t.${tcNotesCol}` : "NULL";
      const tcPhotoExpr = tcPhotoCol ? `t.${tcPhotoCol}` : "NULL";

      tyreChangeEvents = db.prepare(`
        SELECT
          t.id,
          ${tcDateExpr} AS date,
          ${tcPositionExpr} AS position,
          ${tcOutExpr} AS serial_out,
          ${tcInExpr} AS serial_in,
          ${tcHoursExpr} AS hours_at_change,
          ${tcNotesExpr} AS notes,
          ${tcPhotoExpr} AS photo
        FROM ${tyreChangeTable} t
        WHERE t.asset_id = ? ${tcF.sql}
        ORDER BY date DESC, t.id DESC
        LIMIT 200
      `).all(asset.id, ...tcF.params).map((t) => ({
        type: "tyre_change",
        date: t.date,
        title: `Tyre Change #${t.id}${t.position ? ` (${t.position})` : ""}`,
        work_order_id: null,
        details: {
          slip_id: t.id,
          position: t.position || null,
          serial_out: t.serial_out || null,
          serial_in: t.serial_in || null,
          hours_at_change: t.hours_at_change != null ? Number(t.hours_at_change) : null,
          notes: t.notes || null,
          photo: t.photo || null,
        },
      }));
    }

    const tyreInspectTable = firstExistingTable(["tyre_inspections", "tire_inspections"]);
    if (tyreInspectTable) {
      const tiDateCol = firstExistingColumn(tyreInspectTable, ["inspection_date", "check_date", "date", "created_at"]);
      const tiPositionCol = firstExistingColumn(tyreInspectTable, ["position", "wheel_position", "tyre_position"]);
      const tiConditionCol = firstExistingColumn(tyreInspectTable, ["condition", "status", "result"]);
      const tiPressureCol = firstExistingColumn(tyreInspectTable, ["pressure_kpa", "pressure_psi", "pressure"]);
      const tiTreadCol = firstExistingColumn(tyreInspectTable, ["tread_depth", "tread_mm"]);
      const tiNotesCol = firstExistingColumn(tyreInspectTable, ["notes", "comment", "observation"]);
      const tiPhotoCol = firstExistingColumn(tyreInspectTable, ["photo_path", "photo_url", "image_path", "image_url", "image"]);

      const tiDateExpr = tiDateCol ? `DATE(t.${tiDateCol})` : "DATE('now')";
      const tiF = dateFilter(tiDateExpr);
      const tiPositionExpr = tiPositionCol ? `t.${tiPositionCol}` : "NULL";
      const tiConditionExpr = tiConditionCol ? `t.${tiConditionCol}` : "NULL";
      const tiPressureExpr = tiPressureCol ? `t.${tiPressureCol}` : "NULL";
      const tiTreadExpr = tiTreadCol ? `t.${tiTreadCol}` : "NULL";
      const tiNotesExpr = tiNotesCol ? `t.${tiNotesCol}` : "NULL";
      const tiPhotoExpr = tiPhotoCol ? `t.${tiPhotoCol}` : "NULL";

      tyreInspectionEvents = db.prepare(`
        SELECT
          t.id,
          ${tiDateExpr} AS date,
          ${tiPositionExpr} AS position,
          ${tiConditionExpr} AS condition,
          ${tiPressureExpr} AS pressure,
          ${tiTreadExpr} AS tread_depth,
          ${tiNotesExpr} AS notes,
          ${tiPhotoExpr} AS photo
        FROM ${tyreInspectTable} t
        WHERE t.asset_id = ? ${tiF.sql}
        ORDER BY date DESC, t.id DESC
        LIMIT 200
      `).all(asset.id, ...tiF.params).map((t) => ({
        type: "tyre_inspection",
        date: t.date,
        title: `Tyre Inspection #${t.id}${t.position ? ` (${t.position})` : ""}`,
        work_order_id: null,
        details: {
          inspection_id: t.id,
          position: t.position || null,
          condition: t.condition || null,
          pressure: t.pressure != null ? Number(t.pressure) : null,
          tread_depth: t.tread_depth != null ? Number(t.tread_depth) : null,
          notes: t.notes || null,
          photo: t.photo || null,
        },
      }));
    }

    // ---- OIL TOTALS (qty + optional cost)
    const oilF = dateFilter("o.log_date");
    const hasOilTotalCost = hasColumn("oil_logs", "total_cost");
    const hasOilUnitCost = hasColumn("oil_logs", "unit_cost");
    const oilCostExpr = hasOilTotalCost
      ? "COALESCE(SUM(o.total_cost), 0)"
      : hasOilUnitCost
      ? "COALESCE(SUM(o.quantity * o.unit_cost), 0)"
      : "0";

    const oilTotals = hasTable("oil_logs")
      ? db.prepare(`
          SELECT
            COALESCE(SUM(o.quantity), 0) AS oil_qty_total,
            ${oilCostExpr} AS oil_cost_total
          FROM oil_logs o
          WHERE o.asset_id = ? ${oilF.sql}
        `).get(asset.id, ...oilF.params)
      : { oil_qty_total: 0, oil_cost_total: 0 };

    // ---- PARTS USED + COST (via WO issue stock movements)
    const hasSmCreatedAt = hasColumn("stock_movements", "created_at");
    const hasSmMovementDate = hasColumn("stock_movements", "movement_date");
    const smDateCol = hasSmCreatedAt
      ? "DATE(sm.created_at)"
      : hasSmMovementDate
      ? "DATE(sm.movement_date)"
      : "DATE('now')";
    const smF = dateFilter(smDateCol);

    const hasSmUnitCost = hasColumn("stock_movements", "unit_cost");
    const hasSmTotalCost = hasColumn("stock_movements", "total_cost");
    const hasPartsUnitCost = hasColumn("parts", "unit_cost");

    const partCostExpr = hasSmTotalCost
      ? "COALESCE(SUM(ABS(sm.total_cost)), 0)"
      : hasSmUnitCost
      ? "COALESCE(SUM(ABS(sm.quantity) * sm.unit_cost), 0)"
      : hasPartsUnitCost
      ? "COALESCE(SUM(ABS(sm.quantity) * COALESCE(p.unit_cost, 0)), 0)"
      : "0";

    const partTotals = hasTable("stock_movements")
      ? db.prepare(`
          SELECT
            COALESCE(SUM(ABS(sm.quantity)), 0) AS parts_qty_total,
            ${partCostExpr} AS parts_cost_total
          FROM stock_movements sm
          JOIN work_orders w ON sm.reference = ('work_order:' || w.id)
          LEFT JOIN parts p ON p.id = sm.part_id
          WHERE w.asset_id = ?
            AND sm.movement_type = 'out'
            ${smF.sql}
        `).get(asset.id, ...smF.params)
      : { parts_qty_total: 0, parts_cost_total: 0 };

    const history = [
      ...breakdownEvents,
      ...workOrders,
      ...getSlipEvents,
      ...componentSlipEvents,
      ...opsSlipEvents,
      ...damageEvents,
      ...tyreChangeEvents,
      ...tyreInspectionEvents,
    ].sort((a, b) =>
      String(b.date).localeCompare(String(a.date))
    );

    const summary = {
      period_start: startOk ? start : null,
      period_end: endOk ? end : null,
      counts: {
        breakdowns: breakdownEvents.length,
        work_orders: workOrders.length,
        get_slips: getSlipEvents.length,
        component_slips: componentSlipEvents.length,
        ops_slips: opsSlipEvents.length,
        damage_reports: damageEvents.length,
        tyre_changes: tyreChangeEvents.length,
        tyre_inspections: tyreInspectionEvents.length,
        events_total: history.length,
      },
      totals: {
        parts_qty_total: Number(partTotals?.parts_qty_total || 0),
        parts_cost_total: Number(partTotals?.parts_cost_total || 0),
        oil_qty_total: Number(oilTotals?.oil_qty_total || 0),
        oil_cost_total: Number(oilTotals?.oil_cost_total || 0),
      },
      maintenance_cost_total:
        Number(partTotals?.parts_cost_total || 0) + Number(oilTotals?.oil_cost_total || 0),
    };

    return {
      ok: true,
      asset: {
        asset_code: asset.asset_code,
        asset_name: asset.asset_name,
        category: asset.category,
        archived: Number(asset.archived),
        archive_reason: asset.archive_reason,
        archived_at: asset.archived_at,
      },
      summary,
      history,
    };
  });
}