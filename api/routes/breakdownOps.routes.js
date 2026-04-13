// IRONLOG/api/routes/breakdownOps.routes.js — Breakdown Ops slip reports (hose, GET, component, tyre)
import { db } from "../db/client.js";
import path from "node:path";
import {
  buildPdfBuffer,
  sectionTitle,
  kvLine,
  table,
  tryDrawLogo,
  ensurePageSpace,
} from "../utils/pdfGenerator.js";

const SLIP_PICTURE_ALLOWED_MIME = new Set(["image/jpeg", "image/png"]);
const SLIP_PICTURE_MAX_COUNT = 4;
const SLIP_PICTURE_MAX_BYTES = 512 * 1024;

function slipContentWidth(doc) {
  return doc.page.width - doc.page.margins.left - doc.page.margins.right;
}

/** Accept client { mime, data_base64 }[]; strip invalid / oversized. */
function normalizeSlipPictures(body) {
  const raw = body?.pictures;
  if (!Array.isArray(raw) || !raw.length) return [];
  const out = [];
  for (const pic of raw) {
    if (out.length >= SLIP_PICTURE_MAX_COUNT) break;
    const mime = String(pic?.mime || "").trim().toLowerCase();
    const b64 = String(pic?.data_base64 || "").replace(/\s/g, "");
    if (!SLIP_PICTURE_ALLOWED_MIME.has(mime) || !b64) continue;
    let buf;
    try {
      buf = Buffer.from(b64, "base64");
    } catch {
      continue;
    }
    if (!buf.length || buf.length > SLIP_PICTURE_MAX_BYTES) continue;
    out.push({ mime, data_base64: buf.toString("base64") });
  }
  return out;
}

function drawSlipPictures(doc, payload) {
  const pics = Array.isArray(payload?.pictures) ? payload.pictures : [];
  if (!pics.length) return;
  const margin = doc.page.margins.left;
  const maxW = slipContentWidth(doc);
  const maxH = 220;
  sectionTitle(doc, "Photos");
  for (const pic of pics) {
    ensurePageSpace(doc, maxH + 36);
    try {
      const buf = Buffer.from(String(pic.data_base64 || ""), "base64");
      if (!buf.length) continue;
      const y0 = doc.y;
      doc.image(buf, margin, y0, { fit: [maxW, maxH] });
      doc.y = y0 + maxH + 14;
    } catch {
      doc.font("Helvetica").fontSize(9).fillColor("#666666").text("(Photo could not be embedded.)", margin, doc.y);
      doc.moveDown(0.6);
    }
  }
}

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function getSiteCode(req) {
  return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
}

function getUser(req) {
  return String(req.headers["x-user-name"] || "user").trim() || "user";
}

function hasPartsTable() {
  try {
    return Boolean(
      db.prepare(`SELECT 1 FROM sqlite_master WHERE type='table' AND name='parts' LIMIT 1`).get()
    );
  } catch {
    return false;
  }
}

function lookupPart(code) {
  if (!hasPartsTable() || !String(code || "").trim()) return null;
  const row = db
    .prepare(
      `
    SELECT id, part_code, part_name, COALESCE(unit_cost, 0) AS unit_cost
    FROM parts
    WHERE UPPER(TRIM(part_code)) = UPPER(TRIM(?))
    LIMIT 1
  `
    )
    .get(String(code).trim());
  if (!row) return null;
  const unit_cost = Number(row.unit_cost || 0);
  return {
    id: Number(row.id),
    part_code: String(row.part_code || ""),
    part_name: String(row.part_name || ""),
    unit_cost,
    /** Same field — Stock On Hand treats parts.unit_cost as USD after receipt conversion. */
    unit_cost_usd: unit_cost,
  };
}

function compactCell(v, max = 200) {
  const s = String(v ?? "");
  return s.length > max ? `${s.slice(0, max)}…` : s;
}

/** Money on slips: Stores master uses USD (same as Stock On Hand list). */
function fmtUsd(v) {
  const n = Number(v);
  if (!Number.isFinite(n)) return "$0.00";
  return `$${n.toFixed(2)}`;
}

function slipTotalUsd(slipType, payload) {
  if (payload.slip_total_usd != null && Number.isFinite(Number(payload.slip_total_usd))) {
    return Number(payload.slip_total_usd);
  }
  if (slipType === "hose_failure") return Number(payload.hose_cost || 0) + Number(payload.oil_cost || 0);
  if (slipType === "get_change") {
    const a = Number(payload.part_line_est);
    const b = Number(payload.description_line_est);
    return (Number.isFinite(a) ? a : 0) + (Number.isFinite(b) ? b : 0);
  }
  if (slipType === "component_change") return Number(payload.cost || 0);
  if (slipType === "tyre_change") {
    return (Array.isArray(payload.tyres) ? payload.tyres : []).reduce((s, t) => s + Number(t.cost || 0), 0);
  }
  return 0;
}

function drawSlipUsdTotals(doc, slipType, payload) {
  sectionTitle(doc, "Totals (USD)");
  kvLine(
    doc,
    "Pricing basis",
    "Line amounts use Stores parts.unit_cost (USD per unit) × qty; manual overrides are USD line totals."
  );
  kvLine(doc, "This slip total", fmtUsd(slipTotalUsd(slipType, payload)));
}

const SLIP_TYPES = new Set(["hose_failure", "get_change", "component_change", "tyre_change"]);

function normSlipQty(v, def = 1) {
  const n = Number(v);
  if (!Number.isFinite(n) || n <= 0) return def;
  return Math.min(n, 99999);
}

function hasTable(name) {
  try {
    return Boolean(
      db.prepare(`SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1`).get(String(name || ""))
    );
  } catch {
    return false;
  }
}

function getAssetHoursInfoAsOf(assetId, asOfYmd) {
  const aid = Number(assetId || 0);
  const asOf = String(asOfYmd || "").trim();
  if (!aid || !/^\d{4}-\d{2}-\d{2}$/.test(asOf)) {
    return { hours: null, source: "unknown", as_of: null };
  }

  const meter = db.prepare(`
    SELECT closing_hours AS latest_closing
    FROM daily_hours
    WHERE asset_id = ?
      AND closing_hours IS NOT NULL
      AND work_date <= ?
    ORDER BY work_date DESC, id DESC
    LIMIT 1
  `).get(aid, asOf);
  const closing = meter?.latest_closing == null ? null : Number(meter.latest_closing);
  if (closing != null && Number.isFinite(closing)) {
    return { hours: closing, source: "daily_closing", as_of: asOf };
  }

  const fromDailyHours = db.prepare(`
    SELECT COALESCE(SUM(hours_run), 0) AS total_hours
    FROM daily_hours
    WHERE asset_id = ?
      AND is_used = 1
      AND hours_run > 0
      AND work_date <= ?
  `).get(aid, asOf);
  const total = Number(fromDailyHours?.total_hours || 0);
  if (Number.isFinite(total) && total > 0) {
    return { hours: total, source: "daily_sum", as_of: asOf };
  }

  const fallback = db.prepare(`
    SELECT hours
    FROM assets
    WHERE id = ?
    LIMIT 1
  `).get(aid);
  const assetHours = fallback?.hours == null ? null : Number(fallback.hours);
  if (assetHours != null && Number.isFinite(assetHours)) {
    return { hours: assetHours, source: "asset_hours", as_of: asOf };
  }

  return { hours: null, source: "unknown", as_of: asOf };
}

export default async function breakdownOpsRoutes(app) {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS ops_slip_reports (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL DEFAULT 'main',
      slip_type TEXT NOT NULL,
      asset_id INTEGER NOT NULL,
      report_date TEXT NOT NULL,
      payload_json TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      created_by TEXT,
      FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE RESTRICT
    )
  `).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_ops_slip_site_date ON ops_slip_reports(site_code, report_date)`).run();
  db.prepare(`CREATE INDEX IF NOT EXISTS idx_ops_slip_type ON ops_slip_reports(slip_type)`).run();

  app.get("/parts/lookup", async (req, reply) => {
    const part_code = String(req.query?.part_code || req.query?.q || "").trim();
    if (!part_code) return reply.code(400).send({ ok: false, error: "part_code required" });
    const p = lookupPart(part_code);
    if (!p) return reply.send({ ok: true, found: false, part: null });
    return reply.send({
      ok: true,
      found: true,
      part: { ...p, currency: "USD", note: "unit_cost is USD (Stores master, same as stock list)." },
    });
  });

  /** Latest GET / hose lines from asset register (get_change_slips) or last ops slip for this site. */
  app.get("/slip-asset-hints", async (req, reply) => {
    const asset_code = String(req.query?.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code required" });
    const asset = db
      .prepare(`SELECT id, asset_code, asset_name FROM assets WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?)) LIMIT 1`)
      .get(asset_code);
    if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });
    const aid = Number(asset.id);
    const site_code = getSiteCode(req);

    let get_change = null;
    let get_change_source = null;
    if (hasTable("get_change_slips") && hasTable("get_change_items")) {
      const slip = db
        .prepare(
          `SELECT id, slip_date, notes, location FROM get_change_slips WHERE asset_id = ? ORDER BY slip_date DESC, id DESC LIMIT 1`
        )
        .get(aid);
      if (slip) {
        const items = db
          .prepare(
            `SELECT position, part_code, part_name, qty, reason FROM get_change_items WHERE slip_id = ? ORDER BY id ASC LIMIT 2`
          )
          .all(slip.id);
        if (items.length) {
          const a = items[0];
          const b = items[1] || null;
          get_change = {
            hours_fitted: null,
            part_code: String(a.part_code || "").trim(),
            part_qty: normSlipQty(a.qty, 1),
            supplier: String(slip.location || "").trim() || undefined,
            date_changed: String(slip.slip_date || "").trim(),
            description_part_code: b ? String(b.part_code || "").trim() : "",
            description_part_qty: b ? normSlipQty(b.qty, 1) : 1,
            notes: String(slip.notes || "").trim() || undefined,
          };
          get_change_source = "get_change_slip";
        }
      }
    }
    if (!get_change) {
      const row = db
        .prepare(
          `SELECT payload_json, report_date FROM ops_slip_reports
           WHERE site_code = ? AND asset_id = ? AND slip_type = 'get_change'
           ORDER BY report_date DESC, id DESC LIMIT 1`
        )
        .get(site_code, aid);
      if (row) {
        try {
          const p = JSON.parse(String(row.payload_json || "{}"));
          get_change = {
            hours_fitted: p.hours_fitted != null && p.hours_fitted !== "" ? Number(p.hours_fitted) : null,
            part_code: String(p.part_code || "").trim(),
            part_qty: normSlipQty(p.part_qty, 1),
            supplier: String(p.supplier || "").trim() || undefined,
            date_changed: String(p.date_changed || row.report_date || "").trim(),
            description_part_code: String(p.description_part_code || "").trim(),
            description_part_qty: normSlipQty(p.description_part_qty, 1),
            notes: p.notes ? String(p.notes).trim() : undefined,
          };
          get_change_source = "last_ops_slip";
        } catch {
          /* ignore */
        }
      }
    }

    let hose_failure = null;
    let hose_failure_source = null;
    const hr = db
      .prepare(
        `SELECT payload_json, report_date FROM ops_slip_reports
         WHERE site_code = ? AND asset_id = ? AND slip_type = 'hose_failure'
         ORDER BY report_date DESC, id DESC LIMIT 1`
      )
      .get(site_code, aid);
    if (hr) {
      try {
        const p = JSON.parse(String(hr.payload_json || "{}"));
        hose_failure = {
          date_fitted: String(p.date_fitted || hr.report_date || "").trim(),
          reason_fitted: String(p.reason_fitted || "").trim(),
          preventable: Boolean(p.preventable),
          hose_part_code: String(p.hose_part_code || "").trim(),
          hose_qty: normSlipQty(p.hose_qty, 1),
          oil_loss_part_code: String(p.oil_loss_part_code || "").trim(),
          oil_loss_qty: normSlipQty(p.oil_loss_qty, 1),
          notes: p.notes ? String(p.notes).trim() : undefined,
        };
        hose_failure_source = "last_ops_slip";
      } catch {
        /* ignore */
      }
    }

    return {
      ok: true,
      asset_code: asset.asset_code,
      asset_name: asset.asset_name,
      get_change,
      get_change_source,
      hose_failure,
      hose_failure_source,
    };
  });

  function normalizePayload(slip_type, body) {
    if (slip_type === "hose_failure") {
      return {
        date_fitted: String(body.date_fitted || "").trim(),
        reason_fitted: String(body.reason_fitted || "").trim(),
        preventable: Boolean(body.preventable),
        hose_part_code: String(body.hose_part_code || "").trim(),
        oil_loss_part_code: String(body.oil_loss_part_code || "").trim(),
        hose_qty: normSlipQty(body.hose_qty, 1),
        oil_loss_qty: normSlipQty(body.oil_loss_qty, 1),
        hose_cost_manual: body.hose_cost_manual != null && body.hose_cost_manual !== "" ? Number(body.hose_cost_manual) : null,
        oil_cost_manual: body.oil_cost_manual != null && body.oil_cost_manual !== "" ? Number(body.oil_cost_manual) : null,
        notes: String(body.notes || "").trim() || null,
      };
    }
    if (slip_type === "get_change") {
      return {
        hours_fitted: body.hours_fitted != null && body.hours_fitted !== "" ? Number(body.hours_fitted) : null,
        part_code: String(body.part_code || "").trim(),
        part_qty: normSlipQty(body.part_qty, 1),
        supplier: String(body.supplier || "").trim(),
        date_changed: String(body.date_changed || "").trim(),
        description_part_code: String(body.description_part_code || "").trim(),
        description_part_qty: normSlipQty(body.description_part_qty, 1),
        notes: String(body.notes || "").trim() || null,
      };
    }
    if (slip_type === "component_change") {
      return {
        date_changed: String(body.date_changed || "").trim(),
        hours_in_service: body.hours_in_service != null && body.hours_in_service !== "" ? Number(body.hours_in_service) : null,
        reason: String(body.reason || "").trim(),
        component_type: String(body.component_type || "").trim(),
        part_code: String(body.part_code || "").trim(),
        cost_manual: body.cost_manual != null && body.cost_manual !== "" ? Number(body.cost_manual) : null,
        notes: String(body.notes || "").trim() || null,
      };
    }
    if (slip_type === "tyre_change") {
      const raw = Array.isArray(body.tyres) ? body.tyres : [];
      const tyres = raw
        .slice(0, 10)
        .map((t) => ({
          position: String(t?.position || "").trim(),
          serial_removed: String(t?.serial_removed || "").trim(),
          serial_new: String(t?.serial_new || "").trim(),
          tread_left: String(t?.tread_left || "").trim(),
          reason: String(t?.reason || "").trim(),
          hours_in_use: t?.hours_in_use != null && t?.hours_in_use !== "" ? Number(t.hours_in_use) : null,
          hours_fitted: t?.hours_fitted != null && t?.hours_fitted !== "" ? Number(t.hours_fitted) : null,
          part_code: String(t?.part_code || "").trim(),
          cost_manual: t?.cost_manual != null && t?.cost_manual !== "" ? Number(t.cost_manual) : null,
          tyre_make_part_code: String(t?.tyre_make_part_code || "").trim(),
        }))
        .filter(
          (t) =>
            t.position ||
            t.serial_removed ||
            t.serial_new ||
            t.part_code ||
            t.tyre_make_part_code ||
            t.reason
        );
      return { tyres, notes: String(body.notes || "").trim() || null };
    }
    return null;
  }

  function enrichPayload(slip_type, payload) {
    const resolved = {};
    if (slip_type === "hose_failure") {
      const hose = lookupPart(payload.hose_part_code);
      const oil = lookupPart(payload.oil_loss_part_code);
      resolved.hose = hose;
      resolved.oil_loss = oil;
      payload._resolved = resolved;
      const hq = normSlipQty(payload.hose_qty, 1);
      const oq = normSlipQty(payload.oil_loss_qty, 1);
      payload.hose_qty = hq;
      payload.oil_loss_qty = oq;
      const hoseMan = payload.hose_cost_manual != null && Number.isFinite(payload.hose_cost_manual);
      const oilMan = payload.oil_cost_manual != null && Number.isFinite(payload.oil_cost_manual);
      payload.hose_cost_is_manual = hoseMan;
      payload.oil_cost_is_manual = oilMan;
      payload.hose_cost = hoseMan ? Number(payload.hose_cost_manual) : hose ? Number(hose.unit_cost || 0) * hq : 0;
      payload.oil_cost = oilMan ? Number(payload.oil_cost_manual) : oil ? Number(oil.unit_cost || 0) * oq : 0;
      payload.slip_total_usd = Number(payload.hose_cost) + Number(payload.oil_cost);
    } else if (slip_type === "get_change") {
      const p = lookupPart(payload.part_code);
      const d = lookupPart(payload.description_part_code);
      payload.part_qty = normSlipQty(payload.part_qty, 1);
      payload.description_part_qty = normSlipQty(payload.description_part_qty, 1);
      payload._resolved = { part: p, description: d };
      const pq = payload.part_qty;
      const dq = payload.description_part_qty;
      payload.part_line_est = p ? Number(p.unit_cost || 0) * pq : null;
      payload.description_line_est = d ? Number(d.unit_cost || 0) * dq : null;
      const a = payload.part_line_est;
      const b = payload.description_line_est;
      payload.slip_total_usd = (Number.isFinite(a) ? a : 0) + (Number.isFinite(b) ? b : 0);
    } else if (slip_type === "component_change") {
      const p = lookupPart(payload.part_code);
      payload._resolved = { part: p };
      const costMan = payload.cost_manual != null && Number.isFinite(payload.cost_manual);
      payload.cost_is_manual = costMan;
      payload.cost = costMan ? Number(payload.cost_manual) : p ? Number(p.unit_cost || 0) : 0;
      payload.slip_total_usd = Number(payload.cost || 0);
    } else if (slip_type === "tyre_change") {
      payload.tyres = (payload.tyres || []).map((t) => {
        const part = lookupPart(t.part_code);
        const make = lookupPart(t.tyre_make_part_code);
        const costMan = t.cost_manual != null && Number.isFinite(t.cost_manual);
        const cost = costMan ? Number(t.cost_manual) : part ? Number(part.unit_cost || 0) : 0;
        return {
          ...t,
          cost,
          cost_is_manual: costMan,
          _part: part,
          _make: make,
        };
      });
      payload.slip_total_usd = payload.tyres.reduce((s, t) => s + Number(t.cost || 0), 0);
    }
    return payload;
  }

  app.post("/slips", async (req, reply) => {
    const slip_type = String(req.body?.slip_type || "").trim();
    if (!SLIP_TYPES.has(slip_type)) {
      return reply.code(400).send({ ok: false, error: `slip_type must be one of: ${[...SLIP_TYPES].join(", ")}` });
    }
    const asset_code = String(req.body?.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ ok: false, error: "asset_code is required" });
    const report_date = String(req.body?.report_date || "").trim();
    if (!isDate(report_date)) return reply.code(400).send({ ok: false, error: "report_date must be YYYY-MM-DD" });

    const asset = db
      .prepare(`SELECT id, asset_code, asset_name FROM assets WHERE UPPER(TRIM(asset_code)) = UPPER(TRIM(?)) LIMIT 1`)
      .get(asset_code);
    if (!asset) return reply.code(404).send({ ok: false, error: "Asset not found" });

    let payload = normalizePayload(slip_type, req.body || {});
    if (!payload) return reply.code(400).send({ ok: false, error: "Invalid payload for slip type" });
    if (slip_type === "tyre_change" && (!payload.tyres || !payload.tyres.length)) {
      return reply.code(400).send({ ok: false, error: "At least one tyre line is required" });
    }

    payload.pictures = normalizeSlipPictures(req.body || {});
    payload = enrichPayload(slip_type, payload);
    if (payload.current_hours == null || payload.current_hours === "") {
      const hrsInfo = getAssetHoursInfoAsOf(asset.id, report_date);
      if (Number.isFinite(Number(hrsInfo.hours))) {
        payload.current_hours = Number(Number(hrsInfo.hours).toFixed(1));
      } else {
        payload.current_hours = null;
      }
      payload.current_hours_source = hrsInfo.source || "unknown";
      payload.current_hours_as_of = hrsInfo.as_of || report_date;
    }
    const site_code = getSiteCode(req);
    const created_by = getUser(req);

    const ins = db
      .prepare(
        `
      INSERT INTO ops_slip_reports (site_code, slip_type, asset_id, report_date, payload_json, created_by)
      VALUES (?, ?, ?, ?, ?, ?)
    `
      )
      .run(site_code, slip_type, asset.id, report_date, JSON.stringify(payload), created_by);

    const id = Number(ins.lastInsertRowid);
    return reply.send({ ok: true, id, slip_type, asset_id: asset.id });
  });

  app.get("/slips", async (req) => {
    const site_code = getSiteCode(req);
    const from = String(req.query?.from || "").trim();
    const to = String(req.query?.to || "").trim();
    const st = String(req.query?.slip_type || "").trim();
    const params = [site_code];
    let where = "r.site_code = ?";
    if (isDate(from)) {
      where += " AND r.report_date >= ?";
      params.push(from);
    }
    if (isDate(to)) {
      where += " AND r.report_date <= ?";
      params.push(to);
    }
    if (st && SLIP_TYPES.has(st)) {
      where += " AND r.slip_type = ?";
      params.push(st);
    }
    const rows = db
      .prepare(
        `
      SELECT
        r.id,
        r.slip_type,
        r.report_date,
        r.created_at,
        r.created_by,
        a.asset_code,
        a.asset_name
      FROM ops_slip_reports r
      JOIN assets a ON a.id = r.asset_id
      WHERE ${where}
      ORDER BY r.report_date DESC, r.id DESC
      LIMIT 200
    `
      )
      .all(...params);
    return { ok: true, rows };
  });

  app.get("/slips/:id", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
    const site_code = getSiteCode(req);
    const row = db
      .prepare(
        `
      SELECT r.*, a.asset_code, a.asset_name
      FROM ops_slip_reports r
      JOIN assets a ON a.id = r.asset_id
      WHERE r.id = ? AND r.site_code = ?
    `
      )
      .get(id, site_code);
    if (!row) return reply.code(404).send({ ok: false, error: "Not found" });
    let payload = {};
    try {
      payload = JSON.parse(String(row.payload_json || "{}"));
    } catch {}
    return {
      ok: true,
      row: { ...row, payload },
    };
  });

  function drawSlipPdf(doc, row) {
    const logoPath = path.join(process.cwd(), "branding", "logo.png");
    tryDrawLogo(doc, logoPath);

    let payload = {};
    try {
      payload = JSON.parse(String(row.payload_json || "{}"));
    } catch {}

    const slipType = String(row.slip_type || "");
    const titles = {
      hose_failure: "Hose Failure Report",
      get_change: "G.E.T. (Ground Engaging Tools) Change Slip",
      component_change: "Component Change Slip",
      tyre_change: "Tyre Change Slip",
    };
    const title = titles[slipType] || "Ops Slip Report";

    sectionTitle(doc, title);
    kvLine(doc, "Report #", String(row.id));
    kvLine(doc, "Date", String(row.report_date || ""));
    kvLine(doc, "Asset", `${row.asset_code || ""} — ${row.asset_name || ""}`);
    kvLine(doc, "Recorded by", String(row.created_by || "-"));
    if (payload.current_hours != null && payload.current_hours !== "") {
      const src = String(payload.current_hours_source || "").trim();
      const asOf = String(payload.current_hours_as_of || row.report_date || "").trim();
      const bits = [`${Number(payload.current_hours).toFixed(1)} h`];
      if (asOf) bits.push(`as of ${asOf}`);
      if (src) bits.push(src);
      kvLine(doc, "Current hours", bits.join(" · "));
    }

    if (slipType === "hose_failure") {
      sectionTitle(doc, "Details");
      kvLine(doc, "Date fitted", payload.date_fitted || "—");
      kvLine(doc, "Reason fitted", compactCell(payload.reason_fitted, 400) || "—");
      kvLine(doc, "Failure preventable", payload.preventable ? "Yes" : "No");
      kvLine(doc, "Hose part (stores)", payload.hose_part_code || "—");
      const hq = normSlipQty(payload.hose_qty, 1);
      kvLine(doc, "Hose qty", String(hq));
      if (payload._resolved?.hose) {
        kvLine(doc, "Hose description", payload._resolved.hose.part_name || "—");
        if (payload.hose_cost_is_manual !== true) {
          kvLine(doc, "Hose unit (Stores USD)", fmtUsd(payload._resolved.hose.unit_cost));
        }
      }
      kvLine(
        doc,
        "Hose line total (USD)",
        payload.hose_cost_is_manual === true ? `${fmtUsd(payload.hose_cost)} (manual)` : fmtUsd(payload.hose_cost)
      );
      kvLine(doc, "Oil loss part (stores)", payload.oil_loss_part_code || "—");
      const oq = normSlipQty(payload.oil_loss_qty, 1);
      kvLine(doc, "Oil loss qty", String(oq));
      if (payload._resolved?.oil_loss) {
        kvLine(doc, "Oil part description", payload._resolved.oil_loss.part_name || "—");
        if (payload.oil_cost_is_manual !== true) {
          kvLine(doc, "Oil unit (Stores USD)", fmtUsd(payload._resolved.oil_loss.unit_cost));
        }
      }
      kvLine(
        doc,
        "Oil loss line total (USD)",
        payload.oil_cost_is_manual === true ? `${fmtUsd(payload.oil_cost)} (manual)` : fmtUsd(payload.oil_cost)
      );
      if (payload.notes) {
        sectionTitle(doc, "Notes");
        doc.font("Helvetica").fontSize(10).fillColor("#111").text(compactCell(payload.notes, 2000), {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });
      }
    } else if (slipType === "get_change") {
      sectionTitle(doc, "G.E.T. details");
      kvLine(doc, "Hours fitted", payload.hours_fitted != null ? String(payload.hours_fitted) : "—");
      kvLine(doc, "Part number (stores)", payload.part_code || "—");
      kvLine(doc, "Part qty", String(normSlipQty(payload.part_qty, 1)));
      if (payload._resolved?.part) {
        kvLine(doc, "Part name", payload._resolved.part.part_name || "—");
        kvLine(doc, "Part unit (Stores USD)", fmtUsd(payload._resolved.part.unit_cost));
      }
      if (payload.part_line_est != null && Number.isFinite(payload.part_line_est)) {
        kvLine(doc, "Part line total (USD est.)", fmtUsd(payload.part_line_est));
      }
      kvLine(doc, "Supplier", payload.supplier || "—");
      kvLine(doc, "Date changed", payload.date_changed || "—");
      kvLine(doc, "Description part (stores)", payload.description_part_code || "—");
      if (payload.description_part_code) {
        kvLine(doc, "Description part qty", String(normSlipQty(payload.description_part_qty, 1)));
      }
      if (payload._resolved?.description) {
        kvLine(doc, "Description (from stores)", payload._resolved.description.part_name || "—");
        kvLine(doc, "Description unit (Stores USD)", fmtUsd(payload._resolved.description.unit_cost));
      } else if (payload.description_part_code) {
        kvLine(doc, "Description part", "(code not in Stores master)");
      }
      if (
        payload.description_part_code &&
        payload.description_line_est != null &&
        Number.isFinite(payload.description_line_est)
      ) {
        kvLine(doc, "Description line total (USD est.)", fmtUsd(payload.description_line_est));
      }
      if (payload.notes) {
        sectionTitle(doc, "Notes");
        doc.font("Helvetica").fontSize(10).text(compactCell(payload.notes, 2000), {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });
      }
    } else if (slipType === "component_change") {
      sectionTitle(doc, "Component change");
      kvLine(doc, "Date changed", payload.date_changed || "—");
      kvLine(doc, "Hours in service", payload.hours_in_service != null ? String(payload.hours_in_service) : "—");
      kvLine(doc, "Reason for changing", compactCell(payload.reason, 300) || "—");
      kvLine(doc, "Component type", payload.component_type || "—");
      kvLine(doc, "Part number (stores)", payload.part_code || "—");
      if (payload._resolved?.part) {
        kvLine(doc, "Part name", payload._resolved.part.part_name || "—");
        if (payload.cost_is_manual !== true) {
          kvLine(doc, "Part unit (Stores USD)", fmtUsd(payload._resolved.part.unit_cost));
        }
      }
      kvLine(
        doc,
        "Line total (USD)",
        payload.cost_is_manual === true ? `${fmtUsd(payload.cost)} (manual)` : fmtUsd(payload.cost)
      );
      if (payload.notes) {
        sectionTitle(doc, "Notes");
        doc.font("Helvetica").fontSize(10).text(compactCell(payload.notes, 2000), {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });
      }
    } else if (slipType === "tyre_change") {
      sectionTitle(doc, "Tyres");
      const tyres = Array.isArray(payload.tyres) ? payload.tyres : [];
      if (!tyres.length) {
        doc.font("Helvetica").fontSize(10).text("No tyre lines.");
      } else {
        const cols = [
          { key: "pos", label: "Pos", width: 0.06 },
          { key: "serOut", label: "Serial out", width: 0.1 },
          { key: "serIn", label: "Serial in", width: 0.1 },
          { key: "tread", label: "Tread", width: 0.07 },
          { key: "reason", label: "Reason", width: 0.12 },
          { key: "hu", label: "Hrs use", width: 0.07 },
          { key: "hf", label: "Hrs fit", width: 0.07 },
          { key: "part", label: "Part", width: 0.1 },
          { key: "cost", label: "Line USD", width: 0.08 },
          { key: "make", label: "Make", width: 0.1 },
        ];
        const trows = tyres.map((t, i) => ({
          pos: t.position || String(i + 1),
          serOut: compactCell(t.serial_removed, 24),
          serIn: compactCell(t.serial_new, 24),
          tread: compactCell(t.tread_left, 12),
          reason: compactCell(t.reason, 40),
          hu: t.hours_in_use != null ? String(t.hours_in_use) : "",
          hf: t.hours_fitted != null ? String(t.hours_fitted) : "",
          part: compactCell(t.part_code, 20),
          cost: t.cost_is_manual === true ? `${fmtUsd(t.cost)}*` : fmtUsd(t.cost),
          make: compactCell(t.tyre_make_part_code, 16),
        }));
        table(doc, cols, trows, { compact: true, fontSize: 7 });
        if (tyres.some((t) => t.cost_is_manual === true)) {
          doc.moveDown(0.2);
          doc
            .font("Helvetica")
            .fontSize(8)
            .fillColor("#555555")
            .text("* Manual line amount (USD). Other lines use Stores parts.unit_cost (USD).", doc.page.margins.left, doc.y, {
              width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
            });
        }
      }
      if (payload.notes) {
        sectionTitle(doc, "Notes");
        doc.font("Helvetica").fontSize(10).text(compactCell(payload.notes, 2000), {
          width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        });
      }
    }
    drawSlipUsdTotals(doc, slipType, payload);
    drawSlipPictures(doc, payload);
  }

  app.get("/slips/:id/pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ error: "Invalid id" });
    const download = String(req.query?.download || "").trim() === "1";
    const site_code = getSiteCode(req);
    const row = db
      .prepare(
        `
      SELECT r.*, a.asset_code, a.asset_name
      FROM ops_slip_reports r
      JOIN assets a ON a.id = r.asset_id
      WHERE r.id = ? AND r.site_code = ?
    `
      )
      .get(id, site_code);
    if (!row) return reply.code(404).send({ error: "Not found" });

    const slipLabel = String(row.slip_type || "slip").replace(/_/g, "-");
    const pdf = await buildPdfBuffer(
      (doc) => drawSlipPdf(doc, row),
      {
        title: "IRONLOG",
        subtitle: "Breakdown Ops Slip",
        rightText: `#${id}`,
        showPageNumbers: true,
        layout: row.slip_type === "tyre_change" ? "landscape" : "portrait",
      }
    );

    return reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename="ops-slip-${slipLabel}-${id}.pdf"`
      )
      .send(pdf);
  });
}
