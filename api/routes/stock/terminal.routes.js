// IRONLOG/api/routes/stock/terminal.routes.js — the stores counter terminal.
// Registered by routes/stock.routes.js; shared helpers arrive through ctx.
//
// A touch screen at the stores counter (the Toughbook). People sign in with a
// PIN; storemen issue, receive, count and work the workshop's requests;
// technicians collect parts for their own work orders. Scans arrive from a
// barcode scanner (typed into the page) or from a paired phone camera, which
// posts each code here for the terminal to pick up.

import { db } from "../../db/client.js";
import { writeAudit } from "../../utils/audit.js";
import { getRoles } from "../../utils/request.js";
import { technicianMatchesUser } from "../../utils/technicianIdentity.js";
import { jobLine, stockInfo } from "../../utils/techPortal.js";
import { workshopWaitingOnParts } from "../../utils/partsWaiting.js";

export const STORES_ROLES = ["admin", "supervisor", "stores", "storeman", "workshop_admin", "plant_manager", "site_manager"];
const TERMINAL_ROLES = [...STORES_ROLES, "artisan"];
const OPEN_WO = ["open", "assigned", "in_progress", "on_hold", "waiting_parts"];

// Phone-camera scans waiting for a terminal: key → [{ seq, code, at }].
const scanQueues = new Map();
let scanSeq = 0;
const SCAN_TTL_MS = 2 * 60 * 1000;
const MAX_SCAN_KEYS = 200; // a public endpoint: bound what strangers can make it hold

function pruneScans() {
  const cutoff = Date.now() - SCAN_TTL_MS;
  for (const [key, list] of scanQueues) {
    const kept = list.filter((s) => s.at >= cutoff);
    if (kept.length) scanQueues.set(key, kept);
    else scanQueues.delete(key);
  }
}

const validKey = (k) => /^[a-f0-9]{32}$/.test(String(k || ""));
const userOf = (req) => String(req.headers["x-user-name"] || "").trim();
const isStores = (req) => getRoles(req, { fallback: "admin" }).some((r) => STORES_ROLES.includes(r));

/** What a scanned or typed code is: a stock item, a machine or a work order. */
export function lookupCode(database, raw) {
  let code = String(raw || "").trim();
  if (!code) return { kind: "none" };
  // QR labels hold links: pull the identifying value out of them.
  try {
    if (/^https?:\/\//i.test(code)) {
      const u = new URL(code);
      const p = u.searchParams;
      if (p.get("part_code")) code = `PART:${p.get("part_code")}`;
      else if (p.get("asset_code")) code = `ASSET:${p.get("asset_code")}`;
      else if (p.get("wo") || p.get("work_order_id")) code = `WO:${p.get("wo") || p.get("work_order_id")}`;
      else if (/\/(asset|qr)\//i.test(u.pathname)) code = `ASSET:${decodeURIComponent(u.pathname.split("/").pop())}`;
    }
  } catch { /* not a link */ }
  const part = (c) => database.prepare(`SELECT id, part_code, part_name, min_stock FROM parts WHERE UPPER(part_code) = UPPER(?)`).get(c);
  const asset = (c) => database.prepare(`SELECT id, asset_code, asset_name FROM assets WHERE UPPER(asset_code) = UPPER(?)`).get(c);
  const wo = (n) => database.prepare(`SELECT id, asset_id, status FROM work_orders WHERE id = ?`).get(Number(n));
  const out = (kind, row) => ({ kind, ...row });

  const m = /^(PART|ASSET|WO):(.+)$/i.exec(code);
  if (m) {
    const [, kind, value] = m;
    if (kind.toUpperCase() === "PART") { const p = part(value); return p ? out("part", p) : { kind: "unknown", code: value }; }
    if (kind.toUpperCase() === "ASSET") { const a = asset(value); return a ? out("asset", a) : { kind: "unknown", code: value }; }
    const w = wo(String(value).replace(/\D/g, "")); return w ? out("work_order", w) : { kind: "unknown", code: value };
  }
  const p = part(code);
  if (p) return out("part", p);
  const a = asset(code);
  if (a) return out("asset", a);
  const wn = /^(?:WO\s*#?\s*)?(\d{1,7})$/i.exec(code);
  if (wn) { const w = wo(wn[1]); if (w) return out("work_order", w); }
  return { kind: "unknown", code };
}

export default function registerTerminalRoutes(app, ctx) {
  const { getOnHand, getPartByCode, insertAlloc, insertMove, requireRoles } = ctx;

  /** Open work orders this person may collect parts for (storemen: all). */
  function openWorkOrders(req, q = "") {
    const rows = db.prepare(`
      SELECT w.id, w.asset_id, w.status, w.source, w.assigned_artisan_name, w.job_description, w.opened_at,
        a.asset_code, a.asset_name, b.component AS breakdown_component, b.description AS breakdown_description, mp.service_name
      FROM work_orders w
      JOIN assets a ON a.id = w.asset_id
      LEFT JOIN breakdowns b ON b.id = w.reference_id AND w.source = 'breakdown'
      LEFT JOIN maintenance_plans mp ON mp.id = w.reference_id AND w.source = 'service'
      WHERE REPLACE(TRIM(LOWER(COALESCE(w.status, 'open'))), ' ', '_') IN (${OPEN_WO.map(() => "?").join(", ")})
      ORDER BY w.id DESC LIMIT 300
    `).all(...OPEN_WO);
    const me = userOf(req);
    const helperOf = db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'work_order_technicians'`).get()
      ? db.prepare(`SELECT 1 FROM work_order_technicians WHERE work_order_id = ? AND LOWER(username) = LOWER(?)`)
      : null;
    const mine = (w) => technicianMatchesUser(db, w.assigned_artisan_name, me) || Boolean(helperOf?.get(w.id, me));
    const term = String(q || "").trim().toUpperCase();
    return rows
      .filter((w) => isStores(req) || mine(w))
      .filter((w) => !term || String(w.id) === term.replace(/^WO\s*#?/, "") || String(w.asset_code).toUpperCase().includes(term))
      .slice(0, 40)
      .map((w) => ({ id: w.id, asset_code: w.asset_code, asset_name: w.asset_name, status: w.status, job: jobLine(w), lead: w.assigned_artisan_name || null, mine: mine(w) }));
  }

  /** May this person issue to this work order? */
  function mayIssueTo(req, wo) {
    if (isStores(req)) return true;
    const me = userOf(req);
    const w = db.prepare(`SELECT assigned_artisan_name FROM work_orders WHERE id = ?`).get(wo.id);
    if (technicianMatchesUser(db, w?.assigned_artisan_name, me)) return true;
    const helpers = db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'work_order_technicians'`).get();
    return Boolean(helpers && db.prepare(`SELECT 1 FROM work_order_technicians WHERE work_order_id = ? AND LOWER(username) = LOWER(?)`).get(wo.id, me));
  }

  // GET /api/stock/terminal/home — who is signed in and what is waiting.
  app.get("/terminal/home", async (req, reply) => {
    if (!requireRoles(req, reply, TERMINAL_ROLES)) return;
    const stores = isStores(req);
    const waiting = stores ? workshopWaitingOnParts(db, { site: String(req.headers["x-site-code"] || "main") }) : [];
    return {
      ok: true,
      user: userOf(req),
      stores,
      requests_waiting: waiting.filter((w) => w.kind === "request").length,
      requests_in_stock: waiting.filter((w) => w.kind === "request" && w.in_stock).length,
      my_jobs: stores ? null : openWorkOrders(req).length,
    };
  });

  // GET /api/stock/terminal/work-orders?q= — work orders to issue against.
  app.get("/terminal/work-orders", async (req, reply) => {
    if (!requireRoles(req, reply, TERMINAL_ROLES)) return;
    return { ok: true, rows: openWorkOrders(req, req.query?.q) };
  });

  // GET /api/stock/terminal/requests — the workshop's open parts requests (stores only).
  app.get("/terminal/requests", async (req, reply) => {
    if (!requireRoles(req, reply, STORES_ROLES)) return;
    const rows = workshopWaitingOnParts(db, { site: String(req.headers["x-site-code"] || "main") }).filter((r) => r.kind === "request");
    return { ok: true, rows };
  });

  // GET /api/stock/terminal/lookup?code= — a scanned or typed code.
  app.get("/terminal/lookup", async (req, reply) => {
    if (!requireRoles(req, reply, TERMINAL_ROLES)) return;
    const hit = lookupCode(db, req.query?.code);
    if (hit.kind === "part") {
      const s = stockInfo(db, [hit.id]).get(hit.id) || { on_hand: 0, bin: null };
      return { ok: true, ...hit, on_hand: s.on_hand, bin: s.bin };
    }
    if (hit.kind === "work_order") {
      const w = openWorkOrders(req).find((x) => x.id === hit.id);
      return { ok: true, ...hit, ...(w || {}), allowed: Boolean(w) };
    }
    return { ok: true, ...hit };
  });

  // POST /api/stock/terminal/issue { work_order_id | asset_code, lines: [{ part_code, quantity, request_id? }] }
  // All lines are issued together or none.
  app.post("/terminal/issue", async (req, reply) => {
    if (!requireRoles(req, reply, TERMINAL_ROLES)) return;
    const body = req.body || {};
    const me = userOf(req) || null;
    const woId = Number(body.work_order_id || 0) || null;
    let wo = null;
    let asset = null;
    if (woId) {
      wo = db.prepare(`SELECT id, asset_id, status FROM work_orders WHERE id = ?`).get(woId);
      if (!wo) return reply.code(404).send({ ok: false, error: `Work order #${woId} not found` });
      if (!mayIssueTo(req, wo)) return reply.code(403).send({ ok: false, error: "You can only collect parts for your own work orders" });
    } else if (body.asset_code) {
      if (!isStores(req)) return reply.code(403).send({ ok: false, error: "Choose one of your work orders" });
      asset = db.prepare(`SELECT id, asset_code FROM assets WHERE UPPER(asset_code) = UPPER(?)`).get(String(body.asset_code).trim());
      if (!asset) return reply.code(404).send({ ok: false, error: `Machine not found: ${body.asset_code}` });
    } else {
      return reply.code(400).send({ ok: false, error: "Choose the work order or machine" });
    }
    const assetId = wo ? Number(wo.asset_id) : Number(asset.id);

    const lines = [];
    const problems = [];
    for (const [i, l] of (Array.isArray(body.lines) ? body.lines : []).entries()) {
      const part = getPartByCode.get(String(l?.part_code || "").trim());
      const qty = Number(l?.quantity);
      if (!part) { problems.push(`Line ${i + 1}: part not found`); continue; }
      if (!Number.isFinite(qty) || qty <= 0) { problems.push(`${part.part_code}: quantity must be more than 0`); continue; }
      lines.push({ part, qty, requestId: Number(l?.request_id || 0) || null });
    }
    if (!lines.length && !problems.length) problems.push("Add at least one part");
    // Enough stock for every line (the same part on two lines counts once).
    const need = new Map();
    for (const l of lines) need.set(l.part.id, (need.get(l.part.id) || 0) + l.qty);
    for (const [pid, qty] of need) {
      const onHand = Number(getOnHand.get(pid)?.on_hand || 0);
      if (onHand < qty) {
        const p = lines.find((l) => l.part.id === pid).part;
        problems.push(`${p.part_code}: only ${onHand} in stock, ${qty} asked`);
      }
    }
    if (problems.length) return reply.code(400).send({ ok: false, error: problems[0], problems });

    const today = new Date().toISOString().slice(0, 10);
    const reference = wo ? `work_order:${wo.id}` : `asset:${assetId}:stores`;
    const issued = db.transaction(() => lines.map((l) => {
      const before = Number(getOnHand.get(l.part.id)?.on_hand || 0);
      insertMove.run(l.part.id, -Math.abs(l.qty), reference, null, null, null);
      insertAlloc.run(assetId, wo ? wo.id : null, l.part.id, l.qty, today, me, "Issued at the stores terminal", null, null, null);
      if (l.requestId) {
        db.prepare(`
          UPDATE maintenance_parts_requests
          SET status = 'received', status_notes = ?, updated_at = ?
          WHERE id = ? AND LOWER(COALESCE(status, 'requested')) IN ('requested', 'ordered')
        `).run(`Issued at the stores terminal by ${me || "stores"}`, new Date().toISOString(), l.requestId);
      }
      return { part_code: l.part.part_code, part_name: l.part.part_name, quantity: l.qty, on_hand_after: before - l.qty };
    }))();

    writeAudit(db, req, {
      module: "stock",
      action: "terminal_issue",
      entity_type: wo ? "work_order" : "asset",
      entity_id: String(wo ? wo.id : assetId),
      payload: { lines: issued.map((l) => `${l.part_code} x${l.quantity}`), by: me },
    });
    const a = db.prepare(`SELECT asset_code FROM assets WHERE id = ?`).get(assetId);
    return { ok: true, work_order_id: wo ? wo.id : null, asset_code: a?.asset_code || null, lines: issued };
  });

  // ---------------------------------------------------------------- phone as scanner
  // The terminal makes a random key and shows it as a QR; the phone opens the
  // scanner page with it and posts each code; the terminal collects them.

  // POST /api/stock/terminal/scan { key, code }  (public: the key is the secret)
  app.post("/terminal/scan", async (req, reply) => {
    const key = String(req.body?.key || "");
    const code = String(req.body?.code || "").trim().slice(0, 300);
    if (!validKey(key) || !code) return reply.code(400).send({ ok: false, error: "Bad scan" });
    pruneScans();
    if (!scanQueues.has(key) && scanQueues.size >= MAX_SCAN_KEYS) return reply.code(429).send({ ok: false, error: "Too many scanners, try again shortly" });
    const list = scanQueues.get(key) || [];
    scanSeq += 1;
    list.push({ seq: scanSeq, code, at: Date.now() });
    scanQueues.set(key, list.slice(-20));
    return { ok: true };
  });

  // GET /api/stock/terminal/scans?key=&after=  (public: the key is the secret)
  app.get("/terminal/scans", async (req, reply) => {
    const key = String(req.query?.key || "");
    if (!validKey(key)) return reply.code(400).send({ ok: false, error: "Bad key" });
    pruneScans();
    const after = Number(req.query?.after || 0);
    const rows = (scanQueues.get(key) || []).filter((s) => s.seq > after).map((s) => ({ seq: s.seq, code: s.code }));
    return { ok: true, rows, last: scanSeq };
  });
}
