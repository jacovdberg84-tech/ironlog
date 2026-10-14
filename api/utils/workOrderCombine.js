// IRONLOG/api/utils/workOrderCombine.js — combine a machine's open pre-start work orders.
//
// Before pre-start faults joined the open job (see prestartFaults.js), every
// daily check with a fault opened its own work order, so one machine could
// collect ten. Combining keeps one of them (the one furthest along, else the
// oldest) and moves everything from the others onto it: fault lists, progress
// notes, parts requests, stock issued, time, photos, helpers and links. The
// others are closed with "Combined into WO #…" so their history still reads.
import { ensurePrestartLinkTable } from "./prestartFaults.js";

const DONE = ["completed", "approved", "closed"];
const STAGE = { in_progress: 3, assigned: 2, open: 1 };

function hasColumn(db, table, col) {
  try {
    return db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
  } catch {
    return false;
  }
}

/** Every table (other than work_orders) with a work_order_id column. */
function linkedTables(db) {
  return db.prepare(`SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%'`).all()
    .map((r) => r.name)
    .filter((t) => t !== "work_orders" && hasColumn(db, t, "work_order_id"));
}

/** The machine's unfinished pre-start work orders, the one to keep first. */
export function openPrestartWorkOrders(db, assetId) {
  const rows = db.prepare(`
    SELECT w.*, a.asset_code, a.asset_name FROM work_orders w
    LEFT JOIN assets a ON a.id = w.asset_id
    WHERE w.asset_id = ? AND w.source = 'prestart'
      AND LOWER(TRIM(COALESCE(w.status, 'open'))) NOT IN ('completed', 'approved', 'closed')
  `).all(Number(assetId));
  const stage = (s) => STAGE[String(s || "open").trim().toLowerCase()] || 1;
  return rows.sort((a, b) =>
    stage(b.status) - stage(a.status)
    || Number(Boolean(b.assigned_artisan_name)) - Number(Boolean(a.assigned_artisan_name))
    || a.id - b.id);
}

/** The faults on a job ("Engine oil: low; Horn"), else its first line. */
function firstLine(text) {
  const lines = String(text || "").split("\n").map((l) => l.trim()).filter(Boolean);
  const faults = lines.filter((l) => l.startsWith("- ")).map((l) => l.slice(2).replace(/\s+\(again\)$/, ""));
  return (faults.length ? [...new Set(faults)].join("; ") : lines[0] || "").slice(0, 140);
}

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const shortDate = (ymd) => `${Number(ymd.slice(8, 10))} ${MONTHS[Number(ymd.slice(5, 7)) - 1] || ""}`.trim();

/**
 * The kept job's description after combining: every fault once, oldest first,
 * with how often and when it was reported, then any other notes.
 */
export function combinedDescription(workOrders) {
  const faults = new Map();
  const other = [];
  const dates = [];
  const ordered = [...workOrders].sort((a, b) => String(a.opened_at || "").localeCompare(String(b.opened_at || "")) || a.id - b.id);
  for (const w of ordered) {
    let date = String(w.opened_at || "").slice(0, 10);
    for (const raw of String(w.job_description || "").split("\n")) {
      const line = raw.trim();
      if (!line) continue;
      const head = line.match(/(?:reported by|Updated by|Re-checked by) .+? on (\d{4}-\d{2}-\d{2})/i);
      if (head) {
        date = head[1];
        if (!/^Re-checked/i.test(line)) continue;
      }
      if (line.startsWith("- ")) {
        const fault = line.replace(/\s+\(again\)$/, "");
        if (!faults.has(fault)) faults.set(fault, new Set());
        if (date) { faults.get(fault).add(date); dates.push(date); }
        continue;
      }
      if (/^(From WO #\d+:|Combined pre-start faults)/.test(line)) continue;
      other.push(`WO #${w.id}: ${line}`);
    }
  }
  const ids = ordered.map((w) => w.id).sort((a, b) => a - b);
  const sortedDates = [...new Set(dates)].sort();
  const span = sortedDates.length ? ` reported ${sortedDates[0]}${sortedDates.length > 1 ? ` to ${sortedDates[sortedDates.length - 1]}` : ""}` : "";
  const lines = [`Combined pre-start faults from ${ids.length} work orders (${ids.map((id) => `#${id}`).join(", ")})${span}:`];
  for (const [fault, when] of faults) {
    const d = [...when].sort();
    lines.push(d.length > 1 ? `${fault} (reported ${d.length} times: ${d.map(shortDate).join(", ")})` : d.length ? `${fault} (${shortDate(d[0])})` : fault);
  }
  if (other.length) lines.push("", "Other notes:", ...other);
  return lines.join("\n");
}

/** What combining would do, without changing anything. */
export function previewCombine(db, workOrderId) {
  const wo = db.prepare(`SELECT id, asset_id, source, status FROM work_orders WHERE id = ?`).get(Number(workOrderId));
  if (!wo) return { ok: false, error: "Work order not found." };
  if (wo.source !== "prestart") return { ok: false, error: "Only pre-start work orders can be combined." };
  const rows = openPrestartWorkOrders(db, wo.asset_id);
  if (rows.length < 2) return { ok: false, error: "This machine has only one open pre-start work order." };
  const parts = (id) => (hasColumn(db, "maintenance_parts_requests", "work_order_id")
    ? Number(db.prepare(`SELECT COUNT(*) AS n FROM maintenance_parts_requests WHERE work_order_id = ?`).get(id).n) : 0);
  const describe = (r) => ({
    id: r.id,
    status: r.status,
    opened_at: r.opened_at || null,
    technician: r.assigned_artisan_name || null,
    summary: firstLine(r.job_description),
    parts_requests: parts(r.id),
  });
  return {
    ok: true,
    asset_code: rows[0].asset_code,
    asset_name: rows[0].asset_name,
    keep: describe(rows[0]),
    combine: rows.slice(1).map(describe),
  };
}

/**
 * Combine every open pre-start work order on the machine into one.
 * Runs in one transaction; returns { keep_id, combined_ids, moved }.
 */
export function combinePrestartWorkOrders(db, workOrderId, { actor = "admin" } = {}) {
  const preview = previewCombine(db, workOrderId);
  if (!preview.ok) return preview;
  ensurePrestartLinkTable(db);
  if (!hasColumn(db, "work_orders", "merged_into_id")) db.exec(`ALTER TABLE work_orders ADD COLUMN merged_into_id INTEGER`);
  const keepId = preview.keep.id;
  const others = preview.combine.map((r) => r.id);
  const marks = others.map(() => "?").join(",");
  const moved = {};

  db.transaction(() => {
    const all = db.prepare(`SELECT * FROM work_orders WHERE id IN (${[keepId, ...others].join(",")})`).all();
    const keep = all.find((w) => w.id === keepId);
    const rest = others.map((id) => all.find((w) => w.id === id));
    const sets = {};

    // One clean fault list: each fault once, with when it was reported.
    if (hasColumn(db, "work_orders", "job_description")) {
      sets.job_description = combinedDescription([keep, ...rest]);
    }
    if (hasColumn(db, "work_orders", "repair_progress")) {
      const notes = [keep, ...rest].filter((w) => String(w.repair_progress || "").trim())
        .map((w) => (w.id === keepId ? String(w.repair_progress).trim() : `WO #${w.id}: ${String(w.repair_progress).trim()}`));
      if (notes.length) sets.repair_progress = notes.join("\n");
    }
    // Labour and oil already booked on the others still count, on the kept job.
    for (const col of ["labor_hours", "oil_cost"]) {
      if (hasColumn(db, "work_orders", col)) sets[col] = [keep, ...rest].reduce((s, w) => s + Number(w[col] || 0), 0) || keep[col];
    }
    if (hasColumn(db, "work_orders", "labor_rate_per_hour") && keep.labor_rate_per_hour == null) {
      sets.labor_rate_per_hour = rest.find((w) => w.labor_rate_per_hour != null)?.labor_rate_per_hour ?? null;
    }
    if (hasColumn(db, "work_orders", "assigned_artisan_name") && !keep.assigned_artisan_name) {
      const tech = rest.find((w) => w.assigned_artisan_name);
      if (tech) {
        sets.assigned_artisan_name = tech.assigned_artisan_name;
        if (String(keep.status || "open").toLowerCase() === "open") sets.status = "assigned";
      }
    }
    const cols = Object.keys(sets);
    if (cols.length) {
      db.prepare(`UPDATE work_orders SET ${cols.map((c) => `${c} = ?`).join(", ")} WHERE id = ?`).run(...cols.map((c) => sets[c]), keepId);
    }

    // Everything that points at the others now points at the kept job. Rows
    // that would duplicate one already on it (a helper listed twice) stay put.
    for (const t of linkedTables(db)) {
      const n = db.prepare(`UPDATE OR IGNORE ${t} SET work_order_id = ? WHERE work_order_id IN (${marks})`).run(keepId, ...others).changes;
      if (n) moved[t] = n;
    }
    if (hasColumn(db, "breakdowns", "primary_work_order_id")) {
      const n = db.prepare(`UPDATE breakdowns SET primary_work_order_id = ? WHERE primary_work_order_id IN (${marks})`).run(keepId, ...others).changes;
      if (n) moved.breakdowns = n;
    }
    if (hasColumn(db, "stock_movements", "reference")) {
      const n = db.prepare(`UPDATE stock_movements SET reference = ? WHERE reference IN (${others.map(() => "?").join(",")})`)
        .run(`work_order:${keepId}`, ...others.map((id) => `work_order:${id}`)).changes;
      if (n) moved.stock_issued = n;
    }
    // Remember which checks each old job came from, so a resubmit finds the kept job.
    for (const w of rest) {
      db.prepare(`
        INSERT INTO work_order_prestart_checks (work_order_id, check_id) VALUES (?, ?)
        ON CONFLICT(check_id) DO UPDATE SET work_order_id = excluded.work_order_id
      `).run(keepId, Number(w.reference_id));
    }

    const note = `Combined into WO #${keepId} by ${actor}`;
    db.prepare(`
      UPDATE work_orders SET status = 'closed', closed_at = datetime('now'), merged_into_id = ?,
        completion_notes = ?
      WHERE id IN (${marks})
    `).run(keepId, note, ...others);
  })();

  return { ok: true, keep_id: keepId, combined_ids: others, moved };
}
