import { getMachinePrestartTemplate } from "./machinePrestartTemplates.js";
import { andDailyHoursFleetHoursOnly } from "./fleetHoursKpiScope.js";

/** Default non-productive time per completed pre-start (30 minutes). Override via PRESTART_DEDUCTION_HOURS. */
export const PRESTART_DEDUCTION_HOURS = (() => {
  const n = Number(process.env.PRESTART_DEDUCTION_HOURS ?? 0.5);
  return Number.isFinite(n) && n > 0 ? n : 0.5;
})();

export function isPrestartCheckMode(checkMode) {
  const m = String(checkMode || "").trim().toLowerCase();
  return m === "prestart" || m.startsWith("machine_prestart_");
}

export function prestartTypeLabel(checkMode) {
  const m = String(checkMode || "").trim().toLowerCase();
  if (m === "prestart") return "LDV pre-start";
  const pfx = "machine_prestart_";
  if (m.startsWith(pfx)) {
    const id = m.slice(pfx.length).trim();
    const tmpl = id ? getMachinePrestartTemplate(id) : null;
    return tmpl?.title || "Machine pre-start";
  }
  return "Pre-start";
}

const PRESTART_MODE_SQL = `
  (
    LOWER(TRIM(COALESCE(v.check_mode, ''))) = 'prestart'
    OR LOWER(TRIM(COALESCE(v.check_mode, ''))) LIKE 'machine_prestart_%'
  )
`;

/**
 * Completed LDV + machine pre-starts for a calendar date.
 */
export function listDailyPrestarts(db, date) {
  const hasTable = Boolean(
    db.prepare(`SELECT 1 AS ok FROM sqlite_master WHERE type='table' AND name='vehicle_ldv_checks' LIMIT 1`).get()
  );
  if (!hasTable) {
    return { rows: [], deduction_hours_per_check: PRESTART_DEDUCTION_HOURS, total_deduction_hours: 0 };
  }

  const rows = db.prepare(`
    SELECT
      v.id AS check_id,
      v.check_date,
      v.check_mode,
      v.inspector_name,
      v.updated_at,
      v.created_at,
      a.id AS asset_id,
      a.asset_code,
      a.asset_name,
      a.category
    FROM vehicle_ldv_checks v
    JOIN assets a ON a.id = v.asset_id
    WHERE v.check_date = ?
      AND ${PRESTART_MODE_SQL}
      AND COALESCE(a.active, 1) = 1
    ORDER BY a.asset_code ASC, v.id ASC
  `).all(date).map((r) => ({
    check_id: Number(r.check_id),
    check_date: String(r.check_date || ""),
    check_mode: String(r.check_mode || ""),
    check_type: prestartTypeLabel(r.check_mode),
    inspector_name: r.inspector_name ? String(r.inspector_name) : null,
    completed_at: String(r.updated_at || r.created_at || ""),
    asset_id: Number(r.asset_id),
    asset_code: String(r.asset_code || ""),
    asset_name: String(r.asset_name || ""),
    category: r.category ? String(r.category) : null,
    deduction_hours: PRESTART_DEDUCTION_HOURS,
  }));

  return {
    rows,
    deduction_hours_per_check: PRESTART_DEDUCTION_HOURS,
    total_deduction_hours: Number((rows.length * PRESTART_DEDUCTION_HOURS).toFixed(4)),
  };
}

/**
 * Pre-start hours to subtract from fleet availability — only production hour-based rows on that date.
 */
export function prestartDeductionForProductionFleet(db, date) {
  const hasChecks = Boolean(
    db.prepare(`SELECT 1 AS ok FROM sqlite_master WHERE type='table' AND name='vehicle_ldv_checks' LIMIT 1`).get()
  );
  if (!hasChecks) return { count: 0, hours: 0, deduction_hours_per_check: PRESTART_DEDUCTION_HOURS };

  const row = db.prepare(`
    SELECT COUNT(DISTINCT v.asset_id) AS n
    FROM vehicle_ldv_checks v
    JOIN assets a ON a.id = v.asset_id
    JOIN daily_hours dh ON dh.asset_id = v.asset_id AND dh.work_date = ?
    WHERE v.check_date = ?
      AND ${PRESTART_MODE_SQL}
      ${andDailyHoursFleetHoursOnly("dh", "a")}
  `).get(date, date);

  const count = Number(row?.n || 0);
  const hours = Number((count * PRESTART_DEDUCTION_HOURS).toFixed(4));
  return { count, hours, deduction_hours_per_check: PRESTART_DEDUCTION_HOURS };
}
