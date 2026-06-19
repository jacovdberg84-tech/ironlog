/** Standard hourmeter / odometer service grids (plant + LDV). */
export const STANDARD_SERVICE_GRIDS = [250, 500, 1000, 10000];

/** Parse hour interval from names like "500", "500H Service", "1000 hr". */
export function parseIntervalFromServiceName(name) {
  const s = String(name || "").trim();
  if (!s) return null;
  const m = s.match(/(\d+(?:\.\d+)?)\s*(?:h(?:r|our)?s?)?/i);
  if (!m) return null;
  const n = Number(m[1]);
  return Number.isFinite(n) && n > 0 ? n : null;
}

export function isLdvServiceAssetCode(assetCode) {
  return /^V(0[1-9]|1[0-5])AM$/i.test(String(assetCode || "").trim());
}

export function planIntervalHours(plan) {
  const interval = Number(plan?.interval_hours ?? plan?.intervalHours ?? 0);
  if (Number.isFinite(interval) && interval > 0) return interval;
  return parseIntervalFromServiceName(plan?.service_name ?? plan?.serviceName) || 0;
}

/** Map plan interval to a fixed service grid step. */
export function normalizeGridStep(intervalHours, assetCode = null) {
  const iv = Number(intervalHours) || 0;
  if (isLdvServiceAssetCode(assetCode)) {
    if (iv >= 5000 || iv === 10000) return 10000;
    if (iv > 0) return iv;
    return 10000;
  }
  if (STANDARD_SERVICE_GRIDS.includes(iv)) return iv;
  if (iv > 0 && iv <= 250) return 250;
  if (iv <= 500) return 500;
  if (iv <= 1000) return 1000;
  if (iv <= 10000) return 10000;
  return iv > 0 ? iv : 500;
}

/** Round hourmeter reading down to the nearest service milestone on the grid. */
export function snapToServiceMilestone(hours, gridStep) {
  const h = Math.max(0, Number(hours) || 0);
  const step = Math.max(1, Number(gridStep) || 1);
  return Number((Math.floor(h / step) * step).toFixed(2));
}

export function snapLastServiceHours(hours, intervalHours, assetCode = null) {
  const step = normalizeGridStep(intervalHours, assetCode);
  return snapToServiceMilestone(hours, step);
}

export function groupActivePlansByAsset(plans) {
  const map = new Map();
  for (const p of plans || []) {
    if (Number(p.active ?? 1) === 0) continue;
    const assetId = Number(p.asset_id ?? p.assetId ?? 0);
    if (!assetId) continue;
    if (!map.has(assetId)) map.set(assetId, []);
    map.get(assetId).push(p);
  }
  return map;
}

export function hasRotatingSchedule(plans) {
  const intervals = [...new Set((plans || []).map(planIntervalHours).filter((x) => x > 0))];
  return intervals.length >= 2;
}

function gridBaseInterval(plans, assetCode = null) {
  const intervals = [...new Set((plans || []).map(planIntervalHours).filter((x) => x > 0))]
    .map((iv) => normalizeGridStep(iv, assetCode))
    .sort((a, b) => a - b);
  if (!intervals.length) return isLdvServiceAssetCode(assetCode) ? 10000 : 500;
  return intervals[0];
}

/** Next service milestone on a fixed hourmeter grid. */
export function nextMilestoneHours(currentHours, baseInterval) {
  const cur = Math.max(0, Number(currentHours) || 0);
  const base = Math.max(1, Number(baseInterval) || 1);
  if (cur <= 0) return base;
  if (Math.abs(cur % base) < 1e-6) return cur + base;
  return Math.ceil(cur / base) * base;
}

/** Pick service type (250 / 500 / 1000 / 10000) at a grid milestone. */
export function milestoneServiceInterval(milestoneHours, intervals, assetCode = null) {
  const m = Number(milestoneHours);
  const unique = [...new Set(
    (intervals || [])
      .map((iv) => normalizeGridStep(iv, assetCode))
      .filter((x) => x > 0),
  )].sort((a, b) => b - a);
  if (!unique.length) return normalizeGridStep(500, assetCode);
  for (const iv of unique) {
    if (Math.abs(m % iv) < 1e-6) return iv;
  }
  return unique[unique.length - 1];
}

function findPlanForInterval(plans, targetInterval) {
  const target = Number(targetInterval);
  if (!Number.isFinite(target) || target <= 0) return null;
  let exact = null;
  let closest = null;
  let closestDiff = Infinity;
  for (const p of plans) {
    const iv = planIntervalHours(p);
    if (!iv) continue;
    if (Math.abs(iv - target) < 1e-6) {
      exact = p;
      break;
    }
    const diff = Math.abs(iv - target);
    if (diff < closestDiff) {
      closestDiff = diff;
      closest = p;
    }
  }
  return exact || closest;
}

export function resolveLegacyPlanDue(plan, currentHours, assetCode = null) {
  const current = Number(currentHours) || 0;
  const interval = planIntervalHours(plan);
  const gridStep = normalizeGridStep(interval, assetCode);
  const last = snapToServiceMilestone(Number(plan.last_service_hours || 0), gridStep);
  const nextDue = nextMilestoneHours(current, gridStep);
  const remaining = nextDue - current;
  return {
    schedule_mode: "grid",
    plan_id: Number(plan.id ?? plan.plan_id ?? 0),
    asset_id: Number(plan.asset_id || 0),
    service_name: String(plan.service_name || ""),
    interval_hours: interval,
    grid_step_hours: gridStep,
    last_service_hours: last,
    last_service_hours_raw: Number(plan.last_service_hours || 0),
    next_due_hours: Number(nextDue.toFixed(2)),
    remaining_hours: Number(remaining.toFixed(2)),
    next_service_interval: interval,
    is_next_for_asset: true,
  };
}

/** Pick the next service type from active plans on a shared hourmeter grid. */
export function resolveNextServiceForAssetPlans(plans, currentHours, assetCode = null) {
  const activePlans = (plans || []).filter((p) => Number(p.active ?? 1) !== 0);
  if (!activePlans.length) return null;

  const code = assetCode || String(activePlans[0]?.asset_code || "");
  const intervals = [...new Set(activePlans.map(planIntervalHours).filter((x) => x > 0))].sort((a, b) => a - b);
  if (intervals.length < 2) {
    return resolveLegacyPlanDue(activePlans[0], currentHours, code);
  }

  const baseInterval = gridBaseInterval(activePlans, code);
  const current = Number(currentHours) || 0;
  const nextDue = nextMilestoneHours(current, baseInterval);
  const serviceInterval = milestoneServiceInterval(nextDue, intervals, code);
  const matchedPlan = findPlanForInterval(activePlans, serviceInterval);
  const remaining = nextDue - current;
  const lastService = snapToServiceMilestone(Math.max(0, nextDue - serviceInterval), baseInterval);

  return {
    schedule_mode: "rotating",
    plan_id: Number(matchedPlan?.id ?? matchedPlan?.plan_id ?? activePlans[0]?.id ?? 0),
    asset_id: Number(activePlans[0].asset_id || 0),
    service_name: String(matchedPlan?.service_name || `${serviceInterval}`),
    interval_hours: serviceInterval,
    grid_step_hours: baseInterval,
    last_service_hours: lastService,
    next_due_hours: Number(nextDue.toFixed(2)),
    remaining_hours: Number(remaining.toFixed(2)),
    next_service_interval: serviceInterval,
    base_interval_hours: baseInterval,
    major_interval_hours: intervals[intervals.length - 1],
    is_next_for_asset: true,
  };
}

export function enrichPlansWithNextService(plans, getCurrentHours, defaultNearDue = 50) {
  const byAsset = groupActivePlansByAsset(plans);
  const nextByAsset = new Map();
  for (const [assetId, assetPlans] of byAsset) {
    const current = typeof getCurrentHours === "function" ? getCurrentHours(assetId) : 0;
    const code = String(assetPlans[0]?.asset_code || "");
    nextByAsset.set(assetId, resolveNextServiceForAssetPlans(assetPlans, current, code));
  }

  return (plans || []).map((plan) => {
    const assetId = Number(plan.asset_id || 0);
    const assetPlans = byAsset.get(assetId) || [];
    const current = typeof getCurrentHours === "function" ? getCurrentHours(assetId) : 0;
    const code = String(plan.asset_code || assetPlans[0]?.asset_code || "");
    const rotating = hasRotatingSchedule(assetPlans);
    const snappedLast = snapLastServiceHours(
      Number(plan.last_service_hours || 0),
      planIntervalHours(plan),
      code,
    );

    if (rotating) {
      const next = nextByAsset.get(assetId);
      const planId = Number(plan.id ?? plan.plan_id ?? 0);
      const isNext = Number(next?.plan_id || 0) === planId;
      const remaining = isNext ? next?.remaining_hours ?? null : null;
      const dueMeta = isNext && remaining != null
        ? classifyServiceDue(remaining, code, planIntervalHours(plan), defaultNearDue)
        : null;
      return {
        ...plan,
        last_service_hours_snapped: snappedLast,
        schedule_mode: "rotating",
        current_hours: Number(Number(current).toFixed(2)),
        next_due_hours: next?.next_due_hours ?? null,
        remaining_hours: remaining,
        is_next_for_asset: isNext,
        next_service_name: isNext ? next?.service_name : null,
        meter_unit: dueMeta?.meter_unit ?? meterUnitForAsset(code),
        near_due_threshold: dueMeta?.near_due_threshold ?? null,
        status: dueMeta?.status ?? null,
        is_almost_due: dueMeta?.is_almost_due ?? false,
      };
    }

    const legacy = resolveLegacyPlanDue(plan, current, code);
    const dueMeta = classifyServiceDue(legacy.remaining_hours, code, planIntervalHours(plan), defaultNearDue);
    return {
      ...plan,
      last_service_hours_snapped: snappedLast,
      schedule_mode: "grid",
      current_hours: Number(Number(current).toFixed(2)),
      next_due_hours: legacy.next_due_hours,
      remaining_hours: legacy.remaining_hours,
      is_next_for_asset: true,
      next_service_name: legacy.service_name,
      meter_unit: dueMeta.meter_unit,
      near_due_threshold: dueMeta.near_due_threshold,
      status: dueMeta.status,
      is_almost_due: dueMeta.is_almost_due,
    };
  });
}

export function meterUnitForAsset(assetCode) {
  return isLdvServiceAssetCode(assetCode) ? "km" : "hours";
}

/** Plant default 50h; LDV 10000 km cycle flags almost due within 500 km. */
export function nearDueThresholdForAsset(assetCode, serviceInterval = 0, defaultNear = 50) {
  const code = String(assetCode || "");
  const iv = normalizeGridStep(
    Number(serviceInterval) || (isLdvServiceAssetCode(code) ? 10000 : 500),
    code,
  );
  if (isLdvServiceAssetCode(code) && iv >= 10000) return 500;
  if (isLdvServiceAssetCode(code)) return Math.max(100, Number(defaultNear || 50));
  return Math.max(1, Number(defaultNear || 50));
}

export function classifyServiceDue(remaining, assetCode, serviceInterval = 0, defaultNear = 50) {
  const rem = Number(remaining || 0);
  const threshold = nearDueThresholdForAsset(assetCode, serviceInterval, defaultNear);
  let status = "OK";
  if (rem <= 0) status = "OVERDUE";
  else if (rem <= threshold) status = "ALMOST DUE";
  return {
    status,
    near_due_threshold: threshold,
    meter_unit: meterUnitForAsset(assetCode),
    is_overdue: rem <= 0,
    is_almost_due: rem > 0 && rem <= threshold,
  };
}

export function classifyServiceDueStatus(remaining, assetCode, serviceInterval = 0, defaultNear = 50) {
  return classifyServiceDue(remaining, assetCode, serviceInterval, defaultNear).status;
}

function attachDueMeta(row, assetCode, defaultNear = 50) {
  const interval = Number(row.next_service_interval || row.interval_hours || 0);
  const due = classifyServiceDue(row.remaining_hours, assetCode, interval, defaultNear);
  return {
    ...row,
    meter_unit: due.meter_unit,
    near_due_threshold: due.near_due_threshold,
    status: due.status,
    is_overdue: due.is_overdue,
    is_almost_due: due.is_almost_due,
  };
}

/** One due row per asset — uses rotating schedule when multiple intervals exist. */
export function buildDueListFromPlans(plans, getCurrentHours, defaultNearDue = 50) {
  const byAsset = groupActivePlansByAsset(plans);
  const rows = [];

  for (const [, assetPlans] of byAsset) {
    const sample = assetPlans[0];
    const assetId = Number(sample.asset_id || 0);
    const current = typeof getCurrentHours === "function" ? getCurrentHours(assetId) : 0;
    const resolved = resolveNextServiceForAssetPlans(
      assetPlans,
      current,
      String(sample.asset_code || ""),
    );
    if (!resolved) continue;

    const remaining = resolved.remaining_hours;
    rows.push(
      attachDueMeta(
        {
          plan_id: resolved.plan_id,
          asset_id: assetId,
          asset_code: sample.asset_code,
          asset_name: sample.asset_name,
          category: sample.category,
          service_name: resolved.service_name,
          interval_hours: resolved.interval_hours,
          last_service_hours: resolved.last_service_hours,
          grid_step_hours: resolved.grid_step_hours,
          current_hours: Number(Number(current).toFixed(2)),
          next_due_hours: resolved.next_due_hours,
          remaining_hours: resolved.remaining_hours,
          schedule_mode: resolved.schedule_mode,
          next_service_interval: resolved.next_service_interval,
          is_next_for_asset: true,
        },
        String(sample.asset_code || ""),
        defaultNearDue,
      ),
    );
  }

  return rows;
}
