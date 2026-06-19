/** Parse hour interval from names like "500", "500H Service", "1000 hr". */
export function parseIntervalFromServiceName(name) {
  const s = String(name || "").trim();
  if (!s) return null;
  const m = s.match(/(\d+(?:\.\d+)?)\s*(?:h(?:r|our)?s?)?/i);
  if (!m) return null;
  const n = Number(m[1]);
  return Number.isFinite(n) && n > 0 ? n : null;
}

export function planIntervalHours(plan) {
  const interval = Number(plan?.interval_hours ?? plan?.intervalHours ?? 0);
  if (Number.isFinite(interval) && interval > 0) return interval;
  return parseIntervalFromServiceName(plan?.service_name ?? plan?.serviceName) || 0;
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

/** Next service milestone on a fixed hourmeter grid (e.g. every 500h). */
export function nextMilestoneHours(currentHours, baseInterval) {
  const cur = Math.max(0, Number(currentHours) || 0);
  const base = Math.max(1, Number(baseInterval) || 1);
  if (cur <= 0) return base;
  if (Math.abs(cur % base) < 1e-6) return cur + base;
  return Math.ceil(cur / base) * base;
}

export function milestoneServiceInterval(milestoneHours, baseInterval, majorInterval) {
  const m = Number(milestoneHours);
  const base = Math.max(1, Number(baseInterval) || 1);
  const major = Math.max(base, Number(majorInterval) || base);
  if (major > base && Math.abs(m % major) < 1e-6) return major;
  return base;
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

export function resolveLegacyPlanDue(plan, currentHours) {
  const current = Number(currentHours) || 0;
  const last = Number(plan.last_service_hours || 0);
  const interval = planIntervalHours(plan);
  const nextDue = last + interval;
  const remaining = nextDue - current;
  return {
    schedule_mode: "legacy",
    plan_id: Number(plan.id ?? plan.plan_id ?? 0),
    asset_id: Number(plan.asset_id || 0),
    service_name: String(plan.service_name || ""),
    interval_hours: interval,
    last_service_hours: last,
    next_due_hours: Number(nextDue.toFixed(2)),
    remaining_hours: Number(remaining.toFixed(2)),
    next_service_interval: interval,
    is_next_for_asset: true,
  };
}

/** Pick the next service type from active plans on a shared hourmeter grid. */
export function resolveNextServiceForAssetPlans(plans, currentHours) {
  const activePlans = (plans || []).filter((p) => Number(p.active ?? 1) !== 0);
  if (!activePlans.length) return null;

  const intervals = [...new Set(activePlans.map(planIntervalHours).filter((x) => x > 0))].sort((a, b) => a - b);
  if (intervals.length < 2) {
    return resolveLegacyPlanDue(activePlans[0], currentHours);
  }

  const baseInterval = intervals[0];
  const majorInterval = intervals[intervals.length - 1];
  const current = Number(currentHours) || 0;
  const nextDue = nextMilestoneHours(current, baseInterval);
  const serviceInterval = milestoneServiceInterval(nextDue, baseInterval, majorInterval);
  const matchedPlan = findPlanForInterval(activePlans, serviceInterval);
  const remaining = nextDue - current;
  const lastService = Math.max(0, nextDue - serviceInterval);

  return {
    schedule_mode: "rotating",
    plan_id: Number(matchedPlan?.id ?? matchedPlan?.plan_id ?? activePlans[0]?.id ?? 0),
    asset_id: Number(activePlans[0].asset_id || 0),
    service_name: String(matchedPlan?.service_name || `${serviceInterval}`),
    interval_hours: serviceInterval,
    last_service_hours: Number(lastService.toFixed(2)),
    next_due_hours: Number(nextDue.toFixed(2)),
    remaining_hours: Number(remaining.toFixed(2)),
    next_service_interval: serviceInterval,
    base_interval_hours: baseInterval,
    major_interval_hours: majorInterval,
    is_next_for_asset: true,
  };
}

export function enrichPlansWithNextService(plans, getCurrentHours) {
  const byAsset = groupActivePlansByAsset(plans);
  const nextByAsset = new Map();
  for (const [assetId, assetPlans] of byAsset) {
    const current = typeof getCurrentHours === "function" ? getCurrentHours(assetId) : 0;
    nextByAsset.set(assetId, resolveNextServiceForAssetPlans(assetPlans, current));
  }

  return (plans || []).map((plan) => {
    const assetId = Number(plan.asset_id || 0);
    const assetPlans = byAsset.get(assetId) || [];
    const current = typeof getCurrentHours === "function" ? getCurrentHours(assetId) : 0;
    const rotating = hasRotatingSchedule(assetPlans);

    if (rotating) {
      const next = nextByAsset.get(assetId);
      const planId = Number(plan.id ?? plan.plan_id ?? 0);
      const isNext = Number(next?.plan_id || 0) === planId;
      return {
        ...plan,
        schedule_mode: "rotating",
        current_hours: Number(Number(current).toFixed(2)),
        next_due_hours: next?.next_due_hours ?? null,
        remaining_hours: isNext ? next?.remaining_hours ?? null : null,
        is_next_for_asset: isNext,
        next_service_name: isNext ? next?.service_name : null,
      };
    }

    const legacy = resolveLegacyPlanDue(plan, current);
    return {
      ...plan,
      schedule_mode: "legacy",
      current_hours: Number(Number(current).toFixed(2)),
      next_due_hours: legacy.next_due_hours,
      remaining_hours: legacy.remaining_hours,
      is_next_for_asset: true,
      next_service_name: legacy.service_name,
    };
  });
}

/** One due row per asset — uses rotating schedule when multiple intervals exist. */
export function buildDueListFromPlans(plans, getCurrentHours) {
  const byAsset = groupActivePlansByAsset(plans);
  const rows = [];

  for (const [, assetPlans] of byAsset) {
    const sample = assetPlans[0];
    const assetId = Number(sample.asset_id || 0);
    const current = typeof getCurrentHours === "function" ? getCurrentHours(assetId) : 0;
    const resolved = resolveNextServiceForAssetPlans(assetPlans, current);
    if (!resolved) continue;

    const remaining = resolved.remaining_hours;
    rows.push({
      plan_id: resolved.plan_id,
      asset_id: assetId,
      asset_code: sample.asset_code,
      asset_name: sample.asset_name,
      category: sample.category,
      service_name: resolved.service_name,
      interval_hours: resolved.interval_hours,
      last_service_hours: resolved.last_service_hours,
      current_hours: Number(Number(current).toFixed(2)),
      next_due_hours: resolved.next_due_hours,
      remaining_hours: resolved.remaining_hours,
      is_overdue: remaining <= 0,
      schedule_mode: resolved.schedule_mode,
      next_service_interval: resolved.next_service_interval,
      is_next_for_asset: true,
    });
  }

  return rows;
}
