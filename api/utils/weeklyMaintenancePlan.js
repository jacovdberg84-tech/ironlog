// Draft planning only: no work orders, stock reservations, or model arithmetic.
const DAY = 86400000;
export function weeklyWindow(today) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(today) || new Date(today + 'T00:00:00Z').toISOString().slice(0,10) !== today) throw new Error('Invalid planning date');
  const date = new Date(today + 'T00:00:00Z');
  const start = new Date(+date + ((8 - date.getUTCDay()) % 7 || 7) * DAY);
  return { as_of: today, start: start.toISOString().slice(0,10), end: new Date(+start + 6 * DAY).toISOString().slice(0,10) };
}
const money = n => Math.round(n * 100) / 100;
const valid = n => n !== null && n !== undefined && Number.isFinite(Number(n)) && Number(n) >= 0;
export function buildWeeklyMaintenancePlan(rows, today) {
  const range = weeklyWindow(today);
  const services = [], review = [], demand = new Map();
  for (const row of rows) {
    const gaps = [];
    const remaining = Number(row.remaining_hours);
    const usage = row.usage || {};
    const avg = Number(usage.total_run) / 14;
    const meterKm = row.meter_unit === 'km';
    const usageValid = !meterKm && !usage.invalid_days && Number(usage.day_count) > 0 && avg > 0 && avg <= 24;
    let due = null;
    if (row.meter_source !== 'daily_sum' && valid(row.current_hours) && Number.isFinite(remaining)) {
      if (remaining <= 0) due = today;
      else if (usageValid) due = new Date(Date.parse(today + 'T00:00:00Z') + Math.ceil(remaining / avg) * DAY).toISOString().slice(0,10);
    }
    if (!due) { review.push({ asset_code: row.asset_code, service_name: row.service_name, reason: meterKm ? 'Kilometre usage forecast required' : 'Missing or invalid meter/usage readings' }); continue; }
    if (due > range.end) continue;
    const f = row.forecast || {};
    const items = f.manual?.items || [];
    if (!items.length) gaps.push('Required parts/quantities have not been entered');
    for (const item of items) {
      if (!valid(item.unit_cost) || Number(item.unit_cost) <= 0) gaps.push('Price missing for ' + item.part_code);
      const key = String(item.part_code).trim().toUpperCase();
      const previous = demand.get(key) || { part_code: key, qty: 0, on_hand: valid(item.on_hand) ? Number(item.on_hand) : null };
      previous.qty += Number(item.qty) || 0;
      if (!valid(item.on_hand)) previous.on_hand = null;
      else if (previous.on_hand !== null) previous.on_hand = Math.min(previous.on_hand, Number(item.on_hand));
      demand.set(key, previous);
    }
    const kit = valid(f.est_service_kit_cost) && Number(f.est_service_kit_cost) > 0 ? Number(f.est_service_kit_cost) : null;
    const labor = valid(f.est_labor_cost) && Number(f.est_labor_cost) > 0 ? Number(f.est_labor_cost) : null;
    if (kit === null) gaps.push('Service kit cost missing');
    if (labor === null) gaps.push('Labour estimate missing');
    if (!usageValid && !meterKm) gaps.push('Usage forecast unavailable; included because service is already due');
    services.push({ plan_id: row.plan_id, asset_code: row.asset_code, service_name: row.service_name,
      estimated_due: due, priority: remaining <= 0 ? 'OVERDUE / DUE NOW' : due < range.start ? 'DUE BEFORE NEXT WEEK' : 'DUE NEXT WEEK',
      remaining_hours: remaining, cost_source: f.cost_source || 'none', kit_cost: kit, labor_cost: labor,
      known_cost: money((kit || 0) + (labor || 0)), gaps });
  }
  services.sort((a,b) => a.estimated_due.localeCompare(b.estimated_due) || a.remaining_hours-b.remaining_hours);
  const parts = [...demand.values()].map(p => ({...p, qty: money(p.qty), shortage: p.on_hand === null ? null : money(Math.max(0,p.qty-p.on_hand))}));
  const known_cost = money(services.reduce((sum,s)=>sum+s.known_cost,0));
  const lines = [
    'Borris — weekly maintenance draft: ' + range.start + ' to ' + range.end,
    'Based on records as of ' + today + '. Includes overdue work and services due before the week ends.',
    'Forecast uses the last 14 calendar days of recorded running hours. This is a priority queue, not a confirmed workshop schedule.',
    'Known estimated cost: ' + known_cost.toFixed(2) + ' (stored cost units; missing amounts excluded).',
    ...services.map(s => s.asset_code + ' — ' + s.service_name + ': ' + s.priority + '; estimated due ' + s.estimated_due + '; known cost ' + s.known_cost.toFixed(2) + ' (' + s.cost_source + ').' + (s.gaps.length ? ' Check: ' + s.gaps.join('; ') + '.' : '')),
    ...(services.length ? [] : ['No services can currently be placed in this planning window.']),
    'Parts demand (shared stock, not reserved):',
    ...parts.map(p => p.part_code + ': need ' + p.qty + ', on hand ' + (p.on_hand ?? 'unknown') + ', shortage ' + (p.shortage ?? 'unknown')),
    ...review.map(r => 'Needs review: ' + r.asset_code + ' — ' + r.service_name + ': ' + r.reason),
    'Confirm meter readings, parts, labour availability and downtime windows before creating work orders.'
  ];
  return { ok: true, draft: true, range, services, parts, needs_review: review, known_cost, short_answer: lines.join('\n') };
}
