// IRONLOG/api/utils/downtimeCap.js — a machine cannot be down while it runs.
//
// Downtime for a day comes from logged breakdown downtime or, when nothing was
// logged, from the repair work order's labour hours. Neither knows how long
// the machine actually ran that day, so a repair finished early in the shift
// could still show a full repair's worth of downtime (A301AM: ran 9+ hours,
// reported 9 hours down). The day's downtime is therefore capped at the shift
// hours minus the hours it ran, whenever run hours were recorded.

/** Most downtime a machine can have on a day, or Infinity when it did not run. */
export function downtimeCapForRun({ scheduled, run }) {
  const s = Math.max(0, Number(scheduled || 0));
  const r = Math.max(0, Number(run || 0));
  if (!(r > 0)) return Infinity;
  return Math.max(0, s - r);
}

export function capDowntime(hours, { scheduled, run }) {
  return Math.min(Math.max(0, Number(hours || 0)), downtimeCapForRun({ scheduled, run }));
}
