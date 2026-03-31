import { db } from "./client.js";

const breakdownId = 26;

const rows = db.prepare(`
  SELECT
    id,
    asset_id,
    source,
    reference_id,
    status,
    opened_at
  FROM work_orders
  WHERE source = 'breakdown'
    AND reference_id = ?
`).all(breakdownId);

console.log("Work orders for breakdown", breakdownId);
console.table(rows);