import { db } from "../db/client.js";
import { buildReliabilityIncidentsForAssets, woWallClockHoursInRange } from "../utils/reliabilityMetrics.js";

const hasTable = (n) =>
  Boolean(db.prepare("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?").get(n));
const hasColumn = (t, c) =>
  db.prepare(`PRAGMA table_info(${t})`).all().some((r) => r.name === c);

const r = buildReliabilityIncidentsForAssets(db, {
  assetIds: [329],
  start: "2026-01-01",
  end: "2026-06-18",
  hasTable,
  hasColumn,
});
console.log(JSON.stringify(r, null, 2));
console.log(
  "wo 2.5h test",
  woWallClockHoursInRange("2026-06-10 08:00:00", "2026-06-10 10:30:00", "2026-06-01", "2026-06-18"),
);
