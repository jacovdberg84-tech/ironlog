import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { famsSelectedDateRange } from "../utils/famsFuelRange.js";

test("FAMS catch-up range accepts a completed August period", () => {
  assert.deepEqual(
    famsSelectedDateRange({
      startDate: "2026-08-01",
      endDate: "2026-08-31",
      now: new Date("2026-09-14T10:00:00+02:00"),
    }),
    {
      startYmd: "2026-08-01",
      endYmd: "2026-08-31",
      startSlash: "2026/08/01",
      endSlash: "2026/08/31",
    },
  );
});

test("FAMS catch-up range rejects unsafe calendar ranges", () => {
  const now = new Date("2026-09-14T10:00:00+02:00");
  assert.throws(
    () => famsSelectedDateRange({ startDate: "2026-08-31", endDate: "2026-08-01", now }),
    /cannot be after/i,
  );
  assert.throws(
    () => famsSelectedDateRange({ startDate: "2026-02-30", endDate: "2026-03-01", now }),
    /valid calendar date/i,
  );
  assert.throws(
    () => famsSelectedDateRange({ startDate: "2026-09-01", endDate: "2026-09-15", now }),
    /later than today/i,
  );
});

test("FAMS cleanup matches a unique source/day/litres pair even when the legacy meter differs", async (t) => {
  const dbPath = path.join(os.tmpdir(), `ironlog-fams-duplicate-${process.pid}-${Date.now()}.db`);
  process.env.DB_PATH = dbPath;
  const { db } = await import("../db/client.js");
  const {
    ensureFamsFuelSchema,
    previewFamsLegacyDuplicates,
    removeFamsLegacyDuplicates,
  } = await import("../utils/famsFuel.js");

  t.after(async () => {
    db.close();
    await fs.rm(dbPath, { force: true });
  });

  db.exec(`
    CREATE TABLE assets (id INTEGER PRIMARY KEY, asset_code TEXT NOT NULL, asset_name TEXT);
    CREATE TABLE fuel_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      log_date TEXT NOT NULL,
      liters REAL NOT NULL,
      source TEXT
    );
  `);
  ensureFamsFuelSchema();
  db.prepare(`INSERT INTO assets (id, asset_code, asset_name) VALUES (1, 'A300AM', 'Bell B30D')`).run();
  const insert = db.prepare(`
    INSERT INTO fuel_logs (asset_id, log_date, liters, source, meter_unit, meter_run_value, fams_id)
    VALUES (?, ?, ?, ?, ?, ?, ?)
  `);
  insert.run(1, "2026-08-12", 120, "Main Store | Sergio | Driver A", "hours", 20258.2, null);
  // FAMS API records the transaction run; legacy CSVs recorded the absolute SMR.
  insert.run(1, "2026-08-12", 120, "Main Store | Sergio | Driver A", "hours", 12.5, 7001);
  // This has the same asset/day/litres but a different source and must survive.
  insert.run(1, "2026-08-12", 120, "Manual fuel entry", "hours", 20259.2, null);

  const preview = previewFamsLegacyDuplicates({ startDate: "2026-08-01", endDate: "2026-08-31" });
  assert.equal(preview.found, 1);
  assert.equal(preview.rows[0].asset_code, "A300AM");

  const result = removeFamsLegacyDuplicates({ startDate: "2026-08-01", endDate: "2026-08-31" });
  assert.equal(result.removed, 1);
  assert.equal(db.prepare(`SELECT COUNT(*) AS n FROM fuel_logs`).get().n, 2);
  assert.equal(db.prepare(`SELECT COUNT(*) AS n FROM fuel_logs WHERE fams_id = 7001`).get().n, 1);
});
