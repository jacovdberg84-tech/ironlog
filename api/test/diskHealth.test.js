import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { diskHealth, diskLevel, folderSize, ironlogUsage, serverDisks, worstDisk } from "../utils/diskHealth.js";

const GB = 1e9;
const disk = (freeGb, totalGb) => ({ bavail: freeGb * GB / 4096, blocks: totalGb * GB / 4096, bsize: 4096 });

test("the disk warning comes early (10% or 5 GB) and turns urgent near full (5% or 2 GB)", () => {
  assert.equal(diskLevel(disk(24, 75)).level, "ok");
  assert.equal(diskLevel(disk(30, 270)).level, "ok", "11% free on a big disk is still plenty");
  const warn = diskLevel(disk(7, 75));
  assert.equal(warn.level, "warn", "9% free on a 75 GB disk");
  assert.equal(warn.used_pct, 91);
  assert.match(warn.message, /getting full: 7 GB free \(91% used\)/);
  assert.equal(diskLevel(disk(4, 400)).level, "critical", "1% free, even with 4 GB left");
  assert.equal(diskLevel(disk(1.5, 20)).level, "critical", "under 2 GB");
  assert.match(diskLevel(disk(0.16, 75)).message, /almost full: only 0.2 GB free \(100% used\)/);
  assert.equal(diskLevel(disk(4, 20)).level, "warn", "under 5 GB though 20% free");
  assert.equal(diskHealth("/", () => { throw new Error("no statfs"); }), null);
  assert.ok(diskHealth(os.tmpdir()).total_gb > 0, "reads the real disk");
});

test("shows what IronLog keeps on the disk, largest first", () => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "ironlog-disk-"));
  const dbPath = path.join(root, "ironlog.db");
  fs.writeFileSync(dbPath, Buffer.alloc(1000));
  fs.mkdirSync(path.join(root, "backups"));
  for (let i = 0; i < 3; i++) fs.writeFileSync(path.join(root, "backups", `ironlog_${i}.db`), Buffer.alloc(2000));
  fs.mkdirSync(path.join(root, "data", "uploads", "photos"), { recursive: true });
  fs.writeFileSync(path.join(root, "data", "uploads", "photos", "a.jpg"), Buffer.alloc(500));
  const rows = ironlogUsage({ dbPath, backupDir: path.join(root, "backups"), dataDir: path.join(root, "data") });
  assert.deepEqual(rows.map((r) => [r.label, r.bytes, r.files]), [
    ["Database copies on this server", 6000, 3],
    ["Database", 1000, 1],
    ["Photos and uploads", 500, 1],
  ]);
  assert.equal(folderSize(path.join(root, "backups"), 2).partial, true, "very large folders stop counting early");
  assert.deepEqual(folderSize(path.join(root, "missing")), { bytes: 0, files: 0, partial: false });
  fs.rmSync(root, { recursive: true, force: true });
});

test("database copies on a separate volume get their own bar and warning", () => {
  const statfs = (dir) => (dir.includes("volume") ? disk(1, 40) : disk(50, 75));
  const stat = (dir) => ({ dev: dir.includes("volume") ? 2 : 1 });
  const both = serverDisks({ dbPath: "/data/ironlog/api/db/ironlog.db", backupDir: "/mnt/volume/ironlog/aml/db-copies" }, { statfs, stat });
  assert.deepEqual(both.map((d) => [d.name, d.level]), [["Server disk", "ok"], ["Backup volume", "critical"]]);
  assert.match(both[1].message, /^Backup volume almost full: .*New database copies will stop until space is freed; IronLog keeps working/);
  assert.equal(worstDisk(both).name, "Backup volume");
  const same = serverDisks({ dbPath: "/data/ironlog/api/db/ironlog.db", backupDir: "/data/ironlog/api/db/backups" }, { statfs, stat });
  assert.deepEqual(same.map((d) => d.name), ["Server disk"], "copies on the same disk: one bar");
  assert.equal(worstDisk(same).name, "Server disk");
  assert.equal(worstDisk([]), null);
});
