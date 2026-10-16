import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";

const root = fs.mkdtempSync(path.join(os.tmpdir(), "ironlog-vol-"));
process.env.IRONLOG_CODE_DIR = path.join(root, "ironlog");
process.env.IRONLOG_INSTANCES_DIR = path.join(root, "instances");
const vol = createRequire(import.meta.url)("../../deploy/db-copies-volume.cjs");

test("each IronLog is pointed at its own folder on the volume; other settings are kept", () => {
  fs.mkdirSync(path.join(root, "ironlog/api"), { recursive: true });
  fs.writeFileSync(path.join(root, "ironlog/api/.env"), "OPENAI_API_KEY=sk\nPORT=3001\n");
  fs.mkdirSync(path.join(root, "instances/pilot"), { recursive: true });
  fs.writeFileSync(path.join(root, "instances/pilot/instance.env"), `PORT=3101\nIRONLOG_DB_BACKUP_DIR=${root}/instances/pilot/db/backups\n`);
  const volume = path.join(root, "volume");
  fs.mkdirSync(volume);

  assert.throws(() => vol.configure(volume), /not a mounted volume; nothing was changed/, "same disk as / here, so refused");
  assert.equal(vol.isMounted(path.join(root, "missing")), false);

  const plan = vol.configure(volume, { mounted: () => true });
  assert.deepEqual(plan, [
    ["aml", path.join(root, "ironlog/api/db/backups"), path.join(volume, "ironlog/aml/db-copies")],
    ["pilot", path.join(root, "instances/pilot/db/backups"), path.join(volume, "ironlog/pilot/db-copies")],
  ]);
  assert.equal(fs.readFileSync(path.join(root, "ironlog/api/.env"), "utf8"), `OPENAI_API_KEY=sk\nPORT=3001\nIRONLOG_DB_BACKUP_DIR=${volume}/ironlog/aml/db-copies\n`);
  assert.equal(fs.readFileSync(path.join(root, "instances/pilot/instance.env"), "utf8"), `PORT=3101\nIRONLOG_DB_BACKUP_DIR=${volume}/ironlog/pilot/db-copies\n`);
  assert.equal(fs.statSync(path.join(root, "instances/pilot/instance.env")).mode & 0o777, 0o600);
  assert.equal(fs.statSync(path.join(volume, "ironlog/pilot")).mode & 0o777, 0o700, "each company's folder is private");
  const again = vol.configure(volume, { mounted: () => true });
  assert.equal(again[0][1], again[0][2], "running it again finds everything already moved");
});

test("copies are checked before the old file is removed; unfinished files stay", () => {
  const src = path.join(root, "old");
  const dst = path.join(root, "new");
  fs.mkdirSync(src);
  fs.writeFileSync(path.join(src, "ironlog_20261015_101010.db"), "copy-one");
  fs.writeFileSync(path.join(src, "ironlog_auto_20261015_020000.db"), "nightly");
  fs.writeFileSync(path.join(src, "ironlog_20261016_000000.tmp.db"), "half-written");
  fs.mkdirSync(dst);
  fs.writeFileSync(path.join(dst, "ironlog_auto_20261015_020000.db"), "nightly");
  const r = vol.moveFiles(src, dst);
  assert.deepEqual([r.moved, r.already, r.failed], [1, 1, []]);
  assert.deepEqual(fs.readdirSync(src), ["ironlog_20261016_000000.tmp.db"]);
  assert.deepEqual(fs.readdirSync(dst).sort(), ["ironlog_20261015_101010.db", "ironlog_auto_20261015_020000.db"]);
  assert.equal(fs.readFileSync(path.join(dst, "ironlog_20261015_101010.db"), "utf8"), "copy-one");
  assert.deepEqual(vol.moveFiles(path.join(root, "nothing"), dst), { moved: 0, already: 0, bytes: 0, failed: [] });
  fs.rmSync(root, { recursive: true, force: true });
});
