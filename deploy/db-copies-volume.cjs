// deploy/db-copies-volume.cjs — keeps IronLog's on-server database copies on a separate volume.
//
//   node db-copies-volume.cjs configure <volume mount>
//       Points AML (api/.env) and every company instance (instance.env) at
//       <volume>/ironlog/<name>/db-copies via IRONLOG_DB_BACKUP_DIR. Refuses
//       when the volume is not mounted. Prints "name<TAB>old dir<TAB>new dir".
//   node db-copies-volume.cjs move <old dir> <new dir>
//       Moves the copies across: each file is copied, checked (size and
//       SHA-256) and only then removed from the old folder.
// The live databases are not touched. Run by the "Move database copies to
// the backup volume" workflow, which restarts IronLog between the two steps.
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const CODE_DIR = process.env.IRONLOG_CODE_DIR || "/data/ironlog";
const INSTANCES_DIR = process.env.IRONLOG_INSTANCES_DIR || "/data/ironlog-instances";
const KEY = "IRONLOG_DB_BACKUP_DIR";

/** Settings file text with KEY set to value (any earlier KEY lines removed). */
function setEnvLine(text, key, value) {
  const kept = String(text || "").split("\n").filter((l) => !l.startsWith(`${key}=`));
  while (kept.length && kept[kept.length - 1] === "") kept.pop();
  return `${kept.join("\n")}${kept.length ? "\n" : ""}${key}=${value}\n`;
}

function envValue(text, key) {
  const line = String(text || "").split("\n").find((l) => l.startsWith(`${key}=`));
  return line ? line.slice(key.length + 1).trim() : "";
}

/** True when dir is on a different disk from the root (so the volume is really mounted). */
function isMounted(dir, rootDir = "/") {
  try {
    return fs.statSync(dir).isDirectory() && fs.statSync(dir).dev !== fs.statSync(rootDir).dev;
  } catch {
    return false;
  }
}

function targets() {
  const list = [{ name: "aml", envFile: path.join(CODE_DIR, "api/.env"), defaultDir: path.join(CODE_DIR, "api/db/backups") }];
  if (fs.existsSync(INSTANCES_DIR)) {
    for (const name of fs.readdirSync(INSTANCES_DIR).sort()) {
      const envFile = path.join(INSTANCES_DIR, name, "instance.env");
      if (fs.existsSync(envFile)) list.push({ name, envFile, defaultDir: path.join(INSTANCES_DIR, name, "db/backups") });
    }
  }
  return list;
}

function configure(volume, { mounted = isMounted } = {}) {
  if (!mounted(volume)) throw new Error(`${volume} is not a mounted volume; nothing was changed.`);
  const out = [];
  for (const t of targets()) {
    const text = fs.existsSync(t.envFile) ? fs.readFileSync(t.envFile, "utf8") : "";
    const oldDir = path.resolve(path.dirname(t.envFile), envValue(text, KEY) || t.defaultDir);
    const newDir = path.join(volume, "ironlog", t.name, "db-copies");
    fs.mkdirSync(newDir, { recursive: true, mode: 0o700 });
    fs.chmodSync(path.join(volume, "ironlog", t.name), 0o700);
    if (oldDir !== newDir) {
      fs.writeFileSync(t.envFile, setEnvLine(text, KEY, newDir), { mode: 0o600 });
      fs.chmodSync(t.envFile, 0o600);
    }
    out.push([t.name, oldDir, newDir]);
  }
  return out;
}

function sha256(file) {
  return crypto.createHash("sha256").update(fs.readFileSync(file)).digest("hex");
}

/** Moves every finished file (not *.tmp) from src to dst, verifying each copy first. */
function moveFiles(src, dst) {
  const result = { moved: 0, already: 0, bytes: 0, failed: [] };
  if (!fs.existsSync(src) || path.resolve(src) === path.resolve(dst)) return result;
  fs.mkdirSync(dst, { recursive: true, mode: 0o700 });
  for (const name of fs.readdirSync(src).sort()) {
    const from = path.join(src, name);
    const to = path.join(dst, name);
    if (!fs.statSync(from).isFile() || /\.tmp(\.db)?$/.test(name)) continue;
    try {
      const hash = sha256(from);
      if (fs.existsSync(to) && sha256(to) === hash) {
        fs.unlinkSync(from);
        result.already += 1;
        continue;
      }
      const part = `${to}.moving`;
      fs.copyFileSync(from, part);
      const fd = fs.openSync(part, "r");
      fs.fsyncSync(fd);
      fs.closeSync(fd);
      if (fs.statSync(part).size !== fs.statSync(from).size || sha256(part) !== hash) {
        fs.unlinkSync(part);
        throw new Error("copy did not match");
      }
      fs.renameSync(part, to);
      fs.utimesSync(to, fs.statSync(from).atime, fs.statSync(from).mtime);
      result.bytes += fs.statSync(to).size;
      fs.unlinkSync(from);
      result.moved += 1;
    } catch (err) {
      result.failed.push(`${name}: ${err.message}`);
    }
  }
  return result;
}

if (require.main === module) {
  const [cmd, a, b] = process.argv.slice(2);
  try {
    if (cmd === "configure") {
      for (const row of configure(a)) console.log(row.join("\t"));
    } else if (cmd === "move") {
      const r = moveFiles(a, b);
      console.log(`moved ${r.moved} (${(r.bytes / 1e9).toFixed(2)} GB), already there ${r.already}, problems ${r.failed.length}`);
      for (const f of r.failed) console.log(`  kept in old folder: ${f}`);
      if (r.failed.length) process.exitCode = 1;
    } else {
      console.log("Usage: db-copies-volume.cjs configure <volume> | move <old dir> <new dir>");
      process.exitCode = 2;
    }
  } catch (err) {
    console.error(`Stopped: ${err.message}`);
    process.exit(1);
  }
}

module.exports = { setEnvLine, envValue, isMounted, configure, moveFiles };
