const fs = require("node:fs");
const path = require("node:path");
const { execSync } = require("node:child_process");

const ROOT = path.resolve(__dirname, "..");
const DIST = path.join(ROOT, "desktop-dist");
const PKG = path.join(ROOT, "package.json");

function getPidOnPort3001() {
  if (process.platform !== "win32") return null;
  try {
    const out = execSync('netstat -ano | findstr ":3001"', { encoding: "utf8" });
    const line = String(out || "")
      .split(/\r?\n/)
      .find((l) => /\sLISTENING\s/i.test(l));
    if (!line) return null;
    const parts = line.trim().split(/\s+/);
    const pid = Number(parts[parts.length - 1]);
    return Number.isInteger(pid) && pid > 0 ? pid : null;
  } catch {
    return null;
  }
}

function latestInstaller() {
  if (!fs.existsSync(DIST)) return null;
  const files = fs
    .readdirSync(DIST)
    .filter((n) => n.toLowerCase().endsWith(".exe"))
    .map((n) => {
      const p = path.join(DIST, n);
      const st = fs.statSync(p);
      return { name: n, path: p, mtimeMs: st.mtimeMs, size: st.size };
    })
    .sort((a, b) => b.mtimeMs - a.mtimeMs);
  return files[0] || null;
}

function fmtBytes(n) {
  const v = Number(n || 0);
  if (v < 1024) return `${v} B`;
  if (v < 1024 * 1024) return `${(v / 1024).toFixed(1)} KB`;
  if (v < 1024 * 1024 * 1024) return `${(v / (1024 * 1024)).toFixed(1)} MB`;
  return `${(v / (1024 * 1024 * 1024)).toFixed(2)} GB`;
}

function readPkg() {
  try {
    const raw = fs.readFileSync(PKG, "utf8");
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

const pkg = readPkg();
const pid = getPidOnPort3001();
const installer = latestInstaller();

console.log("IRONLOG Desktop Status");
console.log("======================");
console.log(`Package version : ${pkg.version || "unknown"}`);
console.log(`Desktop script  : ${pkg.scripts?.["desktop:dist:safe"] ? "configured" : "missing"}`);
console.log(`API on :3001    : ${pid ? `running (PID ${pid})` : "not running"}`);
if (installer) {
  console.log(`Latest installer: ${installer.name}`);
  console.log(`Path            : ${installer.path}`);
  console.log(`Size            : ${fmtBytes(installer.size)}`);
  console.log(`Updated         : ${new Date(installer.mtimeMs).toISOString()}`);
} else {
  console.log("Latest installer: none found in desktop-dist");
}

