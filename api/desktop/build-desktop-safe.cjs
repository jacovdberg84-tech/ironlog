const { execSync, spawnSync, spawn } = require("node:child_process");

function run(cmd, opts = {}) {
  const isWin = process.platform === "win32";
  const shell = isWin ? "powershell.exe" : true;
  const args = isWin ? ["-NoProfile", "-Command", cmd] : undefined;
  if (isWin) {
    const res = spawnSync(shell, args, { stdio: "inherit", ...opts });
    if (res.status !== 0) throw new Error(`Command failed: ${cmd}`);
    return;
  }
  execSync(cmd, { stdio: "inherit", ...opts });
}

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

async function main() {
  const wasRunning = Boolean(getPidOnPort3001());
  const pid = getPidOnPort3001();
  if (pid) {
    console.log(`[desktop-safe] stopping process on :3001 (PID ${pid})`);
    run(`taskkill /PID ${pid} /F`);
  } else {
    console.log("[desktop-safe] no process detected on :3001");
  }

  try {
    console.log("[desktop-safe] building desktop installer...");
    run("node desktop/build-desktop.cjs");

    console.log("[desktop-safe] restoring better-sqlite3 for Node runtime...");
    run("npm rebuild better-sqlite3 --build-from-source --runtime=node --target=20.20.0");

    if (wasRunning) {
      console.log("[desktop-safe] restarting dev server...");
      if (process.platform === "win32") {
        spawn("powershell.exe", ["-NoProfile", "-Command", "npm run dev"], {
          detached: true,
          stdio: "ignore",
        }).unref();
      } else {
        spawn("npm", ["run", "dev"], { detached: true, stdio: "ignore" }).unref();
      }
    }

    console.log("[desktop-safe] done.");
  } catch (err) {
    console.error("[desktop-safe] failed:", err?.message || err);
    process.exit(1);
  }
}

main();

