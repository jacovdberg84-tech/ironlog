const { app, BrowserWindow, dialog } = require("electron");
const path = require("node:path");
const http = require("node:http");
const fs = require("node:fs");
const os = require("node:os");
const { pathToFileURL } = require("node:url");

const DESKTOP_PORT = Number(process.env.IRONLOG_DESKTOP_PORT || 3210);
const DESKTOP_PORT_MAX = Number(process.env.IRONLOG_DESKTOP_PORT_MAX || DESKTOP_PORT + 20);
const API_CWD = path.resolve(__dirname, "..");
const API_ENTRY = path.join(API_CWD, "index.js");
const APP_ICON = path.join(__dirname, "icon.ico");
let STARTUP_LOG = path.join(os.tmpdir(), "ironlog-desktop-startup.log");

function getDesktopPort() {
  const p = Number(process.env.PORT_EFFECTIVE || process.env.PORT || DESKTOP_PORT);
  return Number.isFinite(p) ? p : DESKTOP_PORT;
}

function appUrl() {
  return process.env.IRONLOG_DESKTOP_URL || `http://127.0.0.1:${getDesktopPort()}/web/index.html`;
}

function healthUrl() {
  return process.env.IRONLOG_DESKTOP_HEALTH || `http://127.0.0.1:${getDesktopPort()}/health`;
}

function logLine(msg) {
  const line = `[${new Date().toISOString()}] ${msg}`;
  try {
    if (STARTUP_LOG) fs.appendFileSync(STARTUP_LOG, `${line}\n`);
  } catch {}
  try {
    // In packaged Windows GUI mode stdout/stderr can be invalid (EBADF).
    if (process.stdout && typeof process.stdout.write === "function") {
      process.stdout.write(`${line}\n`);
    }
  } catch {}
}

function fatalDialog(title, message) {
  try {
    dialog.showErrorBox(title, `${message}\n\nLog file:\n${STARTUP_LOG}`);
  } catch {}
}

process.on("uncaughtException", (err) => {
  const msg = String(err?.stack || err?.message || err || "Unknown uncaught exception");
  logLine(`[desktop] uncaughtException: ${msg}`);
  fatalDialog("IRONLOG Fatal Error", msg);
});

process.on("unhandledRejection", (reason) => {
  const msg = String(reason?.stack || reason?.message || reason || "Unknown unhandled rejection");
  logLine(`[desktop] unhandledRejection: ${msg}`);
  fatalDialog("IRONLOG Fatal Error", msg);
});

function ensureRuntimeEnvironment() {
  // Put writable runtime data under user profile for installed app.
  const dataRoot = path.join(app.getPath("userData"), "data");
  const dbDir = path.join(dataRoot, "db");
  const dbPath = path.join(dbDir, "ironlog.db");
  STARTUP_LOG = path.join(app.getPath("userData"), "desktop-startup.log");
  fs.mkdirSync(dbDir, { recursive: true });
  fs.mkdirSync(path.join(dataRoot, "uploads"), { recursive: true });

  const bundledDb = path.join(API_CWD, "db", "ironlog.db");
  if (!fs.existsSync(dbPath) && fs.existsSync(bundledDb)) {
    fs.copyFileSync(bundledDb, dbPath);
  }

  const packagedWebRoot = path.join(process.resourcesPath || "", "web");
  const asarWebRoot = path.join(API_CWD, "web");
  const webRoot = fs.existsSync(packagedWebRoot) ? packagedWebRoot : asarWebRoot;

  process.env.IRONLOG_DATA_DIR = dataRoot;
  process.env.IRONLOG_APP_DIR = API_CWD;
  process.env.IRONLOG_WEB_DIR = webRoot;
  process.env.DB_PATH = process.env.DB_PATH || dbPath;
  logLine(`Runtime data root: ${dataRoot}`);
  logLine(`Database path: ${process.env.DB_PATH}`);
  logLine(`Web root: ${webRoot}`);
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function probe(url) {
  return new Promise((resolve) => {
    const req = http.get(url, (res) => {
      res.resume();
      resolve(res.statusCode >= 200 && res.statusCode < 400);
    });
    req.on("error", () => resolve(false));
    req.setTimeout(2000, () => {
      req.destroy();
      resolve(false);
    });
  });
}

async function waitForServer(maxWaitMs = 30000) {
  const start = Date.now();
  while (Date.now() - start < maxWaitMs) {
    if (await probe(healthUrl())) return true;
    if (await probe(appUrl())) return true;
    await wait(500);
  }
  return false;
}

async function ensureApiRunning() {
  const alreadyUp = await probe(healthUrl()) || (await probe(appUrl()));
  if (alreadyUp) return;

  // Run API inside Electron main process for packaged reliability.
  process.env.IRONLOG_DESKTOP = "1";
  // Try a dedicated desktop API port range to avoid clashes.
  process.env.PORT = String(DESKTOP_PORT);
  process.env.IRONLOG_DESKTOP_PORT_MAX = String(DESKTOP_PORT_MAX);
  process.env.HOST = process.env.HOST || "0.0.0.0";
  logLine(`Desktop API port range: ${DESKTOP_PORT}-${DESKTOP_PORT_MAX}`);
  logLine(`Booting API from: ${API_ENTRY}`);
  await import(pathToFileURL(API_ENTRY).href);
  logLine(`Desktop API selected port: ${getDesktopPort()}`);

  const up = await waitForServer(45000);
  if (!up) {
    throw new Error("API did not start in time");
  }
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1100,
    minHeight: 720,
    autoHideMenuBar: true,
    icon: APP_ICON,
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  if (process.argv.includes("--devtools")) {
    win.webContents.openDevTools({ mode: "detach" });
  }
  win.webContents.on("did-fail-load", (_event, code, desc, url) => {
    logLine(`[desktop] did-fail-load code=${code} url=${url} desc=${desc}`);
  });
  win.on("unresponsive", () => logLine("[desktop] window unresponsive"));
  win.on("closed", () => logLine("[desktop] window closed"));
  win.loadURL(appUrl()).catch((err) => {
    const msg = String(err?.stack || err?.message || err || "Failed to load app URL");
    logLine(`[desktop] loadURL failed: ${msg}`);
    fatalDialog("IRONLOG Startup Error", msg);
  });
}

app.whenReady().then(async () => {
  try {
    ensureRuntimeEnvironment();
    await ensureApiRunning();
    createWindow();
  } catch (err) {
    const msg = String(err?.stack || err?.message || err || "Unknown startup error");
    logLine(`[desktop] startup failed: ${msg}`);
    fatalDialog("IRONLOG Startup Error", msg);
    app.quit();
  }

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

app.on("before-quit", () => {
  // API runs in-process; app quit stops it automatically.
});

