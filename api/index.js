// IRONLOG/api/index.js
import "./loadEnv.js"; // settings before anything else reads them
import { buildServer } from "./server.js";
import { runHybridPhase1Migration } from "./db/hybridPhase1.js";
import { runHybridPhase2Migration } from "./db/hybridPhase2.js";
import "./db/migrate.js"; // Run schema migration first

import fastifyStatic from "@fastify/static";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";
import { runIronmindAutoScheduler } from "./utils/ironmind.js";
import { runCartrackAutoScheduler } from "./utils/cartrack.js";
import { runUnitechAutoScheduler } from "./utils/unitech.js";
import { runFamsAutoScheduler } from "./utils/famsFuel.js";
import { startDbAutoBackup } from "./db/autoBackup.js";
import { startBackupScheduler } from "./utils/autoBackup.js";
import { getCompanyProfile } from "./utils/companyProfile.js";
import { backupDirPath } from "./routes/backups.routes.js";
import { sendPushToRoles } from "./utils/pushNotify.js";
import { db } from "./db/client.js";
import { bootstrapSmtpFromEnv } from "./utils/mail.js";
import { repairSiteDataOnce } from "./utils/sites.js";
import { ensureAuditTable, writeAudit } from "./utils/audit.js";

// Settings are read by ./loadEnv.js (imported before everything else): api/.env,
// or only an instance's own file when IRONLOG_ENV_FILE is set.

const PORT = process.env.PORT ? Number(process.env.PORT) : 3001;
const HOST = process.env.HOST || "0.0.0.0";
const DESKTOP_PORT_MAX = process.env.IRONLOG_DESKTOP_PORT_MAX
  ? Number(process.env.IRONLOG_DESKTOP_PORT_MAX)
  : PORT;

function runStartupMigrationsSafely() {
  try {
    runHybridPhase1Migration();
  } catch (err) {
    console.error("[startup] phase1 migration skipped:", err?.message || err);
  }
  try {
    runHybridPhase2Migration();
  } catch (err) {
    console.error("[startup] phase2 migration skipped:", err?.message || err);
  }
}

runStartupMigrationsSafely();

const app = buildServer();

// Resolve ../web path (ESM-safe)
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const webCandidates = [
  process.env.IRONLOG_WEB_DIR,
  path.join(__dirname, "web"),
  path.join(__dirname, "../web"),
].filter(Boolean);
const webRoot = webCandidates.find((p) => fs.existsSync(p)) || path.join(__dirname, "../web");

// Serve the web UI at /web/*
app.register(fastifyStatic, {
  root: webRoot,
  prefix: "/web/",
  decorateReply: false, // keeps reply.sendFile available without extra decoration
});

// Uploaded images (manager inspections, vehicle LDV checks, etc.) — same root as maintenance routes
const dataRoot = process.env.IRONLOG_DATA_DIR || process.cwd();
const uploadsRoot = path.join(dataRoot, "uploads");
fs.mkdirSync(uploadsRoot, { recursive: true });
app.register(fastifyStatic, {
  root: uploadsRoot,
  prefix: "/uploads/",
  decorateReply: false,
});

const downloadsRoot = path.join(webRoot, "downloads");
fs.mkdirSync(downloadsRoot, { recursive: true });
app.register(fastifyStatic, {
  root: downloadsRoot,
  prefix: "/web/downloads/",
  decorateReply: false,
});

// Handy redirect so you can just open http://localhost:3001/
app.get("/", async (req, reply) => {
  return reply.redirect("/web/index.html");
});

async function listenWithFallback() {
  if (process.env.IRONLOG_DESKTOP !== "1") {
    await app.listen({ port: PORT, host: HOST });
    process.env.PORT_EFFECTIVE = String(PORT);
    return PORT;
  }

  const start = Number.isFinite(PORT) ? PORT : 3001;
  const end = Number.isFinite(DESKTOP_PORT_MAX) ? DESKTOP_PORT_MAX : start;
  let lastErr = null;
  for (let p = start; p <= end; p += 1) {
    try {
      await app.listen({ port: p, host: HOST });
      process.env.PORT_EFFECTIVE = String(p);
      return p;
    } catch (err) {
      lastErr = err;
      if (err?.code !== "EADDRINUSE") throw err;
    }
  }
  throw lastErr || new Error("No free desktop API port found");
}

try {
  const effectivePort = await listenWithFallback();
  // One-time tidy-up of site codes (utils/sites.js): copies the database first, audits every change.
  try {
    ensureAuditTable(db);
    const system = { headers: { "x-user-name": "system", "x-user-role": "system" } };
    await repairSiteDataOnce(db, {
      backupDir: backupDirPath(),
      audit: (entry) => writeAudit(db, system, { module: "sites", ...entry }),
      log: app.log,
    });
  } catch (err) {
    app.log.warn?.(`[sites] tidy-up skipped: ${err?.message || err}`);
  }
  const stopDbAutoBackup = startDbAutoBackup(app.log);
  process.once("SIGINT", stopDbAutoBackup);
  process.once("SIGTERM", stopDbAutoBackup);
  // Nightly off-site backup and monthly restore test; failures alert the administrators.
  startBackupScheduler({
    db,
    backupDir: backupDirPath(),
    profile: () => getCompanyProfile(db),
    notify: (message) => sendPushToRoles({ roles: ["admin"], kind: "backup_failed", title: "IronLog backup problem", body: message, data: { screen: "admin" } }),
    log: app.log,
  });
  bootstrapSmtpFromEnv({ log: app.log });
  await runIronmindAutoScheduler(app.log);
  runCartrackAutoScheduler(app.log);
  runUnitechAutoScheduler(app.log);
  runFamsAutoScheduler(app.log).catch((err) => {
    app.log.warn?.(`[FAMS] scheduler failed to start: ${err?.message || err}`);
  });
  app.log.info(`IRONLOG API running on http://${HOST}:${effectivePort}`);
  app.log.info(`IRONLOG UI  running on http://${HOST}:${effectivePort}/web/index.html`);
} catch (err) {
  app.log.error(err);
  if (process.env.IRONLOG_DESKTOP === "1") throw err;
  process.exit(1);
}