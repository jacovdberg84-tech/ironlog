// IRONLOG/api/index.js
import { buildServer } from "./server.js";
import { runHybridPhase1Migration } from "./db/hybridPhase1.js";
import { runHybridPhase2Migration } from "./db/hybridPhase2.js";
import "./db/migrate.js"; // Run schema migration first

import dotenv from "dotenv";
import fastifyStatic from "@fastify/static";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";

// Load environment variables from .env in the API root.
dotenv.config({ path: path.join(process.cwd(), ".env") });

const PORT = process.env.PORT ? Number(process.env.PORT) : 3001;
const HOST = process.env.HOST || "0.0.0.0";
const DESKTOP_PORT_MAX = process.env.IRONLOG_DESKTOP_PORT_MAX
  ? Number(process.env.IRONLOG_DESKTOP_PORT_MAX)
  : PORT;

runHybridPhase1Migration();
runHybridPhase2Migration();

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
  app.log.info(`IRONLOG API running on http://${HOST}:${effectivePort}`);
  app.log.info(`IRONLOG UI  running on http://${HOST}:${effectivePort}/web/index.html`);
} catch (err) {
  app.log.error(err);
  if (process.env.IRONLOG_DESKTOP === "1") throw err;
  process.exit(1);
}