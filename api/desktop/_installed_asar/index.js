// IRONLOG/api/index.js
import { buildServer } from "./server.js";
import { runHybridPhase1Migration } from "./db/hybridPhase1.js";
import { runHybridPhase2Migration } from "./db/hybridPhase2.js";

import fastifyStatic from "@fastify/static";
import path from "path";
import fs from "fs";
import { fileURLToPath } from "url";

const PORT = process.env.PORT ? Number(process.env.PORT) : 3001;
const HOST = process.env.HOST || "0.0.0.0";

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

// Handy redirect so you can just open http://localhost:3001/
app.get("/", async (req, reply) => {
  return reply.redirect("/web/index.html");
});

try {
  await app.listen({ port: PORT, host: HOST });
  app.log.info(`IRONLOG API running on http://${HOST}:${PORT}`);
  app.log.info(`IRONLOG UI  running on http://${HOST}:${PORT}/web/index.html`);
} catch (err) {
  app.log.error(err);
  process.exit(1);
}