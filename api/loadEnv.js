// IRONLOG/api/loadEnv.js — reads server settings once, before anything else needs them.
//
// Normally api/.env (then a .env in the working folder fills any gaps). A
// customer instance sharing this code (deploy/ironlog-instance.sh) is started
// with IRONLOG_ENV_FILE pointing at its own settings file, and then ONLY that
// file is read, so it never inherits another company's keys from api/.env.
// Settings already in the process environment always win.
import dotenv from "dotenv";
import path from "node:path";
import { fileURLToPath } from "node:url";

const apiDir = path.dirname(fileURLToPath(import.meta.url));

export function settingsFiles(env = process.env, cwd = process.cwd()) {
  if (env.IRONLOG_ENV_FILE) return [path.resolve(env.IRONLOG_ENV_FILE)];
  return [path.join(apiDir, ".env"), path.join(cwd, ".env")];
}

for (const file of settingsFiles()) dotenv.config({ path: file });
