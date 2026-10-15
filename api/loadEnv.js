// IRONLOG/api/loadEnv.js — reads server settings once, before anything else needs them.
//
// Normally api/.env (then a .env in the working folder fills any gaps), and
// settings already in the process environment win. A customer instance sharing
// this code (deploy/ironlog-instance.sh) is started with IRONLOG_ENV_FILE
// pointing at its own settings file. Then ONLY that file counts: IronLog
// settings inherited from the shell that started it (a deploy carries AML's AI
// key, mailbox and GitHub token) are dropped first, so another company's keys
// can never leak in.
import dotenv from "dotenv";
import path from "node:path";
import { fileURLToPath } from "node:url";

const apiDir = path.dirname(fileURLToPath(import.meta.url));

// Names IronLog reads for keys, connections and storage; system settings (PATH, HOME, NODE_*) are kept.
const APP_SETTING = /^(IRONLOG_|IRONMIND_|BORRIS_|JUNE_|LEMONSLICE_|LIVEKIT_|OPENAI_|AZURE_|FOUNDRY_|USE_FOUNDRY|AI_|LLM_|OLLAMA_|CARTRACK_|UNITECH_|VISIONLINK_|FAMS_|FCM_|SMTP_|TELEMATICS_|FSC_TELEMATICS_|INSPECTPRO_|REPORT_|DOC_|GITHUB_TOKEN$|GH_TOKEN$|DB_PATH$|PORT$|HOST$|SITE_CODE$)/;

export function settingsFiles(env = process.env, cwd = process.cwd()) {
  if (env.IRONLOG_ENV_FILE) return [path.resolve(env.IRONLOG_ENV_FILE)];
  return [path.join(apiDir, ".env"), path.join(cwd, ".env")];
}

/** Instance mode: removes inherited IronLog settings, keeping IRONLOG_ENV_FILE. Returns the names removed. */
export function dropInheritedSettings(env = process.env) {
  if (!env.IRONLOG_ENV_FILE) return [];
  const removed = Object.keys(env).filter((k) => k !== "IRONLOG_ENV_FILE" && APP_SETTING.test(k));
  for (const k of removed) delete env[k];
  return removed;
}

dropInheritedSettings();
for (const file of settingsFiles()) dotenv.config({ path: file });
