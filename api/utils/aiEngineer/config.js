import { getChatModel } from "../llmChat.js";

export const ALLOWED_PATH_PREFIXES = [
  "web/",
  "api/routes/",
  "api/utils/",
  "api/server.js",
];

export const PLANNER_MODEL =
  String(process.env.AI_ENGINEER_PLANNER_MODEL || "").trim() ||
  String(process.env.LLM_MODEL || "").trim() ||
  getChatModel();

export const CODER_MODEL =
  String(process.env.AI_ENGINEER_CODER_MODEL || "").trim() ||
  String(process.env.AI_ENGINEER_MODEL || "").trim() ||
  "gpt-4o";

export const CODER_MAX_TOOL_ROUNDS = Number(process.env.AI_ENGINEER_CODER_MAX_ROUNDS || 10);

export function isAllowedRelativePath(relPath) {
  const p = String(relPath || "").replace(/\\/g, "/").replace(/^\/+/, "");
  if (!p || p.includes("..")) return false;
  if (p.startsWith(".ai-engineer/")) return false;
  if (/\.env/i.test(p) || /credentials/i.test(p) || /\/db\//i.test(p)) return false;
  return ALLOWED_PATH_PREFIXES.some((prefix) => p === prefix || p.startsWith(prefix));
}
