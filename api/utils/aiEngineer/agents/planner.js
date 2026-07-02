import {
  getChatModel,
  isOpenAiCompatibleConfigured,
  openAiCompatibleChatCompletion,
} from "../../llmChat.js";
import { PLANNER_MODEL } from "../config.js";
import { chooseTargetFiles } from "../heuristics.js";
import { extractJsonObject } from "../parse.js";
import { summarizeRepoTree } from "../tools.js";

function buildProposalMarkdown(requestRow, plan) {
  const title = String(requestRow?.title || "").trim();
  const requestText = String(requestRow?.request_text || "").trim();
  const pri = String(requestRow?.priority || "medium");
  const risk = String(plan?.risk_level || requestRow?.risk_level || "medium");
  const steps = Array.isArray(plan?.implementation_steps) ? plan.implementation_steps : [];
  const criteria = Array.isArray(plan?.acceptance_criteria) ? plan.acceptance_criteria : [];
  const targetFiles = Array.isArray(plan?.target_files) ? plan.target_files : [];
  return [
    `# AI Engineer Proposal`,
    ``,
    `## Request`,
    `- ID: ${Number(requestRow?.id || 0)}`,
    `- Title: ${title}`,
    `- Priority: ${pri}`,
    `- Risk: ${risk}`,
    ``,
    `### User request`,
    requestText || "(empty)",
    ``,
    `## Target files`,
    ...targetFiles.map((f) => `- ${f}`),
    ``,
    `## Implementation steps`,
    ...(steps.length ? steps.map((s) => `- ${s}`) : ["- Implement minimal, backwards-compatible changes."]),
    ``,
    `## Acceptance criteria`,
    ...(criteria.length ? criteria.map((s) => `- ${s}`) : ["- Change satisfies the user request without breaking existing flows."]),
    ``,
    `## Validation gates`,
    `- Coder agent produces git-applyable diff`,
    `- Server import validation in worktree`,
    `- Human approval before merge`,
    ``,
  ].join("\n");
}

function normalizePlan(plan, requestRow, repoRoot) {
  const fallbackFiles = chooseTargetFiles(requestRow?.title, requestRow?.request_text);
  const target_files = Array.isArray(plan?.target_files)
    ? plan.target_files.map((f) => String(f || "").replace(/\\/g, "/")).filter(Boolean)
    : fallbackFiles;
  const risk_level = ["low", "medium", "high"].includes(String(plan?.risk_level || "").toLowerCase())
    ? String(plan.risk_level).toLowerCase()
    : String(requestRow?.risk_level || "medium");
  const implementation_steps = Array.isArray(plan?.implementation_steps)
    ? plan.implementation_steps.map((x) => String(x || "").trim()).filter(Boolean).slice(0, 12)
    : [];
  const acceptance_criteria = Array.isArray(plan?.acceptance_criteria)
    ? plan.acceptance_criteria.map((x) => String(x || "").trim()).filter(Boolean).slice(0, 12)
    : [];
  const markdown_preview = String(plan?.markdown_preview || "").trim() || buildProposalMarkdown(requestRow, {
    target_files,
    risk_level,
    implementation_steps,
    acceptance_criteria,
  });
  return {
    agent: "planner",
    model: PLANNER_MODEL,
    target_files: target_files.length ? target_files : fallbackFiles,
    risk_level,
    implementation_steps,
    acceptance_criteria,
    markdown_preview,
    repo_file_sample: summarizeRepoTree(repoRoot).slice(0, 40),
  };
}

export async function runPlannerAgent({ requestRow, repoRoot }) {
  if (!isOpenAiCompatibleConfigured()) {
    const plan = normalizePlan(null, requestRow, repoRoot);
    return {
      ok: true,
      summary: "Planner fallback: heuristic file map (LLM not configured).",
      details: {
        ...plan,
        fallback: true,
        note: "Set OPENAI_API_KEY or OLLAMA_HOST for live planner agent.",
      },
    };
  }

  const fileSample = summarizeRepoTree(repoRoot);
  const system = [
    "You are the IRONLOG Planner Agent.",
    "Output ONLY valid JSON (no markdown prose).",
    "Schema:",
    "{",
    '  "target_files": ["web/...", "api/routes/..."],',
    '  "risk_level": "low|medium|high",',
    '  "implementation_steps": ["..."],',
    '  "acceptance_criteria": ["..."],',
    '  "markdown_preview": "# AI Engineer Proposal\\n..."',
    "}",
    "Rules:",
    "- Prefer 1-5 files under web/ and api/routes/.",
    "- Keep changes minimal and backwards compatible.",
    "- markdown_preview should be a concise proposal markdown string.",
  ].join("\n");

  const user = [
    `REQUEST ID: ${Number(requestRow?.id || 0)}`,
    `TITLE: ${String(requestRow?.title || "")}`,
    `PRIORITY: ${String(requestRow?.priority || "medium")}`,
    `REQUEST: ${String(requestRow?.request_text || "")}`,
    "",
    "REPO FILE SAMPLE:",
    ...fileSample.slice(0, 80).map((f) => `- ${f}`),
  ].join("\n");

  const resp = await openAiCompatibleChatCompletion({
    model: PLANNER_MODEL || getChatModel(),
    temperature: 0.2,
    messages: [
      { role: "system", content: system },
      { role: "user", content: user },
    ],
  });

  const raw = String(resp?.choices?.[0]?.message?.content || "").trim();
  const parsed = extractJsonObject(raw);
  if (!parsed) {
    const plan = normalizePlan(null, requestRow, repoRoot);
    return {
      ok: true,
      summary: "Planner fallback: model returned non-JSON plan.",
      details: {
        ...plan,
        fallback: true,
        raw_preview: raw.slice(0, 2000),
      },
    };
  }

  const plan = normalizePlan(parsed, requestRow, repoRoot);
  return {
    ok: true,
    summary: `Planner agent produced ${plan.target_files.length} target file(s).`,
    details: {
      ...plan,
      proposal: {
        target_files: plan.target_files,
        markdown_preview: plan.markdown_preview,
      },
      raw_preview: raw.slice(0, 2000),
    },
  };
}

export { buildProposalMarkdown };
