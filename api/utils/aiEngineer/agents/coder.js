import OpenAI from "openai";
import {
  getChatModel,
  isOpenAiCompatibleConfigured,
  openAiCompatibleChatCompletion,
} from "../../llmChat.js";
import { CODER_MAX_TOOL_ROUNDS, CODER_MODEL } from "../config.js";
import { chooseTargetFiles } from "../heuristics.js";
import { extractUnifiedDiff } from "../parse.js";
import { createToolRegistry } from "../tools.js";

let openaiClient = null;

function getOpenAIClient() {
  const apiKey = String(process.env.OPENAI_API_KEY || "").trim();
  if (!apiKey) return null;
  if (!openaiClient) openaiClient = new OpenAI({ apiKey });
  return openaiClient;
}

function dispatchTool(registry, name, args) {
  const handler = registry.handlers[name];
  if (!handler) return { error: `Unknown tool: ${name}` };
  try {
    return handler(args || {});
  } catch (err) {
    return { error: String(err?.message || err) };
  }
}

async function runCoderToolLoop({ requestRow, worktreePath, targetFiles }) {
  const client = getOpenAIClient();
  if (!client) {
    return { ok: false, error: "OPENAI_API_KEY required for coder tool loop.", model: CODER_MODEL };
  }

  const registry = createToolRegistry({ repoRoot: worktreePath, worktreePath });
  const system = [
    "You are the IRONLOG Coder Agent.",
    "Implement the user request with minimal, backwards-compatible edits.",
    "Use tools to read/search files before editing.",
    "When ready, call git_apply_check with a unified diff starting with diff --git.",
    "Only edit files under web/ and api/.",
    `Target files: ${targetFiles.join(", ")}`,
  ].join("\n");

  const user = [
    `REQUEST ID: ${Number(requestRow?.id || 0)}`,
    `TITLE: ${String(requestRow?.title || "")}`,
    `REQUEST: ${String(requestRow?.request_text || "")}`,
    "",
    "Produce a valid unified diff via git_apply_check.",
  ].join("\n");

  const messages = [
    { role: "system", content: system },
    { role: "user", content: user },
  ];

  for (let round = 0; round < CODER_MAX_TOOL_ROUNDS; round++) {
    const resp = await client.chat.completions.create({
      model: CODER_MODEL,
      temperature: 0.1,
      messages,
      tools: registry.openAiToolDefs,
      tool_choice: "auto",
    });

    const msg = resp.choices?.[0]?.message;
    if (!msg) {
      return { ok: false, error: "Coder agent returned no message.", model: CODER_MODEL };
    }

    messages.push(msg);

    const toolCalls = msg.tool_calls || [];
    if (!toolCalls.length) {
      const diff = extractUnifiedDiff(msg.content || "");
      if (diff) {
        const check = registry.handlers.git_apply_check({ diff });
        if (check.ok) {
          return {
            ok: true,
            model: CODER_MODEL,
            diff: registry.state.finalizedDiff || diff,
            agent: "coder",
            tool_trace: registry.state.tool_trace,
            raw_preview: String(msg.content || "").slice(0, 2000),
          };
        }
      }
      return {
        ok: false,
        error: "Coder stopped without a valid diff.",
        model: CODER_MODEL,
        raw_preview: String(msg.content || "").slice(0, 2000),
        tool_trace: registry.state.tool_trace,
      };
    }

    for (const call of toolCalls) {
      const name = call.function?.name;
      let args = {};
      try {
        args = JSON.parse(call.function?.arguments || "{}");
      } catch {
        args = {};
      }
      const output = dispatchTool(registry, name, args);
      if (name === "git_apply_check" && output?.ok && registry.state.finalizedDiff) {
        return {
          ok: true,
          model: CODER_MODEL,
          diff: registry.state.finalizedDiff,
          agent: "coder",
          tool_trace: registry.state.tool_trace,
        };
      }
      messages.push({
        role: "tool",
        tool_call_id: call.id,
        content: JSON.stringify(output),
      });
    }
  }

  return {
    ok: false,
    error: "Coder agent exceeded max tool rounds.",
    model: CODER_MODEL,
    tool_trace: registry.state.tool_trace,
  };
}

async function runCoderSingleShot({ requestRow, worktreePath, targetFiles }) {
  const registry = createToolRegistry({ repoRoot: worktreePath, worktreePath });
  const fileBodies = targetFiles
    .map((rel) => {
      try {
        const { content } = registry.handlers.read_file({ path: rel });
        return `FILE: ${rel}\n-----\n${content}\n-----`;
      } catch {
        return `FILE: ${rel}\n-----\n(missing)\n-----`;
      }
    })
    .join("\n\n");

  const system = [
    "You are the IRONLOG Coder Agent.",
    "Output ONLY a unified diff in a ```diff fenced block.",
    "Keep edits minimal and limited to TARGET FILES.",
  ].join("\n");

  const user = [
    `REQUEST ID: ${Number(requestRow?.id || 0)}`,
    `TITLE: ${String(requestRow?.title || "")}`,
    `REQUEST: ${String(requestRow?.request_text || "")}`,
    "",
    "TARGET FILES:",
    ...targetFiles.map((f) => `- ${f}`),
    "",
    fileBodies,
  ].join("\n");

  const resp = await openAiCompatibleChatCompletion({
    model: CODER_MODEL || getChatModel(),
    temperature: 0.1,
    messages: [
      { role: "system", content: system },
      { role: "user", content: user },
    ],
  });

  const raw = String(resp?.choices?.[0]?.message?.content || "").trim();
  const diff = extractUnifiedDiff(raw);
  if (!diff) {
    return {
      ok: false,
      error: "Coder returned no unified diff.",
      model: CODER_MODEL || getChatModel(),
      raw_preview: raw.slice(0, 2000),
      agent: "coder",
      fallback: true,
    };
  }

  const check = registry.handlers.git_apply_check({ diff });
  if (!check.ok) {
    return {
      ok: false,
      error: "Generated diff failed git apply --check.",
      model: CODER_MODEL || getChatModel(),
      diff,
      raw_preview: raw.slice(0, 2000),
      agent: "coder",
      fallback: true,
      tool_trace: registry.state.tool_trace,
    };
  }

  return {
    ok: true,
    model: CODER_MODEL || getChatModel(),
    diff: registry.state.finalizedDiff || diff,
    raw_preview: raw.slice(0, 2000),
    agent: "coder",
    fallback: true,
    tool_trace: registry.state.tool_trace,
  };
}

export async function runCoderAgent({ requestRow, worktreePath, targetFiles, planDetails = null }) {
  const files = Array.isArray(targetFiles) && targetFiles.length
    ? targetFiles
    : Array.isArray(planDetails?.target_files) && planDetails.target_files.length
      ? planDetails.target_files
      : chooseTargetFiles(requestRow?.title, requestRow?.request_text);

  if (!worktreePath) {
    return { ok: false, error: "worktreePath required for coder agent." };
  }

  if (getOpenAIClient()) {
    const result = await runCoderToolLoop({ requestRow, worktreePath, targetFiles: files });
    if (result.ok) return result;
    if (isOpenAiCompatibleConfigured()) {
      const fallback = await runCoderSingleShot({ requestRow, worktreePath, targetFiles: files });
      if (fallback.ok) {
        return {
          ...fallback,
          summary_note: `Tool loop failed (${result.error}); used single-shot fallback.`,
        };
      }
    }
    return result;
  }

  if (!isOpenAiCompatibleConfigured()) {
    return {
      ok: false,
      error: "OPENAI_API_KEY or LLM base URL required for coder agent.",
      model: CODER_MODEL,
    };
  }

  return runCoderSingleShot({ requestRow, worktreePath, targetFiles: files });
}
