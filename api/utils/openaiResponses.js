import OpenAI from "openai";
import { getChatModel, usesCustomChatBase } from "./llmChat.js";

let lastOpenAiResponsesError = "";

/**
 * A direct OpenAI connection is deliberately separate from the existing
 * OpenAI-compatible chat endpoint. This lets Borris use the Responses API
 * without sending a Groq, Ollama, or other provider request to OpenAI.
 */
export function isDirectOpenAiResponsesConfigured() {
  const forceBorrisOpenAi = ["1", "true", "yes", "on"].includes(
    String(process.env.BORRIS_USE_OPENAI_RESPONSES || "").trim().toLowerCase(),
  );
  return Boolean(String(process.env.OPENAI_API_KEY || "").trim()) && (forceBorrisOpenAi || !usesCustomChatBase());
}

export function getLastOpenAiResponsesError() {
  return lastOpenAiResponsesError;
}

export function buildBorrisResponsesInput(messages = []) {
  return messages
    .map((message) => {
      const role = String(message?.role || "user").trim().toUpperCase();
      const content = String(message?.content || "").trim();
      return content ? `${role}:\n${content}` : "";
    })
    .filter(Boolean)
    .join("\n\n");
}

/**
 * Send a short, stateless Borris conversation to OpenAI's Responses API.
 * The caller supplies already-scoped operational context; this helper never
 * reads or writes Ironlog data itself. `store: false` keeps these requests out
 * of stored response history.
 */
export async function openAiResponsesText({
  instructions,
  messages,
  model = getChatModel(),
  maxOutputTokens = 220,
  timeoutMs = 45000,
} = {}) {
  lastOpenAiResponsesError = "";
  if (!isDirectOpenAiResponsesConfigured()) return null;

  const input = buildBorrisResponsesInput(messages);
  if (!input) {
    lastOpenAiResponsesError = "no input supplied";
    return null;
  }

  const safeTimeout = Number.isFinite(Number(timeoutMs)) && Number(timeoutMs) > 0
    ? Math.min(60000, Math.round(Number(timeoutMs)))
    : 45000;
  const safeMaxOutputTokens = Number.isFinite(Number(maxOutputTokens)) && Number(maxOutputTokens) > 0
    ? Math.min(1200, Math.round(Number(maxOutputTokens)))
    : 220;

  try {
    const client = new OpenAI({
      apiKey: String(process.env.OPENAI_API_KEY || "").trim(),
      timeout: safeTimeout,
      maxRetries: 0,
    });
    const response = await client.responses.create({
      model: String(model || getChatModel()).trim() || "gpt-4o-mini",
      instructions: String(instructions || "").trim(),
      input,
      max_output_tokens: safeMaxOutputTokens,
      store: false,
    });
    const text = String(response?.output_text || "").trim();
    if (!text) {
      lastOpenAiResponsesError = "OpenAI Responses returned no output text";
      return null;
    }
    return text;
  } catch (error) {
    const message = String(error?.message || error || "OpenAI Responses request failed");
    lastOpenAiResponsesError = message.slice(0, 500);
    console.warn("[openaiResponses]", lastOpenAiResponsesError);
    return null;
  }
}
