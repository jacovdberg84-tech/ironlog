/**
 * OpenAI-compatible chat completions (/v1/chat/completions).
 * Supports OpenAI Cloud, proxies, and Ollama (set OLLAMA_HOST or OPENAI_BASE_URL).
 */

let lastLlmChatError = "";

export function getLastLlmChatError() {
  return lastLlmChatError;
}

export function clearLastLlmChatError() {
  lastLlmChatError = "";
}

export function getChatModel() {
  return String(process.env.LLM_MODEL || process.env.OPENAI_MODEL || "gpt-4o-mini").trim() || "gpt-4o-mini";
}

export function usesCustomChatBase() {
  return Boolean(
    String(process.env.OPENAI_BASE_URL || "").trim() ||
      String(process.env.LLM_BASE_URL || "").trim() ||
      String(process.env.OLLAMA_HOST || "").trim() ||
      String(process.env.OLLAMA_BASE_URL || "").trim(),
  );
}

/** True when the resolved chat URL is almost certainly local Ollama (OpenAI-compat on :11434). */
export function isOllamaChatUrl(url) {
  try {
    const u = new URL(url);
    return String(u.port || "") === "11434";
  } catch {
    return false;
  }
}

/**
 * Ollama model tags often include :latest; bare names like "llama3.2" may 404 on /v1/chat/completions.
 */
export function normalizeModelForOllamaUrl(model, url) {
  const m = String(model || "").trim();
  if (!m || !isOllamaChatUrl(url)) return m;
  if (m.includes(":") || m.includes("/")) return m;
  return `${m}:latest`;
}

/**
 * Full URL for POST (OpenAI-compatible).
 * Priority: OPENAI_BASE_URL / LLM_BASE_URL → OLLAMA_HOST → default OpenAI cloud.
 */
export function resolveOpenAiCompatibleChatUrl() {
  const explicit = String(process.env.OPENAI_BASE_URL || process.env.LLM_BASE_URL || "").trim();
  if (explicit) {
    let base = explicit.replace(/\/+$/, "");
    if (!/\/v1$/i.test(base)) base = `${base}/v1`;
    return `${base}/chat/completions`;
  }
  const ollama = String(process.env.OLLAMA_HOST || process.env.OLLAMA_BASE_URL || "").trim();
  if (ollama) {
    let base = ollama.replace(/\/+$/, "");
    if (!/\/v1$/i.test(base)) base = `${base}/v1`;
    return `${base}/chat/completions`;
  }
  return "https://api.openai.com/v1/chat/completions";
}

/** True if OPENAI_API_KEY is set, or a non-default chat base is configured (e.g. Ollama). */
export function isOpenAiCompatibleConfigured() {
  if (String(process.env.OPENAI_API_KEY || "").trim()) return true;
  return usesCustomChatBase();
}

function getRequestTimeoutMs(body) {
  const bodyTimeout = Number(body?.timeout_ms);
  if (Number.isFinite(bodyTimeout) && bodyTimeout > 0) return bodyTimeout;
  const envTimeout = Number(process.env.LLM_TIMEOUT_MS || process.env.OLLAMA_TIMEOUT_MS || 0);
  if (Number.isFinite(envTimeout) && envTimeout > 0) return envTimeout;
  return 0;
}

/**
 * POST chat completions. Returns parsed JSON body or null on HTTP/error parse failure.
 */
/** Ollama's own chat endpoint for an OpenAI-style chat URL on the same host. */
export function ollamaNativeChatUrl(openAiUrl) {
  const u = new URL(openAiUrl);
  return `${u.origin}/api/chat`;
}

/**
 * Context window for Borris on Ollama (tokens). Off (0) unless BORRIS_NUM_CTX is
 * set: a different window makes Ollama reload the model with more memory, which
 * a small server may not have. 0 keeps the model's own default.
 */
export function borrisNumCtx() {
  const n = Number(process.env.BORRIS_NUM_CTX ?? 0);
  return Number.isFinite(n) && n > 0 ? Math.min(131072, Math.round(n)) : 0;
}

/**
 * Ollama's OpenAI-compatible endpoint ignores the context size, so callers that
 * pass num_ctx (larger evidence) or json: true go to /api/chat instead. The reply
 * is returned in the OpenAI shape so callers do not care which one answered.
 */
async function ollamaNativeChat(url, payload, timeoutMs) {
  const options = { num_ctx: payload.num_ctx };
  if (payload.temperature != null) options.temperature = payload.temperature;
  if (payload.max_tokens != null) options.num_predict = payload.max_tokens;
  const nativeBody = {
    model: payload.model,
    messages: payload.messages,
    stream: false,
    options,
    ...(payload.json ? { format: "json" } : {}),
  };
  const controller = timeoutMs > 0 ? new AbortController() : null;
  const timer = controller ? setTimeout(() => controller.abort(), timeoutMs) : null;
  try {
    const res = await fetch(ollamaNativeChatUrl(url), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(nativeBody),
      signal: controller?.signal,
    });
    const data = await res.json().catch(() => null);
    if (!res.ok || !data) {
      lastLlmChatError = `Ollama HTTP ${res.status}: ${data?.error || "no body"}`;
      return null;
    }
    const content = data?.message?.content;
    if (typeof content !== "string" || !content.trim()) {
      lastLlmChatError = "Ollama returned no message content";
      return null;
    }
    return { choices: [{ message: { role: "assistant", content } }], model: data.model };
  } catch (e) {
    lastLlmChatError = e?.name === "AbortError" ? `timeout after ${timeoutMs}ms` : `fetch: ${e?.message || e}`;
    return null;
  } finally {
    if (timer) clearTimeout(timer);
  }
}

export async function openAiCompatibleChatCompletion(body) {
  lastLlmChatError = "";
  const url = resolveOpenAiCompatibleChatUrl();
  const toOllama = isOllamaChatUrl(url);
  const timeoutMs = getRequestTimeoutMs(body);
  const headers = { "Content-Type": "application/json" };
  const llmKey = String(process.env.LLM_API_KEY || "").trim();
  const openaiKey = String(process.env.OPENAI_API_KEY || "").trim();
  if (llmKey) headers.Authorization = `Bearer ${llmKey}`;
  else if (openaiKey && !toOllama) headers.Authorization = `Bearer ${openaiKey}`;

  let payload =
    body && typeof body === "object"
      ? { ...body, model: normalizeModelForOllamaUrl(body.model, url) }
      : body;
  if (payload && typeof payload === "object" && Object.prototype.hasOwnProperty.call(payload, "timeout_ms")) {
    delete payload.timeout_ms;
  }
  if (toOllama && payload && (Number(payload.num_ctx) > 0 || payload.json)) {
    if (!(Number(payload.num_ctx) > 0)) delete payload.num_ctx;
    return ollamaNativeChat(url, payload, timeoutMs);
  }
  if (payload && typeof payload === "object") {
    // Not Ollama: num_ctx does not exist and json maps to OpenAI's JSON mode.
    const wantsJson = payload.json;
    delete payload.num_ctx;
    delete payload.json;
    if (wantsJson) payload.response_format = { type: "json_object" };
  }

  const controller = timeoutMs > 0 ? new AbortController() : null;
  const timer = controller
    ? setTimeout(() => {
        controller.abort();
      }, timeoutMs)
    : null;

  let res;
  let data;
  try {
    res = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
      signal: controller?.signal,
    });
    data = await res.json();
  } catch (e) {
    if ((e?.name === "AbortError" || e?.code === "ABORT_ERR") && timeoutMs > 0) {
      lastLlmChatError = `timeout after ${timeoutMs}ms`;
    } else {
      lastLlmChatError = `fetch: ${e?.message || e}`;
    }
    console.warn("[llmChat] fetch failed:", e?.message || e);
    return null;
  } finally {
    if (timer) clearTimeout(timer);
  }


  const safeUrl = url.replace(/^(https?:\/\/[^/]+).*/, "$1/…/chat/completions");

  if (!res.ok) {
    const msg = data?.error?.message || data?.error || `${res.status}`;
    const detail = typeof msg === "string" ? msg : JSON.stringify(msg);
    lastLlmChatError = `HTTP ${res.status}: ${detail}`;
    console.warn("[llmChat]", safeUrl, res.status, detail);
    return null;
  }

  if (data?.error) {
    const msg = typeof data.error === "string" ? data.error : data.error?.message || JSON.stringify(data.error);
    lastLlmChatError = `model: ${msg}`;
    console.warn("[llmChat] response error:", safeUrl, msg);
    return null;
  }

  const content = data?.choices?.[0]?.message?.content;
  if (typeof content !== "string" || !String(content).trim()) {
    lastLlmChatError = "model returned no message content (check LLM_MODEL matches `ollama list`)";
    console.warn("[llmChat] empty content from", safeUrl, JSON.stringify(data).slice(0, 400));
    return null;
  }

  return data;
}

export function chatEndpointSummaryForLogs() {
  const url = resolveOpenAiCompatibleChatUrl();
  return url.replace(/^(https?:\/\/[^/]+).*/, "$1/…");
}


