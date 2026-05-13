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

  const controller = timeoutMs > 0 ? new AbortController() : null;
  const timer = controller
    ? setTimeout(() => {
        controller.abort();
      }, timeoutMs)
    : null;

  let res;
  try {
    res = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(payload),
      signal: controller?.signal,
    });
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

  const data = await res.json().catch(() => ({}));
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
