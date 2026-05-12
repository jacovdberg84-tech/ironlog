/**
 * OpenAI-compatible chat completions (/v1/chat/completions).
 * Supports OpenAI Cloud, proxies, and Ollama (set OLLAMA_HOST or OPENAI_BASE_URL).
 */

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

/**
 * POST chat completions. Returns parsed JSON body or null on HTTP/error parse failure.
 */
export async function openAiCompatibleChatCompletion(body) {
  const url = resolveOpenAiCompatibleChatUrl();
  const headers = { "Content-Type": "application/json" };
  const key = String(process.env.OPENAI_API_KEY || process.env.LLM_API_KEY || "").trim();
  if (key) headers.Authorization = `Bearer ${key}`;

  let res;
  try {
    res = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
    });
  } catch (e) {
    console.warn("[llmChat] fetch failed:", e?.message || e);
    return null;
  }

  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const msg = data?.error?.message || data?.error || `${res.status}`;
    const safeUrl = url.replace(/^(https?:\/\/[^/]+).*/, "$1/…/chat/completions");
    console.warn("[llmChat]", safeUrl, res.status, typeof msg === "string" ? msg : JSON.stringify(msg));
    return null;
  }
  return data;
}

export function chatEndpointSummaryForLogs() {
  const url = resolveOpenAiCompatibleChatUrl();
  return url.replace(/^(https?:\/\/[^/]+).*/, "$1/…");
}
