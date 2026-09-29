import test from "node:test";
import assert from "node:assert/strict";
import { borrisNumCtx, ollamaNativeChatUrl, openAiCompatibleChatCompletion } from "../utils/llmChat.js";

test("native Ollama URL and context default", () => {
  assert.equal(ollamaNativeChatUrl("http://127.0.0.1:11434/v1/chat/completions"), "http://127.0.0.1:11434/api/chat");
  const prev = process.env.BORRIS_NUM_CTX;
  delete process.env.BORRIS_NUM_CTX;
  assert.equal(borrisNumCtx(), 8192);
  process.env.BORRIS_NUM_CTX = "16384";
  assert.equal(borrisNumCtx(), 16384);
  if (prev === undefined) delete process.env.BORRIS_NUM_CTX; else process.env.BORRIS_NUM_CTX = prev;
});

test("num_ctx and json requests go to Ollama /api/chat and come back in OpenAI shape", async () => {
  const env = { ...process.env };
  process.env.OLLAMA_HOST = "http://127.0.0.1:11434";
  delete process.env.OPENAI_BASE_URL;
  delete process.env.LLM_BASE_URL;
  const realFetch = globalThis.fetch;
  let seen = null;
  globalThis.fetch = async (url, init) => {
    seen = { url, body: JSON.parse(init.body) };
    return { ok: true, status: 200, json: async () => ({ model: "qwen2.5:7b", message: { content: '{"ok":true}' } }) };
  };
  try {
    const out = await openAiCompatibleChatCompletion({ model: "qwen2.5:7b", max_tokens: 500, temperature: 0, num_ctx: 16384, json: true, messages: [{ role: "user", content: "hi" }] });
    assert.equal(seen.url, "http://127.0.0.1:11434/api/chat");
    assert.deepEqual(seen.body.options, { num_ctx: 16384, temperature: 0, num_predict: 500 });
    assert.equal(seen.body.format, "json");
    assert.equal(seen.body.stream, false);
    assert.equal(out.choices[0].message.content, '{"ok":true}');
  } finally {
    globalThis.fetch = realFetch;
    process.env = env;
  }
});
