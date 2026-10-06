import test from "node:test";
import assert from "node:assert/strict";
import { buildBorrisResponsesInput, isDirectOpenAiResponsesConfigured } from "../utils/openaiResponses.js";

function withEnv(values, run) {
  const keys = ["OPENAI_API_KEY", "OPENAI_BASE_URL", "LLM_BASE_URL", "OLLAMA_HOST", "OLLAMA_BASE_URL"];
  const before = Object.fromEntries(keys.map((key) => [key, process.env[key]]));
  try {
    for (const key of keys) {
      if (Object.hasOwn(values, key)) {
        if (values[key] == null) delete process.env[key];
        else process.env[key] = values[key];
      }
    }
    return run();
  } finally {
    for (const key of keys) {
      if (before[key] == null) delete process.env[key];
      else process.env[key] = before[key];
    }
  }
}

test("Borris Responses is enabled only for a direct OpenAI configuration", () => {
  withEnv({ OPENAI_API_KEY: "sk-test", OPENAI_BASE_URL: null, LLM_BASE_URL: null, OLLAMA_HOST: null, OLLAMA_BASE_URL: null }, () => {
    assert.equal(isDirectOpenAiResponsesConfigured(), true);
  });
  withEnv({ OPENAI_API_KEY: "sk-test", OPENAI_BASE_URL: "https://api.groq.com/openai/v1" }, () => {
    assert.equal(isDirectOpenAiResponsesConfigured(), false);
  });
});

test("Borris Responses input preserves conversational roles as plain text", () => {
  assert.equal(
    buildBorrisResponsesInput([{ role: "user", content: "Status for G01AM" }, { role: "assistant", content: "One open work order." }]),
    "USER:\nStatus for G01AM\n\nASSISTANT:\nOne open work order.",
  );
});
