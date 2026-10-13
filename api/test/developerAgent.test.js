import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import Fastify from "fastify";
import aiEngineerRoutes, { developerAgentStatus } from "../routes/aiEngineer.routes.js";

test("developer agent status reports the coder without exposing a key", () => {
  const status = developerAgentStatus();
  assert.equal(status.agent, "coder");
  assert.equal(status.name, "IronLog Coder");
  assert.ok(status.coder_model);
  assert.ok(status.planner_model);
  assert.equal(typeof status.coder_ready, "boolean");
  assert.ok(status.max_tool_rounds >= 16);
  assert.ok(status.sandbox.includes("web/"));
  assert.equal(JSON.stringify(status).includes("sk-"), false);
  assert.equal(Object.hasOwn(status, "api_key"), false);
});

test("developer agent status is limited to engineering roles", async () => {
  const app = Fastify();
  await app.register(aiEngineerRoutes, { prefix: "/api/ai-engineer" });
  const denied = await app.inject({
    method: "GET",
    url: "/api/ai-engineer/status",
    headers: { "x-user-role": "operator" },
  });
  assert.equal(denied.statusCode, 403);

  const allowed = await app.inject({
    method: "GET",
    url: "/api/ai-engineer/status",
    headers: { "x-user-role": "admin" },
  });
  assert.equal(allowed.statusCode, 200);
  assert.equal(allowed.json().agent, "coder");
  await app.close();
});

test("the dashboard exposes an admin Developer tab for the coding agent", () => {
  const html = fs.readFileSync(new URL("../../web/index.html", import.meta.url), "utf8");
  const core = fs.readFileSync(new URL("../../web/app/core.js", import.meta.url), "utf8");
  assert.match(html, /data-tab="developer"/);
  assert.match(html, /id="tab-developer"/);
  assert.match(html, /id="devCreateBtn"/);
  assert.match(html, /app\/developer\.js/);
  assert.match(core, /"developer"/);
  assert.match(core, /opt\.value === "admin" \|\| opt\.value === "developer"/);
  assert.match(core, /!list\.includes\("developer"\)/);
});
