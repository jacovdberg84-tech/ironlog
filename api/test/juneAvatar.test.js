import assert from "node:assert/strict";
import test from "node:test";
import { __test, getJuneAvatarStatus } from "../utils/juneAvatar.js";

const KEYS = ["LEMONSLICE_API_KEY", "LIVEKIT_URL", "LIVEKIT_API_KEY", "LIVEKIT_API_SECRET"];

function withEnv(values, callback) {
  const previous = Object.fromEntries(KEYS.map((key) => [key, process.env[key]]));
  try {
    for (const key of KEYS) {
      if (values[key] === undefined) delete process.env[key];
      else process.env[key] = values[key];
    }
    callback();
  } finally {
    for (const key of KEYS) {
      if (previous[key] === undefined) delete process.env[key];
      else process.env[key] = previous[key];
    }
  }
}

test("June avatar reports only missing configuration names, never credentials", () => {
  withEnv({ LIVEKIT_URL: "wss://demo.livekit.cloud" }, () => {
    const status = getJuneAvatarStatus();
    assert.equal(status.configured, false);
    assert.deepEqual(status.missing.sort(), ["LEMONSLICE_API_KEY", "LIVEKIT_API_KEY", "LIVEKIT_API_SECRET"]);
    assert.equal(Object.hasOwn(status, "api_key"), false);
  });
});

test("June avatar creates separate signed LiveKit viewer and publisher tokens", () => {
  const makeToken = (canPublish, canSubscribe) => __test.createLiveKitToken({
    apiKey: "API-test-key",
    apiSecret: "test-secret",
    room: "ironlog-june-test",
    identity: canPublish ? "lemonslice-june" : "viewer-jaco",
    name: "June test",
    canPublish,
    canSubscribe,
  });
  const read = (token) => JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString("utf8"));
  const publisher = read(makeToken(true, true));
  const viewer = read(makeToken(false, true));
  assert.equal(publisher.video.canPublish, true);
  assert.equal(viewer.video.canPublish, false);
  assert.equal(viewer.video.canSubscribe, true);
  assert.equal(publisher.video.room, "ironlog-june-test");
});

test("June asks LemonSlice for the agent built in its web app and reports its refusal reason", async () => {
  const { startJuneAvatar } = await import("../utils/juneAvatar.js");
  const env = { LEMONSLICE_API_KEY: "ls-key", LIVEKIT_URL: "wss://demo.livekit.cloud", LIVEKIT_API_KEY: "lk", LIVEKIT_API_SECRET: "secret" };
  const previous = Object.fromEntries([...Object.keys(env), "JUNE_LEMONSLICE_SOURCE"].map((k) => [k, process.env[k]]));
  Object.assign(process.env, env);
  const realFetch = global.fetch;
  const bodies = [];
  try {
    global.fetch = async (url, init) => {
      bodies.push(JSON.parse(init.body));
      return new Response(JSON.stringify({ detail: "Agent not found" }), { status: 404 });
    };
    await assert.rejects(() => startJuneAvatar({ owner: "jaco" }), /HTTP 404: Agent not found/);
    assert.equal(bodies[0].agent_id, "agent_1cadda06586e6670");
    assert.equal(Object.hasOwn(bodies[0], "agent_image_url"), false, "exactly one avatar source");
    assert.equal(getJuneAvatarStatus().last_error.status, 404);

    process.env.JUNE_LEMONSLICE_SOURCE = "image";
    await assert.rejects(() => startJuneAvatar({ owner: "jaco" }));
    assert.match(bodies[1].agent_image_url, /june-portrait\.png$/);
    assert.equal(Object.hasOwn(bodies[1], "agent_id"), false);
  } finally {
    global.fetch = realFetch;
    for (const [k, v] of Object.entries(previous)) {
      if (v === undefined) delete process.env[k];
      else process.env[k] = v;
    }
  }
});
