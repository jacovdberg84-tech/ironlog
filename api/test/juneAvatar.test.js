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
