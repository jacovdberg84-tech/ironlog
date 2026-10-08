import test from "node:test";
import assert from "node:assert/strict";

// The production server runs a Node without a built-in WebSocket: simulate that.
delete globalThis.WebSocket;
const { WebSocketServer } = await import("ws");
const avatar = await import("../utils/juneAvatar.js");

test("June's LemonSlice tunnel works on Node without a built-in WebSocket", async () => {
  assert.equal(avatar.__testSocket.SocketImpl, avatar.__testSocket.WsWebSocket, "falls back to the ws library");
  const server = new WebSocketServer({ port: 0 });
  await new Promise((r) => server.once("listening", r));
  const received = [];
  server.on("connection", (socket) => socket.on("message", (m) => received.push(JSON.parse(String(m)))));
  const env = { LEMONSLICE_API_KEY: "ls", LIVEKIT_URL: "wss://demo.livekit.cloud", LIVEKIT_API_KEY: "lk", LIVEKIT_API_SECRET: "secret" };
  const previous = Object.fromEntries(Object.keys(env).map((k) => [k, process.env[k]]));
  Object.assign(process.env, env);
  const realFetch = global.fetch;
  global.fetch = async () => new Response(JSON.stringify({ websocket_address: `ws://127.0.0.1:${server.address().port}` }), { status: 200 });
  try {
    const session = await avatar.startJuneAvatar({ owner: "jaco" });
    assert.ok(session.id && session.viewer_token);
    avatar.sendJuneAvatarAudio({ id: session.id, owner: "jaco", audio: Buffer.alloc(320).toString("base64") });
    avatar.finishJuneAvatarTurn({ id: session.id, owner: "jaco" });
    avatar.stopJuneAvatar({ id: session.id, owner: "jaco" });
    for (let i = 0; i < 50 && received.length < 3; i++) await new Promise((r) => setTimeout(r, 20));
    assert.deepEqual(received.map((m) => m.command), ["audio", "audio_end", "terminate"]);
    assert.equal(received[0].sampleRate, 16000);
  } finally {
    global.fetch = realFetch;
    for (const [k, v] of Object.entries(previous)) { if (v === undefined) delete process.env[k]; else process.env[k] = v; }
    await new Promise((r) => server.close(r));
  }
});
