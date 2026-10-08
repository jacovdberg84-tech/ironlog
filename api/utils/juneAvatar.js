// IRONLOG/api/utils/juneAvatar.js
// LemonSlice is deliberately only June's visual layer. June's GPT-Live voice,
// Emma Tool Gateway, session auth, and all Ironlog permissions stay in the
// existing OpenAI WebRTC flow.
import crypto from "node:crypto";

const LEMONSLICE_SESSIONS_URL = "https://lemonslice.com/api/liveai/sessions";
const DEFAULT_AGENT_ID = "agent_1cadda06586e6670";
const DEFAULT_IMAGE_URL = "https://ironlog.ironlogafrica.com/web/assets/june-portrait.png";
const SESSION_TTL_MS = 60 * 60_000;
const SESSION_LIMIT = 8;
const sessions = new Map();
let lastError = null;

function value(name, fallback = "") {
  return String(process.env[name] || fallback).trim();
}

function encodeBase64Url(value) {
  return Buffer.from(value).toString("base64url");
}

function signLiveKitToken(payload, secret) {
  const header = encodeBase64Url(JSON.stringify({ alg: "HS256", typ: "JWT" }));
  const body = encodeBase64Url(JSON.stringify(payload));
  const signature = crypto.createHmac("sha256", secret).update(`${header}.${body}`).digest("base64url");
  return `${header}.${body}.${signature}`;
}

function createLiveKitToken({ apiKey, apiSecret, room, identity, name, canPublish, canSubscribe }) {
  const now = Math.floor(Date.now() / 1000);
  return signLiveKitToken({
    iss: apiKey,
    sub: identity,
    name,
    nbf: now - 5,
    iat: now,
    exp: now + Math.floor(SESSION_TTL_MS / 1000),
    video: { roomJoin: true, room, canPublish, canSubscribe },
  }, apiSecret);
}

function config() {
  const livekitUrl = value("LIVEKIT_URL");
  const livekitApiKey = value("LIVEKIT_API_KEY");
  const livekitApiSecret = value("LIVEKIT_API_SECRET");
  const lemonsliceApiKey = value("LEMONSLICE_API_KEY");
  const missing = [
    !lemonsliceApiKey && "LEMONSLICE_API_KEY",
    !livekitUrl && "LIVEKIT_URL",
    !livekitApiKey && "LIVEKIT_API_KEY",
    !livekitApiSecret && "LIVEKIT_API_SECRET",
  ].filter(Boolean);
  return {
    configured: missing.length === 0,
    missing,
    lemonsliceApiKey,
    livekitUrl,
    livekitApiKey,
    livekitApiSecret,
    agentId: value("JUNE_LEMONSLICE_AGENT_ID", DEFAULT_AGENT_ID),
    imageUrl: value("JUNE_LEMONSLICE_IMAGE_URL", DEFAULT_IMAGE_URL),
    // The character built in the LemonSlice web app is used by default; set
    // JUNE_LEMONSLICE_SOURCE=image to animate the portrait image instead.
    source: value("JUNE_LEMONSLICE_SOURCE", "agent").toLowerCase() === "image" ? "image" : "agent",
    prompt: value("JUNE_LEMONSLICE_PROMPT", "June is a sharp, confident executive assistant. Use natural, attentive expressions and understated hand gestures."),
  };
}

export function getJuneAvatarStatus() {
  const settings = config();
  return {
    provider: "lemonslice",
    configured: settings.configured,
    // This is an identifier only, not a credential. It documents the
    // LemonSlice character Jaco created while Ironlog retains June's brain.
    agent_id: settings.agentId,
    source: settings.source === "agent" && settings.agentId ? "agent" : "image",
    missing: settings.missing,
    last_error: lastError,
  };
}

/** LemonSlice's own reason for a refusal, short and without secrets. */
function providerDetail(data, raw) {
  const pick = data?.detail || data?.error?.message || data?.error || data?.message || "";
  const text = typeof pick === "string" ? pick : JSON.stringify(pick);
  return String(text || raw || "").replace(/\s+/g, " ").replace(/(key|token)[^,;]*/gi, "$1 …").trim().slice(0, 200);
}

function avatarError(message, statusCode = 503) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function removeSession(id, { terminate = false } = {}) {
  const session = sessions.get(id);
  if (!session) return;
  sessions.delete(id);
  try {
    if (terminate && session.socket?.readyState === WebSocket.OPEN) {
      session.socket.send(JSON.stringify({ command: "terminate" }));
    }
    session.socket?.close();
  } catch {}
}

function cleanExpiredSessions(now = Date.now()) {
  for (const [id, session] of sessions) {
    if (Number(session?.expiresAt || 0) <= now) removeSession(id, { terminate: true });
  }
}

function openTunnel(address) {
  if (typeof WebSocket !== "function") {
    throw avatarError("This Ironlog server runtime does not support June's avatar tunnel.");
  }
  return new Promise((resolve, reject) => {
    let settled = false;
    let timeoutId = null;
    const settle = (callback, detail) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeoutId);
      callback(detail);
    };
    let socket;
    try {
      socket = new WebSocket(address);
    } catch {
      reject(avatarError("June could not open LemonSlice's visual connection."));
      return;
    }
    timeoutId = setTimeout(() => {
      try { socket.close(); } catch {}
      settle(reject, avatarError("June's visual connection took too long to start.", 504));
    }, 20_000);
    socket.addEventListener("open", () => settle(resolve, socket), { once: true });
    socket.addEventListener("error", () => settle(reject, avatarError("June could not open LemonSlice's visual connection.")), { once: true });
  });
}

function ownerSession(id, owner) {
  cleanExpiredSessions();
  const session = sessions.get(String(id || ""));
  if (!session || session.owner !== owner) throw avatarError("June's visual session was not found. Please start the conversation again.", 404);
  if (session.socket?.readyState !== WebSocket.OPEN) throw avatarError("June's visual session has ended. Her voice session is still available.", 409);
  return session;
}

function send(session, body) {
  session.socket.send(JSON.stringify(body));
}

export async function startJuneAvatar({ owner, log }) {
  const settings = config();
  if (!settings.configured) {
    throw avatarError("June's LemonSlice visual is not configured on the server.");
  }
  cleanExpiredSessions();
  if (sessions.size >= SESSION_LIMIT) {
    throw avatarError("June is already using all available visual sessions. Please try again shortly.", 429);
  }
  const id = crypto.randomUUID();
  const room = `ironlog-june-${crypto.randomUUID().replace(/-/g, "").slice(0, 18)}`;
  const avatarToken = createLiveKitToken({
    apiKey: settings.livekitApiKey,
    apiSecret: settings.livekitApiSecret,
    room,
    identity: "lemonslice-june",
    name: "June visual",
    canPublish: true,
    canSubscribe: true,
  });
  const viewerToken = createLiveKitToken({
    apiKey: settings.livekitApiKey,
    apiSecret: settings.livekitApiSecret,
    room,
    identity: `june-viewer-${id.slice(0, 12)}`,
    name: "Ironlog administrator",
    canPublish: false,
    canSubscribe: true,
  });
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 45_000);
  let response;
  try {
    response = await fetch(LEMONSLICE_SESSIONS_URL, {
      method: "POST",
      headers: { "X-API-Key": settings.lemonsliceApiKey, "Content-Type": "application/json" },
      body: JSON.stringify({
        transport_type: "websocket-livekit",
        livekit_properties: { livekit_url: settings.livekitUrl, livekit_token: avatarToken },
        // LemonSlice takes exactly one avatar source: the agent built in its web
        // app (agent_id) or an image to animate (agent_image_url).
        ...(settings.source === "agent" && settings.agentId ? { agent_id: settings.agentId } : { agent_image_url: settings.imageUrl }),
        agent_prompt: settings.prompt,
      }),
      signal: controller.signal,
    });
  } catch (error) {
    const message = error?.name === "AbortError"
      ? "LemonSlice took too long to prepare June's visual."
      : "Ironlog could not reach LemonSlice for June's visual.";
    throw avatarError(message, error?.name === "AbortError" ? 504 : 502);
  } finally {
    clearTimeout(timeoutId);
  }
  const raw = await response.text();
  let data = {};
  try { data = JSON.parse(raw); } catch {}
  if (!response.ok) {
    const detail = providerDetail(data, raw);
    log?.warn?.({ status: response.status, detail }, "June LemonSlice session request failed");
    lastError = { at: new Date().toISOString(), status: response.status, message: detail };
    throw avatarError(`LemonSlice refused June's visual (HTTP ${response.status}${detail ? `: ${detail}` : ""}).`, 424);
  }
  lastError = null;
  const address = String(data?.websocket_address || "").trim();
  if (!address) throw avatarError("LemonSlice did not return a visual-session connection.", 424);
  const socket = await openTunnel(address);
  const session = { id, owner, socket, expiresAt: Date.now() + SESSION_TTL_MS };
  sessions.set(id, session);
  socket.addEventListener("close", () => {
    if (sessions.get(id) === session) sessions.delete(id);
  }, { once: true });
  socket.addEventListener("error", () => {
    if (sessions.get(id) === session) sessions.delete(id);
  }, { once: true });
  return {
    id,
    livekit_url: settings.livekitUrl,
    viewer_token: viewerToken,
    expires_at: new Date(session.expiresAt).toISOString(),
  };
}

export function sendJuneAvatarAudio({ id, owner, audio }) {
  const encoded = String(audio || "");
  if (!encoded || encoded.length > 700_000 || encoded.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(encoded)) {
    throw avatarError("June received an invalid visual-audio frame.", 400);
  }
  const bytes = Buffer.from(encoded, "base64");
  if (!bytes.length || bytes.length > 500_000) throw avatarError("June's visual-audio frame is too large.", 400);
  const session = ownerSession(id, owner);
  send(session, { command: "audio", audio: encoded, sampleRate: 16_000, encoding: "PCM16" });
  return { ok: true };
}

export function finishJuneAvatarTurn({ id, owner }) {
  const session = ownerSession(id, owner);
  send(session, { command: "audio_end" });
  return { ok: true };
}

export function interruptJuneAvatar({ id, owner }) {
  const session = ownerSession(id, owner);
  send(session, { command: "interrupt" });
  return { ok: true };
}

export function stopJuneAvatar({ id, owner }) {
  const session = sessions.get(String(id || ""));
  if (!session || session.owner !== owner) return { ok: true };
  removeSession(session.id, { terminate: true });
  return { ok: true };
}

export const __test = { createLiveKitToken, getConfig: config };
