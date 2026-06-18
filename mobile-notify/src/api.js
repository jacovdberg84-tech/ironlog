import AsyncStorage from "@react-native-async-storage/async-storage";

const STORAGE_ORIGIN = "ironlog_notify_origin";
const STORAGE_TOKEN = "ironlog_notify_token";
const STORAGE_USER = "ironlog_notify_user";
const DEFAULT_ORIGIN = "https://ironlog.ironlogafrica.com";

export { DEFAULT_ORIGIN };

export function normalizeOrigin(raw) {
  let s = String(raw || "").trim();
  if (!s) return "";
  if (!/^https?:\/\//i.test(s)) s = `http://${s}`;
  try {
    return new URL(s).origin;
  } catch {
    return s.replace(/\/api\/?$/i, "").replace(/\/$/, "");
  }
}

export function apiBase(origin) {
  const o = normalizeOrigin(origin || DEFAULT_ORIGIN);
  return `${o}/api`;
}

export async function getStoredOrigin() {
  return normalizeOrigin((await AsyncStorage.getItem(STORAGE_ORIGIN)) || DEFAULT_ORIGIN);
}

export async function setStoredOrigin(origin) {
  const o = normalizeOrigin(origin);
  if (o) await AsyncStorage.setItem(STORAGE_ORIGIN, o);
  return o;
}

export async function getStoredToken() {
  return String((await AsyncStorage.getItem(STORAGE_TOKEN)) || "").trim();
}

export async function setStoredSession({ token, user }) {
  if (token) await AsyncStorage.setItem(STORAGE_TOKEN, token);
  else await AsyncStorage.removeItem(STORAGE_TOKEN);
  if (user) await AsyncStorage.setItem(STORAGE_USER, JSON.stringify(user));
  else await AsyncStorage.removeItem(STORAGE_USER);
}

export async function getStoredUser() {
  try {
    const raw = await AsyncStorage.getItem(STORAGE_USER);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

export async function clearStoredSession() {
  await AsyncStorage.multiRemove([STORAGE_TOKEN, STORAGE_USER]);
}

export async function fetchJson(origin, path, opts = {}) {
  const base = apiBase(origin);
  const headers = new Headers(opts.headers || {});
  if (!headers.has("Content-Type") && opts.body) {
    headers.set("Content-Type", "application/json");
  }
  const res = await fetch(`${base}${path}`, { ...opts, headers });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    const err = new Error(data?.error || `HTTP ${res.status}`);
    err.status = res.status;
    err.data = data;
    throw err;
  }
  return data;
}

export async function authFetch(origin, path, token, opts = {}) {
  const headers = new Headers(opts.headers || {});
  if (token) headers.set("Authorization", `Bearer ${token}`);
  return fetchJson(origin, path, { ...opts, headers });
}

export async function fetchPinRoster(origin) {
  return authFetch(origin, "/auth/pin/roster", "");
}

export async function loginPin(origin, username, pin) {
  return fetchJson(origin, "/auth/pin/login", {
    method: "POST",
    body: JSON.stringify({ username, pin }),
  });
}

export async function registerPushToken(origin, token, deviceToken, deviceLabel) {
  return authFetch(origin, "/notifications/register", token, {
    method: "POST",
    body: JSON.stringify({
      token: deviceToken,
      platform: "android",
      device_label: deviceLabel || undefined,
    }),
  });
}

export async function fetchInbox(origin, token, limit = 20) {
  return authFetch(origin, `/notifications/inbox?limit=${limit}`, token);
}

export function workOrderUrl(origin, woId) {
  const o = normalizeOrigin(origin);
  return `${o}/web/workorder-qr.html?wo_id=${encodeURIComponent(String(woId || ""))}`;
}

export function technicianTerminalUrl(origin) {
  const o = normalizeOrigin(origin);
  return `${o}/web/technician-terminal.html`;
}
