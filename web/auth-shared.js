/**
 * Shared sign-in + API helpers for standalone IRONLOG pages
 * (technician terminal, work order QR, work orders board).
 */
(function (global) {
  const API = global.location?.origin || "http://localhost:3001";
  const ROLE_KEY = "ironlog_session_role";
  const ROLES_KEY = "ironlog_session_roles";
  const USER_KEY = "ironlog_session_user";
  const SITE_KEY = "ironlog_session_site";
  const TOKEN_KEY = "ironlog_auth_token";
  const TABS_OVERRIDE_KEY = "ironlog_allowed_tabs";
  const DEFAULT_SITE = "main";

  function getAuthToken() {
    return String(localStorage.getItem(TOKEN_KEY) || sessionStorage.getItem(TOKEN_KEY) || "").trim();
  }

  function setAuthToken(t, remember = true) {
    if (t) {
      if (remember) {
        localStorage.setItem(TOKEN_KEY, t);
        sessionStorage.removeItem(TOKEN_KEY);
      } else {
        sessionStorage.setItem(TOKEN_KEY, t);
        localStorage.removeItem(TOKEN_KEY);
      }
    } else {
      localStorage.removeItem(TOKEN_KEY);
      sessionStorage.removeItem(TOKEN_KEY);
    }
  }

  function normalizeRoles(input, fallbackRole = "operator") {
    const base = Array.isArray(input) ? input : [];
    const out = Array.from(
      new Set(base.map((r) => String(r || "").trim().toLowerCase()).filter(Boolean))
    );
    if (out.length) return out;
    const fb = String(fallbackRole || "operator").trim().toLowerCase() || "operator";
    return [fb];
  }

  function getSessionRole() {
    return String(localStorage.getItem(ROLE_KEY) || "operator").trim().toLowerCase() || "operator";
  }

  function getSessionRoles() {
    const primary = getSessionRole();
    try {
      const raw = localStorage.getItem(ROLES_KEY);
      if (!raw) return [primary];
      return normalizeRoles(JSON.parse(raw), primary);
    } catch {
      return [primary];
    }
  }

  function getSessionUser() {
    return String(localStorage.getItem(USER_KEY) || "").trim();
  }

  function getSessionSite() {
    return String(localStorage.getItem(SITE_KEY) || DEFAULT_SITE).trim().toLowerCase() || DEFAULT_SITE;
  }

  function setSessionContext(user, role, site, roles = null) {
    const rolePrimary = String(role || "operator").trim().toLowerCase() || "operator";
    const roleList = normalizeRoles(Array.isArray(roles) ? roles : [rolePrimary], rolePrimary);
    localStorage.setItem(USER_KEY, String(user || "").trim());
    localStorage.setItem(ROLE_KEY, rolePrimary);
    localStorage.setItem(ROLES_KEY, JSON.stringify(roleList));
    localStorage.setItem(SITE_KEY, String(site || DEFAULT_SITE).trim().toLowerCase() || DEFAULT_SITE);
  }

  function applyUser(user) {
    if (!user) return;
    const u = String(user.username || "").trim();
    const r = String(user.role || "operator").trim().toLowerCase() || "operator";
    const roles = normalizeRoles(user.roles, r);
    const allowedLoc = Array.isArray(user.allowed_locations)
      ? user.allowed_locations.map((x) => String(x || "").trim().toLowerCase()).filter(Boolean)
      : [];
    const currentSite = getSessionSite();
    const nextSite = allowedLoc.length ? (allowedLoc.includes(currentSite) ? currentSite : allowedLoc[0]) : currentSite;
    setSessionContext(u, r, nextSite, roles);
    if (Array.isArray(user.allowed_tabs) && user.allowed_tabs.length) {
      localStorage.setItem(TABS_OVERRIDE_KEY, JSON.stringify(user.allowed_tabs));
    } else {
      localStorage.removeItem(TABS_OVERRIDE_KEY);
    }
  }

  function clearSession() {
    setAuthToken("");
    localStorage.removeItem(TABS_OVERRIDE_KEY);
    localStorage.removeItem(ROLES_KEY);
    localStorage.removeItem(ROLE_KEY);
    localStorage.removeItem(USER_KEY);
  }

  function authHeaders(extra = {}) {
    const tok = getAuthToken();
    if (!tok) return { ...extra };
    const roles = getSessionRoles();
    const h = {
      ...extra,
      "x-user-name": getSessionUser(),
      "x-user-role": getSessionRole(),
      "x-user-roles": roles.join(","),
      "x-site-code": getSessionSite(),
      Authorization: `Bearer ${tok}`,
    };
    return h;
  }

  async function fetchJson(url, opts = {}) {
    const nextOpts = { ...opts };
    const headers = new Headers(nextOpts.headers || {});
    const tok = getAuthToken();
    if (tok) {
      const roles = getSessionRoles();
      headers.set("x-user-name", getSessionUser());
      headers.set("x-user-role", getSessionRole());
      headers.set("x-user-roles", roles.join(","));
      headers.set("x-site-code", getSessionSite());
      headers.set("Authorization", `Bearer ${tok}`);
    }
    if (typeof nextOpts.body === "string" && !headers.has("Content-Type")) {
      headers.set("Content-Type", "application/json");
    }
    nextOpts.headers = headers;

    const res = await fetch(url, nextOpts);
    const text = await res.text();
    let data = null;
    try {
      data = text ? JSON.parse(text) : null;
    } catch {
      data = { raw: text };
    }
    if (!res.ok) {
      const err = new Error(data?.error || data?.message || text || `Request failed (${res.status})`);
      err.status = res.status;
      throw err;
    }
    return data || {};
  }

  async function login(username, password, remember = true) {
    const data = await fetchJson(`${API}/api/auth/login`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username, password }),
    });
    if (data.token) setAuthToken(data.token, remember);
    if (data.user) applyUser(data.user);
    return data;
  }

  async function loginPin(username, pin, remember = false) {
    const data = await fetchJson(`${API}/api/auth/pin/login`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username, pin: String(pin || "").replace(/\D/g, "") }),
    });
    if (data.token) setAuthToken(data.token, remember);
    if (data.user) applyUser(data.user);
    return data;
  }

  async function fetchPinRoster() {
    const data = await fetchJson(`${API}/api/auth/pin/roster`);
    return Array.isArray(data?.technicians) ? data.technicians : [];
  }

  async function trySession() {
    if (!getAuthToken()) return null;
    try {
      const data = await fetchJson(`${API}/api/auth/me`);
      if (data?.user?.id != null) {
        applyUser(data.user);
        return data.user;
      }
      clearSession();
      return null;
    } catch (e) {
      if (e.status === 401) clearSession();
      return null;
    }
  }

  function hasRole(roles) {
    const mine = getSessionRoles();
    const want = Array.isArray(roles) ? roles : [roles];
    return want.some((r) => mine.includes(String(r).toLowerCase()));
  }

  function escapeHtml(s) {
    return String(s || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  global.IronlogAuth = {
    API,
    getAuthToken,
    setAuthToken,
    getSessionRole,
    getSessionRoles,
    getSessionUser,
    getSessionSite,
    setSessionContext,
    applyUser,
    clearSession,
    authHeaders,
    fetchJson,
    login,
    loginPin,
    fetchPinRoster,
    trySession,
    hasRole,
    escapeHtml,
  };
})(window);
