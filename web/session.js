/**
 * Shared session + request helpers for the signed-in IRONLOG pages
 * (maintenance, breakdown ops, work orders). Load before the page script.
 * Defaults match the main app's legacy local mode (admin / main site).
 */
(function (global) {
  const ROLE_KEY = "ironlog_session_role";
  const ROLES_KEY = "ironlog_session_roles";
  const USER_KEY = "ironlog_session_user";
  const SITE_KEY = "ironlog_session_site";
  const TOKEN_KEY = "ironlog_auth_token";

  function getAuthToken() {
    return String(localStorage.getItem(TOKEN_KEY) || sessionStorage.getItem(TOKEN_KEY) || "").trim();
  }

  function getSessionUser() {
    return String(localStorage.getItem(USER_KEY) || "admin").trim() || "admin";
  }

  function getSessionRole() {
    return String(localStorage.getItem(ROLE_KEY) || "admin").trim().toLowerCase() || "admin";
  }

  function getSessionRoles() {
    try {
      const parsed = JSON.parse(String(localStorage.getItem(ROLES_KEY) || "[]"));
      if (Array.isArray(parsed) && parsed.length) {
        return Array.from(new Set(parsed.map((r) => String(r || "").trim().toLowerCase()).filter(Boolean)));
      }
    } catch {}
    return [getSessionRole()];
  }

  function getSessionSite() {
    return String(localStorage.getItem(SITE_KEY) || "main").trim().toLowerCase() || "main";
  }

  function authHeaders(extra = {}) {
    const h = {
      ...extra,
      "x-user-name": getSessionUser(),
      "x-user-role": getSessionRole(),
      "x-user-roles": getSessionRoles().join(","),
      "x-site-code": getSessionSite(),
    };
    const tok = getAuthToken();
    if (tok) h.Authorization = `Bearer ${tok}`;
    return h;
  }

  function escapeHtml(s) {
    return String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  global.IronlogSession = {
    getAuthToken,
    getSessionUser,
    getSessionRole,
    getSessionRoles,
    getSessionSite,
    authHeaders,
    escapeHtml,
  };
})(window);
