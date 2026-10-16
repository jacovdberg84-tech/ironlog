// IRONLOG/api/routes/microsoftSignin.routes.js — "Sign in with Microsoft" (utils/microsoftSignIn.js).
//   GET  /api/sso/microsoft/config     public: is the button shown
//   GET  /api/sso/microsoft/start      public: off to Microsoft (?test=1 from the admin card)
//   GET  /api/sso/microsoft/callback   public: Microsoft sends people back here
//   POST /api/sso/microsoft/complete   public: one-time handoff code → IronLog session token
//   GET/PUT /api/sso/microsoft/settings   administrators (connection, "open to the whole company", linked groups)
//   GET  /api/sso/microsoft/status     administrators: the banner warning before the client secret runs out
// Licensed as the "microsoft_signin" module (utils/licensing.js).
import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";
import { requireAnyRole, getUser } from "../utils/request.js";
import { checkAllowedLocations, knownSites } from "../utils/sites.js";
import { VALID_ROLES } from "./auth.routes.js";
import {
  claimEmail, createSessionHandoff, ensureMicrosoftSignInTables, exchangeCode, getMsSettings, groupAccess, listGroupRoles, matchUser,
  redirectUriFor, saveGroupRoles, saveMsSettings, secretDaysLeft, secretWarning, signingKeys, startSignIn, takeHandoff, takeRequest, verifyIdToken,
} from "../utils/microsoftSignIn.js";

/** IronLog's own web address as people reach it (Cloudflare passes the original host and https). */
export function publicBase(req) {
  const fixed = String(process.env.IRONLOG_PUBLIC_BASE_URL || "").trim();
  if (fixed) return fixed.replace(/\/+$/, "");
  const proto = String(req.headers["x-forwarded-proto"] || req.protocol || "http").split(",")[0].trim();
  const host = String(req.headers["x-forwarded-host"] || req.headers.host || "").split(",")[0].trim();
  return `${proto}://${host}`;
}

const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);

function testPage(ok, lines) {
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Microsoft sign-in test</title>
<style>body{font:16px system-ui,sans-serif;max-width:560px;margin:40px auto;padding:0 16px;color:#111;background:#fff}h1{font-size:20px}.ok{color:#166534}.bad{color:#991b1b}li{margin:6px 0}
@media (prefers-color-scheme:dark){body{background:#0f172a;color:#e5e7eb}.ok{color:#86efac}.bad{color:#fca5a5}}</style></head>
<body><h1 class="${ok ? "ok" : "bad"}">${ok ? "Microsoft sign-in works" : "Microsoft sign-in test failed"}</h1><ul>${lines.map((l) => `<li>${esc(l)}</li>`).join("")}</ul>
<p>This was only a test: nobody was signed in. You can close this tab.</p></body></html>`;
}

export default async function microsoftSigninRoutes(app) {
  ensureMicrosoftSignInTables(db);
  ensureAuditTable(db);

  const audit = (req, action, payload) => {
    try {
      writeAudit(db, req, { module: "auth", action, entity_type: "microsoft_signin", entity_id: payload?.username || null, payload });
    } catch {}
  };

  app.get("/config", async () => ({ ok: true, enabled: getMsSettings(db).enabled }));

  app.get("/start", async (req, reply) => {
    const test = String(req.query?.test || "") === "1";
    try {
      return reply.redirect(startSignIn(db, { redirectUri: redirectUriFor(publicBase(req)), test }));
    } catch (err) {
      if (test) return reply.type("text/html").send(testPage(false, [err.message]));
      return reply.redirect(`/web/index.html?ms_error=${encodeURIComponent(err.code || "failed")}`);
    }
  });

  app.get("/callback", async (req, reply) => {
    const q = req.query || {};
    const pending = takeRequest(db, q.state);
    const test = Boolean(pending?.test);
    const failTo = (code, message, extra = {}) => {
      audit(req, "microsoft_signin.refused", { reason: code, message, ...extra });
      if (test) return reply.type("text/html").send(testPage(false, [message]));
      return reply.redirect(`/web/index.html?ms_error=${encodeURIComponent(code)}`);
    };
    if (!pending) return failTo("expired", "The sign-in link was not recognised or took longer than 10 minutes. Please try again.");
    if (q.error) return failTo("microsoft_error", `Microsoft stopped the sign-in: ${String(q.error_description || q.error).split("\n")[0].slice(0, 200)}`);
    const s = getMsSettings(db);
    let claims;
    try {
      const idToken = await exchangeCode(db, { code: q.code, verifier: pending.verifier, redirectUri: pending.redirect_uri });
      claims = await verifyIdToken(idToken, {
        tenantId: s.tenant_id, clientId: s.client_id, nonce: pending.nonce,
        keys: (refresh) => signingKeys(s.tenant_id, { refresh }),
      });
    } catch (err) {
      return failTo(err.code || "failed", err.message);
    }
    const email = claimEmail(claims);
    if (test) {
      const known = db.prepare(`SELECT username, ms_oid, ms_roles_from_groups FROM users WHERE (ms_tid = ? AND ms_oid = ?) OR LOWER(TRIM(ms_email)) = ?`).get(claims.tid, claims.oid, email);
      const access = groupAccess(db, claims);
      const linkedGroups = listGroupRoles(db).length;
      const groupLine = !linkedGroups
        ? null
        : access.overage
          ? "Microsoft did not list this account's groups (too many). Ask IT to set the app's groups claim to 'Groups assigned to the application'."
          : !Array.isArray(claims.groups)
            ? "Microsoft sent no groups. Ask IT to add a groups claim (Security groups) under the app's Token configuration."
            : access.groups.length
              ? `In linked groups: ${access.groups.map((g) => g.label || g.group_id).join(", ")}. Roles: ${access.roles.join(", ")}. Sites: ${access.sites ? access.sites.join(", ") : "all"}.`
              : "This account is in none of the linked groups.";
      const s2 = getMsSettings(db);
      const who = known
        ? `This account is IronLog user ${known.username} and can sign in.`
        : s2.open_to_company && access.roles.length
          ? "No IronLog user yet: one will be made at the first sign-in, from the linked groups."
          : "No IronLog user has this work email yet: add it to a user under Admin → Users to invite them.";
      audit(req, "microsoft_signin.test", { email, matched: known?.username || null, groups: access.groups.map((g) => g.group_id) });
      return reply.type("text/html").send(testPage(true, [
        `Microsoft confirmed ${claims.name || email} (${email || "no email"}) from your company's directory.`,
        who,
        groupLine,
        s.enabled ? "Microsoft sign-in is switched on." : "Microsoft sign-in is still switched off; tick 'Switched on' and save when you are ready.",
      ].filter(Boolean)));
    }
    let user;
    try {
      user = matchUser(db, claims, { validRoles: VALID_ROLES });
    } catch (err) {
      return failTo(err.code || "failed", err.message, { email });
    }
    const code = createSessionHandoff(db, user.id);
    if (user.created) audit(req, "microsoft_signin.user_created", { username: user.username, email, roles: JSON.parse(user.roles_json || "[]"), sites: user.allowed_locations || null });
    if (user.roles_changed) audit(req, "microsoft_signin.roles_from_groups", { username: user.username, email, from: user.roles_changed.from, to: user.roles_changed.to });
    audit(req, user.created || user.linked_now ? "microsoft_signin.linked" : "microsoft_signin.success", { username: user.username, email });
    return reply.redirect(`/web/index.html?ms_handoff=${encodeURIComponent(code)}`);
  });

  app.post("/complete", async (req, reply) => {
    const token = takeHandoff(db, req.body?.handoff);
    if (!token) return reply.code(400).send({ ok: false, error: "expired", message: "The sign-in took too long. Please try again." });
    return { ok: true, token };
  });

  app.get("/settings", async (req, reply) => {
    if (!requireAnyRole(req, reply, ["admin"])) return;
    const settings = getMsSettings(db);
    const linked = db.prepare(`SELECT COUNT(*) AS n FROM users WHERE ms_oid IS NOT NULL`).get().n;
    const invited = db.prepare(`SELECT COUNT(*) AS n FROM users WHERE ms_email IS NOT NULL AND ms_email <> ''`).get().n;
    const msOnly = db.prepare(`SELECT COUNT(*) AS n FROM users WHERE COALESCE(ms_only, 0) = 1`).get().n;
    const fromGroups = db.prepare(`SELECT COUNT(*) AS n FROM users WHERE COALESCE(ms_roles_from_groups, 0) = 1`).get().n;
    return {
      ok: true, settings, redirect_uri: redirectUriFor(publicBase(req)), secret_days_left: secretDaysLeft(settings),
      invited, linked, ms_only: msOnly, roles_from_groups: fromGroups, groups: listGroupRoles(db), roles: VALID_ROLES.filter((r) => r !== "stores"),
    };
  });

  app.get("/status", async (req, reply) => {
    if (!requireAnyRole(req, reply, ["admin"])) return;
    return { ok: true, warning: secretWarning(db) };
  });

  app.put("/settings", async (req, reply) => {
    if (!requireAnyRole(req, reply, ["admin"])) return;
    const before = getMsSettings(db);
    const body = req.body || {};
    let groupsChanged = null;
    if (body.groups !== undefined) {
      const beforeGroups = listGroupRoles(db);
      const g = saveGroupRoles(db, body.groups, { validRoles: VALID_ROLES, checkSites: (list) => {
        const r = checkAllowedLocations(db, list);
        if (r.ok) return r;
        const known = knownSites(db);
        return { ok: false, error: `not a site: ${list.filter((x) => !known.includes(x.toLowerCase())).join(", ")}. Sites: ${known.join(", ")}. Leave empty for all sites.` };
      }, by: getUser(req) });
      if (!g.ok) return reply.code(400).send({ ok: false, error: g.errors.join(" "), errors: g.errors });
      if (JSON.stringify(beforeGroups) !== JSON.stringify(g.groups)) groupsChanged = { from: beforeGroups, to: g.groups };
    }
    const r = saveMsSettings(db, body, getUser(req));
    if (!r.ok) return reply.code(400).send({ ok: false, error: r.errors.join(" "), errors: r.errors });
    if (groupsChanged) audit(req, "microsoft_signin.groups", groupsChanged);
    audit(req, "microsoft_signin.settings", {
      enabled: { from: before.enabled, to: r.settings.enabled },
      open_to_company: { from: before.open_to_company, to: r.settings.open_to_company },
      tenant_id: r.settings.tenant_id, client_id: r.settings.client_id,
      secret_changed: Boolean(String(req.body?.client_secret || "").trim()),
    });
    return { ok: true, settings: r.settings, secret_days_left: secretDaysLeft(r.settings), groups: listGroupRoles(db) };
  });
}
