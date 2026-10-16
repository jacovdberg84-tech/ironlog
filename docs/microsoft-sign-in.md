# Sign in with Microsoft (Entra ID)

Each IronLog (AML, a pilot company, later customers) can let its staff sign in with
their Microsoft 365 work account. Plan and decisions: the "IronLog — Microsoft 365
(Entra ID) sign-in plan" doc (agreed 16 October 2026).

- Licensed module `microsoft_signin` (Professional and Enterprise packages).
- Admin → **Microsoft sign-in**: the set-up steps for the company's IT, the exact
  redirect address to register, the directory (tenant) ID, client ID and client
  secret (stored encrypted with the instance's own key, never shown again), the date
  the secret runs out, **Switched on**, and **Test sign-in** (opens Microsoft in a new
  tab; nobody is signed in).
- People are **invited**: an admin types the person's work email in
  *Microsoft account* on their user (Admin → Users). At the first sign-in the
  permanent Microsoft account ID is linked; the users table shows *Invited* or *Linked*.
- A Microsoft sign-in gets a normal IronLog session lasting **12 hours**. Passwords,
  PINs (technician portal, stores terminal) and QR pre-starts are unchanged.

- **Microsoft groups → roles (stage 3).** On the same card an admin links a Microsoft
  group (its Object ID) to IronLog roles, and optionally to sites. IT adds a *groups
  claim* to the app (Token configuration; *Groups assigned to the application* when
  people are in many groups). A user marked **Roles and sites come from their Microsoft
  groups** gets the roles of all their linked groups at every sign-in; sites are limited
  only when every one of their groups is limited. Someone taken out of every linked
  group can no longer sign in. Users invited by hand keep the roles the admin gave them.
- **Open to the whole company.** With this on, anyone from the company's directory who
  is in a linked group gets an IronLog user at their first sign-in (username from their
  email, Microsoft-only, roles from groups). People in no linked group are still refused.
  It needs at least one linked group, and removing every group switches it off.
- **Microsoft sign-in only** (Admin → Users): the user has no password or setup code;
  password sign-in, setup and password change are refused. Their PIN still works on the
  technician portal and stores terminal. It needs the person's Microsoft account and
  Microsoft sign-in switched on, and at least one active administrator always keeps a
  password, in case Microsoft sign-in stops working.
- **Secret expiry banner.** Administrators see a banner from 30 days before the client
  secret's date (Admin → Microsoft sign-in) until a new secret and date are saved.
- Sessions already open are not cut when someone leaves a group; the next sign-in
  (within 12 hours) applies the change.

| Route | Who |
|---|---|
| `GET /api/sso/microsoft/config` | public: is the button shown |
| `GET /api/sso/microsoft/start` (`?test=1`) | public: redirect to Microsoft (PKCE, one-time state and nonce) |
| `GET /api/sso/microsoft/callback` | public: Microsoft returns here; the ID token is fully checked |
| `POST /api/sso/microsoft/complete` | public: one-time handoff code (2 minutes) → session token |
| `GET/PUT /api/sso/microsoft/settings` | admin: connection, open to the whole company, linked groups |
| `GET /api/sso/microsoft/status` | admin: the secret-expiry banner |

Code: `api/utils/microsoftSignIn.js`, `api/routes/microsoftSignin.routes.js`; tests in
`api/test/microsoftSignIn.test.js` and `api/test/microsoftSignInGroups.test.js`. Every
sign-in, refusal, link, user created, role change from groups and settings change is in
the audit log (module `auth`, actions `microsoft_signin.*`).
