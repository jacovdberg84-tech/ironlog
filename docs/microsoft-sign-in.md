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

| Route | Who |
|---|---|
| `GET /api/sso/microsoft/config` | public: is the button shown |
| `GET /api/sso/microsoft/start` (`?test=1`) | public: redirect to Microsoft (PKCE, one-time state and nonce) |
| `GET /api/sso/microsoft/callback` | public: Microsoft returns here; the ID token is fully checked |
| `POST /api/sso/microsoft/complete` | public: one-time handoff code (2 minutes) → session token |
| `GET/PUT /api/sso/microsoft/settings` | admin |

Code: `api/utils/microsoftSignIn.js`, `api/routes/microsoftSignin.routes.js`; tests in
`api/test/microsoftSignIn.test.js`. Every sign-in, refusal, link and settings change is
in the audit log (module `auth`, actions `microsoft_signin.*`).

Still to come (stage 3): Microsoft groups to IronLog roles, "open to the whole company",
"Microsoft sign-in only" users, and a banner warning before the secret runs out.
