# API keys for other systems

Other systems (an ERP such as Sage, Power BI, a client portal) read IronLog with an
API key. Readable documentation for the people connecting is at `/web/api-docs.html`
(and `/api/api-keys/openapi.json`).

- Admin → **API keys for other systems**: an administrator names the key, ticks what
  it may read (scopes), optionally limits it to sites and sets an end date. The key
  (`ilk_…`) is shown once; only its SHA-256 hash is stored. Keys can be cancelled at
  any time; the list shows when each was last used and how often.
- Keys are **read-only**: each scope is a short list of exact `GET` addresses
  (`api/utils/apiKeys.js`, `API_SCOPES`). A key never reaches settings, users, admin
  screens, or anything that changes records, whatever headers it sends.
- A key acts with a supervisor's read access on those addresses only, for its sites
  (the same site guard as people with Allowed Locations).
- 300 requests a minute per key (429 beyond that).
- Licensed as the **api_access** module, in the Enterprise package
  (`utils/licenceIssuer.js` catalogue step `api_access`).
- Audit log: `api_keys` module, `api_key.create` and `api_key.cancel`.

| Route | Who |
|---|---|
| `GET/POST /api/api-keys` | admin: list (never the secret) / create |
| `POST /api/api-keys/:id/cancel` | admin |
| `GET /api/api-keys/catalogue` | anyone signed in: what each scope reads |
| `GET /api/api-keys/openapi.json` | anyone signed in |

Code: `api/utils/apiKeys.js`, `api/routes/apiKeys.routes.js`, `api/auth/hook.js`
(`apiKeyAccess`), `api/routes/licence.routes.js`; tests in `api/test/apiKeys.test.js`.
