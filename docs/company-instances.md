# Company instances (pilot customers on the same server)

AML's IronLog runs from `/data/ironlog` as PM2 `ironlog-api` on port 3001. Another
company gets its own **instance**: the same code, but its own data, started by
`deploy/ironlog-instance.sh`.

| | AML | Instance `<name>` |
|---|---|---|
| Folder | `/data/ironlog/api` | `/data/ironlog-instances/<name>/` |
| Settings | `api/.env` | `instance.env` (chmod 600), read alone via `IRONLOG_ENV_FILE` |
| Database, uploads, backups | `api/db`, `api/uploads` | `db/ironlog.db`, `data/`, `db/backups/` |
| Secrets folder | `api/secrets` (licence signing key) | `secrets/` (empty; licence issuing is off) |
| PM2 name, port | `ironlog-api`, 0.0.0.0:3001 | `ironlog-<name>`, 127.0.0.1:31xx |
| Web address | ironlog.ironlogafrica.com | Rule in cloudflared's config file, e.g. pilot.ironlogafrica.com → `http://localhost:3101` |

What keeps them apart:
- `api/loadEnv.js`: with `IRONLOG_ENV_FILE` set, inherited IronLog settings (AI keys,
  GitHub token, Outlook, SMTP, telematics, `DB_PATH`, `PORT`…) are dropped and only
  the instance file is read. The script also starts PM2 with a clean environment.
- Each instance has its own `IRONLOG_AUTH_SECRET` and `JUNE_CONNECTOR_ENCRYPTION_SECRET`,
  generated once and never replaced; sign-in is always required.
- `IRONLOG_LICENCE_ISSUER=0`: an instance cannot issue licences. A new instance starts
  a 30-day trial and gets its licence key from the AML server's Licence issuer.
- No AI key: Borris gives its built-in answers. To switch AI on for one company, add
  `OPENAI_API_KEY` (or `OLLAMA_HOST` and `LLM_MODEL` for the server's local Ollama) to
  that instance's `instance.env` and run the workflow again to restart it.

## Creating one
1. GitHub → Actions → **Create or update a company instance** → Run workflow with the
   name, port, time zone, web address and first administrator; type `CREATE-INSTANCE`.
2. GitHub → Actions → **Publish a company instance (Cloudflare route)** with the web
   address and port; type `ADD-ROUTE`. The `ironlog` tunnel is locally managed, so its
   routes are in cloudflared's config file on the server, not the Cloudflare dashboard
   (`deploy/tunnel-route.sh` adds the rule above the catch-all, checks the file, and
   puts the saved copy back if anything fails).
3. The log shows a one-time setup code (valid 48 hours). The administrator signs in
   with the username, the code in "Setup code (first-time only)" and a new password.
   Running the workflow again before then gives a fresh code.
4. That administrator fills in Admin → Company profile, sets up off-site backups
   (use a separate Storage Box sub-account) and enters the licence key when issued.

Every deploy restarts all instances on the new code (`restart-all` in `deploy.yml`).
`bash deploy/ironlog-instance.sh list` on the server shows them.
