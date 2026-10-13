# IronLog — notes for Claude

IronLog is the maintenance platform for AML's mining fleet in Mozambique (Palma, Cabo Delgado). It runs as one Node server: a Fastify + SQLite API in `api/` that also serves the web app in `web/`. Pushing to `main` deploys to the live server (`.github/workflows/deploy.yml`), so every change goes through a pull request that a person reviews.

## People and roles
- `admin`: full access. `workshop_admin` = plant admin. `plant_manager` = foreman. `storeman` = stores. `artisan` = technician. `operator` = machine operator.
- Technicians and operators speak Portuguese. Their screens (technician portal, stores terminal, checklists) need EN/PT text.
- Borris is the engineering assistant. June is the admin's private voice PA, visible to admins only.

## Code layout
- `api/routes/*.routes.js` are Fastify plugins registered in `api/server.js`. Large areas are split into folders (`api/routes/reports/`, `maintenance/`, `dashboard/` and others). See `docs/api-route-structure.md`.
- `api/utils/` holds shared logic. Tables are created in code with `CREATE TABLE IF NOT EXISTS` (an `ensure…Tables()` helper), never by hand-editing the database.
- `web/index.html` loads `web/app/*.js` as plain scripts that share one global scope. See `docs/web-app-structure.md`. Every new `web/app` file must be added to `index.html`, and `api/test/webAppBundle.test.js` checks this.
- After changing a web file, bump its `?v=` number in the HTML so browsers load the new version. This includes `styles.css`.
- Colours go through CSS variables so dark mode works. Use the tokens at the top of `web/styles.css` (`--card`, `--ink`, `--line`, `--bg-*`, `--tx-*`, `--bd-*`); do not add raw light hex colours.

## Rules
- Run the tests before finishing: `cd api && npm test` (Node's built-in test runner, `api/test/*.test.js`). Add or update tests for what you change. Never skip or delete a test to get green.
- Keep the change focused on the request. Match the surrounding code style and comment density.
- Never commit secrets, `.env` files, databases (`*.db`) or backups. Never change `.github/workflows/deploy.yml` unless the request is about deployment.
- Money is in US dollars. Dates are `YYYY-MM-DD`; the site is on Africa/Maputo time.
- Anything that changes records (stock, work orders, costs) needs a role check on the server, not only a hidden button.
- Write user-facing text in plain, friendly language; the users are workshop people, not developers.
- When you finish, summarise in plain language what changed, how you tested it, and anything a person should check on the live site.
