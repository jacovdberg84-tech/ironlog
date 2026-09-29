# Main web app structure

The main IronLog screen (`web/index.html`) used to load one 22,000-line `web/app.js`. That code now lives in `web/app/`, split by feature. Nothing was rewritten: the files, joined in order, are the old `app.js` apart from a two-line header at the top of each.

## Rules

- `index.html` loads the files in the order below, as plain `<script>` tags. They share one global scope, just like the old single file, so any file can call a function from any other file **once the page has loaded**.
- Code that runs straight away while a file loads (not inside a function) may only use things from the same file or an earlier one. Nearly all code is inside functions, so this rarely matters.
- Add new code to the file for that feature. To add a new file, add a `<script src="./app/…">` tag in `index.html` in the right place. The test `api/test/webAppBundle.test.js` fails if a file on disk is not loaded, or is loaded twice.
- After changing any of these files, bump the `?v=` number on its script tag so browsers fetch the new version.
- `web/my-work.js` (the My Work home screen) loads after these files.

## Files, in load order

| File | What is in it | Lines |
|---|---|---|
| `core.js` | Config, session, auth, fetch helpers, navigation and role visibility, login | 1870 |
| `admin.js` | Users & Access, safety/telematics admin, master data, SMTP, push, PDF settings, backups | 1406 |
| `inspections.js` | Checklist hub, LDV/machine checklists, vehicle check photos | 1370 |
| `ui-helpers.js` | Toasts, status, formatting helpers, thresholds, offline queues | 632 |
| `telematics.js` | Telematics, Cartrack fleet/map/speeding, GPS links, Unitech | 1275 |
| `dashboard.js` | Dashboard KPIs | 578 |
| `borris.js` | Borris (Ironmind) insights, reports and Q&A | 507 |
| `fuel-lube.js` | Lube usage, fuel log/benchmarks, cost settings, shift scenarios | 1560 |
| `stock.js` | Stock monitor, stock reports, stores part orders, parts tracking | 1241 |
| `compliance.js` | Audit trail, approvals, legal documents | 482 |
| `tabs.js` | Tab switching, in-app help, section toggles | 247 |
| `uploads.js` | CSV uploads, FAMS fuel import and templates | 517 |
| `reports.js` | Report downloads (daily, weekly, GM, cost, lube, stock, operations) | 503 |
| `breakdowns.js` | Breakdown ops, operational slips, short breakdowns | 667 |
| `inventory.js` | Part issues, store allocations, manual stock, bins, cycle counts, lube stock | 1127 |
| `procurement.js` | Requisitions, approval chains, purchase orders, journals | 708 |
| `site-ops.js` | Site operations, closing, dispatch, data quality | 866 |
| `daily-input.js` | Daily input grid, prestart section, asset QR sheets, shift self-check | 1804 |
| `assets.js` | Assets register, contractors, plant hire, asset history | 788 |
| `init.js` | App start-up and event wiring | 1901 |
| `documents.js` | Translations, AI documents, dark mode | 432 |
| `tasks.js` | Task workspace | 704 |
| `finance.js` | Finance | 738 |
| `enterprise.js` | Enterprise, integrations, governance, evidence packs, executive | 389 |
| `workshop-docs.js` | Workshop library documents and uploads, Borris OEM | 88 |
