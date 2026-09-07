# Borris weekly maintenance planner

Adds **Plan next week** beside Ask Borris in the Borris AI panel. From Maintenance, select Borris AI, then Plan next week.

The authenticated read-only `GET /api/maintenance/weekly-plan` endpoint returns a draft with the next Monday–Sunday window (Africa/Johannesburg), overdue services, services due before that window ends, combined parts requirements and shortages, known service-kit/labour estimates, and records requiring review. Existing rotating service rules determine the next service per asset. No work orders or reservations are created.

Forecasts use recorded running hours divided by 14 calendar days. Days above 24 hours, missing meter readings, and kilometre-based future services require review. Overdue kilometre services can still appear. Only the next service per asset is included; repeated service cycles, mechanic capacity and downtime windows require planner review.

Cost figures reuse existing manual/store prices and historical averages. Missing estimates are flagged and excluded from the known subtotal; currency is not assumed. Historical costs are estimates, not quotes. Stock is combined across included services but is not reserved or adjusted for other pending demand.

Changed files:
- api/routes/maintenance.routes.js
- api/utils/weeklyMaintenancePlan.js
- api/test/weeklyMaintenancePlan.test.js
- web/app.js
- web/index.html

Validation: `node --test api/test/*.test.js`, plus API syntax and git whitespace checks. Tests include service rotation, week/year boundaries, aggregate stock shortages, missing costs, invalid readings, classic browser script parsing, UI success/failure handling and endpoint wiring. No production database or live LLM was used in tests.

Deployment: deploy the API and web changes together using Ironlog's normal release process, restart the API, then reload the browser. No database migration or model download is needed. The current live site has not been changed by this local implementation.
