# API route structure

`api/routes/reports.routes.js`, `maintenance.routes.js`, `dashboard.routes.js` and `stock.routes.js` used to hold every route in one giant function (about 13,800, 12,700, 4,200 and 2,900 lines). The routes now live in feature files in a folder of the same name:

- `api/routes/reports/*.routes.js`
- `api/routes/maintenance/*.routes.js`
- `api/routes/dashboard/*.routes.js`
- `api/routes/stock/*.routes.js`

## How it fits together

- The main file (for example `reports.routes.js`) still does the one-time setup: file-upload plugin, table columns, background schedulers. It also keeps the shared helper functions.
- At the end of its route function, the main file builds a `ctx` object holding the shared helpers. It then calls each feature file's `register…Routes(app, ctx)`.
- Each feature file imports its own libraries and utilities, and takes the shared helpers it needs from `ctx` at the top (`const { helperA, helperB } = ctx;`).
- The route code itself was moved without changes, so the URLs are exactly the same.

## Adding a route

1. Put it in the feature file it belongs to.
2. If it needs a shared helper that isn't in that file's `const { … } = ctx;` list, add the name there.
3. If the name isn't in the main file's `ctx` object yet, add it there as well.

A helper only one feature file uses can simply live in that file.

## Files

### `api/routes/reports/`

- `fuel-lube.routes.js`: Lube and fuel reports (benchmark, reconciliation, machine history)
- `inspections.routes.js`: Inspection, damage report and legal compliance reports
- `operations-exports.routes.js`: Daily, GM, cost and maintenance-cost Excel exports, rain days
- `period-reports.routes.js`: Daily, weekly, monthly, operations and executive pack reports
- `presentations.routes.js`: Maintenance master, executive and GM presentations and documents
- `settings.routes.js`: Custom report builder, SMTP and PDF settings, report subscriptions
- `stores.routes.js`: Stock monitor, stock movement and part order reports
- `workorder-asset.routes.js`: Work order and asset history PDFs

### `api/routes/maintenance/`

- `insights.routes.js`: Reliability, maintenance insights, governance signals and histogram events
- `inspections.routes.js`: Manager, tyre, undercarriage and artisan inspections
- `mechanic-labor.routes.js`: Mechanic labour entries, settings and timesheets
- `parts-requests.routes.js`: Workshop parts requests and RFQ PDF
- `plans.routes.js`: Maintenance plans, services due, service history and backfill
- `prestart-checks.routes.js`: LDV and machine pre-start checks, checklist hub and damage reports
- `service-templates.routes.js`: Service templates, service planner and service estimates
- `weekly.routes.js`: Borris weekly plan, weekly forum and weekly inspection roster

### `api/routes/dashboard/`

- `asset-kpi.routes.js`: Asset KPI weekly data and exports, LDV pre-start compliance
- `cost-settings.routes.js`: Cost settings, asset rates and part costs
- `fuel.routes.js`: Fuel log, FAMS sync, baselines, comparisons and shift scenarios
- `lube.routes.js`: Lube usage, analytics and mappings
- `overview.routes.js`: Main dashboard, reliability, cost trend and work order nudges

### `api/routes/stock/`

- `cycle-counts.routes.js`: Cycle count sessions and quick counts
- `inventory.routes.js`: Stock on hand, GM stock report, stock monitor, control summary, movement report and FX settings
- `locations.routes.js`: Locations, bins, stock depth, min/max levels and replenishment
- `lube.routes.js`: Lube stock on hand, month stock, minimums, receipts and issues
- `movements.routes.js`: Stock movements, store allocations
- `part-orders.routes.js`: Stores part orders and store QR profile
