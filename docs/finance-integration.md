# IRONLOG Finance Integration Guide

This guide explains how to use the Finance Integration features delivered in IRONLOG:

- Summarized journal posting runs
- Month-end checklist + period lock/close/reopen
- Budget vs Actual
- Rolling 3-month forecast (hybrid model)
- Canonical KPI definitions
- SSOT (single source of truth) finance report exports

---

## 1) Where to Access in UI

Open IRONLOG web app and use the left sidebar:

- `System` -> `Finance`

If `Finance` is not visible:

1. Refresh app (`Ctrl+F5`).
2. If needed, clear saved tab overrides in browser console:
   - `localStorage.removeItem("ironlog:tabs:override")`
3. Refresh again.

---

## 2) Summarized Journal Runs

### Purpose
Create accounting-ready journal runs summarized from maintenance/procurement operations.

### Steps
In `Finance` -> `Summarized Journal Runs`:

1. Select `Start` and `End` dates.
2. Optional: set `Default Cost Center`.
3. Select categories:
   - `Parts`
   - `Labor`
   - `Downtime`
   - `Fuel`
   - `Lube`
   - Optional procurement categories: `Procurement GRN`, `Procurement AP`
4. Click `Build Summarized Run`.
5. Click `Refresh Runs` and select the run ID.
6. Use:
   - `Export CSV`
   - `Export XLSX`
   - `Mark Exported`
   - `Mark Posted` (optionally provide posted reference)
   - `Reverse` (requires reason)

### Notes
- Period lock prevents posting/reversal in locked/closed periods.
- Run detail shows category totals and line balance.

---

## 3) Month-End Checklist + Period Lock

### Purpose
Control close process and prevent accidental changes after close.

### Steps
In `Finance` -> `Month-End Lock + Checklist`:

1. Enter period as `YYYY-MM` (example: `2026-04`).
2. Click `Load Checklist`.
3. Mark items `Done` / `Skip` / `Reset`.
4. Click `Close Period` (or `Force` if approved exception process applies).
5. To reopen:
   - Enter `Reopen Reason`
   - Click `Reopen`

### Access
- Close: admin/supervisor/plant manager
- Reopen: admin only

---

## 4) Budget vs Actual

### Purpose
Compare planned monthly budget against computed actuals.

### Steps
In `Finance` -> `Budget vs Actual`:

1. Enter `Period` (`YYYY-MM`).
2. Pick dimension:
   - `Cost Center`
   - `Site`
   - `Equipment Type`
   - `Category`
   - `Combined`
3. Click `Load`.

To maintain budgets:

1. Paste budget rows JSON in the editor.
2. Click `Save Budgets`.
3. Click `List Budgets` to confirm.

### Example JSON
```json
[
  {
    "period": "2026-04",
    "site_code": "",
    "cost_center_code": "PLANT-A",
    "equipment_type": "",
    "category": "parts",
    "budget_amount": 12000
  }
]
```

---

## 5) Rolling 3-Month Forecast

### Purpose
Generate monthly forecast using:

- Weighted run-rate baseline from prior 3 months
- Driver-based uplift (PM load, open breakdowns, downtime trend)

### Steps
In `Finance` -> `Rolling 3-Month Forecast`:

1. Enter `Start Period` (`YYYY-MM`).
2. Set `Months Ahead` (default `3`).
3. Click `Rebuild Forecast`.
4. Click `Load Forecast`.

---

## 6) SSOT Report + KPI Definitions

### Purpose
Produce one canonical report set and formal KPI definition registry.

### Steps
In `Finance` -> `SSOT Report + Canonical KPI Definitions`:

1. Enter period (`YYYY-MM`).
2. Click `Load SSOT`.
3. Export:
   - `Export CSV`
   - `Export XLSX`
4. Use `Reload KPI Definitions` to refresh formulas.

### Canonical KPIs
- Availability
- Utilization
- MTBF
- MTTR
- Cost per asset-hour

---

## 7) Smoke Test Script

Script path:

- `api/scripts/finance-smoke.js`

Run:

```powershell
cd c:\IRONLOG
$env:API_BASE="http://localhost:3001"
node api/scripts/finance-smoke.js
```

What it validates end-to-end:

1. Summarized run build
2. Run listing/detail
3. Mark exported
4. Mark posted
5. Checklist load/update
6. Period close/reopen
7. Budget upsert
8. Budget vs actual
9. Forecast rebuild/load
10. SSOT + KPI definitions

---

## 8) API Quick Reference

### Procurement (run-based journals)
- `POST /api/procurement/journals/summarize`
- `GET /api/procurement/journals/runs`
- `GET /api/procurement/journals/runs/:id`
- `GET /api/procurement/journals/runs/:id/lines`
- `POST /api/procurement/journals/runs/:id/mark-exported`
- `POST /api/procurement/journals/runs/:id/mark-posted`
- `POST /api/procurement/journals/runs/:id/reverse`
- `GET /api/procurement/journals/runs/:id/export.csv`
- `GET /api/procurement/journals/runs/:id/export.xlsx`

### Finance
- `GET /api/finance/periods/:period`
- `GET /api/finance/periods/:period/checklist`
- `POST /api/finance/periods/:period/checklist/:code`
- `POST /api/finance/periods/:period/close`
- `POST /api/finance/periods/:period/reopen`
- `POST /api/finance/budgets/upsert`
- `GET /api/finance/budgets`
- `GET /api/finance/budgets-vs-actual`
- `POST /api/finance/forecast/rebuild`
- `GET /api/finance/forecast`
- `GET /api/finance/kpis/definitions`
- `GET /api/finance/reports/ssot`
- `GET /api/finance/reports/ssot/export.csv`
- `GET /api/finance/reports/ssot/export.xlsx`

---

## 9) Troubleshooting

- **Finance tab missing**
  - Clear tab override: `localStorage.removeItem("ironlog:tabs:override")`
  - Refresh app.

- **`period locked` errors**
  - Reopen period (admin only) with reason, then retry operation.

- **No forecast rows**
  - Ensure historical months have source activity (parts/labor/fuel/lube/downtime).

- **Run not balanced**
  - Inspect category lines in run detail and verify source data completeness.

