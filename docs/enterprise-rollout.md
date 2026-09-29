# IRONLOG Enterprise Rollout Guide

This guide describes the final big-out rollout:

- Multi-entity support (company -> site, currency, tax)
- Integration Hub (ERP + Payroll) with retry queues, dead-letter, monitoring
- Governance v2 (Segregation of Duties)
- Compliance-ready Evidence Packs (audit / procurement / HSE)
- Executive Command Center + board monthly pack

---

## 1) Where to Access in UI

Left sidebar sections:

- `System` -> `Enterprise`
- `System` -> `Exec Command`

If not visible:

1. Hard refresh (`Ctrl+F5`).
2. In browser console run:
   `localStorage.removeItem("ironlog:tabs:override")`
3. Refresh again.

---

## 2) Multi-Entity (Company -> Site)

### Company Profiles
1. Open `Enterprise` tab -> `Company Profiles`.
2. Fill in: Code, Name, Base Currency, Reporting Currency, Tax Region.
3. Click `Save Company`.

### Site Profiles
1. `Enterprise` tab -> `Site Profiles`.
2. Fill in: Site Code, Company, Name, Local Currency, Region.
3. Click `Save Site`.

### Entity Tree
- Click `Entity Tree` to see `company -> sites` hierarchy view.

### Currency Rates
- Paste JSON rows example:
```json
[{"from_currency":"EUR","to_currency":"USD","rate":1.07,"effective_date":"2026-04-01"}]
```
- Click `Save Rates`.

### Tax Profiles
- Fill in: Code, Label, Rate %, Region.
- Click `Save Tax`.

---

## 3) Integration Hub (ERP + Payroll)

### Create Connection
1. `Enterprise` tab -> `Integration Hub`.
2. Fill in: Connection Code, Connector, Label.
3. Click `Save Connection`.

### Enqueue + Run Job
1. Fill Connection Code, Connector, Payload JSON (example below), Idempotency Key.
2. Click `Enqueue`.
3. Select job ID, click `Run Now`.

#### Example payloads

- ERP Journal Export:
```json
{"run_id": 12}
```
- Payroll Labor Sync:
```json
{"period": "2026-04"}
```

### Retry + Monitoring
- `Retry` manually requeues `retry_wait` or `failed_permanent` jobs.
- `Cancel` cancels queued/retry jobs.
- `Monitor` shows queue depth, failures, top errors, oldest queued.
- `Dead-letter` lists items awaiting triage; acknowledge once resolved.

---

## 4) Governance v2 (SoD)

### Policies
- `Policies` shows active policies and modes (block/warn).
- Defaults seeded automatically:
  - sod_po_create_approve
  - sod_po_approve_receive
  - sod_invoice_capture_post
  - sod_period_close_reopen
  - sod_run_post_reverse

### Evaluate an Action
1. Enter `Action` (example: `finance.journals.reverse`).
2. Enter `Username`.
3. Click `Evaluate`.

### Violations
- `Violations` lists entries with `Open` / `Acknowledged` filter.
- Admin/supervisor/executive can acknowledge with notes.

---

## 5) Compliance Evidence Packs

### Build
1. `Enterprise` tab -> `Compliance Evidence Packs`.
2. Select `Pack Type` (audit / procurement / hse).
3. Set `Period` (YYYY-MM).
4. Click `Build Pack`.

### Integrity
- Each pack includes a sha256 `integrity_hash`.
- Open pack detail to verify `integrity_ok` flag.

---

## 6) Executive Command Center

### Command Center Snapshot
1. `Exec Command` tab.
2. Enter `Period`.
3. Click `Load Command Center`.

### Board Pack
1. Click `Generate Board Pack`.
2. Review narrative highlights + financial + operational + integrations + governance.
3. Click `Export XLSX` for the board-ready workbook.

---

## 7) Smoke Tests

### Finance
```powershell
cd c:\IRONLOG
$env:API_BASE="http://localhost:3001"
node api/scripts/finance-smoke.js
```

### Enterprise
```powershell
cd c:\IRONLOG
$env:API_BASE="http://localhost:3001"
node api/scripts/enterprise-smoke.js
```

---

## 8) API Quick Reference

### Entity (`/api/entity`)
- `POST /companies`
- `GET /companies`
- `POST /sites`
- `GET /sites`
- `GET /tree`
- `POST /currency/rates/upsert`
- `GET /currency/rates`
- `GET /currency/convert?from=X&to=Y&amount=N`
- `POST /tax/profiles/upsert`
- `GET /tax/profiles`

### Integrations (`/api/integrations`)
- `GET /connectors`
- `POST /connections/upsert`
- `GET /connections`
- `POST /jobs/enqueue`
- `POST /jobs/:id/run`
- `POST /jobs/:id/cancel`
- `POST /jobs/:id/retry-now`
- `GET /jobs`
- `GET /jobs/:id`
- `GET /dead-letter`
- `POST /dead-letter/:id/acknowledge`
- `GET /monitoring/summary`

### Governance (`/api/governance`)
- `GET /policies`
- `POST /policies/upsert`
- `POST /evaluate`
- `GET /violations`
- `POST /violations/:id/acknowledge`

### Executive (`/api/executive`)
- `POST /evidence-packs/build`
- `GET /evidence-packs`
- `GET /evidence-packs/:id`
- `GET /command-center?period=YYYY-MM`
- `GET /board-pack?period=YYYY-MM`
- `GET /board-pack/export.xlsx?period=YYYY-MM`

---

## 9) Troubleshooting

- **Tabs missing**: clear saved overrides with `localStorage.removeItem("ironlog:tabs:override")`.
- **Integration job stuck**: click `Retry Now`; check `Dead-letter` list if exhausted.
- **Pack integrity FAIL**: regenerate the pack; do not trust stale row.
- **Board pack narrative empty**: confirm period has source activity (daily_hours, breakdowns, budgets, forecasts).
