# CAT ISO 15143 Pilot (426 TLB)

IRONLOG is wired to Caterpillar's **ISO 15143 (AEMP 2.0)** API from your OpenAPI file:

- Spec: `api/scripts/cat-visionlink-openapi.yaml`
- Base URL: `https://api.cat.com/telematics/iso15143`
- Poller: `api/scripts/poll-cat-visionlink.js`

## Your 426 details

| Field | Value |
|---|---|
| Machine serial | `JZ400729` |
| Model | `426-07LRC` |
| Telematics device serial | `23032800H0ZR0054` |
| VisionLink subscription | Connect |
| Telemetry solution | PL243 Reporting |

## 1) Wait for CAT credentials

Subscribe at [digital.cat.com](https://digital.cat.com) → **APIs** → **ISO 15143-3**.

CAT will email:

- `Client ID`
- `Client Secret`
- Auth method (usually `catFedLogin`; some accounts use Entra ID)

Typical turnaround: **5–7 business days**.

## 2) Register unit in IRONLOG

User Admin → Telematics:

- `asset_code`: your IRONLOG code for the 426
- `device_serial`: `23032800H0ZR0054`
- `unit_model`: `CAT-ISO15143`

## 3) Configure `api/.env`

```env
VISIONLINK_CLIENT_ID=<from CAT>
VISIONLINK_CLIENT_SECRET=<from CAT>
VISIONLINK_AUTH_METHOD=catFedLogin
VISIONLINK_API_MODE=equipment
VISIONLINK_EQUIPMENT_MAKE=CAT
VISIONLINK_EQUIPMENT_MODEL=426-07LRC
VISIONLINK_EQUIPMENT_SERIAL=JZ400729
VISIONLINK_ASSET_MAP_JSON={"JZ400729":"<IRONLOG_ASSET_CODE>"}
VISIONLINK_DEVICE_MAP_JSON={"JZ400729":"23032800H0ZR0054"}
IRONLOG_API_BASE=http://127.0.0.1:3001
```

If CAT issues Entra credentials instead, set:

```env
VISIONLINK_AUTH_METHOD=entraId
VISIONLINK_TOKEN_URL=https://login.microsoftonline.com/ceb177bf-013b-49ab-8a9c-4abce32afc1e/oauth2/v2.0/token
VISIONLINK_SCOPE=api://<application-id-uri>/.default
```

## 4) Dry-run now (no secrets needed)

```powershell
cd c:\IRONLOG\api
node scripts/poll-cat-visionlink.js --dry-run
```

This prints the endpoints and env keys IRONLOG will use.

## 5) First live poll (once credentials arrive)

```powershell
cd c:\IRONLOG\api
npm run telematics:visionlink:poll
```

Expected:

- `processed=1 ingested=1 failed=0`

## 6) Verify

1. Telematics tab → 426 shows live link
2. Engine hours match VisionLink (~3322 hrs and counting)
3. Daily Input → meter source = telematics

## API endpoints used

**Single machine (pilot):**

`GET /fleet/equipment/makeModelSerial/CAT/426-07LRC/JZ400729`

**Whole fleet (later):**

`GET /fleet/1` with pagination via `Links[].Rel = Next`

## Field mapping

| CAT ISO field | IRONLOG ingest |
|---|---|
| `CumulativeOperatingHours.Hour` | `engine_hours` |
| `CumulativeOperatingHours.datetime` | `recorded_at` |
| `CumulativeIdleHours.Hour` | `idle_hours` |
| `Location.Latitude/Longitude` | `latitude` / `longitude` |
| `EngineStatus.Running` | `ignition_on` |
| `EquipmentHeader.SerialNumber` | lookup key for asset/device maps |

## Schedule in production

Run every 10–15 minutes via Task Scheduler or PM2 cron once the first live poll succeeds.
