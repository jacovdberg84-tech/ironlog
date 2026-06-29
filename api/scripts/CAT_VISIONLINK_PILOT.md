# CAT VisionLink Pilot (426 TLB)

Use this checklist to start the first CAT-backed telematics feed into IRONLOG.

## 1) Get CAT API credentials

1. Sign in to [digital.cat.com](https://digital.cat.com).
2. Subscribe to:
   - **ISO 15143-3 API** (recommended first), or
   - **VisionLink API v3** (if you need richer CAT-specific fields).
3. Wait for CAT to issue:
   - `Client ID`
   - `Client Secret`
   - API endpoint URLs and scope (if required).

## 2) Register the machine in IRONLOG

In IRONLOG User Admin -> Telematics:

- `asset_code`: `JZ400729` (or your final IRONLOG code for the 426)
- `device_serial`: `23032800H0ZR0054`
- `unit_model`: `CAT-VISIONLINK`
- `external_id`: optional CAT equipment ID

## 3) Configure environment

Add to `api/.env` (or process env):

```env
VISIONLINK_TOKEN_URL=https://<cat-auth-token-endpoint>
VISIONLINK_CLIENT_ID=<cat-client-id>
VISIONLINK_CLIENT_SECRET=<cat-client-secret>
VISIONLINK_SCOPE=<optional-scope>
VISIONLINK_LASTREPORTED_URL=https://<cat-telemetry-endpoint>
VISIONLINK_ASSET_MAP_JSON={"23032800H0ZR0054":"JZ400729"}
IRONLOG_API_BASE=http://127.0.0.1:3001
```

If ingest auth is enabled, include:

```env
TELEMATICS_API_KEY=<same value expected by /api/telematics/ingest>
```

## 4) Poll once (manual test)

From `c:\IRONLOG\api`:

```powershell
npm run telematics:visionlink:poll
```

Expected output:

- `processed=<n> ingested=<n> failed=0`

## 5) Verify in IRONLOG UI

1. Open Telematics tab.
2. Confirm 426 unit shows as live.
3. Confirm engine hours update.
4. Open Daily Input and verify meter source is telematics and opening/closing align.

## 6) Production schedule

Once manual poll is successful, schedule the poll command every 10-15 minutes
with PM2/Task Scheduler/cron, then monitor for 2-3 days before enabling strict lock policy for additional CAT assets.
