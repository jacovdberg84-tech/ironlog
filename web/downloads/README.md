# IRONLOG Notify APK

Place the built Android APK here as **`ironlog-notify.apk`** so technicians can download it from:

- Technician Terminal → Quick links
- Work Order QR page

## Build the APK

From `c:\IRONLOG\mobile-notify`:

```powershell
npm install
eas build -p android --profile preview
```

Download the APK from EAS and copy it to this folder:

```powershell
Copy-Item .\path\to\build.apk c:\IRONLOG\web\downloads\ironlog-notify.apk
```

## Firebase (server push)

1. Create a Firebase project and add an Android app (`com.aml.ironlog.notify`).
2. Download `google-services.json` into `mobile-notify/`.
3. In Firebase → Project settings → Service accounts → Generate new private key.
4. On the IRONLOG API server, set in `api/.env`:

```
FCM_SERVICE_ACCOUNT_PATH=C:\path\to\firebase-service-account.json
# or inline:
# FCM_SERVICE_ACCOUNT_JSON={"type":"service_account",...}
```

5. Restart the API. `GET /api/notifications/config` should show `"push_enabled": true`.

## Test push

After a technician signs into the mobile app and grants notification permission:

```http
POST /api/notifications/test
Authorization: Bearer <admin token>
Content-Type: application/json

{"username": "jsmith"}
```
