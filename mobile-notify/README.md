# IRONLOG Notify (mobile)

Small Android app for technicians — receives FCM push alerts for breakdowns and assigned work orders.

## Setup

1. **Firebase**
   - Create project at [Firebase Console](https://console.firebase.google.com/).
   - Add Android app: package `com.aml.ironlog.notify.preview` (preview APK) or `com.aml.ironlog.notify` (production).
   - Download `google-services.json` into this folder (`mobile-notify/google-services.json`).
   - Generate a service account key for the API server (see `web/downloads/README.md`).

2. **Expo / EAS**
   - `npm install -g eas-cli` (if needed)
   - `cd mobile-notify && npm install`
   - `eas init` — creates project and updates `app.json` `extra.eas.projectId`
   - Upload FCM credentials: `eas credentials` (Android → Google Service Account)

3. **Build APK**

```powershell
cd c:\IRONLOG\mobile-notify
npm run build:apk
```

Copy the APK to `c:\IRONLOG\web\downloads\ironlog-notify.apk`.

## Local dev

```powershell
npm start
```

Use a physical Android device with a dev build for real FCM tokens (Expo Go uses Expo push, not native FCM).

## App flow

1. Enter IRONLOG server URL (default production).
2. PIN sign-in (same roster as Technician Terminal).
3. App registers FCM device token with `POST /api/notifications/register`.
4. Tapping a notification opens the work order QR page in the browser.
