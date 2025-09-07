# OPD Logger – APK Build (Capacitor)

This project is ready to turn the static `www/` app into an Android APK using Capacitor.

## Prerequisites
- Node.js 18+ and npm
- Android Studio (with Android SDK + Build-Tools)
- Java 17 (Temurin recommended)
- USB debugging enabled on device (optional)

## Install
```bash
npm install
```

## Generate icons (replace resources/icon.png with your 1024x1024 green OPD logo first)
```bash
npx @capacitor/assets generate --android
```

## Add Android platform (first time only)
```bash
npm run cap:add:android
```

## Sync web files into Android project
```bash
npm run cap:sync
```

## Open Android Studio
```bash
npm run cap:open:android
```

Then in **Android Studio**:
- Select **app** module
- Build > Build Bundle(s) / APK(s) > **Build APK(s)**
- Locate the APK in `android/app/build/outputs/apk/debug/` or `release/`

## Notes
- The app uses Capacitor **Filesystem** and **Share** plugins for export. They are already listed in `package.json`.
- The web app is fully offline-enabled via `service-worker.js`. When you update files, bump the cache name inside it.
- If you change `app.js`, also bump the query string in `index.html` like `app.js?v=18`.
