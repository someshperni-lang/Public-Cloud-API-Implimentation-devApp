# SAP DEV Uploads (Windows + Android)

An installable app for bulk uploads to the SAP DEV tenant
`my440263-api.s4hana.cloud.sap`. It covers goods movements, supplier invoices,
purchase orders and products. The same Python code runs as a Windows app and
as an Android app (built with [Flet](https://flet.dev)).

The app talks directly from the PC or phone to SAP. No server is needed in between.

## Before the first build: add the metadata

Copy the four `$metadata` JSON files from the DEV dashboard repo into
`src/metadata/`. They switch on the field checks: unknown fields, field
lengths and data types are all checked before anything is posted.
Without them, the app still works, but it shows a notice that checks are off.

## Build in the cloud (no installs needed)

1. Push this folder to a GitHub repo.
2. In the repo, open **Actions → Build app → Run workflow**.
3. After about 10–20 minutes, open the finished run. Under **Artifacts**, download:
   - `sap-dev-uploads-windows`: unzip it and run `SAP DEV Uploads.exe`. Keep
     the whole folder together.
   - `sap-dev-uploads-apk`: the Android installer (`.apk`).

To build again after changes, run the workflow again, or push a tag such as `v1.0.1`.

## Install on Android

1. Send the `.apk` to the phone (email, Google Drive, WhatsApp or USB).
2. Open it. Android asks you to allow installs from that app (Files, Chrome,
   etc.). Allow it, then tap **Install**.
3. Open **SAP DEV Uploads**. Then go to **Connection**, enter the communication
   user and password, and tap **Test connection**.

## Build on your own PC instead (optional)

```powershell
pip install "flet==1.0.5"
flet build windows      # Windows app, in build\windows
flet build apk          # Android APK, in build\apk
```

The first build downloads the Flutter SDK automatically. The Windows build also
needs Visual Studio 2022 with the "Desktop development with C++" workload.
The APK build needs Android Studio or the Android SDK.

## Try it without building

```powershell
pip install "flet==1.0.5" openpyxl requests
flet run src/main.py
```

## Using the app

- **Upload:** choose the upload type and save the Excel or CSV template.
  Fill it in, then use **Choose file**. Each document is listed as Ready or
  Invalid with the reason. Tap a document to see the exact JSON that will be
  sent. Tick the confirmation and **Post to DEV**. When it finishes, save
  the results as CSV, or save your sheet with a status column added.
- **History:** every post made on this device is listed, with the SAP
  document numbers.
- **Connection:** host, optional client, user and password. "Remember
  password" stores it unencrypted in the app's data, so leave it off on
  shared devices.

## Project layout

```
pyproject.toml                 app name, icon, Android permission, dependencies
.github/workflows/build.yml    cloud build for Windows + Android
src/main.py                    the app screens
src/core/sap_common.py         file reading, conversion, validation, SAP client
src/core/uploads/*.py          one file per upload type (columns, rules, payload)
src/metadata/                  DEV $metadata JSON files (add these)
src/assets/icon.png            app icon
```

To add a field, add it to the column mapping near the top of the upload's file
in `src/core/uploads/`, then rebuild.
