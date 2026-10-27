# SAP bulk uploads - DEV (Windows + Android app)

The installable twin of the Streamlit dashboard in `Sap_Upload_Dashboard_DEV`:
same Incresol login sidebar, banner, tabs, templates, preview, posting,
results, metadata explorer and history, for the DEV tenant
`https://my440263-api.s4hana.cloud.sap`.

The app does **not** have its own upload logic. `src/scripts/` holds the
dashboard's scripts, copied unchanged, and the app calls the same functions
the dashboard does (`build_documents`, `post_documents`, the upsert for PO
texts, the 4-sheet subcontracting workbook ...). A preview or a post in the
app gives the same result as in the dashboard.

## Keeping the app in step with the dashboard

When you change a script in the dashboard, copy it over and rebuild:

```cmd
copy /Y C:\path\to\Sap_Upload_Dashboard_DEV\scripts\*.py src\scripts\
copy /Y C:\path\to\Sap_Upload_Dashboard_DEV\metadata\*_0001.json src\metadata\
git add -A
git commit -m "Sync scripts from the dashboard"
git push
```

Then run **Actions > Build app > Run workflow**. A new tab in the dashboard
also needs a line in `TABS` at the top of `src/main.py`.

## Build

GitHub: **Actions > Build app > Run workflow**, then download the two
artifacts (Windows folder and Android APK) from the finished run.

Locally: `pip install "flet==1.0.5"` then `flet build windows` or
`flet build apk`. Try it without building: `pip install flet==1.0.5
pandas==2.2.2 openpyxl requests` and `flet run src/main.py`.

## Differences from the dashboard

- Login: optional "Remember on this device" (off by default = memory only,
  like the dashboard). On a phone the sidebar is the "Connection" menu.
- Raw file: tap a cell to correct it (the dashboard's editable grid).
- Results of every run are saved in the app's data folder on the device and
  shown in History; each device has its own history.
- History's "Document type" filter picks one type (the dashboard allows several).
