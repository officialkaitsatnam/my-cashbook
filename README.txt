MY CASHBOOK — FRESH BUILD

1. Open https://script.google.com/home and create a NEW Apps Script project.
2. Replace the editor's code with Code.gs contents. Save.
3. Run testAccess() once from the editor. Authorize Google permissions. Confirm execution log has spreadsheet name and sheet list.
4. Deploy > New deployment > Select type: Web app.
   Execute as: Me
   Who has access: Anyone (or the narrowest setting that lets your GitHub Pages app call it).
   Deploy and copy the Web app URL ending in /exec.
5. In index.html, replace PASTE_YOUR_NEW_APPS_SCRIPT_WEB_APP_URL_HERE with that exact URL. Save.
6. Upload index.html and manifest.webmanifest to your GitHub Pages repository and commit/push.
7. Open the GitHub Pages website and press Sync.

IMPORTANT:
- Do not change/delete the existing Google Sheet data.
- The script finds headers like Date, Description/Particulars, Credit, Debit and optional Balance.
- New entries need a month tab named like October-2026, Oct-2026, October 2026 or Oct 2026. If your sheet uses different tab names, tell me the exact names and we can adjust safely.
- Anyone with the deployed web app URL may be able to read/write if you deploy with Anyone access. Keep the URL private.
- Confirm that testAccess() succeeds before publishing the frontend.
