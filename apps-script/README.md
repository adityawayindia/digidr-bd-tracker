# Connecting the tracker to Google Sheets

The tracker reads and writes the Google Sheet through a small Apps Script web app. You set it up once (about 5 minutes).

1. Open the tracker Sheet. Go to **Extensions → Apps Script**.
2. Delete the placeholder code. Paste in the whole of [Code.gs](Code.gs) and save.
3. Optional: set `TEAM_KEY` at the top of the script to a shared password. Anyone using the app will need it.
4. In the function dropdown, pick **setup** and click **Run**. Approve the permissions when asked.
   - This creates the tabs `Doctors`, `Calls`, `Plan`, `Pitch` and `Activity`.
   - It also imports any existing Google Form responses.
5. Click **Deploy → New deployment**. Choose **Web app** and set:
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Click **Deploy** and copy the **Web app URL**. It ends in `/exec`.
7. In the tracker, set `API_URL` near the top of the script in `digidr-bd-tracker.html` to that URL (and `TEAM_KEY` if you set one).

**After editing Code.gs**, go to **Deploy → Manage deployments → Edit → Version: New version → Deploy**. This keeps the same URL.

The old Google Form is no longer used. You can unlink it from the Sheet once the import has run.
