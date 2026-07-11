# Connecting the Checkpoint App to a Google Sheet

The app posts every event (sign-in, each answer, completion) to a URL you set in
`questions.json` → `config.endpoint`. That URL comes from a **Google Apps Script
Web App** attached to your Sheet. One-time setup, ~5 minutes.

## 1. Create the Sheet

1. Go to [sheets.new](https://sheets.new) and create a blank spreadsheet.
2. Name it something like `ETV Checkpoint Responses`.

You don't need to add headers — the script below creates a `Responses` tab with
headers on the first submission.

## 2. Add the Apps Script

1. In the Sheet, open **Extensions → Apps Script**.
2. Delete any placeholder code and paste this:

```javascript
const SHEET_NAME = "Responses";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000); // avoid clobbering when the whole room submits at once
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["timestamp", "session", "email", "event", "checkpoint", "code", "answer", "why"]);
      sheet.setFrozenRows(1);
    }
    const d = JSON.parse(e.postData.contents);
    sheet.appendRow([
      d.ts || new Date().toISOString(),
      d.session || "",
      d.email || "",
      d.event || "",
      d.checkpoint || "",
      d.code || "",
      d.answer || "",
      d.why || ""
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
```

3. Click the 💾 save icon (name the project anything).

## 3. Deploy as a Web App

1. Click **Deploy → New deployment**.
2. Click the gear ⚙️ next to "Select type" and choose **Web app**.
3. Set:
   - **Execute as:** `Me (your@email)`  ← rows are written with *your* permission
   - **Who has access:** `Anyone`  ← required so attendees' browsers can POST without signing in
4. Click **Deploy**, approve the authorization prompts (you may need to click
   "Advanced → Go to … (unsafe)" — it's your own script, this is normal).
5. Copy the **Web app URL**. It looks like:
   `https://script.google.com/macros/s/AKfycb.../exec`

## 4. Paste the URL into questions.json

```json
"config": {
  "endpoint": "https://script.google.com/macros/s/AKfycb.../exec",
  "sessionId": "mathfest-2026"
}
```

Leave `endpoint` as `""` to run in local-only mode (answers stay in each
device's localStorage; recover them via the `#export` trick below).

## 5. Test it

Serve the app over HTTP (the JSON fetch doesn't work from `file://`):

```bash
python3 -m http.server 8000
# then open http://localhost:8000/engagement_app.html
```

Sign in with a test email, unlock a checkpoint (e.g. code `TIDY`), submit an
answer, and check the Sheet — a `Responses` tab should appear with your rows.

You can also test the endpoint directly from a terminal:

```bash
curl -L -d '{"session":"test","email":"me@test.com","event":"answer","checkpoint":1,"code":"TIDY","answer":"hello","why":"testing"}' \
  "https://script.google.com/macros/s/AKfycb.../exec"
```

## Gotchas / good to know

- **Editing the script later?** Changes do NOT go live until you redeploy:
  **Deploy → Manage deployments → ✏️ edit → Version: New version → Deploy.**
  The URL stays the same.
- **Editing questions.json** needs no redeploy of anything — attendees just get
  the new questions on page load. (If you host on GitHub Pages, allow a minute
  or two for its cache.)
- The app sends with `sendBeacon`/`no-cors`, so it can't read the script's
  response — failures are silent by design to keep the room moving. Every
  answer is also saved to the device's localStorage as a backup: visit the
  page with `#export` appended (e.g. `…/engagement_app.html#export`) on any
  device to download that device's answers as CSV.
- One row is written per event. Filter the `event` column to `answer` for the
  actual responses; `signin`/`complete` rows are useful for attendance and
  drop-off.
- To run a second session later, change `sessionId` in `questions.json` — the
  same Sheet keeps collecting, and you can filter by the `session` column.

## Hosting the app for attendees

Any static host works. Simplest: push this folder to GitHub and enable
**Settings → Pages** on the repo — attendees visit
`https://<user>.github.io/<repo>/engagement_app.html`. Make a QR code of that
URL for your opening slide.
