# SOR Attendance System

A mobile-friendly attendance web app that writes straight into a Google Sheet.
It runs as a Google Apps Script bound to the spreadsheet, so you don't need
servers, API keys, or service accounts.

## What you get

| Sheet | Contents |
|---|---|
| **Roster** | `ID · Name · Group · Active`. You maintain this list. |
| **Attendance** | One row per member and one column per date (`yyyy-MM-dd`). Each cell is `Present`, `Late`, `Absent` or `Excused`, color-coded. |
| **Log** | Append-only audit trail of every check-in, admin edit and auto-absent mark. |
| **Settings** | Organization name, late cutoff time, self check-in on/off, auto-absent hour. |

**Web app**
- **Check in**: a member types their ID or full name (names autocomplete).
  They are marked `Present`, or `Late` after the cutoff time. Repeat check-ins
  are ignored, and a check-in never overwrites an `Excused` mark.
- **Admin** (PIN-protected): pick any date, filter by group, set or correct
  statuses, bulk-mark unmarked people, and save. Every change is logged.

**Sheet menu → Attendance**
- *Set up sheets*: creates the four sheets. It is safe to re-run.
- *Set admin PIN…*: the PIN is stored in Script Properties, not in the sheet.
  The admin is locked out for 10 minutes after 10 wrong attempts.
- *Mark unmarked as Absent (today)*
- *Install daily auto-absent trigger*: runs the step above every night.

## Install (about 5 minutes)

### Option A: copy and paste (no tools needed)
1. Create a new Google Sheet.
2. Open **Extensions → Apps Script**.
3. Create files matching `src/`. Paste `Code.js` and `Logic.js` as script
   files (`Code.gs`, `Logic.gs`) and `Index.html` as an HTML file named `Index`.
4. Open **Project Settings**, tick *Show "appsscript.json"*, and replace its
   contents with `src/appsscript.json`.
5. Reload the sheet. Run **Attendance → Set up sheets** and authorize when prompted.
6. Fill in **Roster**, then run **Attendance → Set admin PIN…**.
7. In the Apps Script editor, go to **Deploy → New deployment → Web app**.
   Set *Execute as: Me* and *Who has access: Anyone* (or *Anyone within your
   domain*). Share the URL or turn it into a QR code at the door.

### Option B: `clasp` from this repo
```bash
npm i -g @google/clasp && clasp login
cp .clasp.json.example .clasp.json   # paste the Script ID from Apps Script → Project Settings
npm run push                         # then do steps 5–7 above
```

After editing code, use **Deploy → Manage deployments → Edit → New version**
so the existing URL keeps working.

## Configuration (Settings sheet)
- **Late after (HH:mm)**: check-ins after this time count as `Late`. Leave it
  blank to disable.
- **Self check-in enabled**: set it to `FALSE` to close check-in. Admin
  editing still works.
- **Auto-mark absent time (hour 0-23)**: used when you install the daily trigger.

Times use the spreadsheet's time zone (**File → Settings**).

## Security notes
- Anyone with the link can check in **any** roster member by name or ID. If
  that matters, give members non-guessable IDs and restrict deployment
  access to your Google Workspace domain.
- Admin actions require the PIN on every request, checked server-side.

## Development
```bash
npm test   # logic unit tests plus server tests against an in-memory fake of SpreadsheetApp
```
