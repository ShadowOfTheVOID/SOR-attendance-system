# SOR Attendance System

Members check in on their phones and the result goes straight into a Google Sheet.
Two checks keep it honest:

1. **Google sign-in.** Each member signs in with the Google account listed on
   the roster. Nobody can type in someone else's name.
2. **A rotating 6-digit code on a screen in the room.** The code changes every
   30 seconds, and check-in needs the code that is showing *right now*. Codes
   are shown as a big number and a QR code, so someone at home has nothing to
   enter.

```
 Phone ──Google sign-in──▶ Firebase Auth
   │  (ID token + code)
   ▼
 Firebase Hosting site ──POST──▶ Apps Script API ──▶ Google Sheet
 (check-in / display / admin)    (verifies token,       Roster · Attendance
                                  checks code)          Log · Settings
```

## What's in the sheet

| Sheet | Contents |
|---|---|
| **Roster** | `ID · Name · Group · Active · Email`. You maintain this. **Email** must be the member's Google account. |
| **Attendance** | One row per member and one column per date. Cells are `Present` / `Late` / `Absent` / `Excused`, color-coded. |
| **Log** | Every check-in, admin edit, auto-absent mark and **rejected code attempt**, with the account email. |
| **Settings** | Org name, late cutoff, check-in on/off, auto-absent hour, **admin emails**. |

## The website

| Page | Who | What |
|---|---|---|
| `/` | Members | Sign in with Google, then enter the code (or scan the QR, which fills it in automatically). |
| `/display.html` | Admins | Put this on the projector or a tablet at the door. It shows the live code, a countdown bar and the QR code. |
| `/admin.html` | Admins | Pick a date or group and fix statuses (Excused, Late, …). Each change is logged with the admin's email. |

---

## Setup (about 20 minutes, one time)

You need a Google account, and Node.js installed on your computer
(<https://nodejs.org>, LTS version). Node is only used to upload the website.

### Step 1: Google Sheet and script

1. Create a new Google Sheet (for example "SOR Attendance").
2. Open **Extensions → Apps Script**.
3. Open `Code.gs`, delete what's there, and paste this repo's
   `src/Code.js`.
4. Click **+ → Script**, name it `Logic`, and paste `src/Logic.js`.
5. Click **⚙ Project Settings**, tick **Show "appsscript.json" manifest file**,
   go back to the editor, open `appsscript.json`, and paste `src/appsscript.json`.
6. Save (Ctrl/Cmd+S). Go back to the sheet and **reload the page**. An
   **Attendance** menu appears.
7. Click **Attendance → Set up sheets**. Google asks for permission:
   **Continue → choose your account → Advanced → Go to (project) (unsafe) → Allow**.
   It says "unsafe" only because this is your own unpublished script.
8. Open the **Roster** sheet and replace the example rows with your members.
   Fill in each person's **Google email**.
9. Open the **Settings** sheet. **Admin emails** already contains your email.
   Add other admins, separated by commas.

### Step 2: Firebase project (for Google sign-in and hosting)

1. Go to <https://console.firebase.google.com> → **Create a project**. Name it
   (for example `sor-attendance`). You can turn Google Analytics off.
2. Left menu: **Build → Authentication → Get started → Sign-in method → Google →
   Enable**. Pick a support email and click **Save**.
3. Click **⚙ (Project settings) → General**. Under **Your apps**, click the
   **Web `</>`** icon, give it a nickname, and click **Register app**. You don't
   need to tick Hosting here.
4. Firebase shows a `firebaseConfig = { … }` block. Open `public/config.js` in
   this repo and copy `apiKey`, `authDomain`, `projectId` and `appId` into the
   `firebase` section.
5. Back in the Google Sheet: **Attendance → Set Firebase API key…**. Paste the
   same `apiKey` value. The script uses it to check that sign-ins really came
   from your Firebase project.

### Step 3: Publish the script as an API

1. In the Apps Script editor: **Deploy → New deployment**. Click ⚙ next to
   "Select type" and choose **Web app**.
2. Set **Execute as: Me** and **Who has access: Anyone**. It must be *Anyone*,
   not "Anyone with Google account", or the website can't reach it. The script
   does its own sign-in check.
3. Click **Deploy** and authorize again if asked. Copy the **Web app URL**
   (ends in `/exec`).
4. Paste it into `public/config.js` as `scriptUrl`.

### Step 4: Upload the website

In a terminal, inside this repo folder:

```bash
npm install -g firebase-tools
firebase login
firebase use --add            # pick your project, alias: default
firebase deploy --only hosting
```

It prints your site address, for example `https://sor-attendance.web.app`.

### Step 5: Try it

1. On a laptop or projector, open `https://<your-site>/display.html`, sign in
   with an admin account, and click **Full screen**.
2. On your phone, scan the QR code and sign in with a Google account that is on
   the Roster. You should see ✓ *Present*, and the **Attendance** sheet fills in.
3. Optional: run **Attendance → Install daily auto-absent trigger** so anyone
   who didn't check in is marked `Absent` each night.

---

## Day-to-day

- **Start of class:** open the display page. With no display open, nobody can
  check in.
- **Close check-in:** set *Self check-in enabled* to `FALSE` in Settings.
- **Late:** check-ins after *Late after (HH:mm)* count as `Late`. Leave it
  blank to disable. Times use the sheet's time zone (**File → Settings**).
- **New member:** add a row to Roster with their Google email. No redeploy needed.
- **Someone can't check in:** an admin marks them on `/admin.html`.

## Updating the code later

- **Script changes:** paste the new code, then **Deploy → Manage deployments →
  ✏️ Edit → Version: New version → Deploy**. The URL stays the same. A plain
  *New deployment* creates a new URL.
- **Website changes:** run `firebase deploy --only hosting`.
- Optional: use `clasp` (`npm i -g @google/clasp`, copy `.clasp.json.example`
  to `.clasp.json` with your Script ID, then `npm run push`) instead of
  copy-pasting.

## Troubleshooting

| Message | Fix |
|---|---|
| *Server not configured: run Attendance > Set Firebase API key…* | Do Step 2.5. |
| *Your sign-in expired* on every request | The key from Step 2.5 must be the same `apiKey` as in `config.js`. If you restricted that key in Google Cloud, do **not** use an "HTTP referrers" restriction, because the script calls it from Google's servers. |
| *Server error* / *Failed to fetch* | Check `scriptUrl` ends in `/exec`, and that the deployment's access is **Anyone**. |
| *… isn't on the roster* | The Email cell must match the Google account exactly, and **Active** must be ticked. |
| *auth/unauthorized-domain* | If you use a custom domain, add it under Firebase **Authentication → Settings → Authorized domains**. |
| *Too many wrong codes* | 5 wrong codes lock that account for 10 minutes. The Log sheet shows each attempt. |

## What this does and doesn't stop

- ✅ **Checking in from home.** The code is only visible in the room and dies
  within 30–60 seconds.
- ✅ **Checking in a friend by typing their name.** You can only check in as the
  Google account you're signed into.
- ✅ **Guessing codes.** Codes come from a secret on the server, and there are
  5 tries per 10 minutes.
- ⚠️ **Someone in the room texting the code to a friend** who checks in within
  about a minute is still possible. The Log records exact check-in times and
  emails, so it leaves a trail.
- ⚠️ **Someone signing in on their phone with a friend's Google password.**
  Nothing short of in-person checks stops that.

## Development

```bash
npm test   # logic + API tests against an in-memory fake of Sheets, token check and HMAC
```
