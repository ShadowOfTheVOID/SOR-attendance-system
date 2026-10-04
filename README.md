# SOR Attendance System

Students check in on their phones and it's marked straight into the team's
attendance spreadsheet.
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
 Firebase Hosting site ──POST──▶ Apps Script API ──▶ SOR Signups/Attendance
 (check-in / display / admin)    (verifies token,       student's cell in the
                                  checks code)          shift's column
```

## How it fits the team sheet

It works directly on **"SOR Signups/Attendance '26-'27"**, in the tab named in
settings (default **Offseason 2026**):

- **Each column is a shift.** The script finds the header rows by their
  column-A labels: `Date`, `Shift Number`, `Location`, `Start Time`, `End Time`.
  Edited times like `3:45 PM --> 4:45` use the new time.
- **Students are the rows below `Avg Attendees`.** The name is in column A.
- **A check-in writes `Present`** into that student's cell for the shift. It
  replaces `Not Present`, `Absent Excused` or `Absent Unexcused`, but never
  `Partial` or `Unproductive`. Those stay the leads' call.
- **Nothing else changes.** Formulas (Shift Score, counts), the row-30
  checkboxes, mentor/parent rows and every other tab are left alone.

It adds three small tabs of its own:

| Tab | Contents |
|---|---|
| **Check-in Roster** | Student name (exactly as in column A) and their Google email. Setup fills in the names; you add the emails. |
| **Check-in Log** | Every check-in, **rejected code**, and admin edit, with time and email. |
| **Check-in Settings** | Attendance tab name, check-in on/off, admin emails, org name. |

## The website

| Page | Who | What |
|---|---|---|
| `/` | Students | Sign in with Google, then scan the QR code or type the 6-digit code. |
| `/display.html` | Admins | Put it on a projector or tablet. It shows the live code, a countdown, the QR code and which shift check-ins are going to right now. If a day has more than one location (e.g. Hangar 391 + Online), pick the location; for **Online**, screen-share it on Zoom. |
| `/admin.html` | Admins | Pick a date and shift, then set any student's status from the sheet's own dropdown values. Each change is logged. |

### Which shift does a check-in count for?

The **time of the check-in** decides, using that day's Start/End Time rows for
the display's location:

- **During a shift:** it counts for that shift.
- **During an overlap** (e.g. Shift 1 9:45–1:15 and Shift 2 1:00–4:15): it
  counts for the **next** shift.
- **Up to 30 minutes before a shift:** it counts for that shift, unless another
  shift is still running.
- **Outside all shifts:** check-in is refused, and it doesn't count as a wrong
  code.

A student who stays for two shifts checks in once for each. The display can
stay open all day, because it always shows which shift check-ins are going to.

---

## Setup (about 20 minutes, one time)

You need edit access to the attendance spreadsheet, and Node.js
(<https://nodejs.org>, LTS version) on your computer for uploading the website.

> **Do this from the team account (`sor.frc6059@gmail.com`) if you can.** The
> script runs as whoever sets it up. If that's a personal account and it loses
> access to the sheet later, check-in stops working.

### Step 1: Add the script to the attendance spreadsheet

1. Open **SOR Signups/Attendance '26-'27** and go to **Extensions → Apps Script**.
2. Open `Code.gs`, delete what's there, and paste this repo's `src/Code.js`.
3. Click **+ → Script**, name it `Logic`, and paste `src/Logic.js`.
4. Click **⚙ Project Settings**, tick **Show "appsscript.json" manifest file**,
   go back to the editor, open `appsscript.json`, and paste `src/appsscript.json`.
5. Save (Ctrl/Cmd+S). Go back to the sheet and **reload the page**. A
   **Check-in** menu appears.
6. Click **Check-in → Set up check-in tabs**. Google asks for permission:
   **Continue → choose the account → Advanced → Go to (project) (unsafe) → Allow**.
   It says "unsafe" only because this is your own unpublished script.
   This creates the three tabs and copies every student name into
   **Check-in Roster**.
7. In **Check-in Roster**, fill column B with each student's Google email. You
   can copy them from the *Student Email* column of
   "2026-2027 SOR Student Information & Roster". Each one must be the account
   they'll sign in with.
8. In **Check-in Settings**:
   - **Admin emails** already has the setup account. Add the leads/mentors who
     run the display, comma-separated.
   - **Attendance tab** is `Offseason 2026`. Change it when build season gets
     its own tab.

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
5. Back in the Google Sheet: **Check-in → Set Firebase API key…**. Paste the
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
   with an admin account, and click **Full screen**. Under the location it
   says which shift check-ins are going to.
2. On your phone, scan the QR code and sign in with a Google account that is on
   the Check-in Roster. You should see ✓ *Present for …*, and that student's
   cell in the shift's column changes to `Present`.

---

## Day-to-day

- **Shift days:** open the display page once. It follows the schedule on its
  own. With no display open, nobody can check in.
- **Times matter now:** keep each shift's Start/End Time correct in the sheet,
  because they decide where check-ins go.
- **Close check-in:** set *Self check-in enabled* to `FALSE` in Check-in Settings.
- **New student:** add their row to the attendance tab as usual, then run
  **Check-in → Add student names to Check-in Roster** and fill in their email.
- **New season tab:** change *Attendance tab* in Check-in Settings. The new tab
  needs the same column-A labels (`Date`, `Shift Number`, `Location`,
  `Start Time`, `End Time`, `Avg Attendees`).
- **Partial / Unproductive / excused absences:** set them in the sheet as before,
  or on `/admin.html`.

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
| *Server not configured: run Check-in > Set Firebase API key…* | Do Step 2.5. |
| *Your sign-in expired* on every request | The key from Step 2.5 must be the same `apiKey` as in `config.js`. If you restricted that key in Google Cloud, do **not** use an "HTTP referrers" restriction, because the script calls it from Google's servers. |
| *Server error* / *Failed to fetch* | Check `scriptUrl` ends in `/exec`, and that the deployment's access is **Anyone**. |
| *… is not on the Check-in Roster* | Add the email in column B of Check-in Roster. It must be the Google account they signed in with. |
| *… is not a row in the attendance tab* | The name on Check-in Roster must match column A of the attendance tab (case and extra spaces don't matter). |
| *There is no shift today* / wrong shifts listed | Check the `Date` row for today's column, and that *Attendance tab* in Check-in Settings is right. |
| *Could not find a "… " row in column A* | The attendance tab is missing one of the labels listed under Day-to-day. |
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
npm test   # logic + API tests against a fake copy of the attendance tab layout
```
