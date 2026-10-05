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
- **A check-in writes `Present`** into that student's cell for the shift. A
  check-in more than **30 minutes after the shift starts** writes `Partial`
  instead; change the minutes, or blank them out to turn this off, in Check-in
  Settings. It replaces `Not Present`, `Absent Excused` or `Absent Unexcused`,
  but never a `Partial` or `Unproductive` a lead already set.
- **Nothing else changes.** Formulas (Shift Score, counts), the row-30
  checkboxes, mentor/parent rows and every other tab are left alone.

**Parents** check in the same way, with the same code and the same timing
rules:

- **Already signed up** in one of the shift's parent rows (`Female Parent`,
  `Parent 1`–`Parent 6`): the check-in is confirmed and logged, and the sheet
  isn't changed. Their sign-up can be the full name, a family alias from
  `student-parent-mapping`, or just their first name if no other parent
  shares it.
- **Not signed up:** their full name ("First Last", as Parent Hours Status
  counts it) goes into the first empty `Parent 1`–`Parent 6` slot.
  `Female Parent` is never auto-filled, since leads assign it.
- Parents are never marked `Partial`, and a parent check-in never touches
  student rows.

It adds three small tabs of its own:

| Tab | Contents |
|---|---|
| **Check-in Roster** | Name, Google email, Role (`Student`/`Parent`), and other spellings. Setup fills in students from the attendance tab and parents (with aliases) from `student-parent-mapping`; you add the emails. |
| **Check-in Log** | Every check-in, **rejected code**, and admin edit, with time and email. |
| **Check-in Settings** | Attendance tab name, check-in on/off, admin emails, org name. |

## The website

| Page | Who | What |
|---|---|---|
| `/` | Students & parents | Sign in with Google, then scan the QR code or type the 6-digit code. |
| `/display.html` | Admins | Put it on a projector or tablet. It shows the live code, a countdown, the QR code and which shift check-ins are going to right now. If a day has more than one location (e.g. Hangar 391 + Online), pick the location; for **Online**, screen-share it on Zoom. |
| `/admin.html` | Admins | Pick a date and shift, then set any student's status from the sheet's own dropdown values. Each change is logged. The **Admins** section adds or removes admins by email. |

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

### Step 1: Create a separate Apps Script project

> **Don't put this code in the spreadsheet's existing script.**
> *SOR Signups/Attendance '26-'27* already has its own Apps Script project
> (Code.gs, Prune.gs, SlackTrigger.gs, GroupPastDates.gs). Check-in is a
> separate project that opens the sheet by its ID, so neither can break the
> other.

1. Go to <https://script.google.com> and click **New project**. Rename it
   (top left) to **SOR Check-in**.
2. Open the default `Code.gs`. It's empty apart from `myFunction`, so replace
   all of it with this repo's `src/Code.js`.
3. Click **+ → Script**, name it `Logic`, and paste `src/Logic.js`.
4. Click **⚙ Project Settings**, tick **Show "appsscript.json" manifest file**,
   go back to the editor, open `appsscript.json`, and paste `src/appsscript.json`.
5. Save (Ctrl/Cmd+S).
6. In the toolbar's function dropdown, pick **`setup`** and click **Run**.
   Google asks for permission:
   **Review permissions → choose the account → Advanced → Go to SOR Check-in
   (unsafe) → Allow**. It says "unsafe" only because this is your own
   unpublished script. This adds three tabs to the attendance spreadsheet
   (*Check-in Settings*, *Check-in Roster*, *Check-in Log*), and copies every
   student and parent into **Check-in Roster**.
7. **You don't type any emails.** People are linked to the roster the first
   time they sign in:
   - **Automatically,** if their Google email is on the registration form
     ("2026-2027 SOR Student Information & Roster", tab *RAW DATA*). A
     student's *Student Email* links them to their attendance row. A
     *Parent/Guardian* email links them as that parent, and adds them to the
     roster if they're missing. A middle name on the form doesn't matter
     ("Shravani Swapnil Lad" matches "Shravani Lad"). For a nickname, type
     the form's name in the roster's **Also matches** column (e.g.
     `Katherine Zadorognuk` on Katya's row).
   - **Otherwise,** the page asks *"Who are you?"*. Students pick their name;
     parents pick or type theirs and pick their student. A name whose email
     *is* on the form can't be claimed by another account; the page shows a
     hint (`ab***@gmail.com`) of which account to use.

   Every link is logged in Check-in Log (`Student linked` / `Parent linked` /
   `Parent registered`, marked `Auto` or `Self`). To undo a wrong one, clear
   the email in Check-in Roster. The account that runs the script needs at
   least **view** access to the registration-form spreadsheet.
8. In **Check-in Settings**:
   - You don't need to touch **Admin emails**. The account that created the
     project and `nhstedd@gmail.com` are always admins, and any admin can add
     or remove others on `/admin.html` → **Admins**. Those changes are saved
     in this row.
   - **Attendance tab** is `Offseason 2026`. Change it when build season gets
     its own tab.

The spreadsheet ID and the Firebase API key are already filled in at the top
of `Code.js` (`SPREADSHEET_ID`, `FIREBASE_API_KEY`). Only change them if you
point it at a different spreadsheet or Firebase project.

### Step 2: Firebase project (already done for `sor-attendance`)

The Firebase project **sor-attendance** exists, Google sign-in is enabled,
and `public/config.js` already has its values. For a new project you would:

1. Go to <https://console.firebase.google.com> → **Create a project**.
2. **Security → Authentication → Get started → Sign-in method → Google →
   Enable**, pick a support email, **Save**.
3. **Project Overview → + Add app → Web `</>`** → register. Copy `apiKey`,
   `authDomain`, `projectId` and `appId` into `public/config.js`, and the
   same `apiKey` into `FIREBASE_API_KEY` in `Code.js`.

### Step 3: Publish the script as an API

1. In the **SOR Check-in** Apps Script editor: **Deploy → New deployment**. Click ⚙ next to
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
  `fillRosterNames` from the SOR Check-in editor (function dropdown → **Run**)
  and fill in their email.
- **New season tab:** change *Attendance tab* in Check-in Settings. The new tab
  needs the same column-A labels (`Date`, `Shift Number`, `Location`,
  `Start Time`, `End Time`, `Avg Attendees`).
- **Late arrivals:** marked `Partial` automatically after the *Partial after*
  minutes. A lead can still change it.
- **Unproductive / excused absences:** set them in the sheet as before, or on
  `/admin.html`.

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
| *Your sign-in expired* on every request | `FIREBASE_API_KEY` in `Code.js` must be the same `apiKey` as in `config.js`. If you restricted that key in Google Cloud, do **not** use an "HTTP referrers" restriction, because the script calls it from Google's servers. |
| *Server error* / *Failed to fetch* | Check `scriptUrl` ends in `/exec`, and that the deployment's access is **Anyone**. |
| *… is not on the Check-in Roster* / "Who are you?" keeps appearing | Their Google email isn't on the registration form (or their name there differs from the attendance tab). They can pick their name, or a lead can type the email into column B of Check-in Roster. |
| *… is not a row in the attendance tab* | The name on Check-in Roster must match column A of the attendance tab (case and extra spaces don't matter). |
| *All parent slots for … are full* | Parent 1–6 are all taken for that shift. A lead adds them by hand (e.g. in a spare row or a note). |
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
