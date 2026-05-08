# cadence

Personal Google Apps Script projects that pull data from Strava and Fitbit (via
the Google Health API) into Google Sheets, compute training-load and recovery
metrics on top, and send a weekly AI coaching report.

The pieces are designed to be adopted **independently or together** — fork
just the Strava sync if that's all you want, or stack the whole thing.

> **Not medical advice.** The metrics, readiness scores, and AI-generated
> coaching here are for personal experimentation only. Don't make health or
> training decisions based on them — talk to a qualified provider.

## Layout

```
cadence/
├── workout-tracker/   ← Apps Script project bound to the workout spreadsheet
│   ├── strava-sync.gs           pulls Strava activities → "Strava Workouts" sheet
│   ├── training-load.gs         CTL / ATL / TSB (fitness, fatigue, form)
│   ├── workout-merge.gs         merges Strava + Fitbit workouts, deduplicates
│   └── inspect-strava-activity.gs   one-off Strava activity debug helper
│
├── daily-health/      ← Apps Script project bound to the daily-data spreadsheet
│   ├── daily-health-sync.gs     pulls Google Health metrics → daily sheet
│   ├── readiness-score.gs       1–100 daily readiness score from rolling baselines
│   ├── workout-sync.gs          pulls Google Health workouts → workout sheet
│   └── weekly-coach*.gs         Gemini-powered weekly coach report (8 files)
│
└── docs/              Notes and bug reports
```

Each project subfolder is a separate `clasp` root.

## Getting started

Pick a tier — each builds on the previous, all credentials stay isolated to
the project that needs them:

| Tier | What you get | Depends on |
|---|---|---|
| **A** | Strava → Sheets backup, plus training-load math (CTL/ATL/TSB) | — |
| **B** | Fitbit / Google Health → Sheets backup, plus daily readiness score | — |
| **C** | A + B, with workouts merged from both sources | A and B |
| **D** | Tier C plus a weekly AI coach email | C |

If you only want Tier A, you can skip every step labeled `[B]`, `[C]`, `[D]`,
and ignore the `daily-health/` folder entirely. Same idea in reverse for
Tier B.

### Step 0 — prereqs (everyone)

```bash
# Node + GitHub CLI (clasp runs via npx, no global install needed)
brew install node gh
gh auth login

# Clone this repo
git clone https://github.com/KennethBrandon/cadence.git
cd cadence

# Authenticate clasp against your Google account
npx --yes @google/clasp@latest login
```

You also need the **Google Apps Script API** enabled at
<https://script.google.com/home/usersettings>.

### Step 1 — `[A]` Strava → Sheets

1. **Create the workout spreadsheet.** New Google Sheet, name it whatever
   ("Workout Tracker" works). Add two empty sheets (tabs) inside it:
   `Strava Workouts` and `Personal Records`. The other sheets (Training Load,
   Segment Efforts, etc.) will be auto-created on first run.

2. **Create the bound Apps Script project.** From the spreadsheet, go to
   **Extensions → Apps Script**. The project's URL contains its `scriptId`
   (the long string in `script.google.com/d/<scriptId>/edit`).

3. **Wire `clasp` to your project.** In `workout-tracker/`, create
   `.clasp.json`:

   ```json
   { "scriptId": "<your workout-tracker scriptId>", "rootDir": "." }
   ```

4. **Push the code.**

   ```bash
   cd workout-tracker
   npx --yes @google/clasp@latest push -f
   ```

5. **Get a Strava refresh token.** One-time OAuth dance:

   - Create an app at <https://www.strava.com/settings/api>. Set the callback
     domain to `localhost`. Copy the **Client ID** and **Client Secret**.
   - In a browser, visit (substitute your Client ID):

     ```
     https://www.strava.com/oauth/authorize?client_id=<ID>&response_type=code&redirect_uri=http://localhost&approval_prompt=force&scope=activity:read_all,profile:read_all
     ```

   - Approve. You'll land on a `localhost` page that won't load — that's
     expected. Copy the `code=...` value out of the URL bar.
   - Exchange the code for a refresh token:

     ```bash
     curl -X POST https://www.strava.com/oauth/token \
       -d client_id=<ID> \
       -d client_secret=<SECRET> \
       -d code=<CODE_FROM_URL> \
       -d grant_type=authorization_code
     ```

   - Save the `refresh_token` from the JSON response — it's long-lived.

6. **Set Script Properties.** In the Apps Script editor: **Project Settings
   → Script Properties → Add script property** for each of:

   | Key | Value |
   |---|---|
   | `STRAVA_CLIENT_ID` | from step 5 |
   | `STRAVA_CLIENT_SECRET` | from step 5 |
   | `STRAVA_REFRESH_TOKEN` | from step 5 |

7. **Install the trigger.** In the Apps Script editor, run `INSTALL_TRIGGER`
   once. Verify under **Triggers** in the left sidebar — `runDailySync`
   should be scheduled.

8. **Test.** Run `runDailySync` manually. Recent activities should appear in
   the `Strava Workouts` tab. Run `computeTrainingLoad` to populate
   `Training Load`.

### Step 2 — `[B]` Fitbit / Google Health → Sheets

1. **Create the daily-data spreadsheet.** New Google Sheet ("Daily Health"
   works). No tabs need to pre-exist — `Daily Data` is auto-created on first
   run.

2. **Create the bound Apps Script project** (Extensions → Apps Script). Note
   its `scriptId`.

3. **Wire `clasp`.** In `daily-health/`, create `.clasp.json` with that
   scriptId, then `clasp push -f`.

4. **Create an OAuth client for the Google Health API.**

   - Go to <https://console.cloud.google.com>, create or pick a project.
   - Enable the **Google Health API** (APIs & Services → Library).
   - APIs & Services → Credentials → **Create Credentials → OAuth client ID**.
   - Application type: **Web application**.
   - Authorized redirect URI:

     ```
     https://script.google.com/macros/d/<your daily-health scriptId>/usercallback
     ```

   - Copy the resulting Client ID and Client Secret.

5. **Set Script Properties** in the daily-health Apps Script project:

   | Key | Value |
   |---|---|
   | `GOOGLE_HEALTH_CLIENT_ID` | from step 4 |
   | `GOOGLE_HEALTH_CLIENT_SECRET` | from step 4 |

6. **Authorize the script.** In the Apps Script editor, run `authorize()`
   (in `daily-health-sync.gs`). Open the **Execution log** — it prints an
   authorization URL. Visit it, grant the requested scopes (sleep, HRV,
   activity, etc.), and you'll be redirected to a "Success" page.

7. **Install the trigger.** Run `INSTALL_HEALTH_TRIGGER` once. The script will
   now sync hourly.

8. **Test.** Run `runDailyHealthSync` manually — yesterday's row should
   appear in `Daily Data`. Run `computeReadinessScores` to backfill the
   readiness column.

### Step 3 — `[C]` glue Strava and Fitbit together

You've already done both Tier A and Tier B. To merge the two:

1. **Find your workout spreadsheet's ID** — the long string in the workout
   spreadsheet's URL after `/d/`.

2. **Set one Script Property** on the **daily-health** project:

   | Key | Value |
   |---|---|
   | `WORKOUT_SPREADSHEET_ID` | the workout spreadsheet ID |

Now `daily-health/workout-sync.gs` writes Fitbit workouts into a
`FitBit Activities` tab on your *workout* spreadsheet, and
`workout-tracker/workout-merge.gs` (run via `mergeAndDeduplicateWorkouts`)
joins the two sources into a deduplicated `Merged Workouts` sheet.

The existing `spreadsheets` scope already granted in Step 2 covers the cross-
spreadsheet write — no re-authorization needed.

### Step 4 — `[D]` weekly AI coach email

1. **Get a Gemini API key** at <https://aistudio.google.com/apikey>.

2. **Set the remaining Script Properties** on the **daily-health** project:

   | Key | Value |
   |---|---|
   | `GEMINI_API_KEY` | from step 1 |
   | `WORKOUT_SS_ID` | same value as `WORKOUT_SPREADSHEET_ID` |
   | `REPORT_EMAIL` | where the weekly report is sent |
   | `ATHLETE_PROFILE` | JSON, see below |

   `ATHLETE_PROFILE` shape:

   ```json
   {
     "name": "optional, shown in report footer",
     "dob": "1990-01-01",
     "sex": "M | F",
     "heightInches": 70,
     "primarySports": ["running", "cycling"],
     "maxHR": 190,
     "injuryHistory": ["..."],
     "goals": ["Reach 165 lbs", "Sub-1:30 half marathon"]
   }
   ```

   The goal-progress detector parses `goals` for patterns like
   `"Reach NNN lbs"` or `"half marathon"`.

3. **Test.** Run `sendWeeklyCoachReport` manually. Check your inbox.

4. **Schedule it.** Add a weekly time-based trigger for
   `sendWeeklyCoachReport` (Triggers → Add Trigger → Time-driven → Week timer
   → pick a day/time, e.g. Monday 7am).

## Working on an existing setup

For an owner moving to a new machine, the path is much shorter — clone the
repo, `clasp login`, recreate `.clasp.json` in each project folder with your
existing scriptIds, then `clasp push` or `clasp pull` as needed.

## Pushing local changes to Apps Script

```bash
cd workout-tracker     # or daily-health
npx --yes @google/clasp@latest push -f
```

Pulling the other direction (e.g. after editing in the Apps Script web UI):

```bash
npx --yes @google/clasp@latest pull
```

## Script Properties reference

Quick lookup for everything covered in the setup steps above.

### `workout-tracker`

| Key | Tier | Purpose |
|---|---|---|
| `STRAVA_CLIENT_ID` | A | Strava API client ID |
| `STRAVA_CLIENT_SECRET` | A | Strava API client secret |
| `STRAVA_REFRESH_TOKEN` | A | Long-lived Strava refresh token |

### `daily-health`

| Key | Tier | Purpose |
|---|---|---|
| `GOOGLE_HEALTH_CLIENT_ID` | B | OAuth client ID for the Google Health API |
| `GOOGLE_HEALTH_CLIENT_SECRET` | B | OAuth client secret |
| `WORKOUT_SPREADSHEET_ID` | C | ID of the workout spreadsheet (workout sync writes here) |
| `WORKOUT_SS_ID` | D | Same value as `WORKOUT_SPREADSHEET_ID`; coach reads from it |
| `GEMINI_API_KEY` | D | Google AI Studio key for the weekly coach |
| `REPORT_EMAIL` | D | Where the weekly coach report is sent |
| `ATHLETE_PROFILE` | D | JSON profile (see Step 4) |

## License

Apache License 2.0 — see [LICENSE](LICENSE).
