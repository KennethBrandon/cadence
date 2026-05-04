# cadence

Personal Google Apps Script projects that pull data from Fitbit (via the
Google Health API) and Strava into Google Sheets, compute training-load and
recovery metrics on top, and send a weekly AI coaching report.

## Layout

```
cadence/
├── workout-tracker/   ← Apps Script project bound to the workout spreadsheet
│   ├── strava-sync.gs           pulls Strava activities → "Strava Workouts" sheet
│   ├── workout-merge.gs         merges Strava + Fitbit workouts, deduplicates
│   ├── training-load.gs         CTL / ATL / TSB (fitness, fatigue, form)
│   └── inspect-strava-activity.gs   one-off Strava activity debug helper
│
├── daily-health/      ← Apps Script project bound to the daily-data spreadsheet
│   ├── daily-health-sync.gs     pulls Google Health metrics → daily sheet
│   ├── workout-sync.gs          pulls Google Health workouts → workout sheet
│   ├── readiness-score.gs       1–100 daily readiness score from rolling baselines
│   └── weekly-coach*.gs         Gemini-powered weekly coach report (8 files)
│
└── docs/              Notes and bug reports
```

Each project subfolder is a separate `clasp` root with its own `.clasp.json`.

## First-time setup on a new machine

```bash
# Prereqs
brew install node gh
gh auth login

# Clone the repo
git clone <repo-url> cadence
cd cadence

# Authenticate clasp
npx --yes @google/clasp@latest login
```

You'll also need the **Google Apps Script API** enabled at
<https://script.google.com/home/usersettings>.

## Pushing local changes to Apps Script

```bash
cd workout-tracker     # or daily-health
npx --yes @google/clasp@latest push -f
```

Pulling the other direction (e.g. after editing in the Apps Script web UI):

```bash
npx --yes @google/clasp@latest pull
```

## Required Script Properties

Set these in **Project Settings → Script Properties** in each Apps Script
project. They keep credentials and personal config out of source.

### `workout-tracker`

| Key | Purpose |
|---|---|
| `STRAVA_CLIENT_ID` | Strava API client ID |
| `STRAVA_CLIENT_SECRET` | Strava API client secret |
| `STRAVA_REFRESH_TOKEN` | Long-lived Strava refresh token |

### `daily-health`

| Key | Purpose |
|---|---|
| `GOOGLE_HEALTH_CLIENT_ID` | OAuth client ID for the Google Health API |
| `GOOGLE_HEALTH_CLIENT_SECRET` | OAuth client secret |
| `WORKOUT_SPREADSHEET_ID` | ID of the workout-tracker spreadsheet (workout sync writes here) |
| `WORKOUT_SS_ID` | Same value as above; coach reads from it |
| `GEMINI_API_KEY` | Google AI Studio key for the weekly coach |
| `REPORT_EMAIL` | Where the weekly coach report is sent |
| `ATHLETE_PROFILE` | JSON: `{name?, dob, sex, heightInches, primarySports[], maxHR, injuryHistory[], goals[]}`. Optional `name` is shown in the report footer. The goal-progress detection parses `goals` for patterns like `"Reach NNN lbs"` or `"half marathon"`. |

After setting properties, run `authorize()` (in `daily-health-sync.gs`) once to
grant OAuth scopes for the Google Health API.

## Triggers

Both projects rely on time-based triggers installed via one-time setup
functions inside the script (e.g. `INSTALL_HEALTH_TRIGGER`). Run those once
from the Apps Script editor; verify in **Triggers** in the left sidebar.
