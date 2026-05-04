# Notes for Claude

## What this repo is

Two separate Google Apps Script projects pushed/pulled with `clasp`. Each
project subfolder is a `clasp` root with its own `.clasp.json`.

| Folder | Apps Script project name | Bound to | Purpose |
|---|---|---|---|
| `workout-tracker/` | "Strava Sync" (or similar) | the workout spreadsheet | Strava sync, workout merge, training-load math |
| `daily-health/` | "Fitbit Sync" | the daily-data spreadsheet | Google Health daily metrics + workout sync, readiness, weekly Gemini coach |

Note: `daily-health/workout-sync.gs` writes to the *workout-tracker* spreadsheet
(by ID, via `WORKOUT_SPREADSHEET_ID` script property) — not the daily-data
spreadsheet it's bound to. That cross-project write is intentional.

## Daily commands

`clasp` is **not** globally installed. Always invoke via `npx`:

```bash
cd workout-tracker        # or daily-health
npx --yes @google/clasp@latest push -f       # push local → remote
npx --yes @google/clasp@latest pull          # pull remote → local
npx --yes @google/clasp@latest status        # list tracked/untracked files
```

**`-f` is required for `push`** in non-interactive contexts — without it clasp
prompts for confirmation when files would be renamed or deleted on remote.

## Gotchas you (Claude) will hit

1. **No `--dry-run` flag.** clasp v3 dropped it. To preview a push, clone the
   project to a temp dir and `diff` against local. Don't try `--dry-run`.
2. **Apps Script API must be ON.** If you see `User has not enabled the Apps
   Script API`, the user must toggle it at
   <https://script.google.com/home/usersettings>. You can't fix this for them.
3. **`clasp push` is destructive.** It syncs local → remote — files present
   remotely but not locally get **deleted**. Always `pull` (or clone-to-temp
   diff) before a first push if there's any doubt local is in sync.
4. **`clasp list-scripts` only finds standalone projects.** Both projects here
   are container-bound (to spreadsheets), so they don't appear. Script IDs are
   in the committed `.clasp.json` files — read those.
5. **File extensions.** Local uses `.gs`, clasp converts to `.js` on remote
   transit (and back). Treat them as interchangeable.
6. **Auth lives in `~/.clasprc.json`.** Not the repo. If `clasp` reports
   "you're not logged in", run `npx --yes @google/clasp@latest login`.

## Secrets and PII

**All secrets and personal config live in Apps Script Script Properties** —
never in source. See `README.md` for the full list of required keys per
project. Code reads them via:

```javascript
const FOO = PropertiesService.getScriptProperties().getProperty('FOO');
```

If you add a new secret or piece of personal config:

- Read it via Script Properties.
- Add the key + a one-line description to the table in `README.md`.
- Tell the user the value to set, don't commit a default with the value baked in.

**Pre-commit secret scan.** Before the user pushes anything sensitive, run a
quick grep for the historical secret prefixes/IDs that have already been
exposed (`AIza`, `GOCSPX`, `gho_`, the user's email, etc.) — they should
never reappear in source.

## Don't

- Reintroduce hardcoded secrets or personal data, even temporarily. Goals,
  thresholds, names — derive from `ATHLETE_PROFILE` (which lives in Script
  Properties).
- `clasp push` without first verifying that local matches remote (or that the
  diff is exactly what was intended) — you can clobber remote work.
- Add `clasp` as a global npm install — `npx` is intentional for supply-chain
  reasons. Use the existing `npx --yes @google/clasp@latest` invocation.
