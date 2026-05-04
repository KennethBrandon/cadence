## Summary

`GET /v4/users/me/dataTypes/sleep/dataPoints` returns HTTP 500 "Internal error encountered" for every variant of the query, including the unfiltered request. Other dataTypes (`daily-resting-heart-rate`, `daily-heart-rate-variability`, `weight`, `steps`, `oxygen-saturation`, `total-calories`) return 200 normally with the same auth token.

**Onset:** ~2026-04-20. Reproducible 100% as of 2026-04-26 09:55 PT.

## Endpoint

```
GET https://health.googleapis.com/v4/users/me/dataTypes/sleep/dataPoints
```

Auth: OAuth2 bearer token, scope `https://www.googleapis.com/auth/googlehealth.sleep.readonly` (re-authorized today; other scopes on the same token work fine).

## Probe matrix

All probes run within ~2 seconds of each other:

| Filter | Result |
|---|---|
| *(no filter)* | 500 INTERNAL |
| `sleep.interval.civil_end_time>="2026-04-16"` | 500 INTERNAL |
| `sleep.interval.civil_end_time>="2026-04-21" AND <"2026-04-26"` | 500 INTERNAL |
| `sleep.interval.civil_end_time>="2026-04-25" AND <"2026-04-26"` | 500 INTERNAL |
| `sleep.interval.civil_end_time>="2026-03-22" AND <"2026-03-27"` | 500 INTERNAL *(control, ~5 weeks old)* |
| `sleep.interval.end_time.physical_time>="2026-04-16T00:00:00Z"` | 400 `INVALID_DATA_POINT_FILTER_DATA_TYPE_MEMBER` |
| `sleep.interval.civil_start_time>="2026-04-16"` | 400 `INVALID_DATA_POINT_FILTER_DATA_TYPE_MEMBER` |

The 400s confirm `civil_end_time` is still the correct filter field; the 500s are not filter-shape issues. Even the unfiltered request 500s, so this isn't a date-window problem either.

## Sample raw response

```json
HTTP/1.1 500
{
  "error": {
    "code": 500,
    "message": "Internal error encountered.",
    "status": "INTERNAL"
  }
}
```

## Expected

200 with `dataPoints` for nights of 2026-04-19 through 2026-04-25 (confirmed present and complete in the source app — Fitbit).

## Scope notes

- Other endpoints on the same OAuth credential are healthy: `daily-resting-heart-rate`, `daily-heart-rate-variability`, `weight`, `steps`, `oxygen-saturation`, `total-calories` all return 200 with data for 2026-04-19 through 2026-04-25.
- Sleep data was being returned successfully by this endpoint through 2026-04-18; first 500 observed for queries covering ≥2026-04-19.
- Source-of-truth (Fitbit app) shows full sleep sessions with stages for the affected nights.

## Environment

- Client: Google Apps Script (`UrlFetchApp`), but reproducible via `curl`.
- Region: US.
- API version: v4.
