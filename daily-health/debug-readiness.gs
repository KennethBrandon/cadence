// =============================================================================
// DEBUG READINESS — inspect the readiness column and what's feeding it.
//
// Run debugReadiness() from the Apps Script editor. It logs:
//   1. Current column 17 header + recent values (so we can see if respRate
//      values like 11.8 are landing where readiness should live).
//   2. Per-day inputs (HRV/RHR/sleep/deep/SpO2) for the last N days.
//   3. The component scores readiness-score.gs would produce for each day.
//   4. Raw Google Health API responses for one date (the most useful one
//      to inspect when debugging an upstream change).
//
// Read-only — does not write to the sheet.
// =============================================================================

const DEBUG_DAYS_TO_INSPECT = 7;

function debugReadiness() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Daily Data');
  if (!sheet || sheet.getLastRow() < 2) {
    console.log('No Daily Data sheet found.');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  console.log('=== SHEET LAYOUT ===');
  headers.forEach((h, i) => console.log(`  col ${i + 1} (idx ${i}): ${h}`));

  const col17Header = headers[16];
  console.log(`\nColumn 17 header is: "${col17Header}"`);
  if (col17Header !== 'Readiness') {
    console.log(`  ⚠ Expected "Readiness". If runDailyHealthSync is overwriting`);
    console.log(`    this column with respRate, the header may have been replaced.`);
  }

  // -------------------------------------------------------------------------
  // 1. Recent values in col 17 — are they integers (readiness) or floats
  //    around 11–18 (respiratory rate)?
  // -------------------------------------------------------------------------
  const startRow = Math.max(2, lastRow - DEBUG_DAYS_TO_INSPECT + 1);
  const recentRows = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

  console.log(`\n=== LAST ${recentRows.length} DAYS — column 17 values ===`);
  recentRows.forEach(r => {
    const date = r[0] instanceof Date
      ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(r[0]).substring(0, 10);
    const v = r[16];
    const looksLikeRespRate = typeof v === 'number' && v > 8 && v < 25 && v !== Math.round(v);
    const looksLikeReadiness = typeof v === 'number' && v >= 1 && v <= 100 && v === Math.round(v);
    let tag = '';
    if (looksLikeRespRate)   tag = '  ← looks like respiratory rate (float, 8–25)';
    else if (looksLikeReadiness) tag = '  ← looks like readiness (integer, 1–100)';
    console.log(`  ${date}: ${v}${tag}`);
  });

  // -------------------------------------------------------------------------
  // 2. Recompute readiness components for the last N days from raw inputs
  //    in the sheet — same logic as computeReadinessScores, but verbose.
  // -------------------------------------------------------------------------
  console.log(`\n=== READINESS RECOMPUTE (from sheet inputs) ===`);

  // We need the BASELINE_WINDOW prior days too for rolling averages.
  const baselineStart = Math.max(2, startRow - 14);
  const baselineRows = sheet.getRange(baselineStart, 1, lastRow - baselineStart + 1, 14).getValues();
  // Indices into baselineRows (0-based here): RHR=2, HRV=3, TotalSleep=7, Deep=9, SpO2=13
  const parsed = baselineRows.map(r => ({
    date: r[0] instanceof Date
      ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(r[0]).substring(0, 10),
    rhr:    _toNumDbg(r[2]),
    hrv:    _toNumDbg(r[3]),
    sleep:  _toNumDbg(r[7]),
    deep:   _toNumDbg(r[9]),
    spo2:   _toNumDbg(r[13])
  }));

  // First inspect-window day in `parsed` is at index (startRow - baselineStart).
  const firstInspectIdx = startRow - baselineStart;
  for (let i = firstInspectIdx; i < parsed.length; i++) {
    const r = parsed[i];
    const window = parsed.slice(Math.max(0, i - 14), i);
    const avgHRV  = _avgFieldDbg(window, 'hrv');
    const avgRHR  = _avgFieldDbg(window, 'rhr');
    const avgSpO2 = _avgFieldDbg(window, 'spo2');

    let hrvScore = null, rhrScore = null, sleepScore = null, deepScore = null, spo2Score = null;
    if (r.hrv && avgHRV) hrvScore = 50 + ((r.hrv - avgHRV) / 10) * 50;
    else if (r.hrv) hrvScore = 50;
    if (r.rhr && avgRHR) rhrScore = 50 + ((avgRHR - r.rhr) / 5) * 50;
    else if (r.rhr) rhrScore = 50;
    if (r.sleep) sleepScore = Math.min(100, (r.sleep / 450) * 80);
    if (r.deep && r.sleep) deepScore = Math.min(100, (r.deep / r.sleep) * 100 * 4.5);
    if (r.spo2 && avgSpO2) spo2Score = 80 + ((r.spo2 - avgSpO2) / 1) * 40;
    else if (r.spo2) spo2Score = Math.min(100, (r.spo2 - 88) * 12.5);

    let weighted = 0, totalW = 0;
    if (hrvScore   != null) { weighted += hrvScore   * 0.30; totalW += 0.30; }
    if (rhrScore   != null) { weighted += rhrScore   * 0.25; totalW += 0.25; }
    if (sleepScore != null) { weighted += sleepScore * 0.25; totalW += 0.25; }
    if (deepScore  != null) { weighted += deepScore  * 0.10; totalW += 0.10; }
    if (spo2Score  != null) { weighted += spo2Score  * 0.10; totalW += 0.10; }

    const finalScore = totalW > 0
      ? Math.max(1, Math.min(100, Math.round(weighted / totalW)))
      : null;

    console.log(`\n  ${r.date}`);
    console.log(`    inputs:    hrv=${r.hrv}  rhr=${r.rhr}  sleep=${r.sleep}  deep=${r.deep}  spo2=${r.spo2}`);
    console.log(`    baselines: avgHRV=${_fmtDbg(avgHRV)}  avgRHR=${_fmtDbg(avgRHR)}  avgSpO2=${_fmtDbg(avgSpO2)}`);
    console.log(`    components: hrv=${_fmtDbg(hrvScore)}  rhr=${_fmtDbg(rhrScore)}  sleep=${_fmtDbg(sleepScore)}  deep=${_fmtDbg(deepScore)}  spo2=${_fmtDbg(spo2Score)}`);
    console.log(`    → readiness: ${finalScore}`);
  }

  // -------------------------------------------------------------------------
  // 3. Raw Google Health API dump for the most recent date
  // -------------------------------------------------------------------------
  console.log(`\n=== RAW GOOGLE HEALTH API DUMP (most recent date) ===`);
  if (typeof getHealthService !== 'function') {
    console.log('  getHealthService() not available in this project.');
    return;
  }
  const svc = getHealthService();
  if (!svc.hasAccess()) {
    console.log('  Not authorized. Run authorize() first.');
    return;
  }

  const headersAuth = {
    'Authorization': 'Bearer ' + svc.getAccessToken(),
    'Content-Type': 'application/json'
  };

  const lastDate = parsed[parsed.length - 1].date;
  const [yr, mo, dy] = lastDate.split('-').map(Number);
  const nextDay = Utilities.formatDate(new Date(yr, mo - 1, dy + 1), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const startZ = `${lastDate}T00:00:00Z`;
  const endZ = `${nextDay}T00:00:00Z`;

  console.log(`  Date: ${lastDate}`);

  _dumpEndpointDbg('daily-resting-heart-rate',
    `daily_resting_heart_rate.date>="${lastDate}" AND daily_resting_heart_rate.date<"${nextDay}"`, headersAuth);
  _dumpEndpointDbg('daily-heart-rate-variability',
    `daily_heart_rate_variability.date>="${lastDate}" AND daily_heart_rate_variability.date<"${nextDay}"`, headersAuth);
  _dumpEndpointDbg('oxygen-saturation',
    `oxygen_saturation.sample_time.civil_time>="${lastDate}" AND oxygen_saturation.sample_time.civil_time<"${nextDay}"`, headersAuth);
  _dumpEndpointDbg('sleep',
    `sleep.interval.civil_end_time>="${lastDate}" AND sleep.interval.civil_end_time<"${nextDay}"`, headersAuth);
  _dumpEndpointDbg('daily-respiratory-rate',
    `daily_respiratory_rate.date>="${lastDate}" AND daily_respiratory_rate.date<"${nextDay}"`, headersAuth);
}

// ---------- helpers (suffixed to avoid collisions with readiness-score.gs) ----------
function _toNumDbg(v) {
  if (v === '' || v == null) return null;
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}

function _avgFieldDbg(rows, field) {
  const vals = rows.map(r => r[field]).filter(v => v !== null);
  if (vals.length === 0) return null;
  return vals.reduce((a, b) => a + b, 0) / vals.length;
}

function _fmtDbg(v) {
  if (v == null) return 'null';
  return typeof v === 'number' ? v.toFixed(2) : String(v);
}

function _dumpEndpointDbg(dataType, filter, headers) {
  const url = `https://health.googleapis.com/v4/users/me/dataTypes/${dataType}/dataPoints?filter=${encodeURIComponent(filter)}`;
  const res = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
  console.log(`\n  --- ${dataType} (HTTP ${res.getResponseCode()}) ---`);
  const text = res.getContentText();
  console.log(text.length > 4000 ? text.substring(0, 4000) + '… [truncated]' : text);
}
