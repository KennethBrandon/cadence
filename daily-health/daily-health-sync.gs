// Secrets are stored in Script Properties (Project Settings → Script Properties).
// Required keys: GOOGLE_HEALTH_CLIENT_ID, GOOGLE_HEALTH_CLIENT_SECRET.
const CLIENT_ID = PropertiesService.getScriptProperties().getProperty('GOOGLE_HEALTH_CLIENT_ID');
const CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('GOOGLE_HEALTH_CLIENT_SECRET');


// =============================================================================
// OAUTH
// =============================================================================
function getHealthService() {
  return OAuth2.createService('GoogleHealth')
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('https://www.googleapis.com/auth/googlehealth.health_metrics_and_measurements.readonly https://www.googleapis.com/auth/googlehealth.activity_and_fitness.readonly https://www.googleapis.com/auth/googlehealth.sleep.readonly')
    .setParam('access_type', 'offline')
    .setParam('prompt', 'consent');
}

function authCallback(request) {
  const healthService = getHealthService();
  const isAuthorized = healthService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab.');
  }
}

function authorize() {
  const healthService = getHealthService();
  healthService.reset();
  const authorizationUrl = healthService.getAuthorizationUrl();
  Logger.log('Open the following URL and grant the new permissions:');
  Logger.log(authorizationUrl);
}


// =============================================================================
// TRIGGER INSTALL — run once to enable hourly sync
// =============================================================================
function INSTALL_HEALTH_TRIGGER() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'runDailyHealthSync') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyHealthSync')
    .timeBased()
    .everyHours(1)
    .create();
  console.log("Health trigger installed — will run every hour.");
}


// =============================================================================
// MAIN — runs every hour
// Syncs yesterday fully (including sleep) and today's rolling metrics
// =============================================================================
function runDailyHealthSync() {
  const healthService = getHealthService();
  if (!healthService.hasAccess()) {
    console.error('Not authorized. Run authorize() first.');
    return;
  }
  console.log("Starting Health Sync...");

  const headers = {
    'Authorization': 'Bearer ' + healthService.getAccessToken(),
    'Content-Type': 'application/json'
  };

  const tz        = Session.getScriptTimeZone();
  const today     = new Date();
  const yesterday = new Date();
  yesterday.setDate(today.getDate() - 1);

  const todayStr     = Utilities.formatDate(today,     tz, 'yyyy-MM-dd');
  const yesterdayStr = Utilities.formatDate(yesterday, tz, 'yyyy-MM-dd');

  syncHealthForDate(yesterdayStr, true, headers); // full sync including sleep
  syncHealthForDate(todayStr,     true, headers); // rolling metrics + last night's sleep

  console.log("Health Sync Complete!");

  try { if (typeof runDailyWorkoutSync === "function") runDailyWorkoutSync(); } catch(e) {}
}


// =============================================================================
// SYNC A SINGLE DATE
// includeSleep = true for yesterday, false for today
// =============================================================================
function syncHealthForDate(dateStr, includeSleep, headers) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Daily Data');

  const tz         = Session.getScriptTimeZone();
  // Parse date components directly to avoid UTC/local timezone shifting
  const [yr, mo, dy] = dateStr.split('-').map(Number);
  const nextDayStr      = Utilities.formatDate(new Date(yr, mo-1, dy+1), tz, 'yyyy-MM-dd');
  const dayAfterNextStr = Utilities.formatDate(new Date(yr, mo-1, dy+2), tz, 'yyyy-MM-dd');
  const startZ     = `${dateStr}T00:00:00Z`;
  const endZ       = `${nextDayStr}T00:00:00Z`;

  function fetchHealthAPI(dataType, filterString) {
    const url = `https://health.googleapis.com/v4/users/me/dataTypes/${dataType}/dataPoints?filter=${encodeURIComponent(filterString)}`;
    const res = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    return JSON.parse(res.getContentText());
  }

  function fetchDailyRollup(dataType) {
    const url = `https://health.googleapis.com/v4/users/me/dataTypes/${dataType}/dataPoints:dailyRollUp`;
    const [sy, sm, sd] = dateStr.split('-').map(Number);
    // nextDayStr is always dateStr+1 in local time, so end is always strictly greater than start
    const [ey, em, ed] = nextDayStr.split('-').map(Number);
    const payload = { range: { start: { date: { year: sy, month: sm, day: sd } }, end: { date: { year: ey, month: em, day: ed } } } };
    const res = UrlFetchApp.fetch(url, { method: 'post', headers, payload: JSON.stringify(payload), muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      console.log(`Rollup error for ${dataType}: HTTP ${res.getResponseCode()} — ${res.getContentText()}`);
      return null;
    }
    const result = JSON.parse(res.getContentText());
    return result.rollupDataPoints?.[0] ?? null;
  }

  const date = new Date(`${dateStr}T12:00:00Z`);

  // 1. Resting Heart Rate
  const hrData = fetchHealthAPI("daily-resting-heart-rate", `daily_resting_heart_rate.date>="${dateStr}"`);
  const rhr = hrData?.dataPoints?.find(p => p.dailyRestingHeartRate.date.day === date.getUTCDate())
    ?.dailyRestingHeartRate.beatsPerMinute || "";

  // 2. HRV
  const hrvData = fetchHealthAPI("daily-heart-rate-variability", `daily_heart_rate_variability.date>="${dateStr}"`);
  let avgHrv = "", nonRemHr = "", entropy = "", deepSleepRmssd = "";
  const hrvPoint = hrvData?.dataPoints?.find(p => p.dailyHeartRateVariability.date.day === date.getUTCDate());
  if (hrvPoint) {
    const d      = hrvPoint.dailyHeartRateVariability;
    avgHrv       = d.averageHeartRateVariabilityMilliseconds;
    nonRemHr     = d.nonRemHeartRateBeatsPerMinute;
    entropy      = d.entropy;
    deepSleepRmssd = d.deepSleepRootMeanSquareOfSuccessiveDifferencesMilliseconds;
  }

  // 3. Weight
  const weightData = fetchHealthAPI("weight", `weight.sample_time.physical_time>="${startZ}" AND weight.sample_time.physical_time<"${endZ}"`);
  const weightLbs  = weightData?.dataPoints?.[0]?.weight?.weightGrams
    ? (weightData.dataPoints[0].weight.weightGrams / 1000 * 2.20462).toFixed(1) : "";

  // 4. Sleep (yesterday only)
  let totalSleep = "", lightM = "", deepM = "", remM = "", awakeM = "";
  if (includeSleep) {
    const sleepData = fetchHealthAPI("sleep", `sleep.interval.civil_end_time>="${dateStr}"`);
    if (sleepData?.dataPoints) {
      const tz          = Session.getScriptTimeZone();
      const targetSleep = sleepData.dataPoints.find(p =>
        p.sleep?.interval?.endTime &&
        Utilities.formatDate(new Date(p.sleep.interval.endTime), tz, 'yyyy-MM-dd') === dateStr
      );
      if (targetSleep) {
        let l=0, d=0, r=0, a=0;
        if (targetSleep.sleep.stages) {
          targetSleep.sleep.stages.forEach(stage => {
            const mins = (new Date(stage.endTime) - new Date(stage.startTime)) / 60000;
            if (["LIGHT","ASLEEP"].includes(stage.type))   l += mins;
            else if (stage.type === "DEEP")                d += mins;
            else if (stage.type === "REM")                 r += mins;
            else if (["AWAKE","RESTLESS"].includes(stage.type)) a += mins;
          });
        } else {
          l = (new Date(targetSleep.sleep.interval.endTime) - new Date(targetSleep.sleep.interval.startTime)) / 60000;
        }
        totalSleep = Math.round(l+d+r);
        lightM     = Math.round(l);
        deepM      = Math.round(d);
        remM       = Math.round(r);
        awakeM     = Math.round(a);
      }
    }
  }

  // 5. SpO2
  let spo2 = "";
  const spo2Data = fetchHealthAPI("oxygen-saturation",
    `oxygen_saturation.sample_time.physical_time>="${startZ}" AND oxygen_saturation.sample_time.physical_time<"${endZ}"`);
  if (spo2Data?.dataPoints?.length > 0) {
    let sum = 0, count = 0;
    spo2Data.dataPoints.forEach(p => {
      const pct = p.oxygenSaturation?.percentage;
      if (pct && pct > 75) { sum += pct; count++; }
    });
    if (count > 0) spo2 = (sum / count).toFixed(1);
  }

  // 6. Rollups (Steps, Distance, Calories)
  let totalSteps = "", totalDistMi = "", totalCals = "";

  const stepsRollup = fetchDailyRollup("steps");
  if (stepsRollup?.steps?.countSum) totalSteps = stepsRollup.steps.countSum;

  const distRollup = fetchDailyRollup("distance");
  if (distRollup?.distance?.millimetersSum)
    totalDistMi = (distRollup.distance.millimetersSum / 1609344).toFixed(2);

  const calsRollup = fetchDailyRollup("total-calories");
  if (calsRollup?.totalCalories?.kcalSum) totalCals = Math.round(calsRollup.totalCalories.kcalSum);

  // 7. Respiratory Rate (daily)
  let respRate = "";
  const respData = fetchHealthAPI("daily-respiratory-rate",
    `daily_respiratory_rate.date>="${dateStr}" AND daily_respiratory_rate.date<"${nextDayStr}"`);
  if (respData?.dataPoints?.[0]?.dailyRespiratoryRate?.breathsPerMinute != null) {
    respRate = respData.dataPoints[0].dailyRespiratoryRate.breathsPerMinute;
  }

  // 8. Skin Temp Δ + Sleep Tracked? (Y/N if endpoint succeeds, blank if it fails)
  let tempDelta = "", sleepTracked = "";
  const tempData = fetchHealthAPI("daily-sleep-temperature-derivations",
    `daily_sleep_temperature_derivations.date>="${dateStr}" AND daily_sleep_temperature_derivations.date<"${nextDayStr}"`);
  if (tempData != null) {
    const x = tempData.dataPoints?.[0]?.dailySleepTemperatureDerivations;
    if (x?.nightlyTemperatureCelsius != null && x?.baselineTemperatureCelsius != null) {
      tempDelta = +(x.nightlyTemperatureCelsius - x.baselineTemperatureCelsius).toFixed(3);
    }
    sleepTracked = x ? 'Y' : 'N';
  }

  // 9. VO2 Max (daily, interpolated)
  let vo2Daily = "";
  const vo2DailyData = fetchHealthAPI("daily-vo2-max",
    `daily_vo2_max.date>="${dateStr}" AND daily_vo2_max.date<"${nextDayStr}"`);
  if (vo2DailyData?.dataPoints?.[0]?.dailyVo2Max?.vo2Max != null) {
    vo2Daily = +vo2DailyData.dataPoints[0].dailyVo2Max.vo2Max.toFixed(2);
  }

  // 10. VO2 Max (measured) — highest run-vo2 for the day
  let vo2Run = "";
  const vo2RunData = fetchHealthAPI("run-vo2-max",
    `run_vo2_max.sample_time.physical_time>="${startZ}" AND run_vo2_max.sample_time.physical_time<"${endZ}"`);
  if (vo2RunData?.dataPoints?.length > 0) {
    let max = -Infinity;
    vo2RunData.dataPoints.forEach(p => {
      const v = p.runVo2Max?.runVo2Max;
      if (v != null && v > max) max = v;
    });
    if (max > -Infinity) vo2Run = +max.toFixed(2);
  }

  // 11. Active Zone Minutes (rollup) — fat burn × 1 + cardio × 2 + peak × 2
  let azm = "";
  const azmRollup = fetchDailyRollup("active-zone-minutes");
  if (azmRollup?.activeZoneMinutes) {
    const a = azmRollup.activeZoneMinutes;
    const fb = parseInt(a.sumInFatBurnHeartZone || '0', 10);
    const cd = parseInt(a.sumInCardioHeartZone   || '0', 10);
    const pk = parseInt(a.sumInPeakHeartZone     || '0', 10);
    azm = fb + 2 * (cd + pk);
  }

  // 12. Body Fat
  let bodyFat = "";
  const bodyFatData = fetchHealthAPI("body-fat",
    `body_fat.sample_time.physical_time>="${startZ}" AND body_fat.sample_time.physical_time<"${endZ}"`);
  if (bodyFatData?.dataPoints?.[0]?.bodyFat?.percentage != null) {
    bodyFat = bodyFatData.dataPoints[0].bodyFat.percentage;
  }

  // 13. Upsert to sheet (23 columns). For cols 17–23, preserve existing values
  // when a fetch returned empty so intermittent API failures don't wipe data.
  const rowNum = getOrCreateRowForDate(sheet, dateStr);
  const existingNew = sheet.getRange(rowNum, 17, 1, 7).getValues()[0];
  const newCols = [respRate, tempDelta, vo2Daily, vo2Run, azm, bodyFat, sleepTracked]
    .map((v, i) => v !== "" ? v : existingNew[i]);
  sheet.getRange(rowNum, 1, 1, 23).setValues([[
    dateStr, weightLbs, rhr, avgHrv, nonRemHr, entropy, deepSleepRmssd,
    totalSleep, lightM, deepM, remM, awakeM,
    totalSteps || "", spo2 || "", totalDistMi || "", totalCals || "",
    ...newCols
  ]]);

  console.log(`Synced ${dateStr} — Steps: ${totalSteps}, Cals: ${totalCals}, Sleep: ${totalSleep ? totalSleep + ' min' : 'none found'}, Resp: ${respRate}, AZM: ${azm}`);
}


// =============================================================================
// YEARLY BACKFILL
// Change BACKFILL_YEAR as needed. Safe to re-run — resumes from last row.
// =============================================================================
function runYearlyBackfill() {
  const BACKFILL_YEAR = 2025;

  const healthService = getHealthService();
  if (!healthService.hasAccess()) { console.error('Not authorized.'); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Daily Data');
  if (!sheet) {
    sheet = ss.insertSheet('Daily Data');
    sheet.appendRow(['Date', 'Weight', 'RHR', 'HRV', 'Non-REM HR', 'Entropy', 'RMSSD',
      'Total Sleep', 'Light', 'Deep', 'REM', 'Awake', 'Steps', 'SpO2', 'Distance (mi)', 'Calories']);
  }

  // Resume from last row
  let lastDateStr = "";
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const lastVal = sheet.getRange(lastRow, 1).getValue();
    if (lastVal) {
      lastDateStr = (lastVal instanceof Date)
        ? Utilities.formatDate(lastVal, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(lastVal).substring(0, 10);
    }
  }

  let currentChunkStart = new Date(`${BACKFILL_YEAR}-01-01T00:00:00Z`);
  if (lastDateStr && lastDateStr.startsWith(BACKFILL_YEAR.toString())) {
    currentChunkStart = new Date(`${lastDateStr}T00:00:00Z`);
    currentChunkStart.setDate(currentChunkStart.getDate() + 1);
    console.log(`Resuming from ${Utilities.formatDate(currentChunkStart, Session.getScriptTimeZone(), 'yyyy-MM-dd')}...`);
  } else {
    console.log(`Starting fresh backfill for ${BACKFILL_YEAR}...`);
  }

  const hardEnd = new Date(`${BACKFILL_YEAR + 1}-01-01T00:00:00Z`);
  if (currentChunkStart >= hardEnd) {
    console.log(`Backfill for ${BACKFILL_YEAR} is already complete!`);
    return;
  }

  const headers = {
    'Authorization': 'Bearer ' + healthService.getAccessToken(),
    'Content-Type': 'application/json'
  };

  function fetchDailyRollup(dataType, start, end) {
    const url = `https://health.googleapis.com/v4/users/me/dataTypes/${dataType}/dataPoints:dailyRollUp`;
    const [sy, sm, sd] = start.split('-').map(Number);
    const [ey, em, ed] = end.split('-').map(Number);
    const payload = { range: { start: { date: { year: sy, month: sm, day: sd } }, end: { date: { year: ey, month: em, day: ed } } } };
    const res = UrlFetchApp.fetch(url, { method: 'post', headers, payload: JSON.stringify(payload), muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return [];
    return JSON.parse(res.getContentText()).rollupDataPoints || [];
  }

  function fetchChunkedData(dataType, filterString, maxRetries = 3) {
    let allPoints = [], pageToken = "";
    do {
      let url = `https://health.googleapis.com/v4/users/me/dataTypes/${dataType}/dataPoints?filter=${encodeURIComponent(filterString)}`;
      if (pageToken) url += `&pageToken=${pageToken}`;
      let response, attempt = 0, success = false;
      while (attempt < maxRetries && !success) {
        response = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
        if (response.getResponseCode() === 200) {
          success = true;
        } else if (response.getResponseCode() >= 500) {
          attempt++;
          console.warn(`[Retry ${attempt}/${maxRetries}] 500 error for ${dataType}...`);
          Utilities.sleep(2000);
        } else {
          break;
        }
      }
      if (!success) break;
      const result = JSON.parse(response.getContentText());
      if (result.dataPoints) allPoints = allPoints.concat(result.dataPoints);
      pageToken = result.nextPageToken;
    } while (pageToken);
    return allPoints;
  }

  while (currentChunkStart < hardEnd) {
    let nextChunkStart = new Date(currentChunkStart);
    nextChunkStart.setDate(nextChunkStart.getDate() + 30);
    if (nextChunkStart > hardEnd) nextChunkStart = hardEnd;

    const tz           = Session.getScriptTimeZone();
    const chunkStartStr = Utilities.formatDate(currentChunkStart, tz, 'yyyy-MM-dd');
    const chunkEndStr   = Utilities.formatDate(nextChunkStart,   tz, 'yyyy-MM-dd');
    const chunkStartZ   = `${chunkStartStr}T00:00:00Z`;
    const chunkEndZ     = `${chunkEndStr}T00:00:00Z`;

    const rollupEnd    = new Date(nextChunkStart);
    rollupEnd.setDate(rollupEnd.getDate() - 1);
    const rollupEndStr = Utilities.formatDate(rollupEnd, tz, 'yyyy-MM-dd');

    console.log(`Fetching ${chunkStartStr} to ${rollupEndStr}...`);

    const dailyRecords = new Map();

    // Rollups
    ['steps', 'distance', 'total-calories'].forEach(metric => {
      fetchDailyRollup(metric, chunkStartStr, rollupEndStr).forEach(p => {
        if (!p.civilStartTime?.date) return;
        const d  = p.civilStartTime.date;
        const ds = `${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`;
        if (!dailyRecords.has(ds)) dailyRecords.set(ds, {});
        const r = dailyRecords.get(ds);
        if (p.steps)          r.steps = p.steps.countSum;
        if (p.distance)       r.dist  = (p.distance.millimetersSum / 1609344).toFixed(2);
        if (p.totalCalories?.kcalSum) r.calsSum = p.totalCalories.kcalSum;
      });
    });

    // RHR
    fetchChunkedData("daily-resting-heart-rate",
      `daily_resting_heart_rate.date>="${chunkStartStr}" AND daily_resting_heart_rate.date<"${chunkEndStr}"`
    ).forEach(p => {
      if (!p.dailyRestingHeartRate?.date) return;
      const d  = p.dailyRestingHeartRate.date;
      const ds = `${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`;
      if (!dailyRecords.has(ds)) dailyRecords.set(ds, {});
      dailyRecords.get(ds).rhr = p.dailyRestingHeartRate.beatsPerMinute;
    });

    // HRV
    fetchChunkedData("daily-heart-rate-variability",
      `daily_heart_rate_variability.date>="${chunkStartStr}" AND daily_heart_rate_variability.date<"${chunkEndStr}"`
    ).forEach(p => {
      if (!p.dailyHeartRateVariability?.date) return;
      const d  = p.dailyHeartRateVariability.date;
      const ds = `${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`;
      if (!dailyRecords.has(ds)) dailyRecords.set(ds, {});
      const r    = dailyRecords.get(ds);
      r.hrv      = p.dailyHeartRateVariability.averageHeartRateVariabilityMilliseconds;
      r.nonRemHr = p.dailyHeartRateVariability.nonRemHeartRateBeatsPerMinute;
      r.entropy  = p.dailyHeartRateVariability.entropy;
      r.rmssd    = p.dailyHeartRateVariability.deepSleepRootMeanSquareOfSuccessiveDifferencesMilliseconds;
    });

    // Weight
    fetchChunkedData("weight",
      `weight.sample_time.physical_time>="${chunkStartZ}" AND weight.sample_time.physical_time<"${chunkEndZ}"`
    ).forEach(p => {
      if (!p.weight?.sampleTime?.civilTime?.date) return;
      const d  = p.weight.sampleTime.civilTime.date;
      const ds = `${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`;
      if (!dailyRecords.has(ds)) dailyRecords.set(ds, {});
      dailyRecords.get(ds).weight = (p.weight.weightGrams / 1000 * 2.20462).toFixed(1);
    });

    // SpO2
    fetchChunkedData("oxygen-saturation",
      `oxygen_saturation.sample_time.civil_time>="${chunkStartStr}" AND oxygen_saturation.sample_time.civil_time<"${chunkEndStr}"`
    ).forEach(p => {
      if (!p.oxygenSaturation?.sampleTime?.civilTime?.date) return;
      const d  = p.oxygenSaturation.sampleTime.civilTime.date;
      const ds = `${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`;
      if (!dailyRecords.has(ds)) dailyRecords.set(ds, {});
      dailyRecords.get(ds).spo2 = p.oxygenSaturation.percentage;
    });

    // Sleep
    fetchChunkedData("sleep",
      `sleep.interval.civil_end_time>="${chunkStartStr}" AND sleep.interval.civil_end_time<"${chunkEndStr}"`
    ).forEach(p => {
      if (!p.sleep?.interval?.endTime) return;
      const ds = Utilities.formatDate(new Date(p.sleep.interval.endTime), tz, "yyyy-MM-dd");
      if (ds < chunkStartStr || ds >= chunkEndStr) return;
      if (!dailyRecords.has(ds)) dailyRecords.set(ds, {});
      const r = dailyRecords.get(ds);
      let l=0, d=0, rem=0, a=0;
      if (p.sleep.stages) {
        p.sleep.stages.forEach(s => {
          const mins = (new Date(s.endTime) - new Date(s.startTime)) / 60000;
          if (["LIGHT","ASLEEP"].includes(s.type))        l   += mins;
          else if (s.type === "DEEP")                     d   += mins;
          else if (s.type === "REM")                      rem += mins;
          else if (["AWAKE","RESTLESS"].includes(s.type)) a   += mins;
        });
      } else {
        l = (new Date(p.sleep.interval.endTime) - new Date(p.sleep.interval.startTime)) / 60000;
      }
      r.totalSleep = Math.round(l+d+rem);
      r.light      = Math.round(l);
      r.deep       = Math.round(d);
      r.rem        = Math.round(rem);
      r.awake      = Math.round(a);
    });

    // Write chunk
    const rows = Array.from(dailyRecords.keys()).sort().map(ds => {
      const r = dailyRecords.get(ds);
      return [
        ds, r.weight || "", r.rhr || "", r.hrv || "", r.nonRemHr || "", r.entropy || "", r.rmssd || "",
        r.totalSleep || "", r.light || "", r.deep || "", r.rem || "", r.awake || "",
        r.steps || "", r.spo2 || "", r.dist || "", r.calsSum ? Math.round(r.calsSum) : ""
      ];
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      console.log(`Saved ${rows.length} days.`);
    }

    currentChunkStart = nextChunkStart;
  }
}


// =============================================================================
// BACKFILL: 6 new columns (Resp Rate, Skin Temp Δ, VO2 Max, VO2 Max (measured),
// Active Zone Min, Body Fat %) across every existing dated row. Adds headers
// if missing. Idempotent — only fills cells that are currently empty.
// =============================================================================
function backfillNewColumns() {
  const NEW_HEADERS = ['Resp Rate', 'Skin Temp Δ (°C)', 'VO2 Max', 'VO2 Max (measured)', 'Active Zone Min', 'Body Fat %', 'Sleep Tracked?'];
  const FIRST_NEW_COL = 17; // existing sheet has 16 columns

  const healthService = getHealthService();
  if (!healthService.hasAccess()) { console.error('Not authorized.'); return; }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Daily Data');
  if (!sheet) { console.error("Could not find 'Daily Data' sheet!"); return; }

  // headers rebuilt per chunk so the OAuth token doesn't expire mid-backfill
  let headers = { 'Authorization': 'Bearer ' + healthService.getAccessToken(), 'Content-Type': 'application/json' };
  const refreshHeaders = () => { headers = { 'Authorization': 'Bearer ' + healthService.getAccessToken(), 'Content-Type': 'application/json' }; };
  const tz = Session.getScriptTimeZone();

  // 1. ensure new headers
  const headerRow = sheet.getRange(1, 1, 1, FIRST_NEW_COL + NEW_HEADERS.length - 1).getValues()[0];
  NEW_HEADERS.forEach((h, i) => {
    if (headerRow[FIRST_NEW_COL - 1 + i] !== h) sheet.getRange(1, FIRST_NEW_COL + i).setValue(h);
  });

  // 2. index existing rows by date + find date range.
  // For chunk-skip logic, track which dense columns each date already has filled.
  // (Cols 20=run-vo2 and 22=body-fat are inherently sparse and don't gate.)
  const DENSE_COLS = [17, 18, 19, 21]; // Resp Rate, Skin Temp Δ, VO2 Max, AZM
  const data = sheet.getDataRange().getValues();
  const rowByDate = new Map();
  const denseFilled = new Map(); // dateStr -> Set<colNum>
  let minDate = null, maxDate = null;
  for (let i = 1; i < data.length; i++) {
    const v = data[i][0];
    if (!v) continue;
    const ds = (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).substring(0, 10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) continue;
    rowByDate.set(ds, i + 1);
    const filled = new Set();
    DENSE_COLS.forEach(c => { if (data[i][c - 1] !== '') filled.add(c); });
    denseFilled.set(ds, filled);
    if (!minDate || ds < minDate) minDate = ds;
    if (!maxDate || ds > maxDate) maxDate = ds;
  }
  if (!minDate) { console.log('No dated rows found.'); return; }
  console.log(`Backfilling ${rowByDate.size} rows from ${minDate} to ${maxDate}`);

  function fetchWithAuthRetry(url, options) {
    let res = UrlFetchApp.fetch(url, { ...options, headers, muteHttpExceptions: true });
    if (res.getResponseCode() === 401) {
      refreshHeaders();
      res = UrlFetchApp.fetch(url, { ...options, headers, muteHttpExceptions: true });
    }
    return res;
  }

  // Returns null on failure so callers can distinguish "endpoint failed"
  // from "endpoint succeeded with 0 points".
  function fetchAll(dataType, filter) {
    const out = [];
    let pageToken = '';
    do {
      let url = `https://health.googleapis.com/v4/users/me/dataTypes/${dataType}/dataPoints?filter=${encodeURIComponent(filter)}`;
      if (pageToken) url += `&pageToken=${pageToken}`;
      const res = fetchWithAuthRetry(url, { method: 'get' });
      if (res.getResponseCode() !== 200) {
        console.warn(`  ${dataType} HTTP ${res.getResponseCode()} — ${res.getContentText().replace(/\s+/g, ' ').substring(0, 200)}`);
        return null;
      }
      const j = JSON.parse(res.getContentText());
      if (j.dataPoints) out.push(...j.dataPoints);
      pageToken = j.nextPageToken || '';
    } while (pageToken);
    return out;
  }

  function fetchDailyRollup(dataType, sStr, eStr) {
    const [sy, sm, sd] = sStr.split('-').map(Number);
    const [ey, em, ed] = eStr.split('-').map(Number);
    const payload = { range: { start: { date: { year: sy, month: sm, day: sd } }, end: { date: { year: ey, month: em, day: ed } } } };
    const res = fetchWithAuthRetry(
      `https://health.googleapis.com/v4/users/me/dataTypes/${dataType}/dataPoints:dailyRollUp`,
      { method: 'post', payload: JSON.stringify(payload) }
    );
    if (res.getResponseCode() !== 200) {
      console.warn(`  ${dataType} rollup HTTP ${res.getResponseCode()} — ${res.getContentText().replace(/\s+/g, ' ').substring(0, 200)}`);
      return [];
    }
    return JSON.parse(res.getContentText()).rollupDataPoints || [];
  }

  function dateOf(d) { return `${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`; }

  // 3. iterate 30-day chunks
  let cursor = new Date(`${minDate}T00:00:00Z`);
  const hardEnd = new Date(`${maxDate}T00:00:00Z`);
  hardEnd.setDate(hardEnd.getDate() + 1);

  while (cursor < hardEnd) {
    refreshHeaders();
    const chunkEnd = new Date(cursor);
    chunkEnd.setDate(chunkEnd.getDate() + 30);
    if (chunkEnd > hardEnd) chunkEnd.setTime(hardEnd.getTime());
    const sStr = Utilities.formatDate(cursor,   tz, 'yyyy-MM-dd');
    const eStr = Utilities.formatDate(chunkEnd, tz, 'yyyy-MM-dd');
    const sZ = `${sStr}T00:00:00Z`, eZ = `${eStr}T00:00:00Z`;

    // Skip chunk if every dense column has at least one filled value across
    // the chunk's dates (proves all 4 endpoints succeeded last time).
    const endpointSeen = new Set();
    rowByDate.forEach((_, ds) => {
      if (ds < sStr || ds >= eStr) return;
      denseFilled.get(ds).forEach(c => endpointSeen.add(c));
    });
    if (DENSE_COLS.every(c => endpointSeen.has(c))) {
      console.log(`Chunk ${sStr} → ${eStr} — already done, skipping`);
      cursor = chunkEnd;
      continue;
    }
    const missing = DENSE_COLS.filter(c => !endpointSeen.has(c));
    console.log(`Chunk ${sStr} → ${eStr} — missing cols ${missing.join(',')}`);

    const records = new Map();
    const recFor = (ds) => { if (!records.has(ds)) records.set(ds, {}); return records.get(ds); };

    // Resp rate (daily)
    (fetchAll('daily-respiratory-rate',
      `daily_respiratory_rate.date>="${sStr}" AND daily_respiratory_rate.date<"${eStr}"`
    ) || []).forEach(p => {
      const x = p.dailyRespiratoryRate;
      if (x?.date) recFor(dateOf(x.date)).resp = x.breathsPerMinute;
    });

    // Skin temp delta + Sleep Tracked. If the endpoint succeeded, mark every
    // chunk-date with 'Y' or 'N' based on whether a point exists for it.
    const tempPoints = fetchAll('daily-sleep-temperature-derivations',
      `daily_sleep_temperature_derivations.date>="${sStr}" AND daily_sleep_temperature_derivations.date<"${eStr}"`);
    if (tempPoints !== null) {
      const tracked = new Set();
      tempPoints.forEach(p => {
        const x = p.dailySleepTemperatureDerivations;
        if (!x?.date) return;
        const ds = dateOf(x.date);
        tracked.add(ds);
        const r = recFor(ds);
        r.sleepTracked = 'Y';
        if (x.nightlyTemperatureCelsius != null && x.baselineTemperatureCelsius != null) {
          r.tempDelta = +(x.nightlyTemperatureCelsius - x.baselineTemperatureCelsius).toFixed(3);
        }
      });
      // Fill 'N' for every chunk-date the endpoint didn't return
      rowByDate.forEach((_, ds) => {
        if (ds >= sStr && ds < eStr && !tracked.has(ds)) recFor(ds).sleepTracked = 'N';
      });
    }

    // VO2 max (daily, interpolated)
    (fetchAll('daily-vo2-max',
      `daily_vo2_max.date>="${sStr}" AND daily_vo2_max.date<"${eStr}"`
    ) || []).forEach(p => {
      const x = p.dailyVo2Max;
      if (x?.date) recFor(dateOf(x.date)).vo2Daily = +x.vo2Max.toFixed(2);
    });

    // VO2 max (per-run, ground truth) — keep highest if multiple/day
    (fetchAll('run-vo2-max',
      `run_vo2_max.sample_time.physical_time>="${sZ}" AND run_vo2_max.sample_time.physical_time<"${eZ}"`
    ) || []).forEach(p => {
      const x = p.runVo2Max;
      const date = x?.sampleTime?.civilTime?.date;
      if (!date) return;
      const r = recFor(dateOf(date));
      const v = +x.runVo2Max.toFixed(2);
      if (r.vo2Run == null || v > r.vo2Run) r.vo2Run = v;
    });

    // Active Zone Minutes — daily rollup returns minutes per zone.
    // AZM credit: fat burn × 1, cardio × 2, peak × 2.
    fetchDailyRollup('active-zone-minutes', sStr, eStr).forEach(p => {
      const date = p.civilStartTime?.date;
      const azm = p.activeZoneMinutes;
      if (!date || !azm) return;
      const fb = parseInt(azm.sumInFatBurnHeartZone || '0', 10);
      const cd = parseInt(azm.sumInCardioHeartZone   || '0', 10);
      const pk = parseInt(azm.sumInPeakHeartZone     || '0', 10);
      recFor(dateOf(date)).azm = fb + 2 * (cd + pk);
    });

    // Body fat
    (fetchAll('body-fat',
      `body_fat.sample_time.physical_time>="${sZ}" AND body_fat.sample_time.physical_time<"${eZ}"`
    ) || []).forEach(p => {
      const x = p.bodyFat;
      const date = x?.sampleTime?.civilTime?.date;
      if (date && x.percentage != null) recFor(dateOf(date)).bodyFat = x.percentage;
    });

    // 4. write into existing rows (only fill empty cells)
    let writes = 0;
    records.forEach((r, ds) => {
      const rowNum = rowByDate.get(ds);
      if (!rowNum) return;
      const range = sheet.getRange(rowNum, FIRST_NEW_COL, 1, NEW_HEADERS.length);
      const existing = range.getValues()[0];
      const next = [
        existing[0] === '' && r.resp         != null ? r.resp         : existing[0],
        existing[1] === '' && r.tempDelta    != null ? r.tempDelta    : existing[1],
        existing[2] === '' && r.vo2Daily     != null ? r.vo2Daily     : existing[2],
        existing[3] === '' && r.vo2Run       != null ? r.vo2Run       : existing[3],
        existing[4] === '' && r.azm          != null ? r.azm          : existing[4],
        existing[5] === '' && r.bodyFat      != null ? r.bodyFat      : existing[5],
        existing[6] === '' && r.sleepTracked != null ? r.sleepTracked : existing[6],
      ];
      if (next.some((v, i) => v !== existing[i])) { range.setValues([next]); writes++; }
    });
    console.log(`  wrote ${writes} rows`);

    cursor = chunkEnd;
  }
  console.log('Backfill complete.');
}


// =============================================================================
// UNIVERSAL PATCHER
// Scans every row and fetches missing data for any metric
// =============================================================================
function patchAllMissingData() {
  const healthService = getHealthService();
  if (!healthService.hasAccess()) { console.error('Not authorized.'); return; }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Daily Data');
  if (!sheet) { console.error("Could not find 'Daily Data' sheet!"); return; }

  const data    = sheet.getDataRange().getValues();
  const headers = { 'Authorization': 'Bearer ' + healthService.getAccessToken(), 'Content-Type': 'application/json' };

  console.log("Scanning for rows with missing data...");
  let rowsUpdated = 0;

  for (let i = 1; i < data.length; i++) {
    const row     = data[i];
    const dateVal = row[0];
    if (!dateVal) continue;

    const dateStr    = (dateVal instanceof Date)
      ? Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(dateVal).substring(0, 10);
    const nextDay    = new Date(`${dateStr}T00:00:00Z`);
    nextDay.setDate(nextDay.getDate() + 1);
    const nextDayStr = Utilities.formatDate(nextDay, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const startZ     = `${dateStr}T00:00:00Z`;
    const endZ       = `${nextDayStr}T00:00:00Z`;
    const sleepNextDay = new Date(nextDay);
    sleepNextDay.setDate(sleepNextDay.getDate() + 1);
    const sleepNextDayStr = Utilities.formatDate(sleepNextDay, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const needsWeight = row[1]  === "";
    const needsRHR    = row[2]  === "";
    const needsHRV    = row[3]  === "" || row[4] === "" || row[5] === "" || row[6] === "";
    const needsSleep  = row[7]  === "" || row[8] === "" || row[9] === "" || row[10] === "" || row[11] === "";
    const needsSteps  = row[12] === "";
    const needsSpO2   = row[13] === "";
    const needsDist   = row[14] === "";
    const needsCals   = row[15] === "" || row[15] < 1500;

    if (!needsWeight && !needsRHR && !needsHRV && !needsSleep && !needsSteps && !needsSpO2 && !needsDist && !needsCals) continue;

    console.log(`Patching ${dateStr}...`);
    let didUpdate = false;

    function fetchGap(url, method = 'get', payload = null) {
      const options = { method, headers, muteHttpExceptions: true };
      if (payload) options.payload = JSON.stringify(payload);
      const res = UrlFetchApp.fetch(url, options);
      return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()) : null;
    }

    // Rollups
    if (needsSteps || needsDist || needsCals) {
      const [y, m, dNum]    = dateStr.split('-').map(Number);
      const [ey, em, edNum] = nextDayStr.split('-').map(Number);
      const payload = { range: { start: { date: { year: y, month: m, day: dNum } }, end: { date: { year: ey, month: em, day: edNum } } } };

      if (needsSteps) {
        const res = fetchGap(`https://health.googleapis.com/v4/users/me/dataTypes/steps/dataPoints:dailyRollUp`, 'post', payload);
        if (res?.rollupDataPoints?.[0]?.steps) {
          sheet.getRange(i + 1, 13).setValue(res.rollupDataPoints[0].steps.countSum);
          didUpdate = true;
        }
      }
      if (needsDist) {
        const res = fetchGap(`https://health.googleapis.com/v4/users/me/dataTypes/distance/dataPoints:dailyRollUp`, 'post', payload);
        if (res?.rollupDataPoints?.[0]?.distance) {
          sheet.getRange(i + 1, 15).setValue((res.rollupDataPoints[0].distance.millimetersSum / 1609344).toFixed(2));
          didUpdate = true;
        }
      }
      if (needsCals) {
        const res = fetchGap(`https://health.googleapis.com/v4/users/me/dataTypes/total-calories/dataPoints:dailyRollUp`, 'post', payload);
        if (res?.rollupDataPoints?.[0]?.totalCalories) {
          sheet.getRange(i + 1, 16).setValue(Math.round(res.rollupDataPoints[0].totalCalories.kcalSum));
          didUpdate = true;
        }
      }
    }

    // Weight
    if (needsWeight) {
      const res = fetchGap(`https://health.googleapis.com/v4/users/me/dataTypes/weight/dataPoints?filter=` +
        encodeURIComponent(`weight.sample_time.physical_time>="${startZ}" AND weight.sample_time.physical_time<"${endZ}"`));
      if (res?.dataPoints?.[0]?.weight?.weightGrams) {
        sheet.getRange(i + 1, 2).setValue((res.dataPoints[0].weight.weightGrams / 1000 * 2.20462).toFixed(1));
        didUpdate = true;
      }
    }

    // RHR
    if (needsRHR) {
      const res = fetchGap(`https://health.googleapis.com/v4/users/me/dataTypes/daily-resting-heart-rate/dataPoints?filter=` +
        encodeURIComponent(`daily_resting_heart_rate.date>="${dateStr}" AND daily_resting_heart_rate.date<"${nextDayStr}"`));
      if (res?.dataPoints?.[0]) {
        sheet.getRange(i + 1, 3).setValue(res.dataPoints[0].dailyRestingHeartRate.beatsPerMinute);
        didUpdate = true;
      }
    }

    // HRV
    if (needsHRV) {
      const res = fetchGap(`https://health.googleapis.com/v4/users/me/dataTypes/daily-heart-rate-variability/dataPoints?filter=` +
        encodeURIComponent(`daily_heart_rate_variability.date>="${dateStr}" AND daily_heart_rate_variability.date<"${nextDayStr}"`));
      if (res?.dataPoints?.[0]) {
        const hrv = res.dataPoints[0].dailyHeartRateVariability;
        sheet.getRange(i + 1, 4).setValue(hrv.averageHeartRateVariabilityMilliseconds || "");
        sheet.getRange(i + 1, 5).setValue(hrv.nonRemHeartRateBeatsPerMinute || "");
        sheet.getRange(i + 1, 6).setValue(hrv.entropy || "");
        sheet.getRange(i + 1, 7).setValue(hrv.deepSleepRootMeanSquareOfSuccessiveDifferencesMilliseconds || "");
        didUpdate = true;
      }
    }

    // SpO2
    if (needsSpO2) {
      const res = fetchGap(`https://health.googleapis.com/v4/users/me/dataTypes/oxygen-saturation/dataPoints?filter=` +
        encodeURIComponent(`oxygen_saturation.sample_time.physical_time>="${startZ}" AND oxygen_saturation.sample_time.physical_time<"${endZ}"`));
      if (res?.dataPoints?.length > 0) {
        let sum = 0, count = 0;
        res.dataPoints.forEach(p => {
          const pct = p.oxygenSaturation?.percentage;
          if (pct && pct > 75) { sum += pct; count++; }
        });
        if (count > 0) { sheet.getRange(i + 1, 14).setValue((sum / count).toFixed(1)); didUpdate = true; }
      }
    }

    // Sleep
    if (needsSleep) {
      const res = fetchGap(`https://health.googleapis.com/v4/users/me/dataTypes/sleep/dataPoints?filter=` +
        encodeURIComponent(`sleep.interval.civil_end_time>="${dateStr}" AND sleep.interval.civil_end_time<"${sleepNextDayStr}"`));
      if (res?.dataPoints) {
        const tz = Session.getScriptTimeZone();
        let l=0, d=0, rem=0, a=0;
        res.dataPoints.forEach(p => {
          if (!p.sleep?.interval?.endTime) return;
          if (Utilities.formatDate(new Date(p.sleep.interval.endTime), tz, "yyyy-MM-dd") !== dateStr) return;
          if (p.sleep.stages) {
            p.sleep.stages.forEach(s => {
              const mins = (new Date(s.endTime) - new Date(s.startTime)) / 60000;
              if (["LIGHT","ASLEEP"].includes(s.type))        l   += mins;
              else if (s.type === "DEEP")                     d   += mins;
              else if (s.type === "REM")                      rem += mins;
              else if (["AWAKE","RESTLESS"].includes(s.type)) a   += mins;
            });
          } else {
            l += (new Date(p.sleep.interval.endTime) - new Date(p.sleep.interval.startTime)) / 60000;
          }
        });
        const ts = Math.round(l+d+rem);
        if (ts > 0) {
          sheet.getRange(i + 1, 8).setValue(ts);
          sheet.getRange(i + 1, 9).setValue(Math.round(l));
          sheet.getRange(i + 1, 10).setValue(Math.round(d));
          sheet.getRange(i + 1, 11).setValue(Math.round(rem));
          sheet.getRange(i + 1, 12).setValue(Math.round(a));
          didUpdate = true;
        }
      }
    }

    if (didUpdate) rowsUpdated++;
  }

  console.log(`Done! Patched ${rowsUpdated} days with missing data.`);
}


// =============================================================================
// SPO2 PURIFIER
// Rescans all SpO2 data, filters sensor glitches, rewrites Column N
// =============================================================================
function purifySpO2Column() {
  const healthService = getHealthService();
  if (!healthService.hasAccess()) { console.error('Not authorized.'); return; }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Daily Data');
  if (!sheet) return;

  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = { 'Authorization': 'Bearer ' + healthService.getAccessToken(), 'Content-Type': 'application/json' };
  console.log("Starting SpO2 Purification...");

  const validDates = [];
  for (let i = 1; i < data.length; i++) {
    const d = data[i][0];
    if (d) validDates.push((d instanceof Date)
      ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(d).substring(0, 10));
  }
  if (validDates.length === 0) return;

  let currentChunkStart = new Date(`${validDates[0]}T00:00:00Z`);
  const hardEnd         = new Date(`${validDates[validDates.length - 1]}T00:00:00Z`);
  hardEnd.setDate(hardEnd.getDate() + 1);

  const dailySpO2 = new Map();

  while (currentChunkStart < hardEnd) {
    const nextChunkStart = new Date(currentChunkStart);
    nextChunkStart.setDate(nextChunkStart.getDate() + 30);
    if (nextChunkStart > hardEnd) nextChunkStart.setTime(hardEnd.getTime());

    const tz       = Session.getScriptTimeZone();
    const startStr = Utilities.formatDate(currentChunkStart, tz, 'yyyy-MM-dd');
    const endStr   = Utilities.formatDate(nextChunkStart,   tz, 'yyyy-MM-dd');
    console.log(`Fetching SpO2 ${startStr} to ${endStr}...`);

    let pageToken = "";
    do {
      const url = `https://health.googleapis.com/v4/users/me/dataTypes/oxygen-saturation/dataPoints?filter=` +
        encodeURIComponent(`oxygen_saturation.sample_time.civil_time>="${startStr}" AND oxygen_saturation.sample_time.civil_time<"${endStr}"`) +
        `&pageSize=5000` + (pageToken ? `&pageToken=${pageToken}` : '');
      const res = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        const result = JSON.parse(res.getContentText());
        (result.dataPoints || []).forEach(p => {
          if (!p.oxygenSaturation?.sampleTime?.civilTime?.date) return;
          const d   = p.oxygenSaturation.sampleTime.civilTime.date;
          const ds  = `${d.year}-${String(d.month).padStart(2,'0')}-${String(d.day).padStart(2,'0')}`;
          let pct   = p.oxygenSaturation.percentage;
          if (pct) {
            if (pct <= 1.0) pct = pct * 100;
            if (pct > 75) {
              if (!dailySpO2.has(ds)) dailySpO2.set(ds, { sum: 0, count: 0 });
              dailySpO2.get(ds).sum   += pct;
              dailySpO2.get(ds).count += 1;
            }
          }
        });
        pageToken = result.nextPageToken;
      } else {
        console.error(`API Error for ${startStr}:`, res.getContentText());
        break;
      }
    } while (pageToken);

    currentChunkStart = nextChunkStart;
  }

  const newSpO2Column = [];
  for (let i = 1; i < data.length; i++) {
    const d  = data[i][0];
    let newVal = "";
    if (d) {
      const ds = (d instanceof Date)
        ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(d).substring(0, 10);
      if (dailySpO2.has(ds)) {
        const rec = dailySpO2.get(ds);
        if (rec.count > 0) newVal = (rec.sum / rec.count).toFixed(1);
      }
    }
    newSpO2Column.push([newVal]);
  }

  if (newSpO2Column.length > 0) {
    sheet.getRange(2, 14, newSpO2Column.length, 1).setValues(newSpO2Column);
    console.log(`Done! Purified ${newSpO2Column.length} days of SpO2 data.`);
  }
}


// =============================================================================
// SPO2 DIAGNOSTICS
// =============================================================================
function diagnoseSpO2() {
  const healthService = getHealthService();
  if (!healthService.hasAccess()) { console.error('Not authorized.'); return; }

  const headers = { 'Authorization': 'Bearer ' + healthService.getAccessToken(), 'Content-Type': 'application/json' };
  const twoDaysAgo = new Date();
  twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);

  const url = `https://health.googleapis.com/v4/users/me/dataTypes/oxygen-saturation/dataPoints?filter=` +
    encodeURIComponent(`oxygen_saturation.sample_time.physical_time>="${twoDaysAgo.toISOString()}"`);
  const res = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });

  if (res.getResponseCode() === 200) {
    const data = JSON.parse(res.getContentText());
    if (data.dataPoints?.length > 0) {
      console.log("Raw SpO2 for last 48 hours:");
      console.log(JSON.stringify(data.dataPoints, null, 2));
    } else {
      console.log("API returned zero SpO2 points for the last 48 hours.");
    }
  } else {
    console.error("API Error:", res.getContentText());
  }
}


// =============================================================================
// SLEEP DIAGNOSTICS
// Logs raw sleep API response for the past 10 days, plus a per-day summary
// so you can see exactly which nights the API is/isn't returning.
// =============================================================================
function diagnoseSleep() {
  const healthService = getHealthService();
  if (!healthService.hasAccess()) { console.error('Not authorized.'); return; }

  const headers = { 'Authorization': 'Bearer ' + healthService.getAccessToken(), 'Content-Type': 'application/json' };
  const tz = Session.getScriptTimeZone();

  const todayStr     = Utilities.formatDate(new Date(),                                tz, 'yyyy-MM-dd');
  const yesterdayStr = Utilities.formatDate(new Date(Date.now() - 1*86400000),         tz, 'yyyy-MM-dd');
  const fiveAgoStr   = Utilities.formatDate(new Date(Date.now() - 5*86400000),         tz, 'yyyy-MM-dd');
  const tenAgoStr    = Utilities.formatDate(new Date(Date.now() - 10*86400000),        tz, 'yyyy-MM-dd');
  const monthAgoStr  = Utilities.formatDate(new Date(Date.now() - 35*86400000),        tz, 'yyyy-MM-dd');
  const monthBeforeStr = Utilities.formatDate(new Date(Date.now() - 30*86400000),      tz, 'yyyy-MM-dd');

  // Probes — each tries a different way to see if anything works.
  const probes = [
    ['no filter (any sleep)',                ''],
    ['civil_end_time >= 10 days ago',        `sleep.interval.civil_end_time>="${tenAgoStr}"`],
    ['civil_end_time 5d→today',              `sleep.interval.civil_end_time>="${fiveAgoStr}" AND sleep.interval.civil_end_time<"${todayStr}"`],
    ['civil_end_time yesterday only',        `sleep.interval.civil_end_time>="${yesterdayStr}" AND sleep.interval.civil_end_time<"${todayStr}"`],
    ['civil_end_time 35d→30d ago (control)', `sleep.interval.civil_end_time>="${monthAgoStr}" AND sleep.interval.civil_end_time<"${monthBeforeStr}"`],
    ['physical end_time recent',             `sleep.interval.end_time.physical_time>="${tenAgoStr}T00:00:00Z"`],
    ['civil_start_time recent',              `sleep.interval.civil_start_time>="${tenAgoStr}"`],
  ];

  probes.forEach(([label, filter]) => {
    const url = `https://health.googleapis.com/v4/users/me/dataTypes/sleep/dataPoints` +
      (filter ? `?filter=${encodeURIComponent(filter)}` : '');
    const res = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code === 200) {
      const data = JSON.parse(res.getContentText());
      const n = (data.dataPoints || []).length;
      console.log(`✅ [${label}] HTTP 200 — ${n} point(s)`);
      if (n > 0) {
        // Show one sample so we can see if the schema changed
        console.log(`   sample: ${JSON.stringify(data.dataPoints[0]).substring(0, 400)}...`);
      }
    } else {
      const body = res.getContentText().substring(0, 200);
      console.log(`❌ [${label}] HTTP ${code} — ${body}`);
    }
  });
}


// =============================================================================
// DIAGNOSTIC: probe which dataTypes the API exposes for this account.
// Tries the discovery endpoint, then walks a candidate list (known-good types
// as controls + plausible new types matching Fitbit's expanded sleep metrics).
// Each probe is unfiltered to isolate "type exists" from "filter shape".
// =============================================================================
function diagnoseDataTypes() {
  const healthService = getHealthService();
  if (!healthService.hasAccess()) { console.error('Not authorized.'); return; }

  const headers = { 'Authorization': 'Bearer ' + healthService.getAccessToken(), 'Content-Type': 'application/json' };

  // 1. Discovery — does the API expose a list endpoint?
  const discoveryUrl = 'https://health.googleapis.com/v4/users/me/dataTypes';
  const dRes = UrlFetchApp.fetch(discoveryUrl, { method: 'get', headers, muteHttpExceptions: true });
  console.log(`--- discovery: GET /v4/users/me/dataTypes ---`);
  console.log(`HTTP ${dRes.getResponseCode()} — ${dRes.getContentText().substring(0, 600)}`);
  console.log('');

  // 2. Per-type probes. IDs taken from https://developers.google.com/health/data-types
  // Note: total-calories, active-minutes, floors etc are rollup-only and 400 on list.
  const candidates = [
    // Controls — currently used, expected 200
    'weight', 'daily-resting-heart-rate', 'daily-heart-rate-variability',
    'oxygen-saturation', 'steps', 'distance',
    // The broken one
    'sleep',
    // Sleep-adjacent (the reason we're probing)
    'daily-sleep-temperature-derivations',
    'respiratory-rate-sleep-summary',
    'daily-respiratory-rate',
    'respiratory-rate',
    'heart-rate-variability',
    // Other list-supported types from the docs
    'active-zone-minutes', 'activity-level', 'altitude', 'sedentary-period',
    'swim-lengths-data', 'daily-vo2-max', 'run-vo2-max', 'vo2-max',
    'exercise', 'heart-rate', 'body-fat', 'height', 'daily-oxygen-saturation',
    'hydration-log',
  ];

  console.log(`--- probing ${candidates.length} dataTypes (unfiltered GET) ---`);
  candidates.forEach(type => {
    const url = `https://health.googleapis.com/v4/users/me/dataTypes/${type}/dataPoints`;
    const res = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code === 200) {
      const data = JSON.parse(res.getContentText());
      const n = (data.dataPoints || []).length;
      console.log(`✅ ${type} — HTTP 200, ${n} point(s)`);
    } else {
      const body = res.getContentText().replace(/\s+/g, ' ').substring(0, 160);
      const tag = code === 404 ? '⚪' : code === 500 ? '🔥' : '❌';
      console.log(`${tag} ${type} — HTTP ${code} — ${body}`);
    }
    Utilities.sleep(150); // be polite
  });
}


// =============================================================================
// DIAGNOSTIC: dump full schema for sleep-proxy candidates so we can see
// the field shape and confirm they cover the recent nights `sleep` is dropping.
// =============================================================================
function inspectSleepProxies() {
  const healthService = getHealthService();
  if (!healthService.hasAccess()) { console.error('Not authorized.'); return; }

  const headers = { 'Authorization': 'Bearer ' + healthService.getAccessToken(), 'Content-Type': 'application/json' };
  const types = [
    'daily-sleep-temperature-derivations',
    'respiratory-rate-sleep-summary',
    'heart-rate-variability',
    'daily-respiratory-rate',
  ];

  types.forEach(type => {
    const url = `https://health.googleapis.com/v4/users/me/dataTypes/${type}/dataPoints`;
    const res = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
    const code = res.getResponseCode();
    console.log(`\n========== ${type} (HTTP ${code}) ==========`);
    if (code !== 200) {
      console.log(res.getContentText().substring(0, 400));
      return;
    }
    const data = JSON.parse(res.getContentText());
    const points = data.dataPoints || [];
    console.log(`count: ${points.length}`);
    if (points.length === 0) return;

    // Pull every timestamp-shaped string from each point and keep the max.
    // Schema varies per type, so this is the cheapest way to find "most recent".
    const tsRe = /"\d{4}-\d{2}-\d{2}(?:T[\d:.Z+-]+)?"/g;
    let mostRecent = '';
    let mostRecentPoint = null;
    points.forEach(p => {
      const matches = JSON.stringify(p).match(tsRe) || [];
      matches.forEach(m => {
        const v = m.slice(1, -1);
        if (v > mostRecent) { mostRecent = v; mostRecentPoint = p; }
      });
    });
    console.log(`most recent timestamp seen: ${mostRecent}`);
    console.log(`first point:\n${JSON.stringify(points[0], null, 2)}`);
    if (mostRecentPoint && mostRecentPoint !== points[0]) {
      console.log(`most recent point:\n${JSON.stringify(mostRecentPoint, null, 2)}`);
    }
  });
}


// =============================================================================
// DIAGNOSTIC: probe altitude with date filters + dump VO2 max variants
// to compare their schemas and sources.
// =============================================================================
function inspectExtras() {
  const healthService = getHealthService();
  if (!healthService.hasAccess()) { console.error('Not authorized.'); return; }
  const headers = { 'Authorization': 'Bearer ' + healthService.getAccessToken(), 'Content-Type': 'application/json' };
  const tz = Session.getScriptTimeZone();
  const todayStr   = Utilities.formatDate(new Date(),                            tz, 'yyyy-MM-dd');
  const monthAgoStr = Utilities.formatDate(new Date(Date.now() - 30*86400000),   tz, 'yyyy-MM-dd');
  const monthAgoZ  = `${monthAgoStr}T00:00:00Z`;

  // ---- altitude: try a few filter shapes since unfiltered returned 0 points ----
  console.log(`========== altitude probes ==========`);
  const altProbes = [
    ['unfiltered',                            ''],
    ['altitude.start_time.physical_time 30d', `altitude.start_time.physical_time>="${monthAgoZ}"`],
    ['altitude.sample_time.physical_time 30d', `altitude.sample_time.physical_time>="${monthAgoZ}"`],
    ['altitude.civil_start_time 30d',         `altitude.civil_start_time>="${monthAgoStr}"`],
  ];
  altProbes.forEach(([label, filter]) => {
    const url = `https://health.googleapis.com/v4/users/me/dataTypes/altitude/dataPoints` +
      (filter ? `?filter=${encodeURIComponent(filter)}` : '');
    const res = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code === 200) {
      const data = JSON.parse(res.getContentText());
      const n = (data.dataPoints || []).length;
      console.log(`✅ [${label}] ${n} point(s)`);
      if (n > 0) console.log(`   sample: ${JSON.stringify(data.dataPoints[0]).substring(0, 500)}`);
    } else {
      console.log(`❌ [${label}] HTTP ${code} — ${res.getContentText().replace(/\s+/g, ' ').substring(0, 160)}`);
    }
    Utilities.sleep(150);
  });

  // ---- body-fat + active-zone-minutes: schema check for backfill ----
  ['body-fat', 'active-zone-minutes'].forEach(type => {
    console.log(`\n========== ${type} ==========`);
    const url = `https://health.googleapis.com/v4/users/me/dataTypes/${type}/dataPoints`;
    const res = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) { console.log(`HTTP ${res.getResponseCode()} — ${res.getContentText().substring(0, 200)}`); return; }
    const points = JSON.parse(res.getContentText()).dataPoints || [];
    console.log(`count: ${points.length}`);
    if (points[0]) console.log(`first:\n${JSON.stringify(points[0], null, 2)}`);
  });

  // ---- VO2 max variants: dump 1-2 samples each so we can compare ----
  ['daily-vo2-max', 'run-vo2-max', 'vo2-max'].forEach(type => {
    console.log(`\n========== ${type} ==========`);
    const url = `https://health.googleapis.com/v4/users/me/dataTypes/${type}/dataPoints`;
    const res = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) { console.log(`HTTP ${res.getResponseCode()} — ${res.getContentText().substring(0, 200)}`); return; }
    const points = JSON.parse(res.getContentText()).dataPoints || [];
    console.log(`count: ${points.length}`);
    if (points[0]) console.log(`first:\n${JSON.stringify(points[0], null, 2)}`);
    if (points[1]) console.log(`second:\n${JSON.stringify(points[1], null, 2)}`);
  });
}


// =============================================================================
// UTILITY
// =============================================================================
function getOrCreateRowForDate(sheet, dateStr) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const val      = data[i][0];
    const existing = (val instanceof Date)
      ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(val).substring(0, 10);
    if (existing === dateStr) return i + 1; // 1-based row number
  }
  sheet.appendRow([dateStr]);
  return sheet.getLastRow();
}
