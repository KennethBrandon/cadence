// =========================================================================
// === GOOGLE HEALTH API: WORKOUT DATA SYNC (REPLACES LEGACY FITBIT API) ===
// =========================================================================

// Stored in Script Properties (Project Settings → Script Properties).
// Required key: WORKOUT_SPREADSHEET_ID.
const WORKOUT_SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('WORKOUT_SPREADSHEET_ID');

// --- ONE-TIME HISTORICAL WORKOUT BACKFILL ---
function runHistoricalWorkoutBackfill() {
  const healthService = getHealthService();
  if (!healthService.hasAccess()) {
    console.error('App is not authorized. Please run the authorize() function first.');
    return;
  }

  console.log("Starting Historical Workout Backfill...");

  const startDateStr = "2024-01-01";

  const headers = {
    'Authorization': 'Bearer ' + healthService.getAccessToken(),
    'Content-Type': 'application/json'
  };

  function fetchAllWorkouts(filterString) {
    let allDataPoints = [];
    let pageToken = "";

    do {
      let url = `https://health.googleapis.com/v4/users/me/dataTypes/exercise/dataPoints`;
      let params = [];
      if (filterString) params.push(`filter=${encodeURIComponent(filterString)}`);
      if (pageToken) params.push(`pageToken=${pageToken}`);
      if (params.length > 0) url += "?" + params.join("&");

      const options = { method: 'get', headers: headers, muteHttpExceptions: true };
      const response = UrlFetchApp.fetch(url, options);

      if (response.getResponseCode() !== 200) {
        console.error("Error fetching workouts: " + response.getContentText());
        break;
      }

      const result = JSON.parse(response.getContentText());
      if (result.dataPoints) {
        allDataPoints = allDataPoints.concat(result.dataPoints);
      }
      pageToken = result.nextPageToken;
    } while (pageToken);

    return allDataPoints;
  }

  const workoutPoints = fetchAllWorkouts(`exercise.interval.civil_start_time>="${startDateStr}"`);

  if (workoutPoints.length === 0) {
    console.log("No workouts found in the specified date range.");
    return;
  }

  const rows = [];
  workoutPoints.forEach(p => {
    if (!p.exercise || !p.exercise.interval) return;

    const startTime = new Date(p.exercise.interval.startTime);
    const logId = p.name ? p.name.split('/').pop() : startTime.getTime().toString();

    let durationMins = "";
    if (p.exercise.activeDuration) {
      durationMins = (parseInt(p.exercise.activeDuration.replace('s', '')) / 60).toFixed(2);
    } else {
      const endTime = new Date(p.exercise.interval.endTime);
      durationMins = ((endTime - startTime) / 60000).toFixed(2);
    }

    const activityName = p.exercise.displayName || p.exercise.exerciseType || "Workout";

    // Initialize all extended metrics
    let distanceMiles = "", paceDecMins = "", elevationFt = "", avgHr = "", calories = "";
    let steps = "", azm = "", lightMins = "", modMins = "", vigMins = "", peakMins = "";

    if (p.exercise.metricsSummary) {
      const m = p.exercise.metricsSummary;

      distanceMiles = m.distanceMillimeters ? (m.distanceMillimeters / 1609344).toFixed(2) : "";
      paceDecMins = m.averagePaceSecondsPerMeter ? ((m.averagePaceSecondsPerMeter * 1609.344) / 60).toFixed(2) : "";
      elevationFt = m.elevationGainMillimeters ? (m.elevationGainMillimeters / 304.8).toFixed(0) : "";
      avgHr = m.averageHeartRateBeatsPerMinute || "";
      calories = m.caloriesKcal || "";
      steps = m.steps || "";
      azm = m.activeZoneMinutes || "";

      if (m.heartRateZoneDurations) {
        const hr = m.heartRateZoneDurations;
        lightMins = hr.lightTime ? (parseInt(hr.lightTime.replace('s', '')) / 60).toFixed(1) : "";
        modMins = hr.moderateTime ? (parseInt(hr.moderateTime.replace('s', '')) / 60).toFixed(1) : "";
        vigMins = hr.vigorousTime ? (parseInt(hr.vigorousTime.replace('s', '')) / 60).toFixed(1) : "";
        peakMins = hr.peakTime ? (parseInt(hr.peakTime.replace('s', '')) / 60).toFixed(1) : "";
      }
    }

    rows.push([
      startTime,      // A: Date/Time
      activityName,   // B: Activity
      durationMins,   // C: Duration (mins)
      distanceMiles,  // D: Distance (miles)
      paceDecMins,    // E: Avg Pace (min/mi)
      elevationFt,    // F: Elevation Gain (ft)
      avgHr,          // G: Avg Heart Rate
      calories,       // H: Calories
      steps,          // I: Steps
      azm,            // J: Active Zone Minutes
      lightMins,      // K: Light HR (mins)
      modMins,        // L: Moderate HR (mins)
      vigMins,        // M: Vigorous HR (mins)
      peakMins,       // N: Peak HR (mins)
      logId           // O: Google Health ID
    ]);
  });

  rows.sort((a, b) => a[0] - b[0]);

  const ss = SpreadsheetApp.openById(WORKOUT_SPREADSHEET_ID);

  let sheet = ss.getSheetByName('FitBit Activities');
  if (!sheet) {
    sheet = ss.insertSheet('FitBit Activities');
  }

  const headersRow = [
    "Date/Time", "Activity", "Duration (mins)", "Distance (mi)", "Avg Pace (min/mi)",
    "Elevation (ft)", "Avg Heart Rate", "Calories", "Steps", "Active Zone Mins",
    "Light HR (mins)", "Moderate HR (mins)", "Vigorous HR (mins)", "Peak HR (mins)", "Google Health ID"
  ];

  sheet.clear();
  sheet.appendRow(headersRow);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

  console.log(`BOOM! Successfully backfilled ${rows.length} workouts with ALL telemetry!`);
}


// --- NIGHTLY WORKOUT SYNC (Add to your daily trigger) ---
function runDailyWorkoutSync() {
  const healthService = getHealthService();
  if (!healthService.hasAccess()) {
    console.error('App is not authorized.');
    return;
  }

  console.log("Starting Daily Workout Sync...");

  const lookbackDate = new Date();
  lookbackDate.setDate(lookbackDate.getDate() - 3);
  const lookbackStr = Utilities.formatDate(lookbackDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

  const filterString = `exercise.interval.civil_start_time>="${lookbackStr}"`;
  const url = `https://health.googleapis.com/v4/users/me/dataTypes/exercise/dataPoints?filter=${encodeURIComponent(filterString)}`;

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + healthService.getAccessToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    console.error("API Error: " + response.getContentText());
    return;
  }

  const json = JSON.parse(response.getContentText());
  if (!json.dataPoints || json.dataPoints.length === 0) {
    console.log("No recent workouts found.");
    return;
  }

  const ss = SpreadsheetApp.openById(WORKOUT_SPREADSHEET_ID);

  let sheet = ss.getSheetByName('FitBit Activities');
  if (!sheet) sheet = ss.insertSheet('FitBit Activities');

  const existingData = sheet.getRange(2, 15, Math.max(1, sheet.getLastRow() - 1)).getValues();
  const existingLogIds = existingData.flat().map(String);

  const rowsToAppend = [];

  json.dataPoints.forEach(p => {
    if (!p.exercise || !p.exercise.interval) return;

    const startTime = new Date(p.exercise.interval.startTime);
    const logId = p.name ? p.name.split('/').pop() : startTime.getTime().toString();

    if (!existingLogIds.includes(String(logId))) {
      let durationMins = "";
      if (p.exercise.activeDuration) {
        durationMins = (parseInt(p.exercise.activeDuration.replace('s', '')) / 60).toFixed(2);
      } else {
        const endTime = new Date(p.exercise.interval.endTime);
        durationMins = ((endTime - startTime) / 60000).toFixed(2);
      }

      const activityName = p.exercise.displayName || p.exercise.exerciseType || "Workout";

      let distanceMiles = "", paceDecMins = "", elevationFt = "", avgHr = "", calories = "";
      let steps = "", azm = "", lightMins = "", modMins = "", vigMins = "", peakMins = "";

      if (p.exercise.metricsSummary) {
        const m = p.exercise.metricsSummary;

        distanceMiles = m.distanceMillimeters ? (m.distanceMillimeters / 1609344).toFixed(2) : "";
        paceDecMins = m.averagePaceSecondsPerMeter ? ((m.averagePaceSecondsPerMeter * 1609.344) / 60).toFixed(2) : "";
        elevationFt = m.elevationGainMillimeters ? (m.elevationGainMillimeters / 304.8).toFixed(0) : "";
        avgHr = m.averageHeartRateBeatsPerMinute || "";
        calories = m.caloriesKcal || "";
        steps = m.steps || "";
        azm = m.activeZoneMinutes || "";

        if (m.heartRateZoneDurations) {
          const hr = m.heartRateZoneDurations;
          lightMins = hr.lightTime ? (parseInt(hr.lightTime.replace('s', '')) / 60).toFixed(1) : "";
          modMins = hr.moderateTime ? (parseInt(hr.moderateTime.replace('s', '')) / 60).toFixed(1) : "";
          vigMins = hr.vigorousTime ? (parseInt(hr.vigorousTime.replace('s', '')) / 60).toFixed(1) : "";
          peakMins = hr.peakTime ? (parseInt(hr.peakTime.replace('s', '')) / 60).toFixed(1) : "";
        }
      }

      rowsToAppend.push([
        startTime, activityName, durationMins, distanceMiles, paceDecMins,
        elevationFt, avgHr, calories, steps, azm,
        lightMins, modMins, vigMins, peakMins, logId
      ]);
      existingLogIds.push(String(logId));
    }
  });

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    console.log(`Appended ${rowsToAppend.length} new workout(s).`);
  } else {
    console.log("Checked recent workouts, but none were new.");
  }
}
