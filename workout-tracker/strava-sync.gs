// Secrets are stored in Script Properties (Project Settings → Script Properties).
// Required keys: STRAVA_CLIENT_ID, STRAVA_CLIENT_SECRET, STRAVA_REFRESH_TOKEN.
const STRAVA_CLIENT_ID = PropertiesService.getScriptProperties().getProperty('STRAVA_CLIENT_ID');
const STRAVA_CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('STRAVA_CLIENT_SECRET');
const STRAVA_REFRESH_TOKEN = PropertiesService.getScriptProperties().getProperty('STRAVA_REFRESH_TOKEN');

const WORKOUT_TYPES = { 0: 'Default', 1: 'Race', 2: 'Long Run', 3: 'Workout' };

// Columns in 'Strava Workouts' sheet (1-based):
// --- IDENTITY ---
// A=1  Activity ID hyperlink
// B=2  Date
// C=3  Start Time
// D=4  Name
// E=5  Type
// F=6  Workout Type
// G=7  Description
// H=8  Private Note
// --- CORE METRICS ---
// I=9  Distance
// J=10 Moving Time
// K=11 Elapsed Time
// L=12 Avg HR
// M=13 Max HR
// N=14 Avg Cadence
// O=15 Max Speed
// P=16 Calories
// Q=17 Suffer Score
// --- CONTEXT ---
// R=18 Elevation Gain
// S=19 Temp
// T=20 Trainer
// --- SOCIAL / ACHIEVEMENT ---
// U=21 PR Count
// V=22 Athlete Count
// --- EQUIPMENT ---
// W=23 Gear
// X=24 Device


// =============================================================================
// MAIN — runs every night
// =============================================================================
function runDailySync() {
  console.log("Starting Daily Sync...");

  const tokenResponse = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
    method: 'post',
    payload: {
      client_id: STRAVA_CLIENT_ID, client_secret: STRAVA_CLIENT_SECRET,
      refresh_token: STRAVA_REFRESH_TOKEN, grant_type: 'refresh_token'
    }
  });
  const accessToken = JSON.parse(tokenResponse.getContentText()).access_token;
  const headers = { 'Authorization': `Bearer ${accessToken}` };

  syncNewWorkoutsAndPRs(headers);
  syncStatsAndGear();

  console.log("Daily Sync Complete!");
}


// =============================================================================
// SYNC NEW WORKOUTS, PRs, AND SEGMENT EFFORTS
// =============================================================================
function syncNewWorkoutsAndPRs(headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workoutsSheet = ss.getSheetByName('Strava Workouts');
  const prSheet = ss.getSheetByName('Personal Records');

  const lastRow = workoutsSheet.getLastRow();
  const existingIds = lastRow > 1
    ? workoutsSheet.getRange("A2:A" + lastRow).getValues().flat().map(id => {
        const match = id.toString().match(/\d{10,}/);
        return match ? match[0] : id.toString();
      })
    : [];

  const after = Math.floor(Date.now() / 1000) - (2 * 24 * 60 * 60); // 2 days ago

  const activitiesResponse = UrlFetchApp.fetch(
    `https://www.strava.com/api/v3/athlete/activities?after=${after}&per_page=30`,
    { headers: headers }
  );
  const summaryActivities = JSON.parse(activitiesResponse.getContentText());

  summaryActivities.reverse().forEach(summary => {
    const stravaId = summary.id.toString();

    if (!existingIds.includes(stravaId)) {
      Utilities.sleep(1000);
      const detailResponse = UrlFetchApp.fetch(
        `https://www.strava.com/api/v3/activities/${stravaId}`,
        { headers: headers }
      );
      const activity = JSON.parse(detailResponse.getContentText());

      // Format Date/Time
      const rawLocal = summary.start_date_local;
      const datePart = rawLocal.split('T')[0];
      const timePart = rawLocal.split('T')[1].replace('Z', '');
      const dateObj = new Date(datePart.replace(/-/g, '/'));
      const formattedDate = Utilities.formatDate(dateObj, "GMT", "EEE, MMM dd, yyyy");
      const [h, m] = timePart.split(':');
      const ampm = h >= 12 ? 'PM' : 'AM';
      const displayTime = ((h % 12) || 12) + ":" + m + " " + ampm;
      const linkFormula = `=HYPERLINK("https://www.strava.com/activities/${stravaId}", "${stravaId}")`;

      // 1. Add to Workouts sheet
      workoutsSheet.appendRow([
        linkFormula,                                                          // A: Activity ID
        formattedDate,                                                        // B: Date
        displayTime,                                                          // C: Start Time
        activity.name,                                                        // D: Name
        activity.type,                                                        // E: Type
        WORKOUT_TYPES[activity.workout_type] || "Default",                    // F: Workout Type
        activity.description || "",                                           // G: Description
        activity.private_note || "",                                          // H: Private Note
        (activity.distance * 0.000621371).toFixed(2) + " mi",                // I: Distance
        formatSec(activity.moving_time),                                      // J: Moving Time
        formatSec(activity.elapsed_time),                                     // K: Elapsed Time
        activity.has_heartrate ? Math.round(activity.average_heartrate) : "", // L: Avg HR
        activity.has_heartrate ? activity.max_heartrate : "",                 // M: Max HR
        activity.average_cadence ? String(activity.average_cadence) : "",    // N: Avg Cadence
        activity.max_speed ? (activity.max_speed * 2.23694).toFixed(1) + " mph" : "", // O: Max Speed
        activity.calories || "",                                              // P: Calories
        activity.suffer_score || "0",                                         // Q: Suffer Score
        activity.total_elevation_gain ? Math.round(activity.total_elevation_gain * 3.28) + "ft" : "", // R: Elevation Gain
        activity.average_temp != null ? activity.average_temp + "°C" : "",   // S: Temp
        activity.trainer ? "Yes" : "No",                                      // T: Trainer
        String(activity.pr_count || 0),                                       // U: PR Count
        String(activity.athlete_count || 1),                                  // V: Athlete Count
        activity.gear ? activity.gear.name : "None",                          // W: Gear
        activity.device_name || "Unknown"                                     // X: Device
      ]);

      // 2. Log segment efforts
      logSegmentEfforts(activity, linkFormula, formattedDate);

      // 3. Check for PRs (Top 3 All-Time)
      if (activity.best_efforts) {
        activity.best_efforts.forEach(effort => {
          if (effort.pr_rank && effort.pr_rank <= 3) {
            prSheet.appendRow([
              linkFormula, formattedDate,
              effort.name + ` (Rank: ${effort.pr_rank})`,
              formatSec(effort.elapsed_time)
            ]);
          }
        });
      }
    }
  });
}


// =============================================================================
// SEGMENT EFFORTS HELPER
// =============================================================================
function logSegmentEfforts(activity, linkFormula, formattedDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let segSheet = ss.getSheetByName('Segment Efforts');

  if (!segSheet) {
    segSheet = ss.insertSheet('Segment Efforts');
    segSheet.appendRow([
      'Activity', 'Date', 'Segment', 'Distance (mi)',
      'Elapsed Time', 'Avg HR', 'Max HR', 'Avg Grade %', 'PR Rank'
    ]);
  }

  if (!activity.segment_efforts || activity.segment_efforts.length === 0) return;

  activity.segment_efforts.forEach(effort => {
    segSheet.appendRow([
      linkFormula,
      formattedDate,
      effort.name,
      effort.segment ? (effort.segment.distance * 0.000621371).toFixed(2) : "",
      formatSec(effort.elapsed_time),
      effort.average_heartrate ? Math.round(effort.average_heartrate) : "",
      effort.max_heartrate || "",
      effort.segment ? effort.segment.average_grade : "",
      effort.pr_rank || ""
    ]);
  });
}


// =============================================================================
// STATS & GEAR DASHBOARD
// =============================================================================
function syncStatsAndGear() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statsSheet = ss.getSheetByName('Athlete Stats');
  const workoutsSheet = ss.getSheetByName('Strava Workouts');

  const data = workoutsSheet.getDataRange().getValues();
  const currentYear = new Date().getFullYear().toString();

  let stats = {
    allTimeRun: 0, allTimeHike: 0, allTimeElev: 0, allTimeCal: 0,
    ytdRun: 0, ytdHike: 0, ytdElev: 0, ytdCal: 0
  };
  let gearTracker = {};

  for (let i = 1; i < data.length; i++) {
    const type     = data[i][4]  ? data[i][4].toString()  : "";  // E: Type
    const dateStr  = data[i][1]  ? data[i][1].toString()  : "";  // B: Date
    const distStr  = data[i][8]  ? data[i][8].toString()  : "";  // I: Distance
    const calStr   = data[i][15] ? data[i][15].toString() : "0"; // P: Calories
    const climbStr = data[i][17] ? data[i][17].toString() : "";  // R: Elevation Gain
    const gearName = data[i][22] ? data[i][22].toString() : "None"; // W: Gear

    let dist  = parseFloat(distStr.replace(' mi', '')) || 0;
    let climb = parseInt(climbStr.replace('ft', '')) || 0;
    let cal   = parseInt(calStr.replace(/,/g, '')) || 0;
    let isCurrentYear = dateStr.includes(currentYear);

    stats.allTimeElev += climb;
    stats.allTimeCal  += cal;
    if (isCurrentYear) { stats.ytdElev += climb; stats.ytdCal += cal; }

    if (type === 'Run') {
      stats.allTimeRun += dist;
      if (isCurrentYear) stats.ytdRun += dist;
    } else if (type === 'Hike' || type === 'Walk') {
      stats.allTimeHike += dist;
      if (isCurrentYear) stats.ytdHike += dist;
    }

    if (gearName !== "None" && gearName !== "") {
      if (!gearTracker[gearName]) gearTracker[gearName] = 0;
      gearTracker[gearName] += dist;
    }
  }

  statsSheet.clearContents();

  statsSheet.appendRow(['🏃‍♂️ RUNNING STATS', 'Miles']);
  statsSheet.appendRow(['All-Time Distance Ran', stats.allTimeRun.toFixed(2)]);
  statsSheet.appendRow(['YTD Distance Ran (' + currentYear + ')', stats.ytdRun.toFixed(2)]);
  statsSheet.appendRow(['']);

  statsSheet.appendRow(['🥾 HIKING & WALKING STATS', 'Miles']);
  statsSheet.appendRow(['All-Time Distance Hiked', stats.allTimeHike.toFixed(2)]);
  statsSheet.appendRow(['YTD Distance Hiked (' + currentYear + ')', stats.ytdHike.toFixed(2)]);
  statsSheet.appendRow(['']);

  statsSheet.appendRow(['🏔️ ELEVATION GAIN', 'Feet']);
  statsSheet.appendRow(['All-Time Elevation Gain', Math.round(stats.allTimeElev).toLocaleString()]);
  statsSheet.appendRow(['YTD Elevation Gain (' + currentYear + ')', Math.round(stats.ytdElev).toLocaleString()]);
  statsSheet.appendRow(['']);

  statsSheet.appendRow(['🔥 CALORIES BURNED', 'kcal']);
  statsSheet.appendRow(['All-Time Calories', Math.round(stats.allTimeCal).toLocaleString()]);
  statsSheet.appendRow(['YTD Calories (' + currentYear + ')', Math.round(stats.ytdCal).toLocaleString()]);
  statsSheet.appendRow(['']);

  statsSheet.appendRow(['👟 GEAR MILEAGE LOG', 'Miles']);
  for (const [gear, miles] of Object.entries(gearTracker)) {
    statsSheet.appendRow([gear, miles.toFixed(2)]);
  }
}


// =============================================================================
// ONE-TIME SETUP: Write all headers and format problem columns as plain text
// Run this once before running any backfill
// =============================================================================
function addNewHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Strava Workouts');
  const headers = [
    'Activity ID', 'Date', 'Start Time', 'Name', 'Type', 'Workout Type',
    'Description', 'Private Note', 'Distance', 'Moving Time', 'Elapsed Time',
    'Avg HR', 'Max HR', 'Avg Cadence', 'Max Speed', 'Calories', 'Suffer Score',
    'Elevation Gain', 'Temp', 'Trainer', 'PR Count', 'Athlete Count', 'Gear', 'Device'
  ];
  // Format columns that would otherwise be misread as times or dates
  // K=11 (Elapsed Time), N=14 (Avg Cadence), U=21 (PR Count), V=22 (Athlete Count)
  [11, 14, 21, 22].forEach(col => {
    sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
  });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  console.log("Headers written. Columns K, N, U, V formatted as plain text.");
}


// =============================================================================
// ONE-TIME SETUP: Migrate existing data to new column order
// Run ONCE after addNewHeaders() — reads old layout, rewrites in new layout
// =============================================================================
function migrateColumnOrder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Strava Workouts');
  const data = sheet.getDataRange().getValues();

  // Skip header row — start from row index 1
  const rows = data.slice(1);
  if (rows.length === 0) { console.log("No data to migrate."); return; }

  // Old column indices (0-based) before migration:
  // [0]=ID, [1]=Type, [2]=Date, [3]=Distance, [4]=Moving Time, [5]=Name,
  // [6]=Avg HR, [7]=Max HR, [8]=Calories, [9]=Elevation, [10]=Suffer Score,
  // [11]=Gear, [12]=Device, [13]=Description, [14]=Start Time,
  // [15]=Private Note, [16]=Elapsed Time, [17]=Max Speed, [18]=Avg Cadence,
  // [19]=Temp, [20]=Workout Type, [21]=Trainer, [22]=PR Count, [23]=Athlete Count
  const remapped = rows.map(r => [
    r[0],  // A: Activity ID
    r[2],  // B: Date
    r[14], // C: Start Time
    r[5],  // D: Name
    r[1],  // E: Type
    r[20], // F: Workout Type
    r[13], // G: Description
    r[15], // H: Private Note
    r[3],  // I: Distance
    r[4],  // J: Moving Time
    r[16], // K: Elapsed Time
    r[6],  // L: Avg HR
    r[7],  // M: Max HR
    r[18], // N: Avg Cadence
    r[17], // O: Max Speed
    r[8],  // P: Calories
    r[10], // Q: Suffer Score
    r[9],  // R: Elevation Gain
    r[19], // S: Temp
    r[21], // T: Trainer
    r[22], // U: PR Count
    r[23], // V: Athlete Count
    r[11], // W: Gear
    r[12]  // X: Device
  ]);

  // Clear data rows (keep row 1 header), then write remapped data
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 24).clearContent();
  sheet.getRange(2, 1, remapped.length, 24).setValues(remapped);
  console.log(`Migration complete — ${remapped.length} rows reordered.`);
}


// =============================================================================
// ONE-TIME BACKFILL: Fill new columns for existing workout rows
// Safe to re-run — skips rows where Elapsed Time (K) is already populated
// =============================================================================
function backfillWorkoutColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workoutsSheet = ss.getSheetByName('Strava Workouts');

  const tokenResponse = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
    method: 'post',
    payload: {
      client_id: STRAVA_CLIENT_ID, client_secret: STRAVA_CLIENT_SECRET,
      refresh_token: STRAVA_REFRESH_TOKEN, grant_type: 'refresh_token'
    }
  });
  const accessToken = JSON.parse(tokenResponse.getContentText()).access_token;
  const headers = { 'Authorization': `Bearer ${accessToken}` };

  const data = workoutsSheet.getDataRange().getValues();
  let apiCalls = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Skip if Elapsed Time (K, index 10) is already filled
    if (row[10] !== "") continue;

    const idFormula = row[0].toString();
    const match = idFormula.match(/\d{10,}/);
    if (!match) continue;
    const stravaId = match[0];

    apiCalls++;
    if (apiCalls > 95) {
      console.log("Approaching API limit — run again to continue. Stopped at row " + (i + 1));
      break;
    }

    console.log(`Backfilling row ${i + 1} (ID: ${stravaId})...`);
    Utilities.sleep(1500);

    try {
      const detailResponse = UrlFetchApp.fetch(
        `https://www.strava.com/api/v3/activities/${stravaId}`,
        { headers, muteHttpExceptions: true }
      );
      if (detailResponse.getResponseCode() === 429) {
        console.log("Rate limit hit at row " + (i + 1) + " — wait for the next :00/:15/:30/:45 and re-run.");
        break;
      }
      const activity = JSON.parse(detailResponse.getContentText());

      // Write each new column individually (they are not contiguous in the new layout)
      workoutsSheet.getRange(i + 1, 6).setValue(WORKOUT_TYPES[activity.workout_type] || "Default"); // F: Workout Type
      workoutsSheet.getRange(i + 1, 8).setValue(activity.private_note || "");                        // H: Private Note
      workoutsSheet.getRange(i + 1, 11).setValue(formatSec(activity.elapsed_time));                  // K: Elapsed Time
      workoutsSheet.getRange(i + 1, 14).setValue(activity.average_cadence ? String(activity.average_cadence) : ""); // N: Avg Cadence
      workoutsSheet.getRange(i + 1, 15).setValue(activity.max_speed ? (activity.max_speed * 2.23694).toFixed(1) + " mph" : ""); // O: Max Speed
      workoutsSheet.getRange(i + 1, 19).setValue(activity.average_temp != null ? activity.average_temp + "°C" : ""); // S: Temp
      workoutsSheet.getRange(i + 1, 20).setValue(activity.trainer ? "Yes" : "No");                   // T: Trainer
      workoutsSheet.getRange(i + 1, 21).setValue(String(activity.pr_count || 0));                    // U: PR Count
      workoutsSheet.getRange(i + 1, 22).setValue(String(activity.athlete_count || 1));               // V: Athlete Count
    } catch(e) {
      console.log("Error on row " + (i + 1) + ": " + e);
    }
  }

  console.log("Backfill complete!");
}


// =============================================================================
// ONE-TIME FIX: Re-write time-formatted cells as plain text strings
// Run if Elapsed Time, Avg Cadence, PR Count, or Athlete Count show as times
// Safe to re-run — skips rows already fixed
// =============================================================================
function fixNumericColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Strava Workouts');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Pre-format problem columns as plain text
  [11, 14, 21, 22].forEach(col => {
    sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat("@");
  });

  const tokenResponse = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
    method: 'post',
    payload: {
      client_id: STRAVA_CLIENT_ID, client_secret: STRAVA_CLIENT_SECRET,
      refresh_token: STRAVA_REFRESH_TOKEN, grant_type: 'refresh_token'
    }
  });
  const accessToken = JSON.parse(tokenResponse.getContentText()).access_token;
  const headers = { 'Authorization': `Bearer ${accessToken}` };

  const data = sheet.getDataRange().getValues();
  let apiCalls = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Skip if Max Speed (O, index 14) is empty — row was never backfilled
    if (row[14] === "") continue;
    // Skip if Elapsed Time (K, index 10) is already a plain string — already fixed
    if (typeof row[10] === 'string' && row[10] !== "") continue;

    const idFormula = row[0].toString();
    const match = idFormula.match(/\d{10,}/);
    if (!match) continue;
    const stravaId = match[0];

    apiCalls++;
    if (apiCalls > 95) {
      console.log("API limit — run again to continue. Stopped at row " + (i + 1));
      break;
    }

    console.log("Fixing row " + (i + 1) + "...");
    Utilities.sleep(1500);

    try {
      const detailResponse = UrlFetchApp.fetch(
        `https://www.strava.com/api/v3/activities/${stravaId}`,
        { headers, muteHttpExceptions: true }
      );
      if (detailResponse.getResponseCode() === 429) {
        console.log("Rate limit hit at row " + (i + 1) + " — wait for the next :00/:15/:30/:45 and re-run.");
        break;
      }
      const activity = JSON.parse(detailResponse.getContentText());

      sheet.getRange(i + 1, 11).setValue(formatSec(activity.elapsed_time));                          // K: Elapsed Time
      sheet.getRange(i + 1, 14).setValue(activity.average_cadence ? String(activity.average_cadence) : ""); // N: Avg Cadence
      sheet.getRange(i + 1, 21).setValue(String(activity.pr_count || 0));                            // U: PR Count
      sheet.getRange(i + 1, 22).setValue(String(activity.athlete_count || 1));                       // V: Athlete Count
    } catch(e) {
      console.log("Error on row " + (i + 1) + ": " + e);
    }
  }

  console.log("Fix complete!");
}


// =============================================================================
// ONE-TIME BACKFILL: Historical segment efforts
// Safe to re-run — skips activities already in the Segment Efforts sheet
// =============================================================================
function findHistoricalSegments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workoutsSheet = ss.getSheetByName('Strava Workouts');

  // Segment Efforts sheet — created if missing
  let segSheet = ss.getSheetByName('Segment Efforts');
  if (!segSheet) {
    segSheet = ss.insertSheet('Segment Efforts');
    segSheet.appendRow([
      'Activity', 'Date', 'Segment', 'Distance (mi)',
      'Elapsed Time', 'Avg HR', 'Max HR', 'Avg Grade %', 'PR Rank'
    ]);
  }

  // Tracking sheet — records every activity ID we've scanned, even if it had no segments
  // This prevents re-fetching activities that returned 0 segment efforts
  let scanLog = ss.getSheetByName('Segments Scanned');
  if (!scanLog) {
    scanLog = ss.insertSheet('Segments Scanned');
    scanLog.appendRow(['Activity ID']);
  }
  // Seed from the scan log (tracks no-segment activities)
  const scannedIds = new Set(
    scanLog.getDataRange().getValues().flat().map(v => v.toString()).filter(v => v && v !== 'Activity ID')
  );

  // Also pull IDs already in Segment Efforts so runs before this log existed are covered
  segSheet.getDataRange().getValues().forEach(row => {
    const idMatch = row[0].toString().match(/\d{10,}/);
    if (idMatch) scannedIds.add(idMatch[0]);
  });

  const tokenResponse = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
    method: 'post',
    payload: {
      client_id: STRAVA_CLIENT_ID, client_secret: STRAVA_CLIENT_SECRET,
      refresh_token: STRAVA_REFRESH_TOKEN, grant_type: 'refresh_token'
    }
  });
  const accessToken = JSON.parse(tokenResponse.getContentText()).access_token;
  const headers = { 'Authorization': `Bearer ${accessToken}` };

  const data = workoutsSheet.getDataRange().getValues();
  let apiCalls = 0;

  for (let i = 1; i < data.length; i++) {
    const idFormula = data[i][0].toString();
    const type      = data[i][4].toString(); // E: Type
    const dateStr   = data[i][1];            // B: Date

    if (type !== 'Run' && type !== 'Hike' && type !== 'TrailRun' && type !== 'Walk') continue;

    const match = idFormula.match(/\d{10,}/);
    if (!match) continue;
    const stravaId = match[0];

    // Skip if we've already scanned this activity (with or without segments)
    if (scannedIds.has(stravaId)) continue;

    apiCalls++;
    if (apiCalls > 95) {
      console.log("Approaching API limit — run again to continue. Stopped at row " + (i + 1));
      break;
    }

    console.log(`Fetching segments for row ${i + 1} (${dateStr})...`);
    Utilities.sleep(1500);

    try {
      const detailResponse = UrlFetchApp.fetch(
        `https://www.strava.com/api/v3/activities/${stravaId}`,
        { headers, muteHttpExceptions: true }
      );
      if (detailResponse.getResponseCode() === 429) {
        console.log("Rate limit hit at row " + (i + 1) + " — wait for the next :00/:15/:30/:45 and re-run.");
        break;
      }
      const activity = JSON.parse(detailResponse.getContentText());
      const linkFormula = `=HYPERLINK("https://www.strava.com/activities/${stravaId}", "${stravaId}")`;
      logSegmentEfforts(activity, linkFormula, dateStr);

      // Mark as scanned regardless of whether it had segments
      scanLog.appendRow([stravaId]);
      scannedIds.add(stravaId);
    } catch(e) {
      console.log("Error on row " + (i + 1) + ": " + e);
    }
  }

  console.log("Historical Segment Scan Complete!");
}


// =============================================================================
// ONE-TIME BACKFILL: Historical PRs
// =============================================================================
function findHistoricalPRs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workoutsSheet = ss.getSheetByName('Strava Workouts');
  const prSheet = ss.getSheetByName('Personal Records');

  console.log("Starting Historical PR Scan...");

  const tokenResponse = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
    method: 'post',
    payload: {
      client_id: STRAVA_CLIENT_ID, client_secret: STRAVA_CLIENT_SECRET,
      refresh_token: STRAVA_REFRESH_TOKEN, grant_type: 'refresh_token'
    }
  });
  const accessToken = JSON.parse(tokenResponse.getContentText()).access_token;

  const data = workoutsSheet.getDataRange().getValues();
  const existingPRs = prSheet.getDataRange().getValues().flat().join(",");
  let apiCalls = 0;

  for (let i = 1; i < data.length; i++) {
    const idFormula = data[i][0].toString();
    const type      = data[i][4].toString(); // E: Type
    const dateStr   = data[i][1];            // B: Date

    if (type !== 'Run') continue;

    const match = idFormula.match(/\d{10,}/);
    if (!match) continue;
    const stravaId = match[0];

    if (existingPRs.includes(stravaId)) continue;

    apiCalls++;
    if (apiCalls > 95) {
      console.log("Approaching API limit! Stopping for now. Run again later to finish.");
      break;
    }

    console.log(`Checking Run: ${dateStr}...`);
    Utilities.sleep(1500);

    try {
      const detailResponse = UrlFetchApp.fetch(
        `https://www.strava.com/api/v3/activities/${stravaId}`,
        { headers: { 'Authorization': `Bearer ${accessToken}` }, muteHttpExceptions: true }
      );
      if (detailResponse.getResponseCode() === 429) {
        console.log("Rate limit hit — wait for the next :00/:15/:30/:45 and re-run.");
        break;
      }
      const activity = JSON.parse(detailResponse.getContentText());

      if (activity.best_efforts) {
        activity.best_efforts.forEach(effort => {
          if (effort.pr_rank && effort.pr_rank <= 3) {
            prSheet.appendRow([
              idFormula, dateStr,
              effort.name + ` (Rank: ${effort.pr_rank})`,
              formatSec(effort.elapsed_time)
            ]);
          }
        });
      }
    } catch(e) {
      console.log("Error or rate limit hit on ID: " + stravaId);
    }
  }

  console.log("Historical PR Scan Complete!");
}


// =============================================================================
// TRIGGER INSTALL — run once to enable nightly automation
// =============================================================================
function INSTALL_TRIGGER() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('runDailySync')
    .timeBased()
    .everyHours(6)
    .create();

  console.log("Trigger installed — will run every 6 hours (4x per day).");
}


// =============================================================================
// UTILITY
// =============================================================================
function formatSec(s) {
  if (!s) return "0:00:00";
  const h = Math.floor(s / 3600);
  const m = Math.floor((s % 3600) / 60);
  const sec = s % 60;
  return h + ":" + m.toString().padStart(2, '0') + ":" + sec.toString().padStart(2, '0');
}
