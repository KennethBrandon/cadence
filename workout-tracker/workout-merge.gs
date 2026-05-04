function mergeAndDeduplicateWorkouts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Get the Source Sheets
  const stravaSheet = ss.getSheetByName("Strava Workouts");
  const fitbitSheet = ss.getSheetByName("FitBit Workouts");

  if (!stravaSheet || !fitbitSheet) {
    console.error("Missing one of the source sheets. Please check the names.");
    return;
  }

  const stravaData = stravaSheet.getDataRange().getValues();
  const fitbitData = fitbitSheet.getDataRange().getValues();

  const stravaHeaders = stravaData.shift();
  const fitbitHeaders = fitbitData.shift();

  const getCol = (headers, name) => headers.indexOf(name);

  // 2. Process Strava Data (The Source of Truth)
  const processedStrava = [];
  const stravaBlocks = [];

  for (let i = 0; i < stravaData.length; i++) {
    const row = stravaData[i];
    if (!row[0]) continue;

    const dateVal = row[getCol(stravaHeaders, "Date")];
    const timeVal = row[getCol(stravaHeaders, "Start Time")];
    const startTime = parseStravaDateTime(dateVal, timeVal);

    // Use Elapsed Time for deduplication window (real clock time including pauses),
    // fall back to Moving Time if Elapsed Time column isn't populated yet
    const elapsedTimeVal = row[getCol(stravaHeaders, "Elapsed Time")];
    const movingTimeVal  = row[getCol(stravaHeaders, "Moving Time")];
    const durationMins   = parseDurationToMinutes(elapsedTimeVal || movingTimeVal);
    const endTime        = new Date(startTime.getTime() + (durationMins * 60000));

    stravaBlocks.push({ start: startTime.getTime(), end: endTime.getTime() });

    processedStrava.push({
      platform:     "Strava",
      datetime:     startTime,
      type:         row[getCol(stravaHeaders, "Type")],
      name:         row[getCol(stravaHeaders, "Name")],
      duration:     durationMins.toFixed(2),
      distance:     cleanNumber(row[getCol(stravaHeaders, "Distance")]),
      avgPace:      "",
      avgHr:        row[getCol(stravaHeaders, "Avg HR")],
      calories:     row[getCol(stravaHeaders, "Calories")],
      elevation:    cleanNumber(row[getCol(stravaHeaders, "Elevation Gain")]),
      stravaId:     row[getCol(stravaHeaders, "Activity ID")],
      maxHr:        row[getCol(stravaHeaders, "Max HR")],
      relativeEffort: row[getCol(stravaHeaders, "Suffer Score")],
      gear:         row[getCol(stravaHeaders, "Gear")],
      device:       row[getCol(stravaHeaders, "Device")],
      description:  row[getCol(stravaHeaders, "Description")],
      healthId:     "",
      steps:        "",
      azm:          "",
      lightHr:      "",
      modHr:        "",
      vigHr:        "",
      peakHr:       "",
      // New Strava fields
      privateNote:  row[getCol(stravaHeaders, "Private Note")],
      elapsedTime:  row[getCol(stravaHeaders, "Elapsed Time")],
      maxSpeed:     row[getCol(stravaHeaders, "Max Speed")],
      avgCadence:   row[getCol(stravaHeaders, "Avg Cadence")],
      temp:         row[getCol(stravaHeaders, "Temp")],
      workoutType:  row[getCol(stravaHeaders, "Workout Type")],
      trainer:      row[getCol(stravaHeaders, "Trainer")],
      prCount:      row[getCol(stravaHeaders, "PR Count")],
      athleteCount: row[getCol(stravaHeaders, "Athlete Count")]
    });
  }

  // 3. Process and Filter Fitbit Data
  const processedFitbit = [];
  const BUFFER_MS = 15 * 60000;
  let duplicatesSkipped = 0;

  for (let i = 0; i < fitbitData.length; i++) {
    const row = fitbitData[i];
    if (!row[0]) continue;

    const fitbitStart = new Date(row[getCol(fitbitHeaders, "Date/Time")]).getTime();

    let isDuplicate = false;
    for (const block of stravaBlocks) {
      const windowStart = block.start - BUFFER_MS;
      const windowEnd   = block.end   + BUFFER_MS;
      if (fitbitStart >= windowStart && fitbitStart <= windowEnd) {
        isDuplicate = true;
        break;
      }
    }

    if (isDuplicate) {
      duplicatesSkipped++;
    } else {
      processedFitbit.push({
        platform:     "Fitbit",
        datetime:     new Date(fitbitStart),
        type:         row[getCol(fitbitHeaders, "Activity")],
        name:         "Auto-Detected Activity",
        duration:     row[getCol(fitbitHeaders, "Duration (mins)")],
        distance:     row[getCol(fitbitHeaders, "Distance (mi)")],
        avgPace:      row[getCol(fitbitHeaders, "Avg Pace (min/mi)")],
        avgHr:        row[getCol(fitbitHeaders, "Avg Heart Rate")],
        calories:     row[getCol(fitbitHeaders, "Calories")],
        elevation:    row[getCol(fitbitHeaders, "Elevation (ft)")],
        stravaId:     "",
        maxHr:        "",
        relativeEffort: "",
        gear:         "",
        device:       "",
        description:  "",
        healthId:     row[getCol(fitbitHeaders, "Google Health ID")],
        steps:        row[getCol(fitbitHeaders, "Steps")],
        azm:          row[getCol(fitbitHeaders, "Active Zone Mins")],
        lightHr:      row[getCol(fitbitHeaders, "Light HR (mins)")],
        modHr:        row[getCol(fitbitHeaders, "Moderate HR (mins)")],
        vigHr:        row[getCol(fitbitHeaders, "Vigorous HR (mins)")],
        peakHr:       row[getCol(fitbitHeaders, "Peak HR (mins)")],
        // Strava-only fields — blank for Fitbit rows
        privateNote:  "",
        elapsedTime:  "",
        maxSpeed:     "",
        avgCadence:   "",
        temp:         "",
        workoutType:  "",
        trainer:      "",
        prCount:      "",
        athleteCount: ""
      });
    }
  }

  // 4. Combine and Sort (newest first)
  let mergedData = processedStrava.concat(processedFitbit);
  mergedData.sort((a, b) => b.datetime.getTime() - a.datetime.getTime());

  // 5. Write to the "Merged Workouts" Sheet
  let mergedSheet = ss.getSheetByName("Merged Workouts");
  if (!mergedSheet) {
    mergedSheet = ss.insertSheet("Merged Workouts");
  } else {
    mergedSheet.clear();
  }

  const finalHeaders = [
    "Platform", "Date/Time", "Activity Type", "Name", "Duration (mins)",
    "Distance (mi)", "Avg Pace (min/mi)", "Avg HR", "Calories", "Elevation (ft)",
    "Strava ID", "Max HR", "Relative Effort", "Gear", "Device", "Description",
    "Google Health ID", "Steps", "Active Zone Mins", "Light HR", "Moderate HR",
    "Vigorous HR", "Peak HR",
    // New Strava fields
    "Private Note", "Elapsed Time", "Max Speed", "Avg Cadence", "Temp",
    "Workout Type", "Trainer", "PR Count", "Athlete Count"
  ];

  const outputArray = [finalHeaders];
  mergedData.forEach(row => {
    outputArray.push([
      row.platform, row.datetime, row.type, row.name, row.duration,
      row.distance, row.avgPace, row.avgHr, row.calories, row.elevation,
      row.stravaId, row.maxHr, row.relativeEffort, row.gear, row.device, row.description,
      row.healthId, row.steps, row.azm, row.lightHr, row.modHr, row.vigHr, row.peakHr,
      // New Strava fields
      row.privateNote, row.elapsedTime, row.maxSpeed, row.avgCadence, row.temp,
      row.workoutType, row.trainer, row.prCount, row.athleteCount
    ]);
  });

  mergedSheet.getRange(1, 1, outputArray.length, outputArray[0].length).setValues(outputArray);
  mergedSheet.getRange(2, 2, outputArray.length, 1).setNumberFormat("M/d/yyyy h:mm am/pm");

  console.log(`Merge Complete: ${processedStrava.length} Strava activities. Added ${processedFitbit.length} unique Fitbit activities. Skipped ${duplicatesSkipped} duplicate Fitbit activities.`);
}


// --- HELPER FUNCTIONS ---

function cleanNumber(val) {
  if (!val && val !== 0) return "";
  if (typeof val === 'string') {
    const cleaned = parseFloat(val.replace(/[^0-9.-]/g, ''));
    return isNaN(cleaned) ? "" : cleaned;
  }
  return val;
}

function parseStravaDateTime(dateStr, timeStr) {
  let d = new Date(dateStr);

  if (timeStr instanceof Date) {
    d.setHours(timeStr.getHours(), timeStr.getMinutes(), 0);
  } else if (typeof timeStr === 'string') {
    let timeMatch = timeStr.match(/(\d+):(\d+)\s*(AM|PM)/i);
    if (timeMatch) {
      let hours = parseInt(timeMatch[1], 10);
      let mins  = parseInt(timeMatch[2], 10);
      if (timeMatch[3].toUpperCase() === 'PM' && hours < 12) hours += 12;
      if (timeMatch[3].toUpperCase() === 'AM' && hours === 12) hours = 0;
      d.setHours(hours, mins, 0);
    }
  }
  return d;
}

function parseDurationToMinutes(val) {
  if (!val) return 0;
  if (val instanceof Date) {
    return (val.getHours() * 60) + val.getMinutes() + (val.getSeconds() / 60);
  } else if (typeof val === 'string' && val.includes(':')) {
    let parts = val.split(':').map(Number);
    if (parts.length === 3) return (parts[0] * 60) + parts[1] + (parts[2] / 60);
    if (parts.length === 2) return parts[0] + (parts[1] / 60);
  }
  return parseFloat(val) || 0;
}
