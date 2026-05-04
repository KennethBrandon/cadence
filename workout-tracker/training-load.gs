// =============================================================================
// TRAINING LOAD — CTL / ATL / TSB (Fitness / Fatigue / Form)
//
// Reads Strava suffer scores, computes daily training stress,
// then calculates exponentially-weighted moving averages:
//   CTL (Chronic Training Load)  = 42-day decay  ("Fitness")
//   ATL (Acute Training Load)    = 7-day decay    ("Fatigue")
//   TSB (Training Stress Balance) = CTL - ATL      ("Form")
//
// Positive TSB = fresh/peaked.  Negative TSB = fatigued.
// TSB between -10 and +10 is the typical "ready to race" zone.
// =============================================================================


// =============================================================================
// MAIN — call manually or add to runDailySync()
// =============================================================================
function computeTrainingLoad() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workoutsSheet = ss.getSheetByName('Strava Workouts');
  if (!workoutsSheet || workoutsSheet.getLastRow() < 2) {
    console.log("No Strava Workouts data found.");
    return;
  }

  // -------------------------------------------------------------------------
  // 1. Read all workouts and extract date + suffer score
  // -------------------------------------------------------------------------
  const data = workoutsSheet.getDataRange().getValues();
  const dailyTSS = {};  // { "yyyy-MM-dd": totalSufferScore }

  for (let i = 1; i < data.length; i++) {
    const dateVal  = data[i][1];  // B: Date  (formatted "EEE, MMM dd, yyyy")
    const sufferVal = data[i][16]; // Q: Suffer Score

    if (!dateVal) continue;

    // Parse the date string into yyyy-MM-dd
    const dateStr = parseDateToISO(dateVal);
    if (!dateStr) continue;

    const score = parseFloat(sufferVal) || 0;
    dailyTSS[dateStr] = (dailyTSS[dateStr] || 0) + score;
  }

  // -------------------------------------------------------------------------
  // 2. Build a continuous date range from first workout to today
  // -------------------------------------------------------------------------
  const sortedDates = Object.keys(dailyTSS).sort();
  if (sortedDates.length === 0) {
    console.log("No valid workout dates found.");
    return;
  }

  const startDate = new Date(sortedDates[0] + 'T12:00:00');
  const today     = new Date();
  today.setHours(12, 0, 0, 0);

  const allDates = [];
  const d = new Date(startDate);
  while (d <= today) {
    allDates.push(Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    d.setDate(d.getDate() + 1);
  }

  // -------------------------------------------------------------------------
  // 3. Compute CTL / ATL / TSB using exponential decay
  //    CTL_today = CTL_yesterday * (1 - 1/42) + TSS_today / 42
  //    ATL_today = ATL_yesterday * (1 - 1/7)  + TSS_today / 7
  //    TSB = CTL - ATL
  // -------------------------------------------------------------------------
  const CTL_DAYS = 42;
  const ATL_DAYS = 7;

  let ctl = 0;
  let atl = 0;

  const rows = [];

  for (const dateStr of allDates) {
    const tss = dailyTSS[dateStr] || 0;

    ctl = ctl * (1 - 1 / CTL_DAYS) + tss / CTL_DAYS;
    atl = atl * (1 - 1 / ATL_DAYS) + tss / ATL_DAYS;
    const tsb = ctl - atl;

    rows.push([
      dateStr,
      Math.round(tss),
      parseFloat(ctl.toFixed(1)),
      parseFloat(atl.toFixed(1)),
      parseFloat(tsb.toFixed(1))
    ]);
  }

  // -------------------------------------------------------------------------
  // 4. Write to "Training Load" sheet (newest first)
  // -------------------------------------------------------------------------
  let sheet = ss.getSheetByName('Training Load');
  if (!sheet) {
    sheet = ss.insertSheet('Training Load');
  } else {
    sheet.clear();
  }

  // Reverse so newest days are at the top
  rows.reverse();

  const headers = ['Date', 'Daily TSS', 'CTL (Fitness)', 'ATL (Fatigue)', 'TSB (Form)'];
  sheet.appendRow(headers);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  // Bold + freeze header
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // --- Conditional formatting ---
  const rules = [];

  // Daily TSS (col B) — intensity of training day
  const tssRange = sheet.getRange(2, 2, rows.length, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(80)
    .setBackground('#f4cccc').setRanges([tssRange]).build());   // very hard day
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(40, 79)
    .setBackground('#fff2cc').setRanges([tssRange]).build());   // moderate day
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 39)
    .setBackground('#d9ead3').setRanges([tssRange]).build());   // easy day

  // CTL / Fitness (col C) — higher is fitter, gradient
  const ctlRange = sheet.getRange(2, 3, rows.length, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(20)
    .setBackground('#b6d7a8').setRanges([ctlRange]).build());   // strong fitness
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(10, 19.9)
    .setBackground('#d9ead3').setRanges([ctlRange]).build());   // moderate fitness
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 9.9)
    .setBackground('#fff2cc').setRanges([ctlRange]).build());   // low fitness

  // ATL / Fatigue (col D) — higher means more tired
  const atlRange = sheet.getRange(2, 4, rows.length, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(20)
    .setBackground('#f4cccc').setRanges([atlRange]).build());   // high fatigue
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(10, 19.9)
    .setBackground('#fff2cc').setRanges([atlRange]).build());   // moderate fatigue
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(1, 9.9)
    .setBackground('#d9ead3').setRanges([atlRange]).build());   // low fatigue

  // TSB / Form (col E) — positive = fresh, negative = fatigued
  const tsbRange = sheet.getRange(2, 5, rows.length, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#d9ead3').setRanges([tsbRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(-10, 0)
    .setBackground('#fff2cc').setRanges([tsbRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(-10)
    .setBackground('#f4cccc').setRanges([tsbRange]).build());

  sheet.setConditionalFormatRules(rules);

  // Legend and today's summary at top-right
  const latest = rows[0];  // newest day is now first

  sheet.getRange(1, 7).setValue('Today');
  sheet.getRange(1, 7, 1, 2).setFontWeight('bold');
  sheet.getRange(2, 7).setValue('Fitness (CTL)');
  sheet.getRange(2, 8).setValue(latest[2]);
  sheet.getRange(3, 7).setValue('Fatigue (ATL)');
  sheet.getRange(3, 8).setValue(latest[3]);
  sheet.getRange(4, 7).setValue('Form (TSB)');
  sheet.getRange(4, 8).setValue(latest[4]);

  // Legend — Metric Definitions
  sheet.getRange(6, 7).setValue('Metrics');
  sheet.getRange(6, 7, 1, 2).setFontWeight('bold');

  sheet.getRange(7, 7).setValue('Daily TSS');
  sheet.getRange(7, 8).setValue('Training stress for the day (sum of Strava suffer scores).');
  sheet.getRange(8, 7).setValue('CTL (Fitness)');
  sheet.getRange(8, 8).setValue('42-day training load. Higher = fitter. Builds slowly.');
  sheet.getRange(9, 7).setValue('ATL (Fatigue)');
  sheet.getRange(9, 8).setValue('7-day training load. Spikes after hard efforts. Drops with rest.');
  sheet.getRange(10, 7).setValue('TSB (Form)');
  sheet.getRange(10, 8).setValue('CTL minus ATL. Your freshness balance.');

  // Legend — Color Key
  sheet.getRange(12, 7).setValue('Color Key');
  sheet.getRange(12, 7, 1, 2).setFontWeight('bold');

  // Daily TSS colors
  sheet.getRange(13, 7).setValue('TSS 1-39');
  sheet.getRange(13, 8).setValue('Easy day');
  sheet.getRange(13, 7).setBackground('#d9ead3');
  sheet.getRange(14, 7).setValue('TSS 40-79');
  sheet.getRange(14, 8).setValue('Moderate effort');
  sheet.getRange(14, 7).setBackground('#fff2cc');
  sheet.getRange(15, 7).setValue('TSS 80+');
  sheet.getRange(15, 8).setValue('Very hard day');
  sheet.getRange(15, 7).setBackground('#f4cccc');

  // CTL colors
  sheet.getRange(17, 7).setValue('CTL 20+');
  sheet.getRange(17, 8).setValue('Strong fitness');
  sheet.getRange(17, 7).setBackground('#b6d7a8');
  sheet.getRange(18, 7).setValue('CTL 10-20');
  sheet.getRange(18, 8).setValue('Moderate fitness');
  sheet.getRange(18, 7).setBackground('#d9ead3');
  sheet.getRange(19, 7).setValue('CTL 1-10');
  sheet.getRange(19, 8).setValue('Low fitness base');
  sheet.getRange(19, 7).setBackground('#fff2cc');

  // ATL colors
  sheet.getRange(21, 7).setValue('ATL 1-10');
  sheet.getRange(21, 8).setValue('Well rested');
  sheet.getRange(21, 7).setBackground('#d9ead3');
  sheet.getRange(22, 7).setValue('ATL 10-20');
  sheet.getRange(22, 8).setValue('Moderate fatigue');
  sheet.getRange(22, 7).setBackground('#fff2cc');
  sheet.getRange(23, 7).setValue('ATL 20+');
  sheet.getRange(23, 8).setValue('High fatigue — heavy training block');
  sheet.getRange(23, 7).setBackground('#f4cccc');

  // TSB colors
  sheet.getRange(25, 7).setValue('TSB > 0');
  sheet.getRange(25, 8).setValue('Fresh / peaked — ready to perform');
  sheet.getRange(25, 7).setBackground('#d9ead3');
  sheet.getRange(26, 7).setValue('TSB -10 to 0');
  sheet.getRange(26, 8).setValue('Productive training zone');
  sheet.getRange(26, 7).setBackground('#fff2cc');
  sheet.getRange(27, 7).setValue('TSB < -10');
  sheet.getRange(27, 8).setValue('Deep fatigue — risk of overtraining, consider rest');
  sheet.getRange(27, 7).setBackground('#f4cccc');

  // Auto-resize legend columns
  sheet.autoResizeColumn(7);
  sheet.autoResizeColumn(8);

  // -------------------------------------------------------------------------
  // 5. Chart — CTL, ATL, TSB over time
  // -------------------------------------------------------------------------
  // Remove existing charts so we don't duplicate on re-run
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  const chartDataRows = rows.length + 1; // include header

  const chart = sheet.newChart()
    .asLineChart()
    .addRange(sheet.getRange(1, 1, chartDataRows, 1))  // Date (A)
    .addRange(sheet.getRange(1, 3, chartDataRows, 1))  // CTL (C)
    .addRange(sheet.getRange(1, 4, chartDataRows, 1))  // ATL (D)
    .addRange(sheet.getRange(1, 5, chartDataRows, 1))  // TSB (E)
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setNumHeaders(1)
    .setTitle('Training Load — Fitness / Fatigue / Form')
    .setCurveStyle(Charts.CurveStyle.SMOOTH)
    .setOption('legend', { position: 'bottom' })
    .setOption('series', {
      0: { color: '#38761d', lineWidth: 2 },  // CTL — dark green
      1: { color: '#e69138', lineWidth: 2 },  // ATL — orange
      2: { color: '#3d85c6', lineWidth: 2 }   // TSB — blue
    })
    .setPosition(29, 7, 0, 0)  // Below the legend
    .build();

  sheet.insertChart(chart);

  console.log(`Training Load computed: ${rows.length} days. CTL=${latest[2]}, ATL=${latest[3]}, TSB=${latest[4]}`);
}


// =============================================================================
// UTILITY: Parse Strava date format to yyyy-MM-dd
// Handles "EEE, MMM dd, yyyy" (e.g. "Sat, Jan 06, 2024") and Date objects
// =============================================================================
function parseDateToISO(dateVal) {
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const str = String(dateVal).trim();

  // Try "EEE, MMM dd, yyyy" format
  const match = str.match(/(\w+),\s+(\w+)\s+(\d+),\s+(\d{4})/);
  if (match) {
    const months = {
      Jan:0, Feb:1, Mar:2, Apr:3, May:4, Jun:5,
      Jul:6, Aug:7, Sep:8, Oct:9, Nov:10, Dec:11
    };
    const m = months[match[2]];
    if (m !== undefined) {
      const d = new Date(parseInt(match[4]), m, parseInt(match[3]));
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  }

  // Try ISO format
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str.substring(0, 10);
  }

  return null;
}
