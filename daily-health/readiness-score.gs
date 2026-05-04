// =============================================================================
// READINESS SCORE — Daily 1-100 score from recovery metrics
//
// Uses 14-day rolling baselines so the score adapts to YOUR personal trends.
// No hardcoded thresholds — everything is relative to your own recent data.
//
// Components (weighted):
//   HRV vs 14-day avg     — 30%  (higher = better)
//   RHR vs 14-day avg     — 25%  (lower = better)
//   Sleep duration         — 25%  (vs 7.5hr target)
//   Deep sleep %           — 10%  (% of total sleep in deep stage)
//   SpO2                   — 10%  (deviation from personal baseline)
//
// Writes "Readiness" column (Q=17) to the Daily Data sheet.
// =============================================================================

const SLEEP_TARGET_MINS = 450;  // 7 hours 30 minutes
const BASELINE_WINDOW   = 14;   // days for rolling average


// =============================================================================
// MAIN — call manually or add to runDailyHealthSync()
// =============================================================================
function computeReadinessScores() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Daily Data');
  if (!sheet || sheet.getLastRow() < 2) {
    console.log("No Daily Data found.");
    return;
  }

  // -------------------------------------------------------------------------
  // 1. Read all daily data
  // -------------------------------------------------------------------------
  const data = sheet.getDataRange().getValues();
  const headerRow = data[0];

  // Add header if not present
  if (headerRow.length < 17 || headerRow[16] !== 'Readiness') {
    sheet.getRange(1, 17).setValue('Readiness');
  }

  // Extract columns into arrays for easier rolling-window math
  // Indices: RHR=2, HRV=3, TotalSleep=7, Deep=9, SpO2=13
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    rows.push({
      row:    i + 1,  // 1-based sheet row
      rhr:    toNum(data[i][2]),
      hrv:    toNum(data[i][3]),
      sleep:  toNum(data[i][7]),
      deep:   toNum(data[i][9]),
      spo2:   toNum(data[i][13])
    });
  }

  // -------------------------------------------------------------------------
  // 2. Compute rolling baselines and score each day
  // -------------------------------------------------------------------------
  const scores = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];

    // Need at least some data to score
    if (!r.hrv && !r.rhr && !r.sleep) {
      scores.push([""]);
      continue;
    }

    // Compute 14-day rolling averages (excluding current day)
    const windowStart = Math.max(0, i - BASELINE_WINDOW);
    const window = rows.slice(windowStart, i);

    const avgHRV  = rollingAvg(window, 'hrv');
    const avgRHR  = rollingAvg(window, 'rhr');
    const avgSpO2 = rollingAvg(window, 'spo2');

    // --- HRV Component (30%) ---
    // Score 50 at baseline, +/- based on deviation
    // Each 1ms above baseline = ~2 points, capped at 100
    let hrvScore = 50;
    if (r.hrv && avgHRV) {
      const deviation = r.hrv - avgHRV;
      // Scale: +-10ms from baseline spans the full 0-100 range
      hrvScore = 50 + (deviation / 10) * 50;
    } else if (r.hrv) {
      hrvScore = 50;  // no baseline yet, neutral
    } else {
      hrvScore = null;
    }

    // --- RHR Component (25%) ---
    // Lower is better. Score 50 at baseline.
    let rhrScore = 50;
    if (r.rhr && avgRHR) {
      const deviation = avgRHR - r.rhr;  // positive when RHR is below avg (good)
      // Scale: +-5 bpm from baseline spans the full range
      rhrScore = 50 + (deviation / 5) * 50;
    } else if (r.rhr) {
      rhrScore = 50;
    } else {
      rhrScore = null;
    }

    // --- Sleep Duration Component (25%) ---
    // Score based on how close to target (450 min = 7.5 hrs)
    let sleepScore = null;
    if (r.sleep) {
      // At target = 80, above target = up to 100, below = scales down
      // 6hrs (360min) = ~40, 5hrs (300min) = ~10
      const ratio = r.sleep / SLEEP_TARGET_MINS;
      sleepScore = Math.min(100, ratio * 80);
    }

    // --- Deep Sleep Component (10%) ---
    // Healthy deep sleep is ~15-20% of total. Score based on percentage.
    let deepScore = null;
    if (r.deep && r.sleep && r.sleep > 0) {
      const deepPct = (r.deep / r.sleep) * 100;
      // 20% = score 90, 15% = 70, 10% = 50, 5% = 25
      deepScore = Math.min(100, deepPct * 4.5);
    }

    // --- SpO2 Component (10%) ---
    // Normally very stable (94-96%). Flag meaningful drops.
    let spo2Score = 80;  // default high since usually stable
    if (r.spo2 && avgSpO2) {
      const deviation = r.spo2 - avgSpO2;
      // Each 0.5% drop below baseline is significant
      spo2Score = 80 + (deviation / 1) * 40;
    } else if (r.spo2) {
      // No baseline — score based on absolute value
      // 95+ = great, 93 = ok, <92 = concerning
      spo2Score = Math.min(100, (r.spo2 - 88) * 12.5);
    }

    // --- Weighted composite ---
    let totalWeight = 0;
    let weightedSum = 0;

    if (hrvScore !== null)   { weightedSum += hrvScore * 0.30;   totalWeight += 0.30; }
    if (rhrScore !== null)   { weightedSum += rhrScore * 0.25;   totalWeight += 0.25; }
    if (sleepScore !== null) { weightedSum += sleepScore * 0.25; totalWeight += 0.25; }
    if (deepScore !== null)  { weightedSum += deepScore * 0.10;  totalWeight += 0.10; }
    if (spo2Score !== null)  { weightedSum += spo2Score * 0.10;  totalWeight += 0.10; }

    if (totalWeight > 0) {
      const raw = weightedSum / totalWeight;
      const clamped = Math.max(1, Math.min(100, Math.round(raw)));
      scores.push([clamped]);
    } else {
      scores.push([""]);
    }
  }

  // -------------------------------------------------------------------------
  // 3. Write scores to column Q (17)
  // -------------------------------------------------------------------------
  if (scores.length > 0) {
    sheet.getRange(2, 17, scores.length, 1).setValues(scores);
  }

  // Conditional formatting for readiness column
  const scoreRange = sheet.getRange(2, 17, scores.length, 1);

  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(70)
    .setBackground('#d9ead3')
    .setRanges([scoreRange])
    .build();

  const yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(40, 69)
    .setBackground('#fff2cc')
    .setRanges([scoreRange])
    .build();

  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(40)
    .setBackground('#f4cccc')
    .setRanges([scoreRange])
    .build();

  // Preserve existing rules and add ours
  const existingRules = sheet.getConditionalFormatRules();
  // Remove old readiness rules if re-running
  const filtered = existingRules.filter(rule => {
    const ranges = rule.getRanges();
    return !ranges.some(r => r.getColumn() === 17);
  });
  filtered.push(greenRule, yellowRule, redRule);
  sheet.setConditionalFormatRules(filtered);

  // Log summary
  const validScores = scores.filter(s => s[0] !== "").map(s => s[0]);
  if (validScores.length > 0) {
    const avg = validScores.reduce((a, b) => a + b, 0) / validScores.length;
    const latest = validScores[validScores.length - 1];
    console.log(`Readiness computed for ${validScores.length} days. Latest: ${latest}, Average: ${avg.toFixed(0)}`);
  }
}


// =============================================================================
// UTILITIES
// =============================================================================
function toNum(val) {
  if (val === "" || val === null || val === undefined) return null;
  const n = parseFloat(val);
  return isNaN(n) ? null : n;
}

function rollingAvg(window, field) {
  const vals = window.map(r => r[field]).filter(v => v !== null);
  if (vals.length === 0) return null;
  return vals.reduce((a, b) => a + b, 0) / vals.length;
}
