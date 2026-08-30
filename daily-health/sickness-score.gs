// =============================================================================
// SICKNESS SIGNAL — Daily 0-100 illness probability from overnight metrics
//
// Detects the systemic (feverish) illness signature: RHR up, HRV down, skin
// temp up, respiratory rate up, SpO2 down — each measured against the
// athlete's own trailing 14-day baseline, never a population norm.
//
// Backtested against 933 days (Jan 2024 – Aug 2026) with illnesses mined from
// the athlete's own Strava notes as ground truth: the only two days ever
// scoring ≥ 50 were both real, diary-confirmed illnesses; hard-workout
// look-alikes (e.g. a 2.4 hr kayak day) are attenuated by the training gate.
// Known blind spots: GI bugs without fever and mild colds that never move
// overnight physiology.
//
// Components (weighted, renormalized over whichever signals exist that day):
//   RHR above baseline        — 25%  (ramps +1.0 → +4.5 bpm)
//   HRV below baseline        — 25%  (ramps −10% → −35%)
//   Skin temp above baseline  — 25%  (ramps +0.15 → +0.60 °C)
//   Resp rate above baseline  — 15%  (ramps +0.4 → +1.5 br/min)
//   SpO2 below baseline       — 10%  (ramps −1.0 → −3.0 pts)
//
// Training gate: if the PREVIOUS day had ≥ 120 active-zone minutes or
// ≥ 18,000 steps, today's evidence is cut to 60% — a hard workout is the
// innocent explanation and gets attributed first.
//
// Persistence: prob = 100 × (0.65 × today + 0.35 × yesterday). One bad night
// tops out around 65; crossing 50 effectively requires two consecutive
// elevated mornings. Viruses build; hot bedrooms don't.
//
// Thresholds: ≥ 50 likely fighting something · 35–49 watch · < 35 quiet.
//
// Writes "Sickness Prob" column (Y=25) to the Daily Data sheet. Runs inside
// the same script lock as computeReadinessScores (chained from there).
// =============================================================================

const SICKNESS_COL       = 25;   // col Y — sync owns 1-24, readiness owns 17
const SICKNESS_BASE_WIN  = 14;   // trailing days for baselines
const SICKNESS_MIN_BASE  = 5;    // min valid days before a baseline counts

// Column indices in the Daily Data row array (0-based)
const SICK_IDX = { rhr: 2, hrv: 3, steps: 12, spo2: 13, skin: 17, azm: 20, resp: 23 };


// Run standalone (manual/backfill). Normal operation is the chained call from
// computeReadinessScores, which already holds the lock.
function computeSicknessScores() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    console.warn('Skipping sickness compute: could not acquire script lock.');
    return;
  }
  try {
    _computeSicknessScoresLocked();
  } finally {
    lock.releaseLock();
  }
}

function _computeSicknessScoresLocked() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Daily Data');
  if (!sheet || sheet.getLastRow() < 2) {
    console.log('No Daily Data found.');
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data[0][SICKNESS_COL - 1] !== 'Sickness Prob') {
    sheet.getRange(1, SICKNESS_COL).setValue('Sickness Prob');
  }

  // Sort chronologically without trusting sheet order, keeping sheet rows.
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const raw = data[i][0];
    if (!raw) continue;
    const date = raw instanceof Date ? raw : new Date(String(raw).substring(0, 10) + 'T12:00:00');
    if (isNaN(date.getTime())) continue;
    rows.push({
      sheetRow: i + 1,
      date:  date,
      rhr:   toNum(data[i][SICK_IDX.rhr]),
      hrv:   toNum(data[i][SICK_IDX.hrv]),
      skin:  toNum(data[i][SICK_IDX.skin]),
      resp:  toNum(data[i][SICK_IDX.resp]),
      spo2:  toNum(data[i][SICK_IDX.spo2]),
      azm:   toNum(data[i][SICK_IDX.azm]),
      steps: toNum(data[i][SICK_IDX.steps])
    });
  }
  rows.sort((a, b) => a.date - b.date);

  const WEIGHTS = { rhr: 0.25, hrv: 0.25, skin: 0.25, resp: 0.15, spo2: 0.10 };
  const scoreBySheetRow = {};
  let prevE = 0;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const window = rows.slice(Math.max(0, i - SICKNESS_BASE_WIN), i);

    const base = {};
    for (const k in WEIGHTS) {
      const vals = window.map(w => w[k]).filter(v => v !== null);
      base[k] = vals.length >= SICKNESS_MIN_BASE
        ? vals.reduce((a, b) => a + b, 0) / vals.length
        : null;
    }

    const ev = {};
    if (r.rhr  !== null && base.rhr  !== null) ev.rhr  = sicknessRamp_(r.rhr - base.rhr, 1.0, 4.5);
    if (r.hrv  !== null && base.hrv  !== null) ev.hrv  = sicknessRamp_((base.hrv - r.hrv) / base.hrv, 0.10, 0.35);
    if (r.skin !== null && base.skin !== null) ev.skin = sicknessRamp_(r.skin - base.skin, 0.15, 0.60);
    if (r.resp !== null && base.resp !== null) ev.resp = sicknessRamp_(r.resp - base.resp, 0.4, 1.5);
    if (r.spo2 !== null && base.spo2 !== null) ev.spo2 = sicknessRamp_(base.spo2 - r.spo2, 1.0, 3.0);

    // The two core signals are required; otherwise no score for the day.
    if (!('rhr' in ev) || !('hrv' in ev)) {
      scoreBySheetRow[r.sheetRow] = '';
      prevE = 0;
      continue;
    }

    let wsum = 0, esum = 0;
    for (const k in ev) { wsum += WEIGHTS[k]; esum += WEIGHTS[k] * ev[k]; }
    let E = esum / wsum;

    // Training gate — big prior-day effort attenuates today's evidence
    if (i > 0) {
      const p = rows[i - 1];
      if ((p.azm !== null && p.azm >= 120) || (p.steps !== null && p.steps >= 18000)) {
        E *= 0.6;
      }
    }

    scoreBySheetRow[r.sheetRow] = Math.round(100 * (0.65 * E + 0.35 * prevE));
    prevE = E;
  }

  // Write back in sheet order (blank for header-adjacent rows never scored)
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const v = scoreBySheetRow[i + 1];
    out.push([v === undefined ? '' : v]);
  }
  if (out.length > 0) {
    sheet.getRange(2, SICKNESS_COL, out.length, 1).setValues(out);
  }

  // Conditional formatting: red = likely (≥50), yellow = watch (35–49)
  const scoreRange = sheet.getRange(2, SICKNESS_COL, out.length, 1);
  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(50)
    .setBackground('#f4cccc')
    .setRanges([scoreRange])
    .build();
  const yellowRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(35, 49)
    .setBackground('#fff2cc')
    .setRanges([scoreRange])
    .build();
  const filtered = sheet.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().some(r => r.getColumn() === SICKNESS_COL));
  filtered.push(redRule, yellowRule);
  sheet.setConditionalFormatRules(filtered);

  // rows is chronological; sheet order may not be
  const scoredRows = rows.filter(r => scoreBySheetRow[r.sheetRow] !== '');
  if (scoredRows.length > 0) {
    const latest = scoredRows[scoredRows.length - 1];
    console.log(`Sickness signal computed for ${scoredRows.length} days. Latest (${latest.date.toISOString().substring(0, 10)}): ${scoreBySheetRow[latest.sheetRow]}`);
  }
}

function sicknessRamp_(x, lo, hi) {
  if (x === null || isNaN(x)) return 0;
  if (x <= lo) return 0;
  if (x >= hi) return 1;
  return (x - lo) / (hi - lo);
}
