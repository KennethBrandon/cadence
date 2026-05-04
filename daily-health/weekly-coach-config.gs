// =============================================================================
// WEEKLY COACH — CONFIG & CONTEXT
//
// Constants, athlete profile, and the context builders that emit the static
// "who is this person" + population norms blocks consumed by the prompt.
// =============================================================================

// Config is stored in Script Properties (Project Settings → Script Properties).
// Required keys: GEMINI_API_KEY, WORKOUT_SS_ID, REPORT_EMAIL.
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
const WORKOUT_SS_ID = PropertiesService.getScriptProperties().getProperty('WORKOUT_SS_ID');
const REPORT_EMAIL = PropertiesService.getScriptProperties().getProperty('REPORT_EMAIL');

// Athlete profile — stored as JSON in Script Properties (Project Settings → Script
// Properties). Required key: ATHLETE_PROFILE. Expected shape:
//   { dob, sex, heightInches, primarySports[], maxHR, injuryHistory[], goals[] }
// Age is computed from dob at runtime so it stays current.
const ATHLETE_PROFILE = JSON.parse(
  PropertiesService.getScriptProperties().getProperty('ATHLETE_PROFILE') || 'null'
);

// Normative population ranges, device-calibrated where it matters.
// HRV ranges use Fitbit's nightly rMSSD (measured during sleep), which runs
// lower than waking-spot-check numbers — don't compare these to Oura/Whoop.
// RHR: American Heart Association / Mayo Clinic tables.
// Sleep: National Sleep Foundation.
// Keyed by sex + 10-year age bracket. Extend as needed.
const NORMATIVE_RANGES = {
  male: {
    '40-49': {
      // Fitbit nightly HRV (rMSSD), ms — approximate bands from Fitbit's
      // published aggregate data for this age bracket.
      hrv:      { low: 22, belowAvg: 28, average: 35, aboveAvg: 45, excellent: 55 },
      // Resting HR (bpm), AHA age 40-49 male
      rhr:      { athlete: 54, excellent: 61, good: 65, aboveAvg: 69, average: 73, belowAvg: 77 },
      // Sleep duration (hrs), NSF adult recommendation
      sleepHrs: { min: 7, max: 9 },
      // Sleep stage %, standard clinical ranges
      deepPct:  { min: 13, max: 23 },
      remPct:   { min: 20, max: 25 },
    },
  },
};

function buildNormativeContext() {
  const p = ATHLETE_PROFILE;
  const dob = new Date(p.dob + 'T12:00:00');
  const now = new Date();
  let age = now.getFullYear() - dob.getFullYear();
  const beforeBirthday =
    now.getMonth() < dob.getMonth() ||
    (now.getMonth() === dob.getMonth() && now.getDate() < dob.getDate());
  if (beforeBirthday) age -= 1;

  const bracket = `${Math.floor(age / 10) * 10}-${Math.floor(age / 10) * 10 + 9}`;
  const ranges = NORMATIVE_RANGES[p.sex]?.[bracket];
  if (!ranges) return `No normative ranges available for ${p.sex} age ${bracket}.`;

  return [
    `Population norms for ${p.sex}s age ${bracket} (use for context, not as personal targets — athlete's 90-day baseline is the truer personal reference):`,
    `  HRV (Fitbit nightly rMSSD, ms): low <${ranges.hrv.low} · below-avg ${ranges.hrv.low}-${ranges.hrv.belowAvg} · average ${ranges.hrv.belowAvg}-${ranges.hrv.average} · above-avg ${ranges.hrv.average}-${ranges.hrv.aboveAvg} · excellent >${ranges.hrv.excellent}`,
    `  Resting HR (bpm): athlete <${ranges.rhr.athlete} · excellent ${ranges.rhr.athlete}-${ranges.rhr.excellent} · good ${ranges.rhr.excellent}-${ranges.rhr.good} · average ${ranges.rhr.good}-${ranges.rhr.average} · poor >${ranges.rhr.belowAvg}`,
    `  Sleep duration: ${ranges.sleepHrs.min}-${ranges.sleepHrs.max} hrs recommended`,
    `  Deep sleep: ${ranges.deepPct.min}-${ranges.deepPct.max}% of total sleep`,
    `  REM sleep: ${ranges.remPct.min}-${ranges.remPct.max}% of total sleep`,
  ].join('\n');
}

function buildAthleteContext() {
  const p = ATHLETE_PROFILE;
  const dob = new Date(p.dob + 'T12:00:00');
  const now = new Date();
  let age = now.getFullYear() - dob.getFullYear();
  const beforeBirthday =
    now.getMonth() < dob.getMonth() ||
    (now.getMonth() === dob.getMonth() && now.getDate() < dob.getDate());
  if (beforeBirthday) age -= 1;

  const feet = Math.floor(p.heightInches / 12);
  const inches = p.heightInches % 12;

  const lines = [
    `Age: ${age} (${p.sex})`,
    `Height: ${feet}'${inches}" (${p.heightInches} in)`,
    `Primary sports: ${p.primarySports.join(', ')}`,
    `Max HR: ${p.maxHR} bpm`,
    `Injury history / limitations:`,
    ...p.injuryHistory.map(i => `  - ${i}`),
    `Current goals:`,
    ...p.goals.map(g => `  - ${g}`),
  ];
  return lines.join('\n');
}


// Compute what the athlete actually does on a typical week. Grounds the AI's
// recommendations in reality — without this, it will suggest "cap hikes at 4
// mi" to someone whose baseline is 8 mi/hike for 2 years. Next Week bullets
// should be small deltas from baseline, not arbitrary caps.
function buildAthleteBaseline(ninetyDayHealth, ninetyDayWorkouts) {
  const health = ninetyDayHealth || [];
  const workouts = ninetyDayWorkouts || [];
  const weeksOfData = Math.max(1, Math.round(health.length / 7));

  // Helpers scoped to this fn — we want min/max/median, not just mean.
  const median = arr => {
    const s = arr.slice().sort((a, b) => a - b);
    const n = s.length;
    if (n === 0) return null;
    return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2;
  };
  const avg = arr => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;

  // Group workouts by type family. Columns: Date/Time=1, Type=2, Name=3,
  // Duration=4, Distance=5, Max HR=11.
  const byType = { hike: [], run: [], ride: [], strength: [], other: [] };
  workouts.forEach(r => {
    const t = String(r[2] || '').toLowerCase();
    let bucket = 'other';
    if (t.includes('hike')) bucket = 'hike';
    else if (t.includes('run')) bucket = 'run';
    else if (t.includes('ride') || t.includes('bike') || t.includes('cycle')) bucket = 'ride';
    else if (t.includes('strength') || t.includes('weight')) bucket = 'strength';
    byType[bucket].push({
      duration: parseFloat(r[4]) || 0,
      distance: parseFloat(r[5]) || 0,
      maxHR:    parseFloat(r[11]) || 0,
    });
  });

  const lines = [`Based on the last ~${weeksOfData} weeks of tracking:`];

  const describeActivity = (label, arr) => {
    if (!arr.length) return null;
    const perWeek = (arr.length / weeksOfData).toFixed(1);
    const dists = arr.map(x => x.distance).filter(v => v > 0);
    const durs  = arr.map(x => x.duration).filter(v => v > 0);
    const parts = [`~${perWeek}/week`];
    if (dists.length) {
      parts.push(`typical distance ${median(dists).toFixed(1)} mi (range ${Math.min(...dists).toFixed(1)}–${Math.max(...dists).toFixed(1)})`);
    }
    if (durs.length) {
      parts.push(`typical duration ${Math.round(median(durs))} min (range ${Math.round(Math.min(...durs))}–${Math.round(Math.max(...durs))})`);
    }
    return `${label}: ${parts.join(', ')}`;
  };

  ['hike', 'run', 'ride', 'strength'].forEach(k => {
    const line = describeActivity(k.charAt(0).toUpperCase() + k.slice(1) + 's', byType[k]);
    if (line) lines.push(`  ${line}`);
  });

  // Health baselines
  const sleeps = health.map(r => parseFloat(r[7]) || 0).filter(v => v > 0);
  const steps  = health.map(r => parseFloat(r[12]) || 0).filter(v => v > 0);
  const avgSleepMin = avg(sleeps);
  const avgSteps   = avg(steps);
  if (avgSleepMin) lines.push(`  Sleep: typical ${(avgSleepMin / 60).toFixed(1)} hrs/night`);
  if (avgSteps)    lines.push(`  Steps: typical ${Math.round(avgSteps).toLocaleString()} per day`);

  return lines.join('\n');
}
