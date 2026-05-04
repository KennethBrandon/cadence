// =============================================================================
// WEEKLY AI COACH REPORT — Gemini-powered health analysis
//
// Pulls the previous 7 full days of health data + workouts, compares against
// previous periods (prior week + same week last year), and sends
// a personalized coaching report via Gmail.
//
// This file holds the orchestrator (sendWeeklyCoachReport). Supporting code
// lives in sibling files:
//   weekly-coach-config.gs   — constants, ATHLETE_PROFILE, NORMATIVE_RANGES, context builders
//   weekly-coach-lens.gs     — Lens of the Week (pickLensForWeek + signal detection)
//   weekly-coach-angle.gs    — Interesting Stuff Angle (detectInterestingAngle)
//   weekly-coach-history.gs  — Report History sheet I/O
//   weekly-coach-html.gs     — markdownToHtml + chart layout
//   weekly-coach-charts.gs   — buildBarChart / buildStackedChart / buildLineChart
//   weekly-coach-test.gs     — testWeeklyCoachReportForDate / backfillWeeklyReports
// =============================================================================
function sendWeeklyCoachReport(options) {
  options = options || {};
  const referenceDate = options.referenceDate ? new Date(options.referenceDate) : new Date();
  const skipEmail = options.skipEmail === true;

  console.log(`Generating Weekly Coach Report (reference date: ${referenceDate.toDateString()}${skipEmail ? ', skipEmail=true' : ''})...`);

  const tz = Session.getScriptTimeZone();
  const today = referenceDate;

  // Date ranges — previous 7 full days (ending yesterday, excluding today).
  // Time-of-day is normalized so the window doesn't shift based on when the
  // script runs (otherwise e.g. an early-morning workout on the boundary day
  // can fall outside the filter).
  const thisWeekEnd   = new Date(today);
  thisWeekEnd.setDate(today.getDate() - 1);           // yesterday
  thisWeekEnd.setHours(23, 59, 59, 999);              // end of yesterday
  const thisWeekStart = new Date(thisWeekEnd);
  thisWeekStart.setDate(thisWeekEnd.getDate() - 6);   // 7 days back from yesterday
  thisWeekStart.setHours(0, 0, 0, 0);                 // start of that day

  const prevWeekEnd   = new Date(thisWeekStart);
  prevWeekEnd.setDate(prevWeekEnd.getDate() - 1);
  prevWeekEnd.setHours(23, 59, 59, 999);
  const prevWeekStart = new Date(prevWeekEnd);
  prevWeekStart.setDate(prevWeekStart.getDate() - 6);
  prevWeekStart.setHours(0, 0, 0, 0);

  // Historical comparison periods
  const oneYearStart = new Date(thisWeekStart);
  oneYearStart.setFullYear(oneYearStart.getFullYear() - 1);
  const oneYearEnd = new Date(thisWeekEnd);
  oneYearEnd.setFullYear(oneYearEnd.getFullYear() - 1);

  const twoYearStart = new Date(thisWeekStart);
  twoYearStart.setFullYear(twoYearStart.getFullYear() - 2);
  const twoYearEnd = new Date(thisWeekEnd);
  twoYearEnd.setFullYear(twoYearEnd.getFullYear() - 2);

  // Rolling windows for trend context
  const thirtyDayEnd = new Date(thisWeekStart);
  thirtyDayEnd.setDate(thirtyDayEnd.getDate() - 1);
  thirtyDayEnd.setHours(23, 59, 59, 999);
  const thirtyDayStart = new Date(thirtyDayEnd);
  thirtyDayStart.setDate(thirtyDayStart.getDate() - 29);
  thirtyDayStart.setHours(0, 0, 0, 0);

  const ninetyDayEnd = new Date(thisWeekStart);
  ninetyDayEnd.setDate(ninetyDayEnd.getDate() - 1);
  ninetyDayEnd.setHours(23, 59, 59, 999);
  const ninetyDayStart = new Date(ninetyDayEnd);
  ninetyDayStart.setDate(ninetyDayStart.getDate() - 89);
  ninetyDayStart.setHours(0, 0, 0, 0);

  // -------------------------------------------------------------------------
  // 1. Pull health data from Daily Data sheet
  // -------------------------------------------------------------------------
  const healthSS = SpreadsheetApp.getActiveSpreadsheet();
  const dailySheet = healthSS.getSheetByName('Daily Data');
  const dailyData = dailySheet.getDataRange().getValues();

  function getHealthRows(startDate, endDate) {
    return dailyData.slice(1).filter(row => {
      if (!row[0]) return false;
      const d = row[0] instanceof Date
        ? row[0]
        : new Date(String(row[0]).substring(0, 10) + 'T12:00:00');
      return d >= startDate && d <= endDate;
    });
  }

  const thisWeekHealth  = getHealthRows(thisWeekStart, thisWeekEnd);
  const prevWeekHealth  = getHealthRows(prevWeekStart, prevWeekEnd);
  const oneYearHealth   = getHealthRows(oneYearStart, oneYearEnd);
  const twoYearHealth   = getHealthRows(twoYearStart, twoYearEnd);
  const thirtyDayHealth = getHealthRows(thirtyDayStart, thirtyDayEnd);
  const ninetyDayHealth = getHealthRows(ninetyDayStart, ninetyDayEnd);

  // Find personal bests and worsts from all data for context
  const allRows = dailyData.slice(1).filter(r => r[0]);
  const allHRV = allRows.map(r => parseFloat(r[3])).filter(v => v > 0);
  const allRHR = allRows.map(r => parseFloat(r[2])).filter(v => v > 0);
  const allSleep = allRows.map(r => parseFloat(r[7])).filter(v => v > 0);
  const allReadiness = allRows.map(r => parseFloat(r[16])).filter(v => v > 0);
  const allWeight = allRows.map(r => parseFloat(r[1])).filter(v => v > 0);

  const mean = arr => arr.length > 0 ? arr.reduce((a, b) => a + b, 0) / arr.length : null;

  const personalBests = {
    bestHRV: allHRV.length > 0 ? Math.max(...allHRV) : null,
    bestRHR: allRHR.length > 0 ? Math.min(...allRHR) : null,
    bestSleep: allSleep.length > 0 ? Math.max(...allSleep) : null,
    bestReadiness: allReadiness.length > 0 ? Math.max(...allReadiness) : null,
    avgHRV:       mean(allHRV),
    avgRHR:       mean(allRHR),
    avgSleep:     mean(allSleep),
    avgReadiness: mean(allReadiness),
    avgWeight:    mean(allWeight),
    totalDays: allRows.length,
    // Raw arrays for percentile computation in LONGTERM_PROGRESS
    allHRV, allRHR, allSleep, allReadiness, allWeight,
  };

  // First 90 days of tracking — the "starting line" anchor for LONGTERM_PROGRESS.
  // Only meaningful once we have a decent stretch of history past the first period.
  const sortedByDate = [...allRows].sort((a, b) => {
    const da = a[0] instanceof Date ? a[0] : new Date(a[0]);
    const db = b[0] instanceof Date ? b[0] : new Date(b[0]);
    return da - db;
  });
  const firstPeriodHealth = sortedByDate.slice(0, Math.min(90, sortedByDate.length));

  // -------------------------------------------------------------------------
  // 2. Pull training load data
  // -------------------------------------------------------------------------
  const workoutSS = SpreadsheetApp.openById(WORKOUT_SS_ID);
  let trainingLoadSummary = "No training load data available.";
  let tlData = null;  // promoted so the fitness chart can use it later

  const tlSheet = workoutSS.getSheetByName('Training Load');
  if (tlSheet && tlSheet.getLastRow() > 1) {
    tlData = tlSheet.getDataRange().getValues();

    const startStr = Utilities.formatDate(thisWeekStart, tz, 'yyyy-MM-dd');
    const endStr   = Utilities.formatDate(thisWeekEnd, tz, 'yyyy-MM-dd');
    console.log(`Training Load: looking for dates ${startStr} to ${endStr}`);
    console.log(`Training Load sheet has ${tlData.length - 1} rows. First date: ${tlData[1][0]}, Last date: ${tlData[tlData.length - 1][0]}`);

    const tlRows = tlData.slice(1).filter(row => {
      if (!row[0]) return false;
      const dateStr = row[0] instanceof Date
        ? Utilities.formatDate(row[0], tz, 'yyyy-MM-dd')
        : String(row[0]).substring(0, 10);
      return dateStr >= startStr && dateStr <= endStr;
    });

    console.log(`Training Load: found ${tlRows.length} matching rows`);

    if (tlRows.length > 0) {
      const latest = tlRows[0];
      const totalTSS = tlRows.reduce((sum, r) => sum + (parseFloat(r[1]) || 0), 0);
      trainingLoadSummary = `Weekly TSS Total: ${totalTSS}\n` +
        `Current CTL (Fitness): ${latest[2]}\n` +
        `Current ATL (Fatigue): ${latest[3]}\n` +
        `Current TSB (Form): ${latest[4]}\n` +
        `Daily TSS this week: ${tlRows.map(r => {
          const ds = r[0] instanceof Date
            ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')
            : String(r[0]).substring(0, 10);
          return `${ds}: ${r[1]}`;
        }).join(', ')}`;
      console.log("Training Load summary: " + trainingLoadSummary);
    }
  }

  // -------------------------------------------------------------------------
  // 3. Pull workouts from Merged Workouts sheet
  // -------------------------------------------------------------------------
  let thisWeekWorkouts = "No workout data available.";
  let prevWeekWorkouts = "No previous week workout data.";
  // Hoisted so the lens picker can inspect raw workout rows (e.g. to detect a
  // long-run milestone).
  let thisWkWorkouts = [];
  // Hoisted for baseline computation — typical activity volume over 90 days.
  let ninetyDayWorkouts = [];

  const mergedSheet = workoutSS.getSheetByName('Merged Workouts');
  if (mergedSheet && mergedSheet.getLastRow() > 1) {
    const mergedData = mergedSheet.getDataRange().getValues();
    const mergedHeaders = mergedData[0];

    // Columns we never want to send to the AI — IDs/device noise that eat tokens
    // and add no coaching value.
    const SKIP_WORKOUT_FIELDS = new Set(['Strava ID', 'Google Health ID', 'Device']);

    function getWorkoutRows(data, startDate, endDate) {
      return data.slice(1).filter(row => {
        const d = new Date(row[1]);  // Date/Time is column B (index 1)
        return !isNaN(d) && d >= startDate && d <= endDate;
      });
    }

    thisWkWorkouts = getWorkoutRows(mergedData, thisWeekStart, thisWeekEnd);
    const prevWkWorkouts = getWorkoutRows(mergedData, prevWeekStart, prevWeekEnd);
    ninetyDayWorkouts = getWorkoutRows(mergedData, ninetyDayStart, ninetyDayEnd);

    // Full-row workout blocks: one header line + key:value for every non-empty
    // column. This gives the AI access to Description, Private Note, elevation,
    // cadence, etc. — the context that actually explains how a session went.
    function formatWorkouts(rows) {
      if (rows.length === 0) return "No workouts recorded.";
      return rows.map(row => {
        const date = Utilities.formatDate(new Date(row[1]), tz, 'EEE MM/dd h:mm a');
        const type = row[2] || 'Unknown';
        const name = row[3] || '';
        const headerLine = `--- ${date} | ${type}${name ? ' | ' + name : ''} ---`;
        const fields = mergedHeaders.map((h, i) => {
          if (!h || SKIP_WORKOUT_FIELDS.has(h)) return null;
          // Already in the header line
          if (h === 'Date/Time' || h === 'Activity Type' || h === 'Name') return null;
          const val = row[i];
          if (val === '' || val === null || val === undefined) return null;
          return `  ${h}: ${val}`;
        }).filter(Boolean);
        return [headerLine, ...fields].join('\n');
      }).join('\n\n');
    }

    thisWeekWorkouts = formatWorkouts(thisWkWorkouts);
    prevWeekWorkouts = formatWorkouts(prevWkWorkouts);
  }

  // -------------------------------------------------------------------------
  // 4. Format health summaries
  // -------------------------------------------------------------------------
  function summarizeHealth(rows, label) {
    if (rows.length === 0) return `${label}: No data available.`;

    const avg = (arr) => {
      const valid = arr.filter(v => v !== "" && v !== null && v !== undefined && v > 0);
      return valid.length > 0 ? (valid.reduce((a, b) => a + b, 0) / valid.length) : null;
    };

    const weights    = avg(rows.map(r => parseFloat(r[1])  || 0));
    const rhrs       = avg(rows.map(r => parseFloat(r[2])  || 0));
    const hrvs       = avg(rows.map(r => parseFloat(r[3])  || 0));
    const sleeps     = avg(rows.map(r => parseFloat(r[7])  || 0));
    const deeps      = avg(rows.map(r => parseFloat(r[9])  || 0));
    const rems       = avg(rows.map(r => parseFloat(r[10]) || 0));
    const steps      = avg(rows.map(r => parseFloat(r[12]) || 0));
    const cals       = avg(rows.map(r => parseFloat(r[15]) || 0));
    const readiness  = avg(rows.map(r => parseFloat(r[16]) || 0));

    let summary = `${label} (${rows.length} days):\n`;
    if (weights)   summary += `  Weight: ${weights.toFixed(1)} lbs\n`;
    if (rhrs)      summary += `  Resting HR: ${rhrs.toFixed(0)} bpm\n`;
    if (hrvs)      summary += `  HRV: ${hrvs.toFixed(1)} ms\n`;
    if (sleeps)    summary += `  Total Sleep: ${(sleeps / 60).toFixed(1)} hrs\n`;
    if (deeps)     summary += `  Deep Sleep: ${(deeps / 60).toFixed(1)} hrs\n`;
    if (rems)      summary += `  REM Sleep: ${(rems / 60).toFixed(1)} hrs\n`;
    if (steps)     summary += `  Steps: ${Math.round(steps).toLocaleString()}\n`;
    if (cals)      summary += `  Calories: ${Math.round(cals).toLocaleString()}\n`;
    if (readiness) summary += `  Readiness Score: ${readiness.toFixed(0)}/100\n`;

    // Daily readiness breakdown
    if (rows.some(r => r[16])) {
      summary += `  Daily Readiness: `;
      summary += rows.map(r => {
        const dateStr = r[0] instanceof Date
          ? Utilities.formatDate(r[0], tz, 'EEE')
          : String(r[0]).substring(0, 10);
        const score = r[16] || '—';
        return `${dateStr}: ${score}`;
      }).join(', ') + '\n';
    }

    return summary;
  }

  // -------------------------------------------------------------------------
  // 4b. Pre-compute long-horizon distribution context.
  //
  // For each core metric, locate this week's average inside the athlete's
  // 90-day distribution (percentile band) and report a 30-day rolling trend
  // built from three contiguous 30-day bins. This is the antidote to over-
  // reading week-over-week noise — most weeks land in the "normal" band of
  // the 90-day window and don't deserve dramatic framing. The bin trend
  // catches slow drifts that WoW deltas miss.
  //
  // Returns { summary, bands }. `bands` is keyed by metric label and
  // consumed by computeSignificantChanges so each WoW flag carries its
  // 90-day band into the prompt.
  // -------------------------------------------------------------------------
  function computeLongHorizonContext(thisRows, ninetyDayRows) {
    if (!ninetyDayRows || ninetyDayRows.length < 30) {
      return {
        summary: "Long-horizon context unavailable (need 30+ days of 90-day baseline data).",
        bands: {},
      };
    }

    // Linear-interpolated percentile from a sorted ascending array.
    const percentile = (sortedArr, p) => {
      if (sortedArr.length === 0) return null;
      const idx = (sortedArr.length - 1) * (p / 100);
      const lo = Math.floor(idx), hi = Math.ceil(idx);
      if (lo === hi) return sortedArr[lo];
      return sortedArr[lo] + (sortedArr[hi] - sortedArr[lo]) * (idx - lo);
    };

    const avg = arr => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;

    // Sort baseline rows oldest → newest so we can split into three 30-day
    // bins for the rolling trend. Bin 3 is the 30 days right before this
    // week — the most relevant comparison point for "where the trend was
    // pointing heading into this week."
    const sortedNinety = [...ninetyDayRows].sort((a, b) => {
      const da = a[0] instanceof Date ? a[0] : new Date(a[0]);
      const db = b[0] instanceof Date ? b[0] : new Date(b[0]);
      return da - db;
    });
    const binSize = Math.floor(sortedNinety.length / 3);
    const bin1 = sortedNinety.slice(0, binSize);
    const bin3 = sortedNinety.slice(binSize * 2);

    const valsFromCol = (rows, col) =>
      rows.map(r => parseFloat(r[col])).filter(v => isFinite(v) && v > 0);

    // Labels MUST match those used by computeSignificantChanges so band
    // tagging in the WoW block lines up cleanly.
    const metrics = [
      { label: 'HRV',         col: 3,  unit: 'ms',   fmt: v => v.toFixed(1) },
      { label: 'Resting HR',  col: 2,  unit: 'bpm',  fmt: v => v.toFixed(0) },
      { label: 'Total Sleep', col: 7,  unit: 'hrs',  fmt: v => (v / 60).toFixed(1) },
      { label: 'Readiness',   col: 16, unit: '/100', fmt: v => v.toFixed(0) },
      { label: 'Steps',       col: 12, unit: '',     fmt: v => Math.round(v).toLocaleString() },
      { label: 'Weight',      col: 1,  unit: 'lbs',  fmt: v => v.toFixed(1) },
    ];

    const bands = {};
    const lines = metrics.map(m => {
      const thisVals   = valsFromCol(thisRows, m.col);
      const ninetyVals = valsFromCol(sortedNinety, m.col);
      if (thisVals.length === 0 || ninetyVals.length < 10) {
        return `  ${m.label}: insufficient data`;
      }
      const thisVal = avg(thisVals);
      const sorted90 = [...ninetyVals].sort((a, b) => a - b);
      const p10 = percentile(sorted90, 10);
      const p25 = percentile(sorted90, 25);
      const p50 = percentile(sorted90, 50);
      const p75 = percentile(sorted90, 75);
      const p90 = percentile(sorted90, 90);
      const iqr = p75 - p25;

      // Band by percentile cuts at 10/25/75/90.
      let band;
      if      (thisVal < p10)  band = 'extreme_low';
      else if (thisVal < p25)  band = 'low';
      else if (thisVal <= p75) band = 'normal';
      else if (thisVal <= p90) band = 'high';
      else                     band = 'extreme_high';
      bands[m.label] = band;

      // 30-day bin trend. Threshold = 0.5 * IQR so it self-calibrates per
      // metric (steps and HRV vary very differently in absolute terms).
      const bin1Avg = avg(valsFromCol(bin1, m.col));
      const bin3Avg = avg(valsFromCol(bin3, m.col));
      let trendDesc = '90d trend: insufficient bin data';
      if (bin1Avg !== null && bin3Avg !== null && iqr > 0) {
        const delta = bin3Avg - bin1Avg;
        let direction;
        if (Math.abs(delta) < 0.5 * iqr) direction = 'flat';
        else direction = delta > 0 ? 'trending up' : 'trending down';
        trendDesc = `90d trend: ${direction} (${m.fmt(bin1Avg)} → ${m.fmt(bin3Avg)})`;
      }

      return `  ${m.label}: ${m.fmt(thisVal)} ${m.unit} this week | 90d median ${m.fmt(p50)} (IQR ${m.fmt(p25)}–${m.fmt(p75)}) | band: ${band} | ${trendDesc}`;
    });

    const summary = `Where this week's averages sit in the athlete's recent 90-day distribution.
Bands by percentile: extreme_low (<10th) · low (10–25th) · normal (25–75th) · high (75–90th) · extreme_high (>90th).
Use this to decide what deserves the spotlight. A "normal" band value rarely deserves WoW drama — its 90d trend is usually the more honest story.
${lines.join('\n')}`;

    return { summary, bands };
  }

  // -------------------------------------------------------------------------
  // 4c. Pre-compute significant week-over-week changes.
  // We flag only metrics that cleared a per-metric threshold so the AI doesn't
  // have to decide what's "big" — it just renders what we pass in. Each line
  // is tagged with the metric's 90-day band (from computeLongHorizonContext)
  // so the prompt can soften framing for changes still inside the normal band.
  // -------------------------------------------------------------------------
  function computeSignificantChanges(thisRows, prevRows, bands) {
    bands = bands || {};
    if (thisRows.length === 0 || prevRows.length === 0) {
      return { summary: "No comparison available (missing data for one of the weeks).", count: 0 };
    }

    const avgCol = (rows, idx) => {
      const valid = rows.map(r => parseFloat(r[idx]) || 0).filter(v => v > 0);
      if (valid.length === 0) return null;
      return valid.reduce((a, b) => a + b, 0) / valid.length;
    };

    // Per-metric significance thresholds. Use absThreshold for metrics where
    // % change is misleading (e.g., 2 lbs IS big even though it's only ~1%).
    const metrics = [
      { label: 'Weight',      col: 1,  higherBetter: false, absThreshold: 2,  unit: 'lbs',  fmt: v => v.toFixed(1) },
      { label: 'Resting HR',  col: 2,  higherBetter: false, pctThreshold: 5,  unit: 'bpm',  fmt: v => v.toFixed(0) },
      { label: 'HRV',         col: 3,  higherBetter: true,  pctThreshold: 8,  unit: 'ms',   fmt: v => v.toFixed(1) },
      { label: 'Total Sleep', col: 7,  higherBetter: true,  pctThreshold: 8,  unit: 'hrs',  fmt: v => (v / 60).toFixed(1) },
      { label: 'Deep Sleep',  col: 9,  higherBetter: true,  pctThreshold: 15, unit: 'hrs',  fmt: v => (v / 60).toFixed(1) },
      { label: 'REM Sleep',   col: 10, higherBetter: true,  pctThreshold: 15, unit: 'hrs',  fmt: v => (v / 60).toFixed(1) },
      { label: 'Steps',       col: 12, higherBetter: true,  pctThreshold: 12, unit: '',     fmt: v => Math.round(v).toLocaleString() },
      { label: 'Readiness',   col: 16, higherBetter: true,  absThreshold: 8,  unit: '/100', fmt: v => v.toFixed(0) },
    ];

    const changes = metrics.map(m => {
      const thisVal = avgCol(thisRows, m.col);
      const prevVal = avgCol(prevRows, m.col);
      if (thisVal === null || prevVal === null) return null;
      const absDelta = thisVal - prevVal;
      const pctDelta = prevVal !== 0 ? (absDelta / prevVal) * 100 : 0;
      let significant = false;
      if (m.absThreshold !== undefined && Math.abs(absDelta) >= m.absThreshold) significant = true;
      if (m.pctThreshold !== undefined && Math.abs(pctDelta) >= m.pctThreshold) significant = true;
      return { ...m, thisVal, prevVal, absDelta, pctDelta, significant };
    }).filter(Boolean);

    const sig = changes.filter(c => c.significant);
    if (sig.length === 0) {
      return { summary: "No metric moved enough this week to flag as significant (thresholds: ~8% for most metrics; 2 lbs for weight; 8 pts for readiness; 5% for RHR).", count: 0 };
    }

    // Biggest movers first (by absolute % change)
    sig.sort((a, b) => Math.abs(b.pctDelta) - Math.abs(a.pctDelta));

    const lines = sig.map(c => {
      const arrow = c.absDelta > 0 ? '⬆️' : '⬇️';
      const isGood = (c.higherBetter === (c.absDelta > 0));
      const tag = isGood ? 'improvement' : 'regression';
      const pctStr = `${c.pctDelta > 0 ? '+' : ''}${c.pctDelta.toFixed(1)}%`;
      const bandTag = bands[c.label] ? ` [90d band: ${bands[c.label]}]` : '';
      return `  ${arrow} ${c.label}: ${c.fmt(c.prevVal)} → ${c.fmt(c.thisVal)} ${c.unit} (${pctStr}, ${tag})${bandTag}`;
    });

    return { summary: lines.join('\n'), count: sig.length };
  }

  const horizonContext = computeLongHorizonContext(thisWeekHealth, ninetyDayHealth);
  const sigChanges = computeSignificantChanges(thisWeekHealth, prevWeekHealth, horizonContext.bands);

  const thisWeekSummary   = summarizeHealth(thisWeekHealth, 'THIS WEEK');
  const prevWeekSummary   = summarizeHealth(prevWeekHealth, 'PREVIOUS WEEK');
  const thirtyDaySummary  = summarizeHealth(thirtyDayHealth, 'LAST 30 DAYS (baseline)');
  const ninetyDaySummary  = summarizeHealth(ninetyDayHealth, 'LAST 90 DAYS (baseline)');
  const oneYearSummary    = summarizeHealth(oneYearHealth, 'SAME WEEK 1 YEAR AGO');
  const twoYearSummary    = summarizeHealth(twoYearHealth, 'SAME WEEK 2 YEARS AGO');

  const bestsStr = `ALL-TIME CONTEXT (${personalBests.totalDays} days of data):
  All-time best HRV: ${personalBests.bestHRV} ms
  All-time best RHR: ${personalBests.bestRHR} bpm
  All-time best sleep: ${(personalBests.bestSleep / 60).toFixed(1)} hrs
  All-time best readiness: ${personalBests.bestReadiness}/100
  Lifetime avg HRV: ${personalBests.avgHRV?.toFixed(1)} ms
  Lifetime avg RHR: ${personalBests.avgRHR?.toFixed(0)} bpm`;

  // -------------------------------------------------------------------------
  // 5. Build the prompt and call Gemini
  // -------------------------------------------------------------------------
  const weekLabel = `${Utilities.formatDate(thisWeekStart, tz, 'MMM dd')} – ${Utilities.formatDate(thisWeekEnd, tz, 'MMM dd, yyyy')}`;
  const athleteContext = buildAthleteContext();
  const normativeContext = buildNormativeContext();
  const athleteBaseline = buildAthleteBaseline(ninetyDayHealth, ninetyDayWorkouts);
  const priorReportsContext = buildPriorReportsContext();
  const lastWeekActions = _getLastWeekActionItems();
  const lensPick = pickLensForWeek({
    thisWeekHealth,
    ninetyDayHealth,
    personalBests,
    sigChanges,
    thisWeekWorkoutRows: thisWkWorkouts,
    sortedByDate,
    thisWeekStart,
    thisWeekEnd,
    lastWeekActions,
  });
  const lensForWeek = lensPick.lens;
  console.log(`Lens for this week: ${lensForWeek.name} — reason: ${lensPick.reason}`);

  // If the lens picker surfaced milestone-level bests, pipe them into the
  // prompt so the AI celebrates the specific achievement (not just the
  // generic "good week" vibe the Cheerleader voice implies).
  const celebrationBlock = (lensPick.celebrationSignals && lensPick.celebrationSignals.length)
    ? `\nCELEBRATION SIGNALS — milestone moments from this week. Lead the report with at least one of these; name the specific metric and the time span (e.g. "lowest weight in 2.3 years"). Don't bury it in paragraph three.\n  - ${lensPick.celebrationSignals.join('\n  - ')}\n`
    : '';

  // If last week's suggestions and this week's actual behavior diverged,
  // surface those as observations — never as accusations. The athlete has
  // full autonomy, and a prior suggestion may itself have been miscalibrated
  // to their baseline. "Pattern notes" framing lets any lens (Straight Talk,
  // Storyteller, Physiologist) cite the divergence honestly without sliding
  // into drill-sergeant voice.
  const patternNotes = lensPick.patternNotes || lensPick.violationDetails || [];
  const continuingPattern = lensPick.continuingPattern === true;
  const patternNotesBlock = patternNotes.length
    ? `\nPATTERN NOTES — places where last week's suggestions and this week's actual behavior diverged. Notes, not accusations. Guidance:
  - The athlete has full autonomy. Frame as observation: "I suggested X; you chose Y — here's what the data shows." Not as failure-to-obey.
  - Before citing a divergence, honestly ask whether the prior suggestion was well-calibrated to baseline. If it wasn't (e.g. a 4 mi hike cap suggested to someone whose baseline is 8 mi), say so — that's coach accountability, not athlete fault.
  - Cite at most one or two divergences in the report. Don't enumerate every item; pick what's most meaningful.${continuingPattern ? '\n  - This is a CONTINUING pattern across multiple weeks. Name that plainly — but in the current lens voice, not Straight Talk\'s voice (that one ran recently). Avoid the word "again"; prefer specific data about the pattern\'s physiological cost if any.' : ''}
  Pattern notes detected:
  - ${patternNotes.join('\n  - ')}\n`
    : '';

  const interestingAngle = detectInterestingAngle({
    thisWeekHealth,
    thirtyDayHealth,
    ninetyDayHealth,
    oneYearHealth,
    twoYearHealth,
    firstPeriodHealth,
    personalBests,
    thisWeekWorkoutRows: thisWkWorkouts,
  }, lensForWeek.name);
  console.log(`Interesting Stuff angle: ${interestingAngle.angle} — ${interestingAngle.leadData}`);

const prompt = `You are a personal health and fitness coach writing a CONCISE weekly report. Be specific with numbers but keep it tight — no fluff, no filler. The whole report should be under 700 words.

COACH IDENTITY — read this first and let it shape every sentence you write:
You are an AI coach offering data-informed suggestions to a smart, experienced athlete with full autonomy. They will weigh your input against their own judgment, life schedule, body sense, and history. You are NOT their boss and you do NOT issue directives. When they do something different from what you suggested, the first question is whether your suggestion was well-calibrated to them in the first place — not whether they "failed to comply."

BANNED LANGUAGE (never use these in any section, regardless of lens): "directive", "obey", "blatantly", "ignored", "refused", "failed to execute", "no excuses", "no exceptions", "strictly capped", "strict cap", "master class in self-sabotage", "you have not earned", "unforced errors", "ego-driven", "ego overrides", "self-sabotage". These frame the athlete as a subordinate. If you catch yourself reaching for this tone, rewrite.

PREFERRED TONE: Observational and collaborative. "The data shows..." / "One thing worth noticing..." / "Given that your typical hike is 8 miles, a lighter week might look like..." / "I suggested X last week; you went with Y — here's what the metrics say about that."

LENS FOR THIS WEEK: ${lensForWeek.name}
${lensForWeek.direction}
Do not mention the lens name in the report. Just write in that voice. The lens rotates weekly — the athlete sees the variety; they shouldn't see the label.
${celebrationBlock}${patternNotesBlock}

ATHLETE CONTEXT — use this to tailor every section. Reference age/sex-appropriate norms when relevant, respect injury history (never recommend loading patterns that aggravate known issues), and tie training and recovery back to the athlete's stated goals when it adds something. Don't restate the profile back to them.
${athleteContext}

ATHLETE BASELINE — what a typical week actually looks like for this athlete. Every Next Week suggestion must be anchored here. If you're tempted to suggest an activity cap or change that's a >25% cut from baseline (e.g. "4 mi hike" when typical is 8 mi), either (a) don't suggest it unless there's a specific medical reason stated by the athlete, or (b) frame it honestly as a temporary reduction with a clear rationale and a return path. Prefer small deltas ("try one shorter hike this week — maybe 5-6 mi instead of your usual 8") over arbitrary absolute caps. Never suggest something you can't justify against this baseline.
${athleteBaseline}

NORMATIVE RANGES — use these for population context ("your HRV sits in the above-average band for your age"), but the athlete's OWN 90-day baseline is a truer personal reference. Use norms sparingly and only when they add a real insight; don't force a population comparison into every bullet.
${normativeContext}

PRIOR REPORTS — what you wrote in recent weeks. Three rules: (1) do NOT repeat the same framing, metaphors, or opening hooks you used recently — the athlete reads these consecutively and stale phrasing is the #1 way this report feels templated. (2) The "Interesting Stuff framing used" lines below are off-limits as themes — if you already hooked on glute strain, PT visits, sleep protocols, or any recurring subject in prior weeks, do NOT lead the Interesting Stuff section with that same subject again. Find a different angle from this week's data. (3) Grade follow-through: if last week's action items show up in this week's data (or conspicuously don't), call it out. E.g. "you said you'd nap twice — only managed once, and it shows in Wed's readiness." Reference prior weeks naturally, not mechanically.
${priorReportsContext}

CRITICAL CONTENT: Talk like a coach, not a spreadsheet. Keep the raw numbers for specificity, but use these Strava-style labels for the training-load metrics — never use the acronyms TSS, CTL, ATL, or TSB:
  - CTL → "Fitness score" (the rolling ~6-week load — your baseline capacity)
  - ATL → "Fatigue" (how tired you are from recent workouts)
  - TSB → "Freshness" (positive = well-rested and primed; negative = carrying fatigue)
  - TSS → describe conversationally as the week's overall training load/effort, or skip the raw number if it doesn't add anything. Do NOT call it "training stress points" or "TSS".
For example: "Your Fitness score is holding steady at 12.9 and Fatigue dropped to 7, leaving you with +5.9 Freshness — well-rested and primed for work." NOT "CTL is 12.9, ATL is 7, TSB is +5.9."
HRV, RHR, and readiness are fine to use directly with a brief parenthetical on first mention (e.g. "Heart Rate Variability (HRV)"). Keep sentences punchy and conversational. Do NOT include a glossary or a definitions section.

HORIZON-AWARE FRAMING — read the LONG-HORIZON CONTEXT block (in the data section below) before writing about any metric. Week-over-week deltas over-dramatize normal noise; the 90-day band and rolling trend are the honest reference.
  - Lead each metric mention with where it sits in the 90-day band and the 90d trend, not the WoW delta. WoW alone does NOT justify drama.
  - "normal" band → demote to a brief mention unless one of these signal categories applies, in which case lead with the signal: cross-metric divergence (e.g. sleep was high but readiness was low), intra-week pattern (day-of-week swing), variance change (stable vs choppy), conditional performance (e.g. Friday hikes after Thursday strength), goal trajectory, successful stability under a new stimulus, same-condition historical comp.
  - "low" or "high" band → full treatment with WoW + trend context.
  - "extreme_low" or "extreme_high" → headline-worthy.
  - Same rule applies to vs Last Week: each flagged change carries a [90d band: X] tag. Bullets tagged "normal" should be framed as a normal-range fluctuation rather than dramatized; bullets tagged low/high/extreme get full framing. Lead with the biggest mover that's also out of the normal band.
  - It's fine — and often correct — for a week to have nothing dramatic to say. Don't manufacture drama; surface a pattern instead.

CRITICAL OUTPUT RULES: NEVER include meta-commentary, self-corrections, notes-to-self, placeholders, or asterisk asides like "*omitting*", "*TBD*", or "*Wait…*". Decide silently and ship clean output only — the user sees exactly what you write. In the "vs Last Week" section, only include metrics that have real data for BOTH weeks; if either side is missing the metric, skip that bullet entirely without acknowledging the omission.

CRITICAL FORMATTING: You MUST use ### for all section headers (e.g. ### ${lensForWeek.emoji} Weekly Headline). Do not just bold the text. The Weekly Headline header MUST use the exact emoji "${lensForWeek.emoji}" — this emoji is the lens fingerprint and is non-negotiable. Do not substitute a different one.

Use emojis for section headers and key callouts to make the report visually engaging and easy to scan. Use them naturally — e.g. 🏋️ for training, 😴 for sleep, 💚 for good trends, ⚠️ for concerns, 🔥 for achievements, 📊 for data insights.

SECTIONS:

1. **${lensForWeek.emoji} Weekly Headline** — 2 sentences. What defined this week? Open the header with exactly "### ${lensForWeek.emoji} Weekly Headline".

2. **🏋️ Training** — FIRST, provide a concise bulleted list of the week's workouts (e.g. "* **Tue:** 🏋️ Strength (65m) - *Felt strong*"). Use a bold day, an emoji, the activity, and a brief 1-line summary/note for each. SECOND, write a short narrative (2-3 sentences max) analyzing the week's Fitness/Fatigue/Freshness. Use the "Description" and "Private Note" fields as context to make the analysis land — reference them naturally when it actually adds something (e.g. "the heavy-legs feel on Thursday tracks with your elevated HR"). Never quote them verbatim and never just restate them — the user wrote those words and doesn't need them read back. If a note doesn't change your take, skip it.

3. **😴 Recovery** — Frame each metric using its LONG-HORIZON CONTEXT band and 90d trend, not the WoW change. Sleep vs 7.5hr target, HRV/RHR position in the 90-day distribution, readiness highlights. If everything is in the "normal" band, skip the metric-by-metric tour and lead with the most meaningful pattern instead (intra-week swing, cross-metric divergence, variance shift, conditional performance) — that's where the real signal lives in a steady week. 3-4 sentences.

4. **📊 vs Last Week** — Render ONLY the items in the SIGNIFICANT CHANGES block below. Do not invent additional comparisons, do not include metrics that weren't flagged. Each line carries a [90d band: X] tag — let the band, not the WoW %, set the tone. Bullets tagged "normal" get a soft fragment ("a normal-range fluctuation", "still inside your typical 90-day range"); bullets tagged low/high/extreme get full "why it matters" framing. Lead with the biggest mover whose band is out of normal; if all flagged movers are in the normal band, render them with soft framing only. Keep each bullet to one short line (metric + direction + numbers + 3-6 word fragment). Max 4 bullets. Strip the [90d band: X] tag from your output — it's a directive for you, not text for the report. If the SIGNIFICANT CHANGES block says nothing moved enough, replace this whole section with a single line: "Nothing moved enough to call out — a steady week." and move on. No ➡️ arrows — only ⬆️ and ⬇️.

5. **🔍 The Interesting Stuff** — Lead with the pre-computed INTERESTING STUFF ANGLE below. Open with the "Lead data" in your own words, then use the "Framing" as guidance for how to develop the insight. 2-3 sentences total. You may add ONE additional observation if you genuinely notice something else surprising in the data. Year-over-year comparisons are welcome when the angle is LONGTERM_PROGRESS or TRAJECTORY_SHIFT — use them fully there. For other angles, prefer a fresher hook, but a specific YoY callout is fine if it genuinely adds something (just don't use it as filler). Keep the whole section under 100 words.

6. **🎯 Next Week** — 3 bullet suggestions (not commands). Short and specific. Anchor each to the ATHLETE BASELINE — prefer small deltas ("one shorter hike, maybe 5-6 mi instead of your usual 8") over absolute caps ("cap at 4 mi"). Phrase as offers the athlete can weigh: "consider", "worth trying", "one option is", "you might experiment with", "given how the glute felt, maybe...". Never use "cap", "strictly", "must", "obey", "no exceptions", "no excuses", "you need to", "you have to". If you're suggesting a reduction from baseline, give a brief reason and a timeframe (this week only, until the glute quiets down, etc.).

Use specific numbers (e.g. "HRV rose from 45.2 to 46.9 ms"). Be direct but conversational.

Data for ${weekLabel}:

=== LONG-HORIZON CONTEXT (pre-computed — frame metrics against these bands and trends, not just WoW deltas) ===
${horizonContext.summary}

=== SIGNIFICANT CHANGES vs LAST WEEK (pre-computed — use these verbatim for the vs-Last-Week section) ===
${sigChanges.summary}

=== INTERESTING STUFF ANGLE (pre-computed — lead the Interesting Stuff section with this) ===
Angle: ${interestingAngle.angle}
Lead data: ${interestingAngle.leadData}
Framing: ${interestingAngle.framing}

=== THIS WEEK ===
${thisWeekSummary}

=== PREVIOUS WEEK ===
${prevWeekSummary}

=== 30-DAY BASELINE ===
${thirtyDaySummary}

=== 90-DAY BASELINE ===
${ninetyDaySummary}

=== SAME WEEK 1 YEAR AGO ===
${oneYearSummary}

=== SAME WEEK 2 YEARS AGO ===
${twoYearSummary}

=== ${bestsStr} ===

=== TRAINING LOAD ===
${trainingLoadSummary}

=== THIS WEEK'S WORKOUTS ===
${thisWeekWorkouts}

=== PREVIOUS WEEK'S WORKOUTS ===
${prevWeekWorkouts}`;

  console.log("Calling Gemini API...");

    // Uses "gemini-pro-latest" so it auto-upgrades to the newest Pro model.
  // If this ever breaks, check available models at:
  // https://generativelanguage.googleapis.com/v1beta/models?key=YOUR_KEY
  // and update the model name below.
  const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro-latest:generateContent?key=${GEMINI_API_KEY}`;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 16384
    }
  };

  const response = UrlFetchApp.fetch(geminiUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    console.error("Gemini API error: " + response.getContentText());
    return;
  }

  const result = JSON.parse(response.getContentText("UTF-8"));
  const reportText = result.candidates?.[0]?.content?.parts?.[0]?.text;
  const finishReason = result.candidates?.[0]?.finishReason || 'UNKNOWN';

  console.log(`Gemini response — Finish reason: ${finishReason}, Length: ${reportText ? reportText.length : 0} chars`);

  if (finishReason === 'MAX_TOKENS') {
    console.warn("WARNING: Response was truncated! Increase maxOutputTokens.");
  }

  if (!reportText) {
    console.error("No report generated. Response: " + response.getContentText());
    return;
  }

  // Log first and last 200 chars to verify completeness
  console.log("Report start: " + reportText.substring(0, 200));
  console.log("Report end: " + reportText.substring(reportText.length - 200));

  // -------------------------------------------------------------------------
  // 6. Generate inline charts
  // -------------------------------------------------------------------------
  const inlineImages = {};

  // Readiness chart (7-day bar chart)
  const readinessData = thisWeekHealth.map(r => {
    const day = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'EEE')
      : String(r[0]).substring(5, 10);
    return [day, parseFloat(r[16]) || 0];
  });
  if (readinessData.length > 0) {
    inlineImages.readinessChart = buildBarChart(
      'Readiness Score', readinessData, '#1a73e8', 70
    );
  }

  // Sleep chart (7-day stacked bar)
  const sleepData = thisWeekHealth.map(r => {
    const day = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'EEE')
      : String(r[0]).substring(5, 10);
    return [day,
      (parseFloat(r[9]) || 0) / 60,   // Deep
      (parseFloat(r[8]) || 0) / 60,   // Light
      (parseFloat(r[10]) || 0) / 60   // REM
    ];
  });
  if (sleepData.length > 0) {
    inlineImages.sleepChart = buildStackedChart(
      'Sleep Breakdown (hrs)', sleepData,
      ['Deep', 'Light', 'REM'], ['#1a73e8', '#e69138', '#9b59b6'], 7.5
    );
  }

  // HRV & RHR line chart — use null for missing/zero so line doesn't crash to 0.
  // RHR comes first so it lands on the left axis (bpm); HRV on the right (ms).
  const vitalsData = thisWeekHealth.map(r => {
    const day = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'EEE')
      : String(r[0]).substring(5, 10);
    const hrv = parseFloat(r[3]);
    const rhr = parseFloat(r[2]);
    return [day, rhr > 0 ? rhr : null, hrv > 0 ? hrv : null];
  });
  if (vitalsData.length > 0) {
    inlineImages.vitalsChart = buildLineChart(
      'HRV & RHR', vitalsData,
      ['RHR (bpm)', 'HRV (ms)'], ['#e69138', '#3d85c6'],
      { dualAxis: true }
    );
  }

  // All-time charts (weekly-binned — noisy enough to show real variation,
  // sparse enough to render across multi-year history).
  const weeklyBins = (() => {
    const bins = new Map();   // 'YYYY-MM-DD' (Sun anchor) → { date, weight, sleep, hrv, rhr }
    allRows.forEach(r => {
      const d = r[0] instanceof Date
        ? r[0]
        : new Date(String(r[0]).substring(0, 10) + 'T12:00:00');
      if (isNaN(d)) return;
      // Anchor each bin to the Sunday that starts its week.
      const anchor = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      anchor.setDate(anchor.getDate() - anchor.getDay());
      const key = Utilities.formatDate(anchor, tz, 'yyyy-MM-dd');
      let bin = bins.get(key);
      if (!bin) {
        bin = { date: anchor, weight: [], sleep: [], hrv: [], rhr: [] };
        bins.set(key, bin);
      }
      const w = parseFloat(r[1]); if (w > 0) bin.weight.push(w);
      const s = parseFloat(r[7]); if (s > 0) bin.sleep.push(s);   // minutes
      const h = parseFloat(r[3]); if (h > 0) bin.hrv.push(h);
      const rh = parseFloat(r[2]); if (rh > 0) bin.rhr.push(rh);
    });
    return Array.from(bins.values()).sort((a, b) => a.date - b.date);
  })();

  const avgOrNull = arr => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;

  if (weeklyBins.length > 0) {
    // Only label the bin closest to Jan 1 and Jul 1 of each year covered.
    // Every other bin gets an empty label so the x-axis stays clean.
    const anchorLabels = (() => {
      const out = new Array(weeklyBins.length).fill('');
      const minDate = weeklyBins[0].date;
      const maxDate = weeklyBins[weeklyBins.length - 1].date;
      for (let y = minDate.getFullYear(); y <= maxDate.getFullYear(); y++) {
        [0, 6].forEach(monthIdx => {
          const target = new Date(y, monthIdx, 1);
          if (target < minDate || target > maxDate) return;
          // First bin whose Sunday-anchor is on or after the target date.
          const idx = weeklyBins.findIndex(b => b.date >= target);
          if (idx >= 0 && !out[idx]) {
            out[idx] = `${monthIdx === 0 ? 'Jan' : 'Jul'} ${String(y).slice(-2)}`;
          }
        });
      }
      return out;
    })();

    const weightData = weeklyBins.map((b, i) => [anchorLabels[i], avgOrNull(b.weight)]);
    if (weightData.some(d => d[1] !== null)) {
      inlineImages.allTimeWeightChart = buildLineChart(
        'Weight (all time)', weightData, ['Weight (lbs)'], ['#34a853'],
        { interpolate: true, sparseXLabels: true }
      );
    }

    const sleepTotalData = weeklyBins.map((b, i) => {
      const avgMin = avgOrNull(b.sleep);
      return [anchorLabels[i], avgMin === null ? null : avgMin / 60];
    });
    if (sleepTotalData.some(d => d[1] !== null)) {
      inlineImages.allTimeSleepChart = buildLineChart(
        'Total Sleep (all time)', sleepTotalData,
        ['Sleep (hrs)'], ['#1a73e8'],
        { refLine: 7.5, interpolate: true, sparseXLabels: true }
      );
    }

    // RHR first → left axis (bpm); HRV second → right axis (ms).
    const allTimeVitalsData = weeklyBins.map((b, i) => [
      anchorLabels[i], avgOrNull(b.rhr), avgOrNull(b.hrv),
    ]);
    if (allTimeVitalsData.some(d => d[1] !== null || d[2] !== null)) {
      inlineImages.allTimeVitalsChart = buildLineChart(
        'HRV & RHR (all time)', allTimeVitalsData,
        ['RHR (bpm)', 'HRV (ms)'], ['#e69138', '#3d85c6'],
        { dualAxis: true, interpolate: true, sparseXLabels: true }
      );
    }
  }

  // Fitness vs Fatigue chart — 60-day trend of CTL and ATL from Training Load sheet
  if (tlData) {
    const fitnessWindowEnd = thisWeekEnd;
    const fitnessWindowStart = new Date(fitnessWindowEnd);
    fitnessWindowStart.setDate(fitnessWindowStart.getDate() - 59);  // 60 days

    const fitnessRows = tlData.slice(1)
      .filter(row => {
        if (!row[0]) return false;
        const d = row[0] instanceof Date
          ? row[0]
          : new Date(String(row[0]).substring(0, 10) + 'T12:00:00');
        return d >= fitnessWindowStart && d <= fitnessWindowEnd;
      })
      .sort((a, b) => {
        const da = a[0] instanceof Date ? a[0] : new Date(String(a[0]).substring(0, 10));
        const db = b[0] instanceof Date ? b[0] : new Date(String(b[0]).substring(0, 10));
        return da - db;
      });

    if (fitnessRows.length > 0) {
      const fitnessData = fitnessRows.map(r => {
        const date = r[0] instanceof Date
          ? Utilities.formatDate(r[0], tz, 'MM/dd')
          : String(r[0]).substring(5, 10);
        const ctl = parseFloat(r[2]);
        const atl = parseFloat(r[3]);
        return [date, isFinite(ctl) ? ctl : null, isFinite(atl) ? atl : null];
      });
      inlineImages.fitnessChart = buildLineChart(
        'Fitness vs Fatigue (60 days)', fitnessData,
        ['Fitness', 'Fatigue'], ['#1a73e8', '#e69138']
      );
    }
  }

// -------------------------------------------------------------------------
  // 7. Convert markdown to HTML and send via Gmail
  // -------------------------------------------------------------------------
  const htmlBody = markdownToHtml(reportText, weekLabel, inlineImages);

  // Use MailApp instead of GmailApp to preserve Unicode/Emojis
  if (!skipEmail) {
    MailApp.sendEmail(
      REPORT_EMAIL,
      `Weekly Health Report — ${weekLabel}`,
      reportText,  // plain text fallback
      { htmlBody: htmlBody, inlineImages: inlineImages }
    );
    console.log("Weekly Coach Report sent!");
  } else {
    console.log("skipEmail=true — report generated but NOT emailed.");
  }

  // Persist to Report History so next week's prompt can reference what we said.
  try {
    saveReportToHistory(
      weekLabel,
      lensForWeek ? lensForWeek.name : '',
      interestingAngle ? interestingAngle.angle : '',
      reportText
    );
    console.log(`Report saved to history (lens: ${lensForWeek ? lensForWeek.name : 'none'}, angle: ${interestingAngle ? interestingAngle.angle : 'none'}).`);
  } catch (e) {
    console.error(`Failed to save report to history: ${e.message}`);
  }
}
