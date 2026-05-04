// =============================================================================
// LENS OF THE WEEK — the coach's voice shifts each report so consecutive
// weeks don't read the same. Selection is data-driven: signals in the week's
// data trigger matching lenses (e.g. all-time best → Cheerleader, major
// regression → Straight Talk). When no signal is strong enough, we fall back to
// a rotation through the "neutral" lenses, avoiding whichever one ran most
// recently.
// =============================================================================
// Each lens has an emoji used in the Weekly Headline (replacing the generic
// 📋) so the athlete can visually identify the lens at a glance without us
// labeling it by name. Emojis chosen to not collide with section headers
// already in use (🏋️ Training, 😴 Recovery, 📊 vs Last Week, 🔍 The
// Interesting Stuff, 🎯 Next Week).
const COACH_LENSES = [
  {
    name: 'Storyteller',
    emoji: '📖',
    direction: 'Frame this week as a single narrative thread. Pick one dominant arc (a comeback, a grind, a breakthrough, a reset) and tell the week through that lens. Use vivid, concrete language. Section headers and structure stay the same — only the voice shifts.',
  },
  {
    name: 'Detective',
    emoji: '🕵️',
    direction: 'Write like an analyst investigating anomalies. For every notable metric, ask "what caused this?" and point at the most likely driver from the data (workouts, sleep, prior-week fatigue, life context). Be concrete about cause-and-effect. Okay to speculate if you label it as such.',
  },
  {
    name: 'Strategist',
    emoji: '🧭',
    direction: 'Zoom out. Frame this week in the context of the athlete\'s goals ({{GOALS}}). How did this week move them closer or further? What does the next 2-4 weeks need to look like? Next Week bullets should feel like moves in a larger plan, not isolated tasks.',
  },
  {
    name: 'Physiologist',
    emoji: '🫀',
    direction: 'Explain what\'s actually happening in the athlete\'s body — why HRV responds the way it does, what dropping RHR signals about adaptation, why fatigue accumulates. Teach, don\'t just describe. Keep it accessible; no jargon walls.',
  },
  {
    name: 'Straight Talk',
    emoji: '💬',
    direction: 'Direct, honest, and respectful. Name the patterns the data shows without moralizing. When last week\'s suggestion and this week\'s behavior diverged, explore *why* (was the suggestion misaligned with the athlete\'s baseline? did life get in the way? did the suggestion itself miss context?) rather than treating it as disobedience. The athlete is a partner with full autonomy, not a subordinate. Firm about real safety issues, curious and grounded about everything else. Never use "directive," "obey," "strict cap," "no exceptions," "blatantly," "ignored," "refused," "failed to execute," "you have not earned," "self-sabotage." Next Week bullets are firm-but-reasoned suggestions anchored to baseline, not commands.',
  },
  {
    name: 'Cheerleader',
    emoji: '🎉',
    direction: 'Lean into what went well. Find the wins, however small, and celebrate them with specifics. Reframe setbacks as setups. Still honest about what needs work, but the energy is "you\'re doing the thing, keep going." No toxic positivity — back every compliment with data.',
  },
];

function _getLens(name) {
  const lens = COACH_LENSES.find(l => l.name === name);
  if (!lens) return null;
  if (lens.direction.indexOf('{{GOALS}}') === -1) return lens;
  const goals = (ATHLETE_PROFILE && ATHLETE_PROFILE.goals && ATHLETE_PROFILE.goals.length)
    ? ATHLETE_PROFILE.goals.join(', ')
    : 'their stated goals';
  return Object.assign({}, lens, { direction: lens.direction.replace('{{GOALS}}', goals) });
}

// Pull a target weight (lbs) out of ATHLETE_PROFILE.goals, e.g. a goal like
// "Reach NNN lbs bodyweight" yields NNN. Returns null if no matching goal.
function _targetWeightLbs() {
  if (!ATHLETE_PROFILE || !ATHLETE_PROFILE.goals) return null;
  for (const g of ATHLETE_PROFILE.goals) {
    const m = String(g).match(/(\d+(?:\.\d+)?)\s*lbs?\b/i);
    if (m) return parseFloat(m[1]);
  }
  return null;
}

// True if the athlete has a half-marathon-related goal.
function _hasHalfMarathonGoal() {
  if (!ATHLETE_PROFILE || !ATHLETE_PROFILE.goals) return false;
  return ATHLETE_PROFILE.goals.some(g => /half[\s-]?marathon/i.test(String(g)));
}

// Average a column across rows, filtering out missing/zero values.
function _avgCol(rows, idx) {
  const vals = rows.map(r => parseFloat(r[idx]) || 0).filter(v => v > 0);
  if (vals.length === 0) return null;
  return vals.reduce((a, b) => a + b, 0) / vals.length;
}

// Sample standard deviation.
function _stdevCol(rows, idx) {
  const vals = rows.map(r => parseFloat(r[idx]) || 0).filter(v => v > 0);
  if (vals.length < 2) return null;
  const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
  const variance = vals.reduce((s, v) => s + (v - mean) ** 2, 0) / (vals.length - 1);
  return Math.sqrt(variance);
}

// For a given metric column, find how many days since the last time the
// athlete was at least as good as `thisWeekExtreme`. Scans `sortedByDate`
// (oldest → newest) but ignores any row inside the current week window so
// a self-compare doesn't return 0. Returns:
//   • Infinity  — never matched before (true all-time best)
//   • Number    — days since the most recent prior match
//   • null      — not enough data / bad inputs
function _daysSinceLastAtLeastAsGood(sortedByDate, col, higherBetter, thisWeekExtreme, thisWeekStart, thisWeekEnd) {
  if (thisWeekExtreme == null || !sortedByDate || !sortedByDate.length) return null;
  const startMs = thisWeekStart.getTime();
  const endMs   = thisWeekEnd.getTime();
  for (let i = sortedByDate.length - 1; i >= 0; i--) {
    const row = sortedByDate[i];
    const d = row[0] instanceof Date ? row[0] : new Date(row[0]);
    const ms = d.getTime();
    if (ms >= startMs && ms <= endMs) continue;  // skip this week itself
    const v = parseFloat(row[col]);
    if (!v || v <= 0) continue;
    const asGood = higherBetter ? v >= thisWeekExtreme : v <= thisWeekExtreme;
    if (asGood) {
      return Math.floor((endMs - ms) / 86400000);
    }
  }
  return Infinity;  // never matched before → all-time best
}

// Format a days-since value into a human label. 90+ days graduates to years.
function _formatDaysSince(days) {
  if (days === Infinity) return 'all-time best';
  if (days >= 365) {
    const yrs = (days / 365).toFixed(1);
    return `best in ${days} days (${yrs} years)`;
  }
  if (days >= 90) return `best in ${days} days (~${Math.round(days / 30)} months)`;
  return `best in ${days} days`;
}

// Read the most recent N lens names from Report History (newest first).
function _getRecentLensNames(n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPORT_HISTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  return data.slice(1).slice(-n).reverse().map(r => r[2]).filter(Boolean);
}

// Read the most recent report's "Next Week" section — the action items we
// gave the athlete last week. Used by the lens picker to detect directive
// pattern divergences (Straight Talk) and by the prompt for follow-through grading.
// Returns '' if there's no prior report on file.
function _getLastWeekActionItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPORT_HISTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return '';
  const data = sheet.getDataRange().getValues();
  // Sheet layout: [Generated At, Week Label, Lens, Angle, Full Report]
  const lastRow = data[data.length - 1];
  const fullReport = lastRow[4] || '';
  return extractReportSection(fullReport, 'next week');
}

// Read the most recent N angle names from Report History (newest first).
// Used to throttle angles that shouldn't fire every week.
function _getRecentAngles(n) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPORT_HISTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  // Column layout: [Generated At, Week Label, Lens, Angle, Full Report]
  return data.slice(1).slice(-n).reverse().map(r => r[3]).filter(Boolean);
}

// Detect signals in this week's data. Returns a plain object of booleans
// plus any supporting numbers we want to log. Keep this pure/testable.
function detectLensSignals(ctx) {
  const {
    thisWeekHealth, ninetyDayHealth, personalBests, sigChanges, thisWeekWorkoutRows,
    sortedByDate, thisWeekStart, thisWeekEnd, lastWeekActions,
  } = ctx;
  const signals = [];

  // --- Cheerleader: all-time or multi-month best this week ---
  // For each core metric, find this week's best single-day value, then look
  // backwards through history to find how long it's been since we were at
  // least that good. Anything 90+ days (or all-time) is a genuine milestone
  // worth celebrating. Weight is included because it's tied to a stated goal.
  const thisHRVs       = thisWeekHealth.map(r => parseFloat(r[3])  || 0).filter(v => v > 0);
  const thisRHRs       = thisWeekHealth.map(r => parseFloat(r[2])  || 0).filter(v => v > 0);
  const thisReadinesss = thisWeekHealth.map(r => parseFloat(r[16]) || 0).filter(v => v > 0);
  const thisSleeps     = thisWeekHealth.map(r => parseFloat(r[7])  || 0).filter(v => v > 0);
  const thisWeights    = thisWeekHealth.map(r => parseFloat(r[1])  || 0).filter(v => v > 0);

  const MILESTONE_MIN_DAYS = 90;
  const metricsToScan = [
    { label: 'HRV',       col: 3,  higherBetter: true,  extreme: thisHRVs.length       ? Math.max(...thisHRVs)       : null, fmt: v => `${v.toFixed(1)} ms`   },
    { label: 'RHR',       col: 2,  higherBetter: false, extreme: thisRHRs.length       ? Math.min(...thisRHRs)       : null, fmt: v => `${v.toFixed(0)} bpm`  },
    { label: 'Readiness', col: 16, higherBetter: true,  extreme: thisReadinesss.length ? Math.max(...thisReadinesss) : null, fmt: v => `${v.toFixed(0)}/100` },
    { label: 'Sleep',     col: 7,  higherBetter: true,  extreme: thisSleeps.length     ? Math.max(...thisSleeps)     : null, fmt: v => `${(v / 60).toFixed(1)}h` },
    { label: 'Weight',    col: 1,  higherBetter: false, extreme: thisWeights.length    ? Math.min(...thisWeights)    : null, fmt: v => `${v.toFixed(1)} lbs` },
  ];

  const newBests = [];           // milestone-level (90d+) — triggers Cheerleader
  const allTimeBests = [];       // strict subset — highest priority
  if (sortedByDate && thisWeekStart && thisWeekEnd) {
    metricsToScan.forEach(m => {
      if (m.extreme == null) return;
      const days = _daysSinceLastAtLeastAsGood(sortedByDate, m.col, m.higherBetter, m.extreme, thisWeekStart, thisWeekEnd);
      if (days == null) return;
      if (days === Infinity) {
        allTimeBests.push(`${m.label} ${m.fmt(m.extreme)} — all-time best`);
        newBests.push(`${m.label} ${m.fmt(m.extreme)} — all-time best`);
      } else if (days >= MILESTONE_MIN_DAYS) {
        newBests.push(`${m.label} ${m.fmt(m.extreme)} — ${_formatDaysSince(days)}`);
      }
    });
  }

  // --- Straight Talk: pattern divergence WITH physiological cost ---
  // We notice when last week's suggestions and this week's behavior diverged,
  // but that alone isn't enough — the athlete has autonomy, and a prior
  // suggestion may itself have been miscalibrated to their baseline. Straight
  // Talk fires only when the divergence AND the data show a real cost this
  // week (injury-adjacent readiness collapse, multiple regressions, or a
  // worsening injury signal). Otherwise, pattern notes get handled by a
  // gentler lens (Storyteller, Physiologist) or fall through to rotation.
  const violationDetails = [];
  if (lastWeekActions) {
    const actions = lastWeekActions;
    const workouts = thisWeekWorkoutRows || [];

    // Constraint: cap on cardio volume / intensity. We also catch the softer
    // register the new coach voice uses ("ease into," "hold off on," "keep
    // intensity moderate," "recovery week," "before sprints come back,"
    // "swap sprints for walks," "not yet/another week") — without these, the
    // tone-softening we did in the prompt silently disables the violation
    // detector that depends on seeing directive language.
    const cardioCapStated =
      /\b(cap|limit|keep)\b[^.]*\b(cardio|run|hike|intervals?|sprints?|heart rate|hr)\b/i.test(actions) ||
      /\bno\b[^.]*\b(sprints?|intervals?|hills?|high[- ]intensity|explosive)\b/i.test(actions) ||
      /\bzone\s*2\s*(only|run)?\b/i.test(actions) ||
      /\bstrict(ly)?\b[^.]*\b(zone|hr|heart rate|under)\b/i.test(actions) ||
      /\bhr\b.{0,10}\bunder\b.{0,10}\d/i.test(actions) ||
      // Gentle-start language ("start with a flat, easy 15-minute walk",
      // "test the waters gently", "short, easy movement only")
      /\b(start|test|ease|begin|try)\b[^.]*\b(gentle|easy|flat|short|light|small)\b/i.test(actions) ||
      /\b(gentle|easy|flat|short|light|small)\b[^.]*\b(walk|cardio|run|hike|activity|movement)\b/i.test(actions) ||
      /\b(?:10|15|20|25|30)[- ]?(?:min|minute)\b[^.]*\b(walk|cardio|run|hike)\b/i.test(actions) ||
      // Soft-register: "hold off on sprints," "pause the intervals," "save
      // the hills for later," "park hill sprints for now"
      /\b(hold off|pause|park|save|postpone|shelve)\b[^.]*\b(sprints?|intervals?|hills?|cardio|high[- ]intensity|explosive|runs?)\b/i.test(actions) ||
      // "keep intensity moderate," "stay in conversational pace," "keep effort light"
      /\b(keep|stay|remain|stick to)\b[^.]*\b(intensity|pace|effort|hr|heart rate|cardio)\b[^.]*\b(moderate|low|light|easy|chill|conversational|aerobic)\b/i.test(actions) ||
      // "recovery week," "recovery days," "deload block"
      /\b(recovery|deload|easy|aerobic)\s*(week|days?|block|phase|stretch)\b/i.test(actions) ||
      // "before sprints come back," "until the hips settle, hold off on hills"
      /\bbefore\b[^.]*\b(sprints?|intervals?|hills?|high[- ]intensity|explosive)\b[^.]*\b(comes?|returns?|back|again)\b/i.test(actions) ||
      // "swap the hill sprints for a flat walk"
      /\b(swap|trade|substitute|replace)\b[^.]*\b(sprint|interval|hill|hiit|high[- ]intensity|run)\b[^.]*\bfor\b/i.test(actions) ||
      // "not yet on sprints," "another week before intervals," "few more days before hills"
      /\b(not yet|not this week|another week|few more (weeks?|days?)|give it (a|another) (week|few))\b[^.]*\b(sprint|interval|hill|high[- ]intensity|explosive|cardio)\b/i.test(actions);

    if (cardioCapStated) {
      // Pull numeric HR cap if stated ("HR under 140", "under 140 bpm")
      const hrCapMatch = actions.match(/\bunder\s*(\d{3})\s*(bpm|hr)?/i);
      const hrCap = hrCapMatch ? parseInt(hrCapMatch[1], 10) : 160;  // default "notable HR"

      // Pull duration cap if stated ("cap cardio at 30 minutes")
      const durCapMatch = actions.match(/\b(?:cap|limit|keep|under)\s*(?:cardio|runs?|hikes?)?.{0,20}(\d{2,3})\s*(?:min|minutes)\b/i);
      const durCap = durCapMatch ? parseInt(durCapMatch[1], 10) : 60;

      const offender = workouts.find(r => {
        const type = String(r[2] || '').toLowerCase();
        const name = String(r[3] || '').toLowerCase();
        const duration = parseFloat(r[4]) || 0;
        const distance = parseFloat(r[5]) || 0;
        const maxHR = parseFloat(r[11]) || 0;
        const isCardio = /run|hike|bike|cycle|ride|swim/.test(type);
        if (!isCardio) return false;
        const overDuration = duration > durCap;
        const overHR = maxHR > 0 && maxHR > hrCap;
        const longDistance = distance >= 5;
        const explosiveName = /sprint|interval|hill|hiit/.test(name);
        return overDuration || overHR || (longDistance && hrCap < 150) || explosiveName;
      });

      if (offender) {
        const name = offender[3] || offender[2] || 'activity';
        const dur = parseFloat(offender[4]) || 0;
        const dist = parseFloat(offender[5]) || 0;
        const hr = parseFloat(offender[11]) || 0;
        violationDetails.push(
          `prior directive capped cardio/HR but logged "${name}" (${dur.toFixed(0)}min, ${dist.toFixed(1)}mi${hr ? `, max HR ${hr.toFixed(0)}` : ''})`
        );
      }
    }

    // Constraint: injury protection. Direct-command forms ("avoid explosive
    // work") plus the softer register the coach now uses ("give your hips
    // another week," "hold off on plyos," "before hills come back," "swap
    // sprints for flats," "recovery week for the back," "not yet on jumps").
    const injuryGuardStated =
      /\b(avoid|don'?t|do not|skip|no)\b[^.]*\b(explosive|snap[- ]downs?|deep lunges?|heavy|loaded|plyo|jump|sprint)\b/i.test(actions) ||
      // "give your hips another week," "let the back have more recovery time,"
      // "give your knees a stretch to settle"
      /\b(give|let|allow)\b[^.]*\b(hips?|back|knees?|hamstrings?|shoulders?|ankle|joints?|body|legs?)\b[^.]*\b(more|another|extra|recovery|rest|time|stretch|week|breathe|settle)\b/i.test(actions) ||
      // "hold off on plyos," "pause the hill sprints," "park the heavy lifts"
      /\b(hold off|pause|park|postpone|save|shelve)\b[^.]*\b(explosive|snap|plyo|jump|heavy|loaded|sprint|interval|hill|hiit)\b/i.test(actions) ||
      // "recovery week," "healing phase," "deload block"
      /\b(recovery|healing|deload)\s*(week|days?|block|phase|mode|stretch)\b/i.test(actions) ||
      // "before hill sprints come back," "until the hip settles, no plyos"
      /\bbefore\b[^.]*\b(hill|sprint|explosive|plyo|jump|heavy|loaded|interval)\b[^.]*\b(comes?|returns?|back|again)\b/i.test(actions) ||
      // "swap the sprints for a walk," "trade plyos for mobility work"
      /\b(swap|trade|substitute|replace)\b[^.]*\b(sprint|interval|hill|hiit|jump|plyo|heavy|loaded)\b[^.]*\bfor\b/i.test(actions) ||
      // "not yet on sprints," "another week before jumps," "few more weeks before plyos"
      /\b(not yet|not this week|another week|few more weeks?|give it (a|another) (week|few))\b[^.]*\b(explosive|plyo|jump|sprint|heavy|loaded|hill|interval)\b/i.test(actions);

    if (injuryGuardStated && violationDetails.length === 0) {
      // If they did heavy strength or sprint-type work, that's a violation
      const offender = workouts.find(r => {
        const type = String(r[2] || '').toLowerCase();
        const name = String(r[3] || '').toLowerCase();
        const relEffort = parseFloat(r[12]) || 0;
        const explosiveSignal = /sprint|interval|hill|hiit|jump|plyo|snap/.test(name);
        const heavyStrength = type.includes('strength') && relEffort >= 50;
        return explosiveSignal || heavyStrength;
      });
      if (offender) {
        const name = offender[3] || offender[2] || 'activity';
        violationDetails.push(`prior directive called for injury protection but logged "${name}"`);
      }
    }

    // Constraint: explicit sleep target (7.5+ hours, "secure", "claim", "lock in")
    const sleepTargetStated =
      /\b(7\.?5|8)\s*(hours?|hrs?)\b/i.test(actions) &&
      /\b(sleep|bed|bedtime|claim|secure|lock\s*in|hit|chase|target|get)\b/i.test(actions);
    if (sleepTargetStated && violationDetails.length === 0) {
      const thisSleepMin = _avgCol(thisWeekHealth, 7);
      // Violation if > 30 min below 7.5 hr target (i.e. under 7.0 hrs)
      if (thisSleepMin && thisSleepMin < 420) {
        violationDetails.push(
          `explicit sleep target (7.5+ hrs) but averaged ${(thisSleepMin / 60).toFixed(1)} hrs`
        );
      }
    }
  }
  // Pattern divergence alone isn't a Straight Talk trigger — the athlete has
  // autonomy. We require a *cost signal* in this week's data:
  //   (a) readiness collapsed — either ≥7 pts below 90-day baseline, OR in
  //       the bottom 10th percentile of the 90-day distribution. (Loosened
  //       from -10 flat, which was missing weeks like readiness-47-on-
  //       baseline-56 that feel like a clear cost to anyone reading them.)
  //   (b) 3+ metric regressions this week (from sigChanges), OR
  //   (c) this week's workouts include an injury-adjacent aggravation
  //       (explosive/high-intensity work right after an injury directive was
  //       stated last week).
  const regressionCount = (sigChanges.summary.match(/regression/g) || []).length;
  const thisReadinessAvg = _avgCol(thisWeekHealth, 16);
  const baselineReadinessAvg = _avgCol(ninetyDayHealth, 16);
  let readinessCollapse = false;
  if (thisReadinessAvg && baselineReadinessAvg) {
    if (thisReadinessAvg < baselineReadinessAvg - 7) {
      readinessCollapse = true;
    } else {
      // Percentile check: grab all non-null readiness values from 90-day
      // window, sort, and see if this week's average lands in the bottom
      // 10%. Needs enough samples to be meaningful.
      const readings = ninetyDayHealth
        .map(r => parseFloat(r[16]))
        .filter(v => v > 0)
        .sort((a, b) => a - b);
      if (readings.length >= 20) {
        const p10 = readings[Math.floor(readings.length * 0.10)];
        if (thisReadinessAvg <= p10) readinessCollapse = true;
      }
    }
  }
  const multipleRegressions = regressionCount >= 3;
  // Injury-adjacent: injuryGuard directive was stated AND at least one
  // violation entry mentions explosive/injury context.
  const injuryAdjacentCost = violationDetails.some(v =>
    /injury protection|snap|explosive|sprint|hill|interval/i.test(v));

  const hasCostSignal = readinessCollapse || multipleRegressions || injuryAdjacentCost;
  const straightTalkTriggered = violationDetails.length > 0 && hasCostSignal;

  // Pattern divergence WITHOUT a cost signal is still worth surfacing, just
  // in a gentler lens — we'll tag the ctx so pickLensForWeek can route it to
  // Storyteller/Physiologist with pattern-notes prompt framing.
  const softPatternDivergence = violationDetails.length > 0 && !hasCostSignal;

  // For biasing rotation away from Cheerleader on non-violation rough weeks.
  const readinessBelowBaseline =
    thisReadinessAvg && baselineReadinessAvg && (thisReadinessAvg < baselineReadinessAvg - 3);
  const roughWeekNoViolation = !straightTalkTriggered && violationDetails.length === 0 &&
    (regressionCount >= 2 || readinessBelowBaseline);

  // --- Detective: statistical outlier vs 90-day baseline ---
  // Any of HRV / RHR / readiness >2 SD from 90-day mean (positive or negative).
  const anomalies = [];
  [
    { name: 'HRV',       col: 3 },
    { name: 'RHR',       col: 2 },
    { name: 'Readiness', col: 16 },
  ].forEach(m => {
    const thisAvg = _avgCol(thisWeekHealth, m.col);
    const baseAvg = _avgCol(ninetyDayHealth, m.col);
    const baseSd  = _stdevCol(ninetyDayHealth, m.col);
    if (thisAvg && baseAvg && baseSd && baseSd > 0) {
      const z = (thisAvg - baseAvg) / baseSd;
      if (Math.abs(z) >= 2) {
        anomalies.push(`${m.name} z-score ${z.toFixed(2)} (this ${thisAvg.toFixed(1)} vs baseline ${baseAvg.toFixed(1)} ±${baseSd.toFixed(1)})`);
      }
    }
  });

  // --- Strategist: close to a stated goal ---
  // Weight within 3 lbs of the athlete's stated weight target (if any),
  // OR a long run >= 10 mi happened this week and the athlete has a
  // half-marathon goal. Both thresholds derive from ATHLETE_PROFILE.goals.
  const goalProgress = [];
  const thisWeight = _avgCol(thisWeekHealth, 1);
  const targetWeight = _targetWeightLbs();
  if (targetWeight && thisWeight && Math.abs(thisWeight - targetWeight) <= 3) {
    goalProgress.push(`Weight ${thisWeight.toFixed(1)} within 3 lbs of ${targetWeight} goal`);
  }
  if (_hasHalfMarathonGoal() && thisWeekWorkoutRows && thisWeekWorkoutRows.length) {
    // Distance is column 5 in Merged Workouts (Date/Time=1, Type=2, Name=3, Duration=4, Distance=5)
    // A 10+ mi run is a half-marathon-readiness milestone.
    const longRun = thisWeekWorkoutRows.find(r => {
      const type = String(r[2] || '').toLowerCase();
      const dist = parseFloat(r[5]) || 0;
      return type.includes('run') && dist >= 10;
    });
    if (longRun) {
      goalProgress.push(`Long run ${(parseFloat(longRun[5]) || 0).toFixed(1)} mi (half-marathon milestone)`);
    }
  }

  // --- Physiologist: classic fitness adaptation ---
  // Training load up vs baseline AND RHR trending lower vs baseline.
  // Rough proxy: this-week RHR avg is 2+ bpm below 90-day RHR avg.
  const adaptationNotes = [];
  const thisRHRavg = _avgCol(thisWeekHealth, 2);
  const baseRHRavg = _avgCol(ninetyDayHealth, 2);
  if (thisRHRavg && baseRHRavg && (baseRHRavg - thisRHRavg) >= 2) {
    adaptationNotes.push(`RHR ${thisRHRavg.toFixed(0)} vs baseline ${baseRHRavg.toFixed(0)} — sign of adaptation`);
  }

  return {
    newBests,
    allTimeBests,
    straightTalkTriggered,
    straightTalkReason: straightTalkTriggered ? violationDetails.join('; ') : null,
    violationDetails,
    softPatternDivergence,
    roughWeekNoViolation,
    anomalies,
    goalProgress,
    adaptationNotes,
  };
}

// Pick the lens for this week. Priority rules first; rotation fallback last.
// Returns { lens, reason } — `reason` is logged for debugging and isn't
// shown to the user.
function pickLensForWeek(ctx) {
  const signals = detectLensSignals(ctx);
  const recentNames = _getRecentLensNames(2);  // last two weeks

  // Priority triggers (first match wins)
  if (signals.newBests.length > 0) {
    const tag = signals.allTimeBests.length > 0 ? 'all-time best' : 'milestone best';
    return {
      lens: _getLens('Cheerleader'),
      reason: `${tag}: ${signals.newBests.join('; ')}`,
      celebrationSignals: signals.newBests,
    };
  }
  // Straight Talk — but not two weeks in a row. Serial lecturing doesn't
  // coach; it trains the athlete to tune out. If last week was Straight Talk,
  // route to a different voice (Storyteller → Physiologist → Strategist) and
  // pass the pattern details through so the new voice can still surface the
  // divergence — just in its own register.
  if (signals.straightTalkTriggered) {
    if (recentNames[0] === 'Straight Talk') {
      const ladder = ['Storyteller', 'Physiologist', 'Strategist'];
      const pick = ladder.find(name => !recentNames.includes(name)) || ladder[0];
      return {
        lens: _getLens(pick),
        reason: `pattern divergence w/ cost, but Straight Talk used last week — rotating to ${pick}`,
        patternNotes: signals.violationDetails,
        continuingPattern: true,
      };
    }
    return {
      lens: _getLens('Straight Talk'),
      reason: `pattern divergence w/ physiological cost: ${signals.straightTalkReason}`,
      patternNotes: signals.violationDetails,
    };
  }

  // Soft pattern divergence (suggestions not followed, but no cost this week):
  // surface as notes via Storyteller/Physiologist, not as a scolding.
  if (signals.softPatternDivergence) {
    const ladder = ['Storyteller', 'Physiologist', 'Strategist'];
    const pick = ladder.find(name => !recentNames.includes(name)) || ladder[0];
    return {
      lens: _getLens(pick),
      reason: `soft pattern divergence (no cost signal): ${signals.violationDetails.join('; ')}`,
      patternNotes: signals.violationDetails,
    };
  }

  if (signals.anomalies.length > 0) {
    return { lens: _getLens('Detective'), reason: `anomaly: ${signals.anomalies.join('; ')}` };
  }
  if (signals.goalProgress.length > 0) {
    return { lens: _getLens('Strategist'), reason: `goal progress: ${signals.goalProgress.join('; ')}` };
  }
  if (signals.adaptationNotes.length > 0) {
    return { lens: _getLens('Physiologist'), reason: `adaptation: ${signals.adaptationNotes.join('; ')}` };
  }

  // Fallback rotation. On a rough-but-blameless week (regressions or low
  // readiness with no divergence), skew toward lenses that can handle
  // difficulty with empathy. Never drop Cheerleader on a rough week; that'd
  // read as tone-deaf.
  const fallbackOrder = signals.roughWeekNoViolation
    ? ['Storyteller', 'Physiologist', 'Strategist']
    : ['Storyteller', 'Physiologist', 'Strategist', 'Cheerleader'];
  const available = fallbackOrder.filter(name => !recentNames.includes(name));
  const pick = available[0] || fallbackOrder[0];
  const tag = signals.roughWeekNoViolation ? 'rough week (no violation) rotation' : 'fallback rotation';
  return { lens: _getLens(pick), reason: `${tag} (recent: ${recentNames.join(', ') || 'none'})` };
}
