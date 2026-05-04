// =============================================================================
// INTERESTING STUFF ANGLE — picks the lead framing for the "Interesting Stuff"
// section so the AI has a clear focus each week (instead of defaulting to the
// same "you vs 2 years ago" every time). Skips angles that overlap with the
// lens choice (e.g. APPROACHING_PR doesn't fire when Cheerleader already
// celebrated a new best).
// Priority order:
//   1. APPROACHING_PR      — within 5% of an all-time best
//   2. OUTLIER             — any key metric z-score 1.5 ≤ |z| < 2 vs 90-day baseline
//   3. LONGTERM_PROGRESS   — meaningful positive 1yr/2yr delta AND not used in last 3 reports
//   4. DISCONNECT          — two normally-correlated metrics diverged
//   5. STREAK              — 4+ consecutive days above/below the 90-day baseline
//   6. TRAJECTORY_SHIFT    — 30-day avg has drifted ≥ 0.5 SD from 90-day avg
//   7. FRESH_ANGLE         — fallback; AI picks the hook, YoY discouraged (not banned)
// =============================================================================
function detectInterestingAngle(ctx, lensName) {
  const { thisWeekHealth, thirtyDayHealth, ninetyDayHealth, personalBests } = ctx;

  const thisHRVs       = thisWeekHealth.map(r => parseFloat(r[3]) || 0).filter(v => v > 0);
  const thisRHRs       = thisWeekHealth.map(r => parseFloat(r[2]) || 0).filter(v => v > 0);
  const thisReadinesss = thisWeekHealth.map(r => parseFloat(r[16]) || 0).filter(v => v > 0);

  // ---- 1. APPROACHING_PR (skip if Cheerleader already fired) ----
  if (lensName !== 'Cheerleader') {
    const prChecks = [
      { label: 'HRV',       vals: thisHRVs,       best: personalBests.bestHRV,       higherBetter: true,  fmt: v => `${v.toFixed(1)} ms` },
      { label: 'RHR',       vals: thisRHRs,       best: personalBests.bestRHR,       higherBetter: false, fmt: v => `${v.toFixed(0)} bpm` },
      { label: 'Readiness', vals: thisReadinesss, best: personalBests.bestReadiness, higherBetter: true,  fmt: v => `${v.toFixed(0)}` },
    ];
    for (const c of prChecks) {
      if (!c.best || c.vals.length === 0) continue;
      const thisExtreme = c.higherBetter ? Math.max(...c.vals) : Math.min(...c.vals);
      // Within 5% of best, but not past it (past = Cheerleader territory)
      const pctOfBest = c.higherBetter
        ? thisExtreme / c.best
        : c.best / thisExtreme;
      if (pctOfBest >= 0.95 && pctOfBest < 1.0) {
        return {
          angle: 'APPROACHING_PR',
          leadData: `This week's best ${c.label} hit ${c.fmt(thisExtreme)} — ${(pctOfBest * 100).toFixed(0)}% of the all-time best (${c.fmt(c.best)}). One more solid day could tie or beat it.`,
          framing: `Frame as "knocking on the door." What specifically set up that peak day? What would push it over the line next week?`,
        };
      }
    }
  }

  // ---- 2. OUTLIER (skip if Detective already fired) ----
  if (lensName !== 'Detective') {
    const metrics = [
      { label: 'HRV',          col: 3,  higherBetter: true,  unit: 'ms' },
      { label: 'RHR',          col: 2,  higherBetter: false, unit: 'bpm' },
      { label: 'Readiness',    col: 16, higherBetter: true,  unit: '' },
      { label: 'Total Sleep',  col: 7,  higherBetter: true,  unit: 'min' },
      { label: 'Deep Sleep',   col: 9,  higherBetter: true,  unit: 'min' },
    ];
    let biggest = null;
    for (const m of metrics) {
      const thisAvg = _avgCol(thisWeekHealth, m.col);
      const baseAvg = _avgCol(ninetyDayHealth, m.col);
      const baseSd  = _stdevCol(ninetyDayHealth, m.col);
      if (!thisAvg || !baseAvg || !baseSd || baseSd === 0) continue;
      const z = (thisAvg - baseAvg) / baseSd;
      // 1.2 ≤ |z| < 2 — Detective handles ≥2 so we only pick the "notable but
      // not extreme" range here. (Loosened from 1.5 to 1.2 to reduce how often
      // we fall through to FRESH_ANGLE on otherwise normal weeks.)
      if (Math.abs(z) >= 1.2 && Math.abs(z) < 2) {
        if (!biggest || Math.abs(z) > Math.abs(biggest.z)) {
          biggest = { ...m, thisAvg, baseAvg, baseSd, z };
        }
      }
    }
    if (biggest) {
      const direction = biggest.z > 0 ? 'above' : 'below';
      const good = (biggest.higherBetter === (biggest.z > 0)) ? 'positive' : 'negative';
      return {
        angle: 'OUTLIER',
        leadData: `${biggest.label} ran ${direction} baseline by ${Math.abs(biggest.z).toFixed(1)} SD this week (${biggest.thisAvg.toFixed(1)} vs 90-day avg ${biggest.baseAvg.toFixed(1)} ± ${biggest.baseSd.toFixed(1)}). That's a ${good} move worth understanding.`,
        framing: `Ask: what specifically drove this? Look at workouts, sleep, prior-week fatigue, or anything in the notes that could explain it. Speculate if needed, but label speculation.`,
      };
    }
  }

  // ---- 3. LONGTERM_PROGRESS — lifetime-arc improvement, throttled ----
  // Fires only if we haven't used this angle in the last 3 reports AND there's
  // a meaningful positive delta vs one of FOUR anchors: when tracking began,
  // lifetime average, 2 years ago, or 1 year ago. Picks the most impressive
  // comparison. Skips on regressions (Straight Talk / Detective handle bad news).
  // Throttle: must not have fired in the last 4 reports (up from 3 — at 3 we
  // were hitting the boundary back-to-back, which felt repetitive).
  const recentAngles = _getRecentAngles(4);
  if (!recentAngles.includes('LONGTERM_PROGRESS')) {
    const pb = ctx.personalBests || {};

    // Per-metric rules. `lifetimeAvgKey` points to the lifetime avg on
    // personalBests; `allValuesKey` points to the raw daily array for
    // percentile ranking.
    const ltpChecks = [
      { label: 'HRV',       col: 3,  higherBetter: true,  minDelta: 4,   fmt: v => `${v.toFixed(1)} ms`,     lifetimeAvgKey: 'avgHRV',       allValuesKey: 'allHRV' },
      { label: 'RHR',       col: 2,  higherBetter: false, minDelta: 3,   fmt: v => `${v.toFixed(0)} bpm`,    lifetimeAvgKey: 'avgRHR',       allValuesKey: 'allRHR' },
      { label: 'Readiness', col: 16, higherBetter: true,  minDelta: 5,   fmt: v => `${v.toFixed(0)}/100`,    lifetimeAvgKey: 'avgReadiness', allValuesKey: 'allReadiness' },
      { label: 'Weight',    col: 1,  higherBetter: false, minDelta: 5,   fmt: v => `${v.toFixed(1)} lbs`,    lifetimeAvgKey: 'avgWeight',    allValuesKey: 'allWeight' },
      { label: 'Sleep',     col: 7,  higherBetter: true,  minDelta: 20,  fmt: v => `${(v / 60).toFixed(1)} hrs`, lifetimeAvgKey: 'avgSleep', allValuesKey: 'allSleep' },
    ];

    // Only use first-period anchor once we have enough separation from it
    // (otherwise we'd be comparing this week to data that overlaps recent history).
    const canUseFirstPeriod = (pb.totalDays || 0) >= 180;

    const anchors = [
      canUseFirstPeriod && ctx.firstPeriodHealth && ctx.firstPeriodHealth.length >= 30
        ? { kind: 'firstPeriod', rows: ctx.firstPeriodHealth, label: 'when you started tracking' }
        : null,
      { kind: 'lifetimeAvg', label: 'your lifetime average' },
      ctx.twoYearHealth && ctx.twoYearHealth.length ? { kind: 'rows', rows: ctx.twoYearHealth, label: '2 years ago this week' } : null,
      ctx.oneYearHealth && ctx.oneYearHealth.length ? { kind: 'rows', rows: ctx.oneYearHealth, label: '1 year ago this week' } : null,
    ].filter(Boolean);

    // Percentile helper: what % of all logged days is this week's avg at-least-as-good-as?
    const topPercentile = (val, allValues, higherBetter) => {
      if (!allValues || allValues.length === 0) return null;
      const matches = higherBetter
        ? allValues.filter(v => v >= val).length
        : allValues.filter(v => v <= val).length;
      return (matches / allValues.length) * 100;
    };

    const candidates = [];
    for (const c of ltpChecks) {
      const thisAvg = _avgCol(thisWeekHealth, c.col);
      if (!thisAvg) continue;

      for (const anchor of anchors) {
        let compAvg = null;
        if (anchor.kind === 'lifetimeAvg') {
          compAvg = c.lifetimeAvgKey ? pb[c.lifetimeAvgKey] : null;
        } else if (anchor.rows) {
          compAvg = _avgCol(anchor.rows, c.col);
        }
        if (!compAvg) continue;

        const delta = c.higherBetter ? (thisAvg - compAvg) : (compAvg - thisAvg);
        if (delta >= c.minDelta) {
          candidates.push({
            ...c,
            delta,
            thisAvg,
            compAvg,
            anchorKind: anchor.kind,
            anchorLabel: anchor.label,
            // Bonus weight for firstPeriod — the "starting line" story
            // is the most motivating framing, so bias toward it.
            score: (delta / c.minDelta) * (anchor.kind === 'firstPeriod' ? 1.4 : 1.0),
          });
        }
      }
    }

    if (candidates.length > 0) {
      candidates.sort((a, b) => b.score - a.score);
      const top = candidates[0];
      const pct = topPercentile(top.thisAvg, pb[top.allValuesKey], top.higherBetter);
      const direction = top.higherBetter ? 'up from' : 'down from';
      const totalDays = pb.totalDays || 0;
      const yearsTracked = (totalDays / 365).toFixed(1);

      let leadData = `${top.label} averaged ${top.fmt(top.thisAvg)} this week — ${direction} ${top.fmt(top.compAvg)} ${top.anchorLabel}.`;
      if (pct !== null && pct <= 25) {
        leadData += ` That puts this week in the top ${pct.toFixed(0)}% of the ${totalDays} days you've ever logged.`;
      }

      return {
        angle: 'LONGTERM_PROGRESS',
        leadData,
        framing: `This is the "look how far you've come" moment. Scale: ${totalDays} days (~${yearsTracked} years) of tracking. Connect the dots from then to now — what consistent habits, training choices, or rehab work got the athlete here? Keep it grounded and specific (one or two concrete drivers), not a victory lap. If another metric also improved meaningfully vs the same anchor, it's fine to mention it briefly, but ${top.label} is the lead.`,
      };
    }
  }

  // ---- 4. DISCONNECT — two correlated metrics moved opposite ways ----
  const thisReadinessAvg = _avgCol(thisWeekHealth, 16);
  const baseReadinessAvg = _avgCol(ninetyDayHealth, 16);
  const thisSleepAvg     = _avgCol(thisWeekHealth, 7);
  const baseSleepAvg     = _avgCol(ninetyDayHealth, 7);
  const thisHRVavg       = _avgCol(thisWeekHealth, 3);
  const baseHRVavg       = _avgCol(ninetyDayHealth, 3);

  // Good sleep, bad readiness — something non-training is stressing the system
  if (thisSleepAvg && baseSleepAvg && thisReadinessAvg && baseReadinessAvg &&
      thisSleepAvg >= baseSleepAvg + 10 &&            // 10+ extra minutes (loosened from 15)
      thisReadinessAvg < baseReadinessAvg - 3) {      // 3+ readiness points below (loosened from 5)
    return {
      angle: 'DISCONNECT',
      leadData: `Sleep ran ${(thisSleepAvg / 60).toFixed(1)} hrs vs 90-day avg ${(baseSleepAvg / 60).toFixed(1)} hrs, but readiness dropped to ${thisReadinessAvg.toFixed(0)} (baseline ${baseReadinessAvg.toFixed(0)}). The usual sleep→readiness link broke.`,
      framing: `When the obvious recovery input isn't delivering recovery output, something else is in the way. Look at training spikes, life stress, illness, travel, or nutrition. Name the most likely culprit.`,
    };
  }

  // HRV holding or climbing despite elevated training — resilience
  // (We'd need tlData here to compare training load properly; approximate
  // by asking: HRV above baseline AND at least 1 hard workout logged.)
  const heavyWorkouts = (ctx.thisWeekWorkoutRows || []).filter(r => {
    const relEffort = parseFloat(r[12]) || 0;   // Relative Effort col index 12
    const dist = parseFloat(r[5]) || 0;
    return relEffort >= 70 || dist >= 8;
  });
  if (thisHRVavg && baseHRVavg && heavyWorkouts.length >= 2 &&
      thisHRVavg >= baseHRVavg + 1) {
    return {
      angle: 'DISCONNECT',
      leadData: `${heavyWorkouts.length} demanding sessions this week and HRV still held at ${thisHRVavg.toFixed(1)} ms (above 90-day avg ${baseHRVavg.toFixed(1)}). Normally this kind of load suppresses it.`,
      framing: `Call out the resilience — the nervous system is absorbing work it used to cough on. Point at what's enabling that (sleep, consistency, base fitness).`,
    };
  }

  // ---- 5. STREAK — consecutive days above/below baseline ----
  // Sort this-week rows chronologically; count longest consecutive run above
  // the 90-day baseline for HRV and Readiness.
  const sortedRows = [...thisWeekHealth].sort((a, b) => {
    const da = a[0] instanceof Date ? a[0] : new Date(a[0]);
    const db = b[0] instanceof Date ? b[0] : new Date(b[0]);
    return da - db;
  });

  function longestRun(valuesAbove) {
    let best = 0, run = 0;
    for (const v of valuesAbove) {
      if (v === null) { run = 0; continue; }
      if (v) { run++; best = Math.max(best, run); } else { run = 0; }
    }
    return best;
  }

  for (const m of [
    { label: 'HRV',       col: 3,  baseline: baseHRVavg,       higherBetter: true },
    { label: 'Readiness', col: 16, baseline: baseReadinessAvg, higherBetter: true },
  ]) {
    if (!m.baseline) continue;
    const flags = sortedRows.map(r => {
      const v = parseFloat(r[m.col]);
      if (!v || v <= 0) return null;
      return m.higherBetter ? v >= m.baseline : v <= m.baseline;
    });
    const streak = longestRun(flags);
    if (streak >= 3) {  // loosened from 4 — 3 consecutive days is still a real pattern
      return {
        angle: 'STREAK',
        leadData: `${m.label} has been at or above its 90-day baseline (${m.baseline.toFixed(1)}) for ${streak} consecutive days this week.`,
        framing: `Call out the momentum. What's been consistent across those days — sleep, training pattern, something else? Momentum like this tends to compound if protected.`,
      };
    }
  }

  // ---- 6. TRAJECTORY_SHIFT — 30-day avg drifting from 90-day avg ----
  let bestShift = null;
  for (const m of [
    { label: 'HRV',       col: 3,  higherBetter: true },
    { label: 'RHR',       col: 2,  higherBetter: false },
    { label: 'Readiness', col: 16, higherBetter: true },
  ]) {
    const avg30 = _avgCol(thirtyDayHealth, m.col);
    const avg90 = _avgCol(ninetyDayHealth, m.col);
    const sd90  = _stdevCol(ninetyDayHealth, m.col);
    if (!avg30 || !avg90 || !sd90 || sd90 === 0) continue;
    const shiftSD = (avg30 - avg90) / sd90;
    if (Math.abs(shiftSD) >= 0.3) {  // loosened from 0.5 — smaller drifts are still trends worth naming
      if (!bestShift || Math.abs(shiftSD) > Math.abs(bestShift.shiftSD)) {
        bestShift = { ...m, avg30, avg90, sd90, shiftSD };
      }
    }
  }
  if (bestShift) {
    const direction = bestShift.shiftSD > 0 ? 'climbed' : 'dropped';
    const good = (bestShift.higherBetter === (bestShift.shiftSD > 0)) ? 'positive' : 'concerning';
    return {
      angle: 'TRAJECTORY_SHIFT',
      leadData: `${bestShift.label} has ${direction} over the last 30 days — averaging ${bestShift.avg30.toFixed(1)} vs the 90-day baseline of ${bestShift.avg90.toFixed(1)} (${Math.abs(bestShift.shiftSD).toFixed(1)} SD shift, ${good}).`,
      framing: `This isn't a one-week fluke — it's a trend. Frame what's driving the shift (training block, sleep habits, life changes) and whether it should be preserved or corrected.`,
    };
  }

  // ---- 7. FRESH_ANGLE — fallback ----
  return {
    angle: 'FRESH_ANGLE',
    leadData: `Nothing jumps out statistically this week.`,
    framing: `Find something genuinely surprising in the week's data — an unexpected correlation, a pattern in workout timing, a subtle training-recovery link. Year-over-year comparisons are allowed but prefer a fresher angle; if you do reach for YoY, make sure the comparison is specific and earns its place (don't use it as filler).`,
  };
}
