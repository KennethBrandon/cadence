// =============================================================================
// TEST HELPERS — run the coach for historical weeks, useful for:
//   1. Validating output for a known past week
//   2. Seeding Report History with several weeks so the "prior reports"
//      memory feature has real content before the first live run
//
// USAGE (from the Apps Script editor):
//   testWeeklyCoachReportForDate('2026-04-11')
//     → runs a report for the 7 days ending 2026-04-11 (inclusive)
//
//   backfillWeeklyReports(4)
//     → runs 4 historical reports, oldest first, each ending one week
//       before the next. Emails each one. By the time the most recent
//       runs, it will have 3 prior reports in its memory block.
//
//   backfillWeeklyReports(4, { skipEmail: true })
//     → same thing but doesn't email — useful when you just want to
//       seed history without flooding your inbox.
//
// NOTE: runs are always saved to the Report History sheet (that's the
// whole point). If you re-run the same week twice, you'll get a duplicate
// row — clear the sheet manually between test batches if that matters.
// =============================================================================

function testWeeklyCoachReportForDate(endDateStr, options) {
  options = options || {};
  // The main function uses `today - 1 day` as thisWeekEnd. So to make the
  // window END on endDateStr, we pass referenceDate = endDateStr + 1 day.
  const endDate = new Date(endDateStr + 'T12:00:00');
  if (isNaN(endDate)) {
    throw new Error(`Invalid date: "${endDateStr}". Use YYYY-MM-DD format.`);
  }
  const referenceDate = new Date(endDate);
  referenceDate.setDate(referenceDate.getDate() + 1);

  return sendWeeklyCoachReport({
    referenceDate: referenceDate,
    skipEmail: options.skipEmail === true,
  });
}

// Zero-arg wrappers — pick from the Apps Script editor's function dropdown
// and hit Run. Useful because the dropdown can't pass arguments.
function rerunWeekApr19to25() {
  return testWeeklyCoachReportForDate('2026-04-25');
}

function backfill5Weeks() {
  return backfillWeeklyReports(5);
}

function backfill5WeeksNoEmail() {
  return backfillWeeklyReports(5, { skipEmail: true });
}

function backfillWeeklyReports(weeksBack, options) {
  options = options || {};
  if (!weeksBack || weeksBack < 1) {
    throw new Error('backfillWeeklyReports(weeksBack): weeksBack must be >= 1');
  }
  const skipEmail = options.skipEmail === true;

  console.log(`Backfilling ${weeksBack} weekly reports (skipEmail=${skipEmail})...`);

  // Build the list of week-end dates, OLDEST first, so memory accumulates
  // in the right order when buildPriorReportsContext() runs on each pass.
  //
  // The newest week in the backfill must match what a live run today would
  // produce — sendWeeklyCoachReport uses `today - 1 day` as thisWeekEnd, so
  // we anchor the freshest week-end at `today - 1`. Previously we used
  // `today - 7i` starting at i=weeksBack, which gave weeks ending 7–35 days
  // ago and skipped the most recent complete week entirely.
  const today = new Date();
  const weekEnds = [];
  for (let i = weeksBack - 1; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - 1 - (7 * i));   // i=0 → yesterday, i=N → N weeks before that
    d.setHours(23, 59, 59, 999);
    weekEnds.push(d);
  }

  weekEnds.forEach((weekEnd, idx) => {
    const label = Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    console.log(`\n=== Backfill ${idx + 1}/${weekEnds.length}: week ending ${label} ===`);
    try {
      // referenceDate = day AFTER the desired week end (so thisWeekEnd = weekEnd)
      const ref = new Date(weekEnd);
      ref.setDate(ref.getDate() + 1);
      sendWeeklyCoachReport({ referenceDate: ref, skipEmail: skipEmail });
    } catch (e) {
      console.error(`Backfill for ${label} failed: ${e.message}`);
    }
    // Apps Script rate-limits MailApp and UrlFetchApp; a small pause between
    // runs keeps us well clear of both quotas.
    Utilities.sleep(2000);
  });

  console.log(`Backfill complete — ${weekEnds.length} reports processed.`);
}
