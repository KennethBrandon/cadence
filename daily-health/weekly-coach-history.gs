// =============================================================================
// REPORT HISTORY — save each week's report and feed the recent ones back into
// the next prompt, so the AI can avoid repeating itself and reference follow-
// through on its own prior advice.
// =============================================================================
const REPORT_HISTORY_SHEET = 'Report History';
const RECENT_REPORTS_TO_REMEMBER = 3;

function getOrCreateReportHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(REPORT_HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REPORT_HISTORY_SHEET);
    sheet.appendRow(['Generated At', 'Week Label', 'Lens', 'Angle', 'Full Report']);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(5, 600);
  } else {
    // Auto-migrate older sheets that don't yet have the Angle column.
    // Old layout: Generated At | Week Label | Lens | Full Report
    // New layout: Generated At | Week Label | Lens | Angle | Full Report
    const headerRange = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1));
    const headers = headerRange.getValues()[0];
    if (headers.length < 5 || headers[3] !== 'Angle') {
      sheet.insertColumnAfter(3);                       // new col D
      sheet.getRange(1, 4).setValue('Angle');
      sheet.setColumnWidth(5, 600);
    }
  }
  return sheet;
}

function saveReportToHistory(weekLabel, lens, angle, reportText) {
  const sheet = getOrCreateReportHistorySheet();
  sheet.appendRow([new Date(), weekLabel, lens || '', angle || '', reportText]);
}

// Pull a section by keyword-matching the `###` header. Lenient on the emoji
// since the model sometimes reorders/swaps it.
function extractReportSection(text, headerKeyword) {
  if (!text) return '';
  const pattern = new RegExp(
    `###[^\\n]*${headerKeyword}[^\\n]*\\n([\\s\\S]*?)(?=\\n###|$)`,
    'i'
  );
  const m = text.match(pattern);
  return m ? m[1].trim() : '';
}

function buildPriorReportsContext() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REPORT_HISTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    return "No prior reports on file yet — this is the first one (or the first since history was added).";
  }

  const data = sheet.getDataRange().getValues();
  // Take the most recent N rows (last rows of the sheet)
  const rows = data.slice(1).slice(-RECENT_REPORTS_TO_REMEMBER).reverse();

  return rows.map((r, i) => {
    // Sheet layout: [Generated At, Week Label, Lens, Angle, Full Report]
    const [, weekLabel, lens, angle, fullReport] = r;
    const headline = extractReportSection(fullReport, 'headline');
    const interesting = extractReportSection(fullReport, 'interesting');
    const nextWeek = extractReportSection(fullReport, 'next week');
    const weeksAgo = i === 0 ? 'LAST WEEK' : `${i + 1} WEEKS AGO`;
    const lensNote = lens ? ` — lens: ${lens}` : '';
    const angleNote = angle ? ` — angle: ${angle}` : '';
    const parts = [`--- ${weeksAgo} (${weekLabel})${lensNote}${angleNote} ---`];
    if (headline) parts.push(`Headline: ${headline}`);
    if (interesting) parts.push(`Interesting Stuff framing used:\n${interesting}`);
    if (nextWeek) parts.push(`Action items given:\n${nextWeek}`);
    return parts.join('\n');
  }).join('\n\n');
}
