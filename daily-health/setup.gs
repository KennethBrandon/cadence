/**
 * One-time setup: creates Weekly Summary and Monthly Summary sheets
 * with live QUERY formulas sourced from Daily Data.
 * Run manually from the Apps Script editor after clasp push.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupWeeklySummary_(ss);
  setupMonthlySummary_(ss);
  console.log('setupSheets() complete.');
}

function setupWeeklySummary_(ss) {
  let sheet = ss.getSheetByName('Weekly Summary');
  if (!sheet) sheet = ss.insertSheet('Weekly Summary');
  else sheet.clearContents();

  sheet.getRange('A1').setFormula(
    `=ARRAYFORMULA(QUERY({'Daily Data'!A:A, IF(ROW('Daily Data'!A:A)=1, "Week Starting", 'Daily Data'!A:A - WEEKDAY('Daily Data'!A:A) + 1), 'Daily Data'!B:Q}, "select Col2, avg(Col3), avg(Col4), avg(Col5), avg(Col6), avg(Col7), avg(Col8), avg(Col9)/60, avg(Col10), avg(Col11)/60, avg(Col12), avg(Col13), avg(Col14), avg(Col15), avg(Col16), avg(Col17), avg(Col18) where Col1 is not null group by Col2 order by Col2 asc label Col2 'Week Starting', avg(Col3) 'Avg Weight', avg(Col4) 'Avg RHR', avg(Col5) 'Avg HRV', avg(Col9)/60 'Avg Total Sleep (Hrs)', avg(Col11)/60 'Avg Deep Sleep (Hrs)', avg(Col18) 'Avg Readiness'", 1))`
  );
}

function setupMonthlySummary_(ss) {
  let sheet = ss.getSheetByName('Monthly Summary');
  if (!sheet) sheet = ss.insertSheet('Monthly Summary');
  else sheet.clearContents();

  sheet.getRange('A1').setFormula(
    `=ARRAYFORMULA(QUERY({'Daily Data'!A:A, TEXT('Daily Data'!A:A, "mmmm"), 'Daily Data'!B:Q}, "select year(Col1), Col2, avg(Col3), avg(Col4), avg(Col5), avg(Col6), avg(Col7), avg(Col8), avg(Col9)/60, avg(Col10), avg(Col11)/60, avg(Col12), avg(Col13), avg(Col14), avg(Col15), avg(Col16), avg(Col17), avg(Col18), month(Col1)+1 where Col1 is not null group by year(Col1), Col2, month(Col1)+1 order by year(Col1) asc, month(Col1)+1 asc label year(Col1) 'Year', Col2 'Month', avg(Col3) 'Avg Weight', avg(Col4) 'Avg RHR', avg(Col5) 'Avg HRV', avg(Col9)/60 'Avg Total Sleep (Hrs)', avg(Col11)/60 'Avg Deep Sleep (Hrs)', avg(Col18) 'Avg Readiness', month(Col1)+1 'MonthNum'", 1))`
  );
}
