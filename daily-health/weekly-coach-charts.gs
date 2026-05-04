// =============================================================================
// CHART BUILDERS — generate chart images as blobs for inline email.
// NOTE: .setName('*.png') is required so MailApp embeds the image properly.
// =============================================================================
function buildBarChart(title, data, color, refLine) {
  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Day')
    .addColumn(Charts.ColumnType.NUMBER, title);

  data.forEach(row => dataTable.addRow(row));

  const chartBuilder = Charts.newColumnChart()
    .setDataTable(dataTable.build())
    .setTitle(title)
    .setDimensions(600, 250)
    .setColors([color])
    .setOption('legend', { position: 'none' })
    .setOption('chartArea', { width: '88%', height: '74%', left: '10%', top: '14%' })
    .setOption('vAxis', { minValue: 0, maxValue: 100 });

  if (refLine) {
    chartBuilder.setOption('vAxis', {
      minValue: 0, maxValue: 100,
      ticks: [0, 25, 50, refLine, 100]
    });
  }

  return chartBuilder.build().getAs('image/png').setName('barChart.png');
}

function buildStackedChart(title, data, seriesNames, colors, refLine) {
  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Day');

  seriesNames.forEach(name => dataTable.addColumn(Charts.ColumnType.NUMBER, name));
  data.forEach(row => dataTable.addRow(row));

  const vAxis = { minValue: 0 };
  if (refLine) {
    // Put a labeled gridline at the target (e.g. 8 hrs) and headroom above it
    vAxis.maxValue = Math.max(10, refLine + 2);
    vAxis.ticks = [0, 2, 4, 6, { v: refLine, f: `${refLine} hr target` }, vAxis.maxValue];
  }

  const chartBuilder = Charts.newColumnChart()
    .setDataTable(dataTable.build())
    .setTitle(title)
    .setDimensions(600, 250)
    .setColors(colors)
    .setStacked()
    .setOption('legend', { position: 'bottom' })
    .setOption('chartArea', { width: '88%', height: '68%', left: '10%', top: '14%' })
    .setOption('vAxis', vAxis);

  return chartBuilder.build().getAs('image/png').setName('stackedChart.png');
}

// opts: { refLine, dualAxis, interpolate, sparseXLabels }
//   refLine    — number; draws a labeled vAxis gridline (e.g. 7.5 hr sleep target).
//   dualAxis   — boolean; with 2 series, series 0 → left axis, series 1 → right.
//   interpolate — boolean; bridge null gaps in the line (useful for sparse
//                 metrics like weigh-ins that aren't recorded daily).
//   sparseXLabels — boolean; caller has set most x-axis labels to '' so only a
//                 handful render (e.g. Jan/Jul anchors over years of weekly bins).
//                 Forces showTextEvery: 1 so the auto-thinner doesn't pick
//                 evenly-spaced slots that all happen to be empty.
// X-axis is always STRING-typed — the Apps Script Charts service throws
// "a.getTime is not a function" on DATE columns combined with hAxis.ticks,
// so callers handle their own label thinning by emptying non-anchor labels.
function buildLineChart(title, data, seriesNames, colors, opts) {
  opts = opts || {};

  // refLine (single-axis only) is rendered as a dashed companion series at a
  // constant value, so it shows up as its own line with a legend entry.
  const useRefLineSeries = opts.refLine && !opts.dualAxis;
  const chartData = useRefLineSeries
    ? data.map(row => row.concat([opts.refLine]))
    : data;
  const chartSeriesNames = useRefLineSeries
    ? seriesNames.concat([`${opts.refLine} hr target`])
    : seriesNames;
  const chartColors = useRefLineSeries
    ? colors.concat(['#d93025'])
    : colors;

  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Day');

  chartSeriesNames.forEach(name => dataTable.addColumn(Charts.ColumnType.NUMBER, name));
  chartData.forEach(row => dataTable.addRow(row));

  // Dual-axis charts need horizontal room on both sides for the titled axes;
  // single-axis charts can lean tighter on the right.
  const chartArea = opts.dualAxis
    ? { width: '78%', height: '68%', left: '11%', top: '14%' }
    : { width: '88%', height: '68%', left: '10%', top: '14%' };

  const chartBuilder = Charts.newLineChart()
    .setDataTable(dataTable.build())
    .setTitle(title)
    .setDimensions(600, 250)
    .setColors(chartColors)
    .setCurveStyle(Charts.CurveStyle.SMOOTH)
    .setOption('legend', { position: 'bottom' })
    .setOption('chartArea', chartArea)
    .setOption('interpolateNulls', opts.interpolate === true);

  // When the caller has labeled only a few anchor slots (rest are ''), the
  // default auto-thinner picks evenly-spaced positions and lands on empty
  // slots, rendering nothing. showTextEvery: 1 makes every slot eligible so
  // the few non-empty labels actually render. We don't force a slant — Charts
  // chooses based on label fit, and short anchors like "Jan 24" stay flat.
  if (opts.sparseXLabels) {
    chartBuilder.setOption('hAxis', {
      showTextEvery: 1,
      textStyle: { fontSize: 10 },
    });
  }

  if (opts.dualAxis && seriesNames.length === 2) {
    // Series 0 → left axis, series 1 → right axis. Lets two series with
    // different units (e.g. bpm + ms) each use their own scale instead of
    // getting flattened against a shared range. Each axis is titled with its
    // series name and tick labels are colored to match the line.
    //
    // Auto-compute a tight viewWindow per axis so each line uses most of its
    // vertical band — default scaling leaves a lot of headroom that flattens
    // the visible variation.
    const tightWindow = (colIdx) => {
      const vals = data.map(r => r[colIdx]).filter(v => typeof v === 'number' && isFinite(v));
      if (vals.length === 0) return null;
      const min = Math.min.apply(null, vals);
      const max = Math.max.apply(null, vals);
      const span = Math.max(1, max - min);
      const pad = span * 0.1;
      return { min: min - pad, max: max + pad };
    };
    const w0 = tightWindow(1);
    const w1 = tightWindow(2);

    const axis0 = {
      title: seriesNames[0],
      textStyle: { color: colors[0] },
      titleTextStyle: { color: colors[0], italic: false, bold: true },
    };
    if (w0) axis0.viewWindow = w0;
    const axis1 = {
      title: seriesNames[1],
      textStyle: { color: colors[1] },
      titleTextStyle: { color: colors[1], italic: false, bold: true },
    };
    if (w1) axis1.viewWindow = w1;

    chartBuilder
      .setOption('series', {
        0: { targetAxisIndex: 0 },
        1: { targetAxisIndex: 1 },
      })
      .setOption('vAxes', { 0: axis0, 1: axis1 });
  } else if (useRefLineSeries) {
    // Integer y-axis ticks spanning the data + target, with the appended
    // target series styled as a dashed line in a contrasting color.
    const refLine = opts.refLine;
    const numericVals = [];
    data.forEach(row => {
      for (let i = 1; i < row.length; i++) {
        const v = row[i];
        if (typeof v === 'number' && isFinite(v)) numericVals.push(v);
      }
    });
    const dataMin = numericVals.length ? Math.min.apply(null, numericVals) : refLine;
    const dataMax = numericVals.length ? Math.max.apply(null, numericVals) : refLine;
    const lo = Math.floor(Math.min(dataMin, refLine));
    const hi = Math.ceil(Math.max(dataMax, refLine));
    const ticks = [];
    for (let t = lo; t <= hi; t++) ticks.push(t);

    chartBuilder
      .setOption('vAxis', { minValue: lo, maxValue: hi, ticks: ticks })
      .setOption('series', {
        // Index of the appended target series (last column in chartSeriesNames).
        [chartSeriesNames.length - 1]: { lineDashStyle: [4, 4] },
      });
  }

  return chartBuilder.build().getAs('image/png').setName('lineChart.png');
}
