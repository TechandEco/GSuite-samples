/**
 * Installs a trigger in the Spreadsheet to run upon the Sheet being opened.
 * To learn more about triggers read:
 * https://developers.google.com/apps-script/guides/triggers
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Analytics Helper')
      .addItem('Setup data', 'setupData')
      .addItem('Create summary', 'createSummary')
      .addToUi();
};  
  
/**
 * Uses a macro to record steps in a Google Sheet into a script. This script
 * is called as a function from a custom menu to automatically enter dates and
 * values of the data we wish to pull via the Google Analytics add-on 
 * To learn more about macro recording: 
 * https://developers.google.com/apps-script/guides/sheets/macros
 * To learn more about the Google Analytics add-on:
 * https://developers.google.com/analytics/solutions/google-analytics-spreadsheet-add-on
 */
function setupData() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G3').activate();
  spreadsheet.getRange('B3:B7').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('G2').activate();
  spreadsheet.getCurrentCell().setValue('Top countries');
  spreadsheet.getRange('G4').activate();
  spreadsheet.getCurrentCell().setFormula('=DATE(YEAR(TODAY()),1,1)');
  spreadsheet.getRange('G5').activate();
  spreadsheet.getCurrentCell().setFormula('=TODAY()');
  spreadsheet.getRange('G7').activate();
  spreadsheet.getCurrentCell().setValue('ga:country');
  spreadsheet.getRange('B5').activate();
  spreadsheet.getCurrentCell().setFormula('=TODAY()');
  spreadsheet.getRange('B4').activate();
  spreadsheet.getCurrentCell().setFormula('=TODAY()-WEEKDAY(TODAY())-1');
  spreadsheet.getRange('C3').activate();
  spreadsheet.getRange('B3:B7').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('C2').activate();
  spreadsheet.getCurrentCell().setValue('Last week');
  spreadsheet.getRange('C5').activate();
  spreadsheet.getCurrentCell().setFormula('=B4-1');
  spreadsheet.getRange('C4').activate();
  spreadsheet.getCurrentCell().setFormula('=B4-7');
  spreadsheet.getRange('D3').activate();
  spreadsheet.getRange('B3:C7').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('D2').activate();
  spreadsheet.getCurrentCell().setValue('This year');
  spreadsheet.getRange('E2').activate();
  spreadsheet.getCurrentCell().setValue('Last year');
  spreadsheet.getRange('D7').activate();
  spreadsheet.getCurrentCell().setValue('ga:month');
  spreadsheet.getRange('E7').activate();
  spreadsheet.getCurrentCell().setValue('ga:month');
  spreadsheet.getRange('D4').activate();
  spreadsheet.getCurrentCell().setFormula('=DATE(YEAR(TODAY()),1,1)');
  spreadsheet.getRange('E4').activate();
  spreadsheet.getCurrentCell().setFormula('=D4-365');
  spreadsheet.getRange('F3').activate();
  spreadsheet.getRange('D3:D7').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('F2').activate();
  spreadsheet.getCurrentCell().setValue('Top browsers');
  spreadsheet.getRange('F7').activate();
  spreadsheet.getCurrentCell().setValue('ga:browser');
  spreadsheet.getRange('F8').activate();
};

/**
 * Uses a macro to record steps in a Google Sheet into a script. This script
 * is called as a function from a custom menu to automatically create summary
 * charts from the Google Analytics add-on such as unique pageviews by page
 * title, region, and top browsers used.
 * To learn more about macro recording: 
 * https://developers.google.com/apps-script/guides/sheets/macros
 * To learn more about the Google Analytics add-on:
 * https://developers.google.com/analytics/solutions/google-analytics-spreadsheet-add-on
 */
function createSummary() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('This week');
  spreadsheet.getRange('C1').activate();
  spreadsheet.getCurrentCell().setValue('Last week');
  spreadsheet.getRange('A2').activate();
  spreadsheet.getCurrentCell().setValue('Sun');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A2:A8'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('B2').activate();
  spreadsheet.getCurrentCell().setFormula('=\'This week\'!B16');
  spreadsheet.getRange('C2').activate();
  spreadsheet.getCurrentCell().setFormula('=\'Last week\'!B16');
  spreadsheet.getRange('B2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('B2:B8'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('C2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C2:C8'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('A1:C8').activate();
  var sheet = spreadsheet.getActiveSheet();
  var chart = sheet.newChart()
      .asColumnChart()
      .addRange(spreadsheet.getRange('A1:C8'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'false')
      .setOption('title', 'This week and Last week')
      .setPosition(3, 2, 79, 7)
      .build();
  sheet.insertChart(chart);
  var charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asColumnChart()
      .addRange(spreadsheet.getRange('A1:C8'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'false')
      .setOption('title', 'This week and Last week')
      .setOption('height', 246)
      .setOption('width', 398)
      .setPosition(3, 2, 79, 7)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asColumnChart()
      .addRange(spreadsheet.getRange('A1:C8'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'false')
      .setOption('title', 'This week and Last week')
      .setOption('height', 246)
      .setOption('width', 398)
      .setPosition(1, 4, 1, 0)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asColumnChart()
      .addRange(spreadsheet.getRange('A1:C8'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'false')
      .setOption('title', 'This week and last week (page views)')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setOption('hAxis.textStyle.color', '#000000')
      .setOption('vAxes.0.textStyle.color', '#000000')
      .setOption('height', 246)
      .setOption('width', 398)
      .setPosition(1, 4, 1, 0)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asAreaChart()
      .addRange(spreadsheet.getRange('A1:C8'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'false')
      .setOption('title', 'This week and last week (page views)')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setOption('hAxis.textStyle.color', '#000000')
      .setOption('vAxes.0.textStyle.color', '#000000')
      .setOption('height', 246)
      .setOption('width', 398)
      .setPosition(1, 4, 1, 0)
      .build();
  sheet.insertChart(chart);
  spreadsheet.getRange('B13').activate();
  spreadsheet.getCurrentCell().setValue('This year');
  spreadsheet.getRange('C13').activate();
  spreadsheet.getCurrentCell().setValue('Last year');
  spreadsheet.getRange('A14').activate();
  spreadsheet.getCurrentCell().setValue('Jan');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A14:A25'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('B14').activate();
  spreadsheet.getCurrentCell().setFormula('=\'This year\'!B16');
  spreadsheet.getRange('C14').activate();
  spreadsheet.getCurrentCell().setFormula('=\'Last year\'!B16');
  spreadsheet.getRange('B14').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('B14:B25'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('C14').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C14:C25'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('A13:C25').activate();
  chart = sheet.newChart()
      .asBarChart()
      .addRange(spreadsheet.getRange('A13:C25'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'absolute')
      .setOption('title', 'This year and Last year')
      .setYAxisTitle('')
      .setPosition(6, 2, 79, 7)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asBarChart()
      .addRange(spreadsheet.getRange('A13:C25'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'absolute')
      .setOption('title', 'This year and Last year')
      .setYAxisTitle('')
      .setPosition(13, 4, 1, 0)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asBarChart()
      .addRange(spreadsheet.getRange('A13:C25'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'absolute')
      .setOption('title', 'This year and Last year')
      .setYAxisTitle('')
      .setOption('height', 246)
      .setOption('width', 398)
      .setPosition(13, 4, 1, 0)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asBarChart()
      .addRange(spreadsheet.getRange('A13:C25'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'absolute')
      .setOption('title', 'This year and last year (page views)')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setOption('hAxis.textStyle.color', '#000000')
      .setYAxisTitle('')
      .setOption('vAxes.0.textStyle.color', '#000000')
      .setOption('height', 246)
      .setOption('width', 398)
      .setPosition(13, 4, 1, 0)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asBarChart()
      .addRange(spreadsheet.getRange('A13:C25'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'absolute')
      .setOption('title', 'This year and last year (page views)')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setOption('hAxis.textStyle.color', '#000000')
      .setYAxisTitle('')
      .setOption('vAxes.0.textStyle.color', '#000000')
      .setOption('height', 246)
      .setOption('width', 398)
      .setPosition(13, 4, 1, 0)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asBarChart()
      .addRange(spreadsheet.getRange('A13:C25'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'absolute')
      .setOption('title', 'This year and last year (page views)')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setOption('hAxis.textStyle.color', '#000000')
      .setYAxisTitle('')
      .setOption('vAxes.0.textStyle.color', '#000000')
      .setOption('height', 246)
      .setOption('width', 398)
      .setPosition(13, 4, 1, 0)
      .build();
  sheet.insertChart(chart);
  spreadsheet.getRange('I1').activate();
  spreadsheet.getCurrentCell().setValue('Browsers');
  spreadsheet.getRange('J1').activate();
  spreadsheet.getCurrentCell().setValue('Page views');
  spreadsheet.getRange('I1').activate();
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('I2').activate();
  spreadsheet.getCurrentCell().setFormula('=\'Top browsers\'!A16');
  spreadsheet.getRange('J2').activate();
  spreadsheet.getCurrentCell().setFormula('=\'Top browsers\'!B16');
  spreadsheet.getRange('I2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('I2:I25'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('J2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('J2:J25'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('I1:J25').activate();
  chart = sheet.newChart()
      .asPieChart()
      .addRange(spreadsheet.getRange('I1:J3'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'false')
      .setOption('title', 'Page views')
      .setPosition(17, 2, 79, 7)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asPieChart()
      .addRange(spreadsheet.getRange('I1:J3'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'false')
      .setOption('title', 'Page views')
      .setPosition(14, 9, 54, 8)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asPieChart()
      .addRange(spreadsheet.getRange('I1:J3'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'false')
      .setOption('title', 'Page views')
      .setPosition(8, 11, 4, 4)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asPieChart()
      .addRange(spreadsheet.getRange('I1:J3'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .set3D()
      .setOption('isStacked', 'false')
      .setOption('title', 'Page views')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('pieSliceTextStyle.color', '#000000')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setPosition(8, 11, 4, 4)
      .build();
  sheet.insertChart(chart);
  spreadsheet.getRange('B28').activate();
  spreadsheet.getCurrentCell().setValue('Countries');
  spreadsheet.getRange('C28').activate();
  spreadsheet.getCurrentCell().setValue('Page views');
  spreadsheet.getRange('B28:C28').activate();
  spreadsheet.getActiveRangeList().setFontWeight('bold');
  spreadsheet.getRange('B29').activate();
  spreadsheet.getCurrentCell().setFormula('=\'Top countries\'!A16');
  spreadsheet.getRange('C29').activate();
  spreadsheet.getCurrentCell().setFormula('=\'Top countries\'!B16');
  spreadsheet.getRange('B29').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('B29:B900'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('C29').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C29:C900'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('B28:C900').activate();
  chart = sheet.newChart()
      .asPieChart()
      .addRange(spreadsheet.getRange('B28:C60'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('useFirstColumnAsDomain', true)
      .setOption('isStacked', 'false')
      .setOption('title', 'Page views')
      .setPosition(23, 2, 79, 7)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asPieChart()
      .addRange(spreadsheet.getRange('B28:C60'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('pieHole', 0.5)
      .setOption('isStacked', 'false')
      .setOption('title', 'Page views')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('pieSliceTextStyle.color', '#000000')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setPosition(23, 2, 79, 7)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asPieChart()
      .addRange(spreadsheet.getRange('B28:C60'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('pieHole', 0.5)
      .setOption('isStacked', 'false')
      .setOption('title', 'Page views')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('pieSliceTextStyle.color', '#000000')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setPosition(27, 4, 9, 19)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asPieChart()
      .addRange(spreadsheet.getRange('B28:C60'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('pieHole', 0.5)
      .setOption('isStacked', 'false')
      .setOption('title', 'Page views by country this year')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('pieSliceTextStyle.color', '#000000')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setPosition(27, 4, 9, 19)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asPieChart()
      .addRange(spreadsheet.getRange('B28:C60'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('legend.position', 'top')
      .setOption('pieHole', 0.5)
      .setOption('isStacked', 'false')
      .setOption('title', 'Page views by country this year')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('pieSliceTextStyle.color', '#000000')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setPosition(27, 4, 9, 19)
      .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
      .asPieChart()
      .addRange(spreadsheet.getRange('B28:C60'))
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setTransposeRowsAndColumns(false)
      .setNumHeaders(1)
      .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
      .setOption('bubble.stroke', '#000000')
      .setOption('useFirstColumnAsDomain', true)
      .setOption('legend.position', 'labeled')
      .setOption('pieHole', 0.5)
      .setOption('isStacked', 'false')
      .setOption('title', 'Page views by country this year')
      .setOption('annotations.domain.textStyle.color', '#808080')
      .setOption('textStyle.color', '#000000')
      .setOption('legend.textStyle.color', '#191919')
      .setOption('pieSliceTextStyle.color', '#000000')
      .setOption('titleTextStyle.color', '#757575')
      .setOption('annotations.total.textStyle.color', '#808080')
      .setPosition(27, 4, 9, 19)
      .build();
  sheet.insertChart(chart);
};
