/**
 * Installs a trigger in the Spreadsheet to run upon the sheet being opened.
 * To learn more about triggers read:
 * https://developers.google.com/apps-script/guides/triggers
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Clean my export')
        .addItem('Find misspelled states', 'misspelledStatesPivot')
        .addItem('Lookup office codes', 'lookupOfficeCodes')
        .addItem('Query 2019 records', 'queryRecordsByYear')
        .addItem('Find & Query', 'macroCombined')
        .addToUi();
  };

  /**
   *  Combines two functions.
   */
  function macroCombined() {
    lookupOfficeCodes();
    queryRecordsByYear();
  }

  /**
   * Uses a macro to record steps in a Google Sheet to make a script.
   * This script is called as a function from a custom menu to
   * automatically create a pivot table that shows a count of
   * each item in column A.
   * To learn more about macro recording:
   * https://developers.google.com/apps-script/guides/sheets/macros
   */
  function misspelledStatesPivot() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('A:A').activate();
    var sourceData = spreadsheet.getRange('A:A');
    spreadsheet.insertSheet(spreadsheet.getActiveSheet().getIndex() + 1).activate();
    spreadsheet.getActiveSheet().setHiddenGridlines(true);
    var pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
    pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
    var pivotGroup = pivotTable.addRowGroup(1);
    pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
    var pivotValue = pivotTable.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
    pivotGroup = pivotTable.addRowGroup(1);
    spreadsheet.getActiveSheet().setName('Misspelled states');
  };


  /**
   * Uses a macro to record steps in a Google Sheet to make a script.
   * This script is called as a function from a custom menu to
   * automatically perform a VLOOKUP of office codes in another tab
   * of the sheet called 'Office Codes' that's associated to a US
   * state listed in column A.
   * To learn more about macro recording:
   * https://developers.google.com/apps-script/guides/sheets/macros
   */
  function lookupOfficeCodes() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('E2').activate();
    spreadsheet.getCurrentCell().setFormula('=VLOOKUP(A2,\'OFFICE CODES\'!$A$1:$B$10,2, FALSE)');
    spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('E2:E22'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    spreadsheet.getRange('E2:E22').activate();
  };

  /**
   * Uses a macro to record steps in a Google Sheet to make a script.
   * This script is called as a function from a custom menu to
   * automatically displays a table in the ' Exported Data' sheet
   * in cell G1 with all the rows that contain the date '201901--'
   * (any day within the month of January in the year 2019) in
   * column D.
   * To learn more about macro recording:
   * https://developers.google.com/apps-script/guides/sheets/macros
   */
  function queryRecordsByYear() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('G1').activate();
    spreadsheet.getCurrentCell().setValue('Query January 2019 records');
    spreadsheet.getRange('G2').activate();
    spreadsheet.getCurrentCell().setFormula('=QUERY(A:E,"select B,C,D WHERE D LIKE \'201901__\'", 1)');
    spreadsheet.getRange('G3').activate();
  };
