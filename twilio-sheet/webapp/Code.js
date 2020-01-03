var SHEET_URL = 'ENTER_SPREADSHEET_URL_HERE';

/**
 * Loads main web app page.
 * Learm more about web apps: https://developers.google.com/apps-script/guides/web
 *
 * @param {Object} e - HTTP GET event.
 */
function doGet(e) {
  Logger.log(e.parameter);
  return HtmlService.createHtmlOutputFromFile('Page');
}

/**
 * Writes phone number into spreadsheet .
 *
 * @param {number} phoneNumber - number submitted from web app.
 */
function userClicked(phoneNumber){
  var spreadsheet = SpreadsheetApp.openByUrl(SHEET_URL);
  var sheet = spreadsheet.getSheets()[0];
  
  sheet.appendRow([new Date(), phoneNumber]);
}