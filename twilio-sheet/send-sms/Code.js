// Spreadsheet column names mapped to 0-based index numbers.
var TIME_ENTERED = 0;
var PHONE_NUMBER = 1;
var MESSAGE = 2;
var STATUS = 3;

// Enter your Twilio account information here.
var TWILIO_ACCOUNT_SID = 'PASTE_TWILIO_ACCOUNT_ID_HERE_EX: ACa0072e7863c5d13b703e1e18f7cab1ee';
var TWILIO_SMS_NUMBER = 'PASTE_TWILIO_SMS_NUMBER_HERE_EX: 14155925860';
var TWILIO_AUTH_TOKEN = 'PASTE_TWILIO_AUTH_TOKEN_HERE_EX: AC4z072e7863c5d13b703e1e18f7cab1ee:046077789202b8d7166de7a0d3383d9f';

/**
 * Installs a trigger in the Spreadsheet to run upon the Sheet being opened.
 * To learn more about triggers read: https://developers.google.com/apps-script/guides/triggers
 */
function onOpen() {
  // To learn about custom menus, please read:
  // https://developers.google.com/apps-script/guides/menus
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Send SMS')
      .addItem('Send to all', 'sendSmsToAll')
      .addToUi();
};  

/**
 * Sends text messages listed in the Google Sheet
 *
 */
function sendSmsToAll() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  
  // The `shift` method removes the first row and saves it into `headers`.
  var headers = rows.shift();
  
  // Try to send an SMS to every row and save their status.
  rows.forEach(function(row) {
    row[STATUS] = sendSms(row[PHONE_NUMBER], row[MESSAGE]);
  });
  
  // Write the entire data back into the sheet.
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}


/**
 * Sends a message to a given phone number via SMS through Twilio.
 * To learn more about sending an SMS via Twilio and Sheets: 
 * https://www.twilio.com/blog/2016/02/send-sms-from-a-google-spreadsheet.html
 *
 * @param {number} phoneNumber - phone number to send SMS to.
 * @param {string} message - text to send via SMS.
 * @return {string} status of SMS sent (successful sent date or error encountered).
 */
function sendSms(phoneNumber, message) {
  var twilioUrl = 'https://api.twilio.com/2010-04-01/Accounts/' + TWILIO_ACCOUNT_SID + '/Messages.json';

  try {
    UrlFetchApp.fetch(twilioUrl, {
      method: 'post',
      headers: {
        Authorization: 'Basic ' + Utilities.base64Encode(TWILIO_AUTH_TOKEN)
      },
      payload: {
        To: phoneNumber.toString(),
        Body: message,
        From: TWILIO_SMS_NUMBER,
      },
    });
    return 'sent: ' + new Date();
  } catch (err) {
    return 'error: ' + err;
  }
}
