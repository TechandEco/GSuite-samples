// Declaring column header names as constants allows us to call
// them by their name instead of by their index.
// All these are included in the form except the first three.
const ACTION = 'Action';   // manually added to the sheet
const COMMENT = 'Comment'; // manually added to the sheet
const STATUS = 'Status';   // manually added to the sheet
const TIMESTAMP = 'Timestamp';
const NAME = 'Name';
const EMAIL = 'Email';
const DEPARTMENT = 'Department';
const OFFSITE_BUDGET = 'Offsite budget';
const TRAVEL_BUDGET = 'Travel budget';
const OFFICE_SUPPLIES_BUDGET = 'Office supplies budget';
const MARKETING_BUDGET = 'Marketing budget';
const RECRUITING_BUDGET = 'Recruiting budget';

const TEMPLATES_SHEET = 'Templates';

/**
 * Creates a form to populate the table data.
 * This function must be run from the script editor as
 * a one time setup.
 */
function createForm() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let form = FormApp.create('Submit budget here')
      .setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId())
      .setAllowResponseEdits(true);

  // To learn more about the data types of form items, read:
  // https://developers.google.com/apps-script/reference/forms/item-type
  form.addTextItem().setTitle(NAME);
  form.addTextItem().setTitle(EMAIL);

  let departments = form.addListItem().setTitle(DEPARTMENT);
  departments.setChoices([
    departments.createChoice('Marketing'),
    departments.createChoice('Sales'),
    departments.createChoice('HR'),
    departments.createChoice('Finance'),
    departments.createChoice('Engineering'),
  ]);

  form.addTextItem().setTitle(OFFSITE_BUDGET);
  form.addTextItem().setTitle(TRAVEL_BUDGET);
  form.addTextItem().setTitle(OFFICE_SUPPLIES_BUDGET);
  form.addTextItem().setTitle(MARKETING_BUDGET);
  form.addTextItem().setTitle(RECRUITING_BUDGET);
}

/**
 * Installs a trigger in the Spreadsheet to run upon the Sheet being opened.
 * To learn more about triggers read:
 * https://developers.google.com/apps-script/guides/triggers
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Send email')
      .addItem('Approved', 'processApproved')
      .addItem('Not approved', 'processNotApproved')
      .addItem('Research needed', 'processResearchNeeded')
      .addToUi();
}

/**
 * Wrapper function of `processRows` for the 'Approved' action.
 */
function processApproved() {
  processRows('Approved');
}

/**
 * Wrapper function of `processRows` for the 'Approved' action.
 */
function processNotApproved() {
  processRows('Not approved');
}

/**
 * Wrapper function of `processRows` for the 'Approved' action.
 */
function processResearchNeeded() {
  processRows('Research needed');
}

/**
 * Processes only the rows matching the action.
 * It sends an email if the `STATUS` column is empty.
 * This updates the `STATUS` column in the sheet.
 *
 * @param {string} action - drop down option in Action column.
 * @param {string} emailTemplate - this allows for template's name
 *        to not have to be the exact Action's name.
 *        Defaults to the Action name.
 */
function processRows(action, emailTemplate=null) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get the email template doc URLs into a {key: value} Map
  // in the format {templateName: templateURL}.
  let templateRows = ss.getSheetByName(TEMPLATES_SHEET).getDataRange().getValues();
  let templates = templateRows
      .reduce((result, row) => result.set(row[0], row[1]), new Map());

  // Load the row data and get its headers.
  let dataRange = ss.getActiveSheet().getDataRange();
  let rows = dataRange.getValues();
  let headers = rows.shift();

  // Get the values from the status column.
  // These are the values that we want to write back to the sheet.
  let statusRange = dataRange.offset(1, headers.indexOf(STATUS), rows.length, 1);
  let statusValues = statusRange.getValues();

  // Process each row, send an email if necessary and update the `statusValues`.
  rows
      // Convert the row arrays into objects.
      // Start with an empty object, then create a new field
      // for each header name using the corresponding row value.
      .map(rowArray => headers.reduce((rowObject, fieldName, i) => {
        rowObject[fieldName] = rowArray[i];
        return rowObject;
      }, {}))

      // Add the row index (0-based) to the row object, this is used to update
      // the status of the rows that were modified.
      // We do this because the indices won't match after the next `filter` operation.
      // We use the spread operator to unpack the `row` object.
      // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Operators/Spread_syntax
      .map((row, i) => ({...row, rowIndex: i}))

      // From all the rows, filter out and only keep the ones that match the
      // action and the status is empty.
      .filter(row => row[ACTION] == action && !row[STATUS])

      // Send an email and update the status in `statusValues`.
      // We don't need a return value so we use `forEach` instead of `map`.
      .forEach(row => {
        // We start with the doc template HTML body, and then we replace
        // each '{{fieldName}}' with the row's respective value.
        let emailBody = headers.reduce(
          (result, fieldName) => result.replace(`{{${fieldName.toUpperCase()}}}`, row[fieldName]),
          docToHtml(templates.get(emailTemplate || action))
        );

        // Try to send an email, or get the error if it fails.
        let status;
        try {
          MailApp.sendEmail({
            to: row[EMAIL],
            subject: `Request for budget: ${row[ACTION]}`,
            htmlBody: emailBody,
          });
          status = `${row[ACTION]}: ${new Date}`;
        } catch (e) {
          status = `Error: ${e}`;
        }

        // Update the `statusValues` with the new status or error.
        // We use the `rowIndex` from before to update the correct
        // row in `statusValues`.
        statusValues[row.rowIndex][0] = status;
        Logger.log(`Row ${row.rowIndex+2}: ${status}`);
      });

  // Write statusValues back into the sheet "status" column.
  statusRange.setValues(statusValues);
}

/**
 * Fetches a Google Doc as an HTML string.
 *
 * @param {string} docUrl - The URL of a Google Doc to fetch content from.
 * @return {string} The Google Doc rendered as an HTML string.
 */
function docToHtml(docUrl) {
  let docId = DocumentApp.openByUrl(docUrl).getId();
  return UrlFetchApp.fetch(
    `https://docs.google.com/feeds/download/documents/export/Export?id=${docId}&exportFormat=html`,
    {
      method: 'GET',
      headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true,
    },
  ).getContentText();
}
