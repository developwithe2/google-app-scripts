const sheetName = 'SHEET_NAME'; // Input name of sheet tab from workbook
const scriptProps = PropertiesService.getScriptProperties();

// RUN initialSetup FUNCTION TO SET SCRIPT PROPERTY VALUE
function initialSetup () {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  scriptProps.setProperty('sheetId', spreadsheet.getId());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const spreadsheet = SpreadsheetApp.openById(scriptProp.getProperty('sheetId'));
    const sheet = spreadsheet.getSheetByName(sheetName);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nextRow = sheet.getLastRow() + 1;

    // CREATE SEPARATE VARIABLES FOR EACH MESSAGE VALUE
    const name = e.parameter['Name'];
    const email = e.parameter['Email'];
    const message = e.parameter['Message'];

    const newRow = headers.map(function(header) {
      // RETURNS DATE OF INPUT SUBMISSION
      return header === 'Date' ? new Date() : e.parameter[header];
    });
    sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);
    
    MailApp.sendEmail({
      to: 'INPUT_YOUR_EMAIL_ADDRESS', // Input email address to send message to
      replyTo: email, // Email input from HTML form
      subject: 'INPUT_EMAIL_SUBJECT', // Subject of email message
      body: `Name: ${name}\nEmail: ${email}\nMessage: ${message}`
    });

    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow }))
      .setMimeType(ContentService.MimeType.JSON)
  } catch (e) {
    Logger.log(`An error occurred: ${e}`);
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'error', 'error': e }))
      .setMimeType(ContentService.MimeType.JSON)
  }
  finally {
    lock.releaseLock()
  }
}
