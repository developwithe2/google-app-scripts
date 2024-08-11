const documentTemplateId = 'INPUT_TEMPLATE_ID'; // Input document template Id here
const documentName = 'INPUT_FILE_NAME'; // Input document name here

function onFormSubmit(e) {
  
  // Variables from form by spreadsheet column
  var timestamp = e.values[0];
  const entry = {
    email : e.values[1],
    choices : e.values[2]
  };

  // COPY DOCUMENT TEMPLATE & GET ID
  var newDocumentId = DriveApp.getFileById(documentTemplateId).makeCopy(documentName+'_Test').getId();
  // OPEN NEW DOCUMENT
  var newDocument = DocumentApp.openById(newDocumentId);
  // GET ACCESS TO DOCUMENT BODY
  var newDocumentBody = newDocument.getActiveSection();
  // REPLACE PLACEHOLDERS
  newDocumentBody.replaceText('{{placeholder_1}}', entry.email);
  if (entry.choices == null) {
    newDocumentBody.replaceText('{{placeholder_2}}', entry.choices);
  } else {
    newDocumentBody.replaceText('{{placeholder_2}}', 'No data.');
  }
  // SAVE & CLOSE NEW DOCUMENT
  newDocument.saveAndClose();
  
  // CONVERT NEW DOCUMENT TO PDF
  var pdf = DriveApp.getFileById(newDocumentId).getAs('application/pdf');
  
  // ATTACH PDF TO EMAIL
  const emailTemplate = HtmlService.createTemplateFromFile('emailTemplate');
  emailTemplate.components = entry;
  const message = emailTemplate.evaluate().getContent();
  MailApp.sendEmail({
    name: name,
    to: email,
    replyTo: 'INPUT_EMAIL_ADDRESS',
    subject: 'INPUT_EMAIL_SUBJECT',
    htmlBody: message,
    attachments: pdf
  });
  
  // MOVE NEW DOCUMENT TO TRASH
  DriveApp.getFileById(newDocumentId).setTrashed(true);
}
