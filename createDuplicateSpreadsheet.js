function createDuplicateSpreadsheet() {
  let originalSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let originalName = originalSpreadsheet.getName();
  let newSpreadsheetName = originalName + " - Manager's Copy";
  
  // Set the original owner's email address
  let originalOwnerEmail = 'william@williampasternak.com'; // Replace with the original owner's email address
  
  // Create a copy of the current spreadsheet, allow anyone with link to edit spreadsheet.
  let newSpreadsheet = DriveApp.getFileById(originalSpreadsheet.getId()).makeCopy(newSpreadsheetName);
      newSpreadsheet.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT)
  
  // Pull email address from spreadsheet and send an email to that address 
  let recipientEmail = originalSpreadsheet.getRange('J2').getValue(); // Change cell J2 to any single cell where your desired email is
  
  // Send the requested user a link to the new spreadsheet
  let newSpreadsheetUrl = newSpreadsheet.getUrl();
  GmailApp.sendEmail(recipientEmail, 'Requested Spreadsheet', 'Here is the link to an editable version of the spreadsheet: ' + newSpreadsheetUrl);
  
  // Open the new spreadsheet and protect existing protected ranges
  let newSpreadsheetApp = SpreadsheetApp.openById(newSpreadsheet.getId());
  let sheets = newSpreadsheetApp.getSheets();
  sheets.forEach(function(sheet) {
    let protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

    protections.forEach(function(protection) {
      if (protection.canEdit()) {
        protection.removeEditors(protection.getEditors());
        protection.addEditor(originalOwnerEmail);
        protection.setWarningOnly(false);
      }
    });
  });
}
