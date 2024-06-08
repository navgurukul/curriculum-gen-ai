function onEdit(e) {
  if (!e) {
    return;
  }

  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var userEmail = Session.getActiveUser().getEmail();
  var row = range.getRow();
  var col = range.getColumn();

  Logger.log('Editing sheet: ' + sheet.getName());
  Logger.log('Editing row: ' + row);
  Logger.log('Editing column: ' + col);
  Logger.log('User email: ' + userEmail);

  if (sheet.getName() === "Student" && row > 1) {
    var studentEmail = sheet.getRange(row, 1).getValue().trim();
    var isAdmin = checkAdmin(userEmail);

    Logger.log('Student email: ' + studentEmail);
    Logger.log('Is admin: ' + isAdmin);

    if (userEmail !== studentEmail && !isAdmin) {
      e.range.setValue(e.oldValue); // Restore old value
      SpreadsheetApp.getUi().alert('You are not allowed to edit this row.');
    }
  }

  if (sheet.getName() === "Admins" && row > 1 && col === 1) {
    var isAdmin = checkAdmin(userEmail);

    Logger.log('Is admin: ' + isAdmin);

    if (!isAdmin) {
      e.range.setValue(e.oldValue); // Restore old value
      SpreadsheetApp.getUi().alert('Only admins can add or edit admin emails.');
    }
  }
}

function checkAdmin(email) {
  var adminSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Admins");
  if (adminSheet) {
    var adminEmails = adminSheet.getRange("A2:A" + adminSheet.getLastRow()).getValues().flat().filter(String);
    Logger.log('Admin emails: ' + adminEmails);
    return adminEmails.includes(email);
  } else {
    Logger.log('Admins sheet not found');
    // If the Admins sheet doesn't exist, assume no one is an admin
    return false;
  }
}
