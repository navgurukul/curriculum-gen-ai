function checkResumeLinksAndsendDeadlineEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Student");
  var adminSheet = ss.getSheetByName("Admins");
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  var adminEmails = getAdminEmails(adminSheet);

  for (var i = 1; i < data.length; i++) {
    var studentEmail = data[i][0];
    var resumeLink = data[i][1];
    var deadline = new Date(data[i][4]);
    var today = new Date();

    if (today > deadline && (resumeLink.trim() === "" || !isLinkAccessible(resumeLink))) {
      checkResumeLinksAndsendDeadlineEmails(studentEmail, adminEmails, deadline);
    }
  }
}

function isLinkAccessible(url) {
  try {
    var fileId = extractFileIdFromUrl(url);
    var file = DriveApp.getFileById(fileId);
    var viewers = file.getViewers();
    return viewers.length > 0;
  } catch (e) {
    Logger.log('Error accessing file: ' + e.message);
    return false;
  }
}

function extractFileIdFromUrl(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function checkResumeLinksAndsendDeadlineEmails(studentEmail, adminEmails, deadline) {
  var subject = "Resume Submission Deadline Missed";
  var body = `Dear Student,

You have missed the deadline to upload your resume. The deadline was ${deadline.toDateString()}.

Please upload your resume as soon as possible.

Best regards,
University Admin`;
  var cc = adminEmails.join(",");

  MailApp.sendDeadlineEmails({
    to: studentEmail,
    subject: subject,
    body: body,
    cc: cc
  });
}

function getAdminEmails(adminSheet) {
  var adminEmails = [];
  var adminData = adminSheet.getDataRange().getValues();

  for (var i = 1; i < adminData.length; i++) {
    adminEmails.push(adminData[i][0]);
  }
  return adminEmails;
}

function onEdit(e) {
  if (!e) {
    return;
  }

  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var row = range.getRow();

  if (sheet.getName() === "Student" && row > 1) {
    var resumeLink = sheet.getRange(row, 2).getValue();
    var deadline = new Date(sheet.getRange(row, 4).getValue());
    var today = new Date();
    var studentEmail = sheet.getRange(row, 1).getValue().trim();
    var adminSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Admins");
    var adminEmails = getAdminEmails(adminSheet);

    if (today > deadline && (!isLinkAccessible(resumeLink) || resumeLink.trim() === "")) {
      checkResumeLinksAndsendDeadlineEmails(studentEmail, adminEmails, deadline);
    }
  }
}

function createTimeDrivenTrigger() {
  ScriptApp.newTrigger('checkResumeLinksAndsendDeadlineEmails')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .create();
}

// function onOpen() {
//   var ui = SpreadsheetApp.getUi();
//   ui.createMenu('Custom Menu')
//     .addItem('Check Resume Links', 'checkResumeLinksAndsendDeadlineEmails')
//     .addToUi();
// }

// Run this function manually once to set up the time-driven trigger
function setupTriggers() {
  createTimeDrivenTrigger();
}
