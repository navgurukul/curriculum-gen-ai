/**
 * Runs when the spreadsheet is opened and creates a custom menu in the UI.
 */
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Custom Menu')
        .addItem('Check Resume Links', 'checkResumeLinksAndsendDeadlineEmails')
        .addItem('Generate Feedback', 'generateResumeFeedback')
        .addToUi();
}
