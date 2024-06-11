// function onOpen() {
//   var ui = SpreadsheetApp.getUi();
//   ui.createMenu('Resume Feedback')
//     .addItem('Generate Feedback', 'generateResumeFeedback')
//     .addToUi();
// }

function generateResumeFeedback() {
  // Get the active spreadsheet and the first sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Get the email range in column 1
  var emailRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  var emails = emailRange.getValues();

  // Get the link range in column 2
  var linkRange = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1);
  var pdfLinks = linkRange.getValues();

  // Get the role range in column 3
  var roleRange = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1);
  var roles = roleRange.getValues();

  // Loop through each link in column 2
  for (var i = 0; i < pdfLinks.length; i++) {
    var email = emails[i][0];
    var pdfLink = pdfLinks[i][0];
    var role = roles[i][0];
    var resumeFeedback = sheet.getRange(i + 2, 4).getValue(); // Get existing feedback from the same row in column 4

    // Generate feedback only if there is no existing feedback
    if (!resumeFeedback || resumeFeedback.trim() === '') {
      // Log the link being processed
      Logger.log('PDF link: ' + pdfLink);

      var pdfFileId = extractFileIdFromUrl(pdfLink);
      var pdfText = extractTextFromPdf(pdfFileId);

      if (pdfText) {
        var feedback = generateFeedback(pdfText, role);

        // Write the feedback in the adjacent column (column 4)
        sheet.getRange(i + 2, 4).setValue(feedback);

        // Send an email with the feedback
        sendEmail(email, feedback, role);
      } else {
        sheet.getRange(i + 2, 4).setValue('No feedback generated');
      }
    }
  }
}

function extractFileIdFromUrl(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

function extractTextFromPdf(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);

    // Create a new Google Doc from the PDF file
    var resource = {
      title: file.getName(),
      mimeType: MimeType.GOOGLE_DOCS
    };

    var docFile = Drive.Files.copy(resource, fileId);
    var doc = DocumentApp.openById(docFile.id);
    var text = doc.getBody().getText();

    // Remove the temporary Google Doc
    DriveApp.getFileById(docFile.id).setTrashed(true);

    Logger.log('Extracted text: ' + text);
    return text;
  } catch (e) {
    Logger.log('Error extracting text from PDF: ' + e.message);
    return null;
  }
}

function generateFeedback(text, role) {
  var apiKey = 'AIzaSyCKJQlgx5c-mQZkxT5OkkJXBYohHMKuuAY'; // Replace with your actual API key
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=' + apiKey;
  
  // Define your prompt to encourage structured feedback
  var prompt = `Generate feedback for the resume text for the role of: ${role}. 
  Structure the feedback with subheadings, bullet points, and bold key strengths. Here is the text: ${text}`;

  var payload = {
    "contents": [
      {
        "parts": [
          {
            "text": prompt
          }
        ]
      }
    ]
  };

  // Log the API request payload
  Logger.log('API Request Payload: ' + JSON.stringify(payload));

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    var contentText = response.getContentText();
    Logger.log('API Response Status Code: ' + statusCode);
    Logger.log('API Response Content: ' + contentText);

    var jsonResponse = JSON.parse(contentText);

    // Correct response parsing
    if (jsonResponse && jsonResponse.candidates && jsonResponse.candidates.length > 0 && jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts.length > 0) {
      var generatedContent = jsonResponse.candidates[0].content.parts[0].text;
      Logger.log('Generated Content: ' + generatedContent);

      // Convert generated content to formatted text
      var formattedContent = formatFeedback(generatedContent);
      Logger.log('Formatted Content: ' + formattedContent);
      return formattedContent;
    } else {
      Logger.log('No content found in API response.');
      return 'No feedback found';
    }
  } catch (e) {
    // Log any errors
    Logger.log('Error fetching data from Generative Language API: ' + e.message);
    return 'Error generating feedback.';
  }
}

function formatFeedback(content) {
  // Split the content into lines
  var lines = content.split('\n');
  var formattedContent = '';
  
  lines.forEach(line => {
    line = line.trim();

    if (line === '') {
      return;
    }

    // Remove symbols that decrease readability
    line = line.replace(/[*#~]/g, '');

    // Check for subheadings, bullet points, and bold text
    if (line.match(/^Subheading:/i)) {
      formattedContent += `\n\n<strong>${line.replace(/Subheading:/i, '').trim()}</strong>\n\n`;
    } else if (line.match(/^Bullet:/i)) {
      formattedContent += `\n• ${line.replace(/Bullet:/i, '').trim()}`;
    } else if (line.match(/^Bold:/i)) {
      formattedContent += `\n\n<strong>${line.replace(/Bold:/i, '').trim()}</strong>\n\n`;
    } else {
      formattedContent += `\n${line}`;
    }
  });

  return formattedContent;
}

function sendEmail(email, feedback, role) {
  var subject = 'Resume Feedback for the role of ' + role;
  var body = 'Dear Applicant,\n\nHere is the feedback for your resume for the role of ' + role + ':\n\n' + feedback + '\n\nBest Regards,\nYour Team';

  MailApp.sendEmail(email, subject, body);
}
