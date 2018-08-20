function onFormSubmit(e) {
  Logger.log("[METHOD] onFormSubmit");
  
  sendEmail(e.range);
}

function sendEmail(range) {
  Logger.log("[METHOD] sendEmail");
  
  // FETCH SPREADSHEET //
  var values = range.getValues();
  var row = values[0];
  
  // EXTRACT VALUES //
  var requestCategory = row[4];
  var requestName = row[2];
  var requestRole = row[3];
  var requestEmail = row[1];
  var requestDate = row[0];
  var requestSubject = row[5];
  var requestDescription = row [6];
  var requestPrioritization = row[7];
  var requestAdditionalNotes = row[8];
  var requestAltEmail = row[9];
  var requestScreenshots = row[10];
  var requestEmailSent = row[11];
  
  // CLEAN VARIABLES //
  requestCategory = shortenCategoryText(requestCategory);
  
  if (requestAdditionalNotes.length == 0) 
    requestAdditionalNotes = "None";
  
  var requestFollowUp = "";
  if (requestScreenshots.length != 0) 
    requestFollowUp = requestFollowUp.concat("<li>Ask for screenshots or additional files</li>");
  if 
    (requestAltEmail.length != 0) 
    requestFollowUp = requestFollowUp.concat("<li>Provide status updates to "+requestAltEmail+"</li>");
  if 
    (requestFollowUp.length == 0) 
    requestFollowUp = "<br />None";
  else requestFollowUp = "<ul>"+requestFollowUp+"</ul>";
  
  // PREPARE EMAIL //
  var emailRecipients = emailCategorySelector(requestCategory);
  var emailSubject = requestSubject+": "+requestName+" ("+requestEmail+")";
  var emailBody = "<h3>A request has been reported by "+requestEmail+":</h3><hr /> \
      <p> \
      <h5>This is an automated summary of a recent IT Support Request Ticket. This message has been sent to your email as it was selected under your category of IT specialization. If you believe that this is an error, please contact Josh on Slack or at josh.chan@cus.ca</h5><hr/> \
      </p><p> \
      <h1 style='line-height:90%'>"+requestSubject+": "+requestName+"<br /> \
      <span style='font-size:60%'>("+requestDate+")</span></h1> \
      </p><hr /> \
      <p> \
      <strong>ROLE:</strong><br /> \
      "+requestRole+" \
      </p><p> \
      <strong>REQUEST TYPE:</strong><br /> \
      "+requestCategory+" \
      </p><p> \
      <strong>DESCRIPTION:</strong><br /> \
      "+requestDescription+" \
      </p><p> \
      <strong>PRIORITIZATION [1 - 5]:</strong><br /> \
      "+requestPrioritization+" \
      </p><p> \
      <strong>ADDITIONAL NOTES:</strong><br /> \
      "+requestAdditionalNotes+" \
      </p><p> \
      <strong>FOLLOW UP:</strong> \
      "+requestFollowUp+" \
      </p>";
  
  // SEND EMAIL //
  var confirmSheet = SpreadsheetApp.getActiveSheet();
  var confirmRange = confirmSheet.getRange(confirmSheet.getLastRow(),12,1,1);
  var EMAIL_SENT = "Email Sent!"
  if (requestEmailSent != EMAIL_SENT) { // Prevents sending duplicates
    MailApp.sendEmail({
      to: emailRecipients,
      subject: emailSubject,
      htmlBody: emailBody
  });     
    confirmRange.setValue(EMAIL_SENT);
    confirmRange.setBackgroundRGB(0,255,0);
    SpreadsheetApp.flush();// Make sure the cell is updated right away in case the script is interrupted
  }
}

// HELPER FUNCTION FOR CLEAN VARIABLES //
function shortenCategoryText(text) {
  Logger.log("[METHOD] shortenCategoryText");
  
  switch(text) {
    case "G Suite (Email issues, Requesting email account etc.)":
        return "GSuite";
    case "WordPress (Website access, WordPress issues, Website errors/issues)":
        return "WordPress";
    case "Domain Registration/Website Hosting":
        return "Website/Server";
    case "IT Consulting & Support (General help)":
        return "General IT Support";
    default:
        return "IT Support";
  }
}


// HELPER FUNCTION TO SELECT CORRECT EMAIL //
function emailCategorySelector(text) {
  Logger.log("[METHOD] emailCategorySelector");
  
  if (text == "GSuite")
    return "josh.chan@cus.ca";
  if (text == "WordPress")
    return "rachel.zhu@cus.ca";
  if (text == "Website/Server")
    return "alvin.lo@cus.ca";
  if (text =="General IT Support"|"IT Support")
    return "ansel.hartanto@cus.ca";
}


// TEST FUNCTION //

function test() {
  Logger.log("[METHOD] test");
  
  var ss = SpreadsheetApp.openById("13nHs9-ShOFIosHZLNtZYlW6Avj_qHfAItZ42z218ljg").getActiveSheet();
  var testRange = ss.getRange(8,1,1,ss.getLastColumn());
  sendEmail(testRange);
}
