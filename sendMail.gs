
function getAttachment(var fileId){
  return DriveApp.getFileById(fileId);
}

function getSheet(var spreadSheetId, var sheetName){
  // Get spread sheet
  var ss = SpreadsheetApp.openById(spreadSheetId);
  Logger.log(ss);
  // Get specific sheet form selected spread sheet.
  return ss.getSheetByName(sheetName);
}
var EMAIL_SENT="email sent";

function sendEmails() {
  var spreadSheetId = "Enter Sheet Id";
  var sheetName = "Sheet within SpreadSheet name";
  var sheet = getSheet(spreadSheetId, sheetName);
  Logger.log(sheet);
  //information regarding the spreadsheet
  var startRow = 2;  // First row of data to process
  var numRows = 20;   // Number of rows to process
  var startColumn = 1; // Starting column of spreadsheet
  var numColumns = 2; // No. of columns after start column

  // Get data from sheet based on the parameters provided
  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns)
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var attachmentId = "AttachmentId";
  var attachment = getAttachment(attachmentId);
  for (i in data) {
    var row = data[i];
    var emailAddress = row[0];
    Logger.log("row:")
    Logger.log(row)
    var isEmailSent = row[1];
    if(isEmailSent != EMAIL_SENT){
      var t = parseInt(startRow) + parseInt(i);
      Logger.log(t);
      Logger.log("isEmailSent");
      Logger.log(isEmailSent);
      Logger.log("email:");
      Logger.log(emailAddress);      
      var htmlBody = "<html> <body style='color:black'> <p>Hi, Write your mail body here<br><br> Regards,</body></html>"
      GmailApp.sendEmail(emailAddress, subject,"",{
                   name: "Name of the person",
                   htmlBody: htmlBody,
        attachments: [attachment.getAs(MimeType.PNG)]
                 });
      
      sheet.getRange(t, 2).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
