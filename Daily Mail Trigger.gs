// Function to send an email with the onboarding details
function sendEmailWithDetailsOB() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Onboarding");
    Logger.log("Sheet: " + sheet.getName());

    if (!sheet) {
      throw new Error("Onboarding sheet not found.");
    }

    var today = new Date();
    var todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yyyy");
    Logger.log("Today: " + todayFormatted);

    var data = sheet.getDataRange().getValues();

    var editedData = data.filter(function(row) {
      var editedDate = Utilities.formatDate(new Date(row[7]), Session.getScriptTimeZone(), "MM/dd/yyyy");
      return editedDate === todayFormatted;
    });

    var message = "Hi Team, Please find the onboarding status of these employees for " + todayFormatted + ":\n\n";
    if (editedData.length > 0) {
      message += "<table style='border-collapse: collapse; width: 100%;'>";
      message += "<tr style='background-color: #f2f2f2;'>";
      message += "<th style='border: 1px solid #ddd; padding: 8px;'>User Name</th>";
      message += "<th style='border: 1px solid #ddd; padding: 8px;'>Email ID</th>";
      message += "<th style='border: 1px solid #ddd; padding: 8px;'>Password</th>";
      message += "<th style='border: 1px solid #ddd; padding: 8px;'>Asset if Available</th>";
      message += "<th style='border: 1px solid #ddd; padding: 8px;'>Status</th>";
      message += "</tr>";

      for (var i = 0; i < editedData.length; i++) {
        message += "<tr>";
        var columns = [1, 2, 3, 4, 6]; // Indices for columns B, C, D, E, G
        columns.forEach(function(colIndex) {
          message += "<td style='border: 3px solid #ddd; padding: 8px;'>" + editedData[i][colIndex] + "</td>";
        });
        message += "</tr>";
      }
      message += "</table>";
    } else {
      message += "  ONBOARDING DATA NOT AVAILABLE FOR TODAY.";
    }

    MailApp.sendEmail({
      to: "required mail IDs",
      subject: "[>>] Onboarding Status " + todayFormatted,
      htmlBody: message
    });
    Logger.log("Email sent successfully.");
  } catch (error) {
    Logger.log("Error: " + error.message);
  }
}

// Function to add today's date in column A when a change is made in column B
function onEditAddDateOB(e) {
  var sheet = e.source.getSheetByName("Offboarding");
  if (!sheet) return;

  var range = e.range;
  var column = range.getColumn();
  var row = range.getRow();

  if (column === 2 && row > 1) { // Column B is 2, and assuming there's a header row
    var today = new Date();
    var todayFormatted = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yyyy");
    sheet.getRange(row, 1).setValue(todayFormatted); // Column A is 1
  }
}

// Function to create a time-based trigger for sending the email
function createTriggerOB() {
  ScriptApp.newTrigger("sendEmailWithDetailsOB")
    .timeBased()
    .everyDays(1)
    .atHour(20) // 8 PM
    .create();
}
