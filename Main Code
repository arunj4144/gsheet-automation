// Function to extract username from email address
function getEditedByUsername(e) {
  if (e && e.authMode === 'FULL' && e.user && e.user.email) {
    var username = e.user.email.split('@')[0];
    return username.charAt(0).toUpperCase() + username.slice(1).toLowerCase().replace(/\./g, ' ');
  } else if (e && e.user && e.user.nickname) {
    // Use the nickname as a fallback if the email is not available
    return e.user.nickname.charAt(0).toUpperCase() + e.user.nickname.slice(1).toLowerCase().replace(/\./g, ' ');
  } else {
    return 'Unknown User';
  }
}

// Main onEdit function
function onEdit(e) {
  try {
    Logger.log('Event Object: %s', JSON.stringify(e));

    // Check if the event object, range, value, and source are defined
    if (e && e.range && e.value && e.source) {
      // Check if the edited cell is in column L
      if (e.range.getColumn() == 12) { // Changed to column L
        Logger.log('Column L edited');

        // Get the active sheet
        var sheet = e.range.getSheet();

        // Get the data from the edited row
        var rowData = sheet.getRange(e.range.getRow(), 2, 1, 9).getValues()[0];
        var userName = rowData[0];
        var mailID = rowData[1];
        var password = rowData[2];
        var assets = rowData[3];
        var entity = getEmailEntity(mailID);
        var triggerDate = Utilities.formatDate(new Date(), "GMT+0", "dd-MM-yyyy");
        var currentMonthYear = Utilities.formatDate(new Date(), "GMT+0", "MMMM yyyy");

        // Get the address for asset allocation from column L
        var assetAddress = sheet.getRange(e.range.getRow(), 12).getValue();

        // Additional email addresses
        var additionalEmails = ['required mail IDs'];

        // Compose the HTML email body
        var body = `
          <p>Hi Team,</p>
          <p>We have created a mail for <strong>${userName}</strong> and please find the details below:</p>
          <ul>
            <li><strong>Name:</strong> ${userName}</li>
            <li><strong>Mail ID:</strong> ${mailID}</li>
            <li><strong>Password:</strong> ${password}</li>
            <li><strong>Asset:</strong> ${assets}</li>
            <li><strong>Entity:</strong> ${entity}</li>
            <li><strong>Address for Asset Allocation:</strong></li>
            <p style="margin-left: 20px;"><strong>${assetAddress}</strong></p>
          </ul>
          <p>Trigger Date: ${triggerDate}</p>
        `;

        // Send an HTML email to multiple recipients
        MailApp.sendEmail({
          to: additionalEmails.join(','),
          subject: `User Onboarded on DB: ${triggerDate}`,
          htmlBody: body,
        });

        // Update the "Entity" in column G, "Created" in column J
        sheet.getRange(e.range.getRow(), 7).setValue(entity);  // Column G: Entity
        sheet.getRange(e.range.getRow(), 10).setValue('Created');  // Column J: Created
        // Update the additional information in columns H and I
        sheet.getRange(e.range.getRow(), 8).setValue(triggerDate);  // Column H: Trigger Date
        sheet.getRange(e.range.getRow(), 9).setValue(currentMonthYear);  // Column I: Month Year
      }
    } else {
      Logger.log('Event object is missing necessary properties.');
    }
  } catch (error) {
    Logger.log('Error: %s', error.toString());
  }
}

// Function to extract "entity" from email address (excluding .com)
function getEmailEntity(email) {
  var matches = email.match(/@([^\s@.]+)\./);
  return matches ? matches[1] : null;
}
