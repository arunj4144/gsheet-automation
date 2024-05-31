# gsheet-automation
This project is used to send automatic emails with the data from a Google Sheet when a specific column in the corresponding sheet is updated
Detailed Expalnation


Email Username Extraction and Notification Script
This script is designed to extract a username from an email address and send an HTML email notification when a specific column in a Google Sheet is updated. It can be useful for user onboarding or other scenarios where you need to notify relevant parties about changes.

Prerequisites


Google Sheets: You’ll need a Google Sheet where the data is being updated.

Google Apps Script: This script is written in Google Apps Script, which allows you to automate tasks within Google Workspace.


Usage


Open your Google Sheet.
Click on Extensions > Apps Script.
Paste the provided code into the script editor.
Save the script.
Set up an onEdit trigger to execute the onEdit function whenever a cell is edited in the specified column (Column L in this case).


Functions

getEditedByUsername(e)
Extracts a username from the email address provided in the event object (e).
If the email address is not available, it falls back to using the user’s nickname.
Returns the formatted username.

onEdit(e)


Triggered when a cell is edited in the specified column (Column L).
Retrieves relevant data from the edited row (e.g., username, mail ID, password, assets).
Composes an HTML email body with the extracted information.
Sends an email to additional recipients (specified email addresses).
Updates other columns (e.g., Entity, Created, Trigger Date) in the Google Sheet.

getEmailEntity(email)


Extracts the “entity” from an email address (excluding the .com domain).
Used to determine the relevant entity based on the email address.
Customization
Modify the email body template (body) to suit your needs.
Adjust the column numbers and additional email addresses as required.
