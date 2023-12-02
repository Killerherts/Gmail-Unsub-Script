const LABEL_TO_WATCH = 'Unsubscribe';
const LOG_SHEET_INDEX = 0; // Index of the sheet to log in (0 for the first sheet)

/**
 * Main function to run the script.
 * It sets up the necessary components and then runs the unsubscribe process.
 */
function run() {
  setupLogSheet(); // Ensure log sheet is set up
  unsubscribeAndDelete(); // Run the unsubscribe and delete process
}

/**
 * Sets up the log sheet. It uses the first sheet in the spreadsheet.
 * Adds headers if the sheet is empty.
 */
function setupLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheets()[LOG_SHEET_INDEX];

  if (logSheet.getLastRow() === 0) { // Check if the sheet is empty
    // Add headers
    logSheet.appendRow(['Timestamp', 'Email Subject', 'Status', 'Detail']);
    Logger.log('Headers added to the log sheet.');

    // Apply styling to the header row
    const headerRange = logSheet.getRange(1, 1, 1, 4); // Adjust the range as necessary
    headerRange.setBackground('#4a86e8'); // Set background color
    headerRange.setFontColor('#ffffff'); // Set font color
    headerRange.setFontWeight('bold'); // Make text bold
    headerRange.setHorizontalAlignment('center'); // Center align text
    headerRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID); // Set borders
  } else {
    Logger.log('Log sheet already set up.');
  }
}


/**
 * The main process that unsubscribes and deletes emails.
 * It processes each email thread with the specified label.
 */
function unsubscribeAndDelete() {
  const label = GmailApp.getUserLabelByName(LABEL_TO_WATCH);
  const threads = label.getThreads();
  Logger.log('Processing ' + threads.length + ' email threads.');
  threads.forEach(thread => {
    try {
      const message = thread.getMessages()[0];
      const subject = message.getSubject();
      const body = message.getBody();

      const unsubscribed = attemptUnsubscribe(body);
      let status = unsubscribed ? 'Attempted to Unsubscribe' : 'No Unsubscribe Link Found';

      logAction(subject, status, '');
      thread.moveToTrash(); // Optionally delete the email thread
    } catch (e) {
      logAction('Error processing email', 'Failed', e.message);
    }
  });
}

/**
 * Attempts to find and follow an unsubscribe link in the email body.
 * @param {string} emailBody - The body of the email.
 * @return {boolean} - True if an unsubscribe attempt was made, false otherwise.
 */
function attemptUnsubscribe(emailBody) {
  // Use a broad regex pattern to capture a link that includes the word 'unsubscribe'
  let pattern = /<a\s+(?:[^>]*?\s+)?href="([^"]*)"[^>]*>(?:.*?unsubscribe.*?)<\/a>/i;
  let unsubscribeLink = emailBody.match(pattern);

  // If a link is found, try to fetch it
  if (unsubscribeLink && unsubscribeLink[1]) {
    let link = unsubscribeLink[1].trim();
    // Some additional logging for debugging
    Logger.log('Unsubscribe link found: ' + link);
    try {
      // Making a GET request to the unsubscribe link
      let response = UrlFetchApp.fetch(link, { method: 'get', muteHttpExceptions: true });
      // Log the HTTP response code for confirmation
      Logger.log('Request sent. Response code was: ' + response.getResponseCode());
      return true;
    } catch (e) {
      // Log any errors encountered during the request
      Logger.log('Error following unsubscribe link: ' + e.message);
      return false;
    }
  } else {
    // Log if no unsubscribe link is found
    Logger.log('No unsubscribe link found in the email.');
    return false;
  }
}

/**
 * Logs actions taken by the script to the designated log sheet.
 * @param {string} subject - The subject of the email.
 * @param {string} status - The status of the action.
 * @param {string} detail - Additional details about the action.
 */
function logAction(subject, status, detail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheets()[LOG_SHEET_INDEX];
  logSheet.appendRow([new Date(), subject, status, detail]);
}
