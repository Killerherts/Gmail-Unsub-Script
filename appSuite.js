const LABEL_TO_WATCH = 'Unsubscribe';
const LOG_SHEET_INDEX = 0; // Index of the sheet to log in (0 for the first sheet)

/**
 * Main function to run the script.
 * It sets up the necessary components and then runs the unsubscribe process.
 */
function run() {
  setupTriggers(); // Ensure triggers are set up
  setupLogSheet(); // Ensure log sheet is set up
  unsubscribeAndDelete(); // Run the unsubscribe and delete process
}

/**
 * Checks and sets up the necessary triggers for the script.
 */
function setupTriggers() {
  if (!isTriggerSet('run')) {
    ScriptApp.newTrigger('run')
      .timeBased()
      .everyMinutes(5) // Set this to your desired frequency
      .create();
    Logger.log('Trigger for "run" function created.');
  } else {
    Logger.log('Trigger for "run" function already exists.');
  }
}

/**
 * Checks if a trigger for a specific function name already exists.
 * @param {string} functionName - The name of the function to check for.
 * @return {boolean} - True if the trigger exists, false otherwise.
 */
function isTriggerSet(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      return true;
    }
  }
  return false;
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
  let pattern = /<a\s+(?:[^>]*?\s+)?href="([^"]*)"[^>]*>(?:.*?unsubscribe.*?)<\/a>/i;
  let unsubscribeLink = emailBody.match(pattern);

  if (unsubscribeLink && unsubscribeLink[1]) {
    let link = unsubscribeLink[1].trim();
    Logger.log('Unsubscribe link found: ' + link);
    try {
      // Set the timeout for the request
      let options = {
        method: 'get',
        muteHttpExceptions: true,
        timeoutSeconds: 25 // Timeout set to 25 seconds
      };

      let response = UrlFetchApp.fetch(link, options);

      Logger.log('Request sent. Response code was: ' + response.getResponseCode());
      return true;
    } catch (e) {
      // Handle timeout or other errors
      Logger.log('Error following unsubscribe link or request timed out: ' + e.message);
      return false;
    }
  } else {
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
