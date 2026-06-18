/**
 * Opens the file picker dialog and sets the target to the 'PlanFiles' named range.
 */
function showPlanPicker() {
  const html = HtmlService.createTemplateFromFile('Picker');
  html.targetRange = 'PlanFiles'; // Pass the target range name to the HTML
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(700).setHeight(500), 'Select Plan File');
}

/**
 * Opens the file picker dialog and sets the target to the 'SpecFiles' named range.
 */
function showSpecPicker() {
  const html = HtmlService.createTemplateFromFile('Picker');
  html.targetRange = 'SpecFiles'; // Pass the target range name to the HTML
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(700).setHeight(500), 'Select Specification File');
}

/**
 * A server-side function called by the picker dialog to write the selected file's URL
 * into the specified named range.
 *
 * @param {string} url The URL of the selected file.
 * @param {string} rangeName The named range to update (e.g., 'PlanFiles').
 */
function saveUrlToSheet(url, rangeName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const range = sheet.getRangeByName(rangeName);
  if (range) {
    range.setValue(url);
  } else {
    // If the named range doesn't exist, show an error.
    SpreadsheetApp.getUi().alert(`Error: Named range "${rangeName}" not found.`);
  }
}

/**
 * Gets the user's OAuth 2.0 access token. This is required for the
 * Google Picker API to access the user's Drive files.
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthToken() {
  DriveApp.getRootFolder(); // A simple call to a Drive service to trigger the scope.
  return ScriptApp.getOAuthToken();
}