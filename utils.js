// ---------------------------------------
// Utilities and helpers
// ---------------------------------------

/** Get the user's token */
function _getUserToken() {
  console.info("fn._getUserToken");
  const token = PropertiesService.getUserProperties().getProperty(
    keys.user_token
  );
  console.info("fn._getUserToken.success", token ? "Got Token" : "404");
  return token;
}

function _getAppToken() {
  console.info("fn._getAppToken");
  const token = PropertiesService.getScriptProperties().getProperty(
    keys.app_token
  );
  console.info("fn._getAppToken.success");
  return token;
}

/** We need to write something to the sheet so it will save - undocumented bug in Google sheets  */
function initialiseSheet() {
  console.info("fn.initialiseSheet");
  const sheet = SpreadsheetApp.getActiveSheet();
  const lr = sheet.getLastRow();
  if (lr <= 1) {
    console.info("fn.initialiseSheet", "writing to sheet");
    const range = sheet.getRange(lr + 1, 1);
    range.setValues([[""]]);
  }

  // Get the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const currentName = spreadsheet.getName();

  // Check if the spreadsheet is unnamed or has the default name
  if (currentName === "" || currentName === "Untitled spreadsheet") {
    console.info("fn.initialiseSheet", "Renaming spreadsheet");
    spreadsheet.rename("Cent Budget Spreadsheet");
  }

  console.info("fn.initialiseSheet.success");
}

function isSheetReadOnly() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];

  if (protection && !protection.canEdit()) {
    // The sheet is protected, and the current user can't edit it
    return true;
  }
  // The sheet is not protected or the user can edit it
  return false;
}

/**
 * Helper function to get the ids from a given sheet as a JS array
 * @param {string} sheetName
 * @returns {Array} An array of the ids in the sheet (first element is the header)
 */
function getIds(sheetName) {
  console.info("fn.getIds", sheetName);
  let ss;
  if (sheetName) {
    ss = sheetFromName(sheetName);
  } else {
    ss = SpreadsheetApp.getActiveSheet();
  }

  const lastRow = ss.getLastRow();
  if (lastRow === 0 || !lastRow) {
    console.info("fn.getIds", "No rows in sheet");
    return [];
  }

  const array = ss
    .getRange("A1:A" + lastRow)
    .getValues()
    .join()
    .split(",");

  console.info("fn.getIds.success", array.length);
  return array;
}

/**
 * Helper to get the sheet by name, or create it if it doesn't exist
 * @param {string} name the name of the sheet to get
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} the sheet object
 */
function sheetFromName(name) {
  console.info("fn.sheetFromName", name);
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (sheet === null) {
    console.info("fn.sheetFromName", `Creating new sheet ${name}`);
    sheet = ss.insertSheet(name);
    // Set the timezone
    ss.setSpreadsheetTimeZone("Pacific/Auckland");
    console.info("fn.sheetFromName", "set timezone to Pacific/Auckland");
    createHeaderRow(name);
    createComment(name);
  }
  console.info("fn.sheetFromName.success");
  return sheet;
}
