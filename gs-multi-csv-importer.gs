/**
 * Imports CSV files from a specified Google Drive folder into a Google Spreadsheet.
 *
 * This script prompts the user to enter the folder ID where the CSV files are stored. 
 * It then imports each CSV file into a new sheet within the active spreadsheet.
 * If a sheet with the same name already exists, it deletes the old sheet before importing the new one.
 *
 * Steps of the script:
 * 1. Prompts the user to enter the Google Drive folder ID.
 * 2. Retrieves all CSV files from the specified folder.
 * 3. For each CSV file:
 *    - Reads the CSV data.
 *    - Creates a new sheet with the same name as the CSV file (excluding the .csv extension).
 *    - Deletes the sheet if it already exists.
 *    - Inserts the CSV data into the new sheet.
 *    - Makes the first row bold.
 *    - Freezes the first row.
 * 4. If the user cancels the prompt, displays a cancellation alert.
 *
 * Requirements:
 * - The user must have access to the specified Google Drive folder.
 * - The folder ID should be valid and contain at least one CSV file.
 * - The script must be run from within a Google Spreadsheet.
 *
 * Expected Output:
 * - Each CSV file in the specified folder will be imported into a new sheet within the active spreadsheet.
 * - The new sheets will be named after the CSV files (excluding the .csv extension).
 * - The first row of each new sheet will be bold and frozen.
 */

function importCSVsFromDrive() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Import CSVs', 'Please enter the folder ID:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var folderId = response.getResponseText();
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFilesByType(MimeType.CSV);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  
    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();
      var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
      
      var sheetName = fileName.replace(/\.csv$/i, ""); // Removes .csv from the file name
      var sheet = ss.getSheetByName(sheetName);
      
      if (sheet) {
        ss.deleteSheet(sheet); // Deletes the sheet if it already exists
      }
      
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
      
      // Makes the first row bold
      sheet.getRange(1, 1, 1, csvData[0].length).setFontWeight('bold');
      
      // Freezes the first row
      sheet.setFrozenRows(1);
    }
  } else {
    ui.alert('Operation cancelled.');
  }
}
