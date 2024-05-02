/*
This AppScript code looks for a value in a Google worksheet, and returns the name of the specific sheet the value was found in.
*/

function findSheetName(searchValue) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = spreadsheet.getSheets();
  const currentSheet = SpreadsheetApp.getActiveSheet();
  let foundSheets = [];

  for (const sheet of allSheets) {
    if (sheet.getName() !== currentSheet.getName()) { // Check if it's NOT the current sheet
      const data = sheet.getDataRange().getValues();

      for (let row = 0; row < data.length; row++) {
        for (let col = 0; col < data[row].length; col++) {
          if (data[row][col] === searchValue) {
            foundSheets.push(sheet.getName());
            break; // Stop searching this sheet if a match is found
          }
        }
      }
    }
  }

  if (foundSheets.length > 0) {
    return foundSheets;
  } else {
    return "No Sheet Found";
  }
}
